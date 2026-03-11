Public Sub TEST_FillTimesheetForPeriod(ID_Период As Long, StartDate As Date, EndDate As Date)
  Dim db As DAO.Database
    Dim rsEmployees As DAO.Recordset    ' Список сотрудников с графиками
    Dim rsSchedule As DAO.Recordset     ' Результат GetWorkSchedule
    Dim rsScheduleCode As DAO.Recordset ' Результат GetWorkSchedule c кодировкой вахты (У, Н)
    Dim rsAbsence As DAO.Recordset      ' Отпуска/больничные
    Dim sql As String
    Dim actualStart As Date, actualEnd As Date
    Dim работаВвыходной As Boolean, этоПраздник As Boolean
    Dim absenceCode As String
    Dim hoursValue As String
    Dim hoursKode As String

    If ID_Период < 1 Then
        ID_Период = Nz(DLookup("ID_Период", "тПЕРИОД_ТАБЕЛЯ", _
                      "Год_Период = " & Year(StartDate) & _
                      " AND Месяц_Период = " & Month(StartDate)), 0)
        If ID_Период = 0 Then
            MsgBox "Период не найден в таблице тПЕРИОД_ТАБЕЛЯ!", vbExclamation
            Exit Sub
        End If
    End If
    
    Set db = CurrentDb()
    db.Execute "DELETE FROM тТАБЕЛЬ WHERE ID_Период = " & ID_Период & " AND Nz(РучноеИзменение, False)=False", dbFailOnError
    
  ' 1. Находим ВСЕ назначения, которые активны ХОТЯ БЫ ОДИН ДЕНЬ в этом месяце
    sql = "SELECT ng.ID_Персонал, ng.ID_ГрафикРаботы, ng.ДатаНачала, ng.ДатаОкончания, " & _
            "nd.ID_Штат, nd.ДатаНачала AS ДатаНачДолжности, Nz(nd.ДатаОкончания, #01/01/9999#) AS ДатаОкончДолжности, ng.Причина " & _
          "FROM тНАЗНАЧЕНИЕ_ГРАФИКА ng " & _
          "INNER JOIN тНАЗНАЧЕНИЕ_ДОЛЖНОСТИ nd ON ng.ID_Персонал = nd.ID_Персонал " & _
          "   AND ng.ДатаНачала <= Nz(nd.ДатаОкончания, #01/01/9999#) " & _
          "   AND Nz(ng.ДатаОкончания, #01/01/9999#) >= nd.ДатаНачала " & _
          "WHERE " & _
            "(ng.ДатаНачала <= #" & Format(EndDate, "yyyy-mm-dd") & "#) AND " & _
            "(Nz(ng.ДатаОкончания, #01/01/9999#) >= #" & Format(StartDate, "yyyy-mm-dd") & "#) AND " & _
            "(nd.ДатаНачала <= #" & Format(EndDate, "yyyy-mm-dd") & "#) AND " & _
            "(Nz(nd.ДатаОкончания, #01/01/9999#) >= #" & Format(StartDate, "yyyy-mm-dd") & "#) " & _
            " AND nd.ID_Штат IS NOT NULL" 'нужны только записи, где есть назначение на должность (строгая связь):

'Debug.Print SQL
    Set rsAssignments = db.OpenRecordset(sql, dbOpenSnapshot)  ' #############                 ВСЕ НАЗНАЧЕНИЯ
    
 ' 3. Проверяем отпуска и больничные для сотрудника
     sql = "SELECT ID_Персонал, Начало_ОТПУСК As Нач, Окончание_ОТПУСК As Кон, 'ОТ' AS Тип FROM тОТПУСКА " & _
            "WHERE Начало_ОТПУСК <= #" & Format(EndDate, "mm/dd/yyyy") & "# " & _
            "AND Окончание_ОТПУСК >= #" & Format(StartDate, "mm/dd/yyyy") & "# " & _
            "UNION ALL " & _
            "SELECT ID_Персонал, Начало_БЛ, Окончание_БЛ, 'БЛ' FROM тБОЛЬНИЧНЫЙ " & _
            "WHERE Начало_БЛ <= #" & Format(EndDate, "mm/dd/yyyy") & "# " & _
            "AND Окончание_БЛ >= #" & Format(StartDate, "mm/dd/yyyy") & "#"

    Set rsAbsenceAll = db.OpenRecordset(sql, dbOpenSnapshot)  ' #############                 ВСЕ ОТПУСКА И
    
' 0. До всех циклов — ПРАЗДНИКИ в Dictionary (чтобы убить DCount)
    Dim dictAnnual As Object
    Set dictAnnual = CreateObject("Scripting.Dictionary") ' Ежегодные
    Dim dictSpecific As Object
    Set dictSpecific = CreateObject("Scripting.Dictionary") ' Разовые (конкретные даты)
    
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT * FROM тПРАЗДНИКИ", dbOpenSnapshot)
    Do Until rs.EOF
        If rs!Ежегодный_ПраздДни = True Then
            ' Ключ: строка "день-месяц" (например, "01-01" для Нового года)
            Dim key As String
            key = Format(rs!Дата_ПраздДни, "dd-mm")
            dictAnnual(key) = True
        Else
            ' Ключ: конкретная полная дата
            dictSpecific(rs!Дата_ПраздДни) = True
        End If
        rs.MoveNext
    Loop
    rs.Close    
    
    ' 2. Для каждого назначения получаем часы работы
    Do Until rsAssignments.EOF
        actualStart = IIf(StartDate > rsAssignments!ДатаНачала, StartDate, rsAssignments!ДатаНачала)
        actualStart = IIf(actualStart > rsAssignments!ДатаНачДолжности, actualStart, rsAssignments!ДатаНачДолжности)
        actualEnd = IIf(EndDate < rsAssignments!ДатаОкончания, EndDate, rsAssignments!ДатаОкончания)
        actualEnd = IIf(actualEnd < rsAssignments!ДатаОкончДолжности, actualEnd, rsAssignments!ДатаОкончДолжности)
        If actualStart <= actualEnd Then
            Set rsSchedule = GetWorkSchedule(rsAssignments!ID_ГрафикРаботы, actualStart, actualEnd)
        End If
        
        ' 4. Заполняем табель только для дней действия назначения
        If Not (rsSchedule Is Nothing) And Not (rsSchedule.EOF And rsSchedule.BOF) Then
            Do Until rsSchedule.EOF
                absenceCode = ""
                
                ' Проверяем отпуск
                rsAbsenceAll.Filter = "ID_Персонал = " & rsAssignments!ID_Персонал & _
                                        " AND Нач <= #" & Format(rsSchedule!WorkDate, "mm/dd/yyyy") & "# " & _
                                        " AND Кон >= #" & Format(rsSchedule!WorkDate, "mm/dd/yyyy") & "#"
                Set rsCheck = rsAbsenceAll.OpenRecordset() ' Это мгновенно (In-Memory)
                
                If Not rsCheck.EOF Then
                     absenceCode = rsCheck!Тип
                End If
                rsCheck.Close
                
        ' Проверка Праздника/Выходной праздничный
        этоПраздник = dictAnnual.Exists(Format(rsSchedule!WorkDate, "dd-mm"))
        этоВыходной = dictSpecific.Exists(rsSchedule!WorkDate)
                
                ' Формируем значение для табеля
                If absenceCode <> "" Then
                    hoursValue = "'" & absenceCode & "'" 
                    hoursKode = "NULL" 
                Else
                    If IsNull(rsSchedule!ЧАСЫ) Then
                      hoursValue = "NULL"
                      hoursKode = "NULL" 
                    Else
                       hoursValue = "'" & rsSchedule!ЧАСЫ & "'"
                       hoursKode = "'" & rsSchedule!КОД_ЗП & "'" 
                    End If
                End If
                                
                If DCount("*", "тТАБЕЛЬ", "ID_Персонал = " & rsAssignments!ID_Персонал & _
                                " AND Дата_Табель = # " & Format(rsSchedule!WorkDate, "yyyy-mm-dd") & _
                                " # AND Nz(РучноеИзменение, False)  = True") = 0 Then '  ##########################   02/09/2025 If DCount....Then
                    
                    работаВвыходной = False
                    работаВвыходной = (Nz(rsAssignments!Причина, "") Like "*вых*")
                    
                    Dim итогоВыходной As Long
                    итогоВыходной = работаВвыходной
                    
                     If rsAssignments!ID_ГрафикРаботы = 1 And итогоВыходной Then
                       hoursValue = "'10'"
                       hoursKode = "'У'"
                       итогоВыходной = CLng(работаВвыходной Or этоВыходной)
                     End If                    
                    
'                    ' Вставляем запись
                    sql = "INSERT INTO тТАБЕЛЬ (ID_Персонал, ID_Период, ID_Штат, Дата_Табель, Часы_Табель, Код_Табель, ID_ГрафикРаботы, Работа_в_выходной, Работа_в_праздник) " & _
                          "VALUES (" & rsAssignments!ID_Персонал & ", " & ID_Период & ", " & rsAssignments!ID_Штат & ", " & _
                          "#" & Format(rsSchedule!WorkDate, "yyyy-mm-dd") & "#, " & hoursValue & ", " & hoursKode & ", " & rsAssignments!ID_ГрафикРаботы & ", " & итогоВыходной & ", " & этоПраздник & ")"
'ПОКА ПРОСТО СМОТРЮ РЕЗУЛЬТАТ
Debug.Print sql
'                    db.Execute sql, dbFailOnError
                End If 
                rsSchedule.MoveNext
            Loop
        End If
        
        If Not rsSchedule Is Nothing Then
           rsSchedule.Close
           Set rsSchedule = Nothing
        End If
        If Not rsAbsence Is Nothing Then rsAbsence.Close
        rsAssignments.MoveNext
    Loop
    
    rsAssignments.Close
    
    MsgBox "Табель заполнен!", vbInformation
    Exit Sub
    
CleanUp:
    On Error Resume Next
    If Not rsAssignments Is Nothing Then rsAssignments.Close
    If Not rsSchedule Is Nothing Then rsSchedule.Close
    If Not rsAbsence Is Nothing Then rsAbsence.Close
    Set rsAbsence = Nothing
    Set rsSchedule = Nothing
    Set rsAssignments = Nothing
    Set db = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка в функции заполнения табеля: " & Err.Description & _
           vbCrLf & "Строка: " & Erl, vbCritical
    Resume CleanUp
End Sub