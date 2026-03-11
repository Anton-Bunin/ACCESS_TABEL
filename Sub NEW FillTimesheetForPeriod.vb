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
    
    
'On Error GoTo ErrorHandler

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
     sql = "SELECT ID_Персонал, Начало_ОТПУСК As Нач, Окончание_ОТПУСК As Кон, 'от' AS Тип FROM тОТПУСКА " & _
            "WHERE Начало_ОТПУСК <= #" & Format(EndDate, "mm\/dd\/yyyy") & "# " & _
            "AND Окончание_ОТПУСК >= #" & Format(StartDate, "mm\/dd\/yyyy") & "# " & _
            "UNION ALL " & _
            "SELECT ID_Персонал, Начало_БЛ, Окончание_БЛ, 'бл' FROM тБОЛЬНИЧНЫЙ " & _
            "WHERE Начало_БЛ <= #" & Format(EndDate, "mm\/dd\/yyyy") & "# " & _
            "AND (Окончание_БЛ >= #" & Format(StartDate, "mm\/dd\/yyyy") & "# OR Окончание_БЛ IS NULL)"

    Set rsAbsenceAll = db.OpenRecordset(sql, dbOpenSnapshot)  ' #############                 ВСЕ ОТПУСКА И БОЛЬНИЧНЫЕ
    
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
        
    Dim dictManual As Object
    Set dictManual = CreateObject("Scripting.Dictionary")
    Dim rsManual As DAO.Recordset
    Dim manualKey As String
    
' 11 Загружаем только записи с флагом РучноеИзменение за нужный период
    sql = "SELECT ID_Персонал, Дата_Табель FROM тТАБЕЛЬ " & _
          "WHERE ID_Период = " & ID_Период & " AND Nz(РучноеИзменение, False) = True"
    Set rsManual = db.OpenRecordset(sql, dbOpenSnapshot)
    Do Until rsManual.EOF
        ' Создаем уникальный ключ: "ID-Дата"
        manualKey = rsManual!ID_Персонал & "-" & Format(rsManual!Дата_Табель, "yyyy-mm-dd")
        dictManual(manualKey) = True
        rsManual.MoveNext
    Loop
    rsManual.Close
    Set rsManual = Nothing
            
    Dim dictSchedules As Object
    Set dictSchedules = CreateObject("Scripting.Dictionary")
    Dim strSchedKey As String
    
    
    Dim inTrans As Boolean
    inTrans = False
    DBEngine.BeginTrans
    inTrans = True ' Мы вошли в транзакцию
    
    ' 2. Для каждого назначения получаем часы работы
    Do Until rsAssignments.EOF
         ' Вычисление фактической начальной даты (MAX из трех дат)
        actualStart = IIf(StartDate > rsAssignments!ДатаНачала, StartDate, rsAssignments!ДатаНачала)
        actualStart = IIf(actualStart > rsAssignments!ДатаНачДолжности, actualStart, rsAssignments!ДатаНачДолжности)
        ' Вычисление фактической конечной даты (MIN из трех дат)
        actualEnd = IIf(EndDate < rsAssignments!ДатаОкончания, EndDate, rsAssignments!ДатаОкончания)
        actualEnd = IIf(actualEnd < rsAssignments!ДатаОкончДолжности, actualEnd, rsAssignments!ДатаОкончДолжности)
        
        ' Формируем уникальный ключ (ID графика + даты, чтобы кэш был точным)
        strSchedKey = rsAssignments!ID_ГрафикРаботы & "|" & actualStart & "|" & actualEnd
        If Not dictSchedules.Exists(strSchedKey) Then
        ' Вызываем тяжелую функцию ТОЛЬКО если этого графика еще нет в кэше
           Set dictSchedules(strSchedKey) = GetWorkSchedule(rsAssignments!ID_ГрафикРаботы, actualStart, actualEnd)
        End If
        ' Теперь просто берем готовый объект из памяти
        Set rsSchedule = dictSchedules(strSchedKey)
        ' ВАЖНО: так как Recordset один на всех, сбрасываем его в начало перед использованием!
        If Not (rsSchedule.EOF And rsSchedule.BOF) Then rsSchedule.MoveFirst
        
        
        ' 4. Заполняем табель только для дней действия назначения
        If Not (rsSchedule Is Nothing) And Not (rsSchedule.EOF And rsSchedule.BOF) Then  '  And Not (rsSchedule.EOF And rsSchedule.BOF)  добавил 24.10.2025
            Do Until rsSchedule.EOF
                absenceCode = ""
                
                ' Проверяем отпуск
                rsAbsenceAll.Filter = "ID_Персонал = " & rsAssignments!ID_Персонал & _
                                    " AND Нач <= #" & Format(rsSchedule!WorkDate, "mm\/dd\/yyyy") & "#" & _
                                    " AND (Кон >= #" & Format(rsSchedule!WorkDate, "mm\/dd\/yyyy") & "# OR Кон IS NULL)"
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
                    hoursValue = "'" & absenceCode & "'" ' Код отсутствия в кавычках
                    hoursKode = "NULL" ' 17/02/2026 для расчета Зарплаты
                Else
                    If IsNull(rsSchedule!ЧАСЫ) Then
                      hoursValue = "NULL"
                      hoursKode = "NULL" ' 17/02/2026 для расчета Зарплаты
                    Else
                       hoursValue = "'" & rsSchedule!ЧАСЫ & "'"
                       hoursKode = "'" & rsSchedule!КОД_ЗП & "'" ' 17/02/2026 для расчета Зарплаты
                    End If
                End If
                                

                ' Вместо DCount("*", "тТАБЕЛЬ", ...) пишем:
                manualKey = rsAssignments!ID_Персонал & "-" & Format(rsSchedule!WorkDate, "yyyy-mm-dd")
                If Not dictManual.Exists(manualKey) Then
                
                    ' если в переводке в комментарии "работа в выходной день", то поставим галочку в поле [работа в выходной день] -> потребуется в доплатах
                    работаВвыходной = False
                    работаВвыходной = (Nz(rsAssignments!Причина, "") Like "*вых*")
                    
                    Dim итогоВыходной As Long
                    итогоВыходной = работаВвыходной
                    
                    ' если это график 5/2 (ID=1) нужно указать часы работы,т.к. в графике их нет. '  ##########################   03/03/2026 ЭТО ЧТОБ С ТАБЕЛЕ ПРОСТАВЛЯЛИСЬ ЧАСЫ ДЛЯ РАБОТЫ В ВЫХОДНЫЕ
                    ' пусть будет hoursValue = 10 часов чтоб бросалось в глаза
                     If rsAssignments!ID_ГрафикРаботы = 1 And итогоВыходной Then
                       hoursValue = "'10'"
                       hoursKode = "'У'"
                       итогоВыходной = CLng(работаВвыходной Or этоВыходной)
                     End If
                    
'                    ' Вставляем запись
                    ' 17/02/2026 для расчета Зарплаты добавлено значение Код_Табель ("У","Н")
                    sql = "INSERT INTO тТАБЕЛЬ (ID_Персонал, ID_Период, ID_Штат, Дата_Табель, Часы_Табель, Код_Табель, ID_ГрафикРаботы, Работа_в_выходной, Работа_в_праздник) " & _
                          "VALUES (" & rsAssignments!ID_Персонал & ", " & ID_Период & ", " & rsAssignments!ID_Штат & ", " & _
                          "#" & Format(rsSchedule!WorkDate, "yyyy-mm-dd") & "#, " & hoursValue & ", " & hoursKode & ", " & rsAssignments!ID_ГрафикРаботы & ", " & итогоВыходной & ", " & этоПраздник & ")"
                    db.Execute sql, dbFailOnError
                End If '  ##########################   02/09/2025
                rsSchedule.MoveNext
            Loop
        End If
        
'        If Not rsSchedule Is Nothing Then
'           rsSchedule.Close
'           Set rsSchedule = Nothing
'        End If
        If Not rsAbsence Is Nothing Then rsAbsence.Close
        rsAssignments.MoveNext
    Loop
    

    DBEngine.CommitTrans
    inTrans = False ' Транзакция успешно закрыта
    
    rsAssignments.Close
    
    MsgBox "Табель заполнен!", vbInformation
    Exit Sub
    
CleanUp:
    On Error Resume Next
    If Not rsAssignments Is Nothing Then rsAssignments.Close
    If Not rsSchedule Is Nothing Then rsSchedule.Close
    If Not rsAbsence Is Nothing Then rsAbsence.Close
    If Not dictHolidays Is Nothing Then rs.Close
    
    Set rsAbsence = Nothing
    Set rsSchedule = Nothing
    Set rsAssignments = Nothing
    Set dictHolidays = Nothing
    Set db = Nothing
    Set dictSchedules = Nothing
    Exit Sub
    
ErrorHandler:

    Dim errDesc As String: errDesc = Err.Description ' Сначала сохраняем текст!
    Dim errNum As Long: errNum = Err.Number
    
    On Error Resume Next ' Чтобы сам Rollback не вызвал ошибку, если транзакция не успела открыться
    If inTrans Then
        DBEngine.Rollback ' Откатываем только если транзакция была активна
    End If
    
 MsgBox "Ошибка №" & errNum & ": " & errDesc, vbCritical
    Resume CleanUp
End Sub
