Public Sub FillTimesheetForPeriod(ID_Период As Long, StartDate As Date, EndDate As Date)
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
    
   ' Проверка и получение ID периода, если не указан
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
           
   ' Очищаем старые данные за период (если нужно)
    db.Execute "DELETE FROM тТАБЕЛЬ WHERE ID_Период = " & ID_Период & " AND Nz(РучноеИзменение, False)=False", dbFailOnError 'все кроме ручных правок '  ######### 02/09/2025
  
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
    Set rsAssignments = db.OpenRecordset(sql, dbOpenSnapshot)
    ' 2. Для каждого назначения получаем часы работы
    Do Until rsAssignments.EOF
        
'##########       24.10.2025  ###########################
         ' Вычисление фактической начальной даты (MAX из трех дат)
        actualStart = IIf(StartDate > rsAssignments!ДатаНачала, StartDate, rsAssignments!ДатаНачала)
        actualStart = IIf(actualStart > rsAssignments!ДатаНачДолжности, actualStart, rsAssignments!ДатаНачДолжности)
        
        ' Вычисление фактической конечной даты (MIN из трех дат)
        actualEnd = IIf(EndDate < rsAssignments!ДатаОкончания, EndDate, rsAssignments!ДатаОкончания)
        actualEnd = IIf(actualEnd < rsAssignments!ДатаОкончДолжности, actualEnd, rsAssignments!ДатаОкончДолжности)
        
        ' Убедимся, что диапазон корректный
        If actualStart <= actualEnd Then
            ' Получаем часы работы для этого назначения
            Set rsSchedule = GetWorkSchedule(rsAssignments!ID_ГрафикРаботы, actualStart, actualEnd)
           ' Set rsScheduleCode =
        End If
 '##########       24.10.2025  ###########################

'        ' Получаем часы работы для этого назначения
'        Set rsSchedule = GetWorkSchedule(rsAssignments!ID_ГрафикРаботы, actualStart, actualEnd)                     см.выше  измы от  24.10.2025
            
        ' 3. Проверяем отпуска и больничные для сотрудника
        sql = "SELECT Начало_ОТПУСК, Окончание_ОТПУСК, 'ОТ' AS Тип FROM тОТПУСКА " & _
              "WHERE ID_Персонал = " & rsAssignments!ID_Персонал & _
              " AND Начало_ОТПУСК <= #" & Format(EndDate, "yyyy-mm-dd") & "# " & _
              "AND Окончание_ОТПУСК >= #" & Format(StartDate, "yyyy-mm-dd") & "# " & _
              "UNION ALL " & _
              "SELECT Начало_БЛ, Окончание_БЛ, 'БЛ' AS Тип FROM тБОЛЬНИЧНЫЙ " & _
              "WHERE ID_Персонал = " & rsAssignments!ID_Персонал & _
              " AND Начало_БЛ <= #" & Format(EndDate, "yyyy-mm-dd") & "# " & _
              "AND Окончание_БЛ >= #" & Format(StartDate, "yyyy-mm-dd") & "#"
     
        Set rsAbsence = db.OpenRecordset(sql, dbOpenSnapshot)
     

     
        ' 4. Заполняем табель только для дней действия назначения
        If Not (rsSchedule Is Nothing) And Not (rsSchedule.EOF And rsSchedule.BOF) Then  '  And Not (rsSchedule.EOF And rsSchedule.BOF)  добавил 24.10.2025
            Do Until rsSchedule.EOF
                ' Проверяем, попадает ли дата в отпуск или больничный
                absenceCode = ""
                
                ' Проверяем отпуск
                Set rsCheck = db.OpenRecordset( _
                    "SELECT TOP 1 * FROM тОТПУСКА " & _
                    "WHERE ID_Персонал = " & rsAssignments!ID_Персонал & _
                    " AND #" & Format(rsSchedule!WorkDate, "yyyy-mm-dd") & "# BETWEEN Начало_ОТПУСК AND Окончание_ОТПУСК", _
                    dbOpenSnapshot)
                
                If Not rsCheck.EOF Then
                     absenceCode = Nz(rsCheck!Тип_ОТПУСК, "от") ' МАРКЕР  О Т П У С К
                Else
                   ' Проверяем больничный, если не в отпуске
                    Set rsCheck = db.OpenRecordset( _
                           "SELECT TOP 1 1 FROM тБОЛЬНИЧНЫЙ " & _
                           "WHERE ID_Персонал = " & rsAssignments!ID_Персонал & _
                           " AND #" & Format(rsSchedule!WorkDate, "yyyy-mm-dd") & "# >= Начало_БЛ " & _
                           " AND (Окончание_БЛ IS NULL OR #" & Format(rsSchedule!WorkDate, "yyyy-mm-dd") & "# <= Окончание_БЛ)", _
                           dbOpenSnapshot)
                                                              
                       If Not rsCheck.EOF Then absenceCode = "бл" ' МАРКЕР  Б О Л Ь Н И Ч Н Ы Й
                End If
                rsCheck.Close
                
' Проверка Праздника/Выходной праздничный
этоПраздник = DCount("*", "тПРАЗДНИКИ", "Day([Дата_ПраздДни]) = " & Day(rsSchedule!WorkDate) & " AND Month([Дата_ПраздДни]) = " & Month(rsSchedule!WorkDate) & " AND [Ежегодный_ПраздДни] = True") > 0
этоВыходной = DCount("*", "тПРАЗДНИКИ", "[Дата_ПраздДни] = #" & Format(rsSchedule!WorkDate, "mm\/dd\/yyyy") & "#" & " AND [Ежегодный_ПраздДни] = False") > 0
                
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
                                
                If DCount("*", "тТАБЕЛЬ", "ID_Персонал = " & rsAssignments!ID_Персонал & _
                                " AND Дата_Табель = # " & Format(rsSchedule!WorkDate, "yyyy-mm-dd") & _
                                " # AND Nz(РучноеИзменение, False)  = True") = 0 Then '  ##########################   02/09/2025 If DCount....Then
                    
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
                    
                    
                    ' Вставляем запись
                    ' 17/02/2026 для расчета Зарплаты добавлено значение Код_Табель ("У","Н")
                    sql = "INSERT INTO тТАБЕЛЬ (ID_Персонал, ID_Период, ID_Штат, Дата_Табель, Часы_Табель, Код_Табель, ID_ГрафикРаботы, Работа_в_выходной, Работа_в_праздник) " & _
                          "VALUES (" & rsAssignments!ID_Персонал & ", " & ID_Период & ", " & rsAssignments!ID_Штат & ", " & _
                          "#" & Format(rsSchedule!WorkDate, "yyyy-mm-dd") & "#, " & hoursValue & ", " & hoursKode & ", " & rsAssignments!ID_ГрафикРаботы & ", " & итогоВыходной & ", " & этоПраздник & ")"
                    
                    db.Execute sql, dbFailOnError
                End If '  ##########################   02/09/2025
                
                rsSchedule.MoveNext
            Loop
        End If
        
        If Not rsSchedule Is Nothing Then
           rsSchedule.Close
           Set rsSchedule = Nothing ' Важно!   24/10/25
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