Public Function GetWorkSchedule(brigadeID As Long, StartDate As Date, EndDate As Date) As DAO.Recordset
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim daysSQL As String
    
    ' Создаем временную таблицу дней (если ее еще нет)
    CreateDaysTableIfNotExists
    
    ' Формируем основной запрос
'    sql = "SELECT " & _
'          "DateAdd('d', D.День-1, #" & Format(StartDate, "yyyy-mm-dd") & "#) AS WorkDate, " & _
'          "Choose( " & _
'          "  (DateDiff('d', Г.ДатаНачала_ГрафикРаботы, DateAdd('d', D.День-1, #" & Format(StartDate, "yyyy-mm-dd") & "#)) Mod Г.Цикл_ГрафикРаботы) + 1, " & _
'          "  Г.Д1_ГрафикРаботы, Г.Д2_ГрафикРаботы, Г.Д3_ГрафикРаботы, Г.Д4_ГрафикРаботы, Г.Д5_ГрафикРаботы, " & _
'          "  Г.Д6_ГрафикРаботы, Г.Д7_ГрафикРаботы, Г.Д8_ГрафикРаботы, Г.Д9_ГрафикРаботы, Г.Д10_ГрафикРаботы, Г.Д11_ГрафикРаботы " & _
'          ") AS ЧАСЫ, " & _
'          "Choose( " & _
'          "  (DateDiff('d', Г.ДатаНачала_ГрафикРаботы, DateAdd('d', D.День-1, #" & Format(StartDate, "yyyy-mm-dd") & "#)) Mod Г.Цикл_ГрафикРаботы) + 1, " & _
'          "  Г.К1_ГрафикРаботы, Г.К2_ГрафикРаботы, Г.К3_ГрафикРаботы, Г.К4_ГрафикРаботы, Г.К5_ГрафикРаботы, " & _
'          "  Г.К6_ГрафикРаботы, Г.К7_ГрафикРаботы, Г.К8_ГрафикРаботы, Г.К9_ГрафикРаботы, Г.К10_ГрафикРаботы, Г.К11_ГрафикРаботы " & _
'          ") AS КОД " & _
'          "FROM тГРАФИКРАБОТЫ AS Г, tblDays AS D " & _
'          "WHERE Г.ID_ГрафикРаботы = " & brigadeID & " " & _
'          "AND DateAdd('d', D.День-1, #" & Format(StartDate, "yyyy-mm-dd") & "#) BETWEEN #" & Format(StartDate, "yyyy-mm-dd") & "# AND #" & Format(EndDate, "yyyy-mm-dd") & "# " & _
'          "AND Month(DateAdd('d', D.День-1, #" & Format(StartDate, "yyyy-mm-dd") & "#)) = " & Month(StartDate) & " " & _
'          "ORDER BY DateAdd('d', D.День-1, #" & Format(StartDate, "yyyy-mm-dd") & "#)"
    
'    ' Формируем основной запрос с учетом праздников
''    sql = "SELECT " & _
''          "  Расчет.WorkDate, " & _
''          "  IIf(" & brigadeID & " =1 And [П.ID_ПраздДни] Is Not Null, '', [Расчет].ЧАСЫ_ОРИГ) AS ЧАСЫ, " & _
''          "  IIf(" & brigadeID & " =1 And [П.ID_ПраздДни] Is Not Null, '', [Расчет].КОД_ОРИГ) AS КОД " & _
''          "FROM (" & _
''          "  SELECT " & _
''          "    DateAdd('d', D.День-1, #" & Format(StartDate, "yyyy-mm-dd") & "#) AS WorkDate, " & _
''          "    Choose((DateDiff('d', Г.ДатаНачала_ГрафикРаботы, DateAdd('d', D.День-1, #" & Format(StartDate, "yyyy-mm-dd") & "#)) Mod Г.Цикл_ГрафикРаботы) + 1, " & _
''          "      Г.Д1_ГрафикРаботы, Г.Д2_ГрафикРаботы, Г.Д3_ГрафикРаботы, Г.Д4_ГрафикРаботы, Г.Д5_ГрафикРаботы, " & _
''          "      Г.Д6_ГрафикРаботы, Г.Д7_ГрафикРаботы, Г.Д8_ГрафикРаботы, Г.Д9_ГрафикРаботы, Г.Д10_ГрафикРаботы, Г.Д11_ГрафикРаботы) AS ЧАСЫ_ОРИГ, " & _
''          "    Choose((DateDiff('d', Г.ДатаНачала_ГрафикРаботы, DateAdd('d', D.День-1, #" & Format(StartDate, "yyyy-mm-dd") & "#)) Mod Г.Цикл_ГрафикРаботы) + 1, " & _
''          "      Г.К1_ГрафикРаботы, Г.К2_ГрафикРаботы, Г.К3_ГрафикРаботы, Г.К4_ГрафикРаботы, Г.К5_ГрафикРаботы, " & _
''          "      Г.К6_ГрафикРаботы, Г.К7_ГрафикРаботы, Г.К8_ГрафикРаботы, Г.К9_ГрафикРаботы, Г.К10_ГрафикРаботы, Г.К11_ГрафикРаботы) AS КОД_ОРИГ " & _
''          " FROM тГРАФИКРАБОТЫ AS Г, tblDays AS D " & _
''          " WHERE Г.ID_ГрафикРаботы = " & brigadeID & " " & _
''          ") AS [Расчет] LEFT JOIN тПРАЗДНИКИ AS П ON (" & _
''          "  (П.Ежегодный_ПраздДни = TRUE AND Month(Расчет.WorkDate) = Month(П.Дата_ПраздДни) AND Day(Расчет.WorkDate) = Day(П.Дата_ПраздДни)) " & _
''          "  OR (П.Ежегодный_ПраздДни = FALSE AND Расчет.WorkDate = П.Дата_ПраздДни)" & _
''          ") " & "WHERE Расчет.WorkDate BETWEEN #" & Format(StartDate, "yyyy-mm-dd") & "# AND #" & Format(EndDate, "yyyy-mm-dd") & "# " & _
''          "AND Month(Расчет.WorkDate) = " & Month(StartDate) & " ORDER BY Расчет.WorkDate"
    
    
    sql = "SELECT " & _
    " Расчет.WorkDate, " & _
    " IIf(" & brigadeID & " =1 And [П.ID_ПраздДни] Is Not Null, '', [Расчет].ЧАСЫ_ОРИГ) AS ЧАСЫ, " & _
    " IIf(" & brigadeID & " =1 And [П.ID_ПраздДни] Is Not Null, '', [Расчет].КОД_ОРИГ) AS КОД, " & _
    " IIf(" & brigadeID & " =1 And [П.ID_ПраздДни] Is Not Null, '', [Расчет].КОД_ОРИГ_ЗП) AS КОД_ЗП " & _
    "FROM (" & _
    " SELECT " & _
    " DateAdd('d', D.День-1, #" & Format(StartDate, "yyyy-mm-dd") & "#) AS WorkDate, " & _
    " Choose((DateDiff('d', Г.ДатаНачала_ГрафикРаботы, DateAdd('d', D.День-1, #" & Format(StartDate, "yyyy-mm-dd") & "#)) Mod Г.Цикл_ГрафикРаботы) + 1, " & _
    "    Г.Д1_ГрафикРаботы, Г.Д2_ГрафикРаботы, Г.Д3_ГрафикРаботы, Г.Д4_ГрафикРаботы, Г.Д5_ГрафикРаботы, Г.Д6_ГрафикРаботы, Г.Д7_ГрафикРаботы, Г.Д8_ГрафикРаботы, Г.Д9_ГрафикРаботы, Г.Д10_ГрафикРаботы, Г.Д11_ГрафикРаботы) AS ЧАСЫ_ОРИГ, " & _
    " Choose((DateDiff('d', Г.ДатаНачала_ГрафикРаботы, DateAdd('d', D.День-1, #" & Format(StartDate, "yyyy-mm-dd") & "#)) Mod Г.Цикл_ГрафикРаботы) + 1, " & _
    "    Г.К1_ГрафикРаботы, Г.К2_ГрафикРаботы, Г.К3_ГрафикРаботы, Г.К4_ГрафикРаботы, Г.К5_ГрафикРаботы, Г.К6_ГрафикРаботы, Г.К7_ГрафикРаботы, Г.К8_ГрафикРаботы, Г.К9_ГрафикРаботы, Г.К10_ГрафикРаботы, Г.К11_ГрафикРаботы) AS КОД_ОРИГ, " & _
    " Choose((DateDiff('d', Г.ДатаНачала_ГрафикРаботы, DateAdd('d', D.День-1, #" & Format(StartDate, "yyyy-mm-dd") & "#)) Mod Г.Цикл_ГрафикРаботы) + 1, " & _
    "    Г.ЗП1_ГрафикРаботы, Г.ЗП2_ГрафикРаботы, Г.ЗП3_ГрафикРаботы, Г.ЗП4_ГрафикРаботы, Г.ЗП5_ГрафикРаботы, Г.ЗП6_ГрафикРаботы, Г.ЗП7_ГрафикРаботы, Г.ЗП8_ГрафикРаботы, Г.ЗП9_ГрафикРаботы, Г.ЗП10_ГрафикРаботы, Г.ЗП11_ГрафикРаботы) AS КОД_ОРИГ_ЗП " & _
    " FROM тГРАФИКРАБОТЫ AS Г, tblDays AS D " & _
    " WHERE Г.ID_ГрафикРаботы = " & brigadeID & " " & _
    ") AS [Расчет] LEFT JOIN тПРАЗДНИКИ AS П ON (" & _
    " (П.Ежегодный_ПраздДни = TRUE AND Month(Расчет.WorkDate) = Month(П.Дата_ПраздДни) AND Day(Расчет.WorkDate) = Day(П.Дата_ПраздДни)) " & _
    " OR (П.Ежегодный_ПраздДни = FALSE AND Расчет.WorkDate = П.Дата_ПраздДни)" & _
    ") " & _
    "WHERE Расчет.WorkDate BETWEEN #" & Format(StartDate, "yyyy-mm-dd") & "# AND #" & Format(EndDate, "yyyy-mm-dd") & "# " & _
    "AND Month(Расчет.WorkDate) = " & Month(StartDate) & " " & _
    "ORDER BY Расчет.WorkDate"
    
    
    'Debug.Print sql
    
    On Error GoTo ErrorHandler
    
    Set db = CurrentDb()
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    Set GetWorkSchedule = rs
    Exit Function
    
ErrorHandler:
    MsgBox "Ошибка при выполнении запроса GetWorkSchedule:" & vbCrLf & Err.Description & vbCrLf & vbCrLf & "SQL:" & vbCrLf & sql, vbCritical
    Set GetWorkSchedule = Nothing
End Function