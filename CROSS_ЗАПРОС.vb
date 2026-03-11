Sub CROSS_ЗАПРОС(Период As Long)

    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim strSQL As String
    Dim strCrosstabQuery As String
    Dim strInsertSQL As String
    Dim GROUP_BY As String
    Dim SELECT_BY As String
    Dim WHERE As String
    Dim ORDER_BY As String
    Const crosstabQueryName As String = "qryCrosstabTemp"
    
    Dim DOLGNOST As String
    
    Set db = CurrentDb
    
' Шаг 1: SQL кросс-табличного запроса
' по таб номеру
    If Form_фРЕДАКТОР_ТАБЕЛЯ.Группировка.Value = 1 Then
      SELECT_BY = _
        "SELECT p.ТабНомер_Персонал AS ТАБНОМ, p.Фамилия_Персонал & ' ' & p.Имя_Персонал AS ФИО,   " & _
        "       ш.НаименованиеДолжности_ШТАТ AS Название_должности, " & _
        "       ш.ДолжностьID_ШТАТ, " & _
        "       p.ID_Персонал, " & _
        "       Sum(IIf(IsNumeric(Nz([t].[Часы_Табель], 0)), CDbl(Nz([t].[Часы_Табель], 0)), 0)) AS [Всего часов] "
      GROUP_BY = _
        "GROUP BY p.ТабНомер_Персонал, p.Фамилия_Персонал & ' ' & p.Имя_Персонал,  " & _
        "       ш.НаименованиеДолжности_ШТАТ, " & _
        "       ш.ДолжностьID_ШТАТ, " & _
        "       p.ID_Персонал  "
    End If
    
' по Фамилия
    If Form_фРЕДАКТОР_ТАБЕЛЯ.Группировка.Value = 2 Then
      SELECT_BY = _
        "SELECT p.Фамилия_Персонал & ' ' & p.Имя_Персонал AS ФИО, p.ТабНомер_Персонал AS ТАБНОМ,  " & _
        "       ш.НаименованиеДолжности_ШТАТ AS Название_должности, " & _
        "       ш.ДолжностьID_ШТАТ, " & _
        "       p.ID_Персонал, " & _
        "       Sum(IIf(IsNumeric(Nz([t].[Часы_Табель], 0)), CDbl(Nz([t].[Часы_Табель], 0)), 0)) AS [Всего часов] "
      GROUP_BY = _
        "GROUP BY p.Фамилия_Персонал & ' ' & p.Имя_Персонал, p.ТабНомер_Персонал, " & _
        "       ш.НаименованиеДолжности_ШТАТ, " & _
        "       ш.ДолжностьID_ШТАТ, " & _
        "       p.ID_Персонал  "
    End If
    
' по бригаде
    If Form_фРЕДАКТОР_ТАБЕЛЯ.Группировка.Value = 3 Then
      SELECT_BY = _
        "SELECT p.ТабНомер_Персонал AS ТАБНОМ, p.Фамилия_Персонал & ' ' & p.Имя_Персонал AS ФИО, " & _
        "       гр.Название_ГрафикРаботы AS График, " & _
        "       гр.Бригада_ГрафикРаботы AS Бригада, " & _
        "       ш.НаименованиеДолжности_ШТАТ AS Название_должности, " & _
        "       ш.ДолжностьID_ШТАТ, " & _
        "       p.ID_Персонал, t.ID_ГрафикРаботы, " & _
        "       Sum(IIf(IsNumeric(Nz([t].[Часы_Табель], 0)), CDbl(Nz([t].[Часы_Табель], 0)), 0)) AS [Всего часов] "
      GROUP_BY = _
        "GROUP BY гр.Бригада_ГрафикРаботы, гр.Название_ГрафикРаботы, " & _
        "       ш.НаименованиеДолжности_ШТАТ, " & _
        "       ш.ДолжностьID_ШТАТ, " & _
        "       p.Фамилия_Персонал & ' ' & p.Имя_Персонал, " & _
        "       p.ТабНомер_Персонал, " & _
        "       p.ID_Персонал,  " & _
        "       t.ID_ГрафикРаботы "
    End If
    
    DOLGNOST = ""
    If Me.CheckBox_M.Value Then
        DOLGNOST = " and ш.НаименованиеДолжности_ШТАТ Like '*машин*' "
    Else
        If Me.CheckBox_O.Value Then
            DOLGNOST = " and ш.НаименованиеДолжности_ШТАТ Like '*операт*' "
        Else
            If Me.CheckBox_O.Value And Me.CheckBox_M.Value Then
                DOLGNOST = " AND ш.НаименованиеДолжности_ШТАТ Like '*операт*' AND ш.НаименованиеДолжности_ШТАТ Like '*машин*' "
            End If
        End If
    End If
            
    WHERE = "pd.ID_Период = " & Период & " "                                              ' 1. Базовое обязательное условие
    If Me.Бригада > 0 Then _
                                WHERE = WHERE & " AND гр.ID_ГрафикРаботы = " & Me.Бригада       ' 2. Дополнительное условие для Бригады (если есть)
    If Me.фВыбранныйРаботник > 0 Then _
                                WHERE = WHERE & " AND t.ID_Персонал = " & Me.фВыбранныйРаботник ' 3. Другое дополнительное условие (пример)
    If WHERE <> "" Then
        WHERE = "WHERE " & WHERE        ' 5. Формируем итоговую строку WHERE
    Else
        WHERE = "" ' Обработка случая, если вообще нет условий (крайне редко, но на всякий случай)
    End If
            
'    ORDER_BY = "ORDER BY p.Фамилия_Персонал & ' ' & p.Имя_Персонал ASC "   ' с КРОС-ЗАПРОСОМ ORDER BY ваааще не дружит..... просто 3,14ц
'
'    If Form_фРЕДАКТОР_ТАБЕЛЯ.Группировка.Value = 3 Then
'         ', шт.НаименованиеДолжности_ШТАТ DESC, p.Фамилия_Персонал & ' ' & p.Имя_Персонал ASC
'        'Debug.Print "ORDER BY t.ID_ГрафикРаботы ASC "
'    End If
'
'    If CheckBOX_sortDolgnost.Value Then _
'        ORDER_BY = "ORDER BY шт.НаименованиеДолжности_ШТАТ DESC, p.Фамилия_Персонал & ' ' & p.Имя_Персонал ASC "

    strCrosstabQuery = _
        "TRANSFORM First(t.Часы_Табель) AS Часы " & _
         SELECT_BY & _
        "FROM ((((тПЕРСОНАЛ AS p " & _
        "INNER JOIN тТАБЕЛЬ AS t ON p.ID_Персонал = t.ID_Персонал) " & _
        "INNER JOIN тГРАФИКРАБОТЫ AS гр ON t.ID_ГрафикРаботы = гр.ID_ГрафикРаботы) " & _
        "INNER JOIN тПЕРИОД_ТАБЕЛЯ AS pd ON t.ID_Период = pd.ID_Период) " & _
        "INNER JOIN тШТАТ AS ш ON t.ID_Штат = ш.ID_Штат) " & _
         WHERE & DOLGNOST & " " & _
         GROUP_BY & _
         ORDER_BY & _
        "PIVOT Day(t.Дата_Табель) In (1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31)"

'Debug.Print strCrosstabQuery

  ' Шаг 2: Удалить старый запрос, если существует
    On Error Resume Next
    db.QueryDefs.Delete crosstabQueryName
    db.Execute "DELETE FROM [tTEMP_CROSS]", dbFailOnError
    On Error GoTo 0
    db.Execute "DELETE FROM [tTEMP_CROSS]", dbFailOnError
  ' Шаг 3: Создать новый кросс-таб запрос  "qryCrosstabTemp"
    Set qdf = db.CreateQueryDef(crosstabQueryName, strCrosstabQuery)
  ' Шаг 4: Выполнить вставку
    strInsertSQL = "INSERT INTO tTEMP_CROSS SELECT * FROM " & crosstabQueryName
    db.Execute strInsertSQL, dbFailOnError

End Sub