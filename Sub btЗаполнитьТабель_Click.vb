Private Sub btЗаполнитьТабель_Click()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim startD As Date, endD As Date
    Dim tabelDate As Date
    Dim sql As String
    Dim isConfirmed As Long
    
    If IsNull(Me.Период) Then
        MsgBox "Нужно выбрать табель!", vbExclamation
        Exit Sub
    End If
    
    isConfirmed = MsgBox("Данные табеля за " & Me.Период.Column(1) & " будут перезаписаны!" & _
    vbCrLf & vbCrLf & "(восстановить старые данные будет невозможно)", vbExclamation + vbYesNo, "Подтверждение")
    
    If isConfirmed = vbYes Then
    
            sql = "SELECT * FROM тПЕРИОД_ТАБЕЛЯ WHERE ID_Период = " & Me.Период
            Set db = CurrentDb()
            Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
            
            tabelDate = DateSerial(rs!Год_Период, rs!Месяц_Период, 1)
            
            startD = DateSerial(Year(tabelDate), Month(tabelDate), 1)  ' Первое число текущего месяца
            endD = DateSerial(Year(tabelDate), Month(tabelDate) + 1, 0) ' Последнее число текущего месяца
        
            FillTimesheetForPeriod Me.Период, startD, endD ' заполняетм строки ТАБЕЛЯ
            
            rs.Close
            Set rs = Nothing
    Else
        MsgBox "Пользователь отменил редактирование.", vbInformation
    End If
    
    CROSS_ЗАПРОС Me.Период
    
    Me.пфРЕДАКТОР_ТАБЕЛЯ.Form.Requery
    Me.Запрос_CROSS.Form.Requery
    Me.d_ПЕРИОД.Requery

End Sub
