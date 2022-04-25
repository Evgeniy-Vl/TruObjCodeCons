Module ModSendMaile
    Sub SendMail(WhoCall As String, Optional Code As String = "", Optional NameDoc As String = "", Optional RP As String = "")
        Dim objOutlookApp As Object, objMail As Object
        Dim sBody As String = "", sSubject As String = "", sTo As String = "", sCC As String = "", bCC As String = "", AttachFile As String = ""
        'Application.ScreenUpdating = False
        On Error Resume Next
        objOutlookApp = GetObject(, "Outlook.Application")
        Err.Clear()
        If objOutlookApp Is Nothing Then
            objOutlookApp = CreateObject("Outlook.Application")
        End If
        'objOutlookApp.Session.Logon()
        objMail = objOutlookApp.CreateItem(0)
        bCC = "mgf@list.ru"
        If Err.Number <> 0 Then objOutlookApp = Nothing : objMail = Nothing : Exit Sub
        On Error GoTo ErrFix
        Select Case WhoCall
            Case "DogObjCodeRP"
                If RP = "Евгений" Then
                ElseIf RP <> "" Then
                    sTo = "mgf@list.ru"
                    bCC = ""
                End If
                DBReader.Close()
                sSubject = "Присвоен Код."
                sBody = "  РП - " & RP & " Договору № " & NameDoc & " присвоен код - " & ObjCode
            Case "DogObjCode"
                sTo = "truhin@pmk-411.ru"
                sSubject = "Присвоение Кодов."
                sBody = "Информация о присвоении кодов."
                AttachFile = DBPath & "\Лог\" & Month(Now) & "." & Day(Now) & "_ObjCodeApp.txt"
        End Select
        With objMail
            .To = sTo
            .CC = sCC
            .BCC = bCC
            .Subject = sSubject
            .Body = sBody
            If AttachFile <> "" Then .Attachments.Add(AttachFile)
            .Send()
            objOutlookApp = Nothing : objMail = Nothing
            'Application.ScreenUpdating = True
        End With
        Exit Sub
ErrFix:
        Call ModErrFix.ErrFix("ModSendMaile", "SendMail/WhoCall - " & WhoCall & ", RP - " & RP & "/Номер ошибки - " & Err.Number & "/Ошибка - " & Err.Description)
    End Sub
End Module
