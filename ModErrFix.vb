Imports System.IO
Imports System.Text
Module ModErrFix
    Sub ErrFix(WhoCall As String, Optional ErrData As String = "")
        Dim FileErr As New System.IO.StreamWriter(DBPath & "\Лог\" & Month(Now) & "." & Day(Now) & "_ObjCodeApp.txt", True)
        ' Select Case WhoCall
        'Case "1C_Cons_Start"
        '        ErrData = "-------------------"
        '    Case "1C_Cons_End"
        '        ErrData = "-------------------"
        '    Case "Mod_MC"
        'Case "DogObjCodeErr"

        '    Case "ModRepTru"
        '    Case "ModRepChmngMC"
        '    Case "ModRepEditor"
        '    Case "ModSchFact"
        '    Case "Bank"
        '    Case "PMK_SUP_Cons"
        '    Case "DatePayKS2"
        '    Case "RepTruDirect"
        'End Select
        'Console.Write(vbCrLf & "Номер Ошибки - " & Err.Number & "/Ошибка - " & Err.Description)
        FileErr.WriteLine(Now & "/" & WhoCall & "/" & ErrData & "/" & Environ("UserName") & "/" & Environ("LogonServer") & "/" & Environ("ComputerName"))
        FileErr.Close()
    End Sub
End Module
