Module ModFunc
    Function ReadFile(VarName As String)
        Dim ReadTxtFile As String, STrex As String, STrt As String, N As Single, Sep As String, FileIni As String
        On Error GoTo ErrFix
        FileIni = DBPath & "\BS.ini"
        ReadTxtFile = My.Computer.FileSystem.ReadAllText(FileIni)
        STrt = ""
        N = 1
        Do While N < Len(ReadTxtFile)
            STrex = Chr(CInt("&H" & Mid(ReadTxtFile, N, 2)))
            STrt &= STrex
            N += 2
        Loop
        Sep = Mid(STrt, 1, 1)
        If InStr(STrt, VarName) > 0 Then
            STrt = Mid(STrt, InStr(STrt, VarName) + Len(VarName), Len(STrt) - (InStr(STrt, VarName) + Len(VarName) - 1))
            STrt = Left(STrt, InStr(STrt, Sep) - 1)
        End If
        ReadFile = STrt
        Exit Function
ErrFix:
        Call ModErrFix.ErrFix("ModFunc", "FuncReadFile/Номер ошибки - " & Err.Number & "/Ошибка - " & Err.Description & "/FileIni - " & FileIni)
    End Function
    Function SQLStr(NameFile As String, ColIst As String, ColParam As String, Optional ByVal Param As String = "", Optional ByVal WhoCall As String = "") As String
        DBCom1.CommandText = "select " & ColIst & " from " & NameFile & " where " & ColParam & Param
        DBReader1 = DBCom1.ExecuteReader()
        DBReader1.Read()
        Select Case WhoCall
            Case "Str"
                Try
                    Return DBReader1.GetValue(0)
                Catch Ex As Exception
                    DBReader1.Close()
                    Return ""
                Finally
                    DBReader1.Close()
                End Try
        End Select
        Return ""
        DBReader1.Close()
    End Function
    Function DateConvert(ByVal Param As String, Optional ByVal WhoCall As String = "") As String
        Dim TxtDay As String, TxtMonth As String, TxtYear As String, DogDate As String
        On Error GoTo ErrFix
        TxtDay = LSet(Param, 2)
        TxtMonth = Mid(Param, 4, 2)
        TxtYear = Mid(Param, 7, 4)
        DogDate = "#" & TxtMonth & "/" & TxtDay & "/" & TxtYear & "#"
        Return DogDate
        Exit Function
ErrFix:
        Call ModErrFix.ErrFix("ModFunc", "FuncDateConvert/Номер ошибки - " & Err.Number & "/Ошибка - " & Err.Description & "/WhoCall - " & WhoCall & "/Дата - " & Param)
    End Function
End Module
