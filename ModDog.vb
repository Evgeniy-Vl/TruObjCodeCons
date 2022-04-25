Imports System.Text.RegularExpressions
Module ModDog
    Dim INN As String = 0, KPP As String = 0, INNA As String = 0, KPPA As String = 0 ', Cash As Decimal = 0
    Dim Charterer As String, Podriad As String
    Dim INNInd As Boolean = False, KPPInd As Boolean = False, PAOInd As Boolean = False, ConfirmInd As Boolean = False, DogDateInd As Boolean = False, DogNumInd As Boolean = False
    Dim DogNum As String, DogNumF As String, DogDate As Date = "01.01.1000", K As Integer = 0
    Sub PDFRead()
        Dim TxtLineFromPdf As String, N As Integer, M As Integer, UboundMass As Integer
        Dim DirectName As String(), FilesName As String(), FileName As String, ObjShortName As String, RP As String
        On Error GoTo ErrFix
        DBCom.CommandText = "Select INN from Подрядчики"
        DBReader = DBCom.ExecuteReader()
        While (DBReader.Read())
            If K = 0 Then
                PodrINN(K) = DBReader.GetString(0)
                K += 1
            Else
                ReDim Preserve PodrINN(K)
                PodrINN(K) = DBReader.GetString(0)
                K += 1
            End If
        End While
        DBReader.Close()
        AppWord = CreateObject("Word.Application")
        AppWord.Visible = True
        DocWord = AppWord.Documents.Add
        DirectName = IO.Directory.GetDirectories(DocsScan)
        For M = 0 To UBound(DirectName)
            FilesName = System.IO.Directory.GetFiles(DirectName(M), "*.pdf", IO.SearchOption.AllDirectories)
            RP = Right(DirectName(M), Len(DirectName(M)) - Len(DocsScan) - 1) & "."
            UboundMass = UBound(FilesName)
            For N = 0 To UboundMass
                Console.Write(vbCrLf & "Начало обработки файла - № " & FilesName(N))
                ObjCode = Right(FilesName(N), Len(FilesName(N)) - Len(DocsScan) - 1)
                ObjShortName = Mid(ObjCode, InStr(ObjCode, "-") + 1, InStr(ObjCode, ".pdf") - InStr(ObjCode, "-") - 1)
                ObjCode = Mid(ObjCode, InStr(ObjCode, "\") + 1, 1)
                ObjCode = ModFunc.SQLStr("Основной", "MAX(ObjCode)", " ObjCode Like '1___22%'",, "Str")
                ObjCode = LSet(ObjCode, 3) + 1 & "-" & Mid(CStr(Year(Now)), 3, 2) & "-00"
                Call FileParsing(FilesName(N))
                DocWord.Close
                'INN = 0 : KPP = 0 : INNA = 0 : KPPA = 0 ': Cash = 0 : CashInd = False 
                'Charterer = "" : Podriad = ""
                'DogNum = ""
                'INNInd = False : DogDateInd = False : ConfirmInd = False
                'DocWord = AppWord.Documents.Open(FilesName(N))
                'DocWord.Activate
                'For Each P In DocWord.Paragraphs
                '    TxtLineFromPdf = P.Range.text
                '    TxtLineFromPdf = Replace(TxtLineFromPdf, ChrW(7), "")
                '    TxtLineFromPdf = Replace(TxtLineFromPdf, """", "")
                '    TxtLineFromPdf = Replace(TxtLineFromPdf, "«", "")
                '    TxtLineFromPdf = Replace(TxtLineFromPdf, "»", "")
                '    TxtLineFromPdf = Replace(TxtLineFromPdf, vbCr, "")
                '    TxtLineFromPdf = Replace(TxtLineFromPdf, vbTab, "")
                '    TxtLineFromPdf = Trim(TxtLineFromPdf)
                '    If TxtLineFromPdf <> "" Then
                '        Console.Write(vbCrLf & TxtLineFromPdf)
                '        'Определяем номер договора
                '        If DogNum = "" Then
                '            Call TestDogNum(TxtLineFromPdf)
                '        Else
                '            If ModFunc.SQLStr("Договоры", "DogDeliv", "DogNum = '" & DogNum & "'",, "Str") = "" Then
                '                DogNum = "/00/"
                '                ModErrFix.ErrFix("ModDog.PDFRead", "РП - " & RP & " в файле (" & Mid(FilesName(N), Len(DirectName(M)) + 2, Len(FilesName(N)) - Len(DirectName(M))) & ") невозможно определить номер документа.")
                '                Exit For
                '            End If
                '        End If
                '        If LCase(TxtLineFromPdf) Like " именуе___ в дальнейшем " Then
                '            DogNum = "/00/"
                '        End If
                '        If InStr(LCase(TxtLineFromPdf), "публичное ") > 0 Then
                '            PAOInd = True
                '        End If
                '        If PAOInd = True Then
                '            'If InStr(LCase(TxtLineFromPdf), "пао мтс") > 0 Then
                '            '    INNA = "7740000076"
                '            '    PAOInd = False
                '            'ElseIf InStr(LCase(TxtLineFromPdf), "пао мегафон") > 0 Then
                '            '    INNA = "7812014560"
                '            '    PAOInd = False
                '            'ElseIf InStr(LCase(TxtLineFromPdf), "пао вымпелком") > 0 Then
                '            '    INNA = "7713076301"
                '            '    PAOInd = False
                '            'ElseIf InStr(LCase(TxtLineFromPdf), "пао ростелеком") > 0 Then
                '            '    INNA = "7707049388"
                '            '    PAOInd = False
                '            'End If
                '        ElseIf InStr(TxtLineFromPdf, "Гражданин ") > 0 Then
                '            Charterer = Mid(TxtLineFromPdf, InStr(TxtLineFromPdf, "Гражданин ") + Len("Гражданин "), InStr(TxtLineFromPdf, ",") - InStr(TxtLineFromPdf, "Гражданин ") - Len("Гражданин "))
                '        End If
                '        'Определяем Контрагентов
                '        If Charterer = "" Or Podriad = "" Then
                '            If KPPInd = True Then
                '                Call TestINNKPP(TxtLineFromPdf)
                '            End If
                '            If InStr(LCase(TxtLineFromPdf), "реквизиты сторон") > 0 Or
                '                InStr(LCase(TxtLineFromPdf), "реквизиты и подписи сторон") > 0 Or
                '                InStr(LCase(TxtLineFromPdf), "реквизиты, адреса и подписи сторон") > 0 Or
                '                InStr(LCase(TxtLineFromPdf), "подписи и реквизиты сторон") > 0 Or
                '                InStr(LCase(TxtLineFromPdf), "адреса и реквизиты сторон") > 0 Or
                '                InStr(LCase(TxtLineFromPdf), "адреса, реквизиты и подписи сторон") > 0 Or
                '                InStr(LCase(TxtLineFromPdf), "адреса и платежные реквизиты сторон") > 0 Or
                '                InStr(LCase(TxtLineFromPdf), "адреса и банковские реквизиты сторон") > 0 Or
                '                InStr(LCase(TxtLineFromPdf), "адреса, банковские и почтовые реквизиты и подписи сторон") > 0 Or
                '                InStr(LCase(TxtLineFromPdf), "юридические адреса сторон") > 0 Or
                '                InStr(LCase(TxtLineFromPdf), "юридические адреса и реквизиты сторон") > 0 Or
                '                InStr(LCase(TxtLineFromPdf), "юридические адреса и подписи сторон") > 0 Or
                '                InStr(LCase(TxtLineFromPdf), "юридические адреса и банковские реквизиты") > 0 Or
                '                InStr(LCase(TxtLineFromPdf), "юридические адреса и банковские реквизиты сторон") > 0 Or
                '                InStr(LCase(TxtLineFromPdf), "юридические адреса и платежные реквизиты сторон") > 0 Or
                '                InStr(LCase(TxtLineFromPdf), "юридические адреса, платежные реквизиты и подписи сторон") > 0 Or
                '                InStr(LCase(TxtLineFromPdf), "юридические адреса, банковские реквизиты и подписи сторон") > 0 Or
                '                InStr(LCase(TxtLineFromPdf), "юридические адреса, реквизиты сторон и подписи сторон") > 0 Or
                '                InStr(LCase(TxtLineFromPdf), "почтовые, юридические адреса и банковские реквизиты сторон") > 0 Or
                '                InStr(LCase(TxtLineFromPdf), "почтовые, юридические адреса и банковские реквизиты сторон") > 0 Or
                '                InStr(LCase(TxtLineFromPdf), "реквизиты, печати и подписи уполномоченных лиц сторон") > 0 Then
                '                'If InStr(LCase(TxtLineFromPdf), "инн ") > 0 Or InStr(LCase(TxtLineFromPdf), "инн: ") > 0 Then
                '                KPPInd = True
                '            End If
                '        End If
                '        ' Если ЭДО определяем дату утверждения
                '        If ConfirmInd = True Then
                '            If TxtLineFromPdf Like "* Бушков Александр Алексеевич, *" Then
                '                DogDateInd = True
                '            End If
                '        End If
                '        If InStr(TxtLineFromPdf, "ОЭДО ООО Компания Тензор") > 0 Or InStr(TxtLineFromPdf, "ЭДО АО ПФ СКБ Контур") > 0 Then
                '            ConfirmInd = True
                '        End If
                '    End If
                'Next
                If Not IsNumeric(Charterer) And Charterer <> "" And DogNum <> "" And DogNum <> "/00/" Then
                    ' Присваеваем новый код
                    DogNumF = Replace(DogNum, "/", "-")
                    DogNumF = Replace(DogNumF, "\", "-")
                    DogNumF = Replace(DogNumF, " ", "-")
                    FileName = ObjCode & "_" & "00" & "_" & CStr(DogDate) & "_" & DogNumF & ".pdf"
                    If ModFunc.SQLStr("DocBuh", "FileName", " FileName = '" & FileName & "'",, "Str") = "" Then
                        DBCom.CommandText = "insert into DocBuh (FileName, DocName, KtrAgent, INN, KPP, ObjCode, RP, Podriad, DocNum, DocDate, DateInsert, Changer) values('" & FileName & "', 'Договор', '" & Charterer & "', '" & INNA & "', '" & KPPA & "', '" & ObjCode & "', '" & RP & "', '" & Podriad & "', '" & DogNum & "', '" & DogDate & "', '" & Now & "', 'TruObjCodeCons')" ', '" & Environ("username") & "', '" & Environ("computername") & "')"
                        DBReader = DBCom.ExecuteReader()
                        DBReader.Close()
                        DBCom.CommandText = "insert into Основной (ObjCode, ObjShotName, RP, Podriad, Charterer, NumDog, DogDeliv, DateInsert, Changer, LogUserName, LogCompName) values('" & ObjCode & "', '" & ObjShortName & "', '" & RP & "', '" & Podriad & "', '" & Charterer & "', '" & DogNum & "', '" & DogDate & "', '" & Now & "', 'TruObjCodeCons', '" & Environ("username") & "', '" & Environ("computername") & "')"
                        DBReader = DBCom.ExecuteReader()
                        DBReader.Close()
                        'DocWord.Close
                        IO.File.Move(sourceFileName:=FilesName(N), destFileName:=ArchiveFilePath & "\" & FileName)
                    Else
                        Console.Write(vbCrLf & "Документ № " & FileName & " уже существует в базе.")
                        ModErrFix.ErrFix("ModDog.PDFRead", "Документ № " & FileName & " уже существует в базе.")
                        Exit For
                    End If
                    ' Console.Write(vbCrLf & vbCrLf & ObjCode & vbCrLf & RP & vbCrLf & DogNum & " - " & DogDate & " - " & " (" & ObjShortName & ")" & vbCrLf & Charterer & vbCrLf & Podriad & "Закончена обработки файла - № " & FilesName(N))
                    ModErrFix.ErrFix("ModDog.PDFRead", "Документу № " & Right(FilesName(N), Len(FilesName(N)) - Len(DirectName(M)) - 1) & " РП - " & RP & " присвоен код - " & ObjCode)
                    ModSendMaile.SendMail("DogObjCodeRP", ObjCode, Right(FilesName(N), Len(FilesName(N)) - Len(DirectName(M)) - 1) & " - " & ObjShortName, RP)
                Else
                    ModErrFix.ErrFix("ModDog.PDFRead", "Документ № " & Right(FilesName(N), Len(FilesName(N)) - Len(DirectName(M)) - 1) & " РП - " & RP & " не определён Номер и Дата документа.")
                    Console.Write(vbCrLf & "Документ № " & Right(FilesName(N), Len(FilesName(N)) - Len(DirectName(M)) - 1) & " РП - " & RP & " не определён Номер и Дата документа.")
                End If
                Console.Write(vbCrLf & "Закончена обработки файла - № " & Right(FilesName(N), Len(FilesName(N)) - Len(DirectName(M)) - 1))
            Next
        Next
        AppWord.Quit
        ModSendMaile.SendMail("DogObjCode", ObjCode)
        Exit Sub
ErrFix:
        System.Console.Write(vbCrLf & Err.Number & " - " & Err.Description)
    End Sub
    Sub FileParsing(FileName As String)
        Dim TxtLineFromPdf As String, N As Integer = 0
        INN = 0 : KPP = 0 : INNA = 0 : KPPA = 0 : Charterer = "" : Podriad = "" : DogNum = "" : DogDate = "01.01.1000"
        INNInd = False : DogDateInd = False : DogNumInd = False : ConfirmInd = False
        DocWord = AppWord.Documents.Open(FileName)
        DocWord.Activate
        For Each P In DocWord.Paragraphs
            TxtLineFromPdf = P.Range.text
            TxtLineFromPdf = Replace(TxtLineFromPdf, ChrW(7), "")
            TxtLineFromPdf = Replace(TxtLineFromPdf, """", "")
            TxtLineFromPdf = Replace(TxtLineFromPdf, "«", "")
            TxtLineFromPdf = Replace(TxtLineFromPdf, "»", "")
            TxtLineFromPdf = Replace(TxtLineFromPdf, vbCr, "")
            TxtLineFromPdf = Replace(TxtLineFromPdf, vbTab, "")
            TxtLineFromPdf = Trim(TxtLineFromPdf)
            If TxtLineFromPdf <> "" Then
                Console.Write(vbCrLf & TxtLineFromPdf)
                'Определяем номер договора
                If InStr(LCase(TxtLineFromPdf), "договор") > 0 And N = 0 Then
                    DogNumInd = True
                End If

                If DogNum = "" And DogNumInd = True And N < 5 Then
                    N += 1
                    Call TestDogNum(TxtLineFromPdf)
                ElseIf N = 5 And DogNumInd = True Then
                    Exit Sub
                ElseIf Charterer = "" And DogNum <> "" Then
                    DBCom1.CommandText = "select Charterer, DogDeliv from Договоры where DogNum = '" & DogNum & "'"
                    DBReader1 = DBCom1.ExecuteReader()
                    DBReader1.Read()
                    Try
                        Charterer = DBReader1.GetValue(0)
                        DogDate = DBReader1.GetValue(1)
                    Catch Ex As Exception
                        Console.Write(vbCrLf & "Договор № " & DogNum & "не занесён в базу.")
                        DBReader1.Close()
                        Exit Sub
                    Finally
                        DBReader1.Close()
                    End Try
                End If
                If Charterer <> "" And INNA = 0 Then
                    DBCom1.CommandText = "select INN, KPP from Заказчики where FullName = '" & Charterer & "'"
                    DBReader1 = DBCom1.ExecuteReader()
                    DBReader1.Read()
                    Try
                        INNA = DBReader1.GetValue(0)
                        KPPA = DBReader1.GetValue(1)
                    Catch Ex As Exception
                        Console.Write(vbCrLf & "Заказчик - " & Charterer & "не занесён в базу.")
                        DBReader1.Close()
                        Exit Sub
                    Finally
                        DBReader1.Close()
                    End Try
                End If
                If LCase(TxtLineFromPdf) Like " именуе___ в дальнейшем " And DogNum = "" Then
                    DogNum = "/00/"
                    Exit Sub
                End If
                If InStr(LCase(TxtLineFromPdf), "публичное ") > 0 Then
                    PAOInd = True
                End If
                If PAOInd = True Then
                    'If InStr(LCase(TxtLineFromPdf), "пао мтс") > 0 Then
                    '    INNA = "7740000076"
                    '    PAOInd = False
                    'ElseIf InStr(LCase(TxtLineFromPdf), "пао мегафон") > 0 Then
                    '    INNA = "7812014560"
                    '    PAOInd = False
                    'ElseIf InStr(LCase(TxtLineFromPdf), "пао вымпелком") > 0 Then
                    '    INNA = "7713076301"
                    '    PAOInd = False
                    'ElseIf InStr(LCase(TxtLineFromPdf), "пао ростелеком") > 0 Then
                    '    INNA = "7707049388"
                    '    PAOInd = False
                    'End If
                ElseIf InStr(TxtLineFromPdf, "Гражданин ") > 0 Then
                    Charterer = Mid(TxtLineFromPdf, InStr(TxtLineFromPdf, "Гражданин ") + Len("Гражданин "), InStr(TxtLineFromPdf, ",") - InStr(TxtLineFromPdf, "Гражданин ") - Len("Гражданин "))
                End If
                'Определяем Контрагентов
                If Charterer = "" Or Podriad = "" Then
                    If KPPInd = True Then
                        Call TestINNKPP(TxtLineFromPdf)
                    End If
                    If InStr(LCase(TxtLineFromPdf), "реквизиты сторон") > 0 Or
                            InStr(LCase(TxtLineFromPdf), "реквизиты и подписи сторон") > 0 Or
                            InStr(LCase(TxtLineFromPdf), "реквизиты, адреса и подписи сторон") > 0 Or
                            InStr(LCase(TxtLineFromPdf), "подписи и реквизиты сторон") > 0 Or
                            InStr(LCase(TxtLineFromPdf), "адреса и реквизиты сторон") > 0 Or
                            InStr(LCase(TxtLineFromPdf), "адреса, реквизиты и подписи сторон") > 0 Or
                            InStr(LCase(TxtLineFromPdf), "адреса и платежные реквизиты сторон") > 0 Or
                            InStr(LCase(TxtLineFromPdf), "адреса и банковские реквизиты сторон") > 0 Or
                            InStr(LCase(TxtLineFromPdf), "адреса, банковские и почтовые реквизиты и подписи сторон") > 0 Or
                            InStr(LCase(TxtLineFromPdf), "юридические адреса сторон") > 0 Or
                            InStr(LCase(TxtLineFromPdf), "юридические адреса и реквизиты сторон") > 0 Or
                            InStr(LCase(TxtLineFromPdf), "юридические адреса и подписи сторон") > 0 Or
                            InStr(LCase(TxtLineFromPdf), "юридические адреса и банковские реквизиты") > 0 Or
                            InStr(LCase(TxtLineFromPdf), "юридические адреса и банковские реквизиты сторон") > 0 Or
                            InStr(LCase(TxtLineFromPdf), "юридические адреса и платежные реквизиты сторон") > 0 Or
                            InStr(LCase(TxtLineFromPdf), "юридические адреса, платежные реквизиты и подписи сторон") > 0 Or
                            InStr(LCase(TxtLineFromPdf), "юридические адреса, банковские реквизиты и подписи сторон") > 0 Or
                            InStr(LCase(TxtLineFromPdf), "юридические адреса, реквизиты сторон и подписи сторон") > 0 Or
                            InStr(LCase(TxtLineFromPdf), "почтовые, юридические адреса и банковские реквизиты сторон") > 0 Or
                            InStr(LCase(TxtLineFromPdf), "почтовые, юридические адреса и банковские реквизиты сторон") > 0 Or
                            InStr(LCase(TxtLineFromPdf), "реквизиты, печати и подписи уполномоченных лиц сторон") > 0 Then
                        'If InStr(LCase(TxtLineFromPdf), "инн ") > 0 Or InStr(LCase(TxtLineFromPdf), "инн: ") > 0 Then
                        KPPInd = True
                    End If
                End If
                ' Если ЭДО определяем дату утверждения
                If ConfirmInd = True Then
                    If TxtLineFromPdf Like "* Бушков Александр Алексеевич, *" Then
                        DogDateInd = True
                    End If
                End If
                If InStr(TxtLineFromPdf, "ОЭДО ООО Компания Тензор") > 0 Or InStr(TxtLineFromPdf, "ЭДО АО ПФ СКБ Контур") > 0 Then
                    ConfirmInd = True
                End If
            End If
        Next
    End Sub
    Sub TestDocDate(Prgf As String)
        Dim M As Integer = 0
        Dim myMatches As MatchCollection
        'Dim WRegex As New Regex("\b(?<day>\d{2})\b")
        Dim DRegex As New Regex("\b(?<date>\d{2}\.\d{2}\.\d{4})\b")
        Dim successfulMatch As Match
        'If Prgf Like "*##.##.####*" Then
        myMatches = DRegex.Matches(Prgf)
        'If Charterer = "ПАО МТС" Then
        For Each successfulMatch In myMatches
            'Console.Write(vbCrLf & successfulMatch.Value & " - " & successfulMatch.Index)
            'If WorkStop = DBDate Then
            DogDate = CDate(Prgf)
            'End If
        Next
        'Exit Sub
        ' End If
        '    For Each successfulMatch In myMatches
        '        ' Console.Write(vbCrLf & successfulMatch.Value & " - " & successfulMatch.Index)
        '        If WorkStart = DBDate Then
        '            WorkStart = successfulMatch.Value
        '        ElseIf WorkStop = DBDate Then
        '            Prgf = successfulMatch.Value
        '            If WorkStart <> CDate(Prgf) Then
        '                WorkStop = CDate(Prgf)
        '            End If
        '        End If
        '        M += 1
        '    Next
        'Else
        '    myMatches = WRegex.Matches(Prgf)
        '    For Each successfulMatch In myMatches
        '        'Console.Write(vbCrLf & successfulMatch.Value & " - " & successfulMatch.Index)
        '        If WorkStart = DBDate Then
        '            WorkStart = CDate(Mid(Prgf, successfulMatch.Index, InStr(Prgf, "г") - successfulMatch.Index))
        '        ElseIf WorkStop = DBDate Then
        '            If M = 1 Then
        '                Prgf = Mid(Prgf, InStr(Prgf, "г") + 2, Len(Prgf) - InStr(Prgf, "г") - 3)
        '            Else
        '                Prgf = Mid(Prgf, successfulMatch.Index, InStr(Prgf, "г") - successfulMatch.Index)
        '            End If
        '            If WorkStart <> CDate(Prgf) Then
        '                WorkStop = CDate(Prgf)
        '            End If
        '        End If
        '        M += 1
        '    Next
        'End If
    End Sub
    Sub TestDogNum(Prgf As String)
        Prgf = Replace(Prgf, vbCr, "")
        Prgf = Replace(Prgf, vbTab, "")
        Prgf = Prgf.TrimEnd("/")
        Prgf = Prgf.TrimStart("_") 'Replace(Prgf, "_", "")
        If InStr(LCase(Prgf), "№") > 0 Then
            Prgf = Mid(Prgf, InStr(LCase(Prgf), "№") + 1, Len(Prgf) - InStr(LCase(Prgf), "№"))
            Prgf = Trim(Prgf)
            If InStr(Prgf, " ") > 0 Then
                DogNum = LSet(Prgf, InStr(Prgf, " ") - 1)
            Else
                DogNum = Prgf
            End If
        End If
        If InStr(LCase(Prgf), "договор подряда") > 0 And DogNum = "" Then
            DogNum = Trim(Mid(Prgf, Len("договор подряда") + 1, Len(Prgf) - Len("договор подряда")))
        End If
        If DogNum <> "" Then DogNumInd = False

    End Sub
    Sub TestZacazNum(Prgf As String)
        Dim myMatches As MatchCollection
        Dim mgfRegex As New Regex("\b(?<ZacNum>\d{6}-\w+) \ b")
        Dim mtsRegex As New Regex("\b(?<ZacNum>\d{6}-\w+)\b")
        Dim successfulMatch As Match
        If Charterer = "ПАО МегаФон" Then
            myMatches = mgfRegex.Matches(Prgf)
        Else
            myMatches = mtsRegex.Matches(Prgf)
        End If
        For Each successfulMatch In myMatches
            'Console.Write(vbCrLf & successfulMatch.Value & " - " & successfulMatch.Index)
            ZacazNum = Mid(Prgf, successfulMatch.Index + 1, Len(Prgf) - successfulMatch.Index + 1)
        Next
    End Sub
    Sub TestINNKPP(Prgf As String)
        Dim myMatches As MatchCollection
        Dim INNRegex As New Regex("\b(?<day>\d{10})\b")
        Dim KPPRegex As New Regex("\b(?<Date>\d{9})\b")
        Dim successfulMatch As Match
        myMatches = INNRegex.Matches(Prgf)
        For Each successfulMatch In myMatches
            'Console.Write(vbCrLf & successfulMatch.Value & " - " & successfulMatch.Index)
            INN = successfulMatch.Value
        Next
        myMatches = KPPRegex.Matches(Prgf)
        For Each successfulMatch In myMatches
            'Console.Write(vbCrLf & successfulMatch.Value & " - " & successfulMatch.Index)
            KPP = successfulMatch.Value
            KPPInd = False
        Next
        If INN <> 0 And KPP <> 0 Then
            If Podriad = "" Then
                For K = 0 To UBound(PodrINN)
                    If PodrINN(K) = INN Then
                        Podriad = ModFunc.SQLStr("Подрядчики", "Podriads", "INN = '" & INN & "'",, "Str")
                        INN = 0 : KPP = 0
                        Exit Sub
                    End If
                Next
            End If
            Charterer = ModFunc.SQLStr("Заказчики", " Count(FullName)", "INN = '" & INN & "' and KPP = '" & KPP & "'",, "Str")
            If Charterer = 1 Then
                Charterer = ModFunc.SQLStr("Заказчики", "FullName", "INN = '" & INN & "' and KPP = '" & KPP & "'",, "Str")
                INNA = INN : KPPA = KPP
                INN = 0 : KPP = 0
            Else
                ModErrFix.ErrFix("ModDog.PDFRead", "Существует больше одного Заказчика с ИНН = " & INN & " КПП = " & KPP & " Договору № " & DogNum & " Код не присвоен.")
            End If
        End If
        If Charterer <> "" And Podriad <> "" Then KPPInd = False
    End Sub
End Module
