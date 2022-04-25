
Module ModStart
    Public DBConn As New System.Data.OleDb.OleDbConnection
    Public DBConn1 As New System.Data.OleDb.OleDbConnection
    Public DBCom As New System.Data.OleDb.OleDbCommand
    Public DBCom1 As New System.Data.OleDb.OleDbCommand
    Public DBReader As System.Data.OleDb.OleDbDataReader
    Public DBReader1 As System.Data.OleDb.OleDbDataReader
    Public Const DBDate As String = "01.01.1000", NDS As Decimal = 0.2
    Public AppExcel As Object, WBook As Object
    Public AppWord As Object, DocWord As Object ', P As Object
    Public DBPath As String, DocsPath As String, ArchiveFilePath As String, DocsScan As String
    Public ObjCode As String, ZacazNum As String, PodrINN(0) As String
    Sub Main()
        Dim FilesName As String(), FileNameDest As String = "", LogFile As Object
        On Error GoTo ErrFix
        Console.Write(vbCrLf & "Старт инициализации переменных.")
        Call InitVar()
        Console.Write(vbCrLf & "Инициализация переменных закончена.")
        Console.Write(vbCrLf & "Старт анализа PDF-файлов.")
        Call ModDog.PDFRead()
        Console.Write(vbCrLf & "Анализ PDF-файлов закончен.")
        Exit Sub
ErrFix:
        Console.Write(vbCrLf & "Номер Ошибки - " & Err.Number & "/Ошибка - " & Err.Description)
        Call ModErrFix.ErrFix("SBISSUPErr", "Main/Номер Ошибки - " & Err.Number & "/Ошибка - " & Err.Description)
    End Sub
    Sub InitVar()
        Dim OfficeDir As String = "C:\Program Files\Microsoft Office\", OFile As Object
        'AppPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
        DBPath = "\\Srv5\work_pmk\УПРАВЛЕНИЕ\СУП\База"
        DocsScan = "\\Srv5\work_pmk\Scaner\truhin\Договоры"
        DocsPath = DBPath & "\EDODocs"
        OFile = Dir(OfficeDir & "Office??", vbDirectory)
        OfficeDir = OfficeDir & OFile & "\MSACC.OLB"
        'ArchiveFilePath = DBPath & "\" & "Archive"
        If IO.File.Exists(OfficeDir) = False Then
            IO.File.Copy(DBPath & "\MSACC.OLB", OfficeDir)
        End If
        'LogCompName = Environ("computername")
        'LogCompName = Environ("logonserver")
        'LogUserName =
        If Dir(DBPath, vbDirectory) = "" Then
            If Environ("computername") = "EVGENIY-SOFT" Then
                DBPath = "C:\Users\Евгений\Documents\Разработка\ПМК-411\База\ПМКСУПбд"
            ElseIf Environ("computername") = "SONY-EVGENIY" Then
                DBPath = "C:\Users\Evgeniy\Documents\Разработка\ПМК-411\База\ПМКСУПбд"
                DocsScan = DBPath & "\Scan\Договоры"
            End If
            'EdoDocsPath = DBPath & "\EDODocs\Общая"
            Console.Write(vbCrLf & "Программа работает локальной с базой.")
        Else
            Console.Write(vbCrLf & "Программа работает с базой ПМК-411")
        End If
        ArchiveFilePath = DBPath & "\" & "Archive"
        DBConn.ConnectionString = ReadFile("VBDB")
        DBConn.Open()
        DBCom.Connection = DBConn
        DBConn1.ConnectionString = ReadFile("VBDB")
        DBConn1.Open()
        DBCom1.Connection = DBConn1
    End Sub
End Module
