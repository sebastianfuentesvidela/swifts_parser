Imports System.Web.Mail
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Module Module1
    Private _c As SwiftsJob.AppLog
    Private _cnxstr As String = ""
    Private _logdir As String = ""
    Private _inputdir As String = ""
    Private _swiftdir As String = ""
    Private _smtpserver As String = ""
    Private _smtpdestin As String = ""
    Private _smtpsender As String = ""
    Private _logFilName As String = "JobCargaSwifts_" & Format(System.DateTime.Today, "dd-MMM-yy") & ".log"
    Private _verboso As Boolean = False

    Function Main(ByVal CmdArgs() As String) As Integer
        Dim ArgNum As Integer, OkOut As Boolean = True

        Dim objIniFile As SwiftsJob.IniFile
        Dim strData As String
        objIniFile = New SwiftsJob.IniFile(AppPath(True) & "CargaSwiftsSalien.ini")
        _logdir = objIniFile.GetString("Params", "LogDir", "")
        If Not _logdir.EndsWith("\") Then
            _logdir = _logdir & "\"
        End If
        _inputdir = objIniFile.GetString("Params", "InputDir", "")
        If Not _inputdir.EndsWith("\") Then
            _inputdir = _inputdir & "\"
        End If
        _cnxstr = objIniFile.GetString("Params", "ConnectString", "")
        _verboso = IIf(objIniFile.GetInteger("Params", "VerboseMode", 0) = 1, True, False)
        _swiftdir = objIniFile.GetString("Params", "OutputDir", "")
        If Not _swiftdir.EndsWith("\") Then
            _swiftdir = _swiftdir & "\"
        End If
        _smtpserver = objIniFile.GetString("Params", "SmtpServer", "")
        _smtpdestin = objIniFile.GetString("Params", "MailsDestino", "")
        _smtpsender = objIniFile.GetString("Params", "MaildeEnvio", "")
        objIniFile = Nothing

        'If CmdArgs.Length > 0 Then
        '    For ArgNum = 0 To UBound(CmdArgs)
        '        Select Case UCase(Left(CmdArgs(ArgNum), 2))
        '            Case "-L"
        '                _logdir = CmdArgs(ArgNum).Substring(2)
        '                If Not _logdir.EndsWith("\") Then
        '                    _logdir = _logdir & "\"
        '                End If
        '            Case "-I"
        '                _inputdir = CmdArgs(ArgNum).Substring(2)
        '                If Not _inputdir.EndsWith("\") Then
        '                    _inputdir = _inputdir & "\"
        '                End If
        '            Case "-V"
        '                _verboso = True
        '            Case "-C"
        '                _cnxstr = CmdArgs(ArgNum).Substring(2)
        '            Case "-D"
        '                _swiftdir = CmdArgs(ArgNum).Substring(2)
        '                If Not _swiftdir.EndsWith("\") Then
        '                    _swiftdir = _swiftdir & "\"
        '                End If
        '            Case Else
        '                System.Console.WriteLine("Comando Desconocido: " & Left(CmdArgs(ArgNum), 2))
        '                Return True
        '        End Select
        '    Next ArgNum
        'Else
        '    System.Console.WriteLine("Args: -L""<log_directory>"" -C""<connection_string>"" -D""<proc_directory>"" -V")
        '    System.Console.WriteLine("   -L: Directorio de Logging")
        '    System.Console.WriteLine("   -C: Conexión a Base de Datos")
        '    System.Console.WriteLine("   -I: Directorio de Recepcion Archivos (OUT's)")
        '    System.Console.WriteLine("   -D: Directorio de Archivos Procesados (OUT's)")
        '    System.Console.WriteLine("   -V: Verboso (Mensajes a Consola)")
        '    System.Console.ReadLine()
        '    Return 1
        'End If
        Dim LogOption As Integer = 1
        'If _verboso Then
        '    System.Console.WriteLine("Log Directory: " & _logdir)
        '    System.Console.WriteLine("Input Directory: " & _inputdir)
        '    System.Console.WriteLine("Output Directory: " & _swiftdir)
        '    System.Console.WriteLine("String Conexion BBDD: " & _cnxstr)
        '    System.Console.WriteLine("Server SMTP: " & _smtpserver)
        '    System.Console.WriteLine("Destinos SMTP: " & _smtpdestin)
        '    System.Console.WriteLine("Verbose Mode")
        '    LogOption = 2
        '    If _logdir = "" Then
        '        System.Console.WriteLine("¡Falta Log Directory! ")
        '        System.Console.ReadLine()
        '        Return 1
        '    End If
        '    If _cnxstr = "" Then
        '        System.Console.WriteLine("¡Falta String de Conexión a DDBB! ")
        '        System.Console.ReadLine()
        '        Return 1
        '    End If
        '    If _inputdir = "" Then
        '        System.Console.WriteLine("¡Falta Directorio de Recepcion de Archivos! ")
        '        System.Console.ReadLine()
        '        Return 1
        '    End If
        '    If _swiftdir = "" Then
        '        System.Console.WriteLine("¡Falta Directorio de Destino! ")
        '        System.Console.ReadLine()
        '        Return 1
        '    End If
        '    If _smtpserver = "" Then
        '        System.Console.WriteLine("¡Falta Definir Servidor SMTP! ")
        '        System.Console.ReadLine()
        '        Return 1
        '    End If
        '    If _smtpdestin = "" Then
        '        System.Console.WriteLine("¡Falta Definir Destinos SMTP! ")
        '        System.Console.ReadLine()
        '        Return 1
        '    End If
        '    If _smtpsender = "" Then
        '        System.Console.WriteLine("¡Falta Definir Cuenta de Envio SMTP! ")
        '        System.Console.ReadLine()
        '        Return 1
        '    End If
        'End If
        _c = New SwiftsJob.AppLog(LogOption, _logdir & _logFilName)

        OkOut = ProcJob()

        'If _verboso Then
        '    System.Console.WriteLine("Presione Enter para Terminar")
        '    System.Console.ReadLine()
        'End If
        If OkOut Then
            Return 1
        Else
            Return 0
        End If
    End Function

    Function ProcJob() As Boolean
        Dim di As IO.DirectoryInfo
        Dim diar1 As IO.FileInfo()
        Dim diarf As IO.FileInfo()
        Dim dra As IO.FileInfo
        Dim dra2 As IO.FileInfo
        Dim arcswf As IO.FileInfo
        Dim datswf As SwiftsJob.SwiftData
        Dim okfile As Boolean = False
        Dim idf As Integer = 0
        Dim arFiles As New ArrayList
        Dim datfile As SwiftsJob.FileMap
        Try
            di = New IO.DirectoryInfo(_inputdir)
            diar1 = di.GetFiles("*.out")
        Catch ex As Exception
            _c.Informa(ex.ToString)
        End Try
        _c.Informa("0" & arFiles.Count)
        Try
            For Each dra In diar1
                datfile = New SwiftsJob.FileMap
                datfile.fullName = dra.FullName
                datfile.fileDate = dra.LastWriteTime
                arFiles.Add(datfile)
            Next
        Catch ex As Exception
            ProcJob = False
            Exit Function
        End Try
        For i As Integer = 0 To arFiles.Count - 1
            datfile = arFiles(i)
            okfile = False : idf = 0
            _c.Informa("==============================================")
            _c.Informa(datfile.fullName) 'dra.FullName)
            If ProcFile(datfile.fullName, datfile.fileDate) Then
                Try
                    _c.Informa("Relocalizando Archivo Swifts: " & Left(datfile.Name, datfile.Name.Length - 4) & " en " & _swiftdir)
                    Do While Not okfile
                        dra2 = New FileInfo(_swiftdir & Left(datfile.Name, datfile.Name.Length - 4) & "_" & idf & ".out")
                        If Not dra2.Exists() Then
                            okfile = True
                        Else
                            idf = idf + 1
                        End If
                    Loop
                    dra = New FileInfo(datfile.fullName)
                    dra.CopyTo(_swiftdir & Left(dra.Name, dra.Name.Length - 4) & "_" & idf & ".out")
                    dra.Delete()
                    ProcJob = True
                    _c.Informa("Ok")
                Catch ex As Exception
                    _c.Informa("Error: " & ex.ToString)
                    ProcJob = False
                    Exit For
                End Try
            Else
                _c.Informa("Error: " & datfile.fullName & " abandonado para proceso siguiente")
                ProcJob = False
            End If
        Next

    End Function

    Function ProcFile(ByVal fil As String, ByVal fechafile As Date) As Boolean
        Dim SwiftDat As SwiftsJob.SwiftData
        Dim errMess As String, artxt, arsxt, a, arsub
        Dim p As Integer, MtNum As Integer
        Dim tstr As String, lindat As String
        Dim oFile As System.IO.File
        Dim oRead As System.IO.StreamReader
        'datswf.NomArchivo = arcswf.Name
        _c.Informa("Iniciando Proceso Archivo Swifts: " & fil & " en " & _swiftdir)
        Try
            oRead = oFile.OpenText(fil)
        Catch ex As Exception
            _c.Informa("Error: " & ex.ToString)
        End Try

        Dim arLin() As String
        Dim i As Integer = 0, k As Integer, j As Integer, linea As String
        While oRead.Peek <> -1
            linea = oRead.ReadLine()
            If Left(linea, 1) <> ":" And Left(linea, 1) <> "-" And Left(linea, 1) <> "" Then
                arLin(i - 1) = arLin(i - 1) & vbCrLf & linea
            Else
                ReDim Preserve arLin(i)
                arLin(i) = Trim(linea)
                If Left(arLin(i), Len("-}")) = "-}" And Len(arLin(i)) > 17 Then
                    k = InStr(arLin(i), "{S:{SAC:}{COP:") + "{S:{SAC:}{COP:".Length + 3
                    k = InStr(arLin(i), "{1:")
                    If k > 0 Then
                        'If Asc(Mid(arLin(i), k, 1)) = 3 And Len(Trim(Mid(arLin(i), k + 1))) > 0 Then
                        i = i + 1
                        ReDim Preserve arLin(i)
                        arLin(i) = Trim(Mid(arLin(i - 1), k))
                        arLin(i - 1) = Trim(Left(arLin(i - 1), k - 1))
                    End If
                    '                _c.Informa(" · " & Mid(arLin(i), k, 1) & ") => " & Asc(Mid(arLin(i), k, 1)))
                End If
                i = i + 1
            End If
        End While
        oRead.Close()
        If i = 0 Then
            _c.Informa("Archivo Vacio : " & fil & " , moviendo a respaldo ")
            ProcFile = True : Exit Function
        End If
        SwiftDat = New SwiftsJob.SwiftData
        For i = 0 To arLin.GetUpperBound(0)
            SwiftDat.dat_All = IIf(Len(SwiftDat.dat_All) = 0, arLin(i), SwiftDat.dat_All & vbCrLf & arLin(i))
            SwiftDat.OrigFile = fil
            Select Case True
                Case Left(arLin(i), Len("-}")) = "-}"
                    k = InStr(1, arLin(i), "{CHK:")
                    If k > 0 Then
                        j = InStr(k, arLin(i), "}")
                        linea = Mid(arLin(i), k, (j - k) + 1)
                        SwiftDat.MsgTrailer = linea
                    End If
                    If Left(SwiftDat.MsgText, 2) = vbCrLf Then
                        SwiftDat.MsgText = Mid(SwiftDat.MsgText, 3)
                        SwiftDat.MsgTEng = Mid(SwiftDat.MsgTEng, 3)
                    End If
                    If SaveDataSwift(SwiftDat, errMess) Then
                        SwiftDat = New SwiftsJob.SwiftData
                    Else
                        ProcFile = False
                        Exit Function
                        '                        SwiftDat = New SwiftsJob.SwiftData
                    End If
                Case Left(arLin(i), Len(":13C:")) = ":13C:"
                    SwiftDat.dat_13C = Mid(arLin(i), Len(":13C:") + 1)
                    artxt = Split(SwiftDat.dat_13C, "/")
                    If UBound(artxt) = 2 Then
                        a = artxt(0)
                        SwiftDat.timeType = artxt(1)
                        'SwiftMT.timeType = SwiftDat.timeType
                        a = artxt(2)  ' time+offset
                        SwiftDat.timeTime = Left(a, 4)
                        'SwiftMT.timeTime = SwiftDat.timeTime
                        SwiftDat.timeOffs = Mid(a, 6)
                        'SwiftMT.timeOffs = SwiftDat.timeOffs
                    End If
                    'SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  13C: " & GetTagText(MtNum, "13C")
                    'SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & SwiftDat.dat_13C
                    'SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  13C: " & GetTagText(MtNum, "13C")
                    'SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        " & SwiftDat.dat_13C
                Case Left(arLin(i), Len(":20:")) = ":20:"
                    SwiftDat.dat_20 = Mid(arLin(i), Len(":20:") + 1)
                    SwiftDat.msgReference = Mid(arLin(i), Len(":20:") + 1)
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "   20: " & GetTagText(MtNum, "20")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & SwiftDat.msgReference
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "   20: " & GetTagTEng(MtNum, "20")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        " & SwiftDat.msgReference
                    Do While Right(SwiftDat.msgReference, 1) = "." And Len(SwiftDat.msgReference) > 0
                        SwiftDat.msgReference = Mid(SwiftDat.msgReference, 1, Len(SwiftDat.msgReference) - 1)
                    Loop
                    'SwiftMT.msgReference = SwiftDat.msgReference
                    _c.Informa("   -- Swift Ref: " & SwiftDat.msgReference)
                Case Left(arLin(i), Len(":21:")) = ":21:"
                    SwiftDat.dat_21 = Mid(arLin(i), Len(":21:") + 1)
                    SwiftDat.relReference = SwiftDat.dat_21
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "   21: " & GetTagText(MtNum, "21")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & SwiftDat.relReference
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "   21: " & GetTagTEng(MtNum, "21")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        " & SwiftDat.relReference
                    'SwiftMT.relReference = SwiftDat.relReference
                Case Left(arLin(i), Len(":23:")) = ":23:"
                    SwiftDat.dat_23 = Mid(arLin(i), Len(":23:") + 1)
                    SwiftDat.issuingBankReference = Left(SwiftDat.dat_23, 16)
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "   23: " & GetTagText(MtNum, "23")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & SwiftDat.issuingBankReference
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "   23: " & GetTagTEng(MtNum, "23")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        " & SwiftDat.issuingBankReference
                    'SwiftMT.bankOperatCode = SwiftDat.bankOperatCode
                Case Left(arLin(i), Len(":23B:")) = ":23B:"
                    SwiftDat.dat_23B = Mid(arLin(i), Len(":23B:") + 1)
                    SwiftDat.bankOperatCode = Left(SwiftDat.dat_23B, 4)
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  23B: " & GetTagText(MtNum, "23B")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & SwiftDat.bankOperatCode
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  23B: " & GetTagTEng(MtNum, "23B")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        " & SwiftDat.bankOperatCode
                    'SwiftMT.bankOperatCode = SwiftDat.bankOperatCode
                Case Left(arLin(i), Len(":23E:")) = ":23E:"
                    a = Mid(arLin(i), Len(":23E:") + 1)
                    SwiftDat.dat_23E = IIf(Len(SwiftDat.dat_23E) = 0, a, SwiftDat.dat_23E & vbCrLf & a)
                    SwiftDat.instructCodeCode = IIf(Len(SwiftDat.instructCodeCode) = 0, Left(a, 4), SwiftDat.instructCodeCode & vbCrLf & Left(a, 4))
                    'SwiftMT.instructCodeCode = SwiftDat.instructCodeCode
                    SwiftDat.instructCodeRefr = IIf(Len(SwiftDat.instructCodeCode) = 0, Mid(a, 6), SwiftDat.instructCodeCode & vbCrLf & Mid(a, 6))
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  23E: " & GetTagText(MtNum, "23E")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & SwiftDat.instructCodeCode
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & SwiftDat.instructCodeRefr
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  23E: " & GetTagTEng(MtNum, "23E")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        " & SwiftDat.instructCodeCode
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        " & SwiftDat.instructCodeRefr
                    'SwiftMT.instructCodeRefr = SwiftDat.instructCodeRefr
                Case Left(arLin(i), Len(":25:")) = ":25:"
                    SwiftDat.dat_25 = Mid(arLin(i), Len(":25:") + 1)
                    SwiftDat.accountIdent = SwiftDat.dat_25
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "   25: " & GetTagText(MtNum, "25")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & SwiftDat.accountIdent
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "   25: " & GetTagTEng(MtNum, "25")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        " & SwiftDat.accountIdent
                    'SwiftMT.transacType = SwiftDat.transacType
                Case Left(arLin(i), Len(":26E:")) = ":26E:"
                    SwiftDat.dat_26E = Mid(arLin(i), Len(":26E:") + 1)
                    SwiftDat.numofAmmend = SwiftDat.dat_26E
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  26E: " & GetTagText(MtNum, "26E")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & SwiftDat.numofAmmend
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  26E: " & GetTagTEng(MtNum, "26E")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        " & SwiftDat.numofAmmend
                    'SwiftMT.transacType = SwiftDat.transacType
                Case Left(arLin(i), Len(":26T:")) = ":26T:"
                    SwiftDat.dat_26T = Mid(arLin(i), Len(":26T:") + 1)
                    SwiftDat.transacType = Left(SwiftDat.dat_26T, 3)
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  26T: " & GetTagText(MtNum, "26T")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & SwiftDat.transacType
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  26T: " & GetTagTEng(MtNum, "26T")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        " & SwiftDat.transacType
                    'SwiftMT.transacType = SwiftDat.transacType
                Case Left(arLin(i), Len(":27:")) = ":27:"
                    SwiftDat.dat_27 = Mid(arLin(i), Len(":27:") + 1)
                    Try
                        arsxt = Split(SwiftDat.dat_27, "/")
                        SwiftDat.seqIndex = arsxt(0)
                        If UBound(arsxt) = 1 Then
                            SwiftDat.seqTotal = arsxt(1)
                        End If
                    Catch ex As Exception
                    End Try
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "   27: " & GetTagText(MtNum, "27")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Pagina Seq. : " & SwiftDat.seqIndex
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Total Sequ. : " & SwiftDat.seqTotal
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "   27: " & GetTagTEng(MtNum, "27")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        Seq. Page   : " & SwiftDat.seqIndex
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        of Total Pg.: " & SwiftDat.seqTotal
                Case Left(arLin(i), Len(":30:")) = ":30:"
                    SwiftDat.dat_30 = IIf(Mid(arLin(i), Len(":30:") + 1, 2) <= "50", "20" & Mid(arLin(i), Len(":30:") + 1, 6), "19" & Mid(arLin(i), Len(":30:") + 1, 6))
                    SwiftDat.dateofAmmend = SwiftDat.dat_30
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "   30: " & GetTagText(MtNum, "30")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Fecha   : " & SwiftDat.dateofAmmend
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "   30: " & GetTagTEng(MtNum, "30")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        Date    : " & SwiftDat.dateofAmmend
                Case Left(arLin(i), Len(":31C:")) = ":31C:"
                    SwiftDat.dat_31C = IIf(Mid(arLin(i), Len(":31C:") + 1, 2) <= "50", "20" & Mid(arLin(i), Len(":31C:") + 1, 6), "19" & Mid(arLin(i), Len(":31C:") + 1, 6))
                    SwiftDat.dateofIssue = SwiftDat.dat_31C
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  31C: " & GetTagText(MtNum, "31C")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Fecha   : " & SwiftDat.dateofIssue
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  31C: " & GetTagTEng(MtNum, "31C")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        Date    : " & SwiftDat.dateofIssue
                Case Left(arLin(i), Len(":31D:")) = ":31D:"
                    SwiftDat.dat_31D = Mid(arLin(i), Len(":31D:") + 1)
                    tstr = IIf(Mid(arLin(i), Len(":31D:") + 1, 2) <= "50", "20" & Mid(arLin(i), Len(":31D:") + 1, 6), "19" & Mid(arLin(i), Len(":31D:") + 1, 6))
                    SwiftDat.dateofExpire = tstr
                    SwiftDat.placeofExpire = Mid(SwiftDat.dat_31D, 7)
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  31D: " & GetTagText(MtNum, "31D")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "       Fecha Expir.: " & SwiftDat.dateofExpire
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "       Lugar Expir.: " & SwiftDat.placeofExpire
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  31D: " & GetTagTEng(MtNum, "31D")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        Exp. Date  : " & SwiftDat.dateofExpire
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        Exp. Place : " & SwiftDat.placeofExpire
                Case Left(arLin(i), Len(":31E:")) = ":31E:"
                    SwiftDat.dat_31E = IIf(Mid(arLin(i), Len(":31E:") + 1, 2) <= "50", "20" & Mid(arLin(i), Len(":31E:") + 1, 6), "19" & Mid(arLin(i), Len(":31E:") + 1, 6))
                    SwiftDat.newdateofExpiry = SwiftDat.dat_31E
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  31E: " & GetTagText(MtNum, "31E")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Fecha   : " & SwiftDat.newdateofExpiry
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  31E: " & GetTagTEng(MtNum, "31E")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        Date    : " & SwiftDat.newdateofExpiry
                Case Left(arLin(i), Len(":32A:")) = ":32A:"
                    SwiftDat.Valuta = IIf(Mid(arLin(i), Len(":32A:") + 1, 2) <= "50", "20" & Mid(arLin(i), Len(":32A:") + 1, 6), "19" & Mid(arLin(i), Len(":32A:") + 1, 6))
                    SwiftDat.Moneda = Mid(arLin(i), Len(":32A:") + 7, 3)
                    SwiftDat.Monto = Mid(arLin(i), Len(":32A:") + 10)
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  32A: " & GetTagText(MtNum, "32A")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Fecha   : " & SwiftDat.Valuta
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Moneda  : " & SwiftDat.Moneda
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Monto   : " & SwiftDat.Monto
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  32A: " & GetTagTEng(MtNum, "32A")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        Date    : " & SwiftDat.Valuta
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        Currency: " & SwiftDat.Moneda
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        Amount  : " & SwiftDat.Monto
                Case Left(arLin(i), Len(":32B:")) = ":32B:"
                    SwiftDat.dat_32B = Mid(arLin(i), Len(":32B:") + 1)
                    SwiftDat.incrCrAmountMon = Mid(arLin(i), Len(":32B:") + 1, 3)
                    SwiftDat.incrCrAmountCant = Mid(arLin(i), Len(":32B:") + 4)
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  32B: " & GetTagText(MtNum, "32B")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Moneda  : " & SwiftDat.incrCrAmountMon
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Monto   : " & SwiftDat.incrCrAmountCant
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  32B: " & GetTagTEng(MtNum, "32B")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        Currency: " & SwiftDat.incrCrAmountMon
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        Amount  : " & SwiftDat.incrCrAmountCant
                Case Left(arLin(i), Len(":33A:")) = ":33A:"
                    SwiftDat.dat_33A = Mid(arLin(i), Len(":32B:") + 1)
                    SwiftDat.fecNetAmmo = IIf(Mid(arLin(i), Len(":33A:") + 1, 2) <= "50", "20" & Mid(arLin(i), Len(":33A:") + 1, 6), "19" & Mid(arLin(i), Len(":33A:") + 1, 6))
                    SwiftDat.instructCurr = Mid(arLin(i), Len(":33A:") + 7, 3)
                    SwiftDat.instructAmount = Mid(arLin(i), Len(":33A:") + 10)
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  33A: " & GetTagText(MtNum, "33A")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Fecha   : " & SwiftDat.fecNetAmmo
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Moneda  : " & SwiftDat.instructCurr
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Monto   : " & SwiftDat.instructAmount
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  33A: " & GetTagTEng(MtNum, "33A")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        Date    : " & SwiftDat.fecNetAmmo
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        Currency: " & SwiftDat.instructCurr
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        Amount  : " & SwiftDat.instructAmount
                Case Left(arLin(i), Len(":33B:")) = ":33B:"
                    SwiftDat.dat_33B = Mid(arLin(i), Len(":33B:") + 1)
                    SwiftDat.instructCurr = Left(SwiftDat.dat_33B, 3)
                    SwiftDat.instructAmount = Mid(SwiftDat.dat_33B, 4)
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  33B: " & GetTagText(MtNum, "33B")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Moneda  : " & SwiftDat.instructCurr
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Monto   : " & SwiftDat.instructAmount
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  33B: " & GetTagTEng(MtNum, "33B")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        Currency: " & SwiftDat.instructCurr
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        Amount  : " & SwiftDat.instructAmount
                Case Left(arLin(i), Len(":34B:")) = ":34B:"
                    SwiftDat.dat_34B = Mid(arLin(i), Len(":34B:") + 1)
                    SwiftDat.newCrAftAmmendCurr = Left(SwiftDat.dat_34B, 3)
                    SwiftDat.newCrAftAmmendAmount = Mid(SwiftDat.dat_34B, 4)
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  34B: " & GetTagText(MtNum, "34B")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Moneda  : " & SwiftDat.newCrAftAmmendCurr
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Monto   : " & SwiftDat.newCrAftAmmendAmount
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  34B: " & GetTagTEng(MtNum, "34B")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        Currency: " & SwiftDat.newCrAftAmmendCurr
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        Amount  : " & SwiftDat.newCrAftAmmendAmount
                Case Left(arLin(i), Len(":36:")) = ":36:"
                    SwiftDat.dat_36 = Mid(arLin(i), Len(":36:") + 1)
                    SwiftDat.exchRate = SwiftDat.dat_36
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "   36: " & GetTagText(MtNum, "36")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & SwiftDat.exchRate
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "   36: " & GetTagTEng(MtNum, "36")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        " & SwiftDat.exchRate
                Case Left(arLin(i), Len(":39A:")) = ":39A:"
                    SwiftDat.dat_39A = Mid(arLin(i), Len(":39A:") + 1)
                    Try
                        arsxt = Split(SwiftDat.dat_39A, "/")
                        SwiftDat.percTolMonto1 = arsxt(0)
                        If UBound(arsxt) = 1 Then
                            SwiftDat.percTolMonto2 = arsxt(1)
                        End If
                    Catch ex As Exception
                    End Try
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  39A: " & GetTagText(MtNum, "39A")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Tolerancia1 : " & SwiftDat.percTolMonto1
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Tolerancia2 : " & SwiftDat.percTolMonto2
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  39A: " & GetTagTEng(MtNum, "39A")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        Tolerance1  : " & SwiftDat.percTolMonto1
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        Tolerance2  : " & SwiftDat.percTolMonto2
                Case Left(arLin(i), Len(":39B:")) = ":39B:"
                    SwiftDat.dat_39B = Mid(arLin(i), Len(":39B:") + 1)
                    SwiftDat.maxCredAmmoMsg = SwiftDat.dat_39B
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  39B: " & GetTagText(MtNum, "39B")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & SwiftDat.maxCredAmmoMsg
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  39B: " & GetTagTEng(MtNum, "39B")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        " & SwiftDat.maxCredAmmoMsg
                Case UCase(Left(arLin(i), Len(":39C:"))) = ":39C:"
                    SwiftDat.dat_39C = Mid(arLin(i), Len(":39C:") + 1)
                    SwiftDat.addicAmmoCover = SwiftDat.dat_39C
                    arsxt = Split(SwiftDat.dat_39C, vbCrLf)
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  39C: " & GetTagText(MtNum, "39C")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  39C: " & GetTagTEng(MtNum, "39C")
                    For k = 0 To UBound(arsxt)
                        SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                        SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                    Next
                Case Left(arLin(i), Len(":40A:")) = ":40A:"
                    SwiftDat.dat_40A = Mid(arLin(i), Len(":40A:") + 1)
                    SwiftDat.typeofCredit = SwiftDat.dat_40A
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  40A: " & GetTagText(MtNum, "40A")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & SwiftDat.typeofCredit
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  40A: " & GetTagTEng(MtNum, "40A")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        " & SwiftDat.typeofCredit
                Case Left(arLin(i), Len(":40E:")) = ":40E:"
                    SwiftDat.dat_40E = Mid(arLin(i), Len(":40E:") + 1)
                    SwiftDat.applicRules = SwiftDat.dat_40E
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  40E: " & GetTagText(MtNum, "40E")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & SwiftDat.applicRules
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  40E: " & GetTagTEng(MtNum, "40E")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        " & SwiftDat.applicRules
                Case Left(arLin(i), Len(":40F:")) = ":40F:"
                    SwiftDat.dat_40F = Mid(arLin(i), Len(":40F:") + 1)
                    SwiftDat.applicRules = SwiftDat.dat_40F
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  40F: " & GetTagText(MtNum, "40F")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & SwiftDat.applicRules
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  40F: " & GetTagTEng(MtNum, "40F")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        " & SwiftDat.applicRules
                Case UCase(Left(arLin(i), Len(":41A:"))) = ":41A:"
                    SwiftDat.dat_41A = Mid(arLin(i), Len(":41A:") + 1)
                    arsxt = Split(SwiftDat.dat_41A, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        SwiftDat.availableWith = SwiftDat.availableWith & vbCrLf & arsxt(k)
                    Next
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  41A: " & GetTagText(MtNum, "41A")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  41A: " & GetTagTEng(MtNum, "41A")
                    For k = 0 To UBound(arsxt)
                        SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                        SwiftDat.MsgTEng = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                    Next
                Case UCase(Left(arLin(i), Len(":41D:"))) = ":41D:"
                    SwiftDat.dat_41D = Mid(arLin(i), Len(":41D:") + 1)
                    arsxt = Split(SwiftDat.dat_41D, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        SwiftDat.availableWith = SwiftDat.availableWith & vbCrLf & arsxt(k)
                    Next
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  41D: " & GetTagText(MtNum, "41D")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  41D: " & GetTagTEng(MtNum, "41D")
                    For k = 0 To UBound(arsxt)
                        SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                        SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                    Next
                Case UCase(Left(arLin(i), Len(":42A:"))) = ":42A:"
                    SwiftDat.dat_42A = Mid(arLin(i), Len(":42A:") + 1)
                    Try
                        If Left(SwiftDat.dat_42A, 1) = "/" Then
                            arsxt = Split(Mid(SwiftDat.dat_42A, 2), vbCrLf)
                            SwiftDat.draweeBankPI = arsxt(0)
                            SwiftDat.draweeBank = arsxt(1)
                        Else
                            SwiftDat.draweeBank = SwiftDat.dat_42A
                        End If
                    Catch ex As Exception
                    End Try
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  42A: " & GetTagText(MtNum, "42A")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "            BIC : " & SwiftDat.draweeBank
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  42A: " & GetTagTEng(MtNum, "42A")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "            BIC : " & SwiftDat.draweeBank
                    tstr = GetBankFromBIC(SwiftDat.draweeBank)
                    arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                        End If
                    Next
                Case UCase(Left(arLin(i), Len(":42C:"))) = ":42C:"
                    SwiftDat.dat_42C = Mid(arLin(i), Len(":42C:") + 1)
                    SwiftDat.draftsAt = SwiftDat.dat_42C
                    arsxt = Split(SwiftDat.dat_42C, vbCrLf)
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  42C: " & GetTagText(MtNum, "42C")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  42C: " & GetTagTEng(MtNum, "42C")
                    For k = 0 To UBound(arsxt)
                        SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                        SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                    Next
                Case UCase(Left(arLin(i), Len(":42D:"))) = ":42D:"
                    SwiftDat.dat_42D = Mid(arLin(i), Len(":42D:") + 1)
                    Try
                        If Left(SwiftDat.dat_42D, 1) = "/" Then
                            arsxt = Split(Mid(SwiftDat.dat_42D, 2), vbCrLf)
                            SwiftDat.draweeBankPI = arsxt(0)
                            For k = 1 To UBound(arsxt)
                                SwiftDat.draweeBank = SwiftDat.draweeBank & vbCrLf & arsxt(k)
                            Next
                            SwiftDat.draweeBank = Mid(SwiftDat.draweeBank, 3)
                        Else
                            arsxt = Split(SwiftDat.dat_42D, vbCrLf)
                            For k = 0 To UBound(arsxt)
                                SwiftDat.draweeBank = SwiftDat.draweeBank & vbCrLf & arsxt(k)
                            Next
                            SwiftDat.draweeBank = Mid(SwiftDat.draweeBank, 3)
                        End If
                    Catch ex As Exception
                    End Try
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  42D: " & GetTagText(MtNum, "42D")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Identf: " & SwiftDat.draweeBankPI
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  42D: " & GetTagTEng(MtNum, "42D")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        Identf: " & SwiftDat.draweeBankPI
                    tstr = SwiftDat.draweeBank : arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "      Nombre/Dir: " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "      Name/Addr.: " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                        End If
                    Next
                Case UCase(Left(arLin(i), Len(":42M:"))) = ":42M:"
                    SwiftDat.dat_42M = Mid(arLin(i), Len(":42M:") + 1)
                    SwiftDat.mixPayDetail = SwiftDat.dat_42M
                    arsxt = Split(SwiftDat.dat_42M, vbCrLf)
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  42M: " & GetTagText(MtNum, "42M")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  42M: " & GetTagTEng(MtNum, "42M")
                    For k = 0 To UBound(arsxt)
                        SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                        SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                    Next
                Case UCase(Left(arLin(i), Len(":42P:"))) = ":42P:"
                    SwiftDat.dat_42P = Mid(arLin(i), Len(":42P:") + 1)
                    SwiftDat.deferrPayDetail = SwiftDat.dat_42P
                    arsxt = Split(SwiftDat.dat_42P, vbCrLf)
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  42P: " & GetTagText(MtNum, "42P")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  42P: " & GetTagTEng(MtNum, "42P")
                    For k = 0 To UBound(arsxt)
                        SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                        SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                    Next
                Case UCase(Left(arLin(i), Len(":43P:"))) = ":43P:"
                    SwiftDat.dat_43P = Mid(arLin(i), Len(":43P:") + 1)
                    SwiftDat.partialShipp = SwiftDat.dat_43P
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  43P: " & GetTagText(MtNum, "43P")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & SwiftDat.partialShipp
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  43P: " & GetTagTEng(MtNum, "43P")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & SwiftDat.partialShipp
                Case UCase(Left(arLin(i), Len(":43T:"))) = ":43T:"
                    SwiftDat.dat_43T = Mid(arLin(i), Len(":43T:") + 1)
                    SwiftDat.transShipp = SwiftDat.dat_43T
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  43T: " & GetTagText(MtNum, "43T")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & SwiftDat.transShipp
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  43T: " & GetTagTEng(MtNum, "43T")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & SwiftDat.transShipp
                Case UCase(Left(arLin(i), Len(":44A:"))) = ":44A:"
                    SwiftDat.dat_44A = Mid(arLin(i), Len(":44A:") + 1)
                    SwiftDat.newDispatchOpt = SwiftDat.dat_44A
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  44A: " & GetTagText(MtNum, "44A")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & SwiftDat.newDispatchOpt
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  44A: " & GetTagTEng(MtNum, "44A")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        " & SwiftDat.newDispatchOpt
                Case UCase(Left(arLin(i), Len(":44B:"))) = ":44B:"
                    SwiftDat.dat_44B = Mid(arLin(i), Len(":44B:") + 1)
                    SwiftDat.newFinalDest = SwiftDat.dat_44B
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  44B: " & GetTagText(MtNum, "44B")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & SwiftDat.newFinalDest
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  44B: " & GetTagTEng(MtNum, "44B")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        " & SwiftDat.newFinalDest
                Case Left(arLin(i), Len(":44C:")) = ":44C:"
                    SwiftDat.dat_44C = IIf(Mid(arLin(i), Len(":44C:") + 1, 2) <= "50", "20" & Mid(arLin(i), Len(":44C:") + 1, 6), "19" & Mid(arLin(i), Len(":44C:") + 1, 6))
                    SwiftDat.lateDateofShipp = SwiftDat.dat_44C
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  44C: " & GetTagText(MtNum, "44C")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Fecha   : " & SwiftDat.lateDateofShipp
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  44C: " & GetTagTEng(MtNum, "44C")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        Date    : " & SwiftDat.lateDateofShipp
                Case UCase(Left(arLin(i), Len(":44D:"))) = ":44D:"
                    SwiftDat.dat_44D = Mid(arLin(i), Len(":44D:") + 1)
                    SwiftDat.shippPeriod = SwiftDat.dat_44D
                    arsxt = Split(SwiftDat.dat_44D, vbCrLf)
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  44D: " & GetTagText(MtNum, "44D")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  44D: " & GetTagTEng(MtNum, "44D")
                    For k = 0 To UBound(arsxt)
                        SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                        SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                    Next
                Case Left(arLin(i), Len(":44E:")) = ":44E:"
                    SwiftDat.dat_44E = Mid(arLin(i), Len(":44E:") + 1)
                    SwiftDat.shippOrigin = SwiftDat.dat_44E
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  44E: " & GetTagText(MtNum, "44E")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & SwiftDat.shippOrigin
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  44E: " & GetTagTEng(MtNum, "44E")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        " & SwiftDat.shippOrigin
                Case Left(arLin(i), Len(":44F:")) = ":44F:"
                    SwiftDat.dat_44F = Mid(arLin(i), Len(":44F:") + 1)
                    SwiftDat.shippDestiny = SwiftDat.dat_44F
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  44F: " & GetTagText(MtNum, "44F")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & SwiftDat.shippDestiny
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  44F: " & GetTagTEng(MtNum, "44F")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        " & SwiftDat.shippDestiny
                Case UCase(Left(arLin(i), Len(":45A:"))) = ":45A:"
                    SwiftDat.dat_45A = Mid(arLin(i), Len(":45A:") + 1)
                    SwiftDat.descripGoods = SwiftDat.dat_45A
                    arsxt = Split(SwiftDat.dat_45A, vbCrLf)
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  45A: " & GetTagText(MtNum, "45A")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  45A: " & GetTagTEng(MtNum, "45A")
                    For k = 0 To UBound(arsxt)
                        SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                        SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                    Next
                Case UCase(Left(arLin(i), Len(":46A:"))) = ":46A:"
                    SwiftDat.dat_46A = Mid(arLin(i), Len(":46A:") + 1)
                    SwiftDat.docsRequired = SwiftDat.dat_46A
                    arsxt = Split(SwiftDat.dat_46A, vbCrLf)
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  46A: " & GetTagText(MtNum, "46A")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  46A: " & GetTagTEng(MtNum, "46A")
                    For k = 0 To UBound(arsxt)
                        SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                        SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                    Next
                Case UCase(Left(arLin(i), Len(":47A:"))) = ":47A:"
                    SwiftDat.dat_47A = Mid(arLin(i), Len(":47A:") + 1)
                    SwiftDat.additCondics = SwiftDat.dat_47A
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  47A: " & GetTagText(MtNum, "47A")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  47A: " & GetTagTEng(MtNum, "47A")
                    arsxt = Split(SwiftDat.dat_47A, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                        SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                    Next
                Case UCase(Left(arLin(i), Len(":48:"))) = ":48:"
                    SwiftDat.dat_48 = Mid(arLin(i), Len(":48:") + 1)
                    SwiftDat.periodPresent = SwiftDat.dat_48
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "   48: " & GetTagText(MtNum, "48")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "   48: " & GetTagTEng(MtNum, "48")
                    arsxt = Split(SwiftDat.dat_48, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                        SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                    Next
                Case UCase(Left(arLin(i), Len(":49:"))) = ":49:"
                    SwiftDat.dat_49 = Mid(arLin(i), Len(":49:") + 1)
                    SwiftDat.confirmInstruct = SwiftDat.dat_49
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "   49: " & GetTagText(MtNum, "49")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & SwiftDat.confirmInstruct
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "   49: " & GetTagTEng(MtNum, "49")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & SwiftDat.confirmInstruct
                Case UCase(Left(arLin(i), Len(":50:"))) = ":50:"
                    SwiftDat.dat_50 = Mid(arLin(i), Len(":50:") + 1)
                    arsxt = Split(SwiftDat.dat_50, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        SwiftDat.applicant = SwiftDat.applicant & vbCrLf & arsxt(k)
                    Next
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "   50: " & GetTagText(MtNum, "50")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "   50: " & GetTagTEng(MtNum, "50")
                    For k = 0 To UBound(arsxt)
                        SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                        SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                    Next
                Case UCase(Left(arLin(i), Len(":50A:"))) = ":50A:"
                    SwiftDat.dat_50A = Mid(arLin(i), Len(":50A:") + 1)
                    Try
                        If Left(SwiftDat.dat_50A, 1) = "/" Then
                            arsxt = Split(Mid(SwiftDat.dat_50A, 2), vbCrLf)  ' ordering Customer
                            SwiftDat.orderingCusAcc = arsxt(0)  ' ordering Customer Cuenta
                            If UBound(arsxt) = 1 Then
                                SwiftDat.orderingCusName = arsxt(1) ' ordering Customer Name
                            ElseIf UBound(arsxt) > 1 Then
                                For k = 1 To UBound(arsxt)
                                    SwiftDat.orderingCusName = SwiftDat.orderingCusName & vbCrLf & arsxt(k)
                                Next
                                SwiftDat.orderingCusName = Mid(SwiftDat.orderingCusName, 3)
                            End If
                        Else
                            SwiftDat.orderingCusAcc = ""  ' ordering Customer Cuenta
                            SwiftDat.orderingCusName = SwiftDat.dat_50A ' ordering Customer Name
                        End If
                    Catch ex As Exception
                    End Try
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  50A: " & GetTagText(MtNum, "50A")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Cuenta  : " & SwiftDat.orderingCusAcc
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  50A: " & GetTagTEng(MtNum, "50A")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        Account : " & SwiftDat.orderingCusAcc
                    tstr = SwiftDat.orderingCusName : arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "         Nombre: " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "         Name  : " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                 " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                 " & arsxt(k)
                        End If
                    Next
                Case Left(arLin(i), Len(":50K:")) = ":50K:"
                    SwiftDat.dat_50K = Mid(arLin(i), Len(":50K:") + 1)
                    Try
                        If Left(SwiftDat.dat_50K, 1) = "/" Then
                            arsxt = Split(Mid(SwiftDat.dat_50K, 2), vbCrLf)  ' ordering Customer
                            SwiftDat.orderingCusAcc = arsxt(0)  ' ordering Customer Cuenta
                            If UBound(arsxt) = 1 Then
                                SwiftDat.orderingCusName = arsxt(1) ' ordering Customer Name
                            ElseIf UBound(arsxt) > 1 Then
                                For k = 1 To UBound(arsxt)
                                    SwiftDat.orderingCusName = SwiftDat.orderingCusName & vbCrLf & arsxt(k)
                                Next
                                SwiftDat.orderingCusName = Mid(SwiftDat.orderingCusName, 3)
                            End If
                        Else
                            SwiftDat.orderingCusAcc = ""  ' ordering Customer Cuenta
                            SwiftDat.orderingCusName = SwiftDat.dat_50K ' ordering Customer Name
                        End If
                    Catch ex As Exception
                    End Try
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  50K: " & GetTagText(MtNum, "50K")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Cuenta  : " & SwiftDat.orderingCusAcc
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  50K: " & GetTagTEng(MtNum, "50K")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        Account : " & SwiftDat.orderingCusAcc
                    tstr = SwiftDat.orderingCusName : arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Nombre  : " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        Name    : " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                        End If
                    Next
                Case Left(arLin(i), Len(":50F:")) = ":50F:"
                    SwiftDat.dat_50F = Mid(arLin(i), Len(":50F:") + 1)
                    Try
                        If Left(SwiftDat.dat_50F, 1) = "/" Then
                            arsxt = Split(Mid(SwiftDat.dat_50F, 2), vbCrLf)  ' ordering Customer
                            SwiftDat.orderingCusAcc = arsxt(0)  ' ordering Customer Cuenta
                            If UBound(arsxt) = 1 Then
                                If InStr(arsxt(1), "/") > 0 Then
                                    arsub = Split(arsxt(1), "/")
                                    If UBound(arsub) > 0 Then
                                        SwiftDat.orderingCusName = arsub(1) ' ordering Customer Name
                                    Else
                                        SwiftDat.orderingCusName = arsxt(1) ' ordering Customer Name
                                    End If
                                End If
                            ElseIf UBound(arsxt) > 1 Then
                                For k = 1 To UBound(arsxt)
                                    arsub = Split(arsxt(k), "/")
                                    If IsNumeric(arsub(0)) Then
                                        If UBound(arsub) > 0 Then
                                            lindat = arsub(1)
                                            For j = 2 To UBound(arsub)
                                                lindat = lindat & "/" & arsub(j)
                                            Next
                                            SwiftDat.orderingCusName = SwiftDat.orderingCusName & vbCrLf & lindat 'arsub(1) ' ordering Customer Name
                                        Else
                                            SwiftDat.orderingCusName = SwiftDat.orderingCusName & vbCrLf & arsxt(k) ' ordering Customer Name
                                        End If
                                    Else
                                        SwiftDat.orderingCusName = SwiftDat.orderingCusName & vbCrLf & arsxt(k)
                                    End If
                                Next
                                SwiftDat.orderingCusName = Mid(SwiftDat.orderingCusName, 3)
                            End If
                        Else
                            SwiftDat.orderingCusAcc = ""  ' ordering Customer Cuenta
                            SwiftDat.orderingCusName = SwiftDat.dat_50F ' ordering Customer Name
                        End If
                    Catch ex As Exception
                    End Try
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  50K: " & GetTagText(MtNum, "50K")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Cuenta  : " & SwiftDat.orderingCusAcc
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  50K: " & GetTagTEng(MtNum, "50K")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        Account : " & SwiftDat.orderingCusAcc
                    tstr = SwiftDat.orderingCusName : arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Nombre  : " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        Name    : " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                        End If
                    Next
                Case UCase(Left(arLin(i), Len(":51A:"))) = ":51A:"
                    SwiftDat.dat_51A = Mid(arLin(i), Len(":51A:") + 1)
                    Try
                        If Left(SwiftDat.dat_51A, 1) = "/" Then
                            arsxt = Split(Mid(SwiftDat.dat_51A, 2), vbCrLf)
                            SwiftDat.applicBankPI = arsxt(0)
                            SwiftDat.applicBank = arsxt(1)
                        Else
                            SwiftDat.applicBank = SwiftDat.dat_51A
                        End If
                    Catch ex As Exception
                    End Try
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  51A: " & GetTagText(MtNum, "51A")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "            BIC : " & SwiftDat.applicBank
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  51A: " & GetTagTEng(MtNum, "51A")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "            BIC : " & SwiftDat.applicBank
                    tstr = GetBankFromBIC(SwiftDat.applicBank)
                    arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                        End If
                    Next
                Case UCase(Left(arLin(i), Len(":51D:"))) = ":51D:"
                    SwiftDat.dat_51D = Mid(arLin(i), Len(":51D:") + 1)
                    Try
                        If Left(SwiftDat.dat_51D, 1) = "/" Then
                            arsxt = Split(Mid(SwiftDat.dat_51D, 2), vbCrLf)
                            SwiftDat.applicBankPI = arsxt(0)
                            For k = 1 To UBound(arsxt)
                                SwiftDat.applicBank = SwiftDat.applicBank & vbCrLf & arsxt(k)
                            Next
                            SwiftDat.applicBank = Mid(SwiftDat.applicBank, 3)
                        Else
                            arsxt = Split(SwiftDat.dat_51D, vbCrLf)
                            For k = 0 To UBound(arsxt)
                                SwiftDat.applicBank = SwiftDat.applicBank & vbCrLf & arsxt(k)
                            Next
                            SwiftDat.applicBank = Mid(SwiftDat.applicBank, 3)
                        End If
                    Catch ex As Exception
                    End Try
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  51D: " & GetTagText(MtNum, "51D")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Identf: " & SwiftDat.applicBankPI
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  51D: " & GetTagTEng(MtNum, "51D")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        Identf: " & SwiftDat.applicBankPI
                    tstr = SwiftDat.applicBank : arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "      Nombre/Dir: " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "      Name/Addr.: " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                        End If
                    Next
                Case UCase(Left(arLin(i), Len(":52A:"))) = ":52A:"
                    SwiftDat.dat_52A = Mid(arLin(i), Len(":52A:") + 1)
                    Try
                        If Left(SwiftDat.dat_52A, 1) = "/" Then
                            arsxt = Split(Mid(SwiftDat.dat_52A, 2), vbCrLf)  ' ordering Customer
                            SwiftDat.origBankPI = arsxt(0)
                            SwiftDat.origBank = arsxt(1)
                        Else
                            SwiftDat.origBank = SwiftDat.dat_52A
                        End If
                    Catch ex As Exception
                    End Try
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  52A: " & GetTagText(MtNum, "52A")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "            BIC : " & SwiftDat.origBank
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  52A: " & GetTagTEng(MtNum, "52A")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "            BIC : " & SwiftDat.origBank
                    tstr = GetBankFromBIC(SwiftDat.origBank)
                    arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                        End If
                    Next
                Case UCase(Left(arLin(i), Len(":52D:"))) = ":52D:"
                    SwiftDat.dat_52D = Mid(arLin(i), Len(":52D:") + 1)
                    Try
                        If Left(SwiftDat.dat_52D, 1) = "/" Then
                            arsxt = Split(Mid(SwiftDat.dat_52D, 2), vbCrLf)  ' ordering Customer
                            SwiftDat.origBankPI = arsxt(0)
                            For k = 1 To UBound(arsxt)
                                SwiftDat.origBank = SwiftDat.origBank & vbCrLf & arsxt(k)
                            Next
                            SwiftDat.origBank = Mid(SwiftDat.origBank, 3)
                        Else
                            SwiftDat.origBank = SwiftDat.dat_52D
                        End If
                    Catch ex As Exception
                    End Try
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  52D: " & GetTagText(MtNum, "52D")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  52D: " & GetTagTEng(MtNum, "52D")
                    tstr = SwiftDat.origBank : arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "      Nonbre/Dir: " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "      Name/Addr.: " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                        End If
                    Next
                Case UCase(Left(arLin(i), Len(":53A:"))) = ":53A:"
                    SwiftDat.dat_53A = Mid(arLin(i), Len(":53A:") + 1)
                    Try
                        If Left(SwiftDat.dat_53A, 1) = "/" Then
                            arsxt = Split(Mid(SwiftDat.dat_53A, 2), vbCrLf)  ' sender correspondent
                            SwiftDat.correspSenderPI = arsxt(0)
                            SwiftDat.correspSender = arsxt(1)
                        Else
                            SwiftDat.correspSender = SwiftDat.dat_53A
                        End If
                    Catch ex As Exception
                    End Try
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  53A: " & GetTagText(MtNum, "53A")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "            BIC : " & SwiftDat.correspSender
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  53A: " & GetTagTEng(MtNum, "53A")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "            BIC : " & SwiftDat.correspSender
                    tstr = GetBankFromBIC(SwiftDat.correspSender)
                    arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                        End If
                    Next
                Case UCase(Left(arLin(i), Len(":53B:"))) = ":53B:"
                    SwiftDat.dat_53B = Mid(arLin(i), Len(":53B:") + 1)
                    Try
                        If Left(SwiftDat.dat_53B, 1) = "/" Then
                            arsxt = Split(Mid(SwiftDat.dat_53B, 2), vbCrLf)  ' sender correspondent
                            SwiftDat.correspSenderPI = arsxt(0)
                            SwiftDat.correspSender = arsxt(1)
                        Else
                            SwiftDat.correspSender = SwiftDat.dat_53B
                        End If
                    Catch ex As Exception
                    End Try
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  53B: " & GetTagText(MtNum, "53B")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  53B: " & GetTagTEng(MtNum, "53B")
                    tstr = SwiftDat.correspSender : arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "      Nombre/Dir: " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "      Name/Addr.: " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                        End If
                    Next
                Case UCase(Left(arLin(i), Len(":53D:"))) = ":53D:"
                    SwiftDat.dat_53D = Mid(arLin(i), Len(":53D:") + 1)
                    Try
                        If Left(SwiftDat.dat_53D, 1) = "/" Then
                            arsxt = Split(Mid(SwiftDat.dat_53D, 2), vbCrLf)  ' sender correspondent
                            SwiftDat.correspSenderPI = arsxt(0)
                            For k = 1 To UBound(arsxt)
                                SwiftDat.correspSender = SwiftDat.correspSender & vbCrLf & arsxt(k)
                            Next
                            SwiftDat.correspSender = Mid(SwiftDat.correspSender, 3)
                        Else
                            arsxt = Split(SwiftDat.dat_53D, vbCrLf)  ' sender correspondent
                            For k = 0 To UBound(arsxt)
                                SwiftDat.correspSender = SwiftDat.correspSender & vbCrLf & arsxt(k)
                            Next
                            SwiftDat.correspSender = Mid(SwiftDat.correspSender, 3)
                        End If
                    Catch ex As Exception
                    End Try
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  53D: " & GetTagText(MtNum, "53D")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  53D: " & GetTagTEng(MtNum, "53D")
                    tstr = SwiftDat.correspSender : arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "      Nombre/Dir: " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "      Name/Addr.: " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                        End If
                    Next
                Case UCase(Left(arLin(i), Len(":54A:"))) = ":54A:"
                    SwiftDat.dat_54A = Mid(arLin(i), Len(":54A:") + 1)
                    Try
                        If Left(SwiftDat.dat_54A, 1) = "/" Then
                            arsxt = Split(Mid(SwiftDat.dat_54A, 2), vbCrLf)  ' ordering Customer
                            SwiftDat.correspReceiverPI = arsxt(0)
                            SwiftDat.correspReceiver = arsxt(1)
                        Else
                            SwiftDat.correspReceiver = SwiftDat.dat_54A
                        End If
                    Catch ex As Exception
                    End Try
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  54A: " & GetTagText(MtNum, "54A")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "            BIC : " & SwiftDat.correspReceiver
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  54A: " & GetTagTEng(MtNum, "54A")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "            BIC : " & SwiftDat.correspReceiver
                    tstr = GetBankFromBIC(SwiftDat.correspReceiver)
                    arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                        End If
                    Next
                Case UCase(Left(arLin(i), Len(":54B:"))) = ":54B:"
                    SwiftDat.dat_54B = Mid(arLin(i), Len(":54B:") + 1)
                    Try
                        If Left(SwiftDat.dat_54B, 1) = "/" Then
                            arsxt = Split(Mid(SwiftDat.dat_54B, 2), vbCrLf)  ' ordering Customer
                            SwiftDat.correspReceiverPI = arsxt(0)
                            SwiftDat.correspReceiver = arsxt(1)
                        Else
                            SwiftDat.correspReceiver = SwiftDat.dat_54B
                        End If
                    Catch ex As Exception
                    End Try
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  54B: " & GetTagText(MtNum, "54B")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  54B: " & GetTagTEng(MtNum, "54B")
                    tstr = SwiftDat.correspReceiver : arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "      Nombre/Dir: " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "      Name/Addr.: " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                        End If
                    Next
                Case UCase(Left(arLin(i), Len(":54D:"))) = ":54D:"
                    SwiftDat.dat_54D = Mid(arLin(i), Len(":54D:") + 1)
                    Try
                        If Left(SwiftDat.dat_54D, 1) = "/" Then
                            arsxt = Split(Mid(SwiftDat.dat_54D, 2), vbCrLf)  ' ordering Customer
                            SwiftDat.correspReceiverPI = arsxt(0)
                            For k = 1 To UBound(arsxt)
                                SwiftDat.correspReceiver = SwiftDat.correspReceiver & vbCrLf & arsxt(k)
                            Next
                            SwiftDat.correspReceiver = Mid(SwiftDat.correspReceiver, 3)
                        Else
                            arsxt = Split(SwiftDat.dat_54D, vbCrLf)  ' ordering Customer
                            For k = 0 To UBound(arsxt)
                                SwiftDat.correspReceiver = SwiftDat.correspReceiver & vbCrLf & arsxt(k)
                            Next
                            SwiftDat.correspReceiver = Mid(SwiftDat.correspReceiver, 3)
                        End If
                    Catch ex As Exception
                    End Try
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  54D: " & GetTagText(MtNum, "54D")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  54D: " & GetTagTEng(MtNum, "54D")
                    tstr = SwiftDat.correspReceiver : arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "      Nombre/Dir: " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "      Name/Addr.: " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                        End If
                    Next
                Case UCase(Left(arLin(i), Len(":55A:"))) = ":55A:"
                    SwiftDat.dat_55A = Mid(arLin(i), Len(":55A:") + 1)
                    Try
                        If Left(SwiftDat.dat_55A, 1) = "/" Then
                            arsxt = Split(Mid(SwiftDat.dat_55A, 2), vbCrLf)  ' ordering Customer
                            SwiftDat.terceraInstit = arsxt(1)
                        Else
                            SwiftDat.terceraInstit = SwiftDat.dat_55A
                        End If
                    Catch ex As Exception
                    End Try
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  55A: " & GetTagText(MtNum, "55A")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "            BIC : " & SwiftDat.terceraInstit
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  55A: " & GetTagTEng(MtNum, "55A")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "            BIC : " & SwiftDat.terceraInstit
                    tstr = GetBankFromBIC(SwiftDat.terceraInstit)
                    arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                        End If
                    Next
                Case UCase(Left(arLin(i), Len(":56A:"))) = ":56A:"
                    SwiftDat.dat_56A = Mid(arLin(i), Len(":56A:") + 1)
                    Try
                        If Left(SwiftDat.dat_56A, 1) = "/" Then
                            arsxt = Split(Mid(SwiftDat.dat_56A, 2), vbCrLf)  ' ordering Customer
                            SwiftDat.intermediariaInstitPI = arsxt(0)
                            SwiftDat.intermediariaInstit = arsxt(1)
                        Else
                            SwiftDat.intermediariaInstit = SwiftDat.dat_56A
                        End If
                    Catch ex As Exception
                    End Try
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  56A: " & GetTagText(MtNum, "56A")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "            BIC : " & SwiftDat.intermediariaInstit
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  56A: " & GetTagTEng(MtNum, "56A")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "            BIC : " & SwiftDat.intermediariaInstit
                    tstr = GetBankFromBIC(SwiftDat.intermediariaInstit)
                    arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                        End If
                    Next
                Case UCase(Left(arLin(i), Len(":56D:"))) = ":56D:"
                    SwiftDat.dat_56D = Mid(arLin(i), Len(":56D:") + 1)
                    Try
                        If Left(SwiftDat.dat_56D, 1) = "/" Then
                            arsxt = Split(Mid(SwiftDat.dat_56D, 2), vbCrLf)  ' ordering Customer
                            SwiftDat.intermediariaInstitPI = arsxt(0)
                            SwiftDat.intermediariaInstit = arsxt(1)
                            For k = 1 To UBound(arsxt)
                                SwiftDat.intermediariaInstit = SwiftDat.intermediariaInstit & vbCrLf & arsxt(k)
                            Next
                            SwiftDat.intermediariaInstit = Mid(SwiftDat.intermediariaInstit, 3)
                        Else
                            arsxt = Split(SwiftDat.dat_56D, vbCrLf)  ' ordering Customer
                            For k = 0 To UBound(arsxt)
                                SwiftDat.intermediariaInstit = SwiftDat.intermediariaInstit & vbCrLf & arsxt(k)
                            Next
                            SwiftDat.intermediariaInstit = Mid(SwiftDat.intermediariaInstit, 3)
                        End If
                    Catch ex As Exception
                    End Try
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  56D: " & GetTagText(MtNum, "56D")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  56D: " & GetTagTEng(MtNum, "56D")
                    tstr = SwiftDat.intermediariaInstit : arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "      Nombre/Dir: " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "      Name/Addr.: " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                        End If
                    Next
                Case UCase(Left(arLin(i), Len(":57A:"))) = ":57A:"
                    SwiftDat.dat_57A = Mid(arLin(i), Len(":57A:") + 1) ' accoun with institution 
                    Try
                        If Left(SwiftDat.dat_57A, 1) = "/" Then
                            arsxt = Split(Mid(SwiftDat.dat_57A, 2), vbCrLf)  ' ordering Customer
                            SwiftDat.financialInstitPI = arsxt(0)
                            SwiftDat.financialInstit = arsxt(1)
                        Else
                            SwiftDat.financialInstit = SwiftDat.dat_57A
                        End If
                    Catch ex As Exception
                    End Try
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  57A: " & GetTagText(MtNum, "57A")
                    If SwiftDat.financialInstitPI.Length > 0 Then
                        SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "            Acc : " & SwiftDat.financialInstitPI
                    End If
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "            BIC : " & SwiftDat.financialInstit
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  57A: " & GetTagTEng(MtNum, "57A")
                    If SwiftDat.financialInstitPI.Length > 0 Then
                        SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "            Acc : " & SwiftDat.financialInstitPI
                    End If
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "            BIC : " & SwiftDat.financialInstit
                    tstr = GetBankFromBIC(SwiftDat.financialInstit)
                    arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                        End If
                    Next
                Case UCase(Left(arLin(i), Len(":57B:"))) = ":57B:"
                    SwiftDat.dat_57B = Mid(arLin(i), Len(":57B:") + 1) ' accoun with institution 
                    Try
                        If Left(SwiftDat.dat_57B, 1) = "/" Then
                            arsxt = Split(Mid(SwiftDat.dat_57B, 2), vbCrLf)  ' ordering Customer
                            SwiftDat.financialInstitPI = arsxt(0)
                            SwiftDat.financialInstit = arsxt(1)
                        Else
                            SwiftDat.financialInstit = SwiftDat.dat_57B
                        End If
                    Catch ex As Exception
                    End Try
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  57B: " & GetTagText(MtNum, "57B")
                    If SwiftDat.financialInstitPI.Length > 0 Then
                        SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "            Acc : " & SwiftDat.financialInstitPI
                    End If
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  57B: " & GetTagTEng(MtNum, "57B")
                    If SwiftDat.financialInstitPI.Length > 0 Then
                        SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "            Acc : " & SwiftDat.financialInstitPI
                    End If
                    tstr = SwiftDat.financialInstit : arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "      Nombre/Dir: " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "      Name/Addr.: " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                        End If
                    Next
                Case UCase(Left(arLin(i), Len(":57D:"))) = ":57D:"
                    SwiftDat.dat_57D = Mid(arLin(i), Len(":57D:") + 1) ' accoun with institution 
                    Try
                        If Left(SwiftDat.dat_57D, 1) = "/" Then
                            arsxt = Split(Mid(SwiftDat.dat_57D, 2), vbCrLf)  ' ordering Customer
                            SwiftDat.financialInstitPI = arsxt(0)
                            For k = 1 To UBound(arsxt)
                                SwiftDat.financialInstit = SwiftDat.financialInstit & vbCrLf & arsxt(k)
                            Next
                            SwiftDat.financialInstit = Mid(SwiftDat.financialInstit, 3)
                        Else
                            arsxt = Split(SwiftDat.dat_57D, vbCrLf)  ' ordering Customer
                            For k = 0 To UBound(arsxt)
                                SwiftDat.financialInstit = SwiftDat.financialInstit & vbCrLf & arsxt(k)
                            Next
                            SwiftDat.financialInstit = Mid(SwiftDat.financialInstit, 3)
                        End If
                    Catch ex As Exception
                    End Try
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  57D: " & GetTagText(MtNum, "57D")
                    If SwiftDat.financialInstitPI.Length > 0 Then
                        SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "            Acc : " & SwiftDat.financialInstitPI
                    End If
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  57D: " & GetTagTEng(MtNum, "57D")
                    If SwiftDat.financialInstitPI.Length > 0 Then
                        SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "            Acc : " & SwiftDat.financialInstitPI
                    End If
                    tstr = SwiftDat.financialInstit : arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "      Nombre/Dir: " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "      Name/Addr.: " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                        End If
                    Next
                Case Left(arLin(i), Len(":58A:")) = ":58A:"
                    SwiftDat.dat_58A = Mid(arLin(i), Len(":58A:") + 1)
                    Try
                        If Left(SwiftDat.dat_58A, 1) = "/" Then
                            arsxt = Split(Mid(SwiftDat.dat_58A, 2), vbCrLf)  ' ordering Customer
                            SwiftDat.benefInstitPI = arsxt(0)
                            SwiftDat.benefInstit = arsxt(1)
                        Else
                            SwiftDat.benefInstit = SwiftDat.dat_58A
                        End If
                    Catch ex As Exception
                    End Try
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  58A: " & GetTagText(MtNum, "58A")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "            BIC : " & SwiftDat.benefInstit
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  58A: " & GetTagTEng(MtNum, "58A")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "            BIC : " & SwiftDat.benefInstit
                    tstr = GetBankFromBIC(SwiftDat.benefInstit)
                    arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                        End If
                    Next
                Case Left(arLin(i), Len(":58D:")) = ":58D:"
                    SwiftDat.dat_58D = Mid(arLin(i), Len(":58D:") + 1)
                    Try
                        If Left(SwiftDat.dat_58D, 1) = "/" Then
                            arsxt = Split(Mid(SwiftDat.dat_58D, 2), vbCrLf)  ' benef Instit
                            SwiftDat.benefInstitPI = arsxt(0)
                            For k = 1 To UBound(arsxt)
                                SwiftDat.benefInstit = SwiftDat.benefInstit & vbCrLf & arsxt(k)
                            Next
                            SwiftDat.benefInstit = Mid(SwiftDat.benefInstit, 3)
                        Else
                            arsxt = Split(SwiftDat.dat_58D, vbCrLf)  ' benef Instit
                            For k = 0 To UBound(arsxt)
                                SwiftDat.benefInstit = SwiftDat.benefInstit & vbCrLf & arsxt(k)
                            Next
                            SwiftDat.benefInstit = Mid(SwiftDat.benefInstit, 3)
                        End If
                    Catch ex As Exception
                    End Try
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  58D: " & GetTagText(MtNum, "58D")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  58D: " & GetTagTEng(MtNum, "58D")
                    tstr = SwiftDat.benefInstit : arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "      Nombre/Dir: " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "      Name/Addr.: " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                        End If
                    Next
                Case Left(arLin(i), Len(":59:")) = ":59:"
                    SwiftDat.dat_59 = Mid(arLin(i), Len(":59:") + 1)
                    Try
                        If Left(SwiftDat.dat_59, 1) = "/" Then
                            artxt = Split(SwiftDat.dat_59, vbCrLf)
                            If UBound(artxt) >= 1 Then
                                If InStr(artxt(0), "/") > 0 Then
                                    SwiftDat.benAccount = artxt(0)
                                    For k = 1 To UBound(artxt)
                                        SwiftDat.benName = SwiftDat.benName & vbCrLf & artxt(k)
                                    Next
                                    SwiftDat.benName = Mid(SwiftDat.benName, 3)
                                Else
                                    SwiftDat.benName = SwiftDat.dat_59 ' benef  Name
                                End If
                            Else
                                SwiftDat.benName = SwiftDat.dat_59 ' benef  Name
                            End If
                        Else
                            SwiftDat.benName = SwiftDat.dat_59 ' benef  Name
                        End If
                    Catch ex As Exception
                    End Try
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "   59: " & GetTagText(MtNum, "59")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Cuenta  : " & SwiftDat.benAccount
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "   59: " & GetTagTEng(MtNum, "59")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        Account : " & SwiftDat.benAccount
                    tstr = SwiftDat.benName : arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "      Nombre/Dir: " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "      Name/Addr.: " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                        End If
                    Next
                Case UCase(Left(arLin(i), Len(":59A:"))) = ":59A:"
                    SwiftDat.dat_59A = Mid(arLin(i), Len(":59A:") + 1)
                    Try
                        If Left(SwiftDat.dat_59A, 1) = "/" Then
                            artxt = Split(SwiftDat.dat_59A, vbCrLf)
                            If UBound(artxt) > 1 Then
                                arsxt = Split(artxt(0), "/")  ' benef 
                                SwiftDat.benAccount = arsxt(1)  ' benef Customer account
                                For k = 1 To UBound(artxt)
                                    SwiftDat.benName = SwiftDat.benName & vbCrLf & artxt(k)
                                Next
                                ' benef Customer account
                            Else
                            End If
                        Else
                            SwiftDat.benName = SwiftDat.dat_59A  ' benef Customer account
                        End If
                    Catch ex As Exception
                    End Try
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  59A: " & GetTagText(MtNum, "59A")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Cuenta  :" & SwiftDat.benAccount
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  59A: " & GetTagTEng(MtNum, "59A")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        Account :" & SwiftDat.benAccount
                    tstr = SwiftDat.benName : arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "      Nombre/Dir: " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "      Name/Addr.: " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                            SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                        End If
                    Next
                Case Left(arLin(i), Len(":70:")) = ":70:"
                    SwiftDat.dat_70 = Mid(arLin(i), Len(":70:") + 1)
                    SwiftDat.remitInfo = SwiftDat.dat_70
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "   70: " & GetTagText(MtNum, "70")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "   70: " & GetTagTEng(MtNum, "70")
                    tstr = SwiftDat.remitInfo : arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & arsxt(k)
                        SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        " & arsxt(k)
                    Next
                Case UCase(Left(arLin(i), Len(":71A:"))) = ":71A:"
                    SwiftDat.dat_71A = Mid(arLin(i), Len(":71A:") + 1)
                    SwiftDat.cargoTo = SwiftDat.dat_71A
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  71A: " & GetTagText(MtNum, "71A")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & SwiftDat.cargoTo
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  71A: " & GetTagTEng(MtNum, "71A")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        " & SwiftDat.cargoTo
                Case UCase(Left(arLin(i), Len(":71B:"))) = ":71B:"
                    SwiftDat.dat_71B = Mid(arLin(i), Len(":71B:") + 1)
                    SwiftDat.cargDeducted = SwiftDat.dat_71B
                    arsxt = Split(SwiftDat.cargDeducted, vbCrLf)
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  71B: " & GetTagText(MtNum, "71B")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  71B: " & GetTagTEng(MtNum, "71B")
                    For k = 0 To UBound(arsxt)
                        SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                        SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "                  " & arsxt(k)
                    Next
                Case UCase(Left(arLin(i), Len(":71F:"))) = ":71F:"
                    tstr = Mid(arLin(i), Len(":71F:") + 1)
                    If Trim(SwiftDat.dat_71F) <> "" Then
                        SwiftDat.dat_71F = SwiftDat.dat_71F & "/" & tstr
                        SwiftDat.senderChargCurr = SwiftDat.senderChargCurr & "/" & Left(tstr, 3)
                        SwiftDat.senderChargAmou = SwiftDat.senderChargAmou & "/" & Mid(tstr, 4)
                    Else
                        SwiftDat.dat_71F = tstr
                        SwiftDat.senderChargCurr = Left(tstr, 3)
                        SwiftDat.senderChargAmou = Mid(tstr, 4)
                    End If
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  71F: " & GetTagText(MtNum, "71F")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Moneda  : " & Left(tstr, 3)
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Monto   : " & Mid(tstr, 4)
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  71F: " & GetTagTEng(MtNum, "71F")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        Currency: " & Left(tstr, 3)
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        Amount  : " & Mid(tstr, 4)
                Case UCase(Left(arLin(i), Len(":71G:"))) = ":71G:"
                    tstr = Mid(arLin(i), Len(":71G:") + 1)
                    If Trim(SwiftDat.dat_71G) <> "" Then
                        SwiftDat.dat_71G = SwiftDat.dat_71G & "/" & tstr
                        SwiftDat.recverChargCurr = SwiftDat.recverChargCurr & "/" & Left(tstr, 3)
                        SwiftDat.recverChargAmou = SwiftDat.recverChargAmou & "/" & Mid(tstr, 4)
                    Else
                        SwiftDat.dat_71G = tstr
                        SwiftDat.recverChargCurr = Left(tstr, 3)
                        SwiftDat.recverChargAmou = Mid(tstr, 4)
                    End If
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  71G: " & GetTagText(MtNum, "71G")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Moneda  : " & Left(tstr, 3)
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Monto   : " & Mid(tstr, 4)
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  71G: " & GetTagTEng(MtNum, "71G")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        Currency: " & Left(tstr, 3)
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        Amount  : " & Mid(tstr, 4)
                Case Left(arLin(i), Len(":72:")) = ":72:"
                    SwiftDat.dat_72 = Mid(arLin(i), Len(":72:") + 1)
                    SwiftDat.send2recvInfo = SwiftDat.dat_72
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "   72: " & GetTagText(MtNum, "72")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "   72: " & GetTagTEng(MtNum, "72")
                    tstr = SwiftDat.send2recvInfo : arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "         " & arsxt(k)
                        SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "         " & arsxt(k)
                    Next
                Case UCase(Left(arLin(i), Len(":77B:"))) = ":77B:"
                    SwiftDat.dat_77B = Mid(arLin(i), Len(":77B:") + 1)
                    SwiftDat.regulReport = SwiftDat.dat_77B
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  77B: " & GetTagText(MtNum, "77B")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  77B: " & GetTagTEng(MtNum, "77B")
                    tstr = SwiftDat.regulReport : arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & arsxt(k)
                        SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        " & arsxt(k)
                    Next
                Case UCase(Left(arLin(i), Len(":77T:"))) = ":77T:"
                    SwiftDat.dat_77T = Mid(arLin(i), Len(":77T:") + 1)
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  77T: " & GetTagText(MtNum, "77T")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & SwiftDat.dat_77T
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "  77T: " & GetTagTEng(MtNum, "77T")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        " & SwiftDat.dat_77T
                Case UCase(Left(arLin(i), Len(":78:"))) = ":78:"
                    SwiftDat.dat_78 = Mid(arLin(i), Len(":78:") + 1)
                    SwiftDat.msgtoBank = SwiftDat.dat_78
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "   78: " & GetTagText(MtNum, "78")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "   78: " & GetTagTEng(MtNum, "78")
                    arsxt = Split(SwiftDat.msgtoBank, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & arsxt(k)
                        SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        " & arsxt(k)
                    Next
                Case UCase(Left(arLin(i), Len(":79:"))) = ":79:"
                    SwiftDat.dat_79 = Mid(arLin(i), Len(":79:") + 1)
                    SwiftDat.msgExplanation = SwiftDat.dat_79
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "   79: " & GetTagText(MtNum, "79")
                    SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "   79: " & GetTagTEng(MtNum, "79")
                    tstr = SwiftDat.msgExplanation : arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & arsxt(k)
                        SwiftDat.MsgTEng = SwiftDat.MsgTEng & vbCrLf & "        " & arsxt(k)
                    Next
                Case Left(arLin(i), Len("{1:F")) = "{1:F"
                    k = InStr(arLin(i), "{2:")
                    If k > 0 Then
                        SwiftDat.tipo_mensaje = Mid(arLin(i), k + 4, 3)
                        SwiftDat.Fecha_Ingreso = fechafile
                        If SwiftDat.tipo_mensaje = "103" Then
                            'SwiftDat.Hora_ingreso = Mid(arLin(i), k + 7, 4)
                            SwiftDat.Corresponsal = Mid(arLin(i), k + 7, 12)
                        ElseIf SwiftDat.tipo_mensaje = "202" Then
                            'SwiftDat.Hora_ingreso = Mid(arLin(i), k + 7, 4)
                            SwiftDat.Corresponsal = Mid(arLin(i), k + 7, 12)
                        ElseIf SwiftDat.tipo_mensaje = "910" Then
                            'SwiftDat.Hora_ingreso = Mid(arLin(i), k + 7, 4)
                            SwiftDat.Corresponsal = Mid(arLin(i), k + 7, 12)
                        ElseIf SwiftDat.tipo_mensaje = "799" Then
                            'SwiftDat.Hora_ingreso = Mid(arLin(i), k + 7, 4)
                            SwiftDat.Corresponsal = Mid(arLin(i), k + 7, 12)
                        Else
                            'SwiftDat.Hora_ingreso = Mid(arLin(i), k + 7, 4)
                            SwiftDat.Corresponsal = Mid(arLin(i), k + 7, 12)
                        End If
                        SwiftDat.MsgReceiver = SwiftDat.Corresponsal
                        MtNum = CInt(SwiftDat.tipo_mensaje)
                        k = InStr(arLin(i), "{1:")
                        j = InStr(k, arLin(i), "}")
                        linea = Mid(arLin(i), k + 3, (j - k) - 3)
                        tstr = Mid(linea, 4, 12)
                        tstr = Left(tstr, 8) & Mid(tstr, 10, 3)
                        SwiftDat.MsgSender = tstr
                    End If
                    k = InStr(arLin(i), "{1:") : j = InStr(k, arLin(i), "}")
                    linea = Mid(arLin(i), k + 3, (j - k) - 3)
                    SwiftDat.ISN_mensaje = Right(linea, 6)
                    tstr = Mid(linea, 4, 12)
                    tstr = Left(tstr, 8) & Mid(tstr, 10, 3)
                    'SwiftDat.MsgReceiver = tstr
                    SwiftDat.MsgMUR = ""
                    Try
                        k = InStr(10, arLin(i), "{3:") : j = InStr(k, arLin(i), "}")
                        linea = Mid(arLin(i), k + 3, (j - k) - 3)
                        k = InStr(linea, ":")
                        SwiftDat.MsgMUR = Mid(linea, k + 1)
                    Catch ex As Exception
                    End Try
                    SwiftDat.MsgHead = "Swift OUTPUT : FIN " & SwiftDat.tipo_mensaje & " " & GetTagText(CInt(SwiftDat.tipo_mensaje), "--")
                    SwiftDat.MsgHead = SwiftDat.MsgHead & vbCrLf & "Envia        : " & SwiftDat.MsgSender
                    SwiftDat.MsgHEng = "Swift OUTPUT : FIN " & SwiftDat.tipo_mensaje & " " & GetTagTEng(CInt(SwiftDat.tipo_mensaje), "--")
                    SwiftDat.MsgHEng = SwiftDat.MsgHEng & vbCrLf & "Sender       : " & SwiftDat.MsgSender
                    tstr = GetBankFromBIC(SwiftDat.MsgSender)
                    artxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(artxt)
                        SwiftDat.MsgHead = SwiftDat.MsgHead & vbCrLf & "         " & artxt(k)
                        SwiftDat.MsgHEng = SwiftDat.MsgHEng & vbCrLf & "         " & artxt(k)
                    Next
                    SwiftDat.MsgHead = SwiftDat.MsgHead & vbCrLf & "Recibe       : " & SwiftDat.MsgReceiver
                    SwiftDat.MsgHEng = SwiftDat.MsgHEng & vbCrLf & "Receiver     : " & SwiftDat.MsgReceiver
                    tstr = GetBankFromBIC(SwiftDat.MsgReceiver)
                    artxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(artxt)
                        SwiftDat.MsgHead = SwiftDat.MsgHead & vbCrLf & "         " & artxt(k)
                        SwiftDat.MsgHEng = SwiftDat.MsgHEng & vbCrLf & "         " & artxt(k)
                    Next
                    If SwiftDat.MsgMUR.Length > 0 Then
                        SwiftDat.MsgHead = SwiftDat.MsgHead & vbCrLf & "MUR :  " & SwiftDat.MsgMUR
                        SwiftDat.MsgHEng = SwiftDat.MsgHEng & vbCrLf & "MUR :  " & SwiftDat.MsgMUR
                    End If
                Case Else 'dat_Tail
                    SwiftDat.dat_Tail = SwiftDat.dat_Tail & vbCrLf & arLin(i)
                    'SwiftMT.dat_Tail = IIf(Left(SwiftDat.dat_Tail, 2) = vbCrLf, Mid(SwiftDat.dat_Tail, 3), SwiftDat.dat_Tail)
            End Select
            'rConsola("arLin(" & i & ") > " & arLin(i))
        Next
        ProcFile = True
    End Function
    Function SaveDataSwift(ByVal SwiftData As SwiftsJob.SwiftData, ByVal errMessage As String) As Boolean
        Dim tmomey As Decimal
        Dim sim As String = Mid("" & 1.1, 2, 1)
        Dim msg_799 As String, OkDone As Boolean

        SaveDataSwift = True

        _c.Informa("-----------------------------------")
        _c.Informa(":Mensaje Swift Tipo MT" & SwiftData.tipo_mensaje) '& "----------------------------")
        _c.Informa("     -- ISN_mensaje > " & SwiftData.ISN_mensaje)
        _c.Informa("     -- tipo_mensaje > " & SwiftData.tipo_mensaje)
        _c.Informa("     -- Data Tail > " & SwiftData.dat_Tail)
        'SaveDataSwift = True
        'Exit Function
        'If SwiftData.tipo_mensaje <> "700" Then
        '    SwiftData.tipo_mensaje = "NOPROC_" & SwiftData.tipo_mensaje
        'End If
        If SwiftData.tipo_mensaje = "103" Then
            Dim conx As New SqlConnection(_cnxstr)
            Dim ermess
            Dim strSwf As String = "INSERT INTO T_ECX_SwiftMT103_Sale"
            strSwf = strSwf & " (F_ISN, F_FechaIngreso, F_FechaProceso, F_tipo_mensaje, F_Hora_ingreso, F_Corresponsal, " & _
                        " F_timeType, F_timeTime, F_timeOffs, " & _
                        " F_msgReference, F_bankOperatCode, F_instructCodeCode, " & _
                        " F_instructCodeRefr, F_transacType, F_Valuta, F_Moneda, F_Monto, " & _
                        " F_instructCurr, F_instructAmount, F_exchRate, F_orderingCusAcc, " & _
                        " F_orderingCusName, F_origBank, F_correspSender, F_correspReceiver, " & _
                        " F_terceraInstit, F_intermediariaInstit, F_financialInstit, " & _
                        " F_benAccount, F_benName, F_remitInfo, F_cargoTo, " & _
                        " F_senderChargCurr, F_senderChargAmou, F_recverChargCurr, F_recverChargAmou, " & _
                        " F_send2recvInfo, F_regulReport, F_SwiftMessage, F_DataTail, F_OrigFile, " & _
                        " F_SwiftSender, F_SwiftReceiver, F_SwiftMUR, F_SwiftHead, F_SwiftText, F_SwiftTrailer, " & _
                        " F_SwiftHeng, F_SwiftTeng)"
            'VALUES     (3, '3', '3', '3', '3', '3', '3', '3', '3', '3', '3', 3, '3', 3, '3', 3, 3, '3', '3', '3', '3', '3', '3', '3', '3', '3', '3', '3', '3', '3', '3', 3, '3', 3, '3', '3', '3')
            strSwf = strSwf & " VALUES (@ISN, @FECING, GETDATE(), @TPMSG, @HRING, @CORRESP, " & _
                        " @TMTIP, @TMTIM, @TMOFF, " & _
                        " @MSGREF, @BKOPCOD, @ITRCDCD, " & _
                        " @ITRCDRF, @TRNTIP, @VALUTA, @MONEDA, @MONTO, " & _
                        " @ITRCURR, @ITRMTO, @TPOCAM, @CTAORDCUS, " & _
                        " @NOMORDCUS, @ORDBNK, @CORRSENDR, @CORRECVR, " & _
                        " @TERCINST, @INTERINST, @FINANINST, " & _
                        " @BENACO, @BENNOM, @REMINFO, @CARGOTO, " & _
                        " @SNDCHRCUR, @SNDCHRAMO, @RCVCHRCUR, @RCVCHRAMO, " & _
                        " @SNDRECINF, @REGUREP, @SWIFTMSG, @DATTAIL, @FILEDAT, " & _
                        " @SWFSNDR, @SWFRCVR, @SWFMUR, @SWFHED, @SWFTXT, @SWFTRAI, @SWFHEG, @SWFTEG)"

            Dim swftComm As New SqlCommand(strSwf, conx)
            swftComm.Parameters.Add("@ISN", SqlDbType.Int).Value = CInt(SwiftData.ISN_mensaje)
            swftComm.Parameters.Add("@FECING", SqlDbType.DateTime).Value = SwiftData.Fecha_Ingreso
            swftComm.Parameters.Add("@TPMSG", SqlDbType.Int).Value = CInt(SwiftData.tipo_mensaje)
            swftComm.Parameters.Add("@HRING", SqlDbType.VarChar, 5).Value = Left(SwiftData.Hora_ingreso, 2) & ":" & Mid(SwiftData.Hora_ingreso, 3)
            swftComm.Parameters.Add("@CORRESP", SqlDbType.VarChar, 50).Value = SwiftData.Corresponsal
            swftComm.Parameters.Add("@TMTIP", SqlDbType.VarChar, 8).Value = SwiftData.timeType
            swftComm.Parameters.Add("@TMTIM", SqlDbType.VarChar, 5).Value = Left(SwiftData.timeTime, 2) & ":" & Mid(SwiftData.timeTime, 3)
            swftComm.Parameters.Add("@TMOFF", SqlDbType.VarChar, 5).Value = Left(SwiftData.timeOffs, 2) & ":" & Mid(SwiftData.timeOffs, 3)
            swftComm.Parameters.Add("@MSGREF", SqlDbType.VarChar, 16).Value = SwiftData.msgReference
            swftComm.Parameters.Add("@BKOPCOD", SqlDbType.VarChar, 4).Value = SwiftData.bankOperatCode
            swftComm.Parameters.Add("@ITRCDCD", SqlDbType.VarChar, 4).Value = SwiftData.instructCodeCode
            swftComm.Parameters.Add("@ITRCDRF", SqlDbType.VarChar, 30).Value = SwiftData.instructCodeRefr
            swftComm.Parameters.Add("@TRNTIP", SqlDbType.VarChar, 3).Value = SwiftData.transacType
            swftComm.Parameters.Add("@VALUTA", SqlDbType.DateTime).Value = toDate(SwiftData.Valuta)
            swftComm.Parameters.Add("@MONEDA", SqlDbType.VarChar, 3).Value = SwiftData.Moneda
            If Trim("" & SwiftData.Monto) = "" Then
                tmomey = 0
            Else : tmomey = CDec(SwiftData.Monto.Replace(",", sim))
            End If
            swftComm.Parameters.Add("@MONTO", SqlDbType.Money).Value = tmomey
            swftComm.Parameters.Add("@ITRCURR", SqlDbType.VarChar, 3).Value = SwiftData.instructCurr
            If Trim("" & SwiftData.instructAmount) = "" Then
                tmomey = 0
            Else : tmomey = CDec(SwiftData.instructAmount.Replace(",", sim))
            End If
            swftComm.Parameters.Add("@ITRMTO", SqlDbType.Money).Value = tmomey
            If Trim("" & SwiftData.exchRate) = "" Then
                tmomey = 0
            Else : tmomey = CDec(SwiftData.exchRate.Replace(",", sim))
            End If
            swftComm.Parameters.Add("@TPOCAM", SqlDbType.Money).Value = tmomey
            swftComm.Parameters.Add("@CTAORDCUS", SqlDbType.VarChar, 34).Value = SwiftData.orderingCusAcc
            swftComm.Parameters.Add("@NOMORDCUS", SqlDbType.VarChar, 300).Value = SwiftData.orderingCusName
            swftComm.Parameters.Add("@ORDBNK", SqlDbType.VarChar, 300).Value = SwiftData.origBank
            swftComm.Parameters.Add("@CORRSENDR", SqlDbType.VarChar, 300).Value = SwiftData.correspSender
            swftComm.Parameters.Add("@CORRECVR", SqlDbType.VarChar, 300).Value = SwiftData.correspReceiver
            swftComm.Parameters.Add("@TERCINST", SqlDbType.VarChar, 300).Value = SwiftData.terceraInstit
            swftComm.Parameters.Add("@INTERINST", SqlDbType.VarChar, 300).Value = SwiftData.intermediariaInstit
            swftComm.Parameters.Add("@FINANINST", SqlDbType.VarChar, 300).Value = SwiftData.financialInstit
            swftComm.Parameters.Add("@BENACO", SqlDbType.VarChar, 34).Value = SwiftData.benAccount
            swftComm.Parameters.Add("@BENNOM", SqlDbType.VarChar, 300).Value = SwiftData.benName
            swftComm.Parameters.Add("@REMINFO", SqlDbType.VarChar, 300).Value = SwiftData.remitInfo
            swftComm.Parameters.Add("@CARGOTO", SqlDbType.VarChar, 3).Value = SwiftData.cargoTo
            swftComm.Parameters.Add("@SNDCHRCUR", SqlDbType.VarChar, 50).Value = SwiftData.senderChargCurr
            'If Trim("" & SwiftData.senderChargAmou) = "" Then
            '    tmomey = 0
            'Else : tmomey = CDec(SwiftData.senderChargAmou.Replace(",", sim))
            'End If
            swftComm.Parameters.Add("@SNDCHRAMO", SqlDbType.VarChar, 50).Value = SwiftData.senderChargAmou
            swftComm.Parameters.Add("@RCVCHRCUR", SqlDbType.VarChar, 50).Value = SwiftData.recverChargCurr
            'If Trim("" & SwiftData.recverChargAmou) = "" Then
            '    tmomey = 0
            'Else : tmomey = CDec(SwiftData.recverChargAmou.Replace(",", sim))
            'End If
            swftComm.Parameters.Add("@RCVCHRAMO", SqlDbType.VarChar, 50).Value = SwiftData.recverChargAmou
            swftComm.Parameters.Add("@SNDRECINF", SqlDbType.VarChar, 300).Value = SwiftData.send2recvInfo
            swftComm.Parameters.Add("@REGUREP", SqlDbType.VarChar, 300).Value = SwiftData.regulReport
            swftComm.Parameters.Add("@SWIFTMSG", SqlDbType.Text).Value = SwiftData.dat_All
            swftComm.Parameters.Add("@DATTAIL", SqlDbType.Text).Value = SwiftData.dat_Tail
            swftComm.Parameters.Add("@FILEDAT", SqlDbType.VarChar, 150).Value = SwiftData.OrigFile
            swftComm.Parameters.Add("@SWFSNDR", SqlDbType.VarChar, 12).Value = SwiftData.MsgSender
            swftComm.Parameters.Add("@SWFRCVR", SqlDbType.VarChar, 12).Value = SwiftData.MsgReceiver
            swftComm.Parameters.Add("@SWFMUR", SqlDbType.VarChar, 150).Value = SwiftData.MsgMUR
            swftComm.Parameters.Add("@SWFHED", SqlDbType.VarChar, 500).Value = SwiftData.MsgHead
            swftComm.Parameters.Add("@SWFTXT", SqlDbType.Text).Value = SwiftData.MsgText
            swftComm.Parameters.Add("@SWFTRAI", SqlDbType.VarChar, 200).Value = SwiftData.MsgTrailer
            swftComm.Parameters.Add("@SWFHEG", SqlDbType.VarChar, 500).Value = SwiftData.MsgHEng
            swftComm.Parameters.Add("@SWFTEG", SqlDbType.Text).Value = SwiftData.MsgTEng
            Try
                If conx.State = ConnectionState.Closed Then conx.Open()
                swftComm.ExecuteNonQuery()
                _c.Informa("  SWIFT :  " & SwiftData.tipo_mensaje & " -> " & SwiftData.msgReference & " Ingresado a BBDD")
                SaveDataSwift = True
                AvisaDestinatarioComex(SwiftData) '.MsgReceiver, SwiftData.orderingCusName, SwiftData.msgReference, SwiftData.benName, SwiftData, SwiftData.MsgHEng, SwiftData.MsgTEng)
            Catch ex As Exception
                _c.Informa("  Error: " & ex.ToString)
                SaveDataSwift = False
            Finally
                'swftComm.Dispose()
            End Try
        ElseIf SwiftData.tipo_mensaje = "202" Then
            Dim conx As New SqlConnection(_cnxstr)
            Dim ermess
            Dim strSwf As String = "INSERT INTO T_ECX_SwiftMT202_Sale"
            strSwf = strSwf & " (F_ISN, F_FechaIngreso, F_tipo_mensaje, F_Hora_ingreso, F_Corresponsal, " & _
                        " F_timeType, F_timeTime, F_timeOffs,  " & _
                        " F_msgReference, F_relReference, F_Valuta, F_Moneda, " & _
                        " F_Monto, F_origBank, F_correspSender, F_correspReceiver, " & _
                        " F_intermediariaInstit, F_financialInstit, F_benefInstit, " & _
                        " F_send2recvInfo, F_regulReport, " & _
                        " F_SwiftMessage, F_DataTail, F_OrigFile, " & _
                        " F_SwiftSender, F_SwiftReceiver, F_SwiftMUR, F_SwiftHead, F_SwiftText, F_SwiftTrailer, " & _
                        " F_SwiftHeng, F_SwiftTeng)"
            'VALUES     (3, '3', '3', '3', '3', '3', '3', '3', '3', '3', '3', 3, '3', 3, '3', 3, 3, '3', '3', '3', '3', '3', '3', '3', '3', '3', '3', '3', '3', '3', '3', 3, '3', 3, '3', '3', '3')
            strSwf = strSwf & " VALUES (@ISN, @FECING, @TPMSG, @HRING, @CORRESP, " & _
                        " @TMTIP, @TMTIM, @TMOFF, " & _
                        " @MSGREF, @RELREF, @VALUTA, @MONEDA, @MONTO, " & _
                        " @ORDBNK, @CORRSENDR, @CORRECVR, " & _
                        " @INTERINST, @FINANINST, @BENFINST, " & _
                        " @SNDRECINF, @REGUREP, @SWIFTMSG, @DATTAIL, @FILEDAT, " & _
                        " @SWFSNDR, @SWFRCVR, @SWFMUR, @SWFHED, @SWFTXT, @SWFTRAI, @SWFHEG, @SWFTEG)"

            Dim swftComm As New SqlCommand(strSwf, conx)
            swftComm.Parameters.Add("@ISN", SqlDbType.Int).Value = CInt(SwiftData.ISN_mensaje)
            swftComm.Parameters.Add("@FECING", SqlDbType.DateTime).Value = SwiftData.Fecha_Ingreso
            swftComm.Parameters.Add("@TPMSG", SqlDbType.Int).Value = CInt(SwiftData.tipo_mensaje)
            swftComm.Parameters.Add("@HRING", SqlDbType.VarChar).Value = Left(SwiftData.Hora_ingreso, 2) & ":" & Mid(SwiftData.Hora_ingreso, 3)
            swftComm.Parameters.Add("@CORRESP", SqlDbType.VarChar).Value = SwiftData.Corresponsal
            swftComm.Parameters.Add("@TMTIP", SqlDbType.VarChar).Value = SwiftData.timeType
            swftComm.Parameters.Add("@TMTIM", SqlDbType.VarChar).Value = Left(SwiftData.timeTime, 2) & ":" & Mid(SwiftData.timeTime, 3)
            swftComm.Parameters.Add("@TMOFF", SqlDbType.VarChar).Value = Left(SwiftData.timeOffs, 2) & ":" & Mid(SwiftData.timeOffs, 3)
            swftComm.Parameters.Add("@MSGREF", SqlDbType.VarChar).Value = SwiftData.msgReference
            swftComm.Parameters.Add("@RELREF", SqlDbType.VarChar).Value = SwiftData.relReference
            swftComm.Parameters.Add("@VALUTA", SqlDbType.DateTime).Value = toDate(SwiftData.Valuta)
            swftComm.Parameters.Add("@MONEDA", SqlDbType.VarChar, 3).Value = SwiftData.Moneda
            If Trim("" & SwiftData.Monto) = "" Then
                tmomey = 0
            Else : tmomey = CDec(SwiftData.Monto.Replace(",", sim))
            End If
            swftComm.Parameters.Add("@MONTO", SqlDbType.Money).Value = tmomey
            swftComm.Parameters.Add("@ORDBNK", SqlDbType.VarChar).Value = SwiftData.origBank
            swftComm.Parameters.Add("@CORRSENDR", SqlDbType.VarChar).Value = SwiftData.correspSender
            swftComm.Parameters.Add("@CORRECVR", SqlDbType.VarChar).Value = SwiftData.correspReceiver
            swftComm.Parameters.Add("@INTERINST", SqlDbType.VarChar).Value = SwiftData.intermediariaInstit
            swftComm.Parameters.Add("@FINANINST", SqlDbType.VarChar).Value = SwiftData.financialInstit
            swftComm.Parameters.Add("@BENFINST", SqlDbType.VarChar).Value = SwiftData.benefInstit
            swftComm.Parameters.Add("@SNDRECINF", SqlDbType.VarChar).Value = SwiftData.send2recvInfo
            swftComm.Parameters.Add("@REGUREP", SqlDbType.VarChar).Value = SwiftData.regulReport
            swftComm.Parameters.Add("@SWIFTMSG", SqlDbType.Text).Value = SwiftData.dat_All
            swftComm.Parameters.Add("@DATTAIL", SqlDbType.Text).Value = SwiftData.dat_Tail
            swftComm.Parameters.Add("@FILEDAT", SqlDbType.Text).Value = SwiftData.OrigFile
            swftComm.Parameters.Add("@SWFSNDR", SqlDbType.VarChar).Value = SwiftData.MsgSender
            swftComm.Parameters.Add("@SWFRCVR", SqlDbType.VarChar).Value = SwiftData.MsgReceiver
            swftComm.Parameters.Add("@SWFMUR", SqlDbType.VarChar).Value = SwiftData.MsgMUR
            swftComm.Parameters.Add("@SWFHED", SqlDbType.VarChar).Value = SwiftData.MsgHead
            swftComm.Parameters.Add("@SWFTXT", SqlDbType.Text).Value = SwiftData.MsgText
            swftComm.Parameters.Add("@SWFTRAI", SqlDbType.VarChar).Value = SwiftData.MsgTrailer
            swftComm.Parameters.Add("@SWFHEG", SqlDbType.VarChar).Value = SwiftData.MsgHEng
            swftComm.Parameters.Add("@SWFTEG", SqlDbType.Text).Value = SwiftData.MsgTEng

            Try
                If conx.State = ConnectionState.Closed Then conx.Open()
                swftComm.ExecuteNonQuery()
                _c.Informa("  SWIFT :  " & SwiftData.tipo_mensaje & " -> " & SwiftData.msgReference & " Ingresado a BBDD")
                SaveDataSwift = True
            Catch ex As Exception
                _c.Informa("  Error: " & ex.ToString)
                SaveDataSwift = False
            Finally
                'swftComm.Dispose()
            End Try
        ElseIf SwiftData.tipo_mensaje = "700" Then
            Dim conx As New SqlConnection(_cnxstr)
            Dim ermess
            Dim strSwf As String
            strSwf = "" 'IF NOT exists(SELECT 1 FROM T_ECX_SwiftMT700_Sale WHERE F_docCredNumber = @MSGREF)"
            strSwf = strSwf & " INSERT INTO T_ECX_SwiftMT700_Sale " & _
                        " (F_ISN, F_FechaIngreso, F_FechaProceso, F_tipo_mensaje, " & _
                        " F_Hora_ingreso, F_Corresponsal, " & _
                        " F_docCredNumber, F_ref2PreAdvice, F_seqIndex, F_seqTotal, " & _
                        " F_typeofCredit, F_fechIssue, F_dateofExpire, F_placeofExpire, " & _
                        " F_applicBank, F_applicant, F_benAccount, F_benName, " & _
                        " F_MontoCredCant, F_MontoCredMone, F_percTolMonto1, F_percTolMonto2, " & _
                        " F_MaxCredCode, F_AdicMontoMemo, F_availableWith, F_draftsAt, F_draweeBank, " & _
                        " F_mixPayDetail, F_deferrPayDetail, F_partialShipp, F_transShipp, F_DispatchOpt,  " & _
                        " F_FinalDest, F_lateDateofShipp, F_shippPeriod, F_descripGoods, " & _
                        " F_docsRequired, F_additCondics, F_cargDeducted, F_periodPresent,  " & _
                        " F_appl_rules, F_confirmInstruct, F_reimbursingBank, F_instruct2Bank, " & _
                        " F_adviserBank, F_send2recvInfo, F_SwiftMessage, F_DataTail, " & _
                        " F_OrigFile, F_SwiftSender, F_SwiftReceiver, F_SwiftMUR, " & _
                        " F_SwiftHead, F_SwiftText, F_SwiftTrailer, " & _
                        " F_ptoOrigen, F_ptoDestino, F_SwiftHeng, F_SwiftTeng)"
            strSwf = strSwf & " VALUES (@ISN, @FECING, GETDATE(), @TPMSG, @HRING, @CORRESP, " & _
                        " @MSGREF, @RELREF, @SEQINDX, @SEQTOTL, " & _
                        " @TYPECRED, @FECISS, @FECEXPI, @LUGEXPI, " & _
                        " @APPLICBANK, @APPLIC, @BENACCO, @BENEF, " & _
                        " @CREDMONTO, @CREDMONEDA, @PERCTOT1, @PERCTOT2, " & _
                        " @MAXCREDMOD, @ADICMONTMEMO, @AVAILWITH, @DRAFTAT, @DRAWEEBANK, " & _
                        " @MIXPAYDETL, @DEFERPYDET, @PARTSHIPP, @TRANSSHIPP, @DESPCHOPT, " & _
                        " @FINLDEST, @LASTFECEM, @SHIPPER, @DESCRGOOD, " & _
                        " @DOCSREQUI, @ADICCONDS, @CARGDED, @PERIPRESEN, " & _
                        " @APPLRULES, @CONFRINSTR, @REIMBANK, @INSTR2BANK, " & _
                        " @ADVISBANK, @SNDRECINF, " & _
                        " @SWIFTMSG, @DATTAIL, @FILEDAT, " & _
                        " @SWFSNDR, @SWFRCVR, @SWFMUR, @SWFHED, @SWFTXT, @SWFTRAI, @PTORIG, @PTDEST, @SWFHEG, @SWFTEG)"

            Dim swftComm As New SqlCommand(strSwf, conx)
            swftComm.Parameters.Add("@ISN", SqlDbType.Int).Value = CInt(SwiftData.ISN_mensaje)
            swftComm.Parameters.Add("@FECING", SqlDbType.DateTime).Value = SwiftData.Fecha_Ingreso
            swftComm.Parameters.Add("@TPMSG", SqlDbType.Int).Value = CInt(SwiftData.tipo_mensaje)
            swftComm.Parameters.Add("@HRING", SqlDbType.VarChar, 5).Value = Left(SwiftData.Hora_ingreso, 2) & ":" & Mid(SwiftData.Hora_ingreso, 3)
            swftComm.Parameters.Add("@CORRESP", SqlDbType.VarChar, 50).Value = SwiftData.Corresponsal
            swftComm.Parameters.Add("@MSGREF", SqlDbType.VarChar, 16).Value = SwiftData.msgReference
            swftComm.Parameters.Add("@RELREF", SqlDbType.VarChar, 16).Value = SwiftData.issuingBankReference
            swftComm.Parameters.Add("@SEQINDX", SqlDbType.Int).Value = CInt(SwiftData.seqIndex)
            swftComm.Parameters.Add("@SEQTOTL", SqlDbType.Int).Value = CInt(SwiftData.seqTotal)
            swftComm.Parameters.Add("@TYPECRED", SqlDbType.VarChar, 24).Value = SwiftData.typeofCredit
            swftComm.Parameters.Add("@FECISS", SqlDbType.DateTime).Value = toDate(SwiftData.dateofIssue)
            swftComm.Parameters.Add("@FECEXPI", SqlDbType.DateTime).Value = toDate(SwiftData.dateofExpire)
            swftComm.Parameters.Add("@LUGEXPI", SqlDbType.VarChar, 30).Value = SwiftData.placeofExpire
            swftComm.Parameters.Add("@APPLICBANK", SqlDbType.VarChar, 200).Value = SwiftData.applicBank
            swftComm.Parameters.Add("@APPLIC", SqlDbType.VarChar, 150).Value = SwiftData.applicant
            swftComm.Parameters.Add("@BENACCO", SqlDbType.VarChar, 50).Value = SwiftData.benAccount
            swftComm.Parameters.Add("@BENEF", SqlDbType.VarChar, 150).Value = SwiftData.benName
            If Trim("" & SwiftData.incrCrAmountCant) = "" Then
                tmomey = 0
            Else : tmomey = CDec(SwiftData.incrCrAmountCant.Replace(",", sim))
            End If
            swftComm.Parameters.Add("@CREDMONTO", SqlDbType.Money).Value = tmomey
            swftComm.Parameters.Add("@CREDMONEDA", SqlDbType.Char, 3).Value = SwiftData.incrCrAmountMon
            swftComm.Parameters.Add("@PERCTOT1", SqlDbType.Int).Value = Val(SwiftData.percTolMonto1)
            swftComm.Parameters.Add("@PERCTOT2", SqlDbType.Int).Value = Val(SwiftData.percTolMonto2)
            swftComm.Parameters.Add("@MAXCREDMOD", SqlDbType.VarChar, 13).Value = SwiftData.maxCredAmmoMsg
            swftComm.Parameters.Add("@ADICMONTMEMO", SqlDbType.VarChar, 400).Value = SwiftData.addicAmmoCover
            swftComm.Parameters.Add("@AVAILWITH", SqlDbType.VarChar, 200).Value = SwiftData.availableWith
            swftComm.Parameters.Add("@DRAFTAT", SqlDbType.VarChar, 120).Value = SwiftData.draftsAt
            swftComm.Parameters.Add("@DRAWEEBANK", SqlDbType.VarChar, 160).Value = SwiftData.draweeBank
            swftComm.Parameters.Add("@MIXPAYDETL", SqlDbType.VarChar, 160).Value = SwiftData.mixPayDetail
            swftComm.Parameters.Add("@DEFERPYDET", SqlDbType.VarChar, 160).Value = SwiftData.deferrPayDetail
            swftComm.Parameters.Add("@PARTSHIPP", SqlDbType.VarChar, 35).Value = SwiftData.partialShipp
            swftComm.Parameters.Add("@TRANSSHIPP", SqlDbType.VarChar, 35).Value = SwiftData.transShipp
            swftComm.Parameters.Add("@DESPCHOPT", SqlDbType.VarChar, 65).Value = SwiftData.newDispatchOpt
            swftComm.Parameters.Add("@FINLDEST", SqlDbType.VarChar, 65).Value = SwiftData.newFinalDest
            swftComm.Parameters.Add("@LASTFECEM", SqlDbType.DateTime).Value = toDate(SwiftData.lateDateofShipp)
            swftComm.Parameters.Add("@SHIPPER", SqlDbType.VarChar, 400).Value = SwiftData.shippPeriod
            swftComm.Parameters.Add("@DESCRGOOD", SqlDbType.Text).Value = SwiftData.descripGoods
            swftComm.Parameters.Add("@DOCSREQUI", SqlDbType.Text).Value = SwiftData.docsRequired
            swftComm.Parameters.Add("@ADICCONDS", SqlDbType.Text).Value = SwiftData.additCondics
            swftComm.Parameters.Add("@CARGDED", SqlDbType.VarChar, 250).Value = SwiftData.cargDeducted
            swftComm.Parameters.Add("@PERIPRESEN", SqlDbType.VarChar, 150).Value = SwiftData.periodPresent
            swftComm.Parameters.Add("@APPLRULES", SqlDbType.VarChar, 100).Value = SwiftData.applicRules
            swftComm.Parameters.Add("@CONFRINSTR", SqlDbType.VarChar, 7).Value = SwiftData.confirmInstruct
            swftComm.Parameters.Add("@REIMBANK", SqlDbType.VarChar, 200).Value = SwiftData.correspSender
            swftComm.Parameters.Add("@INSTR2BANK", SqlDbType.VarChar, 800).Value = SwiftData.msgtoBank
            swftComm.Parameters.Add("@ADVISBANK", SqlDbType.VarChar, 200).Value = SwiftData.financialInstit
            swftComm.Parameters.Add("@SNDRECINF", SqlDbType.VarChar, 400).Value = SwiftData.send2recvInfo
            swftComm.Parameters.Add("@SWIFTMSG", SqlDbType.Text).Value = SwiftData.dat_All
            swftComm.Parameters.Add("@DATTAIL", SqlDbType.Text).Value = SwiftData.dat_Tail
            swftComm.Parameters.Add("@FILEDAT", SqlDbType.VarChar, 150).Value = SwiftData.OrigFile
            swftComm.Parameters.Add("@SWFSNDR", SqlDbType.VarChar, 12).Value = SwiftData.MsgSender
            swftComm.Parameters.Add("@SWFRCVR", SqlDbType.VarChar, 12).Value = SwiftData.MsgReceiver
            swftComm.Parameters.Add("@SWFMUR", SqlDbType.VarChar, 150).Value = SwiftData.MsgMUR
            swftComm.Parameters.Add("@SWFHED", SqlDbType.VarChar, 500).Value = SwiftData.MsgHead
            swftComm.Parameters.Add("@SWFTXT", SqlDbType.Text).Value = SwiftData.MsgText
            swftComm.Parameters.Add("@SWFTRAI", SqlDbType.VarChar, 200).Value = SwiftData.MsgTrailer
            swftComm.Parameters.Add("@PTORIG", SqlDbType.VarChar, 100).Value = SwiftData.shippOrigin
            swftComm.Parameters.Add("@PTDEST", SqlDbType.VarChar, 100).Value = SwiftData.shippDestiny
            swftComm.Parameters.Add("@SWFHEG", SqlDbType.VarChar, 500).Value = SwiftData.MsgHEng
            swftComm.Parameters.Add("@SWFTEG", SqlDbType.Text).Value = SwiftData.MsgTEng
            Try
                If conx.State = ConnectionState.Closed Then conx.Open()
                swftComm.ExecuteNonQuery()
                _c.Informa("  SWIFT :  " & SwiftData.tipo_mensaje & " -> " & SwiftData.msgReference & " Ingresado a BBDD")
                SaveDataSwift = True
            Catch ex As Exception
                _c.Informa("  Error: " & ex.ToString)
                SaveDataSwift = False
            Finally
                'swftComm.Dispose()
            End Try
        ElseIf SwiftData.tipo_mensaje = "707" Then
            Dim conx As New SqlConnection(_cnxstr)
            Dim ermess
            Dim strSwf As String = "INSERT INTO T_ECX_SwiftMT707_Sale"
            strSwf = strSwf & " (F_ISN, F_FechaIngreso, F_tipo_mensaje, F_Hora_ingreso, F_Corresponsal, " & _
                        " F_timeType, F_timeTime, F_timeOffs, " & _
                        " F_msgReference, F_receiverRef, F_issBankRef, F_issBankBicName, " & _
                        " F_fechIssue, F_fechModif, F_numeModif, F_benefPreModif, " & _
                        " F_fechNewExpir, F_IncrMontoCant, F_IncrMontoMone, F_DecrMontoCant, " & _
                        " F_DecrMontoMone, F_NewMontoCant, F_NewMontoMone, F_PorceTolerMonto, " & _
                        " F_MaxCredCode, F_AdicMontoMemo, F_OnBoardMemo, F_ForTranspMemo, " & _
                        " F_LastFecEmba, F_ShipPeriodMemo, F_Mensaje79, F_send2recvInfo, " & _
                        " F_SwiftMessage, F_DataTail, F_OrigFile, " & _
                        " F_SwiftSender, F_SwiftReceiver, F_SwiftMUR, F_SwiftHead, " & _
                        " F_SwiftText, F_SwiftTrailer, " & _
                        " F_SwiftHeng, F_SwiftTeng)"
            strSwf = strSwf & " VALUES (@ISN, @FECING, @TPMSG, @HRING, @CORRESP, " & _
                        " @TMTIP, @TMTIM, @TMOFF, " & _
                        " @MSGREF, @RELREF, @IBNKREF, @IBNKBIC, " & _
                        " @FECISS, @FECMOD, @NUMMOD, @BENPRE, " & _
                        " @FECNEWEXP, @INCRMONTO, @INCRMONEDA, @DECRMONTO, " & _
                        " @DECRMONEDA, @NEWMONTO, @NEWMONE, @PRCTOLM, " & _
                        " @MAXCREDMOD, @ADICMONTMEMO, @ONBRDMEM, @FORTRNSMEM, " & _
                        " @LASTFECEM, @SHIPPER, @MSJ79, @SNDRECINF, " & _
                        " @SWIFTMSG, @DATTAIL, @FILEDAT, " & _
                        " @SWFSNDR, @SWFRCVR, @SWFMUR, @SWFHED, @SWFTXT, @SWFTRAI, @SWFHEG, @SWFTEG)"

            'strSwf = strSwf & " (F_ISN, F_FechaIngreso, F_tipo_mensaje, F_Hora_ingreso, F_Corresponsal, " & _
            '           " F_timeType, F_timeTime, F_timeOffs, " & _
            '           " F_msgReference, F_receiverRef, F_issBankRef, F_issBankBicName, " & _
            '           " F_fechIssue, F_fechModif, F_numeModif, F_benefPreModif, " & _
            '           " F_IncrMontoCant, F_IncrMontoMone, F_DecrMontoCant, " & _
            '           " F_DecrMontoMone, F_NewMontoCant, F_NewMontoMone, F_PorceTolerMonto, " & _
            '           " F_MaxCredCode, F_AdicMontoMemo, F_OnBoardMemo, F_ForTranspMemo, " & _
            '           " F_ShipPeriodMemo, F_Mensaje79, F_send2recvInfo, " & _
            '           " F_SwiftMessage, F_DataTail, F_OrigFile, " & _
            '           " F_SwiftSender, F_SwiftReceiver, F_SwiftMUR, F_SwiftHead, " & _
            '           " F_SwiftText, F_SwiftTrailer)"
            'strSwf = strSwf & " VALUES (@ISN, @FECING, @TPMSG, @HRING, @CORRESP, " & _
            '            " @TMTIP, @TMTIM, @TMOFF, " & _
            '            " @MSGREF, @RELREF, @IBNKREF, @IBNKBIC, " & _
            '            " @FECISS, @FECMOD, @NUMMOD, @BENPRE, " & _
            '            " @INCRMONTO, @INCRMONEDA, @DECRMONTO, " & _
            '            " @DECRMONEDA, @NEWMONTO, @NEWMONE, @PRCTOLM, " & _
            '            " @MAXCREDMOD, @ADICMONTMEMO, @ONBRDMEM, @FORTRNSMEM, " & _
            '            " @SHIPPER, @MSJ79, @SNDRECINF, " & _
            '            " @SWIFTMSG, @DATTAIL, @FILEDAT, " & _
            '            " @SWFSNDR, @SWFRCVR, @SWFMUR, @SWFHED, @SWFTXT, @SWFTRAI)"
            Dim swftComm As New SqlCommand(strSwf, conx)
            swftComm.Parameters.Add("@ISN", SqlDbType.Int).Value = Val(SwiftData.ISN_mensaje)
            swftComm.Parameters.Add("@FECING", SqlDbType.DateTime).Value = SwiftData.Fecha_Ingreso
            swftComm.Parameters.Add("@TPMSG", SqlDbType.Int).Value = Val(SwiftData.tipo_mensaje)
            swftComm.Parameters.Add("@HRING", SqlDbType.VarChar).Value = Left(SwiftData.Hora_ingreso, 2) & ":" & Mid(SwiftData.Hora_ingreso, 3)
            swftComm.Parameters.Add("@CORRESP", SqlDbType.VarChar).Value = SwiftData.Corresponsal
            swftComm.Parameters.Add("@TMTIP", SqlDbType.VarChar).Value = SwiftData.timeType
            swftComm.Parameters.Add("@TMTIM", SqlDbType.VarChar).Value = Left(SwiftData.timeTime, 2) & ":" & Mid(SwiftData.timeTime, 3)
            swftComm.Parameters.Add("@TMOFF", SqlDbType.VarChar).Value = Left(SwiftData.timeOffs, 2) & ":" & Mid(SwiftData.timeOffs, 3)
            swftComm.Parameters.Add("@MSGREF", SqlDbType.VarChar).Value = SwiftData.msgReference
            swftComm.Parameters.Add("@RELREF", SqlDbType.VarChar).Value = SwiftData.relReference
            swftComm.Parameters.Add("@IBNKREF", SqlDbType.VarChar).Value = SwiftData.issuingBankReference
            swftComm.Parameters.Add("@IBNKBIC", SqlDbType.VarChar).Value = SwiftData.origBank
            swftComm.Parameters.Add("@FECISS", SqlDbType.DateTime).Value = toDate(SwiftData.dateofIssue)
            swftComm.Parameters.Add("@FECMOD", SqlDbType.DateTime).Value = toDate(SwiftData.dateofAmmend)
            swftComm.Parameters.Add("@NUMMOD", SqlDbType.Int).Value = Val(SwiftData.numofAmmend)
            swftComm.Parameters.Add("@BENPRE", SqlDbType.VarChar).Value = SwiftData.benName
            If Len(SwiftData.newdateofExpiry) > 0 Then
                swftComm.Parameters.Add("@FECNEWEXP", SqlDbType.DateTime).Value = toDate(SwiftData.newdateofExpiry)
            Else
                swftComm.Parameters.Add("@FECNEWEXP", SqlDbType.DateTime).Value = System.DBNull.Value
            End If
            If Trim("" & SwiftData.incrCrAmountCant) = "" Then
                tmomey = 0
            Else : tmomey = CDec(SwiftData.incrCrAmountCant.Replace(",", sim))
            End If
            swftComm.Parameters.Add("@INCRMONTO", SqlDbType.Money).Value = tmomey
            swftComm.Parameters.Add("@INCRMONEDA", SqlDbType.VarChar, 3).Value = SwiftData.incrCrAmountMon
            If Trim("" & SwiftData.instructAmount) = "" Then
                tmomey = 0
            Else : tmomey = CDec(SwiftData.instructAmount.Replace(",", sim))
            End If
            swftComm.Parameters.Add("@DECRMONTO", SqlDbType.Money).Value = tmomey
            swftComm.Parameters.Add("@DECRMONEDA", SqlDbType.VarChar, 3).Value = SwiftData.instructCurr
            If Trim("" & SwiftData.newCrAftAmmendAmount) = "" Then
                tmomey = 0
            Else : tmomey = CDec(SwiftData.newCrAftAmmendAmount.Replace(",", sim))
            End If
            swftComm.Parameters.Add("@NEWMONTO", SqlDbType.Money).Value = tmomey
            swftComm.Parameters.Add("@NEWMONE", SqlDbType.VarChar, 3).Value = SwiftData.newCrAftAmmendCurr
            swftComm.Parameters.Add("@PRCTOLM", SqlDbType.VarChar).Value = SwiftData.dat_39A
            swftComm.Parameters.Add("@MAXCREDMOD", SqlDbType.VarChar).Value = SwiftData.maxCredAmmoMsg
            swftComm.Parameters.Add("@ADICMONTMEMO", SqlDbType.VarChar).Value = SwiftData.addicAmmoCover
            swftComm.Parameters.Add("@ONBRDMEM", SqlDbType.VarChar).Value = SwiftData.newDispatchOpt
            swftComm.Parameters.Add("@FORTRNSMEM", SqlDbType.VarChar).Value = SwiftData.newFinalDest
            If Len(SwiftData.lateDateofShipp) > 0 Then
                swftComm.Parameters.Add("@LASTFECEM", SqlDbType.DateTime).Value = toDate(SwiftData.lateDateofShipp)
            Else
                swftComm.Parameters.Add("@LASTFECEM", SqlDbType.DateTime).Value = System.DBNull.Value
            End If
            swftComm.Parameters.Add("@SHIPPER", SqlDbType.VarChar).Value = SwiftData.shippPeriod
            swftComm.Parameters.Add("@MSJ79", SqlDbType.VarChar).Value = SwiftData.msgExplanation
            swftComm.Parameters.Add("@SNDRECINF", SqlDbType.VarChar).Value = SwiftData.send2recvInfo
            swftComm.Parameters.Add("@SWIFTMSG", SqlDbType.Text).Value = SwiftData.dat_All
            swftComm.Parameters.Add("@DATTAIL", SqlDbType.Text).Value = SwiftData.dat_Tail
            swftComm.Parameters.Add("@FILEDAT", SqlDbType.Text).Value = SwiftData.OrigFile
            swftComm.Parameters.Add("@SWFSNDR", SqlDbType.VarChar).Value = SwiftData.MsgSender
            swftComm.Parameters.Add("@SWFRCVR", SqlDbType.VarChar).Value = SwiftData.MsgReceiver
            swftComm.Parameters.Add("@SWFMUR", SqlDbType.VarChar).Value = SwiftData.MsgMUR
            swftComm.Parameters.Add("@SWFHED", SqlDbType.VarChar).Value = SwiftData.MsgHead
            swftComm.Parameters.Add("@SWFTXT", SqlDbType.Text).Value = SwiftData.MsgText
            swftComm.Parameters.Add("@SWFTRAI", SqlDbType.VarChar).Value = SwiftData.MsgTrailer
            swftComm.Parameters.Add("@SWFHEG", SqlDbType.VarChar).Value = SwiftData.MsgHEng
            swftComm.Parameters.Add("@SWFTEG", SqlDbType.Text).Value = SwiftData.MsgTEng

            Try
                If conx.State = ConnectionState.Closed Then conx.Open()
                swftComm.ExecuteNonQuery()
                _c.Informa("  SWIFT :  " & SwiftData.tipo_mensaje & " -> " & SwiftData.msgReference & " Ingresado a BBDD")
                SaveDataSwift = True
            Catch ex As Exception
                _c.Informa("  Error: " & ex.ToString)
                SaveDataSwift = False
            Finally
                'swftComm.Dispose()
            End Try
        ElseIf SwiftData.tipo_mensaje = "752" Then
            Dim conx As New SqlConnection(_cnxstr)
            Dim ermess
            Dim strSwf As String = "INSERT INTO T_ECX_SwiftMT752_Sale"
            strSwf = strSwf & " (F_ISN, F_FechaIngreso, F_tipo_mensaje, F_Hora_ingreso, F_Corresponsal, " & _
                        " F_timeType, F_timeTime, F_timeOffs, " & _
                        " F_docCredNum, F_presBnkRef, F_furtherIdentif, F_fechAdvise, " & _
                        " F_TotAdvMontoCant, F_TotAdvMontoMone, F_CargDeducMemo, F_fecNetAmmo, " & _
                        " F_NetAmmoCant, F_NetAmmoMone, F_correspSender, F_correspReceiver, " & _
                        " F_send2recvInfo, F_SwiftMessage, F_DataTail, F_OrigFile, F_SwiftSender, " & _
                        " F_SwiftReceiver, F_SwiftMUR, F_SwiftHead, F_SwiftText, F_SwiftTrailer, " & _
                        " F_SwiftHeng, F_SwiftTeng)"
            strSwf = strSwf & " VALUES (@ISN, @FECING, @TPMSG, @HRING, @CORRESP, " & _
                        " @TMTIP, @TMTIM, @TMOFF, " & _
                        " @DOCCREDNUM, @PRSBNKREF, @FURTHID, @FECADV, " & _
                        " @NEWMONTO, @NEWMONE, @CARGDEDU, @FECNETAMO, " & _
                        " @NETMONTO, @NETMONE, @CORRSENDR, @CORRECVR, " & _
                        " @SNDRECINF, @SWIFTMSG, @DATTAIL, @FILEDAT, " & _
                        " @SWFSNDR, @SWFRCVR, @SWFMUR, @SWFHED, @SWFTXT, @SWFTRAI, @SWFHEG, @SWFTEG)"

            Dim swftComm As New SqlCommand(strSwf, conx)
            swftComm.Parameters.Add("@ISN", SqlDbType.Int).Value = CInt(SwiftData.ISN_mensaje)
            swftComm.Parameters.Add("@FECING", SqlDbType.DateTime).Value = SwiftData.Fecha_Ingreso
            swftComm.Parameters.Add("@TPMSG", SqlDbType.Int).Value = CInt(SwiftData.tipo_mensaje)
            swftComm.Parameters.Add("@HRING", SqlDbType.VarChar).Value = Left(SwiftData.Hora_ingreso, 2) & ":" & Mid(SwiftData.Hora_ingreso, 3)
            swftComm.Parameters.Add("@CORRESP", SqlDbType.VarChar).Value = SwiftData.Corresponsal
            swftComm.Parameters.Add("@TMTIP", SqlDbType.VarChar).Value = SwiftData.timeType
            swftComm.Parameters.Add("@TMTIM", SqlDbType.VarChar).Value = Left(SwiftData.timeTime, 2) & ":" & Mid(SwiftData.timeTime, 3)
            swftComm.Parameters.Add("@TMOFF", SqlDbType.VarChar).Value = Left(SwiftData.timeOffs, 2) & ":" & Mid(SwiftData.timeOffs, 3)
            swftComm.Parameters.Add("@DOCCREDNUM", SqlDbType.VarChar).Value = SwiftData.msgReference
            swftComm.Parameters.Add("@PRSBNKREF", SqlDbType.VarChar).Value = SwiftData.relReference
            swftComm.Parameters.Add("@FURTHID", SqlDbType.VarChar).Value = SwiftData.issuingBankReference
            swftComm.Parameters.Add("@FECADV", SqlDbType.DateTime).Value = toDate(SwiftData.dateofAmmend)
            If Trim("" & SwiftData.incrCrAmountCant) = "" Then
                tmomey = 0
            Else : tmomey = CDec(SwiftData.incrCrAmountCant.Replace(",", sim))
            End If
            swftComm.Parameters.Add("@NEWMONTO", SqlDbType.Money).Value = tmomey
            swftComm.Parameters.Add("@NEWMONE", SqlDbType.VarChar, 3).Value = SwiftData.incrCrAmountMon
            swftComm.Parameters.Add("@CARGDEDU", SqlDbType.VarChar).Value = SwiftData.cargDeducted
            swftComm.Parameters.Add("@FECNETAMO", SqlDbType.DateTime).Value = toDate(SwiftData.fecNetAmmo)
            If Trim("" & SwiftData.instructAmount) = "" Then
                tmomey = 0
            Else : tmomey = CDec(SwiftData.instructAmount.Replace(",", sim))
            End If
            swftComm.Parameters.Add("@NETMONTO", SqlDbType.Money).Value = tmomey
            swftComm.Parameters.Add("@NETMONE", SqlDbType.VarChar, 3).Value = SwiftData.instructCurr
            swftComm.Parameters.Add("@CORRSENDR", SqlDbType.VarChar).Value = SwiftData.correspSender
            swftComm.Parameters.Add("@CORRECVR", SqlDbType.VarChar).Value = SwiftData.correspReceiver
            swftComm.Parameters.Add("@SNDRECINF", SqlDbType.VarChar).Value = SwiftData.send2recvInfo
            swftComm.Parameters.Add("@SWIFTMSG", SqlDbType.Text).Value = SwiftData.dat_All
            swftComm.Parameters.Add("@DATTAIL", SqlDbType.Text).Value = SwiftData.dat_Tail
            swftComm.Parameters.Add("@FILEDAT", SqlDbType.Text).Value = SwiftData.OrigFile
            swftComm.Parameters.Add("@SWFSNDR", SqlDbType.VarChar).Value = SwiftData.MsgSender
            swftComm.Parameters.Add("@SWFRCVR", SqlDbType.VarChar).Value = SwiftData.MsgReceiver
            swftComm.Parameters.Add("@SWFMUR", SqlDbType.VarChar).Value = SwiftData.MsgMUR
            swftComm.Parameters.Add("@SWFHED", SqlDbType.VarChar).Value = SwiftData.MsgHead
            swftComm.Parameters.Add("@SWFTXT", SqlDbType.Text).Value = SwiftData.MsgText
            swftComm.Parameters.Add("@SWFTRAI", SqlDbType.VarChar).Value = SwiftData.MsgTrailer
            swftComm.Parameters.Add("@SWFHEG", SqlDbType.VarChar).Value = SwiftData.MsgHEng
            swftComm.Parameters.Add("@SWFTEG", SqlDbType.Text).Value = SwiftData.MsgTEng

            Try
                If conx.State = ConnectionState.Closed Then conx.Open()
                swftComm.ExecuteNonQuery()
                _c.Informa("  SWIFT :  " & SwiftData.tipo_mensaje & " -> " & SwiftData.msgReference & " Ingresado a BBDD")
                SaveDataSwift = True
            Catch ex As Exception
                _c.Informa("  Error: " & ex.ToString)
                SaveDataSwift = False
            Finally
                'swftComm.Dispose()
            End Try
        ElseIf SwiftData.tipo_mensaje = "910" Then
            Dim conx As New SqlConnection(_cnxstr)
            Dim ermess
            Dim strSwf As String = "INSERT INTO T_ECX_SwiftMT910_Sale"
            strSwf = strSwf & " (F_ISN, F_FechaIngreso, F_tipo_mensaje, F_Hora_ingreso, F_Corresponsal, " & _
                        " F_msgReference, F_relReference, F_Valuta, F_Moneda, " & _
                        " F_Monto, F_orderingCusAcc, F_orderingCusName, " & _
                        " F_origBank, F_accountIdent, F_intermediariaInstit, " & _
                        " F_send2recvInfo, F_regulReport, " & _
                        " F_SwiftMessage, F_DataTail, F_OrigFile, " & _
                        " F_SwiftSender, F_SwiftReceiver, F_SwiftMUR, F_SwiftHead, F_SwiftText, F_SwiftTrailer, " & _
                        " F_SwiftHeng, F_SwiftTeng)"
            strSwf = strSwf & " VALUES (@ISN, @FECING, @TPMSG, @HRING, @CORRESP, " & _
                        " @MSGREF, @RELREF, @VALUTA, @MONEDA, @MONTO, " & _
                        " @CTAORDCUS, @NOMORDCUS, " & _
                        " @ORDBNK, @ACOIDEN, @INTERINST, " & _
                        " @SNDRECINF, @REGUREP, @SWIFTMSG, @DATTAIL, @FILEDAT, " & _
                        " @SWFSNDR, @SWFRCVR, @SWFMUR, @SWFHED, @SWFTXT, @SWFTRAI, @SWFHEG, @SWFTEG)"

            Dim swftComm As New SqlCommand(strSwf, conx)
            swftComm.Parameters.Add("@ISN", SqlDbType.Int).Value = CInt(SwiftData.ISN_mensaje)
            swftComm.Parameters.Add("@FECING", SqlDbType.DateTime).Value = SwiftData.Fecha_Ingreso
            swftComm.Parameters.Add("@TPMSG", SqlDbType.Int).Value = CInt(SwiftData.tipo_mensaje)
            swftComm.Parameters.Add("@HRING", SqlDbType.VarChar).Value = Left(SwiftData.Hora_ingreso, 2) & ":" & Mid(SwiftData.Hora_ingreso, 3)
            swftComm.Parameters.Add("@CORRESP", SqlDbType.VarChar).Value = SwiftData.Corresponsal
            swftComm.Parameters.Add("@MSGREF", SqlDbType.VarChar).Value = SwiftData.msgReference
            swftComm.Parameters.Add("@RELREF", SqlDbType.VarChar).Value = SwiftData.relReference
            swftComm.Parameters.Add("@VALUTA", SqlDbType.DateTime).Value = toDate(SwiftData.Valuta)
            swftComm.Parameters.Add("@MONEDA", SqlDbType.VarChar, 3).Value = SwiftData.Moneda
            If Trim("" & SwiftData.Monto) = "" Then
                tmomey = 0
            Else : tmomey = CDec(SwiftData.Monto.Replace(",", sim))
            End If
            swftComm.Parameters.Add("@MONTO", SqlDbType.Money).Value = tmomey
            swftComm.Parameters.Add("@CTAORDCUS", SqlDbType.VarChar).Value = SwiftData.orderingCusAcc
            swftComm.Parameters.Add("@NOMORDCUS", SqlDbType.VarChar).Value = SwiftData.orderingCusName
            swftComm.Parameters.Add("@ORDBNK", SqlDbType.VarChar).Value = SwiftData.origBank
            swftComm.Parameters.Add("@ACOIDEN", SqlDbType.VarChar).Value = SwiftData.accountIdent
            swftComm.Parameters.Add("@INTERINST", SqlDbType.VarChar).Value = SwiftData.intermediariaInstit
            swftComm.Parameters.Add("@SNDRECINF", SqlDbType.VarChar).Value = SwiftData.send2recvInfo
            swftComm.Parameters.Add("@REGUREP", SqlDbType.VarChar).Value = SwiftData.regulReport
            swftComm.Parameters.Add("@SWIFTMSG", SqlDbType.Text).Value = SwiftData.dat_All
            swftComm.Parameters.Add("@DATTAIL", SqlDbType.Text).Value = SwiftData.dat_Tail
            swftComm.Parameters.Add("@FILEDAT", SqlDbType.Text).Value = SwiftData.OrigFile
            swftComm.Parameters.Add("@SWFSNDR", SqlDbType.VarChar).Value = SwiftData.MsgSender
            swftComm.Parameters.Add("@SWFRCVR", SqlDbType.VarChar).Value = SwiftData.MsgReceiver
            swftComm.Parameters.Add("@SWFMUR", SqlDbType.VarChar).Value = SwiftData.MsgMUR
            swftComm.Parameters.Add("@SWFHED", SqlDbType.VarChar).Value = SwiftData.MsgHead
            swftComm.Parameters.Add("@SWFTXT", SqlDbType.Text).Value = SwiftData.MsgText
            swftComm.Parameters.Add("@SWFTRAI", SqlDbType.VarChar).Value = SwiftData.MsgTrailer
            swftComm.Parameters.Add("@SWFHEG", SqlDbType.VarChar).Value = SwiftData.MsgHEng
            swftComm.Parameters.Add("@SWFTEG", SqlDbType.Text).Value = SwiftData.MsgTEng

            Try
                If conx.State = ConnectionState.Closed Then conx.Open()
                '  swftComm.ExecuteNonQuery()
                _c.Informa("  SWIFT :  " & SwiftData.tipo_mensaje & " -> " & SwiftData.msgReference & " Ingresado a BBDD")
                SaveDataSwift = True
            Catch ex As Exception
                _c.Informa("  Error: " & ex.ToString)
                SaveDataSwift = False
            Finally
                'swftComm.Dispose()
            End Try
        ElseIf SwiftData.tipo_mensaje = "799" Then
            Dim conx As New SqlConnection(_cnxstr)
            Dim ermess
            Dim strSwf As String = "INSERT INTO T_ECX_SwiftMT799_Sale"
            strSwf = strSwf & " (F_ISN, F_FechaIngreso, F_tipo_mensaje, F_Hora_ingreso, F_Corresponsal, " & _
                        " F_msgReference, F_relReference, F_Explanation, " & _
                        " F_SwiftMessage, F_DataTail, F_OrigFile, " & _
                        " F_SwiftSender, F_SwiftReceiver, F_SwiftMUR, F_SwiftHead, F_SwiftText, F_SwiftTrailer, " & _
                        " F_SwiftHeng, F_SwiftTeng)"
            strSwf = strSwf & " VALUES (@ISN, @FECING, @TPMSG, @HRING, @CORRESP, " & _
                        " @MSGREF, @RELREF, @EXPLAN, @SWIFTMSG, @DATTAIL, @FILEDAT, " & _
                        " @SWFSNDR, @SWFRCVR, @SWFMUR, @SWFHED, @SWFTXT, @SWFTRAI, @SWFHEG, @SWFTEG)"

            Dim swftComm As New SqlCommand(strSwf, conx)
            swftComm.Parameters.Add("@ISN", SqlDbType.Int).Value = CInt(SwiftData.ISN_mensaje)
            swftComm.Parameters.Add("@FECING", SqlDbType.DateTime).Value = SwiftData.Fecha_Ingreso
            swftComm.Parameters.Add("@TPMSG", SqlDbType.Int).Value = CInt(SwiftData.tipo_mensaje)
            swftComm.Parameters.Add("@HRING", SqlDbType.VarChar).Value = Left(SwiftData.Hora_ingreso, 2) & ":" & Mid(SwiftData.Hora_ingreso, 3)
            swftComm.Parameters.Add("@CORRESP", SqlDbType.VarChar).Value = SwiftData.Corresponsal
            swftComm.Parameters.Add("@MSGREF", SqlDbType.VarChar).Value = SwiftData.msgReference
            swftComm.Parameters.Add("@RELREF", SqlDbType.VarChar).Value = SwiftData.relReference
            swftComm.Parameters.Add("@EXPLAN", SqlDbType.VarChar, 1800).Value = SwiftData.msgExplanation
            swftComm.Parameters.Add("@SWIFTMSG", SqlDbType.Text).Value = SwiftData.dat_All
            swftComm.Parameters.Add("@DATTAIL", SqlDbType.Text).Value = SwiftData.dat_Tail
            swftComm.Parameters.Add("@FILEDAT", SqlDbType.Text).Value = SwiftData.OrigFile
            swftComm.Parameters.Add("@SWFSNDR", SqlDbType.VarChar).Value = SwiftData.MsgSender
            swftComm.Parameters.Add("@SWFRCVR", SqlDbType.VarChar).Value = SwiftData.MsgReceiver
            swftComm.Parameters.Add("@SWFMUR", SqlDbType.VarChar).Value = SwiftData.MsgMUR
            swftComm.Parameters.Add("@SWFHED", SqlDbType.VarChar).Value = SwiftData.MsgHead
            swftComm.Parameters.Add("@SWFTXT", SqlDbType.Text).Value = SwiftData.MsgText
            swftComm.Parameters.Add("@SWFTRAI", SqlDbType.VarChar).Value = SwiftData.MsgTrailer
            swftComm.Parameters.Add("@SWFHEG", SqlDbType.VarChar).Value = SwiftData.MsgHEng
            swftComm.Parameters.Add("@SWFTEG", SqlDbType.Text).Value = SwiftData.MsgTEng

            Try
                If conx.State = ConnectionState.Closed Then conx.Open()
                swftComm.ExecuteNonQuery()
                _c.Informa("  SWIFT :  " & SwiftData.tipo_mensaje & " -> " & SwiftData.msgReference & " Ingresado a BBDD")
                SaveDataSwift = True
            Catch ex As Exception
                _c.Informa("  Error: " & ex.ToString)
                SaveDataSwift = False
            Finally
                'swftComm.Dispose()
            End Try
            'msg_799 = "Señores," & vbCrLf
            'msg_799 = msg_799 & "   Se ha recibido mensaje swift 799, que adjuntamos para su revisión." & vbCrLf
            'msg_799 = msg_799 & vbCrLf
            'msg_799 = msg_799 & SwiftData.MsgHead & vbCrLf & "-------------------------------------------" & vbCrLf & SwiftData.MsgText
            ''msg_799 = msg_799 & SwiftData.dat_All
            'If Not Envia("Mensaje Swift 799 Recibido", msg_799) Then
            '    _c.Informa(" Mensaje MT799 enviado a los respectivos Mails ")
            'Else
            '    _c.Informa(" No pudo ser enviado el Mensaje MT799 a los respectivos Mails ")
            'End If


        Else
            _c.Informa("   -- Tipo Swift: " & SwiftData.tipo_mensaje & " -> " & SwiftData.Moneda)
            SaveDataSwift = True
        End If
        _c.Informa("----------------------------------------------------------------------")
    End Function
    Function SaveDataSwift_back(ByVal SwiftData As SwiftsJob.SwiftData, ByVal errMessage As String) As Boolean
        Dim tmomey As Decimal
        Dim sim As String = Mid("" & 1.1, 2, 1)
        Dim msg_799 As String, OkDone As Boolean
        _c.Informa("-----------------------------------")
        _c.Informa(":Mensaje Swift Tipo MT" & SwiftData.tipo_mensaje) '& "----------------------------")
        _c.Informa("     -- ISN_mensaje > " & SwiftData.ISN_mensaje)
        _c.Informa("     -- tipo_mensaje > " & SwiftData.tipo_mensaje)
        _c.Informa("     -- Data Tail > " & SwiftData.dat_Tail)
        'SaveDataSwift = True
        'Exit Function
        'If SwiftData.tipo_mensaje <> "700" Then
        '    SwiftData.tipo_mensaje = "NOPROC_" & SwiftData.tipo_mensaje
        'End If
        If SwiftData.tipo_mensaje = "103" Then
            Dim conx As New SqlConnection(_cnxstr)
            Dim ermess
            Dim strSwf As String = "INSERT INTO T_ECX_SwiftMT103_Sale"
            strSwf = strSwf & " (F_ISN, F_FechaIngreso, F_tipo_mensaje, F_Hora_ingreso, F_Corresponsal, " & _
                        " F_timeType, F_timeTime, F_timeOffs, " & _
                        " F_msgReference, F_bankOperatCode, F_instructCodeCode, " & _
                        " F_instructCodeRefr, F_transacType, F_Valuta, F_Moneda, F_Monto, " & _
                        " F_instructCurr, F_instructAmount, F_exchRate, F_orderingCusAcc, " & _
                        " F_orderingCusName, F_origBank, F_correspSender, F_correspReceiver, " & _
                        " F_terceraInstit, F_intermediariaInstit, F_financialInstit, " & _
                        " F_benAccount, F_benName, F_remitInfo, F_cargoTo, " & _
                        " F_senderChargCurr, F_senderChargAmou, F_recverChargCurr, F_recverChargAmou, " & _
                        " F_send2recvInfo, F_regulReport, F_SwiftMessage, F_DataTail, F_OrigFile, " & _
                        " F_SwiftSender, F_SwiftReceiver, F_SwiftMUR, F_SwiftHead, F_SwiftText, F_SwiftTrailer)"
            'VALUES     (3, '3', '3', '3', '3', '3', '3', '3', '3', '3', '3', 3, '3', 3, '3', 3, 3, '3', '3', '3', '3', '3', '3', '3', '3', '3', '3', '3', '3', '3', '3', 3, '3', 3, '3', '3', '3')
            strSwf = strSwf & " VALUES (@ISN, @FECING, @TPMSG, @HRING, @CORRESP, " & _
                        " @TMTIP, @TMTIM, @TMOFF, " & _
                        " @MSGREF, @BKOPCOD, @ITRCDCD, " & _
                        " @ITRCDRF, @TRNTIP, @VALUTA, @MONEDA, @MONTO, " & _
                        " @ITRCURR, @ITRMTO, @TPOCAM, @CTAORDCUS, " & _
                        " @NOMORDCUS, @ORDBNK, @CORRSENDR, @CORRECVR, " & _
                        " @TERCINST, @INTERINST, @FINANINST, " & _
                        " @BENACO, @BENNOM, @REMINFO, @CARGOTO, " & _
                        " @SNDCHRCUR, @SNDCHRAMO, @RCVCHRCUR, @RCVCHRAMO, " & _
                        " @SNDRECINF, @REGUREP, @SWIFTMSG, @DATTAIL, @FILEDAT, " & _
                        " @SWFSNDR, @SWFRCVR, @SWFMUR, @SWFHED, @SWFTXT, @SWFTRAI)"

            Dim swftComm As New SqlCommand(strSwf, conx)
            swftComm.Parameters.Add("@ISN", SqlDbType.Int).Value = CInt(SwiftData.ISN_mensaje)
            swftComm.Parameters.Add("@FECING", SqlDbType.DateTime).Value = SwiftData.Fecha_Ingreso
            swftComm.Parameters.Add("@TPMSG", SqlDbType.Int).Value = CInt(SwiftData.tipo_mensaje)
            swftComm.Parameters.Add("@HRING", SqlDbType.VarChar).Value = Left(SwiftData.Hora_ingreso, 2) & ":" & Mid(SwiftData.Hora_ingreso, 3)
            swftComm.Parameters.Add("@CORRESP", SqlDbType.VarChar).Value = SwiftData.Corresponsal
            swftComm.Parameters.Add("@TMTIP", SqlDbType.VarChar).Value = SwiftData.timeType
            swftComm.Parameters.Add("@TMTIM", SqlDbType.VarChar).Value = Left(SwiftData.timeTime, 2) & ":" & Mid(SwiftData.timeTime, 3)
            swftComm.Parameters.Add("@TMOFF", SqlDbType.VarChar).Value = Left(SwiftData.timeOffs, 2) & ":" & Mid(SwiftData.timeOffs, 3)
            swftComm.Parameters.Add("@MSGREF", SqlDbType.VarChar).Value = SwiftData.msgReference
            swftComm.Parameters.Add("@BKOPCOD", SqlDbType.VarChar, 4).Value = SwiftData.bankOperatCode
            swftComm.Parameters.Add("@ITRCDCD", SqlDbType.VarChar, 4).Value = SwiftData.instructCodeCode
            swftComm.Parameters.Add("@ITRCDRF", SqlDbType.VarChar, 30).Value = SwiftData.instructCodeRefr
            swftComm.Parameters.Add("@TRNTIP", SqlDbType.VarChar, 3).Value = SwiftData.transacType
            swftComm.Parameters.Add("@VALUTA", SqlDbType.DateTime).Value = toDate(SwiftData.Valuta)
            swftComm.Parameters.Add("@MONEDA", SqlDbType.VarChar, 3).Value = SwiftData.Moneda
            If Trim("" & SwiftData.Monto) = "" Then
                tmomey = 0
            Else : tmomey = CDec(SwiftData.Monto.Replace(",", sim))
            End If
            swftComm.Parameters.Add("@MONTO", SqlDbType.Money).Value = tmomey
            swftComm.Parameters.Add("@ITRCURR", SqlDbType.VarChar).Value = SwiftData.instructCurr
            If Trim("" & SwiftData.instructAmount) = "" Then
                tmomey = 0
            Else : tmomey = CDec(SwiftData.instructAmount.Replace(",", sim))
            End If
            swftComm.Parameters.Add("@ITRMTO", SqlDbType.Money).Value = tmomey
            If Trim("" & SwiftData.exchRate) = "" Then
                tmomey = 0
            Else : tmomey = CDec(SwiftData.exchRate.Replace(",", sim))
            End If
            swftComm.Parameters.Add("@TPOCAM", SqlDbType.Money).Value = tmomey
            swftComm.Parameters.Add("@CTAORDCUS", SqlDbType.VarChar).Value = SwiftData.orderingCusAcc
            swftComm.Parameters.Add("@NOMORDCUS", SqlDbType.VarChar).Value = SwiftData.orderingCusName
            swftComm.Parameters.Add("@ORDBNK", SqlDbType.VarChar).Value = SwiftData.origBank
            swftComm.Parameters.Add("@CORRSENDR", SqlDbType.VarChar).Value = SwiftData.correspSender
            swftComm.Parameters.Add("@CORRECVR", SqlDbType.VarChar).Value = SwiftData.correspReceiver
            swftComm.Parameters.Add("@TERCINST", SqlDbType.VarChar).Value = SwiftData.terceraInstit
            swftComm.Parameters.Add("@INTERINST", SqlDbType.VarChar).Value = SwiftData.intermediariaInstit
            swftComm.Parameters.Add("@FINANINST", SqlDbType.VarChar).Value = SwiftData.financialInstit
            swftComm.Parameters.Add("@BENACO", SqlDbType.VarChar).Value = SwiftData.benAccount
            swftComm.Parameters.Add("@BENNOM", SqlDbType.VarChar).Value = SwiftData.benName
            swftComm.Parameters.Add("@REMINFO", SqlDbType.VarChar).Value = SwiftData.remitInfo
            swftComm.Parameters.Add("@CARGOTO", SqlDbType.VarChar).Value = SwiftData.cargoTo
            swftComm.Parameters.Add("@SNDCHRCUR", SqlDbType.VarChar).Value = SwiftData.senderChargCurr
            'If Trim("" & SwiftData.senderChargAmou) = "" Then
            '    tmomey = 0
            'Else : tmomey = CDec(SwiftData.senderChargAmou.Replace(",", sim))
            'End If
            swftComm.Parameters.Add("@SNDCHRAMO", SqlDbType.VarChar).Value = SwiftData.senderChargAmou
            swftComm.Parameters.Add("@RCVCHRCUR", SqlDbType.VarChar).Value = SwiftData.recverChargCurr
            'If Trim("" & SwiftData.recverChargAmou) = "" Then
            '    tmomey = 0
            'Else : tmomey = CDec(SwiftData.recverChargAmou.Replace(",", sim))
            'End If
            swftComm.Parameters.Add("@RCVCHRAMO", SqlDbType.VarChar).Value = SwiftData.recverChargAmou
            swftComm.Parameters.Add("@SNDRECINF", SqlDbType.VarChar).Value = SwiftData.send2recvInfo
            swftComm.Parameters.Add("@REGUREP", SqlDbType.VarChar).Value = SwiftData.regulReport
            swftComm.Parameters.Add("@SWIFTMSG", SqlDbType.Text).Value = SwiftData.dat_All
            swftComm.Parameters.Add("@DATTAIL", SqlDbType.Text).Value = SwiftData.dat_Tail
            swftComm.Parameters.Add("@FILEDAT", SqlDbType.Text).Value = SwiftData.OrigFile
            swftComm.Parameters.Add("@SWFSNDR", SqlDbType.VarChar).Value = SwiftData.MsgSender
            swftComm.Parameters.Add("@SWFRCVR", SqlDbType.VarChar).Value = SwiftData.MsgReceiver
            swftComm.Parameters.Add("@SWFMUR", SqlDbType.VarChar).Value = SwiftData.MsgMUR
            swftComm.Parameters.Add("@SWFHED", SqlDbType.VarChar).Value = SwiftData.MsgHead
            swftComm.Parameters.Add("@SWFTXT", SqlDbType.Text).Value = SwiftData.MsgText
            swftComm.Parameters.Add("@SWFTRAI", SqlDbType.VarChar).Value = SwiftData.MsgTrailer

            Try
                If conx.State = ConnectionState.Closed Then conx.Open()
                swftComm.ExecuteNonQuery()
                _c.Informa("  SWIFT :  " & SwiftData.tipo_mensaje & " -> " & SwiftData.msgReference & " Ingresado a BBDD")
                SaveDataSwift_back = True
            Catch ex As Exception
                _c.Informa("  Error: " & ex.ToString)
                SaveDataSwift_back = False
            Finally
                'swftComm.Dispose()
            End Try
        ElseIf SwiftData.tipo_mensaje = "202" Then
            Dim conx As New SqlConnection(_cnxstr)
            Dim ermess
            Dim strSwf As String = "INSERT INTO T_ECX_SwiftMT202_Sale"
            strSwf = strSwf & " (F_ISN, F_FechaIngreso, F_tipo_mensaje, F_Hora_ingreso, F_Corresponsal, " & _
                        " F_timeType, F_timeTime, F_timeOffs,  " & _
                        " F_msgReference, F_relReference, F_Valuta, F_Moneda, " & _
                        " F_Monto, F_origBank, F_correspSender, F_correspReceiver, " & _
                        " F_intermediariaInstit, F_financialInstit, F_benefInstit, " & _
                        " F_send2recvInfo, F_regulReport, " & _
                        " F_SwiftMessage, F_DataTail, F_OrigFile, " & _
                        " F_SwiftSender, F_SwiftReceiver, F_SwiftMUR, F_SwiftHead, F_SwiftText, F_SwiftTrailer)"
            'VALUES     (3, '3', '3', '3', '3', '3', '3', '3', '3', '3', '3', 3, '3', 3, '3', 3, 3, '3', '3', '3', '3', '3', '3', '3', '3', '3', '3', '3', '3', '3', '3', 3, '3', 3, '3', '3', '3')
            strSwf = strSwf & " VALUES (@ISN, @FECING, @TPMSG, @HRING, @CORRESP, " & _
                        " @TMTIP, @TMTIM, @TMOFF, " & _
                        " @MSGREF, @RELREF, @VALUTA, @MONEDA, @MONTO, " & _
                        " @ORDBNK, @CORRSENDR, @CORRECVR, " & _
                        " @INTERINST, @FINANINST, @BENFINST, " & _
                        " @SNDRECINF, @REGUREP, @SWIFTMSG, @DATTAIL, @FILEDAT, " & _
                        " @SWFSNDR, @SWFRCVR, @SWFMUR, @SWFHED, @SWFTXT, @SWFTRAI)"

            Dim swftComm As New SqlCommand(strSwf, conx)
            swftComm.Parameters.Add("@ISN", SqlDbType.Int).Value = CInt(SwiftData.ISN_mensaje)
            swftComm.Parameters.Add("@FECING", SqlDbType.DateTime).Value = SwiftData.Fecha_Ingreso
            swftComm.Parameters.Add("@TPMSG", SqlDbType.Int).Value = CInt(SwiftData.tipo_mensaje)
            swftComm.Parameters.Add("@HRING", SqlDbType.VarChar).Value = Left(SwiftData.Hora_ingreso, 2) & ":" & Mid(SwiftData.Hora_ingreso, 3)
            swftComm.Parameters.Add("@CORRESP", SqlDbType.VarChar).Value = SwiftData.Corresponsal
            swftComm.Parameters.Add("@TMTIP", SqlDbType.VarChar).Value = SwiftData.timeType
            swftComm.Parameters.Add("@TMTIM", SqlDbType.VarChar).Value = Left(SwiftData.timeTime, 2) & ":" & Mid(SwiftData.timeTime, 3)
            swftComm.Parameters.Add("@TMOFF", SqlDbType.VarChar).Value = Left(SwiftData.timeOffs, 2) & ":" & Mid(SwiftData.timeOffs, 3)
            swftComm.Parameters.Add("@MSGREF", SqlDbType.VarChar).Value = SwiftData.msgReference
            swftComm.Parameters.Add("@RELREF", SqlDbType.VarChar).Value = SwiftData.relReference
            swftComm.Parameters.Add("@VALUTA", SqlDbType.DateTime).Value = toDate(SwiftData.Valuta)
            swftComm.Parameters.Add("@MONEDA", SqlDbType.VarChar, 3).Value = SwiftData.Moneda
            If Trim("" & SwiftData.Monto) = "" Then
                tmomey = 0
            Else : tmomey = CDec(SwiftData.Monto.Replace(",", sim))
            End If
            swftComm.Parameters.Add("@MONTO", SqlDbType.Money).Value = tmomey
            swftComm.Parameters.Add("@ORDBNK", SqlDbType.VarChar).Value = SwiftData.origBank
            swftComm.Parameters.Add("@CORRSENDR", SqlDbType.VarChar).Value = SwiftData.correspSender
            swftComm.Parameters.Add("@CORRECVR", SqlDbType.VarChar).Value = SwiftData.correspReceiver
            swftComm.Parameters.Add("@INTERINST", SqlDbType.VarChar).Value = SwiftData.intermediariaInstit
            swftComm.Parameters.Add("@FINANINST", SqlDbType.VarChar).Value = SwiftData.financialInstit
            swftComm.Parameters.Add("@BENFINST", SqlDbType.VarChar).Value = SwiftData.benefInstit
            swftComm.Parameters.Add("@SNDRECINF", SqlDbType.VarChar).Value = SwiftData.send2recvInfo
            swftComm.Parameters.Add("@REGUREP", SqlDbType.VarChar).Value = SwiftData.regulReport
            swftComm.Parameters.Add("@SWIFTMSG", SqlDbType.Text).Value = SwiftData.dat_All
            swftComm.Parameters.Add("@DATTAIL", SqlDbType.Text).Value = SwiftData.dat_Tail
            swftComm.Parameters.Add("@FILEDAT", SqlDbType.Text).Value = SwiftData.OrigFile
            swftComm.Parameters.Add("@SWFSNDR", SqlDbType.VarChar).Value = SwiftData.MsgSender
            swftComm.Parameters.Add("@SWFRCVR", SqlDbType.VarChar).Value = SwiftData.MsgReceiver
            swftComm.Parameters.Add("@SWFMUR", SqlDbType.VarChar).Value = SwiftData.MsgMUR
            swftComm.Parameters.Add("@SWFHED", SqlDbType.VarChar).Value = SwiftData.MsgHead
            swftComm.Parameters.Add("@SWFTXT", SqlDbType.Text).Value = SwiftData.MsgText
            swftComm.Parameters.Add("@SWFTRAI", SqlDbType.VarChar).Value = SwiftData.MsgTrailer

            Try
                If conx.State = ConnectionState.Closed Then conx.Open()
                swftComm.ExecuteNonQuery()
                _c.Informa("  SWIFT :  " & SwiftData.tipo_mensaje & " -> " & SwiftData.msgReference & " Ingresado a BBDD")
                SaveDataSwift_back = True
            Catch ex As Exception
                _c.Informa("  Error: " & ex.ToString)
                SaveDataSwift_back = False
            Finally
                'swftComm.Dispose()
            End Try
        ElseIf SwiftData.tipo_mensaje = "700" Then
            Dim conx As New SqlConnection(_cnxstr)
            Dim ermess
            Dim strSwf As String
            strSwf = "" 'IF NOT exists(SELECT 1 FROM T_ECX_SwiftMT700_Sale WHERE F_docCredNumber = @MSGREF)"
            strSwf = strSwf & " INSERT INTO T_ECX_SwiftMT700_Sale " & _
                        " (F_ISN, F_FechaIngreso, F_FechaProceso, F_tipo_mensaje, " & _
                        " F_Hora_ingreso, F_Corresponsal, " & _
                        " F_docCredNumber, F_ref2PreAdvice, F_seqIndex, F_seqTotal, " & _
                        " F_typeofCredit, F_fechIssue, F_dateofExpire, F_placeofExpire, " & _
                        " F_applicBank, F_applicant, F_benAccount, F_benName, " & _
                        " F_MontoCredCant, F_MontoCredMone, F_percTolMonto1, F_percTolMonto2, " & _
                        " F_MaxCredCode, F_AdicMontoMemo, F_availableWith, F_draftsAt, F_draweeBank, " & _
                        " F_mixPayDetail, F_deferrPayDetail, F_partialShipp, F_transShipp, F_DispatchOpt,  " & _
                        " F_FinalDest, F_lateDateofShipp, F_shippPeriod, F_descripGoods, " & _
                        " F_docsRequired, F_additCondics, F_cargDeducted, F_periodPresent,  " & _
                        " F_appl_rules, F_confirmInstruct, F_reimbursingBank, F_instruct2Bank, " & _
                        " F_adviserBank, F_send2recvInfo, F_SwiftMessage, F_DataTail, " & _
                        " F_OrigFile, F_SwiftSender, F_SwiftReceiver, F_SwiftMUR, " & _
                        " F_SwiftHead, F_SwiftText, F_SwiftTrailer)"
            strSwf = strSwf & " VALUES (@ISN, @FECING, GETDATE(), @TPMSG, @HRING, @CORRESP, " & _
                        " @MSGREF, @RELREF, @SEQINDX, @SEQTOTL, " & _
                        " @TYPECRED, @FECISS, @FECEXPI, @LUGEXPI, " & _
                        " @APPLICBANK, @APPLIC, @BENACCO, @BENEF, " & _
                        " @CREDMONTO, @CREDMONEDA, @PERCTOT1, @PERCTOT2, " & _
                        " @MAXCREDMOD, @ADICMONTMEMO, @AVAILWITH, @DRAFTAT, @DRAWEEBANK, " & _
                        " @MIXPAYDETL, @DEFERPYDET, @PARTSHIPP, @TRANSSHIPP, @DESPCHOPT, " & _
                        " @FINLDEST, @LASTFECEM, @SHIPPER, @DESCRGOOD, " & _
                        " @DOCSREQUI, @ADICCONDS, @CARGDED, @PERIPRESEN, " & _
                        " @APPLRULES, @CONFRINSTR, @REIMBANK, @INSTR2BANK, " & _
                        " @ADVISBANK, @SNDRECINF, " & _
                        " @SWIFTMSG, @DATTAIL, @FILEDAT, " & _
                        " @SWFSNDR, @SWFRCVR, @SWFMUR, @SWFHED, @SWFTXT, @SWFTRAI)"

            Dim swftComm As New SqlCommand(strSwf, conx)
            swftComm.Parameters.Add("@ISN", SqlDbType.Int).Value = CInt(SwiftData.ISN_mensaje)
            swftComm.Parameters.Add("@FECING", SqlDbType.DateTime).Value = SwiftData.Fecha_Ingreso
            swftComm.Parameters.Add("@TPMSG", SqlDbType.Int).Value = CInt(SwiftData.tipo_mensaje)
            swftComm.Parameters.Add("@HRING", SqlDbType.VarChar, 5).Value = Left(SwiftData.Hora_ingreso, 2) & ":" & Mid(SwiftData.Hora_ingreso, 3)
            swftComm.Parameters.Add("@CORRESP", SqlDbType.VarChar, 50).Value = SwiftData.Corresponsal
            swftComm.Parameters.Add("@MSGREF", SqlDbType.VarChar, 16).Value = SwiftData.msgReference
            swftComm.Parameters.Add("@RELREF", SqlDbType.VarChar, 16).Value = SwiftData.issuingBankReference
            swftComm.Parameters.Add("@SEQINDX", SqlDbType.Int).Value = CInt(SwiftData.seqIndex)
            swftComm.Parameters.Add("@SEQTOTL", SqlDbType.Int).Value = CInt(SwiftData.seqTotal)
            swftComm.Parameters.Add("@TYPECRED", SqlDbType.VarChar, 24).Value = SwiftData.typeofCredit
            swftComm.Parameters.Add("@FECISS", SqlDbType.DateTime).Value = toDate(SwiftData.dateofIssue)
            swftComm.Parameters.Add("@FECEXPI", SqlDbType.DateTime).Value = toDate(SwiftData.dateofExpire)
            swftComm.Parameters.Add("@LUGEXPI", SqlDbType.VarChar, 30).Value = SwiftData.placeofExpire
            swftComm.Parameters.Add("@APPLICBANK", SqlDbType.VarChar, 200).Value = SwiftData.applicBank
            swftComm.Parameters.Add("@APPLIC", SqlDbType.VarChar, 150).Value = SwiftData.applicant
            swftComm.Parameters.Add("@BENACCO", SqlDbType.VarChar, 50).Value = SwiftData.benAccount
            swftComm.Parameters.Add("@BENEF", SqlDbType.VarChar, 150).Value = SwiftData.benName
            If Trim("" & SwiftData.incrCrAmountCant) = "" Then
                tmomey = 0
            Else : tmomey = CDec(SwiftData.incrCrAmountCant.Replace(",", sim))
            End If
            swftComm.Parameters.Add("@CREDMONTO", SqlDbType.Money).Value = tmomey
            swftComm.Parameters.Add("@CREDMONEDA", SqlDbType.Char, 3).Value = SwiftData.incrCrAmountMon
            swftComm.Parameters.Add("@PERCTOT1", SqlDbType.Int).Value = Val(SwiftData.percTolMonto1)
            swftComm.Parameters.Add("@PERCTOT2", SqlDbType.Int).Value = Val(SwiftData.percTolMonto2)
            swftComm.Parameters.Add("@MAXCREDMOD", SqlDbType.VarChar, 13).Value = SwiftData.maxCredAmmoMsg
            swftComm.Parameters.Add("@ADICMONTMEMO", SqlDbType.VarChar, 400).Value = SwiftData.addicAmmoCover
            swftComm.Parameters.Add("@AVAILWITH", SqlDbType.VarChar, 200).Value = SwiftData.availableWith
            swftComm.Parameters.Add("@DRAFTAT", SqlDbType.VarChar, 120).Value = SwiftData.draftsAt
            swftComm.Parameters.Add("@DRAWEEBANK", SqlDbType.VarChar, 160).Value = SwiftData.draweeBank
            swftComm.Parameters.Add("@MIXPAYDETL", SqlDbType.VarChar, 160).Value = SwiftData.mixPayDetail
            swftComm.Parameters.Add("@DEFERPYDET", SqlDbType.VarChar, 160).Value = SwiftData.deferrPayDetail
            swftComm.Parameters.Add("@PARTSHIPP", SqlDbType.VarChar, 35).Value = SwiftData.partialShipp
            swftComm.Parameters.Add("@TRANSSHIPP", SqlDbType.VarChar, 35).Value = SwiftData.transShipp
            swftComm.Parameters.Add("@DESPCHOPT", SqlDbType.VarChar, 65).Value = SwiftData.newDispatchOpt
            swftComm.Parameters.Add("@FINLDEST", SqlDbType.VarChar, 65).Value = SwiftData.newFinalDest
            swftComm.Parameters.Add("@LASTFECEM", SqlDbType.DateTime).Value = toDate(SwiftData.lateDateofShipp)
            swftComm.Parameters.Add("@SHIPPER", SqlDbType.VarChar, 400).Value = SwiftData.shippPeriod
            swftComm.Parameters.Add("@DESCRGOOD", SqlDbType.Text).Value = SwiftData.descripGoods
            swftComm.Parameters.Add("@DOCSREQUI", SqlDbType.Text).Value = SwiftData.docsRequired
            swftComm.Parameters.Add("@ADICCONDS", SqlDbType.Text).Value = SwiftData.additCondics
            swftComm.Parameters.Add("@CARGDED", SqlDbType.VarChar, 250).Value = SwiftData.cargDeducted
            swftComm.Parameters.Add("@PERIPRESEN", SqlDbType.VarChar, 150).Value = SwiftData.periodPresent
            swftComm.Parameters.Add("@APPLRULES", SqlDbType.VarChar, 100).Value = SwiftData.applicRules
            swftComm.Parameters.Add("@CONFRINSTR", SqlDbType.VarChar, 7).Value = SwiftData.confirmInstruct
            swftComm.Parameters.Add("@REIMBANK", SqlDbType.VarChar, 200).Value = SwiftData.correspSender
            swftComm.Parameters.Add("@INSTR2BANK", SqlDbType.VarChar, 800).Value = SwiftData.msgtoBank
            swftComm.Parameters.Add("@ADVISBANK", SqlDbType.VarChar, 200).Value = SwiftData.financialInstit
            swftComm.Parameters.Add("@SNDRECINF", SqlDbType.VarChar, 400).Value = SwiftData.send2recvInfo
            swftComm.Parameters.Add("@SWIFTMSG", SqlDbType.Text).Value = SwiftData.dat_All
            swftComm.Parameters.Add("@DATTAIL", SqlDbType.Text).Value = SwiftData.dat_Tail
            swftComm.Parameters.Add("@FILEDAT", SqlDbType.VarChar, 150).Value = SwiftData.OrigFile
            swftComm.Parameters.Add("@SWFSNDR", SqlDbType.VarChar, 12).Value = SwiftData.MsgSender
            swftComm.Parameters.Add("@SWFRCVR", SqlDbType.VarChar, 12).Value = SwiftData.MsgReceiver
            swftComm.Parameters.Add("@SWFMUR", SqlDbType.VarChar, 150).Value = SwiftData.MsgMUR
            swftComm.Parameters.Add("@SWFHED", SqlDbType.VarChar, 500).Value = SwiftData.MsgHead
            swftComm.Parameters.Add("@SWFTXT", SqlDbType.Text).Value = SwiftData.MsgText
            swftComm.Parameters.Add("@SWFTRAI", SqlDbType.VarChar, 200).Value = SwiftData.MsgTrailer
            Try
                If conx.State = ConnectionState.Closed Then conx.Open()
                swftComm.ExecuteNonQuery()
                _c.Informa("  SWIFT :  " & SwiftData.tipo_mensaje & " -> " & SwiftData.msgReference & " Ingresado a BBDD")
                SaveDataSwift_back = True
            Catch ex As Exception
                _c.Informa("  Error: " & ex.ToString)
                SaveDataSwift_back = False
            Finally
                'swftComm.Dispose()
            End Try
        ElseIf SwiftData.tipo_mensaje = "707" Then
            Dim conx As New SqlConnection(_cnxstr)
            Dim ermess
            Dim strSwf As String = "INSERT INTO T_ECX_SwiftMT707_Sale"
            strSwf = strSwf & " (F_ISN, F_FechaIngreso, F_tipo_mensaje, F_Hora_ingreso, F_Corresponsal, " & _
                        " F_timeType, F_timeTime, F_timeOffs, " & _
                        " F_msgReference, F_receiverRef, F_issBankRef, F_issBankBicName, " & _
                        " F_fechIssue, F_fechModif, F_numeModif, F_benefPreModif, " & _
                        " F_fechNewExpir, F_IncrMontoCant, F_IncrMontoMone, F_DecrMontoCant, " & _
                        " F_DecrMontoMone, F_NewMontoCant, F_NewMontoMone, F_PorceTolerMonto, " & _
                        " F_MaxCredCode, F_AdicMontoMemo, F_OnBoardMemo, F_ForTranspMemo, " & _
                        " F_LastFecEmba, F_ShipPeriodMemo, F_Mensaje79, F_send2recvInfo, " & _
                        " F_SwiftMessage, F_DataTail, F_OrigFile, " & _
                        " F_SwiftSender, F_SwiftReceiver, F_SwiftMUR, F_SwiftHead, " & _
                        " F_SwiftText, F_SwiftTrailer)"
            strSwf = strSwf & " VALUES (@ISN, @FECING, @TPMSG, @HRING, @CORRESP, " & _
                        " @TMTIP, @TMTIM, @TMOFF, " & _
                        " @MSGREF, @RELREF, @IBNKREF, @IBNKBIC, " & _
                        " @FECISS, @FECMOD, @NUMMOD, @BENPRE, " & _
                        " @FECNEWEXP, @INCRMONTO, @INCRMONEDA, @DECRMONTO, " & _
                        " @DECRMONEDA, @NEWMONTO, @NEWMONE, @PRCTOLM, " & _
                        " @MAXCREDMOD, @ADICMONTMEMO, @ONBRDMEM, @FORTRNSMEM, " & _
                        " @LASTFECEM, @SHIPPER, @MSJ79, @SNDRECINF, " & _
                        " @SWIFTMSG, @DATTAIL, @FILEDAT, " & _
                        " @SWFSNDR, @SWFRCVR, @SWFMUR, @SWFHED, @SWFTXT, @SWFTRAI)"

            'strSwf = strSwf & " (F_ISN, F_FechaIngreso, F_tipo_mensaje, F_Hora_ingreso, F_Corresponsal, " & _
            '           " F_timeType, F_timeTime, F_timeOffs, " & _
            '           " F_msgReference, F_receiverRef, F_issBankRef, F_issBankBicName, " & _
            '           " F_fechIssue, F_fechModif, F_numeModif, F_benefPreModif, " & _
            '           " F_IncrMontoCant, F_IncrMontoMone, F_DecrMontoCant, " & _
            '           " F_DecrMontoMone, F_NewMontoCant, F_NewMontoMone, F_PorceTolerMonto, " & _
            '           " F_MaxCredCode, F_AdicMontoMemo, F_OnBoardMemo, F_ForTranspMemo, " & _
            '           " F_ShipPeriodMemo, F_Mensaje79, F_send2recvInfo, " & _
            '           " F_SwiftMessage, F_DataTail, F_OrigFile, " & _
            '           " F_SwiftSender, F_SwiftReceiver, F_SwiftMUR, F_SwiftHead, " & _
            '           " F_SwiftText, F_SwiftTrailer)"
            'strSwf = strSwf & " VALUES (@ISN, @FECING, @TPMSG, @HRING, @CORRESP, " & _
            '            " @TMTIP, @TMTIM, @TMOFF, " & _
            '            " @MSGREF, @RELREF, @IBNKREF, @IBNKBIC, " & _
            '            " @FECISS, @FECMOD, @NUMMOD, @BENPRE, " & _
            '            " @INCRMONTO, @INCRMONEDA, @DECRMONTO, " & _
            '            " @DECRMONEDA, @NEWMONTO, @NEWMONE, @PRCTOLM, " & _
            '            " @MAXCREDMOD, @ADICMONTMEMO, @ONBRDMEM, @FORTRNSMEM, " & _
            '            " @SHIPPER, @MSJ79, @SNDRECINF, " & _
            '            " @SWIFTMSG, @DATTAIL, @FILEDAT, " & _
            '            " @SWFSNDR, @SWFRCVR, @SWFMUR, @SWFHED, @SWFTXT, @SWFTRAI)"
            Dim swftComm As New SqlCommand(strSwf, conx)
            swftComm.Parameters.Add("@ISN", SqlDbType.Int).Value = Val(SwiftData.ISN_mensaje)
            swftComm.Parameters.Add("@FECING", SqlDbType.DateTime).Value = SwiftData.Fecha_Ingreso
            swftComm.Parameters.Add("@TPMSG", SqlDbType.Int).Value = Val(SwiftData.tipo_mensaje)
            swftComm.Parameters.Add("@HRING", SqlDbType.VarChar).Value = Left(SwiftData.Hora_ingreso, 2) & ":" & Mid(SwiftData.Hora_ingreso, 3)
            swftComm.Parameters.Add("@CORRESP", SqlDbType.VarChar).Value = SwiftData.Corresponsal
            swftComm.Parameters.Add("@TMTIP", SqlDbType.VarChar).Value = SwiftData.timeType
            swftComm.Parameters.Add("@TMTIM", SqlDbType.VarChar).Value = Left(SwiftData.timeTime, 2) & ":" & Mid(SwiftData.timeTime, 3)
            swftComm.Parameters.Add("@TMOFF", SqlDbType.VarChar).Value = Left(SwiftData.timeOffs, 2) & ":" & Mid(SwiftData.timeOffs, 3)
            swftComm.Parameters.Add("@MSGREF", SqlDbType.VarChar).Value = SwiftData.msgReference
            swftComm.Parameters.Add("@RELREF", SqlDbType.VarChar).Value = SwiftData.relReference
            swftComm.Parameters.Add("@IBNKREF", SqlDbType.VarChar).Value = SwiftData.issuingBankReference
            swftComm.Parameters.Add("@IBNKBIC", SqlDbType.VarChar).Value = SwiftData.origBank
            swftComm.Parameters.Add("@FECISS", SqlDbType.DateTime).Value = toDate(SwiftData.dateofIssue)
            swftComm.Parameters.Add("@FECMOD", SqlDbType.DateTime).Value = toDate(SwiftData.dateofAmmend)
            swftComm.Parameters.Add("@NUMMOD", SqlDbType.Int).Value = Val(SwiftData.numofAmmend)
            swftComm.Parameters.Add("@BENPRE", SqlDbType.VarChar).Value = SwiftData.benName
            If Len(SwiftData.newdateofExpiry) > 0 Then
                swftComm.Parameters.Add("@FECNEWEXP", SqlDbType.DateTime).Value = toDate(SwiftData.newdateofExpiry)
            Else
                swftComm.Parameters.Add("@FECNEWEXP", SqlDbType.DateTime).Value = System.DBNull.Value
            End If
            If Trim("" & SwiftData.incrCrAmountCant) = "" Then
                tmomey = 0
            Else : tmomey = CDec(SwiftData.incrCrAmountCant.Replace(",", sim))
            End If
            swftComm.Parameters.Add("@INCRMONTO", SqlDbType.Money).Value = tmomey
            swftComm.Parameters.Add("@INCRMONEDA", SqlDbType.VarChar, 3).Value = SwiftData.incrCrAmountMon
            If Trim("" & SwiftData.instructAmount) = "" Then
                tmomey = 0
            Else : tmomey = CDec(SwiftData.instructAmount.Replace(",", sim))
            End If
            swftComm.Parameters.Add("@DECRMONTO", SqlDbType.Money).Value = tmomey
            swftComm.Parameters.Add("@DECRMONEDA", SqlDbType.VarChar, 3).Value = SwiftData.instructCurr
            If Trim("" & SwiftData.newCrAftAmmendAmount) = "" Then
                tmomey = 0
            Else : tmomey = CDec(SwiftData.newCrAftAmmendAmount.Replace(",", sim))
            End If
            swftComm.Parameters.Add("@NEWMONTO", SqlDbType.Money).Value = tmomey
            swftComm.Parameters.Add("@NEWMONE", SqlDbType.VarChar, 3).Value = SwiftData.newCrAftAmmendCurr
            swftComm.Parameters.Add("@PRCTOLM", SqlDbType.VarChar).Value = SwiftData.dat_39A
            swftComm.Parameters.Add("@MAXCREDMOD", SqlDbType.VarChar).Value = SwiftData.maxCredAmmoMsg
            swftComm.Parameters.Add("@ADICMONTMEMO", SqlDbType.VarChar).Value = SwiftData.addicAmmoCover
            swftComm.Parameters.Add("@ONBRDMEM", SqlDbType.VarChar).Value = SwiftData.newDispatchOpt
            swftComm.Parameters.Add("@FORTRNSMEM", SqlDbType.VarChar).Value = SwiftData.newFinalDest
            If Len(SwiftData.lateDateofShipp) > 0 Then
                swftComm.Parameters.Add("@LASTFECEM", SqlDbType.DateTime).Value = toDate(SwiftData.lateDateofShipp)
            Else
                swftComm.Parameters.Add("@LASTFECEM", SqlDbType.DateTime).Value = System.DBNull.Value
            End If
            swftComm.Parameters.Add("@SHIPPER", SqlDbType.VarChar).Value = SwiftData.shippPeriod
            swftComm.Parameters.Add("@MSJ79", SqlDbType.VarChar).Value = SwiftData.msgExplanation
            swftComm.Parameters.Add("@SNDRECINF", SqlDbType.VarChar).Value = SwiftData.send2recvInfo
            swftComm.Parameters.Add("@SWIFTMSG", SqlDbType.Text).Value = SwiftData.dat_All
            swftComm.Parameters.Add("@DATTAIL", SqlDbType.Text).Value = SwiftData.dat_Tail
            swftComm.Parameters.Add("@FILEDAT", SqlDbType.Text).Value = SwiftData.OrigFile
            swftComm.Parameters.Add("@SWFSNDR", SqlDbType.VarChar).Value = SwiftData.MsgSender
            swftComm.Parameters.Add("@SWFRCVR", SqlDbType.VarChar).Value = SwiftData.MsgReceiver
            swftComm.Parameters.Add("@SWFMUR", SqlDbType.VarChar).Value = SwiftData.MsgMUR
            swftComm.Parameters.Add("@SWFHED", SqlDbType.VarChar).Value = SwiftData.MsgHead
            swftComm.Parameters.Add("@SWFTXT", SqlDbType.Text).Value = SwiftData.MsgText
            swftComm.Parameters.Add("@SWFTRAI", SqlDbType.VarChar).Value = SwiftData.MsgTrailer

            Try
                If conx.State = ConnectionState.Closed Then conx.Open()
                swftComm.ExecuteNonQuery()
                _c.Informa("  SWIFT :  " & SwiftData.tipo_mensaje & " -> " & SwiftData.msgReference & " Ingresado a BBDD")
                SaveDataSwift_back = True
            Catch ex As Exception
                _c.Informa("  Error: " & ex.ToString)
                SaveDataSwift_back = False
            Finally
                'swftComm.Dispose()
            End Try
        ElseIf SwiftData.tipo_mensaje = "752" Then
            Dim conx As New SqlConnection(_cnxstr)
            Dim ermess
            Dim strSwf As String = "INSERT INTO T_ECX_SwiftMT752_Sale"
            strSwf = strSwf & " (F_ISN, F_FechaIngreso, F_tipo_mensaje, F_Hora_ingreso, F_Corresponsal, " & _
                        " F_timeType, F_timeTime, F_timeOffs, " & _
                        " F_docCredNum, F_presBnkRef, F_furtherIdentif, F_fechAdvise, " & _
                        " F_TotAdvMontoCant, F_TotAdvMontoMone, F_CargDeducMemo, F_fecNetAmmo, " & _
                        " F_NetAmmoCant, F_NetAmmoMone, F_correspSender, F_correspReceiver, " & _
                        " F_send2recvInfo, F_SwiftMessage, F_DataTail, F_OrigFile, F_SwiftSender, " & _
                        " F_SwiftReceiver, F_SwiftMUR, F_SwiftHead, F_SwiftText, F_SwiftTrailer)"
            strSwf = strSwf & " VALUES (@ISN, @FECING, @TPMSG, @HRING, @CORRESP, " & _
                        " @TMTIP, @TMTIM, @TMOFF, " & _
                        " @DOCCREDNUM, @PRSBNKREF, @FURTHID, @FECADV, " & _
                        " @NEWMONTO, @NEWMONE, @CARGDEDU, @FECNETAMO, " & _
                        " @NETMONTO, @NETMONE, @CORRSENDR, @CORRECVR, " & _
                        " @SNDRECINF, @SWIFTMSG, @DATTAIL, @FILEDAT, " & _
                        " @SWFSNDR, @SWFRCVR, @SWFMUR, @SWFHED, @SWFTXT, @SWFTRAI)"

            Dim swftComm As New SqlCommand(strSwf, conx)
            swftComm.Parameters.Add("@ISN", SqlDbType.Int).Value = CInt(SwiftData.ISN_mensaje)
            swftComm.Parameters.Add("@FECING", SqlDbType.DateTime).Value = SwiftData.Fecha_Ingreso
            swftComm.Parameters.Add("@TPMSG", SqlDbType.Int).Value = CInt(SwiftData.tipo_mensaje)
            swftComm.Parameters.Add("@HRING", SqlDbType.VarChar).Value = Left(SwiftData.Hora_ingreso, 2) & ":" & Mid(SwiftData.Hora_ingreso, 3)
            swftComm.Parameters.Add("@CORRESP", SqlDbType.VarChar).Value = SwiftData.Corresponsal
            swftComm.Parameters.Add("@TMTIP", SqlDbType.VarChar).Value = SwiftData.timeType
            swftComm.Parameters.Add("@TMTIM", SqlDbType.VarChar).Value = Left(SwiftData.timeTime, 2) & ":" & Mid(SwiftData.timeTime, 3)
            swftComm.Parameters.Add("@TMOFF", SqlDbType.VarChar).Value = Left(SwiftData.timeOffs, 2) & ":" & Mid(SwiftData.timeOffs, 3)
            swftComm.Parameters.Add("@DOCCREDNUM", SqlDbType.VarChar).Value = SwiftData.msgReference
            swftComm.Parameters.Add("@PRSBNKREF", SqlDbType.VarChar).Value = SwiftData.relReference
            swftComm.Parameters.Add("@FURTHID", SqlDbType.VarChar).Value = SwiftData.issuingBankReference
            swftComm.Parameters.Add("@FECADV", SqlDbType.DateTime).Value = toDate(SwiftData.dateofAmmend)
            If Trim("" & SwiftData.incrCrAmountCant) = "" Then
                tmomey = 0
            Else : tmomey = CDec(SwiftData.incrCrAmountCant.Replace(",", sim))
            End If
            swftComm.Parameters.Add("@NEWMONTO", SqlDbType.Money).Value = tmomey
            swftComm.Parameters.Add("@NEWMONE", SqlDbType.VarChar, 3).Value = SwiftData.incrCrAmountMon
            swftComm.Parameters.Add("@CARGDEDU", SqlDbType.VarChar).Value = SwiftData.cargDeducted
            swftComm.Parameters.Add("@FECNETAMO", SqlDbType.DateTime).Value = toDate(SwiftData.fecNetAmmo)
            If Trim("" & SwiftData.instructAmount) = "" Then
                tmomey = 0
            Else : tmomey = CDec(SwiftData.instructAmount.Replace(",", sim))
            End If
            swftComm.Parameters.Add("@NETMONTO", SqlDbType.Money).Value = tmomey
            swftComm.Parameters.Add("@NETMONE", SqlDbType.VarChar, 3).Value = SwiftData.instructCurr
            swftComm.Parameters.Add("@CORRSENDR", SqlDbType.VarChar).Value = SwiftData.correspSender
            swftComm.Parameters.Add("@CORRECVR", SqlDbType.VarChar).Value = SwiftData.correspReceiver
            swftComm.Parameters.Add("@SNDRECINF", SqlDbType.VarChar).Value = SwiftData.send2recvInfo
            swftComm.Parameters.Add("@SWIFTMSG", SqlDbType.Text).Value = SwiftData.dat_All
            swftComm.Parameters.Add("@DATTAIL", SqlDbType.Text).Value = SwiftData.dat_Tail
            swftComm.Parameters.Add("@FILEDAT", SqlDbType.Text).Value = SwiftData.OrigFile
            swftComm.Parameters.Add("@SWFSNDR", SqlDbType.VarChar).Value = SwiftData.MsgSender
            swftComm.Parameters.Add("@SWFRCVR", SqlDbType.VarChar).Value = SwiftData.MsgReceiver
            swftComm.Parameters.Add("@SWFMUR", SqlDbType.VarChar).Value = SwiftData.MsgMUR
            swftComm.Parameters.Add("@SWFHED", SqlDbType.VarChar).Value = SwiftData.MsgHead
            swftComm.Parameters.Add("@SWFTXT", SqlDbType.Text).Value = SwiftData.MsgText
            swftComm.Parameters.Add("@SWFTRAI", SqlDbType.VarChar).Value = SwiftData.MsgTrailer

            Try
                If conx.State = ConnectionState.Closed Then conx.Open()
                swftComm.ExecuteNonQuery()
                _c.Informa("  SWIFT :  " & SwiftData.tipo_mensaje & " -> " & SwiftData.msgReference & " Ingresado a BBDD")
                SaveDataSwift_back = True
            Catch ex As Exception
                _c.Informa("  Error: " & ex.ToString)
                SaveDataSwift_back = False
            Finally
                'swftComm.Dispose()
            End Try
        ElseIf SwiftData.tipo_mensaje = "910" Then
            Dim conx As New SqlConnection(_cnxstr)
            Dim ermess
            Dim strSwf As String = "INSERT INTO T_ECX_SwiftMT910_Sale"
            strSwf = strSwf & " (F_ISN, F_FechaIngreso, F_tipo_mensaje, F_Hora_ingreso, F_Corresponsal, " & _
                        " F_msgReference, F_relReference, F_Valuta, F_Moneda, " & _
                        " F_Monto, F_orderingCusAcc, F_orderingCusName, " & _
                        " F_origBank, F_accountIdent, F_intermediariaInstit, " & _
                        " F_send2recvInfo, F_regulReport, " & _
                        " F_SwiftMessage, F_DataTail, F_OrigFile, " & _
                        " F_SwiftSender, F_SwiftReceiver, F_SwiftMUR, F_SwiftHead, F_SwiftText, F_SwiftTrailer)"
            strSwf = strSwf & " VALUES (@ISN, @FECING, @TPMSG, @HRING, @CORRESP, " & _
                        " @MSGREF, @RELREF, @VALUTA, @MONEDA, @MONTO, " & _
                        " @CTAORDCUS, @NOMORDCUS, " & _
                        " @ORDBNK, @ACOIDEN, @INTERINST, " & _
                        " @SNDRECINF, @REGUREP, @SWIFTMSG, @DATTAIL, @FILEDAT, " & _
                        " @SWFSNDR, @SWFRCVR, @SWFMUR, @SWFHED, @SWFTXT, @SWFTRAI)"

            Dim swftComm As New SqlCommand(strSwf, conx)
            swftComm.Parameters.Add("@ISN", SqlDbType.Int).Value = CInt(SwiftData.ISN_mensaje)
            swftComm.Parameters.Add("@FECING", SqlDbType.DateTime).Value = SwiftData.Fecha_Ingreso
            swftComm.Parameters.Add("@TPMSG", SqlDbType.Int).Value = CInt(SwiftData.tipo_mensaje)
            swftComm.Parameters.Add("@HRING", SqlDbType.VarChar).Value = Left(SwiftData.Hora_ingreso, 2) & ":" & Mid(SwiftData.Hora_ingreso, 3)
            swftComm.Parameters.Add("@CORRESP", SqlDbType.VarChar).Value = SwiftData.Corresponsal
            swftComm.Parameters.Add("@MSGREF", SqlDbType.VarChar).Value = SwiftData.msgReference
            swftComm.Parameters.Add("@RELREF", SqlDbType.VarChar).Value = SwiftData.relReference
            swftComm.Parameters.Add("@VALUTA", SqlDbType.DateTime).Value = toDate(SwiftData.Valuta)
            swftComm.Parameters.Add("@MONEDA", SqlDbType.VarChar, 3).Value = SwiftData.Moneda
            If Trim("" & SwiftData.Monto) = "" Then
                tmomey = 0
            Else : tmomey = CDec(SwiftData.Monto.Replace(",", sim))
            End If
            swftComm.Parameters.Add("@MONTO", SqlDbType.Money).Value = tmomey
            swftComm.Parameters.Add("@CTAORDCUS", SqlDbType.VarChar).Value = SwiftData.orderingCusAcc
            swftComm.Parameters.Add("@NOMORDCUS", SqlDbType.VarChar).Value = SwiftData.orderingCusName
            swftComm.Parameters.Add("@ORDBNK", SqlDbType.VarChar).Value = SwiftData.origBank
            swftComm.Parameters.Add("@ACOIDEN", SqlDbType.VarChar).Value = SwiftData.accountIdent
            swftComm.Parameters.Add("@INTERINST", SqlDbType.VarChar).Value = SwiftData.intermediariaInstit
            swftComm.Parameters.Add("@SNDRECINF", SqlDbType.VarChar).Value = SwiftData.send2recvInfo
            swftComm.Parameters.Add("@REGUREP", SqlDbType.VarChar).Value = SwiftData.regulReport
            swftComm.Parameters.Add("@SWIFTMSG", SqlDbType.Text).Value = SwiftData.dat_All
            swftComm.Parameters.Add("@DATTAIL", SqlDbType.Text).Value = SwiftData.dat_Tail
            swftComm.Parameters.Add("@FILEDAT", SqlDbType.Text).Value = SwiftData.OrigFile
            swftComm.Parameters.Add("@SWFSNDR", SqlDbType.VarChar).Value = SwiftData.MsgSender
            swftComm.Parameters.Add("@SWFRCVR", SqlDbType.VarChar).Value = SwiftData.MsgReceiver
            swftComm.Parameters.Add("@SWFMUR", SqlDbType.VarChar).Value = SwiftData.MsgMUR
            swftComm.Parameters.Add("@SWFHED", SqlDbType.VarChar).Value = SwiftData.MsgHead
            swftComm.Parameters.Add("@SWFTXT", SqlDbType.Text).Value = SwiftData.MsgText
            swftComm.Parameters.Add("@SWFTRAI", SqlDbType.VarChar).Value = SwiftData.MsgTrailer

            Try
                If conx.State = ConnectionState.Closed Then conx.Open()
                swftComm.ExecuteNonQuery()
                _c.Informa("  SWIFT :  " & SwiftData.tipo_mensaje & " -> " & SwiftData.msgReference & " Ingresado a BBDD")
                SaveDataSwift_back = True
            Catch ex As Exception
                _c.Informa("  Error: " & ex.ToString)
                SaveDataSwift_back = False
            Finally
                'swftComm.Dispose()
            End Try
        ElseIf SwiftData.tipo_mensaje = "799" Then
            Dim conx As New SqlConnection(_cnxstr)
            Dim ermess
            Dim strSwf As String = "INSERT INTO T_ECX_SwiftMT799_Sale"
            strSwf = strSwf & " (F_ISN, F_FechaIngreso, F_tipo_mensaje, F_Hora_ingreso, F_Corresponsal, " & _
                        " F_msgReference, F_relReference, F_Explanation, " & _
                        " F_SwiftMessage, F_DataTail, F_OrigFile, " & _
                        " F_SwiftSender, F_SwiftReceiver, F_SwiftMUR, F_SwiftHead, F_SwiftText, F_SwiftTrailer)"
            strSwf = strSwf & " VALUES (@ISN, @FECING, @TPMSG, @HRING, @CORRESP, " & _
                        " @MSGREF, @RELREF, @EXPLAN, @SWIFTMSG, @DATTAIL, @FILEDAT, " & _
                        " @SWFSNDR, @SWFRCVR, @SWFMUR, @SWFHED, @SWFTXT, @SWFTRAI)"

            Dim swftComm As New SqlCommand(strSwf, conx)
            swftComm.Parameters.Add("@ISN", SqlDbType.Int).Value = CInt(SwiftData.ISN_mensaje)
            swftComm.Parameters.Add("@FECING", SqlDbType.DateTime).Value = SwiftData.Fecha_Ingreso
            swftComm.Parameters.Add("@TPMSG", SqlDbType.Int).Value = CInt(SwiftData.tipo_mensaje)
            swftComm.Parameters.Add("@HRING", SqlDbType.VarChar).Value = Left(SwiftData.Hora_ingreso, 2) & ":" & Mid(SwiftData.Hora_ingreso, 3)
            swftComm.Parameters.Add("@CORRESP", SqlDbType.VarChar).Value = SwiftData.Corresponsal
            swftComm.Parameters.Add("@MSGREF", SqlDbType.VarChar).Value = SwiftData.msgReference
            swftComm.Parameters.Add("@RELREF", SqlDbType.VarChar).Value = SwiftData.relReference
            swftComm.Parameters.Add("@EXPLAN", SqlDbType.VarChar, 1800).Value = SwiftData.msgExplanation
            swftComm.Parameters.Add("@SWIFTMSG", SqlDbType.Text).Value = SwiftData.dat_All
            swftComm.Parameters.Add("@DATTAIL", SqlDbType.Text).Value = SwiftData.dat_Tail
            swftComm.Parameters.Add("@FILEDAT", SqlDbType.Text).Value = SwiftData.OrigFile
            swftComm.Parameters.Add("@SWFSNDR", SqlDbType.VarChar).Value = SwiftData.MsgSender
            swftComm.Parameters.Add("@SWFRCVR", SqlDbType.VarChar).Value = SwiftData.MsgReceiver
            swftComm.Parameters.Add("@SWFMUR", SqlDbType.VarChar).Value = SwiftData.MsgMUR
            swftComm.Parameters.Add("@SWFHED", SqlDbType.VarChar).Value = SwiftData.MsgHead
            swftComm.Parameters.Add("@SWFTXT", SqlDbType.Text).Value = SwiftData.MsgText
            swftComm.Parameters.Add("@SWFTRAI", SqlDbType.VarChar).Value = SwiftData.MsgTrailer

            Try
                If conx.State = ConnectionState.Closed Then conx.Open()
                swftComm.ExecuteNonQuery()
                _c.Informa("  SWIFT :  " & SwiftData.tipo_mensaje & " -> " & SwiftData.msgReference & " Ingresado a BBDD")
                SaveDataSwift_back = True
            Catch ex As Exception
                _c.Informa("  Error: " & ex.ToString)
                SaveDataSwift_back = False
            Finally
                'swftComm.Dispose()
            End Try
            'msg_799 = "Señores," & vbCrLf
            'msg_799 = msg_799 & "   Se ha recibido mensaje swift 799, que adjuntamos para su revisión." & vbCrLf
            'msg_799 = msg_799 & vbCrLf
            'msg_799 = msg_799 & SwiftData.MsgHead & vbCrLf & "-------------------------------------------" & vbCrLf & SwiftData.MsgText
            ''msg_799 = msg_799 & SwiftData.dat_All
            'If Not Envia("Mensaje Swift 799 Recibido", msg_799) Then
            '    _c.Informa(" Mensaje MT799 enviado a los respectivos Mails ")
            'Else
            '    _c.Informa(" No pudo ser enviado el Mensaje MT799 a los respectivos Mails ")
            'End If


        Else
            _c.Informa("   -- Tipo Swift: " & SwiftData.tipo_mensaje & " -> " & SwiftData.Moneda)
            SaveDataSwift_back = False
        End If
        _c.Informa("----------------------------------------------------------------------")
    End Function

    Function GetBankFromBIC(ByVal bic As String) As String
        Dim dtTipo As DataTable = Nothing
        Dim sBnkName As String, sBnkAddr As String
        Dim sqlDataAdapter As SqlDataAdapter
        Dim conn As New SqlConnection(_cnxstr)
        Dim strSQL As String = "exec spr_swf_GetBank @BIC"
        Dim myCommand As New SqlCommand(strSQL, conn)
        myCommand.Parameters.Add("@BIC", SqlDbType.VarChar).Value = bic
        Try
            sqlDataAdapter = New SqlDataAdapter(myCommand)
            dtTipo = New DataTable
            sqlDataAdapter.Fill(dtTipo)
            sBnkName = Trim(dtTipo.Rows(0).Item(0))
            sBnkAddr = Trim(dtTipo.Rows(0).Item(1))
            If sBnkAddr.Length > 0 Then
                GetBankFromBIC = sBnkName & vbCrLf & sBnkAddr
            ElseIf sBnkName.Length > 0 Then
                GetBankFromBIC = sBnkName
            Else
                GetBankFromBIC = ""
            End If
        Catch ex As Exception
            GetBankFromBIC = ""
        Finally
            conn.Close()
            myCommand.Dispose()
            conn.Dispose()
        End Try
        conn = Nothing
    End Function
    Function GetTagText(ByVal iMessType As Integer, ByVal sMessTag As String) As String
        Dim dtTipo As DataTable = Nothing
        Dim sqlDataAdapter As SqlDataAdapter
        Dim conn As New SqlConnection(_cnxstr)
        Dim strSQL As String = "exec spr_swf_mensaje @IdMsg, @MsTag"
        Dim myCommand As New SqlCommand(strSQL, conn)
        myCommand.Parameters.Add("@IdMsg", SqlDbType.Int).Value = iMessType
        myCommand.Parameters.Add("@MsTag", SqlDbType.VarChar).Value = sMessTag
        Try
            sqlDataAdapter = New SqlDataAdapter(myCommand)
            dtTipo = New DataTable
            sqlDataAdapter.Fill(dtTipo)
            GetTagText = Trim(dtTipo.Rows(0).Item(0))
        Catch ex As Exception
            GetTagText = ""
        Finally
            conn.Close()
            myCommand.Dispose()
            conn.Dispose()
        End Try
        conn = Nothing
    End Function
    Function GetTagTEng(ByVal iMessType As Integer, ByVal sMessTag As String) As String
        Dim dtTipo As DataTable = Nothing
        Dim sqlDataAdapter As SqlDataAdapter
        Dim conn As New SqlConnection(_cnxstr)
        Dim strSQL As String = "exec spr_swf_message @IdMsg, @MsTag"
        Dim myCommand As New SqlCommand(strSQL, conn)
        myCommand.Parameters.Add("@IdMsg", SqlDbType.Int).Value = iMessType
        myCommand.Parameters.Add("@MsTag", SqlDbType.VarChar).Value = sMessTag
        Try
            sqlDataAdapter = New SqlDataAdapter(myCommand)
            dtTipo = New DataTable
            sqlDataAdapter.Fill(dtTipo)
            GetTagTEng = Trim(dtTipo.Rows(0).Item(0))
        Catch ex As Exception
            GetTagTEng = ""
        Finally
            conn.Close()
            myCommand.Dispose()
            conn.Dispose()
        End Try
        conn = Nothing
    End Function
    Sub rConsola(ByVal tx As String)
        If _verboso Then System.Console.Write(tx) : System.Console.Write(vbCrLf)
    End Sub
    Function toDate(ByVal yyDate As String) 'As DateTime
        If Len(yyDate) < 8 Then
            toDate = System.DBNull.Value '.DateTime.MinValue
        Else
            toDate = DateSerial(CInt(Left(yyDate, 4)), CInt(Mid(yyDate, 5, 2)), CInt(Mid(yyDate, 7, 2)))
        End If
    End Function
    Public Function AppPath( _
            Optional ByVal backSlash As Boolean = False _
            ) As String
        ' System.Reflection.Assembly.GetExecutingAssembly...
        Dim s As String = _
            IO.Path.GetDirectoryName( _
            System.Reflection.Assembly.GetExecutingAssembly.GetCallingAssembly.Location)
        ' si hay que añadirle el backslash
        If backSlash Then
            s &= "\"
        End If
        Return s
    End Function
    Sub AvisaDestinatarioComex(ByVal SwiftData As SwiftsJob.SwiftData)
        Dim mails As String, ok As Boolean, bnkname As String, arsxt
        Dim subject As String, mensaje As String, cliente As String
        Dim BankDest As String, Cliname As String, Operef As String, BenefNom As String, MsgHead As String, MsgText As String
        BankDest = SwiftData.MsgReceiver
        bnkname = GetBankFromBIC(BankDest)
        If bnkname.Length > 0 Then
            arsxt = Split(bnkname, vbCrLf)
            bnkname = arsxt(0)
        Else
            arsxt = Split(BankDest, vbCrLf)
            bnkname = arsxt(0)
        End If
        Cliname = SwiftData.orderingCusName
        arsxt = Split(Cliname, vbCrLf)
        cliente = arsxt(0)
        Operef = SwiftData.msgReference
        If Operef.StartsWith("F") Then
            subject = "Envío de Transferencia" '"Aviso de Transferencia de Fondos"
            mensaje = "Dears Sirs" & vbCrLf & vbCrLf
            mensaje = mensaje & "      Please find attached copy of message forwarded to "
            mensaje = mensaje & bnkname '" AMERICAN EXPRESS BANK LTD"
            mensaje = mensaje & ", by instruction of our customer "
            mensaje = mensaje & cliente '" COMERCIAL MUNDO LCV LTDA."
            mensaje = mensaje & ", corresponding to payment "
            mensaje = mensaje & " of your collection order, kindly request you "
            mensaje = mensaje & " to check advising us in the case of any discrepancy." & vbCrLf
            mensaje = mensaje & " " & vbCrLf
            mensaje = mensaje & " Atentamente" & vbCrLf
            mensaje = mensaje & " " & vbCrLf
            mensaje = mensaje & " Oscar Herrera" & vbCrLf
            mensaje = mensaje & " " & vbCrLf
            mensaje = mensaje & " Jefe Importaciones Banco Security" & vbCrLf
            mensaje = mensaje & " oherrera@security.cl" & vbCrLf
            mensaje = mensaje & " Teléfono 56-02-5844848" & vbCrLf & vbCrLf

            mensaje = mensaje & SwiftData.MsgHEng
            mensaje = mensaje & vbCrLf
            mensaje = mensaje & "--------------------------------------------" & vbCrLf
            mensaje = mensaje & SwiftData.MsgTEng & vbCrLf
            mails = GetEmailFromPayOrder(Operef)
            If mails.Length > 0 Then
                mails = mails.Replace("·", ";")
                mails = mails.Replace(",", ";")
                'mails = "sefuentes@security.cl;icontreras@security.cl;bdiaz@security.cl"
                ok = Envia(mails, subject, mensaje)
                ' envia msails
            End If
        ElseIf Operef.StartsWith("OPS") Then
            subject = "Envío de Transferencia" '"Aviso de Transferencia de Fondos"
            mensaje = "Dears Sirs" & vbCrLf & vbCrLf
            mensaje = mensaje & "      Please find attached copy of message forwarded to "
            mensaje = mensaje & bnkname '" AMERICAN EXPRESS BANK LTD"
            mensaje = mensaje & ", by instruction of our customer "
            mensaje = mensaje & cliente '" COMERCIAL MUNDO LCV LTDA."
            mensaje = mensaje & ", corresponding to payment "
            mensaje = mensaje & " of your collection order, kindly request you "
            mensaje = mensaje & " to check advising us in the case of any discrepancy." & vbCrLf
            mensaje = mensaje & " " & vbCrLf
            mensaje = mensaje & " Atentamente" & vbCrLf
            mensaje = mensaje & " " & vbCrLf
            mensaje = mensaje & " Oscar Herrera" & vbCrLf
            mensaje = mensaje & " " & vbCrLf
            mensaje = mensaje & " Jefe Importaciones Banco Security" & vbCrLf
            mensaje = mensaje & " oherrera@security.cl" & vbCrLf
            mensaje = mensaje & " Teléfono 56-02-5844848" & vbCrLf & vbCrLf

            mensaje = mensaje & SwiftData.MsgHEng
            mensaje = mensaje & vbCrLf
            mensaje = mensaje & "--------------------------------------------" & vbCrLf
            mensaje = mensaje & SwiftData.MsgTEng & vbCrLf
            mails = GetEmailFromCorePayOrder(Operef)
            If mails.Length > 0 Then
                mails = mails.Replace("·", ";")
                mails = mails.Replace(",", ";")
                'mails = "sefuentes@security.cl;icontreras@security.cl;bdiaz@security.cl"
                ok = Envia(mails, subject, mensaje)
                ' envia msails
            End If

        ElseIf Operef.ToUpper.StartsWith("IC") Then
            subject = "Envío de Transferencia (IC)" '"Aviso de Transferencia de Fondos"
            mensaje = "Dears Sirs" & vbCrLf & vbCrLf
            mensaje = mensaje & "      Please find attached copy of message forwarded to "
            mensaje = mensaje & bnkname '" AMERICAN EXPRESS BANK LTD"
            mensaje = mensaje & ", by instruction of our customer "
            mensaje = mensaje & cliente '" COMERCIAL MUNDO LCV LTDA."
            mensaje = mensaje & ", corresponding to payment "
            mensaje = mensaje & " of your collection order, kindly request you "
            mensaje = mensaje & " to check advising us in the case of any discrepancy." & vbCrLf
            mensaje = mensaje & " " & vbCrLf
            mensaje = mensaje & " Atentamente" & vbCrLf
            mensaje = mensaje & " " & vbCrLf
            mensaje = mensaje & " Oscar Herrera" & vbCrLf
            mensaje = mensaje & " " & vbCrLf
            mensaje = mensaje & " Jefe Importaciones Banco Security" & vbCrLf
            mensaje = mensaje & " oherrera@security.cl" & vbCrLf
            mensaje = mensaje & " Teléfono 56-02-5844848" & vbCrLf & vbCrLf

            mensaje = mensaje & SwiftData.MsgHEng
            mensaje = mensaje & vbCrLf
            mensaje = mensaje & "--------------------------------------------" & vbCrLf
            mensaje = mensaje & SwiftData.MsgTEng & vbCrLf
            BenefNom = SwiftData.benName
            arsxt = Split(BenefNom, vbCrLf)
            BenefNom = arsxt(0)
            mails = GetEmailFromBenefPagoIC(Operef, BenefNom)
            If mails.Length > 0 Then
                mails = mails.Replace("·", ";")
                mails = mails.Replace(",", ";")
                'mails = "sefuentes@security.cl;icontreras@security.cl;bdiaz@security.cl"
                ok = Envia(mails, subject, mensaje)
                ' envia msails
            End If
        ElseIf 1 = 0 And Operef.ToUpper.StartsWith("I") And IsNumeric(Mid(Operef, 2)) Then
            subject = "Letter Credit sent from: " & cliente '"Aviso de Transferencia de Fondos"
            mensaje = "Dears Sirs" & vbCrLf & vbCrLf
            mensaje = mensaje & "      Please find attached copy of message forwarded to "
            mensaje = mensaje & bnkname '" AMERICAN EXPRESS BANK LTD"
            mensaje = mensaje & ", by instruction of our customer "
            mensaje = mensaje & cliente '" COMERCIAL MUNDO LCV LTDA."
            'mensaje = mensaje & ", corresponding to payment "
            'mensaje = mensaje & " of your collection order, kindly request you "
            'mensaje = mensaje & " to check advising us in the case of any discrepancy." & vbCrLf
            mensaje = mensaje & " " & vbCrLf
            mensaje = mensaje & " Atentamente" & vbCrLf
            mensaje = mensaje & " " & vbCrLf
            mensaje = mensaje & " Oscar Herrera" & vbCrLf
            mensaje = mensaje & " " & vbCrLf
            mensaje = mensaje & " Jefe Importaciones Banco Security" & vbCrLf
            mensaje = mensaje & " oherrera@security.cl" & vbCrLf
            mensaje = mensaje & " Teléfono 56-02-5844848" & vbCrLf & vbCrLf

            mensaje = mensaje & SwiftData.MsgHEng
            mensaje = mensaje & vbCrLf
            mensaje = mensaje & "--------------------------------------------" & vbCrLf
            mensaje = mensaje & SwiftData.MsgTEng & vbCrLf
            BenefNom = SwiftData.benName
            arsxt = Split(BenefNom, vbCrLf)
            BenefNom = arsxt(0)
            mails = GetEmailFromBenefLCImp(Operef)
            If mails.Length > 0 Then
                mails = mails.Replace("·", ";")
                mails = mails.Replace(",", ";")
                'mails = "sefuentes@security.cl;icontreras@security.cl;bdiaz@security.cl"
                ok = Envia(mails, subject, mensaje)
                ' envia msails
            End If
        End If
    End Sub
    Function GetEmailFromBenefLCImp(ByVal OpeRef As String) As String
        Dim dtTipo As DataTable = Nothing
        Dim sMailName As String
        Dim sqlDataAdapter As SqlDataAdapter
        Dim conn As New SqlConnection(_cnxstr)
        Dim strSQL As String = "exec spr_swf_GetDestinoAvisoBenefCobImp @REF, @NOM"
        Dim myCommand As New SqlCommand(strSQL, conn)
        myCommand.Parameters.Add("@REF", SqlDbType.VarChar).Value = OpeRef
        'myCommand.Parameters.Add("@NOM", SqlDbType.VarChar).Value = BenefNom
        Try
            sqlDataAdapter = New SqlDataAdapter(myCommand)
            dtTipo = New DataTable
            sqlDataAdapter.Fill(dtTipo)
            If dtTipo.Rows.Count > 0 Then
                sMailName = Trim(dtTipo.Rows(0).Item(0))
            Else
                sMailName = ""
            End If
            'sBnkAddr = Trim(dtTipo.Rows(0).Item(1))
            If sMailName.Length > 0 Then
                GetEmailFromBenefLCImp = sMailName
            Else
                GetEmailFromBenefLCImp = ""
            End If
        Catch ex As Exception
            GetEmailFromBenefLCImp = ""
        Finally
            conn.Close()
            myCommand.Dispose()
            conn.Dispose()
        End Try
        conn = Nothing
    End Function
    Function GetEmailFromBenefPagoIC(ByVal OpeRef As String, ByVal BenefNom As String) As String
        Dim dtTipo As DataTable = Nothing
        Dim sMailName As String
        Dim sqlDataAdapter As SqlDataAdapter
        Dim conn As New SqlConnection(_cnxstr)
        Dim strSQL As String = "exec spr_swf_GetDestinoAvisoBenefCobImp @REF, @NOM"
        Dim myCommand As New SqlCommand(strSQL, conn)
        myCommand.Parameters.Add("@REF", SqlDbType.VarChar).Value = OpeRef
        myCommand.Parameters.Add("@NOM", SqlDbType.VarChar).Value = BenefNom
        Try
            sqlDataAdapter = New SqlDataAdapter(myCommand)
            dtTipo = New DataTable
            sqlDataAdapter.Fill(dtTipo)
            If dtTipo.Rows.Count > 0 Then
                sMailName = Trim(dtTipo.Rows(0).Item(0))
            Else
                sMailName = ""
            End If
            'sBnkAddr = Trim(dtTipo.Rows(0).Item(1))
            If sMailName.Length > 0 Then
                GetEmailFromBenefPagoIC = sMailName
            Else
                GetEmailFromBenefPagoIC = ""
            End If
        Catch ex As Exception
            GetEmailFromBenefPagoIC = ""
        Finally
            conn.Close()
            myCommand.Dispose()
            conn.Dispose()
        End Try
        conn = Nothing
    End Function

    Function GetEmailFromPayOrder(ByVal OpeRef As String) As String
        Dim dtTipo As DataTable = Nothing
        Dim sMailName As String
        Dim sqlDataAdapter As SqlDataAdapter
        Dim conn As New SqlConnection(_cnxstr)
        Dim strSQL As String = "exec spr_swf_GetDestinoAvisoOpago @REF"
        Dim myCommand As New SqlCommand(strSQL, conn)
        myCommand.Parameters.Add("@REF", SqlDbType.VarChar).Value = OpeRef
        Try
            sqlDataAdapter = New SqlDataAdapter(myCommand)
            dtTipo = New DataTable
            sqlDataAdapter.Fill(dtTipo)
            If dtTipo.Rows.Count > 0 Then
                sMailName = Trim(dtTipo.Rows(0).Item(0))
            Else
                sMailName = ""
            End If
            'sBnkAddr = Trim(dtTipo.Rows(0).Item(1))
            If sMailName.Length > 0 Then
                GetEmailFromPayOrder = sMailName
            Else
                GetEmailFromPayOrder = ""
            End If
        Catch ex As Exception
            GetEmailFromPayOrder = ""
        Finally
            conn.Close()
            myCommand.Dispose()
            conn.Dispose()
        End Try
        conn = Nothing
    End Function
    Function GetEmailFromCorePayOrder(ByVal OpeRef As String) As String
        Dim dtTipo As DataTable = Nothing
        Dim sMailName As String
        Dim sqlDataAdapter As SqlDataAdapter
        Dim conn As New SqlConnection(_cnxstr)
        Dim strSQL As String = "exec spr_swf_GetDestinoCoreComexTrfSale @REF"
        Dim myCommand As New SqlCommand(strSQL, conn)
        myCommand.Parameters.Add("@REF", SqlDbType.VarChar).Value = OpeRef
        Try
            sqlDataAdapter = New SqlDataAdapter(myCommand)
            dtTipo = New DataTable
            sqlDataAdapter.Fill(dtTipo)
            If dtTipo.Rows.Count > 0 Then
                sMailName = Trim(dtTipo.Rows(0).Item(0))
            Else
                sMailName = ""
            End If
            'sBnkAddr = Trim(dtTipo.Rows(0).Item(1))
            If sMailName.Length > 0 Then
                GetEmailFromCorePayOrder = sMailName
            Else
                GetEmailFromCorePayOrder = ""
            End If
        Catch ex As Exception
            GetEmailFromCorePayOrder = ""
        Finally
            conn.Close()
            myCommand.Dispose()
            conn.Dispose()
        End Try
        conn = Nothing
    End Function

    Function Envia(ByVal destino As String, ByVal subject As String, ByVal mensaje As String) As Boolean
        Dim conx As New SqlConnection(_cnxstr)
        Dim ermess
        'declare @rc int
        'exec @rc = master.dbo.xp_smtp_sendmail
        '    @FROM       = N'MyEmail@MyDomain.com',
        '    @FROM_NAME  = N'Joe Mailman',
        '    @TO         = N'MyFriend@HisDomain.com',
        '    @CC         = N'MyOtherFriend@HisDomain.com',

        '    @BCC        = N'MyEmail@MyDomain.com',

        '    @priority   = N'HIGH',
        '    @subject    = N'Hello SQL Server SMTP Mail',
        '    @message    = N'Goodbye MAPI, goodbye Outlook',
        '    @type       = N'text/plain',
        '    @attachments= N'c:\attachment1.txt;c:\attachment2.txt',
        '    @server     = N'mail.mydomain.com'
        'select RC = @rc

        'Dim strSwf As String = "EXEC master.dbo.xp_smtp_sendmail @FROM='banco@security.cl',@TO=@destino,@BCC='sefuentes@security.cl;icontreras@security.cl',@subject=@sujeto,@server='smtp.security.cl',@message=@f_message"
        Dim strSwf As String = "EXEC master.dbo.xp_smtp_sendmail @FROM='banco@security.cl',@FROM_NAME='Banco Security - Chile',@TO=@destino,@BCC=@blindcopy,@subject=@sujeto,@server='smtp.security.cl',@message=@f_message"
        Dim swftComm As New SqlCommand(strSwf, conx)

        swftComm.Parameters.Add("@destino", SqlDbType.VarChar, 200).Value = destino '"bdiaz@security.cl" '"sebastian@fuentesv.com" '
        swftComm.Parameters.Add("@blindcopy", SqlDbType.VarChar, 200).Value = _smtpdestin ' "sefuentes@security.cl;icontreras@security.cl" ' 
        swftComm.Parameters.Add("@sujeto", SqlDbType.VarChar, 300).Value = subject
        swftComm.Parameters.Add("@f_message", SqlDbType.VarChar, 7000).Value = mensaje
        Try
            If conx.State = ConnectionState.Closed Then conx.Open()
            _c.Informa("  Enviando:  " & subject & " -> " & destino)
            swftComm.ExecuteNonQuery()
        Catch ex As Exception
            _c.Informa("  Error: " & ex.ToString)
            Envia = True
        Finally
            conx.Close()
            swftComm.Dispose()
            conx.Dispose()
        End Try


    End Function
    Function Envia2(ByVal subject As String, ByVal mensaje As String) As Boolean
        Dim eMail As MailMessage, i As Integer
        eMail = New MailMessage
        '        eMail.To = destino '"nadie@alla.com"
        eMail.From = IIf(_smtpsender.Length = 0, "yo@deaqui.org", _smtpsender)
        'If _smtpdestin.Length > 0 Then
        '    If _smtpdestin.IndexOf(";") > 0 Then
        '        Dim cc2 = _smtpdestin.Split(";")
        '        For i = 0 To UBound(cc2) '- 1
        '            eMail.To = cc2(i)
        '        Next
        '    Else
        eMail.To = _smtpdestin
        '    End If
        'End If
        eMail.Subject = subject '"Alertas Normalización"
        eMail.Body = mensaje '"Hola"
        eMail.BodyFormat = MailFormat.Text '.Html

        SmtpMail.SmtpServer = IIf(_smtpserver.Length = 0, "smpt.security.cl", _smtpserver)
        Try
            SmtpMail.Send(eMail)
        Catch ehttp As System.Web.HttpException
            _c.Informa(ehttp.Message)
            _c.Informa("full error message:")
            _c.Informa(ehttp.ToString())
            Return True
        Catch ex As Exception
            _c.Informa("Error: " & ex.Message)
            Return True
        End Try

    End Function

End Module
