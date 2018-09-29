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
        objIniFile = New SwiftsJob.IniFile(AppPath(True) & "CargaSwiftsEntrantes.ini")
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
        If _verboso Then
            System.Console.WriteLine("Log Directory: " & _logdir)
            System.Console.WriteLine("Input Directory: " & _inputdir)
            System.Console.WriteLine("Output Directory: " & _swiftdir)
            System.Console.WriteLine("String Conexion BBDD: " & _cnxstr)
            System.Console.WriteLine("Server SMTP: " & _smtpserver)
            System.Console.WriteLine("Destinos SMTP: " & _smtpdestin)
            System.Console.WriteLine("Verbose Mode")
            LogOption = 2
            If _cnxstr = "" Then
                System.Console.WriteLine("¡Falta String de Conexión a DDBB! ")
                System.Console.ReadLine()
                Return 1
            End If
            If _logdir = "" Then
                System.Console.WriteLine("¡Falta Log Directory! ")
                System.Console.ReadLine()
                Return 1
            End If
            If _inputdir = "" Then
                System.Console.WriteLine("¡Falta Directorio de Recepcion de Archivos! ")
                System.Console.ReadLine()
                Return 1
            End If
            If _swiftdir = "" Then
                System.Console.WriteLine("¡Falta Directorio de Destino! ")
                System.Console.ReadLine()
                Return 1
            End If
            If _smtpserver = "" Then
                System.Console.WriteLine("¡Falta Definir Servidor SMTP! ")
                System.Console.ReadLine()
                Return 1
            End If
            If _smtpdestin = "" Then
                System.Console.WriteLine("¡Falta Definir Destinos SMTP! ")
                System.Console.ReadLine()
                Return 1
            End If
            If _smtpsender = "" Then
                System.Console.WriteLine("¡Falta Definir Cuenta de Envio SMTP! ")
                System.Console.ReadLine()
                Return 1
            End If
        End If
        _c = New SwiftsJob.AppLog(LogOption, _logdir & _logFilName)

        OkOut = ProcJob()

        If _verboso Then
            System.Console.WriteLine("Presione Enter para Terminar")
            System.Console.ReadLine()
        End If
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
        Dim desdra As IO.FileInfo
        Dim arcswf As IO.FileInfo
        Dim datswf As SwiftsJob.SwiftData
        Dim okfile As Boolean = False
        Dim idf As Integer = 0
        Try
            di = New IO.DirectoryInfo(_inputdir)
            diar1 = di.GetFiles("*.out")
        Catch ex As Exception
            System.Console.WriteLine(ex.ToString)
            System.Console.WriteLine("Presione Enter para Terminar")
            System.Console.ReadLine()
        End Try
        For Each dra In diar1
            rConsola(dra.FullName)
            okfile = False : idf = 0
            If ProcFile(_inputdir & dra.Name, dra.LastWriteTime) Then
                Try
                    desdra = New FileInfo(_swiftdir & dra.Name)
                    _c.Informa("Relocalizando Archivo Swifts: " & Left(dra.Name, dra.Name.Length - 4) & " en " & _swiftdir)
                    Do While Not okfile
                        dra2 = New FileInfo(_swiftdir & Left(dra.Name, dra.Name.Length - 4) & "_" & idf & ".out")
                        If Not dra2.Exists() Then
                            okfile = True
                        Else
                            idf = idf + 1
                        End If
                    Loop
                    dra.CopyTo(_swiftdir & Left(dra.Name, dra.Name.Length - 4) & "_" & idf & ".out")
                    'If Not desdra.Exists() Then
                    '    dra.CopyTo(_swiftdir & dra.Name)
                    'End If
                    dra.Delete()
                    ProcJob = True
                    _c.Informa("Ok")
                Catch ex As Exception
                    _c.Informa("Error: " & ex.ToString)
                    ProcJob = False
                End Try
            Else
                ProcJob = False
                Exit For
            End If
        Next

    End Function

    Function ProcFile(ByVal fil As String, ByVal fechafile As Date) As Boolean
        Dim SwiftDat As SwiftsJob.SwiftData
        Dim SwiftMT As Object
        Dim errMess As String, artxt, arsxt, a
        Dim p As Integer, MtNum As Integer
        Dim tstr As String
        Dim oFile As System.IO.File
        Dim oRead As System.IO.StreamReader
        'datswf.NomArchivo = arcswf.Name
        _c.Informa("Iniciando Proceso Archivo Swifts: " & fil & " en " & _swiftdir)
        oRead = oFile.OpenText(fil)

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
                    '                rConsola(" · " & Mid(arLin(i), k, 1) & ") => " & Asc(Mid(arLin(i), k, 1)))
                End If
                i = i + 1
            End If
        End While
        oRead.Close()
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
                    If Left(SwiftDat.MsgText, 2) = vbCrLf Then SwiftDat.MsgText = Mid(SwiftDat.MsgText, 3)
                    If SaveDataSwift(SwiftDat, errMess) Then
                        SwiftDat = New SwiftsJob.SwiftData
                    Else
                        ProcFile = False
                        Exit Function
                        'SwiftDat = New SwiftsJob.SwiftData
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
                Case Left(arLin(i), Len(":20:")) = ":20:"
                    SwiftDat.dat_20 = Mid(arLin(i), Len(":20:") + 1)
                    SwiftDat.msgReference = Mid(arLin(i), Len(":20:") + 1)
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "   20: " & GetTagText(MtNum, "20")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & SwiftDat.msgReference
                    'SwiftMT.msgReference = SwiftDat.msgReference
                    _c.Informa("   -- Swift Ref: " & SwiftDat.msgReference)
                Case Left(arLin(i), Len(":21:")) = ":21:"
                    SwiftDat.dat_21 = Mid(arLin(i), Len(":21:") + 1)
                    SwiftDat.relReference = SwiftDat.dat_21
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "   21: " & GetTagText(MtNum, "21")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & SwiftDat.relReference
                    'SwiftMT.relReference = SwiftDat.relReference
                Case Left(arLin(i), Len(":23B:")) = ":23B:"
                    SwiftDat.dat_23B = Mid(arLin(i), Len(":23B:") + 1)
                    SwiftDat.bankOperatCode = Left(SwiftDat.dat_23B, 4)
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  23B: " & GetTagText(MtNum, "23B")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & SwiftDat.bankOperatCode
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
                    'SwiftMT.instructCodeRefr = SwiftDat.instructCodeRefr
                Case Left(arLin(i), Len(":25:")) = ":25:"
                    SwiftDat.dat_25 = Mid(arLin(i), Len(":25:") + 1)
                    SwiftDat.accountIdent = SwiftDat.dat_25
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "   25: " & GetTagText(MtNum, "25")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & SwiftDat.accountIdent
                    'SwiftMT.transacType = SwiftDat.transacType
                Case Left(arLin(i), Len(":26T:")) = ":26T:"
                    SwiftDat.dat_26T = Mid(arLin(i), Len(":26T:") + 1)
                    SwiftDat.transacType = Left(SwiftDat.dat_26T, 3)
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  26T: " & GetTagText(MtNum, "26T")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & SwiftDat.transacType
                    'SwiftMT.transacType = SwiftDat.transacType
                Case Left(arLin(i), Len(":32A:")) = ":32A:"
                    SwiftDat.Valuta = IIf(Mid(arLin(i), Len(":32A:") + 1, 2) <= "50", "20" & Mid(arLin(i), Len(":32A:") + 1, 6), "19" & Mid(arLin(i), Len(":32A:") + 1, 6))
                    SwiftDat.Moneda = Mid(arLin(i), Len(":32A:") + 7, 3)
                    SwiftDat.Monto = Mid(arLin(i), Len(":32A:") + 10)
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  32A: " & GetTagText(MtNum, "32A")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Date    : " & SwiftDat.Valuta
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Currency: " & SwiftDat.Moneda
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Amount  : " & SwiftDat.Monto
                Case Left(arLin(i), Len(":33B:")) = ":33B:"
                    SwiftDat.dat_33B = Mid(arLin(i), Len(":33B:") + 1)
                    SwiftDat.instructCurr = Left(SwiftDat.dat_33B, 3)
                    SwiftDat.instructAmount = Mid(SwiftDat.dat_33B, 4)
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  33B: " & GetTagText(MtNum, "33B")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Currency: " & SwiftDat.instructCurr
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Amount  : " & SwiftDat.instructAmount
                Case Left(arLin(i), Len(":36:")) = ":36:"
                    SwiftDat.dat_36 = Mid(arLin(i), Len(":36:") + 1)
                    SwiftDat.exchRate = SwiftDat.dat_36
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "   36: " & GetTagText(MtNum, "36")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & SwiftDat.exchRate
                Case Left(arLin(i), Len(":40E:")) = ":40E:"
                    SwiftDat.dat_40E = Mid(arLin(i), Len(":40E:") + 1)
                    SwiftDat.applicRules = SwiftDat.dat_40E
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  40E: " & GetTagText(MtNum, "40E")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & SwiftDat.applicRules
                Case Left(arLin(i), Len(":40F:")) = ":40F:"
                    SwiftDat.dat_40F = Mid(arLin(i), Len(":40F:") + 1)
                    SwiftDat.applicRules = SwiftDat.dat_40F
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  40F: " & GetTagText(MtNum, "40F")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & SwiftDat.applicRules
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
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Account : " & SwiftDat.orderingCusAcc
                    tstr = SwiftDat.orderingCusName : arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "         Name  : " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                 " & arsxt(k)
                        End If
                    Next
                Case Left(arLin(i), Len(":50F:")) = ":50F:"
                    SwiftDat.dat_50K = Mid(arLin(i), Len(":50F:") + 1)
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
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  50F: " & GetTagText(MtNum, "50K")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Account : " & SwiftDat.orderingCusAcc
                    tstr = SwiftDat.orderingCusName : arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Name    : " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
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
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Account : " & SwiftDat.orderingCusAcc
                    tstr = SwiftDat.orderingCusName : arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Name    : " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                        End If
                    Next
                Case UCase(Left(arLin(i), Len(":51A:"))) = ":51A:"
                    SwiftDat.dat_51A = Mid(arLin(i), Len(":51A:") + 1)
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  51A: " & GetTagText(MtNum, "51A")
                    SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & SwiftDat.dat_51A
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
                    tstr = GetBankFromBIC(SwiftDat.origBank)
                    arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
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
                    'SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "            tf: " & SwiftDat.origBankPI
                    tstr = SwiftDat.origBank : arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Name/Dir: " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
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
                    tstr = GetBankFromBIC(SwiftDat.correspSender)
                    arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
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
                    'SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Identf: " & SwiftDat.correspSenderPI
                    tstr = SwiftDat.correspSender : arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Name/Dir: " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
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
                    'SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Identf: " & SwiftDat.correspSenderPI
                    tstr = SwiftDat.correspSender : arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Name/Dir: " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
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
                    tstr = GetBankFromBIC(SwiftDat.correspReceiver)
                    arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
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
                    'SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Identf: " & SwiftDat.correspReceiverPI
                    tstr = SwiftDat.correspReceiver : arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Name/Dir: " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
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
                    'SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Identf: " & SwiftDat.correspReceiverPI
                    tstr = SwiftDat.correspReceiver : arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Name/Dir: " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
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
                    tstr = GetBankFromBIC(SwiftDat.terceraInstit)
                    arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
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
                    tstr = GetBankFromBIC(SwiftDat.intermediariaInstit)
                    arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                        End If
                    Next
                Case UCase(Left(arLin(i), Len(":56D:"))) = ":56D:"
                    SwiftDat.dat_56D = Mid(arLin(i), Len(":56D:") + 1)
                    Try
                        If Left(SwiftDat.dat_56D, 1) = "/" Then
                            arsxt = Split(Mid(SwiftDat.dat_56D, 2), vbCrLf)  ' ordering Customer
                            SwiftDat.intermediariaInstitPI = arsxt(0)
                            SwiftDat.intermediariaInstit = arsxt(1)
                            For k = 2 To UBound(arsxt)
                                SwiftDat.intermediariaInstit = SwiftDat.intermediariaInstit & vbCrLf & arsxt(k)
                            Next
                            'SwiftDat.intermediariaInstit = Mid(SwiftDat.intermediariaInstit, 3)
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
                    'SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Identf: " & SwiftDat.intermediariaInstitPI
                    tstr = SwiftDat.intermediariaInstit : arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Name/Dir: " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
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
                    tstr = GetBankFromBIC(SwiftDat.financialInstit)
                    arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
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
                    'SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Identf: " & SwiftDat.financialInstitPI
                    tstr = SwiftDat.financialInstit : arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Name/Dir: " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
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
                    'SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Identf: " & SwiftDat.financialInstitPI
                    tstr = SwiftDat.financialInstit : arsxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(arsxt)
                        If k = 0 Then
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Name/Dir: " & arsxt(k)
                        Else
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
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
                        tstr = GetBankFromBIC(SwiftDat.benefInstit)
                        arsxt = Split(tstr, vbCrLf)
                        For k = 0 To UBound(arsxt)
                            If k = 0 Then
                                SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                            Else
                                SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
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
                        'SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Identf: " & SwiftDat.benefInstitPI
                        tstr = SwiftDat.benefInstit : arsxt = Split(tstr, vbCrLf)
                        For k = 0 To UBound(arsxt)
                            If k = 0 Then
                                SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Name/Dir: " & arsxt(k)
                            Else
                                SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
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
                        SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Account : " & SwiftDat.benAccount
                        tstr = SwiftDat.benName : arsxt = Split(tstr, vbCrLf)
                        For k = 0 To UBound(arsxt)
                            If k = 0 Then
                                SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Name/Dir: " & arsxt(k)
                            Else
                                SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
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
                        SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Account :" & SwiftDat.benAccount
                        tstr = SwiftDat.benName : arsxt = Split(tstr, vbCrLf)
                        For k = 0 To UBound(arsxt)
                            If k = 0 Then
                                SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Name/Dir: " & arsxt(k)
                            Else
                                SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "                  " & arsxt(k)
                            End If
                        Next
                Case Left(arLin(i), Len(":70:")) = ":70:"
                        SwiftDat.dat_70 = Mid(arLin(i), Len(":70:") + 1)
                        SwiftDat.remitInfo = SwiftDat.dat_70
                        SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "   70: " & GetTagText(MtNum, "70")
                        tstr = SwiftDat.remitInfo : arsxt = Split(tstr, vbCrLf)
                        For k = 0 To UBound(arsxt)
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & arsxt(k)
                        Next
                Case UCase(Left(arLin(i), Len(":71A:"))) = ":71A:"
                        SwiftDat.dat_71A = Mid(arLin(i), Len(":71A:") + 1)
                        SwiftDat.cargoTo = SwiftDat.dat_71A
                        SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  71A: " & GetTagText(MtNum, "71A")
                        SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & SwiftDat.cargoTo
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
                        SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Currency: " & Left(tstr, 3)
                        SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Amount  : " & Mid(tstr, 4)
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
                        SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Currency: " & Left(tstr, 3)
                        SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        Amount  : " & Mid(tstr, 4)
                Case Left(arLin(i), Len(":72:")) = ":72:"
                        SwiftDat.dat_72 = Mid(arLin(i), Len(":72:") + 1)
                        SwiftDat.send2recvInfo = SwiftDat.dat_72
                        SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "    72: " & GetTagText(MtNum, "72")
                        tstr = SwiftDat.send2recvInfo : arsxt = Split(tstr, vbCrLf)
                        For k = 0 To UBound(arsxt)
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "         " & arsxt(k)
                        Next
                Case UCase(Left(arLin(i), Len(":77B:"))) = ":77B:"
                        SwiftDat.dat_77B = Mid(arLin(i), Len(":77B:") + 1)
                        SwiftDat.regulReport = SwiftDat.dat_77B
                        SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  77B: " & GetTagText(MtNum, "77B")
                        tstr = SwiftDat.regulReport : arsxt = Split(tstr, vbCrLf)
                        For k = 0 To UBound(arsxt)
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & arsxt(k)
                        Next
                Case UCase(Left(arLin(i), Len(":77T:"))) = ":77T:"
                        SwiftDat.dat_77T = Mid(arLin(i), Len(":77T:") + 1)
                        SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  77T: " & GetTagText(MtNum, "77T")
                        SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & SwiftDat.dat_77T
                Case UCase(Left(arLin(i), Len(":79:"))) = ":79:"
                        SwiftDat.dat_79 = Mid(arLin(i), Len(":79:") + 1)
                        SwiftDat.msgExplanation = SwiftDat.dat_79
                        SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "  79: " & GetTagText(MtNum, "79")
                        tstr = SwiftDat.msgExplanation : arsxt = Split(tstr, vbCrLf)
                        For k = 0 To UBound(arsxt)
                            SwiftDat.MsgText = SwiftDat.MsgText & vbCrLf & "        " & arsxt(k)
                        Next
                Case Left(arLin(i), Len("{1:F")) = "{1:F"
                        k = InStr(arLin(i), "{2:")
                        If k > 0 Then
                            SwiftDat.tipo_mensaje = Mid(arLin(i), k + 4, 3)
                            SwiftDat.Fecha_Ingreso = fechafile
                            If SwiftDat.tipo_mensaje = "103" Then
                                SwiftDat.Hora_ingreso = Mid(arLin(i), k + 7, 4)
                                SwiftDat.Corresponsal = Mid(arLin(i), k + 17, 12)
                            ElseIf SwiftDat.tipo_mensaje = "202" Then
                                SwiftDat.Hora_ingreso = Mid(arLin(i), k + 7, 4)
                                SwiftDat.Corresponsal = Mid(arLin(i), k + 7, 12)
                            ElseIf SwiftDat.tipo_mensaje = "910" Then
                                SwiftDat.Hora_ingreso = Mid(arLin(i), k + 7, 4)
                                SwiftDat.Corresponsal = Mid(arLin(i), k + 17, 12)
                            ElseIf SwiftDat.tipo_mensaje = "799" Then
                                SwiftDat.Hora_ingreso = Mid(arLin(i), k + 7, 4)
                                SwiftDat.Corresponsal = Mid(arLin(i), k + 17, 12)
                            Else
                                '   SwiftMT = New SwiftsJob.SwiftData
                            End If
                            MtNum = CInt(SwiftDat.tipo_mensaje)
                            j = InStr(k, arLin(i), "}")
                            linea = Mid(arLin(i), k + 3, (j - k) - 3)
                            tstr = Mid(linea, 15, 12)
                            tstr = Left(tstr, 8) & Mid(tstr, 10, 3)
                            SwiftDat.MsgSender = tstr
                        End If
                    k = InStr(10, arLin(i), "{1:")
                    If k = 0 Then k = 1
                    j = InStr(k, arLin(i), "}")
                    linea = Mid(arLin(i), k + 3, (j - k) - 3)
                    SwiftDat.ISN_mensaje = Right(CStr(Year(fechafile)), 2) & Right(linea, 6)
                    tstr = Mid(linea, 4, 12)
                    tstr = Left(tstr, 8) & Mid(tstr, 10, 3)
                    SwiftDat.MsgReceiver = tstr
                    SwiftDat.MsgMUR = ""
                    Try
                        k = InStr(10, arLin(i), "{3:") : j = InStr(k, arLin(i), "}")
                        linea = Mid(arLin(i), k + 3, (j - k) - 3)
                        k = InStr(linea, ":")
                        SwiftDat.MsgMUR = Mid(linea, k + 1)
                    Catch ex As Exception
                    End Try
                    SwiftDat.MsgHead = "Swift OUTPUT : FIN " & SwiftDat.tipo_mensaje & " " & GetTagText(CInt(SwiftDat.tipo_mensaje), "--")
                    SwiftDat.MsgHead = SwiftDat.MsgHead & vbCrLf & "Sender       : " & SwiftDat.MsgSender
                    tstr = GetBankFromBIC(SwiftDat.MsgSender)
                    artxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(artxt)
                        SwiftDat.MsgHead = SwiftDat.MsgHead & vbCrLf & "         " & artxt(k)
                    Next
                    SwiftDat.MsgHead = SwiftDat.MsgHead & vbCrLf & "Receiver     : " & SwiftDat.MsgReceiver
                    tstr = GetBankFromBIC(SwiftDat.MsgReceiver)
                    artxt = Split(tstr, vbCrLf)
                    For k = 0 To UBound(artxt)
                        SwiftDat.MsgHead = SwiftDat.MsgHead & vbCrLf & "         " & artxt(k)
                    Next
                    If SwiftDat.MsgMUR.Length > 0 Then
                        SwiftDat.MsgHead = SwiftDat.MsgHead & vbCrLf & "MUR :  " & SwiftDat.MsgMUR
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

        rConsola(":Mensaje Swift Tipo MT" & SwiftData.tipo_mensaje & "----------------------------")
        rConsola("     -- ISN_mensaje > " & SwiftData.ISN_mensaje)
        rConsola("     -- tipo_mensaje > " & SwiftData.tipo_mensaje)
        rConsola("     -- Hora_ingreso > " & SwiftData.Hora_ingreso)
        rConsola("     -- Corresponsal > " & SwiftData.Corresponsal)
        rConsola("     -- timeType > " & SwiftData.timeType)
        rConsola("     -- timeTime > " & SwiftData.timeTime)
        rConsola("     -- timeOffs (Offset) > " & SwiftData.timeOffs)
        rConsola("     -- msgReference > " & SwiftData.msgReference)
        rConsola("     -- relReference > " & SwiftData.relReference)
        rConsola("     -- bankOperat    Code > " & SwiftData.bankOperatCode)
        rConsola("     -- instructCodeCode > " & SwiftData.instructCodeCode)
        rConsola("     -- instructCodeRefr > " & SwiftData.instructCodeRefr)
        rConsola("     -- transacType > " & SwiftData.transacType)
        rConsola("     -- Valuta > " & SwiftData.Valuta)
        rConsola("     -- Moneda > " & SwiftData.Moneda)
        rConsola("     -- Monto > " & SwiftData.Monto)
        rConsola("     -- accountIdent > " & SwiftData.accountIdent)
        rConsola("     -- instructCurr > " & SwiftData.instructCurr)
        rConsola("     -- instructAmount > " & SwiftData.instructAmount)
        rConsola("     -- exchRate > " & SwiftData.exchRate)
        rConsola("     -- orderingCusAcc > " & SwiftData.orderingCusAcc)
        rConsola("     -- orderingCusName > " & SwiftData.orderingCusName)
        rConsola("     -- origBankPI > " & SwiftData.origBankPI)
        rConsola("     -- origBank > " & SwiftData.origBank)
        rConsola("     -- correspSenderPI > " & SwiftData.correspSenderPI)
        rConsola("     -- correspSender > " & SwiftData.correspSender)
        rConsola("     -- correspReceiverPI > " & SwiftData.correspReceiverPI)
        rConsola("     -- correspReceiver > " & SwiftData.correspReceiver)
        rConsola("     -- terceraInstit > " & SwiftData.terceraInstit)
        rConsola("     -- intermediariaInstitPI > " & SwiftData.intermediariaInstitPI)
        rConsola("     -- intermediariaInstit > " & SwiftData.intermediariaInstit)
        rConsola("     -- financialInstitPI > " & SwiftData.financialInstitPI)
        rConsola("     -- financialInstit > " & SwiftData.financialInstit)
        rConsola("     -- benefInstitPI > " & SwiftData.benefInstitPI)
        rConsola("     -- benefInstitPI > " & SwiftData.benefInstit)
        rConsola("     -- benAccount > " & SwiftData.benAccount)
        rConsola("     -- benName > " & SwiftData.benName)
        rConsola("     -- remitInfo > " & SwiftData.remitInfo)
        rConsola("     -- cargoTo > " & SwiftData.cargoTo)
        rConsola("     -- senderChargCurr > " & SwiftData.senderChargCurr)
        rConsola("     -- senderChargAmou > " & SwiftData.senderChargAmou)
        rConsola("     -- recverChargCurr > " & SwiftData.recverChargCurr)
        rConsola("     -- recverChargAmou > " & SwiftData.recverChargAmou)
        rConsola("     -- send2recvInfo > " & SwiftData.send2recvInfo)
        rConsola("     -- regulReport > " & SwiftData.regulReport)
        rConsola("     -- Data Tail > " & SwiftData.dat_Tail)
        rConsola("-----------------------------------")
        'SaveDataSwift = True
        'Exit Function
        If SwiftData.tipo_mensaje = "103" And SwiftData.Moneda <> "CLP" Then
            Dim conx As New SqlConnection(_cnxstr)
            Dim ermess
            Dim strSwf As String = "INSERT INTO T_ECX_SwiftMT103"
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
                        " F_SwiftSender, F_SwiftReceiver, F_SwiftMUR, F_SwiftHead, F_SwiftText, F_SwiftTrailer)"
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
                        " @SWFSNDR, @SWFRCVR, @SWFMUR, @SWFHED, @SWFTXT, @SWFTRAI)"

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

            Try
                If conx.State = ConnectionState.Closed Then conx.Open()
                swftComm.ExecuteNonQuery()
                _c.Informa("  SWIFT :  " & SwiftData.tipo_mensaje & " -> " & SwiftData.msgReference & " Ingresado a BBDD")
                SaveDataSwift = True
            Catch ex As Exception
                _c.Informa("  Error: " & ex.ToString)
                If Not Envia("Swifts Recibidos ", ex.ToString) Then
                    _c.Informa(" Mensaje Error " & ex.ToString & " enviado a los respectivos Mails ")
                Else
                    _c.Informa(" No pudo ser enviado el  Mensaje Error")
                End If
                SaveDataSwift = False
            Finally
                'swftComm.Dispose()
            End Try
        ElseIf SwiftData.tipo_mensaje = "202" And SwiftData.Moneda <> "CLP" Then
            Dim conx As New SqlConnection(_cnxstr)
            Dim ermess
            Dim strSwf As String = "INSERT INTO T_ECX_SwiftMT202"
            strSwf = strSwf & " (F_ISN, F_FechaIngreso, F_FechaProceso, F_tipo_mensaje, F_Hora_ingreso, F_Corresponsal, " & _
                        " F_timeType, F_timeTime, F_timeOffs,  " & _
                        " F_msgReference, F_relReference, F_Valuta, F_Moneda, " & _
                        " F_Monto, F_origBank, F_correspSender, F_correspReceiver, " & _
                        " F_intermediariaInstit, F_financialInstit, F_benefInstit, " & _
                        " F_send2recvInfo, F_regulReport, " & _
                        " F_SwiftMessage, F_DataTail, F_OrigFile, " & _
                        " F_SwiftSender, F_SwiftReceiver, F_SwiftMUR, F_SwiftHead, F_SwiftText, F_SwiftTrailer)"
            'VALUES     (3, '3', '3', '3', '3', '3', '3', '3', '3', '3', '3', 3, '3', 3, '3', 3, 3, '3', '3', '3', '3', '3', '3', '3', '3', '3', '3', '3', '3', '3', '3', 3, '3', 3, '3', '3', '3')
            strSwf = strSwf & " VALUES (@ISN, @FECING, GETDATE(), @TPMSG, @HRING, @CORRESP, " & _
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
            swftComm.Parameters.Add("@HRING", SqlDbType.VarChar, 5).Value = Left(SwiftData.Hora_ingreso, 2) & ":" & Mid(SwiftData.Hora_ingreso, 3)
            swftComm.Parameters.Add("@CORRESP", SqlDbType.VarChar, 50).Value = SwiftData.Corresponsal
            swftComm.Parameters.Add("@TMTIP", SqlDbType.VarChar, 8).Value = SwiftData.timeType
            swftComm.Parameters.Add("@TMTIM", SqlDbType.VarChar, 5).Value = Left(SwiftData.timeTime, 2) & ":" & Mid(SwiftData.timeTime, 3)
            swftComm.Parameters.Add("@TMOFF", SqlDbType.VarChar, 5).Value = Left(SwiftData.timeOffs, 2) & ":" & Mid(SwiftData.timeOffs, 3)
            swftComm.Parameters.Add("@MSGREF", SqlDbType.VarChar, 16).Value = SwiftData.msgReference
            swftComm.Parameters.Add("@RELREF", SqlDbType.VarChar, 16).Value = SwiftData.relReference
            swftComm.Parameters.Add("@VALUTA", SqlDbType.DateTime).Value = toDate(SwiftData.Valuta)
            swftComm.Parameters.Add("@MONEDA", SqlDbType.VarChar, 3).Value = SwiftData.Moneda
            If Trim("" & SwiftData.Monto) = "" Then
                tmomey = 0
            Else : tmomey = CDec(SwiftData.Monto.Replace(",", sim))
            End If
            swftComm.Parameters.Add("@MONTO", SqlDbType.Money).Value = tmomey
            swftComm.Parameters.Add("@ORDBNK", SqlDbType.VarChar, 300).Value = SwiftData.origBank
            swftComm.Parameters.Add("@CORRSENDR", SqlDbType.VarChar, 300).Value = SwiftData.correspSender
            swftComm.Parameters.Add("@CORRECVR", SqlDbType.VarChar, 300).Value = SwiftData.correspReceiver
            swftComm.Parameters.Add("@INTERINST", SqlDbType.VarChar, 300).Value = SwiftData.intermediariaInstit
            swftComm.Parameters.Add("@FINANINST", SqlDbType.VarChar, 300).Value = SwiftData.financialInstit
            swftComm.Parameters.Add("@BENFINST", SqlDbType.VarChar, 300).Value = SwiftData.benefInstit
            swftComm.Parameters.Add("@SNDRECINF", SqlDbType.VarChar, 300).Value = SwiftData.send2recvInfo
            swftComm.Parameters.Add("@REGUREP", SqlDbType.VarChar, 300).Value = SwiftData.regulReport
            swftComm.Parameters.Add("@SWIFTMSG", SqlDbType.Text).Value = SwiftData.dat_All
            swftComm.Parameters.Add("@DATTAIL", SqlDbType.Text).Value = SwiftData.dat_Tail
            swftComm.Parameters.Add("@FILEDAT", SqlDbType.Text).Value = SwiftData.OrigFile
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
                SaveDataSwift = True
            Catch ex As Exception
                _c.Informa("  Error: " & ex.ToString)
                If Not Envia("Swifts Recibidos ", ex.ToString) Then
                    _c.Informa(" Mensaje Error " & ex.ToString & " enviado a los respectivos Mails ")
                Else
                    _c.Informa(" No pudo ser enviado el  Mensaje Error")
                End If
                SaveDataSwift = False
            Finally
                'swftComm.Dispose()
            End Try
        ElseIf SwiftData.tipo_mensaje = "910" And SwiftData.Moneda <> "CLP" Then

            Dim conx As New SqlConnection(_cnxstr)
            Dim ermess
            Dim strSwf As String = "INSERT INTO T_ECX_SwiftMT910"
            strSwf = strSwf & " (F_ISN, F_FechaIngreso, F_FechaProceso, F_tipo_mensaje, F_Hora_ingreso, F_Corresponsal, " & _
                        " F_msgReference, F_relReference, F_Valuta, F_Moneda, " & _
                        " F_Monto, F_orderingCusAcc, F_orderingCusName, " & _
                        " F_origBank, F_accountIdent, F_intermediariaInstit, " & _
                        " F_send2recvInfo, F_regulReport, " & _
                        " F_SwiftMessage, F_DataTail, F_OrigFile, " & _
                        " F_SwiftSender, F_SwiftReceiver, F_SwiftMUR, F_SwiftHead, F_SwiftText, F_SwiftTrailer)"
            strSwf = strSwf & " VALUES (@ISN, @FECING, GETDATE(), @TPMSG, @HRING, @CORRESP, " & _
                        " @MSGREF, @RELREF, @VALUTA, @MONEDA, @MONTO, " & _
                        " @CTAORDCUS, @NOMORDCUS, " & _
                        " @ORDBNK, @ACOIDEN, @INTERINST, " & _
                        " @SNDRECINF, @REGUREP, @SWIFTMSG, @DATTAIL, @FILEDAT, " & _
                        " @SWFSNDR, @SWFRCVR, @SWFMUR, @SWFHED, @SWFTXT, @SWFTRAI)"

            Dim swftComm As New SqlCommand(strSwf, conx)
            swftComm.Parameters.Add("@ISN", SqlDbType.Int).Value = CInt(SwiftData.ISN_mensaje)
            swftComm.Parameters.Add("@FECING", SqlDbType.DateTime).Value = SwiftData.Fecha_Ingreso
            swftComm.Parameters.Add("@TPMSG", SqlDbType.Int).Value = CInt(SwiftData.tipo_mensaje)
            swftComm.Parameters.Add("@HRING", SqlDbType.VarChar, 5).Value = Left(SwiftData.Hora_ingreso, 2) & ":" & Mid(SwiftData.Hora_ingreso, 3)
            swftComm.Parameters.Add("@CORRESP", SqlDbType.VarChar, 50).Value = SwiftData.Corresponsal
            swftComm.Parameters.Add("@MSGREF", SqlDbType.VarChar, 16).Value = SwiftData.msgReference
            swftComm.Parameters.Add("@RELREF", SqlDbType.VarChar, 16).Value = SwiftData.relReference
            swftComm.Parameters.Add("@VALUTA", SqlDbType.DateTime).Value = toDate(SwiftData.Valuta)
            swftComm.Parameters.Add("@MONEDA", SqlDbType.VarChar, 3).Value = SwiftData.Moneda
            If Trim("" & SwiftData.Monto) = "" Then
                tmomey = 0
            Else : tmomey = CDec(SwiftData.Monto.Replace(",", sim))
            End If
            swftComm.Parameters.Add("@MONTO", SqlDbType.Money).Value = tmomey
            swftComm.Parameters.Add("@CTAORDCUS", SqlDbType.VarChar, 34).Value = SwiftData.orderingCusAcc
            swftComm.Parameters.Add("@NOMORDCUS", SqlDbType.VarChar, 150).Value = SwiftData.orderingCusName
            swftComm.Parameters.Add("@ORDBNK", SqlDbType.VarChar, 150).Value = SwiftData.origBank
            swftComm.Parameters.Add("@ACOIDEN", SqlDbType.VarChar, 35).Value = SwiftData.accountIdent
            swftComm.Parameters.Add("@INTERINST", SqlDbType.VarChar, 150).Value = SwiftData.intermediariaInstit
            swftComm.Parameters.Add("@SNDRECINF", SqlDbType.VarChar, 210).Value = SwiftData.send2recvInfo
            swftComm.Parameters.Add("@REGUREP", SqlDbType.VarChar, 1000).Value = SwiftData.regulReport
            swftComm.Parameters.Add("@SWIFTMSG", SqlDbType.Text).Value = SwiftData.dat_All
            swftComm.Parameters.Add("@DATTAIL", SqlDbType.Text).Value = SwiftData.dat_Tail
            swftComm.Parameters.Add("@FILEDAT", SqlDbType.Text).Value = SwiftData.OrigFile
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
                SaveDataSwift = True
            Catch ex As Exception
                _c.Informa("  Error: " & ex.ToString)
                If Not Envia("Swifts Recibidos ", ex.ToString) Then
                    _c.Informa(" Mensaje Error " & ex.ToString & " enviado a los respectivos Mails ")
                Else
                    _c.Informa(" No pudo ser enviado el  Mensaje Error")
                End If

                SaveDataSwift = False
            Finally
                'swftComm.Dispose()
            End Try
            'If SaveDataSwift Then
            '    Dim msg_910 As String
            '    msg_910 = "Señores," & vbCrLf
            '    msg_910 = msg_910 & "   Se ha recibido mensaje swift " & SwiftData.tipo_mensaje & ", que adjuntamos para su revisión." & vbCrLf
            '    msg_910 = msg_910 & vbCrLf
            '    msg_910 = msg_910 & SwiftData.MsgHead & vbCrLf & "-------------------------------------------" & vbCrLf & SwiftData.MsgText
            '    'msg_910 = msg_910 & SwiftData.dat_All
            '    If Not Envia("Mensaje Swift " & SwiftData.tipo_mensaje & " Recibido", msg_910) Then
            '        _c.Informa(" Mensaje MT" & SwiftData.tipo_mensaje & " enviado a los respectivos Mails ")
            '    Else
            '        _c.Informa(" No pudo ser enviado el Mensaje MT" & SwiftData.tipo_mensaje & " a los respectivos Mails ")
            '    End If
            'End If

        ElseIf Right(SwiftData.tipo_mensaje, 2) = "99" Then

            Dim conx As New SqlConnection(_cnxstr)
            Dim ermess
            Dim strSwf As String = "INSERT INTO T_ECX_SwiftMT799"
            strSwf = strSwf & " (F_ISN, F_FechaIngreso, F_FechaProceso, F_tipo_mensaje, F_Hora_ingreso, F_Corresponsal, " & _
                        " F_msgReference, F_relReference, F_Explanation, " & _
                        " F_SwiftMessage, F_DataTail, F_OrigFile, " & _
                        " F_SwiftSender, F_SwiftReceiver, F_SwiftMUR, F_SwiftHead, F_SwiftText, F_SwiftTrailer)"
            strSwf = strSwf & " VALUES (@ISN, @FECING, GETDATE(), @TPMSG, @HRING, @CORRESP, " & _
                        " @MSGREF, @RELREF, @EXPLAN, @SWIFTMSG, @DATTAIL, @FILEDAT, " & _
                        " @SWFSNDR, @SWFRCVR, @SWFMUR, @SWFHED, @SWFTXT, @SWFTRAI)"

            Dim swftComm As New SqlCommand(strSwf, conx)
            swftComm.Parameters.Add("@ISN", SqlDbType.Int).Value = CInt(SwiftData.ISN_mensaje)
            swftComm.Parameters.Add("@FECING", SqlDbType.DateTime).Value = SwiftData.Fecha_Ingreso
            swftComm.Parameters.Add("@TPMSG", SqlDbType.Int).Value = CInt(SwiftData.tipo_mensaje)
            swftComm.Parameters.Add("@HRING", SqlDbType.VarChar, 5).Value = Left(SwiftData.Hora_ingreso, 2) & ":" & Mid(SwiftData.Hora_ingreso, 3)
            swftComm.Parameters.Add("@CORRESP", SqlDbType.VarChar, 50).Value = SwiftData.Corresponsal
            swftComm.Parameters.Add("@MSGREF", SqlDbType.VarChar, 16).Value = SwiftData.msgReference
            swftComm.Parameters.Add("@RELREF", SqlDbType.VarChar, 16).Value = SwiftData.relReference
            swftComm.Parameters.Add("@EXPLAN", SqlDbType.VarChar, 1800).Value = SwiftData.msgExplanation
            swftComm.Parameters.Add("@SWIFTMSG", SqlDbType.Text).Value = SwiftData.dat_All
            swftComm.Parameters.Add("@DATTAIL", SqlDbType.Text).Value = SwiftData.dat_Tail
            swftComm.Parameters.Add("@FILEDAT", SqlDbType.Text).Value = SwiftData.OrigFile
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
                SaveDataSwift = True
            Catch ex As Exception
                _c.Informa("  Error: " & ex.ToString)
                If Not Envia("Swifts Recibidos ", ex.ToString) Then
                    _c.Informa(" Mensaje Error " & ex.ToString & " enviado a los respectivos Mails ")
                Else
                    _c.Informa(" No pudo ser enviado el Mensaje")
                End If
                SaveDataSwift = False
            Finally
                'swftComm.Dispose()
            End Try
            If SaveDataSwift Then
                msg_799 = "Señores," & vbCrLf
                msg_799 = msg_799 & "   Se ha recibido mensaje swift " & SwiftData.tipo_mensaje & ", que adjuntamos para su revisión." & vbCrLf
                msg_799 = msg_799 & vbCrLf
                msg_799 = msg_799 & SwiftData.MsgHead & vbCrLf & "-------------------------------------------" & vbCrLf & SwiftData.MsgText
                'msg_799 = msg_799 & SwiftData.dat_All
                If Not Envia("Mensaje Swift " & SwiftData.tipo_mensaje & " Recibido", msg_799) Then
                    _c.Informa(" Mensaje MT" & SwiftData.tipo_mensaje & " enviado a los respectivos Mails ")
                Else
                    _c.Informa(" No pudo ser enviado el Mensaje MT" & SwiftData.tipo_mensaje & " a los respectivos Mails ")
                End If
            End If


        Else
            _c.Informa("   -- Tipo Swift: " & SwiftData.tipo_mensaje & " -> " & SwiftData.Moneda)
            '            SaveDataSwift = False
        End If
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
    Sub rConsola(ByVal tx As String)
        If _verboso Then System.Console.Write(tx) : System.Console.Write(vbCrLf)
    End Sub
    Function toDate(ByVal yyDate As String) As DateTime
        If Len(yyDate) < 8 Then
            toDate = Nothing 'System.DBNull.Value
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
    Function Envia(ByVal subject As String, ByVal mensaje As String) As Boolean
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
