Imports System.IO
Namespace SwiftsJob
    Public Class AppLog
        Private _ventana As Boolean = False
        Private _archivo As Boolean = False
        Private _filedestino As String

        Sub New(Optional ByVal opcionOutPut As Integer = 0, Optional ByVal FileDestino As String = "")
            'If opcionOutPut = 0 Then
            '    _ventana = True
            'ElseIf opcionOutPut = 1 Then
                _archivo = True
                _filedestino = FileDestino
            'ElseIf opcionOutPut = 2 Then
            '    _ventana = True
            '    _archivo = True
            '    _filedestino = FileDestino
            'End If

        End Sub
        Public Sub Informa(ByVal mensaje As String)
            If _ventana Then
                System.Console.WriteLine(mensaje)
            End If
            If _archivo Then
                Dim fs As FileStream
                Try
                    fs = New FileStream(_filedestino, FileMode.Append, FileAccess.Write, FileShare.Read)
                    Dim SwFromFileStream As StreamWriter = New StreamWriter(fs)
                    'SwFromFileStream.WriteLine(Hour(System.DateTime.Now) & ":" & Minute(System.DateTime.Now) & ":" & Second(System.DateTime.Now) & " " & mensaje)
                    SwFromFileStream.WriteLine(Format(System.DateTime.Now, "HH:mm:ss tt") & " " & mensaje)
                    SwFromFileStream.Flush()
                    SwFromFileStream.Close()
                Catch ex As Exception
                    ' System.Console.WriteLine(ex.Message)
                Finally

                End Try
                Try
                    fs.Close()
                Catch ex As Exception

                End Try
                'Dim filepath As String = Path.GetDirectoryName(Server.MapPath(Request.ServerVariables("PATH_INFO")))
                'Dim fullpath As String
                'Dim Ok As Boolean = False
                'Dim bt As Byte, arBt As Byte(), filename As String
            End If
        End Sub

    End Class
End Namespace
