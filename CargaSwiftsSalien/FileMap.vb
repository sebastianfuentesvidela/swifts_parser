Namespace SwiftsJob

Public Class FileMap
        Public fullName As String
        Public fileDate As Date
        Public Function Name() As String
            Name = fullName.Substring(fullName.LastIndexOf("\") + 1)
        End Function
End Class

End Namespace
