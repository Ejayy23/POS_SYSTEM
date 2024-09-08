Module connection
    Public message As String = ""
    Public message1 As String
    Public Function mycon() As OleDb.OleDbConnection
        Return New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\shakehauzDB.accdb")
    End Function
End Module
