Module Module1
    Public CNN As New ADODB.Connection

    Public Sub DBConnect()
        If CNN.State = 1 Then ' 1 means the connection is open
            CNN.Close()
        End If
        CNN.ConnectionString = "Provider=SQLOLEDB.1;Data Source=DANIEL\MSSQLSERVER01;user id=sa;password=1234;Initial Catalog=FINALS"
        CNN.Open()
    End Sub
End Module
