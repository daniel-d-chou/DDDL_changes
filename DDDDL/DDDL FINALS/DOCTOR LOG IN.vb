Imports ADODB

Public Class DOCTOR_LOG_IN
    Private CNN As Connection ' Declare the connection object

    Private Sub DBConnect()
        Try
            If CNN Is Nothing Then
                CNN = New Connection() ' Initialize if not already done
            End If

            If CNN.State = 0 Then ' if connection is closed
                CNN.ConnectionString = "Provider=SQLOLEDB.1;Data Source=DANIEL\MSSQLSERVER01;user id=sa;password=1234;Initial Catalog=FINALS"
                CNN.Open()
            End If
        Catch ex As Exception
            MsgBox("Database connection error: " & ex.Message)
        End Try

        If CNN Is Nothing OrElse CNN.State = 0 Then
            Call DBConnect()
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Home.Show()
        Me.Hide()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim RST As Recordset = Nothing
        Dim STRSQL As String
        Dim MSG As String = ""
        Dim isALLOK As Boolean = True

        If Trim(TextBox1.Text) = "" Then
            MSG &= "Enter Username" & vbCrLf
            isALLOK = False
        End If
        If Trim(TextBox2.Text) = "" Then
            MSG &= "Enter Password" & vbCrLf
            isALLOK = False
        End If

        If isALLOK Then
            Try
                Call DBConnect()
                If CNN Is Nothing OrElse CNN.State = 0 Then
                    Throw New Exception("Database connection is not open.")
                End If

                STRSQL = "SELECT * FROM DOCTOR WHERE Username = '" & Replace(Trim(TextBox1.Text), "'", "''") & "' AND Password='" & Replace(Trim(TextBox2.Text), "'", "''") & "'"
                RST = CNN.Execute(STRSQL)

                If RST.EOF Then
                    MsgBox("Invalid Credentials")
                Else
                    MsgBox("Welcome Dr. " & RST.Fields("First_Name").Value)
                    DOCTOR_TAB.Show()
                    Me.Hide()
                End If
            Catch ex As Exception
                MsgBox("Error: " & ex.Message)
            Finally
                If RST IsNot Nothing Then RST.Close()
            End Try
        Else
            MsgBox(MSG)
        End If
    End Sub

    Private Sub DOCTOR_LOG_IN_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox2.UseSystemPasswordChar = True
    End Sub

    Private Sub DOCTOR_LOG_IN_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Try
            If CNN IsNot Nothing AndAlso CNN.State = 1 Then ' If connection is open
                CNN.Close()
            End If
        Catch ex As Exception
            ' Optional: log or ignore
        End Try
    End Sub
End Class