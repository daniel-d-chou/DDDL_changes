Public Class DOCTOR_LOG_IN

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Home.Show()
        If CNN.State = ADODB.ObjectStateEnum.adStateOpen Then
            CNN.Close()
        End If
        Me.Hide()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim RST As New ADODB.Recordset
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
                If CNN.State = 0 Then
                    CNN.Open()
                End If
                STRSQL = "SELECT * FROM DOCTOR WHERE Username = '" & Replace(Trim(TextBox1.Text), "'", "''") & "' AND Password='" & Replace(Trim(TextBox2.Text), "'", "''") & "'"
                RST.Open(STRSQL, CNN)
                If RST.EOF Then
                    MsgBox("Invalid Credentials")
                Else
                    MsgBox("Welcome Dr. " & RST.Fields("Last_Name").Value)
                    DOCTOR_TAB.Show()
                    Me.Hide()
                End If
                RST.Close()
            Catch ex As Exception
                MsgBox("An error occurred: " & ex.Message)
            End Try
        Else
            MsgBox(MSG)
        End If
    End Sub

    Private Sub DOCTOR_LOG_IN_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            If CNN.State = 0 Then
                CNN.ConnectionString = "Provider=SQLOLEDB.1;Data Source=DANIEL\MSSQLSERVER01;user id=sa;password=1234;Initial Catalog=FINALS"
                CNN.Open()
            End If
            TextBox2.UseSystemPasswordChar = True
        Catch ex As Exception
            MsgBox("Database connection failed: " & ex.Message)
        End Try
    End Sub

    Private Sub DOCTOR_LOG_IN_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        If CNN.State = ADODB.ObjectStateEnum.adStateOpen Then
            CNN.Close()
        End If
    End Sub

End Class
