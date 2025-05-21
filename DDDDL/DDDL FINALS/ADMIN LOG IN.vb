Public Class ADMIN_LOG_IN
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
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
            MSG = MSG & "Enter Username" & vbCrLf
            isALLOK = False
        End If
        If Trim(TextBox2.Text) = "" Then
            MSG = MSG & "Enter Password" & vbCrLf
            isALLOK = False
        End If

        If isALLOK Then
            STRSQL = "SELECT * FROM ADMIN WHERE Username = '" & Replace(Trim(TextBox1.Text), "'", "''") & "' AND Password='" & Replace(Trim(TextBox2.Text), "'", "''") & "'"

            If CNN.State <> ADODB.ObjectStateEnum.adStateOpen Then
                CNN.Open()
            End If

            RST = CNN.Execute(STRSQL)
            If RST.EOF Then
                MsgBox("Invalid Credentials")
            Else
                MsgBox("Welcome " + RST.Fields("Username").Value)
                ADMIN_TAB.Show()
                Me.Hide()
            End If
        Else
            MsgBox(MSG)
        End If
    End Sub
    Private Sub ADMIN_LOG_IN_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CNN.ConnectionString = "Provider=SQLOLEDB.1;Data Source=DANIEL\MSSQLSERVER01;user id=sa;password=1234;Initial Catalog=FINALS"
        CNN.Open()
        TextBox2.UseSystemPasswordChar = True
    End Sub
    Private Sub ADMIN_LOG_IN_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        'CNN.Close()
    End Sub
End Class