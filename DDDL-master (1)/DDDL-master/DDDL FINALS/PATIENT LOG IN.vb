Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class PATIENT_LOG_IN
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
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
            STRSQL = "SELECT * FROM PATIENT WHERE Patient_ID = '" & Replace(Trim(TextBox1.Text), "'", "''") & "' AND Password='" & Replace(Trim(TextBox2.Text), "'", "''") & "'"
            RST = CNN.Execute(STRSQL)
            If RST.EOF Then
                MsgBox("Invalid Credentials")
            Else
                MsgBox("Welcome " + RST.Fields("First_Name").Value)
                Dim tab As New PATIENT_TAB()
                tab.PatientID = RST.Fields("Patient_ID").Value.ToString()
                tab.Show()
                Me.Hide()
            End If

            RST.Close()
        Else
            MsgBox(MSG)
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Home.Show()
        Me.Hide()
    End Sub

    Private Sub PATIENT_LOG_IN_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Module1.DBConnect()
        TextBox2.UseSystemPasswordChar = True
    End Sub
End Class