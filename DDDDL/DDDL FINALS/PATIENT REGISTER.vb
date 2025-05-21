Imports System.Data.OleDb
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class PATIENT_REGISTER

    'Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
    '    Dim STRSQL As String
    '    If CNN.State = 0 Then
    '        CNN.Open("Provider=SQLOLEDB.1;Data Source=DANIEL\MSSQLSERVER01;user id=sa;password=1234;Initial Catalog=FINALS")
    '    End If

    '    STRSQL = "INSERT INTO PATIENT (First_name, Last_Name, Date_of_Birth, Sex, Address, Cellphone_Number, Email_Address, Marital_Status, Emergency_Contact, 
    '    Cellphone_No_of_Emergency_Contact, Registration_Date, Patient_ID, Password) VALUES "
    '    STRSQL = STRSQL & vbCrLf & "('" & TextBox2.Text & "','" & TextBox8.Text & "','" & TextBox9.Text & "','" & TextBox10.Text & "','" & TextBox13.Text & "','" & TextBox11.Text & "','" & TextBox12.Text & "','" & TextBox1.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "','" & TextBox6.Text & "','" & TextBox7.Text & "')"
    '    CNN.Execute(STRSQL)
    '    MsgBox("Successfully Saved")
    '    Me.Close()
    '    Home.Show()
    '    Me.Hide()
    'End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' Check if any textbox is empty
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or
       TextBox5.Text = "" Or TextBox6.Text = "" Or TextBox7.Text = "" Or TextBox8.Text = "" Or
       TextBox9.Text = "" Or TextBox10.Text = "" Or TextBox11.Text = "" Or TextBox12.Text = "" Or
       TextBox13.Text = "" Then
            MsgBox("Please fill in all fields before submitting.")
            Exit Sub
        End If

        Dim STRSQL As String
        If CNN.State = 0 Then
            CNN.Open("Provider=SQLOLEDB.1;Data Source=DANIEL\MSSQLSERVER01;user id=sa;password=1234;Initial Catalog=FINALS")
        End If

        STRSQL = "INSERT INTO PATIENT (First_name, Last_Name, Date_of_Birth, Sex, Address, Cellphone_Number, Email_Address, Marital_Status, Emergency_Contact, 
    Cellphone_No_of_Emergency_Contact, Registration_Date, Patient_ID, Password) VALUES "
        STRSQL &= vbCrLf & "('" & TextBox2.Text & "','" & TextBox8.Text & "','" & TextBox9.Text & "','" & TextBox10.Text & "','" & TextBox13.Text & "','" & TextBox11.Text & "','" & TextBox12.Text & "','" & TextBox1.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "','" & TextBox6.Text & "','" & TextBox7.Text & "')"

        CNN.Execute(STRSQL)
        MsgBox("Successfully Saved")

        Me.Close()
        Home.Show()
        Me.Hide()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Home.Show()
        Me.Hide()
    End Sub

End Class
