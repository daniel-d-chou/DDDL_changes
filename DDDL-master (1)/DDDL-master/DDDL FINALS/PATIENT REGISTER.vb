Public Class PATIENT_REGISTER

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Home.Show()
        Me.Hide()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim controlsToCheck As TextBox() = {TextBox1, TextBox2, TextBox3, TextBox4, TextBox6, TextBox7, TextBox8, TextBox9, TextBox10, TextBox11, TextBox12, TextBox13}
        For Each tb As TextBox In controlsToCheck
            If String.IsNullOrWhiteSpace(tb.Text) Then
                MsgBox("Please fill in all fields before registering.")
                Return
            End If
        Next

        ' Validate date of birth format
        Dim dob As DateTime
        If Not DateTime.TryParse(TextBox9.Text, dob) Then
            MsgBox("Please enter a valid date of birth (e.g., yyyy-MM-dd).")
            Return
        End If

        ' Increment patient ID for next registration before saving
        Dim newPatientId As Long = CLng(TextBox6.Text)
        TextBox6.Text = (newPatientId + 1).ToString()

        ' Open connection if closed
        If CNN.State = 0 Then
            CNN.Open("Provider=SQLOLEDB.1;Data Source=DANIEL\MSSQLSERVER01;user id=sa;password=1234;Initial Catalog=FINALS")
        End If

        ' Prepare INSERT query with incremented Patient ID
        Dim STRSQL As String
        STRSQL = "INSERT INTO PATIENT (First_name, Last_Name, Date_of_Birth, Sex, Address, Cellphone_Number, Email_Address, Marital_Status, Emergency_Contact, Cellphone_No_of_Emergency_Contact, Registration_Date, Patient_ID, Password) VALUES "
        STRSQL &= "('" & Replace(TextBox2.Text, "'", "''") & "','" & Replace(TextBox8.Text, "'", "''") & "','" & Replace(TextBox9.Text, "'", "''") & "','" & Replace(TextBox10.Text, "'", "''") & "','" & Replace(TextBox13.Text, "'", "''") & "','" & Replace(TextBox11.Text, "'", "''") & "','" & Replace(TextBox12.Text, "'", "''") & "','" & Replace(TextBox1.Text, "'", "''") & "','" & Replace(TextBox3.Text, "'", "''") & "','" & Replace(TextBox4.Text, "'", "''") & "','" & Replace(TextBox5.Text, "'", "''") & "','" & newPatientId.ToString() & "','" & Replace(TextBox7.Text, "'", "''") & "')"

        CNN.Execute(STRSQL)

        MsgBox("Your patient ID is: " & newPatientId.ToString())  ' Show the ID just inserted

        Me.Hide()
        Dim newForm As New PATIENT_LOG_IN()
        newForm.Show()
    End Sub

    Private Sub PATIENT_REGISTER_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox6.ReadOnly = True
        TextBox5.Text = DateTime.Now.ToString("yyyy-MM-dd")  ' Set today's date for registration
        TextBox5.ReadOnly = True


        ' Get the next available Patient ID from the database
        Dim RST As New ADODB.Recordset
        Dim nextId As Long = 1
        If CNN.State = 0 Then
            CNN.Open("Provider=SQLOLEDB.1;Data Source=DANIEL\MSSQLSERVER01;user id=sa;password=1234;Initial Catalog=FINALS")
        End If
        RST.Open("SELECT MAX(Patient_ID) AS MaxID FROM PATIENT", CNN)
        If Not RST.EOF And Not IsDBNull(RST.Fields("MaxID").Value) Then
            nextId = CLng(RST.Fields("MaxID").Value) + 1
        End If
        RST.Close()
        TextBox6.Text = nextId.ToString()
    End Sub
End Class