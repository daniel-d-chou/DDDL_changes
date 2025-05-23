Imports ADODB

Public Class DOCTOR_TAB

    Private CNN As Connection

    Private Sub DBConnect()
        Try
            If CNN Is Nothing Then
                CNN = New Connection()
            End If

            If CNN.State = 0 Then
                CNN.ConnectionString = "Provider=SQLOLEDB.1;Data Source=DANIEL\MSSQLSERVER01;user id=sa;password=1234;Initial Catalog=FINALS"
                CNN.Open()
            End If
        Catch ex As Exception
            MsgBox("Database connection error: " & ex.Message)
        End Try
    End Sub

    Private Sub SetupListView()
        Dim x As Integer
        x = ListView1.Width / 5
        ListView1.View = View.Details
        ListView1.Columns.Clear()

        ListView1.Columns.Add("First Name", x)
        ListView1.Columns.Add("Last Name", x)
        ListView1.Columns.Add("Date of Birth", x)
        ListView1.Columns.Add("Sex", x)
        ListView1.Columns.Add("Address", x)
        ListView1.Columns.Add("Cellphone Number", x)
        ListView1.Columns.Add("Email Address", x)
        ListView1.Columns.Add("Marital Status", x)
        ListView1.Columns.Add("Emergency Contact", x)
        ListView1.Columns.Add("Emergency Cell No.", x)
        ListView1.Columns.Add("Registration Date", x)
        ListView1.Columns.Add("Patient ID", x)
        ListView1.Columns.Add("Password", x)
        ListView1.Columns.Add("Date_of_Admission", x)
        ListView1.Columns.Add("Date_of_Discharge", x)
        ListView1.Columns.Add("Patient_Weight", x)
        ListView1.Columns.Add("Patient_Height", x)
        ListView1.Columns.Add("Doctor_in_Charge", x)
        ListView1.Columns.Add("Blood_Type", x)
        ListView1.Columns.Add("Temperature", x)
        ListView1.Columns.Add("Sugar_Level", x)
        ListView1.Columns.Add("Findings", x)
        ListView1.Columns.Add("Treatment", x)
        ListView1.Columns.Add("Follow_up", x)
    End Sub


    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Home.Show()
        If CNN.State = ADODB.ObjectStateEnum.adStateOpen Then
            CNN.Close()
        End If
        Me.Hide()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        DOCTOR_LOG_IN.Show()
        If CNN.State = ADODB.ObjectStateEnum.adStateOpen Then
            CNN.Close()
        End If
        Me.Hide()
    End Sub

    Private Sub DisplayRecords()
        Dim RST As New ADODB.Recordset
        Dim STRSQL As String

        STRSQL = "SELECT * FROM DOCTOR"
        If TextBox8.Text <> "" Then
            STRSQL &= vbCrLf & "WHERE (First_Name LIKE '%" & TextBox8.Text & "%')"
        End If
        RST = CNN.Execute(STRSQL)
        ListView1.Items.Clear()
        While Not RST.EOF
            ListView1.Items.Add(RST.Fields("First_Name").Value)
            ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add(RST.Fields("Last_Name").Value)
            ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add(RST.Fields("Specialization").Value)
            ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add(RST.Fields("Contact_Number").Value)
            ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add(RST.Fields("Username").Value)
            ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add(RST.Fields("Password").Value)
            RST.MoveNext()
        End While
    End Sub

    Private Sub DOCTOR_TAB_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DBConnect()
        SetupListView()
        Call DisplayRecords()
        LoadAllPatients()
    End Sub

    Private Sub LoadAllPatients()
        ListView1.Items.Clear()
        Dim RST As New ADODB.Recordset
        Dim STRSQL As String = "SELECT * FROM PATIENT"
        RST = CNN.Execute(STRSQL)

        Do While Not RST.EOF
            Dim item As New ListViewItem(RST("First_Name").Value.ToString())
            item.SubItems.Add(RST("Last_Name").Value.ToString())
            item.SubItems.Add(RST("Date_of_Birth").Value.ToString())
            item.SubItems.Add(RST("Sex").Value.ToString())
            item.SubItems.Add(RST("Address").Value.ToString())
            item.SubItems.Add(RST("Cellphone_Number").Value.ToString())
            item.SubItems.Add(RST("Email_Address").Value.ToString())
            item.SubItems.Add(RST("Marital_Status").Value.ToString())
            item.SubItems.Add(RST("Emergency_Contact").Value.ToString())
            item.SubItems.Add(RST("Cellphone_No_of_Emergency_Contact").Value.ToString())
            item.SubItems.Add(RST("Registration_Date").Value.ToString())
            item.SubItems.Add(RST("Patient_ID").Value.ToString())
            item.SubItems.Add(RST("Password").Value.ToString())
            item.SubItems.Add(RST("Date_of_Admission").Value.ToString())
            item.SubItems.Add(RST("Date_of_Discharge").Value.ToString())
            item.SubItems.Add(RST("Patient_Weight").Value.ToString())
            item.SubItems.Add(RST("Patient_Height").Value.ToString())
            item.SubItems.Add(RST("Doctor_in_Charge").Value.ToString())
            item.SubItems.Add(RST("Blood_Type").Value.ToString())
            item.SubItems.Add(RST("Temperature").Value.ToString())
            item.SubItems.Add(RST("Sugar_Level").Value.ToString())
            item.SubItems.Add(RST("Findings").Value.ToString())
            item.SubItems.Add(RST("Treatment").Value.ToString())
            item.SubItems.Add(RST("Follow_up").Value.ToString())
            ListView1.Items.Add(item)
            RST.MoveNext()
        Loop
        RST.Close()
    End Sub


    Private Sub Button6_Click(sender As Object, e As EventArgs)
        Dim searchName As String = Trim(TextBox1.Text)
        ListView1.Items.Clear()
        Dim RST As New ADODB.Recordset
        Dim STRSQL As String = "SELECT First_Name, Last_Name FROM PATIENT WHERE First_Name LIKE '%" & Replace(searchName, "'", "''") & "%' OR Last_Name LIKE '%" & Replace(searchName, "'", "''") & "%'"
        RST = CNN.Execute(STRSQL)
        Do While Not RST.EOF
            ListView1.Items.Add(RST.Fields("First_Name").Value & " " & RST.Fields("Last_Name").Value)
            RST.MoveNext()
        Loop
        RST.Close()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If Trim(TextBox1.Text) = "" Then
            LoadAllPatients()
        End If
    End Sub
    Private Sub TextBox10_KeyUp(sender As Object, e As KeyEventArgs) Handles TextBox10.KeyUp
        Call LoadAllPatients()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim RST As New ADODB.Recordset
        Dim STRSQL As String
        Dim isALLOK As Boolean = True
        Dim MSG As String = ""

        If isALLOK Then
            STRSQL = "SELECT * FROM PATIENT WHERE Patient_ID = '" & Replace(Trim(TextBox3.Text), "'", "''") & "'"
            RST = CNN.Execute(STRSQL)

            If RST.EOF Then
                MsgBox("Invalid Credentials / Input Patient's ID")
            Else

                STRSQL = "UPDATE PATIENT SET " &
                 "Date_of_Admission = '" & Replace(Trim(TextBox8.Text), "'", "''") & "', " &
                 "Date_of_Discharge = '" & Replace(Trim(TextBox2.Text), "'", "''") & "', " &
                 "Patient_Weight = '" & Replace(Trim(TextBox5.Text), "'", "''") & "', " &
                 "Patient_Height = '" & Replace(Trim(TextBox4.Text), "'", "''") & "', " &
                 "Doctor_in_Charge = '" & Replace(Trim(TextBox1.Text), "'", "''") & "', " &
                 "Blood_Type = '" & Replace(Trim(TextBox6.Text), "'", "''") & "', " &
                 "Temperature = '" & Replace(Trim(TextBox7.Text), "'", "''") & "'," &
                 "Sugar_Level = '" & Replace(Trim(TextBox9.Text), "'", "''") & "'," &
                 "Findings = '" & Replace(Trim(TextBox12.Text), "'", "''") & "', " &
                 "Treatment = '" & Replace(Trim(TextBox11.Text), "'", "''") & "'," &
                 "Follow_up = '" & Replace(Trim(TextBox13.Text), "'", "''") & "' " &
                 "WHERE Patient_ID = '" & Replace(Trim(TextBox3.Text), "'", "''") & "'"



                CNN.Execute(STRSQL)
                MsgBox("Update Saved")
            End If
        Else
            MsgBox(MSG)
        End If
        Me.Hide()
        Dim newForm As New DOCTOR_TAB()
        newForm.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim RST As New ADODB.Recordset
        Dim STRSQL As String
        Dim isALLOK As Boolean = True
        Dim MSG As String = ""

        If isALLOK Then

            STRSQL = "SELECT * FROM PATIENT WHERE Patient_ID = '" & Replace(Trim(TextBox3.Text), "'", "''")
            RST = CNN.Execute(STRSQL)

            If RST.EOF Then
                MsgBox("Invalid Credentials / Input Patient's ID")
            Else

                If MsgBox("Are you sure you want to delete this account?", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Confirm Delete") = MsgBoxResult.Yes Then
                    STRSQL = "DELETE FROM PATIENT WHERE Patient_ID = '" & Replace(Trim(TextBox3.Text), "'", "''") & "'"
                    CNN.Execute(STRSQL)
MsgBox("Patient Deleted Successfully")


                    Me.Hide()
Dim newForm As New DOCTOR_TAB()
newForm.Show()
End If
End If
Else
MsgBox(MSG)
End If

End Sub
End Class

