Imports ADODB

Public Class DOCTOR_TAB

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
    End Sub

    Private Sub SetupListView()
        ListView1.View = View.Details
        ListView1.Columns.Clear()

        ListView1.Columns.Add("First Name", 100)
        ListView1.Columns.Add("Last Name", 100)
        ListView1.Columns.Add("Date of Birth", 100)
        ListView1.Columns.Add("Sex", 70)
        ListView1.Columns.Add("Address", 150)
        ListView1.Columns.Add("Cellphone Number", 120)
        ListView1.Columns.Add("Email Address", 150)
        ListView1.Columns.Add("Marital Status", 100)
        ListView1.Columns.Add("Emergency Contact", 120)
        ListView1.Columns.Add("Emergency Cell No.", 120)
        ListView1.Columns.Add("Registration Date", 100)
        ListView1.Columns.Add("Patient ID", 80)
        ListView1.Columns.Add("Password", 100)
    End Sub


    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Home.Show()
        Me.Hide()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        DOCTOR_LOG_IN.Show()
        Me.Hide()
    End Sub

    Private Sub DOCTOR_TAB_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DBConnect()
        SetupListView()
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
            ListView1.Items.Add(item)
            RST.MoveNext()
        Loop
        RST.Close()
    End Sub


    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
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

End Class