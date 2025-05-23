Public Class PATIENT_TAB
    ' Property to receive Patient ID from login form
    Public Property PatientID As String

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Home.Show()
        If CNN.State = ADODB.ObjectStateEnum.adStateOpen Then
            CNN.Close()
        End If
        Me.Hide()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        PATIENT_LOG_IN.Show()
        If CNN.State = ADODB.ObjectStateEnum.adStateOpen Then
            CNN.Close()
        End If
        Me.Hide()
    End Sub

    Private Sub PATIENT_TAB_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim STRSQL As String
        Dim RST As New ADODB.Recordset

        TextBox1.ReadOnly = True
        TextBox2.ReadOnly = True
        TextBox3.ReadOnly = True
        TextBox4.ReadOnly = True
        TextBox5.ReadOnly = True
        TextBox6.ReadOnly = True
        TextBox7.ReadOnly = True
        TextBox8.ReadOnly = True
        TextBox9.ReadOnly = True
        TextBox10.ReadOnly = True
        TextBox11.ReadOnly = True
        TextBox12.ReadOnly = True
        ' Open the connection if not already open
        If CNN.State = 0 Then
            CNN.Open("Provider=SQLOLEDB.1;Data Source=DANIEL\MSSQLSERVER01;user id=sa;password=1234;Initial Catalog=FINALS")
        End If

        ' Retrieve logged-in patient's data
        STRSQL = "SELECT * FROM PATIENT WHERE Patient_ID = '" & Replace(PatientID, "'", "''") & "'"
        RST.Open(STRSQL, CNN)

        If Not RST.EOF Then
            TextBox1.Text = RST.Fields("Registration_Date").Value                ' Date of Admission
            TextBox2.Text = RST.Fields("First_Name").Value & " " & RST.Fields("Last_Name").Value   ' Patient Name
            TextBox7.Text = RST.Fields("Date_of_Discharge").Value                ' Date of Discharge
            TextBox3.Text = RST.Fields("Patient_Height").Value                           ' Patient's Height
            TextBox8.Text = RST.Fields("Temperature").Value                      ' Temperature
            TextBox4.Text = RST.Fields("Patient_Weight").Value                           ' Patient's Weight
            TextBox9.Text = RST.Fields("Blood_Type").Value                       ' Blood Type
            TextBox5.Text = RST.Fields("Sugar_Level").Value                      ' Sugar Level
            TextBox1.Text = RST.Fields("Date_of_Admission").Value                   ' Follow Up Check Up
            TextBox6.Text = RST.Fields("Doctor_In_Charge").Value                ' Doctor's in Charge
            TextBox10.Text = RST.Fields("Follow_up").Value
            TextBox11.Text = RST.Fields("Treatment").Value
            TextBox12.Text = RST.Fields("Findings").Value
            ' You do not have controls for Treatment and Findings in the Designer file.
        End If

        RST.Close()
    End Sub
End Class
