Imports System.Security.Cryptography
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class ADMIN_TAB
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Home.Show()
        If CNN.State = ADODB.ObjectStateEnum.adStateOpen Then
            CNN.Close()
        End If
        Me.Hide()
    End Sub
    Private Sub AddCols()
        Dim x As Integer
        x = ListView1.Width / 5

        ListView1.Columns.Clear()

        ListView1.Columns.Add("First Name", x)
        ListView1.Columns.Add("Last Name", x)
        ListView1.Columns.Add("Specialization", x)
        ListView1.Columns.Add("Contact Number", x)
        ListView1.Columns.Add("Username", x)
        ListView1.Columns.Add("Password", x)
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

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim STRSQL As String
        Dim MSG As String = ""
        Dim isALLOK As Boolean = True


        If Trim(TextBox4.Text) = "" Then
            MSG &= "Enter Username" & vbCrLf
            isALLOK = False
        End If

        If Trim(TextBox7.Text) = "" Then
            MSG &= "Enter Password" & vbCrLf
            isALLOK = False
        End If

        If isALLOK Then

            Dim checkRST As New ADODB.Recordset
            Dim checkSQL As String = "SELECT * FROM DOCTOR WHERE Username = '" & Replace(Trim(TextBox4.Text), "'", "''") & "'"
            checkRST = CNN.Execute(checkSQL)

            If Not checkRST.EOF Then
                MsgBox("Username already exists.")
                Exit Sub
            End If


            STRSQL = "INSERT INTO DOCTOR (First_Name, Middle_Name, Last_Name, Specialization, Contact_Number, Username, Password) VALUES (" &
             "'" & Replace(Trim(TextBox1.Text), "'", "''") & "', " &
             "'" & Replace(Trim(TextBox3.Text), "'", "''") & "', " &
             "'" & Replace(Trim(TextBox2.Text), "'", "''") & "', " &
             "'" & Replace(Trim(TextBox5.Text), "'", "''") & "', " &
             "'" & Replace(Trim(TextBox6.Text), "'", "''") & "', " &
             "'" & Replace(Trim(TextBox4.Text), "'", "''") & "', " &
             "'" & Replace(Trim(TextBox7.Text), "'", "''") & "')"

            CNN.Execute(STRSQL)
            MsgBox("Successfully Saved")



        Else
            MsgBox(MSG)
        End If
        Me.Hide()
        Dim newForm As New ADMIN_TAB()
        newForm.Show()
    End Sub

    Private Sub ADMIN_TAB_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If CNN.State = ConnectionState.Open Then
            CNN.Close()
        End If

        CNN.ConnectionString = "Provider=SQLOLEDB.1;Data Source=DANIEL\MSSQLSERVER01;user id=sa;password=1234;Initial Catalog=FINALS"        
        CNN.Open()
        ListView1.View = View.Details
        Call AddCols()
        Call DisplayRecords()
    End Sub

    Private Sub ADMIN_TAB_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        'CNN.Close()
    End Sub

    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView1.SelectedIndexChanged

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim RST As New ADODB.Recordset
        Dim STRSQL As String
        Dim isALLOK As Boolean = True
        Dim MSG As String = ""

        If isALLOK Then

            STRSQL = "SELECT * FROM DOCTOR WHERE Username = '" & Replace(Trim(TextBox4.Text), "'", "''") & "' AND Password='" & Replace(Trim(TextBox7.Text), "'", "''") & "'"
            RST = CNN.Execute(STRSQL)

            If RST.EOF Then
                MsgBox("Invalid Credentials / Input Username & Password")
            Else

                STRSQL = "UPDATE DOCTOR SET " &
                 "First_Name = '" & Replace(Trim(TextBox1.Text), "'", "''") & "', " &
                 "Middle_Name = '" & Replace(Trim(TextBox3.Text), "'", "''") & "', " &
                 "Last_Name = '" & Replace(Trim(TextBox2.Text), "'", "''") & "', " &
                 "Specialization = '" & Replace(Trim(TextBox5.Text), "'", "''") & "', " &
                 "Contact_Number = '" & Replace(Trim(TextBox6.Text), "'", "''") & "', " &
                 "Username = '" & Replace(Trim(TextBox4.Text), "'", "''") & "', " &
                 "Password = '" & Replace(Trim(TextBox7.Text), "'", "''") & "' " &
                 "WHERE Username = '" & Replace(Trim(TextBox4.Text), "'", "''") & "'"

                CNN.Execute(STRSQL)
                MsgBox("Update Saved")
            End If
        Else
            MsgBox(MSG)
        End If
        Me.Hide()
        Dim newForm As New ADMIN_TAB()
        newForm.Show()
    End Sub
    Private Sub TextBox8_KeyUp(sender As Object, e As KeyEventArgs) Handles TextBox8.KeyUp
        Call DisplayRecords()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim RST As New ADODB.Recordset
        Dim STRSQL As String
        Dim isALLOK As Boolean = True
        Dim MSG As String = ""

        If isALLOK Then

            STRSQL = "SELECT * FROM DOCTOR WHERE Username = '" & Replace(Trim(TextBox4.Text), "'", "''") & "' AND Password='" & Replace(Trim(TextBox7.Text), "'", "''") & "'"
            RST = CNN.Execute(STRSQL)

            If RST.EOF Then
                MsgBox("Invalid Credentials / Input Username & Password")
            Else

                If MsgBox("Are you sure you want to delete this account?", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Confirm Delete") = MsgBoxResult.Yes Then
                    STRSQL = "DELETE FROM DOCTOR WHERE Username = '" & Replace(Trim(TextBox4.Text), "'", "''") & "'"
                    CNN.Execute(STRSQL)
                    MsgBox("Account Deleted Successfully")


                    Me.Hide()
                    Dim newForm As New ADMIN_TAB()
                    newForm.Show()
                End If
            End If
        Else
            MsgBox(MSG)
        End If

    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged

    End Sub
End Class
