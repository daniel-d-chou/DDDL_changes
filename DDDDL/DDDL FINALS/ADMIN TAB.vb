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
        x = ListView1.Width / 4

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
            STRSQL &= vbCrLf & "WHERE (Username LIKE '%" & TextBox8.Text & "%')"
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ListView1.SelectedItems.Count = 0 Then
            MsgBox("Please select a record to delete.")
            Exit Sub
        End If

        Dim selectedUsername As String = ListView1.SelectedItems(0).SubItems(3).Text.Trim()
        Dim STRSQL As String = "DELETE FROM DOCTOR WHERE LTRIM(RTRIM(Username)) = ?"

        Using cmd As New OleDb.OleDbCommand(STRSQL, CNN)
            cmd.Parameters.AddWithValue("?", selectedUsername)
            cmd.ExecuteNonQuery()
        End Using

        MsgBox("Record deleted successfully.")
        Call DisplayRecords()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If ListView1.SelectedItems.Count = 0 Then
            MsgBox("Please select a record to update.")
            Exit Sub
        End If

        Dim selectedUsername As String = ListView1.SelectedItems(0).SubItems(3).Text.Trim()

        Dim STRSQL As String = "UPDATE DOCTOR SET First_Name = ?, Middle_Name = ?, Last_Name = ?, Specialization = ?, Contact_Number = ?, Password = ? WHERE LTRIM(RTRIM(Username)) = ?"

        Using cmd As New OleDb.OleDbCommand(STRSQL, CNN)
            cmd.Parameters.AddWithValue("?", TextBox1.Text.Trim())
            cmd.Parameters.AddWithValue("?", TextBox3.Text.Trim())
            cmd.Parameters.AddWithValue("?", TextBox2.Text.Trim())
            cmd.Parameters.AddWithValue("?", TextBox5.Text.Trim())
            cmd.Parameters.AddWithValue("?", TextBox6.Text.Trim())
            cmd.Parameters.AddWithValue("?", TextBox7.Text.Trim())
            cmd.Parameters.AddWithValue("?", selectedUsername)

            cmd.ExecuteNonQuery()
        End Using

        MsgBox("Record updated successfully.")
        Call DisplayRecords()
    End Sub


    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim STRSQL As String = "INSERT INTO DOCTOR (First_Name, Middle_Name, Last_Name, Specialization, Contact_Number, Username, Password) " &
                           "VALUES (?, ?, ?, ?, ?, ?, ?)"

        Using cmd As New OleDb.OleDbCommand(STRSQL, CNN)
            cmd.Parameters.AddWithValue("?", TextBox1.Text.Trim())
            cmd.Parameters.AddWithValue("?", TextBox3.Text.Trim())
            cmd.Parameters.AddWithValue("?", TextBox2.Text.Trim())
            cmd.Parameters.AddWithValue("?", TextBox5.Text.Trim())
            cmd.Parameters.AddWithValue("?", TextBox6.Text.Trim())
            cmd.Parameters.AddWithValue("?", TextBox4.Text.Trim())
            cmd.Parameters.AddWithValue("?", TextBox7.Text.Trim())

            cmd.ExecuteNonQuery()
        End Using

        MsgBox("Successfully Saved")
        Call DisplayRecords()
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

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub
End Class
