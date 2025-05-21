Public Class Home
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        CONTACT.Show()
        Me.Hide()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ADMIN_LOG_IN.Show()
        Me.Hide()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        DOCTOR_LOG_IN.Show()
        Me.Hide()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        PATIENT_REGISTER.Show()
        Me.Hide()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        PATIENT_LOG_IN.Show()
        Me.Hide()
    End Sub

    Private Sub Home_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    Private Sub Home_Closed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed

    End Sub
End Class
