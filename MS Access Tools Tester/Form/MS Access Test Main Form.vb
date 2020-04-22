Imports Access_Tools

Public Class MS_Access_Test
    Private TestObject As New AccessTools
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        MessageBox.Show(TestObject.CreateAccessDatabase(TextBox1.Text))
    End Sub
End Class
