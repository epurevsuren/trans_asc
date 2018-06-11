Imports trans_asc.clFunction

Public Class Form1

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        If (ComboBox1.Text = "Asc-Uni") Then
            TextBox2.Text = Asc2Uni(TextBox1.Text)
        Else
            TextBox2.Text = Uni2Asc(TextBox1.Text)
        End If
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ComboBox1.Items.Add("Asc-Uni")
        ComboBox1.Items.Add("Uni-Asc")
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        If (ComboBox1.Text = "Asc-Uni") Then
            TextBox2.Text = Asc2Uni(TextBox1.Text)
        Else
            TextBox2.Text = Uni2Asc(TextBox1.Text)
        End If
    End Sub
End Class
