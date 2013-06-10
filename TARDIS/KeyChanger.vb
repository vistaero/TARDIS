Public Class KeyChanger

    Public EditingKey As String
    Public EditingKeyName As String
    Public label1text As String
    Public label2text As String
    Public label2_1text As String
    Dim SelectedKey As String

    Private Sub KeyChanger_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Screen1.LanguageComboBox.SelectedIndex = 1 Then
            label1text = "Editando "
            label2text = "La tecla para "
            label2_1text = " será "
        End If
        If Screen1.LanguageComboBox.SelectedIndex = 0 Then
            label1text = "Editing "
            label2text = ""
        End If
        Label1.Text = label1text + EditingKeyName
        ComboBox1.Items.Add(Keys.A)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub
End Class