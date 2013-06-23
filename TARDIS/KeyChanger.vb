Public Class KeyChanger

    Public EditingKey As String
    Public EditingKeyName As String
    Public label1text As String
    Public label2text As String
    Public label2_1text As String
    Dim SelectedKey As String

    Sub Addkeylist()
        ComboBox1.Items.Add(Keys.Enter)
        ComboBox1.Items.Add(Keys.Escape)
        ComboBox1.Items.Add(Keys.Space)
        ComboBox1.Items.Add(Keys.D1)
        ComboBox1.Items.Add(Keys.D2)
        ComboBox1.Items.Add(Keys.D3)
        ComboBox1.Items.Add(Keys.D4)
        ComboBox1.Items.Add(Keys.D5)
        ComboBox1.Items.Add(Keys.D6)
        ComboBox1.Items.Add(Keys.D7)
        ComboBox1.Items.Add(Keys.D8)
        ComboBox1.Items.Add(Keys.D9)
        ComboBox1.Items.Add(Keys.D0)
        ComboBox1.Items.Add(Keys.A)
        ComboBox1.Items.Add(Keys.B)
        ComboBox1.Items.Add(Keys.C)
        ComboBox1.Items.Add(Keys.D)
        ComboBox1.Items.Add(Keys.E)
        ComboBox1.Items.Add(Keys.F)
        ComboBox1.Items.Add(Keys.G)
        ComboBox1.Items.Add(Keys.H)
        ComboBox1.Items.Add(Keys.I)
        ComboBox1.Items.Add(Keys.J)
        ComboBox1.Items.Add(Keys.K)
        ComboBox1.Items.Add(Keys.L)
        ComboBox1.Items.Add(Keys.M)
        ComboBox1.Items.Add(Keys.N)
        ComboBox1.Items.Add(Keys.O)
        ComboBox1.Items.Add(Keys.P)
        ComboBox1.Items.Add(Keys.Q)
        ComboBox1.Items.Add(Keys.R)
        ComboBox1.Items.Add(Keys.S)
        ComboBox1.Items.Add(Keys.T)
        ComboBox1.Items.Add(Keys.U)
        ComboBox1.Items.Add(Keys.V)
        ComboBox1.Items.Add(Keys.W)
        ComboBox1.Items.Add(Keys.X)
        ComboBox1.Items.Add(Keys.Y)
        ComboBox1.Items.Add(Keys.Z)
        ComboBox1.Items.Add(Keys.F1)
        ComboBox1.Items.Add(Keys.F2)
        ComboBox1.Items.Add(Keys.F3)
        ComboBox1.Items.Add(Keys.F4)
        ComboBox1.Items.Add(Keys.F5)
        ComboBox1.Items.Add(Keys.F6)
        ComboBox1.Items.Add(Keys.F7)
        ComboBox1.Items.Add(Keys.F8)
        ComboBox1.Items.Add(Keys.F9)
        ComboBox1.Items.Add(Keys.F10)
        ComboBox1.Items.Add(Keys.F11)
        ComboBox1.Items.Add(Keys.F12)
        ComboBox1.Items.Add(Keys.F13)
        ComboBox1.Items.Add(Keys.F14)
        ComboBox1.Items.Add(Keys.F15)
        ComboBox1.Items.Add(Keys.F16)
        ComboBox1.Items.Add(Keys.F18)
        ComboBox1.Items.Add(Keys.F19)
        ComboBox1.Items.Add(Keys.F20)
        ComboBox1.Items.Add(Keys.F21)
        ComboBox1.Items.Add(Keys.F22)
        ComboBox1.Items.Add(Keys.F23)
        ComboBox1.Items.Add(Keys.F24)
    End Sub

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
        Addkeylist()
        ComboBox1.SelectedIndex = "0"
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

    End Sub

    Private Sub ComboBox1_DropDownClosed(sender As Object, e As EventArgs) Handles ComboBox1.DropDownClosed
        Label2.Text = label2text & EditingKeyName & label2_1text & ComboBox1.SelectedItem
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        EditingKey = ""
        EditingKeyName = ""
        SelectedKey = ""
        Label2.Text = ""
        Me.Close()
    End Sub
End Class