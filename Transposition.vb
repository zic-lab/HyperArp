Public Class Transposition
    Public Bout_Val As String
    Public Transp_Val As Integer
    Public TraiterPianoR As Boolean
    Private Sub Transposer_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Module1.LangueIHM = "fr" Then
            Button2.Text = "Annuler"
            CheckBox1.Text = "Transposer PianRolls"
            Me.Text = "Transposer"
        Else
            Button2.Text = "Cancel"
            CheckBox1.Text = "Transpose PianRolls"
            Me.Text = "Transpose"
        End If
        '
        ComboBox1.Items.Clear()
        '
        ComboBox1.Items.Add("+11             VII   Sensible")
        ComboBox1.Items.Add("+10")
        ComboBox1.Items.Add("+9               VI    Sus Dominante")
        ComboBox1.Items.Add("+8")
        ComboBox1.Items.Add("+7               V     Dominante")
        ComboBox1.Items.Add("+6")
        ComboBox1.Items.Add("+5               IV    Sous Dominante")
        ComboBox1.Items.Add("+4               III   Médiante")
        ComboBox1.Items.Add("+3")
        ComboBox1.Items.Add("+2               II    Sus Tonique")
        ComboBox1.Items.Add("+1")
        ComboBox1.Items.Add("+0")
        ComboBox1.Items.Add("-1               VII   Sensible")
        ComboBox1.Items.Add("-2")
        ComboBox1.Items.Add("-3               VI    Sus Dominante")
        ComboBox1.Items.Add("-4")
        ComboBox1.Items.Add("-5               V     Dominante")
        ComboBox1.Items.Add("-6")
        ComboBox1.Items.Add("-7               IV    Sous Dominante")
        ComboBox1.Items.Add("-8               III   Médiante")
        ComboBox1.Items.Add("-9")
        ComboBox1.Items.Add("-10             II   Sus Tonique")
        ComboBox1.Items.Add("-11")
        '
        ComboBox1.SelectedIndex = 11
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim tbl() As String
        tbl = ComboBox1.Text.Split
        Transp_Val = Convert.ToInt16(tbl(0))
        TraiterPianoR = CheckBox1.Checked
        Bout_Val = "OK"
        Me.Hide()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Bout_Val = "Annuler"
        Transp_Val = -1
        Me.Hide()
    End Sub
End Class