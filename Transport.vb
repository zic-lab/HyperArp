Public Class Transport
    Dim Ouvert As Boolean = True
    ''' <summary>
    ''' Bouton Play
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>                                                                                                          
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim a As String = Form1.Récup_Acc
        If Module1.JeuxActif = False Then
            Form1.PlayHyperArp()
            Button1.Enabled = False
        End If
    End Sub


    ''' <summary>
    ''' Bouton Stop
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Form1.StopPlay()
        Button1.Enabled = True
        Module1.JeuxActif = False

    End Sub

    Private Sub Transport_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        Me.Opacity = 1
        'Panel1.Visible = True
        'Label1.Visible = True
    End Sub
    Private Sub Transport_Deactivate(sender As Object, e As EventArgs) Handles Me.Deactivate
        If Ouvert Then
            Me.Opacity = 1 '0.6 le 21-05-20
            'Panel1.Visible = False
            'Label1.Visible = False
        End If
    End Sub

    Private Sub Transport_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Ouvert = False
        'Me.Visible = False
        e.Cancel = True ' permet d'annuler le déchargement de la form qui est remplacé après par un Hide. La forme reste chargée mais n'est plus visible : elle garde les valeurs de ses composants.
        Me.Hide()
    End Sub

    Private Sub Transport_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim s As New Size(437, 90)
        Me.Size = s
        Me.TopMost = True
        Me.FormBorderStyle = FormBorderStyle.FixedToolWindow
        Label1.Text = "--"
        Label1.ForeColor = Color.Teal ' Color.FromArgb(255, 213, 91)
        If Module1.LangueIHM = "fr" Then
            Me.Text = "BARRE DE TRANSPORT HYPERARP"
        Else
            Me.Text = "HYPERARP TRANPORT BAR"
        End If
        '
        Label2.Location = TextBox1.Location
        Label3.Location = TextBox2.Location
        '
        Label2.Size = TextBox1.Size
        Label3.Size = TextBox2.Size
        '
        'Label2.Font = Label1.Font
        'Label3.Font = Label1.Font
        '
        Label2.TextAlign = Label1.TextAlign
        Label3.TextAlign = Label1.TextAlign
        '
        Label2.BackColor = Color.Teal 'Label1.BackColor
        Label3.BackColor = Color.Teal 'Label1.BackColor
        '
        Label2.ForeColor = Color.Yellow 'Label1.ForeColor
        Label3.ForeColor = Color.Yellow 'Label1.ForeColor
        '
        Label2.Text = Form1.Début.Value.ToString
        Label3.Text = Form1.Terme.Value.ToString
        '
        TextBox1.Text = Form1.Début.Value.ToString
        TextBox2.Text = Form1.Terme.Value.ToString
        '

        If Module1.LangueIHM = "fr" Then
            Label4.Text = "Début"
            Label5.Text = "Fin"

            CheckBox1.Text = "Raccourcis"
        Else
            Label4.Text = "Start"
            Label5.Text = "End"
            CheckBox1.Text = "Shorcuts"
        End If
    End Sub
    Private Sub SortirToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SortirToolStripMenuItem.Click
        Me.Visible = False
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        Dim a As Char = e.KeyChar
        Dim b As String = TextBox1.Text
        Dim i As Integer


        If a = vbCr Then
            If IsNumeric(b) Then
                i = Convert.ToInt16(b)
                TextBox1.BackColor = Color.Teal
                TextBox1.ForeColor = Color.Yellow
                TextBox1.Text = Form1.Début.Value.ToString
                Label2.Text = TextBox1.Text
                If i > 0 Then
                    If Val(b) <= Val(Form1.Terme.Value) Then
                        Label2.Text = Trim(b)
                        Form1.Début.Value = Trim(b)
                    End If
                End If
            End If
            '
            Label2.Visible = True
        End If
    End Sub
    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        Dim a As Char = e.KeyChar
        Dim b As String = TextBox2.Text

        If a = vbCr Then
            If IsNumeric(b) Then
                If Val(b) >= Val(Form1.Début.Value) And Val(b) <= Module1.nbMesures Then
                    TextBox2.BackColor = Color.Teal
                    TextBox2.ForeColor = Color.Yellow
                    '
                    Form1.Terme.Value = Convert.ToInt16(b)
                    Label3.Text = b
                Else
                    TextBox2.BackColor = Color.Teal
                    TextBox2.ForeColor = Color.Yellow
                    TextBox2.Text = Form1.Terme.Value.ToString
                    Label3.Text = TextBox1.Text
                End If
            Else
                TextBox2.BackColor = Color.Teal
                TextBox2.ForeColor = Color.Yellow
                TextBox2.Text = Form1.Terme.Value.ToString
                Label3.Text = TextBox1.Text
            End If
            '
            Label3.Visible = True
        End If
    End Sub

    Private Sub TextBox1_MouseDown(sender As Object, e As MouseEventArgs) Handles TextBox1.MouseDown
        'TextBox1.BackColor = Color.White
        'TextBox1.ForeColor = Color.Black
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub Label2_MouseDown(sender As Object, e As MouseEventArgs) Handles Label2.MouseDown
        Label2.Visible = False
        TextBox1.Focus()
        TextBox1.SelectionStart = Len(TextBox1.Text)
    End Sub

    Private Sub Label3_MouseDown(sender As Object, e As MouseEventArgs) Handles Label3.MouseDown
        Label3.Visible = False
        TextBox2.Focus()
        TextBox2.SelectionStart = Len(TextBox2.Text)
    End Sub

    Private Sub Transport_MouseEnter(sender As Object, e As EventArgs) Handles Me.MouseEnter

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Me.KeyPreview = Not Me.KeyPreview

    End Sub

    Private Sub Transport_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed

    End Sub

    Private Sub Transport_Closed(sender As Object, e As EventArgs) Handles Me.Closed

    End Sub

    Private Sub Transport_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp
        ' PLAY, RECALCUL : F5
        ' *******************
        If e.KeyCode = Keys.F5 Then
            If Not Form1.Horloge1.IsRunning Then
                Form1.PlayHyperArp()
                Button1.Enabled = False
            Else
                Form1.ReCalcul()
            End If
        End If
        '
        '
        ' STOP : F4
        ' *********
        If e.KeyCode = Keys.F4 Then
            Form1.StopPlay()
            Button1.Enabled = True
        End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub
End Class