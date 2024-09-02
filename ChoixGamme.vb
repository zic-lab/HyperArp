Public Class ChoixGamme
    Public retour As String = ""
    Private Sub ChoixGamme_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Lagnue
        If LangueIHM = "fr" Then
            Me.Text = "Choix des gammes"
            Button1.Text = "OK"
            Button2.Text = "Annuler"
            Button3.Text = "Filtre de Gammes"
        Else
            Me.Text = "Choice of scales"
            Button1.Text = "OK"
            Button2.Text = "Cancel"
            Button3.Text = "Scales Filter"
        End If
        '
        Button1.Visible = True
        ' Position du formulaire
        Dim P As New Point()
        P = Form1.Location
        P.X = P.X + 10 '+ 45
        P.Y = P.Y + 538
        '
        Me.Location = P
        '
        ComboBox1.Items.Clear()
        If Trim(Form1.ListGammesJouables) <> "" Then
            Dim TBL() As String = Form1.ListGammesJouables.Split("-")
            '
            For Each a As String In TBL
                ComboBox1.Items.Add(a)
            Next
        Else
            Button1.Visible = False
            ComboBox1.Items.Add(" ")
            If LangueIHM = "fr" Then
                Me.Text = ("Gammes(s) Non Trouvées")
                Label1.Text = ("Gammes(s) Non Trouvées")
                Label2.Text = ("Gammes(s) Non Trouvées")
            Else
                Me.Text = ("Scale(s) not Found")
                Label1.Text = ("Scale(s) not Found")
                Label2.Text = ("Scale(s) not Found")
            End If
        End If
        ComboBox1.SelectedIndex = 0
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        retour = ""
        Me.Hide()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Trim(ComboBox1.Text) <> "" Then
            retour = "OK"
        Else
            retour = ""
        End If
        Me.Hide()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim tbl() As String
        FiltreGammes.ShowDialog()
        If FiltreGammes.Retour <> "" Then
            ComboBox1.Items.Clear()
            tbl = FiltreGammes.Retour.Split()
            For Each a As String In tbl
                ComboBox1.Items.Add(a)
            Next
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim aa As String = Trim(ComboBox1.Text)
        Dim a As String
        Dim gamme As String
        Dim tbl() As String
        Dim oo2 As New RechercheG_v2

        If Trim(Form1.ListGammesJouables) <> "" Then
            tbl = aa.Split()
            gamme = Trim(tbl(1))

            If Trim(gamme) <> "" Then
                Label1.Text = oo2.Det_NotesGammes3(Trim(ComboBox1.Text))
                a = oo2.Det_InfoGamme(Trim(gamme))
                a = a.Replace(";", " ")
                a = a.Replace("Fin", " ")
                a = a.Replace(" 1", "   1")
                Label2.Text = a
            End If
        End If
    End Sub
End Class