Public Class ActionsPianoRoll
    Public Enum TypeAct
        Rien = 0
        Effacer
        Coller
        CollerVers
        Transposer
        DéplacerVers
        MoidifierVel
    End Enum
    Public TypeAction As TypeAct
    Public Début As Integer
    Public Fin As Integer
    Public Destination As Integer
    Public N_PianoR As Integer
    Public ValeurTransp As Integer
    Public N_Can1er As Integer
    Public N_Canal As Integer
    Public ValMin As Integer ' vélocité min
    Public ValMax As Integer ' vélocité max
    Private Sub ActionsPianoRoll_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '
        ComboBox4.Items.Clear()
        ComboBox4.Items.Add("+12")
        ComboBox4.Items.Add("+11")
        ComboBox4.Items.Add("+10")
        ComboBox4.Items.Add("+9")
        ComboBox4.Items.Add("+8")
        ComboBox4.Items.Add("+7")
        ComboBox4.Items.Add("+6")
        ComboBox4.Items.Add("+5")
        ComboBox4.Items.Add("+4")
        ComboBox4.Items.Add("+3")
        ComboBox4.Items.Add("+2")
        ComboBox4.Items.Add("+1")
        ComboBox4.Items.Add("+0")
        ComboBox4.Items.Add("-1")
        ComboBox4.Items.Add("-2")
        ComboBox4.Items.Add("-3")
        ComboBox4.Items.Add("-4")
        ComboBox4.Items.Add("-5")
        ComboBox4.Items.Add("-6")
        ComboBox4.Items.Add("-7")
        ComboBox4.Items.Add("-8")
        ComboBox4.Items.Add("-9")
        ComboBox4.Items.Add("-10")
        ComboBox4.Items.Add("-11")
        ComboBox4.Items.Add("-12")
        '
        ComboBox4.SelectedIndex = 12
        '

        If Module1.LangueIHM = "fr" Then
            Me.Text = "Actions dans une sélection de mesures"
            Label1.Text = "Début"
            Label2.Text = "Fin"
            Label3.Text = "Cible"
            Label4.Text = "Valeurs"
            Label7.Text = "PianoRoll de Destination"
            Label3.Text = "Mesure de Destination"
            Label6.Text = "Mesure de Destination"
            Button1.Text = "Effacer"
            Button2.Text = "Coller"
            Button3.Text = "Transposer"
            Button4.Text = "Annuler"
            Button5.Text = "Coller vers"
            GroupBox1.Text = "SELECTION DES MESURES"
            GroupBox2.Text = "ACTIONS DANS LA SELECTION"
            Label10.Text = "(non compris)"
            '
            TabPage1.Text = "Effacer"
            TabPage2.Text = "Coller"
            TabPage4.Text = "Transposer"
            TabPage3.Text = "Coller Vers"
            TabPage5.Text = "Vélocités aléatoires"
            '
            ToolTip1.SetToolTip(Dest, "N° de mesure")
            ToolTip1.SetToolTip(DestVers, "N° de mesure")
            ToolTip1.SetToolTip(ComboBox1, "Choisir piste PianoRoll")
            ToolTip1.SetToolTip(Debut, "N° mesure")
            ToolTip1.SetToolTip(Terme, "N° mesure (non comprise)")
        Else
            Me.Text = "Actions inside a measures selection"
            Label1.Text = "Start"
            Label2.Text = "End"
            Label3.Text = "Destination measure"
            Label6.Text = "Destination measure"
            Label4.Text = "Values"
            Label7.Text = "PianoRoll Destination"
            Button1.Text = "Clear"
            Button2.Text = "Paste"
            Button3.Text = "Transpose"
            Button4.Text = "Cancel"
            Button5.Text = "Paste to"
            GroupBox1.Text = "MEASURES SELECTION"
            GroupBox2.Text = "ACTIONS IN THE SELECTION"
            Label10.Text = "(not included)"
            '
            TabPage1.Text = "Clear"
            TabPage2.Text = "Paste"
            TabPage4.Text = "Transpose"
            TabPage3.Text = "Paste to"
            TabPage5.Text = "Random velocities"
            '
            '
            ToolTip1.SetToolTip(Dest, "Measure Number")
            ToolTip1.SetToolTip(DestVers, "Measure Number")
            ToolTip1.SetToolTip(ComboBox1, "Choose PianoRoll track")
            ToolTip1.SetToolTip(Debut, "Measure number")
            ToolTip1.SetToolTip(Terme, "Measure number (not included)")
            '
        End If
        '
        Me.StartPosition = FormStartPosition.CenterScreen
        ' 
        Debut.Maximum = Module1.nbMesures
        Debut.Minimum = 1
        '
        Terme.Maximum = Module1.nbMesures + 1
        Terme.Minimum = 1
        '
        Dest.Maximum = Module1.nbMesures
        Dest.Minimum = 1
        '
        ComboBox1.Items.Clear() ' Destination PianoRoll Coller Vers
        ComboBox1.Items.Add(Trim(Form1.Tab_PianoRoll1.Text))
        ComboBox1.Items.Add(Trim(Form1.Tab_PianoRoll2.Text))
        ComboBox1.Items.Add(Trim(Form1.Tab_PianoRoll3.Text))
        '
        ComboBox1.SelectedIndex = (N_Canal + 1) - N_Can1er

    End Sub
    Private Sub ActionsPianoRoll_Activated(sender As Object, e As EventArgs) Handles Me.Activated


    End Sub

    ' Bouton Annuler
    ' **************
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        TypeAction = TypeAct.Rien
        Me.Hide()
    End Sub

    ' Bouton Effacer
    ' **************
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Module1.LangueIHM = "fr" Then
            MessageHV.PTitre = "Avertissement"
            MessageHV.PContenuMess = "Voulez-vous effacer les notes entre la mesure " + Debut.Value.ToString + " et la mesure " + Terme.Value.ToString + " ?"
            MessageHV.PTypBouton = "OuiNon"
        Else
            MessageHV.PTitre = "Warning"
            MessageHV.PContenuMess = "Do you want to clear notes between measure " + Debut.Value.ToString + " and measure " + Terme.Value.ToString + " ?"
            MessageHV.PTypBouton = "OK"
            MessageHV.PTypBouton = "OuiNon"
        End If
        Cacher_FormTransparents()
        MessageHV.ShowDialog()
        If MessageHV.Sortie <> "Non" Then
            Début = Convert.ToInt16(Debut.Value)
            Fin = Convert.ToInt16(Terme.Value)
            TypeAction = TypeAct.Effacer
            Me.Hide()
        End If

    End Sub


    ' Bouton Coller
    ' *************
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Module1.LangueIHM = "fr" Then
            MessageHV.PTitre = "Avertissement"
            MessageHV.PContenuMess = "Voulez-vous coller les notes entre la mesure " + Debut.Value.ToString + " et la mesure " + Terme.Value.ToString + " à partir de la mesure " + Dest.Value.ToString + " ?"
            MessageHV.PTypBouton = "OuiNon"
        Else
            MessageHV.PTitre = "Warning"
            MessageHV.PContenuMess = "Do you want to paste notes between measure " + Debut.Value.ToString + " and measure " + Terme.Value.ToString + " from measure " + Dest.Value.ToString + " ?"
            MessageHV.PTypBouton = "OuiNon"
            MessageHV.PTypBouton = "OK"
            MessageHV.PTypBouton = "OuiNon"
        End If
        Cacher_FormTransparents()
        MessageHV.ShowDialog()
        If MessageHV.Sortie <> "Non" Then
            Début = Convert.ToInt16(Debut.Value)
            Fin = Convert.ToInt16(Terme.Value)
            Destination = Convert.ToInt16(Dest.Value)
            TypeAction = TypeAct.Coller
            Me.Hide()
        End If
    End Sub
    ' Bouton Transposer
    ' *****************
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Module1.LangueIHM = "fr" Then
            MessageHV.PTitre = "Avertissement"
            MessageHV.PContenuMess = "Voulez-vous transposer les notes dans votre sélection comprise entre la mesure " + Debut.Value.ToString + " et la mesure " + Terme.Value.ToString + " ?"
        Else
            MessageHV.PTitre = "Warning"
            MessageHV.PContenuMess = "Do you want to transpose the notes in your selection between the measure " + Debut.Value.ToString + " and measure " + Terme.Value.ToString + " ?"
        End If
        '
        MessageHV.PTypBouton = "OuiNon"
        MessageHV.ShowDialog()
        If MessageHV.Sortie <> "Non" Then
            Début = Convert.ToInt16(Debut.Value)
            Fin = Convert.ToInt16(Terme.Value)
            ValeurTransp = Convert.ToInt16(ComboBox4.Text)
            TypeAction = TypeAct.Transposer
            Me.Hide()
        End If
    End Sub
    ' Bouton CollerVers
    ' *****************
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        If Module1.LangueIHM = "fr" Then
            MessageHV.PTitre = "Avertissement"
            MessageHV.PContenuMess = "Voulez-vous coller les notes entre la mesure " + Debut.Value.ToString + " et la mesure " + Terme.Value.ToString + " à partir de la mesure " + DestVers.Value.ToString + " de la Piste " + ComboBox1.Text + " ?"
            MessageHV.PTypBouton = "OuiNon"
        Else
            MessageHV.PTitre = "Warning"
            MessageHV.PContenuMess = "Do you want to paste notes between measure " + Debut.Value.ToString + " and measure " + Terme.Value.ToString + " from measure " + DestVers.Value.ToString + " of the track " + ComboBox1.Text + " ?"
            MessageHV.PTypBouton = "OuiNon"
            MessageHV.PTypBouton = "OK"
            MessageHV.PTypBouton = "OuiNon"
        End If
        Cacher_FormTransparents()
        MessageHV.ShowDialog()
        If MessageHV.Sortie <> "Non" Then
            Début = Convert.ToInt16(Debut.Value)
            Fin = Convert.ToInt16(Terme.Value)
            Destination = Convert.ToInt16(DestVers.Value)
            TypeAction = TypeAct.CollerVers
            tbl = ComboBox1.Text.Split("-")

            N_PianoR = Convert.ToInt16(tbl(0)) - N_Can1er
            Me.Hide()
        End If
    End Sub
    Private Sub Debut_MouseClick(sender As Object, e As MouseEventArgs) Handles Debut.MouseClick
        Label5.Visible = False
    End Sub
    Private Sub Debut_ValueChanged(sender As Object, e As EventArgs) Handles Debut.ValueChanged
        Label5.Visible = False
        If Debut.Value > Terme.Value Then
            If Terme.Value <> 0 Then
                Debut.Value = Terme.Value
                Label5.Visible = True
                If LangueIHM = "fr" Then
                    Label5.Text = "La début de la sélection ne peut pas être supérieur ou égal à la fin"
                Else
                    Label5.Text = "The beginning of the selection cannot be greater or equal than the end"
                End If
            End If
        End If

    End Sub
    Private Sub Terme_MouseClick(sender As Object, e As MouseEventArgs) Handles Terme.MouseClick
        Label5.Visible = False
    End Sub
    Private Sub Terme_ValueChanged(sender As Object, e As EventArgs) Handles Terme.ValueChanged
        Label5.Visible = False
        If Terme.Value < Debut.Value Then
            Terme.Value = Terme.Value + 1
            Label5.Visible = True
            If LangueIHM = "fr" Then
                Label5.Text = "La fin de la sélection ne peut pas être inférieure ou égale au début"
            Else
                Label5.Text = "The end of the selection cannot be lower or equal than the beginning"
            End If
        End If
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        If Module1.LangueIHM = "fr" Then
            MessageHV.PTitre = "Avertissement"
            MessageHV.PContenuMess = "Voulez-vous modifier les vélocités dans votre sélection comprise entre la mesure " + Debut.Value.ToString + " et la mesure " + Terme.Value.ToString + " ?"
        Else
            MessageHV.PTitre = "Warning"
            MessageHV.PContenuMess = "Do you want to modify velocities in your selection between the measure " + Debut.Value.ToString + " and measure " + Terme.Value.ToString + " ?"
        End If
        '
        MessageHV.PTypBouton = "OuiNon"
        MessageHV.ShowDialog()
        If MessageHV.Sortie <> "Non" Then
            Début = Convert.ToInt16(Debut.Value)
            Fin = Convert.ToInt16(Terme.Value)
            ValMin = Convert.ToInt16(ValeurMin.Text)
            ValMax = Convert.ToInt16(ValeurMax.Text)
            TypeAction = TypeAct.MoidifierVel
            Me.Hide()
        End If
    End Sub
    Private Sub ValeurMin_ValueChanged(sender As Object, e As EventArgs) Handles ValeurMin.ValueChanged
        If ValeurMin.Value > ValeurMax.Value Then
            ValeurMin.Value = ValeurMax.Value
        End If
    End Sub
    Private Sub ValeurMax_ValueChanged(sender As Object, e As EventArgs) Handles ValeurMax.ValueChanged
        If ValeurMin.Value > ValeurMax.Value Then
            ValeurMax.Value = ValeurMin.Value
        End If
    End Sub
End Class