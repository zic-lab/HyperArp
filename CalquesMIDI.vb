Public Class CalquesMIDI

    'Private ReadOnly CoulCalqGammes As Color = Color.Khaki
    'Private ReadOnly CoulCalqAcc As Color = Color.DarkOrange
    'Private ReadOnly CDiv_12_8 As Color = Color.DarkKhaki 'Color.FromArgb(253, 79, 51)
    Private ValNoteCubase As New List(Of String) From {
                   "C-2", "C#-2", "D-2", "D#-2", "E-2", "F-2", "F#-2", "G-2", "G#-2", "A-2", "A#-2", "B-2",
                   "C-1", "C#-1", "D-1", "D#-1", "E-1", "F-1", "F#-1", "G-1", "G#-1", "A-1", "A#-1", "B-1",
                   "C0", "C#0", "D0", "D#0", "E0", "F0", "F#0", "G0", "G#0", "A0", "A#0", "B0",
                    "C1", "C#1", "D1", "D#1", "E1", "F1", "F#1", "G1", "G#1", "A1", "A#1", "B1",
                    "C2", "C#2", "D2", "D#2", "E2", "F2", "F#2", "G2", "G#2", "A2", "A#2", "B2",
                    "C3", "C#3", "D3", "D#3", "E3", "F3", "F#3", "G3", "G#3", "A3", "A#3", "B3",
                    "C4", "C#4", "D4", "D#4", "E4", "F4", "F#4", "G4", "G#4", "A4", "A#4", "B4",
                    "C5", "C#5", "D5", "D#5", "E5", "F5", "F#5", "G5", "G#5", "A5", "A#5", "B5",
                    "C6", "C#6", "D6", "D#6", "E6", "F6", "F#6", "G6", "G#6", "A6", "A#6", "B6",
                    "C7", "C#7", "D7", "D#7", "E7", "F7", "F#7", "G7", "G#7", "A7", "A#7", "B7",
                    "C8", "C#8", "D8", "D#8", "E8", "F8", "F#8", "G8"}


    Private Cadres As New List(Of GroupBox)
    Public Choix As New List(Of CheckBox)
    Private NomCalque As New List(Of Label)
    Private Commentaire As New List(Of Label)
    Public Pédale As New System.Windows.Forms.ComboBox
    Public TessDeb As New System.Windows.Forms.ComboBox
    Public TessFin As New System.Windows.Forms.ComboBox
    Public TessListe As New System.Windows.Forms.ComboBox
    Private Modes As New List(Of RadioButton)
    Public N_PianoR As Integer
    Private Langue As String = "fr"
    Public WriteOnly Property PLangue As String
        Set(ByVal value As String)
            Me.Langue = value
        End Set
    End Property
    Private Sub CalquesMIDI_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim i, ii As Integer
        Dim fontNom = New Font("Calibri", 10, FontStyle.Bold)
        '
        If EnChargementCalquesMIDI Then
            Me.StartPosition = FormStartPosition.CenterScreen
            ' 
            ' Dessin du formulaire
            ' ********************
            ' constantes de positionnement
            Dim PositCadre_Y As Integer = 5
            Dim PositCadre_X As Integer = 5
            Dim Hauteurcadres = 48
            Dim p As New Point
            Dim s As New Size

            ' boucle de construction
            Dim c As New Color
            c = Color.FromArgb(255, 0, 102, 204)
            For i = 0 To nbCalques - 1
                ' cadre
                Cadres.Add(New GroupBox)
                SplitContainer1.Panel2.Controls.Add(Cadres.Item(i))
                s.Width = 460
                s.Height = Hauteurcadres
                Cadres.Item(i).Size = s
                p.X = PositCadre_X
                p.Y = PositCadre_Y + (i * Cadres.Item(i).Size.Height)
                Cadres.Item(i).Location = p
                Cadres.Item(i).BackColor = Color.White

                ' choix
                Choix.Add(New CheckBox)
                Cadres.Item(i).Controls.Add(Choix.Item(i))
                p.X = 5
                p.Y = 10
                s.Height = 15
                s.Width = 15
                Choix.Item(i).Location = p
                Choix.Item(i).Size = s
                Choix.Item(i).Visible = True
                'Choix.Item(i).Checked = PassChoixCalques(i) ' variable globale mise à jour dans le module PianoRoll
                'PassChoixCalques(i) = False ' à l'init le l'index 1 (gamme) est mis à  jour dans cette même procédure après le select i

                'nom
                NomCalque.Add(New Label)
                Cadres.Item(i).Controls.Add(NomCalque.Item(i))
                NomCalque.Item(i).AutoSize = False
                NomCalque.Item(i).BorderStyle = BorderStyle.FixedSingle
                p.X = 25
                p.Y = 10
                s.Height = 36
                s.Width = 70
                NomCalque.Item(i).Size = s
                NomCalque.Item(i).Location = p
                NomCalque.Item(i).Text = "Label" + Str(i)
                NomCalque.Item(i).BackColor = c
                NomCalque.Item(i).ForeColor = Color.White
                NomCalque.Item(i).TextAlign = ContentAlignment.MiddleCenter
                NomCalque.Item(i).Font = fontNom
                '
                'commentaires
                Commentaire.Add(New Label)
                Cadres.Item(i).Controls.Add(Commentaire.Item(i))
                Commentaire.Item(i).AutoSize = False
                Commentaire.Item(i).BorderStyle = BorderStyle.FixedSingle
                p.X = 100
                p.Y = 10
                s.Height = 34
                s.Width = 355
                Commentaire.Item(i).Size = s
                Commentaire.Item(i).Location = p
                Commentaire.Item(i).Text = "Label" + Str(i)
                Commentaire.Item(i).BackColor = c
                Commentaire.Item(i).ForeColor = Color.White
                Commentaire.Item(i).TextAlign = ContentAlignment.MiddleLeft
                Commentaire.Item(i).Visible = True
                '
                'éléments spécifiques
                Select Case i
                    Case 0 ' ton
                        NomCalque.Item(i).BackColor = CoulCalqTon
                        NomCalque.Item(i).ForeColor = Color.White

                        p.X = 100 '265 '280 '5
                        p.Y = 10
                        s.Height = 34
                        s.Width = 365 '190
                        Commentaire.Item(i).Size = s
                        Commentaire.Item(i).Location = p
                        Commentaire.Item(i).BackColor = CoulCalqTon
                        Commentaire.Item(i).ForeColor = Color.White
                        '
                        If Module1.LangueIHM = "fr" Then
                            NomCalque.Item(i).Text = "Ton"
                            Commentaire.Item(i).Text = "Colorisation des notes de la Tonalité courante dans l'éditeur PianoRoll."
                        Else
                            NomCalque.Item(i).Text = "Key"
                            Commentaire.Item(i).Text = "Colorization of the notes of the current Key in the PianoRoll editor."
                        End If
                        '
                    Case 1 ' gamme
                        NomCalque.Item(i).BackColor = CoulCalqGammes
                        Commentaire.Item(i).BackColor = CoulCalqGammes
                        NomCalque.Item(i).ForeColor = Color.Black
                        Commentaire.Item(i).ForeColor = Color.Black
                        '
                        'Choix.Item(i).Checked = True ' de base, à la création,  le calque MIDI activé est le Calques Gamme
                        'PassChoixCalques(i) = True
                        '
                        If Module1.LangueIHM = "fr" Then
                            NomCalque.Item(i).Text = "Gamme"
                            Commentaire.Item(i).Text = "Colorisation des notes de la Gamme courante dans l'éditeur PianoRoll."
                        Else
                            NomCalque.Item(i).Text = "Scale"
                            Commentaire.Item(i).Text = "Colorization of the notes of the current Scale in the PianoRoll editor."
                        End If
                    '
                    Case 2 ' accord
                        NomCalque.Item(i).BackColor = CoulCalqAcc
                        Commentaire.Item(i).BackColor = CoulCalqAcc
                        NomCalque.Item(i).ForeColor = Color.Black
                        Commentaire.Item(i).ForeColor = Color.Black
                        '
                        If Module1.LangueIHM = "fr" Then
                            NomCalque.Item(i).Text = "Accord"
                            Commentaire.Item(i).Text = "Colorisation des notes de l'Accord courant dans l'éditeur PianoRoll."
                        Else
                            NomCalque.Item(i).Text = "Chord"
                            Commentaire.Item(i).Text = "Colorization of the notes of the current Chord in the PianoRoll editor."
                        End If

                    Case 3 ' 12/8
                        NomCalque.Item(i).Text = "12/8"
                        NomCalque.Item(i).BackColor = CDiv_12_8
                        NomCalque.Item(i).ForeColor = Color.Black
                        Commentaire.Item(i).BackColor = CDiv_12_8
                        Commentaire.Item(i).ForeColor = Color.Black
                        '
                        If Module1.LangueIHM = "fr" Then
                            Commentaire.Item(i).Text = "Colorisation d'une métrique 12/8 dans l'éditeur PianoRoll."
                        Else
                            Commentaire.Item(i).Text = "Colorization of a 12/8 meter in the PianoRoll editor."
                        End If
                        '

                    Case 4 ' pédale
                        NomCalque.Item(i).BackColor = CoulCalqPed
                        NomCalque.Item(i).ForeColor = Color.Black
                        p.X = 175
                        p.Y = 10
                        s.Height = 34
                        s.Width = 280
                        Commentaire.Item(i).Size = s
                        Commentaire.Item(i).Location = p
                        Commentaire.Item(i).BackColor = CoulCalqPed
                        Commentaire.Item(i).ForeColor = Color.Black
                        '
                        Cadres.Item(i).Controls.Add(Pédale)
                        p.X = 110
                        p.Y = 17
                        s.Height = 20
                        s.Width = 50
                        Pédale.Location = p
                        Pédale.Size = s
                        '
                        Pédale.Items.Add("C")
                        Pédale.Items.Add("C#")
                        Pédale.Items.Add("D")
                        Pédale.Items.Add("D#")
                        Pédale.Items.Add("E")
                        Pédale.Items.Add("F")
                        Pédale.Items.Add("F#")
                        Pédale.Items.Add("G")
                        Pédale.Items.Add("G#")
                        Pédale.Items.Add("A")
                        Pédale.Items.Add("A#")
                        Pédale.Items.Add("B")
                        Pédale.SelectedIndex = 0

                        '
                        If Module1.LangueIHM = "fr" Then
                            NomCalque.Item(i).Text = "Pédale"
                            Commentaire.Item(i).Text = "Colorisation d'une Note constante dans l'éditeur PianoRoll."
                        Else
                            NomCalque.Item(i).Text = "Pédale"
                            Commentaire.Item(i).Text = "Colorization of a Constant Note in the PianoRoll editor."
                        End If
                        '

                    Case 5 ' tessiture 'ensemble continu des notes qui peuvent jouées par un instrument
                        NomCalque.Item(i).Text = "Tessiture"
                        NomCalque.Item(i).BackColor = CoulCalqTess
                        NomCalque.Item(i).ForeColor = Color.Black
                        Commentaire.Item(i).BackColor = CoulCalqTess
                        Commentaire.Item(i).ForeColor = Color.Black
                        '
                        p.X = 225
                        p.Y = 10
                        s.Height = 61
                        s.Width = 230
                        Commentaire.Item(i).AutoSize = False
                        Commentaire.Item(i).TextAlign = ContentAlignment.TopLeft
                        Commentaire.Item(i).Size = s
                        Commentaire.Item(i).Location = p
                        '
                        If Module1.LangueIHM = "fr" Then
                            Commentaire.Item(i).Text = "Ensemble continu des notes qui peuvent être jouées par un instrument."
                        Else
                            Commentaire.Item(i).Text = "Continuous set of notes that can be played by an instrument."
                        End If
                        '
                        s.Width = 460
                        s.Height = Hauteurcadres + 30
                        Cadres.Item(i).Size = s
                        Cadres.Item(i).Controls.Add(TessDeb)
                        p.X = 110
                        p.Y = 17
                        s.Height = 2
                        s.Width = 50
                        TessDeb.Location = p
                        TessDeb.Size = s
                        '
                        For ii = ValNoteCubase.Count - 1 To 0 Step -1
                            TessDeb.Items.Add(ValNoteCubase(ii))
                        Next
                        TessDeb.SelectedIndex = TessDeb.Items.Count - 1
                        '
                        Cadres.Item(i).Controls.Add(TessFin)
                        p.X = 165
                        p.Y = 17
                        s.Height = 20
                        s.Width = 50
                        TessFin.Location = p
                        TessFin.Size = s
                        '
                        For ii = ValNoteCubase.Count - 1 To 0 Step -1
                            TessFin.Items.Add(ValNoteCubase(ii))
                        Next
                        TessFin.SelectedIndex = 0
                        '
                        Cadres.Item(i).Controls.Add(TessListe)
                        p.X = 25
                        p.Y = 50
                        s.Height = 20
                        s.Width = 190
                        TessListe.Location = p
                        TessListe.Size = s
                        '
                        TessListe.Items.Add("Off")
                        TessListe.Items.Add("*** - Violons - ***")
                        TessListe.Items.Add("Violon - Vln. - G2 C6")
                        TessListe.Items.Add("Violon Alto - Vla. - G2 E5")
                        TessListe.Items.Add("Violoncelles - Vlc. - C1 C4")
                        TessListe.Items.Add("ContreBasse - Cb. - B1 C3")

                        TessListe.Items.Add("*** - Flutes - ***")
                        TessListe.Items.Add("Piccolo - Picc. Alto - D4 C7")
                        TessListe.Items.Add("Alto Flute - Fl. Alto - G2 C6")
                        TessListe.Items.Add("Flute - Fl. - B2 D6")

                        TessListe.Items.Add("*** - Hautbois - ***")
                        TessListe.Items.Add("Hautbois - Htb. - B2 G5")

                        TessListe.Items.Add("*** - Bassons - ***")
                        TessListe.Items.Add("Basson - Bas. - B0 G4")
                        TessListe.Items.Add("Contre Basson - Cbs. - B1 D2")

                        TessListe.Items.Add("*** - Clarinettes - ***")
                        TessListe.Items.Add("Clarinette Eb - Clar. - G2 G5")
                        TessListe.Items.Add("Clarinette Bb - Clar. - D2 G5")
                        TessListe.Items.Add("Clarinette basse - Clar. - B0 F4")

                        TessListe.Items.Add("*** - Cuivres - ***")
                        TessListe.Items.Add("Trompette piccolo - Trp. - D3 C6")
                        TessListe.Items.Add("Trompette - Trp. - G2 F5")
                        TessListe.Items.Add("Cor anglais - Cor angl. - D1 F3")
                        TessListe.Items.Add("Trombone - Trb. - D1 F4")
                        TessListe.Items.Add("Trombone base - Trb. b. - D1 F4")
                        TessListe.Items.Add("Tuba - Tub. - B1 G3")

                        TessListe.Items.Add("*** - Voix - ***")
                        TessListe.Items.Add("Soprano - S. - C3 C5")
                        TessListe.Items.Add("Alto - A. - F2 F4")
                        TessListe.Items.Add("Ténor - T. - C2 G3")
                        TessListe.Items.Add("Basse - B. - C2 G3")
                        '
                        TessListe.SelectedIndex = 0
                        '
                        TessListe.Visible = False
                        AddHandler TessListe.SelectedIndexChanged, AddressOf TessListe_SelectedIndexChanged
                End Select
            Next
            EnChargementCalquesMIDI = False
            PassNouvCalque = False
        End If
        '
        ' Mise à jour du formulaire pour le pianoRoll en cours
        ' ***************************************************
        For i = 0 To nbCalques - 1
            Choix.Item(i).Checked = PassChoixCalques(i)
        Next
        '
        Pédale.SelectedIndex = PassChoixPédale
        TessDeb.SelectedIndex = 127 - ValNoteCubase.IndexOf(PassTessDeb)
        TessFin.SelectedIndex = 127 - ValNoteCubase.IndexOf(PassTessFin)
        TessListe.SelectedIndex = PassTessListe

        '
        ' Maj des boutons de sortie
        ' *************************
        If Module1.LangueIHM = "fr" Then
            Button1.Text = "Annuler"
        Else
            Button1.Text = "Cancel"
        End If
        '
        ' Gestion de l'exclusivité
        ' ************************
        Gest_Eclusiv_EventH()
    End Sub
    Sub Gest_Eclusiv_EventH()
        Cadres.Item(3).Visible = True
        Cadres.Item(4).Visible = False
        Cadres.Item(5).Visible = False
    End Sub
    Sub TessListe_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim i, j As Integer
        Dim tbl1() As String
        Dim tbl2() As String
        '
        If Trim(TessListe.Text) <> "Off" Then
            tbl1 = TessListe.Text.Split("-")
            If Trim(tbl1(0)) <> "***" Then
                tbl2 = Trim(tbl1(2)).Split()
                i = 127 - ValNoteCubase.IndexOf(tbl2(0))
                j = 127 - ValNoteCubase.IndexOf(tbl2(1))
                TessDeb.SelectedIndex = i
                TessFin.SelectedIndex = j
            Else
                RemoveHandler TessListe.SelectedIndexChanged, AddressOf TessListe_SelectedIndexChanged
                TessListe.SelectedIndex = 0
                AddHandler TessListe.SelectedIndexChanged, AddressOf TessListe_SelectedIndexChanged
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        For i = 0 To nbCalques - 1
            Choix.Item(i).Checked = PassChoixCalques(i)
        Next i
        TessDeb.SelectedIndex = 127 - ValNoteCubase.IndexOf(PassTessDeb)
        TessFin.SelectedIndex = 127 - ValNoteCubase.IndexOf(PassTessFin)
        TessListe.SelectedIndex = PassTessListe
        Retour = "Annuler"
        Me.Hide()
    End Sub
    Private Sub ButtonApply_Click(sender As Object, e As EventArgs) Handles ButtonApply.Click
        For i = 0 To nbCalques - 1
            PassChoixCalques(i) = Choix.Item(i).Checked
        Next i
        PassChoixPédale = Pédale.SelectedIndex
        PassTessDeb = Trim(TessDeb.Text)
        PassTessFin = Trim(TessFin.Text)
        PassTessListe = TessListe.SelectedIndex
        Retour = "OK"
        Me.Hide()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        For i = 0 To nbCalques - 1
            Choix.Item(i).Checked = False
            PassChoixCalques(i) = False
        Next
    End Sub

    Private Sub SplitContainer1_Panel2_Paint(sender As Object, e As PaintEventArgs) Handles SplitContainer1.Panel2.Paint

    End Sub
End Class