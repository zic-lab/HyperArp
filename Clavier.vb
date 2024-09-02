Imports System.ComponentModel

Public Class Clavier

    Dim Ouvert As Boolean = True

    Private Sub Clavier_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.KeyPreview = True
        Me.ActiveControl = Nothing
        CheckBox1.Checked = True
        Me.TopMost = True
        If LangueIHM = "fr" Then
            Me.Text = "CLAVIER"
            CheckBox1.Text = "Toujours au dessus"
        Else
            Me.Text = "KEYBOARD"
            CheckBox1.Text = "Always on top"
        End If
        Dim s As New Size
        s.Width = 365
        s.Height = 340
        Me.Size = s
        '
        ComboBox1.Items.Add("1 Modulation")
        ComboBox1.Items.Add("7 Volume")
        ComboBox1.Items.Add("10 Pan")
        ComboBox1.Items.Add("11 Expression")
        If Module1.LangueIHM = "fr" Then
            ComboBox1.Items.Add("50 Libre")
            ComboBox1.Items.Add("51 Libre")
            ComboBox1.Items.Add("52 Libre")
            ComboBox1.Items.Add("53 Libre")
        Else
            ComboBox1.Items.Add("50 Free")
            ComboBox1.Items.Add("51 Free")
            ComboBox1.Items.Add("52 Free")
            ComboBox1.Items.Add("53 Free")
        End If
        '
        ComboBox1.SelectedIndex = 0
        Maj_PRG()
    End Sub

    Sub Maj_PRG()
        If LangueIHM = "fr" Then
            ListPRG.Items.Add("GS/GM off")
            ListPRG.Items.Add("01 Piano acoustique 1")
            ListPRG.Items.Add("02 Piano acoustique 2")
            ListPRG.Items.Add("03 Grand piano électrique")
            ListPRG.Items.Add("04 Piano honkytonk")
            ListPRG.Items.Add("05 Piano électrique 1")
            ListPRG.Items.Add("06 Piano électrique 2")
            ListPRG.Items.Add("07 Clavecin")
            ListPRG.Items.Add("08 Clavicorde")
            ListPRG.Items.Add("09 Célesta")
            ListPRG.Items.Add("10 Carillon")
            ListPRG.Items.Add("11 Boîte à musique")
            ListPRG.Items.Add("12 Vibraphone")
            ListPRG.Items.Add("13 Marimba")
            ListPRG.Items.Add("14 Xylophone")
            ListPRG.Items.Add("15 Cloches tubulaires")
            ListPRG.Items.Add("16 Tympanon")
            ListPRG.Items.Add("17 Orgue à tubes")
            ListPRG.Items.Add("18 Orgue percussif")
            ListPRG.Items.Add("19 Orgue rock")
            ListPRG.Items.Add("20 Orgue d'église")
            ListPRG.Items.Add("21 Orgue vibrato")
            ListPRG.Items.Add("22 Accordéon")
            ListPRG.Items.Add("23 Harmonica")
            ListPRG.Items.Add("24 Bandonéon")
            ListPRG.Items.Add("25 Guitare classique")
            ListPRG.Items.Add("26 Guitare folk")
            ListPRG.Items.Add("27 Guitare jazz")
            ListPRG.Items.Add("28 Guitare élec. pure")
            ListPRG.Items.Add("29 Guitare élec. mute")
            ListPRG.Items.Add("30 Guitare saturée")
            ListPRG.Items.Add("31 Guitare distorsion")
            ListPRG.Items.Add("32 Guitare harmonique")
            ListPRG.Items.Add("33 Basse acoustique")
            ListPRG.Items.Add("34 Basse élec. 1")
            ListPRG.Items.Add("35 Basse élec. 2")
            ListPRG.Items.Add("36 Basse élec. 3")
            ListPRG.Items.Add("37 Basse slap 1")
            ListPRG.Items.Add("38 Basse slap 2")
            ListPRG.Items.Add("39 Basse synth.  1")
            ListPRG.Items.Add("40 Basse synth.  2")
            ListPRG.Items.Add("41 Violon")
            ListPRG.Items.Add("42 Viole")
            ListPRG.Items.Add("43 Violoncelle")
            ListPRG.Items.Add("44 Contrebasse")
            ListPRG.Items.Add("45 Cordes trémolo")
            ListPRG.Items.Add("46 Cordes pizzicato")
            ListPRG.Items.Add("47 Harpe")
            ListPRG.Items.Add("48 Timbales")
            ListPRG.Items.Add("49 Quartet cordes 1")
            ListPRG.Items.Add("50 Quartet cordes 2")
            ListPRG.Items.Add("51 Cordes synth 1")
            ListPRG.Items.Add("52 Cordes synth 2")
            ListPRG.Items.Add("53 Chœurs Aahs")
            ListPRG.Items.Add("54 Voix Oohs")
            ListPRG.Items.Add("55 Voix synthétiseur")
            ListPRG.Items.Add("56 Coup d'orchestre")
            ListPRG.Items.Add("57 Trompette")
            ListPRG.Items.Add("58 Trombone")
            ListPRG.Items.Add("59 Tuba")
            ListPRG.Items.Add("60 Trompette bouchée")
            ListPRG.Items.Add("61 Cors")
            ListPRG.Items.Add("62 Ensemble de cuivres")
            ListPRG.Items.Add("63 Cuivres synthétiseur")
            ListPRG.Items.Add("64 Cuivres synthétiseur")
            ListPRG.Items.Add("65 Saxophone soprano")
            ListPRG.Items.Add("66 Saxophone alto")
            ListPRG.Items.Add("67 Saxophone ténor")
            ListPRG.Items.Add("68 Saxophone baryton")
            ListPRG.Items.Add("69 Hautbois")
            ListPRG.Items.Add("70 Cors anglais")
            ListPRG.Items.Add("71 Basson")
            ListPRG.Items.Add("72 Clarinette")
            ListPRG.Items.Add("73 Flûte piccolo")
            ListPRG.Items.Add("74 Flûte")
            ListPRG.Items.Add("75 Flûte à bec")
            ListPRG.Items.Add("76 Flûte de pan")
            ListPRG.Items.Add("77 Bouteille sifflée")
            ListPRG.Items.Add("78 Shakuhachi")
            ListPRG.Items.Add("79 Sifflet")
            ListPRG.Items.Add("80 Ocarina")
            ListPRG.Items.Add("81 Lead carré")
            ListPRG.Items.Add("82 Lead dents de scie")
            ListPRG.Items.Add("83 Lead orgue")
            ListPRG.Items.Add("84 Lead chiff")
            ListPRG.Items.Add("85 Lead charang")
            ListPRG.Items.Add("86 Lead voix")
            ListPRG.Items.Add("87 Lead quinte)")
            ListPRG.Items.Add("88 Lead basse")
            ListPRG.Items.Add("89 Pad new Age")
            ListPRG.Items.Add("90 Pad warm")
            ListPRG.Items.Add("91 Pad poly")
            ListPRG.Items.Add("92 Pad chœurs")
            ListPRG.Items.Add("93 Pad archet")
            ListPRG.Items.Add("94 Pad métal")
            ListPRG.Items.Add("95 Pad halo")
            ListPRG.Items.Add("96 Pad glissement")
        Else
            ListPRG.Items.Add("GS/GM off")
            ListPRG.Items.Add("01 Acoustic Grand Piano")
            ListPRG.Items.Add("02 Bright Acoustic Piano")
            ListPRG.Items.Add("03 Electric Grand Piano")
            ListPRG.Items.Add("04 Honky-tonk Piano")
            ListPRG.Items.Add("05 Electric Piano 1")
            ListPRG.Items.Add("06 Electric Piano 2")
            ListPRG.Items.Add("07 07 Harpsichord")
            ListPRG.Items.Add("08 Clavinet")
            ListPRG.Items.Add("09 Celesta")
            ListPRG.Items.Add("10 Glockenspiel")
            ListPRG.Items.Add("11 Music Box")
            ListPRG.Items.Add("12 Vibraphone")
            ListPRG.Items.Add("13 Marimba")
            ListPRG.Items.Add("14 Xylophone")
            ListPRG.Items.Add("15 Tubular Bells")
            ListPRG.Items.Add("16 Dulcimer")
            ListPRG.Items.Add("17 Drawbar Organ")
            ListPRG.Items.Add("18 Percussive Organ")
            ListPRG.Items.Add("19 Rock Organ")
            ListPRG.Items.Add("20 Church Organ")
            ListPRG.Items.Add("21 Reed Organ")
            ListPRG.Items.Add("22 Accordion")
            ListPRG.Items.Add("23 Harmonica")
            ListPRG.Items.Add("24 Tango Accordion")
            ListPRG.Items.Add("25 Acoustic nylon")
            ListPRG.Items.Add("26 Acoustic steel")
            ListPRG.Items.Add("27 Electric jazz")
            ListPRG.Items.Add("28 Electric clean")
            ListPRG.Items.Add("29 Electric muted")
            ListPRG.Items.Add("30 Overdriven Guitar")
            ListPRG.Items.Add("31 Distortion Guitar")
            ListPRG.Items.Add("32 Guitar Harmonics")
            ListPRG.Items.Add("33 Acoustic Bass")
            ListPRG.Items.Add("34 Electric Bass finger")
            ListPRG.Items.Add("35 Electric Bass pick")
            ListPRG.Items.Add("36 Fretless Bass")
            ListPRG.Items.Add("37 Slap Bass 1")
            ListPRG.Items.Add("38 Slap Bass 2")
            ListPRG.Items.Add("39 Synth Bass 1")
            ListPRG.Items.Add("40 Synth Bass 2")
            ListPRG.Items.Add("41 Violon")
            ListPRG.Items.Add("42 Viola")
            ListPRG.Items.Add("43 Cello")
            ListPRG.Items.Add("44 Contrabass")
            ListPRG.Items.Add("45 Tremolo Strings")
            ListPRG.Items.Add("46 Pizzicato Strings")
            ListPRG.Items.Add("47 Orchestral Harp")
            ListPRG.Items.Add("48 Timpani")
            ListPRG.Items.Add("49 String Ensemble 1")
            ListPRG.Items.Add("50 String Ensemble 2")
            ListPRG.Items.Add("51 Synth Strings 1")
            ListPRG.Items.Add("52 Synth Strings 2")
            ListPRG.Items.Add("53 Choir Aahs")
            ListPRG.Items.Add("54 Voice Oohs")
            ListPRG.Items.Add("55 Synth Choir")
            ListPRG.Items.Add("56 Orchestra Hit")
            ListPRG.Items.Add("57 Trumpet")
            ListPRG.Items.Add("58 Trombone")
            ListPRG.Items.Add("59 Tuba")
            ListPRG.Items.Add("60 Muted Trumpet")
            ListPRG.Items.Add("61 French Horn")
            ListPRG.Items.Add("62 Brass Section")
            ListPRG.Items.Add("63 Synth Brass 1")
            ListPRG.Items.Add("64 Synth Brass 2")
            ListPRG.Items.Add("65 Soprano Sax")
            ListPRG.Items.Add("66 Alto Sax")
            ListPRG.Items.Add("67 Tenor Sax")
            ListPRG.Items.Add("68 Baritone Sax")
            ListPRG.Items.Add("69 Oboe")
            ListPRG.Items.Add("70 English Horn")
            ListPRG.Items.Add("71 Bassoon")
            ListPRG.Items.Add("72 Clarinet")
            ListPRG.Items.Add("73 Piccolo")
            ListPRG.Items.Add("74 Flute")
            ListPRG.Items.Add("75 Recorder")
            ListPRG.Items.Add("76 Pan Flute")
            ListPRG.Items.Add("77 Blown bottle")
            ListPRG.Items.Add("78 Shakuhachi")
            ListPRG.Items.Add("79 Whistle")
            ListPRG.Items.Add("80 Ocarina")
            ListPRG.Items.Add("81 Lead 1 square")
            ListPRG.Items.Add("82 Lead 2 sawtooth")
            ListPRG.Items.Add("83 Lead 3 calliope")
            ListPRG.Items.Add("84 Lead chiff")
            ListPRG.Items.Add("85 Lead charang")
            ListPRG.Items.Add("86 Lead voice")
            ListPRG.Items.Add("87 Lead 7 fifths")
            ListPRG.Items.Add("88 Lead 8 bass + lead")
            ListPRG.Items.Add("89 Pad new Age")
            ListPRG.Items.Add("90 Pad warm")
            ListPRG.Items.Add("91 Pad 3 polysynth")
            ListPRG.Items.Add("92 Pad 4 choir")
            ListPRG.Items.Add("93 Pad 5 bowed")
            ListPRG.Items.Add("94 Pad 6 metallic")
            ListPRG.Items.Add("95 Pad 7 halo")
            ListPRG.Items.Add("96 Pad 8 sweep")
        End If   '
        ListPRG.SelectedIndex = 0
    End Sub

    Private Sub Clavier_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        Dim note As String = ""
        Dim ValeurNote As Byte
        Dim canal As Byte
        Dim Vélo As Byte
        Dim octave As String = Convert.ToString(OctaveEcoute.Value)
        '
        Dim c As String
        'If Not (OctaveEcoute.Focused) Then
        c = e.KeyCode.ToString
            '
            e.SuppressKeyPress = True

            Select Case c
                Case "s", "S" ' C
                    note = "c"
                    Lettre_S.BackColor = Color.Tan
                Case "e", "E" ' C#
                    note = "c#"
                    Lettre_E.BackColor = Color.Tan
                Case "d", "D" ' D
                    note = "d"
                    Lettre_D.BackColor = Color.Tan
                Case "r", "R" ' D#
                    note = "d#"
                    Lettre_R.BackColor = Color.Tan
                Case "f", "F" ' E
                    note = "e"
                    Lettre_F.BackColor = Color.Tan
                Case "g", "G" ' F
                    note = "f"
                    Lettre_G.BackColor = Color.Tan
                Case "y", "Y" ' F#
                    note = "f#"
                    Lettre_Y.BackColor = Color.Tan
                Case "h", "H" ' G
                    note = "g"
                    Lettre_H.BackColor = Color.Tan
                Case "u", "U" ' G#
                    note = "g#"
                    Lettre_U.BackColor = Color.Tan
                Case "j", "J" ' A
                    note = "a"
                    Lettre_J.BackColor = Color.Tan
                Case "i", "I" ' A#
                    note = "a#"
                    Lettre_I.BackColor = Color.Tan
                Case "k", "K" ' B
                    note = "b"
                    Lettre_K.BackColor = Color.Tan
                Case "l", "L" ' B
                    note = "c"
                    octave = Convert.ToString(Val(octave) + 1)
                    Lettre_L.BackColor = Color.Tan
            End Select
            '
            If Trim(note) <> "" Then
                If Form1.DicPiano(UCase(c)) = False Then
                    Form1.DicPiano(UCase(c)) = True
                    note = note + octave
                    ValeurNote = Convert.ToByte(ListNotesd.IndexOf(note))
                    canal = Convert.ToByte(CanalEcoute.Value)
                    Vélo = Convert.ToByte(VéloEcoute.Value)
                Form1.JouerNote2(ValeurNote, canal, Vélo)
                '
                ' Afficher diode rouge sur le piano
                ' *********************************
                'AfficherNote = True
                '
                'j = ValeurNote
                'Form1.Touche_CouleurPréced(j) = Form1.LabelPiano.Item(j).BackColor 'LabelPianoMidiIn.Item(j).BackColor
                'Form1.LabelPianoMidiIn.Item(j).BackColor = Color.Red
            End If
            End If
        'End If
    End Sub
    Sub Playnote(c As String)
        Dim note As String = ""
        Dim ValeurNote As Byte
        Dim canal As Byte
        Dim Vélo As Byte
        Dim octave As String = Convert.ToString(OctaveEcoute.Value)
        '
        RAZ_CouleurNotes() ' toutes les touches "S,D,F, .. " en blanc

        Select Case c
            Case "s", "S" ' C
                note = "c"
                Lettre_S.BackColor = Color.Tan
            Case "e", "E" ' C#
                note = "c#"
                Lettre_E.BackColor = Color.Tan
            Case "d", "D" ' D
                note = "d"
                Lettre_D.BackColor = Color.Tan
            Case "r", "R" ' D#
                note = "d#"
                Lettre_R.BackColor = Color.Tan
            Case "f", "F" ' E
                note = "e"
                Lettre_F.BackColor = Color.Tan
            Case "g", "G" ' F
                note = "f"
                Lettre_G.BackColor = Color.Tan
            Case "y", "Y" ' F#
                note = "f#"
                Lettre_Y.BackColor = Color.Tan
            Case "h", "H" ' G
                note = "g"
                Lettre_H.BackColor = Color.Tan
            Case "u", "U" ' G#
                note = "g#"
                Lettre_U.BackColor = Color.Tan
            Case "j", "J" ' A
                note = "a"
                Lettre_J.BackColor = Color.Tan
            Case "i", "I" ' A#
                note = "a#"
                Lettre_I.BackColor = Color.Tan
            Case "k", "K" ' B
                note = "b"
                Lettre_K.BackColor = Color.Tan
            Case "l", "L" ' B
                note = "c"
                octave = Convert.ToString(Val(octave) + 1)
                Lettre_L.BackColor = Color.Tan
        End Select
        '
        If Trim(note) <> "" Then
            If Form1.DicPiano(UCase(c)) = False Then
                Form1.DicPiano(UCase(c)) = True
                note = note + octave ' concaténation
                ValeurNote = Convert.ToByte(ListNotesd.IndexOf(note))
                canal = Convert.ToByte(CanalEcoute.Value)
                Vélo = Convert.ToByte(VéloEcoute.Value)
                Form1.JouerNote2(ValeurNote, canal, Vélo)
                '
                ' Afficher diode rouge sur le piano
                ' *********************************
                'AfficherNote = True
                '
                'j = ValeurNote
                'Form1.Touche_CouleurPréced(j) = Form1.LabelPiano.Item(j).BackColor 'LabelPianoMidiIn.Item(j).BackColor
                'Form1.LabelPianoMidiIn.Item(j).BackColor = Color.Red


                'Label19.Text = Trim(UCase(note))
                'Label21.Text = Trim(Str(canal))
            End If
        End If

    End Sub
    Public Sub RAZ_CouleurNotes()
        Lettre_S.BackColor = Color.White
        Lettre_E.BackColor = Color.White
        Lettre_D.BackColor = Color.White
        Lettre_R.BackColor = Color.White
        Lettre_F.BackColor = Color.White
        Lettre_G.BackColor = Color.White
        Lettre_Y.BackColor = Color.White
        Lettre_H.BackColor = Color.White
        Lettre_U.BackColor = Color.White
        Lettre_J.BackColor = Color.White
        Lettre_I.BackColor = Color.White
        Lettre_K.BackColor = Color.White
        Lettre_L.BackColor = Color.White
    End Sub
    Sub Stopnote(c As String)
        Dim note As String = ""
        Dim ValeurNote As Byte
        Dim canal As Byte
        Dim Vélo As Byte
        Dim octave As String = Convert.ToString(OctaveEcoute.Value)


        Select Case c
            Case "s", "S" ' C
                note = "c"
                Lettre_S.BackColor = Color.White
            Case "e", "E" ' C#
                note = "c#"
                Lettre_E.BackColor = Color.White
            Case "d", "D" ' D
                note = "d"
                Lettre_D.BackColor = Color.White
            Case "r", "R" ' D#
                note = "d#"
                Lettre_R.BackColor = Color.White
            Case "f", "F" ' E
                note = "e"
                Lettre_F.BackColor = Color.White
            Case "g", "G" ' F
                note = "f"
                Lettre_G.BackColor = Color.White
            Case "y", "Y" ' F#
                note = "f#"
                Lettre_Y.BackColor = Color.White
            Case "h", "H" ' G
                note = "g"
                Lettre_H.BackColor = Color.White
            Case "u", "U" ' G#
                note = "g#"
                Lettre_U.BackColor = Color.White
            Case "j", "J" ' A
                note = "a"
                Lettre_J.BackColor = Color.White
            Case "i", "I" ' A#
                note = "a#"
                Lettre_I.BackColor = Color.White
            Case "k", "K" ' B
                note = "b"
                Lettre_K.BackColor = Color.White
            Case "l", "L" ' B
                note = "c"
                octave = Convert.ToString(Val(octave) + 1)
                Lettre_L.BackColor = Color.White
        End Select
        '
        If Trim(note) <> "" Then
            Form1.DicPiano(UCase(c)) = False
            note = note + octave
            ValeurNote = Convert.ToByte(ListNotesd.IndexOf(note))
            canal = Convert.ToByte(CanalEcoute.Value)
            Vélo = Convert.ToByte(VéloEcoute.Text)
            Form1.StoperNote2(ValeurNote, canal, Vélo)
            '
            ' Effacer diode rouge sur le piano
            ' *********************************
            'j = ValeurNote
            'If Form1.Touche_CouleurPréced(j) = Color.LightGreen Then
            'Form1.LabelPianoMidiIn.Item(j).BackColor = Color.White 'LabelPiano(j).BackColor 'Touche_CouleurPréced
            'Else
            '   Form1.LabelPianoMidiIn.Item(j).BackColor = Form1.LabelPiano(j).BackColor 'Touche_CouleurPréced
            'End If
        End If
    End Sub
    Private Sub Clavier_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp
        Dim note As String = ""
        Dim ValeurNote As Byte
        Dim canal As Byte
        Dim Vélo As Byte
        Dim c As String
        Dim octave As String = Convert.ToString(OctaveEcoute.Value)

        'If Not (OctaveEcoute.Focused) Then
        c = e.KeyCode.ToString
            Select Case c
                Case "s", "S" ' C
                    note = "c"
                    Lettre_S.BackColor = Color.White
                Case "e", "E" ' C#
                    note = "c#"
                    Lettre_E.BackColor = Color.White
                Case "d", "D" ' D
                    note = "d"
                    Lettre_D.BackColor = Color.White
                Case "r", "R" ' D#
                    note = "d#"
                    Lettre_R.BackColor = Color.White
                Case "f", "F" ' E
                    note = "e"
                    Lettre_F.BackColor = Color.White
                Case "g", "G" ' F
                    note = "f"
                    Lettre_G.BackColor = Color.White
                Case "y", "Y" ' F#
                    note = "f#"
                    Lettre_Y.BackColor = Color.White
                Case "h", "H" ' G
                    note = "g"
                    Lettre_H.BackColor = Color.White
                Case "u", "U" ' G#
                    note = "g#"
                    Lettre_U.BackColor = Color.White
                Case "j", "J" ' A
                    note = "a"
                    Lettre_J.BackColor = Color.White
                Case "i", "I" ' A#
                    note = "a#"
                    Lettre_I.BackColor = Color.White
                Case "k", "K" ' B
                    note = "b"
                    Lettre_K.BackColor = Color.White
                Case "l", "L" ' B
                    note = "c"
                    octave = Convert.ToString(Val(octave) + 1)
                    Lettre_L.BackColor = Color.White
            End Select
        '
        If Trim(note) <> "" Then
            Form1.DicPiano(UCase(c)) = False
            note = note + octave
            ValeurNote = Convert.ToByte(ListNotesd.IndexOf(note))
            canal = Convert.ToByte(CanalEcoute.Value)
            Vélo = Convert.ToByte(VéloEcoute.Text)
            Form1.StoperNote2(ValeurNote, canal, Vélo)
            '
            ' Effacer diode rouge sur le piano
            ' *********************************
            'j = ValeurNote
            'If Form1.Touche_CouleurPréced(j) = Color.LightGreen Then
            'Form1.LabelPianoMidiIn.Item(j).BackColor = Color.White 'LabelPiano(j).BackColor 'Touche_CouleurPréced
            'Else
            'Form1.LabelPianoMidiIn.Item(j).BackColor = Form1.LabelPiano(j).BackColor 'Touche_CouleurPréced
        End If
        'End If
        'End If
    End Sub

    Private Sub ListPRG_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListPRG.SelectedIndexChanged
        If Trim(ListPRG.SelectedIndex - 1) <> -1 Then
            If Not (Form1.SortieMidi.Item(Form1.ChoixSortieMidi).IsOpen) Then
                Form1.SortieMidi.Item(Form1.ChoixSortieMidi).Open()
            End If
            'tbl = Split(Trim(PRG))
            Form1.SortieMidi.Item(Form1.ChoixSortieMidi).SendProgramChange(Convert.ToByte(CanalEcoute.Value - 1), ListPRG.SelectedIndex - 1)
        End If
    End Sub
    Private Sub ListPRG_KeyDown(sender As Object, e As KeyEventArgs) Handles ListPRG.KeyDown
        e.SuppressKeyPress = True
    End Sub
    Private Sub Lettre_S_MouseDown(sender As Object, e As MouseEventArgs) Handles Lettre_S.MouseDown
        Playnote("S")
    End Sub
    Private Sub Lettre_S_MouseUp(sender As Object, e As MouseEventArgs) Handles Lettre_S.MouseUp
        Stopnote("S")
    End Sub
    Private Sub Lettre_E_MouseDown(sender As Object, e As MouseEventArgs) Handles Lettre_E.MouseDown
        Playnote("E")
    End Sub
    Private Sub Lettre_E_MouseUp(sender As Object, e As MouseEventArgs) Handles Lettre_E.MouseUp
        Stopnote("E")
    End Sub
    Private Sub Lettre_D_MouseDown(sender As Object, e As MouseEventArgs) Handles Lettre_D.MouseDown
        Playnote("D")
    End Sub
    Private Sub Lettre_D_MouseUp(sender As Object, e As MouseEventArgs) Handles Lettre_D.MouseUp
        Stopnote("D")
    End Sub
    Private Sub Lettre_R_MouseDown(sender As Object, e As MouseEventArgs) Handles Lettre_R.MouseDown
        Playnote("R")
    End Sub
    Private Sub Lettre_R_MouseUp(sender As Object, e As MouseEventArgs) Handles Lettre_R.MouseUp
        Stopnote("R")
    End Sub
    Private Sub Lettre_F_MouseDown(sender As Object, e As MouseEventArgs) Handles Lettre_F.MouseDown
        Playnote("F")
    End Sub
    Private Sub Lettre_F_MouseUp(sender As Object, e As MouseEventArgs) Handles Lettre_F.MouseUp
        Stopnote("F")
    End Sub
    Private Sub Lettre_G_MouseDown(sender As Object, e As MouseEventArgs) Handles Lettre_G.MouseDown
        Playnote("G")
    End Sub
    Private Sub Lettre_G_MouseUp(sender As Object, e As MouseEventArgs) Handles Lettre_G.MouseUp
        Stopnote("G")
    End Sub
    Private Sub Lettre_Y_MouseDown(sender As Object, e As MouseEventArgs) Handles Lettre_Y.MouseDown
        Playnote("Y")
    End Sub
    Private Sub Lettre_Y_MouseUp(sender As Object, e As MouseEventArgs) Handles Lettre_Y.MouseUp
        Stopnote("Y")
    End Sub
    Private Sub Lettre_H_MouseDown(sender As Object, e As MouseEventArgs) Handles Lettre_H.MouseDown
        Playnote("H")
    End Sub
    Private Sub Lettre_H_MouseUp(sender As Object, e As MouseEventArgs) Handles Lettre_H.MouseUp
        Stopnote("H")
    End Sub
    Private Sub Lettre_U_MouseDown(sender As Object, e As MouseEventArgs) Handles Lettre_U.MouseDown
        Playnote("U")
    End Sub
    Private Sub Lettre_U_MouseUp(sender As Object, e As MouseEventArgs) Handles Lettre_U.MouseUp
        Stopnote("U")
    End Sub
    Private Sub Lettre_J_MouseDown(sender As Object, e As MouseEventArgs) Handles Lettre_J.MouseDown
        Playnote("J")
    End Sub
    Private Sub Lettre_J_MouseUp(sender As Object, e As MouseEventArgs) Handles Lettre_J.MouseUp
        Stopnote("J")
    End Sub
    Private Sub Lettre_I_MouseDown(sender As Object, e As MouseEventArgs) Handles Lettre_I.MouseDown
        Playnote("I")
    End Sub
    Private Sub Lettre_I_MouseUp(sender As Object, e As MouseEventArgs) Handles Lettre_I.MouseUp
        Stopnote("I")
    End Sub
    Private Sub Lettre_K_MouseDown(sender As Object, e As MouseEventArgs) Handles Lettre_K.MouseDown
        Playnote("K")
    End Sub
    Private Sub Lettre_K_MouseUp(sender As Object, e As MouseEventArgs) Handles Lettre_K.MouseUp
        Stopnote("K")
    End Sub
    Private Sub Lettre_L_MouseDown(sender As Object, e As MouseEventArgs) Handles Lettre_L.MouseDown
        Playnote("L")
    End Sub
    Private Sub Lettre_L_MouseUp(sender As Object, e As MouseEventArgs) Handles Lettre_L.MouseUp
        Stopnote("L")
    End Sub

    Private Sub VolumeEcoute_Scroll(sender As Object, e As EventArgs) Handles VolumeEcoute.Scroll
        Dim volume As Byte = CByte(VolumeEcoute.Value)
        Dim canal As Byte = CByte(CanalEcoute.Value - 1)
        If Not (Form1.SortieMidi.Item(Form1.ChoixSortieMidi).IsOpen) Then
            Form1.SortieMidi.Item(Form1.ChoixSortieMidi).Open()
        End If
        Form1.SortieMidi.Item(Form1.ChoixSortieMidi).SendControlChange(canal, 7, volume)
    End Sub
    Private Sub Clavier_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        Me.Opacity = 1
        Panel1.Visible = True
        Panel2.Visible = True
    End Sub
    Private Sub Clavier_Deactivate(sender As Object, e As EventArgs) Handles Me.Deactivate
        If Ouvert Then
            Me.Opacity = 1 ' 0.7 opacité supprimée le 21-05-20
            'Panel1.Visible = False
            'Panel2.Visible = False
        End If
    End Sub

    Private Sub Clavier_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Ouvert = False
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            Me.TopMost = True
        Else
            Me.TopMost = False
        End If
    End Sub

    Private Sub ValCtrl_Scroll(sender As Object, e As EventArgs) Handles ValCtrl.Scroll
        Dim tbl() As String
        tbl = ComboBox1.Text.Split()
        Dim vCtrl As Byte = CByte(ValCtrl.Value)
        Dim Ctrl As Byte = CByte(tbl(0))
        Dim canal As Byte = CByte(CanalEcoute.Value - 1)
        If Not (Form1.SortieMidi.Item(Form1.ChoixSortieMidi).IsOpen) Then
            Form1.SortieMidi.Item(Form1.ChoixSortieMidi).Open()
        End If
        Form1.SortieMidi.Item(Form1.ChoixSortieMidi).SendControlChange(canal, Ctrl, vCtrl)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub OctaveEcoute_ValueChanged(sender As Object, e As EventArgs) Handles OctaveEcoute.ValueChanged

    End Sub
End Class