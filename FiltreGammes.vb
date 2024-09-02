Public Class FiltreGammes
    Dim déjàchargée As Boolean = False

    Private RetourValue As String
    Public ReadOnly Property Retour() As String
        Get
            Return RetourValue
        End Get
    End Property
    Private AccordsValue As String
    Public Property Accords() As String
        Get
            Return AccordsValue
        End Get
        Set(ByVal value As String)
            AccordsValue = value
        End Set
    End Property


    Const NbTypeGammes = 6
    Dim LTitre As New List(Of CheckBox)
    Dim LTypeG As New List(Of String)
    Dim LstUsuelles As New List(Of String)
    Dim Lst5notes As New List(Of String)
    Dim Lst6notes As New List(Of String)
    Dim Lst7notes As New List(Of String)
    Dim Lst8notes As New List(Of String)
    Dim Lst9notes As New List(Of String)

    Dim LUsuelles As New List(Of CheckBox)
    Dim L5notes As New List(Of CheckBox)
    Dim L6notes As New List(Of CheckBox)
    Dim L7notes As New List(Of CheckBox)
    Dim L8notes As New List(Of CheckBox)
    Dim L9notes As New List(Of CheckBox)

    Public oo As New RechercheG_v2
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Dim a As String = My.Resources.FichierListeGammes
        'oo.GammesBases = a ' mise à jour du fichier global des gammes 
        'a = My.Resources.FichierListeAccords
        'oo.AccordsBases = a ' mise à jour du fichier global des accords
        LTypeG.Add("Usuelles")
        LTypeG.Add("5 notes")
        LTypeG.Add("6 notes")
        LTypeG.Add("7 notes")
        LTypeG.Add("8 notes")
        LTypeG.Add("9 notes")
        '
        'a = oo.ApparteanceG("E m7-C-D m7", "Maj MinH")
        'OpéEnsembles()    
        'test_Pentatoniques("E m7")
        If Not déjàchargée Then
            Construction()
        End If
        déjàchargée = True
    End Sub
    Sub Construction()
        Dim p As New Point
        Dim s As New Size

        ' Panneaux d'affichage des gammes
        Panel1.BackColor = Color.FromArgb(135, 135, 204)
        Panel2.BackColor = Color.FromArgb(204, 204, 0)
        Panel3.BackColor = Color.FromArgb(153, 102, 153)
        Panel4.BackColor = Color.FromArgb(255, 204, 51)
        Panel5.BackColor = Color.FromArgb(102, 153, 204)
        Panel6.BackColor = Color.FromArgb(102, 153, 51)

        ' Maj liste listes types de gammes (par nombre de notes)
        Dim line As String
        Dim tbl() As String
        Using sr As IO.StringReader = New IO.StringReader(My.Resources.FichierListeGammes)
            While sr.Peek() >= 0 ' Boucler jusqu'à la fin du fichier
                line = sr.ReadLine() ' Lire chaque ligne
                tbl = line.Split(";")
                Select Case tbl(1)
                    Case "Usuelle"
                        LstUsuelles.Add(tbl(2))
                    Case "5 notes"
                        Lst5notes.Add(tbl(2))
                    Case "6 notes"
                        Lst6notes.Add(tbl(2))
                    Case "7 notes"
                        Lst7notes.Add(tbl(2))
                    Case "8 notes"
                        Lst8notes.Add(tbl(2))
                    Case "9 notes"
                        Lst9notes.Add(tbl(2))
                End Select
            End While
        End Using ' Fermer
        '
        ' Construction des Titres
        For i = 0 To NbTypeGammes - 1
            LTitre.Add(New CheckBox)
            With LTitre(i)
                .Checked = False
                .Visible = True
                .AutoSize = True
                .Text = UCase(LTypeG(i))
                p.X = (i * 129) + 11
                p.Y = 3
                .Location = p
                .Tag = i
            End With
            SplitContainer1.Panel2.Controls.Add(LTitre(i))
            'TabPage1.Controls.Add(LTitre(i))
            AddHandler LTitre(i).CheckedChanged, AddressOf Ltitre_CheckedChanged
        Next
        '
        ' construction des listes
        ' Gammes ususuelles
        LUsuelles.Clear()
        For i = 0 To LstUsuelles.Count - 1
            LUsuelles.Add(New CheckBox)
            With LUsuelles(i)
                .Checked = False
                .Visible = True
                .AutoSize = True
                .Text = LstUsuelles(i)
                If .Text = "PMin" Then
                    .ForeColor = Color.Maroon
                    .BackColor = Color.LightYellow
                Else
                    .ForeColor = Color.White
                End If
                p.X = 2 '(LTitre(0).Location.X + 13)
                p.Y = 5 + (i * 20) ' LTitre(0).Location.Y + 2)
                .Location = p
                .Tag = i
                Panel1.Controls.Add(LUsuelles(i))
            End With
        Next
        ' Gammes 5 notes
        L5notes.Clear()
        For i = 0 To Lst5notes.Count - 1
            L5notes.Add(New CheckBox)
            With L5notes(i)
                .Checked = False
                .Visible = True
                .AutoSize = True
                .Text = Lst5notes(i)
                If .Text = "PMin1" Then
                    .ForeColor = Color.Maroon
                    .BackColor = Color.LightYellow
                End If
                p.X = 2 '(LTitre(1).Location.X + 13)
                p.Y = 5 + (i * 20) '(LTitre(1).Location.Y + 20) + (i * 20)
                .Location = p
                .Tag = i
                Panel2.Controls.Add(L5notes(i))
            End With
        Next
        ' Gammes 6notes
        L6notes.Clear()
        For i = 0 To Lst6notes.Count - 1
            L6notes.Add(New CheckBox)
            With L6notes(i)
                .Checked = False
                .Visible = True
                .AutoSize = True
                .Text = Lst6notes(i)
                .ForeColor = Color.White
                p.X = 2 'p.X = (LTitre(2).Location.X + 13)
                p.Y = 5 + (i * 20) 'p.Y = (LTitre(2).Location.Y + 20) + (i * 20)
                .Location = p
                .Tag = i
                Panel3.Controls.Add(L6notes(i))
            End With
        Next
        ' Gammes 7notes
        L7notes.Clear()
        For i = 0 To Lst7notes.Count - 1
            L7notes.Add(New CheckBox)
            With L7notes(i)
                .Checked = False
                .Visible = True
                .AutoSize = True
                .Text = Lst7notes(i)
                p.X = 2 'p.X = (LTitre(2).Location.X + 13)
                p.Y = 5 + (i * 20) 'p.Y = (LTitre(2).Location.Y + 20) + (i * 20)
                .Location = p
                .Tag = i
                Panel4.Controls.Add(L7notes(i))
            End With
        Next
        '
        ' Gammes 8notes
        L8notes.Clear()
        For i = 0 To Lst8notes.Count - 1
            L8notes.Add(New CheckBox)
            With L8notes(i)
                .Checked = False
                .Visible = True
                .AutoSize = True
                .Text = Lst8notes(i)
                .ForeColor = Color.White
                p.X = 2 'p.X = (LTitre(2).Location.X + 13)
                p.Y = 5 + (i * 20) 'p.Y = (LTitre(2).Location.Y + 20) + (i * 20)
                .Location = p
                .Tag = i
                Panel5.Controls.Add(L8notes(i))
            End With
        Next
        ' Gammes 9notes
        L9notes.Clear()
        For i = 0 To Lst9notes.Count - 1
            L9notes.Add(New CheckBox)
            With L9notes(i)
                .Checked = False
                .Visible = True
                .AutoSize = True
                .Text = Lst9notes(i)
                .ForeColor = Color.White
                p.X = 2 'p.X = (LTitre(2).Location.X + 13)
                p.Y = 5 + (i * 20) 'p.Y = (LTitre(2).Location.Y + 20) + (i * 20)
                .Location = p
                .Tag = i
                Panel6.Controls.Add(L9notes(i))
            End With
        Next
        '
        ' LIbellés
        ' ********
        If Module1.LangueIHM = "fr" Then
            Me.Text = "Filtre de Gammes"
            Label1.Text = "Recherche"
            Label5.Text = "Choisir une gamme"
            Label7.Text = "Notes de la gamme"
            Button3.Text = "Annuler"
        Else
            Me.Text = "Scales Filter"
            Label1.Text = "Search"
            Label5.Text = "Chosse a scale"
            Label7.Text = "Scale notes"
            Button3.Text = "Cancel"
        End If

        ' Titre du formulaire
        ' *******************
        If Module1.LangueIHM = "fr" Then
            Me.Text = "Gammes jouables sur liste d'accords"
        Else
            Me.Text = "Playable scales on chords list"
        End If

    End Sub
    Sub Ltitre_CheckedChanged(sender As Object, e As EventArgs)
        Dim ind As Integer
        Dim com As CheckBox = sender
        ind = Val(com.Tag)

        Select Case ind
            Case 0
                For Each Ch As CheckBox In LUsuelles
                    If LTitre(ind).Checked Then
                        Ch.Checked = True
                    Else
                        Ch.Checked = False
                    End If
                Next
            Case 1
                For Each Ch As CheckBox In L5notes
                    If LTitre(ind).Checked Then
                        Ch.Checked = True
                    Else
                        Ch.Checked = False
                    End If
                Next
            Case 2
                For Each Ch As CheckBox In L6notes
                    If LTitre(ind).Checked Then
                        Ch.Checked = True
                    Else
                        Ch.Checked = False
                    End If
                Next
            Case 3
                For Each Ch As CheckBox In L7notes
                    If LTitre(ind).Checked Then
                        Ch.Checked = True
                    Else
                        Ch.Checked = False
                    End If
                Next
            Case 4
                For Each Ch As CheckBox In L8notes
                    If LTitre(ind).Checked Then
                        Ch.Checked = True
                    Else
                        Ch.Checked = False
                    End If
                Next
            Case 5
                For Each Ch As CheckBox In L9notes
                    If LTitre(ind).Checked Then
                        Ch.Checked = True
                    Else
                        Ch.Checked = False
                    End If
                Next
        End Select

    End Sub
    Sub test1()
        Dim a As String
        a = oo.ApparteanceG("C", "Maj MinH MinM PMin PMin1 PMin2 PMin3 PMin4")
        TextBox3.Text = "C, Maj MinH MinM PMin PMin1"
        TextBox2.Text = a
    End Sub
    Sub test2() ' gammes de 7 notes avec un seul accord mineur
        Dim a As String

        a = oo.ApparteanceG("E m", "Maj PhrygienMaj Enigmatique SuperLocrien Hongroise1 Hongroise2 Napolitaine Grecque Indienne1 Indienne2 Indienne3 MoyenOrientale Enigmatique")
        TextBox3.Text = "E m, Maj PhrygienMaj Enigmatique SuperLocrien Hongroise1 Hongroise2 Napolitaine Grecque Indienne1 Indienne2 Indienne3 MoyenOrientale Enigmatique"
        TextBox2.Text = a
    End Sub
    Sub test3() ' gammes de 7 notes avec séquence d'accords
        Dim a As String
        a = oo.ApparteanceG("E m7-G 7-C", "PhrygienMaj Lydien_b2b7 SuperLocrien Hongroise1 Hongroise2 Napolitaine Grecque Indienne1 Indienne2 Indienne3 MoyenOrientale Enigmatique")
        TextBox3.Text = "E m7-G 7-C" + "PhrygienMaj Lydien_b2b7 SuperLocrien Hongroise1 Hongroise2 Napolitaine Grecque Indienne1 Indienne2 Indienne3 MoyenOrientale Enigmatique"
        TextBox2.Text = a
    End Sub
    Sub test6() ' gammes de 7 notes avec séquence d'accords
        Dim a As String
        a = oo.ApparteanceG("C", "PMaj1 PMaj2 PMin1 PMin2 PMin3 PMin4 PMin5 PMin6")
        TextBox3.Text = "E m7-G 7-C" + "PhrygienMaj Lydien_b2b7 SuperLocrien Hongroise1 Hongroise2 Napolitaine Grecque Indienne1 Indienne2 Indienne3 MoyenOrientale Enigmatique"
        TextBox2.Text = a
    End Sub
    'Dim gam As New List(Of String) From {"PMaj1", "PMaj2", "PMin1", "PMin2", "PMin3", "PMin4", "PMin5", "PMin6"}
    Sub test4() ' gammes de 5 notes avec un seul accord mineur
        Dim a As String
        a = oo.ApparteanceG("G# m", "PMaj1 PMaj2 PMin1 PMin2 PMin3 PMin4 PMin5 PMin6")
        TextBox3.Text = "G# m, Maj MinH"
        TextBox2.Text = a
    End Sub
    Sub test5() ' gammes de 5 notes avec avec séquence d'accords
        Dim a As String
        a = oo.ApparteanceG("E m7-C", "PMaj1 PMaj2 PMin1 PMin2 PMin3 PMin4 PMin5 PMin6")
        TextBox3.Text = "E m7-C, PMaj1 PMaj2 PMin1 PMin2 PMin3 PMin4 PMin5 PMin6"
        TextBox2.Text = a
    End Sub
    Sub test7() ' gammes de 5 notes avec avec séquence d'accords
        Dim a As String
        a = oo.ApparteanceG("E m7", "PMin")
        TextBox3.Text = "E m7, PMin"
        TextBox2.Text = a
    End Sub
    ' DEBUT 
    ' *****
    ' DETERMINATION DES EQUIVALENCE DE GAMMES
    ' OBJECTIF : Savoir si plusieurs gammes possèdent les mêmes notes
    ' Les méthodes utilisées sont
    ' -  CalcEquival(gam As List(Of String)) As List(Of String) ' Entrée : liste des gammes à comparer- Sortie : Liste des Gammes equivalentes
    ' - _5NotesEquival() --> fait appel à CalcEquival
    ' - _6notesEquival() --> fait appel à CalcEquival
    ' - _7notesEquival() --> fait appel à CalcEquival
    ' - _8notesEquival() --> fait appel à CalcEquival
    ' - - JazzEquival()  --> fait appel à CalcEquival
    ' Il faut créer un bouton qui permet d'appeler une des procédures _xxxxEquival()

    Function CalcEquival(gam As List(Of String)) As List(Of String)
        Dim i As Integer
        Dim b1, b2 As String
        Dim chroma As New List(Of String) From {"C", "C#", "D", "D#", "E", "F", "F#", "G", "G#", "A", "A#", "B"}
        Dim L1 As New List(Of String)
        Dim L2 As New List(Of String)
        Dim gamEquiv As New List(Of String)

        For Each a1 As String In chroma
            For Each g1 As String In gam
                b1 = oo.Det_NotesGammes3(a1 + " " + g1)
                L1 = oo.ChaineEnListe(b1, "-")
                For Each a2 As String In chroma
                    For Each g2 As String In gam
                        b2 = oo.Det_NotesGammes3(a2 + " " + g2)
                        L2 = oo.ChaineEnListe(b2, "-")
                        i = 0
                        For Each n As String In L2
                            If L1.Contains(n) Then
                                i += 1
                            End If
                        Next
                        If i = L1.Count And ((a1 + " " + g1) <> (a2 + " " + g2)) Then gamEquiv.Add(a1 + " " + g1 + " équivalent à " + a2 + " " + g2) ' And (g1 <> g2)
                    Next
                Next
            Next
        Next
        Return gamEquiv
    End Function
    Sub _5NotesEquival()
        'Dim gam As New List(Of String) From {"PMaj1", "PMaj2", "PMaj1", "PMaj2", "PMin1", "PMin2", "PMin3", "PMin4", "PMin5", "PMin6", "PMin2"}
        Dim gam As New List(Of String) From {"PMaj1", "PMaj2", "PMin1", "PMin2", "PMin3", "PMin4", "PMin5", "PMin6"} ' avec supression de Maj1, Maj2, Min2 Résultat = 0 équivalence --> OK
        Dim ListEquiv As New List(Of String)
        ListEquiv = CalcEquival(gam)
    End Sub
    Sub _6notesEquival()
        'Dim gam As New List(Of String) From {"Augmentee", "Ptons", "Blues1", "Blues2", "6NotesMaj1", "6NotesMaj2", "6NotesMaj3", "6NotesMaj2", "6NotesMin1", "6NotesMin2"} et 
        Dim gam As New List(Of String) From {"Augmentee", "Ptons", "Blues1", "6NotesMaj1", "6NotesMaj3", "6NotesMaj2", "6NotesMin1", "6NotesMin2"} ' avec supression de Blues2 et 6NotesMaj2 Résultat = 0 équivalence --> OK (il existe des équivalences entre gamme de même nom : C Augmenté est équivalent à E Augmentée et G# Augmentée et C Pton est équivalent à D Pton, E Pton, F# Pton, G# Pton, A# Pton et C 6NotesMaj4 est équivalent à F# 6NotesMaj4
        Dim ListEquiv As New List(Of String)
        ListEquiv = CalcEquival(gam)
    End Sub
    Sub _7notesEquival()
        'Dim gam As New List(Of String) From {"Ionien", "Ionien2b", "Dorien", "Dorienb5", "Phrygien", "PhrygienMaj", "Lydien", "Lydienb7", "Lydien_b2b7", "Lydien#5",
        ''"MixoLyd", "MixoLyd_b13", "Aeolien", "Locrien", "LocrienMaj", "MinLoc", "SuperLocrien", "Hongroise1", "Hongroise2", "Napolitaine", ""Grecque", "Yidish1", "Yidish2", "Indienne1", "Indienne2",
        '"Indienne3", "Indienne4", "MoyenOrientale1", "MoyenOrientale2", "MoyenOrientale", "MoyenOrientale4", "Leading", "Enigmatique"}
        Dim gam As New List(Of String) From {"PhrygienMaj", "Lydien_b2b7", "SuperLocrien", "Hongroise1", "Hongroise2", "Napolitaine", "Grecque", "Indienne1", "Indienne2",
        "Indienne3", "MoyenOrientale", "Enigmatique"} ' Résultat = 0 équivalence --> OK
        Dim ListEquiv As New List(Of String)
        ListEquiv = CalcEquival(gam)
    End Sub

    Sub _8notesEquival()
        'Dim gam As New List(Of String) From {"Bebop1", "Bebop2", "Bebop3", "Diminuee", "Blues2", "Espagnole", "8Notes"}
        Dim gam As New List(Of String) From {"Bebop1", "Diminuee", "Blues2", "Espagnole"} ' Résultat = 0 équivalence --> OK
        Dim ListEquiv As New List(Of String)
        ListEquiv = CalcEquival(gam)
    End Sub
    Sub JazzEquival()
        'Dim gam As New List(Of String) From {"Ionien", "Ionien2b", "Dorien", "Dorienb5", "Phrygien", "PhrygienMaj", "Lydien", "Lydienb7", "Lydien_b2b7", "Lydien#5",
        '"MixoLyd", "MixoLyd_b13", "Aeolien", "Locrien", "LocrienMaj"}
        Dim gam As New List(Of String) From {"Ionien2b", "Dorienb5", "PhrygienMaj", "Lydien_b2b7"} ' Résultat = 0 équivalence --> OK
        Dim ListEquiv As New List(Of String)
        ListEquiv = CalcEquival(gam)
    End Sub
    ' FIN
    ' ***

    ' EXEMPLE OPERATION SUR LES ENSEMBLES DANS .NET
    ' *********************************************
    ' Distinct	Supprime les valeurs en Double d'une collection.	Non applicable.	Enumerable.Distinct
    ' Except	Retourne la différence ensembliste, à savoir les éléments d'une collection qui n’apparaissent pas dans une seconde collection.	Non applicable.	Enumerable.Except
    ' Intersect	Retourne l'intersection ensembliste, à savoir les éléments qui apparaissent dans chacune des deux collections.	Non applicable.	Enumerable.Intersect
    ' Union	Retourne l'union ensembliste, à savoir les éléments uniques qui apparaissent dans l’une ou l’autre des deux collections.	Non applicable.	Enumerable.Union
    '
    ''' <summary>
    ''' OpéEnsembles : union/intersection de 2 liste (au sens théorie des ensembles)
    ''' </summary>
    ''' 
    Private Sub OpéEnsembles()
        Dim a As String = " "
        Dim b As String = " "
        Dim ints1 As Integer() = {5, 3, 9, 7, 5, 9, 3, 7}
        Dim ints2 As Integer() = {8, 3, 6, 4, 4, 9, 1, 0}

        Dim union As IEnumerable(Of Integer) = ints1.Union(ints2)
        Dim intersection As IEnumerable(Of Integer) = ints1.Intersect(ints2)

        For Each i As Integer In union
            a = a + i.ToString + " "
        Next
        MessageBox.Show(a)

        For Each i As Integer In intersection
            b = b + i.ToString + " "
        Next
        MessageBox.Show(b)
    End Sub
    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        Dim a As String = ""
        If e.KeyChar = vbCr Then
            a = oo.Det_NotesGammes3(Trim(TextBox1.Text))
        End If
        Label3.Text = Trim(a)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim b As String = ""
        Dim g As String
        Dim tbl() As String


        ' liste utilisable de la totalité des gammes
        ' ******************************************
        'b = "Maj" + " " + "MinH" + " " + "MinM" + " " + "PMin" + " " + "PMaj1" + " " + "PMaj2" + " " + "PMin1" + " " + "PMin2" + " " + "PMin3" + " " + "PMin4" + " " + "PMin5" + " " + "PMin6" _
        '    + " " + "Augmentee" + " " + "Ptons" + " " + "Blues1" + " " + "6NotesMaj1" + " " + "6NotesMaj2" + " " + "6NotesMaj3" + " " + "6NotesMin1" + " " + "6NotesMin2" _
        '    + " " + "PhrygienMaj" + " " + "Lydien_b2b7" + " " + "Hongroise1" + " " + "Hongroise2" + " " + "Napolitaine" + " " + "SuperLocrien" + " " + "Grecque" _
        '    + " " + "Indienne1" + " " + "Indienne2" + " " + "Indienne3" + " " + "MoyenOrientale" + " " + "MajH" + " " + "Enigmatique" + " " + "Bebop1" + " " + "Diminuee" + " " + "Blues2" + " " + "Espagnole" + " " + "9notes"

        ' Récolter les gammes choisies
        For Each cc As CheckBox In LUsuelles
            If cc.Checked Then b = b + cc.Text + " "
        Next
        For Each cc As CheckBox In L5notes
            If cc.Checked Then b = b + cc.Text + " "
        Next
        For Each cc As CheckBox In L6notes
            If cc.Checked Then b = b + cc.Text + " "
        Next
        For Each cc As CheckBox In L7notes
            If cc.Checked Then b = b + cc.Text + " "
        Next
        For Each cc As CheckBox In L8notes
            If cc.Checked Then b = b + cc.Text + " "
        Next
        For Each cc As CheckBox In L9notes
            If cc.Checked Then b = b + cc.Text + " "
        Next
        '
        ' Déterminer les gammes jouables sur les accords 
        'a = oo.ApparteanceG("C", "Maj MinH MinM PMin PMin1 PMin2 PMin3 PMin4")
        'a = oo.ApparteanceG("C", "PMin PMin1")
        If Trim(b) <> "" Then
            g = oo.ApparteanceG(Trim(Accords), Trim(b))

            tbl = g.Split("-")
            If tbl(0) <> "" Then
                ComboBox1.Enabled = True
                ComboBox1.Items.Clear()
                For Each gam As String In tbl
                    ComboBox1.Items.Add(gam)
                Next
                ComboBox1.SelectedIndex = 0
            Else
                ComboBox1.Text = ""
                ComboBox1.Items.Clear()
                ComboBox1.Enabled = False
                Label7.Text = ""
                If Module1.LangueIHM = "fr" Then
                    mess = "pas de gamme(s) trouvée(s)"
                    titre = "Avertissement"
                Else
                    mess = "No scales found"
                    titre = "Warning"
                End If
                i = MessageBox.Show(mess, titre, MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Else
                If Module1.LangueIHM = "fr" Then
                mess = "Choisissez au moins une gamme"
                titre = "Avertissement"
            Else
                mess = "Choose at least one scale"
                titre = "Warning"
            End If
            ' 
            i = MessageBox.Show(mess, titre, MessageBoxButtons.OK, MessageBoxIcon.Warning)

        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim a As String = ""
        '
        RetourValue = Trim(ComboBox1.Text)
        Me.Hide()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        RetourValue = ""
        Me.Hide()
    End Sub

    Private Sub TabPage1_Click(sender As Object, e As EventArgs) Handles TabPage1.Click

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If Trim(ComboBox1.Text) <> "" Then
            Label7.Text = oo.Det_NotesGammes3(Trim(ComboBox1.Text))
        End If
    End Sub

    Private Sub FiltreGammes_GotFocus(sender As Object, e As EventArgs) Handles Me.GotFocus

    End Sub

    Private Sub FiltreGammes_VisibleChanged(sender As Object, e As EventArgs) Handles Me.VisibleChanged
        ComboBox1.Items.Clear()
        ComboBox1.Text = ""
        Label7.Text = ""
    End Sub

    Private Sub SplitContainer1_Panel1_Paint(sender As Object, e As PaintEventArgs) Handles SplitContainer1.Panel1.Paint

    End Sub

    Private Sub TabPage2_Click(sender As Object, e As EventArgs) Handles TabPage2.Click

    End Sub
End Class
