

Public Module Harmonie
    ' Table des mouvements : mémorise l'ordre des degrés
    Public TMouvements(0 To 3, 0 To 6) As Integer
    ' 
    Public TSauve_MouvSecondes(0 To 3, 0 To 6) As String
    Public TSauve_MouvSelect(0 To 2, 0 To 20) As Boolean
    Public DegréMarqué As String = ""
    Public LigneMarquée As Integer = -1
    '
    Public Structure AA
        Public Check As Boolean
        Public Visible As Boolean
    End Structure
    Public TSauve_MouvFiltres(0 To 2, 0 To 13) As AA
    Public ListNotesd As New ArrayList
    Public ListNotesb As New ArrayList
    Public ListNotesMajd As New List(Of String)
    Public ListNotesMajd_Inv As New List(Of String)


    '
    Public ListNd As New ArrayList
    Public ListNb As New ArrayList

    Public ListNdLatine As New ArrayList
    Public ListNbLatine As New ArrayList
    '
    Public Clef As String ' Clef pour la partie "ecriture accord", est mis à jour à un seul endroit : choix de la tonalité
    Public ClefCompoGrid As String ' Clef pour Compogrid qui peut contenir des Tonalités de clefs différentes. A mettre à jour à chaque utilisation pour chaque tonalité ou Accord/Gamme relatifs.
    '
    Public TabNotesD(0 To 35) As String
    Public TabNotesB(0 To 35) As String

    '
    Public TabNotesMajD(0 To 35) As String
    Public TabNotesMajB(0 To 35) As String
    ' Tableau de notes
    ' ****************
    Public TabNotes(36) As String ' Pour la artie "écriture accord". Dépend de la valeur de la variable Clef.
    Public TabNotesCompoGrid(36) As String ' Pour la partie CompoGrid.Dépend de la valeur de la variable ClefCompoGrid
    ' Listes des tonalités majeures et mineures
    ' *****************************************
    '
    Public listMaj As New List(Of String)
    Public listMin As New List(Of String)
    '
    Public Dico_MajHorsClef As New Dictionary(Of String, String)
    Public Dico_MinHorsClef As New Dictionary(Of String, String)


    ' TableGlobaleAcc : Table de sauvegarde des tous les accords d'une tonalité donnée.
    ' ********************************************************************************
    ' TableGlobaleAcc contient tous les accords de 3, 4 et 5 notes.
    ' ------------------------------------------------------------
    ' Les éléments de la table sont toujours exprimés par mouvement de SECONDES
    ' *************************************************************************
    ' Indice 1 (0 to 3) : Type d'accord : 3 notes, 7e, 9e ,11e 
    ' Indice 2 : (0 to 2) modes : Majeur = 0 , Mineur H = 1 et Mineur M =2 - C'est uassi le N° de ligne dans TabTons
    ' Indice 3 : (0 to 6) : les Accords
    ' Ce tableau contient tous les accords, de tous les modes d'une tonalité donnée
    ' Pour les 9e et 11e, ce tableau ne contient que les 9e et 11e non filtrés.
    '

    Public TableGlobalAcc(0 To 3, 0 To 2, 0 To 6) As String
    Public TableGlobalAccVoisin(0 To 3, 0 To 5, 0 To 6) As String
    Public TableGlobalAccSubsti(0 To 3, 0 To 7, 0 To 6) As String
    Public CAD_TableGlobalAcc(0 To 3, 0 To 2, 0 To 6) As String
    Public Blues_TableGlobalAcc(0 To 3, 0 To 2, 0 To 6) As String ' 
    '
    ' 
    '
    ' TableCoursAcc : Table de sauvegarde des données affichées dans le tableau
    ' *************************************************************************
    ' TableCoursAcc contient les caractéristiques des accords actuellement affichés.
    ' ----------------------------------------------------------------------------
    ' Les éléments de la table sont toujours exprimés par mouvement de SECONDES
    ' *************************************************************************
    ' Indice 1 : (0 to 2) : ligne du tableau affichée (Modes Maj, MinH et MinM)
    ' Indice 2 : (0 to 6) : colonne du tableau affichée - correspond au degré I, II, III  selon le mouvement courant
    ' Ce tableau contient les accords en cours préents dans le tableau affiché dans l'IHM
    '
    Public Structure AccordTab
        Public TyAcc As String  '3 Notes, 7e, 9e, 11e
        Public Accord As String ' tonique + chiffrage 
        Public Octave As String ' -1, 0, +1
        Public OctaveChoisie As Integer
        Public Renversement() As String ' R1, R2, R3 ...
        Public RenvChoisi As Integer ' index du renversement courant dans Renversement()
        Public Marqué As Boolean
        Public Tonalité As String ' pour onglet "cadences"
        '
        Public EtendreNotes() As String
        Public EtendreChecked() As Boolean
        '
    End Structure
    '

    Public TableCoursAcc(0 To 2, 0 To 6) As AccordTab ' ligne , colonne
    Public TableCoursAccVoisins(0 To 5, 0 To 6) As AccordTab ' ligne , colonne
    Public CAD_TableCoursAcc(0 To 6) As AccordTab
    ' Versions  de sauvegarde des tables précédentes utilisées lors des changement de  langues (fr/en)
    ' ************************************************************************************************
    Public S_TableCoursAcc(0 To 2, 0 To 6) As AccordTab ' ligne , colonne
    Public S_CAD_TableCoursAcc(0 To 6) As AccordTab

    ' Tableau de Positionnement des EVENTH  : Oisition, Tonalités, Accords, Gammes dans chaque mesure de Grid3
    ' ********************************************************************************************************
    Public Structure EventH
        Public Position As String
        Public Marqueur As String
        Public NumAcc As Integer
        Public Accord As String
        Public Tonalité As String ' mis à jour par le comboboxlist1 général (tonalité majeure -> tonalité mineure)
        Public Mode As String
        Public Gamme As String
        Public Degré As Integer
        Public Détails As String
        Public Ligne As Integer
        Public Répet As Byte
        Public NumMagnéto As Integer
        Public Racine As String
        Public Vel As String ' utilisé dans HYperVoicing
    End Structure
    ' Index 1 : N° de mesure
    ' Index 2 : temps
    ' Index 3 : contre temps
    Public TableEventH(0 To nbMesures, 0 To 5, 0 To 4) As EventH

    ' Tableaux des notes calculées
    ' ****************************
    Public Structure ParZone
        Public DébutZ As String
        Public TermeZ As String
        Public Racine As String      ' text dans le combo racine
        Public NoteRacine As Integer ' index dans le combo racine
        Public OctaveRacine As Integer
        Public OctavePlus1 As Boolean
        Public OctaveMoins1 As Boolean
        Public VoixAsso_OctaveMoins1 As String
        Public VoixAsso_OctavePlus1 As String
        Public ComboVoixAInd As Integer
        Public ComboVoixBInd As Integer
        Public CombiVoixInd As Integer
    End Structure
    ' Tableau des notes caculées par CalculAutovoicing pour fonctionnement sans zone
    ' ******************************************************************************
    Public TableNotesAccords(0 To nbMesures, 0 To 5, 0 To 4) As String
    Public THorsZone As ParZone

    ' Tableau des notes caculées par CalculAutovoicing pour fonctionnement avec zoneq
    ' *******************************************************************************
    Public TableNotesAccordsZ(0 To nbMesures, 0 To 5, 0 To 4) As String
    Public TZone(0 To (NbZones)) As ParZone
    Public ZoneCourante As Integer

    ' Valeurs de retour d'un formulaire
    ' *********************************
    Public RetourString As String

    'Public TableAccGrid3(0 To nbMesures - 1, 0 To 11)
    Public Enum Modes
        Rien = 0
        Majeur = 1
        MineurH = 2
        MineurM = 3
        '
        Cadence_Majeure = 4
        Cadence_Mineure = 5
        Cadence_Mixte = 6
        '
    End Enum
    Public C_ As Color = Color.Gold                ' C
    Public G_ As Color = Color.DarkOrange                ' G
    Public D_ As Color = Color.Lavender                      ' D
    Public A_ As Color = Color.GreenYellow               ' A
    Public E_ As Color = Color.Yellow                    ' E
    Public B_ As Color = Color.LightCyan                 ' B
    Public Fd_ As Color = Color.LightGreen               ' F#
    Public Cd_ As Color = Color.Violet '                 ' C#
    Public Gd_ As Color = Color.Magenta                  ' G#
    Public Dd_ As Color = Color.AliceBlue                ' D#
    Public Ad_ As Color = Color.LightSteelBlue           ' A#
    Public F_ As Color = Color.Cornsilk                  ' F
    '
    Public C_lettre As Color = Color.Black                ' C
    Public G_lettre As Color = Color.White                ' G
    Public D_lettre As Color = Color.Black                ' D
    Public A_lettre As Color = Color.Black                ' A
    Public E_lettre As Color = Color.Blue                 ' E
    Public B_lettre As Color = Color.Black                ' B
    Public Fd_lettre As Color = Color.Black               ' F#
    Public Cd_lettre As Color = Color.Black '             ' C#
    Public Gd_lettre As Color = Color.Yellow              ' G#
    Public Dd_lettre As Color = Color.Black               ' D#
    Public Ad_lettre As Color = Color.Black               ' A#
    Public F_lettre As Color = Color.Black                ' F
    '
    Public DicoCouleur As New Dictionary(Of String, Color)
    Public DicoCouleurLettre As New Dictionary(Of String, Color)
    '
    ' Tolaité courante (utilisé seulement par quelques procédures particulières (Maj_MenuTousAccords2 --> Mode --> NoteInterval2)
    ' **************************************************************************************************************************
    Public Tonacours As String


    Sub Maj_ListMajMin()
        listMaj.Add("C#")
        listMaj.Add("F#")
        listMaj.Add("B")
        listMaj.Add("E")
        listMaj.Add("A")
        listMaj.Add("D")
        listMaj.Add("G")
        listMaj.Add("C")
        listMaj.Add("F")
        listMaj.Add("Bb")
        listMaj.Add("Eb")
        listMaj.Add("Ab")
        '
        listMin.Add("A#")
        listMin.Add("D#")
        listMin.Add("G#")
        listMin.Add("C#")
        listMin.Add("F#")
        listMin.Add("B")
        listMin.Add("E")
        listMin.Add("A")
        listMin.Add("D")
        listMin.Add("G")
        listMin.Add("C")
        listMin.Add("F")
        '
        '
        Dico_MajHorsClef("Db") = "C#"
        Dico_MajHorsClef("Gb") = "F#"
        Dico_MajHorsClef("A#") = "Bb"
        Dico_MajHorsClef("D#") = "Eb"
        Dico_MajHorsClef("G#") = "Ab"

        Dico_MinHorsClef("Bb") = "A#"
        Dico_MinHorsClef("Eb") = "D#"
        Dico_MinHorsClef("Ab") = "G#"
        Dico_MinHorsClef("Db") = "C#"
        Dico_MinHorsClef("Gb") = "F#"        '

    End Sub



    Public Sub Maj_DicoCouleur()
        DicoCouleur.Add("C", C_)
        DicoCouleur.Add("C#", Cd_)
        DicoCouleur.Add("D", D_)
        DicoCouleur.Add("D#", Dd_)
        DicoCouleur.Add("E", E_)
        DicoCouleur.Add("F", F_)
        DicoCouleur.Add("F#", Fd_)
        DicoCouleur.Add("G", G_)
        DicoCouleur.Add("G#", Gd_)
        DicoCouleur.Add("A", A_)
        DicoCouleur.Add("A#", Ad_)
        DicoCouleur.Add("B", B_)
        '
        DicoCouleur.Add("Db", Cd_)
        DicoCouleur.Add("Eb", Dd_)
        DicoCouleur.Add("Gb", Fd_)
        DicoCouleur.Add("Ab", Gd_)
        DicoCouleur.Add("Bb", Ad_)
        '
        DicoCouleurLettre.Add("C", C_lettre)
        DicoCouleurLettre.Add("C#", Cd_lettre)
        DicoCouleurLettre.Add("D", D_lettre)
        DicoCouleurLettre.Add("D#", Dd_lettre)
        DicoCouleurLettre.Add("E", E_lettre)
        DicoCouleurLettre.Add("F", F_lettre)
        DicoCouleurLettre.Add("F#", Fd_lettre)
        DicoCouleurLettre.Add("G", G_lettre)
        DicoCouleurLettre.Add("G#", Gd_lettre)
        DicoCouleurLettre.Add("A", A_lettre)
        DicoCouleurLettre.Add("A#", Ad_lettre)
        DicoCouleurLettre.Add("B", B_lettre)
        '
        DicoCouleurLettre.Add("Db", Cd_lettre)
        DicoCouleurLettre.Add("Eb", Dd_lettre)
        DicoCouleurLettre.Add("Gb", Fd_lettre)
        DicoCouleurLettre.Add("Ab", Gd_lettre)
        DicoCouleurLettre.Add("Bb", Ad_lettre)
        '
    End Sub

    Public Sub Maj_ListNotesd()
        Dim i As Integer
        Dim j As Integer
        Dim n As String


        j = -3
        For i = 0 To 10
            j = j + 1
            n = Trim(Str(j))
            If i < 10 Then
                ListNotesd.Add("c" + n)
                ListNotesd.Add("c#" + n)
                ListNotesd.Add("d" + n)
                ListNotesd.Add("d#" + n)
                ListNotesd.Add("e" + n)
                ListNotesd.Add("f" + n)
                ListNotesd.Add("f#" + n)
                ListNotesd.Add("g" + n)
                ListNotesd.Add("g#" + n)
                ListNotesd.Add("a" + n)
                ListNotesd.Add("a#" + n)
                ListNotesd.Add("b" + n)
            Else
                ListNotesd.Add("c" + n)
                ListNotesd.Add("c#" + n)
                ListNotesd.Add("d" + n)
                ListNotesd.Add("d#" + n)
                ListNotesd.Add("e" + n)
                ListNotesd.Add("f" + n)
                ListNotesd.Add("f#" + n)
                ListNotesd.Add("g" + n)
            End If
        Next
    End Sub
    Public Sub Maj_ListNotesb()
        Dim i As Integer
        Dim j As Integer
        Dim n As String

        j = -3
        For i = 0 To 10
            j = j + 1
            n = Trim(Str(j))
            If i < 10 Then
                ListNotesb.Add("c" + n)
                ListNotesb.Add("db" + n)
                ListNotesb.Add("d" + n)
                ListNotesb.Add("eb" + n)
                ListNotesb.Add("e" + n)
                ListNotesb.Add("f" + n)
                ListNotesb.Add("gb" + n)
                ListNotesb.Add("g" + n)
                ListNotesb.Add("ab" + n)
                ListNotesb.Add("a" + n)
                ListNotesb.Add("bb" + n)
                ListNotesb.Add("b" + n)
            Else
                ListNotesb.Add("c" + n)
                ListNotesb.Add("db" + n)
                ListNotesb.Add("d" + n)
                ListNotesb.Add("eb" + n)
                ListNotesb.Add("e" + n)
                ListNotesb.Add("f" + n)
                ListNotesb.Add("gb" + n)
                ListNotesb.Add("g" + n)
            End If
        Next
    End Sub
    Public Sub Maj_ListNotesMajd()

        ListNotesMajd.Add("C")
        ListNotesMajd.Add("C#")
        ListNotesMajd.Add("D")
        ListNotesMajd.Add("D#")
        ListNotesMajd.Add("E")
        ListNotesMajd.Add("F")
        ListNotesMajd.Add("F#")
        ListNotesMajd.Add("G")
        ListNotesMajd.Add("G#")
        ListNotesMajd.Add("A")
        ListNotesMajd.Add("A#")
        ListNotesMajd.Add("B")
        ListNotesMajd.Add("C")
        ListNotesMajd.Add("C#")
        ListNotesMajd.Add("D")
        ListNotesMajd.Add("D#")
        ListNotesMajd.Add("E")
        ListNotesMajd.Add("F")
        ListNotesMajd.Add("F#")
        ListNotesMajd.Add("G")
        ListNotesMajd.Add("G#")
        ListNotesMajd.Add("A")
        ListNotesMajd.Add("A#")
        ListNotesMajd.Add("B")
        ListNotesMajd.Add("C")
        ListNotesMajd.Add("C#")
        ListNotesMajd.Add("D")
        ListNotesMajd.Add("D#")
        ListNotesMajd.Add("E")
        ListNotesMajd.Add("F")
        ListNotesMajd.Add("F#")
        ListNotesMajd.Add("G")
        ListNotesMajd.Add("G#")
        ListNotesMajd.Add("A")
        ListNotesMajd.Add("A#")
        ListNotesMajd.Add("B")

        ListNotesMajd_Inv.Add("B")
        ListNotesMajd_Inv.Add("A#")
        ListNotesMajd_Inv.Add("A")
        ListNotesMajd_Inv.Add("G#")
        ListNotesMajd_Inv.Add("G")
        ListNotesMajd_Inv.Add("F#")
        ListNotesMajd_Inv.Add("F")
        ListNotesMajd_Inv.Add("E")
        ListNotesMajd_Inv.Add("D#")
        ListNotesMajd_Inv.Add("D")
        ListNotesMajd_Inv.Add("C#")
        ListNotesMajd_Inv.Add("C")
        ListNotesMajd_Inv.Add("B")
        ListNotesMajd_Inv.Add("A#")
        ListNotesMajd_Inv.Add("A")
        ListNotesMajd_Inv.Add("G#")
        ListNotesMajd_Inv.Add("G")
        ListNotesMajd_Inv.Add("F#")
        ListNotesMajd_Inv.Add("F")
        ListNotesMajd_Inv.Add("E")
        ListNotesMajd_Inv.Add("D#")
        ListNotesMajd_Inv.Add("D")
        ListNotesMajd_Inv.Add("C#")
        ListNotesMajd_Inv.Add("C")
        ListNotesMajd_Inv.Add("B")
        ListNotesMajd_Inv.Add("A#")
        ListNotesMajd_Inv.Add("A")
        ListNotesMajd_Inv.Add("G#")
        ListNotesMajd_Inv.Add("G")
        ListNotesMajd_Inv.Add("F#")
        ListNotesMajd_Inv.Add("F")
        ListNotesMajd_Inv.Add("E")
        ListNotesMajd_Inv.Add("D#")
        ListNotesMajd_Inv.Add("D")
        ListNotesMajd_Inv.Add("C#")
        ListNotesMajd_Inv.Add("C")

    End Sub
    '***********************
    '* Procédures communes *
    '***********************
    Public Sub Maj_TabNotes_Minus(Clef As String)
        If Trim(Clef) = "#" Then
            TabNotes(0) = "c"
            TabNotes(1) = "c#"
            TabNotes(2) = "d"
            TabNotes(3) = "d#"
            TabNotes(4) = "e"
            TabNotes(5) = "f"
            TabNotes(6) = "f#"
            TabNotes(7) = "g"
            TabNotes(8) = "g#"
            TabNotes(9) = "a"
            TabNotes(10) = "a#"
            TabNotes(11) = "b"
            TabNotes(12) = "c"
            TabNotes(13) = "c#"
            TabNotes(14) = "d"
            TabNotes(15) = "d#"
            TabNotes(16) = "e"
            TabNotes(17) = "f"
            TabNotes(18) = "f#"
            TabNotes(19) = "g"
            TabNotes(20) = "g#"
            TabNotes(21) = "a"
            TabNotes(22) = "a#"
            TabNotes(23) = "b"
            TabNotes(24) = "c"
            TabNotes(25) = "c#"
            TabNotes(26) = "d"
            TabNotes(27) = "d#"
            TabNotes(28) = "e"
            TabNotes(29) = "f"
            TabNotes(30) = "f#"
            TabNotes(31) = "g"
            TabNotes(32) = "g#"
            TabNotes(33) = "a"
            TabNotes(34) = "a#"
            TabNotes(35) = "b"
        Else
            TabNotes(0) = "c"
            TabNotes(1) = "db"
            TabNotes(2) = "d"
            TabNotes(3) = "eb"
            TabNotes(4) = "e"
            TabNotes(5) = "f"
            TabNotes(6) = "gb"
            TabNotes(7) = "g"
            TabNotes(8) = "ab"
            TabNotes(9) = "a"
            TabNotes(10) = "bb"
            TabNotes(11) = "b"
            TabNotes(12) = "c"
            TabNotes(13) = "db"
            TabNotes(14) = "d"
            TabNotes(15) = "eb"
            TabNotes(16) = "e"
            TabNotes(17) = "f"
            TabNotes(18) = "gb"
            TabNotes(19) = "g"
            TabNotes(20) = "ab"
            TabNotes(21) = "a"
            TabNotes(22) = "bb"
            TabNotes(23) = "b"
            TabNotes(24) = "c"
            TabNotes(25) = "db"
            TabNotes(26) = "d"
            TabNotes(27) = "eb"
            TabNotes(28) = "e"
            TabNotes(29) = "f"
            TabNotes(30) = "gb"
            TabNotes(31) = "g"
            TabNotes(32) = "ab"
            TabNotes(33) = "a"
            TabNotes(34) = "bb"
            TabNotes(35) = "b"
        End If

    End Sub
    Public Sub Maj_TabNotes_Majus(Clef As String)
        If Trim(Clef) = "#" Then
            TabNotes(0) = "C"
            TabNotes(1) = "C#"
            TabNotes(2) = "D"
            TabNotes(3) = "D#"
            TabNotes(4) = "E"
            TabNotes(5) = "F"
            TabNotes(6) = "F#"
            TabNotes(7) = "G"
            TabNotes(8) = "G#"
            TabNotes(9) = "A"
            TabNotes(10) = "A#"
            TabNotes(11) = "B"
            TabNotes(12) = "C"
            TabNotes(13) = "C#"
            TabNotes(14) = "D"
            TabNotes(15) = "D#"
            TabNotes(16) = "E"
            TabNotes(17) = "F"
            TabNotes(18) = "F#"
            TabNotes(19) = "G"
            TabNotes(20) = "G#"
            TabNotes(21) = "A"
            TabNotes(22) = "A#"
            TabNotes(23) = "B"
            TabNotes(24) = "C"
            TabNotes(25) = "C#"
            TabNotes(26) = "D"
            TabNotes(27) = "D#"
            TabNotes(28) = "E"
            TabNotes(29) = "F"
            TabNotes(30) = "F#"
            TabNotes(31) = "G"
            TabNotes(32) = "G#"
            TabNotes(33) = "A"
            TabNotes(34) = "A#"
            TabNotes(35) = "B"
        Else
            TabNotes(0) = "C"
            TabNotes(1) = "Db"
            TabNotes(2) = "D"
            TabNotes(3) = "Eb"
            TabNotes(4) = "E"
            TabNotes(5) = "F"
            TabNotes(6) = "Gb"
            TabNotes(7) = "G"
            TabNotes(8) = "Ab"
            TabNotes(9) = "A"
            TabNotes(10) = "Bb"
            TabNotes(11) = "B"
            TabNotes(12) = "C"
            TabNotes(13) = "Db"
            TabNotes(14) = "D"
            TabNotes(15) = "Eb"
            TabNotes(16) = "E"
            TabNotes(17) = "F"
            TabNotes(18) = "Gb"
            TabNotes(19) = "G"
            TabNotes(20) = "Ab"
            TabNotes(21) = "A"
            TabNotes(22) = "Bb"
            TabNotes(23) = "B"
            TabNotes(24) = "C"
            TabNotes(25) = "Db"
            TabNotes(26) = "D"
            TabNotes(27) = "Eb"
            TabNotes(28) = "E"
            TabNotes(29) = "F"
            TabNotes(30) = "Gb"
            TabNotes(31) = "G"
            TabNotes(32) = "Ab"
            TabNotes(33) = "A"
            TabNotes(34) = "Bb"
            TabNotes(35) = "B"
        End If

    End Sub


    Public Sub Maj_TabNotesD_B()
        TabNotesD(0) = "c"
        TabNotesD(1) = "c#"
        TabNotesD(2) = "d"
        TabNotesD(3) = "d#"
        TabNotesD(4) = "e"
        TabNotesD(5) = "f"
        TabNotesD(6) = "f#"
        TabNotesD(7) = "g"
        TabNotesD(8) = "g#"
        TabNotesD(9) = "a"
        TabNotesD(10) = "a#"
        TabNotesD(11) = "b"
        TabNotesD(12) = "c"
        TabNotesD(13) = "c#"
        TabNotesD(14) = "d"
        TabNotesD(15) = "d#"
        TabNotesD(16) = "e"
        TabNotesD(17) = "f"
        TabNotesD(18) = "f#"
        TabNotesD(19) = "g"
        TabNotesD(20) = "g#"
        TabNotesD(21) = "a"
        TabNotesD(22) = "a#"
        TabNotesD(23) = "b"
        TabNotesD(24) = "c"
        TabNotesD(25) = "c#"
        TabNotesD(26) = "d"
        TabNotesD(27) = "d#"
        TabNotesD(28) = "e"
        TabNotesD(29) = "f"
        TabNotesD(30) = "f#"
        TabNotesD(31) = "g"
        TabNotesD(32) = "g#"
        TabNotesD(33) = "a"
        TabNotesD(34) = "a#"
        TabNotesD(35) = "b"


        TabNotesB(0) = "c"
        TabNotesB(1) = "db"
        TabNotesB(2) = "d"
        TabNotesB(3) = "eb"
        TabNotesB(4) = "e"
        TabNotesB(5) = "f"
        TabNotesB(6) = "gb"
        TabNotesB(7) = "g"
        TabNotesB(8) = "ab"
        TabNotesB(9) = "a"
        TabNotesB(10) = "bb"
        TabNotesB(11) = "b"
        TabNotesB(12) = "c"
        TabNotes(13) = "db"
        TabNotesB(14) = "d"
        TabNotesB(15) = "eb"
        TabNotesB(16) = "e"
        TabNotesB(17) = "f"
        TabNotesB(18) = "gb"
        TabNotesB(19) = "g"
        TabNotesB(20) = "ab"
        TabNotesB(21) = "a"
        TabNotesB(22) = "bb"
        TabNotesB(23) = "b"
        TabNotesB(24) = "c"
        TabNotesB(25) = "db"
        TabNotesB(26) = "d"
        TabNotesB(27) = "eb"
        TabNotesB(28) = "e"
        TabNotesB(29) = "f"
        TabNotesB(30) = "gb"
        TabNotesB(31) = "g"
        TabNotesB(32) = "ab"
        TabNotesB(33) = "a"
        TabNotesB(34) = "bb"
        TabNotesB(35) = "b"
    End Sub
    Public Sub Maj_TabNotesMajD_B()
        TabNotesD(0) = "C"
        TabNotesD(1) = "C#"
        TabNotesD(2) = "D"
        TabNotesD(3) = "D#"
        TabNotesD(4) = "E"
        TabNotesD(5) = "F"
        TabNotesD(6) = "F#"
        TabNotesD(7) = "G"
        TabNotesD(8) = "G#"
        TabNotesD(9) = "A"
        TabNotesD(10) = "A#"
        TabNotesD(11) = "B"
        TabNotesD(12) = "C"
        TabNotesD(13) = "C#"
        TabNotesD(14) = "D"
        TabNotesD(15) = "D#"
        TabNotesD(16) = "E"
        TabNotesD(17) = "F"
        TabNotesD(18) = "F#"
        TabNotesD(19) = "G"
        TabNotesD(20) = "G#"
        TabNotesD(21) = "A"
        TabNotesD(22) = "A#"
        TabNotesD(23) = "B"
        TabNotesD(24) = "C"
        TabNotesD(25) = "C#"
        TabNotesD(26) = "D"
        TabNotesD(27) = "D#"
        TabNotesD(28) = "E"
        TabNotesD(29) = "F"
        TabNotesD(30) = "F#"
        TabNotesD(31) = "G"
        TabNotesD(32) = "G#"
        TabNotesD(33) = "A"
        TabNotesD(34) = "A#"
        TabNotesD(35) = "B"


        TabNotesB(0) = "C"
        TabNotesB(1) = "Db"
        TabNotesB(2) = "D"
        TabNotesB(3) = "Eb"
        TabNotesB(4) = "E"
        TabNotesB(5) = "F"
        TabNotesB(6) = "Gb"
        TabNotesB(7) = "G"
        TabNotesB(8) = "Ab"
        TabNotesB(9) = "A"
        TabNotesB(10) = "Bb"
        TabNotesB(11) = "B"
        TabNotesB(12) = "C"
        TabNotes(13) = "Db"
        TabNotesB(14) = "D"
        TabNotesB(15) = "Eb"
        TabNotesB(16) = "E"
        TabNotesB(17) = "F"
        TabNotesB(18) = "Gb"
        TabNotesB(19) = "G"
        TabNotesB(20) = "Ab"
        TabNotesB(21) = "A"
        TabNotesB(22) = "Bb"
        TabNotesB(23) = "B"
        TabNotesB(24) = "C"
        TabNotesB(25) = "Db"
        TabNotesB(26) = "D"
        TabNotesB(27) = "Eb"
        TabNotesB(28) = "E"
        TabNotesB(29) = "F"
        TabNotesB(30) = "Gb"
        TabNotesB(31) = "G"
        TabNotesB(32) = "Ab"
        TabNotesB(33) = "A"
        TabNotesB(34) = "Bb"
        TabNotesB(35) = "B"
    End Sub



    Public Sub Maj_TabNotes_Majus()
        TabNotes(0) = "C"
        TabNotes(1) = "C#"
        TabNotes(2) = "D"
        TabNotes(3) = "D#"
        TabNotes(4) = "E"
        TabNotes(5) = "F"
        TabNotes(6) = "F#"
        TabNotes(7) = "G"
        TabNotes(8) = "G#"
        TabNotes(9) = "A"
        TabNotes(10) = "A#"
        TabNotes(11) = "B"
        TabNotes(12) = "C"
        TabNotes(13) = "C#"
        TabNotes(14) = "D"
        TabNotes(15) = "D#"
        TabNotes(16) = "E"
        TabNotes(17) = "F"
        TabNotes(18) = "F#"
        TabNotes(19) = "G"
        TabNotes(20) = "G#"
        TabNotes(21) = "A"
        TabNotes(22) = "A#"
        TabNotes(23) = "B"
        TabNotes(24) = "C"
        TabNotes(25) = "C#"
        TabNotes(26) = "D"
        TabNotes(27) = "D#"
        TabNotes(28) = "E"
        TabNotes(29) = "F"
        TabNotes(30) = "F#"
        TabNotes(31) = "G"
        TabNotes(32) = "G#"
        TabNotes(33) = "A"
        TabNotes(34) = "A#"
        TabNotes(35) = "B"
    End Sub
    Public Sub Maj_TabNotes_Majus_b()
        TabNotes(0) = "C"
        TabNotes(1) = "Db"
        TabNotes(2) = "D"
        TabNotes(3) = "Eb"
        TabNotes(4) = "E"
        TabNotes(5) = "F"
        TabNotes(6) = "Gb"
        TabNotes(7) = "G"
        TabNotes(8) = "Ab"
        TabNotes(9) = "A"
        TabNotes(10) = "Bb"
        TabNotes(11) = "B"
        TabNotes(12) = "C"
        TabNotes(13) = "Db"
        TabNotes(14) = "D"
        TabNotes(15) = "Eb"
        TabNotes(16) = "E"
        TabNotes(17) = "F"
        TabNotes(18) = "Gb"
        TabNotes(19) = "G"
        TabNotes(20) = "Ab"
        TabNotes(21) = "A"
        TabNotes(22) = "Bb"
        TabNotes(23) = "B"
        TabNotes(24) = "C"
        TabNotes(25) = "Db"
        TabNotes(26) = "D"
        TabNotes(27) = "Eb"
        TabNotes(28) = "E"
        TabNotes(29) = "F"
        TabNotes(30) = "Gb"
        TabNotes(31) = "G"
        TabNotes(32) = "Ab"
        TabNotes(33) = "A"
        TabNotes(34) = "Bb"
        TabNotes(35) = "B"
    End Sub
    Public Sub Maj_ListN()
        ListNd.Add("c")
        ListNd.Add("c#")
        ListNd.Add("d")
        ListNd.Add("d#")
        ListNd.Add("e")
        ListNd.Add("f")
        ListNd.Add("f#")
        ListNd.Add("g")
        ListNd.Add("g#")
        ListNd.Add("a")
        ListNd.Add("a#")
        ListNd.Add("b")

        ListNb.Add("c")
        ListNb.Add("db")
        ListNb.Add("d")
        ListNb.Add("eb")
        ListNb.Add("e")
        ListNb.Add("f")
        ListNb.Add("gb")
        ListNb.Add("g")
        ListNb.Add("ab")
        ListNb.Add("a")
        ListNb.Add("bb")
        ListNb.Add("b")
    End Sub
    Public Sub Maj_ListNLatine()
        ListNdLatine.Add("Do")
        ListNdLatine.Add("Do#")
        ListNdLatine.Add("Ré")
        ListNdLatine.Add("Ré#")
        ListNdLatine.Add("Mi")
        ListNdLatine.Add("Fa")
        ListNdLatine.Add("Fa#")
        ListNdLatine.Add("Sol")
        ListNdLatine.Add("Sol#")
        ListNdLatine.Add("La")
        ListNdLatine.Add("La#")
        ListNdLatine.Add("Si")

        ListNbLatine.Add("Do")
        ListNbLatine.Add("Réb")
        ListNbLatine.Add("Ré")
        ListNbLatine.Add("Mib")
        ListNbLatine.Add("Mi")
        ListNbLatine.Add("Fa")
        ListNbLatine.Add("Solb")
        ListNbLatine.Add("Sol")
        ListNbLatine.Add("Lab")
        ListNbLatine.Add("La")
        ListNbLatine.Add("Sib")
        ListNbLatine.Add("Si")
    End Sub
    Public Sub Maj_TMouvements()
        ' cycle des secondes
        TMouvements(0, 0) = 1
        TMouvements(0, 1) = 2
        TMouvements(0, 2) = 3
        TMouvements(0, 3) = 4
        TMouvements(0, 4) = 5
        TMouvements(0, 5) = 6
        TMouvements(0, 6) = 7
        'cycle des tierces
        TMouvements(1, 0) = 3
        TMouvements(1, 1) = 5
        TMouvements(1, 2) = 7
        TMouvements(1, 3) = 2
        TMouvements(1, 4) = 4
        TMouvements(1, 5) = 6
        TMouvements(1, 6) = 1
        'cycle des quartes
        TMouvements(2, 0) = 4
        TMouvements(2, 1) = 7
        TMouvements(2, 2) = 3
        TMouvements(2, 3) = 6
        TMouvements(2, 4) = 2
        TMouvements(2, 5) = 5
        TMouvements(2, 6) = 1
        '
        'cycle des quintes
        TMouvements(3, 0) = 5
        TMouvements(3, 1) = 2
        TMouvements(3, 2) = 6
        TMouvements(3, 3) = 3
        TMouvements(3, 4) = 7
        TMouvements(3, 5) = 4
        TMouvements(3, 6) = 1

    End Sub
    Public Function TradAcc_LatAngl(acc As String) ' traduire un accord du latin en anglais, si l'accord arrive en anglais, il est restitué tel quel
        Dim a As String
        Dim tbl() As String

        TradAcc_LatAngl = Trim(acc)
        '
        'If Langue = "fr" Then
        tbl = Split(Trim(acc), " ")

            Select Case Trim(tbl(0))
                Case "Do", "C"
                    a = "C"
                Case "Do#", "C#"
                    a = "C#"
                Case "Ré"
                    a = "D"
                Case "Ré#"
                    a = "D#"
                Case "Mi"
                    a = "E"
                Case "Fa"
                    a = "F"
                Case "Fa#"
                    a = "F#"
                Case "Sol"
                    a = "G"
                Case "Sol#"
                    a = "G#"
                Case "La"
                    a = "A"
                Case "La#"
                    a = "A#"
                Case "Si"
                    a = "B"
                Case "Réb"
                    a = "Db"
                Case "Mib"
                    a = "Eb"
                Case "Solb"
                    a = "Gb"
                Case "Lab"
                    a = "Ab"
                Case "Sib"
                    a = "Bb"
                Case Else '
                    a = Trim(tbl(0)) ' si non trouvé ça signifie qu'on était déjà en Aglais donc on garde la note
            End Select

            If UBound(tbl) > 0 Then
                a = a + " " + tbl(1)
            End If
            TradAcc_LatAngl = a
        'End If
    End Function
    Public Function TradNote_LatMinMaj(note As String) ' 
        Dim a As String


        TradNote_LatMinMaj = Trim(note)
        '
        Select Case Trim(note)
            Case "do"
                a = "Do"
            Case "do#"
                a = "Do#"
            Case "ré"
                a = "Ré"
            Case "ré#"
                a = "Ré#"
            Case "mi"
                a = "Mi"
            Case "fa"
                a = "Fa"
            Case "fa#"
                a = "Fa#"
            Case "sol"
                a = "Sol"
            Case "sol#"
                a = "Sol#"
            Case "la"
                a = "La"
            Case "la#"
                a = "La#"
            Case "si"
                a = "Si"
            Case "réb"
                a = "Réb"
            Case "mib"
                a = "Mib"
            Case "solb"
                a = "Solb"
            Case "lab"
                a = "Lab"
            Case "sib"
                a = "Sib"
            Case Else '
                a = "Sib"
        End Select
        '
        TradNote_LatMinMaj = a
    End Function


    Public Function TradAcc_LatMinMaj(acc As String) ' 
        Dim a As String
        Dim tbl() As String

        TradAcc_LatMinMaj = Trim(acc)
        '
        If Langue = "fr" Then
            tbl = Split(Trim(acc), " ")

            Select Case Trim(tbl(0))
                Case "do"
                    a = "Do"
                Case "do#"
                    a = "Do#"
                Case "ré"
                    a = "Ré"
                Case "ré#"
                    a = "Ré#"
                Case "mi"
                    a = "Mi"
                Case "fa"
                    a = "Fa"
                Case "fa#"
                    a = "Fa#"
                Case "sol"
                    a = "Sol"
                Case "sol#"
                    a = "Sol#"
                Case "la"
                    a = "La"
                Case "la#"
                    a = "La#"
                Case "si"
                    a = "Si"
                Case "réb"
                    a = "Réb"
                Case "mib"
                    a = "Mib"
                Case "solb"
                    a = "Solb"
                Case "lab"
                    a = "Lab"
                Case "sib"
                    a = "Sib"
                Case Else '
                    a = "Sib"
            End Select

            If UBound(tbl) > 0 Then
                a = a + " " + tbl(1)
            End If
            TradAcc_LatMinMaj = a
        End If
    End Function
    Public Function TradAcc_AnglMinMaj(acc As String) ' traduire un accord du latin en anglais
        Dim a As String
        Dim tbl() As String

        TradAcc_AnglMinMaj = Trim(acc)
        '
        If Langue = "en" Then
            tbl = Split(Trim(acc), " ")

            Select Case Trim(tbl(0))
                Case "c"
                    a = "C"
                Case "c#"
                    a = "C#"
                Case "d"
                    a = "D"
                Case "d#"
                    a = "D#"
                Case "e"
                    a = "E"
                Case "f"
                    a = "F"
                Case "f#"
                    a = "F#"
                Case "g"
                    a = "G"
                Case "g#"
                    a = "G#"
                Case "a"
                    a = "A"
                Case "a#"
                    a = "A#"
                Case "b"
                    a = "B"
                Case "db"
                    a = "Db"
                Case "eb"
                    a = "Eb"
                Case "gb"
                    a = "Gb"
                Case "ab"
                    a = "Ab"
                Case "bb"
                    a = "Bb"
                Case Else '
                    a = "Bb"
            End Select

            If UBound(tbl) > 0 Then
                a = a + " " + tbl(1)
            End If
            TradAcc_AnglMinMaj = a
        End If
    End Function
    Public Function TradNote_AnglMinMaj(note As String) ' traduire un accord du latin en anglais
        Dim a As String

        TradNote_AnglMinMaj = Trim(note)
        '
        Select Case Trim(note)
                Case "c"
                    a = "C"
                Case "do#"
                    a = "Do#"
                Case "d"
                    a = "D"
                Case "d#"
                    a = "D#"
                Case "e"
                    a = "E"
                Case "f"
                    a = "F"
                Case "f#"
                    a = "F#"
                Case "g"
                    a = "G"
                Case "g#"
                    a = "G#"
                Case "a"
                    a = "A"
                Case "a#"
                    a = "A#"
                Case "b"
                    a = "B"
                Case "db"
                    a = "Db"
                Case "eb"
                    a = "Eb"
                Case "gb"
                    a = "Gb"
                Case "ab"
                a = "Ab"
            Case "bb"
                a = "Bb"
            Case Else '"
                a = note ' cas où la note anglaise est déjà en majuscule
        End Select
        TradNote_AnglMinMaj = a
    End Function
    Public Function TradAcc_LatAngl2(acc As String) ' traduire un accord du latin en anglais
        Dim a As String
        Dim tbl() As String

        TradAcc_LatAngl2 = Trim(acc)
        '
        tbl = Split(Trim(acc), " ")

        Select Case Trim(tbl(0))
            Case "Do", "C"
                a = "C"
            Case "Do#", "C#"
                a = "C#"
            Case "Ré"
                a = "D"
            Case "Ré#"
                a = "D#"
            Case "Mi"
                a = "E"
            Case "Fa"
                a = "F"
            Case "Fa#"
                a = "F#"
            Case "Sol"
                a = "G"
            Case "Sol#"
                a = "G#"
            Case "La"
                a = "A"
            Case "La#"
                a = "A#"
            Case "Si"
                a = "B"
            Case "Réb"
                a = "Db"
            Case "Mib"
                a = "Eb"
            Case "Solb"
                a = "Gb"
            Case "Lab"
                a = "Ab"
            Case "Sib"
                a = "Bb"
            Case Else '"Sib"
                a = "Bb"
        End Select
        If UBound(tbl) > 0 Then
            a = a + " " + tbl(1)
        End If
        TradAcc_LatAngl2 = a
    End Function
    Public Function TradAcc_AnglLat(acc As String) ' traduire un accord de l'anglais en latin
        Dim a As String
        Dim tbl() As String

        TradAcc_AnglLat = Trim(acc)

        If Langue = "fr" Then
            tbl = Split(Trim(acc), " ")

            Select Case Trim(tbl(0))
                Case "C", "Do"
                    a = "Do"
                Case "C#"
                    a = "Do#"
                Case "D"
                    a = "Ré"
                Case "D#"
                    a = "Ré#"
                Case "E"
                    a = "Mi"
                Case "F"
                    a = "Fa"
                Case "F#"
                    a = "Fa#"
                Case "G"
                    a = "Sol"
                Case "G#"
                    a = "Sol#"
                Case "A"
                    a = "La"
                Case "A#"
                    a = "La#"
                Case "B"
                    a = "Si"
                Case "Db"
                    a = "Réb"
                Case "Eb"
                    a = "Mib"
                Case "Gb"
                    a = "Solb"
                Case "Ab"
                    a = "Lab"
                Case "Bb"
                    a = "Sib"
                Case ""
                    a = ""
                Case Else '
                    a = "Sib"
            End Select
            TradAcc_AnglLat = Trim(Trim(a))
            If UBound(tbl) > 0 Then
                TradAcc_AnglLat = Trim(Trim(a) + " " + Trim(tbl(1)))
            End If
        End If
    End Function
    Public Function TradNote_AnglLatMaj(acc As String) ' traduire un accord du latin en anglais
        Dim a As String

        TradNote_AnglLatMaj = Trim(acc)

        If Langue = "fr" Then
            Select Case Trim(acc)
                Case "c", "C"
                    a = "Do"
                Case "c#", "C#"
                    a = "Do#"
                Case "d", "D"
                    a = "Ré"
                Case "d#", "D#"
                    a = "Ré#"
                Case "e", "E"
                    a = "Mi"
                Case "f", "F"
                    a = "Fa"
                Case "f#", "F#"
                    a = "Fa#"
                Case "g", "G"
                    a = "Sol"
                Case "g#", "G#"
                    a = "Sol#"
                Case "a", "A"
                    a = "La"
                Case "a#", "A#"
                    a = "La#"
                Case "b", "B"
                    a = "Si"
                Case "db", "Db"
                    a = "Réb"
                Case "eb", "Eb"
                    a = "Mib"
                Case "gb", "Gb"
                    a = "Solb"
                Case "ab", "Ab"
                    a = "Lab"
                Case Else '"Bb"
                    a = "Sib"
            End Select
            '
            TradNote_AnglLatMaj = a
        End If
    End Function
    Public Function TradAcc_AnglLat2(acc As String) ' traduire un accord de l'anglais en latin
        Dim a As String
        Dim tbl() As String

        TradAcc_AnglLat2 = Trim(acc)

        tbl = Split(Trim(acc), " ")

        Select Case Trim(tbl(0))
            Case "C", "Do"
                a = "Do"
            Case "C#"
                a = "Do#"
            Case "D"
                a = "Ré"
            Case "D#"
                a = "Ré#"
            Case "E"
                a = "Mi"
            Case "F"
                a = "Fa"
            Case "F#"
                a = "Fa#"
            Case "G"
                a = "Sol"
            Case "G#"
                a = "Sol#"
            Case "A"
                a = "La"
            Case "A#"
                a = "La#"
            Case "B"
                a = "Si"
            Case "Db"
                a = "Réb"
            Case "Eb"
                a = "Mib"
            Case "Gb"
                a = "Solb"
            Case "Ab"
                a = "Lab"
            Case "Bb"
                a = "Sib"
            Case Else '
                a = "Sib"
        End Select

        If UBound(tbl) > 0 Then
            a = a + " " + tbl(1)
        End If
        TradAcc_AnglLat2 = a
    End Function

    Public Function TradNoteMin_AnglLat(NoteAnglMin As String) ' traduire un accord de l'anglais en latin
        Dim a As String

        Select Case Trim(NoteAnglMin)
            Case "c"
                a = "do"
            Case "c#"
                a = "do#"
            Case "d"
                a = "ré"
            Case "d#"
                a = "ré#"
            Case "e"
                a = "mi"
            Case "f"
                a = "fa"
            Case "f#"
                a = "fa#"
            Case "g"
                a = "sol"
            Case "g#"
                a = "sol#"
            Case "a"
                a = "la"
            Case "a#"
                a = "la#"
            Case "b"
                a = "si"
            Case "db"
                a = "réb"
            Case "eb"
                a = "mib"
            Case "gb"
                a = "solb"
            Case "ab"
                a = "lab"
            Case "bb"
                a = "sib"
            Case Else '"Sib"
                a = "sib"
        End Select

        TradNoteMin_AnglLat = Trim(a)
    End Function
    Public Function TradNoteMin_LatAngl(NoteAnglMin As String) ' traduire un accord du latin en anglais
        Dim a As String

        Select Case Trim(NoteAnglMin)
            Case "do"
                a = "c"
            Case "do#"
                a = "c#"
            Case "ré"
                a = "d"
            Case "ré#"
                a = "d#"
            Case "mi"
                a = "e"
            Case "fa"
                a = "f"
            Case "fa#"
                a = "f#"
            Case "sol"
                a = "g"
            Case "sol#"
                a = "g#"
            Case "la"
                a = "a"
            Case "la#"
                a = "a#"
            Case "si"
                a = "b"
            Case "réb"
                a = "db"
            Case "mib"
                a = "eb"
            Case "solb"
                a = "gb"
            Case "lab"
                a = "ab"
            Case "sib"
                a = "bb"
            Case Else '"bb"
                a = "bb"
        End Select

        TradNoteMin_LatAngl = Trim(a)
    End Function
    Public Function Trad_BemDiesNote(note As String) As String
        Select Case note
            Case "db"
                Return "c#"
            Case "eb"
                Return "d#"
            Case "gb"
                Return "f#"
            Case "ab"
                Return "g#"
            Case "bb"
                Return "a#"
            Case Else
                Return note
        End Select
    End Function
    Public Function Trad_DiesBemNote(note As String) As String
        Select Case note
            Case "c#"
                Return "db"
            Case "d#"
                Return "eb"
            Case "f#"
                Return "gb"
            Case "g#"
                Return "ab"
            Case "a#"
                Return "bb"
            Case Else
                Return note
        End Select

    End Function
    Public Function TradNoteMaj_LatAngl(NoteLatMaj As String) ' traduire un accord du latin en anglais
        Dim a As String
        '
        Select Case Trim(NoteLatMaj)
            Case "Do"
                a = "C"
            Case "Do#"
                a = "C#"
            Case "Ré"
                a = "D"
            Case "Ré#"
                a = "D#"
            Case "Mi"
                a = "E"
            Case "Fa"
                a = "F"
            Case "Fa#"
                a = "F#"
            Case "Sol"
                a = "G"
            Case "Sol#"
                a = "G#"
            Case "La"
                a = "A"
            Case "La#"
                a = "A#"
            Case "Si"
                a = "B"
            Case "Réb"
                a = "Db"
            Case "Mib"
                a = "Eb"
            Case "Solb"
                a = "Gb"
            Case "Lab"
                a = "Ab"
            Case "Sib"
                a = "Bb"
            Case Else '
                a = "Bb"
        End Select
        '
        TradNoteMaj_LatAngl = Trim(a)
    End Function




    Public Function TradNoteEnMaj(Note As String) As String 'traduit la 1ere lettre d'une note en majuscule
        Dim a As String
        Dim b As String

        TradNoteEnMaj = UCase(Note)
        If Len(Note) > 1 Then
            a = UCase(Mid(Note, 1, 1))
            b = Mid(Note, 2, Len(Note) - 1)
            TradNoteEnMaj = a + b
        End If
    End Function

    Public Function TradNoteAngl(signe As String, note As String) As String ' pour les sources Progressions et Tonalités
        Dim i As Integer

        TradNoteAngl = note
        If Langue = "fr" Then
            ' déterminer le signe de la note et son index
            ' *******************************************
            If signe = "b" Then
                i = ListNbLatine.IndexOf(note)
            Else
                i = ListNdLatine.IndexOf(note)
            End If
            '
            ' Détermination de la note latine
            ' *******************************
            If signe = "b" Then
                TradNoteAngl = ListNb(i)
            Else
                TradNoteAngl = ListNd(i)
            End If
        End If
        '
    End Function
    Public Function TradNoteAngl2(note As String) As String
        TradNoteAngl2 = note
        Select Case note
            Case "Do", "do"
                TradNoteAngl2 = "C"
            Case "Do#", "do#"
                TradNoteAngl2 = "C#"
            Case "Ré", "ré"
                TradNoteAngl2 = "D"
            Case "Ré#", "ré#"
                TradNoteAngl2 = "D#"
            Case "Mi", "mi"
                TradNoteAngl2 = "E"
            Case "Fa", "fa"
                TradNoteAngl2 = "F"
            Case "Fa#", "fa#"
                TradNoteAngl2 = "F#"
            Case "Sol", "sol"
                TradNoteAngl2 = "G"
            Case "Sol#", "sol#"
                TradNoteAngl2 = "G#"
            Case "La", "la"
                TradNoteAngl2 = "A"
            Case "La#", "la#"
                TradNoteAngl2 = "A#"
            Case "Si", "si"
                TradNoteAngl2 = "B"
            Case "Réb", "réb"
                TradNoteAngl2 = "Db"
            Case "Mib", "mib"
                TradNoteAngl2 = "Eb"
            Case "Solb", "solb"
                TradNoteAngl2 = "Gb"
            Case "Lab"
                TradNoteAngl2 = "Ab"
            Case "Sib", "sib"
                TradNoteAngl2 = "Bb"
            Case Else
                TradNoteAngl2 = Trim(note) ' cas où la note est déjà en Anglais
        End Select
    End Function
    Public Function TradNoteLat(note As String) As String 'pour les sources Progressions et Tonalités
        Dim i As Integer
        Dim signe As String

        TradNoteLat = note
        If Langue = "fr" Then
            ' déterminer le signe de la note et son index
            ' *******************************************
            note = LCase(note)
            'note = Trad_NoteEnDMin(Trim(note))
            'signe = "#"
            signe = Det_SigneNote(note)
            If signe = "#" Then
                i = ListNd.IndexOf(note)
            Else
                i = ListNb.IndexOf(note)
            End If
            '
            ' Détermination de la note latine
            ' *******************************
            If i <> -1 Then
                If signe = "b" Then
                    TradNoteLat = ListNbLatine(i)
                Else
                    TradNoteLat = ListNdLatine(i)
                End If
            End If
        End If
        '
    End Function
    Function Det_SigneNote(note) As String
        Dim a As String
        Det_SigneNote = "#"
        If Len(note) > 1 Then
            a = Microsoft.VisualBasic.Right(note, 1)
            Select Case a
                Case "#"
                    Det_SigneNote = "#"
                Case "b"
                    Det_SigneNote = "b"
                Case Else
                    Det_SigneNote = "#"
            End Select
        End If
    End Function



    Public Function TradEventHLat(EventH As String) As String
        Dim tbl() As String
        Dim Note As String
        '
        TradEventHLat = EventH
        If Langue = "fr" Then
            tbl = Split(EventH, " ")
            Note = TradNoteLat(Trim(tbl(0)))
            TradEventHLat = Trim(Note)
            If UBound(tbl) > 0 Then
                TradEventHLat = Trim(TradEventHLat) + " " + Trim(tbl(1))
            End If
        End If
    End Function
    Public Function TrouverNoteDansTabNotes(note As String) As Integer
        Dim i As Integer
        '
        TrouverNoteDansTabNotes = 0
        For i = 0 To (TabNotes.Length - 1)
            If note = TabNotes(i) Then
                TrouverNoteDansTabNotes = i
                Exit For
            End If
        Next i
    End Function

    Public Function Mode(Tonique As String, Type_Mode As String, Typ As Integer) As String 'TYpe_Mode = Maj, MinH, MinM Typ= 3 -> 3note,Type=4 -> Accord7,Typ=5->Accord9,Typ=6 -> Accord11
        Dim ton As String

        'ton = LCase(Tonique)
        ton = Trim(Tonique)
        Mode = ""

        Select Case Type_Mode
            Case "Maj"
                Select Case Typ
                    Case 3
                        Mode =
                            Tonique + "-" _
                            + UCaseBémol(NoteInterval2(ton, "9")) + " m" + "-" _
                            + UCaseBémol(NoteInterval2(ton, "3")) + " m" + "-" _
                            + UCaseBémol(NoteInterval2(ton, "4")) + "-" _
                            + UCaseBémol(NoteInterval2(ton, "5")) + "-" _
                            + UCaseBémol(NoteInterval2(ton, "6")) + " m" + "-" _
                            + UCaseBémol(NoteInterval2(ton, "M7")) + " mb5"
                    Case 4 ' 7

                        Mode =
                            Tonique + " M7" + "-" _
                            + UCaseBémol(NoteInterval2(ton, "9")) + " m7" + "-" _
                            + UCaseBémol(NoteInterval2(ton, "3")) + " m7" + "-" _
                            + UCaseBémol(NoteInterval2(ton, "4")) + " M7" + "-" _
                            + UCaseBémol(NoteInterval2(ton, "5")) + " 7" + "-" _
                            + UCaseBémol(NoteInterval2(ton, "6")) + " m7" + "-" _
                            + UCaseBémol(NoteInterval2(ton, "M7")) + " m7b5"
                    Case 5 ' 9

                        Mode =
                            Tonique + " M7(9)" + "-" _
                            + UCaseBémol(NoteInterval2(ton, "9")) + " m7(9)" + "-" _
                            + "___" + "-" _
                            + UCaseBémol(NoteInterval2(ton, "4")) + " M7(9)" + "-" _
                            + UCaseBémol(NoteInterval2(ton, "5")) + " 7(9)" + "-" _
                            + UCaseBémol(NoteInterval2(ton, "6")) + " m7(9)" + "-" _
                            + "___" + "-"
                    Case 6 ' 11

                        Mode =
                            Tonique + " 11" + "-" _
                            + UCaseBémol(NoteInterval2(ton, "9")) + " m11" + "-" _
                            + UCaseBémol(NoteInterval2(ton, "3")) + " m11" + "-" _
                            + UCaseBémol(NoteInterval2(ton, "4")) + " M7(11#)" + "-" _
                            + UCaseBémol(NoteInterval2(ton, "5")) + " 7(11)" + "-" _
                            + UCaseBémol(NoteInterval2(ton, "6")) + " m11" + "-" _
                            + "___"
                End Select
            Case "MinH"
                Select Case Typ
                    Case 3
                        Mode =
                            Tonique + " m" + "-" _
                            + UCaseBémol(NoteInterval2(ton, "9")) + " mb5" + "-" _
                            + UCaseBémol(NoteInterval2(ton, "3m")) + " 5#" + "-" _
                            + UCaseBémol(NoteInterval2(ton, "4")) + " m" + "-" _
                            + UCaseBémol(NoteInterval2(ton, "5")) + "-" _
                            + UCaseBémol(NoteInterval2(ton, "5#")) + "-" _
                            + UCaseBémol(NoteInterval2(ton, "M7")) + " mb5"
                    Case 4 ' 7
                        Mode =
                            Tonique + " mM7" + "-" _
                            + UCaseBémol(NoteInterval2(ton, "9")) + " m7b5" + "-" _
                            + UCaseBémol(NoteInterval2(ton, "3m")) + " M75#" + "-" _
                            + UCaseBémol(NoteInterval2(ton, "4")) + " m7" + "-" _
                            + UCaseBémol(NoteInterval2(ton, "5")) + " 7" + "-" _
                            + UCaseBémol(NoteInterval2(ton, "5#")) + " M7" + "-" _
                            + UCaseBémol(NoteInterval2(ton, "M7")) + " 7dim"
                    Case 5 ' 9

                        Mode =
                           Tonique + " mM7(9)" + "-" _
                            + "___" + "-" _
                            + "___" + "-" _
                            + UCaseBémol(NoteInterval2(ton, "4")) + " m7(9)" + "-" _
                            + UCaseBémol(NoteInterval2(ton, "5")) + " 7(b9)" + "-" _
                            + "___" + "-" _
                            + "___"
                    Case 6 '11
                        Mode =
                              Tonique + " m11" + "-" _
                            + "___" + "-" _
                            + "___" + "-" _
                            + "___" + "-" _
                            + "___" + "-" _
                            + UCaseBémol(NoteInterval2(ton, "5#")) + " M7(11#)" + "-" _
                            + "___"

                End Select
            Case "MinM"
                Select Case Typ
                    Case 3
                        Mode =
                        Tonique + " m" + "-" _
                        + UCaseBémol(NoteInterval2(ton, "9")) + " m" + "-" _
                        + UCaseBémol(NoteInterval2(ton, "3m")) + " 5#" + "-" _
                        + UCaseBémol(NoteInterval2(ton, "4")) + "-" _
                        + UCaseBémol(NoteInterval2(ton, "5")) + "-" _
                        + UCaseBémol(NoteInterval2(ton, "6")) + " mb5" + "-" _
                        + UCaseBémol(NoteInterval2(ton, "M7")) + " mb5"

                    Case 4 ' 7
                        Mode =
                        Tonique + " mM7" + "-" _
                        + UCaseBémol(NoteInterval2(ton, "9")) + " m7" + "-" _
                        + UCaseBémol(NoteInterval2(ton, "3m")) + " M75#" + "-" _
                        + UCaseBémol(NoteInterval2(ton, "4")) + " 7" + "-" _
                        + UCaseBémol(NoteInterval2(ton, "5")) + " 7" + "-" _
                        + UCaseBémol(NoteInterval2(ton, "6")) + " m7b5" + "-" _
                        + UCaseBémol(NoteInterval2(ton, "M7")) + " m7b5"
                    Case 5 '9

                        Mode =
                            Tonique + " mM7(9)" + "-" _
                            + "___" + "-" _
                            + "___" + "-" _
                            + UCaseBémol(NoteInterval2(ton, "4")) + " 7(9)" + "-" _
                            + UCaseBémol(NoteInterval2(ton, "5")) + " 7(9)" + "-" _
                            + "___" + "-" _
                            + "___"
                    Case 6 '11
                        Mode =
                            Tonique + " m11" + "-" _
                            + UCaseBémol(NoteInterval2(ton, "9")) + " m11" + "-" _
                            + "___" + "-" _
                            + "___" + "-" _
                            + UCaseBémol(NoteInterval2(ton, "5")) + " 7(11)" + "-" _
                            + "___" + "-" _
                            + "___"
                End Select
        End Select
    End Function

    Function Det_Clef(Tonique As String) As String
        If Len(Tonique) = 1 Or Right(Tonique, 1) = "#" Then
            Det_Clef = "#"
        Else
            Det_Clef = "b"
        End If

        If Trim(Tonique) = "F" Or Trim(Tonique) = "f" Then
            Det_Clef = "b"
        End If
    End Function
    Function Det_ClefLat(Tonique As String) As String
        Select Case Tonique
            Case "Fa", "Sib", "Mib", "Lab"
                Det_ClefLat = "b"
            Case Else
                Det_ClefLat = "#"
        End Select
    End Function
    Function Det_ClefAngl(Tonique As String) As String
        Select Case Tonique
            Case "F", "Bb", "Eb", "Ab"
                Det_ClefAngl = "b"
            Case Else
                Det_ClefAngl = "#"
        End Select
    End Function
    Function Det_ClefLat2(Tonique As String, Mode As String) As String
        Select Case Mode
            Case "Maj"
                Select Case Tonique
                    Case "Fa", "Sib", "Mib", "Lab"
                        Det_ClefLat2 = "b"
                    Case Else
                        Det_ClefLat2 = "#"
                End Select
            Case Else
                Select Case Tonique
                    Case "Ré", "Sol", "Do", "Fa"
                        Det_ClefLat2 = "b"
                    Case Else
                        Det_ClefLat2 = "#"
                End Select
        End Select
    End Function
    Function Det_ClefEn2(Tonique As String, Mode As String) As String
        Select Case Mode
            Case "Maj"
                Select Case Tonique
                    Case "F", "Bb", "Eb", "Ab"
                        Det_ClefEn2 = "b"
                    Case Else
                        Det_ClefEn2 = "#"
                End Select
            Case Else
                Select Case Tonique
                    Case "D", "G", "C", "F"
                        Det_ClefEn2 = "b"
                    Case Else
                        Det_ClefEn2 = "#"
                End Select
        End Select
    End Function

    Function UCaseBémol(Tonique As String)
        Dim a As String

        If Len(Tonique) = 1 Or Right(Tonique, 1) = "#" Then
            UCaseBémol = UCase(Tonique)
        Else
            a = Left(Tonique, 1)
            a = UCase(a)
            UCaseBémol = a + "b"
        End If
    End Function
    Function LCaseBémol(Tonique As String)
        Dim a As String

        If Len(Tonique) = 1 Or Right(Tonique, 1) = "#" Then
            LCaseBémol = LCase(Tonique)
        Else
            a = Left(Tonique, 1)
            a = LCase(a)
            LCaseBémol = a + "b"
        End If
    End Function
    Public Function Det_NotesAccord(Accord As String) As String
        ' La présente fonction Det_NotesAccord est très proche d'une autre fonction Det_NotesAccord3 
        ' La différence entre ces deux fonctions est la suivante :
        '   1 - Det_NotesAccord3 prend la valeur de Clef dans ses paramètres d'entrée
        '   2 - Det_NotesAccord prend la valeur de Clef dans la fonction NoteInterval qu'elle appelle
        ' En conclusion, si une modification est à faire dans l'un de ces deux fonctions Det_NotesAccord3 ou Det_NotesAccord, elle sera 
        ' sans doute à faire dans l'autre fonction.
        Dim Tonique As String
        Dim Chiffrage As String
        '
        Tonique = LCase(Det_Tonique(Accord))
        Chiffrage = Det_Chiffrage(Accord)
        '
        Det_NotesAccord = ""
        Select Case Trim(Chiffrage)
            Case "Maj"
                Det_NotesAccord = Trim(Tonique + "-" + NoteInterval(Tonique, "3") + "-" + NoteInterval(Tonique, "5"))
            Case "m"
                Det_NotesAccord = Trim(Tonique + "-" + NoteInterval(Tonique, "3m") + "-" + NoteInterval(Tonique, "5"))
            Case "mb5"
                Det_NotesAccord = Trim(Tonique + "-" + NoteInterval(Tonique, "3m") + "-" + NoteInterval(Tonique, "b5"))
            Case "7M", "M7"
                Det_NotesAccord = Trim(Tonique + "-" + NoteInterval(Tonique, "3") + "-" + NoteInterval(Tonique, "5") + "-" + NoteInterval(Tonique, "M7"))
            Case "7"
                Det_NotesAccord = Trim(Tonique + "-" + NoteInterval(Tonique, "3") + "-" + NoteInterval(Tonique, "5") + "-" + NoteInterval(Tonique, "7"))
            Case "m7"
                Det_NotesAccord = Trim(Tonique + "-" + NoteInterval(Tonique, "3m") + "-" + NoteInterval(Tonique, "5") + "-" + NoteInterval(Tonique, "7"))
            Case "m7b5"
                Det_NotesAccord = Trim(Tonique + "-" + NoteInterval(Tonique, "3m") + "-" + NoteInterval(Tonique, "b5") + "-" + NoteInterval(Tonique, "7"))
            Case "5#"
                Det_NotesAccord = Trim(Tonique + "-" + NoteInterval(Tonique, "3") + "-" + NoteInterval(Tonique, "5#"))
            Case "mM7", "m7M"
                Det_NotesAccord = Trim(Tonique + "-" + NoteInterval(Tonique, "3m") + "-" + NoteInterval(Tonique, "5") + "-" + NoteInterval(Tonique, "M7"))
            Case "7M5#", "M75#"
                Det_NotesAccord = Trim(Tonique + "-" + NoteInterval(Tonique, "3") + "-" + NoteInterval(Tonique, "5#") + "-" + NoteInterval(Tonique, "M7"))
            Case "7dim"
                Det_NotesAccord = Trim(Tonique + "-" + NoteInterval(Tonique, "3m") + "-" + NoteInterval(Tonique, "b5") + "-" + NoteInterval(Tonique, "6"))
            Case "7M(9)", "M7(9)"
                Det_NotesAccord = Trim(Tonique + "-" + NoteInterval(Tonique, "3") + "-" + NoteInterval(Tonique, "5") + "-" + NoteInterval(Tonique, "M7") + "-" + NoteInterval(Tonique, "9"))
            Case "m7(9)"
                Det_NotesAccord = Trim(Tonique + "-" + NoteInterval(Tonique, "3m") + "-" + NoteInterval(Tonique, "5") + "-" + NoteInterval(Tonique, "7") + "-" + NoteInterval(Tonique, "9"))
            Case "7(9)"
                Det_NotesAccord = Trim(Tonique + "-" + NoteInterval(Tonique, "3") + "-" + NoteInterval(Tonique, "5") + "-" + NoteInterval(Tonique, "7") + "-" + NoteInterval(Tonique, "9"))
            Case "7(b9)"
                Det_NotesAccord = Trim(Tonique + "-" + NoteInterval(Tonique, "3") + "-" + NoteInterval(Tonique, "5") + "-" + NoteInterval(Tonique, "7") + "-" + NoteInterval(Tonique, "b9"))
            Case "7(9#)"
                Det_NotesAccord = Trim(Tonique + "-" + NoteInterval(Tonique, "3") + "-" + NoteInterval(Tonique, "5") + "-" + NoteInterval(Tonique, "7") + "-" + NoteInterval(Tonique, "9#"))
            Case "b9"
                Det_NotesAccord = Trim(Tonique + "-" + NoteInterval(Tonique, "3") + "-" + NoteInterval(Tonique, "5") + "-" + NoteInterval(Tonique, "b9"))
            Case "9#"
                Det_NotesAccord = Trim(Tonique + "-" + NoteInterval(Tonique, "3") + "-" + NoteInterval(Tonique, "5") + "-" + NoteInterval(Tonique, "9#"))
            Case "m9"
                Det_NotesAccord = Trim(Tonique + "-" + NoteInterval(Tonique, "3m") + "-" + NoteInterval(Tonique, "5") + "-" + NoteInterval(Tonique, "9"))
            Case "9"
                Det_NotesAccord = Trim(Tonique + "-" + NoteInterval(Tonique, "3") + "-" + NoteInterval(Tonique, "5") + "-" + NoteInterval(Tonique, "9"))
            Case "7M(11#)", "M7(11#)"
                Det_NotesAccord = Trim(Tonique + "-" + NoteInterval(Tonique, "3") + "-" + NoteInterval(Tonique, "5") + "-" + NoteInterval(Tonique, "M7") + "-" + NoteInterval(Tonique, "11#"))
            Case "7(11)"
                Det_NotesAccord = Trim(Tonique + "-" + NoteInterval(Tonique, "3") + "-" + NoteInterval(Tonique, "5") + "-" + NoteInterval(Tonique, "7") + "-" + NoteInterval(Tonique, "11"))
            Case "11#"
                Det_NotesAccord = Trim(Tonique + "-" + NoteInterval(Tonique, "3") + "-" + NoteInterval(Tonique, "5") + "-" + NoteInterval(Tonique, "11#"))
            Case "11"
                Det_NotesAccord = Trim(Tonique + "-" + NoteInterval(Tonique, "3") + "-" + NoteInterval(Tonique, "5") + "-" + NoteInterval(Tonique, "11"))
            Case "7sus4"
                Det_NotesAccord = Trim(Tonique + "-" + NoteInterval(Tonique, "4") + "-" + NoteInterval(Tonique, "5") + "-" + NoteInterval(Tonique, "7"))
            Case "sus4"
                Det_NotesAccord = Trim(Tonique + "-" + NoteInterval(Tonique, "4") + "-" + NoteInterval(Tonique, "5"))
            Case "m11"
                Det_NotesAccord = Trim(Tonique + "-" + NoteInterval(Tonique, "3m") + "-" + NoteInterval(Tonique, "5") + "-" + NoteInterval(Tonique, "11"))
            Case "m7(11)"
                Det_NotesAccord = Trim(Tonique + "-" + NoteInterval(Tonique, "3m") + "-" + NoteInterval(Tonique, "5") + "-" + NoteInterval(Tonique, "7") + "-" + NoteInterval(Tonique, "11"))
            Case "11"
                Det_NotesAccord = Trim(Tonique + "-" + NoteInterval(Tonique, "3") + "-" + NoteInterval(Tonique, "5") + "-" + NoteInterval(Tonique, "7") + "-" + NoteInterval(Tonique, "11"))
        End Select
    End Function
    Public Function Det_NotesGammes(Gamme As String) As String
        Dim i As Integer
        Dim Tonique As String
        Dim Tonique2 As String
        Dim tbl() As String
        Dim Sauv_Clef As String
        Det_NotesGammes = ""

        Sauv_Clef = Clef
        If Trim(Gamme <> "") Then

            ' Détermnation de la Tonique
            tbl = Split(Trim(Gamme))
            Tonique = Trim(tbl(0))
            ' Recnnaissance de la Tonique dans TabNotes
            Tonique2 = Trim(LCase(Tonique))
            Clef = Det_Clef(Tonique)
            Maj_TabNotes_Minus(Trim(Clef))
            '
            For i = 0 To UBound(TabNotes)
                If Tonique2 = TabNotes(i) Then
                    Exit For
                End If
            Next i
            ' chaînes des notes de la gamme
            Select Case tbl(1)
                Case "Maj"
                    Det_NotesGammes = LCase(Tonique) + " " + TabNotes(i + 2) + " " + TabNotes(i + 4) + " " + TabNotes(i + 5) _
                    + " " + TabNotes(i + 7) + " " + TabNotes(i + 9) + " " + TabNotes(i + 11)
                Case "MinH"
                    Det_NotesGammes = LCase(Tonique) + " " + TabNotes(i + 2) + " " + TabNotes(i + 3) + " " + TabNotes(i + 5) _
                    + " " + TabNotes(i + 7) + " " + TabNotes(i + 8) + " " + TabNotes(i + 11)
                Case "MinM"
                    Det_NotesGammes = LCase(Tonique) + " " + TabNotes(i + 2) + " " + TabNotes(i + 3) + " " + TabNotes(i + 5) _
                    + " " + TabNotes(i + 7) + " " + TabNotes(i + 9) + " " + TabNotes(i + 11)
                Case "MajH"
                    Det_NotesGammes = LCase(Tonique) + " " + TabNotes(i + 2) + " " + TabNotes(i + 4) + " " + TabNotes(i + 5) _
                    + " " + TabNotes(i + 7) + " " + TabNotes(i + 8) + " " + TabNotes(i + 11)
                Case "Penta"
                    Det_NotesGammes = LCase(Tonique) + " " + TabNotes(i + 3) + " " + TabNotes(i + 5) + " " + TabNotes(i + 7) + " " + TabNotes(i + 10)
                Case "Blues"
                    Det_NotesGammes = LCase(Tonique) + " " + TabNotes(i + 3) + " " + TabNotes(i + 5) + " " + TabNotes(i + 6) + " " + TabNotes(i + 7) + " " + TabNotes(i + 10)
            End Select
        End If
        Clef = Sauv_Clef
    End Function
    Public Function Det_NotesGammes3(Gamme As String, Clef1 As String) As String
        Dim i As Integer
        Dim Tonique As String
        Dim Tonique2 As String
        Dim tbl() As String


        Det_NotesGammes3 = ""

        If Trim(Gamme <> "") Then

            ' Détermnation de la Tonique
            tbl = Split(Trim(Gamme))
            Tonique = Trim(tbl(0))
            ' Reconnaissance de la Tonique dans TabNotes
            Tonique2 = Trim(LCase(Tonique))
            Maj_TabNotes_Minus(Trim(Clef1))
            '
            For i = 0 To UBound(TabNotes)
                If Tonique2 = TabNotes(i) Then
                    Exit For
                End If
            Next i
            ' chaînes des noes de la gamme
            Select Case tbl(1)
                Case "Maj"
                    Det_NotesGammes3 = LCase(Tonique) + " " + TabNotes(i + 2) + " " + TabNotes(i + 4) + " " + TabNotes(i + 5) _
                    + " " + TabNotes(i + 7) + " " + TabNotes(i + 9) + " " + TabNotes(i + 11)
                Case "MinH"
                    Det_NotesGammes3 = LCase(Tonique) + " " + TabNotes(i + 2) + " " + TabNotes(i + 3) + " " + TabNotes(i + 5) _
                    + " " + TabNotes(i + 7) + " " + TabNotes(i + 8) + " " + TabNotes(i + 11)
                Case "MinM"
                    Det_NotesGammes3 = LCase(Tonique) + " " + TabNotes(i + 2) + " " + TabNotes(i + 3) + " " + TabNotes(i + 5) _
                    + " " + TabNotes(i + 7) + " " + TabNotes(i + 9) + " " + TabNotes(i + 11)
                Case "MajH"
                    Det_NotesGammes3 = LCase(Tonique) + " " + TabNotes(i + 2) + " " + TabNotes(i + 4) + " " + TabNotes(i + 5) _
                    + " " + TabNotes(i + 7) + " " + TabNotes(i + 8) + " " + TabNotes(i + 11)
                Case "Penta"
                    Det_NotesGammes3 = LCase(Tonique) + " " + TabNotes(i + 3) + " " + TabNotes(i + 5) + " " + TabNotes(i + 7) + " " + TabNotes(i + 10)
                Case "Blues"
                    Det_NotesGammes3 = LCase(Tonique) + " " + TabNotes(i + 3) + " " + TabNotes(i + 5) + " " + TabNotes(i + 6) + " " + TabNotes(i + 7) + " " + TabNotes(i + 10)
            End Select
        End If
    End Function

    Public Function NoteInterval(Tonique As String, Interval As String) As String
        ' La présente fonction NoteInterval est très proche d'une autre fonction NoteInterval2 
        ' qui se trouve dans le module Form1. La différence entre ces deux fonction est la suivante :
        '   1 - NoteInterval2 prend la valeur de Clef dans la tonalite courante (Combobox1.text)
        '   2 - NoteInterval prend la valeur de Clef dans la tonique d'un accord, d'une gamme ou d'un mode
        ' En conclusion, si une modification est à faire dans l'un de ces deux fonctions, elle sera 
        ' sans doute à faire dans l'autre fonction.

        Dim IndexTonique As Integer
        Dim Sauv_Clef As String
        'Dim Clef As String
        '
        'Clef = Det_Clef(ComboBox20.text)
        'If EnrechercheGammes = True Then
        'Maj_TabNotes(Trim("#"))
        'Else
        Sauv_Clef = Clef
        Clef = Det_Clef(Tonique)
        '
        Maj_TabNotes_Minus(Trim(Clef))
        'End If
        '

        NoteInterval = ""
        IndexTonique = TrouverNoteDansTabNotes(LCase(Tonique))
        '
        Select Case Interval
            Case "b2", "b9"
                NoteInterval = TabNotes(IndexTonique + 1)
            Case "2", "9"
                NoteInterval = TabNotes(IndexTonique + 2)
            Case "9#", "3m"
                NoteInterval = TabNotes(IndexTonique + 3)
            Case "3"
                NoteInterval = TabNotes(IndexTonique + 4)
            Case "4", "11"
                NoteInterval = TabNotes(IndexTonique + 5)
            Case "b5", "4#", "11#"
                NoteInterval = TabNotes(IndexTonique + 6)
            Case "5"
                NoteInterval = TabNotes(IndexTonique + 7)
            Case "5#", "6m", "13b"
                NoteInterval = TabNotes(IndexTonique + 8)
            Case "6", "13"
                NoteInterval = TabNotes(IndexTonique + 9)
            Case "7"
                NoteInterval = TabNotes(IndexTonique + 10)
            Case "7M", "M7"
                NoteInterval = TabNotes(IndexTonique + 11)
        End Select
        '
        Clef = Sauv_Clef
    End Function
    Public Function NoteInterval2(Tonique As String, Interval As String) As String
        ' La présente fonction NoteInterval est très proche d'une autre fonction NoteInterval2 
        ' qui se trouve dans le module Form1. La différence entre ces deux fonction est la suivante :
        '   1 - NoteInterval2 prend la valeur de Clef dans la tonalite courante (Tonacours=Combobox1.text)
        '   2 - NoteInterval prend la valeur de Clef dans la tonique d'un accord, d'une gamme ou d'un mode
        ' En conclusion, si une modification est à faire dans l'un de ces deux fonctions, elle sera 
        ' sans doute à faire dans l'autre fonction.

        Dim IndexTonique As Integer
        Dim Sauv_Clef As String
        Dim a As String
        Dim tbl() As String
        'Dim Clef As String
        '
        'Clef = Det_Clef(ComboBox20.text)
        'If EnrechercheGammes = True Then
        'Maj_TabNotes(Trim("#"))
        'Else
        Sauv_Clef = Clef
        tbl = Split(Tonacours)
        a = Trim(tbl(0))
        'a = Tonique
        If Langue = "fr" Then
            Clef = Det_ClefLat(Trim(a))
        Else
            Clef = Det_ClefAngl(Trim(a))
        End If
        '
        Maj_TabNotes_Minus(Trim(Clef))
        'End If
        '

        NoteInterval2 = ""
        IndexTonique = TrouverNoteDansTabNotes(LCase(Tonique))
        '
        Select Case Interval
            Case "b2", "b9"
                NoteInterval2 = TabNotes(IndexTonique + 1)
            Case "2", "9"
                NoteInterval2 = TabNotes(IndexTonique + 2)
            Case "9#", "3m"
                NoteInterval2 = TabNotes(IndexTonique + 3)
            Case "3"
                NoteInterval2 = TabNotes(IndexTonique + 4)
            Case "4", "11"
                NoteInterval2 = TabNotes(IndexTonique + 5)
            Case "b5", "4#", "11#"
                NoteInterval2 = TabNotes(IndexTonique + 6)
            Case "5"
                NoteInterval2 = TabNotes(IndexTonique + 7)
            Case "5#", "6m", "13b"
                NoteInterval2 = TabNotes(IndexTonique + 8)
            Case "6", "13"
                NoteInterval2 = TabNotes(IndexTonique + 9)
            Case "7"
                NoteInterval2 = TabNotes(IndexTonique + 10)
            Case "7M", "M7"
                NoteInterval2 = TabNotes(IndexTonique + 11)
        End Select
        '
        Clef = Sauv_Clef
    End Function
    Public Function CG_NoteInterval(Tonique As String, Interval As String) As String
        Dim IndexTonique As Integer
        'Dim Clef As String
        '
        'Clef = Det_Clef(ComboBox20.text)
        'Maj_TabNotes(Trim(Clef))
        CG_NoteInterval = ""
        IndexTonique = TrouverNoteDansTabNotes(LCase(Tonique))
        '
        Select Case Interval
            Case "b2", "b9"
                CG_NoteInterval = TabNotes(IndexTonique + 1)
            Case "2", "9"
                CG_NoteInterval = TabNotes(IndexTonique + 2)
            Case "9#", "3m"
                CG_NoteInterval = TabNotes(IndexTonique + 3)
            Case "3"
                CG_NoteInterval = TabNotes(IndexTonique + 4)
            Case "4", "11"
                CG_NoteInterval = TabNotes(IndexTonique + 5)
            Case "b5", "4#", "11#"
                CG_NoteInterval = TabNotes(IndexTonique + 6)
            Case "5"
                CG_NoteInterval = TabNotes(IndexTonique + 7)
            Case "5#", "6m", "13b"
                CG_NoteInterval = TabNotes(IndexTonique + 8)
            Case "6", "13"
                CG_NoteInterval = TabNotes(IndexTonique + 9)
            Case "7"
                CG_NoteInterval = TabNotes(IndexTonique + 10)
            Case "7M", "M7"
                CG_NoteInterval = TabNotes(IndexTonique + 11)
        End Select

    End Function
    Public Function Det_Tonique(EventH As String) As String
        Dim tbl() As String

        tbl = Split(EventH)
        Det_Tonique = tbl(0)
    End Function
    Public Function Det_Chiffrage(EventH As String) As String
        Dim tbl() As String

        tbl = Split(EventH)
        If UBound(tbl) > 0 Then
            Det_Chiffrage = tbl(1)
        Else
            Det_Chiffrage = "Maj"
        End If
    End Function
    Function Det_Modulation(Tonique As String, Interval As String) As Integer
        Dim TonMode As String

        TonMode = NoteInterval(Trim(Tonique), Interval)

        Select Case Trim(TonMode)
            Case "c"
                Det_Modulation = 0
            Case "c#"
                Det_Modulation = 1
            Case "d"
                Det_Modulation = 2
            Case "d#"
                Det_Modulation = 3
            Case "e"
                Det_Modulation = 4
            Case "f"
                Det_Modulation = 5
            Case "f#"
                Det_Modulation = 6
            Case "g"
                Det_Modulation = 7
            Case "g#"
                Det_Modulation = 8
            Case "eb"
                Det_Modulation = 9
            Case "a"
                Det_Modulation = 10
            Case "a#"
                Det_Modulation = 11
            Case "b"
                Det_Modulation = 12
            Case Else
                Det_Modulation = -1
        End Select
    End Function
    Public Function PlacementOctave(Notes As String, Octave As String) As String
        Dim tbl() As String
        Dim i As Integer
        Dim n As Integer
        Dim a As String
        Dim IndexMax As Integer

        a = ""
        Select Case Trim(Octave)
            Case "-2"
                n = 1
            Case "-1"
                n = 2
            Case "0"
                n = 3
            Case "+1"
                n = 4
        End Select
        '
        tbl = Split(Notes, "-")
        Maj_TabNotes_Minus(Clef) ' nécessaire pour la procédure suivante ChangementOctave
        '
        IndexMax = UBound(tbl)
        For i = 0 To UBound(tbl)

            If i <> 0 Then
                If ChangementOctave(tbl(i - 1), tbl(i)) Then
                    n = n + 1
                End If
            End If
            a = a + Trim(tbl(i) + Trim(Str(n))) + " "

        Next i
        '
        PlacementOctave = Trim(a)
        '
    End Function
    Function ChangementOctave(Note1 As String, Note2 As String) As Boolean
        Dim i As Integer
        Dim j As Integer

        ChangementOctave = False
        '
        For i = 0 To 35
            If TabNotes(i) = Note1 Then
                j = i
                Exit For
            End If
        Next i
        '
        For i = j To 35
            If TabNotes(i) = Note2 Then
                Exit For
            End If
        Next
        If i > 11 Then
            ChangementOctave = True
        End If
    End Function
    Function Det_GammesJouablesAccord(Accord As String) As List(Of String)
        Dim tbl() As String
        Dim Tonique As String
        Dim Chiffrage As String
        Dim i As Integer
        Dim ListGamme As New List(Of String)
        '
        tbl = Split(Accord)
        Tonique = LCase(tbl(0))
        '
        Chiffrage = ""
        If UBound(tbl) > 0 Then
            Chiffrage = tbl(1)
        End If
        '
        ' Placement dans de la Tonique dans  Tabnotes
        ' *******************************************
        For i = 0 To 30
            If TabNotes(i) = Tonique Then
                Exit For
            End If

        Next i
        '
        Select Case Trim(Chiffrage)

            Case ""
                ListGamme.Add(UCase(TabNotes(i)) + " Maj")           ' C Maj
                ListGamme.Add(UCase(TabNotes(i + 5)) + " Maj")       ' F Maj
                ListGamme.Add(UCase(TabNotes(i + 7)) + " Maj")       ' G Maj
                '
                ListGamme.Add(UCase(TabNotes(i + 4)) + " MinH")       ' E MinH
                ListGamme.Add(UCase(TabNotes(i + 5)) + " MinH")       ' F MinH

                ListGamme.Add(UCase(TabNotes(i + 5)) + " MinM")       ' F MinM
                ListGamme.Add(UCase(TabNotes(i + 7)) + " MinM")       ' G MinM

                ListGamme.Add(UCase(TabNotes(i + 2)) + " PentaMin")  ' D pentamin
                ListGamme.Add(UCase(TabNotes(i + 4)) + " PentaMin")  ' E pentamin
                ListGamme.Add(UCase(TabNotes(i + 7)) + " PentaMin")  ' G pentamin
                ListGamme.Add(UCase(TabNotes(i + 9)) + " PentaMin")  ' A pentamin
                ListGamme.Add(UCase(TabNotes(i + 11)) + " PentaMin") ' B pentamin
                '
                ListGamme.Add(UCase(TabNotes(i)) + " Blues")           ' C Maj
                ListGamme.Add(UCase(TabNotes(i + 5)) + " Blues")       ' F Maj
                ListGamme.Add(UCase(TabNotes(i + 7)) + " Blues")       ' G Maj
                '

            Case "5#"

                ListGamme.Add(UCase(TabNotes(i + 9)) + " MinH")  'A MinH
                ListGamme.Add(UCase(TabNotes(i + 9)) + " MinM") ' A MinM

            Case "9"

                ListGamme.Add(UCase(TabNotes(i)) + " Maj")     ' C Maj
                ListGamme.Add(UCase(TabNotes(i + 5)) + " Maj") ' F Maj
                ListGamme.Add(UCase(TabNotes(i + 7)) + " Maj") ' G Maj

                ListGamme.Add(UCase(TabNotes(i + 5)) + " MinM")       ' F MinM
                ListGamme.Add(UCase(TabNotes(i + 7)) + " MinM")       ' G MinM
                '                
                ListGamme.Add(UCase(TabNotes(i + 2)) + " PentaMin")  ' D pentamin
                ListGamme.Add(UCase(TabNotes(i + 4)) + " PentaMin")  ' E pentamin
                ListGamme.Add(UCase(TabNotes(i + 7)) + " PentaMin")  ' G pentamin
                ListGamme.Add(UCase(TabNotes(i + 9)) + " PentaMin")  ' A pentamin

            Case "9#"
                ListGamme.Add(UCase(TabNotes(i + 4)) + " MinH")       ' E MinH

            Case "11"
                ListGamme.Add(UCase(TabNotes(i)) + " Maj")            ' C Maj
                ListGamme.Add(UCase(TabNotes(i + 5)) + " Maj")        ' F Maj

                ListGamme.Add(UCase(TabNotes(i + 5)) + " MinH")       ' F MinH
                ListGamme.Add(UCase(TabNotes(i + 5)) + " MinM")       ' F MinM

                ListGamme.Add(UCase(TabNotes(i + 2)) + " PentaMin")  ' D pentamin
                ListGamme.Add(UCase(TabNotes(i + 4)) + " PentaMin")  ' E pentamin
                ListGamme.Add(UCase(TabNotes(i + 9)) + " PentaMin")  ' A pentamin

            Case "11#"
                ListGamme.Add(UCase(TabNotes(i + 7)) + " Maj")       ' G Maj

                ListGamme.Add(UCase(TabNotes(i + 4)) + " MinH")       ' E MinH
                ListGamme.Add(UCase(TabNotes(i + 7)) + " MinM")       ' G MinM

                ListGamme.Add(UCase(TabNotes(i + 4)) + " PentaMin")  ' E pentamin
                ListGamme.Add(UCase(TabNotes(i + 9)) + " PentaMin")  ' A pentamin
                ListGamme.Add(UCase(TabNotes(i + 11)) + " PentaMin") ' B pentamin

            Case "sus4"
                ListGamme.Add(UCase(TabNotes(i)) + " Maj")           ' C Maj
                ListGamme.Add(UCase(TabNotes(i + 5)) + " Maj")       ' F Maj

                ListGamme.Add(UCase(TabNotes(i + 5)) + " MinH")       ' F MinH
                ListGamme.Add(UCase(TabNotes(i + 5)) + " MinM")       ' F MinM

                ListGamme.Add(UCase(TabNotes(i + 2)) + " PentaMin")  ' D pentamin
                ListGamme.Add(UCase(TabNotes(i + 4)) + " PentaMin")  ' E pentamin
                ListGamme.Add(UCase(TabNotes(i + 9)) + " PentaMin")  ' A pentamin
                ListGamme.Add(UCase(TabNotes(i + 11)) + " PentaMin") ' B pentamin

            Case "7M", "M7"
                ListGamme.Add(UCase(TabNotes(i)) + " Maj")           ' C Maj
                ListGamme.Add(UCase(TabNotes(i + 7)) + " Maj")       ' G Maj

                ListGamme.Add(UCase(TabNotes(i + 4)) + " MinH")       ' E MinH

                ListGamme.Add(UCase(TabNotes(i + 2)) + " PentaMin")  ' D pentamin
                ListGamme.Add(UCase(TabNotes(i + 4)) + " PentaMin")  ' E pentamin
                ListGamme.Add(UCase(TabNotes(i + 9)) + " PentaMin")  ' A pentamin
                ListGamme.Add(UCase(TabNotes(i + 11)) + " PentaMin") ' B pentamin

            Case "7M(9)", "M7(9)"
                ListGamme.Add(UCase(TabNotes(i)) + " Maj")           ' C Maj
                ListGamme.Add(UCase(TabNotes(i + 7)) + " Maj")       ' G Maj
                '
                ListGamme.Add(UCase(TabNotes(i + 2)) + " PentaMin")  ' D pentamin
                ListGamme.Add(UCase(TabNotes(i + 4)) + " PentaMin")  ' E pentamin
                ListGamme.Add(UCase(TabNotes(i + 9)) + " PentaMin")  ' A pentamin
                ListGamme.Add(UCase(TabNotes(i + 11)) + " PentaMin") ' B pentamin

            Case "7M(11#)", "M7(11#)"
                ListGamme.Add(UCase(TabNotes(i + 7)) + " Maj")       ' G Maj

                ListGamme.Add(UCase(TabNotes(i + 4)) + " MinH")       ' E MinH

                ListGamme.Add(UCase(TabNotes(i + 4)) + " PentaMin")  ' E pentamin
                ListGamme.Add(UCase(TabNotes(i + 9)) + " PentaMin")  ' A pentamin
                ListGamme.Add(UCase(TabNotes(i + 11)) + " PentaMin") ' B pentamin

            Case "7"
                ListGamme.Add(UCase(TabNotes(i + 5)) + " Maj")       ' F Maj

                ListGamme.Add(UCase(TabNotes(i + 5)) + " MinH")       ' F MinH

                ListGamme.Add(UCase(TabNotes(i + 5)) + " MinM")       ' F MinM
                ListGamme.Add(UCase(TabNotes(i + 7)) + " MinM")       ' G MinM

                ListGamme.Add(UCase(TabNotes(i + 2)) + " PentaMin")  ' D pentamin
                ListGamme.Add(UCase(TabNotes(i + 7)) + " PentaMin")  ' G pentamin
                ListGamme.Add(UCase(TabNotes(i + 9)) + " PentaMin")  ' A pentamin
                '
                ListGamme.Add(UCase(TabNotes(i)) + " Blues")           ' C Maj
                ListGamme.Add(UCase(TabNotes(i + 5)) + " Blues")       ' F Maj
                ListGamme.Add(UCase(TabNotes(i + 7)) + " Blues")       ' G Maj

            Case "7(9)"
                ListGamme.Add(UCase(TabNotes(i + 5)) + " Maj")       ' F Maj

                ListGamme.Add(UCase(TabNotes(i + 5)) + " MinM")       ' F MinM
                ListGamme.Add(UCase(TabNotes(i + 7)) + " MinM")       ' G MinM

                ListGamme.Add(UCase(TabNotes(i + 2)) + " PentaMin")  ' D pentamin
                ListGamme.Add(UCase(TabNotes(i + 7)) + " PentaMin")  ' G pentamin
                ListGamme.Add(UCase(TabNotes(i + 9)) + " PentaMin")  ' A pentamin


            Case "7(9#)"
                ListGamme.Add(UCase(TabNotes(i + 4)) + " MinH")       ' E MinH

            Case "7(b9)"
                ListGamme.Add(UCase(TabNotes(i + 5)) + " MinH")       ' E MinH

            Case "7(11)"
                ListGamme.Add(UCase(TabNotes(i + 5)) + " Maj")        ' F Maj

                ListGamme.Add(UCase(TabNotes(i + 5)) + " MinH")       ' F MinH
                ListGamme.Add(UCase(TabNotes(i + 5)) + " MinM")       ' F MinM
                '
                ListGamme.Add(UCase(TabNotes(i + 2)) + " PentaMin")  ' D pentamin
                ListGamme.Add(UCase(TabNotes(i + 7)) + " PentaMin")  ' G pentamin
                ListGamme.Add(UCase(TabNotes(i + 9)) + " PentaMin")  ' A pentamin

            Case "7sus4"
                ListGamme.Add(UCase(TabNotes(i + 5)) + " Maj")       ' F Maj

                ListGamme.Add(UCase(TabNotes(i + 5)) + " MinH")       ' F MinH
                ListGamme.Add(UCase(TabNotes(i + 5)) + " MinM")       ' F MinM

                ListGamme.Add(UCase(TabNotes(i + 2)) + " PentaMin")  ' D pentamin
                ListGamme.Add(UCase(TabNotes(i + 7)) + " PentaMin")  ' G pentamin
                ListGamme.Add(UCase(TabNotes(i + 9)) + " PentaMin")  ' A pentamin


            Case "m"
                ListGamme.Add(UCase(TabNotes(i + 10)) + " Maj")       ' A# Maj
                ListGamme.Add(UCase(TabNotes(i + 3)) + " Maj")        ' D# Maj 
                ListGamme.Add(UCase(TabNotes(i + 8)) + " Maj")        ' G# Maj

                ListGamme.Add(UCase(TabNotes(i)) + " MinH")           ' C MinH
                ListGamme.Add(UCase(TabNotes(i + 7)) + " MinH")       ' G MinH

                ListGamme.Add(UCase(TabNotes(i)) + " MinM")           ' C MinM
                ListGamme.Add(UCase(TabNotes(i + 10)) + " MinM")      ' A# MinM

                ListGamme.Add(UCase(TabNotes(i)) + " PentaMin")      ' C pentamin
                ListGamme.Add(UCase(TabNotes(i + 2)) + " PentaMin")  ' D pentamin
                ListGamme.Add(UCase(TabNotes(i + 5)) + " PentaMin")  ' F pentamin
                ListGamme.Add(UCase(TabNotes(i + 7)) + " PentaMin")  ' G pentamin
                ListGamme.Add(UCase(TabNotes(i + 10)) + " PentaMin") ' A# pentamin

                '
                ListGamme.Add(UCase(TabNotes(i)) + " Blues")           ' C Maj
                ListGamme.Add(UCase(TabNotes(i + 5)) + " Blues")       ' F Maj
                ListGamme.Add(UCase(TabNotes(i + 7)) + " Blues")       ' G Maj

            Case "mb5"
                ListGamme.Add(UCase(TabNotes(i + 1)) + " Maj")        ' C# Maj

                ListGamme.Add(UCase(TabNotes(i + 1)) + " MinH")       ' C# MinH
                ListGamme.Add(UCase(TabNotes(i + 10)) + " MinH")      ' A# MinH

                ListGamme.Add(UCase(TabNotes(i + 1)) + " MinM")       ' C# MinM
                ListGamme.Add(UCase(TabNotes(i + 3)) + " MinM")       ' D# MinM

                ListGamme.Add(UCase(TabNotes(i + 3)) + " PentaMin")   ' D# pentamin
                ListGamme.Add(UCase(TabNotes(i + 5)) + " PentaMin")   ' F pentamin
                ListGamme.Add(UCase(TabNotes(i + 10)) + " PentaMin")  ' A# pentamin

            Case "m9"
                ListGamme.Add(UCase(TabNotes(i + 10)) + " Maj")       ' A# Maj
                ListGamme.Add(UCase(TabNotes(i + 3)) + " Maj")        ' D# Maj 

                ListGamme.Add(UCase(TabNotes(i)) + " MinH")           ' C MinH
                ListGamme.Add(UCase(TabNotes(i + 7)) + " MinH")       ' G MinH

                ListGamme.Add(UCase(TabNotes(i)) + " MinM")           ' C MinM

                ListGamme.Add(UCase(TabNotes(i)) + " PentaMin")      ' C pentamin
                ListGamme.Add(UCase(TabNotes(i + 2)) + " PentaMin")  ' D pentamin
                ListGamme.Add(UCase(TabNotes(i + 5)) + " PentaMin")  ' F pentamin
                ListGamme.Add(UCase(TabNotes(i + 7)) + " PentaMin")  ' G pentamin

            Case "m11"

                ListGamme.Add(UCase(TabNotes(i + 10)) + " Maj")       ' A# Maj
                ListGamme.Add(UCase(TabNotes(i + 3)) + " Maj")        ' D# Maj 
                ListGamme.Add(UCase(TabNotes(i + 8)) + " Maj")        ' G# Maj

                ListGamme.Add(UCase(TabNotes(i)) + " MinH")           ' C MinH

                ListGamme.Add(UCase(TabNotes(i)) + " MinM")           ' C MinM
                ListGamme.Add(UCase(TabNotes(i + 10)) + " MinM")      ' A# MinM

                ListGamme.Add(UCase(TabNotes(i)) + " PentaMin")      ' C pentamin
                ListGamme.Add(UCase(TabNotes(i + 2)) + " PentaMin")  ' D pentamin
                ListGamme.Add(UCase(TabNotes(i + 5)) + " PentaMin")  ' F pentamin
                ListGamme.Add(UCase(TabNotes(i + 7)) + " PentaMin")  ' G pentamin
                ListGamme.Add(UCase(TabNotes(i + 10)) + " PentaMin") ' A# pentamin

                '
            Case "m7"
                ListGamme.Add(UCase(TabNotes(i + 10)) + " Maj")       ' A# Maj
                ListGamme.Add(UCase(TabNotes(i + 3)) + " Maj")        ' D# Maj 
                ListGamme.Add(UCase(TabNotes(i + 8)) + " Maj")        ' G# Maj

                ListGamme.Add(UCase(TabNotes(i + 7)) + " MinH")        ' G MinH
                ListGamme.Add(UCase(TabNotes(i + 10)) + " MinM")       ' A# MinM

                ListGamme.Add(UCase(TabNotes(i)) + " PentaMin")       ' C pentamin
                ListGamme.Add(UCase(TabNotes(i + 2)) + " PentaMin")   ' D pentamin
                ListGamme.Add(UCase(TabNotes(i + 5)) + " PentaMin")   ' F pentamin
                ListGamme.Add(UCase(TabNotes(i + 7)) + " PentaMin")   ' G pentamin
                ListGamme.Add(UCase(TabNotes(i + 10)) + " PentaMin")  ' A# pentamin
                '
                ListGamme.Add(UCase(TabNotes(i)) + " Blues")           ' C Maj
                ListGamme.Add(UCase(TabNotes(i + 5)) + " Blues")       ' F Maj
                ListGamme.Add(UCase(TabNotes(i + 7)) + " Blues")       ' G Maj
                '

            Case "m7b5" '
                ListGamme.Add(UCase(TabNotes(i + 2)) + " Maj")       ' C# Maj

                ListGamme.Add(UCase(TabNotes(i + 10)) + " MinH")     ' A# MinH

                ListGamme.Add(UCase(TabNotes(i + 1)) + " MinM")      ' C# MinM 
                ListGamme.Add(UCase(TabNotes(i + 3)) + " MinM")      ' D# MinM 

                ListGamme.Add(UCase(TabNotes(i + 3)) + " PentaMin")   ' D# pentamin
                ListGamme.Add(UCase(TabNotes(i + 5)) + " PentaMin")   ' F pentamin
                ListGamme.Add(UCase(TabNotes(i + 10)) + " PentaMin")  ' A# pentamin

            Case "m7(9)"
                ListGamme.Add(UCase(TabNotes(i + 10)) + " Maj")       ' A# Maj
                ListGamme.Add(UCase(TabNotes(i + 3)) + " Maj")        ' D# Maj 

                ListGamme.Add(UCase(TabNotes(i + 7)) + " MinH")        ' G MinH

                ListGamme.Add(UCase(TabNotes(i)) + " PentaMin")       ' C pentamin
                ListGamme.Add(UCase(TabNotes(i + 2)) + " PentaMin")   ' D pentamin
                ListGamme.Add(UCase(TabNotes(i + 5)) + " PentaMin")   ' F pentamin
                ListGamme.Add(UCase(TabNotes(i + 7)) + " PentaMin")   ' G pentamin

        End Select
        Det_GammesJouablesAccord = ListGamme
    End Function
    Public Function Det_NumNote_BasseMoins1(typPiste As String, Zone As Integer) As Integer
        Dim a As String
        Det_NumNote_BasseMoins1 = 0
        Select Case typPiste
            Case "PisteHorsZone"
                a = THorsZone.VoixAsso_OctaveMoins1
                Det_NumNote_BasseMoins1 = Val(Microsoft.VisualBasic.Right(a, 1)) - 1
            Case "PisteZone"
                a = TZone(Zone).VoixAsso_OctaveMoins1
                Det_NumNote_BasseMoins1 = Val(Microsoft.VisualBasic.Right(a, 1)) - 1
        End Select

    End Function
    Public Function Det_NumNote_Large(typPiste As String, Zone As Integer) As Integer
        Dim a As String
        Det_NumNote_Large = 0
        Select Case typPiste
            Case "PisteHorsZone"
                a = THorsZone.VoixAsso_OctavePlus1
                Det_NumNote_Large = Val(Microsoft.VisualBasic.Right(a, 1)) - 1
            Case "PisteZone"
                a = TZone(Zone).VoixAsso_OctavePlus1
                Det_NumNote_Large = Val(Microsoft.VisualBasic.Right(a, 1)) - 1
        End Select

    End Function
    Public Function Trad_ListeNotesEnD(Notes As String, Sep As String) As String
        Dim i, j As Integer
        Dim tbl() As String

        tbl = Split(Notes, Sep)

        For i = 0 To UBound(tbl)
            For j = 0 To TabNotesD.Count - 1
                If Trim(LCase(tbl(i))) = Trim(LCase(TabNotesB(j))) Then
                    tbl(i) = LCase(TabNotesD(j))
                    Exit For
                End If
            Next j
        Next i
        Trad_ListeNotesEnD = Join(tbl, "-")
    End Function
    Public Function Trad_ListeNotesEnB(Notes As String, Sep As String)
        Dim i, j As Integer
        Dim tbl() As String

        tbl = Split(Notes, Sep)
        '
        For i = 0 To UBound(tbl)
            For j = 0 To 35
                If LCase(tbl(i)) = TabNotesD(j) Then
                    tbl(i) = TabNotesB(j)
                    Exit For
                End If
            Next j
        Next i
        Trad_ListeNotesEnB = Join(tbl, "-")
    End Function
    Public Function Trad_NoteEnD(Note As String)

        Trad_NoteEnD = Note
        Select Case Note
            Case "Db", "db"
                Trad_NoteEnD = "C#"
            Case "Eb", "eb"
                Trad_NoteEnD = "D#"
            Case "Gb", "gb"
                Trad_NoteEnD = "F#"
            Case "Ab", "ab"
                Trad_NoteEnD = "G#"
            Case "Bb", "bb"
                Trad_NoteEnD = "A#"
            Case Else
                Trad_NoteEnD = Note
        End Select

    End Function
    Public Function Trad_NoteEnDMin(Note As String)

        Trad_NoteEnDMin = Note
        Select Case Note
            Case "Db", "db"
                Trad_NoteEnDMin = "c#"
            Case "Eb", "eb"
                Trad_NoteEnDMin = "d#"
            Case "Gb", "gb"
                Trad_NoteEnDMin = "f#"
            Case "Ab", "ab"
                Trad_NoteEnDMin = "g#"
            Case "Bb", "bb"
                Trad_NoteEnDMin = "a#"
            Case Else
                Trad_NoteEnDMin = Note
        End Select

    End Function
    Public Function Trad_BemDies_Maj(NOTE As String) As String
        Select Case Trim(NOTE)
            Case "Db"
                Return "C#"
            Case "Eb"
                Return "D#"
            Case "Gb"
                Return "F#"
            Case "Ab"
                Return "G#"
            Case "Bb"
                Return "A#"
            Case Else
                Return NOTE
        End Select
    End Function
    Public Function Trad_DiesBem_Maj(NOTE As String) As String
        Select Case Trim(NOTE)
            Case "C#"
                Return "Db"
            Case "D#"
                Return "Eb"
            Case "F#"
                Return "Gb"
            Case "G#"
                Return "Ab"
            Case "A#"
                Return "Bb"
            Case Else
                Return NOTE
        End Select
    End Function
End Module
