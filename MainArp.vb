Public Module MainArp

    ' ***********
    ' * CLASSES *
    ' ***********
    ' Classe utilisée par l'export MIDI : détermine le nombre de pistes à créer dans le fichier MIDI
    Class NbPistesUtiles
        Public Nb As Integer ' nombre de pistes utilisées
        Public TPisteUtil(0 To nb_TotalPistes) As Boolean ' Arrangement1.Nb_PistesMidi

        Public Sub New()
            Dim i, j, ind As Integer
            Dim NotesAcc() As String = {String.Empty} ' notes des accords
            Dim listnotesAcc As String = Form1.Récup_notesAcc
            Dim MagnétoAcc() As String = {String.Empty} ' Numéros des magnéto
            Dim listMagnéto As String = Form1.Récup_NumMagnéto
            Dim NbAccords As Integer


            ' Initialisation
            ' ***************
            Nb = 0
            For i = 0 To UBound(TPisteUtil)
                TPisteUtil(i) = False
            Next
            ' Détermination du nombre d'accord
            ' ********************************
            NotesAcc = listnotesAcc.Split(",")
            NbAccords = UBound(NotesAcc)
            '
            ' 1 - Détermination de l'activité des pistes de variations
            ' ********************************************************
            MagnétoAcc = listMagnéto.Split(",")
            For i = 0 To NbAccords
                ind = Convert.ToInt16(MagnétoAcc(i))
                For j = 0 To nb_PistesVar - 1 'Arrangement1.Nb_PistesMidi - 1
                    If TPisteUtil(j) = False Then
                        TPisteUtil(j) = Convert.ToBoolean(Form1.Récup_Mute(j, ind))
                    End If
                Next j
            Next i
            '
            ' 2 - Détermination de l'activitéde des pistes pianoroll
            ' ******************************************************
            For i = 0 To nb_PianoRoll - 1
                TPisteUtil(nb_PistesVar + i) = Form1.listPIANOROLL(i).PMute
            Next
            '
            ' 3 - Détermination de l'activité de la piste de Batterie
            ' *******************************************************
            TPisteUtil(nb_PistesVar + nb_PianoRoll) = Form1.Drums.PMute
            ' Détermination du nombre de pistes utilisées
            For i = 0 To UBound(TPisteUtil)
                If TPisteUtil(i) Then Nb = Nb + 1 ' Nb est une propriété publique de la classe
            Next i
        End Sub
    End Class

    ' Classe Partition
    ' ****************
    Public Class Partition
        Public Tempo As String = String.Empty
        Public Métrique = String.Empty
        Public NomFichier As String = String.Empty
        ' Public ReadOnly Nb_Magnetos As Integer = 7
        ' Public ReadOnly Nb_PistesMidi As Integer = NombrePistes - 2 ' cette info fournit l'index de la dernière piste
        ' Public ReadOnly Nb_Pistes As Integer = Nb_Magnetos * 6 ' Pistes (dans NB_Pistes) est pris ici au sens de générateurs
        ' Public ReadOnly nbMesures As Integer = 48
        Public NumAccords As New AffAcc
        Class AffAcc
            Public NumAcc As New List(Of String)
            Public LectEcr As Boolean
            Public PointeurLect As Integer
        End Class
        '
        ' **********************************************************************
        ' Récup_NumAcc : récupération du N° du 1er accord devant être joué     *
        ' **********************************************************************
        Public ReadOnly Property PAcc(num As Integer) As Integer
            Get
                Dim i As Integer = Me.NumAccords.NumAcc(num)
                Return i
            End Get
        End Property
        Public Sub New(oMétrique As String, oNomFichier As String) ' oTempo As String,
            'Tempo = Form1.Tempo.Value.ToString
            Métrique = oMétrique
            NomFichier = oNomFichier
        End Sub
    End Class
    ' Classe DEBUG : outils de debug
    '        *****   ***************
    Public Class DbgParNotes
        Public durée As String
        Public note As String
        Public dyn As String
        Public canal As String
        Public numPiste As String
        Public position As String
    End Class

    ''' <summary>
    ''' CLASSE PISTE : La classe PISTE de décomposer le travail de génération du fichier MIDI et du Scheduler MIDI. Les notes
    ''' et les contrôleuers sont placées dans la liste .part de cette classe sous un format "propriétaire" qui est utilisé p
    ''' our la génération MIDI.
    ''' </summary>
    Public Class Piste
        ' Constantes
        ' **********
        Public ValNote As New List(Of String) From {
                   "C-1", "C#-1", "D-1", "D#-1", "E-1", "F-1", "F#-1", "G-1", "G#-1", "A-1", "A#-1", "B-1",
                   "C0", "C#0", "D0", "D#0", "E0", "F0", "F#0", "G0", "G#0", "A0", "A#0", "B0",
                    "C1", "C#1", "D1", "D#1", "E1", "F1", "F#1", "G1", "G#1", "A1", "A#1", "B1",
                    "C2", "C#2", "D2", "D#2", "E2", "F2", "F#2", "G2", "G#2", "A2", "A#2", "B2",
                    "C3", "C#3", "D3", "D#3", "E3", "F3", "F#3", "G3", "G#3", "A3", "A#3", "B3",
                    "C4", "C#4", "D4", "D#4", "E4", "F4", "F#4", "G4", "G#4", "A4", "A#4", "B4",
                    "C5", "C#5", "D5", "D#5", "E5", "F5", "F#5", "G5", "G#5", "A5", "A#5", "B5",
                    "C6", "C#6", "D6", "D#6", "E6", "F6", "F#6", "G6", "G#6", "A6", "A#6", "B6",
                    "C7", "C#7", "D7", "D#7", "E7", "F7", "F#7", "G7", "G#7", "A7", "A#7", "B7",
                    "C8", "C#8", "D8", "D#8", "E8", "F8", "F#8", "G8", "G#8", "A8", "A#8", "B8",
                    "C9", "C#9", "D9", "D#9", "E9", "F9", "F#9", "G9"}

        Public ValNoteCubase As New List(Of String) From {
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

        ' Attributs
        ' *********
        Public part As New List(Of String)
        ' Propriétés
        ' **********
        ' Piste
        Public NumPiste As Byte = 0
        Public NomPiste As String = String.Empty
        ' 
        Public PositCours As Integer = 0
        Public Début As Integer ' exprimé en nombre de doubles croches (= 0 si départ immédiat, ensuite on part toujours en début d'une autre mesure jamais au milieu d'une mesure)
        Public intervAccent As Integer = 4
        Public valeurAccent As Integer = 10
        Public actifAccent As Boolean = False
        Public NbEvts As Double = -1 '

        ' MIDI
        Public Marq As New List(Of String)
        Public Mute As New List(Of Boolean)
        Public Motifs As New List(Of String)
        Public Durée As New List(Of Double) 'Public DuréeNote As Double = SN ' DuréeNote
        Public Octave As New List(Of Integer)
        Public Volume As Byte = 0
        Public Dyn As New List(Of Byte)
        Public PRG As New List(Of Integer) ' on utilise integer pour gérer off=-1 - piano acc par défaut
        Public Canal As Byte = 0
        Public PAN As New List(Of Byte)
        Public Accent As New List(Of String)
        Public Souche As New List(Of Byte)
        Public Delay As New List(Of Boolean)
        Public DébutSouche As New List(Of Boolean)
        Public Retard As New List(Of Byte)
        Public Start As Integer
        Public SoonDelayed As Boolean ' indique qu'un décalage de +1 sur le poisition des notes a déjà eu lieu sur la piste
        '
        Public LongDerNote As New List(Of Integer)
        Public DelayAvant As Boolean = False
        '
        ' Autres Variables 
        Private SilenceDéjàCompt As Boolean = False
        Public Répétition As Integer
        Public PrésenceNotes As Boolean = False
        '
        Public DbgTabNotes As New List(Of DbgParNotes)


        '
        ' Constructeur
        ' ************
        Public Sub New(NomPiste As String, NumPiste As Byte, Canal As Byte)
            Me.NomPiste = Trim(NomPiste)
            Me.NumPiste = NumPiste
            Me.Canal = Canal ' No canal = No Piste
        End Sub


        Public Sub AddPRG(Position As String, ValPRG As String)
            Dim numPRG As Integer = Convert.ToInt32(ValPRG)
            If numPRG <> -1 Then
                Me.part.Add("PRG" + " " + Trim(Str(Me.NumPiste)) + " " + Trim(Str(Me.Canal)) + " " + Trim(Position) + " " + Trim(Str(numPRG)))
            End If
        End Sub
        Public Sub AddCTRL(Position As String, NumCTRL As String, ValCTRL As String, Magnéto As Integer)
            Select Case Convert.ToInt16(NumCTRL)
                Case CVolume, 54 'volume
                    'If Me.Mute(Magnéto) = True Then ValCTRL = 0
                    Me.part.Add("CTRL" + " " + Trim(Str(Me.NumPiste)) + " " + Trim(Str(Me.Canal)) + " " + Trim(Position) + " " + Trim(NumCTRL) + " " + Trim(ValCTRL))

                Case CPAN
                    Me.part.Add("CTRL" + " " + Trim(Str(Me.NumPiste)) + " " + Trim(Str(Me.Canal)) + " " + Trim(Position) + " " + Trim(NumCTRL) + " " + Trim(ValCTRL))
                Case Else
            End Select
        End Sub
        ''' <summary>
        ''' AddNotes : création des notes générées par HYPERARP. Cette méthode intègre aussi la mise à jour des valeurs d'automation.
        ''' </summary>
        ''' <param name="motif">Liste des notes à ouer selon le motif choisi</param>
        ''' <param name="Octave">Valeur octave à raouter selon le paramètre de la piste HyperARp</param>
        ''' <param name="Dyn">Valeur de la Dynamique selon le paramètre de la piste HyperARp</param>
        ''' <param name="Durée">Durée de la note selon le paramètre de la piste HyperARp. Dans les "Perso" la durée n'esst prise en compte : elle est remplacée par la longueur de la note</param>
        ''' <param name="Chiff">Chiffrage de l'accord en cours</param>
        ''' <param name="Delay">Retard de 1 double croche selon le paramètre de la piste HyperARp</param>
        ''' <param name="Accent">Accent sur 1er temps, ou 1er temps et 3e temps. Paramètre globale à l'appli</param>
        ''' <param name="NBoucle">Nombre de boucles. = 0 si pas de boucle</param>
        Public Sub AddNotes(Motif_ As String, motif As String, Octave As Integer, Dyn As Byte, Durée As Double, Chiff As String, Delay As Boolean, Accent As String, NBoucle As Integer, NRepet As Integer)
            Dim i, j As Integer
            Dim tbl() As String = {String.Empty}
            Dim tbl1() As String = {String.Empty}
            Dim Position As Integer
            Dim Position1 As Integer
            Dim durée1 As Double
            Dim durée2 As Double
            Dim FinTraitement As Integer

            Dim Der_NumAcc As Integer
            Dim Dyn1 As Byte
            Dim T_valeurNote() As Byte = {0, 0, 0, 0, 0, 0, 0, 0}
            Dim T_() As String = {""}


            'décalage de 1 double croche
            tbl = Split(motif) ' lecture des notes générées
            FinTraitement = UBound(tbl) ' index tbl de la dernière note dans la mesure

            Der_NumAcc = Convert.ToString(Form1.Det_NumDerAccord())
            Position = Me.PositCours ' Positcours est la position de début  de la note en nb de db croches dans le morceau - Position : est la colonne en cours traité
            Position1 = Me.PositCours
            For i = 0 To FinTraitement ' 
                'For j = 0 To T_valeurNote.Count - 1
                'T_valeurNote(j) = 0
                'Next

                '**************
                '* AUTOMATION *
                '**************
                ECR_Automation(Me.NumPiste.ToString, Me.Canal.ToString, Position, DivisionMes, NBoucle, 1)
                '
                '************************************
                ' * CALCUL DES PARAMETRES DES NOTES *
                ' ***********************************

                If InStr(tbl(i), "-") <> 0 Then
                    ' PERSO : présence Note avec durée (par exemple C3-4), ou silence avec durée (SIL-16)
                    ' *****
                    tbl1 = Split(tbl(i), "-") 'PERSO : lecture de paramètre de la note Perso
                    durée1 = Convert.ToInt16(tbl1(2)) ' PERSO lecture de la durée de la note Perso
                    '
                    T_(0) = tbl1(0) ' PERSO : lecture de la note
                    T_valeurNote(0) = CalcOctave(Octave, ValNoteCubase.IndexOf(T_(0))) ' PERSO : Prise  en compte du pramètre "Octave" et traduction de la note en chiffre
                    '
                    Position = Position1 + Convert.ToInt16(tbl1(1))     ' tbl1(1) = N° Colonne dans le Perso
                    '
                    If i = FinTraitement Then Me.PositCours = Me.PositCours + (16 * NRepet) '

                Else 'ARPEGES et ACCORDS
                    ' *****************
                    Position = Me.PositCours
                    durée1 = QNvSN * Durée
                    T_ = tbl(i).Split("&")
                    For j = 0 To T_.Count - 1
                        T_valeurNote(j) = CalcOctave(Octave, ValNoteCubase.IndexOf(T_(j)))
                    Next
                    Me.PositCours = Me.PositCours + durée1 ' on positionne ici Me.PositCours pour la prochaine note, la position de ma présente note est dans "position' 
                End If
                '
                ' **********************
                ' * ECRITURE DES NOTES *
                ' **********************
                '
                Dyn1 = CalcDyn(Dyn, Val(Accent), Position) ' calcul de la dynamique pour les notes accentuées

                For j = 0 To T_.Count - 1
                    '  
                    If Not FiltreUnisson(Convert.ToString(Position), Convert.ToString(T_valeurNote(j)), Motif_) Then
                        '                                              piste                         canal                                        note                               Position                  durée                           velo
                        Me.part.Add("Note" + " " + Convert.ToString(Me.NumPiste) + " " + Convert.ToString(Me.Canal) + " " + Convert.ToString(T_valeurNote(j)) + " " + Convert.ToString(Position) + " " + Convert.ToString(durée1) + " " + Convert.ToString(Dyn1))
                    End If

                    ' *************
                    '  NOTES VIEW *
                    ' *************
                    If Me.Répétition = 0 Then
                        Dim a As New DbgParNotes
                        If T_.Count - 1 > 0 Then ' si full acoord
                            a.note = "chrd"
                        Else
                            a.note = Convert.ToString(ValNoteCubase.Item(T_valeurNote(j)))
                        End If
                        a.durée = Convert.ToString(durée1)
                        a.dyn = Convert.ToString(Dyn1)
                        a.canal = Convert.ToString(Me.Canal)
                        a.numPiste = Convert.ToString(Me.NumPiste)
                        a.position = Convert.ToString(Position)
                        Me.DbgTabNotes.Add(a)
                    End If
                Next                '
                ' Ecrire les CTRL sur le reste de la longueur
                If durée1 - 1 > 0 Then
                    ECR_Automation(Me.NumPiste.ToString, Me.Canal.ToString, Position + 1, DivisionMes, NBoucle, durée2)
                End If
                '

            Next i
        End Sub
        ' Fonction de filtrage des unissons dans la génération des arpèges
        ' ***************************************************************
        Private Function FiltreUnisson(position As String, note As String, Motif_ As String) As Boolean
            Dim a As String = Nothing
            Dim b As Boolean = False
            Dim tbl() As String = Motif_.Split()
            '
            If Form1.FiltreUni.Checked Then
                If Not Motif_Accord(Motif_) Then ' on ne traite pas le filtrage des unissons pour les accords
                     If DicoUnisson.TryGetValue(position, a) Then
                        If Not PrésenceNote(a, note) Then
                            DicoUnisson(position) = Trim(a) + "-" + Trim(note)
                            b = False
                        Else
                            b = True
                        End If
                    Else
                        DicoUnisson(position) = Trim(note)
                        b = False
                    End If
                End If
            End If
            Return b
        End Function
        Private Function Motif_Accord(Motif_ As String) As Boolean
            Dim tbl() As String = Motif_.Split()
            Dim b As Boolean = False

            If tbl.Count > 1 Then
                If tbl(1) = "Chord" Then ' on ne filtre pas les unissons sur les accords
                    b = True
                End If
            End If
            Return b
        End Function
        Function PrésenceNote(LNote As String, note As String) As Boolean
            Dim b As Boolean = False
            Dim tbl() As String
            tbl = LNote.Split("-")
            For i = 0 To tbl.Count - 1
                If Trim(tbl(i)) = Trim(note) Then
                    b = True
                End If
            Next
            Return b
        End Function


        ''' <summary>
        ''' ECR_Automation : écriture des valeurs des courbes d'automation
        ''' </summary>
        ''' <param name="NPiste">N° de lapiste : = n° de canal pour le moment</param>
        ''' <param name="Canal">N° de canal MIDI</param>
        ''' <param name="Position">Position de la colonne courante en nombre de doubles croches. Cette position tient compte du bouclage de la lecture</param>
        ''' <param name="DivisionMes">Nombre de divisions par mesures : 16 en 4/4 / 24 en 12/8</param>
        ''' <param name="Nboucle">Nombre de boucles. = 0 si pas de boucles.</param>
        Sub ECR_Automation(NPiste As String, Canal As String, Position As Integer, DivisionMes As Integer, Nboucle As Integer, Durée As Integer)
            Dim i As Integer
            Dim i1 As Integer = Convert.ToInt16(Canal)
            Dim valeurCTRL As String ' valeur du ctrl pour la colonne j
            Dim a As String = ""
            'Dim DébutNote As Integer
            Dim Début As String
            'Dim LongueurPart As Integer = (Form1.Terme.Value - Form1.Début.Value) + 1 ' en nombre de mesures
            Dim colonne As Integer = Position ' Position : position en doubles croches dans le fichier MIDI - La 1ere valeur de position est toujours = à 0
            ' Dim colonne1 As Integer = Position - (Nboucle * (LongueurPart * DivisionMes)) + ((Convert.ToInt16(Form1.Début.Value) - 1) * DivisionMes)

            ' Calcul de la colonne (colonne1) des CTRL dans l'automation
            ' **********************************************************
            ' colonne1 est la position correspondant à colonne (position dans le fichier MIDI) dans l'automation 
            Dim LongueurPart As Integer = (Form1.Terme.Value - Form1.Début.Value) + 1           ' Longueur de la partie joué entre les locateurs 'Début et 'Fin' en nombre de mesures
            '
            Dim DeltaRetir As Integer = Nboucle * (LongueurPart * DivisionMes)                  ' Delta à retirer quand l'exécution HyperARp est en Bouclage en nombre double croches
            Dim DeltaAjout As Integer = (Convert.ToInt16(Form1.Début.Value) - 1) * DivisionMes  ' Delta à ajouter en fonction du locator 'Début' en nombre de doubles croches 

            'Dim colonne1 As Integer = Position - (Nboucle * (LongueurPart * DivisionMes)) + ((Convert.ToInt16(Form1.Début.Value) - 1) * DivisionMes) + 1 ' colonne1 va servir à chercher la valeur du CTRL dans l'automation - on met +1 car les schémas d'automation commencent à colonne 1 et non à la colonne 0
            Dim colonne1 As Integer = Position - DeltaRetir + DeltaAjout + 1
            Position -= 1
            For i = 0 To Durée - 1
                colonne += 1
                Position += 1
                For i2 = 0 To nbCourbes - 1
                    valeurCTRL = Form1.Automation1.ValCtrl(i1, i2, colonne1)
                    If valeurCTRL <> Module1.TValPréced2(i1, i2) And Trim(valeurCTRL) <> "-1" Then ' pas besoin de répéter plusieurs fois la même valeur (surtout le 0)
                        If Form1.Automation1.Canaux(i1).Mute.Checked = True Then
                            If Form1.Automation1.ListCCAct.Item(i1).LCCAct.Item(i2).Checked Then ' on n'écrit l'expression que si le Check est true
                                NumPiste = Trim(NPiste)
                                Canal = Trim(Canal)
                                'DébutNote = (j - 1) - ((Form1.Début.Value - 1) * DivisionMes)
                                Début = Position.ToString 'Convert.ToString(DébutNote + (boucle * LongueurPart))
                                Me.part.Add("CTRL" + " " + Trim(NumPiste) + " " + Trim(Canal) + " " + Trim(Début) + " " + Convert.ToString(Form1.Automation1.Det_CTRL(i2)) + " " + Trim(valeurCTRL))
                                Module1.TValPréced2(i1, i2) = valeurCTRL
                            End If
                        End If
                    End If
                Next i2
            Next i
        End Sub

        Public Sub AddMarq(marq As String, PisteNonMute As Byte)
            Me.part.Add("MRQ" + " " + marq + " " + Convert.ToString(Me.PositCours))
        End Sub
        Public Sub AddAcc(Chiff As String, Numacc As Integer) ' As String
            Dim a As String
            Chiff = Chiff.Replace(" ", "")
            a = "Acc" + " " + Chiff + " " + Convert.ToString(Me.NumPiste) + " " + Convert.ToString(Numacc) + " " + Convert.ToString(Me.PositCours)
            Me.part.Add(a)
        End Sub
        Private Function CalcOctave(Octave As Integer, valeurNote As Integer) As Byte
            Dim OctNote As Integer = 0
            CalcOctave = CByte(valeurNote)

            OctNote = valeurNote + Octave
            If Not (OctNote < 0 Or OctNote > 127) Then
                CalcOctave = CByte(OctNote)
            End If
            'End If
        End Function
        Public Function Det_PositionPrésente() As Integer
            Dim i As Integer
            Dim tbl() As String

            Det_PositionPrésente = 0
            i = IndexNotePrécédente() ' note précédente
            If i <> -1 Then ' avant le 1er enregistrement dans part count = 0 donc Me.part.Count - 1 = -1
                tbl = Split(part(i))
                Det_PositionPrésente = Val(tbl(4))
            End If
        End Function
        Public Function Det_Position(Delay As Boolean) As Integer
            Dim i As Integer
            Dim SommeSilences As Integer = 0
            Dim tbl() As String

            Det_Position = 0
            i = IndexNotePrécédente() ' note précédente
            If i <> -1 Then ' avant le 1er enregistrement dans part count = 0 donc Me.part.Count - 1 = -1 --> Det_Position = 0
                tbl = Split(part(i))
                Det_Position = Val(tbl(4)) + Val(tbl(5)) '+ SommeSilences 'Position=position derniere note+Durée dernière note+ mesures de silences précédentes
            Else ' position 1ere note
                Det_Position = 0 ' + SommeSilences
                If Delay Then Det_Position = 1 ' + SommeSilences
            End If
            '
            Me.PositCours = Det_Position
        End Function

        Function IndexNotePrécédente() As Integer
            Dim i As Integer
            Dim tbl() As String
            IndexNotePrécédente = -1
            For i = Me.part.Count - 1 To 0 Step -1
                tbl = Split(part(i))
                If Trim(tbl(0)) = "Note" Then
                    IndexNotePrécédente = i
                    Exit For
                End If
            Next i
        End Function
        Public Sub SetAccent(valeur As Integer, interv As Integer)
            Me.valeurAccent = valeur
            Me.intervAccent = interv
            Me.actifAccent = True
        End Sub
        Public Function GestAccent(valeurDyn As Integer) As Byte
            Dim DynNote As Integer

            GestAccent = CByte(valeurDyn)
            Me.NbEvts = Me.NbEvts + 1
            If actifAccent Then
                If Me.NbEvts Mod Me.intervAccent = 0 Then
                    DynNote = valeurDyn + Me.valeurAccent
                    If Not (DynNote < 0 Or DynNote > 127) Then
                        GestAccent = CByte(DynNote)
                    End If
                End If
            End If
        End Function

    End Class
    '
    ' *********************************************************
    ' *                                                       *
    ' *                FIN DE DELA CLASSE PISTE               *
    ' *                                                       *
    ' *********************************************************


    ' *********************************************************
    ' *                                                       *
    ' *                DEBUT DU MODULE MAINARP                *
    ' *                                                       *
    ' *********************************************************

    ' ***************************
    ' * VARIABLES ET CONSTANTES *
    ' ***************************
    ' Valeur de durée de notes
    Public Const WN = 4
    Public Const HN = 2
    Public Const QN = 1
    Public Const EN = 0.5
    Public Const SN = 0.25
    '
    Public Const RN = 4
    Public Const BL = 2
    Public Const NR = 1
    Public Const CR = 0.5
    Public Const DC = 0.25

    ' Valeurs de dynamique
    Public Const FFF = 120
    Public Const FF = 100
    Public Const F = 85
    Public Const MF = 70
    Public Const MP = 60
    Public Const P = 50
    Public Const PP = 25
    Public Const PPP = 10

    ' Controleurs MIDI
    Public Const CMod = 1
    Public Const CVolume = 7
    Public Const CPAN = 10

    Public Const QNvSN = 4 ' nombre de double croches dans une noire
    'Public Const TN = 0.125

    ' Liste des notes MIDI
    ' ********************
    ' Remarques sur utilisation
    '   Récupératuin du n0 de note MIDI
    '   Dim i As Byte
    '   i = ValNote.IndexOf("F9")

    '
    Public Arrangement1 As New Partition("4/4", "HyperArp")
    Public LesPistes As New List(Of Piste) ' les objets LesPistes sont créées à la fin de Pistes_Création dans form1
    Public part As New List(Of String)
    Dim nbMesuresUtiles As Integer
    Dim listPad As New List(Of PadNote) ' utilisé seulement pour le motif Pad T,3 et 5
    '
    Dim Numérateur As Integer
    Dim Dénominateur As Integer
    Dim DivisionMes As Integer

    ' NOTES
    ' Syntaxe des notes dans .part de "Les Pistes"
    '     | "Note" | NumPiste | Canal | Valeur Note | début | durée | Vélocité |
    ' ex. | "Note" |     0    |   0   |      72     |   16  |   16  |    90    | 
    ' Les paramètres 'début et 'durée' sont exprimés en nombre de double croches
    ' dans l'ex. la note est une ronde commençant à la mesure 2 (le début commence à 0).

    ' CTRL
    ' Syntaxe des controleurs dans .part de "Les Pistes"
    '     | "CTRL" | NumPiste | Canal | début | N° CTRL | Valeur |
    ' ex. | "CTRL" |     6    |   6   |  0    |   7     |   16   |

    ' PRG
    ' Syntaxe des Programmes dans .part de "Les Pistes"
    '     | "PRG" | NumPiste | Canal | début | N° PRG |
    ' ex. | "PRG" |     6    |   6   |  0    |   7    |

    ''' <summary>
    ''' CalculArp : méthode de génération des notes MIDI pour un fichier MIDI ou pour le scheduler MIDI
    ''' </summary>
    ''' <param name="Midifi">Indique s'il est nécessaire de générer un fichier MIDI</param>
    Public Sub CalculArp(Midifi As Boolean)
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim ind As Integer
        Dim PositFin As String = ""
        Dim MagnétoPréced As Integer
        Dim Boucle As Integer = Det_Répet(Midifi) '
        Dim LongDerNote(0 To nb_PistesVar - 1) As Integer ' pour les pistes Variations
        Dim motif_acc1 As String = String.Empty
        '
        ' Variables de récoltes des info en provenance des variations
        ' ***********************************************************
        Dim Mute As Boolean
        Dim Motif As String = ""
        Dim Durée As Double 'Public DuréeNote As Double = SN ' DuréeNote
        Dim Octave As Integer
        Dim Volume As Byte
        Dim Dyn As Byte
        Dim Accent As String
        Dim PRG As Integer
        'Dim Canal As Byte
        Dim PAN As Byte
        'Dim DébutSouche As Boolean
        Dim Delay As Boolean
        '
        Dim VolumeEnvoyé As Boolean = False

        DicoUnisson.Clear()

        ' Tableaux de récoltes des infos d'entrée en provenance de grid2
        ' **************************************************************
        Dim NotesAcc() As String = {String.Empty} ' notes des accords
        Dim RépetAcc() As String = {String.Empty} ' répétition des accords
        Dim ChiffAcc() As String = {String.Empty} ' chiffrage des accords
        Dim NumAcc() As String = {String.Empty} ' N° des accords (N° colonne contenant l'accord dans Grid2)
        Dim MarqAcc() As String = {String.Empty} ' marqueurs
        Dim MagnétoAcc() As String = {String.Empty} ' Numéros des magnéto

        ' Métrique
        ' ********
        Numérateur = Det_Numérateur(Form1.Métrique.Text)
        Dénominateur = Det_Dénominateur(Form1.Métrique.Text)
        DivisionMes = Det_DivisionMes()

        ' Ajustement des locators
        ' ***********************
        'locator Terme
        Dim loc As Integer = Form1.Det_NumDerAccord()

        If Form1.Terme.Value > loc Then
            Form1.Terme.Value = AjustLocat("Terme", Form1.Terme.Value)
            Form1.Terme.BackColor = Color.Gold
        Else
            Form1.Terme.BackColor = Color.White
        End If
        Form1.Terme.Refresh()
        '
        ' Liste de récoltes des infos d'entrée en provenance de grid2
        ' ***********************************************************
        Dim listnotesAcc As String = Form1.Récup_notesAcc2 ' ok
        Dim listRépet As String = Form1.Récup_Répet2 ' ok
        Dim listChiffAcc As String = Form1.Récup_Acc
        Dim listNumAcc As String = Form1.Récup_NumAcc ' ok
        Dim listMRQ As String = Form1.Récup_MRQ ' ok
        Dim listMagnéto As String = Form1.Récup_NumMagnéto 'ok

        '
        ' Mise à jour des infos d'entrée en provenance de grid2
        ' *****************************************************
        Array.Clear(TValPréced2, 0, TValPréced2.Length) ' POur Maj AUtomation : écriture de 0 dans toutes les cellules du tableau

        NotesAcc = listnotesAcc.Split(",")
        RépetAcc = listRépet.Split(",")
        ChiffAcc = listChiffAcc.Split(",")
        NumAcc = listNumAcc.Split(",")
        MarqAcc = listMRQ.Split(",")
        MagnétoAcc = listMagnéto.Split(",")
        '
        ' INITIALISATIONS
        ' ***************
        RAZ_Lespistes()
        Arrangement1.NumAccords.NumAcc.Clear()
        Arrangement1.NumAccords.LectEcr = True

        ' Init de laPiste Accords
        Dim TraitAcc(0 To UBound(NumAcc)) As Boolean ' pour piste Accords
        For i = 0 To UBound(TraitAcc)
            TraitAcc(i) = False
        Next i
        '
        For i = 0 To nb_PistesVar - 1 ' 

            LesPistes.Item(i).NomPiste = "Piste" + Convert.ToString(i + 1)
            LesPistes.Item(i).Start = Convert.ToInt16(Form1.PDébut - 1)
            LesPistes.Item(i).NumPiste = i
            LesPistes.Item(i).Canal = i
            LesPistes.Item(i).part.Clear()
            LesPistes.Item(i).Marq.Clear()
            LesPistes.Item(i).Motifs.Clear()
            LesPistes.Item(i).Durée.Clear()
            LesPistes.Item(i).Motifs.Clear()
            LesPistes.Item(i).Volume = 0
            LesPistes.Item(i).Mute.Clear()
            LesPistes.Item(i).Octave.Clear()
            LesPistes.Item(i).PRG.Clear()
            LesPistes.Item(i).Dyn.Clear()
            LesPistes.Item(i).PAN.Clear()
            LesPistes.Item(i).Accent.Clear()
            LesPistes.Item(i).Souche.Clear()
            LesPistes.Item(i).Retard.Clear()
            LesPistes.Item(i).Delay.Clear()
            LesPistes.Item(i).DébutSouche.Clear()
            LesPistes.Item(i).LongDerNote.Clear()
            LesPistes.Item(i).DbgTabNotes.Clear()
            LesPistes.Item(i).PositCours = 0
            LesPistes.Item(i).DelayAvant = False
            MagnétoPréced = -1
            '
            For j = 0 To nb_Variations - 1
                LesPistes.Item(i).Durée.Add(Det_Durée(Form1.Récup_Durée(i, j)))
                LesPistes.Item(i).Motifs.Add(Form1.Récup_Motif(i, j))
                'LesPistes.Item(i).Volume.Add(Convert.ToByte((Form1.Récup_Volume(i, j))))
                LesPistes.Item(i).Mute.Add(Convert.ToBoolean(Form1.Récup_Mute(i, j)))
                LesPistes.Item(i).Octave.Add(Convert.ToInt16(Form1.Récup_Octave(i, j)))
                LesPistes.Item(i).PRG.Add(Convert.ToInt16(Form1.Récup_PRG(i, j))) '
                LesPistes.Item(i).Dyn.Add(Form1.Récup_Dyn(i, j))
                LesPistes.Item(i).Accent.Add(Trim(Form1.Récup_Accent(i, j)))
                LesPistes.Item(i).PAN.Add(Convert.ToByte(Form1.Récup_PAN(i, j))) ' = CByte(Val(otbl10(i)))

                LesPistes.Item(i).Souche.Add(Convert.ToByte(Form1.Récup_Souche(i, j)))
                LesPistes.Item(i).Retard.Add(Convert.ToByte(Form1.Récup_Retard(i, j)))
                LesPistes.Item(i).Delay.Add(Convert.ToBoolean(Form1.Récup_Delay(i, j)))
                LesPistes.Item(i).DébutSouche.Add(Convert.ToBoolean(Form1.Récup_DébutSouche(i, j)))
            Next j
            '
            ' TRAITEMENT DES PISTES
            ' *********************
            For k = 0 To Boucle '
                LesPistes.Item(i).Répétition = k
                For j = 0 To UBound(NotesAcc) ' ici on tourne sur les accords contenus dans NotesACC - un accords par mesure
                    If Arrangement1.NumAccords.LectEcr Then Arrangement1.NumAccords.NumAcc.Add(NumAcc(j)) ' liste des N° d'accord pour affichage pendant l'écoute
                    ind = Convert.ToInt16(MagnétoAcc(j)) ' magnéto=variation
                    '
                    ' MAJ pour la mesure k courante et pour pour la variation courante (ind)
                    ' **********************************************************************
                    ' 1 - Maj des variables propres à la variation courante (ind)
                    '     *******************************************************
                    Mute = LesPistes.Item(i).Mute(ind)
                    Motif = LesPistes.Item(i).Motifs(ind)
                    Durée = LesPistes.Item(i).Durée(ind)
                    Octave = LesPistes.Item(i).Octave(ind)
                    Dyn = LesPistes.Item(i).Dyn(ind)
                    Accent = LesPistes.Item(i).Accent(ind)
                    PRG = LesPistes.Item(i).PRG(ind)
                    PAN = LesPistes.Item(i).PAN(ind)
                    '
                    ' 2 - Maj particulière pour la gestion des delay pour la variation courante (ind)
                    '     ***************************************************************************
                    Delay = LesPistes.Item(i).Delay(ind)
                    Posit = Convert.ToString((LesPistes.Item(i).PositCours))
                    If Delay And LesPistes.Item(i).DelayAvant = False Then
                        LesPistes.Item(i).DelayAvant = True
                        LesPistes.Item(i).PositCours = LesPistes.Item(i).PositCours + 1
                    End If
                    '
                    ' 3 - Maj particulière pour la variable Début (Souche) pour la variation courante (ind)
                    '     *********************************************************************************
                    Souche = LesPistes.Item(i).Souche(ind) - 1
                    '
                    ' ECRITURES pour la mesure k courante
                    ' ***********************************
                    ' 1 - Ecriture des Contrôleurs et Marqueurs
                    '     *************************************
                    Posit = Convert.ToString(LesPistes.Item(i).PositCours)
                    'If (Posit = "0") And (ind = 0) And (Form1.Remote.Checked) Then
                    'LesPistes.Item(i).AddCTRL(Posit, 54, 64, 0) ' AddCTRL :  cVolume :valeur du CTRL ->7 otbl4(i) = valeur du volume
                    'End If
                    If (Posit = "0" Or (Delay And (Posit = "1"))) And Form1.Mix.AutorisVol.Checked = True Then ' le volume est marqué une fois seul fois au début de piste et de la 1ere boucle (quans lalecture est répétée).
                        Volume = Form1.Mix.PisteVolume.Item(i).Value 'Form1.Récup_Volume(i) ' 
                        If Form1.Récup_VolumeActif(i) = False Then Volume = 0 ' système de Mute de la table de mixage
                        LesPistes.Item(i).AddCTRL(Posit, Str(CVolume), Volume, ind) ' AddCTRL :  cVolume :valeur du CTRL ->7 otbl4(i) = valeur du volume

                    End If
                    If ind <> MagnétoPréced And Mute Then ' attention MUTE = true signifie que la case cochée et donc que la piste est active (c'est Historique)
                        ' PRG
                        LesPistes.Item(i).AddPRG(Posit, PRG)
                        ' Pan
                        LesPistes.Item(i).AddCTRL(Posit, Str(CPAN), PAN, ind) ' AddCTRL pour PAN
                        'End If
                    End If
                    ' Marq
                    If i = 0 And Trim(MarqAcc(j)) <> "" Then
                        LesPistes.Item(i).AddMarq(MarqAcc(j), 0) ' AddMarq
                    End If
                    '
                    ' 2 - Ecriture des notes de la piste i conditionnées par la variation d'index   ind ( variation a été aussi appelé magnéto)
                    '     *******************************************************************************************************************
                    MagnétoPréced = ind
                    'If (NumAcc(j) >= Form1.PDébut) And (NumAcc(j) <= Form1.PFin) Then
                    If i = 0 Then ' on met a jours les accords que lors du passage sur le canal 1
                        LesPistes.Item(i).AddAcc(ChiffAcc(j), Convert.ToString(NumAcc(j))) ' AddAcc : indication de l'accord dans la piste MIDI
                    End If
                    If Mute = True Then ' si la variable MUTE = True alors la piste est active
                        If Trim(NotesAcc(j)) <> "" Then
                            ' Calcul du motif et MAJ de ces notes dans la piste MIDI
                            motif_acc1 = Det_Motif(i, NotesAcc(j), "4/4", Val(RépetAcc(j)), Durée, Motif, Souche, Delay, ChiffAcc(j)) '
                            LesPistes.Item(i).PrésenceNotes = False
                            If Trim(motif_acc1) <> "" Then
                                LesPistes.Item(i).AddNotes(Motif, motif_acc1, Octave, Dyn, Durée, ChiffAcc(j), Delay, Accent, k, Val(RépetAcc(j))) ' AddNotes : par1:motif, par2:dynamique,par3 : numéro de l'accord,par4 : dernier accord de la piste(true/false)
                                LesPistes.Item(i).PrésenceNotes = True
                            End If
                        End If
                    Else
                        LesPistes.Item(i).PositCours = LesPistes.Item(i).PositCours + (Val(RépetAcc(j)) * (QNvSN * RN)) ' 
                    End If
                    'End If
                Next j

            Next k
            Arrangement1.NumAccords.LectEcr = False
            ' écriture de "FIN"
            If Mute = True Then
                'PositFin = Det_FIN(NumAcc(UBound(NumAcc)), Boucle) '
                PositFin = Det_FIN2(Boucle) '
                LesPistes.Item(i).part.Add("FIN" + " " + Convert.ToString(LesPistes.Item(i).NumPiste) + " " + Convert.ToString(PositFin))
            End If
            '
            Form1.JaugeProgres(20)
        Next i
        '
        ' Syntaxe des notes dans .part de "Les Pistes"
        '     | "Note" | NumPiste | Canal | Valeur Note | début | durée | Vélocité |
        ' ex. | "Note" |     6    |   6   |      72     |   16  |   16  |    90    | 
        ' Les paramètres 'début et 'durée' sont exprimés en nombre de double croches
        ' dans l'ex. la note est une ronde commençant à la mesure 2

        ' Piano Roll et Batterie
        ' **********************

        Dim aa As String = "-"
        Dim b, c As Boolean
        Dim tbl() As String
        Dim tbl1() As String
        Dim NumDerAcc As Integer = Form1.Det_NumDerAccord
        Dim f As Integer = Form1.Det_NumDerAccord
        Dim RépetDerAcc As Integer = Val(Form1.Grid2.Cell(3, f).Text)
        Dim Répéter As Boolean = False
        If Boucle > 0 Then Répéter = True

        For j = 0 To nb_PianoRoll - 1
            LesPistes.Item(j + 6).part.Clear() ' LesPistes s'indexent avec j+6
            LesPistes.Item(j + 6).DbgTabNotes.Clear()
            ' 
            If Form1.listPIANOROLL(j).PMute = True Then ' si =true c'est non muet
                b = Form1.PIANOROLLChargé(j) 'PIANOROLL s'index avec j
                c = Form1.listPIANOROLL(j).PMute
                aa = Form1.listPIANOROLL(j).PListNotes(Répéter, Boucle, Form1.Début.Value, Form1.Terme.Value, NumDerAcc) ' Form1.Terme.Value
                LesPistes.Item(j + 6).PrésenceNotes = Form1.listPIANOROLL(j).PPresNotes
            End If

            If Trim(aa) <> "" Then
                tbl = aa.Split("-")
                With LesPistes.Item(j + 6)

                    For i = 0 To UBound(tbl)
                        .part.Add(tbl(i))
                        '
                        ' *******
                        ' debug *
                        ' *******
                        tbl1 = tbl(i).Split()
                        Dim a As New DbgParNotes
                        If tbl1(0) = "Notes" Then
                            a.numPiste = Trim(tbl1(1))
                            a.canal = Trim(tbl1(2))
                            a.note = Trim(.ValNoteCubase.Item(Val(tbl1(3))))
                            a.position = Trim(tbl1(4))
                            a.durée = Trim(tbl1(5))
                            a.dyn = Trim(tbl1(6))
                            LesPistes.Item(j + 6).DbgTabNotes.Add(a)
                        End If
                    Next i
                    'Next k
                End With
            End If
            ' écriture de "FIN"
            LesPistes.Item(j + 6).part.Add("FIN" + " " + Convert.ToString(LesPistes.Item(j + 6).NumPiste) + " " + Convert.ToString(PositFin))
            '
            Form1.JaugeProgres(20)
        Next j

        ' Traitement du DRUM EDIT
        ' ***********************
        j = 13
        aa = Form1.Drums.PListNotes(Répéter, Boucle, Form1.Début.Value, Form1.Terme.Value, NumDerAcc)

        If Trim(aa) <> "" Then
            LesPistes.Item(j).PrésenceNotes = Form1.Drums.PPresNotes
            tbl = aa.Split("-")
            With LesPistes.Item(j)
                'For k = 0 To Boucle
                For i = 0 To UBound(tbl)
                    .part.Add(tbl(i))
                    ' *******
                    ' debug *
                    ' *******
                    tbl1 = tbl(i).Split()
                    Dim a As New DbgParNotes
                    If tbl1(0) = "Notes" Then
                        a.numPiste = Trim(tbl1(1))
                        a.canal = Trim(tbl1(2))
                        a.note = Trim(.ValNoteCubase.Item(Val(tbl1(3))))
                        a.position = Trim(tbl1(4))
                        a.durée = Trim(tbl1(5))
                        a.dyn = Trim(tbl1(6))
                        LesPistes.Item(j).DbgTabNotes.Add(a)
                    End If
                Next
            End With
            Form1.JaugeProgres(20)
        End If
        'Form1.JaugeProgres(10)
        '
        ' CREATION DU FICHIER MIDI
        ' ************************
        Dim Ob As New NbPistesUtiles

        If Midifi Then
            Dim tbl2(nbMesures) As String
            Dim Midifile1 As New MidifileX(96, Ob.Nb, Module1.nbMesures) '
            Midifile1.AddTempo(Form1.Tempo.Value.ToString)
            Midifile1.AddMétrique(Arrangement1.Métrique)
            Midifile1.AddNomFichier(Arrangement1.NomFichier)
            Dim flagCtrl As Boolean = Form1.Récup_ExportCTRL
            Dim PsteMidF As Integer = -1
            ' Ajout des éléments MIDI
            k = 0 ' pour comptage du nombre de pistes
            For i = 0 To nb_TotalPistes - 1 ' 
                If Ob.TPisteUtil(i) Then
                    PsteMidF = PsteMidF + 1
                    Midifile1.AddNomPiste(PsteMidF, Det_NomPisteMIDI2(i)) ' "HyperArp : " + Convert.ToString(i + 1)
                    For j = 0 To LesPistes.Item(i).part.Count - 1
                        If LesPistes.Item(i).PrésenceNotes Then
                            tbl2 = Split(LesPistes.Item(i).part(j))
                            Select Case tbl2(0)
                                Case "Note"
                                    '                      piste      canal          note        debut        durée         velo
                                    Midifile1.AddNote(PsteMidF, Val(tbl2(2)), Val(tbl2(3)), Val(tbl2(4)), Val(tbl2(5)), Val(tbl2(6)))
                                Case "CTRL" ' comprend tous les controles Volume, Panoralique etc.
                                    If Form1.ExporterCTRL = True Then 'Form1.ExportCTRLMenu.Checked Then
                                        Midifile1.AddCTRL(PsteMidF, Val(tbl2(2)), Val(tbl2(3)), Val(tbl2(4)), Val(tbl2(5)))
                                    End If
                                Case "PRG"
                                    If Form1.ExporterCTRL = True Then 'Form1.ExportCTRLMenu.Checked Then
                                        Midifile1.AddProgram(PsteMidF, Val(tbl2(2)), Val(tbl2(3)), Val(tbl2(4)))
                                    End If
                                Case "MRQ"
                                    Midifile1.AddMarqueur(tbl2(1), Convert.ToInt32(tbl2(2)))
                            End Select
                        End If
                    Next j
                End If
            Next i
            Midifile1.ConstruiredMidFile() ' construction du fichier MIDI
        End If
    End Sub
    Function Det_NomPisteMIDI2(ii As Integer) As String ' N° de pistes commençant par 0
        Dim a As String

        If LangueIHM = "fr" Then
            a = "Canal"
        Else
            a = "Channel"
        End If

        Select Case ii
            Case 0
                Return "HyperArp" + " " + a + " " + "1"
            Case 1
                Return "HyperArp" + " " + a + " " + "2"
            Case 2
                Return "HyperArp" + " " + a + " " + "3"
            Case 3
                Return "HyperArp" + " " + a + " " + "4"
            Case 4
                Return "HyperArp" + " " + a + " " + "5"
            Case 5
                Return "HyperArp" + " " + a + " " + "6"
            Case 6, 7, 8, 9, 10, 11, 12
                Return Form1.TabControl4.TabPages(ii - 5).Text
            Case Else ' la pite de batterie est la piste 13
                If Module1.LangueIHM = "fr" Then
                    Return "Batterie canal 10"
                Else
                    Return "Drums channel 10"
                End If


        End Select
    End Function

    ' Syntaxe des notes dans .part de "Les Pistes"
    '     |   0    |     1    |   2   |      3      |   4   |   5   |     6    |
    '     +--------+----------+-------+-------------+-------+-------+----------+
    '     | "Note" | NumPiste | Canal | Valeur Note | début | durée | Vélocité |
    ' ex. | "Note" |     6    |   6   |      72     |   16  |   16  |    90    | 
    ' Les paramètres 'début et 'durée' sont exprimés en nombre de double croches
    ' dans l'ex. la note est une ronde commençant à la mesure 2
    '
    ' ***************************************************
    '
    ' METHODES DU MODULE MAIN
    '
    ' ***************************************************
    Sub TransfPad(pst As Integer)
        Dim tbl1() As String
        Dim tbl2() As String
        Dim listPad As New List(Of String)
        Dim i As Integer = -1
        Dim j As Integer
        Dim k As Integer
        Dim a As String
        Dim b As String
        Dim c As String

        Do
            i = i + 1
            a = LesPistes.Item(pst).part.Item(i)
            tbl1 = a.Split()
            If tbl1(0) = "Note" Then
                k = tbl1(5)
                For j = i + 1 To LesPistes.Item(pst).part.Count - 1
                    b = LesPistes.Item(pst).part.Item(j)
                    tbl2 = b.Split()
                    If tbl2(0) = "Note" Then
                        If tbl1(3) = tbl2(3) Then
                            k = k + Val(tbl2(5))
                            tbl1(5) = Convert.ToString(k)
                        Else
                            c = Join(tbl1, " ")
                            listPad.Add(c)
                            i = j '- 1
                            Exit For
                        End If
                    Else
                        If tbl2(0) = "FIN" Then
                            c = Join(tbl1, " ")
                            listPad.Add(c)
                        End If
                        listPad.Add(b)
                        i = j ' - 1
                        'Exit For
                    End If
                Next
            Else
                listPad.Add(a)
            End If

        Loop Until i >= LesPistes.Item(pst).part.Count - 1
        '
        LesPistes.Item(pst).part.Clear()
        LesPistes.Item(pst).part = listPad
    End Sub
    Function CalcDyn(Dyn As Byte, Accent As Integer, Position As Integer) As Byte
        Dim Multiple As Byte = 16
        If Form1.Accent1_3 Then Multiple = 8

        CalcDyn = Dyn
        If Form1.IsMultiple(Position, Multiple) Then
            CalcDyn = Dyn + Accent
            If CalcDyn > 127 Then CalcDyn = 127
        End If
    End Function
    Function AjustLocat(Typ As String, Locat1 As Integer) As Integer
        Dim locat As Integer = Locat1

        If Trim(Form1.Grid2.Cell(2, locat).Text) = "" Then
            Do
                locat = locat - 1
            Loop Until Trim(Form1.Grid2.Cell(2, locat).Text) <> "" Or locat <= 0
        End If
        '
        If Typ = "Début" And locat = 0 Then
            locat = Locat1
            Do
                locat = locat + 1
            Loop Until Trim(Form1.Grid2.Cell(2, locat).Text) <> "" Or locat >= nbMesures
        End If
        Return locat
    End Function
    Function Det_FIN(Num_DerAcc As Integer, Boucle As Integer) As Integer
        Dim j As Integer
        Dim k As Integer = 0
        Dim PositFin As Integer = 1

        If Num_DerAcc > Form1.PFin Then
            Num_DerAcc = Form1.PFin
        End If
        '
        For j = Val(Form1.PDébut) To Num_DerAcc
            If Trim(Form1.Grid2.Cell(2, j).Text) <> "" Then
                k = k + Val(Form1.Grid2.Cell(3, j).Text)
            End If
        Next
        PositFin = Convert.ToString((k * (Boucle + 1) * 16) + 16) ' à modifier si on passe à un autre type de métrique que 4/4
        Return PositFin
    End Function
    Function Det_FIN2(Boucle As Integer)
        Dim k As Integer
        Dim PositFin As Integer

        k = (Val(Form1.PFin) - Val(Form1.PDébut)) + 1
        If Form1.PFin = Form1.Det_NumDerAccord() Then
            k = k + (Val((Form1.Grid2.Cell(3, Form1.PFin).Text)) - 1)
        End If
        PositFin = Convert.ToString((k * (Boucle + 1) * 16) + 16) ' à modifier si on passe à un autre type de métrique que 4/4
        Return PositFin
    End Function
    Function Det_PistePrésente(pst As Integer) As Boolean
        Dim i As Integer
        Dim j = 0
        For i = 0 To LesPistes.Item(pst).Mute.Count - 1
            If LesPistes.Item(pst).Mute(i) = True Then
                j = j + 1
            End If
        Next i
        Det_PistePrésente = Not (j = 4)
    End Function
    Sub RAZ_Lespistes()
        For i = 0 To nb_PistesVar - 1
            LesPistes.Item(i).Durée.Clear()
            LesPistes.Item(i).Motifs.Clear()
            LesPistes.Item(i).Volume = 0
            LesPistes.Item(i).Mute.Clear()
            LesPistes.Item(i).Octave.Clear()
            LesPistes.Item(i).PRG.Clear()
            LesPistes.Item(i).Dyn.Clear()
            LesPistes.Item(i).PAN.Clear()
            LesPistes.Item(i).part.Clear()
        Next i
    End Sub
    Sub EcrireFIN(pst As Integer, ListR As String, Boucle As Integer)
        Dim i As Integer
        Dim p As Integer = 0
        Dim tbl() As String
        Dim LongueurPart As Integer

        tbl = Split(ListR, ",")
        'For i = 0 To UBound(tbl)
        'p = (p + (Val(tbl(i)) * 16))
        'Next
        'p = (p * (Boucle + 1)) + 16
        LongueurPart = Val(tbl(1)) - Val(tbl(0))
        For i = 0 To LongueurPart
            p = p + 16
        Next
        p = (p * (Boucle + 1)) + 16 'prise en compte de "Répéter"
        LesPistes.Item(pst).part.Add("FIN" + " " + Convert.ToString(LesPistes.Item(pst).NumPiste) + " " + Convert.ToString(p))
    End Sub
    Function Det_PremPisteNonMute(pst As Integer, Magnéto As Integer) As Integer
        Dim i, j As Integer
        j = 0
        Det_PremPisteNonMute = -1 ' 17 = hors n° canal midi = hors N° de piste --> dans ce cas il n'y a pas de piste non mutée
        For i = 0 To nb_BlocPistes - 1
            If LesPistes.Item(pst).Mute(Magnéto) = False Then
                Det_PremPisteNonMute = i
                Exit For
            End If
        Next i
    End Function
    Function Det_NbPiste() As Integer ' non utilisée
        Dim i, j As Integer
        j = 0
        For i = 0 To nb_PistesVar - 1
            'If LesPistes.Item(i).Mute = False Then
            j = j + 1
            'End If
        Next i
        Det_NbPiste = j
    End Function
    Function Det_Répet(Midifi As Boolean) As Integer
        Det_Répet = 0
        'If Form1.Récup_Répéter = True Then
        If Not (Midifi) Then Det_Répet = Convert.ToUInt16(Form1.NbBoucles.Text) - 1
        'End If
    End Function
    Public Function Det_Motif(Pst As Integer, NotesAcc As String, Métrique As String, Répétition As Integer, Durée As Double, Motif As String, Début As Integer, Delay As Boolean, ChiffAcc As String) As String
        Dim numGrid As Integer
        Dim tbl() As String
        Dim a As String
        'Dim otbl9() As String = {String.Empty}
        'Dim listMotifs As String = Form1.Récup_Motif
        'otbl9 = listMotifs.Split(",")
        Det_Motif = ""
        If Trim(NotesAcc) <> "" And Trim(Répétition) <> 0 Then
            Select Case Motif
                Case "ArpMotif 1"
                    'ArpMotif1(ListNotes As String, metrique As String, nbMesures As Integer, duree As Double)
                    Det_Motif = ArpMotif1(NotesAcc, Métrique, Répétition, Durée, Début) ' par1:notes accord/gammes,par2:métrique,par3:répétition
                Case "ArpMotif 2"
                    Det_Motif = ArpMotif2(NotesAcc, Métrique, Répétition, Durée, Début) ' par1:notes accord/gammes,par2:métrique,par3:répétition
                Case "ArpMotif 3"
                    Det_Motif = ArpMotif3(NotesAcc, Métrique, Répétition, Durée, Début) ' par1:notes accord/gammes,par2:métrique,par3:répétition
                Case "ArpMotif 4"
                    Det_Motif = ArpMotif4(NotesAcc, Métrique, Répétition, Durée, Début) ' par1:notes accord/gammes,par2:métrique,par3:répétition
                Case "1_Répétition 1", "Répétition V1", "Répétition T", "Répétition 3", "Full Chord", "No3rd Chord", "No5th Chord" ' accords
                    Det_Motif = Répé(0, NotesAcc, Métrique, Répétition, Durée, 0, Motif, ChiffAcc)
                Case "3_Répétition 2", "Répétition V2" 'répétition de la tierce
                    Det_Motif = Répé(1, NotesAcc, Métrique, Répétition, Durée, 0, Motif, ChiffAcc)
                Case "5_Répétition 3", "Répétition V3" ' répétition de la quinte
                    Det_Motif = Répé(2, NotesAcc, Métrique, Répétition, Durée, 0, Motif, ChiffAcc)

                Case "Perso 1", "Perso 2", "Perso 3", "Perso 4", "Perso 5", "Perso 6"
                    tbl = Split(Motif)
                    numGrid = Convert.ToInt16(tbl(1)) - 1
                    Det_Motif = Perso(numGrid, NotesAcc, Métrique, Répétition, Début)

                Case "Motif1 Chord"
                    tbl = Split(Motif)
                    Select Case tbl(0)
                        Case "Motif1"
                            a = Répé(0, NotesAcc, Métrique, Répétition, Durée, 0, Motif, ChiffAcc)
                            Det_Motif = Motif1_Chord(a)
                        Case "Motif2"
                    End Select

                Case Else
                    Det_Motif = ArpMotif1(NotesAcc, Métrique, Répétition, Durée, Début) ' par1:notes accord/gammes,par2:métrique,par3:répétition

            End Select
        End If
    End Function
    Class PadNote
        Public note As String = ""
        Public NumMes As Integer = -1
        Public NbFois As Integer = -1
    End Class
    Sub PadCalc(DegréID As Integer, ChiffAcc As String, NumAcc As Integer)

        Dim list2 As New List(Of PadNote)
        Dim ind As Integer
        Dim V As String
        Dim tbl() As String
        m = NumAcc
        V = Trim(TableNotesAccordsZ(m, 1, 1))

        If V <> "" Then
            ind = Det_IndexDsVoicing(DegréID, V, ChiffAcc) ' détermination de l'index de degréID dans un voicing V
            tbl = V.Split()
            Dim a As New PadNote With {
                .note = tbl(ind),
                .NumMes = m
            }
            listPad.Add(a)
        End If

    End Sub

    ' ******************************************************
    ' * Det_IndexDsVoicing : détermination de l'index d'un *
    ' * degré dans un voicing V                             *
    ' * DegréID = 0 --> Tonique                            *
    ' * DegréID = 1 --> Tierce                             *
    ' * DegréID = 2 --> Quinte                             * 
    ' * V : chaine de caractères contenant le voicing      *
    ' ******************************************************
    Function Det_IndexDsVoicing(DegréID As Integer, V As String, ChiffAcc As String) As Integer
        Dim notesAcc As String
        Dim tbl1() As String
        Dim tbl2() As String
        Dim Note As String

        Dim i As Integer
        Dim a As String

        notesAcc = Form1.Det_NotesAccord3(Trim(ChiffAcc), "#")
        tbl1 = notesAcc.Split("-")
        Note = tbl1(DegréID)

        tbl2 = V.Split()

        ' suppression des octaves dans le Voicing
        For i = 0 To UBound(tbl2)
            If InStr(tbl2(i), "#") <> 0 Then
                a = Microsoft.VisualBasic.Left(tbl2(i), 2)
            Else
                a = Microsoft.VisualBasic.Left(tbl2(i), 1)
            End If
            If a = Note Then Exit For
        Next
        Return i

    End Function

    Public Function Det_Durée(D As String) As Double
        Select Case D
            Case "WN", "RN"
                Det_Durée = 4
            Case "HN", "BL"
                Det_Durée = 2
            Case "QN", "NR"
                Det_Durée = 1
            Case "EN", "CR"
                Det_Durée = 0.5
            Case "SN", "DC"
                Det_Durée = 0.25
            Case Else
                Det_Durée = 1
        End Select
    End Function
    Public Function Det_Dyn(D As String) As Byte
        Select Case Trim(D)
            Case "FFF"
                Det_Dyn = 120
            Case "FF"
                Det_Dyn = 100
            Case "F"
                Det_Dyn = 85
            Case "MF"
                Det_Dyn = 70
            Case "MP"
                Det_Dyn = 60
            Case "P"
                Det_Dyn = 50
            Case "PP"
                Det_Dyn = 25
            Case "PPP"
                Det_Dyn = 10
            Case Else
                Det_Dyn = 100
        End Select
    End Function
    Public Function Det_Dyn2(D As String) As Byte

        Select Case Trim(D)
            Case "16"
                Det_Dyn2 = 127
            Case "15"
                Det_Dyn2 = 119
            Case "14"
                Det_Dyn2 = 111
            Case "13"
                Det_Dyn2 = 103
            Case "12"
                Det_Dyn2 = 95
            Case "11"
                Det_Dyn2 = 87
            Case "10"
                Det_Dyn2 = 81
            Case "9"
                Det_Dyn2 = 73
            Case "8"
                Det_Dyn2 = 65
            Case "7"
                Det_Dyn2 = 57
            Case "6"
                Det_Dyn2 = 49
            Case "5"
                Det_Dyn2 = 41
            Case "4"
                Det_Dyn2 = 38
            Case "3"
                Det_Dyn2 = 30
            Case "2"
                Det_Dyn2 = 22
            Case "1"
                Det_Dyn2 = 14
            Case Else
                Det_Dyn2 = 80 'Convert.ToByte(Form1.PianoRoll1.ListDynF1.Text)
        End Select
    End Function

    Function NbNotes(metrique As String, nbMesures As Integer, duree As String) As Integer
        NbNotes = nbMesures
        Select Case metrique
            Case "4/4"
                Select Case duree
                    Case WN, RN
                        NbNotes = nbMesures
                    Case HN, BL
                        NbNotes = nbMesures * 2
                    Case QN, NR
                        NbNotes = nbMesures * 4
                    Case EN, CR
                        NbNotes = nbMesures * 8
                    Case SN, DC
                        NbNotes = nbMesures * 16
                End Select
        End Select
    End Function

    Sub Création_MidiFile1(CheminFichierText As String, NbPistes As Integer)
        Dim Ligne As String
        Dim tbl As Object
        Dim i As Integer = 0

        '   
        ' création du fichier MIDI
        ' ************************
        Dim Midifile1 As New MidifileX(96, NbPistes, nbMesuresUtiles)
        'a = My.Application.Info.DirectoryPath
        'a = a + "\" + "fichemidiHypVoice.txt"
        'FileOpen(1, a, OpenMode.Input)
        Dim fileReader = My.Computer.FileSystem.OpenTextFileReader(CheminFichierText)

        Try
            Do Until fileReader.Peek = -1
                Ligne = fileReader.ReadLine()
                i = i + 1 ' compteur pour debugguer
                tbl = Split(Ligne, ";")
                Select Case tbl(0)
                    Case "NomFichier"
                        Midifile1.AddNomFichier(tbl(1))
                    Case "Tempo"
                        Midifile1.AddTempo(Val(tbl(1)))
                    Case "NomPiste" '             piste    nompiste                     
                        Midifile1.AddNomPiste(Val(tbl(1)), tbl(2))
                    Case "Programme" '           piste       Canal       Début      n° PRG
                        Midifile1.AddProgram(Val(tbl(1)), Val(tbl(2)), Val(tbl(3)), Val(tbl(4)))

                    Case "Controleur" '     piste       Canal        Début       N° ctrl      Valeur ctrl
                        Midifile1.AddCTRL(Val(tbl(1)), Val(tbl(2)), Val(tbl(3)), Val(tbl(4)), Val(tbl(5)))
                    Case "Note"
                        '                      piste      canal          note        debut        durée         velo
                        Midifile1.AddNote(Val(tbl(1)), Val(tbl(2)), Val(tbl(3)), Val(tbl(4)), Val(tbl(5)), Val(tbl(6)))
                    Case "Marqueur"
                        If Trim(tbl(1)) <> "" Then
                            Midifile1.AddMarqueur(tbl(1), Val(tbl(2)))
                        End If
                    Case "Texte"
                        Midifile1.AddTexte(Val(tbl(1)), tbl(2), Val(tbl(3)))
                    Case "Métrique"
                        Midifile1.AddMétrique(Trim(tbl(1)))
                End Select
            Loop
            fileReader.Close()
            Midifile1.ConstruiredMidFile() ' construction du fichier MIDI
            '
        Catch ex As Exception
            If LangueIHM = "fr" Then
                MessageHV.PTitre = "Avertissement"
                MessageHV.PContenuMess = "Détection d'un erreur dans procédure : " + "Création_MidiFile1" + Constants.vbCrLf + "Message  : " + ex.Message
                MessageHV.PTypBouton = "OK"
            Else
                MessageHV.PTitre = "Warning"
                MessageHV.PContenuMess = "Détection d'un erreur dans procédure : " + "Création_MidiFile1" + Constants.vbCrLf + "Message  : " + ex.Message
                MessageHV.PTypBouton = "OK"
            End If
            Cacher_FormTransparents()
            MessageHV.ShowDialog()
        End Try
    End Sub

    Private Function Det_DivisionMes() As Integer
        '
        Select Case Dénominateur
            Case 4
                Det_DivisionMes = Numérateur * 4
            Case 8
                Det_DivisionMes = Numérateur * 2
            Case Else
                Det_DivisionMes = 16
        End Select
    End Function
    Function Det_Numérateur(Métrique As String) As Integer
        Dim tbl() As String = Métrique.Split("/")
        Det_Numérateur = Convert.ToInt16(tbl(0))
    End Function
    Function Det_Dénominateur(Métrique As String) As Integer
        Dim tbl() As String = Métrique.Split("/")
        Det_Dénominateur = Convert.ToInt16(tbl(1))
    End Function
End Module
