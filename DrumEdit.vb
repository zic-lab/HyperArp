Imports System.IO
Imports System.Text
Public Class DrumEdit
    ' PROPRIETES
    ' **********
    Public Property PMute As Boolean
        Get
            Mute = Me.CheckMute.Checked
            Return Mute
        End Get
        Set(ByVal value As Boolean)
            Me.CheckMute.Checked = value
        End Set
    End Property
    Private Langue As String
    Public WriteOnly Property PLangue As String
        Set(ByVal value As String)
            Me.Langue = value
        End Set
    End Property
    Private PresNotes As Boolean
    Public Property PPresNotes As Boolean
        Get
            Return PresNotes
        End Get
        Set(ByVal value As Boolean)
            PresNotes = value
        End Set
    End Property
    Private NbMesu As Integer
    Public WriteOnly Property PNbMesures As Integer
        Set(ByVal value As Integer)
            Me.NbMesu = value
        End Set
    End Property

    Public Métrique As String = "4/4"
    Public WriteOnly Property PMétrique As String
        Set(ByVal value As String)
            Me.Métrique = value
        End Set
    End Property
    Private ListAcc As String
    Public WriteOnly Property PListAcc As String
        Set(ByVal value As String)
            Me.ListAcc = value
            Dim tbl() As String = value.Split("-")
            Me.NbAccords = UBound(tbl) + 1
        End Set
    End Property
    Private ListMarq As String
    Public WriteOnly Property PListMarq As String
        Set(ByVal value As String)
            Me.ListMarq = value
        End Set
    End Property
    Private nbRépétitionMax As Integer = Module1.nbRépétitionMax
    Public WriteOnly Property PnbRépétitionMax As Integer
        Set(ByVal value As Integer)
            Me.nbRépétitionMax = value
        End Set
    End Property
    Private GridOrigine As String
    Public ReadOnly Property PGridOrigine As String
        Get
            Return GridOrigine
        End Get
    End Property
    Private EcrNomPerso_Visible As Boolean = False
    Public ReadOnly Property PEcrNomPerso_Visible As Boolean
        Get
            Return EcrNomPerso_Visible
        End Get
    End Property
    ''' <summary>
    ''' Prpriété retournant une chaine de caractère contenant les notes à jouer d'une piste.
    ''' </summary>
    ''' <param name="Répéter">: boolean-case à cocher oui/non</param>
    ''' <param name="Boucle"> : nombre de réptition</param>
    ''' <param name="Form1_Début"> : Début de la lecture en nombre de mesures: locateur Début de la barre de transport</param>
    ''' <param name="Form1_Fin">: Fin de la lecture en nombres de mesures locateur Terme de la barre de transport</param>
    ''' <param name="NumDerAcc">: n° de mesure du dernier accord dans le projet</param>
    ''' <returns>chaine de caractère contenant les notes à jouer d'une piste</returns>

    Public ReadOnly Property PListNotes(Répéter As Boolean, Boucle As Integer, Form1_Début As Integer, Form1_Fin As Integer, NumDerAcc As Integer) As String
        Get
            Me.ListNotes = Contruction_ListeNotes(Répéter, Boucle, Form1_Début, Form1_Fin, NumDerAcc)
            Return Me.ListNotes
        End Get
    End Property
    Private ListNotes As String = ""
    ' Constantes
    ' **********
    Private Const HautLigne1Grid1 = 61
    Private Const HautLignesGrid1 = 37 '24
    Private Const LargeColNom = 191 '260
    Private Const LargeColNotes = 100
    Private Const NbLignesGrid1 = 51 '13

    ' Fonts
    ' *****
    ReadOnly fnt1 As New System.Drawing.Font("Calibri", 13, FontStyle.Regular)
    ReadOnly fnt2 As New System.Drawing.Font("Calibri", 10, FontStyle.Regular)
    ReadOnly fnt3 As New System.Drawing.Font("Tahoma", 10, FontStyle.Regular)
    ReadOnly fnt4 As New System.Drawing.Font("Calibri", 18, FontStyle.Bold)
    ReadOnly fnt5 As New System.Drawing.Font("Calibri", 12, FontStyle.Regular)
    ReadOnly fnt6 As New System.Drawing.Font("Corbel", 11, FontStyle.Regular)
    ReadOnly fnt7 As New System.Drawing.Font("DejaVu Serif", 13, FontStyle.Regular)
    ReadOnly fnt8 As New System.Drawing.Font("DejaVu Serif", 10, FontStyle.Regular)
    ReadOnly fnt9 As New System.Drawing.Font("Verdana", 8, FontStyle.Regular)

    ' Composants IHM
    ' **************
    '
    Public F2 As New Form ' remarque : l'intégration de F2 dans l'onglet (page) de TabControl4 se fait dans Form1/DRUMS_Création
    '
    Public Panneau1 As New SplitContainer
    Public Panneau2 As New SplitContainer
    Public Grid1 As New FlexCell.Grid ' grille drum edit
    Public Grid2 As New FlexCell.Grid ' grille de time line de positionnement des variation D1..D8
    Dim ComBoVar As New System.Windows.Forms.ComboBox ' liste des présets à choisir dans la time line
    '
    Public CheckMute As New CheckBox ' mute de la piste de batterie
    Public SoloBoutDRM As New Button ' Bouton solo de la piste de batterie
    Dim LVar As New List(Of Button) ' liste des boutons permettant de passer d'une variation à l'autre dans le drum edit
    Private ListTypNote As New System.Windows.Forms.ComboBox ' type de notes à écrire dans DrumEdit : Ronde, blanche, Noire...
    Private LabelTypNote As New Label ' 
    Public ListDynF2 As New System.Windows.Forms.ComboBox ' valeur des dynamiques affectée aux notes lors de l'écriture
    Private LabelDyn As New Label
    Private ToolTip1 As New ToolTip
    Private DockButton As New Button

    Private ListNomGS As New System.Windows.Forms.ComboBox   '  entrer instrument par son Nom
    Private ListNotesGS As New System.Windows.Forms.ComboBox '  entrer instrument par sa note
    Private EntrerNomPerso As New TextBox ' changer le nom d'un instrument
    Private BoutInit As New Button  ' bouton de réinitialisation du drum edit
    Private BoutClear As New Button ' bouton effacement des notes
    Private BoutCopier As New Button ' bouton effacement des notes
    Private BoutColler As New Button ' bouton effacement des notes
    Private ListPrésets As New ListBox ' liste de présets
    Private SauvPrésets As New Button ' bouton de sauvegarde de présets
    Public NomPréset As New TextBox ' changer le nom d'un instrument
    Public NoteClic As New CheckBox ' pour envoi d'une note sur clic dans Grid1

    ' Menu Fichier/Quitter
    ' ********************
    Private Menu1 As New System.Windows.Forms.MenuStrip()
    Private Fichier As New System.Windows.Forms.ToolStripMenuItem()
    Private Quitter As New System.Windows.Forms.ToolStripMenuItem()
    Private Edition As New System.Windows.Forms.ToolStripMenuItem()
    Private Couper As New System.Windows.Forms.ToolStripMenuItem()
    Private Copier As New System.Windows.Forms.ToolStripMenuItem()
    Private Coller As New System.Windows.Forms.ToolStripMenuItem()
    Private Séparateur1 As New System.Windows.Forms.ToolStripSeparator()
    Private Annuler As New System.Windows.Forms.ToolStripMenuItem()
    Private Séparateur2 As New System.Windows.Forms.ToolStripSeparator()
    Private Supprimer As New System.Windows.Forms.ToolStripMenuItem()
    '
    ' Menu contextuel
    ' ***************
    Dim MenuContext1 As New ContextMenuStrip
    Dim itemCouper As New ToolStripMenuItem
    Dim itemCopier As New ToolStripMenuItem
    Dim itemColler As New ToolStripMenuItem
    Dim Séparateur3 As New System.Windows.Forms.ToolStripSeparator()
    Dim itemPresets As New ToolStripMenuItem
    Dim LPrests As New List(Of ToolStripMenuItem)




    Class LClip
        Public lstClip As New List(Of String)
    End Class
    Dim ListClipAnnuler As New List(Of LClip) ' annulation à un seul niveau - cette liste est RAZ à chaque sauvegarde de la dernière action
    '
    ' Variables
    ' *********
    Private Canal As Byte = 10
    Private NbMes As Integer
    Dim NbAccords As Integer
    Private Numérateur As Integer
    Private Dénominateur As Integer
    Private DivisionMes As Integer
    Private Enchargement As Boolean = True
    Private ListNotesNoms As New List(Of String)
    Private BufferCCC As String
    Private LPréset As New List(Of String) ' liste des présets A,B,C ... Mise à jour dans Maj_Divers - contient le nom "A" .."H" des boutons
    Public LNomPréset As New List(Of String)
    ReadOnly PrésetsBase As String = "B0/Acoustic Bass Drum" + "-" + "D1/Acoustic Snare" + "-" + "D#1/Hand Clap" + "-" + "F#1/Closed Hi Hat" + "-" + "A#1/Open Hi Hat" + "-" + "A1/Low Tom" + "-" + "D2/High Tom" + "-" + "C#1/Side Stick" + "-" + "D#4/Claves" + "-" + "F2/Ride Bell" + "-" + "G#2/Cowbell" + "-" + "F#2/Tambourine"
    ReadOnly PrésetsBase2 As String = "B0/Acoustic Bass Drum" + "-" + "C1/Bass Drum 1" + "-" + "C#1/Side Stick" + "-" + "D1/Acoustic Snare" + "-" + "D#1/Hand Clap" _
    + "-" + "E1/Electric Snare" + "-" + "F1/Low Floor Tom" + "-" + "F#1/Closed Hi Hat" + "-" + "G1/High Floor Tom" + "-" + "G#1/Pedal Hi Hat" + "-" + "A1/Low Tom" _
    + "-" + "A#1/Open Hi Hat" + "-" + "B1/Low Mid Tom" + "-" + "C2/Hi Mid Tom" + "-" + "C#2/Crash Cymbal 1" + "-" + "D2/High Tom" + "-" + "D#2/Ride Cymbal 1" _
    + "-" + "E2/Chinese Cymbal" + "-" + "F2/Ride Bell" + "-" + "F#2/Tambourine" + "-" + "G2/Splash Cymbal" + "-" + "G#2/Cowbell" + "-" + "A2/Crash Cymbal 2" + "-" + "Bb2/Vibraslap" _
    + "-" + "B2/Ride Cymbal 2" + "-" + "C3/Hi Bongo" + "-" + "C#3/Low Bongo" + "-" + "D3/Mute Hi Conga" + "-" + "D#3/Open Hi Conga" + "-" + "E3/Low Conga" + "-" + "F3/High Timbale" _
    + "-" + "F#3/Low Timbale" + "-" + "G3/High Agogo" + "-" + "G#3/Low Agogo" + "-" + "A3/Cabasa" + "-" + "A#3/Maracas" + "-" + "B3/Short Whistle" + "-" + "C4/Long Whistle" + "-" + "C#4/Short Guiro" _
    + "-" + "D4/Long Guiro" + "-" + "D#4/Claves" + "-" + "E4/Hi Wood Block" + "-" + "F4/Low Wood Block" + "-" + "F#4/Mute Cuica" + "-" + "G4/Open Cuica" + "-" + "G#4/Mute Triangle" + "-" + "A4/Open Triangle"

    Class CNotes
        Public lig As Integer ' ligne
        Public col As Integer ' colonne
        Public dyn As String ' dynamique
        Public Pos As String ' position
    End Class
    Class CPréset
        Public LNotes As New List(Of CNotes)
    End Class
    Dim Préset As New List(Of CPréset)
    '
    ' Couleurs de base du drum edit
    ' *****************************
    ReadOnly DrmC_MarronOrange As Color = Color.FromArgb(209, 130, 11)
    ReadOnly DrmC_BleuFoncé As Color = Color.FromArgb(58, 68, 88)
    ReadOnly DrmC_Gris As Color = Color.FromArgb(223, 226, 225)
    ReadOnly DrmC_Creme As Color = Color.FromArgb(253, 255, 240)
    ReadOnly DrmC_BleuAzur As Color = Color.FromArgb(0, 135, 167)
    ReadOnly DrmC_VioletClair As Color = Color.FromArgb(211, 196, 227)
    ReadOnly DrmC_RougeClair As Color = Color.FromArgb(221, 105, 114)
    ReadOnly DrmC_VertbleuClair As Color = Color.FromArgb(158, 202, 199)
    ReadOnly DrmC_VertOlive As Color = Color.FromArgb(116, 128, 66)
    ReadOnly DrmC_GrisBleu As Color = Color.FromArgb(205, 218, 227)
    ReadOnly DrmC_OrangeOcre As Color = Color.FromArgb(218, 138, 2)
    ReadOnly DrmC_OrangeFoncé As Color = Color.FromArgb(202, 100, 60)
    ReadOnly DrmC_OrangeClair As Color = Color.FromArgb(255, 204, 102)
    ReadOnly DrmC_RougeFoncé As Color = Color.FromArgb(201, 37, 9)
    ReadOnly DrmC_BleuGris As Color = Color.FromArgb(87, 131, 157)
    ReadOnly DrmC_VertFoncé As Color = Color.FromArgb(89, 110, 70)
    ReadOnly DrmC_VertClair As Color = Color.FromArgb(77, 152, 106)
    ReadOnly DrmC_VioletFoncé As Color = Color.FromArgb(79, 64, 125)
    ReadOnly DrmC_Marron As Color = Color.FromArgb(140, 120, 80)
    ReadOnly DrmC_MarronFoncé As Color = Color.FromArgb(76, 54, 30)
    ReadOnly DrmC_JauneTrèsClair As Color = Color.FromArgb(255, 255, 237)
    ReadOnly DrmC_Fauve As Color = Color.FromArgb(176, 138, 81)
    ReadOnly DrmC_Ecru As Color = Color.FromArgb(225, 210, 188)
    '
    Private ReadOnly CoulA As Color = DrmC_BleuGris
    Private ReadOnly CoulB As Color = DrmC_RougeClair
    Private ReadOnly CoulC As Color = DrmC_VertClair
    Private ReadOnly CoulD As Color = DrmC_OrangeOcre
    Private ReadOnly CoulE As Color = DrmC_Marron
    Private ReadOnly CoulF As Color = DrmC_VertFoncé
    Private ReadOnly CoulG As Color = DrmC_RougeFoncé
    Private ReadOnly CoulH As Color = DrmC_VioletFoncé
    Private ReadOnly CoulRien As Color = Color.Tan

    ' Variable pour jouer note sur clic
    ' *********************************
    Private NoteCourante As Byte = 64
    Dim NoteAEtéJouée As Boolean = False
    '
    ' Variable pour ClipBoardPerso
    ' ****************************
    Public Class ClipB
        Public EcartLigne As Integer
        Public EcartCol As Integer
        Public value As String
    End Class
    Dim ClipBPerso As New List(Of ClipB)


    ' Notes
    ' *****
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

    ' Notes séparées par des " " entre note et numéro
    Public ValNoteCubase_2 As New List(Of String) From {
                   "C -2", "C# -2", "D -2", "D# -2", "E -2", "F -2", "F# -2", "G -2", "G# -2", "A -2", "A# -2", "B -2",
                   "C -1", "C# -1", "D -1", "D# -1", "E -1", "F -1", "F# -1", "G -1", "G# -1", "A -1", "A# -1", "B -1",
                   "C 0", "C# 0", "D 0", "D# 0", "E 0", "F 0", "F #0", "G 0", "G# 0", "A 0", "A# 0", "B 0",
                    "C 1", "C# 1", "D 1", "D# 1", "E 1", "F 1", "F# 1", "G 1", "G# 1", "A 1", "A# 1", "B 1",
                    "C 2", "C# 2", "D 2", "D# 2", "E 2", "F 2", "F# 2", "G 2", "G# 2", "A 2", "A# 2", "B 2",
                    "C 3", "C# 3", "D 3", "D# 3", "E 3", "F 3", "F# 3", "G 3", "G# 3", "A 3", "A# 3", "B 3",
                    "C 4", "C# 4", "D 4", "D# 4", "E 4", "F 4", "F# 4", "G 4", "G# 4", "A 4", "A# 4", "B 4",
                    "C 5", "C# 5", "D 5", "D# 5", "E 5", "F 5", "F# 5", "G 5", "G# 5", "A 5", "A# 5", "B 5",
                    "C 6", "C# 6", "D 6", "D# 6", "E 6", "F 6", "F# 6", "G 6", "G# 6", "A 6", "A# 6", "B 6",
                    "C 7", "C# 7", "D 7", "D# 7", "E 7", "F 7", "F# 7", "G 7", "G# 7", "A 7", "A# 7", "B 7",
                    "C 8", "C# 8", "D 8", "D# 8", "E 8", "F 8", "F# 8", "G 8"}
    ''' <summary>
    ''' New : Constructeur du drumedit
    ''' </summary>
    ''' <param name="Canal">N° Canal MIDI du PianoRoll</param>
    Sub New(Canal As Byte)
        Me.Canal = Canal
        Me.DivisionMes = Det_DivisionMes()
        Me.Numérateur = Det_Numérateur(Me.Métrique)
        Me.Dénominateur = Det_Dénominateur(Me.Métrique)
        Maj_Divers()
        Init_ListDesPrésets()
        '
        AddHandler F2.KeyUp, AddressOf F2_KeyUp

        AddHandler ComBoVar.KeyDown, AddressOf ComBoVar_KeyDown
        AddHandler ListTypNote.KeyDown, AddressOf ListTypNote_KeyDown
        AddHandler ListDynF2.KeyDown, AddressOf ListDynF2_Keydown
        AddHandler ListNomGS.KeyDown, AddressOf ListNomGS__KeyDown
        AddHandler ListNotesGS.KeyDown, AddressOf ListNotesGS_Keydown




        'Me.F2.KeyPreview = True 'pour réception des touches F4 et F5 (stop, play (recalcul))
    End Sub

    Private Sub ComBoVar_KeyDown(Sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        e.SuppressKeyPress = True
    End Sub
    Private Sub ListDynF2_Keydown(Sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        e.SuppressKeyPress = True
    End Sub
    Private Sub ListTypNote_KeyDown(Sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        e.SuppressKeyPress = True
    End Sub
    Private Sub ListNotesGS_Keydown(Sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        e.SuppressKeyPress = True
    End Sub
    Private Sub ListNomGS__KeyDown(Sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        e.SuppressKeyPress = True
    End Sub





    Private Sub F2_KeyUp(sender As Object, e As KeyEventArgs)

        ' PLAY, RECALCUL : F5
        ' *******************
        If e.KeyCode = Keys.F5 Then
            If Not Form1.Horloge1.IsRunning Then
                Form1.PlayHyperArp()
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
        End If
    End Sub

    Sub Maj_Divers()
        Dim i As Integer

        LPréset.Add("A")
        LPréset.Add("B")
        LPréset.Add("C")
        LPréset.Add("D")
        LPréset.Add("E")
        LPréset.Add("F")
        LPréset.Add("G")
        LPréset.Add("H")

        ' Init Nom des présets
        ' ********************
        For i = 0 To NbDrumPrésets - 1
            LNomPréset.Add("")
        Next

        ' Init de la liste des Préset (8 présets)
        ' ***************************************
        For i = 0 To NbDrumPrésets - 1
            Dim aa As New CPréset
            Préset.Add(aa)
        Next


        ' Liste clip Annuler des Préset
        ' *****************************
        For i = 0 To Module1.NbDrumPrésets - 1
            ListClipAnnuler.Add(New LClip)
        Next
    End Sub
    Public Sub FocusSurA()
        ' Focus sur le préset "A"
        'LVar.Item(0).Focus()
        NomPréset.Text = LNomPréset.Item(Det_NumPréset(Trim(Grid1.Cell(0, 0).Text)))

    End Sub


    ''' <summary>
    ''' Construction des SplitContainer Panneau1 et Panneau2.
    ''' Panneau1 est vertical et contient la barre d'outils et le Drum Edit
    ''' Panneau2 est horizontal
    ''' Panneau2 porte Panneau1 dans son Panel2
    ''' Panneau2 porte Grid2(TimeLine) dans son Panel1
    ''' Panneau1 : Split Container contenant la barre d'outils et le drum édit
    ''' Split container contenant la time line de position des pattern A..H 
    ''' </summary>
    Sub Construction_SplitContainer()
        F2.Visible = False
        If LangueIHM = "fr" Then
            F2.Text = "Batterie"
        Else
            F2.Text = "DRUMS"
        End If
        F2.Controls.Add(Panneau2) ' ajout du splitcontainer dans le formulaire F2
        ' Panneau1 : Split Container contenant la barre d'outils et le drum édit
        ' Panneau2 : Split container contenant la time line de position des pattern A..H et le splitcontaire1 avec la barre d'outils et le drum edit 
        '
        ' Construction de Panneau2
        ' ************************
        Panneau2.Dock = DockStyle.Fill
        Panneau2.Orientation = Orientation.Horizontal
        Panneau2.SplitterDistance = 100
        Panneau2.Panel1.BorderStyle = BorderStyle.Fixed3D
        Panneau2.Panel1.BackColor = Color.White
        Panneau2.FixedPanel = FixedPanel.Panel1
        Panneau2.IsSplitterFixed = True
        '
        ' Construction de Panneau1  : Split Container contenant la barre d'outils et le drum édit
        ' ***************************************************************************************
        Panneau1.Dock = DockStyle.Fill
        Panneau1.Orientation = Orientation.Vertical

        Panneau1.Panel1.BorderStyle = BorderStyle.FixedSingle
        Panneau1.Panel1.BackColor = Color.White
        '
        Panneau1.Panel1.BorderStyle = BorderStyle.FixedSingle
        Panneau1.Panel2.BorderStyle = BorderStyle.FixedSingle
        '
        Panneau2.Panel2.Controls.Add(Panneau1)
        '
        Panneau1.IsSplitterFixed = True
        Panneau1.SplitterDistance = 84
        Panneau1.Panel1.BackColor = Color.Beige
        '
        ' ajout des composants grid aux panneaux
        ' **************************************
        Panneau1.Panel2.Controls.Add(Grid1) '  drum edit
        Panneau2.Panel1.Controls.Add(Grid2) '  time line


        '
        F2.FormBorderStyle = FormBorderStyle.SizableToolWindow
        F2.ControlBox = False
    End Sub
    '
    ''' <summary>
    ''' Construction du Drum edit : grille d'édition des notes = Grid1
    ''' </summary>
    Private Sub Construction_Grid1()
        Dim i, j As Integer

        Grid1.AutoRedraw = False
        Grid1.Rows = NbLignesGrid1 '24 + Grid1.FixedRows
        Grid1.AllowUserPaste = ClipboardDataEnum.Text

        ' dimensionnement
        Grid1.Dock = DockStyle.Fill 'DockStyle.Fill
        Grid1.SelectionMode = SelectionModeEnum.Free

        '

        Grid1.AllowDrop = False
        'Grid1.AllowUserPaste = ClipboardDataEnum.None
        Grid1.AllowUserReorderColumn = False
        Grid1.AllowUserResizing = False
        Grid1.AllowUserSort = False

        Dim s As New Size With {
            .Width = 1200,
            .Height = HautLigne1Grid1 + HautLignesGrid1 * NbLignesGrid1
        }
        Grid1.Size = s
        Grid1.ScrollBars = ScrollBarsEnum.None
        Grid1.Cols = (Me.DivisionMes) + 4 ' Me.nbRépétitionMax
        Grid1.FixedCols = 1 ' le minimum est 1
        Grid1.FixedRows = 1

        ' 
        Grid1.Row(0).Height = HautLigne1Grid1
        For i = 1 To Grid1.Rows - 1
            Grid1.Row(i).Height = HautLignesGrid1 ' onprend la hauteur pour faire un carré
        Next
        '
        For j = 1 To Grid1.Cols - 1
            Grid1.Column(j).Width = HautLignesGrid1
        Next j
        '
        Grid1.Column(0).Width = 75
        Grid1.Column(1).Width = 1
        Grid1.Column(2).Width = LargeColNom
        Grid1.Column(3).Width = LargeColNotes
        ' couleurs
        Grid1.Range(1, 4, NbLignesGrid1 - 1, 4).BackColor = DrmC_OrangeClair   ' 1er temps NbLignesGrid1 - 1
        Grid1.Range(1, 8, NbLignesGrid1 - 1, 8).BackColor = DrmC_OrangeClair   ' 2e temps
        Grid1.Range(1, 12, NbLignesGrid1 - 1, 12).BackColor = DrmC_OrangeClair ' 3e temps
        Grid1.Range(1, 16, NbLignesGrid1 - 1, 16).BackColor = DrmC_OrangeClair ' 4e temps

        '
        Grid1.BackColorBkg = DrmC_GrisBleu 'Color.White
        Grid1.BackColor1 = Color.Moccasin 'Color.Lavender 'Color.White
        Grid1.BackColor2 = Color.Moccasin ' 'Color.White
        Grid1.SelectionBorderColor = Color.Red
        Grid1.BackColorActiveCellSel = Color.Yellow
        Grid1.BackColorSel = DrmC_BleuFoncé
        Grid1.BackColorFixedSel = Det_BackColorDrum(0) 'Color.Moccasin 'DrmC_BleuFoncé
        Grid1.BackColorFixed = Det_BackColorDrum(0)
        '
        For i = 1 To Grid1.Rows - 1
            For j = 1 To 3
                Grid1.Cell(i, j).BackColor = Color.White
            Next
        Next
        For j = 0 To Grid1.Cols - 1
            Grid1.Cell(0, j).ForeColor = Color.White
        Next
        For i = 0 To Grid1.Rows - 1
            Grid1.Cell(i, 0).ForeColor = Color.White
        Next
        '
        Panneau1.Panel1.BackColor = Color.Beige 'Color.FromArgb(223, 226, 225)
        ' Styles, Titres et Numérotations
        'Dim f As New System.Drawing.Font("Tahoma", 8, FontStyle.Regular)
        Grid1.DefaultFont = fnt2
        If Me.Langue = "fr" Then
            Grid1.Cell(0, 1).Text = "Modèle"
            Grid1.Cell(0, 2).Text = "Nom"
            Grid1.Cell(0, 3).Text = "Notes"
        Else
            Grid1.Cell(0, 1).Text = "Model"
            Grid1.Cell(0, 2).Text = "Name"
            Grid1.Cell(0, 3).Text = "Notes"
        End If
        '
        For j = 4 To Grid1.Cols - 1
            Select Case j
                Case 4
                    Grid1.Cell(0, j).Text = "1"
                Case 8
                    Grid1.Cell(0, j).Text = "2"
                Case 12
                    Grid1.Cell(0, j).Text = "3"
                Case 16
                    Grid1.Cell(0, j).Text = "4"
            End Select
        Next
        '
        Grid1.Cell(0, 0).BackColor = Color.White
        Grid1.Cell(0, 0).ForeColor = Color.Black
        Grid1.Cell(0, 0).Font = fnt4
        Grid1.Cell(0, 0).Text = "A"
        '
        For i = (Grid1.FixedRows) To Grid1.Rows - 1
            For j = 0 To Grid1.Cols - 1
                Grid1.Cell(i, j).Alignment = AlignmentEnum.CenterCenter
                If j = 0 Then
                    Grid1.Cell(i, 0).Text = Convert.ToString(i) ' numérotation de lere colonne
                    Grid1.Cell(1, 0).FontBold = false
                End If
            Next
        Next
        '
        ' Création de la colonne Indicateur de Models
        ' *******************************************
        For i = 1 To Grid1.Rows - 1
            Grid1.Cell(i, 1).BackColor = Color.White
        Next
        Grid1.Cell(1, 1).BackColor = Couleur_ModelDrum

        ' ***************************************************
        ' Intégration des liste de Nom GS et liste de notes *
        ' ************************************************* *
        Grid1.Controls.Add(ListNomGS)
        Grid1.Controls.Add(ListNotesGS)
        Grid1.Controls.Add(EntrerNomPerso)
        ListNomGS.Visible = False
        ListNomGS.DropDownStyle = ComboBoxStyle.DropDownList 'ComboBoxStyle.DropDownList ' pour interdire l'écriture dans le comboblist
        ListNomGS.BringToFront()
        ListNotesGS.Visible = False
        ListNotesGS.BringToFront()
        EntrerNomPerso.Visible = False
        EntrerNomPerso.BringToFront()
        '
        ListNomGS.BackColor = DrmC_JauneTrèsClair
        ListNotesGS.BackColor = DrmC_JauneTrèsClair
        ListNomGS.ForeColor = Color.Black
        ListNotesGS.ForeColor = Color.Black
        EntrerNomPerso.BackColor = DrmC_OrangeFoncé
        EntrerNomPerso.ForeColor = Color.White
        '
        Dim f As New System.Drawing.Font("Calibri", 10, FontStyle.Regular)
        'Grid1.Range(0, 0, Grid1.Rows - 1, Grid1.Cols - 1).Font = f
        Grid1.DefaultFont = f

        AddHandler Grid1.Scroll, AddressOf Grid1_Scroll
        AddHandler Grid1.MouseDown, AddressOf Grid1_MouseDown
        AddHandler Grid1.MouseUp, AddressOf Grid1_MouseUp
        AddHandler Grid1.KeyDown, AddressOf Grid1_Keydown
        AddHandler Grid1.KeyUp, AddressOf Grid1_KeyUp
        AddHandler ListNomGS.SelectedIndexChanged, AddressOf ListNomGS_SelectedIndexChanged
        AddHandler ListNotesGS.SelectedIndexChanged, AddressOf ListNotesGS_SelectedIndexChanged
        AddHandler EntrerNomPerso.KeyPress, AddressOf EntrerNomPerso_KeyPress
        AddHandler BoutInit.MouseClick, AddressOf BoutInit_MouseClick
        AddHandler BoutClear.MouseClick, AddressOf Boutclear_MouseClick
        AddHandler BoutCopier.MouseClick, AddressOf BoutCopier_MouseClick
        AddHandler BoutColler.MouseClick, AddressOf BoutColler_MouseClick
        AddHandler SauvPrésets.MouseClick, AddressOf SauvPrésets_MouseClick
        AddHandler DockButton.MouseClick, AddressOf DockButton_MouseClick
        AddHandler NomPréset.KeyUp, AddressOf NomPréset_KeyUp
        AddHandler ListPrésets.SelectedIndexChanged, AddressOf ListPrésets_SelectedIndexChanged

        '
        '
        ' Construction des Combolist de Grid1 : ListNomGS,ListNotesGS 
        Init_ListNotesNom()
        'Grid1.Range(1, 4, NbLignesGrid1 - 1, Grid1.Cols - 1).Locked = True
        Grid1.Range(1, 1, NbLignesGrid1 - 1, 3).Locked = True ' vérouiller les cellules de Nom Instruments et Notes instrumets
        'Grid1.Range(1, 4, NbLignesGrid1 - 1, Grid1.Cols - 1).Locked = True
        ' init préset de base
        BasePrésets()
        VérouillerGrid1()
        '
        Grid1.AutoRedraw = True
        Grid1.Refresh()
        '
        'VérouillerGrid1()
    End Sub
    Public Sub VérouilerCell()
        Grid1.Range(1, 4, NbLignesGrid1 - 1, Grid1.Cols - 1).Locked = True
    End Sub
    Public Sub DéVérouilerCell()
        Grid1.Range(1, 4, NbLignesGrid1 - 1, Grid1.Cols - 1).Locked = False
    End Sub
    Public Sub DessinDiv()
        Dim i, j As Integer

        ' raz fond
        Grid1.AutoRedraw = False
        For i = 1 To Grid1.Rows - 1
            For j = 4 To Grid1.Cols - 1
                Grid1.Cell(i, j).BackColor = Color.Moccasin
            Next
        Next
        'raz colonnes
        Grid1.Range(1, 4, NbLignesGrid1 - 1, 4).BackColor = DrmC_OrangeClair   ' 1er temps
        Grid1.Range(1, 8, NbLignesGrid1 - 1, 8).BackColor = DrmC_OrangeClair   ' 2e temps
        Grid1.Range(1, 12, NbLignesGrid1 - 1, 12).BackColor = DrmC_OrangeClair ' 3e temps
        Grid1.Range(1, 16, NbLignesGrid1 - 1, 16).BackColor = DrmC_OrangeClair ' 4e temps
        Grid1.AutoRedraw = True
        Grid1.Refresh()
    End Sub
    Public Function Sauv_LInst() As String
        Dim i As Integer
        Dim a As String = "ListDrumInst,"

        For i = 1 To Grid1.Rows - 1
            a = a + Trim(Grid1.Cell(i, 2).Text) + "/" + Trim(Grid1.Cell(i, 3).Text) + ","
        Next
        a = Microsoft.VisualBasic.Left(a, Len(a) - 1) ' retrait de la dernière ","
        Return a
    End Function
    Public Function Sauv_LPresetNotes(NPrés As Integer) As String
        Dim a As String = "ListDrumNotes," + Convert.ToString(NPrés) + ","

        ' Pour un Préset,la chaine contient
        ' ------------------------                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                -----------------------------------------------------------------
        ' - ListDrumNotes,N°Préset+ " " +Ligne+ " " +Colonne+ " " +Vélocité+ " " +Position+","
        ' --------------------------------------------------------------------------------------
        If Préset.Item(NPrés).LNotes.Count <> 0 Then
            For Each aa As CNotes In Préset.Item(NPrés).LNotes
                If IsNumeric(aa.dyn) Then
                    If (Val(aa.dyn) >= 0) And (Val(aa.dyn) <= 127) Then
                        a = a + Convert.ToString(aa.lig) + " " + Convert.ToString(aa.col) + " " + Convert.ToString(aa.dyn) + " " + Convert.ToString(aa.Pos) + ","
                    End If
                End If
            Next
        End If

        a = Microsoft.VisualBasic.Left(a, Len(a) - 1) ' retrait de la dernière ","
        Return a
    End Function
    Public Function Sauv_LTimeLPres() As String

        Dim Fr As Integer = Grid2.Rows - 1
        Dim Fc As Integer = 1
        Dim Lr As Integer = Grid2.Rows - 1
        Dim Lc As Integer = Grid2.Cols - 1

        Dim a As String = "ListTimeLPres,"
        For i = Fr To Lr
            For j = Fc To Lc
                a = a + Trim(Grid2.Cell(i, j).Text) + " "
            Next
        Next

        Return Trim(a)
    End Function
    Public Function Sauv_NomPréset() As String
        Dim a As String = "NomPréset"

        For Each b As String In LNomPréset
            a = a + "," + b
        Next
        Return Trim(a)
    End Function
    Public Sub Charger_ListTimeLPres(ligne As String)
        Dim Fr As Integer = Grid2.Rows - 1
        Dim Fc As Integer = 1
        Dim Lr As Integer = Grid2.Rows - 1
        Dim Lc As Integer = Grid2.Cols - 1
        Dim tbl1() As String = Split(ligne, ",")
        Dim tbl2() As String = tbl1(1).Split()
        Dim i, j As Integer
        Dim k As Integer = 0
        For i = Fr To Lr
            For j = Fc To Lc
                If k <= UBound(tbl2) Then
                    Grid2.Cell(i, j).Text = tbl2(k)
                    Grid2.Cell(i, j).BackColor = Det_BackColorDrum2(Trim(tbl2(k)))
                    k += 1
                End If
            Next
        Next
    End Sub

    Public Sub Charger_LNomPréset(Ligne As String)
        Dim i As Integer
        Dim tbl() As String = Ligne.Split(",")
        ' mise à jour de la liste
        For i = 1 To tbl.Count - 1
            LNomPréset.Item(i - 1) = tbl(i)
        Next
        ' mise à du nom du préset sélectionné
        NomPréset.Text = Det_NumPréset(Trim(Grid1.Cell(0, 0).Text))
    End Sub


    ''' <summary>
    ''' ChargerPersoP : chargement d'un Preset perso dans la boite à ryhtm  HyperArp
    ''' </summary>
    ''' <param name="Chaine">Chaine de caratère contenant le Nom du preset, ses notes et ses instruments. Caractère de séparation : ";"</param>
    Public Sub ChargerPréset(Chaine As String)
        Dim i, j As Integer
        Dim ii As Integer
        Dim v As String
        Dim tbl1() As String = Trim(Chaine).Split(";")
        Dim tbl2() As String
        Dim tbl3() As String
        Dim mess, titre As String

        If Module1.LangueIHM = "fr" Then
            mess = "Charger le préset ?"
            titre = "Demande de confirmation"
        Else
            mess = "Load preset ?"
            titre = "Confirmation request"
        End If
        ' 
        Cacher_FormTransparents()
        i = MessageBox.Show(mess, titre, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If i = vbYes Then
            Grid1.AutoRedraw = False

            ' Effacer les notes éventuellement présentes
            ' ******************************************
            EffacerTout()
            ' Mise à jour du nom
            ' ******************
            NomPréset.Text = Trim(tbl1(0))
            LNomPréset.Item(Det_NumPréset(Trim(Grid1.Cell(0, 0).Text))) = NomPréset.Text


            ' Mise à jour des notes
            ' *********************
            tbl1(1) = Microsoft.VisualBasic.Left(tbl1(1), Len(tbl1(1)) - 1) ' retrait de la dernière ","
            tbl2 = tbl1(1).Split(",")
            For ii = 0 To tbl2.Count - 1
                tbl3 = Trim(tbl2(ii)).Split()
                i = Convert.ToInt16(tbl3(0))
                j = Convert.ToInt16(tbl3(1))
                v = tbl3(2)
                DéVérouillerGrid1()
                Grid1.Cell(i, j).Text = Trim(v)
                VérouillerGrid1()
            Next
            '
            ' Mise à jour des instruments
            ' ***************************
            'Charger_ListDrumInst(tbl1(2))
            tbl2 = Trim(tbl1(2)).Split(",")
            For ii = 0 To tbl2.Count - 1
                tbl3 = Trim(tbl2(ii)).Split("/")
                'DéVérouillerGrid1()
                Grid1.Cell(ii + 1, 2).Text = tbl3(0)
                Grid1.Cell(ii + 1, 3).Text = tbl3(1)
                'VérouillerGrid1()
            Next
            '
            ' Mise à jour de la liste de Préset
            ' *********************************
            Maj_Préset(Det_NumPréset(Trim(Grid1.Cell(0, 0).Text)))
            '
            Grid1.AutoRedraw = True
            Grid1.Refresh()
        End If
    End Sub
    Public Function Sauv_PrésetConstruct(NPrés As Integer, Nom As String) As String
        Dim a As String = Nom + ";"
        Dim tbl() As String
        ' ----------------------------------
        ' Pour un Préset,la chaine contient
        ' --------------------------------------------------------------------------------------                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                -----------------------------------------------------------------
        ' -  " " +Ligne+ " " +Colonne+ " " +Vélocité+ " " +Position+"," -
        ' --------------------------------------------------------------------------------------
        ' Recherche des notes
        ' *******************
        If Préset.Item(NPrés).LNotes.Count <> 0 Then
            For Each aa As CNotes In Préset.Item(NPrés).LNotes
                a = a + Convert.ToString(aa.lig) + " " + Convert.ToString(aa.col) + " " + Convert.ToString(aa.dyn) + " " + Convert.ToString(aa.Pos) + ","
            Next
        End If
        '
        ' Recherche des Noms des Instruments  et des notes de chaque Instrument  (colonne 2 et 3 dans Drum Edit)
        ' *****************************************************************************************************
        tbl = Trim(a).Split(";")
        If tbl.Length > 1 Then ' signifie a-t-on trouvé des notes dans le For Each précédent ? si non alors on ne va pas plus loin
            a += ";"
            For i = 1 To Grid1.Rows - 1
                a = a + Trim(Grid1.Cell(i, 2).Text) + "/" + Trim(Grid1.Cell(i, 3).Text) + ","
            Next
            'a = Microsoft.VisualBasic.Left(a, Len(a) - 1) ' retrait de la dernière ","
        End If
        a = Microsoft.VisualBasic.Left(a, Len(a) - 1) ' retrait de la dernière ","
        Return a
    End Function
    Public Sub Charger_ListDrumInst(Ninst As Integer, Nom As String)
        Dim tbl() As String = Nom.Split("/")
        'DéVérouillerGrid1()
        Grid1.Cell(Ninst, 2).Text = Trim(tbl(0))
        Grid1.Cell(Ninst, 3).Text = Trim(tbl(1))
        'VérouillerGrid1()
    End Sub
    Public Sub Charger_ListDrumNotes(Ligne As String)


        Dim tbl1() As String = Ligne.Split(",")
        Dim tbl2() As String
        Dim NPrés As Integer = Convert.ToInt16(tbl1(1))
        Préset.Item(NPrés).LNotes.Clear()

        For i = 2 To tbl1.Count - 1
            tbl2 = tbl1(i).Split()
            Dim aa As New CNotes
            aa.lig = Convert.ToInt16(tbl2(0))
            aa.col = Convert.ToInt16(tbl2(1))
            aa.dyn = tbl2(2)
            aa.Pos = tbl2(3)
            '
            Préset.Item(NPrés).LNotes.Add(aa)
        Next
        ' Remarque : ici on charge les notes des 8 présets seulement. Un seul préset est affiché à la fois 
        ' dans le drum edit. 
        ' Au moment du chargement (ouvrir2), on choisit d'afficher par défaut 
        ' le Préset 0 dans le drumedit par appel de Refresh_Drums_Ouvrir
    End Sub

    Public Sub Refresh_Drums_Init()
        Dim i As Integer

        BasePrésets()
        BaseTimeLine()
        EffacerTout()
        '
        ' init choix préset : effacer les notes dans la liste de Préset
        Grid1.Cell(0, 0).Text = "A"
        Grid1.BackColorFixed = Det_BackColorDrum(0)
        For i = 0 To NbDrumPrésets - 1
            Préset.Item(i).LNotes.Clear()
        Next
        ' 
        ChoixPréset(0)
        Ecr_Préset(0)
        '
        NomPréset.Text = ""
        LNomPréset.Clear()
        For i = 0 To NbDrumPrésets - 1
            LNomPréset.Add("")
        Next
        '
        Grid1.Refresh()
        Grid2.Refresh()
    End Sub
    Public Sub Refresh_Drums_Ouvrir()
        ChoixPréset(0)
    End Sub



    Public Sub BasePrésets()
        Dim tbl1() As String = PrésetsBase2.Split("-")
        Dim i As Integer
        'DéVérouillerGrid1()
        For i = 1 To tbl1.Count
            Dim tbl2() As String = tbl1(i - 1).Split("/")

            Grid1.Cell(i, 3).Text = tbl2(0)
            Grid1.Cell(i, 2).Text = tbl2(1)
        Next i
        'VérouillerGrid1()
    End Sub
    Public Sub BaseTimeLine()
        Dim j As Integer

        'DéVérouillerGrid1()
        For j = 1 To Grid2.Cols - 1
            Grid2.Cell(1, j).Text = ""
            Grid2.Cell(2, j).Text = ""
            Grid2.Cell(3, j).Text = "A"
            Grid2.Cell(3, j).BackColor = Det_BackColorDrum(0)
            Grid2.Cell(3, j).ForeColor = Color.White
        Next j
        'VérouillerGrid1()
    End Sub
    Private Sub EffacerTout()
        Dim i, j As Integer '
        'DéVérouillerGrid1()
        For i = 1 To Grid1.Rows - 1
            For j = 4 To Grid1.Cols - 1
                Grid1.Cell(i, j).Text = ""
            Next
        Next
        '
        i = Det_NumPréset(Trim(Grid1.Cell(0, 0).Text))
        Maj_Préset(i)
        'VérouillerGrid1()
    End Sub


    '
    ''' <summary>
    ''' Construction d'une chaîne de caractères contenant toutes les notes/CTRL d'un piano roll
    ''' </summary>
    ''' <param name="Répéter"> : boolean-case à cocher oui/non</param>
    ''' <param name="Boucle"> : nombre de réptition</param>
    ''' <param name="Form1_Début"> : Début de la lecture en nombre de mesures: locateur Début de la barre de transport</param>
    ''' <param name="Form1_Fin">: Fin de la lecture en nombres de mesures locateur Terme de la barre de transport</param>
    ''' <param name="NumDerAcc"> : n° de mesure du dernier accord dans le projet</param>
    ''' <returns>Retourne une chaine de caractéres avec éléments séparés par des blancs </returns>
    Function Contruction_ListeNotes(Répéter As Boolean, Boucle As Integer, Form1_Début As Integer, Form1_Fin As Integer, NumDerAcc As Integer) As String
        Dim nbColonnes As Integer
        Dim DebPart As Integer
        Dim FinPart As Integer
        Dim a As String = ""
        Dim n As String


        Dim _boucle As Integer = 0
        Dim LongueurPart As Integer = (Form1_Fin - Form1_Début) + 1 ' en nombre de mesures
        Dim N_Mes As Integer
        Dim nbColonnesMes As Integer
        Dim Départ As Integer
        Dim Volume As String


        ' NOTES
        ' Syntaxe des notes dans .part de "Les Pistes"
        '     | "Note" | NumPiste | Canal | Valeur Note | début | durée | Vélocité |
        ' ex. | "Note" |     6    |   6   |      72     |   16  |   16  |    90    | 
        ' Les paramètres 'début et 'durée' sont exprimés en nombre de double croches
        ' dans l'ex. la note est une ronde commençant à la mesure 2

        ' CTRL
        ' Syntaxe des controleurs dans .part de "Les Pistes"
        '     | "CTRL" | NumPiste | Canal | début | N° CTRL | Valeur | 
        ' ex. | "CTRL" |     6    |   6   |  0    |   7     |   16   |

        ' PRG
        ' Syntaxe des Programmes dans .part de "Les Pistes"
        '     | "PRG" | NumPiste | Canal | début | N° PRG |
        ' ex. | "PRG" |     6    |   6   |  0    |   7    | 


        If Répéter Then _boucle = Boucle
        nbColonnesMes = NbDivInduit(Me.Dénominateur) * Me.Numérateur ' en 4/4 le résultat donne 16 
        ' Rétablissement de Form1_Fin en fonction du nombre de répétitions de l'acccord qu'il pointe
        If Form1_Fin = NumDerAcc Then
            Form1_Fin = Form1_Fin + Val(Form1.Grid2.Cell(3, Form1_Fin).Text) - 1
        End If
        ' Détermination du DEPART et de la FIN en N° de colonnes de la séquence à jouer(sert à la boucle exécuter plus loin)
        Départ = ((nbColonnesMes) * (Form1_Début - 1)) + 1 ' DEPART en N° de colonnes
        nbColonnes = (nbColonnesMes) * Form1_Fin           ' FIN en N° de colonnes

        '
        ' Détermination du DEPART et de la FIN en N° de mesures de la séquence à jouer : sert à calculer LongueurPart 
        DebPart = Form1_Début   ' DEPART en N° de mesures
        FinPart = Form1_Fin     ' FIN en N° de mesures
        LongueurPart = ((FinPart - DebPart) + 1) * nbColonnesMes ' Longueur de la séquence à jouer en nombre de colonnes
        n = Convert.ToString(Me.Canal) '

        ' Mise à jour du controleur de volume
        If Form1.Mix.AutorisVol.Checked = True Then
            Volume = Convert.ToString(Form1.Récup_Volume(n))
            If Form1.Récup_VolumeActif(n) = False Then Volume = "0" ' gestion du système de Mute de la table de mixage
            '                 N° Piste         N° Canal        Position                   N° CTRL           Valeur CTRL
            a = "CTRL" + " " + Trim(n) + " " + Trim(n) + " " + "0" + " " + Convert.ToString(CVolume) + " " + Volume + "-"
        End If
        ' Mise à jour des Notes
        For nbRépet = 0 To _boucle
            For N_Mes = Form1_Début To Form1_Fin
                a = a + LNotesPréset(N_Mes, Form1_Début, nbColonnesMes, LongueurPart, nbRépet)
            Next
        Next
        If Trim(a) <> "" Then a = Trim(Microsoft.VisualBasic.Left(a, a.Length - 1))

        Return a
    End Function
    Function LNotesPréset(N_Mes As Integer, Form1Début As Integer, nbColonnesMes As Integer, LongueurPart As Integer, nbRépet As Integer) As String
        Dim Canal As String = Me.Canal
        Dim ValeurNote As String
        Dim Pos As String = Convert.ToString(N_Mes)
        Dim Durée As String
        Dim Vélocité As String
        Dim a As String = " "
        Dim Poscol As String
        Dim NPiste As String = Canal

        ' Déterminer le n° du préset en fonction du N° de mesure
        Dim aa As String = Grid2.Cell(3, N_Mes).Text
        Dim NPres As Integer = Det_NumPréset(aa)

        ' Parcourir le préset dans la liste
        If NPres <> -1 Then ' -1 indique que la batterie ne joue sur l mesure concernée
            If Préset.Item(NPres).LNotes.Count <> 0 Then
                Me.PPresNotes = True
                For Each b As CNotes In Préset.Item(NPres).LNotes
                    'If Trim(Grid1.Cell(b.lig, 3).Text) <> "" And IsNumeric(Trim(Grid1.Cell(b.lig, b.col).Text)) Then
                    'If Val(Grid1.Cell(b.lig, b.col).Text) >= 0 And Val(Grid1.Cell(b.lig, b.col).Text) <= 127 Then
                    ValeurNote = Convert.ToString(ValNoteCubase.IndexOf(Grid1.Cell(b.lig, 3).Text))
                            '
                            Durée = "1" 'Det_Durée()
                            Vélocité = b.dyn
                            ' 
                            ' calcul de la position de la note
                            Pos = Convert.ToString(N_Mes) + "." + b.Pos
                            Poscol = Calc_PosNote(Pos, nbColonnesMes, Form1Début, nbRépet, LongueurPart)
                            '                    N° Piste         N° Canal        N° Note          Début          Durée         Vélocité
                            a = a + "Note" + " " + NPiste + " " + Canal + " " + ValeurNote + " " + Poscol + " " + Durée + " " + Vélocité + "-"
                    'End If
                    'End If
                Next
            End If
        End If
        Return Trim(a)
    End Function
    Function Det_Durée() As String
        Dim a As String = "1"
        Select Case ListTypNote.Text
            Case "RN*2", "WN*2"
                a = "32"
            Case "RN", "WN"
                a = "16"
            Case "BL", "HN"
                a = "8"
            Case "NR", "QN"
                a = "4"
            Case "CR", "EN"
                a = "2"
            Case "DC", "SN"
                a = "1"
        End Select
        Return Trim(a)
    End Function
    ''' <summary>
    ''' Calcul de la position d'une note 
    ''' </summary>
    ''' <param name="Posit">Position d'une note sous la forme 1.1.1</param>
    ''' <param name="nbColonnesMes">Nb colonne par mesure</param>
    ''' <param name="Form1Début">Valeur du locator de Départ</param>
    ''' <param name="nbRépet">nombre de répétitions</param>
    ''' <param name="LongueurPart">longueur de la partie à jouer (déterminée par les locateurs)</param>
    Function Calc_PosNote(Posit As String, nbColonnesMes As Integer, Form1Début As Integer, nbRépet As Integer, LongueurPart As Integer) As String ' on reçoit ici une position sous la forme
        Dim tbl() As String = Posit.Split(".")
        Dim i As Integer = Convert.ToInt16(tbl(0)) ' N° mes
        Dim j As Integer = Convert.ToInt16(tbl(1))
        Dim k As Integer = Convert.ToInt16(tbl(2)) ' N° crtTemps
        '
        ' Calcul de la position sans tenir compte du locateur de Départ
        Dim pos As Integer = ((i - 1) * (nbColonnesMes) + (j - 1) * 4) + (k - 1)
        ' Calcul de la position en tenant compte du locateur de Départ 
        pos = pos - ((Form1Début - 1) * nbColonnesMes)
        ' Calcul de la position en  tenant compte de la répétition
        pos = pos + (nbRépet * LongueurPart)
        Dim b As String = Convert.ToString(pos)
        Return b
    End Function

    Sub Init_ListNotesNom()

        Dim tbl() As String
        '
        With ListNotesNoms
            If LangueIHM = "fr" Then
                ListNotesNoms.Add("Effacer")
            Else
                ListNotesNoms.Add("Clear")
            End If
            .Add("C-2/-----")
            .Add("C#-2/-----")
            .Add("D-2/-----")
            .Add("D#-2/-----")
            .Add("E-2/-----")
            .Add("F-2/-----")
            .Add("F#-2/-----")
            .Add("G-2/-----")
            .Add("G#-2/-----")
            .Add("A-2/-----")
            .Add("A#-2/-----")
            .Add("B-2/-----")
            .Add("C-1/-----")
            .Add("C#-1/-----")
            .Add("D-1/-----")
            .Add("D#-1/-----")
            .Add("E-1/-----")
            .Add("F-1/-----")
            .Add("F#-1/-----")
            .Add("G-1/-----")
            .Add("G#-1/-----")
            .Add("A-1/-----")
            .Add("A#-1/-----")
            .Add("B-1/-----")
            .Add("C0/-----")
            .Add("C#0/-----")
            .Add("D0/-----")
            .Add("D#0/-----")
            .Add("E0/-----")
            .Add("F0/-----")
            .Add("F#0/-----")
            .Add("G0/-----")
            .Add("G#0/-----")
            .Add("A0/-----")
            .Add("A#0/-----")
            .Add("B0/Acoustic Bass Drum")
            .Add("C1/Bass Drum 1")
            .Add("C#1/Side Stick")
            .Add("D1/Acoustic Snare")
            .Add("D#1/Hand Clap")
            .Add("E1/Electric Snare")
            .Add("F1/Low Floor Tom")
            .Add("F#1/Closed Hi Hat")
            .Add("G1/High Floor Tom")
            .Add("G#1/Pedal Hi Hat")
            .Add("A1/Low Tom")
            .Add("A#1/Open Hi Hat")
            .Add("B1/Low Mid Tom")
            .Add("C2/Hi Mid Tom")
            .Add("C#2/Crash Cymbal 1")
            .Add("D2/High Tom")
            .Add("D#2/Ride Cymbal 1")
            .Add("E2/Chinese Cymbal")
            .Add("F2/Ride Bell")
            .Add("F#2/Tambourine")
            .Add("G2/Splash Cymbal")
            .Add("G#2/Cowbell")
            .Add("A2/Crash Cymbal 2")
            .Add("A#2/Vibraslap")
            .Add("B2/Ride Cymbal 2")
            .Add("C3/Hi Bongo")
            .Add("C#3/Low Bongo")
            .Add("D3/Mute Hi Conga")
            .Add("D#3/Open Hi Conga")
            .Add("E3/Low Conga")
            .Add("F3/High Timbale")
            .Add("F#3/Low Timbale")
            .Add("G3/High Agogo")
            .Add("G#3/Low Agogo")
            .Add("A3/Cabasa")
            .Add("A#3/Maracas")
            .Add("B3/Short Whistle")
            .Add("C4/Long Whistle")
            .Add("C#4/Short Guiro")
            .Add("D4/Long Guiro")
            .Add("D#4/Claves")
            .Add("E4/Hi Wood Block")
            .Add("F4/Low Wood Block")
            .Add("F#4/Mute Cuica")
            .Add("G4/Open Cuica")
            .Add("G#4/Mute Triangle")
            .Add("A4/Open Triangle")
            .Add("A#4/Autre-Other")
            .Add("B4/Autre-Other")
            .Add("C5/Autre-Other")
            .Add("C#5/Autre-Other")
            .Add("D5/Autre-Other")
            .Add("D#5/Autre-Other")
            .Add("E5/Autre-Other")
            .Add("F5/Autre-Other")
            .Add("F#5/Autre-Other")
            .Add("G5/Autre-Other")
            .Add("G#5/Autre-Other")
            .Add("A5/Autre-Other")
            .Add("A#5/Autre-Other")
            .Add("B5/Autre-Other")
        End With
        '
        For Each a As String In ListNotesNoms
            If a <> "Effacer" And a <> "Clear" Then
                tbl = a.Split("/")
                ListNotesGS.Items.Add(tbl(0))
                If Trim(tbl(1) <> "Autre-Other") Then
                    ListNomGS.Items.Add(tbl(1))
                End If
            Else
                ListNomGS.Items.Add(a)
            End If
        Next
        '
        ListNomGS.SelectedIndex = 1
        ListNotesGS.SelectedIndex = 0
        '
    End Sub
    Sub Grid1_MouseDown(sender As Object, e As EventArgs)
        Me.GridOrigine = "Grid1" ' correspond à la propriété PGridOrigine appelé dans Form1/Edition
        ListNomGS.Visible = False
        ListNotesGS.Visible = False
        '
        EntrerNomPerso.Visible = False ' textbox permettant d'entrer un nom d'instrument personnalisé
        '
        ' Jouer la note sur simple clic
        Dim ii As Integer = Grid1.MouseRow
        Dim jj As Integer = Grid1.MouseCol
        '
        If ii > 0 And jj > 3 And (NoteClic.Checked) Then ' cas où MouseRow ou MouseCol n'arrivent pas  à temps, ils valent -1 ce qui n'est pas utilisable
            If Not EnRecalcul Then
                Dim ValeurNote = ValNoteCubase.IndexOf(Grid1.Cell(ii, 3).Text)
                If IsNumeric(Grid1.Cell(ii, jj).Text) Then
                    JouerNote(ValeurNote, Convert.ToByte(Grid1.Cell(ii, jj).Text))
                Else
                    JouerNote(ValeurNote, Convert.ToByte(ListDynF2.Text))
                End If
            End If
            '
        End If
    End Sub
    Sub Grid1_MouseUp(sender As Object, e As EventArgs)
        Dim i As Integer
        Dim LigneCours As Integer = Grid1.MouseRow 'Grid1.ActiveCell.Row
        Dim ColCours As Integer = Grid1.MouseCol 'Grid1.ActiveCell.Col
        Dim l, c As Integer
        Dim NblignesVisibles As Integer = Grid1.ActiveCell.Row - Grid1.TopRow

        Grid1.AutoRedraw = False
        Grid1.BackColorFixedSel = Det_BackColorDrum(Det_NumPréset(Grid1.Cell(0, 0).Text))

        ' TRAITEMENT DES COLONNES 2 et 3 : Nom et Notes des instruments
        ' *************************************************************
        ' Combolist des Noms d'Intruments GS
        ' **********************************
        If My.Computer.Keyboard.CtrlKeyDown And ColCours = 2 Then ' colonnes des Nom GS d'instruments
            Dim s1 As New Size With {
            .Width = LargeColNom, ' en fait c'est plutôt fnt2 qui détermine la hauteur du combolist (voir plus loin)
            .Height = HautLignesGrid1 'ListNomGS.Size.Height
            }
            ListNomGS.Size = s1
            '
            i = Grid1.Row(1).Height
            Dim p1 As New Point With {
                .X = Grid1.Column(0).Width + Grid1.Column(1).Width + 1,
                .Y = Grid1.Row(0).Height + ((NblignesVisibles) * Grid1.Row(1).Height) + 2'+ 6
            }
            ListNomGS.Location = p1
            '
            ListNomGS.Font = fnt4 ' ' détermine la hauteur du combolist --> taille police
            ListNomGS.Visible = True
            ListNotesGS.Visible = False
            '
        End If
        ' combolist des notes
        ' *******************
        If My.Computer.Keyboard.CtrlKeyDown And ColCours = 3 Then ' colonne des notes
            Dim s2 As New Size With {
            .Width = LargeColNotes, ' en fait c'est plutôt fnt2 qui détermine la hauteur du combolist (voir plus loin)
            .Height = HautLignesGrid1'ListNotesGS.Size.Height
            }
            ListNotesGS.Size = s2

            Dim p2 As New Point With {
                .X = Grid1.Column(0).Width + Grid1.Column(1).Width + Grid1.Column(2).Width + 2,
                .Y = Grid1.Row(0).Height + ((NblignesVisibles) * Grid1.Row(1).Height) + 2'+ 6
            }
            ListNotesGS.Location = p2

            '
            ListNotesGS.Font = fnt4 ' détermine la hauteur du combolist --> taille police
            ListNomGS.Visible = False
            ListNotesGS.Visible = True
        End If
        '
        ' Boite texte de modification des Noms GS
        ' ***************************************
        With My.Computer.Keyboard
            If .AltKeyDown And ColCours = 2 Then ' colonnes des Nom GS d'instruments
                ListNomGS.Visible = False
                Dim s1 As New Size With {
            .Width = LargeColNom,
            .Height = HautLignesGrid1 'ListNomGS.Size.Height
            }
                EntrerNomPerso.Size = s1
                '
                Dim p1 As New Point With {
                .X = Grid1.Column(0).Width + Grid1.Column(1).Width + 1,
                .Y = Grid1.Row(0).Height + ((NblignesVisibles) * Grid1.Row(1).Height) '+ 5
            }
                EntrerNomPerso.Location = p1
                '
                EntrerNomPerso.Font = fnt4
                EntrerNomPerso.Visible = True
                EcrNomPerso_Visible = True
                l = Grid1.ActiveCell.Row
                c = Grid1.ActiveCell.Col
                'DéVérouillerGrid1()

                EntrerNomPerso.Focus()
                EntrerNomPerso.Text = Trim(Grid1.Cell(l, c).Text)
                EntrerNomPerso.SelectionStart = Len(EntrerNomPerso.Text)
                'VérouillerGrid1()
                '
                ' 
            End If
        End With
        '
        '  écrire/effacer des notes : TRAITEMENT DES COLONNES 4 et 19 :
        ' *************************************************************
        If My.Computer.Keyboard.CtrlKeyDown And ColCours >= 4 Then
            Sauv_ClipAnnuler()
            l = Grid1.ActiveCell.Row
            c = Grid1.ActiveCell.Col
            ' 
            Grid1.Cell(l, c).Locked = False
            If IsNumeric(Grid1.Cell(l, c).Text) Then
                Grid1.Cell(l, c).Text = "" ' effacer note
            Else
                Grid1.Cell(l, c).Text = Trim(ListDynF2.Text) ' écrire note (par écriture de la vélocité de la note)
            End If
            Grid1.Cell(l, c).Locked = True
            ' 
            i = Det_NumPréset(Trim(Grid1.Cell(0, 0).Text))
            Maj_Préset(i)
            'Refresh_PianoRoll() ' refresh si check Model Drum est coché dans le piano (model de la position des notes de batterie désignée le carré bleu dans la colonne1)
            'Form1.ReCalcul()
        End If
        '
        ' Gestion check box init calque rythmique
        ' ***************************************
        If ColCours = 1 Then
            'For i = 1 To Grid1.Rows - 1
            'Grid1.Cell(i, 1).BackColor = Color.White
            'Next
            'Grid1.Cell(LigneCours, ColCours).BackColor = Module1.Couleur_ModelDrum
            'Refresh_PianoRoll()
        End If
        '
        Grid1.AutoRedraw = True
        Grid1.Refresh()
        ' Coupure de note jouée avec la souris
        ' ************************************
        If NoteAEtéJouée = True Then
            StoperNote(NoteCourante, 0)
            NoteAEtéJouée = False
        End If
    End Sub
    Sub Grid1_Scroll(sender As Object, e As EventArgs)
        EntrerNomPerso.Visible = False
        ListNotesGS.Visible = False
        ListNomGS.Visible = False
    End Sub

    Public Sub Sauv_ClipAnnuler()
        Dim i, j As Integer
        Dim Fr As Integer = Grid1.Selection.FirstRow
        Dim Fc As Integer = Grid1.Selection.FirstCol
        Dim Lr As Integer = Grid1.Selection.LastRow
        Dim Lc As Integer = Grid1.Selection.LastCol
        Dim a As String
        ' RAZ 
        Dim k = Det_NumPréset2()

        ListClipAnnuler(k).lstClip.Clear()

        For i = Fr To Lr
            For j = Fc To Lc
                a = i.ToString + " " + j.ToString + " " + Trim(Grid1.Cell(i, j).Text)
                ListClipAnnuler(k).lstClip.Add(a)
            Next j
        Next i

    End Sub
    Public Sub Restit_ClipAnnuler()
        Dim i, ii, j As Integer
        Dim tbl() As String
        Dim k = Det_NumPréset(Trim(Grid1.Cell(0, 0).Text))
        If ListClipAnnuler(k).lstClip.Count > 0 Then
            For Each a As String In ListClipAnnuler(k).lstClip
                tbl = a.Split()
                i = Convert.ToInt16(tbl(0))
                j = Convert.ToInt16(tbl(1))
                Grid1.Cell(i, j).Text = Trim(tbl(2))
            Next
        End If
        Grid1.Refresh()
        ii = Det_NumPréset2()
        Maj_Préset(ii)
    End Sub

    Sub Refresh_PianoRoll()
        ' rafraichir les PianoRoll si Check Modele Drum est coché
        For i = 0 To Module1.nb_PianoRoll - 1
            If Form1.PIANOROLLChargé(i) = True Then
                Form1.listPIANOROLL(i).F1_Refresh2()
            End If
        Next
    End Sub

    Private Sub Grid1_Keydown(Sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Dim i As Integer = Grid1.ActiveCell.Row
        Dim j As Integer = Grid1.ActiveCell.Col
        Dim ii As Integer
        Dim k As Integer
        Dim a As String ' = Trim(Grid1.Cell(i, j).Text)
        Dim b As String
        Dim tbl1() As String
        Dim tbl2() As String
        Dim sortir As Boolean = False
        Dim NPres As Integer = Det_NumPréset(Trim(Grid1.Cell(0, 0).Text))

        Grid1.AutoRedraw = False

        'VérouillerGrid1() ' les cellules seront dévverrouilées 

        'Grid1.Range(1, 4, NbLignesGrid1 - 1, Grid1.Cols - 1).Locked = True

        ' Incrémenter/décrémenter les valeurs de dynamique

        If e.KeyCode = Keys.Add Or e.KeyCode = Keys.Subtract Then
            ' arrowdirection
            b = CellDyn()
            ' traitement des dynamiques avec les touches + et -
            If Trim(b) <> "" Then
                tbl1 = Split(b)
                For ii = 0 To UBound(tbl1)
                    tbl2 = Split(tbl1(ii), ",")
                    i = Convert.ToInt16(tbl2(0)) ' ligne - - - 
                    j = Convert.ToInt16(tbl2(1)) ' col
                    k = Convert.ToInt16(Grid1.Cell(i, j).Text)
                    ' détermination de la nécessité d'incrémenter
                    If k >= 127 And e.KeyCode = Keys.Add Then
                        sortir = True ' si l'une des valeurs =127 alors on n'augmente plus aucune valeur de la sélection --> sortir = true
                        Exit For
                    End If
                    ' détermination de la nécessité de décrémenter
                    If k <= 0 And e.KeyCode = Keys.Subtract Then
                        sortir = True ' si l'une des valeurs =127 alors on n'augmente plus aucune valeur de la sélection  --> sortir = true
                        Exit For
                    End If
                    '
                Next
                ' incrémentation/décrémentation (si sortir = false)
                If sortir = False Then
                    For ii = 0 To UBound(tbl1)
                        tbl2 = Split(tbl1(ii), ",")
                        i = Convert.ToInt16(tbl2(0))
                        j = Convert.ToInt16(tbl2(1))
                        a = Trim(Grid1.Cell(i, j).Text)
                        'If IsNumeric(a) Then
                        If e.KeyCode = Keys.Add Then
                            k = Convert.ToInt16(a) + 1
                            If k <= 127 Then
                                'DéVérouillerGrid1()
                                Grid1.Cell(i, j).Text = Convert.ToString(k)
                                'VérouillerGrid1()
                            End If
                        ElseIf e.KeyCode = Keys.Subtract Then
                            k = Convert.ToInt16(a) - 1
                            If k >= 0 Then
                                'DéVérouillerGrid1()
                                Grid1.Cell(i, j).Text = Convert.ToString(k)
                                'VérouillerGrid1()
                            End If
                        End If
                        'End If
                        Maj_Préset(NPres)
                    Next

                End If
            End If
        End If

        Grid1.AutoRedraw = True
        Grid1.Refresh()
    End Sub
    Sub NomPréset_KeyUp(Sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        LNomPréset.Item(Det_NumPréset(Trim(Grid1.Cell(0, 0).Text))) = NomPréset.Text
    End Sub

    Private Sub Grid1_KeyUp(Sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Dim i, ii As Integer
        ' Effacer avec la touche  suppr
        If e.KeyCode = Keys.Delete Then

            Dim Fr As Integer = Grid1.Selection.FirstRow
            Dim Fc As Integer = Grid1.Selection.FirstCol
            Dim Lr As Integer = Grid1.Selection.LastRow
            Dim Lc As Integer = Grid1.Selection.LastCol
            If Fc >= 4 Then
                '
                'DéVérouilerCell()
                For i = Fr To Lr
                    For j = Fc To Lc
                        Grid1.Cell(i, j).Text = ""
                    Next
                Next
                'VérouillerGrid1()
                '
                ii = Det_NumPréset2()
                Maj_Préset(ii)
            End If
        End If
        '

        '
        'DéVérouillerGrid1() ' les cellules ont été vérouillées dans Grid1_KeyDown
    End Sub
    Sub ListPrésets_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim a As String = ListDesPrésets.Item(ListPrésets.SelectedIndex) ' lecture de la chaine du Préset
        ChargerPréset(a)
    End Sub

    Sub ListNomGS_SelectedIndexChanged(sender As Object, e As EventArgs)

        If Not Enchargement Then '

            Dim l As Integer = Grid1.ActiveCell.Row
            Dim c As Integer = Grid1.ActiveCell.Col


            If ListNomGS.Text <> "Effacer" And ListNomGS.Text <> "Clear" Then
                Grid1.Cell(l, c).Text = ListNomGS.Text
                ListNomGS.Visible = False
                '
                Dim a As String = ListNotesNoms.Item(ListNomGS.SelectedIndex)
                Dim tbl() As String = a.Split("/")
                Grid1.Cell(l, c + 1).Text = tbl(0)
                Grid1.Refresh()
                ListNomGS.Visible = False
                '
            Else
                Grid1.Cell(l, c).Text = ""
                Grid1.Cell(l, c + 1).Text = ""
                ListNomGS.Visible = False
            End If
        End If
    End Sub
    Sub ListNotesGS_SelectedIndexChanged(sender As Object, e As EventArgs)
        If Not Enchargement Then
            If Not Enchargement Then '

                Dim l As Integer = Grid1.ActiveCell.Row
                Dim c As Integer = 3 ' ne pas utilser  Grid1.ActiveCell.col
                '
                DéVérouillerGrid1()
                Grid1.Cell(l, c).Text = ListNotesGS.Text
                VérouillerGrid1()
                ListNotesGS.Visible = False
                '
                '
            End If
        End If
    End Sub
    Sub EntrerNomPerso_KeyPress(sender As Object, e As KeyPressEventArgs)
        Dim a As Char
        a = e.KeyChar
        If a = vbCr Then
            DéVérouillerGrid1()
            Dim l As Integer = Grid1.ActiveCell.Row
            Dim c As Integer = 2 '  on n'écrit toujours que dans la ligne 2 avec le textbox EntrerNomPerso -    Grid1.ActiveCell.Col
            Grid1.Cell(l, c).Text = Trim(EntrerNomPerso.Text)
            EcrNomPerso_Visible = False
            EntrerNomPerso.Visible = False
            'VérouillerGrid1()
        End If
    End Sub

    ''' <summary>
    ''' Construction de la time line : répartition des patternes A, B, ... dans Grid2
    ''' </summary>
    Private Sub Construction_Grid2()
        Dim i, j As Integer

        Grid2.AutoRedraw = False
        AddHandler Grid2.KeyDown, AddressOf Grid2_KeyDown
        AddHandler Grid2.MouseDown, AddressOf Grid2_MouseDown

        '
        Grid2.AllowDrop = False
        Grid2.AllowUserPaste = ClipboardDataEnum.TextAndFlexCellFormat
        Grid2.AllowUserReorderColumn = False
        Grid2.AllowUserResizing = False
        Grid2.AllowUserSort = False
        ' dimensionnement
        Grid2.Cols = nbMesures + 1 ' Me.nbRépétitionMax
        Grid2.FixedCols = 1 ' le minimum est 1
        Grid2.FixedRows = 3
        Grid2.Rows = 4
        ' 
        For i = 0 To Grid2.Rows - 1
            Grid2.Row(i).Height = 20
        Next
        '
        Grid2.Column(0).Width = 55
        For j = 1 To Grid2.Cols - 1
            Grid2.Column(j).Width = 70
        Next j
        '

        'Grid2.SelectionBorderColor = Color.Transparent

        Grid2.BackColorFixed = DrmC_GrisBleu 'Color.FromArgb(209, 130, 11)
        Grid2.Cell(3, 0).ForeColor = Color.Black
        For i = 0 To Grid2.Rows - 2
            For j = 0 To Grid2.Cols - 1
                Grid2.Cell(i, j).ForeColor = Color.Black
            Next
        Next
        '
        Grid2.ScrollBars = ScrollBarsEnum.Horizontal
        '
        ' Styles, Titres et Numérotations
        ' *******************************
        Dim f As New System.Drawing.Font("Arial", 12, FontStyle.Bold)
        Grid2.DefaultFont = f
        '
        Dim f1 As New System.Drawing.Font("Tahoma", 7, FontStyle.Regular)
        For i = 1 To 2
            For j = 0 To Grid2.Cols - 1
                Grid2.Cell(i, j).Font = f1
            Next
        Next

        Dim f2 As New System.Drawing.Font("Tahoma", 10, FontStyle.Regular)

        For j = 0 To Grid2.Cols - 1
            Grid2.Cell(0, j).Font = f2
        Next
        Grid2.Cell(0, 0).Font = f1
        '
        If Me.Langue = "fr" Then
            Grid2.Cell(0, 0).Text = "Mesures"
            Grid2.Cell(1, 0).Text = "Marqueurs"
            Grid2.Cell(2, 0).Text = "Accords"
            Grid2.Cell(3, 0).Text = ""
        Else
            Grid2.Cell(0, 0).Text = "Measures"
            Grid2.Cell(1, 0).Text = "Markers"
            Grid2.Cell(2, 0).Text = "Chords"
            Grid2.Cell(3, 0).Text = ""
        End If
        '
        Grid2.Column(0).Alignment = AlignmentEnum.LeftCenter
        For j = 1 To Grid2.Cols - 1
            Grid2.Cell(0, j).Text = Convert.ToString(j)
            Grid2.Column(j).Alignment = AlignmentEnum.CenterCenter
        Next
        ' écriture des variation de base
        For j = 1 To Grid2.Cols - 1
            Grid2.Cell(3, j).Text = "A"
            Grid2.Cell(3, j).BackColor = Det_BackColorDrum(0)
            Grid2.Cell(3, j).ForeColor = Color.White
        Next
        'Grid1.Range(5, 0, Grid1.Rows - 1, Grid1.Cols - 1).Locked = True

        ' mise à jour des accords et gammes

        Grid2.Anchor = 5
        Grid2.Dock = DockStyle.Fill
        '
        ' Création de comboboVar
        Panneau2.Panel1.Controls.Add(ComBoVar)
        '

        ComBoVar.Items.Add("A")
        ComBoVar.Items.Add("B")
        ComBoVar.Items.Add("C")
        ComBoVar.Items.Add("D")
        ComBoVar.Items.Add("E")
        ComBoVar.Items.Add("F")
        ComBoVar.Items.Add("G")
        ComBoVar.Items.Add("H")
        ComBoVar.Items.Add("--")
        '
        ComBoVar.SelectedIndex = 0

        Dim s As New Size With {
            .Width = Grid2.Column(0).Width,
            .Height = Grid2.Row(0).Height}

        Dim y As Integer = Grid2.Row(0).Height + Grid2.Row(1).Height + Grid2.Row(2).Height 'Form1.TabControl4.Location.Y +
        Dim x As Integer = 1
        Dim p As New Point With {
            .X = x,
            .Y = y}

        ComBoVar.Size = s
        ComBoVar.Location = p
        '
        ComBoVar.BringToFront()
        ComBoVar.Visible = True

        Grid2.ScrollBars = FlexCell.ScrollBarsEnum.Horizontal
        Grid2.BoldFixedCell = False

        ' évènements
        AddHandler ComBoVar.SelectedIndexChanged, AddressOf Combovar_SelectedIndexChanged
        'AddHandler Grid2.MouseUp, AddressOf Grid2_MouseUp

        Grid2.SelectionMode = FlexCell.SelectionModeEnum.ByColumn
        Grid2.DisplayFocusRect = True

        Grid2.BackColorActiveCellSel = Color.Transparent

        Grid2.BackColorFixedSel = Color.Beige
        Grid2.BackColorBkg = Color.Beige 'Color.White
        Grid2.BackColor1 = Color.Beige 'Color.Lavender 'Color.White
        Grid2.BackColorSel = Color.Beige
        Grid2.BackColorActiveCellSel = Color.Beige
        '
        ' Fermeture des cellules (locked= true)
        Grid2.Range(0, 0, Grid2.Rows - 1, Grid2.Cols - 1).Locked = True

        Grid2.AutoRedraw = True
        Grid2.Refresh()
        '
    End Sub
    Sub Grid2_MouseDown(sender As Object, e As MouseEventArgs)
        Me.GridOrigine = "Grid2"
        '
        If e.Button = MouseButtons.Right Then
            Me.GridOrigine = "Grid2"
            Dim i As Integer = Grid2.MouseRow
            Dim j As Integer = Grid2.MouseCol
            Dim p As New Point
            p.X = Cursor.Position.X
            p.Y = Cursor.Position.Y
            MenuContext1.Show(p)
        End If

    End Sub
    Private Sub Lprests_Click(sender As Object, e As EventArgs) ' menus flottant contextuel
        Dim coul As Color
        Dim j As Integer
        Dim com As ToolStripMenuItem = sender
        Dim aa As String
        aa = com.Tag.ToString

        Dim Fcol As Integer = Grid2.Selection.FirstCol
        Dim Lcol As Integer = Grid2.Selection.LastCol
         
        Grid2.Range(0, 0, Grid2.Rows - 1, Grid2.Cols - 1).Locked = False
        For j = Fcol To Lcol
            coul = Det_BackColorDrumLettre(Trim(aa))
            Grid2.Cell(3, j).Text = Trim(aa)
            Grid2.Cell(3, j).BackColor = coul
            Grid2.Cell(3, j).ForeColor = Color.White
        Next
        Grid2.Range(0, 0, Grid2.Rows - 1, Grid2.Cols - 1).Locked = True
    End Sub

    Private Sub itemCouper_Click(sender As Object, e As EventArgs) ' menus flottant contextuel
        Dim j As Integer
        Dim Fcol As Integer = Grid2.Selection.FirstCol
        Dim Lcol As Integer = Grid2.Selection.LastCol

        coul = Det_BackColorDrumLettre(Trim("Vide"))


        Grid2.Range(0, 0, Grid2.Rows - 1, Grid2.Cols - 1).Locked = False
        Grid2.Selection.CutData()
        Grid2.Range(0, 0, Grid2.Rows - 1, Grid2.Cols - 1).Locked = True
        For j = Fcol To Lcol
            Grid2.Cell(3, j).Text = "--"
            Grid2.Cell(3, j).BackColor = coul
            Grid2.Cell(3, j).ForeColor = Color.White
        Next
    End Sub
    Private Sub itemCopier_Click(sender As Object, e As EventArgs) ' menus flottant contextuel
        Grid2.Selection.CopyData()
    End Sub

    Private Sub itemColler_Click(sender As Object, e As EventArgs) ' menus flottant contextuel
        Grid2.Range(0, 0, Grid2.Rows - 1, Grid2.Cols - 1).Locked = False
        Grid2.Selection.PasteData()
        Grid2.Range(0, 0, Grid2.Rows - 1, Grid2.Cols - 1).Locked = True
    End Sub
    ''' <summary>
    ''' Cet évènement permet de courtcircuiter le CTRL + X sur Grid2 du DrumEdit
    ''' </summary>
    ''' <param name="sender">Objet qui génère l'évènement</param>
    ''' <param name="e">Evènement concernant les touches du clavier</param>
    Sub Grid2_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs)
        If e.Control And e.KeyCode = Keys.X Then
            e.Handled = True
        End If
    End Sub
    Sub Combovar_SelectedIndexChanged(sender As Object, e As EventArgs)
        If Not Enchargement Then

            Dim j As Integer
            Dim coul As New Color
            Grid2.Range(0, 0, Grid2.Rows - 1, Grid2.Cols - 1).Locked = False
            For j = Grid2.Selection.FirstCol To Grid2.Selection.LastCol
                coul = Det_BackColorDrum(ComBoVar.SelectedIndex)
                Grid2.Cell(3, j).BackColor = coul
                Grid2.Cell(3, j).ForeColor = Color.White
                Grid2.Cell(3, j).Text = Trim(ComBoVar.SelectedItem)
            Next
            Grid2.Range(0, 0, Grid2.Rows - 1, Grid2.Cols - 1).Locked = True
        End If

    End Sub
    ''' <summary>
    ''' Construction_Menu : menu Fichier/Attacher
    ''' </summary>
    Public Sub Construction_Menu()
        '
        Fichier.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Quitter})
        Fichier.Size = New System.Drawing.Size(87, 20)
        Fichier.Text = "Fichier"
        Fichier.Visible = True

        ' Edition
        Edition.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Couper, Copier, Coller, Séparateur1, Annuler, Séparateur2, Supprimer})
        Edition.Size = New System.Drawing.Size(56, 20)
        Edition.Text = "Edition"
        Edition.Visible = True
        ' 
        ' Quitter (de Fichier)
        '
        If Me.Langue = "fr" Then
            Quitter.Text = "Attacher"
        Else
            Quitter.Text = "Dock"
        End If
        Quitter.ShortcutKeys = Shortcut.CtrlD
        Quitter.Size = New System.Drawing.Size(180, 22)

        ' Couper
        If Me.Langue = "fr" Then
            Couper.Text = "Couper"
        Else
            Couper.Text = "Cut"
        End If
        Couper.ShortcutKeys = Shortcut.CtrlX
        Couper.Size = New System.Drawing.Size(180, 22)
        '
        ' Copier
        If Me.Langue = "fr" Then
            Copier.Text = "Copier"
        Else
            Copier.Text = "Copy"
        End If
        Copier.ShortcutKeys = Shortcut.CtrlC
        Copier.Size = New System.Drawing.Size(180, 22)
        '
        ' Coller
        If Me.Langue = "fr" Then
            Coller.Text = "Coller"
        Else
            Coller.Text = "Paste"
        End If
        Coller.ShortcutKeys = Shortcut.CtrlV
        Coller.Size = New System.Drawing.Size(180, 22)
        '
        ' Séparateur
        Séparateur1.Size = New System.Drawing.Size(177, 6)

        ' Annuler
        If Me.Langue = "fr" Then
            Annuler.Text = "Annuler"
        Else
            Annuler.Text = "Cancel"
        End If
        Annuler.ShortcutKeys = Shortcut.CtrlZ
        Annuler.Size = New System.Drawing.Size(180, 22)
        Annuler.Enabled = True

        ' Séparateur
        Séparateur1.Size = New System.Drawing.Size(177, 6)
        '
        If Me.Langue = "fr" Then
            Supprimer.Text = "Supprimer"
        Else
            Supprimer.Text = "Delete"
        End If
        Supprimer.ShortcutKeys = Shortcut.Del
        Supprimer.Size = New System.Drawing.Size(180, 22)
        '
        ' Menu

        Menu1.Text = "Menu"
        Me.Menu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Fichier})
        Me.Menu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Edition})
        Menu1.Location = New System.Drawing.Point(0, 0)
        F2.Controls.Add(Menu1)
        F2.MainMenuStrip = Menu1
        F2.MainMenuStrip.Visible = False
        F2.MainMenuStrip.Enabled = False
        Menu1.Visible = False
        Menu1.Enabled = False

        AddHandler Couper.Click, AddressOf Couper_Click
        AddHandler Copier.Click, AddressOf Copier_Click
        AddHandler Coller.Click, AddressOf Coller_Click
        AddHandler Annuler.Click, AddressOf Annuler_Click
        AddHandler Supprimer.Click, AddressOf Supprimer_Click
        AddHandler Quitter.Click, AddressOf Quitter_Click
    End Sub
    Private Sub Couper_Click(sender As Object, e As EventArgs)
        If Grid1.Selection.FirstCol > 3 And Grid1.Selection.FirstRow > 0 Then
            DéVérouillerGrid1()   ' les cellules ont été vérrouillées dans drums.grid1_Keydown 
            Grid1.AutoRedraw = False
            Sauv_ClipAnnuler()
            Grid1.Selection.CutData()
            DessinDiv()
            ii = Det_NumPréset2()
            Maj_Préset(ii)
            Grid1.AutoRedraw = True
            Grid1.Refresh()
            VérouillerGrid1()     ' les cellules seront déverrouillées à la fin de Drums.Grid1_KeyUp
        End If
    End Sub
    Private Sub Copier_Click(sender As Object, e As EventArgs)
        If Grid1.Selection.FirstCol > 3 And Grid1.Selection.FirstRow > 0 Then
            If PGridOrigine = "Grid1" Then
                Grid1.Selection.CopyData()
            End If
        End If
        If PGridOrigine = "Grid2" Then
            Grid2.Selection.CopyData()
        End If
    End Sub
    Private Sub Coller_Click(sender As Object, e As EventArgs)
        If PGridOrigine = "Grid1" Then
            DéVérouillerGrid1()   ' les cellules ont été vérrouillées dans drums.grid1_Keydown
            Sauv_ClipAnnuler()
            Grid1.Selection.PasteData()
            ii = Det_NumPréset2()
            Maj_Préset(ii)
            DessinDiv()           ' cette méthode contient Autoredraw=false,true et Refresh
            VérouillerGrid1()     ' les cellules seront déverrouillées à la fin de Drums.Grid1_KeyUp
        End If
        '
        If PGridOrigine = "Grid2" Then
            Grid2.Range(0, 0, Grid2.Rows - 1, Grid2.Cols - 1).Locked = False ' dévéroulluer grid2
            Grid2.Selection.PasteData()
            Grid2.Range(0, 0, Grid2.Rows - 1, Grid2.Cols - 1).Locked = True ' vérouiller grid2
        End If
    End Sub
    Private Sub Annuler_Click(sender As Object, e As EventArgs)
        Restit_ClipAnnuler()
    End Sub
    Private Sub Supprimer_Click(sender As Object, e As EventArgs)
        ' la capture de la touche suppr est traitée directement dans KeyUp
    End Sub
    Sub Quitter_Click(Sender As Object, e As EventArgs)
        Dim NumOnglet As Integer = Convert.ToUInt16(F2.Tag)
        Attacher(NumOnglet)
    End Sub
    Sub Attacher(Numonglet)
        Me.F2.FormBorderStyle = FormBorderStyle.None
        Me.F2.TopMost = False   ' un seul des 2 suffit ?
        Me.F2.TopLevel = False
        F2.MainMenuStrip.Visible = False
        Form1.TabControl4.TabPages.Item(Numonglet).Controls.Add(Me.F2)
        Me.F2.Dock = DockStyle.Fill
        '
        If LangueIHM = "fr" Then
            DockButton.Text = "Détacher"
        Else
            DockButton.Text = "UnDock"
        End If
    End Sub
    ''' <summary>
    ''' Construction de la barre d'outils
    ''' </summary>
    Private Sub Construction_Barrout()
        Dim i, j As Integer
        Dim f1 As New System.Drawing.Font("Tahoma", 8, FontStyle.Regular)
        '
        ' Case à cocher de mute
        ' *********************
        Panneau1.Panel1.Controls.Add(CheckMute)
        CheckMute.TabStop = False
        '
        Dim p As New Point With {
                .X = 20,
                .Y = 13
            }
        Dim s As New Size With {
            .Width = 120,
            .Height = 20
        }
        CheckMute.Location = p
        CheckMute.Size = s
        CheckMute.Checked = True
        CheckMute.Visible = True
        CheckMute.Enabled = False
        '
        CheckMute.Font = Module1.fontMutePiste
        CheckMute.ForeColor = Color.DarkRed
        '
        If Me.Langue = "fr" Then
            CheckMute.Text = "PISTE 10"
        Else
            CheckMute.Text = "TRACK 10"
        End If
        '
        AddHandler CheckMute.CheckedChanged, AddressOf CheckMute_CheckedChanged

        Panneau1.Panel1.Controls.Add(SoloBoutDRM)
        SoloBoutDRM.TabStop = False
        '
        SoloBoutDRM.Location = New Point(18, 22)
        SoloBoutDRM.Font = New Font("Calibri", 8, FontStyle.Regular)
        SoloBoutDRM.Size = New Size(50, 20)
        SoloBoutDRM.BackColor = Color.Beige
        SoloBoutDRM.Text = "SOLO"
        SoloBoutDRM.Visible = False
        AddHandler SoloBoutDRM.MouseUp, AddressOf SoloBoutDRM_MouseUp
        '
        ' liste des dynamiques
        ' ********************
        p.Y = 43
        p.X = CheckMute.Location.X
        s.Width = 50
        ListTypNote.Size = s
        ListTypNote.Location = p
        '
        '
        Panneau1.Panel1.Controls.Add(ListDynF2)
        ListDynF2.TabStop = False
        Panneau1.Panel1.Controls.Add(LabelDyn)
        LabelDyn.TabStop = False
        ListDynF2.Size = s
        ListDynF2.Location = p
        ListDynF2.Enabled = True
        '
        If Langue = "fr" Then
            LabelDyn.Text = "Dynamique"
        Else
            LabelDyn.Text = "Velocity"
        End If

        LabelDyn.AutoSize = False
        LabelDyn.TextAlign = ContentAlignment.MiddleLeft
        p.Y = p.Y + 18
        p.X = CheckMute.Location.X - 3
        LabelDyn.Location = p
        '
        For i = 127 To 0 Step -1
            ListDynF2.Items.Add(Convert.ToString(i))
        Next
        '
        If Langue = "fr" Then
            ToolTip1.SetToolTip(ListDynF2, "Dynamique")
        Else
            ToolTip1.SetToolTip(ListDynF2, "Velocity")
        End If
        ListDynF2.Font = fnt9
        ListDynF2.SelectedIndex = 37

        ' Case à cocher de Note sur Clic
        ' *****************************
        Panneau1.Panel1.Controls.Add(NoteClic)
        NoteClic.TabStop = False
        '

        p.X = 20
        p.Y = p.Y + 40

        s.Width = 120
        s.Height = 20

        NoteClic.Location = p
        NoteClic.Size = s
        NoteClic.Checked = True
        NoteClic.Visible = True
        NoteClic.Enabled = True
        NoteClic.AutoSize = True
        '
        NoteClic.Font = fnt9
        NoteClic.ForeColor = Color.Black
        '
        If Me.Langue = "fr" Then
            NoteClic.Text = "Ecoute note sur Clic"
        Else
            NoteClic.Text = "Listen note on clic"
        End If

        '
        ' liste des longueurs de notes
        ' ****************************
        p.Y = 50
        p.X = CheckMute.Location.X + ListTypNote.Location.X + ListTypNote.Size.Width
        ListTypNote.Location = p
        '
        p.Y = 75
        LabelTypNote.Location = p
        '
        s.Width = 40
        '
        Panneau1.Panel1.Controls.Add(ListTypNote)
        ListTypNote.TabStop = False
        Panneau1.Panel1.Controls.Add(LabelTypNote)
        If Langue = "fr" Then
            ListTypNote.Items.Add("RN*2")
            ListTypNote.Items.Add("RN")
            ListTypNote.Items.Add("BL")
            ListTypNote.Items.Add("NR")
            ListTypNote.Items.Add("CR")
            ListTypNote.Items.Add("DC")
            LabelTypNote.Text = "Durée"
            ToolTip1.SetToolTip(ListTypNote, "RN:ronde BL:blanche NR:noire CR:croche DC:Double croche")
        Else
            ListTypNote.Items.Add("WN*2")
            ListTypNote.Items.Add("WN")
            ListTypNote.Items.Add("HN")
            ListTypNote.Items.Add("QN")
            ListTypNote.Items.Add("EN")
            ListTypNote.Items.Add("SN")
            LabelTypNote.Text = "Length"
            ToolTip1.SetToolTip(ListTypNote, "WH:whole HN:half QN:quarter EN:eighth  SN:sixteen")
        End If
        ListTypNote.Font = f1
        ListTypNote.SelectedIndex = 4
        '
        ListTypNote.Visible = False
        LabelTypNote.Visible = False
        '
        '
        ' Affichage/Ecriture nom des Présets
        ' **********************************
        p.Y = 14
        p.X = 200
        s.Width = 160
        s.Height = 30
        Panneau1.Panel1.Controls.Add(NomPréset)
        NomPréset.TabStop = False
        NomPréset.Location = p
        NomPréset.Size = s
        NomPréset.BorderStyle = BorderStyle.FixedSingle
        NomPréset.BackColor = Color.White
        NomPréset.Font = fnt6
        NomPréset.Visible = True

        ' Bouton des Présets : bouton de A à H
        ' ************************************
        p.X = 160
        s.Width = 40
        s.Height = 40
        j = 1
        For i = 0 To NbDrumPrésets - 1
            LVar.Add(New Button)
            LVar.Item(i).Text = LPréset.Item(i) '"D" + Convert.ToString(i + 1)
            LVar.Item(i).Font = f1
            LVar.Item(i).BackColor = Color.DarkGreen 'Color.FromArgb(58, 68, 88)
            LVar.Item(i).ForeColor = Color.Yellow
            p.X = 160 + s.Width * i
            If i <= 3 Then
                p.Y = 50 '+ j + (LVar.Item(i).Size.Height) * i
                p.X = 200 + s.Width * i
            Else
                p.Y = 50 + s.Height ' + 20 + j + (LVar.Item(i).Size.Height) * i
                p.X = 200 + s.Width * (i - 4)
            End If
            LVar.Item(i).Location = p
            LVar.Item(i).Size = s
            LVar.Item(i).Visible = True
            Panneau1.Panel1.Controls.Add(LVar.Item(i))
            LVar.Item(i).TabStop = False
            j = 15 * (i + 1)
            LVar.Item(i).ForeColor = Color.White
            LVar.Item(i).BackColor = Det_BackColorDrum(i)
            LVar.Item(i).Tag = i
            If Langue = "fr" Then
                ToolTip1.SetToolTip(LVar.Item(i), "Motif " + Det_LettrePreset(i))
            Else
                ToolTip1.SetToolTip(LVar.Item(i), "Pattern " + Det_LettrePreset(i))
            End If
            AddHandler LVar.Item(i).Click, AddressOf Lvar_Click
        Next

        ' Boutons
        ' *******
        ' bouton Init
        Panneau1.Panel1.Controls.Add(BoutInit)
        BoutInit.TabStop = False
        'p.X = 20
        p.X = CheckMute.Location.X
        p.Y = p.Y + 52
        s.Width = 100
        s.Height = 40
        '
        BoutInit.FlatStyle = FlatStyle.Flat
        BoutInit.FlatAppearance.BorderColor = Color.Gray
        BoutInit.FlatAppearance.BorderSize = 2
        BoutInit.FlatAppearance.MouseDownBackColor = Color.Gold
        BoutInit.FlatAppearance.MouseOverBackColor = Color.GreenYellow

        BoutInit.BackColor = Color.Khaki       'Color.FromArgb(228, 62, 52) 'orange
        BoutInit.ForeColor = Color.Black
        BoutInit.Font = f1
        BoutInit.Location = p
        BoutInit.Size = s
        BoutInit.Text = "Init mapping"
        '
        If Me.Langue = "fr" Then
            ToolTip1.SetToolTip(BoutInit, "Rétablir le mappage de base des instruments")
            BoutInit.Text = "Init mappage"
        Else
            ToolTip1.SetToolTip(BoutInit, "Restore the basic mapping of intruments")
            BoutInit.Text = "Init mapping"
        End If
        ' Bouton Clear
        Panneau1.Panel1.Controls.Add(BoutClear)
        BoutClear.TabStop = False
        p.X = 20
        p.Y = p.Y + 51
        s.Width = 100
        s.Height = 40
        '
        BoutClear.FlatStyle = FlatStyle.Flat
        BoutClear.FlatAppearance.BorderColor = Color.Gray
        BoutClear.FlatAppearance.BorderSize = 2
        BoutClear.FlatAppearance.MouseDownBackColor = Color.Gold
        BoutClear.FlatAppearance.MouseOverBackColor = Color.GreenYellow

        BoutClear.BackColor = Color.FromArgb(246, 209, 87) ' jaune
        BoutClear.Font = f1
        BoutClear.Location = p
        BoutClear.Size = s
        If Me.Langue = "fr" Then
            ToolTip1.SetToolTip(BoutClear, "Effacer le préset en cours")
            BoutClear.Text = "Effacer"
        Else
            ToolTip1.SetToolTip(BoutClear, "Delete the current preset")
            BoutClear.Text = "Clear"
        End If


        '
        ' Bouton Copier 
        Panneau1.Panel1.Controls.Add(BoutCopier)
        BoutCopier.TabStop = False
        p.X = 20
        p.Y = p.Y + 51
        s.Width = 100
        s.Height = 40
        '
        BoutCopier.FlatStyle = FlatStyle.Flat
        BoutCopier.FlatAppearance.BorderColor = Color.Gray
        BoutCopier.FlatAppearance.BorderSize = 2
        BoutCopier.FlatAppearance.MouseDownBackColor = Color.Gold
        BoutCopier.FlatAppearance.MouseOverBackColor = Color.GreenYellow



        BoutCopier.BackColor = Color.FromArgb(98, 180, 222) 'Bleu
        BoutCopier.Font = f1
        BoutCopier.Location = p
        BoutCopier.Size = s

        If Me.Langue = "fr" Then
            BoutCopier.Text = "Copier"
        Else
            BoutCopier.Text = "Copy"
        End If

        ' Bouton Coller
        Panneau1.Panel1.Controls.Add(BoutColler)
        BoutColler.TabStop = False
        p.X = 20
        p.Y = p.Y + 51
        s.Width = 100
        s.Height = 40
        '
        BoutColler.FlatStyle = FlatStyle.Flat
        BoutColler.FlatAppearance.BorderColor = Color.Gray
        BoutColler.FlatAppearance.BorderSize = 2
        BoutColler.FlatAppearance.MouseDownBackColor = Color.Gold
        BoutColler.FlatAppearance.MouseOverBackColor = Color.GreenYellow

        BoutColler.BackColor = Color.LightYellow 'Color.FromArgb(59, 51, 124) ' violet
        BoutColler.ForeColor = Color.Black
        BoutColler.Font = f1
        BoutColler.Location = p
        BoutColler.Size = s

        If Me.Langue = "fr" Then
            BoutColler.Text = "Coller"
        Else
            BoutColler.Text = "Paste"
        End If
        '
        ' Bouton Attacher/détacher
        Panneau1.Panel1.Controls.Add(DockButton)
        DockButton.TabStop = False
        p.X = 293
        p.Y = p.Y + 105
        s.Width = 50
        s.Height = 45
        '
        DockButton.FlatStyle = FlatStyle.Standard
        DockButton.BackColor = Color.DarkOliveGreen
        DockButton.ForeColor = Color.Yellow
        DockButton.Font = fnt5
        DockButton.Location = p
        DockButton.Size = s
        DockButton.AutoSize = True
        '
        If Me.Langue = "fr" Then
            DockButton.Text = "Détacher"
        Else
            DockButton.Text = "Undock"
        End If
        DockButton.Visible = False
        '
        ' liste de présets
        Panneau1.Panel1.Controls.Add(ListPrésets)
        p = LVar.Item(4).Location
        p.Y = p.Y + 52
        '
        s.Height = 200
        s.Width = 160
        '
        ListPrésets.Location = p
        ListPrésets.Size = s
        '
        Maj_ListPrésets()

        ListPrésets.Font = fnt3
        ' Bouton de sauvegarde de présets
        Panneau1.Panel1.Controls.Add(SauvPrésets)
        SauvPrésets.TabStop = False
        p.Y = p.Y + 208
        s.Width = 160
        s.Height = 50
        '
        SauvPrésets.Location = p
        SauvPrésets.Size = s
        '
        If Me.Langue = "fr" Then
            SauvPrésets.Text = "Gestion des Presets Perso"
        Else
            SauvPrésets.Text = "Perso Presets Management"
        End If


        '
        'AddHandler NoteClic.CheckedChanged, AddressOf NoteClic_CheckedChanged
    End Sub

    Private Sub Maj_ListPrésets()
        Dim tbl() As String
        For Each a As String In ListDesPrésets
            tbl = a.Split(";")
            ListPrésets.Items.Add(tbl(0))
        Next
    End Sub



    Function Det_LettrePreset(i As Integer) As String
        Dim a As String = ""
        Select Case i
            Case 0
                a = "A"
            Case 1
                a = "B"
            Case 2
                a = "C"
            Case 3
                a = "D"
            Case 4
                a = "E"
            Case 5
                a = "F"
            Case 6
                a = "G"
            Case 7
                a = "H"
        End Select
        Return Trim(a)
    End Function


    Sub Maj_ToolTip()

        ToolTip1.RemoveAll() ' le RemoveAll est obligatoire pour faire réapparaître les bulles à chaque Undock (autrement, elles n'apparaissent que sur le 1er Undock) 
        ToolTip1.InitialDelay = 250
        ToolTip1.Active = True
        ToolTip1.ShowAlways = True

        If Langue = "fr" Then
            ListTypNote.Items.Add("RN")
            ListTypNote.Items.Add("BL")
            ListTypNote.Items.Add("NR")
            ListTypNote.Items.Add("CR")
            ListTypNote.Items.Add("DC")
            LabelTypNote.Text = "Durée"
            ToolTip1.SetToolTip(ListTypNote, "RN:ronde BL:blanche NR:noire CR:croche DC:Double croche")
            '
            ToolTip1.SetToolTip(ListDynF2, "Dynamique")
        Else
            ListTypNote.Items.Add("WN")
            ListTypNote.Items.Add("HN")
            ListTypNote.Items.Add("QN")
            ListTypNote.Items.Add("EN")
            ListTypNote.Items.Add("SN")
            LabelTypNote.Text = "Length"
            ToolTip1.SetToolTip(ListTypNote, "WH:whole HN:half QN:quarter EN:eighth  SN:sixteen")
            '
            ToolTip1.SetToolTip(ListDynF2, "Velocity")
        End If
    End Sub


    Public Sub CheckMute_CheckedChanged(sender As Object, e As EventArgs)
        Form1.Mix.muteVolume.Item(Me.Canal).Checked = CheckMute.Checked
    End Sub


    Private Sub SoloBoutDRM_MouseUp(sender As Object, e As EventArgs)
        Form1.Mix.GestMute()
        TraitementSoloDRM(9)
    End Sub
    Public Sub TraitementSoloDRM(CanMidi As Integer)
        Dim k As Integer
        If (Form1.SoloCours2 = True And Form1.CanMidiCours = CanMidi) Or Form1.SoloCours2 = False Then
            Form1.Mix.Gestion_Solo2(CanMidi) ' gestion du solo dans la table de mixage
            ' gestion bouton solo dans toute l'appli
            Form1.GestSolo(Form1.SoloCours2)
            ' activation du mode solo
            If Form1.SoloCours2 = False Then
                k = CanMidi
                SoloBoutDRM.BackColor = Color.OrangeRed
                SoloBoutDRM.ForeColor = Color.Yellow
                SoloBoutDRM.Enabled = True
            Else
                k = -1
                SoloBoutDRM.BackColor = Color.Beige
                SoloBoutDRM.ForeColor = Color.Black
            End If

            Form1.PisteSolo = k
            '
            If Form1.SoloCours2 = False Then
                Form1.CanMidiCours = CanMidi
                Form1.SoloCours2 = True
            Else
                Form1.CanMidiCours = -1
                Form1.SoloCours2 = False
            End If
        End If
    End Sub
    Private Sub BoutInit_MouseClick(sender As Object, e As EventArgs)
        Dim mess, titre As String
        Dim i As Integer

        If Module1.LangueIHM = "fr" Then
            mess = "Initialiser le mappage ?"
            titre = "Demande de confirmation"
        Else
            mess = "Init Mapping ?"
            titre = "Confirmation request"
        End If

        Cacher_FormTransparents()
        i = MessageBox.Show(mess, titre, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If i = vbYes Then
            BasePrésets()
        End If
    End Sub
    Private Sub Boutclear_MouseClick(sender As Object, e As EventArgs)
        Dim mess, titre As String
        Dim i As Integer

        If Module1.LangueIHM = "fr" Then
            mess = "Veuillez confirmer l'effacement."
            titre = "Demande de confirmation"
        Else
            mess = "Please confirm deletion"
            titre = "Confirmation request"
        End If

        Cacher_FormTransparents()
        i = MessageBox.Show(mess, titre, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If i = vbYes Then
            EffacerTout()
        End If
    End Sub
    Private Sub BoutCopier_MouseClick(sender As Object, e As EventArgs)
        Dim i, j As Integer

        BufferCCC = ""
        For i = (Grid1.FixedRows) To Grid1.Rows - 1
            For j = 4 To Grid1.Cols - 1
                If Trim(Grid1.Cell(i, j).Text) <> "" Then
                    BufferCCC = BufferCCC + Convert.ToString(i) + " " + Convert.ToString(j) + " " + Grid1.Cell(i, j).Text + ","
                End If
            Next
        Next
        If BufferCCC <> "" Then
            BufferCCC = Microsoft.VisualBasic.Left(BufferCCC, Len(BufferCCC) - 1)
        End If
    End Sub
    Private Sub BoutColler_MouseClick(sender As Object, e As EventArgs)
        Dim tbl1() As String
        Dim tbl2() As String
        Dim i, l, c, d As Integer
        If BufferCCC <> Nothing Then
            tbl1 = BufferCCC.Split(",")
            For i = 0 To tbl1.Count - 1
                tbl2 = tbl1(i).Split()
                l = tbl2(0)
                c = tbl2(1)
                d = tbl2(2)
                Grid1.Cell(l, c).Text = Trim(d)
            Next
            i = Det_NumPréset(Trim(Grid1.Cell(0, 0).Text))
            Maj_Préset(i)
        End If
    End Sub
    Private Sub DockButton_MouseClick(sender As Object, e As EventArgs)
        Dim NumOnglet As Integer = Convert.ToUInt16(F2.Tag)

        If Me.F2.Dock = DockStyle.Fill Then
            Me.F2.FormBorderStyle = FormBorderStyle.FixedDialog ' la fenêtre drum edit n'a pas besoin d'être sizable
            Me.F2.TopMost = True   ' 
            Me.F2.Dock = DockStyle.None
            F2.MainMenuStrip.Visible = True
            F2.MainMenuStrip.Enabled = True
            Form1.TabControl4.TabPages.Item(NumOnglet).Controls.Remove(Me.F2)
            Dim p As New Point With {
                .X = Me.F2.Location.X,
                .Y = Me.F2.Location.Y + 50
            }
            Me.F2.Location = p
            Me.F2.StartPosition = FormStartPosition.Manual ' permet de tenir compte de la location calculée dans p
            '
            Me.F2.TopLevel = True
            If LangueIHM = "fr" Then
                DockButton.Text = "Attacher"
            Else
                DockButton.Text = "Dock"
            End If
            If CompilPerso = True Then
                'Opac.Visible = True ' réglage de l'opacity du formulaire déDocké
            End If
            Maj_ToolTip()

        Else
            Me.F2.FormBorderStyle = FormBorderStyle.None
            Me.F2.TopMost = False   ' un seul des 2 suffit ?
            Me.F2.TopLevel = False
            F2.MainMenuStrip.Visible = False
            F2.MainMenuStrip.Enabled = False
            Form1.TabControl4.TabPages.Item(NumOnglet).Controls.Add(Me.F2)
            Me.F2.Dock = DockStyle.Fill
            '
            If LangueIHM = "fr" Then
                DockButton.Text = "Détacher"
            Else
                DockButton.Text = "UnDock"
            End If
            Maj_ToolTip()
        End If
    End Sub
    Private Sub SauvPrésets_MouseClick(sender As Object, e As EventArgs)
        ' My.Computer.FileSystem.OpenTextFileWriter("C: \préset.txt", True) ' append = true pour ajouter de nouvelles lignes
        '  My.Computer.FileSystem.FileExists
        '  FileOpen(1, FichierOuvrir, OpenMode.Input) ' Ouvre en lecture
        ' 


        Dim DossierDocuments As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        Module1.Création_CTemp()
        Dim path As String
        path = DossierDocuments + "\HyperArp"
        path = path + "\ConstructionPréset.txt"
        If Not (My.Computer.FileSystem.FileExists(path)) Then
            Dim fs As FileStream = File.Create(path)
            fs.Close()
        End If
        Form2.TopLevel = True
        Form2.TopMost = True
        Form2.ShowDialog()    ' Dialog()
    End Sub
    Private Sub Lvar_Click(sender As Object, e As EventArgs)
        Dim com As Button = sender
        Dim ind As Integer
        Dim Coul As Color

        ind = Val(com.Tag)
        Grid1.AutoRedraw = False
        '
        Coul = Det_BackColorDrum(ind)
        Grid1.BackColorFixedSel = Coul
        Grid1.BackColorFixed = Coul
        Grid1.Cell(0, 0).BackColor = Color.White
        Grid1.Cell(0, 0).ForeColor = Color.Black
        Grid1.Cell(0, 0).Font = fnt4
        Grid1.Cell(0, 0).Text = LPréset(ind) '"D" + Convert.ToString(ind + 1)

        ' raz des notes sur préset actuel
        Dim i, j As Integer

        For i = 1 To Grid1.Rows - 1
            For j = 4 To Grid1.Cols - 1
                Grid1.Cell(i, j).Text = ""
            Next
        Next
        '
        ' écriture des notes du Préset
        If Préset.Item(ind).LNotes.Count <> 0 Then
            For Each aa As CNotes In Préset.Item(ind).LNotes
                Grid1.Cell(aa.lig, aa.col).Text = aa.dyn
            Next
        End If
        '
        ' mise à jour du nom du préset dans le textbox NomPreset
        ' ******************************************************
        NomPréset.Text = LNomPréset.Item(ind)

        Grid1.AutoRedraw = True
        Grid1.Refresh()
        '
    End Sub
    Sub ChoixPréset(ind As Integer)
        Dim i, j As Integer

        'DéVérouillerGrid1()
        Grid1.AutoRedraw = False
        '
        Coul = Det_BackColorDrum(ind)
        Grid1.BackColorFixedSel = Coul
        Grid1.BackColorFixed = Coul
        Grid1.Cell(0, 0).BackColor = Color.White
        Grid1.Cell(0, 0).ForeColor = Color.Black
        Grid1.Cell(0, 0).Font = fnt4
        Grid1.Cell(0, 0).Text = LPréset(ind) '"D" + Convert.ToString(ind + 1)


        ' raz des notes sur préset actuel

        For i = 1 To Grid1.Rows - 1
            For j = 4 To Grid1.Cols - 1
                Grid1.Cell(i, j).Text = ""
            Next
        Next
        '
        ' écriture des notes du Préset
        If Préset.Item(ind).LNotes.Count <> 0 Then
            For Each aa As CNotes In Préset.Item(ind).LNotes
                Grid1.Cell(aa.lig, aa.col).Locked = False
                Grid1.Cell(aa.lig, aa.col).Text = aa.dyn
                Grid1.Cell(aa.lig, aa.col).Locked = True
            Next
        End If

        Grid1.AutoRedraw = True
        Grid1.Refresh()
        'VérouillerGrid1()
    End Sub



    ''' <summary>
    ''' Détermination de la couleur des lignes fixes du tableau de Drum Edit en fonction de la variation
    ''' </summary>
    ''' <param name="numVar">N° de la variation de Drum Edit de 0 à 7</param>
    ''' <returns>RETOUR : une couleur (Color obj)</returns>
    Function Det_BackColorDrum(numVar As Integer) As Color

        Select Case numVar
            Case 0
                Det_BackColorDrum = CoulA
            Case 1
                Det_BackColorDrum = CoulB
            Case 2
                Det_BackColorDrum = CoulC
            Case 3
                Det_BackColorDrum = CoulD
            Case 4
                Det_BackColorDrum = CoulE
            Case 5
                Det_BackColorDrum = CoulF
            Case 6
                Det_BackColorDrum = CoulG
            Case 7
                Det_BackColorDrum = CoulH
            Case 8 '  "--" --> pas de batterie 
                Det_BackColorDrum = CoulRien
        End Select

    End Function
    Function Det_BackColorDrumLettre(Lettre As String) As Color

        Select Case Lettre
            Case "A"
                Det_BackColorDrumLettre = CoulA
            Case "B"
                Det_BackColorDrumLettre = CoulB
            Case "C"
                Det_BackColorDrumLettre = CoulC
            Case "D"
                Det_BackColorDrumLettre = CoulD
            Case "E"
                Det_BackColorDrumLettre = CoulE
            Case "F"
                Det_BackColorDrumLettre = CoulF
            Case "G"
                Det_BackColorDrumLettre = CoulG
            Case "H"
                Det_BackColorDrumLettre = CoulH
            Case "Vide" '  "--" --> pas de batterie 
                Det_BackColorDrumLettre = CoulRien
        End Select

    End Function

    Function Det_BackColorDrum2(Préset As String) As Color

        Select Case Trim(Préset)
            Case "A"
                Det_BackColorDrum2 = CoulA
            Case "B"
                Det_BackColorDrum2 = CoulB
            Case "C"
                Det_BackColorDrum2 = CoulC
            Case "D"
                Det_BackColorDrum2 = CoulD
            Case "E"
                Det_BackColorDrum2 = CoulE
            Case "F"
                Det_BackColorDrum2 = CoulF
            Case "G"
                Det_BackColorDrum2 = CoulG
            Case "H"
                Det_BackColorDrum2 = CoulH
            Case "--" ' 8 '  "--" --> pas de batterie 
                Det_BackColorDrum2 = CoulRien
        End Select
    End Function
    Sub Construction_F2()
        Enchargement = True
        Construction_SplitContainer()
        Construction_Grid1()
        Construction_Grid2()
        Construction_Menu()
        Construction_MenuContext()
        Construction_Barrout()
        Enchargement = False
    End Sub
    ''' <summary>
    ''' Construction_MenuContext : uitilisé sur la time line
    ''' </summary>
    Public Sub Construction_MenuContext()
        Dim i As Integer

        ' Menu Couper
        If Langue = "fr" Then
            itemCouper.Text = "Couper"
        Else
            itemCouper.Text = "Cut"
        End If
        MenuContext1.Items.Add(itemCouper)
        '
        ' Menu Copier
        If Langue = "fr" Then
            itemCopier.Text = "Copier"
        Else
            itemCopier.Text = "Copy"
        End If
        MenuContext1.Items.Add(itemCopier)

        '
        ' Menu Copier
        If Langue = "fr" Then
            itemColler.Text = "Coller"
        Else
            itemColler.Text = "Paste"
        End If
        MenuContext1.Items.Add(itemColler)

        '
        ' Séparateur
        MenuContext1.Items.Add(Séparateur3)
        ' Séparateur1.Size = New System.Drawing.Size(177, 6)
        '


        ' Menu Préset
        itemPresets.Text = "Présets"
        MenuContext1.Items.Add(itemPresets)
        'AddHandler itemPresets.Click, AddressOf itemPresets_Click
        '
        ' Définition de la liste des présets
        For i = 0 To 7

            LPrests.Add(New ToolStripMenuItem)
            itemPresets.DropDownItems.Add(LPrests(i))
            Select Case i
                Case 0
                    LPrests(i).Text = "A"
                    LPrests(i).Tag = "A"
                Case 1
                    LPrests(i).Text = "B"
                    LPrests(i).Tag = "B"
                Case 2
                    LPrests(i).Text = "C"
                    LPrests(i).Tag = "C"
                Case 3
                    LPrests(i).Text = "D"
                    LPrests(i).Tag = "D"
                Case 4
                    LPrests(i).Text = "E"
                    LPrests(i).Tag = "E"
                Case 5
                    LPrests(i).Text = "F"
                    LPrests(i).Tag = "F"
                Case 6
                    LPrests(i).Text = "G"
                    LPrests(i).Tag = "G"
                Case 7
                    LPrests(i).Text = "H"
                    LPrests(i).Tag = "H"
            End Select
            AddHandler LPrests(i).Click, AddressOf LPrests_Click
        Next
        AddHandler itemCouper.Click, AddressOf itemCouper_Click
        AddHandler itemCopier.Click, AddressOf itemCopier_Click
        AddHandler itemColler.Click, AddressOf itemColler_Click
        'AddHandler itemPresets.Click, AddressOf itemPresets_Click

    End Sub
    Sub AddMarq(Marq As String, Position As String)

    End Sub
    Sub DeleteMarq(Position As String)

    End Sub
    Sub F2_Refresh_old()
        Dim i, j As Integer
        Dim tbl() As String
        Dim tbl1() As String
        Dim a As String = Me.ListAcc

        Grid2.AutoRedraw = False
        '
        ' Mise à Jour des couleurs
        ' ***********************
        'Grid2.BackColorFixed = DrmC_GrisBleu

        For i = 0 To Grid2.Rows - 2
            For j = 1 To Grid2.Cols - 1
                Grid2.Cell(i, j).BackColor = DrmC_GrisBleu
            Next
        Next

        Dim DerCol As Integer = Form1.Det_NumDerAccord + (Form1.Det_NumDerRépet - 1)
        For i = 0 To Grid2.Rows - 2
            For j = 1 To DerCol
                Grid2.Cell(i, j).BackColor = Color.Ivory
            Next
        Next
        '
        ' Maj des accords
        ' ***************
        tbl = a.Split("-")
        i = 1
        For Each aa As String In tbl
            Grid2.Cell(2, i).Text = aa
            i = i + 1
        Next
        '
        ' Maj des marqueurs
        ' *****************
        '
        ' raz de tous les marqueurs
        For j = 1 To Grid2.Cols - 1
            Grid2.Cell(1, j).Text = ""
        Next
        '
        ' maj des marqueurs
        a = Me.ListMarq
        If Trim(a) <> "" Then
            tbl = a.Split("&")
            i = 1
            Grid2.Range(1, 1, Grid2.Rows - 1, Grid2.Cols - 1).Locked = False
            For Each aa As String In tbl
                tbl1 = aa.Split(";")
                i = Convert.ToInt16(tbl1(1))
                Grid2.Cell(1, i).Text = Trim(tbl1(0))
                Grid2.Cell(1, i).BackColor = Color.AliceBlue 'DarkOliveGreen
                Grid2.Cell(1, i).ForeColor = Color.Red
                Grid2.Refresh()
            Next
            'Grid2.Range(1, 1, Grid2.Rows - 1, Grid2.Cols - 1).Locked = True
        End If
        '


        Grid2.AutoRedraw = True
        Grid2.Refresh()
    End Sub
    Sub F2_Refresh()
        Dim i, j As Integer
        Dim tbl(), tbl1(), tbl2() As String
        Dim a As String = Me.ListAcc
        Dim b As String = Me.ListMarq
        Dim c As String

        Grid2.AutoRedraw = False
        '
        ' Mise à Jour des couleurs
        ' ***********************
        '
        For i = 0 To Grid2.Rows - 2
            For j = 1 To Grid2.Cols - 1
                Grid2.Cell(i, j).BackColor = DrmC_GrisBleu
                Grid2.Cell(i, j).ForeColor = Color.Black
            Next
        Next
        '
        ' Maj des accords
        ' ***************
        ' raz de toute la ligne des accords
        For j = 1 To Grid2.Cols - 1
            Grid2.Cell(2, j).Text = ""
        Next

        If Trim(a) <> "" Then
            ' maj des accords
            tbl = a.Split(";")
            For i = 0 To tbl.Count - 1
                tbl1 = tbl(i).Split("-")
                tbl2 = tbl1(0).Split(".")
                j = Convert.ToInt16(tbl2(0)) ' N° de mesure de l'accord
                c = EventHDsMesure("Accord", j)
                Grid2.Cell(2, j).Text = Trim(c)
                Grid2.Cell(2, j).BackColor = Color.Ivory
            Next
        End If
        '
        ' Maj des marqueurs
        ' *****************
        '
        ' raz de tous les marqueurs
        For j = 1 To Grid2.Cols - 1
            Grid2.Cell(1, j).Text = ""
        Next
        '
        ' maj des marqueurs
        a = Me.ListMarq
        If Trim(a) <> "" Then
            tbl = a.Split("&")
            i = 1
            Grid2.Range(1, 1, 1, Grid2.Cols - 1).Locked = False
            For Each aa As String In tbl
                tbl1 = aa.Split(";")
                i = Convert.ToInt16(tbl1(1))
                Grid2.Cell(1, i).Text = Trim(tbl1(0))
                Grid2.Cell(1, i).BackColor = Color.AliceBlue 'DarkOliveGreen
                Grid2.Cell(1, i).ForeColor = Color.Red
                Grid2.Refresh()
            Next
            Grid2.Range(1, 1, 1, Grid2.Cols - 1).Locked = True
        End If
        '
        Grid2.AutoRedraw = True
        Grid2.Refresh()
    End Sub
    Private Function EventHDsMesure(Eventh As String, Mesure As Integer) As String
        Dim t As Integer
        Dim a, b As String

        a = ""
        b = ""
        For t = 0 To 5
            For ct = 0 To 4 '
                Select Case Eventh
                    Case "Accord"
                        b = TableEventH(Mesure, t, ct).Accord

                    Case "Gamme"
                        b = TableEventH(Mesure, t, ct).Gamme
                    Case "Marqueur"
                        b = TableEventH(Mesure, t, ct).Marqueur
                End Select
                If Trim(b) <> "" Then
                    a = a + b + "/"
                End If
            Next ct
        Next t
        If Trim(a) <> "" Then
            a = Microsoft.VisualBasic.Left(a, Len(a) - 1)
        End If
        Return a
    End Function
    Function Det_DivisionMes() As Integer
        Select Case Me.Dénominateur
            Case 4
                Det_DivisionMes = Me.Numérateur * 4
            Case 8
                Det_DivisionMes = Me.Numérateur * 2
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
    Public Function Det_Préset() As String
        Det_Préset = Trim(Grid1.Cell(0, 0).Text)
    End Function


    Public Function Det_NumPréset(a As String) As Integer
        Select Case a
            Case "A"
                Det_NumPréset = 0
            Case "B"
                Det_NumPréset = 1
            Case "C"
                Det_NumPréset = 2
            Case "D"
                Det_NumPréset = 3
            Case "E"
                Det_NumPréset = 4
            Case "F"
                Det_NumPréset = 5
            Case "G"
                Det_NumPréset = 6
            Case "H"
                Det_NumPréset = 7
            Case Else
                Det_NumPréset = -1
        End Select
    End Function
    Public Function Det_NumPréset2() As Integer
        Det_NumPréset2 = Det_NumPréset(Trim(Grid1.Cell(0, 0).Text))
    End Function
    ''' <summary>
    ''' CellDyn : détermination des cellules concernées par une décrémentation/incrémentation
    ''' </summary>
    ''' <returns> retourne une sélection de "ligne,colonne".
    ''' Les valeurs "ligne,colonne" sont retournées dans une chaîne de caractères
    ''' et sont séparées par des blancs. Chaque "ligne,colonne" est séparé par une ",".</returns>
    Function CellDyn() As String
        Dim i, j As Integer
        Dim a As String = ""
        CellDyn = a

        For i = Grid1.Selection.FirstRow To Grid1.Selection.LastRow
            For j = Grid1.Selection.FirstCol To Grid1.Selection.LastCol
                If IsNumeric(Grid1.Cell(i, j).Text) Then
                    a = a + Convert.ToString(i) + "," + Convert.ToString(j) + " "
                End If
            Next
        Next
        CellDyn = Trim(a)
    End Function
    Sub Maj_Préset(NPrés As Integer)
        Dim i, j As Integer
        Dim c1, c2 As Integer
        Dim posit As String

        Préset.Item(NPrés).LNotes.Clear()
        '
        For j = 4 To Grid1.Cols - 1
            For i = 1 To Grid1.Rows - 1
                If Trim(Grid1.Cell(i, j).Text) <> "" Then
                    ' calcul position dans le préset
                    c1 = Fix(j / 4)             ' chiffre 1
                    c2 = (j - (4 * c1)) + 1     ' chiffre 2
                    posit = Convert.ToString(c1) + "." + Convert.ToString(c2)
                    Dim aa As New CNotes With {
                        .lig = i,
                        .col = j,
                        .dyn = Trim(Grid1.Cell(i, j).Text),
                        .Pos = Trim(posit)
                    }
                    Préset.Item(NPrés).LNotes.Add(aa)
                End If
            Next
        Next
    End Sub
    Sub Ecr_Préset(Nprés As Integer)
        If Préset.Item(Nprés).LNotes.Count <> 0 Then
            For Each aa As CNotes In Préset.Item(Nprés).LNotes
                Grid1.Cell(aa.lig, aa.col).Text = Trim(aa.dyn)
            Next
        End If
    End Sub

    Private Function NbDivInduit(Dénominateur) As Integer
        Dim i As Integer = 4
        Select Case Dénominateur
            Case 4
                i = 4
            Case 8
                i = 2
        End Select
        Return i
    End Function
    Public Function ChaineModèle() As String
        Dim num_DerAcc As Integer = Form1.Det_NumDerAccord
        Dim LigneModèle As Integer = Det_LigneModèle()
        Dim a As String = ""
        Dim c, m, p As Integer

        If LigneModèle <> -1 Then
            'lecture Préset sur Mesure m
            For m = 1 To num_DerAcc
                p = Det_NumPréset(Grid2.Cell(3, m).Text) ' n° de préset de la mesure m
                If p <> -1 Then ' -1 correspond à une absence de batterie sur la mesure
                    For Each aa As CNotes In Préset.Item(p).LNotes
                        If aa.lig = LigneModèle Then
                            ' calcul colonne
                            c = (aa.col + ((m - 1) * 16)) - 3
                            a = a + c.ToString + " "
                        End If
                    Next
                End If
            Next
        End If
        Return Trim(a)
    End Function
    Function Det_LigneModèle() As Integer
        Dim ligne As String = -1
        For i = 1 To Grid1.Rows - 1
            If Grid1.Cell(i, 1).BackColor = Module1.Couleur_ModelDrum Then
                ligne = i
                Exit For
            End If
        Next
        Return ligne
    End Function

    Public Sub VérouillerGrid1()
        PourVerrou(True)
    End Sub
    Public Sub DéVérouillerGrid1()
        PourVerrou(False)
    End Sub
    Sub PourVerrou(b As Boolean)
        Dim i, j As Integer
        'Grid1.Range(1, 4, Grid1.Rows - 1, 4).Locked = False
        Grid1.AutoRedraw = True
        For i = 1 To Grid1.Rows - 1
            For j = 4 To Grid1.Cols - 1
                Grid1.Cell(i, j).Locked = b
            Next j
        Next i
        '
        Grid1.AutoRedraw = True
        Grid1.Refresh()
    End Sub
    Sub JouerNote(ValeurNote As Byte, Dyn As Byte)
        'Dim Canal As Byte = Microsoft.VisualBasic.Right(Trim(Me.CheckMute.Text), 1) ' N° piste, N° canal
        Try

            If Not (Form1.SortieMidi.Item(Form1.ChoixSortieMidi).IsOpen) Then
                Form1.SortieMidi.Item(Form1.ChoixSortieMidi).Open()
            End If
            '
            Form1.SortieMidi.Item(Form1.ChoixSortieMidi).SendNoteOn(Me.Canal, ValeurNote, Dyn)
            NoteCourante = ValeurNote
            NoteAEtéJouée = True
        Catch ex As Exception
            MessageHV.TypBouton = "OK"
            If Module1.LangueIHM = "fr" Then
                MessageHV.PTitre = "Problème de ressource MIDI"
                MessageHV.PContenuMess = "Warning : détection d'une erreur dans procédure " + "Pianoroll.JouerNote" + Constants.vbCrLf + "- " +
                "Message  : " + ex.Message + Constants.vbCrLf + "- " + "Choisissez une autre sortie MIDI"
            Else
                MessageHV.PTitre = "MIDI Resource Problem"
                MessageHV.PContenuMess = "- " + "Warning : procedure error detection  : " + "Pianoroll.JouerNote" + Constants.vbCrLf + "- " +
                "Message  : " + ex.Message + Constants.vbCrLf + "- " + "Choose another MIDI output"
            End If
            NoteCourante = ValeurNote
            NoteAEtéJouée = True
            Cacher_FormTransparents()
            MessageHV.ShowDialog()
            'End
        End Try
    End Sub
    Sub StoperNote(ValeurNote As Byte, Dyn As Byte)
        'Dim Canal As Byte = Microsoft.VisualBasic.Right(Trim(Me.CheckMute.Text), 1) ' N° piste, N° canal
        Dim i As Integer = Form1.ChoixSortieMidi
        Try
            If Not (Form1.SortieMidi.Item(Form1.ChoixSortieMidi).IsOpen) Then
                Form1.SortieMidi.Item(Form1.ChoixSortieMidi).Open()
            End If
            '
            Form1.SortieMidi.Item(Form1.ChoixSortieMidi).SendNoteOff(Me.Canal, ValeurNote, 0)
            NoteAEtéJouée = False
        Catch ex As Exception
            messa = "Problème de ressource MIDI"
            Avertis = messa + Constants.vbCrLf + "Détection d'une erreur dans procédure : " + "StoperNote" + "." + Constants.vbCrLf +
            "Message  : " + ex.Message
            Cacher_FormTransparents()
            MessageHV.ShowDialog()
            End
        End Try
    End Sub
    Public Sub DrumCopy()
        Dim LigneInitiale As Integer = Grid1.Selection.FirstRow
        Dim ColonneInitiale As Integer = Grid1.Selection.FirstCol



        If Grid1.Selection.FirstCol > 3 And Grid1.Selection.FirstRow > 0 Then
            ClipBPerso.Clear()
            'Grid1.Selection.CopyData()
            For i = Grid1.Selection.FirstRow To Grid1.Selection.LastRow
                For j = Grid1.Selection.FirstCol To Grid1.Selection.LastCol
                    If Trim(Grid1.Cell(i, j).Text) <> "" Then
                        'ClipBPerso.Add(i.ToString + "." + j.ToString + "-" + Trim(Grid1.Cell(i, j).Text + ";")
                        Dim oo As New ClipB With {
                            .EcartLigne = i - LigneInitiale,
                            .EcartCol = j - ColonneInitiale,
                            .value = Trim(Grid1.Cell(i, j).Text)
                        }
                        ClipBPerso.Add(oo)
                    End If
                Next
            Next
        End If
    End Sub
    Public Sub DrumPaste()

        Dim LigneInitiale As Integer = Grid1.Selection.FirstRow
        Dim ColonneInitiale As Integer = Grid1.Selection.FirstCol
        '


        If ClipBPerso.Count <> 0 Then

            ' caclul des décalages sur ligne et colonne entre la sélection copié et l'endroit où coller
            'd1 = ClipBPerso.Item(0).row - i
            'd2 = ClipBPerso.Item(0).col - j
            ' coller
            For Each a As ClipB In ClipBPerso
                i = LigneInitiale + a.EcartLigne
                j = ColonneInitiale + a.EcartCol
                '
                If i <= Grid1.Rows - 1 And j <= Grid1.Cols - 1 Then
                    Grid1.Cell(i, j).Text = Trim(a.value)
                End If

            Next
        End If
    End Sub
    '
    Public Sub DrumCut()
        Dim i As Integer = Grid1.Selection.FirstRow
        Dim j As Integer = Grid1.Selection.FirstCol
        Grid1.AutoRedraw = False
        If Grid1.Selection.FirstCol > 3 And Grid1.Selection.FirstRow > 0 Then
            For i = Grid1.Selection.FirstRow To Grid1.Selection.LastRow
                For j = Grid1.Selection.FirstCol To Grid1.Selection.LastCol
                    Grid1.Cell(i, j).Text = ""
                Next
            Next
        End If
        Grid1.AutoRedraw = True
        Grid1.Refresh()
    End Sub
End Class
