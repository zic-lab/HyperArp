Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Rebar

Public Class PianoRoll

    ' non utilisé
    Private _mRejouer As Boolean = False
    Public Event Rejouer(ByVal ordre As Boolean)
    Public Property PRejouer
        Get
            Return _mRejouer
        End Get
        Set(value)
            RaiseEvent Rejouer(True)
            _mRejouer = value
        End Set
    End Property
    '
    Private PresNotes As Boolean = False
    Public Property PPresNotes As Boolean
        Get
            Return PresNotes
        End Get
        Set(ByVal value As Boolean)
            PresNotes = value
        End Set
    End Property

    Private ToRow As Integer = 60
    Public Property PTopRow As Integer
        Get
            ToRow = Me.Grid1.TopRow
            Return ToRow
        End Get
        '
        Set(ByVal value As Integer)
            Grid1.TopRow = value
        End Set
    End Property
    '                                    
    Private ListTon As String
    Public WriteOnly Property PListTon As String
        Set(ByVal value As String)
            Me.ListTon = value
        End Set
    End Property

    Private Mute As Boolean
    Public Property PMute As Boolean
        Get
            Mute = Me.CheckMute.Checked
            Return Mute
        End Get
        Set(ByVal value As Boolean)
            Me.CheckMute.Checked = value
        End Set
    End Property
    ''' <summary>
    ''' Prpriété retournant une chaine de caractère contenant les notes à jouer d'une piste.
    ''' </summary>
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

    Private Langue As String
    Public WriteOnly Property PLangue As String
        Set(ByVal value As String)
            Me.Langue = value
        End Set
    End Property
    '
    Private NbMes As Integer = nbMesures
    Public WriteOnly Property PNbMesures As Integer
        Set(ByVal value As Integer)
            Me.NbMes = value
        End Set
    End Property
    Public Métrique As String = "4/4"
    Public WriteOnly Property PMétrique As String
        Set(ByVal value As String)
            Me.Métrique = value
        End Set
    End Property
    Private nbRépétitionMax As Integer = Module1.nbRépétitionMax
    Public WriteOnly Property PnbRépétitionMax As Integer
        Set(ByVal value As Integer)
            Me.nbRépétitionMax = value
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
    Private ListGam As String
    Public WriteOnly Property PListGam As String
        Set(ByVal value As String)
            Me.ListGam = value
        End Set
    End Property
    Private ListMarq As String
    Public WriteOnly Property PListMarq As String
        Set(ByVal value As String)
            Me.ListMarq = value
        End Set
    End Property
    Private Nb_CTRLS As Integer
    Public ReadOnly Property PNbCtrls As String
        Get
            Me.Nb_CTRLS = nbCourbes
            Return Me.Nb_CTRLS
        End Get
    End Property
    Public N_Can1erPianoR As Integer
    Public WriteOnly Property PN_Can1erPianoR As Integer
        Set(ByVal value As Integer)
            Me.N_Can1erPianoR = value
        End Set
    End Property
    Private StartMeasure As Integer
    Public WriteOnly Property PStartMeasure As Integer
        Set(ByVal value As Integer)
            Me.StartMeasure = value
            Grid1.LeftCol = ((Me.StartMeasure - 1) * 16) + 1 'positionnement sur la mesure cliquée
            '
        End Set
    End Property

    Public F1 As New Form ' remarque : l'intégration de F1 dans l'onglet (page) de TabControl4 se fait dans Form1/PIANOROLL_Création (dans form1)
    Private Grid1 As New FlexCell.Grid
    ' Note sans " "
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
    '

    ' Fonts
    ' *****
    ReadOnly fnt1 As New System.Drawing.Font("Calibri", 13, FontStyle.Regular)
    ReadOnly fnt2 As New System.Drawing.Font("Calibri", 15, FontStyle.Regular)
    ReadOnly fnt3 As New System.Drawing.Font("Tahoma", 10, FontStyle.Regular)
    ReadOnly fnt4 As New System.Drawing.Font("Calibri", 18, FontStyle.Bold)
    ReadOnly fnt5 As New System.Drawing.Font("Calibri", 12, FontStyle.Regular)
    ReadOnly fnt6 As New System.Drawing.Font("Verdana", 8.25, FontStyle.Regular)
    ReadOnly fnt7 As New System.Drawing.Font("Verdana", 7, FontStyle.Regular)
    ReadOnly fnt7_1 As New System.Drawing.Font("Verdana", 7, FontStyle.Bold)
    ReadOnly fnt8 As New System.Drawing.Font("Verdana", 6, FontStyle.Regular)
    '

    ' couleur des courbes
    'ReadOnly CoulMod As Color = Color.LightSeaGreen
    'ReadOnly CoulExp As Color = Color.Blue
    'ReadOnly CoulCC50 As Color = Color.YellowGreen
    'ReadOnly CoulCC51 As Color = Color.Orange
    'ReadOnly CoulCC52 As Color = Color.MediumTurquoise
    'ReadOnly CoulCC53 As Color = Color.Thistle
    'ReadOnly CoulPB As Color = Color.Yellow

    ' Squelette de l'IHM
    ' ******************
    Private Panneau2 As New SplitContainer ' contient Panneau1 dans sa partie supérieure et Panel1(courbes) dans sa partie inférieure
    Private Panneau1 As New SplitContainer ' contient la barre d'outil dans sa partie supérieur et le piano roll dans sa partie inférieure
    ' Eléments de courbes
    Private PanelCourbes As New Panel
    Private TabCourbes As New TabControl
    Private PageCourbes As New List(Of TabPage)

    Public GridCourbes As New List(Of FlexCell.Grid)

    Public AffCtrls As New CheckBox ' demande affichage de la partie contrôleur dans le pianoroll
    Private ActiveExpression As New CheckBox
    Public CCActif As New List(Of CheckBox)
    Private MidiLearn As New CheckBox
    Private AffMidiLearn As New Label
    '
    Private Numérateur As Integer
    Private Dénominateur As Integer
    Private DivisionMes As Integer
    '
    Private NoteCourante As Byte = 64
    Private NoteAEtéJouée As Boolean = False
    '
    Private NbAccords As Integer
    Private Chargé As Boolean = False
    '
    Private BoutonF1_1 As New Button
    Private BoutonF1_2 As New Button
    Private BoutonFermer As New Button
    '
    Public CheckMute As New CheckBox
    Public SoloBoutPR As New Button

    Private ListTypNote As New System.Windows.Forms.ComboBox
    Private LabelTypNote As New Label
    '
    Public ListDynF1 As New System.Windows.Forms.ComboBox
    Private LabelDyn As New Label
    '
    Private ListPRG As New System.Windows.Forms.ComboBox
    Private LabelPRG As New Label

    Private ListMod As New System.Windows.Forms.ComboBox
    Private LabelMod As New Label

    Dim ActionsPianoR As New Button
    Dim OuvrirCalques As New Button
    '
    Private CadreRadio As New GroupBox
    Private LabelPan As New Label
    '
    Private BRadio1 As New RadioButton
    Private BRadio2 As New RadioButton
    Private BRadio3 As New RadioButton
    '
    Private CheckAcc As New CheckBox
    Private CheckTonique As New CheckBox ' case à cocher Mode dans le cadre "Calque MIDI"
    Private ListTonNotes As New System.Windows.Forms.ComboBox
    Private CheckDrum As New CheckBox
    Private Check12_8 As New CheckBox
    '
    Private CadreMacroSel As New GroupBox
    Private CadreMIDIModels As New GroupBox
    Private DebMacroSel As New NumericUpDown
    Private TermeMacroSel As New NumericUpDown
    Private DestMacroSel As New NumericUpDown
    Private BMacroSel As New Button
    '
    Private CadreAllerVers As New GroupBox
    '
    Private CheckTop As New CheckBox
    Private Destination As New TextBox
    '
    Private DockButton As New Button
    Private Opacité As New NumericUpDown

    Public NomduSon As New TextBox
    Dim AffNomduSon As New Label
    '
    Private ZoomMoins As New CheckBox
    Private Const LargCol = 26
    Private Const HautLig = 26
    Private Const ValZoomMoins = 17
    Private Const ValZoomPlus = LargCol
    '
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
    Private TabNotes(36) As String
    '
    Private DernièreEntrée As String
    '
    ' CONSTANTES
    ' Controleurs MIDI
    ' N° des controleurs MIDI
    ' ***********************
    Private Const CMod = 1
    Private Const CVolume = 7
    Private Const CPAN = 10
    Private Const CExpr = 11
    Private Const CC50 = 50
    Private Const CC51 = 51
    Private Const CC52 = 52
    Private Const CC53 = 53
    Private Const CPed = 64
    '

    '
    Private Const nbCourbes = 7
    ' Timer de coupure de jeu de note sur mousedown
    Dim TimerStop As New Timer ' timer d'arrêt d'une note activée dans grid1_mousedown avant l'arrêt dans le grid1_mouseup

    '
    ReadOnly CAccords As Color = Color.FromArgb(57, 222, 130)
    ReadOnly CDrums As Color = Module1.Couleur_ModelDrum 'Color.FromArgb(1, 153, 229)
    '
    Private CompteurActive As Integer = 0
    Private ToolTip1 As New ToolTip
    Private Canal As Byte
    Private CoulBarOut As Color
    '
    ' trait de base pour les notes PianoRoll
    'Private ReadOnly Trait As String = Convert.ToString(Convert.ToChar(9644)) ' +=43
    ' Caractères spéciaux ASCI
    ' ************************
    Private ReadOnly PointBlanc As String = "◘" ' alt + 8
    Private ReadOnly CercleBlanc As String = "◙" ' alt +10
    Private ReadOnly NoteMusique As String = "♪" ' alt + 13
    Private ReadOnly NoteMusique2 As String = "♫" ' alt + 14
    Private ReadOnly Trait As String = "▬" ' alt + 22


    ' variables pour modifi valeur numérique avec flèches
    'Dim KeyUp As Boolean = False
    'Dim KeyDown As Boolean = False
    ' CLASSE pour Edition CCC afin de pouvoir distinguer si l'origine de commande (ctrl +c,x,v) vient du pianoroll ou de ses courbes 
    ' ******************************************************************************************************************************
    Class OrigPianoR
        Public N_Courbe As Integer = -1
        Public Orig1 As New OrigPianoCourbe
        Public Vérouillage As Boolean = False ' sertà vérrouiller la touche CTRL pour tracer un graphe pendant qu'elle est utilisée pour une Edition CCC

    End Class
    Public Orig_PianoR As New OrigPianoR
    '
    '  Variables pour assitance GridAssist1 pour Consonnace/Dissonance
    '  ***************************************************************
    '
    Public GridAssist1 As New FlexCell.Grid ' grille assistance Intervalles
    ' 1    2    3     4    5    6     7    8    9     10   11    12   13   14    15   16   17    18   19    20   21    22   23    24
    Dim Chromad As New List(Of String) From {"c", "c#", "d", "d#", "e", "f", "f#", "g", "g#", "a", "a#", "b", "c", "c#", "d", "d#", "e", "f", "f#", "g", "g#", "a", "a#", "b"}
    Dim Chromab As New List(Of String) From {"c", "db", "d", "eb", "e", "f", "gb", "g", "ab", "a", "bb", "b", "c", "db", "d", "eb", "e", "f", "gb", "g", "ab", "a", "bb", "b"}
    ' Constantes
    Dim Octave As Integer = 12
    Dim Sept As Integer = 11
    Dim Septb As Integer = 10
    Dim SixteMaj As Integer = 9
    Dim SixteMin As Integer = 8
    Dim Quinte As Integer = 7
    Dim Quinteb As Integer = 6
    Dim Quarte As Integer = 5
    Dim TierceMaj As Integer = 4
    Dim TierceMin As Integer = 3
    Dim Neuf As Integer = 2
    Dim Neufb As Integer = 1
    Dim Tonique As Integer = 12
    ' ******
    Dim EstOctave As Integer = 12
    Dim EstSept As Integer = 1
    Dim EstSeptb As Integer = 2
    Dim EstSixteMaj As Integer = 3
    Dim EstSixteMin As Integer = 4
    Dim EstQuinte As Integer = 5
    Dim EstQuinteb As Integer = 6
    Dim EstQuarte As Integer = 7
    Dim EstTierceMaj As Integer = 8
    Dim EstTierceMin As Integer = 9
    Dim EstNeuf As Integer = 10
    Dim EstNeufb As Integer = 11
    Dim EstTonique As Integer = 12

    Const Col_CF = 13
    Const Maxcol = 12 + 1 + 12
    Dim CouleurMarq As New Color
    '
    Dim ft1 As New System.Drawing.Font("Tahoma", 9, FontStyle.Bold)
    '
    '  Variables pour assitance GridAssist2 pour Tonalités
    '  ***************************************************
    '
    Public GridAssist2 As New FlexCell.Grid ' grille assistance tonalités
    Public Assist2SplitC As New SplitContainer
    Public Assist2Bout As New Button
    Public Assist2Deb As New System.Windows.Forms.NumericUpDown
    Public Assist2Fin As New System.Windows.Forms.NumericUpDown
    Public Assist2Text1 As New TextBox
    Public Assist2Lab2 As New Label
    Public Assist2Lab3 As New Label
    Public Assist2Lab4 As New Label
    '
    Public ListeDesNotes As List(Of String)
    '
    '
    Dim ft2 As New System.Drawing.Font("Tahoma", 7, FontStyle.Regular)
    Dim lNotes As List(Of String)
    Public flagAssist2_Z As Boolean = False ' pour annuler l'écriture accord/gammes/tona créée à partir de PianoRoll

    ' Pour Calques MIDI
    ' *****************
    Public ChoixCalquesLocal(0 To nbCalques - 1) As Boolean
    Public PédaleLocale As Integer = 0
    Public TessDebLocale As String = "C-2"
    Public TessFinLocale As String = "G8"
    Public TessListeLocale As Integer = 0
    Dim ListPédale As New List(Of String)
    Public NouvCalqueLocale As Boolean = False


    ' Menu contextuel (pour vélocité aléatoire)
    ' *****************************************
    ' Menu contextuel
    ' ***************
    Dim MenuContext2 As New ContextMenuStrip
    Dim itemVélo As New ToolStripMenuItem





    ' remarque sur Numérateur et dénominateur : 
    '   - Par exemple dans 3/4 : 3 est le numérateur, 4 est le dénominateur
    '   - Le dénominateur donne l'unité de base constituant une mesure : 
    '       * soit en noires(4), par exemple 4/4
    '       * soit en croches(8), par exemple 7/8
    ' - Le numérateur donne le nombre d'unités constituant une mesure
    '       * par exemple, 7/4 indique qu'une mesure constituée de 7 noires
    '       * autre exemple 7/8 indique qu'une mesure constituée de 7 croches.
    ' - Pour exprimer la longueur d'une mesure avec des divisions en  double croches, il faut tenir compte que :
    '       * l'unité 4 (noire) est composée de 4 double croches, donc une mesure à 4/4 est composée
    '         de 4 noires x 4 (=valeur d'une noire en  double croche) = 16 divisions de double croches , une mesure à 7/4 est composée de
    '         7  noires x x 4 (=valeur d'une noire en double croche) = 28  divisions de double croches dans une mesure.
    '       * L'unité 8 (croche) est composée de 2 doubles croches, donc une mesure en 7/8 est composée
    '         de 7x2 = 14 divisions de double croches, une mesure en 12/8 est composée de 12*2=4 divisions de double croches.

    ''' <summary>
    ''' Cette méthode New est le constructeur de la classe. On y définit pricipalement tous les 
    ''' évènements des éléments graphiques de la classe : forumuaire F1, Grille Grid1 et autres boutons de la barre d'outil.
    ''' La prochaine méthode appelée est : Construction_F1 (appel lors de l'instanciation du Piano Roll dans form1/../ PIANOROLL_Création2)
    ''' </summary>
    ''' <param name="Canal">N° Canal MIDI du PianoRoll</param>
    ''' <param name="CoulBarout">Couleur de la barre d'outils</param>

    Sub New(Canal As Byte, CoulBarout As Color)
        AddHandler F1.Load, AddressOf F1_Load
        AddHandler F1.Activated, AddressOf F1_Activated
        AddHandler F1.Deactivate, AddressOf F1_Deactivate
        AddHandler F1.Resize, AddressOf F1_resize
        AddHandler F1.KeyUp, AddressOf F1_KeyUp

        AddHandler Grid1.MouseDown, AddressOf Grid1_MouseDown2
        AddHandler Grid1.MouseUp, AddressOf Grid1_MouseUp
        AddHandler Grid1.KeyPress, AddressOf Grid1_KeyPress
        AddHandler Grid1.KeyDown, AddressOf Grid1_Keydown
        AddHandler Grid1.KeyUp, AddressOf Grid1_KeyUp
        AddHandler Grid1.Scroll, AddressOf Grid1_Scroll
        '
        AddHandler BoutonF1_1.Click, AddressOf BoutonF1_1_Click
        AddHandler BoutonF1_2.Click, AddressOf BoutonF1_2_Click
        '
        AddHandler BoutonFermer.Click, AddressOf BoutonFermer_Click
        '
        AddHandler ActionsPianoR.Click, AddressOf ActionsPianoR_Click
        '
        AddHandler BRadio1.MouseDown, AddressOf BRadio1_MouseDown
        AddHandler BRadio2.MouseDown, AddressOf BRadio2_MouseDown
        AddHandler BRadio3.MouseDown, AddressOf BRadio3_MouseDown
        '
        AddHandler CheckAcc.CheckedChanged, AddressOf Checkacc_CheckedChanged
        AddHandler CheckDrum.CheckedChanged, AddressOf CheckDrum_CheckedChanged
        AddHandler Check12_8.CheckedChanged, AddressOf Check12_8_CheckedChanged
        AddHandler AffCtrls.CheckedChanged, AddressOf AffCtrls_CheckedChanged
        AddHandler ZoomMoins.CheckedChanged, AddressOf ZoomMoins_CheckedChanged
        '
        AddHandler DebMacroSel.ValueChanged, AddressOf DebMacroSel_ValueChange
        AddHandler TermeMacroSel.ValueChanged, AddressOf TermeMacroSel_ValueChange
        AddHandler BMacroSel.Click, AddressOf BMacroSel_Click
        AddHandler OuvrirCalques.Click, AddressOf OuvrirCalques_Click
        '
        AddHandler DockButton.Click, AddressOf Dockbutton_MouseClick
        AddHandler Opacité.ValueChanged, AddressOf Opacité_ValueChange
        '
        AddHandler Destination.KeyDown, AddressOf Destination_Keydown
        AddHandler Destination.TextChanged, AddressOf Destination_TextChanged
        '
        AddHandler ListDynF1.SelectedIndexChanged, AddressOf ListdynF1_SelectedIndexChanged
        '
        AddHandler Couper.Click, AddressOf Couper_Click
        AddHandler Copier.Click, AddressOf Copier_Click
        AddHandler Coller.Click, AddressOf Coller_Click
        AddHandler Annuler.Click, AddressOf Annuler_Click
        AddHandler Supprimer.Click, AddressOf Supprimer_Click
        AddHandler Quitter.Click, AddressOf Quitter_Click
        '
        AddHandler ListPRG.SelectedIndexChanged, AddressOf ListPRG_SelectedIndexChanged
        '
        AddHandler ListMod.SelectedIndexChanged, AddressOf ListMod_SelectedIndexChanged
        '
        AddHandler TimerStop.Tick, AddressOf TimerStop_Tick

        '
        AddHandler ListTypNote.KeyDown, AddressOf ListTypNote_KeyDown
        AddHandler ListDynF1.KeyDown, AddressOf ListDynF1_KeyDown
        AddHandler ListMod.KeyDown, AddressOf ListMod_KeyDown
        AddHandler ListPRG.KeyDown, AddressOf ListPRG_KeyDown
        AddHandler ListTonNotes.KeyDown, AddressOf ListTonNotes_KeyDown


        Me.Numérateur = Det_Numérateur(Me.Métrique)
        Me.Dénominateur = Det_Dénominateur(Me.Métrique)
        Me.DivisionMes = Det_DivisionMes()
        '
        Me.Canal = Canal
        Me.CoulBarOut = CoulBarout
        Me.F1.KeyPreview = True ' pour réception des tocuhes F4 et F5 pour Recalcul
        '
        BoutonFermer.Visible = False
        '
        For i = 0 To Module1.nbCalques - 1
            ChoixCalquesLocal(i) = False
        Next
        '
        ChoixCalquesLocal(1) = True ' à l'init le calque MIDI gammes est activé

        ListPédale.Add("C")
        ListPédale.Add("C#")
        ListPédale.Add("D")
        ListPédale.Add("D#")
        ListPédale.Add("E")
        ListPédale.Add("F")
        ListPédale.Add("F#")
        ListPédale.Add("G")
        ListPédale.Add("G#")
        ListPédale.Add("A")
        ListPédale.Add("A#")
        ListPédale.Add("B")

    End Sub
    Private Sub ListTypNote_KeyDown(Sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        e.SuppressKeyPress = True
    End Sub
    Private Sub ListDynF1_KeyDown(Sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        e.SuppressKeyPress = True
    End Sub
    Private Sub ListMod_KeyDown(Sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        e.SuppressKeyPress = True
    End Sub
    Private Sub ListPRG_KeyDown(Sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        e.SuppressKeyPress = True
    End Sub
    Private Sub ListTonNotes_KeyDown(Sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        e.SuppressKeyPress = True
    End Sub





    ' *******************************************************
    '  METHODES PUBLIQUES
    ' *******************************************************
    Public Sub F1_Refresh()
        '
        Grid1.AutoRedraw = False
        '
        Refresh_G1_Acc()
        Refresh_G1_Gam()
        Refresh_G1_Marq()
        '
        Clear_graphique()
        Graphique_Gammes()
        '
        If CheckAcc.Checked Then ' affichage des notes de l'accord dans les calques MIDI
            Graphique_Accords()
        End If
        '
        Graphique_Divisions()
        '
        If CheckTonique.Checked Then ' affichage de la Tonique d'un Mode
            Graphique_Tonique()
        End If
        '
        If Check12_8.Checked Then
            Graphique_12_8()
        End If
        '
        If AffCtrls.Checked Then
            Show_Ctrls()
        Else
            Hide_Ctrls()
        End If
        Grid1.AutoRedraw = True
        Grid1.Refresh()
        '
        Me.Chargé = True
        '
    End Sub
    Public Sub F1_Refresh2()
        '
        Grid1.AutoRedraw = False
        '
        'Refresh_G1_Acc()
        'Refresh_G1_Gam()
        'Refresh_G1_Transf()
        'Refresh_G1_Marq()
        '
        Clear_graphique()
        Graphique_Gammes()
        '
        If CheckAcc.Checked Then
            Graphique_Accords()
        End If
        '
        Graphique_Divisions()
        '
        If CheckDrum.Checked Then
            Graphique_Drums()
        End If
        '
        If Check12_8.Checked Then
            Graphique_12_8()
        End If
        Grid1.AutoRedraw = True
        Grid1.Refresh()
        '
        Me.Chargé = True
        '
    End Sub
    Public Sub Construction_MenuContext()

        ' Menu Couper
        If Langue = "fr" Then
            itemVélo.Text = "Vélocités aléatoires"
        Else
            itemVélo.Text = "Random velocities"
        End If
        MenuContext2.Items.Add(itemVélo)
        '
        AddHandler itemVélo.Click, AddressOf itemVélo_Click
    End Sub
    Sub itemVélo_Click()
        'Posx = MenuContext2.Location.X
        'Posy = MenuContext2.Location.Y
        '
        Posx = Cursor.Position.X
        Posy = Cursor.Position.Y
        If Not VéloAléat_Chargé Then

            Dim ii As Decimal = Convert.ToInt16(ListDynF1.Text) - 15
            Dim jj As Decimal = Convert.ToInt16(ListDynF1.Text) + 15

            ii = Convert.ToInt16(ListDynF1.Text) - 15
            jj = Convert.ToInt16(ListDynF1.Text) + 15

            If ii <= 0 Then
                ii = 1
                jj = 30
            Else
                If jj > 128 Then
                    ii = 98
                    jj = 128
                End If
            End If

            ValMn = ii
            ValMx = jj
        Else
            ValMn = VéloAleat.ValeurMin.Value
            ValMx = VéloAleat.ValeurMax.Value
        End If
        '
        VéloAleat.ShowDialog()
        '
        If ValMn <> -1 Then ' ValMn = -1 si sortie par Annuler
            ValAlea_Local(VéloAleat.ValeurMin.Value, VéloAleat.ValeurMax.Value)
        End If
    End Sub
    Sub ValAlea_Local(VMin As Integer, Vmax As Integer)
        Dim rand As New Random() ' Création une instance de la classe Random
        Dim Fr As Integer = Grid1.Selection.FirstRow
        Dim Lr As Integer = Grid1.Selection.LastRow
        Dim Fc As Integer = Grid1.Selection.FirstCol
        Dim Lc As Integer = Grid1.Selection.LastCol
        '
        ' Préparation pour CTRL+Z
        ' ***********************
        Dim aa As New Ann
        'ListAnnulation.Clear()
        'PointAnn = -1
        '
        ListAnnulation.Add(aa)
        PointAnn = PointAnn + 1
        ListAnnulation.Item(PointAnn).Action = ActionEnum.Restituer

        For i = Fr To Lr
            For j = Fc To Lc
                If IsNumeric(Grid1.Cell(i, j).Text) Then
                    m = 0
                    Dim oo As New Ann.SauvAnnuler
                    k = j
                    oo.Vélo = Trim(Grid1.Cell(i, k).Text)
                    oo.Ligne = i
                    oo.Colonne = k
                    Grid1.Cell(i, k).Text = rand.Next(VMin, Vmax)
                    ListAnnulation.Item(PointAnn).ListAnnuler.Add(oo) ' sauvegarde pour CTRL+Z
                    ListAnnulation.Item(PointAnn).Tinfo = TypeInfo.Vélocités ' 
                End If
            Next
        Next

    End Sub

    '
    ''' <summary>
    ''' Construction d'une chaîne de caractères contenant toutes les notes et CTRLs d'un piano roll
    ''' </summary>
    ''' <param name="Répéter"> : boolean-case à cocher oui/non</param>
    ''' <param name="Boucle"> : nombre de réptition</param>
    ''' <param name="Form1_Début"> : Début de la lecture en nombre de mesures: locateur Début de la barre de transport</param>
    ''' <param name="Form1_Fin">: Fin de la lecture en nombres de mesures locateur Terme de la barre de transport</param>
    ''' <param name="NumDerAcc"> : n° de mesure du dernier accord dans le projet</param>
    ''' <returns>Retourne une chaine de caractéres avec éléments séparés par des blancs </returns>
    Function Contruction_ListeNotes(Répéter As Boolean, Boucle As Integer, Form1_Début As Integer, Form1_Fin As Integer, NumDerAcc As Integer) As String
        Dim i, j, k, p As Integer
        Dim nbColonnes As Integer
        Dim DebPart As Integer
        Dim FinPart As Integer
        Dim a As String = ""
        Dim n As String
        Dim pédale As String

        Dim NumPiste As String
        Dim Canal As String
        Dim ValeurCtrl As String
        Dim ValeurCtrlPréced As String = "0"
        Dim TValPréded = {"0", "0", "0", "0", "0", "0", "0"} ' Valeurs précédentes de respectivement Expr, Mode, PAN, cc50, cc51, cc52, cc53
        Dim ValeurNote As String
        Dim Début As String
        Dim Durée As String
        Dim Vélocité As String
        Dim _boucle As Integer = 0
        Dim LongueurPart As Integer = (Form1_Fin - Form1_Début) + 1 ' en nombre de mesures
        Dim N_Mes As Integer
        Dim nbColonnesMes As Integer
        Dim DebutNote As Integer
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
        '
        ' N° ONGLET - INDEX LISTE PIANO ROLL - N0 CANAUX MIDI
        '                       HYPERV | 7-PIANOROLL| 8-PIANOROLL| 9-PIANOROLL|10-BATTERIE  | 11-PIANOROLL| 12-PIANOROLL| 13-PIANOROLL| 14-PIANOROLL
        '----------------------------------------------------------------------------------------------------------------------------------------- 
        ' N° ONGLET         |      0   |     1      |       2    |       3    |       4     |      5      |      6      |       7     |      8
        '-----------------------------------------------------------------------------------------------------------------------------------------
        ' INDEX LIST PIANOR |     --   |     0      |       1    |       2    |       --    |      3      |      4      |       5     |      6
        '-----------------------------------------------------------------------------------------------------------------------------------------
        ' N° CANAUX MIDI    |    0-5   |     6      |       7    |       8    |       9     |      10     |      11     |       12    |      13
        '-----------------------------------------------------------------------------------------------------------------------------------------

        If Répéter Then _boucle = Boucle
        'nbColonnesMes = NbDivInduit(Me.Dénominateur) * Me.Numérateur ' en 4/4 le résultat donne 16
        nbColonnesMes = Me.DivisionMes
        ' Rétablissement de Form1_Fin en fonction du nomre de répétitions de l'acccord qu'il pointe
        If Form1_Fin = NumDerAcc Then
            Form1_Fin = Form1_Fin + Val(Form1.Grid2.Cell(3, Form1_Fin).Text) - 1
        End If
        ' Détermination du DEPART et de la FIN en N° de colonnes de la séquence à jouer(sert à la boucle exécuter plus loins)
        Départ = ((nbColonnesMes) * (Form1_Début - 1)) + 1 ' DEPART en N° de colonnes
        nbColonnes = (nbColonnesMes) * Form1_Fin           ' FIN en N° de colonnes

        '
        ' Détermination du DEPART et de la FIN en N° de mesures de la séquence à jouer : sert à calculer LongueurPart 
        DebPart = Form1_Début   ' DEPART en N° de mesures
        FinPart = Form1_Fin     ' FIN en N° de mesures
        LongueurPart = ((FinPart - DebPart) + 1) * nbColonnesMes ' Longueur de la séquence à jouer en nombre de colonnes
        '
        n = Convert.ToString(Me.Canal) ' 
        '
        ' DEPART : position 0
        ' *******************
        ' Mise à jour des controleurs au départ du jeu
        ' ********************************************
        If Form1.Mix.AutorisVol.Checked = True Then ' envoi volumes seulement si table de mixage activée
            Volume = Convert.ToString(Form1.Récup_Volume(n))
            If Form1.Récup_VolumeActif(n) = False Then Volume = "0" ' gestion du système de Mute de la table de mixage
            a = "CTRL" + " " + Trim(n) + " " + Trim(n) + " " + "0" + " " + Convert.ToString(CVolume) + " " + Volume + "-"
        End If
        '
        Dim PRG As Integer = ListPRG.SelectedIndex - 1
        If PRG <> -1 Then
            a = a + "PRG" + " " + Trim(n) + " " + Trim(n) + " " + "0" + " " + Convert.ToString(PRG) + "-" ' programme change
        End If
        'a = a + "CTRL" + " " + Trim(n) + " " + Trim(n) + " " + "0" + " " + Convert.ToString(CPAN) + " " + Trim(Det_ValPan) + "-"
        'a = a + "CTRL" + " " + Trim(n) + " " + Trim(n) + " " + "0" + " " + Convert.ToString(CMod) + " " + "0" + "-"
        ''a = a + "CTRL" + " " + Trim(n) + " " + Trim(n) + " " + "0" + " " + Convert.ToString(CExpr) + " " + Volume + "-" ' supprimé car inutile, c'est le volume qui prend la main quand l'expression est absente  
        'a = a + "CTRL" + " " + Trim(n) + " " + Trim(n) + " " + "0" + " " + Convert.ToString(CPed) + " " + "0" + "-"
        'a = a + "CTRL" + " " + Trim(n) + " " + Trim(n) + " " + "0" + " " + Convert.ToString(CC50) + " " + "0" + "-"
        'a = a + "CTRL" + " " + Trim(n) + " " + Trim(n) + " " + "0" + " " + Convert.ToString(CC51) + " " + "0" + "-"
        'a = a + "CTRL" + " " + Trim(n) + " " + Trim(n) + " " + "0" + " " + Convert.ToString(CC52) + " " + "0" + "-"
        'a = a + "CTRL" + " " + Trim(n) + " " + Trim(n) + " " + "0" + " " + Convert.ToString(CC53) + " " + "0" + "-"
        '
        ' Mise à jour des Notes
        ' *********************
        Me.PPresNotes = False
        For k = 0 To _boucle
            For j = Départ To nbColonnes
                N_Mes = Fix((j - 1) / nbColonnesMes) + 1 ' N_Mes : N° de mesure
                If (N_Mes >= Form1_Début) And (N_Mes <= Form1_Fin) Then ' 
                    For i = Grid1.FixedRows - 1 To Grid1.Rows - 1 ' 
                        If IsNumeric(Grid1.Cell(i, j).Text) Then
                            Me.PPresNotes = True
                            ' analyse de la grille des notes
                            ' ******************************
                            If i <> Grid1.FixedRows - 1 Then    ' Notes
                                NumPiste = Trim(n)
                                Canal = Trim(n)
                                ValeurNote = Convert.ToString(ValNoteCubase.IndexOf(Grid1.Cell(i, 0).Text))
                                DebutNote = (j - 1) - ((Form1_Début - 1) * nbColonnesMes)
                                Début = Convert.ToString(DebutNote + (k * LongueurPart)) ' longueurpart n'intervient que pour une boucle
                                Durée = Det_LongueurNote2(i, j, nbColonnes) ' longueur de la note en nombre de double croches
                                Vélocité = Trim(Grid1.Cell(i, j).Text)
                                a = a + ("Note" + " " + Trim(NumPiste) + " " + Trim(Canal) + " " + Trim(ValeurNote) + " " + Trim(Début) + " " _
                            + Trim(Durée) + " " + Trim(Vélocité) + "-")
                            End If
                        Else
                            ' LIGNE DE PEDALE
                            pédale = Trim(Trim(Grid1.Cell(i, j).Text))
                            If j = Départ Then
                                pédale = Det_PédaleAvant(j)
                            End If
                            If i = (Grid1.FixedRows - 1) And (Trim(Grid1.Cell(i, j).Text) <> "" Or Trim(pédale) <> "") Then   ' ligne de Pedal 
                                NumPiste = Trim(n)
                                Canal = Trim(n)
                                If Trim(pédale) = "P+" Then
                                    ValeurCtrl = 127
                                Else
                                    ValeurCtrl = 0
                                End If

                                DebutNote = (j - 1) - ((Form1_Début - 1) * nbColonnesMes)
                                Début = Convert.ToString(DebutNote + (k * LongueurPart)) ' longueurpart n'intervient que pour une boucle
                                a = a + "CTRL" + " " + Trim(NumPiste) + " " + Trim(Canal) + " " + Trim(Début) + " " + Convert.ToString(CPed) + " " + Trim(ValeurCtrl) + "-"
                            End If
                        End If
                    Next i
                    ' analyse des courbes de controleurs
                    ' **********************************
                    For p = 0 To nbCourbes - 1
                        If CCActif.Item(p).Checked Then
                            ValeurCtrl = ValCtrl(p, j)
                            If ValeurCtrl <> TValPréded(p) And Trim(ValeurCtrl) <> "-1" Then ' pas besoin de répéter plusieurs fois la même valeur (surtout le 0)
                                ' on n'écrit l'expression que si le check du CTRL est true
                                NumPiste = Trim(n)
                                Canal = Trim(n)
                                DebutNote = (j - 1) - ((Form1_Début - 1) * nbColonnesMes)
                                Début = Convert.ToString(DebutNote + (k * LongueurPart))
                                a = a + "CTRL" + " " + Trim(NumPiste) + " " + Trim(Canal) + " " + Trim(Début) + " " + Convert.ToString(Det_CTRL(p)) + " " + Trim(ValeurCtrl) + "-"
                                TValPréded(p) = ValeurCtrl
                            End If
                        End If
                    Next p
                End If
            Next j
        Next k
        If Trim(a) <> "" Then a = Trim(Microsoft.VisualBasic.Left(a, a.Length - 1))

        Return a
    End Function
    Private Function Det_PédaleAvant(coldeb As Integer) As String
        Dim j As Integer
        Dim i As Integer = Grid1.FixedRows - 1
        Dim r As String = ""

        For j = coldeb To 1 Step -1
            If Trim(Grid1.Cell(i, j).Text) <> "" Then
                r = Trim(Grid1.Cell(i, j).Text)
                Exit For
            End If
        Next
        Return r
    End Function
    ''' <summary>
    ''' ValCtrl : retourne la valeur du CTRL dans la courbe N_ Courbe pour la colonne 'colonne'. (valeur= 0 à 127)
    ''' </summary>
    ''' <param name="N_Courbe"></param> N° de l'onglet portant la courbe (caractérise le CTRL : Coure=0 --> Ctrl= Mod, Courbe = 1 --> Ctrl = Expression etc.
    ''' <param name="colonne"></param> Colonne de la grille des courbes : colonne = postion de l'évènement en double croches à partir du début
    ''' <returns></returns>

    Private Function ValCtrl(N_Courbe As Integer, colonne As Integer) As String
        Dim ligne As Integer
        Dim j As Integer = colonne
        Dim k As Integer = 0

        For ligne = 1 To GridCourbes.Item(N_Courbe).Rows - 1
            If ligne = GridCourbes.Item(N_Courbe).Rows - 1 Then ' pas de valeur trouvée à la fin
                k = CTRL_ValAmont(N_Courbe, colonne)
                Exit For
            Else
                If GridCourbes.Item(N_Courbe).Cell(ligne, j).BackColor <> Color.White Then ' valeur trouvée
                    k = ((GridCourbes.Item(N_Courbe).Rows - 1 - ligne) * 2) - 1
                    Exit For
                End If
            End If
        Next
        k = k - 1
        Return k.ToString
    End Function
    Function CTRL_ValAmont(N_Courbe As Integer, colonne As Integer) As Integer
        Dim j As Integer
        Dim ligne As Integer
        Dim k As Integer = 0
        Dim flag As Boolean = False


        For j = colonne To 1 Step -1
            For ligne = 1 To GridCourbes.Item(N_Courbe).Rows - 1
                If GridCourbes.Item(N_Courbe).Cell(ligne, j).BackColor <> Color.White Then ' valeur trouvée
                    k = ((GridCourbes.Item(N_Courbe).Rows - 1 - ligne) * 2) - 1 ' 1ere valeur trouvée dans un colonne précédent du CTRL
                    flag = True
                    Exit For
                End If
            Next
            If flag Then Exit For
        Next
        Return k
    End Function



    ''' <summary>
    ''' ValCtrl2_Ligne : retourne la 1ere ligne différent de Blanc dans la colonne passée en paramètre. Cela donne la valeur do CC pour la conne indiquée.)
    ''' </summary>
    ''' <param name="N_Courbe">N° de la courbes dans le PianoRoll</param>
    ''' <param name="colonne">Colonne de la courbe à traiter</param>
    ''' <returns></returns>
    Private Function ValCtrl2_Ligne(N_Courbe As Integer, colonne As Integer) As String
        Dim Ligne As Integer
        Dim j As Integer = colonne
        Dim a As String = (GridCourbes.Item(N_Courbe).Rows - 1).ToString
        '
        For Ligne = 1 To GridCourbes.Item(N_Courbe).Rows - 1
            If GridCourbes.Item(N_Courbe).Cell(Ligne, j).BackColor <> Color.White Then ' on recherche ici la 1erre cellule dont la couleur n'est blanche
                'v = ((GridCourbes.Item(N_Courbe).Rows - 1 - Ligne) * 2) + 1 ' (65-1 - ligne)*2
                a = Ligne.ToString ' ligne est le N° de ligne
                Exit For
            End If
        Next

        Return a
    End Function
    Function Det_CTRL(n_Courbe As Integer) As Byte
        Dim k As Integer = 0
        Select Case n_Courbe
            Case 0 ' Mod
                k = CExpr
            Case 1 ' expression
                k = CMod
            Case 2 ' PAN
                k = CPAN
            Case 3 ' Libre 1
                k = CC50
            Case 4 ' Libre 2
                k = CC51
            Case 5 ' Libre 2
                k = CC52
            Case 6 ' Libre 2
                k = CC53
        End Select
        Return k
    End Function
    Function Det_CouleurCTRL(NCourbe As Integer) As Color
        Dim coul As New Color
        Select Case NCourbe
            Case 0
                coul = CoulExp
            Case 1
                coul = CoulMod
            Case 2
                coul = CoulPan
            Case 3
                coul = CoulCC50
            Case 4
                coul = CoulCC51
            Case 5
                coul = CoulCC52
            Case 6
                coul = CoulCC53
            Case 7, 8
                coul = CoulAss
        End Select
        Return coul
    End Function
    Function NumPianoR(Canal As Byte) As String
        Dim a As String
        If Canal <= 8 Then
            a = Convert.ToString(Canal - 5)
        Else
            a = Convert.ToString(Canal - 6)
        End If
        Return a
    End Function

    Public Function Enregistrer_ParamMélo() As String
        Dim n_PianoRoll As Integer = Me.Canal - nb_PianoRoll
        Dim a As String = "PianoRoll," + Convert.ToString((NumPianoR(Me.Canal))) + ",ParamMélo,"
        a = a + "ActCanal&" + Convert.ToString(CheckMute.Checked) + ","
        a = a + "TypNote&" + Convert.ToString(ListTypNote.SelectedIndex) + ","
        a = a + "ModeNote&" + Convert.ToString(ListTonNotes.SelectedIndex) + ","
        a = a + "Dyn&" + Convert.ToString(ListDynF1.SelectedIndex) + ","
        a = a + "PRG&" + Convert.ToString(ListPRG.SelectedIndex) + ","
        a = a + "Acc&" + Convert.ToString(CheckAcc.Checked) + ","
        a = a + "Mode&" + Convert.ToString(CheckTonique.Checked) + ","
        a = a + "Drum&" + Convert.ToString(CheckDrum.Checked) + ","
        a = a + "Radio1&" + Convert.ToString(BRadio1.Checked) + ","
        a = a + "Radio2&" + Convert.ToString(BRadio2.Checked) + ","
        a = a + "Radio3&" + Convert.ToString(BRadio3.Checked) + ","
        a = a + "NomduSon&" + Convert.ToString(NomduSon.Text)
        Return Trim(a)
    End Function
    Public Function Enregistrer_CalquesMIDI() As String
        Dim a As String = "PianoRoll," + Convert.ToString((NumPianoR(Me.Canal))) + ",CalquesMIDI,"
        a = a + Convert.ToString(ChoixCalquesLocal(0)) + ","
        a = a + Convert.ToString(ChoixCalquesLocal(1)) + ","
        a = a + Convert.ToString(ChoixCalquesLocal(2)) + ","
        a = a + Convert.ToString(ChoixCalquesLocal(3)) + ","
        a = a + Convert.ToString(ChoixCalquesLocal(4))
        Return Trim(a)
    End Function
    Public Function Enregistrer_ParamCalquesMIDI() As String
        Dim a As String = "PianoRoll," + Me.Canal.ToString + ",ParamCalquesMIDI," + PédaleLocale.ToString
        Return Trim(a)
    End Function
    Public Function Enregistrer_NotesMélo() As String
        Dim i, j As Integer
        Dim nbColonnesMes As Integer
        Dim a As String = "PianoRoll," + Convert.ToString((NumPianoR(Me.Canal))) + ",NotesMélo,"
        '
        Dim Ligne As String
        Dim Colonne As String
        Dim ValeurCtrl As String
        Dim ValeurNote As String
        Dim Durée As String
        Dim Vélocité As String
        '
        nbColonnesMes = NbDivInduit(Me.Dénominateur) * Me.Numérateur ' en 4/4 le résultat donne 16, en 7/8
        nbColonnes = (nbColonnesMes) * Me.NbAccords ' 
        '
        For j = 1 To nbColonnes
            For i = Grid1.FixedRows - 1 To Grid1.Rows - 1
                If IsNumeric(Grid1.Cell(i, j).Text) Then
                    If i <> Grid1.FixedRows - 1 Then ' notes
                        Ligne = Convert.ToString(i)
                        Colonne = Convert.ToString(j)
                        ValeurNote = Convert.ToString(ValNoteCubase.IndexOf(Grid1.Cell(i, 0).Text))
                        Durée = Det_LongueurNote(i, j) ' longueur de la note en nombre de double croches
                        Vélocité = Trim(Grid1.Cell(i, j).Text)
                        a = a + Ligne + "-" + Colonne + "-" + ValeurNote + "-" + Durée + "-" + Vélocité + ","
                    Else ' CTRL expression
                        Ligne = Convert.ToString(i) ' le N° de ligne = 5 permet dans 'Ouvrir' de reconnaitre le CTRl Modulation = 1 
                        Colonne = Convert.ToString(j)
                        ValeurCtrl = Trim(Grid1.Cell(i, j).Text)
                        a = a + Ligne + "-" + Colonne + "-" + ValeurCtrl + ","
                    End If

                End If
            Next
        Next
        If Trim(a) <> "NotesMélo" Then a = Trim(Microsoft.VisualBasic.Left(a, a.Length - 1))
        Return Trim(a)
    End Function
    Public Function Enregistrer_Ctrls(N_Ctrl As Integer) As String
        Dim ind As Integer = N_Ctrl
        Dim a As String = ""
        Dim Ligne As String
        Dim Colonne As Integer

        a = "PianoRoll," + Convert.ToString((NumPianoR(Me.Canal))) + ",Control," + Convert.ToString(ind) + ","
        For Colonne = 1 To GridCourbes(ind).Cols - 1
            Ligne = ValCtrl2_Ligne(ind, Colonne)
            If Ligne <> (GridCourbes.Item(ind).Rows - 1).ToString Then ' la dernière ligne indique la valeur 0, elle reste en blanc
                a = a + Ligne + " " + Colonne.ToString + ","
            End If
        Next Colonne

        a = Trim(Microsoft.VisualBasic.Left(a, a.Length - 1))
        Return a
    End Function
    Public Function Enregistrer_CtrlPédale() As String
        Dim a As String = ""
        Dim Ligne As String = Grid1.FixedRows - 1
        Dim Colonne As Integer
        '
        a = "PianoRoll," + Convert.ToString((NumPianoR(Me.Canal))) + ",Pédale,"
        For Colonne = 1 To Grid1.Cols - 1
            If Trim(Grid1.Cell(Ligne, Colonne).Text) <> "" Then
                a = a + Trim(Grid1.Cell(Ligne, Colonne).Text) + " " + Ligne.ToString + " " + Colonne.ToString + ","
            End If
        Next Colonne
        a = Trim(Microsoft.VisualBasic.Left(a, a.Length - 1))
        Return a
        '
    End Function
    Public Function Enregistrer_ControlSys() As String
        Dim i As Integer
        Dim a As String = ""

        a = "PianoRoll," + Convert.ToString((NumPianoR(Me.Canal))) + ",ControlSys,"
        For i = 0 To nbCourbes - 1
            a = a + Convert.ToString(CCActif.Item(i).Checked) + ","
        Next i
        a = a + Convert.ToString(AffCtrls.Checked)
        Return a
    End Function
    Public Function Enregistrer_Assist1CTRP() As String
        Dim i, j As Integer
        Dim Marq As String = "Off"

        a = "PianoRoll," + Convert.ToString((NumPianoR(Me.Canal))) + ",Assist1CTRP,"
        For i = 1 To GridAssist1.Rows - 1
            If Trim(GridAssist1.Cell(i, Col_CF).Text) <> "" Then
                ' recherche de la note marquée
                j = 0
                Do
                    j += 1
                    If GridAssist1.Cell(i, j).BackColor = CouleurMarq Then Marq = i.ToString + "-" + j.ToString
                Loop Until j = Maxcol Or GridAssist1.Cell(i, j).BackColor = CouleurMarq

                ' N° ligne                ' Note CF
                a = a + i.ToString + " " + Trim(GridAssist1.Cell(i, Col_CF).Text) + " " + Trim(Marq) + ","
            End If
        Next i
        a = Trim(Microsoft.VisualBasic.Left(a, a.Length - 1))
        Return a
    End Function
    Public Sub Charger_ControlSys(LesCtrlSys As String)
        Dim i As Integer
        Dim tbl1() As String
        tbl1 = LesCtrlSys.Split(",")

        For i = 3 To tbl1.Count - 2
            CCActif.Item(i - 3).Checked = tbl1(i)
        Next
        AffCtrls.Checked = tbl1(tbl1.Count - 1)
    End Sub
    Public Sub Charger_CalquesMIDI(LesCalques As String)
        Dim i As Integer
        Dim tbl1() As String

        tbl1 = LesCalques.Split(",")
        For i = 3 To tbl1.Count - 1
            ChoixCalquesLocal(i - 3) = Convert.ToBoolean(tbl1(i))
            PassChoixCalques(i - 3) = Convert.ToBoolean(tbl1(i))
        Next i
    End Sub
    Public Sub Charger_ParamCalquesMIDI(ParamC As String)
        Dim tbl1() As String
        tbl1 = ParamC.Split(",")
        '
        PédaleLocale = Convert.ToInt16(tbl1(3))
    End Sub
    Public Sub Charger_CtrlPédale(Pédales As String)
        Dim k As Integer
        Dim tbl1() As String
        Dim tbl2() As String
        tbl1 = Pédales.Split(",")

        If tbl1.Count > 3 Then
            For k = 3 To tbl1.Count - 1
                tbl2 = tbl1(k).Split
                Grid1.Cell(Convert.ToInt16(tbl2(1)), Convert.ToInt16(tbl2(2))).Text = Trim(tbl2(0))
            Next k
        End If
    End Sub
    Public Sub Charger_Assist1CTRP(Assist1 As String)
        Dim i As Integer
        Dim tbl1() As String
        Dim tbl2() As String
        Dim tbl3() As String
        Dim ligne As Integer
        Dim NoteCF As String
        Dim Marq As String = "Off"
        Dim r, c As Integer

        tbl1 = Assist1.Split(",")
        '

        For i = 3 To tbl1.Count - 1
            ' note
            tbl2 = tbl1(i).Split()
            ligne = Val(tbl2(0))
            NoteCF = Trim(tbl2(1))
            Marq = Trim(tbl2(2))
            ' maj note dans colonne CF
            GridAssist1.Cell(ligne, Col_CF).Text = NoteCF
            Charger_Assist1_Calcul(ligne, NoteCF) ' calcul des valeurs de consonance /dissonance

            ' Marquage
            If Marq <> "Off" Then
                tbl3 = Marq.Split("-")
                r = Val(tbl3(0))
                c = Val(tbl3(1))
                GridAssist1.Cell(r, c).BackColor = CouleurMarq
            End If
        Next

    End Sub
    Sub Charger_Assist1_Calcul(r As Integer, NoteCF As String)
        Dim i As Integer

        With GridAssist1
            i = -1
            Do
                i += 1
            Loop Until (i = Chromad.Count - 1) Or (NoteCF = Chromad.Item(i))
            '
            If i < Chromad.Count - 1 Then
                ' Voix haute
                .Cell(r, Col_CF + 12).Text = Chromad(i + Octave)
                .Cell(r, Col_CF + 11).Text = Chromad(i + Sept)
                .Cell(r, Col_CF + 10).Text = Chromad(i + Septb)
                .Cell(r, Col_CF + 9).Text = Chromad(i + SixteMaj)
                .Cell(r, Col_CF + 8).Text = Chromad(i + SixteMin)
                .Cell(r, Col_CF + 7).Text = Chromad(i + Quinte)
                .Cell(r, Col_CF + 6).Text = Chromad(i + Quinteb)
                .Cell(r, Col_CF + 5).Text = Chromad(i + Quarte)
                .Cell(r, Col_CF + 4).Text = Chromad(i + TierceMaj)
                .Cell(r, Col_CF + 3).Text = Chromad(i + TierceMin)
                .Cell(r, Col_CF + 2).Text = Chromad(i + Neuf)
                .Cell(r, Col_CF + 1).Text = Chromad(i + Neufb)
                '.Cell(r, Col_CF).Text = Chromad(i + Octave)
                ' Voix Basse
                '.Cell(r, 13).Text = Chromad(i + EstOctave)
                .Cell(r, 12).Text = Chromad(i + EstNeufb)
                .Cell(r, 11).Text = Chromad(i + EstNeuf)
                .Cell(r, 10).Text = Chromad(i + EstTierceMin)
                .Cell(r, 9).Text = Chromad(i + EstTierceMaj)
                .Cell(r, 8).Text = Chromad(i + EstQuarte)
                .Cell(r, 7).Text = Chromad(i + EstQuinteb)
                .Cell(r, 6).Text = Chromad(i + EstQuinte)
                .Cell(r, 5).Text = Chromad(i + EstSixteMin)
                .Cell(r, 4).Text = Chromad(i + EstSixteMaj)
                .Cell(r, 3).Text = Chromad(i + EstSeptb)
                .Cell(r, 2).Text = Chromad(i + EstSept)
                .Cell(r, 1).Text = Chromad(i + EstOctave)
            End If
        End With
    End Sub

    Public Sub Init_ControlSys()
        Dim i As Integer
        For i = 0 To nbCourbes - 1
            CCActif.Item(i).Checked = False
        Next i
        AffCtrls.Checked = False
    End Sub
    Public Sub Clear_Courbes()
        Dim i As Integer
        For i = 0 To nbCourbes - 1
            GridCourbes.Item(i).AutoRedraw = False
            GridCourbes.Item(i).Range(1, 1, GridCourbes.Item(i).Rows - 1, GridCourbes.Item(i).Cols - 1).ClearFormat()
            GridCourbes.Item(i).AutoRedraw = True
            GridCourbes.Item(i).Refresh()
        Next
    End Sub
    Public Sub Charger_Param(LisiParam As String)
        Dim i As Integer
        Dim tbl1() As String
        Dim tbl2() As String

        tbl1 = LisiParam.Split(",")

        For i = 3 To UBound(tbl1)
            tbl2 = tbl1(i).Split("&")
            Select Case Trim(tbl2(0))
                Case "ActCanal"
                    CheckMute.Checked = Convert.ToBoolean(tbl2(1))
                Case "TypNote"
                    ListTypNote.SelectedIndex = Convert.ToInt16(tbl2(1))
                Case "Dyn"
                    ListDynF1.SelectedIndex = Convert.ToInt16(tbl2(1))
                Case "ModeNote"
                    ListTonNotes.SelectedIndex = Convert.ToInt16(tbl2(1))
                Case "PRG"
                    ListPRG.SelectedIndex = Convert.ToInt16(tbl2(1))
                Case "Acc"
                    CheckAcc.Checked = Convert.ToBoolean(tbl2(1))
                Case "Mode"
                    CheckTonique.Checked = Convert.ToBoolean(tbl2(1))
                Case "Drum"
                    CheckDrum.Checked = Convert.ToBoolean(tbl2(1))
                Case "Radio1"
                    BRadio1.Checked = Convert.ToBoolean(tbl2(1))
                Case "Radio2"
                    BRadio2.Checked = Convert.ToBoolean(tbl2(1))
                Case "Radio3"
                    BRadio3.Checked = Convert.ToBoolean(tbl2(1))
                Case "NomduSon"
                    NomduSon.Text = Trim(tbl2(1)) 'sur écriture de NomduSon dans PIANOROLL l'évènement NomduSon_TextChanged est appelé et met à jour le NomduSon dans la table de mixage
            End Select
        Next i

    End Sub
    Public Sub Charger_Notes(LesNotes As String)
        Dim i, j, k, m As Integer
        Dim tbl1() As String
        Dim tbl2() As String
        Dim Durée As Integer
        Dim Vélocité As String

        If Trim(LesNotes) <> "" Then
            tbl1 = LesNotes.Split(",")
            '
            Grid1.Visible = False
            Grid1.AutoRedraw = False
            For k = 3 To UBound(tbl1)
                tbl2 = tbl1(k).Split("-")
                i = Val(tbl2(0)) '+ 1 '< --- POUR ANCIENS FICHIERS
                j = Val(tbl2(1))
                ' ValeurNote = Val(tbl2(2)) ' non utilisé car redondant avec j
                If i <> Grid1.FixedRows - 1 Then
                    Durée = Val(tbl2(3)) ' notes
                    Vélocité = tbl2(4)
                    ' écriture de la note
                    ' *******************
                    Grid1.Cell(i, j).Text = Vélocité
                    Grid1.Cell(i, j).FontBold = True
                    For m = j + 1 To (Durée + j - 1)
                        Grid1.Cell(i, m).Text = Trait
                        Grid1.Cell(i, m).FontBold = True
                    Next
                Else ' CTRL modulation = 1 pour ligne (i) = 5
                    Grid1.Cell(i, j).Text = Trim(tbl2(2))
                End If
            Next k
            Grid1.AutoRedraw = True
            Grid1.Refresh()
            Grid1.Visible = True
        End If
    End Sub
    Public Sub Charger_Ctrl(LesCourbes As String)
        Dim j, ind As Integer
        Dim tbl1() As String
        Dim tbl2() As String
        Dim Ligne, Colonne As Integer
        Dim Fin As Integer

        tbl1 = LesCourbes.Split(",")
        If tbl1.Count >= 5 Then
            ind = tbl1(3) ' lecture du n° de ctrl (0 ..nbcourbes)
            Fin = GridCourbes(ind).Rows - 1
            GridCourbes(ind).Visible = False
            GridCourbes(ind).AutoRedraw = False
            For j = 4 To tbl1.Count - 1
                tbl2 = tbl1(j).Split()
                Ligne = tbl2(0)
                Colonne = tbl2(1)
                GridCourbes(ind).Range(Ligne, Colonne, Fin, Colonne).BackColor = Det_CouleurCTRL(ind)
            Next j
            GridCourbes(ind).AutoRedraw = True
            GridCourbes(ind).Refresh()
            GridCourbes(ind).Visible = True
        End If
    End Sub
    ''' <summary>
    ''' Construction_F1 : Méthode de Construction du graphique du PianoRoll
    ''' 1 - configuration du splitcontainer contenant la barre d'outil et la grille : Panneau1
    ''' 2 - appel méthode Construction_Grid1() : construction de la grille de pianoroll
    ''' 3 - appel méthode Construction_BarreOutils() : contrution de la barre d'outils
    ''' 4 - appel méthodeConstruction_Menu() : contruction du menu édition du pianiroll
    ''' </summary>
    Public Sub Construction_F1()
        Dim s As New Size With {
            .Width = 0,
            .Height = 0
            }
        F1.Visible = False
        F1.Text = "PIANO ROLL" + Str(Me.Canal - 5)
        ' Panneau 2
        ' *********
        F1.Controls.Add(Panneau2) ' ajout du splitcontainer dans le formulaire F1
        Panneau2.FixedPanel = FixedPanel.Panel1
        Panneau2.Dock = DockStyle.Fill
        Panneau2.Orientation = Orientation.Horizontal
        Panneau2.Panel1.MaximumSize = s
        Panneau2.Panel2.MaximumSize = s
        Panneau2.Panel1.MinimumSize = s
        Panneau2.Panel2.MinimumSize = s
        Panneau2.Panel1MinSize = 25
        Panneau2.FixedPanel = FixedPanel.None
        Panneau2.SplitterWidth = 4
        '
        ' Panneau 1
        ' *********
        Panneau2.Panel1.Controls.Add(Panneau1)
        Panneau1.Dock = DockStyle.Fill
        Panneau1.Orientation = Orientation.Horizontal
        Panneau1.SplitterDistance = 50
        'Panneau1.Panel1.BorderStyle = BorderStyle.Fixed3D
        Panneau1.FixedPanel = FixedPanel.Panel1
        Panneau1.IsSplitterFixed = True
        Panneau1.Panel1.MaximumSize = s
        Panneau1.Panel2.MaximumSize = s
        Panneau1.Panel1.MinimumSize = s
        Panneau1.Panel2.MinimumSize = s
        Panneau1.Panel1MinSize = 25
        ' Traitemeent de grid1
        Panneau1.Panel2.Controls.Add(Grid1)
        Grid1.Dock = DockStyle.Fill
        '
        F1.FormBorderStyle = FormBorderStyle.SizableToolWindow
        F1.ControlBox = False
        ' 
        Maj_TabNotes()
        Construction_Grid1()
        Construction_BarreOutils()
        Construction_Menu()
        '
        Grid1.AutoRedraw = False
        Graphique_Gammes()
        Graphique_Divisions()

        ' Construction courbes
        Construction_Courbes()
        '
        Grid1.AutoRedraw = True
        Grid1.Refresh()
        '
        ' Menu contextuel pour vélodités aléatoires
        Construction_MenuContext()

        ' A placer obligatoirement en fin de création de l'IHM
        Position_Init() ' taille du formulaire F1 et Positionnement du splitter du haut(splitter Panneau1)
        Hide_Ctrls()
    End Sub
    Sub Show_Ctrls()
        Dim i As Integer
        MidiLearn.Enabled = True
        'Panneau2.SplitterDistance = Panneau2.Size.Height - PanelCourbes.Size.Height
        '
        Panneau2.IsSplitterFixed = False
        'If Me.F1.Dock = DockStyle.Fill Then
        ''Panneau2.SplitterDistance = 405 '410
        'Panneau2.SplitterDistance = Panneau2.Height - 195
        'Else
        Panneau2.SplitterDistance = Panneau2.Height - 195
        'End If
        Panneau2.IsSplitterFixed = True
        '
        PanelCourbes.Visible = True
        '
        For i = 0 To nbCourbes - 1
            CCActif.Item(i).Enabled = True
        Next
        '
        TabCourbes.BringToFront()

    End Sub
    Sub Hide_Ctrls()
        Dim i As Integer
        MidiLearn.Enabled = False
        MidiLearn.Checked = False
        Panneau2.IsSplitterFixed = False
        Panneau2.SplitterDistance = Panneau2.Size.Height
        Panneau2.IsSplitterFixed = False
        PanelCourbes.Visible = False
        '
        For i = 0 To nbCourbes - 1
            CCActif.Item(i).Enabled = False ' la valeur checked des CTRL est toujours lisible
        Next
    End Sub
    Sub Construction_Courbes()
        Dim i, j As Integer
        ' Conteneur des cours
        Try


            Panneau2.Panel2.Controls.Add(PanelCourbes)   ' ajout du conteneur destiné aux courbes de contrôleurs
            Dim s1 As New Size With {
            .Width = Panneau1.Size.Width,
            .Height = 184}
            PanelCourbes.Size = s1
            PanelCourbes.Dock = DockStyle.Bottom
            PanelCourbes.BackColor = Color.BlueViolet

            ' Boite à onglet des courbes
            PanelCourbes.Controls.Add(TabCourbes)
            TabCourbes.Dock = DockStyle.Fill
            PanelCourbes.Dock = DockStyle.Fill
            TabCourbes.Appearance = TabAppearance.Buttons
            TabCourbes.Font = fnt6

            ' ajout du contrôle d'onglets dans le conteneur PanelCourbes des courbes

            For i = 0 To (nbCourbes - 1)

                GridCourbes.Add(New FlexCell.Grid)  ' création d'une grille
                PageCourbes.Add(New TabPage)       ' création d'un onglet
                PageCourbes.Item(i).Controls.Add(GridCourbes.Item(i)) ' placement de la grille dans l'onglet
                TabCourbes.Controls.Add(PageCourbes.Item(i)) ' placement de l'onglet dans son Tabcontrol
                '
                PageCourbes.Item(i).Tag = i
                GridCourbes.Item(i).Tag = i
                PageCourbes.Item(i).Dock = DockStyle.Fill
                PageCourbes.Item(i).BorderStyle = BorderStyle.Fixed3D
                GridCourbes.Item(i).Dock = DockStyle.Fill
                GridCourbes.Item(i).SelectionMode = FlexCell.SelectionModeEnum.ByColumn
                GridCourbes.Item(i).SelectionBorderColor = Color.White
                '
                GridCourbes.Item(i).ScrollBars = ScrollBarsEnum.Both
                GridCourbes.Item(i).FixedRows = 1
                GridCourbes.Item(i).FixedCols = 1
                '
                GridCourbes.Item(i).Cols = Grid1.Cols

                '
                GridCourbes.Item(i).BackColorFixed = Color.White
                GridCourbes.Item(i).BorderStyle = BorderStyleEnum.FixedSingle
                GridCourbes.Item(i).CellBorderColor = Color.Black
                '
                GridCourbes.Item(i).BackColorSel = Color.Transparent
                GridCourbes.Item(i).Cursor = Cursors.Arrow
                'Grid1.SelectionMode = FlexCell.SelectionModeEnum.ByColumn
                '
                ' maj grid pour les courbes
                GridCourbes.Item(i).Rows = 66
                For j = 0 To (GridCourbes.Item(i).Rows) - 1
                    GridCourbes.Item(i).Row(j).Height = 2
                Next

                For j = 0 To GridCourbes.Item(i).Cols - 1
                    GridCourbes.Item(i).Column(j).Width = LargCol
                Next
                '
                GridCourbes.Item(i).Row(GridCourbes.Item(i).Rows - 1).Height = 7
                '
                Select Case i
                    Case 0
                        PageCourbes.Item(i).Text = "Expression"
                        GridCourbes.Item(i).BackColorFixed = Det_CouleurCTRL(i)
                        GridCourbes.Item(i).BackColorFixedSel = Det_CouleurCTRL(i)
                    Case 1
                        PageCourbes.Item(i).Text = "Modulation"
                        GridCourbes.Item(i).BackColorFixed = Det_CouleurCTRL(i)
                        GridCourbes.Item(i).BackColorFixedSel = Det_CouleurCTRL(i)
                    Case 2
                        PageCourbes.Item(i).Text = "PAN"
                        GridCourbes.Item(i).BackColorFixed = Det_CouleurCTRL(i)
                        GridCourbes.Item(i).BackColorFixedSel = Det_CouleurCTRL(i)
                    Case 3
                        PageCourbes.Item(i).Text = "CC50"
                        GridCourbes.Item(i).BackColorFixed = Det_CouleurCTRL(i)
                        GridCourbes.Item(i).BackColorFixedSel = Det_CouleurCTRL(i)
                    Case 4
                        PageCourbes.Item(i).Text = "CC51"
                        GridCourbes.Item(i).BackColorFixed = Det_CouleurCTRL(i)
                        GridCourbes.Item(i).BackColorFixedSel = Det_CouleurCTRL(i)
                    Case 5
                        PageCourbes.Item(i).Text = "CC52"
                        GridCourbes.Item(i).BackColorFixed = Det_CouleurCTRL(i)
                        GridCourbes.Item(i).BackColorFixedSel = Det_CouleurCTRL(i)
                    Case 6
                        PageCourbes.Item(i).Text = "CC53"
                        GridCourbes.Item(i).BackColorFixed = Det_CouleurCTRL(i)
                        GridCourbes.Item(i).BackColorFixedSel = Det_CouleurCTRL(i)


                End Select
                AddHandler GridCourbes.Item(i).KeyUp, AddressOf GridCourbes_KeyUp
                AddHandler GridCourbes.Item(i).KeyDown, AddressOf GridCourbes_KeyDown2
                AddHandler GridCourbes.Item(i).SelChange, AddressOf GridCourbes_Selchange
                AddHandler GridCourbes.Item(i).MouseDown, AddressOf GridCourbes_MouseDown
                AddHandler GridCourbes.Item(i).MouseUp, AddressOf GridCourbes_MouseUp
                AddHandler GridCourbes.Item(i).Scroll, AddressOf GridCourbes_Scroll

            Next i
            '
            ' Construction des Onglet Support
            ' *******************************
            Construction_GridAssist2_Tonalités(7)
            Construction_GridAssist1_Interv(8)



        Catch ex As Exception
            MsgBox("Procédure pianoroll/Construction_Courbes - index i =" + Str(i))
        End Try

    End Sub
    Private Sub Construction_GridAssist2_Tonalités(N_Onglet As Integer)
        ' Préparation 
        ' ***********
        PageCourbes.Add(New TabPage)       ' création d'un onglet
        PageCourbes.Item(N_Onglet).Controls.Add(Assist2SplitC) ' placement du split container dans l'onglet
        Assist2SplitC.Dock = DockStyle.Left
        Assist2SplitC.Orientation = Orientation.Horizontal
        Assist2SplitC.BorderStyle = BorderStyle.FixedSingle
        Assist2SplitC.SplitterDistance = 22
        Assist2SplitC.IsSplitterFixed = True
        ' Panel 2
        Assist2SplitC.Panel2.Controls.Add(GridAssist2)
        Assist2SplitC.Panel2.BorderStyle = BorderStyle.FixedSingle
        ' Panel 1
        Assist2SplitC.Panel1.BackColor = Color.PaleGoldenrod
        '
        ' bouton Ecriture
        ' ***************
        Dim P As New Point
        P.X = 7
        P.Y = 1
        Dim S As New Size
        S.Width = 100
        S.Height = 37
        Assist2SplitC.Panel1.Controls.Add(Assist2Bout)
        Assist2Bout.Location = P
        Assist2Bout.BackColor = Color.Gold
        'Assist2Bout.ForeColor = Color.Khaki
        Assist2Bout.Size = S
        If LangueIHM = "fr" Then
            Assist2Bout.Text = "Ecrire gamme"
        Else
            Assist2Bout.Text = "Write scale"
        End If
        '
        ' Destination
        ' ***********
        S.Width = 60
        S.Height = 20
        '
        ' liste destination
        Assist2SplitC.Panel1.Controls.Add(Assist2Deb)
        Assist2SplitC.Panel1.Controls.Add(Assist2Lab2)
        P.X = 130
        P.Y = 5
        Assist2Deb.Location = P
        Assist2Deb.Size = S
        Assist2Deb.Maximum = Module1.nbMesures
        Assist2Deb.Minimum = 1
        Assist2Deb.Visible = False

        P.X = 130
        P.Y = 25
        Assist2Lab2.Location = P
        Assist2Lab2.AutoSize = True
        Assist2Lab2.Font = ft2
        If LangueIHM = "fr" Then
            Assist2Lab2.Text = "Destination"
        Else
            Assist2Lab2.Text = "Destination"
        End If
        Assist2Lab2.Visible = False
        '
        ' TextBox d'aide
        ' **************
        P.X = 210
        P.Y = 6
        S.Height = 25
        S.Width = 450
        Assist2SplitC.Panel1.Controls.Add(Assist2Text1)
        Assist2Text1.Location = P
        Assist2Text1.Size = S
        Assist2Text1.AutoSize = False
        Assist2Text1.BackColor = Color.White
        Assist2Text1.ForeColor = Color.Red
        Assist2Text1.AcceptsReturn = True
        Assist2Text1.Multiline = True
        Assist2Text1.WordWrap = True
        Assist2Text1.BorderStyle = BorderStyle.FixedSingle
        If LangueIHM = "fr" Then
            Assist2Text1.Text = "Veuillez sélectionner au moins 3 notes dans le Piano Roll"
        Else
            Assist2Text1.Text = "Please select at least 3 notes in the Piano Roll"
        End If

        Assist2Text1.Visible = False

        AddHandler Assist2Bout.MouseClick, AddressOf Assist2Bout_MouseClick
        ' AddHandler Assist2Deb.ValueChanged, AddressOf Assist2Deb_ValueChanged
        ' AddHandler Assist2Fin.ValueChanged, AddressOf Assist2Fin_ValueChanged
        ' AddHandler Assist2Deb.MouseClick, AddressOf Assist2_RAZ_Erreur ' les 2 évènements poitent sur la même procédure pour effacer la zone de message Assist2Text1
        ' AddHandler Assist2Fin.MouseClick, AddressOf Assist2_RAZ_Erreur
        '
        TabCourbes.Controls.Add(PageCourbes.Item(N_Onglet)) ' placement de l'onglet dans son Tabcontrol
        '
        PageCourbes.Item(N_Onglet).Dock = DockStyle.Fill
        PageCourbes.Item(N_Onglet).AutoScroll = False
        If LangueIHM = "fr" Then
            PageCourbes.Item(N_Onglet).Text = "Gammes (beta)"
        Else
            PageCourbes.Item(N_Onglet).Text = "Scales (beta)"
        End If
        'PageCourbes.Item(N_Onglet).BorderStyle = BorderStyle.Fixed3D
        PageCourbes.Item(N_Onglet).Tag = N_Onglet

        With GridAssist2
            .Tag = N_Onglet
            .Dock = DockStyle.Left
            .BackColor1 = Color.WhiteSmoke
            .BackColorFixed = Color.Ivory
            .SelectionMode = FlexCell.SelectionModeEnum.ByRow
            .SelectionBorderColor = Color.White
            .ScrollBars = ScrollBarsEnum.Vertical
            .FixedRows = 1
            .FixedCols = 1
            .Height = 250
            .Rows = 128 '128
            .Cols = 6
            .Width = 1000
            '
            For r = 0 To .Rows - 1
                .Row(r).Height = 20
            Next r
            '
            For r = 1 To .Rows - 1
                .Cell(r, 0).Text = (r).ToString
            Next r
            '
            For c = 0 To .Cols - 1
                .Column(c).Width = 26
            Next c
            '
            .Column(1).Width = 100 ' nom de la gammes
            .Column(2).Width = 310 ' notes de la gamme
            .Column(3).Width = 106
            .Column(4).Width = 106
            .Column(3).Alignment = AlignmentEnum.CenterCenter
            .Column(4).Alignment = AlignmentEnum.CenterCenter
            '.Column(4).Width = 70
            '.Column(5).Width = 70
            .Width = (26 * 16) + 36 + 218
            Assist2SplitC.Width = .Width
            '
            ' titres
            If LangueIHM = "fr" Then
                .Cell(0, 1).Text = "Gammes"
                .Cell(0, 2).Text = "Notes des gammes"
                .Cell(0, 3).Text = "Tonalité suggérée"
                .Cell(0, 4).Text = "Accord basique"
            Else
                .Cell(0, 1).Text = "Scales"
                .Cell(0, 2).Text = "Scales Notes"
                .Cell(0, 3).Text = "Suggested tone"
                .Cell(0, 4).Text = "Basic Chord"
            End If
            '
        End With
    End Sub
    Sub Assist2_RAZ_Erreur(sender As Object, e As EventArgs)
        Assist2Text1.Text = ""
    End Sub
    Private Sub Assist2Deb_ValueChanged(sender As Object, e As EventArgs)

    End Sub
    Private Sub Assist2Bout_MouseClick(sender As Object, e As EventArgs)
        Dim i As String = GridAssist2.ActiveCell.Row
        Dim Tona As String = Trim(GridAssist2.Cell(i, 3).Text)
        Dim Mode As String = Tona
        Dim Gamme As String = Trim(GridAssist2.Cell(i, 1).Text)
        Dim Acc As String = Trim(GridAssist2.Cell(i, 4).Text)
        Dim Degré As Integer = 0
        Dim ligne As Integer = i
        Dim m As Integer = Assist2Deb.Value
        Dim tbl() As String

        tbl = Gamme.Split()
        If tbl(1) = "Blues" Or tbl(1) = "Blues2" Then
            Degré = 4
        End If

        Form1.SAUV_Annuler(m, m) ' pour CTRL Z
        flagAssist2_Z = True ' pour  CTRL Z de type onglet HyperVoicing

        ' Maj des données de l'accord dans TableEventh
        ' ********************************************
        TableEventH(m, 1, 1).Tonalité = Tona
        TableEventH(m, 1, 1).Mode = Mode '
        TableEventH(m, 1, 1).Gamme = Gamme
        TableEventH(m, 1, 1).Accord = Trim(Acc)
        TableEventH(m, 1, 1).NumAcc = m
        TableEventH(m, 1, 1).Position = Str(m) + ".1" + ".1"
        TableEventH(m, 1, 1).Degré = Degré
        TableEventH(m, 1, 1).Ligne = ligne
        TableEventH(m, 1, 1).Racine = "b1"
        TableEventH(m, 1, 1).Vel = "100"
        '
        ' Ecriture et maj couleur dans Grid2
        ' **********************************
        Form1.Grid2.Cell(2, m).Alignment = FlexCell.AlignmentEnum.CenterCenter
        Form1.Grid2.Cell(11, m).Alignment = FlexCell.AlignmentEnum.CenterCenter
        Form1.Grid2.Cell(2, m).Text = Acc
        Form1.Grid2.Cell(11, m).Text = TableEventH(m, 1, 1).Gamme
        Form1.Grid2.Cell(12, m).Text = TableEventH(m, 1, 1).Racine
        '
        '
        Form1.Grid2.ReadonlyFocusRect = FlexCell.FocusRectEnum.Solid
        '
        Form1.Grid2.AutoRedraw = False
        '
        a = TableEventH(m, 1, 1).Tonalité '
        a = Form1.Det_RelativeMajeure2(a) ' si la tonalité est mineure alors on affiche la couleur de la relative majeure
        tbl = Split(a)
        Form1.Grid2.Cell(2, m).BackColor = DicoCouleur.Item(Trim(tbl(0))) ' la couleur de l'accord est fonction de la tonalité
        Form1.Grid2.Cell(2, m).ForeColor = DicoCouleurLettre.Item(tbl(0))
        Form1.Grid2.Cell(11, m).BackColor = DicoCouleur.Item(Trim(tbl(0))) ' la couleur de la gamme est fonction de la tonalité
        Form1.Grid2.Cell(11, m).ForeColor = DicoCouleurLettre.Item(tbl(0))
        '
        Form1.Grid2.Refresh()
        Form1.Grid2.AutoRedraw = True
        '
        '
        '
        If m = 1 Then
            Form1.Grid2.Cell(2, m).BackColor = Color.Red
            Form1.Grid2.Cell(2, m).ForeColor = Color.Yellow
            Form1.Grid2.Cell(11, m).BackColor = Color.Red
            Form1.Grid2.Cell(11, m).ForeColor = Color.Yellow
        End If
        '
        Form1.Formater_Racine(m) ' formater la racine : couleur et alignement
        '
        Form1.Maj_CopieAcc(m)
        '
        OngletCours_Edition = 0 ' on utilise le TCRT Z de l'onglet Hypervoicing

    End Sub
    Private Sub Assist2Fin_ValueChanged(sender As Object, e As EventArgs)
        Assist2Text1.Text = ""
        If Assist2Fin.Value <= Assist2Deb.Value Then
            Assist2Fin.Value = Assist2Fin.Value + 1

            If LangueIHM = "fr" Then
                Assist2Text1.Text = "La fin de la sélection ne peut pas être inférieure ou égale au début"
            Else
                Assist2Text1.Text = "The end of the selection cannot be lower or equal than the beginning"
            End If

        End If
    End Sub


    Private Sub Construction_GridAssist1_Interv(N_Onglet As Integer)
        Dim r As Integer
        Dim c As Integer

        ' Préparation 
        ' ***********
        PageCourbes.Add(New TabPage)       ' création d'un onglet
        PageCourbes.Item(N_Onglet).Controls.Add(GridAssist1) ' placement de la grille dans l'onglet
        TabCourbes.Controls.Add(PageCourbes.Item(N_Onglet)) ' placement de l'onglet dans son Tabcontrol
        '
        PageCourbes.Item(N_Onglet).Dock = DockStyle.Fill
        PageCourbes.Item(N_Onglet).AutoScroll = False
        PageCourbes.Item(N_Onglet).Text = "Consonnances (beta)"
        'PageCourbes.Item(N_Onglet).BorderStyle = BorderStyle.Fixed3D
        PageCourbes.Item(N_Onglet).Tag = N_Onglet
        With GridAssist1
            .Tag = N_Onglet
            .Dock = DockStyle.Left
            .SelectionMode = FlexCell.SelectionModeEnum.ByCell
            .SelectionBorderColor = Color.LightCyan
            .ScrollBars = ScrollBarsEnum.Vertical
            .FixedRows = 1
            .FixedCols = 1
            .Height = 250
            .Rows = 128 '128
            .Cols = 26
            '
            For r = 0 To .Rows - 1
                .Row(r).Height = 26
            Next r
            For c = 0 To .Cols - 1
                .Column(c).Width = 26
            Next c
            .Column(13).Width = 36
            '.Column(26).Width = 654
            .Width = (26 * 26) + 31 '+ 570

            For c = 1 To 13
                Select Case c
                    Case 1
                        .Cell(0, c).Text = "8"
                        .Cell(0, 26 - c).Text = "8"
                        .Range(1, c, .Rows - 1, c).BackColor = Color.Wheat
                        .Range(1, 26 - c, .Rows - 1, 26 - c).BackColor = Color.Wheat
                    Case 2
                        .Cell(0, c).Text = "7"
                        .Cell(0, 26 - c).Text = "7"
                    Case 3
                        .Cell(0, c).Text = "7b"
                        .Cell(0, 26 - c).Text = "7b"
                    Case 4
                        .Cell(0, c).Text = "6"
                        .Cell(0, 26 - c).Text = "6"
                        .Range(1, c, .Rows - 1, c).BackColor = Color.Wheat
                        .Range(1, 26 - c, .Rows - 1, 26 - c).BackColor = Color.Wheat
                    Case 5
                        .Cell(0, c).Text = "6m"
                        .Cell(0, 26 - c).Text = "6m"
                        .Range(1, c, .Rows - 1, c).BackColor = Color.Wheat
                        .Range(1, 26 - c, .Rows - 1, 26 - c).BackColor = Color.Wheat
                    Case 6
                        .Cell(0, c).Text = "5"
                        .Cell(0, 26 - c).Text = "5"
                        .Range(1, c, .Rows - 1, c).BackColor = Color.Wheat
                        .Range(1, 26 - c, .Rows - 1, 26 - c).BackColor = Color.Wheat
                    Case 7
                        .Cell(0, c).Text = "5b"
                        .Cell(0, 26 - c).Text = "5b"
                    Case 8
                        .Cell(0, c).Text = "4"
                        .Cell(0, 26 - c).Text = "4"
                    Case 9
                        .Cell(0, c).Text = "3"
                        .Cell(0, 26 - c).Text = "3"
                        .Range(1, c, .Rows - 1, c).BackColor = Color.Wheat
                        .Range(1, 26 - c, .Rows - 1, 26 - c).BackColor = Color.Wheat
                    Case 10
                        .Cell(0, c).Text = "3m"
                        .Cell(0, 26 - c).Text = "3m"
                        .Range(1, c, .Rows - 1, c).BackColor = Color.Wheat
                        .Range(1, 26 - c, .Rows - 1, 26 - c).BackColor = Color.Wheat
                    Case 11
                        .Cell(0, c).Text = "9"
                        .Cell(0, 26 - c).Text = "9"
                    Case 12
                        .Cell(0, c).Text = "9b"
                        .Cell(0, 26 - c).Text = "9b"
                    Case 13
                        .Cell(0, c).Text = "CF"
                        .Range(1, c, .Rows - 1, c).BackColor = Color.Gold
                End Select
                '
                '.Cell(0, 26).Text = "Commentaires"
            Next

            For r = 1 To .Rows - 1
                .Cell(r, 0).Text = Str(r)
            Next
            '

            .Range(1, 13, .Rows - 1, 13).Font = ft1
            .Range(1, 1, .Rows - 1, .Cols - 1).Alignment = AlignmentEnum.CenterCenter
            .Range(1, 1, .Rows - 1, .Cols - 1).Locked = True
            .Range(1, 13, .Rows - 1, 13).Locked = False
        End With
        ' Param divers
        ' ************
        CouleurMarq = Color.Tan
        ' Evènements
        ' **********
        AddHandler GridAssist1.KeyUp, AddressOf GridAssist1_KeyUp
        AddHandler GridAssist1.MouseUp, AddressOf GridAssist1_MouseUp


    End Sub
    Sub GridAssist1_MouseUp(Sender As Object, e As MouseEventArgs)
        Dim r As Integer = GridAssist1.ActiveCell.Row
        Dim c As Integer = GridAssist1.ActiveCell.Col
        '
        GridAssist1.AutoRedraw = False
        '
        If c <> 13 Then ' 13 = Cantus Firmus
            If GridAssist1.Cell(r, c).BackColor = CouleurMarq Then
                'effacer
                EffacerCouleurs(r)
            Else
                ' effacer et marquer
                EffacerCouleurs(r)
                GridAssist1.Cell(r, c).BackColor = CouleurMarq
            End If
        End If
        '
        GridAssist1_CoulImpair()
        '
        GridAssist1.AutoRedraw = True
        GridAssist1.Refresh()
    End Sub

    Sub EffacerCouleurs(Ligne As Integer)

        With GridAssist1
            .Range(Ligne, 1, Ligne, GridAssist1.Cols - 1).BackColor = Color.White
            ' 8 
            .Range(Ligne, 1, Ligne, 1).BackColor = Color.Wheat
            .Range(Ligne, 25, Ligne, 25).BackColor = Color.Wheat
            ' sixte
            .Range(Ligne, 4, Ligne, 4).BackColor = Color.Wheat
            .Range(Ligne, 22, Ligne, 22).BackColor = Color.Wheat
            ' sixte mineure
            .Range(Ligne, 5, Ligne, 5).BackColor = Color.Wheat
            .Range(Ligne, 21, Ligne, 21).BackColor = Color.Wheat
            ' quinte
            .Range(Ligne, 6, Ligne, 6).BackColor = Color.Wheat
            .Range(Ligne, 20, Ligne, 20).BackColor = Color.Wheat
            ' Tierce majeure     
            .Range(Ligne, 9, Ligne, 9).BackColor = Color.Wheat
            .Range(Ligne, 17, Ligne, 17).BackColor = Color.Wheat
            ' Tierce mineure
            .Range(Ligne, 10, Ligne, 10).BackColor = Color.Wheat
            .Range(Ligne, 16, Ligne, 16).BackColor = Color.Wheat
            ' Cantus Firmus
            .Range(Ligne, 13, Ligne, 13).BackColor = Color.Gold

        End With
    End Sub

    Sub GridAssist1_KeyUp(Sender As Object, e As KeyEventArgs)
        With GridAssist1
            Dim r As Integer = .ActiveCell.Row
            Dim c As Integer = .ActiveCell.Col
            '
            Dim i As Integer = -1
            Dim j As Integer
            Dim a As String = Trim(.Cell(r, Col_CF).Text)
            '
            'If e.KeyData = 13 Then ' touche entrée
            Do
                i += 1
            Loop Until (i = Chromad.Count - 1) Or (a = Chromad.Item(i))
            '
            If i < Chromad.Count - 1 Then
                ' Voix haute
                .Cell(r, Col_CF + 12).Text = Chromad(i + Octave)
                .Cell(r, Col_CF + 11).Text = Chromad(i + Sept)
                .Cell(r, Col_CF + 10).Text = Chromad(i + Septb)
                .Cell(r, Col_CF + 9).Text = Chromad(i + SixteMaj)
                .Cell(r, Col_CF + 8).Text = Chromad(i + SixteMin)
                .Cell(r, Col_CF + 7).Text = Chromad(i + Quinte)
                .Cell(r, Col_CF + 6).Text = Chromad(i + Quinteb)
                .Cell(r, Col_CF + 5).Text = Chromad(i + Quarte)
                .Cell(r, Col_CF + 4).Text = Chromad(i + TierceMaj)
                .Cell(r, Col_CF + 3).Text = Chromad(i + TierceMin)
                .Cell(r, Col_CF + 2).Text = Chromad(i + Neuf)
                .Cell(r, Col_CF + 1).Text = Chromad(i + Neufb)
                '.Cell(r, Col_CF).Text = Chromad(i + Octave)
                ' Voix Basse
                '.Cell(r, 13).Text = Chromad(i + EstOctave)
                .Cell(r, 12).Text = Chromad(i + EstNeufb)
                .Cell(r, 11).Text = Chromad(i + EstNeuf)
                .Cell(r, 10).Text = Chromad(i + EstTierceMin)
                .Cell(r, 9).Text = Chromad(i + EstTierceMaj)
                .Cell(r, 8).Text = Chromad(i + EstQuarte)
                .Cell(r, 7).Text = Chromad(i + EstQuinteb)
                .Cell(r, 6).Text = Chromad(i + EstQuinte)
                .Cell(r, 5).Text = Chromad(i + EstSixteMin)
                .Cell(r, 4).Text = Chromad(i + EstSixteMaj)
                .Cell(r, 3).Text = Chromad(i + EstSeptb)
                .Cell(r, 2).Text = Chromad(i + EstSept)
                .Cell(r, 1).Text = Chromad(i + EstOctave)
            Else
                .Cell(r, Col_CF).Text = ""
            End If
            'End If
            '
            If e.KeyData = 46 Or e.KeyData = 8 Then ' touche suppr ou touche effacer (back)
                If c = Col_CF Then
                    For j = 1 To Maxcol
                        .Cell(r, j).Text = ""
                    Next j
                End If
            End If
            '
        End With
    End Sub
    Sub GridAssist1_CoulImpair()
        Dim i, j As Integer

        For i = 1 To GridAssist1.Rows - 1
            For j = 1 To GridAssist1.Cols - 1
                If EstImpair(i) And GridAssist1.Cell(i, j).BackColor <> CouleurMarq Then
                    Select Case j
                        Case 2, 3, 7, 8, 11, 12, 14, 15, 18, 19, 23, 24
                            GridAssist1.Cell(i, j).BackColor = Color.Beige
                    End Select
                End If
            Next
        Next
    End Sub

    Public Sub Nouv_GridAssist1()
        With GridAssist1
            .AutoRedraw = False
            '
            For i = 1 To GridAssist1.Rows - 1
                For j = 1 To GridAssist1.Cols - 1
                    GridAssist1.Cell(i, j).Text = ""
                Next
                EffacerCouleurs(i)
            Next
            '
            .Range(1, 13, .Rows - 1, 13).Font = ft1
            .Range(1, 1, .Rows - 1, .Cols - 1).Alignment = AlignmentEnum.CenterCenter
            .Range(1, 1, .Rows - 1, .Cols - 1).Locked = True
            .Range(1, 13, .Rows - 1, 13).Locked = False
            '
            .AutoRedraw = True
            .Refresh()
        End With
    End Sub

    Private Sub GridCourbes_Selchange(Sender As Object, e As Grid.SelChangeEventArgs)
        Tracé_Courbe(Sender)
    End Sub
    Sub Tracé_Courbe(Sender As Object)
        Dim com As FlexCell.Grid = Sender
        Dim ind As Integer = com.Tag
        Dim i As Integer = GridCourbes.Item(ind).MouseRow
        Dim j As Integer = GridCourbes.Item(ind).MouseCol
        Dim ligne As Integer = 0
        '
        GridCourbes.Item(ind).AutoRedraw = False
        '
        If ind <= nbCourbes - 1 Then '- 1
            If i <> -1 And j <> -1 Then
                CCActif.Item(ind).Checked = True ' dès que l'on trace dans une courbe, la case à cocher correspondante dans la barre d'outils est cochée.
                GridCourbes.Item(ind).AutoRedraw = False
                'If My.Computer.Keyboard.ShiftKeyDown Then
                If My.Computer.Keyboard.CtrlKeyDown And Orig_PianoR.Vérouillage = False Then ' vérouillage est mis à jour dans l'évènement CTRL + V : il empêche d'écrire dans la courbe quand on fait un "coller" avec CTRL+V
                    GridCourbes.Item(ind).Range(1, j, GridCourbes.Item(ind).Rows - 1, j).ClearBackColor()
                    '
                    If i <> GridCourbes.Item(ind).Rows - 1 Then ' ce test permet d'obtenir la RAZ de la dernière cellule du bas par Ctrl + clic
                        GridCourbes.Item(ind).Range(GridCourbes.Item(ind).Rows - 1, j, i, j).BackColor = Det_CouleurCTRL(ind)
                    End If
                    GridCourbes.Item(ind).Refresh()
                End If
                GridCourbes.Item(ind).AutoRedraw = True
                GridCourbes.Item(ind).Refresh()
            End If
            '
        End If
    End Sub

    Private Sub GridCourbes_MouseDown(Sender As Object, e As MouseEventArgs)
        Dim com As FlexCell.Grid = Sender
        Dim N_Courbe As Integer = com.Tag

        ' RUSTINE 1 : à chaque fois que l'on clique sur une courbe, la cellule active du piano roll devient la cellule (0,0).
        ' Ceci pour éviter le problème suivant : quand le PianoRoll est détaché et qu'une courbe est sélectionné, l'appui sur 
        ' la touche Suppr provoque non seulement la supression de la courbe sélectionné mais aussi la suppression d'une note 
        ' sélectionnée dans le piano roll. Le Bug ne se reproduit pas si le PianoRoll est attaché.

        Grid1.Cell(0, 0).SetFocus() ' <-- RUSTINE 1


        Orig_PianoR.Orig1 = OrigPianoCourbe.Courbe
        Orig_PianoR.N_Courbe = N_Courbe

        Tracé_Courbe(Sender)
        '

    End Sub
    Private Sub GridCourbes_MouseUp(Sender As Object, e As MouseEventArgs)
        Dim com As FlexCell.Grid = Sender
        Dim N_Courbe As Integer = com.Tag
        Dim ligne As Integer
        Dim j As Integer = GridCourbes.Item(N_Courbe).ActiveCell.Col
        Dim k As Integer = -1
        ' Affichage de la valeur dans AffMidiLearn
        ' ****************************************
        For ligne = 1 To GridCourbes.Item(N_Courbe).Rows - 1
            If GridCourbes.Item(N_Courbe).Cell(ligne, j).BackColor <> Color.White Then ' valeur trouvée
                k = ((GridCourbes.Item(N_Courbe).Rows - 1 - ligne) * 2) - 1
                Exit For
            End If
        Next ligne
        '
        If k = -1 Then
            AffMidiLearn.Text = "off"
        Else
            AffMidiLearn.Text = k.ToString
        End If
    End Sub

    Sub GridCourbes_Scroll(Sender As Object, e As EventArgs)
        Dim com As FlexCell.Grid = Sender
        Dim ind As Integer = com.Tag

        If ind <= 6 Then

            Grid1.LeftCol = GridCourbes.Item(ind).LeftCol
            For i = 0 To nbCourbes - 1
                If i <> ind Then
                    GridCourbes.Item(i).LeftCol = GridCourbes.Item(ind).LeftCol
                End If
            Next
        End If
    End Sub
    Sub GridCourbes_KeyUp(Sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Dim com As FlexCell.Grid = Sender
        Dim ind As Integer = com.Tag

        If (e.KeyCode = Keys.Delete Or e.KeyCode = Keys.Back) And Orig_PianoR.Orig1 = OrigPianoCourbe.Courbe Then
            Dim i As Integer = GridCourbes.Item(ind).Selection.FirstRow
            Dim j As Integer = GridCourbes.Item(ind).Selection.FirstCol
            Dim ii As Integer = GridCourbes.Item(ind).Selection.LastRow
            Dim jj As Integer = GridCourbes.Item(ind).Selection.LastCol
            '
            GridCourbes.Item(ind).AutoRedraw = False
            GridCourbes.Item(ind).Range(i, j, ii, jj).BackColor = Color.White
            GridCourbes.Item(ind).AutoRedraw = True
            GridCourbes.Item(ind).Refresh()
        End If
    End Sub

    Public Sub GridCourbes_KeyDown2(Sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Dim com As FlexCell.Grid = Sender
        Dim ind As Integer = com.Tag

        Dim i As Integer = GridCourbes.Item(ind).Selection.FirstRow ' on n'agit ici que sur une colonne à la fois.
        Dim j As Integer = GridCourbes.Item(ind).Selection.FirstCol
        Dim k As Integer
        Dim vv As String
        Dim v As Byte ' 
        Dim b As String


        If e.KeyCode <> Keys.ControlKey Then
            If Orig_PianoR.Orig1 = OrigPianoCourbe.Courbe Then
                If e.KeyCode = Keys.Add Or e.KeyCode = Keys.Subtract Then
                    b = ValCtrl2_Ligne(ind, j) ' 
                    If Trim(b) <> "" Then
                        k = Convert.ToInt16(Trim(b)) ' lecture de la ligne actuelle
                        '
                        ' Traitement incrémentation/décrémentation (si sortir = false)
                        ' ************************************************************
                        '
                        ' INCREMENTATION
                        ' **************
                        If e.KeyCode = Keys.Add Then
                            k -= 1
                            GridCourbes.Item(ind).AutoRedraw = False
                            GridCourbes.Item(ind).Range(k, j, GridCourbes.Item(ind).Rows - 1, j).BackColor = Det_CouleurCTRL(ind)
                            GridCourbes.Item(ind).AutoRedraw = True
                            GridCourbes.Item(ind).Refresh()
                            vv = ValCtrl(ind, j)
                            If Trim(vv) <> "-1" Then
                                v = Convert.ToByte(vv)
                                AffMidiLearn.Text = Trim(vv) 'v.ToString
                                If MidiLearn.Checked Then
                                    Dim c As Byte = Det_CTRL(ind)
                                    Send_CTRL(c, v, Canal) ' Canal est donné par la constructeur de la classe
                                End If
                            Else
                                AffMidiLearn.Text = "--"
                            End If

                            '
                            ' DECREMENTATION
                            ' **************
                        ElseIf e.KeyCode = Keys.Subtract Then
                            k += 1
                            If k <= GridCourbes.Item(ind).Rows - 1 Then '  And k <= GridCourbes.Item(ind).Rows - 1 
                                GridCourbes.Item(ind).AutoRedraw = False
                                GridCourbes.Item(ind).Range(1, j, GridCourbes.Item(ind).Rows - 1, j).BackColor = Color.White
                                If k <> GridCourbes.Item(ind).Rows - 1 Then
                                    GridCourbes.Item(ind).Range(GridCourbes.Item(ind).Rows - 1, j, k, j).BackColor = Det_CouleurCTRL(ind)
                                End If
                                GridCourbes.Item(ind).AutoRedraw = True
                                GridCourbes.Item(ind).Refresh()
                                vv = ValCtrl(ind, j)
                                If Trim(vv) <> "-1" Then
                                    v = Trim(vv) 'Convert.ToByte(vv)
                                    AffMidiLearn.Text = v.ToString
                                    If MidiLearn.Checked Then
                                        Dim c As Byte = Det_CTRL(ind)
                                        Send_CTRL(c, v, Canal) ' Canal est donné par la constructeur de la classe
                                    End If
                                Else
                                    AffMidiLearn.Text = "--"
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End Sub
    ''' <summary>
    ''' Send_CTRL : envoi d'un controleur sur un Canal donné
    ''' </summary>
    ''' <param name="Controleur">Numéro du controleur</param>
    ''' <param name ="Valeur">Valeur du controleur</param>
    ''' <param name="Canal">Numéro canal MIDI de la piste Pinaoroll (6,7,8)</param>
    Sub Send_CTRL(Controleur As Byte, Valeur As Byte, Canal As Integer)
        If EnChargement = False Then
            If Not (Form1.SortieMidi.Item(Form1.ChoixSortieMidi).IsOpen) Then
                Form1.SortieMidi.Item(Form1.ChoixSortieMidi).Open()
            End If
            Form1.SortieMidi.Item(Form1.ChoixSortieMidi).SendControlChange(Canal, Controleur, Valeur)
        End If
    End Sub
    Public Sub Effacer_Grid1()
        Grid1.Range(2, 1, Grid1.Rows - 1, Grid1.Cols - 1).ClearText()
    End Sub
    Public Sub Clear_Notes()
        Dim i, j As Integer
        'Grid1.Range(2, 1, Grid1.Rows - 1, Grid1.Cols - 1).ClearText()
        Grid1.AutoRedraw = False
        For i = Grid1.FixedRows - 1 To Grid1.Rows - 1
            For j = 1 To Grid1.Cols - 1
                Grid1.Cell(i, j).Text = ""
            Next
        Next
        Grid1.AutoRedraw = True
        Grid1.Refresh()
    End Sub
    Public Sub Init_BarrOut()
        CheckMute.Checked = True      ' activation du canal
        ListTypNote.SelectedIndex = 2 ' liste des types de notes
        ListDynF1.SelectedIndex = 37   ' liste des dynamiques
        ListPRG.SelectedIndex = 0     ' liste des programmes GM
        ListTonNotes.Enabled = False   ' liqste des Tons d'un MODE
        BRadio2.Checked = True        ' bouton radio PAN
        CheckAcc.Checked = False      ' calque Accord
        CheckTonique.Checked = False      ' Ton d'un MODE
        'DebMacroSel.Value = 1         ' macros sélection
        TermeMacroSel.Value = 2
        NomduSon.Text = ""
    End Sub

    ' *******************************************************
    ' EVENEMENTS LIES A F1, Grid1 et barre outil
    ' *******************************************************
    Private Sub DebMacroSel_ValueChange(sender As Object, e As EventArgs)
        If EnChargement = False Then
            If DebMacroSel.Value >= TermeMacroSel.Value Then
                DebMacroSel.Value = DebMacroSel.Value - 1
            End If
        End If
    End Sub
    Private Sub TermeMacroSel_ValueChange(sender As Object, e As EventArgs)
        If EnChargement = False Then
            If TermeMacroSel.Value <= DebMacroSel.Value Then
                TermeMacroSel.Value = TermeMacroSel.Value + 1
            End If
        End If
    End Sub
    Private Sub BMacroSel_Click(sender As Object, e As EventArgs)
        Dim i As Integer
        Dim FirstR As Integer = Grid1.FixedRows '(5)
        Dim LastR As Integer = Grid1.Rows - 1
        Dim FirstC As Integer = (DebMacroSel.Value - 1) * 16 + 1
        Dim LastC As Integer = TermeMacroSel.Value * 16
        Dim SauvTopRow As Integer = Grid1.TopRow

        ' Sélection
        ' *********
        i = Grid1.TopRow
        Grid1.Range(FirstR, FirstC, LastR, LastC).SelectCells()
        Grid1.LeftCol = LastC
        Grid1.TopRow = i
        '
        ' Copie
        ' *****
        CopierData()
        '
        ' Aller Vers (sélection de la cellule de départ de la copie)
        ' ***********************************************************
        '
        Dim ii As Integer = Convert.ToInt16(DestMacroSel.Value)
        Grid1.LeftCol = ((ii - 1) * 16) + 1
        Grid1.TopRow = Grid1.FixedRows
        Grid1.Range(Grid1.TopRow, Grid1.LeftCol, Grid1.TopRow, Grid1.LeftCol).SelectCells()
        '
        ' Coller
        ' ******
        CollerData2()
        'Form1.ReCalcul()
        '
        ' restitution de la toprow de départ
        Grid1.TopRow = SauvTopRow
    End Sub
    Private Sub Dockbutton_MouseClick(sender As Object, e As EventArgs)
        Dim NumOnglet As Integer = Convert.ToUInt16(F1.Tag)

        If Me.F1.Dock = DockStyle.Fill Then
            Me.F1.Visible = False '   --> DETACHER
            Me.F1.FormBorderStyle = FormBorderStyle.SizableToolWindow
            F1.Text = "PIANO ROLL" + Str(Me.Canal - 5)
            ' 
            Me.F1.Dock = DockStyle.None
            Form1.TabControl4.TabPages.Item(NumOnglet).Controls.Remove(Me.F1)
            Dim p As New Point With {
                .X = Me.F1.Location.X,
                .Y = Me.F1.Location.Y + 10
            }
            Me.F1.Location = p
            Me.F1.StartPosition = FormStartPosition.Manual ' permet de tenir compte de la location calculée dans p
            Me.F1.TopMost = True
            Me.F1.TopLevel = True '
            Me.F1.MainMenuStrip.Visible = True
            Me.F1.MainMenuStrip.Enabled = True

            'Menu1.Visible = True
            'Menu1.Enabled = True
            'ActivationHandlerEdition()
            '

            Dim s As New Size With {
                .Width = Panneau1.Width + 10,
                .Height = Panneau1.Height + 240
            }
            Me.F1.Size = s
            Me.F1.FormBorderStyle = FormBorderStyle.Sizable
            '
            If LangueIHM = "fr" Then
                DockButton.Text = "Attacher"
            Else
                DockButton.Text = "Dock"
            End If
            '
            Me.F1.Visible = True
            Refresh_Général()
            Maj_Tooltips()
        Else                            ' --> ATTACHER
            Me.F1.Visible = False
            Me.F1.FormBorderStyle = FormBorderStyle.None ' attacher
            Me.F1.TopMost = False   ' un seul des 2 suffit ?
            Me.F1.TopLevel = False
            Form1.TabControl4.TabPages.Item(NumOnglet).Controls.Add(Me.F1)
            F1.Text = Nothing
            Me.F1.MainMenuStrip.Visible = False
            Me.F1.MainMenuStrip.Enabled = False
            'Menu1.Visible = False
            'Menu1.Enabled = False
            'AnnulationHandlerEdition()
            Me.F1.Dock = DockStyle.Fill

            '
            If LangueIHM = "fr" Then
                DockButton.Text = "Détacher"
            Else
                DockButton.Text = "UnDock"
            End If
            Me.F1.Visible = True
            '
            If AffCtrls.Checked Then
                Show_Ctrls()
            Else
                Hide_Ctrls()
            End If
            '
            Refresh_Général()
            Maj_Tooltips()
        End If
    End Sub
    ''' <summary>
    ''' Attacher : attacher le formulaire. Sert au menu Fichier/Quitter du formulaire
    ''' </summary>
    Sub Attacher(NumOnglet As Integer)
        Me.F1.Visible = False
        Me.F1.FormBorderStyle = FormBorderStyle.None ' attacher
        Me.F1.TopMost = False   ' un seul des 2 suffit ?
        Me.F1.TopLevel = False
        Form1.TabControl4.TabPages.Item(NumOnglet).Controls.Add(Me.F1)
        F1.Text = Nothing
        Me.F1.MainMenuStrip.Visible = False
        Me.F1.MainMenuStrip.Enabled = False
        'Menu1.Visible = False
        'Menu1.Enabled = False
        'AnnulationHandlerEdition()
        Me.F1.Dock = DockStyle.Fill
        '
        If LangueIHM = "fr" Then
            DockButton.Text = "Détacher"
        Else
            DockButton.Text = "UnDock"
        End If
        Me.F1.Visible = True
        Refresh_Général()
        Maj_Tooltips()
    End Sub
    Sub Refresh_Général()
        Dim i As Integer
        For i = 0 To GridCourbes.Count - 1
            GridCourbes.Item(i).Refresh()
        Next
    End Sub
    Private Sub Opacité_ValueChange(sender As Object, e As EventArgs)
        Me.F1.Opacity = Convert.ToDouble(Opacité.Value / 100)
    End Sub
    Private Sub Destination_Keydown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True ' pour supprmer le son sur 'Enter' - ne fonctionne que sur KeyDown
            Aller_Vers()
        End If
    End Sub
    Private Sub Destination_TextChanged(sender As Object, e As EventArgs)
        If Trim(Destination.Text) <> "" And IsNumeric(Trim(Destination.Text)) Then
            If CInt(Destination.Text) <= CInt(nbMesures) Then
                Destination.BackColor = Color.White
                Destination.ForeColor = Color.Black
            Else
                Destination.BackColor = Color.Red
                Destination.ForeColor = Color.White
            End If
        End If
    End Sub
    Sub Aller_Vers()
        If Trim(Destination.Text) <> "" And IsNumeric(Trim(Destination.Text)) Then
            Dim i As Integer = Convert.ToUInt16(Trim(Destination.Text))
            If i >= 1 And i <= nbMesures Then
                Grid1.LeftCol = ((i - 1) * 16) + 1
                If CheckTop.Checked Then
                    Grid1.TopRow = Grid1.FixedRows
                    Grid1.Range(Grid1.TopRow, Grid1.LeftCol, Grid1.TopRow, Grid1.LeftCol).SelectCells()
                End If
            End If

            '
            Aller_VersNote(CInt(Destination.Text))
        End If
    End Sub
    Sub Aller_VersNote(N_Mesure As Integer)
        Dim i As Integer
        Dim j As Integer = 1
        Dim N_Note1 = 0
        Dim N_Mesure1 As Integer = ((N_Mesure - 1) * 16) + 1
        '
        Dim OK As Boolean = False
        If N_Mesure <= nbMesures Then
            Destination.BackColor = Color.White
            Destination.ForeColor = Color.Black
            Do
                ' Calcul sur une colonne
                ' **********************
                For i = 6 To Grid1.Rows - 1
                    For j = N_Mesure1 To N_Mesure1 + 15
                        If Trim(Grid1.Cell(i, j).Text) <> "" Then
                            OK = True
                            Grid1.Cell(i + 5, N_Mesure1).SetFocus() ' affichage de la mesure avec adaptation ligne(i+5) et colonne (N_Mesure1)
                            Exit For
                        End If
                    Next j
                    If OK = True Then Exit For
                Next i
            Loop Until OK = True Or j >= N_Mesure1 + 15
        Else
            Destination.BackColor = Color.Red
            Destination.ForeColor = Color.White
        End If
    End Sub


    Public Sub F1_Load(sender As Object, e As EventArgs)
        Grid1.TopRow = 64
        Grid1.Range(5, 1, 5, 1)
        '
    End Sub
    Private Sub F1_Activated(sender As Object, e As EventArgs)
        'F1.Opacity = 1
        'Panneau1.Visible = True
    End Sub
    Private Sub F1_Deactivate(sender As Object, e As EventArgs)

        'F1.Opacity = 0.6
        'Panneau1.Visible = False

    End Sub
    Private Sub F1_resize(sender As Object, e As EventArgs)
        If Me.F1.Dock = DockStyle.None And (Not EnChargement) And (AffCtrls.Checked) Then
            Panneau2.IsSplitterFixed = False
            If Panneau2.Height - 195 > 0 Then
                Panneau2.SplitterDistance = Panneau2.Height - 195
            End If
            Panneau2.IsSplitterFixed = True

        End If
    End Sub
    Private Sub F1_KeyUp(sender As Object, e As KeyEventArgs)
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
        '
    End Sub

    Sub BRadio1_MouseDown(sender As Object, e As EventArgs)
        Send_Pano(Trim(BRadio1.Text))
    End Sub
    Sub BRadio2_MouseDown(sender As Object, e As EventArgs)
        Send_Pano(Trim(BRadio2.Text))
    End Sub
    Sub BRadio3_MouseDown(sender As Object, e As EventArgs)
        Send_Pano(Trim(BRadio3.Text))
    End Sub
    Sub Check12_8_CheckedChanged(sender As Object, e As EventArgs)
        If Check12_8.Checked Then
            Grid1.AutoRedraw = False
            '
            Clear_graphique()
            Graphique_Gammes()

            Graphique_Divisions()
            If CheckAcc.Checked Then
                Graphique_Accords()
            End If
            If CheckDrum.Checked Then
                Graphique_Drums()
            End If
            '
            Graphique_12_8()
            Grid1.AutoRedraw = True
            Grid1.Refresh()
        Else
            Grid1.AutoRedraw = False
            '
            Clear_graphique()
            Graphique_Gammes()
            If CheckAcc.Checked Then
                Graphique_Accords()
            End If
            Graphique_Divisions()
            If CheckDrum.Checked Then
                Graphique_Drums()
            End If
            '
            Grid1.AutoRedraw = True
            Grid1.Refresh()
        End If
    End Sub
    Sub AffCtrls_CheckedChanged(sender As Object, e As EventArgs)
        If AffCtrls.Checked Then
            Show_Ctrls()
        Else
            Hide_Ctrls()
        End If
    End Sub
    Sub ZoomMoins_CheckedChanged(sender As Object, e As EventArgs)
        Dim i, j As Integer
        Grid1.AutoRedraw = False
        '
        If ZoomMoins.Checked Then
            For i = 0 To Grid1.Rows - 1
                Grid1.Row(i).Height = ValZoomMoins
            Next
            '
            For i = 1 To Grid1.Cols - 1
                Grid1.Column(i).Width = ValZoomMoins
            Next i
            '
            For i = 0 To nbCourbes - 1
                GridCourbes.Item(i).AutoRedraw = False
                For j = 1 To GridCourbes.Item(i).Cols - 1
                    GridCourbes.Item(i).Column(j).Width = ValZoomMoins
                Next j
                GridCourbes.Item(i).AutoRedraw = True
                GridCourbes.Item(i).Refresh()
            Next i
        Else
            For i = 0 To Grid1.Rows - 1
                Grid1.Row(i).Height = ValZoomPlus
            Next
            '
            For i = 1 To Grid1.Cols - 1
                Grid1.Column(i).Width = ValZoomPlus
            Next i
            '
            For i = 0 To nbCourbes - 1
                GridCourbes.Item(i).AutoRedraw = False
                For j = 1 To GridCourbes.Item(i).Cols - 1
                    GridCourbes.Item(i).Column(j).Width = ValZoomPlus
                Next j
                GridCourbes.Item(i).AutoRedraw = True
                GridCourbes.Item(i).Refresh()
            Next i
        End If
        '

        Grid1.AutoRedraw = True
        Grid1.Refresh()
    End Sub





    Sub Checkacc_CheckedChanged(sender As Object, e As EventArgs)
        If CheckAcc.Checked Then
            '
            Grid1.AutoRedraw = False
            '
            Clear_graphique()
            Graphique_Gammes()

            Graphique_Accords()
            Graphique_Divisions()
            If Check12_8.Checked Then
                Graphique_12_8()
            End If
            If CheckDrum.Checked Then
                Graphique_Drums()
            End If
            '
            Grid1.AutoRedraw = True
            Grid1.Refresh()
        Else
            Grid1.AutoRedraw = False
            '
            Clear_graphique()
            Graphique_Gammes()
            Graphique_Divisions()

            If Check12_8.Checked Then
                Graphique_12_8()
            End If
            If CheckDrum.Checked Then
                Graphique_Drums()
            End If
            '
            Grid1.AutoRedraw = True
            Grid1.Refresh()
        End If
    End Sub
    Sub CheckDrum_CheckedChanged(sender As Object, e As EventArgs)
        If CheckDrum.Checked Then

            Grid1.AutoRedraw = False
            '
            Clear_graphique()
            Graphique_Gammes()

            Graphique_Divisions()
            If CheckAcc.Checked Then
                Graphique_Accords()
            End If
            If Check12_8.Checked Then
                Graphique_12_8()
            End If
            Graphique_Drums()
            '
            Grid1.AutoRedraw = True
            Grid1.Refresh()
        Else
            Grid1.AutoRedraw = False
            '
            Clear_graphique()
            Graphique_Gammes()
            If CheckAcc.Checked Then
                Graphique_Accords()
            End If
            Graphique_Divisions()
            '
            If Check12_8.Checked Then
                Graphique_12_8()
            End If
            Grid1.AutoRedraw = True
            Grid1.Refresh()
        End If
    End Sub
    Sub ListPRG_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim com As Windows.Forms.ComboBox = sender
        Dim PRG As Integer
        '
        PRG = ListPRG.SelectedIndex - 1
        If Chargé = True Then
            If Trim(PRG) <> -1 Then
                If Not (Form1.SortieMidi.Item(Form1.ChoixSortieMidi).IsOpen) Then
                    Form1.SortieMidi.Item(Form1.ChoixSortieMidi).Open()
                End If
                'tbl = Split(Trim(PRG))
                Form1.SortieMidi.Item(Form1.ChoixSortieMidi).SendProgramChange(Me.Canal, PRG)
            End If
        End If
    End Sub
    Sub ListdynF1_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim i, j As Integer
        For i = Grid1.Selection.FirstRow To Grid1.Selection.LastRow
            For j = Grid1.Selection.FirstCol To Grid1.Selection.LastCol
                If IsNumeric(Grid1.Cell(i, j).Text) Then
                    Grid1.Cell(i, j).Text = Trim(ListDynF1.Text)
                    'Form1.ReCalcul()
                End If
            Next
        Next
    End Sub
    Sub ListMod_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim ii As Integer = 5
        Dim jj As Integer = Grid1.Selection.FirstCol

        If Not EnChargement Then
            If Grid1.Selection.FirstRow = Grid1.FixedRows And Grid1.Selection.LastRow = Grid1.FixedRows + 127 Then
                'Grid1.Cell(ii, jj).Text = Trim(ListMod.Text)
            End If
        End If

    End Sub

    Private Sub Grid1_KeyPress(Sender As Object, e As KeyPressEventArgs)

    End Sub

    Private Sub Grid1_MouseUp(Sender As Object, e As MouseEventArgs)
        Dim i, j As Integer
        Dim ActC As Cell
        Dim a As String
        Dim a1 As String
        Dim a2 As String

        Dim aa As New Ann

        Dim ii As Integer = Grid1.MouseRow
        Dim jj As Integer = Grid1.MouseCol
        Dim tbl() As String
        Dim tbl1() As String



        ' Ecriture tête de Note
        ' ********************
        If Orig_PianoR.Orig1 = OrigPianoCourbe.Piano Then
            ListAnnulation.Clear()
            PointAnn = -1
            '
            If My.Computer.Keyboard.CtrlKeyDown And (Not EnRecalcul) And ii >= Grid1.FixedRows Then
                If IsNumeric(Grid1.ActiveCell.Text) = False Then
                    ' écriture note
                    ActC = Grid1.ActiveCell ' sauvegarde pour gestion ctrl+z (à la fin de l'ecriture de la  note)
                    Grid1.ActiveCell.Text = Trim(ListDynF1.Text) ' tête de note
                    Grid1.Cell(Grid1.ActiveCell.Row, Grid1.ActiveCell.Col + i).FontBold = True
                    '
                    ' longueur de la note
                    j = NbDiv(ListTypNote.Text)
                    For i = 1 To j - 1
                        If Grid1.ActiveCell.Col + i <= Grid1.Cols - 1 Then
                            If IsNumeric(Grid1.Cell(Grid1.ActiveCell.Row, Grid1.ActiveCell.Col + i).Text) Then
                                Exit For
                            End If
                        Else
                            Exit For
                        End If
                        Grid1.Cell(Grid1.ActiveCell.Row, Grid1.ActiveCell.Col + i).Text = Trait ' longueur de note
                        Grid1.Cell(Grid1.ActiveCell.Row, Grid1.ActiveCell.Col + i).FontBold = True
                    Next
                    '
                    ' gestion ctrl+z
                    ListAnnulation.Add(aa)
                    PointAnn = PointAnn + 1
                    ListAnnulation.Item(PointAnn).Action = ActionEnum.Effacer
                    ListAnnulation.Item(PointAnn).Tinfo = TypeInfo.Notes
                    Dim oo2 As New Ann.SauvAnnuler With {
                        .Ligne = ActC.Row, 'Grid1.ActiveCell.Row,
                        .Colonne = ActC.Col 'Grid1.ActiveCell.Col
                    }
                    oo2.Longueur = Det_LongueurNote(oo2.Ligne, oo2.Colonne)
                    oo2.Vélo = Grid1.ActiveCell.Text
                    ListAnnulation.Item(PointAnn).ListAnnuler.Add(oo2)
                Else
                    ' Effacer Note
                    ' gestion Ctrl+Z
                    ListAnnulation.Add(aa)
                    PointAnn = PointAnn + 1
                    ListAnnulation.Item(PointAnn).Action = ActionEnum.Restituer
                    ListAnnulation.Item(PointAnn).Tinfo = TypeInfo.Notes
                    Dim oo2 As New Ann.SauvAnnuler With {
                        .Ligne = Grid1.ActiveCell.Row,
                        .Colonne = Grid1.ActiveCell.Col
                    }
                    oo2.Longueur = Det_LongueurNote(oo2.Ligne, oo2.Colonne)
                    oo2.Vélo = Grid1.ActiveCell.Text
                    ListAnnulation.Item(PointAnn).ListAnnuler.Add(oo2)
                    ' Raccourcir et Effacer (Effacer => tête de note = "" et  raccourcir à partir du début+1)
                    Grid1.ActiveCell.Text = "" ' effacer tête de note
                    If Grid1.ActiveCell.Col < Grid1.Cols - 1 Then
                        Raccourcir(Grid1.ActiveCell.Row, Grid1.ActiveCell.Col + 1) ' effacer la longueur de la note
                    End If
                End If

                'Form1.ReCalcul()
            End If
            '
            ' Modification longueur de note
            ' *****************************
            If My.Computer.Keyboard.AltKeyDown Then
                ' Modifier longueur de note (pas d'annulation pour une longueur de note)
                ' **********************************************************************
                If Trim(Grid1.ActiveCell.Text) = "" Then 'rallonger
                    ' gestion du ctrl+Z
                    ListAnnulation.Add(aa)
                    PointAnn = PointAnn + 1
                    ListAnnulation.Item(PointAnn).Action = ActionEnum.Restituer
                    ListAnnulation.Item(PointAnn).Tinfo = TypeInfo.NotesRallg
                    Dim oo1 As New Ann.SauvAnnuler With {
                        .Ligne = Grid1.ActiveCell.Row,
                        .Colonne = Det_TêtedeNote(.Ligne, Grid1.ActiveCell.Col)'Grid1.ActiveCell.Col
                    }
                    oo1.Longueur = Det_LongueurNote(oo1.Ligne, oo1.Colonne)
                    Dim z As Integer = Det_TêtedeNote(oo1.Ligne, oo1.Colonne)
                    oo1.Vélo = Trim(Grid1.Cell(oo1.Ligne, z).Text)
                    ListAnnulation.Item(PointAnn).ListAnnuler.Add(oo1)
                    ' rallonger
                    Rallonger(Grid1.ActiveCell.Row, Grid1.ActiveCell.Col)
                Else
                    If Trim(Grid1.ActiveCell.Text) = Trait Then 'raccourcir
                        ' gestion du ctrl+Z
                        ListAnnulation.Add(aa)
                        PointAnn = PointAnn + 1
                        ListAnnulation.Item(PointAnn).Action = ActionEnum.Restituer
                        ListAnnulation.Item(PointAnn).Tinfo = TypeInfo.Notes
                        Dim oo2 As New Ann.SauvAnnuler With {
                        .Ligne = Grid1.ActiveCell.Row,
                        .Colonne = Grid1.ActiveCell.Col
                        }
                        ' raccourcir
                        oo2.Longueur = Det_LongueurNote(oo2.Ligne, oo2.Colonne)
                        oo2.Vélo = Grid1.ActiveCell.Text
                        ListAnnulation.Item(PointAnn).ListAnnuler.Add(oo2)
                        ' raccourcir
                        Raccourcir(Grid1.ActiveCell.Row, Grid1.ActiveCell.Col)
                    End If
                End If
            End If
            '
            ' Coupure de note jouée avec la souris
            If NoteAEtéJouée = True Then
                StoperNote(NoteCourante, 0)
                NoteAEtéJouée = False
            End If
            '
            ' Ecriture/Effacement pédale
            ' **************************
            ' Ecriture P+ d'une pédale et effacement
            If ii = Grid1.FixedRows - 1 Then
                If PointAnn = -1 Then
                    ListAnnulation.Add(aa)
                    PointAnn = PointAnn + 1
                End If
                ListAnnulation.Item(PointAnn).Tinfo = TypeInfo.Pedales

                If My.Computer.Keyboard.CtrlKeyDown And (Not EnRecalcul) Then
                    Dim ooo As New Ann.SauvAnnulerCtrl With {
                    .Colonne = Grid1.ActiveCell.Col,
                    .Valeur = Trim(Grid1.Cell(Grid1.FixedRows - 1, Grid1.ActiveCell.Col).Text)
                        }
                    ListAnnulation.Item(PointAnn).ListAnnulerCTRL.Add(ooo)
                    If Trim(Grid1.Cell(Grid1.FixedRows - 1, Grid1.ActiveCell.Col).Text) <> "" Then
                        Grid1.Cell(Grid1.FixedRows - 1, Grid1.ActiveCell.Col).Text = "" ' effacer modulation
                        ListAnnulation.Item(PointAnn).Action = ActionEnum.Restituer
                    Else
                        Grid1.Cell(Grid1.FixedRows - 1, Grid1.ActiveCell.Col).Text = "P+" ' écrire modulation
                        ListAnnulation.Item(PointAnn).Action = ActionEnum.Effacer
                    End If
                End If
                ' Ecriture P- d'une pédale et effacement
                If My.Computer.Keyboard.AltKeyDown And (Not EnRecalcul) Then
                    Dim ooo As New Ann.SauvAnnulerCtrl With {
                    .Colonne = Grid1.ActiveCell.Col,
                    .Valeur = Trim(Grid1.Cell(Grid1.FixedRows - 1, Grid1.ActiveCell.Col).Text)
                        }
                    ListAnnulation.Item(PointAnn).ListAnnulerCTRL.Add(ooo)
                    If Trim(Grid1.Cell(Grid1.FixedRows - 1, Grid1.ActiveCell.Col).Text) <> "" Then
                        Grid1.Cell(Grid1.FixedRows - 1, Grid1.ActiveCell.Col).Text = "" ' effacer modulation
                        ListAnnulation.Item(PointAnn).Action = ActionEnum.Restituer
                    Else
                        Grid1.Cell(Grid1.FixedRows - 1, Grid1.ActiveCell.Col).Text = "P-" ' écrire modulation
                        ListAnnulation.Item(PointAnn).Action = ActionEnum.Effacer
                    End If
                End If
                '
                ' gestion annulation


            End If
            '
        End If
        ' ********************************************************
        ' TRAITEMENTS DES ASSISTANCES (GridAssist1, GridAssist2) *
        ' ********************************************************
        '
        ' *******************************************
        ' I - Détermination des notes sélectionnées *
        ' *******************************************

        Dim Fr As Integer = Grid1.Selection.FirstRow
        Dim Fc As Integer = Grid1.Selection.FirstCol
        Dim Lr As Integer = Grid1.Selection.LastRow
        Dim Lc As Integer = Grid1.Selection.LastCol
        '
        Fisrtc = Fc ' variable globale utiliser par l'écriture
        LastC = Lc
        '
        Dim L1 As New List(Of String)
        Dim L2 As New List(Of String)
        Dim TabN As New List(Of String) From {"c", "c#", "d", "d#", "e", "f", "f#", "g", "g#", "a", "a#", "b", "c", "c#", "d", "d#", "e", "f", "f#", "g", "g#", "a", "a#", "b"}
        '
        '
        ' Mémorisation du N° de mesure de la 1ere note
        ' ********************************************
        Dim dd As Decimal = Fc / 16
        dd = Math.Truncate(dd) ' Truncate = partie entière d'un nombre
        MesureEcriture = Convert.ToInt16(dd + 1)

        ' Détermination des notes de la sélection
        ' ***************************************
        For j = Fc To Lc
            For i = Fr To Lr
                a = Trim(Grid1.Cell(i, j).Text)
                If IsNumeric(a) Then
                    L1.Add(Trim(Grid1.Cell(i, 0).Text))
                End If
            Next
        Next
        ' Construction de la liste des notes (sans les n° et en minuscules)
        ' *****************************************************************
        If L1.Count > 0 Then '2 Then
            For Each oo As String In L1
                j = Len(oo)
                If IsNumeric(oo(j - 1)) Then
                    a = Microsoft.VisualBasic.Left(oo, j - 1)
                    a = a.ToLower()
                    L2.Add(a)
                End If
            Next
            ListeDesNotes = L2 ' ListeDesNotes -->destiné à Onglet Consonnaces
            L2 = L2.Distinct().ToList() ' suppression des doublons dans la liste
            '
            ' **************************************************
            ' II - Traitement la déterminations des gammes     *
            ' **************************************************
            '
            ' Détermination des gammes possibles
            ' **********************************
            Dim L3 As List(Of String)
            L3 = Form1.Assist_Tonalité2(L2)
            '
            ' Mise à jour du tableau d'assistance Assist2Grid
            ' ***********************************************
            i = 0
            '
            GridAssist2.Range(1, 1, GridAssist2.Rows - 1, GridAssist2.Cols - 1).ClearAll()
            '
            For Each bb As String In L3
                i += 1
                bb = Trad_TonaDiese(Trim(bb))
                tbl = bb.Split()
                Select Case Trim(tbl(1)) ' tbl(1) contient le chiffrage de la gamme
                    Case "Maj"
                        GridAssist2.Cell(i, 1).Text = Trim(bb) ' bb contient la gamme
                        a = Det_NotesGammes(bb)
                        GridAssist2.Cell(i, 2).Text = Trim(a.ToLower) ' a contient les notes de la gamme en minuscules
                        GridAssist2.Cell(i, 3).Text = Trim(bb)
                        GridAssist2.Cell(i, 4).Text = Trim(tbl(0))

                    Case "MinH", "MinM", "PentaMin"
                        tbl1 = bb.Split()
                        'bb = Trim(Trad_TonaDiese(bb))
                        GridAssist2.Cell(i, 1).Text = Trim(bb) 'ecriture du nom de la gamme
                        '
                        a = Det_NotesGammes(bb)
                        GridAssist2.Cell(i, 2).Text = Trim(a.ToLower) ' écriture des notes de la gamme
                        '
                        j = TabN.IndexOf(tbl(0).ToLower)
                        a = TabN(j + 3) ' j+3 donne la tonalité Majeur
                        a = Trim(a.ToUpper + " " + "Maj")
                        a = Trad_TonaDiese(a)
                        GridAssist2.Cell(i, 3).Text = Trim(a) ' écriture tonalitéTrim(a.ToUpper + " " + "Maj")
                        GridAssist2.Cell(i, 4).Text = Trim(tbl(0)) + " " + "m"

                    Case "Blues", "Blues2"
                        GridAssist2.Cell(i, 1).Text = Trim(bb)
                        a = Det_NotesGammes(bb)
                        GridAssist2.Cell(i, 2).Text = Trim(a.ToLower)
                        j = TabN.IndexOf(tbl(0).ToLower)
                        '
                        a = TabN(j + 5)
                        a = Trim(a.ToUpper + " " + "Maj")
                        a = Trad_TonaDiese(a)
                        '
                        a1 = tbl(0) + " " + "7"
                        a2 = tbl(0) + " " + "m7"
                        GridAssist2.Cell(i, 3).Text = Trim(a) 'Trim(a.ToUpper + " " + "Maj")
                        GridAssist2.Cell(i, 4).Text = Trim(a1)
                End Select
            Next
            '
            ' ********************************************
            ' III - Ecriture du CF pour les consonnances *
            ' ********************************************
            ' 
            Raz_Assist1()
            i = 1
            For Each aaa As String In ListeDesNotes
                GridAssist1.Cell(i, Col_CF).Text = Trim(aaa)
                Calc_Consonnances(Trim(GridAssist1.Cell(i, Col_CF).Text), i)
                i += 1
            Next
        End If
    End Sub
    Sub Raz_Assist1()
        Dim i, j As Integer

        For i = 1 To GridAssist1.Rows - 1
            For j = 1 To GridAssist1.Cols - 1
                GridAssist1.Cell(i, j).Text = ""
            Next
        Next
    End Sub
    Sub Calc_Consonnances(Note As String, r As Integer)
        Dim i As Integer = -1

        With GridAssist1
            Do
                i += 1
            Loop Until (i = Chromad.Count - 1) Or (Note = Chromad.Item(i))
            '
            If i < Chromad.Count - 1 Then
                ' Voix haute
                .Cell(r, Col_CF + 12).Text = Chromad(i + Octave)
                .Cell(r, Col_CF + 11).Text = Chromad(i + Sept)
                .Cell(r, Col_CF + 10).Text = Chromad(i + Septb)
                .Cell(r, Col_CF + 9).Text = Chromad(i + SixteMaj)
                .Cell(r, Col_CF + 8).Text = Chromad(i + SixteMin)
                .Cell(r, Col_CF + 7).Text = Chromad(i + Quinte)
                .Cell(r, Col_CF + 6).Text = Chromad(i + Quinteb)
                .Cell(r, Col_CF + 5).Text = Chromad(i + Quarte)
                .Cell(r, Col_CF + 4).Text = Chromad(i + TierceMaj)
                .Cell(r, Col_CF + 3).Text = Chromad(i + TierceMin)
                .Cell(r, Col_CF + 2).Text = Chromad(i + Neuf)
                .Cell(r, Col_CF + 1).Text = Chromad(i + Neufb)
                '.Cell(r, Col_CF).Text = Chromad(i + Octave)
                ' Voix Basse
                '.Cell(r, 13).Text = Chromad(i + EstOctave)
                .Cell(r, 12).Text = Chromad(i + EstNeufb)
                .Cell(r, 11).Text = Chromad(i + EstNeuf)
                .Cell(r, 10).Text = Chromad(i + EstTierceMin)
                .Cell(r, 9).Text = Chromad(i + EstTierceMaj)
                .Cell(r, 8).Text = Chromad(i + EstQuarte)
                .Cell(r, 7).Text = Chromad(i + EstQuinteb)
                .Cell(r, 6).Text = Chromad(i + EstQuinte)
                .Cell(r, 5).Text = Chromad(i + EstSixteMin)
                .Cell(r, 4).Text = Chromad(i + EstSixteMaj)
                .Cell(r, 3).Text = Chromad(i + EstSeptb)
                .Cell(r, 2).Text = Chromad(i + EstSept)
                .Cell(r, 1).Text = Chromad(i + EstOctave)
            Else
                .Cell(r, Col_CF).Text = ""
            End If
        End With
    End Sub
    Function Trad_TonaDiese(bb As String) As String
        Dim tbl() As String = bb.Split()
        Dim a As String

        Trad_TonaDiese = bb
        If Trim(tbl(1)) = "Maj" Then
            Select Case tbl(0)
                Case "A#"
                    a = "Bb"
                Case "D#"
                    a = "Eb"
                Case "G#"
                    a = "Ab"
                Case Else
                    a = Trim(tbl(0))
            End Select
            Return a + " " + tbl(1)
        End If

    End Function



    Private Sub Grid1_MouseDown(Sender As Object, e As EventArgs)
        ' Jouer la note sur simple clic
        Dim ii As Integer = Grid1.MouseRow
        Dim jj As Integer = Grid1.MouseCol
        '
        Orig_PianoR.Orig1 = OrigPianoCourbe.Piano
        Orig_PianoR.N_Courbe = -1

        If ii > 5 And jj > -1 Then ' cas où MouseRow ou MouseCol n'arrivent pas  à temps, ils valent -1 ce qui n'est pas utilisable

            If Not EnRecalcul Then
                Dim ValeurNote = ValNoteCubase.IndexOf(Grid1.Cell(ii, 0).Text)
                If IsNumeric(Grid1.Cell(ii, jj).Text) Then
                    JouerNote(ValeurNote, Convert.ToByte(Grid1.Cell(ii, jj).Text))
                Else
                    JouerNote(ValeurNote, Convert.ToByte(ListDynF1.Text))
                End If
            End If
        End If

    End Sub
    Private Sub TimerStop_Tick()
        TimerStop.Stop()
        StoperNote(NoteCourante, 0)
    End Sub
    '
    Private Sub Grid1_MouseDown2(Sender As Object, e As MouseEventArgs)
        Dim i As Integer = Grid1.MouseRow 'Grid1.ActiveCell.Row 'Grid1.MouseRow
        Dim j As Integer = Grid1.MouseCol 'Grid1.ActiveCell.Col 'Grid1.MouseCol
        Dim ValeurNote As String
        Dim ValeurNote2 As String
        '
        Dim ii As Integer = i
        Dim jj As Integer = j

        Orig_PianoR.Orig1 = OrigPianoCourbe.Piano
        Orig_PianoR.N_Courbe = -1

        If i > Grid1.FixedRows - 1 Then
            If IsNumeric(Grid1.Cell(i, j).Text) Or Trim(Grid1.Cell(i, j).Text) = Trait Or My.Computer.Keyboard.CtrlKeyDown Then
                While Trim(Grid1.Cell(i, j).Text) = Trait And j <> 0
                    j = j - 1
                End While
                ValeurNote = ValNoteCubase.IndexOf(Grid1.Cell(ii, 0).Text)
                If IsNumeric(Grid1.Cell(i, j).Text) Then
                    JouerNote(ValeurNote, Convert.ToByte(Grid1.Cell(i, j).Text))
                Else
                    JouerNote(ValeurNote, Convert.ToByte(ListDynF1.Text))
                End If
            Else
                If ii > Grid1.FixedRows Then
                    ValeurNote2 = ValNoteCubase.IndexOf(Grid1.Cell(ii, 0).Text)
                    JouerNote(ValeurNote2, Convert.ToByte(ListDynF1.Text))
                End If
            End If
        End If
        '
        ' Menu contextuel pour vélocité aléatoires
        ' ****************************************
        If e.Button = MouseButtons.Right Then
            Dim p As New Point
            p.X = Cursor.Position.X
            p.Y = Cursor.Position.Y
            'MenuContext2.Show(p)
            itemVélo_Click()

        End If
    End Sub
    Private Sub Grid1_Keydown(Sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Dim i As Integer '= Grid1.ActiveCell.Row
        Dim j As Integer '= Grid1.ActiveCell.Col
        Dim ii As Integer
        Dim k As Integer
        Dim a As String ' 
        Dim b As String
        Dim tbl1() As String
        Dim tbl2() As String
        Dim sortir As Boolean = False

        If Orig_PianoR.Orig1 = OrigPianoCourbe.Piano Then
            If e.KeyCode = Keys.Add Or e.KeyCode = Keys.Subtract Then
                b = CellDyn() ' valeur de la cellule ou de la sélection de ceulles
                '
                ' Vérification préalable de dépassement des bornes min/max selon touche -/+ 
                ' *************************************************************************
                If Trim(b) <> "" Then
                    tbl1 = Split(b)
                    For ii = 0 To UBound(tbl1)
                        tbl2 = Split(tbl1(ii), ",")
                        i = Convert.ToInt16(tbl2(0)) ' ligne
                        j = Convert.ToInt16(tbl2(1)) ' col
                        k = Convert.ToInt16(Grid1.Cell(i, j).Text) ' lecture de la valeur actuelle de la cellule
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
                    Next
                    ' Traitement incrémentation/décrémentation (si sortir = false)
                    ' ***********************************************************
                    If sortir = False Then
                        For ii = 0 To UBound(tbl1)
                            tbl2 = Split(tbl1(ii), ",")
                            i = Convert.ToInt16(tbl2(0)) ' ligne
                            j = Convert.ToInt16(tbl2(1)) ' colonne 
                            a = Trim(Grid1.Cell(i, j).Text) ' valeur de la dynamqie
                            '
                            ' incrémentation
                            If e.KeyCode = Keys.Add Then
                                k = Convert.ToInt16(a) + 1
                                If k <= 127 Then
                                    Grid1.Cell(i, j).Text = Convert.ToString(k)
                                End If
                                ' décrémentation
                            ElseIf e.KeyCode = Keys.Subtract Then
                                k = Convert.ToInt16(a) - 1
                                If k >= 0 Then
                                    Grid1.Cell(i, j).Text = Convert.ToString(k)
                                End If
                            End If
                        Next
                    End If
                End If
            End If
        End If
        '
        '
    End Sub
    ''' <summary>
    ''' L'objectif de cette fonction est savoir si les touches - et +
    ''' du pavé doivent modifier une valeur de modulation et non pas une valeur 
    ''' de controleur(La difficulté vient du fait que les valeurs de la molette de 
    ''' modulation sont écrite sur une ligne fixe).
    ''' Les conditions sont :
    '''  - la sélection courante est constituée des 1ere  et dernière ligne non fixe
    '''  - il existe une valeur numérique de Modulation à modifier
    ''' </summary>
    ''' <param name="FRow">1ere ligne non fixe</param>
    ''' <param name="LastRow">Dernière ligne non fixe</param>
    ''' <param name="ColCours">Colonne Courante</param>
    ''' <returns>Coordonnée de la cellule à modifier. Si n'existe pas alors Return = ""</returns>
    Function ModifierModulation(FRow As Integer, LastRow As Integer, ColCours As Integer) As String

        ModifierModulation = ""
        If FRow = Grid1.FixedRows And LastRow = Grid1.Rows - 1 Then
            If IsNumeric(Grid1.Cell(FRow - 1, ColCours).Text) Then
                ModifierModulation = (FRow - 1).ToString + "," + ColCours.ToString
            End If
        End If
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
        a = ModifierModulation(Grid1.Selection.FirstRow, Grid1.Selection.LastRow, Grid1.Selection.FirstCol)
        CellDyn = Trim(a)
        If Trim(a) = "" Then
            For i = Grid1.Selection.FirstRow To Grid1.Selection.LastRow
                For j = Grid1.Selection.FirstCol To Grid1.Selection.LastCol
                    If IsNumeric(Grid1.Cell(i, j).Text) Then
                        a = a + Convert.ToString(i) + "," + Convert.ToString(j) + " "
                    End If
                Next
            Next
            CellDyn = Trim(a)
        End If
    End Function
    Private Sub Grid1_KeyUp(Sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Dim pas As Integer
        Dim iii As Integer
        Dim NumOnglet As Integer = Val(F1.Tag)

        iii = NumOnglet - 1
        If NumOnglet > 3 Then iii = NumOnglet - 2

        If e.KeyCode = Keys.Delete And Orig_PianoR.Orig1 = OrigPianoCourbe.Piano Then
            Form1.listPIANOROLL.Item(iii).EffacerSelection()
        End If

        '
        ' ***********************************************************
        ' * Appel à la fonction de quatification mélodique de notes *
        ' ***********************************************************
        If e.KeyCode = Keys.H Then
            pas = -1
        ElseIf e.KeyCode = Keys.B Then
            pas = 1
        End If
        If pas <> 0 Then
            QuantMel(pas)
        End If
    End Sub
    Sub Grid1_Scroll(sender As Object, a As EventArgs)

        For i = 0 To nbCourbes - 1
            GridCourbes.Item(i).LeftCol = Grid1.LeftCol
        Next

    End Sub
    Private Sub BoutonF1_1_Click(sender As Object, e As EventArgs)
        If Form1.WindowState <> FormWindowState.Minimized Then
            F1.Top = Form1.Top + 78
            F1.Left = Form1.Left
            F1.Width = Form1.Width - 612
            F1.Height = Form1.Height - 78
        End If
    End Sub
    Private Sub BoutonF1_2_Click(sender As Object, e As EventArgs)
        If Form1.WindowState <> FormWindowState.Minimized Then
            F1.Top = Form1.Top + 78
            F1.Left = Form1.Left
            F1.Width = Form1.Width - 5
            F1.Height = Form1.Height - 78
        End If
    End Sub
    Private Sub BoutonFermer_Click(sender As Object, e As EventArgs)
        F1.Hide()
        Me.Chargé = False
        'Form1.PPIANOROLL_chargé = False
    End Sub
    Sub OuvrirCalques_Click(sender As Object, e As EventArgs)
        '
        Dim i As Integer


        For i = 0 To PassChoixCalques.Count - 1
            PassChoixCalques(i) = ChoixCalquesLocal(i)
        Next i
        PassChoixPédale = PédaleLocale
        Retour = ""
        CalquesMIDI.StartPosition = FormStartPosition.CenterScreen
        CalquesMIDI.ShowDialog()
        If Retour = "OK" Then
            For i = 0 To PassChoixCalques.Count - 1
                ChoixCalquesLocal(i) = PassChoixCalques(i)
            Next i
        End If
        '
        PédaleLocale = PassChoixPédale
        TessDebLocale = PassTessDeb
        TessFinLocale = PassTessFin
        TessListeLocale = PassTessListe
        NouvCalqueLocale = PassNouvCalque
        '
        Maj_CalquesMIDI()
    End Sub
    Sub Maj_CalquesMIDI()
        Dim a As String = ""
        Dim deb As Integer = 0
        Dim fin = ChoixCalquesLocal.Count - 1

        Grid1.AutoRedraw = False

        Refresh_G1_Acc()
        Refresh_G1_Gam()
        Refresh_G1_Marq()
        '
        Clear_graphique()
        Graphique_Divisions()

        ' Détermination des calques Tons, Accords Gammes à afficher

        For i = 0 To ChoixCalquesLocal.Count - 1
            If ChoixCalquesLocal(i) = True Then
                Select Case i
                    Case 0 ' Ton
                        Graphique_Tons()
                        Graphique_Divisions()
                    Case 1 ' Gammes
                        Graphique_Gammes()
                        Graphique_Divisions()
                    Case 2 ' Accords
                        Graphique_Accords()
                        Graphique_Divisions()
                    Case 3 ' 12/8
                        Graphique_12_8()
                    Case 4 ' Pédale
                        Graphique_Tonique()
                        Graphique_Divisions()

                    Case 5 ' Tessiture
                        Maj_Tessiture(TessDebLocale, TessFinLocale)
                End Select
            End If
        Next i
        '
        Grid1.AutoRedraw = True
        Grid1.Refresh()
    End Sub
    Sub Nouv_CalquesMIDI()
        Grid1.AutoRedraw = False
        '
        For i = 0 To nbCalques - 1
            ChoixCalquesLocal(i) = False
        Next
        ChoixCalquesLocal(1) = True ' activation du calque gamme par défaut
        '
        PédaleLocale = 0
        TessDebLocale = "C-2" '127 - ValNoteCubase.IndexOf("C-2")
        TessFinLocale = "G8"
        TessListeLocale = 0
        InitColonneNotes() ' raz graphique tessiture
        '
        Grid1.AutoRedraw = True
        Grid1.Refresh()
    End Sub
    Sub InitColonneNotes()
        Dim i As Integer
        ' RAZ couleur colonnes des notes
        ' ******************************
        'Grid1.BackColorFixed = Color.Beige
        For i = 6 To Grid1.Rows - 1
            Grid1.Cell(i, 0).BackColor = Color.Beige
            Grid1.Cell(i, 0).ForeColor = Color.Black
        Next
    End Sub
    Sub Maj_Tessiture(Tdeb As String, Tfin As String)
        Dim i As Integer

        ' Recherche de la 1ere note de la tessiture
        ' *****************************************
        i = Grid1.Rows ' - 1
        Do
            i = i - 1
        Loop Until Trim(Grid1.Cell(i, 0).Text) = Tdeb Or i <= 6
        '
        ' Colorisation de la tessiture jusqu'à la dernière note
        ' *****************************************************
        i = i + 1
        Do
            i = i - 1
            Grid1.Cell(i, 0).BackColor = Color.Yellow
            Grid1.Cell(i, 0).ForeColor = Color.Black
        Loop Until Trim(Grid1.Cell(i, 0).Text) = Tfin Or i <= 6
        '
    End Sub
    Sub ActionsPianoR_Click(sender As Object, e As EventArgs)
        '
        ActionsPianoRoll.N_Can1er = N_Can1erPianoR
        ActionsPianoRoll.N_Canal = Me.Canal
        ActionsPianoRoll.ShowDialog()
        Select Case ActionsPianoRoll.TypeAction
            Case ActionsPianoRoll.TypeAct.Effacer    ' EFFACER
                If MessageHV.PSortie = "Oui" Then
                    ActionEffacer()
                End If
            Case ActionsPianoRoll.TypeAct.Coller     ' COLLER
                If MessageHV.PSortie = "Oui" Then
                    ActionColler()
                End If
            Case ActionsPianoRoll.TypeAct.Transposer ' TRANSPOSER
                If MessageHV.PSortie = "Oui" Then
                    ActionTransposer(ActionsPianoRoll.Début, ActionsPianoRoll.Fin, ActionsPianoRoll.ValeurTransp)
                End If
            Case ActionsPianoRoll.TypeAct.CollerVers     ' COLLER VERS
                If MessageHV.PSortie = "Oui" Then
                    ActionCollerVers()
                End If
            Case ActionsPianoRoll.TypeAct.DéplacerVers   ' DEPLACER VERS (nrexiste plus - ne fonctionne pas)
                If MessageHV.PSortie = "Oui" Then
                    ActionCollerVers()
                    ActionEffacer_Déplacer()
                End If
            Case ActionsPianoRoll.TypeAct.MoidifierVel   ' DEPLACER VERS
                If MessageHV.PSortie = "Oui" Then
                    ActionVel_Aléa()
                End If
        End Select
    End Sub
    Sub ActionVel_Aléa()
        Dim i, j, k, m As Integer
        Dim Coldeb, ColFin As Integer
        Dim rand As New Random() ' Création une instance de la classe Random

        ' Préparation pour CTRL+Z
        ' ***********************
        Dim aa As New Ann
        'ListAnnulation.Clear()
        'PointAnn = -1
        '
        ListAnnulation.Add(aa)
        PointAnn = PointAnn + 1
        ListAnnulation.Item(PointAnn).Action = ActionEnum.Restituer
        '
        ' effacer les notes
        ' *****************
        Coldeb = Me.DivisionMes * (ActionsPianoRoll.Début - 1) + 1
        ColFin = Me.DivisionMes * (ActionsPianoRoll.Fin) ' 
        For i = Grid1.FixedRows - 1 To Grid1.Rows - 1
            For j = Coldeb To ColFin
                If IsNumeric(Grid1.Cell(i, j).Text) Then
                    m = 0
                    Dim oo As New Ann.SauvAnnuler
                    k = j
                    oo.Vélo = Trim(Grid1.Cell(i, k).Text)
                    oo.Ligne = i
                    oo.Colonne = k
                    Grid1.Cell(i, k).Text = rand.Next(ActionsPianoRoll.ValMin, ActionsPianoRoll.ValMax)
                    ListAnnulation.Item(PointAnn).ListAnnuler.Add(oo) ' sauvegarde pour CTRL+Z
                    ListAnnulation.Item(PointAnn).Tinfo = TypeInfo.Vélocités ' 
                End If
            Next
        Next
    End Sub



    Sub ActionEffacer()
        Dim i, j, k, m As Integer
        Dim Coldeb, ColFin As Integer

        ' Préparation pour CTRL+Z
        ' ***********************
        Dim aa As New Ann
        ListAnnulation.Clear()
        PointAnn = -1
        '
        ListAnnulation.Add(aa)
        PointAnn = PointAnn + 1
        ListAnnulation.Item(PointAnn).Action = ActionEnum.Restituer
        ' effacer les notes
        ' *****************
        Coldeb = Me.DivisionMes * (ActionsPianoRoll.Début - 1) + 1
        ColFin = Me.DivisionMes * (ActionsPianoRoll.Fin) ' 
        For i = Grid1.FixedRows - 1 To Grid1.Rows - 1
            For j = Coldeb To ColFin
                If IsNumeric(Grid1.Cell(i, j).Text) Then
                    m = 0
                    Dim oo As New Ann.SauvAnnuler
                    k = j
                    oo.Vélo = Trim(Grid1.Cell(i, k).Text)
                    oo.Ligne = i
                    oo.Colonne = k
                    Do
                        Grid1.Cell(i, k).Text = "" ' effacer note : vélo et longueur
                        k = k + 1
                        m = m + 1
                    Loop Until Trim(Grid1.Cell(i, k).Text) = "" Or IsNumeric(Grid1.Cell(i, k).Text) = True
                    oo.Longueur = m
                    ListAnnulation.Item(PointAnn).ListAnnuler.Add(oo) ' sauvegarde pour CTRL+Z
                    ListAnnulation.Item(PointAnn).Tinfo = TypeInfo.All
                End If
            Next
        Next
    End Sub
    Sub ActionEffacer_Déplacer()
        Dim i, j, k, m As Integer
        Dim Coldeb, ColFin As Integer


        ' Préparation pour CTRL+Z
        ' ***********************
        Dim aa As New Ann
        'ListAnnulation.Clear()
        'PointAnn = -1
        '
        ListAnnulation.Add(aa)
        PointAnn += 1
        ListAnnulation.Item(PointAnn).Action = ActionEnum.Restituer_DeplaverVers
        '
        ' effacer les notes
        ' *****************
        Coldeb = Me.DivisionMes * (ActionsPianoRoll.Début - 1) + 1
        ColFin = Me.DivisionMes * (ActionsPianoRoll.Fin - 1) ' 
        For i = Grid1.FixedRows - 1 To Grid1.Rows - 1
            For j = Coldeb To ColFin
                If IsNumeric(Grid1.Cell(i, j).Text) Then
                    m = 0
                    Dim oo As New Ann.SauvAnnuler
                    k = j
                    oo.Vélo = Trim(Grid1.Cell(i, k).Text)
                    oo.Ligne = i
                    oo.Colonne = k
                    Do
                        Grid1.Cell(i, k).Text = "" ' effacer note : vélo et longueur
                        k += 1
                        m += 1
                    Loop Until Trim(Grid1.Cell(i, k).Text) = "" Or IsNumeric(Grid1.Cell(i, k).Text) = True
                    oo.Longueur = m
                    ListAnnulation.Item(PointAnn).ListAnnuler.Add(oo) ' sauvegarde pour CTRL+Z
                    ListAnnulation.Item(PointAnn).Tinfo = TypeInfo.All
                End If
            Next
        Next
    End Sub
    Sub ActionColler()
        Dim i As Integer
        Dim FirstR As Integer = Grid1.FixedRows '(5)
        Dim LastR As Integer = Grid1.Rows - 1
        Dim FirstC As Integer = (ActionsPianoRoll.Début - 1) * 16 + 1
        Dim LastC As Integer = (ActionsPianoRoll.Fin) * 16
        Dim SauvTopRow As Integer = Grid1.TopRow

        ' Sélection
        ' *********
        i = Grid1.TopRow
        Grid1.Range(FirstR, FirstC, LastR, LastC).SelectCells()
        Grid1.LeftCol = LastC
        Grid1.TopRow = i
        '
        ' Copie
        ' *****
        CopierData()
        '
        ' Aller Vers (sélection de la cellule de départ de la copie)
        ' ***********************************************************
        '
        Dim ii As Integer = ActionsPianoRoll.Destination
        Grid1.LeftCol = ((ii - 1) * 16) + 1
        Grid1.TopRow = Grid1.FixedRows
        Grid1.Range(Grid1.TopRow, Grid1.LeftCol, Grid1.TopRow, Grid1.LeftCol).SelectCells() ' sélection pour le collage
        '
        ' Coller
        ' ******
        CollerData2()
        'Form1.ReCalcul()
        '
        ' restitution de la toprow de départ
        Grid1.TopRow = SauvTopRow
    End Sub
    Sub ActionCollerVers()
        Dim i As Integer
        Dim FirstR As Integer = Grid1.FixedRows '(5)
        Dim LastR As Integer = Grid1.Rows - 1
        Dim FirstC As Integer = (ActionsPianoRoll.Début - 1) * 16 + 1
        Dim LastC As Integer = (ActionsPianoRoll.Fin) * 16

        Dim SauvTopRow As Integer = Grid1.TopRow
        Dim SauvLeftCol As Integer = Grid1.LeftCol

        ' Sélection
        ' *********
        i = Grid1.TopRow
        Grid1.Range(FirstR, FirstC, LastR, LastC).SelectCells()
        Grid1.LeftCol = LastC
        Grid1.TopRow = i
        '
        ' Copie
        ' *****
        CopierData()
        '
        ' Aller Vers (sélection de la cellule de départ de la copie)
        ' ***********************************************************
        '
        Dim ii As Integer = ActionsPianoRoll.Destination
        ii = ((ii - 1) * 16) + 1
        Grid1.LeftCol = ii '((ii - 1) * 16) + 1
        Grid1.TopRow = Grid1.FixedRows
        Grid1.Range(Grid1.TopRow, Grid1.LeftCol, Grid1.TopRow, Grid1.LeftCol).SelectCells() ' sélection pour le collage
        '
        ' Coller
        ' ******
        CollerData3()
        'Form1.ReCalcul()
        '
        ' restitution de la toprow de départ et left col de destination
        ' *************************************************************
        Grid1.TopRow = SauvTopRow
        Grid1.LeftCol = SauvLeftCol
    End Sub
    Sub CollerData3()
        Dim i, j, k, m As Integer
        Dim LigneInitial, ColonneInitial As Integer
        Dim aa As New Ann
        Dim NPRoll As Integer = ActionsPianoRoll.N_PianoR
        '
        ' Préparation pour CTRL+Z
        ' ***********************
        With Form1.listPIANOROLL.Item(NPRoll)
            .ListAnnulation.Clear()
            .PointAnn = -1
            '
            .ListAnnulation.Add(aa)
            .PointAnn = .PointAnn + 1
            .ListAnnulation.Item(.PointAnn).Action = ActionEnum.EffacerVers
        End With

        ' Collage Notes
        ' *************
        LigneInitial = Grid1.Selection.FirstRow
        ColonneInitial = Grid1.Selection.FirstCol
        '
        For Each a As NotesCCC In ListCCC
            Dim oo As New Ann.SauvAnnuler
            i = LigneInitial + a.Ecartligne
            j = ColonneInitial + a.EcartCol
            If i <= Grid1.Rows - 1 And j <= Grid1.Cols - 1 Then
                oo.Ligne = i
                oo.Colonne = j
                oo.Vélo = a.Vélo
                '
                With Form1.listPIANOROLL.Item(NPRoll)
                    .Grid1.Cell(i, j).Text = a.Vélo
                    .Grid1.Cell(i, j).FontBold = True
                    m = 1
                    For k = j + 1 To j + a.Longueur - 1
                        If k <= Grid1.Cols - 1 Then
                            If IsNumeric(Grid1.Cell(i, k).Text) = False Then
                                m = m + 1
                                .Grid1.Cell(i, k).Text = Trait
                                .Grid1.Cell(i, k).FontBold = True
                            Else
                                Exit For
                            End If
                        Else
                            Exit For
                        End If
                    Next k

                    oo.Longueur = m
                    .ListAnnulation.Item(.PointAnn).ListAnnuler.Add(oo) ' sauvegarde pour CTRL+Z
                    .ListAnnulation.Item(.PointAnn).Tinfo = TypeInfo.All
                    'ListAnnulation.Item(PointAnn).Action = ActionEnum.EffacerRestituer
                End With
            End If
        Next
        '
        ' Collage CTRL des pédales
        ' ************************
        For Each a As CtrlCCC In ListCCCctrl
            Dim ooo As New Ann.SauvAnnulerCtrl
            i = Grid1.FixedRows - 1
            j = ColonneInitial + a.EcartCol
            '
            ooo.Colonne = j
            If j <= Grid1.Cols - 1 Then
                With Form1.listPIANOROLL.Item(NPRoll)
                    .Grid1.Cell(i, j).Text = a.Valeur
                    ooo.Valeur = a.Valeur
                    .ListAnnulation.Item(.PointAnn).ListAnnulerCTRL.Add(ooo) ' ' sauvegarde pour CTRL+Z
                    .ListAnnulation.Item(.PointAnn).Tinfo = TypeInfo.All
                End With
            End If
        Next
        '
        'Grid1.TopRow = LigneInitial
        'Grid1.LeftCol = ColonneInitial - 10

        'Form1.listPIANOROLL.Item(NPRoll).Grid1.Cell(LigneInitial, ColonneInitial).EnsureVisible()
    End Sub
    Public Sub ActionTransposer(Début As Integer, Fin As Integer, ValTransp As Integer)
        Dim i, ii, j, k, lig, col, m, n, Coldeb, Colfin As Integer

        If Pré_VérifTransp(Début, Fin, ValTransp) Then
            '
            ' Préparation pour CTRL+Z
            ' ***********************
            Dim aa As New Ann
            ListAnnulation.Clear()
            PointAnn = -1
            '
            ListAnnulation.Add(aa)
            PointAnn = PointAnn + 1
            ListAnnulation.Item(PointAnn).Action = ActionEnum.EffacerSelRestituer
            ListAnnulation.Item(PointAnn).SelDeb = Début
            ListAnnulation.Item(PointAnn).SelFin = Fin

            ' Tranposition
            ' ************
            n = ValTransp
            Coldeb = Me.DivisionMes * (Début - 1) + 1
            Colfin = Me.DivisionMes * (Fin) '
            If n > 0 Then
                For i = Grid1.FixedRows - 1 To Grid1.Rows - 1
                    For j = Coldeb To Colfin
                        If IsNumeric(Grid1.Cell(i, j).Text) Then
                            'mémo et effacement note
                            m = 0
                            col = j
                            Dim oo As New Ann.SauvAnnuler

                            k = j
                            oo.Vélo = Trim(Grid1.Cell(i, k).Text)
                            oo.Ligne = i
                            oo.Colonne = k
                            Do
                                Grid1.Cell(i, k).Text = "" ' effacer note : vélo et longueur
                                k = k + 1
                                m = m + 1
                            Loop Until Trim(Grid1.Cell(i, k).Text) = "" Or IsNumeric(Grid1.Cell(i, k).Text) = True Or k = Colfin
                            oo.Longueur = m
                            '
                            ' tranposition de la note
                            lig = i - n
                            If lig > Grid1.FixedRows - 1 And lig < Grid1.Rows - 1 Then
                                Grid1.Cell(lig, col).Text = oo.Vélo
                                For ii = col + 1 To col + (m - 1)
                                    Grid1.Cell(lig, ii).Text = Trait
                                Next
                            End If

                            ListAnnulation.Item(PointAnn).ListAnnuler.Add(oo) ' sauvegarde pour CTRL+Z
                            ListAnnulation.Item(PointAnn).Tinfo = TypeInfo.All
                        End If
                    Next
                Next
            Else
                'For i = Grid1.FixedRows - 1 To Grid1.Rows - 1
                For i = Grid1.Rows - 1 To Grid1.FixedRows - 1 Step -1
                    For j = Coldeb To Colfin
                        If IsNumeric(Grid1.Cell(i, j).Text) Then
                            'mémo et effacement note
                            m = 0
                            col = j
                            Dim oo As New Ann.SauvAnnuler

                            k = j
                            oo.Vélo = Trim(Grid1.Cell(i, k).Text)
                            oo.Ligne = i
                            oo.Colonne = k
                            Do
                                Grid1.Cell(i, k).Text = "" ' effacer note : vélo et longueur
                                k = k + 1
                                m = m + 1
                            Loop Until Trim(Grid1.Cell(i, k).Text) = "" Or IsNumeric(Grid1.Cell(i, k).Text) = True
                            oo.Longueur = m
                            '
                            ' tranposition de la note
                            lig = i - n ' ici n contient une valeur négative, il s'agit d'une addition
                            If lig > Grid1.FixedRows - 1 And lig < Grid1.Rows - 1 Then
                                Grid1.Cell(lig, col).Text = oo.Vélo
                                For ii = col + 1 To col + (m - 1)
                                    Grid1.Cell(lig, ii).Text = Trait
                                Next
                            End If

                            ListAnnulation.Item(PointAnn).ListAnnuler.Add(oo) ' sauvegarde pour CTRL+Z
                            ListAnnulation.Item(PointAnn).Tinfo = TypeInfo.All
                        End If
                    Next
                Next
            End If
        End If
    End Sub
    Public Function Pré_VérifTransp(Début As Integer, Fin As Integer, ValTransp As Integer) As Boolean
        Dim n, coldeb, colfin As Integer

        Pré_VérifTransp = True
        n = ValTransp
        coldeb = Me.DivisionMes * (Début - 1) + 1
        colfin = Me.DivisionMes * (Fin) ' on neretire pas -1 à FIN car la dernière mesure est comprise dans l'effacement
        For i = Grid1.FixedRows - 1 To Grid1.Rows - 1
            For j = coldeb To colfin
                If IsNumeric(Grid1.Cell(i, j).Text) Then
                    'mémo et effacement note
                    m = 0
                    col = j
                    Dim oo As New Ann.SauvAnnuler

                    k = j
                    oo.Vélo = Trim(Grid1.Cell(i, k).Text)
                    oo.Ligne = i
                    oo.Colonne = k
                    Do
                        k = k + 1
                        m = m + 1
                    Loop Until Trim(Grid1.Cell(i, k).Text) = "" Or IsNumeric(Grid1.Cell(i, k).Text) = True Or k = colfin
                    oo.Longueur = m
                    '
                    ' tranposition de la note
                    lig = i - n

                    If Not (lig > Grid1.FixedRows - 1 And lig < Grid1.Rows - 1) Then ' détection du dépassement de tessiture 
                        Dim titre As String
                        Dim mess As String
                        If Module1.LangueIHM = "fr" Then
                            titre = "Avertissement"
                            mess = "Dépassement de tessiture dans PianoRoll" + Convert.ToString(Me.Canal - 5) + ": tranposition annulée."
                        Else
                            titre = "Warning"
                            mess = "Range overrun in PianoRoll" + Convert.ToString(Me.Canal - 5) + " : transposition canceled."
                        End If
                        Cacher_FormTransparents()
                        MessageBox.Show(mess, titre, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Pré_VérifTransp = False
                    End If
                End If
            Next
        Next
    End Function

    Private Sub ListDynF2_SelectedIndexChanged(sender As Object, e As EventArgs)
        ' à programmer
    End Sub
    Private Sub Couper_Click(sender As Object, e As EventArgs)
        If Me.Orig_PianoR.Orig1 = OrigPianoCourbe.Piano Then
            CouperData() ' couper notes
            'Form1.ReCalcul()
        Else
            Dim ind As Byte = Orig_PianoR.N_Courbe
            Me.GridCourbes.Item(ind).Selection.CutData() ' couper courbes
        End If

    End Sub
    Private Sub Copier_Click(sender As Object, e As EventArgs)
        If Me.Orig_PianoR.Orig1 = OrigPianoCourbe.Piano Then
            CopierData()
        Else
            Dim ind As Byte = Orig_PianoR.N_Courbe
            Me.GridCourbes.Item(ind).Selection.CopyData() '
        End If
    End Sub
    Private Sub Coller_Click(sender As Object, e As EventArgs)

        If Me.Orig_PianoR.Orig1 = OrigPianoCourbe.Piano Then
            CollerData2()
            'Form1.ReCalcul()
        Else
            Dim ind As Byte = Orig_PianoR.N_Courbe
            Orig_PianoR.Vérouillage = True ' impossibilité de Tracer pendant un collage
            Me.GridCourbes.Item(ind).Selection.PasteData()
            Orig_PianoR.Vérouillage = False ' retour à la possibilité de Tracer
        End If
    End Sub
    Private Sub Annuler_Click(sender As Object, e As EventArgs)
        AnnulerData2()
    End Sub
    Private Sub Supprimer_Click(sender As Object, e As EventArgs)
        EffacerSelection()
    End Sub
    Sub Quitter_Click(sender As Object, e As EventArgs)
        'Form1.Quitter()
        Dim NumOnglet As Integer = Convert.ToUInt16(F1.Tag)
        Attacher(NumOnglet)
    End Sub
    '
    ' *********************************************************
    ' CONSTRUCTION
    ' *********************************************************

    Sub Construction_Grid1()
        Dim i, j, k As Integer

        Grid1.AutoRedraw = False
        ' autorisation actions
        ' ********************
        Grid1.AllowDrop = False
        Grid1.AllowUserPaste = ClipboardDataEnum.Text
        Grid1.AllowUserReorderColumn = False
        Grid1.AllowUserResizing = False
        Grid1.AllowUserSort = False

        Grid1.SelectionMode = SelectionModeEnum.Free
        Grid1.BackColorBkg = Color.White
        Grid1.SelectionBorderColor = Color.Red
        Grid1.BackColorActiveCellSel = Color.Yellow
        Grid1.BackColorSel = Color.PaleGreen

        Dim f As New System.Drawing.Font("Tahoma", 8, FontStyle.Bold)
        Grid1.DefaultFont = f
        '
        Grid1.Cols = Me.DivisionMes * Me.NbMes + 1 '* Me.nbRépétitionMax
        Grid1.FixedRows = 6
        Grid1.Rows = 128 + Grid1.FixedRows
        '
        ' Numérotations
        i = 1
        j = 1
        k = 1
        Do
            Grid1.Range(2, i, 2, i + (Me.DivisionMes - 1)).Merge() ' Merges pour Accords
            Grid1.Range(3, i, 3, i + (Me.DivisionMes - 1)).Merge() ' Merges pour  Gammes
            Grid1.Range(4, i, 4, i + (Me.DivisionMes - 1)).Locked = False
            Grid1.Range(4, i, 4, i + (Me.DivisionMes - 1)).Merge() ' Merges pour  variations
            Grid1.Range(4, i, 4, i + (Me.DivisionMes - 1)).Locked = True
            Grid1.Cell(0, i).Text = Convert.ToString(k - 1)
            Grid1.Cell(0, i).FontBold = True
            i = j
            j = j + Me.DivisionMes
            k = k + 1
        Loop Until j > (Me.NbMes * Me.DivisionMes) + 1 '* Me.nbRépétitionMax
        '
        ' Dimension des lignes - colonnes
        For i = 0 To Grid1.Rows - 1
            Grid1.Row(i).Height = 26
        Next
        '
        For j = 0 To Grid1.Cols - 1
            Grid1.Column(j).Width = 26
        Next j
        '
        For i = (Grid1.FixedRows + 1) To Grid1.Rows - 1
            For j = 0 To Grid1.Cols - 1
                Grid1.Cell(i, j).Alignment = AlignmentEnum.CenterCenter
            Next
        Next
        '
        Grid1.Range(4, 0, Grid1.Rows - 1, 0).FontBold = False
        Grid1.Range(5, 0, Grid1.Rows - 1, Grid1.Cols - 1).Locked = True
        '

        Grid1.Column(0).Width = 28
        '
        ' Maj colonnes des notes

        j = Grid1.Rows - 1
        For i = 0 To 127
            Grid1.Cell(j, 0).Text = ValNoteCubase.Item(i)
            j = j - 1
        Next
        Grid1.Cell(Grid1.FixedRows - 1, 0).Text = "Ped."

        Dim s As New Size With {
            .Width = Panneau1.Size.Width,
            .Height = 500
        }
        Grid1.Size = s
        Grid1.BringToFront()
        '
        ' mise à jour des accords et gammes
        Refresh_G1_Acc()
        Refresh_G1_Gam()
        Refresh_G1_Transf()
        '
        Grid1.Dock = DockStyle.Fill
        '
        Grid1.AutoRedraw = True
        Grid1.Refresh()
    End Sub
    Sub Construction_Grid1_CRoll()
        Dim i, j As Integer

        Grid1.AutoRedraw = False
        ' autorisation actions
        ' ********************
        Grid1.AllowDrop = False
        Grid1.AllowUserPaste = ClipboardDataEnum.Text
        Grid1.AllowUserReorderColumn = False
        Grid1.AllowUserResizing = False
        Grid1.AllowUserSort = False

        Grid1.SelectionMode = SelectionModeEnum.Free
        Grid1.BackColorBkg = Color.White
        Grid1.SelectionBorderColor = Color.Red
        Grid1.BackColorActiveCellSel = Color.Yellow
        Grid1.BackColorSel = Color.PaleGreen

        Dim f As New System.Drawing.Font("Tahoma", 8, FontStyle.Bold)
        Grid1.DefaultFont = f
        '
        Grid1.Cols = 33
        Grid1.FixedRows = 1
        Grid1.Rows = 16

        ' Dimension des lignes - colonnes
        For i = 0 To Grid1.Rows - 1
            Grid1.Row(i).Height = 27
        Next
        '
        For j = 0 To Grid1.Cols - 1
            Grid1.Column(j).Width = 27
        Next j
        '
        For i = (Grid1.FixedRows + 1) To Grid1.Rows - 1
            For j = 0 To Grid1.Cols - 1
                Grid1.Cell(i, j).Alignment = AlignmentEnum.CenterCenter
            Next
        Next
        '
        'Grid1.Range(1, 0, Grid1.Rows - 1, 0).Font = f
        'Grid1.Range(1, 0, Grid1.Rows - 1, Grid1.Cols - 1).Locked = True
        '
        '
        ' Maj colonnes des notes
        Grid1.Cell(1, 0).Text = "V5+"
        Grid1.Cell(1, 0).ForeColor = Color.Blue
        Grid1.Cell(2, 0).Text = "V4+"
        Grid1.Cell(2, 0).ForeColor = Color.Blue
        Grid1.Cell(3, 0).Text = "V3+"
        Grid1.Cell(3, 0).ForeColor = Color.Blue
        Grid1.Cell(4, 0).Text = "V2+"
        Grid1.Cell(4, 0).ForeColor = Color.Blue
        Grid1.Cell(5, 0).Text = "V1+"
        Grid1.Cell(5, 0).ForeColor = Color.Blue

        Grid1.Cell(6, 0).Text = "V5"
        Grid1.Cell(7, 0).Text = "V4"
        Grid1.Cell(8, 0).Text = "V3"
        Grid1.Cell(9, 0).Text = "V2"
        Grid1.Cell(10, 0).Text = "V1"

        Grid1.Cell(11, 0).Text = "V5-"
        Grid1.Cell(11, 0).ForeColor = Color.Red
        Grid1.Cell(12, 0).Text = "V4-"
        Grid1.Cell(12, 0).ForeColor = Color.Red
        Grid1.Cell(13, 0).Text = "V3-"
        Grid1.Cell(13, 0).ForeColor = Color.Red
        Grid1.Cell(14, 0).Text = "V2-"
        Grid1.Cell(14, 0).ForeColor = Color.Red
        Grid1.Cell(15, 0).Text = "V1-"
        Grid1.Cell(15, 0).ForeColor = Color.Red


        '
        Grid1.Width = 5000
        Grid1.Height = 1000
        '
        Dim s As New Size With {
            .Width = Grid1.Width,
            .Height = Grid1.Height
        }
        Grid1.Size = s
        Grid1.ScrollBars = ScrollBarsEnum.Both
        Grid1.Dock = DockStyle.Fill
        Grid1.BringToFront()
        '


        Grid1.AutoRedraw = True
        Grid1.Refresh()


    End Sub
    Sub Construction_BarreOutils()
        Dim i As Integer
        Dim S As New Size()
        Dim P As New Point()
        Dim f As New System.Drawing.Font("Tahoma", 8, FontStyle.Regular)
        Dim sauv As Integer
        Dim P1 As New Point
        ' 
        Panneau1.Panel1.BackColor = Me.CoulBarOut

        ' Mute Canal
        ' **********
        P.X = 2
        P.Y = 5
        S.Height = 15
        S.Width = 80
        CheckMute.Size = S
        CheckMute.Location = P
        Panneau1.Panel1.Controls.Add(CheckMute)
        CheckMute.TabStop = False
        CheckMute.Checked = True
        If Me.Langue = "fr" Then
            CheckMute.Text = "PISTE" + " " + Trim(Convert.ToString((Me.Canal + 1)))
        Else
            CheckMute.Text = "TRACK" + " " + Trim(Convert.ToString(Str(Me.Canal + 1)))
        End If
        '
        CheckMute.Font = New Font("Arial", 8, FontStyle.Bold) ' Module1.fontMutePiste
        CheckMute.ForeColor = Color.DarkRed
        CheckMute.Visible = True
        CheckMute.Enabled = False
        '
        AddHandler CheckMute.CheckedChanged, AddressOf CheckMute_CheckedChanged
        AddHandler CheckMute.Click, AddressOf CheckMute_Click
        AddHandler CheckMute.KeyDown, AddressOf CheckMute_KeyDown

        ' SoloBout
        ' ********
        Panneau1.Panel1.Controls.Add(SoloBoutPR)
        SoloBoutPR.TabStop = False
        SoloBoutPR.Location = New Point(2, 27)
        SoloBoutPR.Font = New Font("Calibri", 8, FontStyle.Regular)
        SoloBoutPR.Size = New Size(50, 20)
        SoloBoutPR.BackColor = Color.Beige
        SoloBoutPR.Text = "SOLO"
        SoloBoutPR.Visible = False
        AddHandler SoloBoutPR.MouseUp, AddressOf SoloBoutPR_MouseUp
        '
        ' liste des types de notes
        ' ************************
        '
        P.X = P.X + S.Width
        P.Y = 5
        S.Width = 40
        ListTypNote.Size = S
        ListTypNote.Location = P
        '
        P.Y = 27
        LabelTypNote.Location = P
        LabelTypNote.AutoSize = True

        Panneau1.Panel1.Controls.Add(ListTypNote)
        ListTypNote.TabStop = False
        Panneau1.Panel1.Controls.Add(LabelTypNote)
        LabelTypNote.TabStop = False
        If Langue = "fr" Then
            ListTypNote.Items.Add("RN")
            ListTypNote.Items.Add("BL")
            ListTypNote.Items.Add("NR")
            ListTypNote.Items.Add("CR")
            ListTypNote.Items.Add("DC")
            LabelTypNote.Text = "Durée"
        Else
            ListTypNote.Items.Add("WN")
            ListTypNote.Items.Add("HN")
            ListTypNote.Items.Add("QN")
            ListTypNote.Items.Add("EN")
            ListTypNote.Items.Add("SN")
            LabelTypNote.Text = "Length"
        End If
        ListTypNote.Font = f
        ListTypNote.SelectedIndex = 2
        '
        ' liste des dynamiques
        ' ********************
        P.X = P.X + S.Width + 5
        P.Y = 5
        S.Width = 40
        '
        Panneau1.Panel1.Controls.Add(ListDynF1)
        ListDynF1.TabStop = False
        Panneau1.Panel1.Controls.Add(LabelDyn)
        LabelDyn.TabStop = False
        ListDynF1.Size = S
        ListDynF1.Location = P
        ListDynF1.Enabled = True
        '
        LabelDyn.Text = "Dyn"
        LabelDyn.AutoSize = True
        P.Y = 27
        LabelDyn.Location = P
        '
        For i = 127 To 0 Step -1
            ListDynF1.Items.Add(Convert.ToString(i))
        Next
        '
        ListDynF1.Font = f
        ListDynF1.SelectedIndex = 37
        '
        ' Liste des programmes
        ' ********************
        P.X = P.X + S.Width + 5
        P.Y = 5
        '
        S.Height = 32
        S.Width = 150
        '
        Maj_PRG(S, P)

        Panneau1.Panel1.Controls.Add(LabelPRG)
        LabelPRG.Text = "PRG"
        LabelPRG.AutoSize = True
        P.Y = 32
        LabelPRG.Location = P

        '
        ' Bouton Actions
        ' **************
        P.X = P.X + S.Width + 10
            P.Y = 4
            S.Width = 64
            S.Height = 22
            Panneau1.Panel1.Controls.Add(ActionsPianoR)
            ActionsPianoR.TabStop = False
            ActionsPianoR.Size = S
            Dim pp As New Point()
            pp = P
            pp.X = pp.X + 12

            ActionsPianoR.Location = pp
            ActionsPianoR.Text = "Actions"

            ' Bouton appel formulaire des calques midi
            ' ****************************************
            Panneau1.Panel1.Controls.Add(OuvrirCalques)
            P.X = ActionsPianoR.Location.X 'P.X + S.Width + 10
            P.Y = 27
            S.Height = 22
            S.Width = 100
            OuvrirCalques.Location = P
            OuvrirCalques.Size = S
            If Langue = "fr" Then
                OuvrirCalques.Text = "Calques MIDI"
            Else
                OuvrirCalques.Text = "MIDI Layers"
            End If
            OuvrirCalques.Visible = True

            ' Cadre Aller Vers
            ' ****************
            Panneau1.Panel1.Controls.Add(CadreAllerVers)
            '
            P.X = ActionsPianoR.Location.X + ActionsPianoR.Size.Width + 65
            P.Y = 4
            S.Height = 43
            S.Width = 77
            '
            CadreAllerVers.Location = P
            CadreAllerVers.Size = S
            '
            Destination.Location = New Point(15, 16)
            Destination.Size = New Size(45, 25)
            Destination.TextAlign = HorizontalAlignment.Center
            Destination.BorderStyle = BorderStyle.FixedSingle
            Destination.Text = "1"
            If Langue = "fr" Then
                CadreAllerVers.Text = "Aller vers"
            Else
                CadreAllerVers.Text = "Go to"
            End If
            '
            CadreAllerVers.Controls.Add(Destination)
            ' --->

            '
            ' Sélection affichage des contrôleurs (checkbox général)
            ' ******************************************************
            Dim positprov As Integer = 800
            'P.X = positprov + S.Width - 200
            P.X = 570
            P.Y = 1
            S.Width = 200
            Panneau1.Panel1.Controls.Add(AffCtrls)
            AffCtrls.Size = S
            AffCtrls.TabStop = False
            AffCtrls.AutoSize = False
            AffCtrls.Font = fnt7
            AffCtrls.TextAlign = ContentAlignment.MiddleLeft
            AffCtrls.Location = P
            If LangueIHM = "fr" Then
                AffCtrls.Text = "CTRLs et Support"
            Else
                AffCtrls.Text = "CTRLs and Support"
            End If

            AffCtrls.Checked = False
            '
            P.Y = 10 ' pour les composants suivants

            ' Courbes contrôleur
            ' ******************  ' Courbes contrôleur
            S.Width = 70
            sauv = P.X
            P.X = P.X + S.Width

            P1 = P

            Dim ii As Integer = 45
            P1.X = P1.X - 40

            Dim P2 As New Point
            P2 = P1
            P2.Y = P2.Y + 18

            For i = 0 To nbCourbes - 1
                CCActif.Add(New CheckBox)
                Panneau1.Panel1.Controls.Add(CCActif.Item(i))
                CCActif.Item(i).TabStop = False
                CCActif.Item(i).Visible = False
                CCActif.Item(i).Checked = True
                Select Case i
                    Case 0
                        CCActif.Item(i).Text = "Exp"
                        P1.X = P1.X + ii
                        CCActif.Item(i).Location = P1
                    Case 1
                        CCActif.Item(i).Text = "Mod"
                        P1.X = P1.X + ii
                        CCActif.Item(i).Location = P1
                    Case 2
                        CCActif.Item(i).Text = "PAN"
                        P1.X = P1.X + ii
                        CCActif.Item(i).Location = P1
                    Case 3
                        CCActif.Item(i).Text = "50"
                        P2.X = P2.X + ii
                        CCActif.Item(i).Location = P2
                    Case 4
                        CCActif.Item(i).Text = "51"
                        P2.X = P2.X + ii
                        CCActif.Item(i).Location = P2
                    Case 5
                        CCActif.Item(i).Text = "52"
                        P2.X = P2.X + ii
                        CCActif.Item(i).Location = P2
                    Case 6
                        CCActif.Item(i).Text = "53"
                        P2.X = P2.X + ii
                        CCActif.Item(i).Location = P2
                End Select
                CCActif.Item(i).Enabled = False
                CCActif.Item(i).Font = fnt8
                CCActif.Item(i).Checked = False
                '
                CCActif.Item(i).AutoSize = True
            Next


            ' Midi Learn
            ' **********
            P.X = sauv
        Panneau1.Panel1.Controls.Add(MidiLearn)
        MidiLearn.TabStop = False
        'P.X = P.X + S.Width + 10
        P.Y = 30
        S.Width = 140
        If Module1.Langue = "fr" Then
            MidiLearn.Text = "Midi Learn (utilisez les touches de pad +/-)"
        Else
            MidiLearn.Text = "Midi Learn (use +/- pad keys)"
        End If
        MidiLearn.Font = fnt7
        MidiLearn.Checked = False
        MidiLearn.AutoSize = True
        MidiLearn.Location = P 'New Point(300, 5)
        MidiLearn.Visible = False
        MidiLearn.Enabled = False
        MidiLearn.BringToFront()

        ' Affichage des valeurs envoyées pour MIDI Learn
        ' **********************************************
        Panneau1.Panel1.Controls.Add(AffMidiLearn)
        AffMidiLearn.AutoSize = True
        P.X = 750 'P.X + MidiLearn.Size.Width + 150
        P.Y = 15
        S.Width = 30
        'P.X = 2
        AffMidiLearn.Location = P
        AffMidiLearn.Font = fnt7
        AffMidiLearn.Visible = True
        AffMidiLearn.Enabled = True
        AffMidiLearn.Text = "off"
        AffMidiLearn.ForeColor = Color.Red
        AffMidiLearn.BringToFront()

        ' Case à cocher zoom-
        ' *******************
        Panneau1.Panel1.Controls.Add(ZoomMoins)
        AffMidiLearn.AutoSize = True
        P.X = 2 'P.X + AffMidiLearn.Size.Width + 72
        P.Y = 22
        S.Width = 30
        ' 
        ZoomMoins.Location = P
        ZoomMoins.Font = fnt7
        ZoomMoins.Text = "Zoom -"

        '
        ' Bouton Docking
        ' ***************
        P1.X = 830 '570
        P1.Y = 3
        S.Width = 40
        '
        Panneau1.Panel1.Controls.Add(DockButton)
        DockButton.TabStop = False
        '
        DockButton.Size = S
        DockButton.Location = P1
        DockButton.Enabled = True
        DockButton.BackColor = Color.AliceBlue ' 
        DockButton.ForeColor = Color.Red
        '
        If Langue = "fr" Then
            'If Module1.LangueIHM = "fr" Then
            DockButton.Text = "Détacher"
        Else
            DockButton.Text = "Undock"
        End If
        '
        DockButton.Font = fnt5
        DockButton.AutoSize = True
        DockButton.Visible = True
        '
        ' Numric UD Opac
        ' **************
        P.X = P.X + S.Width + 40
        P.Y = 20
        S.Width = 40
        '
        Panneau1.Panel1.Controls.Add(Opacité)
        Opacité.TabStop = False
        Opacité.Size = S
        Opacité.Location = P
        Opacité.Enabled = True
        '
        Opacité.Minimum = 60
        Opacité.Maximum = 100
        Opacité.Value = 100
        Opacité.Visible = False
        '
        ' TextBox Nom du Son
        ' ******************

        Panneau1.Panel1.Controls.Add(NomduSon)
        NomduSon.TabStop = False
        NomduSon.Location = New Point(172, 28)
        NomduSon.Font = Module1.fontNomduSon 'New Font("Comic Sans MS", 8, FontStyle.Bold)
        NomduSon.Size = New Size(150, 8)
        NomduSon.BackColor = Color.Beige
        NomduSon.ForeColor = Color.Black
        NomduSon.ShortcutsEnabled = False
        NomduSon.Multiline = False
        NomduSon.RightToLeft = False
        NomduSon.Visible = True
        NomduSon.BorderStyle = BorderStyle.FixedSingle
        NomduSon.Tag = Me.F1.Tag
        NomduSon.Text = ""
        NomduSon.BringToFront()

        'AddHandler NomduSon.KeyUp, AddressOf NomduSon_KeyUp
        'AddHandler NomduSon.KeyPress, AddressOf NomduSon_KeyPress
        AddHandler NomduSon.TextChanged, AddressOf NomduSon_TextChanged
        ' Label AffNom du Son
        ' ******'Controls.Add(AffNomduSon)
        AffNomduSon.Location = New Point(1120, 27)
        AffNomduSon.TextAlign = ContentAlignment.MiddleLeft
        AffNomduSon.Font = New Font("Rubik", 9, FontStyle.Bold)
        AffNomduSon.Size = New Size(148, 23)
        AffNomduSon.BackColor = Color.Beige
        AffNomduSon.ForeColor = Color.Black
        AffNomduSon.Visible = False
        AffNomduSon.BorderStyle = BorderStyle.FixedSingle
        AffNomduSon.Tag = i


        AddHandler AffNomduSon.MouseUp, AddressOf AffNomduson_MouseUp



        ' Tool Tip (bulles d'aides)
        ' *************************
        If Langue = "fr" Then
            '
            ToolTip1.SetToolTip(BoutonF1_1, "Réduit")
            ToolTip1.SetToolTip(BoutonF1_2, "Large")
            ToolTip1.SetToolTip(ListDynF1, "Dynamique")
            ToolTip1.SetToolTip(ListTypNote, "RN:ronde BL: blanche NR: noire CR: croche DC: Double croche")
            ToolTip1.SetToolTip(ListMod, "Valeur initiale Modulation")
            ToolTip1.SetToolTip(CheckAcc, "Modèle MIDI Accords")
            ToolTip1.SetToolTip(CheckTonique, "Tonique Mode Grec")
            ToolTip1.SetToolTip(ListTonNotes, "Tonique Mode Grec")

            ToolTip1.SetToolTip(DebMacroSel, "Début sélection à coller")
            ToolTip1.SetToolTip(TermeMacroSel, "Fin sélection à coller")
            ToolTip1.SetToolTip(DestMacroSel, "Coller à partir de")
            ToolTip1.SetToolTip(BMacroSel, "Coller sélection")
            ToolTip1.SetToolTip(Destination, "Mesure de destination")
            ToolTip1.SetToolTip(CheckTop, "Sélection Première ligne")
            ToolTip1.SetToolTip(DockButton, "Attacher/Détacher éditeur")
            ToolTip1.SetToolTip(Opacité, "Opacité")
            ToolTip1.SetToolTip(ActiveExpression, "Activation du ctrl expression")
            ToolTip1.SetToolTip(ActionsPianoR, "Actions dans une sélection : Effacer, Coller, Transposer")
            ToolTip1.SetToolTip(NomduSon, "Ecrivez ici le nom du son utilisé")

            ToolTip1.SetToolTip(AffCtrls, "Activer les contrôleurs")
            ToolTip1.SetToolTip(CCActif.Item(0), "Expression - 11")
            ToolTip1.SetToolTip(CCActif.Item(1), "Modulation - 1")
            ToolTip1.SetToolTip(CCActif.Item(2), "Pan - 10")
            ToolTip1.SetToolTip(CCActif.Item(3), "Libre - 50")
            ToolTip1.SetToolTip(CCActif.Item(4), "Libre - 51")
            ToolTip1.SetToolTip(CCActif.Item(5), "Libre - 52")
            ToolTip1.SetToolTip(CCActif.Item(6), "Libre - 53")

        Else
            ToolTip1.SetToolTip(BoutonF1_1, "Small")
            ToolTip1.SetToolTip(BoutonF1_2, "Large")
            ToolTip1.SetToolTip(ListDynF1, "Velocity")
            ToolTip1.SetToolTip(ListTypNote, "WH: whole HN: half QN: quarter EN: eighth  SN: sixteen")
            ToolTip1.SetToolTip(ListMod, "Initial Modulation value")
            ToolTip1.SetToolTip(CheckAcc, "Chords MIDI Model")
            ToolTip1.SetToolTip(CheckTonique, "Greek Mode Tonic")
            ToolTip1.SetToolTip(ListTonNotes, "Greek Mode Tonic")

            ToolTip1.SetToolTip(DebMacroSel, "Start of paste selection")
            ToolTip1.SetToolTip(TermeMacroSel, "End of paste selection")
            ToolTip1.SetToolTip(DestMacroSel, "Paste from")
            ToolTip1.SetToolTip(BMacroSel, "Paste selection")
            ToolTip1.SetToolTip(Destination, "Destination measure")
            ToolTip1.SetToolTip(CheckTop, "Select First line")
            ToolTip1.SetToolTip(DockButton, "Doc/Undock editor")
            ToolTip1.SetToolTip(Opacité, "Opacity")
            ToolTip1.SetToolTip(ActiveExpression, "Ctrl expression activation")
            ToolTip1.SetToolTip(ActionsPianoR, "Actions inside a slection : Clear, Paste, Transpose ")
            ToolTip1.SetToolTip(AffCtrls, "Activate the controllers")
            ToolTip1.SetToolTip(NomduSon, "Write here the name of the sound used")

            ToolTip1.SetToolTip(CCActif.Item(0), "Expression - 11")
            ToolTip1.SetToolTip(CCActif.Item(1), "Modulation - 1")
            ToolTip1.SetToolTip(CCActif.Item(2), "Pan - 10")
            ToolTip1.SetToolTip(CCActif.Item(3), "Free - 50")
            ToolTip1.SetToolTip(CCActif.Item(4), "Free - 51")
            ToolTip1.SetToolTip(CCActif.Item(5), "Free - 52")
            ToolTip1.SetToolTip(CCActif.Item(6), "Free - 53")
        End If

    End Sub
    Sub CheckTonique_CheckedChanged(sender As Object, e As EventArgs)
        'If Not EnChargement Then
        ListTonNotes.Enabled = CheckTonique.Checked
        F1_Refresh()
        'End If
    End Sub
    Sub ListTonNotes_SelectedIndexChanged(sender As Object, e As EventArgs)
        'If Not EnChargement Then
        F1_Refresh()
        'End If
    End Sub
    Public Sub Maj_Tooltips()
        ToolTip1.RemoveAll() ' le RemoveAll est obligatoire pour faire réapparaître les bulles à chaque Undock (autrement, elles n'apparaissent que sur le 1er Undock) 
        ToolTip1.InitialDelay = 250
        ToolTip1.Active = True
        ToolTip1.ShowAlways = True

        If Langue = "fr" Then
            '
            ToolTip1.SetToolTip(BoutonF1_1, "Réduit")
            ToolTip1.SetToolTip(BoutonF1_2, "Large")
            ToolTip1.SetToolTip(ListDynF1, "Dynamique")
            ToolTip1.SetToolTip(ListTypNote, "RN:ronde BL: blanche NR: noire CR: croche DC: Double croche")
            ToolTip1.SetToolTip(ListMod, "Valeur initiale Modulation")
            ToolTip1.SetToolTip(CheckAcc, "Modèle MIDI Accords")

            ToolTip1.SetToolTip(DebMacroSel, "Début sélection à coller")
            ToolTip1.SetToolTip(TermeMacroSel, "Fin sélection à coller")
            ToolTip1.SetToolTip(DestMacroSel, "Coller à partir de")
            ToolTip1.SetToolTip(BMacroSel, "Coller sélection")
            ToolTip1.SetToolTip(Destination, "Mesure de destination")
            ToolTip1.SetToolTip(CheckTop, "Sélection Première ligne")
            ToolTip1.SetToolTip(DockButton, "Attacher/Détacher éditeur")
            ToolTip1.SetToolTip(Opacité, "Opacité")
            ToolTip1.SetToolTip(ActiveExpression, "Activation du ctrl expression")
            'ToolTip1.SetToolTip(Clear_All, "Effacer tout")
            ToolTip1.SetToolTip(AffNomduSon, "Nom du son")
            ToolTip1.SetToolTip(NomduSon, "Nom du son")

            ToolTip1.SetToolTip(AffCtrls, "Activer les contrôleurs")
            ToolTip1.SetToolTip(CCActif.Item(0), "Expression - 11")
            ToolTip1.SetToolTip(CCActif.Item(1), "Modulation - 1")
            ToolTip1.SetToolTip(CCActif.Item(2), "Pan - 10")
            ToolTip1.SetToolTip(CCActif.Item(3), "Libre - 50")
            ToolTip1.SetToolTip(CCActif.Item(4), "Libre - 51")
            ToolTip1.SetToolTip(CCActif.Item(5), "Libre - 52")
            ToolTip1.SetToolTip(CCActif.Item(6), "Libre - 53")

        Else
            '
            ToolTip1.SetToolTip(BoutonF1_1, "Small")
            ToolTip1.SetToolTip(BoutonF1_2, "Large")
            ToolTip1.SetToolTip(ListDynF1, "Velocity")
            ToolTip1.SetToolTip(ListTypNote, "WH: whole HN: half QN: quarter EN: eighth  SN: sixteen")
            ToolTip1.SetToolTip(ListMod, "Initial Modulation value")
            ToolTip1.SetToolTip(CheckAcc, "Chords MIDI Model")

            ToolTip1.SetToolTip(DebMacroSel, "Start of paste selection")
            ToolTip1.SetToolTip(TermeMacroSel, "End of paste selection")
            ToolTip1.SetToolTip(DestMacroSel, "Paste from")
            ToolTip1.SetToolTip(BMacroSel, "Paste selection")
            ToolTip1.SetToolTip(Destination, "Destination measure")
            ToolTip1.SetToolTip(CheckTop, "Select First line")
            ToolTip1.SetToolTip(DockButton, "Doc/Undock editor")
            ToolTip1.SetToolTip(Opacité, "Opacity")
            ToolTip1.SetToolTip(ActiveExpression, "Ctrl expression activation")
            'ToolTip1.SetToolTip(Clear_All, "Clear All")
            ToolTip1.SetToolTip(AffCtrls, "Activate the controllers")
            ToolTip1.SetToolTip(AffNomduSon, "Sound name")
            ToolTip1.SetToolTip(CCActif.Item(0), "Expression - 11")
            ToolTip1.SetToolTip(CCActif.Item(1), "Modulation - 1")
            ToolTip1.SetToolTip(CCActif.Item(2), "Pan - 10")
            ToolTip1.SetToolTip(CCActif.Item(3), "Free - 50")
            ToolTip1.SetToolTip(CCActif.Item(4), "Free - 51")
            ToolTip1.SetToolTip(CCActif.Item(5), "Free - 52")
            ToolTip1.SetToolTip(CCActif.Item(6), "Free - 53")
        End If
    End Sub


    Sub Maj_PRG(S As Size, P As Point)

        Panneau1.Panel1.Controls.Add(ListPRG)

        ListPRG.Visible = True
        ListPRG.Size = S
        ListPRG.Location = P
        ListPRG.Enabled = True
        '
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
        '
    End Sub

    Sub AffNomduson_MouseUp(sender As Object, e As EventArgs)
        Dim com As Label = sender
        Dim ind As Integer
        ind = Val(com.Tag)
        AffNomduSon.Visible = False
        NomduSon.Visible = True
        NomduSon.Location = AffNomduSon.Location 'New Point(1120, 27)
        NomduSon.Text = AffNomduSon.Text
        NomduSon.Focus()
        'NomduSon.Item(ind).Select()

    End Sub
    Public Sub NomduSon_TextChanged(sender As Object, e As EventArgs)
        Dim a As String
        Dim j1, j2, k As Integer
        Dim com As TextBox = sender
        Dim ind As Integer
        ind = Me.Canal
        '
        j1 = InStr(Trim(NomduSon.Text), "&")
        j2 = InStr(Trim(NomduSon.Text), ",")
        If j1 <> 0 Or j2 <> 0 Then
            k = NomduSon.SelectionStart
            RemoveHandler NomduSon.TextChanged, AddressOf NomduSon_TextChanged
            a = NomduSon.Text.Replace("&", "")
            NomduSon.Text = Trim(a)
            a = NomduSon.Text.Replace(",", "")
            NomduSon.Text = Trim(a)
            NomduSon.SelectionStart = k - 1 'Len(a)
            AddHandler NomduSon.TextChanged, AddressOf NomduSon_TextChanged
        End If
        '
        RemoveHandler Form1.Mix.NomduSon.Item(ind).TextChanged, AddressOf Form1.Mix.NomduSon_TextChanged
        Form1.Mix.NomduSon.Item(ind).Text = Trim(NomduSon.Text)
        AddHandler Form1.Mix.NomduSon.Item(ind).TextChanged, AddressOf Form1.Mix.NomduSon_TextChanged
        '
        ' Traitement du nom de l'onglet
        ' *****************************
        a = Trim(NomduSon.Text)
        If Len(NomduSon.Text) <= 25 Then ' pas plus de 25 caractères dans les onglets
            Dim CanalClient As String = (Me.Canal + 1).ToString
            If Trim(a) = "" Then
                Form1.TabControl4.TabPages(Me.F1.Tag).text = CanalClient + " - " + "PianoRoll" + (Me.F1.Tag).ToString
            Else
                Form1.TabControl4.TabPages(Me.F1.Tag).text = CanalClient + " - " + a
            End If
        End If
    End Sub

    Public Sub CheckMute_Click(sender As Object, e As EventArgs)
        If (Form1.Horloge1.IsRunning) Then
            CheckMute.Checked = Not (CheckMute.Checked)
        End If
    End Sub
    Public Sub CheckMute_KeyDown(sender As Object, e As KeyEventArgs)
        e.SuppressKeyPress = True
    End Sub
    Public Sub CheckMute_CheckedChanged(sender As Object, e As EventArgs)
        If Not (Form1.Horloge1.IsRunning) Then
            Form1.Mix.muteVolume.Item(Me.Canal).Checked = CheckMute.Checked
        End If
    End Sub
    Sub SoloBoutPR_MouseUp(sender As Object, e As EventArgs)

        Dim com As Button = sender
        Dim ind As Integer
        ind = Val(com.Tag)
        Form1.Mix.GestMute()
        Dim canMidi As Byte = Me.Canal
        TraitementSoloPR(canMidi)
    End Sub
    Public Sub TraitementSoloPR(CanMidi As Integer)
        Dim i, j, k As Integer
        If (Form1.SoloCours2 = True And Form1.CanMidiCours = CanMidi) Or Form1.SoloCours2 = False Then
            Form1.Mix.Gestion_Solo2(CanMidi) ' gestion du solo dans la table de mixage
            ' gestion des couleurs
            j = 0
            i = CanMidi - 6
            If CanMidi <= 8 Then
                i = CanMidi - 6 ' 
            Else
                i = CanMidi - 7 ' pour dépasser le N° de canal MIDI de la boite à rythm
            End If
            ' gestion des boutons solo dans toute l'appli
            Form1.GestSolo(Form1.SoloCours2)
            ' activation du mode solo
            If Form1.SoloCours2 = False Then
                k = CanMidi
                Form1.listPIANOROLL(i).SoloBoutPR.BackColor = Color.OrangeRed
                Form1.listPIANOROLL(i).SoloBoutPR.ForeColor = Color.Yellow
                Form1.listPIANOROLL(i).SoloBoutPR.Enabled = True
            Else
                k = -1
                Form1.listPIANOROLL(i).SoloBoutPR.BackColor = Color.Beige
                Form1.listPIANOROLL(i).SoloBoutPR.ForeColor = Color.Black
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
    Sub Construction_Menu()
        '
        Fichier.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Quitter})
        Fichier.Size = New System.Drawing.Size(87, 20)
        If Me.Langue = "fr" Then
            Fichier.Text = "Fichier"
        Else
            Fichier.Text = "File"
        End If
        Fichier.Visible = True
        Fichier.BackColor = Me.CoulBarOut
        ' Edition
        Edition.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Couper, Copier, Coller, Séparateur1, Annuler, Séparateur2, Supprimer})
        Edition.Size = New System.Drawing.Size(56, 20)
        Edition.Text = "Edition"
        Edition.Visible = True
        Edition.BackColor = Me.CoulBarOut
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
        F1.Controls.Add(Menu1)
        F1.MainMenuStrip = Menu1
        F1.MainMenuStrip.Visible = False
        F1.MainMenuStrip.Enabled = False
        Menu1.Visible = False
        Menu1.Enabled = False
        'AnnulationHandlerEdition()

    End Sub
    Sub AnnulationHandlerEdition()
        'RemoveHandler Couper.Click, AddressOf Couper_Click
        'RemoveHandler Copier.Click, AddressOf Copier_Click
        'RemoveHandler Coller.Click, AddressOf Coller_Click
        'RemoveHandler Annuler.Click, AddressOf Annuler_Click
        'RemoveHandler Supprimer.Click, AddressOf Supprimer_Click
        'RemoveHandler Quitter.Click, AddressOf Quitter_Click
    End Sub
    Sub ActivationHandlerEdition()
        'AddHandler Couper.Click, AddressOf Couper_Click
        'AddHandler Copier.Click, AddressOf Copier_Click
        'AddHandler Coller.Click, AddressOf Coller_Click
        'AddHandler Annuler.Click, AddressOf Annuler_Click
        'AddHandler Supprimer.Click, AddressOf Supprimer_Click
        'AddHandler Quitter.Click, AddressOf Quitter_Click
    End Sub

    Sub Graphique_Divisions()
        Dim i As Integer
        Dim j As Integer = 1

        If EnChargement = False Then
            '
            Dim fin As Integer = Form1.Det_NumDerAccord()
            '
            Dim RépetAcc() As String = {String.Empty} ' répétition des accords
            Dim listRépet As String = Form1.Récup_AllRépet
            '
            RépetAcc = listRépet.Split(",")
            Dim Répet As Integer = RépetAcc(UBound(RépetAcc))
            '
            DblCrochpMes = 16 ' 4/4
            '
            fin = (fin + (Répet - 1)) * DblCrochpMes

            'If Me.Métrique = "4/4" Then
            'For i = 1 To Grid1.Cols - 1
            For i = 1 To fin
                If j = 1 Then
                    Grid1.Range(0, i, 1, i).BackColor = Color.Gold
                    Grid1.Range(5, i, Grid1.Rows - 1, i).BackColor = Color.Gold
                    Grid1.Cell(1, i).Text = Index_Divisions(j)
                Else
                    If j = 5 Or j = 9 Or j = 13 Then
                        Grid1.Range(0, i, 1, i).BackColor = Color.LightSteelBlue
                        Grid1.Range(5, i, Grid1.Rows - 1, i).BackColor = Color.LightSteelBlue
                        Grid1.Cell(1, i).Text = Index_Divisions(j)
                    End If
                End If
                j = j + 1
                If j = DblCrochpMes + 1 Then j = 1
            Next i
            'End If
        End If
        Grid1.Range(Grid1.FixedRows - 1, 1, Grid1.FixedRows - 1, Grid1.Cols - 1).BackColor = Color.Lavender
        'Grid1.AutoRedraw = True
        'Grid1.Refresh()
    End Sub
    Sub Graphique_12_8_old()
        If Check12_8.Checked = True Then
            Dim Permiertemps As Integer = 1
            Dim RépetAcc() As String = {String.Empty} ' répétition des accords
            Dim listRépet As String = Form1.Récup_AllRépet
            RépetAcc = listRépet.Split(",")
            Dim Répet As Integer = RépetAcc(UBound(RépetAcc))
            Dim fin As Integer = Form1.Det_NumDerAccord()
            Dim DblCrochpMes As Integer = 24


            fin = (fin + (Répet - 1)) * 16

            For i = 1 To fin
                If Permiertemps = 1 Then
                    Grid1.Range(5, i, Grid1.Rows - 1, i).BackColor = CDiv_12_8
                End If
                Permiertemps = Permiertemps + 1
                If Permiertemps = DblCrochpMes + 1 Then Permiertemps = 1
            Next i
        End If
    End Sub
    Sub Graphique_12_8()
        'If Check12_8.Checked = True Then
        Dim Permiertemps As Integer = 1
        'Dim RépetAcc() As String = {String.Empty} ' répétition des accords
        'Dim listRépet As String = Form1.Récup_AllRépet ' ' mis en commentaires lors de l'insertion de pianoroll, drumedit et Mix
        ' RépetAcc = listRépet.Split(",") ' ' mis en commentaires lors de l'insertion de pianoroll, drumedit et Mix
        'Dim Répet As Integer = RépetAcc(UBound(RépetAcc))
        'Dim fin As Integer = Form1.Det_NumDerAccord()
        Dim DblCrochpMes As Integer = 4 '24
        Dim Cmpt_Noires As Integer = 0 ' en 12/8 on change de mesure tous les 36 division

        'fin = (fin + (Répet - 1)) * 16
        'fin = Form1.Det_DerEventH2
        For i = 1 To (nbMesures * 16)
            If Permiertemps = 1 Then
                Grid1.Range(5, i, Grid1.Rows - 1, i).BackColor = CDiv_12_8
                If Cmpt_Noires = 0 Then ' on arrive sur le 1er temps d'une nouvelle mesure
                    Cmpt_Noires = 6
                    Grid1.Range(5, i, Grid1.Rows - 1, i).BackColor = Color.Olive
                End If
                Cmpt_Noires = Cmpt_Noires - 1
            End If
            Permiertemps = Permiertemps + 1
            If Permiertemps = DblCrochpMes + 1 Then
                Permiertemps = 1
            End If
        Next i
        'End If
    End Sub

    Public Sub Graphique_Position() ' mise à jour de Grid1.Top / Remarque : Grid1.leftcol est mis à jour dans la propriété PStartMeasure
        Dim i As Integer
        Dim j As Integer
        Dim coldeb As Integer = ((Me.StartMeasure - 1) * 16) + 1
        Dim colfin As Integer = coldeb + 16
        Dim flagSortir As Boolean = False
        Dim a As String
        '
        Grid1.AutoRedraw = False
        Grid1.TopRow = Grid1.FixedRows + 1
        ' détermination du 1er évènement
        For i = (Grid1.FixedRows + 1) To Grid1.Rows - 1
            For j = coldeb To colfin
                If Trim(Grid1.Cell(i, j).Text) <> "" Then
                    a = Grid1.Cell(i, 0).Text
                    Grid1.TopRow = i
                    flagSortir = True
                    Exit For
                End If
            Next j
            If flagSortir Then Exit For
        Next i
        '
        If flagSortir = False Then Grid1.TopRow = 64
        Grid1.AutoRedraw = True
        Grid1.Refresh()
        'Dim jj As Single = Fix(j / 16) + 1
    End Sub
    Public Sub Graphique_Position2()
        Dim i As Integer
        Dim j As Integer
        Dim coldeb As Integer = ((Me.StartMeasure - 1) * 16) + 1
        Dim colfin As Integer = coldeb + 16
        Dim flagSortir As Boolean = False
        Dim a As String
        '
        Grid1.LeftCol = coldeb
        Grid1.TopRow = Grid1.FixedRows + 1
        ' détermination du 1er évènement
        For i = Grid1.FixedRows To Grid1.Rows - 1
            For j = coldeb To colfin
                If Trim(Grid1.Cell(i, j).Text) <> "" Then
                    a = Grid1.Cell(i, 0).Text
                    Grid1.TopRow = i
                    'Grid1.LeftCol = coldeb
                    flagSortir = True
                    Exit For
                End If
            Next j
            If flagSortir Then Exit For
        Next i
        '
        Grid1.Refresh()
        'Dim jj As Single = Fix(j / 16) + 1
    End Sub
    Public Sub Clear_graphique22()
        Dim tbl() As String
        tbl = Me.ListGam.Split("-")

        Grid1.AutoRedraw = False
        Grid1.Range(5, 1, Grid1.Rows - 1, (UBound(tbl) + 1) * 16).BackColor = Color.White
        Grid1.AutoRedraw = True
        Grid1.Refresh()
    End Sub
    Sub Clear_graphique()
        Dim i, j As Integer
        Dim fin As Integer = Form1.Det_NumDerAccord()
        '
        Dim RépetAcc() As String = {String.Empty} ' répétition des accords
        Dim listRépet As String = Form1.Récup_Répet2
        RépetAcc = listRépet.Split(",")
        Dim Répet As Integer = Convert.ToInt16(RépetAcc(UBound(RépetAcc))) ' on va chercher la dernière répétition pour détermin la variable 'fin' juste après
        '
        fin = (fin + (Répet - 1)) * 16
        'Grid1.Range(5, 1, Grid1.Rows - 1, fin).BackColor = Color.White
        '
        For i = Grid1.FixedRows To Grid1.Rows - 1
            For j = 1 To fin
                Grid1.Cell(i, j).BackColor = Color.White
            Next
        Next
    End Sub
    Sub Clear_graphique2() ' cette méthode ne doit pas être utiliser par F1_Refresh. La différence se situe sur la manière de trouver le dernier accord.
        Dim i, j As Integer
        Dim fin As Integer = Det_NumDerAccord2() ' cette méthode appartient à la classe PianoTRoll contrairement à form1.Det_NumDerAccord utilisée dans Clear_graphique
        '
        If fin <> -1 Then
            Grid1.AutoRedraw = False
            For i = Grid1.FixedRows To Grid1.Rows - 1
                For j = 1 To fin
                    Grid1.Cell(i, j).BackColor = Color.White
                Next
            Next
            Grid1.AutoRedraw = True
            Grid1.Refresh()
        End If
    End Sub
    Sub Graphique_Tonique()
        Dim j As Integer
        Dim ColDeb, ColFin As Integer
        Dim Tonique As String
        '
        If Not Module1.EnChargement Then
            j = (Form1.Det_DerEventH2) * 16 'dernière colonne à traiter graphiquement
            If Trim(Me.ListAcc) <> "" Then
                ColDeb = 1
                ColFin = j 'Grid1.Cols - 1
                '
                'Tonique = Trim(ListTonNotes.Text
                Tonique = ListPédale(PédaleLocale)
                Notes_Tonique(Tonique, ColDeb, ColFin) ' surlignage de la tonique
            End If
        End If
    End Sub
    Sub Graphique_Tonique_old()
        Dim i, j As Integer
        Dim tbl(), tbl1(), tbl2() As String
        Dim ColDeb, ColFin As Integer
        Dim Tonique As String
        '
        If Not Module1.EnChargement Then
            j = (Form1.Det_DerEventH2) * 16 'dernière colonne à traiter graphiquement
            tbl = Me.ListAcc.Split(";")
            If Trim(Me.ListAcc) <> "" Then
                tbl = Me.ListAcc.Split(";")
                For i = 0 To UBound(tbl)
                    tbl1 = Split(tbl(i), "-")
                    ' début du calque MIDI en N° de colonnes
                    ColDeb = PosColonnes(tbl1(0))
                    ' fin du calque MIDI en N° de colonnes
                    If i <> UBound(tbl) Then
                        tbl2 = Split(tbl(i + 1), "-")
                        ColFin = PosColonnes(tbl2(0))
                    Else
                        ColFin = j 'Grid1.Cols - 1
                    End If
                    '
                    Tonique = Trim(ListTonNotes.Text)
                    Notes_Tonique(Tonique, ColDeb + 1, ColFin)
                    '
                Next
            End If
        End If
    End Sub
    Function Det_NumDerAccord2() As Integer
        Dim i As Integer = ((nbMesures) * 16) - 15
        Det_NumDerAccord2 = -1 ' cas où pas d'accord trouvé
        While Trim(Grid1.Cell(2, i).Text) = "" And i >= 1
            i = i - 16
        End While
        If i >= 1 Then Det_NumDerAccord2 = i + 15
    End Function
    Public Sub Clear_AllLayers()
        Grid1.AutoRedraw = False
        Grid1.Range(Grid1.FixedRows - 3, 1, Grid1.FixedRows - 3, Grid1.Cols - 1).BackColor = Grid1.BackColorFixed
        Grid1.Range(Grid1.FixedRows - 3, 1, Grid1.FixedRows - 3, Grid1.Cols - 1).ForeColor = Color.Black

        Grid1.Range(Grid1.FixedRows - 1, 1, Grid1.FixedRows - 1, Grid1.Cols - 1).BackColor = Color.Lavender
        Grid1.Range(Grid1.FixedRows + 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).BackColor = Color.White
        Grid1.AutoRedraw = True
        Grid1.Refresh()
    End Sub

    Sub Graphique_Accords()
        Dim i As Integer
        Dim tbl(), tbl2() As String
        '
        tbl = Me.ListAcc.Split(";")
        If Trim(Me.ListAcc) <> "" Then
            For i = 0 To UBound(tbl)
                tbl2 = tbl(i).Split("-")
                Notes_Acc(i, Trim(tbl2(1))) ' traçage du calque MIDI pour les accords
            Next
        End If
    End Sub
    Private Function PosColonnes(mesure As String) As Integer
        Dim tbl() As String
        Dim m, t, ct As Integer
        Dim m1, t1, ct1 As Integer


        tbl = Split(mesure, ".")
        '
        m1 = Convert.ToInt16(tbl(0))
        t1 = Convert.ToInt16(tbl(1))
        ct1 = Convert.ToInt16(tbl(2))
        '
        m = (m1 - 1) * 16
        t = (t1 - 1) * 4
        ct = (ct1 - 1) * 2
        '
        Return m + t + ct

    End Function

    Private Sub Notes_Tonique(Tonique As String, ColDeb As Integer, ColFin As Integer)
        Dim i As Integer
        Dim tbl() As String
        Dim Ligne As Integer
        '
        For i = 0 To ValNoteCubase_2.Count - 1
            tbl = ValNoteCubase_2(i).Split(" ")
            'For j = 0 To UBound(tbl1)
            If Tonique = tbl(0) Then ' UCase(tbl1(j)) = tbl2(0) Then
                Ligne = (Grid1.Rows - 1) - i
                Grid1.Range(Ligne, ColDeb, Ligne, ColFin).BackColor = CoulCalqPed
            End If
            'Next
        Next
    End Sub






    Public Sub Graphique_Drums()
        Dim a As String
        Dim tbl() As String
        Dim Lignedeb As Integer = Grid1.FixedRows
        Dim Lignefin As Integer = Grid1.Rows - 1

        Grid1.AutoRedraw = False
        '
        If CheckDrum.Checked Then
            a = Form1.Drums.ChaineModèle()
            If Trim(a) <> "" Then
                tbl = a.Split()
                For i = 0 To tbl.Count - 1
                    Grid1.Range(Lignedeb, Convert.ToInt16(tbl(i)), Lignefin, Convert.ToInt16(tbl(i))).BackColor = Module1.Couleur_ModelDrum
                Next
            End If
        End If
        '
        Grid1.AutoRedraw = True
        Grid1.Refresh()
    End Sub

    Sub Notes_Acc(NAcc As Integer, Acc As String)
        Dim i, j As Integer
        Dim a As String
        Dim tbl() As String
        Dim tbl1() As String
        Dim tbl2() As String
        Dim Notes_Acc1 As String
        Dim Ligne, ColDéb, ColFin As Integer
        '
        tbl = Acc.Split()
        a = Trim(Form1.TradD(tbl(0)))  ' les accords doivent être exprimés en #
        If UBound(tbl) = 0 Then
            Acc = a
        Else
            Acc = Trim(a + " " + tbl(1))
        End If
        Notes_Acc1 = Form1.Det_NotesAccord3(Acc, "#")
        tbl1 = Notes_Acc1.Split("-")
        '
        For i = 0 To ValNoteCubase_2.Count - 1
            tbl2 = ValNoteCubase_2(i).Split(" ")
            For j = 0 To UBound(tbl1)
                If UCase(tbl1(j)) = tbl2(0) Then
                    Ligne = (Grid1.Rows - 1) - i
                    ColDéb = ((NAcc) * 16) + 1
                    ColFin = ColDéb + 15
                    Grid1.Range(Ligne, ColDéb, Ligne, ColFin).BackColor = CoulCalqAcc
                    Exit For
                End If
            Next
        Next
    End Sub
    Sub Graphique_Tons() 'calque midi
        Dim i, j As Integer
        Dim tbl(), tbl1(), tbl2() As String
        Dim ColDeb, ColFin As Integer
        '
        If Not Module1.EnChargement Then
            j = (Form1.Det_DerEventH2) * 16 'dernière colonne à traiter graphiquement
            If Trim(Me.ListTon) <> "" Then
                tbl = Me.ListTon.Split(";")
                For i = 0 To UBound(tbl)
                    tbl1 = Split(tbl(i), "-")
                    ' début du calque MIDI en N° de colonnes
                    ColDeb = PosColonnes(tbl1(0))
                    ' fin du calque MIDI en N° de colonnes
                    If i <> UBound(tbl) Then
                        tbl2 = Split(tbl(i + 1), "-")
                        ColFin = PosColonnes(tbl2(0))
                    Else
                        ColFin = j  'Grid1.Cols - 1
                    End If
                    Notes_Ton(tbl1(1), ColDeb + 1, ColFin)
                Next
            End If
        End If
    End Sub
    Sub Graphique_Gammes() 'calque midi
        Dim i As Integer
        Dim tbl() As String

        'Grid1.AutoRedraw = False
        '
        If Trim(Me.ListGam) <> "" Then
            tbl = Me.ListGam.Split("-")
            For i = 0 To UBound(tbl)
                Notes_Gamme(i, tbl(i))
            Next
            '
        End If
        'Grid1.AutoRedraw = True
        'Grid1.Refresh()
    End Sub
    Sub Notes_Gamme(NAcc As Integer, Gamme As String)
        Dim i, j As Integer
        Dim tbl1() As String
        Dim tbl2() As String
        Dim Notes_Gamme1 As String
        Dim Ligne, ColDéb, ColFin As Integer
        Dim oo1 As New RechercheG_v2()
        Dim a As String = My.Resources.FichierListeGammes_Cubasev4
        oo1.GammesBases = a ' mise à jour du fichier global des gammes 
        'a = My.Resources.FichierListeAccords
        'oo1.AccordsBases = a ' mise à jour du fichier global des accords
        '
        Gamme = Trad_BemDies(Gamme)
        Notes_Gamme1 = UCase(oo1.Det_NotesGammes3(Gamme))
        tbl1 = Notes_Gamme1.Split("-")
        '
        For i = 0 To ValNoteCubase_2.Count - 1
            tbl2 = ValNoteCubase_2(i).Split(" ")
            For j = 0 To UBound(tbl1)
                If tbl1(j) = tbl2(0) Then
                    Ligne = (Grid1.Rows - 1) - i
                    ColDéb = ((NAcc) * 16) + 1
                    ColFin = ColDéb + 15
                    Grid1.Range(Ligne, ColDéb, Ligne, ColFin).BackColor = CoulCalqGammes
                    Exit For
                End If
            Next
        Next
    End Sub

    Sub Notes_Ton(Gamme As String, ColDeb As Integer, colFin As Integer)
        Dim i, j As Integer
        Dim tbl1() As String
        Dim tbl2() As String
        Dim Notes_Gamme1 As String
        Dim Ligne As Integer
        Dim oo1 As New RechercheG_v2()
        Dim a As String = My.Resources.FichierListeGammes

        'oo1.GammesBases = a ' mise à jour du fichier global des gammes 
        'a = My.Resources.FichierListeAccords
        'oo1.AccordsBases = a ' mise à jour du fichier global des accords
        '
        Gamme = Trad_BemDies(Gamme)
        tbl1 = Gamme.Split()
        '
        Notes_Gamme1 = UCase(oo1.Det_NotesGammes3(Gamme))
        tbl1 = Notes_Gamme1.Split("-")
        '
        For i = 0 To ValNoteCubase_2.Count - 1
            tbl2 = ValNoteCubase_2(i).Split(" ")
            For j = 0 To UBound(tbl1)
                If tbl1(j) = tbl2(0) Then
                    Ligne = (Grid1.Rows - 1) - i
                    Grid1.Range(Ligne, ColDeb, Ligne, colFin).BackColor = CoulCalqTon
                    Exit For
                End If
            Next
        Next
        '
    End Sub
    Private Function Trad_BemDies(gam As String) As String
        Dim tbl() As String = gam.Split()

        Select Case tbl(0)
            Case "Db"
                Return "C#" + " " + Trim(tbl(1))
            Case "Eb"
                Return "D#" + " " + Trim(tbl(1))
            Case "Gb"
                Return "F#" + " " + Trim(tbl(1))
            Case "Ab"
                Return "G#" + " " + Trim(tbl(1))
            Case "Bb"
                Return "A#" + " " + Trim(tbl(1))
            Case Else
                Return gam
        End Select
    End Function
    Sub AddMarq(Marq As String, Position As String)
        Dim tbl() As String
        Dim Pos As Integer

        tbl = Position.Split(".")
        Pos = Convert.ToInt16(tbl(0))
        '
        Pos = ((Pos - 1) * Me.DivisionMes) + 1
        '
        Grid1.Cell(3, Pos).Text = Trim(Marq)
        Grid1.Cell(3, Pos).Alignment = AlignmentEnum.LeftCenter
        Grid1.Cell(3, Pos).FontBold = True

        Grid1.Cell(3, Pos).BackColor = Color.AliceBlue
        Grid1.Cell(3, Pos).ForeColor = Color.Red
    End Sub
    Sub DeleteMarq(Position As String)
        Dim tbl() As String
        Dim Pos As Integer

        tbl = Position.Split(".")
        Pos = Convert.ToInt16(tbl(0))
        '
        Pos = ((Pos - 1) * Me.DivisionMes) + 1
        '
        Grid1.Cell(3, Pos).Text = "" ' effacer le marqueur
        Grid1.Cell(3, Pos).Alignment = AlignmentEnum.LeftCenter
        Grid1.Cell(3, Pos).FontBold = True

        Grid1.Cell(3, Pos).BackColor = Color.FromArgb(240, 240, 240) ' gris clair
        Grid1.Cell(3, Pos).ForeColor = Color.Black
        '
        'Dim a As String = Form1.Det_ListGam()
        'tbl2 = a.Split("-")
        ''
        'tbl = Position.Split(".")
        'Pos2 = Convert.ToInt16(tbl(0))
        'Pos = ((Pos2 - 1) * 16) + 1
        '
        'Grid1.Cell(3, Pos).Text = Trim(tbl2(Pos2 - 1)) ' ré-écriture de la gamme
    End Sub

    ' *********************************************************
    ' UTILITAIRES
    ' *********************************************************
    Public Class NotesCCC
        Public Ecartligne As Integer
        Public EcartCol As Integer
        Public Longueur As Integer
        Public Vélo As String
    End Class
    Public Class CtrlCCC
        Public EcartCol As Integer
        Public Valeur As String
    End Class
    Private Class SauvAnnuler
        Public Ligne As Integer
        Public Colonne As Integer
        Public Longueur As Integer
        Public Vélo As String
        Public Action As ActionEnum

    End Class
    Enum ActionEnum
        Restituer
        Effacer
        EffacerVers
        EffacerRestituer
        EffacerSelRestituer ' pour la tranposition
        Restituer_DeplaverVers
    End Enum
    Enum TypeInfo
        Notes
        NotesRallg
        Pedales
        All
        Vélocités
    End Enum
    Dim Action As ActionEnum
    Public ListCCC As New List(Of NotesCCC) ' tampon utilisé pour le copier/coller des notes
    Public ListCCCctrl As New List(Of CtrlCCC) ' tampon utilisé pour le copier/coller des Ctrl
    Dim ListAnnuler As New List(Of SauvAnnuler) ' tampon utiliser pour annuler dernière opération
    ' Classe Annulation
    ' *****************
    Class Ann
        Class SauvAnnuler
            Public Ligne As Integer
            Public Colonne As Integer
            Public Longueur As Integer
            Public Vélo As String
        End Class
        Class SauvAnnulerCtrl
            Public Colonne As Integer
            Public Valeur As String
        End Class
        Public ListAnnuler As New List(Of SauvAnnuler)
        Public ListAnnulerCTRL As New List(Of SauvAnnulerCtrl)
        Public Action As ActionEnum
        Public Tinfo As TypeInfo
        Public SelDeb As Integer = -1
        Public SelFin As Integer = -1
    End Class
    '
    Dim ListAnnulation As New List(Of Ann) ' Liste pour CTRL Z 
    Dim PointAnn As Integer = -1 ' pointeur de la liste d'annulation
    Sub JouerNote(ValeurNote As Byte, Dyn As Byte)
        'Dim Canal As Byte = Microsoft.VisualBasic.Right(Trim(Me.CheckMute.Text), 1) ' N° piste, N° canal
        Try
            TimerStop.Interval = 1500
            TimerStop.Start()
            If Not (Form1.SortieMidi.Item(Form1.ChoixSortieMidi).IsOpen) Then
                Form1.SortieMidi.Item(Form1.ChoixSortieMidi).Open()
            End If
            '
            Form1.SortieMidi.Item(Form1.ChoixSortieMidi).SendNoteOn(Me.Canal, ValeurNote, Dyn)
            NoteCourante = ValeurNote
            NoteAEtéJouée = True
        Catch ex As Exception
            TimerStop.Stop()
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
        Try
            TimerStop.Stop()

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
    Sub CouperData()
        CopierData()
        EffacerSelection()
    End Sub
    Sub EffacerSelection()
        Dim i, j, k As Integer
        Dim ColFin As Integer = 0
        Dim ExitBoucle As Boolean = False

        With Grid1.Selection

            'If (Grid1.Selection.FirstRow = Grid1.FixedRows - 1 And Grid1.Selection.LastRow = Grid1.Rows - 1) Then
            'SauvPourAnnulationCouper2(.FirstRow, .LastRow, .FirstCol, .LastCol) ' on sauvegarde ici ce que l'on va couper, avant de pouvoir le restituer par CTRL=Z
            '' effacer CTRL Mod
            '' ****************
            'SauvPourAnnulationCouper2(.FirstRow, .FirstRow, .FirstCol, .LastCol) ' on sauvegarde ici ce que l'on va couper, avant de pouvoir le restituer par CTRL=Z, on ne coupe que la 1ere ligne des modulations
            'For j = .FirstCol To .LastCol
            'If IsNumeric(Grid1.Cell(Grid1.FixedRows - 1, j).Text) Then
            'Grid1.Cell(Grid1.FixedRows - 1, j).Text = ""
            'End If
            'Next
            'Else
            ' effacer CTRL Mod + Notes
            ' ************************
            ' effacer CTRL
            ' ************
            SauvPourAnnulationCouper2(.FirstRow, .LastRow, .FirstCol, .LastCol) ' on sauvegarde ici ce que l'on va couper, avant de pouvoir le restituer par CTRL=Z
            For j = .FirstCol To .LastCol
                If Trim(Grid1.Cell(Grid1.FixedRows - 1, j).Text) <> "" Then
                    Grid1.Cell(Grid1.FixedRows - 1, j).Text = ""
                End If
            Next
            ' Effacer Notes
            ' *************
            Grid1.AutoRedraw = False
            For i = .FirstRow To .LastRow
                For j = .FirstCol To .LastCol
                    If IsNumeric(Grid1.Cell(i, j).Text) Then
                        ' effacer tête de notes
                        Grid1.Cell(i, j).Text = ""
                        ' effacer longueur de la note
                        ExitBoucle = False
                        k = j + 1
                        While Grid1.Cell(i, k).Text = Trait And Not (IsNumeric(Grid1.Cell(i, k).Text)) And Not (ExitBoucle)
                            Grid1.Cell(i, k).Text = ""
                            If Trim(Grid1.Cell(i, k + 1).Text = Trait) And (k + 1 <= Grid1.Cols - 1) Then
                                k = k + 1
                            Else
                                ExitBoucle = True
                            End If
                        End While
                    End If
                Next j
            Next
            Grid1.AutoRedraw = True
            Grid1.Refresh()
            'End If
        End With
    End Sub
    '
    Sub EffacerSelection2()
        Dim i, j, k As Integer
        Dim ColFin As Integer = 0
        Dim ExitBoucle As Boolean = False

        With Grid1.Selection
            SauvPourAnnulationCouper2(.FirstRow, .LastRow, .FirstCol, .LastCol) ' on sauvegarde ici ce que l'on va couper, avant de pouvoir le restituer par CTRL=Z
            For i = .FirstRow To .LastRow
                For j = .FirstCol To .LastCol
                    If IsNumeric(Grid1.Cell(i, j).Text) Then
                        ' effacer tête de notes
                        Grid1.Cell(i, j).Text = ""
                        ' effacer longueur de la note
                        k = j + 1
                        While Grid1.Cell(i, k).Text = Trait And Not (IsNumeric(Grid1.Cell(i, k).Text)) And Not (ExitBoucle)
                            Grid1.Cell(i, k).Text = ""
                            If Trim(Grid1.Cell(i, k + 1).Text = Trait) And (k + 1 <= Grid1.Cols - 1) Then
                                k = k + 1
                            Else
                                ExitBoucle = True
                            End If
                        End While
                    End If
                Next
            Next
        End With
    End Sub
    Sub SauvPourAnnulationCouper(ligneDeb As Integer, ligneFin As Integer, colDeb As Integer, ColFin As Integer)
        Dim i, j As Integer
        ListAnnuler.Clear()

        Action = ActionEnum.Restituer
        For i = ligneDeb To ligneFin
            For j = colDeb To ColFin
                If IsNumeric(Grid1.Cell(i, j).Text) Then
                    Dim oo As New SauvAnnuler With {
                        .Ligne = i,
                        .Colonne = j
                    }
                    oo.Longueur = Det_LongueurNote(oo.Ligne, oo.Colonne)
                    oo.Vélo = Grid1.Cell(i, j).Text
                    ListAnnuler.Add(oo)
                End If
            Next j
        Next i
    End Sub
    Sub SauvPourAnnulationCouper2(ligneDeb As Integer, ligneFin As Integer, colDeb As Integer, ColFin As Integer)
        Dim i, j As Integer

        ' Init tampon d'annulation
        ' ************************
        Dim aa As New Ann
        '
        ListAnnulation.Add(aa)
        PointAnn = PointAnn + 1
        ListAnnulation.Item(PointAnn).Action = ActionEnum.Restituer
        '
        ' Sauvegarde des Notes
        ' ********************
        For i = ligneDeb To ligneFin
            For j = colDeb To ColFin
                If IsNumeric(Grid1.Cell(i, j).Text) Then
                    Dim oo As New Ann.SauvAnnuler With {
                        .Ligne = i,
                        .Colonne = j
                    }
                    oo.Longueur = Det_LongueurNote(oo.Ligne, oo.Colonne)
                    oo.Vélo = Grid1.Cell(i, j).Text
                    ListAnnulation.Item(PointAnn).ListAnnuler.Add(oo)
                    ListAnnulation.Item(PointAnn).Tinfo = TypeInfo.All
                    ListAnnulation.Item(PointAnn).Action = ActionEnum.EffacerRestituer
                End If
            Next j
        Next i
        '
        ' Sauvegarde des pédales
        ' **********************
        For j = colDeb To ColFin
            If Trim(Grid1.Cell(Grid1.FixedRows - 1, j).Text) <> "" Then
                Dim ooo As New Ann.SauvAnnulerCtrl With {
                    .Colonne = j,
                    .Valeur = Trim(Grid1.Cell(Grid1.FixedRows - 1, j).Text)
                }
                ListAnnulation.Item(PointAnn).ListAnnulerCTRL.Add(ooo)
                ListAnnulation.Item(PointAnn).Tinfo = TypeInfo.All
                ListAnnulation.Item(PointAnn).Action = ActionEnum.EffacerRestituer
            End If
        Next
        '
    End Sub
    Sub CopierData()
        Dim i, j As Integer
        Dim ColInitial As Integer
        Dim LigneInitial As Integer
        With Grid1.Selection


            ListCCC.Clear()
            ListCCCctrl.Clear()
            '
            ColInitial = .FirstCol
            LigneInitial = .FirstRow
            '

            For j = .FirstCol To .LastCol
                If Trim(Grid1.Cell(Grid1.FixedRows - 1, j).Text) <> "" Then 'présence P+ ou P-
                    Dim ooo As New CtrlCCC With {
                            .EcartCol = j - ColInitial,
                            .Valeur = Trim(Grid1.Cell(Grid1.FixedRows - 1, j).Text)
                            }
                    ListCCCctrl.Add(ooo)
                End If
                For i = .FirstRow To .LastRow ' présence d'une note
                    If IsNumeric(Grid1.Cell(i, j).Text) Then
                        Dim oo As New NotesCCC With {
                                .EcartCol = j - ColInitial,
                                .Ecartligne = i - LigneInitial,
                                .Vélo = Trim(Grid1.Cell(i, j).Text),
                                .Longueur = Convert.ToInt32(Det_LongueurNote(i, j))                      '
                                }
                        ListCCC.Add(oo)
                    End If
                Next i
            Next j
            'End If
        End With
    End Sub
    Sub CollerData2()
        Dim i, j, k, m As Integer
        Dim LigneInitial, ColonneInitial As Integer
        Dim aa As New Ann
        '
        ' Préparation pour CTRL+Z
        ' ***********************
        ListAnnulation.Clear()
        PointAnn = -1
        '
        ListAnnulation.Add(aa)
        PointAnn = PointAnn + 1
        ListAnnulation.Item(PointAnn).Action = ActionEnum.Effacer

        ' Collage Notes
        ' *************
        LigneInitial = Grid1.Selection.FirstRow
        ColonneInitial = Grid1.Selection.FirstCol
        '
        For Each a As NotesCCC In ListCCC
            Dim oo As New Ann.SauvAnnuler
            i = LigneInitial + a.Ecartligne
            j = ColonneInitial + a.EcartCol
            If i <= Grid1.Rows - 1 And j <= Grid1.Cols - 1 Then
                oo.Ligne = i
                oo.Colonne = j
                oo.Vélo = a.Vélo
                '
                Grid1.Cell(i, j).Text = a.Vélo
                Grid1.Cell(i, j).FontBold = True
                m = 1
                For k = j + 1 To j + a.Longueur - 1
                    If k <= Grid1.Cols - 1 Then
                        If IsNumeric(Grid1.Cell(i, k).Text) = False Then
                            m = m + 1
                            Grid1.Cell(i, k).Text = Trait
                            Grid1.Cell(i, k).FontBold = True
                        Else
                            Exit For
                        End If
                    Else
                        Exit For
                    End If
                Next k
                oo.Longueur = m
                ListAnnulation.Item(PointAnn).ListAnnuler.Add(oo) ' sauvegarde pour CTRL+Z
                ListAnnulation.Item(PointAnn).Tinfo = TypeInfo.All
                'ListAnnulation.Item(PointAnn).Action = ActionEnum.EffacerRestituer
            End If
        Next
        '
        ' Collage CTRL des pédales
        ' ************************
        For Each a As CtrlCCC In ListCCCctrl
            Dim ooo As New Ann.SauvAnnulerCtrl
            i = Grid1.FixedRows - 1
            j = ColonneInitial + a.EcartCol
            '
            ooo.Colonne = j
            If j <= Grid1.Cols - 1 Then
                Grid1.Cell(i, j).Text = a.Valeur
                ooo.Valeur = a.Valeur
                ListAnnulation.Item(PointAnn).ListAnnulerCTRL.Add(ooo) ' ' sauvegarde pour CTRL+Z
                ListAnnulation.Item(PointAnn).Tinfo = TypeInfo.All
            End If
        Next
        '
    End Sub
    Sub CollerData()
        Dim i, j, k, m As Integer
        Dim LigneInitial, ColonneInitial As Integer

        Action = ActionEnum.Effacer
        ListAnnuler.Clear()

        LigneInitial = Grid1.Selection.FirstRow
        ColonneInitial = Grid1.Selection.FirstCol
        '
        For Each a As NotesCCC In ListCCC
            Dim oo As New SauvAnnuler
            i = LigneInitial + a.Ecartligne
            j = ColonneInitial + a.EcartCol
            If i <= Grid1.Rows - 1 And j <= Grid1.Cols - 1 Then
                oo.Ligne = i
                oo.Colonne = j
                Grid1.Cell(i, j).Text = a.Vélo
                Grid1.Cell(i, j).FontBold = True
                oo.Vélo = a.Vélo
                m = 1
                For k = j + 1 To j + a.Longueur - 1
                    If k <= Grid1.Cols - 1 Then
                        If IsNumeric(Grid1.Cell(i, k).Text) = False Then
                            m = m + 1
                            Grid1.Cell(i, k).Text = Trait
                            Grid1.Cell(i, k).FontBold = True
                        Else
                            Exit For
                        End If
                    Else
                        Exit For
                    End If
                Next k
                oo.Longueur = m
                ListAnnuler.Add(oo)
            End If
        Next
    End Sub
    Sub AnnulerData()
        Select Case Action
            Case ActionEnum.Restituer
                For Each oo As SauvAnnuler In ListAnnuler
                    With oo
                        Grid1.Cell(.Ligne, .Colonne).Text = .Vélo
                        Dim i As Integer = .Longueur - 1
                        Dim j As Integer = .Colonne + 1
                        While i <> 0
                            Grid1.Cell(.Ligne, j).Text = Trait
                            j = j + 1
                            i = i - 1
                        End While
                    End With
                Next
            Case ActionEnum.Effacer
                For Each oo As SauvAnnuler In ListAnnuler
                    With oo
                        Grid1.Cell(.Ligne, .Colonne).Text = ""
                        Dim i As Integer = .Longueur - 1
                        Dim j As Integer = .Colonne + 1
                        While i <> 0
                            Grid1.Cell(.Ligne, j).Text = ""
                            j = j + 1
                            i = i - 1
                        End While
                    End With
                Next
        End Select

    End Sub
    Sub AnnulerData2()
        If PointAnn <> -1 Then
            Select Case ListAnnulation.Item(PointAnn).Action
                Case ActionEnum.Restituer, ActionEnum.Restituer_DeplaverVers
                    If ListAnnulation.Item(PointAnn).Tinfo = TypeInfo.Notes Or ListAnnulation.Item(PointAnn).Tinfo = TypeInfo.All Then
                        For Each oo As Ann.SauvAnnuler In ListAnnulation.Item(PointAnn).ListAnnuler
                            With oo
                                Grid1.Cell(.Ligne, .Colonne).Text = .Vélo
                                Dim i As Integer = .Longueur - 1
                                Dim j As Integer = .Colonne + 1
                                While i <> 0
                                    Grid1.Cell(.Ligne, j).Text = Trait
                                    j = j + 1
                                    i = i - 1
                                End While
                            End With
                        Next
                    Else
                        If ListAnnulation.Item(PointAnn).Tinfo = TypeInfo.NotesRallg Then
                            For Each oo As Ann.SauvAnnuler In ListAnnulation.Item(PointAnn).ListAnnuler
                                With oo
                                    Dim i As Integer = .Ligne
                                    Dim j As Integer = .Colonne
                                    ' effacer la note
                                    Do
                                        Grid1.Cell(i, j).Text = ""
                                        j = j + 1
                                    Loop Until Trim(Grid1.Cell(.Ligne, j).Text) <> Trait Or j > Grid1.Cols - 1

                                    ' écrire ancienne note
                                    i = .Ligne
                                    j = .Colonne + 1
                                    Grid1.Cell(.Ligne, .Colonne).Text = .Vélo
                                    For j = .Colonne + 1 To (.Colonne + .Longueur) - 1
                                        Grid1.Cell(.Ligne, j).Text = Trait
                                    Next
                                End With
                            Next
                        Else
                            If ListAnnulation.Item(PointAnn).Tinfo = TypeInfo.Vélocités Then
                                For Each oo As Ann.SauvAnnuler In ListAnnulation.Item(PointAnn).ListAnnuler
                                    With oo
                                        Dim i As Integer = .Ligne
                                        Dim j As Integer = .Colonne
                                        ' écrire ancienne note
                                        i = .Ligne
                                        j = .Colonne + 1
                                        Grid1.Cell(.Ligne, .Colonne).Text = .Vélo
                                    End With
                                Next
                            End If
                        End If
                    End If
                    '
                    If ListAnnulation.Item(PointAnn).Action = ActionEnum.Restituer_DeplaverVers Then
                        PointAnn = PointAnn - 1
                        If PointAnn > -1 Then
                            For Each oo As Ann.SauvAnnuler In ListAnnulation.Item(PointAnn).ListAnnuler
                                With oo
                                    Grid1.Cell(.Ligne, .Colonne).Text = ""
                                    Dim i As Integer = .Longueur - 1
                                    Dim j As Integer = .Colonne + 1
                                    While i <> 0
                                        Grid1.Cell(.Ligne, j).Text = ""
                                        j = j + 1
                                        i = i - 1
                                    End While
                                End With
                            Next
                        End If

                    End If
                    If ListAnnulation.Item(PointAnn).Tinfo = TypeInfo.Pedales Or ListAnnulation.Item(PointAnn).Tinfo = TypeInfo.All Then
                        For Each ooo As Ann.SauvAnnulerCtrl In ListAnnulation.Item(PointAnn).ListAnnulerCTRL
                            With ooo
                                Grid1.Cell(Grid1.FixedRows - 1, .Colonne).Text = ooo.Valeur
                            End With
                        Next
                    End If
                    PointAnn = PointAnn - 1
                Case ActionEnum.Effacer
                    If ListAnnulation.Item(PointAnn).Tinfo = TypeInfo.Notes Or ListAnnulation.Item(PointAnn).Tinfo = TypeInfo.All Then
                        For Each oo As Ann.SauvAnnuler In ListAnnulation.Item(PointAnn).ListAnnuler
                            With oo
                                Grid1.Cell(.Ligne, .Colonne).Text = ""
                                Dim i As Integer = .Longueur - 1
                                Dim j As Integer = .Colonne + 1
                                While i <> 0
                                    Grid1.Cell(.Ligne, j).Text = ""
                                    j = j + 1
                                    i = i - 1
                                End While
                            End With
                        Next
                    End If
                Case ActionEnum.EffacerVers
                    If ListAnnulation.Item(PointAnn).Tinfo = TypeInfo.Notes Or ListAnnulation.Item(PointAnn).Tinfo = TypeInfo.All Then
                        For Each oo As Ann.SauvAnnuler In ListAnnulation.Item(PointAnn).ListAnnuler
                            With oo
                                Grid1.Cell(.Ligne, .Colonne).Text = ""
                                Dim i As Integer = .Longueur - 1
                                Dim j As Integer = .Colonne + 1
                                While i <> 0
                                    Grid1.Cell(.Ligne, j).Text = ""
                                    j = j + 1
                                    i = i - 1
                                End While
                            End With
                        Next
                    End If
                    If ListAnnulation.Item(PointAnn).Tinfo = TypeInfo.Pedales Or ListAnnulation.Item(PointAnn).Tinfo = TypeInfo.All Then
                        For Each ooo As Ann.SauvAnnulerCtrl In ListAnnulation.Item(PointAnn).ListAnnulerCTRL
                            With ooo
                                Grid1.Cell(Grid1.FixedRows - 1, .Colonne).Text = ""
                            End With
                        Next
                    End If
                    PointAnn = PointAnn - 1

                Case ActionEnum.EffacerRestituer
                    If ListAnnulation.Item(PointAnn).Tinfo = TypeInfo.Notes Or ListAnnulation.Item(PointAnn).Tinfo = TypeInfo.All Then
                        For Each oo As Ann.SauvAnnuler In ListAnnulation.Item(PointAnn).ListAnnuler
                            With oo
                                'effacer : on efface ici seulement pour faire de la place à la note que l'on va restituer au cas où une autre déjà présente
                                Dim jj As Integer = .Colonne
                                Do
                                    Grid1.Cell(.Ligne, jj).Text = ""
                                    jj = jj + 1
                                Loop Until Trim(Grid1.Cell(.Ligne, jj).Text = "") Or IsNumeric(Grid1.Cell(.Ligne, jj).Text) Or jj = Grid1.Cols - 1
                                ' restituer
                                Grid1.Cell(.Ligne, .Colonne).Text = .Vélo
                                Dim i As Integer = .Longueur - 1
                                Dim j As Integer = .Colonne + 1
                                While i <> 0
                                    Grid1.Cell(.Ligne, j).Text = Trait
                                    j = j + 1
                                    i = i - 1
                                End While
                            End With
                        Next
                    End If
                    If ListAnnulation.Item(PointAnn).Tinfo = TypeInfo.Pedales Or ListAnnulation.Item(PointAnn).Tinfo = TypeInfo.All Then
                        For Each ooo As Ann.SauvAnnulerCtrl In ListAnnulation.Item(PointAnn).ListAnnulerCTRL
                            With ooo
                                Grid1.Cell(Grid1.FixedRows - 1, .Colonne).Text = ""
                                Grid1.Cell(Grid1.FixedRows - 1, .Colonne).Text = ooo.Valeur
                            End With
                        Next
                    End If
                    PointAnn = PointAnn - 1

                Case ActionEnum.EffacerSelRestituer

                    If ListAnnulation.Item(PointAnn).Tinfo = TypeInfo.Notes Or ListAnnulation.Item(PointAnn).Tinfo = TypeInfo.All Then
                        ' effacer : effacer la sélection concernée
                        Dim Coldeb As Integer = Me.DivisionMes * (ListAnnulation.Item(PointAnn).SelDeb - 1) + 1
                        Dim Colfin As Integer = Me.DivisionMes * (ListAnnulation.Item(PointAnn).SelFin) ' on ne retire pas -1 à FIN car la dernière mesure est comprise dans l'effacement
                        For i = Grid1.FixedRows To Grid1.Rows - 1
                            For j = Coldeb To Colfin
                                Grid1.Cell(i, j).Text = ""
                            Next
                        Next i
                        ' restituer
                        For Each oo As Ann.SauvAnnuler In ListAnnulation.Item(PointAnn).ListAnnuler
                            With oo
                                Grid1.Cell(.Ligne, .Colonne).Text = .Vélo
                                Dim i As Integer = .Longueur - 1
                                Dim j As Integer = .Colonne + 1
                                While i <> 0
                                    Grid1.Cell(.Ligne, j).Text = Trait
                                    j = j + 1
                                    i = i - 1
                                End While
                            End With
                        Next
                    End If



            End Select
            '
            ' Réglage de l'annulation à un niveau
            ' ***********************************
            ListAnnulation.Clear()
            PointAnn = -1
            'End If
        End If
    End Sub
    Private Sub Send_Pano(t As String)

        Dim pano As Byte

        Select Case t
            Case "L", "G"
                pano = 0
            Case "C"
                pano = 64
            Case "R", "D"
                pano = 127
            Case Else
                pano = 0
        End Select
        '
        If Chargé = True Then
            'n = Microsoft.VisualBasic.Right(Trim(Me.CheckMute.Text), 1) ' N° piste, N° canal
            'canal = Val(n) - 1
            If Not (Form1.SortieMidi.Item(Form1.ChoixSortieMidi).IsOpen) Then
                Form1.SortieMidi.Item(Form1.ChoixSortieMidi).Open()
            End If
            Form1.SortieMidi.Item(Form1.ChoixSortieMidi).SendControlChange(Me.Canal, 10, pano)
        End If

    End Sub
    Private Function Det_LongueurNote(Ligne As Integer, Coldeb As Integer) As String
        Dim col As Integer = Coldeb
        Dim i As Integer = 0
        If Coldeb < Grid1.Cols - 1 Then
            Do
                i = i + 1
                col = col + 1
            Loop Until Trim(Grid1.Cell(Ligne, col).Text) = "" Or Information.IsNumeric(Grid1.Cell(Ligne, col).Text) = True Or col = (Grid1.Cols - 1)

        End If
        Return Convert.ToString(i)
    End Function
    Private Function Det_LongueurNote2(Ligne As Integer, Coldeb As Integer, Dercolonne As Integer) As String
        Dim col As Integer = Coldeb
        Dim i As Integer = 0
        Do
            i = i + 1
            col = col + 1
        Loop Until Trim(Grid1.Cell(Ligne, col).Text) = "" Or Information.IsNumeric(Grid1.Cell(Ligne, col).Text) = True Or col = (Grid1.Cols - 1) _
            Or col > Dercolonne
        Return Convert.ToString(i)
    End Function




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
    Private Sub Rallonger(Ligne As Integer, ColDeb As Integer)
        Dim j As Integer
        Dim Col As Integer = ColDeb
        ' Déterminer le point de départ vers l'arrière
        Col = Col + 1
        Do
            Col = Col - 1
        Loop Until Grid1.Cell(Ligne, Col).Text = Trait Or Col <= 1 Or IsNumeric(Grid1.Cell(Ligne, Col).Text) = True
        If Col > 1 Then
            For j = Col + 1 To ColDeb
                Grid1.Cell(Ligne, j).Text = Trait
            Next
        End If
    End Sub
    Private Sub Raccourcir(Ligne As Integer, ColDeb As Integer)
        Dim Col As Integer = ColDeb
        Dim FlagSortir As Boolean = False

        If IsNumeric(Grid1.Cell(Ligne, Col).Text) = True Or Trim(Grid1.Cell(Ligne, Col).Text) = Trait Then
            ''
            Do Until FlagSortir = True Or Col >= Grid1.Cols
                If Trim(Grid1.Cell(Ligne, Col).Text) <> "" And IsNumeric(Grid1.Cell(Ligne, Col).Text) = False Then
                    Grid1.Cell(Ligne, Col).Text = ""
                    Col = Col + 1
                Else
                    FlagSortir = True
                End If
            Loop

        End If
    End Sub
    Private Function NbDiv(Valeur As String) As Integer
        NbDiv = 8
        Select Case Trim(Valeur)
            Case "RN", "WN"
                NbDiv = 16
            Case "BL", "HN"
                NbDiv = 8
            Case "NR", "QN"
                NbDiv = 4
            Case "CR", "EN"
                NbDiv = 2
            Case "DC", "SN"
                NbDiv = 1

        End Select

    End Function

    Private Function Det_TêtedeNote(Ligne As Integer, ColDeb As Integer) As Integer
        Dim Col As Integer = ColDeb
        ' Déterminer le point de départ vers l'arrière
        Col = Col + 1
        Do
            Col = Col - 1
        Loop Until Col <= 1 Or IsNumeric(Grid1.Cell(Ligne, Col).Text) = True
        Return Col
    End Function
    Private Function Det_NotesGammes(Gamme As String) As String
        Dim i As Integer
        Dim Tonique As String
        Dim tbl() As String

        Det_NotesGammes = ""

        ' Détermnation de la Tonique de la gamme
        ' **************************************
        tbl = Split(Trim(Gamme))
        Tonique = Trim(tbl(0))
        '
        Tonique = Form1.TradD(Tonique)
        For i = 0 To UBound(TabNotes)
            If Tonique = TabNotes(i) Then
                Exit For
            End If
        Next i

        ' Détermination des notes de la gamme dans une chaine (string)
        ' ************************************************************
        Select Case tbl(1)
            Case "Maj"
                Det_NotesGammes = Tonique + " " + TabNotes(i + 2) + " " + TabNotes(i + 4) + " " + TabNotes(i + 5) _
                    + " " + TabNotes(i + 7) + " " + TabNotes(i + 9) + " " + TabNotes(i + 11)
            Case "MinH"
                Det_NotesGammes = Tonique + " " + TabNotes(i + 2) + " " + TabNotes(i + 3) + " " + TabNotes(i + 5) _
                    + " " + TabNotes(i + 7) + " " + TabNotes(i + 8) + " " + TabNotes(i + 11)
            Case "MinM"
                Det_NotesGammes = Tonique + " " + TabNotes(i + 2) + " " + TabNotes(i + 3) + " " + TabNotes(i + 5) _
                    + " " + TabNotes(i + 7) + " " + TabNotes(i + 9) + " " + TabNotes(i + 11)
            Case "MajH"
                Det_NotesGammes = Tonique + " " + TabNotes(i + 2) + " " + TabNotes(i + 4) + " " + TabNotes(i + 5) _
                    + " " + TabNotes(i + 7) + " " + TabNotes(i + 8) + " " + TabNotes(i + 11)
            Case "PMin", "PentaMin"
                Det_NotesGammes = Tonique + " " + TabNotes(i + 3) + " " + TabNotes(i + 5) + " " + TabNotes(i + 7) + " " + TabNotes(i + 10)
            Case "Blues"
                Det_NotesGammes = Tonique + " " + TabNotes(i + 3) + " " + TabNotes(i + 5) + " " + TabNotes(i + 6) + " " + TabNotes(i + 7) + " " + TabNotes(i + 10)
            Case "Blues2"
                Det_NotesGammes = Tonique + " " + TabNotes(i + 2) + " " + TabNotes(i + 3) + " " + TabNotes(i + 4) + " " + TabNotes(i + 5) + " " + TabNotes(i + 6) + " " + TabNotes(i + 7) + " " + TabNotes(i + 9) + " " + TabNotes(i + 10)
        End Select
    End Function
    Private Sub Maj_TabNotes()
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
    Function Index_Divisions(division As Integer) As String
        Dim i As Integer = 1
        Select Case division
            Case 1
                i = 1
            Case 5
                i = 2
            Case 9
                i = 3
            Case 13
                i = 4
        End Select
        Return Convert.ToString(i)
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

    Sub Refresh_G1_Acc()
        Dim i, j As Integer
        Dim tbl() As String
        Dim tbl1() As String
        '
        ' effacer le text de la ligne entière
        If Trim(Me.ListAcc) <> "" Then
            'Grid1.Range(2, 1, 2, Grid1.Cols - 1).ClearText() ' effacer le texte de la ligne des Mesures
            'Grid1.Range(4, 1, 4, Grid1.Cols - 1).ClearText() ' effacer le texte de la ligne des Accords
            For j = 1 To Grid1.Cols - 1
                Grid1.Cell(2, j).Text = "" '  effacer le texte de la ligne des Mesures
                Grid1.Cell(4, j).Text = "" ' effacer le texte de la ligne des Accords
            Next
            ' mise à jour des accords
            tbl = Split(Me.ListAcc, ";")
            '
            ' mise à jour propriété NbAccords du PianoRoll
            Me.NbAccords = UBound(tbl) + 1
            '
            j = 1
            For i = 0 To Me.NbAccords - 1
                tbl1 = tbl(i).Split("-")
                Grid1.Cell(2, j).Text = tbl1(0)
                Grid1.Cell(2, j).Alignment = AlignmentEnum.LeftCenter
                Grid1.Cell(2, j).FontBold = True
                '
                Grid1.Cell(4, j).Text = tbl1(1)
                Grid1.Cell(4, j).Alignment = AlignmentEnum.LeftCenter
                Grid1.Cell(4, j).FontBold = True
                j = j + Me.DivisionMes
            Next
            '
            'Grid1.AutoRedraw = True
            Grid1.Refresh()
        End If
    End Sub

    Sub Refresh_G1_Gam()
        Dim i, j As Integer
        Dim tbl() As String

        '
        If Trim(Me.ListGam) <> "" Then
            'Grid1.AutoRedraw = False
            ' effacer le text de la ligne entière
            Grid1.Range(3, 1, 3, Grid1.Cols - 1).ClearText() ' effacer le texte de la ligne des gammes

            ' mise à jour des accords
            tbl = Split(Me.ListGam, "-")
            j = 1
            For i = 0 To UBound(tbl)
                Grid1.Cell(3, j).Text = tbl(i)
                Grid1.Cell(3, j).Alignment = AlignmentEnum.LeftCenter
                Grid1.Cell(3, j).FontBold = True
                j = j + Me.DivisionMes
            Next
            'Grid1.AutoRedraw = True
            'Grid1.Refresh()
        End If
    End Sub

    Sub Refresh_G1_Transf_old() ' lestransformer (Variation) + les Marqueurs

        Dim i, j As Integer
        Dim a As String
        Dim tbl() As String
        '
        Dim MarqAcc() As String = {String.Empty}
        Dim listMRQ As String = Form1.Récup_MRQ
        MarqAcc = listMRQ.Split(",")
        '
        Grid1.Range(4, 1, 4, Grid1.Cols - 1).ClearText()
        '
        a = Form1.Récup_NumMagnéto()
        tbl = a.Split(",")
        j = 1
        For i = 0 To UBound(tbl)
            Grid1.Cell(4, j).Text = Trim(MarqAcc(i) + "      " + "Variation" + Str(Val(tbl(i)) + 1))
            Grid1.Cell(4, j).FontBold = True
            Grid1.Cell(4, j).Alignment = AlignmentEnum.LeftCenter
            j = j + Me.DivisionMes
        Next

    End Sub
    Sub Refresh_G1_Transf() ' les transformer (Variation) + les Marqueurs
        Dim h, i, j, k As Integer
        Dim tbl() As String
        Dim a As String

        If Module1.EnChargement = False Then

            ' effacer anciennes infos de variation dans l'entête
            ' **************************************************
            ' Grid1.Range(4, 1, 4, Grid1.Cols - 1).ClearText() <--- ne fonctionne pas
            For j = 1 To Grid1.Cols - 1
                Grid1.Cell(4, j).Text = ""
            Next
            '
            Dim MarqAcc() As String = {String.Empty}
            Dim listMRQ As String = Form1.Récup_AllMRQ
            MarqAcc = listMRQ.Split(",")
            '
            Dim NumAcc() As String = {String.Empty}
            Dim listNumAcc As String = Form1.Récup_AllNumAcc
            NumAcc = listNumAcc.Split(",")
            '
            Dim Num_Variation() As String
            Dim listVariation As String = Form1.Récup_AllNumMagnéto()
            Num_Variation = listVariation.Split(",")
            '
            Dim Num_Répet() As String
            Dim ListRépet As String = Form1.Récup_AllRépet()
            Num_Répet = ListRépet.Split(",")
            '
            j = 1
            If Trim(listVariation) <> "" Then ' si listvariation="" --> pas d'association (accord,variation) après le "Début" (voir Récup_Magnéto)
                For i = 0 To UBound(Num_Variation)
                    j = ((Val(NumAcc(i)) - 1) * Me.DivisionMes) + 1
                    ' nom de la variation
                    Grid1.Cell(4, j).Text = Trim("Variation" + Str(Val(Num_Variation(i)) + 1))
                    Grid1.Cell(4, j).FontBold = True
                    Grid1.Cell(4, j).Alignment = AlignmentEnum.LeftCenter

                    ' cellules suivantes au cas où elles sont vides
                    For k = 0 To Val(Num_Répet(i)) - 2
                        j = ((Val(NumAcc(i)) + k) * Me.DivisionMes) + 1
                        Grid1.Cell(4, j).Text = Trim("Variation" + Str(Val(Num_Variation(i)) + 1))
                        Grid1.Cell(4, j).FontBold = True
                        Grid1.Cell(4, j).Alignment = AlignmentEnum.LeftCenter
                    Next k
                Next i
                '
                ' couleur de la variation
                Dim coul As Color
                Dim mes As Integer = ((nbMesures - 1) * Me.DivisionMes) + 1
                Dim NouvVar As Boolean = False
                For i = 0 To UBound(Num_Variation)
                    j = ((Val(NumAcc(i)) - 1) * Me.DivisionMes) + 1
                    NouvVar = False
                    a = Grid1.Cell(4, j).Text
                    If Trim(Grid1.Cell(4, j).Text) <> "" Then
                        tbl = Grid1.Cell(4, j).Text.Split()
                        coul = TabCoulMagneto(Val(tbl(1)) - 1)
                        NouvVar = True
                    End If
                    '
                    If NouvVar = True Then
                        ' maj couleur variation
                        For k = j To mes
                            Grid1.Cell(4, k).BackColor = coul
                        Next k
                        ' remettre en gris les mesures non utilisées
                        k = Form1.Det_NumDerAccord()
                        If Convert.ToInt16(NumAcc(i)) = k Then
                            h = Form1.Det_NumDerRépet() - 1
                            j = ((Val(NumAcc(i) + h) * Me.DivisionMes)) + 1
                            For k = j To mes
                                Grid1.Cell(4, k).BackColor = Color.FromArgb(244, 244, 244)
                            Next
                        End If
                        NouvVar = False
                    End If
                Next i
            End If
        End If

    End Sub
    Sub Refresh_G1_Marq()
        Dim i As Integer
        Dim tbl1() As String
        Dim tbl2() As String

        tbl1 = ListMarq.Split("&")
        If Trim(ListMarq) <> "" Then
            For i = 0 To UBound(tbl1)
                tbl2 = tbl1(i).Split(";")
                Me.AddMarq(Trim(tbl2(0)), Trim(tbl2(1)) + ".1.1")
            Next i
        End If

    End Sub

    Sub Position_Init()
        If Form1.WindowState <> FormWindowState.Minimized Then
            'taille du formulaire
            F1.Width = Form1.Width - 500
            F1.Height = Form1.Height - 78
            '
            ' Positionnement du splitter du haut(splitter Panneau1)
            Panneau1.IsSplitterFixed = False
            Panneau1.SplitterDistance = 50
            Panneau1.IsSplitterFixed = True
        End If
    End Sub
    Function Det_ValPan() As String
        Det_ValPan = 64
        If BRadio1.Checked Then
            Det_ValPan = "0"
        ElseIf BRadio2.Checked Then
            Det_ValPan = "64"
        ElseIf BRadio3.Checked Then
            Det_ValPan = "127"
        End If
    End Function
    Public Function Det_LineFirstNote(m As Integer) As Integer
        Dim i, j As Integer
        Dim k As Integer = -1
        Dim coldeb As Integer = (16 * m) - 15 '[(m-1)x16] +1 -> 16m-16 +1 -> 16m-15
        Dim colfin As Integer = coldeb + 15

        For i = 0 To 127
            For j = coldeb To colfin
                If Grid1.Cell(i, j).Text <> "" Then
                    k = i
                    Exit For
                End If

            Next
        Next
        Return k
    End Function
    Class CNote
        Public old_long As Integer
        Public old_dyn As String
        Public old_col As Integer
        Public old_row As Integer
        '
        Public new_row As Integer
    End Class
    '
    Sub QuantMel(pas As Integer)
        Dim listNotes As New List(Of CNote)
        Dim ListRow As New List(Of Integer)
        Dim ListCol As New List(Of Integer)
        Dim i, j As Integer
        Dim ii As Integer
        Dim colold As Integer
        Dim flag_sortir As Boolean = False
        '
        Dim PremRow = Det_PremRowSel()
        Dim DerRow = Det_DerRowSel()

        If (pas = -1 And PremRow > Grid1.FixedRows) Or (pas = 1 And DerRow < Grid1.Rows - 1) Then

            ' Détermination des notes sélectionnées
            ' *************************************
            With Grid1.Selection
                For i = .FirstRow To .LastRow
                    For j = .FirstCol To .LastCol
                        If Trim(Grid1.Cell(i, j).Text) <> "" And Trim(Grid1.Cell(i, j).Text) <> Trait Then
                            '
                            Dim OO As New CNote With {
                               .old_row = i,
                               .old_col = j,
                               .old_dyn = Trim(Grid1.Cell(i, j).Text),
                               .old_long = Det_LongueurNote(i, j),
                               .new_row = -1
                           }

                            ListCol.Add(OO.old_col)
                            '
                            listNotes.Add(OO)

                        End If
                    Next
                Next i
            End With
            '

            '
            ' Cacul déplacement (remontée ou descente)
            ' ****************************************

            For i = 0 To listNotes.Count - 1
                With listNotes(i)
                    ' Cacul de la nouvelle ligne
                    ' **************************
                    ii = .old_row
                    colold = .old_col
                    If Grid1.Cell(ii, .old_col).BackColor = Color.Gold Or           ' Première colonne mesure à 4/4
                      Grid1.Cell(ii, .old_col).BackColor = Color.LightSteelBlue Or ' division mesure à 4/4
                      Grid1.Cell(ii, .old_col).BackColor = Color.Olive Or          ' Première colonne mesure à 12/8
                      Grid1.Cell(ii, .old_col).BackColor = CDiv_12_8 Then          ' division mesure à 12/8

                        colold = .old_col + 1 ' on intervient ici pour éviter de tomber sur les colonnes orange (Gold) ou les colonnes bleu (LightSteelBlue) ou les colonnesvertes de 12/8
                    End If
                    Do
                        ii += pas
                    Loop Until ((Grid1.Cell(ii, colold).BackColor = CoulCalqTon) _ '(ii <= Grid1.FixedRows - 1) Or (ii >= Grid1.Rows - 1)
                       Or (Grid1.Cell(ii, colold).BackColor = CoulCalqGammes) _
                       Or (Grid1.Cell(ii, colold).BackColor = CoulCalqAcc))
                    '
                    .new_row = ii ' maj de la nouvelle ligne
                    ListRow.Add(.new_row)
                    '
                End With
            Next i

            If Not flag_sortir Then
                ' Effacer old notes position
                ' **************************
                For i = 0 To listNotes.Count - 1
                    With listNotes(i)
                        For jj = .old_col To (.old_col + (.old_long - 1))
                            Grid1.Cell(.old_row, jj).Text = ""
                        Next
                    End With
                Next i

                ' Mise à jour new notes position
                ' ******************************
                For i = 0 To listNotes.Count - 1
                    With listNotes(i)
                        ' Ecriture nouvelle position (sans la longueur)
                        ' ---------------------------------------------
                        Grid1.Cell(.new_row, .old_col).Text = .old_dyn
                        ' Ecriture des longueurs
                        ' ----------------------
                        For jj = (.old_col + 1) To ((.old_col + 1) + (.old_long - 2))
                            If Not (IsNumeric(Grid1.Cell(.new_row, jj).Text)) Then
                                Grid1.Cell(.new_row, jj).Text = Trait
                            Else
                                Exit For
                            End If
                        Next
                    End With
                Next i
                '
                ' Gestion de la sélection
                ' ***********************
                ListCol.Sort()
                ListRow.Sort()
                Grid1.Range(ListRow(0), ListCol(0), ListRow(ListRow.Count - 1), ListCol(ListCol.Count - 1)).SelectCells()
            End If
        End If
    End Sub
    Function Det_DerRowSel() As Integer
        Dim list1 As New List(Of Integer)
        list1 = ListRowSort()
        Return list1(list1.Count - 1)
    End Function
    Function Det_PremRowSel() As Integer
        Dim list1 As New List(Of Integer)
        list1 = ListRowSort()
        Return list1(0)
    End Function
    Function ListRowSort() As List(Of Integer)
        Dim i, j As Integer
        Dim list1 As New List(Of CNote)
        Dim listRow As New List(Of Integer)
        With Grid1.Selection
            For i = .FirstRow To .LastRow
                For j = .FirstCol To .LastCol
                    If Trim(Grid1.Cell(i, j).Text) <> "" And Trim(Grid1.Cell(i, j).Text) <> Trait Then
                        '
                        Dim OO As New CNote With {
                           .old_row = i,
                           .old_col = j,
                           .old_dyn = Trim(Grid1.Cell(i, j).Text),
                           .old_long = Det_LongueurNote(i, j),
                           .new_row = -1
                       }
                        listRow.Add(OO.old_row)
                        '
                    End If
                Next
            Next i
        End With
        listRow.Sort()
        Return listRow

    End Function
End Class
