
' *#######################################################################
' ##                                                                    ##
' ##                           DECLARATIONS                             ##
' ##                                                                    ##
' ########################################################################


' *********************************************
' Définition des ContextMenuStrip             *
' Menus contextuelles                         *
' *********************************************
' Dans L'onglet TONALITE : ContextMenuStrip1  *
' Dans l'onglet CADENCES : ContextMenuStrip4  *
' Dans la grille Grid2   : ContextMenuStrip3  *
' *********************************************

Imports System.ComponentModel
Imports System.IO
Imports HyperArp.Module1
Imports Microsoft.Win32

' Imports MidifileX
' variables communes au module form1
' **********************************

'Tempo_StopJeuAccord.Interval = TempoEcouteAccord
'Tempo_StopJeuAccord.Enabled = True ' démarrer la tempo d'arrêt de jeu de l'accord
'Tempo_StopJeuAccord.Start()

Public Class Form1
    Class TamponEdition
        Public Enum TGridCours
            Rien = 0
            Grid1 = 1
            Grid2 = 2
            Grid3 = 3
            TabTon = 4
            Piano = 5
        End Enum '
        Public TGrille As TGridCours
        '
        Public mDeb As Integer
        Public tDeb As Integer
        Public ctDeb As Integer
        '
        Public mFin As Integer
        Public tFin As Integer
        Public ctFin As Integer
        '
    End Class
    Class TamponCopie

        Public Actif As Boolean = False ' appartient à la derniere copie effectuée
        '
        Public Mesure_Effacer As Integer = -1
        '
        Public m As Integer
        Public t As Integer
        Public ct As Integer
        '
        Public Décalage As Integer ' Décalage par rapport à la prmière mesure
        '
        Public Marqueur As String
        Public Tonalité As String
        Public Accord As String
        Public Gamme As String
        Public Mode As String
        Public Degré As Integer
        Public Détails As String
        '
        Public Renversement(0 To 4) As String
        Public RenvChoisi As Integer
        Public Octave As String
        Public OctaveChoisie As Integer
        '
        Public EtendreNotes(0 To 5) As String
        Public EtendreChecked(0 To 5) As Boolean

        Public Nb_Items As Integer ' nomnre items de la lsite pour une copie : valeur mise à jour dans le 1er item seulement
        '
    End Class
    Class TamponInfoGrid1
        Public Enum TGrilleCours
            Rien = 0
            Grid1 = 1
            Grid2 = 2
            Grid3 = 3
            TabTon = 4
            Piano = 5
            Grid4 = 6
        End Enum
        Public Sélection As Integer
        Public LigneCours As Integer
        Public Gammes(0 To 6) As String
        Public Marqueurs As String
        Public Grille As TGrilleCours = TGrilleCours.Rien
    End Class
    ' Fonts utilisés par form1
    ReadOnly fnt1 As New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Regular)
    ReadOnly fnt2 As New System.Drawing.Font("Calibri", 12, FontStyle.Bold)



    ' Variables globales pour le menu Edition
    ' ***************************************
    Dim TCopie As New List(Of TamponCopie)
    Dim TZAnnulationGrid1 As New TamponInfoGrid1
    Dim TZAnnulation As New List(Of TamponCopie)

    Dim OrigineEdition As Module1.GridCours ' donne la grille sur laquelles les actions copier, couper, coller sont effectués
    Dim ZAnnulation_Valide As Boolean = False
    Dim ZAnnulation_FirstCol As Integer
    '
    ' Variables permettant de retrouver la valeur de la répétition de l'ancien dernier accord
    ' ***************************************************************************************
    ' utilisé par Maj_Répétition et Maj_Répétition2, lors, respectivement, d'un annuler ou d'un couper
    Dim ReptAncienDerAcc As String
    Dim ColAncienDerAcc As Integer

    Class TT
        Public Toctave(0 To 10) As Integer
    End Class
    Structure CaractNote
        Public NoteTessiture As String
        Public NoteAccord As String
    End Structure

    ' Pour Piano
    ' **********
    Dim Tessiture(0 To 31) As CaractNote
    Dim QuiScroll As String
    Dim LabelPiano As New List(Of Label)
    Dim LabelPianoMidiIn As New List(Of Label)
    Public DicoNotes As New Dictionary(Of String, TT)
    Public Touche_CouleurPréced(0 To 127) As Color
    Dim Num_octave As Integer ' N° octave de la dernière note jouée pour positionner le clavier
    '
    ' Pour les motifs perso
    Public MotifsPerso As New List(Of FlexCell.Grid)
    Public Bcopier As New List(Of Button)
    Public Bcoller As New List(Of Button)
    Dim ClipBPerso As String = ""


    ' Pour les Pistes (au sens Variation)
    ' **********************************

    '
    ' Pour ancienne Table de mixage
    ' ********************
    Dim PistePanel As New List(Of Panel)
    Public PisteMute As New List(Of CheckBox)
    Public SelBloc As New List(Of CheckBox)
    Public SoloBout As New List(Of Button)
    'Public PisteVolume As New List(Of TrackBar)

    'Dim labelVolume As New List(Of Label)
    'Dim labelPiaRoll As New List(Of Label)
    'Dim labelAff As New List(Of Label)
    'Dim soloVolume As New List(Of CheckBox)
    'Dim soloPiste As New List(Of Button)
    '
    Public PisteSolo As Integer = -1
    Public PisteSolo2 As Integer = -1 ' ne pas supprimer : utilisé dans la nouvelle table de mixage
    Public SoloCours2 As Boolean = False
    Public CanMidiCours As Integer = -1
    '
    Dim PisteDyn As New List(Of Windows.Forms.ComboBox)
    Dim PisteDyn2 As New List(Of Windows.Forms.NumericUpDown)
    Dim labelDyn As New List(Of Label)
    '
    Dim BoutMotif As New List(Of BoutPerso)
    Dim PisteMotif As New List(Of Windows.Forms.ComboBox)
    Dim labelMotif As New List(Of Label)
    '
    Dim PisteDurée As New List(Of Windows.Forms.ComboBox)
    Dim labelDurée As New List(Of Label)
    '
    Dim PisteAccent As New List(Of Windows.Forms.ComboBox)
    Dim labelAccent As New List(Of Label)
    '
    Dim PisteSouche As New List(Of Windows.Forms.ComboBox)
    Dim labelSouche As New List(Of Label)
    '
    Dim PisteRetard As New List(Of Windows.Forms.ComboBox)
    Dim labelRetard As New List(Of Label)
    '
    Dim PisteDelay As New List(Of CheckBox)
    Dim labelDelay As New List(Of Label)
    '
    Dim PisteDébut As New List(Of CheckBox)
    Dim labelDébut As New List(Of Label)
    '
    Dim PisteOctave As New List(Of Windows.Forms.ComboBox)
    Dim labelOctave As New List(Of Label)
    '
    Dim PistePRG As New List(Of Windows.Forms.ComboBox)
    Dim labelPRG As New List(Of Label)
    '
    Dim PisteRadio1 As New List(Of RadioButton)
    Dim PisteRadio2 As New List(Of RadioButton)
    Dim PisteRadio3 As New List(Of RadioButton)
    Dim PistePan As New List(Of GroupBox)
    '
    Dim labelNPiste As New List(Of Label) ' propriété visible = false
    Public NomduSon As New List(Of TextBox)
    Dim AffNomduSon As New List(Of Label)

    '
    ' Pour Onglet Tonalité
    ' ********************
    Dim TabTonsDegrés As New List(Of Label)
    Dim TabTonsVDegrés As New List(Of Label)
    Dim TabTons As New List(Of Label)
    Dim TabTonsVoisins As New List(Of Label)
    Dim TabTonsVoisinsMarq As New List(Of CheckBox)
    Dim TabTonsTitreMode As New List(Of Label)
    Dim TabTonsVTitreMode As New List(Of Label)
    Dim Mode_Expert As Boolean
    Dim LabelCours As Integer ' N° label sur lequel on fait le clic droit
    '
    ' Pour Onglet Cadences
    ' ********************
    Dim TabCadDegrés As New List(Of Label)
    Dim TabCad As New List(Of Label)
    Dim CAD_LabelCours As Integer ' N° label sur lequel on fait le clic droit
    '
    Dim Mode_Cadence As Modes
    Dim ModeSimple_Cadence As String ' cette variable ne peut contenir que "Maj" ou "Min" mis à jour respectivement dans Combo3.selectedIndex et combo4.selectedIndex

    ' Pour Onglet Blues
    ' *****************
    Dim TabBluesDegrés As New List(Of Label)
    Dim TabBluesNMesures As New List(Of Label)
    Dim TabBlues As New List(Of Label)
    Dim BLUES_LabelCours As Integer ' N° label sur lequel on fait le clic droit

    ' Pour Boutons Zones
    ' ******************
    Dim BoutZone As New List(Of Button)
    Dim BoutZoneCouleur As Color
    '
    ' Paramètre MIDI
    ' *************
    Public SortieMidi As New List(Of Midi.OutputDevice)
    Public EntréeMidi As New List(Of Midi.InputDevice)
    Public Exception As New List(Of Midi.DeviceException)
    Public ChoixSortieMidi As Integer = 0
    Public ChoixEntréeMidi As Integer = 0
    Public HorlogeMidi As New Midi.Clock(120)
    ' Dim HorlogeMidi As New Midi.Clock(120)
    Public Horloge1 As New Midi.Clock(120)
    Public ExistenceEntréeMidi As Boolean
    '
    Public ListNumMes As New List(Of String) ' pour affichage des N° de mesures lors de l'exécution
    Dim AffNumMes As Integer = 0
    '
    Dim Vélocité As Byte = 64
    Dim Canal As Byte = 1
    '
    Dim VarCallBack As String
    '
    Public DicoAccords As New Dictionary(Of String, String)
    '
    Dim ListAccords As New List(Of String)
    Dim ListGammes As New List(Of String)
    Dim ListMesures As New List(Of String)
    Dim IndexListAccords As Integer
    Dim IndexListGammes As Integer
    Dim IndexListMesures As Integer

    ' Variables pour gestion menu Edition
    ' ***********************************
    Public ClipEdit As New List(Of EventH)
    Class ClipAcc
        Public LAcc As New List(Of EventH)
    End Class

    Dim ClipAnnuler As New List(Of ClipAcc)
    Dim IndexAnnuler As Integer = -1

    ' Classe pour Cacul chaine de notes des Motifs Perso
    ' **************************************************
    Class ChainePerso
        Public Evt As String
        Public Row As Integer
        Public Col As Integer
        Public LettreLigne As String
        Public Longueur As Integer
        Public Chaine As String
    End Class
    Dim CallB_Aff_Marqueur As Boolean = False
    Dim CallB_Aff_Tonalité As Boolean = False
    Dim CallB_Aff_Accord As Boolean = False
    Dim CallB_Aff_Gamme As Boolean = False
    Dim CallB_Aff_Mesure As Boolean = False

    Dim CallB_Aff_FIN As Boolean = False
    Dim CallB_Aff_Métronome As Boolean = False
    '
    Dim CallB_Aff_Acc As Boolean = False
    Dim CallB_Aff_NumMes As Boolean = False
    Dim NumAcc As Integer ' numéro de l'accord en cours de jeu
    '
    Dim AccordAEtéJoué As Boolean = False
    Dim AccordAEtéJoué1 As String = "Non"
    Dim AccordAEtéJoué2 As String = "Non"
    Dim AccordAEtéJoué3 As String = "Non"
    '
    Dim AccordAEtéJoué_Ext As Boolean = False
    Dim AccordAEtéAff As Boolean = False
    '
    Dim GammeAEtéJouée As Boolean = False


    Dim MesureCourante As Integer
    Dim ArrêterTimer As Boolean = False

    Dim AccordJouerPiano As New NotesJouéesPiano
    Dim NoteJouerPIano_OLdbackcolor As Color
    Dim NoteJouerPiano_OldTextNote As String
    '
    Dim OrigineAccord As Modes = Modes.Rien
    Dim Cad_OrigineAccord As Modes = Modes.Rien
    Dim DerGridCliquée As GridCours = GridCours.Rien
    '
    Dim AccordMarqué As String = ""
    Dim AccordMarquéVoisin As String = ""
    '
    Dim Numérateur As Integer
    Dim Dénominateur As Integer
    '
    Dim HelpOuvert As Boolean = False
    '
    ' Propriétés de l'Entrée des accords en glisser/déposer
    ' *****************************************************
    Dim Entrée_Accord As String
    Dim Entrée_Gamme As String
    Dim Entrée_Tonalité As String
    Dim Entrée_Mode As String
    Dim Entrée_Degré As String
    Dim Entrée_Position As String

    Dim LigneCoursGrid5 As Integer = 0 ' pour onglet modulation
    Dim Lab_1 As String
    Dim Lab_2 As String
    '

    Dim TonV As String ' utlisé dans MouseDown de TabTonVoisin ,  pour identifier la tonalité de la gamme
    ' Sélection lignes dans grid1 par barre rouge et bleue
    ' ****************************************************
    '
    Dim SélectionLignes As Integer
    '
    ' Détermination des gammes communes
    ' *********************************
    Dim TabGammesCom As New List(Of List(Of String))
    Dim GammesCom As New List(Of String)
    '
    ' Détermination de l'entrée Majeure ou mineur
    ' *******************************************
    Dim OrigineTona As String ' valeur = Maj ou min
    '
    ' Informations sur le fichier chargé en cours (mis à jour dans "Ouvrir")
    ' **********************************************************************
    ' Résumé des chamin de fichiers utilisés
    ' ======================================
    ' CheminFichierOuvrir  ' ouverture fichier *.zic3
    ' CheminFichierEnreg   ' enregistrement fichier *.zic3
    ' CheminFichierMIDI    ' export des accords au format MIDI *.mid
    ' CheminFichierCalques ' export calque MIDI *.mid
    ' CheminMarqueursMIDI  ' export des marqueurs MIDI *.mid
    '
    '
    Public CheminFichierOuvrir As String = ""
    Public FichierOuvrir As String = ""
    '
    Public CheminFichierEnreg As String = ""
    Public FichierEnreg As String = ""
    '
    Public CheminFichierMIDI As String = ""
    Public FichierMIDI As String = ""
    Public ExportCTRL As Boolean = True
    '
    Public FichierExportDoc As String = ""
    '
    Public CheminFichierCalques As String = ""
    Public FichierCalques As String = ""
    Public CheminMarqueursMIDI As String = ""
    Public CheminFichierExportDoc As String = ""
    Public IndicateurEnreg As Boolean = False
    Public NomFichier As String = ""
    Public ExporterCTRL As Boolean = False
    '
    ' Jouer un Accord sur une grille
    ' ******************************
    Dim ArrêterAccord As Boolean = False
    '
    ' Paramètres de préférences
    ' *************************
    Dim TempoEcouteAccord As Integer = 1850 ' ms
    '
    ' Paramètre pour le métronome
    ' ***************************
    Dim Beat78 As Boolean = True ' pour mesure 7/8

    ' Variable associée à Det_Présence_Marqueur
    ' *****************************************
    Dim Présence_Marqueur As Boolean = False
    '
    ' Onglet en cours
    ' ***************
    'Dim OngletCours As Integer ' 
    Public OngletCours_Edition As Integer = 0 ' utilisé dans pianoroll (donc Public)
    Dim OngletCours_HyperARP As Integer = 0 ' N° des onglets de Tabcontrol2
    '
    ' Tempo_Aff_MidiIn
    ' ****************
    Dim N_Note As Byte = 0
    Dim AfficherNote As Boolean = False
    '
    ' Sauvegarde de mousecol dans les évènements mousUp ou MouseDown
    ' **************************************************************
    Dim SauveMouseColGrid2 As Integer
    Dim SauveMouseColGrid3 As Integer
    ' Variable globale pour entréer sur clic fichier zic3
    ' ***************************************************
    Dim FichierEntréSurClic As String = ""
    ' Variable globale pour affichage des N° d'accord en cours d'exécution
    '
    Dim AffAccord As String
    Dim PointAffAccord As Integer = 0

    ' **********************************************************************************************************************************
    ' Flag utilisé par les  procédures d'écriture d'un accord par Menu ou par glisser/déposer : 
    '      - procédure d'écriture par déglisser/déposer : Grid2_DragDrop & Grid3_DragDrop --> le flag doit être = true
    '      - procédure d'écriture par menu : Ecr_AccordParMenu --> le flag doit être = true
    '  Ce Flag est utilisé dans les procédures EcritureAccordDsGrid2 et EcritureAccordDsGrid3 appelées par les proc citées précédemment
    ' ***********************************************************************************************************************************
    ' Variables pour Glisser Déposer
    ' ******************************
    Dim Valeur_Drag As String = ""
    Dim Colonne_Drag As Integer = -1
    Dim Ligne_Drag As Integer = -1
    Dim Flag_EcrDragDrop As Boolean = False
    Dim MouseIsDown As Boolean = False
    '
    ' 
    Dim FlagMode As Integer
    'Dim RepèreEcriture As String ' utilisé pour restitution des données en mode étendu (Ecriture_Entrée_Dans_Compogrid)


    ' Variables pour détermination gammes jouables sur accords (2e algorithm)
    ' ***********************************************************************
    Dim LGamMaj As New List(Of String)
    Dim LGamMinH As New List(Of String)
    Dim LGamMinM As New List(Of String)
    Dim LGamMajH As New List(Of String)
    Dim LGamPentaMin As New List(Of String)
    Dim LGamBlues As New List(Of String)
    '

    Dim LongueurCroche_ms As Integer
    Class StandLigne
        Public NLigne As Integer
        Public Position As String
        Public Marqueur As String
        Public Tonalité As String
        Public Accord As String
        Public Gamme As String
    End Class
    Dim LCaractLigne As New List(Of StandLigne)
    Dim Couleur As Color
    Dim SENSGamme As String = "Monter"
    Dim NotesAJouer As String
    Dim tbl_NotesOnG() As String
    Dim tbl_NotesOffG() As Byte
    '
    Dim tbl_NotesOnA() As String
    Dim tbl_NotesOffA() As Byte
    '
    Dim EnvAccord As Boolean
    Dim EnvGamme As Boolean
    '
    ' Variable pour accords impactés
    ' ******************************
    Dim IndexPréced As Integer
    Public Structure AccImpact
        Dim Notes As String
        Dim Accord As String
        Dim Position As String
        Dim ligne As Integer
    End Structure
    ' Onglet tabton Voisins
    ' *********************
    ' position du curseur de désignation de Do Majeur ou La Mineur
    Public Enum PosCur
        Haut = 0
        Bas = 1
    End Enum
    Dim PositionCurs As PosCur = PosCur.Haut
    Dim AjoutBas As Integer
    Dim AjoutBas2 As Integer

    ' Variables pour Sortir
    ' ********************
    Dim FermetureParQuitter As Boolean = False
    Dim FermetureEncours As Boolean = False

    ' Variable Compression dans le mixeur
    ' ***********************************
    Public ValCompress As Integer = -1
    '
    ' Utilisation d'une liste d'objet  PianoRoll
    ' ******************************************
    Public listPIANOROLL As New List(Of PianoRoll)
    Public PIANOROLLChargé As New List(Of Boolean)
    Public PIANOROLLNPiste As New List(Of Integer)
    Dim PIANOROLLNPréced As Integer = -1

    ' Variables d'entrée du formulaire Form3 : calcul interval entre 2 valeur de modulation
    ' *************************************************************************************
    Public MyPoint1 As Integer = 0
    Public MyPoint2 As Integer = 0
    Public MyInterval As Integer = 0

    '  Création DrumEdit
    '  *****************
    Public Drums As New DrumEdit(9)

    ' Création Automation
    ' *******************
    Public Automation1 As New Automation

    ' Création Table de Mixage
    ' ************************
    Public Mix As New Mixage
    ' Variable pour Jauge de calcul des notes pour PLAY
    ' *************************************************
    Dim TailleJauge As New Size ' pour la jauge

    Public WriteOnly Property PPIANOROLLChargé(NPISTE As Integer) As Boolean ' PPIANOROLLChargé(0,1) = True/False
        Set(ByVal value As Boolean)
            Me.PIANOROLLChargé(NPISTE) = value
        End Set
    End Property
    Public WriteOnly Property PPIANOROLLNPiste(NPISTE As Integer) As Integer ' PPIANOROLLNPiste(0,1) = 0 ..1
        Set(ByVal value As Integer)
            Me.PIANOROLLNPiste(NPISTE) = value
        End Set
    End Property

    ' Variable pour utilisation des Accents
    ' *************************************
    Public Accent1_3 As Boolean = True

    ' SUBSTITUTION
    ' ************
    Public LabSubsti As New List(Of Label)
    Public EventhSubsti(0 To 3) As EventH

    ' MODULATION 
    ' **********
    Public LabModulat As New List(Of Label)
    Public RadioModulat As New List(Of RadioButton)
    Dim ligneModulat As Integer

    ' 
    ' Variable pour affichage des gammes
    ' **********************************
    Public ListGammesJouables As String = ""



    Class Transit
        Public Mode As New List(Of String)
        Public Notes_T As New List(Of String)
        Public Origines_N As New Dictionary(Of String, String)
    End Class
    Dim Transitions As New List(Of Transit)
    Enum Ty_Extension
        Tonalités
        Accords
        Gammes
        Modes
        Transit
    End Enum
    Dim Etat_Extension As Ty_Extension
    ' Classe générateur pour copie de variations
    ' ******************************************
    Class Générateur
        Public Mute As Boolean
        Public Motif As String
        Public Durée As Double
        Public Octave As Integer
        Public Volume As Byte
        Public Dyn As Byte
        Public PRG As Integer
        Public PAN As Byte
        Public Souche As Byte
        Public Retard As Byte
        Public Delay As Boolean
    End Class




    ' Variables pour le formualire Piano SDF
    ' **************************************
    Public DicPiano As New Dictionary(Of String, Boolean)
    Dim NoteAEtéJouée As Boolean = False
    Dim NoteCourante As Byte = 255

    ' Variable pour le formulaire F3 utilisé par la clase DrumEdit
    ' ************************************************************
    Public F3_Retour As String = "NOK"
    Public F3_Valeur As String = "D1"
    Public Grid22 As New FlexCell.Grid ' grille de time line de positionnement des variation D1..D8

    ' Variables pour gestion des rcines dns Grid2
    ' *******************************************
    Dim TRacine2 As New List(Of String)
    Dim OK_KeyDown As Boolean = True ' pour incrémentation/décrémentation des racines (évite les répétitions par appui continu)


    ' Variable pour activation du delay dans les variations
    ' *****************************************************
    Public indicDelay As Boolean = False
    Public indicVariation As Integer
    Public indicPiste As Integer




    ' *#######################################################################
    ' ##                                                                    ##
    ' ##                     CONSTRUCTION DE L'IHM                          ##
    ' ##                                                                    ##
    ' ########################################################################
    '
    Sub Maj_Culture()
        Select Case LangueIHM
            Case "fr"
                Fr_Culture()
            Case Else
                En_Culture()
        End Select
    End Sub
    '
    Sub Fr_Culture()
        '
        Try
            Dim i As Integer
            '
            ' Mise à jour des titres des onglets
            ' **********************************
            '
            TabPage1.Text = "Accords"
            TabPage2.Text = "Progressions"
            TabPage3.Text = "Vue Notes"
            TabPage16.Text = "Perso" ' ancien "spécifique"

            Tab_DrumEdit.Text = "10 - Batterie"

            ChangementLangue = True
            '
            FrançaisToolStripMenuItem.Checked = True
            EnglishToolStripMenuItem.Checked = False
            ToolStripMenuItem3.Text = "Guide rapide vidéo"

            '
            ' Tableau de bord
            ' ***************
            '
            Label5.Text = "Métrique"
            Label48.Text = "Choix Tonalité"
            '
            '
            ' Menus
            ' *****
            '
            FichierToolStripMenuItem.Text = "Fichier"
            NouveauToolStripMenuItem.Text = "Nouveau"
            OuvrirToolStripMenuItem.Text = "Ouvrir"
            EnregistrerToolStripMenuItem.Text = "Enregistrer"
            EnregistrerSousToolStripMenuItem.Text = "Enregistrer sous"
            QuitterToolStripMenuItem.Text = "Quitter"
            '
            EditionToolStripMenuItem.Text = "Edition"
            AnnulerToolStripMenuItem.Text = "Annuler"
            CouperToolStripMenuItem.Text = "Couper"
            CopierToolStripMenuItem.Text = "Copier"
            CollerToolStripMenuItem.Text = "Coller"
            EffacerToolStripMenuItem.Text = "Effacer"
            PianoToolStripMenuItem.Text = "Clavier"
            MenuKeyBoard.Text = "Clavier"
            MenuTransportBar.Text = "Barre de transport"
            '
            OutilsMenuItem.Text = "Outils"
            NotesViewToolStripMenuItem.Text = "Visu Notes HyperArp"
            MappingToolStripMenuItem.Text = "Visu Mapping"
            MuteOptimisationToolStripMenuItem.Text = "Supervision de l'automation"
            If GammesPRO.Checked Then
                GammesPRO.Text = "Passer en Gammes simples"
            Else
                GammesPRO.Text = "Passer en Gammes PRO"
            End If
            '
            AideToolStripMenuItem.Text = "Aide"
            AuSujetDeToolStripMenuItem.Text = "Au sujet de"
            SiteWebToolStripMenuItem.Text = "Site Web"
            ToolStripMenuItem6.Text = "Splash"
            ToolStripMenuItem13.Text = "Export Modèle MIDI"
            '
            Label1.Text = "Boucles" ' liste des valeurs de boucke

            ' ToolStripMenuItem6.Text = "Espaces"
            ' **********************************
            '
            OpenFileDialog1.Title = "Ouverture projet"
            '
            Label9.Text = "Début" '
            Label10.Text = "Fin" ' 
            '
            Répéter.Text = "Répéter"
            RejeuDirect.Text = "Rejeu direct"

            ' Toutes les notes sont en Anglais même quandl'IHM est en Français.
            i = ComboBox1.SelectedIndex
            ComboBox1.Items.Clear()
            '
            ComboBox1.Items.Clear()
            ComboBox1.Items.Add(" C# Major")
            ComboBox1.Items.Add(" F# Major")
            ComboBox1.Items.Add(" B Major")
            ComboBox1.Items.Add(" E Major")
            ComboBox1.Items.Add(" A Major")
            ComboBox1.Items.Add(" D Major")
            ComboBox1.Items.Add(" G Major")
            ComboBox1.Items.Add(" C Major")
            ComboBox1.Items.Add(" F Major")
            ComboBox1.Items.Add(" Bb Major")
            ComboBox1.Items.Add(" Eb Major")
            ComboBox1.Items.Add(" Ab Major")
            '
            If EnChargement Then
                ComboBox1.SelectedIndex = 7
            Else
                ComboBox1.SelectedIndex = i
            End If
            '
            i = ComboBox2.SelectedIndex
            '
            ComboBox2.Items.Clear()
            '
            ComboBox2.Items.Add(" A# Minor")
            ComboBox2.Items.Add(" D# Minor")
            ComboBox2.Items.Add(" G# Minor")
            ComboBox2.Items.Add(" C# Minor")
            ComboBox2.Items.Add(" F# Minor")
            ComboBox2.Items.Add(" B Minor")
            ComboBox2.Items.Add(" E Minor")
            ComboBox2.Items.Add(" A Minor")
            ComboBox2.Items.Add(" D Minor")
            ComboBox2.Items.Add(" G Minor")
            ComboBox2.Items.Add(" C Minor")
            ComboBox2.Items.Add(" F Minor")
            '
            If EnChargement Then
                ComboBox2.SelectedIndex = 7
            Else
                ComboBox2.SelectedIndex = i
            End If

            '
            i = ComboBox23.SelectedIndex
            ComboBox23.Items.Clear()
            ComboBox23.Items.Add(" Accords de 3 notes")
            ComboBox23.Items.Add(" Accords de 4 notes (7)")
            '
            ComboBox23.SelectedIndex = i
            If EnChargement = True Then
                ComboBox23.SelectedIndex = 0
            End If
            '
            i = ComboBox9.SelectedIndex
            ComboBox9.Items.Clear()
            ComboBox9.Items.Add(" Accords de 3 notes")
            ComboBox9.Items.Add(" Accords de 4 notes (7)")
            '
            ComboBox9.SelectedIndex = i
            If EnChargement = True Then
                ComboBox9.SelectedIndex = 0
            End If



            i = ComboBox6.SelectedIndex
            ComboBox6.Items.Clear()
            ComboBox6.Items.Add(" Accords de 3 notes")
            ComboBox6.Items.Add(" Accords de 4 notes (7)")
            '
            ComboBox6.SelectedIndex = i
            If EnChargement = True Then
                ComboBox6.SelectedIndex = 0
            End If
            '
            MenuExportsMIDI.Text = "Export fichier MIDI"
            '
            ' Onglet Cadences
            ' ***************
            '
            i = ComboBox4.SelectedIndex
            ComboBox4.Items.Clear()
            ComboBox4.Items.Add("Anatole Min")
            ComboBox4.Items.Add("Pseudo 2-5-1")
            ComboBox4.Items.Add("Plagale Min")
            ComboBox4.Items.Add("Hispanique")
            '
            ComboBox4.SelectedIndex = i
            If EnChargement = True Then
                ComboBox4.SelectedIndex = 0
            End If
            '
            i = ComboBox3.SelectedIndex '
            ComboBox3.Items.Clear()
            ComboBox3.Items.Add("Anatole")
            ComboBox3.Items.Add("Forme2")
            ComboBox3.Items.Add("Forme3")
            ComboBox3.Items.Add("Complète")
            ComboBox3.Items.Add("2-5-1")
            ComboBox3.Items.Add("Demi")
            ComboBox3.Items.Add("Parfaite")
            ComboBox3.Items.Add("Plagale")
            ComboBox3.Items.Add("Plagale2")
            ComboBox3.Items.Add("Rompue")
            ComboBox3.Items.Add("Rompue2")
            ComboBox3.Items.Add("Rompue3")
            ComboBox3.Items.Add("Modale")
            ComboBox3.Items.Add("Modale2")
            ComboBox3.Items.Add("Modale3")
            ComboBox3.Items.Add("Napolitaine")
            '
            ComboBox3.SelectedIndex = i
            If EnChargement = True Then
                ComboBox3.SelectedIndex = 0
            End If
            '
            ' Label résultat de choix d'une cadence
            ' *************************************
            '
            Label28.Text = Trad_NomCadence_EnFr(Trim(Label28.Text))
            ' ToolTips
            ' ********
            ToolTip1.SetToolTip(Button28, "Export Calques MIDI")
            ToolTip1.SetToolTip(Button19, "Vue standard. Vue admin.")

            '
            ToolTip1.SetToolTip(ComboBox1, "Tonalité")
            ToolTip1.SetToolTip(ComboBox2, "Relative mineure")
            '
            ToolTip1.SetToolTip(ComboBox23, "Types Accords")
            ToolTip1.SetToolTip(Insert, "Insérer des accords")
            ToolTip1.SetToolTip(TRacine, "Modifier la racine du voicing")
            ToolTip1.SetToolTip(FiltreUni, "Si 2 pistes possèdent des notes à l'unisson, ce filtre fait disparaître ces notes dans la 2e piste.")
            ToolTip1.SetToolTip(Fournotes, "Pour un accord de 3 notes, répétition d'une note de l'accord à l'octave supérieur")

            ' Télécommande
            Remote.Text = "Télécommande"
            ToolTip1.SetToolTip(Remote, "Envoie sur le canal 1 le CTRL 54 sur Play et le CTRL 55 sur Stop")
            ' Filtre Unisons
            FiltreUni.Text = "Filtre Unissons"

            '
            ChangementLangue = False
            '
            ' Noms des modes
            ' **************
            '
            For i = 0 To 2
                Select Case i
                    Case 0
                        TabTonsTitreMode.Item(i).Text = "Mode Majeur"
                    Case 1
                        TabTonsTitreMode.Item(i).Text = "Mode Mineur Harmonique"
                    Case 2
                        TabTonsTitreMode.Item(i).Text = "Mode Mineur Mélodique"
                End Select
            Next i
            '
            ' Etiquettes des combo des listes de cadences
            ' *******************************************
            '
            Label37.Text = "Majeures"
            Label38.Text = "Mineures"
            '
            ' Combo de copy de variations
            ' ***************************
            '
            ComboCopy1.Items.Clear()
            ComboCopy2.Items.Clear()
            For i = 1 To nb_Variations
                ComboCopy1.Items.Add("Variation" + Str(i))
                ComboCopy2.Items.Add("Variation" + Str(i))
            Next
            '
            ComboCopy1.SelectedIndex = 0
            ComboCopy2.SelectedIndex = 1
            '
            ' Editeur Motifs Perso
            ' ********************
            '
            NomsColMotifs()
            For i = 0 To TabControl1.TabPages.Count - 1
                Bcopier(i).Text = "Copier"
                Bcoller(i).Text = "Coller"
            Next i
            '
            ' Réglages système (à droite)
            ' ***************************
            '
            GroupBox1.Text = "Copie de Variations"
            GroupBox2.Text = "Choix Sortie MIDI"
            Label12.Text = "Copie"

            GroupBox3.Text = "Ecoute Accords"
            Label13.Text = "Canal"
            Label14.Text = "Vélocité"
            Label20.Text = "Dernière note jouée"
            Label22.Text = "Canal"
            ' Accents
            RadioButton3.Text = "Sur temps 1"
            RadioButton4.Text = "Sur temps 1 et 3"

            ' Admin : proposition de sauvegarde avant d'ouvrir un nouveau projet ou de le quitter
            CheckBox1.Text = "Alarme enregistrement projet en cours"
            '
            ' 4 Notes
            ComboBox11.Items.Clear()
            ComboBox11.Items.Add("V1")
            ComboBox11.Items.Add("V2")
            ComboBox11.Items.Add("V3")
            ComboBox11.SelectedIndex = 0
            '
            ' Basse-12
            ComboBox12.Items.Clear()
            ComboBox12.Items.Add("V1")
            ComboBox12.Items.Add("V2")
            ComboBox12.Items.Add("V3")
            ComboBox12.SelectedIndex = 0
            '
            ' Menu Flottant : couper, copier, coller
            ' **************************************
            '
            CouperToolStripMenuItem1.Text = "Couper"
            CopierToolStripMenuItem1.Text = "Copier"
            CollerToolStripMenuItem1.Text = "Coller"

        Catch ex As Exception

            If LangueIHM = "fr" Then
                MessageHV.PTitre = "Avertissement"
                MessageHV.PContenuMess = "Détection d'un erreur dans procédure : " + "Fr_Culture" + Constants.vbCrLf + "Message  : " + ex.Message
                MessageHV.PTypBouton = "OK"
            Else
                MessageHV.PTitre = "Warning"
                MessageHV.PContenuMess = "Détection d'un erreur dans procédure : " + "Fr_Culture" + Constants.vbCrLf + "Message  : " + ex.Message
                MessageHV.PTypBouton = "OK"
            End If
            Cacher_FormTransparents()
            MessageHV.ShowDialog()
            End
        End Try


    End Sub
    Sub En_Culture()
        Try
            Dim i As Integer
            ' mise à jour des titres des onglets
            ' **********************************
            '
            TabPage1.Text = "Chords" 'neighboring tones
            TabPage2.Text = "Progressions"
            TabPage3.Text = "Notes View"
            TabPage16.Text = "Perso" ' ancien "Spécifique"
            '
            Tab_DrumEdit.Text = "10 - Drums"

            ChangementLangue = True
            '
            FrançaisToolStripMenuItem.Checked = False
            EnglishToolStripMenuItem.Checked = True
            ToolStripMenuItem3.Text = "Video quick guide"
            '
            '
            ' tableau de bord
            Label5.Text = "Time"
            Label48.Text = "Tone Choice"
            'p.X = 60
            'p.Y = 32
            'Label48.Location = p
            'Label48.Location = p
            'p.X = 233
            'p.Y = 32
            'Label5.Location = p
            '
            ' Menus
            FichierToolStripMenuItem.Text = "File"
            NouveauToolStripMenuItem.Text = "New"
            OuvrirToolStripMenuItem.Text = "Open"
            EnregistrerToolStripMenuItem.Text = "Save"
            EnregistrerSousToolStripMenuItem.Text = "Save as"
            'ExporterCalqueMIDIToolStripMenuItem.Text = "Export MIDI Layers"
            'ExporterCompoToolStripMenuItem.Text = "Export MIDI Chords"
            QuitterToolStripMenuItem.Text = "Quit"
            '
            EditionToolStripMenuItem.Text = "Edition"
            AnnulerToolStripMenuItem.Text = "Cancel"
            CouperToolStripMenuItem.Text = "Cut"
            CopierToolStripMenuItem.Text = "Copy"
            CollerToolStripMenuItem.Text = "Paste"
            EffacerToolStripMenuItem.Text = "Delete"
            PianoToolStripMenuItem.Text = "Keyboard"
            MenuKeyBoard.Text = "Keyboard"
            MenuTransportBar.Text = "Transport Bar"
            '
            OutilsMenuItem.Text = "Tools"
            NotesViewToolStripMenuItem.Text = "HyperArp Notes View"
            MappingToolStripMenuItem.Text = "Mapping View"
            MuteOptimisationToolStripMenuItem.Text = "Automation supervision"
            '

            If GammesPRO.Checked Then
                GammesPRO.Text = "Switch to Simple scales"
            Else
                GammesPRO.Text = "Switch to PRO scales"
            End If
            '
            AideToolStripMenuItem.Text = "Help"
            AuSujetDeToolStripMenuItem.Text = "About"
            SiteWebToolStripMenuItem.Text = "Web site"
            ToolStripMenuItem6.Text = "Splash"
            ToolStripMenuItem13.Text = "Export MIDI Model"
            '
            'ToolStripMenuItem6.Text = "Spaces"
            '
            OpenFileDialog1.Title = "Open Project"
            'OpenFileDialog1.Filter = "CompoGrid (*.zic2)|*.zic2|HyperArp (*.zic3)|*.zic3"
            '
            ' Pistes
            'TabPage_Accueil.Text = "Home"
            '
            '
            Label9.Text = "Start" 'HyperArp"
            Label10.Text = "End"
            '
            Répéter.Text = "Repeat"
            RejeuDirect.Text = "Direct Replay"
            '
            Label1.Text = "Loops" ' liste des valeurs de boucke

            '
            ComboBox1.Items.Clear()
            ComboBox1.Items.Add(" C# Major")
            ComboBox1.Items.Add(" F# Major")
            ComboBox1.Items.Add(" B Major")
            ComboBox1.Items.Add(" E Major")
            ComboBox1.Items.Add(" A Major")
            ComboBox1.Items.Add(" D Major")
            ComboBox1.Items.Add(" G Major")
            ComboBox1.Items.Add(" C Major")
            ComboBox1.Items.Add(" F Major")
            ComboBox1.Items.Add(" Bb Major")
            ComboBox1.Items.Add(" Eb Major")
            ComboBox1.Items.Add(" Ab Major")

            If EnChargement Then
                ComboBox1.SelectedIndex = 7
            Else
                ComboBox1.SelectedIndex = i
            End If
            '
            i = ComboBox2.SelectedIndex
            '
            ComboBox2.Items.Clear()
            '
            ComboBox2.Items.Add(" A# Minor")
            ComboBox2.Items.Add(" D# Minor")
            ComboBox2.Items.Add(" G# Minor")
            ComboBox2.Items.Add(" C# Minor")
            ComboBox2.Items.Add(" F# Minor")
            ComboBox2.Items.Add(" B Minor")
            ComboBox2.Items.Add(" E Minor")
            ComboBox2.Items.Add(" A Minor")
            ComboBox2.Items.Add(" D Minor")
            ComboBox2.Items.Add(" G Minor")
            ComboBox2.Items.Add(" C Minor")
            ComboBox2.Items.Add(" F Minor")

            If EnChargement Then
                ComboBox2.SelectedIndex = 7
            Else
                ComboBox2.SelectedIndex = i
            End If
            '
            '
            i = ComboBox23.SelectedIndex
            ComboBox23.Items.Clear()
            ComboBox23.Items.Add(" 3 notes Chords")
            ComboBox23.Items.Add(" 4 notes Chords (7)")
            'ComboBox23.Items.Add(" 4 notes Chords (9)")
            'ComboBox23.Items.Add(" 4 notes Chords (11)")
            '
            ComboBox23.SelectedIndex = i
            If EnChargement = True Then
                ComboBox23.SelectedIndex = 0
            End If
            '
            i = ComboBox9.SelectedIndex
            ComboBox9.Items.Clear()
            ComboBox9.Items.Add(" 3 notes Chords")
            ComboBox9.Items.Add(" 4 notes Chords (7)")
            '
            ComboBox9.SelectedIndex = i
            If EnChargement = True Then
                ComboBox9.SelectedIndex = 0
            End If
            '
            i = ComboBox6.SelectedIndex
            ComboBox6.Items.Clear()
            ComboBox6.Items.Add(" 3 notes Chords")
            ComboBox6.Items.Add(" 4 notes Chords (7)")
            'ComboBox6.Items.Add(" 4 notes Chords (9)")
            'ComboBox6.Items.Add(" 4 notes Chords (11)")
            '
            ComboBox6.SelectedIndex = i
            If EnChargement = True Then
                ComboBox6.SelectedIndex = 0
            End If
            '
            'CheckBox1.Text = "Listening chords"
            '
            'Réduire.Text = "Reduce"
            '
            'NotesCommunes.Text = "Auto Voicing"
            '
            MenuExportsMIDI.Text = "MIDI File Export"
            '
            ' Onglet Cadences
            ' ***************
            i = ComboBox4.SelectedIndex
            ComboBox4.Items.Clear()
            ComboBox4.Items.Add("Anatole Min")
            ComboBox4.Items.Add("Pseudo 2-5-1")
            ComboBox4.Items.Add("Minor Plagal")
            ComboBox4.Items.Add("Hispanic")
            ComboBox4.SelectedIndex = i
            If EnChargement = True Then
                ComboBox4.SelectedIndex = 0
            End If
            '
            i = ComboBox3.SelectedIndex
            ComboBox3.Items.Clear()
            ComboBox3.Items.Add("Anatole")
            ComboBox3.Items.Add("Forme2")
            ComboBox3.Items.Add("Forme3")
            ComboBox3.Items.Add("Complete")
            ComboBox3.Items.Add("2-5-1")
            ComboBox3.Items.Add("Half")
            ComboBox3.Items.Add("Perfect")
            ComboBox3.Items.Add("Plagal")
            ComboBox3.Items.Add("Plagal2")
            ComboBox3.Items.Add("Broken")
            ComboBox3.Items.Add("Broken2")
            ComboBox3.Items.Add("Broken3")
            ComboBox3.Items.Add("Modal")
            ComboBox3.Items.Add("Modal2")
            ComboBox3.Items.Add("Modal3")
            ComboBox3.Items.Add("Napolitan")
            '
            ComboBox3.SelectedIndex = i
            If EnChargement = True Then
                ComboBox3.SelectedIndex = 0
            End If
            '
            Label28.Text = Trad_NomCadence_FrEn(Trim(Label28.Text))
            '

            ' ToolTips
            ' ********
            ToolTip1.SetToolTip(Button3, "Notes visibility on keyboard")
            ToolTip1.SetToolTip(Button25, "Choose a scale")
            ToolTip1.SetToolTip(Button28, "MIDI Layer Export")
            ToolTip1.SetToolTip(Button29, "Midi Reset")
            ToolTip1.SetToolTip(Button19, "Standard view. Admin view.")

            ToolTip1.SetToolTip(ComboBox1, "Tone")
            ToolTip1.SetToolTip(ComboBox2, "Relative minor")
            ToolTip1.SetToolTip(ComboBox23, "Chords Type")
            ToolTip1.SetToolTip(Insert, "Insert chords")
            ToolTip1.SetToolTip(TRacine, "Modify the root of the voicing")
            ToolTip1.SetToolTip(Remote, "Send channel 1 CTRL 54 on Play and CTRL 55 on Stop")
            ToolTip1.SetToolTip(FiltreUni, "If 2 tracks have identical notes, this filter makes these notes disappear in the 2nd track.")
            '
            ToolTip1.SetToolTip(FiltreUni, "If 2 tracks have notes in unison, this filter makes these notes disappear in the 2nd track.")
            ToolTip1.SetToolTip(Fournotes, "For a 3-note chord, repeat a note of the chord to the octave above.")
            ChangementLangue = False
            '
            ' Noms des modes
            ' **************
            For i = 0 To 2
                Select Case i
                    Case 0
                        TabTonsTitreMode.Item(i).Text = "Major Mode"
                    Case 1
                        TabTonsTitreMode.Item(i).Text = "Harmonic Minor Mode"
                    Case 2
                        TabTonsTitreMode.Item(i).Text = "Melodic Minor Mode"
                End Select
            Next i
            '
            ' Bouton de sélection
            ' *******************
            'Button3.Text = "Select"
            '
            ' Etiquettes des combo des listes de cadences
            ' *******************************************
            Label37.Text = "Major"
            Label38.Text = "Minor"
            '
            '
            ' Copie des Variations
            ' ********************
            ComboCopy1.Items.Clear()
            ComboCopy1.Items.Add("Variation 1")
            ComboCopy1.Items.Add("Variation 2")
            ComboCopy1.Items.Add("Variation 3")
            ComboCopy1.Items.Add("Variation 4")
            ComboCopy1.Items.Add("Variation 5")
            ComboCopy1.Items.Add("Variation 6")
            ComboCopy1.Items.Add("Variation 7")
            '
            ComboCopy1.SelectedIndex = 0
            '
            ComboCopy2.Items.Clear()
            ComboCopy2.Items.Add("Variation 1")
            ComboCopy2.Items.Add("Variation 2")
            ComboCopy2.Items.Add("Variation 3")
            ComboCopy2.Items.Add("Variation 4")
            ComboCopy2.Items.Add("Variation 5")
            ComboCopy2.Items.Add("Variation 6")
            ComboCopy2.Items.Add("Variation 7")
            '
            ComboCopy2.SelectedIndex = 1
            '


            ' Réglages système (à droite)
            ' ***************************
            GroupBox1.Text = "Variations Copy"
            GroupBox2.Text = "MIDI Out Choice"
            Label12.Text = "Copy"
            '
            GroupBox3.Text = "Listening Chords"
            Label13.Text = "Channel"
            Label14.Text = "Velocity"
            Label20.Text = "Last note Played"
            Label22.Text = "Channel"

            ' Accents
            RadioButton3.Text = "On beat 1"
            RadioButton4.Text = "On beats 1 and 3"
            '
            ' Editeur Motifs Perso
            ' ********************
            NomsColMotifs()
            For i = 0 To TabControl1.TabPages.Count - 1
                Bcopier(i).Text = "Copy"
                Bcoller(i).Text = "Paste"
            Next i

            ' Admin : proposition de sauvegarde avant d'ouvrir un nouveau projet ou de le quitter
            ' ***********************************************************************************
            CheckBox1.Text = "warning for saving current projet"

            ' Télécommande
            Remote.Text = "Remote"
            '
            ' Filtre UNissons
            FiltreUni.Text = "Unisons Filter"

            ' 4 Notes
            ComboBox11.Items.Clear()
            ComboBox11.Items.Add("V1")
            ComboBox11.Items.Add("V2")
            ComboBox11.Items.Add("V3")
            ComboBox11.SelectedIndex = 0
            '
            ' Basse-12
            ComboBox12.Items.Clear()
            ComboBox12.Items.Add("V1")
            ComboBox12.Items.Add("V2")
            ComboBox12.Items.Add("V3")
            ComboBox12.SelectedIndex = 0
            '
            ' Menu Flottant : couper, copier, coller
            ' **************************************
            CouperToolStripMenuItem1.Text = "Cut"
            CopierToolStripMenuItem1.Text = "Copy"
            CollerToolStripMenuItem1.Text = "Paste"


        Catch ex As Exception

            If LangueIHM = "fr" Then
                MessageHV.PTitre = "Avertissement"
                MessageHV.PContenuMess = "Détection d'un erreur dans procédure : " + "En_culture" + Constants.vbCrLf + "Message  : " + ex.Message
                MessageHV.PTypBouton = "OK"
            Else
                MessageHV.PTitre = "Warning"
                MessageHV.PContenuMess = "Détection d'un erreur dans procédure : " + "En_culture" + Constants.vbCrLf + "Message  : " + ex.Message
                MessageHV.PTypBouton = "OK"
            End If
            Cacher_FormTransparents()
            MessageHV.ShowDialog()
            End
        End Try
    End Sub
    Sub CréationTabAccords()
        'ComboBox20.Size = New Size(80, 20)
        'ComboBox20.Location = New Point(5, 14)
        '
        'ComboBox21.Size = New Size(80, 20)
        'ComboBox21.Location = New Point(5 + ComboBox20.Size.Width + 5, 14)
        '
        ComboBox1.Size = New Size(95, 20)
        'ComboBox1.Location = New Point(7, 4)
        '
        ComboBox2.Size = New Size(95, 20)
        'ComboBox2.Location = New Point(ComboBox1.Size.Width + 6, 4)
        '
        ' nombre de notes 
        ComboBox23.Location = New Point(5, 5) 'New Point(173, 5) ' New Point(173, 40)
        ComboBox23.Size = New Size(160, 20) ' 
        '
        '
        ComboBox6.Location = ComboBox23.Location 'New Point(2, 2) ' combo 3/4 notes dans Progressions Accord
        ComboBox6.Size = ComboBox23.Size 'New Size(160, 20) ' 
        'ComboBox6.SelectionLength = 0
        '


        Button17.Location = New Point(5, 2) ' '
        Button17.Size = New Size(40, 25) ' '
        Button17.Text = "Sel"
        '
        Button3.Location = New Point(5, 2) ' '
        Button3.Size = New Size(40, 24) ' '
        Button3.Text = "Sel"
        '
        'TabControl4.Location = New Point(1, 34) ' 70
        'TabControl4.Size = New Size(2000, 290)
    End Sub
    Sub Construction(Sig As String, Quantif As Integer)
        Try
            ' Dénominateur/Numérateur : métrique du morceau : 4/4, 3/4,7/8 ...
            ' Le dénominateur ne peut prendre que les valeurs 4 ou 8
            ' Quantif : à la croche => Quantif = 2 ; à la double croche => Quantif = 4
            ' Numérateur As Integer, Dénominateur As Integer
            Dim i As Integer
            Dim tbl() As String

            '
            ' Dimensions du formulaire
            ' ************************
            Me.WindowState = FormWindowState.Normal
            ' Culture par défaut
            ' ******************
            tbl = Split(Sig, "/")
            Numérateur = Val(tbl(0))
            Dénominateur = Val(tbl(1))
            'Sig = Trim(Str(Numérateur) + "/" + Trim(Str(Dénominateur)))
            '
            Métrique.Text = Trim(Sig)
            '
            ' *********************************************
            ' GRID2 : Initialisaton de Grid2 : navigateur *
            ' *********************************************
            '
            Grid2.AutoRedraw = False
            '
            Grid2.AllowDrop = True ' autorise le glisser/déposer dans la grille grid2 (la source est tabton.item(i).text)
            '
            Grid2.Cols = nbMesures + 12
            Grid2.Rows = 4 + nb_Variations + 1 + 1  ' Magnéto = Transfomer le dernier +1 est la ligne des racines
            '
            Grid2.Column(0).Width = 47
            For i = 1 To Grid2.Cols - 1
                Grid2.Column(i).Width = Largeur_ZoomGrid2 '90
                Grid2.Cell(0, i).Text = Str(i)
                Grid2.Cell(1, i).Text = ""
                Grid2.Column(i).Locked = True
            Next
            '
            Grid2.FixedRows = 1
            Grid2.Row(0).Height = 16 ' ligne N° d'accords
            Grid2.Row(1).Height = 19 '39 ' ligne Marqueurs
            Grid2.Row(2).Height = 31 '74 ' ligne Accords
            Grid2.Row(3).Height = 19 '39 ' ligne Répet
            Grid2.Row(4).Height = 13 ' choix variation 1
            Grid2.Row(5).Height = 13 ' choix variation 2
            Grid2.Row(6).Height = 13 ' choix variation 3
            Grid2.Row(7).Height = 13 ' choix variation 4
            Grid2.Row(8).Height = 13 ' choix variation 5
            Grid2.Row(9).Height = 13 ' choix variation 6
            Grid2.Row(10).Height = 14 ' choix variation 7
            Grid2.Row(11).Height = 25 ' Gammes
            Grid2.Row(12).Height = 25 ' racines
            '
            Grid2.Height = 249 '249 Grid2.Row(0).Height + Grid2.Row(1).Height + Grid2.Row(2).Height + Grid2.Row(3).Height '+ Grid2.Row(4).Height
            Grid2.Width = 560
            With Grid2.Range(0, 0, 3, (nbMesures))
                .Borders(FlexCell.EdgeEnum.InsideHorizontal) = FlexCell.LineStyleEnum.None
                .Borders(FlexCell.EdgeEnum.InsideVertical) = FlexCell.LineStyleEnum.None
                .Alignment = FlexCell.AlignmentEnum.CenterCenter
            End With
            '
            Grid2.SelectionMode = FlexCell.SelectionModeEnum.ByColumn
            Grid2.BackColorSel = Color.Lavender
            Grid2.ScrollBars = FlexCell.ScrollBarsEnum.Horizontal
            '
            Grid2.BoldFixedCell = False
            '
            Grid2.FixedRows = 0
            '
            For i = 1 To nbMesures
                Grid2.Cell(3, i).Text = "1"
            Next i
            '
            Maj_CouleursGrid2()
            '
            i = ColonnesUnités(Dénominateur, Quantif) ' nombre de colonnes pour une unité
            '
            For i = 4 To 10
                Grid2.Cell(i, 0).FontSize = 7
                Grid2.Cell(i, 0).Text = Convert.ToString(i - 3)
            Next
            '
            ' Init 1ere colonne de gauche
            Grid2.LeftCol = 0

            ' Positionnement de Grid2
            ' ***********************
            Grid2.Location = New Point(0, 300)
            Grid2.Width = PosSystem - 1

            ' Position des boutons associés à grid2
            ' *************************************

            Dim P1 As New Point
            Dim S As Size
            ' bouton loupe +
            P1.X = 2
            P1.Y = 325
            S.Width = 45
            S.Height = 25
            Button1.Location = P1
            Button1.Size = S
            ' bouton Loupe -
            P1.X = 2
            P1.Y = 347
            S.Width = 45
            S.Height = 25
            Button2.Location = P1
            Button2.Size = S
            '
            ' Bouton échelle
            P1.X = 2
            P1.Y = 478
            S.Width = 45
            S.Height = 28
            Button25.Location = P1
            Button25.Size = S
            Button25.Image = HyperArp.My.Resources.Resources._32x32_Transp_vert__échelle3_
            Button25.ImageAlign = ContentAlignment.MiddleCenter
            Button25.BackColor = Color.Beige
            'GammesPRO.Image = HyperArp.My.Resources.Resources._32x32_Transp_vert__échelle3_
            Button25.BringToFront()
            '

            ' combo racine
            Dim taille As New Size

            Dim PPHeight As Integer = 0

            TRacine.Visible = True

            taille.Width = Grid2.Column(0).Width - 3
            taille.Height = Grid2.Row(12).Height
            TRacine.Size = taille

            P1.X = 3
            P1.Y = 506 'PPHeight

            TRacine.Location = P1
            Maj_Tracine()  ' combo de mise à jour des racines
            Maj_Tracine2() ' liste utilisée pour incrément/décrément des racines (+/-)
            '

            For i = 1 To Grid2.Cols - 1
                Formater_Racine(i)
                'Grid2.Cell(12, i).Alignment = AlignmentEnum.CenterCenter
                'Grid2.Cell(12, i).Font = fnt2
                'Grid2.Cell(12, i).ForeColor = Color.DarkOrange
            Next
            ' 
            '
            ' Init Divereses dans Construction
            ' ********************************
            ' Première mesure par défaut
            ' **************************
            Dim ligne As Integer
            Dim m As Integer
            Dim t As Integer
            Dim ct As Integer


            ligne = 0
            For m = 0 To UBound(TableEventH, 1) '- 1
                For t = 0 To UBound(TableEventH, 2) '- 1
                    For ct = 0 To UBound(TableEventH, 3) '- 1
                        TableEventH(m, t, ct).Ligne = -1
                        TableEventH(m, t, ct).Position = ""
                        TableEventH(m, t, ct).Marqueur = ""
                        TableEventH(m, t, ct).Tonalité = ""
                        TableEventH(m, t, ct).Accord = ""
                        TableEventH(m, t, ct).Gamme = ""
                        TableEventH(m, t, ct).Détails = ""
                        TableEventH(m, t, ct).Répet = 1
                        '
                        TableNotesAccords(m, t, ct) = "" ' table contenant les notes calculées par l'auto voicing
                        'End If
                    Next ct
                Next t
            Next m
            '
            ' Initialisation des tableaux des notes calculées hors zone et avec zones
            ' ***********************************************************************
            ' Hors zones
            ' **********
            For m = 0 To nbMesuresUtiles '- 1
                For t = 0 To 5
                    For ct = 0 To 4 '
                        TableNotesAccords(m, t, ct) = ""
                    Next ct
                Next t
            Next m
            '
            THorsZone.Racine = "c2"
            THorsZone.NoteRacine = 0
            THorsZone.OctaveRacine = 0
            THorsZone.OctavePlus1 = False
            THorsZone.OctaveMoins1 = False
            '
            For i = 0 To (NbZones)
                TZone(i).DébutZ = "---"
                TZone(i).TermeZ = "---"
                '
                TZone(i).Racine = "c2"
                TZone(i).NoteRacine = 24
                TZone(i).CombiVoixInd = 0
                TZone(i).OctaveRacine = 0
                TZone(i).OctavePlus1 = False
                TZone(i).OctaveMoins1 = False

            Next
            '
            '
            ' Avec zones
            ' **********
            For m = 0 To nbMesures '- 1
                For t = 0 To 5
                    For ct = 0 To 4 '
                        TableNotesAccordsZ(m, t, ct) = ""
                    Next ct
                Next t
            Next m
            '
            Dim bb As String
            bb = "I"
            bb = Trad_DegréRomains(TableEventH(1, 1, 1).Degré)
            '
            '
            ' Boutons au niveau de Grid4 des gammes communes
            ' **********************************************
            Mode_Expert = False

            '
            ' Maj des bornes de la zone 1 qui ne sont jamais modifiables
            ' ********************************************************
            TZone(0).DébutZ = "1"
            TZone(0).TermeZ = Str(nbMesures)
            '
            '
            Grid2.Cell(1, 1).SetFocus()
            '
            ' Mise à jour de la langue
            ' ************************
            bb = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CalquesMIDI\HyperArp\LangueIHM", "LangueIHM", Nothing)
            ' ligne de commentaires
            If Trim(bb) = "fr" Then
                Fr_Culture()
            Else
                En_Culture()
            End If
            '
            Grid2.AutoRedraw = True
            Grid2.Refresh()
            '
            ' Diverses Position Initiales
            ' ***************************e
            '
            'Me.Size = New Size(1395, 682) '682 Dimensions de la fenêtre principale
            SplitContainer7.SplitterDistance = PosSystem ' PosStandard ' position du splitter central
            Panel2.Visible = False

            ' Blocage en écriture de toutes les cellules : N°, Marqueur, Accord, Répétition
            ' **************************************************************************************
            'Grid2.Range(1, 1, Grid2.Rows - 1, Grid2.Cols - 1).Locked = True
            VérouillerGrid2()

            ' Nombre de boucles en mode répétition
            ' ************************************
            NbBoucles.Items.Clear()
            For i = 0 To 15
                NbBoucles.Items.Add(Convert.ToString(i + 1))
            Next
            NbBoucles.SelectedIndex = 0 ' 
            NbBoucles.Visible = True
            '
            Me.Visible = True
            '
            ' ************************************************
            ' Mise à jour Numéric Updown Début - Fin - Tempo *
            ' ************************************************
            '
            Grid2.Cell(2, 1).Text = "C" ' écriture provisoire nécessaire pour ça ne plante pas juste après lors de l'appel à 'Det_NumDerAccord' dans l'évènement Début_ValueChanged déclenché par 'Début.Minimum = 1'


            Début.Minimum = 1
            Début.Maximum = nbMesures
            Début.Value = 1
            '
            Terme.Minimum = 1
            Terme.Maximum = nbMesures
            Terme.Value = 5
            '
            Tempo.Minimum = 1
            Tempo.Maximum = 260
            Tempo.Value = 120
            '
            Tempo.BackColor = Coul_ELActif

        Catch ex As Exception
            If LangueIHM = "fr" Then
                MessageHV.PTitre = "Avertissement"
                MessageHV.PContenuMess = "Détection d'un erreur dans procédure : " + "Construction" + Constants.vbCrLf + "Message  : " + ex.Message
                MessageHV.PTypBouton = "OK"
            Else
                MessageHV.PTitre = "Warning"
                MessageHV.PContenuMess = "Détection d'un erreur dans procédure : " + "Construction" + Constants.vbCrLf + "Message  : " + ex.Message
                MessageHV.PTypBouton = "OK"
            End If
            Cacher_FormTransparents()
            MessageHV.ShowDialog()
            End
        End Try
    End Sub

    Sub VérouillerGrid2()
        Dim i, j As Integer
        For i = 1 To Grid2.Rows - 1
            For j = 1 To Grid2.Cols - 1
                Grid2.Cell(i, j).Locked = True
            Next
        Next
    End Sub
    Sub DéVérouillerGrid2()
        Dim i, j As Integer
        For i = 1 To Grid2.Rows - 1
            For j = 1 To Grid2.Cols - 1
                Grid2.Cell(i, j).Locked = False
            Next
        Next
    End Sub

    Sub Maj_Tracine2()

        TRacine2.Add("c1")
        TRacine2.Add("c#1")
        TRacine2.Add("d1")
        TRacine2.Add("d#1")
        TRacine2.Add("e1")
        TRacine2.Add("f1")
        TRacine2.Add("f#1")
        TRacine2.Add("g1")
        TRacine2.Add("g#1")
        TRacine2.Add("a1")
        TRacine2.Add("a#1")
        TRacine2.Add("b1")
        '
        TRacine2.Add("c2")
        '
        TRacine2.Add("c#2")
        TRacine2.Add("d2")
        TRacine2.Add("d#2")
        TRacine2.Add("e2")
        TRacine2.Add("f2")
        TRacine2.Add("f#2")
        TRacine2.Add("g2")
        TRacine2.Add("g#2")
        TRacine2.Add("a2")
        TRacine2.Add("a#2")
        TRacine2.Add("b2")
        TRacine2.Add("c3")
        TRacine2.Add("c#3")
        TRacine2.Add("d3")
        TRacine2.Add("d#3")
        TRacine2.Add("e3")
        TRacine2.Add("f3")
    End Sub
    Sub Maj_Tracine()
        TRacine.Items.Clear()
        '
        'TRacine.Items.Add("c0")
        'TRacine.Items.Add("c#0")
        'TRacine.Items.Add("d0")
        'TRacine.Items.Add("d#0")
        'TRacine.Items.Add("e0")
        'TRacine.Items.Add("f0")
        'TRacine.Items.Add("f#0")
        'TRacine.Items.Add("g0")
        'TRacine.Items.Add("g#0")
        'TRacine.Items.Add("a0")
        'TRacine.Items.Add("a#0")
        'TRacine.Items.Add("b0")

        TRacine.Items.Add("c3")
        TRacine.Items.Add("b2")
        TRacine.Items.Add("a#2")
        TRacine.Items.Add("a2")
        TRacine.Items.Add("g#2")
        TRacine.Items.Add("g2")
        TRacine.Items.Add("f#2")
        TRacine.Items.Add("f2")
        TRacine.Items.Add("e2")
        TRacine.Items.Add("d#2")
        TRacine.Items.Add("d2")
        TRacine.Items.Add("c#2")
        '
        TRacine.Items.Add("c2")
        '
        TRacine.Items.Add("b1")       '
        TRacine.Items.Add("a#1")
        TRacine.Items.Add("a1")
        TRacine.Items.Add("g#1")
        TRacine.Items.Add("g1")
        TRacine.Items.Add("f#1")
        TRacine.Items.Add("f1")
        TRacine.Items.Add("e1")
        TRacine.Items.Add("d#1")
        TRacine.Items.Add("d1")
        TRacine.Items.Add("c#1")
        TRacine.Items.Add("c1")
        '

        'TRacine.Items.Add("c#3")
        'TRacine.Items.Add("d3")
        'TRacine.Items.Add("d#3")
        'TRacine.Items.Add("e3")
        'TRacine.Items.Add("f3")
        'TRacine.Items.Add("f#3")
        'TRacine.Items.Add("g3")
        'TRacine.Items.Add("g#3")
        'TRacine.Items.Add("a3")
        'TRacine.Items.Add("a#3")
        'TRacine.Items.Add("b3")
        '
        'TRacine.Items.Add("c4")
        'TRacine.Items.Add("c#4")
        'TRacine.Items.Add("d4")
        'TRacine.Items.Add("d#4")
        'TRacine.Items.Add("e4")
        'TRacine.Items.Add("f4")
        'TRacine.Items.Add("f#4")
        'TRacine.Items.Add("g4")
        'TRacine.Items.Add("g#4")
        'TRacine.Items.Add("a4")
        'TRacine.Items.Add("a#5")
        'TRacine.Items.Add("b5")
        '
        If EnChargement Then
            TRacine.SelectedIndex = 13
        End If
    End Sub
    Sub CAD_Construction()
        '
        Try
            Cad_AnatoleMaj()
            CAD_Maj_TableGlobalAcc()
            '
            ComboBox3.Location = New Point(175, 5)
            ComboBox4.Location = New Point(175 + ComboBox3.Size.Width + 8, 5)
            '
            Label37.Location = New Point(175, 24)     ' label "Majeures" pour combobox37 liste des progressions majeures
            Label38.Location = New Point(175 + ComboBox3.Size.Width + 8, 24)   ' label "Mineures" pour combobox38 liste des progressions mineures


            'Button3.Location = New Point(360, 4) '  désélection/sélection des cases à cocher
            'Button3.Size = New Size(108, 25)
            '
            Label28.Location = New Point(1, 43)
            Label28.Size = New Size(90, 10)
            '
            ' Init de la cadence de base dans l'onget progression
            ' ***************************************************
            ComboBox1.SelectedIndex = 7 ' choix de la tonalité de C Maj à l'init
            ComboBox3.SelectedIndex = 0 ' chois progression "Anatole"
            Cad_OrigineAccord = Modes.Cadence_Majeure
            Label28.Text = "Anatole"
            TabCad.Item(0).Text = "C"
            TabCad.Item(1).Text = "A m"
            TabCad.Item(2).Text = "D m"
            TabCad.Item(3).Text = "G"
            TabCad.Item(4).Text = "C"


            Label28.TextAlign = ContentAlignment.MiddleLeft
            'Label28.Font = New System.Drawing.Font("Verdana", 7, FontStyle.Italic)
            '
        Catch ex As Exception

            If LangueIHM = "fr" Then
                MessageHV.PTitre = "Avertissement"
                MessageHV.PContenuMess = "Détection d'un erreur dans procédure : " + "CAD_Construction" + Constants.vbCrLf + "Message  : " + ex.Message
                MessageHV.PTypBouton = "OK"
            Else
                MessageHV.PTitre = "Warning"
                MessageHV.PContenuMess = "Détection d'un erreur dans procédure : " + "CAD_Construction" + Constants.vbCrLf + "Message  : " + ex.Message
                MessageHV.PTypBouton = "OK"
            End If
            Cacher_FormTransparents()
            MessageHV.ShowDialog()
            End
        End Try
    End Sub

    Sub BLUES_Maj()
        If EnChargement = False Then
            '
            'CAD_Init_Aff()
            '
            TabBluesDegrés.Item(0).Text = "I"
            TabBluesDegrés.Item(0).Visible = True ' degré
            '
            TabBluesDegrés.Item(1).Text = "VI"
            TabBluesDegrés.Item(1).Visible = True ' degré

            TabBluesDegrés.Item(2).Text = "II"
            TabBluesDegrés.Item(2).Visible = True ' degré

            TabBluesDegrés.Item(3).Text = "V"
            TabBluesDegrés.Item(3).Visible = True ' degré

            TabBluesDegrés.Item(4).Text = "I"
            TabBluesDegrés.Item(4).Visible = True ' degré
            '
            'CAD_Maj_TableGlobalAcc()
        End If
    End Sub
    Sub Init_Fichier()
        'CheminFichierOuvrir = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CalquesMIDI\HyperArp\Chemins", "CheminFichierOuvrir", Nothing)
        NomFichier = ""
        IndicateurEnreg = False
        Me.Text = "Sans Titre"
    End Sub
    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed

        Try
            ' Tous les paramètres mis à jour par UpandDown sont enregistrés ici une seule fois, à la sortie de l'application
            ' **************************************************************************************************************
            '
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CalquesMIDI\HyperArp\Chemins", "CheminFichierOuvrir", Trim(CheminFichierOuvrir))
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CalquesMIDI\HyperArp\Chemins", "CheminEnregistrer", Trim(CheminFichierEnreg))
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CalquesMIDI\HyperArp\Chemins", "ExportCalquesMIDI", Trim(CheminFichierCalques))
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CalquesMIDI\HyperArp\Chemins", "ExportFichierMIDI", Trim(CheminFichierMIDI))            '
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CalquesMIDI\HyperArp\Langue", "Langue", Trim(Langue))  ' Langue utilisée pour traitement des informations sur les notes (toujour en)
            'My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CalquesMIDI\HyperArp\Préférences", "WarningProjetCourant", Convert.ToString(CheckBox1.Checked))
            '

            ' Nettoyage avant de sortir de l'application
            ' ******************************************
            'fermer les ressources MIDI avant de sortie de l'application
            Fermer_MIDI()
            If ComboMidiOut.Items.Count > 0 Then
                If SortieMidi.Item(ChoixSortieMidi).IsOpen Then
                    SortieMidi.Item(ChoixSortieMidi).Close()
                End If
                'Application.Exit()
            End If
            '
            ' Effacer les fichiers dans MesDocuments/HyperArp
            Effacer_CTemp()
            '

        Catch ex As Exception

            If LangueIHM = "fr" Then
                MessageHV.PTitre = "Avertissement"
                MessageHV.PContenuMess = "Détection erreur dans procédure : " + "Form1_FormClosed" + Constants.vbCrLf + "Message  : " + ex.Message
                MessageHV.PTypBouton = "OK"
            Else
                MessageHV.PTitre = "Warning"
                MessageHV.PContenuMess = "Détection erreur dans procédure : " + "Form1_FormClosed" + Constants.vbCrLf + "Message  : " + ex.Message
                MessageHV.PTypBouton = "OK"
            End If
            Cacher_FormTransparents()
            MessageHV.ShowDialog()
            End
        End Try
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim i, j, k As Integer
        '
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.AppStarting
        Me.Controls.setAllControlsTabstop(False) 'permet de mettre touts les tabStop = false La méthode 'setAllControlsTabstop' se trouve dans le module Extensions
        '
        Try
            Me.Visible = False

            ' Nom document par défaut
            ' ***********************
            EnChargement = True
            '
            ' Initde tableaux membres de la structure "AccordTab" pour la table "TableCoursAcc"
            ' ********************************************************************************
            For i = 0 To 2
                For j = 0 To 6
                    ReDim Preserve TableCoursAcc(i, j).Renversement(5)
                    TableCoursAcc(i, j).Marqué = False
                Next j
            Next i
            For i = 0 To 2
                For j = 0 To 6
                    ReDim Preserve TableCoursAcc(i, j).EtendreNotes(5)
                    ReDim Preserve TableCoursAcc(i, j).EtendreChecked(5)
                    For k = 0 To 5
                        TableCoursAcc(i, j).EtendreNotes(k) = ""
                        TableCoursAcc(i, j).EtendreChecked(k) = False
                    Next
                Next j
            Next i
            '
            ' Initialisation des TabControl pour colorisation de leurs onglet
            ' ***************************************************************
            TabControl2.DrawMode = TabDrawMode.OwnerDrawFixed ' onglets saudessus des sources d'accords HyperArp
            TabControl4.DrawMode = TabDrawMode.OwnerDrawFixed ' Onglets généraux : HyperArp, PianoRoll, DrumEdit et Mix
            TabControl5.DrawMode = TabDrawMode.OwnerDrawFixed ' Onglets des Variations de l'arpégiateur

            '
            ' Association Notes - paramètres BasseMoinbs1 et Large
            ' ****************************************************
            THorsZone.VoixAsso_OctaveMoins1 = "Voix1"
            THorsZone.VoixAsso_OctavePlus1 = "Voix2"
            '
            For i = 0 To (NbZones)
                TZone(i).VoixAsso_OctaveMoins1 = "Voix1"
                TZone(i).VoixAsso_OctavePlus1 = "Voix2"
            Next i
            ' Maj de tables et listes d'informations nécessaires au fonctionnement
            ' ********************************************************************
            Maj_ListNotesd() ' liste des N° notes/Nom Notes en #
            Maj_ListNotesb() ' liste des N° notes/Nom Notes en b
            '
            Maj_ListN()
            Maj_ListNLatine()
            MajTabCoulZone()
            MajTabCoulVariations()
            Maj_DicoCouleur() ' dictionnaire des couleurs - la procédure Construction utilise les dictionnaires de couleur
            '
            Init_Fichier()
            '
            ' Création/Lecture de la base de registre de l'appli
            ' **************************************************
            CréationRegistry()
            '
            ' Créations relative à HyperArp
            ' *****************************
            Maj_TMouvements() ' le mouvement des modes est toujours un mouvement de secondes
            BLOCS_création() ' éléments constituant les variations (1 variation contient 6 blocs)
            '
            TabControl2.Visible = False
            '
            TONS_Création()  ' Création onglet "gammes d'accords "
            SUBSTITUTION_CREATION()
            MODULATION_CREATION()
            MOTIFSPERSO_Création() ' création Onglet Perso
            CréationTabAccords() ' barre outil gammes d'accords
            '
            TabControl2.Visible = True
            TabControl2.Refresh()
            '
            ' Constructions
            ' *************
            Maj_TAccents("4/4")
            Maj_ListMajMin()
            Construction("4/4", 2) ' construction globale
            PIANOROLL_Création2()
            CHORDROLL_Création()
            DRUMS_Création()
            MIXAGE_Création()
            AUTOMATION_Création()
            '
            ' Mise à jour des listes de gammes pour algo. de rechercher de gammes
            ' *******************************************************************
            Maj_LGam()

            ' Couleur combobox des Gammes d'accords
            ' *************************************
            ComboBox1.BackColor = Couleur_Accord_Majeur
            ComboBox2.BackColor = Couleur_Accord_Mineur
            ComboBox3.BackColor = Couleur_Accord_Majeur
            ComboBox4.BackColor = Couleur_Accord_Mineur           '
            '
            SélectionLignes = 0 ' valeur défaut de lignes sélectionnées dans grid1 par barres rouge et bleue < -- ????????????? il n'y a plus de barres rouge et bleu
            '
            ' Init Variables gobales diverses
            ' ******************************
            Entrée_Tonalité = "C Maj"
            EnChargement = False
            OngletCours = 0
            MesureCourante = 1
            ' 
            TabControl2.Visible = False
            '
            CAD_Création()
            CAD_Construction()
            '
            TabControl2.Visible = True
            TabControl2.Refresh()
            '
            ' Mise à jour table, dico, list d'information nécessaires aux fonctionnement
            ' **************************************************************************
            Maj_DicoNotes()
            Maj_TabNotesD_B()
            Maj_ListNotesMajd()
            Maj_ListN()
            '
            ' Ecoute
            ' ******
            For i = 1 To 16
                CanalEcoute.Items.Add(Convert.ToString(i))
            Next
            CanalEcoute.SelectedIndex = 0

            ' Init d'un dictionnaire nécessaire à la création du Mini Clavier
            ' **************************************************************
            Init_Dicpiano() ' pour formulaire CLAVIER
            '
            ' panel des boutons Tonalités, Accords, Gammes, Modes
            ' ***************************************************
            AffTona()

            '
            ' info pour utilisation des touches F4 et F5, respectivement touche de stop et touche de Play (Recalcul)
            ' *****************************************************************************************************
            Me.KeyPreview = True
            PlayMidi.Focus()

            ' Dimensionnement de l'application
            ' ********************************
            PanelBarreOutil.Size = New Size(1385, 55)
            TabControl4.Size = New Size(1385, 580)
            Me.Size = New Size(1385, 590) '590 la taille de me.size est dépendante de la taille de PanelBarreOutil et  tabcontrol4
            Me.FormBorderStyle = FormBorderStyle.FixedSingle
            Me.BringToFront()

            Maj_DelockCell()
            '
            ' Forçage accord de base "C Maj"
            ' *****************************
            ForçageAccordBase()

            ' Mise en place de base de l'application
            ' **************************************
            SplitContainer7.SplitterDistance = PosSystem
            SplitContainer2.SplitterDistance = 550
            SplitContainer2.Panel2.Visible = True
            Panel2.Visible = False
            '
            ' position des onglets sur 1er onglet
            ' ***********************************
            TabControl2.SelectedTab = TabControl2.TabPages(0) ' 1er Onglet HyperArp
            TabControl4.SelectedTab = TabControl4.TabPages(0) ' 1er Onglet Généraux
            TabControl5.SelectedTab = TabControl5.TabPages(0) ' 1er Onglet des variations
            '
            ' Refresh des variations (anciennement magnéto)
            ' *********************************************
            '
            RafraichissementOnglets()
            ' Variables utilisées par le Copier/Coller
            ' ****************************************
            OngletCours_Edition = 0
            OngletCours_HyperARP = 0
            '
            ' Temps désigné par les locators
            ' ******************************

            TextBox2.Text = Calcul_Durée()

            Me.Visible = True
            '
            Clipboard.Clear() ' le clipboard est remis à jour régulièreement dans  'Private Sub Form1_Activated'

            ' Lancement de la splah Image au démarrage
            ' ****************************************
            Splash.ShowDialog()


            '
            ' ****************************
            ' * TEST DES INTERFACES MIDI *
            ' ****************************

            TestInterfaceMIDI2()

            ' System
            ' ******
            Clipboard.Clear()
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default


        Catch ex As Exception

            If LangueIHM = "fr" Then
                MessageHV.PTitre = "Avertissement"
                MessageHV.PContenuMess = "Détection d'une erreur dans procédure : " + "Form1_Load" + "." + Constants.vbCrLf + "Message  : " + ex.Message
                MessageHV.PTypBouton = "OK"
            Else
                MessageHV.PTitre = "Warning"
                MessageHV.PContenuMess = "Détection d'une erreur dans procédure : " + "Form1_Load" + "." + Constants.vbCrLf + "Message  : " + ex.Message
                MessageHV.PTypBouton = "OK"
            End If
            Cacher_FormTransparents()
            MessageHV.ShowDialog()
            End
        End Try

    End Sub
    'Sub Init_Mute_HyperArp()
    'Dim tbl() As String
    'Dim canMIDI As Integer
    '
    'For canMIDI = 0 To Module1.nb_Variations - 1
    '       tbl = N_BLOC_MIDI(canMIDI).Split()
    'For Each a As String In tbl
    '            i = Convert.ToInt16(a)
    'Select Case canMIDI
    'Case 0, 1, 2, 3
    '                   'RemoveHandler PisteMute.Item(i).CheckedChanged, AddressOf PisteMute_CheckedChange
    '                  PisteMute.Item(i).Checked = True ' Mix.soloVolume.Item(ind).Checked
    '                    Mix.soloVolume.Item(canMIDI).Checked = True
    '                    'AddHandler PisteMute.Item(i).CheckedChanged, AddressOf PisteMute_CheckedChange
    'Case 4, 5
    '                    PisteMute.Item(i).Checked = False
    '                    Mix.soloVolume.Item(canMIDI).Checked = False
    'End Select
    'Next
    'Next
    'End Sub


    Sub Init_Dicpiano()
        DicPiano("S") = False
        DicPiano("E") = False
        DicPiano("D") = False
        DicPiano("R") = False
        DicPiano("F") = False
        DicPiano("G") = False
        DicPiano("Y") = False
        DicPiano("H") = False
        DicPiano("U") = False
        DicPiano("J") = False
        DicPiano("I") = False
        DicPiano("K") = False
        DicPiano("L") = False
    End Sub



    Sub Maj_DelockCell()
        For i = 0 To Grid2.Cols - 1
            Grid2.Column(i).Locked = False
        Next
    End Sub
    Sub ForçageAccordBase()
        Dim a, b, c, d As String
        Dim i As Integer
        LabelCours = 0
        Entrée_Tonalité = "C Maj"
        Entrée_Mode = "C Maj"
        Entrée_Gamme = "C Maj"
        '

        TabControl2.SelectedIndex = 0 ' nécessaire pour faire fonctionner la présente méthode "ForçageAccordBase"


        Flag_EcrDragDrop = True
        OngletCours = 1
        EcritureAccordDsGrid2("C", 1) ' 
        '
        ' Mise à jour PianoRoll
        ' *********************
        a = Trim(Det_ListAcc())
        b = Trim(Det_ListGam())
        c = Trim(Det_ListMarq())
        d = Trim(Det_ListTon())
        For i = 0 To nb_PianoRoll - 1
            If PIANOROLLChargé(i) = True Then
                listPIANOROLL(i).PListAcc = a 'Det_ListAcc()
                listPIANOROLL(i).PListGam = b 'Det_ListGam()
                listPIANOROLL(i).PListMarq = c 'Det_ListMarq()
                listPIANOROLL(i).PListTon = d 'Det_ListMarq()
                listPIANOROLL(i).F1_Refresh()
                listPIANOROLL(i).Maj_CalquesMIDI()
            End If
        Next
        '
        Automation1.PListAcc = Det_ListAcc()
        Automation1.PListMarq = Det_ListMarq()
        Automation1.F4_Refresh()

        Drums.PListAcc = Det_ListAcc()
        Drums.PListMarq = Det_ListMarq()
        Drums.F2_Refresh()
        ' Mettre à jour selon auto voicing
        Calcul_AutoVoicingZ()
    End Sub


    Function ComptageColonneMesure(Numérateur As Integer, Dénominateur As Integer, NumColMesure As Integer) As Integer
        '
        ComptageColonneMesure = 0
        Select Case Dénominateur
            Case 4
                If NumColMesure = Numérateur * 2 Then
                    ComptageColonneMesure = 1
                Else
                    ComptageColonneMesure = NumColMesure + 1
                End If
            Case 8
                If NumColMesure = Numérateur Then
                    ComptageColonneMesure = 1
                Else
                    ComptageColonneMesure = NumColMesure + 1
                End If
        End Select
    End Function
    Function ColonnesPourSignature(Numérateur As Integer, Dénominateur As Integer, Quantif As Integer) As Integer
        Dim i As Integer
        Dim j As Integer

        i = ColonnesUnités(Dénominateur, Quantif) ' nombre de colonnes pour une unité
        j = (Numérateur * i) ' nombre de colonnes dans 1 mesure
        ColonnesPourSignature = j * (nbMesures + 11)
        '
    End Function
    Function ColonnesUnités(Dénominateur As Integer, Quantif As Integer) As Integer
        ColonnesUnités = -1
        Select Case Dénominateur
            Case 4
                Select Case Quantif
                    Case 2 ' quantification à la croche
                        ColonnesUnités = 2
                    Case 4 ' quantification à la double croche
                        ColonnesUnités = 4
                End Select
            Case 8
                Select Case Quantif
                    Case 2 ' quantification à la croche
                        ColonnesUnités = 1
                    Case 4 ' quantification à la double croche
                        ColonnesUnités = 2
                End Select
        End Select
    End Function
    Function Trad_NomCadence_FrEn(Etiquette As String) As String
        Dim a As String
        a = "Anatole"
        Select Case Trim(Etiquette)
            Case "Anatole"
                a = "Anatole"
            Case "Complète"
                a = "Complete"
            Case "2-5-1"
                a = "2-5-1"
            Case "Demi"
                a = "Half"
            Case "Parfaite"
                a = "Perfect"
            Case "Plagale"
                a = "Plagal"
            Case "Plagale2"
                a = "Plagal2"
            Case "Rompue"
                a = "Broken"
            Case "Rompue2"
                a = "Broken2"
            Case "Rompue3"
                a = "Broken3"
            Case "Modale"
                a = "Modal"
            Case "Modale2"
                a = "Modal2"
            Case "Modale3"
                a = "Modal3"
            Case "Napolitaine"
                a = "Napolitan"
            Case "Anatole Min"
                a = "Anatole Min"
            Case "Pseudo 2-5-1"
                a = "Pseudo 2-5-1"
            Case "Plagale Min"
                a = "Minor Plagal"
            Case "Hispanique"
                a = "Hispanic"
        End Select
        Trad_NomCadence_FrEn = Trim(a)
    End Function
    Function Trad_NomCadence_EnFr(Etiquette As String) As String
        Dim a As String
        a = "Anatole"
        Select Case Trim(Etiquette)
            Case "Anatole"
                a = "Anatole"
            Case "Complete"
                a = "Complète"
            Case "2-5-1"
                a = "2-5-1"
            Case "Half"
                a = "Demi"
            Case "Perfect"
                a = "Parfaite"
            Case "Plagal"
                a = "Plagale"
            Case "Plagal2"
                a = "Plagale2"
            Case "Broken"
                a = "Rompue"
            Case "Broken2"
                a = "Rompue2"
            Case "Broken3"
                a = "Rompue3"
            Case "Modal"
                a = "Modale"
            Case "Modal2"
                a = "Modale2"
            Case "Modal3"
                a = "Modale3"
            Case "Napolitan"
                a = "Napolitaine"
            Case "Anatole Min"
                a = "Anatole Min"
            Case "Pseudo 2-5-1"
                a = "Pseudo 2-5-1"
            Case "Minor Plagal"
                a = "Plagale Min"
            Case "Hispanic"
                a = "Hispanique"
        End Select
        Trad_NomCadence_EnFr = a
    End Function
    ' *#######################################################################
    ' ##                                                                    ##
    ' ##                               MIDI                                 ##
    ' ##                                                                    ##
    ' ########################################################################
    Public Sub Envoyer_PRG()
        Dim pst As Integer
        Dim PRG As Integer
        Dim canal As Byte

        If EnChargement = False Then
            For pst = 0 To nb_BlocPistes - 1 'Arrangement1.Nb_Pistes - 1
                canal = LesPistes.Item(pst).Canal
                'PRG = LesPistes.Item(pst).PRG
                If Trim(PRG) <> -1 Then
                    If Not (SortieMidi.Item(ChoixSortieMidi).IsOpen) Then
                        SortieMidi.Item(ChoixSortieMidi).Open()
                    End If
                    tbl = Split(Trim(PRG))
                    SortieMidi.Item(ChoixSortieMidi).SendProgramChange(canal, PRG)
                End If
            Next pst
        End If
    End Sub
    '
    Private Sub CheckBox9_CheckedChanged(sender As Object, e As EventArgs)
        'Ouverture_EntréeMidi()
    End Sub


    Sub FermerMidi()
        If ComboMidiOut.Items.Count > 0 Then
            If SortieMidi.Item(ChoixSortieMidi).IsOpen Then
                SortieMidi.Item(ChoixSortieMidi).Close()
            End If
        End If
    End Sub
    Private Sub StopMidi_Click(sender As Object, e As EventArgs) Handles StopMidi.Click
        StopPlay()
    End Sub
    Private Sub PlayMidi_Click(sender As Object, e As EventArgs) Handles PlayMidi.Click
        PlayHyperArp()
    End Sub
    Sub PlayHyperArp()


        '
        If Terme.BackColor <> Color.Red Then
            Dim a As String = Me.Récup_Acc
            If Trim(a) <> "" Then
                Me.Cursor = Cursors.WaitCursor
                Module1.JeuxActif = True ' utilisé par barre de transport libre et Maj_AffVoicing (appelé par Calcul_AutoVoicingZ)
                ' init de la jauge
                ' ****************
                JaugeInit()
                JaugeProgres(5)
                '
                Calcul_AutoVoicingZ()
                'jauge
                JaugeProgres(5)
                CalculArp(False)
                JaugeProgres(5)
                'Maj_Mute()
                FermerEntréeMIDI()
                PlayArp()
                JaugeProgres(5)
                Me.Cursor = Cursors.Default
                '
                ' fin jauge
                JaugeFin()

            End If
        End If
    End Sub

    Public Sub JaugeFin()
        Label23.Visible = False ' Label30 est dessus label23
        Label30.Visible = False
    End Sub
    Public Sub JaugeInit()
        TailleJauge = Label23.Size
        TailleJauge.Width = 5
        '
        Label23.Visible = True ' Label30 est dessus label23
        Label30.Visible = True
        Label30.Size = TailleJauge
        Label30.Location = Label23.Location
    End Sub

    Public Sub JaugeProgres(pas As Integer)
        TailleJauge.Width = TailleJauge.Width + pas
        Label30.Size = TailleJauge
        Label30.BringToFront()
        Label23.Refresh()
        Label30.Refresh()
    End Sub


    Sub Init_CTRLMIDI2()
        Dim i As Integer
        If Not (SortieMidi.Item(ChoixSortieMidi).IsOpen) Then
            SortieMidi.Item(ChoixSortieMidi).Open()
        End If
        For i = 0 To nb_TotalPistes - 2 ' on compte pas le drum edit
            SortieMidi.Item(ChoixSortieMidi).SendControlChange(i, 1, 0) ' Modulation Wheel à 0
            'If Mix.AutorisVol.Checked Then
            'SortieMidi.Item(ChoixSortieMidi).SendControlChange(1, 7, 100) ' Volume à 100 (max) 
            'End If
            SortieMidi.Item(ChoixSortieMidi).SendControlChange(i, 64, 0) ' Ctrl Pedale = 0
            SortieMidi.Item(ChoixSortieMidi).SendControlChange(i, 11, 100) ' Expression à 100 (max) 
            SortieMidi.Item(ChoixSortieMidi).SendControlChange(i, 10, 64) ' Panoramique à 64 : valeur centrale
            SortieMidi.Item(ChoixSortieMidi).SendPitchBend(i, 8192) ' pitch bend = 64 --> valeur centrale
        Next
    End Sub
    Sub FermerEntréeMIDI()
        Dim i As Integer
        Dim a As String

        If ExistenceEntréeMidi = True Then
            For i = 0 To EntréeMidi.Count - 1
                If EntréeMidi.Item(i).IsOpen Then
                    EntréeMidi.Item(i).Close()
                    a = EntréeMidi.Item(i).Name
                End If
            Next i
        End If
    End Sub


    Sub StopPlay()
        If ComboMidiOut.Items.Count > 0 Then '
            FIN()
        End If
        Module1.JeuxActif = False
        MIDIReset()
    End Sub
    Private Sub StopToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StopToolStripMenuItem.Click
        StopPlay()
    End Sub
    Sub PlayArp()
        Dim tbl2() As String
        Dim Evénement As String = "Note"
        Dim i, j As Integer
        Dim k As Integer = 0
        Dim Position, Durée As Single
        Dim FinDemandée As Boolean = False
        'Dim PisteLectAcc As Integer = Det_PremPisteActive()

        Horloge1.Reset()
        Horloge1.BeatsPerMinute = Tempo.Value

        NumMesList22()  ' call back pour affichage des mesures
        For i = 0 To nb_TotalPistes - 1 'Arrangement1.Nb_PistesMidi
            If LesPistes.Item(i).part.Count > 0 Then
                For j = 0 To LesPistes.Item(i).part.Count - 1
                    tbl2 = Split(LesPistes.Item(i).part(j))
                    Select Case tbl2(0)
                        Case "Acc"
                            'tbl2(0)="Acc"
                            'tbl2(1)=Chiffrage
                            'tbl2(2)=Numéro Piste
                            'tbl2(3)=Numéro d'accord
                            'tbl2(4)=Position
                            Position = ((Val(tbl2(4)) / 4) + 1)
                            Horloge1.Schedule(New CallbackMessage(AddressOf EVT_Acc, Position))
                            k = k + 1
                        Case "Note"
                            'tbl2(0)="Note"
                            'tbl2(1)=Numéro Piste
                            'tbl2(2)=Numéro de canal
                            'tbl2(3)=Numéro de Note
                            'tbl2(4)=Position
                            'tbl2(5)=Duration
                            'tbl2(6)=Dyn
                            '
                            ' Remarque sur le calcul de Position net de Durée
                            ' 1 - La dll midi dotnet impose une expression en beat avec virgule flottante (noire) des Positions et Durées
                            ' 2 - HyperArp compte la Position et la Durée en double croche (16 divisions)
                            ' 3 - Donc, pour exprimer les Positions et Durées en noire, il faut les diviser par 4
                            ' 4 - Remarque : si HyperArp décide de compter en triple croches(32 divisions), il faudra faire une division par 8.
                            '
                            ' NoteOnOffMessage(DeviceBase device,Channel channel,Pitch pitch,int velocity,float time,Clock clock, float duration
                            Position = ((Val(tbl2(4)) / 4) + 1)
                            Durée = ((Val(tbl2(5)) / 4) - 0.0625)
                            '                                                                        Canal MIDI            N° Note             Dyn                 Position  Horlorge  Durée
                            Horloge1.Schedule(New NoteOnOffMessage(SortieMidi.Item(ChoixSortieMidi), CByte(Val(tbl2(2))), CByte(Val(tbl2(3))), CByte(Val(tbl2(6))), Position, Horloge1, Durée))

                        Case "PRG"
                            'tbl2(0)="PRG"
                            'tbl2(1)=Numéro Piste
                            'tbl2(2)=Numéro de canal
                            'tbl2(3)=Position
                            'tbl2(4)=numPRG
                            Position = ((Val(tbl2(3)) / 4) + 1)
                            Horloge1.Schedule(New Midi.ProgramChangeMessage(SortieMidi.Item(ChoixSortieMidi), CByte(Val(tbl2(2))), CByte(Val(tbl2(4))), Position))
                        Case "CTRL"
                            'tbl(0)="CTRL"
                            'tbl(1)=Numéro Piste
                            'tbl(2)=Numéro de canal
                            'tbl(3)=Position (en nombre de double corche (16)
                            'tbl(4)=Numéro de controleur
                            'tbl(5)=Val CTRL
                            Position = ((Val(tbl2(3)) / 4) + 1)
                            Horloge1.Schedule(New Midi.ControlChangeMessage(SortieMidi.Item(ChoixSortieMidi), CByte(Val(tbl2(2))), CByte(Val(tbl2(4))), CInt(Val(tbl2(5))), Position))
                        Case "FIN"
                            'tbl(0)="FIN"
                            'tbl(1)=Numéro Piste
                            'tbl(2)=Position
                            If Not FinDemandée Then 'tbl2(1) = 0 Then
                                Position = ((Val(tbl2(2)) / 4) + 1)
                                Horloge1.Schedule(New CallbackMessage(AddressOf EVT_FIN, Position + 1))
                                FinDemandée = True
                            End If
                        Case Else
                    End Select
                Next j
            End If
        Next i
        ' 
        ' timer d'affichage des eventh
        ' ****************************
        NumAcc = 0
        Arrangement1.NumAccords.PointeurLect = 0 ' pour affichage des n° d'accords
        PointAffAccord = 0 ' 
        AffNumMes = 0 ' pointeur de lecteur de la liste ListNumMes contenu les Mesures à afficher
        Tempo_Aff_EventH.Enabled = False
        Tempo_Aff_EventH.Interval = 100
        Tempo_Aff_EventH.Start()
        Label6.Text = "---"

        ' Démarrage Play Back MIDI
        ' ************************
        'Fermer_MIDI()
        'Dim b As Boolean = EntréeMidi.Item(ChoixSortieMidi).IsOpen
        Try
            If Not (SortieMidi.Item(ChoixSortieMidi).IsOpen) Then
                SortieMidi.Item(ChoixSortieMidi).Open()
            End If
            '
            If Horloge1.IsRunning Then
                Horloge1.Stop()
                Horloge1.Reset()
            End If
            ' désactivation des boutons PLAY dans les 2 barres fixe et flottante
            PlayMidi.Enabled = False
            Transport.Button1.Enabled = False
            '
            If Remote.Checked Then Send_CTRL54_Remote()
            Horloge1.Start()

            '
        Catch ex As Exception
            Dim a As String = SortieMidi.Item(ChoixSortieMidi).Name
            If LangueIHM = "fr" Then
                MessageHV.PTitre = "Avertissement"
                MessageHV.PContenuMess = "Alarme : procédure 'PlayArp' : l'interface MIDI pourrait être occupée par une autre application." _
+ vbCrLf + "- choisissez une autre sortie MIDI " + "(" + a + ")" _
+ vbCrLf + "- ou une autre application MIDI pourrait être présente : libérez cette application ou mettez la en tâche de fond," _
+ vbCrLf + "- ou redémarrez votre PC." _
+ vbCrLf + ex.Message
                MessageHV.PTypBouton = "OK"
            Else
                MessageHV.PTitre = "Warning"
                MessageHV.PContenuMess = "Alarme : procédure 'PlayArp :  MIDI interface could be occupied by another application." _
+ vbCrLf + "- choose another MIDI output," + "(" + a + ")" _
+ vbCrLf + "- or another MIDI application might be present: release this application or put it in the background," _
+ vbCrLf + "- or reboot your PC." _
+ vbCrLf + ex.Message
                MessageHV.PTypBouton = "OK"
            End If
            Cacher_FormTransparents()
            MessageHV.ShowDialog()
        End Try
    End Sub
    Sub EVT_Acc()
        CallB_Aff_Acc = True
    End Sub
    Sub EVT_NumMes()
        CallB_Aff_NumMes = True
    End Sub
    Sub EVT_FIN()
        CallB_Aff_Acc = False ' par précaution --> on ne peut pas être à la fois sur un évènement Accord et un évènement FIN
        CallB_Aff_NumMes = False
        CallB_Aff_FIN = True
    End Sub
    Function Det_PremPisteActive() As Integer
        Det_PremPisteActive = -1
        For i = 0 To nb_PistesVar - 1 ' Arrangement1.Nb_PistesMidi - 2 ' on met _1 car ici on ne travaille que sur les pistes des variation ( pas sur le PianoRoll)
            If LesPistes.Item(i).part.Count <> 0 Then
                Det_PremPisteActive = i
                Exit For
            End If
        Next
    End Function
    Sub NumMesList()
        Dim ListNumes As New List(Of String) ' pour le test
        Dim i As Integer
        Dim NumMes As Integer = 0
        Dim position As Single


        For i = 1 To (nbMesures * 16) Step 16
            NumMes = NumMes + 1

            position = i
            ListNumes.Add("NumMes" + " " + Convert.ToString(NumMes) + " " + Convert.ToString(position)) ' + " " + Convert.ToString(NumAcc) + " " + Convert.ToString(Me.PositCours))
            position = (position / 4) + 1
            Horloge1.Schedule(New CallbackMessage(AddressOf EVT_NumMes, position))
        Next
    End Sub
    Sub NumMesList2()
        Dim ListNumes As New List(Of String) ' pour le test
        Dim i, k As Integer
        Dim _boucle As Integer = Det_Répet(False)
        Dim NumMes As Integer
        Dim position As Integer
        Dim DerAcc As Integer = Det_NumDerAccord()
        Dim RepDerAcc As Integer = Convert.ToInt16(Grid2.Cell(3, DerAcc).Text) - 1
        Dim NbMesDerAcc As Integer = DerAcc + (RepDerAcc - 1)

        For k = 0 To _boucle
            NumMes = 0
            For i = 1 To ((DerAcc + RepDerAcc) * 16) Step 16
                NumMes = NumMes + 1

                position = i - 1
                ListNumes.Add("NumMes" + " " + Convert.ToString(NumMes) + " " + Convert.ToString(position)) ' + " " + Convert.ToString(NumAcc) + " " + Convert.ToString(Me.PositCours))
                position = (position / 4) + 1
                Horloge1.Schedule(New CallbackMessage(AddressOf EVT_NumMes, position))
            Next i

        Next k
    End Sub
    ''' <summary>
    ''' NumMesList22 : création d'une liste permettant d'afficher
    ''' les N0 de mesures lors de l'exécution d'un morceau.
    ''' </summary>
    Sub NumMesList22()
        Dim ListNumes As New List(Of String) ' pour le test
        Dim i, j, k As Integer
        Dim _boucle As Integer = Det_Répet(False)
        Dim NumMes As Integer
        Dim position, position2 As Integer
        Dim RepDerAcc As Integer = 0


        Dim DerAcc As Integer = Terme.Value 'Det_NumDerAccord()


        If DerAcc = Det_NumDerAccord() Then
            RepDerAcc = Convert.ToInt16(Grid2.Cell(3, DerAcc).Text) - 1
        End If
        ' Dim NbMesDerAcc As Integer = DerAcc + (RepDerAcc - 1)
        Dim Début2 As Integer = Début.Value - 1

        i = 0
        j = 0
        ListNumMes.Clear()
        '
        'Det_RépetSuiv(AccDépart As Integer) As Integer



        For k = 0 To _boucle
            NumMes = Début2
            i = 0
            Do
                NumMes = NumMes + 1
                ListNumMes.Add(Convert.ToString(NumMes)) ' pour lecture par la tempo de Call back
                position = j
                position2 = (position / 4) + 1 ' position pour schedule
                'ListNumes.Add("NumMes" + " " + Convert.ToString(NumMes) + " " + Convert.ToString(position) + " " + Convert.ToString(position2)) ' + " " + Convert.ToString(NumAcc) + " " + Convert.ToString(Me.PositCours))
                Horloge1.Schedule(New CallbackMessage(AddressOf EVT_NumMes, position2))
                i = i + 16
                j = j + 16
            Loop Until i >= ((DerAcc - Début2) + RepDerAcc) * 16
        Next k
    End Sub
    '
    ' *#######################################################################
    ' ##                                                                    ##
    ' ##                               ONGLETS                              ##
    ' ##                                                                    ##
    ' ########################################################################
    '
    Sub EcritureAccordDsGrid2(Acc As String, Colonne As Integer)
        Dim m As Integer
        Dim t As Integer
        Dim ct As Integer
        Dim tbl() As String
        Dim a As String
        Dim ligne2 As Integer
        Dim degré As Integer
        Dim Indexdegré As Integer
        Dim Acc1 As String
        '
        If Trim(Acc) <> "" Then
            '
            DerGridCliquée = GridCours.Grid2
            '
            m = Colonne
            t = 1 ' 
            ct = 1 '
            '
            Acc1 = Acc
            '
            TableEventH(m, t, ct).Racine = Trim(TRacine.Text)
            If TabControl2.SelectedIndex = 0 Then ' Onglet 1 = Tonalités  Then
                Maj_PropriétésEntrée2() ' mise à jour de Entrée_Tonalité n Entrée_Mode et Entrée_Gamme (ici au moment de l'entrée de l'acccord Mode=Gamme)
                TableEventH(m, t, ct).Accord = Trim(Acc) ' Entrée_Accord
                TableEventH(m, t, ct).Gamme = Trim(Entrée_Gamme) ' maj dans Maj_PropriétésEntrée2
                TableEventH(m, t, ct).NumAcc = m
                TableEventH(m, t, ct).Ligne = m
                'TableEventH(m, t, ct).Racine = "c2"

                If Flag_EcrDragDrop = True Then
                    TableEventH(m, t, ct).Mode = Trim(Entrée_Mode) ' Mode = Gamme - le mode n'est pas affiché mais je garde l'info por le moment
                    TableEventH(m, t, ct).Tonalité = Trim(Entrée_Tonalité) 'Trim(ComboBox1.Text) 'Entrée_Tonalité
                    '
                    Entrée_Position = Trim(Colonne.ToString + ".1" + ".1")
                    TableEventH(m, t, ct).Position = Trim(Entrée_Position)
                    '
                    ligne2 = Det_LigneTableGlobale(LabelCours) ' LabelCours = variable globale mise à jour dans TabTon_MouseDown - ce calcul vient de HyperVoicing (je le garde, pas de pb mais à voir plustard)
                    degré = Det_IndexDegré2(LabelCours)
                    TableEventH(m, t, ct).Degré = degré
                    TableEventH(m, t, ct).Ligne = m 'ligne2
                    '
                End If
            Else
                ' Ecriture accords à partir du tableau des progressions Onglet Progression
                ' ************************************************************************
                '
                If TabControl2.SelectedIndex = 1 Then
                    TableEventH(m, t, ct).Accord = Trim(Acc)
                    TableEventH(m, t, ct).Gamme = Entrée_Gamme
                    TableEventH(m, t, ct).NumAcc = m
                    If Flag_EcrDragDrop = True Then
                        'ligne2 = Det_LigneTableGlobale(LabelCours)
                        TableEventH(m, t, ct).Mode = Entrée_Gamme ' Mode = Gamme
                        TableEventH(m, t, ct).Tonalité = Entrée_Tonalité 'Trim(ComboBox1.Text) 'Entrée_Tonalité
                        '
                        Entrée_Position = Trim(Str(Colonne) + ".1" + ".1")
                        TableEventH(m, t, ct).Position = Entrée_Position
                        '
                        ' Calcul du degré
                        ' ***************
                        Indexdegré = CAD_LabelCours 'Det_IndexDegré(a)
                        TableEventH(m, t, ct).Degré = Indexdegré
                        If Indexdegré = 2 And OrigineAccord = Modes.Cadence_Mixte Then
                            TableEventH(m, t, ct).Degré = 4
                        End If
                        '
                    End If
                Else
                    If TabControl2.SelectedIndex = 2 Or TabControl2.SelectedIndex = 3 Then
                        ' Position
                        Entrée_Position = Trim(Str(Colonne) + ".1" + ".1")
                        TableEventH(m, t, ct).Position = Trim(Entrée_Position)
                        ' Degré
                        TableEventH(m, t, ct).Degré = Entrée_Degré
                        ' Tonalité
                        TableEventH(m, t, ct).Tonalité = Entrée_Tonalité
                        ' Accord
                        TableEventH(m, t, ct).Accord = Trim(Entrée_Accord)
                        ' Mode
                        TableEventH(m, t, ct).Mode = Entrée_Mode
                        ' Gamme
                        TableEventH(m, t, ct).Gamme = Entrée_Gamme
                        '
                    End If
                End If
            End If


            ' Traitement global : ECRITURE Grid2
            ' **********************************
            Grid2.Cell(2, m).Alignment = FlexCell.AlignmentEnum.CenterCenter
            Grid2.Cell(11, m).Alignment = FlexCell.AlignmentEnum.CenterCenter
            Grid2.Cell(2, m).Text = Acc1
            Grid2.Cell(11, m).Text = TableEventH(m, t, ct).Gamme
            Grid2.Cell(12, m).Text = TableEventH(m, t, ct).Racine
            '
            Grid2.ReadonlyFocusRect = FlexCell.FocusRectEnum.Solid
            '
            Grid2.AutoRedraw = False
            '
            a = TableEventH(m, 1, 1).Tonalité '
            a = Det_RelativeMajeure2(a) ' si la tonalité est mineure alors on affiche la couleur de la relative majeure
            tbl = Split(a)

            Grid2.Cell(2, m).BackColor = DicoCouleur.Item(Trim(tbl(0))) ' la couleur de l'accord est fonction de la tonalité
            Grid2.Cell(2, m).ForeColor = DicoCouleurLettre.Item(tbl(0))
            Grid2.Cell(11, m).BackColor = DicoCouleur.Item(Trim(tbl(0))) ' la couleur de la gamme est fonction de la tonalité
            Grid2.Cell(11, m).ForeColor = DicoCouleurLettre.Item(tbl(0))
            If m = 1 Then
                Grid2.Cell(2, m).BackColor = Color.Red
                Grid2.Cell(2, m).ForeColor = Color.Yellow
                Grid2.Cell(11, m).BackColor = Color.Red
                Grid2.Cell(11, m).ForeColor = Color.Yellow
            End If
            '
            Formater_Racine(m) ' formater la racine : couleur et alignement

            Grid2.Refresh()
            Grid2.AutoRedraw = True
        End If
        Maj_Nligne() ' met à jour le N° de ligne dans TableEventH pour chaque EVENTH entrés (<> -1)
    End Sub
    Private Sub TabTonsFiltres_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        DerGridCliquée = GridCours.TabTon
    End Sub
    Private Sub TabTonsFiltres_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        Dim com As CheckBox = sender
        Dim i As Integer
        Dim Typ11 As Integer
        Dim Typ9 As Integer
        Dim Typsus4 As Integer

        Dim ind As Integer
        Dim a As String
        Dim chiff() As Object
        Dim Tonalité As Integer
        Dim Degré As String
        Dim IndexDegré As Integer

        '
        ind = Fix(Val(com.Tag) / 2)
        a = Trim(TabTons.Item(ind).Text)
        Typ11 = InStr(a, "11")
        Typsus4 = InStr(a, "sus4")
        Typ9 = InStr(a, "9")

        ' Accord 9e
        ' *********
        If Typ9 <> 0 Then
            chiff = Split(a)
            Select Case Trim(chiff(1))
                Case "7M(9)", "M7(9)", "7(9)"
                    TabTons.Item(ind).Text = Trim(Det_Tonique(a)) + " 9"
                Case "m7(9)"
                    TabTons.Item(ind).Text = Trim(Det_Tonique(a)) + " m9"
                Case "7(b9)"
                    TabTons.Item(ind).Text = Trim(Det_Tonique(a)) + " b9"
                Case "7(9#)"
                    TabTons.Item(ind).Text = Trim(Det_Tonique(a)) + " 9#"
                Case "9", "m9", "b9", "9#"
                    Tonalité = Det_LigneTableGlobale(ind)
                    i = Det_IndexDansLigne(ind)
                    Degré = TabTonsDegrés.Item(i).Text
                    IndexDegré = Det_IndexDegré(Degré)
                    a = TableGlobalAcc(2, Tonalité, IndexDegré)
                    TabTons.Item(ind).Text = TableGlobalAcc(2, Tonalité, IndexDegré)
            End Select
            '
            'Maj_Renversement(ind)
        End If
        '
        ' Accord 11e
        ' **********
        If Typ11 <> 0 Or Typsus4 <> 0 Then
            chiff = Split(a)
            If EstPair(Val(com.Tag)) Then
                ' filtre sus 4
                Select Case Trim(chiff(1))
                    Case "7(11)"
                        TabTons.Item(ind).Text = Trim(Det_Tonique(a)) + " 7sus4"
                    Case "11"
                        TabTons.Item(ind).Text = Trim(Det_Tonique(a)) + " sus4"
                    Case "7sus4"
                        TabTons.Item(ind).Text = Trim(Det_Tonique(a)) + " 7(11)"
                    Case "sus4"
                        TabTons.Item(ind).Text = Trim(Det_Tonique(a)) + " 11"
                End Select
            Else
                ' filtre 7
                Select Case Trim(chiff(1))
                    Case "7M(#11)", "M7(11#)"
                        TabTons.Item(ind).Text = Trim(Det_Tonique(a)) + " 11#"
                    Case "11#"
                        TabTons.Item(ind).Text = Trim(Det_Tonique(a)) + " M7(11#)"
                    Case "7(11)"
                        TabTons.Item(ind).Text = Trim(Det_Tonique(a)) + " 11"
                    Case "11"
                        TabTons.Item(ind).Text = Trim(Det_Tonique(a)) + " 7(11)"
                    Case "7sus4"
                        TabTons.Item(ind).Text = Trim(Det_Tonique(a)) + " sus4"
                    Case "sus4"
                        TabTons.Item(ind).Text = Trim(Det_Tonique(a)) + " 7sus4"

                End Select
            End If
            'Maj_Renversement(ind)
        End If
    End Sub
    Sub CAD_Création()
        Dim i, j, k As Integer

        Dim AjoutRéglage1, AjoutRéglage2 As Integer

        Dim PosX, PosY As Integer
        '
        Dim Longueur, Hauteur As Integer
        '
        Dim DepL, DepH As Integer
        '
        PosY = 72 '28
        PosX = 2 '
        '
        Longueur = 100
        Hauteur = 180
        '
        DepL = 0
        DepH = 10
        j = 0
        k = 1
        i = 0
        '
        AjoutRéglage1 = -5
        AjoutRéglage2 = -20
        Do '
            If i < 5 Then
                TabCadDegrés.Add(New Label)
                'Me.Panel1.Controls.Add(TabCadDegrés.Item(i))
                Me.Panel4.Controls.Add(TabCadDegrés.Item(i))
                TabCadDegrés.Item(i).Size = New System.Drawing.Size(Longueur, Hauteur / 4)
                TabCadDegrés.Item(i).BorderStyle = BorderStyle.FixedSingle
                TabCadDegrés.Item(i).TextAlign = ContentAlignment.MiddleCenter
                TabCadDegrés.Item(i).Tag = Str(i)
                TabCadDegrés.Item(i).Visible = False
                TabCadDegrés.Item(i).Font = New System.Drawing.Font("Arial", 12, FontStyle.Bold)
                TabCadDegrés.Item(i).Location = New Point(PosX + DepL, PosY + DepH + AjoutRéglage1)
                TabCadDegrés.Item(i).BackColor = Color.Khaki
                TabCadDegrés.Item(i).ForeColor = Color.Black
                TabCadDegrés.Item(i).BorderStyle = BorderStyle.Fixed3D
                TabCadDegrés.Item(i).Tag = i
            Else
                ' labels
                ' ******
                j = i - 5
                TabCad.Add(New Label)
                'Me.Panel1.Controls.Add(TabCad.Item(j))
                Me.Panel4.Controls.Add(TabCad.Item(j))
                TabCad.Item(j).Size = New System.Drawing.Size(Longueur, Hauteur / 2)
                TabCad.Item(j).BorderStyle = BorderStyle.FixedSingle
                TabCad.Item(j).TextAlign = ContentAlignment.MiddleCenter
                TabCad.Item(j).Tag = Str(j)
                TabCad.Item(j).Visible = True
                TabCad.Item(j).Font = New System.Drawing.Font("Arial", 12, FontStyle.Bold)
                TabCad.Item(j).BackColor = Couleur_Accord_Majeur
                TabCad.Item(j).Location = New Point(PosX + DepL, PosY + DepH + AjoutRéglage2)
                TabCad.Item(j).Tag = j
                '
                AddHandler TabCad.Item(j).MouseDown, AddressOf TabCad_MouseDown
                AddHandler TabCad.Item(j).MouseUp, AddressOf TabCad_MouseUp
            End If


            DepL = DepL + Longueur + 8
            i = i + 1
            If i = 5 Then
                DepH = DepH + Hauteur - 85
                DepL = 0
            End If
        Loop Until i > 9 ' Il y a 10 label à créer (2fois 5)
    End Sub






    Private Sub TabTCad_MouseUp(ByVal sender As Object, ByVal e As MouseEventArgs)
        Dim ind As Integer
        Dim com As Label = sender
        '
        ind = Val(com.Tag)
    End Sub
    Private Sub TabCadFiltres4_MouseUp(ByVal sender As Object, ByVal e As MouseEventArgs)
        Dim ind As Integer
        Dim com As CheckBox = sender
        Dim a As String
        Dim chiff() As String
        '
        ind = Val(com.Tag)
        a = Trim(TabCad.Item(ind).Text)
        chiff = Split(a)
        '
        Select Case Trim(chiff(1))
            Case "7(11)"
                TabCad.Item(ind).Text = Trim(Det_Tonique(a)) + " 7sus4"
            Case "11"
                TabCad.Item(ind).Text = Trim(Det_Tonique(a)) + " sus4"
            Case "7sus4"
                TabCad.Item(ind).Text = Trim(Det_Tonique(a)) + " 7(11)"
            Case "sus4"
                TabCad.Item(ind).Text = Trim(Det_Tonique(a)) + " 11"
        End Select
    End Sub
    Sub MODULATION_CREATION()
        Dim i, j As Integer
        Dim p As New Point
        Dim s As New Size
        '
        ' Maj distance du splitter du splitcontainer8
        ' *******************************************
        SplitContainer8.SplitterDistance = 117

        ' Remarque : Combobox9 est mis à jour dans fr_Culture et En_Culture
        ' *****************************************************************
        '
        ' Combobox14 : liste des cadences
        ' *******************************
        If LangueIHM = "fr" Then
            ComboBox13.Items.Add("Cadence Parfaite - V I")
            ComboBox13.Items.Add("Cadence plagale - IV I")
            ComboBox13.Items.Add("Demi_cadence - II V")
            ComboBox13.Items.Add("Demi_cadence - IV V")
            ComboBox13.Items.Add("Demi_cadence - I V")
            ComboBox13.Items.Add("Cadence rompue - V VI")
        Else
            ComboBox13.Items.Add("Perfect cadence - V I")
            ComboBox13.Items.Add("Plagal cadence - IV I")
            ComboBox13.Items.Add("Half_cadence - II V")
            ComboBox13.Items.Add("Half_cadence - IV V")
            ComboBox13.Items.Add("Half_cadence - I V")
            ComboBox13.Items.Add("Decept. cadence - V VI")
        End If
        '
        ComboBox13.SelectedIndex = 0
        '
        ' Construction de la grille
        ' *************************
        'Grid5.AllowDrop = True
        '
        Grid5.Cols = 8
        Grid5.Rows = 20
        '
        Grid5.SelectionMode = FlexCell.SelectionModeEnum.ByRow
        Grid5.Selection.BackColor = Color.Transparent
        Grid5.SelectionBorderColor = Color.Transparent
        '
        Grid5.DefaultFont = New Font("Verdana", 8, FontStyle.Regular)
        Grid5.FixedCols = 1
        Grid5.FixedRows = 2
        '
        Grid5.ScrollBars = ScrollBarsEnum.Both
        '
        For i = 0 To 1
            For j = 0 To Grid5.Cols - 1
                Grid5.Cell(i, j).ForeColor = Color.Black

            Next
        Next
        '
        For i = 0 To Grid5.Rows - 1
            Grid5.Cell(i, 0).ForeColor = Color.Black
        Next
        '
        Grid5.Row(0).Height = 15
        Grid5.Row(1).Height = 35
        '
        For i = 0 To 1
            For j = 1 To Grid5.Cols - 1
                Grid5.Cell(i, j).Font = New Font("Verdana", 8, FontStyle.Regular)
            Next
        Next
        '
        Grid5.SelectionBorderColor = Color.Black
        Grid5.BorderColor = Color.Black
        '
        Grid5.Column(0).Width = 55
        '
        For j = 1 To Grid5.Cols - 1
            Grid5.Column(j).Width = 68
            Grid5.Column(j).Alignment = AlignmentEnum.CenterCenter
        Next
        ''
        For i = 1 To Grid5.Rows - 1
            Grid5.Row(i).Height = 25
        Next
        '
        Aff_Degrés()
        ' Ecriture de "Mode Actuel"
        ' ************************
        Grid5.Range(0, 0, 1, 0).MergeCells() = True
        Grid5.Cell(0, 0).WrapText = True
        Grid5.Cell(0, 0).BackColor = Grid5.BackColorFixed 'Color.Gainsboro
        Grid5.Cell(0, 0).Font = New Font("Verdana", 8, FontStyle.Italic)
        If Module1.LangueIHM = "fr" Then
            Grid5.Cell(0, 0).Text = "Mode" + " départ"
            Label35.Text = "Mode départ"
            Label36.Text = "Accord pivot"
            Label40.Text = "Facultatif"
        Else
            Grid5.Cell(0, 0).Text = "Start" + " Mode"
            Label35.Text = "Start mode"
            Label36.Text = "Pivot chord"
            Label40.Text = "Optional"
        End If

        '
        ' Construction de radiobuttons
        ' ****************************
        '
        'p.Y = 7
        'For i = 0 To 2
        'RadioModulat.Add(New RadioButton)
        'RadioModulat.Item(i).Font = New Font("Verdana", 8, FontStyle.Regular)
        'SplitContainer5.Panel1.Controls.Add(RadioModulat.Item(i))
        'RadioModulat.Item(i).Tag = i
        'RadioModulat.Item(i).AutoSize = True
        ''
        'Select Case i
        'Case 0
        'p.X = 192
        'RadioModulat.Item(i).Text = "C Maj"
        'RadioModulat.Item(i).Checked = True
        'Case 1
        'p.X = 262
        'RadioModulat.Item(i).Text = "A MinH"
        'RadioModulat.Item(i).Checked = False
        'Case 2
        'p.X = 340
        'RadioModulat.Item(i).Text = "A MinM"
        'RadioModulat.Item(i).Checked = False
        'End Select
        '
        'RadioModulat.Item(i).Location = p
        'AddHandler RadioModulat.Item(i).CheckedChanged, AddressOf RadioModulat_CheckedChanged
        'Next i
        '
        ' Création des 3 labels pour le système de modulation (LabModulat.Item(0) est inutilisé)
        ' ***************************************************

        s.Width = 83
        s.Height = 50
        '
        For i = 0 To 3
            LabModulat.Add(New Label)
            SplitContainer8.Panel2.Controls.Add(LabModulat.Item(i))
            LabModulat.Item(i).AllowDrop = True
            LabModulat.Item(i).AutoSize = False
            LabModulat.Item(i).Size = s

            LabModulat.Item(i).BackColor = Color.Moccasin

            LabModulat.Item(i).BorderStyle = BorderStyle.FixedSingle
            LabModulat.Item(i).TextAlign = ContentAlignment.MiddleCenter
            LabModulat.Item(i).Font = New Font("Verdana", 8, FontStyle.Regular)
            '"Label" + Str(i)
            LabModulat.Item(i).Tag = i
            LabModulat.Item(i).Refresh()
            LabModulat.Item(i).Visible = True
            LabModulat.Item(i).Text = "---"

            '
            AddHandler LabModulat.Item(i).MouseDown, AddressOf LabModulat_MouseDown
            AddHandler LabModulat.Item(i).MouseUp, AddressOf LabModulat_MouseUp
        Next i
        '
        'LabModulat.Item(0).Visible = False '  LabModulat.Item(0) est inutilié
        '
        ' POSITIONNEMENTS
        ' ***************

        Dim f As New System.Drawing.Font("Verdana", 8, FontStyle.Regular)
        '
        ' AccordPivot
        ' ***********
        LabModulat.Item(0).Location = Label4.Location  ' Accord Nouvelle tona
        LabModulat.Item(0).Size = Label4.Size
        LabModulat.Item(0).Font = f
        LabModulat.Item(0).BackColor = Label4.BackColor
        LabModulat.Item(0).BringToFront()
        '
        ' Accord Nouvelle Tona
        ' ********************
        LabModulat.Item(1).Location = Label34.Location  ' Accord Nouvelle tona
        LabModulat.Item(1).Size = Label34.Size
        LabModulat.Item(1).Font = f
        LabModulat.Item(1).BackColor = Label34.BackColor
        LabModulat.Item(1).BringToFront()
        '
        ' Cadence
        ' *******
        LabModulat.Item(2).Location = Label7.Location   ' cadence V
        LabModulat.Item(2).Size = Label7.Size
        LabModulat.Item(2).Font = f
        LabModulat.Item(2).BackColor = Label7.BackColor
        LabModulat.Item(2).BringToFront()
        '
        LabModulat.Item(3).Location = Label200.Location  ' cadence I
        LabModulat.Item(3).Size = Label200.Size
        LabModulat.Item(3).Font = f
        LabModulat.Item(3).BackColor = Label200.BackColor
        LabModulat.Item(3).BringToFront()
        '
        Label79.Font = f
        '
        If LangueIHM = "fr" Then
            Label21.Text = "Accord pivot"
            Label38.Text = "2e Accord(optionnel)"
        Else
            Label21.Text = "Pivot chord"
            Label38.Text = "2nd Pivot chord(optional)"
        End If
        ' Mise jour des couleurs
        ' **********************
        EffacerGrid5()

    End Sub
    Private Sub Maj_ModulationRadioB2(colonne As Integer)
        Dim Tonique As String
        Dim typ As Integer = ComboBox9.SelectedIndex + 3
        If Trim(Grid2.Cell(2, colonne).Text) <> "" Then
            Dim tonalité As String = Trim(Grid2.Cell(11, colonne).Text) 'identification de la tonalité de départ
            Dim accord As String = Trim(Grid2.Cell(2, colonne).Text) 'identification de l'accord commun de départ
            Dim tbl() As String = tonalité.Split()
            Dim Maj, Min As String
            Dim listMaj As New List(Of String) From {"C#", "F#", "B", "E", "A", "D", "G", "C", "F", "Bb", "Eb", "Ab"}
            Dim listMin As New List(Of String) From {"A#", "D#", "G#", "C#", "F#", "B", "E", "A", "D", "G", "C", "F"}
            '
            If Trim(tbl(1)) = "Maj" Then
                Maj = Trim(tbl(0))
                Min = listMin(listMaj.IndexOf(Maj))
            Else
                Min = Trim(tbl(0))
                Maj = listMin(listMaj.IndexOf(Min))
            End If
            '
            ' Maj de la tonalité de départ
            ' ****************************
            Label32.Text = tonalité
            Label33.Text = accord
            '
            ' identification de l'accord
            ' **************************
            Dim acc = Trim(Grid2.Cell(2, colonne).Text)
            '
            ' Ecriture de la tonalité de départ (sur ligne 0)
            ' *******************************************
            ' 
            tbl = Label32.Text.Split
            Tonique = Retab_Mode(Trim(tbl(0)), Trim(tbl(1)))
            a = Mode3(Trim(Tonique), Trim(tbl(1)), typ, False)
            tbl = a.Split("-")
            For j = 1 To tbl.Count
                Grid5.Cell(1, j).Text = tbl(j - 1)
            Next

            ' 
            ' Mise à jour des tonalités voisines de la tonalité principale
            ' ************************************************************
            EffacerTonVoisins()
            Maj_TonsVoisins(acc)
            RecheraccordGrid5(acc)
            '
        End If
    End Sub
    Sub EffacerGrid5()
        Dim i, j As Integer

        ' Effacer la grille Grid5
        ' ***********************
        Lab_1 = "---"
        Lab_2 = "---"
        LabModulat.Item(0).Text = "---" ' non utilisée
        LabModulat.Item(1).Text = "---"
        LabModulat.Item(2).Text = "---"
        LabModulat.Item(3).Text = "---"
        'Label79.Text = "---"
        '
        Label79.ForeColor = Color.Red
        Label80.Visible = False
        If LangueIHM = "fr" Then
            Label79.Text = "Pour obtenir des accords de transition, veuillez cliquer sur la ligne du mode choisi."
        Else
            Label79.Text = "To obtain transition chords, click on the line of the chosen mode."
        End If
        '
        For i = 0 To Grid5.Rows - 1
            For j = 0 To Grid5.Cols - 1
                ' Effacer toute les lignes sauf la ligne des degrés
                If i > 0 Then
                    Grid5.Cell(i, j).Text = ""
                End If
                ' Cellules contenant les tonalités d'appartenance
                If i > 1 And j > 0 Then
                    Grid5.Cell(i, j).BackColor = Color.Moccasin
                End If
            Next
        Next
    End Sub
    Sub Aff_Degrés()
        Grid5.Cell(0, 1).Text = "I"
        Grid5.Cell(0, 2).Text = "II"
        Grid5.Cell(0, 3).Text = "III"
        Grid5.Cell(0, 4).Text = "IV"
        Grid5.Cell(0, 5).Text = "V"
        Grid5.Cell(0, 6).Text = "VI"
        Grid5.Cell(0, 7).Text = "VII"
    End Sub
    Private Sub LabModulat_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        Dim com As Label = sender
        Dim ind As Integer = Val(com.Tag) ' numéro du label
        Dim a As String
        Dim tbl() As String

        ' Jouer l'accord avec CTRL + clic
        ' *******************************
        If e.Button() = MouseButtons.Left And My.Computer.Keyboard.CtrlKeyDown Then
            a = Trim(LabModulat(ind).Text)
            If a <> "" And a <> "---" Then
                JouerAccord(a)
            End If
        End If

        ' Mise en place 
        ' *************
        If Not My.Computer.Keyboard.CtrlKeyDown And Trim(LabModulat.Item(ind).Text) <> "---" Then
            Select Case Det_TonInitial()
                Case 0
                    OrigineAccord = Modes.Majeur
                Case 1
                    OrigineAccord = Modes.MineurH
                Case 2
                    OrigineAccord = Modes.MineurM
            End Select
            ' Maf des Entrées
            ' ***************
            If Trim(Lab_2) <> "---" And e.Button() = MouseButtons.Left And Not (My.Computer.Keyboard.CtrlKeyDown) _
            And Not (My.Computer.Keyboard.AltKeyDown) _
            And Not (My.Computer.Keyboard.ShiftKeyDown) Then ' 
                AccordMarqué = Trim(LabModulat.Item(ind).Text)
                Entrée_Accord = Trim(AccordMarqué)                   ' Entrée_Accord
                Entrée_Mode = Trim(Lab_2)                            ' Entrée_Mode
                Entrée_Gamme = Trim(Lab_2)                           ' Entrée_Gamme
                Entrée_Degré = Det_DegréDécimal(LigneCoursGrid5, Trim(AccordMarqué)) 'Det_DegréRomain(Grid5.ActiveCell.Row, Trim(AccordMarqué))  ' Entrée_Dégré
                tbl = Entrée_Mode.Split()                            ' Entrée_Tonalité
                If tbl(1) = "MinH" Or tbl(1) = "MinM" Then
                    Entrée_Tonalité = Det_RelativeMajeure2(Trim(Lab_2))
                Else
                    Entrée_Tonalité = Trim(Lab_2)
                End If
                '
                LabModulat.Item(ind).DoDragDrop(Trim(LabModulat.Item(ind).Text), DragDropEffects.Copy Or DragDropEffects.Move)
                If Trim(Valeur_Drag) <> "" Then
                    Maj_DragDrop()
                End If
                '
            End If
        End If
    End Sub
    Private Sub LabModulat_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        '
        If AccordAEtéJoué = True Then
            CouperJouerAccord()
            AccordAEtéJoué = False
        End If
        '
        RAZ_AffNoteAcc()
    End Sub
    Private Sub RadioModulat_CheckedChanged(sender As Object, e As EventArgs)
        Maj_Modulation()
    End Sub



    Sub SUBSTITUTION_CREATION()
        Dim i As Integer
        Dim p As New Point
        Dim s As New Size

        Label67.Text = "---"
        Label73.Text = ""

        ' Cration des Labels
        ' ******************
        For i = 0 To 3
            LabSubsti.Add(New Label)
            Panel14.Controls.Add(LabSubsti.Item(i))
            LabSubsti.Item(i).BackColor = Color.Goldenrod
            LabSubsti.Item(i).ForeColor = Color.Black
            LabSubsti.Item(i).Font = New Font("Arial Narrow", 18, FontStyle.Bold)
            LabSubsti.Item(i).TextAlign = ContentAlignment.MiddleCenter
            LabSubsti.Item(i).Text = "---"
            LabSubsti.Item(i).BringToFront()
            LabSubsti.Item(i).Tag = i

            Select Case i
                Case 0
                    p.X = 2
                    p.Y = 40
                    s.Width = 175
                    s.Height = 131
                    LabSubsti.Item(i).Location = p
                    LabSubsti.Item(i).Size = s
                Case 1
                    p.X = 191
                    p.Y = 40
                    s.Width = 175
                    s.Height = 60
                    LabSubsti.Item(i).Location = p
                    LabSubsti.Item(i).Size = s
                Case 2
                    p.X = 191
                    p.Y = 112
                    s.Width = 175
                    s.Height = 60
                    LabSubsti.Item(i).Location = p
                    LabSubsti.Item(i).Size = s
                Case 3
                    p.X = 376
                    p.Y = 40
                    s.Width = 175
                    s.Height = 131
                    LabSubsti.Item(i).Location = p
                    LabSubsti.Item(i).Size = s
            End Select
            AddHandler LabSubsti.Item(i).MouseDown, AddressOf LabSubsti_MouseDown
            AddHandler LabSubsti.Item(i).MouseUp, AddressOf LabSubsti_MouseUp
        Next
        If LangueIHM = "fr" Then
            Label70.Text = "Ton mineur"
            Label71.Text = "Diatonique"
            Label72.Text = "V de tonique"
        Else
            Label70.Text = "Minor key"
            Label71.Text = "Diatonic"
            Label72.Text = "V of key"
        End If
    End Sub
    Private Sub LabSubsti_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        Dim com As Label = sender
        Dim ind As Integer = Val(com.Tag) ' numéro du label
        Dim a As String
        Dim tbl() As String


        ' Jouer l'accord avec CTRL + clic
        ' *******************************
        If e.Button() = MouseButtons.Left And My.Computer.Keyboard.CtrlKeyDown Then
            a = Trim(LabSubsti(ind).Text)
            If a <> "" And a <> "---" Then
                'JouerAccordLabel(a)
            End If
        End If

        ' Mise en place du glisser/déposer
        ' ********************************
        If Not My.Computer.Keyboard.CtrlKeyDown And Trim(LabSubsti.Item(ind).Text) <> "---" And Trim(LabSubsti.Item(ind).Text) <> "N/A" Then
            tbl = EventhSubsti(0).Mode.Split
            Select Case Trim(tbl(1))
                Case "Maj"
                    OrigineAccord = Modes.Majeur
                Case "MinH"
                    OrigineAccord = Modes.MineurH
                Case "MinM"
                    OrigineAccord = Modes.MineurM
            End Select
            '
            ' Maf des Entrées pour le glisser déposer/déposer
            ' ***********************************************
            If e.Button() = MouseButtons.Left And Not (My.Computer.Keyboard.CtrlKeyDown) _
            And Not (My.Computer.Keyboard.AltKeyDown) _
            And Not (My.Computer.Keyboard.ShiftKeyDown) Then
                Valeur_Drag = ""
                AccordMarqué = Trim(LabSubsti.Item(ind).Text)
                Entrée_Accord = Trim(AccordMarqué)                         ' Entrée_Accord
                Entrée_Mode = EventhSubsti(ind).Mode                       ' Entrée_Mode
                Entrée_Gamme = EventhSubsti(ind).Gamme                     ' Entrée_Gamme
                Entrée_Tonalité = EventhSubsti(ind).Tonalité               ' Entrée_Tonalité
                Entrée_Degré = EventhSubsti(ind).Degré                     ' Entrée_Degré
                '
                LabSubsti.Item(ind).DoDragDrop(Trim(LabSubsti.Item(ind).Text), DragDropEffects.Copy Or DragDropEffects.Move)
                If Trim(Valeur_Drag) <> "" Then
                    Maj_DragDrop()
                End If
            End If
        End If

    End Sub
    Private Sub LabSubsti_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)

    End Sub
    Private Function Det_TonInitial() As Integer
        Dim ind As Integer = 0
        Dim tbl() As String = Label32.Text.Split()

        If tbl(1) = "Maj" Then
            OrigineAccord = Modes.Majeur
            ind = 0
        Else
            If tbl(1) = "MinH" Then
                OrigineAccord = Modes.MineurH
                ind = 1
            Else
                If tbl(1) = "MinM" Then
                    OrigineAccord = Modes.MineurM
                    ind = 2
                End If
            End If
        End If
        '

        Return ind
    End Function
    Private Sub Old_Maj_ModulationRadioB()
        Dim tbl1() As String = Trim(ComboBox1.Text).Split()
        Dim tbl2() As String = Trim(ComboBox2.Text).Split()
        Dim NoteM As String = Trim(tbl1(0))
        Dim Notemin As String = Trim(tbl2(0))
        '
        ' Maj des boutons radio de l'onglet Modulation
        ' ********************************************
        RadioModulat.Item(0).Text = NoteM + " " + "Maj"
        RadioModulat.Item(0).Checked = True
        RadioModulat.Item(1).Text = Notemin + " " + "MinH"
        RadioModulat.Item(1).Checked = False
        RadioModulat.Item(2).Text = Notemin + " " + "MinM"
        RadioModulat.Item(2).Checked = False
        '
        ' rendre les modèles invisibles
        'RadioB1.Visible = False
        'RadioB2.Visible = False
        'RadioB3.Visible = False
    End Sub

    Private Sub Maj_Modulation()
        Dim ind As Integer = Det_TonInitial()
        Dim typ As Integer = ComboBox9.SelectedIndex

        EffacerGrid5()
        '
        ' Maj de la tonalité principale (sur ligne 0)
        ' *******************************************
        Select Case ind
            Case 0 ' Maj
                For i = 0 To 6
                    Grid5.Cell(1, i + 1).Text = TableGlobalAcc(typ, 0, i) 'tbl1(i) ' TradEventHLat(TableGlobalAcc(Typ, 0, i)) 'tbl1(i)
                Next i
            Case 1 ' MinH
                For i = 0 To 6
                    Grid5.Cell(1, i + 1).Text = TableGlobalAcc(typ, 1, i) 'tbl1(i) ' TradEventHLat(TableGlobalAcc(Typ, 0, i)) 'tbl1(i)
                Next i
            Case 2 ' MinM
                For i = 0 To 6
                    Grid5.Cell(1, i + 1).Text = TableGlobalAcc(typ, 2, i) 'tbl1(i) ' TradEventHLat(TableGlobalAcc(Typ, 0, i)) 'tbl1(i)
                Next i

        End Select
    End Sub
    Sub INIT_Pistes()
        Dim i, ii As Integer

        ' Transf1 : 0, 1,2,3,4,5
        ' Transf2 : 6,7,8,9,10,11
        ' Transf3 : 12,13,14,15,16,17
        ' Transf4 : 18,19,20,21,22,23
        ' Transf5 : 24,25,26,27,28,29
        ' Transf6 : 30,31,32,33,34,35
        ' Transf7 : 36,37,38,39,40,41

        For i = 0 To nb_BlocPistes - 1
            '  Mutes
            '  *****
            Select Case i
                Case 4, 5, 10, 11, 16, 17, 22, 23, 28, 29, 34, 35, 40, 41
                    SelBloc.Item(i).Checked = False
                Case Else
                    SelBloc.Item(i).Checked = True
            End Select
            '

            ' Motifs
            ' ******
            Select Case i
                Case 0, 1, 2, 3, 4, 5, 18, 19, 20, 21, 22, 23, 36, 37, 38, 39, 40, 41
                    PisteMotif.Item(i).SelectedIndex = 0
                    BoutMotif.Item(i).Text = Trim(PisteMotif.Item(i).Text)
                Case 6, 7, 8, 9, 10, 11, 24, 25, 26, 27, 28, 29
                    PisteMotif.Item(i).SelectedIndex = 1
                    BoutMotif.Item(i).Text = Trim(PisteMotif.Item(i).Text)
                Case 12, 13, 14, 15, 17, 30, 31, 32, 33, 34, 35
                    PisteMotif.Item(i).SelectedIndex = 2
                    BoutMotif.Item(i).Text = Trim(PisteMotif.Item(i).Text)
                Case Else
                    PisteMotif.Item(i).SelectedIndex = 3
                    BoutMotif.Item(i).Text = Trim(PisteMotif.Item(i).Text)
            End Select

            ' Durée
            ' *****
            Select Case i
                Case 0, 1, 2, 3 ' ici on ne traite pas les 2 derniers générateurs de chaque magneto
                    PisteDurée.Item(i).SelectedIndex = 4 - i
                Case 6, 7, 8, 9
                    PisteDurée.Item(i).SelectedIndex = 10 - i
                Case 12, 13, 14, 15
                    PisteDurée.Item(i).SelectedIndex = 16 - i
                Case 18, 19, 20, 21
                    PisteDurée.Item(i).SelectedIndex = 22 - i
                Case 24, 25, 26, 27
                    PisteDurée.Item(i).SelectedIndex = 28 - i
                Case 30, 31, 32, 33
                    PisteDurée.Item(i).SelectedIndex = 34 - i
                Case 36, 37, 38, 39
                    PisteDurée.Item(i).SelectedIndex = 40 - i
                Case Else
                    PisteDurée.Item(i).SelectedIndex = 0
            End Select

            ' Accent
            ' ******
            PisteAccent.Item(i).SelectedIndex = 0

            ' Souche
            ' ******
            PisteSouche.Item(i).SelectedIndex = 0

            ' Retard
            ' ******
            PisteRetard.Item(i).SelectedIndex = 0

            ' Delay
            ' *****
            PisteDelay.Item(i).Checked = False

            ' DébutSouche
            ' ***********
            PisteDébut.Item(i).Checked = False

            ' Octave
            ' ******
            ii = Det_IndexPisteMidi(i)
            Select Case ii
                Case 0
                    PisteOctave.Item(i).SelectedIndex = 0 '+24
                Case 1
                    PisteOctave.Item(i).SelectedIndex = 1 '+12
                Case 2
                    PisteOctave.Item(i).SelectedIndex = 3 ' 0
                Case 3
                    PisteOctave.Item(i).SelectedIndex = 5 ' -12
                Case 4
                    PisteOctave.Item(i).SelectedIndex = 2 ' +7
                Case 5
                    PisteOctave.Item(i).SelectedIndex = 4 ' -7
            End Select

            ' Dyn
            ' ***
            ii = Det_IndexPisteMidi(i)
            Select Case ii
                Case 0, 1, 2, 3
                    PisteDyn.Item(i).SelectedIndex = 4 ' F
                Case 4, 5
                    PisteDyn.Item(i).SelectedIndex = 4 ' MF
            End Select

            PisteDyn.Item(i).SelectedIndex = 32

            ' Octave
            ' ******
            PisteRadio1.Item(i).Checked = False
            PisteRadio2.Item(i).Checked = True
            PisteRadio3.Item(i).Checked = False


            ' PRG
            ' ***
            PistePRG.Item(i).SelectedIndex = 0
            'PistePan.Item(i).Enabled = False
            '
            ' Nom du son
            ' **********
            NomduSon.Item(i).Text = ""
        Next
    End Sub
    Sub INIT_Specific()
        Dim numGrid As Integer
        Dim i, j As Integer


        For numGrid = 0 To TabControl1.TabPages.Count - 1
            With MotifsPerso.Item(numGrid).Range(2, 1, MotifsPerso.Item(numGrid).Rows - 1, MotifsPerso.Item(numGrid).Cols - 1)
                .ClearText()
                .ClearBackColor()
            End With
            For i = 2 To MotifsPerso.Item(numGrid).Rows - 1
                For j = 1 To MotifsPerso.Item(numGrid).Cols - 1
                    If j = 1 Or j = 5 Or j = 9 Or j = 13 Then
                        MotifsPerso.Item(numGrid).Cell(i, j).BackColor = Couleur_Div
                    End If
                Next j
            Next i
        Next

    End Sub
    Sub BLOCS_création()
        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim k As Integer = -1
        Dim h As Integer
        Dim h1 As Integer = 0  ' hauteur des éléments de chaque piste
        Dim h2 As Integer = 21 ' hauteur des labels des éléments de chaque piste
        Dim hl As Integer = 17 ' hauteur des labels
        Dim L1 As Integer = 107 ' position / gauche des élément GM (PRG et PAN)
        Dim CoulMagneto As Color
        Dim ForeCoulMagneto As Color

        Me.TabControl5.Visible = False

        Dim s As Size
        s.Height = 500
        s.Width = 535
        TabControl5.Size = s

        Do
            ' Conteneur Piste et répartition des pistes dans les magnétos
            PistePanel.Add(New Panel)
            '
            If i <= 5 Then
                Me.TabControl5.TabPages.Item(0).Controls.Add(PistePanel.Item(i))
                CoulMagneto = TabCoulMagneto(0) 'Color.LightSteelBlue
                ForeCoulMagneto = Color.DarkRed
            ElseIf i >= 6 And i <= 11 Then
                Me.TabControl5.TabPages.Item(1).Controls.Add(PistePanel.Item(i))
                CoulMagneto = TabCoulMagneto(1)
                ForeCoulMagneto = Color.Black
            ElseIf i >= 12 And i <= 17 Then
                Me.TabControl5.TabPages.Item(2).Controls.Add(PistePanel.Item(i))
                CoulMagneto = TabCoulMagneto(2)
                ForeCoulMagneto = Color.Black
            ElseIf i >= 18 And i <= 23 Then
                Me.TabControl5.TabPages.Item(3).Controls.Add(PistePanel.Item(i))
                CoulMagneto = TabCoulMagneto(3)
                ForeCoulMagneto = Color.Black
            ElseIf i >= 24 And i <= 29 Then
                Me.TabControl5.TabPages.Item(4).Controls.Add(PistePanel.Item(i))
                CoulMagneto = TabCoulMagneto(4)
                ForeCoulMagneto = Color.Black
            ElseIf i >= 30 And i <= 35 Then
                Me.TabControl5.TabPages.Item(5).Controls.Add(PistePanel.Item(i))
                CoulMagneto = TabCoulMagneto(5)
                ForeCoulMagneto = Color.Black
            ElseIf i >= 36 And i <= 41 Then
                Me.TabControl5.TabPages.Item(6).Controls.Add(PistePanel.Item(i))
                CoulMagneto = TabCoulMagneto(6)
                ForeCoulMagneto = Color.Black
            End If
            '
            If k = 5 Then k = -1 ' à l'onglet suivant Magnéto 2
            k = k + 1 ' pour positionnement des pistes des magnétos
            '
            PistePanel.Item(i).Size = New System.Drawing.Size(560, 87)
            h = (((PistePanel.Item(i).Size.Height)) * k) '
            '
            ' Construction des éléments de chaque piste
            ' *****************************************
            PistePanel.Item(i).Location = New Point(1, h)
            PistePanel.Item(i).BorderStyle = BorderStyle.Fixed3D
            PistePanel.Item(i).BackColor = CoulMagneto 'Coul_ELActif 'Color.DarkKhaki
            PistePanel.Item(i).ForeColor = ForeCoulMagneto
            PistePanel.Item(i).Tag = i


            ' Sélection d'une piste
            ' *********************
            PisteMute.Add(New CheckBox)
            PistePanel.Item(i).Controls.Add(PisteMute.Item(i))
            PisteMute.Item(i).Location = New Point(2, 5)
            PisteMute.Item(i).Size = New System.Drawing.Size(98, 20)
            PisteMute.Item(i).BackColor = CoulMagneto     'Coul_ELActif ' Color.FromArgb(207, 201, 118)
            PisteMute.Item(i).ForeColor = Color.DarkRed
            PisteMute.Item(i).Font = Module1.fontMutePiste '
            PisteMute.Item(i).Checked = True
            PisteMute.Item(i).Tag = i
            PisteMute.Item(i).Enabled = True


            ii = Det_IndexPisteMidi(i) + 1
            If LangueIHM = "fr" Then
                PisteMute.Item(i).Text = "PISTE " + Convert.ToString(ii)
            Else
                PisteMute.Item(i).Text = "PISTE " + Convert.ToString(ii)
            End If

            AddHandler PisteMute.Item(i).MouseDown, AddressOf PisteMute_MouseDown
            AddHandler PisteMute.Item(i).CheckedChanged, AddressOf PisteMute_CheckedChange


            ' Sélection d'un BLOC
            ' *******************
            AddHandler PistePanel.Item(i).MouseDown, AddressOf PistePanel_MouseDown
            ' Mute : ici, le MUTE est matérialisé par une non sélection de la case à cocher
            SelBloc.Add(New CheckBox)
            PistePanel.Item(i).Controls.Add(SelBloc.Item(i))
            SelBloc.Item(i).Location = New Point(500, 3)
            SelBloc.Item(i).Size = New System.Drawing.Size(85, 17)
            SelBloc.Item(i).Font = New System.Drawing.Font("Calibri", 9, FontStyle.Bold)
            SelBloc.Item(i).BackColor = CoulMagneto     'Coul_ELActif ' Color.FromArgb(207, 201, 118)
            SelBloc.Item(i).ForeColor = ForeCoulMagneto
            SelBloc.Item(i).Checked = True
            SelBloc.Item(i).Tag = i
            ii = Det_IndexPisteMidi(i) + 1
            If LangueIHM = "fr" Then
                SelBloc.Item(i).Text = Convert.ToString(ii) '  "BLOC " +
            Else
                SelBloc.Item(i).Text = Convert.ToString(ii) '  "BLOCK " +
            End If
            '

            '
            ' Bouton Solo
            ' ***********
            SoloBout.Add(New Button)
            PistePanel.Item(i).Controls.Add(SoloBout.Item(i))
            SoloBout.Item(i).Location = New Point(2, 30)
            SoloBout.Item(i).Font = New Font("Calibri", 8, FontStyle.Regular)
            SoloBout.Item(i).Size = New Size(50, 20)
            SoloBout.Item(i).BackColor = Color.Beige
            SoloBout.Item(i).ForeColor = Color.Black
            SoloBout.Item(i).Text = "SOLO"
            SoloBout.Item(i).Tag = i
            AddHandler SoloBout.Item(i).MouseUp, AddressOf SoloBout_MouseUp

            ' TextBox Nom du Son
            ' ******************
            NomduSon.Add(New TextBox)
            PistePanel.Item(i).Controls.Add(NomduSon.Item(i))
            NomduSon.Item(i).Location = New Point(2, 58)
            NomduSon.Item(i).Font = Module1.fontNomduSon 'New Font("Rubik", 9, FontStyle.Bold)
            NomduSon.Item(i).Size = New Size(100, 20)
            NomduSon.Item(i).BackColor = Color.Beige
            NomduSon.Item(i).ForeColor = Color.Black
            NomduSon.Item(i).ShortcutsEnabled = False
            NomduSon.Item(i).Visible = True
            NomduSon.Item(i).BorderStyle = BorderStyle.FixedSingle
            NomduSon.Item(i).Tag = i
            If LangueIHM = "fr" Then
                ToolTip1.SetToolTip(NomduSon.Item(i), "Ecrivez ici le nom du son utilisé pour le bloc")
            Else
                ToolTip1.SetToolTip(NomduSon.Item(i), "Write here the name of the sound used for the block")
            End If
            AddHandler NomduSon.Item(i).TextChanged, AddressOf NomduSon_TextChanged
            'AddHandler NomduSon.Item(i).KeyUp, AddressOf NomduSon_keyUp

            ' Label AffNom du Son
            ' *******************
            AffNomduSon.Add(New Label)
            PistePanel.Item(i).Controls.Add(AffNomduson.Item(i))
            AffNomduSon.Item(i).Location = New Point(2, 64)
            AffNomduSon.Item(i).TextAlign = ContentAlignment.MiddleLeft
            AffNomduson.Item(i).Font = New Font("Rubik", 9, FontStyle.Bold)
            AffNomduSon.Item(i).Size = New Size(100, 18)
            AffNomduSon.Item(i).BackColor = Color.Beige
            AffNomduSon.Item(i).ForeColor = Color.DarkRed
            AffNomduSon.Item(i).Visible = False
            AffNomduSon.Item(i).BorderStyle = BorderStyle.FixedSingle
            AffNomduSon.Item(i).Tag = i

            AddHandler AffNomduson.Item(i).MouseUp, AddressOf AffNomduson_MouseUp

            ' Bouton formulaire des  motifs
            ' *****************************
            BoutMotif.Add(New BoutPerso)
            PistePanel.Item(i).Controls.Add(BoutMotif.Item(i))
            BoutMotif.Item(i).Location = New Point(274, h1)
            BoutMotif.Item(i).Size = New System.Drawing.Size(95, 23)
            BoutMotif.Item(i).Font = New System.Drawing.Font("Calibri", 9, FontStyle.Regular)
            BoutMotif.Item(i).Text = "ArpMotif1"
            BoutMotif.Item(i).Tag = i

            PisteMotif.Add(New Windows.Forms.ComboBox)
            PistePanel.Item(i).Controls.Add(PisteMotif.Item(i))
            PisteMotif.Item(i).Location = New Point(274, h1)
            PisteMotif.Item(i).Size = New System.Drawing.Size(95, 45)
            PisteMotif.Item(i).Font = New System.Drawing.Font("Calibri", 9, FontStyle.Regular)
            PisteMotif.Item(i).Tag = i
            PisteMotif.Item(i).Visible = False ' on est obligé de garder PisteMotif pour être compatible avec les anciens fichiers

            '
            labelMotif.Add(New Label)
            PistePanel.Item(i).Controls.Add(labelMotif.Item(i))
            labelMotif.Item(i).Location = New Point(274, h2 + 2)
            labelMotif.Item(i).Size = New System.Drawing.Size(75, hl)
            labelMotif.Item(i).Font = New System.Drawing.Font("Calibri", 8, FontStyle.Regular)
            '
            '
            For j = 1 To 3
                PisteMotif.Item(i).Items.Add("ArpMotif" + Str(j))
            Next j
            'For j = 1 To 3
            PisteMotif.Item(i).Items.Add("Répétition V1")
            PisteMotif.Item(i).Items.Add("Répétition V2")
            PisteMotif.Item(i).Items.Add("Répétition V3")


            'Next j
            For j = 1 To 6
                PisteMotif.Item(i).Items.Add("Perso" + Str(j))
            Next j
            ' Chords
            PisteMotif.Item(i).Items.Add("Full Chord")
            PisteMotif.Item(i).Items.Add("No3rd Chord")
            PisteMotif.Item(i).Items.Add("No5th Chord")
            PisteMotif.Item(i).Items.Add("Motif1 Chord")
            'PisteMotif.Item(i).Items.Add("Répétition T")
            'PisteMotif.Item(i).Items.Add("Répétition 3")


            labelMotif.Item(i).Text = "Motifs"
            PisteMotif.Item(i).SelectedIndex = 0

            PisteMotif.Item(i).Visible = True
            labelMotif.Item(i).Visible = True
            '
            BoutMotif.Item(i).Visible = False
            '
            AddHandler PisteMotif.Item(i).SelectedIndexChanged, AddressOf PisteMotif_SelectedIndexChanged
            AddHandler PisteMotif.Item(i).KeyDown, AddressOf PisteMotif_Keydown
            AddHandler BoutMotif.Item(i).MouseClick, AddressOf BoutMotif_MouseClick
            '
            ' Liste Durée note
            PisteDurée.Add(New Windows.Forms.ComboBox)
            PistePanel.Item(i).Controls.Add(PisteDurée.Item(i))
            PisteDurée.Item(i).Location = New Point(394, h1)
            PisteDurée.Item(i).Size = New System.Drawing.Size(40, 45)
            PisteDurée.Item(i).Font = New System.Drawing.Font("Calibri", 8, FontStyle.Regular)
            PisteDurée.Item(i).Tag = i
            '
            labelDurée.Add(New Label)
            PistePanel.Item(i).Controls.Add(labelDurée.Item(i))
            labelDurée.Item(i).Location = New Point(394, h2)
            labelDurée.Item(i).Size = New System.Drawing.Size(40, hl)
            labelDurée.Item(i).Font = New System.Drawing.Font("Calibri", 8, FontStyle.Regular)
            '
            If LangueIHM = "fr" Then
                PisteDurée.Item(i).Items.Add("RN")
                PisteDurée.Item(i).Items.Add("BL")
                PisteDurée.Item(i).Items.Add("NR")
                PisteDurée.Item(i).Items.Add("CR")
                PisteDurée.Item(i).Items.Add("DC")
                labelDurée.Item(i).Text = "Durée"
                ToolTip1.SetToolTip(PisteDurée.Item(i), "RN: ronde BL: blanche NR: noire CR: croche DC: Double croche")
            Else
                PisteDurée.Item(i).Items.Add("WN")
                PisteDurée.Item(i).Items.Add("HN")
                PisteDurée.Item(i).Items.Add("QN")
                PisteDurée.Item(i).Items.Add("EN")
                PisteDurée.Item(i).Items.Add("SN")
                labelDurée.Item(i).Text = "Length"
                ToolTip1.SetToolTip(PisteDurée.Item(i), "WH: whole HN: half QN: quarter EN: eighth  SN: sixteen")
            End If
            PisteDurée.Item(i).Visible = True
            labelDurée.Item(i).Visible = True


            ' 
            AddHandler PisteDurée.Item(i).SelectedIndexChanged, AddressOf PisteDurée_SelectedIndexChanged
            AddHandler PisteDurée.Item(i).KeyDown, AddressOf PisteDurée_Keydown
            ' Liste Accent
            PisteAccent.Add(New Windows.Forms.ComboBox)
            PistePanel.Item(i).Controls.Add(PisteAccent.Item(i))
            PisteAccent.Item(i).Location = New Point(330, h1 + 40) 'New Point(394, h1 + 10) ' 394
            PisteAccent.Item(i).Size = New System.Drawing.Size(40, 45)
            PisteAccent.Item(i).Font = New System.Drawing.Font("Calibri", 8, FontStyle.Regular)
            PisteAccent.Item(i).Tag = i
            '
            labelAccent.Add(New Label)
            PistePanel.Item(i).Controls.Add(labelAccent.Item(i))
            labelAccent.Item(i).Location = New Point(330, h1 + 61)
            labelAccent.Item(i).Size = New System.Drawing.Size(45, hl)
            labelAccent.Item(i).Font = New System.Drawing.Font("Calibri", 8, FontStyle.Regular)
            If LangueIHM = "fr" Then
                labelAccent.Item(i).Text = "Accent"
            Else
                labelAccent.Item(i).Text = "Accent"
            End If
            '
            PisteAccent.Item(i).Items.Add("off")
            PisteAccent.Item(i).Items.Add("+5")
            PisteAccent.Item(i).Items.Add("+10")
            PisteAccent.Item(i).Items.Add("+15")
            PisteAccent.Item(i).Items.Add("+20")
            PisteAccent.Item(i).Items.Add("+25")

            PisteAccent.Item(i).Visible = True
            labelAccent.Item(i).Visible = True
            '
            AddHandler PisteAccent.Item(i).SelectedIndexChanged, AddressOf PisteAccent__SelectedIndexChanged
            AddHandler PisteAccent.Item(i).KeyDown, AddressOf PisteAccent_Keydown

            '
            ' Liste Souche (souche = note de début de lecture de l'accord)
            PisteSouche.Add(New Windows.Forms.ComboBox)
            PistePanel.Item(i).Controls.Add(PisteSouche.Item(i))
            PisteSouche.Item(i).Location = New Point(450, h1 + 40) 'New Point(394, h1 + 10) ' 394
            PisteSouche.Item(i).Size = New System.Drawing.Size(40, 45)
            PisteSouche.Item(i).Font = New System.Drawing.Font("Calibri", 8, FontStyle.Regular)
            PisteSouche.Item(i).Tag = i
            '
            labelSouche.Add(New Label)
            PistePanel.Item(i).Controls.Add(labelSouche.Item(i))
            labelSouche.Item(i).Location = New Point(450, h1 + 61)
            labelSouche.Item(i).Size = New System.Drawing.Size(45, hl)
            labelSouche.Item(i).Font = New System.Drawing.Font("Calibri", 8, FontStyle.Regular)

            If LangueIHM = "fr" Then
                labelSouche.Item(i).Text = "Départ"
            Else
                labelSouche.Item(i).Text = "Start"
            End If
            '
            PisteSouche.Item(i).Items.Add("1")
            PisteSouche.Item(i).Items.Add("2")
            PisteSouche.Item(i).Items.Add("3")

            PisteSouche.Item(i).Visible = True
            labelSouche.Item(i).Visible = True
            If i > 5 Then
                PisteSouche.Item(i).Visible = False
                labelSouche.Item(i).Visible = False
            End If
            '
            AddHandler PisteSouche.Item(i).SelectedIndexChanged, AddressOf PisteSouche__SelectedIndexChanged
            AddHandler PisteSouche.Item(i).KeyDown, AddressOf PisteSouche_Keydown
            '
            ' liste retard
            PisteRetard.Add(New Windows.Forms.ComboBox)
            PistePanel.Item(i).Controls.Add(PisteRetard.Item(i))
            PisteRetard.Item(i).Location = New Point(235, h1 + 37) 'New Point(394, h1 + 10)
            PisteRetard.Item(i).Size = New System.Drawing.Size(40, 45)
            PisteRetard.Item(i).Font = New System.Drawing.Font("Calibri", 8, FontStyle.Regular)
            PisteRetard.Item(i).Tag = i
            '
            labelRetard.Add(New Label)
            PistePanel.Item(i).Controls.Add(labelRetard.Item(i))
            labelRetard.Item(i).Location = New Point(235, h1 + 56)
            labelRetard.Item(i).Size = New System.Drawing.Size(45, hl)
            labelRetard.Item(i).Font = New System.Drawing.Font("Calibri", 8, FontStyle.Regular)
            If LangueIHM = "fr" Then
                labelRetard.Item(i).Text = "Retard"
            Else
                labelRetard.Item(i).Text = "Delay"
            End If
            '
            PisteRetard.Item(i).Items.Add("off")
            PisteRetard.Item(i).Items.Add("1")
            PisteRetard.Item(i).Items.Add("2")
            PisteRetard.Item(i).Items.Add("3")

            PisteRetard.Item(i).Visible = False
            labelRetard.Item(i).Visible = False

            ' Delay
            PisteDelay.Add(New CheckBox)
            PistePanel.Item(i).Controls.Add(PisteDelay.Item(i))
            PisteDelay.Item(i).Font = New System.Drawing.Font("Calibri", 8, FontStyle.Regular)
            PisteDelay.Item(i).Location = New Point(395, h1 + 39)
            PisteDelay.Item(i).Size = New System.Drawing.Size(60, 20)
            PisteDelay.Item(i).Tag = i
            PisteDelay.Item(i).Text = "Delay"
            PisteDelay.Item(i).Visible = True
            If i > 5 Then
                PisteDelay.Item(i).Visible = False
            End If

            '
            AddHandler PisteDelay.Item(i).CheckedChanged, AddressOf PisteDelay_CheckedChanged
            AddHandler PisteDelay.Item(i).MouseDown, AddressOf PisteDelay_MouseDown
            AddHandler PisteDelay.Item(i).MouseUp, AddressOf PisteDelay_MouseUp


            ' Début (Souche) 
            PisteDébut.Add(New CheckBox)
            PistePanel.Item(i).Controls.Add(PisteDébut.Item(i))
            PisteDébut.Item(i).Font = New System.Drawing.Font("Calibri", 8, FontStyle.Regular)
            PisteDébut.Item(i).Location = New Point(394, h1 + 39)
            PisteDébut.Item(i).Size = New System.Drawing.Size(60, 20)
            PisteDébut.Item(i).Tag = i
            PisteDébut.Item(i).Text = "Début"
            PisteDébut.Item(i).Visible = False
            PisteDébut.Item(i).Visible = False
            '
            AddHandler PisteDébut.Item(i).CheckedChanged, AddressOf PisteDébut_CheckedChanged
            ' Liste Octave
            PisteOctave.Add(New Windows.Forms.ComboBox)
            PistePanel.Item(i).Controls.Add(PisteOctave.Item(i))
            '
            PisteOctave.Item(i).Location = New Point(450, h1)
            PisteOctave.Item(i).Size = New System.Drawing.Size(40, 45)
            PisteOctave.Item(i).Font = New System.Drawing.Font("Calibri", 8, FontStyle.Regular)
            PisteOctave.Item(i).Tag = i
            '
            labelOctave.Add(New Label)
            PistePanel.Item(i).Controls.Add(labelOctave.Item(i))
            labelOctave.Item(i).Location = New Point(450, h2)
            labelOctave.Item(i).Size = New System.Drawing.Size(40, hl)
            labelOctave.Item(i).Font = New System.Drawing.Font("Calibri", 8, FontStyle.Regular)
            labelOctave.Item(i).Text = "Transp"
            '
            PisteOctave.Item(i).Items.Add("+24")
            PisteOctave.Item(i).Items.Add("+12")
            PisteOctave.Item(i).Items.Add("+7")

            PisteOctave.Item(i).Items.Add("0")

            PisteOctave.Item(i).Items.Add("-7")
            PisteOctave.Item(i).Items.Add("-12")
            PisteOctave.Item(i).Items.Add("-24")
            PisteOctave.Item(i).SelectedIndex = 2
            '
            PisteOctave.Item(i).Visible = True
            labelOctave.Item(i).Visible = True
            '
            AddHandler PisteOctave.Item(i).SelectedIndexChanged, AddressOf PisteOctave_SelectedIndexChanged
            AddHandler PisteOctave.Item(i).KeyDown, AddressOf PisteOctave_Keydown

            ' liste dyn
            PisteDyn.Add(New Windows.Forms.ComboBox)
            PistePanel.Item(i).Controls.Add(PisteDyn.Item(i))
            PisteDyn.Item(i).Location = New Point(275, h1 + 41)
            PisteDyn.Item(i).Size = New System.Drawing.Size(43, 45)
            PisteDyn.Item(i).Font = New System.Drawing.Font("Calibri", 8, FontStyle.Regular)
            PisteDyn.Item(i).Tag = i
            '
            labelDyn.Add(New Label)
            PistePanel.Item(i).Controls.Add(labelDyn.Item(i))
            labelDyn.Item(i).Location = New Point(275, h1 + 61)
            labelDyn.Item(i).Size = New System.Drawing.Size(40, hl)
            labelDyn.Item(i).Font = New System.Drawing.Font("Calibri", 8, FontStyle.Regular)
            labelDyn.Item(i).Text = "Dyn"

            'PisteDyn.Item(i).Items.Add("16")
            'PisteDyn.Item(i).Items.Add("15")
            'PisteDyn.Item(i).Items.Add("14")
            'PisteDyn.Item(i).Items.Add("13")
            'PisteDyn.Item(i).Items.Add("12")
            'PisteDyn.Item(i).Items.Add("11")
            'PisteDyn.Item(i).Items.Add("10")
            'PisteDyn.Item(i).Items.Add("9")
            'PisteDyn.Item(i).Items.Add("8")
            'PisteDyn.Item(i).Items.Add("7")
            'PisteDyn.Item(i).Items.Add("6")
            'PisteDyn.Item(i).Items.Add("5")
            'PisteDyn.Item(i).Items.Add("4")
            'PisteDyn.Item(i).Items.Add("3")
            'PisteDyn.Item(i).Items.Add("2")
            'PisteDyn.Item(i).Items.Add("1")
            '
            For j = 127 To 0 Step -1
                PisteDyn.Item(i).Items.Add(Convert.ToString(j))
            Next
            '
            PisteDyn.Item(i).SelectedIndex = 32
            '
            PisteDyn.Item(i).Visible = True
            labelDyn.Item(i).Visible = True
            '
            AddHandler PisteDyn.Item(i).SelectedIndexChanged, AddressOf PisteDyn_SelectedIndexChanged
            AddHandler PisteDyn.Item(i).KeyDown, AddressOf PisteDyn_Keydown

            ' PisteDyn2 numérique Up&Down (le label n'esst pas traité, il est traité avec PisteDyn)
            ' ************************************************************************************
            PisteDyn2.Add(New Windows.Forms.NumericUpDown)
            PistePanel.Item(i).Controls.Add(PisteDyn2.Item(i))
            PisteDyn2.Item(i).Location = New Point(275, h1 + 41)
            PisteDyn2.Item(i).Size = New System.Drawing.Size(40, 20)
            PisteDyn.Item(i).Font = New System.Drawing.Font("Calibri", 8, FontStyle.Regular)
            PisteDyn2.Item(i).Minimum = 0
            PisteDyn2.Item(i).Maximum = 127
            PisteDyn2.Item(i).Increment = 1
            PisteDyn2.Item(i).Value = 90
            PisteDyn2.Item(i).Tag = i
            PisteDyn2.Item(i).Visible = False

            AddHandler PisteDyn2.Item(i).KeyDown, AddressOf PisteDyn2_Keydown
            '
            ' PAN : Boutons radio
            PistePan.Add(New Windows.Forms.GroupBox)
            PisteRadio1.Add(New RadioButton)
            PisteRadio2.Add(New RadioButton)
            PisteRadio3.Add(New RadioButton)
            PistePanel.Item(i).Controls.Add(PistePan.Item(i))
            '
            PistePan.Item(i).Controls.Add(PisteRadio1.Item(i))
            If LangueIHM = "fr" Then
                PisteRadio1.Item(i).Text = "G"
                ToolTip1.SetToolTip(PisteRadio1.Item(i), "Panoramique")
            Else
                PisteRadio1.Item(i).Text = "L"
                ToolTip1.SetToolTip(PisteRadio1.Item(i), "Panoramic")
            End If
            PisteRadio1.Item(i).Size = New System.Drawing.Size(38, 20)
            PisteRadio1.Item(i).Location = New Point(12, 10)
            PisteRadio1.Item(i).Tag = i
            'PisteRadio1.Item(i).Checked = False
            '
            PistePan.Item(i).Controls.Add(PisteRadio2.Item(i))
            If LangueIHM = "fr" Then
                PisteRadio2.Item(i).Text = "C"
                ToolTip1.SetToolTip(PisteRadio2.Item(i), "Panoramique")
            Else
                PisteRadio2.Item(i).Text = "C"
                ToolTip1.SetToolTip(PisteRadio2.Item(i), "Panoramic")
            End If
            PisteRadio2.Item(i).Location = New Point(60, 10)
            PisteRadio2.Item(i).Size = New System.Drawing.Size(38, 20)
            PisteRadio2.Item(i).Tag = i

            PisteRadio2.Item(i).Visible = True
            'PisteRadio2.Item(i).Checked = True
            '
            PistePan.Item(i).Controls.Add(PisteRadio3.Item(i))
            If LangueIHM = "fr" Then
                PisteRadio3.Item(i).Text = "D"
                ToolTip1.SetToolTip(PisteRadio3.Item(i), "Panoramique")
            Else
                PisteRadio3.Item(i).Text = "R"
                ToolTip1.SetToolTip(PisteRadio3.Item(i), "Panoramic")
            End If
            PisteRadio3.Item(i).Location = New Point(100, 10)
            PisteRadio3.Item(i).Size = New System.Drawing.Size(38, 20)
            PisteRadio3.Item(i).Tag = i
            'PisteRadio3.Item(i).Checked = False
            '
            PistePan.Item(i).Location = New Point(L1, 32) ' 173
            PistePan.Item(i).Size = New System.Drawing.Size(139, 32)
            PistePan.Item(i).Font = New System.Drawing.Font("Calibri", 8, FontStyle.Regular)
            'PisteEntrée.Item(i).Text = "Entrée"
            PistePan.Item(i).BackColor = CoulMagneto
            PistePan.Item(i).ForeColor = ForeCoulMagneto
            PistePan.Item(i).Tag = i
            PistePan.Item(i).Visible = True
            PistePan.Item(i).Enabled = True

            AddHandler PisteRadio1.Item(i).MouseDown, AddressOf PisteRadio1_MouseDown
            AddHandler PisteRadio2.Item(i).MouseDown, AddressOf PisteRadio2_MouseDown
            AddHandler PisteRadio3.Item(i).MouseDown, AddressOf PisteRadio3_MouseDown

            ' liste programmes
            Maj_PRG(i, h1, h2, hl, L1, CoulMagneto, ForeCoulMagneto)

            'Label N° de Pistes
            labelNPiste.Add(New Label)
            PistePanel.Item(i).Controls.Add(labelNPiste.Item(i))
            labelNPiste.Item(i).Location = New Point(25, 5)
            labelNPiste.Item(i).Size = New System.Drawing.Size(100, hl + 20)
            labelNPiste.Item(i).Font = New System.Drawing.Font("Calibri", 18, FontStyle.Regular)
            labelNPiste.Item(i).Visible = False
            ii = Det_IndexPisteMidi(i) + 1
            If LangueIHM = "fr" Then
                labelNPiste.Item(i).Text = "Piste " + Convert.ToString(ii)
            Else
                labelNPiste.Item(i).Text = "Track " + Convert.ToString(ii)
            End If
            i = i + 1
        Loop Until i = nb_BlocPistes 'Arrangement1.Nb_Pistes
        '
        '  Ajout des pistes dans l'objet "LesPistes" (défini dans MainArp)
        '  *************************************************************
        For i = 0 To nb_TotalCurseurs - 1 'nb_TotalPistes - 1 '(Arrangement1.Nb_PistesMidi) ' la dernière piste N°6  est la piste mélodique (pianoroll1)
            LesPistes.Add(New Piste("Piste" + Trim(Str(i + 1)), i, i)) ' NomPiste, N° Piste, Canal
        Next i
        '
        ' Maj de toutes les valeurs par défaut
        ' ************************************
        INIT_Pistes() ' 
        '

        Me.TabControl5.Visible = True
        Me.TabControl5.Refresh()
    End Sub
    '***********************************************************
    '* MAJ_PRG : 
    '***********************************************************
    Sub Maj_PRG(i As Integer, h1 As Integer, h2 As Integer, hl As Integer, L1 As Integer, Couleur As Color, Forecouleur As Color)

        PistePRG.Add(New Windows.Forms.ComboBox)
        PistePanel.Item(i).Controls.Add(PistePRG.Item(i))
        PistePRG.Item(i).Location = New Point(L1, h1)
        PistePRG.Item(i).BackColor = Couleur
        PistePRG.Item(i).ForeColor = Forecouleur
        PistePRG.Item(i).Size = New System.Drawing.Size(140, 45)
        PistePRG.Item(i).Font = New System.Drawing.Font("Calibri", 8, FontStyle.Regular)
        PistePRG.Item(i).Tag = i
        '
        labelPRG.Add(New Label)
        PistePanel.Item(i).Controls.Add(labelPRG.Item(i))
        labelPRG.Item(i).Location = New Point(L1, h2)
        labelPRG.Item(i).Size = New System.Drawing.Size(70, hl)
        labelPRG.Item(i).Font = New System.Drawing.Font("Calibri", 8, FontStyle.Regular)
        '
        labelPRG.Item(i).Visible = True
        PistePRG.Item(i).Visible = True
        '
        If LangueIHM = "fr" Then
            labelPRG.Item(i).Text = "Programmes"
            PistePRG.Item(i).Items.Add("GS/GM off")
            PistePRG.Item(i).Items.Add("01 Piano acoustique 1")
            PistePRG.Item(i).Items.Add("02 Piano acoustique 2")
            PistePRG.Item(i).Items.Add("03 Grand piano électrique")
            PistePRG.Item(i).Items.Add("04 Piano honkytonk")
            PistePRG.Item(i).Items.Add("05 Piano électrique 1")
            PistePRG.Item(i).Items.Add("06 Piano électrique 2")
            PistePRG.Item(i).Items.Add("07 Clavecin")
            PistePRG.Item(i).Items.Add("08 Clavicorde")
            PistePRG.Item(i).Items.Add("09 Célesta")
            PistePRG.Item(i).Items.Add("10 Carillon")
            PistePRG.Item(i).Items.Add("11 Boîte à musique")
            PistePRG.Item(i).Items.Add("12 Vibraphone")
            PistePRG.Item(i).Items.Add("13 Marimba")
            PistePRG.Item(i).Items.Add("14 Xylophone")
            PistePRG.Item(i).Items.Add("15 Cloches tubulaires")
            PistePRG.Item(i).Items.Add("16 Tympanon")
            PistePRG.Item(i).Items.Add("17 Orgue à tubes")
            PistePRG.Item(i).Items.Add("18 Orgue percussif")
            PistePRG.Item(i).Items.Add("19 Orgue rock")
            PistePRG.Item(i).Items.Add("20 Orgue d'église")
            PistePRG.Item(i).Items.Add("21 Orgue vibrato")
            PistePRG.Item(i).Items.Add("22 Accordéon")
            PistePRG.Item(i).Items.Add("23 Harmonica")
            PistePRG.Item(i).Items.Add("24 Bandonéon")
            PistePRG.Item(i).Items.Add("25 Guitare classique")
            PistePRG.Item(i).Items.Add("26 Guitare folk")
            PistePRG.Item(i).Items.Add("27 Guitare jazz")
            PistePRG.Item(i).Items.Add("28 Guitare élec. pure")
            PistePRG.Item(i).Items.Add("29 Guitare élec. mute")
            PistePRG.Item(i).Items.Add("30 Guitare saturée")
            PistePRG.Item(i).Items.Add("31 Guitare distorsion")
            PistePRG.Item(i).Items.Add("32 Guitare harmonique")
            PistePRG.Item(i).Items.Add("33 Basse acoustique")
            PistePRG.Item(i).Items.Add("34 Basse élec. 1")
            PistePRG.Item(i).Items.Add("35 Basse élec. 2")
            PistePRG.Item(i).Items.Add("36 Basse élec. 3")
            PistePRG.Item(i).Items.Add("37 Basse slap 1")
            PistePRG.Item(i).Items.Add("38 Basse slap 2")
            PistePRG.Item(i).Items.Add("39 Basse synth.  1")
            PistePRG.Item(i).Items.Add("40 Basse synth.  2")
            PistePRG.Item(i).Items.Add("41 Violon")
            PistePRG.Item(i).Items.Add("42 Viole")
            PistePRG.Item(i).Items.Add("43 Violoncelle")
            PistePRG.Item(i).Items.Add("44 Contrebasse")
            PistePRG.Item(i).Items.Add("45 Cordes trémolo")
            PistePRG.Item(i).Items.Add("46 Cordes pizzicato")
            PistePRG.Item(i).Items.Add("47 Harpe")
            PistePRG.Item(i).Items.Add("48 Timbales")
            PistePRG.Item(i).Items.Add("49 Quartet cordes 1")
            PistePRG.Item(i).Items.Add("50 Quartet cordes 2")
            PistePRG.Item(i).Items.Add("51 Cordes synth 1")
            PistePRG.Item(i).Items.Add("52 Cordes synth 2")
            PistePRG.Item(i).Items.Add("53 Chœurs Aahs")
            PistePRG.Item(i).Items.Add("54 Voix Oohs")
            PistePRG.Item(i).Items.Add("55 Voix synthétiseur")
            PistePRG.Item(i).Items.Add("56 Coup d'orchestre")
            PistePRG.Item(i).Items.Add("57 Trompette")
            PistePRG.Item(i).Items.Add("58 Trombone")
            PistePRG.Item(i).Items.Add("59 Tuba")
            PistePRG.Item(i).Items.Add("60 Trompette bouchée")
            PistePRG.Item(i).Items.Add("61 Cors")
            PistePRG.Item(i).Items.Add("62 Ensemble de cuivres")
            PistePRG.Item(i).Items.Add("63 Cuivres synthétiseur")
            PistePRG.Item(i).Items.Add("64 Cuivres synthétiseur")
            PistePRG.Item(i).Items.Add("65 Saxophone soprano")
            PistePRG.Item(i).Items.Add("66 Saxophone alto")
            PistePRG.Item(i).Items.Add("67 Saxophone ténor")
            PistePRG.Item(i).Items.Add("68 Saxophone baryton")
            PistePRG.Item(i).Items.Add("69 Hautbois")
            PistePRG.Item(i).Items.Add("70 Cors anglais")
            PistePRG.Item(i).Items.Add("71 Basson")
            PistePRG.Item(i).Items.Add("72 Clarinette")
            PistePRG.Item(i).Items.Add("73 Flûte piccolo")
            PistePRG.Item(i).Items.Add("74 Flûte")
            PistePRG.Item(i).Items.Add("75 Flûte à bec")
            PistePRG.Item(i).Items.Add("76 Flûte de pan")
            PistePRG.Item(i).Items.Add("77 Bouteille sifflée")
            PistePRG.Item(i).Items.Add("78 Shakuhachi")
            PistePRG.Item(i).Items.Add("79 Sifflet")
            PistePRG.Item(i).Items.Add("80 Ocarina")
            PistePRG.Item(i).Items.Add("81 Lead carré")
            PistePRG.Item(i).Items.Add("82 Lead dents de scie")
            PistePRG.Item(i).Items.Add("83 Lead orgue")
            PistePRG.Item(i).Items.Add("84 Lead chiff")
            PistePRG.Item(i).Items.Add("85 Lead charang")
            PistePRG.Item(i).Items.Add("86 Lead voix")
            PistePRG.Item(i).Items.Add("87 Lead quinte)")
            PistePRG.Item(i).Items.Add("88 Lead basse")
            PistePRG.Item(i).Items.Add("89 Pad new Age")
            PistePRG.Item(i).Items.Add("90 Pad warm")
            PistePRG.Item(i).Items.Add("91 Pad poly")
            PistePRG.Item(i).Items.Add("92 Pad chœurs")
            PistePRG.Item(i).Items.Add("93 Pad archet")
            PistePRG.Item(i).Items.Add("94 Pad métal")
            PistePRG.Item(i).Items.Add("95 Pad halo")
            PistePRG.Item(i).Items.Add("96 Pad glissement")
        Else
            labelPRG.Item(i).Text = "Programs"
            PistePRG.Item(i).Items.Add("GS/GM off")
            PistePRG.Item(i).Items.Add("01 Acoustic Grand Piano")
            PistePRG.Item(i).Items.Add("02 Bright Acoustic Piano")
            PistePRG.Item(i).Items.Add("03 Electric Grand Piano")
            PistePRG.Item(i).Items.Add("04 Honky-tonk Piano")
            PistePRG.Item(i).Items.Add("05 Electric Piano 1")
            PistePRG.Item(i).Items.Add("06 Electric Piano 2")
            PistePRG.Item(i).Items.Add("07 07 Harpsichord")
            PistePRG.Item(i).Items.Add("08 Clavinet")
            PistePRG.Item(i).Items.Add("09 Celesta")
            PistePRG.Item(i).Items.Add("10 Glockenspiel")
            PistePRG.Item(i).Items.Add("11 Music Box")
            PistePRG.Item(i).Items.Add("12 Vibraphone")
            PistePRG.Item(i).Items.Add("13 Marimba")
            PistePRG.Item(i).Items.Add("14 Xylophone")
            PistePRG.Item(i).Items.Add("15 Tubular Bells")
            PistePRG.Item(i).Items.Add("16 Dulcimer")
            PistePRG.Item(i).Items.Add("17 Drawbar Organ")
            PistePRG.Item(i).Items.Add("18 Percussive Organ")
            PistePRG.Item(i).Items.Add("19 Rock Organ")
            PistePRG.Item(i).Items.Add("20 Church Organ")
            PistePRG.Item(i).Items.Add("21 Reed Organ")
            PistePRG.Item(i).Items.Add("22 Accordion")
            PistePRG.Item(i).Items.Add("23 Harmonica")
            PistePRG.Item(i).Items.Add("24 Tango Accordion")
            PistePRG.Item(i).Items.Add("25 Acoustic nylon")
            PistePRG.Item(i).Items.Add("26 Acoustic steel")
            PistePRG.Item(i).Items.Add("27 Electric jazz")
            PistePRG.Item(i).Items.Add("28 Electric clean")
            PistePRG.Item(i).Items.Add("29 Electric muted")
            PistePRG.Item(i).Items.Add("30 Overdriven Guitar")
            PistePRG.Item(i).Items.Add("31 Distortion Guitar")
            PistePRG.Item(i).Items.Add("32 Guitar Harmonics")
            PistePRG.Item(i).Items.Add("33 Acoustic Bass")
            PistePRG.Item(i).Items.Add("34 Electric Bass finger")
            PistePRG.Item(i).Items.Add("35 Electric Bass pick")
            PistePRG.Item(i).Items.Add("36 Fretless Bass")
            PistePRG.Item(i).Items.Add("37 Slap Bass 1")
            PistePRG.Item(i).Items.Add("38 Slap Bass 2")
            PistePRG.Item(i).Items.Add("39 Synth Bass 1")
            PistePRG.Item(i).Items.Add("40 Synth Bass 2")
            PistePRG.Item(i).Items.Add("41 Violon")
            PistePRG.Item(i).Items.Add("42 Viola")
            PistePRG.Item(i).Items.Add("43 Cello")
            PistePRG.Item(i).Items.Add("44 Contrabass")
            PistePRG.Item(i).Items.Add("45 Tremolo Strings")
            PistePRG.Item(i).Items.Add("46 Pizzicato Strings")
            PistePRG.Item(i).Items.Add("47 Orchestral Harp")
            PistePRG.Item(i).Items.Add("48 Timpani")
            PistePRG.Item(i).Items.Add("49 String Ensemble 1")
            PistePRG.Item(i).Items.Add("50 String Ensemble 2")
            PistePRG.Item(i).Items.Add("51 Synth Strings 1")
            PistePRG.Item(i).Items.Add("52 Synth Strings 2")
            PistePRG.Item(i).Items.Add("53 Choir Aahs")
            PistePRG.Item(i).Items.Add("54 Voice Oohs")
            PistePRG.Item(i).Items.Add("55 Synth Choir")
            PistePRG.Item(i).Items.Add("56 Orchestra Hit")
            PistePRG.Item(i).Items.Add("57 Trumpet")
            PistePRG.Item(i).Items.Add("58 Trombone")
            PistePRG.Item(i).Items.Add("59 Tuba")
            PistePRG.Item(i).Items.Add("60 Muted Trumpet")
            PistePRG.Item(i).Items.Add("61 French Horn")
            PistePRG.Item(i).Items.Add("62 Brass Section")
            PistePRG.Item(i).Items.Add("63 Synth Brass 1")
            PistePRG.Item(i).Items.Add("64 Synth Brass 2")
            PistePRG.Item(i).Items.Add("65 Soprano Sax")
            PistePRG.Item(i).Items.Add("66 Alto Sax")
            PistePRG.Item(i).Items.Add("67 Tenor Sax")
            PistePRG.Item(i).Items.Add("68 Baritone Sax")
            PistePRG.Item(i).Items.Add("69 Oboe")
            PistePRG.Item(i).Items.Add("70 English Horn")
            PistePRG.Item(i).Items.Add("71 Bassoon")
            PistePRG.Item(i).Items.Add("72 Clarinet")
            PistePRG.Item(i).Items.Add("73 Piccolo")
            PistePRG.Item(i).Items.Add("74 Flute")
            PistePRG.Item(i).Items.Add("75 Recorder")
            PistePRG.Item(i).Items.Add("76 Pan Flute")
            PistePRG.Item(i).Items.Add("77 Blown bottle")
            PistePRG.Item(i).Items.Add("78 Shakuhachi")
            PistePRG.Item(i).Items.Add("79 Whistle")
            PistePRG.Item(i).Items.Add("80 Ocarina")
            PistePRG.Item(i).Items.Add("81 Lead 1 square")
            PistePRG.Item(i).Items.Add("82 Lead 2 sawtooth")
            PistePRG.Item(i).Items.Add("83 Lead 3 calliope")
            PistePRG.Item(i).Items.Add("84 Lead chiff")
            PistePRG.Item(i).Items.Add("85 Lead charang")
            PistePRG.Item(i).Items.Add("86 Lead voice")
            PistePRG.Item(i).Items.Add("87 Lead 7 fifths")
            PistePRG.Item(i).Items.Add("88 Lead 8 bass + lead")
            PistePRG.Item(i).Items.Add("89 Pad new Age")
            PistePRG.Item(i).Items.Add("90 Pad warm")
            PistePRG.Item(i).Items.Add("91 Pad 3 polysynth")
            PistePRG.Item(i).Items.Add("92 Pad 4 choir")
            PistePRG.Item(i).Items.Add("93 Pad 5 bowed")
            PistePRG.Item(i).Items.Add("94 Pad 6 metallic")
            PistePRG.Item(i).Items.Add("95 Pad 7 halo")
            PistePRG.Item(i).Items.Add("96 Pad 8 sweep")
        End If   '
        PistePRG.Item(i).SelectedIndex = 0

        '
        AddHandler PistePRG.Item(i).SelectedIndexChanged, AddressOf PistePRG_SelectedIndexChanged
        AddHandler PistePRG.Item(i).KeyDown, AddressOf PistePRG_Keydown
    End Sub
    ' *************************
    ' *  Evénement des pistes *
    ' *************************
    Private Sub PisteDyn2_Keydown(sender As Object, e As KeyEventArgs)
        Dim c As String = e.KeyCode.ToString
        If c = "Return" Then
            'ReCalcul()
        End If

    End Sub
    Sub BoutMotif_MouseClick(sender As Object, e As EventArgs)
        Dim com As BoutPerso = sender
        Dim ind As Integer
        Dim a As String

        ind = Val(com.Tag)
        'MotifCourant = Trim(BoutMotif.Item(ind).Text)
        MotifsForm.ShowDialog()
        If Trim(MotifsForm.retour) <> "" Then
            a = Trim(MotifsForm.retour)
            BoutMotif.Item(ind).Text = Trim(a)
            tbl = Split(a)
            If Trim(tbl(0)) = "Perso" Then
                PisteDurée.Item(ind).Enabled = False
                PisteSouche.Item(ind).Enabled = False
                'PisteDurée.Item(ind).Visible = False
            Else
                PisteDurée.Item(ind).Enabled = True
                PisteSouche.Item(ind).Enabled = True
                'PisteDurée.Item(ind).Visible = True
            End If

        End If
    End Sub

    Sub PisteMotif_Keydown(sender As Object, e As KeyEventArgs)
        e.SuppressKeyPress = True
    End Sub
    Sub PisteDurée_Keydown(sender As Object, e As KeyEventArgs)
        e.SuppressKeyPress = True
    End Sub
    Sub PisteAccent_Keydown(sender As Object, e As KeyEventArgs)
        e.SuppressKeyPress = True
    End Sub
    Sub PisteOctave_Keydown(sender As Object, e As KeyEventArgs)
        e.SuppressKeyPress = True
    End Sub
    Sub PisteDyn_Keydown(sender As Object, e As KeyEventArgs)
        e.SuppressKeyPress = True
    End Sub
    Sub PistePRG_Keydown(sender As Object, e As KeyEventArgs)
        e.SuppressKeyPress = True
    End Sub
    Sub PisteSouche_Keydown(sender As Object, e As KeyEventArgs)
        e.SuppressKeyPress = True
    End Sub
    Private Sub PisteRadio1_MouseDown(sender As Object, e As EventArgs) 'Handles RadioButton1.MouseDown
        Dim com As RadioButton = sender
        Dim ind As Integer
        Dim pst As Integer

        ind = Val(com.Tag)
        pst = Det_IndexPisteMidi(ind)
        ' Pan left
        If EnChargement = False Then
            'Dim volume As Byte = CByte(PisteVolume.Item(ind).Value)
            Dim canal As Byte = LesPistes.Item(pst).Canal
            If Not (SortieMidi.Item(ChoixSortieMidi).IsOpen) Then
                SortieMidi.Item(ChoixSortieMidi).Open()
            End If
            SortieMidi.Item(ChoixSortieMidi).SendControlChange(canal, 10, 0)

        End If
    End Sub
    Private Sub PisteRadio2_MouseDown(sender As Object, e As EventArgs) 'Handles RadioButton2.MouseDown
        Dim com As RadioButton = sender
        Dim ind As Integer
        Dim pst As Integer

        ind = Val(com.Tag)
        pst = Det_IndexPisteMidi(ind)
        ' pan centre
        If EnChargement = False Then
            'Dim volume As Byte = CByte(PisteVolume.Item(ind).Value)
            Dim canal As Byte = LesPistes.Item(pst).Canal
            If Not (SortieMidi.Item(ChoixSortieMidi).IsOpen) Then
                SortieMidi.Item(ChoixSortieMidi).Open()
            End If
            SortieMidi.Item(ChoixSortieMidi).SendControlChange(canal, 10, 64)
        End If
    End Sub
    Private Sub PisteRadio3_MouseDown(sender As Object, e As EventArgs) 'Handles RadioButton3.MouseDown
        Dim com As RadioButton = sender
        Dim ind As Integer
        Dim pst As Integer

        ind = Val(com.Tag)
        pst = Det_IndexPisteMidi(ind)
        ' pan droit
        If EnChargement = False Then
            'Dim volume As Byte = CByte(PisteVolume.Item(ind).Value)
            Dim canal As Byte = LesPistes.Item(pst).Canal
            If Not (SortieMidi.Item(ChoixSortieMidi).IsOpen) Then
                SortieMidi.Item(ChoixSortieMidi).Open()
            End If
            SortieMidi.Item(ChoixSortieMidi).SendControlChange(canal, 10, 127)
        End If
    End Sub
    Sub PistePanel_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        Dim com As Panel = sender
        Dim ind As Integer

        ind = Val(com.Tag)
    End Sub
    Sub PisteMute_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        Dim com As CheckBox = sender
        Dim ind As Integer

        ind = Val(com.Tag)

    End Sub
    Sub PisteMute_CheckedChange(sender As Object, e As EventArgs)
        Dim com As CheckBox = sender
        Dim ind As Integer
        Dim i As Integer
        Dim canMidi As Byte
        'Dim pst As Integer
        If Not EnChargement Then
            ind = Val(com.Tag)
            canMidi = Convert.ToByte(Det_IndexPisteMidi(ind))
            Mix.muteVolume.Item(canMidi).Checked = PisteMute.Item(ind).Checked
            '
            'Dim a As String = N_BLOC_MIDI(canMidi)

            Dim tbl() As String = N_BLOC_MIDI(canMidi).Split()
            For Each a As String In tbl
                i = Convert.ToInt16(a)
                If i <> ind Then
                    RemoveHandler PisteMute.Item(i).CheckedChanged, AddressOf PisteMute_CheckedChange
                    PisteMute.Item(i).Checked = PisteMute.Item(ind).Checked
                    AddHandler PisteMute.Item(i).CheckedChanged, AddressOf PisteMute_CheckedChange
                End If
            Next
        End If



        'ReCalcul()
    End Sub
    Sub PisteDelay_CheckedChanged(sender As Object, e As EventArgs)
        Dim com As CheckBox = sender
        Dim ind As Integer

        ind = Val(com.Tag)

    End Sub
    Sub PisteDelay_MouseDown(sender As Object, e As EventArgs)
        Dim com As CheckBox = sender
        Dim ind As Integer
        Dim a As String

        ind = Val(com.Tag)
        Dim canMidi As Byte
        Dim LDelay As New List(Of Boolean)
        Dim Variation As Integer
        Dim VariationCourante As Integer = Det_IndexVariation(ind)
        '
        canMidi = Det_IndexPisteMidi(ind)
        indicDelay = False
        indicVariation = -1
        indicPiste = -1
        For Variation = 0 To nb_Variations - 1
            a = Récup_Delay(canMidi, Variation)
            LDelay.Add(Convert.ToBoolean(a))
            If LDelay.Item(Variation) = True And Variation <> VariationCourante Then
                indicDelay = True
                indicVariation = Variation
                indicPiste = canMidi
                RemoveHandler PisteDelay.Item(ind).CheckedChanged, AddressOf PisteDelay_CheckedChanged
                PisteDelay.Item(ind).Checked = Not (PisteDelay.Item(ind).Checked)
                AddHandler PisteDelay.Item(ind).CheckedChanged, AddressOf PisteDelay_CheckedChanged
                Exit For
            End If
        Next Variation

    End Sub
    Sub PisteDelay_MouseUp(sender As Object, e As EventArgs)
        If indicDelay = True Then
            indicDelay = False
            MessageHV.PTypBouton = "OK"
            If LangueIHM = "fr" Then
                MessageHV.PTitre = "Un seul Delay par Piste"
                MessageHV.PContenuMess = "Vous avez déjà un delay sur cette piste" + Str(indicPiste + 1) + " dans la variation" + Str(indicVariation + 1) + " Un seul delay est possible par piste."
            Else
                MessageHV.PTitre = "Un seul Delay par Piste"
                MessageHV.PContenuMess = "You already have a delay on this track" + Str(indicPiste + 1) + " in variation" + Str(indicVariation + 1) + " Only one delay is possible per track."
            End If
            'Cacher_FormTransparents()
            MessageHV.ShowDialog()
        End If
    End Sub
    Sub PisteDébut_CheckedChanged(sender As Object, e As EventArgs)
        Dim com As Windows.Forms.CheckBox = sender
        Dim ind As Integer

        ind = Val(com.Tag)
        'ReCalcul()
    End Sub

    Public Function Det_IntervPistes(Magn As Integer) As String
        Select Case Magn
            Case Magneto_1
                Det_IntervPistes = "0" + " " + "5"
            Case Magneto_2
                Det_IntervPistes = "6" + " " + "11"
            Case Magneto_3
                Det_IntervPistes = "12" + " " + "17"
            Case Magneto_4
                Det_IntervPistes = "18" + " " + "23"
            Case Else
                Det_IntervPistes = "0" + " " + "5"
        End Select
    End Function
    Public Sub NomduSon_TextChanged(sender As Object, e As EventArgs)
        Dim com As TextBox = sender
        Dim ind As Integer
        ind = Val(com.Tag)
        Dim canal As Integer = Det_IndexPisteMidi(ind)
        Dim i, j, k As Integer
        Dim a As String

        'j = Det_IndexGénérateur(canal, i)
        ' controle/suppression caractères spéciaux interdit
        If InStr(Trim(NomduSon.Item(ind).Text), ",") <> 0 Or InStr(Trim(NomduSon.Item(ind).Text), "&") <> 0 Then
            k = NomduSon.Item(ind).SelectionStart
            RemoveHandler NomduSon.Item(ind).TextChanged, AddressOf NomduSon_TextChanged
            a = NomduSon.Item(ind).Text.Replace(",", "")
            NomduSon.Item(ind).Text = a
            a = NomduSon.Item(ind).Text.Replace("&", "")
            NomduSon.Item(ind).Text = a '
            NomduSon.Item(ind).SelectionStart = k - 1 'Len(a)
            AddHandler NomduSon.Item(ind).TextChanged, AddressOf NomduSon_TextChanged
        End If
        '
        ' mise à jour de la piste dans les variations
        For i = 0 To nb_Variations - 1
            If i <> ind Then
                j = Det_IndexGénérateur(canal, i)
                RemoveHandler NomduSon.Item(j).TextChanged, AddressOf NomduSon_TextChanged
                NomduSon.Item(j).Text = NomduSon.Item(ind).Text
                AddHandler NomduSon.Item(j).TextChanged, AddressOf NomduSon_TextChanged
            End If

        Next
        ' Mise à jour dans la table de mixage
        RemoveHandler Mix.NomduSon.Item(canal).TextChanged, AddressOf Mix.NomduSon_TextChanged
        Mix.NomduSon.Item(canal).Text = Trim(NomduSon.Item(j).Text)
        AddHandler Mix.NomduSon.Item(canal).TextChanged, AddressOf Mix.NomduSon_TextChanged
    End Sub
    Sub AffNomduson_MouseUp(sender As Object, e As EventArgs)
        Dim com As Label = sender
        Dim ind As Integer
        ind = Val(com.Tag)

        NomduSon.Item(ind).Visible = True
        NomduSon.Item(ind).Text = AffNomduson.Item(ind).Text
        NomduSon.Item(ind).Focus()


    End Sub
    Sub NomduSon_KeyUp(sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)

        If e.KeyCode = Keys.Enter Then
            Dim com As TextBox = sender
            Dim ind As Integer
            ind = Val(com.Tag)
            Dim canal As Integer = Det_IndexPisteMidi(ind)

            Dim i, j As Integer

            For i = 0 To nb_Variations - 1
                j = Det_IndexGénérateur(canal, i)
                AffNomduSon.Item(j).Text = NomduSon.Item(ind).Text
            Next
            NomduSon.Item(ind).Visible = False
        End If

    End Sub
    Sub SoloBout_MouseUp(sender As Object, e As EventArgs)

        Dim com As Button = sender
        Dim ind As Integer
        ind = Val(com.Tag)
        Mix.GestMute()
        Dim canMidi As Byte = Det_IndexPisteMidi(ind)
        TraitementSoloHypA(canMidi)

    End Sub
    Public Sub TraitementSoloHypA(CanMidi As Integer)
        Dim i, j, k As Integer
        If (SoloCours2 = True And CanMidiCours = CanMidi) Or SoloCours2 = False Then
            Mix.Gestion_Solo2(CanMidi) ' gestion du solo dans la table de mixage
            ' gestion des couleurs
            j = 0
            i = CanMidi
            '
            GestSolo(SoloCours2)
            ' activation du mode solo
            Do
                If SoloCours2 = False Then
                    k = CanMidi
                    SoloBout.Item(i).BackColor = Color.OrangeRed
                    SoloBout.Item(i).ForeColor = Color.Yellow
                    SoloBout.Item(i).Enabled = True
                Else
                    k = -1

                    SoloBout.Item(i).BackColor = Color.Beige
                    SoloBout.Item(i).ForeColor = Color.Black
                End If
                i += 6
                j += 1
            Loop Until j >= Module1.nb_Variations
            '

            PisteSolo = k
            '
            If SoloCours2 = False Then
                CanMidiCours = CanMidi
                SoloCours2 = True
            Else
                CanMidiCours = -1
                SoloCours2 = False
            End If
        End If
    End Sub
    Public Sub GestSolo(solocour2 As Boolean)
        Dim i As Integer
        If SoloCours2 = False Then
            For i = 0 To Module1.nb_BlocPistes - 1
                SoloBout.Item(i).Enabled = False
            Next
            For i = 0 To nb_PianoRoll - 1
                listPIANOROLL.Item(i).SoloBoutPR.Enabled = False
            Next
            Drums.SoloBoutDRM.Enabled = False
        Else
            For i = 0 To Module1.nb_BlocPistes - 1
                SoloBout.Item(i).Enabled = True
            Next
            For i = 0 To nb_PianoRoll - 1
                listPIANOROLL.Item(i).SoloBoutPR.Enabled = True
            Next
            Drums.SoloBoutDRM.Enabled = True
        End If

    End Sub





    Sub PisteVolume_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) ' pas utilisé
        Dim com As TrackBar = sender
        Dim ind As Integer

        ind = Val(com.Tag)
    End Sub

    Sub PisteMotif_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim com As Windows.Forms.ComboBox = sender
        Dim ind As Integer
        Dim a As String
        Dim tbl() As String


        ind = Val(com.Tag)
        a = PisteMotif.Item(ind).Text
        BoutMotif.Item(ind).Text = Trim(a) ' 
        tbl = Split(a)
        If Trim(tbl(0)) = "Perso" Then
            PisteDurée.Item(ind).Enabled = False
            PisteSouche.Item(ind).Enabled = False
            'PisteDurée.Item(ind).Visible = False
        Else
            PisteDurée.Item(ind).Enabled = True
            PisteSouche.Item(ind).Enabled = True
            'PisteDurée.Item(ind).Visible = True
        End If
        'ReCalcul()
    End Sub
    Sub PisteDyn_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim com As Windows.Forms.ComboBox = sender
        Dim ind As Integer

        ind = Val(com.Tag)
        'ReCalcul()

    End Sub
    Sub PisteDurée_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim com As Windows.Forms.ComboBox = sender
        Dim ind As Integer

        ind = Val(com.Tag)
        'ReCalcul()

    End Sub
    Sub PisteOctave_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim com As Windows.Forms.ComboBox = sender
        Dim ind As Integer

        ind = Val(com.Tag)
        'ReCalcul()


    End Sub
    Sub PisteSouche__SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim com As Windows.Forms.ComboBox = sender
        Dim ind As Integer

        ind = Val(com.Tag)
        'ReCalcul()
    End Sub
    Sub PisteAccent__SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim com As Windows.Forms.ComboBox = sender
        Dim ind As Integer

        ind = Val(com.Tag)
        'ReCalcul()
    End Sub
    Public Sub ReCalcul()

        If EnChargement = False And Horloge1.IsRunning And RejeuDirect.Checked Then
            EnRecalcul = True
            '
            StopRecalcul()
            PlayMidi.Enabled = False
            CalculArp(False)
            PlayArp()
            '
            EnRecalcul = False
        Else
            PlayHyperArp()
        End If
    End Sub
    Sub PistePRG_SelectedIndexChanged(sender As Object, e As EventArgs)

        Dim com As Windows.Forms.ComboBox = sender
        Dim ind As Integer
        Dim PRG As Integer
        Dim canal As Byte
        '
        ind = Val(com.Tag)
        Dim pst As Integer = Det_IndexPisteMidi(ind)


        canal = LesPistes.Item(pst).Canal
        PRG = PistePRG.Item(ind).SelectedIndex - 1
        If EnChargement = False Then
            If Trim(PRG) <> -1 Then
                If Not (SortieMidi.Item(ChoixSortieMidi).IsOpen) Then
                    SortieMidi.Item(ChoixSortieMidi).Open()
                End If
                tbl = Split(Trim(PRG))
                SortieMidi.Item(ChoixSortieMidi).SendProgramChange(canal, PRG)
            End If
        End If
    End Sub
    Sub MOTIFSPERSO_Création()
        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim k As Integer = 0
        '
        Dim Couleur As New List(Of Color)
        Dim c As New Color
        '
        For i = 0 To TabControl1.TabPages.Count - 1
            If i = 0 Then
                c = Color.Cornsilk
                Couleur.Add(c)
            ElseIf i = 1 Then
                c = Color.LightBlue
                Couleur.Add(c)
            ElseIf i = 2 Then
                c = Color.Beige
                Couleur.Add(c)
            ElseIf i = 3 Then
                c = Color.Linen
                Couleur.Add(c)
            ElseIf i = 4 Then
                c = Color.NavajoWhite
                Couleur.Add(c)
            ElseIf i = 5 Then
                c = Color.PaleTurquoise
                Couleur.Add(c)
            End If
        Next

        i = 0
        Do
            MotifsPerso.Add(New Grid)
            TabControl1.TabPages(i).Controls.Add(MotifsPerso.Item(i))
            ' Positionnement des grilles
            MotifsPerso.Item(i).Dock = DockStyle.None
            '
            MotifsPerso.Item(i).AllowDrop = False
            MotifsPerso.Item(i).AllowUserPaste = ClipboardDataEnum.None
            MotifsPerso.Item(i).AllowUserReorderColumn = False
            MotifsPerso.Item(i).AllowUserResizing = False
            MotifsPerso.Item(i).AllowUserSort = False


            Dim p As New Point(0, 0)
            MotifsPerso.Item(i).Location = p
            MotifsPerso.Item(i).Width = 540
            MotifsPerso.Item(i).Height = 242
            MotifsPerso.Item(i).ScrollBars = ScrollBarsEnum.None
            '
            MotifsPerso.Item(i).BackColorSel = Color.White

            MotifsPerso.Item(i).Rows = 7
            MotifsPerso.Item(i).Cols = 17
            '
            ' Traitement des colonnes et lignes fixes
            MotifsPerso.Item(i).FixedRows = 2
            MotifsPerso.Item(i).Row(0).Height = 53
            MotifsPerso.Item(i).Row(1).Height = 30
            MotifsPerso.Item(i).Column(0).Width = 24
            MotifsPerso.Item(i).BackColorFixed = Couleur(i)
            '
            ' Traitement des colonnes
            k = 1
            For j = 1 To MotifsPerso.Item(i).Cols - 1
                MotifsPerso.Item(i).Column(j).Width = 32
                MotifsPerso.Item(i).Cell(0, j).Text = Str(j)
                MotifsPerso.Item(i).Cell(1, j).Text = Str(k)
                k = k + 1
                If k > 4 Then k = 1
                '
                If j = 1 Or j = 5 Or j = 9 Or j = 13 Then
                    CouleurDivisions(i, j)
                End If

            Next
            ' Traitement des lignes
            For j = 2 To MotifsPerso.Item(i).Rows - 1
                MotifsPerso.Item(i).Row(j).Height = 31
                If j = 2 Then
                    MotifsPerso.Item(i).Cell(j, 0).Text = "V5" '"A"
                ElseIf j = 3 Then
                    MotifsPerso.Item(i).Cell(j, 0).Text = "V4" '"B"
                ElseIf j = 4 Then
                    MotifsPerso.Item(i).Cell(j, 0).Text = "V3" '"C"
                ElseIf j = 5 Then
                    MotifsPerso.Item(i).Cell(j, 0).Text = "V2" '"D"
                ElseIf j = 6 Then
                    MotifsPerso.Item(i).Cell(j, 0).Text = "V1" '"E"
                End If
            Next

            MotifsPerso.Item(i).Tag = i
            AddHandler MotifsPerso.Item(i).MouseUp, AddressOf MotifsPerso_MouseUp
            AddHandler MotifsPerso.Item(i).KeyDown, AddressOf MotifsPerso_KeyDown
            '
            ' création des boutons c/v (copier/coller)
            ' ****************************************
            Bcopier.Add(New Button)
            Bcoller.Add(New Button)
            TabControl1.TabPages(i).Controls.Add(Bcopier(i))
            TabControl1.TabPages(i).Controls.Add(Bcoller(i))
            MotifsPerso.Item(i).Visible = True

            If LangueIHM = "fr" Then
                Bcopier(i).Text = "Copier"
                Bcoller(i).Text = "Coller"
            Else
                Bcopier(i).Text = "Copy"
                Bcoller(i).Text = "Paste"
            End If

            Dim PP As Point
            Dim S As Size
            PP.X = 1 ' MotifsPerso.Item(i).Location.X
            PP.Y = 1 ' MotifsPerso.Item(i).Location.Y
            Bcopier(i).Location = PP

            S.Width = 58
            S.Height = 28
            Bcopier(i).Size = S
            Bcopier(i).Visible = True
            Bcopier(i).BringToFront()
            Bcopier(i).Tag = i
            '
            PP.X = 1 ' MotifsPerso.Item(i).Location.X
            PP.Y = 28 ' MotifsPerso.Item(i).Location.Y
            Bcoller(i).Location = PP

            S.Width = 58
            S.Height = 28
            Bcoller(i).Size = S
            Bcoller(i).Visible = True
            Bcoller(i).BringToFront()
            Bcoller(i).Tag = i
            '
            AddHandler Bcopier.Item(i).MouseUp, AddressOf Bcopier_MouseUp
            AddHandler Bcoller.Item(i).MouseUp, AddressOf Bcoller_MouseUp
            '
            i = i + 1
        Loop Until i = TabControl1.TabPages.Count
        StyleCelluleMotif()

    End Sub
    Sub NomsColMotifs()
        Dim i, j As Integer

        For i = 0 To i = TabControl1.TabPages.Count - 1
            For j = 2 To MotifsPerso.Item(i).Rows - 1
                MotifsPerso.Item(i).Row(j).Height = 31
                If j = 2 Then
                    MotifsPerso.Item(i).Cell(j, 0).Text = "V5" '"A"
                ElseIf j = 3 Then
                    MotifsPerso.Item(i).Cell(j, 0).Text = "V4" '"B"
                ElseIf j = 4 Then
                    MotifsPerso.Item(i).Cell(j, 0).Text = "V3" '"C"
                ElseIf j = 5 Then
                    MotifsPerso.Item(i).Cell(j, 0).Text = "V2" '"D"
                ElseIf j = 6 Then
                    MotifsPerso.Item(i).Cell(j, 0).Text = "V1" '"E"
                End If
            Next j
        Next i
    End Sub

    Sub CouleurDivisions(numGrid As Integer, colonne As Integer)
        Dim i As Integer
        For i = 2 To MotifsPerso.Item(numGrid).Rows - 1
            MotifsPerso.Item(numGrid).Cell(i, colonne).BackColor = Couleur_Div
        Next
    End Sub
    Sub StyleCelluleMotif()
        Dim h As Integer
        Dim i As Integer
        Dim j As Integer
        For h = 0 To TabControl1.TabPages.Count - 1
            For i = 2 To MotifsPerso.Item(h).Rows - 1
                For j = 1 To MotifsPerso.Item(h).Cols - 1
                    MotifsPerso.Item(h).Cell(i, j).Font = New Font("Arial", 18, FontStyle.Bold)
                    MotifsPerso.Item(h).Cell(i, j).Alignment = AlignmentEnum.CenterGeneral
                Next j
            Next i
            For j = 0 To MotifsPerso.Item(h).Cols - 1
                MotifsPerso.Item(h).Cell(0, j).Font = New Font("Calibri", 9, FontStyle.Bold)
                MotifsPerso.Item(h).Cell(1, j).Font = New Font("Verdana", 6.5, FontStyle.Regular)
            Next
            For i = 0 To MotifsPerso.Item(h).Rows - 1
                MotifsPerso.Item(h).Cell(i, 0).Font = New Font("Calibri", 9, FontStyle.Bold)
            Next
        Next h




    End Sub
    Private Sub MotifsPerso_KeyDown(ByVal Sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Dim com As FlexCell.Grid = Sender
        Dim ind As Integer = Val(com.Tag)

        Dim i As Integer = MotifsPerso.Item(ind).ActiveCell.Row
        Dim j As Integer = MotifsPerso.Item(ind).ActiveCell.Col

        Dim a As String = Trim(MotifsPerso.Item(ind).Cell(i, j).Text)

        ' Prevent the user from cutting the cell content by pressing Ctrl+X,Ctrl+C,Ctrl+V key combination please add the following programming code:
        If (e.Control And e.KeyCode = Keys.X) Or (e.Control And e.KeyCode = Keys.V) Or (e.Control And e.KeyCode = Keys.C) Then
            'e.Handled = True
        End If
        If e.KeyCode = Keys.Up Then
            Select Case a
                Case "-"
                    MotifsPerso.Item(ind).Cell(i, j).Text = "*"
                Case "*"
                    MotifsPerso.Item(ind).Cell(i, j).Text = "+"
            End Select

        End If
        If e.KeyCode = Keys.Down Then
            Select Case a
                Case "+"
                    MotifsPerso.Item(ind).Cell(i, j).Text = "*"
                Case "*"
                    MotifsPerso.Item(ind).Cell(i, j).Text = "-"
            End Select
        End If

    End Sub
    Private Sub Bcopier_MouseUp(Sender As Object, e As MouseEventArgs)
        Dim com As Button = Sender
        Dim ind As Integer = Val(com.Tag)

        ClipBPerso = ListSpecific(ind)
    End Sub
    Private Sub Bcoller_MouseUp(Sender As Object, e As MouseEventArgs)
        Dim com As Button = Sender
        Dim ind As Integer = Val(com.Tag)
        Dim tbl() As String
        If Trim(ClipBPerso) <> "" Then
            tbl = Split(ClipBPerso, ",")
            tbl(1) = ind
            ClipBPerso = Join(tbl, ",")
            MAJ_Specific(ClipBPerso)
        End If
    End Sub
    Private Sub MotifsPerso_MouseUp(Sender As Object, e As MouseEventArgs)
        Dim com As FlexCell.Grid = Sender
        Dim FirstR, Firstc, LastR, LastC, ligne As Integer
        Dim i, ind As Integer
        '
        If e.Button = MouseButtons.Left Then '  And My.Computer.Keyboard.CtrlKeyDown
            ind = Val(com.Tag)
            '
            'If My.Computer.Keyboard.CtrlKeyDown Then
            MotifsPerso.Item(ind).AutoRedraw = False
            ligne = MotifsPerso.Item(ind).Selection.FirstRow
            FirstR = 1
            Firstc = MotifsPerso.Item(ind).Selection.FirstCol
            LastR = MotifsPerso.Item(ind).Rows - 1
            LastC = MotifsPerso.Item(ind).Selection.LastCol
            '
            If My.Computer.Keyboard.CtrlKeyDown Then '
                If Trim(MotifsPerso.Item(ind).Cell(ligne, Firstc).Text) = "" Then ' ECRIRE NOTE
                    'EffacerColonne(ind, Firstc, LastC, Couleur_Div, Color.White)
                    MotifsPerso.Item(ind).Cell(ligne, Firstc).Text = "*"
                    'If My.Computer.Keyboard.AltKeyDown Then
                    'MotifsPerso.Item(ind).Cell(ligne, Firstc).Text = "+"
                    'End If
                    For i = Firstc To LastC
                        MotifsPerso.Item(ind).Cell(ligne, i).BackColor = Couleur_Notes ' dessin de la longueur de la note
                    Next i
                    'Effacer2(ind, ligne, Firstc, LastC) ' ici on efface la longueur de la note qui vient d'être écrite s'il elle se situe au dessus d'une note déjà existante
                Else
                    '' EffacerColonne(ind, Firstc, LastC, Couleur_Div, Color.White) ' EFFACER NOTE
                    EffacerLongueur(ind, ligne, Firstc, Couleur_Div, Color.White)
                End If
            Else
                If My.Computer.Keyboard.AltKeyDown Then
                    ModifierLongeur(ind, ligne, Firstc, Couleur_Div, Color.White) 'MODIFIER LONGUEUR
                End If
            End If
            '
            MotifsPerso.Item(ind).AutoRedraw = True
            MotifsPerso.Item(ind).Refresh()
            '
            'ReCalcul()
        End If
    End Sub

    Sub Effacer2(ind As Integer, ligne As Integer, Firstc As Integer, LastC As Integer)
        Dim i, j As Integer
        Dim jj As Integer
        Dim PrésenceSurLong As Boolean = False
        Dim Col2 As Integer

        MotifsPerso.Item(ind).AutoRedraw = False
        ' réduire la longueur de la note à écrire située au dessus de la note trouvée
        For j = Firstc To LastC
            For i = ligne + 1 To MotifsPerso.Item(ind).Rows - 1
                If Trim(MotifsPerso.Item(ind).Cell(i, j).Text) = "*" Then
                    Col2 = j
                    For jj = Col2 To LastC
                        If jj = 1 Or jj = 5 Or jj = 9 Or jj = 13 Then
                            MotifsPerso.Item(ind).Cell(ligne, jj).BackColor = Couleur_Div
                        Else
                            MotifsPerso.Item(ind).Cell(ligne, jj).BackColor = Color.White
                        End If
                    Next jj
                End If
            Next i
        Next j

        ' réduire la longueur de la note à écrire située en dessous de la note trouvée
        PrésenceSurLong = False
        For j = Firstc To LastC
            For i = ligne - 1 To 1 Step -1
                If Trim(MotifsPerso.Item(ind).Cell(i, j).Text) = "*" Then
                    Col2 = j
                    For jj = Col2 To LastC
                        If jj = 1 Or jj = 5 Or jj = 9 Or jj = 13 Then
                            MotifsPerso.Item(ind).Cell(ligne, jj).BackColor = Couleur_Div
                        Else
                            MotifsPerso.Item(ind).Cell(ligne, jj).BackColor = Color.White
                        End If
                    Next jj
                End If
            Next i
        Next j
        '
        MotifsPerso.Item(ind).AutoRedraw = True
        MotifsPerso.Item(ind).Refresh()
    End Sub
    Sub ModifierLongeur(numGrid As Integer, Ligne As Integer, Col As Integer, Couleur1 As Color, Couleur2 As Color)
        Dim i As Integer = Ligne
        Dim j As Integer = Col
        Dim Presence_Etoile As Boolean = False
        '
        If présenceEtoile(numGrid, Ligne) Then '
            MotifsPerso.Item(numGrid).AutoRedraw = False
            If MotifsPerso.Item(numGrid).Cell(Ligne, j).BackColor = Couleur_Notes Then
                For j = Col To MotifsPerso.Item(numGrid).Cols - 1  ' RACCOURCIR LA NOTE
                    If MotifsPerso.Item(numGrid).BackColor = Color.White _
                                            Or Trim(MotifsPerso.Item(numGrid).Cell(i, j).Text) = "*" Then
                        Exit For
                    End If
                    If j = 1 Or j = 5 Or j = 9 Or j = 13 Then
                        MotifsPerso.Item(numGrid).Cell(i, j).BackColor = Couleur1
                    Else
                        MotifsPerso.Item(numGrid).Cell(i, j).BackColor = Couleur2
                    End If
                Next
            Else
                For j = Col To 1 Step -1
                    If Trim(MotifsPerso.Item(numGrid).Cell(i, j).Text) = "*" Then
                        Presence_Etoile = True
                        Exit For
                    End If
                Next
                If Presence_Etoile Then
                    For j = j To Col
                        MotifsPerso.Item(numGrid).Cell(i, j).BackColor = Couleur_Notes ' RALLONGER LA NOTE
                    Next
                End If
            End If
            MotifsPerso.Item(numGrid).AutoRedraw = True
            MotifsPerso.Item(numGrid).Refresh()
        End If
    End Sub
    Sub EffacerColonne(numGrid As Integer, FirstC As Integer, LastC As Integer, Couleur1 As Color, Couleur2 As Color)
        Dim i As Integer
        Dim j As Integer

        MotifsPerso.Item(numGrid).AutoRedraw = False
        For i = 2 To MotifsPerso.Item(numGrid).Rows - 1
            For j = FirstC To LastC
                If Trim(MotifsPerso.Item(numGrid).Cell(i, j).Text) = "*" Or
                                                            MotifsPerso.Item(numGrid).Cell(i, j).BackColor = Couleur_Notes Then
                    EffacerLongueur(numGrid, i, FirstC, Couleur1, Couleur2)
                Else
                    MotifsPerso.Item(numGrid).Cell(i, j).Text = " "
                    If j = 1 Or j = 5 Or j = 9 Or j = 13 Then
                        MotifsPerso.Item(numGrid).Cell(i, j).BackColor = Couleur1
                    Else
                        MotifsPerso.Item(numGrid).Cell(i, j).BackColor = Couleur2
                    End If
                End If
            Next j
        Next i
        MotifsPerso.Item(numGrid).AutoRedraw = True
        MotifsPerso.Item(numGrid).Refresh()
    End Sub
    Sub EffacerLongueur(numGrid As Integer, Ligne As Integer, Col As Integer, Couleur1 As Color, Couleur2 As Color)

        ' Cette procédure efface la note dont la tête ("*")se trouve dans la même colonne que la tête de note de la nouvelle note
        Dim i As Integer = Ligne
        Dim j As Integer = Col
        MotifsPerso.Item(numGrid).AutoRedraw = False

        For j = Col To 16
            MotifsPerso.Item(numGrid).Cell(i, j).Text = " "
            If j = 1 Or j = 5 Or j = 9 Or j = 13 Then
                MotifsPerso.Item(numGrid).Cell(i, j).BackColor = Couleur1
            Else
                MotifsPerso.Item(numGrid).Cell(i, j).BackColor = Couleur2
            End If
            If j <> 16 Then
                If (MotifsPerso.Item(numGrid).Cell(i, j + 1).BackColor = Color.White) Or (Trim(MotifsPerso.Item(numGrid).Cell(i, j + 1).Text) = "*") Then
                    Exit For
                End If
            End If
        Next j
        'Loop Until j = 16 Or MotifsPerso.Item(numGrid).Cell(i, j).BackColor = Color.White Or Trim(MotifsPerso.Item(numGrid).Cell(i, j).Text) = "*"
        MotifsPerso.Item(numGrid).AutoRedraw = True
        MotifsPerso.Item(numGrid).Refresh()
    End Sub
    Function PrésenceEtoile(numGrid As Integer, Ligne As Integer) As Boolean
        Dim i As Integer = Ligne
        Dim j As Integer
        PrésenceEtoile = False

        For j = 0 To MotifsPerso.Item(numGrid).Cols - 1
            If Trim(MotifsPerso.Item(numGrid).Cell(i, j).Text) = "*" Then
                PrésenceEtoile = True
                Exit For
            End If

        Next
    End Function
    Sub PIANOROLL_Création2()
        Dim i As Integer
        Dim CoulBarout As Color

        For i = 0 To nb_PianoRoll - 1
            Select Case i
                Case 0      ' Pianoroll7
                    CoulBarout = Color.FromArgb(252, 248, 200)
                Case 1      ' Pianoroll8
                    CoulBarout = Color.Khaki 'Color.FromArgb(178, 194, 114) 
                Case 2      ' Pianoroll9
                    CoulBarout = Color.FromArgb(241, 212, 179)
                Case 3      ' Pianoroll1
                    CoulBarout = Color.Gainsboro
                Case 4      ' Pianoroll2
                    CoulBarout = Color.Cornsilk
                Case 5      ' Pianoroll3
                    CoulBarout = Color.SeaShell
                Case 6      ' Pianoroll114
                    CoulBarout = Color.White

            End Select
            ' instenciation des PROLL
            If i < 3 Then
                Dim a As New PianoRoll(i + 6, CoulBarout)
                listPIANOROLL.Add(a)
            Else
                Dim b As New PianoRoll((i + 7), CoulBarout) ' au cas où on rajoute des pianoroll après la batterie
                listPIANOROLL.Add(b)
            End If

            listPIANOROLL.Item(i).F1.Visible = False

            ' Mise à jour des tag de F1 : F1.Tag= N° onglet
            ' ********************************************
            If i < 3 Then
                listPIANOROLL.Item(i).F1.Tag = i + 1
            Else
                listPIANOROLL.Item(i).F1.Tag = i + 2
            End If
            '
            ' Maj des propriétés de Pianoroll

            listPIANOROLL.Item(i).PLangue = LangueIHM
            listPIANOROLL.Item(i).PNbMesures = nbMesures ' Arrangement1.nbMesures
            listPIANOROLL.Item(i).PMétrique = Métrique.Text
            listPIANOROLL.Item(i).PnbRépétitionMax = nbRépétitionMax
            listPIANOROLL.Item(i).PListAcc = Det_ListAcc()
            listPIANOROLL.Item(i).PListGam = Det_ListGam()
            listPIANOROLL.Item(i).PListMarq = Det_ListMarq()
            listPIANOROLL.Item(i).PN_Can1erPianoR = 7
            '
            ' Appel de la construction
            listPIANOROLL(i).Construction_F1()
            ' Maj Paramètres de gestion
            PIANOROLLChargé.Add(True)
            '
            listPIANOROLL.Item(i).F1.Visible = False
            listPIANOROLL.Item(i).F1.FormBorderStyle = FormBorderStyle.None
            listPIANOROLL.Item(i).F1.TopLevel = False  '
            listPIANOROLL.Item(i).F1.TopMost = False   ' un seul des 2 suffit ?
            listPIANOROLL(i).F1.Dock = DockStyle.Fill
            If i < 3 Then
                TabControl4.TabPages.Item(i + 1).Visible = False
                TabControl4.TabPages.Item(i + 1).Controls.Add(listPIANOROLL(i).F1) ' ajout de F1 dans l'onglet
            Else
                TabControl4.TabPages.Item(i + 2).Visible = False
                TabControl4.TabPages.Item(i + 2).Controls.Add(listPIANOROLL(i).F1) ' ajout de F1 dans l'onglet
            End If
            listPIANOROLL.Item(i).F1.Show()
            listPIANOROLL.Item(i).F1.Visible = True
            '
            ' Mise à jour des tags des onglet
            ' *******************************
            If i < 3 Then
                TabControl4.TabPages.Item(i).Tag = i + 1
            Else
                TabControl4.TabPages.Item(i).Tag = i + 2
            End If
        Next
    End Sub
    Sub CHORDROLL_Création()
        Dim i, ii As Integer
        Dim CoulBarout As Color

        For i = 0 To nb_ChordRoll - 1
            Select Case i
                Case 0      ' Pianoroll7
                    CoulBarout = Color.Khaki 'Color.FromArgb(252, 248, 200)
                Case 1      ' Pianoroll8
                    CoulBarout = Color.Khaki 'Color.Khaki 'Color.FromArgb(178, 194, 114) 

            End Select
            ' instenciation des PROLL
            ii = (i) + (nb_PianoRoll)
            Dim b As New PianoRoll((ii + 7), CoulBarout) ' au cas où on rajoute des pianoroll après la batterie
            listPIANOROLL.Add(b)
            '
            listPIANOROLL.Item(ii).F1.Visible = False

            ' Mise à jour des tag de F1 : F1.Tag= N° onglet
            ' ********************************************

            listPIANOROLL.Item(ii).F1.Tag = ii '+ 1
            '
            ' Maj des propriétés de Pianoroll

            listPIANOROLL.Item(ii).PLangue = LangueIHM
            listPIANOROLL.Item(ii).PNbMesures = nbMesures ' Arrangement1.nbMesures
            listPIANOROLL.Item(ii).PMétrique = Métrique.Text
            listPIANOROLL.Item(ii).PnbRépétitionMax = nbRépétitionMax
            listPIANOROLL.Item(ii).PListAcc = Det_ListAcc()
            listPIANOROLL.Item(ii).PListGam = Det_ListGam()
            listPIANOROLL.Item(ii).PListMarq = Det_ListMarq()
            listPIANOROLL.Item(ii).PN_Can1erPianoR = 7
            '
            ' Appel de la construction
            listPIANOROLL(ii).Construction_F1()
            ' Maj Paramètres de gestion
            PIANOROLLChargé.Add(True)
            '
            listPIANOROLL.Item(ii).F1.Visible = False
            listPIANOROLL.Item(ii).F1.FormBorderStyle = FormBorderStyle.None
            listPIANOROLL.Item(ii).F1.TopLevel = False  '
            listPIANOROLL.Item(ii).F1.TopMost = False   ' un seul des 2 suffit ?
            listPIANOROLL(ii).F1.Dock = DockStyle.Fill

            TabControl4.TabPages.Item(ii + 2).Visible = False
            TabControl4.TabPages.Item(ii + 2).Controls.Add(listPIANOROLL(ii).F1) ' ajout de F1 dans l'onglet

            listPIANOROLL.Item(ii).F1.Show()
            listPIANOROLL.Item(ii).F1.Visible = True
            '
            ' Mise à jour des tags des onglets
            ' ********************************
            TabControl4.TabPages.Item(ii).Tag = ii + 2
        Next
    End Sub


    Sub AUTOMATION_Création()
        Dim ind As Integer
        ind = TabControl2.TabPages.IndexOf(TabPage4) ' index de tabpage4 dans Tabcontrol2
        Dim bb As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CalquesMIDI\HyperArp\LangueIHM", "LangueIHM", Nothing)
        Automation1.PLangue = bb
        TabControl2.TabPages.Item(ind).Controls.Add(Automation1.F4)
        Automation1.F4.Dock = DockStyle.Fill
        Automation1.F4.Visible = True
        Automation1.F4.Tag = ind
        Automation1.Construction_Menu()
        Automation1.Construction_Barrout()
        Automation1.Construction_BarroutLocale()
        'Automation1.Maj_Tooltips()
    End Sub
    Sub DRUMS_Création()
        ' L'objet drums est créé en public dans les variables globales de form1
        Drums.F2.Visible = False
        'Drums.F2.Tag = nb_PianoRoll + 1
        Drums.F2.Tag = 4 ' la batterie est toujours en position 4 dans les onglets (sauf si on redescend à moins de 3 pistes).
        '
        ' Maj des propriétés de DrumEdit
        ' ******************************
        Drums.PLangue = LangueIHM
        Drums.PNbMesures = nbMesures '
        Drums.PMétrique = Métrique.Text
        Drums.PnbRépétitionMax = nbRépétitionMax
        Drums.PListAcc = Det_ListAcc()
        Drums.PListMarq = Det_ListMarq()
        ' Appel de la construction
        ' ************************
        Drums.Construction_F2()
        Drums.F2.FormBorderStyle = FormBorderStyle.None
        Drums.F2.TopLevel = False  '
        Drums.F2.TopMost = False   ' un seul des 2 suffit ?
        Drums.F2.Dock = DockStyle.Fill
        TabControl4.TabPages.Item(Convert.ToByte(Drums.F2.Tag)).Controls.Add(Drums.F2)

        Drums.F2.Visible = True
    End Sub
    Sub MIXAGE_Création()
        ' L'objet drums est créé en public dans les variables globales de form1
        Mix.F3.Visible = False
        Mix.F3.Tag = nb_PianoRoll + nb_ChordRoll + 2
        '
        ' Maj des propriétés de DrumEdit
        ' ******************************
        Mix.PLangue = LangueIHM

        ' Appel de la construction
        ' ************************
        Mix.Construction_Formulaire()
        Mix.F3.FormBorderStyle = FormBorderStyle.None
        Mix.F3.TopLevel = False  '
        Mix.F3.TopMost = False   ' un seul des 2 suffit ?
        TabControl4.TabPages.Item(nb_PianoRoll + +nb_ChordRoll + 2).Controls.Add(Mix.F3)
        Mix.F3.Dock = DockStyle.Fill
        Mix.F3.BackColor = Color.Beige
        Mix.F3.Visible = True
    End Sub
    Sub TONS_Création()
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer

        Dim PosX As Integer
        Dim PosY As Integer
        '
        Dim Longueur As Integer
        Dim Hauteur As Integer
        '
        Dim DepL As Integer
        Dim DepH As Integer
        '
        PosY = 75 '65 '28
        PosX = 6 '
        '
        Longueur = 75
        Hauteur = 90 '100
        '
        DepL = 0
        DepH = 2
        j = 0
        k = 1
        i = 0
        '
        Do ' 
            TabTons.Add(New Label)
            Me.Panel8.Controls.Add(TabTons.Item(i))
            j = j + 1
            '
            ' Label des accords
            ' *****************
            TabTons.Item(i).Size = New System.Drawing.Size(Longueur, Hauteur / 2)
            TabTons.Item(i).Location = New Point(PosX + DepL, PosY + DepH)
            TabTons.Item(i).BorderStyle = BorderStyle.FixedSingle

            TabTons.Item(i).Text = "C"
            TabTons.Item(i).Tag = Str(i)
            TabTons.Item(i).Visible = True
            TabTons.Item(i).TextAlign = ContentAlignment.MiddleCenter
            '
            If i <= 6 Then
                TabTons.Item(i).BackColor = DicoCouleur.Item("C") 'Couleur_Accord_Majeur 'Couleur_Accords
                TabTons.Item(i).Font = New System.Drawing.Font("Arial", 10, FontStyle.Bold)
                TabTons.Item(i).Visible = True
            End If
            If i >= 7 And i <= 13 Then
                TabTons.Item(i).BackColor = Couleur_Accord_Mineur 'Color.LightGoldenrodYellow
                TabTons.Item(i).Font = New System.Drawing.Font("Arial", 10, FontStyle.Bold)
                TabTons.Item(i).Visible = True
            End If
            If i >= 14 And i <= 20 Then
                TabTons.Item(i).BackColor = Couleur_Accord_Mineur ' Color.LightGoldenrodYellow
                TabTons.Item(i).Font = New System.Drawing.Font("Arial", 10, FontStyle.Bold)
                TabTons.Item(i).Visible = True
            End If

            ' Evènements label Accords
            ' ************************
            AddHandler TabTons.Item(i).MouseDown, AddressOf TabTons_MouseDown
            AddHandler TabTons.Item(i).MouseUp, AddressOf TabTons_MouseUp
            'AddHandler TabTons.Item(i).MouseMove, AddressOf TabTons_MouseMove
            'AddHandler TabTons.Item(i).MouseLeave, AddressOf TabTons_MouseLeave
            '
            '
            DepL = DepL + Longueur + 1
            '
            If j = 7 Then
                DepH = DepH + Hauteur - 18 '38
                j = 0
                DepL = 0
            End If
            'TabTons.Item(i).Refresh()
            i = i + 1
        Loop Until i > 20
        '
        ' Construction des étiquettes de titres des modes
        ' ***********************************************
        PosX = 10
        DepL = 0
        '
        PosY = 60
        '
        For i = 0 To 2
            TabTonsTitreMode.Add(New Label)
            'Me.TabPage_Tonalité.Controls.Add(TabTonsTitreMode.Item(i))
            Me.Panel8.Controls.Add(TabTonsTitreMode.Item(i))
            '
            ' Titre du mode
            ' *************
            TabTonsTitreMode.Item(i).Size = New System.Drawing.Size(300, 15)
            TabTonsTitreMode.Item(i).BorderStyle = BorderStyle.None
            TabTonsTitreMode.Item(i).Font = New System.Drawing.Font("Papyrus", 9, FontStyle.Regular)
            TabTonsTitreMode.Item(i).Location = New Point((PosX + DepL) - 3, PosY)
            TabTonsTitreMode.Item(i).ForeColor = Color.Black
            Select Case i
                Case 0
                    TabTonsTitreMode.Item(i).Text = "Mode Majeur"
                    PosY = 132
                Case 1
                    TabTonsTitreMode.Item(i).Text = "Mode Mineur Harmonique"
                    PosY = 203
                Case 2
                    TabTonsTitreMode.Item(i).Text = "Mode Mineur Mélodique"
            End Select

            TabTonsTitreMode.Item(i).Visible = True

        Next i
        ' Constructon des étiquettes de degrés
        ' ************************************
        Longueur = 74
        Hauteur = 20
        '
        PosX = 6 '12
        PosY = 38 '3
        '
        DepL = 0
        i = -1
        Do
            i = i + 1
            TabTonsDegrés.Add(New Label)
            Me.Panel8.Controls.Add(TabTonsDegrés.Item(i))
            '
            TabTonsDegrés.Item(i).Size = New System.Drawing.Size(Longueur, Hauteur)
            TabTonsDegrés.Item(i).Font = New System.Drawing.Font("Calibri", 11, FontStyle.Bold)
            TabTonsDegrés.Item(i).Location = New Point(PosX + DepL, PosY)
            TabTonsDegrés.Item(i).BorderStyle = BorderStyle.FixedSingle
            TabTonsDegrés.Item(i).TextAlign = ContentAlignment.MiddleCenter
            TabTonsDegrés.Item(i).Text = "I"
            TabTonsDegrés.Item(i).Tag = Str(i)
            TabTonsDegrés.Item(i).Visible = True
            TabTonsDegrés.Item(i).ForeColor = Color.Black
            '
            Select Case i
                Case 0
                    TabTonsDegrés.Item(i).Text = "I"
                Case 1
                    TabTonsDegrés.Item(i).Text = "II"
                Case 2
                    TabTonsDegrés.Item(i).Text = "III"
                Case 3
                    TabTonsDegrés.Item(i).Text = "IV"
                Case 4
                    TabTonsDegrés.Item(i).Text = "V"
                Case 5
                    TabTonsDegrés.Item(i).Text = "VI"
                Case 6
                    TabTonsDegrés.Item(i).Text = "VII"
            End Select
            '
            ' Evènements sur label Degrés
            ' ***************************
            AddHandler TabTonsDegrés.Item(i).MouseDown, AddressOf TabTonsDegrés_MouseDown
            '
            DepL = DepL + Longueur + 2
        Loop Until i >= 6
        '
    End Sub
    Sub TONS_Création2()
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer

        Dim PosX As Integer
        Dim PosY As Integer
        '
        Dim Longueur As Integer
        Dim Hauteur As Integer
        '
        Dim DepL As Integer
        Dim DepH As Integer
        '
        PosY = 87 '75'65 '28
        PosX = 6 '
        '
        Longueur = 75
        Hauteur = 90 '100
        '
        DepL = 0
        DepH = 2
        j = 0
        k = 1
        i = 0
        '
        '
        Do ' 
            TabTons.Add(New Label)
            Me.Panel8.Controls.Add(TabTons.Item(i))
            j = j + 1
            '
            ' Label des accords
            ' *****************
            TabTons.Item(i).Size = New System.Drawing.Size(Longueur, Hauteur / 2)
            TabTons.Item(i).Location = New Point(PosX + DepL, PosY + DepH)
            TabTons.Item(i).BorderStyle = BorderStyle.FixedSingle

            TabTons.Item(i).Text = "C"
            TabTons.Item(i).Tag = Str(i)
            TabTons.Item(i).Visible = True
            TabTons.Item(i).TextAlign = ContentAlignment.MiddleCenter
            '
            If i <= 6 Then
                TabTons.Item(i).BackColor = Couleur_Accord_Majeur 'Couleur_Accords
                TabTons.Item(i).Font = New System.Drawing.Font("Arial", 10, FontStyle.Bold)
                TabTons.Item(i).Visible = True
            End If
            If i >= 7 And i <= 13 Then
                TabTons.Item(i).BackColor = Couleur_Accord_Mineur 'Color.LightGoldenrodYellow
                TabTons.Item(i).Font = New System.Drawing.Font("Arial", 10, FontStyle.Bold)
                TabTons.Item(i).Visible = True
            End If
            If i >= 14 And i <= 20 Then
                TabTons.Item(i).BackColor = Couleur_Accord_Mineur ' Color.LightGoldenrodYellow
                TabTons.Item(i).Font = New System.Drawing.Font("Arial", 10, FontStyle.Bold)
                TabTons.Item(i).Visible = False
            End If

            ' Evènements label Accords
            ' ************************
            AddHandler TabTons.Item(i).MouseDown, AddressOf TabTons_MouseDown
            AddHandler TabTons.Item(i).MouseUp, AddressOf TabTons_MouseUp
            '
            DepL = DepL + Longueur + 1
            '
            If j = 7 Then
                DepH = DepH + Hauteur '- 18 '38
                j = 0
                DepL = 0
            End If
            i = i + 1
        Loop Until i > 20
        '
        ' Construction des étiquettes de titres des modes
        ' ***********************************************
        PosX = 10
        DepL = 0
        '
        PosY = 70
        '
        For i = 0 To 2
            TabTonsTitreMode.Add(New Label)
            'Me.TabPage_Tonalité.Controls.Add(TabTonsTitreMode.Item(i))
            Me.Panel8.Controls.Add(TabTonsTitreMode.Item(i))
            '
            ' Titre du mode
            ' *************
            TabTonsTitreMode.Item(i).Size = New System.Drawing.Size(300, 15)
            TabTonsTitreMode.Item(i).BorderStyle = BorderStyle.None
            TabTonsTitreMode.Item(i).Font = New System.Drawing.Font("Papyrus", 9, FontStyle.Regular)
            TabTonsTitreMode.Item(i).Location = New Point((PosX + DepL) - 3, PosY)
            TabTonsTitreMode.Item(i).ForeColor = Color.White
            Select Case i
                Case 0
                    TabTonsTitreMode.Item(i).Text = "Mode Majeur"
                    TabTonsTitreMode.Item(i).Visible = True
                    PosY = 160
                Case 1
                    TabTonsTitreMode.Item(i).Text = "Mode Mineur Harmonique"
                    TabTonsTitreMode.Item(i).Visible = True
                    PosY = 203
                Case 2
                    TabTonsTitreMode.Item(i).Text = "Mode Mineur Mélodique"
                    TabTonsTitreMode.Item(i).Visible = False
            End Select
        Next i
        ' Constructon des étiquettes de degrés
        ' ************************************
        Longueur = 74
        Hauteur = 20
        '
        PosX = 6 '12
        PosY = 38 '3
        '
        DepL = 0
        i = -1
        Do
            i = i + 1
            TabTonsDegrés.Add(New Label)
            'Me.TabPage_Tonalité.Controls.Add(TabTonsDegrés.Item(i))
            Me.Panel8.Controls.Add(TabTonsDegrés.Item(i))
            '
            TabTonsDegrés.Item(i).Size = New System.Drawing.Size(Longueur, Hauteur)
            TabTonsDegrés.Item(i).Font = New System.Drawing.Font("Calibri", 11, FontStyle.Regular)
            TabTonsDegrés.Item(i).Location = New Point(PosX + DepL, PosY)
            TabTonsDegrés.Item(i).BorderStyle = BorderStyle.FixedSingle
            TabTonsDegrés.Item(i).TextAlign = ContentAlignment.MiddleCenter
            TabTonsDegrés.Item(i).Text = "I"
            TabTonsDegrés.Item(i).Tag = Str(i)
            TabTonsDegrés.Item(i).Visible = True
            TabTonsDegrés.Item(i).ForeColor = Color.White
            '
            Select Case i
                Case 0
                    TabTonsDegrés.Item(i).Text = "I"
                Case 1
                    TabTonsDegrés.Item(i).Text = "II"
                Case 2
                    TabTonsDegrés.Item(i).Text = "III"
                Case 3
                    TabTonsDegrés.Item(i).Text = "IV"
                Case 4
                    TabTonsDegrés.Item(i).Text = "V"
                Case 5
                    TabTonsDegrés.Item(i).Text = "VI"
                Case 6
                    TabTonsDegrés.Item(i).Text = "VII"
            End Select
            '
            ' Evènements sur label Degrés
            ' ***************************
            AddHandler TabTonsDegrés.Item(i).MouseDown, AddressOf TabTonsDegrés_MouseDown
            '
            DepL = DepL + Longueur + 2
        Loop Until i >= 6
        '
        '
    End Sub

    Private Sub TabTonsDegrés_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        DerGridCliquée = GridCours.TabTon
    End Sub
    Private Sub TabTonsVoisins_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        Dim com As Label = sender
        Dim ind As Integer
        Dim i As Integer
        Dim tbl() As String
        Dim degré As Integer
        Dim ligne As Integer
        Dim a As String

        ind = Val(com.Tag)
        i = ind '
        LabelCours = ind
        '
        ' Marquage d'un Accord dans le tableau
        ' ************************************
        If e.Button() = Windows.Forms.MouseButtons.Left Then
            MarquerAccordVoisin(ind) ' --> marquer accord 
        End If
        DerGridCliquée = GridCours.TabTon
        '
        ind = Val(com.Tag)
        LabelCours = ind ' utilisé lors du glisser déposer
        '
        VOIS_Maj_TousAccordsMnContext1(ind)
        '
        ' Clic droit : menu contextuel
        ' ***************************
        If e.Button() = MouseButtons.Right Then '  And TabTonsSelect.Item(ind).Checked = True ' My.Computer.Keyboard.ShiftKeyDown And 
            If i <> -1 Then
                'MarquerAccord(ind)
                'TabTons.Item(ind).BackColor
                If Trim(TabTonsVoisins.Item(i).Text) <> "___" Then
                    MarquerAccordVoisin(ind)
                    Accord1.Text = Trim(TabTonsVoisins.Item(i).Text) ' valeur de l'accord dans menu contextuel
                    ContextMenuStrip1.Show(CType(sender, Object), e.Location)
                End If
            End If
        End If

        ' Jouer Accord
        ' ************

        If e.Button() = MouseButtons.Left And My.Computer.Keyboard.CtrlKeyDown Then ' 
            'JouerAccord(i)
            a = Trim(ComboBox1.Text)
            tbl = Split(a)
            Clef = Trim(Det_Clef(tbl(0)))
            '
            Refresh()
            If ComboMidiOut.Items.Count > 0 Then
                accord = Trim(TabTonsVoisins(ind).Text)
                'JouerSourceTabTon2(accord, ind)
                JouerAccord(accord)
            End If
        End If
        '
        ' Pour glisser - déposer
        ' **********************
        If e.Button() = MouseButtons.Left And Not (My.Computer.Keyboard.CtrlKeyDown) And Not (My.Computer.Keyboard.AltKeyDown) _
                                        And Not (My.Computer.Keyboard.ShiftKeyDown) Then

            degré = Det_IndexDegré2(ind)
            ligne = Det_LigneTableGlobale(ind)
            TonVois = TableGlobalAccVoisin(0, ligne, 0)
            tbl = Split(TonVois)
            Select Case Val(com.Tag)' Mise à jour de OrigineAccord
                Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
                    OrigineAccord = Modes.Majeur
                    TonVois = tbl(0) + " Maj"
                    ModeVois = tbl(0) + " Maj"
                Case Else
                    OrigineAccord = Modes.MineurH
                    TonVois = Det_RelativeMajeure2(tbl(0) + " MinH")
                    ModeVois = tbl(0) + " MinH"
            End Select
            OrigineEdition = GridCours.TabTonVoisin
            '
            TabTonsVoisins.Item(ind).DoDragDrop(TabTonsVoisins.Item(ind).Text, DragDropEffects.Copy Or DragDropEffects.Move)
            'a = Trim(TabTonsVoisins.Item(ind).Text) + " " + TonVois
            'TabTonsVoisins.Item(ind).DoDragDrop(Trim(a), DragDropEffects.Copy Or DragDropEffects.Move)
        End If
    End Sub
    Private Sub TabTonsVoisins_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        Dim ind As Integer
        Dim com As Label = sender
        '
        ind = Val(com.Tag)
        '
        If AccordAEtéJoué = True Then
            CouperJouerAccord2()
            AccordAEtéJoué = False
        End If
        '
        RAZ_AffNoteAcc()
    End Sub
    Private Sub TabTons_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) ' , e1 As System.Windows.Forms.MouseEventArgs
        Dim com As Label = sender
        Dim ind As Integer
        Dim i As Integer
        Dim a As String
        Dim tbl() As String
        '
        '
        ind = Val(com.Tag)
        LabelCours = ind
        '
        DerGridCliquée = GridCours.TabTon
        AccordMarqué = ""
        '
        'ActivationDesMenus()
        Maj_TousAccordsMnContext1(ind) ' maj menu contextuel de choix d'accords sur clic droit - tous les accords possibles dans une cellule donnée
        Maj_Octave(ind)
        '
        i = ind '
        '
        ' Marquage d'un Accord dans le tableau
        ' ************************************
        If e.Button() = Windows.Forms.MouseButtons.Left Then
            MarquerAccord(ind) ' --> marquer accord 
        End If
        '
        ' Clic droit : menu contextuel
        ' ***************************
        If e.Button() = MouseButtons.Right Then '  And TabTonsSelect.Item(ind).Checked = True ' My.Computer.Keyboard.ShiftKeyDown And 
            If i <> -1 Then
                'MarquerAccord(ind)
                'TabTons.Item(ind).BackColor
                If Trim(TabTons.Item(i).Text) <> "___" Then
                    MarquerAccord(ind)
                    Accord1.Text = Trim(TabTons.Item(i).Text) ' valeur de l'accord dans menu contextuel
                    ContextMenuStrip1.Show(CType(sender, Object), e.Location)
                End If
            End If
        End If
        '
        ' Jouer Accord
        ' ************
        If e.Button() = MouseButtons.Left And My.Computer.Keyboard.CtrlKeyDown Then ' 
            'JouerAccord(i)
            a = Trim(ComboBox1.Text)
            tbl = Split(a)
            Clef = Trim(Det_Clef(tbl(0)))
            '
            Refresh()
            If ComboMidiOut.Items.Count > 0 Then
                accord = Trim(TabTons(ind).Text)
                JouerAccord(accord)
            End If

        End If
        '
        If e.Button() = MouseButtons.Middle And Not (My.Computer.Keyboard.CtrlKeyDown) And Not (My.Computer.Keyboard.AltKeyDown) _
                                        And Not (My.Computer.Keyboard.ShiftKeyDown) Then
            If Trim(TabTons.Item(i).Text) <> "___" Then
                AfficherAccordSource(Trim(TabTons.Item(i).Text))
            End If
        End If
        ' Pour glisser - déposer
        ' **********************
        If e.Button() = MouseButtons.Left And Not (My.Computer.Keyboard.CtrlKeyDown) And Not (My.Computer.Keyboard.AltKeyDown) _
                                        And Not (My.Computer.Keyboard.ShiftKeyDown) Then
            Select Case Val(com.Tag)' Mise à jour de OrigineAccord
                Case 0, 1, 2, 3, 4, 5, 6
                    OrigineAccord = Modes.Majeur
                Case 7, 8, 9, 10, 11, 12, 13
                    OrigineAccord = Modes.MineurH
                Case 14, 15, 16, 17, 18, 19, 20
                    OrigineAccord = Modes.MineurM
            End Select
            '
            TabTons.Item(ind).DoDragDrop(TabTons.Item(ind).Text, DragDropEffects.Copy Or DragDropEffects.Move)
            If Trim(Valeur_Drag) <> "" Then
                Maj_DragDrop()
            End If
        End If
        '        
        ' Ecrire les notes de l'accord dans la text box
        ' *********************************************
        Dim AA As String = Trim(TabTons.Item(ind).Text)
        Dim tbl1() As String = Trim(ComboBox1.Text).Split()

        Select Case tbl1(0)
            Case "Bb", "Eb", "Ab"
                TextBox3.Text = Det_NotesAccord3(AA, "b")
            Case Else
                TextBox3.Text = Det_NotesAccord3(AA, "#")
        End Select


    End Sub
    Sub JouerAccord(accord As String)

        Dim zone As Integer
        Dim a, b As String
        Dim tbl() As String
        Dim m, n As Integer
        Dim i As Integer
        Dim Tonique As String
        Dim canal As Byte = Convert.ToByte(CanalEcoute.SelectedItem)
        Dim NumAccCours As Integer = Grid2.Selection.FirstCol
        '
        Try
            If AccordAEtéJoué = False Then

                '
                Clef = "#"
                'zone = Det_ZoneAccord(NumAccCours)
                accord = TradD(accord) ' 
                a = Det_NotesAccord(Trim(accord))

                'a = Trad_ListeNotesEnD(a, "-") ' si des notes sont en bémol, elles sont traduites en #
                tbl = Split(a, "-")
                Tonique = Trim(tbl(0))

                ' Détermination des notes à jouer
                If zone = -1 Then
                    b = Maj_NotesCommunes(Trim(a), -1) ' l'utilisation de ces fonctions ne modifie pas la table des voicings : TableNotesAccordsZ
                    b = Maj_Large_BasseMoins1(b, "PisteHorsZone", -1)
                Else
                    b = Maj_NotesCommunes(Trim(a), zone)
                    b = Maj_Large_BasseMoins1(b, "PisteZone", zone)
                End If
                '
                tbl = Split(b)
                '
                If Not (SortieMidi.Item(ChoixSortieMidi).IsOpen) Then
                    SortieMidi.Item(ChoixSortieMidi).Open() ' ici ça plante si un autre instrument est ouvert : plantage avec la version d'essai de syntronik
                End If
                '
                Num_octave = -1
                For i = 0 To UBound(tbl)
                    n = Val(Microsoft.VisualBasic.Right(tbl(i), 1))
                    If n > Num_octave Then
                        Num_octave = n
                    End If

                    m = Val(ListNotesd.IndexOf(Trim(tbl(i)))) 'Det_NumNote(tbl(i))
                    ' *****
                    SortieMidi.Item(ChoixSortieMidi).SendNoteOn(canal - 1, m, VéloEcoute.Value)
                    '
                Next i
                '
                AccordAEtéJoué = True
            End If
            'End If
        Catch ex As Exception

            If LangueIHM = "fr" Then
                MessageHV.PTitre = "Avertissement"
                MessageHV.PContenuMess = "Détection d'une erreur dans procédure : " + "JouerAccord" + "." + Constants.vbCrLf + "Message  : " + ex.Message
                MessageHV.PTypBouton = "OK"
            Else
                MessageHV.PTitre = "Warning"
                MessageHV.PContenuMess = "Détection d'une erreur dans procédure : " + "JouerAccord" + "." + Constants.vbCrLf + "Message  : " + ex.Message
                MessageHV.PTypBouton = "OK"
            End If
            Cacher_FormTransparents()
            MessageHV.ShowDialog()
            End
        End Try


    End Sub
    Function CAD_Det_Etendu(Indexdegré As Integer, N_Note As Integer, Note As Byte) As Byte
        Dim Note1 As Integer
        CAD_Det_Etendu = Note
        If CAD_TableCoursAcc(Indexdegré).EtendreChecked(N_Note) = True Then
            Note1 = Note + 12
            If Note1 > 127 Then
                Note1 = Note
            End If
            CAD_Det_Etendu = Note1
        End If

    End Function
    Function CAD_Det_EtenduNumNote(m As Integer, n As Integer, j As Integer) As Integer
        Dim Note1 As Integer
        CAD_Det_EtenduNumNote = j
        If m <> n Then 'signifie que la note a été étendue
            Note1 = j + 12 ' étendre le n° de note pour l'affichage
            If Note1 > 127 Then
                Note1 = j
            End If
            CAD_Det_EtenduNumNote = Note1
        End If

    End Function
    ' Onglet Tonalité = Tabpage1 | index= 0 | OngletCours = 1
    ' Onglet Progression = TabPage2 | index = 1 |  OngletCours = 0
    ' Onglet TonsVoisins = TabPage7 | index = 2 | OngletCours = 7
    ' Onglet Zones Voicing = Tabpage3 | index = 3 | Onglet Cours = 
    ' Onglet Perso = Tabpage16 | index = 4 | Onglet_Cours = 
    ' Onglet Mix = TabPage9 | index = 5 | Onglet_Cours = 
    '
    Private Sub TabPage1_Paint(sender As Object, e As PaintEventArgs) Handles TabPage1.Paint
        OngletCours = 1 '
    End Sub

    Private Sub TabPage2_Paint(sender As Object, e As PaintEventArgs) Handles TabPage2.Paint
        OngletCours = 0 ' onglet progresson
    End Sub


    Private Sub TabPage3_Paint(sender As Object, e As PaintEventArgs) Handles TabPage3.Paint
        OngletCours = 2
    End Sub

    '
    ' *#######################################################################
    ' ##                                                                    ##
    ' ##                   MENU EDITION / FICHIER                           ##
    ' ##                                                                    ##
    ' ########################################################################
    Sub CouperNote(n As Byte)
        If NoteAEtéJouée = True Then
            NoteAEtéJouée = False
            If NoteCourante <> 255 Then
                If Not (SortieMidi.Item(ChoixSortieMidi).IsOpen) Then
                    SortieMidi.Item(ChoixSortieMidi).Open()
                End If
                'SortieMidi.Item(ChoixSortieMidi).SendNoteOff(CanalThru.Value - 1, n, 0)
                NoteCourante = 255
            End If
            'SortieMidi.Item(ChoixSortieMidi).SilenceAllNotes()
        End If
    End Sub

    Private Sub QuitterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles QuitterToolStripMenuItem.Click
        ' remarque 1 : iI existe une autre version de cette procédure : procédure "Sortir". Cette procédure Sortir est utilisée dans Form1_Closing.
        ' Cette procédure tient compte de la variable  e.cancel de "e as CancelEventArgs". 
        ' 
        Quitter()
    End Sub
    ''' <summary>
    ''' Procédure Quitter : iI existe une autre version de cette procédure : procédure "Sortir". Cette procédure Sortir est utilisée dans Form1_Closing.
    ''' Cette procédure tient compte de la variable  e.cancel de "e as CancelEventArgs". 
    ''' </summary>
    Sub Quitter()
        Dim d As DialogResult

        d = NouveauConfirm2() ' la question est : voulez-vous enregistrez préalablement votre projet
        Select Case d
            Case DialogResult.Yes
                Enregistrer()
                FermerMidi()
                FermetureParQuitter = True
                FermetureEncours = True
                Application.Exit() ' End

            Case DialogResult.No
                FermerMidi()
                FermetureParQuitter = True
                FermetureEncours = True
                Application.Exit() ' End
                End
            Case DialogResult.Cancel ' on ne fait rien

        End Select
    End Sub
    Private Sub Sortir(ByRef e As FormClosingEventArgs)
        ' remarque 1 : la varaible "e as EventArgs" doit être renommé par "e as CancelEventargs" dans Form1_Closing.
        ' remarque 2 : il existe une fonction identique dans le menu "Quitter".

        Dim d As DialogResult

        d = NouveauConfirm2() ' la question est : voulez-vous enregistrez préalablement votre projet
        Select Case d
            Case DialogResult.Yes
                Enregistrer()
                FermerMidi()
                FermetureEncours = True
                e.Cancel = False '  validation de la sortie --> e.Cancel = False
            Case DialogResult.No
                FermerMidi()
                FermetureEncours = True
                e.Cancel = False '  validation de la sortie --> e.Cancel = False
            Case DialogResult.Cancel ' on ne fait rien, on reste dans l'appli, annulation de la sortie --> e.Cancel = True
                e.Cancel = True
        End Select
    End Sub

    Sub CréationRegistry()
        Dim CalquesMIDI As RegistryKey
        'Dim Appli As RegistryKey
        Dim a As String
        Dim j As Integer

        ' Ici, on charge les infos de la base de registre. 
        ' Ces infos sont mises à jour lors de la sortie de l'application dans Form1_FormClosed
        ' 
        ' Création/Ouverture du dossier CalquesMIDI/HyperArp dans la bassse de registre
        ' *****************************************************************************
        Dim DossierDocuments As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        CalquesMIDI = My.Computer.Registry.CurrentUser.CreateSubKey("Software\CalquesMIDI") ' CreateSubKey : Crée une sous-clé ou en ouvre une existante
        '
        ' Langues
        ' *******
        a = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CalquesMIDI\HyperArp\LangueIHM", "LangueIHM", Nothing)
        If a = Nothing Then
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CalquesMIDI\HyperArp\LangueIHM", "LangueIHM", "fr")
            LangueIHM = "fr"
        Else
            LangueIHM = a
            If Trim(a) = "fr" Then
                LangueIHM = "fr"
            Else
                LangueIHM = "en"
            End If
        End If
        ' CheminFichierOuvrir
        ' *******************
        a = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CalquesMIDI\HyperArp\Chemins", "CheminFichierOuvrir", Nothing) ' Nothing est la valeur à retourner si pas de valeur dans le registre
        If a = Nothing Then
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CalquesMIDI\HyperArp\Chemins", "CheminFichierOuvrir", DossierDocuments)
            CheminFichierOuvrir = DossierDocuments
        Else
            CheminFichierOuvrir = a
        End If


        ' Chemin CheminEnregistrer
        ' ************************
        a = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CalquesMIDI\HyperArp\Chemins", "CheminEnregistrer", Nothing) ' Nothing est la valeur à retourner si pas de valeur dans le registre
        If a = Nothing Then
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CalquesMIDI\HyperArp\Chemins", "CheminEnregistrer", DossierDocuments)
            CheminFichierEnreg = DossierDocuments
        Else
            CheminFichierEnreg = a
        End If


        ' Chemin ExportCalquesMIDI
        ' ************************
        a = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CalquesMIDI\HyperArp\Chemins", "ExportCalquesMIDI", Nothing) ' Nothing est la valeur à retourner si pas de valeur dans le registre
        If a = Nothing Then
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CalquesMIDI\HyperArp\Chemins", "ExportCalquesMIDI", DossierDocuments)
            CheminFichierCalques = DossierDocuments
        Else
            CheminFichierCalques = a
        End If

        '
        ' Chemin ExportFichierMIDI (fichier pour  les accords MIDI)
        ' ************************
        a = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CalquesMIDI\HyperArp\Chemins", "ExportFichierMIDI", Nothing) ' Nothing est la valeur à retourner si pas de valeur dans le registre
        If a = Nothing Then
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CalquesMIDI\HyperArp\Chemins", "ExportFichierMIDI", DossierDocuments)
            CheminFichierMIDI = DossierDocuments
        Else
            CheminFichierMIDI = a
        End If


        ' Warning sauvegarde projet courant
        ' ********************************
        'a = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CalquesMIDI\HyperArp\Préférences", "WarningProjetCourant", Nothing) ' Nothing est la valeur à retourner si pas de valeur dans le registre
        'If a = Nothing Then
        'My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CalquesMIDI\HyperArp\Préférences", "WarningProjetCourant", "True")
        'CheckBox1.Checked = True
        'Else
        'CheckBox1.Checked = Convert.ToBoolean(a)
        'End If
        ' Sorties MIDI
        ' ************
        Dim o As Object
        Try
            o = Midi.OutputDevice.InstalledDevices
            '
            If o.count > 0 Then ' o.count donne le nombre de device /l'index max est donc o.count-1
                For j = 0 To (o.count - 1) '  j = 0 To (o.count - 1)
                    SortieMidi.Add(OutputDevice.InstalledDevices(j))
                    ComboMidiOut.Items.Add(SortieMidi.Item(j).Name)
                Next
                '
                a = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CalquesMIDI\HyperArp\Préférences", "ChoixSortieMIDI", Nothing) ' Nothing est la valeur à retourner si pas de valeur dans le registre
                If a = Nothing Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CalquesMIDI\HyperArp\Préférences", "ChoixSortieMIDI", "0")
                    ComboMidiOut.SelectedIndex = 0
                    ChoixSortieMidi = ComboMidiOut.SelectedIndex
                Else
                    If Val(a) > o.count - 1 Then
                        ComboMidiOut.SelectedIndex = 0
                        ChoixSortieMidi = ComboMidiOut.SelectedIndex
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CalquesMIDI\HyperArp\Préférences", "ChoixSortieMIDI", "0")
                    Else
                        ComboMidiOut.SelectedIndex = Val(a)
                        ChoixSortieMidi = ComboMidiOut.SelectedIndex
                    End If
                End If
            Else
                ComboMidiOut.Enabled = False
                Exist_MIDIout = False
                If LangueIHM = "fr" Then
                    MessageHV.PTitre = "Avertissement"
                    MessageHV.PContenuMess = "Vous n'avez pas de sorties MIDI sur votre  PC : l'application ne peut pas fonctionner. - " + "CréationRegistry. - Veuillez installer Microsoft GS Wavetable Synth."
                    MessageHV.PTypBouton = "OK"
                Else
                    MessageHV.PTitre = "Warning"
                    MessageHV.PContenuMess = "There is no MIDI Output in your PC : the software can't work" + "CréationRegistry. - Please install Microsoft GS Wavetable Synth."
                    MessageHV.PTypBouton = "OK"
                End If
                Cacher_FormTransparents()
                MessageHV.ShowDialog()
            End If

            '
            ' Recherche gammes PRO
            ' ********************
            'a = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CalquesMIDI\HyperArp\Préférences", "GammesPRO", Nothing)
            'If a = Nothing Then
            'My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CalquesMIDI\HyperArp\Préférences", "GammesPRO", "0")
            'Button25.Image = HyperArp.My.Resources.Resources._32x32_Transp_vert__échelle3_
            'GammesPRO.Checked = False
            'Button25.ImageAlign = ContentAlignment.MiddleCenter
            'Button25.BackColor = Color.Beige
            'If Module1.LangueIHM = "fr" Then
            'GammesPRO.Text = "Passer en Gammes PRO"
            'Else
            'GammesPRO.Text = "Switch to PRO scales"
            'End If
            'GammesPRO.Image = HyperArp.My.Resources.Resources._32x32_Transp_vert__échelle3_
            'Else
            'If Trim(a) = "1" Then
            'Button25.Image = HyperArp.My.Resources.Resources._32x32___échelle2_Blues
            'Button25.ImageAlign = ContentAlignment.MiddleRight
            'Button25.BackColor = Color.DarkOliveGreen
            'GammesPRO.Checked = True
            'If Module1.LangueIHM = "fr" Then
            'GammesPRO.Text = "Passer en Gammes simples"
            'Else
            'GammesPRO.Text = "Switch to Simple scales"
            'End If
            'GammesPRO.Image = HyperArp.My.Resources.Resources._32x32___échelle2_Blues
            '
            'Else
            'Button25.Image = HyperArp.My.Resources.Resources._32x32_Transp_vert__échelle3_
            'Button25.ImageAlign = ContentAlignment.MiddleCenter
            'Button25.BackColor = Color.Beige
            'GammesPRO.Checked = False
            'If Module1.LangueIHM = "fr" Then
            'GammesPRO.Text = "Passer en Gammes PRO"
            'Else
            'GammesPRO.Text = "Switch to PRO scales"
            'End If
            'GammesPRO.Image = HyperArp.My.Resources.Resources._32x32_Transp_vert__échelle3_
            'End If
            'End If

        Catch ex As Exception
            If LangueIHM = "fr" Then
                MessageHV.PTitre = "Avertissement"
                MessageHV.PContenuMess = "Détection d'un erreur dans procédure : " + "CréationRegistry" + Constants.vbCrLf + "Message  : " + ex.Message
                MessageHV.PTypBouton = "OK"
            Else
                MessageHV.PTitre = "Warning"
                MessageHV.PContenuMess = "Détection d'un erreur dans procédure : " + "CréationRegistry" + Constants.vbCrLf + "Message  : " + ex.Message
                MessageHV.PTypBouton = "OK"
            End If
            Cacher_FormTransparents()
            MessageHV.ShowDialog()
            End
        End Try
        ' 
    End Sub
    Private Sub OuvrirToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OuvrirToolStripMenuItem.Click
        Dim i As Integer

        'If Module1.EcritUneFois = True Then
        Cacher_FormTransparents()
        i = NouveauConfirm() ' la question est : voulez-vous enregistrez préalablement votre projet actuel
        Select Case i
            Case DialogResult.Yes
                Enregistrer()
                'NouveauProjet()
                Ouvrir2()
            Case DialogResult.No
                'NouveauProjet()
                Ouvrir2()
            Case DialogResult.Cancel ' on ne fait rien
        End Select
        'Else
        'Cacher_FormTransparents()
        'Ouvrir2()
        'End If

    End Sub
    Sub Ouvrir()
        Dim i, j, k As Integer
        Dim m As Integer
        Dim ind As Integer = -1
        Dim Line As String
        Dim TBL() As String
        Dim TBL1() As String

        Dim FInfo As FileInfo
        '
        EnChargement = True
        Grid2.AutoRedraw = False
        Try
            OpenFileDialog1.FilterIndex = 2
            '
            OpenFileDialog1.InitialDirectory = CheminFichierOuvrir
            OpenFileDialog1.FileName = ""
            OpenFileDialog1.Filter = "HyperArp Files (*.zic4)|*.zic4"
            'OpenFileDialog1.ShowDialog()
            '
            FichierOuvrir = ""
            If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                FichierOuvrir = OpenFileDialog1.FileName

                CheminFichierOuvrir = OpenFileDialog1.FileName 'OpenFileDialog1.InitialDirectory
                NomFichier = Dir(FichierOuvrir)
                If My.Computer.FileSystem.FileExists(FichierOuvrir) Then
                    '
                    FInfo = My.Computer.FileSystem.GetFileInfo(FichierOuvrir)
                    CheminFichierOuvrir = FInfo.DirectoryName
                    '
                    CheminFichierEnreg = CheminFichierOuvrir
                    FichierEnreg = FichierOuvrir
                    '
                    IndicateurEnreg = True
                    '
                    Me.Text = NomFichier
                    '
                    NouvPourOuvrir()
                    FileOpen(1, FichierOuvrir, OpenMode.Input) ' Ouvre en lecture
                    While Not EOF(1) ' Boucler jusqu'à la fin du fichier
                        '
                        Line = ReadLine() ' Lire chaque ligne
                        TBL = Line.Split(",")
                        '
                        Select Case Trim(TBL(0))
                            Case "ListSyst"                               '
                                For i = 0 To UBound(TBL)
                                    TBL1 = TBL(i).Split()
                                    Select Case TBL1(0)
                                        Case "ListSyst" ' ListSyst est l'entête de la ligne  et n'a pas de valeur
                                        Case "Tempo"
                                            Tempo.Value = Convert.ToInt16(TBL1(1))
                                        Case "Métrique" ' on passe ici car la métrique (nécessaire pour l'import dans  HyperVoicing mais pas pour Ouvrir() dans HyperArp
                                        Case "Début", "Debut"
                                            Début.Value = Convert.ToInt16(TBL1(1))
                                        Case "Fin"
                                            Terme.Value = Convert.ToInt16(TBL1(1))
                                        Case "Répéter", "Repeter"
                                            Répéter.Checked = Convert.ToBoolean(TBL1(1))
                                        Case "TopRow"
                                            listPIANOROLL(0).PTopRow = Convert.ToInt32(TBL1(1))
                                        Case "Accent"
                                            If Convert.ToBoolean(TBL1(1)) Then
                                                Accent1_3 = True
                                                RadioButton4.Checked = True
                                                RadioButton3.Checked = False
                                            Else
                                                Accent1_3 = False
                                                RadioButton4.Checked = False
                                                RadioButton3.Checked = True
                                            End If
                                    End Select
                                Next
                            Case "ListMarq"
                                For i = 1 To UBound(TBL)
                                    Grid2.Cell(1, i).Text = TBL(i)
                                Next i
                            Case "ListChiffAcc"
                                For i = 1 To UBound(TBL)
                                    Grid2.Cell(2, i).Text = TBL(i)
                                Next i
                            Case "ListRépet", "ListRepet"
                                For i = 1 To UBound(TBL)
                                    Grid2.Cell(3, i).Text = TBL(i)
                                Next i
                            Case "ListMagnétos", "ListMagnetos"
                                For i = 1 To UBound(TBL)
                                    ChoixVariationGrid2(i, i, TBL(i) + 4)
                                Next i
                            Case "ListMute"
                                For i = 1 To UBound(TBL)
                                    PisteMute(i - 1).Checked = Convert.ToBoolean(TBL(i))
                                Next i
                            Case "ListVolume"
                                For i = 1 To UBound(TBL)
                                    'PisteVolume.Item(i - 1).Value = Convert.ToByte(TBL(i))
                                Next i
                            Case "ListPRG"
                                For i = 1 To UBound(TBL)
                                    PistePRG.Item(i - 1).SelectedIndex = Convert.ToInt16(TBL(i))
                                Next i
                            Case "ListMotifs"
                                For i = 1 To UBound(TBL)
                                    PisteMotif.Item(i - 1).SelectedIndex = Convert.ToInt16(TBL(i))
                                Next i
                            Case "ListDurées", "ListDurees"
                                For i = 1 To UBound(TBL)
                                    PisteDurée.Item(i - 1).SelectedIndex = Convert.ToInt16(TBL(i))
                                Next i
                            Case "ListOctaves"
                                For i = 1 To UBound(TBL)
                                    PisteOctave.Item(i - 1).SelectedIndex = Convert.ToInt16(TBL(i))
                                Next i
                            Case "ListDyn"
                                For i = 1 To UBound(TBL)
                                    PisteDyn.Item(i - 1).SelectedIndex = Convert.ToInt16(TBL(i))
                                Next i
                            Case "ListSouches"
                                For i = 1 To UBound(TBL)
                                    PisteSouche.Item(i - 1).SelectedIndex = Convert.ToInt16(TBL(i))
                                Next i
                            Case "ListRetards"
                                For i = 1 To UBound(TBL)
                                    PisteRetard.Item(i - 1).SelectedIndex = Convert.ToInt16(TBL(i))
                                Next i
                            Case "ListDelay" ' à supprimer
                                For i = 1 To UBound(TBL)
                                    PisteDelay.Item(i - 1).Checked = Convert.ToBoolean(TBL(i))
                                Next i
                            Case "ListDébutSouche" ' à supprimer
                                For i = 1 To UBound(TBL)
                                    PisteDébut.Item(i - 1).Checked = Convert.ToBoolean(TBL(i))
                                Next i
                            Case "ListRadio1"
                                For i = 1 To UBound(TBL)
                                    PisteRadio1.Item(i - 1).Checked = Convert.ToBoolean(TBL(i))
                                Next i
                            Case "ListRadio2"
                                For i = 1 To UBound(TBL)
                                    PisteRadio2.Item(i - 1).Checked = Convert.ToBoolean(TBL(i))
                                Next i
                            Case "ListRadio3"
                                For i = 1 To UBound(TBL)
                                    PisteRadio3.Item(i - 1).Checked = Convert.ToBoolean(TBL(i))
                                Next i
                            Case "ListSpecific"
                                k = Val(TBL(1))
                                If UBound(TBL) >= 2 Then MAJ_Specific(Line)
                            Case "EventH"
                                For i = 1 To UBound(TBL)
                                    m = Val(TBL(1))
                                    With TableEventH(m, 1, 1)
                                        .Ligne = Val(TBL(2))
                                        .Position = TBL(3)
                                        .Marqueur = TBL(4)
                                        .Tonalité = TBL(5)
                                        .Accord = TBL(6)
                                        .Mode = TBL(7)
                                        .Gamme = .Mode
                                        .Degré = Val(TBL(8))
                                        .NumAcc = Val(TBL(9))
                                        .NumMagnéto = Val(TBL(10))
                                        .Répet = Convert.ToByte((TBL(11)))
                                        '
                                        ' Maj couleurs dans Grid2

                                    End With
                                Next i
                            Case "Zone"
                                For i = 1 To UBound(TBL)
                                    j = Val(TBL(1))
                                    With TZone(j)
                                        .DébutZ = TBL(2)
                                        .TermeZ = TBL(3)
                                        .NoteRacine = Val(TBL(4))
                                        .Racine = TBL(5)
                                        .OctaveRacine = Val(TBL(6))
                                        .OctavePlus1 = Val(TBL(7))
                                        .OctaveMoins1 = Val(TBL(8))
                                        .VoixAsso_OctaveMoins1 = TBL(9)
                                        .VoixAsso_OctavePlus1 = TBL(10)
                                        .ComboVoixAInd = Val(TBL(11))
                                        .ComboVoixBInd = Val(TBL(12))
                                        .CombiVoixInd = Val(TBL(12))
                                    End With
                                Next i
                            Case "Notes"
                                listPIANOROLL(0).Charger_Notes(Line)

                        End Select
                        ReDim TBL(0)
                    End While
                    FileClose(1) ' Fermer.
                End If
            End If
            Me.Text = FichierOuvrir

            ' Mise à jour des Grilles
            Maj_Grilles()
            Maj_Répétition()
            EnChargement = False
            IndicateurEnreg = True
            '
            Grid2.AutoRedraw = True
            Grid2.Refresh()
            '
            ' Mise à jour PianoRoll            
            '
            For i = 0 To nb_PianoRoll - 1
                listPIANOROLL(i).PListAcc = Det_ListAcc()
                listPIANOROLL(i).PListGam = Det_ListGam()
                listPIANOROLL(i).F1_Refresh()
                listPIANOROLL(i).Maj_CalquesMIDI()
            Next
            '
            Automation1.PListAcc = Det_ListAcc()
            Automation1.PListMarq = Det_ListMarq()
            Automation1.F4_Refresh()
            '
            Drums.PListAcc = Det_ListAcc()
            Drums.PListMarq = Det_ListMarq()
            Drums.F2_Refresh()
            '
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        Catch ex As Exception
            If LangueIHM = "fr" Then
                MessageHV.PTitre = "Avertissement"
                MessageHV.PContenuMess = "Erreur interne : procédure Ouvrir : " + ex.Message
                MessageHV.PTypBouton = "OK"
            Else
                MessageHV.PTitre = "Warning"
                MessageHV.PContenuMess = "Internal Error : procédure Ouvrir : " + ex.Message
                MessageHV.PTypBouton = "OK"
            End If
            Cacher_FormTransparents()
            MessageHV.ShowDialog()
            End
        End Try
        '
    End Sub
    Sub Ouvrir2()
        Dim i, j, k, l As Integer
        Dim a, b, c, d As String
        Dim m, n As Integer
        Dim ind As Integer = -1
        Dim Line As String
        Dim TBL() As String
        Dim TBL1() As String
        Dim TBL2() As String
        Dim Tonal As Integer = 0
        Dim FInfo As FileInfo
        '
        EnChargement = False
        Cursor.Current = Cursors.WaitCursor
        '
        Try
            If Trim(FichierEntréSurClic) = "" Then
                OpenFileDialog1.FilterIndex = 2
                OpenFileDialog1.InitialDirectory = CheminFichierOuvrir
                OpenFileDialog1.FileName = ""
                OpenFileDialog1.Filter = "HyperArp Files (*.zic4)|*.zic4"

                If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                    FichierOuvrir = OpenFileDialog1.FileName
                    CheminFichierOuvrir = OpenFileDialog1.FileName 'OpenFileDialog1.InitialDirectory
                    NomFichier = Dir(FichierOuvrir)
                    If My.Computer.FileSystem.FileExists(FichierOuvrir) Then
                        FInfo = My.Computer.FileSystem.GetFileInfo(FichierOuvrir)
                        CheminFichierOuvrir = FInfo.DirectoryName
                        CheminFichierEnreg = CheminFichierOuvrir
                        FichierEnreg = FichierOuvrir
                        'IndicateurEnreg = True
                        Me.Text = NomFichier
                    Else
                        FichierOuvrir = ""
                    End If
                Else
                    FichierOuvrir = ""
                End If
            Else
                FichierOuvrir = Trim(FichierEntréSurClic)
                FichierEntréSurClic = ""
            End If
            '
            If Trim(FichierOuvrir) <> "" Then
                '
                '
                ' Construction de l'appli pour le nouveau fichier
                ' ************************************************
                NouvPourOuvrir()
                JaugeInit()
                '
                'SplitContainer7.Visible = False
                Using sr As StreamReader = New StreamReader(FichierOuvrir)
                    While sr.Peek() >= 0 ' Boucler jusqu'à la fin du fichier
                        '
                        Line = sr.ReadLine() ' Lire chaque ligne
                        TBL = Line.Split(",")
                        '
                        Select Case Trim(TBL(0))
                            ' ADMIN SYSTEME
                            ' *************
                            Case "ListSyst"
                                JaugeProgres(1) '
                                For i = 0 To UBound(TBL) ' 9
                                    TBL1 = TBL(i).Split()
                                    Select Case TBL1(0)
                                        Case "ListSyst" ' ListSyst est l'entête de la ligne  et n'a pas de valeur
                                        Case "Tempo"
                                            Tempo.Value = Convert.ToInt16(TBL1(1))
                                        Case "Métrique", "Metrique" ' on passe ici car la métrique (nécessaire pour l'import dans  HyperVoicing mais pas pour Ouvrir() dans HyperArp
                                        Case "Début", "Debut"
                                            Début.Value = Convert.ToInt16(TBL1(1))
                                        Case "Fin"
                                            Terme.Value = Convert.ToInt16(TBL1(1))
                                        Case "Répéter", "Repeter"
                                            Répéter.Checked = Convert.ToBoolean(TBL1(1))
                                        Case "TopRow"
                                            listPIANOROLL(0).PTopRow = Convert.ToInt32(TBL1(1))
                                        Case "Accent"
                                            If Convert.ToBoolean(TBL1(1)) Then
                                                Accent1_3 = True
                                                RadioButton4.Checked = True
                                                RadioButton3.Checked = False
                                            Else
                                                Accent1_3 = False
                                                RadioButton4.Checked = False
                                                RadioButton3.Checked = True
                                            End If
                                        Case "Largeur_ZoomGrid2"
                                            Largeur_ZoomGrid2 = Convert.ToInt32(TBL1(1))
                                        Case "Tonalité", "Tonalite"
                                            Tonal = Convert.ToInt16(TBL1(1)) ' selectedindex de comobox1 qui est mis à jour à la fin de cette procédure
                                        Case "RacineCours"
                                            TRacine.SelectedIndex = Convert.ToInt16(TBL1(1))
                                        Case "FiltreUnissons"
                                            FiltreUni.Checked = Convert.ToBoolean(TBL1(1))
                                        Case "4Notes"
                                            Fournotes.Checked = Convert.ToBoolean(TBL1(1))
                                        Case "V4notes"
                                            ComboBox11.SelectedIndex = Convert.ToInt16(TBL1(1))
                                        Case "Bassemoins1"
                                            Bassemoins1.Checked = Convert.ToBoolean(TBL1(1))
                                        Case "VBassemoins1"
                                            ComboBox12.SelectedIndex = Convert.ToInt16(TBL1(1))
                                    End Select
                                Next
                                ' TIME LINE (liste horizontale des accords)
                                ' *****************************************
                            Case "ListMarq"
                                JaugeProgres(1)
                                'For i = 1 To UBound(TBL) ' 60
                                For i = 1 To Module1.nbMesures
                                    If i <= UBound(TBL) Then
                                        Grid2.Cell(1, i).Text = TBL(i)
                                    End If
                                Next i
                            Case "ListChiffAcc"
                                JaugeProgres(1)
                                'For i = 1 To UBound(TBL) ' 60
                                For i = 1 To Module1.nbMesures
                                    If i <= UBound(TBL) Then
                                        Grid2.Cell(2, i).Text = TBL(i)
                                    End If
                                Next i
                            Case "ListRépet", "ListRepet"
                                JaugeProgres(1)
                                'For i = 1 To UBound(TBL) ' 60
                                For i = 1 To Module1.nbMesures
                                    If i <= UBound(TBL) Then
                                        Grid2.Cell(3, i).Text = TBL(i)
                                    End If
                                Next i
                            Case "ListMagnétos", "ListMagnetos"
                                JaugeProgres(1)
                                'For i = 1 To UBound(TBL) ' 60
                                For i = 1 To Module1.nbMesures
                                    If i <= UBound(TBL) Then
                                        ChoixVariationGrid2(i, i, TBL(i) + 4)
                                    End If
                                Next i
                            Case "ListRacines"
                                For i = 1 To Module1.nbMesures
                                    If i <= UBound(TBL) Then
                                        Grid2.Cell(12, i).Text = TBL(i)
                                        TableEventH(i, 1, 1).Racine = Trim(TBL(i))
                                    End If
                                Next i

                                ' VOLUMES DE LA TABLE DE MIXAGE
                                ' *****************************
                            Case "ListVolume"
                                JaugeProgres(1)
                                l = nb_PistesVar + nb_PianoRoll + nb_DrumEdit
                                For i = 1 To l 'UBound(TBL)
                                    Me.Mix.PisteVolume.Item(i - 1).Value = Convert.ToByte(TBL(i))
                                    Me.Mix.labelAff.Item(i - 1).Text = Convert.ToString(Me.Mix.PisteVolume.Item(i - 1).Value)
                                Next i

                                ' MIX ACTIVATION
                                ' **************
                            Case "AutorisVolumes"
                                JaugeProgres(1)
                                Mix.Maj_AutorisVolumes(Line)

                            Case "AutresVol"
                                JaugeProgres(1)
                                Mix.Maj_AutresVolumes(Line)
                                Mix.Maj_Barr()

                                ' BLOCS DES VARIATIONS
                                ' ********************
                            Case "ListMute"
                                JaugeProgres(1)
                                For i = 1 To UBound(TBL) ' 42
                                    SelBloc(i - 1).Checked = Convert.ToBoolean(TBL(i))
                                Next i
                            Case "ListPRG"
                                JaugeProgres(1)
                                For i = 1 To UBound(TBL) ' 42
                                    PistePRG.Item(i - 1).SelectedIndex = Convert.ToInt16(TBL(i))
                                Next i
                            Case "ListMotifs"
                                JaugeProgres(1)
                                For i = 1 To UBound(TBL) '42
                                    n = PisteMotif.Item(i - 1).Items.Count ' pourt info, n n'est utilisé ensuite
                                    PisteMotif.Item(i - 1).SelectedIndex = Convert.ToInt16(TBL(i)) ' cette liste est à conserver pour la compatibilité avec les anciennes versions
                                    BoutMotif.Item(i - 1).Text = Trim(PisteMotif.Item(i - 1).Text) ' pour le momment le bouton du formulaire est invisible (à voir s'il faut le remettre)
                                Next i
                            Case "ListMotifs_text" ' ListMotifs_text doit toujours passer après "ListMotifs"
                                JaugeProgres(1)
                                For i = 1 To UBound(TBL) '42
                                    BoutMotif.Item(i - 1).Text = Trim(TBL(i))
                                Next i
                            Case "ListDurées", "ListDurees"
                                JaugeProgres(1)
                                For i = 1 To UBound(TBL) '42
                                    PisteDurée.Item(i - 1).SelectedIndex = Convert.ToInt16(TBL(i))
                                Next i
                            Case "ListOctaves"
                                JaugeProgres(1)
                                For i = 1 To UBound(TBL) ' 42
                                    PisteOctave.Item(i - 1).SelectedIndex = Convert.ToInt16(TBL(i))
                                Next i
                            Case "ListDyn"
                                For i = 1 To UBound(TBL) ' 42
                                    PisteDyn.Item(i - 1).SelectedIndex = Convert.ToInt16(TBL(i))
                                Next i
                            Case "ListDyn2"
                                For i = 1 To UBound(TBL) '42
                                    PisteDyn2.Item(i - 1).Value = Convert.ToDecimal(TBL(i))
                                    'PisteDyn2.Item(i - 1).Value = Convert.ToDecimal(Det_Dyn2(PisteDyn.Item(i - 1).Text))
                                Next i
                            Case "ListSouches"
                                JaugeProgres(1)
                                For i = 1 To UBound(TBL) '42
                                    PisteSouche.Item(i - 1).SelectedIndex = Convert.ToInt16(TBL(i))
                                Next i
                            Case "ListRetards"
                                For i = 1 To UBound(TBL) ' 42
                                    PisteRetard.Item(i - 1).SelectedIndex = Convert.ToInt16(TBL(i))
                                Next i
                            Case "ListDelay" ' à supprimer
                                JaugeProgres(1)
                                For i = 1 To UBound(TBL) ' 42
                                    PisteDelay.Item(i - 1).Checked = Convert.ToBoolean(TBL(i))
                                Next i
                            Case "ListDébutSouche", "ListDebutSouche" ' à supprimer
                                JaugeProgres(1)
                                For i = 1 To UBound(TBL) ' 42
                                    PisteDébut.Item(i - 1).Checked = Convert.ToBoolean(TBL(i))
                                Next i
                            Case "ListRadio1"
                                JaugeProgres(1)
                                For i = 1 To UBound(TBL) '42
                                    PisteRadio1.Item(i - 1).Checked = Convert.ToBoolean(TBL(i))
                                Next i
                            Case "ListRadio2"
                                For i = 1 To UBound(TBL) '42
                                    PisteRadio2.Item(i - 1).Checked = Convert.ToBoolean(TBL(i))
                                Next i
                            Case "ListRadio3"
                                JaugeProgres(1)
                                For i = 1 To UBound(TBL) '42
                                    PisteRadio3.Item(i - 1).Checked = Convert.ToBoolean(TBL(i))
                                Next i
                            Case "ListAccent"
                                JaugeProgres(1)
                                For i = 1 To UBound(TBL) '42
                                    PisteAccent.Item(i - 1).SelectedIndex = Convert.ToInt16(TBL(i))
                                Next i
                            Case "ListNomduSon" '
                                JaugeProgres(1)
                                For i = 1 To UBound(TBL) ' 42
                                    NomduSon.Item(i - 1).Text = TBL(i) ' sur écriture de NomduSon dans HyperArp l'évènement NomduSon_TextChanged est appelé et met à jour le NomduSon dans la table de mixage
                                Next
                                ' MOTIFS PERSO
                                ' ************
                            Case "ListSpecific"
                                k = Val(TBL(1))
                                If UBound(TBL) >= 2 Then MAJ_Specific(Line) ' 60
                                ' TABLE DES EVENEMENTS (TableEventH)
                                ' **********************************
                            Case "EventH"
                                JaugeProgres(1)
                                For i = 1 To UBound(TBL) ' 12
                                    m = Val(TBL(1))
                                    For j = 2 To UBound(TBL) ' 
                                        With TableEventH(m, 1, 1)
                                            Select Case j
                                                Case 2
                                                    .Ligne = Val(TBL(2))
                                                Case 3
                                                    .Position = TBL(3)
                                                Case 4
                                                    .Marqueur = TBL(4)
                                                Case 5
                                                    .Tonalité = TBL(5)
                                                Case 6
                                                    .Accord = TBL(6)
                                                Case 7
                                                    .Mode = TBL(7)
                                                    .Gamme = .Mode ' utile pour les fichiers zic4 qui n'ont pas la mise à our des gammes mais normalement c'est touours fait sur Case 12
                                                Case 8
                                                    .Degré = Val(TBL(8))
                                                Case 9
                                                    .NumAcc = Val(TBL(9))
                                                Case 10
                                                    .NumMagnéto = Val(TBL(10))
                                                Case 11
                                                    .Répet = Convert.ToByte((TBL(11)))
                                                Case 12
                                                    TBL2 = TBL(12).Split()
                                                    If Trim(TBL2(1)) = "Blues" Then TBL(12) = Trim(TBL2(0) + " " + "Blues1") ' * RUSTINE : pour les anciens fichiers zic4 qui n'avaient BLues1 mais Blues (Blues1= blues).
                                                    .Gamme = TBL(12)
                                            End Select
                                        End With
                                        '
                                    Next j
                                Next i
                                ' ZONES
                                ' *****
                            Case "Zone"
                                For i = 1 To UBound(TBL) ' 6
                                    TBL1 = TBL(i).Split()
                                    Select Case TBL1(0)
                                        Case "Num"
                                            j = Convert.ToInt32(TBL1(1))
                                        Case "DébutZ", "DebutZ"
                                            TZone(j).DébutZ = Trim(TBL1(1))
                                            If j = 0 Then
                                                TZone(0).DébutZ = "1" ' La 1ere zone commence toujours à 1
                                            End If
                                        Case "TermeZ"
                                            If Convert.ToInt16(TBL1(1)) <= Module1.nbMesures Then
                                                TZone(j).TermeZ = Trim(TBL1(1))
                                            Else
                                                TZone(j).TermeZ = Module1.nbMesures.ToString
                                            End If
                                            If j = 0 Then
                                                TZone(0).TermeZ = Module1.nbMesures.ToString ' La 1ere zone finit toujours au nombre max de mesures
                                            End If
                                        Case "NoteRacine"
                                            TZone(j).NoteRacine = Convert.ToInt32(TBL1(1))
                                        Case "Racine"
                                            TZone(j).Racine = Trim(TBL1(1))
                                        Case "OctaveMoins1"
                                            TZone(j).OctaveMoins1 = Convert.ToBoolean(TBL1(1))
                                        Case "Voix"
                                            TZone(j).VoixAsso_OctaveMoins1 = Trim(TBL1(1))
                                    End Select
                                Next i

                             ' PIANOROLL
                             ' *********
                            Case "PianoRoll"
                                JaugeProgres(1)
                                i = Convert.ToInt16(TBL(1)) - 1
                                If i <= nb_PianoRoll - 1 Then
                                    Select Case Trim(TBL(2))
                                        Case "ParamMélo", "ParamMelo"
                                            listPIANOROLL(i).Charger_Param(Line)
                                        Case "NotesMélo", "NotesMelo"
                                            listPIANOROLL(i).Charger_Notes(Line)
                                        Case "Control"
                                            listPIANOROLL(i).Charger_Ctrl(Line)
                                        Case "ControlSys"
                                            listPIANOROLL(i).Charger_ControlSys(Line)
                                        Case "Pédale" ' 
                                            listPIANOROLL(i).Charger_CtrlPédale(Line)
                                        Case "CalquesMIDI"
                                            listPIANOROLL(i).Charger_CalquesMIDI(Line)
                                        Case "ParamCalquesMIDI"
                                            listPIANOROLL(i).Charger_ParamCalquesMIDI(Line)
                                        Case "Assist1CTRP"
                                            listPIANOROLL(i).Charger_Assist1CTRP(Line)
                                    End Select
                                End If
                                '
                                ' AUTOMATION
                                ' **********
                            Case "Automation"
                                i = Convert.ToInt16(TBL(1)) - 1
                                Select Case Trim(TBL(2))
                                    Case "Control"
                                        Automation1.Charger_Ctrl(Line)
                                    Case "ControlSys"
                                        Automation1.Charger_ControlSys(Line)
                                End Select
                                '
                                ' BATTERIE
                                ' *******
                            Case "ListDrumInst"
                                JaugeProgres(1)
                                For i = 1 To TBL.Count - 1 ' 13
                                    Drums.Charger_ListDrumInst(i, TBL(i))
                                Next
                            Case "ListDrumNotes"
                                Drums.Charger_ListDrumNotes(Line)
                            Case "ListTimeLPres"
                                Drums.Charger_ListTimeLPres(Line)
                            Case "NomPréset"
                                Drums.Charger_LNomPréset(Line)
                                Drums.FocusSurA()
                        End Select
                        ReDim TBL(0)
                    End While
                End Using ' Fermer.

                Me.Text = FichierOuvrir
                ' Mise à jour des Grilles
                Maj_Grilles()
                Maj_Répétition()
                Calcul_AutoVoicingZ()
                EnChargement = False
                IndicateurEnreg = True
                JaugeProgres(10)
                ' Mise à jour entête des PianoRoll 
                ' ********************************
                a = Trim(Det_ListAcc())
                b = Trim(Det_ListGam())
                c = Trim(Det_ListMarq())
                d = Trim(Det_ListTon())
                For i = 0 To nb_PianoRoll - 1
                    JaugeProgres(12)
                    listPIANOROLL(i).PListAcc = a 'Det_ListAcc()
                    listPIANOROLL(i).PListGam = b 'Det_ListGam()
                    listPIANOROLL(i).PListMarq = c 'Det_ListMarq()
                    listPIANOROLL(i).PListTon = d
                    listPIANOROLL(i).F1_Refresh()
                    listPIANOROLL(i).Clear_AllLayers()
                    listPIANOROLL(i).F1_Refresh()
                    listPIANOROLL(i).Maj_CalquesMIDI()
                Next
                ' Mise à jour de drumedit (timeline et drumedit = Grid1 et Grid2)
                '
                Automation1.PListAcc = Det_ListAcc()
                Automation1.PListMarq = Det_ListMarq()
                Automation1.F4_Refresh()
                '
                Drums.PListAcc = Det_ListAcc()
                Drums.PListMarq = Det_ListMarq()
                Drums.F2_Refresh()
                Drums.Refresh_Drums_Ouvrir()
                '
                ' Position splitter central
                SplitContainer7.SplitterDistance = PosSystem
                Panel2.Visible = False ' panneau contenant le gros afficheur des mesures et accords (au milieu)
                Grid2.Width = SplitContainer7.SplitterDistance
                '
                ' position des onglets sur 1er onglet
                ' ***********************************
                TabControl2.SelectedTab = TabControl2.TabPages(0) ' 1er Onglet HyperArp
                TabControl4.SelectedTab = TabControl4.TabPages(0) ' 1er Onglet Généraux
                TabControl5.SelectedTab = TabControl5.TabPages(0) ' 1er Onglet des variations
                '
                SplitContainer7.Visible = True
                Me.Activate() ' permet de mettre en avant l'application
                ' Mis à jour de la tonalité dans combobox1
                ComboBox1.SelectedIndex = Tonal
                Cursor.Current = Cursors.Default
                '
                ' Temps désigné par les locators
                ' ******************************
                TextBox2.Text = Calcul_Durée()

                JaugeFin()
            End If
        Catch ex As Exception
            If LangueIHM = "fr" Then
                MessageHV.PTitre = "Avertissement"
                MessageHV.PContenuMess = "Erreur interne : procédure Ouvrir2 : " + "Traitement de : " + " Message .net : " + ex.Message
                MessageHV.PTypBouton = "OK"
            Else
                MessageHV.PTitre = "Warning"
                MessageHV.PContenuMess = "Internal Error : procédure Ouvrir2 : " + ex.Message
                MessageHV.PTypBouton = "OK"
            End If
            Cacher_FormTransparents()
            MessageHV.ShowDialog()
            End
        End Try
        '
    End Sub
    Sub MAJ_Specific(Line As String)
        Dim i, Ligne, Col As Integer
        Dim TBL() As String
        Dim TBL1() As String
        Dim a As String
        Dim numGrid As Integer

        TBL = Line.Split(",")
        numGrid = TBL(1)
        For i = 2 To UBound(TBL)
            TBL1 = TBL(i).Split(";")
            a = TBL1(0)
            Ligne = TBL1(1)
            Col = TBL1(2)
            If Trim(a) = "*" Then
                MotifsPerso.Item(numGrid).Cell(Ligne, Col).Text = Trim(a)
                MotifsPerso.Item(numGrid).Cell(Ligne, Col).BackColor = Couleur_Notes
            ElseIf Trim(a) = "c" Then
                MotifsPerso.Item(numGrid).Cell(Ligne, Col).BackColor = Couleur_Notes
            End If
        Next i
    End Sub

    Sub MAJ_MUTE(listMUTE As String)
        Dim tbl() As String = listMUTE.Split(";")
        For i = 1 To nb_BlocPistes
            PisteMute.Item(i - 1).Checked = Convert.ToBoolean(tbl(i))
        Next
    End Sub
    Sub MAJ_PRG(listPRG As String)
        Dim tbl() As String = listPRG.Split(";")
        For i = 1 To nb_BlocPistes
            PistePRG.Item(i - 1).SelectedIndex = Convert.ToInt16(tbl(i))
        Next
    End Sub
    Sub MAJ_MOTIFS(listMOTIFS As String)
        Dim tbl() As String = listMOTIFS.Split(";")
        For i = 1 To nb_BlocPistes
            PisteMotif.Item(i - 1).SelectedIndex = Convert.ToInt16(tbl(i))
        Next
    End Sub
    Sub MAJ_OCT(listOCT As String)
        Dim tbl() As String = listOCT.Split(";")
        For i = 1 To nb_BlocPistes
            PisteOctave.Item(i - 1).SelectedIndex = Convert.ToInt16(tbl(i))
        Next
    End Sub
    Sub MAJ_DUREE(listDUREE As String)
        Dim tbl() As String = listDUREE.Split(";")
        For i = 1 To nb_BlocPistes
            PisteDurée.Item(i - 1).SelectedIndex = Convert.ToInt16(tbl(i))
        Next
    End Sub
    Sub MAJ_DYN(listDYN As String)
        Dim tbl() As String = listDYN.Split(";")
        For i = 1 To nb_BlocPistes
            PisteDyn.Item(i - 1).SelectedIndex = Convert.ToInt16(tbl(i))
        Next
    End Sub
    Sub MAJ_PAN(listPAN As String)
        Dim tbl() As String = listPAN.Split(";")
        For i = 1 To nb_BlocPistes
            If tbl(i) = 1 Then PisteRadio1.Item(i - 1).Checked = True 'PisteDyn.Item(i - 1).SelectedIndex = Convert.ToInt16(tbl(i))
            If tbl(i) = 2 Then PisteRadio2.Item(i - 1).Checked = True
            If tbl(i) = 3 Then PisteRadio3.Item(i - 1).Checked = True
        Next
    End Sub


    Sub AffTona()
        Dim tbl() As String
        Dim a As String
        a = "C Maj" 'TableEventH(1, 1, 1).Tonalité
        tbl = Split(Trim(a))
        AffTonaCours(tbl(0))
    End Sub





    Public Function ReadLine() As String
        Dim a As String
        a = LineInput(1) ' Lire chaque ligne

        ReadLine = Mid(a, 2, Len(a) - 2)
    End Function
    Private Sub Ouvrir_NOTES(Positions As String, Notes As String)
        Dim TBL() As String
        Dim m, t, ct As Integer

        TBL = Split(Positions, " ")
        m = Val(TBL(0))
        t = Val(TBL(1))
        ct = Val(TBL(2))
        '
        ' mise à jour
        TableNotesAccords(m, t, ct) = Trim(Notes)

    End Sub




    Private Sub Ouvrir_ONGLET_CADENCES(Intitulé As String, Valeur As String)
        Dim i As Integer
        Dim tbl() As String

        Select Case Intitulé
        'Case "Onglet_Courant"
        '   If Valeur = "TabPage_Cadences" Then
        'TabPage_Cadences.Show()
        '       TabPage_Cadences.Select()
        'TabControl4.SelectedTab = TabPage_Cadences
            '   End If
            Case "Cadences_Majeures"
                ComboBox3.SelectedIndex = Val(Valeur)
            Case "Cadences_Mineures"
                ComboBox4.SelectedIndex = Val(Valeur)
            Case "TypAccord"
                ComboBox6.SelectedIndex = Val(Valeur)
            Case "Accords"
                tbl = Split(Valeur, "-")
                For i = 0 To TabCad.Count - 1
                    TabCad.Item(i).Text = TradAcc_AnglLat(tbl(i))
                    If Trim(tbl(i)) = "" Then
                        TabCadDegrés.Item(i).Visible = False
                        TabCad.Item(i).Visible = False
                    End If
                Next i
                Refresh()

            'Next i
            'Case "Filtres4Visible"
            ' tbl = Split(Valeur, "-")
            'For i = 0 To TabCadFiltres4.Count - 1
            'TabCadFiltres4.Item(i).Visible = False
            'If tbl(i + 1) = "True" Then
            'TabCadFiltres4.Item(i).Visible = True
            'End If
            'Next i
            'Case "Filtres4Checked"
            '   tbl = Split(Valeur, "-")
            '  For i = 0 To TabCadFiltres4.Count - 1
            ' TabCadFiltres4.Item(i).Checked = False
            'If tbl(i + 1) = "True" Then
            'TabCadFiltres4.Item(i).Checked = True
            'End If
            'Next
            'Case "Filtres7Visible"
            '   tbl = Split(Valeur, "-")
            '  For i = 0 To TabCadFiltres7.Count - 1
            ' TabCadFiltres7.Item(i).Visible = False
            'If tbl(i + 1) = "True" Then
            'TabCadFiltres7.Item(i).Visible = True
            'End If
            'Next i
            'Case "Filtres7Checked"
            '   tbl = Split(Valeur, "-")
            '  For i = 0 To TabCadFiltres7.Count - 1
            ' TabCadFiltres7.Item(i).Checked = False
            'If tbl(i + 1) = "True" Then
            'TabCadFiltres7.Item(i).Checked = True
            'End If
            'Next
            Case "indicCadence"
                EnChargement = False
                If Trim(Valeur) = "Maj" Then
                    'OuvrirCadences(Trim(ComboBox3.Text))
                Else
                    'OuvrirCadences(Trim(ComboBox4.Text))
                End If
                EnChargement = True
        End Select
    End Sub

    Private Sub Ouvrir_EVENTH(Ligne As String, Position As String, Marqueur As String, Tonalité As String, Accord As String, Gamme As String, Mode As String, Degré As String, Détails As String, Rép As String)
        Dim tbl() As String
        Dim m, t, ct As Integer

        tbl = Split(Position, ".")
        m = Val(tbl(0))
        t = Val(tbl(1))
        ct = Val(tbl(2))
        '
        TableEventH(m, t, ct).Ligne = Val(Ligne)
        TableEventH(m, t, ct).Position = Trim(Position)
        '
        TableEventH(m, t, ct).Marqueur = Trim(Marqueur)
        TableEventH(m, t, ct).Détails = Trim(Détails)
        TableEventH(m, t, ct).Répet = Convert.ToByte(Rép)

        '
        Grid2.AutoRedraw = True
        TableEventH(m, t, ct).Tonalité = Trim(Tonalité)
        TableEventH(m, t, ct).Accord = Trim(Accord)
        TableEventH(m, t, ct).Gamme = Trim(Gamme)
        TableEventH(m, t, ct).Mode = Trim(Mode)
        TableEventH(m, t, ct).Degré = Val(Degré)
        '
    End Sub


    Private Sub AssurerVisibilitéCelluleGrid2(Col_Active As Integer)
        Dim r As Integer
        Dim c As Integer
        '
        r = Grid2.ActiveCell.Row
        c = Col_Active 'Grid3.ActiveCell.Col
        '
        Grid2.Cell(r, c).EnsureVisible()
    End Sub
    Private Sub NouveauAvecSignatireToolStripMenuItem_Click(sender As Object, e As EventArgs)
        'Signature.ShowDialog()
        If Retour = "OK" Then
            Nouv(RetourSTR)
        End If

    End Sub
    Sub Nouv(RetourSTR1 As String)
        Dim i, j As Integer
        Dim ligne As Integer = 0

        Try


            ' Reinit PianoRoll
            ' ****************
            For i = 0 To nb_PianoRoll - 1
                listPIANOROLL(i).Clear_Notes()
                listPIANOROLL(i).Init_BarrOut()
                listPIANOROLL(i).Clear_AllLayers()
                listPIANOROLL(i).Init_ControlSys()
                listPIANOROLL(i).Clear_Courbes()
                listPIANOROLL(i).Nouv_CalquesMIDI()
                listPIANOROLL(i).Nouv_GridAssist1()
            Next
            '
            Automation1.Clear_Courbes()
            ligne = 3 'ligne + 1
            Automation1.Init_ControlSys()

            ' Reinit des onglets de l'automation
            ' **********************************
            ligne = 4 'ligne + 1
            For i = 0 To nbCourbes - 1
                Automation1.TabCourbes.Item(i).Visible = False
                Automation1.TabCourbes.Item(i).SelectedIndex = 0
                Automation1.TabCourbes.Item(i).Visible = True
            Next i
            '
            ' Réinit DrumEdit
            ' ***************
            Drums.FocusSurA()
            '
            'ligne = 6 'ligne + 1
            '
            ' MIXAGE : Reinit des curseurs de la table de mixage
            ' **************************************************
            For i = 0 To nb_TotalPistes - 1
                Mix.PisteVolume.Item(i).Value = 100
                Mix.muteVolume.Item(i).Checked = True
                Mix.PisteVolume.Item(i).Enabled = True ' nécessaire si soloVolume était à false
            Next i
            '
            ' désativation de la table
            ' ************************
            RemoveHandler Mix.AutorisVol.MouseClick, AddressOf Mix.AutorisVol_MouseClick
            'ligne = 7 'ligne + 1
            Mix.AutorisVol.Checked = False
            AddHandler Mix.AutorisVol.MouseClick, AddressOf Mix.AutorisVol_MouseClick
            'ligne = 8 'ligne + 1
            Mix.VolumesEnabled(False)

            ' désativation des Autres Volumes
            ' *******************************
            RemoveHandler Mix.AutresV.CheckedChanged, AddressOf Mix.AutresV_CheckedChanged
            Mix.AutresV.Checked = False
            Mix.AutresV.Enabled = False
            Mix.Send.Enabled = False
            Mix.VolumesEnabled(False)
            Mix.AutresV_False()
            RemoveHandler Mix.AutresV.CheckedChanged, AddressOf Mix.AutresV_CheckedChanged
            '
            ' Reinit Grid2 : time line des accords avec les variations, marquerus etc..
            ' *************************************************************************
            Grid2.AutoRedraw = False
            '
            Grid2.Range(0, 0, (Grid2.Rows - 1), (Grid2.Cols - 1)).ClearAll()
            Grid2.Range(0, 0, (Grid2.Rows - 1), (Grid2.Cols - 1)).ClearFormat()
            Grid2.Range(0, 0, (Grid2.Rows - 1), (Grid2.Cols - 1)).ClearText()

            '
            'ligne = 10 'ligne + 1
            For i = 0 To Grid2.Rows - 1 ' on place cette boucle car le ClearText semble ne pas fonctionner
                For j = 0 To Grid2.Cols - 1
                    Grid2.Cell(i, j).Text = ""
                Next
            Next
            '
            '
            'ligne = 11 'ligne + 1
            '
            For i = 0 To Grid2.Cols - 1
                ' RAZ 1er ligne de grid2 pour le cas où il y avait des marqeurs
                Grid2.Cell(0, i).BackColor = Color.Beige
                Grid2.Cell(0, i).ForeColor = Color.Black
                '
                Grid2.Cell(0, i).Locked = True
            Next i
            '
            'ligne = 12 'ligne + 1
            '
            For i = 0 To Grid2.Cols - 1
                Grid2.Column(i).Locked = True
            Next i
            '
            Grid2.Range(0, 0, Grid2.Rows - 1, Grid2.Cols - 1).BackColor = Color.White
            Grid2.AutoRedraw = True
            Grid2.Refresh()
            '
            'ligne = 13 'ligne + 1
            '
            Maj_TAccents(Trim(RetourSTR1))
            ' Reconstruction générale de HyperArp
            ' ***********************************
            'ligne = 14 'ligne + 1
            Construction(Trim(RetourSTR1), 2)
            '
            'ligne = 15 'ligne + 1
            '
            ComboBox1.SelectedIndex = 7 ' liste des tonalités majeures
            ComboBox2.SelectedIndex = 7 ' liste des tonalités mineures
            ComboBox23.SelectedIndex = 0
            ComboBox6.SelectedIndex = 0
            '
            ComboBox4.SelectedIndex = 0
            ComboBox3.SelectedIndex = 0
            '
            ' Réinit paramètres de la barre de transport
            ' ******************************************
            '
            '
            Répéter.Checked = False
            '
            ' Réinit de quelques paramètre nécessairees à l'ouverture et à l'enregistrement de fichiers.
            ' *****************************************************************************************
            'ligne = 16 'ligne + 1
            Init_Fichier()
            '
            '
            ' Rinit de la timeline au début
            ' *****************************
            Grid2.LeftCol = 1
            '
            ' Reinit diverses
            ' ***************
            '
            'ligne = 17 'ligne + 1
            AffTona()
            'ligne = 18 'ligne + 1
            INIT_Pistes()
            'ligne = 19 'ligne + 1
            INIT_Specific()
            'ligne = 20 'ligne + 1
            'ligne = 21 'ligne + 1
            Drums.Refresh_Drums_Init()
            '
            'ligne = 22 'ligne + 1
            Maj_DelockCell()
            '
            Grid2.Refresh()
            '
            'ligne = 23 'ligne + 1

            ForçageAccordBase() ' forçage accord de base "C Maj" + Effacement des lignes fixes des piano rolls
            '
            'ligne = 24 'ligne + 1
            VoicingView_Clear() ' effacer les notes dans voicing view
            '
            '
            'ligne = 25 'ligne + 1
            '
            ' Reinit IHM Globale
            ' ******************
            SplitContainer7.SplitterDistance = PosSystem
            Grid2.Width = SplitContainer7.SplitterDistance
            Init_IHM()
            '
            ' Mise de la vue des notes
            ' ************************
            Maj_VueNotes()
        Catch ex As Exception
            '
            If LangueIHM = "fr" Then
                MessageHV.PTitre = "Avertissement"
                MessageHV.PContenuMess = "Erreur interne : procédure 'Nouv' : " + "LIGNE = " + ligne.ToString + "Message : " + ex.Message
                MessageHV.PTypBouton = "OK"
            Else
                MessageHV.PTitre = "Warning"
                MessageHV.PContenuMess = "Erreur interne : procédure 'Nouv' : " + ex.Message
                MessageHV.PTypBouton = "OK"
            End If
            Cacher_FormTransparents()
            MessageHV.ShowDialog()
            End
        End Try
    End Sub
    Sub Init_IHM()
        SplitContainer7.SplitterDistance = PosSystem
        SplitContainer2.SplitterDistance = 550
        SplitContainer2.Panel2.Visible = True
        Panel2.Visible = False
        '
        Largeur_ZoomGrid2 = Grid2Largeur0
        ZoomPlusGrid2()
    End Sub

    Sub NouvPourOuvrir()
        Dim i As Integer
        Dim a As Boolean
        '
        Try
            a = SortieMidi.Item(ChoixSortieMidi).IsOpen
            If a = True Then
                SortieMidi.Item(ChoixSortieMidi).Close() ' fermeture pour re-init de la sortie à la fin de la méthode NouvPourOuvrir sur .Open
            End If
            Grid2.AutoRedraw = False
            '
            '
            For i = 0 To Grid2.Cols - 1
                Grid2.Column(i).Locked = False
            Next i
            '
            '
            Grid2.Range(0, 0, (Grid2.Rows - 1), (Grid2.Cols - 1)).ClearAll()
            '
            Grid2.Range(0, 0, (Grid2.Rows - 1), (Grid2.Cols - 1)).ClearFormat()

            Grid2.Range(0, 0, (Grid2.Rows - 1), (Grid2.Cols - 1)).ClearText()
            '

            '
            For i = 0 To Grid2.Cols - 1
                Grid2.Column(i).Locked = True
            Next i
            'i
            Grid2.Refresh()
            '
            Grid2.AutoRedraw = True
            '
            '
            Grid2.LeftCol = 1
            '
            Nouv("4/4")
            '
            'PianoRoll1.Clear_Notes()
            'PianoRoll1.Init_BarrOut()
            '
            For i = 0 To nb_PianoRoll - 1
                listPIANOROLL(i).Clear_Notes()
                listPIANOROLL(i).Init_BarrOut()
            Next

            a = SortieMidi.Item(ChoixSortieMidi).IsOpen
            If a = False Then
                SortieMidi.Item(ChoixSortieMidi).Open()
            End If

        Catch ex As Exception
            Dim aa As String = SortieMidi.Item(ChoixSortieMidi).Name
            If LangueIHM = "fr" Then
                MessageHV.PTitre = "Avertissement"
                MessageHV.PContenuMess = "Alarme : procédure 'NouvPourOuvrir' : l'interface MIDI pourrait être occupée par une autre application." _
+ vbCrLf + "- choisissez une autre sortie MIDI " + "(" + aa + ")" _
+ vbCrLf + "- ou une autre application MIDI pourrait être présente : libérez cette application ou mettez la en tâche de fond," _
+ vbCrLf + "- ou redémarrez votre PC." _
+ vbCrLf + ex.Message
                MessageHV.PTypBouton = "OK"
            Else
                MessageHV.PTitre = "Warning"
                MessageHV.PContenuMess = "Alarme : procédure 'NouvPourOuvrir' :  MIDI interface could be occupied by another application." _
+ vbCrLf + "- choose another MIDI output," + "(" + aa + ")" _
+ vbCrLf + "- or another MIDI application might be present: release this application or put it in the background," _
+ vbCrLf + "- or reboot your PC." _
+ vbCrLf + ex.Message
                MessageHV.PTypBouton = "OK"
            End If
            Cacher_FormTransparents()
            MessageHV.ShowDialog()
        End Try

    End Sub

    Sub Restit_CoursAcc()
        Dim i, j, k, nbCol As Integer

        nbCol = 7
        If EnChargement = False Then
            For i = 0 To UBound(TableCoursAcc, 1)
                For j = 0 To UBound(TableCoursAcc, 2)
                    k = j + (nbCol * i)
                    TabTons.Item(k).Text = S_TableCoursAcc(i, j).Accord     'TableCoursAcc(i, j)
                Next j
            Next i
            '
            For i = 0 To 4
                TabCad.Item(i).Text = S_CAD_TableCoursAcc(i).Accord
            Next i
        End If
        '
        Refresh()
    End Sub
    Function Det_IndexCadAccCours(Ind As Integer) As Integer
        Dim a As String


        a = TabCadDegrés.Item(Ind).Text
        If Cad_OrigineAccord = Modes.Cadence_Mixte Then
            Select Case Trim(a)
                Case "V", "IV"
                    Det_IndexCadAccCours = Det_IndexDegré(a)
                Case "VI", "III"
                    Det_IndexCadAccCours = Det_IndexDegréMin(a)
                Case Else
                    Det_IndexCadAccCours = -1
            End Select
        Else
            Det_IndexCadAccCours = Det_IndexDegré(a)
        End If
    End Function


    Private Sub TradGrid2_Latin()
        Dim i As Integer
        Dim j As Integer
        Dim a, b As String
        Dim tbl() As String
        '
        If Langue = "fr" Then
            For i = 1 To Grid2.Cols - 1
                a = Trim(Grid2.Cell(1, i).Text)
                If Trim(a) <> "" Then
                    If InStr(a, "/") = 0 Then
                        b = TradAcc_AnglLat(a)
                        Grid2.Cell(1, i).Text = Trim(b)
                    Else
                        tbl = Split(Trim(a), "/")
                        b = ""
                        For j = 0 To UBound(tbl)
                            b = b + TradAcc_AnglLat(tbl(j)) + " / "
                        Next j
                        ' retirer le dernier /
                        ' ********************
                        b = Mid(Trim(b), 1, Len(Trim(b)) - 1)
                        Grid2.Cell(1, i).Text = Trim(b)
                    End If
                End If
            Next i
        End If
    End Sub
    Private Sub TradGrid2_Angl()
        Dim i As Integer
        Dim j As Integer
        Dim a, b As String
        Dim tbl() As String

        If Langue = "en" Then
            For i = 1 To Grid2.Cols - 1
                a = Trim(Grid2.Cell(1, i).Text)
                If Trim(a) <> "" Then
                    If InStr(a, "/") = 0 Then
                        b = TradAcc_LatAngl2(a)
                        Grid2.Cell(1, i).Text = Trim(b)
                    Else
                        tbl = Split(Trim(a), "/")
                        b = ""
                        For j = 0 To UBound(tbl)
                            b = b + TradAcc_LatAngl2(tbl(j)) + " / "
                        Next j
                        ' retirer le dernier /
                        b = Mid(Trim(b), 1, Len(Trim(b)) - 1)
                        Grid2.Cell(1, i).Text = Trim(b)
                    End If
                End If
            Next i
        End If
    End Sub

    Sub Maj_Grilles()
        Dim m, t, ct As Integer
        Dim ligne As Integer
        Dim tbl() As String
        Dim P1 As New Point(1, 100)
        Dim P2 As New Point(1, 100)
        Dim a As String
        Dim tbl1() As String
        '
        'Grid2.Visible = False
        Grid2.AutoRedraw = False
        '
        t = 1
        ct = 1
        For m = 0 To nbMesures '- 1
            ligne = TableEventH(m, t, ct).Ligne
            If ligne = 1 Then
                Entrée_Tonalité = TableEventH(m, t, ct).Tonalité
            End If
            If ligne <> -1 Then
                '
                ' MAJ Marqueurs
                ' *************

                If Trim(TableEventH(m, t, ct).Marqueur) <> "" Then
                    Grid2.AutoRedraw = False
                    Grid2.Cell(1, m).Locked = False
                    Grid2.Cell(1, m).BackColor = Color.AliceBlue
                    Grid2.Cell(1, m).ForeColor = Color.Red
                    Grid2.Cell(1, m).Locked = True
                    Grid2.Cell(1, m).Text = Trim(TableEventH(m, t, ct).Marqueur)
                    Grid2.Cell(1, m).SetFocus()
                    '
                End If
                '
                ' Maj Accords et Gammes
                ' **********************
                a = TableEventH(m, t, ct).Tonalité '
                a = Det_RelativeMajeure2(a) ' si la tonalité est mineure alors on affiche la couleur de la relative majeure
                tbl = Split(a)
                Grid2.Cell(2, m).BackColor = DicoCouleur.Item(Trim(tbl(0))) ' la couleur est fonction de la tonalité
                Grid2.Cell(2, m).ForeColor = DicoCouleurLettre.Item(tbl(0))


                Grid2.Cell(2, m).Alignment = FlexCell.AlignmentEnum.CenterCenter
                Grid2.Cell(11, m).Alignment = FlexCell.AlignmentEnum.CenterCenter
                Grid2.Cell(2, m).Text = TradAcc_AnglLat(Trim(TableEventH(m, t, ct).Accord))
                Grid2.Cell(11, m).Text = TradAcc_AnglLat(Trim(TableEventH(m, t, ct).Gamme))

                '
                ' Maj Couleurs Accords et Gammes 
                ' ******************************
                Grid2.AutoRedraw = False
                a = TableEventH(m, t, ct).Tonalité '
                a = Det_RelativeMajeure2(a) ' si la tonalité est mineure alors on affiche la couleur de la relative majeure
                tbl1 = Split(a)
                Grid2.Cell(2, m).BackColor = DicoCouleur.Item(Trim(tbl1(0))) ' la couleur est fonction de la tonalité
                Grid2.Cell(2, m).ForeColor = DicoCouleurLettre.Item(tbl1(0))
                Grid2.Cell(11, m).BackColor = DicoCouleur.Item(Trim(tbl1(0))) ' la couleur est fonction de la tonalité
                Grid2.Cell(11, m).ForeColor = DicoCouleurLettre.Item(tbl1(0))
                If m = 1 Then
                    Grid2.Cell(2, m).BackColor = Color.Red
                    Grid2.Cell(2, m).ForeColor = Color.Yellow
                    Grid2.Cell(11, m).BackColor = Color.Red
                    Grid2.Cell(11, m).ForeColor = Color.Yellow
                End If
                Grid2.AutoRedraw = True
                Grid2.Refresh()
                ' Maj Répétition
                ' **************
                Grid2.Cell(3, m).Text = Convert.ToString(TableEventH(m, t, ct).Répet)
            End If
        Next m
        ' divers grid2
        ' ************
        For i = 1 To Grid2.Cols - 1
            Grid2.Column(i).Width = Largeur_ZoomGrid2 ' maj largeur colonne grid2 en fonction du zoom au moment de la sauvegarde
        Next
        '
        Grid2.LeftCol = 1
        Grid2.Cell(1, 1).SetFocus()
        '
        Grid2.Refresh()
        Grid2.AutoRedraw = True
        'Grid2.Visible = True
    End Sub
    Function Det_DivParBeat() As Integer
        If Dénominateur = 4 Then
            Det_DivParBeat = 2 ' dénominateur = 4
        Else
            Det_DivParBeat = 3 ' dénominateur = 8
        End If
    End Function
    Private Sub ZoomPlus()
        Dim Largeur As Integer
        '
        Select Case Largeur_Zoom
            Case Grid3Largeur1
                Largeur = Grid3Largeur2
                Largeur_Zoom = Grid3Largeur2
            Case Grid3Largeur2
                Largeur = Grid3Largeur3
                Largeur_Zoom = Grid3Largeur3
            Case Grid3Largeur3
                Largeur = Grid3Largeur4
                Largeur_Zoom = Grid3Largeur4
            Case Grid3Largeur4
                Largeur = Largeur_Zoom
        End Select
        '
    End Sub

    Private Sub ZoomMoins()
        Dim Largeur As Integer

        Select Case Largeur_Zoom
            Case Grid3Largeur1
                Largeur = Largeur_Zoom
            Case Grid3Largeur2
                Largeur = Grid3Largeur1
                Largeur_Zoom = Grid3Largeur1
            Case Grid3Largeur3
                Largeur = Grid3Largeur2
                Largeur_Zoom = Grid3Largeur2
            Case Grid3Largeur4
                Largeur = Grid3Largeur3
                Largeur_Zoom = Grid3Largeur3
        End Select
        '
    End Sub
    Private Sub ZoomPlusGrid2()
        Dim Largeur As Integer
        Dim j As Integer
        '
        Select Case Largeur_ZoomGrid2
            Case Grid2Largeur0
                Largeur = Grid2Largeur1
                Largeur_ZoomGrid2 = Grid2Largeur1
            Case Grid2Largeur1, Grid2Largeur60
                Largeur = Grid2Largeur2
                Largeur_ZoomGrid2 = Grid2Largeur2
            Case Grid2Largeur2
                Largeur = Grid2Largeur3
                Largeur_ZoomGrid2 = Grid2Largeur3
            Case Grid2Largeur3
                Largeur = Grid2Largeur4
                Largeur_ZoomGrid2 = Grid2Largeur4
            Case Grid2Largeur4
                Largeur = Largeur_ZoomGrid2
            Case Else
                Largeur = Grid2Largeur1
        End Select
        '
        Grid2.AutoRedraw = False
        j = Grid2.LeftCol
        For i = 1 To (Grid2.Cols - 1) ' - 1)
            Grid2.Column(i).Width = Largeur
        Next i
        Grid2.LeftCol = j
        '
        Grid2.Refresh()
        Grid2.AutoRedraw = True
        '
        AssurerVisibilitéCelluleGrid2(Grid2.ActiveCell.Col) '(Grid2.LeftCol)
    End Sub
    Private Sub ZoomMoinsGrid2()
        Dim Largeur As Integer
        Dim j As Integer

        Select Case Largeur_ZoomGrid2
            Case Grid2Largeur0
                Largeur = Largeur_ZoomGrid2
            Case Grid2Largeur1, Grid2Largeur60
                Largeur = Grid2Largeur0
                Largeur_ZoomGrid2 = Grid2Largeur0
            Case Grid2Largeur2
                Largeur = Grid2Largeur1
                Largeur_ZoomGrid2 = Grid2Largeur1
            Case Grid2Largeur3
                Largeur = Grid2Largeur2
                Largeur_ZoomGrid2 = Grid2Largeur2
            Case Grid2Largeur4
                Largeur = Grid2Largeur3
                Largeur_ZoomGrid2 = Grid2Largeur3
            Case Else
                Largeur = Grid2Largeur1
        End Select
        '
        Grid2.AutoRedraw = False
        j = Grid2.LeftCol
        For i = 1 To (Grid2.Cols - 1) ' - 1)
            Grid2.Column(i).Width = Largeur
        Next i
        Grid2.LeftCol = j
        '
        Grid2.Refresh()
        Grid2.AutoRedraw = True
        '
        AssurerVisibilitéCelluleGrid2(Grid2.ActiveCell.Col) '(Grid2.LeftCol)
    End Sub
    Private Sub SplitContainer1_SplitterMoved(sender As Object, e As SplitterEventArgs)
        'Label17.Text = Str(SplitContainer1.SplitterDistance)
        'Label21.Text = "SplitContainer1"
    End Sub
    Private Sub SplitContainer3_SplitterMoved(sender As Object, e As SplitterEventArgs)
        'Label17.Text = Str(SplitContainer3.SplitterDistance)
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs)
        Select Case EtatSplit3
            Case "SansDétail"
            'SplitContainer3.SplitterDistance = SplitDétails
            'EtatSplit3 = "Détails"
            Case "Détails", "SansGrid3"
                'SplitContainer3.SplitterDistance = SplitSansDétail
                'EtatSplit3 = "SansDétail"
        End Select
    End Sub

    Private Sub Button2_Click_2(sender As Object, e As EventArgs)
        'ZoomMoins()
        AssurerVisibilitéCelluleGrid2(Grid2.ActiveCell.Col)
        ZoomMoinsGrid2()
    End Sub

    Private Sub Button1_Click_3(sender As Object, e As EventArgs)
        'ZoomPlus()
        AssurerVisibilitéCelluleGrid2(Grid2.ActiveCell.Col)
        ZoomPlusGrid2()
    End Sub
    Private Sub SplitContainer9_SplitterMoved(sender As Object, e As SplitterEventArgs)
        'Label17.Text = Str(SplitContainer9.SplitterDistance)
        'Label21.Text = "SplitContainer9"
    End Sub
    Function Det_Index_De_Ligne(ligne As Integer) As String
        Dim a As String
        Dim m As Integer
        Dim t As Integer
        Dim ct As Integer
        Dim sortir As Boolean
        '
        a = ""
        sortir = False
        For m = 0 To UBound(TableEventH, 1)
            For t = 0 To UBound(TableEventH, 2)
                For ct = 0 To UBound(TableEventH, 3)
                    If TableEventH(m, t, ct).Ligne = ligne Then
                        a = Trim(Str(m) + "-" + Str(t) + "-" + Str(ct))
                        sortir = True
                        Exit For
                    End If
                Next ct
                If sortir Then Exit For
            Next t
            If sortir Then Exit For
        Next m
        Det_Index_De_Ligne = a
    End Function
    Function Det_ValCelluleMarquée() As String
        Dim ind As Integer
        Dim ligne As Integer
        Dim degré As Integer
        '

        Det_ValCelluleMarquée = ""
        For ind = 0 To 20
            degré = Det_IndexDegré2(ind)
            ligne = Det_LigneTableGlobale(ind)
            If TableCoursAcc(ligne, degré).Marqué = True Then
                Det_ValCelluleMarquée = TableCoursAcc(ligne, degré).Accord
            End If
        Next
    End Function
    Private Function IndexVoix(Voix As String) As Integer

        Select Case Voix
            Case "-1"
                Return 0
            Case Else
                Return Val(Voix)
        End Select
    End Function
    Sub Maj_MenuTousAccords(m As Integer, t As Integer, ct As Integer) ' non utilisé
        Dim tbl1() As String
        Dim tbl2() As String
        Dim tbl3() As String
        Dim tbl4() As String
        Dim tbl5() As String
        Dim tbl6() As String
        Dim tbl7() As String
        '
        Dim Accord As String
        Dim Tonique As String
        Dim Chiffrage As String
        Dim Enrich As String
        '
        Dim i, k As Integer

        Dim degré As Integer
        Dim SuiteAcc As String

        k = -1

        degré = TableEventH(m, t, ct).Degré          ' détermination du degré de l'accord
        tbl1 = Split(TableEventH(m, t, ct).Mode, " ") ' détermination du mode de l'accord
        '
        ' version 3notes de l'accord
        ' **************************
        '
        SuiteAcc = Mode(tbl1(0), tbl1(1), 3)
        tbl3 = Split(SuiteAcc, "-")
        k = k + 1
        ReDim Preserve tbl2(k)
        tbl2(k) = tbl3(degré) '' lecture de l'accord

        ' version 4notes7 de l'accord
        ' ***************************
        '
        SuiteAcc = Mode(tbl1(0), tbl1(1), 4)
        tbl3 = Split(SuiteAcc, "-")
        k = k + 1
        ReDim Preserve tbl2(k)
        tbl2(k) = tbl3(degré) ' ' lecture de l'accord

        ' version 4notes9 de l'accord
        ' ***************************
        '
        SuiteAcc = Mode(tbl1(0), tbl1(1), 5)
        tbl3 = Split(SuiteAcc, "-") ' lecture de la suite des accords 9e
        '
        If Trim(tbl3(degré)) <> "___" Then
            tbl4 = Split(Trim(tbl3(degré)), "(") ' tbl3(degré)= l'accord
            tbl6 = Split(Trim(tbl4(0)), " ")
            tbl7 = Split(Trim(tbl4(1)), ")")
            Tonique = Trim(tbl6(0))
            Chiffrage = Microsoft.VisualBasic.Left(Trim(tbl6(1)), 1)
            Enrich = Trim(tbl7(0))
            If Chiffrage = "m" Then
                Accord = Tonique + " " + Chiffrage + Enrich
            Else
                Accord = Tonique + " " + Enrich
            End If

            k = k + 1
            ReDim Preserve tbl2(k)
            tbl2(k) = Accord
            k = k + 1
            ReDim Preserve tbl2(k)
            tbl2(k) = Trim(tbl3(degré))
        End If
        '
        ' version 4notes11 de l'accord
        ' ****************************
        SuiteAcc = Mode(tbl1(0), tbl1(1), 6)
        tbl3 = Split(SuiteAcc, "-") ' lecture de l'accord
        If Trim(tbl3(degré)) <> "___" Then
            tbl5 = Split(Trim(tbl3(degré)), " ") ' tbl5(0) = tonique;tbl5(1) = chiffrage
            Select Case tbl5(1)
                Case "m11"
                    k = k + 1
                    ReDim Preserve tbl2(k)
                    tbl2(k) = Trim(tbl3(degré))
                Case "11"
                    tbl4 = Split(tbl3(degré), " ")
                    Tonique = tbl4(0)
                    Accord = Tonique + " " + "sus4"
                    k = k + 1
                    ReDim Preserve tbl2(k)
                    tbl2(k) = Trim(Accord)
                    '
                    k = k + 1
                    ReDim Preserve tbl2(k)
                    tbl2(k) = Trim(tbl3(degré))
                Case "7M(#11)", "M7(11#)", "7(11)"
                    tbl4 = Split(Trim(tbl3(degré)), "(") ' dans tbl4(1) --> soit "(11" soit "(11#"
                    Select Case tbl4(1)
                        Case "11#)"
                            k = k + 1
                            ReDim Preserve tbl2(k)
                            tbl2(k) = Trim(tbl5(0)) + " 11#"
                            '
                            k = k + 1
                            ReDim Preserve tbl2(k)
                            tbl2(k) = Trim(tbl5(0)) + " M7(11#)"

                        Case "11)"
                            k = k + 1
                            ReDim Preserve tbl2(k)
                            tbl2(k) = Trim(tbl5(0)) + " 11"
                            '
                            k = k + 1
                            ReDim Preserve tbl2(k)
                            tbl2(k) = Trim(tbl5(0)) + " sus4"
                            '
                            k = k + 1
                            ReDim Preserve tbl2(k)
                            tbl2(k) = Trim(tbl5(0)) + " 7sus4"
                            '
                            k = k + 1
                            ReDim Preserve tbl2(k)
                            tbl2(k) = Trim(tbl5(0)) + " 7(11)"
                    End Select
            End Select
        End If
        '
        ' Mise à jour dans le menu
        ' ************************
        '
        Accord11.Visible = False
        Accord12.Visible = False
        Accord13.Visible = False
        Accord14.Visible = False
        Accord15.Visible = False
        Accord16.Visible = False
        Accord17.Visible = False
        Accord18.Visible = False
        '
        For i = 0 To UBound(tbl2)
            Select Case i
                Case 0
                    Accord11.Visible = True
                    Accord11.Text = tbl2(i)
                Case 1
                    Accord12.Visible = True
                    Accord12.Text = tbl2(i)
                Case 2
                    Accord13.Visible = True
                    Accord13.Text = tbl2(i)
                Case 3
                    Accord14.Visible = True
                    Accord14.Text = tbl2(i)
                Case 4
                    Accord15.Visible = True
                    Accord15.Text = tbl2(i)
                Case 5
                    Accord16.Visible = True
                    Accord16.Text = tbl2(i)
                Case 6
                    Accord17.Visible = True
                    Accord17.Text = tbl2(i)
                Case 7
                    Accord18.Visible = True
                    Accord18.Text = tbl2(i)
            End Select
        Next i
    End Sub
    Sub Maj_TousAccordsMnContext1(ind As Integer)
        Dim degré As Integer
        Dim ligne As Integer
        Dim tonique As String
        Dim Mode_ As String
        Dim tbl1() As String
        Dim tbl2() As String
        Dim tbl3() As String
        Dim ListeAcc As String
        Dim signe As String

        degré = Det_IndexDegré2(ind)
        ligne = Det_LigneTableGlobale(ind)
        '
        ' détermination du mode
        ' *********************
        Select Case ligne
            Case 0
                tbl1 = Split(Trim(ComboBox1.Text), " ")
                tonique = Trim(tbl1(0))
                If Langue = "fr" Then
                    signe = Det_ClefLat(Trim(tonique))
                Else
                    signe = Det_ClefAngl(Trim(tonique))
                End If
                tonique = TradNoteAngl(signe, Trim(tonique))
                Mode_ = Trim(UCaseBémol(tonique)) + " Maj"
            Case 1
                tbl1 = Split(Trim(ComboBox2.Text), " ")
                tonique = Trim(tbl1(0))
                If Langue = "fr" Then
                    signe = Det_ClefLat(Trim(tonique))
                Else
                    signe = Det_ClefAngl(Trim(tonique))
                End If
                tonique = TradNoteAngl(signe, Trim(tonique))
                Mode_ = Trim(UCaseBémol(tonique)) + " MinH"
            Case Else
                tbl1 = Split(Trim(ComboBox2.Text), " ")
                tonique = Trim(tbl1(0))
                If Langue = "fr" Then
                    signe = Det_ClefLat(Trim(tonique))
                Else
                    signe = Det_ClefAngl(Trim(tonique))
                End If
                tonique = TradNoteAngl(signe, Trim(tonique))
                Mode_ = Trim(UCaseBémol(tonique)) + " MinM"
        End Select
        '
        ' Détermination de la liste des accords possibles dans la cellule courante
        ' ************************************************************************
        Tonacours = Trim(ComboBox1.Text) ' variable globale utilisée par NoteInterval2 (Maj_MenuTousAccords2 --> Mode --> NoteInterval2
        ListeAcc = Maj_MenuTousAccords2(degré, Mode_)
        tbl2 = Split(ListeAcc, "-")
        '
        For i = 0 To UBound(tbl2)
            tbl3 = Split(tbl2(i), " ")
            tonique = TradNoteLat(LCase(tbl3(0)))
            If Langue <> "fr" Then
                tonique = Trim(tbl3(0)) ' comme TradNoteLat renvoie une note en minuscule quand Langue<>"fr", on reprend la valeur initiale qui était en majuscule
            End If
            tbl2(i) = Trim(tonique)
            If UBound(tbl3) > 0 Then
                tbl2(i) = tbl2(i) + " " + tbl3(1)
            End If
        Next
        '
        ' Mise à jour dans le menu
        ' ************************
        '
        Accord11_1.Visible = False
        Accord11_1.Tag = ind
        '
        Accord12_1.Visible = False
        Accord12_1.Tag = ind
        '
        Accord13_1.Visible = False
        Accord13_1.Tag = ind
        '
        Accord14_1.Visible = False
        Accord14_1.Tag = ind
        '
        Accord15_1.Visible = False
        Accord15_1.Tag = ind
        '
        Accord16_1.Visible = False
        Accord16_1.Tag = ind
        '
        Accord17_1.Visible = False
        Accord17_1.Tag = ind
        '
        Accord18_1.Visible = False
        Accord18_1.Tag = ind
        '
        For i = 0 To UBound(tbl2)
            Select Case i
                Case 0
                    Accord11_1.Visible = True
                    Accord11_1.Text = tbl2(i)
                Case 1
                    Accord12_1.Visible = True
                    Accord12_1.Text = tbl2(i)
                Case 2
                    Accord13_1.Visible = True
                    Accord13_1.Text = tbl2(i)
                Case 3
                    Accord14_1.Visible = True
                    Accord14_1.Text = tbl2(i)
                Case 4
                    Accord15_1.Visible = True
                    Accord15_1.Text = tbl2(i)
                Case 5
                    Accord16_1.Visible = True
                    Accord16_1.Text = tbl2(i)
                Case 6
                    Accord17_1.Visible = True
                    Accord17_1.Text = tbl2(i)
                Case 7
                    Accord18_1.Visible = True
                    Accord18_1.Text = tbl2(i)
            End Select
        Next i
    End Sub

    Sub VOIS_Maj_TousAccordsMnContext1(ind As Integer)
        Dim degré As Integer
        Dim ligne As Integer
        Dim tonique As String
        Dim Mode_ As String
        Dim tbl1() As String
        Dim tbl2() As String
        Dim tbl3() As String
        Dim ListeAcc As String
        Dim signe As String
        Dim a As String

        degré = Det_IndexDegré2(ind)
        ligne = Det_LigneTableGlobale(ind)
        '
        ' détermination du mode
        ' *********************
        a = Trim(TableGlobalAccVoisin(0, ligne, 0)) ' on prend le degré 1 (valeur=0) pour voir la toniqu
        If Langue = "fr" Then
            a = TradAcc_AnglLat(a)
            tbl1 = Split(a)
            tonique = Trim(tbl1(0))
            signe = Det_ClefLat(Trim(tonique))
        Else
            tbl1 = Split(a)
            tonique = Trim(tbl1(0))
            signe = Det_ClefAngl(Trim(tonique))
        End If

        tonique = TradNoteAngl(signe, Trim(tonique))
        Select Case ligne
            Case 0, 1, 2
                Mode_ = Trim(UCaseBémol(tonique)) + " Maj"
                Tonacours = Trim(a) + " Maj"
            Case Else 'ligne 2,3,4 (modes voisins mineurs)
                Mode_ = Trim(UCaseBémol(tonique)) + " MinH"
                Tonacours = Trim(a) + " MinH"
        End Select
        '
        ' Détermination de la liste des accords possibles dans la cellule courante
        ' ************************************************************************
        'Tonacours = Trim(Mode_)
        ListeAcc = Maj_MenuTousAccords2(degré, Mode_)
        tbl2 = Split(ListeAcc, "-")
        '
        For i = 0 To UBound(tbl2)
            tbl3 = Split(tbl2(i), " ")
            tonique = TradNoteLat(LCase(tbl3(0)))
            If Langue <> "fr" Then
                tonique = Trim(tbl3(0)) ' comme TradNoteLat renvoie une note en minuscule quand Langue<>"fr", on reprend la valeur initiale qui était en majuscule
            End If
            tbl2(i) = Trim(tonique)
            If UBound(tbl3) > 0 Then
                tbl2(i) = tbl2(i) + " " + tbl3(1)
            End If
        Next
        '
        ' Mise à jour dans le menu
        ' ************************
        '
        Accord11_1.Visible = False
        Accord11_1.Tag = ind
        '
        Accord12_1.Visible = False
        Accord12_1.Tag = ind
        '
        Accord13_1.Visible = False
        Accord13_1.Tag = ind
        '
        Accord14_1.Visible = False
        Accord14_1.Tag = ind
        '
        Accord15_1.Visible = False
        Accord15_1.Tag = ind
        '
        Accord16_1.Visible = False
        Accord16_1.Tag = ind
        '
        Accord17_1.Visible = False
        Accord17_1.Tag = ind
        '
        Accord18_1.Visible = False
        Accord18_1.Tag = ind
        '
        For i = 0 To UBound(tbl2)
            Select Case i
                Case 0
                    Accord11_1.Visible = True
                    Accord11_1.Text = tbl2(i)
                Case 1
                    Accord12_1.Visible = True
                    Accord12_1.Text = tbl2(i)
                Case 2
                    Accord13_1.Visible = True
                    Accord13_1.Text = tbl2(i)
                Case 3
                    Accord14_1.Visible = True
                    Accord14_1.Text = tbl2(i)
                Case 4
                    Accord15_1.Visible = True
                    Accord15_1.Text = tbl2(i)
                Case 5
                    Accord16_1.Visible = True
                    Accord16_1.Text = tbl2(i)
                Case 6
                    Accord17_1.Visible = True
                    Accord17_1.Text = tbl2(i)
                Case 7
                    Accord18_1.Visible = True
                    Accord18_1.Text = tbl2(i)
            End Select
        Next i
    End Sub
    Sub Maj_TousAccordsMnContextCAD(ind As Integer)
        Dim tonique As String
        Dim Mode_ As String = ""
        Dim tbl1() As String
        Dim tbl2() As String
        Dim tbl3() As String
        Dim ListeAcc, a As String
        Dim IndexDegré As Integer
        Dim signe As String

        ' détermination du degré et du mode
        ' *********************************
        a = TabCadDegrés.Item(ind).Text

        If Cad_OrigineAccord = Modes.Cadence_Majeure Then
            tbl1 = Split(Trim(ComboBox1.Text), " ")
            tonique = tbl1(0)
            If Langue = "fr" Then
                signe = Det_ClefLat(Trim(tbl1(0)))
                tonique = TradNoteAngl(signe, Trim((tbl1(0))))
                tonique = TradNoteEnMaj(tonique)
            End If
            IndexDegré = Det_IndexDegré(a)
            Mode_ = Trim(tonique) + " Maj"
        Else
            Select Case Trim(ComboBox4.Text)
                Case "Hispanique", "Hispanic" ' cas d'une cadence mineure mixte (Hispanique)
                    Select Case Trim(a)
                        Case "VI", "V", "IV"
                            tbl1 = Split(Trim(ComboBox1.Text), " ")
                            tonique = Trim(tbl1(0))
                            If Langue = "fr" Then
                                signe = Det_ClefLat(Trim(tbl1(0)))
                                tonique = TradNoteAngl(signe, Trim((tbl1(0))))
                                tonique = TradNoteEnMaj(tonique)
                            End If
                            IndexDegré = Det_IndexDegré(a)
                            Mode_ = Trim(tonique) + " Maj"
                        Case "III", "I"
                            tbl1 = Split(Trim(ComboBox2.Text), " ")
                            tonique = Trim(tbl1(0))
                            If Langue = "fr" Then
                                signe = Det_ClefLat(Trim(tbl1(0)))
                                tonique = TradNoteAngl(signe, Trim((tbl1(0))))
                                tonique = TradNoteEnMaj(tonique)
                            End If
                            IndexDegré = Det_IndexDegréMin(a)
                            Mode_ = Trim(tonique) + " MinH"
                    End Select
                Case Else ' cas cadence mineure non mixte
                    tbl1 = Split(Trim(ComboBox2.Text), " ")
                    tonique = Trim(tbl1(0))
                    If Langue = "fr" Then
                        signe = Det_ClefLat(Trim(tbl1(0)))
                        tonique = TradNoteAngl(signe, Trim((tbl1(0))))
                        tonique = TradNoteEnMaj(tonique)
                    End If
                    IndexDegré = Det_IndexDegréMin(a)
                    Mode_ = Trim(tonique) + " MinH"
            End Select
        End If
        '
        ' Détermination de la liste des accords possibles dans la cellule courante
        ' ************************************************************************
        ListeAcc = Maj_MenuTousAccords2(IndexDegré, Mode_)
        tbl2 = Split(ListeAcc, "-")
        '
        If Langue = "fr" Then
            For i = 0 To UBound(tbl2)
                tbl3 = Split(tbl2(i), " ")
                tonique = TradNoteLat(LCase(tbl3(0)))
                tbl2(i) = Trim(tonique)
                If UBound(tbl3) > 0 Then
                    tbl2(i) = tbl2(i) + " " + tbl3(1)
                End If
            Next
        End If
        ' Mise à jour dans le menu
        ' ************************
        '
        Accord11_2.Visible = False
        Accord11_2.Tag = ind
        '
        Accord12_2.Visible = False
        Accord12_2.Tag = ind
        '
        Accord13_2.Visible = False
        Accord13_2.Tag = ind
        '
        Accord14_2.Visible = False
        Accord14_2.Tag = ind
        '
        Accord15_2.Visible = False
        Accord15_2.Tag = ind
        '
        Accord16_2.Visible = False
        Accord16_2.Tag = ind
        '
        Accord17_2.Visible = False
        Accord17_2.Tag = ind
        '
        Accord18_2.Visible = False
        Accord18_2.Tag = ind
        '
        For i = 0 To UBound(tbl2)
            Select Case i
                Case 0
                    Accord11_2.Visible = True
                    Accord11_2.Text = tbl2(i)
                Case 1
                    Accord12_2.Visible = True
                    Accord12_2.Text = tbl2(i)
                Case 2
                    Accord13_2.Visible = True
                    Accord13_2.Text = tbl2(i)
                Case 3
                    Accord14_2.Visible = True
                    Accord14_2.Text = tbl2(i)
                Case 4
                    Accord15_2.Visible = True
                    Accord15_2.Text = tbl2(i)
                Case 5
                    Accord16_2.Visible = True
                    Accord16_2.Text = tbl2(i)
                Case 6
                    Accord17_2.Visible = True
                    Accord17_2.Text = tbl2(i)
                Case 7
                    Accord18_2.Visible = True
                    Accord18_2.Text = tbl2(i)
            End Select
        Next i
    End Sub
    Sub Maj_TousAccordsMnContext3(m As Integer, t As Integer, ct As Integer) ' utilisé par MenuContextGrid2Grid3 vérifier si non utilisé
        Dim degré As Integer
        Dim Mode_ As String
        Dim tbl1() As String
        Dim tbl2() As String
        Dim ListeAcc As String


        degré = TableEventH(m, t, ct).Degré           ' détermination du degré de l'accord
        Mode_ = TableEventH(m, t, ct).Mode
        tbl1 = Split(Mode_, " ")

        ListeAcc = Maj_MenuTousAccords2(degré, Mode_)
        tbl2 = Split(ListeAcc, "-")
        '
        ' Mise à jour dans le menu
        ' ************************
        '
        Accord11.Visible = False
        Accord12.Visible = False
        Accord13.Visible = False
        Accord14.Visible = False
        Accord15.Visible = False
        Accord16.Visible = False
        Accord17.Visible = False
        Accord18.Visible = False
        '
        For i = 0 To UBound(tbl2)
            Select Case i
                Case 0
                    Accord11.Visible = True
                    Accord11.Text = tbl2(i)
                    If Langue = "fr" Then
                        Accord11.Text = TradAcc_AnglLat(tbl2(i))
                    End If
                Case 1
                    Accord12.Visible = True
                    Accord12.Text = tbl2(i)
                    If Langue = "fr" Then
                        Accord12.Text = TradAcc_AnglLat(tbl2(i))
                    End If
                Case 2
                    Accord13.Visible = True
                    Accord13.Text = tbl2(i)
                    If Langue = "fr" Then
                        Accord13.Text = TradAcc_AnglLat(tbl2(i))
                    End If
                Case 3
                    Accord14.Visible = True
                    Accord14.Text = tbl2(i)
                    If Langue = "fr" Then
                        Accord14.Text = TradAcc_AnglLat(tbl2(i))
                    End If
                Case 4
                    Accord15.Visible = True
                    Accord15.Text = tbl2(i)
                    If Langue = "fr" Then
                        Accord15.Text = TradAcc_AnglLat(tbl2(i))
                    End If
                Case 5
                    Accord16.Visible = True
                    Accord16.Text = tbl2(i)
                    If Langue = "fr" Then
                        Accord16.Text = TradAcc_AnglLat(tbl2(i))
                    End If
                Case 6
                    Accord17.Visible = True
                    Accord17.Text = tbl2(i)
                    If Langue = "fr" Then
                        Accord17.Text = TradAcc_AnglLat(tbl2(i))
                    End If
                Case 7
                    Accord18.Visible = True
                    Accord18.Text = tbl2(i)
                    If Langue = "fr" Then
                        Accord18.Text = TradAcc_AnglLat(tbl2(i))
                    End If
            End Select
        Next i
    End Sub

    Function Det_EquivMinMaj(degré As Integer) ' cette procédure donne le degré d'un accord en mode mineur selon son degré en mode majeur
        Select Case degré
            Case 0
                Det_EquivMinMaj = 2
            Case 1
                Det_EquivMinMaj = 3
            Case 2
                Det_EquivMinMaj = 4
            Case 3
                Det_EquivMinMaj = 5
            Case 4
                Det_EquivMinMaj = 6
            Case 5
                Det_EquivMinMaj = 0
            Case 6
                Det_EquivMinMaj = 1
            Case Else
                Det_EquivMinMaj = 0
        End Select
    End Function

    Function Maj_MenuTousAccords2(degré As Integer, Mode_ As String) As String
        Dim tbl1() As String
        Dim tbl2() As String
        Dim tbl3() As String
        Dim tbl4() As String
        Dim tbl5() As String
        Dim tbl6() As String
        Dim tbl7() As String
        '
        Dim Accord As String
        Dim Tonique As String
        Dim Chiffrage As String
        Dim Enrich As String
        '
        Dim k As Integer
        Dim SuiteAcc As String

        k = -1

        tbl1 = Split(Mode_, " ") ' décomposition du mode de l'accord en tonalité et chiffrage
        '
        ' version 3notes de l'accord
        ' **************************
        '
        SuiteAcc = Mode(tbl1(0), tbl1(1), 3)
        tbl3 = Split(SuiteAcc, "-")
        k = k + 1
        ReDim Preserve tbl2(k)
        tbl2(k) = tbl3(degré) '' lecture de l'accord

        ' version 4notes7 de l'accord
        ' ***************************
        '
        SuiteAcc = Mode(tbl1(0), tbl1(1), 4)
        tbl3 = Split(SuiteAcc, "-")
        k = k + 1
        ReDim Preserve tbl2(k)
        tbl2(k) = tbl3(degré) ' ' lecture de l'accord

        ' version 4notes9 de l'accord
        ' ***************************
        '
        SuiteAcc = Mode(tbl1(0), tbl1(1), 5)
        tbl3 = Split(SuiteAcc, "-") ' lecture de la suite des accords 9e
        '
        If Trim(tbl3(degré)) <> "___" Then
            tbl4 = Split(Trim(tbl3(degré)), "(") ' tbl3(degré)= l'accord
            tbl6 = Split(Trim(tbl4(0)), " ")
            tbl7 = Split(Trim(tbl4(1)), ")")
            Tonique = Trim(tbl6(0))
            Chiffrage = Microsoft.VisualBasic.Left(Trim(tbl6(1)), 1)
            Enrich = Trim(tbl7(0))
            If Chiffrage = "m" Then
                Accord = Tonique + " " + Chiffrage + Enrich
            Else
                Accord = Tonique + " " + Enrich
            End If

            k = k + 1
            ReDim Preserve tbl2(k)
            tbl2(k) = Accord
            k = k + 1
            ReDim Preserve tbl2(k)
            tbl2(k) = Trim(tbl3(degré))
        End If
        '
        ' version 4notes11 de l'accord
        ' ****************************
        SuiteAcc = Mode(tbl1(0), tbl1(1), 6)
        tbl3 = Split(SuiteAcc, "-") ' lecture de l'accord
        If Trim(tbl3(degré)) <> "___" Then
            tbl5 = Split(Trim(tbl3(degré)), " ") ' tbl5(0) = tonique;tbl5(1) = chiffrage
            Select Case tbl5(1)
                Case "m11"
                    k = k + 1
                    ReDim Preserve tbl2(k)
                    tbl2(k) = Trim(tbl3(degré))
                    If (tbl1(1) = "Maj" And (degré = 1 Or degré = 2 Or degré = 5)) Or (tbl1(1) = "MinM" And degré = 1) Then
                        k = k + 1
                        ReDim Preserve tbl2(k)
                        tbl4 = Split(Trim(tbl3(degré)), "(") ' tbl3(degré)= l'accord
                        tbl6 = Split(Trim(tbl4(0)), " ")
                        tbl2(k) = tbl6(0) + " " + "m7(11)"
                    End If
                Case "11"
                    tbl4 = Split(tbl3(degré), " ")
                    Tonique = tbl4(0)
                    Accord = Tonique + " " + "sus4"
                    k = k + 1
                    ReDim Preserve tbl2(k)
                    tbl2(k) = Trim(Accord)
                    '
                    k = k + 1
                    ReDim Preserve tbl2(k)
                    tbl2(k) = Trim(tbl3(degré))
                Case "7M(11.0#)", "M7(11#)", "7(11)"
                    tbl4 = Split(Trim(tbl3(degré)), "(") ' dans tbl4(1) --> soit "11)" soit "11#)"
                    Select Case tbl4(1)
                        Case "11#)"
                            k = k + 1
                            ReDim Preserve tbl2(k)
                            tbl2(k) = Trim(tbl5(0)) + " 11#"
                            '
                            k = k + 1
                            ReDim Preserve tbl2(k)
                            tbl2(k) = Trim(tbl5(0)) + " M7(11#)"

                        Case "11)"
                            k = k + 1
                            ReDim Preserve tbl2(k)
                            tbl2(k) = Trim(tbl5(0)) + " 11"
                            '
                            k = k + 1
                            ReDim Preserve tbl2(k)
                            tbl2(k) = Trim(tbl5(0)) + " sus4"
                            '
                            k = k + 1
                            ReDim Preserve tbl2(k)
                            tbl2(k) = Trim(tbl5(0)) + " 7Sus4"
                            '
                            k = k + 1
                            ReDim Preserve tbl2(k)
                            tbl2(k) = Trim(tbl5(0)) + " 7(11)"
                    End Select
            End Select
        End If
        '
        Maj_MenuTousAccords2 = Join(tbl2, "-")
    End Function


    Function ChaineAccord(m As Integer) As String
        Dim a As String
        '
        Dim t, ct, nbAccords As Integer

        a = ""
        nbAccords = 0
        For t = 0 To 5 ' t temps
            For ct = 0 To 4 ' ct contre temps 
                If Trim(TableEventH(m, t, ct).Accord) <> "" Then
                    nbAccords = nbAccords + 1
                    If Trim(a) = "" Then
                        a = Trim(TradAcc_AnglLat(Trim(TableEventH(m, t, ct).Accord)))
                    Else
                        a = a + " / " + Trim(TradAcc_AnglLat(Trim(TableEventH(m, t, ct).Accord)))
                    End If
                End If
            Next ct
        Next t
        '
        If nbAccords = 1 Then
            Grid2.Cell(1, m).Alignment = FlexCell.AlignmentEnum.CenterCenter
        Else
            Grid2.Cell(1, m).Alignment = FlexCell.AlignmentEnum.LeftCenter
        End If
        '
        ChaineAccord = Trim(a)
    End Function
    Function Det_NbDivisionMesure() As String
        If Dénominateur = 4 Then
            Det_NbDivisionMesure = Numérateur * 2
        Else '(toujours 8)
            Det_NbDivisionMesure = Numérateur
        End If
    End Function



    Sub Maj_TabTons(Typ As Integer) ' typ est le type d'accord : 3 ntes, 4 notes avec 7, 9e, 11e

        Dim i As Integer
        ' ton majeur
        ' **********
        'tbl1 = Split(Mode("C", "Maj", 3), "-")
        For i = 0 To 6
            TabTons.Item(i).Text = TradEventHLat(TableGlobalAcc(Typ, 0, i)) 'tbl1(i)
        Next i
        ' tons mineurs Harmonique
        ' ***********************
        'tbl1 = Split(Mode("A", "MinH", 3), "-")
        For i = 7 To 13
            TabTons.Item(i).Text = TradEventHLat(TableGlobalAcc(Typ, 1, i - 7)) 'tbl1(i - 7)
        Next i
        '
        ' tons mineurs Mélodiques
        ' ***********************
        'tbl1 = Split(Mode("A", "MinM", 3), "-")
        For i = 14 To 20
            TabTons.Item(i).Text = TradEventHLat(TableGlobalAcc(Typ, 2, i - 14)) 'tbl1(i - 14)
        Next i
        '
    End Sub
    Public Sub Maj_TabTonsVoisins(Typ As Integer) ' typ est le type d'accord : 3 ntes, 4 notes avec 7, 9e, 11e

        Dim i As Integer
        ' I Majeur
        ' **********
        'tbl1 = Split(Mode("C", "Maj", 3), "-")
        For i = 0 To 6
            TabTonsVoisins.Item(i).Text = TradEventHLat(TableGlobalAccVoisin(Typ, 0, i)) 'tbl1(i)
        Next i
        ' IV Majeur
        ' *********
        'tbl1 = Split(Mode("A", "MinH", 3), "-")
        For i = 7 To 13
            TabTonsVoisins.Item(i).Text = TradEventHLat(TableGlobalAccVoisin(Typ, 1, i - 7)) 'tbl1(i - 7)
        Next i
        '
        ' V Majeur
        ' ********
        'tbl1 = Split(Mode("A", "MinM", 3), "-")
        For i = 14 To 20
            TabTonsVoisins.Item(i).Text = TradEventHLat(TableGlobalAccVoisin(Typ, 2, i - 14)) 'tbl1(i - 14)
        Next i
        '
        '
        ' I Mineur
        ' *********
        'tbl1 = Split(Mode("A", "MinM", 3), "-")
        For i = 21 To 27
            TabTonsVoisins.Item(i).Text = TradEventHLat(TableGlobalAccVoisin(Typ, 3, i - 21)) 'tbl1(i - 14)
        Next i
        '
        ' IV Mineur
        ' **********
        'tbl1 = Split(Mode("A", "MinM", 3), "-")
        For i = 28 To 34
            TabTonsVoisins.Item(i).Text = TradEventHLat(TableGlobalAccVoisin(Typ, 4, i - 28)) 'tbl1(i - 14)
        Next i
        ' V Mineur
        ' **********
        'tbl1 = Split(Mode("A", "MinM", 3), "-")
        For i = 35 To 41
            TabTonsVoisins.Item(i).Text = TradEventHLat(TableGlobalAccVoisin(Typ, 5, i - 35)) 'tbl1(i - 14)
        Next i
    End Sub



    Public Sub Maj_TableGlobalAcc(tona As String, tonaMin As String)
        Dim tbl() As String
        Dim i As Integer
        Dim j As Integer

        ' Maj
        For i = 3 To 6
            tbl = Split(Mode2(Trim(tona), "Maj", i, False), "-")
            For j = 0 To 6
                TableGlobalAcc(i - 3, 0, j) = tbl(j)
            Next j
        Next i
        ' MinH
        For i = 3 To 6
            tbl = Split(Mode2(Trim(tonaMin), "MinH", i, False), "-")
            For j = 0 To 6
                TableGlobalAcc(i - 3, 1, j) = tbl(j)
            Next j
        Next i
        ' MinM
        For i = 3 To 6
            tbl = Split(Mode2(Trim(tonaMin), "MinM", i, False), "-")
            For j = 0 To 6
                TableGlobalAcc(i - 3, 2, j) = tbl(j)
            Next j
        Next i

    End Sub

    Public Sub Maj_TableGlobalAccVoisin(tona As String, tonaMin As String) ' les paramètres tona et tonamin sont reçus 
        Dim Mode1Voisin As String = Trim(tona) ' 
        Dim Mode2Voisin As String ' 
        Dim Mode3Voisin As String ' 
        Dim Mode4Voisin As String = Trim(tonaMin) ' 
        Dim Mode5Voisin As String '
        Dim Mode6Voisin As String '
        Dim i, j As Integer


        ' Détermination des Modes 
        ' ***********************
        ' Mode Voisin I majeur (Do Majeur)
        Mode1Voisin = Det_TonStandardMaj(Trim(Mode1Voisin))
        For i = 3 To 6
            tbl = Split(Mode2(Trim(Mode1Voisin), "Maj", i, True), "-") ' Mode1Voisin est mis à jour dans sa déclaration
            For j = 0 To 6
                TableGlobalAccVoisin(i - 3, 0, j) = tbl(j)
            Next j
        Next i
        ' Mode Voisin IV majeur (Fa Majeur)
        Mode2Voisin = NoteInterval(LCaseBémol(tona), "4")
        Mode2Voisin = UCaseBémol(Mode2Voisin)
        Mode2Voisin = Det_TonStandardMaj(Trim(Mode2Voisin))
        For i = 3 To 6
            tbl = Split(Mode2(Trim(Mode2Voisin), "Maj", i, True), "-")
            For j = 0 To 6
                TableGlobalAccVoisin(i - 3, 1, j) = tbl(j)
            Next j
        Next i
        ' Mode voisin V majeur (Sol Majeur)
        Mode3Voisin = NoteInterval(LCaseBémol(tona), "5")
        Mode3Voisin = UCaseBémol(Mode3Voisin)
        Mode3Voisin = Det_TonStandardMaj(Trim(Mode3Voisin))
        For i = 3 To 6
            tbl = Split(Mode2(Trim(Mode3Voisin), "Maj", i, True), "-")
            For j = 0 To 6
                TableGlobalAccVoisin(i - 3, 2, j) = tbl(j)
            Next j
        Next i
        '
        ' Mode Voisin I mineur (La Mineur)
        Mode4Voisin = Det_TonStandardMin(Trim(Mode4Voisin))
        For i = 3 To 6
            tbl = Split(Mode2(Trim(Mode4Voisin), "MinH", i, True), "-") ' Mode1Voisin est mis à jour dans sa déclaration
            For j = 0 To 6
                TableGlobalAccVoisin(i - 3, 3, j) = tbl(j)
            Next j
        Next i
        '
        ' Mode voisin IV mineur (Ré mineur)
        Mode5Voisin = NoteInterval(LCaseBémol(tonaMin), "4")
        Mode5Voisin = UCaseBémol(Mode5Voisin)
        Mode5Voisin = Det_TonStandardMin(Trim(Mode5Voisin))
        For i = 3 To 6
            tbl = Split(Mode2(Trim(Mode5Voisin), "MinH", i, True), "-")
            For j = 0 To 6
                TableGlobalAccVoisin(i - 3, 4, j) = tbl(j)
            Next j
        Next i
        '
        ' Mode voisin V mineur (Mi mineur)
        Mode6Voisin = NoteInterval(LCaseBémol(tonaMin), "5")
        Mode6Voisin = UCaseBémol(Mode6Voisin)
        Mode6Voisin = Det_TonStandardMin(Trim(Mode6Voisin))
        For i = 3 To 6
            tbl = Split(Mode2(Trim(Mode6Voisin), "MinH", i, True), "-")
            For j = 0 To 6
                TableGlobalAccVoisin(i - 3, 5, j) = tbl(j)
            Next j
        Next i
        ' Mise à jour de TableGlobalAccVoisin (4 tonalité à mettre à jour)
        Refresh()
    End Sub
    '
    Public Function Det_TonStandardMaj(Ton As String) As String
        Dim a As String
        Select Case Ton
            Case "Db"
                a = "C#"
            Case "Gb"
                a = "F#"
            Case "A#"
                a = "Bb"
            Case "D#"
                a = "Eb"
            Case "G#"
                a = "Ab"
            Case Else
                a = Trim(Ton)
        End Select
        Det_TonStandardMaj = a
    End Function
    Public Function Det_TonStandardMin(Ton As String)
        Dim a As String
        Select Case Ton
            Case "Bb"
                a = "A#"
            Case "Eb"
                a = "D#"
            Case "Ab"
                a = "G#"
            Case "Db"
                a = "C#"
            Case "Gb"
                a = "F#"
            Case Else
                a = Trim(Ton)
        End Select
        Det_TonStandardMin = a
    End Function


    Public Sub RAZ_AccTransit()
        Dim i As Integer
        '
        For i = 0 To 34
            TabTonsVoisinsMarq.Item(i).Visible = False
        Next i
    End Sub





    ' Menu Contextuel Accords/3Notes
    ' ******************************
    Private Sub NotesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Menu3notes.Click

        Dim i As Integer
        Dim col As Integer
        '
        Dim a As String
        Dim IndexDegré As Integer
        Dim Ligne As Integer
        '
        For i = 0 To 20
            '
            Ligne = Det_LigneTableGlobale(i)
            Select Case Ligne
                Case 0
                    col = i
                Case 1
                    col = i - 7
                Case 2
                    col = i - 14
            End Select
            a = TabTonsDegrés.Item(col).Text
            IndexDegré = Det_IndexDegré(a)
            '
            TabTons.Item(i).Text = TableGlobalAcc(0, Ligne, IndexDegré)
            '
            TableCoursAcc(Ligne, IndexDegré).TyAcc = Menu3notes.Text
            TableCoursAcc(Ligne, IndexDegré).Accord = Trim(TabTons.Item(i).Text)
            '
            'Maj_Renversement(i)
            'End If
        Next i
        '
    End Sub
    Sub Menu3_notes()
        Dim i As Integer
        Dim col As Integer
        '
        Dim a As String
        Dim IndexDegré As Integer
        Dim Ligne As Integer
        '
        For i = 0 To 20
            'If TabTonsSelect.Item(i).Checked = True Then
            '
            Ligne = Det_LigneTableGlobale(i)
            Select Case Ligne
                Case 0
                    col = i
                Case 1
                    col = i - 7
                Case 2
                    col = i - 14
            End Select
            a = TabTonsDegrés.Item(col).Text
            IndexDegré = Det_IndexDegré(a)
            '
            'If ChangeLangue = False Then
            TabTons.Item(i).Text = TradEventHLat(TableGlobalAcc(0, Ligne, IndexDegré))
            'End If
            '

            TableCoursAcc(Ligne, IndexDegré).TyAcc = Menu3notes.Text
            TableCoursAcc(Ligne, IndexDegré).Accord = Trim(TabTons.Item(i).Text)
            '
            'Maj_Renversement(i)
            'End If
        Next i
    End Sub
    Function AccordPourCadenceMixte(ind As Integer, TYaccord As Integer) As String
        Dim a As String
        Dim IndexDegré As Integer

        AccordPourCadenceMixte = ""
        Select Case Trim(ComboBox4.Text)
            Case "Hispanique", "Hispanic"
                a = TabCadDegrés.Item(ind).Text
                If Trim(a) <> "" Then
                    IndexDegré = Det_IndexDegréMin(a)
                    Select Case Trim(a)
                        Case "V", "IV"
                            IndexDegré = Det_IndexDegré(a)
                            AccordPourCadenceMixte = CAD_TableGlobalAcc(TYaccord, 0, IndexDegré)
                        Case "III", "VI"
                            IndexDegré = Det_IndexDegréMin(a)
                            AccordPourCadenceMixte = CAD_TableGlobalAcc(TYaccord, 1, IndexDegré)
                    End Select
                End If
        End Select
    End Function
    Sub CAD_Menu3_notes()
        Dim i As Integer
        '
        Dim a As String
        Dim IndexDegré As Integer
        '
        Cad_RAZ_CouleurMarquée()
        '
        For i = 0 To 4
            '

            '

            a = TabCadDegrés.Item(i).Text
            If Trim(a) <> "" Then
                If Cad_OrigineAccord = Modes.Cadence_Mixte Then

                    Select Case Trim(a)
                        Case "V", "IV"
                            IndexDegré = Det_IndexDegré(a)
                        Case "VI", "III"
                            IndexDegré = Det_IndexDegréMin(a)
                    End Select
                Else
                    IndexDegré = Det_IndexDegré(a)
                End If
                '
                If Cad_OrigineAccord <> Modes.Cadence_Mixte Then ' cadence mixte n'existe plus 
                    If Cad_OrigineAccord = Modes.Cadence_Majeure Then
                        'CAD_TableGlobalAcc(0, 0, IndexDegré) = TableGlobalAcc(0, 0, IndexDegré)
                        TabCad.Item(i).Text = TradEventHLat(CAD_TableGlobalAcc(0, 0, IndexDegré))
                    Else
                        IndexDegré = Det_IndexDegréMin(a)
                        'CAD_TableGlobalAcc(0, 0, IndexDegré) = TableGlobalAcc(0, 0, IndexDegré)
                        TabCad.Item(i).Text = TradEventHLat((CAD_TableGlobalAcc(0, 1, IndexDegré)))
                    End If
                Else
                    TabCad.Item(i).Text = TradEventHLat((AccordPourCadenceMixte(i, ComboBox6.SelectedIndex)))
                End If
                '
                'TabCad.Item(i).Text = CAD_TableCoursAcc(i).Accord
                CAD_TableCoursAcc(IndexDegré).TyAcc = Menu3notes.Text
                CAD_TableCoursAcc(IndexDegré).Accord = TradAcc_LatAngl(Trim(TabCad.Item(i).Text))
                '
                '
            End If
            'End If
        Next i
    End Sub
    Sub CAD_Menu4_notes()
        Dim i As Integer
        '
        Dim a As String
        Dim IndexDegré As Integer
        '
        Cad_RAZ_CouleurMarquée()
        '
        For i = 0 To 4
            '
            '
            'TabCadFiltres4.Item(i).Visible = False
            'TabCadFiltres7.Item(i).Visible = False

            a = TabCadDegrés.Item(i).Text
            If Trim(a) <> "" Then
                IndexDegré = Det_IndexDegré(a)
                '
                'If Cad_OrigineAccord <> Modes.Cadence_Mixte Then
                'TabCad.Item(i).Text = CAD_TableGlobalAcc(0, 0, IndexDegré)
                'Else
                'TabCad.Item(i).Text = AccordPourCadenceMixte(i, 0)
                'End If
                '
                If Cad_OrigineAccord <> Modes.Cadence_Mixte Then
                    If Cad_OrigineAccord = Modes.Cadence_Majeure Then
                        'CAD_TableGlobalAcc(0, 0, IndexDegré) = TableGlobalAcc(0, 0, IndexDegré)
                        TabCad.Item(i).Text = TradEventHLat(CAD_TableGlobalAcc(1, 0, IndexDegré))
                    Else
                        IndexDegré = Det_IndexDegréMin(a)
                        'CAD_TableGlobalAcc(0, 0, IndexDegré) = TableGlobalAcc(0, 0, IndexDegré)
                        TabCad.Item(i).Text = TradEventHLat(CAD_TableGlobalAcc(1, 1, IndexDegré))
                    End If
                Else
                    TabCad.Item(i).Text = TradEventHLat(AccordPourCadenceMixte(i, ComboBox6.SelectedIndex))
                End If
                '

                CAD_TableCoursAcc(IndexDegré).TyAcc = Menu4Notes.Text
                CAD_TableCoursAcc(IndexDegré).Accord = TradAcc_LatAngl(Trim(TabCad.Item(i).Text))
                '

            End If
            'End If
        Next i
    End Sub


    ' Menu Contextuel Accords/3Notes+7
    ' *********************************
    Private Sub NotesToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles Menu4Notes.Click
        Dim i As Integer
        Dim col As Integer
        '
        Dim a As String
        Dim IndexDegré As Integer
        Dim Ligne As Integer
        '
        For i = 0 To 20
            'If TabTonsSelect.Item(i).Checked = True Then
            Ligne = Det_LigneTableGlobale(i)
            Select Case Ligne
                Case 0
                    col = i
                Case 1
                    col = i - 7
                Case 2
                    col = i - 14
            End Select
            a = TabTonsDegrés.Item(col).Text
            IndexDegré = Det_IndexDegré(a)
            '
            TabTons.Item(i).Text = TableGlobalAcc(1, Ligne, IndexDegré)
            '

            TableCoursAcc(Ligne, IndexDegré).TyAcc = Menu4Notes.Text
            TableCoursAcc(Ligne, IndexDegré).Accord = Trim(TabTons.Item(i).Text)
            '
            'Maj_Renversement(i)

            'End If
        Next i
    End Sub
    Sub Menu4_notes()
        Dim i As Integer
        Dim col As Integer
        '
        Dim a As String
        Dim IndexDegré As Integer
        Dim Ligne As Integer
        '
        For i = 0 To 20
            ' If TabTonsSelect.Item(i).Checked = True Then
            Ligne = Det_LigneTableGlobale(i)
            Select Case Ligne
                Case 0
                    col = i
                Case 1
                    col = i - 7
                Case 2
                    col = i - 14
            End Select
            a = TabTonsDegrés.Item(col).Text
            IndexDegré = Det_IndexDegré(a)
            '
            TabTons.Item(i).Text = TradEventHLat(TableGlobalAcc(1, Ligne, IndexDegré))
            '
            TableCoursAcc(Ligne, IndexDegré).TyAcc = Menu4Notes.Text
            TableCoursAcc(Ligne, IndexDegré).Accord = Trim(TabTons.Item(i).Text)
            '
            'Maj_Renversement(i)
            'End If

        Next i
    End Sub

    Private Sub Notes9eToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MenuNotes9.Click
        Dim i As Integer
        Dim col As Integer
        '
        Dim a As String
        Dim IndexDegré As Integer
        Dim Ligne As Integer
        '
        For i = 0 To 20
            'If TabTonsSelect.Item(i).Checked = True Then

            Ligne = Det_LigneTableGlobale(i)
            Select Case Ligne
                Case 0
                    col = i
                Case 1
                    col = i - 7
                Case 2
                    col = i - 14
            End Select
            a = TabTonsDegrés.Item(col).Text
            IndexDegré = Det_IndexDegré(a)
            '
            TabTons.Item(i).Text = TableGlobalAcc(2, Ligne, IndexDegré)
            '

            TableCoursAcc(Ligne, IndexDegré).TyAcc = MenuNotes9.Text
            TableCoursAcc(Ligne, IndexDegré).Accord = Trim(TabTons.Item(i).Text)
            '
            'Maj_Renversement(i)
            'End If
        Next i

        'AffichageFiltres("9")
    End Sub
    ' Menu Contextuel Accords/3Notes+11
    ' *********************************
    Private Sub Notes11eToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MenuNotes11.Click
        Dim i As Integer
        Dim col As Integer
        '
        Dim a As String
        Dim IndexDegré As Integer
        Dim Ligne As Integer
        '
        For i = 0 To 20
            'If TabTonsSelect.Item(i).Checked = True Then

            Ligne = Det_LigneTableGlobale(i)
            Select Case Ligne
                Case 0
                    col = i
                Case 1
                    col = i - 7
                Case 2
                    col = i - 14
            End Select
            a = TabTonsDegrés.Item(col).Text
            IndexDegré = Det_IndexDegré(a)
            '
            TabTons.Item(i).Text = TableGlobalAcc(3, Ligne, IndexDegré)
            '

            TableCoursAcc(Ligne, IndexDegré).TyAcc = MenuNotes11.Text
            TableCoursAcc(Ligne, IndexDegré).Accord = Trim(TabTons.Item(i).Text)
            '
            'Maj_Renversement(i)
            'End If

        Next i
        '
        'AffichageFiltres("11")
    End Sub
    ' Octave +1
    ' *********
    Private Sub ToolStripMenuItem10_Click(sender As Object, e As EventArgs) Handles OctavePlus1.Click
        OctavePlus1.Checked = True
        Octave0.Checked = False
        OctaveMoins1.Checked = False
        OctaveMoins2.Checked = False
        '
        Maj_ChoixOctave(0, "+1")
    End Sub
    ' Octave = 0
    ' **********
    Private Sub ToolStripMenuItem11_Click(sender As Object, e As EventArgs) Handles Octave0.Click
        OctavePlus1.Checked = False
        Octave0.Checked = True
        OctaveMoins1.Checked = False
        OctaveMoins2.Checked = False
        '
        Maj_ChoixOctave(1, "0")
    End Sub
    ' Octave -1
    ' *********
    Private Sub ToolStripMenuItem12_Click(sender As Object, e As EventArgs) Handles OctaveMoins1.Click
        OctavePlus1.Checked = False
        Octave0.Checked = False
        OctaveMoins1.Checked = True
        OctaveMoins2.Checked = False
        '
        Maj_ChoixOctave(2, "-1")
    End Sub
    ' Octave -2
    ' *********
    Private Sub OctaveMoins2_Click(sender As Object, e As EventArgs) Handles OctaveMoins2.Click
        OctavePlus1.Checked = False
        Octave0.Checked = False
        OctaveMoins1.Checked = False
        OctaveMoins2.Checked = True
        '
        Maj_ChoixOctave(3, "-2")
    End Sub
    Sub Maj_ChoixOctave(OctaveChoisie As Integer, Octave As String)
        Dim i As Integer
        Dim ligne As Integer
        Dim degré As Integer

        For i = 0 To 20
            'If TabTonsSelect.Item(i).Checked = True Then
            If TabTons.Item(i).BackColor = Color.Red Then
                ligne = Det_LigneTableGlobale(i)
                degré = Det_IndexDegré2(i)
                '
                TableCoursAcc(ligne, degré).OctaveChoisie = OctaveChoisie
                TableCoursAcc(ligne, degré).Octave = Octave
                '
            End If
        Next i
    End Sub
    Private Sub Button19_Click(sender As Object, e As EventArgs)
        Dim i As Integer
        Dim Tonique() As Object

        Tonique = Split(Trim(ComboBox1.Text))
        Tonique(0) = LCaseBémol(Tonique(0))
        i = Det_Modulation(Tonique(0), "4")
        ComboBox1.SelectedIndex = i
    End Sub
    Private Sub Button20_Click(sender As Object, e As EventArgs)
        Dim i As Integer
        Dim Tonique() As Object

        Tonique = Split(Trim(ComboBox1.Text))
        Tonique(0) = LCaseBémol(Tonique(0))
        i = Det_Modulation(Tonique(0), "5")
        ComboBox1.SelectedIndex = i
    End Sub
    Sub EVT_Mesure()
        CallB_Aff_Mesure = True
    End Sub
    Sub EVT_Accord()
        CallB_Aff_Accord = True
    End Sub
    Sub EVT_Gamme()
        CallB_Aff_Gamme = True
    End Sub
    Sub EVT_Métronome()
        'My.Computer.Audio.Play(My.Resources.rimshotV2, AudioPlayMode.Background)
    End Sub
    '

    Sub Maj_DicoNotes()
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer


        Dim List_T As New List(Of TT)
        Dim Notes(0 To 11) As String
        Dim T0, T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11 As New TT

        List_T.Add(T0) '
        List_T.Add(T1)
        List_T.Add(T2)
        List_T.Add(T3)
        List_T.Add(T4)
        List_T.Add(T5)
        List_T.Add(T6)
        List_T.Add(T7)
        List_T.Add(T8)
        List_T.Add(T9)
        List_T.Add(T10)
        List_T.Add(T11)
        '
        Notes(0) = "c"
        Notes(1) = "c#"
        Notes(2) = "d"
        Notes(3) = "d#"
        Notes(4) = "e"
        Notes(5) = "f"
        Notes(6) = "f#"
        Notes(7) = "g"
        Notes(8) = "g#"
        Notes(9) = "a"
        Notes(10) = "a#"
        Notes(11) = "b"


        'T.Toctave(0) = "A"
        For k = 0 To 11
            j = 0 + k
            For i = 0 To 10
                List_T.Item(k).Toctave(i) = j
                j = j + 12
            Next
            DicoNotes.Add(Notes(k), List_T.Item(k))
            ' Entrée des notes exprimées en "bémol"
            Select Case Notes(k)
                Case "c#"
                    DicoNotes.Add("db", List_T.Item(k))
                Case "d#"
                    DicoNotes.Add("eb", List_T.Item(k))
                Case "f#"
                    DicoNotes.Add("gb", List_T.Item(k))
                Case "g#"
                    DicoNotes.Add("ab", List_T.Item(k))
                Case "a#"
                    DicoNotes.Add("bb", List_T.Item(k))
            End Select
        Next k

    End Sub

    Sub FIN()
        'SortieMidi.Item(ChoixSortieMidi).SilenceAllNotes()
        'SortieMidi.Item(ChoixSortieMidi).SilenceAllNotes()
        If Remote.Checked Then Send_CTRL55_Remote()
        ToutesNotesOff()
        Tempo_Aff_EventH.Enabled = False
        Tempo_Aff_EventH.Stop()  ' arrêter la tempo d'arrêt de jeu de l'accord
        CallB_Aff_Acc = False
        CallB_Aff_NumMes = False
        CallB_Aff_FIN = False
        '
        Init_CTRLMIDI2() ' réinit des ctrl par exemple pédale = 0
        Fermer_MIDI()
        '
        ' réactivation des boutons plays :barrre fixe et barre flottante
        ' **************************************************************
        PlayMidi.Enabled = True
        Transport.Button1.Enabled = True
        ' 
        'TimerToutesNotesOff.Interval = 200
        'TimerToutesNotesOff.Start()
        '
        SortieMidi.Item(ChoixSortieMidi).SilenceAllNotes()

    End Sub
    Sub StopRecalcul()
        ToutesNotesOff()
        Tempo_Aff_EventH.Enabled = False
        Tempo_Aff_EventH.Stop()  ' arrêter la tempo d'arrêt de jeu de l'accord
        CallB_Aff_Acc = False
        CallB_Aff_NumMes = False
        CallB_Aff_FIN = False
        '
        Fermer_MIDI()
        '
        SortieMidi.Item(ChoixSortieMidi).SilenceAllNotes()
    End Sub
    Sub MenuContextGrid2Grid3(m As Integer, t As Integer, ct As Integer)
        Dim a As String

        a = TableEventH(m, t, ct).Accord
        If Trim(a) <> "" Then
            ContextMenu3Accord.Text = Trim(a)
            If Langue = "fr" Then ' langue de traitement et non pas langueIHM
                ContextMenu3Accord.Text = TradAcc_AnglLat(Trim(a))
            End If
            Maj_TousAccordsMnContext3(m, t, ct)
        End If
    End Sub
    Function Det_AccordDsMesure(m As Integer) As String ' Détermination de la mesure d'un accord
        Dim t, ct As Integer
        Det_AccordDsMesure = ""
        Dim sortir As Boolean = False
        For t = 0 To UBound(TableEventH, 2) 'nbTempsMesure '- 1
            For ct = 0 To UBound(TableEventH, 3) 'nbDivTemps '- 1
                If Trim(TableEventH(m, t, ct).Accord) <> "" Then
                    Det_AccordDsMesure = Trim(Str(m)) + " " + Trim(Str(t)) + " " + Trim(Str(ct))
                    sortir = True
                    Exit For
                End If
            Next ct
            If sortir Then
                Exit For
            End If
        Next t
    End Function
    Sub Maj_PropriétésEntrée2()

        Dim A As Modes
        Dim i As Integer = TabControl2.SelectedIndex


        If TabControl2.TabPages.Item(i).Tag = 1 Then
            A = Cad_OrigineAccord
        Else
            A = OrigineAccord
        End If
        '
        Entrée_Accord = AccordMarqué
        Select Case A
            Case Modes.Majeur, Modes.Cadence_Majeure
                Entrée_Gamme = Trim(Det_TonaCours2() + " " + "Maj")
            Case Modes.MineurH, Modes.Cadence_Mineure
                Entrée_Gamme = Trim(Det_TonaMinCours2() + " " + "MinH")
            Case Modes.MineurM
                Entrée_Gamme = Trim(Det_TonaMinCours2() + " " + "MinM")
        End Select
        Dim aa As String = Trim(ComboBox1.Text)
        Dim l As Integer = aa.Length
        Entrée_Tonalité = Microsoft.VisualBasic.Left(Trim(ComboBox1.Text), l - 2) 'Entrée_Gamme 'initialement Tonalité=Gamme
        '
        Entrée_Mode = Entrée_Gamme

        '

    End Sub

    Function Trad_DegréRomains(Num As Integer) As String
        Trad_DegréRomains = "I"
        Select Case Num
            Case 0
                Trad_DegréRomains = "I"
            Case 1
                Trad_DegréRomains = "II"
            Case 2
                Trad_DegréRomains = "III"
            Case 3
                Trad_DegréRomains = "IV"
            Case 4
                Trad_DegréRomains = "V"
            Case 5
                Trad_DegréRomains = "VI"
            Case 6
                Trad_DegréRomains = "VII"
        End Select
    End Function
    Function Det_NbTempsMesure() As Integer
        Select Case Dénominateur
            Case 4
                Det_NbTempsMesure = Numérateur
            Case 8
                Det_NbTempsMesure = Numérateur / 3
            Case Else
                Det_NbTempsMesure = -1
        End Select
    End Function
    Function Det_NbDivTemps() As Integer
        Select Case Dénominateur
            Case 4
                Det_NbDivTemps = 2
            Case 8
                Det_NbDivTemps = 3

            Case Else
                Det_NbDivTemps = -1
        End Select
    End Function
    Function Det_NbDivTemps2(temps As Integer) As Integer
        Select Case Dénominateur
            Case 4
                Det_NbDivTemps2 = 2
            Case 8
                Select Case Numérateur
                    Case 7 ' cas 7/8
                        Det_NbDivTemps2 = 3
                        If temps = 3 Then
                            Det_NbDivTemps2 = 4
                        End If
                    Case Else
                        Det_NbDivTemps2 = 3
                End Select
            Case Else
                Det_NbDivTemps2 = -1
        End Select
    End Function
    Function Det_IndexGrid3_De_ColGrid2(ColGrid2 As Integer) As Integer
        Det_IndexGrid3_De_ColGrid2 = (Det_NbDivisionMesure() * (ColGrid2 - 1)) + 1
    End Function
    Function Det_MarqueurDsMesure(m As Integer) As Boolean
        Dim t As Integer
        Dim ct As Integer
        Dim sortir As Boolean = False
        '
        Det_MarqueurDsMesure = False
        For t = 0 To UBound(TableEventH, 2)
            For ct = 0 To UBound(TableEventH, 3)
                If Trim(TableEventH(m, t, ct).Marqueur) <> "" Then
                    Det_MarqueurDsMesure = True
                    sortir = True
                    Exit Function
                End If
            Next ct
            '
            If sortir = True Then
                Exit Function
            End If
        Next t
    End Function

    Function Det_NomMode(Tonalité As String, Degré As String) As String
        Dim tbl() As String

        tbl = Split(Tonalité, " ")
        Det_NomMode = ""
        Select Case tbl(1)
            Case "Maj"
                Select Case Degré
                    Case "I"
                        If Langue = "fr" Then
                            Det_NomMode = "Ionien"
                        Else
                            Det_NomMode = "Ionian"
                        End If
                    Case "II"
                        If Langue = "fr" Then
                            Det_NomMode = "Dorien"
                        Else
                            Det_NomMode = "Dorian"
                        End If
                    Case "III"
                        If Langue = "fr" Then
                            Det_NomMode = "Prygien"
                        Else
                            Det_NomMode = "Prygian"
                        End If
                    Case "IV"
                        If Langue = "fr" Then
                            Det_NomMode = "Lydien"
                        Else
                            Det_NomMode = "Lydian"
                        End If
                    Case "V"
                        If Langue = "fr" Then
                            Det_NomMode = "MixoLydien"
                        Else
                            Det_NomMode = "MixoLydian"
                        End If
                    Case "VI"
                        If Langue = "fr" Then
                            Det_NomMode = "Aeolien"
                        Else
                            Det_NomMode = "Aeolian"
                        End If
                    Case "VII"
                        If Langue = "fr" Then
                            Det_NomMode = "Locrien"
                        Else
                            Det_NomMode = "Locrian"
                        End If
                End Select
            Case "MinH"
                Select Case Degré
                    Case "I"
                        If Langue = "fr" Then
                            Det_NomMode = "MinH"
                        Else
                            Det_NomMode = "MinH"
                        End If
                    Case "II"
                        If Langue = "fr" Then
                            Det_NomMode = "Locrien B6"
                        Else
                            Det_NomMode = "Locrian B6"
                        End If
                    Case "III"
                        If Langue = "fr" Then
                            Det_NomMode = "Ionien 5#"
                        Else
                            Det_NomMode = "Ionian 5#"
                        End If
                    Case "IV"
                        If Langue = "fr" Then
                            Det_NomMode = "Dorien 4#"
                        Else
                            Det_NomMode = "Dorian 4#"
                        End If
                    Case "V"
                        If Langue = "fr" Then
                            Det_NomMode = "SuperPhrygien"
                        Else
                            Det_NomMode = "SuperPhrygian"
                        End If
                    Case "VI"
                        If Langue = "fr" Then
                            Det_NomMode = "Lydien 9#"
                        Else
                            Det_NomMode = "Lydian 9#"
                        End If
                    Case "VII"
                        If Langue = "fr" Then
                            Det_NomMode = "Altéré"
                        Else
                            Det_NomMode = "Altered"
                        End If
                End Select
            Case "MinM"
                Select Case Degré
                    Case "I"
                        If Langue = "fr" Then
                            Det_NomMode = "MinM"
                        Else
                            Det_NomMode = "MinM"
                        End If
                    Case "II"
                        If Langue = "fr" Then
                            Det_NomMode = "Dorien 2b"
                        Else
                            Det_NomMode = "Dorian 2b"
                        End If
                    Case "III"
                        If Langue = "fr" Then
                            Det_NomMode = "Lydien augmenté"
                        Else
                            Det_NomMode = "Augmented Lydian"
                        End If
                    Case "IV"
                        If Langue = "fr" Then
                            Det_NomMode = "Lydien 7b"
                        Else
                            Det_NomMode = "Lydian 7b"
                        End If
                    Case "V"
                        If Langue = "fr" Then
                            Det_NomMode = "Mixolydien 6b"
                        Else
                            Det_NomMode = "Mixolydian 6b"
                        End If
                    Case "VI"
                        If Langue = "fr" Then
                            Det_NomMode = "Aéolien b5"
                        Else
                            Det_NomMode = "Aéolian b5"
                        End If
                    Case "VII"
                        If Langue = "fr" Then
                            Det_NomMode = "SuperLocrien"
                        Else
                            Det_NomMode = "SuperLocrian"
                        End If
                End Select
            Case "MajH"
                Select Case Degré
                    Case "I"
                        If Langue = "fr" Then
                            Det_NomMode = "MajH"
                        Else
                            Det_NomMode = "MajH"
                        End If
                    Case "II"
                        If Langue = "fr" Then
                            Det_NomMode = "Dorien b5"
                        Else
                            Det_NomMode = "Dorian b5"
                        End If
                    Case "III"
                        If Langue = "fr" Then
                            Det_NomMode = "Phrygien b4"
                        Else
                            Det_NomMode = "Phrygian b4"
                        End If
                    Case "IV"
                        If Langue = "fr" Then
                            Det_NomMode = "Lydien b3"
                        Else
                            Det_NomMode = "Lydian b3"
                        End If
                    Case "V"
                        If Langue = "fr" Then
                            Det_NomMode = "Mixolydien b2"
                        Else
                            Det_NomMode = "Mixolydian b2"
                        End If
                    Case "VI"
                        If Langue = "fr" Then
                            Det_NomMode = "Lydien Aug 2#"
                        Else
                            Det_NomMode = "Lydian Aug 2#"
                        End If
                    Case "VII"
                        If Langue = "fr" Then
                            Det_NomMode = "Locrien bb7"
                        Else
                            Det_NomMode = "Locrian bb7"
                        End If
                End Select
        End Select

    End Function
    Function Trad_NomMode(Mode As String, LangueLoc As String) As String

        'tbl = Split(Mode, " ")
        Trad_NomMode = ""
        '
        Select Case Mode 'tbl(1)

            Case "Ionien", "Ionian"
                If LangueLoc = "fr" Then
                    Trad_NomMode = "Ionien"
                Else
                    Trad_NomMode = "Ionian"
                End If
            Case "Dorien", "Dorian"
                If LangueLoc = "fr" Then
                    Trad_NomMode = "Dorien"
                Else
                    Trad_NomMode = "Dorian"
                End If
            Case "Prygien", "Prygian"
                If LangueLoc = "fr" Then
                    Trad_NomMode = "Prygien"
                Else
                    Trad_NomMode = "Prygian"
                End If
            Case "Lydien", "Lydian"
                If LangueLoc = "fr" Then
                    Trad_NomMode = "Lydien"
                Else
                    Trad_NomMode = "Lydian"
                End If
            Case "MixoLydien", "MixoLydian"
                If LangueLoc = "fr" Then
                    Trad_NomMode = "MixoLydien"
                Else
                    Trad_NomMode = "MixoLydian"
                End If
            Case "Aeolien", "Aeolian"
                If LangueLoc = "fr" Then
                    Trad_NomMode = "Aeolien"
                Else
                    Trad_NomMode = "Aeolian"
                End If
            Case "Locrien", "Locrian"
                If LangueLoc = "fr" Then
                    Trad_NomMode = "Locrien"
                Else
                    Trad_NomMode = "Locrian"
                End If


            Case "MinH"
                If LangueLoc = "fr" Then
                    Trad_NomMode = "MinH"
                Else
                    Trad_NomMode = "MinH"
                End If
            Case "Locrien B6", "Locrian B6"
                If LangueLoc = "fr" Then
                    Trad_NomMode = "Locrien B6"
                Else
                    Trad_NomMode = "Locrian B6"
                End If
            Case "Ionien 5#", "Ionian 5#"
                If LangueLoc = "fr" Then
                    Trad_NomMode = "Ionien 5#"
                Else
                    Trad_NomMode = "Ionian 5#"
                End If
            Case "Dorien 4#", "Dorian 4#"
                If LangueLoc = "fr" Then
                    Trad_NomMode = "Dorien 4#"
                Else
                    Trad_NomMode = "Dorian 4#"
                End If
            Case "SuperPhrygien", "SuperPhrygian"
                If LangueLoc = "fr" Then
                    Trad_NomMode = "SuperPhrygien"
                Else
                    Trad_NomMode = "SuperPhrygian"
                End If
            Case "VI"
                If LangueLoc = "fr" Then
                    Trad_NomMode = "Lydien 9#"
                Else
                    Trad_NomMode = "Lydian 9#"
                End If
            Case "Altéré", "Altered"
                If LangueLoc = "fr" Then
                    Trad_NomMode = "Altéré"
                Else
                    Trad_NomMode = "Altered"
                End If


            Case "MinM"
                If LangueLoc = "fr" Then
                    Trad_NomMode = "MinM"
                Else
                    Trad_NomMode = "MinM"
                End If
            Case "Dorien 2b", "Dorian 2b"
                If LangueLoc = "fr" Then
                    Trad_NomMode = "Dorien 2b"
                Else
                    Trad_NomMode = "Dorian 2b"
                End If
            Case "Lydien augmenté", "Augmented Lydian"
                If LangueLoc = "fr" Then
                    Trad_NomMode = "Lydien augmenté"
                Else
                    Trad_NomMode = "Augmented Lydian"
                End If
            Case "Lydien 7b", "Lydian 7b"
                If LangueLoc = "fr" Then
                    Trad_NomMode = "Lydien 7b"
                Else
                    Trad_NomMode = "Lydian 7b"
                End If
            Case "Mixolydien 6b", "Mixolydian 6b"
                If LangueLoc = "fr" Then
                    Trad_NomMode = "Mixolydien 6b"
                Else
                    Trad_NomMode = "Mixolydian 6b"
                End If
            Case "Aéolien b5", "Aéolian b5"
                If LangueLoc = "fr" Then
                    Trad_NomMode = "Aéolien b5"
                Else
                    Trad_NomMode = "Aéolian b5"
                End If
            Case "SuperLocrien", "SuperLocrian"
                If LangueLoc = "fr" Then
                    Trad_NomMode = "SuperLocrien"
                Else
                    Trad_NomMode = "SuperLocrian"
                End If


            Case "MajH"
                If LangueLoc = "fr" Then
                    Trad_NomMode = "MajH"
                Else
                    Trad_NomMode = "MajH"
                End If
            Case "Dorien b5", "Dorian b5"
                If LangueLoc = "fr" Then
                    Trad_NomMode = "Dorien b5"
                Else
                    Trad_NomMode = "Dorian b5"
                End If
            Case "Phrygien b4", "Phrygian b4"
                If LangueLoc = "fr" Then
                    Trad_NomMode = "Phrygien b4"
                Else
                    Trad_NomMode = "Phrygian b4"
                End If
            Case "Lydien b3", "Lydian b3"
                If LangueLoc = "fr" Then
                    Trad_NomMode = "Lydien b3"
                Else
                    Trad_NomMode = "Lydian b3"
                End If
            Case "Mixolydien b2", "Mixolydian b2"
                If LangueLoc = "fr" Then
                    Trad_NomMode = "Mixolydien b2"
                Else
                    Trad_NomMode = "Mixolydian b2"
                End If
            Case "Lydien Aug 2#", "Lydian Aug 2#"
                If LangueLoc = "fr" Then
                    Trad_NomMode = "Lydien Aug 2#"
                Else
                    Trad_NomMode = "Lydian Aug 2#"
                End If
            Case "Locrien bb7", "Locrian bb7"
                If LangueLoc = "fr" Then
                    Trad_NomMode = "Locrien bb7"
                Else
                    Trad_NomMode = "Locrian bb7"
                End If
            Case Else
                Trad_NomMode = Mode
        End Select


    End Function
    Function Det_LimiteSélection(Sélection As Integer, NLigne As Integer) As Integer
        Dim i As Integer
        '
        i = Sélection - NLigne
        If i >= 0 Then
            Det_LimiteSélection = (Sélection - NLigne) + 1
        Else
            Det_LimiteSélection = 0
        End If
    End Function




    Sub AffTonaCours(tona As String)
        '
    End Sub

    Sub AffTonaChoisie(tona As String)
        ComboBox1.BackColor = DicoCouleur.Item(tona)
        ComboBox1.ForeColor = DicoCouleurLettre.Item(tona)
    End Sub


    Private Sub ComboBox2_MouseDown(sender As Object, e As MouseEventArgs)
        OrigineTona = "Min"
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs)
        If EnChargement = False Then
            ComboBox1.SelectedIndex = ComboBox2.SelectedIndex
        Else
            ComboBox1.SelectedIndex = ComboBox2.SelectedIndex
        End If
        'End If
    End Sub
    Private Sub ComboBox23_SelectedIndexChanged(sender As Object, e As EventArgs)
        If EnChargement = False Then ' 
            ChoixTypeAccord()
            ' CAD_Maj_TableGlobalAcc(Mode_Cadence)
        End If
    End Sub
    Sub ChoixTypeAccord()
        If EnChargement = False Then
            Select Case ComboBox23.SelectedIndex
                Case 0
                    Menu3_notes()
                'CAD_Menu3_notes()
                Case 1
                    Menu4_notes()
                'CAD_Menu4_notes()
                Case 2
                    Accords_9e()
                'CAD_Accords_9e()
                Case 3
                    Accords_4e_11e()
                    'CAD_Accords_4e_11e()
            End Select
        End If
    End Sub


    ' Menu Contextuel Accords/Notes+9
    ' *******************************
    Private Sub EToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EToolStripMenuItem.Click
        Dim col As Integer
        '
        Dim a As String
        Dim IndexDegré As Integer
        Dim Ligne As Integer
        '

        For i = 0 To 20
            'If TabTonsSelect.Item(i).Checked = True Then

            Ligne = Det_LigneTableGlobale(i)
            Select Case Ligne
                Case 0
                    col = i
                Case 1
                    col = i - 7
                Case 2
                    col = i - 14
            End Select
            a = TabTonsDegrés.Item(col).Text
            IndexDegré = Det_IndexDegré(a)
            '
            TabTons.Item(i).Text = TableGlobalAcc(2, Ligne, IndexDegré)
            '

            TableCoursAcc(Ligne, IndexDegré).TyAcc = MenuNotes9.Text
            TableCoursAcc(Ligne, IndexDegré).Accord = Trim(TabTons.Item(i).Text)
            '
            'Maj_Renversement(i)
            'End If
        Next i

        'AffichageFiltres("9")
    End Sub
    Sub Accords_9e()
        Dim col As Integer
        '
        Dim a As String
        Dim IndexDegré As Integer
        Dim Ligne As Integer
        '
        For i = 0 To 20
            'If TabTonsSelect.Item(i).Checked = True Then

            Ligne = Det_LigneTableGlobale(i)
            Select Case Ligne
                Case 0
                    col = i
                Case 1
                    col = i - 7
                Case 2
                    col = i - 14
            End Select
            '
            a = TabTonsDegrés.Item(col).Text
            IndexDegré = Det_IndexDegré(a)
            '
            TabTons.Item(i).Text = TableGlobalAcc(2, Ligne, IndexDegré)
            '

            TableCoursAcc(Ligne, IndexDegré).TyAcc = MenuNotes9.Text
            TableCoursAcc(Ligne, IndexDegré).Accord = Trim(TabTons.Item(i).Text)
            '
            ' Maj_Renversement(i)
            'End If
        Next i

        'AffichageFiltres("9")
    End Sub
    '

    ' Menu Contextuel Accords/3Notes+11
    ' *********************************
    Private Sub EToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles EToolStripMenuItem1.Click
        Dim i As Integer
        Dim col As Integer
        '
        Dim a As String
        Dim IndexDegré As Integer
        Dim Ligne As Integer
        '
        For i = 0 To 20
            'If TabTonsSelect.Item(i).Checked = True Then

            Ligne = Det_LigneTableGlobale(i)
            Select Case Ligne
                Case 0
                    col = i
                Case 1
                    col = i - 7
                Case 2
                    col = i - 14
            End Select
            a = TabTonsDegrés.Item(col).Text
            IndexDegré = Det_IndexDegré(a)
            '
            TabTons.Item(i).Text = TableGlobalAcc(3, Ligne, IndexDegré)
            '

            TableCoursAcc(Ligne, IndexDegré).TyAcc = MenuNotes11.Text
            TableCoursAcc(Ligne, IndexDegré).Accord = Trim(TabTons.Item(i).Text)
            '
            'Maj_Renversement(i)
            'End If

        Next i
        '
        'AffichageFiltres("11")
    End Sub
    Sub Accords_4e_11e()
        Dim i As Integer
        Dim col As Integer
        '
        Dim a As String
        Dim IndexDegré As Integer
        Dim Ligne As Integer
        '
        For i = 0 To 20
            'If TabTonsSelect.Item(i).Checked = True Then

            Ligne = Det_LigneTableGlobale(i)
            Select Case Ligne
                Case 0
                    col = i
                Case 1
                    col = i - 7
                Case 2
                    col = i - 14
            End Select
            a = TabTonsDegrés.Item(col).Text
            IndexDegré = Det_IndexDegré(a)
            '
            TabTons.Item(i).Text = TableGlobalAcc(3, Ligne, IndexDegré)
            '

            TableCoursAcc(Ligne, IndexDegré).TyAcc = MenuNotes11.Text
            TableCoursAcc(Ligne, IndexDegré).Accord = Trim(TabTons.Item(i).Text)
            '
            'Maj_Renversement(i)
            'End If

        Next i
        '
        'AffichageFiltres("11")
    End Sub
    ' **********************************************************************
    ' * Début Classe ENR                                                   *
    ' **********************************************************************
    ''' <summary>
    ''' CLASS ENR : permet de lister les différents Paramètre à Enregistrer
    ''' </summary>
    Class ENR
        Public ListSyst As String = ""
        Public ListnumAcc As String = ""
        Public ListMarq As String = ""
        Public ListChiffAcc As String = ""
        Public ListRépet As String = ""
        Public ListMagnétos As String = ""
        Public ListRacines As String = ""
        Public ListMute As String = ""
        Public ListVolume As String = "" ' < --Volume
        Public ListPRG As String = ""
        Public ListMotifs As String = ""
        Public ListMotifs_text As String = ""
        Public ListDurées As String = ""
        Public ListOctaves As String = ""
        Public ListDyn As String = ""
        Public ListDyn2 As String = ""
        Public ListSouches As String = ""
        Public ListRetards As String = ""
        Public ListDelay As String = ""
        Public ListDébutSouche As String = ""
        Public ListRadio1 As String = ""
        Public ListRadio2 As String = ""
        Public ListRadio3 As String = ""
        Public ListAccent As String = ""
        Public ListNomduSon As String = ""
        ' constructeur
        ' ************
        Sub New()
            Maj_ListEntrées()
            Maj_ListGénérateurs()
            Maj_ListSyst()
        End Sub

        ' Méthodes
        ' ********
        Private Sub Maj_ListEntrées()
            Dim i As Integer
            For i = 1 To nbMesures ' - 1
                ListnumAcc = ListnumAcc + "," + Convert.ToString(TableEventH(i, 1, 1).NumAcc)
                ListMarq = ListMarq + "," + Trim(TableEventH(i, 1, 1).Marqueur)
                ListChiffAcc = ListChiffAcc + "," + Trim(TableEventH(i, 1, 1).Accord)
                ListRépet = ListRépet + "," + Convert.ToString(TableEventH(i, 1, 1).Répet)
                ListMagnétos = ListMagnétos + "," + Convert.ToString(TableEventH(i, 1, 1).NumMagnéto) ' magnétos = variations
                ListRacines = ListRacines + "," + Trim(Form1.Grid2.Cell(12, i).Text)
            Next
            ListnumAcc = "ListnumAcc" + ListnumAcc
            ListMarq = "ListMarq" + ListMarq
            ListChiffAcc = "ListChiffAcc" + ListChiffAcc
            ListRépet = "ListRépet" + ListRépet
            ListMagnétos = "ListMagnétos" + ListMagnétos
            ListRacines = "ListRacines" + ListRacines
        End Sub

        Private Sub Maj_ListGénérateurs() ' générateur = variations
            Dim i As Integer

            For i = 0 To nb_BlocPistes - 1 'Arrangement1.Nb_Pistes - 1
                ListMute = ListMute + "," + Convert.ToString(Form1.SelBloc(i).Checked)
                ListPRG = ListPRG + "," + Convert.ToString(Form1.PistePRG.Item(i).SelectedIndex)
                ListMotifs = ListMotifs + "," + Convert.ToString(Form1.PisteMotif.Item(i).SelectedIndex)
                ListMotifs_text = ListMotifs_text + "," + Trim(Form1.BoutMotif.Item(i).Text)
                ListDurées = ListDurées + "," + Convert.ToString(Form1.PisteDurée(i).SelectedIndex)
                ListOctaves = ListOctaves + "," + Convert.ToString(Form1.PisteOctave.Item(i).SelectedIndex)
                ListDyn = ListDyn + "," + Convert.ToString(Form1.PisteDyn.Item(i).SelectedIndex)
                ListDyn2 = ListDyn2 + "," + Convert.ToString(Form1.PisteDyn2.Item(i).Value)
                ListSouches = ListSouches + "," + Convert.ToString(Form1.PisteSouche.Item(i).SelectedIndex)
                ListRetards = ListRetards + "," + Convert.ToString(Form1.PisteRetard.Item(i).SelectedIndex)
                ListDelay = ListDelay + "," + Convert.ToString(Form1.PisteDelay(i).Checked)
                ListDébutSouche = ListDébutSouche + "," + Convert.ToString(Form1.PisteDébut(i).Checked)
                ListRadio1 = ListRadio1 + "," + Convert.ToString(Form1.PisteRadio1.Item(i).Checked)
                ListRadio2 = ListRadio2 + "," + Convert.ToString(Form1.PisteRadio2.Item(i).Checked)
                ListRadio3 = ListRadio3 + "," + Convert.ToString(Form1.PisteRadio3.Item(i).Checked)
                ListAccent = ListAccent + "," + Convert.ToString(Form1.PisteAccent.Item(i).SelectedIndex)
                ListNomduSon = ListNomduSon + "," + Convert.ToString(Form1.NomduSon.Item(i).Text)
            Next i
            '
            ListMute = "ListMute" + ListMute
            ListPRG = "ListPRG" + ListPRG
            ListMotifs = "ListMotifs" + ListMotifs
            ListMotifs_text = "ListMotifs_text" + ListMotifs_text
            ListDurées = "ListDurées" + ListDurées
            ListOctaves = "ListOctaves" + ListOctaves
            ListDyn = "ListDyn" + ListDyn
            ListDyn2 = "ListDyn2" + ListDyn2
            ListDelay = "ListDelay" + ListDelay ' à supprimer
            ListDébutSouche = "ListDébutSouche" + ListDébutSouche
            ListSouches = "ListSouches" + ListSouches
            ListRetards = "ListRetards" + ListRetards
            ListRadio1 = "ListRadio1" + ListRadio1
            ListRadio2 = "ListRadio2" + ListRadio2
            ListRadio3 = "ListRadio3" + ListRadio3
            ListAccent = "ListAccent" + ListAccent
            ListNomduSon = "ListNomduSon" + ListNomduSon
            '
            ' Liste des Volumes
            For i = 0 To nb_TotalPistes - 1 'nb_TotalPistes - 1 'Arrangement1.Nb_PistesMidi '  Nb_PistesMidi donne l'index de la dernière piste
                ListVolume = ListVolume + "," + Convert.ToString(Form1.Mix.PisteVolume.Item(i).Value)
            Next i
            ListVolume = "ListVolume" + ListVolume
        End Sub
        Sub Maj_ListSyst()
            ' A noter : ' le paramètre Répéter n'est jamais sauvegardé - la valeur de base est toujours False = pas de répétition
            ListSyst = "ListSyst" + "," + "Tempo" + " " + Convert.ToString(Form1.Tempo.Value) _
                + "," + "Métrique" + " " + "4/4" _
                + "," + "Début" + " " + Convert.ToString(Form1.Début.Value) _
                + "," + "Fin" + " " + Convert.ToString(Form1.Terme.Value) _
                + "," + "Répéter" + " " + Convert.ToString(False) _
                + "," + "TopRow" + " " + Convert.ToString(Form1.listPIANOROLL(0).PTopRow) _
                + "," + "Accent" + " " + Convert.ToString(Form1.Accent1_3) _
                + "," + "Largeur_ZoomGrid2" + " " + Convert.ToString(Largeur_ZoomGrid2) _
                + "," + "Tonalité" + " " + Convert.ToString(Form1.ComboBox1.SelectedIndex) _
                + "," + "RacineCours" + " " + Convert.ToString(Form1.TRacine.SelectedIndex) _
                + "," + "FiltreUnissons" + " " + Convert.ToString(Form1.FiltreUni.Checked) _
                + "," + "4Notes" + " " + Convert.ToString(Form1.Fournotes.Checked) _
                + "," + "V4notes" + " " + Convert.ToString(Form1.ComboBox11.SelectedIndex) _
                + "," + "Bassemoins1" + " " + Convert.ToString(Form1.Bassemoins1.Checked) _
                + "," + "VBassemoins1" + " " + Convert.ToString(Form1.ComboBox12.SelectedIndex)
        End Sub
    End Class
    ' ****************************************************
    ' * FIN CLASSE ENR                                   *
    ' ****************************************************
    '
    Private Sub EnregistrerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EnregistrerToolStripMenuItem.Click
        Dim enreg As New ENR
        Enregistrer()
    End Sub
    Sub Enregistrer()

        If IndicateurEnreg = False Then
            If Det_NomFich() Then
                IndicateurEnreg = True
                EnregistrerSous()
            End If
        Else
            If File.Exists(FichierEnreg) Then '< ----- ici
                EnregistrerSous()
            Else
                If LangueIHM = "fr" Then
                    MessageHV.PContenuMess = "Erreur : votre fichier n'existe plus. Vérifier vos connexions USB. Utilisez Enregsitrer Sous"
                Else
                    MessageHV.PContenuMess = "Error : your file no longer exists. Check your USB connections. Use save as"
                End If
                MessageHV.PTypBouton = "OK"
                MessageHV.ShowDialog()
            End If
        End If
    End Sub

    Private Sub EnregistrerSousToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EnregistrerSousToolStripMenuItem.Click
        If Det_NomFich() Then
            EnregistrerSous()
        End If
    End Sub
    Function Det_NomFich() As Boolean ' détermination du chemin fichier pour enregistrement de *.zic3
        Dim FInfo As FileInfo
        Dim NomFichSeul As String

        Det_NomFich = True
        SaveFileDialog1.Filter = "HyperArp Files (*.zic4*)|*.zic4"
        SaveFileDialog1.OverwritePrompt = True ' affichage message "Existe déjà ...."
        'a = My.Computer.FileSystem.GetDirectoryInfo(CheminFichierEnreg)
        'SaveFileDialog1.InitialDirectory = a.ToString 'CheminFichierEnreg
        SaveFileDialog1.InitialDirectory = CheminFichierEnreg
        '
        If LangueIHM = "fr" Then
            SaveFileDialog1.Title = "Enregistrement projet"
        Else
            SaveFileDialog1.Title = "Save project"
        End If
        '
        'PianoRoll1.F1.Hide()
        '
        For i = 0 To nb_PianoRoll - 1
            'listPIANOROLL(i).F1.Hide()
        Next
        '
        Cacher_FormTransparents()

        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            CheminFichierEnreg = SaveFileDialog1.FileName 'FInfo.DirectoryName
            FichierEnreg = SaveFileDialog1.FileName
            FInfo = My.Computer.FileSystem.GetFileInfo(FichierEnreg)
            NomFichSeul = My.Computer.FileSystem.GetName(FichierEnreg)
            SaveFileDialog1.FileName = NomFichSeul
            Me.Text = NomFichSeul
            CheminFichierEnreg = FInfo.DirectoryName
        Else
            Det_NomFich = False
        End If
        '
    End Function

    Function Det_NomCalquesMIDI() As Boolean ' export calques MIDI
        Dim Finfo As FileInfo

        Det_NomCalquesMIDI = True

        If LangueIHM = "fr" Then
            SaveFileDialog1.Title = "Export Calques MIDI"
        Else
            SaveFileDialog1.Title = "MIDI Layers Export"
        End If
        '
        SaveFileDialog1.Filter = "MIDI Files (*.mid*)|*.mid"
        SaveFileDialog1.OverwritePrompt = True ' affichage message "Exise déjà ...."
        SaveFileDialog1.InitialDirectory = CheminFichierCalques
        SaveFileDialog1.FileName = ""
        '
        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            FichierCalques = SaveFileDialog1.FileName
            Finfo = My.Computer.FileSystem.GetFileInfo(FichierCalques)
            CheminFichierCalques = Finfo.DirectoryName
        Else
            Det_NomCalquesMIDI = False
        End If
    End Function
    Function Det_NomMarqueursMIDI() As Boolean ' export marquers
        Dim Finfo As FileInfo

        Det_NomMarqueursMIDI = True

        If LangueIHM = "fr" Then
            SaveFileDialog1.Title = "Export Marqueurs MIDI"
        Else
            SaveFileDialog1.Title = "MIDI Markers Export"
        End If
        SaveFileDialog1.Filter = "MIDI Files (*.mid*)|*.mid"
        SaveFileDialog1.OverwritePrompt = True ' affichage message "Exise déjà ...."
        SaveFileDialog1.InitialDirectory = CheminMarqueursMIDI
        SaveFileDialog1.FileName = ""
        '
        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            FichierMIDI = SaveFileDialog1.FileName
            Finfo = My.Computer.FileSystem.GetFileInfo(FichierMIDI)
            CheminMarqueursMIDI = Finfo.DirectoryName
        Else
            Det_NomMarqueursMIDI = False
        End If
    End Function

    Function Det_NomFichExportExcel() As Boolean ' détermination du chemin fichier Accords *.mid
        Dim Finfo As FileInfo

        Det_NomFichExportExcel = True
        '
        If LangueIHM = "fr" Then
            SaveFileDialog1.Title = "Export Compogrid format Excel"
        Else
            SaveFileDialog1.Title = "Export Compogrid Excel Format"
        End If
        '
        SaveFileDialog1.Filter = "Excel Files (*.xls*)|*.xls"
        SaveFileDialog1.OverwritePrompt = True ' affichage message "Exise déjà ...."
        SaveFileDialog1.InitialDirectory = CheminFichierExportDoc
        SaveFileDialog1.FileName = ""
        '
        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            FichierExportDoc = SaveFileDialog1.FileName
            Finfo = My.Computer.FileSystem.GetFileInfo(FichierExportDoc)
            CheminFichierExportDoc = Finfo.DirectoryName
        Else
            Det_NomFichExportExcel = False
        End If
    End Function
    Function Det_NomFichExportHTM() As Boolean ' détermination du chemin fichier Accords *.mid
        Dim Finfo As FileInfo

        Det_NomFichExportHTM = True
        '
        If LangueIHM = "fr" Then
            SaveFileDialog1.Title = "Export Compogrid format HTML"
        Else
            SaveFileDialog1.Title = "Export Compogrid HTML Format"
        End If
        '
        SaveFileDialog1.Filter = "HTML Files (*.html)|*.html"
        SaveFileDialog1.OverwritePrompt = True ' affichage message "Exise déjà ...."
        SaveFileDialog1.InitialDirectory = CheminFichierExportDoc
        SaveFileDialog1.FileName = ""
        '
        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            FichierExportDoc = SaveFileDialog1.FileName
            Finfo = My.Computer.FileSystem.GetFileInfo(FichierExportDoc)
            CheminFichierExportDoc = Finfo.DirectoryName
        Else
            Det_NomFichExportHTM = False
        End If
    End Function

    Function Det_NomFichExportPDF() As Boolean ' détermination du chemin fichier Accords *.mid
        Dim Finfo As FileInfo

        Det_NomFichExportPDF = True
        '
        If LangueIHM = "fr" Then
            SaveFileDialog1.Title = "Export Compogrid format PDF"
        Else
            SaveFileDialog1.Title = "Export Compogrid PDF Format"
        End If
        '
        SaveFileDialog1.Filter = "PDF Files (*.pdf*)|*.pdf"
        SaveFileDialog1.OverwritePrompt = True ' affichage message "Exise déjà ...."
        SaveFileDialog1.InitialDirectory = CheminFichierExportDoc
        SaveFileDialog1.FileName = ""
        '
        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            FichierExportDoc = SaveFileDialog1.FileName
            Finfo = My.Computer.FileSystem.GetFileInfo(FichierExportDoc)
            CheminFichierExportDoc = Finfo.DirectoryName
        Else
            Det_NomFichExportPDF = False
        End If
    End Function
    Sub EnregistrerSous()
        Dim i, ind As Integer
        Dim a As String
        Dim enreg As New ENR ' le calcul des listes est effectué dans le contructeur de la classe ENR

        Try
            If Trim(CheminFichierEnreg) <> "" Then '  est déterminé dans Det_Nomfich '< ----- ici
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                'FileOpen(1, FichierEnreg, OpenMode.Output) ' Ouvre en écriture.
                Using sw As StreamWriter = New StreamWriter(FichierEnreg)
                    IndicateurEnreg = True
                    ' IHM ' le constructeur New de la classe ENR calcul toutes les listes ci-après
                    ' ***
                    sw.WriteLine(enreg.ListSyst)
                    sw.WriteLine(enreg.ListnumAcc)
                    sw.WriteLine(enreg.ListMarq)
                    sw.WriteLine(enreg.ListChiffAcc)
                    sw.WriteLine(enreg.ListRépet)
                    sw.WriteLine(enreg.ListMagnétos)
                    sw.WriteLine(enreg.ListRacines)
                    sw.WriteLine(enreg.ListMute)
                    sw.WriteLine(enreg.ListVolume) ' <-- liste volume
                    sw.WriteLine(enreg.ListPRG)
                    sw.WriteLine(enreg.ListMotifs)
                    sw.WriteLine(enreg.ListMotifs_text)
                    sw.WriteLine(enreg.ListDurées)
                    sw.WriteLine(enreg.ListOctaves)
                    sw.WriteLine(enreg.ListDyn)
                    sw.WriteLine(enreg.ListDyn2)
                    sw.WriteLine(enreg.ListDelay)
                    sw.WriteLine(enreg.ListDébutSouche)
                    sw.WriteLine(enreg.ListSouches)
                    sw.WriteLine(enreg.ListRetards)
                    sw.WriteLine(enreg.ListRadio1)
                    sw.WriteLine(enreg.ListRadio2)
                    sw.WriteLine(enreg.ListRadio3)
                    sw.WriteLine(enreg.ListAccent)
                    sw.WriteLine(enreg.ListNomduSon)
                    '
                    ' Table EventH
                    ' ************
                    For i = 1 To nbMesures
                        If Trim(TableEventH(i, 1, 1).Accord) <> "" Then
                            With TableEventH(i, 1, 1)
                                a = "EventH" + "," + Convert.ToString(i) + "," + Convert.ToString(.Ligne) + "," + .Position + "," + .Marqueur + "," _
                                + .Tonalité + "," + .Accord + "," + .Mode + "," + Convert.ToString(.Degré) + "," _
                                + Convert.ToString(.NumAcc) + "," + Convert.ToString(.NumMagnéto) + "," + Convert.ToString(.Répet) + "," + Convert.ToString(.Gamme) ' le paramètre racine est mis à jour à partir de listRacines dans enregistrer et dans Ouvrir2
                            End With
                            sw.WriteLine(a)
                            'EventHLigne(i)
                        End If
                    Next i
                    ' Table Zones
                    ' ***********
                    For i = 0 To NbZones
                        If TZone(i).DébutZ <> "---" Then
                            With TZone(i)
                                a = "Zone" +
                                "," + "Num " + Convert.ToString(i) +
                                "," + "DébutZ " + Trim(.DébutZ) + ' Valeur de début de la zone
                                "," + "TermeZ " + Trim(.TermeZ) + ' Valeur de la fin de la zone
                                "," + "NoteRacine " + Convert.ToString(.NoteRacine) + ' index dans le combolist
                                "," + "Racine " + Convert.ToString(.Racine) + ' text de la note racine dans le combolist
                                "," + "OctaveMoins1 " + Convert.ToString(.OctaveMoins1) +
                                "," + "Voix " + .VoixAsso_OctaveMoins1
                            End With
                            sw.WriteLine(a)
                            'ZoneLigne(i)
                        End If
                    Next
                    '
                    ' PERSO : Motifs Spécifiques
                    ' **************************
                    For i = 0 To TabControl1.TabPages.Count - 1
                        sw.WriteLine(ListSpecific(i))
                    Next
                    '
                    ' PianoRoll
                    ' *********
                    For i = 0 To nb_PianoRoll - 1
                        sw.WriteLine(listPIANOROLL(i).Enregistrer_ParamMélo())
                        sw.WriteLine(listPIANOROLL(i).Enregistrer_NotesMélo())
                        For ind = 0 To listPIANOROLL(i).PNbCtrls - 1
                            sw.WriteLine(listPIANOROLL(i).Enregistrer_Ctrls(ind))
                        Next ind
                        sw.WriteLine(listPIANOROLL(i).Enregistrer_ControlSys())
                        sw.WriteLine(listPIANOROLL(i).Enregistrer_CtrlPédale())
                        sw.WriteLine((listPIANOROLL(i).Enregistrer_ParamCalquesMIDI))
                        sw.WriteLine((listPIANOROLL(i).Enregistrer_CalquesMIDI))
                        sw.WriteLine((listPIANOROLL(i).Enregistrer_Assist1CTRP))

                    Next i
                    ' 
                    ' Automation
                    ' **********
                    For i1 = 0 To nb_PistesVar - 1
                        For i2 = 0 To nbCourbes - 1
                            sw.WriteLine(Automation1.Enregistrer_Ctrls(i1, i2))
                        Next i2
                        sw.WriteLine(Automation1.Enregistrer_ControlSys(i1))
                    Next i1
                    '
                    ' Drum edit
                    ' *********
                    sw.WriteLine(Drums.Sauv_LInst)       ' Instruments : liste des instruments
                    For i = 0 To NbDrumPrésets - 1       ' Préset :  notes de chaque préset
                        sw.WriteLine(Drums.Sauv_LPresetNotes(i))
                    Next
                    '
                    sw.WriteLine(Drums.Sauv_LTimeLPres)  ' TimeLine : Liste des Préset dans les mesures
                    '
                    sw.WriteLine(Drums.Sauv_NomPréset) ' nom de chaque préset de "A" à "H"

                    ' Table de mixage
                    ' ***************
                    a = Mix.AutorisVolumes   ' MIX Activation dans table de mixage (autorisation globale d'envoi des volumes)
                    sw.WriteLine(a)
                    a = Mix.AutresVol
                    sw.WriteLine(a)

                    ' FIN
                    ' ***
                    sw.WriteLine("FIN")
                End Using
            End If
            '
            Me.Text = FichierEnreg
            '
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        Catch ex As Exception
            If LangueIHM = "fr" Then
                MessageHV.PTitre = "Avertissement"
                MessageHV.PContenuMess = "Erreur interne : procédure 'EnregistrerSous' : " + ex.Message
                MessageHV.PTypBouton = "OK"
            Else
                MessageHV.PTitre = "Warning"
                MessageHV.PContenuMess = "Erreur interne : procédure 'EnregistrerSous' : " + ex.Message
                MessageHV.PTypBouton = "OK"
            End If
            Cacher_FormTransparents()
            MessageHV.ShowDialog()
            End
        End Try
    End Sub
    Function ListSpecific(numGrid As Integer) As String
        Dim i As Integer
        Dim j As Integer
        Dim listSpec As New List(Of String)
        Dim a, aa As String
        '
        listSpec.Clear()
        '
        listSpec.Add("ListSpecific")
        listSpec.Add(Convert.ToString(numGrid))
        '
        For j = 0 To MotifsPerso.Item(numGrid).Cols - 1
            a = ""
            For i = 0 To MotifsPerso.Item(numGrid).Rows - 1
                aa = Trim(MotifsPerso.Item(numGrid).Cell(i, j).Text)
                If aa = "*" Then ' aa = "-" Or Or aa = "+"
                    a = aa + ";" + Convert.ToString(i) + ";" + Convert.ToString(j)  ' * : début de note / i : ligne / j : colonne
                    listSpec.Add(a)
                End If
                If (MotifsPerso.Item(numGrid).Cell(i, j).BackColor = Couleur_Notes) Then ' And (Trim(a) = ""
                    a = "c" + ";" + Convert.ToString(i) + ";" + Convert.ToString(j) ' c / i : ligne / j : colonne     "c" comme couleur --> toutes les cellules colorisées avec Couleur_Notes
                    listSpec.Add(a)
                End If
            Next i
        Next j
        ListSpecific = String.Join(",", listSpec.ToArray())
    End Function




    Sub EventHLigne(m As Integer)
        Dim a As String = ""
        ' Remarque : Tonalité est écrit 2 fois, une fois pour le champ Tonalité et une 2 fois poir le champ Gamme (Gamme =Tonalité pour Hyper Arp).
        With TableEventH(m, 1, 1)
            a = "EventH" + "," + Convert.ToString(m) + "," + Convert.ToString(.Ligne) + "," + .Position + "," + .Marqueur + "," _
                + .Tonalité + "," + .Accord + "," + .Mode + "," + Convert.ToString(.Degré) + "," _
                + Convert.ToString(.NumAcc) + "," + Convert.ToString(.NumMagnéto) + "," + Convert.ToString(.Répet) + "," + Convert.ToString(.Gamme)
        End With
        WriteLine(1, a)
    End Sub
    Sub ZoneLigne(i As Integer)
        Dim a As String
        With TZone(i)
            a = "Zone" +
                "," + "Num " + Convert.ToString(i) +
                "," + "DébutZ " + Trim(.DébutZ) + ' Valeur de début de la zone
                "," + "TermeZ " + Trim(.TermeZ) + ' Valeur de la fin de la zone
                "," + "NoteRacine " + Convert.ToString(.NoteRacine) + ' index dans le combolist
                "," + "Racine " + Convert.ToString(.Racine) + ' text de la note racine dans le combolist
                "," + "OctaveMoins1 " + Convert.ToString(.OctaveMoins1) +
                "," + "Voix " + .VoixAsso_OctaveMoins1
        End With
        WriteLine(1, a)
    End Sub

    Function Det_PAN(ind As Integer) As Integer
        Det_PAN = 2

        If PisteRadio1.Item(ind).Checked Then Det_PAN = 1
        If PisteRadio2.Item(ind).Checked Then Det_PAN = 2
        If PisteRadio3.Item(ind).Checked Then Det_PAN = 3
    End Function

    Function Det_TonalitéDansCombobox1(tona As String) As Integer
        Dim i As Integer
        Select Case tona
            Case "C Maj"
                i = 0
            Case "C# Maj"
                i = 1
            Case "D Maj"
                i = 2
            Case "D# Maj"
                i = 3
            Case "E Maj"
                i = 4
            Case "F Maj"
                i = 5
            Case "F# Maj"
                i = 6
            Case "G Maj"
                i = 7
            Case "G# Maj"
                i = 8
            Case "A Maj"
                i = 9
            Case "A# Maj"
                i = 10
            Case "B Maj"
                i = 11
        End Select
        '
        Det_TonalitéDansCombobox1 = i
    End Function
    'Function Det_ColGrid2Grid3() As Integer
    '    Select Case DerGridCliquée
    '        Case GridCours.Grid2
    '        Case GridCours.Grid3
    '    End Select
    '
    '   End Function
    Sub CouperJouerAccord()
        Dim i, j, n As Integer

        'Tempo_StopJeuAccord.Enabled = False
        'Tempo_StopJeuAccord.Stop()I deractivecell.col
        Try
            ' Fin jeu d'un accord de grid2
            ' ****************************
            If AccordAEtéJoué = True Then
                For i = 0 To UBound(AccordJouerPiano.Notes)
                    If AccordJouerPiano.Notes(i) <> -1 Then
                        j = AccordJouerPiano.Notes(i)
                        'LabelPiano.Item(j).BackColor = AccordJouerPiano.OldBackColor(i)
                        'LabelPiano.Item(j).Text = ""
                    End If
                Next i
                AccordAEtéJoué = False
                '
                For i = 0 To UBound(AccordJouerPiano.Notes)
                    If AccordJouerPiano.Notes(i) <> -1 Then
                        n = AccordJouerPiano.Notes(i)
                        'SortieMidi.Item(ChoixSortieMidi).SendNoteOff(Convert.ToByte(CanalEcoute.Text) - 1, n, 64) ' 64 = vélocité de relachement
                    End If
                Next i
                SortieMidi.Item(ChoixSortieMidi).SilenceAllNotes()
                '
            End If
        Catch ex As Exception

            If LangueIHM = "fr" Then
                MessageHV.PTitre = "Avertissement"
                MessageHV.PContenuMess = "Détection d'un erreur MIDI dans procédure : " + "CouperJouerAccord" + Constants.vbCrLf + "Message  : " + ex.Message
                MessageHV.PTypBouton = "OK"
            Else
                MessageHV.PTitre = "Warning"
                MessageHV.PContenuMess = "Détection d'un erreur MIDI dans procédure : " + "CouperJouerAccord" + Constants.vbCrLf + "Message  : " + ex.Message
                MessageHV.PTypBouton = "OK"
            End If
            Cacher_FormTransparents()
            MessageHV.ShowDialog()
            End
        End Try
    End Sub
    Sub CouperJouerAccord2()
        MIDIReset()
    End Sub
    Sub ToutesNotesOff()
        Dim i, j As Integer
        '
        If Not (SortieMidi.Item(ChoixSortieMidi).IsOpen) Then
            SortieMidi.Item(ChoixSortieMidi).Open()
        End If
        '
        For i = 0 To nb_TotalPistes - 1  'Arrangement1.Nb_PistesMidi
            For j = 0 To 127
                SortieMidi.Item(ChoixSortieMidi).SendNoteOn(i, j, 0)
                SortieMidi.Item(ChoixSortieMidi).SendNoteOff(i, j, 0)
            Next j
        Next i
    End Sub






    Sub CouperAccord()
        Dim i, j As Integer

        'Tempo_StopJeuAccord.Enabled = False
        'Tempo_StopJeuAccord.Stop() 'I deractivecell.col
        Try
            ' Fin jeu d'un accord de grid2
            ' ****************************
            If AccordAEtéJoué = True Then
                For i = 0 To UBound(AccordJouerPiano.Notes)
                    If AccordJouerPiano.Notes(i) <> -1 Then
                        j = AccordJouerPiano.Notes(i)
                        LabelPiano.Item(j).BackColor = AccordJouerPiano.OldBackColor(i)
                        LabelPiano.Item(j).ForeColor = Color.Blue
                    End If
                Next i
                AccordAEtéJoué = False

                SortieMidi.Item(ChoixSortieMidi).SilenceAllNotes()

                Fermer_MIDI()

                '
            End If
        Catch ex As Exception
            If LangueIHM = "fr" Then
                MessageHV.PTitre = "Avertissement"
                MessageHV.PContenuMess = "Détection d'un erreur dans procédure : " + "CouperAccord" + Constants.vbCrLf + "Message  : " + ex.Message
                MessageHV.PTypBouton = "OK"
            Else
                MessageHV.PTitre = "Warning"
                MessageHV.PContenuMess = "Détection d'un erreur dans procédure : " + "CouperAccord" + Constants.vbCrLf + "Message  : " + ex.Message
                MessageHV.PTypBouton = "OK"
            End If
            Cacher_FormTransparents()
            MessageHV.ShowDialog()
            End
        End Try
    End Sub
    Private Sub NouveauToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NouveauToolStripMenuItem.Click


    End Sub
    Function NouveauProjet() As Integer
        Dim i As Integer
        'Me.Visible = False
        'If EcritUneFois = True Then
        i = NouveauConfirm2() ' la question est : voulez-vous enregistrez préalablement votre projet actuel ?
        Select Case i
            Case DialogResult.Yes
                Enregistrer()
                Nouv("4/4")
            Case DialogResult.No
                Nouv("4/4")
            Case DialogResult.Cancel ' on ne fait rien
        End Select
        'Else
        'Nouv("4/4")
        'End If
        Maj_DelockCell()
        '
        ' Diverses Position Initiales
        ' ***************************e
        'Me.Size = New Size(1395, 500) '682 Dimensions de la fenêtre principale
        SplitContainer7.SplitterDistance = PosSystem ' PosStandard ' position du splitter central
        Panel2.Visible = False
        '
        Me.Visible = True
        Return i
    End Function
    Function NouveauConfirm() As Integer
        Dim message As String
        Dim titre As String

        Cacher_FormTransparents()

        If CheckBox1.Checked Then
            NouveauConfirm = vbNo
            If LangueIHM = "fr" Then
                message = "Voulez-vous sauvegarder le projet actuel avant d'en ouvrir un nouveau ?"
                titre = "ALARME POUR SAUVEGARDE DU PROJET COURANT"
            Else
                message = "Do you want to save the current project before opening a new one ?"
                titre = "WARNING FOR SAVING CURRENT PROJECT"
            End If

            Cacher_FormTransparents()
            NouveauConfirm = MessageBox.Show(message, titre, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
        Else
            NouveauConfirm = vbNo
        End If

        ' Valeurs de retour de messagebox
        ' *******************************
        'vbCancel   2	Cancel
        'vbAbort    3	Abandonner
        'vbRetry    4	Réessayer
        'vbIgnore   5	Ignore
        'vbYes      6	Oui
        'vbNo       7	Non

    End Function

    Function NouveauConfirm2() As Integer
        Dim message As String
        Dim titre As String

        Cacher_FormTransparents()

        If CheckBox1.Checked Then
            NouveauConfirm2 = vbNo
            If LangueIHM = "fr" Then
                message = "Sauvegarder le projet avant de le fermer ?"
                titre = "ALARME POUR SAUVEGARDE DU PROJET COURANT"
            Else
                message = "Save the project before closing it ?"
                titre = "WARNING FOR SAVING CURRENT PROJECT"
            End If

            'Cacher_FormTransparents()
            NouveauConfirm2 = MessageBox.Show(message, titre, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
        Else
            NouveauConfirm2 = vbNo
        End If
    End Function


    Function CalculValMétronome(PositionMétro As Integer) As Integer
        Dim BaseMétronome As Integer
        CalculValMétronome = 0
        Select Case Dénominateur
            Case 4
                BaseMétronome = 4
                CalculValMétronome = PositionMétro + BaseMétronome
            Case 8
                BaseMétronome = 6
                Select Case Numérateur
                    Case 6, 9, 12
                        CalculValMétronome = PositionMétro + BaseMétronome
                    Case 7
                        If Beat78 Then
                            BaseMétronome = 6
                            CalculValMétronome = PositionMétro + BaseMétronome
                            Beat78 = False
                        Else
                            BaseMétronome = 8
                            CalculValMétronome = PositionMétro + BaseMétronome
                            Beat78 = True
                        End If
                End Select
        End Select
    End Function
    Function Det_NbTempsSurCycle() As Integer ' fonction utilisée par le métronome
        Dim k As Integer

        Det_NbTempsSurCycle = 0
        k = nbMesures
        If Répéter.Checked Then
            k = (Terme.Value - Début.Value) + 1 ' calcul du nombre de mesures par cycle
        End If
        '
        Select Case Dénominateur
            Case 4
                Det_NbTempsSurCycle = k * Numérateur
            Case 8

                Select Case Numérateur
                    Case 6, 9, 12
                        Det_NbTempsSurCycle = k * (Numérateur / 3)
                    Case 7
                        Det_NbTempsSurCycle = (k * 6) + (k * 8)
                End Select

        End Select
    End Function
    Function Maj_NotesCommunes(Renv As String, Zone As Integer) As String
        Dim i, j, lg1, lg2 As Integer
        Dim tbl() As String
        Dim Racine As String
        Dim Octave As Integer
        Dim indRacine As Integer
        Dim CalcOctave As Integer
        Dim Retour As String
        Dim NombreNotesAccord As Integer
        Dim a As String

        Retour = ""
        Maj_NotesCommunes = ""
        '
        ' Séparation Note / Octave de la racine pour le calcul
        ' ****************************************************
        'a = Trim(NoteRacine.Text)
        If Zone <> -1 Then
            a = Trim(TZone(Zone).Racine)
        Else
            a = Trim(THorsZone.Racine)
        End If
        '
        If Mid(a, Len(a) - 1, 1) <> "-" Then
            lg1 = Len(a) - 1
            lg2 = 1
        Else
            lg1 = Len(a) - 2
            lg2 = 2
        End If
        '
        Racine = Mid(a, 1, lg1) '
        Octave = Microsoft.VisualBasic.Right(a, lg2)
        '
        tbl = Split(Renv, "-")
        NombreNotesAccord = UBound(tbl) + 1
        '
        Init_Tessiture()
        '
        ' Maj des notes de l'accord dans le tableau Tessiture
        ' ***************************************************
        For i = 0 To UBound(tbl)
            For j = 0 To UBound(Tessiture)
                If tbl(i) = Tessiture(j).NoteTessiture Then
                    Tessiture(j).NoteAccord = tbl(i)
                End If
            Next j
        Next i
        '
        ' Détermination de l'index de la racine dans le tableau Tessiture
        ' ***************************************************************
        For i = 0 To UBound(Tessiture)
            If Racine = Tessiture(i).NoteTessiture Then
                indRacine = i
                Exit For
            End If
        Next i
        '
        ' Détermination du renversement avec l'octave
        ' *******************************************

        CalcOctave = indRacine
        For i = indRacine To UBound(Tessiture)
            If Tessiture(i).NoteAccord <> "" Then
                Retour = Trim(Retour + " " + Tessiture(i).NoteAccord + Trim(Str(Octave)))
            End If
            ' pour détermination d'un changement d'octave
            CalcOctave = CalcOctave + 1
            If CalcOctave > 11 Then
                Octave = Octave + 1 ' changement d'octave
                CalcOctave = 0
            End If
            '
            tbl = Split(Retour)
            '
        Next i
        Retour = ""
        For i = 0 To NombreNotesAccord - 1
            Retour = Retour + " " + tbl(i)
        Next
        '
        Maj_NotesCommunes = Trim(Retour)
    End Function

    Sub Init_Tessiture()
        Dim i As Integer

        For i = 0 To 31
            Tessiture(i).NoteTessiture = TabNotesD(i) ' la cacul de l'autovoicing se fait toujours en #
            Tessiture(i).NoteAccord = ""
        Next

    End Sub

    Function Det_DerNoteCycle(nMCycle As Integer) As String
        Dim m, t, ct As Integer
        Dim sortir As Boolean = False
        Dim a As String
        Dim nM As Integer
        Dim PFIN As Integer

        '
        Det_DerNoteCycle = ""
        For m = UBound(TableEventH, 1) To 0 Step -1
            a = ""
            For t = UBound(TableEventH, 2) To 0 Step -1
                For ct = UBound(TableEventH, 3) To 0 Step -1
                    If TableEventH(m, t, ct).Ligne <> -1 Then
                        nM = m '+ ((EndMeasureNumber.Value - 2))
                        PFIN = Terme.Value
                        If PFIN < nM Then
                            Det_DerNoteCycle = Trim(Str(PFIN + nMCycle) + " " + Trim(Str(t)) + " " + Trim(Str(ct)))
                        Else
                            Det_DerNoteCycle = Trim(Str(nM + nMCycle) + " " + Trim(Str(t)) + " " + Trim(Str(ct)))
                        End If
                        sortir = True
                        Exit For
                    End If
                Next ct
                If sortir Then
                    Exit For
                End If
            Next t
            If sortir Then
                Exit For
            End If
        Next m
        '
    End Function
    Function Det_DerNoteCycle2() As String
        Dim m As Integer
        Dim sortir As Boolean = False
        '
        Det_DerNoteCycle2 = ""
        For m = nbMesures To 0 Step -1
            If TableEventH(m, 1, 1).Ligne <> -1 Then
                Det_DerNoteCycle2 = Trim(Str(m) + "." + "1" + "." + "1")
                sortir = True
                Exit For
            End If
        Next m
        '
    End Function
    Function Det_DerNoteCycle3() As String
        Dim m, mm, t, ct As Integer
        Dim sortir As Boolean = False
        Dim a As String

        '
        Det_DerNoteCycle3 = ""
        'If Répéter.Checked = True Then
        mm = Terme.Value
        'Else
        'mm = UBound(TableEventH, 1)
        'End If
        For m = mm To 0 Step -1
            a = ""
            For t = UBound(TableEventH, 2) To 0 Step -1
                For ct = UBound(TableEventH, 3) To 0 Step -1
                    If TableEventH(m, t, ct).Ligne <> -1 Then
                        Det_DerNoteCycle3 = Trim(Trim(Str(m) + "." + Trim(Str(t)) + "." + Trim(Str(ct))))
                        sortir = True
                        Exit For
                    End If
                Next ct
                If sortir Then
                    Exit For
                End If
            Next t
            If sortir Then
                Exit For
            End If
        Next m

    End Function

    Sub Création_MidiFile(CheminFichierText As String, NbPistes As Integer)
        Dim Ligne As String
        Dim tbl As Object
        Dim i As Integer = 0

        '   
        ' création du fichier MIDI
        ' ************************
        Dim Midifile1 As New MidifileX(96, NbPistes, nbMesures)
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
            '
            If LangueIHM = "fr" Then
                MessageHV.PTitre = "Avertissement"
                MessageHV.PContenuMess = "Erreur interne : procédure Création_MidiFile : " + " i =" + Str(i) + "--" + ex.Message
                MessageHV.PTypBouton = "OK"
            Else
                MessageHV.PTitre = "Warning"
                MessageHV.PContenuMess = "Internal Error : procédure Création_MidiFile : " + ex.Message
                MessageHV.PTypBouton = "OK"
            End If
            MessageHV.ShowDialog()
            End
        End Try


    End Sub
    Sub RAZ_Marqueurs_dans_Tableventh()
        Dim m, t, ct As Integer
        For m = 1 To UBound(TableEventH, 1)
            For t = 0 To UBound(TableEventH, 2)
                For ct = 0 To UBound(TableEventH, 3)
                    TableEventH(m, t, ct).Marqueur = ""
                Next ct
            Next t
        Next m
    End Sub


    Function Taille_Mesure() As Integer ' taille d'une mesure en double croches en fonction de la métrique
        Dim tbl() As String
        Dim Numérateur As Integer
        Dim Dénominateur As Integer

        tbl = Split(Métrique.Text, "/")
        Numérateur = Trim(tbl(0))
        Dénominateur = Trim(tbl(1))

        Taille_Mesure = 4
        Select Case Dénominateur
            Case "4"
                Taille_Mesure = 4 * Val(Numérateur)
            Case "8"
                Taille_Mesure = 2 * Val(Numérateur)
        End Select
    End Function









    Private Delegate Function MyDelegate() As Midi.CallbackMessage.CallbackType()
    Sub EnvoiMess() 'As Midi.CallbackMessage.CallbackType()
        VarCallBack = "Passage"
    End Sub


    Private Sub Fermer_MIDI()
        Dim i As Integer
        '
        For i = 0 To SortieMidi.Count - 1
            If SortieMidi.Item(i).IsOpen Then
                If Horloge1.IsRunning Then
                    Horloge1.Stop()
                    Horloge1.Reset()
                End If
                'SortieMidi.Item(i).Close() --> supprimé car le "Close" remet à 0 le N° de programme GS
            End If
        Next
        '
    End Sub


    '    Tableau théorique DES N° DE NOTES MIDI
    '    ********************************************
    '     notes C   C#  D   D#  E   F   F#  G   G#  A   A#  B
    ' I oct+----------------------------------------------
    ' 0 -2 !    0   1   2   3   4   5   6   7   8   9   10  11 
    ' 1 -1 !    12  13  14  15  16  17  18  19  20  21  22  23 
    ' 2  0 !    24  25  26  27  28  29  30  31  32  33  34  35 
    ' 3  1 !    36  37  38  39  40  41  42  43  44  45  46  47 
    ' 4  2 !    48  49  50  51  52  53  54  55  56  57  58  59
    ' 5  3 !    60  61  62  63  64  65  66  67  68  69  70  71
    ' 6  4 !    72  73  74  75  76  77  78  79  80  81  82  83
    ' 7  5 !    84  85  86  87  88  89  90  91  92  93  94  95
    ' 8  6 !    96  97  98  99  100 101 102 103 104 105 106 107 
    ' 9  7 !    108 109 110 111 112 113 114 115 116 117 118 119
    ' 10 8 !    120 121 122 123 124 125 126 127


    ' Application au dictionnaire DicoNotes
    ' *************************************
    ' -------+-----------------------------------------------------------------
    ' Clefs  +  Valeurs
    ' -------+------------------------------------------------------------------
    ' c      ! Classe TT0  :  0 12 24 36 48 50 72 84 96  108 120 - > tous les "c"
    ' c#     ! Classe TT1  :  1 13 25 37 49 51 73 85 97  109 121 - > tous les "c#"
    ' d      ! Classe TT2  :  2 14 26 38 50 52 74 86 98  110 122 - > tous les "d"
    ' d#     ! Classe TT3  :  3 15 27 39 51 53 75 87 99  111 123 - > tous les "d#"
    ' e      ! Classe TT4  :  4 16 28 40 52 54 76 88 100 101 124 - > tous les "e"
    ' f      ! etc...
    ' f#     ! 
    ' g      !
    ' g#     !
    ' a      !
    ' a#     !
    ' b      ! Classe TT11 : 11 23 35 47 58 71 83 95 107 119 -- 

    Sub Aff_Gamme(Gamme As String) ' Colorisation du piano avec les notes de la gamme
        Dim i, j, k, r As Integer
        Dim NotesGamme As String
        Dim tbl() As String
        Dim t As New TT ' Class TT Public Toctave(0 To 11) As Integer End Class
        Dim Tonique As String = ""
        Dim a, G, note, SigneClef As String
        '
        '
        Maj_TabNotes_Minus("#") ' remarque : les notes des gammes sont toujours calculées en # pour coloriser le piano
        tbl = Split(Gamme, " ")
        a = tbl(0)
        SigneClef = Trim(Det_Clef(a))

        RAZ_CouleursPiano()
        '
        If Langue = "fr" Then
            Gamme = TradAcc_LatAngl(Gamme)
        End If
        '
        G = Trad_AccordEn_D(Trim(Gamme))
        NotesGamme = Det_NotesGammes(G)
        tbl = Split(NotesGamme)
        '
        Tonique = Trim(tbl(0))
        ' label d'affichage nom de la gamme en dessous des réglages
        ' *********************************************************
        '
        For i = 0 To UBound(tbl)
            t = DicoNotes(tbl(i)) ' Public DicoNotes As New Dictionary(Of String, TT)
            For j = 0 To UBound(t.Toctave)
                k = t.Toctave(j) 'k est e N° midi de la note  et j est l'octave
                If k < 128 Then
                    If LabelPiano.Item(k).Height = HauteurTouche Then
                        LabelPiano.Item(k).BackColor = Color.LightGreen
                        'LabelPianoMidiIn.Item(k).BackColor = Color.LightGreen
                    Else
                        LabelPiano.Item(k).BackColor = Color.ForestGreen
                        LabelPianoMidiIn.Item(k).BackColor = Color.ForestGreen
                    End If '
                    LabelPiano.Item(k).ForeColor = Color.Black
                    note = Trad_NoteEnD(Trim(tbl(i))) ' Trad_NoteEnD
                    note = Trim(LCase(note))
                    note = Trim(note + Det_nOctave(j))
                    '
                    If SigneClef = "b" Then
                        r = ListNotesd.IndexOf(note)
                        note = ListNotesb(r) ' traduction de la note en 'b' au cas où cette note serait la tonique, dans ce cas elle doit être écrite sur la touche du piano avec la bonne clef
                    End If
                    If tbl(i) = Tonique Then ' ecriture de la tonique dans le piano
                        If LabelPiano.Item(k).Height = HauteurTouche Then
                            LabelPiano.Item(k).ForeColor = Color.Black
                        Else
                            LabelPiano.Item(k).ForeColor = Color.White
                        End If
                        LabelPiano.Item(k).Text = "T" 'Trim(note)
                    End If
                End If
            Next j
        Next i
        '
    End Sub
    Function Det_nOctave(Ind As Integer) As String
        Select Case Ind
            Case 0
                Det_nOctave = "-2"
            Case 1
                Det_nOctave = "-1"
            Case Else
                Det_nOctave = Trim(Str(Ind - 2))
        End Select
    End Function
    Sub RAZ_CouleursPiano()
        Dim i, j, k As Integer
        For i = 0 To 127 Step 12
            For j = 0 To 11
                k = i + j
                If k = 128 Then Exit For
                '
                LabelPiano.Item(k).Text = ""
                Select Case j
                    Case 1, 3, 6, 8, 10
                        LabelPiano.Item(k).BackColor = Color.Black
                        LabelPianoMidiIn.Item(k).BackColor = Color.Black

                        'LabelPiano.Item(k).BringToFront()
                    Case Else
                        LabelPiano.Item(k).BackColor = Color.White
                        LabelPianoMidiIn.Item(k).BackColor = Color.White
                End Select
            Next j

        Next i
    End Sub
    Private Sub Début_ValueChanged_1(sender As Object, e As EventArgs)
        If EnChargement = False Then
            If Début.Value > Terme.Value Then
                Début.Value = Début.Value - 1
            End If
        End If
        TextBox2.Text = Calcul_Durée()
    End Sub

    Private Sub Terme_ValueChanged(sender As Object, e As EventArgs) Handles Terme.ValueChanged
        'If EnChargement = False Then
        'If Terme.Value < Début.Value Then
        'Terme.Value = Terme.Value + 1
        'End If
        'End If
        Cohérence_Délim()
        Calcul_Durée()
        TextBox2.Text = Calcul_Durée()
    End Sub
    Function Calcul_Durée() As String
        Dim r As Double
        Dim nbbeat As Integer
        Dim temps As Double

        TextBox2.BackColor = Color.WhiteSmoke
        TextBox2.ForeColor = Color.Red
        Label3.BackColor = Color.WhiteSmoke
        Label3.ForeColor = Color.Red
        '
        r = Convert.ToInt16(Tempo.Value) / 60
        nbbeat = Convert.ToInt16((Terme.Value - Début.Value) + 1) * 4
        temps = nbbeat / r
        temps = Math.Truncate(temps)
        Return temps.ToString
    End Function

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs)
        'NotesCommunes2.Checked = NotesCommunes.Checked
    End Sub
    Private Sub ComboMidiOut_SelectedIndexChanged_1(sender As Object, e As EventArgs)
        Dim a As String
        '
        '
        If EnChargement = False Then
            ChoixSortieMidi = ComboMidiOut.SelectedIndex
            '
            a = Trim(Str(ChoixSortieMidi))
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CalquesMIDI\HyperArp\Préférences", "ChoixSortieMIDI", a)
        End If
        '
    End Sub
    Private Sub CopierToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopierToolStripMenuItem.Click, CopierToolStripMenuItem1.Click

        Select Case OngletCours_Edition
            Case 0 ' HyperArp
                If Grid2.SelectionMode = SelectionModeEnum.ByColumn Then
                    Select Case OngletCours_HyperARP

                        Case 0, 1, 2, 3, 127 ' Onglets Gammes d'Accords, Progression d'Accords,Zone de Voicing, Perso,Time Line des accords
                            Grid2.Range(1, 1, 1, Grid2.Cols - 1).Locked = False
                            CopierClip()
                            Grid2.Selection.CopyData()
                            Calcul_AutoVoicingZ()
                            Grid2.Range(1, 1, 1, Grid2.Cols - 1).Locked = True
                        Case Else
                            Dim i1 As Integer = Automation1.PNcanal
                            Dim i2 As Integer = Automation1.PNCourbe
                            Automation1.Canaux(i1).GridCourbes(i2).AutoRedraw = False
                            Automation1.Canaux(i1).GridCourbes(i2).Selection.CopyData()
                            Automation1.Canaux(i1).GridCourbes(i2).AutoRedraw = True
                            Automation1.Canaux(i1).GridCourbes(i2).Refresh()

                    End Select
                End If


            Case 1, 2, 3, 5, 6, 7, 8 ' piano roll
                iii = OngletCours_Edition - 1
                If OngletCours_Edition > 3 Then iii = OngletCours_Edition - 2
                With listPIANOROLL.Item(iii)
                    If .Orig_PianoR.Orig1 = OrigPianoCourbe.Piano Then
                        .CopierData()
                    Else
                        Dim ind As Integer = .Orig_PianoR.N_Courbe
                        .GridCourbes.Item(ind).Selection.CopyData()
                    End If
                End With
            Case 4 ' Drum edit
                If Drums.PGridOrigine = "Grid1" Then
                    If Drums.Grid1.Selection.FirstCol > 3 And Drums.Grid1.Selection.FirstRow > 0 Then
                        Drums.Grid1.Selection.CopyData()
                        'Drums.DrumCopy()
                    End If
                End If
                If Drums.PGridOrigine = "Grid2" Then
                    Drums.Grid2.Range(0, 0, Drums.Grid2.Rows - 1, Drums.Grid2.Cols - 1).Locked = False ' dévéroulluer seulement la lignes des lettres de pattern (A,B
                    Drums.Grid2.Selection.CopyData()
                    Drums.Grid2.Range(0, 0, Drums.Grid2.Rows - 1, Drums.Grid2.Cols - 1).Locked = True ' dévéroulluer grid2
                End If
            Case Else ' Courbes
                i = 0

        End Select

    End Sub

    Private Sub CouperToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CouperToolStripMenuItem.Click, CouperToolStripMenuItem1.Click
        Dim a, b, c, d As String
        Dim i, ii, iii As Integer

        Select Case OngletCours_Edition
            Case 0 ' HyperArp
                If Grid2.SelectionMode = SelectionModeEnum.ByColumn Then
                    Select Case OngletCours_HyperARP
                        Case 0, 1, 2, 3, 127 ' Onglets Gammes d'Accords, Progression d'Accords,Zone de Voicing, Perso, Time Line des accords
                            If Grid2.Selection.FirstCol > 1 Then
                                Grid2.Range(1, 1, Grid2.Rows - 1, Grid2.Cols - 1).Locked = False
                                SAUV_Annuler(Grid2.Selection.FirstCol, Grid2.Selection.LastCol)
                                CopierClip()
                                Grid2.Selection.CutData() ' efface tout : text et format
                                '
                                Grid2.Refresh()
                                CouperClip() ' ici, on rétablit le format (sans le text)
                                Grid2.Refresh()
                                Maj_Répétition2(Grid2.Selection.LastCol)
                                Maj_Nligne()
                                Grid2.Range(1, 1, Grid2.Rows - 1, Grid2.Cols - 1).Locked = True
                                Calcul_AutoVoicingZ()
                                ' Mise à jour entête entête des piano roll
                                ' ****************************************
                                a = Trim(Det_ListAcc())
                                b = Trim(Det_ListGam())
                                c = Trim(Det_ListMarq())
                                d = Trim(Det_ListTon())
                                For i = 0 To nb_PianoRoll - 1
                                    If PIANOROLLChargé(i) = True Then
                                        listPIANOROLL(i).PListAcc = a 'Det_ListAcc()
                                        listPIANOROLL(i).PListGam = b 'Det_ListGam()
                                        listPIANOROLL(i).PListMarq = c 'Det_ListMarq()
                                        listPIANOROLL(i).PListTon = d
                                        listPIANOROLL(i).Clear_graphique2() ' <---
                                        listPIANOROLL(i).F1_Refresh()
                                        listPIANOROLL(i).Maj_CalquesMIDI()
                                    End If
                                Next i
                            Else
                                Dim message, titre As String
                                If LangueIHM = "fr" Then
                                    message = "Il n'est pas possible de supprimer la première colonne."
                                    titre = "Avertissement"
                                Else
                                    message = "It is not possible to delete the first column."
                                    titre = "Avertissement"
                                End If
                                Cacher_FormTransparents()
                                MessageBox.Show(message, titre, MessageBoxButtons.OK, MessageBoxIcon.Information)
                            End If
                            ' Mise à jour de Grid2 de drumedit (timeline drumedit)
                            Automation1.PListAcc = Det_ListAcc()
                            Automation1.PListMarq = Det_ListMarq()
                            Automation1.F4_Refresh()

                            Drums.PListAcc = Det_ListAcc()
                            Drums.PListMarq = Det_ListMarq()
                            Drums.F2_Refresh()
                        Case Else ' Onglet Automation
                            Dim i1 As Integer = Automation1.PNcanal
                            Dim i2 As Integer = Automation1.PNCourbe
                            Automation1.Canaux(i1).GridCourbes(i2).AutoRedraw = False
                            Automation1.Canaux(i1).GridCourbes(i2).Selection.CutData()
                            Automation1.Canaux(i1).GridCourbes(i2).AutoRedraw = True
                            Automation1.Canaux(i1).GridCourbes(i2).Refresh()
                    End Select
                End If

            Case 1, 2, 3, 5, 6, 7, 8 ' pianoroll
                'listPIANOROLL.Item(OngletCours_Edition - 1).CouperData()
                iii = OngletCours_Edition - 1
                If OngletCours_Edition > 3 Then iii = OngletCours_Edition - 2
                If listPIANOROLL.Item(iii).Orig_PianoR.Orig1 = OrigPianoCourbe.Piano Then ' edition dans pianoroll
                    listPIANOROLL.Item(iii).CouperData()
                Else
                    Dim ind As Integer = listPIANOROLL.Item(iii).Orig_PianoR.N_Courbe ' édition dans une courbe du pianoroll
                    listPIANOROLL.Item(iii).GridCourbes.Item(ind).Selection.CutData()
                End If
            Case 4 ' Drums
                If Drums.PGridOrigine = "Grid1" Then
                    If Drums.Grid1.Selection.FirstCol > 3 And Drums.Grid1.Selection.FirstRow > 0 Then
                        Drums.DéVérouillerGrid1()    ' les cellules ont été vérrouillées dans drums.grid1_Keydown 
                        Drums.Grid1.AutoRedraw = False
                        Drums.Sauv_ClipAnnuler()
                        '
                        Drums.Grid1.Selection.CutData()
                        'Drums.DrumCopy()
                        'Drums.DrumCut()
                        '
                        Drums.DessinDiv()
                        ii = Drums.Det_NumPréset2()
                        Drums.Maj_Préset(ii)
                        Drums.Grid1.AutoRedraw = True
                        Drums.Grid1.Refresh()
                        Drums.VérouillerGrid1()     ' les cellules seront déverrouillées à la fin de Drums.Grid1_KeyUp
                    End If
                End If
        End Select
    End Sub
    Private Sub CollerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CollerToolStripMenuItem.Click, CollerToolStripMenuItem1.Click
        Dim i, iii, ind As Integer
        Dim a, b, c, d As String
        Select Case OngletCours_Edition
            Case 0 ' HyperArp

                Select Case OngletCours_HyperARP

                    Case 0, 1, 2, 3, 127 ' Onglets Gammes d'Accords, Progression d'Accords,Zone de Voicing, Perso, Time Line des Accords
                        If Grid2.SelectionMode = SelectionModeEnum.ByColumn Then
                            If ClipEdit.Count <> 0 Then
                                ColAncienDerAcc = Det_NumDerAccord() ' Grid2.Cell(3, d).Text ' N° Ancien dernier accord
                                ReptAncienDerAcc = Trim(Grid2.Cell(3, ColAncienDerAcc).Text) ' Répétition ancien dernier accord

                                Grid2.Range(1, 1, Grid2.Rows - 1, Grid2.Cols - 1).Locked = False
                                If CollerClip() Then
                                    Grid2.Selection.PasteData() ' 
                                    '
                                    ' mise à jour des couleurs en fonction des tonalités
                                    Dim m As Integer = Grid2.Selection.FirstCol ' PasteData() colle déjà les couleurs mais si la sélection contient le 1er accord en rouge, on ne souhaite paas que la couleur rouge apparaissent dans l'élément collé
                                    Dim LstCol As Integer = m + (ClipEdit.Count - 1)
                                    Dim tbl() As String
                                    For j = m To LstCol
                                        a = TableEventH(j, 1, 1).Tonalité '
                                        If j <> 1 Then

                                            If Trim(a) <> "" Then ' cas où le CLip contient des colonnes vides
                                                a = Det_RelativeMajeure2(a) ' si la tonalité est mineure alors on affiche la couleur de la relative majeure
                                                tbl = Split(a)
                                                Grid2.Cell(2, j).BackColor = DicoCouleur.Item(Trim(tbl(0))) ' la couleur de l'accord est fonction de la tonalité
                                                Grid2.Cell(2, j).ForeColor = DicoCouleurLettre.Item(tbl(0))
                                                Grid2.Cell(11, j).BackColor = DicoCouleur.Item(Trim(tbl(0))) ' la couleur de la gamme est fonction de la tonalité
                                                Grid2.Cell(11, j).ForeColor = DicoCouleurLettre.Item(tbl(0))
                                            End If
                                        Else
                                            If Trim(a) <> "" Then ' cas où le CLip contient des colonnes vides
                                                a = Det_RelativeMajeure2(a) ' si la tonalité est mineure alors on affiche la couleur de la relative majeure
                                                tbl = Split(a)
                                                Grid2.Cell(2, j).BackColor = Color.Red 'DicoCouleur.Item(Trim(tbl(0))) ' la couleur de l'accord est fonction de la tonalité
                                                Grid2.Cell(2, j).ForeColor = Color.Yellow ' DicoCouleurLettre.Item(tbl(0))
                                                Grid2.Cell(11, j).BackColor = Color.Red 'DicoCouleur.Item(Trim(tbl(0))) ' la couleur de la gamme est fonction de la tonalité
                                                Grid2.Cell(11, j).ForeColor = Color.Yellow 'DicoCouleurLettre.Item(tbl(0))
                                            End If
                                        End If
                                    Next
                                    Maj_Répétition()
                                    Maj_Nligne()
                                    Calcul_AutoVoicingZ()
                                    CtrlConsistVar()
                                    ' Mise à jour entête des piano roll
                                    ' *********************************
                                    a = Trim(Det_ListAcc())
                                    b = Trim(Det_ListGam())
                                    c = Trim(Det_ListMarq())
                                    d = Trim(Det_ListTon())
                                    For i = 0 To nb_PianoRoll - 1
                                        If PIANOROLLChargé(i) = True Then
                                            listPIANOROLL(i).PListAcc = a 'Det_ListAcc()
                                            listPIANOROLL(i).PListGam = b 'Det_ListGam()
                                            listPIANOROLL(i).PListMarq = c 'Det_ListMarq()
                                            listPIANOROLL(i).PListTon = d
                                            listPIANOROLL(i).Clear_graphique2() ' <---
                                            listPIANOROLL(i).F1_Refresh()
                                            listPIANOROLL(i).Maj_CalquesMIDI()
                                        End If
                                    Next
                                    ' Mise à jour de Grid2 de drumedit (timeline drumedit)
                                    '
                                    Automation1.PListAcc = Det_ListAcc()
                                    Automation1.PListMarq = Det_ListMarq()
                                    Automation1.F4_Refresh()

                                    Drums.PListAcc = Det_ListAcc()
                                    Drums.PListMarq = Det_ListMarq()
                                    Drums.F2_Refresh()
                                End If
                                Grid2.Range(1, 1, Grid2.Rows - 1, Grid2.Cols - 1).Locked = True
                            End If
                        End If
                    Case Else
                        Dim i1 As Integer = Automation1.PNcanal
                        Dim i2 As Integer = Automation1.PNCourbe
                        Automation1.POrig_Autom_Vérrouillage = True ' vérouillage du tracé de courbe dans AUtomation le temps d'effectuer le "coller"
                        Automation1.Canaux(i1).GridCourbes(i2).AutoRedraw = False
                        Automation1.Canaux(i1).GridCourbes(i2).Selection.PasteData()
                        Automation1.Canaux(i1).GridCourbes(i2).AutoRedraw = True
                        Automation1.Canaux(i1).GridCourbes(i2).Refresh()
                        Automation1.POrig_Autom_Vérrouillage = False ' retour à la possibilité de Tracer dans Automation
                End Select
            Case 1, 2, 3, 5, 6, 7, 8
                ' PIANO ROLL
                iii = OngletCours_Edition - 1
                If OngletCours_Edition > 3 Then iii = OngletCours_Edition - 2
                With listPIANOROLL.Item(iii)
                    If .Orig_PianoR.Orig1 = OrigPianoCourbe.Piano Then ' pianoroll
                        .CollerData2()
                    Else                                               ' courbes
                        'Dim ind As Integer = .Orig_PianoR.N_Courbe
                        'ind = .Orig_PianoR.N_Courbe
                        '.GridCourbes.Item(ind).AutoRedraw = True
                        '
                        '.Orig_PianoR.Vérouillage = False ' pour désactiver le Tracé de la courbe le temps qye le "coller" s'opère
                        '
                        '.GridCourbes.Item(ind).AllowUserPaste = True
                        '.GridCourbes.Item(ind).Range(1, 1, .GridCourbes.Item(ind).Rows - 1, .GridCourbes.Item(ind).Cols - 1).Locked = False
                        '.GridCourbes.Item(ind).Selection.PasteData()
                        '.GridCourbes.Item(ind).Range(1, 1, .GridCourbes.Item(ind).Rows - 1, .GridCourbes.Item(ind).Cols - 1).Locked = True
                        '.Orig_PianoR.Vérouillage = True ' reprise de la possibilité du Tracé
                        '
                        '.GridCourbes.Item(ind).AllowUserPaste = False
                        '.GridCourbes.Item(ind).AutoRedraw = True
                        '.GridCourbes.Item(ind).Refresh()

                        ind = .Orig_PianoR.N_Courbe
                        .Orig_PianoR.Vérouillage = True ' pour désactiver le Tracé de la courbe le temps qye le "coller" s'opère
                        .GridCourbes.Item(ind).Selection.PasteData()
                        .Orig_PianoR.Vérouillage = False ' reprise de la possibilité du Tracé

                    End If
                End With
            Case Else
                ' DRUMS EDIT (Grid1 et Grid2)
                If Drums.PGridOrigine = "Grid1" Then
                    If Drums.Grid1.Selection.FirstCol > 3 And Drums.Grid1.Selection.FirstRow > 0 Then
                        Drums.DéVérouillerGrid1()   ' les cellules ont été vérrouillées dans drums.grid1_Keydown
                        Drums.Sauv_ClipAnnuler()
                        '
                        Drums.Grid1.AllowUserPaste = ClipboardDataEnum.Text
                        Drums.Grid1.Selection.PasteData()
                        'Drums.DrumPaste()
                        ii = Drums.Det_NumPréset2()
                        Drums.Maj_Préset(ii)
                        Drums.DessinDiv()           ' cette méthode contient Autoredraw=false,true et Refresh
                        Drums.VérouillerGrid1()     ' les cellules seront déverrouillées à la fin de Drums.Grid1_KeyUp
                    End If
                End If
                If Drums.PGridOrigine = "Grid2" Then
                    Drums.Grid2.Range(0, 0, Drums.Grid2.Rows - 1, Drums.Grid2.Cols - 1).Locked = False
                    Drums.Grid2.Selection.PasteData()
                    Drums.Grid2.Range(0, 0, Drums.Grid2.Rows - 1, Drums.Grid2.Cols - 1).Locked = True ' dévéroulluer grid2
                End If

        End Select
    End Sub
    Sub CtrlConsistVar()
        Dim DerAcc As Integer = Det_NumDerAccord()
        Dim i, j, k As Integer
        Dim NumVariation As Integer
        Dim ii As Integer

        j = 0
        Do
            j = j + 1
            If Trim(Grid2.Cell(2, j).Text) = "" Then
                NumVariation = TableEventH(j - 1, 1, 1).NumMagnéto ' n° variation de la colonne précédente initiale qui définit la variation
                k = j
                Do ' détermination du nombre de colonne sans accord, donc détermination du nombre de colonne à mettre à jour avec la couleur de variation
                    k = k + 1
                Loop Until Trim(Grid2.Cell(2, k).Text) <> ""
                '
                k = (k - j) - 1 ' nombre de colonne sans accord - dont il faut mettre à jour la couleur de variation - par exemple si une seule colonne alors k=0
                For i = j To j + k '
                    ' effacer la colonne
                    For ii = 4 To 10
                        Grid2.Cell(ii, i).BackColor = Color.White
                    Next
                    ' mettre à jour la couleur de la variation dans la colonne
                    Grid2.Cell(NumVariation + 4, i).BackColor = Grid2.Cell(NumVariation + 4, j - 1).BackColor ' J-1 va cherche la couleur initiale de variation
                    TableEventH(i, 1, 1).NumMagnéto = NumVariation
                Next
                '
                j = j + k
            End If
        Loop Until j = DerAcc Or j > nbMesures

    End Sub
    Sub CouperClip()
        Dim i As Integer
        '
        Grid2.AutoRedraw = False

        For i = Grid2.Selection.FirstCol To Grid2.Selection.LastCol
            TableEventH(i, 1, 1).Ligne = -1
            TableEventH(i, 1, 1).Accord = ""
            TableEventH(i, 1, 1).Gamme = ""
            Grid2.Cell(3, i).Text = 1
            '
            ChoixVariationGrid2(Grid2.Selection.FirstCol, Grid2.Selection.LastCol, 4)
        Next
        '
        With Grid2.Range(1, 0, 1, (nbMesures)) ' ligne "Marqueurs"
            .BackColor = Color.AliceBlue
            .ForeColor = Color.Red
            .Alignment = FlexCell.AlignmentEnum.CenterCenter
        End With
        With Grid2.Range(2, 0, 2, (nbMesures))
            .Alignment = FlexCell.AlignmentEnum.CenterCenter
        End With
        With Grid2.Range(3, 0, 3, (nbMesures))
            .Alignment = FlexCell.AlignmentEnum.CenterCenter
        End With

        Grid2.AutoRedraw = True
        Grid2.Refresh()
    End Sub
    Sub CopierClip()
        Dim i As Integer
        Dim a As New EventH

        ClipEdit.Clear()
        ' 

        For i = Grid2.Selection.FirstCol To Grid2.Selection.LastCol
            a.Position = TableEventH(i, 1, 1).Position
            a.Marqueur = TableEventH(i, 1, 1).Marqueur
            a.Mode = TableEventH(i, 1, 1).Mode
            a.Accord = TableEventH(i, 1, 1).Accord
            a.Gamme = TableEventH(i, 1, 1).Gamme
            a.Degré = TableEventH(i, 1, 1).Degré
            a.Tonalité = TableEventH(i, 1, 1).Tonalité
            a.Ligne = TableEventH(i, 1, 1).Ligne
            a.Répet = TableEventH(i, 1, 1).Répet
            a.NumAcc = TableEventH(i, 1, 1).NumAcc
            a.NumMagnéto = TableEventH(i, 1, 1).NumMagnéto
            a.Racine = TableEventH(i, 1, 1).Racine
            ClipEdit.Add(a)
        Next i
    End Sub
    Function CollerClip() As Boolean
        Dim i As Integer
        Dim j As Integer
        Dim a As New EventH
        CollerClip = True

        j = Grid2.Selection.FirstCol
        Dim LstCol As Integer = (j) + (ClipEdit.Count - 1)
        '
        If LstCol <= nbMesures Then

            SAUV_Annuler(Grid2.Selection.FirstCol, LstCol)

            ' Boucle de controle dépassement de nbMesures par collage d'une répétition
            For i = 0 To ClipEdit.Count - 1
                If j + (ClipEdit(i).Répet - 1) > nbMesures Then
                    CollerClip = False
                End If
                j = j + 1
            Next i
            '
            ' Boucle de collage des infos
            If CollerClip = True Then
                j = Grid2.Selection.FirstCol
                For i = 0 To ClipEdit.Count - 1
                    TableEventH(j, 1, 1).Position = Convert.ToString(j) + ".1" + ".1"  'Trim(ClipEdit(i).Position)
                    TableEventH(j, 1, 1).Marqueur = Trim(ClipEdit(i).Marqueur)
                    TableEventH(j, 1, 1).Mode = Trim(ClipEdit(i).Mode)
                    TableEventH(j, 1, 1).Accord = Trim(ClipEdit(i).Accord)
                    TableEventH(j, 1, 1).Gamme = Trim(ClipEdit(i).Gamme) '
                    TableEventH(j, 1, 1).Degré = ClipEdit(i).Degré
                    TableEventH(j, 1, 1).Tonalité = Trim(ClipEdit(i).Tonalité)
                    TableEventH(j, 1, 1).Répet = ClipEdit(i).Répet
                    TableEventH(j, 1, 1).NumAcc = j 'ClipEdit(i).NumAcc
                    TableEventH(j, 1, 1).NumMagnéto = ClipEdit(i).NumMagnéto
                    TableEventH(j, 1, 1).Racine = ClipEdit(i).Racine
                    If Trim(TableEventH(j, 1, 1).Accord) = "" Then
                        TableEventH(j, 1, 1).Ligne = -1 ' la serie des N° de ligne est mise à jour dans Maj_NLIGNE
                    End If
                    '
                    ChoixVariationGrid2(j, j, ClipEdit(i).NumMagnéto + 4)
                    '
                    j = j + 1
                Next
            End If
        Else
            CollerClip = False
        End If
        If CollerClip = False Then
            If LangueIHM = "fr" Then
                MessageHV.PTitre = "Avertissement"
                MessageHV.PContenuMess = "Erreur interne : Procédure 'CollerClip' : Dépassement de la limite des mesures"
                MessageHV.PTypBouton = "OK"
            Else
                MessageHV.PTitre = "Warning"
                MessageHV.PContenuMess = "Internal Error : Procédure 'CollerClip' : Exceeding the measures limit"
                MessageHV.PTypBouton = "OK"
            End If
            Cacher_FormTransparents()
            MessageHV.ShowDialog()
        End If

    End Function
    Sub SAUV_Annuler(FirstC As Integer, LastC As Integer)
        Dim i As Integer
        Dim a As New EventH
        Dim b As New ClipAcc
        ' Supprimer tous les éléments situés après l'index courant
        '
        IndexAnnuler = IndexAnnuler + 1
        '
        For i = FirstC To LastC
            a.Position = TableEventH(i, 1, 1).Position
            a.Marqueur = TableEventH(i, 1, 1).Marqueur
            a.Mode = TableEventH(i, 1, 1).Mode
            a.Accord = TableEventH(i, 1, 1).Accord
            a.Gamme = TableEventH(i, 1, 1).Gamme
            a.Degré = TableEventH(i, 1, 1).Degré
            a.Tonalité = TableEventH(i, 1, 1).Tonalité
            a.Ligne = TableEventH(i, 1, 1).Ligne
            a.Répet = TableEventH(i, 1, 1).Répet
            a.NumMagnéto = TableEventH(i, 1, 1).NumMagnéto
            a.NumAcc = i '
            a.Vel = TableEventH(i, 1, 1).Vel
            a.Racine = TableEventH(i, 1, 1).Racine
            '
            b.LAcc.Add(a)
        Next
        '
        ClipAnnuler.Add(b)


    End Sub

    Sub MAj_Annuler()
        Dim i As Integer
        Dim m As Integer
        Dim a As New EventH
        Dim b As New ClipAcc
        Dim aa As String
        Dim tbl() As String
        Dim a1, b1, c1, d1 As String


        '
        If IndexAnnuler <> -1 Then
            Grid2.AutoRedraw = False
            Grid2.Range(1, 1, 3, Grid2.Cols - 1).Locked = False

            ' Mise à jour des données
            ' ***********************
            For i = 0 To ClipAnnuler(IndexAnnuler).LAcc.Count - 1
                m = ClipAnnuler(IndexAnnuler).LAcc(i).NumAcc
                TableEventH(m, 1, 1).Position = Trim(ClipAnnuler(IndexAnnuler).LAcc(i).Position)
                TableEventH(m, 1, 1).Marqueur = Trim(ClipAnnuler(IndexAnnuler).LAcc(i).Marqueur)
                TableEventH(m, 1, 1).Mode = Trim(ClipAnnuler(IndexAnnuler).LAcc(i).Mode)
                TableEventH(m, 1, 1).Accord = Trim(ClipAnnuler(IndexAnnuler).LAcc(i).Accord)
                TableEventH(m, 1, 1).Gamme = Trim(ClipAnnuler(IndexAnnuler).LAcc(i).Gamme)
                TableEventH(m, 1, 1).Degré = ClipAnnuler(IndexAnnuler).LAcc(i).Degré
                TableEventH(m, 1, 1).Tonalité = Trim(ClipAnnuler(IndexAnnuler).LAcc(i).Tonalité)
                TableEventH(m, 1, 1).Ligne = ClipAnnuler(IndexAnnuler).LAcc(i).Ligne
                TableEventH(m, 1, 1).Répet = ClipAnnuler(IndexAnnuler).LAcc(i).Répet
                TableEventH(m, 1, 1).NumAcc = ClipAnnuler(IndexAnnuler).LAcc(i).NumAcc
                TableEventH(m, 1, 1).NumMagnéto = ClipAnnuler(IndexAnnuler).LAcc(i).NumMagnéto
                TableEventH(m, 1, 1).Racine = ClipAnnuler(IndexAnnuler).LAcc(i).Racine
                TableEventH(m, 1, 1).Vel = ClipAnnuler(IndexAnnuler).LAcc(i).Vel
                '
                Grid2.Cell(1, m).Text = Trim(TableEventH(m, 1, 1).Marqueur)
                Grid2.Cell(2, m).Text = TradAcc_AnglLat(Trim(TableEventH(m, 1, 1).Accord)) ' traduction seulement si l'origine est en latin
                Grid2.Cell(3, m).Text = Trim(TableEventH(m, 1, 1).Répet)
                Grid2.Cell(11, m).Text = Trim(TableEventH(m, 1, 1).Gamme)
                Grid2.Cell(12, m).Text = Trim(TableEventH(m, 1, 1).Racine)
                '
                ' Mise à jour des couleurs et formats
                ' ***********************************
                If Trim(TableEventH(m, 1, 1).Accord) <> "" Then
                    ' couleur indiquant la tonalité sur la ligne des accords et la ligne des gammes
                    aa = TableEventH(m, 1, 1).Tonalité '
                    aa = Det_RelativeMajeure2(aa) ' si la tonalité est mineure alors on affiche la couleur de la relative majeure
                    tbl = Split(aa)
                    Grid2.Cell(2, m).BackColor = DicoCouleur.Item(Trim(tbl(0))) ' la couleur est fonction de la tonalité
                    Grid2.Cell(2, m).ForeColor = DicoCouleurLettre.Item(tbl(0))
                    Grid2.Cell(11, m).BackColor = DicoCouleur.Item(Trim(tbl(0))) ' la couleur est fonction de la tonalité
                    Grid2.Cell(11, m).ForeColor = DicoCouleurLettre.Item(tbl(0))
                    ' afficher la variation (magnéto = variation)
                    ChoixVariationGrid2(m, m, TableEventH(m, 1, 1).NumMagnéto + 4)
                    Formater_Racine(m) ' formater la racine : couleur et alignement
                Else
                    Grid2.Cell(2, m).BackColor = Color.White
                    Grid2.Cell(2, m).ForeColor = Color.Black
                    Grid2.Cell(11, m).BackColor = Color.White
                    Grid2.Cell(11, m).ForeColor = Color.Black
                    ChoixVariationGrid2(m, m, 4)
                End If
            Next
            '
            If m = 1 Then
                Grid2.Cell(2, m).BackColor = Color.Red
                Grid2.Cell(2, m).ForeColor = Color.Yellow
                Grid2.Cell(11, m).BackColor = Color.Red
                Grid2.Cell(11, m).ForeColor = Color.Yellow
            End If
            '
            Maj_Répétition()
            '
            ClipAnnuler.RemoveAt(IndexAnnuler)
            IndexAnnuler = IndexAnnuler - 1
            '
            Grid2.Range(1, 1, 3, Grid2.Cols - 1).Locked = True
            Grid2.AutoRedraw = True
            Grid2.Refresh()
            '
            ' Mise à jour de l'entête des PianoRoll
            ' *************************************
            a1 = Trim(Det_ListAcc())
            b1 = Trim(Det_ListGam())
            c1 = Trim(Det_ListMarq())
            d1 = Trim(Det_ListTon())
            For i = 0 To nb_PianoRoll - 1
                If PIANOROLLChargé(i) = True Then
                    listPIANOROLL(i).PListAcc = a1 'Det_ListAcc()
                    listPIANOROLL(i).PListGam = b1 'Det_ListGam()
                    listPIANOROLL(i).PListMarq = c1 'Det_ListMarq()
                    listPIANOROLL(i).PListTon = d1
                    listPIANOROLL(i).Clear_graphique2() ' <---
                    listPIANOROLL(i).F1_Refresh()
                    listPIANOROLL(i).Maj_CalquesMIDI()
                End If
            Next
            ' Mise à jour de Grid2 de drumedit  (timeline drumedit)
            Automation1.PListAcc = Det_ListAcc()
            Automation1.PListMarq = Det_ListMarq()
            Automation1.F4_Refresh()
            '
            Drums.PListAcc = Det_ListAcc()
            Drums.PListMarq = Det_ListMarq()
            Drums.F2_Refresh()
        End If
    End Sub
    Sub Formater_Racine(nMesure As Integer)
        ' formater la racine : couleur et alignement
        Grid2.Cell(12, nMesure).Alignment = AlignmentEnum.CenterCenter
        Grid2.Cell(12, nMesure).Font = fnt2
        Grid2.Cell(12, nMesure).ForeColor = Color.DarkOrange
    End Sub
    Sub RAZ_TCopie()
        Dim i As Integer
        If TCopie.Count > 0 Then
            TCopie.Clear()
            For i = 0 To TCopie.Count - 1
                TCopie.Item(i).Actif = False
            Next i
        End If
    End Sub
    Sub Copie()
        Dim i As Integer
        Dim m, t, ct As Integer
        Dim e As Integer
        Dim mDeb As Integer

        e = 0
        Select Case DerGridCliquée
            Case GridCours.Grid2
                RAZ_TCopie()
                '
                i = -1
                mDeb = Grid2.Selection.FirstCol
                For m = Grid2.Selection.FirstCol To Grid2.Selection.LastCol
                    For t = 0 To UBound(TableEventH, 2)
                        For ct = 0 To UBound(TableEventH, 3)
                            If TableEventH(m, t, ct).Ligne <> -1 Then
                                OrigineEdition = GridCours.Grid2
                                i += 1

                                If i > TCopie.Count - 1 Then
                                    TCopie.Add(New TamponCopie)
                                End If
                                '
                                TCopie(i).Actif = True
                                TCopie(i).Tonalité = TableEventH(m, t, ct).Tonalité
                                TCopie(i).Accord = TableEventH(m, t, ct).Accord
                                TCopie(i).Gamme = TableEventH(m, t, ct).Gamme
                                TCopie(i).Mode = TableEventH(m, t, ct).Mode
                                TCopie(i).Degré = TableEventH(m, t, ct).Degré
                                TCopie(i).Marqueur = TableEventH(m, t, ct).Marqueur
                                TCopie(i).Détails = TableEventH(m, t, ct).Détails
                                '
                                TCopie(i).m = m
                                TCopie(i).t = t
                                TCopie(i).ct = ct
                                '
                                TCopie(i).Décalage = m - mDeb
                                '
                                TCopie.Item(0).Nb_Items = i ' le nombre d'items de la copie est mis à jors uniquement dans le 1er item
                            End If
                        Next ct
                    Next t
                Next m
            Case GridCours.Grid3

        End Select
    End Sub
    Sub Effacer_Grid2(mDeb As Integer, mFin As Integer)
        Dim i As Integer
        Dim k As Integer
        Dim t As Integer
        Dim ct As Integer
        '
        t = 1
        ct = 1
        '

        ' Effacer Grid2
        ' *************
        k = (mFin - mDeb)

        For i = mDeb To mFin
            Grid2.Column(i).Locked = False
        Next
        '
        Grid2.AutoRedraw = False
        For i = mDeb To mFin
            Grid2.Column(i).Locked = True
            If i <> 1 Then ' on laisse la couleur rouge pour la 1ere mesure
                Grid2.Cell(1, i).BackColor = Color.White
                Grid2.Cell(1, i).ForeColor = Color.Black
            End If
            Grid2.Cell(1, i).Text = ""
            ''
            Grid2.Cell(0, i).Locked = False
            Grid2.Cell(0, i).BackColor = Color.Beige
            Grid2.Cell(0, i).Locked = True
            'Grid2.Cell(0, i).SetFocus()
        Next
        Grid2.Refresh()
        Grid2.AutoRedraw = True
        '
    End Sub
    Sub Effacer_Grid2_3(mDeb As Integer, mFin As Integer)
        Dim i As Integer
        Dim k As Integer
        '
        ' Effacer Grid2
        ' *************

        If mDeb = 1 Then mDeb = 2
        Grid2.AutoRedraw = False

        k = (mFin - mDeb)

        For i = mDeb To mFin
            Grid2.Column(i).Locked = False
        Next
        '
        For i = mDeb To mFin
            If Trim(Grid2.Cell(1, i).Text) <> "" Then
                Grid2.Column(i).Locked = True
                If i <> 1 Then ' on laisse la couleur rouge pour la 1ere mesure
                    Grid2.Cell(1, i).BackColor = Color.White
                    Grid2.Cell(1, i).ForeColor = Color.Black
                End If
                '
                Grid2.Cell(1, i).Text = ""
                '
                Grid2.Cell(0, i).Locked = False
                Grid2.Cell(0, i).BackColor = Color.Beige
                Grid2.Cell(0, i).ForeColor = Color.Black
                Grid2.Cell(0, i).Locked = True
                '
                TableEventH(i, 1, 1).Accord = ""
                TableEventH(i, 1, 1).Gamme = ""
                TableEventH(i, 1, 1).Tonalité = ""
                TableEventH(i, 1, 1).Position = ""
                TableEventH(i, 1, 1).Ligne = -1
                TableEventH(i, 1, 1).Marqueur = ""
                TableEventH(i, 1, 1).Détails = ""
            End If
        Next i
        '
        Grid2.Refresh()
        Grid2.AutoRedraw = True
    End Sub



    Sub ZAnnulation_Sauvegarde(mdeb As Integer, mfin As Integer)
        Dim i As Integer
        Dim m, t, ct As Integer
        Dim Détection_Mesure_Vide As Boolean = True
        RAZ_ZAnnulation()
        i = -1
        Détection_Mesure_Vide = True
        ZAnnulation_FirstCol = mdeb
        For m = mdeb To mfin
            For t = 0 To UBound(TableEventH, 2)
                For ct = 0 To UBound(TableEventH, 3)
                    If TableEventH(m, t, ct).Ligne <> -1 Then
                        '
                        ' activer le ctrl+Z
                        'AnnulerToolStripMenuItem.Enabled = False

                        Détection_Mesure_Vide = False
                        ZAnnulation_Valide = True
                        ' Gestion du nombre de Tampons
                        ' ****************************
                        i = i + 1
                        If i > TZAnnulation.Count - 1 Then
                            TZAnnulation.Add(New TamponCopie)
                        End If
                        ' Maj Tampon annulation
                        ' *********************
                        TZAnnulation(i).Actif = True
                        TZAnnulation(i).Mesure_Effacer = -1
                        TZAnnulation(i).Tonalité = TableEventH(m, t, ct).Tonalité
                        TZAnnulation(i).Accord = TableEventH(m, t, ct).Accord
                        TZAnnulation(i).Gamme = TableEventH(m, t, ct).Gamme
                        TZAnnulation(i).Mode = TableEventH(m, t, ct).Mode
                        TZAnnulation(i).Degré = TableEventH(m, t, ct).Degré
                        TZAnnulation(i).Marqueur = TableEventH(m, t, ct).Marqueur
                        TZAnnulation(i).Détails = TableEventH(m, t, ct).Détails
                        '
                        TZAnnulation(i).m = m
                        TZAnnulation(i).t = t
                        TZAnnulation(i).ct = ct
                        '
                        TZAnnulation(i).Décalage = m - mdeb
                        '
                        TZAnnulation.Item(0).Nb_Items = i ' le nombre d'items de la copie est mis à jors uniquement dans le 1er item
                    End If
                Next ct
            Next t
            If Détection_Mesure_Vide = True Then
                i = i + 1
                If i > TZAnnulation.Count - 1 Then
                    TZAnnulation.Add(New TamponCopie)
                End If
                TZAnnulation(i).Actif = True
                TZAnnulation(i).Mesure_Effacer = m
                TZAnnulation(i).m = m
                TZAnnulation.Item(0).Nb_Items = i
                ZAnnulation_Valide = True
            Else
                Détection_Mesure_Vide = True
            End If
        Next m
        '
    End Sub
    Sub RAZ_ZAnnulation()
        Dim i As Integer

        If TZAnnulation.Count > 0 Then
            TZAnnulation.Clear()
            For i = 0 To TZAnnulation.Count - 1
                TZAnnulation.Item(i).Actif = False
            Next i
        End If
    End Sub
    Sub ZAnnulation_Restitution()
        Dim i, j, k As Integer
        Dim m, t, ct As Integer
        Dim Der_Mesure_aEffacer As Integer
        Dim a As String
        Dim tbl() As String
        Dim Nb_Mesures As Integer

        If ZAnnulation_Valide = True Then
            ZAnnulation_Valide = False
            ' Effacer la zone à copier
            ' ************************
            k = TZAnnulation.Item(0).Nb_Items
            Nb_Mesures = TZAnnulation.Item(k).m - TZAnnulation.Item(0).m
            Der_Mesure_aEffacer = ZAnnulation_FirstCol + Nb_Mesures ' détermination de la dernière mesure à effacer
            Effacer_Grid2_3(ZAnnulation_FirstCol, Der_Mesure_aEffacer)
            '
            i = 0
            '
            'Grid2.LeftCol = ZAnnulation_FirstCol + 3
            Grid2.AutoRedraw = False

            Do
                If TZAnnulation.Item(i).Actif = True Then
                    If TZAnnulation(i).Mesure_Effacer = -1 Then

                        m = ZAnnulation_FirstCol + TZAnnulation.Item(i).Décalage
                        t = TZAnnulation.Item(i).t
                        ct = TZAnnulation.Item(i).ct
                        '
                        ' Mise à jour de TableEventH
                        ' **************************
                        TableEventH(m, t, ct).Accord = TZAnnulation.Item(i).Accord
                        TableEventH(m, t, ct).Gamme = TZAnnulation.Item(i).Gamme
                        TableEventH(m, t, ct).Mode = TZAnnulation.Item(i).Mode ' Mode = Gamme - le mode n'est pas affiché mais je garde l'info por le moment
                        TableEventH(m, t, ct).Tonalité = TZAnnulation.Item(i).Tonalité '
                        TableEventH(m, t, ct).Marqueur = TZAnnulation.Item(i).Marqueur

                        TableEventH(m, t, ct).Degré = TZAnnulation.Item(i).Degré
                        '
                        Entrée_Position = Trim(Str(m)) + "." + Trim(Str(t)) + "." + Trim(Str(ct))
                        TableEventH(m, t, ct).Position = Entrée_Position
                        TableEventH(m, t, ct).Détails = TZAnnulation.Item(i).Détails


                        Grid2.AutoRedraw = False
                        '
                        ' Ecriture de l'accord dans la grille Grid2
                        ' *****************************************
                        If Trim(Grid2.Cell(1, m).Text) = "" Then
                            Grid2.Cell(1, m).Locked = False
                            Grid2.Cell(1, m).Alignment = FlexCell.AlignmentEnum.CenterCenter
                            Grid2.Cell(1, m).Text = TZAnnulation.Item(i).Accord
                            If Langue = "fr" Then
                                Grid2.Cell(1, m).Text = TradAcc_AnglLat(TZAnnulation.Item(i).Accord)
                            End If
                            Grid2.Cell(1, m).Locked = True
                            Grid2.ReadonlyFocusRect = FlexCell.FocusRectEnum.Solid
                        Else
                            Grid2.Cell(1, m).Text = ChaineAccord(m)
                        End If
                        '
                        ' Mise à jour correspondante dans Grid3
                        ' *************************************
                        tbl = Split(Trim(TZAnnulation.Item(i).Tonalité))
                        a = tbl(0) ' note de la tonalité courante
                        j = Det_IndexGrid3_De_ColGrid2(m)
                        j = (j) + ((t - 1) * Det_NbDivTemps2(t)) + (ct - 1) ' position dans grid3
                        If j <> 1 Then
                            '
                            If m <> 1 Then
                                Grid2.Cell(1, m).BackColor = DicoCouleur.Item(a) ' la couleur est fonction de la tonalité
                                Grid2.Cell(1, m).ForeColor = DicoCouleurLettre.Item(a)
                            Else
                                Grid2.Cell(1, m).BackColor = Color.Red ' la couleur est fonction de la tonalité
                                Grid2.Cell(1, m).ForeColor = Color.White
                            End If
                            'Grid2.Refresh()
                            '
                        End If
                        '
                        Grid2.AutoRedraw = True
                        Grid2.Refresh()
                    Else
                        Effacer_Grid2_3(TZAnnulation(i).Mesure_Effacer, TZAnnulation(i).Mesure_Effacer)
                    End If
                    'Ecriture_Entrée_Ds_CompoGrid()
                    i = i + 1
                End If
            Loop Until i > k
        End If
    End Sub

    Private Sub NotesCommunes2_CheckedChanged(sender As Object, e As EventArgs)
        'NotesCommunes.Checked = NotesCommunes2.Checked
    End Sub

    Private Sub AnnulerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AnnulerToolStripMenuItem.Click
        Dim iii As Integer
        Select Case OngletCours_Edition
            Case 0 ' HyperArp

                Select Case OngletCours_HyperARP
                    Case 0, 1, 2, 3, 127 ' Onglets Gammes d'Accords, Progression d'Accords,Zone de Voicing, Perso,Time Line des accords
                        MAj_Annuler()
                        Calcul_AutoVoicingZ()
                    Case Else ' case annulation Ciurbes : on ne fait rien. Pas d'annulation de courbes.
                End Select

            Case 1, 2, 3, 5, 6, 7, 8
                iii = OngletCours_Edition - 1
                If OngletCours_Edition > 3 Then iii = OngletCours_Edition - 2
                With listPIANOROLL.Item(iii)
                    If .Orig_PianoR.Orig1 = OrigPianoCourbe.Piano Then ' il n'y a pas d'annulation sur le trçage des courbes
                        If listPIANOROLL(iii).flagAssist2_Z Then
                            MAj_Annuler()
                            Calcul_AutoVoicingZ()
                            listPIANOROLL.Item(iii).flagAssist2_Z = False
                        Else
                            listPIANOROLL.Item(iii).AnnulerData2()
                        End If
                    End If
                End With

            Case Else ' Drum edit
                If Drums.PGridOrigine = "Grid1" Then
                    Drums.Restit_ClipAnnuler()
                End If
                If Drums.PGridOrigine = "Grid2" Then

                End If

        End Select

    End Sub
    Sub Annuler()
        If Not (TZAnnulationGrid1.Grille = TamponInfoGrid1.TGrilleCours.Grid1 Or
        TZAnnulationGrid1.Grille = TamponInfoGrid1.TGrilleCours.Grid4) Then
            ZAnnulation_Restitution()
        End If
    End Sub


    Private Sub MIDIResetToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MIDIResetToolStripMenuItem.Click
        MIDIReset()
    End Sub
    Sub MIDIReset()
        If Not (SortieMidi.Item(ChoixSortieMidi).IsOpen) Then
            SortieMidi.Item(ChoixSortieMidi).Open()
        End If
        ToutesNotesOff()
        SortieMidi.Item(ChoixSortieMidi).SilenceAllNotes()
        Init_CTRLMIDI2()
    End Sub
    Sub ExportMIDI()
        Dim a As String
        'MessageHV.PTypBouton = "OuiNon"
        'If LangueIHM = "fr" Then
        '
        'MessageHV.PTitre = "Paramètres d'export MIDI"
        'MessageHV.PContenuMess = "Voulez-vous exporter les contrôleurs (PRG, VOL, PAN ..) dans le fichier MIDI  ?"
        'Else
        'MessageHV.PTitre = "MIDI export parameters"
        'MessageHV.PContenuMess = "Do you want to export controlers (PRG, VOL, PAN ..) into MIDI file  ?"
        'End If
        '
        'MessageHV.ShowDialog()
        '
        'ExportCTRL = False
        'If MessageHV.PSortie = "Oui" Then ExportCTRL = True

        '
        Try
            If Det_NomFichMIDI() Then
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                a = Création_CTemp() + "\" + "djgbv58147.mid" '  a = Création_CTemp() + "\" + "HyperArp.mid"
                CalculArp(True)
                If My.Computer.FileSystem.FileExists(FichierMIDI) Then ' FichierMIDI est une variable globale mise à jour dans Det_NomFichMIDI
                    My.Computer.FileSystem.DeleteFile(FichierMIDI)
                End If
                '
                My.Computer.FileSystem.CopyFile(a, FichierMIDI)
            End If
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        Catch ex As Exception

            If LangueIHM = "fr" Then
                MessageHV.PTitre = "Avertissement"
                MessageHV.PContenuMess = "Erreur interne : procédure ExportMIDI : " + ex.Message
                MessageHV.PTypBouton = "OK"
            Else
                MessageHV.PTitre = "Warning"
                MessageHV.PContenuMess = "Internal Error : procédure ExportMIDI : " + ex.Message
                MessageHV.PTypBouton = "OK"
            End If
            Cacher_FormTransparents()
            MessageHV.ShowDialog()
            End
        End Try
    End Sub

    Function Det_Présence_Marqueurs2() As Boolean
        Dim m, t, ct As Integer

        Det_Présence_Marqueurs2 = False
        '
        For m = 0 To UBound(TableEventH, 1) '- 1
            For t = 0 To UBound(TableEventH, 2) '- 1
                For ct = 0 To UBound(TableEventH, 3) '- 1
                    If Trim(TableEventH(m, t, ct).Marqueur) <> "" And Trim(TableEventH(m, t, ct).Position) <> "" Then
                        Det_Présence_Marqueurs2 = True
                        Exit For
                    End If
                Next ct
            Next t
        Next m
        '
    End Function
    Private Sub MenuExportsMIDI_MouseHover(sender As Object, e As EventArgs) Handles MenuExportsMIDI.MouseHover

    End Sub
    Sub CAD_Init_Aff()
        Dim i As Integer

        If EnChargement = False Then
            For i = 0 To 4
                TabCadDegrés.Item(i).Visible = True
                TabCad.Item(i).Visible = True
                TabCadDegrés.Item(i).Text = ""
                TabCad.Item(i).Text = ""
                TabCad.Item(i).Visible = False
                TabCadDegrés.Item(i).Visible = False
                '
                '
            Next i
        End If
    End Sub
    Sub BLUES_Init_Aff()
        Dim i As Integer

        If EnChargement = False Then
            For i = 0 To 6
                TabBluesDegrés.Item(i).Text = ""
                TabBlues.Item(i).Text = ""
            Next i
        End If
    End Sub

    Sub Cad_ComplèteMaj()
        If EnChargement = False Then
            '
            CAD_Init_Aff()
            '
            TabCadDegrés.Item(0).Text = "I"
            '
            TabCadDegrés.Item(0).Visible = True ' degré
            TabCad.Item(0).Visible = True ' nom de l'accord

            '
            TabCadDegrés.Item(1).Text = "IV"
            '
            TabCadDegrés.Item(1).Visible = True ' degré
            TabCad.Item(1).Visible = True ' nom de l'accord


            TabCadDegrés.Item(2).Text = "V"
            '
            TabCadDegrés.Item(2).Visible = True ' degré
            TabCad.Item(2).Visible = True ' nom de l'accord


            TabCadDegrés.Item(3).Text = "I"
            '
            TabCadDegrés.Item(3).Visible = True ' degré
            TabCad.Item(3).Visible = True ' nom de l'accord

            '
            ''Label28.Size = New Size(420, 33)
            '

            CAD_Maj_TableGlobalAcc()
        End If
    End Sub
    Sub Cad_251Maj()
        If EnChargement = False Then
            '
            CAD_Init_Aff()
            '
            TabCadDegrés.Item(0).Text = "I"
            TabCadDegrés.Item(0).Visible = True ' degré
            TabCad.Item(0).Visible = True ' nom de l'accord


            TabCadDegrés.Item(1).Text = "II"
            TabCadDegrés.Item(1).Visible = True ' degré
            TabCad.Item(1).Visible = True ' nom de l'accord


            TabCadDegrés.Item(2).Text = "V"
            TabCadDegrés.Item(2).Visible = True ' degré
            TabCad.Item(2).Visible = True ' nom de l'accord


            TabCadDegrés.Item(3).Text = "I"
            TabCadDegrés.Item(3).Visible = True ' degré
            TabCad.Item(3).Visible = True ' nom de l'accord

            '
            'Label28.Size = New Size(426, 33)
            '
            CAD_Maj_TableGlobalAcc()
        End If
    End Sub
    Sub Cad_DemiMaj()
        If EnChargement = False Then
            '
            CAD_Init_Aff()
            '
            TabCadDegrés.Item(0).Text = "V"
            TabCadDegrés.Item(0).Visible = True ' degré
            TabCad.Item(0).Visible = True ' nom de l'accord


            TabCadDegrés.Item(1).Text = "II"
            TabCadDegrés.Item(1).Visible = True ' degré
            TabCad.Item(1).Visible = True ' nom de l'accord


            TabCadDegrés.Item(2).Text = "V"
            TabCadDegrés.Item(2).Visible = True ' degré
            TabCad.Item(2).Visible = True ' nom de l'accord

            '
            'Label28.Size = New Size(316, 33)
            '
            CAD_Maj_TableGlobalAcc()
        End If
    End Sub
    Sub Cad_ParfaiteMaj()
        If EnChargement = False Then
            '
            CAD_Init_Aff()
            '
            TabCadDegrés.Item(0).Text = "I"
            TabCadDegrés.Item(0).Visible = True ' degré
            TabCad.Item(0).Visible = True ' nom de l'accord


            TabCadDegrés.Item(1).Text = "V"
            TabCadDegrés.Item(1).Visible = True ' degré
            TabCad.Item(1).Visible = True ' nom de l'accord


            TabCadDegrés.Item(2).Text = "I"
            TabCadDegrés.Item(2).Visible = True ' degré
            TabCad.Item(2).Visible = True ' nom de l'accord

            '
            'Label28.Size = New Size(316, 33)
            '
            CAD_Maj_TableGlobalAcc()
        End If
    End Sub
    Sub Cad_PlagaleMaj()
        If EnChargement = False Then
            '
            CAD_Init_Aff()
            '
            TabCadDegrés.Item(0).Text = "I"
            TabCadDegrés.Item(0).Visible = True ' degré
            TabCad.Item(0).Visible = True ' nom de l'accord


            TabCadDegrés.Item(1).Text = "IV"
            TabCadDegrés.Item(1).Visible = True ' degré
            TabCad.Item(1).Visible = True ' nom de l'accord


            TabCadDegrés.Item(2).Text = "I"
            TabCadDegrés.Item(2).Visible = True ' degré
            TabCad.Item(2).Visible = True ' nom de l'accord

            '
            'Label28.Size = New Size(316, 33)
            '
            CAD_Maj_TableGlobalAcc()
        End If
    End Sub
    Sub Cad_Plagale2Maj()
        If EnChargement = False Then
            '
            CAD_Init_Aff()
            '
            TabCadDegrés.Item(0).Text = "I"
            TabCadDegrés.Item(0).Visible = True ' degré
            TabCad.Item(0).Visible = True ' nom de l'accord


            TabCadDegrés.Item(1).Text = "V"
            TabCadDegrés.Item(1).Visible = True ' degré
            TabCad.Item(1).Visible = True ' nom de l'accord


            TabCadDegrés.Item(2).Text = "VI"
            TabCadDegrés.Item(2).Visible = True ' degré
            TabCad.Item(2).Visible = True ' nom de l'accord


            TabCadDegrés.Item(3).Text = "IV"
            TabCadDegrés.Item(3).Visible = True ' degré
            TabCad.Item(3).Visible = True ' nom de l'accord


            TabCadDegrés.Item(4).Text = "I"
            TabCadDegrés.Item(4).Visible = True ' degré
            TabCad.Item(4).Visible = True ' nom de l'accord

            '
            'Label28.Size = New Size(535, 33)
            '
            CAD_Maj_TableGlobalAcc()
        End If
    End Sub
    Sub Cad_ModaleMaj()
        If EnChargement = False Then
            '
            CAD_Init_Aff()
            '
            TabCadDegrés.Item(0).Text = "IV"
            TabCadDegrés.Item(0).Visible = True ' degré
            TabCad.Item(0).Visible = True ' nom de l'accord


            TabCadDegrés.Item(1).Text = "II"
            TabCadDegrés.Item(1).Visible = True ' degré
            TabCad.Item(1).Visible = True ' nom de l'accord


            TabCadDegrés.Item(2).Text = "IV"
            TabCadDegrés.Item(2).Visible = True ' degré
            TabCad.Item(2).Visible = True ' nom de l'accord

            '
            'Label28.Size = New Size(316, 33)
            '
            CAD_Maj_TableGlobalAcc()
        End If
    End Sub
    Sub Cad_RompueMaj()
        If EnChargement = False Then
            '
            CAD_Init_Aff()
            '
            TabCadDegrés.Item(0).Text = "I"
            TabCadDegrés.Item(0).Visible = True ' degré
            TabCad.Item(0).Visible = True ' nom de l'accord


            TabCadDegrés.Item(1).Text = "V"
            TabCadDegrés.Item(1).Visible = True ' degré
            TabCad.Item(1).Visible = True ' nom de l'accord


            TabCadDegrés.Item(2).Text = "VI"
            TabCadDegrés.Item(2).Visible = True ' degré
            TabCad.Item(2).Visible = True ' nom de l'accord

            '
            'Label28.Size = New Size(316, 33)
            '
            CAD_Maj_TableGlobalAcc()
        End If
    End Sub
    Sub Cad_Rompue2Maj()
        If EnChargement = False Then
            '
            CAD_Init_Aff()
            '
            TabCadDegrés.Item(0).Text = "VI"
            TabCadDegrés.Item(0).Visible = True ' degré
            TabCad.Item(0).Visible = True ' nom de l'accord


            TabCadDegrés.Item(1).Text = "V"
            TabCadDegrés.Item(1).Visible = True ' degré
            TabCad.Item(1).Visible = True ' nom de l'accord


            TabCadDegrés.Item(2).Text = "IV"
            TabCadDegrés.Item(2).Visible = True ' degré
            TabCad.Item(2).Visible = True ' nom de l'accord


            TabCadDegrés.Item(3).Text = "V"
            TabCadDegrés.Item(3).Visible = True ' degré
            TabCad.Item(3).Visible = True ' nom de l'accord


            TabCadDegrés.Item(4).Text = "VI"
            TabCadDegrés.Item(4).Visible = True ' degré
            TabCad.Item(4).Visible = True ' nom de l'accord

            '
            'Label28.Size = New Size(535, 33)
            '
            CAD_Maj_TableGlobalAcc()
        End If
    End Sub
    Sub Cad_Rompue3Maj()
        If EnChargement = False Then
            '
            CAD_Init_Aff()
            '
            TabCadDegrés.Item(0).Text = "VI"
            TabCadDegrés.Item(0).Visible = True ' degré
            TabCad.Item(0).Visible = True ' nom de l'accord


            TabCadDegrés.Item(1).Text = "IV"
            TabCadDegrés.Item(1).Visible = True ' degré
            TabCad.Item(1).Visible = True ' nom de l'accord


            TabCadDegrés.Item(2).Text = "V"
            TabCadDegrés.Item(2).Visible = True ' degré
            TabCad.Item(2).Visible = True ' nom de l'accord


            TabCadDegrés.Item(3).Text = "VI"
            TabCadDegrés.Item(3).Visible = True ' degré
            TabCad.Item(3).Visible = True ' nom de l'accord

            '
            'Label28.Size = New Size(426, 33)
            '
            CAD_Maj_TableGlobalAcc()
        End If
    End Sub
    Sub Cad_Modale2Maj()
        If EnChargement = False Then
            '
            CAD_Init_Aff()
            '
            TabCadDegrés.Item(0).Text = "VI"
            TabCadDegrés.Item(0).Visible = True ' degré
            TabCad.Item(0).Visible = True ' nom de l'accord


            TabCadDegrés.Item(1).Text = "III"
            TabCadDegrés.Item(1).Visible = True ' degré
            TabCad.Item(1).Visible = True ' nom de l'accord


            TabCadDegrés.Item(2).Text = "VI"
            TabCadDegrés.Item(2).Visible = True ' degré
            TabCad.Item(2).Visible = True ' nom de l'accord

            '
            'Label28.Size = New Size(316, 33)
            '
            CAD_Maj_TableGlobalAcc()
        End If
    End Sub
    Sub Cad_NapolitaineMaj()
        If EnChargement = False Then
            '
            CAD_Init_Aff()
            '
            TabCadDegrés.Item(0).Text = "III"
            TabCadDegrés.Item(0).Visible = True ' degré
            TabCad.Item(0).Visible = True ' nom de l'accord


            TabCadDegrés.Item(1).Text = "IV"
            TabCadDegrés.Item(1).Visible = True ' degré
            TabCad.Item(1).Visible = True ' nom de l'accord


            TabCadDegrés.Item(2).Text = "III"
            TabCadDegrés.Item(2).Visible = True ' degré
            TabCad.Item(2).Visible = True ' nom de l'accord

            '
            'Label28.Size = New Size(316, 33)
            '
            CAD_Maj_TableGlobalAcc()
        End If
    End Sub
    Sub Cad_Modale3Maj()
        If EnChargement = False Then
            '
            CAD_Init_Aff()
            '
            TabCadDegrés.Item(0).Text = "VI"
            TabCadDegrés.Item(0).Visible = True ' degré
            TabCad.Item(0).Visible = True ' nom de l'accord


            TabCadDegrés.Item(1).Text = "VII"
            TabCadDegrés.Item(1).Visible = True ' degré
            TabCad.Item(1).Visible = True ' nom de l'accord


            TabCadDegrés.Item(2).Text = "VI"
            TabCadDegrés.Item(2).Visible = True ' degré
            TabCad.Item(2).Visible = True ' nom de l'accord

            '
            'Label28.Size = New Size(316, 33)
            '
            CAD_Maj_TableGlobalAcc()
        End If
    End Sub
    Sub CAD_Maj_TableGlobalAcc()
        Dim i, j As Integer

        Dim typAccord As Integer

        '
        For i = 0 To 6
            For j = 0 To 1
                CAD_TableGlobalAcc(j, 0, i) = TableGlobalAcc(j, 0, i)
                CAD_TableGlobalAcc(j, 1, i) = TableGlobalAcc(j, 1, i)
            Next j
        Next
        '
        typAccord = ComboBox6.SelectedIndex
        '
        Select Case typAccord
            Case 0
                CAD_Menu3_notes()
            Case 1
                CAD_Menu4_notes()
        End Select
        '
    End Sub



    Public Sub CAD_Maj_TableCoursAccInit()
        Dim i As Integer

        For i = 0 To 6
            CAD_TableCoursAcc(i).TyAcc = "3 Notes"
            CAD_TableCoursAcc(i).Accord = TableGlobalAcc(0, 0, i)
            CAD_TableCoursAcc(i).Octave = "0"
            '
        Next
        '
        'CAD_Maj_RenversementInit()
    End Sub
    Private Sub TabCad_MouseUp(ByVal sender As Object, ByVal e As MouseEventArgs)
        Dim ind As Integer
        Dim com As Label = sender

        '
        ind = Val(com.Tag)
        '
        If AccordAEtéJoué = True Then
            CouperJouerAccord2()
            AccordAEtéJoué = False
        End If
        '
        RAZ_AffNoteAcc()
    End Sub
    Private Sub TabCad_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs)
        Dim ind As Integer
        Dim com As Label = sender
        Dim a, b, c As String
        Dim IndexDegré As Integer
        Dim tbl() As String
        Dim tbl1() As String
        Dim Tbl2() As String

        '
        ind = Val(com.Tag)
        a = TabCadDegrés.Item(ind).Text
        IndexDegré = Calc_CadDegrés(Trim(a)) 'Det_IndexDegré(a) ' valeur du dégré dans la tonalité majeure, par exemple I--> 0, II --> 1, III--> 2, IV-->3, V--> 4
        '
        CAD_LabelCours = IndexDegré
        '
        If e.Button() = Windows.Forms.MouseButtons.Left Then
            Cad_MarquerAccord(ind)
        End If
        '
        If e.Button() = MouseButtons.Right Then '  And TabTonsSelect.Item(ind).Checked = True ' My.Computer.Keyboard.ShiftKeyDown And 

            'MarquerAccord(ind)
            If Trim(TabCad.Item(ind).Text) <> "___" And Trim(TabCad.Item(ind).Text) <> "" Then
                Cad_MarquerAccord(ind)
                Maj_TousAccordsMnContextCAD(ind)
                CAD_Accord.Text = Trim(TabCad.Item(ind).Text) ' valeur de l'accord dans menu contextuel
                ContextMenuStrip4.Show(CType(sender, Object), e.Location)
            End If
        End If
        '
        a = Det_TonaCours2()
        Entrée_Tonalité = Trim(a + " " + "Maj")
        If Cad_OrigineAccord = Modes.Cadence_Majeure Then
            Entrée_Tonalité = Trim(a + " " + "Maj")
            Entrée_Gamme = Entrée_Tonalité
        End If
        '
        If Cad_OrigineAccord = Modes.Cadence_Mineure Then
            c = Trim(ComboBox2.Text)
            If Langue = "fr" Then
                c = TradAcc_LatAngl(Trim(ComboBox2.Text))
            End If
            Tbl2 = Split(Trim(c), " ")
            b = Tbl2(0)

            Entrée_Tonalité = Trim(a + " " + "Maj")
            Entrée_Gamme = Trim(b + " " + "MinH")
            CAD_LabelCours = Det_IndexDegréMin(Trim(TabCadDegrés.Item(ind).Text))
        End If
        '
        If Cad_OrigineAccord = Modes.Cadence_Mixte Then
            Select Case ComboBox4.Text
                Case "Hispanique", "Hispanic"

                    Entrée_Tonalité = Trim(a + " " + "Maj") 'Hispanique_Tonalité(ind)
                    Entrée_Gamme = Hispanique_Gamme(ind)
                    tbl1 = Split(Trim(Entrée_Gamme))
                    CAD_LabelCours = Calc_CadDegrés(Trim(TabCadDegrés.Item(ind).Text))
            End Select
        End If
        '
        ' Jouer Accord
        ' ************
        If e.Button() = MouseButtons.Left And My.Computer.Keyboard.CtrlKeyDown Then ' 
            '
            a = Trim(ComboBox1.Text)
            tbl = Split(a)
            Clef = Trim(Det_Clef(tbl(0)))
            '
            If ComboMidiOut.Items.Count > 0 Then
                'JouerSourceTabCad(ind)
                JouerAccord(Trim(TabCad.Item(ind).Text))
            End If
        End If
        '
        If e.Button() = MouseButtons.Middle And Not (My.Computer.Keyboard.CtrlKeyDown) And Not (My.Computer.Keyboard.AltKeyDown) _
                                        And Not (My.Computer.Keyboard.ShiftKeyDown) Then
            AfficherAccordSource(Trim(TabCad.Item(ind).Text))
        End If
        '
        ' Pour glisser - déposer
        ' **********************
        If e.Button() = MouseButtons.Left And Not (My.Computer.Keyboard.CtrlKeyDown) And Not (My.Computer.Keyboard.AltKeyDown) _
                                        And Not (My.Computer.Keyboard.ShiftKeyDown) Then
            TabCad.Item(ind).DoDragDrop(TabCad.Item(ind).Text, DragDropEffects.Copy Or DragDropEffects.Move)
            If Trim(Valeur_Drag) <> "" Then
                Maj_DragDrop()
            End If
        End If
    End Sub
    Function Det_DegréMinMaj(degréMaj As Integer) As Integer ' 
        ' Procédure appelée uniquement par TabCad_MouseDown 
        ' Cas d'utilisation : Mode Mixte Hispanique 
        ' Permet de transformer le degré d'un accord exprimé en mode majeur dans le degré du mode dans lequel il doit être considéré
        ' F et G doivent être considérés comme en mode majeur, il ont donc les degrés 3 et 4
        ' E et Am doivent être considérés commen en mode mineur, leur degré est respictement 4 et 0
        Det_DegréMinMaj = 0
        Select Case degréMaj
            Case 0 ' 
                Det_DegréMinMaj = 2
            Case 1 ' 
                Det_DegréMinMaj = 3 '
            Case 2
                Det_DegréMinMaj = 4  'degré de E(m) est 2 (III) en Mode Majeur --> devient 4 (V) dans la cadence Hispanique
            Case 3
                Det_DegréMinMaj = 3  'degré de F est 3 (IV) en Mode Majeur --> reste 3 (IV)dans la cadence Hispanique
            Case 4
                Det_DegréMinMaj = 4 ' degré de G est 4 (V) en Mode Majeur --> reste 4 (V)dans la cadence Hispanique
            Case 5
                Det_DegréMinMaj = 0 ' degré de Am est 5 (VI) en Mode Majeur --> devient 0 (I)dans la cadence Hispanique
            Case 6
                Det_DegréMinMaj = 1
        End Select
    End Function


    Function Hispanique_Gamme(ind As Integer) As String
        Dim a, c As String
        Dim tbl() As String
        '
        Hispanique_Gamme = ""
        a = TabCadDegrés.Item(ind).Text
        '
        Select Case Trim(a)
            Case "V", "IV"
                Hispanique_Gamme = Trim(Det_TonaCours2() + " " + "Maj")
            Case "III", "VI"
                c = Trim(ComboBox2.Text)
                If Langue = "fr" Then
                    c = TradAcc_LatAngl(Trim(ComboBox2.Text))
                End If
                '
                tbl = Split(Trim(c), " ")
                Hispanique_Gamme = Trim(tbl(0) + " " + "MinH")
        End Select

    End Function
    Sub Cad_RAZ_CouleurMarquée()
        Dim i As Integer


        If EnChargement = False Then
            If ModeSimple_Cadence = "Maj" Then
                For i = 0 To 4
                    TabCad.Item(i).BackColor = Couleur_Accord_Majeur
                    TabCad.Item(i).ForeColor = Color.Black
                    TabCadDegrés.Item(i).BackColor = Color.Gold 'Couleur_Accord_Majeur
                    TabCadDegrés.Item(i).ForeColor = Color.Black
                Next i
            Else
                For i = 0 To 4
                    TabCad.Item(i).BackColor = Couleur_Accord_Mineur
                    TabCad.Item(i).ForeColor = Color.Black
                    TabCadDegrés.Item(i).BackColor = Color.Gold 'Couleur_Accord_Mineur
                    TabCadDegrés.Item(i).ForeColor = Color.Black
                Next i
            End If
        End If
    End Sub
    Sub Cad_MarquerAccord(ind As Integer)
        Dim a As String
        Dim IndexDegré As Integer

        If Trim(TabCad.Item(ind).Text) <> "" And TabCad.Item(ind).Text <> "___" Then

            ' Mise à jour de la couleur
            ' *************************
            Cad_RAZ_CouleurMarquée()
            '
            TabCad.Item(ind).BackColor = Couleur_Accord_Marqué


            '
            TabCad.Item(ind).ForeColor = Color.Yellow


            '
            ' Mise à jour de l'accord pour écriture
            ' *************************************
            a = TabCadDegrés.Item(ind).Text
            If Cad_OrigineAccord = Modes.Cadence_Majeure Then
                IndexDegré = Det_IndexDegré(a)
            Else
                IndexDegré = Det_IndexDegréMin(a)
            End If
            '
            CAD_TableCoursAcc(IndexDegré).Marqué = True
            AccordMarqué = Trim(TabCad.Item(ind).Text)
            Entrée_Accord = AccordMarqué
            'la
            TabCad.Item(ind).Refresh()
        End If
        '
        ' OrigineAccord est mis à jour dans les comobobox3 et 4 de cadences majeures et mineures
    End Sub


    Sub OuvrirCadences(IndicCadences As String)
        Select Case IndicCadences
            Case "Anatole"
                Cad_OrigineAccord = Modes.Cadence_Majeure
                Cad_AnatoleMaj()
                Label28.Text = "Anatole"
                ComboBox3.SelectAll()

                '
            Case "Complète"
                Cad_OrigineAccord = Modes.Cadence_Majeure
                Cad_ComplèteMaj()
                Label28.Text = "Complète"
                ComboBox3.SelectAll()


            Case "Complete"
                Cad_OrigineAccord = Modes.Cadence_Majeure
                Cad_ComplèteMaj()
                Label28.Text = "Complete"
                ComboBox3.SelectAll()

                '
            Case "2-5-1"
                Cad_OrigineAccord = Modes.Cadence_Majeure
                Cad_251Maj()
                Label28.Text = "2-5-1"
                ComboBox3.SelectAll()

                '
            Case "Demi"
                Cad_OrigineAccord = Modes.Cadence_Majeure
                Cad_DemiMaj()
                Label28.Text = "Demi"
                ComboBox3.SelectAll()


            Case "Half"
                Cad_OrigineAccord = Modes.Cadence_Majeure
                Cad_DemiMaj()
                Label28.Text = "Half"
                ComboBox3.SelectAll()

                '
            Case "Parfaite"
                Cad_OrigineAccord = Modes.Cadence_Majeure
                Cad_ParfaiteMaj()
                Label28.Text = "Parfaite"
                ComboBox3.SelectAll()


            Case "Perfect"
                Cad_OrigineAccord = Modes.Cadence_Majeure
                Cad_ParfaiteMaj()
                Label28.Text = "Perfect"
                ComboBox3.SelectAll()

                '
            Case "Plagale"
                Cad_OrigineAccord = Modes.Cadence_Majeure
                Cad_PlagaleMaj()
                Label28.Text = "Plagale"
                ComboBox3.SelectAll()


            Case "Plagal"
                Cad_OrigineAccord = Modes.Cadence_Majeure
                Cad_PlagaleMaj()
                Label28.Text = "Plagal"
                ComboBox3.SelectAll()

                '
            Case "Plagale2"
                Cad_OrigineAccord = Modes.Cadence_Majeure
                Cad_Plagale2Maj()
                Label28.Text = "Plagale2"
                ComboBox3.SelectAll()


            Case "Plagal2"
                Cad_OrigineAccord = Modes.Cadence_Majeure
                Cad_Plagale2Maj()
                Label28.Text = "Plagal2"
                ComboBox3.SelectAll()

                '
            Case "Rompue"
                Cad_OrigineAccord = Modes.Cadence_Majeure
                Cad_RompueMaj()
                Label28.Text = "Rompue"
                ComboBox3.SelectAll()


            Case "Broken"
                Cad_OrigineAccord = Modes.Cadence_Majeure
                Cad_RompueMaj()
                Label28.Text = "Broken"
                ComboBox3.SelectAll()

                '
            Case "Rompue2"
                Cad_OrigineAccord = Modes.Cadence_Majeure
                Cad_Rompue2Maj()
                Label28.Text = "Rompue2"
                ComboBox3.SelectAll()


            Case "Broken2"
                Cad_OrigineAccord = Modes.Cadence_Majeure
                Cad_Rompue2Maj()
                Label28.Text = "Broken2"
                ComboBox3.SelectAll()

                '
            Case "Rompue3"
                Cad_OrigineAccord = Modes.Cadence_Majeure
                Cad_Rompue3Maj()
                Label28.Text = "Rompue3"
                ComboBox3.SelectAll()


            Case "Broken3"
                Cad_OrigineAccord = Modes.Cadence_Majeure
                Cad_Rompue3Maj()
                Label28.Text = "Broken3"
                ComboBox3.SelectAll()

            Case "Modale"
                Cad_OrigineAccord = Modes.Cadence_Majeure
                Cad_ModaleMaj()
                Label28.Text = "Modale"
                ComboBox3.SelectAll()


            Case "Modal"
                Cad_OrigineAccord = Modes.Cadence_Majeure
                Cad_ModaleMaj()
                Label28.Text = "Modal"
                ComboBox3.SelectAll()

                '
            Case "Modale2"
                Cad_OrigineAccord = Modes.Cadence_Majeure
                Cad_Modale2Maj()
                Label28.Text = "Modale2"
                ComboBox3.SelectAll()

            Case "Modal2"
                Cad_OrigineAccord = Modes.Cadence_Majeure
                Cad_Modale2Maj()
                Label28.Text = "Modal2"
                ComboBox3.SelectAll()

                '
            Case "Modale3"
                Cad_OrigineAccord = Modes.Cadence_Majeure
                Cad_Modale3Maj()
                Label28.Text = "Modale3"
                ComboBox3.SelectAll()


            Case "Modal3"
                Cad_OrigineAccord = Modes.Cadence_Majeure
                Cad_Modale3Maj()
                Label28.Text = "Modal3"
                ComboBox3.SelectAll()


            Case "Napolitaine"
                Cad_OrigineAccord = Modes.Cadence_Majeure
                Cad_NapolitaineMaj()
                Label28.Text = "Napolitaine"


            Case "Napolitan"
                Cad_OrigineAccord = Modes.Cadence_Majeure
                Cad_NapolitaineMaj()
                Label28.Text = "Napolitan"
                ComboBox3.SelectAll()


            Case "Anatole Min"
                Cad_OrigineAccord = Modes.Cadence_Mineure
                Cad_AnatoleMin()
                Label28.Text = "Anatole Min"
                ComboBox4.SelectAll()


            Case "Pseudo 2-5-1"
                Cad_OrigineAccord = Modes.Cadence_Mineure
                Cad_Pseudo251Min()
                Label28.Text = "Pseudo 2-5-1"
                ComboBox4.SelectAll()


            Case "Plagale", "Plagale Min"
                Cad_OrigineAccord = Modes.Cadence_Mineure
                Cad_PLagalMin()
                Label28.Text = "Plagale Min"
                ComboBox4.SelectAll()


            Case "Plagal", "Minor Plagal"
                Cad_OrigineAccord = Modes.Cadence_Mineure
                Cad_PLagalMin()
                Label28.Text = "Minor Plagal"
                ComboBox4.SelectAll()
                '
            Case "Hispanique"
                Cad_OrigineAccord = Modes.Cadence_Mixte
                OrigineAccord = Modes.Cadence_Mixte
                Mode_Cadence = Cad_OrigineAccord
                Cad_HispaniqueMixte()
                Label28.Text = "Hispanique"
                ComboBox4.SelectAll()

            Case "Hispanic"
                Cad_OrigineAccord = Modes.Cadence_Mixte
                OrigineAccord = Modes.Cadence_Mixte
                Mode_Cadence = Cad_OrigineAccord
                Cad_HispaniqueMixte()
                Label28.Text = "Hispanic"
                ComboBox4.SelectAll()
        End Select

    End Sub



    Sub Cad_AnatoleMaj()
        If EnChargement = False Then
            '
            CAD_Init_Aff()
            '
            TabCadDegrés.Item(0).Text = "I"
            TabCadDegrés.Item(0).Visible = True ' degré
            TabCad.Item(0).Visible = True ' nom de l'accord
            '
            TabCadDegrés.Item(1).Text = "VI"
            TabCadDegrés.Item(1).Visible = True ' degré
            TabCad.Item(1).Visible = True ' nom de l'accord


            TabCadDegrés.Item(2).Text = "II"
            TabCadDegrés.Item(2).Visible = True ' degré
            TabCad.Item(2).Visible = True ' nom de l'accord


            TabCadDegrés.Item(3).Text = "V"
            TabCadDegrés.Item(3).Visible = True ' degré
            TabCad.Item(3).Visible = True ' nom de l'accord


            TabCadDegrés.Item(4).Text = "I"
            TabCadDegrés.Item(4).Visible = True ' degré
            TabCad.Item(4).Visible = True ' nom de l'accord

            '
            'Label28.Size = New Size(535, 33)
            '
            CAD_Maj_TableGlobalAcc()
        End If
    End Sub
    Sub Cad_Forme2()
        If EnChargement = False Then
            '
            CAD_Init_Aff()
            '
            TabCadDegrés.Item(0).Text = "I"
            TabCadDegrés.Item(0).Visible = True ' degré
            TabCad.Item(0).Visible = True ' nom de l'accord
            '
            TabCadDegrés.Item(1).Text = "VI"
            TabCadDegrés.Item(1).Visible = True ' degré
            TabCad.Item(1).Visible = True ' nom de l'accord


            TabCadDegrés.Item(2).Text = "IV"
            TabCadDegrés.Item(2).Visible = True ' degré
            TabCad.Item(2).Visible = True ' nom de l'accord


            TabCadDegrés.Item(3).Text = "V"
            TabCadDegrés.Item(3).Visible = True ' degré
            TabCad.Item(3).Visible = True ' nom de l'accord


            TabCadDegrés.Item(4).Text = "I"
            TabCadDegrés.Item(4).Visible = True ' degré
            TabCad.Item(4).Visible = True ' nom de l'accord

            '
            'Label28.Size = New Size(535, 33)
            '
            CAD_Maj_TableGlobalAcc()
        End If
    End Sub
    Sub Cad_Forme3()
        If EnChargement = False Then
            '
            CAD_Init_Aff()
            '
            TabCadDegrés.Item(0).Text = "I"
            TabCadDegrés.Item(0).Visible = True ' degré
            TabCad.Item(0).Visible = True ' nom de l'accord
            '
            TabCadDegrés.Item(1).Text = "V"
            TabCadDegrés.Item(1).Visible = True ' degré
            TabCad.Item(1).Visible = True ' nom de l'accord


            TabCadDegrés.Item(2).Text = "VI"
            TabCadDegrés.Item(2).Visible = True ' degré
            TabCad.Item(2).Visible = True ' nom de l'accord


            TabCadDegrés.Item(3).Text = "IV"
            TabCadDegrés.Item(3).Visible = True ' degré
            TabCad.Item(3).Visible = True ' nom de l'accord


            TabCadDegrés.Item(4).Text = "I"
            TabCadDegrés.Item(4).Visible = True ' degré
            TabCad.Item(4).Visible = True ' nom de l'accord

            '
            'Label28.Size = New Size(535, 33)
            '
            CAD_Maj_TableGlobalAcc()
        End If
    End Sub
    Sub Cad_AnatoleMin()
        If EnChargement = False Then
            '
            CAD_Init_Aff()
            '
            TabCadDegrés.Item(0).Text = "VI"
            TabCadDegrés.Item(0).Visible = True ' degré
            TabCad.Item(0).Visible = True ' nom de l'accord

            '
            TabCadDegrés.Item(1).Text = "IV"
            TabCadDegrés.Item(1).Visible = True ' degré
            TabCad.Item(1).Visible = True ' nom de l'accord


            TabCadDegrés.Item(2).Text = "VII"
            TabCadDegrés.Item(2).Visible = True ' degré
            TabCad.Item(2).Visible = True ' nom de l'accord


            TabCadDegrés.Item(3).Text = "III"
            TabCadDegrés.Item(3).Visible = True ' degré
            TabCad.Item(3).Visible = True ' nom de l'accord


            TabCadDegrés.Item(4).Text = "VI"
            TabCadDegrés.Item(4).Visible = True ' degré
            TabCad.Item(4).Visible = True ' nom de l'accord

            '
            'Label28.Size = New Size(535, 33)
            '
            CAD_Maj_TableGlobalAcc()
            '
        End If
    End Sub

    Sub Cad_PLagalMin()
        If EnChargement = False Then
            '
            CAD_Init_Aff()
            '
            TabCadDegrés.Item(0).Text = "III"
            TabCadDegrés.Item(0).Visible = True ' degré
            TabCad.Item(0).Visible = True ' nom de l'accord


            TabCadDegrés.Item(1).Text = "VI"
            TabCadDegrés.Item(1).Visible = True ' degré
            TabCad.Item(1).Visible = True ' nom de l'accord


            TabCadDegrés.Item(2).Text = "III"
            TabCadDegrés.Item(2).Visible = True ' degré
            TabCad.Item(2).Visible = True ' nom de l'accord

            '
            'Label28.Size = New Size(316, 33)
            '
            CAD_Maj_TableGlobalAcc()
            '
        End If
    End Sub
    Sub Cad_Pseudo251Min()
        If EnChargement = False Then
            '
            CAD_Init_Aff()
            '
            TabCadDegrés.Item(0).Text = "VI"
            TabCadDegrés.Item(0).Visible = True ' degré
            TabCad.Item(0).Visible = True ' nom de l'accord


            TabCadDegrés.Item(1).Text = "VII"
            TabCadDegrés.Item(1).Visible = True ' degré
            TabCad.Item(1).Visible = True ' nom de l'accord


            TabCadDegrés.Item(2).Text = "III"
            TabCadDegrés.Item(2).Visible = True ' degré
            TabCad.Item(2).Visible = True ' nom de l'accord


            TabCadDegrés.Item(3).Text = "VI"
            TabCadDegrés.Item(3).Visible = True ' degré
            TabCad.Item(3).Visible = True ' nom de l'accord

            '
            'Label28.Size = New Size(426, 33)
            '
            CAD_Maj_TableGlobalAcc()
            '
        End If
    End Sub
    Sub Cad_HispaniqueMixte()
        If EnChargement = False Then
            Cad_OrigineAccord = Modes.Cadence_Mixte
            Mode_Cadence = Cad_OrigineAccord
            CAD_Init_Aff()
            '
            ' partie majeure
            TabCadDegrés.Item(0).Text = "VI"
            TabCadDegrés.Item(0).Visible = True ' degré
            TabCad.Item(0).Visible = True ' nom de l'accord


            TabCadDegrés.Item(1).Text = "V"
            TabCadDegrés.Item(1).Visible = True ' degré
            TabCad.Item(1).Visible = True ' nom de l'accord


            TabCadDegrés.Item(2).Text = "IV"
            TabCadDegrés.Item(2).Visible = True ' degré
            TabCad.Item(2).Visible = True ' nom de l'accord

            ' partie mineure
            TabCadDegrés.Item(3).Text = "III"
            TabCadDegrés.Item(3).Visible = True ' degré
            TabCad.Item(3).Visible = True ' nom de l'accord


            TabCadDegrés.Item(4).Text = "VI"
            TabCadDegrés.Item(4).Visible = True ' degré
            TabCad.Item(4).Visible = True ' nom de l'accord

            '
            'Label28.Size = New Size(535, 33)
            '
            CAD_Maj_TableGlobalAcc()
        End If
    End Sub

    Private Sub CAD_Renversement1_Click(sender As Object, e As EventArgs) Handles CAD_Renversement1.Click
        CAD_Renversement1.Checked = True
        '
        CAD_Renversement2.Checked = False
        CAD_Renversement3.Checked = False
        CAD_Renversement4.Checked = False
        CAD_Renversement5.Checked = False
        '
        ' Maj_ChoixRenv(1)
        CAD_TableCoursAcc(CAD_LabelCours).RenvChoisi = 0
        '
        CAD_RAZEtendre()
        '
    End Sub

    Private Sub CAD_Renversement2_Click(sender As Object, e As EventArgs) Handles CAD_Renversement2.Click
        CAD_Renversement1.Checked = False
        '
        CAD_Renversement2.Checked = True
        CAD_Renversement3.Checked = False
        CAD_Renversement4.Checked = False
        CAD_Renversement5.Checked = False
        '
        ' Maj_ChoixRenv(1)
        CAD_TableCoursAcc(CAD_LabelCours).RenvChoisi = 1
        '
        CAD_RAZEtendre()
    End Sub

    Private Sub CAD_Renversement3_Click(sender As Object, e As EventArgs) Handles CAD_Renversement3.Click
        CAD_Renversement1.Checked = False
        '
        CAD_Renversement2.Checked = False
        CAD_Renversement3.Checked = True
        CAD_Renversement4.Checked = False
        CAD_Renversement5.Checked = False
        '
        ' Maj_ChoixRenv(1)
        CAD_TableCoursAcc(CAD_LabelCours).RenvChoisi = 2
        '
        CAD_RAZEtendre()
    End Sub

    Private Sub CAD_Renversement4_Click(sender As Object, e As EventArgs) Handles CAD_Renversement4.Click
        CAD_Renversement1.Checked = False
        '
        CAD_Renversement2.Checked = False
        CAD_Renversement3.Checked = False
        CAD_Renversement4.Checked = True
        CAD_Renversement5.Checked = False
        '
        ' Maj_ChoixRenv(1)
        CAD_TableCoursAcc(CAD_LabelCours).RenvChoisi = 3
        '
        CAD_RAZEtendre()
    End Sub

    Private Sub CAD_Renversement5_Click(sender As Object, e As EventArgs) Handles CAD_Renversement5.Click
        CAD_Renversement1.Checked = False
        '
        CAD_Renversement2.Checked = False
        CAD_Renversement3.Checked = False
        CAD_Renversement4.Checked = False
        CAD_Renversement5.Checked = True
        '
        ' Maj_ChoixRenv(1)
        CAD_TableCoursAcc(CAD_LabelCours).RenvChoisi = 4
        '
        CAD_RAZEtendre()
    End Sub
    Sub CAD_RAZEtendre()
        CAD_EtendreNote1.Checked = False
        CAD_EtendreNote2.Checked = False
        CAD_EtendreNote3.Checked = False
        CAD_EtendreNote4.Checked = False
        CAD_EtendreNote5.Checked = False
        '
        CAD_TableCoursAcc(CAD_LabelCours).EtendreChecked(0) = False
        CAD_TableCoursAcc(CAD_LabelCours).EtendreChecked(1) = False
        CAD_TableCoursAcc(CAD_LabelCours).EtendreChecked(2) = False
        CAD_TableCoursAcc(CAD_LabelCours).EtendreChecked(3) = False
        CAD_TableCoursAcc(CAD_LabelCours).EtendreChecked(4) = False
    End Sub
    Sub CAD_Maj_ChoixOctave(OctaveChoisie As Integer, Octave As String)
        Dim i As Integer
        Dim a As String
        Dim IndexDegré As Integer

        For i = 0 To 4
            'If TabTonsSelect.Item(i).Checked = True Then
            If TabCad.Item(i).BackColor = Color.Red Then
                a = TabCadDegrés.Item(i).Text
                IndexDegré = Det_IndexDegré(a)
                '
                CAD_TableCoursAcc(IndexDegré).OctaveChoisie = OctaveChoisie
                CAD_TableCoursAcc(IndexDegré).Octave = Octave
                '
            End If
        Next i
    End Sub

    Private Sub CAD_OctavePlus1_Click(sender As Object, e As EventArgs) Handles CAD_OctavePlus1.Click
        CAD_OctavePlus1.Checked = True
        CAD_Octave0.Checked = False
        CAD_OctaveMoins1.Checked = False
        CAD_OctaveMoins2.Checked = False
        '
        CAD_Maj_ChoixOctave(0, "+1")
    End Sub

    Private Sub CAD_Octave0_Click(sender As Object, e As EventArgs) Handles CAD_Octave0.Click
        CAD_OctavePlus1.Checked = False
        CAD_Octave0.Checked = True
        CAD_OctaveMoins1.Checked = False
        CAD_OctaveMoins2.Checked = False
        '
        CAD_Maj_ChoixOctave(1, "0")
    End Sub

    Private Sub CAD_OctaveMoins1_Click(sender As Object, e As EventArgs) Handles CAD_OctaveMoins1.Click
        CAD_OctavePlus1.Checked = False
        CAD_Octave0.Checked = False
        CAD_OctaveMoins1.Checked = True
        CAD_OctaveMoins2.Checked = False
        '
        CAD_Maj_ChoixOctave(2, "-1")
    End Sub
    Private Sub CAD_OctaveMoins2_Click(sender As Object, e As EventArgs) Handles CAD_OctaveMoins2.Click
        CAD_OctavePlus1.Checked = False
        CAD_Octave0.Checked = False
        CAD_OctaveMoins1.Checked = False
        CAD_OctaveMoins2.Checked = True
        '
        CAD_Maj_ChoixOctave(3, "-2")
    End Sub

    Private Sub ComboMidiOut_SelectedIndexChanged_old(sender As Object, e As EventArgs) Handles ComboMidiOut.SelectedIndexChanged
        Dim a As String
        '
        Try
            If EnChargement = False Then
                ChoixSortieMidi = ComboMidiOut.SelectedIndex
                '
                a = Trim(Str(ChoixSortieMidi))
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CalquesMIDI\HyperArp\Préférences", "ChoixSortieMIDI", a)
            End If
            '
        Catch ex As Exception
            If LangueIHM = "fr" Then
                MessageHV.PTitre = "Avertissement"
                MessageHV.PContenuMess = "Problème de ressources MIDI : vérifier votre interface MIDI" + ex.Message
                MessageHV.PTypBouton = "OK"
            Else
                MessageHV.PTitre = "Warning"
                MessageHV.PContenuMess = "MIDI resource problem: check your MIDI interface" + ex.Message
                MessageHV.PTypBouton = "OK"
            End If
            Cacher_FormTransparents()
            MessageHV.ShowDialog()
            End
        End Try
    End Sub
    Private Sub ComboMidiOut_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboMidiOut.SelectedIndexChanged
        Dim a As String
        Dim titre As String
        Dim sauvChoixSortieMidi As Integer = ChoixSortieMidi
        '
        Try
            If EnChargement = False Then
                ChoixSortieMidi = ComboMidiOut.SelectedIndex
                '
                a = Trim(Str(ChoixSortieMidi))
                ' test de l'interface choisie : si problème le catch entre en jeu
                If Not (SortieMidi.Item(ChoixSortieMidi).IsOpen) Then
                    '
                    SortieMidi.Item(ChoixSortieMidi).Open()
                    SortieMidi.Item(ChoixSortieMidi).Close()
                Else
                    SortieMidi.Item(ChoixSortieMidi).Close()
                End If
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CalquesMIDI\HyperVoicing\Préférences", "ChoixSortieMIDI", a)
            End If
            '
        Catch ex As Exception

            If Module1.LangueIHM = "fr" Then
                titre = "Veuillez noter que ..."
                Avertis = "La Sortie MIDI " + ComboMidiOut.Text + " est  indisponible. Veuillez choisir une autre sortie MIDI"
            Else
                titre = "Please note that ..."
                Avertis = "The MIDI output " + ComboMidiOut.Text + " is not avaliable. Please choose an other MIDI Output"
            End If
            MessageHV.PContenuMess = Avertis
            MessageHV.Titre = titre
            MessageHV.PTypBouton = "OK"
            MessageHV.ShowDialog()
            '
        End Try
    End Sub
    Private Sub NoteRacine_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub TabControl4_SelectedIndexChanged(sender As Object, e As EventArgs)
        'Dim tcont As TabControl = sender
        'Dim a As String


        ''OngletCours = Val(tcont.SelectedTab.Tag)
        'a = tcont.SelectedTab.Text


    End Sub
    Private Sub Timer_Aff_MidiIn_Tick(sender As Object, e As EventArgs) Handles Timer_Aff_MidiIn.Tick

        If AfficherNote = True Then
            'LabelPianoMidiIn.Item(N_Note).Refresh()
            'LabelPianoMidiIn.Item(N_Note).BringToFront()
            LabelPianoMidiIn.Item(N_Note).BackColor = Color.Red
        Else
            If Det_NoteEstDiése(N_Note) = "#" Then
                LabelPianoMidiIn.Item(N_Note).BackColor = Color.Black
            Else
                LabelPianoMidiIn.Item(N_Note).BackColor = Color.White
            End If
        End If
    End Sub
    Private Function Det_NoteEstDiése(Note As Byte) As String
        Dim i, j, k As Integer

        Det_NoteEstDiése = ""

        For i = 0 To 127 Step 12
            For j = 0 To 11
                k = i + j
                If k = Note Then
                    Select Case j
                        Case 1, 3, 6, 8, 10
                            Det_NoteEstDiése = "#"
                        Case Else
                            Det_NoteEstDiése = ""
                    End Select
                End If
            Next j
        Next i
    End Function




    Sub EcritureAccordDsGrid2_2(Tonalité As String, Accord As String, Gamme As String, Mode As String,
                            m As Integer, t As Integer, ct As Integer) ' les variables Tonalités, Accord, Gamme et Mode sont exprimées ici en Anglais et en #
        Dim tbl() As String
        Dim a, b As String

        DerGridCliquée = GridCours.Grid2
        ' Pour CTRL Z
        ' ***********
        '
        ' Ecriture accords à partir du tableau des tonalités
        ' **************************************************
        '

        b = TonalitéDsClef(Tonalité) ' cette fonction détermine si la tonalité exprimée en # existe bien dans les tonalités possibles (par exemple A# n'existe pas, c'est Bb qu'il faut choisir)
        If b = "b" Then ' si oui, il est nécessaire de tout transformer en "b"
            Tonalité = TradAcc_DieseEnBem(Tonalité)
            Accord = TradAcc_DieseEnBem(Accord)
            Gamme = TradAcc_DieseEnBem(Gamme)
            Mode = TradAcc_DieseEnBem(Mode)
        End If
        TableEventH(m, t, ct).Tonalité = Trim(Tonalité) '  Trim(ComboBox1.Text) 'Entrée_Tonalité
        TableEventH(m, t, ct).Accord = Trim(Accord)     ' Entrée_Accord
        TableEventH(m, t, ct).Gamme = Trim(Gamme)
        TableEventH(m, t, ct).Mode = Trim(Mode)        ' Mode = Gamme - le mode n'est pas affiché mais je garde l'info por le moment

        '
        ' ECRITURE dans Grid2
        ' *******************
        Grid2.AutoRedraw = False
        If Trim(Grid2.Cell(1, m).Text) = "" Then
            Grid2.Cell(1, m).Alignment = FlexCell.AlignmentEnum.CenterCenter
            Grid2.Cell(1, m).Text = Accord
            If Langue = "fr" Then
                Grid2.Cell(1, m).Text = TradAcc_AnglLat(Accord)
            End If
            Grid2.Cell(1, m).Locked = True
            Grid2.ReadonlyFocusRect = FlexCell.FocusRectEnum.Solid
        Else
            Grid2.Cell(1, m).Text = ChaineAccord(m)
        End If
        Grid2.Refresh()
        Grid2.AutoRedraw = True
        '
        tbl = Split(Trim(Tonalité), " ")
        a = tbl(0) ' note de la tonalité courante
        If m <> 1 Then
            Grid2.AutoRedraw = False
            '
            Grid2.Cell(1, m).BackColor = DicoCouleur.Item(a) ' la couleur est fonction de la tonalité
            Grid2.Cell(1, m).ForeColor = DicoCouleurLettre.Item(a)
            '
            Grid2.Refresh()
            Grid2.AutoRedraw = True
            '
        End If
        '
        Refresh()
    End Sub
    'Private Sub Grid3_DragEnter(sender As Object, e As DragEventArgs)
    'If (e.Data.GetDataPresent(DataFormats.Text)) Then
    '        e.Effect = DragDropEffects.Copy
    'Else
    '        e.Effect = DragDropEffects.None
    'End If
    'End Sub





    Private Sub EtendreNote1_Click(sender As Object, e As EventArgs) Handles EtendreNote1.Click
        If EtendreNote1.Checked = True Then
            EtendreNote1.Checked = False
        Else
            EtendreNote1.Checked = True
        End If
        Maj_ChoixEtendre(0, EtendreNote1.Checked)
    End Sub

    Private Sub EtendreNote2_Click(sender As Object, e As EventArgs) Handles EtendreNote2.Click
        If EtendreNote2.Checked = True Then
            EtendreNote2.Checked = False
        Else
            EtendreNote2.Checked = True
        End If
        '
        Maj_ChoixEtendre(1, EtendreNote2.Checked)
    End Sub

    Private Sub EtendreNote3_Click(sender As Object, e As EventArgs) Handles EtendreNote3.Click
        If EtendreNote3.Checked = True Then
            EtendreNote3.Checked = False
        Else
            EtendreNote3.Checked = True
        End If
        Maj_ChoixEtendre(2, EtendreNote3.Checked)
    End Sub

    Private Sub EtendreNote4_Click(sender As Object, e As EventArgs) Handles EtendreNote4.Click
        If EtendreNote4.Checked = True Then
            EtendreNote4.Checked = False
        Else
            EtendreNote4.Checked = True
        End If
        Maj_ChoixEtendre(3, EtendreNote4.Checked)
    End Sub

    Private Sub EtendreNote5_Click(sender As Object, e As EventArgs) Handles EtendreNote5.Click
        If EtendreNote5.Checked = True Then
            EtendreNote5.Checked = False
        Else
            EtendreNote5.Checked = True
        End If
        Maj_ChoixEtendre(4, EtendreNote5.Checked)
    End Sub
    Sub Maj_ChoixEtendre(EtendreChoisi As Integer, EtendreValeur As Boolean)
        Dim i As Integer
        Dim ligne As Integer
        Dim degré As Integer

        For i = 0 To 20
            'If TabTonsSelect.Item(i).Checked = True Then
            If TabTons.Item(i).BackColor = Color.Red Then
                ligne = Det_LigneTableGlobale(i)
                degré = Det_IndexDegré2(i)
                '
                TableCoursAcc(ligne, degré).EtendreChecked(EtendreChoisi) = EtendreValeur
                '
            End If
        Next i
    End Sub

    Private Sub CAD_EtendreNote1_Click(sender As Object, e As EventArgs) Handles CAD_EtendreNote1.Click
        If CAD_EtendreNote1.Checked = True Then
            CAD_EtendreNote1.Checked = False
        Else
            CAD_EtendreNote1.Checked = True
        End If
        CAD_TableCoursAcc(CAD_LabelCours).EtendreChecked(0) = CAD_EtendreNote1.Checked
    End Sub
    Private Sub CAD_EtendreNote2_Click(sender As Object, e As EventArgs) Handles CAD_EtendreNote2.Click
        If CAD_EtendreNote2.Checked = True Then
            CAD_EtendreNote2.Checked = False
        Else
            CAD_EtendreNote2.Checked = True
        End If
        CAD_TableCoursAcc(CAD_LabelCours).EtendreChecked(1) = CAD_EtendreNote2.Checked
    End Sub
    Private Sub CAD_EtendreNote3_Click(sender As Object, e As EventArgs) Handles CAD_EtendreNote3.Click
        If CAD_EtendreNote3.Checked = True Then
            CAD_EtendreNote3.Checked = False
        Else
            CAD_EtendreNote3.Checked = True
        End If
        CAD_TableCoursAcc(CAD_LabelCours).EtendreChecked(2) = CAD_EtendreNote3.Checked
    End Sub
    Private Sub CAD_EtendreNote4_Click(sender As Object, e As EventArgs) Handles CAD_EtendreNote4.Click
        If CAD_EtendreNote4.Checked = True Then
            CAD_EtendreNote4.Checked = False
        Else
            CAD_EtendreNote4.Checked = True
        End If
        CAD_TableCoursAcc(CAD_LabelCours).EtendreChecked(3) = CAD_EtendreNote4.Checked
    End Sub
    Private Sub CAD_EtendreNote5_Click(sender As Object, e As EventArgs) Handles CAD_EtendreNote5.Click
        If CAD_EtendreNote5.Checked = True Then
            CAD_EtendreNote5.Checked = False
        Else
            CAD_EtendreNote5.Checked = True
        End If
        CAD_TableCoursAcc(CAD_LabelCours).EtendreChecked(4) = CAD_EtendreNote5.Checked
    End Sub
    <SecurityPermission(SecurityAction.LinkDemand, ControlDomainPolicy:=True)>
    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        ' remarque : l'évènement shown arrive toujours après l'évènement Load. La variable 'FichierEntréSurClic'
        ' est récupérée dans le présent évènement Shown.
        '
        Dim aa As AppDomainSetup = AppDomain.CurrentDomain.SetupInformation
        Try
            If Not (aa.ActivationArguments Is Nothing) Then
                If (aa.ActivationArguments.ActivationData.Count >= 1) Then
                    FichierEntréSurClic = aa.ActivationArguments.ActivationData(0)
                    Ouvrir2()
                End If
            End If
        Catch ex As Exception
            'MessageBox.Show("Evt Form1_Shown :  " + ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Calcul_AutoVoicingZ : calcul des voicing de chaque accord suivant l'algorithme des notes communes.
    '''                       cette méthode tient compte de la racine d'une accord défini par des zones dans l'onglet "Zones de voicing".
    ''' </summary>
    Public Sub Calcul_AutoVoicingZ_old()
        Dim m As Integer
        Dim a As String
        Dim i As Integer
        Dim tbl() As String
        '
        ' ********************************************************************************************************
        ' *                                                                                                      *
        ' *  TableNotesAccordsZ(m, t, ct) : Table contenant les voicings calculés                                        *
        ' *  TableEventH(m,1,1) : table contenant les caractéristiques d'une mesure : accord, gamme, tonalité .. *
        ' *                                                                                                      * 
        ' ********************************************************************************************************
        '
        ' Principe de l'algorithme des notes communes : 
        '           1 - deux notes communes à 2 accords concécutifs se trouvent toujours sur la même octave.
        '           2 - la note la plus basse d'un accord est conditionnée par la racine de la zone auquel l'accord appartient

        Try
            RAZ_TableNotesAccordsZ()
            For i = 0 To NbZones
                If TZone(i).DébutZ <> "---" Then
                    mdeb = Val(TZone(i).DébutZ)
                    mfin = Val(TZone(i).TermeZ)
                    '
                    For m = mdeb To mfin
                        a = TableEventH(m, 1, 1).Accord
                        TableNotesAccordsZ(m, 1, 1) = ""
                        If Trim(a) <> "" Then
                            tbl = Split(a)
                            tbl(0) = TradD(tbl(0)) ' les accords choisis sont exprimés en # dans la table
                            If UBound(tbl) > 0 Then a = Trim(tbl(0) + " " + tbl(1))
                            a = Det_NotesAccord(Trim(a))
                            a = Trad_ListeNotesEnD(Trim(a), "-") ' la traduction est nécessaire car Maj_NotesCommunes ne traite que des notes en #
                            TableNotesAccordsZ(m, 1, 1) = Trim(Maj_NotesCommunes(Trim(a), i)) ' construction de la suite de notes avec les accords
                        End If
                    Next m
                End If
            Next i
            Maj_AffVoicing() ' Pour onglet Zones de voicing/Visu notes
            '
        Catch ex As Exception
            If LangueIHM = "fr" Then
                MessageHV.PTitre = "Avertissement"
                MessageHV.PContenuMess = "Erreur interne : procédure Calcul_AutoVoicingZ : " + ex.Message
                MessageHV.PTypBouton = "OK"
            Else
                MessageHV.PTitre = "Warning"
                MessageHV.PContenuMess = "Internal Error : procédure Calcul_AutoVoicingZ :  " + ex.Message
                MessageHV.PTypBouton = "OK"
            End If
            Cacher_FormTransparents()
            MessageHV.ShowDialog()
            End
        End Try
        'End If
    End Sub
    Public Sub Calcul_AutoVoicingZ()
        Dim m As Integer
        Dim a As String
        Dim tbl() As String
        Dim NbAcc As Integer = Det_NumDerAccord()
        'Dim AncienV As String
        '
        ' ********************************************************************************************************
        ' *                                                                                                      *
        ' *  TableNotesAccordsZ(m, t, ct) : Table contenant les voicings calculés                                        *
        ' *  TableEventH(m,1,1) : table contenant les caractéristiques d'une mesure : accord, gamme, tonalité .. *
        ' *                                                                                                      * 
        ' ********************************************************************************************************
        '
        ' Principe de l'algorithme des notes communes : 
        '           1 - deux notes communes à 2 accords concécutifs se trouvent toujours sur la même octave.
        '           2 - la note la plus basse d'un accord est conditionnée par la racine de la zone auquel l'accord appartient
        If NbAcc <> -1 Then

            Try
                RAZ_TableNotesAccordsZ()
                For m = 1 To NbAcc
                    TableNotesAccordsZ(m, 1, 1) = ""
                    a = TableEventH(m, 1, 1).Accord

                    If Trim(a) <> "" Then
                        tbl = Split(a)
                        tbl(0) = TradD(tbl(0)) ' les accords choisis sont exprimés en # dans la table
                        If UBound(tbl) > 0 Then a = Trim(tbl(0) + " " + tbl(1))
                        a = Det_NotesAccord(Trim(a))
                        a = Trad_ListeNotesEnD(Trim(a), "-") ' la traduction est nécessaire car Maj_NotesCommunes ne traite que des notes en #

                        'TableNotesAccordsZ(m, 1, 1) = Trim(Maj_NotesCommunes2(Trim(a), Trim(Grid2.Cell(12, m).Text))) ' construction de la suite de notes avec les accords
                        'TableNotesAccordsZ(m, 1, 1) = Trim(Maj_NotesCommunes2(Trim(a), Trim(TRacine.Text)))

                        TableNotesAccordsZ(m, 1, 1) = Trim(Maj_NotesCommunes2(Trim(a), TableEventH(m, 1, 1).Racine)) ' construction de la suite de notes avec les accords
                        b = TableNotesAccordsZ(m, 1, 1)
                        TableNotesAccordsZ(m, 1, 1) = Trim(Transform(b))
                    End If
                Next m
                '
                'Maj_AffVoicing() ' Pour onglet Zones de voicing/Visu notes
                '
                Maj_VueNotes()

            Catch ex As Exception
                If LangueIHM = "fr" Then
                    MessageHV.PTitre = "Avertissement"
                    MessageHV.PContenuMess = "Erreur interne : procédure Calcul_AutoVoicingZ : " + ex.Message + "Valeur de m = " + Str(m) + " TableEventH(m, 1, 1).Accord = " + TableEventH(m, 1, 1).Accord + "TableNotesAccordsZ(m, 1, 1) =" + TableNotesAccordsZ(m, 1, 1)
                    MessageHV.PTypBouton = "OK"
                Else
                    MessageHV.PTitre = "Warning"
                    MessageHV.PContenuMess = "Internal Error : procédure Calcul_AutoVoicingZ :  " + ex.Message + "Valeur de m = " + Str(m)
                    MessageHV.PTypBouton = "OK"
                End If
                Cacher_FormTransparents()
                MessageHV.ShowDialog()
                End
            End Try
        End If
    End Sub
    Private Function Transform(acc As String) As String
        Dim i As Integer
        Dim tbl() As String
        Dim Note_str As String
        Dim Note_int1 As Integer
        Dim Note_int2 As Integer

        Transform = Trim(acc)

        tbl = acc.Split()
        '
        Note_str = Trim(tbl(ComboBox11.SelectedIndex))
        Note_int1 = ValNoteCubase.IndexOf(UCase(Note_str))
        '
        Note_str = Trim(tbl(ComboBox12.SelectedIndex))
        Note_int2 = ValNoteCubase.IndexOf(UCase(Note_str))

        ' 4Notes : Etendre à quatre notes les accords de 3 notes
        ' *******************************************************
        If Fournotes.Checked And tbl.Count = 3 Then ' on agit uniquement sur les accord de 3 notes
            Note_int1 = Note_int1 + 12
            If Note_int1 <= 127 Then
                Note_str = ValNoteCubase(Note_int1)
                ReDim Preserve tbl(3)
                tbl(3) = LCase(Note_str)
                Transform = Join(tbl, " ")
                ' ordonner les notes
            End If

        End If
        '
        ' Basse -1 : Descendre d'une octave la 1ere note
        ' **********************************************
        If Bassemoins1.Checked Then
            Note_int2 = Note_int2 - 12 ' descendre d'une octave
            If Note_int2 >= 0 Then
                ' Calcul de la note
                Note_str = ValNoteCubase(Note_int2)
                For i = 0 To tbl.Count - 1
                    tbl(i) = UCase(tbl(i)) ' tout mettre en majuscule
                Next
                tbl(ComboBox12.SelectedIndex) = Trim(Note_str)
                ' ordonner les notes
                Dim tbl1(0 To tbl.Count - 1) As Integer
                For i = 0 To tbl.Count - 1
                    tbl1(i) = ValNoteCubase.IndexOf(tbl(i))
                Next
                Array.Sort(tbl1) 'tire des notes (sous forme chiffre)
                For i = 0 To tbl.Count - 1
                    tbl(i) = LCase(ValNoteCubase(tbl1(i)))
                Next
                Transform = Join(tbl, " ")
            End If
        End If
        '

    End Function

    Function Maj_NotesCommunes2(Renv As String, Root As String) As String
        Dim i, j, lg1, lg2 As Integer
        Dim tbl() As String
        Dim Racine As String
        Dim Octave As Integer
        Dim indRacine As Integer
        Dim CalcOctave As Integer
        Dim Retour As String
        Dim NombreNotesAccord As Integer
        Dim a As String

        Retour = ""
        '
        ' Séparation Note / Octave de la racine pour le calcul
        ' ****************************************************
        '
        a = Trim(Root)
        If Mid(a, Len(a) - 1, 1) <> "-" Then
            lg1 = Len(a) - 1
            lg2 = 1
        Else
            lg1 = Len(a) - 2
            lg2 = 2
        End If
        '
        Racine = Mid(a, 1, lg1) '
        Octave = Microsoft.VisualBasic.Right(a, lg2)
        '
        tbl = Split(Renv, "-")
        NombreNotesAccord = UBound(tbl) + 1
        '
        Init_Tessiture()
        '
        ' Maj des notes de l'accord dans le tableau Tessiture
        ' ***************************************************
        For i = 0 To UBound(tbl)
            For j = 0 To UBound(Tessiture)
                If tbl(i) = Tessiture(j).NoteTessiture Then
                    Tessiture(j).NoteAccord = tbl(i)
                End If
            Next j
        Next i
        '
        ' Détermination de l'index de la racine dans le tableau Tessiture
        ' ***************************************************************
        For i = 0 To UBound(Tessiture)
            If Racine = Tessiture(i).NoteTessiture Then
                indRacine = i
                Exit For
            End If
        Next i
        '
        ' Détermination du renversement avec l'octave
        ' *******************************************

        CalcOctave = indRacine
        For i = indRacine To UBound(Tessiture)
            If Tessiture(i).NoteAccord <> "" Then
                Retour = Trim(Retour + " " + Tessiture(i).NoteAccord + Trim(Str(Octave)))
            End If
            ' pour détermination d'un changement d'octave
            CalcOctave = CalcOctave + 1
            If CalcOctave > 11 Then
                Octave = Octave + 1 ' changement d'octave
                CalcOctave = 0
            End If
            '
            tbl = Split(Retour)
            '
        Next i
        Retour = ""
        For i = 0 To NombreNotesAccord - 1
            Retour = Retour + " " + tbl(i)
        Next
        '
        Return Trim(Retour)
    End Function
    Sub Maj_AffVoicing()
        Dim i, j As Integer
        Dim m As Integer
        Dim tbl() As String
        '
        If Not Module1.JeuxActif Then

            For i = 1 To Grid7.Rows - 1
                For j = 1 To Grid7.Cols - 1
                    Grid7.Cell(i, j).Text = ""
                    Grid7.Cell(i, j).BackColor = Color.White
                Next
            Next

            For m = 1 To nbMesures - 1
                If Trim(TableNotesAccordsZ(m, 1, 1)) <> "" Then
                    tbl = Split(Trim(TableNotesAccordsZ(m, 1, 1)))
                    j = 5
                    'For i = UBound(tbl) To 0 Step -1
                    For i = 0 To UBound(tbl)
                        If Trim(Grid7.Cell(j, m).Text) <> tbl(i) And Trim(Grid7.Cell(j, m).Text) <> "" Then
                            Grid7.Cell(j, m).BackColor = Color.Orange
                        Else
                            Grid7.Cell(j, m).BackColor = Color.White
                        End If
                        Grid7.Cell(j, m).Text = tbl(i)
                        j = j - 1
                    Next
                Else
                    For i = 1 To 5
                        Grid7.Cell(i, m).Text = ""
                    Next
                End If
            Next
        End If
    End Sub
    Public Sub VoicingView_Clear()
        Dim i, j As Integer

        For i = 1 To Grid7.Rows - 1
            For j = 1 To Grid7.Cols - 1
                Grid7.Cell(i, j).BackColor = Color.White
                Grid7.Cell(i, j).Text = ""
            Next
        Next
    End Sub

    'Private Sub Calcul_AutoVoicing()
    ' Dim m, t, ct As Integer
    'Dim mdeb, tdeb, ctdeb As Integer
    'Dim mfin, tfin, ctfin As Integer
    ''
    'Dim a, b As String
    ''
    ''Dim Renv As String
    ''Dim NumRenv As Integer
    ''
    'Dim DerEventH As Integer
    '   '
    '  ' Maj mesure de début de traitement
    ' ' *********************************

    'mdeb = 1 ' pour traitement génral
    'tdeb = 0
    'ctdeb = 0
    ''
    '' Maj mesure de fin de traitement
    '' *********************************
    'mfin = Det_DerEventH2() ' pour traitement général
    'tfin = 5
    'ctfin = 4
    ''
    ''
    ''RAZ_MarqAutoVoice()
    ''
    'Grid2.Refresh()
    'Grid2.AutoRedraw = True
    ''
    ''TContext1.AutoVoiceValid = True
    ''
    ''
    'DerEventH = Det_DerEventH()
    'If DerEventH <= mfin Then
    '        mfin = DerEventH
    'End If
    '   '
    '  TContext1.ColDeb = mdeb
    ' TContext1.Colfin = mfin
    'TContext1.Colcours = mdeb
    '
    'For m = mdeb To mfin
    'For t = tdeb To tfin
    'For ct = ctdeb To ctfin
    'a = TableEventH(m, t, ct).Accord
    'If Trim(a) <> "" Then
    ''a = TableEventH(m, t, ct).Accord
    'a = Det_NotesAccord(Trim(a))
    ''
    ''NumRenv = TableEventH(m, t, ct).RenvChoisi
    ''Renv = TableEventH(m, t, ct).Renversement(NumRenv)
    'b = Maj_NotesCommunes(Trim(a), -1)
    'b = Maj_Large_BasseMoins1(b, "PisteHorsZone", -1)
    'TableNotesAccords(m, t, ct) = Trim(b)
    'End If
    'Next ct
    'Next t
    'Next m
    '
    'End Sub
    Function Maj_Large_BasseMoins1(b As String, typPiste As String, Zone As Integer) As String
        Dim i, n As Integer
        Dim tbl() As String
        Dim flag_ As Boolean = False

        tbl = Split(b, " ")
        Select Case typPiste
            Case "PisteHorsZone" ' n'existe plus
                '
            Case "PisteZone"
                If TZone(Zone).OctaveMoins1 = True Then
                    n = Val(TZone(Zone).VoixAsso_OctaveMoins1) - 1
                    If n <= UBound(tbl) And n <> -1 Then
                        i = ListNotesd.IndexOf(tbl(n)) 'i = Det_NumNote(tbl(n))
                        If i - 12 >= 0 Then
                            tbl(n) = ListNotesd(i - 12) ' on descend d'une octave la note de la voix n  de l'accord
                            flag_ = True
                        End If
                    End If
                End If
        End Select
        '
        If flag_ = True Then
            Maj_Large_BasseMoins1 = String.Join(" ", tbl) ' on recontruit la chaine des notes de l'accord
        Else
            Maj_Large_BasseMoins1 = Trim(b) ' on restitue le voicing
        End If

    End Function

    Function AccordBemDies(accord As String) ' savoir si un accord est exprimé avec des Bémols ou des dièess
        Dim tbl() As String
        Dim i As Integer

        tbl = Split(accord, " ")
        AccordBemDies = "#"
        For i = 0 To UBound(tbl)
            If Mid(tbl(i), 2, 1) = "b" Then
                AccordBemDies = "b"
                Exit For
            End If
        Next i
    End Function
    Function Det_NumVoix_OctaveMoins1(typPiste As String, Zone As Integer) As Integer
        Dim a As String
        Det_NumVoix_OctaveMoins1 = 0
        Select Case typPiste
            Case "PisteHorsZone"
                a = THorsZone.VoixAsso_OctaveMoins1
                Det_NumVoix_OctaveMoins1 = Val(Microsoft.VisualBasic.Right(a, 1)) - 1
            Case "PisteZone"
                a = TZone(Zone).VoixAsso_OctaveMoins1
                Det_NumVoix_OctaveMoins1 = Val(Microsoft.VisualBasic.Right(a, 1)) - 1
        End Select

    End Function
    Function Det_NumVoix_OctavePlus1(typPiste As String, Zone As Integer) As Integer
        Dim a As String
        Det_NumVoix_OctavePlus1 = 0
        Select Case typPiste
            Case "PisteHorsZone"
                a = THorsZone.VoixAsso_OctavePlus1
                Det_NumVoix_OctavePlus1 = Val(Microsoft.VisualBasic.Right(a, 1)) - 1
            Case "PisteZone"
                a = TZone(Zone).VoixAsso_OctavePlus1
                Det_NumVoix_OctavePlus1 = Val(Microsoft.VisualBasic.Right(a, 1)) - 1
        End Select
    End Function


    Sub RAZ_MarqAutoVoice()
        Dim i As Integer

        Grid2.AutoRedraw = False
        '
        For i = 0 To Grid2.Cols - 1
            If Grid2.Cell(0, i).BackColor <> Color.Green Then
                Grid2.Cell(0, i).BackColor = Color.Beige
            End If
        Next i
        '
        Grid2.Refresh()
        Grid2.AutoRedraw = True
    End Sub

    Private Sub T_Maj_AutoV_Tick(sender As Object, e As EventArgs) Handles T_Maj_AutoV.Tick
        Dim i As Integer
        If TContext1.Colcours <> TContext1.Colfin Then
            If Grid2.Cell(0, TContext1.Colcours).BackColor <> Color.Green Then
                'sauvegarde ancienne couleur de la cellule en cours
                TContext1.CouleurPréced = Grid2.Cell(0, TContext1.Colcours - 1).BackColor
                ' écriture nouvelle couleur dans la cellule en cours
                Grid2.Cell(0, TContext1.Colcours).BackColor = Color.Gainsboro
                ' restitution ancienne couleur sur cellule précédente
                If TContext1.Colcours <> TContext1.ColDeb Then
                    Grid2.Cell(0, TContext1.Colcours - 1).BackColor = TContext1.CouleurPréced
                    Grid2.Cell(0, TContext1.Colcours).SetFocus()
                End If
            Else
                TContext1.CouleurPréced = Color.Green
            End If
            ' incrémenter la cellule en cours
            TContext1.Colcours = TContext1.Colcours + 1
            '
        Else
            Grid2.AutoRedraw = False
            For i = TContext1.ColDeb To TContext1.Colfin
                If Grid2.Cell(0, i).BackColor <> Color.Green Then
                    Grid2.Cell(0, i).BackColor = Color.Gainsboro
                End If
            Next i
            Grid2.Refresh()
            Grid2.AutoRedraw = True
            '
            T_Maj_AutoV.Enabled = False
            T_Maj_AutoV.Stop()

        End If
    End Sub
    Private Sub Button11_Click_3(sender As Object, e As EventArgs)
        NouveauProjet()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs)
        Ouvrir()
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs)
        Enregistrer()
    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs)
        'Calcul_AutoVoicingZ()
        'ExportAccord()
    End Sub

    Private Sub Button14_Click_1(sender As Object, e As EventArgs)
        Annuler()
        Calcul_AutoVoicingZ()
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs)

        Calcul_AutoVoicingZ()
    End Sub

    Private Sub Button16_Click_2(sender As Object, e As EventArgs)
        Calcul_AutoVoicingZ()
    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs)

        Calcul_AutoVoicingZ()
    End Sub

    Private Sub Button23_Click_1(sender As Object, e As EventArgs)
        MIDIReset()
    End Sub
    Private Sub Button29_Click(sender As Object, e As EventArgs) Handles Button29.Click
        MIDIReset()
    End Sub
    Function Det_DerEventH() As Integer
        Dim i, j As Integer
        Det_DerEventH = Terme.Value
        j = Terme.Value
        For i = j To 1 Step -1
            If Trim(Grid2.Cell(1, i).Text) <> "" Then
                Det_DerEventH = i
                Exit Function
            End If
        Next i
    End Function
    Function Det_DerEventH2() As Integer
        Dim i, j As Integer
        Det_DerEventH2 = nbMesuresUtiles
        j = nbMesuresUtiles
        For i = j To 1 Step -1
            If Trim(Grid2.Cell(2, i).Text) <> "" Then
                Det_DerEventH2 = i
                Exit Function
            End If
        Next i
    End Function


    Private Sub Transp11_Click(sender As Object, e As EventArgs)
        'Transposer2(Transp11.Text)
        '
        'Calcul_AutoVoicing()
        'Copie_THorsZones_Ds_TZones()
        Calcul_AutoVoicingZ()
    End Sub

    Function TonalitéDsClef(Tonalité) As String
        Dim table1(0 To 12) As String
        Dim tbl() As String
        Dim i As Integer
        Dim flag1 As Boolean = False
        ' cette procédure part du principe que la Tonalité en entrée est toujours fournie en #
        table1(0) = "C#"
        table1(1) = "F#"
        table1(2) = "B"
        table1(3) = "E"
        table1(4) = "A"
        table1(5) = "D"
        table1(6) = "G"
        table1(7) = "C"
        table1(8) = "F"
        table1(10) = "Bb"
        table1(11) = "Eb"
        table1(12) = "Ab"
        '
        tbl = Split(Trim(Tonalité), " ")
        For i = 0 To 12
            If Trim(tbl(0)) = table1(i) Then
                flag1 = True
                Exit For
            End If
        Next
        If flag1 = True Then
            TonalitéDsClef = "#"
            If Trim(tbl(0)) = "F" Then
                TonalitéDsClef = "b"
            End If
        Else
            TonalitéDsClef = "b"
        End If
    End Function
    Function TonalitéDsClef2(Tonalité) As String
        Dim table1(0 To 12) As String
        Dim table2(0 To 12) As String
        Dim tbl() As String
        Dim i As Integer
        Dim flag1 As Boolean = False
        Dim RelativeMaj As String
        ' cette procédure part du principe que la Tonalité en entrée est toujours fournie en #
        tbl = Split(Trim(Tonalité), " ")
        table1(0) = "C#"
        table1(1) = "F#"
        table1(2) = "B"
        table1(3) = "E"
        table1(4) = "A"
        table1(5) = "D"
        table1(6) = "G"
        table1(7) = "C"
        table1(8) = "F"
        table1(10) = "Bb"
        table1(11) = "Eb"
        table1(12) = "Ab"
        If tbl(1) = "MinH" Or tbl(1) = "MinM" Then
            table2(0) = "A#"
            table2(1) = "D#"
            table2(2) = "G#"
            table2(3) = "C#"
            table2(4) = "F#"
            table2(5) = "B"
            table2(6) = "E"
            table2(7) = "A"
            table2(8) = "D"
            table2(10) = "G"
            table2(11) = "C"
            table2(12) = "F"
            '
            For i = 0 To 12
                If Trim(tbl(0)) = table2(i) Then
                    flag1 = True
                    Exit For
                End If
            Next
            RelativeMaj = table1(i) + " " + "Maj"
            tbl = Split(Trim(RelativeMaj), " ")
        End If
        '
        For i = 0 To 12
            If Trim(tbl(0)) = table1(i) Then
                flag1 = True
                Exit For
            End If
        Next i
        If flag1 = True Then
            TonalitéDsClef2 = "#"
            If Trim(tbl(0)) = "F" Then
                TonalitéDsClef2 = "b"
            End If
        Else
            TonalitéDsClef2 = "b"
        End If
    End Function
    Function TradAcc_DieseEnBem(accord As String) As String
        Dim tbl() As String
        Dim a, b As String
        Dim i As Integer
        '
        TradAcc_DieseEnBem = Trim(accord)
        '
        tbl = Split(accord)
        a = LCase(tbl(0))
        '
        For i = 0 To 35
            If a = TabNotesD(i) Then
                a = TabNotesB(i) ' on traduit la tonique de l'accord en #
                Exit For
            End If
        Next
        '
        ' Mettre en majuscules
        ' *******************
        If Len(a) = 1 Then
            a = UCase(a)
        Else
            If Mid(a, 2, 2) = "#" Then
                a = UCase(a)
            Else
                b = Mid(a, 1, 1)
                b = UCase(b)
                a = Trim(b) + "b"
            End If
        End If





        If UBound(tbl) > 0 Then
            TradAcc_DieseEnBem = a + " " + tbl(1)
        Else
            TradAcc_DieseEnBem = a
        End If
    End Function

    Function Det_TonalitédDuPremierAccordDsMesure(m As Integer) As String
        Dim t, ct As Integer
        Dim sortir As Boolean

        sortir = False
        For t = 0 To UBound(TableEventH, 2)
            For ct = 0 To UBound(TableEventH, 3)
                If Trim(TableEventH(m, t, ct).Accord) <> "" Then
                    sortir = True
                    Exit For ' t et ct sont trouvés
                End If
            Next ct
            If sortir Then
                Exit For
            End If
        Next t
        Det_TonalitédDuPremierAccordDsMesure = TableEventH(m, t, ct).Tonalité
    End Function
    Private Sub Accord11_Click(sender As Object, e As EventArgs) Handles Accord11.Click
        Ecr_AccordParMenu(Trim(Accord11.Text))
        Calcul_AutoVoicingZ()
    End Sub
    Private Sub Accord12_Click(sender As Object, e As EventArgs) Handles Accord12.Click
        Ecr_AccordParMenu(Trim(Accord12.Text))
        Calcul_AutoVoicingZ()
    End Sub
    Private Sub Accord13_Click(sender As Object, e As EventArgs) Handles Accord13.Click
        Ecr_AccordParMenu(Trim(Accord13.Text))
        Calcul_AutoVoicingZ()
    End Sub
    Private Sub Accord14_Click(sender As Object, e As EventArgs) Handles Accord14.Click
        Ecr_AccordParMenu(Trim(Accord14.Text))
        Calcul_AutoVoicingZ()
    End Sub
    Private Sub Accord15_Click(sender As Object, e As EventArgs) Handles Accord15.Click
        Ecr_AccordParMenu(Trim(Accord15.Text))
        Calcul_AutoVoicingZ()
    End Sub
    Private Sub Accord16_Click(sender As Object, e As EventArgs) Handles Accord16.Click
        Ecr_AccordParMenu(Trim(Accord16.Text))
        Calcul_AutoVoicingZ()
    End Sub
    Private Sub Accord17_Click(sender As Object, e As EventArgs) Handles Accord17.Click
        Ecr_AccordParMenu(Trim(Accord17.Text))
        Calcul_AutoVoicingZ()
    End Sub
    Private Sub Accord18_Click(sender As Object, e As EventArgs) Handles Accord18.Click
        Ecr_AccordParMenu(Trim(Accord18.Text))
        Calcul_AutoVoicingZ()
    End Sub
    Sub Ecr_AccordParMenu(Acc As String)
        Dim tbl() As String
        Dim Chiffrage As String


        Flag_EcrDragDrop = False
        Select Case DerGridCliquée
            Case GridCours.Grid2
                tbl = Split(TableEventH(SauveMouseColGrid2, 1, 1).Mode, " ")
                Chiffrage = tbl(1)
                Select Case Trim(Chiffrage)
                    Case "Maj"
                        OrigineAccord = Modes.Majeur   ' OrigineAccord est mis à jour pour EcritureAccordDsGrid2
                    Case "MinH"
                        OrigineAccord = Modes.MineurH
                    Case "MinM"
                        OrigineAccord = Modes.MineurM
                End Select
                ' Les 3 paramètres ci-après sont utilisés par la procédure "EcritureAccordDsGrid2"
                ' *******************************************************************************
                Entrée_Tonalité = TableEventH(SauveMouseColGrid2, 1, 1).Tonalité
                Entrée_Gamme = TableEventH(SauveMouseColGrid2, 1, 1).Gamme
                Entrée_Mode = TableEventH(SauveMouseColGrid2, 1, 1).Mode
                CAD_LabelCours = TableEventH(SauveMouseColGrid2, 1, 1).Degré

                EcritureAccordDsGrid2(Trim(Acc), SauveMouseColGrid2)
            '

            Case GridCours.Grid3
        End Select
    End Sub
    Private Sub Accord11_1_Click(sender As Object, e As EventArgs) Handles Accord11_1.Click
        Dim com As ToolStripMenuItem = sender
        Dim ind As Integer

        ind = Val(com.Tag)
        ECR_CelluleTabtons(ind, Trim(Accord11_1.Text))
        '
    End Sub
    Private Sub Accord12_1_Click(sender As Object, e As EventArgs) Handles Accord12_1.Click
        Dim com As ToolStripMenuItem = sender
        Dim ind As Integer

        ind = Val(com.Tag)
        '
        ECR_CelluleTabtons(ind, Trim(Accord12_1.Text))

    End Sub
    Private Sub Accord13_1_Click(sender As Object, e As EventArgs) Handles Accord13_1.Click
        Dim com As ToolStripMenuItem = sender
        Dim ind As Integer

        ind = Val(com.Tag)
        '
        ECR_CelluleTabtons(ind, Trim(Accord13_1.Text))
    End Sub
    Private Sub Accord14_1_Click(sender As Object, e As EventArgs) Handles Accord14_1.Click
        Dim com As ToolStripMenuItem = sender
        Dim ind As Integer

        ind = Val(com.Tag)
        '
        ECR_CelluleTabtons(ind, Trim(Accord14_1.Text))
    End Sub
    Private Sub Accord15_1_Click(sender As Object, e As EventArgs) Handles Accord15_1.Click
        Dim com As ToolStripMenuItem = sender
        Dim ind As Integer

        ind = Val(com.Tag)
        '
        ECR_CelluleTabtons(ind, Trim(Accord15_1.Text))

    End Sub
    Private Sub Accord16_1_Click(sender As Object, e As EventArgs) Handles Accord16_1.Click
        Dim com As ToolStripMenuItem = sender
        Dim ind As Integer

        ind = Val(com.Tag)
        '
        ECR_CelluleTabtons(ind, Trim(Accord16_1.Text))

    End Sub
    Private Sub Accord17_1_Click(sender As Object, e As EventArgs) Handles Accord17_1.Click
        Dim com As ToolStripMenuItem = sender
        Dim ind As Integer

        ind = Val(com.Tag)
        '
        ECR_CelluleTabtons(ind, Trim(Accord17_1.Text))
    End Sub
    Private Sub Accord18_1_Click(sender As Object, e As EventArgs) Handles Accord18_1.Click
        Dim com As ToolStripMenuItem = sender
        Dim ind As Integer

        ind = Val(com.Tag)
        '
        ECR_CelluleTabtons(ind, Trim(Accord18_1.Text))
    End Sub

    Sub ECR_CelluleTabtons(N_Cellule As Integer, Accord As String)
        Dim i, j, k As Integer
        Dim ligne, col As Integer
        Dim a As String
        Dim IndexDegré As Integer

        i = N_Cellule
        ' Ecriture du nouvel accord dans Tabtons
        ' **************************************
        TabTons.Item(i).Text = Trim(Accord)
        TabTons.Item(i).Refresh()
        '
        ' Détermination de l'index du degré dans chaque ligne (equivalence de N_Cellule en N° degré)
        ' ******************************************************************************************
        ligne = Det_LigneTableGlobale(i)
        Select Case ligne
            Case 0
                col = i
            Case 1
                col = i - 7
            Case 2
                col = i - 14
        End Select
        a = TabTonsDegrés.Item(col).Text
        IndexDegré = Det_IndexDegré(a)
        '
        ' Mise à jour de l'accord dans TableCoursAcc (avec ligne et Indexdegré)
        ' *********************************************************************
        TableCoursAcc(ligne, IndexDegré).Accord = Trim(Accord)
        '
        ' Mise à jour du type d'accord dans TableCoursAcc 
        ' ***********************************************
        j = InStr(Accord, "7")
        If j <> 0 Then
            TableCoursAcc(ligne, IndexDegré).TyAcc = Menu4Notes.Text ' Accord de 7e
        Else
            j = InStr(Accord, "9")
            If j <> 0 Then
                TableCoursAcc(ligne, IndexDegré).TyAcc = MenuNotes9.Text ' Accord de 9e
            Else
                j = InStr(Accord, "4")
                k = InStr(Accord, "11")
                If j <> 0 Or k <> 0 Then
                    TableCoursAcc(ligne, IndexDegré).TyAcc = MenuNotes11.Text ' Accord de 11e ou 4
                Else
                    TableCoursAcc(ligne, IndexDegré).TyAcc = Menu3notes.Text  ' Accord de 3 notes
                End If
            End If
        End If
        '
        'Maj_Renversement(i)
        'RAZ_FiltresTabtons(i)
    End Sub
    Sub ECR_CelluleTabtonsVoisins(N_Cellule As Integer, Accord As String)
        Dim i, j, k As Integer
        Dim ligne As Integer
        Dim IndexDegré As Integer

        i = N_Cellule
        ' Ecriture du nouvel accord dans Tabtons
        ' **************************************
        TabTonsVoisins.Item(i).Text = Trim(Accord)
        TabTonsVoisins.Item(i).Refresh()
        '
        ' Détermination de l'index du degré dans chaque ligne (equivalence de N_Cellule en N° degré)
        ' ******************************************************************************************
        IndexDegré = Det_IndexDegré2(i)
        ligne = Det_LigneTableGlobale(i)
        '
        ' Mise à jour de l'accord dans TableCoursAcc (avec ligne et Indexdegré)
        ' *********************************************************************
        TableCoursAccVoisins(ligne, IndexDegré).Accord = Trim(Accord)
        '
        ' Mise à jour du type d'accord dans TableCoursAccVoisins 
        ' ******************************************************
        j = InStr(Accord, "7")
        If j <> 0 Then
            TableCoursAccVoisins(ligne, IndexDegré).TyAcc = Menu4Notes.Text ' Accord de 7e
        Else
            j = InStr(Accord, "9")
            If j <> 0 Then
                TableCoursAccVoisins(ligne, IndexDegré).TyAcc = MenuNotes9.Text ' Accord de 9e
            Else
                j = InStr(Accord, "4")
                k = InStr(Accord, "11")
                If j <> 0 Or k <> 0 Then
                    TableCoursAccVoisins(ligne, IndexDegré).TyAcc = MenuNotes11.Text ' Accord de 11e ou 4
                Else
                    TableCoursAccVoisins(ligne, IndexDegré).TyAcc = Menu3notes.Text  ' Accord de 3 notes
                End If
            End If
        End If
        '
    End Sub
    Private Sub Accord11_2_Click(sender As Object, e As EventArgs) Handles Accord11_2.Click
        Dim com As ToolStripMenuItem = sender
        Dim ind As Integer

        ind = Val(com.Tag)
        '
        ECR_CelluleTabtonsCAD(ind, Trim(Accord11_2.Text))
    End Sub
    Private Sub Accord12_2_Click(sender As Object, e As EventArgs) Handles Accord12_2.Click
        Dim com As ToolStripMenuItem = sender
        Dim ind As Integer

        ind = Val(com.Tag)
        '
        ECR_CelluleTabtonsCAD(ind, Trim(Accord12_2.Text))
    End Sub
    Private Sub Accord13_2_Click(sender As Object, e As EventArgs) Handles Accord13_2.Click
        Dim com As ToolStripMenuItem = sender
        Dim ind As Integer

        ind = Val(com.Tag)
        '
        ECR_CelluleTabtonsCAD(ind, Trim(Accord13_2.Text))
    End Sub
    Private Sub Accord14_2_Click(sender As Object, e As EventArgs) Handles Accord14_2.Click
        Dim com As ToolStripMenuItem = sender
        Dim ind As Integer

        ind = Val(com.Tag)
        '
        ECR_CelluleTabtonsCAD(ind, Trim(Accord14_2.Text))
    End Sub
    Private Sub Accord15_2_Click(sender As Object, e As EventArgs) Handles Accord15_2.Click
        Dim com As ToolStripMenuItem = sender
        Dim ind As Integer

        ind = Val(com.Tag)
        '
        ECR_CelluleTabtonsCAD(ind, Trim(Accord15_2.Text))
    End Sub

    Private Sub Accord16_2_Click(sender As Object, e As EventArgs) Handles Accord16_2.Click
        Dim com As ToolStripMenuItem = sender
        Dim ind As Integer

        ind = Val(com.Tag)
        '
        ECR_CelluleTabtonsCAD(ind, Trim(Accord16_2.Text))
    End Sub

    Private Sub Accord17_2_Click(sender As Object, e As EventArgs) Handles Accord17_2.Click
        Dim com As ToolStripMenuItem = sender
        Dim ind As Integer

        ind = Val(com.Tag)
        '
        ECR_CelluleTabtonsCAD(ind, Trim(Accord17_2.Text))
    End Sub

    Private Sub Accord18_2_Click(sender As Object, e As EventArgs) Handles Accord18_2.Click
        Dim com As ToolStripMenuItem = sender
        Dim ind As Integer

        ind = Val(com.Tag)
        '
        ECR_CelluleTabtonsCAD(ind, Trim(Accord18_2.Text))
    End Sub
    Sub ECR_CelluleTabtonsCAD(N_Cellule As Integer, Accord As String)
        Dim i As Integer
        Dim a As String
        Dim typAccord, IndexDegré As Integer
        '
        For i = 0 To 3 ' on recopie les données de TableGlobalAcc dans CAD_TableGlobalAcc
            For k = 0 To 6
                CAD_TableGlobalAcc(i, 0, k) = TableGlobalAcc(i, 0, k) ' mise à jour des accords majeurs
                CAD_TableGlobalAcc(i, 1, k) = TableGlobalAcc(i, 1, k) ' mise à jour des accords mineures
            Next k
        Next i
        '
        i = N_Cellule
        typAccord = ComboBox6.SelectedIndex
        '
        a = TabCadDegrés.Item(i).Text
        If Trim(a) <> "" Then
            IndexDegré = Det_IndexDegré(a)
            '
            CAD_TableCoursAcc(IndexDegré).Accord = TradAcc_LatAngl(Trim(Accord)) 'TableGlobalAcc(typAccord, 0, indexDegré)
            TabCad.Item(i).Text = Trim(Accord) ' CAD_TableCoursAcc(i).Accord
        End If
    End Sub

    Private Sub TabPage_Tonalité_Paint(sender As Object, e As PaintEventArgs)
        OngletCours = 1
    End Sub

    Private Sub TabPage_Cadences_Paint(sender As Object, e As PaintEventArgs)
        OngletCours = 0
    End Sub
    Private Sub OctaveRacine_KeyDown(sender As Object, e As KeyEventArgs)
        e.SuppressKeyPress = True
    End Sub

    Private Sub NoteRacine_KeyDown(sender As Object, e As KeyEventArgs)
        e.SuppressKeyPress = True
    End Sub

    Private Sub ComboBox22_KeyDown(sender As Object, e As KeyEventArgs)
        e.SuppressKeyPress = True
    End Sub



    Private Sub ComboMidiOut_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboMidiOut.KeyDown
        e.SuppressKeyPress = True
    End Sub

    Private Sub ComboMidiIn_KeyDown(sender As Object, e As KeyEventArgs)
        e.SuppressKeyPress = True
    End Sub
    Private Sub NoteRacine_SelectedIndexChanged_2(sender As Object, e As EventArgs)
        'If NotesCommunes.Checked = True Then
        'If VoicingZone.Checked = False Then
        'THorsZone.NoteRacine = NoteRacine.SelectedIndex
        'Calcul_AutoVoicing()
        'Else
        'TZone(ZoneCourante).NoteRacine = NoteRacine.SelectedIndex
        Calcul_AutoVoicingZ()
        'End If

        'End If
    End Sub
    Private Sub NoteRacine_SelectedIndexChanged_3(sender As Object, e As EventArgs)
        'If NotesCommunes.Checked = True Then
        'If VoicingZone.Checked = False Then
        'THorsZone.NoteRacine = NoteRacine.SelectedIndex
        'THorsZone.Racine = Trim(NoteRacine.Text)
        ''Calcul_AutoVoicing()
        'Else

        'TZone(ZoneCourante).NoteRacine = NoteRacine.SelectedIndex
        'TZone(ZoneCourante).Racine = Trim(NoteRacine.Text)
        'Calcul_AutoVoicingZ()
        'End If
        'End If

    End Sub
    Private Sub Octave_Plus1_CheckedChanged(sender As Object, e As EventArgs)
        'If NotesCommunes.Checked = True Then
        'If VoicingZone.Checked = False Then
        'THorsZone.OctavePlus1 = Octave_Plus1.Checked
        'Calcul_AutoVoicing()
        'Else
        'TZone(ZoneCourante).OctavePlus1 = Octave_Plus1.Checked
        Calcul_AutoVoicingZ()
        'End If
        'End If
    End Sub
    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs)

    End Sub
    Private Sub Button30_Click(sender As Object, e As EventArgs)
        Dim a As String

        a = Det_NotesAccord("C M7")
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs)
        RAZ_CouleursPiano()
    End Sub
    Public Sub Maj_TableGlobalAccSubsti(tona As String, tonaMin As String, tonaMinN As String)
        Dim tbl() As String
        Dim i As Integer
        Dim j As Integer
        ' TableGlobalAccSubsti(i,j,k)
        ' ***************************
        ' i=type d'accord 0 = 3note, 1=4notes avec 7e,2=4 notes avec 9e,3=4 notes avec 11e
        ' j=type de gamme 0=Majeure, 1=Minh et 2 = MInM
        ' k = liste de 7 accords pour chaque couple (i,j)²
        '
        ' Maj
        For i = 3 To 6
            tbl = Split(Mode2(Trim(tona), "Maj", i, False), "-")
            For j = 0 To 6
                TableGlobalAccSubsti(i - 3, 0, j) = tbl(j)
            Next j
        Next i
        ' MinH
        For i = 3 To 6
            tbl = Split(Mode2(Trim(tonaMin), "MinH", i, False), "-")
            For j = 0 To 6
                TableGlobalAccSubsti(i - 3, 1, j) = tbl(j)
            Next j
        Next i
        ' MinM
        For i = 3 To 6
            tbl = Split(Mode2(Trim(tonaMin), "MinM", i, False), "-")
            For j = 0 To 6
                TableGlobalAccSubsti(i - 3, 2, j) = tbl(j)
            Next j
        Next i
        ' MinN
        For i = 3 To 4
            tbl = Split(Mode2(Trim(tonaMinN), "MinN", i, False), "-") ' MIN = Mineur naturel 
            For j = 0 To 6
                TableGlobalAccSubsti(i - 3, 3, j) = tbl(j)
            Next j
        Next i
    End Sub

    Private Sub Maj_Substitutions(m, t, ct)

        Dim accord As String = Trim(TableEventH(m, t, ct).Accord)
        Dim degré As String = Det_DegréRomain2(TableEventH(m, t, ct).Degré)
        Dim gamme As String = TableEventH(m, t, ct).Gamme
        Dim mode As String = TableEventH(m, t, ct).Mode
        Dim tonalité As String = TableEventH(m, t, ct).Tonalité
        Dim position As String = TableEventH(m, t, ct).Position
        Dim TypAcc As Integer
        Dim a As String
        Dim deg As Integer = TableEventH(m, t, ct).Degré
        Dim tbl() As String
        Dim tbl1() As String
        Dim tbl2() As String
        Dim accords(0 To 3) As String
        Dim i As Integer


        ' effacer l'accord actuel
        ' **********************
        For i = 0 To 3
            LabSubsti.Item(i).Text = "---"
        Next
        Label73.Text = ""
        Label67.Text = accord
        '
        ' Vérifier que l'accord considéré est un accord de 3 notets oude 4 note avec 7e (sinon, la substitution n'est pas offerte)
        ' *************************************************************************************************************************

        If Verif_PresAcc(accord) Then

            a = "- " + "Position : " + position + Chr(13) _
            + "- " + "Tonalité : " + tonalité + " Mode : " + mode + Chr(13) _
            + "- " + "Accord  : " + accord + " Degré : " + degré + Chr(13) _
            + "- " + "Gamme : " + gamme

            tbl = Det_NotesAccord(Trim(accord)).Split("-")
            TypAcc = tbl.Count ' nombre de notes dans l'accord (3 ou 4)

            ' Calcul Substitution par tonalité mineure
            ' ****************************************
            ' Principe : pour un accord de C Maj on va chercher le même dégré dans la tonalité de C Mineur naturel : par exemple E m dans C Maj peut être substituer par Eb dans C mineur naturel
            tbl = mode.Split()
            If tbl(1) = "Maj" Then
                tbl = Trim(tonalité).Split()
                a = Trim(tbl(0)) + " MinH"   ' Si La tonalité est C Maj, mais on cherche la relative majeur de C Min naturel  pour la subtitution par tonalité mineure
                EventhSubsti(0).Tonalité = Det_RelativeMajeure(a)
                EventhSubsti(0).Mode = Trim(EventhSubsti(0).Tonalité)          ' comme les accords de C MinNaturel sont les mêmes que de Eb --> Mode Eb Maj pour simplifier
                '
                ' Changement du signe (si nécessaire) de la note en fonction de la nouvelle tonalité

                LabSubsti.Item(0).Text = Retab_Note_Tona(EventhSubsti(0).Tonalité, TableGlobalAccSubsti(TypAcc - 3, 3, deg))
                '
                ' Ici --> deg du nouvel accord = deg d'ancien accord
                EventhSubsti(0).Gamme = Trim(gamme)
                EventhSubsti(0).Accord = Trim(LabSubsti.Item(0).Text)
                '
                EventhSubsti(0).Degré = TradDegré_MinNMaj(deg)
            Else
                If LangueIHM = "fr" Then
                    Label73.Text = "L'accord à substituer doit appartenir à un mode Majeur."
                Else
                    Label73.Text = "The chord to be substituted must belong to a Major mode."
                End If
                LabSubsti.Item(0).Text = "N/A"
            End If
            ' Calcul Substitution diatonique
            ' ******************************
            ' 
            tbl1 = (Ind_SubstiDiat(TableEventH(m, t, ct).Degré)).Split() ' degré des accord de substitution
            tbl2 = Det_NotesAccord(Trim(accord)).Split("-")
            'Dim NbNotes As Integer = tbl2.Count ' nombre de notes dans l'accord
            Dim IndMode As Integer = Ind_Modes(Trim(TableEventH(m, t, ct).Mode))
            LabSubsti.Item(1).Text = TableGlobalAccSubsti(TypAcc - 3, IndMode, tbl1(0))
            LabSubsti.Item(2).Text = TableGlobalAccSubsti(TypAcc - 3, IndMode, tbl1(1))
            '
            EventhSubsti(1).Tonalité = Trim(tonalité)
            EventhSubsti(1).Mode = Trim(mode)
            EventhSubsti(1).Degré = tbl1(0) 'deg '
            EventhSubsti(1).Gamme = Trim(gamme)
            EventhSubsti(1).Accord = Trim(LabSubsti.Item(1).Text)
            '
            EventhSubsti(2).Tonalité = Trim(tonalité)
            EventhSubsti(2).Mode = Trim(mode)
            EventhSubsti(2).Degré = tbl1(1) 'deg '
            EventhSubsti(2).Gamme = Trim(gamme)
            EventhSubsti(2).Accord = Trim(LabSubsti.Item(2).Text)           '

            ' Calcul Substitution par Ve degré de tonique
            ' *******************************************
            ' Ici la tonique de l'accord sert de tonique au mode dont on va cherche le Ve degré, Par exemple accord est C, on va chercher le Ve degré de la tonatilé de C = G
            tbl = accord.Split()
            a = Retab_Mode(Trim(tbl(0)), "Maj") ' construction de la tonalité : la tonalité se construit ici à partir de la tonique de l'accord + "Maj"(tr'
            EventhSubsti(3).Tonalité = Trim(Trim(a) + " Maj")
            EventhSubsti(3).Mode = Trim(Trim(a) + " Maj")
            EventhSubsti(3).Degré = 4 ' on va chercher le 5e degré de la tonalité construite
            EventhSubsti(3).Gamme = Trim(gamme)
            '
            a = NoteInterval(Trim(a), "5")
            a = UCaseBémol(a)
            If TypAcc = 4 Then
                a = a + " 7"
            End If
            LabSubsti.Item(3).Text = Trim(a)
            EventhSubsti(3).Accord = Trim(a)
        Else
            If LangueIHM = "fr" Then
                Label73.Text = "Seuls les accords de 3 notes et 4 notes avec 7e sont pris en compte pour une substitution."
            Else
                Label73.Text = "Only 3-notes and 4-notes chords with 7ths are considered for substitution."
            End If
            For i = 0 To 3
                LabSubsti.Item(i).Text = "N/A"
            Next
        End If
    End Sub
    Function Ind_Modes(Mode As String) As Integer
        Dim tbl() As String = Mode.Split()

        Select Case tbl(1)
            Case "Maj"
                Return 0
            Case "MinH"
                Return 1
            Case "MinM"
                Return 2
            Case Else
                Return 0
        End Select

    End Function
    Function Retab_Mode(NoteTonique As String, ChiffMode As String) As String
        Dim a As String = NoteTonique
        Select Case ChiffMode
            Case "Maj"
                Select Case NoteTonique
                    Case "Db"
                        a = "C#"
                    Case "Gb"
                        a = "F#"
                    Case "A#"
                        a = "Bb"
                    Case "D#"
                        a = "Eb"
                    Case "G#"
                        a = "Ab"

                End Select
            Case "MinH", "MinM", "MinN"
                Select Case NoteTonique
                    Case "Gb"
                        a = "F#"
                    Case "Ab"
                        a = "G#"
                    Case "Eb"
                        a = "D#"
                    Case "Bb"
                        a = "A#"
                End Select
        End Select
        Return a
    End Function
    Function Ind_SubstiDiat(degré As Integer) As String
        Select Case degré
            Case 0 ' I
                Return "5 2" ' C --> VI et III
            Case 1 ' II
                Return "6 3" ' D m --> VII et IV
            Case 2 ' III
                Return "0 4" ' E m --> I et V
            Case 3 ' IV
                Return "1 5" ' F --> II et VI
            Case 4 ' V
                Return "2 6" ' G --> II et VII
            Case 5 ' VI
                Return "3 0" ' A m --> IV et I
            Case 6 ' VII
                Return "4 1" ' Bmb5 --> V et II
            Case Else
                Return "5 2" ' C --> VI et III
        End Select
    End Function
    Private Function Verif_PresAcc(Acc As String) As Boolean
        Dim boo As Boolean = False
        ' TableGlobalAccSubsti(i,j,k)
        ' ***************************
        ' i=type d'accord 0 = 3note, 1=4notes avec 7e,2=4 notes avec 9e,3=4 notes avec 11e
        ' j=type de gamme à=Majeure, 1=Minh et 2 = MInM
        ' k = liste de 7 accords pour chaque couple (i,j)

        ' Maj
        ' ***
        For i = 0 To 1
            For j = 0 To 6
                If Trim(Acc) = Trim(TableGlobalAccSubsti(i, 0, j)) Then boo = True
            Next j
        Next i
        '
        ' MinH
        ' ****
        If Not boo Then
            For i = 0 To 1

                For j = 0 To 6
                    If Trim(Acc) = Trim(TableGlobalAccSubsti(i, 1, j)) Then boo = True
                Next j
            Next i
        End If
        '
        ' MinM
        ' ****
        If Not boo Then
            For i = 0 To 1
                For j = 0 To 6
                    If Trim(Acc) = Trim(TableGlobalAccSubsti(i, 2, j)) Then boo = True
                Next j
            Next i
        End If
        ' MinN
        ' ****
        If Not boo Then
            For i = 0 To 1
                For j = 0 To 6
                    If Trim(Acc) = Trim(TableGlobalAccSubsti(i, 3, j)) Then boo = True
                Next j
            Next i
        End If
        Return boo
    End Function
    Function Retab_Note_Tona(Tona As String, Acc As String) As String
        Dim a As String
        Dim tbl1() As String
        Dim tbl2() As String

        tbl1 = Tona.Split()
        Dim clf As String = Det_Clef(Trim(tbl1(0)))

        tbl2 = Acc.Split()

        If clf = "#" Then
            tbl2(0) = Trad_BemDies_Maj(tbl2(0))
        Else
            tbl2(0) = Trad_DiesBem_Maj(tbl2(0))
        End If
        a = Join(tbl2, " ")
        Return a
    End Function
    Function TradDegré_MinNMaj(degMinN As Integer) ' exemple traduction d'un degré MIN en dgré de sa relative majeur (C m --> Eb)
        Dim degMaj As Integer = 0
        Select Case degMinN
            Case 0
                degMaj = 5
            Case 1
                degMaj = 6
            Case 2
                degMaj = 0
            Case 3
                degMaj = 1
            Case 4
                degMaj = 2
            Case 5
                degMaj = 3
            Case 6
                degMaj = 4
        End Select
        Return degMaj
    End Function
    Function Det_RelativeMineure(Mode As String) As String
        Dim tbl() As String
        Dim note As String
        '
        If Langue = "fr" Then ' Langue des notes
            Mode = TradAcc_LatAngl(Mode)
        End If
        '
        Det_RelativeMineure = Trim(Mode)

        tbl = Split(Mode)
        If Trim(tbl(1)) = "Maj" Then
            note = tbl(0)
            Select Case Trim(note)
                Case "C#"
                    Det_RelativeMineure = "A# " + "Min"
                Case "F#"
                    Det_RelativeMineure = "D# " + "Min"
                Case "B"
                    Det_RelativeMineure = "G# " + "Min"
                Case "E"
                    Det_RelativeMineure = "C# " + "Min"
                Case "A"
                    Det_RelativeMineure = "F# " + "Min"
                Case "D"
                    Det_RelativeMineure = "B " + "Min"
                Case "G"
                    Det_RelativeMineure = "E " + "Min"
                Case "C"
                    Det_RelativeMineure = "A " + "Min"
                Case "F"
                    Det_RelativeMineure = "D " + "Min"
                Case "Bb"
                    Det_RelativeMineure = "G " + "Min"
                Case "Eb"
                    Det_RelativeMineure = "C " + "Min"
                Case "Ab"
                    Det_RelativeMineure = "F " + "Min"
            End Select

        End If
    End Function
    Function Det_RelativeMajeure(Mode As String) As String
        Dim tbl() As String
        Dim note As String
        '
        If Langue = "fr" Then
            Mode = TradAcc_LatAngl(Mode)
        End If
        '
        Det_RelativeMajeure = Trim(Mode)

        tbl = Split(Mode)
        If Trim(tbl(1)) = "MinH" Or Trim(tbl(1)) = "MinM" Then
            note = tbl(0)
            Select Case Trim(note)
                Case "A#"
                    Det_RelativeMajeure = "C# " + "Maj"
                Case "D#"
                    Det_RelativeMajeure = "F# " + "Maj"
                Case "G#"
                    Det_RelativeMajeure = "B " + "Maj"
                Case "C#"
                    Det_RelativeMajeure = "E " + "Maj"
                Case "F#"
                    Det_RelativeMajeure = "A " + "Maj"
                Case "B"
                    Det_RelativeMajeure = "D " + "Maj"
                Case "E"
                    Det_RelativeMajeure = "G " + "Maj"
                Case "A"
                    Det_RelativeMajeure = "C " + "Maj"
                Case "D"
                    Det_RelativeMajeure = "F " + "Maj"
                Case "G"
                    Det_RelativeMajeure = "Bb " + "Maj"
                Case "C"
                    Det_RelativeMajeure = "Eb " + "Maj"
                Case "F"
                    Det_RelativeMajeure = "Ab " + "Maj"
            End Select

        End If


    End Function
    Function Det_RelatMaj(acc As String) As String
        Dim tbl() As String
        Dim note As String
        Dim a As String

        Det_RelatMaj = "C"
        If Langue = "fr" Then
            acc = TradAcc_LatAngl(acc)
        End If
        tbl = Split(acc)
        note = tbl(0)
        Select Case Trim(note)
            Case "A#", "Bb"
                a = "C#"
            Case "D#", "Eb"
                a = "F#"
            Case "G#", "Ab"
                a = "B"
            Case "C#", "Db"
                a = "E"
            Case "F#", "Gb"
                a = "A"
            Case "B"
                a = "D"
            Case "E"
                a = "G"
            Case "A"
                a = "C"
            Case "D"
                a = "F"
            Case "G"
                a = "Bb"
            Case "C"
                a = "Eb"
            Case "F"
                a = "Ab"
            Case Else
                a = "C"
        End Select
        '
        If Langue = "fr" Then
            Det_RelatMaj = TradNote_AnglLatMaj(a)
        End If
    End Function
    Function Det_RelativeMajeure2(Mode As String) As String
        Dim tbl() As String
        Dim note As String
        '
        '
        Det_RelativeMajeure2 = Trim(Mode)

        tbl = Split(Mode)
        If Trim(tbl(1)) = "MinH" Or Trim(tbl(1)) = "MinM" Then
            note = tbl(0)
            Select Case Trim(note)
                Case "A#"
                    Det_RelativeMajeure2 = "C# " + "Maj"
                Case "D#"
                    Det_RelativeMajeure2 = "F# " + "Maj"
                Case "G#"
                    Det_RelativeMajeure2 = "B " + "Maj"
                Case "C#"
                    Det_RelativeMajeure2 = "E " + "Maj"
                Case "F#"
                    Det_RelativeMajeure2 = "A " + "Maj"
                Case "B"
                    Det_RelativeMajeure2 = "D " + "Maj"
                Case "E"
                    Det_RelativeMajeure2 = "G " + "Maj"
                Case "A"
                    Det_RelativeMajeure2 = "C " + "Maj"
                Case "D"
                    Det_RelativeMajeure2 = "F " + "Maj"
                Case "G"
                    Det_RelativeMajeure2 = "Bb " + "Maj"
                Case "C"
                    Det_RelativeMajeure2 = "Eb " + "Maj"
                Case "F"
                    Det_RelativeMajeure2 = "Ab " + "Maj"
            End Select
        End If
    End Function
    Public Function Det_DegréRomain2(ind As Integer) ' Ind va de 0 à 6
        Dim a As String = ""
        Select Case ind
            Case 0
                a = "I"
            Case 1
                a = "II"
            Case 2
                a = "III"
            Case 3
                a = "IV"
            Case 4
                a = "V"
            Case 5
                a = "VI"
            Case 6
                a = "VII"
        End Select
        Return a
    End Function
    Public Function Det_DegréDécimal(ligne As Integer, acc As String) ' Ind va de 0 à 6
        Dim j As Integer = 1
        Dim a As String = ""
        Dim b As String
        For j = 1 To Grid5.Cols - 1
            b = Trim(Grid5.Cell(ligne, j).Text)
            If Trim(acc) = Trim(Grid5.Cell(ligne, j).Text) Then
                Exit For
            End If
        Next
        Return j - 1
    End Function

    Sub Maj_LGamMaj()
        LGamMaj.Add("C Maj")
        LGamMaj.Add("C# Maj")
        LGamMaj.Add("D Maj")
        LGamMaj.Add("D# Maj")
        LGamMaj.Add("E Maj")
        LGamMaj.Add("F Maj")
        LGamMaj.Add("F# Maj")
        LGamMaj.Add("G Maj")
        LGamMaj.Add("G# Maj")
        LGamMaj.Add("A Maj")
        LGamMaj.Add("A# Maj")
        LGamMaj.Add("B Maj")
    End Sub
    Sub Maj_LGamMinH()
        LGamMinH.Add("C MinH")
        LGamMinH.Add("C# MinH")
        LGamMinH.Add("D MinH")
        LGamMinH.Add("D# MinH")
        LGamMinH.Add("E MinH")
        LGamMinH.Add("F MinH")
        LGamMinH.Add("F# MinH")
        LGamMinH.Add("G MinH")
        LGamMinH.Add("G# MinH")
        LGamMinH.Add("A MinH")
        LGamMinH.Add("A# MinH")
        LGamMinH.Add("B MinH")
    End Sub
    Sub Maj_LGamMinM()
        LGamMinM.Add("C MinM")
        LGamMinM.Add("C# MinM")
        LGamMinM.Add("D MinM")
        LGamMinM.Add("D# MinM")
        LGamMinM.Add("E MinM")
        LGamMinM.Add("F MinM")
        LGamMinM.Add("F# MinM")
        LGamMinM.Add("G MinM")
        LGamMinM.Add("G# MinM")
        LGamMinM.Add("A MinM")
        LGamMinM.Add("A# MinM")
        LGamMinM.Add("B MinM")
    End Sub
    Sub Maj_LGamMajH()
        LGamMajH.Add("C MajH")
        LGamMajH.Add("C# MajH")
        LGamMajH.Add("D MajH")
        LGamMajH.Add("D# MajH")
        LGamMajH.Add("E MajH")
        LGamMajH.Add("F MajH")
        LGamMajH.Add("F# MajH")
        LGamMajH.Add("G MajH")
        LGamMajH.Add("G# MajH")
        LGamMajH.Add("A MajH")
        LGamMajH.Add("A# MajH")
        LGamMajH.Add("B MajH")
    End Sub
    Sub Maj_LGamPentaMin()
        LGamPentaMin.Add("C PMin")
        LGamPentaMin.Add("C# PMin")
        LGamPentaMin.Add("D PMin")
        LGamPentaMin.Add("D# PMin")
        LGamPentaMin.Add("E PMin")
        LGamPentaMin.Add("F PMin")
        LGamPentaMin.Add("F# PMin")
        LGamPentaMin.Add("G PMin")
        LGamPentaMin.Add("G# PMin")
        LGamPentaMin.Add("A PMin")
        LGamPentaMin.Add("A# PMin")
        LGamPentaMin.Add("B PMin")
    End Sub
    Sub Maj_LGamBlues()
        LGamBlues.Add("C Blues")
        LGamBlues.Add("C# Blues")
        LGamBlues.Add("D Blues")
        LGamBlues.Add("D# Blues")
        LGamBlues.Add("E Blues")
        LGamBlues.Add("F Blues")
        LGamBlues.Add("F# Blues")
        LGamBlues.Add("G Blues")
        LGamBlues.Add("G# Blues")
        LGamBlues.Add("A Blues")
        LGamBlues.Add("A# Blues")
        LGamBlues.Add("B Blues")
    End Sub
    Sub Maj_LGam()
        Maj_LGamMaj()
        Maj_LGamMinH()
        Maj_LGamMinM()
        Maj_LGamMajH()
        Maj_LGamPentaMin()
        Maj_LGamBlues()
    End Sub
    Public Function AccordsSélection() As String
        Dim m As Integer
        Dim n As Integer
        Dim a As String = ""
        '

        AccordsSélection = ""
        n = -1
        Do
            ' 
            If Langue = "fr" Then
                a = TradAcc_LatAngl(a)
            End If
            a = TradAcc_BemEnDiese(a)
            If Trim(a) <> "" Then
                AccordsSélection = AccordsSélection + Trim(a) + ";"
            End If
            m = m - 1
            n = n + 1
        Loop Until n >= SélectionLignes Or m < 1
        '
        If Trim(AccordsSélection) <> "" Then
            AccordsSélection = Mid(AccordsSélection, 1, Len(AccordsSélection) - 1)
        End If
        '
    End Function


    Private Sub AideToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AideToolStripMenuItem.Click
        Dim a As String = "https://compomusic.fr/guide-rapide/"
        '
        Process.Start(Trim(a))
    End Sub

    Private Sub AuSujetDeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AuSujetDeToolStripMenuItem.Click
        AboutBox1.ShowDialog()
    End Sub

    Private Sub FrançaisToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        Process.Start("www.calquesmidi.fr/HV_Manuel/HV_ManuelFR.htm")
    End Sub

    Private Sub EnglishToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        Dim a As String
        Dim titre As String
        a = Langue
        Langue = "en"
        Avertis = "Coming soon ..."
        titre = "Warning"
        ' MessageHV.ShowDialog()
        ' Dim result As DialogResult = MessageBox.Show(Avertis, titre, MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Langue = Trim(a)

    End Sub
    Private Sub ToolStripMenuItem6_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem6.Click
        'If LangueIHM = "fr" Then
        'Process.Start("http://calquesmidi.fr/au-sujet-du-site")
        'Else
        'Process.Start("http://midilayers.org/about-this-web-site")
        'End If
        Splash.Show()
    End Sub

    Private Sub Cartozone_Paint(sender As Object, e As PaintEventArgs)
        OngletCours = 2
    End Sub

    Private Sub SiteWebToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SiteWebToolStripMenuItem.Click
        'If Module1.LangueIHM = "fr" Then
        ' Process.Start("https://www.hyperarp.fr")
        'Else
        'Process.Start("https://www.hyperarp.fr/?lang=en")
        'End If
        Process.Start("https://compomusic.fr")
    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs)
        If EnChargement = False Then ' 
            'ChoixTypeAccord()
            CAD_Maj_TableGlobalAcc()
        End If
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs)
        ModeSimple_Cadence = "Maj" ' pour Ouvrir/Enregsitrer
        ComboBox4.Select(0, 0)
        '
        If EnChargement = True Or ChangementLangue = False Then
            Cad_OrigineAccord = Modes.Cadence_Majeure
            OrigineAccord = Modes.Cadence_Majeure
            Mode_Cadence = Cad_OrigineAccord
            '
            Cad_RAZ_CouleurMarquée()
            '
            Select Case ComboBox3.Text
                Case "Anatole"
                    Cad_AnatoleMaj()
                    Label28.Text = "Anatole"
                '
                Case "Complète"
                    Cad_ComplèteMaj()
                    Label28.Text = "Complète"

                Case "Complete"
                    Cad_ComplèteMaj()
                    Label28.Text = "Complete"
                '
                Case "2-5-1"
                    Cad_251Maj()
                    Label28.Text = "2-5-1"
                '
                Case "Demi"
                    Cad_DemiMaj()
                    Label28.Text = "Demi"

                Case "Half"
                    Cad_DemiMaj()
                    Label28.Text = "Half"
                '
                Case "Parfaite"
                    Cad_ParfaiteMaj()
                    Label28.Text = "Parfaite"

                Case "Perfect"
                    Cad_ParfaiteMaj()
                    Label28.Text = "Perfect"
                '
                Case "Plagale"
                    Cad_PlagaleMaj()
                    Label28.Text = "Plagale"

                Case "Plagal"
                    Cad_PlagaleMaj()
                    Label28.Text = "Plagal"
                '
                Case "Plagale2"
                    Cad_Plagale2Maj()
                    Label28.Text = "Plagale2"

                Case "Plagal2"
                    Cad_Plagale2Maj()
                    Label28.Text = "Plagal2"
                '
                Case "Rompue"
                    Cad_RompueMaj()
                    Label28.Text = "Rompue"

                Case "Broken"
                    Cad_RompueMaj()
                    Label28.Text = "Broken"
                '
                Case "Rompue2"
                    Cad_Rompue2Maj()
                    Label28.Text = "Rompue2"

                Case "Broken2"
                    Cad_Rompue2Maj()
                    Label28.Text = "Broken2"
                '
                Case "Rompue3"
                    Cad_Rompue3Maj()
                    Label28.Text = "Rompue3"

                Case "Broken3"
                    Cad_Rompue3Maj()
                    Label28.Text = "Broken3"

                Case "Modale"
                    Cad_ModaleMaj()
                    Label28.Text = "Modale"

                Case "Modal"
                    Cad_ModaleMaj()
                    Label28.Text = "Modal"
                '
                Case "Modale2"
                    Cad_Modale2Maj()
                    Label28.Text = "Modale2"

                Case "Modal2"
                    Cad_Modale2Maj()
                    Label28.Text = "Modal2"
                '
                Case "Modale3"
                    Cad_Modale3Maj()
                    Label28.Text = "Modale3"

                Case "Modal3"
                    Cad_Modale3Maj()
                    Label28.Text = "Modal3"

                Case "Napolitaine"
                    Cad_NapolitaineMaj()
                    Label28.Text = "Napolitaine"

                Case "Napolitan"
                    Cad_NapolitaineMaj()
                    Label28.Text = "Napolitan"
                    '
            End Select
            'Mode_Cadence = "Maj"
        End If
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs)
        ModeSimple_Cadence = "Min" ' pour Ouvrir/Enregsitrer
        ComboBox3.Select(0, 0)
        '
        If EnChargement = False And ChangementLangue = False Then
            '
            'Mode_Cadence = "Min"
            '

            Cad_OrigineAccord = Modes.Cadence_Mineure
            OrigineAccord = Modes.Cadence_Mineure
            Mode_Cadence = Cad_OrigineAccord
            '
            Cad_RAZ_CouleurMarquée()
            Select Case ComboBox4.Text 'ComboBox4.Text
                Case "Anatole Min"
                    Cad_AnatoleMin()
                    Label28.Text = "Anatole Min"
                Case "Pseudo 2-5-1"
                    Cad_Pseudo251Min()
                    Label28.Text = "Pseudo 2-5-1"
                '
                Case "Plagale", "Plagale Min"
                    Cad_PLagalMin()
                    Label28.Text = "Plagale Min"

                Case "Plagal", "Minor Plagal"
                    Cad_PLagalMin()
                    Label28.Text = "Minor Plagal"
                '
                Case "Hispanique"
                    Cad_OrigineAccord = Modes.Cadence_Mixte
                    OrigineAccord = Modes.Cadence_Mixte
                    Mode_Cadence = Cad_OrigineAccord
                    Cad_HispaniqueMixte()
                    Label28.Text = "Hispanique"
                Case "Hispanic"
                    Cad_OrigineAccord = Modes.Cadence_Mixte
                    OrigineAccord = Modes.Cadence_Mixte
                    Mode_Cadence = Cad_OrigineAccord
                    Cad_HispaniqueMixte()
                    Label28.Text = "Hispanic"
            End Select
        End If
    End Sub



    Private Sub DébutZ_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ComboBox23_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox23.SelectedIndexChanged
        If EnChargement = False Then ' 
            ChoixTypeAccord()
            ' CAD_Maj_TableGlobalAcc(Mode_Cadence)
        End If
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        ModeSimple_Cadence = "Maj" ' pour Ouvrir/Enregsitrer
        ComboBox4.Select(0, 0)
        '
        If EnChargement = False Then ' Or ChangementLangue = False
            Cad_OrigineAccord = Modes.Cadence_Majeure
            OrigineAccord = Modes.Cadence_Majeure
            Mode_Cadence = Cad_OrigineAccord
            '
            Cad_RAZ_CouleurMarquée()
            '
            Select Case ComboBox3.Text
                Case "Anatole"
                    Cad_AnatoleMaj()
                    Label28.Text = "Anatole"

                Case "Forme2"
                    Cad_Forme2()
                    Label28.Text = "Forme2"

                Case "Forme3"
                    Cad_Forme3()
                    Label28.Text = "Forme3"
                '
                Case "Complète"
                    Cad_ComplèteMaj()
                    Label28.Text = "Complète"

                Case "Complete"
                    Cad_ComplèteMaj()
                    Label28.Text = "Complete"
                '
                Case "2-5-1"
                    Cad_251Maj()
                    Label28.Text = "2-5-1"
                '
                Case "Demi"
                    Cad_DemiMaj()
                    Label28.Text = "Demi"

                Case "Half"
                    Cad_DemiMaj()
                    Label28.Text = "Half"
                '
                Case "Parfaite"
                    Cad_ParfaiteMaj()
                    Label28.Text = "Parfaite"

                Case "Perfect"
                    Cad_ParfaiteMaj()
                    Label28.Text = "Perfect"
                '
                Case "Plagale"
                    Cad_PlagaleMaj()
                    Label28.Text = "Plagale"

                Case "Plagal"
                    Cad_PlagaleMaj()
                    Label28.Text = "Plagal"
                '
                Case "Plagale2"
                    Cad_Plagale2Maj()
                    Label28.Text = "Plagale2"

                Case "Plagal2"
                    Cad_Plagale2Maj()
                    Label28.Text = "Plagal2"
                '
                Case "Rompue"
                    Cad_RompueMaj()
                    Label28.Text = "Rompue"

                Case "Broken"
                    Cad_RompueMaj()
                    Label28.Text = "Broken"
                '
                Case "Rompue2"
                    Cad_Rompue2Maj()
                    Label28.Text = "Rompue2"

                Case "Broken2"
                    Cad_Rompue2Maj()
                    Label28.Text = "Broken2"
                '
                Case "Rompue3"
                    Cad_Rompue3Maj()
                    Label28.Text = "Rompue3"

                Case "Broken3"
                    Cad_Rompue3Maj()
                    Label28.Text = "Broken3"

                Case "Modale"
                    Cad_ModaleMaj()
                    Label28.Text = "Modale"

                Case "Modal"
                    Cad_ModaleMaj()
                    Label28.Text = "Modal"
                '
                Case "Modale2"
                    Cad_Modale2Maj()
                    Label28.Text = "Modale2"

                Case "Modal2"
                    Cad_Modale2Maj()
                    Label28.Text = "Modal2"
                '
                Case "Modale3"
                    Cad_Modale3Maj()
                    Label28.Text = "Modale3"

                Case "Modal3"
                    Cad_Modale3Maj()
                    Label28.Text = "Modal3"

                Case "Napolitaine"
                    Cad_NapolitaineMaj()
                    Label28.Text = "Napolitaine"

                Case "Napolitan"
                    Cad_NapolitaineMaj()
                    Label28.Text = "Napolitan"
                    '
            End Select
            'Mode_Cadence = "Maj"
        End If
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        ModeSimple_Cadence = "Min" ' pour Ouvrir/Enregsitrer
        ComboBox3.Select(0, 0)
        '
        If EnChargement = False Then ' And ChangementLangue = False
            '
            'Mode_Cadence = "Min"
            '

            Cad_OrigineAccord = Modes.Cadence_Mineure
            OrigineAccord = Modes.Cadence_Mineure
            Mode_Cadence = Cad_OrigineAccord
            '
            Cad_RAZ_CouleurMarquée()
            Select Case ComboBox4.Text 'ComboBox4.Text
                Case "Anatole Min"
                    Cad_AnatoleMin()
                    Label28.Text = "Anatole Min"


                Case "Pseudo 2-5-1"
                    Cad_Pseudo251Min()
                    Label28.Text = "Pseudo 2-5-1"
                '
                Case "Plagale", "Plagale Min"
                    Cad_PLagalMin()
                    Label28.Text = "Plagale Min"

                Case "Plagal", "Minor Plagal"
                    Cad_PLagalMin()
                    Label28.Text = "Minor Plagal"
                '
                Case "Hispanique"
                    Cad_OrigineAccord = Modes.Cadence_Mixte
                    OrigineAccord = Modes.Cadence_Mixte
                    Mode_Cadence = Cad_OrigineAccord
                    Cad_HispaniqueMixte()
                    Label28.Text = "Hispanique"

                Case "Hispanic"
                    Cad_OrigineAccord = Modes.Cadence_Mixte
                    OrigineAccord = Modes.Cadence_Mixte
                    Mode_Cadence = Cad_OrigineAccord
                    Cad_HispaniqueMixte()
                    Label28.Text = "Hispanic"
            End Select
        End If
    End Sub

    Private Sub ComboBox6_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        If EnChargement = False Then ' 
            'ChoixTypeAccord()
            CAD_Maj_TableGlobalAcc()
        End If
    End Sub
    Private Sub Grid2_DragDrop(sender As Object, e As DragEventArgs) Handles Grid2.DragDrop
        Colonne_Drag = Grid2.MouseCol
        Ligne_Drag = Grid2.MouseRow
        Valeur_Drag = e.Data.GetData(DataFormats.Text)
        Flag_EcrDragDrop = True
        ' GridDest = TGridDest.Grid2
    End Sub
    Sub Maj_DragDrop()
        Dim a, b, c, d As String
        Dim colonne As Integer
        Dim f As Integer

        colonne = Grid2.MouseCol
        f = Det_NumDerAccord()
        If colonne > 0 And colonne <= nbMesures Then
            a = Valeur_Drag ' a = texte à glisser/déposer dans grid2
            Flag_EcrDragDrop = True
            If Insert.Checked And f < nbMesures And Trim(Grid2.Cell(2, colonne).Text) <> "" Then
                SAUV_Annuler(1, f + 1)
                AccInsérer(colonne)
            Else
                SAUV_Annuler(colonne, colonne)
            End If
            'a = TradD(a) ' les accords choisis sont exprimés en #
            EcritureAccordDsGrid2(a, colonne) ' EcritureAccordDsGrid2(a, Grid2.MouseCol)
            Maj_Répétition()
            CtrlConsistVar()
            ' Mise à jour des piano roll
            ' *************************
            a = Det_ListAcc()
            b = Det_ListGam()
            c = Det_ListMarq()
            d = Trim(Det_ListTon())
            For i = 0 To nb_PianoRoll - 1
                If PIANOROLLChargé(i) = True Then
                    listPIANOROLL(i).PListAcc = Trim(a) 'Det_ListAcc()
                    listPIANOROLL(i).PListGam = Trim(b) 'Det_ListGam()
                    listPIANOROLL(i).PListMarq = Trim(c) 'Det_ListMarq()
                    listPIANOROLL(i).PListTon = d
                    listPIANOROLL(i).F1_Refresh()
                    listPIANOROLL(i).Maj_CalquesMIDI()
                End If
            Next

            Automation1.PListAcc = Det_ListAcc()
            Automation1.PListMarq = Det_ListMarq()
            Automation1.F4_Refresh()
            ' mise à jour du DrumEdit
            Drums.PListAcc = Det_ListAcc()
            Drums.PListMarq = Det_ListMarq()
            Drums.F2_Refresh()

            ' Mettre à jour selon auto voicing
            Calcul_AutoVoicingZ()
        End If
        '
        Grid2.Cell(1, colonne).SetFocus()
        '
        ' Onglet modulation
        ' *****************
        Maj_ModulationRadioB2(colonne)
        LabModulat.Item(0).Text = "---"
        LabModulat.Item(1).Text = "---"
        LabModulat.Item(2).Text = "---"
        LabModulat.Item(3).Text = "---"
    End Sub
    Sub Maj_CopieAcc(mesure As Integer)
        ' Pour CTRL Z
        ' ************
        '
        Maj_Répétition()
        CtrlConsistVar()
        '
        ' Mise à jour des piano roll
        ' *************************
        a = Det_ListAcc()
        b = Det_ListGam()
        c = Det_ListMarq()
        d = Trim(Det_ListTon())
        For i = 0 To nb_PianoRoll - 1
            If PIANOROLLChargé(i) = True Then
                listPIANOROLL(i).PListAcc = Trim(a) 'Det_ListAcc()
                listPIANOROLL(i).PListGam = Trim(b) 'Det_ListGam()
                listPIANOROLL(i).PListMarq = Trim(c) 'Det_ListMarq()
                listPIANOROLL(i).PListTon = d
                listPIANOROLL(i).F1_Refresh()
                listPIANOROLL(i).Maj_CalquesMIDI()
            End If
        Next

        Automation1.PListAcc = Det_ListAcc()
        Automation1.PListMarq = Det_ListMarq()
        Automation1.F4_Refresh()

        ' mise à jour du DrumEdit
        Drums.PListAcc = Det_ListAcc()
        Drums.PListMarq = Det_ListMarq()
        Drums.F2_Refresh()

        ' Mettre à jour selon auto voicing
        Calcul_AutoVoicingZ()

    End Sub



    Public Function TradD(note As String) As String ' traduction d'une note bémol en dièse
        Dim a As String = note
        Dim b As String = note


        If InStr(a, "b") <> 0 Then ' détection d'un bémol dans la note
            Select Case a
                Case "Db"
                    b = "C#"
                Case "Eb"
                    b = "D#"
                Case "Gb"
                    b = "F#"
                Case "Ab"
                    b = "G#"
                Case "Bb"
                    b = "A#"

            End Select
        End If
        Return Trim(b)
    End Function
    Public Function TradB(note As String) As String ' traduction d'une note dièse en bémol
        Dim a As String = note
        Dim b As String = note

        'AA = Microsoft.VisualBasic.Left(a, 2)
        If InStr(a, "#") <> 0 Then
            Select Case a
                Case "C#"
                    b = "Db"
                Case "D#"
                    b = "Eb"
                Case "F#"
                    b = "Gb"
                Case "G#"
                    b = "Ab"
                Case "A#"
                    b = "Bb"

            End Select
        End If
        Return Trim(b)
    End Function
    '
    Public Function TradDB(note As String, signe As String) As String
        TradDB = note
        If signe = "#" Then Return TradD(note)
        If signe = "b" Then Return TradB(note)

    End Function
    Sub AccInsérer(colonne As Integer)
        Dim i As Integer
        Dim j As Integer
        Dim AA As String
        Dim f As Integer = Det_NumDerAccord()
        Dim liste1 As New List(Of Accgrid2)
        Dim liste2 As New List(Of EventH)



        ' sauvegarder
        If Insert.Checked And f < nbMesures Then
            For i = colonne To f 'Arrangement1.nbMesures
                'If Trim(Grid2.Cell(2, i).Text) <> "" Then
                Dim oo1 As New Accgrid2 With {
                        .Marqueur = Trim(Grid2.Cell(1, i).Text),
                        .Accord = Trim(Grid2.Cell(2, i).Text),
                        .Répet = Trim(Grid2.Cell(3, i).Text),
                        .Magneto = IdentVariation(i),
                        .Gamme = Trim(Grid2.Cell(11, i).Text),
                        .PosOld = i
                    }
                liste1.Add(oo1)

                '
                Dim oo2 As New EventH With {
                    .Accord = TableEventH(i, 1, 1).Accord,
                    .Tonalité = TableEventH(i, 1, 1).Tonalité,
                .Degré = TableEventH(i, 1, 1).Degré,
                .Détails = TableEventH(i, 1, 1).Détails,
                .Gamme = TableEventH(i, 1, 1).Gamme,
                .Ligne = TableEventH(i, 1, 1).Ligne,
                .Marqueur = TableEventH(i, 1, 1).Marqueur,
                .Mode = TableEventH(i, 1, 1).Mode,
                .NumAcc = i,
                .NumMagnéto = TableEventH(i, 1, 1).NumMagnéto,
                .Répet = TableEventH(i, 1, 1).Répet,
                .Position = Trim(Str(i) + "." + "1" + "." + "1")
                }
                liste2.Add(oo2)
                'End If
            Next
            ' décaler
            For Each a As Accgrid2 In liste1
                j = liste1.IndexOf(a) + colonne + 1
                '
                If TableEventH(j - 1, 1, 1).Ligne <> -1 Then
                    Grid2.Cell(1, j).Text = a.Marqueur ' marqueur
                    Grid2.Cell(2, j).Text = a.Accord   ' accord
                    Grid2.Cell(3, j).Text = a.Répet    ' répétition
                    Grid2.Cell(11, j).Text = a.Gamme   ' gamme

                    AA = TableEventH(j - 1, 1, 1).Tonalité
                    AA = Det_RelativeMajeure2(AA) ' si la tonalité est mineure alors on affiche la couleur de la relative majeure
                    tbl = Split(AA)
                    Grid2.Cell(2, j).BackColor = DicoCouleur.Item(Trim(tbl(0))) ' la couleur est fonction de la tonalité
                    Grid2.Cell(2, j).ForeColor = DicoCouleurLettre.Item(tbl(0))
                    Grid2.Cell(11, j).BackColor = DicoCouleur.Item(Trim(tbl(0))) ' la couleur est fonction de la tonalité
                    Grid2.Cell(11, j).ForeColor = DicoCouleurLettre.Item(tbl(0))
                    ' Magnéto
                    ChoixVariationGrid2(j, j, a.Magneto + 4)
                Else
                    Grid2.Cell(1, j).Text = ""  ' marqueur
                    Grid2.Cell(2, j).Text = ""  ' accord
                    Grid2.Cell(3, j).Text = ""  ' répétition
                    Grid2.Cell(11, j).Text = "" ' gamme
                    Grid2.Cell(2, j).BackColor = Color.White ' la couleur est fonction de la tonalité
                    Grid2.Cell(2, j).ForeColor = Color.Black
                    Grid2.Cell(11, j).BackColor = Color.White  ' la couleur est fonction de la tonalité
                    Grid2.Cell(11, j).ForeColor = Color.Black
                End If
            Next
            For Each a As EventH In liste2
                i = liste2.IndexOf(a)
                j = liste2.IndexOf(a) + colonne + 1
                TableEventH(j, 1, 1).Accord = liste2.Item(i).Accord
                TableEventH(j, 1, 1).Tonalité = liste2.Item(i).Tonalité
                TableEventH(j, 1, 1).Degré = liste2.Item(i).Degré
                TableEventH(j, 1, 1).Répet = liste2.Item(i).Répet
                TableEventH(j, 1, 1).Détails = liste2.Item(i).Détails
                TableEventH(j, 1, 1).Gamme = liste2.Item(i).Gamme
                TableEventH(j, 1, 1).Ligne = liste2.Item(i).Ligne
                TableEventH(j, 1, 1).Marqueur = liste2.Item(i).Marqueur
                TableEventH(j, 1, 1).Mode = liste2.Item(i).Mode
                TableEventH(j, 1, 1).NumAcc = liste2.Item(i).NumAcc
                TableEventH(j, 1, 1).NumMagnéto = liste2.Item(i).NumMagnéto
                TableEventH(j, 1, 1).Position = liste2.Item(i).Position
            Next
        End If
        '
        Insert.Checked = False
        Grid2.Refresh()
    End Sub
    Function Det_PlaceEcriture(col As Integer) As Integer
        Dim i As Integer
        Det_PlaceEcriture = -1
        If col <> 0 Then
            Det_PlaceEcriture = col
            If Trim(Grid2.Cell(2, col).Text) = "" Then
                i = col
                Do
                    i = i - 1
                Loop Until Grid2.Cell(2, i).Text <> "" Or i = 0
                Det_PlaceEcriture = i + 1
            End If
        End If
    End Function



    Private Sub Grid2_DragEnter(sender As Object, e As DragEventArgs) Handles Grid2.DragEnter
        If (e.Data.GetDataPresent(DataFormats.Text)) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub Grid3_DragEnter(sender As Object, e As DragEventArgs)
        If (e.Data.GetDataPresent(DataFormats.Text)) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub Grid1_LeaveCell(Sender As Object, e As Grid.LeaveCellEventArgs)  ' ok
        Dim j As Integer
        '
        j = e.NewRow
        '
        'PositionnerCurseursRougeBleu(j)
    End Sub
    Function TradDegré_ChiffreLettre(Chiffre As Integer) As String
        TradDegré_ChiffreLettre = ""
        Select Case Chiffre
            Case 0
                TradDegré_ChiffreLettre = "I"
            Case 1
                TradDegré_ChiffreLettre = "II"
            Case 2
                TradDegré_ChiffreLettre = "III"
            Case 3
                TradDegré_ChiffreLettre = "IV"
            Case 4
                TradDegré_ChiffreLettre = "V"
            Case 5
                TradDegré_ChiffreLettre = "VI"
            Case 6
                TradDegré_ChiffreLettre = "VII"

        End Select
    End Function
    Private Sub Grid2_KeyDown(Sender As Object, e As KeyEventArgs) Handles Grid2.KeyDown
        Dim a As String
        Dim i As Integer
        Dim k As Integer
        Dim mDeb As Integer
        Dim mFin As Integer


        Dim m As Integer
        Dim t As Integer
        Dim ct As Integer


        '
        ' variables pour incrémenter/décrémenter les vélocités

        Dim sortir As Boolean = False
        Dim nbLignesGrid2 = Grid2.Rows
        a = e.KeyData
        mDeb = Grid2.Selection.FirstCol
        mFin = Grid2.Selection.LastCol
        t = 1
        ct = 1
        '

        ' ********************************
        ' * RAZ des variation dans Grid2 *
        ' ********************************

        'NbDivMesure = Det_NbDivisionMesure()
        'If DerGridCliquée = GridCours.Grid2 Then
        'If a = "46" And Grid2.ActiveCell.Col > 1 Then ' 46 = delete Touche Suppr
        ''
        '' Pour CTRL Z
        '' ***********
        'ZAnnulation_Sauvegarde(mDeb, mFin)
        ''
        'Grid2.Range(1, 1, Grid2.Rows - 1, Grid2.Cols - 1).Locked = False
        ''
        'Grid2.AutoRedraw = False
        ''
        '' Grid2.Selection.ClearText() ' <-- semble ne pas toujours fonctionner
        ''
        'For i = mDeb To mFin
        'Grid2.Column(i).Locked = True
        'Grid2.Cell(Grid2.ActiveCell.Row, i).BackColor = Color.White
        'Grid2.Cell(Grid2.ActiveCell.Row, i).ForeColor = Color.Black
        'Grid2.Cell(Grid2.ActiveCell.Row, i).Text = ""
        'Next
        ''
        ''
        'Grid2.Range(1, 1, Grid2.Rows - 1, Grid2.Cols - 1).Locked = False
        '
        'Grid2.Refresh()
        'Grid2.AutoRedraw = True
        'End If
        'End If

        '****************************************
        ' Incrémenter/Décrémenter les racines *
        '****************************************
        If (e.KeyCode = Keys.Add Or e.KeyCode = Keys.Subtract) And OK_KeyDown = True Then ' And Grid2.Selection.FirstCol = 8
            i = Grid2.ActiveCell.Row
            j = Grid2.ActiveCell.Col

            Grid2.AutoRedraw = False
            'Grid2.Range(0, 8, nbLignesGrid2 - 1, 8).Locked = False ' Vérouillage des cellules
            DéVérouillerGrid2()
            ' arrowdirection
            b = CellDyn_Racines()
            ' traitement des dynamiques avec les touches + et -
            If Trim(b) <> "" Then
                tbl1 = Split(b)
                For ii = 0 To UBound(tbl1)
                    tbl2 = Split(tbl1(ii), ",")
                    i = Convert.ToInt16(tbl2(0)) ' ligne - - - 
                    j = Convert.ToInt16(tbl2(1)) ' col
                    '
                    'k = Convert.ToInt16(Grid1.Cell(i, j).Text)
                    k = TRacine2.IndexOf(Trim(Grid2.Cell(i, j).Text))
                    ' détermination de la nécessité d'incrémenter
                    If k >= TRacine2.Count - 1 And e.KeyCode = Keys.Add Then
                        sortir = True ' si l'une des valeurs =max alors on n'augmente plus aucune valeur de la sélection --> sortir = true
                        Exit For
                    End If
                    ' détermination de la nécessité de décrémenter
                    If k <= 0 And e.KeyCode = Keys.Subtract Then
                        sortir = True ' si l'une des valeurs =max alors on n'augmente plus aucune valeur de la sélection  --> sortir = true
                        Exit For
                    End If
                    '
                Next
                ' incrémentation/décrémentation (si sortir = false)
                If sortir = False Then
                    'For j = Grid2.Selection.FirstCol To Grid2.Selection.LastCol
                    'b = j.ToString + ".1.1" 'Det_Index_De_Ligne(Grid2.ActiveCell.Row)
                    'Next j
                    ''
                    'tbl = Split(b, "-")
                    'm = tbl(0)
                    't = tbl(1)
                    'ct = tbl(2)

                    For ii = 0 To UBound(tbl1)
                        tbl2 = Split(tbl1(ii), ",")
                        i = Convert.ToInt16(tbl2(0))
                        j = Convert.ToInt16(tbl2(1))
                        a = Trim(Grid2.Cell(i, j).Text)
                        '
                        b = Trim(j.ToString + ".1.1")
                        tbl3 = b.Split(".")
                        m = Convert.ToInt16(tbl3(0))
                        t = Convert.ToInt16(tbl3(1))
                        ct = Convert.ToInt16(tbl3(2))
                        'If IsNumeric(a) Then
                        If e.KeyCode = Keys.Add Then
                            'k = Convert.ToInt16(a) + 1
                            k = TRacine2.IndexOf(a)
                            k += 1
                            If k <= TRacine2.Count - 1 Then
                                Grid2.Cell(i, j).Text = TRacine2.Item(k)    '
                                TableEventH(m, t, ct).Racine = Trim(TRacine2.Item(k))
                                JouerAcc(b)
                            End If
                        ElseIf e.KeyCode = Keys.Subtract Then
                            'k = Convert.ToInt16(a) - 1
                            k = TRacine2.IndexOf(Trim(Grid2.Cell(i, j).Text))
                            k -= 1
                            If k >= 0 Then
                                Grid2.Cell(i, j).Text = TRacine2.Item(k)
                                TableEventH(m, t, ct).Racine = Trim(TRacine2.Item(k))
                                JouerAcc(b)
                            End If
                        End If
                    Next
                    ' remarque : Calcul_AutoVoicingZ() est appelé dans JouerAcc

                End If
            End If
            'Grid2.Range(0, 8, Grid2.Rows - 1, Grid2.Cols - 1).Locked = True ' Vérouillage des cellules racines
            OK_KeyDown = False  ' pour incrémentation/décrémentation des racines (évite les répétitions par appui continu)
            VérouillerGrid2()
            Grid2.AutoRedraw = True
            Grid2.Refresh()
        End If

    End Sub
    ''' <summary>
    ''' Joueur accord dans Grid à une position données
    ''' </summary>
    ''' <param name="Position">Position (mesure) de l'accord à jouer</param>
    Sub JouerAcc(Position As String)
        Dim m, t, ct As Integer
        Dim tbl() As String

        If ComboMidiOut.Items.Count > 0 Then
            If EnChargement = False Then        ' jouer "Accord"
                If AccordAEtéJoué = False Then
                    tbl = Split(Trim(Position), ".")
                    m = Val(tbl(0))
                    t = Val(tbl(1))
                    ct = Val(tbl(2))
                    b = TableEventH(m, t, ct).Tonalité
                    tbl = Split(b)
                    Clef = Det_Clef(tbl(0))
                    '
                    JouerAccord123(Trim(Position))
                End If
            End If
            '
        End If
    End Sub
    Sub JouerAccord123(Position As String)
        Dim a As String
        Dim tbl() As String
        Dim tbl2() As String
        Dim m, t, ct As Integer
        Dim n, n1 As Integer
        Dim i As Integer
        Dim Tonique As String
        Dim Note As String
        Dim Sauv_Clef As String


        'i = Val(Label50.Text)
        'i = i + 1
        'Label50.Text = Str(i)

        Try
            Sauv_Clef = Trim(Clef)
            ' Calcul indexes des positions
            ' ****************************

            If AccordAEtéJoué = False And Trim(Position) <> "" Then
                '
                'tbl = Split(ComboBox20.Text)
                tbl = Split(Trim(Position), ".")
                m = Val(tbl(0))
                t = Val(tbl(1))
                ct = Val(tbl(2))
                '
                ' Détermination de la Clef (b ou #) de la tonalité de l'accord
                ' ************************************************************
                a = TableEventH(m, t, ct).Tonalité

                tbl2 = Split(a)
                Clef = Det_Clef(tbl2(0))
                '
                Calcul_AutoVoicingZ()
                '
                If TableEventH(m, t, ct).Ligne <> -1 Then
                    ' Détermination de la Tonique : pour affichage en rouge sur le clavier
                    ' ********************************************************************
                    tbl = Split(Det_NotesAccord(TableEventH(m, t, ct).Accord), "-")
                    Tonique = Trim(tbl(0))
                    Tonique = Trad_ListeNotesEnD(Trim(Tonique), "-") ' traduire la tonique en #
                    ' Identification des notes de l'accord 
                    ' ************************************
                    a = TableNotesAccordsZ(m, t, ct) ' ici les notes sont toujours en # - Maj effectuée par Calcul_AutoVoicingZ()
                    tbl = Split(a)
                    ' Initialisations avant jeu de l'accord
                    ' *************************************
                    If Not (SortieMidi.Item(ChoixSortieMidi).IsOpen) Then
                        SortieMidi.Item(ChoixSortieMidi).Open()
                    End If
                    AccordJouerPiano.Raz_NotesJouéesPiano()
                    '
                    ' Jeu de l'accord (boucle sur les notes de l'accords)
                    ' **************************************************
                    NotesAff = ""
                    Num_octave = -1
                    For i = 0 To UBound(tbl)
                        n = ListNotesd.IndexOf(Trim(tbl(i)))
                        n1 = Val(Microsoft.VisualBasic.Right(tbl(i), 1))
                        If n1 > Num_octave Then
                            Num_octave = n1
                        End If
                        'SortieMidi.Item(ChoixSortieMidi).SendNoteOn(CanalThru.Value - 1, n, PlaybackVelocity.Value)
                        SortieMidi.Item(ChoixSortieMidi).SendNoteOn(Convert.ToByte(CanalEcoute.Text), n, VéloEcoute.Value)
                        'AccordJouerPiano.Notes(i) = n
                        'AccordJouerPiano.OldBackColor(i) = LabelPiano.Item(n).BackColor
                        'LabelPiano.Item(n).BackColor = Color.Yellow
                        '
                        Note = Det_NoteSansOctave(Trim(tbl(i)))
                        ''
                        ' Calcul de la clef pour affichage des notes
                        ' ******************************************
                        'a = TableEventH(m, t, ct).Tonalité
                        'tbl2 = Split(a)
                        'Clef = Trim(Det_Clef(Trim(tbl2(0))))
                        ''
                        'Dim nn As Integer
                        'If Clef = "b" Then
                        'nn = ListNotesb.IndexOf(ListNotesb(n))
                        'b = ListNotesd(nn)
                        'LabelPiano.Item(n).Text = b
                        'NotesAff = NotesAff + b + " "
                        'Else
                        'LabelPiano.Item(n).Text = ListNotesd(n)
                        'NotesAff = NotesAff + ListNotesd(n) + " "
                        'End If

                        'If Note = Tonique Then
                        'LabelPiano.Item(n).ForeColor = Color.Red
                        'Else
                        'LabelPiano.Item(n).ForeColor = Color.Blue
                        'End If
                    Next i
                    '
                    AccordAEtéJoué = True
                End If
            End If
            '
            'NotesAff = Trim(NotesAff)
            Clef = Sauv_Clef
        Catch ex As Exception
            messa = "Problème de ressource MIDI"
            MessageHV.PContenuMess = messa + Constants.vbCrLf + "Détection d'une erreur dans procédure : " + "JouerAccord123" + "." + Constants.vbCrLf +
            "Message  : " + ex.Message
            MessageHV.PTypBouton = "OK"
            MessageHV.ShowDialog()
            End
        End Try
        'End If
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns>Retourne dans une chaine les coordonnées lignes,colonnes cellules concernées avec la syntaxe suivante : ligne1,Colonne1 Ligne2,Colonne 2 Ligne3,Colonne 3 etc..</returns>
    Function CellDyn_Racines() As String
        Dim i As Integer = Grid2.Rows - 1
        Dim j As Integer
        Dim a As String = ""
        CellDyn_Racines = a

        For j = Grid2.Selection.FirstCol To Grid2.Selection.LastCol
            If Trim(Grid2.Cell(i, j).Text) <> "" Then
                a = a + Convert.ToString(i) + "," + Convert.ToString(j) + " "
            End If
        Next
        CellDyn_Racines = Trim(a)
    End Function
    Private Sub Grid2_Scroll(Sender As Object, e As EventArgs) Handles Grid2.Scroll
        TMarqueur.Visible = False
    End Sub


    Private Function ListeNuméroNoteAccG4(Ligne As Integer) As String
        Dim Position As String
        Dim tbl() As String
        Dim a As String
        Dim m, t, ct As Integer
        Dim n As Integer

        ' Détermination des paramètres m,t,ct de l'EventH
        ' ************************************************
        Position = LCaractLigne.Item(Ligne).Position       'Trim(Grid1.Cell(Ligne, 1).Text)
        tbl = Split(Trim(Position), ".")
        m = Val(tbl(0))
        t = Val(tbl(1))
        ct = Val(tbl(2))
        '
        ' Détermination de la Clef (b ou #) de la tonalité de l'accord
        ' ************************************************************
        Maj_TabNotes_Minus("#")
        ' Calcul du Voicing de l'accord
        ' *****************************
        Calcul_AutoVoicingZ()
        ' Identification des notes de l'accord 
        ' ************************************
        a = TableNotesAccordsZ(m, t, ct) ' ici les notes sont toujours en # - Maj effectuée par Calcul_AutoVoicingZ()
        tbl = Split(a)
        a = ""
        ' Détermination des N° de notes de l'accord
        ' ****************************************
        For i = 0 To UBound(tbl)
            n = ListNotesd.IndexOf(Trim(tbl(i))) '
            a = Trim(a + Str(n) + " ")
        Next i
        ListeNuméroNoteAccG4 = Trim(a)
    End Function
    Private Sub EnvNotesOffGamme()
        Dim k As Integer
        For k = 0 To UBound(tbl_NotesOffG)
            If tbl_NotesOffG(k) <> 0 Then
                CouperNote2(tbl_NotesOffG(k))
                tbl_NotesOffG(k) = 0
            End If
        Next k
    End Sub
    Private Sub EnvNotesOffAccord()
        Dim k As Integer
        For k = 0 To UBound(tbl_NotesOffA)
            CouperNote2(tbl_NotesOffA(k))
            tbl_NotesOffA(k) = 0
        Next k
    End Sub

    Private Sub Button34_MouseDown(sender As Object, e As MouseEventArgs)
        'Envoi d'un accord
        ' ****************
        EnvAccord = True
        NotesAJouer = "60" + " " + "64" + " " + "67"
        tbl_NotesOnA = Split(NotesAJouer)
        ReDim Preserve tbl_NotesOffA(UBound(tbl_NotesOnA))
        For i = 0 To UBound(tbl_NotesOffA)
            tbl_NotesOffA(i) = 0
        Next i
        'BGW.WorkerSupportsCancellation = True
        'BGW.RunWorkerAsync()
        'Envoi d'une gamme
        ' ****************
        EnvGamme = True
        NotesAJouer = "60" + " " + "62" + " " + "64" + " " + "65" + " " + "67"
        tbl_NotesOnG = Split(NotesAJouer)
        ReDim Preserve tbl_NotesOffG(UBound(tbl_NotesOnG))
        For i = 0 To UBound(tbl_NotesOffG)
            tbl_NotesOffG(i) = 0
        Next i
        BGW.WorkerSupportsCancellation = True
        BGW.RunWorkerAsync()
    End Sub

    Sub CouperNote2(n As Byte)

        If Not (SortieMidi.Item(ChoixSortieMidi).IsOpen) Then
            SortieMidi.Item(ChoixSortieMidi).Open()
        End If
        'SortieMidi.Item(ChoixSortieMidi).SendNoteOff(CanalThru.Value - 1, n, 0)
        'SortieMidi.Item(ChoixSortieMidi).SilenceAllNotes()
    End Sub
    Function Init_P() As Point
        Dim p As New Point
        Init_P.X = 5 '
        Init_P.Y = 9 '
    End Function
    Function Init_PP() As Point
        Dim p As New Point
        Init_PP.X = 5 '
        Init_PP.Y = 15 '
    End Function
    Public Function Det_NumDerAccord() As Integer

        Dim i As Integer = nbMesures
        Grid2.Refresh()
        Det_NumDerAccord = 1 ' cas où seul le dernier accord est présent
        While Trim(Grid2.Cell(2, i).Text) = "" And i >= 1
            i = i - 1
        End While
        Det_NumDerAccord = i

    End Function
    Public Function Det_NumDerAccord2() As Integer
        Dim chronometre As New Stopwatch()
        chronometre.Start()
        Grid2.Refresh()
        Dim i As Integer = nbMesures
        Det_NumDerAccord2 = -1 ' cas où seul le dernier accord est présent
        While Trim(Grid2.Cell(2, i).Text) = "" And i >= 1
            i = i - 1
        End While
        If i >= 1 Then Det_NumDerAccord2 = i
        chronometre.Stop()
        Dim duree As TimeSpan = chronometre.Elapsed
    End Function
    Public Function Det_NumDerRépet() As Integer
        Dim i As Integer = nbMesures
        Det_NumDerRépet = 1 ' cas où pas d'accord trouvé
        While Trim(Grid2.Cell(2, i).Text) = "" And i >= 1
            i = i - 1
        End While
        If i >= 1 Then Det_NumDerRépet = Convert.ToInt16(Grid2.Cell(3, i).Text)
    End Function
    Sub Maj_Répétition2(DerColSel As Integer)
        Dim i, j, k As Integer
        Dim f As Integer = Det_NumDerAccord()

        If f <> -1 Then ' f=-1 signifie qu'il n'y a aucun accord par exemple après un 'couper'
            ' 1 - traitement du dernier accord
            ' ********************************

            If Trim(Grid2.Cell(3, f).Text) <> "" Then
                If f < DerColSel Then
                    For i = f To nbMesures - 1
                        Grid2.Cell(3, i).Text = "1"
                        TableEventH(i, 1, 1).Répet = "1"
                    Next
                End If
                f = f - 1 ' on traite
            End If

            ' 2 - traitement des répétitions des accords sauf le dernier
            ' **********************************************************
            For i = 1 To f '- 1
                Grid2.Cell(3, i).Text = ""
                If Trim(Grid2.Cell(2, i).Text) <> "" Then
                    j = i + 1
                    k = 1
                    While Trim(Grid2.Cell(2, j).Text) = "" And j <= f
                        j = j + 1
                        k = k + 1
                    End While
                    ' maj répétition
                    Grid2.Cell(3, i).Text = Convert.ToString(k)
                    TableEventH(i, 1, 1).Répet = Convert.ToByte(k)
                    Grid2.Refresh()
                End If
            Next
            '
            ' 3 - Maj Répétition manuel sur dernier accord
            ' ********************************************
            If ColAncienDerAcc = Det_NumDerAccord() Then
                Grid2.Cell(3, ColAncienDerAcc).Text = Trim(ReptAncienDerAcc)
            End If

        End If
    End Sub
    Sub Maj_Répétition()
        Dim i, j, k As Integer
        Dim f As Integer = Det_NumDerAccord()


        ' A VOIR
        ' dans la ligne suivante, on permet de garder la répétition si la copie porte sur le dernier accord
        ' pour le moment c'est supprimé
        ' d'autre part , il faut signaler dans le manuel vidéo qu'il est possible de modifier la répétition du dernier accord 

        ' *********************************************************************************************************
        'If Not (Trim(Grid2.Cell(3, f).Text) = "") Then f = f - 1 ' on ne traite le dernier accord que si il n'a pas de répétition
        ' *********************************************************************************************************



        For i = 1 To f '- 1
            Grid2.Cell(3, i).Text = ""
            If Trim(Grid2.Cell(2, i).Text) <> "" Then
                j = i + 1
                k = 1
                While Trim(Grid2.Cell(2, j).Text) = "" And j <= f 'Arrangement1.nbMesures
                    j = j + 1
                    k = k + 1
                End While
                ' maj répétition
                Grid2.Cell(3, i).Text = Convert.ToString(k)
                TableEventH(i, 1, 1).Répet = Convert.ToByte(k)
                Grid2.Refresh()
            Else
                Grid2.Cell(3, i).Text = "1"
                TableEventH(i, 1, 1).Répet = 1
            End If
        Next
        '
        ' 3 - Maj Répétition manuel sur dernier accord
        ' ********************************************
        If ColAncienDerAcc = Det_NumDerAccord() Then
            Grid2.Cell(3, ColAncienDerAcc).Text = Trim(ReptAncienDerAcc)
        End If

    End Sub
    ' *#######################################################################
    ' ##                                                                    ##
    ' ##                 RECUPERATION DES DONNEES IHM                       ##
    ' ##                                                                    ##
    ' ########################################################################
    ' 
    ' ********************************************************************
    ' Récup_Acc : récupération des accords sous forme d'accords chiffrés *
    ' Les informations sont prévelées directement dans Grid2             *
    ' ********************************************************************
    ' 
    Public ReadOnly Property Récup_Acc() As String
        Get
            Dim tbl() As String
            Dim i As Integer
            Dim j As Integer = 0
            Dim a As String = ""
            Dim deb As Integer

            deb = Début.Value
            ReDim Preserve tbl(0)
            If Trim(Grid2.Cell(2, i).Text) = "" Then
                i = Det_AccordPréced(Début.Value)
                tbl(0) = Trim(Grid2.Cell(2, i).Text)
                deb = Det_AccordSuiv(Début.Value)
                j = j + 1
            End If
            For i = deb To nbMesures '- 1
                If Trim(Grid2.Cell(2, i).Text) <> "" Then
                    If j > UBound(tbl) Then
                        ReDim Preserve tbl(j)
                    End If
                    tbl(j) = (Trim(Grid2.Cell(2, i).Text))
                    j = j + 1
                End If
            Next i
            a = Join(tbl, ",")
            Return a
        End Get
        'Set(ByVal Value As String)
        'Me.Name = Value
        'End Set
    End Property
    ' **********************************************************************
    ' Récup_NumAcc : récupération des n° d'accords dans Grid2.Il s'agit    *
    ' du N° de la colonne contenant l'accord.                              *
    ' **********************************************************************
    ' 
    Public ReadOnly Property Récup_NumAcc() As String
        Get
            Dim tbl() As String
            Dim i As Integer
            Dim j As Integer = 0
            Dim a As String = String.Empty
            Dim deb As Integer
            Dim fin As Integer

            ' gestion du début
            deb = Début.Value
            ReDim Preserve tbl(0)
            If Trim(Grid2.Cell(2, i).Text) = "" Then
                i = Det_AccordPréced(Début.Value)
                tbl(0) = Trim(Grid2.Cell(0, i).Text)
                deb = Det_AccordSuiv(Début.Value)
                j = j + 1
            End If
            '
            ' gestion de la fin
            fin = Det_NumDerAccord()
            'If fin > Terme.Value Then fin = Terme.Value

            For i = deb To fin '- 1
                If Trim(Grid2.Cell(2, i).Text) <> "" Then
                    If j > UBound(tbl) Then
                        ReDim Preserve tbl(j)
                    End If
                    tbl(j) = (Trim(Grid2.Cell(0, i).Text))
                    j = j + 1
                End If
            Next i
            a = Join(tbl, ",")
            Return a
        End Get
        'Set(ByVal Value As String)
        'Me.Name = Value
        'End Set
    End Property
    Public ReadOnly Property Récup_AllNumAcc() As String
        Get
            Dim tbl() As String
            Dim i As Integer
            Dim j As Integer = 0
            Dim a As String = String.Empty
            Dim deb As Integer
            Dim fin As Integer

            ' gestion du début
            deb = 1
            ReDim Preserve tbl(0)
            '
            ' gestion de la fin
            fin = Det_NumDerAccord()
            'If fin > Terme.Value Then fin = Terme.Value

            For i = deb To fin '- 1
                If Trim(Grid2.Cell(2, i).Text) <> "" Then
                    If j > UBound(tbl) Then
                        ReDim Preserve tbl(j)
                    End If
                    tbl(j) = (Trim(Grid2.Cell(0, i).Text))
                    j = j + 1
                End If
            Next i
            a = Join(tbl, ",")
            Return a
        End Get
        'Set(ByVal Value As String)
        'Me.Name = Value
        'End Set
    End Property
    ' ********************************************************************
    ' Récup_NotesAcc : récupération des accords sous forme dde notes     *
    ' Les infos sont prélevées la table des voicing TableNotesAccordsZ   *
    ' ********************************************************************
    Public ReadOnly Property Récup_notesAcc() As String
        Get
            Dim tbl() As String
            Dim i As Integer
            Dim j As Integer = 0
            Dim a As String = String.Empty
            Dim deb As Integer
            Dim fin As Integer


            Calcul_AutoVoicingZ() ' remise à jour de la table TableNotesAccordsZ
            ReDim Preserve tbl(0)

            'gestion début
            deb = Début.Value

            ' gestion fin
            fin = Det_NumDerAccord()
            If fin > Terme.Value Then fin = Terme.Value


            For i = deb To fin '- 1
                a = Trim(TableNotesAccordsZ(i, 1, 1))

                If (a <> "" And a <> Nothing) Then
                    If j > UBound(tbl) Then
                        ReDim Preserve tbl(j)
                    End If
                    tbl(j) = (Trim(UCase(TableNotesAccordsZ(i, 1, 1))))
                    j = j + 1
                End If
            Next i
            a = String.Join(",", tbl)
            Return a
        End Get
    End Property
    '
    Public ReadOnly Property Récup_notesAcc2() As String
        Get
            Dim tbl() As String
            Dim i As Integer
            Dim j As Integer = 0
            Dim a As String = String.Empty
            Dim deb As Integer
            Dim fin As Integer


            Calcul_AutoVoicingZ() ' remise à jour de la table TableNotesAccordsZ
            ReDim Preserve tbl(0)

            ' traitement du début
            deb = Début.Value
            If Trim(TableNotesAccordsZ(deb, 1, 1)) = "" Then
                i = Det_AccordPréced(Début.Value)
                tbl(0) = (Trim(UCase(TableNotesAccordsZ(i, 1, 1))))
                deb = Det_AccordSuiv(Début.Value)
                j = j + 1
            End If

            ' traitement de la fin
            fin = Det_NumDerAccord()
            If fin > Terme.Value Then fin = Terme.Value

            ' traitement des autres accords
            For i = deb To fin '- 1
                a = Trim(Grid2.Cell(2, i).Text)

                If (a <> "" And a <> Nothing) Then
                    If j > UBound(tbl) Then
                        ReDim Preserve tbl(j)
                    End If
                    tbl(j) = (Trim(UCase(TableNotesAccordsZ(i, 1, 1))))
                    j = j + 1
                End If
            Next i
            a = String.Join(",", tbl)
            Return a
        End Get
    End Property
    Function Det_TonalitéIndex(tona As String) As Integer
        Dim ind As Integer = 0
        Select Case tona
            Case "C# Maj", "A# MinH", "A# MinM"
                ind = 0
            Case "F# Maj", "D# MinH", "D# MinM"
                ind = 1
            Case "B Maj", "G# MinH", "G# MinM"
                ind = 2
            Case "E Maj", "C# MinH", "C# MinM"
                ind = 3
            Case "A Maj", "F# MinH", "F# MinM"
                ind = 4
            Case "D Maj", "B MinH", "B MinM"
                ind = 5
            Case "G Maj", "E MinH", "E MinM"
                ind = 6
            Case "C Maj", "A MinH", "A MinM"
                ind = 7
            Case "F Maj", "D MinH", "D MinM"
                ind = 8
            Case "Bb Maj", "G MinH", "G MinM"
                ind = 9
            Case "Eb Maj", "C MinH", "C MinM"
                ind = 10
            Case "Ab Maj", "F MinH", "F MinM"
                ind = 11
        End Select
        Return ind
    End Function


    Function Det_AccordPréced(AccDépart As Integer) As Integer
        Dim i As Integer = AccDépart
        Det_AccordPréced = -1 ' cas où pas d'accord trouvé
        While Trim(Grid2.Cell(2, i).Text) = "" And i >= 1
            i = i - 1
        End While
        If i >= 1 Then Det_AccordPréced = i
    End Function
    '
    Function Det_AccordSuiv(AccDépart As Integer) As Integer
        Dim i As Integer = AccDépart + 1
        Det_AccordSuiv = -1 ' cas où pas d'accord trouvé
        While Trim(Grid2.Cell(2, i).Text) = "" And i <= nbMesures
            i = i + 1
        End While
        If i >= 1 Then Det_AccordSuiv = i
    End Function
    ' **************************************************************************
    ' Récup_Mrq : récupération des marqueurs pour chaque position              *
    ' d'accords existants.                                                     *
    ' Remarque importante : ici on détermine non pas la totalité des marqueurs *
    ' mais les marqueurs des accords joués.                                     *
    ' **************************************************************************
    Public ReadOnly Property Récup_MRQ() As String
        Get
            Dim tbl() As String
            Dim i As Integer
            Dim j As Integer = 0
            Dim a As String = String.Empty
            Dim deb As Integer

            deb = Début.Value
            ReDim Preserve tbl(0)
            If Trim(Grid2.Cell(2, deb).Text) = "" Then
                i = Det_AccordPréced(Début.Value)
                tbl(0) = Trim(Grid2.Cell(1, i).Text)
                deb = Det_AccordSuiv(Début.Value)
                j = j + 1
            End If

            For i = deb To nbMesures '- 1
                If Trim(Grid2.Cell(2, i).Text) <> "" Then ' on détecte uniquement les marqueurs positionné sur des mar
                    If j > UBound(tbl) Then
                        ReDim Preserve tbl(j)
                    End If
                    tbl(j) = (Trim(Grid2.Cell(1, i).Text))
                    j = j + 1
                End If
            Next i
            a = Join(tbl, ",")
            Return a
        End Get
    End Property
    Public ReadOnly Property Récup_AllMRQ() As String
        Get
            Dim tbl() As String
            Dim i As Integer
            Dim j As Integer = 0
            Dim a As String = String.Empty
            Dim deb As Integer

            deb = 1
            ReDim Preserve tbl(0)


            For i = deb To nbMesures '- 1
                If Trim(Grid2.Cell(2, i).Text) <> "" Then ' on détecte uniquement les marqueurs positionné sur des mar
                    If j > UBound(tbl) Then
                        ReDim Preserve tbl(j)
                    End If
                    tbl(j) = (Trim(Grid2.Cell(1, i).Text))
                    j = j + 1
                End If
            Next i
            a = Join(tbl, ",")
            Return a
        End Get
    End Property
    ' ***********************************************************
    ' Récup_Répet : récupération des répététions chaque accord  *
    ' ***********************************************************
    ' 
    Public ReadOnly Property Récup_Répet() As String
        Get
            Dim tbl() As String
            Dim i As Integer
            Dim j As Integer = 0
            Dim a As String = String.Empty
            Dim deb As Integer
            Dim f As Integer = Det_NumDerAccord()

            deb = Début.Value
            ReDim Preserve tbl(0)
            For i = deb To f '- 1
                If Trim(Grid2.Cell(2, i).Text) <> "" And Trim(Grid2.Cell(3, i).Text) <> "" Then ' test présence accord et présence répétition
                    If j > UBound(tbl) Then
                        ReDim Preserve tbl(j)
                    End If
                    tbl(j) = (Trim(Grid2.Cell(3, i).Text))
                    j = j + 1
                End If
            Next i
            a = Join(tbl, ",")
            Return a
        End Get
    End Property

    Public ReadOnly Property Récup_Répet2() As String
        Get
            Dim tbl() As String
            Dim i As Integer
            Dim j As Integer = 0
            Dim k As Integer
            Dim a As String = String.Empty
            Dim deb As Integer
            Dim f As Integer = Terme.Value 'Det_NumDerAccord()

            deb = Début.Value
            ReDim Preserve tbl(0)
            If Trim(Grid2.Cell(2, deb).Text) = "" Then
                tbl(0) = Det_RépetSuiv(deb)
                j = j + 1
                deb = Det_AccordSuiv(Début.Value)
            End If

            For i = deb To f '- 1
                If Trim(Grid2.Cell(2, i).Text) <> "" And Trim(Grid2.Cell(3, i).Text) <> "" Then ' test présence accord et présence répétition
                    If j > UBound(tbl) Then
                        ReDim Preserve tbl(j)
                    End If
                    tbl(j) = (Trim(Grid2.Cell(3, i).Text))
                    j = j + 1
                End If
            Next i
            '
            If Trim(Grid2.Cell(2, i - 1).Text) = "" Then
                tbl(j - 1) = Convert.ToString(Det_RépetPourFin(Terme.Value, Début.Value))
            Else
                k = Det_NumDerAccord()
                If (i - 1) <> k Then
                    tbl(j - 1) = "1"
                End If
            End If
            a = Join(tbl, ",")
            Return a
        End Get
    End Property
    Public ReadOnly Property Récup_AllRépet() As String
        Get
            Dim tbl() As String
            Dim i As Integer
            Dim j As Integer = 0
            Dim a As String = String.Empty
            Dim deb As Integer
            Dim f As Integer = Det_NumDerAccord()

            deb = 1
            ReDim Preserve tbl(0)
            For i = deb To f '- 1
                If Trim(Grid2.Cell(2, i).Text) <> "" And Trim(Grid2.Cell(3, i).Text) <> "" Then ' test présence accord et présence répétition
                    If j > UBound(tbl) Then
                        ReDim Preserve tbl(j)
                    End If
                    tbl(j) = Trim(Grid2.Cell(3, i).Text)
                    j = j + 1
                End If
            Next i
            '


            a = Join(tbl, ",")
            Return a
        End Get
    End Property
    Function Det_RépetSuiv(AccDépart As Integer) As Integer
        Dim i As Integer = AccDépart
        Dim j As Integer = 0
        Det_RépetSuiv = -1 '
        While Trim(Grid2.Cell(2, i).Text) = "" And i <= nbMesures '  While Trim(Grid2.Cell(3, i).Text) = "" And i <= nbMesures
            i = i + 1
            j = j + 1
        End While
        If i >= 1 Then Det_RépetSuiv = j
    End Function
    Function Det_RépetPourFin(Départ As Integer, DebSeq As Integer) As Integer
        Dim i As Integer = Départ
        Dim j As Integer = 1
        Det_RépetPourFin = -1 ' cas où pas d'accord trouvé
        While Trim(Grid2.Cell(2, i).Text) = "" And i <> DebSeq And i >= 1
            i = i - 1
            j = j + 1
        End While
        If i >= 1 Then Det_RépetPourFin = j
    End Function
    '
    ' ********************************************************************************
    ' Récup_NumMagnéto : récupération du N° de magnéto pour chaque accord dans grid2 *
    ' ********************************************************************************
    ' 
    Public ReadOnly Property Récup_NumMagnéto() As String
        Get
            Dim tbl() As String
            Dim i As Integer
            Dim j As Integer = 0
            Dim a As String = String.Empty
            Dim deb As Integer

            deb = Début.Value
            ReDim Preserve tbl(0)
            If Trim(Grid2.Cell(2, deb).Text) = "" Then
                i = Det_AccordPréced(Début.Value)
                tbl(0) = (Convert.ToString(IdentVariation(i)))
                deb = Det_AccordSuiv(Début.Value)
                j = j + 1
            End If

            For i = deb To nbMesures '- 1
                If Trim(Grid2.Cell(2, i).Text) <> "" And Trim(Grid2.Cell(3, i).Text) <> "" Then ' test présence accord et présence répétition
                    If j > UBound(tbl) Then
                        ReDim Preserve tbl(j)
                    End If
                    tbl(j) = (Convert.ToString(IdentVariation(i)))
                    j = j + 1
                End If
            Next i
            a = Join(tbl, ",")
            Return a ' quand aucun accord n'est associé avec une variation à partir de Debut.Value alors a=""
        End Get
    End Property
    Public ReadOnly Property Récup_AllNumMagnéto() As String
        Get
            Dim tbl() As String
            Dim i As Integer
            Dim j As Integer = 0
            Dim a As String = String.Empty
            Dim deb As Integer

            deb = 1
            ReDim Preserve tbl(0)

            For i = deb To nbMesures '- 1
                If Trim(Grid2.Cell(2, i).Text) <> "" And Trim(Grid2.Cell(3, i).Text) <> "" Then ' test présence accord et présence répétition
                    If j > UBound(tbl) Then
                        ReDim Preserve tbl(j)
                    End If
                    tbl(j) = (Convert.ToString(IdentVariation(i)))
                    j = j + 1
                End If
            Next i
            a = Join(tbl, ",")
            Return a ' quand aucun accord n'est associé avec une variation à partir de Debut.Value alors a=""
        End Get
    End Property

    ' **********************************************************************
    ' Récup_Répéter : récupération de la demande de bouclage de lecture    *
    ' du séquenceur                                                        *
    ' **********************************************************************
    Public ReadOnly Property Récup_Répéter() As Boolean
        Get
            Dim b As Boolean = Répéter.Checked
            Return b
        End Get
    End Property
    ' **********************************************************************
    ' Récup_PDébut : récupération du N° du 1er accord devant être joué     *
    ' **********************************************************************
    Public Property PDébut() As Decimal
        Get
            Dim i As Decimal = Début.Value
            Return i
        End Get
        Set(ByVal Value As Decimal)
            Début.Value = Value
        End Set
    End Property
    ' **********************************************************************
    ' Récup_PFIN : récupération du N° du dernierr accord devant être joué  *
    ' **********************************************************************
    Public Property PFin() As Decimal
        Get
            Dim i As Integer = Terme.Value
            Return i
        End Get
        Set(ByVal Value As Decimal)
            Terme.Value = Value
        End Set
    End Property
    ' *******************************************************************************
    ' Récup_ExportCTRL : récupération flag  export des contrôleur pour fichier MIDI *
    ' *******************************************************************************
    Public ReadOnly Property Récup_ExportCTRL() As Boolean
        Get
            Dim a As Boolean = ExportCTRL
            Return a
        End Get
    End Property
    ' **********************************************************************
    ' Récup_Durée : récupération des durées des notes des pistes           *
    ' Les informations sont prévelées dans l'IHM                           *
    ' **********************************************************************
    Public ReadOnly Property Récup_Durée(PisteMidi As Integer, Magnéto As Integer) As String
        Get
            Dim ind As Integer
            ind = Det_IndexGénérateur(PisteMidi, Magnéto) ' générateur = ancien bloc de variation / Magnet) = ancien variation
            Return Trim(PisteDurée.Item(ind).Text)
        End Get
    End Property
    ' **********************************************************************
    ' Récup_Volume : récupération des volumes des pistes                   *
    ' Les informations sont prévelées dans l'IHM                           *
    ' **********************************************************************
    Public ReadOnly Property Récup_Volume(PisteMidi As Integer) As String
        Get
            Return Convert.ToString(Me.Mix.PisteVolume.Item(PisteMidi).Value)
        End Get
    End Property
    ' ***********************************************************************
    ' Récup_VolumeActif : récupération de l'état actif (enabled) des pistes *
    '                     sert aus système de Mute de la table de mixage    *
    ' Les informations sont prévelées dans l'IHM                            *
    ' ***********************************************************************
    Public ReadOnly Property Récup_VolumeActif(PisteMidi As Integer) As Boolean
        Get
            Return Convert.ToString(Me.Mix.muteVolume.Item(PisteMidi).Checked)
        End Get
    End Property
    ' **********************************************************************
    ' Récup_Mute : récupération des Mutes des pistes                       *
    ' Les informations sont prévelées dans l'IHM                           *
    ' **********************************************************************
    Public ReadOnly Property Récup_Mute(PisteMidi As Integer, Magnéto As Integer) As String
        Get
            Dim ind As Integer
            ind = Det_IndexGénérateur(PisteMidi, Magnéto) ' générateur = ancien bloc de variation / Magnet) = ancien variation
            Return Convert.ToString(SelBloc.Item(ind).Checked)
        End Get
    End Property
    ' **********************************************************************
    ' Récup_Delay : récupération des Delay des pistes                       *
    ' Les informations sont prévelées dans l'IHM                           *
    ' **********************************************************************
    Public ReadOnly Property Récup_Delay(PisteMidi As Integer, Magnéto As Integer) As String
        Get
            Dim ind As Integer
            ind = Det_IndexGénérateur(PisteMidi, Magnéto) ' générateur = ancien bloc de variation / Magnet) = ancien variation
            Return Convert.ToString(PisteDelay.Item(ind).Checked)
        End Get
    End Property
    ' **********************************************************************
    ' Récup_Début : récupération des Débutle lecture des chaînes de notes  *
    ' Les informations sont prévelées dans l'IHM                           *
    ' **********************************************************************
    Public ReadOnly Property Récup_DébutSouche(PisteMidi As Integer, Magnéto As Integer) As String
        Get
            Dim ind As Integer
            ind = Det_IndexGénérateur(PisteMidi, Magnéto)
            Return Convert.ToString(PisteDébut.Item(ind).Checked)
        End Get
    End Property
    ' **********************************************************************
    ' Récup_Accent : récupération de la valeur d'un Accent d'un Motif      *
    ' **********************************************************************
    Public ReadOnly Property Récup_Accent(PisteMidi As Integer, Magnéto As Integer) As String
        Get
            Dim ind As Integer
            ind = Det_IndexGénérateur(PisteMidi, Magnéto) ' générateur = ancien bloc de variation / Magnet) = ancien variation
            Return Trim(PisteAccent.Item(ind).Text)
        End Get
    End Property
    ' **********************************************************************
    ' Récup_Octave : récupération des Octaves des pistes                   *
    ' Les informations sont prévelées dans l'IHM                           *
    ' **********************************************************************
    Public ReadOnly Property Récup_Octave(PisteMidi As Integer, Magnéto As Integer) As String
        Get
            Dim ind As Integer
            ind = Det_IndexGénérateur(PisteMidi, Magnéto) ' générateur = ancien bloc de variation / Magnet) = ancien variation
            Return Trim(PisteOctave.Item(ind).Text)
        End Get
    End Property
    ' **********************************************************************
    ' Récup_PRG : récupération des N° de programmes des pistes             *
    ' Les informations sont prévelées dans l'IHM                           *
    ' **********************************************************************
    Public ReadOnly Property Récup_PRG(PisteMidi As Integer, Magnéto As Integer) As String
        Get
            Dim ind As Integer
            ind = Det_IndexGénérateur(PisteMidi, Magnéto) ' générateur = ancien bloc de variation / Magnet) = ancien variation
            Return Convert.ToString(PistePRG.Item(ind).SelectedIndex - 1)
        End Get
    End Property
    ' **********************************************************************
    ' Récup_Dyn : récupération des dynamiques des notes des pistes         *
    ' Les informations sont prévelées dans l'IHM - Cas ancien où les       *
    ' dynamiques étaient proposées dans un combobox                        * 
    ' **********************************************************************
    Public ReadOnly Property Récup_Dyn(PisteMidi As Integer, Magnéto As Integer) As String
        Get
            Dim ind As Integer
            ind = Det_IndexGénérateur(PisteMidi, Magnéto) ' générateur = ancien bloc de variation / Magnet) = ancien variation
            Return Trim((PisteDyn.Item(ind).Text))
        End Get
    End Property
    ' **********************************************************************
    ' Récup_Dyn2 : récupération des dynamiques des notes des pistes        *
    ' Les informations sont prévelées dans l'IHM - Les dynamiques sont     *
    ' récupérées dans un NumeriUpDown                                      *
    ' **********************************************************************
    Public ReadOnly Property Récup_Dyn2(PisteMidi As Integer, Magnéto As Integer) As String
        Get
            Dim ind As Integer
            ind = Det_IndexGénérateur(PisteMidi, Magnéto) ' générateur = ancien bloc de variation / Magnet) = ancien variation
            Return Convert.ToString((PisteDyn2.Item(ind).Value))
        End Get
    End Property
    ' **********************************************************************
    ' Récup_PAN : récupération des valeurs de panoramiques de chaque piste *
    ' **********************************************************************
    Public ReadOnly Property Récup_PAN(PisteMidi As Integer, Magnéto As Integer) As String
        Get

            Dim ind As Integer
            ind = Det_IndexGénérateur(PisteMidi, Magnéto) ' générateur = ancien bloc de variation / Magnet) = ancien variation

            j = Det_PAN(ind)
            If j = 1 Then
                Return "0"      ' gauche
            ElseIf j = 2 Then
                Return "64"     ' centre
            ElseIf j = 3 Then   ' droite
                Return "127"
            Else
                Return "0"
            End If

        End Get
    End Property
    ' **********************************************************************
    ' Récup_Souche : récupération des début de lecture du motif            *
    ' Les informations sont prévelées dans l'IHM                           *
    ' **********************************************************************
    Public ReadOnly Property Récup_Souche(PisteMidi As Integer, Magnéto As Integer) As String
        Get
            Dim ind As Integer
            ind = Det_IndexGénérateur(PisteMidi, Magnéto) ' générateur = ancien bloc de variation / Magnet) = ancien variation
            Return Trim(PisteSouche.Item(ind).Text)
        End Get
    End Property
    ' **********************************************************************
    ' Récup_Retard : récupération du retard de démarrage de la piste       *
    ' Les informations sont prévelées dans l'IHM                           *
    ' **********************************************************************
    Public ReadOnly Property Récup_Retard(PisteMidi As Integer, Magnéto As Integer) As String
        Get
            Dim ind As Integer
            ind = Det_IndexGénérateur(PisteMidi, Magnéto) ' générateur = ancien bloc de variation / Magnet) = ancien variation
            If Trim(PisteRetard.Item(ind).Text) = "off" Then
                Return "0"
            Else
                Return Trim(PisteRetard.Item(ind).Text)
            End If
        End Get
    End Property

    ' **********************************************************************
    ' Récup_Motif : récupération des motifs (Arpège, Répétition, Perso) de *
    ' jeu des notes des pistes les informations sont prévelées dans l'IHM  *                         *
    ' **********************************************************************
    Public ReadOnly Property Récup_Motif(PisteMidi As Integer, Magnéto As Integer) As String

        Get
            Dim ind As Integer
            Dim a As String
            ind = Det_IndexGénérateur(PisteMidi, Magnéto) ' générateur = ancien bloc de variation / Magnet) = ancien variation
            'a = Trim(PisteMotif.Item(ind).Text)
            'If a.Chars(0) = "-" Then
            'a = Trim(PisteMotif.Item(ind + 1).Text)
            'Else
            ' a = Trim(PisteMotif.Item(ind).Text)
            'End If

            a = BoutMotif.Item(ind).Text
            Return Trim(a)
        End Get
    End Property
    ' **********************************************************************
    ' Récup_MotifsPerso : création de la chaine de notes d'un motif Perso  *                         *
    ' **********************************************************************
    Public ReadOnly Property Récup_MotifPerso(numGrid As Integer, Dico As Dictionary(Of String, String), nbMesures As Integer, Début As Integer) As String
        Get
            Dim i, j, k As Integer
            Dim jj As Integer
            Dim b, signe, n, s As String
            Dim a As String = ""
            Dim cmptLong As Integer = 0
            Dim ListNotes As New List(Of ChainePerso)
            'Dim ListSilences As New List(Of ChainePerso)
            Dim ListChaines As New List(Of String)
            Dim ListFinal As New List(Of String)
            Dim ListFinal2 As New List(Of String)
            Dim Tab1(15) As String
            Dim PosSIL As Integer
            Dim LongSIL As Integer
            Dim TBL1() As String
            Dim flagPrésenceNotes As Boolean = False

            ' pour info
            ' Class ChainePerso
            '   Public Evt As String
            '   Public Row As Integer
            '   Public Col As Integer
            '   Public LettreLigne As String
            '   Public Longueur As Integer
            '   Public Chaine As String
            ' End Class

            ' Identification des notes
            ' ************************
            k = -1
            For j = 1 To MotifsPerso.Item(numGrid).Cols - 1
                For i = 2 To MotifsPerso.Item(numGrid).Rows - 1
                    signe = Trim(MotifsPerso.Item(numGrid).Cell(i, j).Text) 'lecture cellule du tableau Perso
                    If Trim(signe) <> "" Then ' If Signe ="*"
                        '
                        ' 1°- Déterminationn de la colonne dans la grille Perso
                        ' *************************************************
                        flagPrésenceNotes = True
                        ListNotes.Add(New ChainePerso)
                        k = k + 1
                        ListNotes.Item(k).Col = j ' Colonne
                        ListNotes.Item(k).Row = i
                        ListNotes.Item(k).Evt = a
                        '
                        ' 2° - Détermination de la  longueur de la note
                        ' *********************************************
                        cmptLong = 0
                        For jj = j To MotifsPerso.Item(numGrid).Cols - 1
                            If (MotifsPerso.Item(numGrid).Cell(i, jj).BackColor = Color.White) Or (MotifsPerso.Item(numGrid).Cell(i, jj).BackColor = Couleur_Div) Or (Trim(MotifsPerso.Item(numGrid).Cell(i, jj).Text) <> "" And jj <> j) Then 'If MotifsPerso.Item(numGrid).Cell(i, jj).BackColor <> Couleur_Notes Then
                                Exit For
                            End If
                            cmptLong = cmptLong + 1
                        Next jj
                        ListNotes.Item(k).Longueur = cmptLong
                        ' 3° - Détermination de la note
                        ' *****************************
                        b = Det_LettreLigne(i)
                        ListNotes.Item(k).LettreLigne = Trim(b)
                        n = Dico(b)
                        n = CalcNoteOtave(n, signe)
                        ' 4° - Construction de la chaine
                        ' ******************************
                        ListNotes.Item(k).Chaine = n + "-" + Convert.ToString(cmptLong)
                        '
                        'Exit For

                    End If
                Next i
            Next j
            '
            ' Mise en place des chaines 'notes' dans tableau de string (Tab(16))
            ' ******************************************************************
            If flagPrésenceNotes = True Then

                For Each xEVT As ChainePerso In ListNotes
                    j = xEVT.Col
                    Tab1(j - 1) = Trim(xEVT.Chaine) ' la chaine contient la note suivie de la longueur
                Next

                ' Détermination des chaines 'silences' à partir de ListNotes
                ' **********************************************************
                ' Position du silence = (COLpres + LONGUEURpres) +1
                ' Longueur du silence = (COLsuiv - (COLpres + LONGUEURpres)) si Longueur du silence =0 alors pas de Silence
                If SilenceAuDébut(numGrid) Then
                    PosSIL = 1
                    LongSIL = ListNotes.Item(0).Col - 1
                    s = "SIL-" + Convert.ToString(LongSIL)
                    Tab1(PosSIL - 1) = Trim(s)
                End If
                '
                For k = 0 To ListNotes.Count - 1

                    PosSIL = ((ListNotes.Item(k).Col) + (ListNotes.Item(k).Longueur))
                    If k < ListNotes.Count - 1 Then
                        LongSIL = (ListNotes.Item(k + 1).Col) - ((ListNotes.Item(k).Col) + (ListNotes.Item(k).Longueur))
                    Else
                        LongSIL = (17) - ((ListNotes.Item(k).Col) + (ListNotes.Item(k).Longueur))
                    End If
                    ' Mise à jour dans Tab1
                    If LongSIL <> 0 Then
                        s = "SIL-" + Convert.ToString(LongSIL)
                        Tab1(PosSIL - 1) = Trim(s)
                    End If
                Next
                '
                ' Création de la chaine globale
                ' ***************************** 
                For i = 0 To nbMesures - 1
                    For j = 0 To UBound(Tab1)
                        If Tab1(j) <> Nothing Then
                            ListFinal.Add(Tab1(j))
                        End If
                    Next j
                Next i

                a = String.Join(" ", ListFinal.ToArray())


                ' Prise en compte de la souche (Début)
                ' ************************************
                TBL1 = Split(a)
                Dim TBL2(UBound(TBL1)) As String
                If Début = 1 Then
                    Array.Copy(TBL1, 1, TBL2, 0, TBL1.Length - 1)
                    TBL2(UBound(TBL1)) = TBL1(0)
                Else
                    Array.Copy(TBL1, 0, TBL2, 0, TBL1.Length) ' pas de souche
                End If
                '
                ' Chaine finale
                ' *************
                a = String.Join(" ", TBL2)
            End If
            Return Trim(a)

        End Get
    End Property
    Public ReadOnly Property Récup_MotifPerso2(numGrid As Integer, Dico As Dictionary(Of String, String), nbMesures As Integer, Début As Integer) As String
        Get
            Dim i, j, k As Integer
            Dim ii, jj As Integer
            Dim b, signe, n As String
            Dim a As String = ""
            Dim cmptLong As Integer = 0
            Dim ListNotes As New List(Of ChainePerso)
            'Dim ListChaines As New List(Of String)
            'Dim ListFinal As New List(Of String)
            'Dim ListFinal2 As New List(Of String)
            Dim LNotesDéjàPres As New List(Of String)
            Dim Tab1(15) As String
            Dim flagPrésenceNotes As Boolean = False
            Dim c As String
            ' Identification des notes
            ' ************************
            k = -1
            For ii = 0 To nbMesures - 1
                For j = 1 To MotifsPerso.Item(numGrid).Cols - 1
                    For i = 2 To MotifsPerso.Item(numGrid).Rows - 1
                        signe = Trim(MotifsPerso.Item(numGrid).Cell(i, j).Text) 'lecture cellule du tableau Perso
                        If Trim(signe) <> "" Then ' If Signe ="*"
                            '
                            ' 1°- Déterminationn de la colonne dans la grille Perso
                            ' *****************************************************
                            ListNotes.Add(New ChainePerso)
                            k = k + 1
                            ListNotes.Item(k).Col = (j - 1) + (ii * 16) ' Colonne - quand ii> 1 la colonne e
                            ListNotes.Item(k).Row = i - 1
                            '
                            ' 2° - Détermination de la  longueur de la note
                            ' *********************************************
                            cmptLong = 0
                            For jj = j To MotifsPerso.Item(numGrid).Cols - 1
                                If (MotifsPerso.Item(numGrid).Cell(i, jj).BackColor = Color.White) Or (MotifsPerso.Item(numGrid).Cell(i, jj).BackColor = Couleur_Div) Or (Trim(MotifsPerso.Item(numGrid).Cell(i, jj).Text) <> "" And jj <> j) Then 'If MotifsPerso.Item(numGrid).Cell(i, jj).BackColor <> Couleur_Notes Then
                                    Exit For
                                End If
                                cmptLong = cmptLong + 1
                            Next jj
                            ListNotes.Item(k).Longueur = cmptLong
                            '
                            ' 3° - Détermination de la note
                            ' *****************************
                            b = Det_LettreLigne(i)
                            ListNotes.Item(k).LettreLigne = Trim(b)
                            n = Dico(b)
                            n = CalcNoteOtave(n, signe)

                            ' 4° - Construction de la chaine
                            ' ******************************
                            ' Note (en lettre)   Colonne de début de la note           Longueur de la note   

                            c = n + " " + ListNotes.Item(k).Col.ToString + " " + ii.ToString ' on repère la note avec sa valeur et sa colonne
                            If LNotesDéjàPres.IndexOf(Trim(c)) = -1 Then ' on vérifie que la note n'est pas déjà présente (cas rare mais arrive qd l'utilisateur écrit  une 4e ou 5e voix alors que l'accord joué est un accord de 3 notes
                                ListNotes.Item(k).Chaine = Trim(n) + "-" + Trim(ListNotes.Item(k).Col.ToString) + "-" + Trim(ListNotes.Item(k).Longueur.ToString)
                                LNotesDéjàPres.Add(Trim(c)) ' ajout de la note (+ colonne) dans la liste des notes déjà présentes.
                            End If
                            '
                        End If
                    Next i
                Next j
            Next ii
            ' For ii = 0 To nbMesures - 1
            For Each note As ChainePerso In ListNotes
                If Trim(note.Chaine) <> "" Then
                    a = a + note.Chaine + " "
                End If
            Next

            Return Trim(a)
        End Get
    End Property
    Function CalcNoteOtave(n As String, signe As String) As String
        Dim i, j As Integer
        Dim nn As String = n

        i = Module1.ValNoteCubase.IndexOf(Trim(n))
        j = i
        If Trim(signe) = "+" Then
            i = i + 12
            If i > 127 Then i = j ' si on dépasse, alors on revient à la note intitiale
            nn = ValNoteCubase.Item(i)
        ElseIf Trim(signe) = "-" Then
            i = i - 12
            If i < 0 Then i = j ' si on est en dessous du min, alors on revient à la note intitiale
            nn = ValNoteCubase.Item(i)
        End If
        Return nn
    End Function
    Function Det_LettreLigne(Ligne) As String

        Select Case Ligne
            Case 2
                Det_LettreLigne = "A"
            Case 3
                Det_LettreLigne = "B"
            Case 4
                Det_LettreLigne = "C"
            Case 5
                Det_LettreLigne = "D"
            Case 6
                Det_LettreLigne = "E"
            Case Else
                Det_LettreLigne = "A"
        End Select
    End Function
    Function SilenceAuDébut(numGrid As Integer) As Boolean
        Dim i As Integer
        SilenceAuDébut = True
        For i = 2 To MotifsPerso.Item(numGrid).Rows - 1
            If Trim(MotifsPerso.Item(numGrid).Cell(i, 1).Text) <> "" Then
                SilenceAuDébut = False
            End If
        Next
    End Function


    ' *#######################################################################
    ' ##                                                                    ##
    ' ##                         EVENEMENTS                                 ##
    ' ##                                                                    ##
    ' ########################################################################

    Private Sub Début_ValueChanged(sender As Object, e As EventArgs) Handles Début.ValueChanged
        'If Not EnChargement Then
        ' If Début.Value > Det_NumDerAccord() Then
        ' Début.Value = Début.Value - 1
        'End If
        ' If Début.Value >= Terme.Value Then
        'If Terme.Value <> 0 Then
        'Début.Value = Terme.Value '- 1
        'End If
        'End If
        'End If
        Cohérence_Délim()
        Calcul_Durée()
        TextBox2.Text = Calcul_Durée()
    End Sub
    Private Sub Début_TextChanged(sender As Object, e As EventArgs) Handles Début.TextChanged
        If Module1.IsLoaded(Transport) Then
            Transport.Label2.Text = Début.Value.ToString
            Transport.TextBox1.Text = Début.Value.ToString
        End If
    End Sub
    Private Sub Terme_TextChanged(sender As Object, e As EventArgs) Handles Terme.TextChanged
        If Module1.IsLoaded(Transport) Then
            Transport.Label3.Text = Terme.Value.ToString
            Transport.TextBox2.Text = Terme.Value.ToString
        End If
    End Sub
    'Private Sub Terme_TextChanged(sender As Object, e As EventArgs) Handles Début.TextChanged

    'End Sub
    Private Sub Fin_ValueChanged(sender As Object, e As EventArgs)
        If Début.Value >= Terme.Value Then
            If Terme.Value <> 0 Then
                Terme.Value = Terme.Value
            End If
        End If
    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs)
        ZoomMoins()
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs)
        ZoomPlus()
    End Sub

    Private Sub Button32_Click_1(sender As Object, e As EventArgs) Handles Button32.Click
        'If Module1.LangueIHM = "fr" Then
        'a = "https://www.hyperarp.fr/guide-rapide"
        'Else
        'a = "https://www.hyperarp.fr/guide-rapide?lang=en"
        'End If
        'Process.Start(Trim(a))
        Process.Start("https://compomusic.fr")
    End Sub


    Private Sub Button33_Click(sender As Object, e As EventArgs)
        Début.Text = TZone(ZoneCourante).DébutZ
        Terme.Text = TZone(ZoneCourante).TermeZ
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim i As Integer = Grid2.LeftCol

        ' Effacer les saisies de marqueur ou répétition du dernier accords
        ' *****************************************************************
        TMarqueur.Visible = False
        TMarqueur2.Visible = False

        AssurerVisibilitéCelluleGrid2(Grid2.ActiveCell.Col)
        ZoomPlusGrid2()
        Grid2.LeftCol = i
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim i As Integer = Grid2.LeftCol
        ' Effacer les saisies de marqueur ou répétition du dernier accords
        ' *****************************************************************
        TMarqueur.Visible = False
        TMarqueur2.Visible = False
        '
        AssurerVisibilitéCelluleGrid2(Grid2.ActiveCell.Col)
        ZoomMoinsGrid2()
        Grid2.LeftCol = i
    End Sub

    Private Sub Button31_Click(sender As Object, e As EventArgs)
        'Création_Schedule(CréationFichierScheduleGamme("FichierGamme", "C Maj", "C", 1))
    End Sub
    Private Sub ListBox1_MouseUp(sender As Object, e As MouseEventArgs)
        ' Couper Jouer Accord
        ' *******************
        If AccordAEtéJoué = True Then
            CouperJouerAccord()
            AccordAEtéJoué = False
        End If
    End Sub

    Private Sub Grid5_KeyDown(Sender As Object, e As KeyEventArgs)

    End Sub


    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If Not (FermetureParQuitter) Then
            Sortir(e)
        End If
    End Sub


    Private Sub Grid2_MouseDown(Sender As Object, e As MouseEventArgs) Handles Grid2.MouseDown
        Dim i As Integer
        Dim j As Integer
        Dim m As Integer
        Dim a As String
        Dim b As String
        Dim tbl() As String
        Dim oDiff As Integer

        i = Grid2.MouseRow
        j = Grid2.MouseCol

        ' Choix du mode de sélection en fonction de la ligne
        ' **************************************************

        If i < 12 Then
            Grid2.SelectionMode = SelectionModeEnum.ByColumn
            Grid2.BackColorSel = Color.Lavender
            Grid2.BackColorActiveCellSel = Color.Transparent
        Else
            Grid2.SelectionMode = SelectionModeEnum.Free
            Grid2.BackColorSel = Color.Yellow
            Grid2.BackColorActiveCellSel = Color.Yellow
        End If

        'EcritureMarqueur() ' possibilité de valider l'écriture d'un marqueur en cliquant sur Grid2


        OngletCours_HyperARP = 127 ' pour copier/coller : permet d'indiquer que l'on se trouve dans HyperArp car on a cliqué dans l'un des accords.
        If Grid2.MouseCol <= nbMesures And Grid2.Selection.LastCol <= nbMesures And Grid2.MouseCol > 0 Then
            '

            '
            If i > -1 And j > -1 Then ' cas où MouseRow ou MouseCol retournent une erreur =-1, on ne fait rien
                MesureCourante = j ' utilisé par JouerSourcexxxxx
                '
                ' Mise à jour de la tonalité en cours dans les combobox1,2  en fonction de la mesure sélectionnée
                ' ***********************************************************************************************
                If Not My.Computer.Keyboard.CtrlKeyDown Then
                    m = Det_AccordPréced(MesureCourante)
                    a = Trim(TableEventH(m, 1, 1).Tonalité)
                    ComboBox1.SelectedIndex = Det_TonalitéIndex(a)
                    '
                End If
                '
                ' détermination de la clef courante
                ' *********************************
                m = j
                b = TableEventH(m, 1, 1).Tonalité
                tbl = Split(b)
                Clef = Det_Clef(tbl(0))
                '
                ' Jouer Accord
                ' ************
                If e.Button() = Windows.Forms.MouseButtons.Left And My.Computer.Keyboard.CtrlKeyDown And i = 2 Then 'And Trim(Grid2.Cell(i, j).Text <> ""
                    If ComboMidiOut.Items.Count > 0 Then
                        If EnChargement = False Then
                            If AccordAEtéJoué = False Then
                                If Trim(Grid2.Cell(i, m).Text) <> "" Then
                                    JouerAccord(Trim(Grid2.Cell(i, m).Text)) ' ici on a toujours i=2
                                End If
                            End If
                        End If
                    End If
                    '
                End If
                ' Mise à jour Tables pour Onglet Substitution
                ' *******************************************
                m = j
                If e.Button() = Windows.Forms.MouseButtons.Left And Trim(TableEventH(m, 1, 1).Accord) <> "" And
                     Not ((My.Computer.Keyboard.CtrlKeyDown) And (My.Computer.Keyboard.AltKeyDown)) Then ' TabControl4.SelectedIndex = 2 --> onglet Substitution

                    Dim TMaj, TMin As String

                    tbl = Split(Trim(TableEventH(m, 1, 1).Tonalité))
                    TMaj = Trim(tbl(0))
                    tbl = Det_RelativeMineure(Trim(TableEventH(m, 1, 1).Tonalité)).Split
                    TMin = tbl(0)
                    Maj_TableGlobalAccSubsti(TMaj, TMin, TMaj)
                    '
                    Maj_Substitutions(m, 1, 1)
                End If


                ' Entrées des marqueurs et des répétitions (sur le dernier accord)
                ' ****************************************************************
                '
                ' a/ entrée des marqueurs
                ' ***********************
                TMarqueur.Visible = False
                TMarqueur2.Visible = False
                '
                If e.Button() = MouseButtons.Left Then
                    Dim taille As New Size
                    Dim LigneVisible As Integer = Grid2.TopRow
                    Dim P1, P2 As New Point
                    '

                    If My.Computer.Keyboard.CtrlKeyDown Then
                        If i = 1 Then '1 = N° de ligne des marqueurs
                            '
                            TMarqueur.Visible = True
                            TMarqueur.BringToFront()
                            TMarqueur.TextAlign = HorizontalAlignment.Center
                            ' écriture du texte actuel dans la textBox
                            TMarqueur.Focus()
                            TMarqueur.Text = Trim(Grid2.Cell(i, j).Text)
                            TMarqueur.SelectionStart = Len(TMarqueur.Text)
                            '
                            taille.Width = Grid2.Column(j).Width
                            taille.Height = Grid2.Row(i).Height '- 2
                            TMarqueur.Size = taille
                            '
                            oDiff = j - Grid2.LeftCol
                            '
                            Dim ii = Me.TabControl2.Size.Height
                            P1.Y = (SplitContainer7.Panel1.Location.Y + 298) + Grid2.Row(0).Height + 1 '+ Grid2.Row(3).Height ' H * (i - (LigneVisible - 1)) + 1 '
                            P1.X = (Grid2.Column(j).Width * oDiff) + Grid2.Column(0).Width + 1
                            '
                            TMarqueur.Location = P1
                            TMarqueur.Refresh()
                            If TMarqueur.CanFocus Then
                                TMarqueur.Focus() ' avoir le focus
                            End If

                        End If
                    Else
                        TMarqueur.Visible = False
                    End If
                    '
                    '  b/ Entrée des Répétitions sur le dernier accord
                    '  ***********************************************
                    If My.Computer.Keyboard.CtrlKeyDown Then
                        If i = 3 And j = Det_NumDerAccord() Then ' 
                            TMarqueur2.Visible = True
                            TMarqueur2.TextAlign = HorizontalAlignment.Center
                            '
                            TMarqueur2.Focus()
                            TMarqueur2.Text = Trim(Grid2.Cell(i, j).Text)
                            TMarqueur2.SelectionStart = Len(TMarqueur2.Text)

                            taille.Width = Grid2.Column(j).Width
                            taille.Height = Grid2.Row(i).Height + 5
                            TMarqueur2.Size = taille
                            '
                            oDiff = j - Grid2.LeftCol
                            P1.Y = (SplitContainer7.Panel1.Location.Y + 298) + Grid2.Row(0).Height + Grid2.Row(1).Height + Grid2.Row(2).Height + 3 '+ Grid2.Row(3).Height ' H * (i - (LigneVisible - 1)) + 1 '
                            P1.X = (Grid2.Column(j).Width * oDiff) + Grid2.Column(0).Width + 1
                            '
                            TMarqueur2.Location = P1
                            TMarqueur2.Refresh()
                            If TMarqueur2.CanFocus Then
                                TMarqueur2.Focus() ' avoir le focus
                            End If
                            '
                        End If ' *
                    Else
                        TMarqueur2.Visible = False
                    End If
                    '
                End If

                DerGridCliquée = GridCours.Grid2

            End If
        End If
        '
        ' Onglet Modulation
        ' *****************
        Maj_ModulationRadioB2(j)
        LabModulat.Item(0).Text = "---"
        LabModulat.Item(1).Text = "---"
        LabModulat.Item(2).Text = "---"
        LabModulat.Item(3).Text = "---"
    End Sub
    Private Sub Grid2_MouseUp(Sender As Object, e As MouseEventArgs) Handles Grid2.MouseUp
        Dim i, j, k As Integer

        Dim m As Integer
        Dim t As Integer
        Dim ct As Integer
        Dim tbl() As String
        Dim Firstc As Integer



        If Grid2.MouseCol <= nbMesures And Grid2.Selection.LastCol <= nbMesures Then
            If Grid2.Selection.LastCol <= nbMesures Then
                '
                ' Choix de la Variation  par ctrl+clic souris
                ' ********************************************
                If e.Button() = Windows.Forms.MouseButtons.Left And My.Computer.Keyboard.CtrlKeyDown Then
                    If Grid2.ActiveCell.Row >= 4 And Grid2.ActiveCell.Row <= 10 Then
                        Firstc = Det_AccordPréced(Grid2.Selection.FirstCol)
                        'ChoixMagnetoGrid2_2(Grid2.Selection.FirstCol, Grid2.ActiveCell.Row) ' maj couleur magnéto et du N° magnéto dans TableEventh
                        ChoixVariationGrid2_3(Firstc, Grid2.ActiveCell.Row) ' maj couleur de la variation et du N° magnéto dans TableEventh
                        '
                        ' Mise à jour PianoRoll
                        ' *********************
                        'a = Trim(Det_ListAcc())
                        'b = Trim(Det_ListGam())
                        'c = Trim(Det_ListMarq())
                        'd = Trim(Det_ListTon())
                        'For i = 0 To nb_PianoRoll - 1
                        'If PIANOROLLChargé(i) = True Then
                        'listPIANOROLL(i).PListAcc = a 'Det_ListAcc()
                        'listPIANOROLL(i).PListGam = b 'Det_ListGam()
                        'listPIANOROLL(i).PListMarq = c 'Det_ListMarq()
                        'listPIANOROLL(i).PListTon = d
                        'listPIANOROLL(i).F1_Refresh()
                        'listPIANOROLL(i).Maj_CalquesMIDI()
                        'End If
                        'Next
                    End If
                    '
                    k = Det_NumDerAccord()
                    '
                Else
                    If Grid2.ActiveCell.BackColor <> Color.White And (Grid2.ActiveCell.Row >= 4 And Grid2.ActiveCell.Row < 11) Then
                        'TabControl5.SelectTab(Grid2.ActiveCell.Row - 4) ' affichage de la variation
                    End If
                End If

                ' Arrêter jeu Accord
                ' ******************
                'SortieMidi.Item(ChoixSortieMidi).SilenceAllNotes()
                '
                If AccordAEtéJoué = True Then
                    CouperJouerAccord2()
                    AccordAEtéJoué = False
                End If
                '
                RAZ_AffNoteAcc() ' RAZ des notes afficher sans jouer les notes (utilisation du bouton médian de la souris)
                '
                i = Grid2.ActiveCell.Row
                j = Grid2.ActiveCell.Col
                '
                m = Val(Grid2.ActiveCell.Col)
                t = 1
                ct = 1
                '
                ' Mettre à jour la Tonalité dans Indicateur de la barre d'outil(complètementà gauche)
                ' ***********************************************************************************
                tbl = Split(Trim(TableEventH(m, t, ct).Tonalité), " ")
                If Trim(Grid2.Cell(i, j).Text) <> "" And Trim(tbl(0)) <> "" Then ' cas où on clique sur une celluke vide
                    AffTonaCours(Trim(tbl(0)))
                End If
                '
                '
                ' Afficher le menu contextuel du réglage des chiffrages d'accords complexes
                ' *************************************************************************
                If (e.Button() = Windows.Forms.MouseButtons.Right) Then
                    m = Grid2.MouseCol
                    SauveMouseColGrid2 = m

                    ContextMenu3Accord.Visible = False
                    DerGridCliquée = GridCours.Grid2
                    Transpo.Visible = True ' afficher le menu de transposition
                    ContextMenuStrip3.Show(CType(Sender, Object), e.Location)
                End If
            End If
        End If
    End Sub
    Private Sub SplitContainer2_SplitterMoving(sender As Object, e As SplitterCancelEventArgs)
        'Label47.Text = Str(SplitContainer2.SplitterDistance)
    End Sub
    Private Sub Form1_Paint(sender As Object, e As PaintEventArgs) Handles MyBase.Paint
        'LabelZoneCourante.Text = Str(Me.Width)

    End Sub

    Private Sub ComboBox1_MouseDown(sender As Object, e As MouseEventArgs)
        OrigineTona = "Maj"
    End Sub

    Sub Maj_TabTonTabCad_ChangeLangue()

        ' Public TableCoursAcc(0 To 2, 0 To 6) As AccordTab ' ligne , colonne
        ' Public CAD_TableCoursAcc(0 To 6) As AccordTab
        ' Calcul des n° de cellule d'un tableau : (N° colonne) + (nbLignes x N° colonne)
        Dim i, j, k As Integer
        Const nbCol = 7

        ' maj Tabton
        ' **********
        For i = 0 To 2
            For j = 0 To 6
                k = j + (nbCol * i)
                TabTons.Item(k).Text = TradEventHLat(TableCoursAcc(i, j).Accord)
            Next
        Next
        '
        Refresh()
    End Sub
    ' *#######################################################################
    ' ##                                                                    ##
    ' ##                         CULTURE                                    ##
    ' ##                                                                    ##
    ' ########################################################################
    Function ListeNotesMin_EnD(liste_ As String, délimit As String) As String
        Dim i As Integer
        Dim tbl() As String


        tbl = Split(liste_, délimit)
        '
        For i = 0 To UBound(tbl)
            Select Case tbl(i)
                Case "dd"
                    tbl(i) = "c#"
                Case "eb"
                    tbl(i) = "d#"
                Case "gb"
                    tbl(i) = "f#"
                Case "ab"
                    tbl(i) = "g#"
                Case "bb"
                    tbl(i) = "a#"
                Case Else
                    tbl(i) = tbl(i)
            End Select
        Next
        ListeNotesMin_EnD = Join(tbl, délimit)
    End Function
    Function Trad_AccordEn_D(Accord As String) As String
        Dim tblA() As String

        tblA = Split(Accord)
        tblA(0) = Trim(Trad_ListeNotesEnD(LCase(tblA(0)), " "))
        Trad_AccordEn_D = Trim(tblA(0))
        If UBound(tblA) > 0 Then
            Trad_AccordEn_D = UCase(tblA(0)) + " " + tblA(1)
        Else
            Trad_AccordEn_D = UCase(tblA(0))
        End If

    End Function
    Function Trad_GammeEn_D(Gamme As String) As String
        Dim tblG() As String

        tblG = Split(Gamme)
        tblG(0) = Trim(Trad_ListeNotesEnD(LCase(tblG(0)), "-"))
        Trad_GammeEn_D = Join(tblG, " ")

    End Function

    Sub ChoixLangueEn()
        '
        Langue = "en"
        LanguesToolStripMenuItem.Text = "Langue"
        EnglishToolStripMenuItem.Checked = True
        FrançaisToolStripMenuItem.Checked = False
        '
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CalquesMIDI\HyperArp\Langue", "Langue", "en")
        En_Culture()
        '
        ' Procédures de traduction Anglais --> Latin
        ' ******************************************

        '
    End Sub
    Sub ChoixLangueFr()
        'Try

        Langue = "fr"
        LanguesToolStripMenuItem.Text = "Language"
        EnglishToolStripMenuItem.Checked = False
        FrançaisToolStripMenuItem.Checked = True
        '
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CalquesMIDI\HyperArp\Langue", "Langue", "fr")
        Fr_Culture()
        'Catch ex As Exception
        'MessageBox.Show(ex.Message + " ChoixLangueFr")
        'End Try
    End Sub
    Private Sub FrançaisToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FrançaisToolStripMenuItem.Click
        'If LangueIHM = "en" Then
        Cacher_FormTransparents()
        LangueIHM = "fr"
        MessageHV.PTitre = "Changement de langue"
        MessageHV.PContenuMess = "Vous devez redémarrer l'application pour valider les modifications de langue."
        MessageHV.PTypBouton = "OK"
        Cacher_FormTransparents()
        MessageHV.ShowDialog()
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CalquesMIDI\HyperArp\LangueIHM", "LangueIHM", "fr")
        'End If

        ' 
    End Sub
    Private Sub EnglishToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EnglishToolStripMenuItem.Click
        'If LangueIHM = "fr" Then
        Cacher_FormTransparents()
        LangueIHM = "en"
        MessageHV.PTitre = "Language change"
        MessageHV.PContenuMess = "You must restart the application to validate the language change."
        MessageHV.PTypBouton = "OK"
        Cacher_FormTransparents()
        MessageHV.ShowDialog()
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CalquesMIDI\HyperArp\LangueIHM", "LangueIHM", "en")
        'End If
        '
    End Sub

    ' *#######################################################################
    ' ##                                                                    ##
    ' ##                         UTILITAIRES                                ##
    ' ##                                                                    ##
    ' ########################################################################

    Function IsMultiple(ByVal W_Chiffre As Double, ByVal W_Multiple As Double) As Boolean
        IsMultiple = ((W_Chiffre Mod W_Multiple) = 0)
    End Function


    Function Det_MaxVolume() As Integer
        Dim i, v As Integer
        Dim MaxV As Integer
        '
        Dim List1 As New List(Of Integer)

        For i = 0 To nb_TotalPistes - 1 'Arrangement1.Nb_PistesMidi
            v = Me.Récup_Volume(i)
            List1.Add(v)
        Next i
        '
        MaxV = List1.Item(0)
        For Each v1 As Integer In List1
            If v1 > MaxV Then
                MaxV = v1
            End If
        Next
        Return MaxV
    End Function
    Function Det_MinVolume() As Integer
        Dim i, v As Integer
        Dim MinV As Integer
        '
        Dim List1 As New List(Of Integer)

        For i = 0 To nb_TotalPistes - 1
            v = Me.Récup_Volume(i)
            List1.Add(v)
        Next i
        '
        MinV = List1.Item(0)
        For Each v1 As Integer In List1
            If v1 < MinV Then
                MinV = v1
            End If
        Next
        Return MinV
    End Function


    Private Function Det_ListAcc_old() As String
        Dim i As Integer
        Dim a As String = " "

        For i = 1 To nbMesures
            If Trim(Grid2.Cell(2, i).Text) <> "" Then
                a = a + Grid2.Cell(2, i).Text + "-"
            End If
        Next
        i = a.Length
        a = Trim(Microsoft.VisualBasic.Left(a, i - 1))
        Return a
    End Function

    Private Function Det_ListAcc() As String
        Dim i, j As Integer
        Dim a As String = " "
        Dim liste1 As New List(Of String)

        For i = 1 To nbMesures
            If Trim(Grid2.Cell(2, i).Text) <> "" Then
                j = Convert.ToInt32(Grid2.Cell(3, i).Text)
                Do
                    j = j - 1
                    a = a + i.ToString + ".1.1" + "-" + Grid2.Cell(2, i).Text + ";"
                Loop Until j <= 0
            End If
        Next
        i = a.Length
        a = Trim(Microsoft.VisualBasic.Left(a, i - 1))
        Return a
    End Function
    Public Function Det_ListTon() As String ' pour insertion pianoroll,drumedit et mix / à revoir
        Dim i As Integer
        Dim a As String = " "
        Dim liste1 As New List(Of String)
        Dim m, t, ct As Integer
        Dim m1, t1, ct1, d1 As String
        '
        For m = 0 To nbMesures - 1
            For t = 0 To 5
                For ct = 0 To 4
                    If Trim(TableEventH(m, t, ct).Accord) <> "" Then
                        m1 = Convert.ToString(m)
                        t1 = Convert.ToString(t)
                        ct1 = Convert.ToString(ct)
                        d1 = m1 + "." + t1 + "." + ct1
                        a = a + Trim(d1) + "-" + Trim(TableEventH(m, t, ct).Mode) + ";"
                    End If
                Next ct
            Next t
        Next m
        i = a.Length
        a = Trim(Microsoft.VisualBasic.Left(a, i - 1))
        Return a
    End Function

    Public Function Det_ListGam() As String
        Dim m, j As Integer
        Dim a As String = " "

        'count = TableEventH.GetLength(1)
        For m = 0 To nbMesures
            If TableEventH(m, 1, 1).Ligne <> -1 Then
                j = Convert.ToInt32(Grid2.Cell(3, m).Text)
                Do
                    j = j - 1
                    a = a + TableEventH(m, 1, 1).Gamme + "-"
                Loop Until j <= 0
            End If
        Next
        m = a.Length
        a = Trim(Microsoft.VisualBasic.Left(a, m - 1))
        Return a
    End Function
    Public Function Det_ListMarq() As String
        Dim m As Integer
        Dim a As String = " "

        'count = TableEventH.GetLength(1)
        For m = 0 To nbMesures
            If Trim(TableEventH(m, 1, 1).Marqueur) <> "" Then
                a = a + TableEventH(m, 1, 1).Marqueur + ";" + Convert.ToString(m) + "&"
            End If
        Next
        m = a.Length
        a = Trim(Microsoft.VisualBasic.Left(a, m - 1))
        Return a
    End Function
    ' **************************************************************************************
    ' ChoixVariationGrid2 : met à jour le choix du magneto effectué par l'utilisateur en     *
    ' en cliquant sur les cellules correspondantes de Grid2                                *
    ' FisrtC : 1ere colonne de la sélection de l'utilisateur                               *
    ' LastC :  dernière colonne de la sélection de l'utilisateur                           *
    ' Ligne :  N° de ligne cliquée = Magneto+4                                             *
    ' **************************************************************************************

    Sub ChoixVariationGrid2(firstC As Integer, lastC As Integer, ligne As Integer)
        Dim i As Integer
        Dim j As Integer
        If ligne >= 4 And ligne < 11 Then
            For i = firstC To lastC
                ' raz colonne
                For j = 4 To 10 ' Grid2.Rows - 1
                    Grid2.Cell(j, i).BackColor = Color.White
                Next j
                ' majcolonne et tables
                Grid2.Cell(ligne, i).BackColor = TabCoulMagneto(ligne - 4)
                ' TabControl5.SelectTab(ligne - 4) <-- supprimé car prend trop de temps à s'afficher
                TableEventH(i, 1, 1).NumMagnéto = ligne - 4
            Next i
        End If

    End Sub
    '
    Sub ChoixVariationGrid2_3(firstC As Integer, ligne As Integer) ' mise à jour des couleurs des magnetos et du n° de magnéto dans TableEventH
        Dim j As Integer = firstC
        Dim DerAcc As Integer = Det_NumDerAccord()

        Grid2.AutoRedraw = False
        Do ' boucle de mise à jour
            ' raz ancienne variation


            For i = 4 To 10
                Grid2.Cell(i, j).BackColor = Color.White
            Next

            ' maj nouvelle variation
            Grid2.Cell(ligne, j).BackColor = TabCoulMagneto(ligne - 4)
            'TabControl5.SelectTab(ligne - 4)
            TableEventH(j, 1, 1).NumMagnéto = ligne - 4
            '
            j = j + 1
        Loop Until Grid2.Cell(2, j).Text <> "" Or j > nbMesures
        Grid2.AutoRedraw = True
        Grid2.Refresh()
    End Sub

    ' ******************************************************************************************
    ' IdentVariation : identifie le N° de variation à utiliser pour jouer un accord            *
    ' NumAcc : n° de l'accord c'est aussi le N° de la colonne de Grid2 qui contient l'accord   *
    ' Return : N° Variation affectée à l'accord                                                *
    ' ******************************************************************************************
    Function IdentVariation(NumAcc As Integer) As Integer
        Dim i As Integer
        IdentVariation = 0
        For i = 4 To Grid2.Rows - 1
            If Grid2.Cell(i, NumAcc).BackColor <> Color.White Then
                IdentVariation = i - 4
                Exit For
            End If
        Next i
    End Function

    ' **************************************************************************************
    ' MajTabCoulVariation : les couleurs des variations sont définies initialement dans    *
    ' dans un tableau Publique                                                             *
    ' **************************************************************************************
    Public Sub MajTabCoulVariations()
        TabCoulMagneto(0) = Color.DarkKhaki
        TabCoulMagneto(1) = Color.LightSteelBlue
        TabCoulMagneto(2) = Color.LightSkyBlue
        TabCoulMagneto(3) = Color.Moccasin ' RGB(255, 255, 0) ' Color.SeaShell 'Coul_Zone4
        TabCoulMagneto(4) = Color.Tan
        TabCoulMagneto(5) = Color.LightGray
        TabCoulMagneto(6) = Color.LightGreen 'Color.MediumTurquoise
    End Sub
    ' **************************************************************************************
    ' Init_Variations : met à jour les couleurs des lignes de magnetos dans Grid2          *
    ' avec affacement (blanc) du magneto 1 pour le rendre actif à l'init                   *
    ' Appelé dans Construction                                                             *
    ' **************************************************************************************
    Sub Init_Variations()
        'Maj_Variations()
        Grid2.Range(4, 1, 4, nbMesuresUtiles).BackColor = TabCoulMagneto(0)
    End Sub

    ' **************************************************************************************
    ' Maj_Variations : met à jour les couleurs des lignes de variation dans Grid2          *
    ' Utilisé à l'init avec affacement (blanc) du magneto 1 pour le rendre actif à l'init  *
    ' **************************************************************************************
    Sub Maj_Variations()
        Dim i As Integer
        For i = 0 To nb_Variations - 1
            Grid2.Range(i + 4, 0, i + 4, nbMesuresUtiles).BackColor = Color.White
        Next i
    End Sub
    ' **********************************************************************************
    ' Det_IndexPisteMidi : fournit le N° de piste MIDI en fonction du N° de Bloc.      *
    ' Entrées :  N° de générateurs de 0 à 23                                           *
    ' Sortie  : N° de Piste MIDI de 0 à 5                                              *
    ' **********************************************************************************
    Public Function Det_IndexPisteMidi(Num_Générateur As Integer) As Integer
        Dim Num_Variation As Integer = Fix(Num_Générateur / 6)
        Det_IndexPisteMidi = Num_Générateur - (Num_Variation * 6)
    End Function
    ' *********************************************************************************
    ' Det_IndexVariation : fournit le N° d'une variation en fonction du N° d'un       *
    ' composant de bloc. Les N° de composants de bloc vont de 0 à 41 (7 x 6)          *
    ' Entrées : N° composant (généralement appelé Ind dans les procédures/évènements) *
    ' Sortie  : N° de Variation                                                       *
    ' *********************************************************************************
    Public Function Det_IndexVariation(Ind As Integer) As Integer
        Det_IndexVariation = Fix(Ind / 6)
    End Function

    ' ****************************************************************************************
    ' Det_IndexGénérateur : fournit l'index d. Pour 4 magnetos                                *
    ' les index de  générateurs vont de 0 à 23 pour un système à 6 pistes MIDI et 4 Mgnetos  *                            *
    ' Entrées : PisteMIDI de 0 à 5, N° Magneto : de 0 à 3                                    *
    ' Sortie : index piste des magnetos de 0 à 23.                                           *
    ' ****************************************************************************************
    Public Function Det_IndexGénérateur(PisteMidi As Integer, Magneto As Integer) As Integer
        Dim tbl(0 To nb_PistesVar - 1, 0 To nb_Variations - 1) As Integer
        Dim ind As Integer = -1
        '
        Det_IndexGénérateur = 0
        For j = 0 To nb_Variations - 1
            For i = 0 To nb_PistesVar - 1 ' Arrangement1.Nb_PistesMidi - 2
                ind = ind + 1
                tbl(i, j) = ind
            Next i
        Next j
        '
        Det_IndexGénérateur = tbl(PisteMidi, Magneto)
    End Function

    ''' <summary>
    ''' N_BLOC_MIDI : fournit les N° de Bloc en fonction du canal MIDI
    ''' </summary>
    ''' <param name="CanMIDI">N° Canal MIDI</param>
    ''' <returns></returns>
    Function N_BLOC_MIDI(CanMIDI As Integer) As String
        Dim i, j, k As Integer
        Dim a As String = ""
        i = 0
        j = CanMIDI
        k = 0
        Do
            a = a + j.ToString + " "
            i = i + 6
            j = CanMIDI + i
            k = k + 1
        Loop Until k >= Module1.nb_Variations
        '
        Return Trim(a)
    End Function

    Sub Maj_CouleursGrid2()
        Grid2.AutoRedraw = False
        '
        With Grid2.Range(0, 1, 0, (nbMesures)) ' ligne N° des accords
            .BackColor = Color.Beige
        End With
        '
        With Grid2.Range(1, 1, 1, (nbMesures)) ' ligne "Marqueurs"
            .BackColor = Color.AliceBlue
            .ForeColor = Color.Red
        End With
        '
        With Grid2.Range(2, 1, 2, (nbMesures)) ' Ligne "Accords"
            .BackColor = Color.White
        End With
        '
        With Grid2.Range(3, 1, 3, (nbMesures)) ' Ligne "Répétions"
            .BackColor = Color.White
        End With


        ' colorisation en gris des mesures non utilisées dans grid2
        For i = 1 To Grid2.Rows - 1
            Grid2.Range(i, (nbMesures) + 1, i, Grid2.Cols - 1).BackColor = Color.Gray
            Grid2.Range(i, (nbMesures) + 1, i, Grid2.Cols - 1).ForeColor = Color.Gainsboro
        Next i
        '
        ' mise à jour des couleurs des variations dans la colonne fixe de Grid2
        Grid2.FixedCols = 0
        For i = 4 To 10
            Grid2.Cell(i, 0).BackColor = TabCoulMagneto(i - 4)
        Next
        Grid2.FixedCols = 1

        ' Initialisation de la variation par défaut
        Grid2.Range(4, 1, 4, nbMesuresUtiles).BackColor = TabCoulMagneto(0)

        Grid2.AutoRedraw = True
        Grid2.Refresh()

    End Sub
    Function Det_NomFichMIDI() As Boolean ' détermination du chemin fichier Accords *.mid
        Dim Finfo As FileInfo
        Dim NomFichSeul As String

        Det_NomFichMIDI = True
        '
        If LangueIHM = "fr" Then
            SaveFileDialog2.Title = "Export Accords MIDI"
        Else
            SaveFileDialog2.Title = "MIDI Chords Export"
        End If
        '
        SaveFileDialog2.Filter = "MIDI Files (*.mid*)|*.mid" + "|All Files|*.*"

        SaveFileDialog2.OverwritePrompt = True ' affichage message "Exise déjà ...."
        SaveFileDialog2.InitialDirectory = CheminFichierMIDI
        'SaveFileDialog2.FileName = ""
        '
        'PianoRoll1.F1.Hide()
        '
        For i = 0 To nb_PianoRoll - 1
            'listPIANOROLL(i).F1.Hide()
        Next
        '
        If SaveFileDialog2.ShowDialog = Windows.Forms.DialogResult.OK Then
            FichierMIDI = SaveFileDialog2.FileName
            Finfo = My.Computer.FileSystem.GetFileInfo(FichierMIDI)
            CheminFichierMIDI = Finfo.DirectoryName
            '
            NomFichSeul = My.Computer.FileSystem.GetName(FichierMIDI)
            SaveFileDialog2.FileName = NomFichSeul
        Else
            Det_NomFichMIDI = False
        End If
    End Function
    Private Sub RAZ_TableNotesAccordsZ()
        Dim m, t, ct As Integer

        '
        For m = 0 To UBound(TableNotesAccordsZ, 1)
            For t = 0 To UBound(TableNotesAccordsZ, 2)
                For ct = 0 To UBound(TableNotesAccordsZ, 3)
                    TableNotesAccordsZ(m, t, ct) = ""
                Next ct
            Next t
        Next m

    End Sub
    Public Sub Maj_Nligne() ' mise à jour des N° de ligne dans TableEventh
        Dim i, m, t, ct As Integer
        i = 0
        For m = 0 To UBound(TableEventH, 1)
            For t = 0 To UBound(TableEventH, 2) 'nbTempsMesure '- 1
                For ct = 0 To UBound(TableEventH, 3) 'nbDivTemps '- 1
                    If Trim(TableEventH(m, t, ct).Accord) <> "" Then
                        TableEventH(m, t, ct).Ligne = i
                        i = i + 1
                    Else
                        TableEventH(m, t, ct).Ligne = -1
                    End If
                Next
            Next
        Next
    End Sub



    Function TradAcc_BemEnDiese(accord As String) As String
        Dim tbl() As String
        Dim a As String
        Dim i As Integer
        '
        TradAcc_BemEnDiese = Trim(accord)
        '
        tbl = Split(accord)
        a = LCase(tbl(0))
        '
        For i = 0 To 35
            If a = TabNotesB(i) Then
                a = TabNotesD(i) ' on traduit la tonique de l'accord en #
                Exit For
            End If
        Next
        If UBound(tbl) > 0 Then
            TradAcc_BemEnDiese = UCase(a) + " " + tbl(1)
        Else
            TradAcc_BemEnDiese = UCase(a)
        End If
    End Function
    Function Trad_RacineEnBem(noteR As String) As String
        Dim i As Integer

        Trad_RacineEnBem = noteR
        For i = 0 To 35
            If noteR = TabNotesD(i) Then
                Trad_RacineEnBem = TabNotesB(i) ' on traduit la racine en b
                Exit For
            End If
        Next
    End Function
    Sub MarquerAccord(ind As Integer)
        Dim ligne As Integer
        Dim degré As Integer

        If Trim(TabTons.Item(ind).Text) <> "" And TabTons.Item(ind).Text <> "___" Then

            RAZ_CouleurMarquée()
            '
            TabTons.Item(ind).BackColor = Couleur_Accord_Marqué
            '
            TabTons.Item(ind).ForeColor = Couleur_lettres_Accord_Marqué 'Color.Yellow

            '
            TabTons.Item(ind).Refresh()

            '
            degré = Det_IndexDegré2(ind)
            ligne = Det_LigneTableGlobale(ind)
            '
            TableCoursAcc(ligne, degré).Marqué = True
            '
            AccordMarqué = Trim(TabTons.Item(ind).Text)
            '
            Select Case ligne
                Case 0
                    OrigineAccord = Modes.Majeur
                Case 1
                    OrigineAccord = Modes.MineurH
                Case 2
                    OrigineAccord = Modes.MineurM
            End Select
            '
            TabTons.Item(ind).Refresh()
        End If
    End Sub
    Sub MarquerAccordVoisin(ind As Integer)
        Dim ligne As Integer
        Dim degré As Integer

        If Trim(TabTonsVoisins.Item(ind).Text) <> "" And TabTonsVoisins.Item(ind).Text <> "___" Then

            RAZ_CouleurMarquéeVoisins()
            '
            TabTonsVoisins.Item(ind).BackColor = Couleur_Accord_Marqué
            '
            TabTonsVoisins.Item(ind).ForeColor = Couleur_lettres_Accord_Marqué 'Color.Yellow
            '
            TabTonsVoisins.Item(ind).Refresh()
            '
            degré = Det_IndexDegré2(ind)
            ligne = Det_LigneTableGlobale(ind)
            '
            TableCoursAccVoisins(ligne, degré).Marqué = True
            '
            AccordMarquéVoisin = Trim(TabTonsVoisins.Item(ind).Text)
            AccordMarqué = AccordMarquéVoisin
            '
            Select Case ligne
                Case 0, 1, 2
                    OrigineAccord = Modes.Majeur
                Case Else
                    OrigineAccord = Modes.MineurH
            End Select
            '
            TabTonsVoisins.Item(ind).Refresh()
        End If
    End Sub
    Function Det_AccordMarqué() As Integer
        Dim i As Integer
        '
        Det_AccordMarqué = -1
        For i = 0 To 20
            If TabTons.Item(i).BackColor = Couleur_Accord_Marqué Then
                Det_AccordMarqué = i
                Exit For
            End If
        Next
    End Function
    Sub RAZ_CouleurMarquée()
        Dim ind As Integer
        Dim ligne As Integer
        Dim degré As Integer
        Dim a As String

        For ind = 0 To 20
            degré = Det_IndexDegré2(ind)
            ligne = Det_LigneTableGlobale(ind)
            If TableCoursAcc(ligne, degré).Marqué = True Then
                TabTons.Item(ind).ForeColor = Color.Black

                Select Case ligne
                    Case 0
                        a = Det_TonaCours2()
                        TabTons.Item(ind).BackColor = DicoCouleur.Item(a) 'Couleur_Accord_Majeur
                        '
                    Case Else
                        TabTons.Item(ind).BackColor = Couleur_Accord_Mineur
                End Select
                '
            End If
        Next
        '
        AccordMarqué = "" ' indication absence accord marqué
    End Sub
    Sub RAZ_CouleurMarquéeVoisins()
        Dim ind As Integer
        Dim ligne As Integer
        Dim degré As Integer

        For ind = 0 To 41
            degré = Det_IndexDegré2(ind)
            ligne = Det_LigneTableGlobale(ind)
            '
            If TableCoursAccVoisins(ligne, degré).Marqué = True Then
                TabTonsVoisins.Item(ind).ForeColor = Color.Black
                Select Case ligne
                    Case 0, 1, 2
                        TabTonsVoisins.Item(ind).BackColor = Couleur_Accord_Majeur
                    Case Else
                        TabTonsVoisins.Item(ind).BackColor = Couleur_Accord_Mineur
                End Select
                '
            End If
        Next
        '
        AccordMarqué = "" ' indication absence accord marqué
    End Sub
    Function Det_Etendu(Ligne As Integer, degré As Integer, N_Note As Integer, Note As Byte) As Byte
        Dim Note1 As Integer
        Det_Etendu = Note
        If TableCoursAcc(Ligne, degré).EtendreChecked(N_Note) = True Then
            Note1 = Note + 12
            If Note1 > 127 Then
                Note1 = Note
            End If
            Det_Etendu = Note1
        End If
    End Function
    Function Det_EtenduNumNote(m As Integer, n As Integer, j As Integer) As Integer
        Dim Note1 As Integer
        Det_EtenduNumNote = j
        If m <> n Then 'signifie que la note a été étendue
            Note1 = j + 12 ' étendre le n° de ote pour l'affichage
            If Note1 > 127 Then
                Note1 = j
            End If
            Det_EtenduNumNote = Note1
        End If

    End Function
    Function Det_NoteSansOctave(a As String)
        Dim i As Integer

        i = InStr(a, "#")
        If i = 0 Then
            Det_NoteSansOctave = Mid(a, 1, 1)
        Else
            Det_NoteSansOctave = Mid(a, 1, 2)
        End If
    End Function
    Function Det_ZoneAccord(mesure As Integer) As Integer ' retour N° de zone d'appartenance de l'accord présent dans la mesure ; retour -1 si pas de zones
        Dim i As Integer
        Dim deb As Integer
        Dim terme As Integer

        Det_ZoneAccord = 0
        '
        For i = NbZones To 0 Step -1
            If Trim(TZone(i).DébutZ) <> "---" Then
                deb = Val(TZone(i).DébutZ)
                terme = Val(TZone(i).TermeZ)
                If mesure >= deb And mesure <= terme Then
                    Det_ZoneAccord = i
                    Exit For
                End If
            End If
        Next i

    End Function
    Sub AfficherAccordSource(acc As String)
        Dim accord As String
        Dim zone As Integer
        Dim tbl() As String
        Dim Note As String
        Dim Tonique As String
        Dim m As Integer


        accord = acc
        ' détermination de la clef et tranformation de l'accord s'il est en français
        ' **************************************************************************
        If Langue = "fr" Then
            tbl = Split(Trim(ComboBox1.Text), " ")
            Clef = Det_ClefLat(tbl(0))
            tbl = Split(accord, " ")
            Note = TradNoteAngl(Clef, tbl(0))
            accord = TradNoteEnMaj(Note)
            If UBound(tbl) > 0 Then
                accord = accord + " " + tbl(1)
            End If
        Else
            tbl = Split(Trim(ComboBox1.Text), " ")
            Clef = Det_Clef(tbl(0))
        End If
        ' détermination des notes de l'accord
        ' ***********************************
        a = Det_NotesAccord(Trim(accord))
        a = Trad_ListeNotesEnD(a, "-") ' si des notes sont en bémol, elles sont traduites en #
        tbl = Split(a, "-")
        Tonique = Trim(tbl(0))
        ' application de la zone courante (Globale =-1)
        ' *********************************************
        zone = Det_ZoneAccord(MesureCourante) '  MesureCourante : variable globale mise à jour dans grid2_mousedown1 et grid3_mousedown1
        If zone = -1 Then
            b = Maj_NotesCommunes(Trim(a), -1)
            b = Maj_Large_BasseMoins1(b, "PisteHorsZone", -1)
        Else
            b = Maj_NotesCommunes(Trim(a), zone)
            b = Maj_Large_BasseMoins1(b, "PisteZone", zone)
        End If
        tbl = Split(b)
        ' boucle d'affichage des notes de l'accord
        ' ****************************************
        AccordJouerPiano.Raz_NotesJouéesPiano()
        For i = 0 To UBound(tbl)
            AccordAEtéAff = True
            m = Val(ListNotesd.IndexOf(Trim(tbl(i)))) 'Det_NumNote(tbl(i))

            ' sauvegarde pour restitution de la note du clavier
            AccordJouerPiano.Notes(i) = m
            AccordJouerPiano.Notes(i) = m 'n
            AccordJouerPiano.OldBackColor(i) = LabelPiano.Item(m).BackColor
            AccordJouerPiano.OldForeColor(i) = LabelPiano.Item(m).ForeColor
            AccordJouerPiano.OldText(i) = LabelPiano.Item(m).Text




            LabelPiano.Item(m).BackColor = LabelPiano.Item(m).BackColor 'Color.Yellow ' note de l'accord en jaune sur le piano
            '               '
            If Trim(Clef) = "#" Then
                EcrNote = ListNotesd(m) 'ListNotesd(n)
            Else
                EcrNote = ListNotesb(m) 'ListNotesb(n)
            End If
            '
            Note = Det_NoteSansOctave(Trim(tbl(i)))
            'If Note = Tonique Then
            If LabelPiano.Item(m).BackColor = Color.Black Or LabelPiano.Item(m).BackColor = Color.ForestGreen Then ' Color.ForestGreen
                LabelPiano.Item(m).ForeColor = Color.PowderBlue 'Color.Red ' la tonique est en rouge
                LabelPiano.Item(m).Text = Trim(EcrNote) 'ListNotesd(n) 'Trim(tbl(i))
            Else
                LabelPiano.Item(m).ForeColor = Color.Black
                LabelPiano.Item(m).Text = Trim(EcrNote) 'ListNotesd(n) 'Trim(tbl(i))
            End If
        Next i
    End Sub
    Public Function Appartenance(Accord As String, Gamme As String) As Boolean
        Dim tblA() As String
        Dim TblG() As String
        Dim A, G As String
        Dim i, j, count_ As Integer

        Appartenance = False
        '
        A = Trad_AccordEn_D(Accord)
        A = Trim(Det_NotesAccord3(Trim(Accord), "#")) ' détermination des notes de l'accord
        '
        G = Trad_GammeEn_D(Trim(Gamme))
        G = Det_NotesGammes(Trim(Gamme)) ' détermination des notes de la gamme
        G = ListeNotesMin_EnD(G, " ")
        '
        tblA = Split(A, "-")
        TblG = Split(G)
        '
        count_ = 0
        '
        For i = 0 To UBound(tblA)
            For j = 0 To UBound(TblG)
                If tblA(i) = TblG(j) Then
                    count_ = count_ + 1
                End If
            Next j
        Next i
        '
        If ((UBound(tblA) + 1) - count_) = 0 Then ' si le nb de notes trouvés dans la gamme = au nb de notes de l'accord, alors les notes de l'accords appartiennent à la gamme
            Appartenance = True
        End If
    End Function

    Sub RAZ_AffNoteAcc()
        If AccordAEtéAff = True Then
            For i = 0 To UBound(AccordJouerPiano.Notes)
                If AccordJouerPiano.Notes(i) <> -1 Then
                    j = AccordJouerPiano.Notes(i)
                    LabelPiano.Item(j).BackColor = AccordJouerPiano.OldBackColor(i)
                    LabelPiano.Item(j).ForeColor = AccordJouerPiano.OldForeColor(i)
                    LabelPiano.Item(j).Text = Trim(AccordJouerPiano.OldText(i))
                End If
            Next i
            AccordAEtéAff = False
        End If
    End Sub
    Sub AfficherAccordRapport(Position As String)
        Dim a As String
        Dim tbl() As String
        Dim tbl2() As String
        Dim m, t, ct As Integer
        Dim Sauv_Clef As String
        Dim n As Integer
        Dim Note As String

        Sauv_Clef = Clef
        ' Détermintation des paramètre m,t,ct
        ' ***********************************
        tbl = Split(Trim(Position), ".")
        m = Val(tbl(0))
        t = Val(tbl(1))
        ct = Val(tbl(2))
        ' Recalcul du voicing en fonction des zones
        ' *****************************************
        a = TableEventH(m, t, ct).Tonalité
        tbl2 = Split(a)
        Clef = Det_Clef(tbl2(0))
        Calcul_AutoVoicingZ()
        ' Affichage
        ' *********
        If TableEventH(m, t, ct).Ligne <> -1 Then
            ' Détermination de la Tonique : pour affichage en rouge sur le clavier
            ' ********************************************************************
            tbl = Split(Det_NotesAccord(TableEventH(m, t, ct).Accord), "-")
            Tonique = Trim(tbl(0))
            Tonique = Trad_ListeNotesEnD(Trim(Tonique), "-") ' traduire la tonique en #
            ' Identification des notes de l'accord 
            ' ************************************
            a = TableNotesAccordsZ(m, t, ct) ' ici les notes sont toujours en # - Maj effectuée par Calcul_AutoVoicingZ()
            tbl = Split(a)
            AccordJouerPiano.Raz_NotesJouéesPiano()
            For i = 0 To UBound(tbl)
                AccordAEtéAff = True
                n = ListNotesd.IndexOf(Trim(tbl(i))) ' identification de la note
                ' sauvegarde pour restitution de la note du clavier
                AccordJouerPiano.Notes(i) = n
                AccordJouerPiano.OldBackColor(i) = LabelPiano.Item(n).BackColor
                AccordJouerPiano.OldForeColor(i) = LabelPiano.Item(n).ForeColor
                AccordJouerPiano.OldText(i) = LabelPiano.Item(n).Text
                '
                LabelPiano.Item(n).BackColor = LabelPiano.Item(n).BackColor 'Color.Yellow
                If Clef = "b" Then
                    LabelPiano.Item(n).Text = ListNotesb(n)
                Else
                    LabelPiano.Item(n).Text = ListNotesd(n)
                End If
                Note = Det_NoteSansOctave(Trim(tbl(i)))
                'If Note = Tonique Then
                If LabelPiano.Item(n).BackColor = Color.Black Or LabelPiano.Item(n).BackColor = Color.ForestGreen Then ' Color.ForestGreen
                    LabelPiano.Item(n).ForeColor = Color.PowderBlue
                Else
                    LabelPiano.Item(n).ForeColor = Color.Blue
                End If
            Next i
        End If
        Clef = Sauv_Clef
    End Sub
    Function Det_NumNote(Note As String) As String
        'Dim tbl() As Object
        'Dim Clef As String

        'tbl = Split(ComboBox20.Text)
        'Clef = Det_Clef(tbl(0))

        If Trim(Clef) = "#" Then
            Det_NumNote = ListNotesd.IndexOf(Trim(Note))
        Else
            Det_NumNote = ListNotesb.IndexOf(Trim(Note))
        End If
    End Function
    Sub ActivationDesMenus()
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim m As Integer

        Dim degré As Integer
        Dim Active9 As Boolean
        Dim Active11 As Boolean
        '
        Dim AuMoinsUneSelection As Boolean
        '
        Dim a As String
        Dim b As String
        '

        Active9 = False
        Active11 = False
        '
        AuMoinsUneSelection = False
        '
        For i = 2 To 3 ' on fait le parcours uniquement les accords 9e et 11e (les autres sont obligtoirement activés)
            m = -1
            For j = 0 To 2 ' mode Maj, MinH, MinM
                For k = 0 To 6
                    m = m + 1
                    '
                    AuMoinsUneSelection = True
                    '
                    a = TabTonsDegrés.Item(k).Text
                    degré = Det_IndexDegré(a)
                    '
                    b = TableGlobalAcc(i, j, degré)
                    If TableGlobalAcc(i, j, degré) <> "___" Then
                        Select Case i
                            Case 2
                                Active9 = True
                            Case 3
                                Active11 = True
                        End Select
                    End If
                    'End If
                Next k
            Next
        Next
        '
        If AuMoinsUneSelection = True Then
            Menu3notes.Enabled = True
            Menu4Notes.Enabled = True
            MenuNotes9.Enabled = Active9
            MenuNotes11.Enabled = Active11
        Else
            Menu3notes.Enabled = False
            Menu4Notes.Enabled = False
            MenuNotes9.Enabled = False
            MenuNotes11.Enabled = False
        End If
        '
    End Sub
    Function Calc_CadDegrés(degré As String) As Integer
        Calc_CadDegrés = 0
        Select Case Cad_OrigineAccord

            Case Modes.Majeur, Modes.Cadence_Majeure
                Calc_CadDegrés = Det_IndexDegré(Trim(degré))

            Case Modes.MineurH, Modes.Cadence_Mineure
                Calc_CadDegrés = Det_IndexDegréMin(Trim(degré))

            Case Modes.Cadence_Mixte
                Select Case Trim(degré)
                    Case "VI"
                        Calc_CadDegrés = 0
                    Case "V"
                        Calc_CadDegrés = 4
                    Case "IV"
                        Calc_CadDegrés = 3
                    Case "III"
                        Calc_CadDegrés = 2
                End Select
        End Select
    End Function
    Function Det_ValeurRenv(i As Integer, j As Integer) As String
        Det_ValeurRenv = ""
        Select Case TableCoursAcc(i, j).RenvChoisi
            Case 0
                Det_ValeurRenv = Renversement1.Text
            Case 1
                Det_ValeurRenv = Renversement2.Text
            Case 2
                Det_ValeurRenv = Renversement3.Text
            Case 3
                Det_ValeurRenv = Renversement4.Text
            Case 4
                Det_ValeurRenv = Renversement5.Text
        End Select
    End Function
    Function Det_ValeurOct(i As Integer, j As Integer) As String
        Det_ValeurOct = ""
        Select Case TableCoursAcc(i, j).OctaveChoisie
            Case 0
                Det_ValeurOct = OctavePlus1.Text
            Case 1
                Det_ValeurOct = Octave0.Text
            Case 2
                Det_ValeurOct = OctaveMoins1.Text
            Case 3
                Det_ValeurOct = OctaveMoins2.Text
        End Select
    End Function
    Sub Maj_Octave(ind As Integer)
        Dim degré As Integer
        Dim ligne As Integer
        '
        'If TabTonsSelect.Item(ind).Checked = True Then
        '
        ' = TabTonsDegrés.Item(Det_IndexDegré2(ind)).Text
        'degré = Det_IndexDegré(a)
        '
        degré = Det_IndexDegré2(ind)
        ligne = Det_LigneTableGlobale(ind)
        '
        Select Case TableCoursAcc(ligne, degré).OctaveChoisie
            Case 0
                OctavePlus1.Checked = True
                Octave0.Checked = False
                OctaveMoins1.Checked = False
                OctaveMoins2.Checked = False
            Case 1
                OctavePlus1.Checked = False
                Octave0.Checked = True
                OctaveMoins1.Checked = False
                OctaveMoins2.Checked = False
            Case 2
                OctavePlus1.Checked = False
                Octave0.Checked = False
                OctaveMoins1.Checked = True
                OctaveMoins2.Checked = False
            Case 3
                OctavePlus1.Checked = False
                Octave0.Checked = False
                OctaveMoins1.Checked = False
                OctaveMoins2.Checked = True
        End Select
        ' End If
        '
    End Sub
    Private Sub TabTons_MouseUp(ByVal sender As Object, ByVal e As MouseEventArgs)
        Dim ind As Integer
        Dim com As Label = sender
        '
        ind = Val(com.Tag)
        '
        If AccordAEtéJoué = True Then
            CouperJouerAccord2()
            AccordAEtéJoué = False
        End If
        '
        RAZ_AffNoteAcc()
        ' Pour glisser - déposer
        ' **********************
        'TabTons.Item(ind).DoDragDrop(TabTons.Item(ind).Text, DragDropEffects.Copy Or DragDropEffects.Move)
    End Sub
    Sub Maj_Sélection(Index As Integer)
        Select Case Index
            Case 0, 7, 14

                '
                TabTons.Item(0).BackColor = Couleur_Accord_DéSélectionné
                TabTons.Item(7).BackColor = Couleur_Accord_DéSélectionné
                TabTons.Item(14).BackColor = Couleur_Accord_DéSélectionné
            '
            Case 1, 8, 15
                '
                TabTons.Item(1).BackColor = Couleur_Accord_DéSélectionné
                TabTons.Item(8).BackColor = Couleur_Accord_DéSélectionné
                TabTons.Item(15).BackColor = Couleur_Accord_DéSélectionné
            Case 2, 9, 16

                '
                TabTons.Item(2).BackColor = Couleur_Accord_DéSélectionné
                TabTons.Item(9).BackColor = Couleur_Accord_DéSélectionné
                TabTons.Item(16).BackColor = Couleur_Accord_DéSélectionné

            Case 3, 10, 17
                '
                TabTons.Item(3).BackColor = Couleur_Accord_DéSélectionné
                TabTons.Item(10).BackColor = Couleur_Accord_DéSélectionné
                TabTons.Item(17).BackColor = Couleur_Accord_DéSélectionné
            Case 4, 11, 18

                '
                TabTons.Item(4).BackColor = Couleur_Accord_DéSélectionné
                TabTons.Item(11).BackColor = Couleur_Accord_DéSélectionné
                TabTons.Item(18).BackColor = Couleur_Accord_DéSélectionné
            Case 5, 12, 19
                '
                TabTons.Item(5).BackColor = Couleur_Accord_DéSélectionné
                TabTons.Item(12).BackColor = Couleur_Accord_DéSélectionné
                TabTons.Item(19).BackColor = Couleur_Accord_DéSélectionné
            Case 6, 13, 20

                '
                TabTons.Item(6).BackColor = Couleur_Accord_DéSélectionné
                TabTons.Item(13).BackColor = Couleur_Accord_DéSélectionné
                TabTons.Item(20).BackColor = Couleur_Accord_DéSélectionné
        End Select
        '

        TabTons.Item(Index).BackColor = Couleur_Accord_Sélectionné
    End Sub
    Public Function Mode2(Tonique As String, Type_Mode As String, Typ As Integer, Voisin As Boolean) As String 'TYpe_Mode = Maj, MinH, MinM Typ= 3 -> 3note,Type=4 -> Accord7,Typ=5->Accord9,Typ=6 -> Accord11
        Dim ton As String
        Dim Sauv_Clef As String
        Dim tbl() As String
        Dim a As String

        ton = LCase(Tonique)
        Mode2 = ""
        Sauv_Clef = Clef
        '
        If Voisin = False Then
            tbl = Split(Trim(ComboBox1.Text)) ' on considère ici la tonalité pour déterminer la clef
            a = Trim(tbl(0))
        Else
            a = Trim(Tonique)
        End If
        '
        If Langue = "fr" Then
            a = TradNoteAngl2(a)
        End If
        Clef = Det_Clef(Trim(a))
        '
        Maj_TabNotes_Minus(Trim(Clef))

        Select Case Type_Mode
            Case "Maj"
                Select Case Typ
                    Case 3
                        Mode2 =
                        Tonique + "-" _
                        + UCaseBémol(NoteInterval3(ton, "9", Clef)) + " m" + "-" _
                        + UCaseBémol(NoteInterval3(ton, "3", Clef)) + " m" + "-" _
                        + UCaseBémol(NoteInterval3(ton, "4", Clef)) + "-" _
                        + UCaseBémol(NoteInterval3(ton, "5", Clef)) + "-" _
                        + UCaseBémol(NoteInterval3(ton, "6", Clef)) + " m" + "-" _
                        + UCaseBémol(NoteInterval3(ton, "M7", Clef)) + " mb5"
                    Case 4 ' 7

                        Mode2 =
                        Tonique + " M7" + "-" _
                        + UCaseBémol(NoteInterval3(ton, "9", Clef)) + " m7" + "-" _
                        + UCaseBémol(NoteInterval3(ton, "3", Clef)) + " m7" + "-" _
                        + UCaseBémol(NoteInterval3(ton, "4", Clef)) + " M7" + "-" _
                        + UCaseBémol(NoteInterval3(ton, "5", Clef)) + " 7" + "-" _
                        + UCaseBémol(NoteInterval3(ton, "6", Clef)) + " m7" + "-" _
                        + UCaseBémol(NoteInterval3(ton, "M7", Clef)) + " m7b5"
                    Case 5 ' 9

                        Mode2 =
                        Tonique + " M7(9)" + "-" _
                        + UCaseBémol(NoteInterval3(ton, "9", Clef)) + " m7(9)" + "-" _
                        + "___" + "-" _
                        + UCaseBémol(NoteInterval3(ton, "4", Clef)) + " M7(9)" + "-" _
                        + UCaseBémol(NoteInterval3(ton, "5", Clef)) + " 7(9)" + "-" _
                        + UCaseBémol(NoteInterval3(ton, "6", Clef)) + " m7(9)" + "-" _
                        + "___" + "-"
                    Case 6 ' 11

                        Mode2 =
                        Tonique + " 11" + "-" _
                        + UCaseBémol(NoteInterval3(ton, "9", Clef)) + " m11" + "-" _
                        + UCaseBémol(NoteInterval3(ton, "3", Clef)) + " m11" + "-" _
                        + UCaseBémol(NoteInterval3(ton, "4", Clef)) + " M7(11#)" + "-" _
                        + UCaseBémol(NoteInterval3(ton, "5", Clef)) + " 7(11)" + "-" _
                        + UCaseBémol(NoteInterval3(ton, "6", Clef)) + " m11" + "-" _
                        + "___"
                End Select
            Case "MinH"
                Select Case Typ
                    Case 3
                        Mode2 =
                        Tonique + " m" + "-" _
                        + UCaseBémol(NoteInterval3(ton, "9", Clef)) + " mb5" + "-" _
                        + UCaseBémol(NoteInterval3(ton, "3m", Clef)) + " 5#" + "-" _
                        + UCaseBémol(NoteInterval3(ton, "4", Clef)) + " m" + "-" _
                        + UCaseBémol(NoteInterval3(ton, "5", Clef)) + "-" _
                        + UCaseBémol(NoteInterval3(ton, "5#", Clef)) + "-" _
                        + UCaseBémol(NoteInterval3(ton, "M7", Clef)) + " mb5"
                    Case 4 ' 7
                        Mode2 =
                        Tonique + " mM7" + "-" _
                        + UCaseBémol(NoteInterval3(ton, "9", Clef)) + " m7b5" + "-" _
                        + UCaseBémol(NoteInterval3(ton, "3m", Clef)) + " M75#" + "-" _
                        + UCaseBémol(NoteInterval3(ton, "4", Clef)) + " m7" + "-" _
                        + UCaseBémol(NoteInterval3(ton, "5", Clef)) + " 7" + "-" _
                        + UCaseBémol(NoteInterval3(ton, "5#", Clef)) + " M7" + "-" _
                        + UCaseBémol(NoteInterval3(ton, "M7", Clef)) + " 7dim"
                    Case 5 ' 9

                        Mode2 =
                        "___" + "-" _
                        + "___" + "-" _
                        + "___" + "-" _
                        + UCaseBémol(NoteInterval3(ton, "4", Clef)) + " m7(9)" + "-" _
                        + UCaseBémol(NoteInterval3(ton, "5", Clef)) + " 7(b9)" + "-" _
                        + UCaseBémol(NoteInterval3(ton, "5#", Clef)) + " 7(9#)" + "-" _
                        + "___"
                    Case 6 '11
                        Mode2 =
                            Tonique + " m11" + "-" _
                        + "___" + "-" _
                        + "___" + "-" _
                        + "___" + "-" _
                        + UCaseBémol(NoteInterval3(ton, "5", Clef)) + " 7(11)" + "-" _
                        + UCaseBémol(NoteInterval3(ton, "5#", Clef)) + " M7(11#)" + "-" _
                        + "___"

                End Select
            Case "MinM"
                Select Case Typ
                    Case 3
                        Mode2 =
                    Tonique + " m" + "-" _
                    + UCaseBémol(NoteInterval3(ton, "9", Clef)) + " m" + "-" _
                    + UCaseBémol(NoteInterval3(ton, "3m", Clef)) + " 5#" + "-" _
                    + UCaseBémol(NoteInterval3(ton, "4", Clef)) + "-" _
                    + UCaseBémol(NoteInterval3(ton, "5", Clef)) + "-" _
                    + UCaseBémol(NoteInterval3(ton, "6", Clef)) + " mb5" + "-" _
                    + UCaseBémol(NoteInterval3(ton, "M7", Clef)) + " mb5"

                    Case 4 ' 7
                        Mode2 =
                    Tonique + " mM7" + "-" _
                    + UCaseBémol(NoteInterval3(ton, "9", Clef)) + " m7" + "-" _
                    + UCaseBémol(NoteInterval3(ton, "3m", Clef)) + " M75#" + "-" _
                    + UCaseBémol(NoteInterval3(ton, "4", Clef)) + " 7" + "-" _
                    + UCaseBémol(NoteInterval3(ton, "5", Clef)) + " 7" + "-" _
                    + UCaseBémol(NoteInterval3(ton, "6", Clef)) + " m7b5" + "-" _
                    + UCaseBémol(NoteInterval3(ton, "M7", Clef)) + " m7b5"
                    Case 5 '9

                        Mode2 =
                        "___" + "-" _
                        + "___" + "-" _
                        + "___" + "-" _
                        + UCaseBémol(NoteInterval3(ton, "4", Clef)) + " 7(9)" + "-" _
                        + UCaseBémol(NoteInterval3(ton, "5", Clef)) + " 7(9)" + "-" _
                        + "___" + "-" _
                        + "___"
                    Case 6 '11
                        Mode2 =
                        Tonique + " m11" + "-" _
                        + UCaseBémol(NoteInterval3(ton, "9", Clef)) + " m11" + "-" _
                        + "___" + "-" _
                        + "___" + "-" _
                        + UCaseBémol(NoteInterval3(ton, "5", Clef)) + " 7(11)" + "-" _
                        + "___" + "-" _
                        + "___"
                End Select
            Case "MinN"
                Sauv_Clef = Clef
                a = Det_RelativeMajeure(Tonique + " MinH")
                Clef = Det_Clef(Trim(a))
                Maj_TabNotes_Minus(Trim(Clef))
                Select Case Typ
                    Case 3
                        Mode2 =
                            Tonique + " m" + "-" _
                            + UCaseBémol(NoteInterval3(ton, "9", Clef)) + " mb5" + "-" _
                            + UCaseBémol(NoteInterval3(ton, "3m", Clef)) + "-" _
                            + UCaseBémol(NoteInterval3(ton, "4", Clef)) + " m" + "-" _
                            + UCaseBémol(NoteInterval3(ton, "5", Clef)) + " m" + "-" _
                            + UCaseBémol(NoteInterval3(ton, "5#", Clef)) + "-" _
                            + UCaseBémol(NoteInterval3(ton, "7", Clef)) + ""
                    Case 4 ' 7
                        Mode2 =
                            Tonique + " m7" + "-" _
                            + UCaseBémol(NoteInterval3(ton, "9", Clef)) + " m7b5" + "-" _
                            + UCaseBémol(NoteInterval3(ton, "3m", Clef)) + " M7" + "-" _
                            + UCaseBémol(NoteInterval3(ton, "4", Clef)) + " m7" + "-" _
                            + UCaseBémol(NoteInterval3(ton, "5", Clef)) + " m7" + "-" _
                            + UCaseBémol(NoteInterval3(ton, "5#", Clef)) + " M7" + "-" _
                            + UCaseBémol(NoteInterval3(ton, "7", Clef)) + " 7"
                End Select
        End Select
        Clef = Sauv_Clef
    End Function

    Public Function NoteInterval3(Tonique As String, Interval As String, clef As String) As String
        ' La présente fonction NoteInterval3 est très proche de NoteInterval 
        ' qui se trouve dans le module Harmonie. La différence entre ces deux fonction est la suivante :
        '   1 - NoteInterval3 prend la valeur de Clef dans la tonalite courante (Combobox1.text)
        '   2 - NoteInterval prend la valeur de Clef dans la tonique d'un accord, d'une gamme ou d'un mode
        ' En conclusion, si une modification est à faire dans l'un de ces deux fonctions, elle sera 
        ' sans doute à faire dans l'autre fonction.
        Dim IndexTonique As Integer
        '
        Maj_TabNotes_Minus(Trim(clef))
        '
        NoteInterval3 = ""
        IndexTonique = TrouverNoteDansTabNotes(LCase(Tonique))
        '
        Select Case Interval
            Case "b2", "b9"
                NoteInterval3 = TabNotes(IndexTonique + 1)
            Case "2", "9"
                NoteInterval3 = TabNotes(IndexTonique + 2)
            Case "9#", "3m"
                NoteInterval3 = TabNotes(IndexTonique + 3)
            Case "3"
                NoteInterval3 = TabNotes(IndexTonique + 4)
            Case "4", "11"
                NoteInterval3 = TabNotes(IndexTonique + 5)
            Case "b5", "4#", "11#"
                NoteInterval3 = TabNotes(IndexTonique + 6)
            Case "5"
                NoteInterval3 = TabNotes(IndexTonique + 7)
            Case "5#", "6m", "13b"
                NoteInterval3 = TabNotes(IndexTonique + 8)
            Case "6", "13"
                NoteInterval3 = TabNotes(IndexTonique + 9)
            Case "7"
                NoteInterval3 = TabNotes(IndexTonique + 10)
            Case "7M", "M7"
                NoteInterval3 = TabNotes(IndexTonique + 11)
        End Select
        '
    End Function
    Public Function Det_NotesAccord3(Accord As String, Clef As String) As String
        ' La présente fonction Det_NotesAccord3 est très proche d'une autre fonction Det_NotesAccord 
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
        Det_NotesAccord3 = ""
        Select Case Trim(Chiffrage)
            Case "Maj"
                Det_NotesAccord3 = Trim(Tonique + "-" + NoteInterval3(Tonique, "3", Clef) + "-" + NoteInterval3(Tonique, "5", Clef))
            Case "m"
                Det_NotesAccord3 = Trim(Tonique + "-" + NoteInterval3(Tonique, "3m", Clef) + "-" + NoteInterval3(Tonique, "5", Clef))
            Case "mb5"
                Det_NotesAccord3 = Trim(Tonique + "-" + NoteInterval3(Tonique, "3m", Clef) + "-" + NoteInterval3(Tonique, "b5", Clef))
            Case "7M", "M7"
                Det_NotesAccord3 = Trim(Tonique + "-" + NoteInterval3(Tonique, "3", Clef) + "-" + NoteInterval3(Tonique, "5", Clef) + "-" + NoteInterval3(Tonique, "M7", Clef))
            Case "7"
                Det_NotesAccord3 = Trim(Tonique + "-" + NoteInterval3(Tonique, "3", Clef) + "-" + NoteInterval3(Tonique, "5", Clef) + "-" + NoteInterval3(Tonique, "7", Clef))
            Case "m7"
                Det_NotesAccord3 = Trim(Tonique + "-" + NoteInterval3(Tonique, "3m", Clef) + "-" + NoteInterval3(Tonique, "5", Clef) + "-" + NoteInterval3(Tonique, "7", Clef))
            Case "m7b5"
                Det_NotesAccord3 = Trim(Tonique + "-" + NoteInterval3(Tonique, "3m", Clef) + "-" + NoteInterval3(Tonique, "b5", Clef) + "-" + NoteInterval3(Tonique, "7", Clef))
            Case "5#"
                Det_NotesAccord3 = Trim(Tonique + "-" + NoteInterval3(Tonique, "3", Clef) + "-" + NoteInterval3(Tonique, "5#", Clef))
            Case "m7M", "mM7"
                Det_NotesAccord3 = Trim(Tonique + "-" + NoteInterval3(Tonique, "3m", Clef) + "-" + NoteInterval3(Tonique, "5", Clef) + "-" + NoteInterval3(Tonique, "M7", Clef))
            Case "7M5#", "M75#"
                Det_NotesAccord3 = Trim(Tonique + "-" + NoteInterval3(Tonique, "3", Clef) + "-" + NoteInterval3(Tonique, "5#", Clef) + "-" + NoteInterval3(Tonique, "M7", Clef))
            Case "7dim"
                Det_NotesAccord3 = Trim(Tonique + "-" + NoteInterval3(Tonique, "3m", Clef) + "-" + NoteInterval3(Tonique, "b5", Clef) + "-" + NoteInterval3(Tonique, "6", Clef))
            Case "7M(9)", "M7(9)"
                Det_NotesAccord3 = Trim(Tonique + "-" + NoteInterval3(Tonique, "3", Clef) + "-" + NoteInterval3(Tonique, "5", Clef) + "-" + NoteInterval3(Tonique, "M7", Clef) + "-" + NoteInterval3(Tonique, "9", Clef))
            Case "m7(9)"
                Det_NotesAccord3 = Trim(Tonique + "-" + NoteInterval3(Tonique, "3m", Clef) + "-" + NoteInterval3(Tonique, "5", Clef) + "-" + NoteInterval3(Tonique, "7", Clef) + "-" + NoteInterval3(Tonique, "9", Clef))
            Case "7(9)"
                Det_NotesAccord3 = Trim(Tonique + "-" + NoteInterval3(Tonique, "3", Clef) + "-" + NoteInterval3(Tonique, "5", Clef) + "-" + NoteInterval3(Tonique, "7", Clef) + "-" + NoteInterval3(Tonique, "9", Clef))
            Case "7(b9)"
                Det_NotesAccord3 = Trim(Tonique + "-" + NoteInterval3(Tonique, "3", Clef) + "-" + NoteInterval3(Tonique, "5", Clef) + "-" + NoteInterval3(Tonique, "7", Clef) + "-" + NoteInterval3(Tonique, "b9", Clef))
            Case "7(9#)"
                Det_NotesAccord3 = Trim(Tonique + "-" + NoteInterval3(Tonique, "3", Clef) + "-" + NoteInterval3(Tonique, "5", Clef) + "-" + NoteInterval3(Tonique, "7", Clef) + "-" + NoteInterval3(Tonique, "9#", Clef))
            Case "b9"
                Det_NotesAccord3 = Trim(Tonique + "-" + NoteInterval3(Tonique, "3", Clef) + "-" + NoteInterval3(Tonique, "5", Clef) + "-" + NoteInterval3(Tonique, "b9", Clef))
            Case "9#"
                Det_NotesAccord3 = Trim(Tonique + "-" + NoteInterval3(Tonique, "3", Clef) + "-" + NoteInterval3(Tonique, "5", Clef) + "-" + NoteInterval3(Tonique, "9#", Clef))
            Case "m9"
                Det_NotesAccord3 = Trim(Tonique + "-" + NoteInterval3(Tonique, "3m", Clef) + "-" + NoteInterval3(Tonique, "5", Clef) + "-" + NoteInterval3(Tonique, "9", Clef))
            Case "9"
                Det_NotesAccord3 = Trim(Tonique + "-" + NoteInterval3(Tonique, "3", Clef) + "-" + NoteInterval3(Tonique, "5", Clef) + "-" + NoteInterval3(Tonique, "9", Clef))
            Case "7M(11#)", "M7(11#)"
                Det_NotesAccord3 = Trim(Tonique + "-" + NoteInterval3(Tonique, "3", Clef) + "-" + NoteInterval3(Tonique, "5", Clef) + "-" + NoteInterval3(Tonique, "M7", Clef) + "-" + NoteInterval3(Tonique, "11#", Clef))
            Case "7(11)"
                Det_NotesAccord3 = Trim(Tonique + "-" + NoteInterval3(Tonique, "3", Clef) + "-" + NoteInterval3(Tonique, "5", Clef) + "-" + NoteInterval3(Tonique, "7", Clef) + "-" + NoteInterval3(Tonique, "11", Clef))
            Case "11#"
                Det_NotesAccord3 = Trim(Tonique + "-" + NoteInterval3(Tonique, "3", Clef) + "-" + NoteInterval3(Tonique, "5", Clef) + "-" + NoteInterval3(Tonique, "11#", Clef))
            Case "11"
                Det_NotesAccord3 = Trim(Tonique + "-" + NoteInterval3(Tonique, "3", Clef) + "-" + NoteInterval3(Tonique, "5", Clef) + "-" + NoteInterval3(Tonique, "11", Clef))
            Case "7sus4"
                Det_NotesAccord3 = Trim(Tonique + "-" + NoteInterval3(Tonique, "4", Clef) + "-" + NoteInterval3(Tonique, "5", Clef) + "-" + NoteInterval3(Tonique, "7", Clef))
            Case "sus4"
                Det_NotesAccord3 = Trim(Tonique + "-" + NoteInterval3(Tonique, "4", Clef) + "-" + NoteInterval3(Tonique, "5", Clef))
            Case "m11"
                Det_NotesAccord3 = Trim(Tonique + "-" + NoteInterval3(Tonique, "3m", Clef) + "-" + NoteInterval3(Tonique, "5", Clef) + "-" + NoteInterval3(Tonique, "11", Clef))
            Case "11"
                Det_NotesAccord3 = Trim(Tonique + "-" + NoteInterval3(Tonique, "3", Clef) + "-" + NoteInterval3(Tonique, "5", Clef) + "-" + NoteInterval3(Tonique, "11", Clef))
        End Select
    End Function
    Public Sub Maj_TableCoursAccInit()
        Dim i As Integer
        Dim j As Integer
        '
        For i = 0 To 2
            For j = 0 To 6
                TableCoursAcc(i, j).TyAcc = "3 Notes"
                TableCoursAcc(i, j).Accord = TableGlobalAcc(0, i, j)
                TableCoursAcc(i, j).Octave = "0"
                '
            Next
        Next
        '
    End Sub
    Public Sub Maj_TableCoursVoisinsAccInit()
        Dim i As Integer
        Dim j As Integer
        '
        For i = 0 To 2
            For j = 0 To 6
                TableCoursAccVoisins(i, j).TyAcc = "3 Notes"
                TableCoursAccVoisins(i, j).Accord = TableGlobalAccVoisin(0, i, j)
                TableCoursAccVoisins(i, j).Octave = "0"
                '
            Next
        Next
        '
    End Sub

    Public Function Det_TonaCours() As String
        Dim tbl() As String
        '
        tbl = Split(Trim(ComboBox1.Text))
        Det_TonaCours = Trim(tbl(0))
        '
    End Function
    '
    ''' <summary>
    ''' Det_TonaCours2 : détermination de la tonalité majeure en cours à partir du Combobox1
    ''' On ne garde que la note, le "Maj" est supprimée.
    ''' </summary>
    ''' <returns></returns>
    Public Function Det_TonaCours2() As String
        Dim tbl() As String
        '
        tbl = Split(Trim(ComboBox1.Text))
        'a = Trim(TradNoteAngl2(tbl(0)))
        '
        Det_TonaCours2 = Trim(tbl(0))
    End Function
    Public Function Det_TonaCours3() As String
        Dim tbl() As String
        '
        tbl = Split(Trim(ComboBox2.Text))
        Det_TonaCours3 = Trim(tbl(0))
        '
    End Function
    Public Function Det_TonaMinCours2() As String
        Dim tbl() As String

        '
        tbl = Split(Trim(ComboBox2.Text))
        'a = Trim(TradNoteAngl2(tbl(0)))
        Det_TonaMinCours2 = Trim(tbl(0))
    End Function
    Public Function Det_TonaMinCours3(tonaMaj As String) As String
        Dim tbl() As String
        Dim i As Integer
        Dim TonaMin As String
        '
        Dim TonaMajEn As New ArrayList
        Dim TonaMajFr As New ArrayList
        Dim TonaMinEn As New ArrayList
        Dim TonaMinFr As New ArrayList

        TonaMajFr.Add("Do# Majeur")
        TonaMajFr.Add("Fa# Majeur")
        TonaMajFr.Add("Si Majeur")
        TonaMajFr.Add("Mi Majeur")
        TonaMajFr.Add("La Majeur")
        TonaMajFr.Add("Ré Majeur")
        TonaMajFr.Add("Sol Majeur")
        TonaMajFr.Add("Do Majeur")
        TonaMajFr.Add("Fa Majeur")
        TonaMajFr.Add("Sib Majeur")
        TonaMajFr.Add("Mib Majeur")
        TonaMajFr.Add("Lab Majeur")
        '
        TonaMinFr.Add("La# Mineur")
        TonaMinFr.Add("Ré# Mineur")
        TonaMinFr.Add("Sol# Mineur")
        TonaMinFr.Add("Do# Mineur")
        TonaMinFr.Add("Fa# Mineur")
        TonaMinFr.Add("Si Mineur")
        TonaMinFr.Add("Mi Mineur")
        TonaMinFr.Add("La Mineur")
        TonaMinFr.Add("Ré Mineur")
        TonaMinFr.Add("Sol Mineur")
        TonaMinFr.Add("Do Mineur")
        TonaMinFr.Add("Fa Mineur")


        TonaMajEn.Add("C# Major")
        TonaMajEn.Add("F# Major")
        TonaMajEn.Add("B Major")
        TonaMajEn.Add("E Major")
        TonaMajEn.Add("A Major")
        TonaMajEn.Add("D Major")
        TonaMajEn.Add("G Major")
        TonaMajEn.Add("C Major")
        TonaMajEn.Add("F Major")
        TonaMajEn.Add("Bb Major")
        TonaMajEn.Add("Eb Major")
        TonaMajEn.Add("Ab Major")


        TonaMinEn.Add("A# Minor")
        TonaMinEn.Add("D# Minor")
        TonaMinEn.Add("G# Minor")
        TonaMinEn.Add("C# Minor")
        TonaMinEn.Add("F# Minor")
        TonaMinEn.Add("B Minor")
        TonaMinEn.Add("E Minor")
        TonaMinEn.Add("A Minor")
        TonaMinEn.Add("D Minor")
        TonaMinEn.Add("G Minor")
        TonaMinEn.Add("C Minor")
        TonaMinEn.Add("F Minor")


        If LangueIHM = "fr" Then
            i = TonaMajFr.IndexOf(Trim(tonaMaj))
            TonaMin = TonaMinFr(i)
            tbl = Split(TonaMin, " ")
            i = ListNdLatine.IndexOf(tbl(0)) ' les notes des tonalités mineures ne sont jamais en b
            If Clef = "b" Then
                Det_TonaMinCours3 = ListNb(i)
            Else
                Det_TonaMinCours3 = ListNd(i)
            End If
        Else
            i = TonaMajEn.IndexOf(Trim(tonaMaj))
            TonaMin = TonaMinEn(i)
            tbl = Split(TonaMin, " ")
            Det_TonaMinCours3 = tbl(0)

        End If
    End Function
    Public Function Det_TonaMinCours4(tonaMaj As String) As String
        Dim tbl() As String
        Dim i As Integer
        Dim TonaMin As String
        '
        Dim TonaMajEn As New ArrayList
        Dim TonaMajFr As New ArrayList
        Dim TonaMinEn As New ArrayList
        Dim TonaMinFr As New ArrayList

        TonaMajFr.Add("Do# Majeur")
        TonaMajFr.Add("Fa# Majeur")
        TonaMajFr.Add("Si Majeur")
        TonaMajFr.Add("Mi Majeur")
        TonaMajFr.Add("La Majeur")
        TonaMajFr.Add("Ré Majeur")
        TonaMajFr.Add("Sol Majeur")
        TonaMajFr.Add("Do Majeur")
        TonaMajFr.Add("Fa Majeur")
        TonaMajFr.Add("Sib Majeur")
        TonaMajFr.Add("Mib Majeur")
        TonaMajFr.Add("Lab Majeur")
        '
        TonaMinFr.Add("La# Mineur")
        TonaMinFr.Add("Ré# Mineur")
        TonaMinFr.Add("Sol# Mineur")
        TonaMinFr.Add("Do# Mineur")
        TonaMinFr.Add("Fa# Mineur")
        TonaMinFr.Add("Si Mineur")
        TonaMinFr.Add("Mi Mineur")
        TonaMinFr.Add("La Mineur")
        TonaMinFr.Add("Ré Mineur")
        TonaMinFr.Add("Sol Mineur")
        TonaMinFr.Add("Do Mineur")
        TonaMinFr.Add("Fa Mineur")


        TonaMajEn.Add("C# Major")
        TonaMajEn.Add("F# Major")
        TonaMajEn.Add("B Major")
        TonaMajEn.Add("E Major")
        TonaMajEn.Add("A Major")
        TonaMajEn.Add("D Major")
        TonaMajEn.Add("G Major")
        TonaMajEn.Add("C Major")
        TonaMajEn.Add("F Major")
        TonaMajEn.Add("Bb Major")
        TonaMajEn.Add("Eb Major")
        TonaMajEn.Add("Ab Major")


        TonaMinEn.Add("A# Minor")
        TonaMinEn.Add("D# Minor")
        TonaMinEn.Add("G# Minor")
        TonaMinEn.Add("C# Minor")
        TonaMinEn.Add("F# Minor")
        TonaMinEn.Add("B Minor")
        TonaMinEn.Add("E Minor")
        TonaMinEn.Add("A Minor")
        TonaMinEn.Add("D Minor")
        TonaMinEn.Add("G Minor")
        TonaMinEn.Add("C Minor")
        TonaMinEn.Add("F Minor")

        i = TonaMajEn.IndexOf(Trim(tonaMaj))
        TonaMin = TonaMinEn(i)
        tbl = Split(TonaMin, " ")
        Det_TonaMinCours4 = tbl(0)

    End Function
    Public Function Det_IndexLigneColonne(Ligne As Integer, Colonne As Integer, NbCols As Integer) As Integer
        Det_IndexLigneColonne = (Ligne * NbCols) + Colonne
    End Function
    Public Function Det_IndexDegréMouv(Degré As String) As Integer '
        Dim i As Integer
        Det_IndexDegréMouv = -1
        For i = 0 To 6
            If Trim(TabTonsDegrés.Item(i).Text) = Degré Then
                Det_IndexDegréMouv = i
            End If
        Next

    End Function
    Public Function Det_IndexDegré(Degré As String) As Integer '

        Select Case Trim(Degré)
            Case "I"
                Det_IndexDegré = 0
            Case "II"
                Det_IndexDegré = 1
            Case "III"
                Det_IndexDegré = 2
            Case "IV"
                Det_IndexDegré = 3
            Case "V"
                Det_IndexDegré = 4
            Case "VI"
                Det_IndexDegré = 5
            Case "VII"
                Det_IndexDegré = 6
            Case Else
                Det_IndexDegré = -1
        End Select

    End Function


    Public Function Det_IndexDegréMin(Degré As String) As Integer '
        'Dim i As Integer
        'For i = 0 To 6
        ' If Trim(TabTonsDegrés.Item(i).Text) = Degré Then
        ' Det_IndexDegré = i
        ' End If
        ' Next
        Select Case Trim(Degré)
            Case "I"
                Det_IndexDegréMin = 2
            Case "II"
                Det_IndexDegréMin = 3
            Case "III"
                Det_IndexDegréMin = 4 ' E est II en majeur et V en mineur
            Case "IV"
                Det_IndexDegréMin = 5
            Case "V"
                Det_IndexDegréMin = 6
            Case "VI"
                Det_IndexDegréMin = 0
            Case "VII"
                Det_IndexDegréMin = 1
            Case Else
                Det_IndexDegréMin = -1
        End Select
    End Function

    '   Det_IndexDegré2 : détermine le n° de colonne d'un Degré dans TableCoursAcc
    Public Function Det_IndexDegré2(ind As Integer) As Integer ' l Table
        Dim c As Integer
        Dim degré As String

        ' détermination de la colonne
        Select Case ind
            Case 0, 7, 14, 21, 28, 35 ' colonne 1
                c = 0
            Case 1, 8, 15, 22, 29, 36 ' colonne 2
                c = 1
            Case 2, 9, 16, 23, 30, 37 ' colonne 3
                c = 2
            Case 3, 10, 17, 24, 31, 38 ' colonne 4
                c = 3
            Case 4, 11, 18, 25, 32, 39 ' colonne 5
                c = 4
            Case 5, 12, 19, 26, 33, 40 ' colonne 6
                c = 5
            Case 6, 13, 20, 27, 34, 41 ' colonne 7
                c = 6
        End Select
        '
        degré = TabTonsDegrés.Item(c).Text ' attention on est dans TabTons mais ça marche aussi pour pour TabTonsVoisins
        Det_IndexDegré2 = Det_IndexDegré(degré)
    End Function
    Private Function Det_LigneTableGlobale(Index As Integer) As Integer
        Select Case Index
            Case 0, 1, 2, 3, 4, 5, 6
                Det_LigneTableGlobale = 0
            Case 7, 8, 9, 10, 11, 12, 13
                Det_LigneTableGlobale = 1
            Case 14, 15, 16, 17, 18, 19, 20
                Det_LigneTableGlobale = 2
            Case 21, 22, 23, 24, 25, 26, 27
                Det_LigneTableGlobale = 3
            Case 28, 29, 30, 31, 32, 33, 34
                Det_LigneTableGlobale = 4
            Case 35, 36, 37, 38, 39, 40, 41
                Det_LigneTableGlobale = 5
            Case Else
                Det_LigneTableGlobale = -1
        End Select
    End Function
    Private Function Det_IndexDansLigne(ind As Integer) As Integer
        Select Case ind
            Case 0, 1, 2, 3, 4, 5, 6
                Det_IndexDansLigne = ind
            Case 7, 8, 9, 10, 11, 12, 13
                Det_IndexDansLigne = ind - 7
            Case 14, 15, 16, 17, 18, 19, 20
                Det_IndexDansLigne = ind - 14

            Case Else
                Det_IndexDansLigne = -1
        End Select
    End Function
    Private Function Det_IndexDegréI(Index As Integer) As Integer
        Select Case Index
            Case 0, 1, 2, 3, 4, 5, 6
                Det_IndexDegréI = 0
            Case 7, 8, 9, 10, 11, 12, 13
                Det_IndexDegréI = 7
            Case 14, 15, 16, 17, 18, 19, 20
                Det_IndexDegréI = 14
            Case 21, 22, 23, 24, 25, 26, 27
                Det_IndexDegréI = 21
            Case 28, 29, 30, 31, 32, 33, 34
                Det_IndexDegréI = 28
            Case 35, 36, 37, 38, 39, 40, 41
                Det_IndexDegréI = 35
            Case Else
                Det_IndexDegréI = -1
        End Select
    End Function
    Sub EffacerGrid2()
        Dim lig As Integer
        Dim col As Integer
        Grid2.AutoRedraw = False
        '
        For lig = 0 To Grid2.Rows - 1
            For col = 0 To Grid2.Cols - 1
                Grid2.Cell(lig, col).Text = ""
            Next
        Next
        '
        Grid2.Refresh()
        Grid2.AutoRedraw = True
    End Sub
    Sub EffacerTableEventH()
        Dim m, t, ct As Integer

        For m = 0 To UBound(TableEventH, 1)
            For t = 0 To UBound(TableEventH, 2)
                For ct = 0 To UBound(TableEventH, 3)
                    TableEventH(m, t, ct).Ligne = -1
                    TableEventH(m, t, ct).Position = ""
                    TableEventH(m, t, ct).Position = ""
                    TableEventH(m, t, ct).Marqueur = ""
                    TableEventH(m, t, ct).Tonalité = ""
                    TableEventH(m, t, ct).Accord = ""
                    TableEventH(m, t, ct).Gamme = ""
                Next ct
            Next t
        Next m
    End Sub
    Sub EffacerTout()
        EffacerGrid2()
        EffacerTableEventH()
    End Sub

    Sub EcritureMarqueur()
        Dim b As Boolean = Me.KeyPreview
        Dim tbl() As String

        TMarqueur.Visible = False
        If Grid2.ActiveCell.Row = 1 Then ' ligne marqueur
            tbl = TMarqueur.Text.Split(vbCrLf) ' on retire le code de la touche 'entrée'

            Grid2.ActiveCell.Text = Trim(tbl(0)) 'Trim(TMarqueur.Text)
            TableEventH(Grid2.ActiveCell.Col, 1, 1).Marqueur = Trim(tbl(0)) 'Trim(TMarqueur.Text)
            ' Mise à jour entêtes des Piano Roll
            ' **********************************
            For i = 0 To nb_PianoRoll - 1
                If PIANOROLLChargé(i) = True Then
                    If Trim(tbl(0)) <> "" Then
                        listPIANOROLL(i).AddMarq(Trim(tbl(0)), Convert.ToString(Grid2.ActiveCell.Col) + ".1.1")
                    Else
                        listPIANOROLL(i).DeleteMarq(Convert.ToString(Grid2.ActiveCell.Col) + ".1.1")
                    End If
                End If
            Next
            '
            ' Mise à jour de Automation
            ' *************************
            Automation1.PListAcc = Det_ListAcc()
            Automation1.PListMarq = Det_ListMarq()
            Automation1.F4_Refresh()

            ' Mise à jour  de Drumedit (Timeline de drumedit)
            ' ****************************************************
            Drums.PListAcc = Det_ListAcc()
            Drums.PListMarq = Det_ListMarq()
            Drums.F2_Refresh()
            '
            ' Mise à jour de la vue des notes
            ' ******************************
            Maj_VueNotes()
        End If
    End Sub
    Sub EcritureMarqueur2()
        Dim i, p As Integer
        Dim a1, b1, c1, d1 As String
        Dim b As Boolean = Me.KeyPreview
        Dim tbl() As String

        TMarqueur2.Visible = False
        If Grid2.ActiveCell.Row = 3 Then ' ligne Répétition
            If IsNumeric(TMarqueur2.Text) Then
                p = Det_NumDerAccord()
                tbl = TMarqueur2.Text.Split(vbCrLf) ' on retire le code de la touche 'entrée'
                i = Convert.ToInt16(tbl(0))
                If i > nbRépétitionMax Then TMarqueur2.Text = Convert.ToString(nbRépétitionMax)
                p = Det_NumDerAccord()
                If p + (i - 1) <= Module1.nbMesures Then
                    If i > 0 Then
                        SAUV_Annuler(Grid2.ActiveCell.Col, Grid2.ActiveCell.Col)
                        Grid2.ActiveCell.Text = tbl(0)
                        TableEventH(Grid2.ActiveCell.Col, 1, 1).Répet = Convert.ToByte(tbl(0))

                        ' Mise à jour de l'entête des PianoRoll
                        ' *************************************
                        a1 = Trim(Det_ListAcc())
                        b1 = Trim(Det_ListGam())
                        c1 = Trim(Det_ListMarq())
                        d1 = Trim(Det_ListTon())
                        For i = 0 To nb_PianoRoll - 1
                            If PIANOROLLChargé(i) = True Then
                                listPIANOROLL(i).PListAcc = a1 'Det_ListAcc()
                                listPIANOROLL(i).PListGam = b1 'Det_ListGam()
                                listPIANOROLL(i).PListMarq = c1 'Det_ListMarq()
                                listPIANOROLL(i).PListTon = d1
                                listPIANOROLL(i).Clear_graphique2() ' <----
                                listPIANOROLL(i).F1_Refresh()
                                listPIANOROLL(i).Maj_CalquesMIDI()
                            End If
                        Next
                        '
                        Automation1.PListAcc = Det_ListAcc()
                        Automation1.PListMarq = Det_ListMarq()
                        Automation1.F4_Refresh()

                        Drums.PListAcc = Det_ListAcc()
                        Drums.PListMarq = Det_ListMarq()
                        Drums.F2_Refresh()
                        '
                        Maj_VueNotes()
                    End If
                End If
            End If
        End If
        'End If
    End Sub

    Private Sub Tempo_ValueChanged(sender As Object, e As EventArgs) Handles Tempo.ValueChanged
        Horloge1.BeatsPerMinute = Tempo.Value
        TextBox2.Text = Calcul_Durée()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Répéter_CheckedChanged(sender As Object, e As EventArgs) Handles Répéter.CheckedChanged

    End Sub

    Private Sub MenuExportsMIDI_Click(sender As Object, e As EventArgs) Handles MenuExportsMIDI.Click
        Cacher_FormTransparents()
        MessageHV.PTypBouton = "OuiNon"
        If LangueIHM = "fr" Then
            MessageHV.PTitre = "Export CTRL"
            MessageHV.PContenuMess = "Souhaitez-vous exporter les contrôleurs MIDI (Volume, Pan, PRG) dans le fichier MIDI ?" +
                 Constants.vbCrLf + Constants.vbCrLf + "Remarque importante : le fichier MIDI est créé entre le délimiteur de début " + Début.Text + " et le délimiteur de fin " + Terme.Text + "."
        Else
            MessageHV.PTitre = "CTRL Export"
            MessageHV.PContenuMess = "Do you wish to export MIDI controlers (Volume, Pan, PRG) in the MIDI file  ?" +
                 Constants.vbCrLf + Constants.vbCrLf + "Important note: the MIDI file is created between the start locator " + Début.Text + " and the end locator " + Terme.Text + "."
        End If
        Cacher_FormTransparents()
        MessageHV.ShowDialog()
        If MessageHV.PSortie = "Oui" Then
            ExporterCTRL = True
        Else
            ExporterCTRL = False
        End If
        ExportMIDI()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub Tempo_Aff_EventH_Tick(sender As Object, e As EventArgs) Handles Tempo_Aff_EventH.Tick
        Dim p As Integer
        Dim a As String

        ' affichage des accords
        If CallB_Aff_Acc Then
            PointAffAccord = PointAffAccord + 1
            p = Arrangement1.NumAccords.PointeurLect
            a = Arrangement1.NumAccords.NumAcc(p)
            '
            Label16.Text = Grid2.Cell(2, Val(a)).Text ' Accord dans vue standard
            Label17.Text = Grid2.Cell(2, Val(a)).Text ' Accord dans vue standard
            '
            CallB_Aff_Acc = False
            Arrangement1.NumAccords.PointeurLect = Arrangement1.NumAccords.PointeurLect + 1
        End If

        ' affichage des N° de mesures
        If CallB_Aff_NumMes = True Then
            a = ListNumMes(AffNumMes)
            AffNumMes = (AffNumMes + 1)
            '
            If Transport.Visible = True Then
                Transport.Label1.Text = a
            End If
            Label6.Text = a
            Label31.Text = a 'Convert.ToString(AffNumMes) 'a  ' N° accord dans vue admin
            Label15.Text = a 'Convert.ToString(AffNumMes) ' a  ' N° accord dans vue standard
            CallB_Aff_NumMes = False
        End If
        If CallB_Aff_FIN Then
            StopPlay()
        End If
    End Sub
    Function Det_Numacc() As String
        Det_Numacc = ""
        NumAcc = NumAcc + 1
        For i = NumAcc To nbMesures '- 1
            If Grid2.Cell(2, i).Text <> "" Then
                Det_Numacc = i
                NumAcc = i
                Exit For
            End If
        Next
    End Function

    Private Sub Button4_Click_1(sender As Object, e As EventArgs)
        Dim i As Integer
        i = Det_IndexGénérateur(2, 3)
        i = Det_IndexPisteMidi(21)
    End Sub

    Private Sub Button5_Click_1(sender As Object, e As EventArgs)
        Calcul_AutoVoicingZ()
        CalculArp(False)
        VisuN.ShowDialog()
    End Sub

    Private Sub TMarqueur_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim i As Integer
        Dim a As String
        Dim tona As String
        Dim tonaMin As String
        Dim tbl() As String
        '
        tbl = Split(Trim(ComboBox1.Text))

        If LangueIHM = "fr" Then
            Clef = Det_ClefLat(Trim(tbl(0)))
        Else
            Clef = Det_Clef(Trim(tbl(0)))
        End If
        '

        If EnChargement = False Then ' 
            ' Maj du combo de la relative mineure
            ComboBox2.SelectedIndex = ComboBox1.SelectedIndex
            ComboBox2.Refresh()
            ' Mise à jour onglet Modes (Tabage1)
            ' **********************************
            'tona = Det_TonaCours2()
            tbl = Split(Trim(ComboBox1.Text))
            tona = Trim(tbl(0))
            tonaMin = Det_TonaMinCours4(ComboBox1.Text) '
            '
            tonaMin = TradNote_AnglMinMaj(tonaMin)
            Maj_TableGlobalAcc(Trim(tona), Trim(tonaMin)) ' TableGlobaleAcc contient tous les accords de 3, 4 et 5 notes pour la tonalité choisie
            '
            Maj_TabTons(Trim(ComboBox23.SelectedIndex)) ' maj du contenu de l'onglet TabPage1 (Modes)
            ComboBox2.Refresh()
            '
            ' Mise à jour Modulation
            ' ***********************        '
            'Maj_ModulationRadioB()
            'Maj_Modulation()

            Refresh()
            CAD_Maj_TableGlobalAcc()
            Refresh()

            RAZ_CouleurMarquée()
            '
            a = Trim(Det_TonaCours2())
            Entrée_Tonalité = Trim(a + " " + "Maj")
            '
            ' Mise à jour des couleurs types des tonalités dans les tableaux des ton et le tableau des progressions
            For i = 0 To 6
                TabTons.Item(i).BackColor = DicoCouleur.Item(a) ' couleurs tons
            Next
            For i = 0 To 4
                TabCadDegrés.Item(i).BackColor = DicoCouleur.Item(a) ' couleurs progressions
            Next

            Refresh()
        Else
            '
            Maj_TableGlobalAcc("C", "A") ' valeurs par défaut au démarrage
            Maj_TabTons(0)
            Maj_TableCoursAccInit()
            CAD_Maj_TableCoursAccInit()

            ' Mise à jour Modulation
            ' **********************        '
            'Maj_ModulationRadioB()
            'For i = 0 To 6
            'Grid5.Cell(1, i + 1).Text = TableGlobalAcc(0, 0, i) 'tbl1(i) ' mise à jour de la gamme de C Maj dans l'onglet Modulation
            'Next i
        End If
        'End If
    End Sub


    Private Sub ComboBox2_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        '
        If EnChargement = False Then
            ComboBox1.SelectedIndex = ComboBox2.SelectedIndex
        Else
            ComboBox1.SelectedIndex = ComboBox2.SelectedIndex
        End If
        'End If
    End Sub
    Private Sub TimerToutesNotesOff_Tick(sender As Object, e As EventArgs) Handles TimerToutesNotesOff.Tick
        ToutesNotesOff()
        TimerToutesNotesOff.Stop()
        'PlayMidi.Enabled = True
    End Sub


    Private Sub ButCopy_Click(sender As Object, e As EventArgs) Handles ButCopy.Click
        Dim pst As Integer
        Dim GénéOrig, GénéDest As Integer
        Dim Orig, Dest As Integer
        '


        If ComboCopy1.SelectedIndex <> ComboCopy2.SelectedIndex Then
            Orig = ComboCopy1.SelectedIndex
            Dest = ComboCopy2.SelectedIndex
            '
            MessageHV.PTypBouton = "OuiNon"
            If LangueIHM = "fr" Then
                MessageHV.PTitre = "Copie de Variations"
                MessageHV.PContenuMess = "Confirmez-vous la copie de Variation " + Str(Orig + 1) + " vers Variation " + Str(Dest + 1) + " ?"
            Else
                MessageHV.PTitre = "Variations copy"
                MessageHV.PContenuMess = "Do you confirm the copy from Variation " + Str(Orig + 1) + " to Variation " + Str(Dest + 1) + " ?"
            End If
            Cacher_FormTransparents()
            MessageHV.ShowDialog()
            If MessageHV.PSortie = "Oui" Then
                GénéOrig = Det_IndexGénérateur(pst, Orig)
                GénéDest = Det_IndexGénérateur(pst, Dest)
                For pst = 0 To nb_PistesVar - 1 ' pst contient les index des pistes des variations

                    ' copie 
                    SelBloc(GénéDest + pst).Checked = SelBloc(GénéOrig + pst).Checked
                    PisteMotif(GénéDest + pst).SelectedIndex = PisteMotif(GénéOrig + pst).SelectedIndex
                    PisteDurée(GénéDest + pst).SelectedIndex = PisteDurée(GénéOrig + pst).SelectedIndex
                    PisteOctave(GénéDest + pst).SelectedIndex = PisteOctave(GénéOrig + pst).SelectedIndex

                    PisteDyn(GénéDest + pst).SelectedIndex = PisteDyn(GénéOrig + pst).SelectedIndex
                    PisteAccent(GénéDest + pst).SelectedIndex = PisteAccent(GénéOrig + pst).SelectedIndex
                    PistePRG(GénéDest + pst).SelectedIndex = PistePRG(GénéOrig + pst).SelectedIndex
                    PisteRadio1(GénéDest + pst).Checked = PisteRadio1(GénéOrig + pst).Checked
                    PisteRadio2(GénéDest + pst).Checked = PisteRadio2(GénéOrig + pst).Checked
                    PisteRadio3(GénéDest + pst).Checked = PisteRadio3(GénéOrig + pst).Checked
                    PisteSouche(GénéDest + pst).SelectedIndex = PisteSouche(GénéOrig + pst).SelectedIndex
                    PisteRetard(GénéDest + pst).SelectedIndex = PisteRetard(GénéOrig + pst).SelectedIndex
                    'PisteDelay(GénéDest + pst).Checked = PisteDelay(GénéOrig + pst).Checked ' pas de copie de DElay car delay n'existe plus que sur le bloc 1
                    PisteDébut(GénéDest + pst).Checked = PisteDébut(GénéOrig + pst).Checked
                Next
            End If
        End If
    End Sub

    Private Sub EffacerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EffacerToolStripMenuItem.Click
        Dim a, b, c, d As String
        Select Case OngletCours_Edition
            Case 0 ' HyperArp
                If Grid2.Selection.FirstCol > 1 Then
                    Grid2.Range(1, 1, 12, Grid2.Cols - 1).Locked = False
                    SAUV_Annuler(Grid2.Selection.FirstCol, Grid2.Selection.LastCol)
                    'CopierClip()
                    Grid2.Selection.CutData()
                    CouperClip()
                    Maj_Répétition2(Grid2.Selection.LastCol)
                    Calcul_AutoVoicingZ()
                    '
                    ' Mise à jour PianoRoll
                    ' *********************
                    a = Trim(Det_ListAcc())
                    b = Trim(Det_ListGam())
                    c = Trim(Det_ListMarq())
                    d = Trim(Det_ListTon())
                    For i = 0 To nb_PianoRoll - 1
                        If PIANOROLLChargé(i) = True Then
                            listPIANOROLL(i).PListAcc = a 'Det_ListAcc()
                            listPIANOROLL(i).PListGam = b 'Det_ListGam()
                            listPIANOROLL(i).PListMarq = c 'Det_ListMarq()
                            listPIANOROLL(i).PListTon = d
                            listPIANOROLL(i).Clear_graphique2() ' <---
                            listPIANOROLL(i).F1_Refresh()
                            listPIANOROLL(i).Maj_CalquesMIDI()
                        End If
                    Next
                    '
                    Grid2.Range(1, 1, 12, Grid2.Cols - 1).Locked = True
                    '
                    Automation1.PListAcc = Det_ListAcc()
                    Automation1.PListMarq = Det_ListMarq()
                    Automation1.F4_Refresh()

                    Drums.PListAcc = Det_ListAcc()
                    Drums.PListMarq = Det_ListMarq()
                    Drums.F2_Refresh()
                Else
                    Dim message, titre As String
                    If LangueIHM = "fr" Then
                        message = "Il n'est pas possible de supprimer la première colonne."
                        titre = "Avertissement"
                    Else
                        message = "It is not possible to delete the first column."
                        titre = "Avertissement"
                    End If
                    Cacher_FormTransparents()
                    MessageBox.Show(message, titre, MessageBoxButtons.OK, MessageBoxIcon.Information)

                End If
            Case Else
                ' dans les autres objet,  la capture de la touche suppr est traitée directement dans KeyUp de l'objet concerné
        End Select
    End Sub
    Class Accgrid2
        Public Marqueur As String
        Public Accord As String
        Public Répet As String
        Public Magneto As Integer
        Public Gamme As String
        Public PosOld As Integer
    End Class
    Sub AccConcentrer()
        Dim i As Integer
        Dim j As Integer
        Dim AA As String

        Dim liste1 As New List(Of Accgrid2)
        Dim liste2 As New List(Of EventH)
        '
        Grid2.Refresh()
        ' sauvegarder
        For i = 1 To nbMesures
            If Trim(Grid2.Cell(2, i).Text) <> "" Then
                Dim oo1 As New Accgrid2 With {
                    .Marqueur = Trim(Grid2.Cell(1, i).Text),
                    .Accord = Trim(Grid2.Cell(2, i).Text),
                    .Magneto = IdentVariation(i),
                    .PosOld = i
                }
                liste1.Add(oo1)

                '
                Dim oo2 As New EventH With {
                .Accord = TableEventH(i, 1, 1).Accord,
            .Degré = TableEventH(i, 1, 1).Degré,
            .Détails = TableEventH(i, 1, 1).Détails,
            .Gamme = TableEventH(i, 1, 1).Gamme,
            .Ligne = TableEventH(i, 1, 1).Ligne,
            .Marqueur = TableEventH(i, 1, 1).Marqueur,
            .Mode = TableEventH(i, 1, 1).Mode,
            .NumAcc = i,
            .NumMagnéto = TableEventH(i, 1, 1).NumMagnéto,
            .Position = Trim(Str(i) + "." + "1" + "." + "1")
            }
                liste2.Add(oo2)
            End If
        Next
        ' effacer
        EffacerGrid2_4()
        ' replacer
        For Each a As Accgrid2 In liste1
            j = liste1.IndexOf(a) + 1
            'marqueur
            Grid2.Cell(1, j).Text = a.Marqueur
            'Accord
            Grid2.Cell(2, j).Text = a.Accord
            AA = TableEventH(j, 1, 1).Tonalité '
            AA = Det_RelativeMajeure2(AA) ' si la tonalité est mineure alors on affiche la couleur de la relative majeure
            tbl = Split(AA)
            Grid2.Cell(2, j).BackColor = DicoCouleur.Item(Trim(tbl(0))) ' la couleur est fonction de la tonalité
            Grid2.Cell(2, j).ForeColor = DicoCouleurLettre.Item(tbl(0))
            ' Magnéto
            ChoixVariationGrid2(j, j, a.Magneto + 4)
        Next
        '
        For Each a As EventH In liste2
            i = liste2.IndexOf(a)
            TableEventH(i + 1, 1, 1).Accord = liste2.Item(i).Accord
            TableEventH(i + 1, 1, 1).Degré = liste2.Item(i).Degré
            TableEventH(i + 1, 1, 1).Détails = liste2.Item(i).Détails
            TableEventH(i + 1, 1, 1).Gamme = liste2.Item(i).Gamme
            TableEventH(i + 1, 1, 1).Ligne = liste2.Item(i).Ligne
            TableEventH(i + 1, 1, 1).Marqueur = liste2.Item(i).Marqueur
            TableEventH(i + 1, 1, 1).Mode = liste2.Item(i).Mode
            TableEventH(i + 1, 1, 1).NumAcc = liste2.Item(i).NumAcc
            TableEventH(i + 1, 1, 1).NumMagnéto = liste2.Item(i).NumMagnéto
            TableEventH(i + 1, 1, 1).Position = liste2.Item(i).Position
        Next
        '
        For i = i + 1 To nbMesures
            TableEventH(i, 1, 1).Ligne = -1
            TableEventH(i, 1, 1).Accord = ""
        Next

    End Sub

    Sub EffacerGrid2_4()
        Dim i, j As Integer
        Grid2.AutoRedraw = False

        For i = 0 To Grid2.Cols - 1
            Grid2.Column(i).Locked = False
        Next i
        '
        For j = 1 To Grid2.Cols - 1
            Grid2.Cell(1, j).Text = ""
            Grid2.Cell(2, j).Text = ""
            Grid2.Cell(2, j).BackColor = Color.White
        Next

        '
        'For i = 0 To Grid2.Cols - 1
        ' RAZ 1er ligne de grid2 pour le cas où il y avait des marqeurs
        'Grid2.Cell(0, i).BackColor = Color.Beige
        'Grid2.Cell(0, i).ForeColor = Color.Black
        ''
        '
        'Grid2.Cell(0, i).Locked = True
        'Next i
        '
        Grid2.AutoRedraw = True
        Grid2.Refresh()
    End Sub

    Private Sub TabPage3_Leave(sender As Object, e As EventArgs) Handles TabPage3.Leave
        'Définir.Checked = True
    End Sub
    Private Sub Grid2_DoubleClick(Sender As Object, e As EventArgs) Handles Grid2.DoubleClick
        'Ouverture_PianoRoll()
    End Sub

    Sub Ouverture_PianoRoll2(n_PRoll As Integer)
        Dim i As Integer = n_PRoll
        Dim a As String
        Me.Cursor = Cursors.WaitCursor
        listPIANOROLL(i).PStartMeasure = Grid2.ActiveCell.Col '

        a = Det_ListAcc()
        If Trim(a) <> "" Then
            listPIANOROLL(i).PListAcc = a
            listPIANOROLL(i).PListGam = Det_ListGam()
            '
            PIANOROLLChargé(i) = True
            listPIANOROLL(i).F1.StartPosition = FormStartPosition.CenterScreen
            listPIANOROLL(i).PStartMeasure = Grid2.ActiveCell.Col
            '
            listPIANOROLL(i).F1_Refresh()
            '
            'listPIANOROLL(i).F1.TopMost = True
            listPIANOROLL(i).F1.Show()
            listPIANOROLL(i).Graphique_Position() ' doit être placé après le "Show"
            '
        End If
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub Form1_Move(sender As Object, e As EventArgs) Handles Me.Move
        'PianoRoll.F2.Top = Me.Top + 78
        'PianoRoll.F2.Left = Me.Left
    End Sub
    ' Actions de transposition à partir du ContextMenuStrip3
    ' ******************************************************
    Private Sub Transpo1_Click(sender As Object, e As EventArgs)
        SAUV_Annuler(Grid2.Selection.FirstCol, Grid2.Selection.LastCol)
        Transposer(1, Grid2.Selection.FirstCol, Grid2.Selection.LastCol)
    End Sub
    Private Sub Transpo2_Click(sender As Object, e As EventArgs)
        SAUV_Annuler(Grid2.Selection.FirstCol, Grid2.Selection.LastCol)
        Transposer(2, Grid2.Selection.FirstCol, Grid2.Selection.LastCol)
    End Sub
    Private Sub Transpo3_Click(sender As Object, e As EventArgs)
        SAUV_Annuler(Grid2.Selection.FirstCol, Grid2.Selection.LastCol)
        Transposer(3, Grid2.Selection.FirstCol, Grid2.Selection.LastCol)
    End Sub
    Private Sub Transpo4_Click(sender As Object, e As EventArgs)
        SAUV_Annuler(Grid2.Selection.FirstCol, Grid2.Selection.LastCol)
        Transposer(4, Grid2.Selection.FirstCol, Grid2.Selection.LastCol)
    End Sub
    Private Sub Transpo5_Click(sender As Object, e As EventArgs)
        SAUV_Annuler(Grid2.Selection.FirstCol, Grid2.Selection.LastCol)
        Transposer(5, Grid2.Selection.FirstCol, Grid2.Selection.LastCol)
    End Sub
    Private Sub Transpo6_Click(sender As Object, e As EventArgs)
        SAUV_Annuler(Grid2.Selection.FirstCol, Grid2.Selection.LastCol)
        Transposer(6, Grid2.Selection.FirstCol, Grid2.Selection.LastCol)
    End Sub
    Private Sub Transpo7_Click(sender As Object, e As EventArgs)
        SAUV_Annuler(Grid2.Selection.FirstCol, Grid2.Selection.LastCol)
        Transposer(7, Grid2.Selection.FirstCol, Grid2.Selection.LastCol)
    End Sub
    Private Sub Transpo8_Click(sender As Object, e As EventArgs)
        Transposer(8, Grid2.Selection.FirstCol, Grid2.Selection.LastCol)
    End Sub
    Private Sub Transpo9_Click(sender As Object, e As EventArgs)
        SAUV_Annuler(Grid2.Selection.FirstCol, Grid2.Selection.LastCol)
        Transposer(9, Grid2.Selection.FirstCol, Grid2.Selection.LastCol)
    End Sub
    Private Sub Transpo10_Click(sender As Object, e As EventArgs)
        Transposer(10, Grid2.Selection.FirstCol, Grid2.Selection.LastCol)
    End Sub
    Private Sub Transpo11_Click(sender As Object, e As EventArgs)
        SAUV_Annuler(Grid2.Selection.FirstCol, Grid2.Selection.LastCol)
        Transposer(11, Grid2.Selection.FirstCol, Grid2.Selection.LastCol)
    End Sub
    Sub Transposer(Num As Integer, FirstC As Integer, LastC As Integer)
        Dim i, j As Integer
        Dim a, b, d As String
        Dim NoteM As String
        Dim NoteT As String
        Dim ListNotes As New List(Of String)
        Dim tbl() As String
        Dim tbl1() As String

        ' Maj de la liste
        ListNotes.Add("C")
        ListNotes.Add("C#")
        ListNotes.Add("D")
        ListNotes.Add("D#")
        ListNotes.Add("E")
        ListNotes.Add("F")
        ListNotes.Add("F#")
        ListNotes.Add("G")
        ListNotes.Add("G#")
        ListNotes.Add("A")
        ListNotes.Add("A#")
        ListNotes.Add("B")
        '
        Harmonie.Maj_ListNotesMajd()

        For i = FirstC To LastC
            a = Trim(Grid2.Cell(2, i).Text)
            If a <> "" Then
                ' accord
                tbl = a.Split(" ")
                tbl(0) = TradD(tbl(0))
                j = ListNotes.IndexOf(tbl(0))
                tbl(0) = ListNotesMajd.Item(j + Num)
                Grid2.Cell(2, i).Text = Join(tbl)
                TableEventH(i, 1, 1).Accord = Trim(Grid2.Cell(2, i).Text)
                ' mode 
                tbl = TableEventH(i, 1, 1).Mode.Split(" ")
                tbl(0) = TradD(tbl(0))
                j = ListNotes.IndexOf(tbl(0))
                NoteM = ListNotesMajd.Item(j + Num)
                TableEventH(i, 1, 1).Mode = Trim(NoteM + " " + tbl(1))
                ' gamme
                tbl = TableEventH(i, 1, 1).Gamme.Split(" ")
                tbl(0) = TradD(tbl(0))
                j = ListNotes.IndexOf(tbl(0))
                NoteM = ListNotesMajd.Item(j + Num)
                TableEventH(i, 1, 1).Gamme = Trim(NoteM + " " + tbl(1))
                Grid2.Cell(11, i).Text = Trim(TableEventH(i, 1, 1).Gamme)
                ' tonalité
                tbl = TableEventH(i, 1, 1).Tonalité.Split(" ")
                tbl(0) = TradD(tbl(0))
                j = ListNotes.IndexOf(tbl(0))
                NoteT = ListNotesMajd.Item(j + Num)
                TableEventH(i, 1, 1).Tonalité = Trim(NoteT + " " + tbl(1))
                ' mise à jour de la couleur en fonction de la tonalité
                b = TableEventH(i, 1, 1).Tonalité '
                b = Det_RelativeMajeure2(b) ' si la tonalité est mineure alors on affiche la couleur de la relative majeure
                tbl1 = Split(b)
                Grid2.Cell(2, i).BackColor = DicoCouleur.Item(Trim(tbl1(0))) ' la couleur de l'accord est fonction de la tonalité
                Grid2.Cell(2, i).ForeColor = DicoCouleurLettre.Item(tbl1(0))
                Grid2.Cell(11, i).BackColor = DicoCouleur.Item(Trim(tbl1(0))) ' la couleur de la gamme est fonction de la tonalité
                Grid2.Cell(11, i).ForeColor = DicoCouleurLettre.Item(tbl1(0))
                '
                ' correction éventuelle si la tonalité n'est pas exprimée ave le bon signe
                If TableEventH(i, 1, 1).Tonalité <> "C Maj" Then Det_Alter(i)
            End If

        Next
        ' Mise à jour des pianoRoll avec les nouvelles données d'accord et de gammes
        ' **************************************************************************
        a = Det_ListAcc()
        b = Det_ListGam()
        d = Trim(Det_ListTon())
        For i = 0 To nb_PianoRoll - 1
            If PIANOROLLChargé(i) = True Then
                listPIANOROLL(i).PListAcc = Trim(a) 'Det_ListAcc()
                listPIANOROLL(i).PListGam = Trim(b) 'Det_ListGam()
                listPIANOROLL(i).PListTon = d
                listPIANOROLL(i).F1_Refresh()
            End If
        Next
    End Sub
    Sub Transposer3(Num As Integer, FirstC As Integer, LastC As Integer)
        Dim i As Integer
        Dim a, b, d As String
        Dim ListNotes As New List(Of String)
        Dim tbl() As String
        Dim tbl1() As String


        If Num <> 0 Then
            For i = FirstC To LastC
                a = Trim(Grid2.Cell(2, i).Text)
                If a <> "" Then
                    ' accord
                    tbl = a.Split(" ")
                    tbl(0) = Trim(Transposer_Note(Num, tbl(0)))
                    Grid2.Cell(2, i).Text = Join(tbl) 'Join(tbl)
                    TableEventH(i, 1, 1).Accord = Trim(Grid2.Cell(2, i).Text)
                    ' mode 
                    tbl = TableEventH(i, 1, 1).Mode.Split(" ")
                    tbl(0) = Trim(Transposer_Note(Num, tbl(0)))
                    TableEventH(i, 1, 1).Mode = Join(tbl) 'Trim(Transposer_Note(Num, tbl(0)) + " " + tbl(1))
                    ' gamme
                    tbl = TableEventH(i, 1, 1).Gamme.Split(" ")
                    tbl(0) = Trim(Transposer_Note(Num, tbl(0)))
                    TableEventH(i, 1, 1).Gamme = Join(tbl) 'Trim(Transposer_Note(Num, tbl(0)) + " " + tbl(1))
                    Grid2.Cell(11, i).Text = Trim(TableEventH(i, 1, 1).Gamme)
                    ' tonalité
                    tbl = TableEventH(i, 1, 1).Tonalité.Split(" ")
                    tbl(0) = Trim(Transposer_Note(Num, tbl(0)))
                    TableEventH(i, 1, 1).Tonalité = Join(tbl) 'Trim(Transposer_Note(Num, tbl(0)) + " " + tbl(1))
                    '
                    ' mise à jour de la couleur en fonction de la tonalité
                    b = TableEventH(i, 1, 1).Tonalité '
                    b = Det_RelativeMajeure2(b) ' si la tonalité est mineure alors on affiche la couleur de la relative majeure
                    tbl1 = Split(b)
                    Grid2.Cell(2, i).BackColor = DicoCouleur.Item(Trim(tbl1(0))) ' la couleur de l'accord est fonction de la tonalité
                    Grid2.Cell(2, i).ForeColor = DicoCouleurLettre.Item(tbl1(0))
                    Grid2.Cell(11, i).BackColor = DicoCouleur.Item(Trim(tbl1(0))) ' la couleur de la gamme est fonction de la tonalité
                    Grid2.Cell(11, i).ForeColor = DicoCouleurLettre.Item(tbl1(0))
                    '
                    ' correction éventuelle si la tonalité n'est pas exprimée ave le bon signe
                    If TableEventH(i, 1, 1).Tonalité <> "C Maj" Then Det_Alter(i)
                End If

            Next
            ' Mise à jour des pianoRoll avec les nouvelles données d'accord et de gammes
            ' **************************************************************************
            a = Det_ListAcc()
            b = Det_ListGam()
            d = Trim(Det_ListTon())
            For i = 0 To nb_PianoRoll - 1
                If PIANOROLLChargé(i) = True Then
                    listPIANOROLL(i).PListAcc = Trim(a) 'Det_ListAcc()
                    listPIANOROLL(i).PListGam = Trim(b) 'Det_ListGam()
                    listPIANOROLL(i).PListTon = d
                    listPIANOROLL(i).F1_Refresh()
                    listPIANOROLL(i).Maj_CalquesMIDI()
                End If
            Next
        End If

    End Sub
    Function Transposer_Note(num As Integer, nomNote As String) As String
        Dim ListNotesPlus As New List(Of String)
        Dim ListNotesMoins As New List(Of String)
        Dim r As String = "C"
        Dim j As Integer

        nomNote = Trim(nomNote)

        ListNotesPlus.Add("C")
        ListNotesPlus.Add("C#")
        ListNotesPlus.Add("D")
        ListNotesPlus.Add("D#")
        ListNotesPlus.Add("E")
        ListNotesPlus.Add("F")
        ListNotesPlus.Add("F#")
        ListNotesPlus.Add("G")
        ListNotesPlus.Add("G#")
        ListNotesPlus.Add("A")
        ListNotesPlus.Add("A#")
        ListNotesPlus.Add("B")
        '
        ListNotesMoins.Add("B")
        ListNotesMoins.Add("A#")
        ListNotesMoins.Add("A")
        ListNotesMoins.Add("G#")
        ListNotesMoins.Add("G")
        ListNotesMoins.Add("F#")
        ListNotesMoins.Add("F")
        ListNotesMoins.Add("E")
        ListNotesMoins.Add("D#")
        ListNotesMoins.Add("D")
        ListNotesMoins.Add("C#")
        ListNotesMoins.Add("C")
        '
        Harmonie.Maj_ListNotesMajd()

        '
        If nomNote <> "" Then
            ' accord
            nomNote = TradD(nomNote)
            If num > 0 Then
                j = ListNotesPlus.IndexOf(nomNote)
                r = ListNotesMajd.Item(j + num)
            Else
                j = ListNotesMoins.IndexOf(nomNote)
                r = ListNotesMajd_Inv.Item(j - num)
            End If
            '
            Return r
        End If
        Return r
    End Function
    Sub Det_Alter(m As Integer)
        Dim signe As String = "b"
        Dim tbl() As String

        ' TONALITE : détermination des modifications à apporter sur la tonalié : la tonalité est toujours majeur (nombre d'accidents à la clef)
        tbl = TableEventH(m, 1, 1).Tonalité.Split(" ")
        If listMaj.IndexOf(tbl(0)) = -1 Then
            tbl(0) = Dico_MajHorsClef(tbl(0)) ' correction de la tonalité si nécessaire par exemple D# Maj devient Eb Maj
            TableEventH(m, 1, 1).Tonalité = Join(tbl)
        End If
        If listMaj.IndexOf(tbl(0)) <= 6 Then signe = "#" ' détermination du signe d"altération à utiliser pour l'expression des accords

        ' MODE : détermination des modifications à apporter sur le mode : le mode peut être majeur ou mineur
        tbl = TableEventH(m, 1, 1).Mode.Split(" ")
        If tbl(1) = "Maj" Then
            If listMaj.IndexOf(tbl(0)) = -1 Then
                tbl(0) = Dico_MajHorsClef(tbl(0)) ' correction du mode si nécessaire par exemple D# Maj devient Eb Maj
                TableEventH(m, 1, 1).Mode = Join(tbl)
            End If
        Else
            If listMin.IndexOf(tbl(0)) = -1 Then
                tbl(0) = Dico_MinHorsClef(tbl(0))
                TableEventH(m, 1, 1).Mode = Join(tbl)
            End If
        End If
        '
        ' ACCORD : détermination des modifications à apporter sur l'accord en fonction du signe
        tbl = TableEventH(m, 1, 1).Accord.Split(" ")
        tbl(0) = TradDB(tbl(0), signe)
        TableEventH(m, 1, 1).Accord = Join(tbl)
        Grid2.Cell(2, m).Text = Trim(TableEventH(m, 1, 1).Accord)

        ' GAMME  : détermination des modifications à apporter sur la gamme en fonction du signe
        tbl = TableEventH(m, 1, 1).Gamme.Split(" ")
        tbl(0) = TradDB(tbl(0), signe)
        TableEventH(m, 1, 1).Gamme = Join(tbl)
        Grid2.Cell(11, m).Text = Trim(TableEventH(m, 1, 1).Gamme)
    End Sub
    Private Sub GamesPRO_Click(sender As Object, e As EventArgs) Handles GammesPRO.Click
        'If GammesPRO.Checked Then
        'My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CalquesMIDI\HyperArp\Préférences", "GammesPRO", "0")
        'GammesPRO.Checked = False
        'If Module1.LangueIHM = "fr" Then
        'GammesPRO.Text = "Passer en Gammes PRO"
        'Else
        'GammesPRO.Text = "Switch to PRO scales"
        'End If
        '
        'Button25.BackColor = Color.Beige
        'Button25.ImageAlign = ContentAlignment.MiddleCenter
        'Button25.Image = HyperArp.My.Resources.Resources._32x32_Transp_vert__échelle3_
        'GammesPRO.Image = HyperArp.My.Resources.Resources._32x32_Transp_vert__échelle3_
        'Else
        'My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CalquesMIDI\HyperArp\Préférences", "GammesPRO", "1")
        'GammesPRO.Checked = True
        'If Module1.LangueIHM = "fr" Then
        'GammesPRO.Text = "Passer en Gammes simples"
        'Else
        'GammesPRO.Text = "Switch to Simple scales"
        'End If
        'GammesPRO.Image = HyperArp.My.Resources.Resources._32x32___échelle2_Blues

        'Button25.Image = HyperArp.My.Resources.Resources._32x32___échelle2_Blues
        'Button25.ImageAlign = ContentAlignment.MiddleRight
        'Button25.BackColor = Color.DarkOliveGreen
        'End If
    End Sub

    Private Sub NotesViewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NotesViewToolStripMenuItem.Click
        Calcul_AutoVoicingZ()
        CalculArp(False)
        Cacher_FormTransparents()
        VisuN.ShowDialog()
    End Sub

    Private Sub MIDIResetToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles MIDIResetToolStripMenuItem1.Click
        MIDIReset()
    End Sub




    Public Sub Send_AllVolumes() ' ne pas supprimer : utilisé par la nouvelle table de misage
        If EnChargement = False Then
            Dim pst As Integer
            Dim volume As Byte
            Dim canal As Byte
            '
            If Not (SortieMidi.Item(ChoixSortieMidi).IsOpen) Then
                SortieMidi.Item(ChoixSortieMidi).Open()
            End If

            For pst = 0 To NombrePistes - 1 ' 
                volume = CByte(Mix.PisteVolume.Item(pst).Value)
                canal = CByte(LesPistes.Item(pst).Canal)
                SortieMidi.Item(ChoixSortieMidi).SendControlChange(canal, 7, volume)
            Next
        End If

    End Sub

    Private Sub Form1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Me.KeyPress
        'If TMarqueur.Visible = False And Drums.PEcrNomPerso_Visible = False Then ' lors de l'ecrituree d'un marqueur on peut utiliser la barre d'espace qui alors ne doit pas déclencher le jeu (PlayMidi)
        ' Remarque importante : pour que ça fonctionne ici, il faut que la propriété de form1 KeyPreview =  true
        'If e.KeyChar.ToString() = " " Then ' barre espace
        'If PlayMidi.Enabled = True Then
        'PlayMidi.Focus()
        'Else
        'StopMidi.Focus()
        'End If
        'End If
        'If PlayMidi.Enabled = True And e.KeyChar.ToString() = " " Then
        'PlayHyperArp()
        'Else
        '   If PlayMidi.Enabled = False And e.KeyChar.ToString() = " " Then
        '  StopPlay()
        'End If
        'End If
        'End If

    End Sub

    Sub JouerNote2(ValeurNote As Byte, canal As Byte, Dyn As Byte)

        Try
            If Not (SortieMidi.Item(ChoixSortieMidi).IsOpen) Then
                SortieMidi.Item(ChoixSortieMidi).Open()
            End If
            '
            SortieMidi.Item(ChoixSortieMidi).SendNoteOn(canal - 1, ValeurNote, Dyn)
            NoteCourante = ValeurNote
            NoteAEtéJouée = True
        Catch ex As Exception
            MessageHV.TypBouton = "OK"
            If Module1.LangueIHM = "fr" Then
                MessageHV.PTitre = "Problème de ressource MIDI"
                MessageHV.PContenuMess = "Warning : détection d'une erreur dans procédure " + "JouerNote2" +
                Constants.vbCrLf + "- " + "Message  : " + ex.Message +
                Constants.vbCrLf + "- " + "Votre sortie MIDI pourrait être occupé par une autre application" +
                Constants.vbCrLf + "- " + "Essayer de choisir une autre sortie MIDI"
            Else
                MessageHV.PTitre = "MIDI Resource Problem"
                MessageHV.PContenuMess = "- " + "Warning : procedure error detection  : " + "JouerNote2" +
                    Constants.vbCrLf + "- " + "Message  : " + ex.Message +
                    Constants.vbCrLf + "- " + "Your MIDI output may be occupied by another application." +
                    Constants.vbCrLf + "- " + "Try to choose another MIDI output"
            End If
            Init_Dicpiano()
            If Clavier.Visible Then Clavier.RAZ_CouleurNotes()
            NoteCourante = ValeurNote
            NoteAEtéJouée = True
            Cacher_FormTransparents()
            MessageHV.ShowDialog()
        End Try
    End Sub
    Sub StoperNote2(ValeurNote As Byte, canal As Byte, Dyn As Byte)

        Try
            If Not (SortieMidi.Item(ChoixSortieMidi).IsOpen) Then
                SortieMidi.Item(ChoixSortieMidi).Open()
            End If
            '
            SortieMidi.Item(ChoixSortieMidi).SendNoteOff(canal - 1, ValeurNote, Dyn)
            NoteCourante = ValeurNote
            NoteAEtéJouée = True
        Catch ex As Exception
            MessageHV.PTitre = "Problème de ressource MIDI"
            MessageHV.PContenuMess = messa + Constants.vbCrLf + "Détection d'une erreur dans procédure : " + "JouerNote" + "." + Constants.vbCrLf +
            "Message  : " + ex.Message
            MessageHV.PTypBouton = "OK"
            Cacher_FormTransparents()
            MessageHV.ShowDialog()
        End Try
    End Sub
    Private Sub Button19_Click_1(sender As Object, e As EventArgs)
        Dim i As Integer
        Dim Fin As Integer = 16 * 48
        Dim liste1 As New List(Of Integer)

        For i = 0 To Fin
            If IsMultiple(i, 8) Then
                liste1.Add(i)
            End If
        Next
    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        Accent1_3 = True
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        Accent1_3 = False
    End Sub

    Private Sub Button19_Click_2(sender As Object, e As EventArgs) Handles Button19.Click
        '
        ' L'application est contenue dans une table d'onglets  Tabcontrol4  : 1..6 HyperArp, PianoRoll 1, PianoRoll 3, PianoRoll 3, DrumEdit, Mix
        ' Le 1er Onglet 1..6 HyperArp est organisé de la façon suivante : 
        ' - il y a 2 splitcontainer en jeu : splitcontainer7 et splitcontainer2
        ' - splicontainer7.panel1 contient toute la partie Accords et TimeLine des Accords
        ' - splicontainer7.panel2 contient splicontainer2
        ' - splicontainer2.panel1 contient les onglets de HyperArp
        ' - splicontainer2.panel2 contient les réglages systèmes

        ' Le présent bouton  permet de changer la position d'affichage du compteur de mesure et de faire disparaître le panel "admin système" complètement à droite.
        ' Il y a 2 compteurs de mesures : Label31 en affichage admin système, et Label15 en affichage standard
        ' Il faut aussi savoir que label15 est contenu  dans "Panel2" qui lui-même est contenu dans "SplitContainer7.Panel1". "Pnel2" est placé très à droite de splitcontainer7.Panel, dans sa partie non visible en affichage admin système 
        ' Donc, pour voir le compteur en standard Label15, il faut déplacer le splitter de Splicontainer7 vers la droite
        ' La position en mode système de splitcontainer7.splitterdistance = 535
        ' Pour voir le compteur en standard Label15, il faut déplacer splitcontainer7.splitterdistance = 800 à peu près.Le remettre ensuite à 535.

        If SplitContainer7.SplitterDistance = PosSystem Then
            SplitContainer7.SplitterDistance = PosStandard
            SplitContainer2.SplitterDistance = 750 '750
            SplitContainer2.Panel2.Visible = False
            Panel2.Visible = True
        Else

            SplitContainer7.SplitterDistance = PosSystem
            SplitContainer2.SplitterDistance = 545 '550
            SplitContainer2.Panel2.Visible = True
            Panel2.Visible = False
        End If
    End Sub
    ''' <summary>
    '''  Choix d'une Gamme par clic sur Botton35
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        Dim firstc As Integer = Grid2.Selection.FirstCol
        Dim lastc As Integer = Grid2.Selection.LastCol
        Dim a As String = ""
        Dim bb As String = ""
        Dim c As String
        Dim i As Integer
        Dim tbl() As String
        Dim b1 As String
        Dim GammeChoisie As String

        '
        ' Effacer les saisies de marqueur ou répétition du dernier accords
        ' *****************************************************************
        TMarqueur.Visible = False
        TMarqueur2.Visible = False
        'Button25.Image = HyperArp.My.Resources.Resources._32x32_Transp_vert__échelle3_
        'Button25.Image = HyperArp.My.Resources.Resources._32x32___échelle2_Blues

        If Grid2.SelectionMode = SelectionModeEnum.ByColumn Then

            ' 1 - détermination des accords sélectionnés dans une chaine a
            '     ********************************************************
            For i = firstc To lastc
                bb = Trim(Grid2.Cell(2, i).Text)
                If bb <> "" Then
                    tbl = bb.Split()
                    c = TradD(tbl(0))
                    If UBound(tbl) > 0 Then
                        c = c + " " + tbl(1)
                    End If
                    a = a + Trim(c) + "-" 'constitution de la liste d'accords séparée par des '-'
                End If
            Next
            If Trim(a) <> "" Then
                a = a.Substring(0, a.Length - 1) ' on retire le dernier séparateur '-' de la liste des accords 
                'Dim b As String = "Maj" + " " + "MinH" + " " + "MinM" + " " + "Penta" ' 2 - Détermination des gammes usuelles (non PRO) ' "Maj" + " " + "MinH" + " " + "MinM" + " " + "PMin" + " " + "Blues"
                'Dim b As String = "Maj" + " " + "MinH" + " " + "MinM" + " " + "Penta" + " " + "Blues1" + " " + "Blues2" + " " + "Blues3" _
                '    + " " + "MixoLyd.9/11" + " " + "Lyd.dim." + " " + "Hongrois1" + " " + "Hongrois2" _
                '    + " " + "Arabe" + " " + "Oriental" + " " + "Perse" + " " + "RagaTodi" + " " + "Japonais1" + " " + "Japonais2" _
                '    + " " + "Diminue" + " " + "Mesotonique" + " " + "Majeur_augm."


                ' Dim b As String = "Maj" + " " + "MinH" + " " + "MinM" + " " + "PentaMin" + " " + "Blues" + " " + "Blues2"

                Dim b = "Maj" + " " + "MinH" + " " + "MinM" + " " + "PentaMin" + " " + "Blues" + " " + "Blues2" + " " + "Hongrois1" + " " + "Hongrois2" + " " + "Oriental" + " " + "Arabe"


                '+ " " + "PentaMaj1" + " " + "PentaMaj2" + " " + "PentaMin1" + " " + "PentaMin2" + " " + "PentaMin3" _
                '+ " " + "Blues1" + " " + "Blues2" + " " + "Blues3" _
                '+ " " + "MixoLyd.9/11" + " " + "Lyd.dim." + " " + "Hongrois1" + " " + "Hongrois2" _
                '+ " " + "Arabe" + " " + "Oriental" + " " + "Perse" + " " + "RagaTodi" + " " + "Japonais1" + " " + "Japonais2" _
                '+ " " + "Diminue" + " " + "Mesotonique" + " " + "Majeur_augm." + " " + "Enigmatique"
                Dim oo As New RechercheG_v2
                ListGammesJouables = oo.ApparteanceG(a, b) ' 3 - Calcul des gammes jouables sur les accords sélectionnés pour le mode non PRO
                GammeChoisie = ""
                'If GammesPRO.Checked = False Then
                ChoixGamme.ShowDialog() ' Mode Non PRO
                If ChoixGamme.retour <> "" Then
                    GammeChoisie = Trim(ChoixGamme.ComboBox1.Text)
                End If
                'Else
                'FiltreGammes.Accords = Trim(a) ' Mode  PRO --> a est la liste des accords avec séparateur "-"
                'FiltreGammes.ShowDialog()
                'GammeChoisie = FiltreGammes.Retour ' valeur du combobox1 de Filtreamme (choix de l'utilisateur)
                'End If
                '
                ' Mise à jour des gammes choisies
                ' *******************************
                If Trim(GammeChoisie) <> "" Then
                    SAUV_Annuler(firstc, lastc)
                    For i = firstc To lastc
                        If Trim(Grid2.Cell(2, i).Text) <> "" Then
                            Grid2.Cell(11, i).Text = Trim(GammeChoisie)
                            TableEventH(i, 1, 1).Gamme = Trim(GammeChoisie)
                        End If
                    Next
                    '
                    b1 = Trim(Det_ListGam())
                    For i = 0 To nb_PianoRoll - 1
                        listPIANOROLL(i).PListGam = b1 'Det_ListGam()
                        listPIANOROLL(i).F1_Refresh()
                        listPIANOROLL(i).Maj_CalquesMIDI()
                    Next
                End If
            Else
                If Module1.LangueIHM = "fr" Then
                    mess = "Veuillez sélectionner au moins un accord"
                    titre = "Information"
                Else
                    mess = "Please, select at least one chord"
                    titre = "Information"
                End If
                ' 
                Cacher_FormTransparents()
                i = MessageBox.Show(mess, titre, MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End If
    End Sub
    Private Sub Grid2_Load(sender As Object, e As EventArgs)
    End Sub
    Private Sub SplitContainer7_SplitterMoved(sender As Object, e As SplitterEventArgs) Handles SplitContainer7.SplitterMoved
        If SplitContainer7.SplitterDistance < PosSystem Then
            SplitContainer7.SplitterDistance = PosSystem
            PositActuel = PosSystem
            Grid2.Width = PosSystem - 1
            Panel2.Visible = False
        Else
            Grid2.Width = SplitContainer7.SplitterDistance
            Panel2.Visible = True
        End If
    End Sub

    ''' <summary>
    ''' Ecriture du nomre de Répétition du dernier accord
    ''' </summary>
    ''' <param name="sender">Object textox envoyant l'évènement</param>
    ''' <param name="e">Evènements lié au sender</param>
    Private Sub TMarqueur2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TMarqueur2.KeyPress
        Dim a As Char
        a = e.KeyChar
        If a = vbCr Then
            EcritureMarqueur2()
            Me.Focus()
        End If
    End Sub
    Private Sub OctaveEcoute_KeyDown(sender As Object, e As KeyEventArgs) Handles OctaveEcoute.KeyDown
        e.SuppressKeyPress = True
    End Sub
    Private Sub CanalEcoute_KeyDown(sender As Object, e As KeyEventArgs) Handles CanalEcoute.KeyDown
        e.SuppressKeyPress = True
    End Sub
    Private Sub VéloEcoute_KeyDown(sender As Object, e As KeyEventArgs) Handles VéloEcoute.KeyDown
        e.SuppressKeyPress = True
    End Sub
    Private Sub ComboCopy1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboCopy1.KeyDown
        e.SuppressKeyPress = True
    End Sub
    Private Sub ComboCopy2_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboCopy2.KeyDown
        e.SuppressKeyPress = True
    End Sub

    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
        e.SuppressKeyPress = True
    End Sub

    Private Sub ComboBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox2.KeyDown
        e.SuppressKeyPress = True
    End Sub

    Private Sub ComboBox23_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox23.KeyDown
        e.SuppressKeyPress = True
    End Sub

    Private Sub ComboBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox3.KeyDown
        e.SuppressKeyPress = True
    End Sub

    Private Sub ComboBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox4.KeyDown
        e.SuppressKeyPress = True
    End Sub

    Private Sub ComboBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox6.KeyDown
        e.SuppressKeyPress = True
    End Sub

    Private Sub NoteRacineZ_KeyDown(sender As Object, e As KeyEventArgs)
        e.SuppressKeyPress = True
    End Sub

    Private Sub ComboOctaveDown_KeyDown(sender As Object, e As KeyEventArgs)
        e.SuppressKeyPress = True
    End Sub

    Private Sub ClavierToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Clavier.Show()
    End Sub

    Private Sub QuickBluesToolStripMenuItem_Click(sender As Object, e As EventArgs)
        QuickBlues.ShowDialog()
        ComboBox1.SelectedIndex = QuickBlues.ComboBox1.SelectedIndex
    End Sub

    Private Sub MappingToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MappingToolStripMenuItem.Click
        Cacher_FormTransparents()
        Mapping.ShowDialog()
    End Sub

    Private Sub OctaveToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        Ouverture_PianoRoll2(0)
    End Sub

    Private Sub EtendreToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        Ouverture_PianoRoll2(1)
    End Sub
    Sub PianoRoll_Docking(NPRoll As Integer)
        If listPIANOROLL(NPRoll).F1.Dock = DockStyle.Fill Then
            listPIANOROLL(NPRoll).F1.FormBorderStyle = FormBorderStyle.Sizable
            listPIANOROLL(NPRoll).F1.TopMost = True   ' 
            listPIANOROLL(NPRoll).F1.Dock = DockStyle.None
            TabControl4.TabPages.Item(NPRoll + 1).Controls.Remove(listPIANOROLL(NPRoll).F1)
            Dim p As New Point With {
                .X = Me.Location.X,
                .Y = Me.Location.Y + 50
            }
            listPIANOROLL(NPRoll).F1.Location = p
            listPIANOROLL(NPRoll).F1.StartPosition = FormStartPosition.Manual ' permet de tenir compte de la location calculée dans p
            listPIANOROLL(NPRoll).F1.TopLevel = True '
            listPIANOROLL(NPRoll).F1.MainMenuStrip.Visible = True
        Else
            listPIANOROLL(NPRoll).F1.FormBorderStyle = FormBorderStyle.None
            listPIANOROLL(NPRoll).F1.TopMost = False   ' un seul des 2 suffit ?
            listPIANOROLL(NPRoll).F1.TopLevel = False
            TabControl4.TabPages.Item(NPRoll + 1).Controls.Add(listPIANOROLL(NPRoll).F1)
            listPIANOROLL(NPRoll).F1.Dock = DockStyle.Fill
            listPIANOROLL(NPRoll).F1.MainMenuStrip.Visible = False
        End If
        '
    End Sub


    Private Sub Grid2_Click(Sender As Object, e As EventArgs) Handles Grid2.Click

    End Sub

    Private Sub ToolStripMenuItem13_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem13.Click

        ExportCalqueMIDI()
    End Sub
    Sub ExportCalqueMIDI()
        Dim a As String
        '
        Try
            If Det_NomCalquesMIDI() Then
                Cacher_FormTransparents()
                ' création Fichier Calques MIDI dont le nom est donné dans a
                ' **********************************************************
                a = CréationFichierTexteCalquesMIDI("FichierCalquesMIDI") ' FichierCalquesMIDI.mid est le nom du fichier MIDI qui va être généré.
                Création_MidiFile(a, 2)
                ' copie du fichier Calques MIDI dans le fichier choisi
                ' ****************************************************
                a = Création_CTemp() + "\" + "FichierCalquesMIDI" + ".mid" '
                If My.Computer.FileSystem.FileExists(FichierCalques) Then
                    My.Computer.FileSystem.DeleteFile(FichierCalques)
                End If
                My.Computer.FileSystem.CopyFile(a, FichierCalques) ' la variable FichierCalques est mis à jour dans Det_NomCalquesMIDI
            End If

        Catch ex As Exception
            a = ex.Message
        End Try
    End Sub
    '
    Function CréationFichierTexteCalquesMIDI(NomFichierCalques As String) As String
        Dim h, i, k, n, t As Integer
        Dim m As Integer
        Dim a, b, c, d, notes As String
        Dim ligne1 As String = ""
        Dim tbl1() As String = Nothing
        Dim tbl2() As String = Nothing
        '
        Dim FIN As Double
        Dim NPiste As Integer
        '
        Dim DuréeFIN As String
        Dim LongueurFinDerMesure As Integer = 0
        Dim nM As Integer = 0
        Dim nMCycle As Integer = 0
        Dim Position As String
        Dim PositionDansCycle As String = ""
        Dim tt As New TT
        '
        Dim Sortir1 As Boolean = False
        Dim Sortir2 As Boolean = False
        '
        Dim DernièreNoteCycle As String

        CréationFichierTexteCalquesMIDI = ""
        Try
            '
            ' création du fichier texte de base
            ' *********************************
            '
            NPiste = 2 ' Une piste accord, Une piste gamme
            '
            a = Création_CTemp() ' création du répertoire Documents/HyperArp
            a = a + "\" + "CalquesMIDI.txt" ' prendre en compte le nom du fichier

            If My.Computer.FileSystem.FileExists(a) Then
                My.Computer.FileSystem.DeleteFile(a)
            End If
            '
            Dim fileWriter = My.Computer.FileSystem.OpenTextFileWriter(a, True) ' création du fichier dans le dossier HyperArp
            '
            ' Début écriture avec Nom, Tempo et Métrique
            ' ******************************************
            fileWriter.WriteLine("NomFichier;" + Trim(NomFichierCalques))
            fileWriter.WriteLine("Tempo;" + Trim(Str(Tempo.Value)))
            fileWriter.WriteLine("Métrique;" + Trim(Métrique.Text))
            '
            ' iniitialisations 
            FIN = -1
            DuréeFIN = "0"
            '
            i = Det_NumDerAccord()
            LongueurFinDerMesure = Convert.ToUInt16(Grid2.Cell(3, i).Text) ' nombre de répétitions du dernier accord
            '
            For k = 0 To NPiste - 1 '
                Select Case k
                    Case 0
                        If Langue = "fr" Then
                            fileWriter.WriteLine("NomPiste;" + Trim(Str(k)) + ";Calque_Accords(Muter la piste)")
                        Else
                            fileWriter.WriteLine("NomPiste;" + Trim(Str(k)) + ";Chords_Layer(Mute track")
                        End If
                    Case 1
                        If Langue = "fr" Then
                            fileWriter.WriteLine("NomPiste;" + Trim(Str(k)) + ";Calque_Gammes(Muter la piste)")
                        Else
                            fileWriter.WriteLine("NomPiste;" + Trim(Str(k)) + ";Scales_Layer(Mute track)")
                        End If
                End Select
                '
                For m = 0 To UBound(TableEventH, 1)
                    If TableEventH(m, 1, 1).Ligne <> -1 Then
                        Position = Trim(Str(m)) + "." + "1" + "." + "1"
                        Maj_TabNotes_Minus("#")
                        '
                        Select Case k
                            Case 0 ' traitement Accord
                                fileWriter.WriteLine("Accord : " + TableEventH(m, 1, 1).Accord)
                                notes = Det_NotesAccord(TableEventH(m, 1, 1).Accord)
                                tbl2 = Split(notes, "-")
                            Case 1 ' traitement Gammes
                                d = TableEventH(m, 1, 1).Gamme
                                fileWriter.WriteLine("Gamme : " + TableEventH(m, 1, 1).Gamme)
                                notes = Det_NotesGammes(TableEventH(m, 1, 1).Gamme)
                                tbl2 = Split(notes)
                        End Select
                        '
                        For i = 0 To UBound(tbl2)
                            tt = DicoNotes(tbl2(i)) '  -- Public DicoNotes As New Dictionary(Of String, TT)
                            For h = 0 To UBound(tt.Toctave) ' tt.Toctave contient les N° MIDI d'une note sur toute les octaves
                                n = tt.Toctave(h)
                                If n < 128 Then
                                    ligne1 = "Note;" + Trim(Str(k)) + ";" + Trim(Str(16)) + ";" ' canal pour les calques MIDI = 16
                                    ligne1 = ligne1 + Trim(Str(n)) + ";" ' 16;64"
                                    b = DébutEVT3(Position) + ";"
                                    DernièreNoteCycle = Det_DerNoteCycle2()
                                    If Position <> DernièreNoteCycle Then
                                        t = Convert.ToUInt16(Grid2.Cell(3, m).Text) * 16 'Nom bre de répétition x Longueur d'une mesure en double-croches
                                        c = Convert.ToString(t) + ";" 'Nom bre de répétition x Longueur d'une mesure en double-croches
                                    Else
                                        t = LongueurFinDerMesure * 16
                                        c = Convert.ToString(t) + ";"  '
                                    End If
                                    ligne1 = ligne1 + Trim(b) + Trim(c) + "0" ' Str(1) : Vélocité pour les calques MIDI = 0
                                    '
                                    fileWriter.WriteLine(ligne1)
                                End If
                            Next h
                        Next i
                        ligne1 = ""
                    End If
                    '
                Next m
            Next k
            '
            fileWriter.Close()
            '
            ' Remise du chemin du fichier à la procédure appelante
            ' ****************************************************
            CréationFichierTexteCalquesMIDI = a ' remise du chemin + fichier
            '
        Catch ex As Exception
            Avertis = "Erreur interne : procédure 'CréationFichierTexteCalquesMIDI' : " + ex.Message
            Cacher_FormTransparents()
            MessageHV.ShowDialog()
        End Try
    End Function
    Function DébutEVT3(Position As String) As String
        Dim Tbli() As String
        Dim m As Integer
        '
        Tbli = Split(Position, ".")
        m = Val(Tbli(0))
        DébutEVT3 = Trim(Str((m - 1) * 16)) ' la longueur d'une mesure est  16 doubles croches  pour 4/4
    End Function

    Private Sub TabControl4_Deselected(sender As Object, e As TabControlEventArgs) Handles TabControl4.Deselected
        PIANOROLLNPréced = e.TabPageIndex - 1
    End Sub
    Private Sub TabControl4_Selected(sender As Object, e As TabControlEventArgs) Handles TabControl4.Selected
        OngletCours_Edition = e.TabPageIndex

        Dim Der_PianoRoll As Integer = nb_PianoRoll
        '
        ' Transfert du clipboard de PIANOROLL précédent dans le clipboard (List CCC) du nouveau PIANOROLL sélectionné
        ' Cela permet de copier le contenu d'un pianoroll dans un autre (cela ne concerne le CTRL Z mais seulement CTRL C,V,X(listCCC))
        Dim o As Integer = PIANOROLLNPréced      ' Origine
        Dim d As Integer = e.TabPageIndex - 1    ' Destinataire
        If (o >= 0 And o < Der_PianoRoll) And (d >= 0 And d < Der_PianoRoll) Then '  ' si o=2, on vient de drum edit et il en faut prendre en compte cet évènement ici (idem pour d)
            listPIANOROLL(d).ListCCC.Clear()
            For Each el In listPIANOROLL(o).ListCCC
                listPIANOROLL(d).ListCCC.Add(el)
            Next
        End If
        ' Mise à jour du modèle de batterie au cas où la batterie a été modifiée <- plus cours - les modèles de batterie n'apparaissent plus dans les piano roll car le système est trop couteux en temps
        ' Drums.Refresh_PianoRoll()

        ' Sélection du bouton de changement fraphique de l'onglet HyperArp
        ' ****************************************************************
        If OngletCours_Edition = 0 Then
            Button19.Visible = True
        Else
            Button19.Visible = False
        End If
    End Sub
    Private Sub BPianoRoll1_Click(sender As Object, e As EventArgs)
        PianoRoll_Docking(0)
        '
    End Sub
    Private Sub BPianoRoll2_Click(sender As Object, e As EventArgs)
        PianoRoll_Docking(1)
        '

    End Sub
    Private Sub Terme_KeyDown(sender As Object, e As KeyEventArgs) Handles Terme.KeyDown
        'e.SuppressKeyPress = True
        Dim c As String = e.KeyCode.ToString
        If c = "Return" Then
            If OngletCours_Edition > 0 And OngletCours_Edition <= Module1.nb_PianoRoll Then
                listPIANOROLL(OngletCours_Edition - 1).PStartMeasure = Terme.Value ' positionnement du pianoroll en cours sur la mesure indiquée dans le locateur "Terme" (fin) (simple utilisation de leftcol la grille du pianoroll)
            End If
        End If
    End Sub
    Private Sub BPianoRoll1_KeyDown(sender As Object, e As KeyEventArgs)
        e.SuppressKeyPress = False
    End Sub

    Private Sub BPianoRoll2_KeyDown(sender As Object, e As KeyEventArgs)
        e.SuppressKeyPress = False
    End Sub

    Private Sub Début_KeyDown(sender As Object, e As KeyEventArgs) Handles Début.KeyDown
        Dim c As String = e.KeyCode.ToString
        If c = "Return" Then
            If OngletCours_Edition > 0 And OngletCours_Edition <= Module1.nb_PianoRoll Then
                listPIANOROLL(OngletCours_Edition - 1).PStartMeasure = Début.Value ' positionnement du pianoroll en cours sur la mesure indiquée dans le locateur "Début"
            End If
        End If
    End Sub

    Private Sub Button30_Click_1(sender As Object, e As EventArgs)
        Dim d As New DrumEdit(10) ' canal=10

        Drums.F2.Visible = False
        Drums.F2.Tag = nb_PianoRoll + 1
        '
        ' Maj des propriétés de DrumEdit
        ' ******************************
        Drums.PLangue = LangueIHM
        Drums.PNbMesures = nbMesures '
        Drums.PMétrique = Métrique.Text
        Drums.PnbRépétitionMax = nbRépétitionMax
        Drums.PListAcc = Det_ListAcc()
        Drums.PListMarq = Det_ListMarq()
        Drums.PLangue = LangueIHM
        ' Appel de la construction
        ' ************************
        Drums.Construction_F2()
        Drums.F2.FormBorderStyle = FormBorderStyle.None
        Drums.F2.TopLevel = False  '
        Drums.F2.TopMost = False   ' un seul des 2 suffit ?
        Drums.F2.Dock = DockStyle.Fill
        TabControl4.TabPages.Item(nb_PianoRoll + 1).Controls.Add(d.F2)
        Drums.F2.Visible = True
    End Sub

    Private Sub Button30_Click_2(sender As Object, e As EventArgs)
        ' PListNotes(Répéter As Boolean, Boucle As Integer, Form1_Début As Integer, Form1_Fin As Integer, NumDerAcc As Integer) As String
        ' Contruction_ListeNotes(Répéter, Boucle, Form1_Début, Form1_Fin, NumDerAcc)
        Dim a As String = Drums.PListNotes(False, 2, 1, 1, 1)
    End Sub

    Private Sub Form1_CursorChanged(sender As Object, e As EventArgs) Handles Me.CursorChanged

    End Sub

    Private Sub PianoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PianoToolStripMenuItem.Click
        Clavier.Show()
    End Sub

    Private Sub Panel8_Paint(sender As Object, e As PaintEventArgs) Handles Panel8.Paint

    End Sub

    Private Sub QuickBluesToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles QuickBluesToolStripMenuItem1.Click
        Cacher_FormTransparents()
        QuickBlues.ShowDialog()
        ComboBox1.SelectedIndex = QuickBlues.ComboBox1.SelectedIndex
    End Sub

    Private Sub Button7_Click_1(sender As Object, e As EventArgs)
        MyPoint1 = 50
        MyPoint2 = 0
        MyInterval = 10
        Transition.ShowDialog()

    End Sub

    Private Sub TransportBarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TransportBarToolStripMenuItem.Click
        If Transport.Visible = False Then
            Transport.TopLevel = True
            Transport.TopMost = True
            Transport.StartPosition = FormStartPosition.CenterScreen
            Transport.Show() ' form non modale
        Else
            Transport.Visible = False
        End If
    End Sub


    Private Sub TabControl2_Selected(sender As Object, e As TabControlEventArgs) Handles TabControl2.Selected
        OngletCours_HyperARP = e.TabPageIndex
    End Sub
    ''' <summary>
    ''' TMarqueur_KeyUp : ici on supprime (pendant l'écriture au clavier) l'entrée des caractères " " ,  ";" et "et commercial" qui sont utilisés comme
    ''' délimiteurs dans la liste des marqueurs
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub TMarqueur_KeyUp(sender As Object, e As KeyEventArgs) Handles TMarqueur.KeyUp

        Dim tbl() As String
        If Microsoft.VisualBasic.Right(TMarqueur.Text, 1) = ";" Then
            tbl = TMarqueur.Text.Split(";")
            TMarqueur.Text = Trim(tbl(0))
        End If
        '
        If Microsoft.VisualBasic.Right(TMarqueur.Text, 1) = "&" Then
            tbl = TMarqueur.Text.Split("&")
            TMarqueur.Text = Trim(tbl(0))
        End If
        '
        If Microsoft.VisualBasic.Right(TMarqueur.Text, 1) = " " Then
            tbl = TMarqueur.Text.Split(" ")
            TMarqueur.Text = Trim(tbl(0))
        End If
        TMarqueur.SelectionStart = Len(TMarqueur.Text)
    End Sub
    ''' <summary>
    ''' Ecriture d'un Marqeur dans la Time Line
    ''' </summary>
    ''' <param name="sender">Object textox envoyant l'évènement</param>
    ''' <param name="e">Evènements lié au sender</param>
    Private Sub TMarqueur_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TMarqueur.KeyPress
        Dim a As Char = e.KeyChar

        If a = vbCr Then
            EcritureMarqueur()
            Maj_VueNotes()
            Me.Focus()
        End If
    End Sub

    Private Sub MuteOptimisationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MuteOptimisationToolStripMenuItem.Click
        AutomationSup.ShowDialog()
    End Sub

    Private Sub Form1_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp

        ' PLAY, RECALCUL : F5
        ' *******************
        If e.KeyCode = Keys.F5 Then
            If Not Me.Horloge1.IsRunning Then
                Me.PlayHyperArp()
            Else
                Me.ReCalcul()
            End If
        End If
        '
        ' STOP : F4
        ' *********
        If e.KeyCode = Keys.F4 Then
            StopPlay()
        End If
    End Sub

    Private Sub Form1_Activated(sender As Object, e As EventArgs) Handles Me.Activated

        Clipboard.Clear()
    End Sub
    Sub RafraichissementOnglets()
        Dim i As Integer
        TabControl5.Visible = False
        For i = 0 To TabControl5.TabPages.Count - 1
            TabControl5.SelectedTab = TabControl5.TabPages.Item(i)
            TabControl5.TabPages.Item(i).Refresh()
            TabControl5.TabPages.Item(i).Show()
            TabControl5.Refresh()
        Next i
        TabControl5.SelectedTab = TabControl5.TabPages.Item(0)
        TabControl5.TabPages.Item(0).Show()
        TabControl5.Visible = True
    End Sub

    Private Sub CheckBox1_KeyPress(sender As Object, e As KeyPressEventArgs)

    End Sub
    ''' <summary>
    ''' TabControl2_DrawItem : Tabcontrol HyperArp
    ''' Couleur des onglets de source d'accords
    ''' </summary>
    ''' ForeColor noir = SystemBrushes.MenuText donne le ForeColor de l'onglet = noir
    ''' ForeColor Blanc = SystemBrushes.HighLight donne le ForeColor de l'onglet = blanc
    ''' <param name="sender"></param>
    ''' <param name="e"></param>

    Private Sub TabControl2_DrawItem(sender As Object, e As DrawItemEventArgs) Handles TabControl2.DrawItem
        Dim paddedBounds As Rectangle = e.Bounds
        paddedBounds.Inflate(0, -2)


        Select Case e.Index
            Case 0, 1  ' Gammes d'accords et progression
                e.Graphics.FillRectangle(New SolidBrush(Color.LightSteelBlue), e.Bounds) ' backcolor
                e.Graphics.DrawString(TabControl2.TabPages(e.Index).Text, fnt1, SystemBrushes.MenuText, paddedBounds) ' ForeColor
            Case 2, 3       ' Substitution
                e.Graphics.FillRectangle(New SolidBrush(Color.PaleGoldenrod), e.Bounds) ' backcolor
                e.Graphics.DrawString(TabControl2.TabPages(e.Index).Text, fnt1, SystemBrushes.MenuText, paddedBounds) ' ForeColor
            Case 4     ' Vue notes
                e.Graphics.FillRectangle(New SolidBrush(Color.Maroon), e.Bounds) ' backcolor
                e.Graphics.DrawString(TabControl2.TabPages(e.Index).Text, fnt1, SystemBrushes.HighlightText, paddedBounds) ' ForeColor
            Case 5, 6  ' Perso et Automation
                e.Graphics.FillRectangle(New SolidBrush(Color.DarkOliveGreen), e.Bounds)
                e.Graphics.DrawString(TabControl2.TabPages(e.Index).Text, fnt1, SystemBrushes.HighlightText, paddedBounds)
        End Select

    End Sub

    ''' <summary>
    ''' TabControl5_DrawItem : Couleur des onglets des variations
    ''' </summary>
    ''' ForeColor noir = SystemBrushes.MenuText donne le ForeColor de l'onglet = noir
    ''' ForeColor Blanc = SystemBrushes.HighLight donne le ForeColor de l'onglet = blanc
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub TabControl5_DrawItem(sender As Object, e As DrawItemEventArgs) Handles TabControl5.DrawItem
        Dim paddedBounds As Rectangle = e.Bounds
        paddedBounds.Inflate(0, -2)
        ' e.Graphics.FillRectangle(New SolidBrush(Color.DarkOliveGreen), e.Bounds)
        e.Graphics.FillRectangle(New SolidBrush(TabCoulMagneto(e.Index)), e.Bounds)
        e.Graphics.DrawString(TabControl5.TabPages(e.Index).Text, fnt1, SystemBrushes.MenuText, paddedBounds)

    End Sub

    Private Sub TabControl4_DrawItem(sender As Object, e As DrawItemEventArgs) Handles TabControl4.DrawItem
        ' Couleur ForeColor des titres des onglets -->
        ' SystemBrushes.MenuText donne le ForeColor de l'onglet = noir
        ' SystemBrushes.HighLight donne le ForeColor de l'onglet = blanc

        Dim paddedBounds As Rectangle = e.Bounds
        paddedBounds.Inflate(0, -2)

        Select Case e.Index
            Case 0 ' hyperarp
                e.Graphics.FillRectangle(New SolidBrush(Color.Gold), e.Bounds)
                e.Graphics.DrawString(TabControl4.TabPages(e.Index).Text, fnt1, SystemBrushes.MenuText, paddedBounds)
            Case 1, 2, 3, 5, 6 ' PianoRoll
                e.Graphics.FillRectangle(New SolidBrush(Color.Beige), e.Bounds)
                e.Graphics.DrawString(TabControl4.TabPages(e.Index).Text, fnt1, SystemBrushes.MenuText, paddedBounds)
            Case 4 ' drumedit
                e.Graphics.FillRectangle(New SolidBrush(Color.LightCoral), e.Bounds)
                e.Graphics.DrawString(TabControl4.TabPages(e.Index).Text, fnt1, SystemBrushes.MenuText, paddedBounds)
            Case 7, 8 ' ChordsRoll
                e.Graphics.FillRectangle(New SolidBrush(Color.Khaki), e.Bounds)
                e.Graphics.DrawString(TabControl4.TabPages(e.Index).Text, fnt1, SystemBrushes.MenuText, paddedBounds)
            Case Else ' mix
                e.Graphics.FillRectangle(New SolidBrush(Color.DarkSeaGreen), e.Bounds)
                e.Graphics.DrawString(TabControl4.TabPages(e.Index).Text, fnt1, SystemBrushes.MenuText, paddedBounds)
        End Select
    End Sub

    Private Sub Transpo_Click(sender As Object, e As EventArgs) Handles Transpo.Click
        Transposition.ShowDialog()
        If Transposition.Bout_Val = "OK" And Transposition.Transp_Val <> 0 Then
            ' Pré-Vérification dépassement de tessiture
            ' *****************************************
            Dim bb As Boolean = True

            For i = 0 To nb_PianoRoll - 1
                If PIANOROLLChargé(i) = True Then
                    bb = listPIANOROLL(i).Pré_VérifTransp(Grid2.Selection.FirstCol, Grid2.Selection.LastCol, Transposition.Transp_Val)
                    If Not bb Then Exit For
                End If
            Next
            '
            If bb Then
                ' Tranposition HyperArp
                SAUV_Annuler(Grid2.Selection.FirstCol, Grid2.Selection.LastCol)
                Transposer3(Transposition.Transp_Val, Grid2.Selection.FirstCol, Grid2.Selection.LastCol)
                ' Transposition des PianoRolls
                If Transposition.TraiterPianoR Then
                    For i = 0 To nb_PianoRoll - 1
                        If PIANOROLLChargé(i) = True Then
                            listPIANOROLL(i).Actiontransposer(Grid2.Selection.FirstCol, Grid2.Selection.LastCol, Transposition.Transp_Val)
                        End If
                    Next
                End If
            End If
        End If
    End Sub
    Private Sub TRacine_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TRacine.SelectedIndexChanged
        Dim a As String = Trim(TRacine.Text)
        Dim i As Integer

        ' Effacer les saisies de marqueur ou répétition du dernier accords
        ' *****************************************************************
        TMarqueur.Visible = False
        TMarqueur2.Visible = False

        'If Grid2.SelectionMode = SelectionModeEnum.Free Then
        If Grid2.Selection.LastRow = 12 Then
            For i = Grid2.Selection.FirstCol To Grid2.Selection.LastCol
                If Trim(Grid2.Cell(2, i).Text) <> "" Then
                    Grid2.Cell(12, i).Text = a
                    TableEventH(i, 1, 1).Racine = a
                End If
            Next
        End If
        Calcul_AutoVoicingZ()
        'End If
    End Sub

    Private Sub TRacine_MouseDown(sender As Object, e As MouseEventArgs) Handles TRacine.MouseDown
        ' Effacer les saisies de marqueur ou répétition du dernier accords
        ' *****************************************************************
        TMarqueur.Visible = False
        TMarqueur2.Visible = False
    End Sub

    Private Sub MenuTransportBar_Click(sender As Object, e As EventArgs) Handles MenuTransportBar.Click
        If Transport.Visible = False Then
            Transport.TopLevel = True
            Transport.TopMost = True
            Transport.StartPosition = FormStartPosition.CenterScreen
            Transport.Show() ' form non modale
        Else
            Transport.Visible = False
        End If
    End Sub
    Private Sub MenuKeyBoard_Click(sender As Object, e As EventArgs) Handles MenuKeyBoard.Click
        Clavier.TopMost = True
        Clavier.CheckBox1.Checked = True
        Clavier.Show() ' form non modale
    End Sub
    Private Sub MenuQuickBlues_Click(sender As Object, e As EventArgs) Handles MenuQuickBlues.Click
        'Cacher_FormTransparents()
        QBlues2.ShowDialog()
        'ComboBox1.SelectedIndex = QuickBlues.ComboBox1.SelectedIndex
    End Sub
    Private Sub ToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem3.Click
        If Module1.LangueIHM = "fr" Then
            Process.Start("https://www.hyperarp.fr/guide-rapide")
        Else
            Process.Start("https://www.hyperarp.fr/guide-rapide?lang=en")
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Process.Start("https://www.hyperarp.fr")
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Process.Start("https://www.hyperarp.fr/?lang=en")
    End Sub

    Private Sub Début_KeyUp(sender As Object, e As KeyEventArgs) Handles Début.KeyUp
        Cohérence_Délim()
    End Sub
    Private Sub Cohérence_Délim()

        If Terme.Value >= Début.Value Then
            Terme.BackColor = Color.White
            Début.BackColor = Color.White
            Terme.ForeColor = Color.Black
            Début.ForeColor = Color.Black
        Else
            Terme.BackColor = Color.Red
            Début.BackColor = Color.Red
            Terme.ForeColor = Color.White
            Début.ForeColor = Color.White
        End If
    End Sub

    Private Sub Terme_KeyUp(sender As Object, e As KeyEventArgs) Handles Terme.KeyUp
        Cohérence_Délim()
    End Sub

    Private Sub Grid2_KeyUp(Sender As Object, e As KeyEventArgs) Handles Grid2.KeyUp
        OK_KeyDown = True  ' pour incrémentation/décrémentation des racines (évite les répétitions par appui continu)

        ' Couper Jouer Accord
        ' *******************
        If AccordAEtéJoué = True Then
            CouperJouerAccord2()
            AccordAEtéJoué = False
        End If

    End Sub

    Private Sub NbBoucles_SelectedIndexChanged(sender As Object, e As EventArgs) Handles NbBoucles.SelectedIndexChanged

    End Sub

    Private Sub NbBoucles_KeyDown(sender As Object, e As KeyEventArgs) Handles NbBoucles.KeyDown
        e.SuppressKeyPress = True
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If Not (SortieMidi.Item(ChoixSortieMidi).IsOpen) Then
            SortieMidi.Item(ChoixSortieMidi).Open()
        End If
        SortieMidi.Item(ChoixSortieMidi).SendControlChange(0, 54, 64) ' Canal, CTRL, Valeur CTRL
        'SortieMidi.Item(ChoixSortieMidi).SendNoteOn(1, 120, 64)
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        If Not (SortieMidi.Item(ChoixSortieMidi).IsOpen) Then
            SortieMidi.Item(ChoixSortieMidi).Open()
        End If
        SortieMidi.Item(ChoixSortieMidi).SendControlChange(0, 55, 64) ' Canal, CTRL, Valeur CTRL
        'SortieMidi.Item(ChoixSortieMidi).SendNoteOn(1, 119, 64)
    End Sub


    Sub Send_CTRL54_Remote()
        If Not (SortieMidi.Item(ChoixSortieMidi).IsOpen) Then
            SortieMidi.Item(ChoixSortieMidi).Open()
        End If
        SortieMidi.Item(ChoixSortieMidi).SendControlChange(0, 54, 64) ' Canal, CTRL, Valeur CTRL
    End Sub
    Sub Send_CTRL55_Remote()
        If Not (SortieMidi.Item(ChoixSortieMidi).IsOpen) Then
            SortieMidi.Item(ChoixSortieMidi).Open()
        End If
        SortieMidi.Item(ChoixSortieMidi).SendControlChange(0, 55, 64) ' Canal, CTRL, Valeur CTRL
    End Sub

    Private Sub Fournotes_CheckedChanged(sender As Object, e As EventArgs) Handles Fournotes.CheckedChanged
        Calcul_AutoVoicingZ()
    End Sub

    Private Sub Bassemoins1_CheckedChanged(sender As Object, e As EventArgs) Handles Bassemoins1.CheckedChanged
        Calcul_AutoVoicingZ()
    End Sub
    Private Sub Maj_VueNotes()
        Dim m, t, ct As Integer
        Dim tbl() As String
        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim _1sep As Integer = 1 ' nombre de ligne de séparation
        Dim _2sep As Integer = 2
        Dim SepH As Integer = 2
        '
        Grid7.AutoRedraw = False

        Grid7.Cols = 1
        Grid7.Rows = 20
        '
        '
        For i = 0 To Grid7.Rows - 1
            Grid7.Row(i).Height = 19
        Next
        '
        Grid7.Column(0).Alignment = AlignmentEnum.LeftCenter
        Grid7.Column(0).Width = 50
        Grid7.Row(2).Height = SepH ' 2e ligne de séparation
        Grid7.Row(8).Height = SepH ' 2e ligne de séparation
        '
        Grid7.Cell(1, 0).Text = "Position"
        Grid7.Cell(2 + _1sep, 0).Text = "V5"
        Grid7.Cell(3 + _1sep, 0).Text = "V4"
        Grid7.Cell(4 + _1sep, 0).Text = "V3"
        Grid7.Cell(5 + _1sep, 0).Text = "V2"
        Grid7.Cell(6 + _1sep, 0).Text = "V1"
        If LangueIHM = "fr" Then
            Grid7.Cell(7 + _2sep, 0).Text = "Racine"
            Grid7.Cell(8 + _2sep, 0).Text = "Accord"
            Grid7.Cell(9 + _2sep, 0).Text = "Marqueur"
            Grid7.Cell(10 + _2sep, 0).Text = "Tonalité"
            Grid7.Cell(11 + _2sep, 0).Text = "Mode"
            Grid7.Cell(12 + _2sep, 0).Text = "Gamme"
        Else
            Grid7.Cell(7 + _2sep, 0).Text = "Root"
            Grid7.Cell(8 + _2sep, 0).Text = "Chord"
            Grid7.Cell(9 + _2sep, 0).Text = "Marker"
            Grid7.Cell(10 + _2sep, 0).Text = "Tonality"
            Grid7.Cell(11 + _2sep, 0).Text = "Mode"
            Grid7.Cell(12 + _2sep, 0).Text = "Scale"
        End If
        '
        For m = 0 To nbMesures - 1
            For t = 0 To 5
                For ct = 0 To 4
                    If Trim(TableNotesAccordsZ(m, t, ct)) <> "" Then
                        j = j + 1
                        Grid7.Cols = Grid7.Cols + 1
                        Grid7.Column(j).Width = 50
                        Grid7.Column(j).Alignment = AlignmentEnum.CenterCenter
                        '
                        Grid7.Cell(0, j).Text = j.ToString
                        tbl = Trim(TableNotesAccordsZ(m, t, ct)).Split
                        For i = 0 To tbl.Count - 1
                            Grid7.Cell((6 - i) + _1sep, j).Text = tbl(i)
                        Next
                        '
                        'TableEventH(m, t, ct).
                        Grid7.Cell(1, j).Orientation = FlexCell.TextOrientationEnum.Horizontal
                        Grid7.Cell(1, j).Text = m.ToString() + "." + t.ToString + "." + ct.ToString
                        '
                        Grid7.Cell(7 + _2sep, j).Text = TableEventH(m, t, ct).Racine
                        Grid7.Cell(8 + _2sep, j).Text = TableEventH(m, t, ct).Accord
                        '
                        Grid7.Cell(9 + _2sep, j).ForeColor = Color.Red
                        Grid7.Cell(9 + _2sep, j).Text = TableEventH(m, t, ct).Marqueur
                        '
                        Grid7.Cell(10 + _2sep, j).Text = TableEventH(m, t, ct).Tonalité
                        Grid7.Cell(11 + _2sep, j).Text = TableEventH(m, t, ct).Mode
                        Grid7.Cell(12 + _2sep, j).Text = TableEventH(m, t, ct).Gamme
                    End If
                Next ct
            Next t
        Next m
        '
        '
        Grid7.Cols = Grid7.Cols + 10 ' pour faire en sorte que le scroll horizontal soit toujours apparent
        ' Colorisation
        ' ************
        Grid7.Range(1, 1, 1, Grid7.Cols - 1).BackColor = Color.PaleGoldenrod ' N° mesres

        Grid7.Range(2, 1, 2, Grid7.Cols - 1).BackColor = Color.DarkKhaki ' 1er ligne séparatrice
        '
        Grid7.Range(3, 1, 7, Grid7.Cols - 1).BackColor = Color.Maroon ' Notes  des accords
        Grid7.Range(3, 1, 7, Grid7.Cols - 1).ForeColor = Color.Yellow
        '
        Grid7.Range(8, 1, 8, Grid7.Cols - 1).BackColor = Color.DarkKhaki ' 2e  ligne séparatrice
        '
        Grid7.Range(9, 1, 14, Grid7.Cols - 1).BackColor = Color.PaleGoldenrod ' info complémentaires
        '
        Grid7.AutoRedraw = True
        Grid7.Refresh()
    End Sub

    Private Sub Grid2_Load_1(sender As Object, e As EventArgs) Handles Grid2.Load

    End Sub

    Private Sub Grid5_MouseDown(Sender As Object, e As MouseEventArgs) Handles Grid5.MouseDown
        Dim i As Integer = Grid5.MouseRow ' remarque : ActiveCell.Row retourne 1 tandis que MouseRow retourne 0 (ça vient du fait que l'on clique sur une ligne fixe)
        Dim j As Integer = Grid5.MouseCol ' idem
        If Grid5.Cell(i, j).Text <> Nothing Or Trim(Grid5.Cell(i, j).Text) <> "" Then
            Dim a As String = Trim(Grid5.Cell(i, j).Text)
            '
            Dim r As Integer = Grid5.Selection.FirstRow
            Dim rr As Integer = Grid5.Selection.LastRow
            Dim c As Integer = Grid5.Selection.FirstCol
            Dim cc As Integer = Grid5.Selection.LastCol
            '
            LigneCoursGrid5 = i
            '
            ' Jouer accord de grid5
            ' *********************
            If a <> "" And e.Button() = MouseButtons.Left And My.Computer.Keyboard.CtrlKeyDown Then
                JouerAccord(a)
            End If
            '
            ' Mise à jour du tableau Grid5 et des label de modulation
            ' *******************************************************
            'If i = 1 And (j > 0 And j < 8) Then ' ligne de la tonalité courante
            ' Effacerselections()
            '
            'If i = 1 Then
            'LabModulat.Item(0).Text = "---" ' non utilisée
            'LabModulat.Item(1).Text = "---"
            'LabModulat.Item(2).Text = "---"
            'LabModulat.Item(3).Text = "---"
            'Label79.Text = "---"
            'Label79.ForeColor = Color.Red
            'Label80.Visible = False
            'If LangueIHM = "fr" Then
            'Label79.Text = "Veuillez choisir une Tonalité de modulation dans le tableau"
            'Else
            'Label79.Text = "Please choose a modulation tonality in the grid"
            'End If
            'End If
            ' 
            'LabModulat.Item(0).Text = Trim(Label33.Text) ' mise à jour Accord Pivot
            ''
            'EffacerTonVoisins()
            '' Maj_TonsVoisins(Trim(Grid5.Cell(1, j).Text))
            'Maj_TonsVoisins(Trim(Grid5.Cell(1, j).Text))
            'RecheraccordGrid5(Trim(Grid5.Cell(1, j).Text))
            ''
            'If r > 0 Then
            'Lab_1 = Grid5.Cell(r, c).Text 'accord sélectionné
            'Lab_2 = Trim(Grid5.Cell(r, 0).Text) ' sauvegarde de la "nouvelle tonalité"
            'Maj_Cadence(Lab_1, Lab_2)
            'Label79.Text = Trim(Lab_2)
            'Label80.Visible = True
            'End If

            'End If
            '
            Lab_1 = Trim(Grid5.Cell(i, j).Text) ' sauvegarde de l'accord cliqué dans Grid5
            '
            If i > 1 Then ' lignes des tonalités du tableau Grid5
                Lab_1 = Grid5.Cell(i, j).Text 'accord sélectionné
                Lab_2 = Trim(Grid5.Cell(i, 0).Text) ' sauvegarde de la "nouvelle tonalité"
                Maj_Cadence(Lab_1, Lab_2)
                '
                LabModulat.Item(0).Text = Trim(Label33.Text)
                LabModulat.Item(1).Text = Trim(Grid5.Cell(i, j).Text) ' Accord supplémentaire (avant la cadence)
                '
                Label79.ForeColor = Color.Black
                Label79.Text = Trim(Grid5.Cell(i, 0).Text)            ' Mise à jour de la "nouvelle tonalité"
                Label80.Visible = True
            End If
        End If
    End Sub
    Sub Maj_TonsVoisins(Accord)
        Dim i As Integer
        Dim tbl() As String = Accord.split()
        Dim Note As String = Trad_BemDiesNote(tbl(0)) ' La note est exprimée en #
        Dim Typ As Integer = ComboBox9.SelectedIndex + 3 ' traduction de l'index de comobobox9 en Typ d'accord pour Mode2
        Dim chiffrage As String ' chiffrage de l'acord
        Dim cc As Char

        If tbl.Length = 1 Then
            chiffrage = ""
        Else
            chiffrage = Trim(tbl(1))
        End If
        ligneModulat = 1 ' cette variable est incrémentée dans Maj_Total
        ligne = 1
        ' Calcul de l'index i
        ' *******************
        If Len(Note) = 2 Then
            cc = Note(1)
        End If
        Dim clef1 As String = Trim(cc.ToString)
        Maj_TabNotes_Majus(clef1)
        For i = 0 To TabNotes.Length - 1
            If Note = TabNotes(i) Then
                Exit For
            End If
        Next i
        Select Case chiffrage
            Case ""
                'Maj
                Maj_Total(i, "Maj", Typ, clef1)            ' C Maj
                Maj_Total(i + 5, "Maj", Typ, clef1)        ' F Maj
                Maj_Total(i + 7, "Maj", Typ, clef1)        ' G Maj

                'MinH
                Maj_Total(i + 4, "MinH", Typ, clef1)       ' E MinH
                Maj_Total(i + 5, "MinH", Typ, clef1)       ' F MinH

                'MinM
                Maj_Total(i + 5, "MinM", Typ, clef1)       ' F MinM
                Maj_Total(i + 7, "MinM", Typ, clef1)       ' G MinM

            Case "m"
                'Maj
                Maj_Total(i + 3, "Maj", Typ, clef1)        ' Eb Maj
                Maj_Total(i + 8, "Maj", Typ, clef1)        ' Ab MaJ
                Maj_Total(i + 10, "Maj", Typ, clef1)       ' Bb Maj

                'MinH
                Maj_Total(i, "MinH", Typ, clef1)           ' C MinH
                Maj_Total(i + 7, "MinH", Typ, clef1)       ' G MinH

                'MinM
                Maj_Total(i, "MinM", Typ, clef1)           ' C  MinM
                Maj_Total(i + 10, "MinM", Typ, clef1)      ' Bb MinM

            Case "mb5" 'ex Cmb5
                'Maj
                Maj_Total(i + 1, "Maj", Typ, clef1)            ' Db Maj

                'MinH
                Maj_Total(i + 10, "MinH", Typ, clef1)          ' Bb MinH
                Maj_Total(i + 1, "MinH", Typ, clef1)           ' Db MinH

                'MinM
                Maj_Total(i + 1, "MinM", Typ, clef1)           ' Db  MinM
                Maj_Total(i + 3, "MinM", Typ, clef1)           ' Eb MinM

            Case "M7" 'ex C M7
                'Maj
                Maj_Total(i, "Maj", Typ, clef1)        ' C Maj
                Maj_Total(i + 7, "Maj", Typ, clef1)    ' G MAJ

                'MinH
                Maj_Total(i + 4, "MinH", Typ, clef1)   ' E MinH

            Case "7" 'ex C7
                'Maj
                Maj_Total(i + 5, "Maj", Typ, clef1)    ' F Maj

                'MinH
                Maj_Total(i + 5, "MinH", Typ, clef1)   ' F MinH

                'MinM
                Maj_Total(i + 5, "MinM", Typ, clef1)   ' F MinM
                Maj_Total(i + 7, "MinM", Typ, clef1)   ' G MinM

            Case "m7" 'ex Cm7 ou Eb6
                'Maj
                Maj_Total(i + 10, "Maj", Typ, clef1)       'Bb Maj
                Maj_Total(i + 8, "Maj", Typ, clef1)        'Ab MaJ
                Maj_Total(+3, "Maj", Typ, clef1)           'Eb Maj

                'MinH
                Maj_Total(i + 7, "MinH", Typ, clef1)       'G MinH

                'MinM
                Maj_Total(i + 10, "MinM", Typ, clef1)      'Bb MinM

            Case "m7b5" 'ex Cm7b5
                'Maj
                Maj_Total(i + 1, "Maj", Typ, clef1)         'Db Maj

                'MinH
                Maj_Total(i + 10, "MinH", Typ, clef1)       'Bb MinH

                'MinM
                Maj_Total(i + 1, "MinM", Typ, clef1)        'Db  MinM
                Maj_Total(i + 3, "MinM", Typ, clef1)        'Eb MinM

            Case "7Dim" 'ex C7dim"
            'MinH
            'Pas d'autres tonalités contenant accord de 7Dim

            Case "mM7"
                'MinH
                Maj_Total(i, "MinH", Typ, clef1)               ' C MinH

                'MinM
                Maj_Total(i, "MinM", Typ, clef1)               ' C MinM
        End Select

    End Sub


    Sub Effacerselections()
        Dim i, j As Integer
        For i = 2 To Grid5.Rows - 1
            For j = 1 To Grid5.Cols - 1
                Grid5.Cell(i, j).BackColor = Color.Moccasin
            Next
        Next
    End Sub
    Sub EffacerTonVoisins()
        Dim i, j As Integer
        For i = 2 To Grid5.Rows - 1
            For j = 1 To Grid5.Cols - 1
                Grid5.Cell(i, j).Text = ""
                Grid5.Cell(i, j).BackColor = Color.Moccasin
            Next
        Next
    End Sub

    Sub Maj_Total(i As Integer, Mode As String, typ As Integer, clef1 As String)
        Dim j, k As Integer
        Dim a As String
        Dim Tonique As String
        Dim tbl() As String
        Dim TonCours As String

        k = Det_TonInitial()
        a = Trim(Label32.Text)

        'Maj_TabNotes_Majus() ' cette méthode remplit TabNotes avec des notes majuscules en #
        Maj_TabNotes_Majus(clef1)
        Tonique = Trim(TabNotes(i))
        Tonique = Det_ToniqueClef(Trim(Tonique), Trim(Mode))
        TonCours = Trim(Tonique) + " " + Trim(Mode)
        If a <> TonCours Then
            ligneModulat = ligneModulat + 1

            Grid5.Cell(ligneModulat, 0).Text = Trim(TonCours)
            Grid5.Cell(ligneModulat, 0).Font = New Font("Verdana", 8, FontStyle.Regular)

            ' Calcul des accords
            Tonique = Retab_Mode(Trim(Tonique), Trim(Mode))
            a = Mode3(Trim(Tonique), Trim(Mode), typ, False)
            tbl = a.Split("-")
            ' Ecriture des accords
            For j = 0 To tbl.Length - 1
                Grid5.Cell(ligneModulat, j + 1).Text = tbl(j)
            Next j
        End If
    End Sub
    Function Det_ToniqueClef(Note As String, Mode As String) As String ' Note est transmis en #, c'est la tonique d'un Mode
        Dim a As Integer
        Det_ToniqueClef = Trim(Note)
        Clef = "#"
        If Mode = "Maj" Then
            Select Case Note
                Case "A#", "D#", "G#"
                    Det_ToniqueClef = Trad_DiesBem_Maj(Note)
                    Clef = "b"
                Case "a#", "d#", "g#"
                    a = Trad_DiesBemNote(Note)
                    a = UCaseBémol(a)
                    Clef = "b"
                Case Else
                    Clef = "#"
            End Select
        End If
    End Function
    Public Function Mode3(Tonique As String, Type_Mode As String, Typ As Integer, Voisin As Boolean) As String 'TYpe_Mode = Maj, MinH, MinM Typ= 3 -> 3note,Type=4 -> Accord7,Typ=5->Accord9,Typ=6 -> Accord11
        Dim ton As String
        Dim toni As String
        Dim tbl() As String
        Dim Sauv_Clef As String = Clef

        Mode3 = ""
        'If TabControl2.SelectedTab Is TabPage6 Then ' 3 = Onglet Modulation

        ton = LCase(Tonique)
            Maj_TabNotes_Minus(Trim(Clef))

            ' OngletCours = 0   Onglet Progression
            ' OngletCours = 1   Onglet Tonalité
            ' OngletCours = 7   Onglet Modulation
            ' OngletCours = 16  Onglet Substitution
            '

            'If OngletCours = 0 Or OngletCours = 1 Then
            'tbl = Trim(ComboBox1.Text).Split()
            'Clef = Det_Clef(Trim(tbl(0)))
            'Else
            If Type_Mode <> "Maj" Then
                toni = Det_RelativeMajeure(Trim(Tonique) + " " + Trim(Type_Mode)) ' on passe le mode en entier
                tbl = toni.Split
                toni = Trim(tbl(0))
            Else
                toni = Tonique
            End If
            Clef = Det_Clef(Trim(toni))
            '
            If Clef = "b" Then
                ton = Trad_DiesBemNote(ton)
            Else
                ton = Trad_BemDies(ton)
            End If

            Maj_TabNotes_Minus(Trim(Clef))
            '

            Select Case Type_Mode
                Case "Maj"
                    Select Case Typ
                        Case 3
                            Mode3 =
                                Tonique + "-" _
                                + UCaseBémol(NoteInterval3(ton, "9", Clef)) + " m" + "-" _
                                + UCaseBémol(NoteInterval3(ton, "3", Clef)) + " m" + "-" _
                                + UCaseBémol(NoteInterval3(ton, "4", Clef)) + "-" _
                                + UCaseBémol(NoteInterval3(ton, "5", Clef)) + "-" _
                                + UCaseBémol(NoteInterval3(ton, "6", Clef)) + " m" + "-" _
                                + UCaseBémol(NoteInterval3(ton, "M7", Clef)) + " mb5"
                        Case 4 ' 7

                            Mode3 =
                                Tonique + " M7" + "-" _
                                + UCaseBémol(NoteInterval3(ton, "9", Clef)) + " m7" + "-" _
                                + UCaseBémol(NoteInterval3(ton, "3", Clef)) + " m7" + "-" _
                                + UCaseBémol(NoteInterval3(ton, "4", Clef)) + " M7" + "-" _
                                + UCaseBémol(NoteInterval3(ton, "5", Clef)) + " 7" + "-" _
                                + UCaseBémol(NoteInterval3(ton, "6", Clef)) + " m7" + "-" _
                                + UCaseBémol(NoteInterval3(ton, "M7", Clef)) + " m7b5"
                        Case 5 ' 9

                            Mode3 =
                                Tonique + " M7(9)" + "-" _
                                + UCaseBémol(NoteInterval3(ton, "9", Clef)) + " m7(9)" + "-" _
                                + "___" + "-" _
                                + UCaseBémol(NoteInterval3(ton, "4", Clef)) + " M7(9)" + "-" _
                                + UCaseBémol(NoteInterval3(ton, "5", Clef)) + " 7(9)" + "-" _
                                + UCaseBémol(NoteInterval3(ton, "6", Clef)) + " m7(9)" + "-" _
                                + "___" + "-"
                        Case 6 ' 11

                            Mode3 =
                                Tonique + " 11" + "-" _
                                + UCaseBémol(NoteInterval3(ton, "9", Clef)) + " m11" + "-" _
                                + UCaseBémol(NoteInterval3(ton, "3", Clef)) + " m11" + "-" _
                                + UCaseBémol(NoteInterval3(ton, "4", Clef)) + " M7(11#)" + "-" _
                                + UCaseBémol(NoteInterval3(ton, "5", Clef)) + " 7(11)" + "-" _
                                + UCaseBémol(NoteInterval3(ton, "6", Clef)) + " m11" + "-" _
                                + "___"
                    End Select
                Case "MinH"
                    Select Case Typ
                        Case 3
                            Mode3 =
                                Tonique + " m" + "-" _
                                + UCaseBémol(NoteInterval3(ton, "9", Clef)) + " mb5" + "-" _
                                + UCaseBémol(NoteInterval3(ton, "3m", Clef)) + " 5#" + "-" _
                                + UCaseBémol(NoteInterval3(ton, "4", Clef)) + " m" + "-" _
                                + UCaseBémol(NoteInterval3(ton, "5", Clef)) + "-" _
                                + UCaseBémol(NoteInterval3(ton, "5#", Clef)) + "-" _
                                + UCaseBémol(NoteInterval3(ton, "M7", Clef)) + " mb5"
                        Case 4 ' 7
                            Mode3 =
                                Tonique + " mM7" + "-" _
                                + UCaseBémol(NoteInterval3(ton, "9", Clef)) + " m7b5" + "-" _
                                + UCaseBémol(NoteInterval3(ton, "3m", Clef)) + " M75#" + "-" _
                                + UCaseBémol(NoteInterval3(ton, "4", Clef)) + " m7" + "-" _
                                + UCaseBémol(NoteInterval3(ton, "5", Clef)) + " 7" + "-" _
                                + UCaseBémol(NoteInterval3(ton, "5#", Clef)) + " M7" + "-" _
                                + UCaseBémol(NoteInterval3(ton, "M7", Clef)) + " 7Dim"
                        Case 5 ' 9

                            Mode3 =
                                "___" + "-" _
                                + "___" + "-" _
                                + "___" + "-" _
                                + UCaseBémol(NoteInterval3(ton, "4", Clef)) + " m7(9)" + "-" _
                                + UCaseBémol(NoteInterval3(ton, "5", Clef)) + " 7(b9)" + "-" _
                                + UCaseBémol(NoteInterval3(ton, "5#", Clef)) + " 7(9#)" + "-" _
                                + "___"
                        Case 6 '11
                            Mode3 =
                                  Tonique + " m11" + "-" _
                                + "___" + "-" _
                                + "___" + "-" _
                                + "___" + "-" _
                                + UCaseBémol(NoteInterval3(ton, "5", Clef)) + " 7(11)" + "-" _
                                + UCaseBémol(NoteInterval3(ton, "5#", Clef)) + " M7(11#)" + "-" _
                                + "___"

                    End Select
                Case "MinM"
                    Select Case Typ
                        Case 3
                            Mode3 =
                            Tonique + " m" + "-" _
                            + UCaseBémol(NoteInterval3(ton, "9", Clef)) + " m" + "-" _
                            + UCaseBémol(NoteInterval3(ton, "3m", Clef)) + " 5#" + "-" _
                            + UCaseBémol(NoteInterval3(ton, "4", Clef)) + "-" _
                            + UCaseBémol(NoteInterval3(ton, "5", Clef)) + "-" _
                            + UCaseBémol(NoteInterval3(ton, "6", Clef)) + " mb5" + "-" _
                            + UCaseBémol(NoteInterval3(ton, "M7", Clef)) + " mb5"

                        Case 4 ' 7
                            Mode3 =
                            Tonique + " mM7" + "-" _
                            + UCaseBémol(NoteInterval3(ton, "9", Clef)) + " m7" + "-" _
                            + UCaseBémol(NoteInterval3(ton, "3m", Clef)) + " M75#" + "-" _
                            + UCaseBémol(NoteInterval3(ton, "4", Clef)) + " 7" + "-" _
                            + UCaseBémol(NoteInterval3(ton, "5", Clef)) + " 7" + "-" _
                            + UCaseBémol(NoteInterval3(ton, "6", Clef)) + " m7b5" + "-" _
                            + UCaseBémol(NoteInterval3(ton, "M7", Clef)) + " m7b5"
                        Case 5 '9

                            Mode3 =
                                "___" + "-" _
                                + "___" + "-" _
                                + "___" + "-" _
                                + UCaseBémol(NoteInterval3(ton, "4", Clef)) + " 7(9)" + "-" _
                                + UCaseBémol(NoteInterval3(ton, "5", Clef)) + " 7(9)" + "-" _
                                + "___" + "-" _
                                + "___"
                        Case 6 '11
                            Mode3 =
                                Tonique + " m11" + "-" _
                                + UCaseBémol(NoteInterval3(ton, "9", Clef)) + " m11" + "-" _
                                + "___" + "-" _
                                + "___" + "-" _
                                + UCaseBémol(NoteInterval3(ton, "5", Clef)) + " 7(11)" + "-" _
                                + "___" + "-" _
                                + "___"
                    End Select
            End Select
            Clef = Sauv_Clef
        'End If
    End Function
    Sub RecheraccordGrid5(Accord As String)
        Dim tbl() As String

        Dim i, j As Integer
        tbl = Accord.Split()
        If Len(tbl(0)) = 1 Then
            For i = 2 To Grid5.Rows - 1
                For j = 1 To Grid5.Cols - 1
                    If Trim(Accord) = Trim(Grid5.Cell(i, j).Text) Then
                        Grid5.Cell(i, j).BackColor = Color.Tan
                    End If
                Next
            Next
        Else
            Dim note As String = Trad_Acc_DB(tbl(0))
            Dim Accord2 As String
            If tbl.Count = 1 Then
                Accord2 = Trim(note)
            Else
                Accord2 = note + " " + tbl(1)
            End If
            For i = 2 To Grid5.Rows - 1
                For j = 1 To Grid5.Cols - 1
                    If Trim(Accord) = Trim(Grid5.Cell(i, j).Text) Or Trim(Accord2) = Trim(Grid5.Cell(i, j).Text) Then
                        Grid5.Cell(i, j).BackColor = Color.Tan
                    End If
                Next
            Next
        End If
    End Sub

    Function Trad_Acc_DB(note As String) As String
        Dim cc As Char
        Dim a As String
        cc = note(1)
        a = cc.ToString
        '
        If a = "#" Then
            Select Case note
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
                    Return note

            End Select
        Else
            Select Case note
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
                    Return note
            End Select
        End If
    End Function




    Sub Maj_Cadence(Accord As String, Ton As String)


        If Trim(Accord) <> "---" And Trim(Ton) <> "---" Then
            Dim i, j As Integer
            Dim a As String = ""
            Dim b As String = ""

            Dim tbl1() As String
            Dim tbl2() As String
            Dim tbl3() As String
            '
            ' Détermination des notes du ton dans un tableau
            ' *********************************************
            For i = 2 To Grid5.Rows - 1  ' Rechercher la ligne i du mode dans le tableau
                If Trim(Grid5.Cell(i, 0).Text = Trim(Ton)) Then
                    Exit For
                End If
            Next
            For j = 1 To 7  ' détermination des notes du ton
                a = Trim(Grid5.Cell(i, j).Text)
                b = b + a + "-"
            Next
            b = Microsoft.VisualBasic.Left(b, Len(b) - 1)
            tbl1 = b.Split("-") ' les notes du ton sont placées dans un tableau avec index de 0 à 7
            '
            ' Détermination des 2 degrés de la cadence
            ' ****************************************
            tbl2 = ComboBox13.Text.Split("-")
            tbl3 = Trim(tbl2(1)).Split()
            ' Détermination des 2 accords de la cadence
            ' ***************************************
            LabModulat.Item(2).Text = tbl1(Det_IndexDegré(tbl3(0)))
            LabModulat.Item(3).Text = tbl1(Det_IndexDegré(tbl3(1)))
            '
        End If
    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub

    Private Sub ComboBox9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox9.SelectedIndexChanged
        If Not EnChargement Then
            Dim i As Integer
            '
            EffacerTonVoisins()
            '
            ' Mise à jour de la tonalité en cours
            ' ***********************************
            For i = 0 To 6
                Grid5.Cell(1, i + 1).Text = TableGlobalAcc(ComboBox9.SelectedIndex, RadioModulat_SelectedIndex, i) 'tbl1(i) ' mise de la gamme de C Maj dans l'onglet Modulation
            Next i
            '
        End If
    End Sub

    Function RadioModulat_SelectedIndex() As Integer
        Dim i As Integer
        For i = 0 To RadioModulat.Count - 1
            If RadioModulat.Item(i).Checked Then
                Exit For
            End If
        Next

        Return i
    End Function

    Private Sub ComboBox13_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox13.SelectedIndexChanged
        If Not EnChargement Then
            Maj_Cadence(Lab_1, Lab_2)
        End If
    End Sub

    Private Sub RadioB1_CheckedChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs)
        a = "c d e"
        Dim L As List(Of String)
        L = Assist_Tonalité(a)
    End Sub
    Function GenerateModes(tonic As Integer, notes As List(Of Integer)) As List(Of String)
        Dim modes As Dictionary(Of String, List(Of Integer)) = New Dictionary(Of String, List(Of Integer)) From {
            {"Ionian", New List(Of Integer) From {0, 2, 4, 5, 7, 9, 11}},
            {"Dorian", New List(Of Integer) From {0, 2, 3, 5, 7, 9, 10}},
            {"Phrygian", New List(Of Integer) From {0, 1, 3, 5, 7, 8, 10}},
            {"Lydian", New List(Of Integer) From {0, 2, 4, 6, 7, 9, 11}},
            {"Mixolydian", New List(Of Integer) From {0, 2, 4, 5, 7, 9, 10}},
            {"Aeolian", New List(Of Integer) From {0, 2, 3, 5, 7, 8, 10}},
            {"Locrian", New List(Of Integer) From {0, 1, 3, 5, 6, 8, 10}}
        }

        Dim possibleModes As List(Of String) = New List(Of String)() ' pour le résultat

        For Each Mode1 In modes

            Dim modeName As String = Mode1.Key ' nom du mode
            Dim intervals As List(Of Integer) = Mode1.Value ' intervalles du mode
            Dim modeNotes As List(Of Integer) = New List(Of Integer)()

            ' Calcul des notes dans la gamme chromatique
            ' ******************************************       
            For Each interval In intervals
                modeNotes.Add((tonic + interval) Mod 12) ' tonic correspond à noteindex dans la procédure appelante
            Next

            Dim allNotesInMode As Boolean = True ' flag pour sortie du For suivant

            For Each note In modeNotes
                If Not notes.Contains(note) Then ' notes liste des notes entrées ( en paramètre dans la fonction
                    allNotesInMode = False ' on sort de la boucle si 
                    Exit For
                End If
            Next

            If allNotesInMode Then
                possibleModes.Add(modeName)
            End If
        Next

        Dim allNotes As List(Of String) = New List(Of String) From {"C", "C#", "D", "D#", "E", "F", "F#", "G", "G#", "A", "A#", "B"}

        'Console.WriteLine("Veuillez entrer les cinq notes parmi les 11 notes de musique (séparées par des espaces) :")
        Dim ligne As String = "D" + " " + "E" ' "C" + " " + 
        Dim inputNotes() As String = ligne.Split 'Console.ReadLine().Split()

        Dim inputNotesIndices As List(Of Integer) = New List(Of Integer)()
        For Each note In inputNotes
            inputNotesIndices.Add(allNotes.IndexOf(note.ToUpper())) ' calcul de l'index dans la listemise en majuscules pour la liste allnotes
        Next ' cacul 

        'Console.WriteLine(vbCrLf & "Modes majeurs possibles :")
        For Each noteIndex In inputNotesIndices
            'Dim possibleModes As List(Of String) = GenerateModes(noteIndex, inputNotesIndices)
            'Console.WriteLine($"Pour la note {allNotes(noteIndex)} : {String.Join(", ", possibleModes)}")
            Dim msgBoxResult = (possibleModes)
        Next
        Return possibleModes
    End Function
    Public Function Assist_Tonalité(listNotes As String) As List(Of String)
        Dim a As String = "nok"
        Dim tbl() As String = listNotes.Split()
        Dim listNotes2 As New List(Of String)
        Dim listInterval As New List(Of Integer)
        Dim ListResult As New List(Of String)
        Dim nbNotes As Integer = tbl.Count
        Dim i, ii, j As Integer


        Dim modes As Dictionary(Of String, List(Of Integer)) = New Dictionary(Of String, List(Of Integer)) From {
        {"Maj", New List(Of Integer) From {0, 2, 4, 5, 7, 9, 11}},
        {"MinH", New List(Of Integer) From {0, 2, 3, 5, 7, 8, 11}},
        {"MinM", New List(Of Integer) From {0, 2, 3, 5, 7, 9, 11}},
        {"PentaMin", New List(Of Integer) From {0, 3, 5, 7, 10}}}

        ' 1 - Recopie en double des notes dans la listNotes2
        ' **************************************************
        For i = 0 To 1
            For j = 0 To tbl.Count - 1
                listNotes2.Add(tbl(j).ToUpper) ' copier 2 fois et mettre en majuscule
            Next
        Next
        '
        ' 2 -  Calcul des modes d'appartenance.
        ' ************************************
        For i = 0 To nbNotes - 1
            ' Cacul des intervalles entre les notes en considérant une comme tonique
            For j = i To ((nbNotes - 1) + i)
                ii = Calcul_Intervalle(listNotes2.Item(i), listNotes2.Item(j))
                listInterval.Add(ii)
            Next j
            ' détermination des modes contenant les notes
            For Each oo As String In modes.Keys
                If AppartientA2(listInterval, modes(oo)) Then ' oo est la clef - mode(oo) est la liste
                    ListResult.Add(listNotes2.Item(i) + " " + oo)
                End If
            Next
            listInterval.Clear()
        Next
        ' 
        Return ListResult
    End Function
    Function Assist_Tonalité2(listNotes As List(Of String)) As List(Of String)

        Dim listNotes2 As New List(Of String)
        Dim listInterval As New List(Of Integer)
        Dim ListResult As New List(Of String)
        Dim LModes As New List(Of String)

        ' Version 1 de la liste avec Blues 2
        ' **********************************
        'Dim modes As Dictionary(Of String, List(Of Integer)) = New Dictionary(Of String, List(Of Integer)) From {
        '{"Maj", New List(Of Integer) From {0, 2, 4, 5, 7, 9, 11}},
        '{"MinH", New List(Of Integer) From {0, 2, 3, 5, 7, 8, 11}},
        '{"MinM", New List(Of Integer) From {0, 2, 3, 5, 7, 9, 11}},
        '{"PentaMin", New List(Of Integer) From {0, 3, 5, 7, 10}},
        '{"Blues", New List(Of Integer) From {0, 3, 5, 6, 7, 10}},
        '{"Blues2", New List(Of Integer) From {0, 2, 3, 4, 5, 6, 7, 9, 10}}
        '}
        ' Version 1 de la liste sans Blues 2
        ' **********************************
        Dim modes As Dictionary(Of String, List(Of Integer)) = New Dictionary(Of String, List(Of Integer)) From {
        {"Maj", New List(Of Integer) From {0, 2, 4, 5, 7, 9, 11}},
        {"MinH", New List(Of Integer) From {0, 2, 3, 5, 7, 8, 11}},
        {"MinM", New List(Of Integer) From {0, 2, 3, 5, 7, 9, 11}},
        {"PentaMin", New List(Of Integer) From {0, 3, 5, 7, 10}},
        {"Blues", New List(Of Integer) From {0, 3, 5, 6, 7, 10}}}


        Dim LN As New List(Of String) From {"c", "c#", "d", "d#", "e", "f", "g", "g#", "a", "a#", "b"}
        Dim LN2 As New List(Of String) From {"c", "c#", "d", "d#", "e", "f", "f#", "g", "g#", "a", "a#", "b", "c", "c#", "d", "d#", "e", "f", "f#", "g", "g#", "a", "a#", "b"}
        Dim Gamme As New List(Of String)
        Dim Result As New List(Of String)

        'Maj_TabNotes_Minus("#")
        '
        For Each clef As KeyValuePair(Of String, List(Of Integer)) In modes
            For Each Ton As String In LN
                For Each interv As Integer In modes(clef.Key)
                    i = LN2.IndexOf(Ton)
                    Gamme.Add(LN2(i + interv))
                Next
                If AppartientA3(listNotes, Gamme) Then
                    Result.Add(Ton.ToUpper + " " + clef.Key)
                End If
                Gamme.Clear()
            Next
        Next
        ' 
        Return Result
    End Function
    Function Calcul_Intervalle(note1 As String, note2 As String) As Integer
        Dim i, interv, ii As Integer
        Dim allNotes As List(Of String) = New List(Of String) From {"C", "C#", "D", "D#", "E", "F", "F#", "G", "G#", "A", "A#", "B",
            "C", "C#", "D", "D#", "E", "F", "F#", "G", "G#", "A", "A#", "B"}

        i = allNotes.IndexOf(note1) ' index de note2 dans allNotes
        interv = -1
        For ii = i To allNotes.Count - 1
            interv = (interv + 1)
            If allNotes.Item(ii) = note2 Then
                Exit For
            End If
        Next

        Return interv
    End Function
    Function AppartientA(ByVal Linterv As List(Of Integer), ByVal LMode As List(Of Integer)) As Boolean
        Dim flag As Boolean = True
        For Each oo As Integer In Linterv
            If Not LMode.Contains(oo) Then
                flag = False
                Exit For
            End If
        Next
        Return flag
    End Function
    Function AppartientA2(ByVal Linterv As List(Of Integer), ByVal LMode As List(Of Integer)) As Boolean
        Dim flag As Boolean = True
        For Each oo As Integer In Linterv
            If Not LMode.Contains(oo) Then
                flag = False
                Exit For
            End If
        Next
        Return flag
    End Function
    Function AppartientA3(ByVal Linterv As List(Of String), ByVal LMode As List(Of String)) As Boolean
        Dim flag As Boolean = True
        For Each oo As String In Linterv
            If Not LMode.Contains(oo) Then
                flag = False
                Exit For
            End If
        Next
        Return flag
    End Function

    Private Sub Standard_Click(sender As Object, e As EventArgs) Handles Standard.Click
        Me.Cursor = Cursors.WaitCursor
        'Cacher_FormTransparents()
        Dim result As Integer = NouveauProjet()
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Quick_Blues_Click(sender As Object, e As EventArgs) Handles Quick_Blues.Click
        Me.Cursor = Cursors.WaitCursor
        'Cacher_FormTransparents()
        Dim result As Integer = NouveauProjet()
        If result <> DialogResult.Cancel Then
            QBlues2.ShowDialog()
        End If
        Me.BringToFront()
        Me.Calcul_AutoVoicingZ()
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button8_Click_1(sender As Object, e As EventArgs)
        Dim rand As New Random() ' Créez une instance de la classe Random

        Dim valeurMinimale As Integer = 84 ' Définissez la valeur minimale souhaitée
        Dim valeurMaximale As Integer = 94 ' Définissez la valeur maximale souhaitée

        ' Utilisez la méthode Next() de l'objet Random pour générer un nombre aléatoire
        ' compris entre valeurMinimale (inclus) et valeurMaximale (exclus)
        Dim nombreAleatoire As Integer = rand.Next(valeurMinimale, valeurMaximale)
    End Sub

    Private Sub Magneto1_Click(sender As Object, e As EventArgs) Handles Magneto1.Click

    End Sub

    Private Sub Button8_Click_2(sender As Object, e As EventArgs) Handles Button8.Click
        If Not (SortieMidi.Item(ChoixSortieMidi).IsOpen) Then
            SortieMidi.Item(ChoixSortieMidi).Open()
        End If
        SortieMidi.Item(ChoixSortieMidi).SendNoteOn(0, ValNoteCubase.IndexOf("C3"), 64)
        'SortieMidi.Item(ChoixSortieMidi).SendPitchBend()
    End Sub

    Private Sub Button12_Click_1(sender As Object, e As EventArgs) Handles Button12.Click
        If Not (SortieMidi.Item(ChoixSortieMidi).IsOpen) Then
            SortieMidi.Item(ChoixSortieMidi).Open()
        End If
        SortieMidi.Item(ChoixSortieMidi).SendNoteOff(0, ValNoteCubase.IndexOf("C3"), 64)

    End Sub

    Private Sub Button13_Click_1(sender As Object, e As EventArgs) Handles Button13.Click
        If Not (SortieMidi.Item(ChoixSortieMidi).IsOpen) Then
            SortieMidi.Item(ChoixSortieMidi).Open()
        End If
        ' Dim a As String = &H7F00 ' "0111111100000000"
        SortieMidi.Item(ChoixSortieMidi).SendPitchBend(0, 16383)
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        If Not (SortieMidi.Item(ChoixSortieMidi).IsOpen) Then
            SortieMidi.Item(ChoixSortieMidi).Open()
        End If
        SortieMidi.Item(ChoixSortieMidi).SendPitchBend(0, 8192)
    End Sub

    Private Sub Label23_Click(sender As Object, e As EventArgs) Handles Label23.Click

    End Sub

    Private Sub FiltreUni_CheckedChanged(sender As Object, e As EventArgs) Handles FiltreUni.CheckedChanged

    End Sub

    Private Sub TabPage10_Click(sender As Object, e As EventArgs) Handles TabPage10.Click

    End Sub

    Private Sub Tab_ChordRoll2_Click(sender As Object, e As EventArgs) Handles Tab_ChordRoll2.Click

    End Sub
End Class