Imports System.Windows.Interop
Imports FlexCell
Imports HyperArp.Module1

Public Class Automation
    Public Orig_Autom_Vérrouillage As Boolean = False ' variable utilsée pour désactiver le tracé d'un courbe le temps d'un "Coller"
    Public WriteOnly Property POrig_Autom_Vérrouillage As Boolean
        Set(ByVal value As Boolean)
            Me.Orig_Autom_Vérrouillage = value
        End Set
    End Property
    Private NumCanal As Integer = 0
    Public ReadOnly Property PNcanal As Integer
        Get
            Return NumCanal
        End Get
    End Property
    Private NumCourbe As Integer = 0
    Public ReadOnly Property PNCourbe As Integer
        Get
            Return NumCourbe
        End Get
    End Property
    Private Langue As String
    Public WriteOnly Property PLangue As String
        Set(ByVal value As String)
            Me.Langue = value
        End Set
    End Property
    Private ListAcc As String
    Public WriteOnly Property PListAcc As String
        Set(ByVal value As String)
            Me.ListAcc = value
        End Set
    End Property
    Private ListMarq As String
    Public WriteOnly Property PListMarq As String
        Set(ByVal value As String)
            Me.ListMarq = value
        End Set
    End Property
    Public Métrique As String = "4/4"
    Public WriteOnly Property PMétrique As String
        Set(ByVal value As String)
            Me.Métrique = value
        End Set
    End Property
    ' Constantes
    Private Const TailleCel = 18
    Private Const TailleColFix = 10
    Private Const TailleCanaux = 1150
    Private Const FenêtreWidth = 1180
    Private Const FenêtreHeight = 770 ' 380
    Private Const HautLigneCourbes = 2
    '
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
    ' couleur des courbes
    'ReadOnly CoulExp As Color = Color.FromArgb(&HCCCC99) 'Color.Blue
    'ReadOnly CoulMod As Color = Color.LightSeaGreen 'Color.Ivory 'Color.LemonChiffon 'Color.FromArgb(&HFFCC66) 'Color.LightSeaGreen
    'ReadOnly CoulPan As Color = Color.Red 'FromArgb(&HCCCC99) 'Color.Blue
    'ReadOnly CoulCC50 As Color = Color.YellowGreen
    'ReadOnly CoulCC51 As Color = Color.Orange
    'ReadOnly CoulCC52 As Color = Color.MediumTurquoise
    'ReadOnly CoulCC53 As Color = Color.Thistle
    'ReadOnly CoulPB As Color = Color.Yellow
    '
    ReadOnly CoulBarrout As Color = Color.FromArgb(&HEEEEEE) 'Color.AliceBlue

    Private ToolTip11 As New ToolTip

    ' composants du formulaire
    Public F4 As New Form
    Dim Panneau1 As New SplitContainer
    Dim Panneau2 As New SplitContainer
    Dim Panneau3 As New SplitContainer
    Dim PaneldesCourbes As New Panel ' panel destiné à porter les courbes et leur barre d'outils verticale c.a.d. destiné à recevoir Panneau2, ce Panel permet de scroller tout le Panneau2 dans son entier
    '
    Class Canal
        Public GridCourbes As New List(Of Grid) ' grille

        Public PageCourbes As New List(Of TabPage) ' onglet

        Public Mute As New CheckBox
    End Class
    Public Canaux As New List(Of Canal)
    Public TabCourbes As New List(Of TabControl) ' table d'onglets
    Public MidiLearn As New CheckBox
    Public AffMidiLearn As New Label '
    ' Elément de la barre d'outils principales
    ' ****************************************
    Dim LSyncScroll As New List(Of RadioButton)
    Dim GroupScroll As New GroupBox
    Dim DockButton As New Button

    Dim BoutonTaille As New Button
    Dim BoutonTaillePréced As Integer = 770


    ' Entête
    ' ******
    Dim GridEntête As New FlexCell.Grid
    '
    ' Mute et noms des Canaux
    ' ***********************
    Public ListCCAct As New List(Of SLCCAct)
    Class SLCCAct
        Public LCCAct As New List(Of CheckBox)
    End Class
    '
    ' Aller vers
    ' **********
    Dim LabMesures As New Label
    Dim LabMarqueurs As New Label
    Dim TextMesures As New TextBox
    Dim ListMarqueurs As New System.Windows.Forms.ComboBox
    Dim ListMarqRefresh As New List(Of String)


    Dim ListPréced As New List(Of Button)
    Dim ListSuiv As New List(Of Button)
    '
    Private Numérateur As Integer
    Private Dénominateur As Integer
    Private DivisionMes As Integer

    ' Font
    ReadOnly fnt4 As New System.Drawing.Font("Calibri", 11, FontStyle.Regular)
    ReadOnly fnt5 As New System.Drawing.Font("Calibri", 12, FontStyle.Regular)
    ReadOnly fnt6 As New System.Drawing.Font("Verdana", 8, FontStyle.Regular)
    ReadOnly fnt7 As New System.Drawing.Font("Verdana", 7, FontStyle.Regular)
    ReadOnly fnt7_1 As New System.Drawing.Font("Verdana", 7, FontStyle.Bold)
    ReadOnly fnt8 As New System.Drawing.Font("Verdana", 6, FontStyle.Regular)
    ReadOnly fnt9 As New System.Drawing.Font("Verdana", 8, FontStyle.Regular)
    '
    ' Menu
    ' ****
    Private Menu1 As New System.Windows.Forms.MenuStrip()
    Private Edition As New System.Windows.Forms.ToolStripMenuItem()
    Private Couper As New System.Windows.Forms.ToolStripMenuItem()
    Private Copier As New System.Windows.Forms.ToolStripMenuItem()
    Private Coller As New System.Windows.Forms.ToolStripMenuItem()
    Private Séparateur1 As New System.Windows.Forms.ToolStripSeparator()
    Private Annuler As New System.Windows.Forms.ToolStripMenuItem()
    Private Séparateur2 As New System.Windows.Forms.ToolStripSeparator()
    Private Supprimer As New System.Windows.Forms.ToolStripMenuItem()

    Private Fichier As New System.Windows.Forms.ToolStripMenuItem()
    Private Quitter As New System.Windows.Forms.ToolStripMenuItem()


    Sub New()
        Me.Numérateur = Det_Numérateur(Me.Métrique)
        Me.Dénominateur = Det_Dénominateur(Me.Métrique)
        Me.DivisionMes = Det_DivisionMes()
        '
        Construction_Formulaire()
        Construction_GridEntête()
        Construction_Courbes()
        '
        AddHandler DockButton.MouseClick, AddressOf Dockbutton_MouseClick
        AddHandler F4.KeyUp, AddressOf F4_KeyUp
        '
        Me.F4.KeyPreview = True ' pour réception de la touche F4 et F5 pour Recalcul
        Panneau1.Panel2.Enabled = False ' Les courbes d'automation ne sont pas utilisables tant que le formulaire n'est pas détaché.
    End Sub
    Private Sub F4_KeyUp(sender As Object, e As KeyEventArgs)
        ' Play, Recalcul et Stop
        ' **********************
        If e.KeyCode = Keys.F5 Then
            Form1.ReCalcul()
        End If
        '
        If e.KeyCode = Keys.F4 Then
            Form1.StopPlay()
        End If
    End Sub
    Public Sub Construction_Formulaire()
        ' Formulaire de base
        F4.TopMost = False   ' un seul des 2 suffit ?
        F4.TopLevel = False ' absolument nécessaire pour intégrer un formulaire dans un container.
        F4.FormBorderStyle = FormBorderStyle.None
        F4.ControlBox = False
        F4.Text = "HyperArp : Automation"
        'F4.Width = fenêtre
        '
        ' PANNEAU1 : panel1 reçoit l'entête et panel les courbes
        ' *******************************************************
        Panneau1.Orientation = Orientation.Horizontal
        Panneau1.Dock = DockStyle.Fill
        Panneau1.Width = 1200
        Panneau1.Panel2.AutoScroll = True
        Panneau1.Panel2.BackColor = Color.FromArgb(&HEEEEEE) 'color.red
        Panneau1.Panel1.BackColor = CoulBarrout 'Color.Beige
        'Panneau1.Panel2.BackColor = CoulBarrout 'Color.Beige
        Panneau1.FixedPanel = FixedPanel.Panel1
        Panneau1.SplitterDistance = 57
        Panneau1.IsSplitterFixed = True
        Panneau1.Visible = True
        '
        ' PANNEAU3 : Configuration du splitcontainer
        ' ******************************************
        Panneau3.Orientation = Orientation.Horizontal
        Panneau3.Panel2.Controls.Add(Panneau1)
        Panneau3.Dock = DockStyle.Fill
        Panneau3.Panel2.AutoScroll = True
        Panneau3.Panel1.BackColor = CoulBarrout ' Color.Beige
        Panneau3.Panel2.BackColor = CoulBarrout ' Color.Beige
        Panneau3.FixedPanel = FixedPanel.Panel1
        Panneau3.IsSplitterFixed = False
        Panneau3.SplitterDistance = 48
        Panneau3.IsSplitterFixed = True
        Panneau3.Visible = True
        Panneau3.Panel2.Controls.Add(Panneau1)
        '
        F4.Controls.Add(Panneau3) ' incorporation du splitcontainer dans le formulaire
    End Sub
    Public Sub Construction_BarroutLocale()
        Dim i As Integer
        Dim p As Point

        For i = 0 To Module1.nb_PistesVar - 1

            '
            p.Y = (TabCourbes.Item(i).Size.Height * i) + 3
            '
            Panneau1.Panel2.Controls.Add(Canaux.Item(i).Mute)
            With Canaux.Item(i).Mute
                .Size = New Size(20, 20)
                .AutoSize = True
                If Module1.LangueIHM = "fr" Then
                    .Text = "PISTE " + (i + 1).ToString '+ " - CTRLs"
                Else
                    .Text = "TRACK " + (i + 1).ToString '+ " - CTRLs"
                End If
                .Checked = False
                p.X = 415
                .Location = p
                .Visible = True
                .Tag = i
                .BringToFront()
            End With
            '
            AddHandler Canaux.Item(i).Mute.CheckedChanged, AddressOf CanauxMute_ChekedChange

            ListCCAct.Add(New SLCCAct)
            For j = 0 To nbCourbes - 1
                ListCCAct.Item(i).LCCAct.Add(New CheckBox)
                Panneau1.Panel2.Controls.Add(ListCCAct.Item(i).LCCAct.Item(j))
                With ListCCAct.Item(i).LCCAct.Item(j)
                    .Visible = True
                    .Size = New Size(20, 20)
                    .Checked = True
                    If j > 2 Then
                        .Checked = False
                    End If
                    .AutoSize = True
                    .Font = fnt6
                    Select Case j
                        Case 0
                            .Text = "Exp"
                            p.X = 585
                            .Location = p
                        Case 1
                            .Text = "Mod"
                            p.X = 650
                            .Location = p
                        Case 2
                            .Text = "Pan"
                            p.X = 715
                            .Location = p
                        Case 3
                            .Text = "CC50"
                            p.X = 770
                            .Location = p
                        Case 4
                            .Text = "CC51"
                            p.X = 850
                            .Location = p
                        Case 5
                            .Text = "CC52"
                            p.X = 915
                            .Location = p
                        Case 6
                            .Text = "CC53"
                            p.X = 985
                            .Location = p
                    End Select

                    .BringToFront()
                End With
            Next j
        Next i
        Init_CheckPISTE()
    End Sub
    Sub CanauxMute_ChekedChange(sender As Object, e As EventArgs)
        Dim com As CheckBox = sender
        Dim ind As Integer = Val(com.Tag)
        Dim i As Integer


        If Canaux.Item(ind).Mute.Checked Then
            For i = 0 To nbCourbes - 1
                ListCCAct.Item(ind).LCCAct.Item(i).Enabled = True
            Next
        Else
            For i = 0 To nbCourbes - 1
                ListCCAct.Item(ind).LCCAct.Item(i).Enabled = False
            Next
        End If

    End Sub
    Sub Init_CheckPISTE()

        Dim i1, i2 As Integer
        For i1 = 0 To Module1.nb_PistesVar - 1
            Canaux.Item(i1).Mute.Checked = False
            For i2 = 0 To nbCourbes - 1
                ListCCAct.Item(i1).LCCAct.Item(i2).Enabled = False
                ListCCAct.Item(i1).LCCAct.Item(i2).Checked = False
            Next
        Next i1
    End Sub


    Public Sub Maj_Tooltips()
        Dim i As Integer

        ToolTip11.RemoveAll() ' le RemoveAll est obligatoire pour faire réapparaître les bulles à chaque Undock (autrement, elles n'apparaissent que sur le 1er Undock) 
        ToolTip11.InitialDelay = 250
        ToolTip11.Active = True
        ToolTip11.ShowAlways = True

        For i = 0 To Module1.nb_PistesVar - 1

            If Module1.LangueIHM = "fr" Then
                ToolTip11.SetToolTip(Canaux.Item(i).Mute, "Activer les controleurs")
                ToolTip11.SetToolTip(ListCCAct.Item(i).LCCAct.Item(0), "Expression - 11")
                ToolTip11.SetToolTip(ListCCAct.Item(i).LCCAct.Item(1), "Modulation - 1")
                ToolTip11.SetToolTip(ListCCAct.Item(i).LCCAct.Item(2), "PAN - 10")
                ToolTip11.SetToolTip(ListCCAct.Item(i).LCCAct.Item(3), "Libre - 50")
                ToolTip11.SetToolTip(ListCCAct.Item(i).LCCAct.Item(4), "Libre - 51")
                ToolTip11.SetToolTip(ListCCAct.Item(i).LCCAct.Item(5), "Libre - 52")
                ToolTip11.SetToolTip(ListCCAct.Item(i).LCCAct.Item(6), "Libre - 53")
            Else
                ToolTip11.SetToolTip(Canaux.Item(i).Mute, "Activate the controlers")
                ToolTip11.SetToolTip(ListCCAct.Item(i).LCCAct.Item(0), "Expression - 11")
                ToolTip11.SetToolTip(ListCCAct.Item(i).LCCAct.Item(1), "Modulation - 1")
                ToolTip11.SetToolTip(ListCCAct.Item(i).LCCAct.Item(2), "PAN - 10")
                ToolTip11.SetToolTip(ListCCAct.Item(i).LCCAct.Item(3), "Free - 50")
                ToolTip11.SetToolTip(ListCCAct.Item(i).LCCAct.Item(4), "Free - 51")
                ToolTip11.SetToolTip(ListCCAct.Item(i).LCCAct.Item(5), "Free - 52")
                ToolTip11.SetToolTip(ListCCAct.Item(i).LCCAct.Item(6), "Free - 53")

            End If
        Next i
    End Sub

    Public Sub Construction_Barrout()
        Dim p As New Point
        Dim s As New Size
        ' 
        ' 
        ' Bouton Attacher / Détacher
        ' **************************
        p.X = 5
        p.Y = 5
        s.Width = 70
        s.Height = 35
        '
        Panneau3.Panel1.Controls.Add(DockButton)
        '
        DockButton.Size = s
        DockButton.Location = p
        DockButton.Enabled = True
        DockButton.BackColor = Color.DarkOliveGreen ' 
        DockButton.ForeColor = Color.Yellow
        '
        DockButton.Font = fnt6
        'DockButton.AutoSize = True
        DockButton.Visible = True
        '
        If Langue = "fr" Then
            DockButton.Text = "Détacher"
        Else
            DockButton.Text = "UnDock"
        End If

        ' Bouton radio synchro scroll
        ' ***************************
        p.Y = 0
        Panneau3.Panel1.Controls.Add(GroupScroll)
        p.X = p.X + 100
        GroupScroll.Text = "Scroll Mode"
        s.Width = 150
        s.Height = 47
        GroupScroll.Size = s
        GroupScroll.Location = p
        GroupScroll.Font = fnt7

        p.Y = 20
        For i = 0 To 1
            LSyncScroll.Add(New RadioButton)
            GroupScroll.Controls.Add(LSyncScroll.Item(i))
            With LSyncScroll.Item(i)
                .Size = New Size(20, 20)
                .AutoSize = True
                .Visible = True
                .Enabled = True
                Select Case i
                    Case 0
                        If Me.Langue = "fr" Then
                            .Text = "Courant"
                        Else
                            .Text = "Current"
                        End If
                        p.X = 10
                        .Checked = True
                        'Case 1
                        '   .Text = "Visible"
                        '  p.X = 90
                       ' .Checked = False
                    Case 1
                        If Me.Langue = "fr" Then
                            .Text = "Tout"
                        Else
                            .Text = "All"
                        End If
                        p.X = 90
                        .Checked = False
                End Select
                .Location = p
                .BringToFront()

            End With
        Next i

        ' MIDI Learn
        ' **********
        Panneau3.Panel1.Controls.Add(MidiLearn)
        With MidiLearn
            p.X = 340
            .Size = New Size(20, 20)
            If Me.Langue = "fr" Then
                .Text = "Midi Learn (avec +/-)"
            Else
                .Text = "Midi Learn (width +/-)"
            End If
            .Font = fnt7
            .Checked = False
            .AutoSize = True
            .Location = p 'New Point(300, 5)
            .Visible = True
            .Enabled = True
            .BringToFront()
        End With
        Panneau3.Panel1.Controls.Add(AffMidiLearn)
        With AffMidiLearn
            p.X = 490
            .Size = New Size(20, 20)
            .Location = p
            .Font = fnt7
            .Visible = True
            .Enabled = True
            .AutoSize = True
            .Text = "100"
            .BringToFront()
        End With

        ' Bouton talle fenêtre
        ' ********************
        'Panneau3.Panel1.Controls.Add(BoutonTaille)
        'With BoutonTaille
        ' p.X = 520
        ' p.Y = 10
        ' .Font = fnt6
        ' .Size = New Size(20, 30)
        ' .AutoSize = True
        ' .Location = p
        '.Visible = True
        '.Enabled = True
        '.Text = "DIM"
        '.BringToFront()
        '.Visible = False
        'End With

        ' Aller vers
        ' **********
        Panneau3.Panel1.Controls.Add(LabMesures)
        Panneau3.Panel1.Controls.Add(LabMarqueurs)
        Panneau3.Panel1.Controls.Add(TextMesures)
        Panneau3.Panel1.Controls.Add(ListMarqueurs)
        '
        LabMesures.AutoSize = True
        If Me.Langue = "fr" Then
            LabMesures.Text = "Aller vers Mesure"
        Else
            LabMesures.Text = "Goto Measure"
        End If

        LabMarqueurs.AutoSize = True
        If Me.Langue = "fr" Then
            LabMarqueurs.Text = "Aller vers Marqueur"
        Else
            LabMarqueurs.Text = "Goto Marker"
        End If
        '
        p.X = 600
        p.Y = 2
        LabMesures.Location = p
        LabMesures.Font = fnt6
        p.X = 750
        p.Y = 2
        LabMarqueurs.Location = p
        LabMarqueurs.Font = fnt6
        '
        s.Width = 120
        s.Height = 300
        '
        TextMesures.Size = s
        '
        p.X = 600
        p.Y = 20
        TextMesures.Location = p
        TextMesures.TextAlign = HorizontalAlignment.Center
        TextMesures.Font = fnt7
        '
        p.X = 750
        p.Y = 20
        s.Width = 200
        ListMarqueurs.Location = p
        ListMarqueurs.Items.Add("No Markers")
        ListMarqueurs.Font = fnt7
        ListMarqueurs.Size = s
        '

        AddHandler TextMesures.KeyUp, AddressOf TextMesures_KeyUp
        AddHandler ListMarqueurs.SelectedIndexChanged, AddressOf ListMarqueurs_SelectedIndexChanged
        'AddHandler BoutonTaille.MouseClick, AddressOf BoutonTaille_MouseClick
    End Sub
    Sub TextMesures_KeyUp(Sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If IsNumeric(TextMesures.Text) Then
            Dim N As Integer = Convert.ToInt16(TextMesures.Text)
            If e.KeyCode = Keys.Enter And (N >= 1 And N <= Module1.nbMesures) Then
                Dim i As Integer = ((N - 1) * Me.DivisionMes) + 1
                AllerVersMesure(i)
            End If
        End If
    End Sub
    Sub AllerVersMesure(N_Mesure As Integer)
        Dim i, j As Integer

        For i = 0 To nb_PistesVar - 1
            For j = 0 To nbCourbes - 1
                RemoveHandler Canaux.Item(i).GridCourbes.Item(j).Scroll, AddressOf GridCourbes_Scroll
            Next j
        Next i
        GridEntête.LeftCol = N_Mesure
        For i = 0 To nb_PistesVar - 1
            For j = 0 To nbCourbes - 1
                Canaux.Item(i).GridCourbes.Item(j).AutoRedraw = False
                Canaux.Item(i).GridCourbes.Item(j).LeftCol = N_Mesure
                Canaux.Item(i).GridCourbes.Item(j).AutoRedraw = True
                Canaux.Item(i).GridCourbes.Item(j).Refresh()
            Next j
        Next i
        For i = 0 To nb_PistesVar - 1
            For j = 0 To nbCourbes - 1
                AddHandler Canaux.Item(i).GridCourbes.Item(j).Scroll, AddressOf GridCourbes_Scroll
            Next j
        Next i

    End Sub
    Sub ListMarqueurs_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim tbl() As String
        If EnChargement = False Then
            If ListMarqueurs.Items.Count > 1 Then
                tbl = ListMarqRefresh.Item(ListMarqueurs.SelectedIndex).Split(";")
                If IsNumeric(tbl(1)) Then
                    Dim N As Integer = Convert.ToInt16(tbl(1))
                    Dim i As Integer = ((N - 1) * Me.DivisionMes) + 1
                    AllerVersMesure(i)
                End If
            End If
        End If
    End Sub
    Sub BoutonTaille_MouseClick(sender As Object, e As EventArgs)
        If BoutonTaillePréced = FenêtreHeight Then
            BoutonTaillePréced = 380
            F4.Height = 380
            F4.Width = FenêtreWidth
        Else
            BoutonTaillePréced = FenêtreHeight
            F4.Height = FenêtreHeight
            F4.Width = FenêtreWidth

        End If
    End Sub


    Function Det_ModeScroll() As Integer
        Dim i As Integer
        For i = 0 To 2
            If LSyncScroll.Item(i).Checked Then Exit For
        Next
        Return i
    End Function


    Sub Construction_GridEntête()
        Dim i, j, k As Integer
        Dim p As New Point
        Dim s As New Size
        Dim Div As Integer = Det_DivisionMes()
        Panneau1.Panel1.Controls.Add(GridEntête)
        '
        GridEntête.FixedRows = 3
        GridEntête.FixedCols = 1
        GridEntête.Cols = 1 + (Module1.nbMesures * 16)
        GridEntête.Rows = 3
        '
        GridEntête.BackColorFixed = Color.Ivory
        For i = 1 To GridEntête.Cols - 1
            GridEntête.Column(i).Width = TailleCel
        Next
        GridEntête.Column(0).Width = TailleColFix - 3
        GridEntête.Width = TailleCanaux - 8 'TailleColFix + (TailleCel * (GridEntête.Cols - 1))
        GridEntête.Height = (GridEntête.Rows * TailleCel) + 7
        GridEntête.ScrollBars = ScrollBarsEnum.None
        GridEntête.Font = fnt6
        '
        For i = 0 To 1
            j = 1
            k = 1
            Do
                GridEntête.Range(i, j, i, j + 15).Merge()
                GridEntête.Cell(2, j).Text = Convert.ToString(k)
                j += Div
                k += 1
            Loop Until j > GridEntête.Cols - 1
        Next i
        '
        p.X = 3
        p.Y = 0
        GridEntête.Location = p
        '
        GridEntête.Enabled = False
    End Sub
    Sub Construction_Courbes()
        Dim i1, i, j As Integer
        Dim P As Point
        Dim s As Size
        Try
            For i1 = 0 To Module1.nb_PistesVar - 1

                Canaux.Add(New Canal)

                For i = 0 To nbCourbes - 1
                    TabCourbes.Add(New TabControl)
                    TabCourbes.Item(i1).Tag = i1
                    TabCourbes.Item(i1).Appearance = TabAppearance.Buttons
                    Panneau1.Panel2.Controls.Add(TabCourbes.Item(i1))
                    '
                    With Canaux.Item(i1)

                        ' autorisation actions
                        ' ********************


                        .GridCourbes.Add(New FlexCell.Grid) ' courbes
                        .PageCourbes.Add(New TabPage)       ' onglet

                        .GridCourbes.Item(i).AutoRedraw = False ' <-- laisser à cet endroit, i n'existe pas tant que le add n'a pas eu lieu
                        '
                        .GridCourbes.Item(i).AllowDrop = False
                        '.GridCourbes.Item(i).AllowUserPaste = ClipboardDataEnum.None
                        .GridCourbes.Item(i).AllowUserReorderColumn = False
                        .GridCourbes.Item(i).AllowUserResizing = False
                        .GridCourbes.Item(i).AllowUserSort = False

                        .GridCourbes.Item(i).AllowDrop = False
                        '.GridCourbes.Item(i).AllowUserPaste = ClipboardDataEnum.None
                        .GridCourbes.Item(i).AllowUserReorderColumn = False
                        .GridCourbes.Item(i).AllowUserResizing = False
                        .GridCourbes.Item(i).AllowUserSort = False

                        TabCourbes.Item(i1).Controls.Add(.PageCourbes.Item(i))  ' placement de l'onglet dans son Tabcontrol
                        .PageCourbes.Item(i).Controls.Add(.GridCourbes.Item(i)) ' placement de la grille dans l'onglet
                        '
                        .PageCourbes.Item(i).Tag = i1.ToString + " " + i.ToString
                        .GridCourbes.Item(i).Tag = i1.ToString + " " + i.ToString
                        '
                        .PageCourbes.Item(i).Dock = DockStyle.Fill
                        .GridCourbes.Item(i).Dock = DockStyle.Fill

                        .GridCourbes.Item(i).Font = fnt6
                        '
                        .GridCourbes.Item(i).Cols = 1 + (Module1.nbMesures * 16)
                        .GridCourbes.Item(i).Rows = 66
                        '
                        .GridCourbes.Item(i).FixedRows = 1
                        .GridCourbes.Item(i).FixedCols = 1
                        '
                        .GridCourbes.Item(i).SelectionMode = FlexCell.SelectionModeEnum.ByColumn
                        .GridCourbes.Item(i).SelectionBorderColor = Color.White
                        '

                        .GridCourbes.Item(i).ScrollBars = ScrollBarsEnum.Horizontal
                        '
                        .GridCourbes.Item(i).BorderStyle = BorderStyleEnum.FixedSingle
                        .GridCourbes.Item(i).CellBorderColor = Color.Black
                        '
                        .GridCourbes.Item(i).BackColorSel = Color.Transparent
                        .GridCourbes.Item(i).Cursor = Cursors.Arrow
                        '
                        For j = 0 To (.GridCourbes.Item(i).Cols) - 1
                            .GridCourbes.Item(i).Column(j).Width = TailleCel
                        Next
                        .GridCourbes.Item(i).Column(0).Width = TailleColFix - 3
                        '
                        For j = 0 To (.GridCourbes.Item(i).Rows) - 1
                            .GridCourbes.Item(i).Row(j).Height = HautLigneCourbes
                        Next
                        .GridCourbes.Item(i).Row(.GridCourbes.Item(i).Rows - 1).Height = 7
                        .GridCourbes.Item(i).Row(0).Height = TailleCel
                        '
                        ' positionnement et dimensionnement de tabcourbes
                        ' **********************************************
                        s.Width = TailleCanaux '
                        s.Height = (.GridCourbes.Item(i).Rows * HautLigneCourbes) + 69
                        TabCourbes.Item(i1).Size = s

                        '
                        P.Y = (i1 * TabCourbes.Item(i1).Size.Height) ' 
                        P.X = 0
                        TabCourbes.Item(i1).Location = P
                        '
                        ' Mise à jour des onglets
                        ' ***********************
                        Select Case i
                            Case 0
                                .PageCourbes.Item(i).Text = "Expression"
                                .GridCourbes.Item(i).BackColorFixed = CoulExp
                                .GridCourbes.Item(i).BackColorFixedSel = CoulExp

                            Case 1
                                .PageCourbes.Item(i).Text = "Modulation"
                                .GridCourbes.Item(i).BackColorFixed = CoulMod
                                .GridCourbes.Item(i).BackColorFixedSel = CoulMod
                            Case 2
                                .PageCourbes.Item(i).Text = "PAN"
                                .GridCourbes.Item(i).BackColorFixed = CoulPan
                                .GridCourbes.Item(i).BackColorFixedSel = CoulPan
                            '
                            Case 3
                                .PageCourbes.Item(i).Text = "CC50"
                                .GridCourbes.Item(i).BackColorFixed = CoulCC50
                                .GridCourbes.Item(i).BackColorFixedSel = CoulCC50
                            Case 4
                                .PageCourbes.Item(i).Text = "CC51"
                                .GridCourbes.Item(i).BackColorFixed = CoulCC51
                                .GridCourbes.Item(i).BackColorFixedSel = CoulCC51
                            Case 5
                                .PageCourbes.Item(i).Text = "CC52"
                                .GridCourbes.Item(i).BackColorFixed = CoulCC52
                                .GridCourbes.Item(i).BackColorFixedSel = CoulCC52
                            Case 6
                                .PageCourbes.Item(i).Text = "CC53"
                                .GridCourbes.Item(i).BackColorFixed = CoulCC53
                                .GridCourbes.Item(i).BackColorFixedSel = CoulCC53
                        End Select
                        '
                        ' Numérotation des mesures
                        ' ************************
                        Dim ii, jj, kk As Integer
                        For ii = 0 To 1
                            jj = 1
                            kk = 1
                            Do
                                'GridEntête.Cell(0, jj).Text = Convert.ToString(kk)
                                .GridCourbes.Item(i).Cell(0, jj).Text = Convert.ToString(kk)
                                jj += 16
                                kk += 1
                            Loop Until jj > .GridCourbes.Item(i).Cols - 1
                        Next ii
                        '
                        ' Colorisation des beats
                        ' **********************
                        For ii = 0 To 1
                            jj = 1
                            kk = 1
                            Do
                                If Trim(.GridCourbes.Item(i).Cell(0, jj).Text) <> "" Then
                                    .GridCourbes.Item(i).Cell(0, jj).Font = fnt9
                                    .GridCourbes.Item(i).Cell(0, jj).ForeColor = Color.Black
                                    .GridCourbes.Item(i).Cell(0, jj).BackColor = Color.White
                                Else
                                    .GridCourbes.Item(i).Cell(0, jj).ForeColor = Color.Gold
                                    .GridCourbes.Item(i).Cell(0, jj).BackColor = Color.DarkGreen
                                End If
                                'End If
                                jj += 4
                                kk += 1
                            Loop Until jj > .GridCourbes.Item(i).Cols - 1
                        Next ii

                        AddHandler .GridCourbes.Item(i).KeyUp, AddressOf GridCourbes_KeyUp
                        AddHandler .GridCourbes.Item(i).KeyDown, AddressOf GridCourbes_KeyDown
                        AddHandler .GridCourbes.Item(i).SelChange, AddressOf GridCourbes_Selchange
                        AddHandler .GridCourbes.Item(i).Scroll, AddressOf GridCourbes_Scroll
                        AddHandler .GridCourbes.Item(i).MouseDown, AddressOf GridCourbes_MouseDown
                        'AddHandler .GridCourbes.Item(i).ButtonClick, AddressOf GridCourbes_ButtonClick

                        .GridCourbes.Item(i).AutoRedraw = True
                        .GridCourbes.Item(i).Refresh()
                        '
                    End With
                Next i
            Next i1
        Catch
            MsgBox("Erreur dans procédure PianoRoll/Construction_Courbes - index i =" + Str(i))
        End Try
    End Sub
    Public Sub Construction_Menu()
        '
        Fichier.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Quitter})
        Fichier.Size = New System.Drawing.Size(87, 20)
        If Me.Langue = "fr" Then
            Fichier.Text = "Fichier"
        Else
            Fichier.Text = "File"
        End If
        Fichier.Visible = True
        'Fichier.BackColor = Me.CoulBarOut
        ' Edition
        Edition.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Couper, Copier, Coller, Séparateur1, Annuler, Séparateur2, Supprimer})
        Edition.Size = New System.Drawing.Size(56, 20)
        Edition.Text = "Edition"
        Edition.Visible = True
        'Edition.BackColor = Me.CoulBarOut
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
        Annuler.Enabled = False

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
        F4.Controls.Add(Menu1)
        F4.MainMenuStrip = Menu1
        F4.MainMenuStrip.Visible = False
        '
        AddHandler Couper.Click, AddressOf Couper_Click
        AddHandler Copier.Click, AddressOf Copier_Click
        AddHandler Coller.Click, AddressOf Coller_Click
        'AddHandler Annuler.Click, AddressOf Annuler_Click
        'AddHandler Supprimer.Click, AddressOf Supprimer_Click
        AddHandler Quitter.Click, AddressOf Quitter_Click

    End Sub
    Sub Couper_Click(Sender As Object, e As EventArgs) '
        Canaux(Me.NumCanal).GridCourbes(Me.NumCourbe).AutoRedraw = False
        Canaux(Me.NumCanal).GridCourbes(Me.NumCourbe).Selection.CutData()
        Canaux(Me.NumCanal).GridCourbes(Me.NumCourbe).AutoRedraw = True
        Canaux(Me.NumCanal).GridCourbes(Me.NumCourbe).Refresh()
    End Sub
    Sub Copier_Click(Sender As Object, e As EventArgs)
        Canaux(Me.NumCanal).GridCourbes(Me.NumCourbe).AutoRedraw = False
        Canaux(Me.NumCanal).GridCourbes(Me.NumCourbe).Selection.CopyData()
        Canaux(Me.NumCanal).GridCourbes(Me.NumCourbe).AutoRedraw = True
        Canaux(Me.NumCanal).GridCourbes(Me.NumCourbe).Refresh()
    End Sub
    Sub Coller_Click(Sender As Object, e As EventArgs)
        Orig_Autom_Vérrouillage = True ' Impossibilité de faire un tracé en cas de collage
        Canaux(Me.NumCanal).GridCourbes(Me.NumCourbe).AutoRedraw = False
        Canaux(Me.NumCanal).GridCourbes(Me.NumCourbe).Selection.PasteData()
        Canaux(Me.NumCanal).GridCourbes(Me.NumCourbe).AutoRedraw = True
        Canaux(Me.NumCanal).GridCourbes(Me.NumCourbe).Refresh()
        Orig_Autom_Vérrouillage = False ' retour à la possibilité de Tracer
    End Sub
    Sub Quitter_Click(Sender As Object, e As EventArgs)
        Dim NumOnglet As Integer = Convert.ToUInt16(F4.Tag)
        Attacher(NumOnglet)
    End Sub
    Sub GridCourbes_MouseDown(Sender As Object, e As EventArgs)
        Dim com As FlexCell.Grid = Sender
        Dim a As String = com.Tag
        Dim tbl() As String = a.Split()
        Dim i1 As Integer = Convert.ToInt16(tbl(0)) ' n° du tabcontrol
        Dim i2 As Integer = Convert.ToInt16(tbl(1))  ' n° de la courbe dans le tabcontrol


        ' mise à jour des proprétés pour cas couper, copier, coller dans form1 quand le formulaire est détaché
        Me.NumCanal = i1
        Me.NumCourbe = i2

        ' pour controle isolé d'une seule colonne avec ctrl+clic
        Tracé_Courbe(Sender)

    End Sub
    Sub GridCourbes_Scroll(Sender As Object, e As EventArgs)
        Dim com As FlexCell.Grid = Sender
        Dim a As String = com.Tag
        Dim tbl() As String = a.Split()
        Dim i1 As Integer = Convert.ToInt16(tbl(0)) ' n° du tabcontrol
        Dim i As Integer = Convert.ToInt16(tbl(1))  ' n° de la courbe dans le tabcontrol
        Dim ind, j As Integer
        Dim valsync As Integer = Canaux.Item(i1).GridCourbes.Item(i).LeftCol
        Dim Mode As Integer = Det_ModeScroll()


        For i = 0 To nb_PistesVar - 1
            For j = 0 To nbCourbes - 1
                'Canaux.Item(i).GridCourbes.Item(j).AutoRedraw = False
                RemoveHandler Canaux.Item(i).GridCourbes.Item(j).Scroll, AddressOf GridCourbes_Scroll

            Next j
        Next i
        '
        Select Case Mode
            Case 0 ' Courant
                GridEntête.LeftCol = valsync
            Case 1 ' Visible
                GridEntête.LeftCol = valsync
                For i = 0 To nb_PistesVar - 1
                    ind = TabCourbes.Item(i).SelectedIndex
                    'Canaux.Item(i).GridCourbes.Item(ind).AutoRedraw = False
                    Canaux.Item(i).GridCourbes.Item(ind).LeftCol = valsync
                    'Canaux.Item(i).GridCourbes.Item(ind).AutoRedraw = True
                    'Canaux.Item(i).GridCourbes.Item(ind).Refresh()
                Next i
            Case 2 'Tout
                GridEntête.LeftCol = valsync

                For i = 0 To nb_PistesVar - 1
                    For j = 0 To nbCourbes - 1
                        'Canaux.Item(i).GridCourbes.Item(j).AutoRedraw = False
                        Canaux.Item(i).GridCourbes.Item(j).LeftCol = valsync
                        'Canaux.Item(i).GridCourbes.Item(j).AutoRedraw = True
                        'Canaux.Item(i).GridCourbes.Item(j).Refresh()
                        ')
                    Next j
                Next i
        End Select

        '
        For i = 0 To nb_PistesVar - 1
            For j = 0 To nbCourbes - 1
                AddHandler Canaux.Item(i).GridCourbes.Item(j).Scroll, AddressOf GridCourbes_Scroll ' correspond au remove handler du début
                'Canaux.Item(i).GridCourbes.Item(j).AutoRedraw = True
                'Canaux.Item(i).GridCourbes.Item(j).Refresh()
            Next j
        Next i
    End Sub
    Sub GridCourbes_KeyUp(Sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Dim com As FlexCell.Grid = Sender
        Dim a As String = com.Tag
        Dim tbl() As String = a.Split()
        Dim i1 As Integer = Convert.ToInt16(tbl(0)) ' N° du Tabcontrol
        Dim i2 As Integer = Convert.ToInt16(tbl(1)) ' N° de la courbe (Grid)
        ' effacer totalement une colonne
        If e.KeyCode = Keys.Delete Or e.KeyCode = Keys.Back Then
            Dim i As Integer = Canaux(i1).GridCourbes.Item(i2).Selection.FirstRow
            Dim j As Integer = Canaux(i1).GridCourbes.Item(i2).Selection.FirstCol
            Dim ii As Integer = Canaux(i1).GridCourbes.Item(i2).Selection.LastRow
            Dim jj As Integer = Canaux(i1).GridCourbes.Item(i2).Selection.LastCol
            Canaux(i1).GridCourbes.Item(i2).AutoRedraw = False
            Canaux(i1).GridCourbes.Item(i2).Range(i, j, ii, jj).BackColor = Color.White
            Canaux(i1).GridCourbes.Item(i2).AutoRedraw = True
            Canaux(i1).GridCourbes.Item(i2).Refresh()
        End If
        '


    End Sub
    Public Sub GridCourbes_KeyDown(Sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Dim com As FlexCell.Grid = Sender
        Dim a As String = com.Tag
        Dim tbl() As String
        Dim i1 As Integer ' N° du canal
        Dim i2 As Integer ' N° de la courbe dans le canal
        tbl = a.Split

        i1 = Convert.ToInt16(tbl(0))
        i2 = Convert.ToInt16(tbl(1))

        Dim i As Integer = Canaux.Item(i1).GridCourbes.Item(i2).Selection.FirstRow ' on n'agit ici que sur une colonne à la fois.
        Dim j As Integer = Canaux.Item(i1).GridCourbes.Item(i2).Selection.FirstCol ' j est constant dans cette méthode
        Dim k As Integer
        Dim vv As String
        Dim v As Byte ' 
        Dim b As String

        With Canaux.Item(i1)
            If e.KeyCode <> Keys.ControlKey Then
                If e.KeyCode = Keys.Add Or e.KeyCode = Keys.Subtract Then
                    b = ValCtrl2_Ligne(i1, i2, j) ' 
                    If Trim(b) <> "" Then
                        k = Convert.ToInt16(Trim(b)) ' lecture de la ligne actuelle différente de "blanc" = valeur actuelle du CC pour la colonne j
                        '
                        ' Traitement incrémentation/décrémentation (si sortir = false)
                        ' ************************************************************
                        '
                        ' INCREMENTATION
                        ' **************
                        If e.KeyCode = Keys.Add Then
                            k -= 1
                            .GridCourbes.Item(i2).AutoRedraw = False
                            .GridCourbes.Item(i2).Range(k, j, .GridCourbes.Item(i2).Rows - 1, j).BackColor = Det_CouleurCTRL(i2)
                            '
                            vv = ValCtrl(i1, i2, j)
                            If Trim(vv) <> "-1" Then
                                v = Convert.ToByte(vv)
                                AffMidiLearn.Text = Trim(vv) 'v.ToString
                                If MidiLearn.Checked Then
                                    Dim c As Byte = Det_CTRL(i2)
                                    Send_CTRL(c, v, i1) ' Canal est donné par la constructeur de la classe
                                End If
                            Else
                                AffMidiLearn.Text = "--"
                            End If
                            '
                            .GridCourbes.Item(i2).AutoRedraw = True
                            .GridCourbes.Item(i2).Refresh()

                            '
                            ' DECREMENTATION
                            ' **************
                        ElseIf e.KeyCode = Keys.Subtract Then
                            k += 1
                            If k <= .GridCourbes.Item(i2).Rows - 1 Then '   
                                .GridCourbes.Item(i2).AutoRedraw = False
                                .GridCourbes.Item(i2).Range(1, j, .GridCourbes.Item(i2).Rows - 1, j).BackColor = Color.White
                                If k <> .GridCourbes.Item(i2).Rows - 1 Then
                                    .GridCourbes.Item(i2).Range(k, j, .GridCourbes.Item(i2).Rows - 1, j).BackColor = Det_CouleurCTRL(i2)
                                End If

                                vv = ValCtrl(i1, i2, j)
                                If Trim(vv) <> "-1" Then
                                    v = Convert.ToByte(vv)
                                    AffMidiLearn.Text = Trim(vv) ' v.ToString
                                    If MidiLearn.Checked Then
                                        Dim c As Byte = Det_CTRL(i2)
                                        Send_CTRL(c, v, i1) ' Canal est donné par la constructeur de la classe
                                    End If
                                Else
                                    AffMidiLearn.Text = "--"
                                End If
                                '
                                .GridCourbes.Item(i2).AutoRedraw = True
                                .GridCourbes.Item(i2).Refresh()

                            End If
                        End If
                    End If
                End If
            End If
        End With
    End Sub

    ''' <summary>
    ''' ValCtrl2_Ligne : retourne la 1ere ligne différent de Blanc dans la colonne passée en paramètre. Cela donne la valeur do CC pour la conne indiquée.)
    ''' </summary>
    ''' <param name="i1">N° du Canal</param>
    ''' <param name="i2">N° de la courbe dans le canal</param>
    ''' <param name="colonne">Colonne de la courbe à traiter</param>
    ''' <returns></returns>
    Private Function ValCtrl2_Ligne(i1 As Integer, i2 As Integer, colonne As Integer) As String
        Dim Ligne As Integer
        Dim j As Integer = colonne
        Dim a As String = (Canaux(i1).GridCourbes.Item(i2).Rows - 1).ToString
        '
        With Canaux(i1)
            For Ligne = 1 To .GridCourbes.Item(i2).Rows - 1
                If .GridCourbes.Item(i2).Cell(Ligne, j).BackColor <> Color.White Then ' on recherche ici la 1erre cellule dont la couleur n'est blanche
                    a = Ligne.ToString ' ligne est le N° de ligne
                    Exit For
                End If
            Next
        End With
        Return a
    End Function
    ''' <summary>
    ''' ValCtrl : retourne la valeur du CTRL pour le canal i1,  pour la courbe i2 et pour la colonne 'colonne'. (valeur= 0 à 127)
    ''' </summary>
    ''' <param name="i1"></param> N° du canal
    ''' <param name="i2"></param> N° du canal
    ''' <param name="colonne"></param> Colonne de la grille des courbes : colonne = postion de l'évènement en double croches à partir du début
    ''' <returns>Valeur de CC dans la colonne 'colonne'                                                                                                                                                                                                                                                                                                                                                                                      </returns>

    Public Function ValCtrl(i1 As Integer, i2 As Integer, colonne As Integer) As String
        Dim ligne As Integer
        Dim j As Integer = colonne
        Dim k As Integer = 0
        Dim m = Canaux(i1).GridCourbes.Item(i2).Rows ' pour debug
        With Canaux(i1)
            For ligne = 1 To .GridCourbes.Item(i2).Rows - 1
                If ligne = .GridCourbes.Item(i2).Rows - 1 Then ' rows = 66
                    k = 0
                    k = CTRL_ValAmont(i1, i2)
                    Exit For
                    Exit For
                Else
                    If .GridCourbes.Item(i2).Cell(ligne, j).BackColor <> Color.White Then
                        k = ((.GridCourbes.Item(i2).Rows - 1 - ligne) * 2) - 1 ' calcul de la valeur de CC pour la colonne j ' -1
                        Exit For
                    End If
                End If
            Next
        End With
        k = k - 1
        Return k.ToString
    End Function
    Function CTRL_ValAmont(N_Courbe As Integer, colonne As Integer) As Integer
        Dim j As Integer
        Dim ligne As Integer
        Dim k As Integer = 0
        Dim flag As Boolean = False

        With Canaux(N_Courbe)

            For j = colonne To 1 Step -1
                For ligne = 1 To .GridCourbes.Item(N_Courbe).Rows - 1
                    If .GridCourbes.Item(N_Courbe).Cell(ligne, j).BackColor <> Color.White Then ' valeur trouvée
                        k = ((.GridCourbes.Item(N_Courbe).Rows - 1 - ligne) * 2) - 1 ' 1ere valeur trouvée dans un colonne précédent du CTRL
                        flag = True
                        Exit For
                    End If
                Next
                If flag Then Exit For
            Next
        End With
        Return k
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
    Private Sub GridCourbes_Selchange(Sender As Object, e As Grid.SelChangeEventArgs)
        Tracé_Courbe(Sender)
    End Sub
    Sub Tracé_Courbe(Sender As Object)
        Dim com As FlexCell.Grid = Sender
        Dim a As String = com.Tag
        Dim tbl() As String
        tbl = a.Split
        Dim i1 As Integer = tbl(0)
        Dim i2 As Integer = tbl(1)
        Dim i As Integer = Canaux(i1).GridCourbes.Item(i2).MouseRow
        Dim j As Integer = Canaux(i1).GridCourbes.Item(i2).MouseCol
        Dim ligne As Integer = 0

        With Canaux(i1)
            .GridCourbes.Item(i2).AutoRedraw = False
            '
            If i <> -1 And j <> -1 Then
                .GridCourbes.Item(i2).AutoRedraw = False
                ListCCAct.Item(i1).LCCAct.Item(i2).Checked = True ' dès que l'on trace dans une courbe, la case à cocher correspondante, dans la barre d'outils, est cochée.
                If My.Computer.Keyboard.CtrlKeyDown And Orig_Autom_Vérrouillage = False Then ' vérouillage est mis à jour dans l'évènement CTRL + V : il empêche d'écrire dans la courbe quand on fait un "coller" avec CTRL+V
                    .GridCourbes.Item(i2).AutoRedraw = False
                    .GridCourbes.Item(i2).Range(1, j, .GridCourbes.Item(i2).Rows - 1, j).ClearBackColor()

                    If i <> .GridCourbes.Item(i2).Rows - 1 Then ' ce test permet d'obtenir la RAZ de la dernière cellule du bas par Ctrl + clic
                        .GridCourbes.Item(i2).Range(.GridCourbes.Item(i2).Rows - 1, j, i, j).BackColor = Det_CouleurCTRL(i2)
                    End If
                    .GridCourbes.Item(i2).AutoRedraw = True
                    .GridCourbes.Item(i2).Refresh()
                End If
                .GridCourbes.Item(i2).AutoRedraw = True
                .GridCourbes.Item(i2).Refresh()
            End If
            '
        End With
    End Sub
    Private Sub Dockbutton_MouseClick(sender As Object, e As EventArgs)
        Dim NumOnglet As Integer = Convert.ToUInt16(F4.Tag)

        ' Raz du cli^board de la partie HyperARp
        Form1.ClipEdit.Clear()

        If Me.F4.Dock = DockStyle.Fill Then ' détacher
            Panneau1.Panel2.Enabled = True
            Me.F4.Visible = False
            Me.F4.FormBorderStyle = FormBorderStyle.Sizable

            ' 
            Me.F4.Dock = DockStyle.None
            Form1.TabControl2.TabPages.Item(NumOnglet).Controls.Remove(Me.F4)
            Dim p As New Point With {
            .X = Me.F4.Location.X + 50,
            .Y = Me.F4.Location.Y + 100
        }
            Me.F4.Location = p
            Me.F4.StartPosition = FormStartPosition.Manual ' permet de tenir compte de la location calculée dans p
            Me.F4.TopLevel = True '
            Me.F4.MainMenuStrip.Visible = True
            Me.F4.MainMenuStrip.Enabled = True
            '
            Dim s As New Size With {
            .Width = FenêtreWidth, '+ 500, 'Panneau3.Width,
            .Height = FenêtreHeight 'Panneau3.Height + 400
        }

            Me.F4.FormBorderStyle = FormBorderStyle.Sizable
            Me.F4.Size = s
            '
            If LangueIHM = "fr" Then
                DockButton.Text = "Attacher"
            Else
                DockButton.Text = "Dock"
            End If
            '
            Me.F4.Visible = True
            Refresh_Général()
            Maj_Tooltips()

        Else ' attacher
            Panneau1.Panel2.Enabled = False
            Me.F4.Visible = False
            F4.FormBorderStyle = FormBorderStyle.None ' absolument nécessaire pour que l'attachement se passe bien et autre pour que le titre du formulaire disparaisse.
            Me.F4.TopMost = False   ' un seul des 2 suffit ?
            Me.F4.TopLevel = False
            Form1.TabControl2.TabPages.Item(NumOnglet).Controls.Add(Me.F4)
            Me.F4.Dock = DockStyle.Fill

            Me.F4.MainMenuStrip.Visible = False
            '
            If Langue = "fr" Then
                DockButton.Text = "Détacher"
            Else
                DockButton.Text = "UnDock"
            End If
            Me.F4.Visible = True
            Refresh_Général()
        End If
    End Sub

    Sub Refresh_Général()
        Dim i1 As Integer
        Dim i2 As Integer
        For i1 = 0 To nb_PistesVar - 1
            For i2 = 0 To nbCourbes - 1
                Canaux.Item(i1).GridCourbes.Item(i2).Refresh()
            Next
        Next
    End Sub

    Public Sub F4_Refresh()
        Dim i, j As Integer
        Dim tbl1() As String
        Dim tbl2() As String
        Dim tbl3() As String
        Dim Div As Integer = Me.DivisionMes ' nombre de doubles croches par mesures
        '
        GridEntête.AutoRedraw = False
        '
        ' effacer le text de la ligne entière des marqueurs et des accords
        GridEntête.Range(0, 1, 0, GridEntête.Cols - 1).ClearText() ' effacer le texte de la ligne des marqueurs
        GridEntête.Range(1, 1, 1, GridEntête.Cols - 1).ClearText() ' effacer le texte de la ligne des accords
        ' mise à jour des accords et des marqueurs
        tbl1 = Split(Me.ListMarq, "&")
        tbl2 = Split(Me.ListAcc, "-")
        '
        GridEntête.Range(0, 1, 0, GridEntête.Cols - 1).BackColor = Color.Ivory
        GridEntête.Range(0, 1, 0, GridEntête.Cols - 1).ForeColor = Color.Black
        If Trim(Me.ListMarq) <> "" Then
            For i = 0 To tbl1.Count - 1
                tbl3 = tbl1(i).Split(";")
                j = ((Convert.ToInt16(tbl3(1)) - 1) * Div) + 1
                GridEntête.Cell(0, j).Text = tbl3(0)
                GridEntête.Cell(0, j).Alignment = AlignmentEnum.CenterCenter
                GridEntête.Cell(0, j).Font = fnt6 'verdana 8
                GridEntête.Cell(0, j).BackColor = Color.DarkOliveGreen
                GridEntête.Cell(0, j).ForeColor = Color.Yellow
            Next i
        End If
        '
        ListMarqueurs.Items.Clear()
        If Trim(Me.ListMarq) <> "" Then
            ListMarqRefresh.Clear()

            For i = 0 To tbl1.Count - 1
                ListMarqRefresh.Add(tbl1(i))
                tbl2 = tbl1(i).Split(";")
                ListMarqueurs.Items.Add(tbl2(0))
            Next
            ListMarqueurs.SelectedIndex = 0
        Else
            ListMarqueurs.Items.Add("No Markers")
            ListMarqueurs.SelectedIndex = 0
        End If

        If Trim(Me.ListAcc) <> "" Then
            j = 1
            For i = 0 To tbl2.Count - 1
                GridEntête.Cell(1, j).Text = tbl2(i)
                GridEntête.Cell(1, j).Alignment = AlignmentEnum.CenterCenter
                GridEntête.Cell(1, j).Font = fnt6 'verdana 8
                j += Div
            Next
        End If
        '
        GridEntête.AutoRedraw = True
        GridEntête.Refresh()

    End Sub
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
    Public Function Enregistrer_Ctrls(N_Canal As Integer, N_Courbe As Integer) As String

        Dim a As String = ""
        Dim colonne As Integer
        Dim ligne As String

        With Canaux.Item(N_Canal)
            a = "Automation," + (N_Canal + 1).ToString + ",Control," + N_Courbe.ToString + ","
            For colonne = 1 To .GridCourbes(N_Courbe).Cols - 1
                ligne = ValCtrl2_Ligne(N_Canal, N_Courbe, colonne)
                If ligne <> (.GridCourbes.Item(N_Courbe).Rows - 1).ToString Then ' la dernière ligne indique la valeur 0, elle reste en blanc
                    a = a + ligne + " " + colonne.ToString + ","
                End If
            Next colonne
        End With

        a = Trim(Microsoft.VisualBasic.Left(a, a.Length - 1))
        Return a
    End Function
    ''' <summary>
    ''' Enregistrer_ControlSys : sauvegarde des paramètres de la barre local des Automations.
    ''' Ici, on ne sauvegarde le MIDI Learn (ce n'est pas la peine)
    ''' </summary>
    ''' <param name="N_Canal">N° du Canal d'automation 1..6</param>
    ''' <returns></returns>
    Public Function Enregistrer_ControlSys(N_Canal As Integer) As String
        Dim i As Integer
        Dim a As String = ""

        With Canaux.Item(N_Canal)
            a = "Automation," + Convert.ToString((N_Canal + 1).ToString) + ",ControlSys,"
            a = a + (.Mute.Checked).ToString + ","
            For i = 0 To nbCourbes - 1
                a = a + (ListCCAct.Item(N_Canal).LCCAct(i).Checked).ToString + ","
            Next i
            'a = a + (.MidiLearn.Checked).ToString 
            a = Trim(Microsoft.VisualBasic.Left(a, a.Length - 1))
            Return a
        End With
    End Function
    Public Sub Charger_Ctrl(LesCourbes As String)
        Dim j As Integer
        Dim tbl1() As String
        Dim tbl2() As String
        Dim Ligne, Colonne As Integer
        tbl1 = LesCourbes.Split(",")
        Dim i1 As Integer = Convert.ToInt16(tbl1(1)) - 1
        Dim i2 As Integer = Convert.ToInt16(tbl1(3))
        Dim Fin As Integer = Canaux.Item(i1).GridCourbes(i2).Rows - 1


        If tbl1.Count >= 5 Then
            'For i = 1 To Canaux.Item(i1).GridCourbes(i2).Cols - 1
            Canaux.Item(i1).GridCourbes(i2).Visible = False
            Canaux.Item(i1).GridCourbes(i2).AutoRedraw = False
            For j = 4 To tbl1.Count - 1
                tbl2 = tbl1(j).Split()
                Ligne = tbl2(0)
                Colonne = tbl2(1)
                Canaux.Item(i1).GridCourbes(i2).Range(Ligne, Colonne, Fin, Colonne).BackColor = Det_CouleurCTRL(i2)
            Next j
            Canaux.Item(i1).GridCourbes(i2).AutoRedraw = True
            Canaux.Item(i1).GridCourbes(i2).Refresh()
            Canaux.Item(i1).GridCourbes(i2).Visible = True
            'Next i
        End If


    End Sub
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
        End Select
        Return coul
    End Function
    Public Sub Charger_ControlSys(LesCtrlSys As String)
        Dim i As Integer
        Dim tbl1() As String
        tbl1 = LesCtrlSys.Split(",")
        Dim i1 As Integer = Convert.ToInt16(tbl1(1)) - 1
        Canaux.Item(i1).Mute.Checked = tbl1(3)
        For i = 4 To (nbCourbes + 3)
            ListCCAct.Item(i1).LCCAct.Item(i - 4).Checked = tbl1(i)
        Next
    End Sub

    Public Sub Clear_Courbes()
        Dim i1 As Integer
        Dim i2 As Integer

        For i1 = 0 To nb_PistesVar - 1
            With Canaux.Item(i1)
                For i2 = 0 To nbCourbes - 1
                    .GridCourbes.Item(i2).AutoRedraw = False
                    .GridCourbes.Item(i2).Range(1, 1, .GridCourbes.Item(i2).Rows - 1, .GridCourbes.Item(i2).Cols - 1).ClearFormat()
                    .GridCourbes.Item(i2).AutoRedraw = True
                    .GridCourbes.Item(i2).Refresh()
                Next i2
            End With
        Next i1

    End Sub
    Public Sub Init_ControlSys()
        Dim i1 As Integer
        Dim i2 As Integer

        For i1 = 0 To nb_PistesVar - 1
            Canaux.Item(i1).Mute.Checked = False
            For i2 = 0 To nbCourbes - 1
                ListCCAct.Item(i1).LCCAct(i2).Checked = False
            Next i2
        Next i1

    End Sub

    ''' <summary>
    ''' Attacher : attacher le formulaire. Sert au menu Fichier/Quitter du formulaire
    ''' </summary>
    Sub Attacher(NumOnglet As Integer)
        Me.F4.Visible = False
        F4.FormBorderStyle = FormBorderStyle.None ' absolument nécessaire pour que l'attachement se passe bien et autre pour que le titre du formulaire disparaisse.
        Me.F4.TopMost = False   ' un seul des 2 suffit ?
        Me.F4.TopLevel = False
        Form1.TabControl2.TabPages.Item(NumOnglet).Controls.Add(Me.F4)
        Me.F4.Dock = DockStyle.Fill

        Me.F4.MainMenuStrip.Visible = False
        '
        If Langue = "fr" Then
            DockButton.Text = "Détacher"
        Else
            DockButton.Text = "UnDock"
        End If
        Me.F4.Visible = True
        Refresh_Général()

    End Sub
End Class
