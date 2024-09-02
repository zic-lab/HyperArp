Public Class Mixage
    Public F3 As New Form
    Private Langue As String
    Public WriteOnly Property PLangue As String
        Set(ByVal value As String)
            Me.Langue = value
        End Set
    End Property
    Public PisteVolume As New List(Of TrackBar) ' ces 3 composants sont publiques car ils sont accèdés par Form1.INIT_Piste
    Public labelAff As New List(Of Label)
    Public muteVolume As New List(Of CheckBox) ' Mutes des pistes
    ReadOnly soloPiste As New List(Of Button) ' bouton solo 
    ReadOnly LSauvMute As New List(Of Boolean)
    '
    ReadOnly labelVolume As New List(Of Label)
    ReadOnly LabelNom As New List(Of Label)
    ReadOnly Panneau As New SplitContainer
    ReadOnly Titre As New Label
    Public NomduSon As New List(Of TextBox)
    Private ReadOnly DockButton As New Button


    Public AutorisVol As New CheckBox ' checkbox de désactivation de la table de misage
    ' 
    ReadOnly Plus5 As New Button
    ReadOnly Moins5 As New Button
    ReadOnly Max As New Button
    '
    Public ReadOnly Send As New Button
    Public ReadOnly AutresV As New CheckBox
    '
    ' Menu4
    Private ReadOnly Menu1 As New System.Windows.Forms.MenuStrip()
    Private ReadOnly Fichier As New System.Windows.Forms.ToolStripMenuItem()
    Private ReadOnly Quitter As New System.Windows.Forms.ToolStripMenuItem()


    ' Fonts
    ' *****
    'ReadOnly fnt1 As New System.Drawing.Font("Calibri", 13, FontStyle.Regular)
    'ReadOnly fnt2 As New System.Drawing.Font("Calibri", 15, FontStyle.Regular)
    'ReadOnly fnt3 As New System.Drawing.Font("Tahoma", 10, FontStyle.Regular)
    'ReadOnly fnt4 As New System.Drawing.Font("Calibri", 18, FontStyle.Bold)
    ReadOnly fnt5 As New System.Drawing.Font("Calibri", 12, FontStyle.Regular)
    ReadOnly fnt6 As New Font("Verdana", 8, FontStyle.Regular)
    'ReadOnly fnt7 As New Font("Tahoma", 8.25, FontStyle.Regular)
    ReadOnly fnt8 As New Font("Tahoma", 8.25, FontStyle.Bold)
    ReadOnly fnt9 As New Font("Mistral", 15, FontStyle.Regular)
    'Sub New()


    'End Sub
    Public Sub Construction_Formulaire()
        Construction_F3()
        Construction_Barrout()
        Construction_Table()
        Construction_Menu()
        AutresV_True() ' affichage des curseurs de 11 à 14
        AddHandler F3.KeyUp, AddressOf F3_KeyUp
        Me.F3.KeyPreview = True ' pour réception des touches F4 et F5 (stop, play (recalcul) 
    End Sub
    Private Sub F3_KeyUp(sender As Object, e As KeyEventArgs)

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
        ' STOP : F4
        ' *********
        If e.KeyCode = Keys.F4 Then
            Form1.StopPlay()
        End If
    End Sub
    Sub Construction_F3()
        F3.Controls.Add(Panneau)
        Panneau.Dock = DockStyle.Fill
        Panneau.Orientation = Orientation.Horizontal
        Panneau.SplitterDistance = 50
        Panneau.IsSplitterFixed = True
        Panneau.BorderStyle = BorderStyle.None
        Panneau.Panel1.BorderStyle = BorderStyle.FixedSingle
        Panneau.Panel1.BackColor = Color.FromArgb(240, 240, 240)
        Panneau.Panel2.BackColor = Color.FromArgb(240, 240, 240)
        Panneau.FixedPanel = FixedPanel.Panel1
        '
        If LangueIHM = "fr" Then
            F3.Text = "MIX"
        Else
            F3.Text = "MIX"
        End If
        F3.ControlBox = False
    End Sub
    Sub Construction_Barrout()
        '
        ' bouton -5
        Panneau.Panel1.Controls.Add(Moins5)
        Moins5.Text = "-5"
        Moins5.Size = New Size(45, 45)
        Moins5.Location = New Point(3, 2) '70,10
        Moins5.BackColor = Color.BurlyWood
        Moins5.TabStop = False
        Moins5.TabIndex = 1
        Moins5.ResumeLayout(False)

        '
        ' bouton +5
        Panneau.Panel1.Controls.Add(Plus5)
        Plus5.Text = "+5"
        Plus5.Size = New Size(45, 45)
        Plus5.Location = New Point(70, 2) ' 3,10
        Plus5.BackColor = Color.BurlyWood
        '
        ' bouton normaliser
        Panneau.Panel1.Controls.Add(Max)
        If Me.Langue = "fr" Then
            Max.Text = "Normaliser"
        Else
            Max.Text = "Normalize"
        End If
        '
        Max.Size = New Size(80, 45)
        Max.Location = New Point(135, 2)
        Max.BackColor = Color.BurlyWood
        '
        ' Activation/Désactivation de la table de mixage
        ' **********************************************
        Panneau.Panel1.Controls.Add(AutorisVol)
        If Me.Langue = "fr" Then
            AutorisVol.Text = "MIX Activation"
        Else
            AutorisVol.Text = "MIX Activation"
        End If
        '
        AutorisVol.Size = New Size(100, 42)
        AutorisVol.Font = fnt8
        AutorisVol.Location = New Point(238, 2)
        AutorisVol.BackColor = Color.BurlyWood
        AutorisVol.Checked = True
        '
        ' Bouton Send
        ' ***********
        Panneau.Panel1.Controls.Add(Send)
        Send.Size = New Size(210, 30)
        Send.Location = New Point(1000, 18) '70,10
        Send.UseVisualStyleBackColor = True
        Send.FlatStyle = FlatStyle.Standard
        Send.TextAlign = ContentAlignment.MiddleLeft
        Send.Enabled = False
        Send.TabStop = False
        Send.ResumeLayout(False)
        '
        If LangueIHM = "fr" Then
            Send.Text = "Envoyer tous les volumes"
        Else
            Send.Text = "Send allvolumes"
        End If
        '
        ' Check AutresV
        ' *************
        Panneau.Panel1.Controls.Add(AutresV)
        AutresV.Location = New Point(1000, 1)
        AutresV.AutoSize = True
        If LangueIHM = "fr" Then
            AutresV.Text = "Utiliser les autres volumes MIDI"
        Else
            AutresV.Text = "Using others MIDI volume"
        End If
        AutresV.Visible = False
        AutresV.Checked = False


        ' boutton attacher/détacher
        Panneau.Panel1.Controls.Add(DockButton)
        Dim P As New Point(595, 2)
        Dim s As New Size(100, 45)
        '
        DockButton.FlatStyle = FlatStyle.Standard
        DockButton.BackColor = Color.DarkOliveGreen
        DockButton.ForeColor = Color.Yellow
        DockButton.Font = fnt5
        DockButton.Location = P
        DockButton.Size = s
        DockButton.AutoSize = True
        '
        If LangueIHM = "fr" Then
            DockButton.Text = "Détacher"
        Else
            DockButton.Text = "Undock"
        End If
        '
        DockButton.Visible = True


        'Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) 
        AddHandler AutresV.CheckedChanged, AddressOf AutresV_CheckedChanged
        AddHandler Send.Click, AddressOf Send_Click
        AddHandler Moins5.Click, AddressOf Moins5_Click
        AddHandler Plus5.Click, AddressOf Plus5_Click
        AddHandler Max.Click, AddressOf Max_Click
        AddHandler DockButton.MouseClick, AddressOf DockButton_MouseClic
        AddHandler AutorisVol.MouseClick, AddressOf AutorisVol_MouseClick
    End Sub
    Sub Send_Click(Sender As Object, e As EventArgs)

        Dim pst As Integer
        Dim i As Integer = Form1.ChoixSortieMidi
        Dim canal As Byte

        For pst = 0 To nb_TotalPistes - 1
            If Not (EnChargement) And AutorisVol.Checked And PisteVolume.Item(pst).Enabled And muteVolume.Item(pst).Checked Then
                Dim volume As Byte = CByte(PisteVolume.Item(pst).Value)
                labelAff.Item(pst).Text = Convert.ToString(volume)
                canal = LesPistes.Item(pst).Canal
                If Not (Form1.SortieMidi.Item(Form1.ChoixSortieMidi).IsOpen) Then
                    Form1.SortieMidi.Item(Form1.ChoixSortieMidi).Open()
                End If
                Form1.SortieMidi.Item(Form1.ChoixSortieMidi).SendControlChange(canal, CVolume, volume)
            End If
        Next
    End Sub
    Sub AutresV_CheckedChanged(sender As Object, e As EventArgs)
        If AutresV.Checked Then
            AutresV_True()
            Send.Enabled = True
        Else
            AutresV_False()
            Send.Enabled = False
        End If

    End Sub
    Public Sub AutresV_True()
        Dim i As Integer
        For i = 10 To nb_TotalPistes - 1
            muteVolume.Item(i).Visible = True ' mute
            NomduSon.Item(i).Visible = True
            PisteVolume.Item(i).Visible = True
            labelAff.Item(i).Visible = True
        Next
    End Sub
    Public Sub AutresV_False()
        Dim i As Integer
        For i = 10 To nb_TotalPistes - 1
            muteVolume.Item(i).Visible = False ' mute
            NomduSon.Item(i).Visible = False
            PisteVolume.Item(i).Visible = False
            labelAff.Item(i).Visible = False
        Next
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
        ' Quitter (de Fichier)
        '
        If Me.Langue = "fr" Then
            Quitter.Text = "Attacher"
        Else
            Quitter.Text = "Dock"
        End If
        Quitter.ShortcutKeys = Shortcut.CtrlD
        Quitter.Size = New System.Drawing.Size(180, 22)
        '
        Menu1.Text = "Menu"
        Me.Menu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Fichier})
        F3.MainMenuStrip = Menu1
        F3.MainMenuStrip.Visible = False
        Menu1.Location = New System.Drawing.Point(0, 0)
        F3.Controls.Add(Menu1)
        F3.MainMenuStrip = Menu1

        AddHandler Quitter.Click, AddressOf Quitter_Click
    End Sub
    Sub Quitter_Click(Sender As Object, e As EventArgs)
        Dim NumOnglet As Integer = Convert.ToUInt16(F3.Tag)
        Attacher(NumOnglet)
    End Sub
    Sub Attacher(NumOnglet As Integer)
        Me.F3.FormBorderStyle = FormBorderStyle.None
        Me.F3.TopMost = False   ' un seul des 2 suffit ?
        Me.F3.TopLevel = False
        F3.MainMenuStrip.Visible = False
        Form1.TabControl4.TabPages.Item(NumOnglet).Controls.Add(Me.F3)
        Me.F3.Dock = DockStyle.Fill
        '
        If LangueIHM = "fr" Then
            DockButton.Text = "Détacher"
        Else
            DockButton.Text = "UnDock"
        End If
    End Sub


    Private Sub DockButton_MouseClic(sender As Object, e As EventArgs)
        Dim com As Button = sender
        Dim ind As Integer = com.Tag

        Dim NumOnglet As Integer = F3.Tag
        If Me.F3.Dock = DockStyle.Fill Then
            Me.F3.FormBorderStyle = FormBorderStyle.Sizable ' l
            Me.F3.TopMost = True   ' 
            Me.F3.Dock = DockStyle.None
            Form1.TabControl4.TabPages.Item(NumOnglet).Controls.Remove(Me.F3)
            Dim p As New Point With {
                .X = Me.F3.Location.X + 20,
                .Y = Me.F3.Location.Y + 50
            }
            Dim s As New Size With {
            .Width = Me.F3.Size.Width + 25, '+ 500, 'Panneau3.Width,
            .Height = Me.F3.Size.Height + 18'Panneau3.Height + 400
        }
            Me.F3.Location = p
            Me.F3.Size = s
            '
            Me.F3.StartPosition = FormStartPosition.Manual ' permet de tenir compte de la location calculée dans p
            '
            Me.F3.TopMost = True
            Me.F3.TopLevel = True
            If LangueIHM = "fr" Then
                DockButton.Text = "Attacher"
            Else
                DockButton.Text = "Dock"
            End If
            F3.MainMenuStrip.Visible = True
        Else
            Me.F3.FormBorderStyle = FormBorderStyle.None
            Me.F3.TopMost = False   ' un seul des 2 suffit ?
            Me.F3.TopLevel = False
            F3.MainMenuStrip.Visible = False
            Form1.TabControl4.TabPages.Item(NumOnglet).Controls.Add(Me.F3)
            Me.F3.Dock = DockStyle.Fill
            '
            If LangueIHM = "fr" Then
                DockButton.Text = "Détacher"
            Else
                DockButton.Text = "UnDock"
            End If
        End If
    End Sub
    Private Sub Moins5_Click(sender As Object, e As EventArgs)
        '
        Dim minV, i, j As Integer
        minV = Det_MinVolume()
        j = minV - 5
        If j >= 0 Then
            For i = 0 To nb_TotalPistes - 1
                PisteVolume.Item(i).Value = PisteVolume.Item(i).Value - 5
                labelAff.Item(i).Text = Convert.ToString(PisteVolume.Item(i).Value)
            Next i
        End If
        '
        Form1.Send_AllVolumes()
    End Sub
    Private Sub Plus5_Click(sender As Object, e As EventArgs)

        Dim maxV, i, j As Integer
        maxV = Det_MaxVolume()
        j = maxV + 5
        If j <= 127 Then
            For i = 0 To nb_TotalPistes - 1
                PisteVolume.Item(i).Value = PisteVolume.Item(i).Value + 5
                labelAff.Item(i).Text = Convert.ToString(PisteVolume.Item(i).Value)
            Next i
        End If
        '
        Form1.Send_AllVolumes()

    End Sub

    Private Sub Max_Click(sender As Object, e As EventArgs)

        Dim maxV, i, j As Integer
        maxV = Det_MaxVolume()
        j = 127 - maxV
        '
        If j > 0 Then
            For i = 0 To nb_TotalPistes - 1
                PisteVolume.Item(i).Value = PisteVolume.Item(i).Value + j
                labelAff.Item(i).Text = Convert.ToString(PisteVolume.Item(i).Value)
            Next i
            Form1.ValCompress = j
        Else
            If Form1.ValCompress <> -1 Then
                For i = 0 To nb_TotalPistes - 1
                    PisteVolume.Item(i).Value = PisteVolume.Item(i).Value - Form1.ValCompress
                    labelAff.Item(i).Text = Convert.ToString(PisteVolume.Item(i).Value)
                Next i
            End If
        End If
        Form1.Send_AllVolumes()

    End Sub
    Function Det_MaxVolume() As Integer
        Dim i, v As Integer
        Dim MaxV As Integer
        '
        Dim List1 As New List(Of Integer)

        For i = 0 To nb_TotalPistes - 1 'Arrangement1.Nb_PistesMidi
            v = PisteVolume.Item(i).Value ' Me.Récup_Volume(i) 'PisteVolume.Item(i).Value
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
            v = PisteVolume.Item(i).Value 'Me.Récup_Volume(i)
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
    Sub Construction_Table()
        Dim i, j As Integer
        Dim PP As Integer = 100 ' constant servant au positionnement des éléments sur l'axe des x
        Dim iii, jjj As Integer
        '
        Panneau.Panel1.BackColor = Color.LightGray
        Panneau.Panel2.BackColor = Color.LightGray
        For i = 0 To Module1.nb_TotalPistes - 1
            '
            jjj = 0
            If i > 5 Then jjj = 10
            '
            PisteVolume.Add(New TrackBar)
            labelVolume.Add(New Label)
            labelAff.Add(New Label)
            muteVolume.Add(New CheckBox) ' activation de la tranche (en haut de la track bar)
            soloPiste.Add(New Button)    ' bouton de passage en mode solo
            LabelNom.Add(New Label)
            NomduSon.Add(New TextBox)
            '
            Panneau.Panel2.Controls.Add(PisteVolume.Item(i))
            Panneau.Panel2.Controls.Add(labelVolume.Item(i))
            Panneau.Panel2.Controls.Add(labelAff.Item(i))
            Panneau.Panel2.Controls.Add(muteVolume.Item(i))
            Panneau.Panel2.Controls.Add(soloPiste.Item(i))
            Panneau.Panel2.Controls.Add(LabelNom.Item(i))
            Panneau.Panel2.Controls.Add(NomduSon.Item(i))
            Panneau.Panel2.AutoScroll = True
            '
            ' Mute des pistes
            muteVolume.Item(i).Location = New Point(5 + (i * PP) - jjj, 420) ' checkbox
            muteVolume.Item(i).Visible = True
            muteVolume.Item(i).Checked = True
            muteVolume.Item(i).Text = "PISTE" + Str(i + 1)
            muteVolume.Item(i).Size = New Size(85, 30)
            muteVolume.Item(i).Font = Module1.fontMutePiste
            muteVolume.Item(i).Tag = i
            muteVolume.Item(i).BringToFront()
            '
            ' Nom du son (en bas)
            NomduSon.Item(i).Location = New Point(5 + (i * PP) - jjj, 455) ' checkbox
            NomduSon.Item(i).Visible = True
            'If i = 9 Then NomduSon.Item(i).Visible = False ' pas de nom du son pour la batterie
            NomduSon.Item(i).Text = ""
            NomduSon.Item(i).Size = New Size(82, 20)
            NomduSon.Item(i).Font = Module1.fontNomduSon
            NomduSon.Item(i).BackColor = Color.Beige
            NomduSon.Item(i).ForeColor = Color.Black
            NomduSon.Item(i).BorderStyle = BorderStyle.FixedSingle
            NomduSon.Item(i).Tag = i
            AddHandler NomduSon.Item(i).TextChanged, AddressOf NomduSon_TextChanged

            ' Nom des Pistes 
            LabelNom.Item(i).Location = New Point(5 + (i * PP) - jjj, 5) ' checkbox
            LabelNom.Item(i).Visible = True
            'LabelNom.Item(i).Text = "Piste"
            LabelNom.Item(i).Font = fnt8
            LabelNom.Item(i).AutoSize = True
            LabelNom.Item(i).Tag = i
            '
            Select Case i
                Case 0, 1, 2, 3, 4, 5
                    LabelNom.Item(i).Visible = False
                    LabelNom.Item(i).ForeColor = Color.Red
                    muteVolume.Item(i).ForeColor = Color.DarkRed
                    If Me.Langue = "fr" Then
                        LabelNom.Item(i).Text = "Canal" + Str(i + 1)
                    Else
                        LabelNom.Item(i).Text = "Channel" + Str(i + 1)
                    End If
                    PisteVolume.Item(i).BackColor = Color.DarkKhaki

                Case 6, 7, 8, 10, 11, 12, 13
                    j = i
                    If i >= 10 Then j -= 1
                    LabelNom.Item(i).ForeColor = Color.MediumSeaGreen
                    iii = j + 1
                    If j > 8 Then iii = j + 2
                    LabelNom.Item(i).Text = (iii).ToString + "-PianoRoll"
                    muteVolume.Item(i).ForeColor = Color.Blue
                    PisteVolume.Item(i).BackColor = Color.MediumSeaGreen
                Case 9
                    LabelNom.Item(i).ForeColor = Color.RoyalBlue
                    muteVolume.Item(i).ForeColor = Color.RoyalBlue
                    If Langue = "fr" Then
                        LabelNom.Item(i).Text = "Batterie"
                    Else
                        LabelNom.Item(i).Text = "Drums"
                    End If
                    PisteVolume.Item(i).BackColor = Color.LightSkyBlue
            End Select

            '

            PisteVolume.Item(i).Location = New Point(1 + (i * PP) - jjj, 50) 'track bar ' la valeur 57 dans le calcul de X permet de placer 2 autres Tracbar (9 et 10)
            PisteVolume.Item(i).Orientation = Orientation.Vertical
            PisteVolume.Item(i).Minimum = 0
            PisteVolume.Item(i).Maximum = 127
            PisteVolume.Item(i).Value = 100
            PisteVolume.Item(i).Size = New Size(15, 320) '
            PisteVolume.Item(i).Visible = True
            PisteVolume.Item(i).LargeChange = 1
            PisteVolume.Item(i).SmallChange = 1
            PisteVolume.Item(i).Tag = i
            PisteVolume.Item(i).TickStyle = TickStyle.Both
            '
            '
            labelAff.Item(i).Location = New Point(5 + (i * PP) - jjj, 30) ' valeur du volume en haut de la trackbar
            labelAff.Item(i).Visible = True
            labelAff.Item(i).AutoSize = True
            labelAff.Item(i).BackColor = Color.Beige 'Color.FromArgb(240, 240, 240)
            labelAff.Item(i).ForeColor = Color.Blue
            labelAff.Item(i).Font = fnt6
            labelAff.Item(i).Text = Convert.ToString(PisteVolume.Item(i).Value)
            labelAff.Item(i).BorderStyle = BorderStyle.None
            '
            soloPiste.Item(i).Location = New Point(3 + (i * PP) - jjj, 390) ' bouton solo sous les track bars
            soloPiste.Item(i).Font = New Font("Calibri", 8, FontStyle.Regular)
            soloPiste.Item(i).Size = New Size(50, 20)
            soloPiste.Item(i).BackColor = Color.Beige
            soloPiste.Item(i).Text = "SOLO"
            soloPiste.Item(i).Tag = i
            soloPiste.Item(i).Visible = False
            '
            labelVolume.Item(i).Location = New Point(7 + (i * PP) - jjj, 420) ' label nom des track bar en bas des track bars
            labelVolume.Item(i).Visible = False
            labelVolume.Item(i).Font = fnt6
            labelVolume.Item(i).AutoSize = True ' New Size(55, 20)
            If Module1.LangueIHM = "fr" Then
                labelVolume.Item(i).Text = "Piste" + Str(i + 1)
            Else
                labelVolume.Item(i).Text = "Track" + Str(i + 1)
            End If
            '
            'AddHandler PisteVolume.Item(i).MouseDown, AddressOf PisteVolume_MouseDown
            AddHandler PisteVolume.Item(i).Scroll, AddressOf PisteVolume_Scroll
            'AddHandler muteVolume.Item(i).MouseUp, AddressOf muteVolume_MouseUp
            AddHandler muteVolume.Item(i).CheckedChanged, AddressOf muteVolume_CheckedChanged
            AddHandler soloPiste.Item(i).MouseUp, AddressOf SoloPiste_MouseUp

        Next i

        ' Activation de la table de mixage
        ' ********************************
        RemoveHandler AutorisVol.MouseClick, AddressOf AutorisVol_MouseClick
        AutorisVol.Checked = False
        AddHandler AutorisVol.MouseClick, AddressOf AutorisVol_MouseClick
        VolumesEnabled(False)




        Panneau.Panel2.Controls.Add(Titre)
        ' Titre des pistes HyperArp
        ' 
        Titre.Location = New Point(5, 3)
        Titre.Visible = True
        Titre.AutoSize = False
        Titre.BackColor = Color.DarkKhaki
        Titre.ForeColor = Color.DarkRed
        Titre.Size = New Size(530, 25)
        Titre.TextAlign = ContentAlignment.TopCenter
        Titre.Font = fnt9
        Titre.Text = "HyperArp"
        Titre.BringToFront()
    End Sub
    Public Sub AutorisVol_MouseClick(sender As Object, e As EventArgs)
        VolumesEnabled(AutorisVol.Checked)
    End Sub
    Public Sub VolumesEnabled(act As Boolean)
        '
        Moins5.Enabled = act
        Plus5.Enabled = act
        Max.Enabled = act
        'Send.Enabled = act
        AutresV.Enabled = act
        'Devant.Enabled = act
        'DockButton.Enabled = act

        Panneau.Panel2.Enabled = act
        '
        If act Then
            AutorisVol.BackColor = Color.BurlyWood
            AutorisVol.ForeColor = Color.Black
        Else
            AutorisVol.BackColor = Color.Red
            AutorisVol.ForeColor = Color.White
            Send.Enabled = False
        End If

    End Sub
    Public Function AutorisVolumes() As String
        Dim a As String = "AutorisVolumes,"

        a += AutorisVol.Checked.ToString()
        Return a
    End Function
    Public Function AutresVol() As String
        Dim a As String = "AutresVol,"

        a += AutresV.Checked.ToString()
        Return a
    End Function
    Public Sub Maj_AutorisVolumes(autorisV As String)
        Dim tbl() As String

        RemoveHandler AutorisVol.MouseClick, AddressOf AutorisVol_MouseClick
        tbl = autorisV.Split(",")
        If tbl(1) = "True" Then
            AutorisVol.Checked = True
            VolumesEnabled(True)
        Else
            AutorisVol.Checked = False
            VolumesEnabled(False)
        End If
        AddHandler AutorisVol.MouseClick, AddressOf AutorisVol_MouseClick
    End Sub
    Public Sub Maj_AutresVolumes(ligne As String)
        Dim tbl() As String

        RemoveHandler AutresV.CheckedChanged, AddressOf AutresV_CheckedChanged
        tbl = ligne.Split(",")
        If tbl(1) = "True" Then
            AutresV.Checked = True
            If AutorisVol.Checked = True Then
                Send.Enabled = True
            End If
        End If
        AddHandler AutresV.CheckedChanged, AddressOf AutresV_CheckedChanged
    End Sub
    Sub Barr_00()
        RemoveHandler AutorisVol.MouseClick, AddressOf AutorisVol_MouseClick
        AutorisVol.Checked = False
        VolumesEnabled(False)
        AddHandler AutorisVol.MouseClick, AddressOf AutorisVol_MouseClick
        '
        RemoveHandler AutresV.CheckedChanged, AddressOf AutresV_CheckedChanged
        AutresV.Checked = False
        AutresV.Enabled = False
        Send.Enabled = False
        AddHandler AutresV.CheckedChanged, AddressOf AutresV_CheckedChanged
    End Sub
    Sub Barr_01()
        Barr_00()
    End Sub
    Sub Barr_10()
        RemoveHandler AutorisVol.MouseClick, AddressOf AutorisVol_MouseClick
        AutorisVol.Checked = True
        VolumesEnabled(True)
        AddHandler AutorisVol.MouseClick, AddressOf AutorisVol_MouseClick
        '
        RemoveHandler AutresV.CheckedChanged, AddressOf AutresV_CheckedChanged
        AutresV.Checked = False
        AutresV.Enabled = True
        Send.Enabled = False
        AddHandler AutresV.CheckedChanged, AddressOf AutresV_CheckedChanged
    End Sub
    Sub Barr_11()
        RemoveHandler AutorisVol.MouseClick, AddressOf AutorisVol_MouseClick
        AutorisVol.Checked = True
        VolumesEnabled(True)
        AddHandler AutorisVol.MouseClick, AddressOf AutorisVol_MouseClick
        '
        RemoveHandler AutresV.CheckedChanged, AddressOf AutresV_CheckedChanged
        AutresV.Enabled = True
        AutresV.Checked = True
        Send.Enabled = True
        AddHandler AutresV.CheckedChanged, AddressOf AutresV_CheckedChanged
        '
        AutresV_True()
    End Sub
    Public Sub Maj_Barr()
        If Not (AutorisVol.Checked) And Not (AutresV.Checked) Then
            Barr_00()
        End If
        If Not (AutorisVol.Checked) And AutresV.Checked Then
            Barr_01()
        End If
        If AutorisVol.Checked And Not (AutresV.Checked) Then
            Barr_10()
        End If
        If AutorisVol.Checked And AutresV.Checked Then
            Barr_11()
        End If
    End Sub
    Public Sub NomduSon_TextChanged(sender As Object, e As EventArgs)
        Dim com As TextBox = sender
        Dim ind As Integer
        ind = Val(com.Tag)
        Dim canal As Integer = ind
        Dim i, j, k As Integer
        Dim a As String

        If InStr(Trim(NomduSon.Item(ind).Text), ",") <> 0 Or InStr(Trim(NomduSon.Item(ind).Text), "&") <> 0 Then ' on retire les caractères qui peuvent servir de séparateurs
            k = NomduSon.Item(ind).SelectionStart
            a = NomduSon.Item(ind).Text.Replace(",", "")
            NomduSon.Item(ind).Text = Trim(a)
            a = NomduSon.Item(ind).Text.Replace("&", "")
            NomduSon.Item(ind).Text = Trim(a) '
            NomduSon.Item(ind).SelectionStart = k - 1
        End If
        '
        Select Case canal
            Case 0, 1, 2, 3, 4, 5
                For i = 0 To nb_Variations - 1
                    j = Form1.Det_IndexGénérateur(canal, i)
                    RemoveHandler Form1.NomduSon.Item(j).TextChanged, AddressOf Form1.NomduSon_TextChanged
                    Form1.NomduSon.Item(j).Text = Trim(NomduSon.Item(ind).Text)
                    AddHandler Form1.NomduSon.Item(j).TextChanged, AddressOf Form1.NomduSon_TextChanged
                Next
            Case 6, 7, 8
                RemoveHandler Form1.listPIANOROLL.Item(ind - 6).NomduSon.TextChanged, AddressOf Form1.listPIANOROLL.Item(ind - 6).NomduSon_TextChanged
                'Form1.Mix.NomduSon.Item(ind - 1).Text = Trim(NomduSon.Item(ind).Text)
                Form1.listPIANOROLL.Item(ind - 6).NomduSon.Text = Trim(NomduSon.Item(ind).Text)
                AddHandler Form1.listPIANOROLL.Item(ind - 6).NomduSon.TextChanged, AddressOf Form1.listPIANOROLL.Item(ind - 6).NomduSon_TextChanged
        End Select
    End Sub
    Sub PisteVolume_Scroll(sender As Object, e As EventArgs)
        Dim com As TrackBar = sender
        Dim pst As Integer
        Dim i As Integer = Form1.ChoixSortieMidi

        pst = Val(com.Tag)

        If EnChargement = False Then
            Dim volume As Byte = CByte(PisteVolume.Item(pst).Value)
            labelAff.Item(pst).Text = Convert.ToString(volume)
            Dim canal As Byte = LesPistes.Item(pst).Canal
            If Not (Form1.SortieMidi.Item(Form1.ChoixSortieMidi).IsOpen) Then
                Form1.SortieMidi.Item(Form1.ChoixSortieMidi).Open()
            End If
            Form1.SortieMidi.Item(Form1.ChoixSortieMidi).SendControlChange(canal, CVolume, volume)
        End If
    End Sub
    Sub muteVolume_CheckedChanged(sender As Object, e As EventArgs)
        Dim i As Integer
        Dim com As CheckBox = sender
        Dim ind As Integer
        ind = Val(com.Tag)
        Dim b As Boolean = muteVolume.Item(ind).Checked
        '
        If b Then
            PisteVolume.Item(ind).Enabled = True
            RétablirVolume(ind)
        Else
            PisteVolume.Item(ind).Enabled = False
            CouperVolume(ind)
        End If
        '
        If ind >= 0 And ind <= 5 Then ' Pistes des variations
            Dim tbl() As String = Form1.N_BLOC_MIDI(ind).Split()
            For Each a As String In tbl
                i = Convert.ToInt16(a)
                RemoveHandler Form1.PisteMute.Item(i).CheckedChanged, AddressOf Form1.PisteMute_CheckedChange
                Form1.PisteMute.Item(i).Checked = muteVolume.Item(ind).Checked
                AddHandler Form1.PisteMute.Item(i).CheckedChanged, AddressOf Form1.PisteMute_CheckedChange
            Next
        Else
            'If ind <> 9 Then
            If ind >= 6 And ind <= 8 Then ' Pistes Piano roll
                i = ind - 6
                If ind >= 10 Then i -= 1
                RemoveHandler Form1.listPIANOROLL.Item(i).CheckMute.CheckedChanged, AddressOf Form1.listPIANOROLL(i).CheckMute_CheckedChanged
                Form1.listPIANOROLL(i).PMute = muteVolume.Item(ind).Checked
                AddHandler Form1.listPIANOROLL.Item(i).CheckMute.CheckedChanged, AddressOf Form1.listPIANOROLL(i).CheckMute_CheckedChanged
            Else
                If ind = 9 Then ' batterie
                    RemoveHandler Form1.Drums.CheckMute.CheckedChanged, AddressOf Form1.Drums.CheckMute_CheckedChanged
                    Form1.Drums.PMute = muteVolume.Item(ind).Checked
                    AddHandler Form1.Drums.CheckMute.CheckedChanged, AddressOf Form1.Drums.CheckMute_CheckedChanged
                End If
            End If
        End If
    End Sub
    Sub RétablirVolume(pst As Integer)
        If EnChargement = False Then
            Dim volume As Byte = CByte(PisteVolume.Item(pst).Value)
            labelAff.Item(pst).Text = Convert.ToString(volume)
            Dim canal As Byte = LesPistes.Item(pst).Canal
            If Not (Form1.SortieMidi.Item(Form1.ChoixSortieMidi).IsOpen) Then
                Form1.SortieMidi.Item(Form1.ChoixSortieMidi).Open()
            End If
            Form1.SortieMidi.Item(Form1.ChoixSortieMidi).SendControlChange(canal, 7, volume)
        End If
    End Sub
    Sub CouperVolume(pst As Integer)
        If EnChargement = False Then
            Dim volume As Byte = 0
            labelAff.Item(pst).Text = Convert.ToString(volume)
            Dim canal As Byte = LesPistes.Item(pst).Canal
            If Not (Form1.SortieMidi.Item(Form1.ChoixSortieMidi).IsOpen) Then
                Form1.SortieMidi.Item(Form1.ChoixSortieMidi).Open()
            End If
            Form1.SortieMidi.Item(Form1.ChoixSortieMidi).SendControlChange(canal, 7, volume)
        End If
    End Sub
    Public Sub GestMute()
        ' sauvegarde/restit état des mutes
        ' ********************************
        If Form1.PisteSolo = -1 Then
            LSauvMute.Clear()
            For i = 0 To PisteVolume.Count - 1
                LSauvMute.Add(muteVolume.Item(i).Checked)
            Next i
        Else
            For i = 0 To PisteVolume.Count - 1
                muteVolume.Item(i).Checked = LSauvMute.Item(i)
            Next i
        End If
    End Sub
    Sub SoloPiste_MouseUp(sender As Object, e As EventArgs) ' bouton solo de la piste
        Dim com As Button = sender
        Dim ind As Integer
        Dim i As Integer
        '
        ind = Val(com.Tag)
        '
        GestMute()
        ' Gestion de la table de mixage
        Gestion_Solo2(ind)

        ' Gestion des solo de HyperArp, PianoRoll et Batterie
        Select Case ind
            Case 0, 1, 2, 3, 4, 5 ' --> HyperArp
                Form1.TraitementSoloHypA(ind)

            Case 6, 7, 8 ' --> PianoRoll
                i = ind - 6
                If i >= 10 Then
                    i = ind - 7
                End If
                Form1.listPIANOROLL(i).TraitementSoloPR(ind)

            Case 9 ' --> Batterie
                Form1.Drums.TraitementSoloDRM(ind)

        End Select

    End Sub
    ''' <summary>
    ''' Gestion_Solo :procédure réservée pour la gestion des solos à partir de la table de mixage
    ''' </summary>
    ''' <param name="ind"></param>
    Public Sub Gestion_Solo(ind As Integer)
        Dim i As Integer

        If ind <> Form1.PisteSolo Then ' activation du mode solo
            For i = 0 To PisteVolume.Count - 1
                If i <> ind Then
                    PisteVolume.Item(i).Enabled = False  ' PisteVolume est le trackbar du volume
                    muteVolume.Item(i).Checked = False   ' muteVolume est le checkbox d'activation d'une tranche placé en haut des trackbar
                    CouperVolume(i)
                End If
            Next
            PisteVolume.Item(ind).Enabled = True
            muteVolume.Item(ind).Checked = True
            Form1.PisteSolo = ind
            RétablirVolume(ind)
            '
            soloPiste.Item(ind).BackColor = Color.OrangeRed
            soloPiste.Item(ind).ForeColor = Color.Yellow
        Else                  ' rétablissement du mode normal
            For i = 0 To PisteVolume.Count - 1
                PisteVolume.Item(i).Enabled = True
                muteVolume.Item(i).Checked = True
                RétablirVolume(i)
            Next
            Form1.PisteSolo = -1
            '
            soloPiste.Item(ind).BackColor = Color.Beige
            soloPiste.Item(ind).ForeColor = Color.Black
        End If
    End Sub

    ''' <summary>
    ''' Gestion_Solo2 : procédure réservée pour la gestion des solos des pistes HyperArp (1-6) à partir d'HyperArp, PianoRoll et Batterie
    ''' </summary>
    ''' <param name="ind">N° canal MIDI</param>
    Public Sub Gestion_Solo2(ind As Integer)
        Dim i As Integer

        If ind <> Form1.PisteSolo Then ' activation du mode solo
            For i = 0 To PisteVolume.Count - 1
                If i <> ind Then
                    PisteVolume.Item(i).Enabled = False  ' PisteVolume est le trackbar du volume
                    muteVolume.Item(i).Checked = False   ' muteVolume est le checkbox d'activation d'une tranche placé en haut des trackbar
                    soloPiste.Item(i).Enabled = False     '  sauvegarde de l'état des mutes
                    CouperVolume(i)
                End If
            Next
            PisteVolume.Item(ind).Enabled = True
            muteVolume.Item(ind).Checked = True
            soloPiste.Item(ind).Enabled = True
            Form1.PisteSolo = ind
            soloPiste.Item(ind).BackColor = Color.Red
            soloPiste.Item(ind).ForeColor = Color.Yellow
            RétablirVolume(ind)
        Else                  ' rétablissement du mode normal
            For i = 0 To PisteVolume.Count - 1
                PisteVolume.Item(i).Enabled = True
                soloPiste.Item(i).Enabled = True
                If muteVolume.Item(i).Checked = True Then
                    RétablirVolume(i)
                End If
            Next
            soloPiste.Item(ind).BackColor = Color.Beige
            soloPiste.Item(ind).ForeColor = Color.Black
            Form1.PisteSolo = -1
        End If
    End Sub
End Class
