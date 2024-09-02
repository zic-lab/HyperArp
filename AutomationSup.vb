Public Class AutomationSup
    ' Constantes pour les lignes contenant les checkBox
    ' *************************************************
    Private Const LigHyperArp1 = 3
    Private Const LigHyperArp2 = 4
    Private Const LigHyperArp3 = 5
    Private Const LigHyperArp4 = 6
    Private Const LigHyperArp5 = 7
    Private Const LigHyperArp6 = 8
    '
    Private Const LigPianoRoll1 = 12
    Private Const LigPianoRoll2 = 13
    Private Const LigPianoRoll3 = 14
    Private Const LigPianoRoll4 = 15
    Private Const LigPianoRoll5 = 16

    Private Const ColBout2 = 3
    '
    ' Constantes pour les colonnes
    ' ****************************
    Private Const ColCanal = 3
    Private Const ColExp = 5
    Private Const ColMod = 6
    Private Const ColPan = 7
    Private Const ColCC50 = 8
    Private Const ColCC51 = 9
    Private Const ColCC52 = 10
    Private Const ColCC53 = 11

    Private Sub AutomationSup_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim i1, i2 As Integer
        Dim b As String

        ' Dimension de la fenêtre
        ' ***********************
        Me.Height = 418
        Me.Width = 765
        '
        ' Langue
        ' ******
        If Module1.LangueIHM = "fr" Then
            Button2.Text = "Annuler"
            Button3.Text = "Effacer tout"
            Me.Text = "Supervision de l'Automation"
        Else
            Button2.Text = "Cancel"
            Button3.Text = "Clear all"
            Button2.Text = "Automation Supervision"
        End If
        '

        ' Chargement de l'IHM à partir du fichier 'optimisationHA.flx' placé à la racine de l'exe
        ' ***************************************************************************************
        Dim a As String
        a = Module1.Création_CTemp() + "\" + "optimisationHA.flx" 'Pistes(0).Nom + ".mid" ' \FichierMIDI.mid" My.Application.Info.DirectoryPath
        My.Computer.FileSystem.WriteAllBytes(a, My.Resources.optimisationHA, False)
        Grid1.OpenFile(a)
        '
        ' Mise à jour Automation des Canaux HyperArp
        ' ******************************************
        For i1 = 0 To Form1.Automation1.ListCCAct.Count - 1
            For i2 = 0 To Form1.Automation1.ListCCAct.Item(i1).LCCAct.Count - 1
                b = "0"
                If Form1.Automation1.ListCCAct.Item(i1).LCCAct(i2).Checked Then
                    b = "1"
                End If
                Grid1.Cell(Det_NumHypArp1(i1), Det_NumCourbes(i2)).Text = b ' Det_NumHypArp1 : fournit le N° Ligne / Det_NumCourbes : fournit le N° de colonnes
            Next
        Next
        '
        ' Mise à jour des MUTES généraux des Automations HyperArp (par canaux)
        ' ********************************************************************
        For i1 = 0 To Form1.Automation1.Canaux.Count - 1
            i2 = Det_NumHypArp1(i1)
            Grid1.Cell(i2, ColCanal).Locked = False
            b = "0"
            If Form1.Automation1.Canaux.Item(i1).Mute.Checked Then
                b = "1"
            End If
            ' 
            Grid1.Cell(i2, ColCanal).Text = b
        Next
        '
        ' Mise à jour Automation des PIANOROLL
        ' ************************************
        For i1 = 0 To Form1.listPIANOROLL.Count - 1
            For i2 = 0 To Form1.listPIANOROLL.Item(i1).CCActif.Count - 1
                b = "0"
                If Form1.listPIANOROLL.Item(i1).CCActif.Item(i2).Checked Then
                    b = "1"
                End If
                Grid1.Cell(Det_NumPianoR1(i1), Det_NumCourbes(i2)).Text = b ' Det_NumPianoR1 : fournit le N° Ligne / Det_NumCourbes : fournit le N° de colonnes
            Next
        Next
        '
        ' Mise à jour des MUTES des Automations PIANOROLL
        ' ***********************************************
        For i1 = 0 To Form1.listPIANOROLL.Count - 1
            b = "0"
            If Form1.listPIANOROLL.Item(i1).AffCtrls.Checked Then
                b = "1"
            End If

            Grid1.Cell(Det_NumPianoR1(i1), ColCanal).Text = b
        Next
        '
    End Sub
    ''' <summary>
    ''' Grid1_ButtonClick : clic sur bouton de la grille
    ''' </summary>
    ''' <param name="Sender"></param>
    ''' <param name="e">Evènement de flexcell, contient entre autre le n° de la colonne contenant le bouton sur lequel on vient de cliquer</param>
    Private Sub Grid1_ButtonClick(ByVal Sender As System.Object, ByVal e As FlexCell.Grid.ButtonClickEventArgs) Handles Grid1.ButtonClick
        Select Case True
            Case e.Row >= LigHyperArp1 And e.Row <= LigHyperArp6  ' lignes HyperArp
                If AuMoins1Zéro(e.Row) Then
                    MettreLigneA1(e.Row)
                Else
                    MettreLigneA0(e.Row)
                End If

            Case e.Row >= LigPianoRoll1 And e.Row <= LigPianoRoll3 ' lignes PianoRoll
                If AuMoins1Zéro(e.Row) Then
                    MettreLigneA1(e.Row)
                Else
                    MettreLigneA0(e.Row)
                End If

        End Select

    End Sub
    Private Sub Grid2_ButtonClick(ByVal Sender As System.Object, ByVal e As FlexCell.Grid.ButtonClickEventArgs)

    End Sub
    ''' <summary>
    ''' AuMoins1Zéro : au moins 1 zéro sur la ligne dont le N° est passé en paramètre
    ''' </summary>
    ''' <param name="i">N) de la Ligne concernée</param>
    ''' <returns></returns>
    Function AuMoins1Zéro(i As Integer) As Boolean
        Dim j As Integer
        Dim b As Boolean = False

        For j = ColExp To ColCC53
            If Trim(Grid1.Cell(i, j).Text) = "0" Or Trim(Grid1.Cell(i, j).Text) = "" Then
                b = True
                Exit For
            End If
        Next j
        Return b
    End Function
    Sub MettreLigneA1(i As Integer)
        Dim j As Integer
        For j = ColExp To ColCC53
            Grid1.Cell(i, j).Text = "1"
        Next
    End Sub
    Sub MettreLigneA0(i As Integer)
        Dim j As Integer
        For j = ColExp To ColCC53
            Grid1.Cell(i, j).Text = "0"
        Next
    End Sub
    Function Det_NumCourbes(Courbe As Integer) As Integer
        Dim i As Integer = 0
        Select Case Courbe
            Case 0
                i = ColExp
            Case 1
                i = ColMod
            Case 2
                i = ColPan
            Case 3
                i = ColCC50
            Case 4
                i = ColCC51
            Case 5
                i = ColCC52
            Case 6
                i = ColCC53
        End Select
        Return i
    End Function

    Function Det_NumHypArp1(canal As Integer) As Integer
        Dim i As Integer = 0
        Select Case canal
            Case 0
                i = LigHyperArp1
            Case 1
                i = LigHyperArp2
            Case 2
                i = LigHyperArp3
            Case 3
                i = LigHyperArp4
            Case 4
                i = LigHyperArp5
            Case 5
                i = LigHyperArp6
        End Select
        Return i
    End Function
    Function Det_NumHypArp2(ligne As Integer) As Integer
        Dim i As Integer = 0
        Select Case ligne
            Case LigHyperArp1
                i = 0
            Case LigHyperArp2
                i = 1
            Case LigHyperArp3
                i = 2
            Case LigHyperArp4
                i = 3
            Case LigHyperArp5
                i = 4
            Case LigHyperArp6
                i = 5
        End Select
        Return i
    End Function
    Function Det_NumPianoR1(NCourbe As Integer) As Integer
        Dim i As Integer = 0
        Select Case NCourbe
            Case 0
                i = LigPianoRoll1
            Case 1
                i = LigPianoRoll2
            Case 2
                i = LigPianoRoll3
            Case 3
                i = LigPianoRoll4
            Case 4
                i = LigPianoRoll5
        End Select
        Return i
    End Function
    Function Det_NumPianoR2(ligne As Integer) As Integer
        Dim i As Integer = 0
        Select Case ligne
            Case LigPianoRoll1
                i = 0
            Case LigPianoRoll2
                i = 1
            Case LigPianoRoll3
                i = 2
            Case LigPianoRoll4
                i = 3
            Case LigPianoRoll5
                i = 4
        End Select
        Return i
    End Function

    ''' <summary>
    ''' Button2_Click : Bouton Annuler
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Close()
    End Sub
    ''' <summary>
    ''' Button1_Click : Bouton OK
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim i1, i2 As Integer
        Dim b As Boolean
        '
        ' Mise à jour Automation des CTRL d'Automation HyperArp
        ' *****************************************************
        For i1 = 0 To Form1.Automation1.ListCCAct.Count - 1
            For i2 = 0 To Form1.Automation1.ListCCAct.Item(i1).LCCAct.Count - 1
                b = False
                If Grid1.Cell(Det_NumHypArp1(i1), Det_NumCourbes(i2)).Text = "1" Then
                    b = True
                End If
                Form1.Automation1.ListCCAct.Item(i1).LCCAct(i2).Checked = b ' Det_NumHypArp1 : fournit le N° Ligne / Det_NumCourbes : fournit le N° de colonnes
                'Form1.Automation1.Canaux.Item(i).Mute.Checked

            Next
        Next
        '
        ' Mise à jour des MUTES des Automations HyperArp par canaux
        ' *********************************************************
        For i1 = 0 To Form1.Automation1.Canaux.Count - 1
            i2 = Det_NumHypArp1(i1)
            Grid1.Cell(i2, ColCanal).Locked = False
            b = False
            If Grid1.Cell(i2, ColCanal).Text = "1" Then
                b = True
            End If
            Form1.Automation1.Canaux.Item(i1).Mute.Checked = b ' 

        Next
        '
        ' Mise à jour des CTRM d'Automation des PianoRoll
        ' ***********************************************
        For i1 = 0 To Form1.listPIANOROLL.Count - 1
            For i2 = 0 To Form1.listPIANOROLL.Item(i1).CCActif.Count - 1
                b = False
                If Grid1.Cell(Det_NumPianoR1(i1), Det_NumCourbes(i2)).Text = "1" Then
                    b = True
                End If
                Form1.listPIANOROLL.Item(i1).CCActif.Item(i2).Checked = b ' Det_NumPianoR1 : fournit le N° Ligne / Det_NumCourbes : fournit le N° de colonnes
            Next
        Next
        ' 
        ' Mise à jour des MUTES des Automations PIANOROLL
        ' ***********************************************
        For i1 = 0 To Form1.listPIANOROLL.Count - 1
            b = False
            If Grid1.Cell(Det_NumPianoR1(i1), ColCanal).Text = "1" Then
                b = True
            End If
            Form1.listPIANOROLL.Item(i1).AffCtrls.Checked = b
        Next
        Close()
    End Sub

    ''' <summary>
    ''' Button3_Click : Effacer tout : mettre tout à false
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        For i = LigHyperArp1 To LigHyperArp6
            MettreLigneA0(i)
        Next
        For i = LigPianoRoll1 To LigPianoRoll3
            MettreLigneA0(i)
        Next
    End Sub

    Private Sub Grid1_Load(sender As Object, e As EventArgs) Handles Grid1.Load

    End Sub
End Class