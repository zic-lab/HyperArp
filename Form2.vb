Imports System
Imports System.IO
Imports System.Text
Public Class Form2
    'Dim file As System.IO.StreamWriter
    'file = My.Computer.FileSystem.OpenTextFileWriter(NomFichier, True)
    'file.WriteLine(a)
    'file.Close()
    'Me.Visible = False
    Dim ListPresets As New List(Of String)
    Dim NomFichier As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\HyperVoicing" + "\ConstructionPréset.txt" ' le fichier est créer dans la classe DrumEdit, procédure  Private Sub SauvPrésets_MouseClick
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim tbl() As String


        Me.ListBox1.Items.Clear()
        ListPresets.Clear()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox2.ForeColor = Color.Blue
        TextBox3.Text = ""
        TextBox3.ForeColor = Color.Blue


        If My.Computer.FileSystem.FileExists(NomFichier) Then
            'ListPresets.AddRange((System.IO.File.ReadAllLines(NomFichier, System.Text.Encoding.Default)))
            Using sr As New StreamReader(NomFichier) 'Stream pour la lecture
                Dim ligne As String ' Variable contenant le texte de la ligne
                While sr.Peek() >= 0
                    ligne = sr.ReadLine
                    ListPresets.Add(Trim(ligne))
                    tbl = Trim(ligne).Split(";")
                    Me.ListBox1.Items.Add(tbl(0))
                End While
            End Using
        Else
            Dim fs As FileStream = File.Create(NomFichier)
        End If
        Me.Text = "PRESETS PERSO"
        If Module1.LangueIHM = "fr" Then
            TextBox2.Text = "Fichier de Présets :  " + NomFichier
            'Label3.Text = "Fichier de Présets :  " + NomFichier
            Label1.Text = "Liste des Presets Perso"
            Label2.Text = "Nom * "
            TextBox3.Text = "Pour créer un preset, entrer un Nom et cliquer sur Ajouter."
            Button1.Text = "Ajouter"
            Button3.Text = "Retirer"
            Button4.Text = "Annuler"
            Button5.Text = "Charger"
        Else
            TextBox2.Text = "Presets file  : " + NomFichier
            'Label3.Text = "Presets file  : " + NomFichier
            Label1.Text = "Presets Perso List"
            Label2.Text = "New preset name *"
            TextBox3.Text = "To create a preset, enter a name  and click on  Add."
            Button1.Text = "Add"
            Button3.Text = "Remove"
            Button4.Text = "Cancel"
            Button5.Text = "Load"
        End If
    End Sub
    Public Function ReadLine() As String
        Dim a As String
        a = LineInput(1) ' Lire chaque ligne

        ReadLine = Mid(a, 2, Len(a) - 2)
    End Function
    ''' <summary>
    ''' Button1_Click : Add - ajouter un preset dans  la liste. Le fichier n'est mis à 
    ''' jour que sur un clic sur le bouton "OK".
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        AjouterPreset()
    End Sub
    Sub AjouterPreset()
        Dim i As Integer
        Dim a As String
        Dim tbl() As String

        If Trim(TextBox1.Text) <> "" Then
            i = Form1.Drums.Det_NumPréset(Form1.Drums.Det_Préset)
            a = Form1.Drums.Sauv_PrésetConstruct(i, Trim(TextBox1.Text)) ' détermination de la chaine correspondant au préset DrumEditen cours - textbox1 contient le titre à donner au préset
            tbl = a.Split(";")
            If tbl.Length > 1 Then ' 
                ListPresets.Add(a) ' Mise à jour de la liste avec titre + notes + instrument séparés par des ";"
                ListBox1.Items.Add(Trim(Trim(TextBox1.Text))) ' Mise à jour de la ListBox seulement avec le titre
                '
                Form1.Drums.NomPréset.Text = Trim(Trim(TextBox1.Text)) ' Mise à jour du nouveau nom dans le drum edit
                Form1.Drums.LNomPréset.Item(Form1.Drums.Det_NumPréset(Trim(Form1.Drums.Grid1.Cell(0, 0).Text))) = Form1.Drums.NomPréset.Text ' mise à jour des noms dans chaque variation "A".."H"

                If Module1.LangueIHM = "fr" Then
                    TextBox3.Text = "Nouveau preset crée : " + tbl(0)
                Else
                    TextBox3.Text = "New created preset : " + tbl(0)
                End If
            Else ' pas de notes trouvées dans le préset en cours
                If Module1.LangueIHM = "fr" Then
                    TextBox3.Text = "Préset non ajouté : pas de note."
                Else
                    TextBox3.Text = "Preset not added, no note."
                End If
            End If
        Else
            If Module1.LangueIHM = "fr" Then
                TextBox3.Text = "Pas de nom : préset non ajouté."
            Else
                TextBox3.Text = "No name : preset not added."
            End If
        End If
    End Sub
    ''' <summary>
    ''' Button3_Click : Remove - retirer un preset de la liste. Le fichier n'est mis à 
    ''' jour que sur un clic sur le bouton "OK".
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If ListBox1.SelectedIndex <> -1 Then
            ListPresets.RemoveAt(ListBox1.SelectedIndex)
            If ListBox1.SelectedIndex <> -1 Then
                ListBox1.Items.RemoveAt(ListBox1.SelectedIndex)
            End If
            If Module1.LangueIHM = "fr" Then
                TextBox3.Text = "Preset effacé : pour le restaurer sortir par Annuler."
            Else
                TextBox3.Text = "Preset deleted : to restore it quit with Cancel."
            End If
        Else
            If Module1.LangueIHM = "fr" Then
                TextBox3.Text = "Veuillez sélectionner un preset avant de l'effacer."
            Else
                TextBox3.Text = "Please select a preset before deleting it."
            End If
        End If
    End Sub

    ''' <summary>
    ''' Button5_Click : Load - Charger un préset dans le drum edit
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If ListBox1.SelectedIndex <> -1 Then
            Dim a As String = ListPresets(ListBox1.SelectedIndex)
            Form1.Drums.ChargerPréset(a)
            'Maj_Fichier()
            'Close()
        Else
            If Module1.LangueIHM = "fr" Then
                TextBox3.Text = "Veuillez sélectionner un preset avant de le charger."
            Else
                TextBox3.Text = "Please select a preset before loading it."
            End If
        End If
    End Sub

    ''' <summary>
    ''' Maj_Fichier : mise à jour du fichier 'NomFichier" contenant tous les préset perso.
    ''' La mise à jour s'effectue quen clic sur bouton "OK"
    ''' </summary>
    Private Sub Maj_Fichier()
        Using sw As StreamWriter = New StreamWriter(NomFichier)
            For Each a As String In ListPresets
                If Trim(a) <> "" Then
                    sw.WriteLine(Trim(a))
                End If
            Next
        End Using
    End Sub
    ''' <summary>
    ''' Button2_Click : Bouton OK
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Maj_Fichier()
        Close()
    End Sub
    ''' <summary>
    ''' Button4_Click: Bouton Annuler
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Close()
    End Sub

    Private Sub TextBox1_KeyUp(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyUp
        If e.KeyCode = Keys.Enter Then
            AjouterPreset()
        End If
    End Sub


End Class