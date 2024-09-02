Public Class QuickBlues
    Dim ListNotesD As New List(Of String)
    Dim listNotesMaj As New List(Of String)
    Dim Ligne As Integer = 0
    Private Sub QuickBlues_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        GroupBox1.BackColor = Color.FromArgb(2, 129, 205)
        GroupBox1.ForeColor = Color.White
        GroupBox2.BackColor = Color.FromArgb(2, 129, 205)
        GroupBox2.ForeColor = Color.White

        Dim p As New Point
        p.X = ComboBox4.Location.X
        p.Y = Button3.Location.Y
        Button3.Location = p

        Button1.Size = Button3.Size
        Button2.Size = Button4.Size
        Button1.Location = Button3.Location
        Button2.Location = Button4.Location

        Button1.ForeColor = Color.Black
        Button2.ForeColor = Color.Black
        Button3.ForeColor = Color.Black
        Button4.ForeColor = Color.Black
        '
        Button1.BackColor = Color.White
        Button2.BackColor = Color.White
        Button3.BackColor = Color.White
        Button4.BackColor = Color.White

        Me.StartPosition = FormStartPosition.CenterScreen

        Maj_NotesMajeur()
        Maj_NotesMajeurStandard()

        ComboBox2.Items.Clear()
        If LangueIHM = "fr" Then
            ComboBox2.Items.Add("Blues Majeur")
            ComboBox2.Items.Add("Blues Mineur Harmonique")
            ComboBox2.Items.Add("Blues Mineur Mélodique")
            '
            Button2.Text = "Annuler"
            '
            GroupBox1.Text = "Blues Mineur (I m7 - IV m7 - V 7)"
        Else
            ComboBox2.Items.Add("Major Blues")
            ComboBox2.Items.Add("Harmonic Minor  Blues")
            ComboBox2.Items.Add("Melodic Minor  Blues")
            '
            Button2.Text = "Cancel"
            '
            GroupBox1.Text = "Tonal Blues"
        End If
        '
        ComboBox2.SelectedIndex = 0
        '
        ' 
        ListNotesD.Clear()

        For i = 0 To 2
            ListNotesD.Add("C")
            ListNotesD.Add("C#")
            ListNotesD.Add("D")
            ListNotesD.Add("D#")
            ListNotesD.Add("E")
            ListNotesD.Add("F")
            ListNotesD.Add("F#")
            ListNotesD.Add("G")
            ListNotesD.Add("G#")
            ListNotesD.Add("A")
            ListNotesD.Add("A#")
            ListNotesD.Add("B")
        Next i
        '
        listNotesMaj.Clear()

        listNotesMaj.Add("C#")
        listNotesMaj.Add("F#")
        listNotesMaj.Add("B")
        listNotesMaj.Add("E")
        listNotesMaj.Add("A")
        listNotesMaj.Add("D")
        listNotesMaj.Add("G")
        listNotesMaj.Add("C")
        listNotesMaj.Add("F")
        listNotesMaj.Add("Bb")
        listNotesMaj.Add("Eb")
        listNotesMaj.Add("Ab")

    End Sub
    Sub Maj_NotesMajeurStandard()
        ComboBox4.Items.Clear()

        ComboBox4.Items.Add("C#")
        ComboBox4.Items.Add("F#")
        ComboBox4.Items.Add("B")
        ComboBox4.Items.Add("E")
        ComboBox4.Items.Add("A")
        ComboBox4.Items.Add("D")
        ComboBox4.Items.Add("G")
        ComboBox4.Items.Add("C")
        ComboBox4.Items.Add("F")
        ComboBox4.Items.Add("Bb")
        ComboBox4.Items.Add("Eb")
        ComboBox4.Items.Add("Ab")
        '
        ComboBox4.SelectedIndex = 7
    End Sub
    Sub Maj_NotesMajeur()
        ComboBox1.Items.Clear()

        ComboBox1.Items.Add("C#")
        ComboBox1.Items.Add("F#")
        ComboBox1.Items.Add("B")
        ComboBox1.Items.Add("E")
        ComboBox1.Items.Add("A")
        ComboBox1.Items.Add("D")
        ComboBox1.Items.Add("G")
        ComboBox1.Items.Add("C")
        ComboBox1.Items.Add("F")
        ComboBox1.Items.Add("Bb")
        ComboBox1.Items.Add("Eb")
        ComboBox1.Items.Add("Ab")
        '
        ComboBox1.SelectedIndex = 7
    End Sub
    Sub Maj_NotesMineur()
        ComboBox1.Items.Clear()

        ComboBox1.Items.Add("A#")
        ComboBox1.Items.Add("D#")
        ComboBox1.Items.Add("G#")
        ComboBox1.Items.Add("C#")
        ComboBox1.Items.Add("F#")
        ComboBox1.Items.Add("B")
        ComboBox1.Items.Add("E")
        ComboBox1.Items.Add("A")
        ComboBox1.Items.Add("D")
        ComboBox1.Items.Add("G")
        ComboBox1.Items.Add("C")
        ComboBox1.Items.Add("F")
        '
        ComboBox1.SelectedIndex = 7
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Hide()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim ton As String
        ton = Trim(ComboBox4.Text)

        If ComboBox4.SelectedIndex > 7 Then
            ton = Form1.TradD(ton)
        End If

        Dim ind_I As Integer = ListNotesD.IndexOf(ton)



        Dim I As String = Trim(ton) + " m7"
        Dim IV As String = ListNotesD(ind_I + 5) + " m7"
        Dim V As String = ListNotesD(ind_I + 7) + " 7"
        '
        '
        Ligne = 0
        ' Ecriture I sur 1
        ECR_Standard(I, 5, 1)
        ' Ecriture IV sur 5
        ECR_Standard(IV, 5, 5)
        ' Ecriture I sur 7
        ECR_Standard(I, 5, 7)
        ' Ecriture V sur 9
        ECR_Standard(V, 5, 9)
        ' Ecriture IV sur 10
        ECR_Standard(IV, 5, 10)
        ' Ecriture I sur 11
        ECR_Standard(I, 5, 11)
        ' Ecriture V sur 12
        ECR_Standard(V, 5, 12)


        Me.Hide()
    End Sub
    Function AlterClef(Ton) As String
        Dim Result As String = "-1"

        If ComboBox2.SelectedIndex = 0 Then ' = 0 --> mode majeur
            Select Case Ton
                Case "C#", "F#", "B", "E", "A", "D", "G"
                    Result = "#"
                Case "C", "F", "Bb", "Eb", "Ab"
                    Result = "b"
            End Select
        Else
            Select Case Ton
                Case "A#", "D#", "G#", "C#", "F#", "B", "E"
                    Result = "#"
                Case "A", "D", "G", "C", "F"
                    Result = "b"
            End Select
        End If
        Return Result
    End Function
    Sub ECR_Blues(I As String, IV As String, V As String, tona As String, mode As String, gamme As String)
        Ligne = 0
        ' Position 1
        ECR_Acc(I, 1, 1, tona, mode, gamme)
        ' Position 5
        ECR_Acc(IV, 4, 5, tona, mode, gamme)
        ' Position 7
        ECR_Acc(I, 1, 7, tona, mode, gamme)
        ' Position 9
        ECR_Acc(V, 5, 9, tona, mode, gamme)
        ' Position 10
        ECR_Acc(IV, 4, 10, tona, mode, gamme)
        ' Position 11
        ECR_Acc(I, 1, 11, tona, mode, gamme)
        ' Position 12
        ECR_Acc(V, 5, 12, tona, mode, gamme)
        '
        'Form1.Maj_Nligne()
    End Sub

    Sub ECR_Acc(Acc As String, Degré As Integer, Colonne As Integer, Tona As String, Mode As String, Gamme As String)
        Dim m As Integer = Colonne
        Dim tbl() As String

        ' Maj des données de l'accord dans TableEventh
        ' ********************************************
        TableEventH(m, 1, 1).Tonalité = Tona
        TableEventH(m, 1, 1).Mode = Mode '
        TableEventH(m, 1, 1).Gamme = Gamme
        TableEventH(m, 1, 1).Accord = Trim(Acc)
        TableEventH(m, 1, 1).NumAcc = Colonne
        TableEventH(m, 1, 1).Position = Str(Colonne) + ".1" + ".1"
        TableEventH(m, 1, 1).Degré = Degré - 1
        TableEventH(m, 1, 1).Ligne = Ligne
        '
        ' Ecriture et maj couleur dans Grid2
        ' **********************************
        Form1.Grid2.Cell(2, m).Alignment = FlexCell.AlignmentEnum.CenterCenter
        Form1.Grid2.Cell(11, m).Alignment = FlexCell.AlignmentEnum.CenterCenter
        Form1.Grid2.Cell(2, m).Text = Acc
        Form1.Grid2.Cell(11, m).Text = TableEventH(m, 1, 1).Gamme
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
        Form1.Maj_Répétition()
        '
        Ligne = Ligne + 1
    End Sub
    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
        e.SuppressKeyPress = True
    End Sub

    Private Sub ComboBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox2.KeyDown
        e.SuppressKeyPress = True
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.SelectedIndex = 0 Then
            Maj_NotesMajeur()
        Else
            Maj_NotesMineur()
        End If
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim ton As String
        ton = Trim(ComboBox4.Text)

        If ComboBox4.SelectedIndex > 7 Then
            ton = Form1.TradD(ton)
        End If

        Dim ind_I As Integer = ListNotesD.IndexOf(ton)

        'Dim I As String = Trim(ComboBox4.Text) + " 7"

        Dim I As String = Trim(ton) + " 7"
        Dim IV As String = ListNotesD(ind_I + 5) + " 7"
        Dim V As String = ListNotesD(ind_I + 7) + " 7"
        '
        '
        Ligne = 0
        ' Ecriture I sur 1
        ECR_Standard(I, 5, 1)
        ' Ecriture IV sur 5
        ECR_Standard(IV, 5, 5)
        ' Ecriture I sur 7
        ECR_Standard(I, 5, 7)
        ' Ecriture V sur 9
        ECR_Standard(V, 5, 9)
        ' Ecriture IV sur 10
        ECR_Standard(IV, 5, 10)
        ' Ecriture I sur 11
        ECR_Standard(I, 5, 11)
        ' Ecriture V sur 12
        ECR_Standard(V, 5, 12)


        Me.Hide()
    End Sub
    Sub ECR_Standard(Accord As String, Degré As String, Position As String)
        Dim tbl() As String
        Dim ii As Integer
        Dim Tona, Mode, Gamme As String

        tbl = Accord.Split
        If ComboBox4.SelectedIndex > 7 Then
            tbl(0) = Form1.TradD(tbl(0))
        End If
        ii = ListNotesD.IndexOf(tbl(0))
        Tona = ListNotesD(ii + 5)
        If listNotesMaj.IndexOf(Tona) = -1 Then
            Tona = Form1.TradDB(Tona, "b")
            tbl = Accord.Split()
            tbl(0) = Form1.TradDB(tbl(0), "b")
            Accord = tbl(0) + " " + tbl(1)
        End If
        Tona = Tona + " Maj"
        Mode = Tona
        Gamme = ListNotesD(ii) + " Blues"
        ECR_Acc(Accord, Degré, Position, Tona, Mode, Gamme)

    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Hide()
    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub
End Class