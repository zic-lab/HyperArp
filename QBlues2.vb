Imports System.Windows.Controls
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class QBlues2
    Dim ligne As Integer = 0
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Hide()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub QBlues2_Load(sender As Object, e As EventArgs) Handles Me.Load
        ComboBox1.Items.Clear()

        ComboBox1.Items.Add("C")
        ComboBox1.Items.Add("C#")
        ComboBox1.Items.Add("D")
        ComboBox1.Items.Add("D#")
        ComboBox1.Items.Add("E")
        ComboBox1.Items.Add("F")
        ComboBox1.Items.Add("F#")
        ComboBox1.Items.Add("G")
        ComboBox1.Items.Add("G#")
        ComboBox1.Items.Add("A")
        ComboBox1.Items.Add("A#")
        ComboBox1.Items.Add("B")
        '
        ComboBox1.SelectedIndex = 0
        '
        majBlues("C", "Majeur")
    End Sub
    Sub majBlues(Tonique As String, Mode As String)
        Maj_TabNotes_Majus()
        Tonique = Trim(Tonique)
        Dim sDomin As String = TabNotes(Array.IndexOf(TabNotes, Tonique) + 5)
        Dim Domin As String = TabNotes(Array.IndexOf(TabNotes, Tonique) + 7)
        '
        Select Case Trim(Mode)
            Case "Majeur"
                Label1.Text = Tonique + " 7"
                Label2.Text = sDomin + " 7"
                Label3.Text = Tonique + " 7"
                Label4.Text = Domin + " 7"
                Label5.Text = sDomin + " 7"
                Label6.Text = Tonique + " 7"
                Label7.Text = Domin + " 7"
            Case "Mineur"
                Label1.Text = Tonique + " m7"
                Label2.Text = sDomin + " m7"
                Label3.Text = Tonique + " m7"
                Label4.Text = Domin + " 7"
                Label5.Text = sDomin + " m7"
                Label6.Text = Tonique + " m7"
                Label7.Text = Domin + " 7"
        End Select
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            majBlues(Trim(ComboBox1.Text), "Majeur")
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            majBlues(Trim(ComboBox1.Text), "Mineur")
        End If
    End Sub

    Sub ECR_Standard(Accord As String, Degré As String, Position As String)

        Dim tbl() As String = Accord.Split()
        Maj_TabNotes_Majus()
        Dim i As Integer = Array.IndexOf(TabNotes, Trim(tbl(0))) ' index de la tonique de l'accord
        Dim Tona As String = Trim(TabNotes(i + 5)) + " Maj"
        Dim Mode As String = Tona
        Dim Gamme As String = Trim(tbl(0)) + " Blues"

        ECR_Acc(Accord, Degré, Position, Tona, Mode, Gamme)
        '
        Me.Hide()

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
        Form1.Maj_CopieAcc(m)
        '
        If m = 1 Then
            Form1.Grid2.Cell(2, m).BackColor = Color.Red
            Form1.Grid2.Cell(2, m).ForeColor = Color.Yellow
            Form1.Grid2.Cell(11, m).BackColor = Color.Red
            Form1.Grid2.Cell(11, m).ForeColor = Color.Yellow
        End If

        ligne = Ligne + 1
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim I As String = Trim(Label1.Text)
        Dim IV As String = Trim(Label2.Text)
        Dim V As String = Trim(Label7.Text)

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
    End Sub
End Class