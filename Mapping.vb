Public Class Mapping
    Private Sub Mapping_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim i, j As Integer


        Me.Text = "Mapping PRO"
        Grid1.Cols = 50
        Grid1.Rows = 32
        Grid1.Column(0).Width = 170
        Grid1.Column(0).Alignment = AlignmentEnum.LeftCenter
        For i = 1 To Grid1.Cols - 1
            Grid1.Cell(0, i).Text = Convert.ToString(i)
        Next
        '
        Grid1.Cell(1, 0).Text = "Position"
        Grid1.Cell(2, 0).Text = "Marqueur - Marker"
        Grid1.Cell(3, 0).Text = "N° Accord - Chord number"
        Grid1.Cell(4, 0).Text = "Accord - Chord"
        Grid1.Cell(5, 0).Text = "Tonalité - Tonality"
        Grid1.Cell(6, 0).Text = "Mode"
        Grid1.Cell(7, 0).Text = "Gamme - Scale"
        Grid1.Cell(8, 0).Text = "Degré - Degree"
        Grid1.Cell(9, 0).Text = "Répétition"
        Grid1.Cell(10, 0).Text = "Variation"
        Grid1.Cell(11, 0).Text = "Ligne"
        Grid1.Cell(12, 0).Text = "Détails"
        Grid1.Cell(13, 0).Text = "Zone de Voicing - Voicing Zone"
        Grid1.Cell(14, 0).Text = "Racine de Voicing - Voicing Root"

        For j = 1 To nbMesures
            If TableEventH(j, 1, 1).Ligne <> -1 Then
                Grid1.Cell(1, j).Text = TableEventH(j, 1, 1).Position
                Grid1.Cell(2, j).Text = TableEventH(j, 1, 1).Marqueur
                Grid1.Cell(3, j).Text = TableEventH(j, 1, 1).NumAcc
                Grid1.Cell(4, j).Text = TableEventH(j, 1, 1).Accord
                Grid1.Cell(5, j).Text = TableEventH(j, 1, 1).Tonalité
                Grid1.Cell(6, j).Text = TableEventH(j, 1, 1).Mode
                Grid1.Cell(7, j).Text = TableEventH(j, 1, 1).Gamme
                Grid1.Cell(8, j).Text = TableEventH(j, 1, 1).Degré
                Grid1.Cell(9, j).Text = TableEventH(j, 1, 1).Répet
                Grid1.Cell(10, j).Text = TableEventH(j, 1, 1).NumMagnéto
                Grid1.Cell(11, j).Text = TableEventH(j, 1, 1).Ligne
                Grid1.Cell(12, j).Text = TableEventH(j, 1, 1).Détails
            End If
        Next
    End Sub
End Class