Public Class F2
    Private Sub Grid1_Load(sender As Object, e As EventArgs) Handles Grid1.Load
        Dim i As Integer
        For i = 1 To Grid1.Cols - 1
            Grid1.Cell(0, i).FontBold = True
            Grid1.Cell(0, i).Text = Convert.ToString(i)
            Grid1.Column(i).Width = 20
        Next i
    End Sub


    Private Sub Grid1_Scroll(Sender As Object, e As EventArgs) Handles Grid1.Scroll
        F1.Grid1.LeftCol = Me.Grid1.LeftCol
    End Sub

    Private Sub Grid1_MouseUp(Sender As Object, e As MouseEventArgs) Handles Grid1.MouseUp
        Dim i As Integer = Grid1.Selection.FirstRow
        Dim j As Integer = Grid1.Selection.FirstCol
        '
        Dim ii As Integer = Grid1.Selection.LastRow
        Dim jj As Integer = Grid1.Selection.LastCol
        '
        If Grid1.Cell(i, j).BackColor = Color.White Then
            Grid1.Range(i, j, ii, jj).BackColor = Color.Blue
        Else
            Grid1.Range(i, j, ii, jj).BackColor = Color.White
        End If
    End Sub
    Private Sub F2_Move(sender As Object, e As EventArgs) Handles Me.Move
        F1.Location = Me.Location
    End Sub
    Private Sub F2_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        F1.Close()
    End Sub
    Private Sub F2_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        F1.Size = Me.Size
    End Sub
End Class