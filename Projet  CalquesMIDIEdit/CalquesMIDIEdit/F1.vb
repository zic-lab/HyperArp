
Public Class F1


    Private Sub Grid1_Load(sender As Object, e As EventArgs) Handles Grid1.Load

    End Sub

    Private Sub Grid1_MouseDown(Sender As Object, e As MouseEventArgs) Handles Grid1.MouseDown

    End Sub

    Private Sub F1_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim i As Integer

        For i = 1 To Grid1.Cols - 1
            Grid1.Cell(0, i).FontBold = True
            Grid1.Cell(0, i).Text = Convert.ToString(i)
            Grid1.Column(i).Width = 20
        Next i
        '
        F2.Show()
        F2.TopMost = True

        ' i = HyperArp.MainArp.Arrangement1.Tempo

    End Sub
End Class
