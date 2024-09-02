Public Class VéloAleat
    Private Sub VéloAleat_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim P As New Point
        P.X = Posx
        P.Y = Posy
        Me.Location = P
        '
        ValeurMin.Value = ValMn
        ValeurMax.Value = ValMx
        '
        If LangueIHM = "fr" Then
            Me.Text = "Vélocités aléatoires"
        Else
            Me.Text = "Random velocities"
        End If
        '
        VéloAléat_Chargé = True
    End Sub

    ' Retour par OK
    ' *************
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Hide()
    End Sub
    ' Retour par Annuler
    ' ******************
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ValMn = -1
        Me.Hide()
    End Sub

    Private Sub ValeurMin_ValueChanged(sender As Object, e As EventArgs) Handles ValeurMin.ValueChanged
        If ValeurMin.Value > ValeurMax.Value Then
            ValeurMin.Value = ValeurMax.Value
        End If
    End Sub

    Private Sub ValeurMax_ValueChanged(sender As Object, e As EventArgs) Handles ValeurMax.ValueChanged
        If ValeurMin.Value > ValeurMax.Value Then
            ValeurMax.Value = ValeurMin.Value
        End If
    End Sub
End Class