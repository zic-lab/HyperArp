Public Class MotifsForm
    Private Retour_ As String
    Public Property retour() As String
        Get
            Return Retour_
        End Get
        Set(ByVal value As String)
            Retour_ = value
        End Set
    End Property
    ''' <summary>
    ''' OK : Click sur Bouton OK
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Retour_ = RadioBCoché()
        Me.Hide()

    End Sub
    ''' <summary>
    ''' Annuler : Click sur bouton annuler
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Retour_ = ""
        Me.Hide()
    End Sub
    Private Function RadioBCoché() As String
        Dim a As String = ""
        For Each R As RadioButton In Panel1.Controls
            If R.Checked Then
                a = R.Text
                Exit For
            End If
        Next
        Return Trim(a)
    End Function

    Private Sub MotifsForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Module1.LangueIHM = "fr" Then
            Label1.Text = "Arpèges"
            Label2.Text = "Arpèges Accords"
            Label3.Text = "Perso"
            Label4.Text = "Voix"
            Label5.Text = "Accords"
            '
            Label6.Text = "ARPEGES"
            Label7.Text = "REPETITIONS"
        Else
            Label1.Text = "Arpeggios"
            Label2.Text = "Chords Arpeggios"
            Label3.Text = "Perso"
            Label4.Text = "Voices"
            Label5.Text = "Chords"
            '
            Label6.Text = "ARPEGGIOS"
            Label7.Text = "REPETITIONS"
        End If
        Dim a As String = RadioButton1.Text

        For Each rad As RadioButton In Panel1.Controls
            If Trim(rad.Text) = Trim(MotifCourant) Then
                rad.Checked = True
                Exit For
            End If
        Next
    End Sub
End Class