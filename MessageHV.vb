Public Class MessageHV
    Public Titre As String
    Public ContenuMess As String
    Public TypBouton As String
    Public Sortie As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Oui.Click
        Sortie = "Oui"
        Me.Hide()
    End Sub

    Private Sub MessageHV_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = Titre
        TextBox1.Text = ContenuMess
        Me.TopMost = True
        Oui.Text = "OK"
        Oui.Visible = True

        If LangueIHM = "fr" Then
            Select Case TypBouton
                Case "OK", "Oui"
                    Non.Visible = False
                Case "OuiNon"
                    Non.Visible = True
                    Non.Text = "Non"
            End Select
        Else
            Select Case TypBouton
                Case "OK"
                    Non.Visible = False
                Case "OuiNon"
                    Non.Visible = True
                    Non.Text = "No"
            End Select
        End If
    End Sub

    Private Sub Non_Click(sender As Object, e As EventArgs) Handles Non.Click
        Sortie = "Non"
        Me.Hide()
    End Sub
    Public Property PTitre() As String
        Get
            Dim a As String = Titre
            Return a
        End Get
        Set(ByVal Value As String)
            Me.Titre = Value
        End Set
    End Property
    Public Property PContenuMess() As String
        Get
            Dim a As String = ContenuMess
            Return a
        End Get
        Set(ByVal Value As String)
            Me.ContenuMess = Value
        End Set
    End Property
    Public Property PTypBouton() As String
        Get
            Dim a As String = TypBouton
            Return a
        End Get
        Set(ByVal Value As String)
            Me.TypBouton = Value
        End Set
    End Property
    Public Property PSortie() As String
        Get
            Dim a As String = Sortie
            Return a
        End Get
        Set(ByVal Value As String)
            Me.Sortie = Value
        End Set
    End Property
End Class