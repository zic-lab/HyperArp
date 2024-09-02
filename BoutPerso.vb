Imports System.Drawing.Drawing2D
Public Class BoutPerso : Inherits Button

    Private _cliqué As Boolean
    Private _couleur1 As Color
    Private _couleur2 As Color
    Private _couleur3 As Color
    Private _couleurFond As Color

    Public Property BtnCouleur1() As Color
        Get
            Return _couleur1
        End Get
        Set(value As Color)
            _couleur1 = value
            Refresh()
        End Set
    End Property

    Public Property BtnCouleur2() As Color
        Get
            Return _couleur2
        End Get
        Set(value As Color)
            _couleur2 = value
            t = _couleur2
            Refresh()
        End Set
    End Property

    Public Property BtnMouseHover() As Color
        Get
            Return _couleur3
        End Get
        Set(value As Color)
            _couleur3 = value
            Refresh()
        End Set
    End Property

    Public Property BtnCouleurFond() As Color
        Get
            Return _couleurFond
        End Get
        Set(value As Color)
            _couleurFond = value
            Refresh()
        End Set
    End Property

    Dim t As Color

    Public Sub New()

        _cliqué = False
        _couleur1 = Color.White ' couleur du haut du bouton
        _couleur2 = Color.White ' 
        _couleur3 = Color.AntiqueWhite
        _couleurFond = Color.FromArgb(60, 60, 60)

        t = _couleur2

        Me.Size = New Size(200, 30)
        Me.ForeColor = Color.Black
        Cursor = Cursors.Hand

    End Sub

    Protected Overrides Sub OnPaint(pevent As PaintEventArgs)
        MyBase.OnPaint(pevent)

        Dim G As Graphics = pevent.Graphics ' déclaration d'une variable 'graphique'
        G.Clear(_couleurFond) ' effacer couleur par defaut
        Dim R As New Rectangle(1, 1, Width - 2, Height - 2) ' le rectanle sera tracé sur le bouton 
        If _cliqué = False Then
            Dim GB As New LinearGradientBrush(R, _couleur1, _couleur2, 90) ' donne l'effet de 2 couleurs mélangées (90 est l'angle)
            G.FillRectangle(GB, R) ' remplir le rectangle
        Else
            Dim GB As New LinearGradientBrush(R, _couleur2, _couleur1, 90) ' si le bouton n'est pas cliqué on inverse les couleurs.
            G.FillRectangle(GB, R) ' remplir le rectangle
        End If

        Dim m As SizeF = G.MeasureString(Text, Font)  ' cacul de la taille du texte (en fonction du font)
        Dim x As Single = (Width / 2) - m.Width / 2   ' calcul coordonnée x du bouton
        Dim y As Single = (Height / 2) - m.Height / 2 ' calcul coordonnée y du bouton
        Dim pf As New PointF(x, y) ' point de position du texte
        Dim pf2 As New PointF(x + 1, y + 1)

        G.DrawString(Text, Font, Brushes.White, pf)
        G.DrawString(Text, Font, Brushes.Black, pf2) ' dessinner le texte Brush.black = couleur du texte

    End Sub

    ' changement de couleur quend la souris survol le bouton
    ' ******************************************************
    Protected Overrides Sub OnMouseHover(e As EventArgs)
        MyBase.OnMouseHover(e)
        t = _couleur2
        _couleur2 = Color.DarkKhaki
        Refresh()
    End Sub
    ' évènement quand la souris quitte le bouton
    ' ******************************************
    Protected Overrides Sub OnMouseLeave(e As EventArgs) ' 
        MyBase.OnMouseLeave(e)
        _couleur2 = t
        Refresh()
    End Sub

    ' Clic sur le bouton
    ' ******************
    Protected Overrides Sub OnMouseDown(mevent As MouseEventArgs)
        MyBase.OnMouseDown(mevent)
        _cliqué = True
        Refresh()
    End Sub
    ' Relachement  de la souris après clic sur le bouton
    ' **************************************************
    Protected Overrides Sub OnMouseUp(mevent As MouseEventArgs)
        MyBase.OnMouseUp(mevent)
        _cliqué = False
        Refresh()
    End Sub
End Class
