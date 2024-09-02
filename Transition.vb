Public Class Transition



    Public RésultatTrans As String

    Private Chart2 As New DataVisualization.Charting.Chart
    Private ChartArea1 As New DataVisualization.Charting.ChartArea
    Private Series1 As New DataVisualization.Charting.Series
    Private Générer As New Button
    Private Annuler As New Button
    Private Effacer As New Button
    Private AffVal As New TextBox
    Private Enchargement As Boolean = True


    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        '
        If ENchargement Then
            COnstructionGraphe()
            Enchargement = False
        End If
        Maj_Série()
    End Sub
    Function CalculSérie() As String
        Dim i, j As Integer
        Dim a, b As Double
        Dim Result As String = ""

        ' Point A = (0,Mypoint1)
        ' Point B = (Interval, MyPoint2)

        ' Calcul a= Coefficient directeur Delta Y / Delta X
        i = Form1.MyPoint2 - Form1.MyPoint1 ' delta Y
        j = Form1.MyInterval - 0        ' delta X
        a = i / j '
        ' Calcul de b : b=y-ax
        b = Form1.MyPoint2 - a * Form1.MyInterval
        '
        '
        ' Détermination des valeurs entre MyPoint1 et MyPoint2
        For i = 0 To Form1.MyInterval
            y = Fix((a * i) + b)
            Result = Result + y.ToString + " "
        Next
        Return Trim(Result)
    End Function
    Sub ConstructionGraphe()
        Dim tbl() As String
        Dim i As Integer
        Series1.ChartType = DataVisualization.Charting.SeriesChartType.Line
        Chart2.ChartAreas.Add(ChartArea1)
        Chart2.BackColor = Color.Beige


        Series1.ChartArea = "ChartArea1"
        Series1.BorderWidth = 5
        ' création de 2 points
        RésultatTrans = CalculSérie()
        tbl = RésultatTrans.Split()
        For i = 0 To tbl.Count - 1
            Series1.Points.Add(Convert.ToInt16(tbl(i)))
        Next
        'Series1.Points.Add(10)
        'Series1.Points.Add(50)
        ' ajouter la serie au chart

        ' Position le chart
        Chart2.Location = New System.Drawing.Point(0, 0)
        ' Dimension du chart
        Chart2.Size = New System.Drawing.Size(220, 130)
        ' 
        ' Affichage du résultat
        AffVal.Location = New Point(1, 130)
        AffVal.Size = New Size(211, 15)
        AffVal.Text = ""
        AffVal.Visible = True
        Me.Controls.Add(AffVal)

        ' Boutons des actions
        Générer.Location = New Point(20, 153)
        Générer.Size = New Size(65, 30)
        If Module1.LangueIHM = "fr" Then
            Générer.Text = "Générer"
        Else
            Générer.Text = "Generate"
        End If
        Générer.Visible = True
        Me.Controls.Add(Générer)
        '
        Effacer.Location = New Point(90, 153)
        Effacer.Size = New Size(65, 30)
        If Module1.LangueIHM = "fr" Then
            Effacer.Text = "Effacer"
        Else
            Effacer.Text = "Delete"
        End If
        Effacer.Visible = True
        Me.Controls.Add(Effacer)
        '
        Annuler.Location = New Point(165, 153)
        Annuler.Size = New Size(65, 30)
        If Module1.LangueIHM = "fr" Then
            Annuler.Text = "Annuler"
        Else
            Annuler.Text = "Cancel"
        End If
        Annuler.Visible = True
        Me.Controls.Add(Annuler)
        ' Ajouter le chart au formulaire
        Me.Controls.Add(Chart2)
        Me.Height = 225
        Me.Width = 250
        Me.Text = "Transition"
        '
        AddHandler Générer.Click, AddressOf Générer_Click
        AddHandler Effacer.Click, AddressOf Effacer_Click
        AddHandler Annuler.Click, AddressOf Annuler_Click
    End Sub
    Sub Maj_Série()
        Dim i As Integer
        Dim tbl() As String
        ' création de 2 points
        Series1.Points.Clear()
        RésultatTrans = CalculSérie()
        AffVal.Text = Trim(RésultatTrans)
        If Trim(RésultatTrans) <> "" Then
            'Series1.Points.Add(Form1.MyPoint1)
            tbl = RésultatTrans.Split()
            For i = 0 To tbl.Count - 1
                Series1.Points.Add(Convert.ToInt16(tbl(i)))
            Next
            'Series1.Points.Add(Form1.MyPoint2)
        End If
        Chart2.Series.Clear()
        Chart2.Series.Add(Series1)
    End Sub
    Sub Générer_Click(sender As Object, e As EventArgs)
        Me.Hide()
    End Sub
    Sub Annuler_Click(sender As Object, e As EventArgs)
        RésultatTrans = ""
        Me.Hide()
    End Sub
    Sub Effacer_Click(sender As Object, e As EventArgs)
        RésultatTrans = "Effacer"
        Me.Hide()
    End Sub
End Class