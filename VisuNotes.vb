Public Class VisuNotes
    Dim téléchargement As Boolean = True
    Private Sub Debug_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim i, j As Integer

        DbgGrid.AutoRedraw = False
        '
        Dim p As Size
        p.Height = 260
        p.Width = Me.Width
        Me.Size = p
        '
        DbgGrid.Rows = nb_TotalPistes + 1 'Arrangement1.Nb_PistesMidi + 2
        DbgGrid.Cols = nbMesures * 16 ' Arrangement1.nbMesures * 16
        '
        'DbgGrid.Range(0, 0, DbgGrid.Rows - 1, DbgGrid.Cols - 1).ClearText()


        '
        j = 0
        For i = 1 To DbgGrid.Cols - 1
            DbgGrid.Column(i).Width = 25
        Next i
        '
        j = 0
        For i = 1 To DbgGrid.Cols - 1 Step 16
            j = j + 1
            DbgGrid.Cell(0, i).Text = Str(j)
        Next i
        '
        j = 0
        For i = 1 To DbgGrid.Cols - 1 Step 4
            For k = 1 To nb_TotalPistes 'Arrangement1.Nb_PistesMidi + 1
                DbgGrid.Cell(k, i).BackColor = Color.Gainsboro
            Next k
        Next i
        DbgGrid.Column(0).Width = 65
        '
        For i = 1 To nb_TotalPistes 'Arrangement1.Nb_PistesMidi + 1
            DbgGrid.Cell(i, 0).Alignment = AlignmentEnum.LeftCenter
            If Module1.LangueIHM = "fr" Then
                DbgGrid.Cell(i, 0).Text = "Canal" + Str(i)
                Me.Text = "Visu Notes PRO"
            Else
                DbgGrid.Cell(i, 0).Text = "Channel" + Str(i)
                Me.Text = "Notes View PRO"
            End If
        Next i
        ' Raz du texte
        For i = 1 To DbgGrid.Rows - 1
            For j = 1 To DbgGrid.Cols - 1
                DbgGrid.Cell(i, j).Text = ""
            Next
        Next
        ' Affichage notes
        Dbg_AffPar("Note")
        '
        DbgGrid.AutoRedraw = True
        DbgGrid.Refresh()
        '
        RadioButton1.Checked = True
        téléchargement = False

        ' Maj des locators utilisés

        Label1.Text = Form1.Début.Text
        Label2.Text = Form1.Terme.Text

    End Sub
    Sub Dbg_AffPar(par As String)
        Dim i, j, k, pst As Integer

        DbgGrid.Range(1, 1, DbgGrid.Rows - 1, DbgGrid.Cols - 1).ClearText()
        For pst = 0 To nb_TotalPistes - 1 'Arrangement1.Nb_PistesMidi '
            For k = 0 To LesPistes(pst).DbgTabNotes.Count - 1
                i = Val(LesPistes(pst).DbgTabNotes.Item(k).numPiste) + 1
                j = (Val(LesPistes(pst).DbgTabNotes.Item(k).position) + 1) + (LesPistes.Item(i).Start * 16)
                Select Case par
                    Case "Note"
                        DbgGrid.Cell(i, j).Text = Trim(LesPistes(pst).DbgTabNotes.Item(k).note)
                        If pst >= nb_PistesVar Then
                            DbgGrid.Cell(i, j).ForeColor = Color.Red
                        End If
                    Case "Dyn"
                        DbgGrid.Cell(i, j).Text = Trim(LesPistes(pst).DbgTabNotes.Item(k).dyn)
                    Case "Durée"
                        DbgGrid.Cell(i, j).Text = Trim(LesPistes(pst).DbgTabNotes.Item(k).durée)
                    Case "Canal"
                        DbgGrid.Cell(i, j).Text = Trim(LesPistes(pst).DbgTabNotes.Item(k).canal)
                End Select
            Next k
        Next pst
    End Sub

    Private Sub DbgGrid_Load(sender As Object, e As EventArgs)

    End Sub


    Private Sub SplitContainer1_Panel2_Paint(sender As Object, e As PaintEventArgs) Handles SplitContainer1.Panel2.Paint

    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If téléchargement = False Then
            Dbg_AffPar("Dyn")
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If téléchargement = False Then
            Dbg_AffPar("Note")
        End If
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        If téléchargement = False Then
            Dbg_AffPar("Durée")
        End If
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub
End Class