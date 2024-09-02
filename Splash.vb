Public Class Splash
    Private Sub Panel1_MouseUp(sender As Object, e As MouseEventArgs)
        FermetureSplash()
    End Sub

    Private Sub FermetureSplash()
        Me.Visible = False
        Me.Close()
        'Me.Dispose()

        Timer1.Stop() ' timer de durée de la splash image
        Form1.Visible = True
        '
        If Exist_MIDIout = False Then
            If Langue = "fr" Then
                Avertis = "Pas de sortie MIDI : l'écoute ne pourra pas fonctionner."
            Else
                Avertis = "No MIDI out : listening will not work."
            End If
            Cacher_FormTransparents()
            MessageHV.ShowDialog()
            End
        End If
    End Sub
    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs)
        ' Specify that the link was visited.
        'Me.LinkLabel1.LinkVisited = True
        ' Navigate to a URL.
        System.Diagnostics.Process.Start("https://sourceforge.net/p/hypervoicing/wiki/Home/")
    End Sub
    Private Sub Label2_MouseUp(sender As Object, e As MouseEventArgs) Handles Label2.MouseUp
        Me.Close()
        'Me.Dispose()
        Form1.Visible = True
    End Sub
    Private Sub Label3_MouseUp(sender As Object, e As MouseEventArgs)
        Me.Close()
        'Me.Dispose()
        Form1.Visible = True
    End Sub
    Private Sub Label1_MouseUp(sender As Object, e As MouseEventArgs)
        Me.Close()
        'Me.Dispose()
        Form1.Visible = True
    End Sub
    Private Sub Splash_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim a As System.DateTime
        Dim b As String
        Dim c As String
        '
        a = System.DateTime.Today
        b = a.ToString("dd/MM/yy")
        c = a.ToString("yyyy")

        Label2.Text = "278 79"
        Label4.Text = "Lab13"
        Label5.Text = "VersionDev :" + NumVersionDev
        If LangueIHM = "fr" Then
            Label1.Text = "Tous droits réservés- - HyperArp - MIDI Model - " + Trim(c)
            Label3.Text = "Version : " + NumVersion ' + "  -  " + Dateversion
        Else
            Label1.Text = "All rights reserved - HyperArp - MIDI Model - " + Trim(c)
            Label3.Text = "Version : " + NumVersion '+ "  -  " + Dateversion
        End If
        ' Timer de durée d'affichage de la splah image
        ' ********************************************
        Timer1.Interval = 10000
        Timer1.Start()


    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        FermetureSplash()
    End Sub
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        FermetureSplash()
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub
End Class