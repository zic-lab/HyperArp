Public NotInheritable Class AboutBox1

    Private Sub AboutBox1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim a As System.DateTime
        Dim b As String
        Dim c As String
        '
        a = System.DateTime.Today
        b = a.ToString("dd/MM/yy")
        c = Dateversion 'a.ToString("yyyy")
        '
        If Langue = "fr" Then
            Label3.Text = "Version : " + NumVersion
        Else
            Label3.Text = "Version :" + NumVersion
        End If
        Label4.Text = Trim(Dateversion)
        If Langue = "fr" Then
            Label6.Text = "Tous droits réservés - HyperArp - Midi Model " + Trim(c)
            Label5.Text = "Arpégiateur"
        Else
            Label6.Text = "All rights reserved - HyperArp - Midi Model " + Trim(c)
            Label5.Text = "Arpegiator"
        End If
        ' Définissez le titre du formulaire.
        Dim ApplicationTitle As String
        If My.Application.Info.Title <> "" Then
            ApplicationTitle = My.Application.Info.Title
        Else
            ApplicationTitle = System.IO.Path.GetFileNameWithoutExtension(My.Application.Info.AssemblyName)
        End If

        ' Initialisez tout le texte affiché dans la boîte de dialogue À propos de.
        ' TODO: personnalisez les informations d'assembly de l'application dans le volet "Application" de la 
        '    boîte de dialogue Propriétés du projet (sous le menu "Projet").
        'Me.LabelProductName.Text = My.Application.Info.ProductName
        'Me.LabelVersion.Text = String.Format("Version {0}", My.Application.Info.Version.ToString)
        'Me.LabelCopyright.Text = My.Application.Info.Copyright
        'Me.LabelCompanyName.Text = My.Application.Info.CompanyName
        'Me.TextBoxDescription.Text = My.Application.Info.Description
    End Sub

    Private Sub OKButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
End Class
