Imports System.Windows.Forms.Control
Module Extensions

    ' Cette méthode permet de mettre toutes les propriétés 'TabStop = false' des contrôles créés avec la Barre à Outils VB
    ' Remarque : les propriétés 'TabStop' des controles créés dans le code (sans la Barre à Outils) prennent la valeur 'False' automatiquement au moment de la création de leur contrôle.
    <System.Runtime.CompilerServices.Extension()>
    Public Sub setAllControlsTabstop(ByVal instance As ControlCollection, ByVal value As Boolean)
        For Each ctrl As System.Windows.Forms.Control In instance
            If ctrl.TabIndex <> Nothing Then
                ctrl.TabStop = value
            End If
            ctrl.Controls.setAllControlsTabstop(value)
        Next
    End Sub
End Module
