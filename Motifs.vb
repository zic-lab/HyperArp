Module Motifs
    ' **************************************************************
    ' CalcMotifs : remplit un certain nom de mesures avec le motif *
    ' considéré. Applée par tous les motifs                        *
    ' olong est le nombre d'éléments remis par le motif            *
    ' **************************************************************

    Function CalcMotif(Liste1 As String, metrique As String, nbMesures As Integer, duree As Double, oLong As Integer, Début As Integer) As String
        Dim nombreNotes As Integer = NbNotes(metrique, nbMesures, duree)
        Dim i As Integer = Début
        Dim nn As Integer = 0
        Dim oList As String = String.Empty
        Dim tbl1() As String = {String.Empty}
        Dim tbl2() As String = {String.Empty}
        CalcMotif = "C3"


        ' répétition de nombreNotes fois de (ou des notes) à jour en fonction de la longueur de la note
        tbl1 = Split(Trim(Liste1))
        Do While nn < nombreNotes

            ReDim Preserve tbl2(0 To nn)
            tbl2(nn) = tbl1(i)
            i = i + 1
            nn = nn + 1
            If i = oLong Then i = 0
        Loop
        Return Join(tbl2, " ")
    End Function
    ' **************************************************************
    ' * ArpMotif 1	C	E	G	B	D		C	G	E	B	G	D  *
    ' **************************************************************

    Public Function ArpMotif1(ListNotes As String, metrique As String, nbMesures As Integer, duree As Double, Début As Integer) As String
        Dim Liste1 As String = ""
        Dim oLong As Integer ' longueur du motif d'arpège en nombre de notes
        Dim tbl1() As String = Split(ListNotes)
        ArpMotif1 = ""
        Select Case UBound(tbl1) + 1
            Case 3 ' 3 notes reçues -> 3 notes d'arpèges générées
                Dim oTab(3) As String
                oTab(0) = tbl1(0)
                oTab(1) = tbl1(2)
                oTab(2) = tbl1(1)
                Liste1 = Join(oTab, " ")
                oLong = 3
            Case 4 ' 4 notes reçues -> 5 notes d'arpège générées
                Dim oTab(5) As String
                oTab(0) = tbl1(0)
                oTab(1) = tbl1(2)
                oTab(2) = tbl1(1)
                oTab(3) = tbl1(3)
                oTab(4) = tbl1(2)
                Liste1 = Join(oTab, " ")
                oLong = 5
            Case 5 ' 5 notes reçues -> 7 notes d'arpèges générées
                Dim oTab(6) As String
                oTab(0) = tbl1(0)
                oTab(1) = tbl1(2)
                oTab(2) = tbl1(1)
                oTab(3) = tbl1(3)
                oTab(4) = tbl1(2)
                oTab(5) = tbl1(4)
                oTab(6) = tbl1(3)
                Liste1 = Join(oTab, " ")
                oLong = 7
            Case 6 ' 6 notes reçues -> 9 notes d'arpèges générées
                Dim oTab(6) As String
                oTab(0) = tbl1(0)
                oTab(1) = tbl1(2)
                oTab(2) = tbl1(1)
                oTab(3) = tbl1(3)
                oTab(4) = tbl1(2)
                oTab(5) = tbl1(4)
                oTab(6) = tbl1(3)
                oTab(7) = tbl1(5)
                oTab(8) = tbl1(4)
                Liste1 = Join(oTab, " ")
                oLong = 9
            Case 7 ' 7 notes reçues -> 11 notes d'arpèges générées
                Dim oTab(6) As String
                oTab(0) = tbl1(0)
                oTab(1) = tbl1(2)
                oTab(2) = tbl1(1)
                oTab(3) = tbl1(3)
                oTab(4) = tbl1(2)
                oTab(5) = tbl1(4)
                oTab(6) = tbl1(3)
                oTab(7) = tbl1(5)
                oTab(8) = tbl1(4)
                oTab(9) = tbl1(6)
                oTab(10) = tbl1(5)
                Liste1 = Join(oTab, " ")
                oLong = 11
        End Select
        ArpMotif1 = CalcMotif(Liste1, metrique, nbMesures, duree, oLong, Début)
    End Function
    ' *********************************************************************
    ' ArpMotif 2	C	E	G	B	D		C	E	C	G	C	B	C	D *
    ' *********************************************************************
    Public Function ArpMotif2(ListNotes As String, metrique As String, nbMesures As Integer, duree As Double, Début As Integer) As String
        Dim Liste1 As String = ""
        Dim oLong As Integer
        Dim tbl1() As String = Split(ListNotes)
        ArpMotif2 = ""
        Select Case UBound(tbl1) + 1
            Case 3 ' 3 notes reçues -> 3 notes d'arpèges générées
                Dim oTab(3) As String
                oTab(0) = tbl1(0)
                oTab(1) = tbl1(2)
                oTab(2) = tbl1(0)
                oTab(3) = tbl1(1)
                Liste1 = Join(oTab, " ")
                oLong = 4
            Case 4 ' 4 notes reçues -> 5 notes d'arpège générées
                Dim oTab(5) As String
                oTab(0) = tbl1(0)
                oTab(1) = tbl1(3)
                oTab(2) = tbl1(0)
                oTab(3) = tbl1(2)
                oTab(4) = tbl1(0)
                oTab(5) = tbl1(1)
                Liste1 = Join(oTab, " ")
                oLong = 6
            Case 5 ' 5 notes reçues -> 7 notes d'arpèges générées
                Dim oTab(7) As String
                oTab(0) = tbl1(0)
                oTab(1) = tbl1(4)
                oTab(2) = tbl1(0)
                oTab(3) = tbl1(3)
                oTab(4) = tbl1(0)
                oTab(5) = tbl1(2)
                oTab(6) = tbl1(0)
                oTab(7) = tbl1(1)

                Liste1 = Join(oTab, " ")
                oLong = 8
            Case 6 ' 6 notes reçues -> 9 notes d'arpèges générées
                Dim oTab(9) As String
                oTab(0) = tbl1(0)
                oTab(1) = tbl1(5)
                oTab(2) = tbl1(0)
                oTab(3) = tbl1(4)
                oTab(4) = tbl1(0)
                oTab(5) = tbl1(3)
                oTab(6) = tbl1(0)
                oTab(7) = tbl1(2)
                oTab(8) = tbl1(0)
                oTab(9) = tbl1(1)
                Liste1 = Join(oTab, " ")
                oLong = 10
            Case 7 ' 7 notes reçues -> 11 notes d'arpèges générées
                Dim oTab(11) As String
                oTab(0) = tbl1(0)
                oTab(1) = tbl1(6)
                oTab(2) = tbl1(0)
                oTab(3) = tbl1(5)
                oTab(4) = tbl1(0)
                oTab(5) = tbl1(4)
                oTab(6) = tbl1(0)
                oTab(7) = tbl1(3)
                oTab(8) = tbl1(0)
                oTab(9) = tbl1(2)
                oTab(10) = tbl1(0)
                oTab(11) = tbl1(1)
                Liste1 = Join(oTab, " ")
                oLong = 12
        End Select
        ArpMotif2 = CalcMotif(Liste1, metrique, nbMesures, duree, oLong, Début)
    End Function
    ' *********************************************************************
    ' ArpMotif 3	C	E	G	B	D		D	B	D	G	D	E	D	C *
    ' *********************************************************************
    Public Function ArpMotif3(ListNotes As String, metrique As String, nbMesures As Integer, duree As Double, Début As Integer) As String
        Dim Liste1 As String = ""
        Dim oLong As Integer
        Dim tbl1() As String = Split(ListNotes)
        ArpMotif3 = ""
        Select Case UBound(tbl1) + 1
            Case 3                ' 3 notes reçues -> 3 notes d'arpèges générées
                Dim oTab(3) As String
                oTab(0) = tbl1(2)
                oTab(1) = tbl1(1)
                oTab(2) = tbl1(2)
                oTab(3) = tbl1(0)
                Liste1 = Join(oTab, " ")
                oLong = 4
            Case 4 ' 4 notes reçues -> 5 notes d'arpège générées
                Dim oTab(5) As String
                oTab(0) = tbl1(3)
                oTab(1) = tbl1(2)
                oTab(2) = tbl1(3)
                oTab(3) = tbl1(1)
                oTab(4) = tbl1(3)
                oTab(5) = tbl1(0)
                Liste1 = Join(oTab, " ")
                oLong = 6
            Case 5 ' 5 notes reçues -> 7 notes d'arpèges générées
                Dim oTab(7) As String
                oTab(0) = tbl1(4)
                oTab(1) = tbl1(3)
                oTab(2) = tbl1(4)
                oTab(3) = tbl1(2)
                oTab(4) = tbl1(4)
                oTab(5) = tbl1(1)
                oTab(6) = tbl1(4)
                oTab(7) = tbl1(0)

                Liste1 = Join(oTab, " ")
                oLong = 8
            Case 6 ' 6 notes reçues -> 9 notes d'arpèges générées
                Dim oTab(9) As String
                oTab(0) = tbl1(5)
                oTab(1) = tbl1(4)
                oTab(2) = tbl1(5)
                oTab(3) = tbl1(3)
                oTab(4) = tbl1(5)
                oTab(5) = tbl1(2)
                oTab(6) = tbl1(5)
                oTab(7) = tbl1(1)
                oTab(8) = tbl1(5)
                oTab(9) = tbl1(0)
                Liste1 = Join(oTab, " ")
                oLong = 10
            Case 7 ' 7 notes reçues -> 11 notes d'arpèges générées
                Dim oTab(11) As String
                oTab(0) = tbl1(6)
                oTab(1) = tbl1(5)
                oTab(2) = tbl1(6)
                oTab(3) = tbl1(4)
                oTab(4) = tbl1(6)
                oTab(5) = tbl1(3)
                oTab(6) = tbl1(6)
                oTab(7) = tbl1(2)
                oTab(8) = tbl1(6)
                oTab(9) = tbl1(1)
                oTab(10) = tbl1(6)
                oTab(11) = tbl1(0)
                Liste1 = Join(oTab, " ")
                oLong = 12
        End Select
        ArpMotif3 = CalcMotif(Liste1, metrique, nbMesures, duree, oLong, Début)
    End Function
    ' *******************************************************************
    ' ArpMotif 4	C	E	G	B	D		C	E	C	E	C	E ...   *
    ' *******************************************************************
    Public Function ArpMotif4(ListNotes As String, metrique As String, nbMesures As Integer, duree As Double, Début As Integer) As String
        Dim Liste1 As String = ""
        Dim oLong As Integer
        Dim tbl1() As String = Split(ListNotes)
        ArpMotif4 = ""
        Select Case UBound(tbl1) + 1
            Case 3, 4, 5, 6, 7 ' 3 notes reçues -> 3 notes d'arpèges générées
                Dim oTab(3) As String
                oTab(0) = tbl1(0)
                oTab(1) = tbl1(1)
                Liste1 = Join(oTab, " ")
                oLong = 2
        End Select
        ArpMotif4 = CalcMotif(Liste1, metrique, nbMesures, duree, oLong, Début)
    End Function
    'Public Function Motif1_Chord(ListNotes As String, metrique As String, nbMesures As Integer, duree As Double, Début As Integer) As String
    Function Motif1_Chord(ListNotes As String) As String
        Dim i As Integer
        Dim Liste1 As String = ""
        Dim tbl1() As String
        Dim tbl2() As String = Split(ListNotes)
        Dim L_Motifs As New List(Of String)
        Dim nbAcc As Integer = tbl2.Count
        Dim n As Integer = 0
        Dim r As String = ""


        Motif1_Chord = ""
        tbl1 = tbl2(0).Split("&")
        ' réordonnance des notes de chaque accords
        Select Case tbl1.Count
            Case 3
                For i = 0 To nbAcc - 1
                    Select Case n
                        Case 0
                            L_Motifs.Add(tbl1(0) + "&" + tbl1(1) + "&" + tbl1(2)) ' C E G
                        Case 1
                            L_Motifs.Add(tbl1(1) + "&" + tbl1(2) + "&" + OctPlus(tbl1(0))) 'E G C
                        Case 2
                            L_Motifs.Add(tbl1(2) + "&" + OctPlus(tbl1(0)) + "&" + OctPlus(tbl1(1))) ' G C E
                        Case 3
                            L_Motifs.Add(tbl1(1) + "&" + tbl1(2) + "&" + OctPlus(tbl1(0))) ' E G C
                    End Select
                    n += 1
                    If n = 4 Then n = 0
                Next

            Case 4
                For i = 0 To nbAcc - 1
                    Select Case n
                        Case 0
                            L_Motifs.Add(tbl1(0) + "&" + tbl1(1) + "&" + tbl1(2) + "&" + tbl1(3)) ' C E G B
                        Case 1
                            L_Motifs.Add(tbl1(1) + "&" + tbl1(2) + "&" + tbl1(3) + "&" + OctPlus(tbl1(0))) ' E G B C
                        Case 2
                            L_Motifs.Add(tbl1(2) + "&" + tbl1(3) + "&" + OctPlus(tbl1(0)) + "&" + OctPlus(tbl1(1))) ' G B C E
                        Case 3
                            L_Motifs.Add(tbl1(3) + "&" + OctPlus(tbl1(0)) + "&" + OctPlus(tbl1(1)) + "&" + OctPlus(tbl1(2))) ' B C E G
                        Case 4
                            L_Motifs.Add(tbl1(2) + "&" + tbl1(3) + "&" + OctPlus(tbl1(0)) + "&" + OctPlus(tbl1(1))) ' G B C E
                        Case 5
                            L_Motifs.Add(tbl1(1) + "&" + tbl1(2) + "&" + tbl1(3) + "&" + OctPlus(tbl1(0))) ' E G B C
                    End Select
                    n += 1
                    If n = 6 Then n = 0
                Next
            Case 5
                For i = 0 To nbAcc - 1
                    Select Case n
                        Case 0
                            L_Motifs.Add(tbl1(0) + "&" + tbl1(1) + "&" + tbl1(2) + "&" + tbl1(3) + "&" + tbl1(4)) ' C E G B D
                        Case 1
                            L_Motifs.Add(tbl1(1) + "&" + tbl1(2) + "&" + tbl1(3) + "&" + tbl1(4) + "&" + OctPlus(tbl1(0))) ' E G B D C
                        Case 2
                            L_Motifs.Add(tbl1(2) + "&" + tbl1(3) + "&" + tbl1(4) + "&" + OctPlus(tbl1(0)) + "&" + OctPlus(tbl1(1))) ' G B D C E
                        Case 3
                            L_Motifs.Add(tbl1(3) + "&" + tbl1(4) + "&" + OctPlus(tbl1(0)) + "&" + OctPlus(tbl1(1)) + "&" + OctPlus(tbl1(2))) ' B D C E G
                        Case 4
                            L_Motifs.Add(tbl1(4) + "&" + OctPlus(tbl1(0)) + "&" + OctPlus(tbl1(1)) + "&" + OctPlus(tbl1(2)) + "&" + OctPlus(tbl1(3)))' D C E G B
                        Case 5
                            L_Motifs.Add(tbl1(3) + "&" + tbl1(4) + "&" + OctPlus(tbl1(0)) + "&" + OctPlus(tbl1(1)) + "&" + OctPlus(tbl1(2))) ' B D C E G
                        Case 6
                            L_Motifs.Add(tbl1(2) + "&" + tbl1(3) + "&" + tbl1(4) + "&" + OctPlus(tbl1(0)) + "&" + OctPlus(tbl1(1))) ' G B D C E
                        Case 7
                            L_Motifs.Add(tbl1(1) + "&" + tbl1(2) + "&" + tbl1(3) + "&" + tbl1(4) + "&" + OctPlus(tbl1(0))) ' E G B D C
                    End Select
                    n += 1
                    If n = 8 Then n = 0
                Next

        End Select
        ' transorfer liste en chaine de caractérèe
        For Each a As String In L_Motifs
            r = r + a + " "
        Next
        Return (Trim(r))
    End Function
    'Select Case n
    'Case 0
    '                        L_Motifs.Add(tbl1(0) + "&" + tbl1(1) + "&" + tbl1(2))
    ' Case 1
    '                        L_Motifs.Add(tbl1(2) + "&" + OctPlus(tbl1(0)) + "&" + OctPlus(tbl1(1)))
    '                       'L_Motifs.Add(tbl1(1) + "&" + tbl1(2) + "&" + OctPlus(tbl1(0)))
    'Case 2
    '                       'L_Motifs.Add(tbl1(2) + "&" + OctPlus(tbl1(0)) + "&" + OctPlus(tbl1(1)))
    '                       L_Motifs.Add(tbl1(0) + "&" + tbl1(1) + "&" + tbl1(2))
    'Case 3
    '                        L_Motifs.Add(tbl1(1) + "&" + tbl1(2) + "&" + OctPlus(tbl1(0)))
    'End Select
    '               n += 1
    'If n = 4 Then n = 0
    Function OctPlus(Note As String) As String
        Dim i As Integer
        Dim n As String
        n = Note
        i = ValNoteCubase.IndexOf(Note)
        If i + 12 <= ValNoteCubase.Count - 1 Then
            n = ValNoteCubase(i + 12)
        End If
        Return n
    End Function
    Public Function Répé_old(TypNote As Integer, Accord As String, ListNotes As String, metrique As String, nbMesures As Integer, duree As Double, Début As Integer) As String
        Dim i As Integer
        Dim a As String
        Dim oLong As Integer
        Dim tbl() As String
        Dim tbl1() As String = Split(ListNotes)

        a = Det_NotesAccord(Trim(Accord))
        a = Trad_ListeNotesEnD(Trim(a), "-")
        tbl = Split(a, "-")
        a = UCase(tbl(TypNote))

        For i = 0 To UBound(tbl1)
            If InStr(tbl1(i), a) <> 0 Then
                Exit For
            End If
        Next i

        a = Trim(tbl1(i))
        oLong = 1
        Return CalcMotif(a, metrique, nbMesures, duree, oLong, Début)
    End Function
    Public Function Répé(Voix As Integer, ListNotes As String, metrique As String, nbMesures As Integer, duree As Double, Début As Integer, Motif As String, ChiffAcc As String) As String

        Dim oLong As Integer
        Dim tbl1() As String = Split(ListNotes)
        Dim tbl2() As String = Motif.Split()
        Dim tbl3() As String = ChiffAcc.Split()
        Dim a As String = ""
        Dim tr As String
        Dim qt As String
        Dim n As String
        Dim r As String = ""

        '
        ' détermination des notes à prendre en compte a=liste de notes - a peut posséder une seule note


        If tbl2(1) <> "Chord" Then
            If Voix <= UBound(tbl1) Then
                a = Trim(tbl1(Voix))
            Else
                a = Trim(tbl1(UBound(tbl1)))
            End If
        Else
            'If tbl2(0) = "Full" Then
            '    Listesnotes = Trim(ListNotes) 
            '    a = ListNotes.Replace(" ", "&")
            'Else
            'End If
            Select Case tbl2(0)
                Case "Full", "Motif1"
                    Listesnotes = Trim(ListNotes)
                    a = ListNotes.Replace(" ", "&")
                Case "No3rd"
                    n = Trad_BemDies(tbl3(0))
                    i = ListNotesMajd.IndexOf(n)
                    If Len(ChiffAcc) > 1 Then
                        If ChiffAcc.Chars(2) = "m" Then
                            tr = ListNotesMajd(i + 3) ' tierce lmineure
                        Else
                            tr = ListNotesMajd(i + 4) ' tierce majeure
                        End If
                    Else
                        tr = ListNotesMajd(i + 4) ' tierce majeure
                    End If
                    '
                    For Each a1 As String In tbl1 ' retrait de la tierce
                        n = a1.Chars(0)
                        If Trim(n) <> Trim(tr) Then
                            r = r + a1 + "&"
                        End If
                    Next
                    a = r.Substring(0, r.Length - 1)
                Case "No5th"
                    n = Trad_BemDies(tbl3(0))
                    i = ListNotesMajd.IndexOf(n)
                    qt = ListNotesMajd(i + 7) ' Quinte
                    '
                    For Each a1 As String In tbl1 ' retrait de la tierce
                        n = a1.Chars(0)
                        If Trim(n) <> Trim(qt) Then
                            r = r + a1 + "&"
                        End If
                    Next
                    a = r.Substring(0, r.Length - 1)

            End Select
        End If
        oLong = 1
        Return CalcMotif(a, metrique, nbMesures, duree, oLong, Début)
    End Function

    Public Function Perso(numGrid As Integer, ListNotes As String, metrique As String, nbMesures As Integer, Début As Integer) As String
        Dim Liste1 As String = ""
        Dim Dico As New Dictionary(Of String, String)
        Dim tbl1() As String = Split(ListNotes)
        Perso = ""
        Select Case UBound(tbl1) + 1
            Case 3
                Dico("A") = tbl1(2)
                Dico("B") = tbl1(2)
                Dico("C") = tbl1(2)
                Dico("D") = tbl1(1)
                Dico("E") = tbl1(0)
            Case 4
                Dico("A") = tbl1(3)
                Dico("B") = tbl1(3)
                Dico("C") = tbl1(2)
                Dico("D") = tbl1(1)
                Dico("E") = tbl1(0)
            Case 5
                Dico("A") = tbl1(4)
                Dico("B") = tbl1(3)
                Dico("C") = tbl1(2)
                Dico("D") = tbl1(1)
                Dico("E") = tbl1(0)
        End Select
        Perso = Form1.Récup_MotifPerso2(numGrid, Dico, nbMesures, Début)
    End Function
End Module
