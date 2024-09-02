Public Class RechercheG
    Class LGAccord
        Public LGAcc As New List(Of String)
    End Class
    Dim ListGammesAcc As New List(Of LGAccord)
    Dim ListeGlobale As New List(Of String)
    Dim Tabnotes As New List(Of String)
    Public Sub Maj_TabNotes()
        Dim i As Integer
        For i = 0 To 2
            Tabnotes.Add("C")
            Tabnotes.Add("C#")
            Tabnotes.Add("D")
            Tabnotes.Add("D#")
            Tabnotes.Add("E")
            Tabnotes.Add("F")
            Tabnotes.Add("F#")
            Tabnotes.Add("G")
            Tabnotes.Add("G#")
            Tabnotes.Add("A")
            Tabnotes.Add("A#")
            Tabnotes.Add("B")
        Next
    End Sub
    '
    ' *****************************************
    ' *                                       *
    ' *            CONSTRUCTEUR               *
    ' *                                       *
    ' *****************************************
    Sub New()
        Maj_TabNotes()
    End Sub




    ' *****************************************
    ' *                                       *
    ' *        METHODES DE RECHERCHE          *
    ' *                                       *
    ' *****************************************
    Function Supp_Doublons(liste As String) As String
        Dim Dic As New Dictionary(Of String, String)
        Dim TBL3() As String = liste.Split("-")
        Dim ListRetour As String = ""

        For Each a As String In TBL3
            If Not Dic.ContainsKey(a) Then
                Dic.Add(a, a)
            End If
        Next
        ' Récupération des valeurs du dictionnaire
        Dim valueColl As Dictionary(Of String, String).ValueCollection = Dic.Values
        '
        For Each a As String In valueColl
            Listretour = Trim(Listretour) + a + "-"
        Next
        Return ListRetour
    End Function
    ' Appartenance d'un Accord à un type de gamme (12 gammes par type)
    ' ****************************************************************
    Private Function GammesJouables(Accord As String, TyPG As String) As String
        Dim L As String = Form1.Det_NotesAccord3(Accord, "#")
        Dim tblA() As String = L.Split("-") ' notes de l'accord
        Dim TBLG() As String ' notes de la Gamme
        Dim Chroma As New List(Of String) From {"c", "c#", "d", "d#", "e", "f", "f#", "g", "g#", "a", "a#", "b"}
        Dim LG As String = "" ' liste de gammes
        Dim G As String
        Dim flag As Boolean = True
        Dim i As Integer

        For Each n As String In Chroma
            ' détermination des notes de la gamme
            G = n + " " + TyPG
            TBLG = Det_NotesGammes3(G, "#").Split
            flag = True
            ' appartenance des notes de l'accord à la gamme
            For Each a As String In tblA
                If TBLG.Contains(a) = False Then flag = False
            Next
            If flag Then LG = Trim(LG) + UCase(n) + " " + TyPG + "-"
        Next
        If Trim(LG) <> "" Then
            LG = Trim(LG.Substring(0, LG.Length - 1))

            ' traitement des penta mineures en fonction de l'appartenance d'un accord à une (ou plusieurs) gammes  Majeurs
            ' ************************************************************************************************************
            Dim TBLM() As String = LG.Split("-")
            Dim tbl() As String
            Dim LPmin As String = ""

            For Each a As String In TBLM
                tbl = a.Split()
                If Trim(tbl(1)) = "Maj" Then ' les pentatoniques fournies ici sont tirées directement d'une gamme majeure, elles sont construites à partir des notes de la gamme majeure jouable
                    i = Tabnotes.IndexOf(tbl(0))
                    LPmin = Trim(LPmin) + Tabnotes(i + 2) + " " + "PMin" + "-" + Tabnotes(i + 4) + " " + "PMin" + "-" + Tabnotes(i + 9) + " " + "PMin" + "-"
                End If
            Next
            If Trim(LPmin) <> "" Then
                LPmin = Trim(LPmin.Substring(0, LPmin.Length - 1))
                LG = Trim(LG) + "-" + Trim(LPmin)
            End If
        End If
        Return LG

    End Function
    ' Constitution d'une liste des Gammes pour un Accord
    ' **************************************************
    Private Function ListeDesGammes(Accord As String, ListeTypeG As String) As String
        Dim a As String = ""
        Dim TBLT() As String = ListeTypeG.Split()
        Dim LGammes As String = ""

        For Each T As String In TBLT ' T type (chiffrage) de gammes
            a = GammesJouables(Accord, T)
            If Trim(a) <> "" Then LGammes = Trim(LGammes) + a + "-"
        Next
        '
        Return Trim(LGammes.Substring(0, LGammes.Length - 1))
        '
    End Function

    ' Constitution d'une liste des Gammes pour tous les Accords
    ' *********************************************************
    Private Sub CalcListesGammes(ListAccord As String, ListeTypeG As String)
        Dim TBLA() As String = ListAccord.Split("-")
        Dim aa As String

        For Each Acc As String In TBLA ' TBLA contient les accords
            aa = ListeDesGammes(Acc, ListeTypeG)
            ListeGlobale.Add(aa)
        Next
    End Sub

    ' Détermination des gammes communes
    ' *********************************

    ''' <summary>
    ''' ApparteanceG : Calcul des gammes jouables pour une liste d'accord
    ''' </summary>
    ''' <param name="ListAccord">Liste d'accords séparés par des "-"</param>
    ''' <param name="ListeTypeG">Liste des gammes séparées par des " "</param>
    ''' <returns></returns>
    Public Function ApparteanceG(ListAccord As String, ListeTypeG As String) As String
        Dim i As Integer
        Dim Result As String
        Dim OccurDeAcc As Integer = 0


        If Trim(ListAccord) <> "" Then
            CalcListesGammes(ListAccord, ListeTypeG) ' calcul des gammes jouables sur chaque accord de ListAccord : le résultat est dans ListeGlobale

            Dim TBLA() As String = ListAccord.Split("-")
            If TBLA.Count > 1 Then
                '
                ' Liste des gammes du 1er accord
                Dim TBL1() As String = ListeGlobale(0).Split("-")

                ' Parcours de la Liste des gammes des autres accords à partir de list du 1ere accord
                Dim TBL2() As String
                Dim ListGamCom As String = ""

                For Each Acc In TBL1 ' lecture des éléments de la première liste
                    For i = 1 To ListeGlobale.Count - 1
                        TBL2 = ListeGlobale(i).Split("-")

                        If TBL2.Contains(Acc) Then
                            OccurDeAcc = OccurDeAcc + 1 ' 
                            If OccurDeAcc = ListeGlobale.Count - 1 Then  ' l"accord doit apparaître dans toutes les listes de TBL2
                                ListGamCom = Trim(ListGamCom) + Acc + "-"
                            End If
                        End If
                    Next i
                    OccurDeAcc = 0
                Next
                If ListGamCom.Length <> 0 Then
                    Result = ListGamCom.Substring(0, ListGamCom.Length - 1)
                Else
                    Result = ""
                End If
            Else
                Result = Trim(ListeGlobale(0)) ' résultat de 1 seul accord demandé
            End If
        Else
            Result = ""
        End If
        '
        Result = Supp_Doublons(Result)
        Return Result
    End Function

End Class


