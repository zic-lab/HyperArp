Public Class RechercheG_v2

    ' *****************************************
    ' *                                       *
    ' *            VARIABLES                  *
    ' *                                       *
    ' *****************************************
    Dim Exemple_ As String
    Dim ListGammesBase As New List(Of String)
    Public DicoGammesBase As New Dictionary(Of String, String)
    Dim ListAccordsBase As New List(Of String)
    Dim DicoAccordsBase As New Dictionary(Of String, String)


    '
    ' *****************************************
    ' *                                       *
    ' *            PROPRIETES                 *
    ' *                                       *
    ' *****************************************
    ' Propriété type
    Property ExempleType() As String
        Get
            Return Exemple_
        End Get
        Set(ByVal Value As String)
            Exemple_ = Value
        End Set
    End Property
    ' Gammes Bases
    ' ************
    Dim GammesBases_ As String
    Property GammesBases() As String
        Get
            Return GammesBases_
        End Get
        Set(ByVal Value As String)
            GammesBases_ = Value
            ChargementGammesBase()
        End Set
    End Property

    ' Accords Bases
    Dim AccordsBases_ As String
    Property AccordsBases() As String
        Get
            Return AccordsBases_
        End Get
        Set(ByVal Value As String)
            AccordsBases_ = Value
            ChargementAccordsBase()
        End Set
    End Property
    Class LGAccord
        Public LGAcc As New List(Of String)
    End Class
    Dim ListGammesAcc As New List(Of LGAccord)
    Dim ListeGlobale As New List(Of String)
    Dim Tabnotes As New List(Of String)
    Dim mTabnotes As New List(Of String)
    Dim DIntv As New Dictionary(Of String, Integer) From {{"1", 0}, {"#1", 1}, {"b2", 1}, {"2", 2}, {"#2", 3}, {"b3", 3}, {"3", 4}, {"b4", 4}, {"4", 5},
     {"#4", 6}, {"b5", 6}, {"5", 7}, {"#5", 8}, {"b6", 8}, {"6", 9}, {"#6", 10}, {"b7", 10}, {"7", 11}, {"b9", 13}, {"9", 14}, {"#9", 15}, {"11", 17}, {"#11", 18}}
    Public Sub Maj_TabNotes()
        Dim i As Integer
        For i = 0 To 3
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
            '
            mTabnotes.Add("c")
            mTabnotes.Add("c#")
            mTabnotes.Add("d")
            mTabnotes.Add("d#")
            mTabnotes.Add("e")
            mTabnotes.Add("f")
            mTabnotes.Add("f#")
            mTabnotes.Add("g")
            mTabnotes.Add("g#")
            mTabnotes.Add("a")
            mTabnotes.Add("a#")
            mTabnotes.Add("b")
        Next
    End Sub
    ''' <summary>
    ''' ChargementGammesBase : chargement de la totalité des gammes possibles dans une Liste et dans un dictionnaire (clef=nom de la gamme)
    ''' </summary>
    Sub ChargementGammesBase()
        Dim line As String
        Dim tbl() As String

        ListGammesBase.Clear()
        DicoGammesBase.Clear()

        Using sr As IO.StringReader = New IO.StringReader(Me.GammesBases)
            While sr.Peek() >= 0 ' Boucler jusqu'à la fin du fichier
                line = sr.ReadLine() ' Lire chaque ligne
                ListGammesBase.Add(line)
                tbl = line.Split(";")
                DicoGammesBase.Add((tbl(2)), line) ' tbl(2) chiffrage de l'accord)
            End While
        End Using ' Fermer

    End Sub
    ''' <summary>
    ''' ChargementAccordsBase :  chargement de la totalité des accords possible dans une Liste et dans un dictionnaire (clef=chiffrage de l'accord)
    ''' </summary>
    Sub ChargementAccordsBase()
        Dim line As String
        Dim tbl() As String

        ListAccordsBase.Clear()
        DicoAccordsBase.Clear()

        Using sr As IO.StringReader = New IO.StringReader(Me.AccordsBases)
            While sr.Peek() >= 0 ' Boucler jusqu'à la fin du fichier
                line = sr.ReadLine() ' Lire chaque ligne
                ListAccordsBase.Add(line)
                tbl = line.Split(";")
                DicoAccordsBase.Add((tbl(0)), line)
            End While
        End Using ' Fermer
    End Sub
    '
    ' *****************************************
    ' *                                       *
    ' *            CONSTRUCTEUR               *
    ' *                                       *
    ' *****************************************
    Sub New()
        Maj_TabNotes() ' mise des TabNote et de mTabNotes
        'Dim a As String = My.Resources.FichierListeGammes
        'Dim a As String = My.Resources.FichierListeGammes_Cubasev2 ' FichierListeGammes_Cubasev4
        Dim a As String = My.Resources.FichierListeGammes_Cubasev4
        Me.GammesBases = a ' mise à jour du fichier global des gammes 
        a = My.Resources.FichierListeAccords
        Me.AccordsBases = a ' mise à jour du fichier global des accords
    End Sub
    '
    ' *****************************************
    ' *                                       *
    ' *        METHODES DE RECHERCHE          *
    ' *                                       *
    ' *****************************************
    '

    Function Supp_Doublons2(liste As String) As String
        Dim L As New List(Of String)
        Dim tbl() As String = liste.Split("-")
        For Each a As String In tbl
            L.Add(a)
        Next
        ' Pour test doublons
        'L.Add("C Maj")
        'L.Add("D Maj")
        'L.Add("G Maj")
        'L.Add("B MinH")
        Dim distinctGammes As IEnumerable(Of String) = L.Distinct()
        Dim c As String = ""
        For Each b As String In distinctGammes
            c = c + b + "-"
        Next
        c = c.Substring(0, c.Length - 1)
        Return c
    End Function

    ''' <summary>
    ''' GammesJouables : Appartenance d'un Accord à un type de gamme (12 gammes par type)
    ''' </summary>
    ''' <param name="Accord"></param>
    ''' <param name="TyPG">Chiffrage de la gamme</param>
    ''' <returns></returns>
    Private Function GammesJouables(Accord As String, TyPG As String) As String
        Dim L As String = Det_NotesAccord3(Accord)
        Dim tblA() As String = L.Split("-") ' notes de l'accord
        Dim TBLG() As String ' notes de la Gamme
        Dim Chroma As New List(Of String) From {"C", "C#", "D", "D#", "E", "F", "F#", "G", "G#", "A", "A#", "B"}
        Dim LG As String = "" ' liste de gammes
        Dim G As String
        Dim flag As Boolean = True

        For Each n As String In Chroma
            ' détermination des notes de la gamme
            G = n + " " + TyPG
            TBLG = Det_NotesGammes3(G).Split("-")
            flag = True
            ' appartenance des notes de l'accord à la gamme
            For Each a As String In tblA
                If TBLG.Contains(a) = False Then
                    flag = False
                    Exit For
                End If
            Next
            If flag Then LG = Trim(LG) + UCase(n) + " " + TyPG + "-"
        Next
        '
        If Trim(LG) <> "" Then LG = Trim(LG.Substring(0, LG.Length - 1)) ' retrait du dernier "-"
        Return LG

    End Function
    ' Constitution d'une liste des Gammes jouables sur  un Accord
    ' ************************************************************
    Private Function ListeDesGammes(Accord As String, ListeTypeG As String) As String

        Dim i As Integer
        Dim a As String = ""
        Dim aa As String = ""
        Dim TBLTG() As String = ListeTypeG.Split()
        Dim LGammes As String = ""
        Dim tbl1(), tbl2() As String
        Dim Tbl3() As String
        Dim bb As Char
        Dim s As String = ""
        Dim tonique As String

        Tbl3 = Accord.Split()
        For Each T As String In TBLTG
            'If Trim(T) <> "PMin" Then ' PMIN vient de PMIN1 dans la listes des gammes Usuelles. On transforme PMin1 par PMin au moment de la sélection de la liste de gamme. On repère ainsi le PMin1 qui vient des gammes Usuelles pour le traiter de manière différente que PMin1 dans les 5Notes.
            If T <> "Blues" And T <> "Blues2" Then
                a = GammesJouables(Accord, T)
            End If
            If Trim(a) <> "" Then LGammes = Trim(LGammes) + a + "-"
            If Trim(T) = "Maj" And Trim(a) <> "" Then
                tbl1 = a.Split("-")
                For Each b As String In tbl1
                    tbl2 = b.Split()
                    i = Tabnotes.IndexOf(tbl2(0))
                    tonique = Trim(Tabnotes(i + 9))
                    LGammes = Trim(LGammes) + tonique + " " + "PentaMin" + "-"
                Next
            End If
            '
        Next
        '
        tonique = Trim(Tbl3(0))
        If Tbl3.Count > 1 Then
            bb = Tbl3(1)(0)
            If Len(Tbl3(1)) >= 2 Then
                s = Microsoft.VisualBasic.Left(Tbl3(1), 2)
            End If
        End If
        ' BLUES : si présence d'une 7e de dominante
        ' ****************************************
        If bb.ToString = "7" Or s = "m7" Then
            '
            i = Tabnotes.IndexOf(tonique)
            'tonique = Tabnotes(i + 2)
            a = tonique + " " + "Blues" + "-" + tonique + " " + "Blues2" + "-" ' + Tabnotes(i + 7) + " " + "Blues" + "-" 
            LGammes = Trim(LGammes) + a
        End If
        If Trim(LGammes) <> "" Then LGammes = Trim(LGammes.Substring(0, LGammes.Length - 1))
        Return LGammes
        '
    End Function



    ' Constitution d'une liste des Gammes pour tous les Accords
    ' *********************************************************
    Private Sub CalcListesGammes(ListAccord As String, ListeTypeG As String)
        Dim TBLA() As String = ListAccord.Split("-")
        Dim aa As String

        ListeGlobale.Clear()

        For Each Acc As String In TBLA ' TBLA contient les accords
            aa = ListeDesGammes(Acc, ListeTypeG)
            ListeGlobale.Add(aa)
        Next
    End Sub

    ' Détermination des gammes communes
    ' *********************************

    ''' <summary>
    ''' ApparteanceG : Méthode principales - Calcul des gammes jouables pour une liste d'accord 6Méh
    ''' </summary>
    ''' <param name="ListAccord">Liste d'accords séparés par des "-"</param>
    ''' <param name="ListeTypeG">Liste des gammes séparées par des " "</param>
    ''' <returns></returns>
    Public Function ApparteanceG(ListAccord As String, ListeTypeG As String) As String
        Dim i, j As Integer
        Dim LResult As New List(Of String)
        Dim Result As String
        Dim OccurDeAcc As Integer = 0
        'Dim GamJouables As New List(Of String)
        Dim List1 As New List(Of String)
        Dim List2 As New List(Of String)
        Dim intersection As IEnumerable(Of String) = {}
        Dim GamJouables As New List(Of IEnumerable(Of String))

        If Trim(ListAccord) <> "" And Trim(ListeTypeG) <> "" Then
            ' I) Détermination des Gammes jouables pour chaque Accord : résultat dans ListeGlobale(0 à n-1) n = nombre d'accords. ListeGlobale est une liste de listes
            '    *****************************************************************************************************************************************************
            CalcListesGammes(ListAccord, ListeTypeG) ' calcul des gammes jouables sur chaque accord de ListAccord : le résultat est dans ListeGlobale

            ' II) Détermination des gammes communes à tous les accords
            '     ****************************************************
            ' Parcours de la Liste des gammes des autres accords à partir de list du 1ere accord
            'GamJouables = ChaineEnListe(ListeGlobale(0))
            List1 = ChaineEnListe(ListeGlobale(0), "-")
            If ListeGlobale.Count > 1 Then
                List2 = ChaineEnListe(ListeGlobale(1), "-")
                GamJouables.Add(intersection) ' GamJouables est une liste de IEnumerable
                GamJouables(0) = List1.Intersect(List2)
                '
                i = 2
                j = 1
                While i <= ListeGlobale.Count - 1
                    List1 = ChaineEnListe(ListeGlobale(i), "-")
                    GamJouables.Add(intersection)
                    GamJouables(j) = GamJouables(j - 1).Intersect(List1)
                    j += 1
                    i += 1
                End While

                Result = EnumérableEnChaine(GamJouables(j - 1))
            Else
                Result = Trim(ListeGlobale(0))
            End If
        Else
            Result = ""
        End If
        ' Suppression des doublons s'il en existe
        Result = Supp_Doublons2(Result)
        Return Result
    End Function
    Public Function ChaineEnListe(a As String, sep As String) As List(Of String)
        Dim i As Integer
        Dim tbl() As String
        Dim List1 As New List(Of String)
        '
        tbl = a.Split(sep)
        '
        For i = 0 To tbl.Count - 1
            List1.Add(tbl(i))
        Next
        '
        Return List1
    End Function
    Function EnumérableEnChaine(enu As IEnumerable(Of String)) As String
        Dim i As Integer
        Dim a As String = " "
        Dim tbl() As String
        Dim List1 As New List(Of String)
        '
        tbl = enu.ToArray()
        '
        For i = 0 To tbl.Count - 1
            a = a + Trim(tbl(i)) + "-"
        Next
        '
        a = Trim(a)
        If a <> "" Then a = Trim(a.Substring(0, a.Length - 1))
        Return a
    End Function

    ' *****************************************
    ' *                                       *
    ' *             SUBROUTINES               *
    ' *                                       *
    ' *****************************************

    Private Function Création_CTemp() As String
        Dim DossierDocuments As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        Dim a As String
        a = DossierDocuments + "\HyperArp"
        Création_CTemp = a
        With My.Computer.FileSystem
            If Not (.DirectoryExists(a)) Then
                .CreateDirectory(a)
            End If
        End With
    End Function
    Private Function Det_NotesAccord3(Accord As String) As String
        Dim indTon As Integer
        Dim tbl1() As String
        Dim tbl2() As String
        Dim info As String
        Dim a As String = ""
        Dim b As String = ""
        Dim i, j As Integer

        tbl1 = Accord.Split()
        If tbl1.Count = 1 Then
            Accord = Accord + " " + "Maj"
        End If
        tbl1 = Accord.Split()
        '
        indTon = Det_IndexTonique(tbl1(0)) ' détermination de l'index de la tonique dans TabNotes / tbl1(0) = tonique de l'accord
        '
        info = DicoAccordsBase(tbl1(1))    ' lecture de la ligne d'info de l'accord / tbl1(1) = chiffrage de l'accord
        tbl2 = info.Split(";")
        '
        For i = 2 To tbl2.Count - 1
            j = DIntv(tbl2(i)) 'lecture index de la note de l'accord
            b = mTabnotes(indTon + j)
            a = a + " " + b
        Next
        '
        a = Trim(a)
        a = a.Replace(" ", "-")
        Return Trim(a)
    End Function
    Public Function Det_NotesGammes3(Gamme As String) As String
        Dim i, j As Integer
        Dim tbl1() As String
        Dim tbl2() As String
        Dim a As String = ""
        Dim b As String = ""

        Try
            tbl1 = Gamme.Split() ' tbl(0) Tonique de la gamme ; tbl(1)=Chiffrage de la gamme
            Dim indTonique As String = Det_IndexTonique(Trim(tbl1(0)))

            ' RUSTINE 2 : tranformation de Blues en Blues1 pour rendre compatible les anciens fichiers Microsoft wave Table Synth
            ''If Trim(tbl1(1)) = "Blues" Then
            'tbl1(1) = "Blues1"
            'End If

            Dim info As String = DicoGammesBase(Trim(tbl1(1)))

            tbl2 = info.Split(";")
            i = 3

            While tbl2(i) <> "Fin"
                j = DIntv(tbl2(i)) 'lecture index de la note de l'accord
                b = mTabnotes(indTonique + j)
                a = a + " " + b
                i += 1
            End While
            '
            a = Trim(a)
            a = a.Replace(" ", "-")
            Return Trim(a)
        Catch
            Dim result As DialogResult
            Dim bouton As MessageBoxButtons = MessageBoxButtons.OK
            Dim message As String = "Gamme inconnue " + "Gamme : " + Gamme
            Dim titre As String = "Erreur détectée dans oo.Det_NotesGammes3"
            '
            result = MessageBox.Show(message, titre, bouton)
            Return Trim(a)
        End Try
    End Function
    Public Function Det_InfoGamme(Gamme As String) As String
        Dim info As String = ""
        If Trim(Gamme <> "") Then
            info = DicoGammesBase(Trim(Gamme))
        End If
        Return info
    End Function
    Private Function Det_IndexTonique(tonique As String) As Integer
        Dim i As Integer = -1
        Do
            i += 1
        Loop Until tonique = Tabnotes(i) Or i >= Tabnotes.Count ' 
        Return i
        '
    End Function
End Class


