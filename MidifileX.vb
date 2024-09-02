Public Class MidifileX
    ' Constantes
    ' **********
    Const Rien = &HFF
    '
    Const CnoteOn = &H90
    Const CnoteOff = &H80
    Const CController = &HB0
    Const CAfterTouch = &HD0 'aftertouch monophonique
    Const CProgram = &HC0
    Const CnbCtrl = 5 ' nom de controleurs pount être enovyés simultanément sur une piste
    ' Structures
    ' **********
    Structure NteOn
        Dim note As Byte
        Dim velo As Byte
        Dim canal As Byte
        Dim pos As Integer
    End Structure
    Structure NteOff
        Dim note As Byte
        Dim velo As Byte
        Dim canal As Byte
        Dim pos As Integer
    End Structure
    Structure Ctrl ' CP Ctrl ou Program
        Dim ctrl As Byte
        Dim canal As Byte
        Dim pos As Integer
        Dim valeur As Byte
    End Structure
    Structure Prg ' CP Ctrl ou Program
        Dim prog As Byte
        Dim canal As Byte
        Dim pos As Integer
    End Structure
    Structure AfterTo ' CP Ctrl ou Program
        Dim aftert As Byte
        Dim canal As Byte
        Dim pos As Integer
    End Structure
    Structure ParPistes
        Dim Nom As String
        Dim Tempo As Integer ' utilisé uniquement dans la Piste(0)
        Dim Métrique As String ' utilisé uniquement dans la Piste(0)
    End Structure
    Structure Delt
        Dim PrésenceEvent As Boolean
        Dim Delta As Integer ' delta time en nombre de double croches
    End Structure
    ' VARIABLES PRIVEES
    ' *****************
    ' Varaibles typiques Midi
    ' ***********************
    Private NoteOn(-1, -1, -1) As NteOn ' (N° de pistes, nombre division, polyphonie)
    Private NoteOff(-1, -1, -1) As NteOff '(N° de pistes, nombre division, polyphonie) 
    Private Contrl(-1, -1, -1) As Ctrl ' (N° de pistes, nombre division, polyphonie)
    Private AfterT(-1, -1) As AfterTo ' (N° de pistes, nombre division, polyphonie)
    Private Program(-1, -1) As Prg ' (N° de pistes, nombre divisions)
    Private Deltatime(-1, -1) As Delt ' (N° de pistes, nombre divisions)
    Private Pistes(-1) As ParPistes ' (N° de pistes)
    Private Marq(-1) As String ' nombre de divisions
    Private MidiTexte(-1, -1) As String ' N° de piste, nombre de divisions
    '
    ' Variables typique fichier Midi
    ' ******************************
    Private TabFichierBin(-1) As Byte ' pour génération du fichier midi en binaire
    Private Réso As UShort ' nombre de ticks par noire (Pour HYperVoicing Réso=96 ticks/noire, pour Cubase Réso = 484 ticks/noire)
    Private MFType As UShort = 1 ' 0 : une seule piste; 1 : multipiste non passé dans le constructeur pour lemoment
    Private NbDivisions As Integer ' pour HyperVoicing = (16 * (NbMesures)) + 32 ' 16 correspond à un système de divisions de 16 doubles croches dans une mesure à 4/4 (32 correspond à 2 mesures de fin ne pouvant contenir que des noteOff, aucune noteOn ne peut y être écrite par l'utilisateur)
    Private NbPistes As Byte
    Private NbNotes As Byte ' nombre de notes pouvant apparaître simultanément - Pour Hypervoicing ce nombre est de 128 à cause des calques MIDI
    ' Variables pour calculs binaire du fichier Midi en particulier pour compter la longueur des pistes (longueur à placer en fin de piste).
    Private IndexLongueur As UInteger ' position du dernier octet d'une piste en cours de construction dans le fichier binaire
    Private ValLongueur As UInteger ' variable comptage de la longueur d'une piste en cours de construction
    Private Variation_Piste_Marq As Byte = 0
    '
    ' *********************************************************************************************************************
    ' piste :  N° de la piste - commence à 0
    ' canal :  de 0 à 15 (canal MIDI)
    ' note  :  de 0 à 127
    ' debut :  debut de l'éènement MIDI(Note,CTRL,Program) en nombre de divisions unitaires (pas en ticks, c'est calculer après) - division unitaire = croche
    ' durée :  longueur de la note en nombre de divisions unitaires - division unitaire = croche
    ' velo  :  vélocité de la notede 0 à 127
    ' ctrl  :  n° de controleur de 0 à 127 : par exemple ctrl = 7 --> Volume
    ' program : n° de programme de 0 à 127
    ' tempo : de 1 à 300
    ' nompiste : longueur variable
    ' **********************************************************************************************************************
    '
    ' ****************************
    ' * Contructeur de la classe *
    ' ****************************
    '
    ' Résolution : Nombre de ticks par Noire
    ' NBPistes   : Nombre de pistes MIDI sans compter la piste Marqueur (comptées à partir de 1)
    ' NB Mesures : Nombre de mesures des pistes.
    ' 
    Sub New(Résolution As UShort, NbPst As UShort, NbMesures As Integer) ' , Présence_Marqueurs As Boolean
        Dim i As Integer
        '

        '
        NbNotes = 128 ' polyphonie possible : nombre de notes pouvant être jouées simultanément sur une piste : un grand nombre est nécessaire pour les Calques MIDI
        NbDivisions = (16 * (NbMesures)) + 32 ' 16 correspnd à un système de division de 16 doubles croches dans une mesure à 4/4 (32 correspond à 2 mesures de fin ne pouvant contenir que des noteOff, aucune noteOn ne peut y être écrite par l'utilisateur)
        ' Si on a 12/8, c.a.d. 12 croches, c.a.d. 6 noires et 24 doubles criches  il faut alors compter 24*(NbMesures)+32
        NbPistes = NbPst + 1 '+ Variation_Piste_Marq  ' on rajoute une piste pour la piste système piste=0, une autre pour la piste marqueurs piste = 1 et éventuellement une piste pour les Marqueurs 
        Réso = Résolution
        MFType = 1 ' la présente classe utilise toujours un format MIDI file type 1 (plusieurs pistes)
        '
        ReDim NoteOn(NbPistes - 1, NbDivisions, NbNotes - 1)  '(nombre de pistes, nombre division, polyphonie)
        ReDim NoteOff(NbPistes - 1, (NbDivisions), NbNotes - 1) '(nombre de pistes, nombre division, polyphonie) 
        ReDim Contrl(NbPistes - 1, NbDivisions, CnbCtrl - 1) '(nombre de pistes, nombre division, polyphonie)
        ReDim Program(NbPistes - 1, NbDivisions) '(nombre de pistes, nombre division, polyphonie)
        ReDim AfterT(NbPistes - 1, NbDivisions) '(nombre de pistes, nombre division, polyphonie)
        ReDim Deltatime(NbPistes - 1, NbDivisions) ' (nombre de pistes, nombre divisions)
        ReDim Pistes(NbPistes - 1) '(nombre de pistes)
        ReDim Marq(NbDivisions)    ' nombre de division
        ReDim MidiTexte(NbPistes - 1, NbDivisions) ' nombre de pistes, nombre de division
        '
        For i = 0 To NbPistes - 1
            Pistes(i).Nom = ""
            For j = 0 To NbDivisions - 1
                For k = 0 To NbNotes - 1
                    NoteOn(i, j, k).note = Rien
                Next k
                '
                For k = 0 To NbNotes - 1
                    NoteOff(i, j, k).note = Rien
                Next k
                '
                For k = 0 To CnbCtrl - 1
                    Contrl(i, j, k).ctrl = Rien
                Next k
                '
                Program(i, j).prog = Rien
                Deltatime(i, j).PrésenceEvent = False
                '
                Marq(j) = ""
                MidiTexte(i, j) = ""
            Next j
        Next i
        '
        'i = UBound(Pistes, 1)
        '
    End Sub
    ' Propriétés accessibles
    Public Property Nb_Pistes() As String
        Get
            Return NbPistes
        End Get

        Set(ByVal Value As String)
            Me.NbPistes = Value
        End Set
    End Property

    ' Méthodes
    Public Sub AddNote(piste As Byte, canal As Byte, note As Byte, debut As Integer, duree As Integer, velo As Byte)
        Dim i As Integer
        piste = piste + 1 'Variation_Piste_Marq
        For i = 0 To (NbNotes - 1)
            If NoteOn(piste, debut, i).note = Rien Then
                NoteOn(piste, debut, i).note = note
                NoteOn(piste, debut, i).canal = canal
                NoteOn(piste, debut, i).velo = velo
                NoteOn(piste, debut, i).pos = debut
                Exit For
            End If
        Next i
        For i = 0 To (NbNotes - 1)
            If NoteOff(piste, debut + duree, i).note = Rien Then
                NoteOff(piste, debut + duree, i).note = note
                NoteOff(piste, debut + duree, i).canal = canal
                NoteOff(piste, debut + duree, i).velo = 0 ' noteoff
                NoteOff(piste, debut + duree, i).pos = debut + duree
                Exit For
            End If
        Next i
    End Sub

    Public Sub AddCTRL(piste As Byte, canal As Byte, debut As Integer, ctrl As Byte, valeur As Byte)
        Dim i As Integer
        piste = piste + 1 'Variation_Piste_Marq
        For i = 0 To (CnbCtrl - 1)
            If Contrl(piste, debut, i).ctrl = Rien Then
                Contrl(piste, debut, i).canal = canal
                Contrl(piste, debut, i).pos = debut
                Contrl(piste, debut, i).ctrl = ctrl
                Contrl(piste, debut, i).valeur = valeur
                Exit For
            End If
        Next
    End Sub
    Public Sub AddProgram(piste As Byte, canal As Byte, debut As Integer, prg As Byte)
        piste = piste + 1 'Variation_Piste_Marq
        Program(piste, debut).prog = prg
        Program(piste, debut).canal = canal
        Program(piste, debut).pos = debut
    End Sub
    Public Sub AddAfterT(piste As Byte, canal As Byte, debut As Integer, prg As Byte)
        piste = piste + 1 'Variation_Piste_Marq
        AfterT(piste, debut).aftert = prg
        AfterT(piste, debut).canal = canal
        AfterT(piste, debut).pos = debut
    End Sub
    Public Sub AddTempo(tmp As Integer)
        Pistes(0).Tempo = tmp
    End Sub
    Public Sub AddMétrique(mtr As String)
        Pistes(0).Métrique = mtr
    End Sub
    Public Sub AddNomFichier(NomFichier As String)
        Pistes(0).Nom = Trim(NomFichier)
    End Sub
    Public Sub AddNomPiste(piste As Byte, nompiste As String)
        piste = piste + 1 'Variation_Piste_Marq
        Pistes(piste).Nom = Trim(nompiste)
    End Sub
    Public Sub AddMarqueur(Marqueur As String, debut As Integer)
        Marq(debut) = Trim(Marqueur)
    End Sub
    Public Sub AddTexte(piste As Byte, MidText As String, debut As Integer)
        piste = piste + 1 'Variation_Piste_Marq
        MidiTexte(piste, debut) = Trim(MidText)
    End Sub
    Public Sub ConstruiredMidFile()
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim m As Integer
        Dim c As Byte
        Dim Tempo As Double
        Dim a As String
        Dim deltime As UInteger
        Dim deltimeStr As String
        Dim B1 As String
        Dim B2 As String
        Dim B3 As String

        Try
            '
            ' Mse à jour de la table de calcul des deltatime en nombre de divisions ( Dim Deltatime(-1, -1) As Delt ' (N° de pistes, nombre divisions))
            ' *****************************************************************************************************************************************
            Maj_DeltaTimeDiv2()
            ' Entête
            ' ******
            a = "4D 54 68 64" ' MThd
            Ecrire(a)
            ' Header chunk : MF Type, Nb Piste, Résolution
            ' ********************************************
            a = "00 00 00 06" ' longueur Header chunk
            Ecrire(a)
            ' MF ype
            a = Hex(MFType) ' MFType NbPistes Réso
            a = ConstruireChaine(a, 2)
            Ecrire(a)
            '
            ' Nombre de pistes
            a = Hex(NbPistes)
            a = ConstruireChaine(a, 2)
            Ecrire(a)
            '
            ' Résolution
            a = Hex(Réso)
            a = ConstruireChaine(a, 2)
            Ecrire(a)
            '
            ' Pistes
            ' ******
            For i = 0 To NbPistes - 1
                Select Case i
                    Case 0 ' traitement de la PISTE SYSTEME
                        a = "4D 54 72 6B" ' MTrk
                        Ecrire(a)
                        a = "00 00 00 00" ' initialisation de la longueur
                        Ecrire(a)
                        IndexLongueur = UBound(TabFichierBin) ' position du dernier octet d'une piste en cours de construction dans le fichier binaire
                        ValLongueur = 0 ' variable de comptage du nombre d'octets de la piste en cours de création
                        ' NOM DU FICHIER
                        a = "00 FF 03" 'delta time du nom de la piste + Meta event nom de piste
                        Ecrire(a)
                        '
                        a = Hex(Len(Pistes(0).Nom)) ' écriture de la longueur du nom de la piste
                        a = ConstruireChaine(a, 1) ' on considère que le nombre de caratère d'un om de piste tient sur un octet
                        Ecrire(a)
                        '
                        For m = 1 To Len(Pistes(0).Nom) 'écriture du nom de la piste
                            c = Asc(Mid(Pistes(0).Nom, m, 1))
                            a = Hex(c)
                            Ecrire(a)
                        Next
                        ' TEMPO
                        a = "00 FF 51 03" 'delta time du tempo + Meta event tempo + longueur
                        Ecrire(a)
                        Tempo = (1 / (Val(Pistes(0).Tempo) / 60)) * 1000000
                        a = Hex(CUInt(Tempo))
                        a = ConstruireChaine(a, 3)
                        Ecrire(a)
                        ' Métrique
                        a = "00 FF 58 04" 'delta time métrique + Meta event signature + longueur
                        Ecrire(a)
                        ' la métrique est écrite "en dur" : 4/4, métronome = 1 click/par noire = 24(H18) ,quadruple croches/signal horloge =8
                        'a = "04 02 18 08" ' 4 est le numérateur ; 02 est puissance de 2 indiquant le numérateur 4; le 3e byte indique un click métronome par noire; le 4e byte est constant (8)
                        a = Det_ChaineMétrique(Pistes(0).Métrique)
                        Ecrire(a)
                        ' Fermeture de la piste système : deltatime + meta event + ecriture longeur
                        EcrireLongueurChunk()
                    Case Else ' piste marqueurs ou pistes de notes (non système)

                        a = "4D 54 72 6B" ' MTrk
                        Ecrire(a)
                        a = "00 00 00 00" ' initialisation de la longueur de la piste - la longueur de la piste en octets ne peut être déterminée qu'à partir du moment où l'écriture de la piste est achvée
                        Ecrire(a)
                        IndexLongueur = UBound(TabFichierBin) ' position du dernier octet d'une piste en cours de construction dans le fichier binaire --> ici c'est l'index du dernier des 4 octets d'une longeur de piste
                        ValLongueur = 0 ' variable de comptage du nombre d'octets de la piste en cours de création
                        '
                        a = "00 FF 03" 'delta time du nom de la piste + Meta event nom de piste
                        Ecrire(a)
                        '
                        a = Hex(Len(Pistes(i).Nom)) ' écriture de la longueur du nom de la piste
                        a = ConstruireChaine(a, 1) ' on considère que le nombre de caratères d'un nom de piste tient dans un octet
                        Ecrire(a)
                        '
                        For m = 1 To Len(Pistes(i).Nom) 'écriture du nom de la piste
                            c = Asc(Mid(Pistes(i).Nom, m, 1))
                            a = Hex(c)
                            Ecrire(a)
                        Next
                        '
                        For j = 0 To NbDivisions - 1
                            If Deltatime(i, j).PrésenceEvent = True Then
                                deltime = Deltatime(i, j).Delta * (Réso / 4) ' calcul du deta Time en nombre de ticks
                                a = Hex(deltaTimeCompact(deltime))
                                deltimeStr = ConstruireChaine(a, LongueurDeltaTime(a))
                                ' PROGRAMME
                                If Program(i, j).prog <> Rien Then
                                    Ecrire(deltimeStr) ' ecriture du delta time
                                    B1 = Hex(CProgram Or Program(i, j).canal)
                                    '
                                    B2 = Hex(Program(i, j).prog)
                                    B2 = ConstruireChaine(B2, 1)
                                    '
                                    a = (B1 + " " + B2)
                                    Ecrire(a)
                                    deltimeStr = "00"
                                End If
                                '
                                ' AFTER TOUCH
                                'If AfterT(i, j).aftert <> Rien Then
                                'Ecrire(deltimeStr) ' ecriture du delta time
                                'B1 = Hex(CAfterTouch Or AfterT(i, j).canal)
                                ''
                                'B2 = Hex(AfterT(i, j).aftert)
                                'B2 = ConstruireChaine(B2, 1)
                                ''
                                'a = (B1 + " " + B2)
                                'Ecrire(a)
                                'deltimeStr = "00"
                                'End If
                                ' CONTROLEUR
                                For k = 0 To CnbCtrl - 1
                                    If Contrl(i, j, k).ctrl <> Rien Then
                                        Ecrire(deltimeStr) ' ecriture du delta time
                                        B1 = Hex(CController Or Contrl(i, j, k).canal)
                                        '
                                        B2 = Hex(Contrl(i, j, k).ctrl)
                                        B2 = ConstruireChaine(B2, 1)
                                        '
                                        B3 = Hex(Contrl(i, j, k).valeur)
                                        B3 = ConstruireChaine(B3, 1)
                                        '
                                        a = (B1 + " " + B2 + " " + B3)
                                        Ecrire(a)
                                        deltimeStr = "00"
                                    End If
                                Next k
                                'NOTEOFF
                                For k = 0 To NbNotes - 1
                                    If NoteOff(i, j, k).note <> Rien Then
                                        Ecrire(deltimeStr) ' ecriture du delta time
                                        B1 = Hex(CnoteOff Or NoteOff(i, j, k).canal)
                                        '
                                        B2 = Hex(NoteOff(i, j, k).note)
                                        B2 = ConstruireChaine(B2, 1)
                                        '
                                        B3 = Hex(NoteOff(i, j, k).velo)
                                        B3 = ConstruireChaine(B3, 1)
                                        '
                                        a = (B1 + " " + B2 + " " + B3)
                                        Ecrire(a)
                                        deltimeStr = "00"
                                    End If
                                Next k
                                'NOTEON
                                For k = 0 To NbNotes - 1
                                    If NoteOn(i, j, k).note <> Rien Then
                                        Ecrire(deltimeStr) ' ecriture du delta time
                                        B1 = Hex(CnoteOn Or NoteOn(i, j, k).canal)
                                        '
                                        B2 = Hex(NoteOn(i, j, k).note)
                                        B2 = ConstruireChaine(B2, 1)
                                        '
                                        B3 = Hex(NoteOn(i, j, k).velo)
                                        B3 = ConstruireChaine(B3, 1)
                                        '
                                        a = (B1 + " " + B2 + " " + B3)
                                        Ecrire(a)
                                        deltimeStr = "00"
                                    End If
                                Next k

                                ' TEXTE
                                If Trim(MidiTexte(i, j)) <> "" Then
                                    Ecrire(deltimeStr) ' ecriture du delta time
                                    a = "FF 01" 'Meta event texte
                                    Ecrire(a)
                                    '
                                    a = Len(MidiTexte(i, j)) ' écriture de la longueur du nom du texte
                                    a = Hex(a)
                                    a = ConstruireChaine(a, 1)
                                    Ecrire(a)
                                    '
                                    'For m = 1 To Len(MidiTexte(i, j)) 'écriture du nom de la piste
                                    'c = Asc(Mid(MidiTexte(i, j), m, 1))
                                    'a = Hex(c)
                                    'a = ConstruireChaine(a, 1)
                                    'Ecrire(a)
                                    'Next m
                                    deltimeStr = "00"
                                End If
                                ' MARQUEURS
                                If i = NbPistes - 1 And Trim(Marq(j)) <> "" Then ' la dernière piste est la piste des marqueurs
                                    Ecrire(deltimeStr) ' ecriture du delta time
                                    a = "FF 06" 'Meta event marqueur
                                    Ecrire(a)
                                    '
                                    a = Len(Marq(j)) ' écriture de la longueur du nom du Marqueur
                                    a = Hex(a)
                                    a = ConstruireChaine(a, 1)
                                    Ecrire(a)
                                    '
                                    For m = 1 To Len(Marq(j)) ' écriture du nom de la piste
                                        c = Asc(Mid(Marq(j), m, 1))
                                        a = Hex(c)
                                        a = ConstruireChaine(a, 1)
                                        Ecrire(a)
                                    Next
                                    deltimeStr = "00"
                                End If
                            End If
                        Next j
                        '  fermeture de la piste : deltatime + meta event + ecriture longeur
                        EcrireLongueurChunk()
                End Select
            Next i
            '
            a = Création_CTemp() + "\" + "djgbv58147.mid" 'Pistes(0).Nom + ".mid" ' \FichierMIDI.mid" My.Application.Info.DirectoryPath
            My.Computer.FileSystem.WriteAllBytes(a, TabFichierBin, False)

            '
        Catch ex As Exception

            If LangueIHM = "fr" Then
                MessageHV.PTitre = "Avertissement"
                MessageHV.PContenuMess = "Détection d'un erreur dans procédure : " + "ConstruireMidFile" + Constants.vbCrLf + "Message  : " + ex.Message + "Valeur de i : " + Str(i)
                MessageHV.PTypBouton = "OK"
            Else
                MessageHV.PTitre = "Warning"
                MessageHV.PContenuMess = "Détection d'un erreur dans procédure : " + "ConstruireMidFile" + Constants.vbCrLf + "Message  : " + ex.Message
                MessageHV.PTypBouton = "OK"
            End If
            Cacher_FormTransparents()
            MessageHV.ShowDialog()
            End
        End Try
    End Sub
    Sub EcrireLongueurChunk()
        Dim a As String
        Dim tbl()

        a = "00 FF 2F 00"
        Ecrire(a)
        ' - écriture de la longueur du chunk
        a = Hex(ValLongueur)
        a = ConstruireChaine(a, 4)
        tbl = Split(a)
        TabFichierBin(IndexLongueur - 3) = ValeurByte(tbl(0))
        TabFichierBin(IndexLongueur - 2) = ValeurByte(tbl(1))
        TabFichierBin(IndexLongueur - 1) = ValeurByte(tbl(2))
        TabFichierBin(IndexLongueur) = ValeurByte(tbl(3))
    End Sub

    Function LongueurDeltaTime(a As String) As Byte
        Dim i As Byte
        LongueurDeltaTime = 0
        i = Len(a)
        '
        Select Case i
            Case 1, 2
                LongueurDeltaTime = 1
            Case 3, 4
                LongueurDeltaTime = 2
            Case 5, 6
                LongueurDeltaTime = 3
        End Select

    End Function
    Sub Maj_DeltaTimeDiv2()
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer

        ' Mise à jour de la table DetaTime
        ' Dim Deltatime(-1, -1) As Delt ' (N° de pistes, nombre divisions)
        ' Structure Delt
        '    Dim PrésenceEvent As Boolean
        '    Dim Delta As Integer ' delta time en nombre de double croches
        ' End Structure


        ' postionnement des Note On - Note off
        ' ***********************************
        For i = 0 To NbPistes - 1
            For j = 0 To NbDivisions - 1
                For k = 0 To NbNotes - 1
                    If NoteOn(i, j, k).note <> Rien Then
                        Deltatime(i, j).PrésenceEvent = True ' la valeur Delta as integer est mise à jour plus loin - on note juste la présence d'evts ici.
                    End If
                    'Else
                    If NoteOff(i, j, k).note <> Rien Then
                        Deltatime(i, j).PrésenceEvent = True
                    End If
                    'End If
                Next k
            Next j
        Next i
        '
        ' CONTROLEURS
        ' ***********
        For i = 0 To NbPistes - 1
            For j = 0 To NbDivisions - 1
                For k = 0 To CnbCtrl - 1
                    If Contrl(i, j, k).ctrl <> Rien Then
                        Deltatime(i, j).PrésenceEvent = True
                    End If
                Next k
            Next j
        Next i
        '
        ' PROGRAMMES
        ' **********
        For i = 0 To NbPistes - 1
            For j = 0 To NbDivisions - 1
                If Program(i, j).prog <> Rien Then
                    Deltatime(i, j).PrésenceEvent = True
                End If
            Next j
        Next i

        '
        ' MARQUEURS
        ' *********

        For j = 0 To NbDivisions - 1
            If Trim(Marq(j)) <> "" Then
                Deltatime(NbPistes - 1, j).PrésenceEvent = True
            End If
        Next j

        ' TEXTE
        ' *****
        For i = 0 To NbPistes - 1
            For j = 0 To NbDivisions - 1
                If Trim(MidiTexte(i, j)) <> "" Then
                    Deltatime(i, j).PrésenceEvent = True
                End If
            Next j
        Next i
        ' Déermination des Delta time en nombre de divisions
        ' **************************************************
        For i = 0 To NbPistes - 1
            For j = 0 To NbDivisions - 1
                If Deltatime(i, j).PrésenceEvent = True Then
                    Deltatime(i, j).Delta = DeltaDiv(i, j) ' DeltaDiv est une fonction qui calcule le deltatime en nombre de divisions
                End If
            Next j
        Next i
    End Sub

    Sub Maj_DeltaTimeDiv()
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer

        ' Mse à jour delta time en nombre de divisions
        ' ********************************************
        '
        ' NOTEON et NOTEOFF
        ' *****************
        For i = 0 To NbPistes - 1
            For j = 0 To NbDivisions - 1
                For k = 0 To NbNotes - 1
                    If NoteOn(i, j, k).note <> Rien Then
                        Deltatime(i, j).Delta = DeltaDiv(i, j) ' DeltaDiv est une fonction qui calcule le deltatime en nombre de divisions
                        Deltatime(i, j).PrésenceEvent = True
                    Else
                        If NoteOff(i, j, k).note <> Rien Then
                            Deltatime(i, j).Delta = DeltaDiv(i, j)
                            Deltatime(i, j).PrésenceEvent = True
                        End If
                    End If
                Next k
            Next j
        Next i
        '
        ' CONTROLEURS
        ' ***********
        For i = 0 To NbPistes - 1
            For j = 0 To NbDivisions - 1
                For k = 0 To CnbCtrl - 1
                    If Contrl(i, j, k).ctrl <> Rien Then
                        Deltatime(i, j).Delta = DeltaDiv(i, j)
                        Deltatime(i, j).PrésenceEvent = True
                    End If
                Next k
            Next j
        Next i
        '
        ' PROGRAMMES
        ' **********
        For i = 0 To NbPistes - 1
            For j = 0 To NbDivisions - 1
                If Program(i, j).prog <> Rien Then
                    Deltatime(i, j).Delta = DeltaDiv(i, j)
                    Deltatime(i, j).PrésenceEvent = True
                End If
            Next j
        Next i

        '
        ' MARQUEURS
        ' *********

        For j = 0 To NbDivisions - 1
            If Marq(j) <> "" Then
                Deltatime(1, j).Delta = DeltaDiv(1, j)
                Deltatime(1, j).PrésenceEvent = True
            End If
        Next j

        ' TEXTE
        ' *****
        For i = 0 To NbPistes - 1
            For j = 0 To NbDivisions - 1
                If MidiTexte(i, j) <> "" Then
                    Deltatime(i, j).Delta = DeltaDiv(i, j)
                    Deltatime(i, j).PrésenceEvent = True
                End If
            Next j
        Next i
    End Sub

    Function DeltaDiv(NPiste As Byte, Posit As Integer) As Integer
        Dim pos As Integer
        Dim compteur As UInteger
        pos = Posit
        compteur = 0
        If pos <> 0 Then
            'pos = pos - 1
            Do
                pos = pos - 1
                compteur = compteur + 1 ' compteur de divisions jusqu'au 1er Evt en amont
                DeltaDiv = pos
            Loop Until Deltatime(NPiste, pos).PrésenceEvent = True Or pos = 0
            DeltaDiv = compteur
        Else
            DeltaDiv = 0
        End If
    End Function

    Function deltaTimeCompact(DeltaVal As UInteger) As Integer
        Dim A As Integer
        Dim B As Integer
        Dim B1 As Integer
        Dim B2 As Integer
        Dim B3 As Integer
        Dim C As Integer
        Dim D As Integer
        '
        A = DeltaVal
        If DeltaVal <= 127 Then
            A = A And &H7F ' Mettre à 0 le bit 7 de l'octet de poids faible
            D = A
        End If
        If DeltaVal >= 128 And DeltaVal <= 16383 Then
            B = A And &H3F80 ' lire les bits de 7 à 13
            B = B << 1 ' décalage des bits de 7 à 13 de une fois vers la gauche
            C = A And &HC07F ' effacer les bits de 7 à 13 dans la valeur initiale
            D = B Or C ' replacer dans la valeur initiale les Bits de 7 à  13 décalés
            D = D And &HFF7F ' mettre à 0 le bit 7
            D = D Or &H8000 ' mettre à 1 le bit 15
        End If
        '
        If DeltaVal >= 16384 And DeltaVal <= 2097151 Then
            B1 = A And &H3F80 'lecture des bits 7 à 13
            B1 = B1 << 1 ' décalage de 1 vers la gauche des bits 7 à 13
            B2 = A And &H1FC000 ' lecture des bits 14 à 20
            B2 = B2 << 2 ' décaalage de 2 vers la gauche des Bits de 14 à 20
            C = A And &HFFC07F ' Effacer les bits de 7 à 13 dans la valeur initiale
            D = B1 Or C ' Replacerdans la valeur initiale les bits de 7 à 13 décalés de 1
            D = D And &HE03FFF ' Effacer les bits de 14 à 20 dans la valeur initiale
            D = B2 Or D ' Replacer dans la valeur initiale les bits de 14 à 20 décalés de 2
            D = D And &HFFFF7F ' Mettre à 0 le bit 7 de l'octet de poids faible
            D = D Or &H808000 ' Mettre à 1 le bit 7 de l'octet de poids fort
        End If
        '
        If DeltaVal >= 2097152 And DeltaVal <= 268435455 Then
            B1 = A And &H3F80 ' Lire les bits de 7 à 13
            B1 = B1 << 1 ' Décaler les bits de 7 à 13 de 1 fois vers la gauche
            B2 = A And &H1FC000 ' Lire les bits de 14 à 20
            B2 = B2 << 2 'Décaler les bits de 14 à 20 de 2 fois vers la gauche
            B3 = A And &HFE00000 ' Lire les bits de 21 à 27
            B3 = B3 << 3 ' Décaler les bits de 21 à 27 de 2 fois vers la gauche
            C = A And &HFFFFC07F ' Effacer les bits de 7 à 13 dans la valeur initiale
            D = B1 Or C ' Replacer dans la valeur initiale les bits de 7 à 13 décalés de 1
            D = D And &HFFE03FFF ' Effacer les bits de 14 à 20 dans la valeur initiale
            D = B2 Or D ' Replacer dans la valeur initiale les bits de 14 à 20 décalés de 2
            D = D And &HF01FFFFF ' Effacer les bits de 21 à 27 dans la valeur initiale
            D = B3 Or D ' Replacer dans la valeur initiale les bits de 21 à 27 décalés de 3
            D = D And &HFFFFFF7F ' Mettre à 0 le bit 7 de l'octet de poids faible
            D = D Or &H80808000 ' Mettre à 1 le bit 7 de l'octet de poids fort
        End If
        Return D
    End Function
    '
    Function ValeurByte(Valeur As String) As Byte
        Dim a As String
        Dim b As String
        Dim i As Byte
        Dim j As Byte

        a = Right(Valeur, 1)
        b = Left(Valeur, 1)
        '
        i = TradhHexa(a)
        j = TradhHexa(b)
        '
        ValeurByte = (j * 16) + i
    End Function
    Function TradhHexa(a As String) As Byte
        Dim i As Byte
        '
        Select Case a
            Case "A"
                i = 10
            Case "B"
                i = 11
            Case "C"
                i = 12
            Case "D"
                i = 13
            Case "E"
                i = 14
            Case "F"
                i = 15
            Case Else
                i = Val(a)
        End Select
        '
        TradhHexa = i
    End Function
    Sub EcritureOctet(a As String)
        '
        ReDim Preserve TabFichierBin(UBound(TabFichierBin, 1) + 1)
        TabFichierBin(UBound(TabFichierBin, 1)) = ValeurByte(a)
        '
        ValLongueur = ValLongueur + 1 ' comptage des octets écrit pour maj des longueur de pistes
        '
    End Sub
    Sub Ecrire(a As String)
        Dim Tbl()
        Tbl = Split(a)
        For i = 0 To UBound(Tbl)
            EcritureOctet(Tbl(i))
        Next
    End Sub
    Function ConstruireChaine(a As String, nboctets As Byte) As String
        Dim b As String
        b = RajouterDesSéparateurs(a)
        Select Case Len(Trim(b)) ' nombre de caractrèes dans a
            Case 1
                Select Case nboctets
                    Case 1
                        b = "0" + Trim(b)
                    Case 2
                        b = "00 0" + Trim(b)
                    Case 3
                        b = "00 00 0" + Trim(b)
                    Case 4
                        b = "00 00 00 0" + Trim(b)
                End Select
            Case 2 ' exemple : F0
                Select Case nboctets
                    Case 2
                        b = "00 " + Trim(b)
                    Case 3
                        b = "00 00 " + Trim(b)
                    Case 4
                        b = "00 00 00 " + Trim(b)
                End Select
            Case 4 ' exemple : 1 F0
                Select Case nboctets
                    Case 2
                        b = "0" + Trim(b)
                    Case 3
                        b = "00 0" + Trim(b)
                    Case 4
                        b = "00 00 0" + Trim(b)
                End Select
            Case 5 ' exemple 21 F0
                Select Case nboctets
                    Case 3
                        b = "00 " + Trim(b)
                    Case 4
                        b = "00 00 " + Trim(b)
                End Select
            Case 7 ' exemple   5 21 F0
                Select Case nboctets
                    Case 3
                        b = "0" + Trim(b)
                    Case 4
                        b = "00 0" + Trim(b)
                End Select
        End Select
        '
        ConstruireChaine = b
    End Function
    Function RajouterDesSéparateurs(a As String) As String
        Dim b As String
        Dim tbl(3) As String
        b = a
        Select Case Len(a) ' nombre de caractrèes dans a
            Case 3
                tbl(0) = Left(a, 1)
                tbl(1) = Right(a, 2)
                b = tbl(0) + " " + tbl(1)
            Case 4
                tbl(0) = Left(a, 2)
                tbl(1) = Right(a, 2)
                b = tbl(0) + " " + tbl(1)
            Case 5
                tbl(0) = Left(a, 1)
                tbl(1) = Mid(a, 2, 2)
                tbl(2) = Right(a, 2)
                b = tbl(0) + " " + tbl(1) + " " + tbl(2)
            Case 6
                tbl(0) = Left(a, 2)
                tbl(1) = Mid(a, 3, 2)
                tbl(2) = Right(a, 2)
                b = tbl(0) + " " + tbl(1) + " " + tbl(2)
        End Select
        RajouterDesSéparateurs = b
    End Function
    Function Det_ChaineMétrique(Métrique As String) As String
        Dim tbl() As String
        Dim Numérateur As String
        Dim Dénominateur As String


        Det_ChaineMétrique = "04 02 18 08" ' valeur par défaut avec métrique 4/4

        tbl = Split(Métrique, "/")
        '
        If Trim(tbl(0)) <> "12" Then
            Numérateur = "0" + Trim(tbl(0))
        Else
            Numérateur = "0C"
        End If
        '
        Dénominateur = "02"
        If Trim(tbl(1)) = "8" Then
            Dénominateur = "03"
        End If

        Det_ChaineMétrique = Trim(Numérateur) + " " + Trim(Dénominateur) + " 18 08"

    End Function
End Class
