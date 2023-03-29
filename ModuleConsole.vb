﻿Imports System.Math
Imports System.Text


Module ModuleConsole

    Public Sub Main(ByVal CmdArgs() As String)

        '---------------------------------------------------------------
        'Chemin d'accès des fichiers
        Dim FileInput As String
        Dim FileOutput As String
        FileInput = "D:\2-Boite à outils de calcul\5-Poutre\EC2AP-IN\230326\Input.txt"
        FileOutput = "D:\2-Boite à outils de calcul\5-Poutre\EC2AP-IN\230326\Output.txt"
        If CmdArgs.Length = 1 Then
            FileInput = CmdArgs(0)
        End If
        If CmdArgs.Length = 2 Then
            FileInput = CmdArgs(0)
            FileOutput = CmdArgs(1)
        End If

        '----------------------------------------------------------------
        'Lecture du fichier entree
        Dim I As Integer = 0
        Try
            Dim SR As StreamReader = New StreamReader(FileInput) 'Stream pour la lecture
            Dim ligne As String ' Variable contenant le texte de la ligne
            Do
                I += 1
                ligne = SR.ReadLine()
                'Geometrie poutre
                If I = 8 Then eBw = CDbl(ligne)
                If I = 10 Then eHt = CDbl(ligne)
                If I = 12 Then Te = ligne
                If I = 14 Then ebf = CDbl(ligne)
                If I = 16 Then ehf = CDbl(ligne)
                'Dispo ferraillage horizontal
                If I = 19 Then eds = CDbl(ligne) / 100
                If I = 21 Then edi = CDbl(ligne) / 100
                If I = 23 Then eenrobage = CDbl(ligne) / 1000
                If I = 25 Then ePHImoyen = CDbl(ligne) / 1000
                If I = 27 Then eEspace = CDbl(ligne) / 1000
                'Geometrie travee
                If I = 29 Then travee1.L = CDbl(ligne)
                If I = 31 Then travee1.Ag_L = CDbl(ligne)
                If I = 33 Then travee1.Ad_L = CDbl(ligne)
                'Qualité materiaux
                If I = 36 Then Beton = ligne
                If I = 38 Then efyk = CDbl(ligne)
                If I = 40 Then ClasseAcier = ligne
                'Continuité
                If I = 44 Then travee1.AGType = ligne
                If I = 46 Then ContAppW = CDbl(ligne)
                If I = 48 Then travee1.ADType = ligne
                If I = 50 Then ContAppE = CDbl(ligne)
                If I = 52 Then ContTr = CDbl(ligne)
                'Chargement
                'Option poids propre poutre
                If I = 55 Then PPOption = ligne
                'Charge surfacique 1
                If I = 59 Then travee1.d1_cotes = CDbl(ligne)
                If I = 61 Then travee1.d1_L = CDbl(ligne)
                If I = 63 Then travee1.d1_g = CDbl(ligne)
                If I = 65 Then travee1.d1_gp1 = CDbl(ligne)
                If I = 67 Then travee1.d1_gp2 = CDbl(ligne)
                If I = 69 Then travee1.d1_q = CDbl(ligne)
                'Charge surfacique 2
                If I = 73 Then travee1.d2_cotes = CDbl(ligne)
                If I = 75 Then travee1.d2_L = CDbl(ligne)
                If I = 77 Then travee1.d2_g = CDbl(ligne)
                If I = 79 Then travee1.d2_gp1 = CDbl(ligne)
                If I = 81 Then travee1.d2_gp2 = CDbl(ligne)
                If I = 83 Then travee1.d2_q = CDbl(ligne)
                'Chargement ponctuel
                If I = 87 Then iponc = CDbl(ligne)
                ReDim Preserve travee1.a(iponc - 1)
                ReDim Preserve travee1.g(iponc - 1)
                ReDim Preserve travee1.gp1(iponc - 1)
                ReDim Preserve travee1.gp2(iponc - 1)
                ReDim Preserve travee1.q(iponc - 1)
                ReDim Preserve travee1.R(iponc - 1)
                ReDim Preserve travee1.Pelu(iponc - 1)
                travee1.p_nbre = iponc
                Select Case iponc
                    Case = 0
                        Exit Select
                    Case = 1
                        If I = 90 Then travee1.a(0) = CDbl(ligne)
                        If I = 92 Then travee1.g(0) = CDbl(ligne)
                        If I = 94 Then travee1.gp1(0) = CDbl(ligne)
                        If I = 96 Then travee1.gp2(0) = CDbl(ligne)
                        If I = 98 Then travee1.q(0) = CDbl(ligne)
                        If I = 100 Then travee1.R(0) = ligne
                    Case = 2
                        If I = 90 Then travee1.a(0) = CDbl(ligne)
                        If I = 92 Then travee1.g(0) = CDbl(ligne)
                        If I = 94 Then travee1.gp1(0) = CDbl(ligne)
                        If I = 96 Then travee1.gp2(0) = CDbl(ligne)
                        If I = 98 Then travee1.q(0) = CDbl(ligne)
                        If I = 100 Then travee1.R(0) = ligne
                        If I = 104 Then travee1.a(1) = CDbl(ligne)
                        If I = 106 Then travee1.g(1) = CDbl(ligne)
                        If I = 108 Then travee1.gp1(1) = CDbl(ligne)
                        If I = 110 Then travee1.gp2(1) = CDbl(ligne)
                        If I = 112 Then travee1.q(1) = CDbl(ligne)
                        If I = 114 Then travee1.R(1) = ligne
                    Case = 3
                        If I = 90 Then travee1.a(0) = CDbl(ligne)
                        If I = 92 Then travee1.g(0) = CDbl(ligne)
                        If I = 94 Then travee1.gp1(0) = CDbl(ligne)
                        If I = 96 Then travee1.gp2(0) = CDbl(ligne)
                        If I = 98 Then travee1.q(0) = CDbl(ligne)
                        If I = 100 Then travee1.R(0) = ligne
                        If I = 104 Then travee1.a(1) = CDbl(ligne)
                        If I = 106 Then travee1.g(1) = CDbl(ligne)
                        If I = 108 Then travee1.gp1(1) = CDbl(ligne)
                        If I = 110 Then travee1.gp2(1) = CDbl(ligne)
                        If I = 112 Then travee1.q(1) = CDbl(ligne)
                        If I = 114 Then travee1.R(1) = ligne
                        If I = 118 Then travee1.a(2) = CDbl(ligne)
                        If I = 120 Then travee1.g(2) = CDbl(ligne)
                        If I = 122 Then travee1.gp1(2) = CDbl(ligne)
                        If I = 124 Then travee1.gp2(2) = CDbl(ligne)
                        If I = 126 Then travee1.q(2) = CDbl(ligne)
                        If I = 128 Then travee1.R(2) = ligne
                    Case = 4
                        If I = 90 Then travee1.a(0) = CDbl(ligne)
                        If I = 92 Then travee1.g(0) = CDbl(ligne)
                        If I = 94 Then travee1.gp1(0) = CDbl(ligne)
                        If I = 96 Then travee1.gp2(0) = CDbl(ligne)
                        If I = 98 Then travee1.q(0) = CDbl(ligne)
                        If I = 100 Then travee1.R(0) = ligne
                        If I = 104 Then travee1.a(1) = CDbl(ligne)
                        If I = 106 Then travee1.g(1) = CDbl(ligne)
                        If I = 108 Then travee1.gp1(1) = CDbl(ligne)
                        If I = 110 Then travee1.gp2(1) = CDbl(ligne)
                        If I = 112 Then travee1.q(1) = CDbl(ligne)
                        If I = 114 Then travee1.R(1) = ligne
                        If I = 118 Then travee1.a(2) = CDbl(ligne)
                        If I = 120 Then travee1.g(2) = CDbl(ligne)
                        If I = 122 Then travee1.gp1(2) = CDbl(ligne)
                        If I = 124 Then travee1.gp2(2) = CDbl(ligne)
                        If I = 126 Then travee1.q(2) = CDbl(ligne)
                        If I = 128 Then travee1.R(2) = ligne
                        If I = 132 Then travee1.a(3) = CDbl(ligne)
                        If I = 134 Then travee1.g(3) = CDbl(ligne)
                        If I = 136 Then travee1.gp1(3) = CDbl(ligne)
                        If I = 138 Then travee1.gp2(3) = CDbl(ligne)
                        If I = 140 Then travee1.q(3) = CDbl(ligne)
                        If I = 142 Then travee1.R(3) = ligne
                    Case = 5
                        If I = 90 Then travee1.a(0) = CDbl(ligne)
                        If I = 92 Then travee1.g(0) = CDbl(ligne)
                        If I = 94 Then travee1.gp1(0) = CDbl(ligne)
                        If I = 96 Then travee1.gp2(0) = CDbl(ligne)
                        If I = 98 Then travee1.q(0) = CDbl(ligne)
                        If I = 100 Then travee1.R(0) = ligne
                        If I = 104 Then travee1.a(1) = CDbl(ligne)
                        If I = 106 Then travee1.g(1) = CDbl(ligne)
                        If I = 108 Then travee1.gp1(1) = CDbl(ligne)
                        If I = 110 Then travee1.gp2(1) = CDbl(ligne)
                        If I = 112 Then travee1.q(1) = CDbl(ligne)
                        If I = 114 Then travee1.R(1) = ligne
                        If I = 118 Then travee1.a(2) = CDbl(ligne)
                        If I = 120 Then travee1.g(2) = CDbl(ligne)
                        If I = 122 Then travee1.gp1(2) = CDbl(ligne)
                        If I = 124 Then travee1.gp2(2) = CDbl(ligne)
                        If I = 126 Then travee1.q(2) = CDbl(ligne)
                        If I = 128 Then travee1.R(2) = ligne
                        If I = 132 Then travee1.a(3) = CDbl(ligne)
                        If I = 134 Then travee1.g(3) = CDbl(ligne)
                        If I = 136 Then travee1.gp1(3) = CDbl(ligne)
                        If I = 138 Then travee1.gp2(3) = CDbl(ligne)
                        If I = 140 Then travee1.q(3) = CDbl(ligne)
                        If I = 142 Then travee1.R(3) = ligne
                        If I = 146 Then travee1.a(4) = CDbl(ligne)
                        If I = 148 Then travee1.g(4) = CDbl(ligne)
                        If I = 150 Then travee1.gp1(4) = CDbl(ligne)
                        If I = 152 Then travee1.gp2(4) = CDbl(ligne)
                        If I = 154 Then travee1.q(4) = CDbl(ligne)
                        If I = 156 Then travee1.R(4) = ligne

                End Select
                If I > 200 Then Exit Do

            Loop Until ligne Is Nothing
            SR.Close()
        Catch ex As Exception
            MsgBox("Une erreur est survenue au cours de la lecture du fichier." & vbCrLf & vbCrLf & "Veuillez vérifier l'emplacement : " & FileInput, MsgBoxStyle.Critical, "Erreur lors de l'ouverture du fichier")
        End Try

        '-----------------------------------------------------------------------
        'Paramètres par défaut
        'Materiaux
        'Beton = "C30/37"
        'ClasseAcier = "A"
        egb = 1.5
        acc = 1.0
        'efyk = 500
        egs = 1.15
        'Coefficient de continuite
        'travee1.AGType = True : travee1.ADType = True
        If travee1.AGType = True Then
            travee1.cg = False
        Else
            travee1.cg = True
        End If
        If travee1.ADType = True Then
            travee1.cd = False
        Else
            travee1.cd = True
        End If
        'travee1.cd = 0
        'ContAppW = 0.15
        'ContTr = 1.0
        'ContAppE = 0.15
        'Coef chargement
        coefqp = 0.3
        gpu = 1.35
        gqu = 1.5
        ciment = "N"
        'Absence de chargement ponctuel ou lineaire
        'iponc = 0
        ilin = 0
        'Paramètre Hygrometrie
        RH = 60
        'Parametre ouvertures des fissures
        ek1w = 0.8
        ek2w = 0.5
        ek3w = 3.4
        ek4w = 0.425
        ewkmax = 0.3 / 1000
        'entrer du paramètre de limite de contrainte en traction des armatures sous ELS caractéristique
        ep3 = 0.8
        'entrer du paramètre de limite de contrainte max du béton sous ELS caractéristique
        ep1 = 1.0
        'Parametre ELS - coefficient d'equivalence
        ocoeq = True
        'Parametre d'age de chargement pour le calcul de fleche
        et1 = 10
        et2 = 60
        et3 = 60
        et4 = 365
        et = 3650
        ts = 3
        'entrer durée de chargement
        eduree = False
        'effort tranchant
        k1T = 0.15
        'combobox ratio
        ratiotype = False
        'checkbox option retrait/iterations
        PriseRetrait = False
        IteratFleche = True
        IteratContrainte = True

        '------------------------------------------------------------------------
        'Progamme
        Call Calcul()

        '------------------------------------------------------------------------
        'Ecriture des resultats

        Dim SW As New StreamWriter(FileOutput)
        SW.WriteLine("***Fichier Sortie - EC2-PoutreAP v1.2.XX***")
        SW.WriteLine("Fichier entree:" & FileInput)
        SW.Write(Chr(13) + Chr(10))

        SW.WriteLine("-----------------------------------------------------------")
        SW.WriteLine("DONNEES")
        SW.WriteLine("-----------------------------------------------------------")
        SW.WriteLine("Largeur de la poutre (m)")
        SW.WriteLine(eBw)
        SW.WriteLine("Hauteur de la poutre (m)")
        SW.WriteLine(eHt)
        SW.WriteLine("Forme de la section en Te:")
        SW.WriteLine(Te)
        SW.WriteLine("Largeur table compression = (m)")
        SW.WriteLine(ebf)
        SW.WriteLine("Hauteur table compression = (m)")
        SW.WriteLine(ehf)
        SW.Write(Chr(13) + Chr(10))
        SW.WriteLine("Distance Armature superieure = (cm)")
        SW.WriteLine(eds * 100)
        SW.WriteLine("Distance Armature inferieure = (cm)")
        SW.WriteLine(edi * 100)
        SW.WriteLine("Enrobage = (mm)")
        SW.WriteLine(eenrobage * 1000)
        SW.WriteLine("Diametre moyen de l'armature longitudinale = (mm)")
        SW.WriteLine(ePHImoyen * 1000)
        SW.WriteLine("Espacement armature longitudinale = (mm)")
        SW.WriteLine(eEspace * 1000)

        'Geometrie de la travee
        SW.WriteLine("Portee entre nus = (m)")
        SW.WriteLine(travee1.L)
        SW.WriteLine("Profondeur appui gauche = (m)")
        SW.WriteLine(travee1.Ag_L)
        SW.WriteLine("Profondeur appui droit = (m)")
        SW.WriteLine(travee1.Ad_L)
        SW.Write(Chr(13) + Chr(10))

        'Matériaux
        SW.WriteLine("Qualite de beton = (Cx/x selon EC2)")
        SW.WriteLine(Beton)
        SW.WriteLine("Limite d'elasticite de l'armature = (MPa)")
        SW.WriteLine(efyk)
        SW.WriteLine("Classe de l'acier = (A, B ou C)")
        SW.WriteLine(ClasseAcier)
        SW.Write(Chr(13) + Chr(10))

        'Continuité
        SW.WriteLine("Continuite:")
        SW.WriteLine("Pas de continuite sur l'appui gauche = (True/False)")
        SW.WriteLine(travee1.AGType)
        SW.WriteLine("Coefficient de continuite à l'appui gauche = (-)")
        SW.WriteLine(ContAppW)
        SW.WriteLine("Pas de continuite sur l'appui droit = (True/False)")
        SW.WriteLine(travee1.ADType)
        SW.WriteLine("Coefficient de continuite à l'appui droit = (-)")
        SW.WriteLine(ContAppE)
        SW.WriteLine("Coefficient de continuite en travee = (-)")
        SW.WriteLine(ContTr)
        SW.Write(Chr(13) + Chr(10))

        'Chargement poids propre poutre
        SW.WriteLine("Poids propre poutre :")
        SW.WriteLine(PPOption)
        SW.Write(Chr(13) + Chr(10))

        'Chargement surfacique dalle 1
        SW.WriteLine("Chargement surfacique Dalle 1")
        SW.WriteLine("Type dalle 1 =")
        SW.WriteLine(travee1.d1_cotes)
        SW.WriteLine("Portee dalle 1 = (m)")
        SW.WriteLine(travee1.d1_L)
        SW.WriteLine("Poids propre dalle 1= (kPa)")
        SW.WriteLine(travee1.d1_g)
        SW.WriteLine("Charge permanente dalle 1 = (kPa)")
        SW.WriteLine(travee1.d1_gp1)
        SW.WriteLine("Charge permanente dalle 1 = (kPa)")
        SW.WriteLine(travee1.d1_gp2)
        SW.WriteLine("Surcharge exploitation dalle 1 = (kPa)")
        SW.WriteLine(travee1.d1_q)
        SW.Write(Chr(13) + Chr(10))

        'Chargement surfacique dalle 2
        SW.WriteLine("Chargement surfacique Dalle 2")
        SW.WriteLine("Type dalle 2 =")
        SW.WriteLine(travee1.d2_cotes)
        SW.WriteLine("Portee dalle 2 = (m)")
        SW.WriteLine(travee1.d2_L)
        SW.WriteLine("Poids propre dalle 2= (kPa)")
        SW.WriteLine(travee1.d2_g)
        SW.WriteLine("Charge permanente dalle 2 = (kPa)")
        SW.WriteLine(travee1.d2_gp1)
        SW.WriteLine("Charge permanente dalle 2 = (kPa)")
        SW.WriteLine(travee1.d2_gp2)
        SW.WriteLine("Surcharge exploitation dalle 2 = (kPa)")
        SW.WriteLine(travee1.d2_q)
        SW.Write(Chr(13) + Chr(10))

        'Chargement ponctuel
        SW.WriteLine("Chargement ponctuel")
        SW.WriteLine("Nombre de charges ponctuelles (limite à 5)")
        SW.WriteLine(iponc)
        Select Case iponc
            Case = 0
                Exit Select
            Case = 1
                SW.WriteLine("Charge n1")
                SW.WriteLine("Distance par rapport au nu de la poutre= (m)")
                SW.WriteLine(travee1.a(0))
                SW.WriteLine("Charge Poids propre = (kN)")
                SW.WriteLine(travee1.g(0))
                SW.WriteLine("Charge permanente 1 = (kN)")
                SW.WriteLine(travee1.gp1(0))
                SW.WriteLine("Charge permanente 2 = (kN)")
                SW.WriteLine(travee1.gp2(0))
                SW.WriteLine("Surcharge exploitation = (kN)")
                SW.WriteLine(travee1.q(0))
                SW.WriteLine("Charge relevee = (True/False)")
                SW.WriteLine(travee1.R(0))
                SW.Write(Chr(13) + Chr(10))
            Case = 2
                SW.WriteLine("Charge n1")
                SW.WriteLine("Distance par rapport au nu de la poutre= (m)")
                SW.WriteLine(travee1.a(0))
                SW.WriteLine("Charge Poids propre = (kN)")
                SW.WriteLine(travee1.g(0))
                SW.WriteLine("Charge permanente 1 = (kN)")
                SW.WriteLine(travee1.gp1(0))
                SW.WriteLine("Charge permanente 2 = (kN)")
                SW.WriteLine(travee1.gp2(0))
                SW.WriteLine("Surcharge exploitation = (kN)")
                SW.WriteLine(travee1.q(0))
                SW.WriteLine("Charge relevee = (True/False)")
                SW.WriteLine(travee1.R(0))
                SW.Write(Chr(13) + Chr(10))
                SW.WriteLine("Charge n2")
                SW.WriteLine("Distance par rapport au nu de la poutre= (m)")
                SW.WriteLine(travee1.a(1))
                SW.WriteLine("Charge Poids propre = (kN)")
                SW.WriteLine(travee1.g(1))
                SW.WriteLine("Charge permanente 1 = (kN)")
                SW.WriteLine(travee1.gp1(1))
                SW.WriteLine("Charge permanente 2 = (kN)")
                SW.WriteLine(travee1.gp2(1))
                SW.WriteLine("Surcharge exploitation = (kN)")
                SW.WriteLine(travee1.q(1))
                SW.WriteLine("Charge relevee = (True/False)")
                SW.WriteLine(travee1.R(1))
                SW.Write(Chr(13) + Chr(10))
            Case = 3
                SW.WriteLine("Charge n1")
                SW.WriteLine("Distance par rapport au nu de la poutre= (m)")
                SW.WriteLine(travee1.a(0))
                SW.WriteLine("Charge Poids propre = (kN)")
                SW.WriteLine(travee1.g(0))
                SW.WriteLine("Charge permanente 1 = (kN)")
                SW.WriteLine(travee1.gp1(0))
                SW.WriteLine("Charge permanente 2 = (kN)")
                SW.WriteLine(travee1.gp2(0))
                SW.WriteLine("Surcharge exploitation = (kN)")
                SW.WriteLine(travee1.q(0))
                SW.WriteLine("Charge relevee = (True/False)")
                SW.WriteLine(travee1.R(0))
                SW.Write(Chr(13) + Chr(10))
                SW.WriteLine("Charge n2")
                SW.WriteLine("Distance par rapport au nu de la poutre= (m)")
                SW.WriteLine(travee1.a(1))
                SW.WriteLine("Charge Poids propre = (kN)")
                SW.WriteLine(travee1.g(1))
                SW.WriteLine("Charge permanente 1 = (kN)")
                SW.WriteLine(travee1.gp1(1))
                SW.WriteLine("Charge permanente 2 = (kN)")
                SW.WriteLine(travee1.gp2(1))
                SW.WriteLine("Surcharge exploitation = (kN)")
                SW.WriteLine(travee1.q(1))
                SW.WriteLine("Charge relevee = (True/False)")
                SW.WriteLine(travee1.R(1))
                SW.Write(Chr(13) + Chr(10))
                SW.WriteLine("Charge n3")
                SW.WriteLine("Distance par rapport au nu de la poutre= (m)")
                SW.WriteLine(travee1.a(2))
                SW.WriteLine("Charge Poids propre = (kN)")
                SW.WriteLine(travee1.g(2))
                SW.WriteLine("Charge permanente 1 = (kN)")
                SW.WriteLine(travee1.gp1(2))
                SW.WriteLine("Charge permanente 2 = (kN)")
                SW.WriteLine(travee1.gp2(2))
                SW.WriteLine("Surcharge exploitation = (kN)")
                SW.WriteLine(travee1.q(2))
                SW.WriteLine("Charge relevee = (True/False)")
                SW.WriteLine(travee1.R(2))
                SW.Write(Chr(13) + Chr(10))
            Case = 4
                SW.WriteLine("Charge n1")
                SW.WriteLine("Distance par rapport au nu de la poutre= (m)")
                SW.WriteLine(travee1.a(0))
                SW.WriteLine("Charge Poids propre = (kN)")
                SW.WriteLine(travee1.g(0))
                SW.WriteLine("Charge permanente 1 = (kN)")
                SW.WriteLine(travee1.gp1(0))
                SW.WriteLine("Charge permanente 2 = (kN)")
                SW.WriteLine(travee1.gp2(0))
                SW.WriteLine("Surcharge exploitation = (kN)")
                SW.WriteLine(travee1.q(0))
                SW.WriteLine("Charge relevee = (True/False)")
                SW.WriteLine(travee1.R(0))
                SW.Write(Chr(13) + Chr(10))
                SW.WriteLine("Charge n2")
                SW.WriteLine("Distance par rapport au nu de la poutre= (m)")
                SW.WriteLine(travee1.a(1))
                SW.WriteLine("Charge Poids propre = (kN)")
                SW.WriteLine(travee1.g(1))
                SW.WriteLine("Charge permanente 1 = (kN)")
                SW.WriteLine(travee1.gp1(1))
                SW.WriteLine("Charge permanente 2 = (kN)")
                SW.WriteLine(travee1.gp2(1))
                SW.WriteLine("Surcharge exploitation = (kN)")
                SW.WriteLine(travee1.q(1))
                SW.WriteLine("Charge relevee = (True/False)")
                SW.WriteLine(travee1.R(1))
                SW.Write(Chr(13) + Chr(10))
                SW.WriteLine("Charge n3")
                SW.WriteLine("Distance par rapport au nu de la poutre= (m)")
                SW.WriteLine(travee1.a(2))
                SW.WriteLine("Charge Poids propre = (kN)")
                SW.WriteLine(travee1.g(2))
                SW.WriteLine("Charge permanente 1 = (kN)")
                SW.WriteLine(travee1.gp1(2))
                SW.WriteLine("Charge permanente 2 = (kN)")
                SW.WriteLine(travee1.gp2(2))
                SW.WriteLine("Surcharge exploitation = (kN)")
                SW.WriteLine(travee1.q(2))
                SW.WriteLine("Charge relevee = (True/False)")
                SW.WriteLine(travee1.R(2))
                SW.Write(Chr(13) + Chr(10))
                SW.WriteLine("Charge n4")
                SW.WriteLine("Distance par rapport au nu de la poutre= (m)")
                SW.WriteLine(travee1.a(3))
                SW.WriteLine("Charge Poids propre = (kN)")
                SW.WriteLine(travee1.g(3))
                SW.WriteLine("Charge permanente 1 = (kN)")
                SW.WriteLine(travee1.gp1(3))
                SW.WriteLine("Charge permanente 2 = (kN)")
                SW.WriteLine(travee1.gp2(3))
                SW.WriteLine("Surcharge exploitation = (kN)")
                SW.WriteLine(travee1.q(3))
                SW.WriteLine("Charge relevee = (True/False)")
                SW.WriteLine(travee1.R(3))
                SW.Write(Chr(13) + Chr(10))
            Case = 5
                SW.WriteLine("Charge n1")
                SW.WriteLine("Distance par rapport au nu de la poutre= (m)")
                SW.WriteLine(travee1.a(0))
                SW.WriteLine("Charge Poids propre = (kN)")
                SW.WriteLine(travee1.g(0))
                SW.WriteLine("Charge permanente 1 = (kN)")
                SW.WriteLine(travee1.gp1(0))
                SW.WriteLine("Charge permanente 2 = (kN)")
                SW.WriteLine(travee1.gp2(0))
                SW.WriteLine("Surcharge exploitation = (kN)")
                SW.WriteLine(travee1.q(0))
                SW.WriteLine("Charge relevee = (True/False)")
                SW.WriteLine(travee1.R(0))
                SW.Write(Chr(13) + Chr(10))
                SW.WriteLine("Charge n2")
                SW.WriteLine("Distance par rapport au nu de la poutre= (m)")
                SW.WriteLine(travee1.a(1))
                SW.WriteLine("Charge Poids propre = (kN)")
                SW.WriteLine(travee1.g(1))
                SW.WriteLine("Charge permanente 1 = (kN)")
                SW.WriteLine(travee1.gp1(1))
                SW.WriteLine("Charge permanente 2 = (kN)")
                SW.WriteLine(travee1.gp2(1))
                SW.WriteLine("Surcharge exploitation = (kN)")
                SW.WriteLine(travee1.q(1))
                SW.WriteLine("Charge relevee = (True/False)")
                SW.WriteLine(travee1.R(1))
                SW.Write(Chr(13) + Chr(10))
                SW.WriteLine("Charge n3")
                SW.WriteLine("Distance par rapport au nu de la poutre= (m)")
                SW.WriteLine(travee1.a(2))
                SW.WriteLine("Charge Poids propre = (kN)")
                SW.WriteLine(travee1.g(2))
                SW.WriteLine("Charge permanente 1 = (kN)")
                SW.WriteLine(travee1.gp1(2))
                SW.WriteLine("Charge permanente 2 = (kN)")
                SW.WriteLine(travee1.gp2(2))
                SW.WriteLine("Surcharge exploitation = (kN)")
                SW.WriteLine(travee1.q(2))
                SW.WriteLine("Charge relevee = (True/False)")
                SW.WriteLine(travee1.R(2))
                SW.Write(Chr(13) + Chr(10))
                SW.WriteLine("Charge n4")
                SW.WriteLine("Distance par rapport au nu de la poutre= (m)")
                SW.WriteLine(travee1.a(3))
                SW.WriteLine("Charge Poids propre = (kN)")
                SW.WriteLine(travee1.g(3))
                SW.WriteLine("Charge permanente 1 = (kN)")
                SW.WriteLine(travee1.gp1(3))
                SW.WriteLine("Charge permanente 2 = (kN)")
                SW.WriteLine(travee1.gp2(3))
                SW.WriteLine("Surcharge exploitation = (kN)")
                SW.WriteLine(travee1.q(3))
                SW.WriteLine("Charge relevee = (True/False)")
                SW.WriteLine(travee1.R(3))
                SW.Write(Chr(13) + Chr(10))
                SW.WriteLine("Charge n5")
                SW.WriteLine("Distance par rapport au nu de la poutre= (m)")
                SW.WriteLine(travee1.a(4))
                SW.WriteLine("Charge Poids propre = (kN)")
                SW.WriteLine(travee1.g(4))
                SW.WriteLine("Charge permanente 1 = (kN)")
                SW.WriteLine(travee1.gp1(4))
                SW.WriteLine("Charge permanente 2 = (kN)")
                SW.WriteLine(travee1.gp2(4))
                SW.WriteLine("Surcharge exploitation = (kN)")
                SW.WriteLine(travee1.q(4))
                SW.WriteLine("Charge relevee = (True/False)")
                SW.WriteLine(travee1.R(4))
                SW.Write(Chr(13) + Chr(10))
        End Select
        'Sorties des resultats
        SW.WriteLine("-----------------------------------------------------------")
        SW.WriteLine("RESULTATS")
        SW.WriteLine("-----------------------------------------------------------")
        SW.Write(Chr(13) + Chr(10))
        SW.WriteLine("Erreur > 0 => Non valide")
        SW.WriteLine("Erreur= ")
        SW.WriteLine(err)

        SW.Write(Chr(13) + Chr(10))

        SW.WriteLine("Moment travee ELU = " & Round(travee1.Mhyp, 0) & " kN.m")
        SW.WriteLine("Aciers inf travee = (cm2)")
        SW.WriteLine(Round(travee1.MtrAinf, 1))
        SW.WriteLine("Aciers sup travee = " & Round(travee1.MtrAsup, 1) & " cm2")
        SW.WriteLine(Round(travee1.MtrAsup, 1))

        SW.Write(Chr(13) + Chr(10))

        SW.WriteLine("Moment appui gauche ELU = " & Round(travee1.MtAppG, 0) & " kN.m")

        SW.WriteLine("Aciers inf appui gauche = " & Round(travee1.MtAppGAinf, 1) & "cm2")
        SW.WriteLine("Aciers sup appui gauche = " & Round(travee1.MtAppGAsup, 1) & " cm2")
        SW.WriteLine(Round(travee1.MtAppGAsup, 1))
        SW.WriteLine("Reaction appui gauche ELU = " & Round(travee1.RAppG, 0) & " kN")
        SW.WriteLine("Contrainte tangente appui gauche = " & Round(Abs(travee1.RAppG) * 10 ^ (-3) / eBw / d, 1) & "MPa")

        SW.Write(Chr(13) + Chr(10))

        SW.WriteLine("Moment appui droit ELU = " & Round(travee1.MtAppD, 0) & " kN.m")
        SW.WriteLine("Aciers inf appui droit = " & Round(travee1.MtAppDAinf, 1) & " cm2")
        SW.WriteLine("Aciers sup appui droit = " & Round(travee1.MtAppDAsup, 1) & " cm2")
        SW.WriteLine(Round(travee1.MtAppDAsup, 1))
        SW.WriteLine("Reaction appui droit ELU = " & Round(travee1.RAppD, 0) & " kN")
        SW.WriteLine("Contrainte tangente appui droit = " & Round(Abs(travee1.RAppD) * 10 ^ (-3) / eBw / d, 1) & " MPa")

        SW.Write(Chr(13) + Chr(10))

        SW.WriteLine("Justification contraintes ELS caracteristique :")
        SW.WriteLine("Contraine beton = " & Round(travee1.sCmax, 1) & " MPa")
        SW.WriteLine("Contraine armature = " & Round(travee1.sAmax, 0) & " MPa")

        SW.Write(Chr(13) + Chr(10))

        SW.WriteLine("Justification contraintes ELS quasi permanent :")
        SW.WriteLine("Ouverture des fissures max = (mm)")
        SW.WriteLine(Round(travee1.wmax * 10 ^ 3, 3))
        SW.WriteLine("Fleche totale = " & Round(travee1.ftot * 10 ^ 3, 1) & " < " & Round(travee1.ftotadm * 10 ^ 3, 1) & " mm")
        SW.WriteLine(Round(travee1.ftot * 10 ^ 3, 1))
        SW.WriteLine("Fleche nuisible = " & Round(travee1.fnui * 10 ^ 3, 1) & " < " & Round(travee1.fnuiadm * 10 ^ 3, 1) & " mm")
        SW.WriteLine(Round(travee1.fnui * 10 ^ 3, 1))
        SW.Write(Chr(13) + Chr(10))

        SW.WriteLine("Ratio d'armature :")
        SW.WriteLine("Ratio optimum = (kg/m3)")
        SW.WriteLine(Round(travee1.ratio1, 0))
        SW.WriteLine("Ratio maximum = (kg/m3)")
        SW.WriteLine(Round(travee1.ratio2, 0))

        SW.Close()

    End Sub

End Module
