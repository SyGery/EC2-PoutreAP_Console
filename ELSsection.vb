Imports System.Math
Module ELSsection

    'Données entrée
    Public et1, et, et2, et3, et4 As Double
    Public eMscar As Double
    Public eMsqp As Double
    Public Ms As Double
    Public Msa As Double
    Public RH As Double
    Public RapM As Double

    Public SstatNf As Double
    Public Sstatf As Double

    Public Asts As Double
    Public Ascs As Double

    'Fonction du signe du moment
    'Si section en Té, =true si compression dans la table
    '=false si traction dans la table 
    Public MomentTestELS As Boolean

    'Vérification des contraintes
    'Coef. armature
    Public ep3 As Double
    'Valeur admissible contrainte armature
    Public fyl As Double

    'Coef. béton
    Public ep1 As Double
    'Valeur admissible contrainte béton
    Public fcl As Double

    'Vérification de l'ouverture des fissures
    'Critère admissible d'ouverture de fissure
    Public ewkmax As Double

    'Durée de chargement - False=Courte durée
    Public eduree As Boolean
    'Coef. distance entre ouverture de fissure
    Public ek1w, ek2w, ek3w, ek4w As Double
    'enrobage des armatures longitudinales
    Public eenrobage As Double
    'Diamètre moyen des armatures longitudinales
    Public ePHImoyen As Double
    'Espacement entre armatures longitudinales
    Public eEspace As Double

    'Résultats vérification contrainte

    Public Ecef As Double
    Public neq As Double
    Public Phi As Double

    'Etat non fissurée
    Public Infis As Double
    Public Mcr As Double
    Public dg As Double
    Public EtatFis As Boolean
    Public scr As Double

    'Inertie fissurée
    Public Ifis As Double
    'Distance à l'axe neutre en partant du nu de la partie comprimée
    Public y As Double

    Public Ks As Double

    Public scb As Double
    Public sAsc As Double
    Public sAst As Double

    'Résultats ouverture fissures
    Public hcef As Double
    Public Srmax As Double
    Public Alldif As Double
    Public wk As Double

    'Résultats courbure
    Public r1 As Double
    Public dzeta As Double

    'resultat courbure retrait
    Public rcs As Double
    Public PriseRetrait As Boolean


    'coefficient de combinaison quasi permanent
    'Gestion erreurs
    Public errS As Integer
    Public warS As Integer
    Public Sub ContrainteELS(Phi1 As Double)

        'Call CaracAcier()
        'Call CaracBeton()
        'Call PropSection()

        If Ms >= 0 Then
            'Table de compression en fibre supérieure
            dt = edi
            dc = eds
            yg = yb
            Asts = Ainf
            Ascs = Asup

            MomentTestELS = True
        Else
            'Table de compression en fibre inférieure
            dt = eds
            dc = edi
            yg = eHt - yb

            Asts = Asup
            Ascs = Ainf
            MomentTestELS = False
        End If

        d = eHt - dt

        'Paramètre fluage
        'pour déterminer le module et le coefficient d'équivalence
        'Phi = Fluage(RH, et1, et, h0)
        Ecef = Ecm / (1 + Phi1)
        neq = Eacier / Ecef

        Call NonFissuration()

        'Comparaison Moment ELS avec Moment de première fissuration
        If Mcr > Msa Then
            EtatFis = False
            Ks = Msa / Infis
            scb = Ks * dg
            sAsc = Ks * neq * (dg - dt)
            sAst = -Ks * neq * (d - dg)
        Else
            EtatFis = True

            'Discussion en fonction de la forme de la section
            If Te = False Then
                Call ELSAxeNeutreRect()
                Ks = Msa / Ifis
                scb = Ks * y
                sAsc = Ks * neq * (y - dt)
                sAst = -Ks * neq * (d - y)
            Else
                Call ELSAxeNeutreTe()
                Ks = Msa / Ifis
                scb = Ks * y
                sAsc = Ks * neq * (y - dt)
                sAst = -Ks * neq * (d - y)
            End If
        End If

        If Ascs = 0 Then
            sAsc = 0
        End If


    End Sub
    Public Sub VerifContrainte()

        'Limite de contrainte des armatures de traction
        fyl = -ep3 * efyk

        'Limite de contrainte du béton
        fcl = ep1 * fck

        'Vérification des contraintes de l'acier et du béton
        errS = 0
        If fyl > sAst Then
            errS = 1
        End If

        If fcl < scb Then
            errS = errS + 2
        End If

    End Sub
    Public Sub Ouverture()

        Dim r As Double
        Dim s As Double
        Dim t As Double
        Dim Acef As Double
        Dim rocef As Double

        Dim sr1 As Double
        Dim sr2 As Double

        Dim kt As Double
        Dim alldif1 As Double
        Dim alldif2 As Double

        If EtatFis = False Then
            wk = 0
            Exit Sub
        End If

        'Distance max entre fissures, sr,max
        r = 2.5 * (eHt - d)
        s = (eHt - y) / 3
        t = eHt / 2

        'Minimum de r,s et t
        If r <= s Then
            hcef = r
        Else
            hcef = s
        End If

        If hcef > t Then
            hcef = t
        End If

        'Attention; on considère que la section tendue est dans l'ame de la poutre même en section en té
        'A détailler
        Acef = hcef * eBw
        rocef = Asts / Acef

        sr1 = ek3w * eenrobage + ek1w * ek2w * ek4w * ePHImoyen / rocef
        sr2 = 1.3 * (eHt - y)

        If eEspace - ePHImoyen <= 5 * (eenrobage + ePHImoyen / 2) Then
            Srmax = sr1
        Else
            Srmax = sr2
        End If

        'allongement differentiel

        If eduree = False Then
            kt = 0.4
        Else
            kt = 0.6
        End If

        alldif1 = (-sAst - kt * fctm / rocef * (1 + neq * rocef)) / Eacier
        alldif2 = 0.6 * -sAst / Eacier

        If alldif1 >= alldif2 Then
            Alldif = alldif1
        Else
            Alldif = alldif2
        End If

        wk = Srmax * Alldif

        If wk > ewkmax Then
            errS = errS + 4
        End If

    End Sub
    Public Sub NonFissuration()

        Dim Ahom As Double
        Dim Mstat As Double

        If Te = False Then
            'Section rectangulaire
            'Position axe neutre
            'Aire section homogène
            Ahom = eBw * eHt + neq * Asts + neq * Ascs
            'Moment statique
            Mstat = eBw * eHt ^ 2 / 2 + neq * Ascs * dc + neq * Asts * d
            'Axe neutre
            dg = Mstat / Ahom
            'Inertie non fissurée de la section homogénéisée
            Infis = eBw * dg ^ 3 / 3 + eBw * (eHt - dg) ^ 3 / 3 + neq * Asts * (eHt - dg - dt) ^ 2 + neq * Ascs * (dg - dc) ^ 2

            'Moment critique de fissuration
            Mcr = fctm * Infis / (eHt - dg)
            'Contrainte traction armature
            scr = Mcr / Infis * neq * (d - dg)



        Else
            ' Section Té
            Ahom = eBw * eHt + (ebf - eBw) * ehf + neq * Asts + neq * Ascs
            Mstat = eBw * eHt ^ 2 / 2 + (ebf - eBw) * ehf ^ 2 / 2 + neq * Asts * d + neq * Ascs * dc

            dg = Mstat / Ahom
            Infis = ebf * dg ^ 3 / 3 - (ebf - eBw) * (dg - ehf) ^ 3 / 3 + eBw * (eHt - dg) ^ 3 / 3 + neq * Asts * (d - dg) ^ 2 + neq * Ascs * (dg - dc) ^ 2
            Mcr = fctm * Infis / (eHt - dg)
            scr = Mcr / Infis * neq * (d - dg)

        End If

        SstatNf = Asts * (d - dg)
    End Sub

    Public Sub courbure(M)

        Dim beta As Double

        If eduree = False Then
            beta = 0.5
        Else
            beta = 1
        End If

        dzeta = 1 - beta * (scr / sAst) ^ 2


        If EtatFis = True Then
            r1 = dzeta * M / Ecef / Ifis + (1 - dzeta) * M / Ecef / Infis
        Else
            r1 = M / Ecef / Infis
        End If

    End Sub


    Private Sub ELSAxeNeutreRect()

        Dim r As Double
        Dim s As Double
        Dim t As Double

        'Position de l'axe neutre de la section homogénéisée fissurée
        r = eBw / 2
        s = neq * Ascs + neq * Asts
        t = -neq * Ascs * dt - neq * Asts * d

        delta = s ^ 2 - 4 * r * t
        y = (-s + delta ^ 0.5) / 2 / r

        'Inertie fissurée
        Ifis = eBw * y ^ 3 / 3 + neq * Ascs * (y - dt) ^ 2 + neq * Asts * (d - y) ^ 2

        Sstatf = Asts * (d - y)



    End Sub

    Private Sub ELSAxeNeutreTe()

        Dim k As Double
        Dim r As Double
        Dim s As Double
        Dim t As Double

        k = ebf * ehf ^ 2 / 2 + neq * Ascs * (ehf - dt) - neq * Asts * (d - ehf)

        If MomentTestELS = True Then

            If k <= 0 Then
                'Position de l'axe neutre de la section homogénéisée fissurée
                'dans la nervure
                r = eBw / 2
                s = (ebf - eBw) * ehf + neq * Ascs + neq * Asts
                t = -(ebf - eBw) * ehf ^ 2 / 2 - neq * Ascs * dt - neq * Asts * d

                delta = s ^ 2 - 4 * r * t
                y = (-s + delta ^ 0.5) / 2 / r

                'Inertie fissurée
                Ifis = -(ebf - eBw) * (y - ehf) ^ 3 / 3 + ebf * y ^ 3 / 3 + neq * Ascs * (y - dt) ^ 2 + neq * Asts * (d - y) ^ 2

            Else
                'Position de l'axe neutre de la section homogénéisée fissurée
                'dans la table. Poutre rectangulaire ebf*eHt
                r = ebf / 2
                s = neq * Ascs + neq * Asts
                t = -neq * Ascs * dt - neq * Asts * d

                delta = s ^ 2 - 4 * r * t
                y = (-s + delta ^ 0.5) / 2 / r

                'Inertie fissurée
                Ifis = ebf * y ^ 3 / 3 + neq * Ascs * (y - dt) ^ 2 + neq * Asts * (d - y) ^ 2
            End If
        Else
            'Te inversé équivalent à une section rectangulaire eBw*Ht
            r = eBw / 2
            s = neq * Ascs + neq * Asts
            t = -neq * Ascs * dt - neq * Asts * d

            delta = s ^ 2 - 4 * r * t
            y = (-s + delta ^ 0.5) / 2 / r

            'Inertie fissurée
            Ifis = eBw * y ^ 3 / 3 + neq * Ascs * (y - dt) ^ 2 + neq * Asts * (d - y) ^ 2
        End If

        Sstatf = Asts * (d - y)

    End Sub

    Public Sub coubureretrait()

        Dim beta As Double

        If eduree = False Then
            beta = 0.5
        Else
            beta = 1
        End If

        dzeta = 1 - beta * (scr / sAst) ^ 2

        If EtatFis = True Then
            rcs = dzeta * Retrait(RH, ts, et, h0) * neq * Sstatf / Ifis + (1 - dzeta) * Retrait(RH, ts, et, h0) * neq * SstatNf / Infis
        Else
            rcs = Retrait(RH, ts, et, h0) * neq * SstatNf / Infis
        End If

    End Sub

    Public Sub ContrainteFleche(M As Double, phi As Double)

        If M >= 0 Then
            'Table de compression en fibre supérieure
            dt = edi
            dc = eds
            yg = yb
            Asts = Ainf
            Ascs = Asup

            MomentTestELS = True
        Else
            'Table de compression en fibre inférieure
            dt = eds
            dc = edi
            yg = eHt - yb

            Asts = Asup
            Ascs = Ainf
            MomentTestELS = False
        End If

        d = eHt - dt

        'Paramètre fluage
        'pour déterminer le module et le coefficient d'équivalence
        Ecef = Ecm / (1 + phi)
        neq = Eacier / Ecef

        Call NonFissuration()

        'Comparaison Moment ELS avec Moment de première fissuration
        If Mcr > M Then
            EtatFis = False
            Ks = M / Infis
            scb = Ks * dg
            sAsc = Ks * neq * (dg - dt)
            sAst = -Ks * neq * (d - dg)
        Else
            EtatFis = True
            'Discussion en fonction de la forme de la section
            If Te = False Then
                Call ELSAxeNeutreRect()
                Ks = M / Ifis
                scb = Ks * y
                sAsc = Ks * neq * (y - dt)
                sAst = -Ks * neq * (d - y)
            Else
                Call ELSAxeNeutreTe()
                Ks = M / Ifis
                scb = Ks * y
                sAsc = Ks * neq * (y - dt)
                sAst = -Ks * neq * (d - y)
            End If
        End If

        If Ascs = 0 Then
            sAsc = 0
        End If

    End Sub

End Module
