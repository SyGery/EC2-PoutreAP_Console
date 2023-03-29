Imports System.Math
Module materiaux

    'Declaration Entree Beton
    Public fck, fcd, fyd As Double
    Public fckcube, fcm, fctm, Ecm, Ec, ec1, ecu1, ec2, ecu2, ec3, ecu3 As Double
    Public Beton As String
    Public egb As Double
    Public n As Double
    Public acc As Double
    Public ciment As String
    'Declaration Entree Acier
    Public efyk As Double
    Public k, euk, eud, ek As Double
    Public Const Eacier As Double = 200000
    Public ClasseAcier As String
    Public egs As Double
    'retrait
    Public ts As Double

    Public Sub CaracBeton()
        If Beton = "C12/15" Then
            fck = 12
            fckcube = 15
            fctm = 1.6
            Ecm = 27000
            ec1 = 1.8 * 10 ^ (-3)
            ecu1 = 3.5 * 10 ^ (-3)
            ec2 = 2.0# * 10 ^ (-3)
            ecu2 = 3.5 * 10 ^ (-3)
            ec3 = 1.75 * 10 ^ (-3)
            ecu3 = 3.5 * 10 ^ (-3)
            n = 2


        End If
        If Beton = "C16/20" Then
            fck = 16
            fckcube = 20
            fctm = 1.9
            Ecm = 29000
            ec1 = 1.9 * 10 ^ (-3)
            ecu1 = 3.5 * 10 ^ (-3)
            ec2 = 2.0# * 10 ^ (-3)
            ecu2 = 3.5 * 10 ^ (-3)
            ec3 = 1.75 * 10 ^ (-3)
            ecu3 = 3.5 * 10 ^ (-3)
            n = 2

        End If
        If Beton = "C20/25" Then
            fck = 20
            fckcube = 25
            fctm = 2.2
            Ecm = 30000
            ec1 = 2.0# * 10 ^ (-3)
            ecu1 = 3.5 * 10 ^ (-3)
            ec2 = 2.0# * 10 ^ (-3)
            ecu2 = 3.5 * 10 ^ (-3)
            ec3 = 1.75 * 10 ^ (-3)
            ecu3 = 3.5 * 10 ^ (-3)
            n = 2

        End If
        If Beton = "C25/30" Then
            fck = 25
            fckcube = 30
            fctm = 2.6
            Ecm = 31000
            ec1 = 2.1 * 10 ^ (-3)
            ecu1 = 3.5 * 10 ^ (-3)
            ec2 = 2.0# * 10 ^ (-3)
            ecu2 = 3.5 * 10 ^ (-3)
            ec3 = 1.75 * 10 ^ (-3)
            ecu3 = 3.5 * 10 ^ (-3)
            n = 2

        End If
        If Beton = "C30/37" Then
            fck = 30
            fckcube = 37
            fctm = 2.9
            Ecm = 33000
            ec1 = 2.2 * 10 ^ (-3)
            ecu1 = 3.5 * 10 ^ (-3)
            ec2 = 2.0# * 10 ^ (-3)
            ecu2 = 3.5 * 10 ^ (-3)
            ec3 = 1.75 * 10 ^ (-3)
            ecu3 = 3.5 * 10 ^ (-3)
            n = 2

        End If

        If Beton = "C35/45" Then
            fck = 35
            fckcube = 45
            fctm = 3.2
            Ecm = 34000
            ec1 = 2.25 * 10 ^ (-3)
            ecu1 = 3.5 * 10 ^ (-3)
            ec2 = 2.0# * 10 ^ (-3)
            ecu2 = 3.5 * 10 ^ (-3)
            ec3 = 1.75 * 10 ^ (-3)
            ecu3 = 3.5 * 10 ^ (-3)
            n = 2

        End If

        If Beton = "C40/50" Then
            fck = 40
            fckcube = 50
            fctm = 3.5
            Ecm = 35000
            ec1 = 2.3 * 10 ^ (-3)
            ecu1 = 3.5 * 10 ^ (-3)
            ec2 = 2.0# * 10 ^ (-3)
            ecu2 = 3.5 * 10 ^ (-3)
            ec3 = 1.75 * 10 ^ (-3)
            ecu3 = 3.5 * 10 ^ (-3)
            n = 2

        End If

        If Beton = "C45/55" Then
            fck = 45
            fckcube = 55
            fctm = 3.8
            Ecm = 36000
            ec1 = 2.4 * 10 ^ (-3)
            ecu1 = 3.5 * 10 ^ (-3)
            ec2 = 2.0# * 10 ^ (-3)
            ecu2 = 3.5 * 10 ^ (-3)
            ec3 = 1.75 * 10 ^ (-3)
            ecu3 = 3.5 * 10 ^ (-3)
            n = 2

        End If
        If Beton = "C50/60" Then
            fck = 50
            fckcube = 60
            fctm = 4.1
            Ecm = 37000
            ec1 = 2.45 * 10 ^ (-3)
            ecu1 = 3.5 * 10 ^ (-3)
            ec2 = 2.0# * 10 ^ (-3)
            ecu2 = 3.5 * 10 ^ (-3)
            ec3 = 1.75 * 10 ^ (-3)
            ecu3 = 3.5 * 10 ^ (-3)
            n = 2

        End If
        If Beton = "C55/67" Then
            fck = 55
            fckcube = 67
            fctm = 4.2
            Ecm = 38000
            ec1 = 2.5 * 10 ^ (-3)
            ecu1 = 3.2 * 10 ^ (-3)
            ec2 = 2.2 * 10 ^ (-3)
            ecu2 = 3.1 * 10 ^ (-3)
            ec3 = 1.8 * 10 ^ (-3)
            ecu3 = 3.1 * 10 ^ (-3)
            n = 1.75

        End If
        If Beton = "C60/75" Then
            fck = 60
            fckcube = 75
            fctm = 4.4
            Ecm = 39000
            ec1 = 2.6 * 10 ^ (-3)
            ecu1 = 3.0# * 10 ^ (-3)
            ec2 = 2.3 * 10 ^ (-3)
            ecu2 = 2.9 * 10 ^ (-3)
            ec3 = 1.9 * 10 ^ (-3)
            ecu3 = 2.9 * 10 ^ (-3)
            n = 1.6

        End If
        If Beton = "C70/85" Then
            fck = 70
            fckcube = 85
            fctm = 4.6
            Ecm = 41000
            ec1 = 2.7 * 10 ^ (-3)
            ecu1 = 2.8 * 10 ^ (-3)
            ec2 = 2.4 * 10 ^ (-3)
            ecu2 = 2.7 * 10 ^ (-3)
            ec3 = 2.0# * 10 ^ (-3)
            ecu3 = 2.7 * 10 ^ (-3)
            n = 1.45

        End If
        If Beton = "C80/95" Then
            fck = 80
            fckcube = 95
            fctm = 4.8
            Ecm = 42000
            ec1 = 2.8 * 10 ^ (-3)
            ecu1 = 2.8 * 10 ^ (-3)
            ec2 = 2.5 * 10 ^ (-3)
            ecu2 = 2.6 * 10 ^ (-3)
            ec3 = 2.2 * 10 ^ (-3)
            ecu3 = 2.6 * 10 ^ (-3)
            n = 1.4

        End If

        If Beton = "C90/105" Then
            fck = 90
            fckcube = 105
            fctm = 5.0
            Ecm = 44000
            ec1 = 2.8 * 10 ^ (-3)
            ecu1 = 2.9 * 10 ^ (-3)
            ec2 = 2.6 * 10 ^ (-3)
            ecu2 = 2.6 * 10 ^ (-3)
            ec3 = 2.3 * 10 ^ (-3)
            ecu3 = 2.6 * 10 ^ (-3)
            n = 1.4

        End If

        fcm = fck + 8
        Ec = 1.05 * Ecm
        fcd = fck / egb * acc

    End Sub
    Public Sub CaracAcier()

        If ClasseAcier = "A" Then
            k = 1.05
            euk = -0.025
        End If
        If ClasseAcier = "B" Then
            k = 1.08
            euk = -0.05
        End If
        If ClasseAcier = "C" Then
            k = 1.15
            euk = -0.075
        End If

        fyd = efyk / egs
        eud = 0.9 * euk
        ek = -fyd / Eacier


    End Sub


    Public Function Fluage(ByVal RH As Double, ByVal t1 As Double, ByVal t As Double, ByVal h0 As Double) As Double

        Dim Bt0 As Double
        Dim Bfcm As Double
        Dim PhiRH As Double
        Dim Bh As Double
        Dim Bc As Double

        Dim a1 As Double
        Dim a2 As Double
        Dim a3 As Double

        a1 = (35 / fcm) ^ (0.7)
        a2 = (35 / fcm) ^ (0.2)
        a3 = (35 / fcm) ^ (0.5)

        'Calcul de Bc(t,t0)
        If fcm <= 35 Then
            Bh = 1.5 * (1 + (0.012 * RH) ^ (18)) * h0 * 10 ^ 3 + 250
            If Bh > 1500 Then
                Bh = 1500
            End If
        Else
            Bh = 1.5 * (1 + (0.012 * RH) ^ (18)) * h0 * 10 ^ 3 + 250 * a3
            If Bh > 1500 * a3 Then
                Bh = 1500 * a3
            End If
        End If

        Bc = ((t - t1) / (Bh + t - t1)) ^ (0.3)

        'Calcul de Phi0
        Bt0 = (0.1 + t1 ^ (0.2)) ^ (-1)
        Bfcm = 16.8 / fcm ^ (0.5)

        If fcm <= 35 Then
            PhiRH = 1 + (1 - RH / 100) / (0.1 * (h0 * 10 ^ 3) ^ (1 / 3))
        Else
            PhiRH = (1 + (1 - RH / 100) / (0.1 * (h0 * 10 ^ 3) ^ (1 / 3)) * a1) * a2
        End If

        'Calcul de Phi à la date t
        'A partir de 10 ans, le fluage à l'infini est pris en compte
        'Le fluage n'est pas considéré non linéaire voir EC2 3.1.4
        'Les contrainte de compression dans une poutre en flexion simple est supposée inférieure à 0.45 fck(to)
        If t > 3649 Then
            Fluage = PhiRH * Bfcm * Bt0
        Else
            Fluage = PhiRH * Bfcm * Bt0 * Bc
        End If

    End Function

    Public Function Retrait(ByVal RH As Double, ByVal ts As Double, ByVal t As Double, ByVal h0 As Double) As Double

        Dim ecaf As Double
        Dim bas As Double
        Dim eca As Double

        Dim ads1 As Double
        Dim ads2 As Double
        Dim BRH As Double
        Dim ecd0 As Double
        Dim kh As Double
        Dim ecdf As Double
        Dim Bds As Double
        Dim ecd As Double

        ' Dim ecsf As Double
        'Dim ecs As Double


        'Retrait endogène
        ecaf = 2.5 * (fck - 10) * 10 ^ (-6)
        bas = 1 - Exp(-0.2 * t ^ (0.5))
        eca = bas * ecaf

        Select Case ciment
            Case Is = "S"
                ads1 = 3
                ads2 = 0.13
            Case Is = "N"
                ads1 = 4
                ads2 = 0.12
            Case Is = "R"
                ads1 = 6
                ads2 = 0.11
        End Select

        'Retrait de dessication
        BRH = 1.55 * (1 - (RH / 100) ^ 3)
        ecd0 = 0.85 * ((220 + 110 * ads1) * Exp(-ads2 * fcm / 10)) * 10 ^ (-6) * BRH

        'Kh §3.1.4 Tableau 3.3
        'Valeurs interpolées entre valeurs du tableau (soumis à interprétation)
        Select Case h0
            Case Is <= 100
                kh = 1.0#
            Case Is <= 200
                kh = 1.0# - 0.15 / 100 * (h0 - 100)
            Case Is <= 300
                kh = 0.85 - 0.1 / 100 * (h0 - 200)
            Case Is < 500
                kh = 0.75 - 0.05 / 200 * (h0 - 300)
            Case Is >= 500
                kh = 0.7
        End Select

        ecdf = kh * ecd0
        Bds = (t - ts) / (t - ts + 0.04 * h0 ^ (3 / 2))
        ecd = Bds * ecdf

        'Retrait total
        Retrait = eca + ecd

    End Function
End Module
