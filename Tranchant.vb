Imports System.Math
Module Tranchant


    'Déclaration variables externes au calcul de l'effort tranchant
    'Données d'entrée
    Public Ved As Double
    Public Ned As Double = 0

    Public SEd As Double
    Public TauEd As Double

    'Déclaration variables propre à l'effort tranchant

    'Section d'armature transversale (cadres+épingles+étriers..)
    Public Asw As Double
    'Espacement des aciers transversaux
    Public spacing As Double
    'Bras de levier calcul de section ELU
    Public z As Double
    'Résultat armature transversale m²/ml
    Public Asws As Double
    Public Aswsmin As Double

    'Déclaration Sollicitations

    'Contrainte compression sous ELU
    Public Scp As Double

    'Résistance au tranchant du béton seul
    Public France As Boolean = True
    Public Vrdc As Double
    Public Taurdc As Double
    Public Resist As String
    Public rol As Double
    Public Crdc As Double

    'Résistance au tranchant de l'armature
    Public Vrds As Double
    Public Taurds As Double
    Public n1 As Double

    'Résistance maximale pour vérifier la compression de la bielle
    Public Vrdmax As Double
    Public TauRdmax As Double

    'Résultats
    Public Vrd As Double
    Public Tranchant As Boolean

    'Paramétrage Eurocode 2
    Public k1T As Double
    Public vmin As Double
    Public Precontrainte As Boolean

    'Gestion des erreurs
    Public errT As Integer


    Sub VResistConcrete()

        Dim k As Double
        Dim ka As Double
        Dim TauRdmin As Double

        'Calcul VRdc EC2§6.2.2
        'k
        ka = 1 + (200 / (d * 10 ^ (3))) ^ 0.5
        If ka > 2 Then
            k = 2
        Else
            k = ka
        End If

        'rol, pourcentage d'armature
        rol = Ast / eBw / d
        If rol > 0.02 Then
            rol = 0.02
        End If
        'Crdc
        Crdc = 0.18 / egb

        'Sigma cp
        Scp = Ned / eBw / eHt

        'vmin
        If France = False Then
            vmin = 0.035 * k ^ (3 / 2) * fck ^ 0.5
        Else
            vmin = 0.053 / egb * k ^ (3 / 2) * fck ^ 0.5

        End If



        'TauRdmin/ TauRdc
        Select Case Scp
            Case Is >= 0

                If Scp >= 0.2 * fcd Then
                    'MsgBox("Compressive stress limited to 0.2 x fcd", vbInformation, "Warning")
                    Taurdc = Crdc * k * (100 * rol * fck) ^ (1 / 3) + k1T * 0.2 * fcd
                    TauRdmin = vmin + k1T * 0.2 * fcd

                Else

                    Taurdc = Crdc * k * (100 * rol * fck) ^ (1 / 3) + k1T * Scp
                    TauRdmin = vmin + k1T * Scp

                End If

            Case Is < 0
                Taurdc = 0
                TauRdmin = 0
        End Select

        If Taurdc < TauRdmin Then
            Taurdc = TauRdmin
        End If

        Vrdc = Taurdc * eBw * d

    End Sub

    Sub VresistRCCheck()

        'Calcul VRd,s EC2§6.2.3
        Asws = Asw / spacing
        Vrds = Asws * z * fyd
        Taurds = Vrds / eBw / d

    End Sub

    Sub VresistRCDim()

        Asws = Abs(Ved) / z / fyd
        Vrds = Abs(Ved)
        Taurds = Vrds / eBw / d

    End Sub

    Sub VresistRCmax()

        Dim aw As Double

        'Calcul Vrdmax EC2§6.2.3

        n1 = 0.6 * (1 - fck / 250)

        'Alpha cw en cas de compression traction
        If Ned < 0 Then
            aw = 1 + SEd / fctm
        Else : aw = 1
        End If

        Vrdmax = aw * eBw * z * n1 * fcd / 2
        TauRdmax = Vrdmax / eBw / d
    End Sub

    Sub MainDimTranchant()

        'Initialisation erreur
        errT = 0


        'Hauteur utile
        d = eHt - dt
        'Contrainte moyenne dans la section
        SEd = Ned / eBw / d

        'Contrainte tangente
        TauEd = Ved / eBw / d

        If SEd < -fctm Then
            errT = 1
            'MsgBox("Case not adressed in EC2: Normal stress higher than concrete tension capacity fctm", vbOKOnly, "Error")
            Exit Sub
        End If

        'Calcul de la resistance du béton
        Call VResistConcrete()
        Call VresistRCmax()

        If Abs(Ved) > Vrdmax Then
            errT = 2
            Exit Sub
        End If


        'Calcul armature minimale effort tranchant
        Aswsmin = eBw * 0.08 * (fck) ^ 0.5 / efyk

        'Détermination de Vrd
        If Abs(Ved) > Vrdc Then
            Resist = "Armature transversale requise"
        Else
            Resist = "Section béton suffisante"
            Vrd = Vrdc
            Asws = Aswsmin
            Vrds = 0
            Taurds = 0
            Exit Sub
        End If

        VresistRCDim()

        If Asws < Aswsmin Then
            Asws = Aswsmin
        End If

    End Sub

    Public Function AboutVerif(ByVal typ As Boolean, ByVal Vu As Double, ByVal bw As Double, ByVal a As Double, ByVal f As Double, ByVal g As Double) As Boolean

        'Contrainte de compression dans la bielle d'about
        Dim sig As Double
        sig = 2 * Abs(Vu) * 10 ^ (-3) / bw / a

        'Capacité de compression de la bielle d'un noeud soumis à compression et à traction §6.5.4
        Dim sigadm As Double
        sigadm = 0.85 * (1 - f / 250) * f / g

        If sig <= sigadm Then
            AboutVerif = True
        Else
            AboutVerif = False
        End If

    End Function

    Public Function GlissementVerif(ByVal typ As Boolean, ByVal Vu As Double, ByVal a As Double, ByVal f As Double, ByVal g As Double) As Boolean

        'Dimensionnement des aciers de glissement
        Dim az As Double
        az = Abs(Vu) * 10 ^ (-3) / 2 / f * g

        If az <= a Then
            GlissementVerif = True
        Else
            GlissementVerif = False
        End If

    End Function

End Module
