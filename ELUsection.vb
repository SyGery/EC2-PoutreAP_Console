Imports System.Math
Module ELUsection
    'Déclaration sorties
    Public Ast, Asc As Double
    Public dt, dc, yg, d As Double
    Public pivot As String

    Public coefqp As Double
    Public gpu As Double
    Public gqu As Double

    Public eMu As Double

    Public Mua As Double
    Public MA As Double
    Public MA1, MA2, MA3 As Double
    Public DIF As Double
    Public eb As Double
    Public est As Double
    Public esc As Double

    Public FAsc As Double
    Public FAst As Double
    Public Fb As Double

    Public mlim As Double
    Public zs As Double

    Public SigmaAst As Double
    Public SigmaAsc As Double
    Public SigmaA As Double
    Public epsA As Double
    Public SigmaBt As Double
    Public sigmaB As Double

    Public epsB As Double
    Public a As Double


    Public delta As Double
    Public ag As Double
    Public s1 As Double
    Public ContrainteBeton As Double
    Public ave1 As Double
    Public ave2 As Double

    Public Asmin As Double
    Public Asmax As Double
    Public Amin As Double

    Public Asup As Double
    Public Ainf As Double
    Public momentest As Double

    Public s2 As Double

    Public yy As Double
    Public l As Double
    Public yy1, Fb1, delta1, yy2, Fb2, delta2, yy3, Fb3, delta3 As Double

    'Détail position et qualité des armatures
    Public ArmDet(3) As Double
    Public Arm(7, 1) As Double

    'Gestion erreurs
    Public war As Byte = 0
    Public err As Byte = 0


    Public Sub CalculELU()


        If eMu >= 0 Then
            dt = edi
            dc = eds
            yg = yb
            momentest = 1

        Else
            dt = eds
            dc = edi
            yg = eHt - yb
            momentest = 2
        End If

        Mua = abs(eMu)
        If Mua < 10 ^ -4 Then
            Mua = 0
        End If

        d = eHt - dt

        'Frontières
        'Calcul des moments particuliers
        'Limite pivot A - pivot B, MA1
        yy = d * ecu2 / (ecu2 - eud)
        l = yy * ec2 / ecu2

        effort()
        yy1 = yy
        delta1 = delta
        Fb1 = Fb
        MA1 = MA

        'Limite pivot B optimum, Mlima

        yy = d * ecu2 / (ecu2 - ek)
        l = yy * ec2 / ecu2

        effort()

        yy2 = yy
        delta2 = delta
        Fb2 = Fb
        MA2 = MA
        mlim = MA

        'Limite pivotB - pivotC, MA3
        yy = eHt

        effort()
        yy3 = yy
        delta3 = delta
        Fb3 = Fb
        MA3 = MA


        If Mua <= MA1 Then

            pivot = "A"
            est = eud
            yy = d * ecu2 / (ecu2 - eud)
            l = yy * ec2 / ecu2

            effort()
            DIF = (Mua - MA) / MA

            Do While abs(DIF) > 10 ^ (-4)
                effort()
                DIF = (Mua - MA) / MA

                yy = yy * (1 + DIF / 2)
                eb = est * yy / (yy - d)
                l = yy * ec2 / eb

                FAsc = 0
                esc = 0
                Asc = 0
            Loop
        Else


            If Mua <= mlim Then

                'Sans aciers comprimés
                'Départ boucle Pivot A/B

                pivot = "B"
                eb = ecu2
                yy = d * ecu2 / (ecu2 - eud)

                effort()
                DIF = (Mua - MA) / MA

                Do While abs(DIF) > 10 ^ (-4)

                    effort()
                    DIF = (Mua - MA) / MA

                    yy = yy * (1 + DIF / 2)
                    est = eb * (yy - d) / yy
                    l = yy * ec2 / eb

                    FAsc = 0
                    esc = 0
                    Asc = 0

                Loop


            Else
                'Avec aciers comprimés
                pivot = "B - with compression reinforcement"
                eb = ecu2
                yy = d * ecu2 / (ecu2 - eud)

                effort()
                DIF = (mlim - MA) / MA

                Do While abs(DIF) > 10 ^ (-4)

                    effort()
                    DIF = (mlim - MA) / MA

                    yy = yy * (1 + DIF / 2)
                    est = eb * (yy - d) / yy
                    l = yy * ec2 / eb

                    esc = ecu2 * (yy - dc) / yy
                    zs = eHt - dc - dt
                    FAsc = (Mua - mlim) / zs
                Loop

                If FAsc * zs > 0.4 * Mua Then
                    war = 1
                End If
            End If
        End If


        ' effort()

        epsA = est
        SigmaAcier()
        SigmaAst = SigmaA

        epsA = esc
        SigmaAcier()
        SigmaAsc = SigmaA




        If SigmaAsc <> 0 Then
            Asc = abs(FAsc / SigmaAsc)
        Else
            Asc = 0
        End If

        FAst = -Fb - FAsc
        Ast = abs(FAst / SigmaAst)


        'Condition de non fragilité et armature max
        Amin = Asmina(fctm, efyk, eBw, d)
        If Te = True Then
            Asmax = 0.04 * (eBw * eHt + (ebf - eBw) * ehf)
        Else
            Asmax = 0.04 * (eBw * eHt)
        End If


        If Ast < Amin Then
            Ast = Amin
        End If

        If Ast + Asc >= Asmax Then
            err = 1
        End If

        'position des resultats
        If eMu >= 0 Then
            Asup = Asc
            Ainf = Ast
        Else
            Ainf = Asc
            Asup = Ast
        End If

    End Sub

    Public Function Asmina(z As Double, e As Double, r As Double, t As Double) As Double

        Dim Asmin1 As Double
        Dim Asmin2 As Double

        Asmin1 = 0.26 * z / e * r * t
        Asmin2 = 0.0013 * r * t

        If Asmin1 > Asmin2 Then
            Asmina = Asmin1
        Else
            Asmina = Asmin2
        End If

    End Function
    Public Sub effort()

        Dim Mi As Double
        Dim Fbi As Double

        Select Case Te
            Case True

                If momentest = 1 Then

                    'Section en T
                    'Condition Axe neutre inférieur à epaisseur de table - Section comprimée rectangulaire
                    If yy <= ehf Then

                        If yy <= l Then
                            'Diagramme compression parabole incomplète
                            a = yy
                            Parabole()
                            Fb = s1 * ebf
                            Mi = Fb * ag

                        Else
                            'Diagramme compression parabole complète
                            a = l
                            Parabole()
                            Fb = s1 * ebf
                            Mi = Fb * (ag + yy - l)
                            Fbi = fcd * (yy - l) * ebf
                            Fb = Fb + Fbi
                            Mi = Mi + Fbi * (yy - l) / 2
                        End If

                    Else
                        'Condition Axe neutre supérieur à épaisseur de table - Section comprimée en Té
                        If yy <= l Then
                            'Diagramme compression parabole incomplète - Improbable
                            'Warning: Section comprimée en Té dont la parabole n'est pas complète => Compression des ailes est supprimée
                            a = yy
                            Parabole()
                            Fb = s1 * eBw
                            Mi = Fb * ag
                        Else
                            ' Diagramme compression parabole complète
                            If (yy - l) <= ehf Then
                                ' Parabole comprise dans la table de compression
                                a = l
                                Parabole()
                                Fb = s1 * ebf
                                Mi = Fb * (yy - l + ag)
                                a = yy - ehf
                                Parabole()
                                Fbi = s1 * (ebf - eBw)
                                Fb = Fb - Fbi
                                Mi = Mi - Fbi * (ehf + ag)
                                Fbi = fcd * (yy - l) * ebf
                                Fb = Fb + Fbi
                                Mi = Mi + Fbi * (yy - l) / 2
                            Else
                                ' Parabole complète sous la table sous la table de compression
                                a = l
                                Parabole()
                                Fb = s1 * eBw
                                'Effort part parabole
                                Mi = Fb * (ag + yy - l)
                                'Effort part rectangle de la largeur de l'aile
                                Fbi = fcd * (yy - l) * ebf
                                Fb = Fb + Fbi
                                Mi = Mi + Fbi * (yy - l) / 2
                                'Effort part rectangle à soustraire entre aile et parabole
                                Fbi = fcd * (yy - l - ehf) * (ebf - eBw)
                                Fb = Fb - Fbi
                                Mi = Mi - Fbi * (ehf + (yy - l - ehf) / 2)
                            End If
                        End If
                    End If

                Else
                    If yy <= l Then
                        a = yy
                        Parabole()
                        Fb = s1 * eBw
                        Mi = Fb * ag
                    Else
                        a = l
                        Parabole()
                        Fb = s1 * eBw
                        Mi = Fb * (ag + yy - l)
                        Fbi = fcd * (yy - l) * eBw
                        Fb = Fb + Fbi
                        Mi = Mi + Fbi * (yy - l) / 2
                    End If

                End If


            Case False
                If yy <= l Then
                    a = yy
                    Parabole()
                    Fb = s1 * eBw
                    Mi = Fb * ag
                Else
                    a = l
                    Parabole()
                    Fb = s1 * eBw
                    Mi = Fb * (ag + yy - l)
                    Fbi = fcd * (yy - l) * eBw
                    Fb = Fb + Fbi
                    Mi = Mi + Fbi * (yy - l) / 2
                End If

        End Select

        delta = Mi / Fb
        MA = Fb * (d - delta)

        epsB = eb
        SigmaBeton()
        SigmaBt = sigmaB


    End Sub

    Public Sub Parabole()

        Dim u As Double
        u = 1 - a / l
        'Calcul de l'effort de compression de la part de parabole et de la position de son centre de gravité ag
        'Fonction parabole f(x)=1-(1-x/ec2)^n
        s1 = l * fcd * (1 - 1 / (n + 1) - u + u ^ (n + 1) / (n + 1))
        'Valeur approchée de l'effort résultant pour le cas de la parabole n=2
        'Erreur importante si prise en compte pour fck> 50 MPa
        's1 = fcd * (a / l) ^ 2 * (l - a / 3)
        ag = a - l ^ 2 / a * (1 / 2 - 1 / ((n + 1) * (n + 2)) + (1 - a / l) ^ 2 / 2 - (1 - a / l) ^ (n + 2) / (n + 2) - (1 - a / l) + (1 - a / l) ^ (n + 1) / (n + 1)) / (1 - l / (a * (n + 1)) + l * (1 - a / l) ^ (n + 1) / (a * (n + 1)))
        'Valeur approchée du centre de gravité par rapport à la fibre supérieure pour parabole n=2
        'Erreur négligeable si prise en compte pour fck> 50MPa
        'ag = (4 * l - a) / (3 * l - a) * a / 4

    End Sub

    Public Sub SigmaAcier()

        'Contrainte dans la section d'acier en fonction de l'allongement

        If epsA <= 0 Then

            Select Case epsA
                Case Is < ek
                    SigmaA = -fyd * (1 + (k - 1) * (epsA - ek) / (eud - ek))
                Case Is >= ek
                    SigmaA = epsA * Eacier
            End Select

        End If

        'Contrainte dans la section d'acier en fonction du raccourcissement

        If epsA > 0 Then

            Select Case epsA
                Case Is > ek
                    SigmaA = fyd * (1 + (k - 1) * (epsA + ek) / (-eud + ek))
                Case Is <= ek
                    SigmaA = epsA * Eacier
            End Select

        End If

    End Sub

    Public Sub SigmaBeton()
        'Contrainte de la fibre supérieure du béton en fonction de l'allongement

        If epsB < 0 Then
            sigmaB = 0
        ElseIf epsB < ec2 Then
            sigmaB = fcd * (1 - (epsB / ec2 - 1) ^ 2)
        Else
            sigmaB = fcd
        End If

    End Sub

    Private Function abs(p As Double) As Double
        If p < 0 Then
            abs = -p
        Else
            abs = p
        End If
    End Function

    Public Sub ArmDetail(e As Double, h As Double, Ac As Double, co As Double)

        Dim A As Double
        Dim i As Integer
        Dim j As Integer
        Dim HAlim As Integer
        Dim Imax As Integer
        Dim val As Boolean
        Dim rapport As Double


        'Nombre de barres par lits
        'Hypothèse enrobage latérale 10cm
        '1 barre tous les 10cm environ
        ArmDet(0) = Math.Round((e - 0.1) / 0.1, 0) + 1

        'Nomenclature des aciers
        Arm(0, 0) = 8
        Arm(1, 0) = 10
        Arm(2, 0) = 12
        Arm(3, 0) = 14
        Arm(4, 0) = 16
        Arm(5, 0) = 20
        Arm(6, 0) = 25
        Arm(7, 0) = 32
        Arm(0, 1) = 0.00005
        Arm(1, 1) = 0.000079
        Arm(2, 1) = 0.000113
        Arm(3, 1) = 0.000154
        Arm(4, 1) = 0.0002
        Arm(5, 1) = 0.000314
        Arm(6, 1) = 0.00049
        Arm(7, 1) = 0.000803

        'Bonne disposition de ferraillage
        'Limitation du nb de lits et du diamètre des barres en fonction de la largeur
        If e > 0.4 Then
            HAlim = 7
            Imax = 7
        Else
            If e >= 0.3 Then
                HAlim = 7
                Imax = 6
            Else
                If e >= 0.25 Then
                    HAlim = 6
                    Imax = 5
                Else
                    If e >= 0.2 Then
                        HAlim = 5
                        Imax = 4
                    Else
                        HAlim = 4
                        Imax = 3
                    End If
                End If
            End If
        End If


        'Boucle pour obtenir la ditribution des barres correspondant à Ac
        For i = 0 To Imax
            For j = 0 To HAlim

                A = ArmDet(0) * Arm(j, 1) + ArmDet(0) * i * Arm(HAlim, 1)
                If A >= Ac Then
                    ArmDet(1) = i + 1
                    ArmDet(2) = Arm(j, 0)
                    val = True
                    GoTo CalculCdG
                Else
                    val = False
                End If
            Next j
        Next i

        'Erreur si le nombre de lits max doit être dépassé
        If val = False Then
            err = 3
        End If

CalculCdG:
        Select Case ArmDet(1)

            Case Is = 1
                ArmDet(3) = co + 0.008 + (ArmDet(2) * 10 ^ (-3)) / 2
            Case Is = 2
                ArmDet(3) = co + 0.008 + (Arm(HAlim, 0) / 2000 * Arm(HAlim, 1) + (Arm(HAlim, 0) / 1000 + ArmDet(2) / 2000) * 3.14 * ArmDet(2) ^ 2 / 4 * 10 ^ (-6)) / (Arm(HAlim, 1) + 3.14 * ArmDet(2) ^ 2 / 4 * 10 ^ (-6))
            Case Is = 3
                ArmDet(3) = co + 0.008 + (Arm(HAlim, 0) / 2000 * Arm(HAlim, 1) + Arm(HAlim, 0) * 1.5 / 1000 * Arm(HAlim, 1) + (2 * Arm(HAlim, 0) / 1000 + 1.5 * ArmDet(2) / 1000) * 3.14 * ArmDet(2) ^ 2 / 4 * 10 ^ (-6)) / (2 * Arm(HAlim, 1) + 3.14 * ArmDet(2) ^ 2 / 4 * 10 ^ (-6))
            Case Is = 4
                ArmDet(3) = co + 0.008 + (Arm(HAlim, 0) / 2000 * Arm(HAlim, 1) + Arm(HAlim, 0) * 1.5 / 1000 * Arm(HAlim, 1) + Arm(HAlim, 0) * 3.5 / 1000 * Arm(HAlim, 1) + (4 * Arm(HAlim, 0) / 1000 + 0.5 * ArmDet(2) / 1000) * 3.14 * ArmDet(2) ^ 2 / 4 * 10 ^ (-6)) / (3 * Arm(HAlim, 1) + 3.14 * ArmDet(2) ^ 2 / 4 * 10 ^ (-6))
            Case Is = 5
                ArmDet(3) = co + 0.008 + (Arm(HAlim, 0) / 2000 * Arm(HAlim, 1) + Arm(HAlim, 0) * 1.5 / 1000 * Arm(HAlim, 1) + Arm(HAlim, 0) * 3.5 / 1000 * Arm(HAlim, 1) + Arm(HAlim, 0) * 4.5 / 1000 * Arm(HAlim, 1) + (6 * Arm(HAlim, 0) / 1000 + 0.5 * ArmDet(2) / 1000) * 3.14 * ArmDet(2) ^ 2 / 4 * 10 ^ (-6)) / (4 * Arm(HAlim, 1) + 3.14 * ArmDet(2) ^ 2 / 4 * 10 ^ (-6))
            Case Is = 6
                ArmDet(3) = co + 0.008 + (Arm(HAlim, 0) / 2000 * Arm(HAlim, 1) + Arm(HAlim, 0) * 1.5 / 1000 * Arm(HAlim, 1) + Arm(HAlim, 0) * 3.5 / 1000 * Arm(HAlim, 1) + Arm(HAlim, 0) * 4.5 / 1000 * Arm(HAlim, 1) + 6.5 * Arm(HAlim, 0) / 1000 * Arm(HAlim, 1) + (7 * Arm(HAlim, 0) / 1000 + 0.5 * ArmDet(2) / 1000) * 3.14 * ArmDet(2) ^ 2 / 4 * 10 ^ (-6)) / (5 * Arm(HAlim, 1) + 3.14 * ArmDet(2) ^ 2 / 4 * 10 ^ (-6))
            Case Is = 7
                ArmDet(3) = co + 0.008 + (Arm(HAlim, 0) / 2000 * Arm(HAlim, 1) + Arm(HAlim, 0) * 1.5 / 1000 * Arm(HAlim, 1) + Arm(HAlim, 0) * 3.5 / 1000 * Arm(HAlim, 1) + Arm(HAlim, 0) * 4.5 / 1000 * Arm(HAlim, 1) + 6.5 * Arm(HAlim, 0) / 1000 * Arm(HAlim, 1) + 7.5 * Arm(HAlim, 0) / 1000 * Arm(HAlim, 1) + (9 * Arm(HAlim, 0) / 1000 + 0.5 * ArmDet(2) / 1000) * 3.14 * ArmDet(2) ^ 2 / 4 * 10 ^ (-6)) / (6 * Arm(HAlim, 1) + 3.14 * ArmDet(2) ^ 2 / 4 * 10 ^ (-6))
            Case Is = 8
                ArmDet(3) = co + 0.008 + (Arm(HAlim, 0) / 2000 * Arm(HAlim, 1) + Arm(HAlim, 0) * 1.5 / 1000 * Arm(HAlim, 1) + Arm(HAlim, 0) * 3.5 / 1000 * Arm(HAlim, 1) + Arm(HAlim, 0) * 4.5 / 1000 * Arm(HAlim, 1) + 6.5 * Arm(HAlim, 0) / 1000 * Arm(HAlim, 1) + 7.5 * Arm(HAlim, 0) / 1000 * Arm(HAlim, 1) + 9.5 * Arm(HAlim, 0) / 1000 * Arm(HAlim, 1) + (10 * Arm(HAlim, 0) / 1000 + 0.5 * ArmDet(2) / 1000) * 3.14 * ArmDet(2) ^ 2 / 4 * 10 ^ (-6)) / (7 * Arm(HAlim, 1) + 3.14 * ArmDet(2) ^ 2 / 4 * 10 ^ (-6))
        End Select

        'Contrôle rapport hauteur utile / hauteur section

        rapport = (h - ArmDet(3)) / h

        'Erreur si la hauteur utile représent moins de 75% de la hauteur de la section
        If rapport < 0.75 Then
            err = 3
        End If


    End Sub


End Module
