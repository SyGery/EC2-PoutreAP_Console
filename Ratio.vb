Imports System.Math
Module Ratio
    Public Assmax As Double
    Public Aimax As Double
    Public Mmax As Double
    Public Mapp As Double
    Public MtAppGAsup As Double
    Public MtAppGAinf As Double

    'Déclarations des poids d'acier
    'Ai aciers inférieurs  1=ratio avec épure
    'As aciers supérieurs  2=ratio maximum
    'At aciers transversaux
    Public PAi1, PAi2 As Double
    Public PAs1, PAs2 As Double
    Public PAt1, PAt2 As Double
    Public Pp As Double

    Public ratiotype As Boolean

    Public Function poidsAinf(ByVal Amx As Double, ByVal nbresection As Double, ByVal ASinf() As Double, ByVal a() As Double, ByVal L As Double, ByVal phi As Double, ByVal gtyp As Boolean, ByVal dtyp As Boolean) As Integer

        Dim lanc As Double

        PAi1 = 0

        'For i = 0 To nbresection
        'If ASinf(i) < 0.25 * Aimax Then
        'asiinf(i) = 0.25 * Aimax * 1.05
        'Else
        'asiinf(i) = ASinf(i) * 1.05
        'End If
        'Next

        'Intégration du poids
        For i = 0 To nbresection - 1
            PAi1 = PAi1 + ASinf(i) * (a(i + 1) - a(i)) * 7850
        Next

        'Longueur d'ancrage sur appui
        If gtyp = True Then
            If dtyp = True Then
                lanc = 80 * ePHImoyen
            Else
                lanc = 50 * ePHImoyen
            End If
        Else
            If dtyp = True Then
                lanc = 50 * ePHImoyen
            Else
                lanc = 20 * ePHImoyen
            End If
        End If

        PAi1 = PAi1 + 0.5 * Amx * lanc * 7850
        PAi2 = Amx * (L + lanc) * 7850

        Return 0
    End Function

    Public Function poidsAsup(AE As Double, AW As Double, nbresection As Double, AsEpure() As Double, a() As Double, L As Double, Gtyp As Boolean, Dtype As Boolean, Lg As Double, Ld As Double, phi As Double)

        Dim Amx As Double
        Dim volume As Double
        Dim Lmax As Double

        'Aciers sup max
        If AE > AW Then
            Amx = AE
        Else
            Amx = AW
        End If

        'Largeur d'appui max
        If Lg < Ld Then
            Lmax = Ld
        Else
            Lmax = Lg
        End If

        'Poids intérieur travée
        For i = 0 To nbresection - 1
            PAs1 = PAs1 + AsEpure(i) * (a(i + 1) - a(i)) * 7850
        Next

        'Ancrage sur appuis
        If Gtyp = False And Dtype = False Then
            volume = (40 * phi + Lmax) * Amx
        Else
            If Gtyp = True Then
                volume = 40 * phi * AW
            Else
                volume = (40 * phi + Lg) * AW
            End If

            If Dtype = True Then
                volume = volume + 40 * phi * AE
            Else
                volume = volume + (40 * phi + Ld) * AE
            End If

        End If

        'Poids total aciers supérieurs
        PAs1 = PAs1 + volume * 7850
        'Poids total aciers supérieurs maximum
        PAs2 = (Amx * L + volume) * 7850

        Return 0
    End Function

    Public Function poidstranchant(nbresection As Double, aswss() As Double, a() As Double, L As Double, nbponc As Double, R() As Boolean, pelu() As Double, fyd As Double) As Double
        Dim cadre As Double
        Dim aswmax As Double
        Dim aswr(nbponc) As Double
        Dim PAtr As Double

        'Longueur de ferraillage transversal
        If eBw < 0.2 Then
            cadre = eHt + eBw
        Else
            cadre = 0.2 * (eHt + eBw) / eBw + (eHt + 0.1) * (1 - 0.2 / eBw)
        End If

        'Poids d'aciers généré par le relevage des charges ponctuelles

        PAtr = 0
        For i = 0 To nbponc - 1
            If R(i) = True Then
                PAtr = PAtr + cadre * pelu(i) / fyd * 7850
            End If
        Next

        'Poids estimé
        For i = 0 To nbresection - 1
            PAt1 = PAt1 + cadre * aswss(i) * (a(i + 1) - a(i)) * 7850
        Next

        PAt1 = PAt1 + PAtr

        'Poids maximum
        For i = 0 To nbresection
            If aswmax < aswss(i) Then
                aswmax = aswss(i)
            End If
        Next

        PAt2 = aswmax * cadre * L * 7850 + PAtr

        Return 0
    End Function

    Function PoidsPeau(Ht As Double, hd As Double, c As Double, fctm As Double, fyk As Double, L As Double)

        Dim Ap As Double
        Ap = 2 * (Ht - hd) * c * 0.5 * fctm / fyk

        Pp = Ap * L * 7850

        Return 0
    End Function

    Public Function CalRatio(ByVal typ As Boolean, ByVal pi As Double, ByVal ps As Double, ByVal pt As Double, ByVal pp As Double, ByVal l1 As Double, ByVal l2 As Double, ByVal ht As Double, ByVal bw As Double)

        If typ = True Then
            CalRatio = (pi + ps + pt + pp) / (l1 * ht * bw)
        Else
            CalRatio = (pi + ps + pt + pp) / (l2 * ht * bw)
        End If
        'MsgBox(CalRatio)
    End Function

    Public Function EcretageATrans(total As Double, pa As Double, rest As Double, aws() As Double, awsEp() As Double) As Double

        If pa Mod (2) = 0 Then
            For i = 0 To pa / 2 - 1
                For j = 0 To 3
                    If i = 0 Then
                        awsEp(4 * i + j) = aws(1) * 1.05
                    Else
                        awsEp(4 * i + j) = aws(4 * i) * 1.05
                    End If
                Next
            Next
            For i = 0 To pa / 2
                For j = 0 To 3
                    If i = 0 Then
                        awsEp(total - 4 * i - j) = aws(total - 1) * 1.05
                    Else
                        awsEp(total - 4 * i - j) = aws(total - 4 * i) * 1.05
                    End If
                Next
            Next
            For i = 1 To rest
                awsEp(2 * pa + i - 1) = aws(2 * pa + 1) * 1.05
            Next
        Else
            For i = 0 To (pa - 1) / 2 - 1
                For j = 0 To 3
                    If i = 0 Then
                        awsEp(4 * i + j) = aws(1) * 1.05
                    Else
                        awsEp(4 * i + j) = aws(4 * i) * 1.05
                    End If
                Next
            Next
            For i = 0 To (pa - 1) / 2 - 1
                For j = 0 To 3
                    If i = 0 Then
                        awsEp(total - 4 * i - j) = aws(total - 1) * 1.05
                    Else
                        awsEp(total - 4 * i - j) = aws(total - 4 * i) * 1.05
                    End If
                Next
            Next

            For i = 1 To 4 + rest
                awsEp(4 * (pa - 1) / 2 + i - 1) = aws(4 * (pa / 2 - 1) + 1) * 1.05
            Next
        End If

        Return 0

    End Function

End Module
