Module ProprietesSection
    'Declaration Entree Geometrie	
    Public eBw, ebf, eHt, ehf, eds, edi As Double

    'Declaration Proprietes Section
    Public Sb, Mstat, yb, hutile, Ib As Double
    Public Te As Boolean
    Public h0 As Double


    Public Sub PropSection()

        If Te = False Then
            ebf = 0
            ehf = 0
        Else
            If eBw > ebf Then
                eBw = ebf
            End If
        End If

        Sb = eHt * eBw + ehf * (ebf - eBw)

        Mstat = eBw * (eHt ^ 2) / 2 + ehf * (ebf - eBw) * (eHt - ehf / 2)

        yb = Mstat / Sb

        Ib = eBw * (eHt ^ 3) / 12 + eBw * eHt * ((eHt / 2 - yb) ^ 2) + (ehf ^ 3) * (ebf - eBw) / 12 + ehf * (ebf - eBw) * ((eHt - ehf / 2 - yb) ^ 2)

        If Te = True Then
            h0 = 2 * (eBw * eHt + (ebf - eBw) * ehf) / (2 * ebf + 2 * (eHt - ehf))
        Else
            h0 = 2 * (eBw * eHt) / (2 * eBw + 2 * (eHt - ehf))
        End If

    End Sub
End Module
