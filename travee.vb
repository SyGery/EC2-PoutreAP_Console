Imports System.Math
Public Class travee
    Inherits chargeponc

    '---------------
    'GEOMETRIE
    'Portée entre nus
    Public L As Double
    'Portée de calcul
    Public L_calc As Double
    'Portée de ratio
    Public L_ratio As Double
    'Appuis gauche et droite
    'Type d'appui
    'Si AT=true alors appui de rive
    Public AGType As Boolean
    Public ADType As Boolean
    'Largeur d'appui
    Public Ag_L As Double
    Public Ad_L As Double
    Public NuAppG As Double
    Public NuAppD As Double
    'Continuite
    Public cg As Boolean
    Public cd As Boolean
    '---------------
    'CHARGEMENT
    'Charges surfaciques
    'Nombre de bords portant la dalle
    Public d1_cotes As Integer
    'Longueur portée de la dalle gauche
    Public d1_L As Double
    'Valeurs des chargements surfaciques de la dalle gauche (kPa)
    Public d1_g, d1_gp1, d1_gp2, d1_q As Double
    'Idem pour dalle droite d2
    Public d2_cotes As Integer
    Public d2_L As Double
    Public d2_g, d2_gp1, d2_gp2, d2_q As Double
    'Longueur sur laquelle est multipliée la charge surfacique
    Public Chsurf_l As Double
    'Valeurs
    Public chsurf_g_g As Double
    Public chsurf_g_gp1 As Double
    Public chsurf_g_gp2 As Double
    Public chsurf_g_q As Double
    Public chsurf_d_d As Double
    Public chsurf_d_g As Double
    Public chsurf_d_gp1 As Double
    Public chsurf_d_gp2 As Double
    Public chsurf_d_q As Double
    'Charges ponctuelles
    'Nombre de charges ponctuelles par travées
    Public p_nbre As Integer
    'Indice de la charge ponctuelle
    Public p_num As Integer
    'Tableau de charges ponctuelles
    Public p_ch() As chargeponc
    'Charges linéaires
    'Nombre de charges linéaires
    Public l_nbre As Integer
    'Indice de la charge linéaire
    Public l_num As Integer
    'Tableau de charges linéaires
    Public l_ch() As TLin_ch
    '------------------------------------
    'DISCRETISATION
    'Discrétisation de la portée de la travée
    'Nombre de sections de calcul
    Public Nbre_sect As Integer
    'Tableau des abcisses des points
    Public a_sect() As Double
    'Abcisse du changement de charge surfacique de la dalle de gauche
    Public chsurf_g_d As Double
    'longueur du pas entre deux points remarquables
    Public lg As Double
    'Longueur maximale du pas
    Public delta_L As Double

    '----------------------------------
    'RESULTATS
    'Liste des sollicitations par points
    'Moment
    Public T_eff_iso_mt() As Double
    'Effort tranchant à gauche du point
    Public T_eff_iso_vp() As Double
    'Effort tranchant à droite du point
    Public T_eff_iso_vm() As Double
    'Reaction appui gauche
    Public RAppG As Double
    'Reaction appui droit
    Public RAppD As Double

    'MOMENTS ISOSTATIQUES
    'Moment max ELU iso
    Public Miso As Double
    'Abscisse Moment Max ELU iso
    Public aMiso As Double
    'Moment max ELS car
    Public MisoCar As Double
    'Abscisse Moment Max ELS Iso Car
    Public aMisoCar As Double
    'Moment max Iso ELS qp
    Public MisoQp As Double
    'Abscisse Moment Max ELS Iso QP
    Public aMisoQp As Double
    'Abscisse Moment Nul gauche W et droite E
    Public aMNulW As Double
    Public aMNulE As Double

    'MOMENTS HYPERSTATIQUES
    'Moment HYP appui gauche
    Public M1W As Double
    Public M2W As Double
    Public M1WCar As Double
    Public M2WCar As Double
    Public M1WQp As Double
    Public M2WQp As Double
    'Moment HYP appui droite
    Public M1E As Double
    Public M2E As Double
    Public M1ECar As Double
    Public M2ECar As Double
    Public M1EQp As Double
    Public M2EQp As Double
    'Moment travee Hyp ELU
    Public M2tr As Double
    Public M2trCar As Double
    Public M2trQp As Double

    'Moment max ELU hyp
    Public Mhyp As Double
    'Abscisse Moment Max ELU hyp
    Public aMhyp As Double
    'Acier sup MtTr
    Public MtrAsup As Double
    'Acier inf Mt Max
    Public MtrAinf As Double
    'Moment Appui gauche
    Public MtAppG As Double
    'Acier sup MtAppG
    Public MtAppGAsup As Double
    'Acier Ing MtAppG
    Public MtAppGAinf As Double
    'Moment Appui droit
    Public MtAppD As Double
    'Acier sup MtAppD
    Public MtAppDAsup As Double
    'Acier Inf MtAppD
    Public MtAppDAinf As Double

    'Armature
    'Acier max appui gauche
    Public ASupEMax As Double
    'Acier max appui droit
    Public ASupWMax As Double
    'Acier max inf travee
    Public AInfMax As Double

    'Ouverture fissure max
    Public wmax As Double
    'Contrainte armature tendue max
    Public sAmax As Double
    'Contrainte beton max
    Public sCmax As Double
    'Flèche totale admissible
    Public ftotadm As Double
    'Flèche totale
    Public ftotale() As Double
    Public ftot As Double
    'Flèche nuisible admissible
    Public fnuiadm As Double
    'Flèche nuisible
    Public fnuisible() As Double
    Public fnui As Double

    'Ratio armature
    Public ratio1 As Double
    'Ratio armature max
    Public ratio2 As Double

    Public Function chsurf(L As Double, n1 As Double, d1_L As Double, n2 As Double, d2_L As Double, t As Boolean, Ht As Double, hf As Double, bw As Double, b As Double, Opt As Boolean) As Integer
        'Détermination des abcisses liées aux chargements surfaciques

        'Chsurf_g_g, panneau de gauche  
        If n1 = 1 Then
            chsurf_g_d = 0
            Chsurf_l = d1_L
        ElseIf n1 = 2 Then
            chsurf_g_d = 0
            Chsurf_l = d1_L / 2
        ElseIf n1 = 4 Then
            If d1_L < L Then
                chsurf_g_d = d1_L / 2 - eBw / 2
                Chsurf_l = d1_L / 2
            Else
                chsurf_g_d = L / 2
                Chsurf_l = L / 2
            End If
        End If

        'Détermination de la charge poids propre en fonction de l'option pp de poutre et des dimensions de dalle
        If t = True Then
            If Opt = True Then
                chsurf_g_g = d1_g * Chsurf_l + (Ht - hf) * bw / 2 * 25
            Else
                chsurf_g_g = d1_g * Chsurf_l
            End If
        Else
            If Opt = True Then
                chsurf_g_g = d1_g * Chsurf_l + (Ht - d1_g / 25) * bw / 2 * 25
            Else
                chsurf_g_g = d1_g * Chsurf_l
            End If
        End If

        'chsurf_g_g = d1_g * Chsurf_l
        chsurf_g_gp1 = d1_gp1 * Chsurf_l
        chsurf_g_gp2 = d1_gp2 * Chsurf_l
        chsurf_g_q = d1_q * Chsurf_l

        'ChSurf_d_g, panneau de droite
        If n2 = 1 Then
            chsurf_d_d = 0
            Chsurf_l = d2_L
        ElseIf n2 = 2 Then
            chsurf_d_d = 0
            Chsurf_l = d2_L / 2
        ElseIf n2 = 4 Then
            If d2_L < L / 2 Then
                chsurf_d_d = d2_L / 2
            Else
                chsurf_d_d = L / 2
                Chsurf_l = chsurf_d_d
            End If
        End If

        If t = True Then
            If Opt = True Then
                chsurf_d_g = d2_g * Chsurf_l + (Ht - hf) * bw / 2 * 25
            Else
                chsurf_d_g = d2_g * Chsurf_l
            End If
        Else
            If Opt = True Then
                chsurf_d_g = d2_g * Chsurf_l + (Ht - d2_g / 25) * bw / 2 * 25
            Else
                chsurf_d_g = d2_g * Chsurf_l
            End If

        End If

        'chsurf_d_g = d2_g * Chsurf_l
        chsurf_d_gp1 = d2_gp1 * Chsurf_l
        chsurf_d_gp2 = d2_gp2 * Chsurf_l
        chsurf_d_q = d2_q * Chsurf_l
        Return 0

    End Function
    Public Function disc(sec() As Double, L As Double, Ag_L As Double, Tr As Integer, c() As Double, i As Integer, il As Integer, p1_a() As Double, p2_a() As Double) As Integer
        'Discretisation des points de calcul
        Dim v As Integer = 0
        Dim b As Integer
        Dim x As Double
        NuAppG = Min(eHt / 2, Ag_L / 2)
        NuAppD = Min(eHt / 2, Ad_L / 2)
        Nbre_sect = 5
        ReDim Preserve sec(Nbre_sect - 1)
        sec(0) = 0
        sec(1) = NuAppG
        sec(2) = NuAppG + L / 2
        sec(3) = NuAppG + L
        sec(4) = L_calc
        '''' les charges surfaciques
        If chsurf_g_d > 0 Then
            v += 1
            ReDim Preserve sec(4 + v)
            sec(v + 4) = NuAppG + chsurf_g_d
            v += 1
            ReDim Preserve sec(4 + v)
            sec(v + 4) = NuAppG + L - chsurf_g_d
        End If
        If chsurf_d_d > 0 Then
            v += 1
            ReDim Preserve sec(4 + v)
            sec(v + 4) = NuAppG + chsurf_d_d
            v += 1
            ReDim Preserve sec(4 + v)
            sec(v + 4) = NuAppG + L - chsurf_d_d
        End If
        Nbre_sect = 5 + v
        '''' les charges ponctuelles
        If i <> 0 Then
            For b = 0 To i - 1
                ReDim Preserve sec(5 + b + v)
                sec(5 + b + v) = NuAppG + c(b)
            Next
            Nbre_sect = Nbre_sect + i
        End If
        '''' Les charges lineaires
        If il <> 0 Then
            For s = 0 To il - 1
                ReDim Preserve sec(5 + i + v + s)
                sec(5 + b + v + s) = NuAppG + p1_a(s)
                ReDim Preserve sec(5 + i + v + s + 1)
                sec(5 + b + v + s + 1) = NuAppG + p2_a(s)
            Next
            Nbre_sect = Nbre_sect + 2 * il
        End If
        '''' triage ordre croissant
        For j = sec.GetUpperBound(0) - 1 To 0 Step -1
            For v = 0 To j
                If sec(v + 1) < sec(v) Then
                    x = sec(v + 1)
                    sec(v + 1) = sec(v)
                    sec(v) = x
                End If
            Next
        Next
        '''' effacer les duplicats
        Dim f As Integer = 0
        While f < Nbre_sect - 1
            If sec(f + 1) = sec(f) Then
                For v = f + 1 To Nbre_sect - 2
                    sec(v) = sec(v + 1)
                Next
                Nbre_sect = Nbre_sect - 1
            Else
                f += 1
            End If
        End While
        ReDim Preserve sec(Nbre_sect - 1)
        'Longueur minimale d'un pas de poutre
        delta_L = L_calc / 20
        Dim y As Integer
        For j = 0 To Nbre_sect - 2
            lg = (sec(j + 1) - sec(j)) / delta_L
            If lg > Int(lg) Then
                y = Int(lg)
            Else
                y = Int(lg) - 1
            End If
            If y > 0 Then
                ReDim Preserve sec(Nbre_sect - 1 + y)
                Nbre_sect = Nbre_sect + y
                lg = (sec(j + 1) - sec(j)) / (y + 1)
                For nt = 1 To y
                    sec(Nbre_sect - 1 - y + nt) = sec(j) + nt * lg
                Next
            End If
        Next
        For j = sec.GetUpperBound(0) - 1 To sec.GetLowerBound(0) Step -1
            For v = sec.GetLowerBound(0) To j
                If sec(v + 1) < sec(v) Then
                    x = sec(v + 1)
                    sec(v + 1) = sec(v)
                    sec(v) = x
                End If
            Next
        Next
        'Nbre_sect = sec.GetUpperBound(0)
        ReDim Preserve a_sect(Nbre_sect - 1)
        For f = 0 To Nbre_sect - 1
            a_sect(f) = sec(f)
        Next
        Return 0
    End Function
    Public Function Calc_iso_chponc(t As Integer, a() As Double, ch() As Double, str1 As Integer, str2 As Integer, Nbre_travees As Integer, iponc As Integer) As Integer
        'Calcul des sollicitations crées par les charges ponctuelles
        Dim Ra, Rb, Ma, Mb As Double
        Dim s As Integer
        Dim C_eff_vm() As Double
        Dim C_eff_mt() As Double
        Dim C_eff_vp() As Double
        ReDim Preserve T_eff_iso_mt(Nbre_sect - 1)
        ReDim Preserve T_eff_iso_vm(Nbre_sect - 1)
        ReDim Preserve T_eff_iso_vp(Nbre_sect - 1)
        ReDim Preserve C_eff_vm(Nbre_sect - 1)
        ReDim Preserve C_eff_vp(Nbre_sect - 1)
        ReDim Preserve C_eff_mt(Nbre_sect - 1)
        For i = 0 To iponc - 1
            If t = 1 And str1 = 2 Then
                Rb = ch(i)
                Ra = 0
                Ma = 0
                Mb = -ch(i) * (L_calc - a(i))
            ElseIf t = Nbre_travees And str2 = 2 Then
                Rb = 0
                Ra = ch(i)
                Ma = -ch(i) * a(i)
                Mb = 0
            Else
                Rb = ch(i) * (NuAppG + a(i)) / L_calc
                Ra = ch(i) - Rb
                Ma = 0
                Mb = 0
            End If
            '''' effort tranchant 
            s = 0
            While (a_sect(s) < (NuAppG + a(i)) And (s <= Nbre_sect - 1))
                C_eff_vm(s) = -Ra
                C_eff_vp(s) = -Ra
                s += 1
            End While
            C_eff_vm(s) = -Ra
            C_eff_vp(s) = Rb
            s = s + 1
            While s <= Nbre_sect - 1
                C_eff_vm(s) = Rb
                C_eff_vp(s) = Rb
                s += 1
            End While
            C_eff_vm(0) = 0
            C_eff_vp(Nbre_sect - 1) = 0
            '''' moment fléchissant
            s = 0
            While (a_sect(s) <= (NuAppG + a(i)) And (s <= Nbre_sect - 1))
                C_eff_mt(s) = Ma + Ra * a_sect(s)
                s += 1
            End While
            While s <= Nbre_sect - 1
                C_eff_mt(s) = Ma + (Ra * a_sect(s)) - (ch(i) * (a_sect(s) - (NuAppG + a(i))))
                s += 1
            End While
            For s = 0 To Nbre_sect - 1
                T_eff_iso_mt(s) = T_eff_iso_mt(s) + C_eff_mt(s)
                T_eff_iso_vp(s) = T_eff_iso_vp(s) + C_eff_vp(s)
                T_eff_iso_vm(s) = T_eff_iso_vm(s) + C_eff_vm(s)
            Next
        Next
        Array.Clear(C_eff_mt, 0, C_eff_mt.Length)
        Array.Clear(C_eff_vp, 0, C_eff_vp.Length)
        Array.Clear(C_eff_vm, 0, C_eff_vm.Length)
        Return 0
    End Function
    Public Function Calc_iso_chlin(t As Integer, a1() As Double, a2() As Double, ch1() As Double, Nbre_travees As Integer, str1 As Integer, str2 As Integer, ilin As Integer) As Integer
        'Calcul des sollicitations crées par les charges linéaires
        Dim Ra, Rb, Ma, Mb As Double
        Dim s As Integer = 0
        Dim C_eff_vm() As Double
        Dim C_eff_mt() As Double
        Dim C_eff_vp() As Double
        Dim ch_C_eff_vm() As Double
        Dim ch_C_eff_vp() As Double
        Dim ch_t, ch_cdg As Double
        ReDim Preserve T_eff_iso_mt(Nbre_sect - 1)
        ReDim Preserve T_eff_iso_vm(Nbre_sect - 1)
        ReDim Preserve T_eff_iso_vp(Nbre_sect - 1)
        ReDim Preserve C_eff_vm(Nbre_sect - 1)
        ReDim Preserve C_eff_vp(Nbre_sect - 1)
        ReDim Preserve C_eff_mt(Nbre_sect - 1)
        ReDim Preserve ch_C_eff_vm(Nbre_sect - 1)
        ReDim Preserve ch_C_eff_vp(Nbre_sect - 1)
        For i = 0 To ilin - 1
            s = 0
            While a_sect(s) < (NuAppG + a1(i)) And s <= Nbre_sect - 1
                ch_C_eff_vm(s) = 0
                ch_C_eff_vp(s) = 0
                s += 1
            End While
            ch_C_eff_vm(s) = 0
            ch_C_eff_vp(s) = ch1(i)
            s = s + 1
            While (a_sect(s) < (NuAppG + a2(i))) And s <= Nbre_sect - 1
                ch_C_eff_vm(s) = ch1(i)
                ch_C_eff_vp(s) = ch1(i)
                s = s + 1
            End While
            ch_C_eff_vm(s) = ch1(i)
            ch_C_eff_vp(s) = 0
            s = s + 1
            While s <= Nbre_sect - 1
                ch_C_eff_vm(s) = 0
                ch_C_eff_vp(s) = 0
                s = s + 1
            End While
            '''' calcul actions et moments sur appuis
            'action totale
            ch_t = (ch1(i)) * (a2(i) - a1(i))
            'abcsisse d'application de la charge totale
            ch_cdg = NuAppG + (a1(i) + a2(i)) / 2
            ''''Reactions d'appui
            If (t = 1) And str1 = 2 Then
                Rb = ch_t
                Ra = 0
                Ma = 0
                Mb = -ch_t * (L_calc - ch_cdg)
            ElseIf (t = Nbre_travees) And (str2 = 2) Then
                Rb = 0
                Ra = ch_t
                Ma = -ch_t * ch_cdg
                Mb = 0
            Else
                Rb = ch_t * ch_cdg / L_calc
                Ra = ch_t - Rb
                Ma = 0
                Mb = 0
            End If
            '''' efforts tranchants
            C_eff_vm(0) = 0
            C_eff_vp(0) = -Ra
            s = 1
            While s <= Nbre_sect - 1
                C_eff_vm(s) = C_eff_vp(s - 1) + (ch_C_eff_vp(s - 1) + ch_C_eff_vm(s)) / 2 * (a_sect(s) - a_sect(s - 1))
                C_eff_vp(s) = C_eff_vm(s)
                s = s + 1
            End While
            C_eff_vp(Nbre_sect - 1) = 0
            ''''moments flechissants
            C_eff_mt(0) = Ma
            s = 1
            While s <= Nbre_sect - 1
                C_eff_mt(s) = C_eff_mt(s - 1) - (C_eff_vp(s - 1) + C_eff_vm(s)) / 2 * (a_sect(s) - a_sect(s - 1))
                s = s + 1
            End While
            '''' cumul des efforts
            For s = 0 To Nbre_sect - 1
                T_eff_iso_mt(s) = T_eff_iso_mt(s) + C_eff_mt(s)
                T_eff_iso_vp(s) = T_eff_iso_vp(s) + C_eff_vp(s)
                T_eff_iso_vm(s) = T_eff_iso_vm(s) + C_eff_vm(s)
            Next
        Next
        Array.Clear(C_eff_mt, 0, C_eff_mt.Length)
        Array.Clear(C_eff_vp, 0, C_eff_vp.Length)
        Array.Clear(C_eff_vm, 0, C_eff_vm.Length)
        Return 0
    End Function
    Public Function Calc_iso_chtrap(t As Integer, a1 As Double, a2 As Double, ch_v As Double, ch_d As Double, Nbre_travees As Double) As Integer
        'Calcul des sollicitations crées par les charges surfaciques pour un côté de dalle
        Dim Ra, Rb, Ma, Mb As Double
        'p= pente ch_t=charge totale de la travée de la charge surfacique
        'ch_cdg= centre de gravité d'application de la charge totale
        Dim p, ch_t, ch_cdg As Double
        Dim s As Integer
        'Contribution de chaque pas à la sollicitation 
        Dim C_eff_vm() As Double
        Dim C_eff_mt() As Double
        Dim C_eff_vp() As Double
        Dim ch_C_eff_vm() As Double
        Dim ch_C_eff_vp() As Double
        Dim ch() As Double
        ReDim Preserve T_eff_iso_mt(Nbre_sect - 1)
        ReDim Preserve T_eff_iso_vm(Nbre_sect - 1)
        ReDim Preserve T_eff_iso_vp(Nbre_sect - 1)
        ReDim Preserve ch(Nbre_sect - 1)
        ReDim Preserve C_eff_vm(Nbre_sect - 1)
        ReDim Preserve C_eff_vp(Nbre_sect - 1)
        ReDim Preserve C_eff_mt(Nbre_sect - 1)
        ReDim Preserve ch_C_eff_vm(Nbre_sect - 1)
        ReDim Preserve ch_C_eff_vp(Nbre_sect - 1)
        '''' discrétisation de charge trapézoïdale en charge élémentaire
        If ch_d <> 0 Then
            p = ch_v / ch_d
            s = 0
            While (a_sect(s) < a1) And s <= Nbre_sect - 1
                ch(s) = 0
                s += 1
            End While
            While (a_sect(s) <= (a1 + ch_d)) And s <= Nbre_sect - 1
                ch(s) = (a_sect(s) - a1) * p
                s += 1
            End While
            While (a_sect(s) <= (a2 - ch_d)) And s <= Nbre_sect - 1
                ch(s) = ch_v
                s += 1
            End While
            While (a_sect(s) <= a2) And s < Nbre_sect - 1
                ch(s) = ch_v - (a_sect(s) - (a2 - ch_d)) * p
                s += 1
            End While
            While s <= Nbre_sect - 1
                ch(s) = 0
                s += 1
            End While
        Else
            s = 0
            While a_sect(s) < a1 And s <= Nbre_sect - 1
                ch_C_eff_vm(s) = 0
                ch_C_eff_vp(s) = 0
                s += 1
            End While
            ch_C_eff_vm(s) = 0
            ch_C_eff_vp(s) = ch_v
            s = s + 1
            While (a_sect(s) < a2) And s <= Nbre_sect - 1
                ch_C_eff_vm(s) = ch_v
                ch_C_eff_vp(s) = ch_C_eff_vm(s)
                s += 1
            End While
            ch_C_eff_vm(s) = ch_v
            ch_C_eff_vp(s) = 0
            s = s + 1
            While s <= Nbre_sect - 1
                ch_C_eff_vm(s) = 0
                ch_C_eff_vp(s) = 0
                s += 1
            End While
        End If
        'Reactions d'appuis
        '''' calcul actions et moments sur appuis
        ch_t = ch_v * (a2 - a1 - ch_d)
        ch_cdg = (a1 + a2) / 2
        If t = 1 Then
            Rb = ch_t
            Ra = 0
            Ma = 0
            Mb = -ch_t * ch_cdg
        ElseIf t = Nbre_travees Then
            Rb = 0
            Ra = ch_t
            Ma = -ch_t * ch_cdg
            Mb = 0
        Else
            Rb = ch_t * ch_cdg / L_calc
            Ra = ch_t - Rb
            Ma = 0
            Mb = 0
        End If
        'Sollicitations élémentaires par abcisse

        If ch_d <> 0 Then
            ''''Effort tranchant
            C_eff_vm(0) = 0
            C_eff_vp(0) = -Ra
            s = 1
            While s <= Nbre_sect - 1
                C_eff_vm(s) = C_eff_vp(s - 1) + (ch(s - 1) + ch(s)) / 2 * (a_sect(s) - a_sect(s - 1))
                C_eff_vp(s) = C_eff_vm(s)
                s += 1
            End While
            C_eff_vp(Nbre_sect - 1) = 0
            ''''Moment fléchissant
            C_eff_mt(0) = Ma
            s = 1
            While s <= Nbre_sect - 1
                C_eff_mt(s) = C_eff_mt(s - 1) - (C_eff_vp(s - 1) + C_eff_vm(s)) / 2 * (a_sect(s) - a_sect(s - 1))
                s += 1
            End While
        Else
            ''Effort tranchant
            C_eff_vm(0) = 0
            C_eff_vp(0) = -Ra
            s = 1
            While s <= Nbre_sect - 1
                C_eff_vm(s) = C_eff_vp(s - 1) + (ch_C_eff_vp(s - 1) + ch_C_eff_vm(s)) / 2 * (a_sect(s) - a_sect(s - 1))
                C_eff_vp(s) = C_eff_vm(s)
                s += 1
            End While
            C_eff_vp(Nbre_sect - 1) = 0
            ''''Moment fléchissant
            C_eff_mt(0) = Ma
            s = 1
            While s <= Nbre_sect - 1
                C_eff_mt(s) = C_eff_mt(s - 1) - (C_eff_vp(s - 1) + C_eff_vm(s)) / 2 * (a_sect(s) - a_sect(s - 1))
                s += 1
            End While
        End If
        'Somme des solliciations pour obtenir les diagrammes
        For s = 0 To Nbre_sect - 1
            T_eff_iso_mt(s) = T_eff_iso_mt(s) + C_eff_mt(s)
            T_eff_iso_vp(s) = T_eff_iso_vp(s) + C_eff_vp(s)
            T_eff_iso_vm(s) = T_eff_iso_vm(s) + C_eff_vm(s)
        Next
        Array.Clear(C_eff_mt, 0, C_eff_mt.Length)
        Array.Clear(C_eff_vp, 0, C_eff_vp.Length)
        Array.Clear(C_eff_vm, 0, C_eff_vm.Length)
        Return 0
    End Function
End Class
Public Class chargeponc
    Inherits TLin_ch
    'Attributs de la charge ponctuelle
    Public a() As Double
    Public g() As Double
    Public gp1() As Double
    Public gp2() As Double
    Public q() As Double
    Public R() As Boolean
    Public Pelu() As Double
End Class
Public Class TLin_ch
    'Attributs de la charge linéaire
    Public p1_a(), p2_a() As Double
    Public p_g(), p_gp1(), p_gp2(), p_q() As Double
End Class

