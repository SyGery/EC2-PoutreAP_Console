
Imports System.Math
Module ModuleCalcul


    '--------------------------------------------------------------------
    'DONNEES
    'Parametre lu
    Public travee1 As New travee
    'Poids propre poutre
    Public PPOption As Boolean
    'Parametre defini par défaut
    Public Tr As Integer
    Public iponc As Integer
    Public ilin As Integer
    Public chponc As New chargeponc
    'Continuite
    Public ContAppW As Double
    Public ContAppE As Double
    Public ContTr As Double
    'ANF prise en compte des rapports de moment ELS qp sur moment ELS car
    Public ocoeq As Boolean
    'Iteration automatique si erreur fleche
    Public IteratFleche As Boolean
    'Iteration automatique si erreur contrainte
    Public IteratContrainte As Boolean

    '-----------------------------------------------------------------------
    'Paramètre de calcul
    Public Miso_elsqp() As Double
    Public Miso_elsqp_graph() As Double
    Public Miso_elscar() As Double
    Public Miso_elscar_graph() As Double
    Public Miso_elu() As Double
    Public Miso_elu_graph() As Double
    Dim M2W As Double
    Dim M2E As Double
    Public Mhyp_eluenv_min() As Double
    Public Mhyp_eluenv_max() As Double
    Public Tiso_elu_vp() As Double
    Public Tiso_elu_vm() As Double
    Public Tiso_elu() As Double
    Public Thyper_elu_vp() As Double
    Public Thyper_elu_vm() As Double
    Public Thyper_elu() As Double
    Public Tiso_elu_graph() As Double
    Dim MessAppuiG As String
    Dim MessAppuiD As String
    Public a_graph() As Double
    Public w() As Double
    Public sA() As Double
    Public sC() As Double
    Public r() As Double
    Public retr() As Double
    Public rsw() As Double
    Public rsd() As Double
    Public r2sw() As Double
    Public rsd2() As Double
    Public r2sd() As Double
    Public rq() As Double
    Public r2sd2() As Double
    Public vtot() As Double
    Public rinstsw() As Double
    Public rinstsd() As Double
    Public rinstsw2() As Double
    Public vinst1() As Double
    Public vnui() As Double
    Public vretr() As Double
    Public ASinf() As Double
    Public ASinfEpure() As Double
    Public di() As Double
    Public diEpure() As Double
    Public ASsup() As Double
    Public ASsupEpure() As Double
    Public ds() As Double
    Public dsEpure() As Double
    Dim z0 As Double
    Dim zL As Double
    Dim zMtMaxhyp As Double
    Public aswss() As Double
    Public aswssEpure() As Double
    Public v() As Double
    Dim Moment() As Double
    Dim MomentHyper() As Double
    Public flechMq() As Double
    Public flechHypMq() As Double
    Public flechMsd2() As Double
    Public flechHypMsd2() As Double
    Public flechMsd1() As Double
    Public flechHypMsd1() As Double
    Public flechMsw() As Double
    Public flechHypMsw() As Double
    Public iteration As String

    '-------------------------------------------------
    'SORTIES
    Public VerifAncrageGauche As Boolean
    Public VerifBielleGauche As Boolean
    Public VerifAncrageDroit As Boolean
    Public VerifBielleDroit As Boolean
    Public IteraOK As Double


    Public Sub Calcul()

        '------------------------------------------------------------------------------------------
        'INITIALISATION
        Call Initialisation()
        '------------------------------------------------------------------------------
        'CALCUL POUTRE

        'Materiaux
        Call CaracBeton()
        Call CaracAcier()
        Call PropSection()

        'Parametrage de la poutre
        'interpretation du cas de charge surfacique (equivalent charge linéaire; position discontinuité)
        travee1.chsurf(travee1.L, travee1.d1_cotes, travee1.d1_L, travee1.d2_cotes, travee1.d2_L, Te, eHt, ehf, eBw, eb, PPOption)

        'Discretisation de la poutre en n sections
        Dim sec(4) As Double
        travee1.disc(sec, travee1.L, travee1.Ag_L, Tr, travee1.a, travee1.p_nbre, travee1.l_nbre, travee1.p1_a, travee1.p2_a)

        '-----------------------------------------------------------------------
        'SOLLICITATIONS

        'ISOSTATIQUE

        'COMBINAISON ELU

        'Charges ponctuelles
        'Tableau des charges ponctuelles
        Dim ELUchponcg(travee1.Nbre_sect - 1) As Double
        Dim ELUchponcgp1(travee1.Nbre_sect - 1) As Double
        Dim ELUchponcgp2(travee1.Nbre_sect - 1) As Double
        Dim ELUchponcq(travee1.Nbre_sect - 1) As Double
        'Ponderation
        For i = 0 To iponc - 1
            ELUchponcg(i) = gpu * travee1.g(i)
            ELUchponcgp1(i) = gpu * travee1.gp1(i)
            ELUchponcgp2(i) = gpu * travee1.gp2(i)
            ELUchponcq(i) = gqu * travee1.q(i)
            travee1.Pelu(i) = (ELUchponcg(i) + ELUchponcgp1(i) + ELUchponcgp2(i) + ELUchponcq(i)) * 10 ^ (-3)
        Next
        travee1.Calc_iso_chponc(Tr, travee1.a, ELUchponcg, travee1.cg, travee1.cd, 1, iponc)
        travee1.Calc_iso_chponc(Tr, travee1.a, ELUchponcgp1, travee1.cg, travee1.cd, 1, iponc)
        travee1.Calc_iso_chponc(Tr, travee1.a, ELUchponcgp2, travee1.cg, travee1.cd, 1, iponc)
        travee1.Calc_iso_chponc(Tr, travee1.a, ELUchponcq, travee1.cg, travee1.cd, 1, iponc)

        'Charges linéaires
        Dim ELUchlinq(travee1.Nbre_sect - 1) As Double
        Dim ELUchling(travee1.Nbre_sect - 1) As Double
        Dim ELUchlingp1(travee1.Nbre_sect - 1) As Double
        Dim ELUchlingp2(travee1.Nbre_sect - 1) As Double
        'Ponderation
        For i = 0 To ilin - 1
            ELUchling(i) = gpu * travee1.p_g(i)
            ELUchlingp1(i) = gpu * travee1.p_gp1(i)
            ELUchlingp2(i) = gpu * travee1.p_gp2(i)
            ELUchlinq(i) = gqu * travee1.p_q(i)
        Next
        travee1.Calc_iso_chlin(Tr, travee1.p1_a, travee1.p2_a, ELUchling, 1, travee1.cg, travee1.cd, ilin)
        travee1.Calc_iso_chlin(Tr, travee1.p1_a, travee1.p2_a, ELUchlingp1, 1, travee1.cg, travee1.cd, ilin)
        travee1.Calc_iso_chlin(Tr, travee1.p1_a, travee1.p2_a, ELUchlingp2, 1, travee1.cg, travee1.cd, ilin)
        travee1.Calc_iso_chlin(Tr, travee1.p1_a, travee1.p2_a, ELUchlinq, 1, travee1.cg, travee1.cd, ilin)

        'Charges surfaciques
        travee1.Calc_iso_chtrap(Tr, travee1.NuAppG, travee1.NuAppG + travee1.L, gpu * travee1.chsurf_g_g, travee1.chsurf_g_d, 1)
        travee1.Calc_iso_chtrap(Tr, travee1.NuAppG, travee1.NuAppG + travee1.L, gpu * travee1.chsurf_g_gp1, travee1.chsurf_g_d, 1)
        travee1.Calc_iso_chtrap(Tr, travee1.NuAppG, travee1.NuAppG + travee1.L, gpu * travee1.chsurf_g_gp2, travee1.chsurf_g_d, 1)
        travee1.Calc_iso_chtrap(Tr, travee1.NuAppG, travee1.NuAppG + travee1.L, gqu * travee1.chsurf_g_q, travee1.chsurf_g_d, 1)
        travee1.Calc_iso_chtrap(Tr, travee1.NuAppG, travee1.NuAppG + travee1.L, gpu * travee1.chsurf_d_g, travee1.chsurf_d_d, 1)
        travee1.Calc_iso_chtrap(Tr, travee1.NuAppG, travee1.NuAppG + travee1.L, gpu * travee1.chsurf_d_gp1, travee1.chsurf_d_d, 1)
        travee1.Calc_iso_chtrap(Tr, travee1.NuAppG, travee1.NuAppG + travee1.L, gpu * travee1.chsurf_d_gp2, travee1.chsurf_d_d, 1)
        travee1.Calc_iso_chtrap(Tr, travee1.NuAppG, travee1.NuAppG + travee1.L, gqu * travee1.chsurf_d_q, travee1.chsurf_d_d, 1)

        'Construction des listes de sollicitations
        ReDim Miso_elu(travee1.Nbre_sect - 1)
        ReDim Miso_elu_graph(travee1.Nbre_sect - 1)

        ReDim Tiso_elu(travee1.Nbre_sect - 1)
        ReDim Tiso_elu_vp(travee1.Nbre_sect - 1)
        ReDim Tiso_elu_vm(travee1.Nbre_sect - 1)

        For f = 0 To travee1.Nbre_sect - 1
            Miso_elu(f) = travee1.T_eff_iso_mt(f)
            Miso_elu_graph(f) = -travee1.T_eff_iso_mt(f)
            Tiso_elu_vp(f) = travee1.T_eff_iso_vp(f)
            Tiso_elu_vm(f) = travee1.T_eff_iso_vm(f)
        Next

        'Localisation du moment maximum en travee
        'Indice de la section à moment max
        Dim MtMax_sect As Integer
        For f = 0 To travee1.Nbre_sect - 1
            If Miso_elu(f) > travee1.Miso Then
                travee1.Miso = Miso_elu(f)
                MtMax_sect = f
            End If
        Next
        travee1.aMiso = travee1.a_sect(MtMax_sect)

        Array.Clear(travee1.T_eff_iso_mt, 0, travee1.T_eff_iso_mt.Length)
        Array.Clear(travee1.T_eff_iso_vp, 0, travee1.T_eff_iso_vp.Length)
        Array.Clear(travee1.T_eff_iso_vm, 0, travee1.T_eff_iso_vm.Length)

        'COMBINAISON ELS CARACTERISTIQUE
        travee1.Calc_iso_chponc(Tr, travee1.a, travee1.g, travee1.cg, travee1.cd, n, iponc)
        travee1.Calc_iso_chponc(Tr, travee1.a, travee1.gp1, travee1.cg, travee1.cd, n, iponc)
        travee1.Calc_iso_chponc(Tr, travee1.a, travee1.gp2, travee1.cg, travee1.cd, n, iponc)
        travee1.Calc_iso_chponc(Tr, travee1.a, travee1.q, travee1.cg, travee1.cd, n, iponc)
        travee1.Calc_iso_chlin(Tr, travee1.p1_a, travee1.p2_a, travee1.p_g, n, travee1.cg, travee1.cd, ilin)
        travee1.Calc_iso_chlin(Tr, travee1.p1_a, travee1.p2_a, travee1.p_gp1, n, travee1.cg, travee1.cd, ilin)
        travee1.Calc_iso_chlin(Tr, travee1.p1_a, travee1.p2_a, travee1.p_gp2, n, travee1.cg, travee1.cd, ilin)
        travee1.Calc_iso_chlin(Tr, travee1.p1_a, travee1.p2_a, travee1.p_q, n, travee1.cg, travee1.cd, ilin)
        travee1.Calc_iso_chtrap(Tr, travee1.NuAppG, travee1.NuAppG + travee1.L, travee1.chsurf_g_g, travee1.chsurf_g_d, 1)
        travee1.Calc_iso_chtrap(Tr, travee1.NuAppG, travee1.NuAppG + travee1.L, travee1.chsurf_g_gp1, travee1.chsurf_g_d, 1)
        travee1.Calc_iso_chtrap(Tr, travee1.NuAppG, travee1.NuAppG + travee1.L, travee1.chsurf_g_gp2, travee1.chsurf_g_d, 1)
        travee1.Calc_iso_chtrap(Tr, travee1.NuAppG, travee1.NuAppG + travee1.L, travee1.chsurf_g_q, travee1.chsurf_g_d, 1)
        travee1.Calc_iso_chtrap(Tr, travee1.NuAppG, travee1.NuAppG + travee1.L, travee1.chsurf_d_g, travee1.chsurf_d_d, 1)
        travee1.Calc_iso_chtrap(Tr, travee1.NuAppG, travee1.NuAppG + travee1.L, travee1.chsurf_d_gp1, travee1.chsurf_d_d, 1)
        travee1.Calc_iso_chtrap(Tr, travee1.NuAppG, travee1.NuAppG + travee1.L, travee1.chsurf_d_gp2, travee1.chsurf_d_d, 1)
        travee1.Calc_iso_chtrap(Tr, travee1.NuAppG, travee1.NuAppG + travee1.L, travee1.chsurf_d_q, travee1.chsurf_d_d, 1)

        ReDim Miso_elscar(travee1.Nbre_sect - 1)
        ReDim Miso_elscar_graph(travee1.Nbre_sect - 1)
        For f = 0 To travee1.Nbre_sect - 1
            Miso_elscar(f) = travee1.T_eff_iso_mt(f)
            Miso_elscar_graph(f) = -travee1.T_eff_iso_mt(f)
        Next

        Array.Clear(travee1.T_eff_iso_mt, 0, travee1.T_eff_iso_mt.Length)
        Array.Clear(travee1.T_eff_iso_vp, 0, travee1.T_eff_iso_vp.Length)
        Array.Clear(travee1.T_eff_iso_vm, 0, travee1.T_eff_iso_vm.Length)

        'Localisation du moment maximum en travee
        'Indice de la section à moment max
        Dim MtMaxCar_sect As Integer
        For f = 0 To travee1.Nbre_sect - 1
            If Miso_elscar(f) > travee1.MisoCar Then
                travee1.MisoCar = Miso_elscar(f)
                MtMaxCar_sect = f
            End If
        Next
        travee1.aMisoCar = travee1.a_sect(MtMaxCar_sect)

        'COMBINAISON ELS QUASI PERMANENTE
        Dim qq(travee1.Nbre_sect - 1) As Double
        For i = 0 To iponc - 1
            qq(i) = coefqp * travee1.q(i)
        Next
        travee1.Calc_iso_chponc(Tr, travee1.a, travee1.g, travee1.cg, travee1.cd, n, iponc)
        travee1.Calc_iso_chponc(Tr, travee1.a, travee1.gp1, travee1.cg, travee1.cd, n, iponc)
        travee1.Calc_iso_chponc(Tr, travee1.a, travee1.gp2, travee1.cg, travee1.cd, n, iponc)
        travee1.Calc_iso_chponc(Tr, travee1.a, qq, travee1.cg, travee1.cd, n, iponc)
        Dim pq(travee1.Nbre_sect - 1) As Double
        For i = 0 To ilin - 1
            pq(i) = coefqp * travee1.p_q(i)
        Next
        travee1.Calc_iso_chlin(Tr, travee1.p1_a, travee1.p2_a, travee1.p_g, n, travee1.cg, travee1.cd, ilin)
        travee1.Calc_iso_chlin(Tr, travee1.p1_a, travee1.p2_a, travee1.p_gp1, n, travee1.cg, travee1.cd, ilin)
        travee1.Calc_iso_chlin(Tr, travee1.p1_a, travee1.p2_a, travee1.p_gp2, n, travee1.cg, travee1.cd, ilin)
        travee1.Calc_iso_chlin(Tr, travee1.p1_a, travee1.p2_a, pq, n, travee1.cg, travee1.cd, ilin)

        'Chargement trapézoidale
        travee1.Calc_iso_chtrap(Tr, travee1.NuAppG, travee1.NuAppG + travee1.L, travee1.chsurf_g_g, travee1.chsurf_g_d, 1)
        travee1.Calc_iso_chtrap(Tr, travee1.NuAppG, travee1.NuAppG + travee1.L, travee1.chsurf_g_gp1, travee1.chsurf_g_d, 1)
        travee1.Calc_iso_chtrap(Tr, travee1.NuAppG, travee1.NuAppG + travee1.L, travee1.chsurf_g_gp2, travee1.chsurf_g_d, 1)
        travee1.Calc_iso_chtrap(Tr, travee1.NuAppG, travee1.NuAppG + travee1.L, coefqp * travee1.chsurf_g_q, travee1.chsurf_g_d, 1)
        travee1.Calc_iso_chtrap(Tr, travee1.NuAppG, travee1.NuAppG + travee1.L, travee1.chsurf_d_g, travee1.chsurf_d_d, 1)
        travee1.Calc_iso_chtrap(Tr, travee1.NuAppG, travee1.NuAppG + travee1.L, travee1.chsurf_d_gp1, travee1.chsurf_d_d, 1)
        travee1.Calc_iso_chtrap(Tr, travee1.NuAppG, travee1.NuAppG + travee1.L, travee1.chsurf_d_gp2, travee1.chsurf_d_d, 1)
        travee1.Calc_iso_chtrap(Tr, travee1.NuAppG, travee1.NuAppG + travee1.L, coefqp * travee1.chsurf_d_q, travee1.chsurf_d_d, 1)

        ReDim Miso_elsqp(travee1.Nbre_sect - 1)
        ReDim Miso_elsqp_graph(travee1.Nbre_sect - 1)
        For f = 0 To travee1.Nbre_sect - 1
            Miso_elsqp(f) = travee1.T_eff_iso_mt(f)
            Miso_elsqp_graph(f) = -travee1.T_eff_iso_mt(f)
        Next

        'Localisation du moment maximum en travee
        'Indice de la section à moment max
        Dim MtMaxQp_sect As Integer
        For f = 0 To travee1.Nbre_sect - 1
            If Miso_elsqp(f) > travee1.MisoQp Then
                travee1.MisoQp = Miso_elsqp(f)
                MtMaxQp_sect = f
            End If
        Next
        travee1.aMisoQp = travee1.a_sect(MtMaxQp_sect)

        'HYPERSTATIQUE

        'Combinaison ELU
        Dim alphaMtr As Double

        travee1.M1W = -ContAppW * travee1.Miso
        travee1.M1E = -ContAppE * travee1.Miso
        travee1.M2tr = ContTr * travee1.Miso
        alphaMtr = travee1.aMiso / travee1.L_calc

        Call Continuite(travee1.M2tr, travee1.Miso, alphaMtr, travee1.M1W, travee1.M1E)
        travee1.M2W = M2W
        travee1.M2E = M2E

        'Construction de la liste des efforts tranchants ELU

        'Effort tranchant généré par la continuité
        Dim VHyper As Double
        If ContAppW <= 0.15 And ContAppE <= 0.15 Then
            VHyper = 0
        Else
            VHyper = (travee1.M1W - travee1.M1E) / travee1.L_calc
        End If

        ReDim Thyper_elu(travee1.Nbre_sect - 1)
        ReDim Thyper_elu_vp(travee1.Nbre_sect - 1)
        ReDim Thyper_elu_vm(travee1.Nbre_sect - 1)

        For f = 0 To travee1.Nbre_sect - 1
            Thyper_elu(f) = Tiso_elu_vp(f) + VHyper
            Thyper_elu_vp(f) = Tiso_elu_vp(f) + VHyper
            Thyper_elu_vm(f) = Tiso_elu_vm(f) + VHyper
        Next

        'Reprise de la valeur de l'effort tranchant en fin de travee avec la valeur à gauche vm
        Thyper_elu(travee1.Nbre_sect - 1) = Tiso_elu_vm(travee1.Nbre_sect - 1) + VHyper
        TranchantGraph(Thyper_elu_vm, Thyper_elu_vp, travee1.p_nbre, travee1.a)



        'Construction de la liste des moments enveloppe ELU
        Dim Mhyp_elu_m1(travee1.Nbre_sect - 1) As Double
        Dim Mhyp_elu_m2(travee1.Nbre_sect - 1) As Double
        ReDim Mhyp_eluenv_min(travee1.Nbre_sect - 1)
        ReDim Mhyp_eluenv_max(travee1.Nbre_sect - 1)
        Dim Mhyp_eluenv_min_graph(travee1.Nbre_sect - 1) As Double
        Dim Mhyp_eluenv_max_graph(travee1.Nbre_sect - 1) As Double

        For i = 0 To travee1.Nbre_sect - 1
            Mhyp_elu_m1(i) = Miso_elu(i) + travee1.M1W * (1 - travee1.a_sect(i) / travee1.L_calc) + travee1.M1E * (travee1.a_sect(i) / travee1.L_calc)
            Mhyp_elu_m2(i) = Miso_elu(i) + travee1.M2W * (1 - travee1.a_sect(i) / travee1.L_calc) + travee1.M2E * (travee1.a_sect(i) / travee1.L_calc)

            If Mhyp_elu_m1(i) <= 0 Then
                Mhyp_eluenv_min(i) = Mhyp_elu_m1(i)
                Mhyp_eluenv_min_graph(i) = -Mhyp_elu_m1(i)
            Else
                Mhyp_eluenv_min(i) = 0
                Mhyp_eluenv_min_graph(i) = 0
            End If

            If Mhyp_elu_m2(i) >= 0 Then
                Mhyp_eluenv_max(i) = Mhyp_elu_m2(i)
                Mhyp_eluenv_max_graph(i) = -Mhyp_elu_m2(i)
            Else
                Mhyp_eluenv_min(i) = 0
                Mhyp_eluenv_max_graph(i) = 0
            End If
        Next

        'Localisation du moment maximum en travee hyperstatique
        'Indice de la section à moment max
        Dim MtMaxHyp_sect As Double
        For f = 0 To travee1.Nbre_sect - 1
            If Mhyp_eluenv_max(f) > travee1.Mhyp Then
                travee1.Mhyp = Mhyp_eluenv_max(f)
                MtMaxHyp_sect = f
            End If
        Next
        travee1.aMhyp = travee1.a_sect(MtMaxHyp_sect)



        'Combinaison ELS
        travee1.M1WCar = -ContAppW * travee1.MisoCar
        travee1.M1ECar = -ContAppE * travee1.MisoCar
        travee1.M2trCar = ContTr * travee1.MisoCar

        Call Continuite(travee1.M2trCar, travee1.MisoCar, alphaMtr, travee1.M1WCar, travee1.M1ECar)
        travee1.M2WCar = M2W
        travee1.M2ECar = M2E

        travee1.M1WQp = -ContAppW * travee1.MisoQp
        travee1.M1EQp = -ContAppE * travee1.MisoQp
        travee1.M2trQp = ContTr * travee1.MisoQp


        Call Continuite(travee1.M2trQp, travee1.MisoQp, alphaMtr, travee1.M1WQp, travee1.M1EQp)
        travee1.M2WQp = M2W
        travee1.M2EQp = M2E

        'Construction de la liste des moments combinaison ELS caractéristique et quasi permanent
        'Position correspondant à toutes les travées chargées
        'Moment maximum en travée
        Dim Mhyp_elscar(travee1.Nbre_sect - 1) As Double
        Dim Mhyp_elsqp(travee1.Nbre_sect - 1) As Double
        Dim RapMels(travee1.Nbre_sect - 1) As Double

        For i = 0 To travee1.Nbre_sect - 1
            Mhyp_elscar(i) = Miso_elscar(i) + travee1.M2WCar * (1 - travee1.a_sect(i) / travee1.L_calc) + travee1.M2ECar * (travee1.a_sect(i) / travee1.L_calc)
            Mhyp_elsqp(i) = Miso_elsqp(i) + travee1.M2WQp * (1 - travee1.a_sect(i) / travee1.L_calc) + travee1.M2EQp * (travee1.a_sect(i) / travee1.L_calc)
        Next

        RapMels(0) = 1
        For i = 1 To travee1.Nbre_sect - 2
            RapMels(i) = Mhyp_elsqp(i) / Mhyp_elscar(i)
        Next
        RapMels(travee1.Nbre_sect - 1) = 1

        '-------------------------------------------------------------------------
        'Construction Epure de répartition des aciers longitudinaux
        'Index des tronçons de moment nul
        Dim MtNulW_sect As Double
        Dim MtNulE_sect As Double

        For i = 0 To travee1.Nbre_sect - 1
            If Mhyp_elu_m1(i) >= 0 Then
                MtNulE_sect = i
            End If
        Next

        'Abcisse de droite du tronçon de moment nul
        travee1.aMNulE = travee1.a_sect(MtNulE_sect)

        For i = 0 To travee1.Nbre_sect - 1
            If Mhyp_elu_m1(travee1.Nbre_sect - 1 - i) >= 0 Then
                MtNulW_sect = travee1.Nbre_sect - 1 - i
            End If
        Next

        'Abcisse de gauche du tronçon de moment nul
        travee1.aMNulW = travee1.a_sect(MtNulW_sect)

        'Construction liste des armatures inférieures et supérieures
        'index du tronçon correspondant à l'arrêt des barres
        Dim aASupW_sect As Double
        Dim aASupE_sect As Double

        'Gauche W
        'Maximum entre pt de moment nul + décalage 0.8ht et 0.8Ht + Lg ancrage
        Dim distsupW As Double
        If travee1.aMNulW + 0.8 * eHt > 0.8 * eHt + 40 * ePHImoyen Then
            distsupW = travee1.aMNulW + 0.8 * eHt
        Else
            distsupW = 0.8 * eHt + 40 * ePHImoyen
        End If

        'Droite E
        'Maximum entre pt de moment nul + décalage 0.8ht et 0.8Ht + Lg ancrage
        Dim distsupE As Double
        If travee1.L_calc - travee1.aMNulE + 0.8 * eHt > 0.8 * eHt + 40 * ePHImoyen Then
            distsupE = travee1.aMNulE - 0.8 * eHt
        Else
            distsupE = travee1.L_calc - 0.8 * eHt - 40 * ePHImoyen
        End If

        'Recherche du pas correspondant à un décalage de 0.8H par rapport à l'annulation du moment
        For i = 0 To travee1.Nbre_sect - 1
            If distsupW - travee1.a_sect(travee1.Nbre_sect - 1 - i) < 0 Then
                aASupW_sect = travee1.Nbre_sect - 1 - i
            End If

            If distsupE - travee1.a_sect(i) > 0 Then
                aASupE_sect = i
            End If
        Next

        '--------------------------------------------------------------------------
        'DIMENSIONNEMENT BETON ARME

        'Dimensionnement des tableaux en fonction de la discretisation de la poutre
        ReDim r(travee1.Nbre_sect - 1)
        ReDim retr(travee1.Nbre_sect - 1)
        ReDim rsw(travee1.Nbre_sect - 1)
        ReDim rsd(travee1.Nbre_sect - 1)
        ReDim r2sw(travee1.Nbre_sect - 1)
        ReDim rsd2(travee1.Nbre_sect - 1)
        ReDim r2sd(travee1.Nbre_sect - 1)
        ReDim rq(travee1.Nbre_sect - 1)
        ReDim r2sd2(travee1.Nbre_sect - 1)
        ReDim rinstsw(travee1.Nbre_sect - 1)
        ReDim rinstsd(travee1.Nbre_sect - 1)
        ReDim rinstsw2(travee1.Nbre_sect - 1)
        ReDim vinst1(travee1.Nbre_sect - 1)
        ReDim vtot(travee1.Nbre_sect - 1)
        ReDim vnui(travee1.Nbre_sect - 1)
        ReDim vretr(travee1.Nbre_sect - 1)
        ReDim ASinf(travee1.Nbre_sect - 1)
        ReDim di(travee1.Nbre_sect - 1)
        ReDim ASsup(travee1.Nbre_sect - 1)
        ReDim ds(travee1.Nbre_sect - 1)
        ReDim aswss(travee1.Nbre_sect - 1)
        ReDim w(travee1.Nbre_sect - 1)
        ReDim sA(travee1.Nbre_sect - 1)
        ReDim sC(travee1.Nbre_sect - 1)
        ReDim travee1.ftotale(travee1.Nbre_sect - 1)
        ReDim travee1.fnuisible(travee1.Nbre_sect - 1)
        Dim wk2(travee1.Nbre_sect - 1) As Double
        Dim wkmax As Double = 0

        'Construction des listes d'armature théorique ASsup et ASinf
        For f = 0 To travee1.Nbre_sect - 1

            '---------------------------------------------------------------------
            Dim Ainf1 As Double
            Dim di1 As Double
            Dim Ainf2 As Double
            Dim di2 As Double
            Dim Asup1 As Double
            Dim ds1 As Double
            Dim Asup2 As Double
            Dim ds2 As Double

            'Dimensionnement ELU
            eMu = Mhyp_elu_m1(f) * 10 ^ -3
            Call CalculELU()

            'Itération sur la position du centre de gravité des aciers
            'Travail sur les aciers principaux tendus en fonction du signe du moment
            If eMu >= 0 Then
                'Détermination de la position du centre de gravité
                Call ArmDetail(eBw, eHt, Ainf, eenrobage)

                'Itération si nécessaire
                While Abs(ArmDet(3) - edi) / edi > 0.2
                    edi = ArmDet(3)
                    Call CalculELU()
                    Call ArmDetail(eBw, eHt, Ainf, eenrobage)
                End While
                edi = ArmDet(3)
            Else
                Call ArmDetail(eBw, eHt, Asup, eenrobage)

                While Abs(ArmDet(3) - eds) / eds > 0.2
                    eds = ArmDet(3)
                    Call CalculELU()
                    Call ArmDetail(eBw, eHt, Asup, eenrobage)
                End While
                eds = ArmDet(3)
            End If

            Ainf1 = Ainf
            di1 = edi
            Asup1 = Asup
            ds1 = eds

            'Bras de levier pour épure aciers sup
            If f = 0 Then
                z0 = z
            End If

            If f = travee1.Nbre_sect Then
                zL = z
            End If

            eMu = Mhyp_elu_m2(f) * 10 ^ -3
            Call CalculELU()

            'Itération sur la position du centre de gravité des aciers
            'Travail sur les aciers principaux tendus en fonction du signe du moment
            If eMu >= 0 Then
                'Détermination de la position du centre de gravité
                Call ArmDetail(eBw, eHt, Ainf, eenrobage)

                'Itération si nécessaire
                While Abs(ArmDet(3) - edi) / edi > 0.2
                    edi = ArmDet(3)
                    Call CalculELU()
                    Call ArmDetail(eBw, eHt, Ainf, eenrobage)
                End While
                edi = ArmDet(3)
            Else
                Call ArmDetail(eBw, eHt, Asup, eenrobage)

                While Abs(ArmDet(3) - eds) / eds > 0.2
                    eds = ArmDet(3)
                    Call CalculELU()
                    Call ArmDetail(eBw, eHt, Asup, eenrobage)
                End While
                eds = ArmDet(3)
            End If

            Ainf2 = Ainf
            di2 = edi
            Asup2 = Asup
            ds2 = eds

            'Bras de levier pour épure aciers inf
            If f = MtMaxHyp_sect Then
                zMtMaxhyp = z
            End If

            If Ainf1 < Ainf2 Then
                ASinf(f) = Ainf2
                di(f) = di2
            Else
                ASinf(f) = Ainf1
                di(f) = di1
            End If

            If Asup1 < Asup2 Then
                ASsup(f) = Asup2
                ds(f) = ds2
            Else
                ASsup(f) = Asup1
                ds(f) = ds1
            End If

            '------------------------------------------------------------------------
            'Dimensionnement Effort tranchant
            Ved = Thyper_elu(f) * 10 ^ -3
            z = d - ag
            Call MainDimTranchant()
            aswss(f) = Asws

        Next

        '-----------------------------------------------------------------------------
        'Arret de la première boucle de dimensionnement à l'ELU
        'Construction du ferraillage réel
        '-----------------------------------------------------------------------------
        Dim ASupWMax As Double
        Dim ASupEMax As Double
        Dim dsW As Double
        Dim dsE As Double
        ReDim ASsupEpure(travee1.Nbre_sect - 1)
        ReDim dsEpure(travee1.Nbre_sect - 1)

        'ACIERS SUPERIEURS
        'Max Aciers longitudinaux supérieurs gauche
        For i = 0 To aASupW_sect
            If ASsup(i) > ASupWMax Then
                ASupWMax = ASsup(i)
                dsW = ds(i)
            End If
        Next
        travee1.ASupWMax = ASupWMax * 1.05

        'MAX Aciers longitudinaux supérieurs droits
        For i = aASupE_sect To travee1.Nbre_sect - 1
            If ASsup(i) > ASupEMax Then
                ASupEMax = ASsup(i)
                dsE = ds(i)
            End If
        Next
        travee1.ASupEMax = ASupEMax * 1.05

        'Gauche
        For i = 0 To aASupW_sect
            ASsupEpure(i) = ASupWMax * 1.05
            dsEpure(i) = dsW
        Next

        'Centre - max aciers de calcul et aciers de construction
        For i = aASupW_sect + 1 To aASupE_sect - 1
            ASsupEpure(i) = Max(ASsup(i), eBw * 10 * 0.5 * 10 ^ (-4))
            dsEpure(i) = ds(i)
        Next

        'Droit
        For i = aASupE_sect To travee1.Nbre_sect - 1
            ASsupEpure(i) = ASupEMax * 1.05
            dsEpure(i) = dsE
        Next


        'ACIERS INFERIEURS
        'Maximum des aciers longitudinaux inférieurs
        Dim ASinfMax As Double
        Dim dinfMax As Double
        Dim dinfMax2 As Double
        ReDim ASinfEpure(travee1.Nbre_sect - 1)
        ReDim diEpure(travee1.Nbre_sect - 1)

        For i = 0 To travee1.Nbre_sect - 1
            If ASinf(i) > ASinfMax Then
                ASinfMax = ASinf(i)
                dinfMax = di(i)
            End If
        Next

        'Vérification de Ainf > Amin
        If ASinfMax >= Asmin Then
            travee1.AInfMax = ASinfMax * 1.05
        Else
            travee1.AInfMax = Asmin * 1.05
            dinfMax = 0.05
        End If

        'Approche décalage des moments par deux sections
        'Coté gauche
        'index gauche section ASinfmax/2 
        Dim aMtW_sect As Double
        For i = 0 To MtMaxHyp_sect
            If ASinf(MtMaxHyp_sect - i) - ASinfMax / 2 > 0 Then
                aMtW_sect = MtMaxHyp_sect - i
            End If
        Next

        'Décalage de 0.8Ht
        'Abcisse du décalage
        Dim aMtwdecal As Double
        aMtwdecal = travee1.a_sect(aMtW_sect) - 0.8 * eHt

        'Condition décalage ne peut être hors de la travée
        If aMtwdecal < 0 Then
            aMtwdecal = 0
        End If

        'Index gauche decalage = arrêt de gauche du 2ème lit Ainf
        Dim aMtwdecal_sect As Double

        For i = 0 To travee1.Nbre_sect - 1
            If aMtwdecal - travee1.a_sect(travee1.Nbre_sect - 1 - i) < 0 Then
                aMtwdecal_sect = travee1.Nbre_sect - 1 - i - 1
            End If
        Next

        'Coté droit
        'index droit section ASinfmax/2 
        Dim aMtE_sect As Double
        For i = MtMaxHyp_sect To travee1.Nbre_sect - 1
            If ASinf(i) - ASinfMax / 2 > 0 Then
                aMtE_sect = i
            End If
        Next

        'Décalage de 0.8Ht
        'Abcisse du décalage
        Dim aMtedecal As Double
        aMtedecal = travee1.a_sect(aMtE_sect) + 0.8 * eHt

        'Index droit decalage = arrêt de droite du 2ème lit Ainf
        Dim aMtedecal_sect As Double

        For i = 0 To travee1.Nbre_sect - 1
            If aMtedecal - travee1.a_sect(i) > 0 Then
                aMtedecal_sect = i + 1
            End If
        Next

        'Position du CdG de AinfMax/2
        Call ArmDetail(eBw, eHt, travee1.AInfMax / 2, eenrobage)
        dinfMax2 = ArmDet(3)

        'If aMtW_sect < 0 Then
        'aMtW_sect = 0
        'End If

        'If aMtE_sect > travee1.Nbre_sect - 1 Then
        'aMtE_sect = travee1.Nbre_sect - 1
        'End If

        'Construction de la liste d'armature réelle

        If aMtwdecal_sect <> 0 Then
            For i = 0 To aMtwdecal_sect - 1
                ASinfEpure(i) = travee1.AInfMax / 2
                diEpure(i) = dinfMax2
            Next
        End If

        If aMtedecal_sect <> travee1.Nbre_sect Then

            For i = aMtwdecal_sect To aMtedecal_sect
                ASinfEpure(i) = travee1.AInfMax
                diEpure(i) = dinfMax
            Next

            For i = aMtedecal_sect + 1 To travee1.Nbre_sect - 1
                ASinfEpure(i) = travee1.AInfMax / 2
                diEpure(i) = dinfMax2
            Next
        Else
            For i = aMtwdecal_sect To aMtedecal_sect - 1
                ASinfEpure(i) = travee1.AInfMax
                diEpure(i) = dinfMax
            Next

        End If


        'Aciers transversaux
        ReDim aswssEpure(travee1.Nbre_sect - 1)
        Dim pas As Double
        Dim reste As Double

        'Pas où le ferraillage transversale est constant
        pas = Int(travee1.Nbre_sect / 4)
        'Nombre de sections restantes hors rythme des pas
        reste = travee1.Nbre_sect - 4 * pas

        Call EcretageATrans(travee1.Nbre_sect - 1, pas, reste, aswss, aswssEpure)

        '-----------------------------------------------------------------------------------------
        'VERIFICATION DES APPUIS
        'Appui gauche
        'Vérification de la bielle d'about

        VerifBielleGauche = AboutVerif(travee1.AGType, Thyper_elu(1), eBw, travee1.Ag_L, fck, egb)
        'Vérification des aciers de glissement
        VerifAncrageGauche = GlissementVerif(travee1.AGType, Thyper_elu(1), ASinfEpure(1), efyk, egs)

        'Appui droit
        'Vérification de la bielle d'about
        VerifBielleDroit = AboutVerif(travee1.ADType, Thyper_elu(travee1.Nbre_sect - 2), eBw, travee1.Ad_L, fck, egb)
        'Vérification des aciers de glissement
        VerifAncrageDroit = GlissementVerif(travee1.ADType, Thyper_elu(travee1.Nbre_sect - 2), ASinfEpure(travee1.Nbre_sect - 2), efyk, egs)




        '-------------------------------------------------------------------------------
        'Vérification des critères de service (contraintes sous combinaison rare, ouverture fissures sous quasi permanente
        'flèche sous quasi permanent)
        'Reprise de la boucle avec le ferraillage réel
        '--------------------------------------------------------------------------
        For f = 0 To travee1.Nbre_sect - 1
            Asup = ASsupEpure(f)
            eds = dsEpure(f)
            Ainf = ASinfEpure(f)
            edi = diEpure(f)
            '----------------------------------------------------------------------
            'Verification ELS Contrainte
            eMscar = Mhyp_elscar(f) * 10 ^ -3
            Ms = eMscar
            Msa = Abs(eMscar)
            'Coef de fluage
            Phi = Fluage(RH, et1, et, h0)
            If ocoeq = True Then
                Phi = Phi * RapMels(f)
            End If
            Call ContrainteELS(Phi)
            sA(f) = sAst
            sC(f) = scb
            Call VerifContrainte()
            '----------------------------------------------------------------------
            'Verification ELS fissure
            eMsqp = Mhyp_elsqp(f) * 10 ^ -3
            Ms = eMsqp
            Msa = Abs(eMsqp)
            Phi = Fluage(RH, et1, et, h0)
            If ocoeq = True Then
                Phi = Phi * RapMels(f)
            End If
            Call ContrainteELS(Phi)
            Call Ouverture()
            w(f) = wk
            Call courbure(eMsqp)
            Call coubureretrait()
            r(f) = r1
            retr(f) = rcs

            '-----------------------------------------------------------------------------------------------
            'Itération sur le ferraillage longitudinal inférieur si la contrainte de traction de l'acier est trop importante
            If IteratContrainte = True Then
                While errS = 1 Or errS = 3 Or errS = 5 Or errS = 7
                    If ASinf(f) > Asmax Then
                        err = 1
                        'MsgBox("Iteration armature-> Ferraillage longitudinal > 4% section", MsgBoxStyle.Critical, "Error")
                        Exit For
                    End If
                    ASinfEpure(f) = ASinfEpure(f) + 1 / 10000

                    Call ArmDetail(eBw, eHt, ASinfEpure(f), eenrobage)
                    edi = ArmDet(3)
                    diEpure(f) = edi
                    If err = 3 Then
                        Exit For
                    End If

                    Msa = Abs(eMscar)
                    Phi = Fluage(RH, et1, et, h0)
                    If ocoeq = True Then
                        Phi = Phi * RapMels(f)
                    End If
                    Call ContrainteELS(Phi)
                    sA(f) = sAst
                    sC(f) = scb
                    Call VerifContrainte()
                    'Ouverture des fissures
                    Msa = Abs(eMsqp)
                    Phi = Fluage(RH, et1, et, h0)
                    If ocoeq = True Then
                        Phi = Phi * RapMels(f)
                    End If
                    Call ContrainteELS(Phi)
                    Call Ouverture()
                    w(f) = wk
                End While

            End If

            If err <> 0 Then
                GoTo Erreur
            End If

            '---------------------------------------------------------------------------
            'Calcul des courbures pour chaque combinaisons de chargement, temps

            'courbure de la flèche totale
            'Listes des moments par type de charge 
            fleche()

            'Msw Moment Poids propre G
            If ocoeq = True Then
                Phi = Fluage(RH, et1, et, h0) * RapMels(f)
            Else
                Phi = Fluage(RH, et1, et, h0)
            End If
            Call ContrainteFleche(flechHypMsw(f), Phi)
            Call courbure(flechHypMsw(f))
            rsw(f) = r1

            'Msd1 Moment Poids propre G + G'1
            If ocoeq = True Then
                Phi = Fluage(RH, et2, et, h0) * RapMels(f)
            Else
                Phi = Fluage(RH, et2, et, h0)
            End If
            Call ContrainteFleche(flechHypMsd1(f), Phi)
            Call courbure(flechHypMsd1(f))
            rsd(f) = r1

            Call ContrainteFleche(flechHypMsw(f), Phi)
            Call courbure(flechHypMsw(f))
            r2sw(f) = r1

            'Msd2 Moment Poids propre G + G'1 + G'2
            If ocoeq = True Then
                Phi = Fluage(RH, et3, et, h0) * RapMels(f)
            Else
                Phi = Fluage(RH, et3, et, h0)
            End If

            Call ContrainteFleche(flechHypMsd2(f), Phi)
            Call courbure(flechHypMsd2(f))
            rsd2(f) = r1
            Call ContrainteFleche(flechHypMsd1(f), Phi)
            Call courbure(flechHypMsd1(f))
            r2sd(f) = r1

            'Mq Moment Poids propre G+ G'1+ G'2+ Psi Q
            If ocoeq = True Then
                Phi = Fluage(RH, et4, et, h0) * RapMels(f)
            Else
                Phi = Fluage(RH, et4, et, h0)
            End If

            Call ContrainteFleche(flechHypMq(f), Phi)
            Call courbure(flechHypMq(f))
            rq(f) = r1
            Call ContrainteFleche(flechHypMsd2(f), Phi)
            Call courbure(flechHypMsd2(f))
            r2sd2(f) = r1

            'courbure de la flèche d'installation
            If ocoeq = True Then
                Phi = Fluage(RH, et1, et2, h0) * RapMels(f)
            Else
                Phi = Fluage(RH, et1, et2, h0)
            End If

            Call ContrainteFleche(flechHypMsw(f), Phi)
            Call courbure(flechHypMsw(f))
            rinstsw(f) = r1
            Call ContrainteFleche(flechHypMsd1(f), 0)
            Call courbure(flechHypMsd1(f))
            rinstsd(f) = r1
            Call ContrainteFleche(flechHypMsw(f), 0)
            Call courbure(flechHypMsw(f))
            rinstsw2(f) = r1
        Next

        Array.Clear(travee1.T_eff_iso_mt, 0, travee1.T_eff_iso_mt.Length)
        Array.Clear(travee1.T_eff_iso_vp, 0, travee1.T_eff_iso_vp.Length)
        Array.Clear(travee1.T_eff_iso_vm, 0, travee1.T_eff_iso_vm.Length)

        '---------------------------------------------------------------------------
        'Calcul de flèche

        'Calcul des composantes de flèche

        ''Calcul de la fleche d'installation qui sera soustraite à la flèche totale -> flèche nuisible
        'Flèche nuisible = flèche apparaissant après l'installation des premières charges permanentes g'1
        calculfleche(rinstsw)
        For s = 0 To travee1.Nbre_sect - 1
            vinst1(s) = v(s)
        Next
        calculfleche(rinstsd)
        For s = 0 To travee1.Nbre_sect - 1
            vinst1(s) = vinst1(s) + v(s)
        Next
        calculfleche(rinstsw2)
        For s = 0 To travee1.Nbre_sect - 1
            vinst1(s) = vinst1(s) - v(s)
        Next

        '' calcul des composantes de la fleche totale
        'Calcul de la flèche de retrait
        Call calculfleche(retr)
        For s = 0 To travee1.Nbre_sect - 1
            vretr(s) = v(s)
        Next

        'Flèche SW long terme
        calculfleche(rsw)
        For s = 0 To travee1.Nbre_sect - 1
            vtot(s) = v(s)
        Next

        'Flèche SW+G'1 long terme
        calculfleche(rsd)
        For s = 0 To travee1.Nbre_sect - 1
            vtot(s) = v(s) + vtot(s)
        Next
        'Flèche SW à partir installation G'1 à soustraire
        calculfleche(r2sw)
        For s = 0 To travee1.Nbre_sect - 1
            vtot(s) = vtot(s) - v(s)
        Next

        'Flèche SW+G'1+G'2 long terme
        calculfleche(rsd2)
        For s = 0 To travee1.Nbre_sect - 1
            vtot(s) = vtot(s) + v(s)
        Next
        'Flèche SW+G'1 à partir installation G'2 à soustraire
        calculfleche(r2sd)
        For s = 0 To travee1.Nbre_sect - 1
            vtot(s) = vtot(s) - v(s)
        Next

        'Flèche SW+G'1+G'2+alpha Q long terme
        calculfleche(rq)
        For s = 0 To travee1.Nbre_sect - 1
            vtot(s) = vtot(s) + v(s)
        Next

        'Dernière opération pour obtenir la flèche totale
        'Flèche SW+G'1+G'2 à partir de Q à soustraire
        calculfleche(r2sd2)
        If PriseRetrait = True Then
            For s = 0 To travee1.Nbre_sect - 1
                travee1.ftotale(s) = vtot(s) - v(s) + vretr(s)
            Next
        Else
            For s = 0 To travee1.Nbre_sect - 1
                travee1.ftotale(s) = vtot(s) - v(s)
            Next
        End If

        'Calcul de la flèche nuisible
        For s = 0 To travee1.Nbre_sect - 1
            travee1.fnuisible(s) = travee1.ftotale(s) - vinst1(s)
        Next

        'Définition des critères de flèche totale et nuisible
        travee1.ftotadm = -travee1.L / 250
        travee1.fnuiadm = -travee1.L / 500

        'Verification des critères de flèche
        Dim FlecheVerif As Boolean
        FlecheVerif = True

        For f = 0 To travee1.Nbre_sect - 1
            If travee1.ftotadm > travee1.ftotale(f) Or travee1.fnuiadm > travee1.fnuisible(f) Then
                FlecheVerif = False
            End If
        Next


        '---------------------------------------------------------------------------------------------------------------------
        'Itération sur l'armature longitudinale inférieure pour le respect des critères de flèche

        If IteratFleche = True And FlecheVerif = False Then

            IteraOK = 1

            For s = 0 To travee1.Nbre_sect - 1
                While travee1.ftotadm > travee1.ftotale(s) Or travee1.fnuiadm > travee1.fnuisible(s)
                    For i = 0 To travee1.Nbre_sect - 1
                        ASinfEpure(i) = ASinfEpure(i) + 1 / 10000
                        If ASinfEpure(i) + ASsupEpure(i) >= Asmax Then
                            err = 1
                        Else
                            err = 0
                        End If
                    Next

                    Select Case err
                        Case 1
                            'MsgBox("Iteration fleche-> Ferraillage longitudinal > 4% section", MsgBoxStyle.Critical, "Erreur")
                            Exit For
                    End Select

                    'Détermination des courbures après itération i sur les aciers
                    For f = 0 To travee1.Nbre_sect - 1
                        Ainf = ASinfEpure(f)
                        Call ArmDetail(eBw, eHt, Ainf, eenrobage)
                        edi = ArmDet(3)
                        diEpure(f) = edi
                        If err = 3 Then
                            Exit While
                        End If

                        'COURBURE DE LA FLECHE TOTALE
                        'rsw, SW t1-infinie
                        Phi = Fluage(RH, et1, et, h0)
                        Call ContrainteFleche(flechMsw(f), Phi)
                        Call courbure(flechMsw(f))
                        rsw(f) = r1
                        'rsd, SW+G'1 t2-infinie
                        Phi = Fluage(RH, et2, et, h0)
                        Call ContrainteFleche(flechMsd1(f), Phi)
                        Call courbure(flechMsd1(f))
                        rsd(f) = r1
                        'r2sw, SW t2-infinie
                        Call ContrainteFleche(flechMsw(f), Phi)
                        Call courbure(flechMsw(f))
                        r2sw(f) = r1
                        'rsd2, Sw+G'1+G'2 t3-infinie
                        Phi = Fluage(RH, et3, et, h0)
                        Call ContrainteFleche(flechMsd2(f), Phi)
                        Call courbure(flechMsd2(f))
                        rsd2(f) = r1
                        'r2sd, SW+G'1 t3-infinie
                        Call ContrainteFleche(flechMsd1(f), Phi)
                        Call courbure(flechMsd1(f))
                        r2sd(f) = r1
                        'rq, SW+G'1+G'2+Q t4-infinie
                        Phi = Fluage(RH, et4, et, h0)
                        Call ContrainteFleche(flechMq(f), Phi)
                        Call courbure(flechMq(f))
                        rq(f) = r1
                        'r2sd2, SW+G'1+G'2 t4-infinie
                        Call ContrainteFleche(flechMsd2(f), Phi)
                        Call courbure(flechMsd2(f))
                        r2sd2(f) = r1
                        'COURBURE DE LA FLECHE INSTALLATION
                        Phi = Fluage(RH, et1, et2, h0)
                        Call ContrainteFleche(flechMsw(f), Phi)
                        Call courbure(flechMsw(f))
                        rinstsw(f) = r1
                        Call ContrainteFleche(flechMsd1(f), 0)
                        Call courbure(flechMsd1(f))
                        rinstsd(f) = r1
                        Call ContrainteFleche(flechMsw(f), 0)
                        Call courbure(flechMsw(f))
                        rinstsw2(f) = r1
                    Next
                    calculfleche(rsw)
                    For ss = 0 To travee1.Nbre_sect - 1
                        vtot(ss) = v(ss)
                    Next
                    calculfleche(rsd)
                    For ss = 0 To travee1.Nbre_sect - 1
                        vtot(ss) = v(ss) + vtot(ss)
                    Next
                    calculfleche(r2sw)
                    For ss = 0 To travee1.Nbre_sect - 1
                        vtot(ss) = vtot(ss) - v(ss)
                    Next
                    calculfleche(rsd2)
                    For ss = 0 To travee1.Nbre_sect - 1
                        vtot(ss) = vtot(ss) + v(ss)
                    Next
                    calculfleche(r2sd)
                    For ss = 0 To travee1.Nbre_sect - 1
                        vtot(ss) = vtot(ss) - v(ss)
                    Next
                    calculfleche(rq)
                    For ss = 0 To travee1.Nbre_sect - 1
                        vtot(ss) = vtot(ss) + v(ss)
                    Next
                    calculfleche(r2sd2)
                    For ss = 0 To travee1.Nbre_sect - 1
                        travee1.ftotale(ss) = vtot(ss) - v(ss)

                    Next

                    ''''fleche d'installation
                    calculfleche(rinstsw)
                    For ss = 0 To travee1.Nbre_sect - 1
                        vinst1(ss) = v(ss)
                    Next
                    calculfleche(rinstsd)
                    For ss = 0 To travee1.Nbre_sect - 1
                        vinst1(ss) = vinst1(ss) + v(ss)
                    Next
                    calculfleche(rinstsw2)
                    For ss = 0 To travee1.Nbre_sect - 1
                        vinst1(ss) = vinst1(ss) - v(ss)
                        travee1.fnuisible(ss) = travee1.ftotale(ss) - vinst1(ss)

                    Next
                End While

                If err <> 0 Then
                    GoTo Erreur
                End If


                '----------------------------------------------------------------------
                'Verification ELS Contrainte
                eMscar = Mhyp_elscar(s) * 10 ^ -3
                Ms = eMscar
                Msa = Abs(eMscar)
                Ainf = ASinfEpure(s)
                Phi = Fluage(RH, et1, et, h0)
                If ocoeq = True Then
                    Phi = Phi * RapMels(s)
                End If
                Call ContrainteELS(Phi)
                sA(s) = sAst
                sC(s) = scb
                'Call VerifContrainte()
                '----------------------------------------------------------------------
                'Verification ELS fissure
                eMsqp = Mhyp_elsqp(s) * 10 ^ -3
                Ms = eMsqp
                Msa = Abs(eMsqp)
                Phi = Fluage(RH, et1, et, h0)
                If ocoeq = True Then
                    Phi = Phi * RapMels(s)
                End If
                Call ContrainteELS(Phi)
                Call Ouverture()
                w(s) = wk
                'Form3.DataGridView1.Rows.Item(s).Cells(22).Value = Round(wk2(s), 2)
            Next
        End If

        '------------------------------------------------------------------------------------------------------
        'For s = 0 To travee1.Nbre_sect - 1
        'vnui(s) = vtot(s) - vinst1(s)
        'Next

        'Sorties de la section de ferraillage au point de moment max et min
        travee1.MtrAinf = ASinfEpure(MtMaxHyp_sect) * 10 ^ 4
        travee1.MtrAsup = ASsupEpure(MtMaxHyp_sect) * 10 ^ 4

        travee1.MtAppG = Mhyp_elu_m1(0)
        travee1.MtAppGAinf = ASinfEpure(0) * 10 ^ 4
        travee1.MtAppGAsup = ASsupEpure(0) * 10 ^ 4

        travee1.MtAppD = Mhyp_elu_m1(travee1.Nbre_sect - 1)
        travee1.MtAppDAinf = ASinfEpure(travee1.Nbre_sect - 1) * 10 ^ 4
        travee1.MtAppDAsup = ASsupEpure(travee1.Nbre_sect - 1) * 10 ^ 4


Erreur:
        'Arret si dépassement du ferraillage maximal/table de compression non valide/ferraillage incompatible section
        Select Case err
            Case 1
                Exit Sub
            Case 2
                Exit Sub
            Case 3
                Exit Sub
        End Select

        'Recherche de la contrainte de traction de l'armature tendue min
        travee1.sAmax = 0
        For f = 0 To travee1.Nbre_sect - 1
            If sA(f) < travee1.sAmax Then
                travee1.sAmax = sA(f)
            End If
        Next

        'Recherche de la contrainte de compression max du béton
        travee1.sCmax = 0
        For f = 0 To travee1.Nbre_sect - 1
            If sC(f) > travee1.sCmax Then
                travee1.sCmax = sC(f)
            End If
        Next

        'Recherche de l'ouverture de fissure maximale
        travee1.wmax = 0
        For f = 0 To travee1.Nbre_sect - 1
            If w(f) > travee1.wmax Then
                travee1.wmax = w(f)
            End If
        Next

        '''' fleche totale
        'calculfleche(r)
        'For f = 0 To travee1.Nbre_sect - 1
        ' Form3.DataGridView1.Rows.Item(f).Cells(12).Value = Round(v(f) * 1000, 2)
        ' Next

        '''' fleche retrait
        'Double intégration des courbures retr()
        Call calculfleche(retr)
        'For f = 0 To travee1.Nbre_sect - 1

        'Next

        '------------------------------------------------------------------------------
        'Formalisation des résultats de flèche totale
        Dim ftotverif As Boolean
        travee1.ftot = 0
        ftotverif = True

        For f = 0 To travee1.Nbre_sect - 1
            If travee1.ftotale(f) < travee1.ftot Then
                travee1.ftot = travee1.ftotale(f)
            End If
        Next
        If travee1.ftot < travee1.ftotadm Then
            ftotverif = False

        End If


        'Formalisation des résultats de flèche nuisible
        Dim fnuiverif As Boolean
        travee1.fnui = 0

        fnuiverif = True
        For f = 0 To travee1.Nbre_sect - 1
            If travee1.fnuisible(f) < travee1.fnui Then
                travee1.fnui = travee1.fnuisible(f)
            End If
        Next

        If travee1.fnui < travee1.fnuiadm Then
            fnuiverif = False

        End If

        '---------------------------------------------------------------------------
        'Ratio d'armature
        Call poidsAinf(travee1.MtrAinf * 10 ^ (-4), travee1.Nbre_sect - 1, ASinfEpure, travee1.a_sect, travee1.L, ePHImoyen, travee1.AGType, travee1.ADType)
        Call poidsAsup(travee1.ASupEMax, travee1.ASupWMax, travee1.Nbre_sect - 1, ASsupEpure, travee1.a_sect, travee1.L, travee1.AGType, travee1.ADType, travee1.Ag_L, travee1.Ad_L, ePHImoyen)
        Call poidstranchant(travee1.Nbre_sect - 1, aswssEpure, travee1.a_sect, travee1.L, travee1.p_nbre, travee1.R, travee1.Pelu, efyk / egs)
        Call PoidsPeau(eHt, ehf, 0.04, fctm, efyk, travee1.L_calc)

        'MsgBox(Round(PAi1) & "  " & Round(PAi2) & "  " & Round(PAs1) & "   " & Round(PAs2) & "  " & Round(PAt1) & "   " & Round(PAt2) & "   " & Round(Pp))

        travee1.ratio1 = CalRatio(ratiotype, PAi1, PAs1, PAt1, Pp, travee1.L, travee1.L_ratio, eHt, eBw)
        travee1.ratio2 = CalRatio(ratiotype, PAi2, PAs2, PAt2, Pp, travee1.L, travee1.L_ratio, eHt, eBw)

    End Sub

    Sub Initialisation()

        'Portée de calul
        travee1.L_calc = travee1.L + Min(eHt / 2, travee1.Ag_L / 2) + Min(eHt / 2, travee1.Ad_L / 2)
        'Ratio
        travee1.L_ratio = travee1.L + travee1.Ag_L / 2 + travee1.Ad_L / 2

        travee1.l_nbre = 0
        ReDim travee1.p1_a(0)
        ReDim travee1.p2_a(0)

        Tr = 0

        'Verification geometrique de la section
        If Te = False Then
            ehf = eHt
            ebf = eBw
        End If

        If ebf < eBw Then
            ebf = eBw
            Te = False
            Exit Sub
        End If

        If ehf > eHt Then
            ehf = eHt
            Exit Sub
        End If

        'Vérification largeur de table de compression §5.3.2.1
        Dim b1 As Double
        Dim b2 As Double
        Dim beff1 As Double
        Dim beff2 As Double
        Dim l0 As Double

        If ContAppW = 0.15 Then
            If ContAppE = 0.15 Then
                l0 = 1
            Else
                l0 = 0.85
            End If
        Else
            If ContAppE = 0.15 Then
                l0 = 0.85
            Else
                l0 = 0.7
            End If
        End If

        'Gauche
        Select Case travee1.d1_cotes
            Case 1
                b1 = travee1.d1_L
            Case 2
                b1 = travee1.d1_L / 2
            Case 4
                b1 = travee1.d1_L / 2
        End Select

        Dim a As Double
        If b1 < 0.2 * l0 * travee1.L_calc Then
            a = b1
        Else
            a = 0.2 * l0 * travee1.L_calc
        End If

        beff1 = 0.2 * b1 + 0.1 * l0 * travee1.L_calc
        If beff1 > a Then
            beff1 = a
        End If

        'Droite
        Select Case travee1.d2_cotes
            Case 1
                b2 = travee1.d2_L
            Case 2
                b2 = travee1.d2_L / 2
            Case 4
                b2 = travee1.d2_L / 2
        End Select

        If b2 < 0.2 * l0 * travee1.L_calc Then
            a = b2
        Else
            a = 0.2 * l0 * travee1.L_calc
        End If

        beff2 = 0.2 * b2 + 0.1 * l0 * travee1.L_calc
        If beff2 > a Then
            beff2 = a
        End If

        Dim btable As Double
        btable = beff1 + beff2 + eBw

        If btable < ebf Then
            err = 2
        End If

        'Initialisation sorties
        With travee1
            .RAppG = 0 : .RAppD = 0 : .Miso = 0 : .aMiso = 0 : .MisoCar = 0 : .aMisoCar = 0 : .MisoQp = 0 : .aMisoQp = 0
            .aMNulW = 0 : .aMNulE = 0 : .M1W = 0 : .M2W = 0 : .M1WCar = 0 : .M2WCar = 0 : .M1WQp = 0 : .M2WQp = 0
            .M1E = 0 : .M2E = 0 : .M1ECar = 0 : .M2ECar = 0 : .M1EQp = 0 : .M2EQp = 0 : .M2tr = 0 : .M2trCar = 0 : .M2trQp = 0
            .Mhyp = 0 : .aMhyp = 0 : .MtrAsup = 0 : .MtrAinf = 0 : .MtAppG = 0 : .MtAppGAsup = 0 : .MtAppGAinf = 0 : .MtAppD = 0 : .MtAppDAsup = 0 : .MtAppDAinf = 0
            .wmax = 0 : .sAmax = 0 : .sCmax = 0 : .ftotadm = 0 : .ftot = 0 : .fnuiadm = 0 : .fnui = 0 : .Nbre_sect = 0 : .ratio1 = 0 : .ratio2 = 0 : IteraOK = 0
        End With

        'erreurs
        err = 0 : errS = 0 : war = 0

        'Poids Aciers
        PAi1 = 0 : PAs1 = 0 : PAs2 = 0 : PAt1 = 0 : PAi2 = 0 : PAt2 = 0

    End Sub

    Public Function calculfleche(r) As Double

        'Double intégration des courbures
        'Rotation, premiere integration
        Dim w(travee1.Nbre_sect - 1) As Double
        'Fleche deuxieme integration
        ReDim v(travee1.Nbre_sect - 1)

        Dim ss As Double = 0
        Dim s2 As Double = 0
        Array.Clear(w, 0, w.Length)
        Array.Clear(v, 0, v.Length)

        For f = 0 To travee1.Nbre_sect - 2
            ss = ss + (-travee1.a_sect(f) + travee1.a_sect(f + 1)) / 2 * (r(f) + r(f + 1)) * (travee1.L_calc - travee1.a_sect(f + 1))
        Next
        For f = 1 To travee1.Nbre_sect - 2
            s2 = s2 + (((-travee1.a_sect(f - 1) + travee1.a_sect(f)) ^ 2) / 6 + ((-travee1.a_sect(f) + travee1.a_sect(f + 1)) ^ 2) / 3) * r(f)
        Next

        'Tableau de rotations
        'Rotation initiale
        w(0) = (-ss - s2 - ((-travee1.a_sect(0) + travee1.a_sect(1)) ^ 2) / 3 * r(0) - ((travee1.a_sect(travee1.Nbre_sect - 1) - travee1.a_sect(travee1.Nbre_sect - 2)) ^ 2) / 6 * r(travee1.Nbre_sect - 1)) / travee1.L_calc

        For f = 0 To travee1.Nbre_sect - 2
            w(f + 1) = w(f) + (-travee1.a_sect(f) + travee1.a_sect(f + 1)) / 2 * (r(f) + r(f + 1))
        Next

        'Tableau de flèche
        'Déformée à l'appui gauche
        v(0) = 0
        For f = 0 To travee1.Nbre_sect - 2
            v(f + 1) = v(f) + w(f) * (-travee1.a_sect(f) + travee1.a_sect(f + 1)) + ((-travee1.a_sect(f) + travee1.a_sect(f + 1)) ^ 2) / 6 * (2 * r(f) + r(f + 1))
        Next

        'Déformé à l'appui droite
        v(travee1.Nbre_sect - 1) = 0

        Return 0

    End Function

    Public Sub fleche()

        'Calcul de listes de moments par type de chargement G G' Q
        'Initialisation des listes de moments
        Array.Clear(travee1.T_eff_iso_mt, 0, travee1.T_eff_iso_mt.Length)
        Array.Clear(travee1.T_eff_iso_vp, 0, travee1.T_eff_iso_vp.Length)
        Array.Clear(travee1.T_eff_iso_vm, 0, travee1.T_eff_iso_vm.Length)

        '----
        'Msw
        travee1.Calc_iso_chponc(Tr, travee1.a, travee1.g, travee1.cg, travee1.cd, n, iponc)
        travee1.Calc_iso_chlin(Tr, travee1.p1_a, travee1.p2_a, travee1.p_g, n, travee1.cg, travee1.cd, ilin)
        travee1.Calc_iso_chtrap(Tr, travee1.NuAppG, travee1.NuAppG + travee1.L, travee1.chsurf_g_g, travee1.chsurf_g_d, 1)
        travee1.Calc_iso_chtrap(Tr, travee1.NuAppG, travee1.NuAppG + travee1.L, travee1.chsurf_d_g, travee1.chsurf_d_d, 1)

        ReDim Moment(travee1.Nbre_sect - 1)
        ReDim MomentHyper(travee1.Nbre_sect - 1)

        ReDim flechMsw(travee1.Nbre_sect - 1)
        ReDim flechHypMsw(travee1.Nbre_sect - 1)

        For f = 0 To travee1.Nbre_sect - 1
            Moment(f) = travee1.T_eff_iso_mt(f) * 10 ^ -3
            flechMsw(f) = travee1.T_eff_iso_mt(f) * 10 ^ -3
        Next

        Call MomentHyperFleche()

        For f = 0 To travee1.Nbre_sect - 1
            flechHypMsw(f) = MomentHyper(f)
        Next

        '------
        'Msd1

        Array.Clear(Moment, 0, Moment.Length)
        Array.Clear(MomentHyper, 0, MomentHyper.Length)

        travee1.Calc_iso_chponc(Tr, travee1.a, travee1.gp1, travee1.cg, travee1.cd, n, iponc)
        travee1.Calc_iso_chlin(Tr, travee1.p1_a, travee1.p2_a, travee1.p_gp1, n, travee1.cg, travee1.cd, ilin)
        travee1.Calc_iso_chtrap(Tr, travee1.NuAppG, travee1.NuAppG + travee1.L, travee1.chsurf_g_gp1, travee1.chsurf_g_d, 1)
        travee1.Calc_iso_chtrap(Tr, travee1.NuAppG, travee1.NuAppG + travee1.L, travee1.chsurf_d_gp1, travee1.chsurf_d_d, 1)

        ReDim flechMsd1(travee1.Nbre_sect - 1)
        ReDim flechHypMsd1(travee1.Nbre_sect - 1)

        For f = 0 To travee1.Nbre_sect - 1
            Moment(f) = travee1.T_eff_iso_mt(f) * 10 ^ -3
            flechMsd1(f) = travee1.T_eff_iso_mt(f) * 10 ^ -3
        Next

        Call MomentHyperFleche()

        For f = 0 To travee1.Nbre_sect - 1
            flechHypMsd1(f) = MomentHyper(f)
        Next

        '------
        'Msd2
        Array.Clear(Moment, 0, Moment.Length)
        Array.Clear(MomentHyper, 0, MomentHyper.Length)

        travee1.Calc_iso_chponc(Tr, travee1.a, travee1.gp2, travee1.cg, travee1.cd, n, iponc)
        travee1.Calc_iso_chlin(Tr, travee1.p1_a, travee1.p2_a, travee1.p_gp2, n, travee1.cg, travee1.cd, ilin)
        travee1.Calc_iso_chtrap(Tr, travee1.NuAppG, travee1.NuAppG + travee1.L, travee1.chsurf_g_gp2, travee1.chsurf_g_d, 1)
        travee1.Calc_iso_chtrap(Tr, travee1.NuAppG, travee1.NuAppG + travee1.L, travee1.chsurf_d_gp2, travee1.chsurf_d_d, 1)

        ReDim flechMsd2(travee1.Nbre_sect - 1)
        ReDim flechHypMsd2(travee1.Nbre_sect - 1)

        For f = 0 To travee1.Nbre_sect - 1
            Moment(f) = travee1.T_eff_iso_mt(f) * 10 ^ -3
            flechMsd2(f) = travee1.T_eff_iso_mt(f) * 10 ^ -3
        Next

        Call MomentHyperFleche()

        For f = 0 To travee1.Nbre_sect - 1
            flechHypMsd2(f) = MomentHyper(f)
        Next

        '-----
        'Mq
        Array.Clear(Moment, 0, Moment.Length)
        Array.Clear(MomentHyper, 0, MomentHyper.Length)

        'Multiplication des charge élémentaires de q par le coefficient PSI quasi permanent
        'Charges ponctuelles
        Dim qq(travee1.Nbre_sect - 1) As Double
        For i = 0 To iponc - 1
            qq(i) = coefqp * travee1.q(i)
        Next
        travee1.Calc_iso_chponc(Tr, travee1.a, qq, travee1.d1_cotes, travee1.d2_cotes, n, iponc)
        'Charges lineaires
        Dim pq(travee1.Nbre_sect - 1) As Double
        For i = 0 To ilin - 1
            pq(i) = coefqp * travee1.p_q(i)
        Next
        travee1.Calc_iso_chlin(Tr, travee1.p1_a, travee1.p2_a, pq, n, travee1.d1_cotes, travee1.d2_cotes, ilin)
        'Charge trapezoidales
        travee1.Calc_iso_chtrap(Tr, travee1.NuAppG, travee1.NuAppG + travee1.L, coefqp * travee1.chsurf_g_q, travee1.chsurf_g_d, 1)
        travee1.Calc_iso_chtrap(Tr, travee1.NuAppG, travee1.NuAppG + travee1.L, coefqp * travee1.chsurf_d_q, travee1.chsurf_d_d, 1)

        ReDim flechMq(travee1.Nbre_sect - 1)
        ReDim flechHypMq(travee1.Nbre_sect - 1)

        For f = 0 To travee1.Nbre_sect - 1
            Moment(f) = travee1.T_eff_iso_mt(f) * 10 ^ -3
            flechMq(f) = travee1.T_eff_iso_mt(f) * 10 ^ -3
        Next

        Call MomentHyperFleche()

        For f = 0 To travee1.Nbre_sect - 1
            flechHypMq(f) = MomentHyper(f)
        Next

    End Sub

    Private Function Continuite(ByVal M As Double, ByVal M0 As Double, ByVal Alpha As Double, ByVal Mw As Double, ByVal Mea As Double) As Integer
        Dim Mm As Double
        Mm = M - M0

        M2W = (Mw - Mea) * Alpha + Mm
        M2E = Mea + (Mw - Mea) * Alpha - Mw + Mm

        If M2W >= 0 Then
            M2W = 0
        End If

        If M2E >= 0 Then
            M2E = 0
        End If

        Return 0
    End Function

    Private Sub MomentHyperFleche()
        'Recherche du moment max iso et de son abcisse
        'Localisation du moment maximum en travee
        'Indice de la section à moment max
        Dim MtMax_section As Integer
        Dim Miso As Double
        Dim aMiso As Double
        For f = 0 To travee1.Nbre_sect - 1
            If Moment(f) > Miso Then
                Miso = Moment(f)
                MtMax_section = f
            End If
        Next
        aMiso = travee1.a_sect(MtMax_section)

        'Détermination des entrées pour le calcul de coef. de continuité
        Dim alpha As Double
        Dim M1W As Double
        Dim M1E As Double
        Dim M2Tr As Double

        M1W = -ContAppW * Miso
        M1E = -ContAppE * Miso
        M2Tr = ContTr * Miso
        alpha = aMiso / travee1.L_calc

        Call Continuite(M2Tr, Miso, alpha, M1W, M1E)

        'Construction de la liste des moments hyperstatiques
        For i = 0 To travee1.Nbre_sect - 1
            MomentHyper(i) = Moment(i) + M2W * (1 - travee1.a_sect(i) / travee1.L_calc) + M2E * (travee1.a_sect(i) / travee1.L_calc)
        Next

    End Sub

    Private Function TranchantGraph(Tgauche() As Double, Tdroite() As Double, Pnb As Integer, Ponca() As Double) As Integer

        Dim s As Integer = 0

        'Par défaut
        'Dimension liste pour tranchant sans charge ponctuelle
        ReDim Tiso_elu_graph(travee1.Nbre_sect - 1)
        ReDim a_graph(travee1.Nbre_sect - 1)

        'Cas de presence de charges ponctuelles
        If Pnb <> 0 Then
            ReDim Tiso_elu_graph(travee1.Nbre_sect - 1 + Pnb)
            ReDim a_graph(travee1.Nbre_sect - 1 + Pnb)

            Select Case Pnb

                Case 1
                    While (travee1.a_sect(s) < (travee1.NuAppG + Ponca(0)) And (s <= travee1.Nbre_sect))
                        Tiso_elu_graph(s) = Tdroite(s)
                        a_graph(s) = travee1.a_sect(s)
                        s = s + 1
                    End While

                    Tiso_elu_graph(s) = Tgauche(s)
                    a_graph(s) = travee1.a_sect(s)
                    Tiso_elu_graph(s + 1) = Tdroite(s)
                    a_graph(s + 1) = travee1.a_sect(s)
                    s = s + 2

                    While s <= travee1.Nbre_sect - 1
                        Tiso_elu_graph(s) = Tdroite(s - 1)
                        a_graph(s) = travee1.a_sect(s - 1)
                        s = s + 1
                    End While

                    Tiso_elu_graph(s) = Tgauche(s - 1)
                    a_graph(s) = travee1.a_sect(s - 1)


                Case 2
                    While (travee1.a_sect(s) < (travee1.NuAppG + Ponca(0)) And (s <= travee1.Nbre_sect))
                        Tiso_elu_graph(s) = Tdroite(s)
                        a_graph(s) = travee1.a_sect(s)
                        s = s + 1
                    End While

                    Tiso_elu_graph(s) = Tgauche(s)
                    a_graph(s) = travee1.a_sect(s)
                    Tiso_elu_graph(s + 1) = Tdroite(s)
                    a_graph(s + 1) = travee1.a_sect(s)
                    s = s + 2

                    While (travee1.a_sect(s - 1) < (travee1.NuAppG + Ponca(1)) And (s <= travee1.Nbre_sect))
                        Tiso_elu_graph(s) = Tdroite(s - 1)
                        a_graph(s) = travee1.a_sect(s - 1)
                        s = s + 1
                    End While

                    Tiso_elu_graph(s) = Tgauche(s - 1)
                    a_graph(s) = travee1.a_sect(s - 1)
                    Tiso_elu_graph(s + 1) = Tdroite(s - 1)
                    a_graph(s + 1) = travee1.a_sect(s - 1)
                    s = s + 2

                    While s <= travee1.Nbre_sect
                        Tiso_elu_graph(s) = Tdroite(s - 2)
                        a_graph(s) = travee1.a_sect(s - 2)
                        s = s + 1
                    End While

                    Tiso_elu_graph(s) = Tgauche(s - 2)
                    a_graph(s) = travee1.a_sect(s - 2)


                Case 3

                    While (travee1.a_sect(s) < (travee1.NuAppG + Ponca(0)) And (s <= travee1.Nbre_sect))
                        Tiso_elu_graph(s) = Tdroite(s)
                        a_graph(s) = travee1.a_sect(s)
                        s = s + 1
                    End While

                    Tiso_elu_graph(s) = Tgauche(s)
                    a_graph(s) = travee1.a_sect(s)
                    Tiso_elu_graph(s + 1) = Tdroite(s)
                    a_graph(s + 1) = travee1.a_sect(s)
                    s = s + 2

                    While (travee1.a_sect(s - 1) < (travee1.NuAppG + Ponca(1)) And (s <= travee1.Nbre_sect))
                        Tiso_elu_graph(s) = Tdroite(s - 1)
                        a_graph(s) = travee1.a_sect(s - 1)
                        s = s + 1
                    End While

                    Tiso_elu_graph(s) = Tgauche(s - 1)
                    a_graph(s) = travee1.a_sect(s - 1)
                    Tiso_elu_graph(s + 1) = Tdroite(s - 1)
                    a_graph(s + 1) = travee1.a_sect(s - 1)
                    s = s + 2

                    While (travee1.a_sect(s - 2) < (travee1.NuAppG + Ponca(2)) And (s <= travee1.Nbre_sect))
                        Tiso_elu_graph(s) = Tdroite(s - 2)
                        a_graph(s) = travee1.a_sect(s - 2)
                        s = s + 1
                    End While

                    Tiso_elu_graph(s) = Tgauche(s - 2)
                    a_graph(s) = travee1.a_sect(s - 2)
                    Tiso_elu_graph(s + 1) = Tdroite(s - 2)
                    a_graph(s + 1) = travee1.a_sect(s - 2)
                    s = s + 2

                    While s <= travee1.Nbre_sect + 1
                        Tiso_elu_graph(s) = Tdroite(s - 3)
                        a_graph(s) = travee1.a_sect(s - 3)
                        s = s + 1
                    End While

                    Tiso_elu_graph(s) = Tgauche(s - 3)
                    a_graph(s) = travee1.a_sect(s - 3)

            End Select

            'Calcul de la reaction d'appui
            travee1.RAppG = Tiso_elu_graph(0)
            travee1.RAppD = Tiso_elu_graph(travee1.Nbre_sect - 1 + Pnb)

            'Cas en absence de charges ponctuelles
        Else

            While s <= travee1.Nbre_sect - 2
                Tiso_elu_graph(s) = Tdroite(s)
                a_graph(s) = travee1.a_sect(s)
                s = s + 1
            End While

            Tiso_elu_graph(s) = Tgauche(s)
            a_graph(s) = travee1.a_sect(s)

            'Calcul de la reaction d'appui
            travee1.RAppG = Tiso_elu_graph(0)
            travee1.RAppD = Tiso_elu_graph(travee1.Nbre_sect - 1)
        End If

        Return 0
    End Function



End Module
