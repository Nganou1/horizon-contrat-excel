Attribute VB_Name = "modReporting"
Option Explicit
'==============================================================================
' Horizon_Contrat - fiches magasin en PDF
'
' GenererFichesPDF : une fiche par magasin dans Sorties\AAAA-MM-JJ\.
' Magasins fermés : exclus par défaut (paramètre Inclure_Magasins_Fermes).
' HC-063, fermé le 30/04/2026, a encore 749 863 EUR de CA sur l'exercice
' (sept.-avril) : c'est la date de fermeture qui l'écarte, pas un CA nul.
' Résultat attendu avec les paramètres livrés : 119 PDF, 1 fermé, 0 sans CA.
'==============================================================================

Public Sub GenererFichesPDF()
    Dim wsFiche As Worksheet, lo As ListObject, v As Variant
    Dim rRegion As Range, rMag As Range, rCA As Range, rValide As Range
    Dim iCode As Long, iNom As Long, iRegion As Long, iFerm As Long
    Dim i As Long, n As Long, lTraites As Long, lSansCA As Long, lFermes As Long
    Dim sDossier As String, sFichier As String, sErr As String, sNote As String
    Dim sRegionAvant As String, sMagAvant As String, sFermes As String
    Dim dDebut As Double, dCA As Double, bInclureFermes As Boolean, bEtatSauve As Boolean
    Dim bScreen As Boolean, bEvents As Boolean, bAlerts As Boolean, lCalc As XlCalculation

    dDebut = Timer
    bScreen = Application.ScreenUpdating
    bEvents = Application.EnableEvents
    bAlerts = Application.DisplayAlerts
    lCalc = Application.Calculation
    ' Échap / Ctrl+Pause lèvent l'erreur 18 au lieu d'arrêter net la macro :
    ' sans ça, « Fin » dans la boîte d'interruption saute la restauration.
    Application.EnableCancelKey = xlErrorHandler
    On Error GoTo Echec

    Set wsFiche = Feuille(ONGLET_FICHE)
    Set rRegion = Plage("Region_Fiche")
    Set rMag = Plage("Mag_Selectionne")
    Set rCA = Plage("Fiche_CA_Net")
    Set rValide = Plage("F_Valide")
    Set lo = Tableau(ONGLET_REF_MAGASINS, TBL_MAGASINS)
    iCode = IndexColonne(lo, "Code_Magasin")
    iNom = IndexColonne(lo, "Nom_Magasin")
    iRegion = IndexColonne(lo, "Region")
    iFerm = IndexColonne(lo, "Date_Fermeture")
    v = lo.DataBodyRange.Value                 ' les 120 magasins en un seul accès
    n = UBound(v, 1)
    bInclureFermes = CBool(ParamOuDefaut("Inclure_Magasins_Fermes", False))

    sDossier = CreerDossierDuJour(CheminSorties())
    sRegionAvant = CStr(rRegion.Value)
    sMagAvant = CStr(rMag.Value)
    bEtatSauve = True

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    MettreEnPageFiche wsFiche

    For i = 1 To n
        If Not bInclureFermes And EstRenseigne(v(i, iFerm)) Then
            lFermes = lFermes + 1
            sFermes = sFermes & IIf(sFermes = "", "", ", ") & v(i, iCode)
        Else
            rRegion.Value = v(i, iRegion)       ' F_Valide exige un magasin de la région choisie
            rMag.Value = v(i, iCode)
            Application.Calculate

            dCA = 0
            If rValide.Value = True And IsNumeric(rCA.Value) Then dCA = CDbl(rCA.Value)

            If dCA > 0 Then
                EnTeteFiche wsFiche, CStr(v(i, iCode)) & " - " & CStr(v(i, iNom))
                sFichier = sDossier & "Fiche_" & NomFichierSain(CStr(v(i, iCode))) & ".pdf"
                wsFiche.ExportAsFixedFormat Type:=xlTypePDF, Filename:=sFichier, _
                    Quality:=xlQualityStandard, IncludeDocProperties:=False, _
                    IgnorePrintAreas:=False, OpenAfterPublish:=False
                lTraites = lTraites + 1
            Else
                lSansCA = lSansCA + 1
            End If
        End If

        If i Mod 5 = 0 Or i = n Then              ' pas à chaque tour : la barre d'état coûte cher
            Application.StatusBar = "Horizon_Contrat - fiches : " & i & " / " & n & _
                                    " (" & lTraites & " PDF)  -  Échap pour interrompre"
            DoEvents
        End If
    Next i

    sNote = IIf(bInclureFermes, "magasins fermés inclus", _
                lFermes & " fermé(s) exclu(s)" & IIf(sFermes <> "", " : " & sFermes, "")) & _
            " ; " & lSansCA & " sans CA ; dossier " & sDossier
    Journaliser "GenererFichesPDF", lTraites, lFermes + lSansCA, Timer - dDebut, sNote

Fin:
    On Error Resume Next
    If bEtatSauve Then
        rRegion.Value = sRegionAvant
        rMag.Value = sMagAvant
    End If
    Application.ScreenUpdating = bScreen
    Application.EnableEvents = bEvents
    Application.DisplayAlerts = bAlerts
    Application.Calculation = lCalc
    If lCalc = xlCalculationManual Then Application.Calculate
    Application.StatusBar = False
    Application.EnableCancelKey = xlInterrupt

    If sErr <> "" Then
        Journaliser "GenererFichesPDF", lTraites, lFermes + lSansCA, Timer - dDebut, "ERREUR : " & sErr
        MsgBox "Interrompu après " & lTraites & " fiche(s) : " & sErr & vbLf & vbLf & _
               "Excel est revenu à son état normal (affichage, événements, mode de calcul).", _
               vbExclamation, "Horizon_Contrat"
    Else
        If MsgBox(lTraites & " fiche(s) PDF générée(s) en " & FormatDuree(Timer - dDebut) & "." & vbLf & _
                  lFermes & " magasin(s) fermé(s) exclu(s)" & IIf(sFermes <> "", " (" & sFermes & ")", "") & _
                  ", " & lSansCA & " sans CA." & vbLf & vbLf & _
                  "Dossier : " & sDossier & vbLf & vbLf & "Ouvrir le dossier ?", _
                  vbYesNo + vbInformation, "Horizon_Contrat") = vbYes Then
            ThisWorkbook.FollowHyperlink sDossier
        End If
    End If
    Exit Sub

Echec:
    If Err.Number = 18 Then
        sErr = "interrompu par l'utilisateur"
    Else
        sErr = Err.Description & " (erreur " & Err.Number & ")"
    End If
    Resume Fin
End Sub

'------------------------------------------------------------------------------
' Mise en page, réglée une fois avant la boucle
'------------------------------------------------------------------------------
Private Sub MettreEnPageFiche(ByVal ws As Worksheet)
    Dim rZone As Range, rTop As Range, sExercice As String
    Set rZone = PlageNommee("Zone_Impression_Fiche")
    sExercice = CStr(ParamOuDefaut("Exercice_Courant", ""))

    Set rTop = PlageNommee("Zone_Top_Produits")

    ' Sans imprimante installée, Excel refuse certaines propriétés de PageSetup.
    ' L'export PDF fonctionne quand même : on n'en fait pas une erreur bloquante.
    On Error Resume Next
    ' Le surlignage de ligne (Worksheet_SelectionChange) ne doit pas finir dans les PDF.
    If Not rTop Is Nothing Then rTop.Interior.Pattern = xlNone
    Application.PrintCommunication = False
    With ws.PageSetup
        If Not rZone Is Nothing Then .PrintArea = rZone.Address
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(1)
        .TopMargin = Application.CentimetersToPoints(1.6)
        .BottomMargin = Application.CentimetersToPoints(1.3)
        .HeaderMargin = Application.CentimetersToPoints(0.6)
        .FooterMargin = Application.CentimetersToPoints(0.5)
        .RightHeader = "&""Segoe UI""&9Exercice " & sExercice
        .LeftFooter = "&""Segoe UI""&8Horizon_Contrat - document interne - généré le " & Format$(Date, "dd/mm/yyyy")
        .RightFooter = "&""Segoe UI""&8Page &P / &N"
    End With
    Application.PrintCommunication = True
    On Error GoTo 0
End Sub

' En-tête avec le nom du magasin, mis à jour à chaque fiche.
' &B (gras) plutôt que "Cambria,Bold" : le nom du style de police dépend de la
' langue d'Excel ("Gras" en français), &B non.
Private Sub EnTeteFiche(ByVal ws As Worksheet, ByVal sTexte As String)
    On Error Resume Next
    Application.PrintCommunication = False
    ws.PageSetup.LeftHeader = "&""Cambria""&B&11" & Replace$(sTexte, "&", "&&")
    Application.PrintCommunication = True
    On Error GoTo 0
End Sub

Private Function EstRenseigne(ByVal v As Variant) As Boolean
    If IsEmpty(v) Or IsError(v) Then Exit Function
    EstRenseigne = (Len(Trim$(CStr(v))) > 0)
End Function

'------------------------------------------------------------------------------
' Ouvre la fiche d'un magasin (double-clic du tableau d'alerte du DASHBOARD)
'------------------------------------------------------------------------------
Public Sub AfficherFiche(ByVal sCode As String)
    Dim lo As ListObject, v As Variant, i As Long, iCode As Long, iRegion As Long
    If Trim$(sCode) = "" Then Exit Sub
    Set lo = Tableau(ONGLET_REF_MAGASINS, TBL_MAGASINS)
    iCode = IndexColonne(lo, "Code_Magasin")
    iRegion = IndexColonne(lo, "Region")
    v = lo.DataBodyRange.Value
    For i = 1 To UBound(v, 1)
        If StrComp(CStr(v(i, iCode)), sCode, vbTextCompare) = 0 Then
            Plage("Region_Fiche").Value = v(i, iRegion)
            Plage("Mag_Selectionne").Value = v(i, iCode)
            Feuille(ONGLET_FICHE).Activate
            Exit Sub
        End If
    Next i
End Sub
