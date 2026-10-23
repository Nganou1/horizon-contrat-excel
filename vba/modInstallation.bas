Attribute VB_Name = "modInstallation"
Option Explicit
'==============================================================================
' Horizon_Contrat - installation de la couche reporting dans le classeur
'
' InstallerReporting (Alt+F8), une fois, après l'import des modules.
' Relançable sans risque : ce qui existe déjà est laissé tel quel.
'
' Ajoute :
'   - l'onglet JOURNAL et son tableau tblJournal ;
'   - l'onglet SAISIE_RELEVES et son tableau tblSaisieReleves ;
'   - les paramètres manquants en bas de PARAMETRES (un nom par clé) ;
'   - les noms Fiche_CA_Net, Zone_Impression_Fiche, Zone_Tableau_Alerte,
'     Zone_Top_Produits ;
'   - quatre boutons en haut du DASHBOARD.
' Renseigne Chemin_Donnees et Chemin_Sorties s'ils sont encore « À RENSEIGNER ».
'==============================================================================

Private Const COULEUR_VERT As Long = 6254091        ' RGB(11, 110, 95)
Private Const COULEUR_TERRE As Long = 3102132       ' RGB(180, 85, 47)
Private Const COULEUR_GRIS As Long = 13292245       ' RGB(213, 210, 202)

Public Sub InstallerReporting()
    Dim sBilan As String, sErr As String, bScreen As Boolean

    If MsgBox("Installer la couche reporting dans ce classeur ?" & vbLf & vbLf & _
              "- onglets JOURNAL et SAISIE_RELEVES" & vbLf & _
              "- paramètres d'envoi et de saisie en bas de PARAMETRES" & vbLf & _
              "- quatre boutons en haut du DASHBOARD" & vbLf & vbLf & _
              "Rien d'existant n'est supprimé.", vbOKCancel + vbQuestion, "Horizon_Contrat") <> vbOK Then Exit Sub

    bScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo Echec

    sBilan = sBilan & CreerJournal()
    sBilan = sBilan & CreerSaisieReleves()
    sBilan = sBilan & AjouterParametres()
    sBilan = sBilan & RenseignerChemins()
    sBilan = sBilan & AjouterNoms()
    sBilan = sBilan & CreerBoutons()
    Journaliser "InstallerReporting", 1, 0, 0, Replace$(Mid$(sBilan, 2), vbLf, " ; ")

Fin:
    On Error Resume Next
    Application.ScreenUpdating = bScreen
    Feuille(ONGLET_DASHBOARD).Activate
    If sErr <> "" Then
        MsgBox "Installation interrompue : " & sErr & vbLf & vbLf & "Déjà fait :" & sBilan, vbExclamation, "Horizon_Contrat"
    Else
        If ThisWorkbook.FileFormat <> xlOpenXMLWorkbookMacroEnabled Then
            sBilan = sBilan & vbLf & vbLf & "ATTENTION : enregistre le classeur en .xlsm (Fichier > Enregistrer sous)," & _
                     " sinon les macros seront perdues à la fermeture."
        End If
        MsgBox "Installation terminée." & vbLf & sBilan, vbInformation, "Horizon_Contrat"
    End If
    Exit Sub

Echec:
    sErr = Err.Description & " (erreur " & Err.Number & ")"
    Resume Fin
End Sub

'------------------------------------------------------------------------------
Private Function CreerJournal() As String
    Dim ws As Worksheet
    If FeuilleExiste(ONGLET_JOURNAL) Then
        CreerJournal = vbLf & "- JOURNAL : déjà présent"
        Exit Function
    End If
    Set ws = NouvelOnglet(ONGLET_JOURNAL, ONGLET_PARAMETRES, COULEUR_TERRE)
    TitreOnglet ws, "JOURNAL - traces d'exécution VBA", _
                "Une ligne par exécution : qui, quoi, combien, en combien de temps, avec quelle erreur. Écrit par Journaliser (modJournal)."
    CreerTableau ws, "B4", TBL_JOURNAL, _
        Array("Horodatage", "Utilisateur", "Procedure", "Traites", "Ignores", "Duree_s", "Message"), _
        Array(19, 14, 26, 9, 9, 9, 90), _
        Array("dd/mm/yyyy hh:mm:ss", "@", "@", "0", "0", "0.0", "@")
    CreerJournal = vbLf & "- JOURNAL : créé (tblJournal)"
End Function

Private Function CreerSaisieReleves() As String
    Dim ws As Worksheet
    If FeuilleExiste(ONGLET_SAISIE) Then
        CreerSaisieReleves = vbLf & "- SAISIE_RELEVES : déjà présent"
        Exit Function
    End If
    Set ws = NouvelOnglet(ONGLET_SAISIE, ONGLET_JOURNAL, COULEUR_GRIS)
    TitreOnglet ws, "SAISIE_RELEVES - relevés journaliers saisis par le formulaire", _
                "Ne pas saisir ici à la main : le formulaire (bouton « Saisir un relevé ») contrôle chaque ligne avant de l'écrire."
    CreerTableau ws, "B4", TBL_SAISIE, _
        Array("Horodatage", "Code_Magasin", "Date_Releve", "Canal", "CA_Brut_TTC", "Remises_TTC", _
              "CA_Net_HT", "Cout_Matiere_HT", "Nb_Tickets", "Nb_Articles", "Alertes", "Saisi_Par"), _
        Array(19, 13, 12, 12, 13, 12, 12, 15, 11, 11, 60, 14), _
        Array("dd/mm/yyyy hh:mm", "@", "dd/mm/yyyy", "@", "#,##0.00", "#,##0.00", _
              "#,##0.00", "#,##0.00", "#,##0", "#,##0", "@", "@")
    CreerSaisieReleves = vbLf & "- SAISIE_RELEVES : créé (tblSaisieReleves)"
End Function

Private Function NouvelOnglet(ByVal sNom As String, ByVal sApres As String, ByVal lCouleur As Long) As Worksheet
    Dim ws As Worksheet
    If FeuilleExiste(sApres) Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(sApres))
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    End If
    ws.Name = sNom
    ws.Tab.Color = lCouleur
    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 9.5
    ws.Columns("A").ColumnWidth = 2
    ActiveWindow.DisplayGridlines = False
    Set NouvelOnglet = ws
End Function

Private Sub TitreOnglet(ByVal ws As Worksheet, ByVal sTitre As String, ByVal sSousTitre As String)
    With ws.Range("B1")
        .Value = sTitre
        .Font.Name = "Cambria"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(20, 51, 44)
    End With
    ws.Rows(1).RowHeight = 34
    ws.Range("B2").Value = sSousTitre
    ws.Range("B2").Font.Color = RGB(107, 117, 112)
End Sub

Private Sub CreerTableau(ByVal ws As Worksheet, ByVal sCoin As String, ByVal sNom As String, _
                         ByVal vEntetes As Variant, ByVal vLargeurs As Variant, ByVal vFormats As Variant)
    Dim lo As ListObject, rEntete As Range, i As Long, n As Long
    n = UBound(vEntetes) - LBound(vEntetes) + 1
    Set rEntete = ws.Range(sCoin).Resize(1, n)
    rEntete.Value = vEntetes
    Set lo = ws.ListObjects.Add(xlSrcRange, rEntete.Resize(2), , xlYes)
    lo.Name = sNom
    lo.TableStyle = "TableStyleLight15"
    For i = 1 To n
        ws.Columns(rEntete.Cells(1, i).Column).ColumnWidth = vLargeurs(LBound(vLargeurs) + i - 1)
        lo.ListColumns(i).DataBodyRange.NumberFormat = vFormats(LBound(vFormats) + i - 1)
    Next i
    ws.Activate
    ws.Range(sCoin).Offset(1, 0).Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
    ws.Range("B1").Select
End Sub

'------------------------------------------------------------------------------
' Paramètres : ajoutés en bas de PARAMETRES, un nom par clé, si absents
'------------------------------------------------------------------------------
Private Function AjouterParametres() As String
    Dim ws As Worksheet, vP As Variant, p As Variant, r As Long, n As Long, bSection As Boolean

    Set ws = Feuille(ONGLET_PARAMETRES)
    vP = Array( _
        Array("Mail_Copie_Direction", "direction@horizon-contrat.example", "Texte", "@", _
              "Copie de chaque message aux franchisés. Adresse fictive livrée (.example) : remplace-la par une vraie pour un envoi réel."), _
        Array("Domaine_Mail_Franchises", "horizon-contrat.example", "Texte", "@", _
              "Adresse d'un franchisé = code magasin en minuscules @ ce domaine (hc-002@...). Un domaine .example n'est jamais envoyé."), _
        Array("Inclure_Magasins_Fermes", False, "Booléen", "General", _
              "VRAI = une fiche PDF aussi pour les magasins fermés (période partielle). FAUX = exclus, et comptés à part dans JOURNAL."), _
        Array("Mois_Ouvert", DateSerial(Year(Date), Month(Date), 1), "Date", "dd/mm/yyyy", _
              "Premier jour du mois ouvert à la saisie des relevés. Une date hors de ce mois est refusée par le formulaire."), _
        Array("Seuil_Ecart_Ticket_Moyen", 0.25, "Fraction", "0%", _
              "Écart toléré entre le ticket moyen saisi et la moyenne des 30 derniers jours (avertissement, pas blocage)."), _
        Array("Tolerance_Cout_Matiere_Pt", 0.05, "Fraction", "0.0%", _
              "Coût matière saisi toléré jusqu'au taux de coût de l'exercice du magasin + cette marge (avertissement)."), _
        Array("Derniere_Actualisation", Empty, "Date", "dd/mm/yyyy hh:mm", _
              "Écrit par ActualiserTout : date et heure de la dernière actualisation complète."))

    r = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row + 2
    For Each p In vP
        If PlageNommee(CStr(p(0))) Is Nothing Then
            If Not bSection Then
                ws.Range("B38:E38").Copy Destination:=ws.Range("B" & r & ":E" & r)   ' mise en forme d'en-tête
                ws.Range("B" & r).Value = "Couche VBA - reporting, envoi et saisie (ajouté par InstallerReporting)"
                r = r + 1
                bSection = True
            End If
            ' Copie directe (sans presse-papiers) d'une ligne de clé existante pour
            ' en reprendre la mise en forme ; les quatre valeurs sont réécrites.
            ws.Range("B41:E41").Copy Destination:=ws.Range("B" & r & ":E" & r)
            ws.Range("B" & r).Value = p(0)
            ws.Range("C" & r).NumberFormat = p(3)
            ws.Range("C" & r).Value = p(1)
            ws.Range("D" & r).Value = p(2)
            ws.Range("E" & r).Value = p(4)
            ThisWorkbook.Names.Add Name:=CStr(p(0)), RefersTo:="='" & ws.Name & "'!$C$" & r
            r = r + 1
            n = n + 1
        End If
    Next p
    AjouterParametres = vbLf & "- PARAMETRES : " & n & " clé(s) ajoutée(s)"
End Function

Private Function RenseignerChemins() As String
    Dim s As String, sDonnees As String, sSorties As String
    sDonnees = AvecSeparateur(ThisWorkbook.Path) & "Donnees" & Application.PathSeparator
    If CheminRenseigne("Chemin_Donnees") = "" And DossierExiste(sDonnees & "02_Ventes") Then
        EcrireParam "Chemin_Donnees", sDonnees
        s = s & vbLf & "- Chemin_Donnees = " & sDonnees
    End If
    If CheminRenseigne("Chemin_Sorties") = "" Then
        sSorties = AvecSeparateur(ThisWorkbook.Path) & "Sorties" & Application.PathSeparator
        EcrireParam "Chemin_Sorties", sSorties
        s = s & vbLf & "- Chemin_Sorties = " & sSorties
    End If
    RenseignerChemins = s
End Function

Private Function AjouterNoms() As String
    Dim n As Long
    n = n + NomSiAbsent("Fiche_CA_Net", "='" & ONGLET_FICHE & "'!$G$7")
    n = n + NomSiAbsent("Zone_Impression_Fiche", "='" & ONGLET_FICHE & "'!$B$1:$Q$44")
    n = n + NomSiAbsent("Zone_Tableau_Alerte", "='" & ONGLET_DASHBOARD & "'!$B$32:$M$37")
    n = n + NomSiAbsent("Zone_Top_Produits", "='" & ONGLET_FICHE & "'!$B$21:$E$26")
    AjouterNoms = vbLf & "- noms ajoutés : " & n
End Function

Private Function NomSiAbsent(ByVal sNom As String, ByVal sRef As String) As Long
    If PlageNommee(sNom) Is Nothing Then
        ThisWorkbook.Names.Add Name:=sNom, RefersTo:=sRef
        NomSiAbsent = 1
    End If
End Function

'------------------------------------------------------------------------------
' Boutons (formes reliées à des macros, plus robustes que des contrôles ActiveX)
'------------------------------------------------------------------------------
Private Function CreerBoutons() As String
    Dim ws As Worksheet, i As Long, shp As Shape, dGauche As Double, dDroite As Double, dLarg As Double
    Dim vTextes As Variant, vMacros As Variant, bProtegee As Boolean
    Const ECART As Double = 6

    Set ws = Feuille(ONGLET_DASHBOARD)
    bProtegee = ws.ProtectContents
    If bProtegee Then ws.Unprotect MOT_DE_PASSE_FEUILLES
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 6) = "btnHC_" Then ws.Shapes(i).Delete
    Next i

    vTextes = Array("Actualiser les données", "Générer les fiches PDF", "Préparer les mails", "Saisir un relevé")
    vMacros = Array("ActualiserTout", "GenererFichesPDF", "EnvoyerFichesFranchises", "OuvrirSaisieReleve")
    dGauche = ws.Range("K1").Left
    dDroite = ws.Range("S1").Left + ws.Range("S1").Width
    dLarg = (dDroite - dGauche - 3 * ECART) / 4

    For i = 0 To 3
        Set shp = ws.Shapes.AddShape(5, dGauche + i * (dLarg + ECART), ws.Range("K1").Top + 6, dLarg, 22)
        shp.Name = "btnHC_" & vMacros(i)
        shp.OnAction = vMacros(i)
        shp.Placement = xlFreeFloating
        shp.Line.Visible = 0
        shp.Fill.ForeColor.RGB = IIf(i = 0, RGB(107, 117, 112), COULEUR_VERT)
        With shp.TextFrame2
            .VerticalAnchor = 3
            .MarginLeft = 2
            .MarginRight = 2
            .TextRange.Text = vTextes(i)
            .TextRange.ParagraphFormat.Alignment = 2
            .TextRange.Font.Size = 9
            .TextRange.Font.Bold = -1
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        End With
    Next i
    If bProtegee Then ProtegerFeuille ws
    CreerBoutons = vbLf & "- DASHBOARD : 4 boutons"
End Function
