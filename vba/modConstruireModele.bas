Attribute VB_Name = "modConstruireModele"
Option Explicit
'==============================================================================
' Horizon_Contrat - construction du modèle Power Query + Power Pivot
'
' Lancer : ConstruireModele (Alt+F8). Durée : 2 à 5 minutes.
' Relançable : l'ancien modèle (requêtes, tables, relations, mesures, TCD de
' contrôle) est supprimé puis reconstruit à l'identique.
'
' Ce que la macro crée :
'   - 19 requêtes Power Query (code M également lisible dans PowerQuery_M\) ;
'   - 11 tables chargées dans le modèle de données (connexion uniquement) ;
'   - 17 relations ;
'   - 31 mesures DAX (listées dans Mesures_DAX.txt) ;
'   - une feuille TCD_CONTROLE avec un TCD branché sur le modèle.
'
' Ce qu'elle ne fait pas (l'objet VBA d'Excel ne l'expose pas) :
'   - marquer Dim_Dates comme table de dates ;
'   - trier Mois_Nom par Mois_Num.
'   Ces deux réglages se font en deux clics dans Power Pivot (voir LISEZ-MOI).
'
' Prérequis : Excel 365 ou 2021 pour Windows, classeur enregistré en .xlsm
' dans le dossier qui contient le sous-dossier Donnees.
'==============================================================================

Private Const PREFIXE_CNX As String = "Requete - "
Private m_sEchecs As String
Private m_bPointVirgule As Boolean

Public Sub ConstruireModele()
    Dim t0 As Double, sChemin As String
    Dim nTables As Long, nRel As Long, nMes As Long, nLignes As Double
    Dim lCalc As XlCalculation, bScreen As Boolean, sErr As String

    If Val(Application.Version) < 16 Then
        MsgBox "Cette macro demande Excel 2016 ou plus récent (Power Query intégré).", vbExclamation, "Horizon_Contrat"
        Exit Sub
    End If
    If ThisWorkbook.Path = "" Then
        MsgBox "Enregistre d'abord le classeur en .xlsm, dans le dossier qui contient Donnees.", vbExclamation, "Horizon_Contrat"
        Exit Sub
    End If

    sChemin = VerifierCheminDonnees()
    If sChemin = "" Then Exit Sub

    If MsgBox("Construction du modèle à partir de :" & vbLf & sChemin & vbLf & vbLf & _
              "Les requêtes, le modèle de données et la feuille TCD_CONTROLE existants seront remplacés." & vbLf & _
              "Durée : 2 à 5 minutes. Continuer ?", vbOKCancel + vbQuestion, "Horizon_Contrat") <> vbOK Then Exit Sub

    t0 = Timer
    m_sEchecs = ""
    m_bPointVirgule = False
    lCalc = Application.Calculation
    bScreen = Application.ScreenUpdating
    Application.EnableCancelKey = xlErrorHandler
    On Error GoTo Echec

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Etape "1/6 - Suppression de l'ancien modèle..."
    SupprimerModele

    Etape "2/6 - Création des requêtes Power Query..."
    ThisWorkbook.Queries.FastCombine = True
    CreerRequetes

    Etape "3/6 - Chargement des tables dans le modèle (le plus long)..."
    nTables = ChargerTablesModele()

    Etape "4/6 - Création des relations..."
    nRel = CreerRelations()

    Etape "5/6 - Création des mesures DAX..."
    nMes = CreerMesures()

    Etape "6/6 - TCD de contrôle..."
    CreerTCDControle

    On Error Resume Next
    nLignes = TableModele("Fait_Ventes").RecordCount
    On Error GoTo Echec

Fin:
    On Error Resume Next
    Application.Calculation = lCalc
    Application.ScreenUpdating = bScreen
    Application.StatusBar = False
    Application.EnableCancelKey = xlInterrupt
    If sErr <> "" Then
        MsgBox "Construction interrompue : " & sErr & vbLf & vbLf & _
               "Étape en cours au moment de l'erreur : voir la barre d'état ou relancer." & _
               IIf(m_sEchecs <> "", vbLf & vbLf & "Détails :" & m_sEchecs, ""), vbExclamation, "Horizon_Contrat"
    Else
        MsgBox "Modèle construit en " & Format$(Timer - t0, "0") & " secondes." & vbLf & vbLf & _
               "Tables chargées : " & nTables & " / 11" & vbLf & _
               "Relations : " & nRel & " / 17" & vbLf & _
               "Mesures DAX : " & nMes & " / 31" & vbLf & _
               "Lignes de Fait_Ventes : " & Format$(nLignes, "#,##0") & " (attendu : 288 844)" & vbLf & vbLf & _
               "À faire à la main dans Power Pivot (Données > Gérer le modèle de données) :" & vbLf & _
               "  1. Dim_Dates > Conception > Marquer comme table de dates (colonne Date)" & vbLf & _
               "  2. Dim_Dates > colonne Mois_Nom > Trier par colonne > Mois_Num" & _
               IIf(m_sEchecs <> "", vbLf & vbLf & "Éléments non créés :" & m_sEchecs, ""), _
               IIf(m_sEchecs = "", vbInformation, vbExclamation), "Horizon_Contrat"
        On Error Resume Next
        ThisWorkbook.Worksheets("TCD_CONTROLE").Activate
    End If
    Exit Sub

Echec:
    sErr = Err.Description & " (erreur " & Err.Number & ")"
    Resume Fin
End Sub

'------------------------------------------------------------------------------
' Chemin des données
'------------------------------------------------------------------------------
Private Function VerifierCheminDonnees() As String
    Dim sActuel As String, sDefaut As String, sSaisi As String
    On Error Resume Next
    sActuel = CStr(ThisWorkbook.Names("Chemin_Donnees").RefersToRange.Value)
    On Error GoTo 0
    sActuel = AvecSeparateur(sActuel)
    If DossierValide(sActuel) Then
        VerifierCheminDonnees = sActuel
        Exit Function
    End If
    sDefaut = AvecSeparateur(ThisWorkbook.Path) & "Donnees" & Application.PathSeparator
    If Not DossierValide(sDefaut) Then
        sSaisi = InputBox("Dossier des données introuvable." & vbLf & _
                          "Colle le chemin du dossier Donnees (celui qui contient 02_Ventes) :", _
                          "Horizon_Contrat", sDefaut)
        If sSaisi = "" Then Exit Function
        sDefaut = AvecSeparateur(sSaisi)
        If Not DossierValide(sDefaut) Then
            MsgBox "Ce dossier ne contient pas 02_Ventes : " & sDefaut, vbExclamation, "Horizon_Contrat"
            Exit Function
        End If
    End If
    ThisWorkbook.Names("Chemin_Donnees").RefersToRange.Value = sDefaut
    VerifierCheminDonnees = sDefaut
End Function

Private Function DossierValide(ByVal s As String) As Boolean
    If Len(s) < 3 Then Exit Function
    On Error Resume Next
    DossierValide = (Dir(s & "02_Ventes", vbDirectory) <> vbNullString) And _
                    (Dir(s & "01_Referentiels" & Application.PathSeparator & "Referentiels_Horizon_Contrat.xlsx") <> vbNullString)
End Function

Private Function AvecSeparateur(ByVal s As String) As String
    s = Trim$(s)
    If s <> "" And Right$(s, 1) <> Application.PathSeparator Then s = s & Application.PathSeparator
    AvecSeparateur = s
End Function

Private Sub Etape(ByVal sTexte As String)
    Application.StatusBar = "Horizon_Contrat - " & sTexte
    DoEvents
End Sub

'------------------------------------------------------------------------------
' Suppression de l'existant
'------------------------------------------------------------------------------
Private Sub SupprimerModele()
    Dim i As Long, v As Variant
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("TCD_CONTROLE").Delete
    For i = ThisWorkbook.Connections.Count To 1 Step -1
        If Left$(ThisWorkbook.Connections(i).Name, Len(PREFIXE_CNX)) = PREFIXE_CNX Then ThisWorkbook.Connections(i).Delete
    Next i
    For Each v In NomsRequetes(True)
        ThisWorkbook.Queries(CStr(v)).Delete
    Next v
    Err.Clear
    On Error GoTo 0
    Application.DisplayAlerts = True
End Sub

'------------------------------------------------------------------------------
' Requêtes Power Query
'------------------------------------------------------------------------------
Private Sub AjouterRequete(ByVal sNom As String, ByVal sFormule As String, ByVal sDescription As String)
    On Error Resume Next
    ThisWorkbook.Queries(sNom).Delete
    On Error GoTo 0
    ThisWorkbook.Queries.Add sNom, sFormule, sDescription
End Sub

Private Function ChargerTablesModele() As Long
    Dim v As Variant, sNom As String, t As ModelTable, n As Long
    For Each v In Array("Fait_Ventes", "Dim_Magasins", "Dim_Produits", "Dim_Fournisseurs", "Dim_Employes", "Fait_Budget", "Fait_Mix_Produits", "Fait_Stocks", "Fait_Achats", "Fait_RH", "Dim_Dates")
        sNom = CStr(v)
        Etape "3/6 - Chargement de " & sNom & "..."
        On Error Resume Next
        ThisWorkbook.Connections.Add2 PREFIXE_CNX & sNom, "Connexion à la requête " & sNom, _
            "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & sNom & ";Extended Properties=""""", _
            """" & sNom & """", 6, True, False
        If Err.Number <> 0 Then
            m_sEchecs = m_sEchecs & vbLf & "Table " & sNom & " : " & Err.Description
            Err.Clear
        Else
            Set t = TableModele(sNom)
            If Not t Is Nothing Then
                If t.RecordCount = 0 Then ThisWorkbook.Connections(PREFIXE_CNX & sNom).Refresh
                n = n + 1
            Else
                m_sEchecs = m_sEchecs & vbLf & "Table " & sNom & " : connexion créée mais table absente du modèle"
            End If
        End If
        Err.Clear
        On Error GoTo 0
    Next v
    ChargerTablesModele = n
End Function

Private Function TableModele(ByVal sNom As String) As ModelTable
    Dim t As ModelTable, s As String
    For Each t In ThisWorkbook.Model.ModelTables
        If StrComp(t.Name, sNom, vbTextCompare) = 0 Then
            Set TableModele = t
            Exit Function
        End If
    Next t
    For Each t In ThisWorkbook.Model.ModelTables
        s = ""
        On Error Resume Next
        s = t.SourceWorkbookConnection.Name
        On Error GoTo 0
        If StrComp(s, PREFIXE_CNX & sNom, vbTextCompare) = 0 Then
            Set TableModele = t
            Exit Function
        End If
    Next t
End Function

Private Function ColonneModele(ByVal t As ModelTable, ByVal sCol As String) As ModelTableColumn
    Dim c As ModelTableColumn
    For Each c In t.ModelTableColumns
        If StrComp(c.Name, sCol, vbTextCompare) = 0 Then
            Set ColonneModele = c
            Exit Function
        End If
    Next c
End Function

'------------------------------------------------------------------------------
' Relations
'------------------------------------------------------------------------------
Private Function AjouterRelation(ByVal sFait As String, ByVal sColFait As String, _
                                 ByVal sDim As String, ByVal sColDim As String) As Boolean
    Dim tF As ModelTable, tD As ModelTable, cF As ModelTableColumn, cD As ModelTableColumn
    Set tF = TableModele(sFait)
    Set tD = TableModele(sDim)
    If tF Is Nothing Or tD Is Nothing Then
        m_sEchecs = m_sEchecs & vbLf & "Relation " & sFait & " > " & sDim & " : table absente"
        Exit Function
    End If
    Set cF = ColonneModele(tF, sColFait)
    Set cD = ColonneModele(tD, sColDim)
    If cF Is Nothing Or cD Is Nothing Then
        m_sEchecs = m_sEchecs & vbLf & "Relation " & sFait & "[" & sColFait & "] > " & sDim & "[" & sColDim & "] : colonne absente"
        Exit Function
    End If
    On Error Resume Next
    ThisWorkbook.Model.ModelRelationships.Add cF, cD
    If Err.Number <> 0 Then
        m_sEchecs = m_sEchecs & vbLf & "Relation " & sFait & "[" & sColFait & "] > " & sDim & " : " & Err.Description
    Else
        AjouterRelation = True
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Function CreerRelations() As Long
    Dim n As Long
    If AjouterRelation("Fait_Ventes", "Date", "Dim_Dates", "Date") Then n = n + 1
    If AjouterRelation("Fait_Budget", "Mois", "Dim_Dates", "Date") Then n = n + 1
    If AjouterRelation("Fait_Mix_Produits", "Mois", "Dim_Dates", "Date") Then n = n + 1
    If AjouterRelation("Fait_Stocks", "Mois", "Dim_Dates", "Date") Then n = n + 1
    If AjouterRelation("Fait_Achats", "Date_Commande", "Dim_Dates", "Date") Then n = n + 1
    If AjouterRelation("Fait_RH", "Mois", "Dim_Dates", "Date") Then n = n + 1
    If AjouterRelation("Fait_Ventes", "Code_Magasin", "Dim_Magasins", "Code_Magasin") Then n = n + 1
    If AjouterRelation("Fait_Budget", "Code_Magasin", "Dim_Magasins", "Code_Magasin") Then n = n + 1
    If AjouterRelation("Fait_Mix_Produits", "Code_Magasin", "Dim_Magasins", "Code_Magasin") Then n = n + 1
    If AjouterRelation("Fait_Stocks", "Code_Magasin", "Dim_Magasins", "Code_Magasin") Then n = n + 1
    If AjouterRelation("Fait_Achats", "Code_Magasin", "Dim_Magasins", "Code_Magasin") Then n = n + 1
    If AjouterRelation("Fait_RH", "Code_Magasin", "Dim_Magasins", "Code_Magasin") Then n = n + 1
    If AjouterRelation("Fait_Mix_Produits", "Code_Produit", "Dim_Produits", "Code_Produit") Then n = n + 1
    If AjouterRelation("Fait_Stocks", "Code_Produit", "Dim_Produits", "Code_Produit") Then n = n + 1
    If AjouterRelation("Fait_Achats", "Code_Produit", "Dim_Produits", "Code_Produit") Then n = n + 1
    If AjouterRelation("Fait_Achats", "Code_Fournisseur", "Dim_Fournisseurs", "Code_Fournisseur") Then n = n + 1
    If AjouterRelation("Fait_RH", "Matricule", "Dim_Employes", "Matricule") Then n = n + 1
    CreerRelations = n
End Function

'------------------------------------------------------------------------------
' Mesures DAX
'------------------------------------------------------------------------------
Private Function FmtEuro(ByVal nDec As Long) As Object
    Dim o As Object
    On Error Resume Next
    Set o = ThisWorkbook.Model.ModelFormatCurrency("€", nDec)
    If o Is Nothing Then Set o = ThisWorkbook.Model.ModelFormatDecimalNumber(True, nDec)
    Set FmtEuro = o
End Function

Private Function FmtPct(ByVal nDec As Long) As Object
    Set FmtPct = ThisWorkbook.Model.ModelFormatPercentageNumber(False, nDec)
End Function

Private Function FmtDec(ByVal nDec As Long) As Object
    Set FmtDec = ThisWorkbook.Model.ModelFormatDecimalNumber(True, nDec)
End Function

Private Function FmtEntier() As Object
    Set FmtEntier = ThisWorkbook.Model.ModelFormatWholeNumber(True)
End Function

Private Function AjouterMesure(ByVal sNom As String, ByVal sFormule As String, ByVal oFormat As Object) As Boolean
    Dim tb As ModelTable, sAlt As String
    Set tb = TableModele("Fait_Ventes")
    If tb Is Nothing Then
        m_sEchecs = m_sEchecs & vbLf & "Mesure " & sNom & " : table Fait_Ventes absente"
        Exit Function
    End If
    sAlt = Replace(sFormule, ",", ";")
    On Error Resume Next
    If m_bPointVirgule Then
        ThisWorkbook.Model.ModelMeasures.Add sNom, tb, sAlt, oFormat
    Else
        ThisWorkbook.Model.ModelMeasures.Add sNom, tb, sFormule, oFormat
        If Err.Number <> 0 Then
            Err.Clear
            ThisWorkbook.Model.ModelMeasures.Add sNom, tb, sAlt, oFormat
            If Err.Number = 0 Then m_bPointVirgule = True
        End If
    End If
    If Err.Number <> 0 Then
        m_sEchecs = m_sEchecs & vbLf & "Mesure " & sNom & " : " & Err.Description
    Else
        AjouterMesure = True
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Function CreerMesures() As Long
    Dim n As Long, f As String
    ' --- Base
    f = ""
    f = f & "SUM ( Fait_Ventes[CA_Brut_HT] )" & vbLf
    If AjouterMesure("CA Brut HT", f, FmtEuro(0)) Then n = n + 1
    f = ""
    f = f & "SUM ( Fait_Ventes[Remises_HT] )" & vbLf
    If AjouterMesure("Remises HT", f, FmtEuro(0)) Then n = n + 1
    f = ""
    f = f & "[CA Brut HT] - [Remises HT]" & vbLf
    If AjouterMesure("CA Net HT", f, FmtEuro(0)) Then n = n + 1
    f = ""
    f = f & "SUM ( Fait_Ventes[Cout_Matiere_HT] )" & vbLf
    If AjouterMesure("Coût Matière HT", f, FmtEuro(0)) Then n = n + 1
    f = ""
    f = f & "[CA Net HT] - [Coût Matière HT]" & vbLf
    If AjouterMesure("Marge Brute HT", f, FmtEuro(0)) Then n = n + 1
    f = ""
    f = f & "DIVIDE ( [Marge Brute HT], [CA Net HT] )" & vbLf
    If AjouterMesure("Taux de Marge", f, FmtPct(2)) Then n = n + 1
    ' --- Volumes
    f = ""
    f = f & "SUM ( Fait_Ventes[Nb_Tickets] )" & vbLf
    If AjouterMesure("Tickets", f, FmtEntier()) Then n = n + 1
    f = ""
    f = f & "SUM ( Fait_Ventes[Nb_Articles] )" & vbLf
    If AjouterMesure("Articles", f, FmtEntier()) Then n = n + 1
    f = ""
    f = f & "DIVIDE ( [CA Net HT], [Tickets] )" & vbLf
    If AjouterMesure("Ticket Moyen", f, FmtEuro(2)) Then n = n + 1
    f = ""
    f = f & "DIVIDE ( [Articles], [Tickets] )" & vbLf
    If AjouterMesure("Articles par Ticket", f, FmtDec(2)) Then n = n + 1
    f = ""
    f = f & "DISTINCTCOUNT ( Fait_Ventes[Code_Magasin] )" & vbLf
    If AjouterMesure("Nb Magasins Actifs", f, FmtEntier()) Then n = n + 1
    ' --- Temps
    f = ""
    f = f & "CALCULATE ( [CA Net HT], SAMEPERIODLASTYEAR ( Dim_Dates[Date] ) )" & vbLf
    If AjouterMesure("CA Net N-1", f, FmtEuro(0)) Then n = n + 1
    f = ""
    f = f & "DIVIDE ( [CA Net HT] - [CA Net N-1], [CA Net N-1] )" & vbLf
    If AjouterMesure(ChrW(916) & " CA Net %", f, FmtPct(1)) Then n = n + 1
    f = ""
    f = f & "VAR Ex = MAX ( Dim_Dates[Annee_Exercice] )" & vbLf
    f = f & "VAR MoisEx = MAX ( Dim_Dates[Mois_Exercice] )" & vbLf
    f = f & "RETURN" & vbLf
    f = f & "    CALCULATE (" & vbLf
    f = f & "        [CA Net HT]," & vbLf
    f = f & "        ALL ( Dim_Dates )," & vbLf
    f = f & "        Dim_Dates[Annee_Exercice] = Ex," & vbLf
    f = f & "        Dim_Dates[Mois_Exercice] <= MoisEx" & vbLf
    f = f & "    )" & vbLf
    If AjouterMesure("CA Net Cumul Exercice", f, FmtEuro(0)) Then n = n + 1
    f = ""
    f = f & "CALCULATE ( [CA Net HT], DATESINPERIOD ( Dim_Dates[Date], MAX ( Dim_Dates[Date] ), -12, MONTH ) )" & vbLf
    If AjouterMesure("CA Net 12 Mois Glissants", f, FmtEuro(0)) Then n = n + 1
    f = ""
    f = f & "CALCULATE ( [Taux de Marge], SAMEPERIODLASTYEAR ( Dim_Dates[Date] ) )" & vbLf
    If AjouterMesure("Taux de Marge N-1", f, FmtPct(2)) Then n = n + 1
    f = ""
    f = f & "IF ( ISBLANK ( [Taux de Marge N-1] ) || ISBLANK ( [Taux de Marge] ), BLANK (), ( [Taux de Marge] - [Taux de Marge N-1] ) * 100 )" & vbLf
    If AjouterMesure(ChrW(916) & " Taux de Marge pt", f, FmtDec(2)) Then n = n + 1
    f = ""
    f = f & "CALCULATE ( [CA Net HT], DATEADD ( Dim_Dates[Date], -1, MONTH ) )" & vbLf
    If AjouterMesure("CA Net Mois Précédent", f, FmtEuro(0)) Then n = n + 1
    f = ""
    f = f & "DIVIDE ( [CA Net HT] - [CA Net Mois Précédent], [CA Net Mois Précédent] )" & vbLf
    If AjouterMesure(ChrW(916) & " CA Net Mois %", f, FmtPct(1)) Then n = n + 1
    ' --- Périmètre
    f = ""
    f = f & "VAR Debut = MIN ( Dim_Dates[Date] )" & vbLf
    f = f & "VAR Fin = MAX ( Dim_Dates[Date] )" & vbLf
    f = f & "VAR DebutN1 = EDATE ( Debut, -12 )" & vbLf
    f = f & "VAR Comparables =" & vbLf
    f = f & "    FILTER (" & vbLf
    f = f & "        ALL ( Dim_Magasins )," & vbLf
    f = f & "        Dim_Magasins[Date_Ouverture] <= DebutN1" & vbLf
    f = f & "            && ( ISBLANK ( Dim_Magasins[Date_Fermeture] )" & vbLf
    f = f & "                 || Dim_Magasins[Date_Fermeture] >= Fin )" & vbLf
    f = f & "    )" & vbLf
    f = f & "RETURN" & vbLf
    f = f & "    CALCULATE ( [CA Net HT], KEEPFILTERS ( Comparables ) )" & vbLf
    If AjouterMesure("CA Net Périmètre Comparable", f, FmtEuro(0)) Then n = n + 1
    f = ""
    f = f & "VAR Debut = MIN ( Dim_Dates[Date] )" & vbLf
    f = f & "VAR Fin = MAX ( Dim_Dates[Date] )" & vbLf
    f = f & "VAR DebutN1 = EDATE ( Debut, -12 )" & vbLf
    f = f & "RETURN" & vbLf
    f = f & "    COUNTROWS (" & vbLf
    f = f & "    FILTER (" & vbLf
    f = f & "        ALL ( Dim_Magasins )," & vbLf
    f = f & "        Dim_Magasins[Date_Ouverture] <= DebutN1" & vbLf
    f = f & "            && ( ISBLANK ( Dim_Magasins[Date_Fermeture] )" & vbLf
    f = f & "                 || Dim_Magasins[Date_Fermeture] >= Fin )" & vbLf
    f = f & "    )" & vbLf
    f = f & "    )" & vbLf
    If AjouterMesure("Nb Magasins Comparables", f, FmtEntier()) Then n = n + 1
    f = ""
    f = f & "VAR Debut = MIN ( Dim_Dates[Date] )" & vbLf
    f = f & "VAR Fin = MAX ( Dim_Dates[Date] )" & vbLf
    f = f & "VAR DebutN1 = EDATE ( Debut, -12 )" & vbLf
    f = f & "VAR Comparables =" & vbLf
    f = f & "    FILTER (" & vbLf
    f = f & "        ALL ( Dim_Magasins )," & vbLf
    f = f & "        Dim_Magasins[Date_Ouverture] <= DebutN1" & vbLf
    f = f & "            && ( ISBLANK ( Dim_Magasins[Date_Fermeture] )" & vbLf
    f = f & "                 || Dim_Magasins[Date_Fermeture] >= Fin )" & vbLf
    f = f & "    )" & vbLf
    f = f & "VAR Courant = CALCULATE ( [CA Net HT], KEEPFILTERS ( Comparables ) )" & vbLf
    f = f & "VAR Precedent =" & vbLf
    f = f & "    CALCULATE ( [CA Net HT], KEEPFILTERS ( Comparables ), SAMEPERIODLASTYEAR ( Dim_Dates[Date] ) )" & vbLf
    f = f & "RETURN" & vbLf
    f = f & "    DIVIDE ( Courant - Precedent, Precedent )" & vbLf
    If AjouterMesure(ChrW(916) & " CA Comparable %", f, FmtPct(1)) Then n = n + 1
    ' --- Budget
    f = ""
    f = f & "CALCULATE (" & vbLf
    f = f & "    SUM ( Fait_Budget[Budget_CA_HT] )," & vbLf
    f = f & "    ALL ( Dim_Dates )," & vbLf
    f = f & "    TREATAS ( VALUES ( Dim_Dates[Mois_Debut] ), Fait_Budget[Mois] )" & vbLf
    f = f & ")" & vbLf
    If AjouterMesure("Budget CA HT", f, FmtEuro(0)) Then n = n + 1
    f = ""
    f = f & "VAR MoisBudgetes =" & vbLf
    f = f & "    FILTER (" & vbLf
    f = f & "        VALUES ( Dim_Dates[Mois_Debut] )," & vbLf
    f = f & "        NOT ISBLANK ( CALCULATE ( [Budget CA HT], ALL ( Dim_Magasins ) ) )" & vbLf
    f = f & "    )" & vbLf
    f = f & "RETURN" & vbLf
    f = f & "    IF (" & vbLf
    f = f & "        ISBLANK ( [Budget CA HT] )," & vbLf
    f = f & "        BLANK ()," & vbLf
    f = f & "        CALCULATE ( [CA Net HT], MoisBudgetes ) - [Budget CA HT]" & vbLf
    f = f & "    )" & vbLf
    If AjouterMesure("Écart au Budget €", f, FmtEuro(0)) Then n = n + 1
    f = ""
    f = f & "VAR MoisBudgetes =" & vbLf
    f = f & "    FILTER (" & vbLf
    f = f & "        VALUES ( Dim_Dates[Mois_Debut] )," & vbLf
    f = f & "        NOT ISBLANK ( CALCULATE ( [Budget CA HT], ALL ( Dim_Magasins ) ) )" & vbLf
    f = f & "    )" & vbLf
    f = f & "RETURN" & vbLf
    f = f & "    DIVIDE ( CALCULATE ( [CA Net HT], MoisBudgetes ), [Budget CA HT] )" & vbLf
    If AjouterMesure("Atteinte Budget %", f, FmtPct(1)) Then n = n + 1
    f = ""
    f = f & "COUNTROWS (" & vbLf
    f = f & "    FILTER (" & vbLf
    f = f & "        VALUES ( Dim_Magasins[Code_Magasin] )," & vbLf
    f = f & "        NOT ISBLANK ( [Budget CA HT] ) && [Atteinte Budget %] < 9 / 10" & vbLf
    f = f & "    )" & vbLf
    f = f & ")" & vbLf
    If AjouterMesure("Nb Magasins Sous Objectif", f, FmtEntier()) Then n = n + 1
    ' --- Opérations
    f = ""
    f = f & "DIVIDE ( SUM ( Fait_Stocks[Pertes] ), SUM ( Fait_Stocks[Sorties] ) )" & vbLf
    If AjouterMesure("Taux de Perte Stock", f, FmtPct(2)) Then n = n + 1
    f = ""
    f = f & "AVERAGE ( Fait_Achats[Retard_Jours] )" & vbLf
    If AjouterMesure("Retard Livraison Moyen", f, FmtDec(2)) Then n = n + 1
    f = ""
    f = f & "DIVIDE (" & vbLf
    f = f & "    CALCULATE ( COUNTROWS ( Fait_Achats ), Fait_Achats[Retard_Jours] > 0 )," & vbLf
    f = f & "    CALCULATE ( COUNTROWS ( Fait_Achats ), FILTER ( Fait_Achats, NOT ISBLANK ( Fait_Achats[Retard_Jours] ) ) )" & vbLf
    f = f & ")" & vbLf
    If AjouterMesure("Taux de Livraisons en Retard", f, FmtPct(1)) Then n = n + 1
    f = ""
    f = f & "AVERAGE ( Fait_RH[Jours_Absence] )" & vbLf
    If AjouterMesure("Jours Absence Moyens", f, FmtDec(2)) Then n = n + 1
    f = ""
    f = f & "DIVIDE ( [CA Net HT], SUM ( Fait_RH[Heures_Travaillees] ) )" & vbLf
    If AjouterMesure("CA Net par Heure Travaillée", f, FmtEuro(2)) Then n = n + 1
    CreerMesures = n
End Function

'------------------------------------------------------------------------------
' TCD de contrôle branché sur le modèle
'------------------------------------------------------------------------------
Private Sub CreerTCDControle()
    Dim ws As Worksheet, pc As PivotCache, pt As PivotTable, v As Variant
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = "TCD_CONTROLE"
    ActiveWindow.DisplayGridlines = False
    ws.Range("B1").Value = "TCD DE CONTRÔLE - branché sur le modèle Power Pivot"
    ws.Range("B1").Font.Bold = True
    ws.Range("B1").Font.Size = 14
    ws.Range("B2").Value = "Compare le TCD (ligne 4) aux valeurs attendues (ligne 10), calculées hors Excel sur les mêmes fichiers. " & _
                           "Un écart = une étape Power Query ou une mesure à revoir."
    ws.Range("B2").Font.Italic = True
    EcrireValeursAttendues ws
    On Error Resume Next
    Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlExternal, _
                SourceData:=ThisWorkbook.Connections("ThisWorkbookDataModel"), Version:=6)
    If pc Is Nothing Then
        m_sEchecs = m_sEchecs & vbLf & "TCD de contrôle : connexion ThisWorkbookDataModel introuvable"
        Exit Sub
    End If
    Set pt = pc.CreatePivotTable(TableDestination:=ws.Range("B4"), TableName:="tcdControle")
    If pt Is Nothing Then
        m_sEchecs = m_sEchecs & vbLf & "TCD de contrôle : création impossible (" & Err.Description & ")"
        Exit Sub
    End If
    pt.CubeFields("[Dim_Dates].[Exercice]").Orientation = xlRowField
    For Each v In Array("CA Net HT", "Taux de Marge", "Tickets", "Ticket Moyen", ChrW(916) & " CA Net %", ChrW(916) & " CA Comparable %", "Atteinte Budget %", "Nb Magasins Actifs")
        Err.Clear
        pt.AddDataField pt.CubeFields("[Measures].[" & CStr(v) & "]")
        If Err.Number <> 0 Then m_sEchecs = m_sEchecs & vbLf & "TCD : mesure " & CStr(v) & " non ajoutée"
    Next v
    pt.RowAxisLayout xlTabularRow
    pt.ColumnGrand = False
    pt.RowGrand = False
    ws.Columns("B:J").AutoFit
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub EcrireValeursAttendues(ByVal ws As Worksheet)
    With ws
        .Range("B10").Value = "Valeurs attendues"
        .Range("B10").Font.Bold = True
        .Range("B11:J11").Value = Array("Exercice", "CA Net HT", "Taux de Marge", "Tickets", "Ticket Moyen", _
                                        ChrW(916) & " CA Net %", ChrW(916) & " CA Comparable %", "Atteinte Budget %", "Nb Magasins Actifs")
        .Range("B12:J12").Value = Array("2024/25", 141471440.65, 0.69857, 11179848, 12.654, "(vide)", "(vide)", 0.92664, 118)
        .Range("B13:J13").Value = Array("2025/26", 159556531.46, 0.69554, 11899174, 13.409, 0.12784, 0.10221, 0.98057, 120)
        .Range("B11:J11").Font.Bold = True
        .Range("B11:J11").Interior.Color = RGB(242, 241, 237)
        .Range("C12:C13").NumberFormat = "#,##0 ""€"""
        .Range("D12:D13").NumberFormat = "0.00%"
        .Range("E12:E13").NumberFormat = "#,##0"
        .Range("F12:F13").NumberFormat = "0.00 ""€"""
        .Range("G12:I13").NumberFormat = "0.0%"
        .Range("J12:J13").NumberFormat = "0"
        .Range("G12:H12").HorizontalAlignment = xlRight
        .Range("B15").Value = "Notes de lecture"
        .Range("B15").Font.Bold = True
        .Range("B16").Value = "- 2024/25 n'a pas de N-1 dans les fichiers : les deux évolutions sont vides, c'est normal."
        .Range("B17").Value = "- L'atteinte budget ne compare que les mois budgétés : janv.-août 2025 pour 2024/25 (le budget commence en janvier 2025), " & _
                              "les douze mois pour 2025/26. Sur janv.-août 2026 seul (filtre Annee = 2026), elle vaut 97,2 %, comme dans le classeur."
        .Range("B18").Value = "- Lignes de Fait_Ventes : 288 844 (289 506 lues, 240 doublons et 422 codes magasin vides écartés)."
        .Range("B16:B18").Font.Color = RGB(107, 117, 112)
    End With
End Sub

'------------------------------------------------------------------------------
' Liste et code des requêtes
'------------------------------------------------------------------------------
Private Function NomsRequetes(Optional ByVal bOrdreInverse As Boolean = False) As Variant
    Dim a As Variant, b() As Variant, i As Long, n As Long
    a = Array("pChemin", "fxDateSouple", "fxNombreSouple", "fxNettoyerTexte", "fxEcartJours", "fxBudgetAnnee", "qry_Ventes_Brutes", "Fait_Ventes", "qry_Rejets_Ventes", "Dim_Magasins", "Dim_Produits", "Dim_Fournisseurs", "Dim_Employes", "Fait_Budget", "Fait_Mix_Produits", "Fait_Stocks", "Fait_Achats", "Fait_RH", "Dim_Dates")
    If Not bOrdreInverse Then
        NomsRequetes = a
        Exit Function
    End If
    n = UBound(a)
    ReDim b(0 To n)
    For i = 0 To n
        b(i) = a(n - i)
    Next i
    NomsRequetes = b
End Function

Private Sub CreerRequetes()
    AjouterRequete "pChemin", M_pChemin(), "Chemin du dossier Donnees, lu dans la cellule nommée Chemin_Donnees"
    AjouterRequete "fxDateSouple", M_fxDateSouple(), "Convertit JJ/MM/AAAA, AAAA-MM-JJ ou une vraie date en date"
    AjouterRequete "fxNombreSouple", M_fxNombreSouple(), "Convertit un nombre écrit avec virgule ou point décimal"
    AjouterRequete "fxNettoyerTexte", M_fxNettoyerTexte(), "Casse de titre, particules en minuscule après un tiret (Aix-en-Provence)"
    AjouterRequete "fxEcartJours", M_fxEcartJours(), "Retard de livraison en jours au-delà du délai contractuel (null si non livré)"
    AjouterRequete "fxBudgetAnnee", M_fxBudgetAnnee(), "Lit un fichier Budget_AAAA.xlsx et le dépivote"
    AjouterRequete "qry_Ventes_Brutes", M_qry_Ventes_Brutes(), "Les 24 fichiers de ventes lus et convertis, sans filtre"
    AjouterRequete "Fait_Ventes", M_Fait_Ventes(), "Ventes nettoyées : canal normalisé, doublons et codes vides retirés"
    AjouterRequete "qry_Rejets_Ventes", M_qry_Rejets_Ventes(), "Lignes sans code magasin, comptées par fichier"
    AjouterRequete "Dim_Magasins", M_Dim_Magasins(), "Référentiel magasins dédoublonné, villes normalisées"
    AjouterRequete "Dim_Produits", M_Dim_Produits(), "Référentiel produits"
    AjouterRequete "Dim_Fournisseurs", M_Dim_Fournisseurs(), "Référentiel fournisseurs"
    AjouterRequete "Dim_Employes", M_Dim_Employes(), "Référentiel salariés, taux horaires texte convertis"
    AjouterRequete "Fait_Budget", M_Fait_Budget(), "Budgets 2025 et 2026 dépivotés : 120 magasins x 12 mois x 2 ans"
    AjouterRequete "Fait_Mix_Produits", M_Fait_Mix_Produits(), "Mix produits : magasin x mois x produit"
    AjouterRequete "Fait_Stocks", M_Fait_Stocks(), "Stocks mensuels par magasin et produit"
    AjouterRequete "Fait_Achats", M_Fait_Achats(), "Lignes de commande fournisseurs, avec retard de livraison"
    AjouterRequete "Fait_RH", M_Fait_RH(), "Heures mensuelles par salarié"
    AjouterRequete "Dim_Dates", M_Dim_Dates(), "Calendrier du 01/09/2024 au 31/08/2026, exercice septembre-août"
End Sub

Private Function M_pChemin() As String
    Dim s As String
    s = s & "let" & vbLf
    s = s & "    Source = Excel.CurrentWorkbook(){[Name = ""Chemin_Donnees""]}[Content]," & vbLf
    s = s & "    Valeur = Text.Trim(Text.From(Source{0}[Column1]))," & vbLf
    s = s & "    AvecSeparateur = if Text.EndsWith(Valeur, ""\"") then Valeur else Valeur & ""\""" & vbLf
    s = s & "in" & vbLf
    s = s & "    AvecSeparateur" & vbLf
    M_pChemin = s
End Function

Private Function M_fxDateSouple() As String
    Dim s As String
    s = s & "(t as any) as nullable date =>" & vbLf
    s = s & "let" & vbLf
    s = s & "    Resultat =" & vbLf
    s = s & "        if t = null then null" & vbLf
    s = s & "        else if t is date then t" & vbLf
    s = s & "        else if t is datetime then Date.From(t)" & vbLf
    s = s & "        else" & vbLf
    s = s & "            let" & vbLf
    s = s & "                s = Text.Trim(Text.From(t))" & vbLf
    s = s & "            in" & vbLf
    s = s & "                if s = """" then null" & vbLf
    s = s & "                else if Text.Contains(s, ""/"") then" & vbLf
    s = s & "                    let p = Text.Split(s, ""/"") in" & vbLf
    s = s & "                    #date(Number.FromText(p{2}), Number.FromText(p{1}), Number.FromText(p{0}))" & vbLf
    s = s & "                else" & vbLf
    s = s & "                    #date(Number.FromText(Text.Middle(s, 0, 4)), Number.FromText(Text.Middle(s, 5, 2)), Number.FromText(Text.Middle(s, 8, 2)))" & vbLf
    s = s & "in" & vbLf
    s = s & "    Resultat" & vbLf
    M_fxDateSouple = s
End Function

Private Function M_fxNombreSouple() As String
    Dim s As String
    s = s & "(t as any) as nullable number =>" & vbLf
    s = s & "let" & vbLf
    s = s & "    Resultat =" & vbLf
    s = s & "        if t = null then null" & vbLf
    s = s & "        else if t is number then t" & vbLf
    s = s & "        else" & vbLf
    s = s & "            let" & vbLf
    s = s & "                s = Text.Trim(Text.From(t))" & vbLf
    s = s & "            in" & vbLf
    s = s & "                if s = """" then null" & vbLf
    s = s & "                else if Text.Contains(s, "","") then Number.FromText(s, ""fr-FR"")" & vbLf
    s = s & "                else Number.FromText(s, ""en-US"")" & vbLf
    s = s & "in" & vbLf
    s = s & "    Resultat" & vbLf
    M_fxNombreSouple = s
End Function

Private Function M_fxNettoyerTexte() As String
    Dim s As String
    s = s & "(t as nullable text) as nullable text =>" & vbLf
    s = s & "let" & vbLf
    s = s & "    Particules = {""de"", ""du"", ""des"", ""la"", ""le"", ""les"", ""sur"", ""en""}," & vbLf
    s = s & "    Majuscule = (w as text) as text =>" & vbLf
    s = s & "        if w = """" then w else Text.Upper(Text.Start(w, 1)) & Text.Lower(Text.Range(w, 1))," & vbLf
    s = s & "    TraiterMot = (m as text) as text =>" & vbLf
    s = s & "        let" & vbLf
    s = s & "            Morceaux = Text.Split(m, ""-"")," & vbLf
    s = s & "            Traites = List.Transform(" & vbLf
    s = s & "                List.Positions(Morceaux)," & vbLf
    s = s & "                (i) => if i > 0 and List.Contains(Particules, Text.Lower(Morceaux{i}))" & vbLf
    s = s & "                       then Text.Lower(Morceaux{i})" & vbLf
    s = s & "                       else Majuscule(Morceaux{i}))" & vbLf
    s = s & "        in" & vbLf
    s = s & "            Text.Combine(Traites, ""-"")," & vbLf
    s = s & "    Resultat =" & vbLf
    s = s & "        if t = null then null" & vbLf
    s = s & "        else Text.Combine(List.Transform(List.Select(Text.Split(Text.Trim(t), "" ""), each _ <> """"), TraiterMot), "" "")" & vbLf
    s = s & "in" & vbLf
    s = s & "    Resultat" & vbLf
    M_fxNettoyerTexte = s
End Function

Private Function M_fxEcartJours() As String
    Dim s As String
    s = s & "(dCommande as nullable date, dLivraison as nullable date, delai as nullable number) as nullable number =>" & vbLf
    s = s & "    if dCommande = null or dLivraison = null or delai = null then null" & vbLf
    s = s & "    else Duration.Days(dLivraison - dCommande) - delai" & vbLf
    M_fxEcartJours = s
End Function

Private Function M_fxBudgetAnnee() As String
    Dim s As String
    s = s & "(annee as number) as table =>" & vbLf
    s = s & "let" & vbLf
    s = s & "    Classeur = Excel.Workbook(File.Contents(pChemin & ""07_Budget\Budget_"" & Number.ToText(annee) & "".xlsx""), null, true)," & vbLf
    s = s & "    Feuille = Classeur{0}[Data]," & vbLf
    s = s & "    SansTitres = Table.Skip(Feuille, 4)," & vbLf
    s = s & "    EnTetes = Table.PromoteHeaders(SansTitres, [PromoteAllScalars = true])," & vbLf
    s = s & "    SansVides = Table.SelectRows(EnTetes, each [Code_Magasin] <> null and [Code_Magasin] <> """")," & vbLf
    s = s & "    Depivote = Table.UnpivotOtherColumns(SansVides, {""Code_Magasin"", ""Nom_Magasin"", ""Region""}, ""Mois_Texte"", ""Budget_CA_HT"")," & vbLf
    s = s & "    Mois = Table.AddColumn(Depivote, ""Mois"", each" & vbLf
    s = s & "        #date(Number.FromText(Text.AfterDelimiter([Mois_Texte], ""/"")), Number.FromText(Text.BeforeDelimiter([Mois_Texte], ""/"")), 1), type date)," & vbLf
    s = s & "    Colonnes = Table.SelectColumns(Mois, {""Code_Magasin"", ""Mois"", ""Budget_CA_HT""})," & vbLf
    s = s & "    Types = Table.TransformColumnTypes(Colonnes, {{""Code_Magasin"", type text}, {""Budget_CA_HT"", type number}})" & vbLf
    s = s & "in" & vbLf
    s = s & "    Types" & vbLf
    M_fxBudgetAnnee = s
End Function

Private Function M_qry_Ventes_Brutes() As String
    Dim s As String
    s = s & "let" & vbLf
    s = s & "    Source = Folder.Files(pChemin & ""02_Ventes"")," & vbLf
    s = s & "    Fichiers = Table.SelectRows(Source, each Text.Lower([Extension]) = "".csv"" and Text.StartsWith([Name], ""Ventes_""))," & vbLf
    s = s & "    AvecMois = Table.AddColumn(Fichiers, ""Mois_Fichier"", each Text.BetweenDelimiters([Name], ""Ventes_"", "".csv""), type text)," & vbLf
    s = s & "    Lues = Table.AddColumn(AvecMois, ""Contenu"", each" & vbLf
    s = s & "        Table.TransformColumnNames(" & vbLf
    s = s & "            Table.PromoteHeaders(" & vbLf
    s = s & "                Csv.Document([Content], [Delimiter = "";"", Encoding = 65001, QuoteStyle = QuoteStyle.Csv])," & vbLf
    s = s & "                [PromoteAllScalars = true])," & vbLf
    s = s & "            (n) => Text.Replace(n, Character.FromNumber(65279), """")))," & vbLf
    s = s & "    Reduite = Table.SelectColumns(Lues, {""Mois_Fichier"", ""Contenu""})," & vbLf
    s = s & "    Etendue = Table.ExpandTableColumn(Reduite, ""Contenu""," & vbLf
    s = s & "        {""Date"", ""Code_Magasin"", ""Canal"", ""Nb_Tickets"", ""Nb_Articles"", ""CA_Brut_HT"", ""Remises_HT"", ""Cout_Matiere_HT""})," & vbLf
    s = s & "    Dates = Table.TransformColumns(Etendue, {{""Date"", fxDateSouple, type date}})," & vbLf
    s = s & "    Nombres = Table.TransformColumns(Dates, {" & vbLf
    s = s & "        {""CA_Brut_HT"", fxNombreSouple, type number}," & vbLf
    s = s & "        {""Remises_HT"", fxNombreSouple, type number}," & vbLf
    s = s & "        {""Cout_Matiere_HT"", fxNombreSouple, type number}," & vbLf
    s = s & "        {""Nb_Tickets"", fxNombreSouple, Int64.Type}," & vbLf
    s = s & "        {""Nb_Articles"", fxNombreSouple, Int64.Type}})" & vbLf
    s = s & "in" & vbLf
    s = s & "    Nombres" & vbLf
    M_qry_Ventes_Brutes = s
End Function

Private Function M_Fait_Ventes() As String
    Dim s As String
    s = s & "let" & vbLf
    s = s & "    Source = qry_Ventes_Brutes," & vbLf
    s = s & "    CanalPropre = Table.TransformColumns(Source, {{""Canal"", (c) =>" & vbLf
    s = s & "        let k = Text.Lower(Text.Trim(c)) in" & vbLf
    s = s & "        if k = ""sur place"" then ""Sur place""" & vbLf
    s = s & "        else if k = ""à emporter"" then ""À emporter""" & vbLf
    s = s & "        else if k = ""livraison"" then ""Livraison""" & vbLf
    s = s & "        else if k = ""drive"" then ""Drive""" & vbLf
    s = s & "        else Text.Trim(c), type text}})," & vbLf
    s = s & "    SansDoublons = Table.Distinct(CanalPropre," & vbLf
    s = s & "        {""Date"", ""Code_Magasin"", ""Canal"", ""Nb_Tickets"", ""Nb_Articles"", ""CA_Brut_HT"", ""Remises_HT"", ""Cout_Matiere_HT""})," & vbLf
    s = s & "    SansVides = Table.SelectRows(SansDoublons, each [Code_Magasin] <> null and Text.Trim([Code_Magasin]) <> """")," & vbLf
    s = s & "    CANet = Table.AddColumn(SansVides, ""CA_Net_HT"", each [CA_Brut_HT] - [Remises_HT], type number)," & vbLf
    s = s & "    Types = Table.TransformColumnTypes(CANet, {{""Code_Magasin"", type text}, {""Mois_Fichier"", type text}})" & vbLf
    s = s & "in" & vbLf
    s = s & "    Types" & vbLf
    M_Fait_Ventes = s
End Function

Private Function M_qry_Rejets_Ventes() As String
    Dim s As String
    s = s & "let" & vbLf
    s = s & "    Source = qry_Ventes_Brutes," & vbLf
    s = s & "    Vides = Table.SelectRows(Source, each [Code_Magasin] = null or Text.Trim([Code_Magasin]) = """")," & vbLf
    s = s & "    ParFichier = Table.Group(Vides, {""Mois_Fichier""}, {{""Nb_Rejets"", each Table.RowCount(_), Int64.Type}})" & vbLf
    s = s & "in" & vbLf
    s = s & "    ParFichier" & vbLf
    M_qry_Rejets_Ventes = s
End Function

Private Function M_Dim_Magasins() As String
    Dim s As String
    s = s & "let" & vbLf
    s = s & "    Classeur = Excel.Workbook(File.Contents(pChemin & ""01_Referentiels\Referentiels_Horizon_Contrat.xlsx""), null, true)," & vbLf
    s = s & "    Feuille = Classeur{[Item = ""Magasins"", Kind = ""Sheet""]}[Data]," & vbLf
    s = s & "    EnTetes = Table.PromoteHeaders(Feuille, [PromoteAllScalars = true])," & vbLf
    s = s & "    SansDoublons = Table.Distinct(EnTetes)," & vbLf
    s = s & "    Villes = Table.TransformColumns(SansDoublons, {{""Ville"", fxNettoyerTexte, type text}})," & vbLf
    s = s & "    Dates = Table.TransformColumns(Villes, {{""Date_Ouverture"", fxDateSouple, type date}, {""Date_Fermeture"", fxDateSouple, type date}})," & vbLf
    s = s & "    Types = Table.TransformColumnTypes(Dates, {" & vbLf
    s = s & "        {""Code_Magasin"", type text}, {""Nom_Magasin"", type text}, {""Region"", type text}," & vbLf
    s = s & "        {""Type_Emplacement"", type text}, {""Statut"", type text}, {""Surface_m2"", Int64.Type}," & vbLf
    s = s & "        {""Nb_Places_Assises"", Int64.Type}, {""Drive"", type text}, {""Franchise_Responsable"", type text}})" & vbLf
    s = s & "in" & vbLf
    s = s & "    Types" & vbLf
    M_Dim_Magasins = s
End Function

Private Function M_Dim_Produits() As String
    Dim s As String
    s = s & "let" & vbLf
    s = s & "    Classeur = Excel.Workbook(File.Contents(pChemin & ""01_Referentiels\Referentiels_Horizon_Contrat.xlsx""), null, true)," & vbLf
    s = s & "    Feuille = Classeur{[Item = ""Produits"", Kind = ""Sheet""]}[Data]," & vbLf
    s = s & "    EnTetes = Table.PromoteHeaders(Feuille, [PromoteAllScalars = true])," & vbLf
    s = s & "    Types = Table.TransformColumnTypes(EnTetes, {" & vbLf
    s = s & "        {""Code_Produit"", type text}, {""Libelle_Produit"", type text}, {""Categorie"", type text}," & vbLf
    s = s & "        {""Sous_Categorie"", type text}, {""Taux_TVA"", type number}, {""Taux_Cout_Matiere"", type number}," & vbLf
    s = s & "        {""Prix_Vente_TTC_Actuel"", type number}, {""Actif"", type text}})" & vbLf
    s = s & "in" & vbLf
    s = s & "    Types" & vbLf
    M_Dim_Produits = s
End Function

Private Function M_Dim_Fournisseurs() As String
    Dim s As String
    s = s & "let" & vbLf
    s = s & "    Classeur = Excel.Workbook(File.Contents(pChemin & ""01_Referentiels\Referentiels_Horizon_Contrat.xlsx""), null, true)," & vbLf
    s = s & "    Feuille = Classeur{[Item = ""Fournisseurs"", Kind = ""Sheet""]}[Data]," & vbLf
    s = s & "    EnTetes = Table.PromoteHeaders(Feuille, [PromoteAllScalars = true])," & vbLf
    s = s & "    Types = Table.TransformColumnTypes(EnTetes, {" & vbLf
    s = s & "        {""Code_Fournisseur"", type text}, {""Nom_Fournisseur"", type text}, {""Categorie_Approvisionnee"", type text}," & vbLf
    s = s & "        {""Delai_Livraison_Contractuel_Jours"", Int64.Type}, {""Conditions_Paiement"", type text}, {""Note_Qualite"", type number}})" & vbLf
    s = s & "in" & vbLf
    s = s & "    Types" & vbLf
    M_Dim_Fournisseurs = s
End Function

Private Function M_Dim_Employes() As String
    Dim s As String
    s = s & "let" & vbLf
    s = s & "    Classeur = Excel.Workbook(File.Contents(pChemin & ""01_Referentiels\Referentiels_Horizon_Contrat.xlsx""), null, true)," & vbLf
    s = s & "    Feuille = Classeur{[Item = ""Employes"", Kind = ""Sheet""]}[Data]," & vbLf
    s = s & "    EnTetes = Table.PromoteHeaders(Feuille, [PromoteAllScalars = true])," & vbLf
    s = s & "    Taux = Table.TransformColumns(EnTetes, {{""Taux_Horaire_Brut"", (v) =>" & vbLf
    s = s & "        if v is number then v" & vbLf
    s = s & "        else Number.FromText(Text.Remove(Text.From(v), {""€"", "" "", Character.FromNumber(160), Character.FromNumber(8239)}), ""fr-FR"")," & vbLf
    s = s & "        type number}})," & vbLf
    s = s & "    Dates = Table.TransformColumns(Taux, {{""Date_Embauche"", fxDateSouple, type date}})," & vbLf
    s = s & "    Types = Table.TransformColumnTypes(Dates, {" & vbLf
    s = s & "        {""Matricule"", type text}, {""Code_Magasin"", type text}, {""Nom_Prenom"", type text}," & vbLf
    s = s & "        {""Poste"", type text}, {""Type_Contrat"", type text}})" & vbLf
    s = s & "in" & vbLf
    s = s & "    Types" & vbLf
    M_Dim_Employes = s
End Function

Private Function M_Fait_Budget() As String
    Dim s As String
    s = s & "let" & vbLf
    s = s & "    Combine = Table.Combine({fxBudgetAnnee(2025), fxBudgetAnnee(2026)})" & vbLf
    s = s & "in" & vbLf
    s = s & "    Combine" & vbLf
    M_Fait_Budget = s
End Function

Private Function M_Fait_Mix_Produits() As String
    Dim s As String
    s = s & "let" & vbLf
    s = s & "    Source = Csv.Document(File.Contents(pChemin & ""03_Mix_Produits\Mix_Produits_Mensuel.csv""), [Delimiter = "";"", Encoding = 65001, QuoteStyle = QuoteStyle.Csv])," & vbLf
    s = s & "    EnTetes = Table.PromoteHeaders(Source, [PromoteAllScalars = true])," & vbLf
    s = s & "    Noms = Table.TransformColumnNames(EnTetes, (n) => Text.Replace(n, Character.FromNumber(65279), """"))," & vbLf
    s = s & "    Mois = Table.TransformColumns(Noms, {{""Mois"", (m) => fxDateSouple(m & ""-01""), type date}})," & vbLf
    s = s & "    Nombres = Table.TransformColumns(Mois, {" & vbLf
    s = s & "        {""Quantite_Vendue"", fxNombreSouple, Int64.Type}," & vbLf
    s = s & "        {""CA_HT"", fxNombreSouple, type number}," & vbLf
    s = s & "        {""Cout_Matiere_HT"", fxNombreSouple, type number}})," & vbLf
    s = s & "    Types = Table.TransformColumnTypes(Nombres, {{""Code_Magasin"", type text}, {""Code_Produit"", type text}})" & vbLf
    s = s & "in" & vbLf
    s = s & "    Types" & vbLf
    M_Fait_Mix_Produits = s
End Function

Private Function M_Fait_Stocks() As String
    Dim s As String
    s = s & "let" & vbLf
    s = s & "    Source = Csv.Document(File.Contents(pChemin & ""04_Stocks\Stocks_Mensuels.csv""), [Delimiter = "";"", Encoding = 65001, QuoteStyle = QuoteStyle.Csv])," & vbLf
    s = s & "    EnTetes = Table.PromoteHeaders(Source, [PromoteAllScalars = true])," & vbLf
    s = s & "    Noms = Table.TransformColumnNames(EnTetes, (n) => Text.Replace(n, Character.FromNumber(65279), """"))," & vbLf
    s = s & "    Mois = Table.TransformColumns(Noms, {{""Mois"", (m) => fxDateSouple(m & ""-01""), type date}})," & vbLf
    s = s & "    Nombres = Table.TransformColumns(Mois, {" & vbLf
    s = s & "        {""Stock_Initial"", fxNombreSouple, Int64.Type}, {""Entrees"", fxNombreSouple, Int64.Type}," & vbLf
    s = s & "        {""Sorties"", fxNombreSouple, Int64.Type}, {""Pertes"", fxNombreSouple, Int64.Type}," & vbLf
    s = s & "        {""Stock_Final"", fxNombreSouple, Int64.Type}})," & vbLf
    s = s & "    Taux = Table.AddColumn(Nombres, ""Taux_Perte"", each" & vbLf
    s = s & "        if [Sorties] = null or [Sorties] = 0 then null else [Pertes] / [Sorties], type number)," & vbLf
    s = s & "    Types = Table.TransformColumnTypes(Taux, {{""Code_Magasin"", type text}, {""Code_Produit"", type text}})" & vbLf
    s = s & "in" & vbLf
    s = s & "    Types" & vbLf
    M_Fait_Stocks = s
End Function

Private Function M_Fait_Achats() As String
    Dim s As String
    s = s & "let" & vbLf
    s = s & "    Source = Csv.Document(File.Contents(pChemin & ""05_Achats\Achats_Fournisseurs.csv""), [Delimiter = "";"", Encoding = 65001, QuoteStyle = QuoteStyle.Csv])," & vbLf
    s = s & "    EnTetes = Table.PromoteHeaders(Source, [PromoteAllScalars = true])," & vbLf
    s = s & "    Noms = Table.TransformColumnNames(EnTetes, (n) => Text.Replace(n, Character.FromNumber(65279), """"))," & vbLf
    s = s & "    Dates = Table.TransformColumns(Noms, {" & vbLf
    s = s & "        {""Date_Commande"", fxDateSouple, type date}," & vbLf
    s = s & "        {""Date_Livraison_Reelle"", fxDateSouple, type date}})," & vbLf
    s = s & "    Nombres = Table.TransformColumns(Dates, {" & vbLf
    s = s & "        {""Quantite_Commandee"", fxNombreSouple, Int64.Type}," & vbLf
    s = s & "        {""Prix_Achat_Unitaire_HT"", fxNombreSouple, type number}," & vbLf
    s = s & "        {""Montant_HT"", fxNombreSouple, type number}})," & vbLf
    s = s & "    Jointure = Table.NestedJoin(Nombres, {""Code_Fournisseur""}, Dim_Fournisseurs, {""Code_Fournisseur""}, ""F"", JoinKind.LeftOuter)," & vbLf
    s = s & "    Delai = Table.ExpandTableColumn(Jointure, ""F"", {""Delai_Livraison_Contractuel_Jours""}, {""Delai_Contractuel""})," & vbLf
    s = s & "    Retard = Table.AddColumn(Delai, ""Retard_Jours"", each" & vbLf
    s = s & "        fxEcartJours([Date_Commande], [Date_Livraison_Reelle], [Delai_Contractuel]), type number)," & vbLf
    s = s & "    MoisCommande = Table.AddColumn(Retard, ""Mois_Commande"", each Date.StartOfMonth([Date_Commande]), type date)," & vbLf
    s = s & "    Types = Table.TransformColumnTypes(MoisCommande, {" & vbLf
    s = s & "        {""Num_Commande"", type text}, {""Code_Magasin"", type text}," & vbLf
    s = s & "        {""Code_Fournisseur"", type text}, {""Code_Produit"", type text}})" & vbLf
    s = s & "in" & vbLf
    s = s & "    Types" & vbLf
    M_Fait_Achats = s
End Function

Private Function M_Fait_RH() As String
    Dim s As String
    s = s & "let" & vbLf
    s = s & "    Source = Csv.Document(File.Contents(pChemin & ""06_RH\Heures_Mensuelles.csv""), [Delimiter = "";"", Encoding = 65001, QuoteStyle = QuoteStyle.Csv])," & vbLf
    s = s & "    EnTetes = Table.PromoteHeaders(Source, [PromoteAllScalars = true])," & vbLf
    s = s & "    Noms = Table.TransformColumnNames(EnTetes, (n) => Text.Replace(n, Character.FromNumber(65279), """"))," & vbLf
    s = s & "    Mois = Table.TransformColumns(Noms, {{""Mois"", (m) => fxDateSouple(m & ""-01""), type date}})," & vbLf
    s = s & "    Nombres = Table.TransformColumns(Mois, {" & vbLf
    s = s & "        {""Heures_Travaillees"", fxNombreSouple, type number}," & vbLf
    s = s & "        {""Heures_Supplementaires"", fxNombreSouple, type number}," & vbLf
    s = s & "        {""Jours_Absence"", fxNombreSouple, type number}})," & vbLf
    s = s & "    Types = Table.TransformColumnTypes(Nombres, {" & vbLf
    s = s & "        {""Matricule"", type text}, {""Code_Magasin"", type text}, {""Poste"", type text}})" & vbLf
    s = s & "in" & vbLf
    s = s & "    Types" & vbLf
    M_Fait_RH = s
End Function

Private Function M_Dim_Dates() As String
    Dim s As String
    s = s & "let" & vbLf
    s = s & "    Debut = #date(2024, 9, 1)," & vbLf
    s = s & "    Fin = #date(2026, 8, 31)," & vbLf
    s = s & "    Jours = List.Dates(Debut, Duration.Days(Fin - Debut) + 1, #duration(1, 0, 0, 0))," & vbLf
    s = s & "    Base = Table.FromList(Jours, Splitter.SplitByNothing(), {""Date""})," & vbLf
    s = s & "    Typee = Table.TransformColumnTypes(Base, {{""Date"", type date}})," & vbLf
    s = s & "    C1 = Table.AddColumn(Typee, ""Annee"", each Date.Year([Date]), Int64.Type)," & vbLf
    s = s & "    C2 = Table.AddColumn(C1, ""Mois_Num"", each Date.Month([Date]), Int64.Type)," & vbLf
    s = s & "    C3 = Table.AddColumn(C2, ""Mois_Debut"", each Date.StartOfMonth([Date]), type date)," & vbLf
    s = s & "    C4 = Table.AddColumn(C3, ""Mois_Nom"", each Date.ToText([Date], ""MMMM"", ""fr-FR""), type text)," & vbLf
    s = s & "    C5 = Table.AddColumn(C4, ""Mois_Abrege"", each Date.ToText([Date], ""MMM yy"", ""fr-FR""), type text)," & vbLf
    s = s & "    C6 = Table.AddColumn(C5, ""Trimestre"", each ""T"" & Number.ToText(Date.QuarterOfYear([Date])), type text)," & vbLf
    s = s & "    C7 = Table.AddColumn(C6, ""Semaine"", each Date.WeekOfYear([Date], Day.Monday), Int64.Type)," & vbLf
    s = s & "    C8 = Table.AddColumn(C7, ""Jour_Semaine"", each Date.DayOfWeek([Date], Day.Monday) + 1, Int64.Type)," & vbLf
    s = s & "    C9 = Table.AddColumn(C8, ""Jour_Semaine_Nom"", each Date.ToText([Date], ""dddd"", ""fr-FR""), type text)," & vbLf
    s = s & "    C10 = Table.AddColumn(C9, ""Est_Weekend"", each [Jour_Semaine] >= 6, type logical)," & vbLf
    s = s & "    C11 = Table.AddColumn(C10, ""Annee_Exercice"", each if [Mois_Num] >= 9 then [Annee] else [Annee] - 1, Int64.Type)," & vbLf
    s = s & "    C12 = Table.AddColumn(C11, ""Exercice"", each" & vbLf
    s = s & "        Number.ToText([Annee_Exercice]) & ""/"" & Text.End(Number.ToText([Annee_Exercice] + 1), 2), type text)," & vbLf
    s = s & "    C13 = Table.AddColumn(C12, ""Mois_Exercice"", each if [Mois_Num] >= 9 then [Mois_Num] - 8 else [Mois_Num] + 4, Int64.Type)," & vbLf
    s = s & "    C14 = Table.AddColumn(C13, ""Mois_Index"", each [Annee] * 100 + [Mois_Num], Int64.Type)" & vbLf
    s = s & "in" & vbLf
    s = s & "    C14" & vbLf
    M_Dim_Dates = s
End Function
