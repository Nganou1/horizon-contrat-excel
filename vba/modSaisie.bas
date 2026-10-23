Attribute VB_Name = "modSaisie"
Option Explicit
'==============================================================================
' Horizon_Contrat - données du formulaire de saisie des relevés
'
' Le formulaire frmSaisieReleve ne gère que l'interface ; toutes les lectures
' et écritures dans tblSaisieReleves passent par ce module.
'
' tblSaisieReleves : Horodatage | Code_Magasin | Date_Releve | Canal |
'   CA_Brut_TTC | Remises_TTC | CA_Net_HT | Cout_Matiere_HT | Nb_Tickets |
'   Nb_Articles | Alertes | Saisi_Par
'==============================================================================

Public Sub OuvrirSaisieReleve()
    Dim sManque As String
    sManque = VerifierPrerequisSaisie()
    If sManque <> "" Then
        MsgBox "Saisie impossible :" & sManque & vbLf & vbLf & "Lance InstallerReporting (Alt+F8).", _
               vbExclamation, "Horizon_Contrat"
        Exit Sub
    End If
    ' Appel par le nom : ce module compile même si le formulaire n'existe pas encore.
    On Error GoTo Echec
    VBA.UserForms.Add("frmSaisieReleve").Show
    Exit Sub
Echec:
    MsgBox "Ouverture du formulaire impossible : " & Err.Description & " (erreur " & Err.Number & ")." & vbLf & vbLf & _
           "Si frmSaisieReleve n'existe pas encore : LISEZ-MOI, étape 6.", vbExclamation, "Horizon_Contrat"
End Sub

Public Function VerifierPrerequisSaisie() As String
    Dim s As String, v As Variant
    If Not FeuilleExiste(ONGLET_SAISIE) Then
        s = s & vbLf & "- onglet " & ONGLET_SAISIE & " absent"
    ElseIf TableauAbsent(ONGLET_SAISIE, TBL_SAISIE) Then
        s = s & vbLf & "- tableau " & TBL_SAISIE & " absent"
    End If
    For Each v In Array("Mois_Ouvert", "Taux_TVA_Restauration", "Ratio_Articles_Min", "Ratio_Articles_Max", _
                        "Plafond_Remise", "Seuil_Ecart_Ticket_Moyen", "Tolerance_Cout_Matiere_Pt")
        If PlageNommee(CStr(v)) Is Nothing Then s = s & vbLf & "- paramètre " & CStr(v) & " absent"
    Next v
    VerifierPrerequisSaisie = s
End Function

Private Function TableauAbsent(ByVal sFeuille As String, ByVal sTable As String) As Boolean
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ThisWorkbook.Worksheets(sFeuille).ListObjects(sTable)
    TableauAbsent = lo Is Nothing
End Function

Private Function TableSaisie() As ListObject
    Set TableSaisie = Tableau(ONGLET_SAISIE, TBL_SAISIE)
End Function

Private Function Colonne(ByVal sEntete As String) As Range
    Set Colonne = TableSaisie().ListColumns(sEntete).DataBodyRange
End Function

'------------------------------------------------------------------------------
' Magasins
'------------------------------------------------------------------------------
Public Function NomCanal(ByVal i As Long) As String
    NomCanal = Choose(i, "Sur place", "À emporter", "Livraison", "Drive")
End Function

' Magasins ouverts seulement : un champ qui ne propose pas le choix impossible
' vaut mieux qu'un message d'erreur après coup (HC-063 n'apparaît pas).
' Renvoie un tableau (0..n-1, 0..1) : code, "Ville - type d'emplacement".
Public Function ListeMagasinsOuverts() As Variant
    Dim lo As ListObject, v As Variant, r() As Variant, i As Long, n As Long
    Dim iCode As Long, iVille As Long, iType As Long, iFerm As Long

    Set lo = Tableau(ONGLET_REF_MAGASINS, TBL_MAGASINS)
    iCode = IndexColonne(lo, "Code_Magasin")
    iVille = IndexColonne(lo, "Ville")
    iType = IndexColonne(lo, "Type_Emplacement")
    iFerm = IndexColonne(lo, "Date_Fermeture")
    v = lo.DataBodyRange.Value

    ReDim r(0 To UBound(v, 1) - 1, 0 To 1)
    For i = 1 To UBound(v, 1)
        If Len(Trim$(CStr(v(i, iFerm)))) = 0 Then
            r(n, 0) = v(i, iCode)
            r(n, 1) = v(i, iVille) & " - " & v(i, iType)
            n = n + 1
        End If
    Next i
    ListeMagasinsOuverts = TronquerLignes(r, n, 2)
End Function

Public Function MagasinADrive(ByVal sCode As String) As Boolean
    Dim v As Variant
    v = ValeurReferentiel(sCode, "Drive")
    MagasinADrive = (StrComp(Trim$(CStr(v)), "Oui", vbTextCompare) = 0)
End Function

Public Function NbCanauxMagasin(ByVal sCode As String) As Long
    NbCanauxMagasin = IIf(MagasinADrive(sCode), 4, 3)
End Function

Private Function ValeurReferentiel(ByVal sCode As String, ByVal sColonne As String) As Variant
    Dim lo As ListObject, k As Variant
    Set lo = Tableau(ONGLET_REF_MAGASINS, TBL_MAGASINS)
    k = Application.Match(sCode, lo.ListColumns("Code_Magasin").DataBodyRange, 0)
    If IsError(k) Then
        ValeurReferentiel = Empty
    Else
        ValeurReferentiel = lo.ListColumns(sColonne).DataBodyRange.Cells(CLng(k), 1).Value
    End If
End Function

' Taux de coût matière de l'exercice du magasin (1 - taux de marge), réseau en secours.
Public Function TauxCoutMagasin(ByVal sCode As String) As Double
    Dim k As Variant, v As Variant
    k = Application.Match(sCode, Plage("CM_Code_Magasin"), 0)
    If Not IsError(k) Then v = Plage("CM_Taux_Marge_N").Cells(CLng(k), 1).Value
    If IsNumeric(v) And Not IsEmpty(v) Then
        TauxCoutMagasin = 1 - CDbl(v)
    Else
        TauxCoutMagasin = 1 - CDbl(Plage("D_Taux_Reseau").Value)
    End If
End Function

'------------------------------------------------------------------------------
' Relevés
'------------------------------------------------------------------------------
Public Function ReleveExiste(ByVal sCode As String, ByVal d As Date, ByVal sCanal As String) As Boolean
    If NbLignesUtiles(TableSaisie()) = 0 Then Exit Function
    ReleveExiste = Application.WorksheetFunction.CountIfs( _
        Colonne("Code_Magasin"), sCode, Colonne("Date_Releve"), CDbl(d), Colonne("Canal"), sCanal) > 0
End Function

Public Function NbRelevesMois(ByVal sCode As String, ByVal dMois As Date) As Long
    If NbLignesUtiles(TableSaisie()) = 0 Then Exit Function
    NbRelevesMois = Application.WorksheetFunction.CountIfs(Colonne("Code_Magasin"), sCode, _
        Colonne("Date_Releve"), ">=" & CLng(DebutDeMois(dMois)), _
        Colonne("Date_Releve"), "<=" & CLng(FinDeMois(dMois)))
End Function

' Ticket moyen de référence : 30 jours de relevés avant la date (ratio de
' sommes), ou à défaut (moins de 5 relevés) le ticket moyen de l'exercice.
Public Function TicketMoyenReference(ByVal sCode As String, ByVal d As Date, ByRef sSource As String) As Double
    Dim dNet As Double, dTickets As Double, n As Long, k As Variant, v As Variant
    Dim sDebut As String, sFin As String

    If NbLignesUtiles(TableSaisie()) > 0 Then
        sDebut = ">=" & CLng(d - 30)
        sFin = "<" & CLng(d)
        With Application.WorksheetFunction
            n = .CountIfs(Colonne("Code_Magasin"), sCode, Colonne("Date_Releve"), sDebut, Colonne("Date_Releve"), sFin)
            If n >= 5 Then
                dNet = .SumIfs(Colonne("CA_Net_HT"), Colonne("Code_Magasin"), sCode, _
                               Colonne("Date_Releve"), sDebut, Colonne("Date_Releve"), sFin)
                dTickets = .SumIfs(Colonne("Nb_Tickets"), Colonne("Code_Magasin"), sCode, _
                                   Colonne("Date_Releve"), sDebut, Colonne("Date_Releve"), sFin)
            End If
        End With
    End If
    If dTickets > 0 Then
        sSource = "moyenne des 30 derniers jours"
        TicketMoyenReference = dNet / dTickets
        Exit Function
    End If

    k = Application.Match(sCode, Plage("CM_Code_Magasin"), 0)
    If Not IsError(k) Then v = Plage("CM_Ticket_Moyen_N").Cells(CLng(k), 1).Value
    If IsNumeric(v) And Not IsEmpty(v) Then
        sSource = "moyenne de l'exercice"
        TicketMoyenReference = CDbl(v)
    End If
End Function

Public Function CanalSuivantLibre(ByVal sCode As String, ByVal d As Date, ByVal iDepart As Long) As Long
    Dim i As Long, k As Long, nCanaux As Long
    nCanaux = NbCanauxMagasin(sCode)
    For k = 1 To nCanaux
        i = ((iDepart - 1 + k) Mod nCanaux) + 1
        If Not ReleveExiste(sCode, d, NomCanal(i)) Then
            CanalSuivantLibre = i
            Exit Function
        End If
    Next k
End Function

' Derniers relevés du magasin pour le ListBox : (0..n-1, 0..4) =
' Date | Canal | CA net HT | Tickets | Alertes ; Empty si aucun.
Public Function DerniersReleves(ByVal sCode As String, ByVal nMax As Long) As Variant
    Dim lo As ListObject, v As Variant, r() As Variant, idx() As Long
    Dim i As Long, j As Long, n As Long, t As Long, k As Long
    Dim iCode As Long, iDate As Long, iCanal As Long, iNet As Long, iTk As Long, iAl As Long, iHor As Long

    Set lo = TableSaisie()
    If NbLignesUtiles(lo) = 0 Then Exit Function
    iHor = IndexColonne(lo, "Horodatage")
    iCode = IndexColonne(lo, "Code_Magasin")
    iDate = IndexColonne(lo, "Date_Releve")
    iCanal = IndexColonne(lo, "Canal")
    iNet = IndexColonne(lo, "CA_Net_HT")
    iTk = IndexColonne(lo, "Nb_Tickets")
    iAl = IndexColonne(lo, "Alertes")
    v = lo.DataBodyRange.Value

    ReDim idx(1 To UBound(v, 1))
    For i = 1 To UBound(v, 1)
        If StrComp(CStr(v(i, iCode)), sCode, vbTextCompare) = 0 Then
            n = n + 1
            idx(n) = i
        End If
    Next i
    If n = 0 Then Exit Function

    ' Tri partiel (date puis horodatage décroissants) : seuls les nMax premiers comptent.
    For j = 1 To IIf(n < nMax, n, nMax)
        For k = j + 1 To n
            If PlusRecent(v, idx(k), idx(j), iDate, iHor) Then
                t = idx(j): idx(j) = idx(k): idx(k) = t
            End If
        Next k
    Next j

    If n > nMax Then n = nMax
    ReDim r(0 To n - 1, 0 To 4)
    For j = 1 To n
        i = idx(j)
        r(j - 1, 0) = Format$(v(i, iDate), "dd/mm/yyyy")
        r(j - 1, 1) = v(i, iCanal)
        r(j - 1, 2) = Format$(v(i, iNet), "#,##0.00") & " €"
        r(j - 1, 3) = v(i, iTk)
        r(j - 1, 4) = v(i, iAl)
    Next j
    DerniersReleves = r
End Function

' v passé ByRef : ByVal copierait tout le tableau à chaque comparaison.
Private Function PlusRecent(ByRef v As Variant, ByVal a As Long, ByVal b As Long, _
                            ByVal iDate As Long, ByVal iHor As Long) As Boolean
    If v(a, iDate) <> v(b, iDate) Then
        PlusRecent = (v(a, iDate) > v(b, iDate))
    Else
        PlusRecent = (v(a, iHor) > v(b, iHor))
    End If
End Function

Public Sub EnregistrerReleve(ByVal sCode As String, ByVal d As Date, ByVal sCanal As String, _
                             ByVal dBrut As Double, ByVal dRemises As Double, ByVal dNet As Double, _
                             ByVal dCout As Double, ByVal lTickets As Long, ByVal lArticles As Long, _
                             ByVal sAlertes As String)
    Dim lo As ListObject, r As Range, bEvents As Boolean, lNum As Long, sDesc As String
    bEvents = Application.EnableEvents
    On Error GoTo Echec
    Set lo = TableSaisie()
    Application.EnableEvents = False
    Set r = NouvelleLigne(lo)
    r.Cells(1, IndexColonne(lo, "Horodatage")).Value = Now
    r.Cells(1, IndexColonne(lo, "Code_Magasin")).Value = sCode
    r.Cells(1, IndexColonne(lo, "Date_Releve")).Value = d
    r.Cells(1, IndexColonne(lo, "Canal")).Value = sCanal
    r.Cells(1, IndexColonne(lo, "CA_Brut_TTC")).Value = Round(dBrut, 2)
    r.Cells(1, IndexColonne(lo, "Remises_TTC")).Value = Round(dRemises, 2)
    r.Cells(1, IndexColonne(lo, "CA_Net_HT")).Value = Round(dNet, 2)
    r.Cells(1, IndexColonne(lo, "Cout_Matiere_HT")).Value = Round(dCout, 2)
    r.Cells(1, IndexColonne(lo, "Nb_Tickets")).Value = lTickets
    r.Cells(1, IndexColonne(lo, "Nb_Articles")).Value = lArticles
    r.Cells(1, IndexColonne(lo, "Alertes")).Value = sAlertes
    r.Cells(1, IndexColonne(lo, "Saisi_Par")).Value = Environ$("USERNAME")
    Application.EnableEvents = bEvents
    Exit Sub
Echec:
    lNum = Err.Number
    sDesc = Err.Description
    Application.EnableEvents = bEvents
    Err.Raise lNum, "EnregistrerReleve", sDesc
End Sub

'------------------------------------------------------------------------------
Private Function TronquerLignes(ByRef r() As Variant, ByVal n As Long, ByVal nCol As Long) As Variant
    Dim t() As Variant, i As Long, j As Long
    If n = 0 Then Exit Function
    ReDim t(0 To n - 1, 0 To nCol - 1)
    For i = 0 To n - 1
        For j = 0 To nCol - 1
            t(i, j) = r(i, j)
        Next j
    Next i
    TronquerLignes = t
End Function
