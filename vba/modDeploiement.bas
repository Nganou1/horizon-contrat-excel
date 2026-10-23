Attribute VB_Name = "modDeploiement"
Option Explicit
'==============================================================================
' Horizon_Contrat - protection et masquage avant de donner le classeur
'
' La protection Excel est un garde-fou contre les fausses manipulations, pas
' une sécurité : le mot de passe d'une feuille se retire en quelques secondes.
' Le dire en entretien est une qualité, prétendre le contraire une erreur.
'==============================================================================

Public Sub ProtegerRestitution()
    VerrouillerSaufSelection Feuille(ONGLET_DASHBOARD), "C4,H4"
    VerrouillerSaufSelection Feuille(ONGLET_FICHE), "C4,E4"
    MsgBox "DASHBOARD et FICHE_MAGASIN protégés." & vbLf & _
           "Restent modifiables : les filtres (C4, H4) et la sélection du magasin (C4, E4).", _
           vbInformation, "Horizon_Contrat"
End Sub

Public Sub DeprotegerRestitution()
    On Error Resume Next
    Feuille(ONGLET_DASHBOARD).Unprotect MOT_DE_PASSE_FEUILLES
    Feuille(ONGLET_FICHE).Unprotect MOT_DE_PASSE_FEUILLES
    On Error GoTo 0
    MsgBox "DASHBOARD et FICHE_MAGASIN déprotégés.", vbInformation, "Horizon_Contrat"
End Sub

Private Sub VerrouillerSaufSelection(ByVal ws As Worksheet, ByVal sCellesLibres As String)
    ws.Unprotect MOT_DE_PASSE_FEUILLES
    ws.Cells.Locked = True
    ws.Range(sCellesLibres).Locked = False
    ProtegerFeuille ws
End Sub

' Formats autorisés : la mise en surbrillance de FICHE_MAGASIN (SelectionChange)
' modifie la couleur de fond. Les boutons (formes) restent cliquables.
Public Sub ProtegerFeuille(ByVal ws As Worksheet)
    ws.Protect Password:=MOT_DE_PASSE_FEUILLES, DrawingObjects:=True, Contents:=True, _
               Scenarios:=True, AllowFormattingCells:=True
    ws.EnableSelection = xlNoRestrictions
End Sub

'------------------------------------------------------------------------------
' Feuilles techniques en « très masqué » : invisibles même par clic droit >
' Afficher. Les formules et les macros continuent de les lire normalement.
'------------------------------------------------------------------------------
Public Sub MasquerFeuillesTechniques()
    Dim ws As Worksheet, n As Long
    Feuille(ONGLET_DASHBOARD).Activate
    For Each ws In ThisWorkbook.Worksheets
        If EstTechnique(ws.Name) Then
            ws.Visible = xlSheetVeryHidden
            n = n + 1
        End If
    Next ws
    MsgBox n & " feuille(s) technique(s) très masquée(s)." & vbLf & _
           "Pour les retrouver : AfficherFeuillesTechniques (Alt+F8).", vbInformation, "Horizon_Contrat"
End Sub

Public Sub AfficherFeuillesTechniques()
    Dim ws As Worksheet, n As Long
    For Each ws In ThisWorkbook.Worksheets
        If EstTechnique(ws.Name) And ws.Visible <> xlSheetVisible Then
            ws.Visible = xlSheetVisible
            n = n + 1
        End If
    Next ws
    MsgBox n & " feuille(s) technique(s) réaffichée(s).", vbInformation, "Horizon_Contrat"
End Sub

Private Function EstTechnique(ByVal sNom As String) As Boolean
    EstTechnique = (sNom Like "REF_*") Or (sNom Like "DATA_*") Or (sNom Like "CALC_*")
End Function
