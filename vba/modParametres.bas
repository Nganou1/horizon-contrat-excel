Attribute VB_Name = "modParametres"
Option Explicit
'==============================================================================
' Horizon_Contrat - paramètres et noms d'objets
'
' Règle du projet : aucun chemin, seuil ou adresse dans le code. Tout vient de
' PARAMETRES par Param("Cle"). Chaque clé de PARAMETRES est aussi une cellule
' nommée du même nom : Param lit ce nom, et cherche la clé dans la colonne B en
' secours.
'
' Les noms d'onglets sont regroupés ici : si un onglet est renommé, une seule
' ligne change. (Variante du corrigé : CodeName des feuilles, voir LISEZ-MOI.)
'==============================================================================

Public Const ONGLET_DASHBOARD As String = "DASHBOARD"
Public Const ONGLET_FICHE As String = "FICHE_MAGASIN"
Public Const ONGLET_PARAMETRES As String = "PARAMETRES"
Public Const ONGLET_JOURNAL As String = "JOURNAL"
Public Const ONGLET_SAISIE As String = "SAISIE_RELEVES"
Public Const ONGLET_REF_MAGASINS As String = "REF_MAGASINS"

Public Const TBL_MAGASINS As String = "tblMagasins"
Public Const TBL_JOURNAL As String = "tblJournal"
Public Const TBL_SAISIE As String = "tblSaisieReleves"

' Garde-fou contre les fausses manipulations, pas une sécurité : un mot de
' passe de feuille Excel se retire en quelques secondes avec des outils publics.
Public Const MOT_DE_PASSE_FEUILLES As String = "HC-2026"

Public Function Param(ByVal sCle As String) As Variant
    Dim r As Range, c As Range

    Set r = PlageNommee(sCle)
    If Not r Is Nothing Then
        Param = r.Cells(1, 1).Value
        Exit Function
    End If

    Set c = Feuille(ONGLET_PARAMETRES).Columns("B").Find(What:=sCle, LookIn:=xlValues, _
                                                         LookAt:=xlWhole, MatchCase:=False)
    If Not c Is Nothing Then
        Param = c.Offset(0, 1).Value
        Exit Function
    End If

    Err.Raise vbObjectError + 513, "Param", "Paramètre introuvable : " & sCle & _
              ". Lance InstallerReporting, ou ajoute la clé dans PARAMETRES."
End Function

Public Function ParamOuDefaut(ByVal sCle As String, ByVal vDefaut As Variant) As Variant
    Dim v As Variant
    On Error GoTo Absent
    v = Param(sCle)
    If IsEmpty(v) Then v = vDefaut
    ParamOuDefaut = v
    Exit Function
Absent:
    ParamOuDefaut = vDefaut
End Function

Public Sub EcrireParam(ByVal sCle As String, ByVal vValeur As Variant)
    Dim r As Range
    Set r = PlageNommee(sCle)
    If r Is Nothing Then Err.Raise vbObjectError + 514, "EcrireParam", "Paramètre introuvable : " & sCle
    r.Cells(1, 1).Value = vValeur
End Sub

' Un chemin de PARAMETRES encore à l'état "[À RENSEIGNER] ..." n'est pas un chemin.
Public Function CheminRenseigne(ByVal sCle As String) As String
    Dim s As String
    s = Trim$(CStr(ParamOuDefaut(sCle, "")))
    If s = "" Or Left$(s, 1) = "[" Then Exit Function
    CheminRenseigne = AvecSeparateur(s)
End Function

Public Function CheminSorties() As String
    Dim s As String
    s = CheminRenseigne("Chemin_Sorties")
    If s = "" Then s = AvecSeparateur(ThisWorkbook.Path) & "Sorties" & Application.PathSeparator
    CheminSorties = s
End Function
