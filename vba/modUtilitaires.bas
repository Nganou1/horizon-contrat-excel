Attribute VB_Name = "modUtilitaires"
Option Explicit
'==============================================================================
' Horizon_Contrat - fonctions utilitaires partagées
'==============================================================================

Public Function Feuille(ByVal sNom As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sNom)
    On Error GoTo 0
    If ws Is Nothing Then
        Err.Raise vbObjectError + 520, "Feuille", "Onglet introuvable : " & sNom & _
                  " (renommé ? Corrige la constante dans modParametres.)"
    End If
    Set Feuille = ws
End Function

Public Function FeuilleExiste(ByVal sNom As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sNom)
    FeuilleExiste = Not ws Is Nothing
End Function

Public Function Tableau(ByVal sFeuille As String, ByVal sTable As String) As ListObject
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ThisWorkbook.Worksheets(sFeuille).ListObjects(sTable)
    On Error GoTo 0
    Set Tableau = lo
    If lo Is Nothing Then
        Err.Raise vbObjectError + 521, "Tableau", "Tableau " & sTable & " introuvable sur " & sFeuille & _
                  ". Lance InstallerReporting."
    End If
End Function

Public Function PlageNommee(ByVal sNom As String) As Range
    On Error Resume Next
    Set PlageNommee = ThisWorkbook.Names(sNom).RefersToRange
End Function

Public Function Plage(ByVal sNom As String) As Range
    Dim r As Range
    Set r = PlageNommee(sNom)
    If r Is Nothing Then Err.Raise vbObjectError + 522, "Plage", "Nom introuvable : " & sNom & " (lance InstallerReporting)"
    Set Plage = r
End Function

' Ligne à écrire dans un tableau structuré : réutilise la ligne vide d'un
' tableau qui vient d'être créé, sinon ListRows.Add (jamais de recherche xlUp).
Public Function NouvelleLigne(ByVal lo As ListObject) As Range
    If lo.ListRows.Count = 1 Then
        If Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then
            Set NouvelleLigne = lo.DataBodyRange.Rows(1)
            Exit Function
        End If
    End If
    Set NouvelleLigne = lo.ListRows.Add.Range
End Function

Public Function NbLignesUtiles(ByVal lo As ListObject) As Long
    Dim n As Long
    If lo.DataBodyRange Is Nothing Then Exit Function
    n = lo.ListRows.Count
    If n = 1 Then
        If Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then n = 0
    End If
    NbLignesUtiles = n
End Function

Public Function AvecSeparateur(ByVal s As String) As String
    s = Trim$(s)
    If s <> "" And Right$(s, 1) <> Application.PathSeparator Then s = s & Application.PathSeparator
    AvecSeparateur = s
End Function

Public Function DossierExiste(ByVal s As String) As Boolean
    If Len(s) < 3 Then Exit Function
    On Error Resume Next
    DossierExiste = (Dir$(AvecSeparateur(s), vbDirectory) <> vbNullString)
End Function

' ...\Sorties\AAAA-MM-JJ\ ; crée la racine (un niveau) puis le dossier du jour.
Public Function CreerDossierDuJour(ByVal sRacine As String) As String
    Dim sChemin As String
    sRacine = AvecSeparateur(sRacine)
    If Not DossierExiste(sRacine) Then MkDir Left$(sRacine, Len(sRacine) - 1)
    sChemin = sRacine & Format$(Date, "yyyy-mm-dd") & Application.PathSeparator
    If Not DossierExiste(sChemin) Then MkDir Left$(sChemin, Len(sChemin) - 1)
    CreerDossierDuJour = sChemin
End Function

Public Function DossierDuJour() As String
    DossierDuJour = CheminSorties() & Format$(Date, "yyyy-mm-dd") & Application.PathSeparator
End Function

' Les noms de magasin contiennent "Gare / Aéroport" : le "/" fait échouer
' ExportAsFixedFormat avec une erreur 1004 qui ne cite ni le fichier ni le caractère.
Public Function NomFichierSain(ByVal sNom As String) As String
    Dim v As Variant
    For Each v In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        sNom = Replace$(sNom, CStr(v), "-")
    Next v
    NomFichierSain = Trim$(sNom)
End Function

Public Function FormatDuree(ByVal dSecondes As Double) As String
    If dSecondes < 60 Then
        FormatDuree = Format$(dSecondes, "0") & " s"
    Else
        FormatDuree = Int(dSecondes / 60) & " min " & Format$(dSecondes - 60 * Int(dSecondes / 60), "00") & " s"
    End If
End Function

' Nombre saisi avec virgule ou point, espaces de milliers tolérés.
' Val() ne connaît que le point, et ne dépend donc pas des paramètres régionaux.
Public Function ValeurNum(ByVal v As Variant) As Double
    Dim s As String
    s = Trim$(CStr(v))
    s = Replace$(s, " ", "")
    s = Replace$(s, Chr$(160), "")
    s = Replace$(s, ChrW(8239), "")
    s = Replace$(s, "€", "")
    s = Replace$(s, ",", ".")
    If s = "" Then Exit Function
    ValeurNum = Val(s)
End Function

' Date saisie JJ/MM/AAAA (ou JJ-MM-AAAA, JJ.MM.AAAA, année sur 2 chiffres).
' Pas de CDate/IsDate : leur lecture dépend des paramètres régionaux du poste,
' et "03/04/2026" devient le 4 mars sur un Windows réglé en anglais.
Public Function LireDateFr(ByVal s As String, ByRef dDate As Date) As Boolean
    Dim p() As String, j As Long, m As Long, a As Long
    s = Trim$(s)
    s = Replace$(Replace$(s, "-", "/"), ".", "/")
    p = Split(s, "/")
    If UBound(p) <> 2 Then Exit Function
    If Not (IsNumeric(p(0)) And IsNumeric(p(1)) And IsNumeric(p(2))) Then Exit Function
    j = CLng(p(0)): m = CLng(p(1)): a = CLng(p(2))
    If a < 100 Then a = a + 2000
    If m < 1 Or m > 12 Or j < 1 Or j > 31 Or a < 2000 Or a > 2100 Then Exit Function
    If Day(DateSerial(a, m, j)) <> j Then Exit Function
    dDate = DateSerial(a, m, j)
    LireDateFr = True
End Function

Public Function FinDeMois(ByVal d As Date) As Date
    FinDeMois = DateSerial(Year(d), Month(d) + 1, 0)
End Function

Public Function DebutDeMois(ByVal d As Date) As Date
    DebutDeMois = DateSerial(Year(d), Month(d), 1)
End Function

' Accès en lecture à une colonne de tableau structuré par son en-tête.
Public Function IndexColonne(ByVal lo As ListObject, ByVal sEntete As String) As Long
    Dim n As Long
    On Error Resume Next
    n = lo.ListColumns(sEntete).Index
    On Error GoTo 0
    If n = 0 Then
        Err.Raise vbObjectError + 523, "IndexColonne", "Colonne " & sEntete & " absente de " & lo.Name
    End If
    IndexColonne = n
End Function

Public Function EchapperHtml(ByVal s As String) As String
    s = Replace$(s, "&", "&amp;")
    s = Replace$(s, "<", "&lt;")
    s = Replace$(s, ">", "&gt;")
    EchapperHtml = s
End Function
