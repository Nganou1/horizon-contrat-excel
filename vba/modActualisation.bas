Attribute VB_Name = "modActualisation"
Option Explicit
'==============================================================================
' Horizon_Contrat - actualisation synchrone des requêtes, du modèle et des TCD
'
' RefreshAll rend la main tout de suite : sans BackgroundQuery = False et
' l'attente qui suit, la ligne suivante s'exécute sur les données de la veille.
'==============================================================================

Public Sub ActualiserTout()
    Dim cn As WorkbookConnection, dDebut As Double, sErr As String
    Dim lCalc As XlCalculation, bScreen As Boolean, n As Long

    If ThisWorkbook.Queries.Count = 0 Then
        MsgBox "Aucune requête Power Query dans ce classeur." & vbLf & _
               "Lance d'abord ConstruireModele (Alt+F8).", vbInformation, "Horizon_Contrat"
        Exit Sub
    End If

    dDebut = Timer
    lCalc = Application.Calculation
    bScreen = Application.ScreenUpdating
    Application.EnableCancelKey = xlErrorHandler
    On Error GoTo Echec

    Application.ScreenUpdating = False
    Application.StatusBar = "Horizon_Contrat - actualisation des requêtes et du modèle..."

    ' Certaines connexions (modèle) n'exposent pas BackgroundQuery : on ignore
    ' l'erreur pour celles-là seulement, puis on remet la gestion normale.
    For Each cn In ThisWorkbook.Connections
        On Error Resume Next
        Select Case cn.Type
            Case xlConnectionTypeOLEDB
                cn.OLEDBConnection.BackgroundQuery = False
            Case xlConnectionTypeODBC
                cn.ODBCConnection.BackgroundQuery = False
        End Select
        On Error GoTo Echec
        n = n + 1
    Next cn

    ' En calcul manuel, CalculationState reste "en attente" indéfiniment :
    ' on passe en automatique le temps de l'actualisation (restauré dans Fin).
    Application.Calculation = xlCalculationAutomatic
    ThisWorkbook.RefreshAll
    Application.CalculateUntilAsyncQueriesDone
    Do While Application.CalculationState <> xlDone
        If Timer - dDebut > 900 Then Err.Raise vbObjectError + 530, , "actualisation toujours en cours après 15 minutes"
        DoEvents
    Loop

    Application.StatusBar = "Horizon_Contrat - recalcul complet..."
    Application.CalculateFull

    On Error Resume Next
    EcrireParam "Derniere_Actualisation", Now   ' absent si InstallerReporting n'a pas tourné
    On Error GoTo Echec
    Journaliser "ActualiserTout", n, 0, Timer - dDebut, ""

Fin:
    On Error Resume Next
    Application.ScreenUpdating = bScreen
    Application.Calculation = lCalc
    Application.StatusBar = False
    Application.EnableCancelKey = xlInterrupt
    If sErr <> "" Then
        Journaliser "ActualiserTout", 0, 0, Timer - dDebut, sErr
        MsgBox "L'actualisation a échoué : " & sErr, vbExclamation, "Horizon_Contrat"
    Else
        MsgBox n & " connexion(s) actualisée(s) en " & FormatDuree(Timer - dDebut) & ".", _
               vbInformation, "Horizon_Contrat"
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
