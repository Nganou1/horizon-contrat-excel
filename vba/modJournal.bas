Attribute VB_Name = "modJournal"
Option Explicit
'==============================================================================
' Horizon_Contrat - journal d'exécution (tableau tblJournal, onglet JOURNAL)
'
' Colonnes : Horodatage | Utilisateur | Procedure | Traites | Ignores |
'            Duree_s | Message
'==============================================================================

Public Sub Journaliser(ByVal sProcedure As String, _
                       ByVal lTraites As Long, _
                       ByVal lIgnores As Long, _
                       ByVal dSecondes As Double, _
                       ByVal sMessage As String)
    Dim lo As ListObject, r As Range, bEvents As Boolean

    ' Le seul On Error Resume Next légitime du projet : la journalisation ne
    ' doit jamais faire échouer l'opération qu'elle journalise.
    On Error Resume Next
    bEvents = Application.EnableEvents
    Application.EnableEvents = False

    Set lo = ThisWorkbook.Worksheets(ONGLET_JOURNAL).ListObjects(TBL_JOURNAL)
    If Not lo Is Nothing Then
        Set r = NouvelleLigne(lo)
        If Not r Is Nothing Then
            r.Cells(1, 1).Value = Now
            r.Cells(1, 2).Value = Environ$("USERNAME")
            r.Cells(1, 3).Value = sProcedure
            r.Cells(1, 4).Value = lTraites
            r.Cells(1, 5).Value = lIgnores
            r.Cells(1, 6).Value = Round(dSecondes, 1)
            r.Cells(1, 7).Value = Left$(sMessage, 1000)
        End If
    End If

    Application.EnableEvents = bEvents
End Sub
