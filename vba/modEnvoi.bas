Attribute VB_Name = "modEnvoi"
Option Explicit
'==============================================================================
' Horizon_Contrat - envoi des fiches aux franchisés par Outlook
'
' Sécurité, non négociable :
'   - Mode_Envoi_Reel = FAUX (valeur livrée) : les messages vont dans les
'     BROUILLONS d'Outlook (.Save). Rien ne part.
'   - Mode_Envoi_Reel = VRAI : confirmation explicite avant le premier envoi,
'     et jamais d'envoi vers une adresse en .example (adresses fictives
'     livrées) : ces messages restent en brouillon.
'   - Late binding (CreateObject, variables Object) : le classeur compile et
'     s'ouvre sur un poste sans Outlook.
'==============================================================================

Private Const OL_MAIL_ITEM As Long = 0

Public Sub EnvoyerFichesFranchises()
    Dim oOutlook As Object, lo As ListObject, v As Variant
    Dim iCode As Long, iNom As Long, iVille As Long, iResp As Long
    Dim i As Long, lPrets As Long, lCrees As Long, lEnvoyes As Long, lEchecs As Long, lFictifs As Long
    Dim sDossier As String, sPJ As String, sErr As String, sEchecs As String, sResultat As String
    Dim bReel As Boolean, dDebut As Double, dMois As Date

    dDebut = Timer
    On Error GoTo Echec

    sDossier = DossierDuJour()
    If Not DossierExiste(sDossier) Then
        MsgBox "Aucun PDF du jour dans :" & vbLf & sDossier & vbLf & vbLf & _
               "Lance d'abord GenererFichesPDF.", vbExclamation, "Horizon_Contrat"
        Exit Sub
    End If

    Set lo = Tableau(ONGLET_REF_MAGASINS, TBL_MAGASINS)
    iCode = IndexColonne(lo, "Code_Magasin")
    iNom = IndexColonne(lo, "Nom_Magasin")
    iVille = IndexColonne(lo, "Ville")
    iResp = IndexColonne(lo, "Franchise_Responsable")
    v = lo.DataBodyRange.Value

    For i = 1 To UBound(v, 1)
        If Dir$(CheminPJ(sDossier, CStr(v(i, iCode)))) <> vbNullString Then lPrets = lPrets + 1
    Next i
    If lPrets = 0 Then
        MsgBox "Le dossier du jour ne contient aucune fiche.", vbExclamation, "Horizon_Contrat"
        Exit Sub
    End If

    bReel = CBool(ParamOuDefaut("Mode_Envoi_Reel", False))
    If bReel Then
        If MsgBox("ENVOI RÉEL : " & lPrets & " message(s) vont partir maintenant." & vbLf & _
                  "(Les adresses en .example restent en brouillon.)" & vbLf & vbLf & _
                  "Confirmer l'envoi ?", vbYesNo + vbExclamation + vbDefaultButton2, _
                  "Horizon_Contrat - envoi réel") <> vbYes Then Exit Sub
    Else
        If MsgBox(lPrets & " message(s) vont être créés dans les BROUILLONS d'Outlook." & vbLf & _
                  "Rien ne sera envoyé (Mode_Envoi_Reel = FAUX)." & vbLf & vbLf & "Continuer ?", _
                  vbOKCancel + vbInformation, "Horizon_Contrat - brouillons") <> vbOK Then Exit Sub
    End If

    On Error Resume Next
    Set oOutlook = CreateObject("Outlook.Application")
    On Error GoTo Echec
    If oOutlook Is Nothing Then
        MsgBox "Outlook n'est pas disponible sur ce poste.", vbExclamation, "Horizon_Contrat"
        Exit Sub
    End If

    dMois = DernierMoisVentes()
    For i = 1 To UBound(v, 1)
        sPJ = CheminPJ(sDossier, CStr(v(i, iCode)))
        If Dir$(sPJ) <> vbNullString Then            ' pas de PDF, pas de message
            Application.StatusBar = "Horizon_Contrat - message " & (lCrees + lEchecs + 1) & " / " & lPrets
            sResultat = CreerMessage(oOutlook, CStr(v(i, iCode)), CStr(v(i, iNom)), CStr(v(i, iVille)), _
                                     CStr(v(i, iResp)), sPJ, dMois, bReel)
            Select Case sResultat
                Case "envoye": lCrees = lCrees + 1: lEnvoyes = lEnvoyes + 1
                Case "brouillon": lCrees = lCrees + 1
                Case "fictif": lCrees = lCrees + 1: lFictifs = lFictifs + 1
                Case Else
                    lEchecs = lEchecs + 1
                    sEchecs = sEchecs & vbLf & v(i, iCode) & " : " & sResultat
            End Select
            DoEvents
        End If
    Next i

Fin:
    On Error Resume Next
    Set oOutlook = Nothing
    Application.StatusBar = False
    If sErr <> "" Then
        Journaliser "EnvoyerFichesFranchises", lCrees, lEchecs, Timer - dDebut, "ERREUR : " & sErr
        MsgBox "Envoi interrompu : " & sErr & vbLf & lCrees & " message(s) traités avant l'erreur.", _
               vbExclamation, "Horizon_Contrat"
    Else
        Journaliser "EnvoyerFichesFranchises", lCrees, lEchecs, Timer - dDebut, _
                    IIf(bReel, "Envoi réel : " & lEnvoyes & " envoyé(s), " & lFictifs & " adresse(s) fictive(s) en brouillon", _
                               "Brouillons") & IIf(sEchecs <> "", " ; échecs :" & Replace$(sEchecs, vbLf, " "), "")
        MsgBox IIf(bReel, lEnvoyes & " message(s) envoyé(s), " & (lCrees - lEnvoyes) & " laissé(s) en brouillon.", _
                          lCrees & " brouillon(s) créé(s) dans Outlook.") & _
               IIf(lEchecs > 0, vbLf & lEchecs & " échec(s) :" & sEchecs, ""), _
               IIf(lEchecs > 0, vbExclamation, vbInformation), "Horizon_Contrat"
    End If
    Exit Sub

Echec:
    sErr = Err.Description & " (erreur " & Err.Number & ")"
    Resume Fin
End Sub

'------------------------------------------------------------------------------
' Un message. Renvoie "envoye", "brouillon", "fictif" ou le texte de l'erreur :
' une erreur sur un franchisé ne bloque pas les suivants.
'------------------------------------------------------------------------------
Private Function CreerMessage(ByVal oOutlook As Object, ByVal sCode As String, ByVal sNom As String, _
                              ByVal sVille As String, ByVal sResponsable As String, ByVal sPJ As String, _
                              ByVal dMois As Date, ByVal bReel As Boolean) As String
    Dim oMail As Object, sA As String, sCopie As String, bFictif As Boolean
    On Error GoTo Echec

    sA = AdresseFranchise(sCode)
    sCopie = Trim$(CStr(ParamOuDefaut("Mail_Copie_Direction", "")))
    bFictif = EstAdresseFictive(sA)

    Set oMail = oOutlook.CreateItem(OL_MAIL_ITEM)
    With oMail
        .To = sA
        If sCopie <> "" Then
            If Not (bReel And EstAdresseFictive(sCopie)) Then .CC = sCopie
        End If
        .Subject = "Horizon_Contrat - rapport " & NomMois(dMois) & " - " & sCode & " " & sVille
        .HTMLBody = CorpsMail(sCode, sNom, sResponsable, dMois)
        .Attachments.Add sPJ
        If bReel And Not bFictif Then
            .Send
            CreerMessage = "envoye"
        Else
            .Save                                  ' dossier Brouillons
            CreerMessage = IIf(bReel, "fictif", "brouillon")
        End If
    End With
    Set oMail = Nothing
    Exit Function

Echec:
    CreerMessage = Err.Description
End Function

'------------------------------------------------------------------------------
' Contenu
'------------------------------------------------------------------------------
Private Function CorpsMail(ByVal sCode As String, ByVal sNom As String, ByVal sResponsable As String, _
                           ByVal dMois As Date) As String
    Dim dCA As Double, dCout As Double, lTickets As Double, dCA1 As Double
    Dim dTaux As Double, dTM As Double, dEvol As Double, bEvol As Boolean
    Dim vTauxEx As Variant, vSeuil As Variant, vAtteinte As Variant, dSeuilBudget As Double
    Dim s As String, sAlerte As String, dMoisN1 As Date

    dMoisN1 = DateSerial(Year(dMois) - 1, Month(dMois), 1)
    dCA = SommeVentes("V_CANet", sCode, dMois)
    dCout = SommeVentes("V_Cout", sCode, dMois)
    lTickets = SommeVentes("V_Tickets", sCode, dMois)
    dCA1 = SommeVentes("V_CANet", sCode, dMoisN1)
    If dCA > 0 Then dTaux = (dCA - dCout) / dCA
    If lTickets > 0 Then dTM = dCA / lTickets
    If dCA1 > 0 Then
        dEvol = dCA / dCA1 - 1
        bEvol = True
    End If

    vTauxEx = ValeurMagasin("CM_Taux_Marge_N", sCode)
    vSeuil = ParamNomme("D_Seuil_Taux_Alerte")
    vAtteinte = ValeurMagasin("CM_Atteinte_Budget", sCode)
    dSeuilBudget = CDbl(ParamOuDefaut("Seuil_Atteinte_Budget", 0.9))
    If IsNumeric(vTauxEx) And IsNumeric(vSeuil) And Not IsEmpty(vTauxEx) Then
        If CDbl(vTauxEx) < CDbl(vSeuil) Then
            sAlerte = sAlerte & "<li>Votre taux de marge sur l'exercice (" & Pct(CDbl(vTauxEx), 1) & _
                      ") est sous le seuil d'alerte du réseau (" & Pct(CDbl(vSeuil), 1) & _
                      "). Votre animateur réseau vous contactera pour un point sur le coût matière.</li>"
        End If
    End If
    If IsNumeric(vAtteinte) And Not IsEmpty(vAtteinte) Then
        If CDbl(vAtteinte) < dSeuilBudget Then
            sAlerte = sAlerte & "<li>Votre atteinte du budget 2026 (janvier-août) est de " & Pct(CDbl(vAtteinte), 1) & _
                      ", sous l'objectif de " & Pct(dSeuilBudget, 0) & ".</li>"
        End If
    End If

    s = "<div style=""font-family:Segoe UI,Arial,sans-serif;font-size:10.5pt;color:#14332C"">"
    s = s & "<p>Bonjour " & EchapperHtml(sResponsable) & ",</p>"
    s = s & "<p>Vous trouverez en pièce jointe la fiche de pilotage de <b>" & EchapperHtml(sNom) & "</b> (" & sCode & ").</p>"
    s = s & "<p><b>Vos indicateurs de " & NomMois(dMois) & "</b></p>"
    s = s & "<table style=""border-collapse:collapse;font-size:10.5pt"">"
    s = s & LigneHtml("CA net HT", Euros(dCA, 0) & IIf(bEvol, " (" & IIf(dEvol >= 0, "+", "") & Pct(dEvol, 1) & " vs " & NomMois(dMoisN1) & ")", ""))
    s = s & LigneHtml("Taux de marge brute", Pct(dTaux, 1))
    s = s & LigneHtml("Ticket moyen HT", Euros(dTM, 2))
    s = s & "</table>"
    If sAlerte <> "" Then
        s = s & "<p style=""color:#B4552F""><b>Points d'attention</b></p><ul style=""color:#B4552F"">" & sAlerte & "</ul>"
    End If
    s = s & "<p>Bonne réception,<br>Le contrôle de gestion Horizon_Contrat</p>"
    s = s & "<p style=""font-size:8.5pt;color:#6B7570"">Message préparé automatiquement à partir du classeur de pilotage.</p>"
    s = s & "</div>"
    CorpsMail = s
End Function

Private Function LigneHtml(ByVal sLibelle As String, ByVal sValeur As String) As String
    LigneHtml = "<tr><td style=""padding:3px 18px 3px 0;color:#6B7570"">" & sLibelle & _
                "</td><td style=""padding:3px 0;text-align:right""><b>" & sValeur & "</b></td></tr>"
End Function

'------------------------------------------------------------------------------
' Adresses : construites par formule à partir du code magasin et du domaine
' de PARAMETRES. Le domaine livré est fictif (.example, réservé par la norme
' RFC 2606 : aucun message ne peut y arriver).
'------------------------------------------------------------------------------
Public Function AdresseFranchise(ByVal sCode As String) As String
    AdresseFranchise = LCase$(Trim$(sCode)) & "@" & Trim$(CStr(ParamOuDefaut("Domaine_Mail_Franchises", "horizon-contrat.example")))
End Function

Public Function EstAdresseFictive(ByVal sAdresse As String) As Boolean
    sAdresse = LCase$(Trim$(sAdresse))
    EstAdresseFictive = (Right$(sAdresse, 8) = ".example") Or (InStr(sAdresse, "@") = 0)
End Function

'------------------------------------------------------------------------------
' Lectures dans le classeur
'------------------------------------------------------------------------------
Private Function CheminPJ(ByVal sDossier As String, ByVal sCode As String) As String
    CheminPJ = sDossier & "Fiche_" & NomFichierSain(sCode) & ".pdf"
End Function

Private Function DernierMoisVentes() As Date
    DernierMoisVentes = CDate(Application.WorksheetFunction.Max(Plage("V_Mois")))
End Function

Private Function SommeVentes(ByVal sColonne As String, ByVal sCode As String, ByVal dMois As Date) As Double
    SommeVentes = Application.WorksheetFunction.SumIfs(Plage(sColonne), Plage("V_Mag"), sCode, _
                                                      Plage("V_Mois"), CDbl(dMois))
End Function

Private Function ValeurMagasin(ByVal sColonne As String, ByVal sCode As String) As Variant
    Dim k As Variant
    k = Application.Match(sCode, Plage("CM_Code_Magasin"), 0)
    If IsError(k) Then
        ValeurMagasin = Empty
    Else
        ValeurMagasin = Plage(sColonne).Cells(CLng(k), 1).Value
    End If
End Function

Private Function ParamNomme(ByVal sNom As String) As Variant
    Dim r As Range
    Set r = PlageNommee(sNom)
    If r Is Nothing Then
        ParamNomme = Empty
    Else
        ParamNomme = r.Cells(1, 1).Value
    End If
End Function

Private Function NomMois(ByVal d As Date) As String
    NomMois = Choose(Month(d), "janvier", "février", "mars", "avril", "mai", "juin", "juillet", _
                     "août", "septembre", "octobre", "novembre", "décembre") & " " & Year(d)
End Function

' Nombres écrits à la main : Format$ suit les paramètres régionaux du poste,
' et un corps de mail ne doit pas changer de séparateurs selon le PC qui l'envoie.
Private Function NombreFr(ByVal d As Double, ByVal nDec As Long) As String
    Dim bNeg As Boolean, dEnt As Double, lDec As Long, sEnt As String, r As String
    bNeg = (d < 0)
    d = Round(Abs(d), nDec)
    dEnt = Int(d)
    lDec = CLng(Round((d - dEnt) * 10 ^ nDec, 0))
    If nDec > 0 And lDec >= 10 ^ nDec Then
        dEnt = dEnt + 1
        lDec = 0
    End If
    sEnt = Format$(dEnt, "0")
    Do While Len(sEnt) > 3
        r = "&nbsp;" & Right$(sEnt, 3) & r
        sEnt = Left$(sEnt, Len(sEnt) - 3)
    Loop
    r = sEnt & r
    If nDec > 0 Then r = r & "," & Format$(lDec, String$(nDec, "0"))
    If bNeg And (dEnt > 0 Or lDec > 0) Then r = "-" & r
    NombreFr = r
End Function

Private Function Euros(ByVal d As Double, ByVal nDec As Long) As String
    Euros = NombreFr(d, nDec) & "&nbsp;&euro;"
End Function

Private Function Pct(ByVal d As Double, ByVal nDec As Long) As String
    Pct = NombreFr(d * 100, nDec) & "&nbsp;%"
End Function
