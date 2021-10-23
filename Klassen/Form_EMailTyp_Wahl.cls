VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EMailTyp_Wahl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub startMail_Click()
On Error GoTo Err_startMail_Click
    Dim stAnrede As String
    Dim stVorname As String
    Dim stNachname As String
    Dim stAdresse As String
    Dim stPostleitzahl As String
    Dim stOrt As String
    Dim stMailPrivat As String
    Dim stLieb As String
    Dim stRebenart As String
    Dim lgRebeNr As Long
    Dim stInternetlink As String
    Dim stFirma As String
    Dim DatumVersand As Date
    
    Dim stDocName As String
    Dim stqueryname As String
    Dim stsql As String
    
    Dim tgtdb As Database
    Dim tgtqry As QueryDef
    Dim tgtrs As Recordset
    
    Dim emailtypenqry As QueryDef
    Dim emailtypenrs As Recordset
    
    Dim stEMailBetreff As String
    Dim stAttachement As String
    Dim stAbfragEmailAdressen As String
    Dim lgEMailTypID  As String
    
    Dim stEMailNachricht As String
    
    Dim emtxtqry As QueryDef 'E-Mail Textschablone
    Dim emtxtrs As Recordset
    
    Dim varRegion As Variant
    Dim varPreisKat As Variant
    Dim varStatusNr As Variant
    Dim antwort As Variant
    
    Dim intCheckFelder As Integer
    antwort = MsgBox("Wollen Sie wirklich e-mails verschicken?", vbYesNo + vbCritical + vbDefaultButton1, "Warnhinweis")
    If Not (antwort = vbYes) Then GoTo Exit_startMail_Click
    Set tgtdb = CurrentDb
    'E-Mail Typ: Eigenschaften bestimmen
    stqueryname = "EMailTyp_Selektion"
    lgEMailTypID = Me![EMailTypID]
    stsql = "SELECT a.EMailTypID, " & vbNewLine
    stsql = stsql & "a.EMailBetreff, a.Attachement, a.AbfragEmailAdressen" & vbNewLine
    stsql = stsql & "FROM EMAILTYPEN AS a" & vbNewLine
    stsql = stsql & "WHERE a.EMailTypID = " & lgEMailTypID & ";"
    Set emailtypenqry = tgtdb.QueryDefs(stqueryname)
    emailtypenqry.SQL = stsql
    Set emailtypenrs = emailtypenqry.OpenRecordset(dbOpenSnapshot)
    If emailtypenrs.EOF Then Err.Raise vbObjectError + 83
    stEMailBetreff = emailtypenrs!EMailBetreff
    stAttachement = Nz(emailtypenrs!Attachement, conDefaultNullString)
    stAbfragEmailAdressen = emailtypenrs!AbfragEmailAdressen
    emailtypenqry.Close
    emailtypenrs.Close
    
   
    'Statische Datensätze für die E-Mail Nachricht bereitstellen (Textschablone)
    DatumVersand = Date
       
    stqueryname = "Emailtexte_selektion"
    lgEMailTypID = Me![EMailTypID]
    stsql = "SELECT e.ReihenfolgeNr, e.Emailtext, e.LeerschlagAnzahl, e.SeriendruckfeldID, e.ZeilenumbruchAnzahl" & vbNewLine
    stsql = stsql & "FROM EMAILTEXTE AS e" & vbNewLine
    stsql = stsql & "WHERE e.EMailTypID = " & lgEMailTypID & vbNewLine
    stsql = stsql & "ORDER BY e.ReihenfolgeNr;"
    Set emtxtqry = tgtdb.QueryDefs(stqueryname)
    emtxtqry.SQL = stsql
    Set emtxtrs = emtxtqry.OpenRecordset(dbOpenSnapshot)

    If Me![Versandart] = conEmailMassenversand Then
        'Adressen abfragen
        Set tgtqry = tgtdb.QueryDefs(stAbfragEmailAdressen)
        varRegion = Forms!EMAILTYPEN!Region
        varPreisKat = Forms!EMAILTYPEN!PreiskatNr
        varStatusNr = Forms!EMAILTYPEN!StatusNr
        If (Not (IsNull(varRegion)) And (Not (IsNull(varPreisKat))) And ((IsNull(varStatusNr)) Or (tgtqry.Parameters.Count = 2))) Then
            tgtqry.Parameters(0) = varRegion
            tgtqry.Parameters(1) = varPreisKat
        End If
        If (Not (IsNull(varRegion)) And (Not (IsNull(varPreisKat))) And (Not (IsNull(varStatusNr))) And (tgtqry.Parameters.Count = 3)) Then
            tgtqry.Parameters(0) = varRegion
            tgtqry.Parameters(1) = varPreisKat
            tgtqry.Parameters(2) = varStatusNr
            
        End If
        Set tgtrs = tgtqry.OpenRecordset(dbOpenSnapshot)
        
        If tgtrs.EOF Then Err.Raise vbObjectError + 81
        intCheckFelder = checkAddressFields(tgtrs)
        
    
        'E-Mails ausfüllen und verschicken
        Do While tgtrs.EOF <> True
            stAnrede = Nz(tgtrs!Anrede, conDefaultNullString)
            stVorname = Nz(tgtrs!Vorname, conDefaultNullString)
            stNachname = Nz(tgtrs!Nachname, conDefaultNullString)
            stAdresse = Nz(tgtrs!Adresse, conDefaultNullString)
            stPostleitzahl = Nz(tgtrs!Postleitzahl, conDefaultNullString)
            stOrt = Nz(tgtrs!Ort, conDefaultNullString)
            stMailPrivat = Nz(tgtrs!MailPrivat, conDefaultNullString)
            stLieb = Nz(tgtrs!Lieb, conDefaultNullString)
            'Bugfix Ph. Flach - 17. Jan. 2017: stFirma nicht initialisiert
            stFirma = Nz(tgtrs!Firma, conDefaultNullString)
            If intCheckFelder = conFelderAlleVorhanden Then
                stRebenart = Nz(tgtrs![Rebenart], conDefaultNullString)
                lgRebeNr = Nz(tgtrs![RebeNr], conDefaultNullZahl)
                stInternetlink = Nz(tgtrs![Internetlink], conDefaultNullString)
            ElseIf intCheckFelder = conFeldRebenartVorhanden Then
                stRebenart = Nz(tgtrs![Rebenart], conDefaultNullString)
                lgRebeNr = conDefaultNullZahl
                stInternetlink = conDefaultNullString
            ElseIf intCheckFelder = conFeldRebenartInternetLinkVorhanden Then
                stRebenart = Nz(tgtrs![Rebenart], conDefaultNullString)
                lgRebeNr = conDefaultNullZahl
                stInternetlink = Nz(tgtrs![Internetlink], conDefaultNullString)
            ElseIf intCheckFelder = conFeldRebeNrVorhanden Then
                lgRebeNr = Nz(tgtrs![RebeNr], conDefaultNullZahl)
                stRebenart = conDefaultNullString
                stInternetlink = conDefaultNullString
            ElseIf intCheckFelder = conFeldRebeNrInternetLinkVorhanden Then
                lgRebeNr = Nz(tgtrs![RebeNr], conDefaultNullZahl)
                stRebenart = conDefaultNullString
                stInternetlink = Nz(tgtrs![Internetlink], conDefaultNullString)
            ElseIf intCheckFelder = conFeldRebenartRebeNrFehlt Then
                lgRebeNr = conDefaultNullZahl
                stRebenart = conDefaultNullString
                stInternetlink = Nz(tgtrs![Internetlink], conDefaultNullString)
            ElseIf intCheckFelder = conFeldInternetLinkFehlt Then
                stRebenart = Nz(tgtrs![Rebenart], conDefaultNullString)
                lgRebeNr = Nz(tgtrs![RebeNr], conDefaultNullZahl)
                stInternetlink = conDefaultNullString
            Else
                stRebenart = conDefaultNullString
                lgRebeNr = conDefaultNullZahl
                stInternetlink = conDefaultNullString
            End If
            'Textschablone ausfüllen
            stEMailNachricht = getEmailNachricht(emtxtrs, DatumVersand, stVorname, stNachname, _
              stAdresse, stPostleitzahl, stOrt, stMailPrivat, stLieb, stAnrede, stRebenart, lgRebeNr, stInternetlink, stFirma)
            'Für Debugging:
            'MsgBox stEMailNachricht
            'E-Mail versenden
            Call mailversand(stMailPrivat, stEMailBetreff, stEMailNachricht, stAttachement)
            'Zur nächsten Adresse gehen
            tgtrs.MoveNext
        Loop
        tgtqry.Close
        tgtrs.Close
        emtxtqry.Close

    Else
        stAnrede = Nz(Me![Anrede], conDefaultNullString)
        stVorname = Nz(Me![Vorname], conDefaultNullString)
        stNachname = Nz(Me![Nachname], conDefaultNullString)
        stAdresse = Nz(Me![Adresse], conDefaultNullString)
        stPostleitzahl = Nz(Me![Postleitzahl], conDefaultNullString)
        stOrt = Nz(Me![Ort], conDefaultNullString)
        stMailPrivat = Nz(Me![MailPrivat], conDefaultNullString)
        stLieb = Nz(Me![Lieb], conDefaultNullString)
        stRebenart = Nz(Me![Rebenart], conDefaultNullString)
        lgRebeNr = Nz(Me![RebeNr], conDefaultNullZahl)
        stInternetlink = Nz(Me![Internetlink], conDefaultNullString)
        stFirma = ""
        'Textschablone ausfüllen
        stEMailNachricht = getEmailNachricht(emtxtrs, DatumVersand, stVorname, stNachname, _
          stAdresse, stPostleitzahl, stOrt, stMailPrivat, stLieb, stAnrede, stRebenart, lgRebeNr, stInternetlink, stFirma)
        'E-Mail versenden
        'MsgBox stEMailNachricht
        Call mailversand(stMailPrivat, stEMailBetreff, stEMailNachricht, stAttachement)
    
    End If 'Versandart
    emtxtqry.Close
    emtxtrs.Close
    
Exit_startMail_Click:
    Exit Sub

Err_startMail_Click:
    If Err.Number = vbObjectError + 81 Then
        MsgBox "Keine E-mail-Adrsse", vbCritical
    ElseIf Err.Number = vbObjectError + 83 Then
        MsgBox "Ungültiger E-Mail Typ " & vbNewLine & _
        "(kein passender Datensatz in EMAILTYPEN gefunden!)", vbCritical
    Else
        MsgBox Err.Description
    End If
    Resume Exit_startMail_Click
    
End Sub


Private Sub EMailVorschau_Click()
On Error GoTo Err_EMailVorschau_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim stqueryname As String
    Dim lgEMailTypID As Long
    Dim stsql As String
    Dim emailtypenqry As QueryDef
    Dim emailtypenrs As Recordset
    Dim stEMailBetreff As String
    Dim stEMailNachricht As String
    Dim stAttachement As String
    Dim stAbfragEmailAdressen As String
    Dim tgtdb As Database
    Dim stAnrede As String
    Dim stVorname As String
    Dim stNachname As String
    Dim stAdresse As String
    Dim stFirma As String
    Dim stPostleitzahl As String
    Dim stOrt As String
    Dim stMailPrivat As String
    Dim stLieb As String
    Dim stRebenart As String
    Dim lgRebeNr As Long
    Dim stInternetlink As String
    Dim DatumVersand As Date
    Dim emtxtqry As QueryDef
    Dim emtxtrs As Recordset

    Set tgtdb = CurrentDb
    'E-Mail Typ: Eigenschaften bestimmen
    stqueryname = "EMailTyp_Selektion"
    lgEMailTypID = Me![EMailTypID]
    stsql = "SELECT a.EMailTypID, " & vbNewLine
    stsql = stsql & "a.EMailBetreff, a.Attachement, a.AbfragEmailAdressen" & vbNewLine
    stsql = stsql & "FROM EMAILTYPEN AS a" & vbNewLine
    stsql = stsql & "WHERE a.EMailTypID = " & lgEMailTypID & ";"
    Set emailtypenqry = tgtdb.QueryDefs(stqueryname)
    emailtypenqry.SQL = stsql
    Set emailtypenrs = emailtypenqry.OpenRecordset(dbOpenSnapshot)
    If emailtypenrs.EOF Then Err.Raise vbObjectError + 83
    stEMailBetreff = emailtypenrs!EMailBetreff
    stAttachement = Nz(emailtypenrs!Attachement, conDefaultNullString)
    stAbfragEmailAdressen = emailtypenrs!AbfragEmailAdressen
    emailtypenqry.Close
    emailtypenrs.Close


    'Statische Datensätze für die E-Mail Nachricht bereitstellen (Textschablone)
    DatumVersand = Date
    stqueryname = "Emailtexte_selektion"
    lgEMailTypID = Me![EMailTypID]
    stsql = "SELECT e.ReihenfolgeNr, e.Emailtext, e.LeerschlagAnzahl, e.SeriendruckfeldID, e.ZeilenumbruchAnzahl" & vbNewLine
    stsql = stsql & "FROM EMAILTEXTE AS e" & vbNewLine
    stsql = stsql & "WHERE e.EMailTypID = " & lgEMailTypID & vbNewLine
    stsql = stsql & "ORDER BY e.ReihenfolgeNr;"
    Set emtxtqry = tgtdb.QueryDefs(stqueryname)
    emtxtqry.SQL = stsql
    Set emtxtrs = emtxtqry.OpenRecordset(dbOpenSnapshot)

    
    'Textschablone ausfüllen
    
    If Me![Versandart] = conEmailMassenversand Then
        stAnrede = "Herr"
        stVorname = "Manfred"
        stNachname = "Mustermann"
        stAdresse = "Musterstrasse 99"
        stPostleitzahl = 9909
        stOrt = "Musterhausen"
        stMailPrivat = "m.mustermann@musterdomain.ch"
        stLieb = "Sehr geehrter Herr Mustermann"
        stRebenart = "CD"
        lgRebeNr = 19
        stInternetlink = "www.zizerser.ch/pdf/preisliste.doc"
        stFirma = "Musterhotel"
    Else
        stAnrede = Nz(Me!Anrede, conDefaultNullString)
        stVorname = Nz(Me!Vorname, conDefaultNullString)
        stNachname = Nz(Me!Nachname, conDefaultNullString)
        stAdresse = Nz(Me!Adresse, conDefaultNullString)
        stPostleitzahl = Nz(Me!Postleitzahl, conDefaultNullString)
        stOrt = Nz(Me!Ort, conDefaultNullString)
        stMailPrivat = Nz(Me!MailPrivat, conDefaultNullString)
        stLieb = Nz(Me!Lieb, conDefaultNullString)
        stRebenart = Nz(Me!Rebenart, conDefaultNullString)
        lgRebeNr = Nz(Me!RebeNr, conDefaultNullZahl)
        stInternetlink = Nz(Me!Internetlink, conDefaultNullString)
        stFirma = ""
    End If
    stEMailNachricht = getEmailNachricht(emtxtrs, DatumVersand, stVorname, stNachname, _
      stAdresse, stPostleitzahl, stOrt, stMailPrivat, stLieb, stAnrede, stRebenart, lgRebeNr, stInternetlink, stFirma)


    stDocName = "EMail_Testform"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    With Forms![EMail_Testform]
        ![EMailNachricht] = stEMailNachricht
        ![Betreff] = stEMailBetreff
    End With
    emtxtqry.Close
    emtxtrs.Close
Exit_EMailVorschau_Click:
    Exit Sub

Err_EMailVorschau_Click:

    If Err.Number = vbObjectError + 83 Then
        MsgBox "Ungültiger E-Mail Typ " & vbNewLine & _
        "(kein passender Datensatz in EMAILTYPEN gefunden!)", vbCritical
    Else
        MsgBox Err.Description
    End If
    Resume Exit_EMailVorschau_Click
    
End Sub
Private Sub Befehl23_Click()
On Error GoTo Err_Befehl23_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "EMAILTYPEN_HF"
    
    stLinkCriteria = "[EMailTypID]=" & Me![EMailTypID]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Befehl23_Click:
    Exit Sub

Err_Befehl23_Click:
    MsgBox Err.Description
    Resume Exit_Befehl23_Click
    
End Sub
Private Sub op_Adresskontrolle_Click()
On Error GoTo Err_op_Adresskontrolle_Click
    Dim tgtdb As Database
    Dim stDocName As String
    Dim stqueryname As String
    Dim stsql As String
    Dim lgEMailTypID As Long
    Dim emailtypenqry As QueryDef
    Dim adrkontrqry As QueryDef
    Dim emailtypenrs As Recordset
    Dim stEmailAbfrage As String
    Set tgtdb = CurrentDb
    'E-Mail Typ: Eigenschaften bestimmen
    stqueryname = "EMailTyp_Selektion"
    lgEMailTypID = Me![EMailTypID]
    stsql = "SELECT a.EMailTypID, " & vbNewLine
    stsql = stsql & "a.AbfragEmailAdressen" & vbNewLine
    stsql = stsql & "FROM EMAILTYPEN AS a" & vbNewLine
    stsql = stsql & "WHERE a.EMailTypID = " & lgEMailTypID & vbNewLine
    stsql = stsql & "and a.AbfragEmailAdressen is not null;"
    Set emailtypenqry = tgtdb.QueryDefs(stqueryname)
    emailtypenqry.SQL = stsql
    Set emailtypenrs = emailtypenqry.OpenRecordset(dbOpenSnapshot)
    If emailtypenrs.EOF Then Err.Raise vbObjectError + 83
    stEmailAbfrage = emailtypenrs!AbfragEmailAdressen
    emailtypenqry.Close
    emailtypenrs.Close
    'Adressabfrage bilden, je nach Email-Typ
    stsql = "SELECT a.PersNr, a.Vorname, a.Nachname, a.Firma, a.Postleitzahl, a.Ort, a.MailPrivat" & vbNewLine
    stsql = stsql + "FROM " & stEmailAbfrage & " AS a " & vbNewLine
    stsql = stsql + "ORDER BY a.Nachname;"
    stDocName = "Adresskontrolle_Emailversand"
    Set adrkontrqry = tgtdb.QueryDefs(stDocName)
    adrkontrqry.SQL = stsql
    adrkontrqry.Close
    DoCmd.OpenQuery stDocName, acNormal, acEdit

Exit_op_Adresskontrolle_Click:
    Exit Sub

Err_op_Adresskontrolle_Click:
    If Err.Number = vbObjectError + 83 Then
        MsgBox "Ungültiger E-Mail Typ " & vbNewLine & _
        "(kein passender Datensatz in EMAILTYPEN gefunden!)", vbCritical
    Else
        MsgBox Err.Description
    End If

    Resume Exit_op_Adresskontrolle_Click
    
End Sub
