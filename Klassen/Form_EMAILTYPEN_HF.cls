VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EMAILTYPEN_HF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub EMailVorschau_Click()
On Error GoTo Err_EMailVorschau_Click

    Dim stDocName As String
    Dim stqueryname As String
    Dim stsql As String
    Dim emailtypenqry As QueryDef
    Dim emailtypenrs As Recordset
    Dim stEMailBetreff As String
    Dim stEMailNachricht As String
    Dim tgtdb As Database
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
    stInternetlink = "www.zizerser.ch/pdf/preisliste.pdf"
    stFirma = "Musterfirma"
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
