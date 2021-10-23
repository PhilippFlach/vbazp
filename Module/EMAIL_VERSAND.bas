Attribute VB_Name = "EMAIL_VERSAND"
Option Compare Database

Option Explicit
Public Const conEmailEinzelversand As String = "Einzelversand"
Public Const conEmailMassenversand  As String = "Massenversand"
Public Const conDefaultNullZahl As Long = -1
Public Const conDefaultNullString As String = ""
Public Const conFehlendeObligatorischeFelder As Integer = 9900
Public Const conFeldRebenartVorhanden As Integer = 11
Public Const conFeldRebeNrVorhanden As Integer = 12
Public Const conFeldRebenartRebeNrFehlt As Integer = 1112
Public Const conFeldInternetLinkFehlt As Integer = 13
Public Const conFeldRebeNrInternetLinkVorhanden As Integer = 1113
Public Const conFeldRebenartInternetLinkVorhanden As Integer = 1213
Public Const conFeldRebenartRebeNrUndInternetLinkFehlen As Integer = 11213 'default
Public Const conFelderAlleVorhanden = 0

Function getEmailNachricht(rsemtxt As Recordset, date1Datum As Date, st2Vorname As String, st3Nachname As String, _
st4Adresse As String, st5PLZ As String, st6Ort As String, _
st7email As String, st8Lieb As String, st9Anrede As String, _
st11Rebenart As String, _
lg12RebeNr As Long, st13Internetlink As String, st14Firma As String) As String
'rsemtxt fragt gefilterte Datensätze aus Tabelle EMAILTEXTE ab
'Erwartete Variablen in rsemtxt:
'  !Emailtext
'  !LeerschlagAnzahl
'  !SeriendruckfeldID
'  !ZeilenumbruchAnzahl

On Error GoTo Err_getEmailNachricht
    Dim varEmail As String 'temporärer E-mail Text
    Dim lgcounter As Long 'Zählvariable
    varEmail = conDefaultNullString
    rsemtxt.MoveFirst
    Do While rsemtxt.EOF <> True
        varEmail = varEmail & rsemtxt!emailtext
        lgcounter = 0
        Do While lgcounter < rsemtxt!LeerschlagAnzahl
            varEmail = varEmail & " "
            lgcounter = lgcounter + 1
        Loop
        Select Case rsemtxt!SeriendruckfeldID
        Case 1
          varEmail = varEmail & Format(date1Datum)
        Case 2
          varEmail = varEmail & st2Vorname
        Case 3
          varEmail = varEmail & st3Nachname
        Case 4
          varEmail = varEmail & st4Adresse
        Case 5
          varEmail = varEmail & st5PLZ & " " & st6Ort
        Case 6
          varEmail = varEmail & st6Ort
        Case 7
          varEmail = varEmail & st7email
        Case 8
          varEmail = varEmail & st8Lieb
        Case 9
          varEmail = varEmail & st9Anrede
        Case 11
          varEmail = varEmail & st11Rebenart
        Case 12
          If Not (lg12RebeNr = conDefaultNullZahl) Then
            varEmail = varEmail & Format(lg12RebeNr)
          Else
            varEmail = varEmail
          End If
        Case 13
            varEmail = varEmail & st13Internetlink
        Case 14
            varEmail = varEmail & st14Firma
        Case Else
          varEmail = varEmail
        End Select
        lgcounter = 0
        Do While lgcounter < rsemtxt!ZeilenumbruchAnzahl
            varEmail = varEmail & vbNewLine
            lgcounter = lgcounter + 1
        Loop
        rsemtxt.MoveNext
    Loop


Exit_getEmailNachricht:
    getEmailNachricht = varEmail
    Exit Function

Err_getEmailNachricht:
    MsgBox Err.Description
    MsgBox Err.Number
    Resume Exit_getEmailNachricht
    
End Function

Public Sub mailversand(sttoEmailAdresse As String, stsubject As String, varEmailNachricht As String, Optional stAttachment As String)
On Error GoTo Err_mailversand
   If sttoEmailAdresse = conDefaultNullString Then Err.Raise vbObjectError + 81
   If InStr(1, sttoEmailAdresse, "@") = 0 Then Err.Raise vbObjectError + 81
   If stAttachment = conDefaultNullString Then
    DoCmd.SendObject acSendNoObject, , , sttoEmailAdresse, , , stsubject, varEmailNachricht, False
   Else
    DoCmd.SendObject acSendNoObject, , , sttoEmailAdresse, , , stsubject, varEmailNachricht, True, stAttachment
   End If
Exit_mailversand:
    Exit Sub

Err_mailversand:
    If Err.Number = vbObjectError + 81 Then
       MsgBox "Ungültige E-mail-Adresse"
    Else
        MsgBox Err.Description
    End If
    Resume Exit_mailversand
    
End Sub

Function checkAddressFields(tgtrs As Recordset) As Integer
'TRUE = In tgtrs sind die erforderlichen Felder vorhanden fragt gefilterte Datensätze aus Tabelle EMAILTEXTE ab
'FALSE = Es fehlt ein Feld in tgtrs, welches für den E-mail versand benötigt wird

On Error GoTo Err_checkAddressFields

    Dim inttemp As Integer 'temporärer Rückgabewert
    
    Dim fldLoop As Field
    Dim stFeldname As String
    
    Dim boolAnrede As Boolean
    Dim boolVorname As Boolean
    Dim boolNachname As Boolean
    Dim boolAdresse As Boolean
    Dim boolPostleitzahl As Boolean
    Dim boolOrt As Boolean
    Dim boolMailPrivat As Boolean
    Dim boolLieb As Boolean
    Dim boolRebenart As Boolean
    Dim boolRebeNr As Boolean
    Dim boolInternetlink As Boolean
    
    'Initialisierung:
    'Es wurde noch kein Feld gefunden
    'Daher alle Variablen auf False
    boolAnrede = False
    boolVorname = False
    boolNachname = False
    boolAdresse = False
    boolPostleitzahl = False
    boolOrt = False
    boolMailPrivat = False
    boolLieb = False
    boolRebenart = False
    boolRebeNr = False
    boolInternetlink = False
    inttemp = conFeldRebenartRebeNrUndInternetLinkFehlen
    
    For Each fldLoop In tgtrs.Fields
    'Prüfen welche Feldnamen in der Abfrage vorkommen
        stFeldname = fldLoop.Name
        If stFeldname = "Anrede" Then boolAnrede = True
        If stFeldname = "Vorname" Then boolVorname = True
        If stFeldname = "Nachname" Then boolNachname = True
        If stFeldname = "Adresse" Then boolAdresse = True
        If stFeldname = "Postleitzahl" Then boolPostleitzahl = True
        If stFeldname = "Ort" Then boolOrt = True
        If stFeldname = "MailPrivat" Then boolMailPrivat = True
        If stFeldname = "Lieb" Then boolLieb = True
        If stFeldname = "Rebenart" Then boolRebenart = True
        If stFeldname = "RebeNr" Then boolRebeNr = True
        If stFeldname = "Internetlink" Then boolInternetlink = True
    Next fldLoop

    'Prüfen ob obligatorisches Feld fehlt
   If Not (boolAnrede And _
    boolVorname And _
    boolNachname And _
    boolAdresse And _
    boolPostleitzahl And _
    boolOrt And _
    boolMailPrivat And _
    boolLieb) Then Err.Raise vbObjectError + 90 'obligatorisches Feld fehlt
    
    'Prüfen ob fakultative Felder fehlen
    If (Not boolRebenart) And boolRebeNr And (Not boolInternetlink) Then inttemp = conFeldRebeNrVorhanden
    If boolRebenart And (Not boolRebeNr) And (Not boolInternetlink) Then inttemp = conFeldRebenartVorhanden
    If (Not boolRebenart) And boolRebeNr And boolInternetlink Then inttemp = conFeldRebeNrInternetLinkVorhanden
    If boolRebenart And (Not boolRebeNr) And boolInternetlink Then inttemp = conFeldRebenartInternetLinkVorhanden
    If (Not boolRebenart) And (Not boolRebeNr) And boolInternetlink Then inttemp = conFeldRebenartRebeNrFehlt  'Beide Felder fehlen
    If boolRebenart And boolRebeNr And boolInternetlink Then inttemp = conFelderAlleVorhanden   'Alle Felder vorhanden
    If boolRebenart And boolRebeNr And Not (boolInternetlink) Then inttemp = conFeldInternetLinkFehlt
Exit_checkAddressFields:
    checkAddressFields = inttemp
    Exit Function

Err_checkAddressFields:
    If Err.Number = vbObjectError + 90 Then
        If Not boolAnrede Then MsgBox "Das Feld Anrede fehlt in der Abfrage, welche in EMAILTYPEN Tabelle definiert ist."
        If Not boolVorname Then MsgBox "Das Feld Vorname fehlt in der Abfrage, welche in EMAILTYPEN Tabelle definiert ist."
        If Not boolNachname Then MsgBox "Das Feld Nachname fehlt in der Abfrage, welche in EMAILTYPEN Tabelle definiert ist."
        If Not boolAdresse Then MsgBox "Das Feld Adresse fehlt in der Abfrage, welche in EMAILTYPEN Tabelle definiert ist."
        If Not boolPostleitzahl Then MsgBox "Das Feld Postleitzahl fehlt in der Abfrage, welche in EMAILTYPEN Tabelle definiert ist."
        If Not boolOrt Then MsgBox "Das Feld Ort fehlt in der Abfrage, welche in EMAILTYPEN Tabelle definiert ist."
        If Not boolMailPrivat Then MsgBox "Das Feld MailPrivat fehlt in der Abfrage, welche in EMAILTYPEN Tabelle definiert ist."
        If Not boolLieb Then MsgBox "Das Feld Lieb fehlt in der Abfrage, welche in EMAILTYPEN Tabelle definiert ist."
    Else
        MsgBox Err.Description
        MsgBox Err.Number
    End If
        inttemp = conFehlendeObligatorischeFelder
    Resume Exit_checkAddressFields
    
End Function

'Public Function SendeMail(strMail As String, strBetreff As String, _
'                          strText As String, strAttach As String) As Boolean
''strMail enthält alle Mailadressen (wie man sie in Outloook eingeben würde)
''strBetreff die Betreffzeile für die Mail
''strText den eMail-Text
'    Static myMail       As Outlook.MailItem
'    Static myOutlApp    As Outlook.Application
'    Dim iAttach         As Integer
'
'    SendeMail = False
'    'Outlook öffnen
'    Set myOutlApp = New Outlook.Application
'    'neue Mail öffnen
'    Set myMail = myOutlApp.CreateItem(olMailItem)
'    With myMail
'        'Mail Felder befüllen
'        .To = strMail
'        .Subject = strBetreff
'        .Body = strText
'        'Dateien Anhängen
'        .Attachments.Add strAttach
'        'Abschicken (ich gebe zu, dass man auftretende Fehler auch eleganter abfangen kann ;-)
'        On Error Resume Next
'        .Send
'        If Err.Number <> 0 Then SendeMail = False
'        On Error GoTo 0
'    End With
'    myOutlApp.Quit
'    Set myMail = Nothing
'    Set myOutlApp = Nothing
'End Function


Function getPreiskatGruppe(lgPreiskatNr As Long) As String
'fragt Preiskategorie-Gruppe aus Tabelle PREISKATEGORIEN ab
'Input: Gewählte Preiskategorien-Nummer

On Error GoTo Err_getPreiskatGruppe

    Dim db As Database
    Dim stsql As String
    Dim stqueryname As String
    Dim rs As Recordset
    Set db = CurrentDb
    Dim qry As QueryDef
    Dim stResult As String
    
    'Preis-Kategorie: Gruppe bestimmen
    stqueryname = "Preiskategorie_lookup_gruppe"
    stsql = "SELECT a.PreiskatNr, a.PreiskatGruppe" & vbNewLine
    stsql = stsql & "FROM PREISKATEGORIEN AS a" & vbNewLine
    stsql = stsql & "WHERE a.PreiskatNr = " & lgPreiskatNr & ";"
    Set qry = db.QueryDefs(stqueryname)
    qry.SQL = stsql
    Set rs = qry.OpenRecordset(dbOpenSnapshot)
    If rs.EOF Then
        stResult = ""
    Else
        stResult = rs!PreiskatGruppe
    End If
    qry.Close
    rs.Close
Exit_getPreiskatGruppe:
    getPreiskatGruppe = stResult
    Exit Function

Err_getPreiskatGruppe:
    MsgBox Err.Description
    MsgBox Err.Number
    Resume Exit_getPreiskatGruppe
    
End Function


