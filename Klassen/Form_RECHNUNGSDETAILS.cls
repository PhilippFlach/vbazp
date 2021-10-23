VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_RECHNUNGSDETAILS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Function boolRabattUpdate() As Boolean
On Error GoTo Err_boolRabattUpdate
    Dim db As Database
    Dim qry As QueryDef
    Dim stsql As String
    Dim stqueryname As String
    Dim lgRechn As Long
    Dim varSpezialrabatt As Variant
    Dim varRabattStandard As Variant
    Dim varRabattTatsaechlich As Variant
    Dim boolresult As Boolean
    
    'Initialisierung
    lgRechn = Forms![RECHNUNGSDETAILS]![RechNr]
    varSpezialrabatt = Forms![RECHNUNGSDETAILS]![RechnRabatt]
    varRabattStandard = Forms![RECHNUNGSDETAILS]![RabattUF].Form![Rabatt]
    boolresult = False
    
    If IsNull(varSpezialrabatt) Then
    'Falls kein Spezialrabatt gilt Standard
        varRabattTatsaechlich = varRabattStandard
    Else
    'Spezialrabatt anwenden
        varRabattTatsaechlich = varSpezialrabatt
    End If
    'Prüfen ob ein Rabatt vorliegt
    If IsNumeric(varRabattTatsaechlich) Then
        'Rabatt übertragen
        Set db = CurrentDb
        stqueryname = "Rechnung RechnRabatt update"
        stsql = "UPDATE RECHNUNGEN AS a SET a.RechnRabatt = " & varRabattTatsaechlich & vbNewLine
        stsql = stsql & "WHERE a.RechNr = " & lgRechn
        Set qry = db.QueryDefs(stqueryname)
        qry.SQL = stsql
        qry.Execute
        qry.Close
    End If
    boolresult = True
    
Exit_boolRabattUpdate:
    boolRabattUpdate = boolresult
    Exit Function

Err_boolRabattUpdate:
    MsgBox Err.Description
    boolresult = False
    Resume Exit_boolRabattUpdate

End Function

Private Sub Befehl24_Click()

End Sub

Private Sub cmdOpenRechnSC_Click()
On Error GoTo Err_cmdOpenRechnSC_Click
    Dim strQRCode As String
    Dim HatRabattUpdateErfolgreich As Boolean
    Dim HatRechnungsdatenAufbereitet As Boolean
    Dim stRechnNr As String
    'Zuerst Daten im Formular speichern
    DoCmd.RunMacro "Speichern"
    'Rabatt festlegen
    HatRabattUpdateErfolgreich = boolRabattUpdate()
    'Rechnung drucken
    If Not HatRabattUpdateErfolgreich Then Err.Raise vbObjectError + 57
    HatRechnungsdatenAufbereitet = SetupRechnungSwisscode()
    If Not HatRechnungsdatenAufbereitet Then Err.Raise vbObjectError + 58
    fDebug
    strQRCode = fQRString
    Call GetQRCode2(strQRCode, 543, 543)                 'ausgelegt für 46mm
    ' hier kommt der Druckvorgang
    DoCmd.OpenReport "Rechnung Kopf SC", acViewPreview, , , acWindowNormal
    stRechnNr = Format(Forms![RECHNUNGSDETAILS]![RechNr])
    Reports![Rechnung Kopf SC].Caption = stRechnNr & "_QR_Rechnung"
Exit_cmdOpenRechnSC_Click:
    Exit Sub

Err_cmdOpenRechnSC_Click:
    If Err.Number = vbObjectError + 57 Then
        MsgBox "Berechnung Rabatt fehlgeschlagen"
    ElseIf Err.Number = vbObjectError + 58 Then
        MsgBox "Aufbereitung Rechnungsdaten fehlgeschlagen"
    Else
        MsgBox Err.Description
    End If
    Resume Exit_cmdOpenRechnSC_Click
End Sub

Private Sub cmDoppelSC_Click()
On Error GoTo Err_cmDoppelSC_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim stRechnNr As String

    stDocName = "Rechnung Doppel SC"
        
    DoCmd.OpenReport stDocName, acViewPreview
    stRechnNr = Format(Forms![RECHNUNGSDETAILS]![RechNr])
    Reports![Rechnung Doppel SC].Caption = stRechnNr & "_Doppel"
Exit_cmDoppelSC_Click:
    Exit Sub

Err_cmDoppelSC_Click:
    MsgBox Err.Description
    Resume Exit_cmDoppelSC_Click
End Sub

Private Sub Gastroinfo_Click()
On Error GoTo Err_Gastroinfo_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "GASTRO"
    
    stLinkCriteria = "[AdressNr]=" & Me![AdressNr2]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Gastroinfo_Click:
    Exit Sub

Err_Gastroinfo_Click:
    MsgBox Err.Description
    Resume Exit_Gastroinfo_Click
    
End Sub
Private Sub RechnBetrOeffn_Click()
On Error GoTo Err_RechnBetrOeffn_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "RechnBetr"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_RechnBetrOeffn_Click:
    Exit Sub

Err_RechnBetrOeffn_Click:
    MsgBox Err.Description
    Resume Exit_RechnBetrOeffn_Click
    
End Sub
Private Sub Befehl33_Click()
On Error GoTo Err_Befehl33_Click

    Dim stDocName As String

    stDocName = "Rechnung Proforma GG"
    DoCmd.OpenReport stDocName, acPreview

Exit_Befehl33_Click:
    Exit Sub

Err_Befehl33_Click:
    MsgBox Err.Description
    Resume Exit_Befehl33_Click
    
End Sub
Private Sub Befehl35_Click()
On Error GoTo Err_Befehl35_Click

    Dim stDocName As String

    stDocName = "RechnungAdressePacket"
    DoCmd.OpenReport stDocName, acPreview

Exit_Befehl35_Click:
    Exit Sub

Err_Befehl35_Click:
    MsgBox Err.Description
    Resume Exit_Befehl35_Click
    
End Sub

Private Sub Rechnungdrucken_Click()
  Dim boolErfolgreich As Boolean
  'Zuerst Daten im Formular speichern
  DoCmd.RunMacro "Speichern"
  'Rabatt festlegen
  boolErfolgreich = boolRabattUpdate()
  'Rechnung drucken
  If boolErfolgreich Then
    DoCmd.RunMacro "Rechnung drucken"
  End If
End Sub

Private Sub RechnungProFormadrucken_Click()
  Dim boolErfolgreich As Boolean
  'Zuerst Daten im Formular speichern
  DoCmd.RunMacro "Speichern"
  'Rabatt festlegen
  boolErfolgreich = boolRabattUpdate()
  'Rechnung drucken
  If boolErfolgreich Then
    DoCmd.RunMacro "Rechnung drucken Proforma"
  End If

End Sub
