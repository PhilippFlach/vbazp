Attribute VB_Name = "SwissCodeIntegration"
Option Compare Database

Option Explicit


Public Function SetupRechnungSwisscode() As Boolean
   On Error GoTo Err_SetupRechnungSwisscode
    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim db As Database
    Dim qry As QueryDef
    Dim rs As Recordset
    Dim stsql As String
    Dim result As Boolean
    
    Dim LProdukte2 As String
    Dim LArray() As String
    Dim Zahlungsempf As String
    Dim WeingutImTschalaer As String
    Dim Rechnungsbetrag As String
    Dim ZahlerPLZOrt As String
    Dim UnstruktInfo As String
    Dim stIBAN As String
    Dim tv As TempVar

    
    'Initialisierungen
    result = True
    Set db = CurrentDb
    Set qry = db.QueryDefs("Rechnung Kopf")
    qry.Parameters(0) = Forms![RECHNUNGSDETAILS]![RechNr]
    
    Set rs = qry.OpenRecordset(dbOpenSnapshot)
    If rs.EOF Then Err.Raise vbObjectError + 121
    TempVars.RemoveAll
    With rs
        LArray() = Split(![Produkte2].Value, " ")
        Zahlungsempf = StrConv(LArray(0), vbProperCase)
        TempVars![ZEName] = Zahlungsempf
        WeingutImTschalaer = Trim(StrConv(LArray(1), vbProperCase) & " " & ![FamilieG].Value)
        TempVars![ZEStrtNmOrAdrLine1] = WeingutImTschalaer
        TempVars![ZEBldgNbOrAdrLine2] = ![ZizersG].Value
        Rechnungsbetrag = ![TotalFr].Value & "." & Format(![TotalRp].Value, "00")
        TempVars![Amt] = Rechnungsbetrag
        TempVars![ZPName] = ![Zeile1].Value
        TempVars![ZPStrtNmOrAdrLine1] = ![AdressZeile1].Value
        ZahlerPLZOrt = ![Postleitzahl].Value & " " & ![Ort].Value
        'TempVars![ZPBldgNbOrAdrLine2] = ![AdressZeile2].Value
        TempVars![ZPBldgNbOrAdrLine2] = ZahlerPLZOrt
        'TempVars![ZPPstCd] = CStr(![Postleitzahl].Value)
        'TempVars![ZPTwnNm] = ![Ort].Value
        UnstruktInfo = "Rgdat: " & ![RechDatum].Value & " Rgnr: " & ![RechNr].Value
        TempVars![Ustrd] = UnstruktInfo

    End With
    
    'Statische Elemente
    TempVars![QRTyp] = "SPC"
    TempVars![Version] = "0200"
    TempVars![Coding] = "1"
    stIBAN = Forms![RECHNUNGSDETAILS]![IBAN].Value
    TempVars![ZEIBAN] = stIBAN
    TempVars![ZEAdrTp] = "K"
    TempVars![ZEPstCd] = ""
    TempVars![ZETwnNm] = ""
    TempVars![ZECtry] = "CH"
    TempVars![Ccy] = "CHF"
    TempVars![ZPAdrTp] = "K"
    TempVars![ZPPstCd] = ""
    TempVars![ZPTwnNm] = ""
    TempVars![ZPCtry] = "CH"
    TempVars![Tp] = "NON"
    TempVars![Ref] = ""
    TempVars![AddInf] = ""
    TempVars![Trailer] = "EPD"
    TempVars![StrdBkgInf] = ""
    TempVars![AltPmt1] = ""
    TempVars![AltPmt2] = ""
    TempVars!Sprache = 1
    
    'Only debugging
'    Debug.Print "TempVars.Count: " & TempVars.Count
'    For Each tv In TempVars
'        Debug.Print tv.Name, tv.Value
'    Next

Exit_SetupRechnungSwisscode:
    SetupRechnungSwisscode = result

   Exit Function

Err_SetupRechnungSwisscode:
   'Check for any errors you might be expecting
   If Err.Number = vbObjectError + 121 Then
      MsgBox "Keine Daten für Rechnung Kopf", vbCritical, "Fehler in Abfrage Rechnung Kopf"
      result = False
      Resume Exit_SetupRechnungSwisscode
   Else
      'Give the user some info on this error
      MsgBox Err.Description
      result = False
      Resume Exit_SetupRechnungSwisscode
   End If
End Function

Public Sub GetQRCode2(Content As String, Width As Integer, Height As Integer)
Dim ByteData() As Byte
Dim XmlHttp As Object
Dim HttpReq As String
Dim ReturnContent As String
Dim EncContent As String
Dim QRImage As String
          
          EncContent = EncodeURL2(Content)
          
          HttpReq = "https://api.qrserver.com/v1/create-qr-code/?data=" & EncContent & "&size=" & Width & "x" & Height & "&ecc=M &charset-target=UTF8 &format=png"
          'HttpReq = "https://www.qrcode-monkey.com, size:250, download:true"
          'HttpReq = "https://qrcode.tec-it.com/API/QRCode?data=" & EncContent & " & errorcorrection:M"
          'HttpReq = "https://chart.googleapis.com/chart?cht=qr, chs=<250>x<250>, chl=" & EncContent & ", chld=M"
          'Debug.Print EncContent
          'Debug.Print HttpReq
          Set XmlHttp = CreateObject("MSXML2.XmlHttp")
          XmlHttp.Open "GET", HttpReq, False
          XmlHttp.Send
          ByteData = XmlHttp.responseBody
          Set XmlHttp = Nothing
          
          ReturnContent = StrConv(ByteData, vbUnicode)
          Call ExportImage2(ReturnContent)

End Sub
Public Function EncodeURL2(Str As String)
Dim ScriptEngine As Object
Dim encoded As String
Dim Temp As String

          Temp = Replace(Str, " ", "%20")                             'ersetzt Leerzeichen
          Temp = Replace(Temp, "#", "%23")                     'ersetzt #
          Temp = Replace(Temp, vbCrLf, "%0a")              'ersetzt den Zeilenumbruch
          EncodeURL2 = Temp
          
End Function
Public Sub ExportImage2(image As String)
Dim FilePath As String

On Error GoTo NoSave

          FilePath = Application.CurrentProject.Path & "\qr.png"
          
          Open FilePath For Binary As #1
          Put #1, 1, image
          Close #1
          
Exit Sub

NoSave:
          MsgBox "Could not save the QR Code Image! Reason: " & Err.Description, vbCritical, "File Save Error"

End Sub
