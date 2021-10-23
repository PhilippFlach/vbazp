Attribute VB_Name = "Diverse Hilfsfunktionen"
Option Compare Database
Option Explicit
Function Adressname(Firma, Nachname, Vorname, Adresse, Ort)
If IsNull(Ort) Then
Adressname = Trim$(Firma & " " & Nachname & " " & Vorname & " " & Adresse & " " & Ort)
End If
If Not IsNull(Ort) And IsNull(Firma) And Not IsNull(Adresse) Then
Adressname = Trim$(Nachname & " " & Vorname & ", " & Adresse & ", " & Ort)
End If
If Not IsNull(Ort) And Not IsNull(Firma) And Not IsNull(Nachname) And Not IsNull(Adresse) Then
Adressname = Trim$(Firma & ", " & Nachname & " " & Vorname & ", " & Adresse & ", " & Ort)
End If
If Not IsNull(Ort) And Not IsNull(Firma) And IsNull(Nachname) And Not IsNull(Adresse) Then
Adressname = Trim$(Firma & ", " & Adresse & ", " & Ort)
End If
If Not IsNull(Ort) And IsNull(Firma) And IsNull(Adresse) Then
Adressname = Trim$(Nachname & " " & Vorname & ", " & Ort)
End If
If Not IsNull(Ort) And Not IsNull(Firma) And Not IsNull(Nachname) And IsNull(Adresse) Then
Adressname = Trim$(Firma & ", " & Nachname & " " & Vorname & ", " & Ort)
End If
If Not IsNull(Ort) And Not IsNull(Firma) And IsNull(Nachname) And IsNull(Adresse) Then
Adressname = Trim$(Firma & ", " & Ort)
End If

'Wird verwendet in Abfrage Adressbezeichnung und Bestellungen Zeitraum
End Function


Function BestellQuittiert()
BestellQuittiert = 1
'Wird verwendet in Makro Bestell quitt:
'1 = Lieferstatus: "Lieferschein gedruckt"
End Function

Function ersatz(wert, alternativwert)
If IsNull(wert) Then
ersatz = alternativwert
Else
ersatz = wert
End If
'Verwendet in Abfrage Rechnadr und in Abfrage Rechnung mit Rabattangabe
End Function

Function fuenfer(wert)
If IsNull(wert) Or (Not IsNumeric(wert)) Then
fuenfer = 0
Else
fuenfer = CDbl((CLng(wert * 20)) / 20)
End If
End Function

Function fuenfer2(wert)
If IsNull(wert) Or (Not IsNumeric(wert)) Then
fuenfer2 = 0
Else
fuenfer2 = Format$(((CLng(wert * 20)) / 20), "0.00")
End If
End Function

Function InBearb()
InBearb = 0
'0 = Lieferstatus "In Bearbeitung"
End Function


Function naechsterFreitag(aktuellesdatum)
If (6 - WeekDay(aktuellesdatum)) <= 0 Then
naechsterFreitag = aktuellesdatum + (6 - WeekDay(aktuellesdatum)) + 7
Else
naechsterFreitag = aktuellesdatum + (6 - WeekDay(aktuellesdatum))
End If
End Function

Function NameundVorname(Nachname, Vorname)
'Wird verwendet in Abfrage Namen
NameundVorname = Nachname & ", " & Vorname
If IsNull(Vorname) Then
NameundVorname = Nachname
End If
If IsNull(Nachname) Then
NameundVorname = Vorname
End If
End Function


Function RechnungGedruckt()
RechnungGedruckt = 2
'Entspricht Lieferstatus 2: Rechnung gedruckt
'Wird verwendet in Abfrage Rechnung gedruckt
End Function

Function Standardpreiskategorie()
Standardpreiskategorie = 0
'Kategorie, die verwendet wird, falls keine andere Preiskategorie zutrifft
'wird in Abfrage BestelldetailsEPausStandard verwendet
End Function

Function MassgRabatt(GebindeNr, Rabattwert)
'GebindeNr ist leer: Lieferdetail ist Ware
'GebindeNr ist nicht leer: Lieferdetail ist Gebinde

If IsNull(GebindeNr) Then
MassgRabatt = Rabattwert
Else
MassgRabatt = 0
'Wird in Rechnungsbeträge  verwendet
'um zu berechnen, ob es Lieferdetail eine Ware ist
'welche zu Rabatt berechtigt.
End If

End Function
Function mwstber(inkl, MwSt)
If Abs(MwSt) < 0.00001 Then
mwstber = 0
Else
mwstber = inkl / (1 + 1 / MwSt)
'wird in Abfrage Rechnungen Folge verwendet
End If

End Function
Function einleitung(Nr, Text, Vorname, Anrede, Nachname)
If ((Nr = 1) Or (Nr = 2) Or (Nr = 3)) Then
einleitung = Text & " " & Vorname
Else
einleitung = Text & " " & Trim$(Anrede & " " & Nachname)
End If

End Function
Function runden2(wert)
If IsNumeric(wert) Then
runden2 = Format$(wert, "0.00")
Else
runden2 = ""
End If

End Function
Function division(x, y) As Double
If (Abs(y) <= 0.000001) Or (Not IsNumeric(x)) Or (Not IsNumeric(y)) Then
division = 999999
Else
division = x / y
End If

End Function
Function aktuellerauftrag()
If IsLoaded("BESTELL") Then
aktuellerauftrag = [Forms]![Bestell]![BestellNr]
End If
If IsLoaded("LIEFERDETAILS") Then
aktuellerauftrag = [Forms]![Lieferdetails]![BestellNr]
End If
End Function
Function IsLoaded(ByVal strFormName As String) As Boolean
 ' Gibt den Wert "True" zurück, wenn das angegebene Formular in Formularansicht
 ' oder Datenblattansicht geöffnet ist.
    
    Const conObjStateClosed = 0
    Const conDesignView = 0
    
    If SysCmd(acSysCmdGetObjectState, acForm, strFormName) <> conObjStateClosed Then
        If Forms(strFormName).CurrentView <> conDesignView Then
            IsLoaded = True
        End If
    End If
    
End Function

Function aktuellerbereich()
If IsLoaded("BESTELL") Then
aktuellerbereich = [Forms]![Bestell]![BereichNr]
End If
If IsLoaded("LIEFERDETAILS") Then
aktuellerbereich = [Forms]![Lieferdetails]![BereichNr]
End If
End Function
Public Function lieferscheinoeffnen(Bestell As Variant) As Boolean
On Error GoTo Err_Lieferscheinoeffnen
    Dim db As Database
    Dim rs As Recordset
    Dim stquery As String
    Dim stsql As String
    Dim stlieferscheininfo As String
    Dim stqueryname As Variant
    Dim qry As QueryDef
    Dim lgadressnr As Variant
    Dim stlieferinfo As String
    If Nz(Bestell, 0) = 1 Then
        lgadressnr = Forms![Bestell]![AdressNr]
    Else
        lgadressnr = Forms![Lieferdetails]![AdressNr]
    End If
    Set db = CurrentDb
    stqueryname = "Personen_Info_UF_liefer"
    stsql = "SELECT a.AdressNr, " & vbNewLine
    stsql = stsql & "a.Vorname, a.Natel, a.MailPrivat, a.Avisemailflag" & vbNewLine
    stsql = stsql & "FROM Personen_Info_UF AS a" & vbNewLine
    stsql = stsql & "WHERE a.AdressNr = " & lgadressnr & ";"
    Set qry = db.QueryDefs(stqueryname)
    qry.SQL = stsql
    Set rs = qry.OpenRecordset(dbOpenSnapshot)
    stlieferinfo = ""
    While Not rs.EOF
     stlieferinfo = stlieferinfo & rs!Vorname
     stlieferinfo = stlieferinfo & IIf(IsNull(rs!Natel), "", "   " & rs!Natel)
     stlieferinfo = stlieferinfo & IIf(IsNull(rs!MailPrivat), "", "   " & rs!MailPrivat)
     stlieferinfo = stlieferinfo & IIf(IsNull(rs!Avisemailflag), "", "   " & rs!Avisemailflag)
     stlieferinfo = stlieferinfo & vbNewLine
     rs.MoveNext
    Wend
    qry.Close
    rs.Close
    stqueryname = "Logo_privatpersoninfo_update"
    stsql = "UPDATE LOGO " & vbNewLine
    stsql = stsql & "SET LOGO.privatpersoninfo = """ & stlieferinfo & """" & vbNewLine
    Set qry = db.QueryDefs(stqueryname)
    qry.SQL = stsql
    qry.Execute
    qry.Close
    DBEngine.Idle dbRefreshCache
Exit_Lieferscheinoeffnen:
    lieferscheinoeffnen = True
    Exit Function

Err_Lieferscheinoeffnen:
    MsgBox Err.Description
    lieferscheinoeffnen = False
    Resume Exit_Lieferscheinoeffnen

End Function

Function getAnzahlFlaschenDepot(AnzEinh As Variant, HatFlaschenDepot As Boolean) As Long
Dim lgAnzEinh As Long
Dim tmpResult As Long

If IsNumeric(AnzEinh) Then
    lgAnzEinh = Round(AnzEinh, 0)
Else
    lgAnzEinh = 0
End If

If HatFlaschenDepot Then
  tmpResult = lgAnzEinh
Else
  tmpResult = 0
End If

getAnzahlFlaschenDepot = tmpResult

End Function

Function getStandardGebindeName(ArtNr As Variant, PreiskatNr As Variant) As Variant
On Error GoTo Err_getStandardGebindeName

Dim db As Database
Dim qry As QueryDef
Dim rs As Recordset
Dim stqueryname As String
Dim stsql As String
Dim tmpRes As Variant
    If Not (IsNumeric(ArtNr) And IsNumeric(PreiskatNr)) Then
     tmpRes = ""
     GoTo Exit_getStandardGebindeName
    End If

    Set db = CurrentDb
    'E-Mail Typ: Eigenschaften bestimmen
    stqueryname = "Bestelldetails_UF_Preiskatgebinde_sel"
    
    stsql = "SELECT a.ArtBez " & vbNewLine
    stsql = stsql & "FROM Bestelldetails_UF_Preiskatgebinde AS a" & vbNewLine
    stsql = stsql & "WHERE a.ArtNr = " & ArtNr & vbNewLine
    stsql = stsql & "AND a.PreiskatNr = " & PreiskatNr & vbNewLine
    
    Set qry = db.QueryDefs(stqueryname)
    qry.SQL = stsql
    Set rs = qry.OpenRecordset(dbOpenSnapshot)
    If Not rs.EOF Then
       tmpRes = rs!ArtBez
    End If
    rs.Close
    
    If IsNull(tmpRes) Or tmpRes < "0" Then
        stqueryname = "Bestelldetails_UF_Standardgebinde_sel"
        
        stsql = "SELECT a.ArtBez " & vbNewLine
        stsql = stsql & "FROM Bestelldetails_UF_Standardgebinde AS a" & vbNewLine
        stsql = stsql & "WHERE a.ArtNr = " & ArtNr & vbNewLine
        Set qry = db.QueryDefs(stqueryname)
        qry.SQL = stsql
        Set rs = qry.OpenRecordset(dbOpenSnapshot)
        If Not rs.EOF Then
            tmpRes = rs!ArtBez
        End If
    
    End If
    
    
Exit_getStandardGebindeName:
    getStandardGebindeName = tmpRes
    Exit Function

Err_getStandardGebindeName:
    MsgBox Err.Description
    MsgBox Err.Number
    Resume Exit_getStandardGebindeName
  
End Function


Function getStandardGebindeCode(ArtNr As Variant, PreiskatNr As Variant) As Long
On Error GoTo Err_getStandardGebindeCode

Dim db As Database
Dim qry As QueryDef
Dim rs As Recordset
Dim stqueryname As String
Dim stsql As String
Dim tmpRes As Variant
    If Not (IsNumeric(ArtNr) And IsNumeric(PreiskatNr)) Then
     tmpRes = 0
     GoTo Exit_getStandardGebindeCode
    End If

    Set db = CurrentDb
    'E-Mail Typ: Eigenschaften bestimmen
    stqueryname = "Bestelldetails_UF_Preiskatgebinde_sel"
    
    stsql = "SELECT a.Gebindecode " & vbNewLine
    stsql = stsql & "FROM Bestelldetails_UF_Preiskatgebinde AS a" & vbNewLine
    stsql = stsql & "WHERE a.ArtNr = " & ArtNr & vbNewLine
    stsql = stsql & "AND a.PreiskatNr = " & PreiskatNr & vbNewLine
    
    Set qry = db.QueryDefs(stqueryname)
    qry.SQL = stsql
    Set rs = qry.OpenRecordset(dbOpenSnapshot)
    If Not rs.EOF Then
       tmpRes = rs!Gebindecode
    End If
    rs.Close
    
    If IsNull(tmpRes) Or tmpRes < "0" Then
        stqueryname = "Bestelldetails_UF_Standardgebinde_sel"
        
        stsql = "SELECT a.Gebindecode " & vbNewLine
        stsql = stsql & "FROM Bestelldetails_UF_Standardgebinde AS a" & vbNewLine
        stsql = stsql & "WHERE a.ArtNr = " & ArtNr & vbNewLine
        Set qry = db.QueryDefs(stqueryname)
        qry.SQL = stsql
        Set rs = qry.OpenRecordset(dbOpenSnapshot)
        If Not rs.EOF Then
            tmpRes = rs!Gebindecode
        End If
    
    End If
    
    
Exit_getStandardGebindeCode:
    getStandardGebindeCode = tmpRes
    Exit Function

Err_getStandardGebindeCode:
    MsgBox Err.Description
    MsgBox Err.Number
    Resume Exit_getStandardGebindeCode
  
End Function
Function getAddressLine(Address As Variant) As String
On Error GoTo Err_getAddressLine
Dim lgPositionCrLf1 As Long
Dim lgPositionCrLf2 As Long
Dim lgAddressLength As Long
Dim stringCompare As String

Dim tmpRes As String
    
    tmpRes = ""
    If IsNull(Address) Then
        tmpRes = ""
        GoTo Exit_getAddressLine
    End If
    
    'Suche Zeilenumbruch
    stringCompare = vbCrLf
    lgPositionCrLf1 = InStr(1, Address, stringCompare, vbBinaryCompare)
    
    'Hilfsgrösse: Länge der gesamten Adresse
    lgAddressLength = Len(Address)
    
    If lgPositionCrLf1 = 0 Then
        'Enthält nur eine Adresszeile
        'Gib erste Adresszeile zurück
        tmpRes = Address
    Else
        'Enthält zwei Adresszeilen
        'Gib zweite Adresszeile zurück
        'Feststellen wo die 2. Adresszeile aufhoert
         lgPositionCrLf2 = InStr(lgPositionCrLf1 + 2, Address, stringCompare, vbBinaryCompare)
        If lgPositionCrLf2 = 0 Then
          'Es hat nur 2 Zeilen im Adresse Feld, die 2. Zeile zurückgeben
          tmpRes = Mid(Address, (lgPositionCrLf1 + 2), (lgAddressLength - lgPositionCrLf1))
        Else
          'Es hat mehr als 2 Adresszeilen
          'Nur die 2. Zeile zurückgeben
           tmpRes = Mid(Address, (lgPositionCrLf1 + 2), (lgPositionCrLf2 - lgPositionCrLf1 - 2))
        End If
    End If

Exit_getAddressLine:
    getAddressLine = tmpRes
    Exit Function

Err_getAddressLine:
    If Err.Number = vbObjectError + 93 Then
        MsgBox "Address-Zeilen-Nr. muss 1 oder 2 sein"
    Else
        MsgBox Err.Description
    End If
    Resume Exit_getAddressLine

End Function

