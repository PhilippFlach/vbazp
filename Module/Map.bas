Attribute VB_Name = "Map"
Option Compare Database

Option Explicit
Public Enum MapTyp
    Gastro = 1
    Wiederverkauf = 2
End Enum

Public Sub exportMap(exportMapTyp As MapTyp)
On Error GoTo Err_exportMap

'    Dim stDocName As String
'    Dim stLinkCriteria As String
'    Dim queryname As String
'    Dim stFileName As String
'
'    Dim db As Database
'
'    Dim sttest As String
'    If exportMapTyp = Gastro Then
'        stJSON = "x"
'    ElseIf exportMapTyp = Wiederverkauf Then
'        stJSON = "x"
'    Else
'       MsgBox "Ungültiges Argumnent in exportMap"
'       GoTo Exit_exportMap
'    End If
'

Exit_exportMap:
    Exit Sub

Err_exportMap:
    MsgBox Err.Description
    Resume Exit_exportMap
    
End Sub


Public Sub SchreibeMapdaten(rsMapInfo As Recordset, stFileName As String)


On Error GoTo Err_SchreibeMapdaten
Dim DNr As Integer 'Dateinummer für Output
Dim Dateiname As String
Dim zaehler  As Long
Dim MaxAnzahlZeilen As Long
Dim stIdent As String 'Linker Einzug
'ob zusätzlich noch Rezeptkopf geschrieben werden muss
Dim boolErsterEintrag As Boolean 'True = Der aktuelle Record ist der erste Eintrag

Dim stQuoteMarks As String 'Wird benötigt für JSON Textbegrenzung

'Variablen initialisieren
stQuoteMarks = Chr(34)
MaxAnzahlZeilen = 20000 'Maximale Anzahl Records
stIdent = "    "
DNr = FreeFile
Dateiname = stFileName & ".temp.json"
If rsMapInfo.EOF Then Err.Raise vbObjectError + 42 'keine Daten
'Datei  initialisieren
Open Dateiname For Output As #DNr    ' Datei zur Ausgabe öffnen.

'Initialisierungen
boolErsterEintrag = True
zaehler = 0
Print #DNr, "["
'Loop über alle Adressenn:
'Markerdaten in Datei schreiben
Do While Not rsMapInfo.EOF
   If boolErsterEintrag Then
        Print #DNr, "  {"
        boolErsterEintrag = False
   Else 'zweiter oder folgende Marker
        Print #DNr, "  },"
        Print #DNr, "  {"
   End If
        Print #DNr, stIdent & JSONName("id") & stQuoteMarks & rsMapInfo!GINR & stQuoteMarks & ","
        Print #DNr, stIdent & JSONName("map_id") & stQuoteMarks & "1" & stQuoteMarks & ","
        Print #DNr, stIdent & JSONName("address") & stQuoteMarks & json_Encode(rsMapInfo!WebAdresse) & stQuoteMarks & ","
        Print #DNr, stIdent & JSONName("description") & stQuoteMarks & json_Encode(rsMapInfo!WebTextProdukte) & stQuoteMarks & ","
        Print #DNr, stIdent & JSONName("pic") & stQuoteMarks & stQuoteMarks & ","
    If IsNull(rsMapInfo!Internet) Then
         Print #DNr, stIdent & JSONName("link") & stQuoteMarks & stQuoteMarks & ","
    ElseIf LCase(Left(rsMapInfo!Internet, 4)) = "http" Then
        Print #DNr, stIdent & JSONName("link") & stQuoteMarks & rsMapInfo!Internet & stQuoteMarks & ","
    Else
         Print #DNr, stIdent & JSONName("link") & stQuoteMarks & "http://" & rsMapInfo!Internet & stQuoteMarks & ","
    End If
        Print #DNr, stIdent & JSONName("icon") & stQuoteMarks & stQuoteMarks & ","
   If IsNull(rsMapInfo!WebLat) Then
        Print #DNr, stIdent & JSONName("lat") & stQuoteMarks & stQuoteMarks & ","
   Else
        Print #DNr, stIdent & JSONName("lat") & stQuoteMarks & rsMapInfo!WebLat & stQuoteMarks & ","
   End If
   If IsNull(rsMapInfo!WebIng) Then
        Print #DNr, stIdent & JSONName("lng") & stQuoteMarks & stQuoteMarks & ","
   Else
        Print #DNr, stIdent & JSONName("lng") & stQuoteMarks & rsMapInfo!WebIng & stQuoteMarks & ","
   End If
        Print #DNr, stIdent & JSONName("anim") & stQuoteMarks & "0" & stQuoteMarks & ","
        Print #DNr, stIdent & JSONName("title") & stQuoteMarks & json_Encode(rsMapInfo!Firma) & stQuoteMarks & ","
        Print #DNr, stIdent & JSONName("infoopen") & stQuoteMarks & "0" & stQuoteMarks & ","
        Print #DNr, stIdent & JSONName("category") & stQuoteMarks & stQuoteMarks & ","
        Print #DNr, stIdent & JSONName("approved") & stQuoteMarks & "1" & stQuoteMarks & ","
        Print #DNr, stIdent & JSONName("retina") & stQuoteMarks & "1" & stQuoteMarks & ","
        Print #DNr, stIdent & JSONName("type") & stQuoteMarks & "0" & stQuoteMarks & ","
        Print #DNr, stIdent & JSONName("did") & stQuoteMarks & stQuoteMarks & ","
        Print #DNr, stIdent & JSONName("sticky") & stQuoteMarks & "0" & stQuoteMarks & ","
        Print #DNr, stIdent & JSONName("other_data") & stQuoteMarks & stQuoteMarks

    zaehler = zaehler + 1
    If zaehler > MaxAnzahlZeilen Then Err.Raise vbObjectError + 53 'Zuviele Records
    
    rsMapInfo.MoveNext
Loop
Print #DNr, "  }"
Print #DNr, "]"
Close #DNr    ' Datei schließen.
MsgBox "Access hat " & zaehler & " Orte in die Datei " & vbNewLine _
        & Dateiname & vbNewLine & "geschrieben! ", vbInformation, "Exportvorgang erfolgreich"
        
Exit_SchreibeMapdaten:
    Exit Sub

Err_SchreibeMapdaten:
    If Err.Number = vbObjectError + 53 Then
        MsgBox ("Mehr als " & MaxAnzahlZeilen & " Daten für den Export" _
        & " gefunden! ")
        Close #DNr    ' Datei schließen.
    ElseIf Err.Number = vbObjectError + 42 Then
        MsgBox ("Keine Daten für " _
        & "Export-Datei gefunden! ")
    Else
        MsgBox Err.Description
    End If
    Resume Exit_SchreibeMapdaten
End Sub

Public Function JSONName(Name As String)
    Dim stQuoteMarks As String 'Wird benötigt für JSON Textbegrenzung

    'Variablen initialisieren
    stQuoteMarks = Chr(34)
    JSONName = stQuoteMarks & Name & stQuoteMarks & ": "

End Function
Public Function json_Encode(ByVal json_Text As Variant) As String
    ' Reference: http://www.ietf.org/rfc/rfc4627.txt
    ' Escape: ", \, /, backspace, form feed, line feed, carriage return, tab
    Dim json_Index As Long
    Dim json_Char As String
    Dim json_AscCode As Long
    Dim json_Buffer As String
    Dim json_BufferPosition As Long
    Dim json_BufferLength As Long

    For json_Index = 1 To VBA.Len(json_Text)
        json_Char = VBA.Mid$(json_Text, json_Index, 1)
        json_AscCode = VBA.AscW(json_Char)

        ' When AscW returns a negative number, it returns the twos complement form of that number.
        ' To convert the twos complement notation into normal binary notation, add 0xFFF to the return result.
        ' https://support.microsoft.com/en-us/kb/272138
        If json_AscCode < 0 Then
            json_AscCode = json_AscCode + 65536
        End If

        ' From spec, ", \, and control characters must be escaped (solidus is optional)

        Select Case json_AscCode
        Case 34
            ' " -> 34 -> \"
            json_Char = "\"""
        Case 92
            ' \ -> 92 -> \\
            json_Char = "\\"
'        Case 47
'            ' / -> 47 -> \/ (optional)
'            If JsonOptions.EscapeSolidus Then
'                json_Char = "\/"
'            End If
        Case 8
            ' backspace -> 8 -> \b
            json_Char = "\b"
        Case 12
            ' form feed -> 12 -> \f
            json_Char = "\f"
        Case 10
            ' line feed -> 10 -> \n
            json_Char = "\n"
        Case 13
            ' carriage return -> 13 -> \r
            json_Char = "\r"
        Case 9
            ' tab -> 9 -> \t
            json_Char = "\t"
        Case 0 To 31, 127 To 65535
            ' Non-ascii characters -> convert to 4-digit hex
            json_Char = "\u" & VBA.Right$("0000" & VBA.Hex$(json_AscCode), 4)
        End Select

        json_BufferAppend json_Buffer, json_Char, json_BufferPosition, json_BufferLength
    Next json_Index

    json_Encode = json_BufferToString(json_Buffer, json_BufferPosition)
End Function
Public Sub json_BufferAppend(ByRef json_Buffer As String, _
                              ByRef json_Append As Variant, _
                              ByRef json_BufferPosition As Long, _
                              ByRef json_BufferLength As Long)
    ' VBA can be slow to append strings due to allocating a new string for each append
    ' Instead of using the traditional append, allocate a large empty string and then copy string at append position
    '
    ' Example:
    ' Buffer: "abc  "
    ' Append: "def"
    ' Buffer Position: 3
    ' Buffer Length: 5
    '
    ' Buffer position + Append length > Buffer length -> Append chunk of blank space to buffer
    ' Buffer: "abc       "
    ' Buffer Length: 10
    '
    ' Put "def" into buffer at position 3 (0-based)
    ' Buffer: "abcdef    "
    '
    ' Approach based on cStringBuilder from vbAccelerator
    ' http://www.vbaccelerator.com/home/VB/Code/Techniques/RunTime_Debug_Tracing/VB6_Tracer_Utility_zip_cStringBuilder_cls.asp
    '
    ' and clsStringAppend from Philip Swannell
    ' https://github.com/VBA-tools/VBA-JSON/pull/82

    Dim json_AppendLength As Long
    Dim json_LengthPlusPosition As Long

    json_AppendLength = VBA.Len(json_Append)
    json_LengthPlusPosition = json_AppendLength + json_BufferPosition

    If json_LengthPlusPosition > json_BufferLength Then
        ' Appending would overflow buffer, add chunk
        ' (double buffer length or append length, whichever is bigger)
        Dim json_AddedLength As Long
        json_AddedLength = IIf(json_AppendLength > json_BufferLength, json_AppendLength, json_BufferLength)

        json_Buffer = json_Buffer & VBA.Space$(json_AddedLength)
        json_BufferLength = json_BufferLength + json_AddedLength
    End If

    ' Note: Namespacing with VBA.Mid$ doesn't work properly here, throwing compile error:
    ' Function call on left-hand side of assignment must return Variant or Object
    Mid$(json_Buffer, json_BufferPosition + 1, json_AppendLength) = CStr(json_Append)
    json_BufferPosition = json_BufferPosition + json_AppendLength
End Sub
Public Function json_BufferToString(ByRef json_Buffer As String, ByVal json_BufferPosition As Long) As String
    If json_BufferPosition > 0 Then
        json_BufferToString = VBA.Left$(json_Buffer, json_BufferPosition)
    End If
End Function
