Attribute VB_Name = "modSprachen"
Option Compare Database
Option Explicit

Public gFehler As Integer                                              ' kann für eigene Fehlerbehandlung gebraucht werden
Public Const cCaption As Integer = 1                        ' Konstante für die Beschriftung des Feldes
Public Const cToolTipText As Integer = 2              ' Konstante für den Text des Tool Tips
Public Const cStatusText As Integer = 3                  ' Konstante für den Status Text (Info-Zeile unten)
Public Const cValidationText As Integer = 4          ' Konstante für den Überprüfungstext
Public gFormName As String                                        ' enthält den Namen des aktuellen Formulars

Function fctgSprachID() As Integer

          fctgSprachID = TempVars!Sprache

End Function

Function GetMessage(MsgID) As String
'
' Meldungstext zu einer Message gemäss Spracheinstellung
' ermitteln
'
Dim varVar As Variant
                                            ' Text aus der Tabelle "tblSprachen" lesen
    varVar = RstLookup("[lngText]", "tblSprachen", "[shtText]=" & [MsgID], _
        "[lguID]=" & TempVars!Sprache)
                                            ' prüfen, ob eine Meldung gefunden wurde
    If IsNull(varVar) Then
        GetMessage = "Kein Meldungstext für ID " & Str$(MsgID) & " definiert"
    Else
        GetMessage = varVar
    End If
 
End Function

Private Function RstLookup( _
ByVal strFieldName As String, _
ByVal strSource As String, _
Optional ByVal strCriteria1 As String = vbNullString, _
Optional ByVal strCriteria2 As String = vbNullString _
) As Variant

Dim rst As DAO.Recordset
Dim strSQL As String

          ' SQL String aufbauen, zuerst Feldname und Tabelle
          strSQL = "SELECT " & strFieldName & _
                    " FROM " & strSource
          If strCriteria1 > vbNullString Then             ' 1.Kriterium anfügen falls vorhanden
                    strSQL = strSQL & " WHERE " & strCriteria1
          End If
          If strCriteria2 > vbNullString Then             ' 2.Kriterium anfügen falls vorhanden
                    strSQL = strSQL & " AND " & strCriteria2
          End If
          Set rst = CurrentDb.OpenRecordset(strSQL, dbOpenSnapshot)
          
          With rst
                    If .EOF Then
                              RstLookup = Null
                    Else
                              RstLookup = .Fields(0)
                    End If
                    .Close
          End With
          Set rst = Nothing

End Function

Sub SetFormBeschriftung(actFrm As Form)
'
' Alle Felder für ein Formular entsprechend
' der Spracheinstellung setzen.
'
Dim actCtrl As Control

On Error GoTo err_SetFormBeschriftung

          gFormName = actFrm.Name
                                              ' Formulartitel (Überschrift)
          actFrm.Caption = GetBeschriftung(actFrm.Name, cCaption)
                                              ' Alle Steuerelemente des Formulars prüfen
          For Each actCtrl In actFrm.Controls
                    If actCtrl.Visible = True Then
                                                    ' Prüfen, ob der Steuerelement-Typ eine
                                                    ' Beschriftung beinhaltet.
                                                    ' Dies sind:
                                                    ' + Bezeichungsfelder: acLabel
                                                    ' + Pushbuttons: acCommandButton
                              If actCtrl.ControlType = acLabel Or _
                                        actCtrl.ControlType = acCommandButton Then
                                                  actCtrl.Caption = GetBeschriftung(actCtrl.Name, cCaption)
                              End If
                                                                  ' Tool-Tip-Text für alle Steuerelemente setzen
                              actCtrl.ControlTipText = GetBeschriftung(actCtrl.Name, cToolTipText)
                                                                  ' Statuszeilentext für alle Steuerelemente (außer Labels) setzen
                              If actCtrl.ControlType <> acLabel Then
                                        actCtrl.StatusBarText = GetBeschriftung(actCtrl.Name, cStatusText)
                              End If
                                                                  ' Wenn notwendig, dann die Gültigkeitsmeldung setzen
                              If actCtrl.ControlType <> acLabel And _
                                        actCtrl.ControlType <> acCommandButton Then
                                                  If actCtrl.ValidationRule <> "" Then
                                                            actCtrl.ValidationText = GetBeschriftung(actCtrl.Name, cValidationText)
                                                  End If
                              End If
                    End If
          Next actCtrl

Exit_SetFormBeschriftung:
          Exit Sub
          
err_SetFormBeschriftung:
          If Err.Number = 438 Then
                    Resume Next
          Else
                    Debug.Print Err.Number & " " & Err.Description
          End If
          Resume Exit_SetFormBeschriftung
 
End Sub

Sub SetReportBeschriftung(actRpt As Report)
'
' Alle Felder für ein Report entsprechend
' der Spracheinstellung setzen.
'
Dim actCtrl As Control
          gFormName = actRpt.Name
                                              ' Formulartitel (Überschrift)
          actRpt.Caption = GetBeschriftung(actRpt.Name, cCaption)
                                              ' Alle Steuerelemente des Formulars prüfen
          For Each actCtrl In actRpt.Controls
                    If actCtrl.Visible = True Then
                                                    ' Prüfen, ob der Steuerelement-Typ eine
                                                    ' Beschriftung beinhaltet.
                                                    ' Dies sind:
                                                    ' + Bezeichungsfelder: acLabel
                                                    ' + Pushbuttons: acCommandButton
                              If actCtrl.ControlType = acLabel Or _
                                        actCtrl.ControlType = acCommandButton Then
                                                  actCtrl.Caption = GetBeschriftung(actCtrl.Name, cCaption)
                              End If
                    End If
          Next actCtrl
 
End Sub

Function GetBeschriftung(txtCtrlName As String, intArt As Integer) As String
'
' Ermittelt die Beschriftung zu einem Steuerelement
' Entweder RstLookup oder DLookup, je nachdem was schneller ist
' Lokal istDLookup schneller, im Netz ist RstLookup schneller

Dim varVar As Variant
                                        ' prüfen, ob ein Controlname angegeben ist
          If Trim$(txtCtrlName) = "" Then
                    MsgBox "Kein Controlname angegeben", "GetBeschriftung"
                    Exit Function
          End If
                                              ' Auswahl der BeschriftungsintArt
          Select Case intArt
                    Case cToolTipText:
                                                        ' Tool-Tip-Text aus Tabelle "tblFormSprachen" lesen
                              varVar = RstLookup("[ttpText]", "tblFormSprachen", "[ctrlName]='" & txtCtrlName & "'", _
                                  "[lguID]=" & TempVars!Sprache)
                              'varVar = DLookup("[ttpText]", "tblFormSprachen", _
                              '               "[ctrlName]='" & txtCtrlName & "' AND " & _
                              '               "[lguID]=" & Str$(TempVars!Sprache))
                    Case cCaption:
                                                        ' Beschriftung aus Tabelle "tblFormSprachen" lesen
                              varVar = RstLookup("[ttlField]", "tblFormSprachen", "[ctrlName]='" & txtCtrlName & "'", _
                                  "[lguID]=" & TempVars!Sprache)
                              'varVar = DLookup("[ttlField]", "tblFormSprachen", _
                              '               "[ctrlName]='" & txtCtrlName & "' AND " & _
                              '               "[lguID]=" & Str$(TempVars!Sprache))
                    Case cStatusText:
                                                        ' Statuszeilentext aus Tabelle "tblFormSprachen" lesen
                              varVar = RstLookup("[sttText]", "tblFormSprachen", "[ctrlName]='" & txtCtrlName & "'", _
                                  "[lguID]=" & TempVars!Sprache)
                              'varVar = DLookup("[sttText]", "tblFormSprachen", _
                              '               "[ctrlName]='" & txtCtrlName & "' AND " & _
                              '               "[lguID]=" & Str$(TempVars!Sprache))
                    Case cValidationText:
                                                        ' Gültigkeitmeldung aus Tabelle "tblFormSprachen" lesen
                              varVar = RstLookup("[gkmText]", "tblFormSprachen", "[ctrlName]='" & txtCtrlName & "'", _
                                  "[lguID]=" & TempVars!Sprache)
                              'varVar = DLookup("[gkmText]", "tblFormSprachen", _
                              '               "[ctrlName]='" & txtCtrlName & "' AND " & _
                              '               "[lguID]=" & Str$(TempVars!Sprache))
          End Select
                                              ' prüfen, ob eine Beschriftung gefunden wurde
          If IsNull(varVar) Then
                    GetBeschriftung = ""
                                ' Den Namen des Steuerelementes in die Tabelle tblFormSprachen eintragen,
                                ' damit man weiß was noch zu ergänzen ist. Ausser für Feld ID
                                ' und nur falls nicht schon vorhanden
                    If Not txtCtrlName = "ID" And FindEintrag(TempVars!Sprache, txtCtrlName, gFormName) = False Then
                        CurrentDb.Execute "INSERT INTO tblFormSprachen (lguID, ctrlName, Herkunft)" & _
                                      "VALUES ('" & TempVars!Sprache & "', '" & txtCtrlName & "', '" & gFormName & "');"
                    End If
          Else
                    GetBeschriftung = varVar
          End If
 
End Function

Function FindEintrag(intSprache As Integer, txtCtrl As String, txtForm As String) As Boolean
' überprüft ob ein Eintrag schon vorhanden ist und gibt True oder False zurück

Dim rst As DAO.Recordset

          Set rst = CurrentDb.OpenRecordset( _
                    "SELECT [ctrlName] " & _
                    "FROM tblFormSprachen " & _
                    "WHERE [lguID] = " & intSprache & _
                    "AND [ctrlName] ='" & txtCtrl & _
                    "'AND [Herkunft] ='" & txtForm & "'")
          
          With rst
                    If Not .EOF Then
                              FindEintrag = True
                    Else
                              FindEintrag = False
                    End If
          End With

          rst.Close
          Set rst = Nothing

End Function


