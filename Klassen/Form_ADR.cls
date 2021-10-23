VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ADR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_BeforeUpdate(Cancel As Integer)
DatumAdress = Date
End Sub


Sub AdressSuche_Click()
On Error GoTo Err_AdressSuche_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Such Adress"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_AdressSuche_Click:
    Exit Sub

Err_AdressSuche_Click:
    MsgBox Err.Description
    Resume Exit_AdressSuche_Click
    
End Sub

Private Sub Postleitzahl_Exit(Cancel As Integer)
On Error GoTo Err_Postleitzahl_Exit
    Dim tgtdb As Database
    Dim stsql As String
    Dim lgPostleitzahl As Long
    Dim plzqry As QueryDef
    Dim plzrs As Recordset
    Dim stDocName As String
    Dim stqueryname As String
    Dim varort As Variant
    Set tgtdb = CurrentDb
    'E-Mail Typ: Eigenschaften bestimmen
    stqueryname = "PLZ_Lookup"
    lgPostleitzahl = Nz(Me![Postleitzahl], 0)
    varort = Me![Ort]
    stsql = "SELECT d.Postleitzahl, d.Ort" & vbNewLine
    stsql = stsql & "FROM REGIONEN_LINK AS d" & vbNewLine
    stsql = stsql & "WHERE d.Postleitzahl = " & lgPostleitzahl & ";"
    Set plzqry = tgtdb.QueryDefs(stqueryname)
    plzqry.SQL = stsql
    Set plzrs = plzqry.OpenRecordset(dbOpenSnapshot)
    If plzrs.EOF Then Err.Raise vbObjectError + 37
    plzrs.MoveFirst
    'If IsNull(varort) Then Me![Ort] = plzrs!Ort
    Me![Ort] = plzrs!Ort
    plzqry.Close
    plzrs.Close

Exit_Postleitzahl_Exit:
    Exit Sub

Err_Postleitzahl_Exit:
    If Err.Number = vbObjectError + 37 Then
        MsgBox "Postleitzahl nicht vorhanden"
        stDocName = "Such Region"
        DoCmd.OpenForm stDocName
    Else
        MsgBox Err.Description
    End If
    Resume Exit_Postleitzahl_Exit

End Sub

Sub SchliesseADR_Click()
On Error GoTo Err_SchliesseADR_Click


    DoCmd.Close
    DoCmd.Close acForm, "Such Adress"
Exit_SchliesseADR_Click:
    Exit Sub

Err_SchliesseADR_Click:
    MsgBox Err.Description
    Resume Exit_SchliesseADR_Click
    
End Sub
Sub Gastro_Info_Click()
On Error GoTo Err_Gastro_Info_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "GASTRO"
    
    stLinkCriteria = "[AdressNr]=" & Me![AdressNr]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Gastro_Info_Click:
    Exit Sub

Err_Gastro_Info_Click:
    MsgBox Err.Description
    Resume Exit_Gastro_Info_Click
    
End Sub
Sub Personen_Info_Click()
On Error GoTo Err_Personen_Info_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "PERS"
    
    stLinkCriteria = "[AdressNr]=" & Me![AdressNr]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Personen_Info_Click:
    Exit Sub

Err_Personen_Info_Click:
    MsgBox Err.Description
    Resume Exit_Personen_Info_Click
    
End Sub
Sub Rechnungs_Adresse_Click()
On Error GoTo Err_Rechnungs_Adresse_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "RECHADR"
    
    stLinkCriteria = "[AdressNr]=" & Me![AdressNr]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Rechnungs_Adresse_Click:
    Exit Sub

Err_Rechnungs_Adresse_Click:
    MsgBox Err.Description
    Resume Exit_Rechnungs_Adresse_Click
    
End Sub


Private Sub Firmbezug_Click()
On Error GoTo Err_Firmbezug_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "FIRMENBEZ"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Firmbezug_Click:
    Exit Sub

Err_Firmbezug_Click:
    MsgBox Err.Description
    Resume Exit_Firmbezug_Click
    
End Sub
Private Sub Befehl53_Click()
On Error GoTo Err_Befehl53_Click

    Dim stDocName As String

    stDocName = "Kundeninformation"
    DoCmd.OpenReport stDocName, acPreview

Exit_Befehl53_Click:
    Exit Sub

Err_Befehl53_Click:
    MsgBox Err.Description
    Resume Exit_Befehl53_Click
    
End Sub
Private Sub WähleTelNr1_Click()
On Error GoTo Err_WähleTelNr1_Click

    Dim stDialStr As String
    Dim PrevCtl As Control
    Const ERR_OBJNOTEXIST = 2467
    Const ERR_OBJNOTSET = 91
    Const ERR_CANTMOVE = 2483

    Set PrevCtl = Screen.PreviousControl
    

      stDialStr = "0" & Me!Telefon1

    Application.Run "utility.wlib_AutoDial", stDialStr

Exit_WähleTelNr1_Click:
    Exit Sub

Err_WähleTelNr1_Click:
    If (Err = ERR_OBJNOTEXIST) Or (Err = ERR_OBJNOTSET) Or (Err = ERR_CANTMOVE) Then
      Resume Next
    End If
    MsgBox Err.Description
    Resume Exit_WähleTelNr1_Click
    
End Sub
Private Sub Wähle_TelNr_1_Click()
On Error GoTo Err_Wähle_TelNr_1_Click

    Dim stDialStr As String
    Dim PrevCtl As Control
    Const ERR_OBJNOTEXIST = 2467
    Const ERR_OBJNOTSET = 91
    Const ERR_CANTMOVE = 2483

    Set PrevCtl = Screen.PreviousControl
    
    
      stDialStr = "0" & Me!Telefon1
        
    Application.Run "utility.wlib_AutoDial", stDialStr

Exit_Wähle_TelNr_1_Click:
    Exit Sub

Err_Wähle_TelNr_1_Click:
    If (Err = ERR_OBJNOTEXIST) Or (Err = ERR_OBJNOTSET) Or (Err = ERR_CANTMOVE) Then
      Resume Next
    End If
    MsgBox Err.Description
    Resume Exit_Wähle_TelNr_1_Click
    
End Sub
