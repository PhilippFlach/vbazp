VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EMAILTYPEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Edit_Email_Click()
On Error GoTo Err_Edit_Email_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "EMAILTYPEN_HF"
    
    stLinkCriteria = "[EMailTypID]=" & Me![EMailTypID]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Edit_Email_Click:
    Exit Sub

Err_Edit_Email_Click:
    MsgBox Err.Description
    Resume Exit_Edit_Email_Click
    
End Sub
Private Sub EMailVersandMehrfach_Click()
On Error GoTo Err_EMailVersandMehrfach_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim lgEMailTypID As Long
    Const stmehrfachauswahl  As String = "Mehrere Datensätze aus"
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    lgEMailTypID = Me![EMailTypID]
    stDocName = "EMailTyp_Wahl"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    With Forms![EMailTyp_Wahl]
        ![Versandart] = conEmailMassenversand
        ![EMailTypID] = lgEMailTypID
        ![Vorname] = stmehrfachauswahl
        ![Nachname] = Me![AbfragEmailAdressen]
    End With
Exit_EMailVersandMehrfach_Click:
    Exit Sub

Err_EMailVersandMehrfach_Click:
    MsgBox Err.Description
    Resume Exit_EMailVersandMehrfach_Click
    
End Sub

Private Sub PreiskatNr_AfterUpdate()
On Error GoTo Err_PreiskatNr_AfterUpdate
    Dim varPreiskatNr As Variant
    varPreiskatNr = Me!PreiskatNr
'    If varPreiskatNr = 2 Then
'        Me!AbfragEmailAdressen = "AvisierungMails"
'    Else
'        Me!AbfragEmailAdressen = "AvisierungMails_Geschäftl"
'    End If
Exit_PreiskatNr_AfterUpdate:
    Exit Sub

Err_PreiskatNr_AfterUpdate:
    MsgBox Err.Description
    Resume Exit_PreiskatNr_AfterUpdate

End Sub
Private Sub Befehl17_Click()
On Error GoTo Err_Befehl17_Click


    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70

Exit_Befehl17_Click:
    Exit Sub

Err_Befehl17_Click:
    MsgBox Err.Description
    Resume Exit_Befehl17_Click
    
End Sub
