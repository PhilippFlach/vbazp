VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Auswahl Adress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Sub AnfangGenau_Click()
On Error GoTo Err_AnfangGenau_Click


    DoCmd.DoMenuItem acFormBar, acRecordsMenu, 2, , acMenuVer70

Exit_AnfangGenau_Click:
    Exit Sub

Err_AnfangGenau_Click:
    MsgBox Err.Description
    Resume Exit_AnfangGenau_Click
    
End Sub
Sub NachnameGenau_Click()
On Error GoTo Err_NachnameGenau_Click


    DoCmd.DoMenuItem acFormBar, acRecordsMenu, 2, , acMenuVer70

Exit_NachnameGenau_Click:
    Exit Sub

Err_NachnameGenau_Click:
    MsgBox Err.Description
    Resume Exit_NachnameGenau_Click
    
End Sub
Private Sub OeffneAdresseInfo_Click()
On Error GoTo Err_OeffneAdresseInfo_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "ADR"
    
    stLinkCriteria = "[AdressNr]=" & Me![AdressNr]
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.Close acForm, "Auswahl Adress"

Exit_OeffneAdresseInfo_Click:
    Exit Sub

Err_OeffneAdresseInfo_Click:
    MsgBox Err.Description
    Resume Exit_OeffneAdresseInfo_Click
    
End Sub
Private Sub cmdNeueBestellung_Click()
On Error GoTo Err_cmdNeueBestellung_Click

    Dim stDocName As String

    stDocName = "oeffne Such Bestell"
    DoCmd.RunMacro stDocName
    Forms![Such Bestell]!Nr = Me!AdressNr
    stDocName = "Neue Bestellung"
    DoCmd.RunMacro stDocName
    DoCmd.Close acForm, "Auswahl Adress"
    DoCmd.Close acForm, "Such Adress"

Exit_cmdNeueBestellung_Click:
    Exit Sub

Err_cmdNeueBestellung_Click:
    MsgBox Err.Description
    Resume Exit_cmdNeueBestellung_Click
    
End Sub
