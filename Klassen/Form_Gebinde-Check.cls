VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Gebinde-Check"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Sub Bestellkorrg_Click()
On Error GoTo Err_Bestellkorrg_Click


    DoCmd.Close

Exit_Bestellkorrg_Click:
    Exit Sub

Err_Bestellkorrg_Click:
    MsgBox Err.Description
    Resume Exit_Bestellkorrg_Click
    
End Sub
Sub OK_Click()
On Error GoTo Err_OK_Click


    DoCmd.Close
    DoCmd.Close acForm, "BESTELL"
    DoCmd.OpenForm "Such Bestell"
    DoCmd.GoToControl "Nr"

Exit_OK_Click:
    Exit Sub

Err_OK_Click:
    MsgBox Err.Description
    Resume Exit_OK_Click
    
End Sub
