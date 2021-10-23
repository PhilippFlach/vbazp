VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ARTIKEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Beenden_Click()
On Error GoTo Err_Beenden_Click


    DoCmd.Close
    DoCmd.Close acForm, "Such Artikel"
Exit_Beenden_Click:
    Exit Sub

Err_Beenden_Click:
    MsgBox Err.Description
    Resume Exit_Beenden_Click

End Sub


Private Sub Form_BeforeUpdate(Cancel As Integer)
DatumArtikel = Date
End Sub


Sub ArtikelSuchen_Click()
On Error GoTo Err_ArtikelSuchen_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Such Artikel"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    
Exit_ArtikelSuchen_Click:
    Exit Sub

Err_ArtikelSuchen_Click:
    MsgBox Err.Description
    Resume Exit_ArtikelSuchen_Click
    
End Sub
Private Sub Befehl56_Gebinde_Inhalt_Click()
On Error GoTo Err_Befehl56_Gebinde_Inhalt_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "ARTIKELGEBINDE"
    
    stLinkCriteria = "[ArtNr]=" & Me![ArtNr]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Befehl56_Gebinde_Inhalt_Click:
    Exit Sub

Err_Befehl56_Gebinde_Inhalt_Click:
    MsgBox Err.Description
    Resume Exit_Befehl56_Gebinde_Inhalt_Click
    
End Sub
