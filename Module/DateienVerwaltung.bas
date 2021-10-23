Attribute VB_Name = "DateienVerwaltung"
Option Compare Database
Option Explicit

Public Function IsHeaderTrailerFilePresent(stPathFile As String) As Boolean
On Error GoTo Err_IsHeaderTrailerFilePresent
Dim result As Boolean
Dim pathinfo As Variant

  pathinfo = Dir(stPathFile)
  If Len(pathinfo) < 1 Then
        Err.Raise vbObjectError + 67
  Else
        result = True
  End If


Exit_IsHeaderTrailerFilePresent:
    IsHeaderTrailerFilePresent = result
    Exit Function

Err_IsHeaderTrailerFilePresent:
    If Err.Number = vbObjectError + 67 Then
        MsgBox stPathFile & " nicht vorhanden!"
    Else
        MsgBox Err.Description
    End If
    result = False
    Resume Exit_IsHeaderTrailerFilePresent
End Function

Public Sub VerbindeDateien(stHeaderFileName As String, stMiddleFileName As String, stTrailerFileName As String, stCombinedFileName As String)
On Error GoTo Err_VerbindeDateien
    Dim process_id As Long
    Dim stCommand As String

    ' Start the program.
    
    stCommand = "D:\AuftragZP\zusammensetzen.bat " & stHeaderFileName & " " & stMiddleFileName & " " & stTrailerFileName & " " & stCombinedFileName
    
    process_id = Shell(stCommand, 1)
    
'    Dim wsh As Object
'    Set wsh = VBA.CreateObject("WScript.Shell")
'    Dim waitOnReturn As Boolean: waitOnReturn = True
'    Dim windowStyle As Integer: windowStyle = 1
'

    'wsh.Run stCommand, windowStyle, waitOnReturn

    MsgBox "Access hat die Dateien " & Dir(stHeaderFileName) & ", " & Dir(stMiddleFileName) & " und " & vbNewLine & _
        stTrailerFileName & " kombiniert zur Datei " & stCombinedFileName & "! ", vbInformation, "Kombinationsvorgang erfolgreich"
        
Exit_VerbindeDateien:
    Exit Sub

Err_VerbindeDateien:
   
    MsgBox Err.Description
   
    Resume Exit_VerbindeDateien
End Sub
