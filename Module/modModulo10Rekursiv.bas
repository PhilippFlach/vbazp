Attribute VB_Name = "modModulo10Rekursiv"
Option Compare Database
Option Explicit

Private Sub check()
' dient zum testen der Modulo 10 Berechnung

Dim strRef As String
Dim strCheck As String

          strRef = "21000000000313947143000901"
          strCheck = Mod10R(strRef)
          strRef = strRef & strCheck
          
End Sub

Public Function Mod10R(ByVal strRef As String) As Integer
Dim i As Integer
Dim p As Integer
Dim c As Integer
Dim z As Integer
Dim e As Integer

On Error GoTo err_Mod10R

          c = 0
          p = 0
          For i = 1 To Len(strRef)
                    z = CInt(Mid(strRef, i, 1))
                    c = p + z
                    Select Case c
                              Case 0
                                        p = 0
                              Case 1
                                        p = 9
                              Case 2
                                        p = 4
                              Case 3
                                        p = 6
                              Case 4
                                        p = 8
                              Case 5
                                        p = 2
                              Case 6
                                        p = 7
                              Case 7
                                        p = 1
                              Case 8
                                        p = 3
                              Case 9
                                        p = 5
                               Case 10
                                        p = 0
                              Case 11
                                        p = 9
                              Case 12
                                        p = 4
                              Case 13
                                        p = 6
                              Case 14
                                        p = 8
                              Case 15
                                        p = 2
                              Case 16
                                        p = 7
                              Case 17
                                        p = 1
                              Case 18
                                        p = 3
                              Case Else
                                        MsgBox "Es ist ein Fehler aufgetreten.", vbCritical, "Error"
                                        GoTo err_Mod10R
                              End Select
                    Next
                    Select Case p
                              Case 0
                                        e = 0
                              Case 1
                                        e = 9
                              Case 2
                                        e = 8
                              Case 3
                                        e = 7
                              Case 4
                                        e = 6
                              Case 5
                                        e = 5
                              Case 6
                                        e = 4
                              Case 7
                                        e = 3
                              Case 8
                                        e = 2
                              Case 9
                                        e = 1
                              Case Else
                                        MsgBox "Es ist ein Fehler aufgetreten.", vbCritical, "Error"
                                        GoTo err_Mod10R
                              End Select
          Mod10R = e
                                       
Exit_Mod10R:
          Exit Function
          
err_Mod10R:
          Debug.Print Err.Number & " " & Err.Description
          Resume Exit_Mod10R
                    
End Function
