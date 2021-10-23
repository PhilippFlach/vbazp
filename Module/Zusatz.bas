Attribute VB_Name = "Zusatz"
Option Compare Database
Function fehlend()
fehlend = Forms!DetCheck.RecordsetClone.RecordCount
'Wird verwendet in Bestellquitt
End Function


Function fehlendGLOB()
fehlendGLOB = Forms!DetCheckGLOB.RecordsetClone.RecordCount
'Wird verwendet in BestellquittGLOB

End Function


