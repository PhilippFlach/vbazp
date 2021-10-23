Attribute VB_Name = "modVariables"
Option Compare Database
Option Explicit

Public Const OnTheFly As Boolean = True
Public Const TimeOut As Integer = 300
Public Const AutoSelect As Byte = 0
Public gstrComboControlSource As String
Public gMeldung As Long
Public gbolSprache As Boolean

Public Const cZGUltmtCdtr As String = ""
Public Const cZGAdrTp As String = ""
Public Const cZGName As String = ""
Public Const cZGStrtNmOrAdrLine1 As String = ""
Public Const cZGBldgNbOrAdrLine2 As String = ""
Public Const cZGPstCd As String = ""
Public Const cZGTwnNm As String = ""
Public Const cZGCtry As String = ""

Public varArray(26, 2) As Variant
Public Const cPZ As Integer = 0
