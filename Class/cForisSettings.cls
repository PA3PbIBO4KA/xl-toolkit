VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cForisSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' Name: cMacroReferences
' Author: Mikhail Krasyuk
' Date: 19.10.2024
' Update: 04.11.2024

Option Explicit

Private pWsSettings As Worksheet

Private pRngForisHeader As Range
Private pRootFolder As String

Private pIsSettingsChanged As Boolean

Public Sub Initialize(SettingsSheet As Worksheet)
    
    Set pWsSettings = SettingsSheet
    
    With pWsSettings
        Set pRngForisHeader = .Range("H2:Z2")
        Let pRootFolder = .Range("A2").Value
    End With
    
    Let pIsSettingsChanged = False
    
End Sub

Public Property Get ForisHeader()
    ForisHeader = pRngForisHeader
End Property

Public Sub Save()
    If pIsSettingsChanged Then
        ThisWorkbook.Save
    End If
End Sub
