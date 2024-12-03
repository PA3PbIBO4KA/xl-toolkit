Attribute VB_Name = "mMain"

' Name: mMain
' Author: Mikhail Krasyuk
' Date: 03.10.2024
' Update: 05.11.2024

Option Explicit

Sub Show()
    Let Application.Windows(ThisWorkbook.Name).Visible = True
End Sub

Sub FORIS()
    
    Let Application.Windows(ThisWorkbook.Name).Visible = False
    
    Call mCommonLib.Init
    
    Load frmPrimary
    Load frmSettings
    
    Call frmPrimary.Show
    
End Sub

Sub GpGprs()
    
    Dim objGpGprs As New cGpGprsHandler
    
    Let Application.Windows(ThisWorkbook.Name).Visible = False
    
    Call mCommonLib.Init
    
    With objGpGprs
        .Initialize ActiveSheet
        .UpdateTable
    End With
    
    Call Lib.ExitMacros
    
End Sub
