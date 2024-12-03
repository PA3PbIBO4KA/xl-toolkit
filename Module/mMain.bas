Attribute VB_Name = "mMain"

' Name: mMain
' Author: Mikhail Krasyuk
' Date: 06.10.2024
' Update: 05.11.2024

Option Explicit

Public Sub CloseMacrosWorkbook()
    Workbooks(ThisWorkbook.Name).Close SaveChanges:=False
End Sub

Public Sub ExitMacros()
    Application.OnTime Now + TimeValue("00:00:01"), "CloseMacrosWorkbook"
End Sub

Sub Main()
    Application.Windows(ThisWorkbook.Name).Visible = False
End Sub

Sub Show()
    frmMemo.Show
End Sub

Sub CloseMacro()
    Call CloseMacrosWorkbook
End Sub
