Attribute VB_Name = "mCommonLib"

' Name: mCommonLib
' Author: Mikhail Krasyuk
' Date: 13.06.2024

Option Explicit

Global Lib As cCommonLib

Global FileSystem As Object

Global mcWhiteColor As Long
Global mcLightGreyColor As Long
Global mcYellowColor As Long
Global mcGreenColor As Long

Global Const mcMillisecond As Double = 0.000000011574

Global Const mcByFiltering = 0
Global Const mcByEnd = 1
Global Const mcByCycle = 2

Global Const mcEmpty = ""
Global Const mcWholeRange = "A1:XFD1048576"
Global Const mcBackslash = "\"

Global mcErrorExplanation As String

Public Sub Init()
    
    mcWhiteColor = RGB(255, 255, 255)
    mcLightGreyColor = RGB(220, 220, 220)
    mcYellowColor = RGB(255, 255, 0)
    mcGreenColor = RGB(0, 190, 0)
    
    Set FileSystem = CreateObject("Scripting.FileSystemObject")
    Set Lib = New cCommonLib
    
End Sub

Public Sub EnableDebugMode()
    Stop
End Sub

Public Sub CloseMacrosWorkbook()
    Workbooks(ThisWorkbook.Name).Close SaveChanges:=False
End Sub



