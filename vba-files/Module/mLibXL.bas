Attribute VB_Name = "mLibXL"

' Name: mLibXL
' Author: Mikhail Krasyuk
' Date: 13.06.2024

Option Explicit

' Class instance
Global Lib As CLibXL

' FSO
Global FileSystem As Object

' Colors
Global mcWhiteColor As Long
Global mcLightGreyColor As Long
Global mcYellowColor As Long
Global mcGreenColor As Long

' One millisecond
Global Const mcMillisecond As Double = 0.000000011574

' Modes for calculating columns and rows
Global Const mcByFiltering = 0
Global Const mcByEnd = 1
Global Const mcByCycle = 2

' Character constants
Global Const mcWholeRange = "A1:XFD1048576"
Global Const mcBackslash = "\"

' For error handler
Global mcErrorExplanation As String

' Initialization
Public Sub Init()
    
    mcWhiteColor = RGB(255, 255, 255)
    mcLightGreyColor = RGB(220, 220, 220)
    mcYellowColor = RGB(255, 255, 0)
    mcGreenColor = RGB(0, 190, 0)
    
    Set FileSystem = CreateObject("Scripting.FileSystemObject")
    Set Lib = New CLibXL
    
End Sub

Public Sub EnableDebugMode()
    Stop
End Sub

' Will be called at the end of the program, via countdown
Public Sub CloseMacrosWorkbook()
    Workbooks(ThisWorkbook.Name).Close SaveChanges:=False
End Sub