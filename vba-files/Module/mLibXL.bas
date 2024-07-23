Attribute VB_Name = "mLibXL"

' Name: mLibXL
' Author: Михаил Красюк
' Date: 13.06.2024

Option Explicit

' Экземпляр класса
Global Lib As CLibXL

' FSO
Global FileSystem As Object

' Цвета
Global mcWhiteColor As Long
Global mcLightGreyColor As Long
Global mcYellowColor As Long
Global mcGreenColor As Long

' Одна миллисекунда
Global Const mcMillisecond As Double = 0.000000011574

' Режимы расчета столбцов и строк
Global Const mcByFiltering = 0
Global Const mcByEnd = 1
Global Const mcByCycle = 2

' Символьные константы
Global Const mcWholeRange = "A1:XFD1048576"
Global Const mcBackslash = "\"

' Для обработчика ошибок
Global mcErrorExplanation As String

' Инициализация
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

' Будет вызываться в конце программы, через обратный отсчет
Public Sub CloseMacrosWorkbook()
    Workbooks(ThisWorkbook.Name).Close SaveChanges:=False
End Sub

