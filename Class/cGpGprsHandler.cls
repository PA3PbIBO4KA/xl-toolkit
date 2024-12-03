VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cGpGprsHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' Name: cGpGprsHandler
' Author: Mikhail Krasyuk
' Date: 04.11.2024
' Update: 04.11.2024

Option Explicit

Private pWsGpGprs As Worksheet

Private pGpGprsRowsCount As Long

Public Sub Initialize(GpGprsSheet As Worksheet)
    Set pWsGpGprs = GpGprsSheet
    Let pGpGprsRowsCount = pWsGpGprs.Cells(Rows.Count, 1).End(xlUp).Row
End Sub

Public Sub UpdateTable()
    
    Dim Sheet As Worksheet
    
    Dim RowIndex, ColumnIndex As Long
    
    With pWsGpGprs
        
        For ColumnIndex = 1 To 4
            .Columns(53).EntireColumn.Insert Shift:=xlToRight
        Next ColumnIndex
        
        For ColumnIndex = 1 To 4
            .Columns(17).EntireColumn.Insert Shift:=xlToRight
        Next ColumnIndex
        
        .Columns(11).EntireColumn.Insert Shift:=xlToRight
        
        .Cells(1, 11).Value = "volume_Kb"
        
        .Cells(1, 18).Value = "origTime_Year"
        .Cells(1, 19).Value = "origTime_Month"
        .Cells(1, 20).Value = "origTime_Day"
        .Cells(1, 21).Value = "origTime_Hour"
        
        .Cells(1, 58).Value = "dateTime_Cp04_Year"
        .Cells(1, 59).Value = "dateTime_Cp04_Month"
        .Cells(1, 60).Value = "dateTime_Cp04_Day"
        .Cells(1, 61).Value = "dateTime_Cp04_Hour"
        
        For RowIndex = 2 To pGpGprsRowsCount
            
            .Cells(RowIndex, 11).Value = .Cells(RowIndex, 10).Value / 1024
            
            .Cells(RowIndex, 18).Value = Year(.Cells(RowIndex, 17).Value)
            .Cells(RowIndex, 19).Value = Month(.Cells(RowIndex, 17).Value)
            .Cells(RowIndex, 20).Value = Day(.Cells(RowIndex, 17).Value)
            .Cells(RowIndex, 21).Value = Hour(.Cells(RowIndex, 17).Value)
            
            .Cells(RowIndex, 58).Value = Year(.Cells(RowIndex, 57).Value)
            .Cells(RowIndex, 59).Value = Month(.Cells(RowIndex, 57).Value)
            .Cells(RowIndex, 60).Value = Day(.Cells(RowIndex, 57).Value)
            .Cells(RowIndex, 61).Value = Hour(.Cells(RowIndex, 57).Value)
            
        Next RowIndex
        
        .Name = "GPGPRS"
        .Tab.Color = RGB(244, 176, 132)
        
    End With
    
End Sub
