Attribute VB_Name = "mHotKeys"

' Name: mHotKeys
' Author: Mikhail Krasyuk
' Date: 06.10.2024
' Update: 05.11.2024

Option Explicit

' Shift + Ctrl + H
Public Sub HideMacrosWorkbook()
Attribute HideMacrosWorkbook.VB_ProcData.VB_Invoke_Func = "H\n14"
    Application.Windows(ThisWorkbook.Name).Visible = IIf(ThisWorkbook.Windows(1).Visible, False, True)
End Sub

' Shift + Ctrl + A
Sub CreateTable()
Attribute CreateTable.VB_ProcData.VB_Invoke_Func = "A\n14"
    
    Dim Sheet As Worksheet: Set Sheet = ActiveSheet
    
    Dim arrNeedlessSymbols As Variant: arrNeedlessSymbols = Array(":", ".", ",")
    Dim arrNeedlessFields As Variant: arrNeedlessFields = Array("time")
    
    Dim UpperRow As Long: Let UpperRow = 1
    Dim LeftistColumn As Long: Let LeftistColumn = 1
    
    Dim TableRange As Range
    Dim TableObject As ListObject
    
    Dim IMSIIndex As Long: Let IMSIIndex = 0
    Dim MSISDNColumn As Long: Let MSISDNColumn = 0
    
    Dim ColumnIndex, RowIndex As Long
    
    Dim RowsCount As Long: Let RowsCount = Sheet.Cells(Rows.Count, LeftistColumn).End(xlUp).Row
    Dim ColumnsCount As Long: Let ColumnsCount = Sheet.Cells(UpperRow, 1).End(xlToRight).Column
    
    With Sheet
        
        Dim Cell As Variant
        Dim IsAllCellsNumerical As Boolean
        
        Set TableRange = .Range(.Cells(UpperRow, LeftistColumn), .Cells(RowsCount, ColumnsCount))
        
        On Error Resume Next
        Set TableObject = .ListObjects(1)
        On Error GoTo 0
        
        If Not TableObject Is Nothing Then
            TableObject.TableStyle = ""
            TableObject.Unlist
        End If
        
        .Range(.Cells(1, LeftistColumn), .Cells(1, ColumnsCount)).Font.Bold = True
        .Rows(UpperRow).AutoFilter
        
        For ColumnIndex = 1 To ColumnsCount
            
            Let IsAllCellsNumerical = True
            
            If mAdditional.IsInArray(CStr(.Cells(1, ColumnIndex).Value), arrNeedlessFields, mcFully) Then
                Let IsAllCellsNumerical = False
                GoTo NextColumnIndex
            End If
            
            For Each Cell In .Range(.Cells(2, ColumnIndex), .Cells(RowsCount, ColumnIndex))
                If Not IsNumeric(CStr(Cell.Value)) And mAdditional.IsInArray(CStr(Cell.Value), arrNeedlessSymbols, mcPartly) Then
                    Let IsAllCellsNumerical = False
                    Exit For
                End If
            Next Cell
            
            If IsAllCellsNumerical Then
                .Range(.Cells(UpperRow + 1, ColumnIndex), .Cells(RowsCount, ColumnIndex)).NumberFormat = "0"
                Let IsAllCellsNumerical = False 'îáíóëåíèå ïåðåìåííîé
            End If
            
NextColumnIndex:
        Next ColumnIndex
        
    End With
    
    TableRange.Columns.AutoFit
    
End Sub

' Shift + Ctrl + P
Sub CreatePivotTable()
Attribute CreatePivotTable.VB_ProcData.VB_Invoke_Func = "P\n14"
    
    Dim wbActive As Workbook: Set wbActive = Workbooks(ActiveWorkbook.Name)
    Dim wsData As Worksheet: Set wsData = ActiveSheet
    Dim wsPivots As Worksheet
    
    On Error Resume Next
    Set wsPivots = wbActive.Worksheets("PIVOTS")
    On Error GoTo 0
    
    If wsPivots Is Nothing Then
        Set wsPivots = wbActive.Worksheets.Add
        wsPivots.Name = "PIVOTS"
        wsPivots.Tab.Color = RGB(219, 219, 219)
    Else
    End If
    
    Dim PivotField As String
    
    Dim RowIndex, ColumnIndex As Long
    Dim RowsCount As Long: Let RowsCount = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    Dim ColumnsCount As Long: Let ColumnsCount = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column
    Dim DataRange As Range: Set DataRange = wsData.Range(wsData.Cells(1, 1), wsData.Cells(RowsCount, ColumnsCount))
    
    Dim pcData As PivotCache: Set pcData = wbActive.PivotCaches.Create( _
                                                        SourceType:=xlDatabase, _
                                                        SourceData:=DataRange)
    
    Dim ptData As PivotTable: Set ptData = pcData.CreatePivotTable( _
                                                TableDestination:=wsPivots.Cells(1, 1), _
                                                TableName:="PivotTable")
    
    For ColumnIndex = 1 To ColumnsCount
        If Application.WorksheetFunction.CountA(wsData.Columns(ColumnIndex)) = RowsCount Then
            PivotField = wsData.Cells(1, ColumnIndex).Value
            Exit For
        End If
    Next ColumnIndex
    
    If PivotField <> "" Then
        With ptData
            .AddDataField .PivotFields(PivotField), "Êîë-âî çàïèñåé", xlCount
        End With
    End If
    
End Sub

' Shift + Ctrl + R
' Êîïèðóåò êîë-âî ñòðî÷åê èç àêòèâíîãî ëèñòà
Sub GetRowsCount()
Attribute GetRowsCount.VB_ProcData.VB_Invoke_Func = "R\n14"
    
    Dim DataObject As New MSForms.DataObject
    Dim Sheet As Worksheet: Set Sheet = ActiveSheet
    Dim RowsCount As Long: Let RowsCount = Sheet.Cells(Rows.Count, 1).End(xlUp).Row - 1
    
    DataObject.SetText CStr(RowsCount)
    DataObject.PutInClipboard
    
End Sub

' Shift + Ctrl + C
Sub TransformSelectionToClipboard()
Attribute TransformSelectionToClipboard.VB_ProcData.VB_Invoke_Func = "C\n14"
    
    Dim DataObject As New MSForms.DataObject
    
    Dim SelectedRange      As Range: Set SelectedRange = Selection
    Dim ValuesArray()      As Variant: ValuesArray = SelectedRange.Value
    Dim TransformedValues  As String: TransformedValues = Join(Application.Transpose(Application.Index(ValuesArray, 0, 0)), ",")
    
    DataObject.SetText TransformedValues
    DataObject.PutInClipboard
    
End Sub

' Shift + Ctrl + S
Sub SaveFile()
Attribute SaveFile.VB_ProcData.VB_Invoke_Func = "S\n14"
    
    Application.DisplayAlerts = False
    
    With ActiveWorkbook
        .SaveAs Filename:=.Path & "\" & Left(.Name, InStrRev(.Name, ".") - 1) & ".xlsx", FileFormat:=xlOpenXMLWorkbook
    End With
    
    Application.DisplayAlerts = True
    
End Sub
