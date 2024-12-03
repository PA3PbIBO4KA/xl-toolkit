VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFORISBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' Name: cForisBuilder
' Author: Mikhail Krasyuk
' Date: 02.10.2024
' Update: 04.11.2024

Option Explicit

Private pTargetWorksheet As Worksheet

Private pRowsToDeleteCollection As Collection

Private pXLSFilesFolderPath As String
Private pMSISDNArray() As String
Private pMSISDN As String

Private pXLSFilesCounter As Long
Private pTargetWorksheetRowsCount As Long
Private pTargetWorksheetColumnsCount As Long

Private pIsMSISDNSingle As Boolean

Private Sub Class_Initialize()
End Sub

Public Property Let XLSFilesFolderPath(ByVal FolderPath As String)
    Let pXLSFilesFolderPath = FolderPath
    Let pXLSFilesCounter = 1
    Let pTargetWorksheetRowsCount = 0
End Property

Public Property Let MSISDNs(ByVal MSISDNs As String)

    If InStr(MSISDNs, ",") = 0 Then
        Let pMSISDN = MSISDNs
        Let pIsMSISDNSingle = True
    Else
        Let pMSISDNArray = Split(MSISDNs, ",")
        Let pIsMSISDNSingle = False
    End If
    
End Property

Public Property Get IsMSISDNSingle()
    IsMSISDNSingle = pIsMSISDNSingle
End Property

Private Sub CopyPasteValues(ByVal XLSWorksheet As Worksheet)
    
    Dim ExcessRowIndex As Variant
    
    Dim RowIndex As Long
    
    If Not Lib.IsInArray(XLSWorksheet.Range("C4"), pMSISDNArray) Then
        Exit Sub
    End If
    
    If pXLSFilesCounter = 1 Then
        Let RowIndex = 2
    Else
        Let RowIndex = pTargetWorksheetRowsCount + 1
    End If
    
    Set pRowsToDeleteCollection = New Collection
    
    Dim ColumnsCount As Long: Let ColumnsCount = Lib.GetColumnsCount(Row:=7, Sheet:=XLSWorksheet)
    Dim RowsCount    As Long: Let RowsCount = Lib.GetRowsCount(Sheet:=XLSWorksheet)
    
    If RowsCount = 7 Then
        Exit Sub
    End If
    
    Dim MSISDN As String: Let MSISDN = XLSWorksheet.Range("C4").Value
    
    With XLSWorksheet
        If pXLSFilesCounter = 1 Then
            .Range(.Cells(7, 1), .Cells(RowsCount, ColumnsCount)).Copy
        Else
            .Range(.Cells(8, 1), .Cells(RowsCount, ColumnsCount)).Copy
        End If
    End With
    
    Let pTargetWorksheetRowsCount = Lib.GetRowsCount(Sheet:=pTargetWorksheet, Method:=mcByEnd)
    
    If pXLSFilesCounter = 1 Then
        pTargetWorksheet.Range("B1").PasteSpecial xlPasteValuesAndNumberFormats
    Else
        pTargetWorksheet.Range("B" & pTargetWorksheetRowsCount + 1).PasteSpecial xlPasteValuesAndNumberFormats
    End If
    
    pTargetWorksheet.Range("A1").Value = "Абоненты"
    
    Let pTargetWorksheetRowsCount = Lib.GetRowsCount(Sheet:=pTargetWorksheet, Column:=2, Method:=mcByEnd)
    
    For RowIndex = RowIndex To pTargetWorksheetRowsCount
        pTargetWorksheet.Range("A" & RowIndex).Value = MSISDN
    Next RowIndex
    
    With pRowsToDeleteCollection
        .Add pTargetWorksheetRowsCount
        .Add pTargetWorksheetRowsCount - 1
    End With
    
    For Each ExcessRowIndex In pRowsToDeleteCollection
        pTargetWorksheet.Cells(ExcessRowIndex, 1).EntireRow.Delete
    Next ExcessRowIndex
    
    Let pTargetWorksheetRowsCount = Lib.GetRowsCount(Sheet:=pTargetWorksheet, Column:=2, Method:=mcByEnd)
    Let pTargetWorksheetColumnsCount = Lib.GetColumnsCount(Sheet:=pTargetWorksheet, Row:=1, Method:=mcByEnd)
    
End Sub

Public Sub CreateTable()
    
    If IsEmpty(pXLSFilesFolderPath) Then
        Exit Sub
    End If
    
    Dim XLSWorkbook As Workbook
    Dim XLSWorksheet As Worksheet
    
    Dim XLSFilePath As String: Let XLSFilePath = Dir(pXLSFilesFolderPath & "*.xls")
    
    Dim RowIndex As Long
    
    If Not pIsMSISDNSingle Then
        If XLSFilePath <> "" Then
            Dim pTargerWorkbook As Workbook: Set pTargerWorkbook = Workbooks.Add
            Set pTargetWorksheet = pTargerWorkbook.Worksheets(1)
        End If
    End If
    
    Do While XLSFilePath <> ""
        
        Set XLSWorkbook = Workbooks.Open(pXLSFilesFolderPath & XLSFilePath)
        Set XLSWorksheet = XLSWorkbook.Worksheets(1)
        
        If Not pIsMSISDNSingle Then
            Call CopyPasteValues(XLSWorksheet)
        Else
            If pMSISDN = CStr(XLSWorksheet.Range("C4").Value) Then
                Exit Sub
            End If
        End If
        
        XLSWorkbook.Close False
        
        pXLSFilesCounter = pXLSFilesCounter + 1
        
        XLSFilePath = Dir
        
    Loop
    
    With pTargetWorksheet
        
        .Name = "FORIS"
        
        .Tab.Color = RGB(155, 194, 230)
        
        .Columns(3).EntireColumn.Insert Shift:=xlToRight
        .Columns(5).EntireColumn.Insert Shift:=xlToRight
        .Columns(5).EntireColumn.Insert Shift:=xlToRight
        .Columns(8).EntireColumn.Insert Shift:=xlToRight
        
        .Range("A1:S1").Value = ThisWorkbook.Worksheets(1).Range("H2:Z2").Value
        
        .ListObjects.Add(xlSrcRange, .Range(.Cells(1, 1), .Cells(pTargetWorksheetRowsCount, 19)), , xlYes).Name = "ForisTable"
        .ListObjects.Item(1).TableStyle = "TableStyleMedium16"
        
        For RowIndex = 2 To pTargetWorksheetRowsCount
            
            .Range("C" & RowIndex).Value = Day(.Range("B" & RowIndex))
            .Range("E" & RowIndex).Value = Hour(.Range("D" & RowIndex))
            .Range("F" & RowIndex).Value = Minute(.Range("D" & RowIndex))
            
            If .Range("I" & RowIndex).Value Like "<--" & "*" Then
                .Range("H" & RowIndex).Value = 20
            Else
                .Range("H" & RowIndex).Value = 10
            End If
            
        Next RowIndex
        
        .Columns("N:N").Replace What:="Kb", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, _
                                            MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
        .Rows.AutoFit
        .Columns.AutoFit
        
        .Range("A1").Select
        
    End With
    
End Sub
