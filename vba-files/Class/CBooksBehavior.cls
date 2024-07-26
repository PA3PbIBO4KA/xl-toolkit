VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBooksBehavior"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' Name: CBooksBehavior
' Author: Mikhail Krasyuk
' Date: 27.07.2024

Option Explicit

Private m_wbMacros As Workbook
Private m_wsMacrosMainMenu As Worksheet
Private m_wsMacrosReferences As Worksheet

Private m_wbCommon As Workbook
Private m_wsCommonPrimary As Worksheet

Private m_wbPreviousCommon As Workbook
Private m_wsPreviousCommonPrimary As Worksheet

Private m_wbForm As Workbook
Private m_wsFormPrimary As Worksheet

Private m_wbASR As Workbook
Private m_wsASRPrimary As Worksheet

Private m_wbSplitting As Workbook
Private m_wsSplittingPrimary As Worksheet

Private m_sCommonBookName As String

' Initialization
Private Sub class_initialize()

    On Error GoTo ErrorHandler_OpenFile

    Lib.AddLogNote "Macros Launched"

    Set m_wbMacros = Workbooks(ThisWorkbook.Name)

    With m_wbMacros
        Set m_wsMacrosMainMenu = .Worksheets("Main Menu")
        Set m_wsMacrosReferences = .Worksheets("References")
    End With

    Exit Sub

ErrorHandler_OpenFile:

    Call Lib.FatalError("Failed to open file")
    Lib.DisableOptimization
    End

End Sub

' Renaming workbooks based on date
Private Sub RenameWorkbook(ByRef Book As Workbook, ByRef Reference As String)

    On Error GoTo ErrorHandler_Rename

    Dim sYearToday As String: sYearToday = CStr(Year(Now))
    Dim sOldName   As String: sOldName = Book.Name
    Dim sOldRef    As String: sOldRef = Reference
    Dim iYear      As Long

    For iYear = Year(Now) To 2024 Step -1
        If InStr(sOldName, CStr(iYear)) Then

            sOldName = Replace(sOldName, CStr(iYear), sYearToday, 1, 1)
            Reference = Replace(sOldRef, CStr(iYear), sYearToday, 1, 1)

            Exit For

        End If
    Next iYear

    Exit Sub

ErrorHandler_Rename:

    Call Lib.FatalError("Failed to rename file")
    Lib.DisableOptimization
    End

End Sub

' Opening previous summary
Private Sub OpenPreviousCommon(Optional ByVal ReadOnly As Boolean = False)

    On Error GoTo ErrorHandler_OpenFile

    Dim sReference  As String
    Dim i           As Long
    
    For i = 13 To 2 Step -1
        If m_wsCommonPrimary.Cells(2, 35).Value Like m_wsMacrosReferences.Cells(i, 6).Value & "*" Then

            'NOTE -- We take the previous month because the report is generated for the previous month
            If Month(Now) = 1 Then
                sReference = Year(Now) & "-" & "12" & "\1\" & "*" & ".xlsx"
            ElseIf Month(Now) > 10 Then
                sReference = Year(Now) & "-" & Month(Now) - 1 & "\1\" & "*" & ".xlsx"
            Else
                sReference = Year(Now) & "-" & "0" & Month(Now) - 1 & "\1\" & "*" & ".xlsx"
            End If

            If Dir(sReference) <> "" Then
                Exit For
            Else
                Exit Sub
            End If

        End If
    Next i

    Set m_wbPreviousCommon = Workbooks.Open(Filename:=sReference, ReadOnly:=ReadOnly)
    Set m_wsPreviousCommonPrimary = m_wbPreviousCommon.Worksheets(1)

    Exit Sub

ErrorHandler_OpenFile:

    Call Lib.FatalError("Failed to open file")
    Lib.DisableOptimization
    End

End Sub

Public Sub CreateXLSXFile(sPath As String)
    
    On Error GoTo ErrorHandler_CreateFile
    
    ' Loading 5%
    With frmLoading
        .SetInfo "Creating .xlsx file..."
        .SetProgress 5
    End With
    
    Call Lib.Delay(2000)
    
    ' Generated name for the summary summary
    If Month(Now) > 9 Then
        m_sCommonBookName = "Common " & Month(Now) & "." & Year(Now)
    Else
        m_sCommonBookName = "Common 0" & Month(Now) & "." & Year(Now)
    End If
    
    ' Delete the existing summary summary if there is one
    If Lib.IsFileExists(sPath & "\" & m_sCommonBookName & ".xlsx") Then
        Call Kill(sPath & "\" & m_sCommonBookName & ".xlsx")
    End If
    
    ' Create a summary file
    Set m_wbCommon = Workbooks.Add: Set m_wsCommonPrimary = m_wbCommon.Worksheets(1)
        
    With m_wbCommon
        .Worksheets(1).Name = "Common"
        .SaveAs sPath & "\" & m_sCommonBookName & ".xlsx"
    End With
    
    Exit Sub
    
ErrorHandler_CreateFile:

    Call Lib.FatalError("Failed to create .xlsx file")
    Lib.DisableOptimization
    End
    
End Sub

' Creating table
Public Sub CreateCommonTable()
    
    On Error GoTo ErrorHandler_CreateCommonTable
    
    ' Loading 15%
    With frmLoading
        .SetInfo "Forming a table header..."
        .SetProgress 15
    End With
    
    Call Lib.Delay(2000)
    
    ' Number of links to summary files by region and link to the proforma table
    Dim nRegionalReferences As Long: nRegionalReferences = m_wsMacrosReferences.Cells(2, 1).End(xlDown).Row
    Dim sHeadersReference   As String: sHeadersReference = m_wsMacrosReferences.Cells(2, 3).Value
    
    ' For the first month
    Dim rngPeriod As Range
    
    'Book with template
    Dim wbHeaders As Workbook:  Set wbHeaders = Workbooks.Open(Filename:=sHeadersReference, ReadOnly:=True)
    Dim wsHeaders As Worksheet: Set wsHeaders = wbHeaders.Worksheets("Common")
    
    ' Removes filters, expands columns
    With wsHeaders
        .AutoFilterMode = False
        .Columns.Hidden = False
        .Rows.Hidden = False
    End With
    
    ' Headers width
    Dim nHeaderColumns As Long: nHeaderColumns = Lib.GetColumnsCount(Row:=3, Sheet:=wsHeaders)
    
    Dim i, j, k, iPasteRow As Long: iPasteRow = 5
    
    ' Years, months, template header and their initialization
    
    Dim arrYears(), arrMonth() As String
    Dim arrHeaders()           As String
    Dim arrRefs()              As String
    
    ReDim arrYears(2 To 5)
    ReDim arrMonth(2 To 13)
    ReDim arrHeaders(1 To nHeaderColumns)
    ReDim arrRefs(2 To nRegionalReferences)
    
    For i = 2 To 5
        arrYears(i) = m_wsMacrosReferences.Cells(i, 5).Value
    Next i
    
    For i = 2 To 13
        arrMonth(i) = m_wsMacrosReferences.Cells(i, 6).Value
    Next i
    
    For i = 1 To nHeaderColumns
        arrHeaders(i) = LCase(wsHeaders.Cells(3, i).Value)
    Next i
    
    ' Adds a header to the summary summary
    wsHeaders.Range("A1:" & "AQ4").Copy Destination:=m_wsCommonPrimary.Range("A1")
    
    ' Duplicate header periods
    ' NOTE -- minus 2 because we start from scratch - once; two - because we form for the last month
    For i = 0 To Month(Now) - 2
        With m_wsCommonPrimary
            .Range(.Cells(1, 35 + (i * 9)), .Cells(4, 43 + (i * 9))).Copy Destination:=.Range(.Cells(1, 35 + ((i + 1) * 9)), .Cells(4, 43 + ((i + 1) * 9)))
        End With
    Next i
    
    ' Formats header names
    i = 0
    For j = Month(Now) + 1 To 2 Step -1
    
        With m_wsCommonPrimary
            If Not arrMonth(j) = "December" Then
                .Cells(2, 35 + (i * 9)).Value = m_wsMacrosReferences.Cells(j, 6).Value & " " & Year(Now)
            Else
                .Cells(2, 35 + (i * 9)).Value = m_wsMacrosReferences.Cells(j, 6).Value & " " & Year(Now) - 1
            End If
        End With
        
        i = i + 1
        
    Next j
    
    Dim nCommonColumns    As Long: nCommonColumns = Lib.GetColumnsCount(Row:=3, Method:=mcByFiltering, Sheet:=m_wsCommonPrimary)
    Dim wbRegional        As Workbook
    Dim wsRegionalPrimary As Worksheet
    
    With m_wsCommonPrimary
        .Range(.Cells(5, 1), .Cells(1048576, nCommonColumns)).NumberFormat = "@"
    End With
    
    For i = 2 To nRegionalReferences
        
        ' Assign a link to the .xlsx file
        Dim sRef As String: sRef = m_wsMacrosReferences.Cells(i, 1).Value
        
        On Error Resume Next
        Set wbRegional = Workbooks(sRef)
        On Error GoTo ErrorHandler_CreateCommonTable
        
        If Not wbRegional Is Nothing Then
            If Not InStr(wbRegional.Name, CStr(Year(Now))) > 0 Then
                Call RenameWorkbook(wbRegional, sRef)
            End If
        End If
        
        arrRefs(i) = CStr(sRef)
        
    Next i
    
    Lib.AddLogNote "Form of common table"
    
    ' Forming a final summary
    For i = 2 To nRegionalReferences
        
        ' Open the .xlsx file using the link
        Application.DisplayAlerts = False
        Set wbRegional = Workbooks.Open(Filename:=arrRefs(i), ReadOnly:=True)
        Application.DisplayAlerts = True
        
        ' Loading +5% for each regional summary
        With frmLoading
            .SetInfo wbRegional.Name & "..."
            .IncreaseProgress 5
        End With
        
        Set wsRegionalPrimary = wbRegional.Worksheets("Common")
        
        ' Removes filters, expands columns
        With wsRegionalPrimary
            .AutoFilterMode = False
            .Columns.Hidden = False
        End With
        
        ' Number of columns and rows of a specific open file
        Dim nRows    As Long: nRows = Lib.GetRowsCount(Column:=1, Method:=mcByEnd, Sheet:=wsRegionalPrimary)
        Dim nColumns As Long: nColumns = Lib.GetColumnsCount(Row:=3, Method:=mcByFiltering, Sheet:=wsRegionalPrimary)
        
        ' Buffer variable
        Dim rngBuffer As Range
        
        Dim iYear, iMonth, iMonthOld As Long
        
        ' Inserting rows into the first range
        For j = 1 To nHeaderColumns - 9
            For k = 1 To nColumns
                If arrHeaders(j) = LCase(wsRegionalPrimary.Cells(3, k).Value) Then
                    
                    ' Copy-paste cells
                    
                    With wsRegionalPrimary
                        Set rngBuffer = .Range(.Cells(5, k), .Cells(nRows, k))
                    End With
                    
                    With m_wsCommonPrimary
                        
                        ' NOTE -- on the first run, all cell data is inserted
                        If i = 2 Then
Common_PasteAll:
                            rngBuffer.Copy
                            .Cells(iPasteRow, j).PasteSpecial xlPasteAll
                            Application.CutCopyMode = False
                            
                        Else
                            
                            ' In rare cases, there may be a buffer overflow, so you will have to insert the entire column along with cell formats
                            On Error GoTo Common_PasteAll
                            
                            .Range(.Cells(iPasteRow, j), .Cells(rngBuffer.Cells.Count + (iPasteRow - 1), j)).Value = rngBuffer.Value
                            
                        End If
                        
                    End With
                    
                    ' Return to default Error Handler
                    On Error GoTo ErrorHandler_CreateCommonTable
                    
                End If
            Next k
        Next j
        
        ' Select the first month from which the common table begins
        
        With wsRegionalPrimary
            Set rngPeriod = .Range(.Cells(2, 1), .Cells(2, nColumns))
        End With
        
        For Each rngBuffer In rngPeriod.Columns
            For iYear = Year(Now) To 2023 Step -1
                For iMonth = 13 To 2 Step -1
                    
                    Dim sDate   As String:  sDate = arrMonth(iMonth) & " " & iYear
                    Dim dt      As Date:    dt = CDate(sDate)
                    
                    If rngBuffer.Value = dt Then
                    
                        iMonth = Month(Now) - (iMonth - 1)
                        iMonthOld = iMonth
                        
                        GoTo BuildPeriods
                        
                    End If
                    
                Next iMonth
            Next iYear
        Next rngBuffer
        
BuildPeriods:
        
        ' Number of the 35th row in the open table
        Set rngBuffer = wsRegionalPrimary.Cells(4, 1).EntireRow.Find("35", LookIn:=xlValues, LookAt:=xlPart)
        Dim n As Long: n = rngBuffer.Column
        
        ' Inserting rows into period ranges
        For iMonth = iMonth To Month(Now)
            For j = 35 To 43
                For k = n + ((iMonth - iMonthOld) * 9) To (n + 8) + ((iMonth - iMonthOld) * 9)
                    If arrHeaders(j) = LCase(wsRegionalPrimary.Cells(3, k).Value) Then
                        
                        ' Copy-paste cells
                        
                        With wsRegionalPrimary
                            Set rngBuffer = .Range(.Cells(5, k), .Cells(nRows, k))
                        End With
                        
                        With m_wsCommonPrimary
                            
                            ' NOTE -- on the first run, all cell data is inserted
                            If i = 2 Then
Periods_PasteAll:
                                rngBuffer.Copy
                                .Cells(iPasteRow, j + (iMonth * 9)).PasteSpecial xlPasteAll
                                Application.CutCopyMode = False
                                
                            Else
                                
                                ' In rare cases, there may be a buffer overflow, so you will have to insert the entire column along with cell formats
                                On Error GoTo Periods_PasteAll
                                
                                .Range(.Cells(iPasteRow, j + (iMonth * 9)), .Cells(rngBuffer.Cells.Count + (iPasteRow - 1), j + (iMonth * 9))).Value = rngBuffer.Value
                                
                            End If
                            
                        End With
                        
                        ' Return to default Error Handler
                        On Error GoTo ErrorHandler_CreateCommonTable
                        
                        Exit For
                        
                    End If
                Next k
            Next j
        Next iMonth
        
        ' Move the place of our next copy-paste down one line
        iPasteRow = Lib.GetRowsCount(Column:=1, Sheet:=m_wsCommonPrimary) + 1
        
        wbRegional.Close SaveChanges:=False
        
        Call Lib.Delay(3000)
        
    Next i
    
    Lib.AddLogNote "Form of additional table"
    
    ' Number of lines in the final summary
    Dim nCommonRows As Long: nCommonRows = Lib.GetRowsCount(Column:=1, Sheet:=m_wsCommonPrimary)
    
    nCommonColumns = Lib.GetColumnsCount(Row:=3, Method:=mcByFiltering, Sheet:=m_wsCommonPrimary)
    
    ' Formats the table
    With m_wsCommonPrimary
    
        Call Lib.AcceptCellsBorders(TableRange:=.Range(.Cells(5, 1), .Cells(5, nCommonColumns)))
        
        .Range(.Cells(5, 1), .Cells(5, nCommonColumns)).Font.Size = 8
        .Range(.Cells(5, 1), .Cells(5, nCommonColumns)).Font.Name = "Arial"
        
        .Range(.Cells(5, 1), .Cells(5, nCommonColumns)).Copy
        .Range(.Cells(6, 1), .Cells(nCommonRows, nCommonColumns)).PasteSpecial xlPasteFormats
        
        .Range(.Cells(5, 1), .Cells(5, nCommonColumns)).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range(.Cells(5, nCommonColumns), .Cells(nCommonRows, nCommonColumns)).Borders(xlEdgeRight).LineStyle = xlContinuous
        
    End With
    
    ' Loading 99%
    With frmLoading
        .SetInfo "Formation of a comparative table..."
        .SetProgress 99
    End With
    
    Call Lib.Delay(2000)
    
    ' Sheet of Previous Common table
    Call OpenPreviousCommon(ReadOnly:=True)
    
    ' NOTE -- Common Table for the previous month may not have been generated by the time the macro is run
    If Not m_wbPreviousCommon Is Nothing Then
        
        With m_wbCommon
            .Worksheets.Add After:=m_wbCommon.Worksheets("Common")
            .Worksheets(2).Name = "for comparison"
        End With
        
        ' Comparison sheet
        Dim wsAdditional As Worksheet: Set wsAdditional = m_wbCommon.Worksheets(2)
        
        ' Form a table for comparison
        With wsAdditional
        
            Call Lib.AcceptCellsBorders(TableRange:=.Range("A1:C3"))
            
            .Cells(1, 2).Value = m_sCommonBookName
            
            If Month(Now) > 9 Then
                .Cells(1, 3).Value = Month(Now) & "." & Year(Now) & " COMMON TABLE"
            Else
                .Cells(1, 3).Value = "0" & Month(Now) & "." & Year(Now) & " COMMON TABLE"
            End If
            
            .Cells(2, 1).Value = "DZ Sum"
            .Cells(3, 1).Value = "DZ Reserve Sum"
            
            .Range("A1:C1").Interior.Color = mcGreenColor
            .Range("A2:A3").Interior.Color = mcLightGreyColor
            
            .Range("A1:C1").Font.Bold = True
            .Range("A1:A3").Font.Bold = True
            
            .Range("B2:C3").NumberFormat = "0.00"
            
            ' Inserts values
            
            .Cells(2, 2).Value = m_wsPreviousCommonPrimary.Cells(1, 36).Value
            .Cells(3, 2).Value = m_wsPreviousCommonPrimary.Cells(1, 41).Value
            
            .Cells(2, 3).Value = m_wsCommonPrimary.Cells(1, 36).Value
            .Cells(3, 3).Value = m_wsCommonPrimary.Cells(1, 41).Value
            
            .Range("A:C").Columns.AutoFit
            
        End With
        
        With m_wsCommonPrimary
            .Activate
            .Range("A1").Select
            .Rows(4).AutoFilter
        End With
        
        wsAdditional.Activate: wsAdditional.Range("A1").Select
        
        m_wbPreviousCommon.Close SaveChanges:=False
        
    Else
        With m_wsCommonPrimary
            .Activate
            .Range("A1").Select
            .Rows(4).AutoFilter
        End With
    End If
    
    Application.CutCopyMode = False
    
    wbHeaders.Close SaveChanges:=False
    
    Lib.AddLogNote "Forming has finished"
    
    Exit Sub
    
ErrorHandler_CreateCommonTable:

    Call Lib.FatalError("Failed to generate summary summary")
    Lib.DisableOptimization
    End
    
End Sub

Public Sub UpdateRegionalTable()
    
    On Error GoTo ErrorHandler_UpdateRegionalTable
    
    Lib.AddLogNote "Updating a regional table"
    
    Dim wbRegional        As Workbook: Set wbRegional = Workbooks.Open(Filename:=frmMain.txtRegionalPath.Value, ReadOnly:=False)
    Dim wsRegionalPrimary As Worksheet: Set wsRegionalPrimary = wbRegional.Worksheets("Common")
    
    Dim nRegionalRows As Long: nRegionalRows = wsRegionalPrimary.Cells(wsRegionalPrimary.Rows.Count, 1).End(xlUp).Row
    Dim i, j, k       As Long
    
    Dim colFilteredIndices As New Collection
    Dim arrRegional()      As Variant: ReDim arrRegional(5 To nRegionalRows)
    
    Dim dictForm      As Object: Set dictForm = CreateObject("Scripting.Dictionary")
    Dim dictASR       As Object: Set dictASR = CreateObject("Scripting.Dictionary")
    Dim dictSplitting As Object: Set dictSplitting = CreateObject("Scripting.Dictionary")
    
    ' Check using Form 7.4
    If Not frmMain.txtFormPath.Value = "" Then
        
        ' Loading 15%
        With frmLoading
            .SetInfo "Check using Form 7.4..."
            .SetProgress 15
        End With
        
        Call Lib.Delay(2000)
        
        With wsRegionalPrimary
            .AutoFilterMode = False
            .Range("A:AA").AutoFilter Field:=6, Criteria1:="OEBS"
        End With
        
        ' Initializes an array of transactions
        For i = 5 To nRegionalRows
            
            arrRegional(i) = CStr(wsRegionalPrimary.Cells(i, 1).Value)
            
            If wsRegionalPrimary.Range("A" & i).EntireRow.Hidden = False Then
                
                colFilteredIndices.Add i
                
                If wsRegionalPrimary.Cells(i, 9).Value = "OK" Then
                    With wsRegionalPrimary
                        .Range(.Cells(i, 9), .Cells(i, 9)).Interior.Color = mcYellowColor
                    End With
                End If
                
            End If
            
        Next i
        
        Set m_wbForm = Workbooks.Open(Filename:=CStr(frmMain.txtFormPath.Value), ReadOnly:=True)
        Set m_wsFormPrimary = m_wbForm.Worksheets(1)
        
        Dim nFormRows As Long:    nFormRows = Lib.GetRowsCount(Column:=1, Sheet:=m_wsFormPrimary)
        Dim arrForm() As Variant: ReDim arrForm(2 To nFormRows)
        
        For i = 2 To nFormRows
            arrForm(i) = CStr(m_wsFormPrimary.Cells(i, 7).Value)
        Next i
        
        For i = 2 To nFormRows
            If Not dictForm.Exists(arrForm(i)) Then
                dictForm.Add arrForm(i), True
            End If
        Next i
        
        For Each i In colFilteredIndices
            If Not dictForm.Exists(arrRegional(i)) Then
                With wsRegionalPrimary
                    .Range(.Cells(i, 9), .Cells(i, 9)).Interior.Color = mcWhiteColor
                    .Cells(i, 9).Value = "OK"
                End With
            End If
        Next
        
        Set colFilteredIndices = Nothing
        
        m_wbForm.Close SaveChanges:=False
        
    End If
    
    ' Check according to ACP Calculations
    If Not frmMain.txtASRPath.Value = "" Then
        
        ' Loading 45%
        With frmLoading
            .SetInfo "Checking against ASR Calculations..."
            .SetProgress 45
        End With
        
        Call Lib.Delay(2000)
        
        With wsRegionalPrimary
            .AutoFilterMode = False
            .Range("A:AA").AutoFilter Field:=6, Criteria1:="ASR Repayments"
        End With
        
        ' Initializes an array of transactions
        For i = 5 To nRegionalRows
            
            arrRegional(i) = CStr(wsRegionalPrimary.Cells(i, 1).Value)
            
            If wsRegionalPrimary.Range("A" & i).EntireRow.Hidden = False Then
                
                colFilteredIndices.Add i
            
                If wsRegionalPrimary.Cells(i, 9).Value = "OK" Then
                    With wsRegionalPrimary
                        .Range(.Cells(i, 9), .Cells(i, 9)).Interior.Color = mcYellowColor
                    End With
                End If
                
            End If
            
        Next i
        
        Set m_wbASR = Workbooks.Open(Filename:=CStr(frmMain.txtASRPath.Value), ReadOnly:=True)
        Set m_wsASRPrimary = m_wbASR.Worksheets(1)
        
        Dim nASRRows   As Long:    nASRRows = Lib.GetRowsCount(Column:=1, Sheet:=m_wsASRPrimary)
        Dim arrASR()   As Variant: ReDim arrASR(2 To nASRRows)
        
        For i = 2 To nASRRows
            arrASR(i) = CStr(m_wsASRPrimary.Cells(i, 4).Value)
        Next i
        
        For i = 2 To nASRRows
            If Not dictASR.Exists(arrASR(i)) Then
                dictASR.Add arrASR(i), True
            End If
        Next i
        
        For Each i In colFilteredIndices
            If Not dictASR.Exists(arrRegional(i)) Then
                With wsRegionalPrimary
                    .Range(.Cells(i, 9), .Cells(i, 9)).Interior.Color = mcWhiteColor
                    .Cells(i, 9).Value = "OK"
                End With
            End If
        Next
        
        Set colFilteredIndices = Nothing
        
        m_wbASR.Close SaveChanges:=False
        
    End If
    
    ' Check for separate prepayment
    If Not frmMain.txtSplittingPath.Value = "" Then
        
        ' Loading 75%
        With frmLoading
            .SetInfo "Checking for separate prepayment..."
            .SetProgress 75
        End With
        
        Call Lib.Delay(2000)
        
        With wsRegionalPrimary
            .AutoFilterMode = False
            .Range("A:AA").AutoFilter Field:=6, Criteria1:="Creditors"
        End With
        
        ' Initializes an array of transactions
        For i = 5 To nRegionalRows
        
            arrRegional(i) = CStr(wsRegionalPrimary.Cells(i, 1).Value)
            
            If wsRegionalPrimary.Range("A" & i).EntireRow.Hidden = False Then
                
                colFilteredIndices.Add i
            
                If wsRegionalPrimary.Cells(i, 9).Value = "OK" Then
                    With wsRegionalPrimary
                        .Range(.Cells(i, 9), .Cells(i, 9)).Interior.Color = mcYellowColor
                    End With
                End If
            
            End If
            
        Next i
        
        Set m_wbSplitting = Workbooks.Open(Filename:=CStr(frmMain.txtSplittingPath.Value), ReadOnly:=True)
        Set m_wsSplittingPrimary = m_wbSplitting.Worksheets(1)
        
        Dim nSplittingRows   As Long:    nSplittingRows = Lib.GetRowsCount(Column:=1, Sheet:=m_wsSplittingPrimary)
        Dim arrSplitting()   As Variant: ReDim arrSplitting(14 To nSplittingRows)
        
        For i = 14 To nSplittingRows
            arrSplitting(i) = CStr(m_wsSplittingPrimary.Cells(i, 7).Value)
        Next i
        
        For i = 14 To nSplittingRows
            If Not dictSplitting.Exists(arrSplitting(i)) Then
                dictSplitting.Add arrSplitting(i), True
            End If
        Next i
        
        For Each i In colFilteredIndices
            If Not dictSplitting.Exists(arrRegional(i)) Then
                With wsRegionalPrimary
                    .Range(.Cells(i, 9), .Cells(i, 9)).Interior.Color = mcWhiteColor
                    .Cells(i, 9).Value = "OK"
                End With
            End If
        Next
        
        Set colFilteredIndices = Nothing
        
        m_wbSplitting.Close SaveChanges:=False
        
    End If
    
    ' Loading 99%
    With frmLoading
    	.SetInfo ""
        .SetProgress 99
    End With
        
    Call Lib.Delay(2000)
    
    Application.CutCopyMode = False
    wsRegionalPrimary.AutoFilterMode = False
    
    ' Resets the selected range after the code runs
    With wsRegionalPrimary
        .Activate
        .Range("A1").Select
        .Rows(4).AutoFilter
    End With
    
    Lib.AddLogNote "Updating has finished"
    
    Exit Sub
    
ErrorHandler_UpdateRegionalTable:

    Call Lib.FatalError("Failed to update the summary file for the previous period")
    Lib.DisableOptimization
    End
    
End Sub
