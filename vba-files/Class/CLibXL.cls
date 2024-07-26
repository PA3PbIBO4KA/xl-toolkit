VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLibXL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' Name: CLibXL
' Author: Mikhail Krasyuk
' Date: 23.07.2024

Option Explicit

Private Const m_LogFilesPath As String = ""

Private m_LogFile    As Object
Private m_LogFolder  As Object

Private m_asFileLine() As String
Private m_sFolderPath  As String
Private m_sFileName    As String

Private m_nFileLines As Long

' Class constructor
Private Sub class_initialize()
    
    m_sFolderPath = m_LogFilesPath & Application.UserName & mcBackslash
    m_sFileName = ""
    
    ReDim m_asFileLine(1 To 1000)
    
    Set m_LogFile = Nothing
    Set m_LogFolder = Nothing
    
    m_nFileLines = 1
    
End Sub

' For log files and Debug.Print
Private Function Today() As String
    
    Today = "[" & Now & "]: "
    
End Function

' Create a log file. Called at the end of the program.
Public Sub CreateLogFile()
    
    If m_nFileLines = 1 Then
        Exit Sub
    End If
    
    Dim iLine As Long
    
    m_sFileName = "log " & Day(Now) & "." & Month(Now) & "." & Year(Now) & " " & Hour(Now) & "#" & Month(Now) & "#" & Second(Now) & ".txt"
        
    If Not FileSystem.FolderExists(m_sFolderPath) Then
        Call FileSystem.CreateFolder(m_sFolderPath)
    End If
    
    Set m_LogFile = FileSystem.CreateTextFile(m_sFolderPath & m_sFileName)
    
    For iLine = 1 To m_nFileLines
        m_LogFile.WriteLine m_asFileLine(iLine)
    Next iLine
    
    m_LogFile.Close
    
End Sub

' Adds a line of text to the future log file
Public Sub AddLogNote(ByVal Text As String)
    
    m_asFileLine(m_nFileLines) = Today & Text
    
    m_nFileLines = m_nFileLines + 1
    
End Sub

'Disables the UI while the program is running. Optimizes the macro execution process
Public Sub EnableOptimization()
    
    With Application
        .AskToUpdateLinks = False
        .EnableEvents = False
        .ScreenUpdating = False
        .EnableAnimations = False
        .DisplayStatusBar = False
        .PrintCommunication = False
    End With
    
End Sub

' Enable UI after program execution. Return to default settings
Public Sub DisableOptimization()

    With Application
        .AskToUpdateLinks = True
        .EnableEvents = True
        .ScreenUpdating = True
        .EnableAnimations = True
        .DisplayStatusBar = True
        .PrintCommunication = True
    End With
    
End Sub

' Returns the number of rows in the table
Public Function GetRowsCount(ByVal Column As Long, Optional Method As Long = mcByFiltering, Optional Sheet As Worksheet = Nothing) As Long
    
    If Sheet Is Nothing Then Set Sheet = ActiveSheet
    
    If (Method = mcByFiltering) Then
        GetRowsCount = Sheet.Columns(Column).Find(What:="*", LookIn:=xlValues, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    ElseIf (Method = mcByEnd) Then
        GetRowsCount = Sheet.Cells(Rows.Count, Column).End(xlUp).Row
    ElseIf (Method = mcByCycle) Then
        
        Dim i As Long
        
        For i = 1 To 1048576
            If Sheet.Cells(i, 1) = "" Then
                GetRowsCount = i - 1
                Exit For
            End If
        Next i
        
    End If
    
End Function

' Returns the number of columns in the table
Public Function GetColumnsCount(ByVal Row As Long, Optional Method As Long = mcByFiltering, Optional Sheet As Worksheet = Nothing) As Long
    
    If Sheet Is Nothing Then Set Sheet = ActiveSheet
    
    If (Method = mcByFiltering) Then
        GetColumnsCount = Sheet.Rows(Row).Find(What:="*", LookIn:=xlValues, SearchDirection:=xlPrevious, SearchOrder:=xlByColumns).Column
    ElseIf (Method = mcByEnd) Then
        GetColumnsCount = Sheet.Cells(Row, 1).End(xlToRight).Column
    ElseIf (Method = mcByCycle) Then
        
        Dim i As Long
        
        For i = 1 To 16384
            If Sheet.Cells(1, i) = "" Then
                GetColumnsCount = i - 1
                Exit For
            End If
        Next i
        
    End If
    
End Function

' Check that the corresponding .xlsx file is open
Public Function IsWorkbookOpen(Filename As String) As Boolean

    Dim iFF As Integer
    iFF = FreeFile
    
    On Error Resume Next
    Open Filename For Random Access Read Write Lock Read Write As #iFF
    Close #iFF
    
    IsWorkbookOpen = Err
    
End Function

' Checking if the corresponding .xlsx file is open
Function IsFileExists(Filename As String) As Boolean
    
    If FileSystem.FileExists(Filename) Then
        IsFileExists = True
    Else
        IsFileExists = False
    End If
    
End Function

' Checking whether a sheet exists in the .xlsx file
Public Function IsSheetExists(ByVal SheetName As String, Optional Workbook As Workbook = Nothing) As Boolean

    On Error Resume Next
    
    If Workbook Is Nothing Then
        Workbook = ActiveWorkbook
    End If
    
    IsSheetExists = (Not Workbook.Sheets(SheetName) Is Nothing)
    
    On Error GoTo 0
    
End Function

' Pauses the macro
Public Sub Delay(Optional ByVal Milliseconds As Long = 5000)
    
    If Milliseconds = 0 Then Exit Sub
    
    Application.Wait Now + (Milliseconds * mcMillisecond)
    
End Sub

' Template for user selection of .xlsx file via UI
Public Function ChooseXLSXFile() As String
    
    Dim sTitle As String: sTitle = "Select Excel file"

    Dim sFileName As String: sFileName = Application.GetOpenFilename("All files(*.*),*.*", 1, sTitle, , False)

    If VarType(sFileName) = vbBoolean Then
        Exit Function
    End If

    If Me.IsFileExists(CStr(sFileName)) Then
    Else
        Call Workbooks.Open(sFileName)
    End If

    ChooseXLSXFile = sFileName

End Function

' Closes all .xlsx files specified in the array
Public Sub CloseAllWorkbooks(ByRef ArrayWorkbooks() As Variant, Optional ByVal Saving As Boolean = False)
    
    Dim wb As Workbook
    Dim i As Long
    
    For i = LBound(ArrayWorkbooks) To UBound(ArrayWorkbooks)
    
        Set wb = ArrayWorkbooks(i)
        
        If IsWorkbookOpen(wb.Name) Then
            wb.Close SaveChanges:=Saving
        End If
        
    Next i
    
End Sub

' Converts the column number to a letter
Public Function GetColumnLetter(ByVal Column As Long) As String

    Dim sLetter As String
    
    If Column < 1 Then
        GetColumnLetter = ""
    Else
        While Column > 0
            If Column > 26 Then
                sLetter = Chr((Column Mod 26) + 64) & sLetter
                Column = Int(Column / 26) - 1
            Else
                sLetter = Chr((Column Mod 26) + 65) & sLetter
                Column = Int(Column / 26)
            End If
        Wend
    End If
    
    GetColumnLetter = sLetter
    
End Function

' Highlights the boundaries of a specified range
Public Sub AcceptCellsBorders(TableRange As Range)
    
    With TableRange
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    End With
    
End Sub

' Exit the .xlsm file after executing the macro
Public Sub ExitMacros()
    
    Application.OnTime Now + TimeValue("00:00:01"), "CloseMacrosWorkbook"
    
End Sub

' Error handler
Public Sub FatalError(ByVal Explanation As String)
    
    ' Initializing the crash reporter
    mcErrorExplanation = Explanation
    frmCrashReporter.Show
    
    Rem !_OLD VERSION_! Rem
    'MsgBox Title:="Fatal Error", _
    '       Prompt:=Explanation & vbCrLf & vbCrLf & _
    '               "Error code: " & Err.Number & vbCrLf & _
    '               "Decs: " & Err.Description & vbCrLf & _
    '               "Src: " & Err.Source
    
    ' Creating a log file
    With Me
        .AddLogNote "Error: " & Err.Description
        .CreateLogFile
    End With
    
    End
    
End Sub

