VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCommonLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' Name: cCommonLib
' Author: Mikhail Krasyuk
' Date: 13.06.2024

Option Explicit

Private Const m_LogFilesPath As String = ""

Private m_LogFile    As Object
Private m_LogFolder  As Object

Private m_asFileLine() As String
Private m_sFolderPath  As String
Private m_sFileName    As String

Private m_nFileLines As Long

' Êîíñòðóêòîð êëàññà
Private Sub Class_Initialize()
    
    m_sFolderPath = m_LogFilesPath & Application.UserName & mcBackslash
    m_sFileName = ""
    
    ReDim m_asFileLine(1 To 1000)
    
    Set m_LogFile = Nothing
    Set m_LogFolder = Nothing
    
    m_nFileLines = 1
    
End Sub

' Äëÿ ëîã ôàéëîâ è Debug.Print
Private Function Today() As String
    Today = "[" & Now & "]: "
End Function

' Ñîçäàíèå ëîã ôàéëà. Âûçûâàåòñÿ â êîíöå ïðîãðàììû.
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

' Äîáàâëÿåò ñòðî÷êó ñ òåêñòîì â áóäóùèé ëîã ôàéë
Public Sub AddLogNote(ByVal Text As String)
    m_asFileLine(m_nFileLines) = Today & Text
    m_nFileLines = m_nFileLines + 1
End Sub

' Îòêëþ÷åíèå UI âî âðåìÿ âûïîëíåíèÿ ïðîãðàììû. Îïòèìèçèðóåò ïðîöåññ âûïîëíåíèÿ ìàêðîñà
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

' Âêëþ÷åíèå UI ïîñëå âûïîëíåíèÿ ïðîãðàììû. Âîçâðàò ê íàñòðîéêàì ïî-óìîë÷àíèþ
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

' Âîçâðàùàåò êîë-âî ñòðîê â òàáëèöå
Public Function GetRowsCount(Optional Column As Long = 1, Optional Method As Long = mcByFiltering, Optional Sheet As Worksheet = Nothing) As Long
    
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

' Âîçâðàùàåò êîë-âî ñòîëáöîâ â òàáëèöå
Public Function GetColumnsCount(Optional Row As Long = 1, Optional Method As Long = mcByFiltering, Optional Sheet As Worksheet = Nothing) As Long
    
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

Public Function IsInArray(Value As String, arr As Variant) As Boolean
    
    Dim i
    
    For i = LBound(arr) To UBound(arr)
        If arr(i) = Value Then
            IsInArray = True
            Exit Function
        End If
    Next i
    
    IsInArray = False

End Function

Public Function IsValueInCollection(Collection As Collection, Value As Variant) As Boolean

    Dim Item As Variant
    
    For Each Item In Collection
        If Item = Value Then
            IsValueInCollection = True
            Exit Function
        End If
    Next Item
    
    IsValueInCollection = False
    
End Function

' Ïðîâåðêà îòêðûò ëè ñîîòâåòñòâóþùèé .xlsx ôàéë
Public Function IsWorkbookOpen(Filename As String) As Boolean

    Dim iFF As Integer
    iFF = FreeFile
    
    On Error Resume Next
    Open Filename For Random Access Read Write Lock Read Write As #iFF
    Close #iFF
    
    IsWorkbookOpen = Err
    
End Function

' Ïðîâåðêà ñóùåñòâóåò ëè ôàéë â ïàïêå
Function IsFileExists(Filename As String) As Boolean
    If FileSystem.FileExists(Filename) Then
        IsFileExists = True
    Else
        IsFileExists = False
    End If
End Function

' Ïðîâåðêà ñóùåñòâóåò ëè ëèñò â .xlsx ôàéëå
Public Function IsSheetExists(ByVal SheetName As String, Optional Workbook As Workbook = Nothing) As Boolean

    On Error Resume Next
    
    If Workbook Is Nothing Then
        Workbook = ActiveWorkbook
    End If
    
    IsSheetExists = (Not Workbook.Sheets(SheetName) Is Nothing)
    
    On Error GoTo 0
    
End Function

' Ñòàâèò íà ïàóçó ìàêðîñ
Public Sub Delay(Optional ByVal Milliseconds As Long = 5000)
    
    If Milliseconds = 0 Then Exit Sub
    
    Application.Wait Now + (Milliseconds * mcMillisecond)
    
End Sub

' Øàáëîí äëÿ âûáîðà .xlsx ôàéë ïîëüçîâàòåëåì ÷åðåç UI
Public Function ChooseXLSXFile() As String
    
    Dim sTitle As String: sTitle = "Âûáåðåòå ôàéë Excel"

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

' Çàêðûâàåò âñå .xlsx ôàéëû çàäàíûå â ìàññèâå
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

' Ïðåîáðàçóåò íîìåð ñòîëáöà â áóêâó
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

' Âûäåëÿåò ãðàíèöû çàäàííîãî äèàïîçîíà
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

' Âûõîä èç .xlsm ôàéëà ïîñëå âûïîëíåíèÿ ìàêðîñà
Public Sub ExitMacros()
    Application.OnTime Now + TimeValue("00:00:01"), "CloseMacrosWorkbook"
End Sub

' Îáðàáîò÷èê îøèáîê
Public Sub FatalError(ByVal Explanation As String)
    
    ' Èíèöèàëèçàöèÿ êðàø ðåïîðòåðà
    mcErrorExplanation = Explanation
    frmCrashReporter.Show
    
    Rem !_ÑÒÀÐÛÉ ÂÀÐÈÀÍÒ_! Rem
    'MsgBox Title:="Fatal Error", _
    '       Prompt:=Explanation & vbCrLf & vbCrLf & _
    '               "Error code: " & Err.Number & vbCrLf & _
    '               "Decs: " & Err.Description & vbCrLf & _
    '               "Src: " & Err.Source
    
    ' Ñîçäàíèå ëîã ôàéëà
    With Me
        .AddLogNote "Error: " & Err.Description
        .CreateLogFile
    End With
    
    End
    
End Sub

