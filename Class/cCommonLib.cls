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

' ����������� ������
Private Sub Class_Initialize()
    
    m_sFolderPath = m_LogFilesPath & Application.UserName & mcBackslash
    m_sFileName = ""
    
    ReDim m_asFileLine(1 To 1000)
    
    Set m_LogFile = Nothing
    Set m_LogFolder = Nothing
    
    m_nFileLines = 1
    
End Sub

' ��� ��� ������ � Debug.Print
Private Function Today() As String
    Today = "[" & Now & "]: "
End Function

' �������� ��� �����. ���������� � ����� ���������.
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

' ��������� ������� � ������� � ������� ��� ����
Public Sub AddLogNote(ByVal Text As String)
    m_asFileLine(m_nFileLines) = Today & Text
    m_nFileLines = m_nFileLines + 1
End Sub

' ���������� UI �� ����� ���������� ���������. ������������ ������� ���������� �������
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

' ��������� UI ����� ���������� ���������. ������� � ���������� ��-���������
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

' ���������� ���-�� ����� � �������
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

' ���������� ���-�� �������� � �������
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

' �������� ������ �� ��������������� .xlsx ����
Public Function IsWorkbookOpen(Filename As String) As Boolean

    Dim iFF As Integer
    iFF = FreeFile
    
    On Error Resume Next
    Open Filename For Random Access Read Write Lock Read Write As #iFF
    Close #iFF
    
    IsWorkbookOpen = Err
    
End Function

' �������� ���������� �� ���� � �����
Function IsFileExists(Filename As String) As Boolean
    If FileSystem.FileExists(Filename) Then
        IsFileExists = True
    Else
        IsFileExists = False
    End If
End Function

' �������� ���������� �� ���� � .xlsx �����
Public Function IsSheetExists(ByVal SheetName As String, Optional Workbook As Workbook = Nothing) As Boolean

    On Error Resume Next
    
    If Workbook Is Nothing Then
        Workbook = ActiveWorkbook
    End If
    
    IsSheetExists = (Not Workbook.Sheets(SheetName) Is Nothing)
    
    On Error GoTo 0
    
End Function

' ������ �� ����� ������
Public Sub Delay(Optional ByVal Milliseconds As Long = 5000)
    
    If Milliseconds = 0 Then Exit Sub
    
    Application.Wait Now + (Milliseconds * mcMillisecond)
    
End Sub

' ������ ��� ������ .xlsx ���� ������������� ����� UI
Public Function ChooseXLSXFile() As String
    
    Dim sTitle As String: sTitle = "�������� ���� Excel"

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

' ��������� ��� .xlsx ����� ������� � �������
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

' ����������� ����� ������� � �����
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

' �������� ������� ��������� ���������
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

' ����� �� .xlsm ����� ����� ���������� �������
Public Sub ExitMacros()
    Application.OnTime Now + TimeValue("00:00:01"), "CloseMacrosWorkbook"
End Sub

' ���������� ������
Public Sub FatalError(ByVal Explanation As String)
    
    ' ������������� ���� ���������
    mcErrorExplanation = Explanation
    frmCrashReporter.Show
    
    Rem !_������ �������_! Rem
    'MsgBox Title:="Fatal Error", _
    '       Prompt:=Explanation & vbCrLf & vbCrLf & _
    '               "Error code: " & Err.Number & vbCrLf & _
    '               "Decs: " & Err.Description & vbCrLf & _
    '               "Src: " & Err.Source
    
    ' �������� ��� �����
    With Me
        .AddLogNote "Error: " & Err.Description
        .CreateLogFile
    End With
    
    End
    
End Sub
