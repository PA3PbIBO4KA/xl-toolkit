VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "Common Builder"
   ClientHeight    =   7875
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10155
   OleObjectBlob   =   "frmMain.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Name: frmMain
' Author: Mikhail Krasyuk
' Date: 16.07.2024

Option Explicit

' Initialization
Private Sub UserForm_Initialize()
    
    On Error GoTo ErrorHandler_Initialize
    
    Dim sFolderPath As String
    Dim sFileName   As String
    Dim file        As Object
    
    ' The path should be written here
    sFolderPath = ""
    
    For Each file In CreateObject("Scripting.FileSystemObject").GetFolder(sFolderPath).Files
        
        sFileName = file.Name
        
        If sFileName Like "Form 7.4" & "*" Then
            txtFormPath.Value = sFolderPath & sFileName
            GoTo NextFile
        End If
        
        If sFileName Like "ASR Calculations" & "*" Then
            txtASRPath.Value = sFolderPath & sFileName
            GoTo NextFile
        End If
        
        If sFileName Like "Prepayment Split" & "*" Then
            txtSplittingPath.Value = sFolderPath & sFileName
            GoTo NextFile
        End If
        
NextFile:
    Next file
    
    Exit Sub
    
ErrorHandler_Initialize:

    Call Lib.FatalError("Failed to initialize frmMain")
    Lib.DisableOptimization
    End
    
End Sub

' Selecting a folder for the new common table
Private Sub btnChooseFolder_Click()
    
    On Error GoTo ErrorHandler_Click
    
    Dim sFolderName As String
    
    With Application.FileDialog(msoFileDialogFolderPicker)
    
        .Title = "Select settings"
        .Show
        
        If .SelectedItems.Count > 0 Then
            sFolderName = .SelectedItems(1)
        Else
            Exit Sub
        End If
        
    End With
    
    txtFolderPath.Value = FileSystem.GetFolder(sFolderName).Path
    
    Exit Sub
    
ErrorHandler_Click:

    Call Lib.FatalError("Failed to process button press")
    Lib.DisableOptimization
    End
    
End Sub

' Next comes the selection of files to check the repayment through the explorer

Private Sub btnChooseRegionalFile_Click()
    
    On Error Resume Next
    txtRegionalPath.Value = Lib.ChooseXLSXFile()
    
End Sub

Private Sub btnChooseFormFile_Click()
    
    On Error Resume Next
    txtFormPath.Value = Lib.ChooseXLSXFile()
    
End Sub

Private Sub btnChooseASRFile_Click()
    
    On Error Resume Next
    txtASRPath.Value = Lib.ChooseXLSXFile()
    
End Sub

Private Sub btnChooseSplittingFile_Click()
    
    On Error Resume Next
    txtSplittingPath.Value = Lib.ChooseXLSXFile()
    
End Sub

' Creating .xlsx file for a new common table
Private Sub btnCreateFile_Click()
    
    On Error GoTo ErrorHandler_Click
    
    If txtFolderPath.Value = "" Then
        Exit Sub
    End If
    
    frmLoading.Show
    
    Call Lib.Delay(1500)
    
    Lib.EnableOptimization
    Call Books.CreateXLSXFile(sPath:=txtFolderPath.Value)
    Call Books.CreateCommonTable
    Lib.DisableOptimization
    
    MsgBox "Common table is built"
    
    With Lib
        .AddLogNote "Macros Closed"
        .CreateLogFile
    End With
    
    End
    
ErrorHandler_Click:

    Call Lib.FatalError("Failed to process button press")
    Lib.DisableOptimization
    End
    
End Sub

' Checking repayments
Private Sub btnCheckRepayment_Click()
    
    On Error GoTo ErrorHandler_Click
    
    If txtRegionalPath.Value = "" Then
        Exit Sub
    End If
    
    frmLoading.Show
    
    Call Lib.Delay(1500)
    
    Lib.EnableOptimization
    Call Books.UpdateRegionalTable
    Lib.DisableOptimization
       
    frmLoading.Hide
       
    MsgBox "Checking repayments is complete"
       
    With Lib
        .AddLogNote "Macros Closed"
        .CreateLogFile
    End With
    
    End
    
ErrorHandler_Click:

    Call Lib.FatalError("Failed to process button press")
    Lib.DisableOptimization
    End
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        
        'Cancel = True
        
        With Lib
            .AddLogNote "Macros Closed"
            .CreateLogFile
        End With
        
        End
        
    End If
End Sub
