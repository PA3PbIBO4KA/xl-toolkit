VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSettings 
   Caption         =   "Íàñòðîéêè"
   ClientHeight    =   2310
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7485
   OleObjectBlob   =   "frmSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Name: frmSettings
' Author: Mikhail Krasyuk
' Date: 25.10.2024
' Update: 04.11.2024

Option Explicit

Private pFileSystem As Object
Private pRootPath As String

Private Sub UserForm_Initialize()
    Set pFileSystem = CreateObject("Scripting.FileSystemObject")
    Let Me.txtFolderPath.Value = ThisWorkbook.Worksheets(1).Range("A2").Value
    Let pRootPath = Me.txtFolderPath.Value
End Sub

Private Sub btnChooseFolder_Click()
    
    Dim FolderName As String
    
    With Application.FileDialog(msoFileDialogFolderPicker)
    
        .Title = "Âûáåðåòå ïàïêó"
        .Show
        
        If .SelectedItems.Count > 0 Then
            Let FolderName = .SelectedItems(1)
        Else
            Exit Sub
        End If
        
    End With
    
    Let txtFolderPath.Value = pFileSystem.GetFolder(FolderName).Path
    
End Sub

Public Sub Cancel()
    
    Let Me.txtFolderPath.Value = pRootPath
    
    Me.Hide
    
End Sub

Private Sub btnAccept_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    With Me.btnAccept
        .BackColor = frmPrimary.InteractionButtonColor
        .BorderColor = frmPrimary.InteractionButtonBorderColor
    End With
End Sub

Private Sub btnCancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    With Me.btnCancel
        .BackColor = frmPrimary.InteractionButtonColor
        .BorderColor = frmPrimary.InteractionButtonBorderColor
    End With
End Sub

Private Sub framePrimary_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    With Me.btnAccept
        .BackColor = frmPrimary.DefaultButtonColor
        .BorderColor = frmPrimary.DefaultButtonBorderColor
    End With
    
    With Me.btnCancel
        .BackColor = frmPrimary.DefaultButtonColor
        .BorderColor = frmPrimary.DefaultButtonBorderColor
    End With
    
End Sub

Private Sub btnCancel_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call Me.Cancel
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
    End If
End Sub
