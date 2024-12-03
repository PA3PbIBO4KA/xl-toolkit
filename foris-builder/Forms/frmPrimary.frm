VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPrimary 
   Caption         =   "Ñáîðùèê FORIS"
   ClientHeight    =   2940
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9900.001
   OleObjectBlob   =   "frmPrimary.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPrimary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Name: frmPrimary
' Author: Mikhail Krasyuk
' Date: 02.10.2024
' Update: 24.10.2024

Option Explicit

Private pFORISBuilder As cFORISBuilder
Private pFileSystem As Object

Public DefaultButtonColor As Long
Public DefaultButtonBorderColor As Long
Public InteractionButtonColor As Long
Public InteractionButtonBorderColor As Long

Private Sub UserForm_Initialize()
    
    Set pFileSystem = CreateObject("Scripting.FileSystemObject")
    
    Let DefaultButtonColor = RGB(255, 255, 255)
    Let DefaultButtonBorderColor = RGB(169, 169, 169)
    
    Let InteractionButtonColor = RGB(211, 240, 224)
    Let InteractionButtonBorderColor = RGB(134, 191, 160)
    
End Sub

Private Sub btnCreateFile_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    With Me.btnCreateFile
        .BackColor = InteractionButtonColor
        .BorderColor = InteractionButtonBorderColor
    End With
End Sub

Private Sub btnSettings_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    With Me.btnSettings
        .BackColor = InteractionButtonColor
        .BorderColor = InteractionButtonBorderColor
    End With
End Sub

Private Sub frameFORIS_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    With Me.btnCreateFile
        .BackColor = DefaultButtonColor
        .BorderColor = DefaultButtonBorderColor
    End With
    
    With Me.btnSettings
        .BackColor = DefaultButtonColor
        .BorderColor = DefaultButtonBorderColor
    End With
    
End Sub

Private Sub framePrimary_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
     With Me.btnCreateFile
        .BackColor = DefaultButtonColor
        .BorderColor = DefaultButtonBorderColor
    End With
    
    With Me.btnSettings
        .BackColor = DefaultButtonColor
        .BorderColor = DefaultButtonBorderColor
    End With
    
End Sub

Private Sub btnSettings_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    frmSettings.Show
End Sub

Private Sub btnCreateFile_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    With Me.btnCreateFile
        .BackColor = InteractionButtonColor
        .BorderColor = InteractionButtonBorderColor
    End With
    
    Lib.EnableOptimization
    
    Set pFORISBuilder = New cFORISBuilder
    
    With pFORISBuilder
        .XLSFilesFolderPath = frmSettings.txtFolderPath.Value
        .MSISDNs = Me.txtMSISDNs.Value
        .CreateTable
    End With
    
    Lib.DisableOptimization
    Lib.ExitMacros
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        
        'Cancel = True
        
        Lib.ExitMacros
        
        End
        
    End If
End Sub
