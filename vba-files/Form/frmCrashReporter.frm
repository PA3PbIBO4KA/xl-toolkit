VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCrashReporter 
   Caption         =   "Критическая ошибка"
   ClientHeight    =   5550
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7620
   OleObjectBlob   =   "frmCrashReporter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCrashReporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Name: frmCrashReporter
' Author: Михаил Красюк
' Date: 16.07.2024

Option Explicit

' Инициализация
Private Sub UserForm_Initialize()
    
    lblExplanation.Caption = "Explanation: " & mcErrorExplanation
    lblErrorCode.Caption = "Error Code: " & Err.Number
    lblSource.Caption = "Source: " & Err.Source
    lblDescription.Caption = "Description: " & Err.Description
    lblFile.Caption = "File: " & Err.HelpFile
    lblLine.Caption = "Line: " & Erl
    
End Sub

Private Sub btnClose_Click()
    frmCrashReporter.Hide
End Sub

Private Sub btnEnableDebugMode_Click()
    mLibXL.EnableDebugMode
End Sub
