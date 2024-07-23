VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLoading 
   Caption         =   "Загрузка"
   ClientHeight    =   2400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7005
   OleObjectBlob   =   "frmLoading.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Name: frmLoading
' Author: Михаил Красюк
' Date: 16.07.2024

Option Explicit

Private m_fScale      As Double

Private m_iPercentage As Long
Private m_nSymbols    As Long

Private m_sLoadingBar As String

' Инициализация
Private Sub UserForm_Initialize()
    
    On Error GoTo ErrorHandler_Initialize
    
    frmMain.Hide
    
    m_fScale = 1.07
    
    m_iPercentage = 0
    m_nSymbols = 1
    
    m_sLoadingBar = "|"
    lblPercentage = "1%"
    lblInfo = ""
    
    Exit Sub
    
ErrorHandler_Initialize:

    Call Lib.FatalError("Не удалось инициализировать frmLoading")
    Lib.DisableOptimization
    End
    
End Sub

Public Sub SetInfo(ByVal Text As String)
    
    On Error GoTo ErrorHandler_SetInfo
    
    lblInfo = Text
    
    Exit Sub
    
ErrorHandler_SetInfo:
    
    Call Lib.FatalError("Не удалось изменить надпись на экране загрузки")
    Lib.DisableOptimization
    End
    
End Sub

Public Sub SetProgress(ByVal Percentage As Long)
    
    On Error GoTo ErrorHandler_SetProgress
    
    ' Реинициализация
    Dim i As Long: m_sLoadingBar = ""
    
    ' Обновление приватных переменных
    m_iPercentage = Percentage
    m_nSymbols = Round(m_iPercentage * m_fScale)
    
    ' Заполнение лоадинг бара
    For i = 1 To m_nSymbols
        m_sLoadingBar = m_sLoadingBar + "|"
    Next i
    
    ' Обновление интерфейса
    lblPercentage = CStr(m_iPercentage & "%")
    lblLoadingBar = m_sLoadingBar
    
    Exit Sub
    
ErrorHandler_SetProgress:

    Call Lib.FatalError("Не удалось изменить Loading Bar")
    Lib.DisableOptimization
    End
    
End Sub

Public Sub IncreaseProgress(Optional ByVal Percentage As Long = 10)
    
    On Error GoTo ErrorHandler_IncreaseProgress
    
    ' Реинициализация
    Dim i As Long: m_sLoadingBar = ""
    
    ' Обновление приватных переменных
    m_iPercentage = m_iPercentage + Percentage
    m_nSymbols = Round(m_iPercentage * m_fScale)
    
    ' Заполнение лоадинг бара
    For i = 1 To m_nSymbols
        m_sLoadingBar = m_sLoadingBar + "|"
    Next i
    
    ' Обновление интерфейса
    lblPercentage = CStr(m_iPercentage & "%")
    lblLoadingBar = m_sLoadingBar
    
    Exit Sub
    
ErrorHandler_IncreaseProgress:

    Call Lib.FatalError("Не удалось изменить Loading Bar")
    Lib.DisableOptimization
    End
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        
    End If
End Sub
