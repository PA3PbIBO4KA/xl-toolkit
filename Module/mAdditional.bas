Attribute VB_Name = "mAdditional"

' Name: mAdditional
' Author: Mikhail Krasyuk
' Date: 06.11.2024
' Update: 06.11.2024

Option Explicit

Global Const mcFully = 1
Global Const mcPartly = 2

Public Function IsInArray(Value As String, arr As Variant, Mode As Long) As Boolean
    
    Dim i
    
    If Mode = mcFully Then
        For i = LBound(arr) To UBound(arr)
            If Value = arr(i) Then
                IsInArray = True
                Exit Function
            End If
        Next i
    ElseIf Mode = mcPartly Then
        For i = LBound(arr) To UBound(arr)
            If Value Like "*" & arr(i) & "*" Then
                IsInArray = True
                Exit Function
            End If
        Next i
    End If
    
    IsInArray = False

End Function
