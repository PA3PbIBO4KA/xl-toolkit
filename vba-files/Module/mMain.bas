Attribute VB_Name = "mMain"

' Name: mMain
' Author: ������ ������
' Date: 13.06.2024

Option Explicit

Global Books As CBooksBehavior

' �������������
Public Sub Init()
    Set Books = New CBooksBehavior
End Sub

' ����� ����� � ������
Sub Launch()
Attribute Launch.VB_ProcData.VB_Invoke_Func = "m\n14"
    
    mLibXL.Init
    mMain.Init
    
    frmMain.Show
    
End Sub
