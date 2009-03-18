Attribute VB_Name = "modMain"
Public Type SINITCOMMONCONTROLSEX
   dwSize  As Long  ' - Размер структуры
   dwICC   As Long  ' - Какие классы загружать
End Type
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (lpInitCtrls As SINITCOMMONCONTROLSEX) As Boolean
Public Const ICC_USEREX_CLASSES = &H200 ' - Класс, который нам нужен.
Option Explicit



Private Function InitCommonControlsVB() As Boolean
   On Error Resume Next
   Dim iccex As SINITCOMMONCONTROLSEX
   ' Ensure CC available:
   With iccex
       .dwSize = LenB(iccex)
       .dwICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)
End Function

Public Sub Main()
    InitCommonControlsVB
    Load MainForm
    MainForm.Show
End Sub
