Attribute VB_Name = "modMain"
Option Explicit

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean

'--------
'controls
'--------
Private Const ICC_USEREX_CLASSES = &H200

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type


'--------------------
'Try use WinXP styles
'--------------------
Private Function InitCommonControlsVB() As Boolean
   On Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   ' Ensure CC available:
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)
   On Error GoTo 0
End Function

'-----------------------
'Application entry point
'-----------------------
Public Sub Main()
    InitCommonControlsVB
    If (Not App.PrevInstance) Then
        Load MainForm
    Else
        MsgBox "FCSHServer is already running!", vbCritical, "FCSHServer"
    End If
End Sub




