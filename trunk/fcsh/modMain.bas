Attribute VB_Name = "modMain"
'***********************************************************************************
'* nimrod97@gmail.com                                                              *
'* Project homepage http://code.google.com/p/fsch/                                 *
'* Adobe Flex Compiler Shell wrapper                                               *
'* 2008                                                                            *
'***********************************************************************************


Public Const START_FCSH As String = "START_FCSH"
Public Const STOP_FCSH As String = "STOP_FCSH"
Public Const LOG_CLEAR As String = "LOG_CLEAR"
Public Const BUILD_TASK As String = "BUILD"
Public Const INCREMENTAL_ON As String = "INCREMENTAL_ON"
Public Const INCREMENTAL_OFF As String = "INCREMENTAL_OFF"
Public Const TARGET_INFO As String = "TARGET_INFO"
Public Const OPTIONS As String = "OPTIONS"
Public Const ON_TOP As String = "ON_TOP"
Public Const Alpha As String = "ALPHA"
Public Const ABOUT As String = "ABOUT"
Public Const BUILD_PROGRESS As String = "BUILD_PROGRESS"
Public Const APP_APPEARANCE As String = "APP_APPEARANCE"
Public Const KEYBOARD As String = "KEYBOARD"

Public Const ERROR_PNG As String = "ERROR_PNG"
Public Const EXEC_PNG As String = "EXEC_PNG"
Public Const STOPPED_PNG As String = "STOPPED_PNG"
Public Const IDLE_PNG As String = "IDLE_PNG"


Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Private Const ICC_USEREX_CLASSES = &H200
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean


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

Public Sub Main()
    InitCommonControlsVB
    MainForm.Show
End Sub


Public Function getPath() As String
    If (Len(app.path) = 3) Then
        getPath = app.path
    Else
        getPath = app.path + "\"
    End If
End Function
