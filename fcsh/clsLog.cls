VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***********************************************************************************
'* nimrod97@gmail.com                                                              *
'* Project homepage http://code.google.com/p/fsch/                                 *
'* Adobe Flex Compiler Shell wrapper                                               *
'* 2008                                                                            *
'***********************************************************************************

Option Explicit

Private logger As RichTextBox
Private color As Long
Private config As clsConfiguration

Public Sub clsLog(ByRef log As RichTextBox, cfg As clsConfiguration)
    Set logger = log
    Set config = cfg
    color = RGB(0, 126, 255)
End Sub

Public Sub xDebug(ByVal Msg As String)
    If (config.LOG_DEBUG) Then
        output "[DEBUG] " + Msg + vbCrLf
    Else
        Write #1, Msg
    End If
End Sub

Public Sub xInfo(ByVal Msg As String)
    output "[INFO]  " + Msg + vbCrLf
End Sub

Public Sub xError(ByVal Msg As String)
    output "[ERROR] " + Msg + vbCrLf
End Sub

Public Sub xFcsh(ByVal Msg As String)
    output "[FCSH] " + Msg + vbCrLf
End Sub

Public Sub xExec(ByVal Msg As String)
    output "[EXEC] " + Msg + vbCrLf
End Sub

Public Sub Text(Msg As String)
    output Msg
End Sub


Public Sub Clear()
    logger.Text = ""
End Sub

'logs to console
Public Sub output(Text As String)
    logger.Text = logger.Text + Text
    logger.SelStart = Len(logger.Text) - Len(Text)
    logger.SelLength = Len(Text)
    logger.SelColor = color
    logger.SelStart = Len(logger.Text)
    Print #1, Text
End Sub

 
Private Sub Class_Initialize()
Open app.path + "/" + "fcsh.log" For Append As #1
Print #1, vbCrLf
Print #1, Now
Print #1, "--------------------------------------------------------------------------"
End Sub

Private Sub Class_Terminate()
Close 1
End Sub


