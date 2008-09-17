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
Option Explicit

Private LOGGER As RichTextBox
Private color As Long

Public Sub clsLog(ByRef log As RichTextBox)
    Set LOGGER = log
    color = vbBlue
End Sub

Public Sub xDebug(ByVal msg As String)
    output "[DEBUG] " + msg + vbCrLf
End Sub

Public Sub xInfo(ByVal msg As String)
    output "[INFO]  " + msg + vbCrLf
End Sub

Public Sub xError(ByVal msg As String)
    output "[ERROR] " + msg + vbCrLf
End Sub

Public Sub Text(msg As String)
    output msg
End Sub


Public Sub Clear()
    LOGGER.Text = ""
End Sub

'logs to console
Public Sub output(Text As String)
    LOGGER.Text = LOGGER.Text + Text
    LOGGER.SelStart = Len(LOGGER.Text) - Len(Text)
    LOGGER.SelLength = Len(Text)
    LOGGER.SelColor = color
    LOGGER.SelStart = Len(LOGGER.Text)
End Sub

 
