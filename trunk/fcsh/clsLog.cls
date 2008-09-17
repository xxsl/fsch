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

Public Sub clsLog(ByRef log As RichTextBox)
    Set LOGGER = log
End Sub

Public Sub xDebug(ByVal msg As String)
    output msg, vbBlue
End Sub

Public Sub xInfo(ByVal msg As String)
    output msg, vbGrayText
End Sub

Public Sub xError(ByVal msg As String)
    output msg, vbRed
End Sub

Public Sub Clear()
    LOGGER.text = ""
End Sub

'logs to console
Public Sub output(text As String, color As Long)
    LOGGER.TextRTF
    LOGGER.SelColor = vbBlack
    LOGGER.text = LOGGER.text + text
    LOGGER.SelStart = Len(LOGGER.text) - Len(text)
    LOGGER.SelLength = Len(text)
    LOGGER.SelColor = color
    LOGGER.SelStart = Len(LOGGER.text)
End Sub

 
