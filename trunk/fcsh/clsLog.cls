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

Private Logger As RichTextBox
Private color As Long

Public Sub clsLog(ByRef log As RichTextBox)
    Set Logger = log
    color = vbBlue
End Sub

Public Sub xDebug(ByVal Msg As String)
    output "[DEBUG] " + Msg + vbCrLf
End Sub

Public Sub xInfo(ByVal Msg As String)
    output "[INFO]  " + Msg + vbCrLf
End Sub

Public Sub xError(ByVal Msg As String)
    output "[ERROR] " + Msg + vbCrLf
End Sub

Public Sub text(Msg As String)
    output Msg
End Sub


Public Sub Clear()
    Logger.text = ""
End Sub

'logs to console
Public Sub output(text As String)
    Logger.text = Logger.text + text
    Logger.SelStart = Len(Logger.text) - Len(text)
    Logger.SelLength = Len(text)
    Logger.SelColor = color
    Logger.SelStart = Len(Logger.text)
    Print #1, text
End Sub

 
Private Sub Class_Initialize()
Open App.path + "/" + "fcsh.log" For Append As #1
Print #1, vbCrLf
Print #1, Now
Print #1, "--------------------------------------------------------------------------"
End Sub

Private Sub Class_Terminate()
Close 1
End Sub
