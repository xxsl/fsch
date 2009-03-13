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

Private Sub Class_Initialize()
    Open App.path + "/" + "FCSHServer.log" For Append As #1
    Print #1, vbCrLf
    Print #1, "--------------------------------------------------------------------------"
End Sub

Public Sub xDebug(msg As String)
    output msg, "[DEBUG]"
End Sub

Public Sub xInfo(msg As String)
    output msg, "[INFO]"
End Sub

Public Sub xWarning(msg As String)
    output msg, "[WARN]"
End Sub

Public Sub xError(msg As String)
    output msg, "[ERROR]"
End Sub

Private Sub output(msg As String, prefix As String)
    Debug.Print prefix + " " & Now & " " + msg
    Print #1, prefix + " " & Now & " " + msg
End Sub


Private Sub Class_Terminate()
    Close #1
End Sub
