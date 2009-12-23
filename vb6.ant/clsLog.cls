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

Private Const cDEBUG As String = "debug"
Private Const cINFO As String = "info"
Private Const cERROR As String = "error"
Private Const cWARN As String = "warn"

'-----------
'Preferences
'-----------
Private prefs As New clsPreferences

'---------------
'Main log window
'---------------
Private logWindow As RichTextBox

'------------
'Error log window
'------------
Private errWindow As RichTextBox

'------------------
'Warning log window
'------------------
Private warnWindow As RichTextBox


'------------------------
'Setup log output to file
'------------------------
Private Sub Class_Initialize()
    prefs.initialize Me
    
    Open getLogFilePath() For Append As #1
    Close #1
End Sub


'------------------------------
'Set all log windows for output
'------------------------------
Public Sub setWindow(ByRef textBox As RichTextBox, ByRef errBox As RichTextBox, ByRef warnBox As RichTextBox)
    Set logWindow = textBox
    Set errWindow = errBox
    Set warnWindow = warnBox
End Sub


'------------------
'Logs debug message
'------------------
Public Sub xDebug(msg As String)
    If (isLogLevelEnabled(cDEBUG)) Then
        output msg, "[DEBUG]"
    End If
End Sub


'-----------------
'Logs info message
'-----------------
Public Sub xInfo(msg As String)
    If (isLogLevelEnabled(cINFO)) Then
        output msg, "[INFO]"
    End If
End Sub


'--------------------
'Logs warning message
'--------------------
Public Sub xWarning(msg As String)
    If (isLogLevelEnabled(cWARN)) Then
        output msg, "[WARN]"
    End If
End Sub


'------------------
'Logs error message
'------------------
Public Sub xError(msg As String)
    If (isLogLevelEnabled(cERROR)) Then
        output msg, "[ERROR]"
    End If
End Sub


'----------------
'Logs fcsh output
'----------------
Public Sub xFcsh(msg As String)
    logWindow.Text = logWindow.Text + msg
    logWindow.SelStart = Len(logWindow.Text) - Len(msg)
    logWindow.SelLength = Len(msg)
    logWindow.SelColor = RGB(90, 90, 90)
    
    xDebug msg
End Sub


'---------------
'Logs fcsh error
'---------------
Public Sub xFcshErr(msg As String)
    errWindow.Text = errWindow.Text + msg + vbCrLf
    errWindow.SelStart = Len(errWindow.Text) - Len(msg)
    errWindow.SelLength = Len(msg)
    'errWindow.SelColor = RGB(150, 0, 0)
    
    xError msg
End Sub


'-----------------
'Logs fcsh warning
'-----------------
Public Sub xFcshWarn(msg As String)
    warnWindow.Text = warnWindow.Text + msg + vbCrLf
    warnWindow.SelStart = Len(warnWindow.Text) - Len(msg)
    warnWindow.SelLength = Len(msg)
    'warnWindow.SelColor = RGB(223, 112, 0)
    
    xWarning msg
End Sub


'------------------
'Logs debug message
'------------------
Public Sub output(msg As String, prefix As String)
    Debug.Print prefix + " " & Now & " " + msg
    
    Open getLogFilePath() For Append As #1
    Print #1, prefix + " " & Now & " " + msg
    Close #1
End Sub


'----------------------------
'Check if loglevel is enabled
'----------------------------
Private Function isLogLevelEnabled(level As String) As Boolean
    Dim logString As String
    Dim levels() As String
    Dim i As Long
    
    logString = prefs.logLevels
    
    levels = Split(logString, ",")
    

    If (Len(Trim(logString)) > 0) Then
    
        For i = 0 To UBound(levels)
            If (UCase(Trim(levels(i))) = UCase(Trim(level))) Then
                isLogLevelEnabled = True
                Exit Function
            End If
        Next i
    End If
    
    isLogLevelEnabled = False
End Function


'-------------------
'Builds logfile path
'-------------------
Private Function getLogFilePath() As String
    Dim fso As New FileSystemObject
    getLogFilePath = fso.BuildPath(App.path, "FCSHServer.log")
End Function

