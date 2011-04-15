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
Private logWindow As ListBox



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
Public Sub setWindow(ByRef list As ListBox)
    Set logWindow = list
End Sub


'------------------
'Logs debug message
'------------------
Public Sub xDebug(msg As String)
    If (isLogLevelEnabled(cDEBUG)) Then
        output msg, "DEBUG"
    End If
End Sub


'-----------------
'Logs info message
'-----------------
Public Sub xInfo(msg As String)
    If (isLogLevelEnabled(cINFO)) Then
        output msg, "INFO"
    End If
End Sub


'--------------------
'Logs warning message
'--------------------
Public Sub xWarning(msg As String)
    If (isLogLevelEnabled(cWARN)) Then
        output msg, "WARN"
    End If
End Sub


'------------------
'Logs error message
'------------------
Public Sub xError(msg As String)
    If (isLogLevelEnabled(cERROR)) Then
        output msg, "ERROR"
    End If
End Sub


'------------------
'Logs debug message
'------------------
Public Sub output(msg As String, prefix As String)

    Dim result As String
    
    result = prefix + " [" & Now & "] " + msg
    
    Debug.Print result
    
    Open getLogFilePath() For Append As #1
    Print #1, result
    Close #1
    
    logWindow.AddItem result
    logWindow.TopIndex = logWindow.NewIndex
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
    getLogFilePath = fso.BuildPath(App.path, "profiler.log")
End Function

