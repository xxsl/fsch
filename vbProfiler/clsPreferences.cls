VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPreferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const Profiler    As String = "Profiler"

Const Application   As String = "Application"

Private prefs       As New ARINIManager

Private Log         As clsLog

Private initialized As Boolean

Public Sub initialize(appLog As clsLog)

    On Error GoTo prefs_error

    If (Not initialized) Then
        Set Log = appLog

        Dim path As String

        path = App.path + "\" + "profiler.ini"
        prefs.INIFile = path
        initialized = True
    
        If (Not FileExists(path)) Then
            MsgBox "Preferences file is not found: " + path, vbCritical
        End If
    End If

    Exit Sub

prefs_error:

    If Err.Number <> 0 Then
        Err.Raise Number:=3, Description:="Preference system error: " + Err.Description & ":" & Err.Number, Source:="clsPreferences.initialize"
    End If

End Sub

'-----------------------
'Profiler
'-----------------------

Public Property Get port() As Long
    port = Val(prefs.getValue(Profiler, "profiler.port", "9999"))
End Property

Public Property Let port(ByVal vNewValue As Long)
    prefs.WriteValue Profiler, "profiler.port", Trim(str(vNewValue))
End Property


Public Property Get host() As String
    host = prefs.getValue(Profiler, "profiler.host", "localhost")
End Property

Public Property Let host(ByVal vNewValue As String)
    prefs.WriteValue Profiler, "profiler.host", vNewValue
End Property


Public Property Get profilerSWF() As String
    profilerSWF = prefs.getValue(Profiler, "profiler.profilerSWF", "c:\Users\Dookie\ProfilerAgent.swf")
End Property

Public Property Let profilerSWF(ByVal vNewValue As String)
    prefs.WriteValue Profiler, "profiler.profilerSWF", vNewValue
End Property


Public Property Get urlPattern() As String
    urlPattern = prefs.getValue(Profiler, "profiler.url.pattern", "${file}?host=${host}&port=${port}")
End Property

Public Property Let urlPattern(ByVal vNewValue As String)
    prefs.WriteValue Profiler, "profiler.url.pattern", vNewValue
End Property


'-----------------------
'Application
'-----------------------
Public Property Get alwaysOnTop() As Boolean
    alwaysOnTop = Val(prefs.getValue(Application, "AlwaysOnTop", "0")) = 1
End Property

Public Property Let alwaysOnTop(ByVal vNewValue As Boolean)

    Dim result As Long

    If (vNewValue) Then
        result = 1
    End If

    prefs.WriteValue Application, "AlwaysOnTop", Trim(str(result))
End Property

Public Property Get logViewer() As String
    logViewer = prefs.getValue(Application, "logViewer", "notepad.exe")
End Property

Public Property Let logViewer(ByVal vNewValue As String)
    prefs.WriteValue Application, "logViewer", vNewValue
End Property

Public Property Get logLevels() As String
    logLevels = prefs.getValue(Application, "logLevels", "info,warn,error")
End Property

Public Property Let logLevels(ByVal vNewValue As String)
    prefs.WriteValue Application, "logLevels", vNewValue
End Property
