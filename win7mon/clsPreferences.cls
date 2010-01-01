VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPreferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const Application As String = "Application"

Private prefs As New ARINIManager
Private initialized As Boolean

'------------------------------------------------
'Check server.ini file exists, setup ARINIManager
'------------------------------------------------
Private Sub Class_Initialize()
    On Error GoTo prefs_error
    If (Not initialized) Then
        Dim path As String
        path = App.path + "\" + "NetMon.ini"
        prefs.INIFile = path
        initialized = True
        If (Not FileExists(path)) Then
            MsgBox "Preferences file is not found: " + path, vbCritical
        End If
    End If
    Exit Sub
prefs_error:
    If Err.Number <> 0 Then
        Err.Raise _
        Number:=3, _
        Description:="Preference system error: " + Err.Description & ":" & Err.Number, _
        Source:="clsPreferences.initialize"
    End If
End Sub


'------------------------------------------------------------------
'Server port, used to pass commands, currently must not be modified
'------------------------------------------------------------------
Public Property Get serverPort() As Long
    serverPort = Val(prefs.GetValue(Application, "port", "40000"))
End Property

Public Property Let serverPort(ByVal vNewValue As Long)
    prefs.WriteValue Application, "port", Trim(Str(vNewValue))
End Property



'--------------------------
'Java virual machine to use
'--------------------------
Public Property Get java() As String
    java = prefs.GetValue(Application, "java", "")
End Property

Public Property Let java(ByVal vNewValue As String)
    prefs.WriteValue Application, "java", vNewValue
End Property



'--------------
'Java vmoptions
'--------------
Public Property Get vmoptions() As String
    vmoptions = prefs.GetValue(Application, "vmoptions", "")
End Property

Public Property Let vmoptions(ByVal vNewValue As String)
    prefs.WriteValue Application, "vmoptions", vNewValue
End Property


'----------------------------
'Flex SDK to use, must be set
'----------------------------
Public Property Get sdk() As String
    sdk = prefs.GetValue(Application, "sdk", "c:\flex\3.3\")
End Property

Public Property Let sdk(ByVal vNewValue As String)
    prefs.WriteValue Application, "sdk", vNewValue
End Property



'-------------------------------
'Command line for fcsh execution
'-------------------------------
Public Property Get cmd() As String
    Dim result As String
    
    If (Len(java) > 0) And (FileExists(java)) Then
        Dim fso As New FileSystemObject
        Dim JavaBinPath As String
        JavaBinPath = fso.BuildPath(java, "bin")
        
        result = "cmd.exe /C cd /D " + JavaBinPath + " && java.exe "
    Else
        result = "java.exe "
    End If
    
    result = result + vmoptions
    result = result + " -Dapplication.home=" + GetShortName(sdk) + " -jar " + Chr(34) + fso.BuildPath(sdk, "lib") + "\fcsh.jar" + Chr(34)
    cmd = result
End Property

Public Property Let cmd(ByVal vNewValue As String)
    prefs.WriteValue Application, "vmoptions", vNewValue
End Property



'-----------------------------------
'Inreface Id
'-----------------------------------
Public Property Get interfaceId() As Long
    interfaceId = Val(prefs.GetValue(Application, "interfaceId", "13"))
End Property

Public Property Let interfaceId(ByVal vNewValue As Long)
    prefs.WriteValue Application, "interfaceId", Trim(Str(vNewValue))
End Property



'--------------------------------------------
'Time in milliseconds to wait for fcsh output
'--------------------------------------------
Public Property Get consoleWait() As Long
    consoleWait = Val(prefs.GetValue(Application, "console.wait", "500"))
End Property

Public Property Let consoleWait(ByVal vNewValue As Long)
    prefs.WriteValue Application, "console.wait", Trim(Str(vNewValue))
End Property



'------------------------------
'External logfile viewer to use
'------------------------------
Public Property Get logViewer() As String
    logViewer = prefs.GetValue(Application, "logViewer", "notepad.exe")
End Property

Public Property Let logViewer(ByVal vNewValue As String)
    prefs.WriteValue Application, "logViewer", vNewValue
End Property


'-----------------------------------------
'Log levels to use (info,warn,error,debug)
'-----------------------------------------
Public Property Get logLevels() As String
    logLevels = prefs.GetValue(Application, "logLevels", "")
End Property

Public Property Let logLevels(ByVal vNewValue As String)
    prefs.WriteValue Application, "logLevels", vNewValue
End Property
