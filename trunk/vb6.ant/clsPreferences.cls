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
'***********************************************************************************
'* nimrod97@gmail.com                                                              *
'* Project homepage http://code.google.com/p/fsch/                                 *
'* Adobe Flex Compiler Shell wrapper                                               *
'* 2008                                                                            *
'***********************************************************************************

Option Explicit

Const FCSHServer As String = "FCSHServer"

Private prefs As New ARINIManager
Private log As clsLog
Private initialized As Boolean


'------------------------------------------------
'Check server.ini file exists, setup ARINIManager
'------------------------------------------------
Public Sub initialize(appLog As clsLog)
    If (Not initialized) Then
        Set log = appLog
        log.xInfo "Initialize preferences"
        Dim path As String
        path = App.path + "\" + "server.ini"
        prefs.INIFile = path
        initialized = True

        If (Not FileExists(path)) Then
           log.xError "Preferences file is not found: " + path
           MsgBox "Preferences file is not found: " + path, vbCritical
        Else
           log.xInfo "Preferences initialized"
        End If
    End If
End Sub


'------------------------------------------------------------------
'Server port, used to pass commands, currently must not be modified
'------------------------------------------------------------------
Public Property Get serverPort() As Long
    serverPort = Val(prefs.GetValue(FCSHServer, "port", "40000"))
End Property
Public Property Let serverPort(ByVal vNewValue As Long)
    prefs.WriteValue FCSHServer, "port", Trim(str(vNewValue))
End Property



'--------------------------
'Java virual machine to use
'--------------------------
Public Property Get java() As String
    java = prefs.GetValue(FCSHServer, "java", "")
End Property
Public Property Let java(ByVal vNewValue As String)
    prefs.WriteValue FCSHServer, "java", vNewValue
End Property




'-------------------------------
'Command line for fcsh execution
'-------------------------------
Public Property Get cmd() As String
    Dim result As String, sdk As String
    sdk = prefs.GetValue(FCSHServer, "sdk", "C:\flex")
    result = prefs.GetValue(FCSHServer, "vmoptions", "")
    result = "java.exe " + result
    result = result + " -Dapplication.home=" + sdk + " -jar " + sdk + "\lib\fcsh.jar"
    log.xInfo result
    cmd = result
End Property
Public Property Let cmd(ByVal vNewValue As String)
    prefs.WriteValue FCSHServer, "vmoptions", vNewValue
End Property



'-----------------------------------
'Buffer for console output, in bytes
'-----------------------------------
Public Property Get consoleBuffer() As Long
    consoleBuffer = Val(prefs.GetValue(FCSHServer, "console.buffer", "8192"))
End Property
Public Property Let consoleBuffer(ByVal vNewValue As Long)
    prefs.WriteValue FCSHServer, "console.buffer", Trim(str(vNewValue))
End Property



'--------------------------------------------
'Time in milliseconds to wait for fcsh output
'--------------------------------------------
Public Property Get consoleWait() As Long
    consoleWait = Val(prefs.GetValue(FCSHServer, "console.wait", "500"))
End Property

Public Property Let consoleWait(ByVal vNewValue As Long)
    prefs.WriteValue FCSHServer, "console.wait", Trim(str(vNewValue))
End Property
