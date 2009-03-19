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

Private Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function PathIsDirectory Lib "shlwapi" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long

Private Const FCSHServer As String = "FCSHServer"

Public prefs As New ARINIManager
Private log As clsLog
Private initialized As Boolean



Public Sub initialize(appLog As clsLog)
    If (Not initialized) Then
        Set log = appLog
        log.xInfo "initialize preferences"
        Dim path As String
        path = App.path + "\" + "server.ini"
        prefs.INIFile = path
        initialized = True

        If (Not FileExists(path)) Then
           log.xError "Preferences file is not found: " + path
           MsgBox "Preferences file is not found: " + path, vbCritical
        End If
    End If
End Sub

Private Function FileExists(ByVal sPath As String) As Boolean
      If (PathFileExists(sPath)) And Not (PathIsDirectory(sPath)) Then FileExists = True
End Function

'server port --------------------------
Public Property Get SERVER_PORT() As Long
    SERVER_PORT = Val(prefs.GetValue(FCSHServer, "port", "40000"))
End Property

Public Property Let SERVER_PORT(ByVal vNewValue As Long)
    prefs.WriteValue FCSHServer, "port", Trim(str(vNewValue))
End Property

'server host --------------------------
Public Property Get SERVER_HOST() As String
    SERVER_HOST = prefs.GetValue(FCSHServer, "host", "localhost")
End Property

Public Property Let SERVER_HOST(ByVal vNewValue As String)
    prefs.WriteValue FCSHServer, "host", vNewValue
End Property

'exec command
Public Property Get CMD() As String
    Dim result As String, sdk As String
    sdk = prefs.GetValue(FCSHServer, "sdk", "C:\flex")
    result = prefs.GetValue(FCSHServer, "vmoptions", "fcsh.exe")
    result = "java.exe " + result
    result = result + " -Dapplication.home=" + sdk + " -jar " + sdk + "\lib\fcsh.jar"
    log.xInfo result
    CMD = result
End Property

Public Property Let CMD(ByVal vNewValue As String)
    prefs.WriteValue FCSHServer, "vmoptions", vNewValue
End Property

Public Property Get CONSOLE_BUFFER() As Long
    CONSOLE_BUFFER = Val(prefs.GetValue(FCSHServer, "console.buffer", "8192"))
End Property

Public Property Let CONSOLE_BUFFER(ByVal vNewValue As Long)
    prefs.WriteValue FCSHServer, "console.buffer", Trim(str(vNewValue))
End Property

Public Property Get CONSOLE_WAIT() As Long
    CONSOLE_WAIT = Val(prefs.GetValue(FCSHServer, "console.wait", "8192"))
End Property

Public Property Let CONSOLE_WAIT(ByVal vNewValue As Long)
    prefs.WriteValue FCSHServer, "console.wait", Trim(str(vNewValue))
End Property
