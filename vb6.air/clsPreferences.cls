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

