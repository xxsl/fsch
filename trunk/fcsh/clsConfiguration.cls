VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'api
Private Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function PathIsDirectory Lib "shlwapi" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long

'local vars
Private log As clsLog
Private prefs As New ARINIManager
Private initialized As Boolean

Public SERVER_PORT As Long
Public LOG_DEBUG As Boolean
Public APPLICATIONS As Long
Public SHOW_BALOON As Boolean

Const PREFERENCE As String = "preferences"


Public Sub Load()
    setupConfiguration
    SERVER_PORT = Val(prefs.GetValue(PREFERENCE, "port", "44000"))
    
    Dim useDebug As Long
    useDebug = Val(prefs.GetValue(PREFERENCE, "debug", "0"))
    LOG_DEBUG = (useDebug > 0)
    
    Dim showBaloon As Long
    showBaloon = Val(prefs.GetValue(PREFERENCE, "showBaloonTips", "0"))
    SHOW_BALOON = (showBaloon > 0)
    
    APPLICATIONS = Val(prefs.GetValue("applications", "count", "0"))
End Sub

Public Function LoadApplication(index As Long) As clsTarget
    Dim target As New clsTarget
    Dim num As String
    num = Trim(Str(index))
    target.fCommand = prefs.GetValue(num, "command", "mxmlc")
    target.fName = prefs.GetValue(num, "name", "")
    target.fSource = prefs.GetValue(num, "source", "")
    target.fOutput = prefs.GetValue(num, "output", "")
    target.fDebug = prefs.GetValue(num, "debug", "")
    target.fServices = prefs.GetValue(num, "services", "")
    target.fContext = prefs.GetValue(num, "context", "")
    target.fLibraries = prefs.GetValue(num, "libraries", "")
    Set LoadApplication = target
End Function

Private Sub setupConfiguration()
    If (Not initialized) Then
        Dim path As String
        path = app.path + "/" + "fcsh.ini"
        prefs.INIFile = path
        initialized = True
        If (Not log Is Nothing) Then
            If (FileExists(path)) Then
                log.xInfo "Preferences have loaded successfully"
            Else
                log.xError "Preferences are not found: fcsh.ini"
            End If
        End If
    End If
End Sub

Private Function FileExists(ByVal sPath As String) As Boolean
      If (PathFileExists(sPath)) And Not (PathIsDirectory(sPath)) Then FileExists = True
End Function


Public Property Let logger(ByVal vNewValue As clsLog)
    Set log = vNewValue
End Property
