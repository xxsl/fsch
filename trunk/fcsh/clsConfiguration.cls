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



Public Sub Load()
    setupConfiguration
    SERVER_PORT = Val(prefs.GetValue("preferences", "port", "44000"))
End Sub



Private Sub setupConfiguration()
    If (Not initialized) Then
        Dim path As String
        path = App.path + "/" + "fcsh.ini"
        prefs.INIFile = path
        initialized = True
        If (FileExists(path)) Then
            log.xInfo "Preferences have loaded successfully"
        Else
            log.xError "Preferences are not found: fcsh.ini"
        End If
    End If
End Sub

Private Function FileExists(ByVal sPath As String) As Boolean
      If (PathFileExists(sPath)) And Not (PathIsDirectory(sPath)) Then FileExists = True
End Function


Public Property Let Logger(ByVal vNewValue As clsLog)
    Set log = vNewValue
End Property
