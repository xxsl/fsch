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

'local vars
Private log As clsLog
Private prefs As New ARINIManager
Private initialized As Boolean

Private fSERVER_PORT As Long
Private fLOG_DEBUG As Boolean
Private fAPPLICATIONS As Long
Private fSHOW_BALOON As Boolean

Const PREFERENCE As String = "preferences"


Public Property Let logger(ByVal vNewValue As clsLog)
    Set log = vNewValue
End Property

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

Public Sub Load()
    setupConfiguration
    fSERVER_PORT = Val(prefs.GetValue(PREFERENCE, "port", "44000"))
    
    Dim useDebug As Long
    useDebug = Val(prefs.GetValue(PREFERENCE, "debug", "0"))
    fLOG_DEBUG = (useDebug > 0)
    
    Dim showBaloon As Long
    showBaloon = Val(prefs.GetValue(PREFERENCE, "showBaloonTips", "0"))
    fSHOW_BALOON = (showBaloon > 0)
    
    fAPPLICATIONS = Val(prefs.GetValue("applications", "count", "0"))
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

Public Sub saveApplication(index As Long, ByRef target As clsTarget)
    Dim num As String
    num = Trim(Str(index))
    
    prefs.WriteValue num, "command", target.fCommand
    prefs.WriteValue num, "name", target.fName
    prefs.WriteValue num, "source", target.fSource
    prefs.WriteValue num, "output", target.fOutput
    prefs.WriteValue num, "debug", target.fDebug
    prefs.WriteValue num, "services", target.fServices
    prefs.WriteValue num, "context", target.fContext
    prefs.WriteValue num, "libraries", target.fLibraries
End Sub


Public Sub Clear()
Open app.path + "/" + "fcsh.ini" For Output As #1
Close #1
End Sub

'save boolean -------------------------
Private Sub saveBool(node As String, key As String, value As Boolean)
    If (value) Then
        prefs.WriteValue node, key, 1
    Else
        prefs.WriteValue node, key, 0
    End If
End Sub

'server port --------------------------
Public Property Get SERVER_PORT() As Long
    SERVER_PORT = fSERVER_PORT
End Property

Public Property Let SERVER_PORT(ByVal vNewValue As Long)
    fSERVER_PORT = vNewValue
    prefs.WriteValue PREFERENCE, "port", Str(fSERVER_PORT)
End Property

'log debug ----------------------------
Public Property Get LOG_DEBUG() As Boolean
    LOG_DEBUG = fLOG_DEBUG
End Property

Public Property Let LOG_DEBUG(ByVal vNewValue As Boolean)
    fLOG_DEBUG = vNewValue
    saveBool PREFERENCE, "debug", fLOG_DEBUG
End Property

'show baloons
Public Property Get SHOW_BALOON() As Boolean
    SHOW_BALOON = fSHOW_BALOON
End Property

Public Property Let SHOW_BALOON(ByVal vNewValue As Boolean)
    fSHOW_BALOON = vNewValue
    saveBool PREFERENCE, "showBaloonTips", fSHOW_BALOON
End Property

'apps count
Public Property Get APPLICATIONS() As Long
    APPLICATIONS = fAPPLICATIONS
End Property

Public Property Let APPLICATIONS(ByVal vNewValue As Long)
    fAPPLICATIONS = vNewValue
    prefs.WriteValue "applications", "count", Str(fAPPLICATIONS)
End Property

