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
'***********************************************************************************
'* nimrod97@gmail.com                                                              *
'* Project homepage http://code.google.com/p/fsch/                                 *
'* Adobe Flex Compiler Shell wrapper                                               *
'* 2008                                                                            *
'***********************************************************************************

Option Explicit

'local vars
Private prefs As New ARINIManager
Private initialized As Boolean

Private fSERVER_PORT As Long
Private fLOG_DEBUG As Boolean
Private fAPPLICATIONS As Long
Private fSHOW_BALOON As Boolean
Private fFCSH_PATH As String
Private fALPHA As Byte
Private fBACKCOLOR As Long
Private fFORECOLOR As Long

Public RECOMPILE As New clsHotKey

Const PREFERENCE As String = "preferences"


Private Sub setupConfiguration()
    If (Not initialized) Then
        Dim path As String
        path = app.path + "/" + "fcsh.ini"
        prefs.INIFile = path
        initialized = True

        If (Not FileExists(path)) Then
           MsgBox "Preferences are not found: fcsh.ini", vbCritical
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
    
    fFCSH_PATH = prefs.GetValue(PREFERENCE, "fcsh_exe_path", "fcsh.exe")
    
    fALPHA = CByte(Val(prefs.GetValue(PREFERENCE, "alpha", "128")))
    
    fBACKCOLOR = Val(prefs.GetValue(PREFERENCE, "backcolor", "-2147483633"))
    
    fFORECOLOR = Val(prefs.GetValue(PREFERENCE, "forecolor", "16743936"))
    
    fAPPLICATIONS = Val(prefs.GetValue("applications", "count", "0"))
    
    RECOMPILE.init prefs, "recompile_hotkey"
End Sub

Public Function LoadApplication(index As Long) As clsTarget
    Dim target As New clsTarget
    Dim Num As String
    Num = Trim(Str(index))
    target.fCommand = prefs.GetValue(Num, "command", "mxmlc")
    target.fName = prefs.GetValue(Num, "name", "")
    target.fSource = prefs.GetValue(Num, "source", "")
    target.fOutput = prefs.GetValue(Num, "output", "")
    target.fDebug = prefs.GetValue(Num, "debug", "")
    target.fServices = prefs.GetValue(Num, "services", "")
    target.fContext = prefs.GetValue(Num, "context", "")
    target.fLibraries = prefs.GetValue(Num, "libraries", "")
    'target.fOther = prefs.GetValue(num, "other", "")
    Set LoadApplication = target
End Function

Public Sub saveApplication(index As Long, ByRef target As clsTarget)
    Dim Num As String
    Num = Trim(Str(index))
    
    prefs.WriteValue Num, "command", target.fCommand
    prefs.WriteValue Num, "name", target.fName
    prefs.WriteValue Num, "source", target.fSource
    prefs.WriteValue Num, "output", target.fOutput
    prefs.WriteValue Num, "debug", target.fDebug
    prefs.WriteValue Num, "services", target.fServices
    prefs.WriteValue Num, "context", target.fContext
    prefs.WriteValue Num, "libraries", target.fLibraries
    'prefs.WriteValue num, "other", target.fOther
End Sub


Public Sub clear()
Open app.path + "/" + "fcsh.ini" For Output As #1
Close #1
End Sub

'save boolean -------------------------
Private Sub saveBool(node As String, KEY As String, Value As Boolean)
    If (Value) Then
        prefs.WriteValue node, KEY, 1
    Else
        prefs.WriteValue node, KEY, 0
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


'fcsh path
Public Property Get FCSH_PATH() As String
    FCSH_PATH = fFCSH_PATH
End Property

Public Property Let FCSH_PATH(ByVal vNewValue As String)
    fFCSH_PATH = vNewValue
    prefs.WriteValue PREFERENCE, "fcsh_exe_path", fFCSH_PATH
End Property

'transparency
Public Property Get Alpha() As Byte
    Alpha = fALPHA
End Property

Public Property Let Alpha(ByVal vNewValue As Byte)
    fALPHA = vNewValue
    prefs.WriteValue PREFERENCE, "alpha", Trim(Str(fALPHA))
End Property

'log backcolor
Public Property Get BackColor() As Long
    BackColor = fBACKCOLOR
End Property

Public Property Let BackColor(ByVal vNewValue As Long)
    fBACKCOLOR = vNewValue
    prefs.WriteValue PREFERENCE, "backcolor", Str(fBACKCOLOR)
End Property

'log forecolor
Public Property Get ForeColor() As Long
    ForeColor = fFORECOLOR
End Property

Public Property Let ForeColor(ByVal vNewValue As Long)
    fFORECOLOR = vNewValue
    prefs.WriteValue PREFERENCE, "forecolor", Str(fFORECOLOR)
End Property


