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
    'target.fOther = prefs.GetValue(num, "other", "")
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
    'prefs.WriteValue num, "other", target.fOther
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


'fcsh path
Public Property Get FCSH_PATH() As String
    FCSH_PATH = fFCSH_PATH
End Property

Public Property Let FCSH_PATH(ByVal vNewValue As String)
    fFCSH_PATH = vNewValue
    prefs.WriteValue PREFERENCE, "fcsh_exe_path", fFCSH_PATH
End Property

'transparency
Public Property Get ALPHA() As Byte
    ALPHA = fALPHA
End Property

Public Property Let ALPHA(ByVal vNewValue As Byte)
    fALPHA = vNewValue
    prefs.WriteValue PREFERENCE, "alpha", Trim(Str(fALPHA))
End Property

'log backcolor
Public Property Get BACKCOLOR() As Long
    BACKCOLOR = fBACKCOLOR
End Property

Public Property Let BACKCOLOR(ByVal vNewValue As Long)
    fBACKCOLOR = vNewValue
    prefs.WriteValue PREFERENCE, "backcolor", Str(fBACKCOLOR)
End Property

'log forecolor
Public Property Get FORECOLOR() As Long
    FORECOLOR = fFORECOLOR
End Property

Public Property Let FORECOLOR(ByVal vNewValue As Long)
    fFORECOLOR = vNewValue
    prefs.WriteValue PREFERENCE, "forecolor", Str(fFORECOLOR)
End Property
