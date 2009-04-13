VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Sections
Private Const APPLICATION_SECTION As String = "Application"
Private Const HOTKEY_SECTION As String = "Hotkey"

'Keys
Private Const ALWAYSONTOP_KEY As String = "ShowBaloon"
Private Const MINIMIZE_KEY As String = "MinimizeOnClose"
Private Const HOTKEYENABLED_KEY As String = "HotkeyEnabled"
Private Const ANTPATH_KEY As String = "AntPath"


Public prefs As New ARINIManager


Private Sub Class_Initialize()
    prefs.INIFile = App.Path + "/AntRunner.ini"
    prefs.Refresh
End Sub

Public Sub DeleteBuild(section As String)
    prefs.Sections.Remove section
End Sub

Public Property Get Builds() As Collection
    Dim result As New Collection
    Dim build As AntBuild
    Dim i As Long
    For i = 1 To prefs.Sections.Count
        If (Mid(prefs.Sections.Item(i).Name, 1, Len("Build_")) = "Build_") Then
            Set build = New AntBuild
            build.load prefs, prefs.Sections.Item(i).Name
            If (build.isParsed) Then
                result.Add build
            End If
        End If
    Next i
    Set Builds = result
End Property

Public Property Get ShowBaloon() As Boolean
    Dim value As Long
    value = CLng(prefs.getValue(APPLICATION_SECTION, ALWAYSONTOP_KEY, "0"))
    ShowBaloon = (value = 1)
End Property

Public Property Let ShowBaloon(ByVal vNewValue As Boolean)
    Dim value As String
    If (vNewValue) Then
        value = "1"
    Else
        value = "0"
    End If
    prefs.WriteValue APPLICATION_SECTION, ALWAYSONTOP_KEY, value
End Property

Public Property Get MinimizeOnClose() As Boolean
    Dim value As Long
    value = CLng(prefs.getValue(APPLICATION_SECTION, MINIMIZE_KEY, "0"))
    MinimizeOnClose = (value = 1)
End Property

Public Property Let MinimizeOnClose(ByVal vNewValue As Boolean)
    Dim value As String
    If (vNewValue) Then
        value = "1"
    Else
        value = "0"
    End If
    prefs.WriteValue APPLICATION_SECTION, MINIMIZE_KEY, value
End Property


Public Property Get HotKeyEnabled() As Boolean
    Dim value As Long
    value = CLng(prefs.getValue(APPLICATION_SECTION, HOTKEYENABLED_KEY, "0"))
    HotKeyEnabled = (value = 1)
End Property

Public Property Let HotKeyEnabled(ByVal vNewValue As Boolean)
    Dim value As String
    If (vNewValue) Then
        value = "1"
    Else
        value = "0"
    End If
    prefs.WriteValue APPLICATION_SECTION, HOTKEYENABLED_KEY, value
End Property

Public Property Let Key(ByVal vNewValue As Long)
    prefs.WriteValue HOTKEY_SECTION, "key", Str(vNewValue)
End Property

'use hot keys
Public Property Get Key() As Long
    Key = Val(prefs.getValue(HOTKEY_SECTION, "key", "0"))
End Property

Public Property Let MODIFYER(ByVal vNewValue As Long)
    prefs.WriteValue HOTKEY_SECTION, "key_modifyer", Str(vNewValue)
End Property

'use hot keys
Public Property Get MODIFYER() As Long
    MODIFYER = Val(prefs.getValue(HOTKEY_SECTION, "key_modifyer", "0"))
End Property

Public Property Let AntPath(ByVal vNewValue As String)
    prefs.WriteValue APPLICATION_SECTION, ANTPATH_KEY, vNewValue
End Property

'use hot keys
Public Property Get AntPath() As String
    AntPath = prefs.getValue(APPLICATION_SECTION, ANTPATH_KEY, "ant.bat")
End Property
