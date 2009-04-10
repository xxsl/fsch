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

'Keys
Private Const ALWAYSONTOP_KEY As String = "AlwaysOnTop"

Public prefs As New ARINIManager


Private Sub Class_Initialize()
    prefs.INIFile = App.Path + "/AntRunner.ini"
    prefs.Refresh
End Sub

Public Sub DeleteBuild(section As String)
    prefs.Sections.Remove section
End Sub

Public Property Get Builds() As Collection
    Dim Result As New Collection
    Dim build As AntBuild
    Dim i As Long
    For i = 1 To prefs.Sections.Count
        If (Mid(prefs.Sections.Item(i).Name, 1, Len("Build_")) = "Build_") Then
            Set build = New AntBuild
            build.load prefs, prefs.Sections.Item(i).Name
            Result.Add build
        End If
    Next i
    Set Builds = Result
End Property

Public Property Get AlwaysOnTop() As Boolean
    Dim value As Long
    value = CLng(prefs.getValue(APPLICATION_SECTION, ALWAYSONTOP_KEY, "0"))
    AlwaysOnTop = (value = 1)
End Property

Public Property Let AlwaysOnTop(ByVal vNewValue As Boolean)
    Dim value As String
    If (vNewValue) Then
        value = "1"
    Else
        value = "0"
    End If
    prefs.WriteValue APPLICATION_SECTION, ALWAYSONTOP_KEY, value
End Property
