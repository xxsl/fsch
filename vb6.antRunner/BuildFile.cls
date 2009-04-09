VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AntBuild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public DisplayName As String

Private id As String
Private xml As New DOMDocument
Private xmlFile As String

Private Sub Class_Initialize()
    xml.async = False
End Sub

Private Sub parseBuildXML()
    xml.load buildfile
End Sub

Public Property Get Name() As Variant
    Dim projectName As IXMLDOMAttribute
    Set projectName = xml.selectSingleNode("//project/@name")
    Name = projectName.Text
End Property

Public Property Get Targets() As Collection
    Dim result As New Collection
    Dim AntTask As AntTarget
    
    Dim TargetsList As IXMLDOMNodeList
    Dim Target As IXMLDOMElement
    
    Set TargetsList = xml.selectNodes("//project/target")

    If (Not xml.parsed) Then
        MsgBox xml.parseError.reason & " at " & xml.parseError.Line & ", position " & xml.parseError.linepos, vbCritical, "Parser error"
    End If

    For Each Target In TargetsList
        Set AntTask = New AntTarget
        AntTask.fromXMLNode Target
        Set AntTask.Build = Me
        result.Add AntTask, AntTask.Name
    Next
    
    Set Targets = result
End Property

Public Function getTargetByName(Name As String) As AntTarget
    Dim AntTask As AntTarget
    
    Dim TargetsList As IXMLDOMNodeList
    Dim Target As IXMLDOMElement
    
    Set TargetsList = xml.selectNodes("//project/target")

    For Each Target In TargetsList
        Set AntTask = New AntTarget
        AntTask.fromXMLNode Target
        Set AntTask.Build = Me
        If (AntTask.Name = Name) Then
            Exit For
        End If
    Next
    
    Set getTargetByName = AntTask
End Function

Public Property Get buildfile() As String
    buildfile = xmlFile
End Property

Public Property Let buildfile(ByVal file As String)
    xmlFile = file
    parseBuildXML
End Property

Public Sub load(ByRef prefs As ARINIManager, section As String)
    buildfile = prefs.getValue(section, "file")
    DisplayName = prefs.getValue(section, "name")
    id = section
End Sub

Public Sub save(ByRef prefs As ARINIManager)
    prefs.WriteValue id, "file", buildfile
    prefs.WriteValue id, "name", DisplayName
End Sub






