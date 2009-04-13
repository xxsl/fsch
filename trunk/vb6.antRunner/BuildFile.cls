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

Public Id As String

Private xml As New DOMDocument
Private xmlFile As String

Private Sub Class_Initialize()
    xml.async = False
End Sub

Private Sub parseBuildXML()
    xml.load BuildFile
End Sub

Public Property Get Name() As Variant
    Dim projectName As IXMLDOMAttribute
    Set projectName = xml.selectSingleNode("//project/@name")
    If (Not projectName Is Nothing) Then
        Name = projectName.Text
    Else
        Name = ""
    End If
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
        Set AntTask.build = Me
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
        Set AntTask.build = Me
        If (AntTask.Name = Name) Then
            Exit For
        End If
    Next
    
    Set getTargetByName = AntTask
End Function

Public Property Get isParsed() As Boolean
    isParsed = xml.parsed
End Property

Public Property Get BuildFile() As String
    BuildFile = xmlFile
End Property

Public Property Let BuildFile(ByVal file As String)
    xmlFile = file
    parseBuildXML
End Property

Public Sub GenerateId()
    Id = "Build_" & Rnd
End Sub

Public Sub load(ByRef prefs As ARINIManager, section As String)
    BuildFile = prefs.getValue(section, "file")
    DisplayName = prefs.getValue(section, "name")
    Id = section
End Sub

Public Sub save(ByRef prefs As ARINIManager)
    prefs.WriteValue Id, "file", BuildFile
    prefs.WriteValue Id, "name", DisplayName
End Sub






