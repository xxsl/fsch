VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBuildXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public content As New DOMDocument

    
    

Private Sub Class_Initialize()
    content.async = False
End Sub



Public Sub LoadXML(url As String)
    content.Load url
End Sub

Public Function getTargets() As Variant
    Dim result() As String
    Dim i As Long
    Dim targets As IXMLDOMNodeList
    Dim target As IXMLDOMElement
    
    Set targets = content.selectNodes("//project/target")
    
    If (targets.length = 0) Then
        MsgBox "Selected file is not valid Ant build!", vbCritical, "Parser error"
    End If
    
    ReDim result(targets.length - 1)
    
    For Each target In targets
        result(i) = target.Attributes.getNamedItem("name").Text
        i = i + 1
    Next
    getTargets = result
End Function
