VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AntTarget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Build As AntBuild

Public Name As String
Public Depends As String
Public Description As String


Public Sub fromXMLNode(node As IXMLDOMElement)
    Dim attr1 As IXMLDOMAttribute
    Set attr1 = node.Attributes.getNamedItem("name")
    If (Not attr1 Is Nothing) Then
        Name = attr1.Text
    Else
        Name = "Not Set"
    End If
    
    Dim attr2 As IXMLDOMAttribute
    Set attr2 = node.Attributes.getNamedItem("depends")
    If (Not attr2 Is Nothing) Then
        Depends = attr2.Text
    Else
        Depends = "Not Set"
    End If
    
    Dim attr3 As IXMLDOMAttribute
    Set attr3 = node.Attributes.getNamedItem("description")
    If (Not attr3 Is Nothing) Then
        Description = attr3.Text
    Else
        Description = "Not Set"
    End If
End Sub

Public Function getCommand()
    getCommand = "ant " & Build.buildfile & " " & Name
End Function
