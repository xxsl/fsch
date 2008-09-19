VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cldTarget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public fCommand As String
Public fName As String
Public fSource As String
Public fOutput As String
Public fLibraries() As String
Public fDebug As Boolean


Public Function getExecCommand() As String
    Dim result As String
    result = fCommand + " " + fSource + " " + fOutput + " " + fDebug
End Function

'
Public Property Get OutputPath() As String
    If (Len(Trim(fOutput)) > 0) Then
        OutputPath = OUTPUT_PATH + fOutput + " "
    Else
        OutputPath = ""
    End If
End Property

Public Property Let OutputPath(ByVal vNewValue As String)

End Property


Public Property Get OutputPath() As String

End Property

Public Property Let OutputPath(ByVal vNewValue As String)

End Property

Public Property Get OutputPath() As String

End Property

Public Property Let OutputPath(ByVal vNewValue As String)

End Property

Public Property Get OutputPath() As String

End Property

Public Property Let OutputPath(ByVal vNewValue As String)

End Property

Public Property Get OutputPath() As String

End Property

Public Property Let OutputPath(ByVal vNewValue As String)

End Property

Public Property Get OutputPath() As String

End Property

Public Property Let OutputPath(ByVal vNewValue As String)

End Property
