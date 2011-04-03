VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const SECTION As String = "custom"

Public extractedFiles As New Collection

Public Function getResourceByName(NAME As String, Optional resource As String = ".png") As String
    Dim Data() As Byte
    Dim path As String
    
    Data = LoadResData(NAME, SECTION)
    path = getPath & NAME & resource
    
    Open path For Output As #2
    Close #2
    Open path For Binary As #2
        Put #2, , Data()
    Close #2
    extractedFiles.Add path
    getResourceByName = path
End Function

Public Function getIndex(NAME As String) As Long
    Dim i As Long
    Dim item As Variant
    Dim path As String
    For Each item In extractedFiles
        i = i + 1
        path = CStr(item)
        If (InStr(1, path, NAME) > 0) Then
            getIndex = i
            Exit Function
        End If
    Next
    getIndex = -1
End Function

Public Sub clear()
    Dim item As Variant
    Dim path As String
    
    For Each item In extractedFiles
        path = CStr(item)
        If (FileExists(path)) Then
            Kill path
        End If
    Next
    Set extractedFiles = New Collection
End Sub
