VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private value As String

Private key   As Long

Public Function canDeserialise(firstByte As Byte) As Boolean
    canDeserialise = (firstByte = C_STRING)
End Function

Public Sub parse(ByRef data As clsSocketData)
    value = readString(data)
    key = read32X(data)
End Sub

Public Function bytesAvailable(ByRef data As clsSocketData) As Boolean

    Dim length As Long, Result As Boolean

    If (data.size >= 2) Then
        length = read16P(data.getBytes, 0)
        Result = ((length + 6) <= data.size)
    Else
        Result = False
    End If

    bytesAvailable = Result
End Function

Public Function getValue() As String
    getValue = value
End Function

Public Function getKey() As Long
    getKey = key
End Function

