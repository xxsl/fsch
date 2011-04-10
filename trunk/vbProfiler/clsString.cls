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

Private Value As String

Private key   As Long

Public Function canDeserialise(firstByte As Byte) As Boolean
    canDeserialise = (firstByte = C_STRING)
End Function

Public Sub parse(ByRef Data As clsSocketData)
    Value = readString(Data)
    key = read32X(Data)
End Sub

Public Function bytesAvailable(ByRef Data As clsSocketData) As Boolean

    Dim Length As Long, result As Boolean, bytes() As Byte

    If (Data.Size >= 2) Then
        Data.getBytes bytes
        Length = read16P(bytes, 0)
        result = ((Length + 6) <= Data.Size)
    Else
        result = False
    End If

    bytesAvailable = result
End Function

Public Function getValue() As String
    getValue = Value
End Function

Public Function getKey() As Long
    getKey = key
End Function

