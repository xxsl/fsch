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

Private dataSize As Long

Public Function canDeserialise(firstByte As Byte) As Boolean
    canDeserialise = (firstByte = C_STRING)
End Function

Public Sub parse(ByRef Data As clsSocketData)
    Dim buff() As Byte, length As Integer, pos As Long, lng As Long
    Data.ReadData buff, dataSize
    read16P buff, pos, length
    
    If (length > 0) Then
        Dim strbuff() As Byte
        ReDim strbuff(length)
        lng = length
        CopyMemory strbuff(0), buff(pos), lng
        pos = pos + lng
        ConvertUtf8BytesToString strbuff, Value
    Else
        Value = ""
    End If

    read32P buff, pos, key
End Sub

Public Function bytesAvailable(ByRef Data As clsSocketData) As Boolean

    Dim length As Integer, result As Boolean, bytes() As Byte

    If (Data.socket.BytesReceived >= 2) Then
        Data.getBytes bytes
        read16P bytes, 0, length
        dataSize = length + 6
        result = (dataSize <= Data.socket.BytesReceived)
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

