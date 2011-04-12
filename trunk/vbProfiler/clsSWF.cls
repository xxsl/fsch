VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSWF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Value() As Byte

Private Size    As Long

Private key     As Long

Private dataSize As Long

Public Function canDeserialise(firstByte As Byte) As Boolean
    canDeserialise = (firstByte = C_SWF)
End Function

Public Sub parse(ByRef Data As clsSocketData)
    Dim buffer() As Byte, pos As Long, length As Long
    Data.ReadData buffer, dataSize
    read32P buffer, pos, key
    read32P buffer, pos, Size
    length = UBound(buffer) - pos
    ReDim Value(length - 1)
    CopyMemory Value(0), buffer(pos), length
End Sub

Public Function bytesAvailable(ByRef Data As clsSocketData) As Boolean

    Dim length As Long, result As Boolean, bytes() As Byte

    If (Data.socket.BytesReceived >= 4) Then
        Data.getBytes bytes
        read32P bytes, 4, length
        dataSize = length + 8
        result = (dataSize <= Data.socket.BytesReceived)
    Else
        result = False
    End If

    bytesAvailable = result
End Function

Public Function getValue() As Byte()
    getValue = Value
End Function

Public Function getSize() As Long
    getSize = Size
End Function

Public Function getKey() As Long
    getKey = key
End Function
