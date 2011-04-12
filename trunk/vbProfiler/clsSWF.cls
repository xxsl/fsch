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

Private size    As Long

Private key     As Long

Private dataSize As Long

Public Function canDeserialise(firstByte As Byte) As Boolean
    canDeserialise = (firstByte = C_SWF)
End Function

Public Sub parse(ByRef Data As clsSocketData)
    Dim Buffer() As Byte, pos As Long, Length As Long
    Data.ReadData Buffer, dataSize
    read32P Buffer, pos, key
    read32P Buffer, pos, size
    Length = UBound(Buffer) - pos
    ReDim Value(Length - 1)
    CopyMemory Value(0), Buffer(pos), Length
End Sub

Public Function bytesAvailable(ByRef Data As clsSocketData) As Boolean

    Dim Length As Long, result As Boolean, bytes() As Byte, pos As Long

    If (Data.bytesAvailable >= 4) Then
        Data.getBytes bytes, pos
        pos = pos + 4
        read32P bytes, pos, Length
        dataSize = Length + 8
        result = (dataSize <= Data.bytesAvailable)
    Else
        result = False
    End If

    bytesAvailable = result
End Function

Public Function getValue() As Byte()
    getValue = Value
End Function

Public Function getSize() As Long
    getSize = size
End Function

Public Function getKey() As Long
    getKey = key
End Function
