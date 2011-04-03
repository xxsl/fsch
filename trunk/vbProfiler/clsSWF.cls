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

Private value() As Byte

Private size    As Long

Private key     As Long

Public Function canDeserialise(firstByte As Byte) As Boolean
    canDeserialise = (firstByte = C_SWF)
End Function

Public Sub parse(ByRef data As clsSocketData)
    key = read32X(data)
    size = read32X(data)
    value = readX(data, size)
End Sub

Public Function bytesAvailable(ByRef data As clsSocketData) As Boolean

    Dim length As Long, Result As Boolean

    If (data.size >= 4) Then
        length = read32P(data.getBytes, 4)
        Result = ((length + 8) <= data.size)
    Else
        Result = False
    End If

    bytesAvailable = Result
End Function

Public Function getValue() As Byte()
    getValue = value
End Function

Public Function getSize() As Long
    getSize = size
End Function

Public Function getKey() As Long
    getKey = key
End Function
