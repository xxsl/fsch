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

Private Size    As Long

Private key     As Long

Public Function canDeserialise(firstByte As Byte) As Boolean
    canDeserialise = (firstByte = C_SWF)
End Function

Public Sub parse(ByRef Data As clsSocketData)
    Dim Length As Long
    Data.read32 key
    Data.read32 Size
    Data.ReadData value, Size
End Sub

Public Function getValue() As Byte()
    getValue = value
End Function

Public Function getSize() As Long
    getSize = Size
End Function

Public Function getKey() As Long
    getKey = key
End Function
