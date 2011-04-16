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

Public dataSize As Long

Public Sub parse(ByRef Data As clsSocketData)
    Data.skip dataSize
End Sub

Public Sub bytesAvailable(ByRef Data As clsSocketData, ByRef result As Boolean)

    Dim Pos As Long

    If (Data.bytesAvailable >= 4) Then
        Data.getPos Pos
        Data.read32P Pos, key
        Data.read32P Pos, Size
        dataSize = Size + 8&
        result = (dataSize <= Data.bytesAvailable)
        
        If (result) Then
            Data.ReadDataP Pos, value, Size
        End If
    Else
        result = False
    End If
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
