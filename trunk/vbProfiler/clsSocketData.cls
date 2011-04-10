VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSocketData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public socket As Winsock

Private bytes()   As Byte

Public Size    As Long

Private Sub Class_Initialize()
    ReDim bytes(1024)
End Sub

Public Sub append(ByRef buffer() As Byte)
    Dim Length   As Long
    
    Length = UBound(buffer) + 1
    
    If (Size + Length - 1 > 1024) Then
        ReDim Preserve bytes(Size + Length - 1)
    End If
    
    CopyMemory bytes(Size), buffer(0), Length
    
    Size = Size + Length
End Sub


Public Sub getBytes(ByRef buffer() As Byte)
    buffer = bytes
End Sub

Public Sub PeekData(ByRef buffer() As Byte, ByVal Size As Long)
    
    ReDim Preserve buffer(Size - 1)
   
    CopyMemory buffer(0), bytes(0), Size
End Sub

Public Function ReadData(ByRef buffer() As Byte, ByVal rSize As Long)

    PeekData buffer, rSize
    
    If (Size - rSize > 0) Then
        CopyMemory bytes(0), bytes(rSize), Size - rSize
    End If
    
    Size = Size - rSize
End Function


