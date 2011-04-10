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

Private Const x32 As Long = 4
Private Const x16 As Long = 2
Private Const x8 As Long = 1

Public socket As Winsock

Private bytes()   As Byte

Public Size    As Long

Private Sub Class_Initialize()
    ReDim bytes(2048)
End Sub

Public Sub append(ByRef Buffer() As Byte)
    Dim Length   As Long
    
    Length = UBound(Buffer) + 1
    
    If (Size + Length - 1 > 1024) Then
        ReDim Preserve bytes(Size + Length - 1)
    End If
    
    CopyMemory bytes(Size), Buffer(0), Length
    
    Size = Size + Length
End Sub


Public Sub getBytes(ByRef Buffer() As Byte)
    Buffer = bytes
End Sub

Public Sub PeekData(ByRef Buffer() As Byte, ByVal Size As Long)
    ReDim Buffer(Size - 1&)
   
    CopyMemory Buffer(0), bytes(0), Size
End Sub

Public Function ReadData(ByRef Buffer() As Byte, ByVal rSize As Long)

    PeekData Buffer, rSize
    
    If (Size - rSize > 0) Then
        CopyMemory bytes(0), bytes(rSize), Size - rSize
    End If
    
    Size = Size - rSize
End Function

Public Function ReadData32(ByRef Buffer() As Byte)
      
    CopyMemory Buffer(0), bytes(0), x32
    
    If (Size - x32 > 0) Then
        CopyMemory bytes(0), bytes(x32), Size - x32
    End If
    
    Size = Size - x32
End Function

Public Function ReadData16(ByRef Buffer() As Byte)
      
    CopyMemory Buffer(0), bytes(0), x16
    
    If (Size - x16 > 0) Then
        CopyMemory bytes(0), bytes(x16), Size - x16
    End If
    
    Size = Size - x16
End Function

Public Function ReadData8(ByRef Buffer() As Byte)
      
    CopyMemory Buffer(0), bytes(0), x8
    
    If (Size - x8 > 0) Then
        CopyMemory bytes(0), bytes(x8), Size - x8
    End If
    
    Size = Size - x8
End Function

