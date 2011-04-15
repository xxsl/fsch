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

Private Declare Function ntohl Lib "wsock32.dll" (ByVal hostlong As Long) As Long

Private Declare Function ntohs Lib "wsock32.dll" (ByVal hostlong As Integer) As Integer

Public socket As SocketWrench

Public Function ReadData(ByRef Buffer() As Byte, ByVal rSize As Long)
    ReDim Buffer(rSize - 1)
    socket.ReadStream Buffer, rSize
End Function

Public Sub read32(ByRef result As Long)
    Dim lngBuff(3) As Byte
    socket.ReadStream lngBuff, 4&
    CopyMemory result, lngBuff(0), 4&
    result = ntohl(result)
End Sub

Public Sub read16(ByRef result As Integer)
    Dim intBuff(1) As Byte
    socket.ReadStream intBuff, 2&
    CopyMemory result, intBuff(0), 2&
    result = ntohs(result)
End Sub

Public Function Read8(ByRef result As Byte)
    Dim byteBuff(0) As Byte
    socket.ReadStream byteBuff, 1&
    result = byteBuff(0)
End Function
