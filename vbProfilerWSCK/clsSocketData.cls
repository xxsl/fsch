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

Private Declare Sub GetMem4 Lib "msvbvm60" (ByVal Ptr As Long, value As Any)

Private Declare Sub GetMem2 Lib "msvbvm60" (ByVal Ptr As Long, value As Any)

Public socket As Winsock

Public bytesAvailable As Long

Private cursor As Long

Private buff() As Byte

Public Sub Refresh()
    cursor = 0
    socket.PeekData buff, vbArray
End Sub

Public Sub Complete()
    If (cursor > 0) Then
        socket.GetData buff, vbArray, cursor
    End If
End Sub

Public Sub skip(ByVal bytes As Long)
    cursor = cursor + bytes
    bytesAvailable = bytesAvailable - bytes
End Sub

Public Sub getPos(ByRef Pos As Long)
     Pos = cursor + 0
End Sub

Public Function ReadData(ByRef Buffer() As Byte, ByVal rSize As Long)
    ReDim Buffer(rSize - 1)
    CopyMemB Buffer(0), buff(cursor), rSize
    cursor = cursor + rSize
    bytesAvailable = bytesAvailable - rSize
End Function

Public Sub read32(ByRef result As Long)
    'GetMem4 VarPtr(buff(cursor)), result
    CopyMemB result, buff(cursor), 4&
    result = ntohl(result)
    cursor = cursor + 4
    bytesAvailable = bytesAvailable - 4
End Sub

Public Sub read16(ByRef result As Integer)
    'GetMem2 VarPtr(buff(cursor)), result
    CopyMemB result, buff(cursor), 2&
    result = ntohs(result)
    cursor = cursor + 2
    bytesAvailable = bytesAvailable - 2
End Sub

Public Function ReadDataP(ByRef position As Long, ByRef Buffer() As Byte, ByVal rSize As Long)
    ReDim Buffer(rSize - 1)
    CopyMemB Buffer(0), buff(position), rSize
    position = position + rSize
End Function

Public Sub read32P(ByRef position As Long, ByRef result As Long)
    'GetMem4 VarPtr(buff(position)), result
    CopyMemB result, buff(position), 4&
    result = ntohl(result)
    position = position + 4
End Sub

Public Sub read16P(ByRef position As Long, ByRef result As Integer)
    'GetMem2 VarPtr(buff(position)), result
    CopyMemB result, buff(position), 2&
    result = ntohs(result)
    position = position + 2
End Sub

Public Function ReadData8(ByRef result As Byte)
    result = buff(cursor)
    cursor = cursor + 1
    bytesAvailable = bytesAvailable - 1
End Function




