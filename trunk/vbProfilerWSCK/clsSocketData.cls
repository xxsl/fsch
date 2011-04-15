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
    CopyMemory Buffer(0), buff(cursor), rSize
    cursor = cursor + rSize
    bytesAvailable = bytesAvailable - rSize
End Function

Public Sub read32(ByRef result As Long)
    'GetMem4 VarPtr(buff(cursor)), result
    CopyMemory result, buff(cursor), 4&
    result = ntohl(result)
    cursor = cursor + 4
    bytesAvailable = bytesAvailable - 4
End Sub

Public Sub read16(ByRef result As Integer)
    'GetMem2 VarPtr(buff(cursor)), result
    CopyMemory result, buff(cursor), 2&
    result = ntohs(result)
    cursor = cursor + 2
    bytesAvailable = bytesAvailable - 2
End Sub

Public Sub read32P(ByRef position As Long, ByRef result As Long)
    'GetMem4 VarPtr(buff(position)), result
    CopyMemory result, buff(position), 4&
    result = ntohl(result)
    position = position + 4
End Sub

Public Sub read16P(ByRef position As Long, ByRef result As Integer)
    'GetMem2 VarPtr(buff(position)), result
    CopyMemory result, buff(position), 2&
    result = ntohs(result)
    position = position + 2
End Sub

Public Function ReadData8(ByRef result As Byte)
    result = buff(cursor)
    cursor = cursor + 1
    bytesAvailable = bytesAvailable - 1
End Function


'Public Function ReadData32(ByRef result As Long)
'    socket.GetData result, vbLong
'    result = ntohl(result)
'End Function
'
'Public Function ReadData16(ByRef result As Integer)
'    socket.GetData result, vbInteger
'    result = ntohs(result)
'End Function

'Public Sub readString(ByRef result As String)
'    Dim strBuff() As Byte
'    Dim Length    As Integer
'
'    ReadData16 Length
'
'    If (Length > 0) Then
'        ReadData strBuff, Length
'        ConvertUtf8BytesToString strBuff, result
'    Else
'        result = ""
'    End If
'End Sub


