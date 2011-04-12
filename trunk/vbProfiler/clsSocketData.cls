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

Public socket As Winsock

Public bytesAvailable As Long

Private buff() As Byte

Public Sub Refresh()
    socket.PeekData buff, vbArray
End Sub

Public Sub getPos(ByRef Pos As Long)
     Pos = UBound(buff) + 1 - bytesAvailable
End Sub

Public Sub getBytes(ByRef Buffer() As Byte, ByRef Pos As Long)
     Pos = UBound(buff) + 1 - bytesAvailable
     Buffer = buff
End Sub

Public Function ReadData(ByRef Buffer() As Byte, ByVal rSize As Long)
    socket.GetData Buffer, vbArray, rSize
    bytesAvailable = bytesAvailable - rSize
End Function


Public Sub read32P(ByRef position As Long, ByRef result As Long)
    CopyMemory result, buff(position), 4&
    result = ntohl(result)
    position = position + 4
End Sub

Public Sub read16P(ByRef position As Long, ByRef result As Integer)
    CopyMemory result, buff(position), 2&
    result = ntohs(result)
    position = position + 2
End Sub

'Public Function ReadData32(ByRef result As Long)
'    socket.GetData result, vbLong
'    result = ntohl(result)
'End Function
'
'Public Function ReadData16(ByRef result As Integer)
'    socket.GetData result, vbInteger
'    result = ntohs(result)
'End Function

Public Function ReadData8(ByRef result As Byte)
    socket.GetData result, vbByte
    bytesAvailable = bytesAvailable - 1
End Function

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


