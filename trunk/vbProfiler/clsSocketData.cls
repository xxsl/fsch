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

Public Sub getBytes(ByRef Buffer() As Byte)
     socket.PeekData Buffer, vbArray
End Sub

Public Function ReadData(ByRef Buffer() As Byte, ByVal rSize As Long)
    socket.GetData Buffer, vbArray, rSize
End Function

Public Function ReadData32(ByRef result As Long)
    socket.GetData result, vbLong
    result = ntohl(result)
End Function

Public Function ReadData16(ByRef result As Integer)
    socket.GetData result, vbInteger
    result = ntohs(result)
End Function

Public Function ReadData8(ByRef result As Byte)
    socket.GetData result, vbByte
End Function

Public Sub readString(ByRef result As String)
    Dim strBuff() As Byte
    Dim Length    As Integer

    ReadData16 Length
   
    If (Length > 0) Then
        ReadData strBuff, Length
        ConvertUtf8BytesToString strBuff, result
    Else
        result = ""
    End If
End Sub


