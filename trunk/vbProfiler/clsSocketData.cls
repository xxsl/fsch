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


Private Const x32 As Long = 4
Private Const x16 As Long = 2
Private Const x8 As Long = 1

Public socket As Winsock


Public Sub getBytes(ByRef Buffer() As Byte)
    socket.PeekData Buffer, vbArray
End Sub

Public Sub PeekData(ByRef Buffer() As Byte, ByVal Size As Long)
    socket.PeekData Buffer, vbArray, Size
End Sub

Public Function ReadData(ByRef Buffer() As Byte, ByVal rSize As Long)
    socket.GetData Buffer, vbArray, rSize
End Function

Public Function ReadData32(ByRef result As Long)
    socket.GetData result, vbLong, x32
    result = ntohl(result)
End Function

Public Function ReadData16(ByRef result As Integer)
    socket.GetData result, vbInteger, x16
    result = ntohs(result)
End Function

Public Function ReadData8(ByRef result As Byte)
    socket.GetData result, vbByte, x8
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
