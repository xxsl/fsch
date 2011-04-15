VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private value As String

Private key   As Long

Public Function canDeserialise(firstByte As Byte) As Boolean
    canDeserialise = (firstByte = C_STRING)
End Function

Public Sub parse(ByRef Data As clsSocketData)
    Dim Length As Integer, Pos As Long, lng As Long
    
    Data.read16 Length
    
    If (Length > 0) Then
        Dim strBuff() As Byte
        lng = Length
        Data.ReadData strBuff, lng
        ConvertUtf8BytesToString strBuff, value
    Else
        value = ""
    End If

    Data.read32 key
End Sub


Public Function getValue() As String
    getValue = value
End Function

Public Function getKey() As Long
    getKey = key
End Function

