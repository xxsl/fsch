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

Private Value As String

Private key   As Long

Private dataSize As Long

Public Function canDeserialise(firstByte As Byte) As Boolean
    canDeserialise = (firstByte = C_STRING)
End Function

Public Sub parse(ByRef Data As clsSocketData)
    Dim buff() As Byte, Length As Integer, Pos As Long, lng As Long
    Data.ReadData buff, dataSize
    read16P buff, Pos, Length
    
    If (Length > 0) Then
        Dim strBuff() As Byte
        ReDim strBuff(Length)
        lng = Length
        CopyMemory strBuff(0), buff(Pos), lng
        Pos = Pos + lng
        ConvertUtf8BytesToString strBuff, Value
    Else
        Value = ""
    End If

    read32P buff, Pos, key
End Sub

Public Function bytesAvailable(ByRef Data As clsSocketData) As Boolean

    Dim Length As Integer, result As Boolean, Pos As Long

    If (Data.bytesAvailable >= 2) Then
        Data.getPos Pos
        Data.read16P Pos, Length
        dataSize = Length + 6
        result = (dataSize <= Data.bytesAvailable)
    Else
        result = False
    End If

    bytesAvailable = result
End Function

Public Function getValue() As String
    getValue = Value
End Function

Public Function getKey() As Long
    getKey = key
End Function

