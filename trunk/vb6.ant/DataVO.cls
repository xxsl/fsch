VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataVO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public data As String
Public target As String

Public Function isClass(ByRef dataInput() As Byte, Optional ByRef pos As Long = 0) As Boolean
     Dim result As String
     result = readString(dataInput, pos)
     isClass = (result = AIR_DATAVO)
End Function

Public Sub deSerialize(ByRef dataInput() As Byte, ByRef pos As Long)
    data = readString(dataInput, pos)
    target = readString(dataInput, pos)
End Sub

Public Sub serialize(ByRef dataOutput() As Byte)
    writeString dataOutput, AIR_DATAVO
    writeString dataOutput, data
    writeString dataOutput, target
End Sub

Public Function toString() As String
    toString = "Class=" & AIR_DATAVO & "; target=" & target & "; data=" & data & ";"
End Function
