VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function canDeserialise(firstByte As Byte) As Boolean
    canDeserialise = (firstByte = C_GC)
End Function

Public Sub parse(ByRef data As clsSocketData)

End Sub

Public Function bytesAvailable(ByRef data As clsSocketData) As Boolean
    bytesAvailable = True
End Function

