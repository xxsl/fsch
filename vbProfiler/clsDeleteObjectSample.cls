VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDeleteObjectSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'    private void readDeleteObjectSample() throws IOException
'    {
'        long time = read64() / 1000L;
'        int id = read32();
'        int size = read32();
'
'        initTimeOffset(time);
'
'        this.playerSession.deleteObjectSample(time + this.timeOffset, id, size);
'    }

Private id   As Long

Private size   As Long

Public Function canDeserialise(firstByte As Byte) As Boolean
    canDeserialise = (firstByte = C_DELETE_OBJECT_SAMPLE)
End Function


Public Sub parse(ByRef data As clsSocketData)
    Dim buff() As Byte
    buff = readX(data, 8)

    id = read32X(data)

    size = read32X(data)
End Sub

Public Function bytesAvailable(ByRef data As clsSocketData) As Boolean

    Dim length As Long, Result As Boolean

    If (data.size >= 16) Then
        Result = True
    Else
        Result = False
    End If

    bytesAvailable = Result
End Function


Public Function getId() As Long
    getId = id
End Function


Public Function getSize() As Long
    getSize = size
End Function



