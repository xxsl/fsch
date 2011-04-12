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

Private Size   As Long

Private dataSize As Long

Private Sub Class_Initialize()
    dataSize = 16
End Sub

Public Function canDeserialise(firstByte As Byte) As Boolean
    canDeserialise = (firstByte = C_DELETE_OBJECT_SAMPLE)
End Function


Public Sub parse(ByRef Data As clsSocketData)
    Dim buff() As Byte, Length As Long, Pos As Long
    
    Data.ReadData buff, 16
    Pos = 8

    read32P buff, Pos, id
    read32P buff, Pos, Size
End Sub

Public Function bytesAvailable(ByRef Data As clsSocketData) As Boolean

    Dim Length As Long, result As Boolean

    If (Data.bytesAvailable >= dataSize) Then
        result = True
    Else
        result = False
    End If

    bytesAvailable = result
End Function


Public Function getId() As Long
    getId = id
End Function


Public Function getSize() As Long
    getSize = Size
End Function




