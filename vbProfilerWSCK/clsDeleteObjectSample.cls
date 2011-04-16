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

Public Sub parse(ByRef Data As clsSocketData)
    Data.skip 8
    Data.read32 id
    Data.read32 Size
End Sub

Public Sub bytesAvailable(ByRef Data As clsSocketData, ByRef result As Boolean)
    result = Data.bytesAvailable >= dataSize
End Sub


Public Function getId() As Long
    getId = id
End Function


Public Function getSize() As Long
    getSize = Size
End Function




