VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNewObjectSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'    private void readNewObjectSample() throws IOException
'    {
'        long time = read64() / 1000L;
'        StackTrace stackTrace = null;
'
'        if (this.memoryStackTracesEnabled)
'        {
'            stackTrace = readStackTrace();
'        } else
'        {
'            stackTrace = new StackTrace(0);
'        }
'
'        int id = read32();
'        String type = this.playerSession.getString(read32());
'        String url = this.playerSession.getString(read32());
'        int size = read32();
'
'        initTimeOffset(time);
'        this.playerSession.newObjectSample(time + this.timeOffset, stackTrace, id, type, url, size);
'    }

Public strings As Collection

Private Value As New clsStackTrace

Private id   As Long

Private typeName   As String

Private url   As String

Private Size   As Long

Private dataSize As Long

Public Function canDeserialise(firstByte As Byte) As Boolean
    canDeserialise = (firstByte = C_NEW_OBJECT_SAMPLE)
End Function


Public Sub parse(ByRef Data As clsSocketData)
    Data.skip 8
    
    Dim Count As Long, i As Long
    
    Data.read32 Count

    Dim lng As Long
    
    For i = 1 To Count
        Dim line As New clsStackTraceLine
        Data.read32 lng
        line.NAME = strings.item(lng)
        Data.read32 lng
        line.FileName = strings.item(lng)
        Data.read32 lng
        line.line = lng
        Value.lines.Add line
    Next i
    
    Data.read32 id
    
    Data.read32 lng
    typeName = strings.item(lng)
    Data.read32 lng
    url = strings.item(lng)
    Data.read32 Size
End Sub

Public Function bytesAvailable(ByRef Data As clsSocketData) As Boolean

    Dim Length As Long, result As Boolean, Pos As Long

    If (Data.bytesAvailable >= 12) Then
        Data.getPos Pos
        Pos = Pos + 8
        Data.read32P Pos, Length
        dataSize = Length * 12 + 28
        result = (dataSize <= Data.bytesAvailable)
    Else
        result = False
    End If

    bytesAvailable = result
End Function

Public Function getValue() As clsStackTrace
    getValue = Value
End Function

Public Function getId() As Long
    getId = id
End Function

Public Function getType() As String
    getType = typeName
End Function

Public Function getURL() As String
    getURL = url
End Function

Public Function getSize() As Long
    getSize = Size
End Function


