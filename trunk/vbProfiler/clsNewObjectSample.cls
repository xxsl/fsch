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

Private value As New clsStackTrace

Private id   As Long

Private typeName   As String

Private url   As String

Private size   As Long

Public Function canDeserialise(firstByte As Byte) As Boolean
    canDeserialise = (firstByte = C_NEW_OBJECT_SAMPLE)
End Function


Public Sub parse(ByRef data As clsSocketData)
    Dim buff() As Byte
    buff = readX(data, 8)
    
    Dim count As Long, i As Long
    count = read32X(data)

    For i = 1 To count
        Dim line As New clsStackTraceLine
        line.name = strings.Item(read32X(data))
        line.fileName = strings.Item(read32X(data))
        line.line = read32X(data)
        value.lines.Add line
    Next i
    
    id = read32X(data)
    
    typeName = strings.Item(read32X(data))
    url = strings.Item(read32X(data))
    size = read32X(data)
End Sub

Public Function bytesAvailable(ByRef data As clsSocketData) As Boolean

    Dim length As Long, Result As Boolean

    If (data.size >= 12) Then
        length = read32P(data.getBytes, 8)
        Result = ((length * 12 + 28) <= data.size)
    Else
        Result = False
    End If

    bytesAvailable = Result
End Function

Public Function getValue() As clsStackTrace
    getValue = value
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
    getSize = size
End Function


