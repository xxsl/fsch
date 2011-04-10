VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'    private void readSample() throws IOException
'    {
'        long time = read64() / 1000L;
'        StackTrace stackTrace = readStackTrace();
'
'        initTimeOffset(time);
'
'        this.playerSession.sample(time + this.timeOffset, stackTrace);
'    }

'    private StackTrace readStackTrace() throws IOException
'    {
'        int length = read32();
'        StackTrace result = new StackTrace(length);
'
'        for (int i = 0; i < length; i++)
'        {
'            result.setName(i, this.playerSession.getString(read32()));
'            result.setFileName(i, this.playerSession.getString(read32()));
'            result.setLineNumber(i, read32());
'        }
'
'        return result;
'    }

Public strings As Collection

Private Value As New clsStackTrace

Public Function canDeserialise(firstByte As Byte) As Boolean
    canDeserialise = (firstByte = C_SAMPLE)
End Function


Public Sub parse(ByRef Data As clsSocketData)
    Dim buff() As Byte
    readX buff, Data, 8
    
    Dim Size As Long, i As Long
    Size = read32X(Data)

    For i = 1 To Size
        Dim line As New clsStackTraceLine
        line.NAME = strings.item(read32X(Data))
        line.FileName = strings.item(read32X(Data))
        line.line = read32X(Data)
        Value.lines.Add line
    Next i
End Sub

Public Function bytesAvailable(ByRef Data As clsSocketData) As Boolean

    Dim Length As Long, result As Boolean, bytes() As Byte

    If (Data.Size >= 12) Then
        Data.getBytes bytes
        Length = read32P(bytes, 8)
        result = ((Length * 12 + 12) <= Data.Size)
    Else
        result = False
    End If

    bytesAvailable = result
End Function

Public Function getValue() As clsStackTrace
    getValue = Value
End Function


