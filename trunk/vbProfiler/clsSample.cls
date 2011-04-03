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

Private value As New clsStackTrace

Public Function canDeserialise(firstByte As Byte) As Boolean
    canDeserialise = (firstByte = C_SAMPLE)
End Function


Public Sub parse(ByRef data As clsSocketData)
    Dim buff() As Byte
    buff = readX(data, 8)
    
    Dim size As Long, i As Long
    size = read32X(data)

    For i = 1 To size
        Dim line As New clsStackTraceLine
        line.name = strings.Item(read32X(data))
        line.fileName = strings.Item(read32X(data))
        line.line = read32X(data)
        value.lines.Add line
    Next i
End Sub

Public Function bytesAvailable(ByRef data As clsSocketData) As Boolean

    Dim length As Long, Result As Boolean

    If (data.size >= 12) Then
        length = read32P(data.getBytes, 8)
        Result = ((length * 12 + 12) <= data.size)
    Else
        Result = False
    End If

    bytesAvailable = Result
End Function

Public Function getValue() As clsStackTrace
    getValue = value
End Function


