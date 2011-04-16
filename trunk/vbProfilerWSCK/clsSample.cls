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

Public strings As cCollection

Private value As New clsStackTrace

Private dataSize As Long

Public Function canDeserialise(firstByte As Byte) As Boolean
    canDeserialise = (firstByte = C_SAMPLE)
End Function


Public Sub parse(ByRef Data As clsSocketData)
    Data.skip 8
    
    Dim Size As Long, i As Long
    
    Data.read32 Size

    Dim lng As Long
    For i = 1 To Size
        Dim line As New clsStackTraceLine
        
        Data.read32 lng
        line.Name = strings.item(lng)
        Data.read32 lng
        line.FileName = strings.item(lng)
        Data.read32 lng
        line.line = lng
        value.lines.Add line
    Next i
End Sub

Public Function bytesAvailable(ByRef Data As clsSocketData) As Boolean

    Dim Length As Long, result As Boolean, Pos As Long

    If (Data.bytesAvailable >= 12) Then
        Data.getPos Pos
        Pos = Pos + 8
        Data.read32P Pos, Length
        dataSize = Length * 12 + 12
        result = (dataSize <= Data.bytesAvailable)
    Else
        result = False
    End If

    bytesAvailable = result
End Function

Public Function getValue() As clsStackTrace
    getValue = value
End Function


