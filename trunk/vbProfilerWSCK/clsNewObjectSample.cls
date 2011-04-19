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

Private value    As New clsStackTrace

Private id       As Long

Private typeName As Long

Private url      As Long

Private Size     As Long

Private dataSize As Long

Private success  As Boolean

Public Sub parse(ByRef Data As clsSocketData)
     Data.skip dataSize
End Sub

Public Sub bytesAvailable(ByRef Data As clsSocketData, ByRef result As Boolean)

    result = Data.bytesAvailable >= 12&

    If (result) Then

        Dim Count As Long, Pos As Long, i As Long

        Data.getPos Pos
        Pos = Pos + 8&
    
        Data.read32P Pos, Count
        
        dataSize = Count * 12& + 28&
        result = (dataSize <= Data.bytesAvailable)

        If (result) Then

'            For i = 1 To Count
'                Dim line As New clsStackTraceLine
'
'                Data.read32P Pos, line.Name
'                Data.read32P Pos, line.FileName
'                Data.read32P Pos, line.line
'                value.lines.Add line
'            Next i
            Pos = Pos + Count * 12
    
            Data.read32P Pos, id
            Data.read32P Pos, typeName
            Data.read32P Pos, url
            Data.read32P Pos, Size
            
            success = True
        End If
        
    End If

End Sub

Public Function getValue() As clsStackTrace
    getValue = value
End Function

Public Function getId() As Long
    getId = id
End Function

Public Function getType() As Long
    getType = typeName
End Function

Public Function getURL() As Long
    getURL = url
End Function

Public Function getSize() As Long
    getSize = Size
End Function

