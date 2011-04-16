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

Public strings   As cCollection

Private value    As New clsStackTrace

Private id       As Long

Private typeName As String

Private url      As String

Private Size     As Long

Private dataSize As Long

Private success  As Boolean

Public Function canDeserialise(firstByte As Byte) As Boolean
    canDeserialise = (firstByte = C_NEW_OBJECT_SAMPLE)
End Function

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

            Dim lng As Long
    
            For i = 1 To Count

                Dim line As New clsStackTraceLine

                Data.read32P Pos, lng
                line.Name = strings.item(lng)
                Data.read32P Pos, lng
                line.FileName = strings.item(lng)
                Data.read32P Pos, lng
                line.line = lng
                value.lines.Add line
            Next i
    
            Data.read32P Pos, id
    
            Data.read32P Pos, lng

            typeName = strings.item(lng)

            Data.read32P Pos, lng
            url = strings.item(lng)
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

Public Function getType() As String
    getType = typeName
End Function

Public Function getURL() As String
    getURL = url
End Function

Public Function getSize() As Long
    getSize = Size
End Function

