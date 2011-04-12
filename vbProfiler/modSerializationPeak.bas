Attribute VB_Name = "modSerializationPeak"
Option Explicit

Public Sub read32P(ByRef buffer() As Byte, ByRef position As Long, ByRef result As Long)

    Dim strBuff() As Byte

    Dim i As Long

    ReDim strBuff(3)
    
    For i = (position + 3) To position Step -1
        strBuff(result) = buffer(i)
        result = result + 1
    Next i
    
    CopyMemory result, strBuff(0), 4&

    position = position + 4
End Sub

Public Sub read16P(ByRef buffer() As Byte, ByRef position As Long, ByRef result As Integer)

    Dim strBuff(0 To 1) As Byte

    Dim i As Long

    For i = (position + 1) To position Step -1
        strBuff(result) = buffer(i)
        result = result + 1
    Next i
    
    CopyMemory result, strBuff(0), 2&

    position = position + 2
End Sub

