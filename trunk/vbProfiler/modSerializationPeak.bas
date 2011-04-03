Attribute VB_Name = "modSerializationPeak"
Option Explicit

Public Function read32P(ByRef buffer() As Byte, ByRef position As Long) As Long

    Dim strBuff() As Byte

    Dim length    As Long, i As Long

    ReDim strBuff(3)
    
    For i = (position + 3) To position Step -1
        strBuff(length) = buffer(i)
        length = length + 1
    Next i
    
    CopyMemory length, strBuff(0), 4&

    position = position + 4
    read32P = length
End Function

Public Function read16P(ByRef buffer() As Byte, ByRef position As Long) As Integer

    Dim strBuff() As Byte

    Dim length    As Long, i As Long

    ReDim strBuff(1)
    
    For i = (position + 1) To position Step -1
        strBuff(length) = buffer(i)
        length = length + 1
    Next i
    
    CopyMemory length, strBuff(0), 2&

    position = position + 2
    read16P = length
End Function

Public Function read8P(ByRef buffer() As Byte, ByRef position As Long) As Byte

    Dim Result As Byte

    Result = buffer(position)
    position = position + 1
    read8P = Result
End Function

Public Function readStringP(ByRef buffer() As Byte, ByRef position As Long) As String

    Dim strBuff() As Byte

    Dim length    As Long

    length = read16P(buffer, position)
   
    ReDim strBuff(length - 1)
    
    CopyMemory strBuff(0), buffer(position), length
    
    position = position + length
    readStringP = ConvertUtf8BytesToString(strBuff)
End Function
