Attribute VB_Name = "modSerializationPeak"
Option Explicit

Public Function read32P(ByRef Buffer() As Byte, ByRef position As Long) As Long

    Dim strBuff() As Byte

    Dim Length    As Long, i As Long

    ReDim strBuff(3)
    
    For i = (position + 3) To position Step -1
        strBuff(Length) = Buffer(i)
        Length = Length + 1
    Next i
    
    CopyMemory Length, strBuff(0), 4&

    position = position + 4
    read32P = Length
End Function

Public Function read16P(ByRef Buffer() As Byte, ByRef position As Long) As Integer

    Dim strBuff() As Byte

    Dim Length    As Long, i As Long

    ReDim strBuff(1)
    
    For i = (position + 1) To position Step -1
        strBuff(Length) = Buffer(i)
        Length = Length + 1
    Next i
    
    CopyMemory Length, strBuff(0), 2&

    position = position + 2
    read16P = Length
End Function

Public Function read8P(ByRef Buffer() As Byte, ByRef position As Long) As Byte

    Dim result As Byte

    result = Buffer(position)
    position = position + 1
    read8P = result
End Function

Public Function readStringP(ByRef Buffer() As Byte, ByRef position As Long) As String

    Dim strBuff() As Byte
    Dim str As String
    Dim Length    As Long

    Length = read16P(Buffer, position)
   
    ReDim strBuff(Length - 1)
    
    CopyMemory strBuff(0), Buffer(position), Length
    
    position = position + Length
    ConvertUtf8BytesToString strBuff, str
    readStringP = str
End Function
