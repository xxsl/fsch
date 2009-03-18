Attribute VB_Name = "modSerialization"
Option Explicit

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)

Public Function writeDouble(ByRef buffer() As Byte, ByVal str As Double)
    Dim first As Boolean
    If (UBound(buffer) = 0) Then
        first = True
    End If

    Dim length() As Byte
    length = DoubleToByteArray(str)
    
    Dim oldLen As Long, i As Long
    oldLen = UBound(buffer)
    
    If (first) Then
        ReDim Preserve buffer(UBound(buffer) + 7)
    Else
        ReDim Preserve buffer(UBound(buffer) + 8)
    End If
    
    For i = 0 To 7
        buffer(UBound(buffer) - i) = length(i)
    Next i
End Function

Public Function readDouble(ByRef buffer() As Byte, ByRef position As Long) As Double

    On Error GoTo check
    
    Dim i As Long
    Dim k As Long
    
    Const size As Long = 7
    
    Dim strBuff(0 To size) As Byte
   
    For i = (position + size) To position Step -1
        strBuff(k) = buffer(i)
        k = k + 1
    Next i
    
    Dim result As Double
    CopyMemory result, strBuff(0), (size + 1)

    position = position + (size + 1)
    readDouble = result
check:
    'Debug.Print Err.Description
End Function


Public Function readBoolean(ByRef buffer() As Byte, ByRef position As Long) As Boolean
    Dim result As Byte
    result = buffer(position)
    position = position + 1
    readBoolean = (result = 1)
End Function

Public Function writeBoolean(ByRef buffer() As Byte, ByVal str As Boolean)
    Dim first As Boolean
    If (UBound(buffer) = 0) Then
        Debug.Print "First time"
        first = True
    End If
    
    If (first) Then
        'do nothing
    Else
        ReDim Preserve buffer(UBound(buffer) + 1)
    End If
    
    If (str) Then
        buffer(UBound(buffer)) = 1
    Else
        buffer(UBound(buffer)) = 0
    End If
End Function


Public Function readLong(ByRef buffer() As Byte, ByRef position As Long) As Long
    Dim strBuff() As Byte
    Dim length As Long, i As Long

    ReDim strBuff(3)
    
    For i = (position + 3) To position Step -1
        strBuff(length) = buffer(i)
        length = length + 1
    Next i
    
    CopyMemory length, strBuff(0), 4&

    position = position + 4
    readLong = length
End Function

Public Function writeLong(ByRef buffer() As Byte, ByVal str As Long)
    Dim first As Boolean
    If (UBound(buffer) = 0) Then
        'Debug.Print "First time"
        first = True
    End If

    Dim length() As Byte
    length = LongToByteArray(str)
    
    Dim oldLen As Long, i As Long
    oldLen = UBound(buffer)
    
    If (first) Then
        ReDim Preserve buffer(UBound(buffer) + 3)
    Else
        ReDim Preserve buffer(UBound(buffer) + 4)
    End If
    
    For i = 0 To 3
        buffer(UBound(buffer) - i) = length(i)
    Next i
End Function

Public Function readString(ByRef buffer() As Byte, ByRef position As Long) As String
    Dim strBuff() As Byte
    Dim length As Long, i As Long

    ReDim strBuff(3)
    
    For i = (position + 3) To position Step -1
        strBuff(length) = buffer(i)
        length = length + 1
    Next i
    
    CopyMemory length, strBuff(0), 4&
    'Debug.Print "String length is " & length
    
   
    ReDim strBuff(length - 1)
    
    CopyMemory strBuff(0), buffer(position + 4), length
    
    position = position + 4 + length
    readString = StrConv(UTF16toANSI(strBuff), vbUnicode)
End Function

Public Function writeString(ByRef buffer() As Byte, ByVal str As String)
    Dim first As Boolean
    If (UBound(buffer) = 0) Then
        'Debug.Print "First time"
        first = True
    End If
    
    
    Dim strBuff() As Byte
    strBuff = StrConv(str, vbFromUnicode)
    
    strBuff = ANSItoUTF16(strBuff)
    
    Dim length() As Byte
    'Debug.Print "String length is " & (UBound(strBuff) + 1)
    length = LongToByteArray(UBound(strBuff) + 1)
    
    Dim oldLen As Long, i As Long
    oldLen = UBound(buffer)
    
    ReDim Preserve buffer(UBound(buffer) + 4)
    For i = 0 To 3
        buffer(UBound(buffer) - i) = length(i)
    Next i
       
    oldLen = UBound(buffer)
    
    ReDim Preserve buffer(UBound(buffer) + UBound(strBuff) + 1)

    i = UBound(strBuff)
    CopyMemory buffer(oldLen + 1), strBuff(0), i + 1
    
    oldLen = UBound(buffer)
    
    If (first) Then
        CopyMemory buffer(0), buffer(1), oldLen
        ReDim Preserve buffer(oldLen - 1)
    End If
End Function

Public Function LongToByteArray(ByVal lng As Long) As Byte()
    Dim byteArray(0 To 3) As Byte
    CopyMemory byteArray(0), ByVal VarPtr(lng), 4&
    LongToByteArray = byteArray
End Function

Public Function DoubleToByteArray(ByVal lng As Double) As Byte()
    Dim byteArray(0 To 7) As Byte
    CopyMemory byteArray(0), ByVal VarPtr(lng), 8&
    DoubleToByteArray = byteArray
End Function





