Attribute VB_Name = "modSerialization"
Option Explicit

Public Declare Sub CopyMemory _
               Lib "kernel32" _
               Alias "RtlMoveMemory" (Destination As Any, _
                                      Source As Any, _
                                      ByVal length As Long)

'---------------------------------
'Serialize double to amf bytearray
'---------------------------------
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

'-----------------------------------
'Deserialize amf bytearray to Double
'-----------------------------------
Public Function readDouble(ByRef buffer() As Byte, ByRef position As Long) As Double

    On Error GoTo check
    
    Dim i                  As Long

    Dim k                  As Long
    
    Const size             As Long = 7
    
    Dim strBuff(0 To size) As Byte
   
    For i = (position + size) To position Step -1
        strBuff(k) = buffer(i)
        k = k + 1
    Next i
    
    Dim Result As Double

    CopyMemory Result, strBuff(0), (size + 1)

    position = position + (size + 1)
    readDouble = Result
check:
    'Debug.Print Err.Description
End Function

'------------------------------------
'Deserialize amf bytearray to Boolean
'------------------------------------
Public Function readBoolean(ByRef buffer() As Byte, ByRef position As Long) As Boolean

    Dim Result As Byte

    Result = buffer(position)
    position = position + 1
    readBoolean = (Result = 1)
End Function

'---------------------------------
'Serialize double to amf bytearray
'---------------------------------
Public Sub writeBoolean(ByRef buffer() As Byte, ByVal str As Boolean)

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

End Sub

'-------------------------------
'Serialize Long to amf bytearray
'-------------------------------
Public Function write32(ByRef buffer() As Byte, ByVal str As Long)

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

'-------------------------------
'Serialize Integer to amf bytearray
'-------------------------------
Public Function write16(ByRef buffer() As Byte, ByVal str As Integer)

    Dim first As Boolean

    If (UBound(buffer) = 0) Then
        first = True
    End If

    Dim length() As Byte

    length = LongToByteArray(str)
    
    Dim oldLen As Long, i As Long

    oldLen = UBound(buffer)
    
    If (first) Then
        ReDim Preserve buffer(UBound(buffer) + 1)
    Else
        ReDim Preserve buffer(UBound(buffer) + 2)
    End If
    
    For i = 0 To 1
        buffer(UBound(buffer) - i) = length(i)
    Next i

End Function

'-----------------------------------
'Deserialize amf bytearray to String
'-----------------------------------
Public Function readString(ByRef data As clsSocketData) As String

    Dim strBuff() As Byte

    Dim length    As Long

    length = read16X(data)
   
    If (length > 0) Then
        strBuff = readX(data, length)
        readString = ConvertUtf8BytesToString(strBuff)
    Else
        readString = ""
    End If
End Function

'----------------------------------
'Serialize dString to amf bytearray
'----------------------------------
Public Function writeString(ByRef buffer() As Byte, ByVal str As String)

    Dim first As Boolean

    If (UBound(buffer) = 0) Then
        first = True
    End If
    
    Dim strBuff() As Byte

    strBuff = ConvertStringToUtf8Bytes(str)
    
    Dim length() As Byte

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

'---------------------------------------------
'Convert Long to bytearray using amf byteorder
'---------------------------------------------
Public Function LongToByteArray(ByVal lng As Long) As Byte()

    Dim byteArray(0 To 3) As Byte

    CopyMemory byteArray(0), ByVal VarPtr(lng), 4&
    LongToByteArray = byteArray
End Function

'---------------------------------------------
'Convert Long to bytearray using amf byteorder
'---------------------------------------------
Public Function DoubleToByteArray(ByVal lng As Double) As Byte()

    Dim byteArray(0 To 7) As Byte

    CopyMemory byteArray(0), ByVal VarPtr(lng), 8&

    DoubleToByteArray = byteArray
End Function

Public Function read32X(ByRef data As clsSocketData) As Long

    Dim buffer() As Byte, length As Long
    
    buffer = reverse(readX(data, 4))
    
    CopyMemory length, buffer(0), 4&

    read32X = length
End Function

Public Function read16X(ByRef data As clsSocketData) As Integer

    Dim buffer() As Byte, length As Integer
    
    buffer = reverse(readX(data, 2))
    
    CopyMemory length, buffer(0), 2&

    read16X = length
End Function

Public Function read8X(ByRef data As clsSocketData) As Byte

    Dim buffer() As Byte

    buffer = readX(data, 1)
    read8X = buffer(0)
End Function

Public Function readX(ByRef data As clsSocketData, ByVal size As Long) As Byte()
    readX = data.ReadData(size)
End Function

Public Function reverse(ByRef bytes() As Byte) As Byte()

    Dim length    As Long

    Dim i         As Long

    Dim k         As Long

    Dim shifted() As Byte
    
    length = UBound(bytes)
    
    ReDim shifted(length)
    
    For i = length To 0 Step -1
        shifted(k) = bytes(i)
        k = k + 1
    Next i

    reverse = shifted
End Function

