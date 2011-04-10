Attribute VB_Name = "modSerialization"
Option Explicit

Public Declare Sub CopyMemory _
               Lib "kernel32" _
               Alias "RtlMoveMemory" (Destination As Any, _
                                      Source As Any, _
                                      ByVal Length As Long)

'---------------------------------
'Serialize double to amf bytearray
'---------------------------------
Public Function writeDouble(ByRef buffer() As Byte, ByVal str As Double)

    Dim first As Boolean

    If (UBound(buffer) = 0) Then
        first = True
    End If

    Dim Length() As Byte

    Length = DoubleToByteArray(str)
    
    Dim oldLen As Long, i As Long

    oldLen = UBound(buffer)
    
    If (first) Then
        ReDim Preserve buffer(UBound(buffer) + 7)
    Else
        ReDim Preserve buffer(UBound(buffer) + 8)
    End If
    
    For i = 0 To 7
        buffer(UBound(buffer) - i) = Length(i)
    Next i

End Function

'-----------------------------------
'Deserialize amf bytearray to Double
'-----------------------------------
Public Function readDouble(ByRef buffer() As Byte, ByRef position As Long) As Double

    On Error GoTo check
    
    Dim i                  As Long

    Dim k                  As Long
    
    Const Size             As Long = 7
    
    Dim strBuff(0 To Size) As Byte
   
    For i = (position + Size) To position Step -1
        strBuff(k) = buffer(i)
        k = k + 1
    Next i
    
    Dim result As Double

    CopyMemory result, strBuff(0), (Size + 1)

    position = position + (Size + 1)
    readDouble = result
check:
    'Debug.Print Err.Description
End Function

'------------------------------------
'Deserialize amf bytearray to Boolean
'------------------------------------
Public Function readBoolean(ByRef buffer() As Byte, ByRef position As Long) As Boolean

    Dim result As Byte

    result = buffer(position)
    position = position + 1
    readBoolean = (result = 1)
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

    Dim Length() As Byte

    Length = LongToByteArray(str)
    
    Dim oldLen As Long, i As Long

    oldLen = UBound(buffer)
    
    If (first) Then
        ReDim Preserve buffer(UBound(buffer) + 3)
    Else
        ReDim Preserve buffer(UBound(buffer) + 4)
    End If
    
    For i = 0 To 3
        buffer(UBound(buffer) - i) = Length(i)
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

    Dim Length() As Byte

    Length = LongToByteArray(str)
    
    Dim oldLen As Long, i As Long

    oldLen = UBound(buffer)
    
    If (first) Then
        ReDim Preserve buffer(UBound(buffer) + 1)
    Else
        ReDim Preserve buffer(UBound(buffer) + 2)
    End If
    
    For i = 0 To 1
        buffer(UBound(buffer) - i) = Length(i)
    Next i

End Function

'-----------------------------------
'Deserialize amf bytearray to String
'-----------------------------------
Public Function readString(ByRef Data As clsSocketData) As String

    Dim strBuff() As Byte

    Dim Length    As Long

    Length = read16X(Data)
   
    If (Length > 0) Then
        readX strBuff, Data, Length
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
    
    Dim Length() As Byte

    Length = LongToByteArray(UBound(strBuff) + 1)
    
    Dim oldLen As Long, i As Long

    oldLen = UBound(buffer)
    
    ReDim Preserve buffer(UBound(buffer) + 4)

    For i = 0 To 3
        buffer(UBound(buffer) - i) = Length(i)
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

Public Function read32X(ByRef Data As clsSocketData) As Long

    Dim buffer() As Byte, Length As Long
    
    readX buffer, Data, 4
    
    Reverse buffer
    
    CopyMemory Length, buffer(0), 4&

    read32X = Length
End Function

Public Function read16X(ByRef Data As clsSocketData) As Integer

    Dim buffer() As Byte, Length As Integer
    
    readX buffer, Data, 2
    
    Reverse buffer
    
    CopyMemory Length, buffer(0), 2&

    read16X = Length
End Function

Public Function read8X(ByRef Data As clsSocketData) As Byte
    Dim buffer() As Byte
    Data.ReadData buffer, 1&
    read8X = buffer(0)
End Function

Public Sub readX(ByRef buffer() As Byte, ByRef Data As clsSocketData, ByVal Size As Long)
     Data.ReadData buffer, Size
End Sub


Private Sub Reverse(ByRef s() As Byte)
  Dim i As Long
  Dim sSwap As Byte
  Dim Length As Long
  
  Length = UBound(s)
  
  For i = 0 To (Length - 1) \ 2
    sSwap = s(Length - i)
    s(Length - i) = s(i)
    s(i) = sSwap
  Next i
End Sub

