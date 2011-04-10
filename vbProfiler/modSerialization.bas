Attribute VB_Name = "modSerialization"
Option Explicit

Public Declare Sub CopyMemory _
               Lib "kernel32" _
               Alias "RtlMoveMemory" (Destination As Any, _
                                      Source As Any, _
                                      ByVal Length As Long)
                                      
Public Declare Function ntohl Lib "wsock32.dll" (ByVal hostlong As Long) As Long

Public Declare Function ntohs Lib "wsock32.dll" (ByVal hostlong As Integer) As Integer

Private lngBuffer(0 To 3) As Byte

Private intBuffer(0 To 1) As Byte

Private bytBuffer(0 To 0) As Byte

Public Sub readString(ByRef Data As clsSocketData, ByRef result As String)

    Dim strBuff() As Byte

    Dim Length    As Integer

    read16X Data, Length
   
    If (Length > 0) Then
        readX strBuff, Data, Length
        ConvertUtf8BytesToString strBuff, result
    Else
        result = ""
    End If

End Sub

Public Sub read32X(ByRef Data As clsSocketData, ByRef result As Long)
    
    Data.ReadData32 lngBuffer
    
    CopyMemory result, lngBuffer(0), 4&
    
    result = ntohl(result)
End Sub

Public Sub read16X(ByRef Data As clsSocketData, ByRef result As Integer)
    
    Data.ReadData16 intBuffer
    
    CopyMemory result, intBuffer(0), 2&
    
    result = ntohs(result)
End Sub

Public Sub read8X(ByRef Data As clsSocketData, ByRef result As Byte)

    Data.ReadData8 bytBuffer
    
    result = bytBuffer(0)
    
End Sub

Public Sub readX(ByRef Buffer() As Byte, ByRef Data As clsSocketData, ByVal Size As Long)
    Data.ReadData Buffer, Size
End Sub
