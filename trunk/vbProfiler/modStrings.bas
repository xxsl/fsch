Attribute VB_Name = "modString"
Option Explicit

Public Declare Function CharToOem _
               Lib "user32" _
               Alias "CharToOemA" (ByVal lpszSrc As String, _
                                   ByVal lpszDst As String) As Long
      
Public Declare Function OemToChar _
               Lib "user32" _
               Alias "OemToCharA" (ByVal lpszSrc As String, _
                                   ByVal lpszDst As String) As Long

Private objStream As ADODB.Stream



Public Function OemToCharS(sOutput As String)

    Dim outputstr As String

    outputstr = Space$(Len(sOutput))
    OemToChar sOutput, outputstr
    OemToCharS = outputstr
End Function

Public Function ToOEM(sourcestring As String)

    Dim deststring As String

    Dim code       As Long
   
    deststring = Space$(Len(sourcestring))
    code = CharToOem(sourcestring, deststring)
    ToOEM = deststring
End Function

'Public Function ConvertStringToUtf8Bytes(ByRef strText As String) As Byte()
'
'    Dim objStream As ADODB.Stream
'
'    Dim Data()    As Byte
'
'    ' init stream
'    Set objStream = New ADODB.Stream
'    objStream.Charset = "utf-8"
'    objStream.Mode = adModeReadWrite
'    objStream.Type = adTypeText
'    objStream.Open
'
'    ' write bytes into stream
'    objStream.WriteText strText
'    objStream.Flush
'
'    ' rewind stream and read text
'    objStream.position = 0
'    objStream.Type = adTypeBinary
'    objStream.Read 3 ' skip first 3 bytes as this is the utf-8 marker
'    Data = objStream.Read()
'
'    ' close up and return
'    objStream.Close
'    ConvertStringToUtf8Bytes = Data
'
'End Function

Public Sub ConvertUtf8BytesToString(ByRef Data() As Byte, ByRef str As String)
   
    If (objStream Is Nothing) Then
        Set objStream = New ADODB.Stream
        objStream.Charset = "utf-8"
        objStream.Mode = adModeReadWrite
    End If
    
    objStream.Type = adTypeBinary
    objStream.Open
    
    ' write bytes into stream
    objStream.Write Data
    objStream.Flush
    
    ' rewind stream and read text
    objStream.position = 0
    objStream.Type = adTypeText
    str = objStream.ReadText
    
    ' close up and return
    objStream.Close
End Sub
