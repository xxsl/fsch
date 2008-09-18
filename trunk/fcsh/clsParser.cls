VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public isError As Boolean
Public isIdAssigned As Boolean


Private logger As clsLog
Private output() As String



Public Sub Begin()
    logger.xInfo "Start of operation"
    ReDim Preserve output(0)
End Sub

Public Function Add(data As String) As Boolean
    Dim newOutput() As String
    Dim i As Long
    Dim length As Long
    
    newOutput = Split(data, vbLf)
    length = UBound(newOutput)
    ReDim Preserve output(UBound(output) + UBound(newOutput))
    For i = 0 To UBound(newOutput)
        logger.xFcsh newOutput(i)
        output(UBound(output) - length + i) = newOutput(i)
    Next i
    Add = detectEnd
End Function

Private Function detectEnd() As Boolean
    Dim i As Long
    Dim line As String
    For i = 0 To UBound(output)
        line = Trim(LCase(output(i)))
        If ((InStr(1, line, "error") > 0) Or (line = "(fcsh)")) Then
            detectEnd = True
            logger.xInfo "End of operation /" & Trim(output(i)) & "/"
            Exit For
        End If
    Next i
    detectEnd = False
End Function

Public Property Let log(ByRef vNewValue As clsLog)
    Set logger = vNewValue
End Property
