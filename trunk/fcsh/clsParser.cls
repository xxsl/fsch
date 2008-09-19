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
Public ErrorMessage As String
Public isIdAssigned As Boolean
Public TargetID As Long


Private logger As clsLog
Private output() As String



Public Sub Begin()
    logger.xDebug "Start of operation"
    ErrorMessage = ""
    isError = False
    isIdAssigned = False
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
    Add = detectEnd(newOutput)
End Function

Private Function detectEnd(newOutput() As String) As Boolean
    Dim result As Boolean
    Dim i As Long
    Dim line As String
    For i = 0 To UBound(newOutput)
        line = Trim(LCase(newOutput(i)))
        If (InStr(1, line, "error") > 0) Then
            isError = True
            ErrorMessage = Trim(newOutput(i))
            logger.xDebug "Error: /" & Trim(newOutput(i)) & "/"
        End If
        If (InStr(1, line, "assigned") > 0) Then
            isIdAssigned = True
            parseID (line)
            logger.xDebug "Target ID assigned: /" & Trim(newOutput(i)) & "/"
        End If
        If (line = "(fcsh)") Then
            result = True
            logger.xDebug "End of operation: end /" & Trim(newOutput(i)) & "/"
            Exit For
        End If
    Next i
    detectEnd = result
End Function

Private Sub parseID(data As String)
        Dim myRegExp As RegExp
        Dim myMatches As MatchCollection
        Dim myMatch As Match
        
        Set myRegExp = New RegExp
        myRegExp.IgnoreCase = True
        myRegExp.Global = True
        myRegExp.Pattern = "[0-9]+"
        Set myMatches = myRegExp.Execute(data)
        For Each myMatch In myMatches
          TargetID = Val(myMatch.Value)
        Next
End Sub

Public Property Let log(ByRef vNewValue As clsLog)
    Set logger = vNewValue
End Property
