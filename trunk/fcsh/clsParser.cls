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
'***********************************************************************************
'* nimrod97@gmail.com                                                              *
'* Project homepage http://code.google.com/p/fsch/                                 *
'* Adobe Flex Compiler Shell wrapper                                               *
'* 2008                                                                            *
'***********************************************************************************

Option Explicit

Public isWarning As Boolean
Public isError As Boolean
Public ErrorMessage As String
Public isIdAssigned As Boolean
Public TargetID As Long
Public isFcshStart As Boolean


Private logger As clsLog
Private output() As String



Public Sub Begin()
    logger.xDebug "Start of operation"
    ErrorMessage = ""
    isFcshStart = False
    isError = False
    isWarning = False
    isIdAssigned = False
    TargetID = 0
    ReDim Preserve output(0)
End Sub

Public Function Add(Data As String) As Boolean
    Dim newOutput() As String
    Dim i As Long
    Dim Length As Long
    
    newOutput = Split(Data, vbLf)
    Length = UBound(newOutput)
    ReDim Preserve output(UBound(output) + UBound(newOutput))
    
    Add = detectEnd(newOutput)
    
    For i = 0 To UBound(newOutput) - 1
        logger.xFcsh newOutput(i)
        output(UBound(output) - Length + i) = newOutput(i)
    Next i
    
    If (Not Add) Then
        output(UBound(output) - Length + UBound(newOutput)) = newOutput(UBound(newOutput))
    End If
End Function

Private Function detectEnd(newOutput() As String) As Boolean
    Dim result As Boolean
    Dim i As Long
    Dim line As String
    For i = 0 To UBound(newOutput)
        line = Trim(LCase(newOutput(i)))
        If (InStr(1, line, "warning:") > 0) Then
            isWarning = True
            ErrorMessage = Trim(newOutput(i))
            logger.xDebug "Warn: /" & Trim(newOutput(i)) & "/"
        End If
        If (InStr(1, line, "error:") > 0) Then
            isError = True
            ErrorMessage = Trim(newOutput(i))
            logger.xDebug "Error: /" & Trim(newOutput(i)) & "/"
        End If
        If (InStr(1, line, "assigned") > 0) Then
            isIdAssigned = True
            parseID (line)
            logger.xDebug "Target ID assigned: /" & Trim(newOutput(i)) & "/"
        End If
        If (InStr(1, line, LCase("Adobe Flex Compiler SHell")) > 0) Then
            isFcshStart = True
        End If
        If (line = "(fcsh)") Then
            result = True
            logger.xDebug "End of operation: end /" & Trim(newOutput(i)) & "/"
            Exit For
        End If
    Next i
    detectEnd = result
End Function

Private Sub parseID(Data As String)
        Dim myRegExp As RegExp
        Dim myMatches As MatchCollection
        Dim myMatch As Match
        
        Set myRegExp = New RegExp
        myRegExp.IgnoreCase = True
        myRegExp.Global = True
        myRegExp.Pattern = "[0-9]+"
        Set myMatches = myRegExp.Execute(Data)
        For Each myMatch In myMatches
          TargetID = Val(myMatch.Value)
        Next
End Sub

Public Property Let log(ByRef vNewValue As clsLog)
    Set logger = vNewValue
End Property
