VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsController"
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

'-----------
'FCSH Output
'-----------
Public output As String

'-------------------------------------------------------------
'Dispatched when end of build detected (success/warning/error)
'-------------------------------------------------------------
Public Event Result(id As Build, assigned As Long)

Private log As clsLog


'------------------------------------------
'Setup controller, set logger & fcshwrapper
'------------------------------------------
Public Sub initialize(ByRef logger As clsLog)
    Set log = logger
End Sub


'----------------------------
'On console data output event
'----------------------------
Public Sub addStdout(out As String)
    'add to buffer
    output = output + out
    
    'print to log file and window
    log.xFcsh out
    
    'check if compilation finished
    If (isFinished()) Then
        'If error dispatch BUILD_ERROR, id = -1
        If (isError) Then
            log.xFcsh "-------------------[BUILD FAILED]-------------------" + vbCrLf
            RaiseEvent Result(BUILD_ERROR, -1)
        Else
        
            'If warning dispatch BUILD_WARNING, id = assingned
            If (isWarn()) Then
                log.xFcsh "-------------------[BUILD WARNING]-------------------" + vbCrLf
                RaiseEvent Result(BUILD_WARNING, isAssigned())
            Else
                'If no errors/warnings dispatch BUILD_SUCCESSFULL, id = assingned
                log.xFcsh "-------------------[BUILD SUCCESSFULL]-------------------" + vbCrLf
                RaiseEvent Result(BUILD_SUCCESSFULL, isAssigned())
            End If
        End If
       
        log.xInfo "[FCSH] Build completed"
    End If
End Sub


'---------------------------
'Check if target is assigned
'---------------------------
Public Function isAssigned() As Long
    Dim regExp As regExp
    Dim matches As MatchCollection
    Dim match As match
    
    Set regExp = New regExp
    regExp.IgnoreCase = True
    regExp.Global = True
    regExp.Pattern = "fcsh: Assigned ([0-9]+) as the compile target id"
    
    Set matches = regExp.Execute(output)
    
    If (matches.Count > 0) Then
        log.xInfo "[FCSH] Target id assigned " & CLng(matches.Item(0).SubMatches.Item(0))
        isAssigned = CLng(matches.Item(0).SubMatches.Item(0))
    Else
        isAssigned = -1
    End If
End Function


'--------------------------------------
'Check if target compiled with warnings
'--------------------------------------
Public Function isWarn() As Boolean
    Dim regExp As regExp
    Dim matches As MatchCollection
    Dim match As match
    
    Set regExp = New regExp
    regExp.IgnoreCase = True
    regExp.Global = True
    regExp.Pattern = "(.*Warning:[^\r]+)"
    
    Set matches = regExp.Execute(output)
    
    For Each match In matches
        log.xFcshWarn "[FCSH] " & match.value
    Next
    
    isWarn = (matches.Count > 0)
End Function


'---------------------------------
'Check if target failed to compile
'---------------------------------
Public Function isError() As Boolean
    Dim regExp As regExp
    Dim matches As MatchCollection
    Dim match As match
    
    Set regExp = New regExp
    regExp.IgnoreCase = True
    regExp.Global = True
    regExp.Pattern = "(.*Error:[^\r]+)"
    
    Set matches = regExp.Execute(output)
    
    For Each match In matches
        log.xFcshErr "[FCSH] " & match.value
    Next
    
    isError = (matches.Count > 0)
End Function


'-------------------------------------------------
'Check if target compilation finished (any status)
'-------------------------------------------------
Public Function isFinished() As Boolean
    Dim regExp As regExp
    Dim matches As MatchCollection
    Dim match As match
    
    Set regExp = New regExp
    regExp.IgnoreCase = True
    regExp.Global = True
    regExp.Pattern = "\(fcsh\)"
    
    Set matches = regExp.Execute(output)
    
    isFinished = (matches.Count > 0)
End Function


'--------------------
'Clears output buffer
'--------------------
Public Sub reset()
    output = ""
End Sub
