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

Public output As String
Private log As clsLog


Public Event Result(id As Build, assigned As Long)


Public Sub initialize(ByRef logger As clsLog)
    Set log = logger
End Sub


Public Sub addStdout(out As String)
    output = output + out
    'log.xDebug "[FCSH] " + out
    If (isFinished()) Then
        If (isError) Then
            log.xError "[FCSH] Build error"
            RaiseEvent Result(BUILD_ERROR, -1)
        Else
            If (isWarn()) Then
                log.xWarning "[FCSH] Build warnings"
                RaiseEvent Result(BUILD_WARNING, isAssigned())
            Else
                log.xInfo "[FCSH] Build successfull"
                RaiseEvent Result(BUILD_SUCCESSFULL, isAssigned())
            End If
        End If
    End If
End Sub

Private Function isAssigned() As Long
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

Private Function isWarn() As Boolean
    Dim regExp As regExp
    Dim matches As MatchCollection
    Dim match As match
    Set regExp = New regExp
    regExp.IgnoreCase = True
    regExp.Global = True
    regExp.Pattern = "(.*Warning:[^\r]+)"
    Set matches = regExp.Execute(output)
    isWarn = (matches.Count > 0)
End Function


Private Function isError() As Boolean
    Dim regExp As regExp
    Dim matches As MatchCollection
    Dim match As match
    Set regExp = New regExp
    regExp.IgnoreCase = True
    regExp.Global = True
    regExp.Pattern = "(.*Error:[^\r]+)"
    Set matches = regExp.Execute(output)
    isError = (matches.Count > 0)
End Function


Private Function isFinished() As Boolean
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

Public Sub reset()
    output = ""
End Sub
