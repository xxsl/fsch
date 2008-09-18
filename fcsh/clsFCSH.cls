VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFCSH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const SHELL As String = "fcsh.exe"

Private log As clsLog
Private isInitialized As Boolean
Private outputParser As New clsParser

Public isRunning As Boolean
Public currentID As Long
Public isExec As Boolean


Public Event onStop()
Public Event onStart()
Public Event onIdAssigned(ByVal id As Long)
Public Event onFinish()
Public Event onError(ByVal Msg As String)

Dim WithEvents fcsh As RedirectLib.Application
Attribute fcsh.VB_VarHelpID = -1



'set log and params
Public Sub Initialize(ByRef logger As clsLog)
    If (Not isInitialized) Then
        Set log = logger
        Set fcsh = New RedirectLib.Application
        fcsh.BufferSize = 8192
        fcsh.Wait = 100
        fcsh.Name = SHELL
        
        outputParser.log = logger
                
        currentID = 0
        isExec = False
        isRunning = False
        isInitialized = True
        log.xInfo "fcsh wrapper is initilized"
    End If
End Sub

'start shell
Public Sub Start()
    If (Not isRunning) Then
        Select Case fcsh.Start
           Case laAlreadyRunning
                log.xError "fcsh is already running"
           Case laWindowsError
                log.xError "Windows error: " & CStr(fcsh.LastErrorNumber)
                isRunning = False
                RaiseEvent onError("Windows error: " & CStr(fcsh.LastErrorNumber) & vbCrLf)
           Case laOk
                isRunning = True
                outputParser.Begin
                RaiseEvent onStart
                log.xInfo "fcsh started"
        End Select
    Else
        log.xError "fcsh is already running"
    End If
End Sub


Public Sub Quit()
    If (isRunning) Then
        fcsh.Stop
        currentID = 0
        isExec = False
    End If
End Sub

'stdIn
Public Sub exec(args As String)
    log.xDebug vbCrLf
    log.xDebug "[Exec]: " + args + vbCrLf
    log.xDebug "-----------------------------------------" + vbCrLf
    'output ">>[  exec  ] @ " & Now & vbCrLf
    If (isRunning) Then
        isExec = True
        fcsh.Write ToOEM(args) + vbCrLf
    Else
        log.xError "Cant exec: fcsh stopped" + vbCrLf
        RaiseEvent onError(args)
    End If
End Sub


'on terminate quit from fcsh
Private Sub Class_Terminate()
    If (isRunning) Then
        fcsh.Stop
    End If
End Sub

'on fcsh output
Private Sub fcsh_DataReceived(ByVal sData As String)
    Dim output As String
    output = OemToCharS(sData)
    outputParser.Add output
    'log.xDebug OemToCharS(sData)
    'parse OemToCharS(sData)
End Sub

Private Sub parse(ByVal data As String)
    If (InStr(1, data, "Error") > 0) Then
        RaiseEvent onError(data)
    End If
    If (InStr(1, data, "Assigned") > 0) Then
        parseID data
    End If
    If (InStr(1, data, "(fcsh) ") > 0) Then
        'output vbCrLf
        'output "<<[finished] @ " & Now & vbCrLf
        'output "-----------------------------------------" + vbCrLf
        isExec = False
        RaiseEvent onFinish
    End If
End Sub


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
          currentID = Val(myMatch.Value)
          RaiseEvent onIdAssigned(currentID)
        Next
End Sub

'on fcsh quit
Private Sub fcsh_ProcessEnded()
    currentID = 0
    isExec = False
    isRunning = False
    RaiseEvent onStop
End Sub

