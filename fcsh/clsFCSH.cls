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

Private log As clsLog

Const SHELL As String = "fcsh.exe"
Private isInitialized As Boolean

Private outputParser As New clsParser

Public isRunning As Boolean
Public isExec As Boolean

'***************************************************************************************
'fcsh.exe stop
Public Event onStop()
'scsh.exe launch
Public Event onStart()
'compile target id assigned
Public Event onIdAssigned(ByVal id As Long)
'command finished
Public Event onFinish()
'error while processing command
Public Event onError(ByVal Msg As String)
'***************************************************************************************

'console stdOut capture
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
                 RaiseEvent onError("fcsh is already running")
           Case laWindowsError
                log.xError "Windows error: " & CStr(fcsh.LastErrorNumber)
                isRunning = False
                RaiseEvent onError("Windows error: " & CStr(fcsh.LastErrorNumber) & vbCrLf)
           Case laOk
                isRunning = True
                outputParser.Begin
                RaiseEvent onStart
                log.xDebug "fcsh started"
        End Select
    Else
        log.xError "fcsh is already running"
    End If
End Sub

'stop fcsh.exe
Public Sub Quit()
    If (isRunning) Then
        fcsh.Stop
    End If
End Sub

'on stop fcsh.exe event handler
Private Sub fcsh_ProcessEnded()
    isExec = False
    isRunning = False
    log.xDebug "fcsh stopped"
    RaiseEvent onStop
End Sub

'write to fcsh stdIn
Public Sub exec(args As String)
    log.xExec args

    If (isRunning) Then
        isExec = True
        fcsh.Write ToOEM(args) + vbCrLf
    Else
        log.xError "Cant exec: fcsh stopped" + vbCrLf
        RaiseEvent onError("Cant exec: fcsh stopped" + vbCrLf)
    End If
End Sub


'on fcsh output event handler
Private Sub fcsh_DataReceived(ByVal sData As String)
    Dim output As String
    output = OemToCharS(sData)
    If (outputParser.Add(output)) Then
        log.xDebug "outputParser detected finish"
        isExec = False
        If (outputParser.isError) Then
            'if error any error happend
            RaiseEvent onError(outputParser.ErrorMessage)
        Else
            'if new compile id was assigned
            If (outputParser.isIdAssigned) Then
                RaiseEvent onIdAssigned(outputParser.TargetID)
            'if all was succesfull
            Else
                RaiseEvent onFinish
            End If
        End If
    End If
End Sub

'on terminate stop fcsh.exe
Private Sub Class_Terminate()
    If (isRunning) Then
        fcsh.Stop
    End If
End Sub
