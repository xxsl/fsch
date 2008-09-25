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
'***********************************************************************************
'* nimrod97@gmail.com                                                              *
'* Project homepage http://code.google.com/p/fsch/                                 *
'* Adobe Flex Compiler Shell wrapper                                               *
'* 2008                                                                            *
'***********************************************************************************

Option Explicit

Private log As clsLog
Private outputParser As New clsParser
Private config As clsConfiguration


Private isInitialized As Boolean



Public isRunning As Boolean
Public isExec As Boolean
Private currentTarget As clsTarget


'***************************************************************************************
'fcsh.exe stop
Public Event onStop()
'fcsh.exe launch
Public Event onStart()
'compile target id assigned
Public Event onIdAssigned(target As clsTarget)
'command finished
Public Event onFinish(target As clsTarget)
'error while processing command
Public Event onError(target As clsTarget)
'***************************************************************************************

'console stdOut capture
Dim WithEvents fcsh As RedirectLib.Application
Attribute fcsh.VB_VarHelpID = -1



'set log and params
Public Sub Initialize(ByRef logger As clsLog, ByRef cfg As clsConfiguration)
    If (Not isInitialized) Then
        Set log = logger
        Set config = cfg
        Set fcsh = New RedirectLib.Application
        fcsh.BufferSize = 8192
        fcsh.Wait = 100
        fcsh.name = config.FCSH_PATH
        
        outputParser.log = logger
                
        isExec = False
        isRunning = False
        isInitialized = True
        log.xInfo "fcsh wrapper is initilized"
    End If
End Sub

'start shell
Public Sub Start()
Dim fakeTarget As New clsTarget

    If (Not isRunning) Then
        Select Case fcsh.Start
           Case laAlreadyRunning
                log.xError "fcsh is already running"
                fakeTarget.fMessage = "fcsh is already running"
                RaiseEvent onError(fakeTarget)
           Case laWindowsError
                fakeTarget.fMessage = "Windows error: " & CStr(fcsh.LastErrorNumber)
                log.xError "Windows error: " & CStr(fcsh.LastErrorNumber)
                isRunning = False
                RaiseEvent onError(fakeTarget)
           Case laOk
                isRunning = True
                outputParser.Begin
                Set currentTarget = Nothing
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
    Set currentTarget = Nothing
    RaiseEvent onStop
End Sub

'write to fcsh stdIn
Public Sub exec(target As clsTarget, Optional incremental As Boolean = True)
    Set currentTarget = target
    
    outputParser.Begin
    
    If (isRunning) Then
        isExec = True
        If (currentTarget.fTargetID = 0) Then
            log.xExec currentTarget.getExecCommand
            fcsh.Write ToOEM(currentTarget.getExecCommand) + vbCrLf
        Else
            If (incremental) Then
                log.xExec currentTarget.getExecRecompile
                fcsh.Write ToOEM(currentTarget.getExecRecompile) + vbCrLf
            Else
                log.xExec currentTarget.getExecCommand
                fcsh.Write ToOEM(currentTarget.getExecCommand) + vbCrLf
            End If
        End If
    Else
        log.xError "Cant exec: fcsh stopped" + vbCrLf
        target.fMessage = "Cant exec: fcsh stopped"
        RaiseEvent onError(target)
    End If
End Sub

'write to fcsh stdIn
Public Sub info(target As clsTarget)
    Set currentTarget = target
    
    outputParser.Begin
    
    If (isRunning) Then
        isExec = True
        If (currentTarget.fTargetID > 0) Then
            log.xExec currentTarget.getExecInfo
            fcsh.Write ToOEM(currentTarget.getExecInfo) + vbCrLf
        End If
    Else
        log.xError "Cant exec: fcsh stopped" + vbCrLf
        target.fMessage = "Cant exec: fcsh stopped"
        RaiseEvent onError(target)
    End If
End Sub


'on fcsh output event handler
Private Sub fcsh_DataReceived(ByVal sData As String)
    Dim output As String
    output = OemToCharS(sData)
    If (outputParser.Add(output)) Then
        log.xDebug "outputParser detected finish"
        isExec = False
        
        If (outputParser.isFcshStart) Then Exit Sub
        
        If (outputParser.isError) Then
            'if error any error happend
            currentTarget.fMessage = outputParser.ErrorMessage
            RaiseEvent onError(currentTarget)
        Else
            'if new compile id was assigned
            If (outputParser.isIdAssigned) Then
                currentTarget.fTargetID = outputParser.TargetID
                RaiseEvent onIdAssigned(currentTarget)
            'if all was succesfull
            Else
                RaiseEvent onFinish(currentTarget)
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
