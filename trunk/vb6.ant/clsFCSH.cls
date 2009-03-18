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

Private isInitialized As Boolean
Private log As clsLog
Private targets As New Dictionary
Private commandLine As String


Public isRunning As Boolean
Public isExec As Boolean

Private output As String

'***************************************************************************************
'fcsh.exe stop
Public Event onStop(message As DataVO)
'fcsh.exe launch
Public Event onStart(message As DataVO)
'data input
Public Event onComplete(data As DataVO)
'error event
Public Event onError(message As ErrorVO)
'***************************************************************************************

'console stdOut capture
Dim WithEvents fcsh As RedirectLib.Application
Attribute fcsh.VB_VarHelpID = -1

'parsing
Dim WithEvents parser As clsController
Attribute parser.VB_VarHelpID = -1



Public Sub initialize(ByRef appLog As clsLog, ByRef prefs As clsPreferences)
    If (Not isInitialized) Then
        Set parser = New clsController
        parser.initialize appLog
        Set fcsh = New RedirectLib.Application
        Set log = appLog
        fcsh.BufferSize = prefs.CONSOLE_BUFFER
        fcsh.Wait = prefs.CONSOLE_WAIT
        fcsh.name = prefs.CMD
        
        isExec = False
        isRunning = False
        isInitialized = True
        parser.reset
    End If
End Sub


Public Sub Start(path As String)
    If (Not isRunning) Then
        'fcsh.name = path
        Select Case fcsh.Start
           Case laAlreadyRunning
                RaiseError FCSH_ALREADY_STATRED, "fcsh is already running"
           Case laWindowsError
                isRunning = False
                RaiseError FCSH_START_FAILURE, "Windows error: " & CStr(fcsh.LastErrorNumber) & ". Make sure you have configured fcsh.exe path."
           Case laOk
                isRunning = True
                Dim message As New DataVO
                message.target = FCSH_STARTED
                message.data = "empty"
                RaiseEvent onStart(message)
        End Select
    Else
        RaiseError FCSH_ALREADY_STATRED, "fcsh is already running"
    End If
End Sub


Public Sub Quit()
    If (isRunning) Then
        fcsh.Stop
    End If
End Sub


Private Sub fcsh_ProcessEnded()
    targets.RemoveAll
    parser.reset
    isExec = False
    isRunning = False
    Dim message As New DataVO
    message.target = FCSH_STOPPED
    message.data = "empty"
    RaiseEvent onStop(message)
End Sub


Public Sub exec(command As String)
    If (isRunning) Then
        If (Not isExec) Then
            isExec = True
            commandLine = command
            log.xInfo "fcsh execute: " + command
            If (targets.Exists(commandLine)) Then
                log.xInfo "Recompile target: " + command
                fcsh.Write ToOEM("compile " & targets.Item(command)) + vbCrLf
            Else
                log.xInfo "Compile target: " + command
                fcsh.Write ToOEM(command) + vbCrLf
            End If
        Else
            RaiseError FCSH_BUSY, "Cant exec: fcsh is busy"
        End If
    Else
        RaiseError FCSH_NOT_STATRED, "Cant exec: fcsh stopped"
    End If
End Sub


Private Sub fcsh_DataReceived(ByVal sData As String)
    If (isExec) Then
        parser.addStdout sData
    End If
End Sub


Private Sub RaiseError(id As FCSHErrors, description As String)
    Dim message As New ErrorVO
    message.description = description
    message.id = id
    RaiseEvent onError(message)
End Sub


Private Sub Class_Terminate()
    Quit
End Sub

Private Sub parser_Result(id As Build, assigned As Long)
    Dim isAssigned As Boolean
    If (assigned > 0 And Not targets.Exists(commandLine)) Then
        targets.Add commandLine, assigned
        isAssigned = True
    End If
    
    Dim data As New DataVO
    data.data = parser.output
    Select Case id
    Case Build.BUILD_SUCCESSFULL:
         data.target = FCSH_BUILD_SUCCESSFULL
         If (isAssigned) Then
            DisplayBalloon "Build", "Build successfull. Assigned target id " & assigned & ".", NIIF_INFO
         Else
            DisplayBalloon "Build", "Build successfull", NIIF_INFO
         End If
    Case Build.BUILD_WARNING:
         data.target = FCSH_BUILD_WARNING
         If (isAssigned) Then
            Dim msg As String
            msg = "Build successfull, but there were warnings." + vbCrLf + "Assigned target id " & assigned & "."
            DisplayBalloon "Build", msg, NIIF_WARNING
         Else
            DisplayBalloon "Build", "Build successfull, but there were warnings.", NIIF_WARNING
         End If
    Case Build.BUILD_ERROR:
         data.target = FCSH_BUILD_ERROR
         DisplayBalloon "Build", "Build failed", NIIF_ERROR
    End Select
    
    isExec = False
    parser.reset
    RaiseEvent onComplete(data)
End Sub
