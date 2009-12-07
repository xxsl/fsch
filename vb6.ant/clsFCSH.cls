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

'------
'Logger
'------
Private log As clsLog

'----------------------------------------
'fcsh targets map (targetID, commandline)
'----------------------------------------
Public targets As New Dictionary

'----------------------------
'Flex Compiler SHell starting
'----------------------------
Public isStarting As Boolean

'---------------------
'Last executed command
'---------------------
Private commandLine As String

'-----------
'Preferences
'-----------
Private prefs As clsPreferences

'-------------
'Wrapper state
'-------------
Public isRunning As Boolean

Private isExec_prop As Boolean

'-------------
'EVENTS
'fcsh.exe stop
Public Event onStop(message As DataVO)
'fcsh.exe launch
Public Event onStart(message As DataVO)
'data input
Public Event onComplete(data As DataVO)
'error event
Public Event onError(message As ErrorVO)
'complex event to disable/enable controls
Public Event CommandsEnabled(enable As Boolean)
'-------------

'-------------
'Wrapper class
'-------------
Dim WithEvents fcsh As RedirectLib.Application
Attribute fcsh.VB_VarHelpID = -1

'-----------------------------
'Wrapper console output parser
'-----------------------------
Dim WithEvents parser As clsController
Attribute parser.VB_VarHelpID = -1


'-------------------------
'Setup wrapper, log, prefs
'-------------------------
Public Sub initialize(ByRef appLog As clsLog, ByRef preferences As clsPreferences)
    If (Not isInitialized) Then
        'console output parser
        Set parser = New clsController
        parser.initialize appLog
        
        'wrapper
        Set fcsh = New RedirectLib.Application
        
        'logger
        Set log = appLog

        'preferences
        Set prefs = preferences
        
        'wrapper options
        fcsh.BufferSize = prefs.consoleBuffer
        fcsh.Wait = prefs.consoleWait
        fcsh.Name = prefs.cmd
        
        isExec = False
        isRunning = False
        isStarting = False
        isInitialized = True
        parser.reset
    End If
End Sub


'---------------------------
'Launch Flex Compiler SHell
'---------------------------
Public Sub Start()
    If (Not isRunning) Then
        Select Case fcsh.Start
           Case laAlreadyRunning
                RaiseError FCSH_ALREADY_STATRED, "[FCSH] Flex Compiler SHell is already running"
           Case laWindowsError
                isRunning = False
                RaiseError FCSH_START_FAILURE, "[FCSH] Launch error: " & CStr(fcsh.LastErrorNumber) & ". Make sure you have configured fcsh.exe path."
           Case laOk
                isRunning = True
                isExec = True
                isStarting = True
                log.xInfo "[FCSH] Flex Compiler SHell launched"
        End Select
    Else
        RaiseError FCSH_ALREADY_STATRED, "[FCSH] Flex Compiler SHell is already running"
    End If
End Sub


'------------------------
'Stop Flex Compiler SHell
'------------------------
Public Sub Quit()
    If (isRunning) Then
        fcsh.Stop
    End If
End Sub


'-------------------------------------
'Wrapper process exited (assume error)
'-------------------------------------
Private Sub fcsh_ProcessEnded()
    'log error and console dump
    log.xError "[FCSH] Flex Compiler SHell failed. Check logs/ini."
    log.xError "Output dump:"
    log.output parser.output, "[DUMP]"
    log.xError "Dump end"
    
    'show error tooltip
    DisplayBalloon "Flex Compiler SHell", "Failed to start. Check your server.ini", NIIF_ERROR
    
    mainform.FillTargetsBox
    
    targets.RemoveAll
    
    parser.reset
    
    isExec = False
    isRunning = False
    
    'send onStop event
    Dim message As New DataVO
    message.target = FCSH_STOPPED
    message.data = "empty"
    RaiseEvent onStop(message)
End Sub



'---------------
'Execute command
'---------------
Public Sub exec(command As String)
    If (isRunning) Then
        If (Not isExec) Then
            isExec = True
            commandLine = command
            log.xDebug "fcsh execute: " + command
            
            If (targets.Exists(commandLine)) Then
                'recompile
                log.xInfo "[FCSH] Recompile target: " + command
                log.xFcsh "(fcsh) compile " & targets.Item(command) & vbCrLf
                fcsh.Write ToOEM("compile " & targets.Item(command)) + vbCrLf
            Else
                'compile as new
                log.xInfo "[FCSH] Compile target: " + command
                log.xFcsh "(fcsh) " + command + vbCrLf
                fcsh.Write ToOEM(command) + vbCrLf
            End If
        Else
            RaiseError FCSH_BUSY, "[FCSH] Cant exec: Flex Compiler SHell is busy"
        End If
    Else
        RaiseError FCSH_NOT_STATRED, "[FCSH] Cant exec: Flex Compiler SHell is stopped"
    End If
End Sub


'------------------------------
'Clears targets from fcsh cache
'------------------------------
Public Sub clear(command As String)
    If (isRunning) Then
        If (Not isExec) Then
              commandLine = command
              fcsh.Write ToOEM(command) + vbCrLf
        End If
    End If
End Sub


'-------------------------
'On wrapper console output
'-------------------------
Private Sub fcsh_DataReceived(ByVal sData As String)
    If (isExec) Then
        parser.addStdout (sData)
    End If
End Sub


'----------------------------------
'On build completed, process result
'----------------------------------
Private Sub parser_Result(id As Build, assigned As Long)
    log.xDebug "parser_Result id=" & id & ", assigned=" & assigned
    
    Dim data As New DataVO
    data.data = parser.output
    isExec = False
    
    'clear output buffer
    parser.reset
    
    'detect Flex Compiler SHell startup
    If (isStarting) Then
       isStarting = False
       If (id = BUILD_SUCCESSFULL) Then
            log.xInfo "[FCSH] Flex Compiler SHell started successfully"
            DisplayBalloon "Flex Compiler SHell", "Started successfully", NIIF_INFO
            
            Dim message As New DataVO
            message.target = FCSH_STARTED
            message.data = "empty"
            RaiseEvent onStart(message)
       End If
       Exit Sub
    End If
    
    Dim isAssigned As Boolean
    
    'detect new target
    If (assigned > 0 And Not targets.Exists(commandLine)) Then
        targets.Add commandLine, assigned
        isAssigned = True
        mainform.FillTargetsBox
    End If
    

    Select Case id
        Case Build.BUILD_SUCCESSFULL:
             data.target = FCSH_BUILD_SUCCESSFULL
             If (isAssigned) Then
                DisplayBalloon "Build", "Build successfull. Assigned target id " & assigned & ".", NIIF_INFO
             Else
                DisplayBalloon "Build", "Build successfull. Id is " & targets.Item(commandLine), NIIF_INFO
             End If
        Case Build.BUILD_WARNING:
             data.target = FCSH_BUILD_WARNING
             If (isAssigned) Then
                Dim msg As String
                msg = "Build successfull, but there were warnings." + vbCrLf + "Assigned target id " & assigned & "."
                DisplayBalloon "Build", msg, NIIF_WARNING
             Else
                DisplayBalloon "Build", "Build successfull, but there were warnings. Id is " & targets.Item(commandLine), NIIF_WARNING
             End If
        Case Build.BUILD_ERROR:
             data.target = FCSH_BUILD_ERROR
             If (targets.Exists(commandLine)) Then
                DisplayBalloon "Build", "Build failed. Id is " & targets.Item(commandLine), NIIF_ERROR
             Else
                DisplayBalloon "Build", "Build failed", NIIF_ERROR
             End If
    End Select
    
    'dispatch build completed event
    RaiseEvent onComplete(data)
End Sub



'----------------------------------------------
'Compiler state (Busy/true, Idle/false)
'----------------------------------------------
Public Property Get isExec() As Boolean
    isExec = isExec_prop
End Property

Public Property Let isExec(ByVal vNewValue As Boolean)
    isExec_prop = vNewValue
    RaiseEvent CommandsEnabled((Not isExec_prop) And isRunning)
End Property


'--------------------------------
'Kill wrapper process onTerminate
'--------------------------------
Private Sub Class_Terminate()
    Quit
End Sub


'----------
'Send error
'----------
Private Sub RaiseError(id As FCSHErrors, description As String)
    Dim message As New ErrorVO
    message.description = description
    message.id = id
    RaiseEvent onError(message)
End Sub
