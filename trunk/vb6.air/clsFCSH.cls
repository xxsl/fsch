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



Public isRunning As Boolean
Public isExec As Boolean


'***************************************************************************************
'fcsh.exe stop
Public Event onStop()
'fcsh.exe launch
Public Event onStart()
'data input
Public Event onDataArrival(data As String)
'error event
Public Event onError(msg As String)
'***************************************************************************************

'console stdOut capture
Dim WithEvents fcsh As RedirectLib.Application
Attribute fcsh.VB_VarHelpID = -1


'set log and params
Public Sub initialize(appLog As clsLog)
    If (Not isInitialized) Then
        Set fcsh = New RedirectLib.Application
        Set log = appLog
        fcsh.BufferSize = 4096
        fcsh.Wait = 500
        
        isRunning = False
        isInitialized = True
    End If
End Sub

'start shell
Public Sub Start(path As String)
    If (Not isRunning) Then
        fcsh.name = path
        Select Case fcsh.Start
           Case laAlreadyRunning
                RaiseEvent onError("fcsh is already running")
           Case laWindowsError
                isRunning = False
                RaiseEvent onError("Windows error: " & CStr(fcsh.LastErrorNumber) & ". Make sure you have configured fcsh.exe path.")
           Case laOk
                isRunning = True
                RaiseEvent onStart
        End Select
    Else
        RaiseEvent onError("fcsh is already running")
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
    isRunning = False
    RaiseEvent onStop
End Sub

'write to fcsh stdIn
Public Sub exec(command As String)
    If (isRunning) Then
        log.xInfo "fcsh execute: " + command
        fcsh.Write ToOEM(command) + vbCrLf
    Else
        RaiseEvent onError("Cant exec: fcsh stopped")
    End If
End Sub

'on fcsh output event handler
Private Sub fcsh_DataReceived(ByVal sData As String)
    Dim output As String
    'output = OemToCharS(sData)
    RaiseEvent onDataArrival(sData)
End Sub

'on terminate stop fcsh.exe
Private Sub Class_Terminate()
    If (isRunning) Then
        fcsh.Stop
    End If
End Sub
