VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CShell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*****************************
'* Win32 Function Stubs . . .
'*****************************
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetPriorityClass Lib "kernel32" (ByVal hProcess As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForInputIdle Lib "user32" (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long


'*******************
'* Win32 Types . . .
'*******************
Private Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessId As Long
   dwThreadId As Long
End Type

Private Type STARTUPINFO
   cb As Long
   lpReserved As String
   lpDesktop As String
   lpTitle As String
   dwX As Long
   dwY As Long
   dwXSize As Long
   dwYSize As Long
   dwXCountChars As Long
   dwYCountChars As Long
   dwFillAttribute As Long
   dwFlags As Long
   wShowWindow As Integer
   cbReserved2 As Integer
   lpReserved2 As Long
   hStdInput As Long
   hStdOutput As Long
   hStdError As Long
End Type

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * 6400
End Type

'***********************
'* Win32 Constants . . .
'***********************
Private Const INFINITE As Long = &HFFFF
Private Const TH32CS_SNAPPROCESS As Long = 2&
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const WINAPI_TRUE = 1
Private Const PROCESS_TERMINATE = 1
Private Const CREATE_SUSPENDED As Long = &H4

'************************************
'* Applications priority class . . .
'************************************
Public Enum PROCESS_PRIORITY1
    ABOVE_NORMAL_PRIORITY_CLASS = &H8000
    BELOW_NORMAL_PRIORITY_CLASS = &H4000
    HIGH_PRIORITY_CLASS = &H80
    IDLE_PRIORITY_CLASS = &H40
    NORMAL_PRIORITY_CLASS = &H20
    REALTIME_PRIORITY_CLASS = &H100
End Enum

'******************************
'* Object State variables . . .
'******************************

Private mApplication As String
Private mArguments As String
Private mProcessInformation As PROCESS_INFORMATION
Private mSuspended As Boolean

Public Sub Initialise(ePriority As PROCESS_PRIORITY1, Optional Suspended As Boolean = False, Optional Synchronous As Boolean = False)

    On Error GoTo ERR_Initialise
    
    Dim lRet As Long
    Dim lCreationFlags As Long
    Dim uStartUpInfo As STARTUPINFO
    Dim uProcessInformation As PROCESS_INFORMATION
     
    '*******************************************************
    '* Check to see if Application has been specified . . .
    '*******************************************************
    If LenB(mApplication) = 0 Then
        Err.Raise vbObjectError + 512, , "Application Not Specified"
    End If
    
    '***************************************************
    '* Check to see if Application actually exists . . .
    '***************************************************
    If LenB(Dir(mApplication)) = 0 Then
        Err.Raise vbObjectError + 512, , "Application '" & mApplication & "' Does Not Exist"
    End If

    '****************************************
    '* Deal with suspended applications . . .
    '****************************************
    If Suspended Then
        lCreationFlags = CREATE_SUSPENDED Or ePriority
        mSuspended = True
    Else
        lCreationFlags = ePriority
        mSuspended = False
    End If
    
    '**************************
    '* Start the Process . . .
    '**************************
    uStartUpInfo.cb = Len(uStartUpInfo)
    lRet = CreateProcessA(vbNullString, mApplication & " " & mArguments, ByVal 0&, ByVal 0&, 1&, lCreationFlags, ByVal 0&, vbNullString, uStartUpInfo, uProcessInformation)
    
    If lRet <> 0 Then
    
        If Synchronous Then
            
            '*****************************************
            '* Wait until application terminates . . .
            '*****************************************
            Call WaitForSingleObject(uProcessInformation.hProcess, INFINITE)
            
            '******************************
            '* Release stored handles . . .
            '******************************
            Call CloseHandle(mProcessInformation.hThread)
            Call CloseHandle(mProcessInformation.hProcess)
            
        Else
        
            '************************************
            '* Remember the process details . . .
            '************************************
            mProcessInformation = uProcessInformation
            
        End If
    
    Else
        Err.Raise vbObjectError + 512, , "Create Process Failed"
    End If
    
    Exit Sub
    
ERR_Initialise:
    Err.Raise Err.Number, "ProcessManager.CShell.Initialise", "Unable to Start Application '" & mApplication & " " & mArguments & _
    "' Because '" & Err.Description & "'"
End Sub


Public Sub Terminate()

    On Error GoTo ERR_Terminate

    Dim ProcessTree() As PROCESS_TREE
    
    If Not mProcessInformation.dwProcessId = 0 Then
    
        '*************************
        '* Kill off the kids . . .
        '*************************
        ProcessTree = GetProcessList()
        KillProcessTree ProcessTree, mProcessInformation.dwProcessId
        
        '***********************************
        '* Kill initial (root) process . . .
        '***********************************
        KillProcess mProcessInformation.dwProcessId
    
    End If
    
    Exit Sub
    
ERR_Terminate:
    Err.Raise Err.Number, "ProcessManager.CShell.Terminate", "Unable To Terminate Application '" & mApplication & "' Because '" & Err.Description & "'"
End Sub


Public Function IsRunning() As Boolean

    On Error GoTo ERR_IsRunning
    
    '***********************************************************************
    '* If we cannot get the apps priority class, the app cannot exist . . .
    '***********************************************************************
    If GetPriorityClass(mProcessInformation.hProcess) > 0 Then
         IsRunning = True
    Else
         IsRunning = False
    End If

    Exit Function
    
ERR_IsRunning:
    Err.Raise vbObjectError + 512, "ProcessManager.CShell.IsRunning", "Unable to Determine whether '" & mApplication & "' is running because '" & Err.Description & "'"
End Function

Public Sub Release()

    On Error GoTo ERR_Release
    
    '******************************************************
    '* Ensure that the thread is currently suspended . . .
    '******************************************************
    If mSuspended Then
        Call ResumeThread(mProcessInformation.hThread)
        mSuspended = False
    End If
    
    Exit Sub
    
ERR_Release:
    Err.Raise vbObjectError + 512, "ProcessManager.CShell.Release", "Unable To Release Thread Suspension Because '" & Err.Description & "'"
End Sub

Private Sub KillProcessTree(ProcessTree() As PROCESS_TREE, ParentProcessId As Long)

    Dim lCtr As Long
    
    '*********************************************
    '* Check every process for it's children . . .
    '*********************************************
    For lCtr = 0 To UBound(ProcessTree)
        If ProcessTree(lCtr).ParentProcessId = ParentProcessId Then
            KillProcessTree ProcessTree, ProcessTree(lCtr).ProcessId
            KillProcess ProcessTree(lCtr).ProcessId
        End If
    Next
        
End Sub

Private Function GetProcessList() As PROCESS_TREE()

    On Error GoTo ERR_GetProcessTree

    Dim hSnapShot As Long
    Dim hProcess As Long
    Dim uProcessEntry As PROCESSENTRY32
    Dim lSuccess As Long
    Dim ProcessTree() As PROCESS_TREE
    Dim lCtr As Long
    
    '************************************************************
    '* Get a snapshot of all of the processes in the system . . .
    '************************************************************
    hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    
    '***********************************************
    '* If we don't have a snapshot then finish . . .
    '***********************************************
    If hSnapShot = INVALID_HANDLE_VALUE Then
        Err.Raise vbObjectError + 512, , "Unable To Get Process Snapshot"
    Else
    
        '*********************************
        '* Get first process in list . . .
        '*********************************
        uProcessEntry.dwSize = Len(uProcessEntry)
        lSuccess = ProcessFirst(hSnapShot, uProcessEntry)
        
        If lSuccess = WINAPI_TRUE Then
        
            lCtr = 0
            
            '**********************************
            '* Loop through all processes . . .
            '**********************************
            Do Until lSuccess <> WINAPI_TRUE
            
                ReDim Preserve ProcessTree(lCtr)
                With ProcessTree(lCtr)
                    .ParentProcessId = uProcessEntry.th32ParentProcessID
                    .ProcessId = uProcessEntry.th32ProcessID
                End With
                
                lCtr = lCtr + 1
                lSuccess = ProcessNext(hSnapShot, uProcessEntry)
                
            Loop
        
        Else
            Err.Raise vbObjectError + 512, , "Unable To Get First Process In Snapshot"
        End If
    
    End If
    
    '********************************
    '* Release handle resources . . .
    '********************************
    CloseHandle (hSnapShot)
    
    GetProcessList = ProcessTree
    Exit Function
    
ERR_GetProcessTree:

    CloseHandle (hSnapShot)
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Private Sub KillProcess(ProcessId As Long)

    On Error GoTo ERR_KillProcess

    Dim hProcess As Long
    Dim lExitCode As Long

    '*************************************************
    '* Kill the process, and release the handle . . .
    '*************************************************
    hProcess = OpenProcess(PROCESS_TERMINATE, False, ProcessId)
    Call TerminateProcess(hProcess, lExitCode)
    Call CloseHandle(hProcess)
    
    Exit Sub
    
ERR_KillProcess:
    Call CloseHandle(hProcess)
End Sub

Property Let Application(sApplication As String)
    mApplication = sApplication
End Property


Property Get Application() As String
    Application = mApplication
End Property


Property Let Arguments(sArguments As String)
    mArguments = sArguments
End Property


Property Get Arguments() As String
    Arguments = mArguments
End Property
