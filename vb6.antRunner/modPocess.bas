Attribute VB_Name = "modPocess"
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
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long


'*******************
'* Win32 Types . . .
'*******************
Public Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessId As Long
   dwThreadId As Long
End Type

Public Type STARTUPINFO
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
Public Enum PROCESS_PRIORITY
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
Public Type PROCESS_TREE
    ProcessId As Long
    ParentProcessId As Long
End Type

Public Function GetProcess() As PROCESS_TREE

    On Error GoTo ERR_GetProcessTree
    
    Dim thisPID As Long
    Call GetWindowThreadProcessId(Runner.hwnd, thisPID)
    

    Dim hSnapShot As Long
    Dim hProcess As Long
    Dim uProcessEntry As PROCESSENTRY32
    Dim lSuccess As Long
    Dim ProcessTree As PROCESS_TREE
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
                If (uProcessEntry.th32ProcessID = thisPID) Then
                    
                    With ProcessTree
                        .ParentProcessId = uProcessEntry.th32ParentProcessID
                        .ProcessId = uProcessEntry.th32ProcessID
                    End With
                    Debug.Print "process found"
                    Exit Do
                Else
                    lCtr = lCtr + 1
                    lSuccess = ProcessNext(hSnapShot, uProcessEntry)
                End If
            Loop
        
        Else
            Err.Raise vbObjectError + 512, , "Unable To Get First Process In Snapshot"
        End If
    
    End If
    
    '********************************
    '* Release handle resources . . .
    '********************************
    CloseHandle (hSnapShot)
    
    GetProcess = ProcessTree
    Exit Function
    
ERR_GetProcessTree:

    CloseHandle (hSnapShot)
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Public Function GetProcessList() As PROCESS_TREE()

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

Public Sub KillProcessTree(ProcessTree() As PROCESS_TREE, ParentProcessId As Long)

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
