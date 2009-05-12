Attribute VB_Name = "modFileSystem"
Option Explicit

Public Const WAIT_FAILED = &HFFFFFFFF 'Our WaitForSingleObject failed to wait and returned -1
Public Const WAIT_OBJECT_0 = &H0& 'The waitable object got signaled
Public Const WAIT_ABANDONED = &H80& 'We got out of the waitable object
Public Const WAIT_TIMEOUT = &H102& 'the interval we used, timed out.
Public Const STANDARD_RIGHTS_ALL = &H1F0000 'No special user rights needed to open this process

Public Const FILE_NOTIFY_CHANGE_ATTRIBUTES = &H4
Public Const FILE_NOTIFY_CHANGE_DIR_NAME = &H2
Public Const FILE_NOTIFY_CHANGE_FILE_NAME = &H1
Public Const FILE_NOTIFY_CHANGE_SIZE = &H8
Public Const FILE_NOTIFY_CHANGE_LAST_WRITE = &H10
Public Const FILE_NOTIFY_CHANGE_SECURITY = &H100
Public Const FILE_NOTIFY_CHANGE_ALL = &H4 Or &H2 Or &H1 Or &H8 Or &H10 Or &H100

Private Declare Function FindFirstChangeNotification Lib "kernel32" Alias "FindFirstChangeNotificationA" (ByVal lpPathName As String, ByVal bWatchSubtree As Long, ByVal dwNotifyFilter As Long) As Long
Private Declare Function FindCloseChangeNotification Lib "kernel32" (ByVal hChangeHandle As Long) As Long
Private Declare Function FindNextChangeNotification Lib "kernel32" (ByVal hChangeHandle As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function ResetEvent Lib "kernel32" (ByVal hEvent As Long) As Long

Public Function WaitForFileChange(ByVal fn As String, Optional ByVal flags = FILE_NOTIFY_CHANGE_ALL, Optional ByVal CheckSubDir As Boolean = False, Optional ByVal TimeOut As Long = -1) As Long
    'waits until a change occurs in directory fn
    Dim hNote As Long
    Dim result As Long
   
    hNote = FindFirstChangeNotification(fn, CheckSubDir, flags) 'Set the hook
    
    result = WaitForSingleObject(hNote, TimeOut)  'Wait for the event
    
    FindCloseChangeNotification hNote 'Clear the hook
    
    WaitForFileChange = result
End Function
