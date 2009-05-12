Attribute VB_Name = "modFileSystem"
Option Explicit
      ' Enumerations
      Public Enum FolderEnum
          feApp = 0  ' \Program Files\Project (more reliable than App.Path)
          feCommonAppData = 35 ' \Docs & Settings\All Users\Application Data
          feCommonAdminTools = 47 ' \Docs & Settings\All Users\Start Menu\Programs\Administrative Tools
          feCommonDesktop = 25 ' \Docs & Settings\All Users\Desktop
          feCommonDocs = 46 ' \Docs & Settings\All Users\Documents
          feCommonPics = 54 ' \Docs & Settings\All Users\Documents\Pictures
          feCommonMusic = 53 ' \Docs & Settings\All Users\Documents\Music
          feCommonStartMenu = 22 ' \Docs & Settings\All Users\Start Menu
          feCommonStartMenuPrograms = 23 ' \Docs & Settings\All Users\Start Menu\Programs
          feCommonTemplates = 45 ' \Docs & Settings\All Users\Templates
          feCommonVideos = 55 ' \Docs & Settings\All Users\Documents\My Videos
          feLocalAppData = 28 ' \Docs & Settings\User\Local Settings\Application Data
          feLocalCDBurning = 59 ' \Docs & Settings\User\Local Settings\Application Data\Microsoft\CD Burning
          feLocalHistory = 34 ' \Docs & Settings\User\Local Settings\History
          feLocalTempInternetFiles = 32 ' \Docs & Settings\User\Local Settings\Temporary Internet Files
          feProgramFiles = 38 ' \Program Files
          feProgramFilesCommon = 43 ' \Program Files\Common Files
          'feRecycleBin = 10 ' ???
          feUser = 40 ' \Docs & Settings\User
          feUserAdminTools = 48 ' \Docs & Settings\User\Start Menu\Programs\Administrative Tools
          feUserAppData = 26 ' \Docs & Settings\User\Application Data
          feUserCache = 32 ' \Docs & Settings\User\Local Settings\Temporary Internet Files
          feUserCookies = 33 ' \Docs & Settings\User\Cookies
          feUserDesktop = 16 ' \Docs & Settings\User\Desktop
          feUserDocs = 5 ' \Docs & Settings\User\My Documents
          feUserFavorites = 6 ' \Docs & Settings\User\Favorites
          feUserMusic = 13 ' \Docs & Settings\User\My Documents\My Music
          feUserNetHood = 19 ' \Docs & Settings\User\NetHood
          feUserPics = 39 ' \Docs & Settings\User\My Documents\My Pictures
          feUserPrintHood = 27 ' \Docs & Settings\User\PrintHood
          feUserRecent = 8 ' \Docs & Settings\User\Recent
          feUserSendTo = 9 ' \Docs & Settings\User\SendTo
          feUserStartMenu = 11 ' \Docs & Settings\User\Start Menu
          feUserStartMenuPrograms = 2 ' \Docs & Settings\User\Start Menu\Programs
          feUserStartup = 7 ' \Docs & Settings\User\Start Menu\Programs\Startup
          feUserTemplates = 21 ' \Docs & Settings\User\Templates
          feUserVideos = 14  ' \Docs & Settings\User\My Documents\My Videos
          feWindows = 36 ' \Windows
          feWindowFonts = 20 ' \Windows\Fonts
          feWindowsResources = 56 ' \Windows\Resources
          feWindowsSystem = 37 ' \Windows\System32
      End Enum
      Private Declare Function SHGetFolderPath Lib "shfolder" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal pszPath As String) As Long
      Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
      Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
       


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



Public Function SpecialFolder(pfe As FolderEnum) As String
          Const MAX_PATH = 260
          Dim strPath As String
          Dim strBuffer As String
          Dim lngHandle As Long
          Dim lngLen As Long
         
          strBuffer = Space$(MAX_PATH)
          If pfe = feApp Then
              lngHandle = GetModuleHandle(App.EXEName)
              lngLen = GetModuleFileName(lngHandle, strBuffer, MAX_PATH)
              strPath = Left$(strBuffer, lngLen)
              strPath = Left$(strPath, InStrRev(strPath, "\") - 1)
              If InStr(strPath, "Microsoft Visual Studio") > 0 Then strPath = App.path
          Else
              If SHGetFolderPath(0, pfe, 0, 0, strBuffer) = 0 Then strPath = Left$(strBuffer, InStr(strBuffer, vbNullChar) - 1)
          End If
          If Right$(strPath, 1) = "\" Then strPath = Left$(strPath, Len(strPath) - 1)
          SpecialFolder = strPath
End Function
