Attribute VB_Name = "modMain"
'***********************************************************************************
'* nimrod97@gmail.com                                                              *
'* Project homepage http://code.google.com/p/fsch/                                 *
'* Adobe Flex Compiler Shell wrapper                                               *
'* 2008                                                                            *
'***********************************************************************************

Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

Public Declare Function PathIsDirectory Lib "shlwapi" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long

Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean

Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'----------------
'window constants
'----------------
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_SHOWWINDOW = &H40

'--------
'controls
'--------
Public Const ICC_USEREX_CLASSES = &H200

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

'------------------------
'reference to System Tray
'------------------------
Public AppTray As frmSysTray

'--------------------
'Try use WinXP styles
'--------------------
Private Function InitCommonControlsVB() As Boolean
   On Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   ' Ensure CC available:
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)
   On Error GoTo 0
End Function

'-----------------------
'Application entry point
'-----------------------
Public Sub Main()
    If (Not InitCommonControlsVB) Then
        Dim log As New clsLog
        log.xError "Critical error: " + Err.description & ":" & Err.Number
    End If
    
    If (Not App.PrevInstance) Then
        Load mainform
    Else
        MsgBox "FCSHServer is already running!", vbCritical, "FCSHServer"
    End If
End Sub


'------------------------
'Set window Always On Top
'------------------------
Public Sub SetAlwaysOnTopMode(hWndOrForm As Variant, Optional ByVal OnTop As Boolean = True)
    Dim hwnd As Long
    ' get the hWnd of the form to be move on top
    If VarType(hWndOrForm) = vbLong Then
        hwnd = hWndOrForm
    Else
        hwnd = hWndOrForm.hwnd
    End If
    SetWindowPos hwnd, IIf(OnTop, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
End Sub


'--------------------------------------
'Check file/folder exists on filesystem
'--------------------------------------
Public Function FileExists(ByVal sPath As String) As Boolean
      If (PathFileExists(sPath)) And Not (PathIsDirectory(sPath)) Then FileExists = True
End Function


'------------------
'Get DOS short path
'------------------
Public Function GetShortName(sFile As String) As String
    Dim sShortFile As String * 256
    Dim lResult As Long

    'Make a call to the GetShortPathName API
    lResult = GetShortPathName(sFile, sShortFile, Len(sShortFile))

    'Trim out unused characters from the string.
    GetShortName = Left$(sShortFile, lResult)
End Function
