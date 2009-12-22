Attribute VB_Name = "modUtil"
Option Explicit

Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal HWND As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

Public Declare Function PathIsDirectory Lib "shlwapi" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long

Public Declare Function SetWindowPos Lib "user32" (ByVal HWND As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long

'----------------
'window constants
'----------------
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_SHOWWINDOW = &H40

'------------------------
'Returns application path
'------------------------
Public Function GetPath() As String
    If (Len(App.Path) > 3) Then
        GetPath = App.Path & "\" & App.EXEName
    Else
        GetPath = App.Path & App.EXEName
    End If
End Function

'------------------------
'Set window Always On Top
'------------------------
Public Sub SetAlwaysOnTopMode(hWndOrForm As Variant, Optional ByVal OnTop As Boolean = True)
    Dim HWND As Long
    ' get the hWnd of the form to be move on top
    If VarType(hWndOrForm) = vbLong Then
        HWND = hWndOrForm
    Else
        HWND = hWndOrForm.HWND
    End If
    SetWindowPos HWND, IIf(OnTop, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
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
