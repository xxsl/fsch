Attribute VB_Name = "modMain"
Option Explicit

Public Declare Sub CopyMemory _
               Lib "kernel32" _
               Alias "RtlCopyMemory" (Destination As Any, _
                                      Source As Any, _
                                      ByVal Length As Long)

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

Public Declare Function PathIsDirectory Lib "shlwapi" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long

Public Declare Function SetWindowPos _
               Lib "user32" (ByVal hWnd As Long, _
                             ByVal hWndInsertAfter As Long, _
                             ByVal X As Long, _
                             ByVal Y As Long, _
                             ByVal cX As Long, _
                             ByVal cY As Long, _
                             ByVal wFlags As Long) As Long

Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean

Public Declare Function GetShortPathName _
               Lib "kernel32" _
               Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
                                          ByVal lpszShortPath As String, _
                                          ByVal cchBuffer As Long) As Long

Public Declare Function SendMessage _
               Lib "user32" _
               Alias "SendMessageA" (ByVal hWnd As Long, _
                                     ByVal wMsg As Long, _
                                     ByVal wParam As Long, _
                                     lParam As Any) As Long

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

    ReDim lngBuffer(3)
    
    ReDim intBuffer(1)
    
    ReDim bytBuffer(0)

    On Error GoTo 0

End Function

'-----------------------
'Application entry point
'-----------------------
Public Sub Main()

    If (Not InitCommonControlsVB) Then
        'Dim Log As New clsLog
        'Log.xError "Critical error: " + Err.Description & ":" & Err.Number
    End If
    
    If (Not App.PrevInstance) Then
        Load frmMain
        frmMain.Show
    Else
        MsgBox "Profiler is already running!", vbCritical, "Profiler"
    End If

End Sub

'------------------------
'Set window Always On Top
'------------------------
Public Sub SetAlwaysOnTopMode(hWndOrForm As Variant, Optional ByVal OnTop As Boolean = True)

    Dim hWnd As Long

    ' get the hWnd of the form to be move on top
    If VarType(hWndOrForm) = vbLong Then
        hWnd = hWndOrForm
    Else
        hWnd = hWndOrForm.hWnd
    End If

    SetWindowPos hWnd, IIf(OnTop, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
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

    Dim lResult    As Long

    'Make a call to the GetShortPathName API
    lResult = GetShortPathName(sFile, sShortFile, Len(sShortFile))

    'Trim out unused characters from the string.
    GetShortName = Left$(sShortFile, lResult)
End Function

Public Function getPath() As String

    If (Len(App.path) = 3) Then
        getPath = App.path
    Else
        getPath = App.path + "\"
    End If

End Function
