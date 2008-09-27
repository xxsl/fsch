Attribute VB_Name = "modApi"
'***********************************************************************************
'* nimrod97@gmail.com                                                              *
'* Project homepage http://code.google.com/p/fsch/                                 *
'* Adobe Flex Compiler Shell wrapper                                               *
'* 2008                                                                            *
'***********************************************************************************


Option Explicit

Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long

Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long

Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long

Private Const STD_OUTPUT_HANDLE = -11&

Private Const STD_INPUT_HANDLE = -10&

Public Declare Function OemToChar Lib "user32" Alias "OemToCharA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long

Public Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)

Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Public Declare Function IsWindowVisible& Lib "user32" (ByVal hWnd As Long)

Public Declare Function GetParent& Lib "user32" (ByVal hWnd As Long)

Public Declare Function CharToOem Lib "user32" Alias "CharToOemA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long

Public Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long

Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long

Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, wParam As Any, lParam As Any) As Long

Public Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

Public Declare Function PathIsDirectory Lib "shlwapi" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long

Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
    
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_SHOWWINDOW = &H40
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000
Public Const LWA_ALPHA = &H2&
    
    
Sub SetAlwaysOnTopMode(hWndOrForm As Variant, Optional ByVal OnTop As Boolean = _
    True)
    Dim hWnd As Long
    ' get the hWnd of the form to be move on top
    If VarType(hWndOrForm) = vbLong Then
        hWnd = hWndOrForm
    Else
        hWnd = hWndOrForm.hWnd
    End If
    SetWindowPos hWnd, IIf(OnTop, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, _
        SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
End Sub
    
Public Function GetShortName(ByVal fFileName As String) As String
  Dim bufTxt As String * 261
  Dim RetLen As Long

  RetLen = GetShortPathName(fFileName, bufTxt, 260)
  GetShortName = Left$(bufTxt, RetLen)
End Function
    


Public Function OemToCharS(sOutput As String)
   Dim outputstr As String
   outputstr = Space$(Len(sOutput))
   OemToChar sOutput, outputstr
   OemToCharS = outputstr
End Function

Public Function ToOEM(sourcestring As String)
    Dim deststring As String  ' получаемая строка
    Dim code As Long
    
    deststring = Space$(Len(sourcestring)) 'получаем перекодированную строку
    code = CharToOem(sourcestring, deststring)
    ToOEM = deststring
End Function



Public Function WriteStdOut(ByVal Text As String) As Long
    Dim StdOut As Long
    Dim result As Long
    Dim BytesWritten As Long
    StdOut = GetStdHandle(STD_OUTPUT_HANDLE)
    result = WriteFile(StdOut, ByVal Text, Len(Text), BytesWritten, ByVal 0&)
    If result = 0 Then
        WriteStdOut = 1001 ', , "Unable to write to standard output"
    ElseIf BytesWritten < Len(Text) Then
        WriteStdOut = 1002 ', , "Incomplete write operation"
    End If
End Function

Public Function FileExists(ByVal sPath As String) As Boolean
      If (PathFileExists(sPath)) And Not (PathIsDirectory(sPath)) Then FileExists = True
End Function
