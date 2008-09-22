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

Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Public Declare Function IsWindowVisible& Lib "user32" (ByVal hwnd As Long)

Public Declare Function GetParent& Lib "user32" (ByVal hwnd As Long)

Public Declare Function CharToOem Lib "user32" Alias "CharToOemA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long

Public Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, wParam As Any, lParam As Any) As Long

Public Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

Public Declare Function PathIsDirectory Lib "shlwapi" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long

Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
    
Dim sPattern As String, hFind As Long




Public Function GetShortName(sFile As String) As String
    Dim sShortFile As String * 67
    Dim lResult As Long

    'Make a call to the GetShortPathName API
    lResult = GetShortPathName(sFile, sShortFile, _
    Len(sShortFile))

    'Trim out unused characters from the string.
    GetShortName = Left$(sShortFile, lResult)

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


Function EnumWinProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
  Dim k As Long, sName As String
  hFind = 0
  'IsWindowVisible(hwnd) And
  If GetParent(hwnd) = 0 Then
     sName = Space$(128)
     k = GetWindowText(hwnd, sName, 128)
     If k > 0 Then
        sName = Left$(sName, k)
        If lParam = 0 Then sName = UCase(sName)
        If sName Like sPattern Then
           hFind = hwnd
           EnumWinProc = 0
           Exit Function
        End If
     End If
  End If
  EnumWinProc = 1
End Function

Public Function FindWindowWild(sWild As String, Optional bMatchCase As Boolean = True) As Long
  sPattern = sWild
  If Not bMatchCase Then sPattern = UCase(sPattern)
  EnumWindows AddressOf EnumWinProc, bMatchCase
  FindWindowWild = hFind
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
