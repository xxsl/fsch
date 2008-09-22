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
    Dim deststring As String  ' ���������� ������
    Dim code As Long
    
    deststring = Space$(Len(sourcestring)) '�������� ���������������� ������
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