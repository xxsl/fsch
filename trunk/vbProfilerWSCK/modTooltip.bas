Attribute VB_Name = "modTooltip"
'***********************************************************************************
'* nimrod97@gmail.com                                                              *
'* Project homepage http://code.google.com/p/fsch/                                 *
'* Adobe Flex Compiler Shell wrapper                                               *
'* 2008                                                                            *
'***********************************************************************************

Option Explicit


Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal HWND As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal HWND As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function GetClientRect Lib "user32" (ByVal HWND As Long, lpRect As RECT) As Long

Private Declare Function DestroyWindow Lib "user32" (ByVal HWND As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type TOOLINFO
    cbSize As Long
    uFlags As Long
    HWND As Long
    uID As Long
    RECT As RECT
    hinst As Long
    lpszText As String
    lParam As Long
End Type


Private Const CW_USEDEFAULT = &H80000000


Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOSIZE = &H1
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOP = 0
Private Const HWND_TOPMOST = -1
Private Const HWND_BOTTOM = 1


Private Const WS_POPUP = &H80000000
Private Const WS_EX_TOPMOST = &H8&


Private Const WM_USER = &H400

Private Const TTDT_AUTOMATIC = 0
Private Const TTDT_AUTOPOP = 2
Private Const TTDT_INITIAL = 3
Private Const TTDT_RESHOW = 1

' All of the flags for tool tip windows.
Private Const TTF_ABSOLUTE = &H80
Private Const TTF_CENTERTIP = &H2
Private Const TTF_DI_SETITEM = &H8000
Private Const TTF_IDISHWND = &H1
Private Const TTF_RTLREADING = &H4
Private Const TTF_SUBCLASS = &H10
Private Const TTF_TRACK = &H20
Private Const TTF_TRANSPARENT = &H100

' All of the available messages for tool tip windows.
Private Const TTM_ACTIVATE = (WM_USER + 1)
Private Const TTM_ADDTOOLA = (WM_USER + 4)
Private Const TTM_ADDTOOLW = (WM_USER + 50)
Private Const TTM_ADJUSTRECT = (WM_USER + 31)
Private Const TTM_DELTOOLA = (WM_USER + 5)
Private Const TTM_DELTOOLW = (WM_USER + 51)
Private Const TTM_ENUMTOOLSA = (WM_USER + 14)
Private Const TTM_ENUMTOOLSW = (WM_USER + 58)
Private Const TTM_GETBUBBLESIZE = (WM_USER + 30)
Private Const TTM_GETCURRENTTOOLA = (WM_USER + 15)
Private Const TTM_GETCURRENTTOOLW = (WM_USER + 59)
Private Const TTM_GETDELAYTIME = (WM_USER + 21)
Private Const TTM_GETMARGIN = (WM_USER + 27)
Private Const TTM_GETMAXTIPWIDTH = (WM_USER + 25)
Private Const TTM_GETTEXTA = (WM_USER + 11)
Private Const TTM_GETTEXTW = (WM_USER + 56)
Private Const TTM_GETTIPBKCOLOR = (WM_USER + 22)
Private Const TTM_GETTIPTEXTCOLOR = (WM_USER + 23)
Private Const TTM_GETTOOLCOUNT = (WM_USER + 13)
Private Const TTM_GETTOOLINFOA = (WM_USER + 8)
Private Const TTM_GETTOOLINFOW = (WM_USER + 53)
Private Const TTM_HITTESTA = (WM_USER + 10)
Private Const TTM_HITTESTW = (WM_USER + 55)
Private Const TTM_NEWTOOLRECTA = (WM_USER + 6)
Private Const TTM_NEWTOOLRECTW = (WM_USER + 52)
Private Const TTM_POP = (WM_USER + 28)
Private Const TTM_RELAYEVENT = (WM_USER + 7)
Private Const TTM_SETDELAYTIME = (WM_USER + 3)
Private Const TTM_SETMARGIN = (WM_USER + 26)
Private Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)
Private Const TTM_SETTIPBKCOLOR = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
Private Const TTM_SETTITLEA = (WM_USER + 32)
Private Const TTM_SETTITLEW = (WM_USER + 33)
Private Const TTM_SETTOOLINFOA = (WM_USER + 9)
Private Const TTM_SETTOOLINFOW = (WM_USER + 54)
Private Const TTM_TRACKACTIVATE = (WM_USER + 17)
Private Const TTM_TRACKPOSITION = (WM_USER + 18)
Private Const TTM_UPDATE = (WM_USER + 29)
Private Const TTM_UPDATETIPTEXTA = (WM_USER + 12)
Private Const TTM_UPDATETIPTEXTW = (WM_USER + 57)
Private Const TTM_WINDOWFROMPOINT = (WM_USER + 16)
Private Const TTS_ALWAYSTIP = &H1
'

Private Const TTS_BALLOON = &H40
'

Private Const TTS_NOANIMATE = &H10
'

Private Const TTS_NOFADE = &H20
'
Private Const TTS_NOPREFIX = &H2



Private Const TOOLTIPS_CLASS = "tooltips_class"
Private Const TOOLTIPS_CLASSA = "tooltips_class32"

'------------------------
'Tooltip window reference
'------------------------
Dim hwndTT As Long


'----------------------
'Destroy tooltip window
'----------------------
Public Sub DestroyTooltip()
    If (hwndTT <> 0) Then
        DestroyWindow hwndTT
        hwndTT = 0
    End If
End Sub


'------------
'Show tooltip
'------------
Public Sub DisplayTooltip(ControlHandle As Long, TooltipText As String, ApphInstance As Long)
    DestroyTooltip
    Dim ti As TOOLINFO
    
    Dim RECT As RECT
    
    Dim uID As Long
    uID = 0
    
    
    Dim strPntr As String
    strPntr = TooltipText
    
    Dim RetVal As Long
    
    
    
    hwndTT = CreateWindowEx(WS_EX_TOPMOST, TOOLTIPS_CLASSA, vbNullString, TTS_BALLOON Or WS_POPUP Or TTS_NOPREFIX, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, ControlHandle, 0, ApphInstance, 0)
    
    GetClientRect ControlHandle, RECT
    
    SetWindowPos hwndTT, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
    
    
    
    ti.cbSize = Len(ti)
    ti.uFlags = TTF_SUBCLASS
    '
    ' The hwnd of the control having the tool tip applied.
    ti.HWND = ControlHandle
    ti.hinst = ApphInstance
    ti.uID = uID
    ti.lpszText = strPntr
    ti.RECT = RECT

    
    
    RetVal = SendMessage(hwndTT, TTM_ADDTOOLA, 0, ti)
    
    RetVal = SendMessage(hwndTT, TTM_SETMAXTIPWIDTH, 0, 80)
    
    RetVal = SendMessage(hwndTT, TTM_SETTIPBKCOLOR, RGB(247, 252, 203), 0) 'cream
    RetVal = SendMessage(hwndTT, TTM_SETTIPTEXTCOLOR, RGB(0, 0, 0), 0)
    
    RetVal = SendMessage(hwndTT, TTM_UPDATETIPTEXTA, 0, ti)

End Sub

