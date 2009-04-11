Attribute VB_Name = "modTray"
'***********************************************************************************
'* nimrod97@gmail.com                                                              *
'* Project homepage http://code.google.com/p/fsch/                                 *
'* Adobe Flex Compiler Shell wrapper                                               *
'* 2008                                                                            *
'***********************************************************************************

Option Explicit

'-------------------------------------------------
Private Const NOTIFYICON_VERSION = &H3
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_NULL = &H0

Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIM_SETFOCUS = &H3
Private Const NIM_SETVERSION = &H4
Private Const NIM_VERSION = &H5

Private Const NIS_HIDDEN = &H1
Private Const NIS_SHAREDICON = &H2

   
Private Const WM_USER = &H400
Private Const NIN_BALLOONSHOW = (WM_USER + 2)
Private Const NIN_BALLOONHIDE = (WM_USER + 3)
Private Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
Private Const NIN_BALLOONUSERCLICK = (WM_USER + 5)

'shell version / NOTIFIYICONDATA struct size constants
Private Const NOTIFYICONDATA_V1_SIZE As Long = 88  'pre-5.0 structure size
Private Const NOTIFYICONDATA_V2_SIZE As Long = 488 'pre-6.0 structure size
Private Const NOTIFYICONDATA_V3_SIZE As Long = 504 '6.0+ structure size

Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

Private Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 128
  dwState As Long
  dwStateMask As Long
  szInfo As String * 256
  uTimeoutAndVersion As Long
  szInfoTitle As String * 64
  dwInfoFlags As Long
  guidItem As GUID
End Type

Public Enum InfoIcon
        NIIF_NONE = &H0
        NIIF_INFO = &H1
        NIIF_WARNING = &H2
        NIIF_ERROR = &H3
        NIIF_GUID = &H5         ' Не используется
        NIIF_ICON_MASK = &HF    ' Не используется
        NIIF_NOSOUND = &H10
End Enum

'[Return Events]
Public Enum TrayRetunEventEnum
    MouseMove = &H200       'On Mousemove
    LeftUp = &H202          'Left Button Mouse Up
    LeftDown = &H201        'Left Button MouseDown
    LeftDbClick = &H203     'Left Button Double Click
    RightUp = &H205         'Right Button Up
    RightDown = &H204       'Right Button Down
    RightDbClick = &H206    'Right Button Double Click
    MiddleUp = &H208        'Middle Button Up
    MiddleDown = &H207      'Middle Button Down
    MiddleDbClick = &H209   'Middle Button Double Click
End Enum

'[Modify Items]
Public Enum ModifyItemEnum
    ToolTip = 1             'Modify ToolTip
    Icon = 2                'Modify Icon
End Enum

'[API]
Private TrayIcon As NOTIFYICONDATA
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Private Const NOTIFYICONDATA_SIZE As Long = 504 '6.0+ structure size

'[Add to Tray]
Public Sub TrayAdd(hwnd As Long, Icon As Picture, _
                    ToolTip As String, ReturnCallEvent As TrayRetunEventEnum)
    With TrayIcon
        .cbSize = Len(TrayIcon)
        .hwnd = hwnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = ReturnCallEvent
        .hIcon = Icon
        .szTip = ToolTip & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, TrayIcon
End Sub

'[Remove From tray]
Public Sub TrayDelete()
    Shell_NotifyIcon NIM_DELETE, TrayIcon
End Sub

'[Modify the tray]
Public Sub TrayModify(Item As ModifyItemEnum, vNewValue As Variant)
    Select Case Item
        Case ToolTip
            TrayIcon.szTip = vNewValue & vbNullChar
        Case Icon
            TrayIcon.hIcon = vNewValue
    End Select
    Shell_NotifyIcon NIM_MODIFY, TrayIcon
End Sub

Public Sub DisplayBalloon(ByVal sTitle As String, ByVal sText As String, ByVal info As InfoIcon)

Dim ret As Long
  
   With TrayIcon
      .cbSize = NOTIFYICONDATA_SIZE
      .uFlags = NIF_INFO
      .dwInfoFlags = info
      .szInfoTitle = sTitle & vbNullChar
      .szInfo = sText & vbNullChar
   End With

   ret = Shell_NotifyIcon(NIM_MODIFY, TrayIcon)
End Sub
