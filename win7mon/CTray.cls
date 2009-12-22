VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean


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
        NIIF_GUID = &H5
        NIIF_ICON_MASK = &HF
        NIIF_NOSOUND = &H10
End Enum

'-------------
'Return Events
'-------------
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

'------------
'Modify Items
'------------
    

Private Const NOTIFYICONDATA_SIZE As Long = 504 '6.0+ structure size


'-------------------
'Hold application tray reference
'-------------------
Private TrayIcon As NOTIFYICONDATA

Private isTrayAdded As Boolean


'----------------
'Add icon to Tray
'----------------
Public Sub Add(hwnd As Long, Icon As Picture, ToolTip As String, ReturnCallEvent As TrayRetunEventEnum)
    If (Not isTrayAdded) Then
        With TrayIcon
            .cbSize = Len(TrayIcon)
            .hwnd = hwnd
            .uID = vbNull
            .uFlags = NIF_ICON Or NIF_TIP
            .uCallbackMessage = ReturnCallEvent
            .hIcon = Icon
            .szTip = ToolTip & vbNullChar
        End With
        Shell_NotifyIcon NIM_ADD, TrayIcon
        isTrayAdded = True
    Else
        Err.Raise _
        Number:=1, _
        Description:="Tray already added.", _
        Source:="CTray.Add"
    End If
End Sub


'----------------
'Remove application icon from tray
'----------------
Public Sub Delete()
    If (isTrayAdded) Then
        Shell_NotifyIcon NIM_DELETE, TrayIcon
        isTrayAdded = False
    End If
End Sub


'---------------------------------
'Modify tray icon
'---------------------------------
Public Sub ModifyIcon(vNewValue As Variant)
    If (isTrayAdded) Then
        With TrayIcon
            .uFlags = NIF_ICON
            .hIcon = vNewValue
            .cbSize = Len(TrayIcon)
        End With
        Shell_NotifyIcon NIM_MODIFY, TrayIcon
    Else
        Err.Raise _
        Number:=2, _
        Description:="Tray is not added.", _
        Source:="CTray.ModifyIcon"
    End If
End Sub

'---------------------------------
'Modify info
'---------------------------------
Public Sub ModifyTooltip(vNewValue As Variant)
    If (isTrayAdded) Then
        With TrayIcon
            .szTip = vNewValue & vbNullChar
            .uFlags = NIF_TIP
            .cbSize = Len(TrayIcon)
        End With
        Shell_NotifyIcon NIM_MODIFY, TrayIcon
    Else
        Err.Raise _
        Number:=2, _
        Description:="Tray is not added.", _
        Source:="CTray.ModifyIcon"
    End If
End Sub


'-------------------
'Show baloon tooltip
'-------------------
Public Sub DisplayBalloon(ByVal sTitle As String, ByVal sText As String, ByVal info As InfoIcon)
    If (isTrayAdded) Then
       Dim result As Long
      
       With TrayIcon
          .uFlags = NIF_INFO
          .dwInfoFlags = info
          .szInfoTitle = sTitle & vbNullChar
          .szInfo = sText & vbNullChar
       End With
    
       TrayIcon.cbSize = Len(TrayIcon)
       result = Shell_NotifyIcon(NIM_MODIFY, TrayIcon)
    Else
        Err.Raise _
        Number:=2, _
        Description:="Tray is not added.", _
        Source:="CTray.ModifyIcon"
    End If
End Sub

