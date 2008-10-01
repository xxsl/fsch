VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFloat 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   570
   ClientLeft      =   15120
   ClientTop       =   11190
   ClientWidth     =   1395
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   570
   ScaleWidth      =   1395
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1005
      ButtonWidth     =   1085
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "  Build  "
            Object.ToolTipText     =   "Build"
            Style           =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "State"
            Object.ToolTipText     =   "Build state"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmFloat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()

Private Const ERROR_PNG As Long = 15
Private Const EXEC_PNG As Long = 14
Private Const STOPPED_PNG As Long = 13
Private Const IDLE_PNG As Long = 12

Private Const BUILD_STOPPED As String = "Stopped"
Private Const BUILD_ACTIVE As String = "Active"
Private Const BUILD_IDLE As String = "Idle"
Private Const BUILD_ERROR As String = "Error"

Private Sub Form_Load()
    Toolbar1.ImageList = MainForm.pngImages
    Toolbar1.Buttons(1).Image = 3
    Toolbar1.Buttons(2).Image = STOPPED_PNG
    Toolbar1.Buttons(2).ToolTipText = BUILD_STOPPED
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 1:
            MainForm.rebuild
            Toolbar1.Buttons(2).Image = EXEC_PNG
            Toolbar1.Buttons(2).ToolTipText = BUILD_ACTIVE
    End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim Index As Long
    Index = Val(ButtonMenu.key)
    MainForm.build Index
    Toolbar1.Buttons(2).Image = EXEC_PNG
    Toolbar1.Buttons(2).ToolTipText = BUILD_ACTIVE
End Sub

Private Sub Toolbar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Const WM_NCLBUTTONDOWN = &HA1
  Const HTCAPTION = 2
  If Button = vbLeftButton Then
    ReleaseCapture
    Call SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
End Sub

Public Sub idle()
    Toolbar1.Buttons(2).Image = IDLE_PNG
    Toolbar1.Buttons(2).ToolTipText = BUILD_IDLE
End Sub

Public Sub stopped()
    Toolbar1.Buttons(2).Image = STOPPED_PNG
    Toolbar1.Buttons(2).ToolTipText = BUILD_STOPPED
End Sub

Public Sub error()
    Toolbar1.Buttons(2).Image = ERROR_PNG
    Toolbar1.Buttons(2).ToolTipText = BUILD_ERROR
End Sub


