VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C18424FD-277B-4362-B7A8-2788E7DBF8B4}#1.0#0"; "QProGIF.ocx"
Begin VB.Form frmFloat 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   540
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   1245
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmFloat.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   540
   ScaleWidth      =   1245
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   540
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   953
      ButtonWidth     =   767
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Build"
            Style           =   5
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   600
      Left            =   720
      TabIndex        =   2
      Top             =   0
      Width           =   600
   End
   Begin prjQProGIF.QProGIF Gif 
      Height          =   495
      Left            =   730
      TabIndex        =   1
      Top             =   35
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Filename        =   "ajax-loader.gif"
   End
   Begin VB.Image buildIcon 
      Height          =   240
      Left            =   840
      Picture         =   "frmFloat.frx":000C
      Top             =   120
      Width           =   240
   End
End
Attribute VB_Name = "frmFloat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()

Private Const BUILD_STOPPED As String = "Stopped"
Private Const BUILD_ACTIVE As String = "Active"
Private Const BUILD_IDLE As String = "Success"
Private Const BUILD_ERROR As String = "Error"
Private Const BUILD_WARN As String = "Warning"

Private Sub Form_Load()
    Gif.FileName = MainForm.preloader.extractedFiles(MainForm.preloader.getIndex(EXEC_PNG))
    Toolbar1.ImageList = MainForm.pngImages
    Toolbar1.Buttons(1).Image = MainForm.preloader.getIndex(BUILD_TASK)
    showPicture STOPPED_PNG, BUILD_STOPPED
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Const WM_NCLBUTTONDOWN = &HA1
  Const HTCAPTION = 2
  If Button = vbLeftButton Then
    ReleaseCapture
    Call SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
End Sub


Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Const WM_NCLBUTTONDOWN = &HA1
  Const HTCAPTION = 2
  If Button = vbLeftButton Then
    ReleaseCapture
    Call SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = 2) Then
        PopupMenu MainForm.mnuShell
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 1:
            showAnimation
            MainForm.rebuild
    End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim Index As Long
    Index = Val(ButtonMenu.KEY)
    showAnimation
    MainForm.BUILD Index
End Sub

Private Sub buildIcon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Const WM_NCLBUTTONDOWN = &HA1
  Const HTCAPTION = 2
  If Button = vbLeftButton Then
    ReleaseCapture
    Call SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
End Sub

Public Sub idle()
    showPicture IDLE_PNG, BUILD_IDLE
End Sub

Public Sub stopped()
    showPicture STOPPED_PNG, BUILD_STOPPED
End Sub

Public Sub error()
    showPicture ERROR_PNG, BUILD_ERROR
End Sub

Public Sub warning()
    showPicture WARNING_PNG, BUILD_WARN
End Sub

Public Sub active()
    showAnimation
End Sub


Private Sub showPicture(Image As String, Text As String)
    Gif.Visible = False
    Gif.PauseAnimation = True
    buildIcon.Visible = True
    buildIcon.Picture = MainForm.pngImages.ListImages(MainForm.preloader.getIndex(Image)).ExtractIcon
    Label1.ToolTipText = Text
End Sub

Private Sub showAnimation()
    Gif.Visible = True
    Gif.PauseAnimation = False
    buildIcon.Visible = False
End Sub


