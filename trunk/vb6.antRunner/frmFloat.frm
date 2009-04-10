VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C18424FD-277B-4362-B7A8-2788E7DBF8B4}#1.0#0"; "QProGIF.ocx"
Begin VB.Form frmFloat 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   300
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   795
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmFloat.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   300
   ScaleWidth      =   795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList Icons 
      Left            =   1560
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFloat.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "Icons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin VB.Image ResultIcon 
      Height          =   225
      Index           =   1
      Left            =   525
      Picture         =   "frmFloat.frx":027D
      Top             =   30
      Width           =   240
   End
   Begin VB.Image ResultIcon 
      Height          =   240
      Index           =   0
      Left            =   520
      Picture         =   "frmFloat.frx":037F
      Top             =   30
      Width           =   240
   End
   Begin prjQProGIF.QProGIF Gif 
      Height          =   255
      Left            =   520
      TabIndex        =   0
      Top             =   30
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Filename        =   "C:\work\google.code\vb6.antRunner\icons\ajax-loader.gif"
   End
End
Attribute VB_Name = "frmFloat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()


Private Sub MoveWindow(Button As Integer)
  Const WM_NCLBUTTONDOWN = &HA1
  Const HTCAPTION = 2
  If Button = vbLeftButton Then
    ReleaseCapture
    Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveWindow Button
End Sub


Public Property Let State(value As Long)
    Select Case value
        'successfull
        Case 1:
                Gif.Visible = False
                ResultIcon(0).Visible = True
                ResultIcon(1).Visible = False
        'failed
        Case 2:
                Gif.Visible = False
                ResultIcon(0).Visible = False
                ResultIcon(1).Visible = True
        'running
        Case 3:
                Gif.Visible = True
                ResultIcon(0).Visible = False
                ResultIcon(1).Visible = False
    End Select
End Property


Private Sub ResultIcon_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveWindow Button
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1:
                Runner.RunBuildTarget
    End Select
End Sub
