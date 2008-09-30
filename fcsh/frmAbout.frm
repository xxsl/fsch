VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   3915
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2702.203
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList iconList 
      Left            =   240
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   255
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAbout.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAbout.frx":08DC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   720
      Left            =   120
      Picture         =   "frmAbout.frx":0C2E
      ScaleHeight     =   505.68
      ScaleMode       =   0  'User
      ScaleWidth      =   505.68
      TabIndex        =   1
      Top             =   120
      Width           =   720
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4320
      TabIndex        =   0
      Top             =   3480
      Width           =   1260
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   1200
      Top             =   2280
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   1200
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Nimrod97@gmail.com"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   2040
      TabIndex        =   8
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "http://code.google.com/p/fsch/"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "http://www.famfamfam.com/lab/icons/silk/"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   3120
      Width           =   3135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1770.409
      Y2              =   1770.409
   End
   Begin VB.Label lblDescription 
      Caption         =   $"frmAbout.frx":2928
      ForeColor       =   &H00000000&
      Height          =   810
      Left            =   1050
      TabIndex        =   2
      Top             =   1125
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   4
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1780.762
      Y2              =   1780.762
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   225
      Left            =   1050
      TabIndex        =   5
      Top             =   780
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "This application uses Silk icon set made by Mark James:"
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   120
      TabIndex        =   3
      Top             =   2745
      Width           =   3975
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "About"
    lblVersion.Caption = "Version " & app.Major & "." & app.Minor & "." & app.Revision
    lblTitle.Caption = "Adobe Flex Compiler Shell Wrapper"
    Image1.Picture = iconList.ListImages(2).ExtractIcon
    Image2.Picture = iconList.ListImages(1).ExtractIcon
End Sub
