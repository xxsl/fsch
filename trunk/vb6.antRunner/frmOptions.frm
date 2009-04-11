VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{EAFBDB7D-3DF2-42CB-86DA-2383D2449EA3}#1.0#0"; "HotKeyConfig.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   2445
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   6150
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD 
      Left            =   3840
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAntpath 
      Caption         =   "..."
      Height          =   285
      Left            =   5400
      TabIndex        =   16
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox txtAntpath 
      Height          =   285
      Left            =   960
      TabIndex        =   15
      Top             =   1080
      Width           =   4335
   End
   Begin VB.CheckBox chkMinimize 
      Caption         =   "Minimize on close"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   1695
   End
   Begin VB.CheckBox chkHotkey 
      Caption         =   "Enabled"
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   1560
      Width           =   975
   End
   Begin HotKeyConfig.HotKey HotKey 
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotKeyModifier  =   0
   End
   Begin VB.CheckBox chkBaloon 
      Caption         =   "Use baloon tooltips"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   7
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   6
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   5
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   1935
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Ant path"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "* Restart required"
      Height          =   255
      Left            =   3360
      TabIndex        =   13
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Build key"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   855
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private config As New clsConfig


Private Sub cmdAntpath_Click()
   CD.Filter = "Ant (*.bat)|*.bat"
   CD.ShowOpen
   If (Len(CD.FileName) > 0) Then
      txtAntpath.Text = CD.FileName
   End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    config.ShowBaloon = chkBaloon.value = 1
    config.MinimizeOnClose = chkMinimize.value = 1
    config.HotKeyEnabled = chkHotkey.value = 1
    config.Key = HotKey.HotKey
    config.Modifyer = HotKey.HotKeyModifier
    config.AntPath = txtAntpath.Text
    Unload Me
End Sub


Private Sub Form_Load()
    If (config.ShowBaloon) Then
        chkBaloon.value = 1
    Else
        chkBaloon.value = 0
    End If
    
    If (config.MinimizeOnClose) Then
        chkMinimize.value = 1
    Else
        chkMinimize.value = 0
    End If
    
    If (config.HotKeyEnabled) Then
        chkHotkey.value = 1
    Else
        chkHotkey.value = 0
    End If
    
    HotKey.HotKey = config.Key
    HotKey.HotKeyModifier = config.Modifyer
    
    txtAntpath.Text = config.AntPath
End Sub
