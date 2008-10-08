VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{EAFBDB7D-3DF2-42CB-86DA-2383D2449EA3}#1.0#0"; "HotKeyConfig.ocx"
Begin VB.Form frmOptions 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   4650
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   10335
   FillColor       =   &H00FFFF80&
   ForeColor       =   &H8000000D&
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   10335
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab TabDialog 
      Height          =   4215
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   7435
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      ForeColor       =   8404992
      MouseIcon       =   "frmOptions.frx":058A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   " Main"
      TabPicture(0)   =   "frmOptions.frx":05A6
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "UpDown1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtFcsh"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chkBaloon"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "chkDebug"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtPort"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdFcsh"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   " Appearance"
      TabPicture(1)   =   "frmOptions.frx":05C2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label12"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label11"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label10"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label9"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "UpDown2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "picFore"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "picBack"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtAlpha"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      TabCaption(2)   =   " Build"
      TabPicture(2)   =   "frmOptions.frx":05DE
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "toolbarIcons"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "picAppFrame"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lstApps"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdAdd"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdRemove"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      TabCaption(3)   =   " Hot Keys"
      TabPicture(3)   =   "frmOptions.frx":05FA
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label7"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "chkUseKey"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "hkConfig"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).ControlCount=   3
      Begin HotKeyConfig.HotKey hkConfig 
         Height          =   285
         Left            =   1920
         TabIndex        =   61
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
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
      Begin VB.CheckBox chkUseKey 
         Height          =   255
         Left            =   3600
         TabIndex        =   59
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtAlpha 
         Height          =   285
         Left            =   -73080
         TabIndex        =   53
         Top             =   150
         Width           =   840
      End
      Begin VB.PictureBox picBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -73080
         ScaleHeight     =   225
         ScaleWidth      =   825
         TabIndex        =   52
         ToolTipText     =   "Select color"
         Top             =   510
         Width           =   855
      End
      Begin VB.PictureBox picFore 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -73080
         ScaleHeight     =   225
         ScaleWidth      =   825
         TabIndex        =   51
         ToolTipText     =   "Select color"
         Top             =   870
         Width           =   855
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   375
         Left            =   -74040
         TabIndex        =   50
         ToolTipText     =   "Remove application"
         Top             =   3360
         Width           =   855
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   -74880
         TabIndex        =   49
         ToolTipText     =   "Add application"
         Top             =   3360
         Width           =   855
      End
      Begin VB.ListBox lstApps 
         Height          =   3180
         Left            =   -74880
         TabIndex        =   48
         Top             =   120
         Width           =   1695
      End
      Begin VB.PictureBox picAppFrame 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000010&
         Height          =   3735
         Left            =   -73080
         ScaleHeight     =   249
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   545
         TabIndex        =   20
         Top             =   80
         Width           =   8175
         Begin VB.CommandButton cmdRemoveOther 
            Caption         =   "-"
            Height          =   285
            Left            =   7560
            TabIndex        =   38
            ToolTipText     =   "remove option"
            Top             =   3250
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdOther 
            Caption         =   "+"
            Height          =   285
            Left            =   7080
            TabIndex        =   37
            ToolTipText     =   "Add option"
            Top             =   3250
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.ComboBox cmbOptions 
            Height          =   315
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   3250
            Visible         =   0   'False
            Width           =   4455
         End
         Begin VB.CommandButton cmdRemoveLib 
            Caption         =   "-"
            Height          =   285
            Left            =   7560
            TabIndex        =   35
            ToolTipText     =   "Remove library"
            Top             =   1440
            Width           =   375
         End
         Begin VB.ComboBox cmbLibs 
            Height          =   315
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   1440
            Width           =   4455
         End
         Begin VB.ComboBox cmbCommand 
            Height          =   315
            ItemData        =   "frmOptions.frx":0616
            Left            =   2520
            List            =   "frmOptions.frx":0618
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   315
            Width           =   5415
         End
         Begin VB.ComboBox cmbDebug 
            Height          =   315
            ItemData        =   "frmOptions.frx":061A
            Left            =   2520
            List            =   "frmOptions.frx":061C
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   2880
            Width           =   5415
         End
         Begin VB.CommandButton cmdLib 
            Caption         =   "+"
            Height          =   285
            Left            =   7080
            TabIndex        =   31
            ToolTipText     =   "Add library"
            Top             =   1440
            Width           =   375
         End
         Begin VB.CommandButton cmdOutput 
            Caption         =   "..."
            Height          =   285
            Left            =   7560
            TabIndex        =   30
            ToolTipText     =   "Select folder"
            Top             =   1800
            Width           =   375
         End
         Begin VB.CommandButton cmdServices 
            Caption         =   "..."
            Height          =   285
            Left            =   7560
            TabIndex        =   29
            ToolTipText     =   "Select file"
            Top             =   2160
            Width           =   375
         End
         Begin VB.CommandButton cmdMxml 
            Caption         =   "..."
            Height          =   285
            Left            =   7560
            TabIndex        =   28
            ToolTipText     =   "Select Flex application"
            Top             =   1080
            Width           =   375
         End
         Begin VB.TextBox txtTarget 
            Height          =   315
            Index           =   1
            Left            =   2520
            TabIndex        =   27
            Top             =   705
            Width           =   5415
         End
         Begin VB.TextBox txtTarget 
            Height          =   285
            Index           =   2
            Left            =   2520
            TabIndex        =   26
            Top             =   1080
            Width           =   4935
         End
         Begin VB.TextBox txtTarget 
            Height          =   285
            Index           =   4
            Left            =   2520
            TabIndex        =   25
            Top             =   1820
            Width           =   4935
         End
         Begin VB.TextBox txtTarget 
            Height          =   285
            Index           =   5
            Left            =   2520
            TabIndex        =   24
            Top             =   2160
            Width           =   4935
         End
         Begin VB.TextBox txtTarget 
            Height          =   285
            Index           =   6
            Left            =   2520
            TabIndex        =   23
            Top             =   2520
            Width           =   4935
         End
         Begin VB.TextBox txtTarget 
            Height          =   285
            Index           =   3
            Left            =   240
            TabIndex        =   22
            Top             =   2760
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtTarget 
            Height          =   285
            Index           =   0
            Left            =   240
            TabIndex        =   21
            Top             =   480
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Other options"
            Height          =   255
            Left            =   240
            TabIndex        =   47
            Top             =   3240
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.Label lblApp 
            Alignment       =   1  'Right Justify
            Caption         =   "Command"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   46
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label lblApp 
            Alignment       =   1  'Right Justify
            Caption         =   "Name"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   45
            Top             =   720
            Width           =   2295
         End
         Begin VB.Label lblApp 
            Alignment       =   1  'Right Justify
            Caption         =   "Application path (*.mxml, *.css)"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   44
            Top             =   1080
            Width           =   2295
         End
         Begin VB.Label lblApp 
            Alignment       =   1  'Right Justify
            Caption         =   "Output path"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   43
            Top             =   1800
            Width           =   2295
         End
         Begin VB.Label lblApp 
            Alignment       =   1  'Right Justify
            Caption         =   "Service config path (*.xml)"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   42
            Top             =   2160
            Width           =   2295
         End
         Begin VB.Label lblApp 
            Alignment       =   1  'Right Justify
            Caption         =   "Context root"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   41
            Top             =   2520
            Width           =   2295
         End
         Begin VB.Label lblApp 
            Alignment       =   1  'Right Justify
            Caption         =   "Debug"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   40
            Top             =   2880
            Width           =   2295
         End
         Begin VB.Label lblApp 
            Alignment       =   1  'Right Justify
            Caption         =   "Include libraries (separated by ;)"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   39
            Top             =   1440
            Width           =   2295
         End
      End
      Begin VB.CommandButton cmdFcsh 
         Caption         =   "..."
         Height          =   285
         Left            =   -67680
         TabIndex        =   13
         Top             =   1230
         Width           =   495
      End
      Begin VB.TextBox txtPort 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73080
         TabIndex        =   12
         Text            =   "44000"
         Top             =   150
         Width           =   825
      End
      Begin VB.CheckBox chkDebug 
         Height          =   255
         Left            =   -73080
         TabIndex        =   11
         Top             =   510
         Width           =   255
      End
      Begin VB.CheckBox chkBaloon 
         Height          =   255
         Left            =   -73080
         TabIndex        =   10
         Top             =   870
         Width           =   255
      End
      Begin VB.TextBox txtFcsh 
         Height          =   285
         Left            =   -73080
         TabIndex        =   9
         Top             =   1230
         Width           =   5295
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   -72254
         TabIndex        =   14
         Top             =   150
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   327681
         Value           =   44000
         BuddyControl    =   "txtPort"
         BuddyDispid     =   196633
         OrigLeft        =   2760
         OrigTop         =   360
         OrigRight       =   3015
         OrigBottom      =   735
         Max             =   65000
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComctlLib.ImageList toolbarIcons 
         Left            =   -74280
         Top             =   1800
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16711935
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":061E
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":0970
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":0CC2
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin ComCtl2.UpDown UpDown2 
         Height          =   285
         Left            =   -72239
         TabIndex        =   54
         Top             =   150
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   327681
         BuddyControl    =   "txtAlpha"
         BuddyDispid     =   196611
         OrigLeft        =   7560
         OrigTop         =   360
         OrigRight       =   7815
         OrigBottom      =   735
         Max             =   255
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Recompile key"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   270
         Width           =   1650
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Main window alpha"
         Height          =   255
         Left            =   -74880
         TabIndex        =   58
         Top             =   150
         Width           =   1650
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Console backcolor"
         Height          =   255
         Left            =   -74880
         TabIndex        =   57
         Top             =   510
         Width           =   1650
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Console forecolor"
         Height          =   255
         Left            =   -74880
         TabIndex        =   56
         Top             =   870
         Width           =   1650
      End
      Begin VB.Label Label12 
         Caption         =   "Clear log to see changes"
         Height          =   255
         Left            =   -71640
         TabIndex        =   55
         Top             =   750
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Server port"
         Height          =   255
         Left            =   -74880
         TabIndex        =   19
         Top             =   150
         Width           =   1650
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Show debug messages"
         Height          =   195
         Left            =   -74880
         TabIndex        =   18
         Top             =   510
         Width           =   1650
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Show baloon tips"
         Height          =   195
         Left            =   -74880
         TabIndex        =   17
         Top             =   870
         Width           =   1650
      End
      Begin VB.Label Label4 
         Caption         =   "Restart required!"
         Height          =   255
         Left            =   -71760
         TabIndex        =   16
         Top             =   150
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Path to fcsh.exe"
         Height          =   255
         Left            =   -74880
         TabIndex        =   15
         Top             =   1230
         Width           =   1650
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   960
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Browse"
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   7920
      TabIndex        =   7
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   9120
      TabIndex        =   6
      Top             =   4200
      Width           =   1095
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   5
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
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   4
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
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   3
         Top             =   300
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************
'* nimrod97@gmail.com                                                              *
'* Project homepage http://code.google.com/p/fsch/                                 *
'* Adobe Flex Compiler Shell wrapper                                               *
'* 2008                                                                            *
'***********************************************************************************

Option Explicit

Private Declare Function SHBrowseForFolder Lib "shell32" (lpBI As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Type BrowseInfo
         hWndOwner      As Long
         pIDLRoot       As Long
         pszDisplayName As Long
         lpszTitle      As Long
         ulFlags        As Long
         lpfnCallback   As Long
         lParam         As Long
         iImage         As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Const MAIN_TAB As Long = 0
Private Const APPEARANCE_TAB As Long = 1
Private Const BUILD_TAB As Long = 2
Private Const HOTKEYS_TAB As Long = 3

Private appsCollection As Collection

Private config As clsConfiguration
Private log As clsLog

Private isLoading As Boolean

Private optionDebug As New clsFixedCombo
Private optionCommand As New clsFixedCombo
Private optionLibs As New clsStringCombo
Private optionOther As New clsStringCombo








Private Sub cmdOther_Click()
    Dim other As String
    other = InputBox("Enter new option e.g. -debug=true", "Add option")
    
    If (Len(Trim(other)) > 0) Then
        optionOther.Add other
    End If

    Dim target As Long
    If (lstApps.ListIndex >= 0 And Not isLoading) Then
        target = lstApps.ListIndex + 1
        appsCollection.item(target).fOther = optionOther.Property
    End If
End Sub

Private Sub cmdRemoveLib_Click()
    If (cmbLibs.ListIndex >= 0) Then
        optionLibs.Remove cmbLibs.List(cmbLibs.ListIndex)
    End If
    Dim target As Long
    If (lstApps.ListIndex >= 0 And Not isLoading) Then
        target = lstApps.ListIndex + 1
        appsCollection.item(target).fLibraries = optionLibs.Property
    End If
End Sub



Private Sub cmdRemoveOther_Click()
    If (cmbOptions.ListIndex >= 0) Then
        optionOther.Remove cmbOptions.List(cmbOptions.ListIndex)
    End If
    Dim target As Long
    If (lstApps.ListIndex >= 0 And Not isLoading) Then
        target = lstApps.ListIndex + 1
        appsCollection.item(target).fOther = optionOther.Property
    End If
End Sub

Private Sub Form_Load()
    Dim debugItems(1 To 2) As String
    debugItems(1) = "false"
    debugItems(2) = "true"
    optionDebug.Bind cmbDebug, debugItems
    
    Dim commandItems(1 To 2) As String
    commandItems(1) = "mxmlc"
    commandItems(2) = "compc"
    optionCommand.Bind cmbCommand, commandItems
    
    optionLibs.Bind cmbLibs, ";"
    optionOther.Bind cmbOptions, " "
    
    'draw frame
    Dim padding As Long, corner As Long
    corner = 10
    padding = 4
    RoundRect picAppFrame.hDC, padding, padding, picAppFrame.ScaleWidth - padding, picAppFrame.ScaleHeight - padding, corner, corner

    TabDialog.TabPicture(MAIN_TAB) = MainForm.pngImages.ListImages(MainForm.preloader.getIndex(START_FCSH)).Picture
    TabDialog.TabPicture(APPEARANCE_TAB) = MainForm.pngImages.ListImages(MainForm.preloader.getIndex(APP_APPEARANCE)).Picture
    TabDialog.TabPicture(BUILD_TAB) = MainForm.pngImages.ListImages(MainForm.preloader.getIndex(BUILD_TASK)).Picture
    TabDialog.TabPicture(HOTKEYS_TAB) = MainForm.pngImages.ListImages(MainForm.preloader.getIndex(KEYBOARD)).Picture
End Sub


Public Sub loadPrefs(ByRef cfg As clsConfiguration, ByRef logger As clsLog)
    Set log = logger
    Set config = cfg
    Set appsCollection = New Collection
    lstApps.clear
    picAppFrame.ENABLED = False
    resetControls
    
    txtPort.Text = config.SERVER_PORT
    
    If (config.LOG_DEBUG) Then
        chkDebug.Value = 1
    Else
        chkDebug.Value = 0
    End If
    
    If (config.SHOW_BALOON) Then
        chkBaloon.Value = 1
    Else
        chkBaloon.Value = 0
    End If
    
    txtFcsh.Text = config.FCSH_PATH
    
    txtAlpha.Text = config.Alpha
    
    picBack.BackColor = config.BackColor
    picFore.BackColor = config.ForeColor
    
    chkUseKey.Value = config.RECOMPILE.ENABLED
    
    hkConfig.hotkey = config.RECOMPILE.KEY
    hkConfig.HotKeyModifier = config.RECOMPILE.MODIFYER
    
    Dim i As Long
    Dim app As clsTarget
    
    For i = 1 To config.APPLICATIONS
        Set app = config.LoadApplication(i)
        appsCollection.Add app
        lstApps.AddItem app.fName
    Next i
    
End Sub


Private Sub AppToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.index
            Case 1:
                    Dim newApp As New clsTarget
                    Dim appName As String
                    appName = InputBox("Enter application name", "New application")
                    If (uniqueName(appName)) Then
                        newApp.fCommand = "mxmlc"
                        newApp.fName = appName
                        newApp.fDebug = "false"
                        appsCollection.Add newApp
                    Else
                        MsgBox "This name already exists - " + appName, vbCritical
                    End If
            Case 2:
                    Dim index As Long
                    If (lstApps.ListIndex >= 0) Then
                        Dim i As Long
                        Dim name As String
                        name = lstApps.List(lstApps.ListIndex)
                        Dim app As clsTarget
                       
                        i = 1
                        For Each app In appsCollection
                            If (app.fName = name) Then
                                appsCollection.Remove i
                            End If
                            i = i + 1
                        Next
                      
                    End If
    End Select
    
    lstApps.clear
    For Each app In appsCollection
        lstApps.AddItem app.fName
    Next
    resetControls
End Sub

Private Function uniqueName(name As String) As Boolean
    Dim app As clsTarget
    Dim isResult As Boolean
    
    isResult = True
    
    isResult = (Len(Trim(name)) > 0)
    
    For Each app In appsCollection
        If (LCase(name) = LCase(app.fName)) Then
            isResult = False
        End If
    Next
    
    uniqueName = isResult
End Function



Private Sub cmbCommand_Click()
    Dim target As Long
    If (lstApps.ListIndex >= 0 And Not isLoading) Then
        target = lstApps.ListIndex + 1
        appsCollection.item(target).fCommand = optionCommand.Property
    End If
End Sub

Private Sub cmbDebug_Click()
    Dim target As Long
    If (lstApps.ListIndex >= 0 And Not isLoading) Then
        target = lstApps.ListIndex + 1
        appsCollection.item(target).fDebug = optionDebug.Property
    End If
End Sub


Private Sub cmdRemove_Click()
    Dim app As clsTarget
    Dim index As Long
    If (lstApps.ListIndex >= 0) Then
        Dim i As Long
        Dim name As String
        name = lstApps.List(lstApps.ListIndex)
        
        i = 1
        For Each app In appsCollection
            If (app.fName = name) Then
                appsCollection.Remove i
            End If
            i = i + 1
        Next
    End If
    lstApps.clear
    For Each app In appsCollection
       lstApps.AddItem app.fName
    Next
    resetControls
End Sub


Private Sub cmdAdd_Click()
   Dim app As clsTarget
   Dim newApp As New clsTarget
   Dim appName As String
   appName = InputBox("Enter application name", "New application")
   If (uniqueName(appName)) Then
       newApp.fCommand = "mxmlc"
       newApp.fName = appName
       newApp.fDebug = "false"
       appsCollection.Add newApp
   Else
       MsgBox "This name already exists - " + appName, vbCritical
   End If
   
   lstApps.clear
   For Each app In appsCollection
       lstApps.AddItem app.fName
   Next
   resetControls
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdFcsh_Click()
    CD1.FileName = ""
    CD1.Filter = "Adobe Flex Compiler Shell (fcsh.exe)|fcsh.exe"
    CD1.ShowOpen
    If (Len(CD1.FileName) > 0) Then
        txtFcsh.Text = CD1.FileName
    End If
End Sub


Private Sub cmdLib_Click()
    CD1.FileName = ""
    CD1.Filter = "Flex library (*.swc)|*.swc"
    CD1.ShowOpen
    If (Len(CD1.FileName) > 0) Then
        optionLibs.Add CD1.FileName
    End If
    Dim target As Long
    If (lstApps.ListIndex >= 0 And Not isLoading) Then
        target = lstApps.ListIndex + 1
        appsCollection.item(target).fLibraries = optionLibs.Property
    End If
End Sub

Private Sub cmdMxml_Click()
    CD1.FileName = ""
    CD1.Filter = "MXML document (*.mxml)|*.mxml|CSS document (*.css)|*.css"
    CD1.ShowOpen
    If (Len(CD1.FileName) > 0) Then
        txtTarget(2).Text = GetShortName(CD1.FileName)
    End If
End Sub

Private Sub cmdOutput_Click()
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo

    szTitle = "This is the title"
    With tBrowseInfo
            .hWndOwner = Me.hWnd
            .lpszTitle = lstrcat(szTitle, "")
            .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With

    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    End If
    
    Dim ext As String
    If (LCase(Trim(cmbCommand.Text)) = "mxmlc") Then
        ext = ".swf"
    Else
        ext = ".swc"
    End If
    
    If (Len(Trim(sBuffer)) = 3) Then
        txtTarget(4).Text = GetShortName(sBuffer) + txtTarget(1).Text + ext
    End If
    If (Len(Trim(sBuffer)) > 3) Then
        txtTarget(4).Text = GetShortName(sBuffer) + "\" + txtTarget(1).Text + ext
    End If
End Sub



Private Sub cmdServices_Click()
    CD1.FileName = ""
    CD1.Filter = "XML document (*.xml)|*.xml"
    CD1.ShowOpen
    If (Len(CD1.FileName) > 0) Then
        txtTarget(5).Text = GetShortName(CD1.FileName)
    End If
End Sub

Private Sub cmdSave_Click()
    'config.Clear

    config.LOG_DEBUG = (chkDebug.Value = 1)
    config.SERVER_PORT = txtPort.Text
    config.SHOW_BALOON = (chkBaloon = 1)
    config.FCSH_PATH = txtFcsh.Text
    config.Alpha = txtAlpha.Text
    
    config.BackColor = picBack.BackColor
    config.ForeColor = picFore.BackColor
    
    config.RECOMPILE.ENABLED = chkUseKey.Value
    config.RECOMPILE.KEY = hkConfig.hotkey
    config.RECOMPILE.MODIFYER = hkConfig.HotKeyModifier
    
    
    Dim i As Long
    Dim app As clsTarget
    i = 0
    For Each app In appsCollection
        i = i + 1
        config.saveApplication i, app
    Next
    
    config.APPLICATIONS = i
    
    Dim hotkeySetup As New clsHotKeySetup
    hotkeySetup.SetupKey config.RECOMPILE, MainForm.hotkey
    
    MainForm.loadApps
    Me.Hide
End Sub

Private Sub lstApps_Click()
    isLoading = True
    Dim index As Long
    If (lstApps.ListIndex >= 0) Then
        picAppFrame.ENABLED = True
        index = lstApps.ListIndex + 1
        optionCommand.Property = appsCollection.item(index).fCommand
        txtTarget(1).Text = appsCollection.item(index).fName
        txtTarget(2).Text = appsCollection.item(index).fSource
        optionLibs.Property = appsCollection.item(index).fLibraries
        txtTarget(4).Text = appsCollection.item(index).fOutput
        txtTarget(5).Text = appsCollection.item(index).fServices
        txtTarget(6).Text = appsCollection.item(index).fContext
        optionDebug.Property = appsCollection.item(index).fDebug
    Else
        picAppFrame.ENABLED = False
    End If
    isLoading = False
End Sub


Private Sub picBack_Click()
    CD1.Color = picBack.BackColor
    CD1.ShowColor
    picBack.BackColor = CD1.Color
End Sub

Private Sub picFore_Click()
    CD1.Color = picFore.BackColor
    CD1.ShowColor
    picFore.BackColor = CD1.Color
End Sub


Private Sub txtTarget_Change(index As Integer)
    Dim target As Long
    If (lstApps.ListIndex >= 0 And Not isLoading) Then
        target = lstApps.ListIndex + 1
        appsCollection.item(target).fCommand = cmbCommand.Text
        appsCollection.item(target).fName = txtTarget(1).Text
        appsCollection.item(target).fSource = txtTarget(2).Text
        appsCollection.item(target).fLibraries = optionLibs.Property
        appsCollection.item(target).fOutput = txtTarget(4).Text
        appsCollection.item(target).fServices = txtTarget(5).Text
        appsCollection.item(target).fContext = txtTarget(6).Text
        appsCollection.item(target).fDebug = optionDebug.Property
    End If
End Sub

Private Sub resetControls()
    Dim i As Long
    For i = 0 To txtTarget.Count - 1
        txtTarget(i).Text = ""
    Next i
    optionDebug.Reset
    optionCommand.Reset
    optionLibs.Reset
    optionOther.Reset
    picAppFrame.ENABLED = False
End Sub
