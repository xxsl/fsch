VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmOptions 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   6675
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
   ScaleHeight     =   6675
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picFrameNoName 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000010&
      Height          =   3975
      Left            =   120
      ScaleHeight     =   265
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   673
      TabIndex        =   21
      Top             =   2160
      Width           =   10095
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   375
         Left            =   1020
         TabIndex        =   47
         ToolTipText     =   "Remove application"
         Top             =   3420
         Width           =   855
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   180
         TabIndex        =   46
         ToolTipText     =   "Add application"
         Top             =   3420
         Width           =   855
      End
      Begin VB.PictureBox picAppFrame 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000010&
         Height          =   3735
         Left            =   1920
         ScaleHeight     =   249
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   537
         TabIndex        =   23
         Top             =   120
         Width           =   8055
         Begin VB.CommandButton cmdRemoveOther 
            Caption         =   "-"
            Height          =   285
            Left            =   7560
            TabIndex        =   53
            ToolTipText     =   "remove option"
            Top             =   3250
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdOther 
            Caption         =   "+"
            Height          =   285
            Left            =   7080
            TabIndex        =   52
            ToolTipText     =   "Add option"
            Top             =   3250
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.ComboBox cmbOptions 
            Height          =   315
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   51
            Top             =   3250
            Visible         =   0   'False
            Width           =   4455
         End
         Begin VB.CommandButton cmdRemoveLib 
            Caption         =   "-"
            Height          =   285
            Left            =   7560
            TabIndex        =   49
            ToolTipText     =   "Remove library"
            Top             =   1440
            Width           =   375
         End
         Begin VB.ComboBox cmbLibs 
            Height          =   315
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   1440
            Width           =   4455
         End
         Begin VB.ComboBox cmbCommand 
            Height          =   315
            ItemData        =   "frmOptions.frx":058A
            Left            =   2520
            List            =   "frmOptions.frx":058C
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Top             =   315
            Width           =   5415
         End
         Begin VB.ComboBox cmbDebug 
            Height          =   315
            ItemData        =   "frmOptions.frx":058E
            Left            =   2520
            List            =   "frmOptions.frx":0590
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   2880
            Width           =   5415
         End
         Begin VB.CommandButton cmdLib 
            Caption         =   "+"
            Height          =   285
            Left            =   7080
            TabIndex        =   43
            ToolTipText     =   "Add library"
            Top             =   1440
            Width           =   375
         End
         Begin VB.CommandButton cmdOutput 
            Caption         =   "..."
            Height          =   285
            Left            =   7560
            TabIndex        =   42
            ToolTipText     =   "Select folder"
            Top             =   1800
            Width           =   375
         End
         Begin VB.CommandButton cmdServices 
            Caption         =   "..."
            Height          =   285
            Left            =   7560
            TabIndex        =   41
            ToolTipText     =   "Select file"
            Top             =   2160
            Width           =   375
         End
         Begin VB.CommandButton cmdMxml 
            Caption         =   "..."
            Height          =   285
            Left            =   7560
            TabIndex        =   40
            ToolTipText     =   "Select Flex application"
            Top             =   1080
            Width           =   375
         End
         Begin VB.TextBox txtTarget 
            Height          =   315
            Index           =   1
            Left            =   2520
            TabIndex        =   31
            Top             =   705
            Width           =   5415
         End
         Begin VB.TextBox txtTarget 
            Height          =   285
            Index           =   2
            Left            =   2520
            TabIndex        =   30
            Top             =   1080
            Width           =   4935
         End
         Begin VB.TextBox txtTarget 
            Height          =   285
            Index           =   4
            Left            =   2520
            TabIndex        =   29
            Top             =   1820
            Width           =   4935
         End
         Begin VB.TextBox txtTarget 
            Height          =   285
            Index           =   5
            Left            =   2520
            TabIndex        =   28
            Top             =   2160
            Width           =   4935
         End
         Begin VB.TextBox txtTarget 
            Height          =   285
            Index           =   6
            Left            =   2520
            TabIndex        =   27
            Top             =   2520
            Width           =   4935
         End
         Begin VB.TextBox txtTarget 
            Height          =   285
            Index           =   3
            Left            =   240
            TabIndex        =   26
            Top             =   2760
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtTarget 
            Height          =   285
            Index           =   0
            Left            =   240
            TabIndex        =   25
            Top             =   480
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Other options"
            Height          =   255
            Left            =   240
            TabIndex        =   50
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
            TabIndex        =   39
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label lblApp 
            Alignment       =   1  'Right Justify
            Caption         =   "Name"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   38
            Top             =   720
            Width           =   2295
         End
         Begin VB.Label lblApp 
            Alignment       =   1  'Right Justify
            Caption         =   "Application path (*.mxml, *.css)"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   37
            Top             =   1080
            Width           =   2295
         End
         Begin VB.Label lblApp 
            Alignment       =   1  'Right Justify
            Caption         =   "Output path"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   36
            Top             =   1800
            Width           =   2295
         End
         Begin VB.Label lblApp 
            Alignment       =   1  'Right Justify
            Caption         =   "Service config path (*.xml)"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   35
            Top             =   2160
            Width           =   2295
         End
         Begin VB.Label lblApp 
            Alignment       =   1  'Right Justify
            Caption         =   "Context root"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   34
            Top             =   2520
            Width           =   2295
         End
         Begin VB.Label lblApp 
            Alignment       =   1  'Right Justify
            Caption         =   "Debug"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   33
            Top             =   2880
            Width           =   2295
         End
         Begin VB.Label lblApp 
            Alignment       =   1  'Right Justify
            Caption         =   "Include libraries (separated by ;)"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   32
            Top             =   1440
            Width           =   2295
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   " Application properties "
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   225
            TabIndex        =   24
            Top             =   -30
            Width           =   1605
         End
      End
      Begin VB.ListBox lstApps 
         Height          =   3180
         Left            =   180
         TabIndex        =   22
         Top             =   180
         Width           =   1695
      End
      Begin MSComctlLib.ImageList toolbarIcons 
         Left            =   840
         Top             =   2160
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
               Picture         =   "frmOptions.frx":0592
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":08E4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":0C36
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picFramePrefs 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      FillColor       =   &H80000011&
      ForeColor       =   &H80000010&
      Height          =   1935
      Left            =   120
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   673
      TabIndex        =   8
      Top             =   120
      Width           =   10095
      Begin VB.PictureBox picFore 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6840
         ScaleHeight     =   225
         ScaleWidth      =   825
         TabIndex        =   60
         ToolTipText     =   "Select color"
         Top             =   1080
         Width           =   855
      End
      Begin VB.PictureBox picBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6840
         ScaleHeight     =   225
         ScaleWidth      =   825
         TabIndex        =   59
         ToolTipText     =   "Select color"
         Top             =   720
         Width           =   855
      End
      Begin ComCtl2.UpDown UpDown2 
         Height          =   285
         Left            =   7680
         TabIndex        =   56
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   327681
         BuddyControl    =   "txtAlpha"
         BuddyDispid     =   196632
         OrigLeft        =   7560
         OrigTop         =   360
         OrigRight       =   7815
         OrigBottom      =   735
         Max             =   255
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtAlpha 
         Height          =   285
         Left            =   6840
         TabIndex        =   55
         Top             =   360
         Width           =   840
      End
      Begin VB.CommandButton cmdFcsh 
         Caption         =   "..."
         Height          =   285
         Left            =   4800
         TabIndex        =   20
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtPort 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1920
         TabIndex        =   13
         Text            =   "44000"
         Top             =   360
         Width           =   825
      End
      Begin VB.CheckBox chkDebug 
         Height          =   255
         Left            =   1920
         TabIndex        =   11
         Top             =   720
         Width           =   255
      End
      Begin VB.CheckBox chkBaloon 
         Height          =   255
         Left            =   1920
         TabIndex        =   10
         Top             =   1080
         Width           =   255
      End
      Begin VB.TextBox txtFcsh 
         Height          =   285
         Left            =   1920
         TabIndex        =   9
         Top             =   1440
         Width           =   2775
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   2745
         TabIndex        =   12
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   327681
         Value           =   44000
         BuddyControl    =   "txtPort"
         BuddyDispid     =   196634
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
      Begin VB.Label Label12 
         Caption         =   "Clear log to see changes"
         Height          =   375
         Left            =   7920
         TabIndex        =   61
         Top             =   960
         Width           =   2055
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   352
         X2              =   352
         Y1              =   4
         Y2              =   124
      End
      Begin VB.Label Label11 
         Caption         =   "Log forecolor"
         Height          =   255
         Left            =   5520
         TabIndex        =   58
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Log backcolor"
         Height          =   255
         Left            =   5520
         TabIndex        =   57
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Transparency"
         Height          =   255
         Left            =   5520
         TabIndex        =   54
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   " Preferences "
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   225
         TabIndex        =   19
         Top             =   -30
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Server port"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1680
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Show debug messages"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   1650
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Show baloon tips"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   1650
      End
      Begin VB.Label Label4 
         Caption         =   "Restart required!"
         Height          =   255
         Left            =   3000
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Path to fcsh.exe"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   1650
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   360
      Top             =   6120
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
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   9100
      TabIndex        =   6
      Top             =   6240
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

Private Declare Function SHBrowseForFolder Lib "shell32" _
                                        (lpBI As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" _
                                        (ByVal pidList As Long, _
                                        ByVal lpBuffer As String) As Long

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
                                        (ByVal lpString1 As String, ByVal _
                                        lpString2 As String) As Long
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
    
    Dim padding As Long, corner As Long
    corner = 10
    padding = 4
    RoundRect picFramePrefs.hDC, padding, padding, picFramePrefs.ScaleWidth - padding, picFramePrefs.ScaleHeight - padding, corner, corner
    RoundRect picFrameNoName.hDC, padding, padding, picFrameNoName.ScaleWidth - padding, picFrameNoName.ScaleHeight - padding, corner, corner
    RoundRect picAppFrame.hDC, padding, padding, picAppFrame.ScaleWidth - padding, picAppFrame.ScaleHeight - padding, corner, corner

End Sub


Public Sub loadPrefs(ByRef cfg As clsConfiguration, ByRef logger As clsLog)
    Set log = logger
    Set config = cfg
    Set appsCollection = New Collection
    lstApps.clear
    picAppFrame.Enabled = False
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
    
    Dim I As Long
    Dim app As clsTarget
    
    For I = 1 To config.APPLICATIONS
        Set app = config.LoadApplication(I)
        appsCollection.Add app
        lstApps.AddItem app.fName
    Next I
    
End Sub


Private Sub AppToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
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
                    Dim Index As Long
                    If (lstApps.ListIndex >= 0) Then
                        Dim I As Long
                        Dim name As String
                        name = lstApps.List(lstApps.ListIndex)
                        Dim app As clsTarget
                       
                        I = 1
                        For Each app In appsCollection
                            If (app.fName = name) Then
                                appsCollection.Remove I
                            End If
                            I = I + 1
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
    Dim Index As Long
    If (lstApps.ListIndex >= 0) Then
        Dim I As Long
        Dim name As String
        name = lstApps.List(lstApps.ListIndex)
        
        I = 1
        For Each app In appsCollection
            If (app.fName = name) Then
                appsCollection.Remove I
            End If
            I = I + 1
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
    
    Dim I As Long
    Dim app As clsTarget
    I = 0
    For Each app In appsCollection
        I = I + 1
        config.saveApplication I, app
    Next
    
    config.APPLICATIONS = I
    
    MainForm.loadApps
    Me.Hide
End Sub







Private Sub lstApps_Click()
    isLoading = True
    Dim Index As Long
    If (lstApps.ListIndex >= 0) Then
        picAppFrame.Enabled = True
        Index = lstApps.ListIndex + 1
        optionCommand.Property = appsCollection.item(Index).fCommand
        txtTarget(1).Text = appsCollection.item(Index).fName
        txtTarget(2).Text = appsCollection.item(Index).fSource
        optionLibs.Property = appsCollection.item(Index).fLibraries
        txtTarget(4).Text = appsCollection.item(Index).fOutput
        txtTarget(5).Text = appsCollection.item(Index).fServices
        txtTarget(6).Text = appsCollection.item(Index).fContext
        optionDebug.Property = appsCollection.item(Index).fDebug
    Else
        picAppFrame.Enabled = False
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

Private Sub txtTarget_Change(Index As Integer)
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
    Dim I As Long
    For I = 0 To txtTarget.Count - 1
        txtTarget(I).Text = ""
    Next I
    optionDebug.Reset
    optionCommand.Reset
    optionLibs.Reset
    optionOther.Reset
    picAppFrame.Enabled = False
End Sub
