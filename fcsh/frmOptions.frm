VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmOptions 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   6330
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   8520
   FillColor       =   &H00FFFF80&
   ForeColor       =   &H8000000D&
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   8520
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   120
      ScaleHeight     =   3585
      ScaleWidth      =   8265
      TabIndex        =   21
      Top             =   2160
      Width           =   8295
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   3375
         Left            =   2040
         ScaleHeight     =   3345
         ScaleWidth      =   6105
         TabIndex        =   24
         Top             =   120
         Width           =   6135
         Begin VB.ComboBox cmbDebug 
            Height          =   315
            ItemData        =   "frmOptions.frx":1CFA
            Left            =   2400
            List            =   "frmOptions.frx":1D04
            TabIndex        =   45
            Top             =   2880
            Width           =   3495
         End
         Begin VB.CommandButton cmdLib 
            Caption         =   "+"
            Height          =   285
            Left            =   5520
            TabIndex        =   44
            Top             =   1440
            Width           =   375
         End
         Begin VB.CommandButton cmdOutput 
            Caption         =   "..."
            Height          =   285
            Left            =   5520
            TabIndex        =   43
            Top             =   1800
            Width           =   375
         End
         Begin VB.CommandButton cmdServices 
            Caption         =   "..."
            Height          =   285
            Left            =   5520
            TabIndex        =   42
            Top             =   2160
            Width           =   375
         End
         Begin VB.CommandButton cmdMxml 
            Caption         =   "..."
            Height          =   285
            Left            =   5520
            TabIndex        =   41
            Top             =   1080
            Width           =   375
         End
         Begin VB.TextBox txtTarget 
            Height          =   315
            Index           =   1
            Left            =   2400
            TabIndex        =   32
            Top             =   705
            Width           =   3495
         End
         Begin VB.TextBox txtTarget 
            Height          =   285
            Index           =   2
            Left            =   2400
            TabIndex        =   31
            Top             =   1080
            Width           =   3015
         End
         Begin VB.TextBox txtTarget 
            Height          =   285
            Index           =   4
            Left            =   2400
            TabIndex        =   30
            Top             =   1800
            Width           =   3015
         End
         Begin VB.TextBox txtTarget 
            Height          =   285
            Index           =   5
            Left            =   2400
            TabIndex        =   29
            Top             =   2160
            Width           =   3015
         End
         Begin VB.TextBox txtTarget 
            Height          =   285
            Index           =   6
            Left            =   2400
            TabIndex        =   28
            Top             =   2520
            Width           =   3495
         End
         Begin VB.TextBox txtTarget 
            Height          =   285
            Index           =   3
            Left            =   2400
            TabIndex        =   27
            Top             =   1440
            Width           =   3015
         End
         Begin VB.TextBox txtTarget 
            Height          =   285
            Index           =   0
            Left            =   2400
            TabIndex        =   26
            Top             =   360
            Width           =   3495
         End
         Begin VB.Label lblApp 
            Alignment       =   1  'Right Justify
            Caption         =   "Command"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   40
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label lblApp 
            Alignment       =   1  'Right Justify
            Caption         =   "Name"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   39
            Top             =   720
            Width           =   2295
         End
         Begin VB.Label lblApp 
            Alignment       =   1  'Right Justify
            Caption         =   "Application path (*.mxml)"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   38
            Top             =   1080
            Width           =   2295
         End
         Begin VB.Label lblApp 
            Alignment       =   1  'Right Justify
            Caption         =   "Output path"
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   37
            Top             =   1800
            Width           =   2295
         End
         Begin VB.Label lblApp 
            Alignment       =   1  'Right Justify
            Caption         =   "Service config path (*.xml)"
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   36
            Top             =   2160
            Width           =   2295
         End
         Begin VB.Label lblApp 
            Alignment       =   1  'Right Justify
            Caption         =   "Context root"
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   35
            Top             =   2520
            Width           =   2295
         End
         Begin VB.Label lblApp 
            Alignment       =   1  'Right Justify
            Caption         =   "Debug"
            Height          =   255
            Index           =   6
            Left            =   0
            TabIndex        =   34
            Top             =   2880
            Width           =   2295
         End
         Begin VB.Label lblApp 
            Alignment       =   1  'Right Justify
            Caption         =   "Include libraries (separated by ;)"
            Height          =   255
            Index           =   7
            Left            =   0
            TabIndex        =   33
            Top             =   1440
            Width           =   2295
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Application properties "
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   -10
            TabIndex        =   25
            Top             =   -10
            Width           =   1665
         End
      End
      Begin VB.ListBox lstApps 
         Appearance      =   0  'Flat
         Height          =   2955
         Left            =   120
         TabIndex        =   23
         Top             =   525
         Width           =   1815
      End
      Begin MSComctlLib.Toolbar AppToolbar 
         Height          =   330
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Appearance      =   1
         Style           =   1
         ImageList       =   "toolbarIcons"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Add aplication"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Rermove application"
               ImageIndex      =   2
            EndProperty
         EndProperty
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
               Picture         =   "frmOptions.frx":1D15
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":2067
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":23B9
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      FillColor       =   &H80000011&
      ForeColor       =   &H8000000D&
      Height          =   1935
      Left            =   120
      ScaleHeight     =   1905
      ScaleWidth      =   8265
      TabIndex        =   8
      Top             =   120
      Width           =   8295
      Begin VB.CommandButton cmdFcsh 
         Caption         =   "..."
         Height          =   285
         Left            =   7800
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
         Width           =   720
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
         Width           =   5775
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   2640
         TabIndex        =   12
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   327681
         Value           =   44000
         BuddyControl    =   "UpDown1"
         BuddyDispid     =   196621
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
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Preferences "
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   -15
         TabIndex        =   19
         Top             =   -15
         Width           =   975
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
         Left            =   0
         TabIndex        =   14
         Top             =   1440
         Width           =   1815
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   120
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Browse"
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7320
      TabIndex        =   6
      Top             =   5880
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

Private appsCollection As Collection

Private config As clsConfiguration
Private log As clsLog

Private isLoading As Boolean

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib "shell32" _
                                        (lpbi As BrowseInfo) As Long

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


Public Sub loadPrefs(ByRef cfg As clsConfiguration, ByRef logger As clsLog)
    Set log = logger
    Set config = cfg
    Set appsCollection = New Collection
    lstApps.Clear
    resetControls
    
    txtPort.Text = config.SERVER_PORT
    
    If (config.LOG_DEBUG) Then
        chkDebug.value = 1
    Else
        chkDebug.value = 0
    End If
    
    If (config.SHOW_BALOON) Then
        chkBaloon.value = 1
    Else
        chkBaloon.value = 0
    End If
    
    txtFcsh.Text = config.FCSH_PATH
    
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
    
    lstApps.Clear
    For Each app In appsCollection
        lstApps.AddItem app.fName
    Next
    resetControls
End Sub

Private Function uniqueName(name As String) As Boolean
    Dim app As clsTarget
    Dim isResult As Boolean
    
    isResult = True
    
    For Each app In appsCollection
        If (LCase(name) = LCase(app.fName)) Then
            isResult = False
        End If
    Next
    
    uniqueName = isResult
End Function


Private Sub cmbDebug_Click()
    Dim target As Long
    If (lstApps.ListIndex >= 0 And Not isLoading) Then
        target = lstApps.ListIndex + 1
        appsCollection.Item(target).fDebug = Str(cmbDebug.ListIndex = 0)
    End If
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
        If (Len(Trim(txtTarget(3).Text)) = 0) Then
            txtTarget(3).Text = CD1.FileName
        Else
            txtTarget(3).Text = txtTarget(3).Text + ";" + CD1.FileName
        End If
    End If
End Sub

Private Sub cmdMxml_Click()
    CD1.FileName = ""
    CD1.Filter = "MXML document (*.mxml)|*.mxml"
    CD1.ShowOpen
    If (Len(CD1.FileName) > 0) Then
        txtTarget(2).Text = CD1.FileName
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
    If (Len(Trim(sBuffer)) = 3) Then
        txtTarget(4).Text = sBuffer + txtTarget(1).Text + ".swf"
    End If
    If (Len(Trim(sBuffer)) > 3) Then
        txtTarget(4).Text = sBuffer + "\" + txtTarget(1).Text + ".swf"
    End If
End Sub

Private Sub cmdServices_Click()
    CD1.FileName = ""
    CD1.Filter = "XML document (*.xml)|*.xml"
    CD1.ShowOpen
    If (Len(CD1.FileName) > 0) Then
        txtTarget(5).Text = CD1.FileName
    End If
End Sub

Private Sub cmdSave_Click()
    'config.Clear

    config.LOG_DEBUG = (chkDebug.value = 1)
    config.SERVER_PORT = txtPort.Text
    config.SHOW_BALOON = (chkBaloon = 1)
    config.FCSH_PATH = txtFcsh.Text
    
    Dim i As Long
    Dim app As clsTarget
    i = 0
    For Each app In appsCollection
        i = i + 1
        config.saveApplication i, app
    Next
    
    config.APPLICATIONS = i
    
    MainForm.loadApps
    Me.Hide
End Sub




Private Sub lstApps_Click()
    isLoading = True
    Dim index As Long
    If (lstApps.ListIndex >= 0) Then
        index = lstApps.ListIndex + 1
        txtTarget(0).Text = appsCollection.Item(index).fCommand
        txtTarget(1).Text = appsCollection.Item(index).fName
        txtTarget(2).Text = appsCollection.Item(index).fSource
        txtTarget(3).Text = appsCollection.Item(index).fLibraries
        txtTarget(4).Text = appsCollection.Item(index).fOutput
        txtTarget(5).Text = appsCollection.Item(index).fServices
        txtTarget(6).Text = appsCollection.Item(index).fContext
        If (LCase(appsCollection.Item(index).fDebug) = "true") Then
            cmbDebug.ListIndex = 0
        Else
            cmbDebug.ListIndex = 1
        End If
    End If
    isLoading = False
End Sub



Private Sub txtTarget_Change(index As Integer)
    Dim target As Long
    If (lstApps.ListIndex >= 0 And Not isLoading) Then
        target = lstApps.ListIndex + 1
        appsCollection.Item(target).fCommand = txtTarget(0).Text
        appsCollection.Item(target).fName = txtTarget(1).Text
        appsCollection.Item(target).fSource = txtTarget(2).Text
        appsCollection.Item(target).fLibraries = txtTarget(3).Text
        appsCollection.Item(target).fOutput = txtTarget(4).Text
        appsCollection.Item(target).fServices = txtTarget(5).Text
        appsCollection.Item(target).fContext = txtTarget(6).Text
        appsCollection.Item(target).fDebug = Str(cmbDebug.ListIndex = 0)
    End If
End Sub

Private Sub resetControls()
    Dim i As Long
    For i = 0 To txtTarget.Count - 1
        txtTarget(i).Text = ""
    Next i
    cmbDebug.Text = ""
End Sub
