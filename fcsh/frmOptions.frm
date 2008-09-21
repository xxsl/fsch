VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   5970
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   8520
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   8520
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Applications"
      Height          =   3735
      Left            =   120
      TabIndex        =   17
      Top             =   1680
      Width           =   8295
      Begin VB.Frame Frame3 
         Height          =   3495
         Left            =   2040
         TabIndex        =   20
         Top             =   120
         Width           =   6135
         Begin VB.TextBox txtTarget 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   2520
            TabIndex        =   37
            Top             =   480
            Width           =   3495
         End
         Begin VB.TextBox txtTarget 
            Height          =   285
            Index           =   7
            Left            =   2520
            TabIndex        =   36
            Top             =   3000
            Width           =   3495
         End
         Begin VB.TextBox txtTarget 
            Height          =   285
            Index           =   3
            Left            =   2520
            TabIndex        =   34
            Top             =   1560
            Width           =   3495
         End
         Begin VB.TextBox txtTarget 
            Height          =   285
            Index           =   6
            Left            =   2520
            TabIndex        =   33
            Top             =   2640
            Width           =   3495
         End
         Begin VB.TextBox txtTarget 
            Height          =   285
            Index           =   5
            Left            =   2520
            TabIndex        =   32
            Top             =   2280
            Width           =   3495
         End
         Begin VB.TextBox txtTarget 
            Height          =   285
            Index           =   4
            Left            =   2520
            TabIndex        =   31
            Top             =   1920
            Width           =   3495
         End
         Begin VB.TextBox txtTarget 
            Height          =   285
            Index           =   2
            Left            =   2520
            TabIndex        =   30
            Top             =   1200
            Width           =   3495
         End
         Begin VB.TextBox txtTarget 
            Height          =   315
            Index           =   1
            Left            =   2520
            TabIndex        =   29
            Top             =   825
            Width           =   3495
         End
         Begin VB.Label Label5 
            Caption         =   "Aplication properties"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label lblApp 
            Alignment       =   1  'Right Justify
            Caption         =   "Include libraries (separated by ;)"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   28
            Top             =   1560
            Width           =   2295
         End
         Begin VB.Label lblApp 
            Alignment       =   1  'Right Justify
            Caption         =   "Debug"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   27
            Top             =   3000
            Width           =   2295
         End
         Begin VB.Label lblApp 
            Alignment       =   1  'Right Justify
            Caption         =   "Context root"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   26
            Top             =   2640
            Width           =   2295
         End
         Begin VB.Label lblApp 
            Alignment       =   1  'Right Justify
            Caption         =   "Service config path (*.xml)"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   25
            Top             =   2280
            Width           =   2295
         End
         Begin VB.Label lblApp 
            Alignment       =   1  'Right Justify
            Caption         =   "Output path"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   24
            Top             =   1920
            Width           =   2295
         End
         Begin VB.Label lblApp 
            Alignment       =   1  'Right Justify
            Caption         =   "Application path (*.mxml)"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   23
            Top             =   1200
            Width           =   2295
         End
         Begin VB.Label lblApp 
            Alignment       =   1  'Right Justify
            Caption         =   "Name"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   22
            Top             =   840
            Width           =   2295
         End
         Begin VB.Label lblApp 
            Alignment       =   1  'Right Justify
            Caption         =   "Command"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   21
            Top             =   480
            Width           =   2295
         End
      End
      Begin MSComctlLib.Toolbar AppToolbar 
         Height          =   330
         Left            =   120
         TabIndex        =   19
         Top             =   240
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
         Left            =   960
         Top             =   2280
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16711935
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":1CFA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":204C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.ListBox lstApps 
         Appearance      =   0  'Flat
         Height          =   2955
         Left            =   120
         TabIndex        =   18
         Top             =   645
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   6120
      TabIndex        =   8
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7320
      TabIndex        =   7
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Preferences"
      Height          =   1455
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   8295
      Begin VB.CheckBox chkBaloon 
         Height          =   255
         Left            =   2040
         TabIndex        =   15
         Top             =   1080
         Width           =   255
      End
      Begin VB.CheckBox chkDebug 
         Height          =   255
         Left            =   2040
         TabIndex        =   14
         Top             =   720
         Width           =   255
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   2760
         TabIndex        =   13
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   327681
         Value           =   44000
         BuddyControl    =   "txtPort"
         BuddyDispid     =   196620
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
      Begin VB.TextBox txtPort 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         TabIndex        =   12
         Text            =   "44000"
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label4 
         Caption         =   "(restart required)"
         Height          =   255
         Left            =   3120
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Show baloon tips"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   1650
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Show debug messages"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   1650
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Server port"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1680
      End
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
Option Explicit

Private appsCollection As Collection

Private config As clsConfiguration
Private log As clsLog

Private isLoading As Boolean

Public Sub loadPrefs(ByRef cfg As clsConfiguration, ByRef logger As clsLog)
    Set log = logger
    Set config = cfg
    Set appsCollection = New Collection
    lstApps.Clear
    
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


Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdSave_Click()
    'config.Clear

    config.LOG_DEBUG = (chkDebug.value = 1)
    config.SERVER_PORT = txtPort.Text
    config.SHOW_BALOON = (chkBaloon = 1)
    
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
        txtTarget(7).Text = appsCollection.Item(index).fDebug
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
        appsCollection.Item(target).fDebug = txtTarget(7).Text
    End If
End Sub

Private Sub resetControls()
    Dim i As Long
    For i = 0 To txtTarget.Count - 1
        txtTarget(i).Text = ""
    Next i
End Sub
