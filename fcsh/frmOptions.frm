VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   6000
   ClientLeft      =   2565
   ClientTop       =   1380
   ClientWidth     =   6675
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Applications"
      Height          =   3735
      Left            =   120
      TabIndex        =   17
      Top             =   1680
      Width           =   6495
      Begin MSComctlLib.Toolbar Toolbar1 
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
               Picture         =   "frmOptions.frx":000C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":035E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.ListBox lstApps 
         Appearance      =   0  'Flat
         Height          =   2955
         Left            =   120
         TabIndex        =   18
         Top             =   650
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5520
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
      Width           =   6495
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
         BuddyDispid     =   196614
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
         Appearance      =   0  'Flat
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

Private config As clsConfiguration

Public Sub loadPrefs(ByRef cfg As clsConfiguration)
    Set config = cfg
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
End Sub


Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdSave_Click()
    config.LOG_DEBUG = (chkDebug.value = 1)
    config.SERVER_PORT = txtPort.Text
    config.SHOW_BALOON = (chkBaloon = 1)
    Me.Hide
End Sub


