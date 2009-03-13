VERSION 5.00
Object = "{EAFBDB7D-3DF2-42CB-86DA-2383D2449EA3}#1.0#0"; "HotKeyConfig.ocx"
Begin VB.Form MainForm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "HotKey setup"
   ClientHeight    =   945
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Enabled"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin HotKeyConfig.HotKey HotKey1 
      Height          =   300
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   529
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
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private prefs As New clsHotKey
Private manager As New ARINIManager


Private Sub Command1_Click()
    prefs.KEY = HotKey1.HotKey
    prefs.MODIFYER = HotKey1.HotKeyModifier
    prefs.ENABLED = Check1.Value
End Sub

Private Sub Form_Load()
    manager.INIFile = App.Path + "\server.ini"
    prefs.init manager, "hotkey"
    
    HotKey1.HotKey = prefs.KEY
    HotKey1.HotKeyModifier = prefs.MODIFYER
    Check1.Value = prefs.ENABLED
End Sub
