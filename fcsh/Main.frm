VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Main 
   AutoRedraw      =   -1  'True
   Caption         =   "Flex compiler shell"
   ClientHeight    =   6585
   ClientLeft      =   3840
   ClientTop       =   2220
   ClientWidth     =   9750
   Icon            =   "Main.frx":0000
   ScaleHeight     =   6585
   ScaleWidth      =   9750
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   6960
      TabIndex        =   5
      Top             =   4200
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Top             =   4200
      Width           =   615
   End
   Begin MSComctlLib.ImageList disabledIcons 
      Left            =   8520
      Top             =   2760
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
            Picture         =   "Main.frx":355A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":38AC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList enabledIcons 
      Left            =   8520
      Top             =   2040
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
            Picture         =   "Main.frx":3BFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":3F50
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton fakeTray 
      Caption         =   "fakeTray"
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSWinsockLib.Winsock Abort 
      Left            =   7920
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Service 
      Left            =   7920
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   2
      Top             =   6315
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   476
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "enabledIcons"
      DisabledImageList=   "disabledIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Start fcsh"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtbLog 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   9975
      _Version        =   393217
      BackColor       =   -2147483633
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Main.frx":42A2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSWinsockLib.Winsock Server 
      Left            =   7920
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   44000
   End
   Begin VB.Menu mnuShell 
      Caption         =   "Shell"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private log As New clsLog
Public config As New clsConfiguration
Dim WithEvents fcsh As clsFCSH
Attribute fcsh.VB_VarHelpID = -1

Private isServerBusy As Boolean
Private PORT As Long



'***********************************************************************************************
'fcsh event handling
'***********************************************************************************************
Private Sub fcsh_onError(ByVal Msg As String)
    log.xError "fcsh:" + Msg
    DisplayBalloon "Flex compiler shell", Msg, NIIF_ERROR
End Sub

'on command success
Private Sub fcsh_onFinish()
    log.xFcsh "Exec completed"
    log.Text vbCrLf
    DisplayBalloon "Flex compiler shell", "Build successfull", NIIF_INFO
End Sub

'on new compile target id
Private Sub fcsh_onIdAssigned(ByVal id As Long)
    log.xInfo "Target id = " & id
    DisplayBalloon "Flex compiler shell", "Assigned Target id " & id, NIIF_INFO
End Sub

'on fcsh.exe start
Private Sub fcsh_onStart()
   Toolbar.Buttons.Item(1).Image = 2
   Toolbar.Buttons.Item(1).ToolTipText = "Stop fcsh"
End Sub

'on fcsh.exe stop
Private Sub fcsh_onStop()
   Toolbar.Buttons.Item(1).Image = 1
   Toolbar.Buttons.Item(1).ToolTipText = "Start fcsh"
End Sub

'***********************************************************************************************
'Start app
'***********************************************************************************************
'start application
Private Sub Form_Load()
    'set up logging
    log.clsLog rtbLog
    
    'init prefs
    config.logger = log
    config.Load
    
    Dim target As clsTarget
    Set target = config.LoadApplication(config.APPLICATIONS)
        
    'set loglevel
    log.LogLevel = config.LOG_DEBUG
        
    'init vars
    isServerBusy = False
    initSockets 'listen for requests
    
    'setup fcsh
    Set fcsh = New clsFCSH
    fcsh.Initialize log
        
    'add tray icon
    TrayAdd fakeTray.hwnd, Me.Icon, "Flex compiler shell", MouseMove
    
    'log and show tooltip
    log.xDebug "Application initialized"
End Sub



'***********************************************************************************************
'Sockets
'***********************************************************************************************
'init main server socket
Private Sub initSockets()
    log.xInfo "Server is listening on port " & config.SERVER_PORT
    Server.Close
    Server.LocalPort = config.SERVER_PORT
    On Error Resume Next
    Server.Listen
    If Err.Number <> 0 Then
        log.xError "Cant start server: " + Err.Description
        Err.Clear
    End If
End Sub




'one more connection
Private Sub Server_ConnectionRequest(ByVal requestID As Long)
    If (Not isServerBusy) Then
        log.xDebug "Accepted connection request " & requestID
        Service.Close
        Service.Accept requestID
    Else
        log.xError "Server is busy. Ignoring connection request " & requestID
        Abort.Accept requestID
        Abort.SendData "busy"
        Abort.Close
    End If
End Sub

Private Sub Abort_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    log.xError "Server socket: " + Description
    Server.Close
End Sub

Private Sub Server_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    log.xError "Abort socket: " + Description
    Abort.Close
End Sub

Private Sub Service_Close()
    log.xDebug "Service connection closed"
End Sub

Private Sub Service_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    log.xError "Service socket: " + Description
    Service.Close
End Sub

Private Sub Service_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    Dim intCnt As Integer

    Service.GetData strData, vbString
    
    log.xDebug "Socket in: " + strData
End Sub



'***********************************************************************************************
'toolbar
'***********************************************************************************************
'click
Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.index
        Case 1:
                If (Not fcsh.isRunning) Then
                    fcsh.Start
                Else
                    fcsh.Quit
                End If
    End Select
End Sub


'***********************************************************************************************
'basic
'***********************************************************************************************
'tray events
Private Sub fakeTray_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cEvent As Single
    cEvent = X / Screen.TwipsPerPixelX
    Select Case cEvent
        Case MouseMove
            Debug.Print "MouseMove"
        Case LeftUp
            log.xDebug "Left Up"
        Case LeftDown
            log.xDebug "Left Down"
        Case LeftDbClick
            log.xDebug "LeftDbClick"
        Case MiddleUp
            log.xDebug "MiddleUp"
        Case MiddleDown
            log.xDebug "MiddleDown"
        Case MiddleDbClick
            log.xDebug "MiddleDbClick"
        Case RightUp
            log.xDebug "RightUp" ': PopupMenu mnuShell
        Case RightDown
            log.xDebug "RightDown"
        Case RightDbClick
            log.xDebug "RightDbClick"
    End Select
End Sub

'on resize
Private Sub Form_Resize()
    rtbLog.Width = Me.Width - 130
    
    Dim logWidth As Long
    logWidth = Me.Height - rtbLog.Top - 500 - StatusBar.Height
    If (logWidth > 0) Then
        rtbLog.Height = logWidth
    End If
End Sub

'on quit
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    log.xDebug "Application stopped"
    TrayDelete
    Server.Close
End Sub

'todo remove--------------------
Private Sub Command1_Click()
    fcsh.exec Text1.Text
End Sub

