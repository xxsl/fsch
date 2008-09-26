VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MainForm 
   AutoRedraw      =   -1  'True
   Caption         =   "Flex compiler shell"
   ClientHeight    =   6450
   ClientLeft      =   3840
   ClientTop       =   2220
   ClientWidth     =   13785
   Icon            =   "Main.frx":0000
   ScaleHeight     =   6450
   ScaleWidth      =   13785
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSComctlLib.ImageList disabledIcons 
      Left            =   5520
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":08DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0C2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0F80
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":12D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":16A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":19F6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList enabledIcons 
      Left            =   5520
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1D48
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":209A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":23EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":273E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":2A90
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":2E62
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":31B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":3506
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":3858
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton fakeTray 
      Caption         =   "fakeTray"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSWinsockLib.Winsock Abort 
      Left            =   4920
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Service 
      Left            =   4920
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13785
      _ExtentX        =   24315
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "enabledIcons"
      DisabledImageList=   "disabledIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Start fcsh"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Build"
            ImageIndex      =   3
            Style           =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Incremental build"
            ImageIndex      =   7
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Show target info"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Options"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Clear log"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Place window on top"
            ImageIndex      =   8
            Style           =   1
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Make window transparent"
            ImageIndex      =   9
            Style           =   1
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtbLog 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   6376
      _Version        =   393217
      BackColor       =   -2147483633
      BorderStyle     =   0
      ScrollBars      =   3
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Main.frx":3BAA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSWinsockLib.Winsock Server 
      Left            =   4920
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   44000
   End
   Begin VB.Menu mnuShell 
      Caption         =   "Shell"
      Visible         =   0   'False
      Begin VB.Menu mnuRecompile 
         Caption         =   "Recompile"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRun 
         Caption         =   "Run/Stop fcsh"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "MainForm"
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

Private Const BUILD_BUTTON As Long = 2
Private Const RUN_BUTTON As Long = 1
Private Const BUILD_FAILED As String = "Build failed"
Private Const BUILD_SUCESSFULL As String = "Build successfull"

Private log As New clsLog
Public config As New clsConfiguration
Dim WithEvents fcsh As clsFCSH
Attribute fcsh.VB_VarHelpID = -1

Private targets As New Dictionary
Private lastTarget As clsTarget

Private isRemote As Boolean
Private isServerBusy As Boolean
Private responce As String

Private lastState As Long




'***********************************************************************************************
'fcsh event handling
'***********************************************************************************************
Private Sub fcsh_onError(target As clsTarget)
    Dim Msg As String
    Msg = target.fMessage
    
    log.xError Msg
    log.Text vbCrLf
    DisplayBalloon "Flex compiler shell", BUILD_FAILED + ". " + Msg, NIIF_ERROR
    
    If ((target.fTargetID = 0) And (targets.Exists(target.fName))) Then
        targets.Remove target.fName
    End If
    
    sendRemote Msg + vbCrLf + BUILD_FAILED
End Sub

'on command success
Private Sub fcsh_onFinish(target As clsTarget)
    log.xFcsh "Exec completed"
    log.Text vbCrLf
    DisplayBalloon "Flex compiler shell", BUILD_SUCESSFULL, NIIF_INFO
    
    If ((target.fTargetID = 0) And (targets.Exists(target.fName))) Then
        targets.Remove target.fName
    End If
    
    sendRemote BUILD_SUCESSFULL
End Sub

'on new compile target id
Private Sub fcsh_onIdAssigned(target As clsTarget)
    log.xFcsh "Exec completed. id is " & target.fTargetID
    log.Text vbCrLf
    DisplayBalloon "Flex compiler shell", BUILD_SUCESSFULL + ". Assigned Target id " & target.fTargetID, NIIF_INFO
    
    If (Not targets.Exists(target.fName)) Then
        targets.Add target.fName, target
    End If
    
    sendRemote BUILD_SUCESSFULL
End Sub

'on fcsh.exe start
Private Sub fcsh_onStart()
   Toolbar.Buttons.item(RUN_BUTTON).Image = 2
   Toolbar.Buttons.item(RUN_BUTTON).ToolTipText = "Stop fcsh"
   targets.RemoveAll
   Set lastTarget = Nothing
   DisplayBalloon "Flex compiler shell", "fcsh is started", NIIF_INFO
End Sub

'on fcsh.exe stop
Private Sub fcsh_onStop()
   Toolbar.Buttons.item(RUN_BUTTON).Image = 1
   Toolbar.Buttons.item(RUN_BUTTON).ToolTipText = "Start fcsh"
   targets.RemoveAll
   Set lastTarget = Nothing
   log.Text vbCrLf
   DisplayBalloon "Flex compiler shell", "fcsh is stopped", NIIF_WARNING
End Sub

'***********************************************************************************************
'Start app
'***********************************************************************************************


'start application
Private Sub Form_Load()
        'init prefs
        config.Load
        
        'set up logging
        log.clsLog rtbLog, config

          
        'init vars
        isServerBusy = False
        initSockets 'listen for requests
        
        'setup fcsh
        Set fcsh = New clsFCSH
        fcsh.Initialize log, config
            
        'add tray icon
        TrayAdd fakeTray.hWnd, Me.Icon, "Flex compiler shell", MouseMove
        
        'log and show tooltip
        log.xDebug "Application initialized"
        
        'load configured apps
        loadApps
        
End Sub

Public Sub loadApps()
    Toolbar.Buttons(BUILD_BUTTON).ButtonMenus.Clear
    Dim i As Long
    Dim app As clsTarget
    Dim key As String
    Dim ButtonMenu As MSComctlLib.ButtonMenu
    
    For i = 1 To config.APPLICATIONS
        Set app = config.LoadApplication(i)
        key = i & "app"
        Toolbar.Buttons(BUILD_BUTTON).ButtonMenus.Add i, key, app.fName
    Next i
End Sub


'***********************************************************************************************
'Sockets Server
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




Private Sub mnuExit_Click()
    Form_QueryUnload 0, 0
    End
End Sub

Private Sub mnuRecompile_Click()
   If (Not (lastTarget Is Nothing)) Then
       fcsh.exec lastTarget
   Else
       log.xError "No targets were assigned yet. Nothing to recompile."
       log.Text vbCrLf
       DisplayBalloon "Flex compiler shell", "No targets were assigned yet. Nothing to recompile.", NIIF_WARNING
   End If
End Sub

Private Sub mnuRun_Click()
If (fcsh.isRunning) Then
    fcsh.Quit
Else
    fcsh.Start
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
    isRemote = False
End Sub

Private Sub Service_Connect()
    isRemote = True
End Sub

Private Sub Service_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    log.xError "Service socket: " + Description
    Service.Close
    isRemote = False
End Sub

Private Sub Service_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    Dim intCnt As Integer

    Service.GetData strData, vbString
    
    log.xDebug "Socket in: " + strData
    
    isRemote = True
    
    remoteExec strData
End Sub


Private Sub remoteExec(arg As String)
    Dim i As Long
    Dim app As clsTarget
    Dim appFound As Boolean
    
    For i = 1 To config.APPLICATIONS
        Set app = config.LoadApplication(i)
        If (LCase(app.fName) = LCase(arg)) Then
            appFound = True
            Exit For
        End If
    Next i
    
    If (Not fcsh.isRunning) Then
        sendRemote "fsch is stopped" + vbCrLf + BUILD_FAILED
        Exit Sub
    End If
    
    If (fcsh.isExec) Then
        sendRemote "fsch is busy" + vbCrLf + BUILD_FAILED
        Exit Sub
    End If
    
    If (Not appFound) Then
        sendRemote "Application not found [" + arg + "]" + vbCrLf + BUILD_FAILED
    Else
        build i
    End If
End Sub

Private Sub sendRemote(Msg As String)
   If (isRemote) Then
       Service.SendData Msg
   End If
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
        Case BUILD_BUTTON:
                If ((lastTarget Is Nothing)) Then
                    log.xError "No targets were assigned yet. Nothing to recompile."
                    log.Text vbCrLf
                    Exit Sub
                End If
                If (lastTarget.fTargetID = 0) Then
                    log.xError "No targets were assigned yet. Nothing to recompile."
                    log.Text vbCrLf
                    Exit Sub
                End If
                fcsh.exec lastTarget, (Toolbar.Buttons(3).value = tbrPressed)
        Case 4:
                If ((lastTarget Is Nothing)) Then
                    log.xError "No targets were assigned yet. Nothing to recompile."
                    log.Text vbCrLf
                    Exit Sub
                End If
                If (lastTarget.fTargetID = 0) Then
                    log.xError "No targets were assigned yet. Nothing to recompile."
                    log.Text vbCrLf
                    Exit Sub
                End If
                fcsh.info lastTarget
                
        Case 6:
                frmOptions.loadPrefs config, log
                frmOptions.Show 1, Me
        Case 8:
                log.Clear
        Case 10:
                If (Toolbar.Buttons(10).value = tbrPressed) Then
                    SetAlwaysOnTopMode Me.hWnd, True
                Else
                    SetAlwaysOnTopMode Me.hWnd, False
                End If
        Case 11:
                If (Toolbar.Buttons(11).value = tbrPressed) Then
                     Dim bytOpacity As Byte
                     'Set the transparency level
                     bytOpacity = config.ALPHA
                     Call SetWindowLong(Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
                     Call SetLayeredWindowAttributes(Me.hWnd, 0, bytOpacity, LWA_ALPHA)
                Else
                    Call SetWindowLong(Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) And (Not WS_EX_LAYERED))
                End If
    End Select
End Sub

Private Sub Toolbar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim index As Long
    index = Val(ButtonMenu.key)
    build index
End Sub

Private Sub build(index As Long)
    Dim app As clsTarget
    Set app = config.LoadApplication(index)
        
    If (fcsh.isRunning) Then
        If (targets.Exists(app.fName)) Then
            Set lastTarget = targets.item(app.fName)
            fcsh.exec lastTarget, (Toolbar.Buttons(3).value = tbrPressed)
        Else
            Set lastTarget = app
            targets.Add app.fName, app
            fcsh.exec app
        End If
    Else
        app.fMessage = "Cant exec: fcsh stopped"
        fcsh_onError app
    End If
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
            If (Me.Visible) Then
                Me.Visible = False
                lastState = Me.WindowState
                Me.WindowState = vbMinimized
            Else
                Me.Visible = True
                Me.WindowState = lastState
            End If
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
            log.xDebug "RightUp": PopupMenu mnuShell
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
    logWidth = Me.Height - rtbLog.Top - 500
    If (logWidth > 0) Then
        rtbLog.Height = logWidth
    End If
End Sub

'on quit
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    log.xDebug "Application stopped"
    TrayDelete
    Server.Close
    Unload frmOptions
End Sub



