VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{95D85F43-414D-432F-909E-2ED57BBC389C}#1.2#0"; "MCLHotkey.ocx"
Begin VB.Form MainForm 
   AutoRedraw      =   -1  'True
   Caption         =   "Flex Compiler SHell Wrapper"
   ClientHeight    =   6450
   ClientLeft      =   3840
   ClientTop       =   2220
   ClientWidth     =   10590
   Icon            =   "Main.frx":0000
   ScaleHeight     =   6450
   ScaleWidth      =   10590
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MCLHotkey.VBHotKey HotKey 
      Left            =   1200
      Top             =   5280
      _ExtentX        =   794
      _ExtentY        =   794
      VKey            =   1
      WinKey          =   0   'False
      Enabled         =   0   'False
   End
   Begin MSComctlLib.ImageList fakeList 
      Left            =   5640
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1CFA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList pngImages 
      Left            =   5640
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
   End
   Begin VB.PictureBox picClear 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   5280
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   4
      Top             =   3480
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox picIconLoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   4920
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   3
      Top             =   3480
      Visible         =   0   'False
      Width           =   270
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
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10590
      _ExtentX        =   18680
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "fakeList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "run"
            Object.ToolTipText     =   "Start fcsh"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Clear log"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "build"
            Object.ToolTipText     =   "Build"
            ImageIndex      =   1
            Style           =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "type"
            Object.ToolTipText     =   "Incremental build"
            ImageIndex      =   1
            Style           =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "info"
            Object.ToolTipText     =   "Show target info"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "options"
            Object.ToolTipText     =   "Options"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ontop"
            Object.ToolTipText     =   "Place window on top"
            ImageIndex      =   1
            Style           =   1
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "transparent"
            Object.ToolTipText     =   "Make window transparent"
            ImageIndex      =   1
            Style           =   1
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "about"
            Object.ToolTipText     =   "About"
            ImageIndex      =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
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
      TextRTF         =   $"Main.frx":3A04
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
      Begin VB.Menu mnuFloat 
         Caption         =   "Show floating window"
      End
      Begin VB.Menu mnuSep3 
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

Private Const ABOUT_BUTTON As Long = 13
Private Const TRANSPARENT_BUTTON As Long = 11
Private Const ONTOP_BUTTON As Long = 10
Private Const CLEAR_BUTTON As Long = 2
Private Const OPTIONS_BUTTON As Long = 8
Private Const INFO_BUTTON As Long = 6
Private Const TYPE_BUTTON As Long = 5
Private Const BUILD_BUTTON As Long = 4
Private Const RUN_BUTTON As Long = 1

Private Const BUILD_FAILED As String = "Build failed"
Private Const BUILD_SUCESSFULL As String = "Build successfull"

Private log As New clsLog

Public config As New clsConfiguration
Public preloader As New clsResource

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
    
    frmFloat.error
    
    
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
    
    frmFloat.idle
    
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
    
    frmFloat.idle
    
    sendRemote BUILD_SUCESSFULL
End Sub

'on fcsh.exe start
Private Sub fcsh_onStart()
   Toolbar.Buttons.item(RUN_BUTTON).Image = preloader.getIndex(STOP_FCSH)
   Toolbar.Buttons.item(RUN_BUTTON).ToolTipText = "Stop fcsh"
   'Toolbar.Buttons.item(RUN_BUTTON).Caption = "Stop fcsh"
   targets.RemoveAll
   Set lastTarget = Nothing
   DisplayBalloon "Flex compiler shell", "fcsh is started", NIIF_INFO
   frmFloat.idle
End Sub

'on fcsh.exe stop
Private Sub fcsh_onStop()
   Toolbar.Buttons.item(RUN_BUTTON).Image = preloader.getIndex(START_FCSH)
   Toolbar.Buttons.item(RUN_BUTTON).ToolTipText = "Start fcsh"
   'Toolbar.Buttons.item(RUN_BUTTON).Caption = "Start fcsh"
   targets.RemoveAll
   Set lastTarget = Nothing
   log.Text vbCrLf
   DisplayBalloon "Flex compiler shell", "fcsh is stopped", NIIF_WARNING
   frmFloat.stopped
End Sub

'***********************************************************************************************
'Start app
'***********************************************************************************************


'start application
Private Sub Form_Load()
        Me.icon = frmAbout.icon

        'init prefs
        config.Load
        
        'set up logging
        log.clsLog rtbLog, config

        'load png images
        LoadPNG
          
        'init vars
        isServerBusy = False
        initSockets 'listen for requests
        
        'setup fcsh
        Set fcsh = New clsFCSH
        fcsh.Initialize log, config
            
        'add tray icon
        TrayAdd fakeTray.hWnd, Me.icon, "Flex compiler shell", MouseMove
        
        'log and show tooltip
        log.xDebug "Application initialized"
        
        'load configured apps
        loadApps
        
        'hotkeys
        Dim hotkeySetup As New clsHotKeySetup
        hotkeySetup.SetupKey config.RECOMPILE, HotKey
End Sub

Private Sub LoadPNG()
    'extract files and save
    preloader.getResourceByName START_FCSH
    preloader.getResourceByName STOP_FCSH
    preloader.getResourceByName LOG_CLEAR
    preloader.getResourceByName BUILD_TASK
    preloader.getResourceByName INCREMENTAL_ON
    preloader.getResourceByName INCREMENTAL_OFF
    preloader.getResourceByName TARGET_INFO
    preloader.getResourceByName OPTIONS
    preloader.getResourceByName ON_TOP
    preloader.getResourceByName Alpha
    preloader.getResourceByName ABOUT
    preloader.getResourceByName APP_APPEARANCE
    preloader.getResourceByName KEYBOARD
    
    preloader.getResourceByName IDLE_PNG
    preloader.getResourceByName STOPPED_PNG
    preloader.getResourceByName ERROR_PNG
    preloader.getResourceByName EXEC_PNG, ".gif"
    
    'load
    Dim pngLoader As New clsPngToImageList
    pngLoader.Initialize picIconLoad, picClear, pngImages, log
    pngLoader.LoadIcons preloader.extractedFiles
    
    
    'setup toolbar
    Set Toolbar.ImageList = pngImages
    Toolbar.Buttons(RUN_BUTTON).Image = preloader.getIndex(START_FCSH)
    Toolbar.Buttons(BUILD_BUTTON).Image = preloader.getIndex(BUILD_TASK)
    Toolbar.Buttons(TYPE_BUTTON).Image = preloader.getIndex(INCREMENTAL_ON)
    Toolbar.Buttons(INFO_BUTTON).Image = preloader.getIndex(TARGET_INFO)
    Toolbar.Buttons(OPTIONS_BUTTON).Image = preloader.getIndex(OPTIONS)
    Toolbar.Buttons(CLEAR_BUTTON).Image = preloader.getIndex(LOG_CLEAR)
    Toolbar.Buttons(ONTOP_BUTTON).Image = preloader.getIndex(ON_TOP)
    Toolbar.Buttons(TRANSPARENT_BUTTON).Image = preloader.getIndex(Alpha)
    Toolbar.Buttons(ABOUT_BUTTON).Image = preloader.getIndex(ABOUT)
End Sub

Public Sub loadApps()
    Toolbar.Buttons(BUILD_BUTTON).ButtonMenus.clear
    frmFloat.Toolbar1.Buttons(1).ButtonMenus.clear
    Dim i As Long
    Dim app As clsTarget
    Dim KEY As String
    Dim ButtonMenu As MSComctlLib.ButtonMenu
    
    For i = 1 To config.APPLICATIONS
        Set app = config.LoadApplication(i)
        KEY = i & "app"
        Toolbar.Buttons(BUILD_BUTTON).ButtonMenus.Add i, KEY, app.fName
        frmFloat.Toolbar1.Buttons(1).ButtonMenus.Add i, KEY, app.fName
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
        Err.clear
    End If
End Sub


'hot key build
Private Sub HotKey_HotkeyPressed()
    If (config.RECOMPILE.ENABLED = 1) Then
        rebuild
    End If
End Sub

Private Sub mnuExit_Click()
    Form_Unload 0
    End
End Sub

Private Sub mnuFloat_Click()
    mnuFloat.Checked = Not mnuFloat.Checked
    If (mnuFloat.Checked = True) Then
        frmFloat.Show
        SetAlwaysOnTopMode frmFloat.hWnd, True
        Dim bytOpacity As Byte
        'Set the transparency level
        bytOpacity = config.FLOATALPHA
        Call SetWindowLong(frmFloat.hWnd, GWL_EXSTYLE, GetWindowLong(frmFloat.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
        Call SetLayeredWindowAttributes(frmFloat.hWnd, 0, bytOpacity, LWA_ALPHA)
    Else
        frmFloat.Hide
    End If
End Sub

Private Sub mnuRecompile_Click()
   rebuild
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
        BUILD i
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
        Case RUN_BUTTON:
                If (Not fcsh.isRunning) Then
                    fcsh.Start
                Else
                    fcsh.Quit
                End If
        Case BUILD_BUTTON:
                rebuild
        Case TYPE_BUTTON:
                If (Toolbar.Buttons(TYPE_BUTTON).Value = tbrUnpressed) Then
                    Toolbar.Buttons(TYPE_BUTTON).Image = preloader.getIndex(INCREMENTAL_ON)
                Else
                    Toolbar.Buttons(TYPE_BUTTON).Image = preloader.getIndex(INCREMENTAL_OFF)
                End If
        Case INFO_BUTTON:
                If ((lastTarget Is Nothing)) Then
                    log.xError "No targets were assigned yet. Nothing to show."
                    log.Text vbCrLf
                    Exit Sub
                End If
                If (lastTarget.fTargetID = 0) Then
                    log.xError "No targets were assigned yet. Nothing to show."
                    log.Text vbCrLf
                    Exit Sub
                End If
                fcsh.info lastTarget
                
        Case OPTIONS_BUTTON:
                frmOptions.loadPrefs config, log
                SetAlwaysOnTopMode frmOptions.hWnd, (Toolbar.Buttons(ONTOP_BUTTON).Value = tbrPressed)
                frmOptions.Move (Screen.Width - frmOptions.Width) \ 2, ((Screen.Height - frmOptions.Height) \ 2)
        Case CLEAR_BUTTON:
                log.clear
        Case ONTOP_BUTTON:
                If (Toolbar.Buttons(ONTOP_BUTTON).Value = tbrPressed) Then
                    SetAlwaysOnTopMode Me.hWnd, True
                Else
                    SetAlwaysOnTopMode Me.hWnd, False
                End If
        Case TRANSPARENT_BUTTON:
                If (Toolbar.Buttons(TRANSPARENT_BUTTON).Value = tbrPressed) Then
                     Dim bytOpacity As Byte
                     'Set the transparency level
                     bytOpacity = config.Alpha
                     Call SetWindowLong(Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
                     Call SetLayeredWindowAttributes(Me.hWnd, 0, bytOpacity, LWA_ALPHA)
                Else
                    Call SetWindowLong(Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) And (Not WS_EX_LAYERED))
                End If
        Case ABOUT_BUTTON:
                SetAlwaysOnTopMode frmAbout.hWnd, (Toolbar.Buttons(ONTOP_BUTTON).Value = tbrPressed)
                frmAbout.Move (Screen.Width - frmAbout.Width) \ 2, ((Screen.Height - frmAbout.Height) \ 2)
    End Select
End Sub

Private Sub Toolbar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim index As Long
    index = Val(ButtonMenu.KEY)
    BUILD index
End Sub

Public Sub rebuild()
    Dim target As New clsTarget
    If ((lastTarget Is Nothing)) Then
       If (fcsh.isRunning) Then
           target.fMessage = "No targets were assigned yet. Nothing to recompile."
       Else
           target.fMessage = "Cant exec: fcsh stopped"
       End If
       fcsh_onError target
       Exit Sub
    End If
    If (lastTarget.fTargetID = 0) Then
       target.fMessage = "No targets were assigned yet. Nothing to recompile."
       fcsh_onError target
    End If
    fcsh.exec lastTarget, (Toolbar.Buttons(TYPE_BUTTON).Value = tbrUnpressed)
End Sub

Public Sub BUILD(index As Long)
    Dim app As clsTarget
    Set app = config.LoadApplication(index)
        
    If (fcsh.isRunning) Then
        frmFloat.active
        If (targets.Exists(app.fName)) Then
            Set lastTarget = targets.item(app.fName)
            fcsh.exec lastTarget, (Toolbar.Buttons(TYPE_BUTTON).Value = tbrUnpressed)
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
Private Sub Form_Unload(Cancel As Integer)
    preloader.clear
    log.xDebug "Application stopped"
    TrayDelete
    Server.Close
    Unload frmOptions
    Unload frmAbout
    Unload frmFloat
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Me.Visible = False
    lastState = Me.WindowState
    Me.WindowState = vbMinimized
    Cancel = 1
End Sub



