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
   ClientWidth     =   10590
   Icon            =   "Main.frx":0000
   ScaleHeight     =   6450
   ScaleWidth      =   10590
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
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
            Picture         =   "Main.frx":058A
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
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10590
      _ExtentX        =   18680
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "fakeList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Start fcsh"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Build"
            ImageIndex      =   1
            Style           =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Incremental build"
            ImageIndex      =   1
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Show target info"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Options"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Clear log"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Place window on top"
            ImageIndex      =   1
            Style           =   1
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Make window transparent"
            ImageIndex      =   1
            Style           =   1
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "About"
            ImageIndex      =   1
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
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Main.frx":2294
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

Private Const ABOUT_BUTTON As Long = 14
Private Const TRANSPARENT_BUTTON As Long = 12
Private Const ONTOP_BUTTON As Long = 11
Private Const CLEAR_BUTTON As Long = 9
Private Const OPTIONS_BUTTON As Long = 7
Private Const INFO_BUTTON As Long = 5
Private Const TYPE_BUTTON As Long = 4
Private Const BUILD_BUTTON As Long = 3
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
        TrayAdd fakeTray.hWnd, Me.icon, "Flex compiler shell", MouseMove
        
        'log and show tooltip
        log.xDebug "Application initialized"
        
        'load configured apps
        loadApps
        
        'load png images
        LoadPNG
        
End Sub

Private Sub LoadPNG()
    'extract files and save
    Dim files As New Collection
    Dim imgArray() As Byte
    Dim I As Long
    
    For I = 101 To 111
        imgArray = LoadResData(I, "custom")
        
        Open I & ".png" For Output As #2
        Close #2
        
        Open I & ".png" For Binary As #2
           Put #2, , imgArray()
        Close #2
        files.Add I & ".png"
    Next I
    
    'load
    Dim pngLoader As New clsPngToImageList
    pngLoader.Initialize picIconLoad, picClear, pngImages, log
    pngLoader.LoadIcons files
    
    For I = 101 To 111
        If (FileExists(I & ".png")) Then
            Kill I & ".png"
        Else
            log.xDebug "File not found " & I & ".png"
        End If
    Next I
    
    'setup toolbar
    Set Toolbar.ImageList = pngImages
    Toolbar.Buttons(RUN_BUTTON).Image = 1
    Toolbar.Buttons(BUILD_BUTTON).Image = 3
    Toolbar.Buttons(TYPE_BUTTON).Image = 4
    Toolbar.Buttons(INFO_BUTTON).Image = 6
    Toolbar.Buttons(OPTIONS_BUTTON).Image = 7
    Toolbar.Buttons(CLEAR_BUTTON).Image = 8
    Toolbar.Buttons(ONTOP_BUTTON).Image = 9
    Toolbar.Buttons(TRANSPARENT_BUTTON).Image = 10
    Toolbar.Buttons(ABOUT_BUTTON).Image = 11
End Sub

Public Sub loadApps()
    Toolbar.Buttons(BUILD_BUTTON).ButtonMenus.clear
    Dim I As Long
    Dim app As clsTarget
    Dim key As String
    Dim ButtonMenu As MSComctlLib.ButtonMenu
    
    For I = 1 To config.APPLICATIONS
        Set app = config.LoadApplication(I)
        key = I & "app"
        Toolbar.Buttons(BUILD_BUTTON).ButtonMenus.Add I, key, app.fName
    Next I
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
    Dim I As Long
    Dim app As clsTarget
    Dim appFound As Boolean
    
    For I = 1 To config.APPLICATIONS
        Set app = config.LoadApplication(I)
        If (LCase(app.fName) = LCase(arg)) Then
            appFound = True
            Exit For
        End If
    Next I
    
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
        build I
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
    Select Case Button.Index
        Case RUN_BUTTON:
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
                fcsh.exec lastTarget, (Toolbar.Buttons(TYPE_BUTTON).Value = tbrPressed)
        Case TYPE_BUTTON:
                If (Toolbar.Buttons(TYPE_BUTTON).Value = tbrPressed) Then
                    Toolbar.Buttons(TYPE_BUTTON).Image = 4
                Else
                    Toolbar.Buttons(TYPE_BUTTON).Image = 5
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
                frmOptions.Show 1, Me
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
                frmAbout.Show 1
    End Select
End Sub

Private Sub Toolbar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim Index As Long
    Index = Val(ButtonMenu.key)
    build Index
End Sub

Private Sub build(Index As Long)
    Dim app As clsTarget
    Set app = config.LoadApplication(Index)
        
    If (fcsh.isRunning) Then
        If (targets.Exists(app.fName)) Then
            Set lastTarget = targets.item(app.fName)
            fcsh.exec lastTarget, (Toolbar.Buttons(TYPE_BUTTON).Value = tbrPressed)
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
    Unload frmAbout
End Sub



