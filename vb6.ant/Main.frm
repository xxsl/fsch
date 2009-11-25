VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form MainForm 
   BackColor       =   &H8000000C&
   Caption         =   "Flex Compiler SHell Server"
   ClientHeight    =   7560
   ClientLeft      =   4290
   ClientTop       =   3675
   ClientWidth     =   10260
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   10260
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton fakeTray 
      Caption         =   "fakeTray"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   8040
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSWinsockLib.Winsock Service 
      Left            =   1920
      Top             =   8520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Server 
      Left            =   1920
      Top             =   8040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox frmTargets 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   120
      ScaleHeight     =   3615
      ScaleWidth      =   9975
      TabIndex        =   3
      Top             =   3960
      Width           =   9975
      Begin VB.CommandButton cmdHide 
         Cancel          =   -1  'True
         Caption         =   "Hide"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6360
         TabIndex        =   7
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton cmdRecompile 
         Caption         =   "Recompile"
         Default         =   -1  'True
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   3120
         Width           =   1215
      End
      Begin VB.ListBox lstTargets 
         Height          =   2460
         IntegralHeight  =   0   'False
         ItemData        =   "Main.frx":058A
         Left            =   120
         List            =   "Main.frx":058C
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   7455
      End
      Begin VB.Label Label1 
         Caption         =   "Compiler cache:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.PictureBox frmFcsh 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   120
      ScaleHeight     =   3615
      ScaleWidth      =   9975
      TabIndex        =   2
      Top             =   120
      Width           =   9975
      Begin TabDlg.SSTab SSTab 
         Height          =   3375
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   5953
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   617
         WordWrap        =   0   'False
         ShowFocusRect   =   0   'False
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "FCSH Output"
         TabPicture(0)   =   "Main.frx":058E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "cmdClearLog"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "rtbLog"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Build Errors"
         TabPicture(1)   =   "Main.frx":069F
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "rtbError"
         Tab(1).Control(1)=   "cmdClearErr"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Build Warnings"
         TabPicture(2)   =   "Main.frx":07AF
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "rtbWarn"
         Tab(2).Control(1)=   "cmdClearWarn"
         Tab(2).ControlCount=   2
         TabCaption(3)   =   "Preferences"
         TabPicture(3)   =   "Main.frx":08BF
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "chkOnTop"
         Tab(3).ControlCount=   1
         Begin RichTextLib.RichTextBox rtbWarn 
            Height          =   2295
            Left            =   -74880
            TabIndex        =   16
            Top             =   480
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   4048
            _Version        =   393217
            BackColor       =   16777215
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            ScrollBars      =   3
            Appearance      =   0
            RightMargin     =   65000
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"Main.frx":0A05
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
         Begin RichTextLib.RichTextBox rtbError 
            Height          =   2295
            Left            =   -74880
            TabIndex        =   15
            Top             =   480
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   4048
            _Version        =   393217
            BackColor       =   16777215
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            ScrollBars      =   3
            Appearance      =   0
            RightMargin     =   65000
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"Main.frx":0A83
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
         Begin RichTextLib.RichTextBox rtbLog 
            Height          =   2295
            Left            =   120
            TabIndex        =   14
            Top             =   480
            Width           =   9375
            _ExtentX        =   16536
            _ExtentY        =   4048
            _Version        =   393217
            BackColor       =   16777215
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            ScrollBars      =   3
            Appearance      =   0
            RightMargin     =   65000
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"Main.frx":0B01
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
         Begin VB.CommandButton cmdClearWarn 
            Caption         =   "Clear log"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -66600
            TabIndex        =   13
            Top             =   2880
            Width           =   1215
         End
         Begin VB.CommandButton cmdClearLog 
            Caption         =   "Clear log"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8400
            TabIndex        =   12
            Top             =   2880
            Width           =   1215
         End
         Begin VB.CommandButton cmdClearErr 
            Caption         =   "Clear log"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -66600
            TabIndex        =   11
            Top             =   2880
            Width           =   1215
         End
         Begin VB.CheckBox chkOnTop 
            Caption         =   "Always on top"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   -74880
            TabIndex        =   10
            Top             =   480
            Width           =   1695
         End
      End
   End
   Begin FCSHServer.ctlSplitterEx ctlSplitterEx1 
      Height          =   7575
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   13361
   End
   Begin VB.Menu mnu_shell 
      Caption         =   "SHell"
      Visible         =   0   'False
      Begin VB.Menu mnu_about 
         Caption         =   "About"
      End
      Begin VB.Menu mnu_space3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_show_window 
         Caption         =   "Compiler cache"
      End
      Begin VB.Menu mnu_space1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_log 
         Caption         =   "View log"
      End
      Begin VB.Menu mnu_space2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_exit 
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

'-------------------------------
'Fix HorizontalScroll in ListBox
'-------------------------------
Const LB_SETHORIZONTALEXTENT = &H194
Const LB_GETHORIZONTALEXTENT = &H193


'---------------------------------
'Flex Compiler SHEll Wrapper class
'---------------------------------
Private WithEvents fcsh As clsFCSH
Attribute fcsh.VB_VarHelpID = -1

'-----------------------
'Application preferences
'-----------------------
Private prefs As New clsPreferences

'-----------------------
'Log to file and console
'-----------------------
Private log As New clsLog

'--------------
'resize handler
'--------------
Private resizeHandle As New clsResizeHandle


'-------------------------------------
'PREFERENCES TAB, set AlwaysOnTop mode
'-------------------------------------
Private Sub chkOnTop_Click()
    prefs.alwaysOnTop = (chkOnTop.value = 1)
    If (chkOnTop.value = 1) Then
        SetAlwaysOnTopMode Me.HWND, True
    Else
        SetAlwaysOnTopMode Me.HWND, False
    End If
End Sub



'-----------------
'Clear log windows
'-----------------

'Error window
Private Sub cmdClearErr_Click()
    rtbError.Text = ""
End Sub

'Main log window
Private Sub cmdClearLog_Click()
    rtbLog.Text = ""
End Sub

'Warnings window
Private Sub cmdClearWarn_Click()
    rtbWarn.Text = ""
End Sub


'--------------------------------------
'Clears selected target from FCSH cache
'--------------------------------------
Private Sub cmdClear_Click()
    Dim key As String, value As String
    
    If (lstTargets.ListIndex <> -1) Then
        key = lstTargets.List(lstTargets.ListIndex)
        value = CStr(fcsh.targets.Item(key))
        If (fcsh.isRunning And Not fcsh.isExec) Then
            fcsh.targets.Remove key
            fcsh.clear "clear " + value
            
            log.xInfo "Target cleared: (" & value & ") " & key
        Else
            log.xFcsh "[ERROR] Flex Compiler SHell is stopped or busy" & vbCrLf
        End If
    Else
        log.xFcsh "[ERROR] No target was selected for removal" & vbCrLf
    End If
    
    FillTargetsBox
End Sub

'-----------------
'Fills targets box
'-----------------
Public Sub FillTargetsBox()
    Dim key As Variant
    
    lstTargets.clear
    
    For Each key In fcsh.targets
        lstTargets.AddItem CStr(key)
    Next
    
    DestroyTooltip
    
    SetHorizontalExtent
End Sub



Private Sub cmdHide_Click()
    Me.Hide
End Sub


Private Sub cmdRecompile_Click()
    Dim key As String
    
    If (lstTargets.ListIndex <> -1) Then
        key = lstTargets.List(lstTargets.ListIndex)
        If (fcsh.isRunning And Not fcsh.isExec) Then
            fcsh.exec key
        Else
            log.xFcsh "[ERROR] Flex Compiler SHell is stopped or busy" & vbCrLf
        End If
    Else
        log.xFcsh "[ERROR] No target was selected for recompile" & vbCrLf
    End If
End Sub


Private Sub fcsh_CommandsEnabled(enable As Boolean)
    cmdClear.Enabled = enable
    cmdRecompile.Enabled = enable
End Sub



Private Sub Form_Load()
    Me.ctlSplitterEx1.AttachObjects Me.frmFcsh, Me.frmTargets, False
    Me.ctlSplitterEx1.TileMode = TILE_HORIZONTALLY
    
    log.setWindow rtbLog, rtbError, rtbWarn
    log.xInfo "Application started"
    
    prefs.initialize log
    
    Dim port As Long
    port = prefs.serverPort

    log.xInfo "Server is listening on port " & port
    Server.Close
    Server.LocalPort = port
    
   
    TrayAdd fakeTray.HWND, Me.Icon, "Flex Compiler SHell Server", MouseMove
    
   
    Set fcsh = New clsFCSH
    fcsh.initialize log, prefs
    
    fcsh.Start
    
    SetHorizontalExtent
    
    resizeHandle.setup Me
End Sub

Sub SetHorizontalExtent()
    Dim maxWidth As Long
    Dim Item As Long
    maxWidth = 0
    If (lstTargets.ListCount > 0) Then
        For Item = 0 To lstTargets.ListCount - 1
            If (maxWidth < TextWidth(CStr(lstTargets.List(Item)))) Then
                maxWidth = TextWidth(CStr(lstTargets.List(Item)))
            End If
        Next
    End If

    maxWidth = maxWidth / Screen.TwipsPerPixelX
    SendMessage lstTargets.HWND, LB_SETHORIZONTALEXTENT, maxWidth, ByVal 0&
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Result As VbMsgBoxResult
    
    If (Me.Visible) Then
        Result = MsgBox("This action will stop fcsh server. Are you sure?", vbExclamation + vbOKCancel, "Confirm")
        If (Result <> vbOK) Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub lstTargets_Click()
    Dim Result As String
    DestroyTooltip
    If (lstTargets.ListIndex <> -1) Then
        Dim lines() As String
        Dim i As Long
        Result = lstTargets.List(lstTargets.ListIndex)
        lines = Split(Result, " ")
        Result = ""
        For i = 0 To UBound(lines) - 1
            Result = Result & lines(i) & vbCrLf
        Next i
        DisplayTooltip lstTargets.HWND, Result, App.hInstance
    End If
End Sub


Private Sub mnu_log_Click()
    On Error Resume Next
    log.xInfo prefs.logViewer + " " + Chr(34) + App.path + "\FCSHServer.log" + Chr(34)
    
    Shell prefs.logViewer + " " + Chr(34) + App.path + "\FCSHServer.log" + Chr(34), vbNormalFocus
    If (Err.Number > 0) Then
        MsgBox Err.description, vbCritical, "Error"
        Err.clear
    End If
End Sub

Private Sub mnu_show_window_Click()
    If (prefs.alwaysOnTop) Then
        SetAlwaysOnTopMode Me.HWND, True
        chkOnTop.value = vbChecked
    Else
        Me.Show
        chkOnTop.value = vbUnchecked
    End If
End Sub



Private Sub Server_ConnectionRequest(ByVal requestID As Long)
    If (Not fcsh.isExec) Then
        log.xInfo "Accepted connection request " & requestID
        Service.Close
        Service.Accept requestID
    Else
        log.xInfo "Connection request ignored " & requestID
    End If
End Sub

Private Sub Server_Error(ByVal Number As Integer, description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    log.xError "Server socket error: " + description
End Sub

Private Sub Service_Connect()
    log.xInfo "Connection established"
End Sub

Private Sub Service_Close()
    log.xInfo "Service connection closed"
    'DisplayBalloon "Info", "Client disconnected", NIIF_INFO
End Sub

Private Sub Service_Error(ByVal Number As Integer, description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    log.xError "Service socket error: " + description
    Service.Close
End Sub

Private Sub Service_DataArrival(ByVal bytesTotal As Long)
       'wait for header int with length of the structure
       If (bytesTotal >= 4) Then
           'get structure length
           Dim objectLength As Long, i As Long, buffer() As Byte
           Service.PeekData buffer, , 4
           objectLength = convertInt(buffer)
           log.xDebug "Object length: " & objectLength & " bytesTotal: " & bytesTotal
            
           'if structure arrived completely then deserialize
           If ((bytesTotal - 4) = objectLength) Then
               Service.GetData buffer, , 4
               Dim objectData() As Byte
               Service.GetData objectData, vbArray
               
               deSerialize objectData
           End If
       End If
End Sub

Private Sub deSerialize(ByRef byteArray() As Byte)
    Dim objectType As String
    Dim pos As Long
    objectType = readString(byteArray, pos)
    
    log.xDebug "Object type: " & objectType
    
    Select Case objectType
        Case AIR_COMMANDVO:
                            Dim command As New CommandVO
                            command.deSerialize byteArray, pos
                            log.xDebug command.toString
                            executeCommand command
        Case AIR_ERRORVO:
                            Dim error As New ErrorVO
                            error.deSerialize byteArray, pos
                            log.xDebug error.toString
                            processError error
        Case AIR_DATAVO:
                            Dim data As New DataVO
                            data.deSerialize byteArray, pos
                            log.xDebug data.toString
                            processData data
        Case AIR_BALOONVO:
                            Dim baloon As New baloonVO
                            baloon.deSerialize byteArray, pos
                            log.xDebug baloon.toString
                            showBaloon baloon
    End Select
End Sub

Private Sub processData(data As DataVO)
    log.xError "Not implemented: Private Sub processData(data As DataVO)"
End Sub


Private Sub processError(error As ErrorVO)
    log.xError "Not implemented: Private Sub processError(error As ErrorVO)"
End Sub

Private Sub executeCommand(command As CommandVO)
    log.xDebug "Command: target=" + command.target + " command=" + command.command
    Select Case command.target
        Case "fcsh":
                    fcsh.exec command.command
        Case "fcsh_start":
                    fcsh.Start
        Case "fcsh_stop":
                    fcsh.Quit
        Case "fcsh_getstate":
                    If (fcsh.isRunning) Then
                        fcsh_calllback "true", "fcsh_getstate"
                    Else
                        fcsh_calllback "false", "fcsh_getstate"
                    End If
        Case "system_open":
                    Shell command.command, vbNormalFocus
    End Select
End Sub

Private Sub showBaloon(baloonVO As baloonVO)
    DisplayBalloon baloonVO.title, baloonVO.message, baloonVO.baloon_type
End Sub


Private Sub sendByteArray(ByRef byteArray() As Byte)
    If (Service.State = sckConnected) Then
        log.xDebug "Sending " & (UBound(byteArray) + 1) & " bytes"
        Dim size() As Byte
        ReDim size(0)
        writeLong size, UBound(byteArray) + 1
        Service.SendData size
        Service.SendData byteArray
    Else
        log.xDebug "Network falure. There are no clients connected to the server"
        'DisplayBalloon "Network falure", "There are no clients connected to the server", NIIF_ERROR
    End If
End Sub


Private Function convertInt(ByRef byteArray() As Byte) As Long
    Dim buffer(0 To 3) As Byte, i As Long
    For i = 3 To 0 Step -1
        buffer(3 - i) = byteArray(i)
    Next i
    CopyMemory convertInt, buffer(0), 4
End Function

Private Sub fcsh_onComplete(value As DataVO)
    sendDataVO value
End Sub

Private Sub fcsh_onStop(value As DataVO)
    cmdRecompile.Enabled = False
    cmdClear.Enabled = False
    sendDataVO value
End Sub

Private Sub fcsh_onStart(value As DataVO)
    On Error Resume Next
    Server.Listen
    If Err.Number <> 0 Then
        log.xError "Cant start server: " + Err.description
        Err.clear
    End If
    sendDataVO value
End Sub

Private Sub fcsh_onError(value As ErrorVO)
    log.xError value.toString
    Dim byteArray() As Byte
    ReDim byteArray(0)
    value.serialize byteArray
    sendByteArray byteArray
End Sub

Private Sub sendDataVO(data As DataVO)
    log.xDebug data.toString
    Dim byteArray() As Byte
    ReDim byteArray(0)
    data.serialize byteArray
    sendByteArray byteArray
End Sub


Private Sub fcsh_calllback(data As String, target As String)
    log.xDebug "Callback: target=" + target + " data=" + data
    Dim byteArray() As Byte
    ReDim byteArray(0)
    Dim dataObject As New DataVO
    dataObject.target = target
    dataObject.data = data
    dataObject.serialize byteArray
    sendByteArray byteArray
End Sub


Private Sub fakeTray_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cEvent As Single
    cEvent = X / Screen.TwipsPerPixelX
    Select Case cEvent
        Case MouseMove
            'Debug.Print "MouseMove"
        Case LeftUp
            'Debug.Print "Left Up"
            'showHide
        Case LeftDown
            'Debug.Print "Left Down"
        Case LeftDbClick
            'Debug.Print "LeftDbClick"
        Case MiddleUp
            'Debug.Print "MiddleUp"
        Case MiddleDown
            'Debug.Print "MiddleDown"
        Case MiddleDbClick
            'Debug.Print "MiddleDbClick"
        Case RightUp
            'Debug.Print "RightUp" ': PopupMenu mnuShell
            PopupMenu mnu_shell
        Case RightDown
            'Debug.Print "RightDown"
        Case RightDbClick
            'Debug.Print "RightDbClick"
    End Select
End Sub



Private Sub mnu_about_Click()
    Dim Result As String
    Result = "Flex Compiler SHell Server. Version " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & vbCrLf
    If (Server.State = sckListening) Then
        Result = Result + "Server is listening on port " & Server.LocalPort & vbCrLf & vbCrLf
    Else
        Result = Result + "Warning: Server socket state " & Server.State & vbCrLf & vbCrLf
    End If
    
    Result = Result + "Is client connected: " & (Service.State = sckConnected) & vbCrLf & vbCrLf

    Result = Result + "Is fcsh running: " & (fcsh.isRunning)
    
    MsgBox Result, vbOKOnly, "About"
End Sub

Private Sub mnu_exit_Click()
    Form_Unload 0
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    log.xInfo "Application stopped"
    Server.Close
    Service.Close
    TrayDelete
    fcsh.Quit
End Sub


Private Sub Service_SendComplete()
    log.xDebug "[send complete]"
End Sub

Private Sub sendCommand(data As String, target As String)
    log.xDebug "Send command: target=" + target + " data=" + data
    Dim byteArray() As Byte
    ReDim byteArray(0)
    Dim dataObject As New CommandVO
    dataObject.target = target
    dataObject.command = data
    dataObject.serialize byteArray
    sendByteArray byteArray
End Sub
