VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form MainForm 
   Caption         =   "Flex Compiler SHell Server"
   ClientHeight    =   6450
   ClientLeft      =   4290
   ClientTop       =   3600
   ClientWidth     =   7695
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
   ScaleHeight     =   6450
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkOnTop 
      Caption         =   "Always on top"
      Height          =   360
      Left            =   5880
      TabIndex        =   9
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton cmdClearLog 
      Caption         =   "Clear log"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   8
      Top             =   2880
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox rtbLog 
      Height          =   2415
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   4260
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Main.frx":058A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdHide 
      Cancel          =   -1  'True
      Caption         =   "Hide"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   5
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdRecompile 
      Caption         =   "Recompile"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   6000
      Width           =   1095
   End
   Begin VB.ListBox lstTargets 
      Height          =   2460
      ItemData        =   "Main.frx":0626
      Left            =   120
      List            =   "Main.frx":0628
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   3360
      Width           =   7455
   End
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
   Begin VB.Label Label2 
      Caption         =   "Flex Compiler SHell output:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   75
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Compiler cache:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   1575
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
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal _
    hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long
Const LB_SETHORIZONTALEXTENT = &H194
Const LB_GETHORIZONTALEXTENT = &H193




Private WithEvents fcsh As clsFCSH
Attribute fcsh.VB_VarHelpID = -1
Private prefs As New clsPreferences
Private log As New clsLog


Private Sub chkOnTop_Click()
    If (chkOnTop.value = 1) Then
        SetAlwaysOnTopMode Me.hwnd, True
    Else
        SetAlwaysOnTopMode Me.hwnd, False
    End If
End Sub

Private Sub cmdClear_Click()
    Dim KEY As String
    If (lstTargets.ListIndex <> -1) Then
        KEY = lstTargets.List(lstTargets.ListIndex)
        If (fcsh.isRunning And Not fcsh.isExec) Then
            fcsh.clear "clear " + CStr(fcsh.targets.Item(KEY))
            fcsh.targets.Remove KEY
        End If
    End If
    fillView
End Sub

Private Sub cmdClearLog_Click()
    rtbLog.Text = ""
End Sub

Private Sub cmdHide_Click()
    Me.Hide
End Sub

Private Sub cmdRecompile_Click()
    Dim KEY As String
    
    If (lstTargets.ListIndex <> -1) Then
        KEY = lstTargets.List(lstTargets.ListIndex)
        If (fcsh.isRunning And Not fcsh.isExec) Then
            fcsh.exec KEY
        End If
    End If
End Sub


Private Sub fcsh_CommandsEnabled(enable As Boolean)
    cmdClear.Enabled = enable
    cmdRecompile.Enabled = enable
End Sub



Private Sub Form_Load()
    log.setWindow rtbLog
    log.xInfo "Application started"
    prefs.initialize log
    
    Dim port As Long
    port = prefs.SERVER_PORT

    log.xInfo "Server is listening on port " & port
    Server.Close
    Server.LocalPort = port
    
    On Error Resume Next
    Server.Listen
    If Err.Number <> 0 Then
        log.xError "Cant start server: " + Err.description
        Err.clear
    End If
    
    TrayAdd fakeTray.hwnd, Me.Icon, "Flex Compiler SHell Server", MouseMove
    
   
    Set fcsh = New clsFCSH
    fcsh.initialize log, prefs
    
    fcsh.Start
    
    SetHorizontalExtent
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
    SendMessage lstTargets.hwnd, LB_SETHORIZONTALEXTENT, maxWidth, ByVal 0&
End Sub



Public Sub fillView()
    Dim KEY As Variant
    
    lstTargets.clear
    For Each KEY In fcsh.targets
        lstTargets.AddItem CStr(KEY)
    Next
    SetHorizontalExtent
End Sub


Private Sub Form_Resize()
    lstTargets.Width = Me.Width - lstTargets.Left * 2 - 150
    rtbLog.Width = Me.Width - rtbLog.Left * 2 - 150
    
    Dim listHeight As Long
    listHeight = Me.Height - lstTargets.Top - 1000
    If (listHeight > 0) Then
        lstTargets.Height = listHeight
    End If
    
   
    listHeight = lstTargets.Top + lstTargets.Height + 100
    cmdClear.Top = listHeight
    cmdHide.Top = listHeight
    cmdRecompile.Top = listHeight
    cmdClearLog.Left = lstTargets.Left + lstTargets.Width - cmdClearLog.Width
    chkOnTop.Left = lstTargets.Left + lstTargets.Width - chkOnTop.Width
    cmdHide.Left = lstTargets.Left + lstTargets.Width - cmdHide.Width
End Sub


Private Sub mnu_log_Click()
    Shell "notepad.exe " + App.path + "\FCSHServer.log", vbNormalFocus
End Sub

Private Sub mnu_show_window_Click()
    Me.Show
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
    sendDataVO value
End Sub

Private Sub fcsh_onStart(value As DataVO)
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
    Dim result As String
    result = "Flex Compiler SHell Server. Version " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & vbCrLf
    If (Server.State = sckListening) Then
        result = result + "Server is listening on port " & Server.LocalPort & vbCrLf & vbCrLf
    Else
        result = result + "Warning: Server socket state " & Server.State & vbCrLf & vbCrLf
    End If
    
    result = result + "Is client connected: " & (Service.State = sckConnected) & vbCrLf & vbCrLf

    result = result + "Is fcsh running: " & (fcsh.isRunning)
    
    MsgBox result, vbOKOnly, "About"
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
