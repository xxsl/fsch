VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{95D85F43-414D-432F-909E-2ED57BBC389C}#1.2#0"; "MCLHotkey.ocx"
Begin VB.Form MainForm 
   Caption         =   "Form1"
   ClientHeight    =   1590
   ClientLeft      =   4290
   ClientTop       =   3600
   ClientWidth     =   3000
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1590
   ScaleWidth      =   3000
   Visible         =   0   'False
   Begin MCLHotkey.VBHotKey VBHotKey1 
      Left            =   1920
      Top             =   1080
      _ExtentX        =   794
      _ExtentY        =   794
      WinKey          =   0   'False
      Enabled         =   0   'False
   End
   Begin VB.CommandButton fakeTray 
      Caption         =   "fakeTray"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin MSWinsockLib.Winsock Service 
      Left            =   1920
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Server 
      Left            =   1920
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnu_shell 
      Caption         =   "SHell"
      Visible         =   0   'False
      Begin VB.Menu mnu_about 
         Caption         =   "About"
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

Private WithEvents fcsh As clsFCSH
Attribute fcsh.VB_VarHelpID = -1
Private prefs As New clsPreferences
Private log As New clsLog


Private Sub Form_Load()
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
        Err.Clear
    End If
    
    TrayAdd fakeTray.hwnd, Me.Icon, "Flex Compiler SHell Server", MouseMove
    
    Dim hotConfig As New clsHotKeySetup
    hotConfig.SetupKey prefs.prefs, VBHotKey1
    
    Set fcsh = New clsFCSH
    fcsh.initialize log
End Sub

Private Sub mnu_log_Click()
    Shell "notepad.exe " + App.path + "\FCSHServer.log", vbNormalFocus
End Sub

Private Sub Server_ConnectionRequest(ByVal requestID As Long)
    log.xInfo "Accepted connection request " & requestID
    Service.Close
    Service.Accept requestID
End Sub

Private Sub Server_Error(ByVal Number As Integer, description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    log.xError "Server socket error: " + description
End Sub

Private Sub Service_Connect()
    log.xInfo "Connection established"
End Sub

Private Sub Service_Close()
    log.xInfo "Service connection closed"
    DisplayBalloon "Info", "Client disconnected", NIIF_INFO
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
           log.xInfo "Object length: " & objectLength
            
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
    
    log.xInfo "Object type: " & objectType
    
    Select Case objectType
        Case AIR_COMMANDVO:
                            Dim command As New CommandVO
                            command.deSerialize byteArray, pos
                            log.xInfo command.toString
                            executeCommand command
        Case AIR_ERRORVO:
                            Dim error As New ErrorVO
                            error.deSerialize byteArray, pos
                            log.xInfo error.toString
                            processError error
        Case AIR_DATAVO:
                            Dim data As New DataVO
                            data.deSerialize byteArray, pos
                            log.xInfo data.toString
                            processData data
        Case AIR_BALOONVO:
                            Dim baloon As New baloonVO
                            baloon.deSerialize byteArray, pos
                            log.xInfo baloon.toString
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
    log.xInfo "Command: target=" + command.target + " command=" + command.command
    Select Case command.target
        Case "fcsh":
                    fcsh.exec command.command
        Case "fcsh_start":
                    fcsh.Start command.command
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
        log.xInfo "Sending " & (UBound(byteArray) + 1) & " bytes"
        Dim size() As Byte
        ReDim size(0)
        writeLong size, UBound(byteArray) + 1
        Service.SendData size
        Service.SendData byteArray
    Else
        log.xWarning "Network falure. There are no clients connected to the server"
        DisplayBalloon "Network falure", "There are no clients connected to the server", NIIF_ERROR
    End If
End Sub


Private Function convertInt(ByRef byteArray() As Byte) As Long
    Dim buffer(0 To 3) As Byte, i As Long
    For i = 3 To 0 Step -1
        buffer(3 - i) = byteArray(i)
    Next i
    CopyMemory convertInt, buffer(0), 4
End Function

Private Sub fcsh_onStop()
    log.xInfo "fcsh stopped"
    Dim byteArray() As Byte
    ReDim byteArray(0)
    Dim data As New DataVO
    data.target = "fcsh_stop"
    data.data = "Flex Compiler SHell stopped"
    data.serialize byteArray
    sendByteArray byteArray
End Sub

Private Sub fcsh_onStart()
    log.xInfo "fcsh started"
    Dim byteArray() As Byte
    ReDim byteArray(0)
    Dim dataObject As New DataVO
    dataObject.target = "fcsh_start"
    dataObject.data = "Flex Compiler SHell started"
    dataObject.serialize byteArray
    sendByteArray byteArray
End Sub

Private Sub fcsh_onError(msg As String)
    log.xError msg
    Dim byteArray() As Byte
    ReDim byteArray(0)
    Dim error As New ErrorVO
    error.id = 1&
    error.description = msg
    error.serialize byteArray
    sendByteArray byteArray
End Sub

Private Sub fcsh_onDataArrival(data As String)
    log.xDebug data
    Dim byteArray() As Byte
    ReDim byteArray(0)
    Dim dataObject As New DataVO
    dataObject.target = "fcsh_data"
    dataObject.data = data
    dataObject.serialize byteArray
    sendByteArray byteArray
End Sub

Private Sub fcsh_calllback(data As String, target As String)
    log.xInfo "Callback: target=" + target + " data=" + data
    Dim byteArray() As Byte
    ReDim byteArray(0)
    Dim dataObject As New DataVO
    dataObject.target = target
    dataObject.data = data
    dataObject.serialize byteArray
    sendByteArray byteArray
End Sub

Private Sub sendCommand(data As String, target As String)
    log.xInfo "Send command: target=" + target + " data=" + data
    Dim byteArray() As Byte
    ReDim byteArray(0)
    Dim dataObject As New CommandVO
    dataObject.target = target
    dataObject.command = data
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
    log.xInfo "[send complete]"
End Sub

Private Sub VBHotKey1_HotkeyPressed()
    log.xInfo "HotKey pressed"
    sendCommand "empty", "system_hotkey"
End Sub
