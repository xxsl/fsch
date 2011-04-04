VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Begin VB.Form frmMain 
   Caption         =   "Profiler"
   ClientHeight    =   4965
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10560
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   Begin vbalIml6.vbalImageList vbalImageList 
      Left            =   8760
      Top             =   2040
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   4710
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3245
            MinWidth        =   3245
            Text            =   "Memory usage: N/A"
            TextSave        =   "Memory usage: N/A"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab tabs 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   7011
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Live View"
      TabPicture(0)   =   "Main.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LiveView"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "vbalGrid1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Log"
      TabPicture(1)   =   "Main.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Log"
      Tab(1).ControlCount=   1
      Begin vbAcceleratorSGrid6.vbalGrid vbalGrid1 
         Height          =   2055
         Left            =   3600
         TabIndex        =   5
         Top             =   840
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   3625
         BackgroundPictureHeight=   0
         BackgroundPictureWidth=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DisableIcons    =   -1  'True
      End
      Begin VB.ListBox Log 
         Height          =   2700
         IntegralHeight  =   0   'False
         Left            =   -74880
         TabIndex        =   3
         Top             =   480
         Width           =   7455
      End
      Begin MSComctlLib.ListView LiveView 
         Height          =   1935
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   3413
         SortKey         =   1
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ColHdrIcons     =   "imlIcons"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Type"
            Object.Width           =   12347
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "name"
            Text            =   "Size"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Instances"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Timer LiveTimer 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   9120
      Top             =   3120
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   8280
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":05C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0B5C
            Key             =   "up"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0EAE
            Key             =   "down"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnu_main 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu mnu_test 
         Caption         =   "Test"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'server socket
Private WithEvents server    As JBSOCKETSERVERLib.server
Attribute server.VB_VarHelpID = -1

'tray handle
Private WithEvents m_AppTray As frmSysTray
Attribute m_AppTray.VB_VarHelpID = -1

Private socketData           As clsSocketData

Private processor            As New clsProcessor

Private isCommandPending     As Boolean

Private selectedLiveObject As String

Private logx As New clsLog

Private prefs As New clsPreferences

'Profiler connected
Private Sub server_OnConnectionEstablished(ByVal Socket As JBSOCKETSERVERLib.ISocket)
    logx.xInfo "connection established: " & addressAsString(Socket)
    'prevent other other connections
    server.StopListening
    
    Set socketData = New clsSocketData
    Socket.UserData = socketData
    
    'wait for data
    Socket.RequestRead
    
    LiveTimer.Enabled = True
End Sub

'Profiler disconnected
Private Sub server_OnConnectionClosed(ByVal Socket As JBSOCKETSERVERLib.ISocket)
    logx.xInfo "connection closed: " & addressAsString(Socket)
    
    Socket.UserData = Null
    LiveTimer.Enabled = False
    
    selectedLiveObject = ""
    
    'we are ready for new session
    server.StartListening
    'TODO clean up session data
End Sub

'Profiler data processing
Private Sub server_OnDataReceived(ByVal Socket As JBSOCKETSERVERLib.ISocket, ByVal Data As JBSOCKETSERVERLib.IData)
  
    Dim sckData As clsSocketData

    Set sckData = Socket.UserData
    sckData.append Data
    
    'logger "data counter=" & sckData.increment
    
    If (Not sckData.locked) Then
        processor.processCommand sckData, Socket
    Else
        'logger "waiting for data"
    End If
    
    Socket.RequestRead
End Sub

Private Function addressAsString(Socket As JBSOCKETSERVERLib.ISocket) As String
    addressAsString = Socket.RemoteAddress.Address & " : " & Socket.RemoteAddress.port
End Function

Private Sub Form_Resize()
    If (Me.WindowState = vbNormal) Then
        tabs.Width = Me.Width - 500
        tabs.Height = Me.Height - 1000
        
        If (tabs.Tab = 0) Then
            LiveView.Width = tabs.Width - LiveView.Left * 2
            LiveView.Height = tabs.Height - LiveView.Top - 120
        End If
        
        Log.Width = tabs.Width - 240
        Log.Height = tabs.Height - 600
    End If
End Sub

Private Sub tabs_Click(PreviousTab As Integer)
    Form_Resize
End Sub

Private Sub LiveView_ItemClick(ByVal item As MSComctlLib.ListItem)
   selectedLiveObject = item.Text
End Sub

Private Sub LiveTimer_Timer()

    Dim types  As New Dictionary

    Dim s      As clsNewObjectSample
    
    Dim liveObj As clsLiveObject

    Dim sample As Variant

    Dim i      As Long

    Dim memory As Long

    Dim str    As String
    
    For Each sample In processor.samples

        Set s = sample

        If (types.Exists(s.getType)) Then

            Set liveObj = types.item(s.getType)

            types.Remove (s.getType)
            
            liveObj.instances = liveObj.instances + 1
            liveObj.Size = liveObj.Size + s.getSize

            types.Add s.getType, liveObj

        Else

            Set liveObj = New clsLiveObject
            liveObj.NAME = s.getType
            liveObj.instances = 1
            liveObj.Size = s.getSize
            types.Add s.getType, liveObj

        End If

        memory = memory + s.getSize
    Next
    
    LiveView.Visible = False
    
    Dim Size As Long

    Size = LiveView.ListItems.Count - 1
    
    For i = 0 To Size
        LiveView.ListItems.Remove 1
    Next
    
    Dim selectedItem As ListItem
   
    For Each sample In types.Keys

        str = sample
        Set liveObj = types.item(str)

        Dim Row As ListItem

        Set Row = LiveView.ListItems.Add
        Row.Text = liveObj.NAME

        Dim col As ListSubItem

        Set col = Row.ListSubItems.Add
        col.Text = "" & liveObj.Size
        
        Set col = Row.ListSubItems.Add
        col.Text = "" & liveObj.instances
        
        If (selectedLiveObject = liveObj.NAME) Then
            Set selectedItem = Row
        End If
    Next
    
    Dim ColumnHeader As MSComctlLib.ColumnHeader

    For Each sample In LiveView.ColumnHeaders

        Set ColumnHeader = sample

        Select Case ColumnHeader.icon

            Case "down"
                ColumnHeader.icon = Empty
                LiveView_ColumnClick ColumnHeader

            Case "up"
                ColumnHeader.icon = "down"
                LiveView_ColumnClick ColumnHeader
        End Select

    Next
    
    If (Not selectedItem Is Nothing) Then
        Set LiveView.selectedItem = selectedItem
        'selectedItem.EnsureVisible
    End If
    
    LiveView.Visible = True
    
    status.Panels.item(1).Text = "Memory usage:" & memory \ 1024 & " kb"
End Sub

Private Sub LiveView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call ClearHeaderIcons(ColumnHeader.Index)

    Select Case ColumnHeader.Index

        Case 1

            Select Case ColumnHeader.icon

                Case "down"
                    ColumnHeader.icon = "up"
                    Call SortColumn(LiveView, ColumnHeader.Index, sortDescending, sortAlpha)

                Case "up"
                    ColumnHeader.icon = "down"
                    Call SortColumn(LiveView, ColumnHeader.Index, sortAscending, sortAlpha)

                Case Else
                    ColumnHeader.icon = "down"
                    Call SortColumn(LiveView, ColumnHeader.Index, sortAscending, sortAlpha)
            End Select

        Case 2

            Select Case ColumnHeader.icon

                Case "down"
                    ColumnHeader.icon = "up"
                    Call SortColumn(LiveView, ColumnHeader.Index, sortDescending, sortNumeric)

                Case "up"
                    ColumnHeader.icon = "down"
                    Call SortColumn(LiveView, ColumnHeader.Index, sortAscending, sortNumeric)

                Case Else
                    ColumnHeader.icon = "down"
                    Call SortColumn(LiveView, ColumnHeader.Index, sortAscending, sortNumeric)
            End Select
        Case 3

            Select Case ColumnHeader.icon

                Case "down"
                    ColumnHeader.icon = "up"
                    Call SortColumn(LiveView, ColumnHeader.Index, sortDescending, sortNumeric)

                Case "up"
                    ColumnHeader.icon = "down"
                    Call SortColumn(LiveView, ColumnHeader.Index, sortAscending, sortNumeric)

                Case Else
                    ColumnHeader.icon = "down"
                    Call SortColumn(LiveView, ColumnHeader.Index, sortAscending, sortNumeric)
            End Select
    End Select

End Sub

Private Sub ClearHeaderIcons(CurrentHeader As Integer)

    Dim i As Integer

    For i = 1 To LiveView.ColumnHeaders.Count

        If LiveView.ColumnHeaders(i).Index <> CurrentHeader Then
            LiveView.ColumnHeaders(i).icon = Empty
        End If

    Next

End Sub


'Public preloader As New clsResource
'
'Private Sub LoadPNG()
'    'extract files and save
'    preloader.getResourceByName START_FCSH
'    preloader.getResourceByName STOP_FCSH
'    preloader.getResourceByName LOG_CLEAR
'    preloader.getResourceByName BUILD_TASK
'    preloader.getResourceByName INCREMENTAL_ON
'    preloader.getResourceByName INCREMENTAL_OFF
'    preloader.getResourceByName TARGET_INFO
'    preloader.getResourceByName Options
'    preloader.getResourceByName ON_TOP
'    preloader.getResourceByName Alpha
'    preloader.getResourceByName ABOUT
'    preloader.getResourceByName APP_APPEARANCE
'    preloader.getResourceByName KEYBOARD
'    preloader.getResourceByName WARNING_PNG
'    preloader.getResourceByName CUSTOM_COMMAND
'
'    preloader.getResourceByName IDLE_PNG
'    preloader.getResourceByName STOPPED_PNG
'    preloader.getResourceByName ERROR_PNG
'    preloader.getResourceByName EXEC_PNG, ".gif"
'
'    'load
'    Dim pngLoader As New clsPngToImageList
'    pngLoader.Initialize picIconLoad, picClear, pngImages, Log
'    pngLoader.LoadIcons preloader.extractedFiles
'
'
'    'setup toolbar
'    Set Toolbar.ImageList = pngImages
'    Toolbar.Buttons(RUN_BUTTON).Image = preloader.getIndex(START_FCSH)
'    Toolbar.Buttons(BUILD_BUTTON).Image = preloader.getIndex(BUILD_TASK)
'    Toolbar.Buttons(TYPE_BUTTON).Image = preloader.getIndex(INCREMENTAL_ON)
'    Toolbar.Buttons(INFO_BUTTON).Image = preloader.getIndex(TARGET_INFO)
'    Toolbar.Buttons(CUSTOM_BUTTON).Image = preloader.getIndex(CUSTOM_COMMAND)
'    Toolbar.Buttons(OPTIONS_BUTTON).Image = preloader.getIndex(Options)
'    Toolbar.Buttons(CLEAR_BUTTON).Image = preloader.getIndex(LOG_CLEAR)
'    Toolbar.Buttons(ONTOP_BUTTON).Image = preloader.getIndex(ON_TOP)
'    Toolbar.Buttons(TRANSPARENT_BUTTON).Image = preloader.getIndex(Alpha)
'    Toolbar.Buttons(ABOUT_BUTTON).Image = preloader.getIndex(ABOUT)
'End Sub

'start up
Private Sub Form_Load()
    Load frmSysTray
    Set m_AppTray = New frmSysTray
    
    m_AppTray.AddToTray imlIcons.ListImages(1).Picture.Handle
    m_AppTray.ToolTip = "Flex profiler"

    logx.setWindow Log

    Set processor.logx = logx

    Set server = CreateSocketServer(9999)
    server.StartListening
End Sub

'exit dialog
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim result As VbMsgBoxResult
    
    If (Me.Visible) Then
        result = MsgBox("This action will stop profiler. Are you sure?", vbExclamation + vbOKCancel, "Confirm")

        If (result <> vbOK) Then
            Cancel = 1
        End If
    End If

End Sub

'quit
Private Sub Form_Unload(Cancel As Integer)
    server.StopListening
    Set processor.logx = Nothing
    Unload m_AppTray
    Unload frmSysTray
End Sub

'Tray menu handler
Private Sub m_AppTray_SysTrayMouseDown(ByVal eButton As MouseButtonConstants)

    Select Case eButton

        Case vbRightButton
            PopupMenu mnu_main
    
    End Select

End Sub


