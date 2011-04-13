VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
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
   Begin VB.Timer tmrResize 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9960
      Top             =   3840
   End
   Begin VB.Timer tmrSocketSpeed 
      Interval        =   5000
      Left            =   8520
      Top             =   3840
   End
   Begin MSWinsockLib.Winsock Service 
      Left            =   9120
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Server 
      Left            =   8400
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
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
      TabIndex        =   3
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
      TabIndex        =   1
      Top             =   4710
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   4233
            MinWidth        =   4233
            Text            =   "Socket speed: N/A"
            TextSave        =   "Socket speed: N/A"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   4233
            MinWidth        =   4233
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
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Log"
      TabPicture(1)   =   "Main.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Log"
      Tab(1).ControlCount=   1
      Begin vbAcceleratorSGrid6.vbalGrid LiveView 
         Height          =   2055
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   3625
         RowMode         =   -1  'True
         GridLines       =   -1  'True
         BackgroundPictureHeight=   0
         BackgroundPictureWidth=   0
         AlternateRowBackColor=   16053492
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HeaderHotTrack  =   0   'False
         DisableIcons    =   -1  'True
      End
      Begin VB.ListBox Log 
         Height          =   2700
         IntegralHeight  =   0   'False
         Left            =   -74880
         TabIndex        =   2
         Top             =   480
         Width           =   7455
      End
   End
   Begin VB.Timer LiveTimer 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   9240
      Top             =   3840
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
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":05C2
            Key             =   ""
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

Private WithEvents m_AppTray As frmSysTray
Attribute m_AppTray.VB_VarHelpID = -1

Private socketData           As clsSocketData

Private processor            As New clsProcessor

Private isCommandPending     As Boolean

Private selectedLiveObject   As String

Private logx                 As New clsLog

Private prefs                As New clsPreferences

Private dataCount            As Double

'last window state
Private prevState            As Long

'last sort column
Private lastSortColumn       As Long

Private Sub Server_ConnectionRequest(ByVal requestID As Long)

    If (Not Service.State = sckConnected) Then
        logx.xInfo "Accepted connection request " & requestID
        Service.Close
        Service.Accept requestID
        Set socketData = New clsSocketData
        Set socketData.socket = Service
        LiveTimer.Enabled = False
    Else
        logx.xInfo "Connection request ignored " & requestID
    End If

End Sub

Private Sub Server_Error(ByVal Number As Integer, _
                         description As String, _
                         ByVal Scode As Long, _
                         ByVal Source As String, _
                         ByVal HelpFile As String, _
                         ByVal HelpContext As Long, _
                         CancelDisplay As Boolean)
    logx.xError "Server socket error: " + description
End Sub

Private Sub Service_Connect()
    logx.xInfo "Connection established"
End Sub

Private Sub Service_Close()
    logx.xInfo "Service connection closed"
    'DisplayBalloon "Info", "Client disconnected", NIIF_INFO
    LiveTimer.Enabled = False
End Sub

Private Sub Service_Error(ByVal Number As Integer, _
                          description As String, _
                          ByVal Scode As Long, _
                          ByVal Source As String, _
                          ByVal HelpFile As String, _
                          ByVal HelpContext As Long, _
                          CancelDisplay As Boolean)
    logx.xError "Service socket error: " + description
    Service.Close
End Sub

Private Sub Service_DataArrival(ByVal bytesTotal As Long)
    If (processor.minSize <= bytesTotal) Then
        socketData.bytesAvailable = bytesTotal
        socketData.Refresh
        processor.processCommand socketData
        dataCount = dataCount + (bytesTotal - socketData.bytesAvailable)
    End If
End Sub

Private Sub Form_Resize()

    If (Me.WindowState <> vbMinimized) Then
        tabs.Width = Me.Width - 500
        tabs.Height = Me.Height - 1000 - Toolbar.Height
        
        If (tabs.Tab = 0) Then
            LiveView.Width = tabs.Width - LiveView.Left * 2
            LiveView.Height = tabs.Height - LiveView.Top - 120
        End If
        
        Log.Width = tabs.Width - 240
        Log.Height = tabs.Height - 600
    End If

    If (Me.WindowState <> prevState) Then
        prevState = Me.WindowState
        tmrResize.Enabled = True
    End If

End Sub

Private Sub tmrResize_Timer()

    Form_Resize
    tmrResize.Enabled = False
End Sub

Private Sub tabs_Click(PreviousTab As Integer)

    Form_Resize
End Sub

Private Sub LiveView_ItemClick(ByVal item As MSComctlLib.ListItem)
    selectedLiveObject = item.Text
End Sub

Private Sub tmrSocketSpeed_Timer()
    status.Panels.item(1).Text = "Socket speed: " & FormatNumber(dataCount / 5 / 1024, 2) & " kb/s"
    dataCount = 0
End Sub

Private Sub LiveTimer_Timer()

    Dim types   As New Dictionary

    Dim s       As clsNewObjectSample
    
    Dim liveObj As clsLiveObject

    Dim sample  As Variant

    Dim i       As Long

    Dim memory  As Long

    Dim str     As String
    
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
   
    LiveView.Redraw = False
    
    LiveView.clear
    
    LiveView.Rows = UBound(types.Keys) + 1
   
    For i = 1 To UBound(types.Keys) + 1
        
        str = types.Keys(i - 1)
        Set liveObj = types.item(str)

        LiveView.CellText(i, 1) = liveObj.NAME
        LiveView.CellText(i, 2) = "" & liveObj.Size
        LiveView.CellText(i, 3) = "" & liveObj.instances
    Next
    
    LiveViewRestoreSort

    LiveView.Redraw = True
    
    status.Panels.item(2).Text = "Memory usage: " & FormatNumber(memory / 1024, 2) & " kb"
End Sub

Private Sub LiveViewRestoreSort()
    Dim iCol     As Long

    Dim iSortCol As Long

    Dim sJunk()  As String, eJunk() As ECGSortOrderConstants

    With LiveView.SortObject
        .ClearNongrouped
        iSortCol = .IndexOf(lastSortColumn)

        If (iSortCol <= 0) Then
            iSortCol = .Count + 1
        End If
      
        .SortColumn(iSortCol) = lastSortColumn

        If (LiveView.ColumnSortOrder(lastSortColumn) = CCLOrderNone) Or (LiveView.ColumnSortOrder(lastSortColumn) = CCLOrderAscending) Then
            .SortOrder(iSortCol) = CCLOrderAscending
        Else
            .SortOrder(iSortCol) = CCLOrderDescending
        End If

        LiveView.ColumnSortOrder(lastSortColumn) = .SortOrder(iSortCol)
        .SortType(iSortCol) = LiveView.ColumnSortType(lastSortColumn)
      
    End With
   
    Screen.MousePointer = vbHourglass
    LiveView.Sort
    Screen.MousePointer = vbDefault

End Sub

Private Sub LiveView_ColumnClick(ByVal lCol As Long)

    lastSortColumn = lCol

    Dim iCol     As Long

    Dim iSortCol As Long

    Dim sJunk()  As String, eJunk() As ECGSortOrderConstants

    With LiveView.SortObject
        .ClearNongrouped
        iSortCol = .IndexOf(lCol)

        If (iSortCol <= 0) Then
            iSortCol = .Count + 1
        End If
      
        .SortColumn(iSortCol) = lCol

        If (LiveView.ColumnSortOrder(lCol) = CCLOrderNone) Or (LiveView.ColumnSortOrder(lCol) = CCLOrderDescending) Then
            .SortOrder(iSortCol) = CCLOrderAscending
        Else
            .SortOrder(iSortCol) = CCLOrderDescending
        End If

        LiveView.ColumnSortOrder(lCol) = .SortOrder(iSortCol)
        .SortType(iSortCol) = LiveView.ColumnSortType(lCol)
      
        ' Place ascending/descending icon:
        '      For iCol = 1 To grdLib.Columns
        '         If (iCol <> lCol) Then
        '            If Not (grdLib.ColumnIsGrouped(iCol)) Then
        '               If grdLib.ColumnImage(iCol) > -1 Then
        '                  grdLib.ColumnImage(iCol) = -1
        '               End If
        '            End If
        '         ElseIf grdLib.ColumnHeader(iCol) <> "" Then
        '            grdLib.ColumnImageOnRight(iCol) = True
        '            If (.SortOrder(iSortCol) = CCLOrderAscending) Then
        '               grdLib.ColumnImage(iCol) = 0
        '            Else
        '               grdLib.ColumnImage(iCol) = 1
        '            End If
        '         End If
        '      Next iCol
      
    End With
   
    Screen.MousePointer = vbHourglass
    LiveView.Sort
    Screen.MousePointer = vbDefault

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
    
    LiveView.AddColumn "Name", "Name", ecgHdrTextALignLeft, , 500, True, False, , True, , False, CCLSortString
    LiveView.AddColumn "Size", "Size", ecgHdrTextALignLeft, , , True, False, , True, , False, CCLSortNumeric
    LiveView.AddColumn "Instances", "Instances", ecgHdrTextALignLeft, , , True, False, , True, , False, CCLSortNumeric

    LiveView.StretchLastColumnToFit = True
    
    lastSortColumn = 1
    
    Server.Close
    Server.LocalPort = 9999
    Server.Listen
    
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
    Server.Close
    Service.Close
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
