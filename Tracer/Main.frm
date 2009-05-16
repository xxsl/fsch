VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form MainWindow 
   Caption         =   "Tracer"
   ClientHeight    =   2385
   ClientLeft      =   12675
   ClientTop       =   9270
   ClientWidth     =   4230
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2385
   ScaleWidth      =   4230
   Begin ComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4230
      _ExtentX        =   7461
      _ExtentY        =   1111
      ButtonWidth     =   1138
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ToolbarIcons"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   5
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Monitor"
            Object.Tag             =   ""
            ImageIndex      =   1
            Style           =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   1e-4
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Clear"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "On Top"
            Object.Tag             =   ""
            ImageIndex      =   5
            Style           =   1
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Alpha"
            Object.Tag             =   ""
            ImageIndex      =   4
            Style           =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.CommandButton cmdFake 
      Caption         =   "cmdFake"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin Tracer.TagAnchor TagAnchor1 
      Left            =   1080
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   820
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Default         =   -1  'True
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Tag             =   "FFFT"
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   50
      TabIndex        =   2
      Tag             =   "FFFT"
      Text            =   "type text to find"
      Top             =   2040
      Width           =   1935
   End
   Begin RichTextLib.RichTextBox Log 
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Tag             =   "TTTT"
      Top             =   645
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2355
      _Version        =   393217
      BackColor       =   15987699
      BorderStyle     =   0
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      RightMargin     =   65000
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Main.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblSearch 
      AutoSize        =   -1  'True
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   2880
      TabIndex        =   4
      Top             =   2070
      Width           =   45
   End
   Begin ComctlLib.ImageList ToolbarIcons 
      Left            =   240
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":04DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":0A30
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":0F82
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":14D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":1A26
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "MainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private prefs As clsPreferences
Private searchPos As Long
Private lastWindowState As Long

Private Const filename As String = "flashlog.txt"




Private Sub Form_Load()
     Set prefs = New clsPreferences
     TrayAdd cmdFake.hWnd, Me.Icon, "Flash Tracer", MouseMove
     lastWindowState = vbNormal
     prefs.initialize
     ReloadLog
End Sub

Private Sub Form_Resize()
    TagAnchor1.DoResize
    
    
    If (Me.WindowState <> vbMinimized) Then
        lastWindowState = Me.WindowState
    End If
    
    If (prefs.minimize) Then
        If Me.WindowState = vbMinimized Then
            Me.Visible = False
        End If
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    TrayDelete
    End
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As ComctlLib.Button)
    Debug.Print "Button selected is " & Button.Index
    Select Case Button.Index
        Case 1:
               If (Button.value = tbrPressed) Then
                    Debug.Print "Run monitoring..."
                    ReloadLog
                    RunMonitoring
               End If
               
        Case 3:
                Dim file As String
                file = prefs.LogDir + filename
                If (FileExists(file)) Then
                    Debug.Print "Clear " + file
                    Open file For Output As #1
                    Close 1
                    ReloadLog
                End If


        Case 4:
                If (Button.value = tbrPressed) Then
                    SetAlwaysOnTopMode Me.hWnd, True
                Else
                    SetAlwaysOnTopMode Me.hWnd, False
                End If
        Case 5:
                If (Button.value = tbrPressed) Then
                     Dim bytOpacity As Byte
                     'Set the transparency level
                     bytOpacity = prefs.alpha
                     Call SetWindowLong(Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
                     Call SetLayeredWindowAttributes(Me.hWnd, 0, bytOpacity, LWA_ALPHA)
                Else
                    Call SetWindowLong(Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) And (Not WS_EX_LAYERED))
                End If

    End Select
End Sub


Private Sub RunMonitoring()
    Dim path As String, filePath As String
    Dim result As Long
    
    path = prefs.LogDir
    filePath = path + filename
    
    Debug.Print "Monitoring folder: " + path

    If (Not FileExists(filePath)) Then
        MsgBox "File " + filePath + " does not exist. Monitoring stopped.", vbCritical
        Toolbar.Buttons(1).value = tbrUnpressed
        Exit Sub
    End If
    
    Do
        result = WaitForFileChange(path, FILE_NOTIFY_CHANGE_SIZE, True, prefs.wait)
     
        Select Case result
         
             Case WAIT_TIMEOUT
                 DoEvents
             
             Case WAIT_FAILED Or WAIT_ABANDONED
                 MsgBox "System failed to monitor " + path, vbCritical
             
             Case WAIT_OBJECT_0
                 ReloadLog
                    
        End Select
    
        If (Toolbar.Buttons(1).value = tbrUnpressed) Then
            Debug.Print "Stop monitoring..."
            Exit Do
        End If
    
    Loop
End Sub

Private Sub ReloadLog()
    Dim filePath As String
    filePath = prefs.LogDir + filename
    If (FileExists(filePath)) Then
        searchPos = 0
        On Error Resume Next
        Log.LoadFile (filePath)
        Log.SelStart = Len(Log.Text)
        If (Err.Number <> 0) Then
            Debug.Print Err.Description
            Err.Clear
        End If
    End If
End Sub

Private Sub txtFind_Change()
    searchPos = 0
    searchPos = Log.Find(txtFind.Text, searchPos, Len(Log.Text))
    Debug.Print "Found text at " & searchPos
    If (searchPos >= 0) Then
        Debug.Print "Found text at " & searchPos
        Log.SelStart = searchPos
        Log.SelLength = Len(txtFind.Text)
        searchPos = searchPos + Len(txtFind.Text)
        lblSearch.Caption = ""
    Else
        lblSearch.Caption = "Text not found"
    End If
End Sub

Private Sub cmdNext_Click()
    If (searchPos >= 0) Then
        searchPos = Log.Find(txtFind.Text, searchPos, Len(Log.Text))
        Debug.Print "Found text at " & searchPos
        If (searchPos >= 0) Then
            Log.SelStart = searchPos
            Log.SelLength = Len(txtFind.Text)
            searchPos = searchPos + Len(txtFind.Text)
            lblSearch.Caption = ""
        Else
            lblSearch.Caption = "Text not found"
        End If
    Else
        txtFind_Change
    End If
End Sub

Private Sub txtFind_GotFocus()
    txtFind.SelStart = 0
    txtFind.SelLength = Len(txtFind.Text)
End Sub

Private Sub cmdFake_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cEvent As Single
    cEvent = X / Screen.TwipsPerPixelX
    Select Case cEvent
        Case MouseMove
            'Debug.Print "MouseMove"
        Case LeftUp
            If (prefs.minimize) Then
                If (Me.Visible = False) Then
                    Me.Visible = True
                    Me.WindowState = lastWindowState
                Else
                    lastWindowState = Me.WindowState
                    Me.WindowState = vbMinimized
                    Me.Visible = False
                End If
            End If
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
            'PopupMenu mnu_shell
        Case RightDown
            'Debug.Print "RightDown"
        Case RightDbClick
            'Debug.Print "RightDbClick"
    End Select
End Sub
