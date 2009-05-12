VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form MainWindow 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Tracer"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   7245
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   741
      ButtonWidth     =   609
      Appearance      =   1
      ImageList       =   "ToolbarIcons"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   1
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   1
            Style           =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin RichTextLib.RichTextBox Log 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   420
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   11456
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      RightMargin     =   65000
      TextRTF         =   $"Main.frx":0000
   End
   Begin ComctlLib.ImageList ToolbarIcons 
      Left            =   0
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":0084
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


Private Sub Form_Load()
     Set prefs = New clsPreferences
     prefs.initialize
End Sub

Private Sub Form_Resize()
    Log.Width = Me.Width - Log.Left - Log.Left - 120
    
    Dim logheight As Long
    logheight = Me.Height - Log.Top - 410
    If (logheight > 0) Then
        Log.Height = Me.Height - Log.Top - 410
    End If
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Debug.Print "QueryUnload"
    Toolbar.Buttons(1).Value = tbrUnpressed
End Sub


Private Sub Toolbar_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Index
        Case 1:
               If (Button.Value = tbrPressed) Then
                    Debug.Print "Run monitoring..."
                    RunMonitoring
               End If
    End Select
End Sub


Private Sub RunMonitoring()
    Dim path As String
    Dim result As Long
    
    path = prefs.LogDir
    Debug.Print "Monitoring folder: " + path
    
    Do
        result = WaitForFileChange(path, FILE_NOTIFY_CHANGE_SIZE, True, 10)
     
        Select Case result
         
             Case WAIT_TIMEOUT
                 DoEvents
             
             Case WAIT_FAILED Or WAIT_ABANDONED
                 'MsgBox "Wait failed or abandoned"
                 'Exit Do
             
             Case WAIT_OBJECT_0
                 'MsgBox "The state of the specified object is signaled"
                 'Exit Do
                 Log.LoadFile (path + "flashlog.txt")
                 Log.SelStart = Len(Log.Text)
                    
        End Select
    
        If (Toolbar.Buttons(1).Value = tbrUnpressed) Then
            Debug.Print "Stop monitoring..."
            Exit Do
        End If
    
    Loop
End Sub

'        Case ONTOP_BUTTON:
'                If (Toolbar.Buttons(ONTOP_BUTTON).Value = tbrPressed) Then
'                    SetAlwaysOnTopMode Me.hWnd, True
'                Else
'                    SetAlwaysOnTopMode Me.hWnd, False
'                End If
'        Case TRANSPARENT_BUTTON:
'                If (Toolbar.Buttons(TRANSPARENT_BUTTON).Value = tbrPressed) Then
'                     Dim bytOpacity As Byte
'                     'Set the transparency level
'                     bytOpacity = config.Alpha
'                     Call SetWindowLong(Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
'                     Call SetLayeredWindowAttributes(Me.hWnd, 0, bytOpacity, LWA_ALPHA)
'                Else
'                    Call SetWindowLong(Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) And (Not WS_EX_LAYERED))
'                End If
