VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Begin VB.Form MainWindow 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Tracer"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   6420
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStopMonitoring 
      Caption         =   "Stop Monitoring"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin RichTextLib.RichTextBox Log 
      Height          =   6255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   11033
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   3
      RightMargin     =   65000
      TextRTF         =   $"Main.frx":0000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Monitoring"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "MainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private StopWhatching As Boolean
Private LogFilePath As String


Private Sub Form_Load()
    StopWhatching = False
    LogFilePath = "c:\Documents and Settings\aturtsevitch.AMP\Application Data\Macromedia\Flash Player\Logs\"
End Sub

Private Sub cmdStopMonitoring_Click()
    StopWhatching = True
End Sub




Private Sub Command1_Click()
    StopWhatching = False

    Dim result As Long
    
    Do
        result = WaitForFileChange(LogFilePath, FILE_NOTIFY_CHANGE_SIZE, True, 10)
     
        Select Case result
         
             Case WAIT_TIMEOUT
                 DoEvents
             
             Case WAIT_FAILED Or WAIT_ABANDONED
                 'MsgBox "Wait failed or abandoned"
                 'Exit Do
             
             Case WAIT_OBJECT_0
                 'MsgBox "The state of the specified object is signaled"
                 'Exit Do
                 Log.LoadFile (LogFilePath + "flashlog.txt")
                 Log.SelStart = Len(Log.Text)
                    
        End Select
    
        If (StopWhatching) Then
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

