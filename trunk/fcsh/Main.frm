VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Main 
   Caption         =   "Flex compiler shell"
   ClientHeight    =   6585
   ClientLeft      =   3840
   ClientTop       =   2220
   ClientWidth     =   9750
   Icon            =   "Main.frx":0000
   ScaleHeight     =   6585
   ScaleWidth      =   9750
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Service 
      Left            =   7920
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   2
      Top             =   6315
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   476
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtbLog 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   9975
      _Version        =   393217
      BackColor       =   -2147483633
      BorderStyle     =   0
      ScrollBars      =   3
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Main.frx":355A
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
   Begin MSWinsockLib.Winsock Server 
      Left            =   7920
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   44000
   End
   Begin VB.Menu mnuShell 
      Caption         =   "Shell"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private log As clsLog

Private isServerBusy As Boolean
Private port As Long



Private Sub Form_Load()
    'set up logging
    Set log = New clsLog
    log.clsLog rtbLog
    
    'init vars
    isServerBusy = False
    port = 44000
    initSockets 'listen for requests
    
    log.xDebug "started"
    log.xInfo "started"
    log.xError "started"
End Sub


'init main server socket
Private Sub initSockets()
    log.xInfo "Server is listening on port " & port
    Server.Close
    Server.LocalPort = port
    Server.Listen
End Sub

'one more connection
Private Sub Server_ConnectionRequest(ByVal requestID As Long)
    If (Not isServerBusy) Then
        Service.Accept requestID
    End If
End Sub


Private Sub sockMain_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strData As String
    Dim intCnt As Integer
    
    sockMain(Index).GetData strData, vbString
    txtStatus.Text = txtStatus.Text & _
        strData & vbCrLf

    'This sends the data back to the other clients
    For intCnt = 1 To intSockCnt
        If sockMain(intCnt).State = sckConnected Then
            sockMain(intCnt).SendData strData
        End If
    Next intCnt
End Sub

'on resize
Private Sub Form_Resize()
    rtbLog.Width = Me.Width - 130
    
    Dim logWidth As Long
    logWidth = Me.Height - rtbLog.Top - 500 - StatusBar.Height
    If (logWidth > 0) Then
        rtbLog.Height = logWidth
    End If
End Sub

'on quit
Private Sub Form_Terminate()
    Server.Close
End Sub
