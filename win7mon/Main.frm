VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monitor"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   7380
   FillColor       =   &H0080FF80&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmAdapters 
      Caption         =   "Adapters"
      Height          =   855
      Left            =   120
      TabIndex        =   20
      Top             =   4680
      Width           =   7215
      Begin VB.ComboBox cmbInterfaces 
         Height          =   315
         ItemData        =   "Main.frx":0000
         Left            =   120
         List            =   "Main.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   360
         Width           =   6855
      End
   End
   Begin VB.Frame frmData 
      Caption         =   "Traffic"
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   7215
      Begin VB.Label lblRecv 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1320
         TabIndex        =   19
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblSent 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1320
         TabIndex        =   18
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sent"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1095
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Received"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1335
         Width           =   975
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bytes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         TabIndex        =   15
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bytes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         TabIndex        =   14
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bps"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2760
         TabIndex        =   13
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bps"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2760
         TabIndex        =   12
         Top             =   480
         Width           =   375
      End
      Begin VB.Label lblBPSr 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1080
         TabIndex        =   11
         Top             =   720
         Width           =   795
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Received"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblBPSs 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1080
         TabIndex        =   9
         Top             =   480
         Width           =   795
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sent"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblAvrS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   1920
         TabIndex        =   7
         Top             =   480
         Width           =   795
      End
      Begin VB.Label lblAvrR 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   1920
         TabIndex        =   6
         Top             =   720
         Width           =   795
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Current"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Average"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1920
         TabIndex        =   4
         Top             =   240
         Width           =   795
      End
      Begin VB.Label lblType 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.Frame frmGraph 
      Caption         =   "Graph"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.PictureBox Chart 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         FillColor       =   &H00008000&
         ForeColor       =   &H80000008&
         Height          =   2115
         Left            =   120
         ScaleHeight     =   141
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   465
         TabIndex        =   1
         Top             =   240
         Width           =   6975
      End
   End
   Begin VB.Timer tmrGraph 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3480
      Top             =   7320
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3480
      Top             =   6720
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2760
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0004
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0160
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":02BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0418
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ChartSend() As Long
Private ChartRecv() As Long
Private LastSend    As Long
Private LastRecv    As Long
Private currSend    As Long
Private currRecv    As Long
Private gScale      As Long
Private CurrPoss    As Long

Private m_objIpHelper As CIpHelper
Private BytesSent     As Long
Private BytesReceived As Long

'tray control
Private m_tray As New CTray

Private Sub Chart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Chart.ToolTipText = "Received: " & Abs(ChartRecv(X)) & vbCrLf & " Sent: " & Abs(ChartSend(X))
End Sub

Private Sub Form_Click()
    Me.WindowState = vbNormal
End Sub

'CSEH: ErrReportAndTrace
Private Sub Form_Load()

    If (CreateManifest) Then End
    ReDim ChartSend(Chart.ScaleWidth)
    ReDim ChartRecv(Chart.ScaleWidth)
    CurrPoss = 1
    gScale = 1
    Set m_objIpHelper = New CIpHelper

    Dim objInterface As CInterface

    For Each objInterface In m_objIpHelper.Interfaces

        cmbInterfaces.AddItem "[#" & objInterface.InterfaceIndex & "] " + "[Type: " + objInterface.InterfaceTypes & "] " & objInterface.InterfaceDescription
    Next

    If (m_objIpHelper.Interfaces.Count > 0) Then
        cmbInterfaces.ListIndex = 10
    End If

    Set objInterface = m_objIpHelper.Interfaces(1)
    BytesReceived = objInterface.OctetsReceived
    BytesSent = objInterface.OctetsSent

    m_tray.Add Me.HWND, ImageList1.ListImages(4).Picture, "Bytes received: " & Trim$(Format$(BytesReceived, "###,###,###,###")) & vbCrLf & "Bytes sent: " & Trim$(Format$(BytesSent, "###,###,###,###")) & vbNullChar, MouseMove
    m_tray.DisplayBalloon App.EXEName, App.EXEName + " started!", NIIF_INFO
    

    PlotChart
    tmrGraph.Enabled = True
    Timer1.Enabled = True
    
    Me.Show
End Sub

Private Sub PlotChart()

    Dim gCount As Long
    Dim gPos   As Long
    Dim rc     As Long
    Dim sc     As Long
    Dim avrS   As Long
    Dim avrR   As Long
    Dim maxRS  As Long
    Dim scval  As Long
    
    Chart.Cls

    Dim median As Long

    median = (Chart.ScaleHeight \ 2)
    scval = CStr(CSng(gScale) * (Chart.ScaleHeight \ 2))
    Debug.Print "Net Monitor (Full Scale = " & CStr(scval) & ")"

    Dim size As Long

    size = UBound(ChartSend)

    For gCount = 1 To size
        gPos = gCount
        Chart.Line (gCount - 1, median)-(gCount, median + (ChartSend(gCount)) / gScale), vbRed, BF
        Chart.Line (gCount - 1, median)-(gCount, median - (ChartRecv(gCount)) / gScale), vbBlue, BF

        If ChartSend(gPos) <> 0 Then
            sc = sc + 1
            avrS = avrS + ChartSend(gPos)

            If maxRS > ChartSend(gPos) Then maxRS = ChartSend(gPos)
        End If

        If ChartRecv(gPos) <> 0 Then
            rc = rc + 1
            avrR = avrR + ChartRecv(gPos)

            If maxRS > ChartRecv(gPos) Then maxRS = ChartRecv(gPos)
        End If

    Next

    If sc <> 0 Then Me.lblAvrS = -(CLng(avrS / sc)) Else Me.lblAvrS = 0

    If rc <> 0 Then Me.lblAvrR = -(CLng(avrR / rc)) Else Me.lblAvrR = 0

    'draw lines
    Dim ScaleWidth As Long, percent25 As Long, percent75 As Long

    percent75 = (Chart.ScaleHeight \ 8) * 3
    percent25 = (Chart.ScaleHeight \ 8)
    ScaleWidth = Chart.ScaleWidth
    Chart.ForeColor = vb3DDKShadow
    Chart.Line (0, median)-(ScaleWidth, median)
    Chart.Line (0, median - percent75)-(ScaleWidth, median - percent75)
    Chart.Line (0, median + percent75)-(ScaleWidth, median + percent75)
    Chart.DrawStyle = 2
    Chart.Line (0, median - percent25)-(ScaleWidth, median - percent25)
    Chart.Line (0, median + percent25)-(ScaleWidth, median + percent25)
    Chart.DrawStyle = 0
    Chart.ForeColor = vbYellow
    Chart.Line (0, median)-(ScaleWidth, median)
    Chart.Line (0, median)-(ScaleWidth, median)
    gScale = -(maxRS / median) + 1
End Sub

Private Sub UpdateInterfaceInfo()
    Dim objInterface       As CInterface
    Static st_objInterface As CInterface
    Static lngBytesRecv    As Long
    Static lngBytesSent    As Long
    Dim blnIsRecv          As Boolean
    Dim blnIsSent          As Boolean
    If st_objInterface Is Nothing Then Set st_objInterface = New CInterface
    Set objInterface = m_objIpHelper.Interfaces(cmbInterfaces.ListIndex + 1)

    Select Case objInterface.InterfaceType

        Case MIB_IF_TYPE_ETHERNET
            lblType.Caption = "Ethernet"

        Case MIB_IF_TYPE_FDDI
            lblType.Caption = "FDDI"

        Case MIB_IF_TYPE_LOOPBACK
            lblType.Caption = "Loopback"

        Case MIB_IF_TYPE_OTHER
            lblType.Caption = "Other"

        Case MIB_IF_TYPE_PPP
            lblType.Caption = "PPP"

        Case MIB_IF_TYPE_SLIP
            lblType.Caption = "SLIP"

        Case MIB_IF_TYPE_TOKENRING
            lblType.Caption = "TokenRing"
    End Select

    BytesReceived = objInterface.OctetsReceived
    BytesSent = objInterface.OctetsSent
    lblRecv.Caption = Trim$(Format$(BytesReceived, "###,###,###,###"))
    lblSent.Caption = Trim$(Format$(BytesSent, "###,###,###,###"))
    Set st_objInterface = objInterface
    '---------------
    blnIsRecv = (BytesReceived > lngBytesRecv)
    blnIsSent = (BytesSent > lngBytesSent)

    lngBytesRecv = BytesReceived
    lngBytesSent = BytesSent
    
   
    m_tray.ModifyIcon GetIconForConnection(blnIsRecv, blnIsSent)
    m_tray.ModifyTooltip "Bytes received: " & Trim$(Format$(BytesReceived, "###,###,###,###")) & vbCrLf & "Bytes sent: " & Trim$(Format$(BytesSent, "###,###,###,###")) & vbNullChar
End Sub

Private Function GetIconForConnection(ByVal blnIsRecv As Boolean, ByVal blnIsSent As Boolean) As StdPicture
    If blnIsRecv And blnIsSent Then
        Set GetIconForConnection = ImageList1.ListImages(4).Picture
        Exit Function
    ElseIf (Not blnIsRecv) And blnIsSent Then
        Set GetIconForConnection = ImageList1.ListImages(3).Picture
        Exit Function
    ElseIf blnIsRecv And (Not blnIsSent) Then
        Set GetIconForConnection = ImageList1.ListImages(2).Picture
        Exit Function
    ElseIf Not (blnIsRecv And blnIsSent) Then
        Set GetIconForConnection = ImageList1.ListImages(1).Picture
        Exit Function
    End If
End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.WindowState = vbNormal
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim msg As Long

    If Me.ScaleMode = vbPixels Then
        msg = X
    Else
        msg = X / Screen.TwipsPerPixelX
    End If

    Select Case msg

        Case TrayRetunEventEnum.LeftDown '515 restore form window

            If Me.WindowState = vbNormal Then
                Me.WindowState = vbMinimized
                Me.Visible = False
            Else
                Me.WindowState = vbNormal
                Me.Visible = True
                Me.SetFocus
            End If

    End Select

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    m_tray.Delete
End Sub

Private Sub Form_Terminate()
    m_tray.Delete
End Sub

Private Sub Form_Unload(Cancel As Integer)
    m_tray.Delete
End Sub

Private Sub Timer1_Timer()
    Call UpdateInterfaceInfo
End Sub

Private Sub tmrGraph_Timer()

    Dim bs As Long, br As Long

    bs = BytesSent
    br = BytesReceived

    If LastRecv = 0 Then LastRecv = br

    If LastSend = 0 Then LastSend = bs
    currRecv = LastRecv - br
    currSend = LastSend - bs
    LastRecv = br
    LastSend = bs
    ChartSend(CurrPoss) = currSend
    ChartRecv(CurrPoss) = currRecv
    lblBPSr = -(currRecv)
    lblBPSs = -(currSend)
    Chart.ForeColor = vbButtonFace
    Chart.Line (CurrPoss, 0)-(CurrPoss, Chart.ScaleHeight)

    If CurrPoss = Chart.ScaleWidth Then
        CurrPoss = 1
    Else
        CurrPoss = CurrPoss + 1
    End If

    PlotChart
    Chart.ForeColor = vbGreen
    Chart.Line (CurrPoss, 0)-(CurrPoss, Chart.ScaleHeight)
End Sub

