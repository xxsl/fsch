VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Runner 
   Caption         =   "Ant Runner"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   10590
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox Log 
      Height          =   4815
      Left            =   3960
      TabIndex        =   4
      Top             =   1440
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   8493
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      RightMargin     =   65000
      TextRTF         =   $"Runner.frx":0000
   End
   Begin MSComctlLib.ImageList Icons 
      Left            =   3600
      Top             =   6720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Runner.frx":008D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   2880
      Top             =   6840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   3
      Top             =   7440
      Width           =   10590
      _ExtentX        =   18680
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   1032
            MinWidth        =   988
            Text            =   "Status:"
            TextSave        =   "Status:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10590
      _ExtentX        =   18680
      _ExtentY        =   635
      ButtonWidth     =   609
      Appearance      =   1
      Style           =   1
      ImageList       =   "Icons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.ListBox List1 
      Height          =   4860
      IntegralHeight  =   0   'False
      Left            =   360
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1320
      Width           =   3375
   End
   Begin AntRunner.ctlSplitterEx VSplitter 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   9975
      _extentx        =   18230
      _extenty        =   12091
   End
End
Attribute VB_Name = "Runner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public build As New clsBuildXML



Private Sub Form_Load()
    VSplitter.AttachObjects List1, Log
    VSplitter.TileMode = TILE_VERTICALLY
End Sub

Private Sub Form_Resize()
    Dim size As Long
    
    VSplitter.Width = Me.Width - 380
    
    size = Me.Height - VSplitter.Top - (Me.Height - Status.Top) - 100
    If (size > 0) Then
        VSplitter.Height = size
    End If
End Sub


Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1:
                loadBuildFile
    End Select
End Sub

Private Sub loadBuildFile()
   Dim i As Long

    CD.Filter = "XML Files (*.xml,*.xsl)|*.xml;*xsl"
   CD.ShowOpen
   If (Len(CD.FileName) > 0) Then
       build.LoadXML CD.FileName
       
       For i = 0 To UBound(build.getTargets)
          List1.AddItem build.getTargets(i)
       Next i
   End If
End Sub

