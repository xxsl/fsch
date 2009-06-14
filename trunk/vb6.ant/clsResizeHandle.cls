VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsResizeHandle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***********************************************************************************
'* nimrod97@gmail.com                                                              *
'* Project homepage http://code.google.com/p/fsch/                                 *
'* Adobe Flex Compiler Shell wrapper                                               *
'* 2008                                                                            *
'***********************************************************************************

Option Explicit

Private WithEvents mainform As Form
Attribute mainform.VB_VarHelpID = -1

Private WithEvents frmFcsh As PictureBox
Attribute frmFcsh.VB_VarHelpID = -1

Private WithEvents SSTab1 As SSTab
Attribute SSTab1.VB_VarHelpID = -1

Private WithEvents frmTargets As PictureBox
Attribute frmTargets.VB_VarHelpID = -1



Public Sub setup(appForm As Form)
    Set mainform = appForm
    
    Set frmFcsh = mainform.frmFcsh
    
    Set SSTab1 = mainform.SSTab
    
    Set frmTargets = mainform.frmTargets
End Sub


'--------------------------
'Force resize on tab select
'--------------------------
Private Sub SSTab1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmFcsh_Resize
End Sub


'------------------
'Top section resize
'------------------
Private Sub frmFcsh_Resize()
    Dim listHeight As Long
    
    listHeight = mainform.frmFcsh.Width - mainform.SSTab.Left * 2
    If (listHeight > 0) Then
        mainform.SSTab.Width = listHeight
    End If
    listHeight = mainform.frmFcsh.Height - mainform.SSTab.Top - 150
    If (listHeight > 0) Then
        mainform.SSTab.Height = listHeight
    End If
    
    
    listHeight = mainform.SSTab.Width
    
    If (mainform.SSTab.Tab = 0) Then
        mainform.rtbLog.Width = listHeight - mainform.rtbLog.Left * 2
    End If
    
    If (mainform.SSTab.Tab = 1) Then
        mainform.rtbError.Width = listHeight - mainform.rtbError.Left * 2
    End If
    
    If (mainform.SSTab.Tab = 2) Then
        mainform.rtbWarn.Width = listHeight - mainform.rtbWarn.Left * 2
    End If
    
    listHeight = mainform.SSTab.Height - mainform.rtbLog.Top - 600
    
    If (listHeight > 0 And mainform.SSTab.Tab = 0) Then
        mainform.rtbLog.Height = listHeight
    End If
    
    listHeight = mainform.SSTab.Height - mainform.rtbError.Top - 600
    If (listHeight > 0 And mainform.SSTab.Tab = 1) Then
        mainform.rtbError.Height = listHeight
    End If
    
    listHeight = mainform.SSTab.Height - mainform.rtbWarn.Top - 600
    If (listHeight > 0 And mainform.SSTab.Tab = 2) Then
        mainform.rtbWarn.Height = listHeight
    End If
   
    listHeight = mainform.rtbLog.Top + mainform.rtbLog.Height + 100
    If (mainform.SSTab.Tab = 0) Then
        mainform.cmdClearLog.Top = listHeight
        mainform.cmdClearLog.Left = mainform.rtbLog.Left + mainform.rtbLog.Width - mainform.cmdClearLog.Width
    End If
    
    listHeight = mainform.rtbError.Top + mainform.rtbError.Height + 100
    If (mainform.SSTab.Tab = 1) Then
        mainform.cmdClearErr.Top = listHeight
        mainform.cmdClearErr.Left = mainform.rtbError.Left + mainform.rtbError.Width - mainform.cmdClearErr.Width
    End If
    
    listHeight = mainform.rtbWarn.Top + mainform.rtbWarn.Height + 100
    If (mainform.SSTab.Tab = 2) Then
        mainform.cmdClearWarn.Top = listHeight
        mainform.cmdClearWarn.Left = mainform.rtbWarn.Left + mainform.rtbWarn.Width - mainform.cmdClearWarn.Width
    End If
End Sub



'---------------------
'Resize bottom section
'---------------------
Private Sub frmTargets_Resize()
    mainform.lstTargets.Width = mainform.frmTargets.Width - mainform.lstTargets.Left * 2
    
    Dim listHeight As Long
    
    listHeight = mainform.frmTargets.Height - mainform.lstTargets.Top - 700
    
    If (listHeight > 0) Then
        mainform.lstTargets.Height = listHeight
    End If
    
    listHeight = mainform.frmTargets.Height - 600
    mainform.cmdClear.Top = listHeight
    mainform.cmdHide.Top = listHeight
    mainform.cmdRecompile.Top = listHeight
    
    mainform.cmdHide.Left = mainform.lstTargets.Left + mainform.lstTargets.Width - mainform.cmdHide.Width
End Sub



'-----------
'Form resize
'-----------
Private Sub mainform_Resize()
    mainform.ctlSplitterEx1.Width = mainform.Width - 120
    mainform.ctlSplitterEx1.Height = mainform.Height - 410
End Sub
