VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form MainForm 
   Caption         =   "Compiler"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock Controller 
      Left            =   3360
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "localhost"
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3360
      Top             =   2040
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************
'* nimrod97@gmail.com                                                              *
'* Project homepage http://code.google.com/p/fsch/                                 *
'* Adobe Flex Compiler Shell wrapper                                               *
'* 2008                                                                            *
'***********************************************************************************

Option Explicit

Public config As New clsConfiguration

Private ms As Long

Const BUILD_FAILED As String = "Build failed"
Const BUILD_SUCESSFULL As String = "Build successfull"

Private Sub Form_Load()
    If (Len(Trim(Command)) > 0) Then
        config.Load
        Dim result As Long
        result = WriteStdOut("fcsh remote compiler [" & App.Revision & "]" & vbCrLf)
        If (result <> 0) Then
            MsgBox "Relink executable to support stdOut"
            End
        End If
        Timer1.Enabled = True
        Controller.RemotePort = config.SERVER_PORT
        Controller.LocalPort = 0
        Controller.Connect
    Else
        result = WriteStdOut("fcsh remote compiler [" & App.Revision & "]" & vbCrLf)
        If (result <> 0) Then
            MsgBox "Relink executable to support stdOut"
            End
        End If
        WriteStdOut "No target" + vbCrLf
        End
    End If
End Sub

Private Sub Timer1_Timer()
    ms = ms + 1
End Sub

'***********************************************************************************************
'Sockets Client
'***********************************************************************************************
Private Sub Controller_Close()
    End
End Sub

Private Sub Controller_Connect()
    Controller.SendData Command
End Sub

Private Sub Controller_DataArrival(ByVal bytesTotal As Long)
    Dim s As String
    On Error Resume Next
    Controller.GetData s, vbString, bytesTotal
    responce = responce + s
    WriteStdOut s
    If (InStr(1, responce, BUILD_FAILED) > 0 Or InStr(1, responce, BUILD_SUCESSFULL) > 0) Then
        WriteStdOut vbCrLf
        WriteStdOut "Build time: " & ms * 100 & " ms" & vbCrLf
        End
    End If
    If (Err.Number <> 0) Then
        WriteStdOut "Error: " & Err.Number & " " & Err.Description & vbCrLf
        If (Err.Number = 10054) Then
            WriteStdOut "fcsh is stopped"
        End If
        Err.Clear
        WriteStdOut vbCrLf
        WriteStdOut "Build time: " & ms * 100 & " ms" & vbCrLf
        End
    End If
End Sub

Private Sub Controller_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    WriteStdOut Description + vbCrLf + BUILD_FAILED
    Controller.Close
    End
End Sub
