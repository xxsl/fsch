VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Event onBuildStart()
Public Event onBuildProgess(data As String)
Public Event onBuildSuccess()
Public Event onBuildError()


Private WithEvents Ant As RedirectLib.Application
Attribute Ant.VB_VarHelpID = -1

Private BuildSuccessfull As Boolean

Private Sub Class_Initialize()
    Set Ant = New RedirectLib.Application
    Ant.BufferSize = 8192
End Sub

Public Sub Run(ByRef Target As AntTarget)
    Dim Name As String
    BuildSuccessfull = False
    Name = Target.getCommand
    Ant.Name = Name
    
    Select Case Ant.Start
       Case laAlreadyRunning:
            RaiseEvent onBuildProgess("Ant is already running")
       Case laWindowsError
            RaiseEvent onBuildProgess("Launch error: " & CStr(Ant.LastErrorNumber) & vbCrLf & Name)
       Case laOk
            RaiseEvent onBuildStart
    End Select
End Sub


Private Sub ant_DataReceived(ByVal sData As String)
    RaiseEvent onBuildProgess(sData)
    isSuccess sData
End Sub

Public Function isSuccess(data As String) As Boolean
    If (InStr(1, data, "BUILD SUCCESSFUL") > 0) Then
        BuildSuccessfull = True
    End If
End Function

Private Sub ant_ProcessEnded()
    If (BuildSuccessfull) Then
        RaiseEvent onBuildSuccess
    Else
        RaiseEvent onBuildError
    End If
End Sub


