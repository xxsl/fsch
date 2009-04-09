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


Private WithEvents ant As RedirectLib.Application
Attribute ant.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set ant = New RedirectLib.Application
    ant.BufferSize = 8192
End Sub

Public Sub Run(ByRef target As AntTarget)
    ant.Name = "ant " & target.build.BuildFile & " " & target.Name
End Sub


Private Sub ant_DataReceived(ByVal sData As String)
    'todo parse errors
    RaiseEvent onBuildProgess(sData)
End Sub


Private Sub ant_ProcessEnded()
    'todo
    RaiseEvent onBuildSuccess
End Sub


