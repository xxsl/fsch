VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTarget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public fCommand As String
Public fName As String
Public fSource As String
Public fOutput As String
Public fLibraries As String
Public fDebug As String
Public fContext As String
Public fServices As String


Public Function getExecCommand() As String
    Dim result As String
    result = fCommand + " " + fSource + " " + Add(OUTPUT_PATH, fOutput) + _
                Add(INCLUDE_LIBRARIES, fLibraries) + Add(SERVICES, fServices) + _
                Add(COMPILE_DEBUG, fDebug)
End Function

Private Function Add(property As String, arg As String)
    If (Len(Trim(arg)) > 0) Then
        Add = property + arg + " "
            Else
        Add = " "
    End If
End Function
