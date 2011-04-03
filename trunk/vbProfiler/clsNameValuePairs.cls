VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNameValuePairs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strings() As String

Private ids()  As Long

Private m_size As Long


Public Sub setData(index As Long, name As String, Optional id As Long = -1)
    strings(index) = name
    If (id <> -1) Then
        ids(index) = id
    End If
End Sub


Public Function getName(index As Long) As String
    getName = strings(index)
End Function

Public Function getId(index As Long) As Long
    getId = ids(index)
End Function

Public Function getSize() As Long
    getSize = m_size
End Function

Public Function setSize(value As Long)
    m_size = value
    ReDim strings(m_size - 1)
    ReDim ids(m_size - 1)
End Function
