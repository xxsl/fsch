VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFixedCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private combo As ComboBox

Public Sub Bind(ByRef cmb As ComboBox, ByRef items() As String)
    Set combo = cmb

    Dim i As Long
    For i = LBound(items) To UBound(items)
        combo.AddItem items(i)
    Next i
End Sub



Public Property Get Property() As String
    Property = combo.List(combo.ListIndex)
End Property

Public Property Let Property(ByVal vNewValue As String)
    Dim i As Long
    For i = 0 To combo.ListCount - 1
        If (LCase(Trim(combo.List(i))) = LCase(Trim(vNewValue))) Then
            combo.ListIndex = i
            Exit For
        End If
    Next i
End Property

Public Sub Reset()
    combo.ListIndex = -1
End Sub
