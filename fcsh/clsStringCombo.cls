VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStringCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private combo As ComboBox
Private separator As String

Public Sub Bind(ByRef cmb As ComboBox, sep As String)
    separator = sep
    Set combo = cmb
End Sub



Public Property Get Property() As String
    Dim i As Long, value As String
    
    For i = 0 To combo.ListCount - 1
        value = value + combo.List(i) + separator
    Next i
    If (Len(value) > 0) Then
        value = Mid(value, 1, Len(value) - 1)
    End If
    Property = value
End Property

Public Property Let Property(ByVal vNewValue As String)
    Dim values() As String
    
    combo.Clear
    
    values = Split(vNewValue, separator)
    
    Dim i As Long
    For i = LBound(values) To UBound(values)
        combo.AddItem values(i)
    Next i
    If (combo.ListCount >= 1) Then
       combo.ListIndex = combo.ListCount - 1
    End If
End Property

Public Sub Add(item As String)
    Dim i As Long, hasItem As Boolean
    
    For i = 0 To combo.ListCount - 1
        If (LCase(Trim(combo.List(i))) = LCase(Trim(item))) Then
            hasItem = True
            Exit For
        End If
    Next i
    If (Not hasItem) Then
       combo.AddItem item
       If (combo.ListCount >= 1) Then
           combo.ListIndex = combo.ListCount - 1
       End If
    End If
End Sub

Public Sub Remove(item As String)
    Dim i As Long
    
    For i = 0 To combo.ListCount - 1
        If (LCase(Trim(combo.List(i))) = LCase(Trim(item))) Then
            combo.RemoveItem i
            Exit For
        End If
    Next i
    
    If (combo.ListCount >= 1) Then
        combo.ListIndex = combo.ListCount - 1
    End If
End Sub

Public Sub Reset()
    combo.Clear
End Sub

