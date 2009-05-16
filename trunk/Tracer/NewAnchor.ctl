VERSION 5.00
Begin VB.UserControl TagAnchor 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   495
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   495
   ScaleWidth      =   495
End
Attribute VB_Name = "TagAnchor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim LastHeight As Long, LastWidth As Long
Dim MinHeight As Long, MinWidth As Long
'********************************************
' DoResize checks the control tag and resize
'it.
'
' HeightChange and WidthChange are both save
'changing in height or width
'
' MinWidth and MinHeight save the first value
'of form and don't let form to get smaller
'
' Use Tag this way: Left,Top,Right and Bottom
' If U want to resize your control from all 4 ways
'use Tag this way: TTTT
'********************************************
Public Sub DoResize()
    Dim HeightChange As Long, WidthChange As Long
    Dim Frm As Object
    Dim Tg As String
    Dim i As Integer
    
    ' Set The Parent of Control
    Set Frm = Extender.Parent
    
    ' Exit sub on Minimize
    If Frm.WindowState = vbMinimized Then Exit Sub
    
    ' Check the form for Min Values
    If Frm.Height <= MinHeight Then Frm.Height = MinHeight
    If Frm.Width <= MinWidth Then Frm.Width = MinWidth
    
    'Calculate the Changes
    HeightChange = Frm.Height - LastHeight
    WidthChange = Frm.Width - LastWidth
    'If this is not the first time of resize
    If LastHeight <> 0 And LastWidth <> 0 Then
        For i = 0 To Frm.Controls.Count - 1
            
            Tg = Frm.Controls(i).Tag
            'Checking Tag
            If Right(Tg, 1) = "T" Then
                If Mid(Tg, 2, 1) = "T" Then
                    Frm.Controls(i).Height = Frm.Controls(i).Height + HeightChange
                Else
                    Frm.Controls(i).Top = Frm.Controls(i).Top + HeightChange
                End If
            End If
            
            If Mid(Tg, 3, 1) = "T" Then
                If Left(Tg, 1) = "T" Then
                    Frm.Controls(i).Width = Frm.Controls(i).Width + WidthChange
                Else
                    Frm.Controls(i).Left = Frm.Controls(i).Left + WidthChange
                End If
            End If
        Next i
    Else
        'This is the first Resize
        MinHeight = Frm.Height
        MinWidth = Frm.Width
    End If
    'Save Last values
    LastHeight = Frm.Height
    LastWidth = Frm.Width
End Sub
Private Sub UserControl_Resize()
   Width = 480
   Height = 465
End Sub
