VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSocketData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private bytes()   As Byte

Private m_locked  As Boolean

Private m_counter As Long

Private m_size    As Long

Public Sub append(ByRef data As JBSOCKETSERVERLib.IData)

    Dim buffer() As Byte

    Dim length   As Long
    
    buffer = data.Read
    length = UBound(buffer) + 1
    
    ReDim Preserve bytes(m_size + length - 1)
    CopyMemory bytes(m_size), buffer(0), length
    
    m_size = m_size + length
End Sub

Public Property Get locked() As Boolean
    locked = m_locked
End Property

Public Property Get size() As Long
    size = m_size
End Property

Public Function increment() As Long
    m_counter = m_counter + 1
    increment = m_counter
End Function

Public Function getBytes() As Byte()
    getBytes = bytes
End Function

Public Function PeekData(ByVal size As Long) As Byte()

    Dim buffer() As Byte

    ReDim Preserve buffer(size - 1)
    
    CopyMemory buffer(0), bytes(0), size
    PeekData = buffer
End Function

Public Function ReadData(ByVal size As Long) As Byte()

    Dim buffer()   As Byte

    Dim newBytes() As Byte
    
    buffer = PeekData(size)
    
    If (m_size - size > 0) Then
        ReDim newBytes(m_size - size - 1)
        CopyMemory newBytes(0), bytes(size), m_size - size
    End If
    
    bytes = newBytes
    
    m_size = m_size - size
    
    ReadData = buffer
End Function
