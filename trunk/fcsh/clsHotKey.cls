VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHotKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private prefs As ARINIManager
Private section As String


Public Sub init(ByRef manager As ARINIManager, name As String)
    Set prefs = manager
    section = name
End Sub


Public Property Let ENABLED(ByVal vNewValue As Long)
    prefs.WriteValue section, "key_enabled", Str(vNewValue)
End Property

'use hot keys
Public Property Get ENABLED() As Long
    ENABLED = Val(prefs.GetValue(section, "key_enabled", "0"))
End Property

Public Property Let KEY(ByVal vNewValue As Long)
    prefs.WriteValue section, "key", Str(vNewValue)
End Property

'use hot keys
Public Property Get KEY() As Long
    KEY = Val(prefs.GetValue(section, "key", "0"))
End Property

Public Property Let MODIFYER(ByVal vNewValue As Long)
    prefs.WriteValue section, "key_modifyer", Str(vNewValue)
End Property

'use hot keys
Public Property Get MODIFYER() As Long
    MODIFYER = Val(prefs.GetValue(section, "key_modifyer", "0"))
End Property
