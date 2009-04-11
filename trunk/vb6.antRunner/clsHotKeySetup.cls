VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHotKeySetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub SetupKey(ByRef config As clsConfig, ByRef control As VBHotKey)
Dim MODIFYER As Long
MODIFYER = config.MODIFYER

control.VKey = config.KEY

control.AltKey = (MODIFYER = echkModifierKeys.HOTKEYF_ALT) Or _
                 (MODIFYER = echkModifierKeys.HOTKEYF_ALTSHIFT) Or _
                 (MODIFYER = echkModifierKeys.HOTKEYF_CONTROLALT) Or _
                 (MODIFYER = echkModifierKeys.HOTKEYF_CONTROLALTSHIFT)

control.CtrlKey = (MODIFYER = echkModifierKeys.HOTKEYF_CONTROL) Or _
                  (MODIFYER = echkModifierKeys.HOTKEYF_CONTROLALT) Or _
                  (MODIFYER = echkModifierKeys.HOTKEYF_CONTROLALTSHIFT) Or _
                  (MODIFYER = echkModifierKeys.HOTKEYF_SHIFTCONTROL)
                  
control.ShiftKey = (MODIFYER = echkModifierKeys.HOTKEYF_CONTROLALTSHIFT) Or _
                   (MODIFYER = echkModifierKeys.HOTKEYF_ALTSHIFT) Or _
                   (MODIFYER = echkModifierKeys.HOTKEYF_SHIFT) Or _
                   (MODIFYER = echkModifierKeys.HOTKEYF_SHIFTCONTROL)
                   
control.ENABLED = config.HotKeyEnabled
End Sub
