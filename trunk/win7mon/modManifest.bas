Attribute VB_Name = "modManifest"
Option Explicit

Public Function CreateManifest() As Boolean
    Dim f   As TextStream
    Dim fso As New FileSystemObject
    Dim ApplicationPath As String

    ApplicationPath = GetPath & ".exe.manifest"
    
    If fso.FileExists(ApplicationPath) = False Then
        fso.CreateTextFile ApplicationPath, False
        Set f = fso.OpenTextFile(ApplicationPath, ForAppending, TristateFalse)
        f.Write LoadResString("101")
        f.Close

        DoEvents
        MsgBox "No manifest file found please restart the application", vbInformation + vbOKOnly, "Load ManifestFile"

        End

        CreateManifest = True
    Else
        CreateManifest = False
    End If

End Function



