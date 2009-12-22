Attribute VB_Name = "modManifest"
Option Explicit

Public Function CreateManifest() As Boolean

    Dim f   As TextStream

    Dim fso As New FileSystemObject

    If fso.FileExists(App.Path & App.EXEName & ".exe" & ".manifest") = False Then
        fso.CreateTextFile App.Path & App.EXEName & ".exe" & ".manifest", False
        Set f = fso.OpenTextFile(App.Path & App.EXEName & ".exe" & ".manifest", ForAppending, TristateFalse)
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

