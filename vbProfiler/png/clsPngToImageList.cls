VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPngToImageList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***********************************************************************************
'* nimrod97@gmail.com                                                              *
'* Project homepage http://code.google.com/p/fsch/                                 *
'* Adobe Flex Compiler Shell wrapper                                               *
'* 2008                                                                            *
'***********************************************************************************


Private cImage As c32bppDIB
Private log As clsLog
Private picIcon As PictureBox
Private picClear As PictureBox
Private imgList As MSComctlLib.ImageList

Public Sub Initialize(ByRef icon As PictureBox, ByRef clear As PictureBox, ByRef pngList As MSComctlLib.ImageList, ByRef logger As clsLog)
    Set picIcon = icon
    Set picClear = clear
    Set imgList = pngList
    Set log = logger
End Sub

Public Sub LoadIcons(files As Collection)
    Dim item As Variant
    Dim path As String
    
    imgList.ListImages.clear
    
    For Each item In files
        path = CStr(item)
        If (FileExists(path)) Then
            picIcon.Picture = picClear.Picture
            cImage.InitializeDIB 16, 16
            cImage.LoadPicture_File path, 16, 16, False
            log.xDebug "Loading " + path + " :" + getResult
            cImage.Render picIcon.hDC
            cImage.DestroyDIB
            picIcon.Picture = picIcon.Image
            imgList.ListImages.Add , , picIcon.Image
        End If
    Next
End Sub

'class routines
Private Function getResult() As String
    Select Case cImage.ImageType ' want to know source image's format?
           Case imgNone, imgError:     getResult = "Image was not loaded"
           Case imgBitmap:             getResult = "Format: Standard Bitmap or JPG"
           Case imgEMF:                getResult = "Format: Extended Windows Metafile"
           Case imgWMF:                getResult = "Format: Standard Windows Metafile"
           Case imgIcon:               getResult = "Format: Standard Icon"
           Case imgBmpARGB:            getResult = "Format: 32bpp Bitmap with ARGB"
           Case imgBmpPARGB:           getResult = "Format: 32bpp Bitmap with pARGB"
           Case imgCursor:             getResult = "Format: Standard Cursor"
           Case imgCursorARGB:         getResult = "Format: Alpha Cursor"
           Case imgIconARGB:           getResult = "Format: Alpha Icon"
           Case imgPNG:                getResult = "Format: PNG"
           Case imgPNGicon:            getResult = "Format: PNG in Vista Icon"
           Case imgGIF
               If cImage.Alpha = True Then
                                       getResult = "Format: Transparent GIF"
               Else
                                       getResult = "Format: GIF"
               End If
           Case imgTGA
               If cImage.Alpha = True Then
                                       getResult = "Format: Transparent TGA"
               Else
                                       getResult = "Format: TGA (Targa)"
               End If
           Case Else:                  getResult = "..."
       End Select
End Function

Private Sub Class_Initialize()
    Set cImage = New c32bppDIB
End Sub

Private Sub Class_Terminate()
    Set cImage = Nothing
    Set picIcon = Nothing
    Set picClear = Nothing
    Set imgList = Nothing
End Sub
