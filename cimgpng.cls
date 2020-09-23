VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cImgPng"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type GdiplusStartupInput
   GdiplusVersion           As Long
   DebugEventCallback       As Long
   SuppressBackgroundThread As Long
   SuppressExternalCodecs   As Long
End Type
Private Declare Sub GdiplusShutdown Lib "gdiplus.dll" (ByVal Token As Long)
Private Declare Function GdipLoadImageFromFile Lib "gdiplus.dll" (ByVal Filename As Long, GpImage As Long) As Long
Private Declare Function GdiplusStartup Lib "gdiplus.dll" (Token As Long, gdipInput As GdiplusStartupInput, GdiplusStartupOutput As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus.dll" (ByVal Hdc As Long, GpGraphics As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal InterMode As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus.dll" (ByVal Graphics As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus.dll" (ByVal Image As Long) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus.dll" (ByVal Image As Long, Width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus.dll" (ByVal Image As Long, Height As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus.dll" ( _
      ByVal Graphics As Long, _
      ByVal GpImage As Long, _
      ByVal dstx As Long, _
      ByVal dsty As Long, _
      ByVal dstwidth As Long, _
      ByVal dstheight As Long, _
      ByVal srcx As Long, _
      ByVal srcy As Long, _
      ByVal srcwidth As Long, _
      ByVal srcheight As Long, _
      ByVal srcUnit As Long, _
      ByVal imageAttributes As Long, _
      ByVal callback As Long, _
      ByVal callbackData As Long) As Long

Private Const InterpolationModeHighQualityBicubic = 7
Private Const UnitPixel = 2

Private m_token As Long      'token to GDI engine
Private m_Image As Long      'original file loaded
Private m_Graphic As Long    'adjusted image
Private m_FileName As String
Private ImageWidth As Long
Private ImageHeight As Long

Private Sub Class_Initialize()
   'start the GDI engine
  On Error GoTo Handler
  Dim GpInput As GdiplusStartupInput

  GpInput.GdiplusVersion = 1
  GdiplusStartup m_token, GpInput, ByVal 0&
  Exit Sub
Handler:
  ' ReportError Err.Number, "Initialising GDI+", Err.Description
  Resume Next
End Sub

Private Sub Class_Terminate()
   If Initialised Then
      DisposeImage
      Call GdiplusShutdown(m_token)
   End If
End Sub

Private Function DisposeImage()
  GdipDisposeImage m_Image
  m_Image = 0
End Function

Private Property Get Initialised() As Boolean
   If m_token <> 0 Then
      Initialised = True
   End If
End Property

Public Function Load(sFileName As String) As Boolean
  Dim retval As Long
   DisposeImage
   m_FileName = sFileName
   retval = GdipLoadImageFromFile(StrPtr(sFileName), m_Image)
   If retval = 0 Then
    GdipGetImageWidth m_Image, ImageWidth
    GdipGetImageHeight m_Image, ImageHeight
    Load = True
   End If
End Function

Public Function StretchDC(ByVal Hdc As Long, X As Long, Y As Long, W As Long, h As Long)
   GdipCreateFromHDC Hdc, m_Graphic
   If W = 0 Then W = h * ImageWidth / ImageHeight
   If h = 0 Then h = W * ImageHeight / ImageWidth
   GdipSetInterpolationMode m_Graphic, InterpolationModeHighQualityBicubic
   GdipDrawImageRectRectI m_Graphic, m_Image, X, Y, W, h, 0, 0, ImageWidth, ImageHeight, UnitPixel, 0, 0, 0
   GdipDeleteGraphics m_Graphic
End Function
