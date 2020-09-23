VERSION 5.00
Begin VB.Form FrmFlipX 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00270707&
   BorderStyle     =   0  'None
   Caption         =   "Flip Album"
   ClientHeight    =   13095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19590
   ControlBox      =   0   'False
   FillColor       =   &H00050505&
   ForeColor       =   &H00050505&
   LinkTopic       =   "Form1"
   ScaleHeight     =   873
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1306
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmFlipX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'************************************************************'
' Special Christmas Gift 2010                                '
' Christmas Flipping Book                                     '
' FlipPages  with tansparent background and background music '
'************************************************************'

' ****************************************************
' FlipPages   Animate a page or picture being flipped.
' Flip Perspective and sine wave -> Page Curve
'
' Extra feature : Tansparent background, page shadow and background music
'                Keypress to flip left <-> right
'                <Sapce bar> to toggle stop / start flip
'                <Esc> to exit
' ****************************************************
' mailto: tmax_visiber@yahoo.com


Private Type PageL      'Page parameter
        Left As Long
        Top As Long
        Width As Long
        Height As Long
End Type

' variant use by FlipPage
Dim R As Long                       ' Radius of Page (measure from center page bottom to the top left or top right)
Dim radY As Long                    ' Y pos of radius
Dim dw As Long                      ' Increase in width
Dim StartX, StartY As Long          ' Start of X position
Dim StartHeight, EndHeight As Long  ' Start of Y position
Dim OutWidth As Long                ' Output width
Dim OutYOffset As Long              ' Output Y position offset from origin

Dim Page As PageL           ' Page Display
Dim AutoFlip As Boolean     ' Auto flip pages
Dim Dx As Long              ' Differ in X direction -> for auto direction flippage
Dim AppPath  As String      ' Application path
Dim ImagePath As String     ' photo's path
Dim ResourcePath As String  ' Resources
Dim CurrentPhoto As Integer ' Current photo display

' Runtime control object
Dim File1 As FileListBox
Dim Pic1 As PictureBox
Dim Pic2 As PictureBox
Dim PicShadow(1) As cImgPng
Dim PicCover As cImgPng
Dim PicBorder As cImgPng
Dim cpng As cImgPng

' **********************************
' Keypress to flip Left < - > Right
' **********************************
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Ready2Exit
If KeyCode = vbKeyLeft Then R2L
If KeyCode = vbKeyRight Then L2R
If KeyCode = vbKeySpace Then AutoFlip = Not AutoFlip: FlipPage
End Sub

' *********************************************
' Loading all the runtime controls & parameters
' *********************************************
Private Sub Form_Load()
AppPath = App.Path & IIf(Right(App.Path, 1) <> "\", "\", "")
ImagePath = AppPath & "a_christmas\" '"Images\"    'Images folder
ResourcePath = AppPath & "Resource\"  'Resource folder
mciPlay ResourcePath & "jingle_bells.mid"
'PlayWave ResourcePath & "jingle_bells.mp3"
Set File1 = Me.Controls.Add("VB.filelistbox", "File1")
File1.Path = ImagePath
File1.Pattern = "*.jpg"
Set Pic1 = Me.Controls.Add("VB.PictureBox", "pic1")
Pic1.AutoRedraw = True
Pic1.AutoSize = True
Pic1.ScaleMode = 3
Set Pic2 = Me.Controls.Add("VB.PictureBox", "pic2")
Pic2.AutoRedraw = True
Pic2.AutoSize = True
Pic2.ScaleMode = 3
Pic1.Picture = LoadPicture(ImagePath & File1.List(CurrentPhoto))
Pic2.Picture = LoadPicture(ImagePath & File1.List(CurrentPhoto + 1))

Set PicShadow(0) = New cImgPng
PicShadow(0).Load ResourcePath & "shadow_R1.png"
Set PicShadow(1) = New cImgPng
PicShadow(1).Load ResourcePath & "shadow_L1.png"
Set PicBorder = New cImgPng
PicBorder.Load ResourcePath & "border.png"
Set PicCover = New cImgPng
PicCover.Load ResourcePath & "PageCover.png"
AutoFlip = True 'False
CurrentPhoto = 0
End Sub

' *********
' Drag Form
' *********
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then DragForm Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Ready2Exit
End Sub

Private Sub Form_Resize()
On Error Resume Next
SetTransparent Me.Hwnd, Me.BackColor   '&H00270707&
Me.Width = 3 / 2 * Me.Height
Page.Width = 900                  ' Preset to 900
Page.Height = 2 / 3 * Page.Width  ' for 4R Landscape photo (4" x 6" ) 2/3 ratio
Page.Left = (Me.ScaleWidth - Page.Width) / 2
Page.Top = (Me.ScaleHeight - Page.Height) / 2 + 80

PicCover.StretchDC Me.Hdc, Page.Left - 10, Page.Top, Page.Width + 20, Page.Height + 10
SetStretchBltMode Me.Hdc, COLORONCOLOR
StretchBlt Me.Hdc, Page.Left, Page.Top, Page.Width, Page.Height, Pic1.Hdc, 0, 0, Pic1.ScaleWidth, Pic1.ScaleHeight, vbSrcCopy
PicBorder.StretchDC Me.Hdc, Page.Left, Page.Top, Page.Width, Page.Height
Me.Refresh
R = 5 / 6 * Page.Width            'R = Sqr(Page.Height ^2 + (Page.Width / 2) ^2)
FlipPage
End Sub

Private Sub Form_Unload(Cancel As Integer)
Ready2Exit
End Sub

' ***********************
' Flip from Right to Left
' ***********************
Private Sub R2L()
  Pic1.Picture = LoadPicture(ImagePath & File1.List(CurrentPhoto))
  CurrentPhoto = CurrentPhoto - 1
  If CurrentPhoto < 0 Then CurrentPhoto = File1.ListCount - 1
   Pic2.Picture = LoadPicture(ImagePath & File1.List(CurrentPhoto))
  Shadow

  For dw = Page.Width To 0 + Page.Width / 50 Step -20
    If dw >= Page.Width / 2 Then
        Blting True
    Else
        Blting False
    End If
    Delay 5
  Next dw

  Me.Cls
  PicCover.StretchDC Me.Hdc, Page.Left - 10, Page.Top, Page.Width + 20, Page.Height + 10
  SetStretchBltMode Me.Hdc, COLORONCOLOR
  Pic2.Picture = LoadPicture(ImagePath & File1.List(CurrentPhoto))
  StretchBlt Me.Hdc, Page.Left, Page.Top, Page.Width, Page.Height, Pic2.Hdc, 0, 0, Pic2.ScaleWidth, Pic2.ScaleHeight, vbSrcCopy
  PicBorder.StretchDC Me.Hdc, Page.Left, Page.Top, Page.Width, Page.Height
  Me.Refresh
End Sub

' ***********************
' Flip from Left to Right
' ***********************
Private Sub L2R()
  Pic2.Picture = LoadPicture(ImagePath & File1.List(CurrentPhoto))
  CurrentPhoto = CurrentPhoto + 1
  If CurrentPhoto > File1.ListCount - 1 Then CurrentPhoto = 0
  Pic1.Picture = LoadPicture(ImagePath & File1.List(CurrentPhoto))
  Shadow

  For dw = 0 To Page.Width - Page.Width / 50 Step 20
    If dw <= Page.Width / 2 Then
        Blting False
    Else
        Blting True
    End If
    Delay 5
  Next dw

  Me.Cls
  PicCover.StretchDC Me.Hdc, Page.Left - 10, Page.Top, Page.Width + 20, Page.Height + 10
  SetStretchBltMode Me.Hdc, COLORONCOLOR
  Pic1.Picture = LoadPicture(ImagePath & File1.List(CurrentPhoto))
  StretchBlt Me.Hdc, Page.Left, Page.Top, Page.Width, Page.Height, Pic1.Hdc, 0, 0, Pic1.ScaleWidth, Pic1.ScaleHeight, vbSrcCopy
  PicBorder.StretchDC Me.Hdc, Page.Left, Page.Top, Page.Width, Page.Height
  Me.Refresh
End Sub

' **********************************
' Compute the parameter for FlipSBlt
' **********************************
Sub Blting(Reverse As Boolean)
  radY = (R - Page.Height) * Sin((dw / Page.Width) * Pi)    ' radY = Sqr(R ^2 - (Page.Width / 2 - i)^2) - Page.Height
  StartX = Page.Left + dw
  StartY = Page.Top - radY
  StartHeight = Page.Height
  EndHeight = Page.Height
  OutWidth = (Page.Width / 2) - dw
  OutYOffset = radY
  Me.Cls
  PicCover.StretchDC Me.Hdc, Page.Left - 10, Page.Top, Page.Width + 20, Page.Height + 10
  SetStretchBltMode Me.Hdc, COLORONCOLOR
  StretchBlt Me.Hdc, Page.Left, Page.Top, Page.Width / 2, Page.Height, Pic1.Hdc, 0, 0, Pic1.ScaleWidth / 2, Pic1.ScaleHeight, vbSrcCopy
  StretchBlt Me.Hdc, Page.Left + Page.Width / 2, Page.Top, Page.Width / 2, Page.Height, Pic2.Hdc, Pic2.ScaleWidth / 2, 0, Pic2.ScaleWidth / 2, Pic2.ScaleHeight, vbSrcCopy
  If Not Reverse Then
    Call FlipSBlt(Me.Hdc, StartX, StartY, OutWidth, StartHeight, EndHeight, OutYOffset, Pic2.Hdc, Pic2.ScaleWidth / 2, Pic2.ScaleHeight, False)
  Else
    Call FlipSBlt(Me.Hdc, StartX, StartY, OutWidth, StartHeight, EndHeight, OutYOffset, Pic1.Hdc, Pic1.ScaleWidth / 2, Pic1.ScaleHeight, True)
  End If
  Me.Refresh
End Sub

' *******************************
' Perspective Blt with sine curve
' *******************************
Sub FlipSBlt(ByVal outDC As Long, ByVal outX As Long, ByVal outY As Long, _
    ByVal OutWidth As Long, ByVal outStartHeight As Long, ByVal outEndHeight As Long, _
    ByVal outYOff As Long, ByVal inDC As Long, ByVal inWidth As Long, ByVal inHeight As Long, Optional Reverse As Boolean = False)
  Dim loopx As Long
  Dim InterpPos As Single
  Dim InterpH As Long
  Dim StartLoop As Long
  Dim EndLoop As Long
  Dim rady1 As Long
  If OutWidth = 0 Then Exit Sub
  StartLoop = 0
  EndLoop = OutWidth
  If OutWidth < 0 Then
    StartLoop = OutWidth
    EndLoop = 0
  End If
  SetStretchBltMode outDC, COLORONCOLOR
  For loopx = StartLoop To EndLoop
    InterpPos = loopx / OutWidth
    InterpH = InterpPos * (outEndHeight - outStartHeight)
    rady1 = outEndHeight / 20 * Sin((InterpPos) * 3.14159)
    If Not Reverse Then
      StretchBlt outDC, loopx + outX, outY + (InterpPos * outYOff) - rady1, 1, outStartHeight + InterpH, inDC, InterpPos * inWidth, 0, 1, inHeight, vbSrcCopy
    Else
      StretchBlt outDC, loopx + outX, outY + (InterpPos * outYOff) - rady1, 1, outStartHeight + InterpH, inDC, (2 - InterpPos) * inWidth, 0, 1, inHeight, vbSrcCopy
    End If
  Next loopx
End Sub

' ********
' AutoFlip
' ********
Sub FlipPage()
  Do While AutoFlip
    L2R
    Delay 1200
  Loop
End Sub

' **********
' Time delay
' **********
Sub Delay(tSet As Long)
  Dim tStart, tEnd As Long
  tStart = GetTickCount
  Do While tEnd < tSet
    tEnd = GetTickCount - tStart
    DoEvents
  Loop
End Sub

' ***********************
' SetShadow with Png file
' ***********************
Sub Shadow()
  PicShadow(0).StretchDC Pic1.Hdc, 0, 0, Pic1.ScaleWidth, Pic1.ScaleHeight
  PicShadow(1).StretchDC Pic2.Hdc, 0, 0, Pic2.ScaleWidth, Pic2.ScaleHeight
  PicBorder.StretchDC Pic1.Hdc, 0, 0, Pic1.ScaleWidth, Pic1.ScaleHeight
  PicBorder.StretchDC Pic2.Hdc, 0, 0, Pic2.ScaleWidth, Pic2.ScaleHeight
End Sub

' **************
' Clear All data
' **************
Sub Ready2Exit()
  mciStop
  'StopWave
  Set PicShadow(0) = Nothing
  Set PicShadow(1) = Nothing
  Set PicCover = Nothing
  Set PicBorder = Nothing
  End
End Sub
