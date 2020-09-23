Attribute VB_Name = "ModAll"
Option Explicit

' Wav player Api
Const SND_NOWAIT = &H2000        'don't wait if the driver is busy
Const SND_ASYNC = &H1            'Play asynchronously
Const SND_NODEFAULT = &H2        'silence not default, if sound not found
Const SND_MEMORY = &H4           'lpszSoundName points to a memory file
Const SND_LOOP = &H8             'loop the sound until next sndPlaySound
Const SND_NOSTOP = &H10          'don't stop any currently playing sound
Const SND_SYNC = &H0             'play synchronously (default)
Const SND_FILENAME& = &H20000
Const SND_ALIAS& = &H10000

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

' Midi player
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
    
' SetTranparent Api
Const LW_KEY = &H1
Const G_E = (-20)
Const W_E = &H80000
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal Hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

' DragForm Api
Const WM_NCLBUTTONDOWN = &HA1
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

' Paint stretch gdi
Public Const COLORONCOLOR = 3           '**IMPORTANT **  settting for StretchBlt
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal Hdc As Long, ByVal nStretchMode As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" ( _
      ByVal Hdc As Long, _
      ByVal X As Long, _
      ByVal Y As Long, _
      ByVal nWidth As Long, _
      ByVal nHeight As Long, _
      ByVal hSrcDC As Long, _
      ByVal xSrc As Long, _
      ByVal ySrc As Long, _
      ByVal nSrcWidth As Long, _
      ByVal nSrcHeight As Long, _
      ByVal dwRop As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Const Pi As Single = 3.141593

'**********
' Drag Form
'**********
Function DragForm(frm As Form)
  Dim ret As Long
  ret = ReleaseCapture()
  ret = SendMessage(frm.Hwnd, WM_NCLBUTTONDOWN, 2&, 0&)
End Function

'***************************
' Set background transparent
'***************************
Sub SetTransparent(Hwnd As Long, TransColor As Long)
  On Error Resume Next
  Dim ret As Long
  ret = GetWindowLong(Hwnd, G_E)
  ret = ret Or W_E
  SetWindowLong Hwnd, G_E, ret
  SetLayeredWindowAttributes Hwnd, TransColor, 0, LW_KEY
End Sub

'****************
' Wav file player
'****************
Sub PlayWave(Filename As String)
  sndPlaySound Filename, SND_ASYNC Or SND_NODEFAULT Or SND_NOSTOP Or SND_NOWAIT Or SND_LOOP
End Sub

Sub StopWave()
  sndPlaySound vbNullString, SND_ASYNC
End Sub

'****************
'Midi file player
'****************
Sub mciPlay(Filename As String)
  Dim sFile As String
  Dim sShortFile As String * 67
  Dim lResult As Long
  Dim errstr As String * 255
  Dim ret As Long
    
  lResult = GetShortPathName(Filename, sShortFile, Len(sShortFile))
  sFile = Left$(sShortFile, lResult)
  ret = mciSendString("open " & sFile & " type sequencer alias Midi2010", ByVal 0&, 0, 0)
  If ret <> 0 Then
    ret = mciGetErrorString(ret, errstr, 255)
    MsgBox errstr
    Exit Sub
  End If
  mciSendString "play Midi2010", ByVal 0&, 0, 0
End Sub

Sub mciStop()
   mciSendString "stop Midi2010", 0, 0, 0
   mciSendString "close Midi2010", 0, 0, 0
End Sub


