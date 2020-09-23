VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3468
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5064
   ForeColor       =   &H0000FF00&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   99  'Custom
   ScaleHeight     =   289
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   422
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer MusicTime 
      Interval        =   1000
      Left            =   2760
      Top             =   240
   End
   Begin MCI.MMControl MMControl1 
      Height          =   612
      Left            =   840
      TabIndex        =   2
      Top             =   2760
      Visible         =   0   'False
      Width           =   2832
      _ExtentX        =   4995
      _ExtentY        =   1080
      _Version        =   393216
      DeviceType      =   "Sequencer"
      FileName        =   ""
   End
   Begin VB.PictureBox SavePic 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1332
      Left            =   2280
      ScaleHeight     =   1332
      ScaleWidth      =   1452
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.PictureBox LoadPic 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1332
      Left            =   480
      ScaleHeight     =   1332
      ScaleWidth      =   1452
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Timer FramesT 
      Interval        =   1000
      Left            =   480
      Top             =   240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Key As Byte
Public Cur As Long
Public MTime As Byte

Sub CrankItUp()
Cur = ShowCursor(0)
ModDX7.Init Me.hwnd
End Sub

Sub ShutItDown()
ShowCursor Cur
ModDX7.EndIt Me.hwnd
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Key = KeyCode

Select Case KeyCode

Case vbKeyEscape
  ModSurfaces.UnloadAllPics
  ModDX7.RestoreDisplayMode
  Form1.ShutItDown
  End

Case vbKeyUp
  If Snake.Facing <> dDOWN Then TurnSnake dUP

Case vbKeyDown
  If Snake.Facing <> dUP Then TurnSnake dDOWN

Case vbKeyLeft
  If Snake.Facing <> dRIGHT Then TurnSnake dLEFT

Case vbKeyRight
  If Snake.Facing <> dLEFT Then TurnSnake dRIGHT

End Select

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Key = 0
End Sub
    
Private Sub Form_Load()
Move 0, 0, 640, 480
MMControl1.Filename = App.Path & "\Soundz\Music.mid"
End Sub

Private Sub FramesT_Timer()
Debug.Print "FPS = " & Frames

Frames = 0
End Sub

Sub ConvertEm()
LoadPic = LoadPicture(App.Path & "\Graphix\Background.jpg")
SavePic = LoadPic
SavePicture SavePic.Picture, App.Path & "\Graphix\Background.bmp"

LoadPic = LoadPicture(App.Path & "\Graphix\Intro.jpg")
SavePic = LoadPic
SavePicture SavePic.Picture, App.Path & "\Graphix\Intro.bmp"

LoadPic = LoadPicture(App.Path & "\Graphix\Goodbye.jpg")
SavePic = LoadPic
SavePicture SavePic.Picture, App.Path & "\Graphix\Goodbye.bmp"
End Sub

Sub Play()
MMControl1.Command = "Open"
MMControl1.Command = "Play"
MusicTime.Enabled = True
End Sub

Sub ShutYerNoise()
MMControl1.Command = "Stop"
MMControl1.Command = "Close"
MusicTime.Enabled = False
End Sub

Private Sub MusicTime_Timer()
MTime = MTime + 1
If MTime = 113 Then
  ShutYerNoise
  Play
  MTime = 0
End If
End Sub


