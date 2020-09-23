Attribute VB_Name = "DXsnakesMAIN"
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Public Frames As Integer

Public Str As String

Public Diff As Byte

Public Type tSnakePart
  x As Integer
  y As Integer
End Type
  
Public Type tSnake
  Part() As tSnakePart
  xm As Integer
  ym As Integer
  Facing As Byte
End Type

Public Snake As tSnake

Public Type tApple
  x As Integer
  y As Integer
End Type

Public Apple As tApple

Public Const HEAD = 0

Public Const dUP = 1
Public Const dRIGHT = 2
Public Const dDOWN = 3
Public Const dLEFT = 4

Public GameOverReason As Byte

Public Const OK = 0
Public Const TANGLE = 1
Public Const EDGE = 2
Public Const EAT = 3

Public x As Integer
Public y As Integer
Public i As Integer
Public i2 As Integer

Public Score As Integer


Sub Main()
MsgBox "Welcome to DX Snakes by Simon Price! This game is like the one you find on mobile phones, except with a bit better graphics and a larger screen! Use the arrow keys to control your snake, and press Escape to quit the game. Have fun!", vbInformation, "DX Snakes by Simon Price"
Form1.MousePointer = vbHourglass
'convert jpegs to bitmaps
Form1.ConvertEm
'start up DX
Form1.CrankItUp
PlayAgain:
'set the screen res
ModDX7.SetDisplayMode 640, 480, 16
'load all the graphix
ModSurfaces.LoadAllPics

'start music
Form1.Play
'do intro screen
DoIntro
'load up the level
LoadLevel
'enter main game loop
MainGameLoop

If UserWantsToPlayAgain Then GoTo PlayAgain

SayGoodbye

'stop music
Form1.ShutYerNoise
'unload all the graphix
ModSurfaces.UnloadAllPics
'put screen back to norm
ModDX7.RestoreDisplayMode
'close DX
Form1.ShutItDown

End
End Sub

Sub DoIntro()

'show the intro screen
ModDX7.SetRect SrcRect, 0, 0, 640, 480
BackBuffer.BltFast 0, 0, Intro, SrcRect, DDBLTFAST_WAIT
View.Flip Nothing, DDFLIP_WAIT

'wait for user to press a valid key
Do

DoEvents
Select Case Form1.Key
  Case vbKey1 To vbKey9
    Diff = Form1.Key - vbKey0
    Exit Do
End Select

Loop

End Sub

Sub LoadLevel()

'set snake length
ReDim Snake.Part(0 To 110 - 10 * Diff)

'turn him right
TurnSnake dRIGHT

'position snake
For i = 0 To 110 - 10 * Diff
  Snake.Part(i).x = 250 - i * Diff
  Snake.Part(i).y = 240
Next

'set the apple
MoveApple

'reset the score
Score = 0

'new life
GameOverReason = 0

End Sub

Sub MainGameLoop()

Do
DoEvents

Frames = Frames + 1

Select Case MoveSnake
Case OK
Case EDGE
GameOverReason = EDGE
Case TANGLE
GameOverReason = TANGLE
Case EAT
  DoEvents
  MoveApple
End Select

DoEvents

'draw everything

'draw the background
ModDX7.SetRect SrcRect, 0, 0, 640, 480
BackBuffer.BltFast 0, 0, Background, SrcRect, DDBLTFAST_WAIT

'draw the snake
DrawSnake

'draw the score
BackBuffer.DrawText 100, 100, "Score = " & Score, True

'draw the apple
ModDX7.SetRect SrcRect, 150, 0, 30, 30
BackBuffer.BltFast Apple.x - 15, Apple.y - 15, Sprites, SrcRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY

DoEvents

'copy everything from backbuffer into sight
View.Flip Nothing, DDFLIP_WAIT

Loop Until GameOverReason

End Sub

Function UserWantsToPlayAgain() As Boolean
'unload all the graphix
ModSurfaces.UnloadAllPics
'close DX
Form1.ShutItDown

'construct a few sentences based on your performance
Select Case GameOverReason
Case EDGE
Str = "You hit the edge of the screen!"
Case TANGLE
Str = "You got tangled up!"
End Select

Str = Str & " You collected " & Score \ Diff & " apples and scored " & Score & " points."

Select Case Score
Case 0 To 30
Str = Str & " That's a rubbish score!"
Case 30 To 60
Str = Str & " That's not very good."
Case 60 To 100
Str = Str & " That's not a bad score."
Case 100 To 150
Str = Str & " That's a good score."
Case 150 To 250
Str = Str & " That's a great score!"
Case 250 To 500
Str = Str & " That's an amazing score!"
Case Is > 500
Str = Str & " That's an unbeleiveably high score!"
End Select

Str = Str & " Do you want to play again?"

If MsgBox(Str, vbQuestion + vbYesNo, "Game Over!") = vbYes Then UserWantsToPlayAgain = True

'start up DX
Form1.CrankItUp
'load all the graphix
ModSurfaces.LoadAllPics

End Function

Sub SayGoodbye()
'start up DX
Form1.CrankItUp
'set the screen res
ModDX7.SetDisplayMode 640, 480, 16
'load all the graphix
ModSurfaces.LoadAllPics
ModDX7.CreateSurfaceFromFile Intro, IntroDesc, App.Path & "\Graphix\Goodbye.bmp", 640, 480
ModDX7.SetRect SrcRect, 0, 0, 640, 480
BackBuffer.BltFast 0, 0, Intro, SrcRect, DDBLTFAST_WAIT
View.Flip Nothing, DDFLIP_WAIT
Form1.Key = 0
Do
DoEvents
Loop Until Form1.Key
End Sub

Function MoveSnake() As Byte

'shift each snake part along one
For i = UBound(Snake.Part) To 1 Step -1
  Snake.Part(i) = Snake.Part(i - 1)
Next

'move the head
Snake.Part(HEAD).x = Snake.Part(HEAD).x + Snake.xm
Snake.Part(HEAD).y = Snake.Part(HEAD).y + Snake.ym

'see if snake has crashed into edge of screen
Select Case Snake.Part(HEAD).x
Case 15 To 625
  Case Else
  MoveSnake = EDGE
End Select

Select Case Snake.Part(HEAD).y
Case 15 To 465
  Case Else
  MoveSnake = EDGE
End Select

'see if snake has crashed into itself
For i = 70 \ Diff To UBound(Snake.Part)
  Select Case Snake.Part(HEAD).x
    Case Snake.Part(i).x - 30 To Snake.Part(i).x + 30
  Select Case Snake.Part(HEAD).y
    Case Snake.Part(i).y - 30 To Snake.Part(i).y + 30
       'snake has hit itself
       MoveSnake = TANGLE
  End Select
  End Select
Next

'see if snake has got the apple
Select Case Snake.Part(HEAD).x
  Case Apple.x - 30 To Apple.x + 30
Select Case Snake.Part(HEAD).y
  Case Apple.y - 30 To Apple.y + 30
    'snake has got the apple
    MoveSnake = EAT
End Select
End Select
      
End Function

Sub DrawSnake()
'draw each bit of snake, from tail to head
ModDX7.SetRect SrcRect, 0, 0, 30, 30

For i = UBound(Snake.Part) To 1 Step -1 - 10 + Diff
  BackBuffer.BltFast Snake.Part(i).x - 15, Snake.Part(i).y - 15, Sprites, SrcRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
Next
  
'draw the head
ModDX7.SetRect SrcRect, Snake.Facing * 30, 0, 30, 30
BackBuffer.BltFast Snake.Part(HEAD).x - 15, Snake.Part(HEAD).y - 15, Sprites, SrcRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
End Sub

Sub TurnSnake(Dir As Byte)
Snake.Facing = Dir

Select Case Dir
Case dUP
  Snake.xm = 0
  Snake.ym = -Diff
Case dDOWN
  Snake.xm = 0
  Snake.ym = Diff
Case dLEFT
  Snake.xm = -Diff
  Snake.ym = 0
Case dRIGHT
  Snake.xm = Diff
  Snake.ym = 0
End Select

End Sub

Sub MoveApple()
'moves the apple to a new position
TryAgain:
Apple.x = Int(Rnd * 610) + 15
Apple.y = Int(Rnd * 450) + 15

'check it's not under the snake
For i = HEAD To UBound(Snake.Part)
  Select Case Snake.Part(i).x
    Case Apple.x - 30 To Apple.x + 30
      Select Case Snake.Part(i).y
        Case Apple.y - 30 To Apple.y + 30
           'apple is under snake so try again
           GoTo TryAgain
      End Select
  End Select
Next
      
'increase score for eating apple
Score = Score + Diff

'increase the length of the snake
ReDim Preserve Snake.Part(0 To UBound(Snake.Part) + Diff * 10)
'Snake.Part(UBound(Snake.Part)) = Snake.Part(UBound(Snake.Part) - 1)

End Sub

