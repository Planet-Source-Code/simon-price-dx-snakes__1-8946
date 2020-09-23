Attribute VB_Name = "ModSurfaces"
Public BackBuffer As DirectDrawSurface7
Public View As DirectDrawSurface7
Public Sprites As DirectDrawSurface7
Public Background As DirectDrawSurface7
Public Intro As DirectDrawSurface7
'Public Words As DirectDrawSurface7

Public ViewDesc As DDSURFACEDESC2
Public BackBufferDesc As DDSURFACEDESC2
Public SpritesDesc As DDSURFACEDESC2
Public BackgroundDesc As DDSURFACEDESC2
'Public WordsDesc As DDSURFACEDESC2
Public IntroDesc As DDSURFACEDESC2

Public BackBufferCaps As DDSCAPS2

Public ColorKey As DDCOLORKEY

Sub CreatePrimaryAndBackBuffer()
Set View = Nothing
Set BackBuffer = Nothing

ViewDesc.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
ViewDesc.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
ViewDesc.lBackBufferCount = 1
Set View = DX_Draw.CreateSurface(ViewDesc)

BackBufferCaps.lCaps = DDSCAPS_BACKBUFFER
Set BackBuffer = View.GetAttachedSurface(BackBufferCaps)
BackBuffer.GetSurfaceDesc ViewDesc

BackBuffer.SetFontTransparency True
End Sub

Sub LoadAllPics()
Dim Path As String

CreatePrimaryAndBackBuffer

Set Sprites = Nothing

ModDX7.CreateSurfaceFromFile Sprites, SpritesDesc, App.Path & "\Graphix\Sprites.bmp", 180, 30
ModDX7.CreateSurfaceFromFile Background, BackgroundDesc, App.Path & "\Graphix\Background.bmp", 640, 480
'ModDX7.CreateSurfaceFromFile Words, WordsDesc, App.Path & "\Graphix\Words.bmp", 640, 480
ModDX7.CreateSurfaceFromFile Intro, IntroDesc, App.Path & "\Graphix\Intro.bmp", 640, 480

ModDX7.AddColorKey Sprites, ColorKey, vbBlack, vbBlack
'ModDX7.AddColorKey Words, ColorKey, vbBlack, vbBlack

End Sub

Sub UnloadAllPics()
Set Sprites = Nothing
Set BackBuffer = Nothing
Set View = Nothing
End Sub

'Press a number from 1 to 9 to start the game at that level of difficulty (1 = easy peasy, 9 = well 'ard). Collect the apples to score points. The amount of points per apple depends on the difficulty level. The game ends when you crash (into either yourself or a wall). The action will get faster and your snake will get longer as the game progresses.
