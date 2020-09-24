Attribute VB_Name = "modSFX"
Option Explicit

Public AudioOn As Boolean

Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Public Const SND_ASYNC = &H1         '  play asynchronously
Public Const SND_FILENAME = &H20000     '  name is a file name
Public Const SND_LOOP = &H8         '  loop the sound until next sndPlaySound
Public Const SND_NOWAIT = &H2000      '  don't wait if the driver is busy
Public Const SND_NOSTOP = &H10        '  don't stop any currently playing sound

' wav filenames
Public Const sfxEat As String = "\sounds\eat.wav"
Public Const sfxPacDie As String = "\sounds\pdie.wav"
Public Const sfxGhostDie As String = "\sounds\gdie.wav"
Public Const sfxWin As String = "\sounds\win.wav"
Public Const sfxGameOver As String = "\sounds\gameover.wav"
