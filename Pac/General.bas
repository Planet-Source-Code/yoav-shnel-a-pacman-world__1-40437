Attribute VB_Name = "modGeneral"
Option Explicit

' db connection
Public cn As New Connection

'code timer
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const CellSize As Integer = 32 ' size of sprites

Public Const FrameNumber As Integer = 2 ' number of different frames
Public Const FrameInterval As Integer = 4 ' number of ticks until frame changes

Public Enum Directions
    East = 0
    South = 1
    West = 2
    North = 3
    None = 4
End Enum

' our hero
Public Type Pacman
    X As Integer
    Y As Integer
    Direction As Directions
    FrameCount As Integer
    Speed As Integer
    Lives As Integer
    Score As Integer
End Type

' the bad guys
Public Type Ghost
    X As Integer
    Y As Integer
    Direction As Directions
    FrameCount As Integer
    Speed As Integer
    Color As Integer
    Scared As Boolean
    ScaredCount As Integer ' number of ticks since scared
End Type

Public Function GetAppPath() As String
    ' check for final '\' (in case app.path is a root drive like "C:\"
    If Right(App.Path, 1) = "\" Then
        GetAppPath = Left(App.Path, Len(App.Path) - 1)
    Else
        GetAppPath = App.Path
    End If
End Function


