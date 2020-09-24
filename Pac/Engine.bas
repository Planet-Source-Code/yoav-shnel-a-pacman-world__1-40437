Attribute VB_Name = "modEngine"
Option Explicit

Public Playing As Boolean
Public EndFlag As Boolean
Public PauseFlag As Boolean

Public GameSpeed As Integer

' the players
Public Pac As Pacman
Dim Ghosts() As Ghost

' ghost statistics
Public GhostSpeed As Integer
Public ScaredInterval As Integer

Public CameraX As Integer
Public CameraY As Integer

Private Function DetectCollision(ByRef g As Ghost) As Boolean
    Dim HorizontalCollision As Boolean
    Dim VerticalCollision As Boolean
    
    If Pac.X >= g.X And Pac.X < g.X + CellSize - 10 Then ' caught from west (maybe)
        HorizontalCollision = True
    ElseIf Pac.X >= g.X - CellSize + 10 And Pac.X < g.X Then ' caught from east (maybe)
        HorizontalCollision = True
    End If
    
    If Pac.Y >= g.Y And Pac.Y < g.Y + CellSize - 10 Then ' caught from north (maybe)
        VerticalCollision = True
    ElseIf Pac.Y >= g.Y - CellSize + 10 And Pac.Y < g.Y Then ' caught from south (maybe)
        VerticalCollision = True
    End If
    
    DetectCollision = HorizontalCollision And VerticalCollision
End Function

Private Sub Move(ByVal Direction As Directions, ByRef X As Integer, ByRef Y As Integer, ByVal Speed As Integer, ByVal bufDC As Long, ByVal srcDC As Long, ByVal bkgDC, Eat As Boolean, Optional DoMove As Boolean = True)
    If Eat Then ' blit empty background onto map
        BitBlt bkgDC, X + 1, Y + 1, CellSize - 2, CellSize - 2, 0, 0, 0, vbBlackness ' erase old image
    End If
    
    If DoMove Then
        ' move according to direction
        Select Case Direction
            Case East:
                X = X + Speed
            Case South:
                Y = Y + Speed
            Case West:
                X = X - Speed
            Case North:
                Y = Y - Speed
        End Select
    End If
    
    ' corrections
    If X < 0 Then X = 0
    If Y < 0 Then Y = 0
    If X > (MapWidth - 1) * CellSize Then X = (MapWidth - 1) * CellSize
    If Y > (MapHeight - 1) * CellSize Then Y = (MapWidth - 1) * CellSize
    
    ' blit sprite onto buffer
    BitBlt bufDC, X, Y, CellSize, CellSize, srcDC, 0, CellSize, vbSrcAnd   ' blit mask for new image
    BitBlt bufDC, X, Y, CellSize, CellSize, srcDC, 0, 0, vbSrcPaint  ' blit the new image
End Sub

' scare a ghost after Pacman eats super pill
Private Sub Scare(g As Ghost)
    g.Scared = True
    g.ScaredCount = 0
End Sub

' check for pills after Pacman moves
' returns true if pac has just eaten
Private Function PacCheck() As Boolean
    Dim Cell As Long
    Dim i As Integer
    
    ' get current cell from map
    Select Case Pac.Direction
        Case North:
            Cell = GetPixel(frmMain.picMap.hdc, Pac.X + CellSize / 2, Pac.Y + 3)
        Case East:
            Cell = GetPixel(frmMain.picMap.hdc, Pac.X + CellSize - 4, Pac.Y + CellSize / 2)
        Case South:
            Cell = GetPixel(frmMain.picMap.hdc, Pac.X + CellSize / 2, Pac.Y + CellSize - 4)
        Case West:
            Cell = GetPixel(frmMain.picMap.hdc, Pac.X + 3, Pac.Y + CellSize / 2)
    End Select
    'Cell = Map(Pac.X, Pac.Y)
    
    'If (Cell Mod 32) >= 16 Then ' check 5th bit for regular pill
    If Cell = MapPillColor Then
        PacCheck = True
        Pac.Score = Pac.Score + 1
        Pills = Pills - 1
        If AudioOn Then PlaySound GetAppPath & sfxEat, CLng(0), SND_ASYNC + SND_FILENAME + SND_NOWAIT + SND_NOSTOP
    End If
    'If (Cell Mod 64) >= 32 Then ' check 6th bit for super pill
    If Cell = MapSuperPillColor Then
        PacCheck = True
        For i = LBound(Ghosts) To UBound(Ghosts)
            Scare Ghosts(i)
        Next i
        If AudioOn Then PlaySound GetAppPath & sfxEat, CLng(0), SND_ASYNC + SND_FILENAME + SND_NOWAIT + SND_NOSTOP
    End If
    ClearMap Pac.X, Pac.Y ' clear pills from map
End Function

' check for kills after ghost moves
Private Sub GhostCheck(g As Ghost)
    Dim Cell As Byte

    ' check if ghost is scared
    If g.Scared Then
        ' check if he should stay scared
        g.ScaredCount = g.ScaredCount + 1
        If g.ScaredCount = ScaredInterval Then
            g.Scared = False
        End If
    End If

    ' look for collisions
    If DetectCollision(g) Then
        If g.Scared Then
            KillGhost g
        Else
            KillPac
        End If
    End If
End Sub

' main game function: plays one turn
Public Sub Play(bufDC As Long, bkgDC As Long)
    Dim i As Integer
    Dim dir As Integer
    Dim PacDC As Long
    Dim GhostDC As Long
    Dim OldDir As Integer
    Dim NewDir As Integer
    Dim Chasing As Boolean
    
    Playing = True
    
    ' blit background onto buffer
    BitBlt bufDC, 0, 0, MapWidth * CellSize, MapHeight * CellSize, bkgDC, 0, 0, vbSrcCopy
    
    ' move Pacman
    PacDC = PacSprite(Pac)
    Move Pac.Direction, Pac.X, Pac.Y, Pac.Speed, bufDC, PacDC, bkgDC, PacCheck, CanMove(Pac.Direction, Pac.X, Pac.Y)
    
    ' move ghosts
    For i = LBound(Ghosts) To UBound(Ghosts)
        Chasing = CanChase(Ghosts(i))
        If (Not Chasing) Or Ghosts(i).Scared Then
            ' unless the ghost is chasing pac (he won't chase him if his scared)
            ' it moves around randomely
            dir = Int(Rnd * 2) ' choose random direction to turn
            If dir = 0 Then dir = -1
            OldDir = Ghosts(i).Direction
            ' if centered in cell check for a possibility to make a random turn
            If (Ghosts(i).X = (Ghosts(i).X \ CellSize) * CellSize) And (Ghosts(i).Y = (Ghosts(i).Y \ CellSize) * CellSize) And Int(Rnd * 2) Then
                ' get new direction
                NewDir = OldDir + dir
                If NewDir = -1 Then NewDir = 3
                If NewDir = 4 Then NewDir = 0
                ' if we can turn we turn
                If CanMove(NewDir, Ghosts(i).X, Ghosts(i).Y) Then
                    Ghosts(i).Direction = NewDir
                End If
            End If
            ' if the ghost is stuck he scrolls through the
            ' directions until he finds one open
            Do While (Not CanMove(Ghosts(i).Direction, Ghosts(i).X, Ghosts(i).Y)) ' Or Ghosts(i).LeaveHouse
                Ghosts(i).Direction = Ghosts(i).Direction + dir
                If Ghosts(i).Direction = -1 Then
                    Ghosts(i).Direction = 3
                ElseIf Ghosts(i).Direction = 4 Then
                    Ghosts(i).Direction = 0
                End If
                If Ghosts(i).Direction = OldDir Then ' infinite loop bug workaround!!
                    Ghosts(i).Direction = (OldDir + 2) Mod 4 ' go back from where we came
                    Exit Do ' exit infinite loop
                End If
            Loop
        End If
        If Ghosts(i).Direction <> OldDir Then ' after turning we center ghost to axis
            Select Case Ghosts(i).Direction
                Case East, West:
                    ' center on y axis
                    Ghosts(i).Y = (Ghosts(i).Y \ CellSize) * CellSize
                Case North, South:
                    ' center on x axis
                    Ghosts(i).X = (Ghosts(i).X \ CellSize) * CellSize
            End Select
        End If
        GhostDC = GhostSprite(Ghosts(i))
        Move Ghosts(i).Direction, Ghosts(i).X, Ghosts(i).Y, Ghosts(i).Speed, bufDC, GhostDC, bkgDC, False
        GhostCheck Ghosts(i)
    Next i
    
    Playing = False
End Sub

Public Sub SetDirection(Direction As Integer)
    If Direction <> Pac.Direction Then ' we have to turn
        If Direction Mod 2 <> Pac.Direction Mod 2 Then ' center to axis
            Select Case Direction
                Case East, West:
                    ' center on y axis
                    Pac.Y = ((Pac.Y + CellSize / 2) \ CellSize) * CellSize
                Case North, South:
                    ' center on x axis
                    Pac.X = ((Pac.X + CellSize / 2) \ CellSize) * CellSize
            End Select
        End If
        Pac.Direction = Direction
    End If
End Sub

Private Sub KillPac()
    If AudioOn Then PlaySound GetAppPath & sfxPacDie, CLng(0), SND_ASYNC + SND_FILENAME + SND_NOWAIT
    If Pac.Lives = 0 Then
        frmMain.GameOver
    Else
        Pac.Lives = Pac.Lives - 1
        frmMain.Killed
        StartPac
        StartGhosts
    End If
End Sub

Private Sub KillGhost(g As Ghost)
    Pac.Score = Pac.Score + 10
    If AudioOn Then PlaySound GetAppPath & sfxGhostDie, CLng(0), SND_ASYNC + SND_FILENAME + SND_NOWAIT
    g.Scared = False
    g.X = InitGhostsX
    g.Y = InitGhostsY
End Sub

Public Sub StartGhosts(Optional BackHome As Boolean = True)
    Dim i As Integer
    
    For i = LBound(Ghosts) To UBound(Ghosts)
        Ghosts(i).Scared = False
        Ghosts(i).Color = i Mod 4
        Ghosts(i).Direction = i Mod 4
        Ghosts(i).Speed = GhostSpeed
        If BackHome Then Ghosts(i).X = InitGhostsX
        If BackHome Then Ghosts(i).Y = InitGhostsY
    Next i
End Sub

Private Sub StartPac()
    Pac.Direction = 0
    Pac.X = InitPacX
    Pac.Y = InitPacY
End Sub

Public Sub StartGame()
    'CurrentLevel = 1
    Pac.Lives = 3
    Pac.Speed = 4
    Pac.Score = 0
    GhostSpeed = 2
    ScaredInterval = 250
    ReDim Ghosts(3)
    StartPac
    StartGhosts
End Sub

' game loop
Public Sub PlayGame()
    Dim t1 As Long
    Dim t2 As Long
    
    Static HCameraDirection As Directions
    Static VCameraDirection As Directions
    
    t2 = GetTickCount
    
    Do
        DoEvents
        t1 = GetTickCount
        
        If (t1 - t2) >= GameSpeed And Not PauseFlag And Not EndFlag Then ' check appropiate interval has passed
            Play frmMain.picCanvas.hdc, frmMain.picMap.hdc ' play one turn
            
            If frmMain.picCanvas.ScaleWidth > frmMain.picView.ScaleWidth Then
                If Pac.X <= CameraX + 50 Then
                    HCameraDirection = West
                End If
                If Pac.X >= CameraX + 200 And Pac.X + CellSize <= CameraX + frmMain.picView.ScaleWidth - 200 Then
                    HCameraDirection = None
                End If
                If Pac.X + CellSize >= CameraX + frmMain.picView.ScaleWidth - 50 Then
                    HCameraDirection = East
                End If
                
                Select Case HCameraDirection
                    Case East
                        CameraX = CameraX + Pac.Speed * 1.5
                    Case West
                        CameraX = CameraX - Pac.Speed * 1.5
                End Select
                
                If CameraX > MapWidth * CellSize - frmMain.picView.ScaleWidth Then CameraX = MapWidth * CellSize - frmMain.picView.ScaleWidth
                If CameraX < 0 Then CameraX = 0
            Else
                CameraX = 0
            End If
            
            If frmMain.picCanvas.ScaleHeight > frmMain.picView.ScaleHeight Then
                If Pac.Y <= CameraY + 50 Then
                    VCameraDirection = North
                End If
                If Pac.Y >= CameraY + 200 And Pac.Y + CellSize <= CameraY + frmMain.picView.ScaleHeight - 200 Then
                    VCameraDirection = None
                End If
                If Pac.Y + CellSize >= CameraY + frmMain.picView.ScaleHeight - 50 Then
                    VCameraDirection = South
                End If
                
                Select Case VCameraDirection
                    Case South
                        CameraY = CameraY + Pac.Speed * 1.5
                    Case North
                        CameraY = CameraY - Pac.Speed * 1.5
                End Select
                
                If CameraY > MapHeight * CellSize - frmMain.picView.ScaleHeight Then CameraY = MapHeight * CellSize - frmMain.picView.ScaleHeight
                If CameraY < 0 Then CameraY = 0
            Else
                CameraY = 0
            End If
            
            ' blit buffer onto screen
            BitBlt frmMain.picView.hdc, 0, 0, frmMain.picView.ScaleWidth, frmMain.picView.ScaleHeight, frmMain.picCanvas.hdc, CameraX, CameraY, vbSrcCopy
            frmMain.picView.Refresh
            ' refresh score
            frmMain.lblScore = Pac.Score
            If Pills = 0 Then ' finished the level
                If AudioOn Then PlaySound GetAppPath & sfxWin, CLng(0), SND_ASYNC + SND_FILENAME + SND_NOWAIT
                If ScaredInterval > 100 Then
                    ' decrease time the ghosts are scared
                    ScaredInterval = ScaredInterval - 50
                End If
                If GhostSpeed < Pac.Speed Then
                    ' increase ghost speed
                    GhostSpeed = GhostSpeed + 2
                Else
                    ' add another ghost
                    ReDim Ghosts(UBound(Ghosts) + 1)
                End If
                frmMain.NextLevel
                StartPac
                StartGhosts
            End If
            t2 = GetTickCount
        End If
    Loop Until EndFlag
End Sub

' this function checks to see if pac is in the line of site of a ghost
' if so, it returns true and set the ghost's direction to chase pac
Public Function CanChase(ByRef g As Ghost) As Boolean
    Dim i As Integer
    Dim iStep As Integer
    Dim dir As Directions
    Dim gx As Integer
    Dim gy As Integer
    Dim px As Integer
    Dim py As Integer
    
    If Pac.X = g.X Or Pac.Y = g.Y Then
        
        gx = (g.X \ CellSize) * CellSize ' lined up ghost x
        gy = (g.Y \ CellSize) * CellSize ' lined up ghost y
        px = (Pac.X \ CellSize) * CellSize ' lined up pac x
        py = (Pac.Y \ CellSize) * CellSize ' lined up pac y
        
        If py = gy Then
            If g.Y <> gy Then Exit Function ' only chase if ghost is lined up
            If px < gx Then
                dir = West
                iStep = CellSize
            Else
                dir = East
                iStep = -CellSize
            End If
            
            For i = px + iStep To gx - iStep Step iStep
                If (Map(i, py) Mod 8) >= 4 Or (Map(i, py) Mod 2) >= 1 Then Exit For ' wall between them
            Next i
            
            If i = gx Then  ' no walls between them
                If CanMove(dir, g.X, g.Y) Then
                    ' chase
                    CanChase = True
                    g.Direction = dir
                End If
            End If
        
        ElseIf px = gx Then
            If g.X <> gx Then Exit Function ' only chase if ghost is lined up
            If py < gy Then
                dir = North
                iStep = CellSize
            Else
                dir = South
                iStep = -CellSize
            End If
                
            For i = py + iStep To gy - iStep Step iStep
                If (Map(px, i) Mod 16) >= 8 Or (Map(px, i) Mod 4) >= 2 Then Exit For ' wall between them
            Next i
            
            If i = gy Then ' no walls between them
                If CanMove(dir, g.X, g.Y) Then
                    ' chase
                    CanChase = True
                    g.Direction = dir
                End If
            End If
            
        End If
    End If
End Function
