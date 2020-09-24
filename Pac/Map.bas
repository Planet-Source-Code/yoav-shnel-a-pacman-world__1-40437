Attribute VB_Name = "modMap"
Option Explicit

' map variables
Dim MapArr() As Byte
Public MapWidth As Integer
Public MapHeight As Integer

' initial location for pacman
Public InitPacX As Integer
Public InitPacY As Integer

' initial location for ghosts
Public InitGhostsX As Integer
Public InitGhostsY As Integer

' color scheme
Public Const MapBorderColor As Long = &HC0FFC0
Public Const MapSuperPillColor As Long = vbBlue
Public Const MapPillColor As Long = vbYellow
Public Const MapHouseColor As Long = &H404040

' final score with which the level is finished (number of pills on map)
Public Pills As Integer

Public Function Map(ByVal X As Integer, ByVal Y As Integer) As Byte
    Map = MapArr(X \ CellSize, Y \ CellSize)
End Function

Public Sub ClearMap(ByVal X As Integer, ByVal Y As Integer)
    ' clear super pill
    If (MapArr(X \ CellSize, Y \ CellSize) Mod 64) >= 32 Then MapArr(X \ CellSize, Y \ CellSize) = MapArr(X \ CellSize, Y \ CellSize) - 32
    ' clear pill
    If (MapArr(X \ CellSize, Y \ CellSize) Mod 32) >= 16 Then MapArr(X \ CellSize, Y \ CellSize) = MapArr(X \ CellSize, Y \ CellSize) - 16
End Sub

Public Function CanMove(ByVal d As Directions, ByVal X As Integer, ByVal Y As Integer) As Boolean
    Select Case d
        Case North:
            CanMove = ((MapArr(X \ CellSize, (Y + CellSize - 1) \ CellSize) Mod 16) < 8) ' check 4th bit
        Case West:
            CanMove = ((MapArr((X + CellSize - 1) \ CellSize, Y \ CellSize) Mod 8) < 4) ' check 3rd bit
        Case South:
            CanMove = ((MapArr(X \ CellSize, Y \ CellSize) Mod 4) < 2)  ' check 2nd bit
        Case East:
            CanMove = ((MapArr(X \ CellSize, Y \ CellSize) Mod 2) < 1)  ' check 1st bit
    End Select
End Function

Public Function LoadMap(MapID As Integer) As Boolean
    On Error GoTo NoLoad
    Dim rs As New Recordset
    Dim Cell As Byte
    Dim X As Integer
    Dim Y As Integer
        
    ' get level information
    Set rs.ActiveConnection = cn
    rs.CursorType = adOpenForwardOnly
    rs.LockType = adLockReadOnly
    rs.Open "select * from Levels where id=" & MapID
    
    If rs.EOF Then GoTo NoLoad ' level doesn't exist
    
    MapWidth = rs("width")
    MapHeight = rs("height")
    
    ReDim MapArr(MapWidth - 1, MapHeight - 1) ' redim map matrix
    
    rs.Close
    
    ' get map information
    rs.Open "select * from Maps where id=" & MapID & " order by x,y"
    
    frmMain.picCanvas.Width = CellSize * Screen.TwipsPerPixelX + 100
    frmMain.picCanvas.Height = CellSize * Screen.TwipsPerPixelY + 100
    frmMain.picMap.Width = CellSize * MapWidth * Screen.TwipsPerPixelX + 100
    frmMain.picMap.Height = CellSize * MapHeight * Screen.TwipsPerPixelY + 100
    
    frmMain.picMap.Cls
    frmMain.picMap.Refresh
    
    If rs.EOF Then GoTo NoLoad ' no map information
    
    Pills = 0
    
    For X = 0 To MapWidth - 1
        For Y = 0 To MapHeight - 1
            If rs("x") <> X Or rs("y") <> Y Then GoTo NoLoad ' missing map information
            Cell = rs("cell")
            frmMain.picCanvas.Cls
            If (Cell Mod 16) >= 8 Then ' north wall
                frmMain.picCanvas.Line (0, 0)-(CellSize, 0), MapBorderColor
            End If
            If (Cell Mod 8) >= 4 Then ' west wall
                frmMain.picCanvas.Line (0, 0)-(0, CellSize), MapBorderColor
            End If
            If (Cell Mod 4) >= 2 Then ' south wall
                frmMain.picCanvas.Line (0, CellSize - 1)-(CellSize, CellSize - 1), MapBorderColor
            End If
            If (Cell Mod 2) >= 1 Then ' east wall
                frmMain.picCanvas.Line (CellSize - 1, 0)-(CellSize - 1, CellSize), MapBorderColor
            End If
            If (Cell Mod 64) >= 32 Then ' draw super pill
                frmMain.picCanvas.FillColor = MapSuperPillColor
                frmMain.picCanvas.Circle (CellSize / 2 - 1, CellSize / 2 - 1), CellSize / 16 + 1, MapSuperPillColor
            ElseIf (Cell Mod 128) < 64 Then ' draw regular pill (and validate that cell has one
                Pills = Pills + 1
                frmMain.picCanvas.FillColor = MapPillColor
                frmMain.picCanvas.Circle (CellSize / 2 - 1, CellSize / 2 - 1), CellSize / 16, MapPillColor
                If (Cell Mod 32) < 16 Then ' put regular pill value in cell
                    Cell = Cell + 16
                End If
            End If
            If (Cell Mod 128) >= 64 Then ' ghost house
                frmMain.picCanvas.FillColor = MapHouseColor
                frmMain.picCanvas.Line (0, 0)-(CellSize, CellSize), MapHouseColor, BF
                InitGhostsX = X * CellSize
                InitGhostsY = Y * CellSize
                ' make sure ghost house is closed
                'If (Cell Mod 2) < 1 Then Cell = Cell + 1 ' close east wall
                'If (Cell Mod 4) < 2 Then Cell = Cell + 2 ' close south wall
                'If (Cell Mod 8) < 4 Then Cell = Cell + 4 ' close west wall
                'If (Cell Mod 16) < 8 Then Cell = Cell + 8 ' close north wall
            End If
            If Cell >= 128 Then ' initial location for pacman
                InitPacX = X * CellSize
                InitPacY = Y * CellSize
            End If
            MapArr(X, Y) = Cell ' populate map matrix
            ' blit cell onto map
            BitBlt frmMain.picMap.hdc, X * CellSize, Y * CellSize, CellSize, CellSize, frmMain.picCanvas.hdc, 0, 0, vbSrcPaint
            rs.MoveNext
        Next Y
    Next X
    
    frmMain.picCanvas.Refresh
        
    ' get canvas ready for game (set to map size)
    frmMain.picCanvas.Width = frmMain.picMap.Width
    frmMain.picCanvas.Height = frmMain.picMap.Height
    
    ' blit map onto buffer
    frmMain.picCanvas.Cls
    frmMain.picCanvas.Refresh
    BitBlt frmMain.picCanvas.hdc, 0, 0, MapWidth * CellSize, MapHeight * CellSize, frmMain.picMap.hdc, 0, 0, vbSrcCopy
    frmMain.picCanvas.Refresh
    
    LoadMap = True ' map loaded successfully
     
NoLoad:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
End Function
