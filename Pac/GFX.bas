Attribute VB_Name = "modGFX"
Option Explicit

' The following API calls are for:

' blitting
Public Declare Function BitBlt Lib "gdi32" (ByVal hdestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

' creating buffers / loading sprites
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, _
  ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, _
  ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, _
  ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

' loading sprites
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" _
  (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

' cleanup
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

' for the editor
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Const IMAGE_BITMAP As Long = &O0    ' Used with LoadImage.
Private Const LR_LOADFROMFILE As Long = 16  ' Used with LoadImage.

Dim PacSprites(1, 3) As Long
Dim GhostSprites(1, 3, 1) As Long
Dim ScaredSprites(1) As Long

' This will load a sprite file into a DC and create a mask for it based on the transcolor passed.
Public Function LoadSprite(ByVal FileName As String, ByRef Width As Long, ByRef Height As Long, _
  ByVal TransColor As Long, Optional ByVal SecondaryTransColor As Long = -1) As Long
    Dim lngDestDC As Long    ' DC of the final sprite picture.
    Dim lngDestBMP As Long   ' Bitmap for the DC above.
    Dim lngLoadDC As Long    ' DC of the initial loaded picture.
    Dim lngMonoDC As Long    ' DC of the temporary monochrome mask.
    Dim lngMonoBMP As Long   ' Bitmap for the DC above.
    Dim lngBrush As Long     ' Used to fill the DC with specific color.
    Dim udtRect As RECT      ' Used to define DC area to be filled.

    ' Load the sprite picture from the file.
    lngLoadDC = LoadPicture(FileName, Width, Height)
    If lngLoadDC = 0 Then GoTo Done

    ' Convert the secondary transparent color to the primary one.
    If SecondaryTransColor > -1 Then
        ' Create a dc and bitmap for a mono mask and destination holder.
        ' Then select the new bitmaps into their respective dc's.
        lngMonoDC = CreateCompatibleDC(lngLoadDC)
        If lngMonoDC = 0 Then GoTo Done
        lngMonoBMP = CreateBitmap(Width, Height, 1, 1, ByVal 0&)
        If lngMonoBMP = 0 Then GoTo Done
        DeleteObject SelectObject(lngMonoDC, lngMonoBMP)

        lngDestDC = CreateCompatibleDC(lngLoadDC)
        If lngDestDC = 0 Then GoTo Done
        lngDestBMP = CreateCompatibleBitmap(lngLoadDC, Width, Height * 2)
        If lngDestBMP = 0 Then GoTo Done
        DeleteObject SelectObject(lngDestDC, lngDestBMP)

        ' Copy the loaded dc into the destination dc.
        Call BitBlt(lngDestDC, 0, 0, Width, Height, lngLoadDC, 0, 0, vbSrcCopy)

        ' Now create the mask bitmap and copy it to the destination.
        Call SetBkColor(lngLoadDC, SecondaryTransColor)
        Call BitBlt(lngMonoDC, 0, 0, Width, Height, lngLoadDC, 0, 0, vbSrcCopy)
        Call BitBlt(lngDestDC, 0, Height, Width, Height, lngMonoDC, 0, 0, vbSrcCopy)

        ' Now create the inverse of the mask and AND it with the sprite picture in the destination dc.
        Call BitBlt(lngMonoDC, 0, 0, Width, Height, lngMonoDC, 0, 0, vbNotSrcCopy)
        Call BitBlt(lngDestDC, 0, 0, Width, Height, lngMonoDC, 0, 0, vbSrcAnd)

        ' Fill the load dc with the transcolor.
        lngBrush = CreateSolidBrush(TransColor)
        With udtRect
            .Left = 0: .Top = 0: .Right = Width: .Bottom = Height
        End With
        Call FillRect(lngLoadDC, udtRect, lngBrush)
        Call DeleteObject(lngBrush)

        ' Now apply the destination to the original picture like you would normal display a sprite.
        Call BitBlt(lngLoadDC, 0, 0, Width, Height, lngDestDC, 0, Height, vbSrcAnd)
        Call BitBlt(lngLoadDC, 0, 0, Width, Height, lngDestDC, 0, 0, vbSrcPaint)

        ' Cleanup temporary varaiables.
        Call DeleteObject(lngMonoBMP)
        Call DeleteObject(lngDestBMP)
        Call DeleteDC(lngMonoDC)
        Call DeleteDC(lngDestDC)
    End If

    ' Create a dc and bitmap for a mono mask and destination holder.
    ' Then select the new bitmaps into their respective dc's.
    lngMonoDC = CreateCompatibleDC(lngLoadDC)
    If lngMonoDC = 0 Then GoTo Done
    lngMonoBMP = CreateBitmap(Width, Height, 1, 1, ByVal 0&)
    If lngMonoBMP = 0 Then GoTo Done
    DeleteObject SelectObject(lngMonoDC, lngMonoBMP)

    lngDestDC = CreateCompatibleDC(lngLoadDC)
    If lngDestDC = 0 Then GoTo Done
    lngDestBMP = CreateCompatibleBitmap(lngLoadDC, Width, Height * 2)
    If lngDestBMP = 0 Then GoTo Done
    DeleteObject SelectObject(lngDestDC, lngDestBMP)

    ' Copy the loaded dc into the destination dc.
    Call BitBlt(lngDestDC, 0, 0, Width, Height, lngLoadDC, 0, 0, vbSrcCopy)

    ' Now create the mask bitmap and copy it to the destination.
    Call SetBkColor(lngLoadDC, TransColor)
    Call BitBlt(lngMonoDC, 0, 0, Width, Height, lngLoadDC, 0, 0, vbSrcCopy)
    Call BitBlt(lngDestDC, 0, Height, Width, Height, lngMonoDC, 0, 0, vbSrcCopy)

    ' Now create the inverse of the mask and AND it with the sprite picture in the destination dc.
    Call BitBlt(lngMonoDC, 0, 0, Width, Height, lngMonoDC, 0, 0, vbNotSrcCopy)
    Call BitBlt(lngDestDC, 0, 0, Width, Height, lngMonoDC, 0, 0, vbSrcAnd)

    ' Pass back the completed sprite.
    LoadSprite = lngDestDC

Done:
    ' We should now be done, so delete all unused objects.
    Call DeleteObject(lngMonoBMP)
    Call DeleteObject(lngDestBMP)
    Call DeleteDC(lngLoadDC)
    Call DeleteDC(lngMonoDC)
End Function

' This loads in a bitmap to a device context from a file.
Public Function LoadPicture(ByVal FileName As String, ByRef Width As Long, ByRef Height As Long) As Long
    Dim lngBitmap As Long
    Dim lngDC As Long
    Dim udtBitmapData As BITMAP

    ' Create a Device Context.
    lngDC = CreateCompatibleDC(0)
    If lngDC = 0 Then Exit Function

    ' Load the image.
    lngBitmap = LoadImage(0, FileName, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE)
    If lngBitmap = 0 Then
        ' Failure in loading bitmap.
        Call DeleteDC(lngDC)
        Exit Function
    End If

    ' Get Width & Height of the bitmap.
    GetObject lngBitmap, Len(udtBitmapData), udtBitmapData
    Width = udtBitmapData.bmWidth
    Height = udtBitmapData.bmHeight

    ' Put the Bitmap into the device context.
    DeleteObject SelectObject(lngDC, lngBitmap)
    DeleteObject lngBitmap

    ' Return the device context.
    LoadPicture = lngDC
End Function

Public Function PacSprite(ByRef Pac As Pacman) As Long
    ' increase frame counter
    Pac.FrameCount = Pac.FrameCount + 1
    If Pac.FrameCount = FrameInterval * FrameNumber Then Pac.FrameCount = 0
    
    ' load sprite
    PacSprite = PacSprites(Pac.FrameCount \ FrameInterval, Pac.Direction)
End Function

Public Function GhostSprite(ByRef g As Ghost) As Long
    ' increase frame counter
    g.FrameCount = g.FrameCount + 1
    If g.FrameCount = FrameInterval * FrameNumber Then g.FrameCount = 0
    
    ' load sprite
    If g.Scared Then
        GhostSprite = ScaredSprites(g.FrameCount \ FrameInterval)
        If g.ScaredCount > ScaredInterval - 25 Then
            ' make ghost blink
            If g.FrameCount Mod 5 Then GhostSprite = 0
        End If
    Else
        GhostSprite = GhostSprites(g.FrameCount \ FrameInterval, g.Color, g.Direction \ 2)
    End If
End Function

Public Sub LoadGFX()
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim Path As String
    
    Path = GetAppPath
    
    For i = 0 To 1
        ScaredSprites(i) = LoadSprite(Path & "\pics\Scared_" & i & ".bmp", CellSize, CellSize, vbBlack)
        For j = 0 To 3
            PacSprites(i, j) = LoadSprite(Path & "\pics\Pac" & i & "_" & j & ".bmp", CellSize, CellSize, vbBlack)
            For k = 0 To 1
                GhostSprites(i, j, k) = LoadSprite(Path & "\pics\Ghost" & i & "_" & j & "_" & k & ".bmp", CellSize, CellSize, vbBlack)
            Next k
        Next j
    Next i
End Sub

Public Sub UnloadGFX()
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim Path As String
    
    Path = GetAppPath
    
    For i = 0 To 1
        DeleteDC ScaredSprites(i)
        For j = 0 To 3
            DeleteDC PacSprites(i, j)
            For k = 0 To 1
                DeleteDC GhostSprites(i, j, k)
            Next k
        Next j
    Next i
End Sub
