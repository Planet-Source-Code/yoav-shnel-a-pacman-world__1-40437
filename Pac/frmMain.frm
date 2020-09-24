VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   Caption         =   "Pac World"
   ClientHeight    =   7515
   ClientLeft      =   285
   ClientTop       =   630
   ClientWidth     =   8850
   ForeColor       =   &H0000FFFF&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   8850
   Begin MSComDlg.CommonDialog cmnDialog 
      Left            =   5760
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      HelpContext     =   1000
      HelpFile        =   "pachelp.hlp"
   End
   Begin VB.PictureBox picView 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   7815
      Left            =   0
      ScaleHeight     =   517
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   357
      TabIndex        =   2
      Top             =   0
      Width           =   5415
      Begin VB.Label lblHS 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   735
         Index           =   4
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label lblHS 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   735
         Index           =   3
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label lblHS 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   735
         Index           =   2
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label lblHS 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   735
         Index           =   1
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label lblHSName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Index           =   4
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   3015
      End
      Begin VB.Label lblHSName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Index           =   3
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   3015
      End
      Begin VB.Label lblHSName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Index           =   2
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   3015
      End
      Begin VB.Label lblHSName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Index           =   1
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   3015
      End
      Begin VB.Label lblHS 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   735
         Index           =   0
         Left            =   3480
         TabIndex        =   9
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblHSName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label lblMessage 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "David"
            Size            =   36
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   1335
         Left            =   600
         TabIndex        =   3
         Top             =   2400
         Width           =   4575
      End
   End
   Begin VB.PictureBox picMap 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   2415
      Left            =   5640
      ScaleHeight     =   157
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   245
      TabIndex        =   1
      Top             =   5400
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   1935
      HelpContextID   =   1000
      Left            =   6240
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   0
      Top             =   2880
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   7800
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label lblLives 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   5880
      TabIndex        =   6
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblScoreTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Score: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   7920
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblLivesTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Lives: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuGamePlay 
         Caption         =   "&New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuGamePause 
         Caption         =   "&Pause Game"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuGameEnd 
         Caption         =   "&End Game"
         Enabled         =   0   'False
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuOpt 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptSpeed 
         Caption         =   "Speed"
         Begin VB.Menu mnuSpeed 
            Caption         =   "Slow"
            Index           =   0
         End
         Begin VB.Menu mnuSpeed 
            Caption         =   "Medium"
            Index           =   1
         End
         Begin VB.Menu mnuSpeed 
            Caption         =   "Fast"
            Index           =   2
         End
      End
      Begin VB.Menu mnuOptSound 
         Caption         =   "Sound FX"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "&Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Cheat As String

Dim MapID As Integer
Dim rsMap As New Recordset
Public UserName As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    
    Select Case KeyCode
        Case vbKeyUp:
            SetDirection North
        Case vbKeyDown:
            SetDirection South
        Case vbKeyLeft:
            SetDirection West
        Case vbKeyRight:
            SetDirection East
        Case vbKeyA To vbKeyZ:
            Cheat = Cheat & Chr(KeyCode)
        Case vbKeyEscape:
            If Cheat = "NEXT" Then
                Pills = 1
            ElseIf Cheat = "SLOW" Then
                If GhostSpeed > 2 Then
                    GhostSpeed = GhostSpeed - 2
                    StartGhosts (False)
                End If
            End If
            Cheat = ""
    End Select
End Sub

Private Sub Form_Load()
    Dim Speed As Integer

    Randomize Timer
     
    ' load sprites
    LoadGFX
    
    ' window setting
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 7500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 7500)
    
    ' audio settings
    AudioOn = GetSetting(App.Title, "Settings", "Sound", True)
    mnuOptSound.Checked = AudioOn
    
    ' speed settings
    Speed = GetSetting(App.Title, "Settings", "Speed", 1)
    mnuSpeed(Speed).Checked = True
    GameSpeed = 15 + (2 - Speed) * 35
    
    ' open connection
    cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=pac.mdb;Persist Security Info=False"
    cn.Open
    
    ' Load map ids
    rsMap.ActiveConnection = cn
    rsMap.CursorType = adOpenKeyset
    rsMap.LockType = adLockReadOnly
    rsMap.Open "select id from Levels"
    
    ' display high scores
    HSPrint
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload
    End
End Sub

Private Sub Form_Resize()
    Dim i As Integer

    If WindowState <> vbMinimized Then
        If Width < lblLives.Width + 450 * Screen.TwipsPerPixelX Then
            Width = lblLives.Width + 450 * Screen.TwipsPerPixelX
        End If
        
        If Height < 520 * Screen.TwipsPerPixelY Then
            Height = 520 * Screen.TwipsPerPixelY
        End If
        
        picView.Top = 0
        picView.Left = 0
        picView.Height = ScaleHeight
        picView.Width = ScaleWidth - lblLives.Width
        lblScoreTitle.Top = 800
        lblScoreTitle.Left = picView.Width
        lblScoreTitle.Width = lblLives.Width
        lblScore.Top = lblScoreTitle.Top + lblScoreTitle.Height
        lblScore.Left = picView.Width
        lblScore.Width = lblLives.Width
        lblLivesTitle.Top = lblScoreTitle.Height + lblScore.Height + 1500
        lblLivesTitle.Left = picView.Width
        lblLivesTitle.Width = lblLives.Width
        lblLives.Top = lblLivesTitle.Top + lblLivesTitle.Height
        lblLives.Left = picView.Width
        lblMessage.Left = 0
        lblMessage.Top = (picView.ScaleHeight - lblMessage.Height) / 2
        lblMessage.Width = picView.ScaleWidth
        picView.ScaleMode = vbTwips
        For i = 0 To 4
            lblHSName(i).Left = 800
            lblHSName(i).Top = 700 + i * (lblHSName(i).Height + 400)
            'lblHSName(i).Width = picView.Width
            lblHS(i).Left = lblHSName(i).Width + 800
            lblHS(i).Top = lblHSName(i).Top
        Next i
        picView.ScaleMode = vbPixels
        picView.Cls
        picView.Refresh
        BitBlt picView.hdc, 0, 0, picView.ScaleWidth, picView.ScaleHeight, picCanvas.hdc, CameraX, CameraY, vbSrcCopy
        picView.Refresh
    Else
        If mnuGamePause.Enabled Then
            PauseGame
        End If
    End If
End Sub

Private Sub PauseGame()
    PauseFlag = True
    lblMessage.Caption = "Paused"
    mnuGamePause.Caption = "&Resume"
End Sub

Private Sub Unload()
    Dim i As Integer
    
    UnloadGFX
    
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
    
    SaveSetting App.Title, "Settings", "Sound", AudioOn
    
    For i = 0 To 2
        If mnuSpeed(i).Checked = True Then
            SaveSetting App.Title, "Settings", "Speed", i
            Exit For
        End If
    Next i

    rsMap.Close
    Set rsMap = Nothing
    cn.Close
    Set cn = Nothing
    End
End Sub

Private Sub mnuGameEnd_Click()
    GameOver
End Sub

Private Sub mnuGameExit_Click()
    Unload
    End
End Sub

Private Sub mnuGamePause_Click()
    If mnuGamePause.Caption = "&Pause Game" Then
        PauseGame
    Else
        PauseFlag = False
        lblMessage.Caption = ""
        mnuGamePause.Caption = "&Pause Game"
    End If
End Sub

Private Sub mnuGamePlay_Click()
    Dim i As Integer
    
    ' clear high score labels
    For i = 0 To 4
        lblHS(i).Visible = False
        lblHSName(i).Visible = False
    Next i

    ' start a new game
    lblMessage.Caption = "Get Ready..."
    rsMap.MoveFirst
    MapID = rsMap(0)
    If LoadMap(MapID) Then
        BitBlt picView.hdc, 0, 0, picView.ScaleWidth, picView.ScaleHeight, picMap.hdc, CameraX, CameraY, vbSrcCopy
        picView.Refresh
        EndFlag = False
        StartGame
        DoEvents
        Sleep 1500
        lblMessage.Caption = ""
        mnuGamePause.Enabled = True
        lblLives.Caption = Pac.Lives
        lblScore.Caption = Pac.Score
        mnuGameEnd.Enabled = True
        mnuGamePlay.Enabled = False
        PlayGame
    Else
        MsgBox "Unable to load level.", , "PacWorld"
        lblMessage.Caption = ""
    End If
End Sub

Public Sub GameOver()
    mnuGameEnd.Enabled = False
    mnuGamePlay.Enabled = True
    EndFlag = True
    lblMessage.Caption = "GAME OVER"
    DoEvents
    Sleep 3000
    If AudioOn Then PlaySound GetAppPath & sfxGameOver, CLng(0), SND_ASYNC + SND_FILENAME + SND_NOWAIT
    lblMessage.Caption = ""
    mnuGamePause.Enabled = False
    picView.Cls
    picCanvas.Cls
    picView.Refresh
    HSUpdate
End Sub

Public Sub NextLevel()
    Pac.Score = Pac.Score + 100
    lblMessage.Caption = "Yeah!"
    DoEvents
    Sleep 3000
    picView.Cls
    picView.Refresh
    Do
        MapID = NextMap
    Loop Until LoadMap(MapID)
    lblMessage.Caption = "Get Ready..."
    BitBlt picView.hdc, 0, 0, picView.ScaleWidth, picView.ScaleHeight, picMap.hdc, CameraX, CameraY, vbSrcCopy
    picView.Refresh
    DoEvents
    Sleep 1500
    lblMessage.Caption = ""
End Sub

Public Sub Killed()
    lblMessage.Caption = "Oops!"
    lblLives.Caption = Pac.Lives
    DoEvents
    Sleep 2000
    lblMessage.Caption = "Get Ready..."
    BitBlt picView.hdc, 0, 0, picView.ScaleWidth, picView.ScaleHeight, picMap.hdc, CameraX, CameraY, vbSrcCopy
    picView.Refresh
    DoEvents
    Sleep 1500
    lblMessage.Caption = ""
End Sub

Private Sub HSPrint()
    Dim i As Integer
    Dim rs As New Recordset
    
    rs.ActiveConnection = cn
    rs.CursorType = adOpenForwardOnly
    rs.LockType = adLockReadOnly
     
    rs.Open "select * from Scores order by score desc"
    
    While Not rs.EOF
        If Not IsNull(rs("name")) Then
            lblHSName(i).Caption = rs("name")
        End If
        If Not IsNull(rs("score")) Then
            lblHS(i).Caption = rs("score")
        End If
        lblHSName(i).Visible = True
        lblHS(i).Visible = True
        i = i + 1
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub HSUpdate()
    On Error GoTo EndSub
    Dim rs As New Recordset
    
    rs.ActiveConnection = cn
    rs.CursorType = adOpenForwardOnly
    rs.LockType = adLockPessimistic
    
    rs.Open "select * from Scores where score=(select min(score) from Scores)"
    
    If Pac.Score > rs("score") Then
        frmGetName.Show vbModal
        rs("score") = Pac.Score
        rs("name") = UserName
        rs.Update
    End If
    
    HSPrint
    
    rs.Close
    Set rs = Nothing
EndSub:
End Sub

Private Function NextMap() As Integer
    rsMap.MoveNext
    If rsMap.EOF Then rsMap.MoveFirst
    NextMap = rsMap(0)
End Function

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuHelpHelp_Click()
    cmnDialog.HelpFile = GetAppPath & "\pachelp.hlp"
    cmnDialog.HelpContext = 1000
    cmnDialog.HelpCommand = cdlHelpContext
    cmnDialog.ShowHelp
End Sub

Private Sub mnuOptSound_Click()
    mnuOptSound.Checked = Not mnuOptSound.Checked
    AudioOn = mnuOptSound.Checked
End Sub

Private Sub mnuSpeed_Click(Index As Integer)
    Dim i As Integer
    
    For i = 0 To 2
        If i = Index Then
            mnuSpeed(i).Checked = True
        Else
            mnuSpeed(i).Checked = False
        End If
    Next i
    
    GameSpeed = 15 + (2 - Index) * 35
End Sub
