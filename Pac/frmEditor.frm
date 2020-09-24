VERSION 5.00
Begin VB.Form frmEditor 
   Caption         =   "PacWorld Editor"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9855
   Icon            =   "frmEditor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   537
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   657
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   8400
      TabIndex        =   3
      Text            =   "My Level"
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Done"
      Height          =   615
      Left            =   8160
      TabIndex        =   9
      Top             =   4800
      Width           =   1695
   End
   Begin VB.TextBox txtWidth 
      Height          =   285
      Left            =   8400
      TabIndex        =   5
      Text            =   "20"
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox txtHeight 
      Height          =   285
      Left            =   8400
      TabIndex        =   4
      Text            =   "20"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Frame fraDrawOptions 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   8040
      TabIndex        =   8
      Top             =   2880
      Width           =   1815
      Begin VB.OptionButton optDraw 
         Caption         =   "Erase"
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   7
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton optDraw 
         Caption         =   "Draw"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.VScrollBar VScroll 
      Enabled         =   0   'False
      Height          =   3615
      Left            =   7440
      TabIndex        =   2
      Top             =   120
      Width           =   255
   End
   Begin VB.HScrollBar HScroll 
      Enabled         =   0   'False
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   3960
      Width           =   6855
   End
   Begin VB.PictureBox picView 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   3615
      Left            =   240
      ScaleHeight     =   237
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   461
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   6975
      Begin VB.PictureBox picMap 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00404040&
         Height          =   2415
         Left            =   600
         ScaleHeight     =   157
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   309
         TabIndex        =   12
         Top             =   360
         Width           =   4695
         Begin VB.PictureBox picHouse 
            BackColor       =   &H00FF00FF&
            BorderStyle     =   0  'None
            Height          =   855
            Left            =   840
            ScaleHeight     =   57
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   65
            TabIndex        =   13
            Top             =   600
            Width           =   975
         End
      End
   End
   Begin VB.Label lblName 
      Caption         =   "Name:"
      Height          =   255
      Left            =   8400
      TabIndex        =   14
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblWidth 
      Caption         =   "Width :"
      Height          =   255
      Left            =   8400
      TabIndex        =   11
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label lblHeight 
      Caption         =   "Height :"
      Height          =   255
      Left            =   8400
      TabIndex        =   10
      Top             =   1200
      Width           =   1335
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim EMapWidth As Integer
Dim EMapHeight As Integer
Private Const CellSize As Integer = 20 ' override general cellsize

Private Sub cmdSave_Click()
    'On Error GoTo ErrHandler
    Dim rs As New Recordset
    Dim MapID As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim Cell As Byte
    
    If txtName.Text = "" Then
        MsgBox "Enter a level name.", , "PacWorld"
        txtName.SetFocus
        Exit Sub
    End If
          
    cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=pac.mdb;Persist Security Info=False"
    cn.Open
    
    
    ' get new id for level
    rs.ActiveConnection = cn
    rs.CursorType = adOpenForwardOnly
    rs.LockType = adLockReadOnly
    rs.Open "select max(id) from Levels"
    
    If IsNull(rs(0)) Then
        MapID = 0
    Else
        MapID = rs(0) + 1
    End If
    
    rs.Close
    
    ' insert level into database
    cn.Execute "insert into Levels (id,width,height,name) values (" & MapID & "," & EMapWidth & "," & EMapHeight & ",'" & txtName.Text & "')"
    
    ' save map into database
    For X = 0 To EMapWidth - 1
        For Y = 0 To EMapHeight - 1
            'cell creation
            Cell = 0
            If GetPixel(picMap.hdc, (X + 1) * CellSize - 1, (Y + 0.5) * CellSize) = MapBorderColor Or X = EMapWidth - 1 Then
                Cell = Cell + 1 ' east wall
            End If
            If GetPixel(picMap.hdc, (X + 0.5) * CellSize, (Y + 1) * CellSize - 1) = MapBorderColor Or Y = EMapHeight - 1 Then
                Cell = Cell + 2 ' south wall
            End If
            If GetPixel(picMap.hdc, X * CellSize, (Y + 0.5) * CellSize) = MapBorderColor Or X = 0 Then
                Cell = Cell + 4 ' west wall
            End If
            If GetPixel(picMap.hdc, (X + 0.5) * CellSize, Y * CellSize) = MapBorderColor Or Y = 0 Then
                Cell = Cell + 8 ' north wall
            End If
            If GetPixel(picMap.hdc, (X + 0.5) * CellSize, (Y + 0.5) * CellSize) = MapSuperPillColor Then
                Cell = Cell + 32 ' super pill
            ElseIf GetPixel(picMap.hdc, (X + 0.5) * CellSize, (Y + 0.5) * CellSize) = vbCyan Then
                Cell = Cell + 64 ' ghost house
            Else
                Cell = Cell + 16 ' regular pill
            End If
            
            ' save cell in database
            cn.Execute "insert into Maps (id,x,y,cell) values (1," & X & "," & Y & "," & _
                Cell & ")"
        Next Y
    Next X
    
    MsgBox "Level Saved", , "PacWorld"
    
NoSave:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    
    Exit Sub
    
ErrHandler:
    MsgBox Err.Number & " - " & Err.Description
    GoTo NoSave
End Sub

Private Sub Form_Load()
    EMapHeight = 20
    EMapWidth = 20
End Sub

Private Sub Form_Resize()
    picView.Top = 0
    picView.Left = 0
    picView.Height = ScaleHeight - HScroll.Height
    picView.Width = fraDrawOptions.Left - 20 - VScroll.Width
    picMap.Top = 0
    picMap.Left = 0
    picMap.Height = EMapHeight * CellSize + 2
    picMap.Width = EMapWidth * CellSize + 2
    picHouse.Top = CellSize * 9
    picHouse.Left = CellSize * 9
    picHouse.Height = CellSize * 2
    picHouse.Width = CellSize * 2
    VScroll.Top = 0
    VScroll.Left = picView.Width
    VScroll.Height = picView.Height
    HScroll.Top = picView.Height
    HScroll.Left = 0
    HScroll.Width = picView.Width
End Sub

Private Sub HScroll_Change()
    picView.Left = -HScroll.Value * CellSize
End Sub

Private Sub picHouse_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim NewX As Long
    Dim NewY As Long

    If Button = vbLeftButton Then
        NewX = (X \ CellSize) * CellSize
        NewY = (Y \ CellSize) * CellSize
        If NewX < 0 Then NewX = 0
        If NewX > (EMapWidth - 1) * CellSize Then NewX = (EMapWidth - 1) * CellSize
        If NewY < 0 Then NewY = 0
        If NewY > (EMapHeight - 1) * CellSize Then NewY = (EMapHeight - 1) * CellSize
        If picHouse.Left <> NewX Or picHouse.Top <> NewY Then
            picHouse.Move (X \ CellSize) * CellSize, (Y \ CellSize) * CellSize
            'picMap.Refresh
        End If
    End If
End Sub

Private Sub picMap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim NewX As Long
    Dim NewY As Long
    
    If Button = vbLeftButton Then
        NewX = CLng(X / CellSize) * CellSize
        NewY = CLng(Y / CellSize) * CellSize
        If NewX < 0 Then NewX = 0
        If NewX > EMapWidth * CellSize Then NewX = EMapWidth * CellSize
        If NewY < 0 Then NewY = 0
        If NewY > EMapHeight * CellSize Then NewY = EMapHeight * CellSize
        If optDraw(0).Value Then ' draw
            If (Y > NewY - 5 And Y < NewY + 5) And Not (X > NewX - 5 And X < NewX + 5) Then
                picMap.Line (NewX, NewY - 1)-(NewX + CellSize, NewY), MapBorderColor, BF
            End If
            If (X > NewX - 5 And X < NewX + 5) And Not (Y > NewY - 5 And Y < NewY + 5) Then
                picMap.Line (NewX - 1, NewY)-(NewX, NewY + CellSize), MapBorderColor, BF
            End If
        Else ' erase
            If X > NewX - 5 And X < NewX + 5 Then
                picMap.Line (NewX, NewY - 1)-(NewX + CellSize, NewY), picMap.BackColor, BF
            End If
            If Y > NewY - 5 And Y < NewY + 5 Then
                picMap.Line (NewX - 1, NewY)-(NewX, NewY + CellSize), picMap.BackColor, BF
            End If
        End If
    End If
End Sub

Private Sub picMap_Resize()
    ' enable scrollbars if necessary
    If picMap.Width > picView.Width Then
        HScroll.Enabled = True
        HScroll.Max = picMap.Width \ CellSize
    Else
        HScroll.Enabled = False
    End If
    If picMap.Height > picView.Height Then
        VScroll.Enabled = True
        VScroll.Max = picMap.Height \ CellSize
    Else
        VScroll.Enabled = False
    End If
End Sub

Private Sub txtHeight_Validate(Cancel As Boolean)
    If Not IsNumeric(txtHeight.Text) Then
        MsgBox "Only numbers please!", , "PacWorld"
        txtHeight.SelStart = 0
        txtHeight.SelLength = Len(txtHeight.Text)
        Cancel = True
    ElseIf Val(txtHeight.Text) > 200 Then
        MsgBox "Height can't be more than 200.", , "PacWorld"
        txtHeight.SelStart = 0
        txtHeight.SelLength = Len(txtHeight.Text)
        Cancel = True
    Else
        EMapHeight = Val(txtHeight.Text)
        picMap.Height = EMapHeight * CellSize + 2
    End If
End Sub

Private Sub txtWidth_Validate(Cancel As Boolean)
    If Not IsNumeric(txtWidth.Text) Then
        MsgBox "Only numbers please!", , "PacWorld"
        txtWidth.SelStart = 0
        txtWidth.SelLength = Len(txtWidth.Text)
        Cancel = True
    ElseIf Val(txtWidth.Text) > 200 Then
        MsgBox "Width can't be more than 200.", , "PacWorld"
        txtWidth.SelStart = 0
        txtWidth.SelLength = Len(txtHeight.Text)
        Cancel = True
    Else
        EMapWidth = Val(txtWidth.Text)
        picMap.Width = EMapWidth * CellSize + 2
    End If
End Sub

Private Sub VScroll_Change()
    picMap.Top = -VScroll.Value * CellSize
End Sub
