VERSION 5.00
Begin VB.Form CtrlPanel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "控制器"
   ClientHeight    =   1572
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   4932
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   131
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   411
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Start 
      Caption         =   "开始模拟"
      Default         =   -1  'True
      Height          =   612
      Left            =   3720
      TabIndex        =   14
      Top             =   840
      Width           =   1092
   End
   Begin VB.Frame Frame7 
      Caption         =   "抖动速度"
      Height          =   612
      Left            =   2520
      TabIndex        =   12
      Top             =   840
      Width           =   1092
      Begin VB.TextBox ShakeSpeed 
         Height          =   264
         Left            =   120
         MaxLength       =   3
         TabIndex        =   13
         Text            =   "0.2"
         Top             =   240
         Width           =   852
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "人行速度"
      Height          =   612
      Left            =   1320
      TabIndex        =   10
      Top             =   840
      Width           =   1092
      Begin VB.TextBox WalkSpeed 
         Height          =   264
         Left            =   120
         MaxLength       =   3
         TabIndex        =   11
         Text            =   "2.0"
         Top             =   240
         Width           =   852
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "场地人数"
      Height          =   612
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1092
      Begin VB.TextBox PersonCount 
         Height          =   264
         Left            =   120
         MaxLength       =   3
         TabIndex        =   9
         Text            =   "5"
         Top             =   240
         Width           =   852
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "扇形半径"
      Height          =   612
      Left            =   3720
      TabIndex        =   6
      Top             =   120
      Width           =   1092
      Begin VB.TextBox SectorRadius 
         Height          =   264
         Left            =   120
         MaxLength       =   3
         TabIndex        =   7
         Text            =   "720"
         Top             =   240
         Width           =   852
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "弧角大小"
      Height          =   612
      Left            =   2520
      TabIndex        =   4
      Top             =   120
      Width           =   1092
      Begin VB.TextBox MaxAngle 
         Height          =   264
         Left            =   120
         MaxLength       =   3
         TabIndex        =   5
         Text            =   "120"
         Top             =   240
         Width           =   852
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "通道宽度"
      Height          =   612
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   1092
      Begin VB.TextBox PsgwayWidth 
         ForeColor       =   &H00000000&
         Height          =   264
         Left            =   120
         TabIndex        =   3
         Text            =   "80"
         Top             =   240
         Width           =   852
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "人体半径"
      Height          =   612
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1092
      Begin VB.TextBox HumanRadius 
         Height          =   264
         Left            =   120
         MaxLength       =   3
         TabIndex        =   1
         Text            =   "20"
         Top             =   240
         Width           =   852
      End
   End
End
Attribute VB_Name = "CtrlPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private HumanRadius_prev As String
Private PsgwayWidth_prev As String
Private MaxAngle_prev As String
Private SectorRadius_prev As String
Private PersonCount_prev As String
Private WalkSpeed_prev As String
Private ShakeSpeed_prev As String

Private Sub Form_Load()
    HumanRadius_prev = HumanRadius
    PsgwayWidth_prev = PsgwayWidth
    MaxAngle_prev = MaxAngle
    SectorRadius_prev = SectorRadius
    PersonCount_prev = PersonCount
    WalkSpeed_prev = WalkSpeed
    ShakeSpeed_prev = ShakeSpeed
    Me.Move 32, 32
    Simulator.Show
    Simulator.Move Me.Left + Me.Width, Me.Top
    RefreshSimulator
End Sub
Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub HumanRadius_Change()
    If isNumeric(HumanRadius) Then
        RefreshSimulator
    Else
        HumanRadius = HumanRadius_prev
    End If
End Sub
Private Sub HumanRadius_GotFocus()
    HumanRadius.SelStart = 0
    HumanRadius.SelLength = Len(HumanRadius)
End Sub

Private Sub PsgwayWidth_Change()
    If isNumeric(PsgwayWidth) Then
        RefreshSimulator
    Else
        PsgwayWidth = PsgwayWidth_prev
    End If
End Sub
Private Sub PsgwayWidth_GotFocus()
    PsgwayWidth.SelStart = 0
    PsgwayWidth.SelLength = Len(PsgwayWidth)
End Sub

Private Sub MaxAngle_Change()
    If isNumeric(MaxAngle) Then
        RefreshSimulator
    Else
        MaxAngle = MaxAngle_prev
    End If
End Sub
Private Sub MaxAngle_GotFocus()
    MaxAngle.SelStart = 0
    MaxAngle.SelLength = Len(MaxAngle)
End Sub

Private Sub SectorRadius_Change()
    If isNumeric(SectorRadius) Then
        RefreshSimulator
    Else
        SectorRadius = SectorRadius_prev
    End If
End Sub
Private Sub SectorRadius_GotFocus()
    SectorRadius.SelStart = 0
    SectorRadius.SelLength = Len(SectorRadius)
End Sub

Private Sub PersonCount_Change()
    If isNumeric(PersonCount) Then
        RefreshSimulator
    Else
        PersonCount = PersonCount_prev
    End If
End Sub
Private Sub PersonCount_GotFocus()
    PersonCount.SelStart = 0
    PersonCount.SelLength = Len(PersonCount)
End Sub

Private Sub WalkSpeed_Change()
    If isNumeric(WalkSpeed) Then
        RefreshSimulator
    Else
        WalkSpeed = WalkSpeed_prev
    End If
End Sub
Private Sub WalkSpeed_GotFocus()
    WalkSpeed.SelStart = 0
    WalkSpeed.SelLength = Len(WalkSpeed)
End Sub

Private Sub ShakeSpeed_Change()
    If isNumeric(ShakeSpeed) Then
        RefreshSimulator
    Else
        ShakeSpeed = ShakeSpeed_prev
    End If
End Sub
Private Sub ShakeSpeed_GotFocus()
    ShakeSpeed.SelStart = 0
    ShakeSpeed.SelLength = Len(ShakeSpeed)
End Sub


Private Sub RefreshSimulator()
    With Simulator
        .Width = (SectorRadius * Sin(AngleToRadian(MaxAngle) / 2) * 2 + PsgwayWidth) * Screen.TwipsPerPixelX
        .Height = (HumanRadius * 3 + SectorRadius + 64) * Screen.TwipsPerPixelY
        
        .LeftWall.Left = 0
        .LeftWall.Width = SectorRadius * Sin(AngleToRadian(MaxAngle) / 2)
        .RightWall.Width = SectorRadius * Sin(AngleToRadian(MaxAngle) / 2)
        .RightWall.Left = .ScaleWidth - .RightWall.Width
        
        .LeftWall.Height = HumanRadius * 3
        .RightWall.Height = HumanRadius * 3
        .fps.Top = .LeftWall.Height + 8
        
        .LeftEnd.Y1 = .LeftWall.Height
        .RightEnd.Y1 = .RightWall.Height
        .LeftEnd.X1 = .LeftWall.Width
        .RightEnd.X1 = .RightWall.Left
        .LeftEnd.X2 = 0
        .RightEnd.X2 = .ScaleWidth
        .LeftEnd.Y2 = SectorRadius * Cos(AngleToRadian(MaxAngle) / 2) + .LeftWall.Height
        .RightEnd.Y2 = SectorRadius * Cos(AngleToRadian(MaxAngle) / 2) + .LeftWall.Height
    End With
End Sub

Private Sub Start_Click()
    Simulator.Start HumanRadius, PsgwayWidth, MaxAngle, SectorRadius, PersonCount, WalkSpeed, ShakeSpeed
End Sub
