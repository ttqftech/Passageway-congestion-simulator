VERSION 5.00
Begin VB.Form Simulator 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simulator"
   ClientHeight    =   2436
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   203
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Mover 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1800
      Top             =   0
   End
   Begin VB.Shape Person 
      BorderColor     =   &H0041D900&
      BorderWidth     =   2
      FillColor       =   &H00FFAB1A&
      FillStyle       =   0  'Solid
      Height          =   372
      Index           =   0
      Left            =   3360
      Shape           =   3  'Circle
      Top             =   0
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Line RightEnd 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      X1              =   90
      X2              =   140
      Y1              =   20
      Y2              =   70
   End
   Begin VB.Line LeftEnd 
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      X1              =   60
      X2              =   10
      Y1              =   20
      Y2              =   70
   End
   Begin VB.Shape RightWall 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00135F80&
      FillStyle       =   0  'Solid
      Height          =   252
      Left            =   1320
      Top             =   0
      Width           =   372
   End
   Begin VB.Shape LeftWall 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00135F80&
      FillStyle       =   0  'Solid
      Height          =   252
      Left            =   120
      Top             =   0
      Width           =   372
   End
End
Attribute VB_Name = "Simulator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private HumanRadius As Integer
Private PsgwayWidth As Integer
Private MaxAngle As Integer
Private SectorRadius As Integer
Private PersonCount As Integer
Private WalkSpeed As Double
Private ShakeSpeed As Double
'Private Center As POINT
Private ViewCenter As POINT
Private FieldWidth As Double
Private Fieldheight As Double
Private realPoint() As POINT

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

'初始化
Public Sub Start(HumanRadius_ As Integer, PsgwayWidth_ As Integer, MaxAngle_ As Integer, SectorRadius_ As Integer, PersonCount_ As Integer, WalkSpeed_ As Double, ShakeSpeed_ As Double)
    HumanRadius = HumanRadius_
    PsgwayWidth = PsgwayWidth_
    MaxAngle = MaxAngle_
    SectorRadius = SectorRadius_
    PersonCount = PersonCount_
    WalkSpeed = WalkSpeed_
    ShakeSpeed = ShakeSpeed_
    
    ReDim realPoint(1 To PersonCount) As POINT
    FieldWidth = SectorRadius * Sin(AngleToRadian(MaxAngle) / 2) * 2 + PsgwayWidth
    Fieldheight = SectorRadius
    ViewCenter.X = FieldWidth / 2
    ViewCenter.Y = LeftWall.Height
    Randomize
    Person(0).Width = HumanRadius * 2
    Person(0).Height = HumanRadius * 2
    Person(0).Left = -Person(0).Width * 2
    Dim tempPoint As POINT
    For n = 1 To PersonCount
        Do
            tempPoint.X = (Rnd() - 0.5) * FieldWidth
            tempPoint.Y = Rnd() * Fieldheight
        Loop While isOutOfRange(tempPoint, SectorRadius, MaxAngle, PsgwayWidth)
        realPoint(n) = tempPoint
        Load Person(n)
        Person(n).Visible = True
    Next
    refreshDisplay
    Mover.Enabled = True
End Sub

'初始化人员位置时判断有没有出界
Private Function isOutOfRange(point_ As POINT, ByVal SectorRadius_ As Double, ByVal MaxAngle_ As Double, ByVal extraWidth As Double) As Boolean
    isOutOfRange = True
    If Abs(point_.X) < extraWidth / 2 Then                                                                  '在中间区域内
        If point_.Y < SectorRadius_ Then
            isOutOfRange = False
        End If
    ElseIf point_.Y / (Abs(point_.X) - Abs(wxtrawidth)) >= Abs(Tan(AngleToRadian(90 - MaxAngle_ / 2))) Then '在扇形区域内
        If point_.X ^ 2 + point_.Y ^ 2 <= SectorRadius ^ 2 Then
            isOutOfRange = False
        End If
    End If
End Function

'把数组中的人同步到画面
Private Sub refreshDisplay()
    For i = 1 To UBound(realPoint)
        Person(i).Move realPoint(i).X - HumanRadius + ViewCenter.X, realPoint(i).Y - HumanRadius + ViewCenter.Y
    Next
End Sub

'简陋的碰撞检测，返回应该回弹的距离
Private Function calcForce(ByVal index As Integer) As POINT
    Dim d As Double                         '临时变量，距离
    Dim direction As Double, depth As Double
    For i = 1 To UBound(realPoint)          '遍历所有人，找有没有碰的
        If index <> i Then
            d = Sqr((realPoint(i).X - realPoint(index).X) ^ 2 + (realPoint(i).Y - realPoint(index).Y) ^ 2)
            depth = HumanRadius * 2 - d    '嵌进去的深度
            If depth > 0 Then               '发生碰撞
                direction = -Atn(realPoint(i).Y - realPoint(index).Y) / (realPoint(i).X - realPoint(index).X)
                'Debug.Print direction
                calcForce.X = calcForce.X + Cos(direction) * WalkSpeed * depth / HumanRadius
                If calcForce.X < 0 Then Stop
                calcForce.Y = calcForce.Y + Sin(direction) * WalkSpeed * depth / HumanRadius
            End If
        End If
    Next
    
End Function

Private Sub Mover_Timer()
    ReDim nextPoint(UBound(realPoint)) As POINT
    For i = 1 To UBound(realPoint)
        Dim movement As POINT
        movement = calcForce(i)
        nextPoint(i).X = realPoint(i).X + movement.X
        nextPoint(i).Y = realPoint(i).Y + movement.Y
    Next
    For i = 1 To UBound(realPoint)
        realPoint(i) = nextPoint(i)
    Next
    refreshDisplay
End Sub


