Attribute VB_Name = "Caluculation"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Type POINT
    X As Double
    Y As Double
End Type

Public Const PI = 3.14159265358979

Public Function isNumeric(text As String) As Boolean
    On Error GoTo err:
    If " " & text = Val(text) Then
        isNumeric = True
    End If
    Exit Function
err:
    isNumeric = False
End Function

Public Function AngleToRadian(ByVal Angle As Double)        '角度转弧度
    '1° = π/180 * 1 ≈ 0.0174532925199433 rad
    AngleToRadian = Angle * 1.74532925199433E-02
End Function
Public Function RadianToAngle(ByVal Radian As Double)       '弧度转角度
    RadianToAngle = Radian * 57.2957795130823
End Function

Public Function Arcsin(Value As Double)
    Arcsin = Atn(Value / Sqr(-Value * Value + 1))
End Function

Public Function Arccos(Value As Double)
    Arccos = Atn(-Value / Sqr(-Value * Value + 1)) + 2 * Atn(1)
End Function

Public Function Arctan(Value As Double)
    Arctan = Atn(Value)
End Function
