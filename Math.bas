Attribute VB_Name = "Math"

'##############################
'# Distance betwen two points #
'##############################
Public Function Distance(X1, Y1, X2, Y2) As Integer
    h = Abs(Y1 - Y2)
    w = Abs(X1 - X2)
    Distance = Sqr((h * h) + (w * w))
End Function

'###############################################
'# Distance betwen center and edge of a circle #
'###############################################
Public Sub Rotation(Angle As Integer, Radius As Integer, ByRef Xout As Integer, ByRef Yout As Integer)
    Dim AngleT As Single

    If DegreesRotation > 359 Then
        AngleT = (Angle Mod 360) * 1.74532925199433E-02
    Else
        AngleT = Angle * 1.74532925199433E-02
    End If

    Xout = (Cos(AngleT) * Radius)
    Yout = (Sin(AngleT) * Radius)
End Sub

'###########################
'# Force betwen two points #
'###########################
Public Sub ForceBetwenPoints(X1, Y1, X2, Y2, LinkLenth, LinkFlex, IsRope, ByRef LinkStress, ByRef Xout1, ByRef Yout1, ByRef Xout2, ByRef Yout2)
On Error Resume Next
dis = Distance(X1, Y1, X2, Y2)
Xout1 = ((X1 - X2) * (LinkLenth - dis)) / 1000 / LinkFlex
Yout1 = ((Y1 - Y2) * (LinkLenth - dis)) / 1000 / LinkFlex
Xout2 = 0 - Xout1
Yout2 = 0 - Yout1

If IsRope = True And dis + 10 < LinkLenth Then
    Xout1 = 0
    Yout1 = 0
    Xout2 = 0
    Yout2 = 0
End If
'Stress Caculation
If dis > LinkLenth Then
    LinkStress = 0 - Abs(Xout1 + Yout1) * 2
Else
    LinkStress = Abs(Xout1 + Yout1) * 2
End If
End Sub

Public Function Over0(num) 'Make numbers under 0 to 0
Over0 = num
If num < 0 Then Over0 = 0
End Function

Public Function TrueOrFalse(inp) As Boolean ' covert 1 or 0 to bolean (for checkboxes)
If Val(inp) > 0 Then TrueOrFalse = True
End Function


Public Function OneOrZero(inp As Boolean) 'boolean to 1 or 0
OneOrZero = 0
If inp = True Then OneOrZero = 1
End Function
