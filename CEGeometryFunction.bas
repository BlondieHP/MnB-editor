Attribute VB_Name = "CEGeometryFunction"
Public Const Pi = 3.14159265358979
Public Const CEG_PARALLEL = 11
Public Const CEG_INFINITY = 6
Public Const CEG_VACANT = 0
Public PosO As tPoint

Public Type tPoint
X As Single
Y As Single
End Type

Public Type tRectangle
Top As Single
Left As Single
Height As Single
Width As Single
End Type

Public Type tCircler
Cnt As tPoint
Radius As Single
End Type

Public Type tLinekb       'y=kx+b
k As Single
b As Single
End Type

Public Type tLineABC      'Ax+By+C=0
a As Single               'k=-A/B
b As Single               'b=-C/B
c As Single
End Type

Public Type tSize
Height As Single
Width As Single
End Type

Public Type tRecEx
Cnt As tPoint
Size As tSize
Degree As Single
End Type

Public Type tDeviation
dX As Single
dY As Single
End Type

Public Type tLineSegment
Pos(1) As tPoint
End Type

Public Type tSpeedD
Value As Single
Direction As Single
End Type

Public Type tSpeedXY
Vx As Single
Vy As Single
End Type

Public Type tInterval
Num(1) As Single
End Type

Public Function GetDistance(Pos1 As tPoint, Pos2 As tPoint) As Single
'tPoint
Dim H As Single, l As Single
H = (Pos1.Y - Pos2.Y) ^ 2
l = (Pos1.X - Pos2.X) ^ 2
GetDistance = Sqr(H + l)
End Function

Public Function IsPInRec(Pos As tPoint, LimitArea As tRectangle) As Boolean
'tPoint|tRectangle
If Pos.X > LimitArea.Left And Pos.X < LimitArea.Left + LimitArea.Width _
And Pos.Y > LimitArea.Top And Pos.Y < LimitArea.Top + LimitArea.Height Then
IsPInRec = True
Else
IsPInRec = False
End If
End Function

Public Function IsPinC(Pos As tPoint, LimitArea As tCircler) As Boolean
'tPoint|tCircler
Dim Distance As Single, Pose As tPoint
Pose.X = LimitArea.Cnt.X
Pose.Y = LimitArea.Cnt.Y
Distance = GetDistance(Pos, Pose)
If Distance < LimitArea.Radius Then
IsPinC = True
Else
IsPinC = False
End If
End Function

Public Function IsCOverC(Cir1 As tCircler, Cir2 As tCircler) As Boolean
Dim Distance As Single
Distance = GetDistance(Cir1.Cnt, Cir2.Cnt)

If Distance < Cir1.Radius + Cir2.Radius Then
IsCOverC = True
Else
IsCOverC = False
End If
End Function

Public Function IsSeparate(Rec1 As tRectangle, Rec2 As tRectangle) As Boolean
If Rec1.Top < Rec2.Top And Rec1.Top + Rec1.Height < Rec2.Top Then
IsSeparate = True
Exit Function
End If
If Rec2.Top < Rec1.Top And Rec2.Top + Rec2.Height < Rec1.Top Then
IsSeparate = True
Exit Function
End If
If Rec1.Left < Rec2.Left And Rec1.Left + Rec1.Width < Rec2.Left Then
IsSeparate = True
Exit Function
End If
If Rec2.Left < Rec1.Left And Rec2.Left + Rec2.Width < Rec1.Left Then
IsSeparate = True
Exit Function
End If
IsSeparate = False
End Function

Public Function RadToDeg(Rad As Single) As Single
RadToDeg = Rad * (180 / Pi)
End Function

Public Function DegToRad(Deg As Single) As Single
DegToRad = Deg * (Pi / 180)
End Function

Public Function SinD(Num As Single) As Single
SinD = Sin(DegToRad(Num))
End Function

Public Function CosD(Num As Single) As Single
CosD = Cos(DegToRad(Num))
End Function

Public Function TanD(Num As Single) As Single
TanD = Tan(DegToRad(Num))
End Function

Public Function AtnD(Num As Single) As Single
AtnD = RadToDeg(Atn(Num))
End Function

Public Function GetDegree(Pos As tPoint, Cnt As tPoint) As Single
'tPoint
If Pos.X = Cnt.X And Pos.Y = Cnt.Y Then
GetDegree = 0: Exit Function
End If
Dim Cita As Single
If Pos.X - Cnt.X = 0 Then
   If Pos.Y >= Cnt.Y Then
   GetDegree = 0.5 * Pi
   Else
   GetDegree = 1.5 * Pi
   End If
   Exit Function
End If
Cita = Atn((Pos.Y - Cnt.Y) / (Pos.X - Cnt.X))

If Pos.X < Cnt.X Then
Cita = Cita - Pi
End If
GetDegree = Cita
End Function

Public Function RecExtoRec(RecEx As tRecEx) As tRectangle
With RecExtoRec
    .Top = RecEx.Cnt.Y - 0.5 * RecEx.Size.Height
    .Left = RecEx.Cnt.X - 0.5 * RecEx.Size.Width
    .Height = RecEx.Size.Height
    .Width = RecEx.Size.Width
End With
End Function

Public Function PoltoRec(Lenth As Single, Degree As Single, Cnt As tPoint) As tPoint
With PoltoRec
   .Y = Lenth * Sin(Degree) + Cnt.Y
   .X = Lenth * Cos(Degree) + Cnt.X
End With
End Function

Public Function RotatePoint(Pos As tPoint, Cnt As tPoint, Degree As Single) As tPoint
Dim Cita As Single, l As Single
Cita = GetDegree(Pos, Cnt)
Cita = Cita + Degree
l = GetDistance(Pos, Cnt)
RotatePoint = PoltoRec(l, Cita, Cnt)
End Function

Public Function IsPinRecEx(Pos As tPoint, RecEx As tRecEx) As Boolean
Dim TemRec As tRectangle, TemP As tPoint, l As Single
TemRec = RecExtoRec(RecEx)
l = GetDistance(RecEx.Cnt, Pos)
TemP = PoltoRec(l, -RecEx.Degree, RecEx.Cnt)
IsPinRecEx = IsPInRec(TemP, TemRec)
End Function

Public Function GetLineSegmentFunction(LineS As tLineSegment, oFunction As tLinekb) As Long
If LineS.Pos(0).X = LineS.Pos(1).X Then
GetLineSegmentFunction = 0: Exit Function
End If
oFunction.k = (LineS.Pos(0).Y - LineS.Pos(1).Y) / (LineS.Pos(0).X - LineS.Pos(1).X)
oFunction.b = LineS.Pos(0).Y - oFunction.k * LineS.Pos(0).X
GetLineSegmentFunction = 1
End Function

Public Function IsPbetween2PX(Cnt As tPoint, Pos1 As tPoint, Pos2 As tPoint, Optional IsOpen As Boolean = False) As Boolean
If Not IsOpen Then
IsPbetween2PX = (Cnt.X - Pos1.X) * (Cnt.X - Pos2.X) <= 0
Else
IsPbetween2PX = (Cnt.X - Pos1.X) * (Cnt.X - Pos2.X) < 0
End If
End Function

Public Function IsPbetween2PY(Cnt As tPoint, Pos1 As tPoint, Pos2 As tPoint, Optional IsOpen As Boolean = False) As Boolean
If Not IsOpen Then
IsPbetween2PY = (Cnt.Y - Pos1.Y) * (Cnt.Y - Pos2.Y) <= 0
Else
IsPbetween2PY = (Cnt.Y - Pos1.Y) * (Cnt.Y - Pos2.Y) < 0
End If
End Function

Public Function GetCrossPoint(Line1 As tLinekb, Line2 As tLinekb) As tPoint
With GetCrossPoint
    .X = (Line2.b - Line1.b) / (Line1.k - Line2.k)
    .Y = Line1.k * .X + Line1.b
End With
End Function

Public Function IsPbetween2P(Cnt As tPoint, Pos1 As tPoint, Pos2 As tPoint, Optional IsOpen As Boolean = False) As Boolean
IsPbetween2P = IsPbetween2PY(Cnt, Pos1, Pos2, IsOpen) And IsPbetween2PX(Cnt, Pos1, Pos2, IsOpen)
End Function

Public Function IsPonLineS(Pos As tPoint, LineS As tLineSegment, Optional Accuracy As Single = 0.0001) As Boolean
Dim TemL As tLinekb, s As Long, Y As Single, TemY As Single
s = GetLineSegmentFunction(LineS, TemL)
   If s = 1 Then
   Y = Int((TemL.k * Pos.X + TemL.b) / Accuracy) * Accuracy
   TemY = Int(Pos.Y / Accuracy) * Accuracy
   IsPonLineS = (Y = TemY) And IsPbetween2P(Pos, LineS.Pos(0), LineS.Pos(1))
   Else
   IsPonLineS = (Int(Pos.X / Accuracy) * Accuracy = Int(LineS.Pos(0).X / Accuracy) * Accuracy) And IsPbetween2P(Pos, LineS.Pos(0), LineS.Pos(1))
   End If
End Function

Public Function PEPonLineS(Pos1 As tPoint, Pos2 As tPoint, LineS As tLineSegment, Arrival As tPoint, Optional Accuracy As Single = 0.001) As Long
Dim TemLS As tLineSegment, TemP As tPoint, TemL(0 To 1) As tLineABC, s(0 To 1) As Long, ER As Long
TemLS.Pos(0) = Pos1: TemLS.Pos(1) = Pos2
TemL(0) = GetLineSegmentFunctionABC(TemLS)
TemL(1) = GetLineSegmentFunctionABC(LineS)
TemP = GetCrossPABC(TemL(0), TemL(1), , ER)
   
   If IsPbetween2P(TemP, LineS.Pos(0), LineS.Pos(1), False) And IsPbetween2P(TemP, Pos1, Pos2, False) Then
       If ER = CEG_VACANT Then
       Arrival = TemP: PEPonLineS = 1
       ElseIf ER = CEG_PARALLEL Then
       Arrival = Pos2
       End If
   Else
   Arrival = Pos2: PEPonLineS = 0
   Exit Function
   End If
End Function

Public Sub DegreeStandardize(Degree As Single)
Do While Degree < 0
   Degree = Degree + 2 * Pi
Loop
Do While Degree >= 2 * Pi
   Degree = Degree - 2 * Pi
Loop
End Sub


Public Function FuncDegreeStandardize(ByVal Degree As Single) As Single
FuncDegreeStandardize = Degree
Do While FuncDegreeStandardize < 0
   FuncDegreeStandardize = FuncDegreeStandardize + 2 * Pi
Loop
Do While FuncDegreeStandardize >= 2 * Pi
   FuncDegreeStandardize = FuncDegreeStandardize - 2 * Pi
Loop
End Function

Public Function DegreeStandardize2(Degree As Single) As Single
Do While Degree < -Pi
   Degree = Degree + 2 * Pi
Loop
Do While Degree > Pi
   Degree = Degree - 2 * Pi
Loop
End Function

Public Function FuncDegreeStandardize2(Degree As Single) As Single
FuncDegreeStandardize2 = Degree
Do While FuncDegreeStandardize2 < -Pi
   FuncDegreeStandardize2 = FuncDegreeStandardize2 + 2 * Pi
Loop
Do While FuncDegreeStandardize2 > Pi
   FuncDegreeStandardize2 = FuncDegreeStandardize2 - 2 * Pi
Loop
End Function

Public Function DegreeStandardizeinPi(Degree As Single) As Single
DegreeStandardize Degree
Do While Degree >= Pi
   Degree = Degree - Pi
Loop
End Function

Public Sub DegreeStandardizeinRight(Degree As Single)
DegreeStandardize Degree
Do While Degree > Pi / 2
   Degree = Degree - Pi
Loop

Do While Degree < -Pi / 2
   Degree = Degree + Pi
Loop

End Sub

Public Function FuncDegreeStandardizeinRight(ByVal Degree As Single) As Single
FuncDegreeStandardizeinRight = Degree
Do While FuncDegreeStandardizeinRight > Pi / 2
   FuncDegreeStandardizeinRight = FuncDegreeStandardizeinRight - Pi
Loop

Do While FuncDegreeStandardizeinRight < -Pi / 2
   FuncDegreeStandardizeinRight = FuncDegreeStandardizeinRight + Pi
Loop

End Function

Public Function DegreeStandardizeinDown(Degree As Single) As Single
DegreeStandardize Degree
If Degree < 0 Then
Degree = -Degree
End If
End Function

Public Sub DrawRad(Paper As PictureBox, Cnt As tPoint, Radius As Single, SRad As Single, ERad As Single, m As Single, Optional Color As Long = vbBlack, Optional AntiRec As Single = 0.1, Optional LastPointX As Single, Optional LastPointY As Single)
DegreeStandardize SRad: DegreeStandardize ERad
Dim LastP As tPoint, i As Single, TemP As tPoint
LastP = PoltoRec(Radius, SRad, Cnt)
If SRad > ERad Then SRad = SRad - 2 * Pi
Dim H As Single, w As Single
w = Radius: H = Radius * m
For i = SRad To ERad Step AntiRec
    TemP.X = w * Cos(i) + Cnt.X
    TemP.Y = H * Sin(i) + Cnt.Y
    Paper.Line (LastP.X, LastP.Y)-(TemP.X, TemP.Y), Color
    LastP = TemP
Next i

TemP.X = w * Cos(ERad) + Cnt.X
TemP.Y = H * Sin(ERad) + Cnt.Y
    
Paper.Line (LastP.X, LastP.Y)-(TemP.X, TemP.Y), Color

If Not IsMissing(LastPointX) Then
LastPointX = LastP.X
End If
If Not IsMissing(LastPointY) Then
LastPointY = LastP.Y
End If
End Sub

Public Function GetSign(Num As Single) As Integer
If Num > 0 Then
GetSign = 1
ElseIf Num < 0 Then
GetSign = -1
ElseIf Num = 0 Then
GetSign = 0
End If
End Function

Public Function GetMidPoint(Pos1 As tPoint, Pos2 As tPoint) As tPoint
With GetMidPoint
   .X = (Pos1.X + Pos2.X) / 2
   .Y = (Pos1.Y + Pos2.Y) / 2
End With
End Function

Public Function SpeedDtoXY(V As tSpeedD) As tSpeedXY
With SpeedDtoXY
    .Vy = V.Value * Sin(V.Direction)
    .Vx = V.Value * Cos(V.Direction)
End With
End Function

Public Function SpeedXYtoD(V As tSpeedXY) As tSpeedD
Dim vP As tPoint, tP As tPoint
With SpeedXYtoD
               .Value = Sqr(V.Vy ^ 2 + V.Vx ^ 2)
               If V.Vx <> 0 Then
               tP.X = V.Vx
               tP.Y = V.Vy
               .Direction = GetDegree(tP, vP)
               Else
               .Direction = GetSign(V.Vy) * Pi / 2
               End If
End With
End Function

Public Function GetDistancePnLine(Pos As tPoint, Linekb As tLinekb) As Single
GetDistancePnLine = Abs(Linekb.k * Pos.X - Pos.Y + Linekb.b) / Sqr(Linekb.k ^ 2 + 1)
End Function

Public Function IsPbetween2PEx(Pos As tPoint, LineS As tLineSegment) As Boolean
Dim tP As tPoint, tSP As tPoint, tD As Single
tD = GetDegree(LineS.Pos(1), LineS.Pos(0))
tSP = RotatePoint(LineS.Pos(1), LineS.Pos(0), -tD)
tP = RotatePoint(Pos, LineS.Pos(0), -tD)
IsPbetween2PEx = IsPbetween2PX(tP, tSP, LineS.Pos(0))
End Function

Public Function GetDistancePnLineABC(Pos As tPoint, LineS As tLineSegment) As Single
Dim TL As tLineABC
TL = GetLineSegmentFunctionABC(LineS)

With TL
     GetDistancePnLineABC = Abs(.a * Pos.X + .b * Pos.Y + .c) / Sqr(.a ^ 2 + .b ^ 2)
End With
End Function

Public Function GetDistancePnLineS(Pos As tPoint, LineS As tLineSegment) As Single
Dim TL As tLinekb, tLen1 As Single, tLen2 As Single
If IsPbetween2PEx(Pos, LineS) Then
    If GetLineSegmentFunction(LineS, TL) = False Then
    GetDistancePnLineS = Abs(LineS.Pos(0).X - Pos.X)
    Else
    GetDistancePnLineS = GetDistancePnLine(Pos, TL)
    End If
Else
   tLen1 = GetDistance(LineS.Pos(0), Pos)
   tLen2 = GetDistance(LineS.Pos(1), Pos)
      If tLen1 < tLen2 Then
      GetDistancePnLineS = tLen1
      Else
      GetDistancePnLineS = tLen2
      End If
End If

End Function

Public Function Accuracize(Value As Single, Accuracy As Single) As Single
Accuracize = Int(Value / Accuracy) * Accuracy
End Function

Public Function SetPonLineX(Pos As tPoint, Linekb As tLinekb) As tPoint
SetPonLineX.Y = Linekb.k * Pos.X + Linekb.b
SetPonLineX.X = Pos.X
End Function

Public Function GetLineSegmentFunctionABC(LineS As tLineSegment) As tLineABC
With GetLineSegmentFunctionABC
     .a = LineS.Pos(1).Y - LineS.Pos(0).Y
     .b = LineS.Pos(0).X - LineS.Pos(1).X
     .c = -LineS.Pos(0).Y * .b - .a * LineS.Pos(0).X
End With
End Function

Public Function GetVerticalLineABC(LineABC As tLineABC, CrossP As tPoint) As tLineABC
With GetVerticalLineABC
    .a = -LineABC.b
    .b = LineABC.a
    .c = -.a * CrossP.X - .b * CrossP.Y
End With
End Function

Public Function GetCrossPABC(Line1 As tLineABC, Line2 As tLineABC, Optional Pixelize As Boolean = False, Optional ErrReason As Long = CEG_VACANT) As tPoint
With GetCrossPABC
    If Line1.a * Line2.b = Line2.a * Line1.b Then
    ErrReason = CEG_PARALLEL: Exit Function
    Else
    ErrReason = CEG_VACANT
    .X = (Line1.b * Line2.c - Line2.b * Line1.c) / (Line1.a * Line2.b - Line2.a * Line1.b)
    .Y = (Line2.a * Line1.c - Line1.a * Line2.c) / (Line1.a * Line2.b - Line2.a * Line1.b)
    End If
End With
If Pixelize Then
PointPixelize GetCrossPABC
End If
End Function

Public Function Swap(Value1 As Variant, Value2 As Variant) As Long
On Error GoTo Errline
Dim t As Variant
t = Value1
Value1 = Value2
Value2 = t
Swap = 1
Exit Function
Errline:
Swap = 0
End Function


Public Function GetVerticalDegree(Pos1 As tPoint, LineS As tLineSegment) As Single
Dim tD As Single, tP As tPoint, tS As Integer
tP = Pos1
tD = GetDegree(LineS.Pos(1), LineS.Pos(0))
tP = RotatePoint(tP, LineS.Pos(0), -tD)
tS = GetSign(tP.Y - LineS.Pos(0).Y)
GetVerticalDegree = tD + tS * Pi / 2
End Function

Public Sub PointPixelize(Pos As tPoint)
With Pos
   .X = Int(.X)
   .Y = Int(.Y)
End With
End Sub


Public Function PEPonLineSEx(Pos1 As tPoint, Pos2 As tPoint, LineS As tLineSegment, Arrival As tPoint, Optional Pixelize As Boolean, Optional Accuracy As Single = 0.001) As Long
Dim LSF As tLineABC, CP As tPoint, Reason As Long, PPF As tLineABC, tLS As tLineSegment

tLS.Pos(0) = Pos1
tLS.Pos(1) = Pos2
PPF = GetLineSegmentFunctionABC(tLS)
LSF = GetLineSegmentFunctionABC(LineS)
CP = GetCrossPABC(LSF, PPF, Pixelize, Reason)
    If Reason = CEG_PARALLEL Then
    Arrival = Pos2
    PEPonLineSEx = 0: GoTo Last
    Else
        If GetDistancePnLineS(CP, LineS) < 1 * Accuracy And IsPbetween2P(CP, Pos1, Pos2, False) Then
        Dim FP As tPoint, vL As tLineABC
        vL = GetVerticalLineABC(LSF, Pos2)
        FP = GetCrossPABC(LSF, vL, True)
            Dim tD As Single
            tD = GetVerticalDegree(Pos1, LineS)
            Arrival = PoltoRec(1 * Accuracy, tD, FP)
            PEPonLineSEx = 1
        Else
        Arrival = Pos2
        PEPonLineSEx = 0: GoTo Last
        End If
    End If
Last:
If Pixelize Then
  PointPixelize Arrival
End If
End Function

Public Function PEPConPoint(Cir As tCircler, Pos As tPoint, Arrival As tCircler, Optional Offset As Single = 1) As Single
Dim tD As Single
D = GetDistance(Pos, Cir.Cnt)
Arrival = Cir

If D < Cir.Radius Then
     PEPConPoint = Cir.Radius - D
     Arrival = Cir
     tD = GetDegree(Cir.Cnt, Pos)
     Arrival.Cnt = PoltoRec(Cir.Radius + Offset, tD, Pos)
End If

End Function

Public Function PEPConPointEx(Cir1 As tCircler, Cir2 As tCircler, Pos As tPoint, Arrival As tCircler, Optional Offset As Single = 1) As Single
Dim D As Single, tLS As tLineSegment, tD As Single, tC As tCircler, tP As tPoint, l As Single

If (Cir1.Cnt.X <> Cir2.Cnt.X) And (Cir1.Cnt.Y <> Cir2.Cnt.Y) Then
Arrival = Cir2
tLS = SetLineSegement(Cir1.Cnt.X, Cir1.Cnt.Y, Cir2.Cnt.X, Cir2.Cnt.Y)
D = GetDistancePnLineS(Pos, tLS)

If D < Cir2.Radius Then
     PEPConPointEx = Cir2.Radius - D
     
     'Rotate
      tD = GetDegree(Cir2.Cnt, Cir1.Cnt)
     
     tC.Radius = Cir2.Radius
     tC.Cnt = RotatePoint(Cir2.Cnt, Cir1.Cnt, -tD)
     tP = RotatePoint(Pos, Cir1.Cnt, -tD)
     
     l = Sqr(Cir1.Radius ^ 2 - D ^ 2)
     With Arrival
         .Cnt.Y = tC.Cnt.Y
         .Cnt.X = tP.X - Sgn(tC.Cnt.X - Cir1.Cnt.X) * (l + Offset)
         .Radius = Cir1.Radius
     End With
     
     Arrival.Cnt = RotatePoint(Arrival.Cnt, Cir1.Cnt, tD)
End If

Else
PEPConPointEx = PEPConPoint(Cir2, Pos, Arrival)
End If
End Function

Public Function PEPConLineS(Cir As tCircler, LineS As tLineSegment, Arrival As tCircler, Optional Offset As Single = 0.1) As Single
Dim q(2) As Single, tCir As tCircler, D As Single, r As Single, tLS As tLineSegment, Cita As Single
tCir = Cir
r = Cir.Radius
q(0) = PEPConPoint(tCir, LineS.Pos(0), tCir)
q(1) = PEPConPoint(tCir, LineS.Pos(1), tCir)

Arrival = tCir

D = GetDistancePnLineS(tCir.Cnt, LineS)

If D < r Then
q(2) = D
     Cita = GetDegree(LineS.Pos(1), LineS.Pos(0))
     With tLS
           .Pos(0) = LineS.Pos(0)
           .Pos(1) = RotatePoint(LineS.Pos(1), tLS.Pos(0), -Cita)
     End With

     With Arrival
           .Cnt = RotatePoint(tCir.Cnt, tLS.Pos(0), -Cita)
           
           .Cnt.Y = .Cnt.Y + (r - D) * Sgn(.Cnt.Y - tLS.Pos(0).Y)
           
           .Cnt = RotatePoint(Arrival.Cnt, tLS.Pos(0), Cita)
     End With
End If

PEPConLineS = q(0) Or q(1) Or q(2)
End Function

Public Function IsLineSCrossed(LineS1 As tLineSegment, LineS2 As tLineSegment) As Long
Dim RC As tPoint, RA As Single, tLS As tLineSegment, q(1) As Long
'Rotate 1
    RC = LineS1.Pos(0)
    RA = GetDegree(LineS1.Pos(1), RC)
    
    With tLS
         .Pos(0) = RotatePoint(LineS2.Pos(0), RC, -RA)
         .Pos(1) = RotatePoint(LineS2.Pos(1), RC, -RA)
    End With
    
    If (Sgn(tLS.Pos(0).Y - RC.Y) <> Sgn(tLS.Pos(1).Y - RC.Y)) Or Sgn(tLS.Pos(0).Y - RC.Y) = 0 Or Sgn(tLS.Pos(1).Y - RC.Y) = 0 Then
          q(0) = 1
    End If
'Rotate 2
    RC = LineS2.Pos(0)
    RA = GetDegree(LineS2.Pos(1), RC)
    
    With tLS
         .Pos(0) = RotatePoint(LineS1.Pos(0), RC, -RA)
         .Pos(1) = RotatePoint(LineS1.Pos(1), RC, -RA)
    End With
    
    If (Sgn(tLS.Pos(0).Y - RC.Y) <> Sgn(tLS.Pos(1).Y - RC.Y)) Or Sgn(tLS.Pos(0).Y - RC.Y) = 0 Or Sgn(tLS.Pos(1).Y - RC.Y) = 0 Then
          q(1) = 1
    End If
    
'Conclusion
IsLineSCrossed = q(0) And q(1)
End Function

Public Function PEPConLineSEx(Cir1 As tCircler, Cir2 As tCircler, LineS As tLineSegment, Arrival As tCircler, Optional Offset As Single = 0.1) As Single
Dim CP As tPoint, Reason As Long, NC1 As tCircler, NC2 As tCircler, RD As Single, NewLineS As tLineSegment, q(5) As Boolean, H As Single, r As Single
Dim Distance As Single, InvC As tInterval, InvP As tInterval, tCir2 As tCircler

tCir2 = Cir2

'Q(4) = PEPConPointEx(Cir1, tCir2, LineS.Pos(0), tCir2)
'Q(5) = PEPConPointEx(Cir1, tCir2, LineS.Pos(1), tCir2)
'PEPConLineSEx = Q(4) Or Q(5)

r = Cir1.Radius

'Crash1
RD = -GetDegree(tCir2.Cnt, Cir1.Cnt)          'C0 is the CNT

'Transform C-C Line
     NC1 = Cir1
     NC2.Cnt = RotatePoint(tCir2.Cnt, NC1.Cnt, RD)
     NC2.Radius = r

With NewLineS                                    'LineS Laydown
    .Pos(0) = RotatePoint(LineS.Pos(0), Cir1.Cnt, RD)
    .Pos(1) = RotatePoint(LineS.Pos(1), Cir1.Cnt, RD)
End With


'X_Projection
InvC = CreateCirtoCirProjectionX(NC1, NC2)
InvP = CreatePostoPosProjectionX(NewLineS.Pos(0), NewLineS.Pos(1))
q(0) = IsIntervalOverlap(InvC, InvP)

'Y_Projection
InvC = CreateCirtoCirProjectionY(NC1, NC2)
InvP = CreatePostoPosProjectionY(NewLineS.Pos(0), NewLineS.Pos(1))
q(1) = IsIntervalOverlap(InvC, InvP)


'Crash2
RD = -GetDegree(LineS.Pos(1), LineS.Pos(0))      'P0 is the CNT

With NewLineS                                    'LineS Laydown
    .Pos(0) = LineS.Pos(0)
    .Pos(1) = RotatePoint(LineS.Pos(1), LineS.Pos(0), RD)
End With

'Transform C-C Line
     NC1.Cnt = RotatePoint(Cir1.Cnt, LineS.Pos(0), RD)
     NC1.Radius = r
     NC2.Cnt = RotatePoint(tCir2.Cnt, LineS.Pos(0), RD)
     NC2.Radius = r

'X_Projection
InvC = CreateCirtoCirProjectionX(NC1, NC2)
InvP = CreatePostoPosProjectionX(NewLineS.Pos(0), NewLineS.Pos(1))
q(2) = IsIntervalOverlap(InvC, InvP)

'Y_Projection
InvC = CreateCirtoCirProjectionY(NC1, NC2)
InvP = CreatePostoPosProjectionY(NewLineS.Pos(0), NewLineS.Pos(1))
q(3) = IsIntervalOverlap(InvC, InvP)
    
    
If q(0) And q(1) And q(2) And q(3) Then
   Distance = Abs(NC2.Cnt.Y - (NewLineS.Pos(0).Y + Sgn(NC1.Cnt.Y - NewLineS.Pos(0).Y) * r)) + Offset
   
   'Circle Correction
   With Arrival.Cnt
       .X = NC2.Cnt.X
       .Y = NC2.Cnt.Y + Sgn(NC1.Cnt.Y - NewLineS.Pos(0).Y) * Distance
   End With
   Arrival.Radius = r


   'Untransform
   Arrival.Cnt = RotatePoint(Arrival.Cnt, LineS.Pos(0), -RD)

Else
   Arrival = tCir2
End If


PEPConLineSEx = Distance
End Function



Public Function LineSAddLenth(LineS As tLineSegment, AddL As Single) As tLineSegment
Dim MP As tPoint, l As Single, Deg As Single
l = GetDistance(LineS.Pos(0), LineS.Pos(1))
MP = GetMidPoint(LineS.Pos(0), LineS.Pos(1))
Deg = GetDegree(LineS.Pos(0), MP)

With LineSAddLenth
    .Pos(0) = PoltoRec(l / 2 + AddL, Deg, MP)
    .Pos(1) = PoltoRec(l / 2 + AddL, Pi + Deg, MP)
End With
End Function

Public Function PEPonCnLineS(Cir1 As tCircler, Cir2 As tCircler, LineS As tLineSegment, Arrival As tCircler, Direction As Single, Optional Accuracy As Single = 0.001) As Long
Dim Pos1 As tPoint, Pos2 As tPoint, tHD As Single, AP As tPoint, NewLineS As tLineSegment
NewLineS = LineSAddLenth(LineS, Cir1.Radius)
Arrival = Cir2
tHD = GetVerticalDegree(Cir1.Cnt, LineS) + Pi
Direction = tHD
DegreeStandardize Direction
Pos1 = PoltoRec(Cir1.Radius, tHD, Cir1.Cnt)
Pos2 = PoltoRec(Cir2.Radius, tHD, Cir2.Cnt)
PEPonCnLineS = PEPonLineS(Pos1, Pos2, NewLineS, AP, Accuracy)
Arrival.Cnt = PoltoRec(Cir2.Radius, Pi + tHD, AP)
End Function

Public Function SimplifyPEPonCnLineS(Cir1 As tCircler, Cir2 As tCircler, LineS As tLineSegment, Optional Accuracy As Single = 0.001, Optional Direction As Single) As Long
Dim Pos1 As tPoint, Pos2 As tPoint, tHD As Single, AP As tPoint, NewLineS As tLineSegment
NewLineS = LineSAddLenth(LineS, Cir1.Radius)

tHD = GetVerticalDegree(Cir1.Cnt, LineS)
Direction = tHD + Pi
DegreeStandardize Direction
Pos1 = PoltoRec(Cir1.Radius, tHD, Cir1.Cnt)
Pos2 = PoltoRec(Cir2.Radius, tHD + Pi, Cir2.Cnt)
SimplifyPEPonCnLineS = PEPonLineS(Pos1, Pos2, NewLineS, AP, Accuracy)

End Function

Public Function PEPonCircle(Pos1 As tCircler, Pos2 As tCircler, Cir As tCircler, Arrival As tCircler, Optional Accuracy As Single = 0.001) As Long
Dim TL As Single, tLS As tLineSegment
With tLS
   .Pos(0) = Pos1.Cnt
   .Pos(1) = Pos2.Cnt
End With

TL = GetDistancePnLineS(Cir.Cnt, tLS)
TL = Accuracize(TL, Accuracy)

If TL <= Cir.Radius Then
PEPonCircle = 1

End If
End Function

Public Sub DrawRadEx(Paper As PictureBox, Cnt As tPoint, Radius As Single, SRad As Single, ERad As Single, RotateDegree As Single, m As Single, Optional Color As Long = vbBlack, Optional AntiRec As Single = 0.1, Optional LastPointX As Single, Optional LastPointY As Single)
'DegreeStandardize SRad: DegreeStandardize ERad
Dim LastP As tPoint, i As Single, TemP As tPoint
LastP = PoltoRec(Radius, SRad, Cnt)
LastP = RotatePoint(LastP, Cnt, RotateDegree)
If SRad > ERad Then SRad = SRad - 2 * Pi
Dim H As Single, w As Single
w = Radius: H = Radius * m
For i = SRad To ERad Step AntiRec
    TemP.X = w * Cos(i) + Cnt.X
    TemP.Y = H * Sin(i) + Cnt.Y
    TemP = RotatePoint(TemP, Cnt, RotateDegree)
    Paper.Line (LastP.X, LastP.Y)-(TemP.X, TemP.Y), Color
    LastP = TemP
Next i
If Not IsMissing(LastPointX) Then
LastPointX = LastP.X
End If
If Not IsMissing(LastPointY) Then
LastPointY = LastP.Y
End If
End Sub

Public Function GetRelativePoint(CntRelativePoint As tPoint, Pos As tPoint, Cnt As tPoint, ByVal dScale As Single) As tPoint
Dim TemP As tPoint
TemP.X = CntRelativePoint.X + (Pos.X - Cnt.X) * dScale
TemP.Y = CntRelativePoint.Y + (Pos.Y - Cnt.Y) * -dScale

GetRelativePoint = TemP
End Function

'Absolute
Public Function GetAbsolutePoint(CntAbsolutePoint As tPoint, Pos As tPoint, Cnt As tPoint, ByVal dScale As Single) As tPoint
Dim TemP As tPoint
TemP.X = CntAbsolutePoint.X + (Pos.X - Cnt.X) / dScale
TemP.Y = CntAbsolutePoint.Y + (Pos.Y - Cnt.Y) / -dScale

GetAbsolutePoint = TemP
End Function

Public Function PointSubStract(Minuend As tPoint, Subtrahend As tPoint) As tPoint
With PointSubStract
     .X = Minuend.X - Subtrahend.X
     .Y = Minuend.Y - Subtrahend.Y
End With
End Function

Public Function SetPoint(ByVal X As Single, ByVal Y As Single) As tPoint
With SetPoint
       .X = X
       .Y = Y
End With
End Function

Public Function SetCircler(ByVal X As Single, ByVal Y As Single, ByVal r As Single) As tCircler
With SetCircler
       .Cnt.X = X
       .Cnt.Y = Y
       .Radius = r
End With
End Function

Public Function SetLineSegement(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As tLineSegment
With SetLineSegement
      .Pos(0).X = X1
      .Pos(0).Y = Y1
      .Pos(1).X = X2
      .Pos(1).Y = Y2
End With
End Function

Public Function PointDot(Dot1 As tPoint, Dot2 As tPoint) As Single
PointDot = Dot1.X * Dot2.X + Dot1.Y * Dot2.Y
End Function

Public Function PointAdd(Pos As tPoint, Cnt As tPoint) As tPoint
With PointAdd
     .X = Cnt.X + Pos.X
     .Y = Cnt.Y + Pos.Y
End With
End Function

Public Function SetPointLenth(iVector As tPoint, ByVal Lenth As Single) As tPoint
Dim tP As tPoint, tD As Single
tP = iVector
tD = GetDegree(tP, PosO)
tP = PoltoRec(Lenth, tD, PosO)
SetPointLenth = tP
End Function

Public Function SetPointDegree(iVector As tPoint, ByVal Degree As Single) As tPoint
Dim tP As tPoint, l As Single
tP = iVector
l = GetDistance(tP, PosO)
tP = PoltoRec(l, Degree, PosO)
SetPointDegree = tP
End Function

Public Sub SngSwap(Num1 As Single, Num2 As Single)
Dim t As Single
t = Num1
Num1 = Num2
Num2 = t
End Sub

Public Function IsIntervalOverlap(Inv1 As tInterval, Inv2 As tInterval) As Boolean
Dim q As Boolean

q = ((Inv2.Num(0) > Inv1.Num(0)) And (Inv2.Num(1) > Inv1.Num(0))) And ((Inv2.Num(0) > Inv1.Num(1)) And (Inv2.Num(1) > Inv1.Num(1))) Or _
    ((Inv2.Num(0) < Inv1.Num(0)) And (Inv2.Num(1) < Inv1.Num(0))) And ((Inv2.Num(0) < Inv1.Num(1)) And (Inv2.Num(1) < Inv1.Num(1)))
    
IsIntervalOverlap = Not q
End Function

Public Function IntervalAssisgnment(ByVal Num1 As Single, ByVal Num2 As Single) As tInterval
With IntervalAssisgnment
      .Num(0) = Num1
      .Num(1) = Num2
End With
End Function

Public Sub IntervalStandardize(inV As tInterval)

If inV.Num(1) < inV.Num(0) Then
    SngSwap inV.Num(0), inV.Num(1)
End If

End Sub

Public Function CreateCirtoCirProjectionX(Cir1 As tCircler, Cir2 As tCircler) As tInterval
Dim r As Single
r = Cir1.Radius
With CreateCirtoCirProjectionX
    .Num(0) = Cir1.Cnt.X + Sgn(Cir1.Cnt.X - Cir2.Cnt.X) * r
    .Num(1) = Cir2.Cnt.X + Sgn(Cir2.Cnt.X - Cir1.Cnt.X) * r
End With
IntervalStandardize CreateCirtoCirProjectionX
End Function

Public Function CreateCirtoCirProjectionY(Cir1 As tCircler, Cir2 As tCircler) As tInterval
Dim r As Single
r = Cir1.Radius
If Sgn(Cir1.Cnt.Y - Cir2.Cnt.Y) <> 0 Then
With CreateCirtoCirProjectionY
    .Num(0) = Cir1.Cnt.Y + Sgn(Cir1.Cnt.Y - Cir2.Cnt.Y) * r
    .Num(1) = Cir2.Cnt.Y + Sgn(Cir2.Cnt.Y - Cir1.Cnt.Y) * r
End With

Else
With CreateCirtoCirProjectionY
    .Num(0) = Cir1.Cnt.Y - r
    .Num(1) = Cir1.Cnt.Y + r
End With
End If

IntervalStandardize CreateCirtoCirProjectionY
End Function

Public Function CreatePostoPosProjectionX(Pos1 As tPoint, Pos2 As tPoint) As tInterval
With CreatePostoPosProjectionX
    .Num(0) = Pos1.X
    .Num(1) = Pos2.X
End With
IntervalStandardize CreatePostoPosProjectionX
End Function

Public Function CreatePostoPosProjectionY(Pos1 As tPoint, Pos2 As tPoint) As tInterval
With CreatePostoPosProjectionY
    .Num(0) = Pos1.Y
    .Num(1) = Pos2.Y
End With
IntervalStandardize CreatePostoPosProjectionY
End Function

Public Function IsSameCircler(Cir1 As tCircler, Cir2 As tCircler) As Long

 If (Cir1.Cnt.X = Cir2.Cnt.X) And (Cir1.Cnt.Y = Cir2.Cnt.Y) And (Cir1.Radius = Cir2.Radius) Then
 IsSameCircler = 1
 End If
 
End Function

Public Function PEPConLineSEx_V2000(Cir1 As tCircler, Cir2 As tCircler, LineS As tLineSegment, Arrival As tCircler, Optional Offset As Single = 0.1) As Long
Dim q(3) As Long, tCir2 As tCircler, i As Long
If IsSameCircler(Cir1, Cir2) Then
    PEPConLineSEx_V2000 = PEPConLineS(Cir2, LineS, Arrival, Offset)
Else
    tCir2 = Cir2
    
    For i = 0 To 1
      q(i) = PEPConPointEx(Cir1, tCir2, LineS.Pos(i), tCir2, Offset)
    Next i
    
    q(2) = PEPConLineS(tCir2, LineS, tCir2, Offset)
    
    Arrival = tCir2
    
    q(3) = PEPConLineSEx_V2000_Judgement(Cir1, tCir2, LineS)
    
    If q(3) Then
        PEPConLineSEx_V2000_Correction Cir1, tCir2, LineS, Arrival, Offset
    End If
    
    PEPConLineSEx_V2000 = q(0) Or q(1) Or q(2) Or q(3)
End If

End Function

Private Function PEPConLineSEx_V2000_Judgement(Cir1 As tCircler, Cir2 As tCircler, LineS As tLineSegment) As Long
Dim ML(2) As tLineSegment, q(2) As Long, i As Long
PEPConLineSEx_V2000_CreateMotionLocus Cir1, Cir2, ML(0), ML(1), ML(2)

For i = 0 To 2
   q(i) = IsLineSCrossed(ML(i), LineS)
Next i

PEPConLineSEx_V2000_Judgement = q(0) Or q(1) Or q(2)
End Function

Private Sub PEPConLineSEx_V2000_Correction(Cir1 As tCircler, Cir2 As tCircler, LineS As tLineSegment, Arrival As tCircler, Optional Offset As Single = 0.1)
Dim NewLineS As tLineSegment, RD As Single, NC1 As tCircler, NC2 As tCircler, Distance As Single, r As Single
RD = -GetDegree(LineS.Pos(1), LineS.Pos(0))      'P0 is the CNT

r = Cir1.Radius

With NewLineS                                    'LineS Laydown
    .Pos(0) = LineS.Pos(0)
    .Pos(1) = RotatePoint(LineS.Pos(1), LineS.Pos(0), RD)
End With

'Transform C-C Line
     NC1.Cnt = RotatePoint(Cir1.Cnt, LineS.Pos(0), RD)
     NC1.Radius = r
     NC2.Cnt = RotatePoint(Cir2.Cnt, LineS.Pos(0), RD)
     NC2.Radius = r


   Distance = Abs(NC2.Cnt.Y - (NewLineS.Pos(0).Y + Sgn(NC1.Cnt.Y - NewLineS.Pos(0).Y) * r)) + Offset
   
   'Circle Correction
   With Arrival.Cnt
       .X = NC2.Cnt.X
       .Y = NC2.Cnt.Y + Sgn(NC1.Cnt.Y - NewLineS.Pos(0).Y) * Distance
   End With
   Arrival.Radius = r
   
   'Untransform
   Arrival.Cnt = RotatePoint(Arrival.Cnt, LineS.Pos(0), -RD)
End Sub

Private Sub PEPConLineSEx_V2000_CreateMotionLocus(Cir1 As tCircler, Cir2 As tCircler, oLineS1 As tLineSegment, oLineS2 As tLineSegment, oLineS3 As tLineSegment)
Dim Deg As Single, r As Single
r = Cir1.Radius

Deg = GetDegree(Cir2.Cnt, Cir1.Cnt)

With oLineS1
     .Pos(0) = PoltoRec(r, Deg - Pi, Cir1.Cnt)
     .Pos(1) = PoltoRec(r, Deg, Cir2.Cnt)
End With

With oLineS2
     .Pos(0) = PoltoRec(r, Deg + Pi / 2, Cir1.Cnt)
     .Pos(1) = PoltoRec(r, Deg + Pi / 2, Cir2.Cnt)
End With

With oLineS3
     .Pos(0) = PoltoRec(r, Deg - Pi / 2, Cir1.Cnt)
     .Pos(1) = PoltoRec(r, Deg - Pi / 2, Cir2.Cnt)
End With

End Sub

Public Function IsFalseLineSegement(LineS As tLineSegment) As Boolean
IsFalseLineSegement = (LineS.Pos(0).X = LineS.Pos(1).X) And (LineS.Pos(0).Y = LineS.Pos(1).Y)
End Function
