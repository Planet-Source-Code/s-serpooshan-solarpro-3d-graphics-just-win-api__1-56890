Attribute VB_Name = "ModuleTower"
Option Explicit

'Special Geometry functions for 100 meter CoolingTowers
'Special Object #2: ('Gonbad')
'R=a+bH^2+cH^2.5+dH^3    (0<H<30) -> (50>R>0)
'a=30 , b=.61071176  , c=-.22324169 , d=.01929
DefSng A-Z
Public Const TowerDefColor = &HBAAFAF  '&H999999
Public Const TowerInsideColDef = &H444444

Sub TowerGeoParams(ByVal TowerH As Single, ByVal TowerTeta As Single, TowerR As Single, TowerFi As Single, Px As Single, Py As Single, Vn As Vector3DType)
 'Input:TowerH,TowerTeta (H:height in meter, Teta: angle by south(x axis) )
 'Output: TowerR,TowerFi,Vn,Ax,Ay
 'TowerR:Radius of tower cross section
 'TowerFi(angle of Vn by z axis) in Rad
 'Px,Py: x,y of point of tower corresponding to given H,Teta
 'it is known that Pz=H
 'Vn: normal Vector for point P on the tower
 Dim hh, ct, st, sf
 hh = Round((TowerH - 78.06041), 14)
 TowerR = 24.58 * Sqr(1 + (hh / 72.23705) ^ 2)
 If hh = 0 Then
  TowerFi = 1.5707963267949 'Pi/2 '90
 Else
  TowerFi = -Atn(8.63686871 * TowerR / hh)
  If TowerFi < 0 Then TowerFi = TowerFi + 3.14159265358979
 End If
 ct = Cos(TowerTeta): st = Sin(TowerTeta): sf = Sin(TowerFi)
 Px = TowerR * ct
 Py = TowerR * st
 Vn.dx = sf * ct
 Vn.dy = sf * st
 Vn.dz = Cos(TowerFi)
End Sub

Function TowerElementPoly(ByVal TowerH, ByVal TowerTeta, ByVal dTowerTeta, ByVal dTowerH, Optional ColRGB As Long = TowerDefColor) As Page3DType
 'Angles Must be in Rad mode.
 Dim a, x, y, z, t, r
 With TowerElementPoly
  .ColRGB = ColRGB
  .NumPoints = 4
  ReDim .Point(4)
  z = TowerH: GoSub TowerElem10
  t = TowerTeta
  .Point(1) = Point3D(r * Cos(t), r * Sin(t), z)
  t = t + dTowerTeta
  .Point(2) = Point3D(r * Cos(t), r * Sin(t), z)
  z = z + dTowerH: GoSub TowerElem10
  .Point(3) = Point3D(r * Cos(t), r * Sin(t), z)
  t = TowerTeta
  .Point(4) = Point3D(r * Cos(t), r * Sin(t), z)
  
  .Point(0) = MidPoint3D(.Point(1), .Point(3)) 'Center
 
 End With
 Exit Function
 
TowerElem10:
  a = (z - 78.06041) / 72.23705
  r = 24.58 * Sqr(1 + a * a)
Return

End Function

Function MakeTowerObj(Optional ByVal nTowerH1 As Long = 10, Optional ByVal nTowerTeta1 As Long = 40, Optional ColRGB As Long = TowerDefColor, Optional ByVal ColEdge As Long = 0, Optional ByVal iCode As Integer = 2, Optional ObjStrName As String = "Tower") As Object3DType
 Dim nTowerH As Long, nTowerTeta As Long
 Dim TowerH, TowerTeta, dTowerH, dTowerTeta
 Dim ih As Long, i As Long, c As Long
 
 nTowerH = nTowerH1: nTowerTeta = nTowerTeta1
 
 'dTowerH = 5 'delta H Tower (m)
 'nTowerTeta = 40 '(must be 4*k) Num of Parts in each cross (Teta)
 
 dTowerH = (100 - 19.4) / nTowerH
 dTowerTeta = 2 * Pi / nTowerTeta
 
 With MakeTowerObj
  ReDim Preserve .SpValues(10)
  .SpValues(9) = nTowerH
  .SpValues(10) = nTowerTeta
  
  .NameStr = ObjStrName
  .NumPages = nTowerTeta * nTowerH
  .iCode = iCode
  .ColEdge = ColEdge
  .CenterPoint = Point3D(0, 0, (19.4 * 1.3 + 100 * 1) / 2.3)
  ReDim .Page(.NumPages - 1)
  
  c = 0
  For ih = 0 To nTowerH - 1
   TowerH = 19.4 + ih * dTowerH
   For i = 0 To nTowerTeta - 1
    TowerTeta = i * dTowerTeta
    .Page(c) = TowerElementPoly(TowerH, TowerTeta, dTowerTeta, dTowerH, ColRGB)
    c = c + 1
   Next
  Next
 End With

End Function

Public Sub TowerDrawVisibleParts(TowerObj1 As Object3DType, Optional InsideCol As Long = TowerInsideColDef)
'Draw Tower Visible parts from view point
'InsideCol: specify inside color of tower for better visibility. ( if be -1 don't use it)
Dim nTowerH As Long, nTowerTeta As Long
Dim i As Long, c As Long, ih As Long, n As Long
Dim j As Long, j2 As Long, ja As Long, Col1 As Long
Dim iTeta1 As Long, ReverseH As Long
Dim VTeta As Single, VFi As Single
Dim ColEdge As Long

nTowerH = TowerObj1.SpValues(9)
nTowerTeta = TowerObj1.SpValues(10)

VTeta = View3DTeta * kRad2Deg: VFi = View3DFi * kRad2Deg
If VFi < 0 Then VTeta = VTeta + 180
iTeta1 = ((VTeta Mod 360) + 90) * nTowerTeta / 360

If iTeta1 < 0 Then iTeta1 = iTeta1 + nTowerTeta
If iTeta1 >= nTowerTeta Then iTeta1 = iTeta1 - nTowerTeta

If Abs(VFi) < 17 Then ja = 2 Else ja = 1
j2 = nTowerTeta \ 2 + 1 - ja
If InsideCol = -1 Or Abs(VFi) < 10 Then j2 = -1
If 180 - Abs(VFi) < 17 Then j2 = j2 + 2: ja = 0
If 180 - Abs(VFi) < 10 Then j2 = nTowerTeta + 1: ja = 0

With TowerObj1
  n = .NumPages
  ColEdge = .ColEdge
  If Abs(VFi) > 90 Then 'for fi [+-90 to +-180] must draw from top to bottom
   ReverseH = -1: c = n - nTowerTeta
  Else
   ReverseH = 1: c = 0
  End If
  For ih = 0 To nTowerH - 1
   i = iTeta1: j = 1
   Do
    If j > ja And j < j2 Then
     Col1 = .Page(c + i).ColRGB
     .Page(c + i).ColRGB = InsideCol
     DrawPage3D .Page(c + i), ColEdge '&H333333
     .Page(c + i).ColRGB = Col1
    Else
     DrawPage3D .Page(c + i), ColEdge
    End If
    i = i + 1: If i >= nTowerTeta Then i = 0
    j = j + 1
   Loop Until i = iTeta1
   c = c + nTowerTeta * ReverseH
  Next
End With

End Sub


Function TowerRadiation(s1 As SolarType, ByVal TowerH As Single, ByVal TowerTeta As Single, GDir As Single, Gdif As Single, Optional sIsInRad As Boolean = False) As Single
 Dim TowerR, TowerFi, ax, ay, Az
 Dim CosIncidence, CF, Fws
 Dim s As SolarType
 Dim Vn As Vector3DType, Vs As Vector3DType
 s = s1
 If Not sIsInRad Then ConvertRadDegMode s, False
 CF = s.CF
 
 Vs = MakeSunVector(s.Beta, s.fi, True)
 TowerGeoParams TowerH, TowerTeta, TowerR, TowerFi, ax, ay, Vn
 CosIncidence = CosineV1V2(Vs, Vn)
 Fws = (1 + CosineV1V2(Vn, VectorXYZ(0, 0, 1))) / 2
 
 
 GDir = GDirect(s.Beta, CosIncidence, CF)
 Gdif = Gdiffuse(s.Beta, CF, Fws)
 
 TowerRadiation = GDir + Gdif
 
End Function

Function MakeKhabgah() As Object3DType
 Dim i As Long, j As Long, z As Single, y2 As Single, dh As Single, dl As Single, c As Long, r As Long, rCol As Long
 Dim Col(1) As Long
 Col(0) = RGB(150, 150, 150): Col(1) = RGB(100, 200, 200)
 'X Is Zero In All Points
 With MakeKhabgah
  .NumPages = 2000
  ReDim .Page(.NumPages - 1)
  .ColEdge = 0
  .iCode = 4
  .NameStr = "Khabgah"
  
  dh = 6: dl = 10
  c = 0: r = 0: rCol = 0
  z = 0: y2 = 0
  Do
  r = r + 1 'floor (Row)
  z = (r - 1) * dh
  If z <= 18 Then y2 = y2 - dl
  If z = 30 Then y2 = -160
  With .Page(c)
    .ColRGB = Col(rCol)
    .NumPoints = 4
    ReDim .Point(4)
    .Point(1) = Point3D(0, 0, z)
    .Point(2) = Point3D(0, y2, z)
    .Point(3) = Point3D(0, y2, z + dh)
    .Point(4) = Point3D(0, 0, z + dh)
    .Point(0) = MidPoint3D(.Point(1), .Point(3))
    c = c + 1
    rCol = (rCol + 1) Mod 2
  End With '.page(c)
  Loop Until r = 13
  
 End With 'MakeKhabgah

End Function
