Attribute VB_Name = "ModuleSolar"
'Solar Main Module, By S.Serpooshan, 2001

'* You Must Call
'  'SetSolarConstants' /OR/ 'main' sub
'  and
'  'SetDayOfYear' /OR/ 'SetDate'
'  sub to Set Some Params.

'**********************************************************************
'
'  x axis is in South   &   y axis is in East
'           z axis is Upward to Zenith
'
'  Beta: angle of sun by horizon(solar altitude angle)
'  Fi  : angle of solar beam's prjection on the
'        horizontal plain of earth by south axis(azimuth angle)
'  Teta: angle between normal vector of arbitrary
'        plate and solar beam. (incidence angle)
'**********************************************************************

Option Explicit

'* Note!:
'      Vector3DType is a Type Used by this module
'      but because of it is defined in another modules
'      it's definition is Remarked.
'      Remove it's Comment mark if you need!

'15 degree = 1 hour = 60 minutes (h<=>LST)
'h is + for morning & - for afternoon
'h=(12-LST/60)*15  , h:degree , LST:minutes

'at solar noon when HourAngle is zero we have Fi=0
'LCT Shiraz = LCT Tehran + DifLongitudeToEast * 4 (minutes)

'Tehran: 35.68 North , 51.43 West
'Shiraz: 29.62 North , 52.54 West

DefSng A-Z

Public Type SolarType
 CityName As String    'City Name
 DifLong2Est As Single 'difference between longitude of your local location and capital center (cities in East of center is positive) such as Shiraz,Tehran
 l As Single           'Latitude: degree of local city from equator to north (or south) -> arze joghrafiaee
 Altitude As Single           'Altitude (meter) from sea level -> is needed in hottel method
 Month As String       'Milady Month Name/Number (1-12)
 nDay As Integer       '1 to 365: day in year (1=1january)
 EOT As Single         'Equation Of Time = func. of nDay
 CF As Single          'Cloud Factor (0=Clear Sky ... 1=Full Cloudy)
 d As Single           'declination = func. of nDay : d is angle of sun beam and equator plain (-23.5 to 23.5)
 LCTstring As String   'Local Civil Time Of Center (Tehran)
 LCT As Single         'LCT in minutes
 LST As Single         'Local Solar Time in minutes
 h As Single           'hour Angle (h=0 <=> LST=12noon)
 Beta As Single        'beta (solar beam angle by hor. plain : 0-90)
 SinBeta As Single     'sine of beta
 fi As Single          'fi (solar azimuth angle: angle of projection of beam by south: -180 to +180)
 RadMode As Boolean    'True=Radian  False=Degree  Angles
End Type

Type GRadiationType
 Direct As Single
 diffuse As Single
 Reflect As Single
 Total As Single
End Type


'Note:
'because of choosing x axis on South and y axis on East:
'h & Fi are + for morning and - for afternoon

Type ConstantsSolarType
 'Daneshyar:
 GDirectC As Single
 GdiffuseA1 As Single
 GdiffuseA2 As Single
 GdiffuseA3 As Single
 'ASHRAE:
 ASHRAE_A As Single
 ASHRAE_B As Single
 ASHRAE_C As Single
 'Hottel:
 Hottel_Gon As Single
 Hottel_a0 As Single
 Hottel_a1 As Single
 Hottel_k As Single
End Type


'Remove comment marks of Vector3DType definition
'if this type is not defined in another modules!!

'Type Vector3DType
' dx As single
' dy As single
' dz As single
'End Type


Public ConstSolar As ConstantsSolarType
Public sUser As SolarType 'store user solar params in public format
Public GRadiationUser As GRadiationType
Public UserScene As Scene3DType
Public UserdMinutes As Single, UserDayStep As Long, UserTimeTotal As Single, UserDayTotal As Long, UserTraceAsCollector As Integer, UserDoEventsOnPlate As Integer
Public UserShadowCode As Integer, UserIndexObj As Long, UserIndexPage As Long
Public RoGround As Single  'Reflectance of ground
Public MethodOfSolarCalc As Integer  '(1,2,3)Determine Methof of calculating Radiation
Public DayLightSavingTime As Integer 'true means: we must sub one hour in spring & summer
Public UseStatisticDataForCF As Integer
Public FixedCF As Single, UseSpecificHorizG As Integer, SpecificHorizG As Single, SpecificHorizGdif As Single
Public nDayBeginEnd(1) As Long, nTimeBeginEnd(1) As String, TimeIsLST As Integer
Public GNormalDirect As Single, Towbeam As Single
Public Const Pi = 3.14159265358979  'pi=4*Atn(1) 'although it is not exact value of pi but in computer is better value for computation
Public Const Pi2 = 1.5707963267949         'pi/2
Public Const kRad2Deg = 57.2957795130823 '180/pi

Sub Main()
 
 'Dim point1 As Point3DType, page1 As Page3DType
 'Dim x, y, z
 'point1 = Point3D(4, 0, 1.5)
 'page1 = MakePage3D(Point3DArrayByVal(0, 0, 0, 0, 0, 0, 10, 0, 0, 10, 5, 2, 0, 5, 2))
 'TransferSelfPage3D page1, Vector3D(-20, 30, -40)
 'TransferSelfPoint3D point1, Vector3D(-20, 30, -40)
 'a = IfLinePassThroughPage(page1, VnOfPage(page1), point1, Vector3D(2, 1, -3))
 'MsgBox CStr(a)
 'End
 
 'aa$ = "S.Serpooshan": For i = 1 To Len(aa$): Print "chr(" + CStr(Asc(Mid(aa$, i, 1))) + ")+";: Next: Print
 On Error Resume Next
 Serpoosh = Chr(83) + Chr(46) + Chr(83) + Chr(101) + Chr(114) + Chr(112) + Chr(111) + Chr(111) + Chr(115) + Chr(104) + Chr(97) + Chr(110)
 ComponentG = 7 '1+2+4 = All Components:Direct+diffuse+Reflect
 TargetObjIndex = -1 'no target object
 SpCaseObjectChange = True
 
 ShouldDrawSun = True: ShouldDrawSunBeam = True: ShouldDrawAxises = True
 ShouldDrawHorizonPlate = True: ShouldDrawInActiveObjects = True
 gdi_BackgrFillCol = -10
 iObjTransform = -1 'no scene object is avalable for applying transforms
 Call SetSolarConstants
 SetColGradArray

RecordFolder = GetSetting(App.EXEName, "Settings", "RecPath", "")

 Dim r As Long
 r = Val(GetSetting("Shine", "Sec1", "R", 0))
 r = r + 1
 SaveSetting "Shine", "Sec1", "R", CStr(r)
 'ShouldRStp = r
 
 'FormTest.Show: Exit Sub
 
 MDIFormMain.Show
 'FormSolar1.Show
End Sub


'------------------------------------------------------
'********** P A R T 1 : Solar Parameters Functions ******
'------------------------------------------------------

Public Sub SetLSTasLCTstr(s As SolarType)
 'Set .LST in minutes as a func. Of StandardCenterTime
 'StandardCenterTime -> Time in Tehran (eq. "13:45:10")
 s.LST = TimeToMinutes(s.LCTstring) + s.DifLong2Est * 4 + s.EOT
End Sub

Public Sub SetLCTasLST(s As SolarType)
 Dim a
 s.LCT = s.LST - s.DifLong2Est * 4 - s.EOT  'time in minutes
 s.LCTstring = MakeTimeFromMin(s.LCT)
End Sub

Public Function FindSaiOfPlate(Vn As Vector3DType, Optional RadMode As Boolean = False) As Single
 ' Sai angle is angle between 'projection of vector Vn on the horizontal' and South
 ' x axis is on south and y axis is on East so:
 ' Sai is angle from South to East (+)
 Dim a, dx, dy, m
 dx = Vn.dx: dy = Vn.dy
 
 If dx = 0 Then
  If RadMode Then a = Pi2 Else a = 90
  a = a * Sgn(dy)
 Else
  a = Atn(dy / dx)
  If dx < 0 Then a = Pi * Sgn(dy) + a
  If Not RadMode Then a = a * kRad2Deg
 End If
 
 FindSaiOfPlate = a

End Function

Public Sub SetSolarParamsByLST(s As SolarType)
 'Find h,Beta,Fi as a function of LST
 s.h = 180 - s.LST / 4   ' or: (12 - LST/60)*15
 FindSunBetaFiByHour s
End Sub

Public Sub FindSunBetaFiByHour(s As SolarType)
 'Find Beta,Fi as a function of Hour Angle (h)
 'Fi is + in morning and - in afternoon
 'in solar noon Fi is zero same as h
 Dim h, l, d, SinB, Beta, RadMode
 RadMode = s.RadMode
 h = s.h: l = s.l: d = s.d
 If RadMode Then
  SinB = Cos(l) * Cos(h) * Cos(d) + Sin(l) * Sin(d)
 Else
  SinB = Cosd(l) * Cosd(h) * Cosd(d) + Sind(l) * Sind(d)
 End If
 Beta = ArcSin(SinB, s.RadMode)
 s.Beta = Beta
 s.SinBeta = SinB
 If RadMode Then
  s.fi = Sgn(h) * ArcCoS((Sin(Beta) * Sin(l) - Sin(d)) / (Cos(Beta) * Cos(l)), True)
 Else
  s.fi = Sgn(h) * ArcCoS((Sind(Beta) * Sind(l) - Sind(d)) / (Cosd(Beta) * Cosd(l)), False)
 End If
End Sub

Public Function FindHourByBeta(s As SolarType) As Single
 'Find 2 value for HourAngle as a function of beta
 'these are + and - of what this function return
 's.RadMode : if be True this func. assume all angles are in Rad not in degree
 Dim ch, l, d, b
 l = s.l: d = s.d: b = s.Beta
 If s.RadMode Then
  ch = (Sin(b) - Sin(l) * Sin(d)) / (Cos(l) * Cos(d))
 Else
  ch = (Sind(b) - Sind(l) * Sind(d)) / (Cosd(l) * Cosd(d))
 End If
 FindHourByBeta = ArcCoS(ch, s.RadMode)
End Function

Public Function FindLSTbyHourAngle(ByVal h As Single) As Single
 FindLSTbyHourAngle = (180# - h) * 4
End Function

Public Function FindSolarNoon(s As SolarType) As String
 Dim s1 As SolarType
 s1 = s
 s1.LST = 12 * 60
 SetLCTasLST s1
 FindSolarNoon = s1.LCTstring
End Function

Public Sub FindSunRiseSunSetNoon(s As SolarType, SunRiseTime As String, SunSetTime As String, SunNoonTime As String, DayLength As String, LSTExact() As Single)
 'LSTExact(0,1,2)=SRise,SSet,SNoon
 Dim s1 As SolarType
 s1 = s: If s1.RadMode Then ConvertRadDegMode s1, True
 s1.Beta = 0: s1.SinBeta = 0
 s1.h = Abs(FindHourByBeta(s1))
 s1.LST = FindLSTbyHourAngle(s1.h)
 SetLCTasLST s1
 LSTExact(0) = s1.LST
 SunRiseTime = s1.LCTstring
 s1.LST = FindLSTbyHourAngle(-s1.h)
 SetLCTasLST s1
 LSTExact(1) = s1.LST
 SunSetTime = s1.LCTstring
 DayLength = MakeTimeFromMin(Abs(s1.h * 8))
 SunNoonTime = FindSolarNoon(s)
 LSTExact(2) = 12 * 60

End Sub


Function MakeSunVector(ByVal Beta As Single, ByVal fi As Single, Optional RadMode As Boolean = False) As Vector3DType
 'Make a vector by Beta & Fi for the sun (from Origin to Sun)
 Dim a As Vector3DType  'sun vector
 Dim c
 If RadMode Then
  c = Cos(Beta)
  a.dx = c * Cos(fi)
  a.dy = c * Sin(fi)
  a.dz = Sin(Beta)
 Else
  c = Cosd(Beta)
  a.dx = c * Cosd(fi)
  a.dy = c * Sind(fi)
  a.dz = Sind(Beta)
 End If
 
 'a = Vector3D(1, 0.15, 0.4)

 MakeSunVector = a
 
End Function

Public Function CosineV1V2(Vs As Vector3DType, Vn As Vector3DType) As Single
 'Return Cos(Teta) where Teta is angle between Vs and Vn
 'you can use this func to find incidence angle (V1=Vsun , V2=Vn plate)
 
 Dim dx1, dy1, dz1
 Dim dx2, dy2, dz2
 Dim c, m1, m2
 dx1 = Vs.dx: dy1 = Vs.dy: dz1 = Vs.dz
 m1 = Sqr(dx1 * dx1 + dy1 * dy1 + dz1 * dz1)
 
 dx2 = Vn.dx: dy2 = Vn.dy: dz2 = Vn.dz
 m2 = Sqr(dx2 * dx2 + dy2 * dy2 + dz2 * dz2)
 If m1 = 0 Or m2 = 0 Then CosineV1V2 = 0: Exit Function 'bad case (one of the vectors are zero)
 CosineV1V2 = (dx1 * dx2 + dy1 * dy2 + dz1 * dz2) / (m1 * m2)
End Function

Public Function FindVnOpt_NSCollector(ByVal Beta, ByVal fi, Optional RadMode As Boolean = False) As Vector3DType
 With FindVnOpt_NSCollector
 .dx = 0
 If RadMode Then
  .dy = Sin(fi)
  .dz = Tan(Beta)
 Else
  .dy = Sin(fi / kRad2Deg)
  .dz = Tan(Beta / kRad2Deg)
 End If
 End With
End Function

Public Function FindTetaCosineOpt_NSCollector(s As SolarType) As Single
 'return Optimum Cos(Teta) for a N-S Collector
 'the collector must have this normal vector:
 'Vn_OptNS=(0,sin(Fi),tg(Beta))
 Dim c
 If s.RadMode Then
  c = Sqr(Sin(s.h) ^ 2 * Cos(s.d) ^ 2 + s.SinBeta ^ 2)
 Else
  c = Sqr(Sind(s.h) ^ 2 * Cosd(s.d) ^ 2 + s.SinBeta ^ 2)
 End If
 FindTetaCosineOpt_NSCollector = c
End Function




'------------------------------------------------------
'************ P A R T 2 : Radiation Functions *********
'------------------------------------------------------


Sub SetSolarConstants()  'Call this sub initially!
 Dim k As Single
 'Constants For DANESHYAR Method In Iran:
 k = 11.63 'convert cal/(hr.cm2) to W/m2
 '0.97(Dir) & 0.99(diff) may multiply to obtain light better data for a Sigma average against of errors of shorter actual day, converting,...
 With ConstSolar
  .GDirectC = 950.61  'Gsc for australia (same as Iran) is 81.738 cal/(hr.cm2)
  .GdiffuseA1 = k * 0.132 'A1,A2,A3 are .218,.299,17.269 for USA -> multiply by .604 for Iran (high altitude)
  .GdiffuseA2 = k * 0.181 * kRad2Deg
  .GdiffuseA3 = k * 10.43
 End With
   
End Sub

Public Sub SetABCforASHRAE(ByVal nDay As Long)
 Dim x As Single
 x = nDay 'Milady
 With ConstSolar
  .ASHRAE_A = 1147.58687626 + 57.49849592 * Sin(0.01742766958555 * x + 1.478218072)
  .ASHRAE_B = 0.163947122 + 0.02368994 * Sin(0.0201714862896 * x + 4.013066167)
  .ASHRAE_C = 0.120682346 + 0.017896423 * Sin(0.0202929536132 * x + 3.97985854)
 End With
End Sub

Public Sub SetabkForHottel(s As SolarType)
 '******** these are for less than 2.5 km altitude And 23km visibility. *******
 'set a,b,k using Altitude,latitude
 Dim a As Single, r0 As Single, r1 As Single, rk As Single, l As Single
 l = Abs(s.l): If s.RadMode Then l = l * kRad2Deg
 a = sUser.Altitude / 1000 'convert to km
 If a > 2.5 Then a = 2.5 'these are for less than 2.5 km altitude And 23km visibility.
 r0 = 1: r1 = 1: rk = 1
 If s.nDay > 0 Then
  If l < 25 Then 'Tropical
   r0 = 0.95: r1 = 0.98: rk = 1.02
  ElseIf l < 60 Then 'Mid-latitude
    If s.l > 0 Then 'north semisphere
     If s.nDay <= 90 Then r0 = 1.03: r1 = 1.01: rk = 1 Else If s.nDay > 6 * 30 And s.nDay < 9 * 30 Then r0 = 0.97: r1 = 0.99: rk = 1.02 'winter else summer
    Else 'South semisphere
      If s.nDay <= 90 Then r0 = 0.97: r1 = 0.99: rk = 1.02 Else If s.nDay > 6 * 30 And s.nDay < 9 * 30 Then r0 = 1.03: r1 = 1.01: rk = 1 'summer else winter
    End If
  Else
   r0 = 0.99: r1 = 0.99: rk = 1.01
  End If
 End If
 With ConstSolar
  .Hottel_a0 = r0 * (0.4237 - 0.00821 * (6 - a) * (6 - a))
  .Hottel_a1 = r1 * (0.5055 + 0.00595 * (6.5 - a) * (6.5 - a))
  .Hottel_k = rk * (0.2711 + 0.01858 * (2.5 - a) * (2.5 - a))
 End With
End Sub

Function GDirect(BetaRad As Single, CosTeta As Single, CF As Single) As Single
 'Beta: solar Beta angle in Rad
 'CosTeta: Cosine of Teta (incidence angle)
 'CF: Cloud Factor (0=clear day ... 1=full cloudy)
 'this Function return zero automatically for plates that can't be seen from the sun (they have cosTeta<0 or in some cases beta<0)
 '* if change in this or in GDirect functions MUST modify something when 'UseSpecificHorizG' is True in UserFindRadiationPlate
 Dim G As Single
 On Error Resume Next
 
 With ConstSolar
 Select Case MethodOfSolarCalc
 Case 0, 1
  'DANESHYAR Method:
  G = (1 - CF) * .GDirectC * (1 - Exp(-4.2972 * BetaRad)) * CosTeta
 Case 2
  'ASHRAE:
  If Abs(BetaRad) <= 0.001 Then GNormalDirect = 0 Else GNormalDirect = .ASHRAE_A / Exp(.ASHRAE_B / Sin(BetaRad))
  G = (1 - CF) * GNormalDirect * CosTeta
 Case 3
  'Hottel:
  With ConstSolar
   Towbeam = .Hottel_a0 + .Hottel_a1 * Exp(-.Hottel_k / Sin(BetaRad))
   G = .Hottel_Gon * Towbeam * (1 - CF) * CosTeta
  End With
 End Select
 
 If G < 0 Or CosTeta < 0 Or BetaRad < 0 Then G = 0
 GDirect = G
 
 End With
 
End Function

Function Gdiffuse(ByVal BetaRad As Single, ByVal CF As Single, ByVal Fws As Single, Optional GNormalDirectVal As Single = -1, Optional TowBeamVal As Single = -1) As Single
 'Return diffuse rad.
 'Note: if your method is 2 (ASHRAE) You Must First Call GDirect to computing GNormalDirect
 'Fws: Fwall/Sky = (1+cos(a))/2 , a=angle of Vn by z axis
 '* if change in this or in GDirect functions MUST modify something when 'UseSpecificHorizG' is True in UserFindRadiationPlate
 Dim Gd As Single, a As Single
 
 With ConstSolar
 
 Select Case MethodOfSolarCalc
 Case 1
  Gd = (.GdiffuseA1 + .GdiffuseA2 * BetaRad + .GdiffuseA3 * CF) * Fws
 Case 2
  If GNormalDirectVal = -1 Then 'user don't send computed value to this sub, so we calculate it ourselves
   If Abs(BetaRad) <= 0.001 Then GNormalDirect = 0 Else GNormalDirect = .ASHRAE_A / Exp(.ASHRAE_B / Sin(BetaRad))
  End If
  '1/CN^2(ClearNess Factor) ~= (1.130092837 + 1.062490392 * CF)
  Gd = .ASHRAE_C * GNormalDirect * Fws * (1.1 + 1.1 * CF)
 Case 3
  Dim td As Single
  If TowBeamVal = -1 Then 'user don't send computed value so calc. it
   If Abs(BetaRad) <= 0.001 Then Towbeam = .Hottel_a0 Else Towbeam = .Hottel_a0 + .Hottel_a1 * Exp(-.Hottel_k / Sin(BetaRad))
  End If
  td = 0.271 - 0.2939 * Towbeam: If td < 0 Then td = 0
  Gd = .Hottel_Gon * Sin(BetaRad) * td * Fws * (1.55 + 1.55 * CF)
 End Select
 
 End With
 
 If Gd < 0 Then Gd = 0
 Gdiffuse = Gd
End Function

Function GReflectGround(ByVal Beta As Single, ByVal SinBeta As Single, ByVal CF As Single, ByVal RoGroundVal As Single, ByVal Fwg As Single) As Single
 'Fwg: F Wall->Ground = (1-cos(alfa))/2 ,alfa=angle between z axis and normal vector of any plate
 'RoGround:Reflectance of ground (such as .2)
 Dim GDir As Single, GtH As Single 'GtH:Gtotal for a Horizontal page (ground)
 'for a horizontal plate cos(teta) is sin(beta)
 GDir = GDirect(Beta, SinBeta, CF) 'must be in 2 lines to be sure computing GNormalDirect
 GtH = GDir + Gdiffuse(Beta, CF, 1, GNormalDirect) 'total G for a horizontal plate
 GReflectGround = GtH * RoGroundVal * Fwg
End Function

'------------------------------------------------------
'*************** G E N E R A L functions **************
'------------------------------------------------------

Public Function MonthByDay(ByVal DayOfYear As Integer) As Integer
'return Month that day is in it (MILADY)
Dim v As Variant, i As Integer
v = Array(0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334, 366)
For i = 0 To 12
 If DayOfYear <= v(i + 1) Then Exit For
Next
MonthByDay = i + 1
End Function

Public Function BadTimeSyntax(STime As String) As Boolean
 'return true if the time has bad syntax
 Dim i As Long, s As String
 On Error Resume Next
 s = STime
 If InStr(s, ":") = 0 Then s = s + ":0"
 i = InStr(s, "."): If i Then s = Left(s, i - 1)
 i = Hour(s) + Minute(s) + Second(s)
 If Err Then BadTimeSyntax = True: Err.Clear Else BadTimeSyntax = False
End Function

Public Function TimeToMinutes(ByVal STime As String) As Single
 On Error Resume Next
 If InStr(STime, ":") = 0 Then STime = STime + ":0"
 TimeToMinutes = Hour(STime) * 60 + Minute(STime) + Second(STime) / 60
 If Err Then Err.Clear: TimeToMinutes = 6 * 60 'bad time -> assume 6:00 morning
End Function

Public Function MakeTimeByHourMinSec(h, m, s) As String
 MakeTimeByHourMinSec = Format(h, "00") + ":" + Format(m, "00") + ":" + Format(s, "00")
End Function

Public Function MakeTimeFromMin(m) As String
 Dim hh, mm, ss
 hh = m \ 60
 mm = m Mod 60
 ss = m - Int(m)
 MakeTimeFromMin = MakeTimeByHourMinSec(hh, mm, ss)
End Function

Sub ConvertRadDegMode(s As SolarType, Rad2Deg As Boolean)
'Convert angles(Beta,Fi,l,d,h) from radian to degree or reverse
'This may used to speed up computations
'Rad2Deg=True => Convert Rad to Deg
'Rad2Deg=False => Convert Deg to Rad
Dim b, f, d, l, h
b = s.Beta: f = s.fi: d = s.d: l = s.l: h = s.h

If Rad2Deg Then 'Radian to Degree
 s.Beta = b * kRad2Deg: s.fi = f * kRad2Deg: s.d = d * kRad2Deg: s.l = l * kRad2Deg: s.h = h * kRad2Deg
 s.RadMode = False
Else 'Degree to Radian
 s.Beta = b / kRad2Deg: s.fi = f / kRad2Deg: s.d = d / kRad2Deg: s.l = l / kRad2Deg:: s.h = h / kRad2Deg
 s.RadMode = True
End If

End Sub

Public Function SolveEqNewton(ByVal EqCase, ByVal x0 As Variant)

MdSolarSolvEq:
Select Case EqCase
Case 1:
Case 2:
End Select
Return

End Function

Public Function ArcSin(ByVal x1, Optional AnsInRad As Boolean = False) As Single
 Dim b As Single, x As Single
 x = Round(x1, 15)
 If Abs(x) = 1 Then b = Pi2 * Sgn(x) Else b = Atn(x / Sqr(Abs(1 - x * x)))
 If Not AnsInRad Then If Abs(x) = 1 Then b = 90 * Sgn(x) Else b = b * kRad2Deg
 ArcSin = b
End Function

Public Function ArcCoS(ByVal x1, Optional AnsInRad As Boolean = False) As Single
 Dim b As Single, x
 x = Round(x1, 15)
 If x = 0 Then
  If AnsInRad Then b = Pi2 Else b = 90
 Else
  b = Atn(Sqr(Abs(1 - x * x)) / x): If x < 0 Then b = b + Pi
  If Not AnsInRad Then b = b * kRad2Deg
 End If
 ArcCoS = b
End Function

Public Function Sind(ByVal x) As Single
 'return Sin of x in degree
 Sind = Sin(x / kRad2Deg)
End Function

Public Function Cosd(ByVal x) As Single
 'return Cos of x in degree
 Cosd = Cos(x / kRad2Deg)
End Function

Public Function RadToDegree(ByVal mRad As Single) As Single
 RadToDegree = mRad * kRad2Deg
End Function

Public Function DegreeToRad(ByVal mDeg As Single) As Single
 DegreeToRad = mDeg / kRad2Deg
End Function

Public Function ViewFactor12(Obj1 As Object3DType, Obj2 As Object3DType) As Single
'return ViewFactor 1 -> 2
Dim i As Long, j As Long, CosTi As Single, dAi As Single, CosTj As Single, dAj As Single, R2ij As Single
Dim Pgi As Page3DType, Pgj As Page3DType, Vni As Vector3DType, Vnj As Vector3DType, Vr As Vector3DType
Dim F12 As Single, Ft As Single, AToti As Single, n1 As Long, n2 As Long
Dim FProgress As Form
On Error Resume Next

Set FProgress = New FormProgress
FProgress.SetProperties "View Factor Calculation ..."
FProgress.Show

F12 = 0
AToti = 0
n1 = Obj1.NumPages - 1
n2 = Obj2.NumPages - 1

For i = 0 To n1
 If i Mod 5 = 0 Then If AToti > 0 Then FProgress.Label1.Caption = "Integrated F : " + CStr(Round(F12 / AToti / Pi, 5))
 FProgress.SetProgress (i + 1) / (n1 + 1)
 Pgi = Obj1.Page(i)
 dAi = AreaPage(Pgi): Vni = VnOfPage(Pgi): AToti = AToti + dAi
 Ft = 0
 For j = 0 To n2
  If j Mod 20 = 0 Or j = n2 Then
    DoEvents
    If FProgress.CancelPressed Then i = Obj1.NumPages - 1: Exit For
  End If
  Pgj = Obj2.Page(j)
  dAj = AreaPage(Pgj): Vnj = VnOfPage(Pgj)
  Vr = VectorOf2Points(Pgi.Point(0), Pgj.Point(0))
  R2ij = LengthVector(Vr, True) '|Vr| ^ 2
  CosTi = CosineV1V2(Vr, Vni)
  If CosTi > 0 Then
   CosTj = -CosineV1V2(Vr, Vnj) 'because Vr is from i to j and here we must use -Vr or use a math. rule: cos(pi-a)=-cos(a)
   If CosTj > 0 Then
    Ft = Ft + (CosTi * CosTj / R2ij) * dAj
   End If
  End If
 Next 'j
 F12 = F12 + Ft * dAi
Next 'i

If AToti > 0 Then F12 = Round(F12 / AToti / Pi, 5) Else F12 = 0

If F12 > 1 Then F12 = 1 Else If F12 < 0 Then F12 = 0

Unload FProgress
Set FProgress = Nothing

ViewFactor12 = F12

End Function

Function MonthNum(MonthStr As String, Optional IranDate As Boolean = False, Optional OutputSameFormat As Boolean = True) As Integer
 'Return Month Number by it's name
 'if OutputSameFormat be true output is same as input format that determine by IranDate
  Dim i As Integer, j As Integer, m As String
  i = Val(MonthStr)
  If IranDate Then
      If i = 0 Then
        m = LCase("Far Ord Kho Tir Mor Sha Meh Aba Aza Dey Bah Esf ")
        j = InStr(m, LCase(Left(MonthStr, 3)) + " ")
        i = (j - 1) \ 4 + 1
      End If
      If Not OutputSameFormat Then i = i + 3: If i > 12 Then i = i - 12
  Else
      If i = 0 Then
        m = LCase("Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec ")
        j = InStr(m, LCase(Left(MonthStr, 3)) + " ")
        i = (j - 1) \ 4 + 1
      End If
      If Not OutputSameFormat Then i = i + 9: If i > 12 Then i = i - 12
  End If
    
  MonthNum = i
End Function

Function VectorXYZ(dx, dy, dz) As Vector3DType
 'make a vector by dx,dy,dz
 Dim a As Vector3DType
 With a
  .dx = dx
  .dy = dy
  .dz = dz
 End With
 VectorXYZ = a
End Function

Public Sub SolveMatrix(a() As Single, n As Integer, x() As Single)
'Input Matrix A is (n) * (n+1)
'Output Answer Matrix x() is n*1
'A(i,n+1) is constants c1,c2,...(without go them to left side of eq.) in this form:
'a11 x1 + a12 x2 + ... + a1n xn = c1 ,-> A(1,n+1)=c1
'...

Dim i As Integer, j As Integer, i2 As Integer, j2 As Integer
Dim m As Integer
Dim Mai As Single, aii As Single, ai2i As Single, sigma As Single
Dim tmp As Single
Dim Time1 As Single

Time1 = Timer
DoEvents

For i = 1 To n - 1
 'make aii be the biggest in column i and below it:
 Mai = Abs(a(i, i)): m = i
 For i2 = i + 1 To n
  If Abs(a(i2, i)) > Mai Then Mai = Abs(a(i2, i)): m = i2
 Next
 If m <> i Then 'must pivot:  (swap row i and m)
  For j = 1 To n + 1
   tmp = a(i, j): a(i, j) = a(m, j): a(m, j) = tmp
  Next
 End If

 'Forward Elemination:
 aii = a(i, i)
 If aii = 0 Then
  MsgBox "Error!, Devision by Zero." + vbCrLf + "In 'SolveMatrix' Function" ': Stop: End
 End If
 For i2 = i + 1 To n      'make a(i2,i) to Zero by use a(i,i)
  ai2i = a(i2, i)
  If ai2i <> 0 Then
    a(i2, i) = 0
    For j = i + 1 To n + 1
     a(i2, j) = a(i2, j) - ai2i * a(i, j) / aii
    Next j
  End If
 Next i2
If i Mod 5 = 0 Then DoEvents
If i Mod 10 = 0 Then
 'b$ = FormatPercent(i / (n - 1), 2) '  Format(i / (n - 1), "0.00%")
 'If heatMess.WindowState = vbMinimized Then heatMess.Caption = "Analyse " + b$ Else heatMess.Caption = acap1$
 'heatMess.LabelComputed.Caption = b$
 'heatMess.LabelTime.Caption = Format(Timer - time1, "0") + " Sec"
 'If heatMess.Tag = "Cancel" Then
 ' heatMess.Caption = acap1$
 ' Exit Sub
 'End If
End If

Next i

'    ////////////////////////////////////////////////////////////////////
'    /                                i-1                               /
'    /                    a(n-i,n+1) - ä a(n-i,n-j).x(n-j)              /
'    /                                j=0                               /
'    / Formula:  x(n-i) = -------------------------------- , 0<=i<=n-1  /
'    /                                                                  /
'    /                              a(n-i,n-i)                          /
'    /                                                                  /
'    ////////////////////////////////////////////////////////////////////

'heatMess.Caption = acap1$

For i = 0 To n - 1
 sigma = 0
 For j = 0 To i - 1
  sigma = sigma + a(n - i, n - j) * x(n - j)
 Next
 If a(n - i, n - i) = 0 Then
  MsgBox "Error!, Find Exact Value Of x Matrix Is Impossible."
  'Stop
  'End
 End If
 x(n - i) = (a(n - i, n + 1) - sigma) / a(n - i, n - i)
Next i

End Sub

Function CheckPageIsInShadow(Vs As Vector3DType, ShadowCode As Integer) As Integer
 'UserIndexObj,UserIndexPage: Index of object and Page in the scene
 'Vs: Sun Vector
 'ShadowCode:   0=NoShadow
 '              1=Parent Object
 '              2=Other Objects
 '              3=1+2=Consider All Shadows
 'for some simplicity this sub use UserVariables instead of get from params
 
 If ShadowCode = 0 Then CheckPageIsInShadow = 0: Exit Function
 Dim iObj As Long, iPage As Long, page0 As Page3DType, p0 As Point3DType, Pi As Point3DType
 
 With UserScene
   page0 = .Object(UserIndexObj).Page(UserIndexPage)
   p0 = page0.Point(0)
   If ShadowCode And 2 = 0 Then iObj = UserIndexObj Else iObj = 0
   For iObj = iObj To .NumObjects - 1
    If iObj <> UserIndexObj Or (ShadowCode And 1) <> 0 Then
     With .Object(iObj)
      If .iCode And 8 Then 'if this object can cause shadow
        For iPage = 0 To .NumPages - 1
          If iObj <> UserIndexObj Or iPage <> UserIndexPage Then 'skip page itself
            If IfLinePassThroughPage(.Page(iPage), VnOfPage(.Page(iPage)), p0, Vs) Then
              CheckPageIsInShadow = 1: Exit Function
            End If
          End If
        Next
      End If
     End With
     If ShadowCode And 2 = 0 Then Exit For 'ShadowCode=1
    End If
   Next 'iobj
   
 End With
 
 CheckPageIsInShadow = 0
 
End Function

Function SigmaGtByHour(s1 As SolarType, Vn As Vector3DType, h1Deg As Single, h2Deg As Single, ByVal dtSec As Long, GDir As Single, Gdif As Single, Optional TraceAsCollector As Integer = 0, Optional s_IsInRad As Boolean = False)
'Return: Gtotal = GDir + Gdif
'dtSec : delta t in secound. time period to computing Sigma
'Vn : normal vector of plate
'h1,h2 : Integral Hour Angle limits, h of begin and end in 'Degree'
'** if h1=h2 then this sub compute G (w/m2=j/s.m2) instead of jule
'GDir: GDirect
'Gdif: Gdiffuse
'TraceAsCollector: if be true G will compute for the North-South axis collector!
's_IsInRad : this function use Rad mode to faster computation, so
'              if your angles (l,d,h,Beta,Fi) are not in Rad set this to true
'              for convert them automatically.

Dim s As SolarType
Dim l, d, sl, cl, sd, cd, a1, a2
Dim Beta, fi, SinB, dh

Dim G As Single, Gd As Single 'Direct & diffuse
Dim h As Single, h1 As Single, h2 As Single, i As Long
Dim Vs As Vector3DType, Vz As Vector3DType
Dim CosTeta, CF, Fws
s = s1
If Not s_IsInRad Then ConvertRadDegMode s, False
CF = s.CF

l = s.l: d = s.d
sl = Sin(l): cl = Cos(l): sd = Sin(d): cd = Cos(d)
a1 = cl * cd: a2 = sl * sd

If Not TraceAsCollector Then
 'for a fix plate:
 Fws = (1 + CosineV1V2(Vn, VectorXYZ(0, 0, 1))) / 2
End If

h1 = h1Deg / kRad2Deg: h2 = h2Deg / kRad2Deg
If h1 > h2 Then h = h1: h1 = h2: h2 = h 'swap h1,h2
dh = (Pi * dtSec / 43200) 'delta hour in Rad (24 hours = 360 deg = 2Pi Rad)

If h2 = h1 Then dh = 1 'to exit surely after one loop

G = 0: Gd = 0
i = 0

Do
 
 h = h1 + i * dh: If h > h2 Then Exit Do
 i = i + 1
 SinB = a1 * Cos(h) + a2
 Beta = ArcSin(SinB, True)
 fi = Sgn(h) * ArcCoS((SinB * sl - sd) / (Cos(Beta) * cl), True)
 
 If TraceAsCollector = 1 Then
  CosTeta = Sqr(Sin(h) ^ 2 * cd * cd + SinB * SinB)
  Fws = (1 + CosineV1V2(Vn, VectorXYZ(0, Sin(fi), Tan(Beta)))) / 2
 Else
  Vs = MakeSunVector(Beta, fi, True)
  CosTeta = CosineV1V2(Vs, Vn)
 End If
 G = G + GDirect(Beta, CosTeta, CF)
 Gd = Gd + Gdiffuse(Beta, CF, Fws)
 
Loop

If h1 <> h2 Then 'is computed for a point
 GDir = G * dtSec
 Gdif = Gd * dtSec
Else
 GDir = G
 Gdif = Gd
End If

SigmaGtByHour = GDir + Gdif

End Function


Function SigmaGtDaily(s As SolarType, Vn As Vector3DType, ByVal dtSec As Single, GDir As Single, Gdif As Single, Optional DayLengthSec As Single, Optional TraceAsCollector As Integer = 0, Optional EditAsShorterDay As Boolean = True) As Single
 'Return: Total(Dir+dif) Energy J/m2 for this day
 'GDir:GDirect / Gdif:Gdiffuse
 'DayLengthSec (output): length of day in secound
 
 'Input parameters:
 'Vn: normal vector of plate ex: Vn=(0,0,1) for Horizontal
 'dtSec: delta t in secound for computing
 
 'EditAsShorterDay: actually, we have
 'some thing such as mountains that
 'prevent from absorbtion of solar rad. so:
 '0.97 & 0.99 may multiply to obtain light better data for a Sigma average against of errors of shorter actual day, converting,...
  
 SetDayOfYear s.nDay, s
 Dim h1, h2, G
 s.Beta = 0
 h1 = FindHourByBeta(s) 'SunRise
 h2 = -h1 'SunSet
 G = SigmaGtByHour(s, Vn, h1, h2, dtSec, GDir, Gdif, TraceAsCollector)
 If EditAsShorterDay Then
  G = GDir * 0.97 + Gdif * 0.99
 End If
 DayLengthSec = 2 * h1 * 3600& / 15
 SigmaGtDaily = G
 
End Function

Public Sub SetDayOfYear(ByVal nDay As Integer, s As SolarType)
 'Get nDay (DayOfYear in Milady)
 'Set nDayOfYear,EOT,declination of s
 'EOT is in minutes -> LST(Local Solar Time)=LCT(Local Civil Time)+EOT(Equation Of Time)
 'Use nDayOfYear -> between 1 to 365  (1=1 january)
 Dim n As Single, d As Single
 s.nDay = nDay
 n = (nDay - 1) * 2! * Pi / 365
 
 s.EOT = 229.2 * (0.000075 + 0.001868 * Cos(n) - 0.032077 * Sin(n) - 0.014615 * Cos(2 * n) - 0.04089 * Sin(2 * n))
 
 If MethodOfSolarCalc = 3 Then
   d = 23.45 * Sin(2 * Pi * (284 + nDay) / 365)
 Else
   d = 0.3963723 - 22.9132745 * Cos(n) + 4.0254304 * Sin(n) _
 - 0.387205 * Cos(2 * n) + 0.05196728 * Sin(2 * n) - 0.1545267 * Cos(3 * n) _
 + 0.08479777 * Sin(3 * n)
 End If
 
 If s.RadMode Then s.d = d / kRad2Deg Else s.d = d
 
 If MethodOfSolarCalc = 2 Then
  SetABCforASHRAE nDay
 ElseIf MethodOfSolarCalc = 3 Then
  ConstSolar.Hottel_Gon = 1370 * (1 + 0.033 * Cos(2 * Pi / 365 * nDay))
 End If
 
 s.Month = MonthByDay(nDay)
 SetCityData s
 
End Sub

Public Sub SetDate(ByVal nDayOfMonth As Integer, ByVal MonthStr As String, s As SolarType, Optional ByVal IranDate As Boolean = False)
 'MonthStr: Month Name Or Number; such as Jan,Feb,.. or 1,2,..
 'IranDate: True means that date is given as 'hejri shamsi'
 Dim Ai As Variant, m As String, i As Long, n As Long
 i = MonthNum(MonthStr, IranDate)
 If IranDate Then
  Select Case i
  Case 1 To 6: n = (i - 1) * 31 + nDayOfMonth
  Case 7 To 12: n = 6 * 31 + (i - 7) * 30 + nDayOfMonth
  End Select
  n = n + 79
  If n > 365 Then n = n - 365
 Else
  Ai = Array(0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334)
  n = nDayOfMonth + Ai(i - 1)
 End If
 
 SetDayOfYear n, s
 
End Sub

Public Function DayByDate(ByVal nDayOfMonth As Integer, ByVal MonthStr As String, Optional ByVal IranDate As Boolean = False) As Long
 'MonthStr: Month Name Or Number; such as Jan,Feb,.. or 1,2,..
 'IranDate: True means that date is given as 'hejri shamsi'
 'Return: Milady Day Number (1 to 365)
 Dim Ai As Variant, m As String, i As Long, n As Long
 i = MonthNum(MonthStr, IranDate)
 If IranDate Then
  Select Case i
  Case 1 To 6: n = (i - 1) * 31 + nDayOfMonth
  Case 7 To 12: n = 6 * 31 + (i - 7) * 30 + nDayOfMonth
  End Select
  n = n + 79
  If n > 365 Then n = n - 365
 Else
  Ai = Array(0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334)
  n = nDayOfMonth + Ai(i - 1)
 End If
 
 DayByDate = n
 
End Function

Sub SetCityData(s As SolarType)
 's.Month must be as Milady
 'SET CF (Cloud Factor) And Latitude/Longtitude Of City
 'Longtitudes are not correct except for shiraz/tehran
 On Error Resume Next
 Dim i As Integer
 Dim c As Variant, CityStr As String
 CityStr = s.CityName
 
 If UseStatisticDataForCF = 0 Then s.CF = FixedCF
 
 If LCase(CityStr) <> LCase("ArbitraryLocation") Then '/////////// Set Properties (l,CF) by city name:
 
 If LTrim(CityStr) = "" Then Exit Sub
 'first 12 data are CF and 2 Last data are Latitude(y),Longtitude(x)
 Select Case UCase(Replace(CityStr, " ", ""))
 Case "RASHT": c = Array(0.707, 0.77, 0.79, 0.672, 0.514, 0.54, 0.5, 0.573, 0.605, 0.625, 0.649, 0.691, 37.31, 51.43, -7)
 Case "RAMSAR": c = Array(0.607, 0.651, 0.753, 0.667, 0.579, 0.545, 0.526, 0.614, 0.656, 0.678, 0.597, 0.641, 36.9, 51.43, -20)
 Case "BABOLSAR": c = Array(0.532, 0.587, 0.649, 0.591, 0.484, 0.436, 0.434, 0.483, 0.531, 0.542, 0.52, 0.562, 36.71, 51.43, -21)
 Case "KHOY", "KHOI": c = Array(0.633, 0.612, 0.548, 0.521, 0.494, 0.402, 0.333, 0.245, 0.303, 0.366, 0.519, 0.595, 38.13, 51.43, 1160)
 Case "TABRIZ": c = Array(0.586, 0.53, 0.528, 0.5, 0.379, 0.198, 0.192, 0.178, 0.183, 0.304, 0.392, 0.541, 38.13, 51.43, 1349)
 Case "REZAEYEH", "REZAEIEH": c = Array(0.645, 0.55, 0.563, 0.51, 0.389, 0.205, 0.176, 0.152, 0.201, 0.337, 0.431, 0.56, 37.53, 51.43, 1000)
 Case "SAGHEZ": c = Array(0.638, 0.563, 0.56, 0.532, 0.38, 0.261, 0.227, 0.201, 0.194, 0.221, 0.418, 0.605, 36.25, 51.43, 1476)
 Case "TEHRAN": c = Array(0.429, 0.391, 0.456, 0.4, 0.352, 0.209, 0.216, 0.186, 0.18, 0.242, 0.316, 0.423, 35.68, 51.43, 1191)
 Case "HAMEDAN": c = Array(0.575, 0.536, 0.501, 0.502, 0.387, 0.194, 0.214, 0.201, 0.226, 0.297, 0.419, 0.525, 35.2, 51.43, 1644)
 Case "KERMANSHAH", "BAKHTARAN": c = Array(0.574, 0.543, 0.486, 0.477, 0.359, 0.157, 0.168, 0.166, 0.181, 0.291, 0.412, 0.52, 34.31, 51.43, 1322)
 Case "KASHAN": c = Array(0.559, 0.453, 0.497, 0.482, 0.431, 0.43, 0.38, 0.292, 0.268, 0.37, 0.441, 0.465, 33.98, 51.43, 955)
 Case "KHORRAMABAD", "KHORAMABAD": c = Array(0.468, 0.431, 0.465, 0.478, 0.35, 0.184, 0.187, 0.193, 0.194, 0.212, 0.383, 0.488, 33.48, 51.43, 1134)
 Case "SHAHREKORD": c = Array(0.464, 0.393, 0.392, 0.441, 0.282, 0.208, 0.243, 0.191, 0.151, 0.173, 0.319, 0.356, 32.31, 51.43, 2066)
 Case "SHAHROUD", "SHAHROOD", "SHAHROD", "SHAHRUD": c = Array(0.407, 0.4, 0.422, 0.412, 0.353, 0.25, 0.233, 0.174, 0.189, 0.235, 0.284, 0.397, 36.41, 51.43, 1000)
 Case "MASHHAD", "MASHAD": c = Array(0.447, 0.502, 0.591, 0.499, 0.343, 0.188, 0.174, 0.142, 0.171, 0.265, 0.317, 0.441, 36.26, 51.43, 985)
 Case "SABZEVAR": c = Array(0.458, 0.41, 0.49, 0.433, 0.326, 0.252, 0.209, 0.17, 0.18, 0.221, 0.304, 0.428, 36.21, 51.43, 94)
 Case "SEMNAN": c = Array(0.386, 0.395, 0.453, 0.43, 0.358, 0.246, 0.226, 0.184, 0.186, 0.229, 0.277, 0.348, 35.55, 51.43, 1138)
 Case "TORBATHEIDARIEH", "TORBATEHEIDARIEH": c = Array(0.461, 0.489, 0.523, 0.397, 0.327, 0.203, 0.175, 0.132, 0.168, 0.222, 0.278, 0.474, 35.26, 51.43, 1000)
 Case "TABAS": c = Array(0.35, 0.361, 0.373, 0.386, 0.297, 0.225, 0.204, 0.186, 0.147, 0.141, 0.276, 0.377, 33.6, 51.43, 691)
 Case "BIRJAND": c = Array(0.372, 0.387, 0.398, 0.384, 0.294, 0.207, 0.18, 0.158, 0.127, 0.145, 0.276, 0.319, 32.86, 51.43, 1456)
 Case "ISFAHAN", "ISFEHAN": c = Array(0.304, 0.324, 0.299, 0.349, 0.32, 0.162, 0.206, 0.168, 0.149, 0.159, 0.277, 0.33, 32.61, 51.43, 1590)
 Case "YAZD": c = Array(0.419, 0.348, 0.379, 0.423, 0.338, 0.274, 0.241, 0.164, 0.141, 0.148, 0.272, 0.358, 31.9, 51.43, 1230)
 Case "AHWAZ", "AHVAZ": c = Array(0.458, 0.438, 0.465, 0.457, 0.379, 0.267, 0.266, 0.217, 0.177, 0.258, 0.357, 0.346, 31.33, 51.43, 18)
 Case "ZABOL": c = Array(0.397, 0.377, 0.395, 0.345, 0.273, 0.205, 0.219, 0.206, 0.178, 0.164, 0.244, 0.36, 31.03, 51.43, 487)
 Case "ABADAN": c = Array(0.418, 0.377, 0.36, 0.398, 0.347, 0.258, 0.262, 0.232, 0.185, 0.228, 0.386, 0.415, 30.36, 51.43, 11)
 Case "KERMAN": c = Array(0.391, 0.367, 0.359, 0.435, 0.297, 0.244, 0.235, 0.207, 0.192, 0.219, 0.226, 0.32, 30.25, 51.43, 1749)
 Case "SHIRAZ": c = Array(0.342, 0.307, 0.347, 0.381, 0.223, 0.147, 0.195, 0.184, 0.131, 0.13, 0.239, 0.306, 29.53, 52.53, 1491)
 Case "ZAHEDAN": c = Array(0.422, 0.366, 0.403, 0.409, 0.317, 0.264, 0.257, 0.201, 0.171, 0.203, 0.227, 0.313, 29.46, 51.43, 1370)
 Case "BAM": c = Array(0.292, 0.326, 0.34, 0.354, 0.248, 0.211, 0.222, 0.183, 0.172, 0.127, 0.163, 0.2, 29.1, 51.43, 1067)
 Case "BOSHEHR", "BUSHEHR": c = Array(0.295, 0.382, 0.337, 0.366, 0.26, 0.182, 0.216, 0.2, 0.168, 0.168, 0.253, 0.339, 28.98, 51.43, 4)
 Case "BANDARABBAS", "BANDARABAS": c = Array(0.325, 0.304, 0.353, 0.354, 0.212, 0.254, 0.357, 0.299, 0.251, 0.176, 0.195, 0.246, 27.21, 51.43, 10)
 Case "IRANSHAHR": c = Array(0.263, 0.298, 0.315, 0.319, 0.231, 0.312, 0.364, 0.266, 0.19, 0.122, 0.117, 0.183, 27.2, 51.43, 566)
 Case "CHABAHAR": c = Array(0.261, 0.254, 0.262, 0.249, 0.215, 0.314, 0.373, 0.382, 0.291, 0.165, 0.16, 0.217, 25.28, 51.43, 7)
 
 Case "NEWYORK":
 
 Case Else: MsgBox "Cant Find City!": End: Exit Sub 'stop: City Not Found
 End Select
 
 i = MonthNum(s.Month, False): If i <= 0 Then i = 1
 If IsEmpty(c) Then Exit Sub
 
 s.Altitude = c(14)
 If UseStatisticDataForCF Then s.CF = c(i - 1)
 If s.RadMode Then s.l = c(12) / kRad2Deg Else s.l = c(12)
 s.DifLong2Est = 0 'c(13) - 51.43
 
 End If '/////////////////////////////////// <>ArbitraryLocation
 
 If MethodOfSolarCalc = 3 Then
   SetabkForHottel s
 End If

End Sub


'------------------------------------------------------
'******************************* USER FUNCTIONS:
'------------------------------------------------------

Public Sub UserSetLocation(ByVal CityStr As String)
  sUser.CityName = CityStr
End Sub

Public Sub UserSetDate(ByVal nDayOfMonth As Integer, ByVal MonthStr As String, Optional ByVal IranDate As Boolean = False)
 'MonthStr: Month Name Or Number; such as Jan,Feb,.. or 1,2,..
 'IranDate: True means that date is given as 'hejri shamsi'
 SetDate nDayOfMonth, MonthStr, sUser, IranDate
End Sub

Sub UserFindRadiationScene(ByVal dMinutes As Single, Optional ByVal ComponentG As Integer = 7, Optional ShowProgress As Integer = True, Optional TraceAsCollector As Integer = 0, Optional ColorMode As Integer = 0, Optional OnlyOneObject As Long = -1)
 'ComponentG : determine what component of G must use for color of plates:
 '               1=GDir  2=Gdiff  4=GRef
 '
 'UserShadowCode:  determine how should consider shadow of other pages when calculating Radiation (this affect Direct Radiation)
 '              0=NoShadow
 '              1=Consider Shadow Of pages Of parent Object
 '              2=Consider Shadow Of Other Objects
 '              3=1+2=Consider All Shadows
 'ColorMode:
 '           0=Normal set color of pages
 '           1,2,4=use single color (R or G or B)  -> (ex: 3=R+G=yellow)
 '           8=gradient relative to solar noon
 '           16=gradient relative to a special Value
 '           -1=Don't set colors by G
 ' OnlyOneObject:
 '            -1=calculate all objects of the scene OTHER => only one object (this means UserFindRadiationObject)
 '
 '- Return Answers in SpValues() and Some Global variables
 '- this sub Compute Radiation in each object of
 '  the Userscene that has iCode 4 (or 4+...) and set it's color
 '
 'Note: this sub cause to define SunObject for half time of first day in UserFindRadiationPlate for drawing it as an average sun position
 
 Dim n As Long, nTot As Long, pTot As Long, iRadObj() As Long, iTot As Long, iObj As Long, i As Long, j As Long, Vn As Vector3DType, Ai As Single, AObj As Single
 Dim GDir As Single, Gdif As Single, GRef As Single, Gtot As Single, G As Single, SigmaGdA As Single
 Dim SigmaQDir As Single, SigmaQdif As Single, SigmaQRef As Single
 Dim GPage() As Single, GMin As Single, GMax As Single, GAve As Single, GMinTot As Single, GMaxTot As Single
 Dim uR As Integer, uG As Integer, uB As Integer, uRGB As Integer, c As Integer
 
 n = 0: pTot = 0: nTot = 0
 iTot = 0
 
 'first find objects that should calculate their Radiation
 With UserScene
 If .NumObjects <= 0 Then Exit Sub
 ReDim iRadObj(.NumObjects - 1)
 If OnlyOneObject > -1 Then
   iRadObj(nTot) = OnlyOneObject: nTot = nTot + 1: pTot = pTot + .Object(OnlyOneObject).NumPages
 Else
   For iObj = 0 To .NumObjects - 1
     With .Object(iObj)
      If .iCode And 4 Then
       iRadObj(nTot) = iObj
       nTot = nTot + 1: pTot = pTot + .NumPages
      End If
     End With
   Next
 End If
 End With
     
 'first keep radiation of each page and calculate Min,Max of them. then set col of pages.
 If ShowProgress Then
  Dim FProg As New FormProgress
  FProg.SetProperties "Calculate..."
  FProg.Show
  FProg.Refresh
 End If

 If pTot > 0 Then ReDim GPage(pTot - 1)
 iTot = -1
 
 With UserScene
  For iObj = 0 To nTot - 1
   UserIndexObj = iRadObj(iObj)
   With .Object(UserIndexObj)
    '***** Each Object:
    
    SigmaGdA = 0: AObj = 0
    SigmaQDir = 0: SigmaQdif = 0: SigmaQRef = 0
    'uR: UseRed or not ,uG,uB ...
    uR = Sgn(ColorMode And 1): uG = Sgn(ColorMode And 2): uB = Sgn(ColorMode And 4)
    uRGB = uR Or uG Or uB
    
    n = .NumPages - 1
    
    For i = 0 To n
      UserIndexPage = i
      iTot = iTot + 1
      Vn = VnOfPage(.Page(i))
      Ai = AreaPage(.Page(i))
      AObj = AObj + Ai
      
      If ShowProgress Then
        If iTot Mod 50 = 0 Or iTot >= pTot - 1 Then FProg.SetProgress iTot / pTot
      End If
      
      If CancelPressedPublic Then iObj = nTot - 1: Exit For 'to exit from both 'For-Next' loops
      
      'If i = 1 Then Stop
      UserFindRadiationPlate Vn, dMinutes, TraceAsCollector
      
      With GRadiationUser
       GDir = .Direct
       Gdif = .diffuse
       GRef = .Reflect
       Gtot = .Total
      End With
      G = 0
      If ComponentG And 1 Then G = G + GDir
      If ComponentG And 2 Then G = G + Gdif
      If ComponentG And 4 Then G = G + GRef
      GPage(iTot) = G: SigmaGdA = SigmaGdA + G * Ai
      SigmaQDir = SigmaQDir + GDir * Ai: SigmaQdif = SigmaQdif + Gdif * Ai: SigmaQRef = SigmaQRef + GRef * Ai
      If i = 0 Then
       GMin = G: GMax = G
      Else
       If G < GMin Then GMin = G
       If G > GMax Then GMax = G
      End If
    Next
    
    GAve = SigmaGdA / AObj 'Gave=Integral(G.dA)/Integral(dA)
    
    'set special values of this object
    ReDim Preserve .SpValues(10)
    .SpValues(0) = AObj
    .SpValues(1) = GMin 'these values are based on ComponentG (these are Max,Min,Ave of Direct or diffuse or... that set by ComponentG)
    .SpValues(2) = GMax
    .SpValues(3) = GAve
    
    'these values are absolute and don't attention to ComponentG
    .SpValues(4) = SigmaQDir
    .SpValues(5) = SigmaQdif
    .SpValues(6) = SigmaQRef
    'Total Time:
    .SpValues(7) = UserTimeTotal
    
   End With 'End With .Object(RadiObj(iobj))
   '********** End Each Object
   
   'check for Total GMin,Max (Min,Max Of Userscene)
   If iObj = 0 Then
     GMinTot = GMin: GMaxTot = GMax 'initiate GMinMaxTotal in first compare
   Else
     If GMin < GMinTot Then GMinTot = GMin
     If GMax > GMaxTot Then GMaxTot = GMax
   End If
    
   
  Next 'next iobj 'next object
  
  If ShowProgress Then Unload FProg
  
  If UseAbsoluteGMinMax Then 'replace with specified values
    GMinTot = AbsoluteGMin
    GMaxTot = AbsoluteGMax
  End If
  
  'set GMinTot,GMaxTot for the Userscene:
  ReDim .SpValues(1)
  .SpValues(0) = GMinTot
  .SpValues(1) = GMaxTot
      
  If ColorMode <> -1 Then
   'set color of pages:
   iTot = -1
   For iObj = 0 To nTot - 1
    
    With .Object(iRadObj(iObj))
     n = .NumPages - 1
     For i = 0 To n
      iTot = iTot + 1
      If uRGB Then
       c = InterpolateCol(20, 255, GMinTot, GMaxTot, GPage(iTot))
       .Page(i).ColRGB = RGB(c * uR, c * uG, c * uB)
      Else
       .Page(i).ColRGB = ColGrad(InterpolateCol(0, ColGradMaxI, GMinTot, GMaxTot, GPage(iTot)))
      End If
     Next
     
    End With '.Object
    
   Next 'iObj
  End If 'ColorMode<>-1
  
  
 End With 'End With Userscene

End Sub

Sub UserFindRadiationObject(ObjIndex As Long, ByVal dMinutes As Single, Optional ByVal ComponentG As Integer = 7, Optional ShowProgress As Integer = True, Optional TraceAsCollector As Integer = 0, Optional ColorMode As Integer = 0)
 
 'see UserFindRadiationScene for params

 UserFindRadiationScene dMinutes, ComponentG, ShowProgress, TraceAsCollector, ColorMode, ObjIndex
 
End Sub


Sub UserFindRadiationPlate(Vn As Vector3DType, Optional ByVal dMinutes As Single = 15, Optional TraceAsCollector As Integer = 0)
'Use sUser,nDayBeginEnd(),nTimeBeginEnd()
'dMinutes: delta time in minutes to integrating
'Return:
'       G is returned at GRadiationUser
'       UserTimeTotal (reset by this sub) -> total time in minutes that calculate for.

'Note: in this sub SunObject will define automatically for half time of first day

Dim s As SolarType
Dim LSTBeginEnd(1) As Single
Dim l, d, sl, cl, sd, cd, a1, a2, k
Dim Beta, fi, SinB
Dim nDay As Integer
Dim CosTeta, CF, Fws, Fwg, Shad As Integer
Dim nTime As Single
Dim DayCount As Long, TimeCount As Long
Dim st As String, LSTExact(2) As Single, MustLoopYear As Integer

Dim G As Single, Gd As Single, Gr As Single  'Direct , diffuse & Reflect For Times Of Each Day
Dim SigmaG As Single, SigmaGd As Single, SigmaGr As Single
Dim DailyG As Single, DailyGd As Single, DailyGr As Single  'Direct , diffuse & Reflect For Days

Dim TimeTot As Single, t As Single, CalRate As Integer

Dim h As Single, h1 As Single, h2 As Single, i As Long
Dim Vs As Vector3DType, Vz As Vector3DType
s = sUser
ConvertRadDegMode s, False

l = s.l: sl = Sin(l): cl = Cos(l)
Vz = Vector3D(0, 0, 1)

If Not TraceAsCollector Then
 'for a fix plate:
 k = CosineV1V2(Vn, VectorXYZ(0, 0, 1))
 Fws = (1 + k) / 2
 Fwg = (1 - k) / 2
End If

SigmaG = 0: SigmaGd = 0: SigmaGr = 0
DayCount = 0: TimeCount = 0
TimeTot = 0

If UserDayStep = 0 Then UserDayStep = 1

MustLoopYear = nDayBeginEnd(1) < nDayBeginEnd(0) 'in this case we should go to end of year and continue from first day until reach to nDayBeginEnd(1)
nDay = nDayBeginEnd(0)

'Loop for count days:
Do ' Milady Days

 SetDayOfYear nDay, s
 d = s.d: sd = Sin(d): cd = Cos(d)
 a1 = cl * cd: a2 = sl * sd
 CF = s.CF
 
 'Set LSTBeginEnd() using nTimeBeginEnd():
 For i = 0 To 1
  st = UCase(nTimeBeginEnd(i)) 'first should translate SpecialKeyWords(SR,SS,SN)
  If InStr(st, "S") > 0 Then
    FindSunRiseSunSetNoon s, "1", "1", "1", "1", LSTExact()
    st = Replace(st, "SR", CStr(LSTExact(0)))
    st = Replace(st, "SS", CStr(LSTExact(1)))
    st = Replace(st, "SN", CStr(LSTExact(2)))
    LSTBeginEnd(i) = Val(st)
  Else
    'nTimeBeginEnd(i) is a text such as 17:45:30
    If Not TimeIsLST Then 'should convert to LST
      s.LCTstring = st: SetLSTasLCTstr s: LSTBeginEnd(i) = s.LST
    Else
      LSTBeginEnd(i) = TimeToMinutes(st)
    End If
  End If
 Next
  
   
 If DayCount = 0 Then
  'to graphical show: Find Sun in half time for first day (this must be corrected to mid day but now i use it for simplicity)
  nTime = (LSTBeginEnd(0) + LSTBeginEnd(1)) / 2 'near to Average Time
  h = Pi * (1 - nTime / 720)
  SinB = a1 * Cos(h) + a2: Beta = ArcSin(SinB, True)
  fi = Sgn(h) * ArcCoS((SinB * sl - sd) / (Cos(Beta) * cl), True)
  Vs = MakeSunVector(Beta, fi, True)
  If SunDistance <= 0 Then SunDistance = 180
  SetLengthVector Vs, SunDistance
  SunObject = MakeSunObject(ConvertVector2Point(Vs), 4, &HEEFF&, &HF5FF&)
 End If
 
 G = 0: Gd = 0: Gr = 0
 
 If LSTBeginEnd(1) = LSTBeginEnd(0) Or dMinutes = 0 Then CalRate = 1 Else CalRate = 0
 
 'Loop for count times: ////////////
 For nTime = LSTBeginEnd(0) To LSTBeginEnd(1) Step dMinutes  'Minutes
  
  'TimeCount = TimeCount + 1 'number of summations in each day
  
  If CalRate = 0 Then
   't = time period for this step (delta t)
   If nTime = LSTBeginEnd(0) Then 'is first (begin time for this day)
    t = dMinutes / 2
   Else
    t = dMinutes
   End If
   If nTime + dMinutes / 2 > LSTBeginEnd(1) Then
    t = nTime - dMinutes / 2
    If t < LSTBeginEnd(0) Then t = nTime - LSTBeginEnd(0)
    t = LSTBeginEnd(1) - t
   End If
  Else 'CalRate<>0
   t = 1
  End If 'CalRate=0
  
  h = Pi * (1 - nTime / 720) '(hour angle) Deg: 180 - nTime / 4   ' or: (12 - LST/60)*15
  SinB = a1 * Cos(h) + a2: If SinB < 0 Then SinB = 0 'sometimes is very small negative value for SunRise or SunSet
  Beta = ArcSin(SinB, True)
  fi = Sgn(h) * ArcCoS((SinB * sl - sd) / (Cos(Beta) * cl), True)
  
  'check for use specified radiation for a horizontal page:
  If UseSpecificHorizG Then 'find special value for C.F. by given GHoriz.
    'replacing CosTeta by SinBeta (because this is for a horizontal plate) and Fws by 1
   With ConstSolar
    Select Case MethodOfSolarCalc
    Case 0, 1
     k = .GDirectC * (1 - Exp(-4.2972 * Beta)) * SinB
     CF = (SpecificHorizG - .GdiffuseA1 - .GdiffuseA2 * Beta - k) / (.GdiffuseA3 - k)
    Case 2
     GNormalDirect = .ASHRAE_A / Exp(.ASHRAE_B / Sin(Beta))
     CF = (SpecificHorizG / GNormalDirect - SinB - 1.1 * .ASHRAE_C) / (1.1 * .ASHRAE_C - SinB)
    Case 3
 
    End Select
   End With
  End If
  
  If TraceAsCollector = 0 Then
   Vs = MakeSunVector(Beta, fi, True) 'vector from origin to sun (or from any point to the sun)
   CosTeta = CosineV1V2(Vs, Vn) 'negative values means an angle bigger than 90 deg so it can't be seen from the sun!. GDirect function return zero automatically for such a this case.
  ElseIf TraceAsCollector = 1 Then 'NS-Collector
   CosTeta = Sqr(Sin(h) ^ 2 * cd * cd + SinB * SinB)
   k = CosineV1V2(Vector3D(0, 0, 1), Vector3D(0, Sin(fi), Tan(Beta)))  'Vn=Vector3D(0, Sin(Fi), Tan(Beta))
   Fws = (1 + k) / 2
   Fwg = (1 - k) / 2
  ElseIf TraceAsCollector = 2 Then 'EW-Collector
   
  ElseIf TraceAsCollector = 3 Then 'Full-Trace
   CosTeta = 1
   'Vs is Vn in this case of full trace
   k = CosineV1V2(Vz, MakeSunVector(Beta, fi, True))
   'k = CosineV1V2(Vector3D(0, 0, 1), Vector3D(0, Sin(fi), Tan(Beta)))  'Vn=Vector3D(0, Sin(Fi), Tan(Beta))
   Fws = (1 + k) / 2
   Fwg = (1 - k) / 2
  End If
  

  'check for shadow calculation
  'If UserIndexObj = 2 Then UserShadowCode = 2 Else UserShadowCode = 0
  If CosTeta > 0 Then
   If UserShadowCode Then Shad = CheckPageIsInShadow(Vs, UserShadowCode) Else Shad = 0
   If Shad = 0 Then 'if is not in shadow calculate Direct term
    G = G + GDirect(Beta, CosTeta, CF) * t
   End If
  End If 'CosTeta>
    
  
  Gd = Gd + Gdiffuse(Beta, CF, Fws, GNormalDirect, Towbeam) * t
  Gr = Gr + GReflectGround(Beta, SinB, CF, RoGround, Fwg) * t

  If CalRate Then Exit For 'must compute Rate (in an instant)
  
  TimeCount = TimeCount + 1
  If UserDoEventsOnPlate Then If TimeCount Mod 2000 = 0 Or nTime + dMinutes >= LSTBeginEnd(1) Then TimeCount = 0: DoEvents 'should remark//////
  
 Next 'nTime
 
 t = LSTBeginEnd(1) - LSTBeginEnd(0)
 If CalRate Then t = 1 Else TimeTot = TimeTot + t
 'find Gaverage for the day
 SigmaG = SigmaG + G     'G is integral(G.dt)
 SigmaGd = SigmaGd + Gd  'Gd " " "      Gd.dt
 SigmaGr = SigmaGr + Gr  'Gr " " "      Gr.dt
 
  
 DayCount = DayCount + 1 'should increse one for each loop, regardless of step of days!! (because we divide summation by this to get average)
 
 If nDay = nDayBeginEnd(1) Then Exit Do
 nDay = nDay + UserDayStep: If nDay > 365 Then If MustLoopYear Then nDay = 1: MustLoopYear = 0
 If nDay > nDayBeginEnd(1) And MustLoopYear = 0 Then nDay = nDayBeginEnd(1)
Loop 'nDay
 

UserTimeTotal = TimeTot
UserDayTotal = DayCount

'return answer:(average of (average of each day))
If TimeTot = 0 Then TimeTot = DayCount 'TimeTot=0 => average of some day instants ELSE average of total time
With GRadiationUser
 .Direct = SigmaG / TimeTot
 .diffuse = SigmaGd / TimeTot
 .Reflect = SigmaGr / TimeTot
 .Total = (SigmaG + SigmaGd + SigmaGr) / TimeTot
 'If .Total < 864 Then Stop
End With


End Sub


Public Sub DoEnd()
    
    On Error Resume Next
    
    SaveSetting App.EXEName, "Settings", "RecPath", RecordFolder
    
    Unload FormGr
    End

End Sub


