Attribute VB_Name = "ModuleGDI"
'              Graphics Using Win-API
'
'              By S.Serpooshan, 2001
'
'***************************************************
'SAMPLE OF USE:
'
'    InitializeGDI Pic1
'
'    Dim p(0 To 3) As POINTAPI
'    p(0).x = 0: p(0).y = 0
'    ...
'    DeleteObject Selectobject(gdi_BufferHdc, QBPen(0,0))
'    'SelectQbPenBrush 0, 14
'    ClearGDIwindow
'    DrawPolygon p(0), 4
'    Selectobject gdi_BufferHdc, QBBrush(16) 'no fill
'    DrawPolygon p(0), 4
'    RefreshGDIwindow
'
'    ReleaseGDI
'**************************************************

Option Explicit

DefSng A-Z

'these types are defined again in Graphix3D, so remark them here.
'Public Type Point3DType
'        x As Single 'single
'        y As Single 'single
'        z As Single 'single
'End Type
'
'Public Type Page3DType
'        ColRGB As Long 'RGB Color : -1=no specify color => don't change brush (use current brush, so if current be a Null Brush don't fill page)
'        NumPoints As Long
'        Point() As Point3DType 'Point(0)=Center Of Page , Point(1 to NumPoints)=Points
'        SpValue As Single 'any arbitrary value such as magnitude of Radiation on this page
'End Type
'
'Public Type Object3DType
'        CenterPoint As Point3DType
'        NameStr As String
'        iCode As Integer '1=Sphere / 2=SpObject-Tower / 4=Find Radiation / 8=Can Make Shadow
'                         '16=semiSphere / 32=Dome / 64=SunObject(=>Radius=SpValues(0))
'        ColEdge As Long  'color for Edges (-1 = NoEdge)
'        SpValues() As Single 'array for Special Values such as Temperature or G_Radiation
'        NumPages As Long
'        Page() As Page3DType
'End Type
'
'Public Type ViewParamsType
'        Teta As Single
'        fi As Single
'        View3DScale As Single
'        ViewXshift As Long '2D shift
'        ViewYshift As Long '2D shift
'        View3DShiftOrig As Point3DType '3D shift (set Target)
'        RadMode As Integer
'End Type
'
'Public Type View3DAnglesType
'        SinTeta As Single 'single
'        CosTeta As Single 'single
'        SinFi As Single 'single
'        CosFi As Single 'single
'End Type
'
'Public Type Scene3DType
'        NameStr As String
'        DefaultView As ViewParamsType
'        CenterPoint As Point3DType
'        NumObjects As Long
'        Object() As Object3DType
'        SpValues() As Single 'array for Special Values
'End Type
'
'Public Type SinCos3AngType
'        SinX As Single
'        CosX As Single
'        SinY As Single
'        CosY As Single
'        SinZ As Single
'        CosZ As Single
'End Type
'
'Public Type Vector3DType
' dx As Single 'Long
' dy As Single 'Long
' dz As Single 'Long
'End Type
'
'Public Type Line3DType
'        p1 As Point3DType
'        p2 As Point3DType
'End Type


'-- Constants for drawing functions
Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
Public Const SRCERASE = &H440328
Public Const SRCINVERT = &H660046
Public Const SRCPAINT = &HEE0086
Public Const BLACKNESS = &H42&
Public Const WHITENESS = &HFF0062

'-- Constants for Pens and Brush functions
Public Const PS_SOLID = 0
Public Const PS_DASH = 1
Public Const PS_DOT = 2
Public Const PS_DASHDOT = 3
Public Const PS_DASHDOTDOT = 4
Public Const PS_NULL = 5
Public Const PS_INSIDEFRAME = 6
Public Const NULL_BRUSH = 5 'used by GetStockBrush

'my 3D constants:
Public Const CS_PolyGoTo = 987987987 'goto next point
Public Const CS_PolyRel = 999987987 'reletive mode
Public Const CS_PolyAbs = 888987987 'absolute mode

'-- Types for drawing functions - API (2D)
Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

'-- Types for drawing functions - User Defined (3D)

Public Type PolyDraw3DType
        CommandStr As String
        Point() As Point3DType
End Type



'-- Device Context functions
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'-- Drawing functions
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long

'-- Pens and Brushs
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long

'Other functions
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'Variables used for Device Context and Drawing functions
Public gdi_BufferHdc As Long      'To store the HDC (Handle to Device Context) to a Memory Device Context to be used as a gdi_BufferHdc
Dim gdi_bkBitmap As Long

Public gdi_Hwnd As Long, gdi_HdcOut As Long
Public gdi_PicObj As PictureBox
Public gdi_dx As Long, gdi_dy As Long
Public gdi_Initiated As Integer

'Variables to define View
Public View3DAngles As View3DAnglesType
Public View3DTeta As Single, View3DFi As Single 'Radian

Public View3DShiftOrig As Point3DType 'Shift Origin
Public ViewXshift As Long, ViewYshift As Long
Public View3DScale As Single

Private Const Pi = 3.14159265
Private Const Pi2 = 1.5707963 'Pi/2
Private Const kRad2Deg = 57.2957795130823 '180/pi

'Other variables
Public FleshAxis(2) As Page3DType
Public TextXdraw As PolyDraw3DType, TextYdraw As PolyDraw3DType
Public SunObject As Object3DType, SunDistance As Single
Public DoRotateZY As Boolean 'if be true use other order for rotation instead xyz->use zy only such as teta and fi)

'-- Variables to store Pen and Brush handles
Public QBPen() As Long
Public QBBrush() As Long
Public gdi_BackgrFillCol As Long
Public ColGrad() As Long, ColGradMaxI As Integer '0 to ColGradMaxI
Public DontChangeCurPenBrush As Integer 'if be true then DrawPage3D don't change PenBrush to fill pages (higher speed - Wire frame)
Public DontSortPages As Integer 'if be true then don't sort pages of objects in DrawObject3D sub
'To Have a Wire-Frame Model Use:
'       DontChangeCurPenBrush = True : DontSortPages = True
'       SelectQbPenBrush 15, 16 '16=Null Brush
'       DrawScene3D UserScene
Private tmpLastcPen As Long, tmpLastcBrush As Long, tmpLastcPenDeleted As Long, tmpLastcBrushDeleted As Long

'Variable(s) Use For QSort sub
Const CS_upNumQSort = 2000
Private finishP(CS_upNumQSort) As Long, lowP(CS_upNumQSort) As Long, cline(CS_upNumQSort) As Byte


Sub InitializeGDI(PictureboxObj As PictureBox)
    Set gdi_PicObj = PictureboxObj
    With PictureboxObj
    'gdi_dx,gdi_dy: width and height of window
    If .AutoRedraw = False Then .AutoRedraw = True 'must be true
    gdi_Hwnd = .hwnd
    gdi_dx = .ScaleX(.ScaleWidth, .ScaleMode, vbPixels)
    gdi_dy = .ScaleY(.ScaleHeight, .ScaleMode, vbPixels)
    'ScreenDC = GetDC(0) 'Hwnd=0 is On top of all windows
    'If hdc = 0 Then gdi_HdcOut = GetDC(gdi_Hwnd) Else gdi_HdcOut = hdc
    gdi_HdcOut = .hdc
    
    'set an array for gradient color (Blue>Cyan>Yellow>Red):
    'SetColGradArray
    'set x,y,z axis variables flesh and x,y texts:
    SetAxisParams
    'Creating Memory gdi_BufferHdc compatible to the current Screen mode
    gdi_BufferHdc = CreateCompatibleDC(gdi_HdcOut)
    gdi_bkBitmap = CreateCompatibleBitmap(gdi_HdcOut, gdi_dx, gdi_dy)
    DeleteObject SelectObject(gdi_BufferHdc, gdi_bkBitmap)
     
     
    Call CreateQBPensBrushs
    gdi_Initiated = True
    
    End With
End Sub

Sub ReleaseGDI()
    If Not gdi_Initiated Then Exit Sub
    On Error Resume Next
    ReleaseDC gdi_Hwnd, gdi_HdcOut
    DeleteObject gdi_bkBitmap
    DeleteDC gdi_BufferHdc
    DeleteQbPensBrushs
    gdi_Initiated = 0
End Sub

Sub CreateQBPensBrushs()
    '0, 1..16 -> SOLID Pen
    '1, ...   -> DASH  Pen
    '2, ...   -> DOT   Pen
    '3, ...   -> DASH_DOT
    
    '0-15=QB-Cols
    '16=Null
    
    Dim i As Long
    ReDim QBPen(3, 16), QBBrush(16)
    For i = 0 To 15
     QBPen(0, i) = CreatePen(PS_SOLID, 1, QBColor(i))
     QBPen(1, i) = CreatePen(PS_DASH, 1, QBColor(i))
     QBPen(2, i) = CreatePen(PS_DOT, 1, QBColor(i))
     QBPen(3, i) = CreatePen(PS_DASHDOT, 1, QBColor(i))
     
     QBBrush(i) = CreateSolidBrush(QBColor(i))
    Next
    For i = 0 To 3
     QBPen(i, 16) = CreatePen(PS_NULL, 0, 0)
    Next
    QBBrush(16) = GetStockObject(NULL_BRUSH)
    
End Sub

Sub DeleteQbPensBrushs()
    Dim i As Integer
    For i = 0 To UBound(QBBrush)
     DeleteObject QBPen(i)
     DeleteObject QBBrush(i)
    Next
End Sub

Sub RefreshGDIwindow()
        'Dim a As Long
        'Debug.Print gdi_HdcOut, FormGr.Pic1.hdc
        BitBlt gdi_HdcOut, 0, 0, gdi_dx, gdi_dy, gdi_BufferHdc, 0, 0, SRCCOPY
        'a = FormGr.Pic1.Image.Width
        gdi_PicObj.Refresh 'new
End Sub

Sub ClearGDIwindow()
        BitBlt gdi_BufferHdc, 0, 0, gdi_dx, gdi_dy, gdi_BufferHdc, 0, 0, WHITENESS  'Clear window
        'you can use Rectangle to fill background by any color
End Sub

Sub SelectQbPenBrush(Optional ByVal iPen As Long = -1, Optional ByVal iBrush As Long = -1, Optional iPenStyle As Long = 0)
            '-1 => don't change
            'iPen:0-16(QBCols)
            'iBrush:0-16(QBCols + 16=NoFill)
            'iPenStyle:
            '0=Solid , 1=Dash , 2=Dot , 3=Dash_Dot
            If iPen >= 0 Then DeleteObject SelectObject(gdi_BufferHdc, QBPen(iPenStyle, iPen))
            If iBrush >= 0 Then
                DeleteObject SelectObject(gdi_BufferHdc, QBBrush(iBrush))
            End If
End Sub

Sub SelectPenBrush(Optional ByVal ColPen As Long = -1, Optional ByVal ColBrush As Long = -1, Optional PenStyle As Long = 0)
    'iPen,iBrush < 0 => don't change
    If tmpLastcPen <> tmpLastcPenDeleted Or tmpLastcBrush <> tmpLastcBrushDeleted Then Call DeleteLastPenBrush 'programmer fergot this. so we do ourselves! ;-)
    If ColPen >= 0 Then
     tmpLastcPen = CreatePen(PenStyle, 1, ColPen)
     DeleteObject SelectObject(gdi_BufferHdc, tmpLastcPen)
    Else
     tmpLastcPen = -10
    End If
    If ColBrush >= 0 Then
     tmpLastcBrush = CreateSolidBrush(ColBrush)
     DeleteObject SelectObject(gdi_BufferHdc, tmpLastcBrush)
    Else
     tmpLastcBrush = -10
    End If
    'Note: must delete them after drawing! using DeleteLastPenBrush sub
End Sub

Sub DeleteLastPenBrush()
    'use to delete pen,Brush objects defined by SelectPenBrush
    If tmpLastcPen <> -10 Then tmpLastcPenDeleted = tmpLastcPen: DeleteObject tmpLastcPen: tmpLastcPen = -10
    If tmpLastcBrush <> -10 Then tmpLastcBrushDeleted = tmpLastcBrush: DeleteObject tmpLastcBrush: tmpLastcBrush = -10
End Sub

Sub GetRGBofCol(Col As Long, rr As Integer, gg As Integer, bb As Integer)
    'return red,green,blue of a color
    Dim c As Long
    c = Col
    rr = c Mod 256: c = c \ 256
    gg = c Mod 256
    bb = c \ 256
End Sub

Function GetRealNearestColor(ByVal hdc1 As Long, ByVal Col As Long) As Long
 Dim c As Long
 'NOTE!: to get real color (using 'getnearestcolor' API function may have some problems! and is not same as this sub
 c = GetPixel(hdc1, 1, 1)
 If c <> -1 Then
  GetRealNearestColor = SetPixel(hdc1, 1, 1, Col)
  SetPixelV hdc1, 1, 1, c
 Else
  GetRealNearestColor = Col 'faild to test
 End If

End Function


Sub FillHwndOut()
 'this sub fill background by the current color
 'SelectQbPenBrush iPen, iPen
 If gdi_BackgrFillCol < 0 Then gdi_BackgrFillCol = &HFFFFFF
 SelectPenBrush gdi_BackgrFillCol, gdi_BackgrFillCol
 Rectangle gdi_BufferHdc, 0, 0, gdi_dx, gdi_dy
 Call DeleteLastPenBrush
End Sub

Sub DrawPolygon(points As POINTAPI, iCount As Long)
    'iCount:4 for cube
    Polygon gdi_BufferHdc, points, iCount
End Sub


'*************************************************************
'  3D Graphics Simulation Functions:
'*************************************************************

Public Function Point3D(ByVal x, ByVal y, ByVal z) As Point3DType
 With Point3D
  .x = x
  .y = y
  .z = z
 End With
End Function

Public Function Vector3D(ByVal dx As Single, ByVal dy As Single, ByVal dz As Single) As Vector3DType
 'make a vector by dx,dy,dz
 With Vector3D
  .dx = dx
  .dy = dy
  .dz = dz
 End With
End Function

Sub SetViewShift3D(POrig As Point3DType)
'Cause to change Target (rotate about another point not origin)
 View3DShiftOrig = POrig
End Sub

Sub SetViewShift2D(ByVal XShift As Long, ByVal YShift As Long, Optional AddToShift As Boolean = False)
 'AddToShift: Determine if should add to current values
 If AddToShift Then
  ViewXshift = ViewXshift + XShift
  ViewYshift = ViewYshift + YShift
 Else
  ViewXshift = XShift
  ViewYshift = YShift
 End If
End Sub

Sub SetViewScale3D(ByVal VScale As Single)
 View3DScale = VScale
End Sub

Sub SetView3D(ByVal Teta, ByVal fi, Optional ByVal IsInRadMod As Boolean = False)
 'This Sub Set Viewer Location
 '
 'if IsInRad=True => treat Angles in Rad
 '**  when viewer rotate Teta about z objects rotate
 '    (-Teta) about it. so we change Teta,Fi to -Teta,-Fi
 
 If Not IsInRadMod Then
  Teta = Teta / kRad2Deg: fi = fi / kRad2Deg
 End If
 
 View3DTeta = Teta: View3DFi = fi
 
 With View3DAngles
  .SinTeta = -Sin(Teta)
  .CosTeta = Cos(Teta)
  .SinFi = -Sin(fi)
  .CosFi = Cos(fi)
 End With
 
End Sub

Sub SetViewWholeParams(ViewParams As ViewParamsType)
 'set all view parameters
 With ViewParams
  SetViewShift2D .ViewXshift, .ViewYshift
  SetViewShift3D .View3DShiftOrig
  SetViewScale3D .View3DScale
  SetView3D .Teta, .fi, .RadMode
 End With
End Sub

Function ViewParams(Teta1 As Single, Fi1 As Single, View3DScale1 As Single, ViewXshift1 As Long, ViewYshift1 As Long, View3DShiftOrig1 As Point3DType) As ViewParamsType
 With ViewParams
  .Teta = Teta1
  .fi = Fi1
  .View3DScale = View3DScale1
  .ViewXshift = ViewXshift1
  .ViewYshift = ViewYshift1
  .View3DShiftOrig = View3DShiftOrig1
 End With
End Function

Function SinCos3Angles(ByVal ax, ByVal ay, ByVal Az, Optional IsInRad As Boolean = False) As SinCos3AngType
 If Not IsInRad Then
  ax = ax / kRad2Deg
  ay = ay / kRad2Deg
  Az = Az / kRad2Deg
 End If
 With SinCos3Angles
  .SinX = Sin(ax)
  .CosX = Cos(ax)
  .SinY = Sin(ay)
  .CosY = Cos(ay)
  .SinZ = Sin(Az)
  .CosZ = Cos(Az)
 End With
End Function

Function Point2Dof3D(p1 As Point3DType) As POINTAPI
'Call 'SetView3D' sub to set viewer location

Dim xx As Single, x2 As Long, y2 As Long
Dim x As Single, y As Single, z As Single
If View3DScale = 0 Then View3DScale = 1

With p1
 x = (.x - View3DShiftOrig.x) * View3DScale: y = (.y - View3DShiftOrig.y) * View3DScale: z = (.z - View3DShiftOrig.z) * View3DScale
End With

With View3DAngles
 xx = x * .CosTeta - y * .SinTeta
 x2 = z * .SinFi + xx * .CosFi
 y2 = x * .SinTeta + y * .CosTeta
 'z2 = z * .CosFi - xx * .SinFi ' DON'T NEED IT HERE
End With

With Point2Dof3D
 .x = ViewXshift + y2
 .y = ViewYshift + x2
End With

End Function

Function ZHeight3D(p1 As Point3DType) As Single
 'Return z (Height) of a 3d point as seem from view point
 'when change view Teta,Fi this value change as objects and points rotate in reverse angles
 'This is Important to find Drawing order (from far objects to near objects)
 With View3DAngles
  ZHeight3D = (p1.z - View3DShiftOrig.z) * .CosFi - ((p1.x - View3DShiftOrig.x) * .CosTeta - (p1.y - View3DShiftOrig.y) * .SinTeta) * .SinFi
 End With
 
End Function

Function TransferPoint3D(p1 As Point3DType, v As Vector3DType) As Point3DType
 With TransferPoint3D
  .x = p1.x + v.dx
  .y = p1.y + v.dy
  .z = p1.z + v.dz
 End With
End Function

Sub TransferSelfPoint3D(p1 As Point3DType, v As Vector3DType)
 With p1
  .x = .x + v.dx
  .y = .y + v.dy
  .z = .z + v.dz
 End With
End Sub


Function TransferPage3D(page1 As Page3DType, v As Vector3DType) As Page3DType
 Dim n As Long, i As Long
 With TransferPage3D
  n = page1.NumPoints
  .ColRGB = page1.ColRGB
  .NumPoints = n
  ReDim .Point(n)
  For i = 0 To n
   .Point(i) = TransferPoint3D(page1.Point(i), v)
  Next
 End With
End Function

Sub TransferSelfPage3D(page1 As Page3DType, v As Vector3DType)
 Dim n As Long, i As Long
 With page1
  For i = 0 To page1.NumPoints
   TransferSelfPoint3D .Point(i), v
  Next
 End With
End Sub

Sub TransferSelfObject3D(Obj1 As Object3DType, v As Vector3DType)
 Dim i As Long
 With Obj1
  .CenterPoint = TransferPoint3D(.CenterPoint, v)
  For i = 0 To .NumPages - 1
   TransferSelfPage3D .Page(i), v
  Next
 End With
End Sub

Function TransferObject3D(Obj1 As Object3DType, v As Vector3DType) As Object3DType
 TransferObject3D = Obj1
 TransferSelfObject3D TransferObject3D, v
End Function

Function RotatePoint3D(p1 As Point3DType, SinCos3A As SinCos3AngType, Optional SelfMode As Boolean = False) As Point3DType
 'SelfMode=True => return none, p1 rotate itself (answer return in p1 already)
 'this sub USE THIS ORDER: Ax then Ay then Az
 'IT IS DIFFERENT THAT ROTATE A POINT OR OBJECT (FISRT ABOUT x AXIS THEN ABOUT y,Z) OR IN OTHER ORDER!
 'public var 'DoRotateZY' -> if true=> rotate about z and then y. (no rotation about x)
 Dim xx As Single, zz As Single
 Dim x As Single, y As Single, z As Single
 Dim x2 As Single, y2 As Single, z2 As Single
 
 With p1
  x = .x:  y = .y:  z = .z
 End With
 With SinCos3A
  If DoRotateZY = False Then 'default
   x2 = x
   y2 = y * .CosX - z * .SinX
   z2 = y * .SinX + z * .CosX
   x2 = z2 * .SinY + x * .CosY
   If SelfMode Then
     p1.x = x2 * .CosZ - y2 * .SinZ
     p1.y = x2 * .SinZ + y2 * .CosZ
     p1.z = z2 * .CosY - x * .SinY 'x not x2
   Else
     RotatePoint3D.x = x2 * .CosZ - y2 * .SinZ
     RotatePoint3D.y = x2 * .SinZ + y2 * .CosZ
     RotatePoint3D.z = z2 * .CosY - x * .SinY 'x not x2
   End If
  Else 'DoRotateZY=True
   x2 = x * .CosZ - y * .SinZ
   If SelfMode Then
     p1.y = x * .SinZ + y * .CosZ
     p1.x = z * .SinY + x2 * .CosY
     p1.z = z * .CosY - x2 * .SinY
   Else
     RotatePoint3D.y = x * .SinZ + y * .CosZ
     RotatePoint3D.x = z * .SinY + x2 * .CosY
     RotatePoint3D.z = z * .CosY - x2 * .SinY
   End If
  End If
 End With
 
End Function

Sub RotateSelfPage3D(page1 As Page3DType, SinCos3A As SinCos3AngType)
 Dim i As Long
 For i = 0 To page1.NumPoints
  RotatePoint3D page1.Point(i), SinCos3A, True
 Next

End Sub

Function RotatePage3D(page1 As Page3DType, SinCos3A As SinCos3AngType) As Page3DType
 Dim i As Long, n As Long
 n = page1.NumPoints
 With RotatePage3D
  .NumPoints = n
  .ColRGB = page1.ColRGB
  .SpValue = page1.SpValue
  ReDim .Point(n)
 
  For i = 0 To n
   .Point(i) = RotatePoint3D(page1.Point(i), SinCos3A)
  Next
 End With
 
End Function

Function RotateObject3D(Obj1 As Object3DType, SinCos3A As SinCos3AngType) As Object3DType
 Dim i As Long, n As Long
 RotateObject3D = Obj1
 With RotateObject3D
  n = Obj1.NumPages
  .CenterPoint = RotatePoint3D(Obj1.CenterPoint, SinCos3A)
  For i = 0 To n - 1
   .Page(i) = RotatePage3D(Obj1.Page(i), SinCos3A)
  Next
 End With
End Function

Sub RotateSelfObject3D(Obj1 As Object3DType, SinCos3A As SinCos3AngType)
 Dim i As Long, n As Long
 With Obj1
  n = .NumPages
  .CenterPoint = RotatePoint3D(.CenterPoint, SinCos3A)
  For i = 0 To n - 1
   RotateSelfPage3D .Page(i), SinCos3A
  Next
 End With
End Sub

Sub ScaleSelfPage3D(page1 As Page3DType, Scale1 As Single, ScaleOrigPoint As Point3DType)
 Dim i As Long, x As Single, y As Single, z As Single
 With ScaleOrigPoint
  x = .x
  y = .y
  z = .z
 End With
 With page1
  For i = 0 To .NumPoints
   With .Point(i)
    .x = (.x - x) * Scale1 + x
    .y = (.y - y) * Scale1 + y
    .z = (.z - z) * Scale1 + z
   End With
  Next
 End With
End Sub

Sub ScaleSelfObject3D(Obj1 As Object3DType, Scale1 As Single, ScaleOrigPoint As Point3DType)
 
 Dim i As Long, iPage As Long, x As Single, y As Single, z As Single
 With ScaleOrigPoint
  x = .x
  y = .y
  z = .z
 End With
 
 With Obj1
 For iPage = 0 To .NumPages - 1
  With .Page(iPage)
   For i = 0 To .NumPoints
    With .Point(i)
     .x = (.x - x) * Scale1 + x
     .y = (.y - y) * Scale1 + y
     .z = (.z - z) * Scale1 + z
    End With
   Next
  End With
 Next
 End With
 
End Sub

Function Line3D(p1 As Point3DType, p2 As Point3DType) As Line3DType
 With Line3D
  .p1 = p1
  .p2 = p2
 End With
End Function

Function MidPoint3D(p1 As Point3DType, p2 As Point3DType, Optional ByVal Weight1 As Single = 1) As Point3DType
 'return mid of 2 points
 With MidPoint3D
  .x = (p1.x * Weight1 + p2.x) / (1 + Weight1)
  .y = (p1.y * Weight1 + p2.y) / (1 + Weight1)
  .z = (p1.z * Weight1 + p2.z) / (1 + Weight1)
 End With
End Function

Function VectorOf2Points(p1 As Point3DType, p2 As Point3DType) As Vector3DType
 With VectorOf2Points
  .dx = p2.x - p1.x
  .dy = p2.y - p1.y
  .dz = p2.z - p1.z
 End With
End Function

Function VectorOfPoint(p As Point3DType) As Vector3DType
 With VectorOfPoint
  .dx = p.x
  .dy = p.y
  .dz = p.z
 End With
End Function

Function IsZeroVector(v As Vector3DType) As Integer
 With v
  IsZeroVector = .dx = 0 And .dy = 0 And .dz = 0
 End With
End Function

Function VnOfPage(page1 As Page3DType) As Vector3DType
 'return normal vector of a page using Cross product of first two vectors of plate(points 0,1,2)
 Dim v1 As Vector3DType, v2 As Vector3DType
 With page1
 '.Point(0) is Center! not a corner point
 If .NumPoints < 3 Then Exit Function 'invalid page
 v1 = VectorOf2Points(.Point(1), .Point(2))
 v2 = VectorOf2Points(.Point(2), .Point(3))
 If IsZeroVector(v1) Then
  If .NumPoints >= 4 Then
   v1 = v2
   v2 = VectorOf2Points(.Point(2), .Point(4))
  End If
 End If
 
 VnOfPage = CrossProduct(v1, v2)
 End With
End Function

Function Distance2Point(p1 As Point3DType, p2 As Point3DType, Optional ReturnSquareAns As Boolean = False) As Single
 'return distance between 2 points
 'if ReturnSquareAns be true then return square of answer.
 If ReturnSquareAns Then
  Distance2Point = (p2.x - p1.x) * (p2.x - p1.x) + (p2.y - p1.y) * (p2.y - p1.y) + (p2.z - p1.z) * (p2.z - p1.z)
 Else
  Distance2Point = Sqr((p2.x - p1.x) * (p2.x - p1.x) + (p2.y - p1.y) * (p2.y - p1.y) + (p2.z - p1.z) * (p2.z - p1.z))
 End If
End Function

Function LengthVector(v1 As Vector3DType, Optional ReturnSquareAns As Boolean = False) As Single
 'return length of a vector
 'if ReturnSquareAns be true then return square of answer.
 With v1
  If ReturnSquareAns Then
   LengthVector = .dx * .dx + .dy * .dy + .dz * .dz
  Else
   LengthVector = Sqr(.dx * .dx + .dy * .dy + .dz * .dz)
  End If
 End With
End Function

Function Area3Angular(p1 As Point3DType, p2 As Point3DType, p3 As Point3DType) As Single
 'return area of a Triangular
 Dim a As Single, b As Single, c As Single, p As Single
 a = Distance2Point(p1, p2)
 b = Distance2Point(p2, p3)
 c = Distance2Point(p3, p1)
 p = (a + b + c) / 2 'half of perimiter
 'in MAKING EXE WITH Advanced Options FOR FAST have some problem (-1.#IND) if ABS() remove!!
 Area3Angular = Sqr(Abs(p * (p - a) * (p - b) * (p - c))) 'don't remove ABS function
 'If Left(CStr(Area3Angular), 4) = "-1.#" Then MsgBox CStr(Area3Angular) + "  =>   " + CStr(Area3Angular)
 
End Function

Function AreaPage(page1 As Page3DType) As Single
 'Return area of a page (polygon) by dividing it to n-2 threeangles
 Dim i As Long, s As Single
 s = 0
 With page1
 For i = 2 To .NumPoints - 1
  s = s + Area3Angular(.Point(1), .Point(i), .Point(i + 1))
 Next
 End With
 AreaPage = s
End Function

Function PageLineIncidencePoint(PagePoint As Point3DType, PageVn As Vector3DType, LinePoint As Point3DType, LineVector As Vector3DType) As Point3DType
 'return incidence point of a page with a line
 'to faster radiation calculation both page and line are specified by two params: a point and a vector
 
 Dim t
 t = (PageVn.dx * LineVector.dx + PageVn.dy * LineVector.dy + PageVn.dz * LineVector.dz)
 If t = 0 Then t = 1E-20
 t = (PageVn.dx * (PagePoint.x - LinePoint.x) + PageVn.dy * (PagePoint.y - LinePoint.y) + PageVn.dz * (PagePoint.z - LinePoint.z)) / t
 With PageLineIncidencePoint
  .x = LineVector.dx * t + LinePoint.x
  .y = LineVector.dy * t + LinePoint.y
  .z = LineVector.dz * t + LinePoint.z
 End With
 
End Function

Sub GetTetaFiOfPoint(p As Point3DType, Teta1 As Single, Fi1 As Single, Optional AnsInRad As Boolean = True)
 Dim d
 If AnsInRad Then 'radian mode
  If p.x = 0 Then
   Teta1 = Pi2 * Sgn(p.y)
  ElseIf p.y = 0 Then If p.x > 0 Then Teta1 = 0 Else Teta1 = Pi
  Else
   Teta1 = Atn(p.y / p.x)
   If p.x < 0 Then Teta1 = (Pi - Teta1) * Sgn(p.y)
  End If
  d = Sqr(p.x * p.x + p.y * p.y)
  Fi1 = Atn(d / p.z): If Fi1 < 0 Then Fi1 = Fi1 + Pi
 Else
  If p.x = 0 Then
   Teta1 = 90 * Sgn(p.y)
  ElseIf p.y = 0 Then If p.x > 0 Then Teta1 = 0 Else Teta1 = 180
  Else
   Teta1 = Atn(p.y / p.x) * kRad2Deg
   If p.x < 0 Then Teta1 = (180 - Teta1) * Sgn(p.y)
  End If
  d = Sqr(p.x * p.x + p.y * p.y)
  Fi1 = Atn(d / p.z) * kRad2Deg: If Fi1 < 0 Then Fi1 = Fi1 + 180
 End If
End Sub

Function GetSinCosOfVector(Vn As Vector3DType) As SinCos3AngType
 'Return Sin and Cos Of Ax,Ay,Az Of a vector
 On Error Resume Next
 Dim a As Single, b As Single
 With GetSinCosOfVector
  a = Sqr(Vn.dx * Vn.dx + Vn.dy * Vn.dy)
  .CosX = Vn.dx / a
  .SinX = Vn.dy / a
  .CosY = .SinX
  .SinY = .CosX
  b = Sqr(a * a + Vn.dz * Vn.dz)
  .CosZ = Vn.dz / b
  .SinZ = a / b
 End With
End Function

Function GetSinCosForRotation_X_To_Vector(Vn As Vector3DType) As SinCos3AngType
 'Return SinCos3Ang needed to apply for rotate a point on the X axis to the position of Vn vector.
 'Rotate Order: X (0 degree) -> Y (-[90-az] degree) -> Z (Ax degree)
 On Error Resume Next
 Dim a As Single, b As Single
 With GetSinCosForRotation_X_To_Vector
  a = Sqr(Vn.dx * Vn.dx + Vn.dy * Vn.dy)
  b = Sqr(a * a + Vn.dz * Vn.dz) 'vector length
  .CosX = 1
  .SinX = 0
  .SinY = -Vn.dz / b
  .CosY = a / b
  .CosZ = Vn.dx / a
  .SinZ = Vn.dy / a
 End With
End Function

Function IsPointInPage(PointA As Point3DType, PageP As Page3DType, PageVn As Vector3DType) As Boolean
 'determine if the page point 'PointA' is in limited area that is defined by the page 'PageP' or not
 'PageVn: normal vector of 'PageP'
 'because of point is in the page P we can check for limited area as 2 dimentional problem
 'Note: if we don't know point is in page or not, we should first check it by mathematical formula of the page
 Dim i As Long, j As Long, p(2) As Point3DType
 Dim i1 As Long, i2 As Long
 Dim a As Single, b As Single, c As Single, e As Single
 
 Dim point1 As Point3DType, page1 As Page3DType, Vn As Vector3DType
 Dim SinCosAzy As SinCos3AngType
 point1 = PointA: page1 = PageP: Vn = PageVn
 
 'First export 3D problem into 2D problem:
 'Method: consider Vn by it's three angles it makes by axises x,y,z (Ax,Ay,Az)
 '        first rotate page and point by the magnitude of -(Ax) about Z axis,
 '        then by magnitude of -Az about y axis.
 
 If Vn.dx <> 0 Or Vn.dy <> 0 Then 'dx,dy=0 means a 2d problem such we want for.
  With SinCosAzy
   a = Sqr(Vn.dx * Vn.dx + Vn.dy * Vn.dy)
   .CosZ = Vn.dx / a
   .SinZ = -Vn.dy / a
   b = Sqr(a * a + Vn.dz * Vn.dz) 'vector length
   .CosY = Vn.dz / b
   .SinY = -a / b
  End With
  i1 = DoRotateZY: DoRotateZY = True
  RotatePoint3D point1, SinCosAzy, True
  RotateSelfPage3D page1, SinCosAzy
  DoRotateZY = i1
 End If
 
 With page1
  p(0) = .Point(1)
  For i = 2 To .NumPoints - 1
    p(1) = .Point(i): p(2) = .Point(i + 1)
    'p(0,1,2) define a Triangular part of page
    'p(j).x
    For i1 = 0 To 2
     i2 = (i1 + 1) Mod 3
     a = p(i2).y - p(i1).y
     b = p(i1).x - p(i2).x
     c = -a * p(i1).x - b * p(i1).y
     'e=ax+by+c : line define by p(i1),p(i2) (e<0 e=0 e>0)
     j = (i2 + 1) Mod 3 'p(j) is third point of Triangular
     e = Sgn(Round(a * point1.x + b * point1.y + c, 8))
     If e <> Sgn(a * p(j).x + b * p(j).y + c) And e <> 0 Then Exit For 'is not inside this 3angular
    Next 'i1
    If i1 = 3 Then 'is inside this 3angle so is in page
     IsPointInPage = True: Exit Function
    End If
  Next
 End With
 
 IsPointInPage = False
 
End Function

Function IfLinePassThroughPage(page1 As Page3DType, page1Vn As Vector3DType, LinePoint As Point3DType, LineVector As Vector3DType, Optional SideC As Integer = 1) As Integer
 'determine if the line pass from <<Limited Area>> define by a page
 Dim p As Point3DType
 p = PageLineIncidencePoint(page1.Point(1), page1Vn, LinePoint, LineVector)
 If IsVectorsSameOrOppositSide(VectorOf2Points(LinePoint, p), LineVector) <> SideC Then IfLinePassThroughPage = 0: Exit Function
 
 IfLinePassThroughPage = IsPointInPage(p, page1, page1Vn)
 
End Function

Function Interpolate(ByVal y1 As Single, ByVal y2 As Single, ByVal x1 As Single, ByVal x2 As Single, ByVal x As Single) As Single
If x2 = x1 Then
  Interpolate = (y1 + y2) / 2
 Else
  If x < x1 Then
    Interpolate = y1
  ElseIf x > x2 Then
    Interpolate = y2
  Else
    Interpolate = (y2 - y1) * (x - x1) / (x2 - x1) + y1
  End If
 End If
End Function

Function InterpolateCol(col1 As Integer, col2 As Integer, x1 As Single, x2 As Single, x As Single) As Integer
 'same as up except that for integer Integer Color Interpolation
 If x2 = x1 Then
  InterpolateCol = (col1 + col2) / 2
 Else
  If x < x1 Then
    InterpolateCol = col1
  ElseIf x > x2 Then
    InterpolateCol = col2
  Else
    InterpolateCol = (col2 - col1) * (x - x1) / (x2 - x1) + col1
  End If
 End If
End Function

Sub SetColGradArray(Optional ColorMode As Integer = 0)
 'Gradient color (Blue>Cyan>Yellow>Red)
 'require 1080 array elements (0 to ColGradMaxI)
 Dim r As Long, G As Long, b As Long, i As Long
 Static ColGradArrayMode As Integer, frst As Integer
 
 If frst = 0 Then frst = 1: ColGradArrayMode = -100 'invalid value for first time
 If ColGradArrayMode = ColorMode Then Exit Sub Else ColGradArrayMode = ColorMode
 
 ReDim ColGrad(3000)
 
 If ColGradArrayMode = 0 Then
 i = 0
 For b = 100 To 255 'dark blue -> blue
  ColGrad(i) = RGB(0, 0, b): i = i + 1
 Next
 
 'you can remove this but change next for wich: ColGrad(i) = RGB(0, 255 - B, B)
 For G = 0 To 255 'blue -> cyan
  ColGrad(i) = RGB(0, G, 255): i = i + 1
 Next
 
 For b = 255 To 0 Step -1 'cyan -> green
  ColGrad(i) = RGB(0, 255, b): i = i + 1
 Next
 
 For r = 0 To 255 'green -> yellow
  ColGrad(i) = RGB(r, 255, 0): i = i + 1
 Next
 
 '------------------------------------
 If False Then 'if add magenta?
  'yellow -> magenta
  For G = 255 To 0 Step -1
   ColGrad(i) = RGB(255, G, 255 - G): i = i + 1
  Next
  '255,0,255
 
  For G = 255 To 0 Step -1 'magenta -> Red
   ColGrad(i) = RGB(255, 0, G): i = i + 1
  Next
  
 Else
  
  For G = 255 To 0 Step -1 'yellow -> Red
   ColGrad(i) = RGB(255, G, 0): i = i + 1
  Next
  
 End If 'addmagenta
 '--------------------------------------
 
 For r = 255 To 80 Step -1 'Red -> dark Red
  ColGrad(i) = RGB(r, 0, 0): i = i + 1
 Next
 
 End If 'ColGradMode
 
 ColGradMaxI = i - 1
 ReDim Preserve ColGrad(ColGradMaxI)
 
End Sub

Function MakeXFlesh(ByVal L1 As Long, ByVal d As Long, Optional ColRGB As Long = 0) As Page3DType
Dim d2 As Long
d2 = d \ 2
With MakeXFlesh
 .NumPoints = 6
 .ColRGB = ColRGB
 ReDim .Point(6)
 .Point(0) = Point3D(L1 \ 2, 0, 0) 'Center Point
 .Point(1) = Point3D(0, 0, 0)
 .Point(2) = Point3D(L1 - d2, 0, 0)
 .Point(3) = Point3D(L1 - d, -d2, 0)
 .Point(4) = Point3D(L1, 0, 0)
 .Point(5) = Point3D(L1 - d, d2, 0)
 .Point(6) = .Point(2)
End With
End Function

Function MakeFlesh3D(ByVal L1 As Long, ByVal d As Long, SinCosA As SinCos3AngType, Optional ColRGB As Long = 0) As Page3DType
 MakeFlesh3D = RotatePage3D(MakeXFlesh(L1, d, ColRGB), SinCosA)
End Function

Sub SetAxisParams()
 'x,y,z Axises
 Dim Col As Long
 Col = 0
 FleshAxis(0) = MakeXFlesh(50, 3, Col)
 FleshAxis(1) = MakeFlesh3D(50, 3, SinCos3Angles(0, 0, 90), Col)
 FleshAxis(2) = MakeFlesh3D(110, 3, SinCos3Angles(0, -90, 0), Col)
 With TextXdraw
  .CommandStr = "MLML" 'M:Move , L:Line
  ReDim .Point(3)
  .Point(0) = Point3D(56, 2, 0)
  .Point(1) = Point3D(-4, -4, 0)
  .Point(2) = Point3D(0, 4, 0)
  .Point(3) = Point3D(4, -4, 0)
 End With
 With TextYdraw
  .CommandStr = "MLML"
  ReDim .Point(3)
  .Point(0) = Point3D(2, 56, 0)
  .Point(1) = Point3D(-4, -4, 0)
  .Point(2) = Point3D(4, 0, 0)
  .Point(3) = Point3D(-2, 2, 0)
 End With
End Sub

Function MakeHorizRectangle(ByVal dx, ByVal dy, Optional ColRGB As Long = &HFFFFFF) As Page3DType
 With MakeHorizRectangle
    .NumPoints = 4
    .ColRGB = ColRGB
    ReDim .Point(4)
    .Point(0) = Point3D(dx / 2, dy / 2, 0) 'Center
    .Point(1) = Point3D(0, 0, 0)
    .Point(2) = Point3D(dx, 0, 0)
    .Point(3) = Point3D(dx, dy, 0)
    .Point(4) = Point3D(0, dy, 0)
 End With
End Function

Function MakePage3D(p() As Point3DType, Optional ColRGB As Long = &HFFFFFF) As Page3DType
 'p(0) must be Center Of Page
 'p(1) to p(NumPoints) are 3D points
 Dim n As Long, i As Long
 n = UBound(p)
 With MakePage3D
    .NumPoints = n
    .ColRGB = ColRGB
    ReDim .Point(n)
    For i = 0 To n
     .Point(i) = p(i)
    Next
 End With
End Function

Function MakePageObjectnXnY(p As Single, q As Single, nx As Long, ny As Long) As Object3DType

Dim dx As Single, dy As Single, dx2 As Single, dy2 As Single, ix As Long, iy As Long, c As Long
Dim x As Single, y As Single

dx = p / nx: dy = q / ny
dx2 = dx / 2: dy2 = dy / 2
c = 0

With MakePageObjectnXnY
.NumPages = nx * ny
ReDim .Page(.NumPages - 1)

For ix = 0 To nx - 1
 x = ix * dx
 For iy = 0 To ny - 1
  y = iy * dy
  .Page(c) = MakePage3D(Point3DArrayByVal(x + dx2, y + dy2, 0, x, y, 0, x + dx, y, 0, x + dx, y + dy, 0, x, y + dy, 0))
  c = c + 1
 Next
Next

End With 'obj1

End Function

Function Point3DArrayByVal(ParamArray valueXYZ() As Variant) As Point3DType()
 'Input:  valueXYZ of points : 1,1,1 , 2,3,4 , 5,10,3 , ...
 'Output: point3D array()
 Dim n As Long, i As Long, c As Long, p() As Point3DType
 n = UBound(valueXYZ) 'must be 2 or 5 or 3k-1
 ReDim p((n + 1) \ 3 - 1)
 c = 0
 For i = 0 To n Step 3
  p(c) = Point3D(valueXYZ(i), valueXYZ(i + 1), valueXYZ(i + 2))
  c = c + 1
 Next
 Point3DArrayByVal = p
End Function

Function MakeSunObject(PCenter As Point3DType, ByVal Radius2D As Long, ByVal Col As Long, ByVal ColEdge As Long, Optional ByVal iCode As Integer = 8) As Object3DType
 With MakeSunObject
  .CenterPoint = PCenter
  .iCode = iCode
  .ColEdge = ColEdge
  .NumPages = 1
  ReDim .Page(0)
  .Page(0).ColRGB = Col
  ReDim .SpValues(0)
  .SpValues(0) = Radius2D
 End With
End Function

Sub SetLengthVector(v As Vector3DType, ByVal Length1 As Single)
Dim a As Single
With v
 a = Sqr(.dx * .dx + .dy * .dy + .dz * .dz)
 .dx = (.dx / a) * Length1
 .dy = (.dy / a) * Length1
 .dz = (.dz / a) * Length1
End With

End Sub

Function ConvertPoint2Page(p1 As Point3DType) As Page3DType
 With ConvertPoint2Page
  .ColRGB = 0
  .NumPoints = 1
  ReDim .Point(1)
  .Point(0) = p1
  .Point(1) = p1
 End With
End Function

Function ConvertVector2Point(v1 As Vector3DType) As Point3DType
 With ConvertVector2Point
  .x = v1.dx
  .y = v1.dy
  .z = v1.dz
 End With
End Function

Function ConvertPage2Obj(page1 As Page3DType, Optional iCode As Integer = 0, Optional ColEdge As Long = 0, Optional ObjNameStr As String = "") As Object3DType
 With ConvertPage2Obj
  .NumPages = 1
  .iCode = iCode
  .ColEdge = ColEdge
  .NameStr = ObjNameStr
  'Center Of Obj = Center Of Page:
  .CenterPoint = page1.Point(0)
  ReDim .Page(0)
  .Page(0) = page1
 End With
End Function

Sub ClearScene(Scene As Scene3DType)
 'clear a Scene from memory.
 'this sub is needed sometimes before using 'AddObjectToScene'
 With Scene
  Erase .Object(), .SpValues()
  .NumObjects = 0
  .CenterPoint = Point3D(0, 0, 0)
  .NameStr = ""
 End With
End Sub

Sub RemoveObjectFromScene(Scene As Scene3DType, ObjIndex As Long)
 'Remove object #ObjIndex from the scene
 Dim i As Long, ni As Long
 With Scene
  ni = .NumObjects - 1
  If ObjIndex < 0 Or ObjIndex > ni Then Exit Sub 'invalid object index
  For i = ObjIndex To ni - 1
   .Object(i) = .Object(i + 1)
  Next
  .NumObjects = ni
  ReDim Preserve .Object(ni)
 End With
 
End Sub

Function MakeCube(ByVal dx, ByVal dy, ByVal dz, Optional ColRGB As Long = &HFFFFFF, Optional ColEdge As Long, Optional iCode As Integer = 0) As Object3DType
 With MakeCube
  .NumPages = 6
  .iCode = iCode
  .ColEdge = ColEdge
  .CenterPoint = Point3D(dx / 2, dy / 2, dz / 2)
  ReDim .Page(6 - 1)
  .Page(0) = MakeHorizRectangle(dx, dy, ColRGB)
  .Page(1) = RotatePage3D(MakeHorizRectangle(dx, dz, ColRGB), SinCos3Angles(Pi2, 0, 0, True))
  .Page(2) = RotatePage3D(MakeHorizRectangle(dz, dy, ColRGB), SinCos3Angles(0, -Pi2, 0, True))
  
  .Page(3) = TransferPage3D(.Page(0), Vector3D(0, 0, dz))
  .Page(4) = TransferPage3D(.Page(1), Vector3D(0, dy, 0))
  .Page(5) = TransferPage3D(.Page(2), Vector3D(dx, 0, 0))
 End With
 
End Function

Function MakeRegularHorizPoly(ByVal n As Long, r As Single, Optional ColRGB As Long = &HFFFFFF) As Page3DType
 Dim i As Long, t As Single, dt As Single
 With MakeRegularHorizPoly
  .ColRGB = ColRGB
  .NumPoints = n
  ReDim .Point(n)
  .Point(0) = Point3D(0, 0, 0)
  dt = 2 * Pi / n
  For i = 0 To n - 1
   t = i * dt
   With .Point(i + 1)
    .x = r * Cos(t)
    .y = r * Sin(t)
    .z = 0
   End With
  Next
 End With
End Function

Function MakeHorizArcPoly(ByVal TetaRad As Single, ByVal n As Long, r As Single, Optional ColRGB As Long = &HFFFFFF) As Page3DType
 Dim i As Long, t As Single, dt As Single
 With MakeHorizArcPoly
  .ColRGB = ColRGB
  .NumPoints = n
  ReDim .Point(n)
  .Point(0) = Point3D(r * Cos(TetaRad / 2) / 2, r * Sin(TetaRad / 2) / 2, 0)
  dt = TetaRad / n
  For i = 0 To n - 1
   t = i * dt
   With .Point(i + 1)
    .x = r * Cos(t)
    .y = r * Sin(t)
    .z = 0
   End With
  Next
 End With
End Function

Function MakeCylinderArc(h As Single, TetaRad As Single, nTeta As Long, r As Single, Optional iCode As Integer = 0, Optional EndClose As Integer = 3, Optional ColEdge As Long = -1, Optional ColRGB As Long = &HFFFFFF, Optional SideColRGB As Long = &HFFFFFF) As Object3DType
    Dim pg1 As Page3DType, Obj1 As Object3DType, n As Long
    pg1 = MakeHorizArcPoly(TetaRad, nTeta, r, ColRGB)
    TransferSelfPage3D pg1, Vector3D(0, 0, -h / 2)
    Obj1 = Join2Page(pg1, TransferPage3D(pg1, Vector3D(0, 0, h)), SideColRGB, ColEdge, EndClose, iCode)
    ReversePage3D Obj1.Page(0)
    MakeCylinderArc = Obj1
End Function

Function MakeSemiCylinder(h As Single, TetaRad As Single, nTeta As Long, r As Single, Optional iCode As Integer = 0, Optional ColEdge As Long = -1, Optional ColRGB As Long = &HFFFFFF, Optional SideColRGB As Long = &HFFFFFF) As Object3DType
    Dim pg1 As Page3DType, Obj1 As Object3DType, n As Long
    pg1 = MakeHorizArcPoly(TetaRad, nTeta, r, ColRGB)
    TransferSelfPage3D pg1, Vector3D(0, 0, -h / 2)
    Obj1 = Join2Page(pg1, TransferPage3D(pg1, Vector3D(0, 0, h)), SideColRGB, ColEdge, 3, iCode) 'if change EndClose except 3 Must change next codes that remove base floor plate
    With Obj1
        ReversePage3D .Page(0)
        pg1 = .Page(.NumPages - 1) 'save last page: one of ends
        RemoveEndPages Obj1, 2 'remove last two page: the floor plate and one end
        AddPage2Obj pg1, Obj1
    End With
    MakeSemiCylinder = Obj1
End Function

Sub AddPage2Obj(pg1 As Page3DType, Obj1 As Object3DType)
    'add a page to end of pages of obj1
    Dim n As Long
    With Obj1
        n = .NumPages
        n = n + 1
        .NumPages = n
        ReDim Preserve .Page(n - 1)
        .Page(n - 1) = pg1
    End With
End Sub

Sub RemoveEndPages(Obj1 As Object3DType, nPgFromEnd As Long)
    'remove n pages from end of pages of obj1
    Dim n As Long
    With Obj1
        n = .NumPages
        n = n - nPgFromEnd: If n < 0 Then n = 0
        .NumPages = n
        If n >= 1 Then ReDim Preserve .Page(n - 1) Else Erase .Page
    End With
End Sub

Sub ReversePage3D(pg1 As Page3DType)
    'reverse points definition to reverse normal vector of the page
    Dim i As Long, j As Long, p As Point3DType
    With pg1
        j = .NumPoints
        For i = 1 To j \ 2
            p = .Point(i): .Point(i) = .Point(j): .Point(j) = p 'swap points i,j
            j = j - 1
        Next
    End With
End Sub

Function EvalPoint3D(p1 As Point3DType, ByVal ZAdd As Single) As Point3DType
 With EvalPoint3D
  .x = p1.x
  .y = p1.y
  .z = p1.z + ZAdd
 End With
End Function

Function EvalPage3D(page1 As Page3DType, ByVal ZAdd As Single) As Page3DType
 Dim n As Long, i As Long
 
 With EvalPage3D
  .ColRGB = page1.ColRGB
  .SpValue = page1.SpValue
  n = page1.NumPoints
  .NumPoints = n
  ReDim .Point(n)
  For i = 0 To n
   .Point(i) = EvalPoint3D(page1.Point(i), ZAdd)
  Next
 End With
 
End Function

Function Join2Page(page1 As Page3DType, page2 As Page3DType, Optional SideColRGB As Long = &HFFFFFF, Optional ColEdge As Long = 0, Optional EndClose As Integer = 3, Optional iCode As Integer = 0) As Object3DType
 Dim i As Long, a As Long, t As Single, dt As Single
 Dim x As Single, y As Single
 Dim n As Long, nTot As Long, n2 As Long
 Dim i2 As Long
 
 'EndClose: (0,1,2,3) -> 0=NoAdd  1=Add page1  2=Add Page2  3=Add Both page1,page2
 'in NoAdd mode we have only Side Pages
 
 '>> number of points for page1 and page2 must be equal
 '      Else for case that page2 is a single point!
 
 With Join2Page
 
  n = page1.NumPoints
  n2 = page2.NumPoints
  
  Select Case EndClose
  Case 0: nTot = n: a = -1
  Case 1: nTot = n + 1: a = 0
  Case 2: nTot = n + 1: a = -1
  Case 3: nTot = n + 2: a = 0
  End Select
  
  .CenterPoint = MidPoint3D(page1.Point(0), page2.Point(0))
  .iCode = iCode
  .ColEdge = ColEdge
  .NumPages = nTot
  ReDim .Page(nTot - 1)
  
  If EndClose And 1 Then .Page(0) = page1
  If EndClose And 2 Then .Page(nTot - 1) = page2
  '.page(0).ColRGB = sideColRGB
  '.page(ntot - 1).ColRGB = sideColRGB
  
  
  For i = 1 To n
   With .Page(i + a)
     .ColRGB = SideColRGB
     .NumPoints = 4
     ReDim .Point(4)
     If i < n Then i2 = i + 1 Else i2 = 1
     .Point(1) = page1.Point(i)
     .Point(2) = page1.Point(i2)
     If n2 > 1 Then
      .Point(3) = page2.Point(i2)
      .Point(4) = page2.Point(i)
     Else
      .Point(3) = page2.Point(1)
      .Point(4) = page2.Point(1)
     End If
     .Point(0) = MidPoint3D(.Point(1), .Point(3)) 'Center
   End With
  Next
  
 End With
End Function

Function CylinderElementPoly(ByVal iCode As Integer, ByVal CylinderH, ByVal CylinderTeta, ByVal dCylinderTeta, ByVal dCylinderH, Optional ColRGB As Long = &HFFFFFF) As Page3DType
 'Angles Must be in Rad mode.
 Dim a As Single, x As Single, y As Single
 Dim z As Single, t As Single, r As Single
 With CylinderElementPoly
  .ColRGB = ColRGB
  .NumPoints = 4
  ReDim .Point(4)
  z = CylinderH: GoSub CylinderElem10
  t = CylinderTeta
  .Point(1) = Point3D(r * Cos(t), r * Sin(t), z)
  t = t + dCylinderTeta
  .Point(2) = Point3D(r * Cos(t), r * Sin(t), z)
  z = z + dCylinderH: GoSub CylinderElem10
  .Point(3) = Point3D(r * Cos(t), r * Sin(t), z)
  t = CylinderTeta
  .Point(4) = Point3D(r * Cos(t), r * Sin(t), z)
  
  .Point(0) = MidPoint3D(.Point(1), .Point(3)) 'Center
 
 End With
 Exit Function
 
CylinderElem10:
If iCode And 1 Then 'Sphere
  r = Sqr(30 * 30 - (30 - z) * (30 - z))
ElseIf iCode And 16 Then 'SemiSphere
  r = Sqr(30 * 30 - z * z)
ElseIf iCode And 32 Then 'Dome
  a = z * z
  r = 24.28516401 + 0.080955265 * a - 0.02083452 * z ^ 2.5 + 0.001228612 * z * a
End If
Return

End Function

Function SphereElementPoly(ByVal iCode As Integer, ByVal SphereFi, ByVal SphereTeta, ByVal dSphereTeta, ByVal dSphereFi, Optional ColRGB As Long = &HFFFFFF) As Page3DType
 'Angles Must be in Rad mode.
 Dim a As Single, z As Single
 Dim fi As Single, t As Single, r As Single
 With SphereElementPoly
  .ColRGB = ColRGB
  .NumPoints = 4
  ReDim .Point(4)
  fi = SphereFi
  t = SphereTeta + dSphereTeta: GoSub SphereElem10
  .Point(1) = Point3D(r * Cos(t), r * Sin(t), z)
  t = SphereTeta
  .Point(2) = Point3D(r * Cos(t), r * Sin(t), z)
  fi = fi + dSphereFi: GoSub SphereElem10
  .Point(3) = Point3D(r * Cos(t), r * Sin(t), z)
  t = t + dSphereTeta
  .Point(4) = Point3D(r * Cos(t), r * Sin(t), z)
  .Point(0) = MidPoint3D(.Point(1), .Point(3)) 'Center
 
 End With
 Exit Function
 
SphereElem10:
 'Sphere (or Semi Sphere)
 z = 30 * Cos(fi)
 r = 30 * Sin(fi) 'r=R*sin(fi)
Return

End Function

Function MakeCylinderObj(Optional ByVal iCode As Integer = 1, Optional ByVal nCylinderH1 As Long = 10, Optional ByVal nCylinderTeta1 As Long = 40, Optional ColRGB As Long = &HFFFFFF, Optional ByVal ColEdge As Long = 0, Optional ObjNameStr As String = "") As Object3DType
 'iCode: '1=Sphere / 2=SpObject-Tower / 4=Find Radiation / 64=SunObject(=>Radius=SpValues(0))
         '16=SemiSphere / 32=Dome
 
 Dim nCylinderH, nCylinderTeta, CylinderH, CylinderTeta, dCylinderH, dCylinderTeta
 Dim ih As Long, i As Long, c As Long, UseSphere As Integer
 
 nCylinderH = nCylinderH1: nCylinderTeta = nCylinderTeta1
 UseSphere = 0
 'dCylinderH = 5 'delta H Cylinder (m)
 'nCylinderTeta = 40 '(must be 4*k) Num of Parts in each cross (Teta)
 
 With MakeCylinderObj
  
  .NameStr = ObjNameStr
  
  If iCode And 1 Then 'Sphere
   'dCylinderH = 60 / nCylinderH
   dCylinderH = Pi / nCylinderH: UseSphere = 1
   .CenterPoint = Point3D(0, 0, 30): .NameStr = "Sphere"
  ElseIf iCode And 16 Then 'SemiSphere
   'dCylinderH = 30 / nCylinderH
   dCylinderH = Pi / 2 / nCylinderH: UseSphere = 1
   .CenterPoint = Point3D(0, 0, 4 * 30 / (3 * 3.14159)): .NameStr = "Semi-Sphere"
  ElseIf iCode And 32 Then 'Dome
   dCylinderH = 60 / nCylinderH
   .CenterPoint = Point3D(0, 0, 17): .NameStr = "Dome"
  End If
  dCylinderTeta = 2 * Pi / nCylinderTeta
  
  .NumPages = nCylinderTeta * nCylinderH
  .iCode = iCode
  .ColEdge = ColEdge
  ReDim .Page(.NumPages - 1)
  
  c = 0
  For ih = 0 To nCylinderH - 1
   CylinderH = ih * dCylinderH
   For i = 0 To nCylinderTeta - 1
    CylinderTeta = i * dCylinderTeta
    If UseSphere Then
      .Page(c) = SphereElementPoly(iCode, CylinderH, CylinderTeta, dCylinderTeta, dCylinderH, ColRGB)
    Else
      .Page(c) = CylinderElementPoly(iCode, CylinderH, CylinderTeta, dCylinderTeta, dCylinderH, ColRGB)
    End If
    c = c + 1
   Next
  Next
 End With

End Function

Sub MakeObjByRotateCurve(ReturnObj As Object3DType, ByVal R_as_fZ As String, zBottom As Single, zTop As Single, nZ As Long, Teta1 As Single, Teta2 As Single, nTeta As Long, Optional iCode As Long = 0, Optional NameStr As String = "UserDefined", Optional ColEdge As Long = &HFFFFFF, Optional ColRGB As Long = &HCCCCCC, Optional OutSideOrInSide As Integer = 1)
 'Rotate a curve about an axis: curve is defined by R=f(Z) , this formula is in 'R_as_fZ' string
 'OutSideOrInside: 1 = normal vector Outside, 2 = Inside
 '** Angles Must be in Rad mode. **
 Dim z As Single, Teta As Single, iZ As Long, iTeta As Long, dz As Single, dTeta As Single, r As Single
 Dim c As Long, r1 As Single, r2 As Single, z1 As Single
 Dim Cost As Single, Sint As Single
 If nZ = 0 Or nTeta = 0 Then Exit Sub
 dz = (zTop - zBottom) / nZ
 dTeta = (Teta2 - Teta1) / nTeta
 
 R_as_fZ = UCase(R_as_fZ)
 
 With ReturnObj
 .NumPages = nTeta * nZ
 .iCode = iCode
 .NameStr = NameStr
 .ColEdge = ColEdge
 ReDim .Page(.NumPages - 1)
 
 c = 0
 
 For iZ = 0 To nZ - 1
  z = zBottom + iZ * dz
  GoSub ObjRotCur10
  r1 = r: z1 = z 'r,z for bottom floor
  z = z + dz: GoSub ObjRotCur10
  'r,z are r2,z2 (next floor)
  For iTeta = 0 To nTeta - 1
   Teta = Teta1 + iTeta * dTeta
   With .Page(c)
    .ColRGB = ColRGB
    .NumPoints = 4
    ReDim .Point(4)
    .Point(1) = Point3D(r1 * Cos(Teta), r1 * Sin(Teta), z1)
    Cost = Cos(Teta + dTeta): Sint = Sin(Teta + dTeta)
    .Point(2) = Point3D(r1 * Cost, r1 * Sint, z1)
    'Next Floor
    .Point(3) = Point3D(r * Cost, r * Sint, z)
    .Point(4) = Point3D(r * Cos(Teta), r * Sin(Teta), z)
      
    .Point(0) = MidPoint3D(.Point(1), .Point(3)) 'Center
     
   End With 'page(c)
   c = c + 1
  Next
 Next
 
 z = (zBottom + zTop) / 2
 .CenterPoint = Point3D(0, 0, z) 'Object Center
 
 End With 'ReturnObj
 
 Exit Sub

ObjRotCur10:
  r = FormScriptEditor.ScriptControl1.Eval(Replace(R_as_fZ, "Z", z))
Return

End Sub

Sub MakeObjByExtrude(ReturnObj As Object3DType, ByVal Y_as_fX As String, ByVal xMin As Single, ByVal xMax As Single, ByVal ndx As Long, ByVal LengthZ As Single, Optional ByVal nLengthZ As Long = 1, Optional iCode As Long = 0, Optional NameStr As String = "UserExtrude", Optional ColEdge As Long = &HFFFFFF, Optional ColRGB As Long = &HCCCCCC, Optional InSideOrOutSide As Integer = 0)
'VnInsideOutside: determine normal vector side of pages (think for y=x^2, inside=upward , outside=downward)
'                 0=Inside , 1=Outside

 Dim x As Single, y As Single, x1 As Single, y1 As Single, z As Single, z1 As Single
 Dim dx As Single, dz As Single
 Dim ix As Long, iZ As Long, i As Long, c As Long
 
 dx = (xMax - xMin) / ndx
 dz = LengthZ / nLengthZ
 z1 = -LengthZ / 2
 Y_as_fX = UCase(Y_as_fX)
 
 If InSideOrOutSide = 0 Then dz = -dz: z1 = -z1
 
 With ReturnObj
  
  .NameStr = NameStr
  
  .NumPages = ndx * nLengthZ
  .iCode = iCode
  .ColEdge = ColEdge
  ReDim .Page(.NumPages - 1)
  
  c = 0
  For ix = 0 To ndx - 1
   x = xMin + ix * dx: z = z1
   GoSub MakeObjByExtrude_10
   x1 = x: y1 = y: x = x + dx: GoSub MakeObjByExtrude_10
   For iZ = 0 To nLengthZ - 1
    z = z1 + iZ * dz
    With .Page(c)
     .ColRGB = ColRGB
     .NumPoints = 4
     ReDim .Point(4)
     .Point(1) = Point3D(x1, y1, z)
     .Point(2) = Point3D(x, y, z)
     .Point(3) = Point3D(x, y, z + dz)
     .Point(4) = Point3D(x1, y1, z + dz)
     .Point(0) = MidPoint3D(.Point(1), .Point(3)) 'Center point of page
    End With
    c = c + 1
   Next
  Next
 End With

Exit Sub

MakeObjByExtrude_10: 'find y as x
 y = FormScriptEditor.ScriptControl1.Eval(Replace(Y_as_fX, "X", x))
Return

End Sub

Function MakeYfXObj(ByVal xMin As Single, ByVal xMax As Single, ByVal ndx As Long, ByVal a As Single, ByVal b As Single, ByVal LengthZ As Single, Optional ByVal nLengthZ As Long = 1, Optional ByVal iCode As Integer = 128, Optional ColRGB As Long = &HFFFFFF, Optional ByVal ColEdge As Long = 0, Optional ByVal VnInsideOutside As Integer = 0, Optional ObjNameStr As String = "") As Object3DType
 'iCode: '1=Sphere / 2=SpObject-Tower / 4=Find Radiation / 8=cause shadow
         '16=SemiSphere / 32=Dome
         '64=SunObject(=>Radius=SpValues(0))
         '128=parabolic (collector)=> y=a*x^2+b*x (a=1/8,b=0) x:xMin to xMax step [(xMax-xMin)/ndx]

'VnInsideOutside: determine normal vector side of pages (think for y=x^2, inside=upward , outside=downward)
'                 0=Inside , 1=Outside

 Dim x As Single, y As Single, x1 As Single, y1 As Single, z As Single, z1 As Single
 Dim dx As Single, dz As Single
 Dim ix As Long, iZ As Long, i As Long, c As Long
 
 dx = (xMax - xMin) / ndx
 dz = LengthZ / nLengthZ
 z1 = -LengthZ / 2
 If VnInsideOutside = 0 Then dz = -dz: z1 = -z1
 
 With MakeYfXObj
  
  .NameStr = ObjNameStr
  
  If iCode And 128 Then 'Parabolic
  
  End If
  
  .NumPages = ndx * nLengthZ
  .iCode = iCode
  .ColEdge = ColEdge
  ReDim .Page(.NumPages - 1)
  
  c = 0
  For ix = 0 To ndx - 1
   x = xMin + ix * dx: z = z1
   GoSub MakeYfXObj10
   x1 = x: y1 = y: x = x + dx: GoSub MakeYfXObj10
   For iZ = 0 To nLengthZ - 1
    z = z1 + iZ * dz
    With .Page(c)
     .ColRGB = ColRGB
     .NumPoints = 4
     ReDim .Point(4)
     .Point(1) = Point3D(x1, y1, z)
     .Point(2) = Point3D(x, y, z)
     .Point(3) = Point3D(x, y, z + dz)
     .Point(4) = Point3D(x1, y1, z + dz)
     .Point(0) = MidPoint3D(.Point(1), .Point(3)) 'Center point of page
    End With
    c = c + 1
   Next
  Next
 End With

Exit Function

MakeYfXObj10: 'find y as x
 If iCode And 128 Then 'parabolic
   y = x * (a * x + b) 'y = ax2 + bx
 End If
Return

End Function

Function Triangular3D(p1 As Point3DType, p2 As Point3DType, p3 As Point3DType, ColRGB As Long) As Page3DType
 'make a three angle by it's three points
 With Triangular3D
  .NumPoints = 3
  ReDim .Point(3)
  .Point(1) = p1
  .Point(2) = p2
  .Point(3) = p3
  .Point(0) = MidPoint3D(p1, MidPoint3D(p2, p3), 2) 'Center Of Mass
  .ColRGB = ColRGB
 End With
End Function

Sub Make3AngularGridObj(Obj1 As Object3DType)
 Dim i As Long, j As Long, Obj2 As Object3DType, c As Long, n As Long, p As Point3DType
 c = 0 'number of pages of obj2
 Obj2 = Obj1
 With Obj1
 n = .NumPages
 For i = 0 To .NumPages - 1
  With .Page(i)
    If .NumPoints = 3 Then 'this is a Triangular -> so we half it to two Triangular
      If n < c + 2 Then 'add some array for new pages
        n = c + 2 + 200 'for more speed we add 100 elements each time we need
        ReDim Preserve Obj2.Page(n - 1)
      End If
      p = MidPoint3D(.Point(2), .Point(3))
      Obj2.Page(c) = Triangular3D(.Point(1), .Point(2), p, .ColRGB)
      Obj2.Page(c + 1) = Triangular3D(.Point(1), p, .Point(3), .ColRGB)
      c = c + 2
    Else
      For j = 2 To .NumPoints - 1
       c = c + 1 'add a new page to obj2
       If n < c Then 'add some array for new pages
         n = c + 200 'for more speed we add 100 elements each time we need
         ReDim Preserve Obj2.Page(n - 1)
       End If
       Obj2.Page(c - 1) = Triangular3D(.Point(1), .Point(j), .Point(j + 1), .ColRGB)
      Next 'j
    End If '.NumPoints=3
  End With
 Next 'i
 End With
 
 ReDim Preserve Obj2.Page(c)
 Obj2.NumPages = c
 Obj1 = Obj2
 
End Sub

Sub Make3AngularGridScene(Scene1 As Scene3DType)
 Dim i As Long, c As Long
 With Scene1
  For i = 0 To .NumObjects - 1
   c = .Object(i).iCode
   If (c And 4) > 0 Or (c And 8) > 0 Then
    Make3AngularGridObj .Object(i)
   End If
  Next
 End With
 
End Sub

Function CentroidOf3Anglular(TriangularPg As Page3DType) As Point3DType
 With TriangularPg
  CentroidOf3Anglular = MidPoint3D(.Point(1), MidPoint3D(.Point(2), .Point(3)), 2)
 End With
End Function

Sub AddObjectToScene(Scene As Scene3DType, AddObj As Object3DType)
 Dim n As Long
 With Scene
  n = .NumObjects
  .NumObjects = n + 1
  .CenterPoint = MidPoint3D(.CenterPoint, AddObj.CenterPoint, n)
  
  If n = 0 Then
   ReDim .Object(0) 'UBound of Null array cause an error!! so don't merge these
  ElseIf UBound(.Object) < n Then
   ReDim Preserve .Object(n)
  End If
  .Object(n) = AddObj
 End With
End Sub

Sub DrawLine3D(L1 As Line3DType)
Dim p As POINTAPI
With L1
 p = Point2Dof3D(.p1)
 MoveToEx gdi_BufferHdc, p.x, p.y, p
 p = Point2Dof3D(.p2)
 LineTo gdi_BufferHdc, p.x, p.y
End With

End Sub

Sub DrawPlusMark(p As Point3DType, ByVal l As Single, Optional UseScale As Single = -1, Optional AddCross As Integer = False)
 'Draw a + mark in the Horizontal plate
 'x,y: center of '+' mark  ,  l: length of each four line
 'UseScale: -1: draw mark by current scale, Other: use that for scaling (use a value such as 1 for a fix size Mark)
 Dim p3 As Point3DType, p2 As POINTAPI
 If UseScale > 0 Then l = l / View3DScale * UseScale
 p3 = p
 'p2 = Point2Dof3D(p3)
 'Ellipse gdi_BufferHdc, p2.x - 2, p2.y - 2, p2.x + 2, p2.y + 2
 
 p3.x = p.x - l: p3.y = p.y
 p2 = Point2Dof3D(p3)
 MoveToEx gdi_BufferHdc, p2.x, p2.y, p2
 p3.x = p.x + l
 p2 = Point2Dof3D(p3)
 LineTo gdi_BufferHdc, p2.x, p2.y
 
 p3.x = p.x: p3.y = p.y - l
 p2 = Point2Dof3D(p3)
 MoveToEx gdi_BufferHdc, p2.x, p2.y, p2
 p3.y = p.y + l
 p2 = Point2Dof3D(p3)
 LineTo gdi_BufferHdc, p2.x, p2.y
 
 If AddCross Then 'draw a x
  p3.x = p.x - l: p3.y = p.y - l
  p2 = Point2Dof3D(p3)
  MoveToEx gdi_BufferHdc, p2.x, p2.y, p2
  p3.x = p.x + l: p3.y = p.y + l
  p2 = Point2Dof3D(p3)
  LineTo gdi_BufferHdc, p2.x, p2.y
 
  p3.x = p.x + l: p3.y = p.y - l
  p2 = Point2Dof3D(p3)
  MoveToEx gdi_BufferHdc, p2.x, p2.y, p2
  p3.x = p.x - l: p3.y = p.y + l
  p2 = Point2Dof3D(p3)
  LineTo gdi_BufferHdc, p2.x, p2.y
 End If
 
End Sub

Sub DrawPolyDraw3D(PDraw As PolyDraw3DType, Optional iRelativeMode As Boolean = True)

'- this sub begin from first point and
'  continue in Relative or Absolute Mode
'  as is determined iRelativeMode
'PDraw.Point contain 3D points
'PDraw.CommandStr: such as "MLLLML.."
'M=MoveTo   L=LineTo

Dim p As POINTAPI, p3D As Point3DType, p0 As Point3DType
Dim x1 As Long, y1 As Long, z1 As Long, x2 As Long, y2 As Long, z2 As Long
Dim ch As String, cmd As String, i As Long

With PDraw
 cmd = UCase(.CommandStr)
 For i = 0 To Len(cmd) - 1
   p3D = .Point(i)
   If i <> 0 And iRelativeMode Then
    p3D.x = p3D.x + p0.x: p3D.y = p3D.y + p0.y: p3D.z = p3D.z + p0.z
   End If
   p = Point2Dof3D(p3D)
   ch = Mid(cmd, i + 1, 1)
   Select Case ch
   Case "M" 'MoveTo
        MoveToEx gdi_BufferHdc, p.x, p.y, p
   Case "L" 'LineTo
        LineTo gdi_BufferHdc, p.x, p.y
   End Select
   p0 = p3D
 Next
End With

End Sub

Sub DrawPage3D(page1 As Page3DType, Optional ColEdge As Long = 0)
 'use ColRGB=-1 to don't change fill color
 
 'ColEdge:
 '        -1 => set pen color to each page's fill color
 '        -2 => don't change pen color(Use Current Pen)
 '     other => use that color for Edges
 
 Dim i As Long, n As Long, cPen As Long, cBrush As Long
  
 With page1
  n = .NumPoints
  If n <= 0 Then Exit Sub
  Static p2() As POINTAPI
  ReDim Preserve p2(n - 1)
  
  For i = 1 To n
   p2(i - 1) = Point2Dof3D(.Point(i))
  Next
  
  cPen = ColEdge: cBrush = .ColRGB
  If cPen = -1 Then cPen = cBrush
  
  If DontChangeCurPenBrush = 0 Then SelectPenBrush cPen, cBrush
  Polygon gdi_BufferHdc, p2(0), n
  If DontChangeCurPenBrush = 0 Then DeleteLastPenBrush
  
 End With

End Sub

Sub DrawObject3D(Obj1 As Object3DType)
 Dim i As Long, n As Long, a() As Single, b() As Long, iCode As Integer, NoEdge As Boolean
 Dim ColEdge As Long
 
 With Obj1
  n = .NumPages - 1
  If n < 0 Then Exit Sub
  ColEdge = .ColEdge
  iCode = .iCode
  'check for some visibility
  If (iCode And (4 + 8)) = 0 Then 'no 4 no 8
   If .NameStr = DefaultHorizontalPlateStr Then
      If ShouldDrawHorizonPlate = 0 Then Exit Sub 'dont draw
   Else
      If ShouldDrawInActiveObjects = 0 Then Exit Sub 'dont draw
   End If
  End If
  
  'Check for SpCases:
  'this is remarked to use standard drawing method such as all other objects
  If iCode And 2 Then 'SpCase: TowerObject
   'TowerDrawVisibleParts obj1
   'Exit Sub
  End If
  
  If iCode And 64 Then 'SpCase: Sun Object
   DrawSunObject Obj1
   Exit Sub
  End If
  
  If DontSortPages = 0 Then
    'preparing to sort:
    ReDim a(n), b(n)
    For i = 0 To n
     a(i) = ZHeight3D(.Page(i).Point(0))  'point(0)=Center Point Of page
     b(i) = i
    Next
    QSort a(), b(), 0, n
    'must use b() as Drawing Order.
    'drawing:
    For i = 0 To n
         DrawPage3D .Page(b(i)), ColEdge
    Next
  Else
    'drawing without Sort & Hide:
    For i = 0 To n
         DrawPage3D .Page(i), ColEdge
    Next
  End If
 End With
End Sub

Sub DrawSunObject(SunObj As Object3DType)
 Dim p As POINTAPI, r As Long
 Dim cPen As Long, cBrush As Long
 On Error Resume Next
 With SunObj
  
  p = Point2Dof3D(.CenterPoint)
  r = .SpValues(0)
  cPen = .ColEdge: cBrush = .Page(0).ColRGB
  If cPen = -1 Then cPen = cBrush
  SelectPenBrush cPen, cBrush
  Ellipse gdi_BufferHdc, p.x - r, p.y - r, p.x + r, p.y + r
  DeleteLastPenBrush
  
 End With

End Sub

Sub DrawScene3D(Scene1 As Scene3DType)
 Dim n As Long, i As Long, a() As Single, b() As Long
 With Scene1
  n = .NumObjects - 1
  If n < 0 Then Exit Sub
  'preparing to sort:
  ReDim a(n), b(n)
  For i = 0 To n
   a(i) = ZHeight3D(.Object(i).CenterPoint)
   b(i) = i
  Next
  QSort a(), b(), 0, n
  'must use b() as Drawing Order.
  'drawing:
  For i = 0 To n
   DrawObject3D .Object(b(i))
  Next
 End With
End Sub

'Some Math-Functions:

Function CentroidPoint3D(p() As Point3DType, nfinish As Long) As Point3DType
 'x=sigma(x)/n  , y=sigma(y)/n  ,  z=sigma(z)/n
 Dim i As Long, x As Single, y As Single, z As Single
 For i = 0 To nfinish
  x = x + p(i).x
  y = y + p(i).y
  z = z + p(i).z
 Next
 With CentroidPoint3D
  .x = x / nfinish
  .y = y / nfinish
  .z = z / nfinish
 End With
End Function

Public Function CrossProduct(u As Vector3DType, v As Vector3DType) As Vector3DType
 'return UxV
 With CrossProduct
  .dx = u.dy * v.dz - u.dz * v.dy
  .dy = u.dz * v.dx - u.dx * v.dz
  .dz = u.dx * v.dy - u.dy * v.dx
 End With
End Function

Public Function IsVectorsSameOrOppositSide(v1 As Vector3DType, v2 As Vector3DType) As Integer
 'return +1 for same side , -1 for opposit side
 'if vectors be parallel this func return nearest answer
 Dim a As Integer
 a = Sgn(v1.dx) = Sgn(v2.dx) And Sgn(v1.dy) = Sgn(v2.dy) And Sgn(v1.dz) = Sgn(v2.dz)
 If a Then a = 1 Else a = -1
 IsVectorsSameOrOppositSide = a
End Function

'Quick-Sort Algoritm:
Public Sub QSort(a1() As Single, b1() As Long, ByVal start1 As Long, ByVal finish1 As Long)
'a1() -> Main Data
'b1() -> Sorted Index base on Initial values.
'        This is needed for find order of drawing pages/objects

Dim start As Long, finish As Long
Dim low As Long, middle As Long, p As Long, high As Long
Dim v1 As Single, i1 As Long
Dim check As Single

start = start1: finish = finish1
If finish <= start Then Exit Sub


p = 0
finishP(p) = finish

QSort_st: '------------------------------------------------------------------
finish = finishP(p)
low = start: high = finish
middle = (low + high) \ 2


check = a1(middle)

Do

         Do While a1(low) < check
          low = low + 1
         Loop


         Do While a1(high) > check
          high = high - 1
         Loop

        If low <= high Then
         'SWAP a1(low), a1(high):
         v1 = a1(low): a1(low) = a1(high): a1(high) = v1
         i1 = b1(low): b1(low) = b1(high): b1(high) = i1
         low = low + 1
         high = high - 1
        End If
Loop Until low > high

lowP(p) = low


If start < high Then 'CALL QSort(start, high)
 cline(p) = 0
 p = p + 1: finishP(p) = high: GoTo QSort_st
End If

QSort_1:
If low < finish Then 'CALL QSort(low, finish)
 cline(p) = 1
 p = p + 1: start = low: finishP(p) = finish: GoTo QSort_st
End If

QSort_END:
p = p - 1
If p >= 0 Then If cline(p) = 0 Then low = lowP(p): finish = finishP(p): GoTo QSort_1 Else GoTo QSort_END


End Sub

Public Sub SaveObjectToFile(FileStr As String, obj As Object3DType)
On Error Resume Next
Dim f As Long, i As Long, Pg As Long
f = FreeFile
Open FileStr For Output As f
If Err Then Close f: Exit Sub

With obj
 Print #f, "'Object Definition Created By Solar Calculator V1.0"
 Print #f,
 Print #f, "'Object Properties:"
 Print #f, "Name:"; .NameStr
 Print #f, "Center Point:("; CStr(.CenterPoint.x); ","; CStr(.CenterPoint.y); ","; CStr(.CenterPoint.z); ")"
 Print #f, "Code:"; CStr(.iCode)
 Print #f, "Wire Color:&h"; Hex(.ColEdge)
 Print #f, "Number Of Plates:"; CStr(.NumPages)
 Print #f,
 Print #f, "'Plates... :"
 Print #f, "'Color , Center-Point X,Y,Z , Point1 X,Y,Z , Point2 X,Y,Z , ..."
 Print #f,
 For Pg = 0 To .NumPages - 1
  With .Page(Pg)
   Print #f, "&h"; Hex(.ColRGB); " , ";
   For i = 0 To .NumPoints
    With .Point(i)
     Print #f, .x; ","; .y; ","; .z;
    End With 'point(i)
    If i <> .NumPoints Then Print #f, ","; Else Print #f, 'Next Line
   Next 'i
  End With 'Page(Pg)
 Next 'Pg
End With 'Obj

Close f

End Sub


Public Sub LoadObjectFromFile(FileStr As String, obj As Object3DType)
On Error Resume Next
Dim f As Long, i As Long, Pg As Long, a As String, v As Variant, n As Long, c As Long
f = FreeFile
Open FileStr For Input As f
If Err Then Close f: Exit Sub

With obj
 
 Do: Line Input #f, a: a = Replace(a, " ", ""): Loop While EOF(f) = 0 And (Left(a, 1) = "'" Or a = "")
 .NameStr = Mid(a, InStr(1, a, ":") + 1)
 Do: Line Input #f, a: a = Replace(a, " ", ""): Loop While EOF(f) = 0 And (Left(a, 1) = "'" Or a = "")
 With .CenterPoint
  i = InStr(1, a, "("): If i = 0 Then i = InStr(1, a, ":")
  v = Split(Mid(a, i + 1), ",")
  .x = Val(v(0))
  .y = Val(v(1))
  .z = Val(v(2))
  If Err Then Err.Clear
 End With
 Do: Line Input #f, a: a = Replace(a, " ", ""): Loop While EOF(f) = 0 And (Left(a, 1) = "'" Or a = "")
 .iCode = Val(Mid(a, InStr(1, a, ":") + 1))
 Do: Line Input #f, a: a = Replace(a, " ", ""): Loop While EOF(f) = 0 And (Left(a, 1) = "'" Or a = "")
 .ColEdge = Val(Mid(a, InStr(1, a, ":") + 1))
 Do: Line Input #f, a: a = Replace(a, " ", ""): Loop While EOF(f) = 0 And (Left(a, 1) = "'" Or a = "")
 .NumPages = Val(Mid(a, InStr(1, a, ":") + 1))
 ReDim .Page(.NumPages - 1)
 
 
 For Pg = 0 To .NumPages - 1
  With .Page(Pg)
   Do: Line Input #f, a: a = Replace(a, " ", ""): Loop While EOF(f) = 0 And (Left(a, 1) = "'" Or a = "")
   v = Split(a, ",")
   n = UBound(v)
   .ColRGB = CLng(v(0))
   .NumPoints = (n - 2 - 1) \ 3
   ReDim .Point(.NumPoints)
   c = 0
   For i = 1 To n Step 3
    With .Point(c)
      .x = CSng(v(i))
      .y = CSng(v(i + 1))
      .z = CSng(v(i + 2))
    End With 'point(i)
    c = c + 1
   Next 'i
  End With 'Page(Pg)
 Next 'Pg
End With 'Obj

Close f

End Sub


