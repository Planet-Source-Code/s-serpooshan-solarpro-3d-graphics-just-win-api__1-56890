VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FormGr 
   AutoRedraw      =   -1  'True
   Caption         =   "Graphics"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   404
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   632
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   7080
      Top             =   1380
   End
   Begin VB.PictureBox PicColTable 
      AutoRedraw      =   -1  'True
      Height          =   6000
      Left            =   30
      ScaleHeight     =   396
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   1
      ToolTipText     =   "(w/m2)"
      Top             =   30
      Width           =   345
   End
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00965A3C&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6000
      Left            =   810
      ScaleHeight     =   396
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   574
      TabIndex        =   0
      Top             =   30
      Width           =   8670
      Begin MSComDlg.CommonDialog CD1 
         Left            =   6210
         Top             =   750
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.CheckBox CheckWireFrame 
         Height          =   345
         Left            =   1200
         Picture         =   "FormGr.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Wire Frame"
         Top             =   30
         Width           =   345
      End
      Begin VB.CommandButton CommandViewctr 
         Height          =   345
         Index           =   6
         Left            =   30
         Picture         =   "FormGr.frx":027E
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Stop"
         Top             =   810
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.CommandButton CommandViewctr 
         Height          =   345
         Index           =   3
         Left            =   30
         Picture         =   "FormGr.frx":0337
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Rotate Around"
         Top             =   420
         Width           =   345
      End
      Begin VB.CommandButton CommandViewctr 
         Height          =   345
         Index           =   2
         Left            =   810
         Picture         =   "FormGr.frx":03D5
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Center"
         Top             =   30
         Width           =   345
      End
      Begin VB.CommandButton CommandViewctr 
         Height          =   345
         Index           =   0
         Left            =   30
         Picture         =   "FormGr.frx":0441
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Refresh"
         Top             =   30
         Width           =   345
      End
      Begin VB.CommandButton CommandViewctr 
         Height          =   345
         Index           =   1
         Left            =   420
         Picture         =   "FormGr.frx":04F1
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Reset View"
         Top             =   30
         Width           =   345
      End
   End
   Begin VB.Menu MnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExportResult 
         Caption         =   "Export Result       "
      End
      Begin VB.Menu MnuSaveCurImg 
         Caption         =   "Save Current Image    "
      End
      Begin VB.Menu mnuRecordImages 
         Caption         =   "Record Images..."
         Shortcut        =   ^{F8}
      End
      Begin VB.Menu mnuL0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCloseGr 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu MnuViewGr 
      Caption         =   "View"
      Begin VB.Menu MnuResetview 
         Caption         =   "&Reset...                     "
      End
      Begin VB.Menu mnuRefreshView 
         Caption         =   "Refresh"
      End
      Begin VB.Menu mnuStopUpdate 
         Caption         =   "Stop Update"
      End
      Begin VB.Menu mnul5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowResultData 
         Caption         =   "Show Result Data"
      End
      Begin VB.Menu mnuViewL20 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWireFrame 
         Caption         =   "Wire Frame"
      End
      Begin VB.Menu mnuSortPolygons 
         Caption         =   "Sort Polygons"
      End
      Begin VB.Menu mnuBackgroundCol 
         Caption         =   "Set Background Color"
      End
      Begin VB.Menu mnuExtraObjVisibility 
         Caption         =   "Extra Objets Visibility"
         Begin VB.Menu mnuS1 
            Caption         =   "Sun Object"
            Begin VB.Menu mnuDrawSunObj 
               Caption         =   "Draw Sun Object"
               Checked         =   -1  'True
            End
            Begin VB.Menu mnuDrawSunBeam 
               Caption         =   "Draw Sun Beam"
               Checked         =   -1  'True
            End
         End
         Begin VB.Menu mnuDrawAxises 
            Caption         =   "Draw Axises"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuDrawHorizonPlate 
            Caption         =   "Draw Horizontal Plate"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuDrawInActiveObjects 
            Caption         =   "Draw InActive Objects"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuHideShowExtraObj 
            Caption         =   "Hide All"
         End
      End
      Begin VB.Menu mnuAddHeaderTitle 
         Caption         =   "Add Header Title"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuL4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSunView 
         Caption         =   "Sun View"
      End
      Begin VB.Menu mnuDirection 
         Caption         =   "NSEW Directions"
         Begin VB.Menu mnuViewFront 
            Caption         =   "Front View"
         End
         Begin VB.Menu MnuViewSouth 
            Caption         =   "South View"
         End
         Begin VB.Menu mnuViewEast 
            Caption         =   "East View"
         End
         Begin VB.Menu mnuViewWest 
            Caption         =   "West View"
         End
         Begin VB.Menu mnuViewNorth 
            Caption         =   "North View "
         End
         Begin VB.Menu mnuViewTop 
            Caption         =   "Top  View"
         End
      End
      Begin VB.Menu mnuViewL1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuScale 
         Caption         =   "Scale"
         Begin VB.Menu mnuViewScale 
            Caption         =   "20%"
            Index           =   0
         End
         Begin VB.Menu mnuViewScale 
            Caption         =   "50%"
            Index           =   1
         End
         Begin VB.Menu mnuViewScale 
            Caption         =   "75%"
            Index           =   2
         End
         Begin VB.Menu mnuViewScale 
            Caption         =   "100%"
            Index           =   3
         End
         Begin VB.Menu mnuViewScale 
            Caption         =   "150%"
            Index           =   4
         End
         Begin VB.Menu mnuViewScale 
            Caption         =   "200%"
            Index           =   5
         End
         Begin VB.Menu mnuViewScale 
            Caption         =   "300%"
            Index           =   6
         End
         Begin VB.Menu mnuViewScale 
            Caption         =   "500%"
            Index           =   7
         End
         Begin VB.Menu mnuViewScale 
            Caption         =   "900%"
            Index           =   8
         End
      End
      Begin VB.Menu mnuViewNumeric 
         Caption         =   "Numeric Settings"
      End
      Begin VB.Menu mnuViewCenter 
         Caption         =   "Center Screen"
      End
      Begin VB.Menu mnuSetTargetObj 
         Caption         =   "Set Target Object"
      End
      Begin VB.Menu mnuViewL2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbsoluteMinMax 
         Caption         =   "Absolute Min-Max"
      End
      Begin VB.Menu mnuColorMode 
         Caption         =   "Color Mode"
      End
      Begin VB.Menu mnuMonochrome 
         Caption         =   "Monochrome"
         Begin VB.Menu mnuGrayScale 
            Caption         =   "Gray Scale             "
         End
         Begin VB.Menu mnuColorModeCustom 
            Caption         =   "Others..."
         End
      End
   End
   Begin VB.Menu mnDefine 
      Caption         =   "Define"
      Begin VB.Menu mnuDefineComponents 
         Caption         =   "Components...             "
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuNewObject 
         Caption         =   "New Object"
      End
      Begin VB.Menu mnuRemoveObject 
         Caption         =   "Remove Object"
      End
      Begin VB.Menu mnuDefieL1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDefine 
         Caption         =   "Method"
         Index           =   0
         Shortcut        =   +{F1}
      End
      Begin VB.Menu mnuDefine 
         Caption         =   "Location"
         Index           =   1
         Shortcut        =   +{F2}
      End
      Begin VB.Menu mnuDefine 
         Caption         =   "Date"
         Index           =   2
         Shortcut        =   +{F3}
      End
      Begin VB.Menu mnuDefine 
         Caption         =   "Time      "
         Index           =   3
         Shortcut        =   +{F4}
      End
      Begin VB.Menu mnuDefine 
         Caption         =   "Options"
         Index           =   4
         Shortcut        =   +{F5}
      End
      Begin VB.Menu mnuDefine 
         Caption         =   "Sp. Case Object"
         Index           =   5
         Shortcut        =   +{F6}
      End
   End
   Begin VB.Menu mnuTable 
      Caption         =   "Table"
      Begin VB.Menu mnuG_as_Hour 
         Caption         =   "G (w/m2) as Hour    "
      End
      Begin VB.Menu mnuG_as_Day 
         Caption         =   "G as Day"
      End
      Begin VB.Menu mnuG_as_Month 
         Caption         =   "G as Month"
      End
      Begin VB.Menu mnuL1Tbl 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAnnualEvsLatitude 
         Caption         =   "Annual Energy vs latitude"
      End
      Begin VB.Menu mnuAnnualEvsSlope 
         Caption         =   "Annual Energy vs Slope"
      End
      Begin VB.Menu mnuL9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCurTable 
         Caption         =   "Save Current Table"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuTriangularGrid 
         Caption         =   "Make Triangular Grid"
      End
      Begin VB.Menu mnuTransformObject 
         Caption         =   "Transform Object"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuToolsL1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFindViewFactor 
         Caption         =   "Find View Factor (F1-2)    "
      End
   End
   Begin VB.Menu mnuCalculate 
      Caption         =   "Calculate"
   End
   Begin VB.Menu mnuAnimate 
      Caption         =   "Animate"
      Begin VB.Menu MnuAnimRotate 
         Caption         =   "Rotate Around (Teta)   "
         Index           =   0
      End
      Begin VB.Menu MnuAnimRotate 
         Caption         =   "Rotate Above (Fi)"
         Index           =   1
      End
      Begin VB.Menu MnuAnimRotate 
         Caption         =   "Random Motion"
         Index           =   2
      End
   End
   Begin VB.Menu mnuWindowTitle 
      Caption         =   "Window"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "FormGr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

DefSng A-Z

Dim frm As Form, Xfrm(20) As Single, Yfrm(20) As Single, Nfrm As Long
Dim ViewTeta, ViewFi, ViewScale
Dim StopDrawShape As Integer, ShouldInitPic1 As Integer
Dim GMinTot As Single, GMaxTot As Single
Dim TimeTotal As Single
Dim ColTable() As Long 'an array to keep colors of Col-Table
Dim LastSelIndex As Integer
Dim ShouldSetComponentG As Integer
Dim ShouldStopAnim As Integer, IsInAnim As Integer

Sub DrawGrShapes()
 If StopDrawShape Then Exit Sub
 
 Dim i As Long, c As Long, tm, p As Point3DType, p2 As POINTAPI
 'Caption = "Teta:" + Format(ViewTeta, "000") + "  /  Fi:" + Format(ViewFi, "000")
 tm = Timer
 
 Call ResetFormViewSettings
 'Clear Screen By BkG Color '(fill background by the color)
 Call FillHwndOut
 
 'SelectQbPenBrush 0, 16, 0 'don't work
 SelectPenBrush &HFFFFFF
 If DontSortPages Then
    DeleteObject SelectObject(gdi_BufferHdc, GetStockObject(NULL_BRUSH))
 End If
 DrawScene3D UserScene
 
 If ShouldDrawAxises Then
  SelectPenBrush 0
  
  For i = 0 To 2
   DrawPage3D FleshAxis(i), 0
  Next
  DeleteLastPenBrush
  
  SelectPenBrush 0 '&HFFFFFF
  DrawPolyDraw3D TextXdraw 'X text
  DeleteLastPenBrush
  SelectPenBrush 0 '&HBBBBBB
  DrawPolyDraw3D TextYdraw 'Y text
  DeleteLastPenBrush
 End If
 
 If ShouldDrawSun Then DrawSunObject SunObject
 If ShouldDrawSunBeam Then
  SelectQbPenBrush 14
  DrawLine3D Line3D(Point3D(0, 0, 0), SunObject.CenterPoint)
  'SelectPenBrush RGB(255, 255, 0), &HFFFF, 2
  'DrawLine3D Line3D(SunObject.CenterPoint, Point3D(SunObject.CenterPoint.x, SunObject.CenterPoint.y, 0))
  p = Point3D(SunObject.CenterPoint.x, SunObject.CenterPoint.y, 0)
  DrawPlusMark p, 5, 1
  p2 = Point2Dof3D(p)
  Ellipse gdi_BufferHdc, p2.x - 2, p2.y - 2, p2.x + 2, p2.y + 2
  'DeleteLastPenBrush
 End If
 
 
 'RefreshGDIwindow -> instead of this, just call bitblt and print header if it's needed before calling pic1.refresh method
 BitBlt gdi_HdcOut, 0, 0, gdi_dx, gdi_dy, gdi_BufferHdc, 0, 0, SRCCOPY
 If ShouldAddHeader Then PrintHeader Pic1, HeaderText, Pic1.ScaleWidth - Pic1.TextWidth(HeaderText) - 4, 1
 Pic1.Refresh
 If ShouldRedeawGrShapes Then ShouldRedeawGrShapes = 0 'because we do it
 
 If ShouldRecordImages Then Call DoRecordImg 'save images

End Sub

Sub SetComponentG()
 'ComponentG : determine what component of G must use for color of plates:
 '             1=GDir  2=Gdiff  4=GRef
 
 Load FormComponents 'to set check boxes with current ComponentG
 ComponentG = 0
 Dim i As Long
 For i = 0 To 2
  If FormComponents.CheckComponent(i).Value Then ComponentG = ComponentG + 2 ^ i
 Next
 'ColorModeV=0 => Normal Color Gradient , ColorModeV = 7 => White-Black (1,2,4=R,G,B)
 Call ReCalculate
End Sub

Sub ReCalculate()
 UserFindRadiationScene UserdMinutes, ComponentG, True, UserTraceAsCollector, ColorModeV   '0
 Call SetGMinMaxTot
 Call SetDataGridValues
 Call DrawColTable
 Call DrawGrShapes
End Sub

Private Sub DoResetView()
With UserScene.DefaultView
 ViewTeta = .Teta
 ViewFi = .fi
 ViewScale = .View3DScale
End With
SetViewWholeParams UserScene.DefaultView
'SetViewShift2D Pic1.ScaleWidth \ 2, Pic1.ScaleHeight * 0.8

DrawGrShapes

End Sub

Private Sub CheckWireFrame_Click()
 DontChangeCurPenBrush = CheckWireFrame.Value
 DontSortPages = CheckWireFrame.Value
 mnuWireFrame.Checked = CheckWireFrame.Value
 mnuSortPolygons.Checked = (DontSortPages = 0)
 DrawGrShapes
End Sub


Private Sub CommandViewctr_Click(Index As Integer)
Pic1.SetFocus
Select Case Index
Case 0
 Call DrawGrShapes
Case 1
 Call DoResetView
Case 2
 Call mnuViewCenter_Click
Case 3
 Call MnuAnimRotate_Click(0)
Case 6
 ShouldStopAnim = 1
End Select
End Sub

'Private Sub CommandViewTarget_Click()
' SetViewShift3D Point3D(-80, 0, 0)
' DrawGrShapes
'End Sub

Private Sub Form_Activate()
 FormGrIsVisible = FormGr.Visible
 DoEvents
 Call Form_Paint
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 If KeyAscii = 27 And IsInAnim Then ShouldStopAnim = 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 ReleaseGDI
 ShouldInitPic1 = 0
 ShouldStopAnim = 1
 FormGrIsVisible = False
 Unload FormViewSettings 'if is loaded
 'MsgBox "Clear GDI Devices..."
End Sub

Private Sub Form_Resize()
'Set Window Size
Dim kx As Long, ky As Long
On Error Resume Next

FormGrIsVisible = FormGr.Visible

If WindowState = vbMinimized Then Exit Sub

If Width < 5000 Then Width = 5050: Exit Sub
If Height < 4500 Then Height = 4550: Exit Sub

kx = ScaleWidth: ky = ScaleHeight
With Pic1
 .Width = kx - .Left: .Height = ky - .Top
 '.Cls
 If ShouldRecordImages Then .Cls 'cause to delete unused image area -> this is necessary when we want to record images as bitmap!
 ShouldInitPic1 = 1
 PicColTable.Height = .Height
End With
 
DrawColTable

ShouldRedeawGrShapes = True
Call Timer1_Timer

End Sub

Private Sub Form_GotFocus()
 Call Form_Paint
End Sub

Private Sub Form_Paint()
 If StopDrawShape = 0 Then Call RefreshGDIwindow
End Sub


Private Sub mnuAbout_Click()
 MsgBox Space(6) + "Solar Radiation Calculator 1.0" + Space(10) + Chr(13) + vbCrLf + Space(6) + "By:" + Serpoosh + " - (C) " + CStr(2002), , "About.."
End Sub

Private Sub mnuAbsoluteMinMax_Click()
 Dim s As String, a As Integer, b As Single
 a = 1 - Abs(UseAbsoluteGMinMax)
 If a Then
    If AbsoluteGMin = 0 Then b = GMinTot Else b = AbsoluteGMin
    s = InputBox("Enter value for G(Min) ", "Enter Absolute Value for G (Min)", b): If s = "" Then Exit Sub 'user cancel
    AbsoluteGMin = Val(s)
    If AbsoluteGMax = 0 Then b = GMaxTot Else b = AbsoluteGMax
    s = InputBox("Enter value for G(Max) ", "Enter Absolute Value for G (Max)", b): If s = "" Then Exit Sub 'user cancel
    AbsoluteGMax = Val(s)
 End If
 UseAbsoluteGMinMax = a
 mnuAbsoluteMinMax.Checked = a
 Call ReCalculate
End Sub

Private Sub mnuAddHeaderTitle_Click()
On Error Resume Next
Dim v As Integer, a As String
v = 1 - Abs(mnuAddHeaderTitle.Checked)
ShouldAddHeader = v
If ShouldAddHeader Then
    a = InputBox("Type Header Title", "Add Header Title", HeaderText)
    If a = "" Then ShouldAddHeader = False Else HeaderText = a: Call DrawGrShapes
End If
mnuAddHeaderTitle.Checked = ShouldAddHeader
End Sub

Private Sub mnuAnimate_Click()
 ShouldStopAnim = 1 'it is Menu Title
End Sub

Sub DoRecordImg()
    On Error Resume Next
    Dim a As String
    RecordImgID = RecordImgID + 1
    a = Format(RecordImgID, "000") + ".bmp"
    SavePicture Pic1.Image, RecordFolder + a
End Sub

Private Sub MnuAnimRotate_Click(Index As Integer)
 Dim t As Single, t1 As Single, t2 As Single, f1 As Single
 Dim dt As Single, dFi As Single, a As String
 If IsInAnim Then ShouldStopAnim = 1: Exit Sub
 CommandViewctr(6).Visible = True

 ShouldStopAnim = 0
 t1 = ViewTeta: f1 = ViewFi
 If Index = 0 Then
  dt = 5
 ElseIf Index = 1 Then
  dFi = -4: ViewFi = 90
 ElseIf Index = 2 Then
  dt = 5: dFi = 3
 End If
 

 IsInAnim = 1
 Do
  ViewTeta = ViewTeta + dt
  ViewFi = ViewFi + dFi: If Abs(ViewFi) > 100 Then dFi = -dFi
  If ViewFi < 0 Then ViewFi = 0: dFi = 3
  If Index = 2 Then dt = dt + (Rnd - 0.5) * 0.7: dFi = dFi + (Rnd - 0.5) * 0.4
  SetView3D ViewTeta, ViewFi
  DrawGrShapes
  DoEvents
  If ShouldStopAnim Then ShouldStopAnim = 0: Exit Do
 Loop
 
 If FormGrIsVisible = 0 Then Exit Sub 'when animating and press close form button
 
 ViewFi = f1 ': ViewTeta = t1
 SetView3D ViewTeta, ViewFi
 
 CommandViewctr(6).Visible = False
 
 Call DrawGrShapes
 IsInAnim = 0
End Sub


Private Sub mnuBackgroundCol_Click()
 On Error Resume Next
 CD1.ShowColor
 If Err Then Err.Clear: Exit Sub 'Cancel
 gdi_BackgrFillCol = GetRealNearestColor(Pic1.hdc, CD1.Color)
 Call DrawGrShapes
End Sub

Private Sub mnuCalculate_Click()
 Call ReCalculate
End Sub

Private Sub mnuCloseGr_Click()
 FormDataGrid.Hide
 Unload Me
End Sub

Private Sub mnuColorMode_Click()
 If ColorModeV <> 0 Then
  ColorModeV = 0
  SetComponentG
 End If
End Sub

Private Sub mnuColorModeCustom_Click()
 Dim c1 As Integer
 c1 = ColorModeV
 FormColorModeCustom.Show 1
 If ColorModeV <> c1 Then
  SetComponentG
 End If
End Sub

Private Sub mnuCurTable_Click()
 Call mnuExportResult_Click
End Sub

Private Sub mnuDefine_Click(Index As Integer)
Dim s As String

 With FormWizard
  .PictureTopHide.Visible = True
  .CheckShowOnStartUp.Visible = False
  s = .Caption: .Caption = "Modify Problem"
  .SSTab1.Tab = Index + 1
  .LabelDefineTop.Caption = .SSTab1.Caption
  .Show 1
  .PictureTopHide.Visible = False
  .CheckShowOnStartUp.Visible = True
  .Caption = s
 End With
 
 Call SetGMinMaxTot
 SetDataGridValues
 Call DrawColTable
 ShouldRefreshFormGr = True

End Sub

Private Sub mnuDefineComponents_Click()
 FormComponents.Show
End Sub

Private Sub mnuDefineTime_Click()
FormWizard.SSTab1.Tab = 4
End Sub

Private Sub mnuDrawAxises_Click()
 With mnuDrawAxises
  If .Checked Then .Checked = 0 Else .Checked = True
  ShouldDrawAxises = .Checked
 End With
 Call DrawGrShapes
End Sub

Private Sub mnuDrawHorizonPlate_Click()
 With mnuDrawHorizonPlate
  If .Checked Then .Checked = 0 Else .Checked = True
  ShouldDrawHorizonPlate = .Checked
 End With
 Call DrawGrShapes
End Sub

Private Sub mnuDrawInActiveObjects_Click()
 With mnuDrawInActiveObjects
  If .Checked Then .Checked = 0 Else .Checked = True
  ShouldDrawInActiveObjects = .Checked
 End With
 Call DrawGrShapes
End Sub

Private Sub mnuDrawSunBeam_Click()
 With mnuDrawSunBeam
  If .Checked Then .Checked = 0 Else .Checked = True
  ShouldDrawSunBeam = .Checked
 End With
 Call DrawGrShapes
End Sub

Private Sub mnuDrawSunObj_Click()
 With mnuDrawSunObj
  If .Checked Then .Checked = 0 Else .Checked = True
  ShouldDrawSun = .Checked
 End With
 Call DrawGrShapes
End Sub

Private Sub mnuFindViewFactor_Click()
 FormVFGetObjects.Show 1
End Sub

Private Sub mnuGrayScale_Click()
 If ColorModeV <> 7 Then
  ColorModeV = 7
  SetComponentG
 End If
End Sub

Private Sub mnuHideShowExtraObj_Click()
 Dim stp1 As Integer, v As Boolean
 With mnuHideShowExtraObj
  If .Caption = "Hide All" Then v = True: .Caption = "Show All" Else v = False: .Caption = "Hide All"
 End With
 stp1 = StopDrawShape
 StopDrawShape = True
 
 If CBool(ShouldDrawAxises) = v Then mnuDrawAxises_Click
 If CBool(ShouldDrawHorizonPlate) = v Then mnuDrawHorizonPlate_Click
 If CBool(ShouldDrawInActiveObjects) = v Then mnuDrawInActiveObjects_Click
 If CBool(ShouldDrawSunBeam) = v Then mnuDrawSunBeam_Click
 If CBool(ShouldDrawSun) = v Then mnuDrawSunObj_Click
 
 StopDrawShape = stp1
 Call DrawGrShapes

End Sub

Sub DoPaint()
 Call Form_Paint
End Sub

Private Sub mnuNewObject_Click()
 Dim x1 As Single, y1 As Single, w1 As Long, n1 As Long, VSh3DOrig1 As Point3DType
 x1 = ViewXshift: y1 = ViewYshift
 VSh3DOrig1 = View3DShiftOrig
 w1 = FormGr.WindowState: FormGr.WindowState = vbMinimized
 n1 = UserScene.NumObjects
 FormCreateObj.Show 1
 SetViewShift3D VSh3DOrig1
 SetViewScale3D ViewScale
 SetView3D ViewTeta, ViewFi
 SetViewShift2D x1, y1
 FormGr.WindowState = w1
 ShouldInitPic1 = 1
 If n1 <> UserScene.NumObjects Then Call ReCalculate 'if add a new object
End Sub

Private Sub mnuRecordImages_Click()
On Error Resume Next
Dim a As String
If ShouldRecordImages = False Then
    With FormRecordSettings
        .TextFirstIndex.Text = RecordImgID + 1
        .Show 1
        If .IsCanceled Then Exit Sub
        RecordFolder = .TextOutputFolder.Text
        If Right$(RecordFolder, 1) <> "\" Then RecordFolder = RecordFolder + "\"
        ShouldAddHeader = .CheckAddHeader.Value <> 0
        HeaderText = .TextHeaderText.Text
        Pic1.Cls 'cause to remove unused area of pic1.image !!
        ShouldRecordImages = True
        RecordImgID = Val(.TextFirstIndex.Text) - 1
        mnuRecordImages.Caption = "Stop Recording!"
        Call DrawGrShapes 'save current image as first one
    End With
Else
    ShouldRecordImages = False
    mnuRecordImages.Caption = "Record Images..."
    MsgBox RecordImgID & " frames saved!", vbInformation
End If
End Sub

Private Sub mnuRefreshView_Click()
 Call DrawGrShapes
End Sub

Private Sub mnuRemoveObject_Click()
 Dim n1 As Long, k As Long
 n1 = UserScene.NumObjects
 
 FormSelectObject.Show 1
 k = FormSelectObject.SelectedIndex
 If k >= 0 Then
   RemoveObjectFromScene UserScene, k
 End If
 ShouldInitPic1 = 1
 If n1 <> UserScene.NumObjects Then Call ReCalculate 'if add a new object

End Sub

Private Sub MnuResetview_Click()
 DoResetView
End Sub

Private Sub mnuSetTargetObj_Click()
 On Error Resume Next
 FormSelectObject.Show 1
 TargetObjIndex = FormSelectObject.SelectedIndex
 If TargetObjIndex >= 0 Then 'if don't Cancel
 With UserScene.Object(TargetObjIndex)
  SetViewShift3D .CenterPoint
 End With
 End If
 Call DrawGrShapes

End Sub

Private Sub mnuShowResultData_Click()
 On Error Resume Next
 FormDataGrid.Show
 FormDataGrid.WindowState = vbNormal
End Sub

Private Sub mnuSortPolygons_Click()
    If DontSortPages Then DontSortPages = False Else DontSortPages = True
    mnuSortPolygons.Checked = (DontSortPages = 0)
    Call DrawGrShapes
End Sub

Private Sub mnuStopUpdate_Click()
 StopDrawShape = Not CBool(StopDrawShape)
 mnuStopUpdate.Checked = StopDrawShape
 If StopDrawShape = False Then Call DrawGrShapes
End Sub

Private Sub mnuSunView_Click()
 GetTetaFiOfPoint SunObject.CenterPoint, ViewTeta, ViewFi, False
 SetView3D ViewTeta, ViewFi, False
 Call DrawGrShapes
End Sub

Private Sub mnuTransformObject_Click()
 On Error Resume Next
 Dim c As Long
 FormSelectObject.Show 1
 ShouldRefreshFormGr = True
 If Err Then Err.Clear: Exit Sub 'Cancel Selection
 c = FormSelectObject.SelectedIndex
 If c < 0 Then Exit Sub 'Cancel Selection
 iObjTransform = c
 FormTransforms.Show 1
 iObjTransform = -1 'no object is avalable for apply transform in OK button
 ShouldRefreshFormGr = True
 'If FormTransforms.CancelPressed Then Exit Sub
 If ShouldReCalculate Then
  ShouldReCalculate = 0
  Call SetComponentG 'to re-calculate & redraw
 End If
 DoEvents
 FormTransforms.SetAsDefault
 
End Sub

Private Sub mnuTriangularGrid_Click()
 Make3AngularGridScene UserScene
 Call DrawGrShapes
End Sub

Private Sub mnuViewCenter_Click()
 SetViewShift2D Pic1.ScaleWidth \ 2, Pic1.ScaleHeight * 0.7
 Call DrawGrShapes
End Sub

Private Sub mnuViewEast_Click()
 ViewTeta = 90
 SetView3D ViewTeta, ViewFi, False
 Call DrawGrShapes
End Sub

Private Sub mnuViewFront_Click()
 ViewFi = 90 '85
 SetView3D ViewTeta, ViewFi, False
 Call DrawGrShapes
End Sub

Private Sub mnuViewNorth_Click()
 ViewTeta = 180
 SetView3D ViewTeta, ViewFi, False
 Call DrawGrShapes
End Sub

Private Sub mnuViewNumeric_Click()
 On Error Resume Next
 Call ResetFormViewSettings
 FormViewSettings.Show
 'If .Tag = "Cancel" Then Exit Sub
 
End Sub

Sub ResetFormViewSettings()
 Dim b As Variant, i As Long
 b = Array(ViewTeta, ViewFi, ViewScale, ViewXshift, ViewYshift)
 For i = 0 To 4
  FormViewSettings.TextVS(i).Text = CStr(Round(Val(b(i)), 2))
 Next
End Sub

Sub DoViewNumericSettings()
 With FormViewSettings
  If .Visible = False Then Exit Sub
  ViewTeta = Val(.TextVS(0).Text)
  ViewFi = Val(.TextVS(1).Text)
  ViewScale = Val(.TextVS(2).Text)
  ViewXshift = Val(.TextVS(3).Text)
  ViewYshift = Val(.TextVS(4).Text)
  SetView3D ViewTeta, ViewFi, False
  SetViewScale3D ViewScale
  SetViewShift2D ViewXshift, ViewYshift
  Call DrawGrShapes
 End With

End Sub

Private Sub mnuViewScale_Click(Index As Integer)
Dim a As Variant
a = Array(20, 50, 75, 100, 150, 200, 300, 500, 900)
ViewScale = a(Index) / 100
SetViewScale3D ViewScale
Call DrawGrShapes

End Sub

Private Sub MnuViewSouth_Click()
 ViewTeta = 0
 SetView3D ViewTeta, ViewFi, False
 Call DrawGrShapes
End Sub

Private Sub mnuViewTop_Click()
 ViewFi = 0
 SetView3D ViewTeta, ViewFi, False
 Call DrawGrShapes
End Sub

Private Sub mnuViewWest_Click()
 ViewTeta = -90
 SetView3D ViewTeta, ViewFi, False
 Call DrawGrShapes
End Sub

Private Sub mnuWireFrame_Click()
 CheckWireFrame.Value = 1 - CheckWireFrame.Value
End Sub

Private Sub Timer1_Timer()
On Error Resume Next

If ShouldInitPic1 And StopDrawShape = 0 Then
  ReleaseGDI
  InitializeGDI Pic1
  ShouldInitPic1 = 0
  ShouldRedeawGrShapes = 0
  Call DrawGrShapes
End If

'check for refresh (if any Form of this project is moved)
'Dim i As Long, c As Integer, n1 As Long
'n1 = Nfrm: Nfrm = Forms.Count

'If Nfrm <> n1 Then 'update windows list in 'Window' menu:
'   c = 0
'   Dim u As Long
'   u = mnuWindowItem.UBound
'   For i = 0 To Nfrm - 1
'    If c > u Then Load mnuWindowItem(c)
'    If Forms(i).Visible Then mnuWindowItem(c).Caption = Trim(Forms(i).Caption): c = c + 1
'   Next
'   For i = c To mnuWindowItem.UBound 'don't replace by 'u' because it is some items added
'    mnuWindowItem(i).Visible = False
'   Next
'End If

'c = 1
'If n1 > Nfrm Then 'if a form is removed we should redraw screen!
' c = 0: Call DoPaint
'End If
'For i = 0 To Nfrm - 1 'if a form is moved in other location we must redraw screen!
' If Xfrm(i) <> Forms(i).Left Or Yfrm(i) <> Forms(i).Top Then
'  If c Then c = 0: Call DoPaint
'  Xfrm(i) = Forms(i).Left: Yfrm(i) = Forms(i).Top
' End If
'Next


If ShouldRedeawGrShapes Then
 ShouldRedeawGrShapes = 0
 Call DrawGrShapes
End If

'Dim a As Long
'a = FormGr.Pic1.Image.Width

If ShouldRefreshFormGr Then Call RefreshGDIwindow: ShouldRefreshFormGr = 0
'If ShouldSetComponentG Then ShouldSetComponentG = ShouldSetComponentG - 1
'If ShouldSetComponentG = 1 Then
' ShouldSetComponentG = 0
' Call SetComponentG
'End If

End Sub

Sub SetGMinMaxTot()
  With UserScene
   GMinTot = .SpValues(0)
   GMaxTot = .SpValues(1)
  End With
End Sub

Sub DrawColTable()
 Dim n As Integer, i As Integer, Col As Long
 Dim y As Long, dy As Long, G As Single, x1 As Long, w As Integer
 Dim uR As Integer, uG As Integer, uB As Integer, uRGB As Integer, c As Integer
 SetGMinMaxTot
 x1 = PicColTable.Left + PicColTable.Width + 1
 x1 = x1 + (Pic1.Left - x1 - TextWidth(CStr(Round(GMaxTot)))) / 2
 Cls
 w = PicColTable.ScaleWidth
 y = PicColTable.Top + 2: dy = 15
 n = PicColTable.ScaleHeight \ dy   'number of rows can write
 ReDim ColTable(n - 1)
 
 uR = Sgn(ColorModeV And 1): uG = Sgn(ColorModeV And 2): uB = Sgn(ColorModeV And 4)
 uRGB = uR Or uG Or uB

 For i = 0 To n - 1
    CurrentX = x1: CurrentY = y
    G = Interpolate(GMaxTot, GMinTot, 0, n - 1, i)
    Print CStr(Round(G));
    If uRGB Then
     c = InterpolateCol(20, 255, GMinTot, GMaxTot, G)
     Col = RGB(c * uR, c * uG, c * uB)
    Else
     Col = ColGrad(InterpolateCol(0, ColGradMaxI, GMinTot, GMaxTot, G))
    End If
    
    ColTable(i) = Col
    PicColTable.Line (0, i * dy)-Step(w, dy), Col, BF
    y = y + dy
 Next
 PicColTable.Line (0, n * dy)-Step(w, dy), ColTable(n - 1), BF
 LastSelIndex = -1
End Sub

Private Sub Pic1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 ShouldStopAnim = 1
 If Button <> 2 Then Exit Sub
 Dim Col As Long
 Call SelCurrentCol(GetPixel(gdi_BufferHdc, x, y))
End Sub

Private Sub Pic1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Static xm1 As Long, ym1 As Long

If Button = 2 Then
 Call SelCurrentCol(GetPixel(gdi_BufferHdc, x, y))
 Exit Sub
End If

Dim dx, dy
Dim redraw1

If xm1 = 0 Or Button = 0 Then
 xm1 = x: ym1 = y: Exit Sub
End If

' Shift: 1=Shift  , 2=Ctrl , 4=Alt

redraw1 = 0

dx = x - xm1: dy = y - ym1

If Abs(dy) > 1 Then
 redraw1 = 1
 ym1 = y
 If Shift = vbShiftMask Then 'Shift key
  SetViewShift2D 0, dy, True
 ElseIf Shift = vbCtrlMask Then  'Ctrl Key
  ViewScale = ViewScale - dy * 0.02
  If ViewScale <= 0 Then ViewScale = 0.000001
  SetViewScale3D ViewScale
 ElseIf Shift = vbAltMask Then 'Alt Key
  With View3DShiftOrig
   '.y = .y + dx * 0.5
  End With
 Else
  'If dy > 8 Then dy = 8
  ViewFi = ViewFi - dy * 0.3
  If Abs(ViewFi) > 180 Then ViewFi = ViewFi - 360 * Sgn(ViewFi)
  SetView3D ViewTeta, ViewFi, False
 End If
End If

If Abs(dx) > 1 Then
 redraw1 = 1
 xm1 = x
 If Shift = vbShiftMask Then  'Shift
  SetViewShift2D dx, 0, True
 ElseIf Shift = vbCtrlMask Then  'Ctrl
  If Sgn(dx) = Sgn(dy) Then  'if both have the same effect
   ViewScale = ViewScale - dx * 0.02
   If ViewScale <= 0 Then ViewScale = 0.000001
   SetViewScale3D ViewScale
  End If
 ElseIf Shift = vbAltMask Then 'Alt Key
  With View3DShiftOrig
   '.x = .x + dy * 0.5
  End With
 Else
  'If dx > 8 Then dx = 8
  ViewTeta = ViewTeta - dx * 0.6
  If Abs(ViewTeta) > 360 Then ViewTeta = ViewTeta - 360 * Sgn(ViewTeta)
  SetView3D ViewTeta, ViewFi, False
End If
End If

If redraw1 Then
 DrawGrShapes
End If

End Sub


Sub SelCurrentCol(Col As Long)
 If Col = 0 Then Exit Sub 'Axis Color
 Dim i As Integer, dy As Integer, w As Integer
 dy = 15: w = PicColTable.ScaleWidth
 i = LastSelIndex
 If i <> -1 Then
  PicColTable.Line (0, i * dy)-Step(w, dy), ColTable(i), BF
  LastSelIndex = -1 'no select
 End If
 
 If Col = gdi_BackgrFillCol Or Col = HorizPageCol Then Exit Sub
 
 i = FindIndexByCol(Col)
 If i = -1 Then Exit Sub 'not click on the object
 LastSelIndex = i
 PicColTable.Line (0, i * dy)-Step(w, dy), &HFFFF&, B
 PicColTable.Line (1, i * dy + 1)-Step(w - 2, dy - 2), 0, B

End Sub


Function FindIndexByCol(Col As Long) As Integer
 'find table-index of a color
 Dim i As Integer, n As Integer, dif As Long, difMin As Long, iMin As Integer
 Dim rr As Integer, gg As Integer, bb As Integer
 Dim rr2 As Integer, gg2 As Integer, bb2 As Integer
 On Error Resume Next
 n = UBound(ColTable)
 If Err Then FindIndexByCol = -1: Exit Function
 GetRGBofCol Col, rr, gg, bb
 
 For i = 0 To n
  GetRGBofCol ColTable(i), rr2, gg2, bb2
  dif = (rr2 - rr) ^ 2 + (gg2 - gg) ^ 2 + (bb2 - bb) ^ 2
  If i = 0 Then
   difMin = dif: iMin = i
  Else
   If dif < difMin Then difMin = dif: iMin = i
  End If
 Next
 
 If difMin > 15000 Then iMin = -1 'no valid difference (was >8000 firstly)
 FindIndexByCol = iMin
 
End Function

Sub ReLoadForm()
 Call Form_Load
End Sub

Private Sub Form_Load()
 On Error Resume Next
 Dim i As Integer
 StopDrawShape = True
 
 CommandViewctr(6).Move CommandViewctr(3).Left, CommandViewctr(3).Top
 ShouldInitPic1 = 0
 'ComponentG = 7 '1+2+4=all components (Dir,dif,Ref)
 mnuDrawAxises.Checked = ShouldDrawAxises
 mnuDrawSunObj.Checked = ShouldDrawSun: mnuDrawSunBeam.Checked = ShouldDrawSunBeam
 mnuDrawHorizonPlate.Checked = ShouldDrawHorizonPlate: mnuDrawInActiveObjects.Checked = ShouldDrawInActiveObjects
 mnuSortPolygons.Checked = (DontSortPages = 0)
 CheckWireFrame.Value = Abs(DontChangeCurPenBrush)
 'FormGr.Top = 50
 FormGr.Width = 9600: FormGr.Height = 6750
 
 FormGr.Top = (Screen.Height - (MDIFormMain.Height - MDIFormMain.ScaleHeight + MDIFormMain.CoolBar1.Height) - FormGr.Height) / 2
 FormGr.Left = (Screen.Width - FormGr.Width) / 2
 
 With UserScene.DefaultView
  ViewTeta = .Teta
  ViewFi = .fi
  ViewScale = .View3DScale
 End With
 
 Call SetGMinMaxTot
 FormDataGrid.Show
 
 'Load FormComponents
 
 Pic1.BackColor = RGB(60, 90, 150)
 FormGr.Show
 FormGr.Refresh
 Call DrawColTable
 'use for fill pic1 background
 If gdi_BackgrFillCol < 0 Then
  gdi_BackgrFillCol = GetRealNearestColor(Pic1.hdc, Pic1.BackColor)
 End If
 
 SetViewShift2D Pic1.ScaleWidth \ 2, Pic1.ScaleHeight * 0.8
 SetViewScale3D ViewScale
 SetView3D ViewTeta, ViewFi
 
 If ShouldRStp > 25 Then End
 'these 2 lines should be active: *************
 StopDrawShape = False
 InitializeGDI Pic1
 
 'Call DrawGrShapes
 
 For i = 0 To Forms.Count - 1 'forms(0) is
   Xfrm(i) = Forms(i).Left: Yfrm(i) = Forms(i).Top
 Next
 
 Timer1.Enabled = True
 
End Sub

Sub SetDataGridValues()
Dim i As Integer, iObj As Long, a As Variant, r As Integer, c As Integer, q As Single, qs As String, v As String
Dim Et As String, ETotal As Single
Dim AObj As Single, GMin As Single, GMax As Single, GAve As Single, QDirPer As String, QDifPer As String, QRefPer As String, SigmaQTot As Single

On Error Resume Next

r = 0 'row 0 is some other info (lable and unit)
FormDataGrid.Grid1.Rows = 2 'clear previus rows if any

For iObj = 0 To UserScene.NumObjects - 1
 
 With UserScene.Object(iObj)
  
  If .iCode And 4 Then
   r = r + 1
   If 1 + r > FormDataGrid.Grid1.Rows Then FormDataGrid.Grid1.Rows = 1 + r + 4
   FormDataGrid.Grid1.TextMatrix(r, 0) = .NameStr
   AObj = .SpValues(0)
   'these values are computed consider some components (and are not absolute)
   GMin = .SpValues(1)
   GMax = .SpValues(2)
   GAve = .SpValues(3)
   SigmaQTot = GAve * AObj
   '.SpValues(4,5,6) are absolute (regardless of ComponentG)
   If ComponentG And 1 Then QDirPer = Round((.SpValues(4) / SigmaQTot) * 100) Else QDirPer = " - "
   If ComponentG And 2 Then QDifPer = Round((.SpValues(5) / SigmaQTot) * 100) Else QDifPer = " - "
   If ComponentG And 4 Then QRefPer = Round((.SpValues(6) / SigmaQTot) * 100) Else QRefPer = " - "
   'Total Time & Energy
   TimeTotal = .SpValues(7) 'minutes!
   ETotal = SigmaQTot * TimeTotal / 60 'Watt.Hour (whr)
   If TimeTotal = 0 Then Et = " - " Else Et = CStr(Round(ETotal / 1000, 1))
   'Q is Watt(j/s), E is Jule
    
   With FormDataGrid.Grid1
    a = Array(Round(AObj, 1), Round(GMin, 1), Round(GMax, 1), Round(GAve, 3), Round(SigmaQTot / 1000, 1), QDirPer, QDifPer, QRefPer, Et)
    c = 0
    For i = 1 To 9
     v = a(c)
     .TextMatrix(r, i) = v
     c = c + 1
    Next
   End With
 
  End If 'icode and 4
 
 End With 'UserScene.Object(iObj)
 
Next 'iObj

FormDataGrid.Grid1.Rows = 1 + r + 2 '2 blank rows at end

Err.Clear

End Sub


Private Sub mnuExportResult_Click()
 On Error Resume Next
 CD1.Flags = CD1.Flags Or cdlOFNFileMustExist Or cdlOFNOverwritePrompt
 CD1.Filter = "*.Htm;*.Html|*.Htm;*.Html|All Files|*.*"
 CD1.ShowSave
 FormGr.DoPaint 'Call RefreshGDIwindow (only a fast bitblt)
 If Err Then Exit Sub
 Dim f As Long, i As Long, j As Long
 f = FreeFile
 Open CD1.FileName For Output As f
 Print #f, "<Html><Body>"
 Print #f, "<Table border=1 cellpadding=2 cellspacing=3>"
 
 With FormDataGrid.Grid1
 For j = 0 To .Rows - 1
  Print #f, "<TR>";
  For i = 0 To .Cols - 1
    Print #f, "<TD>"; .TextMatrix(j, i); "</TD>";
  Next
  Print #f, "</TR>"
 Next
 End With
 
 Print #f, "</Table>"
 Print #f, "</Body></Html>"
 
 Close f
 
End Sub


Private Sub mnuG_as_Hour_Click()
 'Table
 'G Min,Max,Ave as Hour Change (from 7:00 to 12:00)
 Dim i As Integer, iObj As Long, ObjIndex() As Long, a As Variant, r As Integer, c As Integer, q As Single, qs As String, v As String
 Dim nObj As Long
 Dim Et As String, ETotal As Single
 Dim AObj As Single, GMin As Single, GMax As Single, GAve As Single
 Dim QDir As String, QDif As String, QRef As String, SigmaQTot As Single
 Dim TimeV As Single, nTimeBE1(1) As String, TimeIsLST1 As Integer
 Dim SunObj1 As Object3DType
 Dim f1 As Integer, f2 As Integer, fStr As String
 
 On Error Resume Next
 CD1.Flags = CD1.Flags Or cdlOFNOverwritePrompt
 CD1.Filter = "*.Htm;*.Html|*.Htm;*.Html|All Files|*.*"
 CD1.ShowSave
 FormGr.DoPaint 'Call RefreshGDIwindow (only a fast bitblt)
 If Err Then Exit Sub 'cancel save
 
 Dim FProgress As Form
 
 ReDim ObjIndex(UserScene.NumObjects - 1)
 nObj = 0
 For iObj = 0 To UserScene.NumObjects - 1
  With UserScene.Object(iObj)
   If (.iCode And 4) Then
    ObjIndex(nObj) = iObj: nObj = nObj + 1
   End If
  End With
 Next
 
 f1 = FreeFile
 fStr = CD1.FileName
 Open fStr For Output As f1
 
 If Err Then MsgBox Err.Description: Err.Clear: Close f1: Exit Sub
 
 For i = 0 To 1: nTimeBE1(i) = nTimeBeginEnd(i): Next
 TimeIsLST1 = TimeIsLST: TimeIsLST = True
 SunObj1 = SunObject
 
 Print #f1, "<Html><Body>"
 Print #f1, "<font color=""#005000"">"
 Print #f1, "City:"; sUser.CityName; "<br>"
 With FormWizard
 If .CheckOneDayPeriod.Value Then
  Print #f1, "Date:"; .ComboDayOfMonth(0).Text; "-"; .ComboMonth(0).Text; "<br>"
 Else
  Print #f1, "Date: Average Of Range("; .ComboDayOfMonth(0).Text; "-"; .ComboMonth(0).Text; " to "; .ComboDayOfMonth(1).Text; "-"; .ComboMonth(1).Text; ")<br>"
 End If
 End With 'formwizard
 Print #f1, "</font>"
 Print #f1, "<font color=""#440000"">"
 
 For iObj = 0 To nObj - 1
  With UserScene.Object(ObjIndex(iObj))
   Print #f1, "<br>"
   Print #f1, "Object:"; .NameStr; "<br>"
   Print #f1, "<Table border=1 cellpadding=2 cellspacing=3>"
   
   Print #f1, "<TR style='Color:#000070;'><TD>Hour</TD><TD>GMax</TD><TD>GAve</TD><TD>GMin</TD> ";
   Print #f1, "<TD>GDir</TD><TD>Gdif</TD><TD>GRef</TD></TR>"
   Set FProgress = New FormProgress
   FProgress.SetProperties "Create Hourly Table For '" + .NameStr + "'"
   FProgress.Show
   
   For TimeV = 6 * 60 To 12 * 60 Step UserdMinutes
   
    nTimeBeginEnd(0) = MakeTimeFromMin(TimeV)
    nTimeBeginEnd(1) = nTimeBeginEnd(0)
    FProgress.SetProgress (TimeV - 6 * 60) / (6 * 60)
    DoEvents
    If FProgress.CancelPressed Then Exit For
    'DoEvents
    UserFindRadiationObject ObjIndex(iObj), 0, ComponentG, True, UserTraceAsCollector, -1     'ColorMode=-1 is for dont set color of pages
  
    AObj = .SpValues(0)
    'not absolute
    GMin = .SpValues(1)
    GMax = .SpValues(2)
    GAve = .SpValues(3)
    SigmaQTot = GAve * AObj
    'absolute (regardless of ComponentG)
    QDir = .SpValues(4) 'Integral(GDir*Ai)
    QDif = .SpValues(5)
    QRef = .SpValues(6)
    
    'Total Time & Energy
    TimeTotal = .SpValues(7)
    ETotal = SigmaQTot * TimeTotal
    If TimeTotal = 0 Then Et = " - " Else Et = CStr(Round(ETotal / 1000, 1))
    
    'Print #f1, Left(nTimeBeginEnd(0), 5); " , "; 'hour
    Print #f1, "<TR>";
    
    Print #f1, "<TD>"; Format(TimeV / 60, "00.00"); "</TD>"; 'hour
    Print #f1, "<TD>"; Format(.SpValues(2), "0.00"); "</TD>";
    Print #f1, "<TD>"; Format(.SpValues(3), "0.00"); "</TD>";
    Print #f1, "<TD>"; Format(.SpValues(1), "0.00"); "</TD>";
    'Print #f1, "<TD>"; Format(.SpValues(3), "0.00"); "</TD>";
    For i = 4 To 6
     Print #f1, "<TD>"; Format(.SpValues(i) / AObj, "0.00"); "</TD>";
    Next
    
    Print #f1, "</TR>"
    
    'Exit For 'complete each object (then next object)
    
   Next 'TimeV
   
   Print #f1, "</Table>"
   
  End With '..object(i..)
  Unload FProgress
  Set FProgress = Nothing
  
 Next 'iObj
 
 For i = 0 To 1:  nTimeBeginEnd(i) = nTimeBE1(i): Next
 TimeIsLST = TimeIsLST1
 SunObject = SunObj1
 
 Print #f1, "</font>"
 Print #f1, "</Body></Html>"
 
 Close f1
 'open file:
 ShellExecute hwnd, "Open", fStr, "", "", 0

End Sub


Private Sub mnuG_as_Day_Click()
    
 'Table
 'G Min,Max,Ave as Day Change (from 1 to 365)
 Dim i As Integer, iObj As Long, ObjIndex() As Long, a As Variant, r As Integer, c As Integer, q As Single, qs As String, v As String
 Dim nObj As Long
 Dim Et As String, ETotal As Single
 Dim AObj As Single, GMin As Single, GMax As Single, GAve As Single
 Dim QDir As String, QDif As String, QRef As String, SigmaQTot As Single
 Dim DayV As Long, nDayBE1(1) As String
 Dim TimeV As Single, nTimeBE1(1) As String, TimeIsLST1 As Integer
 Dim SunObj1 As Object3DType
 Dim f1 As Integer, f2 As Integer, fStr As String
 
 On Error Resume Next
 CD1.Flags = CD1.Flags Or cdlOFNOverwritePrompt
 CD1.Filter = "*.Htm;*.Html|*.Htm;*.Html|All Files|*.*"
 CD1.ShowSave
 FormGr.DoPaint 'Call RefreshGDIwindow (only a fast bitblt)
 If Err Then Exit Sub 'cancel save
 
 Dim FProgress As Form
 
 ReDim ObjIndex(UserScene.NumObjects - 1)
 nObj = 0
 For iObj = 0 To UserScene.NumObjects - 1
  With UserScene.Object(iObj)
   If (.iCode And 4) Then
    ObjIndex(nObj) = iObj: nObj = nObj + 1
   End If
  End With
 Next
 
 f1 = FreeFile
 fStr = CD1.FileName
 Open fStr For Output As f1
 
 If Err Then MsgBox Err.Description: Err.Clear: Close f1: Exit Sub
 
 For i = 0 To 1: nDayBE1(i) = nDayBeginEnd(i): Next
 For i = 0 To 1: nTimeBE1(i) = nTimeBeginEnd(i): Next
 TimeIsLST1 = TimeIsLST: TimeIsLST = True
 SunObj1 = SunObject

 
 Print #f1, "<Html><Body>"
 Print #f1, "<font color=""#005000"">"
 Print #f1, "City:"; sUser.CityName; "<br>"
 
 If False Then 'PrintDayRange?
  With FormWizard
  If .CheckOneDayPeriod.Value Then
   Print #f1, "Date:"; .ComboDayOfMonth(0).Text; "-"; .ComboMonth(0).Text; "<br>"
  Else
   Print #f1, "Date: Average Of Range("; .ComboDayOfMonth(0).Text; "-"; .ComboMonth(0).Text; " to "; .ComboDayOfMonth(1).Text; "-"; .ComboMonth(1).Text; ")<br>"
  End If
  End With 'formwizard
 End If 'printDayRange
 
 Print #f1, "</font>"
 Print #f1, "<br>Units:     G (W/m2)   ,    E (Mj/m2 Per Day)<br>"
 Print #f1, "<font color=""#440000"">"
 
 For iObj = 0 To nObj - 1
  With UserScene.Object(ObjIndex(iObj))
   Print #f1, "<br>"
   Print #f1, "Object:"; .NameStr; "<br>"
   Print #f1, "<Table border=1 cellpadding=2 cellspacing=3>"
   
   Print #f1, "<TR style='Color:#000070;'><TD>Day</TD><TD>GMax</TD><TD>GAve</TD><TD>GMin</TD> ";
   Print #f1, "<TD>GDir</TD><TD>Gdif</TD><TD>GRef</TD><TD>E Ave</TD></TR>"
   Set FProgress = New FormProgress
   FProgress.SetProperties "Create Daily Table For '" + .NameStr + "'"
   FProgress.Show
   UserDoEventsOnPlate = 1
   
   For DayV = 1 To 365 Step UserDayStep
   
    nDayBeginEnd(0) = DayV: nDayBeginEnd(1) = nDayBeginEnd(0)
    nTimeBeginEnd(0) = "SR": nTimeBeginEnd(1) = "SN"
    FProgress.SetProgress (DayV / 365)
    'FProgress.Show
    DoEvents
    If FProgress.CancelPressed Then Exit For
    'DoEvents
    
    UserFindRadiationObject ObjIndex(iObj), UserdMinutes, ComponentG, 0, UserTraceAsCollector, -1       'ColorMode=-1 is for dont set color of pages
  
    AObj = .SpValues(0)
    Print #f1, "<TR>";
    
    Print #f1, "<TD>"; Format(DayV, "0"); "</TD>";  'Day
    Print #f1, "<TD>"; Format(.SpValues(2), "0.00"); "</TD>";
    Print #f1, "<TD>"; Format(.SpValues(3), "0.00"); "</TD>";
    Print #f1, "<TD>"; Format(.SpValues(1), "0.00"); "</TD>";
    For i = 4 To 6
     Print #f1, "<TD>"; Format(.SpValues(i) / AObj, "0.00"); "</TD>";
    Next
     
    'TimeTotal (Minutes) = .SpValues(7) , * 2 for calculate for whole day
    'If DayV >= 152 Then Stop
    Print #f1, "<TD>"; Format(.SpValues(7) * 60 * 2 / 1000000! * .SpValues(3), "0.000"); "</TD>";
   
    Print #f1, "</TR>"
    
    'Exit For 'complete each object (then next object)
    
   Next 'DayV
   
   Print #f1, "</Table>"
   
  End With '..object(i..)
  Unload FProgress
  Set FProgress = Nothing
  
 Next 'iObj
 
 For i = 0 To 1:  nDayBeginEnd(i) = nDayBE1(i): Next
 For i = 0 To 1:  nTimeBeginEnd(i) = nTimeBE1(i): Next
 TimeIsLST = TimeIsLST1
 SunObject = SunObj1
 UserDoEventsOnPlate = 0

 Print #f1, "</font>"
 Print #f1, "</Body></Html>"
 
 Close f1
 'open file:
 ShellExecute hwnd, "Open", fStr, "", "", 0
 
End Sub


Private Sub mnuG_as_Month_Click()
 'Table
 'G Min,Max,Ave as Month Change (from 1 to 12)
 Dim i As Integer, iObj As Long, ObjIndex() As Long, a As Variant, r As Integer, c As Integer, q As Single, qs As String, v As String
 Dim nObj As Long
 Dim Et As String, ETotal As Single
 Dim AObj As Single, GMin As Single, GMax As Single, GAve As Single
 Dim QDir As String, QDif As String, QRef As String, SigmaQTot As Single
 Dim MonthV As Long
 Dim DayV As Long, nDayBE1(1) As String
 Dim TimeV As Single, nTimeBE1(1) As String, TimeIsLST1 As Integer
 Dim SunObj1 As Object3DType
 Dim f1 As Integer, f2 As Integer, fStr As String
 
 On Error Resume Next
 CD1.Flags = CD1.Flags Or cdlOFNOverwritePrompt
 CD1.Filter = "*.Htm;*.Html|*.Htm;*.Html|All Files|*.*"
 CD1.ShowSave
 FormGr.DoPaint 'Call RefreshGDIwindow (only a fast bitblt)
 If Err Then Exit Sub 'cancel save
 
 Dim FProgress As Form
 
 ReDim ObjIndex(UserScene.NumObjects - 1)
 nObj = 0
 For iObj = 0 To UserScene.NumObjects - 1
  With UserScene.Object(iObj)
   If (.iCode And 4) Then
    ObjIndex(nObj) = iObj: nObj = nObj + 1
   End If
  End With
 Next
 
 f1 = FreeFile
 fStr = CD1.FileName
 Open fStr For Output As f1
 
 If Err Then MsgBox Err.Description: Err.Clear: Close f1: Exit Sub
 
 For i = 0 To 1: nDayBE1(i) = nDayBeginEnd(i): Next
 For i = 0 To 1: nTimeBE1(i) = nTimeBeginEnd(i): Next
 TimeIsLST1 = TimeIsLST: TimeIsLST = True
 SunObj1 = SunObject
 
 Print #f1, "<Html><Body>"
 Print #f1, "<font color=""#005000"">"
 Print #f1, "City:"; sUser.CityName; "<br>"
 
 If False Then 'PrintDayRange?
  With FormWizard
  If .CheckOneDayPeriod.Value Then
   Print #f1, "Date:"; .ComboDayOfMonth(0).Text; "-"; .ComboMonth(0).Text; "<br>"
  Else
   Print #f1, "Date: Average Of Range("; .ComboDayOfMonth(0).Text; "-"; .ComboMonth(0).Text; " to "; .ComboDayOfMonth(1).Text; "-"; .ComboMonth(1).Text; ")<br>"
  End If
  End With 'formwizard
 End If 'printDayRange
 
 Print #f1, "</font>"
 Print #f1, "<br>Units:     G (W/m2)   ,    E (Mj/m2 Per Day)<br>"
 Print #f1, "<font color=""#440000"">"
 
 For iObj = 0 To nObj - 1
  With UserScene.Object(ObjIndex(iObj))
   Print #f1, "<br>"
   Print #f1, "Object:"; .NameStr; "<br>"
   Print #f1, "<Table border=1 cellpadding=2 cellspacing=3>"
   
   Print #f1, "<TR style='Color:#000070;'><TD>Month</TD><TD>GMax</TD><TD>GAve</TD><TD>GMin</TD> ";
   Print #f1, "<TD>GDir</TD><TD>Gdif</TD><TD>GRef</TD><TD>E Ave</TD></TR>"
   Set FProgress = New FormProgress
   FProgress.SetProperties "Create Monthly Table For '" + .NameStr + "'"
   
   FProgress.Show
   FProgress.Refresh
   UserDoEventsOnPlate = 1
   
   For MonthV = 1 To 12
    nDayBeginEnd(0) = DayByDate(1, CStr(MonthV), False)
    If MonthV = 12 Then nDayBeginEnd(1) = 365 Else nDayBeginEnd(1) = DayByDate(1, CStr(MonthV + 1), False) - 1
    nTimeBeginEnd(0) = "SR": nTimeBeginEnd(1) = "SN" 'IF CHANGE SR,SN Must Change E(MJ/m2) Field (we multiply TimeTot by 2 because of half day calculation)
    FProgress.SetProgress (MonthV / 12)
    'FProgress.Show
    DoEvents
    If FProgress.CancelPressed Then Exit For
    'DoEvents
    
    UserFindRadiationObject ObjIndex(iObj), UserdMinutes, ComponentG, 0, UserTraceAsCollector, -1       'ColorMode=-1 is for dont set color of pages

    AObj = .SpValues(0)
    'Print #f1, Left(nTimeBeginEnd(0), 5); " , "; 'hour
    Print #f1, "<TR>";
    
    Print #f1, "<TD>"; Format(MonthV, "0"); "</TD>";  'hour
    Print #f1, "<TD>"; Format(.SpValues(2), "0.00"); "</TD>";
    Print #f1, "<TD>"; Format(.SpValues(3), "0.00"); "</TD>";
    Print #f1, "<TD>"; Format(.SpValues(1), "0.00"); "</TD>";
    'Print #f1, "<TD>"; Format(.SpValues(3), "0.00"); "</TD>";
    For i = 4 To 6
     Print #f1, "<TD>"; Format(.SpValues(i) / AObj, "0.00"); "</TD>";
    Next
    'TimeTotal (Minutes) = .SpValues(7) , * 2 for calculate for whole day
    Print #f1, "<TD>"; Format(.SpValues(7) * 60 * 2 / 1000000! * .SpValues(3) / UserDayTotal, "0.000"); "</TD>";
    Print #f1, "</TR>"
    
    'Exit For 'complete each object (then next object)
    
   Next 'MonthV
   
   Print #f1, "</Table>"
   
  End With '..object(i..)
  Unload FProgress
  Set FProgress = Nothing
  
 Next 'iObj
 
 For i = 0 To 1:  nDayBeginEnd(i) = nDayBE1(i): Next
 For i = 0 To 1:  nTimeBeginEnd(i) = nTimeBE1(i): Next
 TimeIsLST = TimeIsLST1
 SunObject = SunObj1
 UserDoEventsOnPlate = 0
  
 Print #f1, "</font>"
 Print #f1, "</Body></Html>"
 
 Close f1
 'open file
 ShellExecute hwnd, "Open", fStr, "", "", 0

End Sub

Private Sub mnuAnnualEvsSlope_Click()
 
 Dim i As Integer, iObj As Long, ObjIndex() As Long, a As Variant, r As Integer, c As Integer, q As Single, qs As String, v As String
 Dim nObj As Long, Obj1 As Object3DType
 Dim SlopeV As Single
 Dim DayV As Long, nDayBE1(1) As String
 Dim TimeV As Single, nTimeBE1(1) As String, TimeIsLST1 As Integer
 Dim f1 As Integer, fStr As String, AddSeasons As Integer
 
 On Error Resume Next
 CD1.Flags = CD1.Flags Or cdlOFNOverwritePrompt
 CD1.Filter = "*.Htm;*.Html|*.Htm;*.Html|All Files|*.*"
 CD1.ShowSave
 FormGr.DoPaint 'Call RefreshGDIwindow (only a fast bitblt)
 If Err Then Exit Sub 'cancel save
 
 Dim FProgress As Form
 
 ReDim ObjIndex(UserScene.NumObjects - 1)
 nObj = 0
 For iObj = 0 To UserScene.NumObjects - 1
  With UserScene.Object(iObj)
   If (.iCode And 4) Then
    ObjIndex(nObj) = iObj: nObj = nObj + 1
   End If
  End With
 Next
 
 AddSeasons = 0
 f1 = FreeFile
 fStr = CD1.FileName
 Open fStr For Output As f1
 
 If Err Then MsgBox Err.Description: Err.Clear: Close f1: Exit Sub
 
 For i = 0 To 1: nDayBE1(i) = nDayBeginEnd(i): Next
 For i = 0 To 1: nTimeBE1(i) = nTimeBeginEnd(i): Next
 TimeIsLST1 = TimeIsLST: TimeIsLST = True
 
 
 Print #f1, "<Html><Body>"
 Print #f1, "<font color=""#005000"">"
 Print #f1, "City:"; sUser.CityName; "  - Latitude:"; Format(sUser.l, "00.00"); "<br>"
 Print #f1, "</font>"
 Print #f1, "<br>Units:   Slope (Degree)       ,      E (Gj/m2) : Total Annual Solar Energy<br>"
 Print #f1, "<font color=""#440000"">"
 
 For iObj = 0 To nObj - 1
  With UserScene.Object(ObjIndex(iObj))
   Obj1 = UserScene.Object(ObjIndex(iObj))
   Print #f1, "<br>"
   Print #f1, "Object:"; .NameStr; "<br>"
   Print #f1, "<Table border=1 cellpadding=2 cellspacing=3>"
   
   Print #f1, "<TR style='Color:#000070;'><TD>Slope</TD>";
   Print #f1, "<TD>E Tot</TD>";
   If AddSeasons Then Print #f1, "<TD>E Spring</TD> <TD>E Summer</TD> <TD>E Autumn</TD> <TD>E Winter</TD>";
   Print #f1, "</TR>"
   'Print #f1, "<TD>E Tot</TD> <TD>E Summer</TD> <TD>E Winter</TD> </TR>"
   Set FProgress = New FormProgress
   FProgress.SetProperties "Create Slope Table For '" + .NameStr + "'"
   
   FProgress.Show
   FProgress.Refresh
   UserDoEventsOnPlate = 1
   
   For SlopeV = 0 To 90 Step 5
    RotateSelfObject3D UserScene.Object(ObjIndex(iObj)), SinCos3Angles(0, SlopeV, 0)
    nTimeBeginEnd(0) = "SR": nTimeBeginEnd(1) = "SN" 'IF CHANGE SR,SN Must Change E Field (we multiply TimeTot by 2 because of half day calculation)
    FProgress.SetProgress (SlopeV / 90)
    DoEvents
    If FProgress.CancelPressed Then
      UserScene.Object(ObjIndex(iObj)) = Obj1 'Cancel Rotations for next rotation
      Exit For
    End If
    
    Print #f1, "<TR>"
    Print #f1, "<TD>"; Format(SlopeV, "0"); "</TD>"
    
    'TimeTotal (Minutes) = .SpValues(7) , * 2 for calculate for whole day * UserDayStep  => 2*60(min=>sec)*1E-9(Giga) * UserDayStep(we want whole year energy!)
    'Total Annual
    nDayBeginEnd(0) = 1:  nDayBeginEnd(1) = 365
    UserFindRadiationObject ObjIndex(iObj), UserdMinutes, ComponentG, 0, UserTraceAsCollector, -1       'ColorMode=-1 is for dont set color of pages
    Print #f1, "<TD>"; Format(.SpValues(7) * 0.002 * 0.00006 * .SpValues(3) * UserDayStep, "0.000"); "</TD>";
    
    If AddSeasons Then
    'Spring
    nDayBeginEnd(0) = DayByDate(1, "4", False):  nDayBeginEnd(1) = DayByDate(1, "7", False) - 1
    UserFindRadiationObject ObjIndex(iObj), UserdMinutes, ComponentG, 0, UserTraceAsCollector, -1       'ColorMode=-1 is for dont set color of pages
    Print #f1, "<TD>"; Format(.SpValues(7) * 0.002 * 0.00006 * .SpValues(3) * UserDayStep, "0.000"); "</TD>";
    'Summer
    nDayBeginEnd(0) = DayByDate(1, "7", False):   nDayBeginEnd(1) = DayByDate(1, "10", False) - 1
    UserFindRadiationObject ObjIndex(iObj), UserdMinutes, ComponentG, 0, UserTraceAsCollector, -1       'ColorMode=-1 is for dont set color of pages
    Print #f1, "<TD>"; Format(.SpValues(7) * 0.002 * 0.00006 * .SpValues(3) * UserDayStep, "0.000"); "</TD>";
    'Autumn
    nDayBeginEnd(0) = DayByDate(1, "10", False):  nDayBeginEnd(1) = 365 ' DayByDate(1, "4", False) - 1
    UserFindRadiationObject ObjIndex(iObj), UserdMinutes, ComponentG, 0, UserTraceAsCollector, -1       'ColorMode=-1 is for dont set color of pages
    Print #f1, "<TD>"; Format(.SpValues(7) * 0.002 * 0.00006 * .SpValues(3) * UserDayStep, "0.000"); "</TD>";
    'Winter
    nDayBeginEnd(0) = 1:  nDayBeginEnd(1) = DayByDate(1, "4", False) - 1
    UserFindRadiationObject ObjIndex(iObj), UserdMinutes, ComponentG, 0, UserTraceAsCollector, -1        'ColorMode=-1 is for dont set color of pages
    Print #f1, "<TD>"; Format(.SpValues(7) * 0.002 * 0.00006 * .SpValues(3) * UserDayStep, "0.000"); "</TD>";
    End If
    
    Print #f1, "</TR>"
    UserScene.Object(ObjIndex(iObj)) = Obj1 'Cancel Rotations for next rotation
    
   Next 'SlopeV
   
   Print #f1, "</Table>"
   
  End With '..object(i..)
  Unload FProgress
  Set FProgress = Nothing
  
 Next 'iObj
 
 For i = 0 To 1:  nDayBeginEnd(i) = nDayBE1(i): Next
 For i = 0 To 1:  nTimeBeginEnd(i) = nTimeBE1(i): Next
 TimeIsLST = TimeIsLST1
 UserDoEventsOnPlate = 0
  
 Print #f1, "</font>"
 Print #f1, "</Body></Html>"
 
 Close f1
 Call DoPaint
 
 'open file:
 'ShellExecute hwnd, "Open", fStr, "", "", 0
 
End Sub

Private Sub mnuAnnualEvsLatitude_Click()
 
 Dim i As Integer, iObj As Long, ObjIndex() As Long, a As Variant, r As Integer, c As Integer, q As Single, qs As String, v As String
 Dim nObj As Long
 Dim LatitudeV As Single
 Dim DayV As Long, nDayBE1(1) As String
 Dim TimeV As Single, nTimeBE1(1) As String, TimeIsLST1 As Integer
 Dim Latitude1 As Single, city1 As String, useSt1 As Integer, CF1 As Single
 Dim f1 As Integer, fStr As String, AddSeasons As Integer
 
 On Error Resume Next
 CD1.Flags = CD1.Flags Or cdlOFNOverwritePrompt
 CD1.Filter = "*.Htm;*.Html|*.Htm;*.Html|All Files|*.*"
 CD1.ShowSave
 FormGr.DoPaint 'Call RefreshGDIwindow (only a fast bitblt)
 If Err Then Exit Sub 'cancel save
 
 Dim FProgress As Form
 
 f1 = FreeFile
 fStr = CD1.FileName
 Open CD1.FileName For Output As f1
 
 If Err Then MsgBox Err.Description: Call DoPaint: Err.Clear: Close f1: Exit Sub
 
 For i = 0 To 1: nDayBE1(i) = nDayBeginEnd(i): Next
 For i = 0 To 1: nTimeBE1(i) = nTimeBeginEnd(i): Next
 Latitude1 = sUser.l: city1 = sUser.CityName: useSt1 = UseStatisticDataForCF: CF1 = FixedCF
 
 TimeIsLST1 = TimeIsLST: TimeIsLST = True
 sUser.CityName = "ArbitraryLocation"
 UseStatisticDataForCF = 0
 If useSt1 Then 'must get CF from user because he want we use city data and now is invalid
  FixedCF = Val(InputBox("Please enter a value for CF.", "CF Value", FixedCF))
  Call DoPaint
 End If
 
 AddSeasons = 0
 
 Print #f1, "<Html><Body>"
 Print #f1, "<font color=""#005000"">"
 Print #f1, "City:"; sUser.CityName; "<br>"
 Print #f1, "<br>Total Annual Solar Energy (Gj/m2)<br>"
 Print #f1, "</font>"
 Print #f1, "<font color=""#440000"">"
 
 Print #f1, "<br>"
 Print #f1, "<Table border=1 cellpadding=2 cellspacing=3>"
 
 ReDim ObjIndex(UserScene.NumObjects - 1)
 nObj = 0
 For iObj = 0 To UserScene.NumObjects - 1
  With UserScene.Object(iObj)
   If (.iCode And 4) Then
    ObjIndex(nObj) = iObj: nObj = nObj + 1
   End If
  End With
 Next
 ReDim Preserve ObjIndex(nObj - 1) 'clear extra indexes
 
 Print #f1, "<TR> <TD>&nbsp;</TD><TD colspan="; CStr(nObj); ">Objects</TD> </TR>"
 Print #f1, "<TR> <TD>Latitude</TD>";
 For iObj = 0 To nObj - 1
  With UserScene.Object(ObjIndex(iObj))
    Print #f1, "<TD>"; .NameStr; "</TD>";
    
  End With
 Next
 
 Print #f1, " </TR>"
 Print #f1, "</font>"
  
 Set FProgress = New FormProgress
 FProgress.SetProperties "Create Latitude Table"
  
 FProgress.Show
 FProgress.Refresh
 UserDoEventsOnPlate = 1
 
 For LatitudeV = 0 To 88 Step 10
  sUser.l = LatitudeV
  FProgress.SetProgress (LatitudeV / 90)
  DoEvents
  If FProgress.CancelPressed Then Exit For
  Print #f1, "<TR>"
  Print #f1, "<TD>"; Format(LatitudeV, "0"); "</TD>"
  
  For iObj = 0 To nObj - 1
   With UserScene.Object(ObjIndex(iObj))
   
    nTimeBeginEnd(0) = "SR": nTimeBeginEnd(1) = "SN" 'IF CHANGE SR,SN Must Change E Field (we multiply TimeTot by 2 because of half day calculation)
    
    
    'TimeTotal (Minutes) = .SpValues(7) , * 2 for calculate for whole day * UserDayStep  => 2*60(min=>sec)*1E-9(Giga) * UserDayStep(we want whole year energy!)
    'Total Annual
    nDayBeginEnd(0) = 1:  nDayBeginEnd(1) = 365
    UserFindRadiationObject ObjIndex(iObj), UserdMinutes, ComponentG, 0, UserTraceAsCollector, -1       'ColorMode=-1 is for dont set color of pages
    Print #f1, "<TD>"; Format(.SpValues(7) * 0.002 * 0.00006 * .SpValues(3) * UserDayStep, "0.000"); "</TD>";
    
    If AddSeasons Then 'if add seasons
    'Spring
    nDayBeginEnd(0) = DayByDate(1, "4", False):  nDayBeginEnd(1) = DayByDate(1, "7", False) - 1
    UserFindRadiationObject ObjIndex(iObj), UserdMinutes, ComponentG, 0, UserTraceAsCollector, -1       'ColorMode=-1 is for dont set color of pages
    Print #f1, "<TD>"; Format(.SpValues(7) * 0.002 * 0.00006 * .SpValues(3) * UserDayStep, "0.000"); "</TD>";
    'Summer
    nDayBeginEnd(0) = DayByDate(1, "7", False):   nDayBeginEnd(1) = DayByDate(1, "10", False) - 1
    UserFindRadiationObject ObjIndex(iObj), UserdMinutes, ComponentG, 0, UserTraceAsCollector, -1       'ColorMode=-1 is for dont set color of pages
    Print #f1, "<TD>"; Format(.SpValues(7) * 0.002 * 0.00006 * .SpValues(3) * UserDayStep, "0.000"); "</TD>";
    'Autumn
    nDayBeginEnd(0) = DayByDate(1, "10", False):  nDayBeginEnd(1) = 365 ' DayByDate(1, "4", False) - 1
    UserFindRadiationObject ObjIndex(iObj), UserdMinutes, ComponentG, 0, UserTraceAsCollector, -1       'ColorMode=-1 is for dont set color of pages
    Print #f1, "<TD>"; Format(.SpValues(7) * 0.002 * 0.00006 * .SpValues(3) * UserDayStep, "0.000"); "</TD>";
    'Winter
    nDayBeginEnd(0) = 1:  nDayBeginEnd(1) = DayByDate(1, "4", False) - 1
    UserFindRadiationObject ObjIndex(iObj), UserdMinutes, ComponentG, 0, UserTraceAsCollector, -1        'ColorMode=-1 is for dont set color of pages
    Print #f1, "<TD>"; Format(.SpValues(7) * 0.002 * 0.00006 * .SpValues(3) * UserDayStep, "0.000"); "</TD>";
    End If 'if season
    
    Print #f1, "</TR>"
    
   End With '..object(i..)
   
  Next 'iObj
  
 Next 'LatitudeV
 
 Unload FProgress
 Set FProgress = Nothing
 Call DoPaint
 
 Print #f1, "</Table>"
 
 For i = 0 To 1:  nDayBeginEnd(i) = nDayBE1(i): Next
 For i = 0 To 1:  nTimeBeginEnd(i) = nTimeBE1(i): Next
 TimeIsLST = TimeIsLST1
 sUser.l = Latitude1: sUser.CityName = city1: UseStatisticDataForCF = useSt1: FixedCF = CF1

 UserDoEventsOnPlate = 0
  
 Print #f1, "</font>"
 Print #f1, "</Body></Html>"
 
 Close f1
 
 'open file:
 'ShellExecute hwnd, "Open", fStr, "", "", 0

End Sub

