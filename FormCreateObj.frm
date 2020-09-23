VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FormCreateObj 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Object"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8415
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   357
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   561
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CommandLoadObj 
      Caption         =   "Load"
      Height          =   375
      Left            =   1140
      TabIndex        =   37
      Top             =   4860
      Width           =   975
   End
   Begin VB.CommandButton CommandSaveObj 
      Caption         =   "Save"
      Height          =   375
      Left            =   1140
      TabIndex        =   36
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox TextNameStr 
      Height          =   315
      Left            =   120
      TabIndex        =   34
      Text            =   "UserDefined"
      Top             =   3540
      Width           =   1695
   End
   Begin VB.CheckBox CheckCalculRad 
      Caption         =   "Calculate Radiation"
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   4020
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CommandButton CommandTransform 
      Caption         =   "Transform"
      Height          =   375
      Left            =   60
      TabIndex        =   19
      Top             =   4860
      Width           =   975
   End
   Begin VB.CommandButton CommandAddToScene 
      Caption         =   "Add To Scene"
      Height          =   375
      Left            =   2220
      TabIndex        =   18
      Top             =   4440
      Width           =   1275
   End
   Begin VB.CommandButton CommandCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2220
      TabIndex        =   17
      Top             =   4860
      Width           =   1275
   End
   Begin VB.CommandButton CommandDraw 
      Caption         =   "Create"
      Default         =   -1  'True
      Height          =   375
      Left            =   60
      TabIndex        =   16
      Top             =   4440
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2895
      Left            =   60
      TabIndex        =   1
      Top             =   120
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   5106
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabHeight       =   529
      TabCaption(0)   =   "Rotate"
      TabPicture(0)   =   "FormCreateObj.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(7)=   "TextEq"
      Tab(0).Control(8)=   "TextZ1"
      Tab(0).Control(9)=   "TextZ2"
      Tab(0).Control(10)=   "TextNZ"
      Tab(0).Control(11)=   "TextTeta1"
      Tab(0).Control(12)=   "TextTeta2"
      Tab(0).Control(13)=   "TextNTeta"
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Extrude"
      TabPicture(1)   =   "FormCreateObj.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label10"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label12"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label13"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label14"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label15"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label16"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "TextZnExtrude"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "TextZlength"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "TextnX"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "TextX2"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "TextX1"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "TextEqExtrude"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).ControlCount=   12
      Begin VB.TextBox TextEqExtrude 
         Height          =   315
         Left            =   60
         TabIndex        =   25
         Text            =   "x^2 / 12"
         Top             =   720
         Width           =   3315
      End
      Begin VB.TextBox TextX1 
         Height          =   315
         Left            =   60
         TabIndex        =   24
         Text            =   "-15"
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox TextX2 
         Height          =   315
         Left            =   1200
         TabIndex        =   23
         Text            =   "15"
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox TextnX 
         Height          =   315
         Left            =   2340
         TabIndex        =   22
         Text            =   "20"
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox TextZlength 
         Height          =   315
         Left            =   60
         TabIndex        =   21
         Text            =   "50"
         Top             =   2460
         Width           =   975
      End
      Begin VB.TextBox TextZnExtrude 
         Height          =   315
         Left            =   2340
         TabIndex        =   20
         Text            =   "2"
         Top             =   2460
         Width           =   975
      End
      Begin VB.TextBox TextNTeta 
         Height          =   315
         Left            =   -72660
         TabIndex        =   15
         Text            =   "40"
         Top             =   2460
         Width           =   975
      End
      Begin VB.TextBox TextTeta2 
         Height          =   315
         Left            =   -73800
         TabIndex        =   13
         Text            =   "6.283185"
         Top             =   2460
         Width           =   975
      End
      Begin VB.TextBox TextTeta1 
         Height          =   315
         Left            =   -74940
         TabIndex        =   11
         Text            =   "0"
         Top             =   2460
         Width           =   975
      End
      Begin VB.TextBox TextNZ 
         Height          =   315
         Left            =   -72660
         TabIndex        =   9
         Text            =   "5"
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox TextZ2 
         Height          =   315
         Left            =   -73800
         TabIndex        =   7
         Text            =   "50"
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox TextZ1 
         Height          =   315
         Left            =   -74940
         TabIndex        =   5
         Text            =   "0"
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox TextEq 
         Height          =   315
         Left            =   -74940
         TabIndex        =   2
         Text            =   "z/2"
         Top             =   720
         Width           =   3315
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Equation Of Curve: y = f (x)"
         Height          =   195
         Left            =   60
         TabIndex        =   31
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "x1"
         Height          =   195
         Left            =   60
         TabIndex        =   30
         Top             =   1320
         Width           =   180
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "x2"
         Height          =   195
         Left            =   1200
         TabIndex        =   29
         Top             =   1320
         Width           =   180
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "devision (N)"
         Height          =   195
         Left            =   2340
         TabIndex        =   28
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "z Length"
         Height          =   195
         Left            =   60
         TabIndex        =   27
         Top             =   2220
         Width           =   615
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "devision (N)"
         Height          =   195
         Left            =   2340
         TabIndex        =   26
         Top             =   2220
         Width           =   855
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "devision (N)"
         Height          =   195
         Left            =   -72660
         TabIndex        =   14
         Top             =   2220
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Teta2"
         Height          =   195
         Left            =   -73800
         TabIndex        =   12
         Top             =   2220
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Teta1"
         Height          =   195
         Left            =   -74940
         TabIndex        =   10
         Top             =   2220
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "devision (N)"
         Height          =   195
         Left            =   -72660
         TabIndex        =   8
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Upper Z"
         Height          =   195
         Left            =   -73800
         TabIndex        =   6
         Top             =   1320
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Lower Z"
         Height          =   195
         Left            =   -74940
         TabIndex        =   4
         Top             =   1320
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Equation Of Curve: R = f (Z)"
         Height          =   195
         Left            =   -74940
         TabIndex        =   3
         Top             =   480
         Width           =   2070
      End
   End
   Begin VB.PictureBox PicPrev 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00965A3C&
      ForeColor       =   &H80000008&
      Height          =   5235
      Left            =   3540
      ScaleHeight     =   347
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   319
      TabIndex        =   0
      Top             =   60
      Width           =   4815
   End
   Begin VB.CheckBox CheckCauseShadow 
      Caption         =   "Cause Shadow"
      Height          =   255
      Left            =   2160
      TabIndex        =   33
      Top             =   4020
      Width           =   1395
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Left            =   120
      TabIndex        =   35
      Top             =   3300
      Width           =   405
   End
End
Attribute VB_Name = "FormCreateObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ViewTeta, ViewFi, ViewScale

Private Sub CommandCancel_Click()
 ReleaseGDI
 Me.Hide
End Sub

Private Sub CommandDraw_Click()
Select Case SSTab1.Tab
 Case 0 'Rotate
  TextEq.Text = Replace(Replace(LCase(TextEq.Text), " ", ""), "r=", "")
  MakeObjByRotateCurve UDefObj, TextEq.Text, Val(TextZ1.Text), Val(TextZ2.Text), Val(TextNZ.Text), Val(TextTeta1.Text), Val(TextTeta2.Text), Val(TextNTeta.Text), 4, TextNameStr.Text, , &H999999
 Case 1 'Extrude
  TextEq.Text = Replace(Replace(LCase(TextEq.Text), " ", ""), "y=", "")
  MakeObjByExtrude UDefObj, TextEqExtrude.Text, Val(TextX1), Val(TextX2), Val(TextnX.Text), Val(TextZlength), Val(TextZnExtrude), 4, TextNameStr.Text, , &H999999, 0
 End Select
 Call DrawCrObjShapes
 
End Sub

Sub DrawCrObjShapes()
 Dim v As Vector3DType, Azy As SinCos3AngType, dr1 As Integer, VnFlesh As Page3DType
 
 Call FillHwndOut
 
 SelectQbPenBrush 15, 16
 DrawObject3D UDefObj
 
 With UDefObj
  If .NumPages > 0 Then
    v = VnOfPage(.Page(0))
    Azy = GetSinCosForRotation_X_To_Vector(v)
    VnFlesh = MakeXFlesh(6, 2, RGB(150, 0, 0))
    VnFlesh = RotatePage3D(VnFlesh, Azy)
    TransferSelfPage3D VnFlesh, VectorOfPoint(.Page(0).Point(0))
    'DrawLine3D Line3D(.Page(0).Point(0), TransferPoint3D(.Page(0).Point(0), v))
    DrawPage3D VnFlesh, RGB(50, 0, 0)
  End If
 End With
 
 'If ShouldDrawAxises Then
  SelectQbPenBrush 0, 16, 0
  For i = 0 To 2
   DrawPage3D FleshAxis(i), -2
  Next
  
  SelectQbPenBrush 15, 16, 0
  DrawPolyDraw3D TextXdraw
  SelectQbPenBrush 7, 16, 0
  DrawPolyDraw3D TextYdraw
 'End If

 RefreshGDIwindow
 
End Sub

Private Sub CommandAddToScene_Click()
 Dim c As Long
 c = 0
 If CheckCalculRad.Value Then c = c + 4
 If CheckCauseShadow.Value Then c = c + 8
 
 UDefObj.ColEdge = -1
 UDefObj.iCode = c
 UDefObj.NameStr = TextNameStr.Text
 AddObjectToScene UserScene, UDefObj
 Hide
 Call FormWizard.Form_Activate 'Enable UserDefined Object Check Box
End Sub

Private Sub CommandLoadObj_Click()
 On Error Resume Next
 With MDIFormMain.CD1
  .Flags = .Flags Or cdlOFNFileMustExist Or cdlOFNOverwritePrompt
  .Filter = "*.Obj|*.Obj|All Files|*.*"
  .ShowOpen
  If Err Then Exit Sub 'cancel
  LoadObjectFromFile .FileName, UDefObj
 End With
 Call DrawCrObjShapes
 
End Sub

Private Sub CommandSaveObj_Click()
 On Error Resume Next
 With MDIFormMain.CD1
  .Flags = .Flags Or cdlOFNFileMustExist Or cdlOFNOverwritePrompt
  .Filter = "*.Obj|*.Obj|All Files|*.*"
  .ShowSave
  If Err Then Exit Sub 'cancel
  SaveObjectToFile .FileName, UDefObj
 End With
 
End Sub

Private Sub CommandTransform_Click()
 On Error Resume Next
 FormTransforms.Show 1
 ShouldRefreshFormGr = True
 If FormTransforms.CancelPressed Then Exit Sub
 ApplyTransformsToObj UDefObj
 FormTransforms.SetAsDefault
 Call DrawCrObjShapes

End Sub

Private Sub Form_Activate()
 ReleaseGDI
 With PicPrev
  If gdi_BackgrFillCol < 0 Then gdi_BackgrFillCol = GetRealNearestColor(.hdc, .BackColor)
  InitializeGDI PicPrev
  SetViewShift2D .ScaleWidth \ 2, .ScaleHeight * 0.6
  SetViewShift3D Point3D(0, 0, 0)
 End With
 If ViewScale = 0 Then
  ViewScale = 4: ViewTeta = 45: ViewFi = 45
 End If
 SetViewScale3D ViewScale
 SetView3D ViewTeta, ViewFi
 Call DrawCrObjShapes
 DoEvents
 RefreshGDIwindow
 
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 ReleaseGDI

End Sub

Private Sub PicPrev_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Static xm1 As Long, ym1 As Long

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
 Call DrawCrObjShapes
End If

End Sub

