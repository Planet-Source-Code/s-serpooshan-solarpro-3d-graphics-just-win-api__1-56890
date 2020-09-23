VERSION 5.00
Begin VB.Form FormTest 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11145
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   11145
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   5295
      Left            =   2460
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "FormTest.frx":0000
      Top             =   660
      Visible         =   0   'False
      Width           =   4515
   End
End
Attribute VB_Name = "FormTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyPress(KeyAscii As Integer)
 If KeyAscii = 27 Then End
End Sub

Private Sub Form_Load()
 
 'Call VFtest: End
 
 Dim GDirMin, GDirMax
 Dim n As Long
'On Error Resume Next
 n = 13
 TimeIsLST = True
 UseStatisticDataForCF = False
 sUser.l = 40
 sUser.Altitude = 1650
 sUser.CityName = "ArbitraryLocation"
 
 Show
 
 
 For FixedCF = 0 To 0 Step 0.1
 'n = n + 35: If n > 365 Then n = 1 '360 * Rnd + 1
 Print n
 MethodOfSolarCalc = 2
 SetDayOfYear n, sUser
 'SetabkForHottel sUser
 nTimeBeginEnd(0) = "9:30": nTimeBeginEnd(1) = nTimeBeginEnd(0)
 nDayBeginEnd(0) = n: nDayBeginEnd(1) = n

 MethodOfSolarCalc = 3
 SetDayOfYear n, sUser
 UserFindRadiationPlate Vector3D(0, 0, 1), 0
 With GRadiationUser
   Print "           Hottel:   GDirect="; Format(.Direct, "0.00"); "   Gdiffuse="; Format(.diffuse, ".00")
 End With
 
 Next
 
 
 Exit Sub
 
 
 MethodOfSolarCalc = 1
 TimeIsLST = True
 UseStatisticDataForCF = True
 sUser.CityName = "Shiraz"
 SetCityData sUser
 
 GDirMin = 1000: GDirMax = -1
 
 For n = 1 To 356
  SetDayOfYear n, sUser
  nTimeBeginEnd(0) = "12:00": nTimeBeginEnd(1) = "12:00"
  nDayBeginEnd(0) = n: nDayBeginEnd(1) = n
  UserFindRadiationPlate Vector3D(0, 0, 1), 0
  G = GRadiationUser.Direct
  If G < GDirMin Then GDirMin = G: nmin = n
  If G > GDirMax Then GDirMax = G: nmax = n
  
  a = a + CStr(n) + ": " + CStr(G) + "  ,  " + CStr(GRadiationUser.diffuse) + vbCrLf
 Next
 Print GDirMax; nmax
 Print GDirMin; nmin
 Text1.Text = a
 
End Sub

Sub VFtest()

Dim p As Single, q As Single, nx As Long, ny As Long
Dim dx As Single, dy As Single, ix As Long, iy As Long, c As Long
Dim x As Single, y As Single
Dim Obj1 As Object3DType, Obj2 As Object3DType
p = 200
q = 2
nx = 20
ny = 20
Obj1 = MakePageObjectnXnY(p, q, nx, ny)
Obj2 = TransferObject3D(Obj1, Vector3D(-p, 0, 0))
RotateSelfObject3D Obj2, SinCos3Angles(0, 180, 0)  'pi-alfa => alfa=angle between surfaces
TransferSelfObject3D Obj2, Vector3D(0, 0, 10)
F12 = ViewFactor12(Obj1, Obj2)
MsgBox CStr(F12)
  
End Sub



