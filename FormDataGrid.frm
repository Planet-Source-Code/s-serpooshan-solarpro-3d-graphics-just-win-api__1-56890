VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FormDataGrid 
   AutoRedraw      =   -1  'True
   Caption         =   "Data"
   ClientHeight    =   1230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6975
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
   MDIChild        =   -1  'True
   ScaleHeight     =   82
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   465
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   1200
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6960
      _ExtentX        =   12277
      _ExtentY        =   2117
      _Version        =   393216
      Rows            =   5
      Cols            =   10
      ForeColorFixed  =   -2147483640
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483630
      BackColorBkg    =   -2147483633
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
   End
End
Attribute VB_Name = "FormDataGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
 On Error Resume Next
 Dim k As Long, n As Long, i As Integer, a As Variant, b As Variant
 k = Screen.TwipsPerPixelX

 With FormDataGrid.Grid1
  a = Array("Object", "Area  (m2)", "Gmin (w/m2)", "Gmax (w/m2)", "Gave (w/m2)", "q  (kWatt)", "qDir. (%)", "qdif. (%)", "qRef. (%)", "E (kW.hr)")
  'b = Array("Unit", "m2", "Watt/m2", "Watt/m2", "Watt/m2", "KWatt", "%", "%", "%", "Kjule")
  n = UBound(a) + 1
  .Cols = n
  For i = 0 To n - 1
   .ColWidth(i) = 72 * k 'k / n
   .TextMatrix(0, i) = a(i)
   '.TextMatrix(1, i) = b(i)
   .ColAlignment(i) = 1
  Next
  .Col = 0: .Row = 0:  .CellAlignment = 4
  .ColWidth(0) = 80 * k
  '.ColAlignment(0) = 1: .ColAlignment(1) = 1: .ColAlignment(2) = 4
  '.Col = 1: .Row = 0: .CellAlignment = 4
  .Col = 1: .Row = 1
 End With
 Top = 0
 Call FormGr.SetDataGridValues
 
End Sub

Private Sub Form_Resize()
 On Error Resume Next
 If ScaleWidth > 80 And Grid1.Width <> ScaleWidth Then Grid1.Width = ScaleWidth
 Height = (Grid1.Height - ScaleHeight) * Screen.TwipsPerPixelY + Height
 'If ScaleHeight > 80 And Grid1.Height <> ScaleHeight Then
 ' Grid1.Height = ScaleHeight
 'End If
End Sub

Private Sub Grid1_LostFocus()
With Grid1
 .BackColorSel = &H80000005
 .ForeColorSel = &H80000012
End With
End Sub

Private Sub Grid1_SelChange()
With Grid1
 .BackColorSel = &H880000
 .ForeColorSel = &HFFFFFF
End With
End Sub
