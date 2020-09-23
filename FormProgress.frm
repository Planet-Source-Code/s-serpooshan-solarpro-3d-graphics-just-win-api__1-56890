VERSION 5.00
Begin VB.Form FormProgress 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wait..."
   ClientHeight    =   765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3765
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   51
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   251
   Tag             =   "FormProgressType"
   Begin VB.CommandButton CommandCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2910
      TabIndex        =   2
      Top             =   30
      Width           =   795
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   60
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   239
      TabIndex        =   0
      Top             =   330
      Width           =   3645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Progress:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   120
      Width           =   690
   End
End
Attribute VB_Name = "FormProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_CancelPressed As Integer
Dim m_ProgressPer As Single
Dim dx As Single, dy As Single

Sub SetProgress(ByVal PercentComp As Single)
 m_ProgressPer = PercentComp
 Call ReDrawProgress
End Sub

Sub ReDrawProgress()
 Picture1.Line (0, 0)-(dx * m_ProgressPer, dy), , B
 'Me.Refresh
 Picture1.Refresh
End Sub

Sub SetProperties(CaptionStr As String, Optional LabelStr As String = "Progress:")
 Caption = CaptionStr
 Label1.Caption = LabelStr
End Sub

Public Property Get CancelPressed() As Integer
 CancelPressed = m_CancelPressed
End Property

Public Property Let CancelPressed(ByVal vNewValue As Integer)
 m_CancelPressed = vNewValue
 CancelPressedPublic = vNewValue
End Property

Private Sub CommandCancel_Click()
 CancelPressed = True
End Sub

Private Sub Form_Load()
 Dim a As Form, Top1 As Long, t As Long, x As Long
 CancelPressedPublic = 0
 For Each a In Forms
  If a.Tag = "FormProgressType" And a.Visible = True Then
   t = a.Top + a.Height
   x = a.Left
   If t > Top1 Then Top1 = t
  End If
 Next
 If Top1 = 0 Then
  Me.Top = (Screen.Height - Me.Height) \ 2
  Me.Left = (Screen.Width - Me.Width) \ 2
 Else
  Me.Move x, Top1
 End If
End Sub

Private Sub Form_Resize()
 dx = Picture1.ScaleWidth
 dy = Picture1.ScaleHeight
 Me.Refresh
 Call ReDrawProgress
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Me.Hide
 DoEvents
 If FormGrIsVisible Then FormGr.DoPaint
End Sub
