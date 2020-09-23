VERSION 5.00
Begin VB.Form FormTransforms 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transforms..."
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3825
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   3825
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CommandCancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   2580
      TabIndex        =   8
      Top             =   900
      Width           =   975
   End
   Begin VB.CommandButton CommandOK 
      Caption         =   "Apply"
      Default         =   -1  'True
      Height          =   375
      Left            =   2580
      TabIndex        =   7
      Top             =   420
      Width           =   975
   End
   Begin VB.CommandButton CommandDefault 
      Caption         =   "Defaults"
      Height          =   375
      Left            =   2580
      TabIndex        =   6
      Top             =   1950
      Width           =   975
   End
   Begin VB.TextBox TextScale 
      Height          =   285
      Left            =   360
      TabIndex        =   5
      Text            =   "100"
      Top             =   2040
      Width           =   1395
   End
   Begin VB.TextBox TextRotation 
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Text            =   "0,0,0"
      Top             =   1260
      Width           =   1395
   End
   Begin VB.TextBox TextTransform 
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Text            =   "0,0,0"
      Top             =   480
      Width           =   1395
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Scale Percent:"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   1800
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Transform Vector:"
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Rotation (x-y-z):"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   1020
      Width           =   1215
   End
End
Attribute VB_Name = "FormTransforms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_Cancel As Integer

Private Sub CommandCancel_Click()
m_Cancel = True
Unload Me
End Sub

Private Sub CommandDefault_Click()
 Call SetAsDefault
End Sub

Sub SetAsDefault()
 TextTransform.Text = "0,0,0"
 TextRotation.Text = "0,0,0"
 TextScale.Text = "100"
End Sub

Private Sub CommandOk_Click()
 m_Cancel = 0
 
 If iObjTransform >= 0 Then
    ApplyTransformsToObj UserScene.Object(iObjTransform)
    ShouldReCalculate = True
    If TargetObjIndex > -1 Then
     SetViewShift3D UserScene.Object(TargetObjIndex).CenterPoint
    End If
    Call SetAsDefault
    If FormGrIsVisible Then Call FormGr.DrawGrShapes 'to re-calculate & redraw
 Else 'some times such as when in Create Object Form Press Tranform button
    Me.Hide
 End If
 
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
m_Cancel = True
End Sub

Public Property Get CancelPressed() As Variant
CancelPressed = m_Cancel
End Property

Public Property Let CancelPressed(ByVal vNewValue As Variant)
 m_Cancel = vNewValue
End Property

