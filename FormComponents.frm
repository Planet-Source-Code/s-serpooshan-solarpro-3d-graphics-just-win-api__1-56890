VERSION 5.00
Begin VB.Form FormComponents 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Components"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   4680
   Begin VB.CommandButton CommandCancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   3158
      TabIndex        =   5
      Top             =   1388
      Width           =   1005
   End
   Begin VB.CheckBox CheckComponent 
      Caption         =   "Ground Reflect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   578
      TabIndex        =   3
      ToolTipText     =   "Ground Reflect Component"
      Top             =   1508
      Value           =   1  'Checked
      Width           =   1425
   End
   Begin VB.CheckBox CheckComponent 
      Caption         =   "diffuse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   578
      TabIndex        =   2
      ToolTipText     =   "diffuse Componet"
      Top             =   1178
      Value           =   1  'Checked
      Width           =   825
   End
   Begin VB.CheckBox CheckComponent 
      Caption         =   "Direct"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   578
      TabIndex        =   1
      ToolTipText     =   "Direct Componet"
      Top             =   848
      Value           =   1  'Checked
      Width           =   825
   End
   Begin VB.CommandButton CommandOk 
      Caption         =   "Calculate"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3158
      MouseIcon       =   "FormComponents.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   788
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Consider these Components:"
      Height          =   195
      Left            =   518
      TabIndex        =   4
      Top             =   368
      Width           =   2085
   End
End
Attribute VB_Name = "FormComponents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandCancel_Click()
 Unload Me
End Sub

Private Sub CommandOk_Click()
 Call FormGr.SetComponentG
End Sub

Private Sub Form_Load()
 Dim i As Integer
 For i = 0 To 2
  CheckComponent(i).Value = IIf(ComponentG And 2 ^ i, 1, 0)
 Next
End Sub
