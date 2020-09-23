VERSION 5.00
Begin VB.Form FormVFGetObjects 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View Factor"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4185
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
   ScaleHeight     =   3195
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CommandCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2145
      TabIndex        =   6
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton CommandCalcul 
      Caption         =   "Calculate  F1-2"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   645
      TabIndex        =   5
      Top             =   2400
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   645
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1740
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   645
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "«Calculate 'View Factor' of two objects»"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   3705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Secound Object:"
      Height          =   195
      Left            =   645
      TabIndex        =   1
      Top             =   1500
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "First Object:"
      Height          =   195
      Left            =   645
      TabIndex        =   0
      Top             =   720
      Width           =   900
   End
End
Attribute VB_Name = "FormVFGetObjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ObjIndexes() As Long

Private Sub CommandCalcul_Click()
 Dim F12 As Single
 If Combo1.ListIndex = -1 Or Combo2.ListIndex = -1 Or Combo1.ListIndex = Combo2.ListIndex Then
  MsgBox "Please specify two different objects to proceed !", , "Error": Exit Sub
 End If
 Me.Hide
 With UserScene
  F12 = ViewFactor12(.Object(ObjIndexes(Combo1.ListIndex)), .Object(ObjIndexes(Combo2.ListIndex)))
 End With
 MsgBox "       F = " + CStr(F12) + Space(20) + vbCrLf, , "View Factor"
End Sub

Private Sub CommandCancel_Click()
 Me.Hide
 DoEvents
 If FormGrIsVisible Then FormGr.DoPaint
End Sub

Private Sub Form_Activate()
 Dim i As Long, c As Long
 Combo1.Clear
 Combo2.Clear
 c = 0
 With UserScene
  ReDim ObjIndexes(.NumObjects - 1)
  For i = 0 To .NumObjects - 1
   If .Object(i).NameStr <> DefaultHorizontalPlateStr Then
    Combo1.AddItem .Object(i).NameStr
    Combo2.AddItem .Object(i).NameStr
    ObjIndexes(c) = i
    c = c + 1
   End If
  Next
 End With
 If c <= 1 Then
  MsgBox "There is only " + CStr(c) + " object in the scene!" + vbCrLf + vbCrLf + "We need two objects at least!"
  CommandCancel_Click
 End If
End Sub

