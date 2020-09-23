VERSION 5.00
Begin VB.Form FormSelectObject 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Object"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3645
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
   ScaleHeight     =   2430
   ScaleWidth      =   3645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   405
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   960
      Width           =   2775
   End
   Begin VB.CommandButton CommandOk 
      Caption         =   "OK"
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
      Left            =   405
      TabIndex        =   1
      Top             =   1620
      Width           =   1275
   End
   Begin VB.CommandButton CommandCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   1620
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Object:"
      Height          =   195
      Left            =   405
      TabIndex        =   4
      Top             =   720
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Please select an object from the list:"
      Height          =   195
      Left            =   405
      TabIndex        =   3
      Top             =   360
      Width           =   2610
   End
End
Attribute VB_Name = "FormSelectObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ObjIndexes() As Long

Private Sub CommandOk_Click()
 If Combo1.ListIndex = -1 Then
  MsgBox "Please specify an object to proceed !", , "Error": Exit Sub
 End If
 Me.Hide
End Sub

Private Sub CommandCancel_Click()
 Me.Hide
 'Error 1 'to say caller that user press cancel! (such as ComonDialog control)
 Combo1.ListIndex = 0: ObjIndexes(Combo1.ListIndex) = -1 'means cancel
End Sub

Private Sub Form_Activate()
 Dim i As Long, s As String, c As Long
 Combo1.Clear
 With UserScene
  ReDim ObjIndexes(.NumObjects - 1)
  For i = 0 To .NumObjects - 1
   s = .Object(i).NameStr
   If s <> DefaultHorizontalPlateStr Then Combo1.AddItem s: ObjIndexes(c) = i: c = c + 1
  Next
 End With
 If Combo1.ListCount = 1 Then Combo1.ListIndex = 0: Call CommandOk_Click  'we have not any choise

End Sub

Public Property Get SelectedIndex() As Variant
 SelectedIndex = ObjIndexes(Combo1.ListIndex)
End Property

Public Property Let SelectedIndex(ByVal vNewValue As Variant)
End Property
