VERSION 5.00
Begin VB.Form FormViewSettings 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Settings"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3435
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormViewSettings.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   3435
   Begin VB.CommandButton CommandLoadViewSetting 
      Caption         =   "Load.."
      Height          =   375
      Left            =   2280
      TabIndex        =   13
      Top             =   1380
      Width           =   1035
   End
   Begin VB.CommandButton CommandSaveSetting 
      Caption         =   "Save.."
      Height          =   375
      Left            =   2280
      TabIndex        =   12
      Top             =   960
      Width           =   1035
   End
   Begin VB.CommandButton CommandExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   375
      Left            =   2280
      TabIndex        =   11
      Top             =   540
      Width           =   1035
   End
   Begin VB.CommandButton CommandApply 
      Caption         =   "Apply"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   120
      Width           =   1035
   End
   Begin VB.TextBox TextVS 
      Height          =   285
      Index           =   4
      Left            =   1260
      TabIndex        =   9
      Top             =   1380
      Width           =   855
   End
   Begin VB.TextBox TextVS 
      Height          =   285
      Index           =   3
      Left            =   1260
      TabIndex        =   8
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox TextVS 
      Height          =   285
      Index           =   2
      Left            =   1260
      TabIndex        =   7
      Top             =   780
      Width           =   855
   End
   Begin VB.TextBox TextVS 
      Height          =   285
      Index           =   1
      Left            =   1260
      TabIndex        =   6
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox TextVS 
      Height          =   285
      Index           =   0
      Left            =   1260
      TabIndex        =   5
      Top             =   180
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Y-Shift(2D)"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1425
      Width           =   795
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "X-Shift(2D)"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1125
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Scale (0-10)"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   825
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fi (Above)"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   525
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Teta (Around)"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   225
      Width           =   1020
   End
End
Attribute VB_Name = "FormViewSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Commandexit_Click()
 Unload Me
End Sub

Private Sub CommandApply_Click()
 If FormGrIsVisible Then FormGr.DoViewNumericSettings
End Sub

Private Sub CommandLoadViewSetting_Click()
 Dim f As Long, f1 As String, a As String, i As Long
 On Error Resume Next
 With FormGr.CD1
  f1 = .Filter
  .Filter = "*.txt|*.txt"
  .ShowOpen
  .Filter = f1
  FormGr.DoPaint 'Call RefreshGDIwindow (only a fast bitblt)
  If Err Then Exit Sub
  f = FreeFile
  Open .FileName For Input As f
  For i = 0 To 4
   Input #f, a
   TextVS(i).Text = a
  Next
  Close f
 End With
 CommandApply_Click
End Sub

Private Sub CommandSaveSetting_Click()
 Dim f As Long, f1 As String, i As Long
 On Error Resume Next
 With FormGr.CD1
  f1 = .Filter
  .Filter = "*.txt|*.txt"
  .ShowSave
  FormGr.DoPaint 'Call RefreshGDIwindow (only a fast bitblt)
  .Filter = f1
  If Err Then Exit Sub
  f = FreeFile
  Open .FileName For Output As f
  For i = 0 To 4
   Print #f, TextVS(i).Text; ", ";
  Next
  Close f
 End With
End Sub

Private Sub Form_Load()
 Me.Move FormGr.Left, FormGr.Top
End Sub
