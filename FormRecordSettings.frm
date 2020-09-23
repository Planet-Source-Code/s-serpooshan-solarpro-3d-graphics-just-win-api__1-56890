VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FormRecordSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Record Settings"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4410
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormRecordSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   4410
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TextPrefix 
      Height          =   285
      Left            =   1380
      TabIndex        =   11
      Top             =   1500
      Width           =   795
   End
   Begin VB.CommandButton CommandOkCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   2340
      TabIndex        =   9
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton CommandOkCancel 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   8
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox TextHeaderText 
      Height          =   315
      Left            =   1680
      TabIndex        =   7
      Top             =   2580
      Width           =   2535
   End
   Begin VB.CheckBox CheckAddHeader 
      Caption         =   "Add header text"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2610
      Width           =   1515
   End
   Begin VB.TextBox TextFirstIndex 
      Height          =   285
      Left            =   2880
      TabIndex        =   4
      Top             =   1920
      Width           =   795
   End
   Begin VB.CommandButton CommandBrowseFolder 
      Caption         =   "..."
      Height          =   315
      Left            =   3840
      TabIndex        =   2
      Top             =   900
      Width           =   435
   End
   Begin VB.TextBox TextOutputFolder 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   900
      Width           =   3615
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   3120
      Top             =   660
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label Label4 
      Caption         =   "File name prefix:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1515
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "» Record graphical output images to disk"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   180
      Width           =   3555
   End
   Begin VB.Label Label2 
      Caption         =   "First saving file 'Index' to begin from:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1935
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Path/Folder for output images:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   660
      Width           =   2295
   End
End
Attribute VB_Name = "FormRecordSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_IsCanceled As Boolean

Private Sub CheckAddHeader_Click()
TextHeaderText.Enabled = CheckAddHeader.Value
End Sub

Private Sub CommandBrowseFolder_Click()
Dim a As String, i As Long, b As String
On Error Resume Next
CD1.Filter = "*.bmp|*.bmp"
CD1.DialogTitle = "Save folder & file name prefix"
CD1.ShowSave
CD1.DialogTitle = ""
If Err Then Exit Sub
b = CD1.FileTitle
a = CD1.FileName
a = Left$(a, Len(a) - Len(b))
If Left$(a, 1) <> "\" Then a = a + "\"
TextOutputFolder.Text = a
TextOutputFolder.SelStart = Len(a): TextOutputFolder.SelLength = 0
i = InStrRev(b, ".")
If i Then b = Left$(b, i - 1)
TextPrefix.Text = b
End Sub

Public Property Get IsCanceled() As Boolean
IsCanceled = m_IsCanceled
End Property

Public Property Let IsCanceled(ByVal vNewValue As Boolean)
m_IsCanceled = vNewValue
End Property

Private Sub CommandOkCancel_Click(Index As Integer)
If Index = 0 Then 'OK
    m_IsCanceled = False
Else 'Cancel
    m_IsCanceled = True
End If
Me.Hide
End Sub

Private Sub Form_Activate()
Call CheckAddHeader_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbFormControlMenu Then 'user close the form
    Cancel = True 'don't unload the form
    CommandOkCancel_Click (1) 'same as cancel button
End If
End Sub
