VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form FormScriptEditor 
   Caption         =   "Script Editor"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7995
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   495
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   533
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   5400
      Top             =   2100
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
   End
   Begin VB.ListBox ListSpKeyWords 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      ItemData        =   "FormEditor.frx":0000
      Left            =   2370
      List            =   "FormEditor.frx":0034
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.ComboBox ComboSpKeyWords 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      ItemData        =   "FormEditor.frx":00DE
      Left            =   2370
      List            =   "FormEditor.frx":0112
      TabIndex        =   4
      ToolTipText     =   "KeyWords"
      Top             =   0
      Width           =   2160
   End
   Begin VB.CommandButton CommandOpen 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   30
      Picture         =   "FormEditor.frx":01BC
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Open"
      Top             =   0
      Width           =   345
   End
   Begin VB.CommandButton CommandSave 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   420
      Picture         =   "FormEditor.frx":0672
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Save"
      Top             =   0
      Width           =   345
   End
   Begin VB.CommandButton CommandRunScript 
      Height          =   315
      Left            =   810
      Picture         =   "FormEditor.frx":0B28
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Run (F5)"
      Top             =   0
      Width           =   345
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   330
      Width           =   7965
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ctrl+Space:"
      Height          =   195
      Left            =   1530
      TabIndex        =   6
      Top             =   60
      Width           =   825
   End
End
Attribute VB_Name = "FormScriptEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Dim DontInsertText As Integer

Sub InsertText(t As String)
 Dim t1 As String
 With Text1
  .SetFocus
  ListSpKeyWords.Visible = False
  ComboSpKeyWords.Text = ""
  t1 = Clipboard.GetText
  Clipboard.Clear
  Clipboard.SetText t '+ " "
  SendKeys "^v" 'ctrl+v = paste
  DoEvents
  Clipboard.SetText t1 'previous
 End With
End Sub

Private Sub ComboSpKeyWords_GotFocus()
 ListSpKeyWords.Visible = True
End Sub

Private Sub ComboSpKeyWords_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
  ListSpKeyWords.SetFocus: DoEvents
  If KeyCode = vbKeyDown Then SendKeys "{DOWN}" Else SendKeys "{UP}"
 End If
End Sub

Private Sub ComboSpKeyWords_KeyPress(KeyAscii As Integer)
 'Exit Sub
 If KeyAscii = 8 Then Exit Sub
 If KeyAscii = 27 Then
  KeyAscii = 0: ListSpKeyWords.Visible = False
  Text1.SetFocus: Text1.SelLength = 0: ComboSpKeyWords.Text = "": Exit Sub
 End If
 Dim s As String
 If KeyAscii = 13 Then s = "": KeyAscii = 32 Else s = Chr(KeyAscii)
 If (KeyAscii < 65 Or KeyAscii > 90) And (KeyAscii < 97 Or KeyAscii > 122) Then
  KeyAscii = 0
  InsertText ListSpKeyWords.List(ListSpKeyWords.ListIndex) + s
 End If
End Sub

Private Sub ComboSpKeyWords_LostFocus()
 If ActiveControl.Name <> ListSpKeyWords.Name Then ListSpKeyWords.Visible = False
End Sub

Private Sub CommandRunScript_Click()
 'On Error Resume Next
 'Text1.SetFocus
 'Graphix3D1.ClearScene 'Because Of AddObject
 'With ScriptControl1
 ' .Reset
 ' .AddObject "Gr3D", Graphix3D1, True
 ' .ExecuteStatement Text1.Text
 'End With
 
 'If Err Then
 ' MsgBox Err.Description
 ' Err.Clear
 ' Exit Sub
 'End If
 
 'Graphix3D1.CopyScene UserScene
 
End Sub

Sub CommandOpen_Click()
 On Error Resume Next
 Text1.SetFocus
 With MDIFormMain.CD1
   .Flags = cdlOFNOverwritePrompt + cdlOFNFileMustExist
   .Filter = "VB Script Files (*.Vbs)|*.Vbs|All Files|*.*"
   .ShowOpen
   If Err Then Err.Clear: Exit Sub 'user cancel opening
   Dim f As Long
   f = FreeFile
   Open .FileName For Input As f
   Text1.Text = Input$(LOF(f), #1)
   Close f
   If Err Then MsgBox "Error during load:" + vbCrLf + vbCrLf + Err.Description, , "Error": Err.Clear
   Caption = .FileTitle + " - Script Editor"
   Tag = .FileName
 End With
End Sub

Private Sub CommandSave_Click()
 On Error Resume Next
 Text1.SetFocus
 With MDIFormMain.CD1
   .FileName = Tag
   If .Flags And cdlOFNOverwritePrompt Then .Flags = .Flags Xor cdlOFNOverwritePrompt
   If Tag = "" Then .Flags = cdlOFNOverwritePrompt
   .Flags = .Flags Or cdlOFNFileMustExist
   .Filter = "VB Script Files (*.Vbs)|*.Vbs|All Files|*.*"
   .ShowSave
   If FormGrIsVisible Then FormGr.DoPaint 'Call RefreshGDIwindow (only a fast bitblt)
   If Err Then Err.Clear: Exit Sub 'user cancel opening
   Dim f As Long
   f = FreeFile
   Open .FileName For Output As f
   Print #f, Text1.Text
   Close f
   If Err Then MsgBox "Error during save:" + vbCrLf + vbCrLf + Err.Description, , "Error": Err.Clear
   Caption = .FileTitle + " - Script Editor"
   Tag = .FileName
 End With
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF5 Then CommandRunScript_Click
 If KeyCode = 27 Then ListSpKeyWords.Visible = False: Text1.SetFocus
End Sub

Private Sub Form_Load()
 Cursor2End
End Sub

Sub Cursor2End()
 Text1.SelStart = Len(Text1.Text)
End Sub

Private Sub Form_Resize()
 If ScaleWidth > 100 Then Text1.Width = ScaleWidth - 1
 If ScaleHeight > 100 Then Text1.Height = ScaleHeight - 1 - Text1.Top
End Sub

Private Sub ListSpKeyWords_Click()
 With ListSpKeyWords
  'If DontInsertText = 0 Then InsertText .List(.ListIndex)
  If ActiveControl.Name <> "ComboSpKeyWords" Then ComboSpKeyWords.Text = .List(.ListIndex)
 End With
End Sub

Sub CompleteSpKeyWords()
 Dim t As String, n As Long, i As Long
 t = LCase(ComboSpKeyWords.Text)
 'With ComboSpKeyWords
 With ListSpKeyWords
  n = Len(t)
  For i = 0 To .ListCount - 1
   If LCase(Left(.List(i), n)) = t Then
     DontInsertText = 1: .ListIndex = i: DontInsertText = 0
     Exit For
   End If
  Next
 End With
End Sub

Private Sub ComboSpKeyWords_Change()
 If ActiveControl.Name <> "ComboSpKeyWords" Then Exit Sub
 ListSpKeyWords.Visible = True
 CompleteSpKeyWords
End Sub

Private Sub ComboSpKeyWords_Click()
 'Exit Sub
 If ActiveControl.Name <> "ComboSpKeyWords" Then Exit Sub
 Dim t As String
 t = ComboSpKeyWords.Text
 InsertText t
End Sub

Private Sub ListSpKeyWords_DblClick()
 Call ListSpKeyWords_KeyPress(13)
End Sub

Private Sub ListSpKeyWords_KeyPress(KeyAscii As Integer)
 With ListSpKeyWords
 Dim s As String
 If KeyAscii <> 13 Then s = Chr(KeyAscii) Else s = ""
 If KeyAscii = 13 Or ((KeyAscii < 65 Or KeyAscii > 90) And (KeyAscii < 97 Or KeyAscii > 122)) Then
  KeyAscii = 0
  .Visible = False
  InsertText .List(.ListIndex) + s
 End If
 End With
End Sub

Private Sub Text1_GotFocus()
 'If ListSpKeyWords.Visible Then ListSpKeyWords.Visible = False 'for more safety
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
 With Text1
  If KeyCode = 32 And (Shift = 2) Then 'Ctrl+Space
   Dim i As Long, c As Long, i2 As Long, w As String, m As String, a As Integer
   If .SelLength > 0 Then
    w = .SelText
   Else
    'find last word near cursor and store it to w:
    w = ""
    i2 = .SelStart: c = 0
    For i = i2 To 1 Step -1
      m = Mid(.Text, i, 1)
      a = Asc(UCase(m))
      If a < 65 Or a > 90 Then Exit For
      w = m + w: c = c + 1
    Next
    If c <> 0 Then
     .SelStart = i2 - c
     .SelLength = c
    End If
   End If
   With ComboSpKeyWords
    .SetFocus
    .Text = w
    .SelLength = 0: .SelStart = Len(w)
   End With
  End If
 End With
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 If ActiveControl.Name <> Text1.Name Then KeyAscii = 0: Exit Sub
End Sub
