VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm MDIFormMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Solar Radiation Analyzer"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8700
   Icon            =   "MDIFormSolar.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox CoolBar1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   8640
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   8700
      Begin VB.PictureBox PicToolbar 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   30
         ScaleHeight     =   375
         ScaleWidth      =   8580
         TabIndex        =   1
         Top             =   30
         Width           =   8580
         Begin VB.CommandButton CommandEditor 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   420
            Picture         =   "MDIFormSolar.frx":0E42
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Script Editor"
            Top             =   30
            Width           =   345
         End
         Begin VB.CommandButton CommandNewWizard 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   30
            Picture         =   "MDIFormSolar.frx":241E
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "New Wizard Problem"
            Top             =   30
            Width           =   345
         End
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   1230
      Top             =   870
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Menu MnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNewWizard 
         Caption         =   "&New Problem Wizard                "
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuNewScriptEditor 
         Caption         =   "New &Script Editor"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuL1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReadProblem 
         Caption         =   "&Read Problem"
      End
      Begin VB.Menu mnuOpenScript 
         Caption         =   "Open Script"
      End
      Begin VB.Menu mnuL2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuWindowTitle 
      Caption         =   "&Window"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "MDIFormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandEditor_Click()
 mnuNewScriptEditor_Click
End Sub

Private Sub CommandNewWizard_Click()
 PicToolbar.SetFocus
 mnuNewWizard_Click
End Sub

Private Sub MDIForm_Activate()
 Static TimeActivate As Integer
 
 'FormScriptEditor.Show
 
 If TimeActivate = 0 Then
  TimeActivate = 1
  MDIFormMain.Show
  Dim i As Integer
  i = FormWizard.CheckShowOnStartUp.Value
  i = Val(GetSetting("SolarRadiationSSerp", "Sec1", "ShowWizardAtStartUp", i))
  FormWizard.CheckShowOnStartUp.Value = i
  If i Then Call mnuNewWizard_Click
 End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Call DoEnd
End Sub

Private Sub mnuAbout_Click()
 MsgBox Space(6) + "Solar Radiation Calculator 1.0" + Space(10) + Chr(13) + vbCrLf + Space(6) + "By:" + Serpoosh + " - (C) " + CStr(2002), , "About.."
End Sub

Private Sub mnuExit_Click()
Call DoEnd
End Sub

Private Sub mnuNewScriptEditor_Click()
 Dim f As New FormScriptEditor
 f.Show
End Sub

Private Sub mnuNewWizard_Click()
 If FormDataGrid.Visible Then FormDataGrid.Visible = False
 SpCaseObjectChange = True
 Load FormWizard
 FormWizard.SSTab1.Tab = 1
 FormWizard.Show 1
 DoEvents
 If FormGrIsVisible Then FormGr.SetFocus
End Sub

Private Sub mnuOpenScript_Click()
 On Error Resume Next
 mnuNewScriptEditor_Click
 MDIFormMain.ActiveForm.CommandOpen_Click
 
End Sub

