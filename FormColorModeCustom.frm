VERSION 5.00
Begin VB.Form FormColorModeCustom 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choose Color..."
   ClientHeight    =   795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2775
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
   ScaleHeight     =   795
   ScaleWidth      =   2775
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Cyan"
      Height          =   315
      Index           =   5
      Left            =   1860
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "6"
      Top             =   420
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF80FF&
      Caption         =   "Magenta"
      Height          =   315
      Index           =   4
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "5"
      Top             =   420
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Yellow"
      Height          =   315
      Index           =   3
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "3"
      Top             =   420
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF0000&
      Caption         =   "Blue"
      Height          =   315
      Index           =   2
      Left            =   1860
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "4"
      Top             =   60
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "Green"
      Height          =   315
      Index           =   1
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "2"
      Top             =   60
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Red"
      Height          =   315
      Index           =   0
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "1"
      Top             =   60
      Width           =   855
   End
End
Attribute VB_Name = "FormColorModeCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
 ColorModeV = Val(Command1(Index).Tag)
 Me.Hide
End Sub

