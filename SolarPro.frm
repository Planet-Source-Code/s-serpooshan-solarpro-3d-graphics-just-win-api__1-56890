VERSION 5.00
Begin VB.Form FormSolar1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00CACACA&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solar Radiation "
   ClientHeight    =   8220
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11910
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
   MDIChild        =   -1  'True
   ScaleHeight     =   548
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   794
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFEDED&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   5880
      ScaleHeight     =   285
      ScaleWidth      =   2505
      TabIndex        =   92
      Top             =   180
      Width           =   2535
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1020
         TabIndex        =   93
         Top             =   0
         Width           =   435
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFEDED&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3060
      ScaleHeight     =   285
      ScaleWidth      =   2505
      TabIndex        =   90
      Top             =   180
      Width           =   2535
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1020
         TabIndex        =   91
         Top             =   0
         Width           =   390
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFEDED&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   240
      ScaleHeight     =   285
      ScaleWidth      =   2505
      TabIndex        =   88
      Top             =   180
      Width           =   2535
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   900
         TabIndex        =   89
         Top             =   0
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFEDED&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      ScaleHeight     =   345
      ScaleWidth      =   2505
      TabIndex        =   86
      Top             =   3600
      Width           =   2535
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Method"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   900
         TabIndex        =   87
         Top             =   60
         Width           =   630
      End
   End
   Begin VB.PictureBox PictureF 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFEDED&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   4
      Left            =   3060
      ScaleHeight     =   345
      ScaleWidth      =   2505
      TabIndex        =   84
      Top             =   3600
      Width           =   2535
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Constants"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   780
         TabIndex        =   85
         Top             =   60
         Width           =   840
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E5CDCD&
      Caption         =   "Method"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   4395
      Index           =   3
      Left            =   240
      TabIndex        =   70
      Top             =   3600
      Width           =   2535
      Begin VB.CheckBox CheckRoGround 
         BackColor       =   &H00E5CDCD&
         Caption         =   "Ground Reflection"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   300
         TabIndex        =   76
         Top             =   3780
         Value           =   1  'Checked
         Width           =   1995
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E5CDCD&
         Caption         =   "diffuse Radiation"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   300
         TabIndex        =   75
         Top             =   3360
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.OptionButton OptionMethod 
         BackColor       =   &H00E5CDCD&
         Caption         =   "Hottel Trans. Method"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   73
         Top             =   1980
         Width           =   2175
      End
      Begin VB.OptionButton OptionMethod 
         BackColor       =   &H00E5CDCD&
         Caption         =   "ASHRAE Method"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   72
         Top             =   1500
         Width           =   1995
      End
      Begin VB.OptionButton OptionMethod 
         BackColor       =   &H00E5CDCD&
         Caption         =   "Daneshyar Formula"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   71
         Top             =   1020
         Value           =   -1  'True
         Width           =   1995
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Considerations:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   300
         TabIndex        =   77
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Calculate Radiation By:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   300
         TabIndex        =   74
         Top             =   540
         Width           =   1965
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFE5E5&
      Caption         =   "Table"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10140
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   6120
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E5CDCD&
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   3135
      Index           =   2
      Left            =   5880
      TabIndex        =   33
      Top             =   180
      Width           =   2535
      Begin VB.CommandButton CommandTime 
         BackColor       =   &H00E5CDCD&
         Caption         =   "Noon"
         Height          =   375
         Index           =   2
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   1200
         Width           =   1035
      End
      Begin VB.CommandButton CommandTime 
         BackColor       =   &H00E5CDCD&
         Caption         =   "Sun Set"
         Height          =   375
         Index           =   1
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   840
         Width           =   1035
      End
      Begin VB.CommandButton CommandTime 
         BackColor       =   &H00E5CDCD&
         Caption         =   "Sun Rise"
         Height          =   375
         Index           =   0
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   480
         Width           =   1035
      End
      Begin VB.TextBox TextTime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E5CDCD&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   180
         TabIndex        =   43
         Text            =   "09:30:00"
         ToolTipText     =   "Central Civil Time"
         Top             =   2520
         Width           =   1035
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E5CDCD&
         Caption         =   "Time:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   180
         TabIndex        =   42
         Top             =   2220
         Width           =   510
      End
      Begin VB.Label LabelSunTime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "09:30:00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   41
         Top             =   1740
         Width           =   975
      End
      Begin VB.Label LabelSunTime 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "09:30:00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   40
         Top             =   1260
         Width           =   975
      End
      Begin VB.Label LabelSunTime 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "09:30:00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   39
         Top             =   900
         Width           =   975
      End
      Begin VB.Label LabelSunTime 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "09:30:00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   38
         Top             =   540
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Day Length:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   180
         TabIndex        =   37
         Top             =   1740
         Width           =   1005
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "S. Noon:"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   240
         TabIndex        =   36
         Top             =   1260
         Width           =   795
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sun Set:"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   35
         Top             =   900
         Width           =   750
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sun Rise:"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   34
         Top             =   540
         Width           =   855
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFE5E5&
      Caption         =   "Graphics"
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
      Left            =   9060
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5CDCD&
      Caption         =   "Constants"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4395
      Index           =   4
      Left            =   3060
      TabIndex        =   13
      Top             =   3600
      Width           =   2535
      Begin VB.TextBox TextRoGround 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5CDCD&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1020
         MaxLength       =   6
         TabIndex        =   97
         Text            =   "0.2"
         Top             =   3840
         Width           =   615
      End
      Begin VB.OptionButton OptionRoGround 
         BackColor       =   &H00E5CDCD&
         Caption         =   "User Defined"
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   96
         Top             =   3540
         Width           =   1275
      End
      Begin VB.OptionButton OptionRoGround 
         BackColor       =   &H00E5CDCD&
         Caption         =   "Normal"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   95
         Top             =   3240
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.OptionButton OptionRoGround 
         BackColor       =   &H00E5CDCD&
         Caption         =   "Snowy"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   94
         Top             =   3240
         Width           =   1035
      End
      Begin VB.TextBox TextGsc 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5CDCD&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   900
         MouseIcon       =   "SolarPro.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Text            =   "950.6"
         ToolTipText     =   "Solar Constant Of Radiation"
         Top             =   900
         Width           =   615
      End
      Begin VB.TextBox TextdiffA 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5CDCD&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   900
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Text            =   "1.53516"
         Top             =   1815
         Width           =   855
      End
      Begin VB.TextBox TextdiffA 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5CDCD&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   900
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Text            =   "2.10503"
         Top             =   2100
         Width           =   855
      End
      Begin VB.TextBox TextdiffA 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5CDCD&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   900
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Text            =   "121.301"
         Top             =   2385
         Width           =   855
      End
      Begin VB.Label LabelRo2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "g ="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   660
         TabIndex        =   80
         Top             =   3840
         Width           =   255
      End
      Begin VB.Label LabelRo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "r"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   540
         TabIndex        =   79
         Top             =   3765
         Width           =   120
      End
      Begin VB.Label LabelGrRefTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ground Reflectance:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   300
         MousePointer    =   99  'Custom
         TabIndex        =   78
         Top             =   2880
         Width           =   1740
      End
      Begin VB.Label LabelGsc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gsc ="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   300
         TabIndex        =   24
         ToolTipText     =   "Solar Constant Of Radiation"
         Top             =   900
         Width           =   495
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A1 ="
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   300
         TabIndex        =   23
         Top             =   1800
         Width           =   465
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A2 ="
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   300
         TabIndex        =   22
         Top             =   2100
         Width           =   465
      End
      Begin VB.Label LabeldifTit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diffuse Radiation:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   300
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   1500
         Width           =   1515
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Direct Beam Constant:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   300
         TabIndex        =   20
         Top             =   540
         Width           =   1920
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A3 ="
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   300
         TabIndex        =   19
         Top             =   2385
         Width           =   465
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(w/m2)"
         Height          =   195
         Left            =   1560
         TabIndex        =   18
         Top             =   930
         Width           =   510
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E5CDCD&
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   3135
      Index           =   1
      Left            =   3060
      TabIndex        =   8
      Top             =   180
      Width           =   2535
      Begin VB.CheckBox CheckPersianDate 
         BackColor       =   &H00E5CDCD&
         Caption         =   "Persian Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   32
         Top             =   2520
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.ComboBox ComboDayOfMonth 
         BackColor       =   &H00E5CDCD&
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "SolarPro.frx":0152
         Left            =   360
         List            =   "SolarPro.frx":0154
         TabIndex        =   12
         Text            =   "1"
         Top             =   1620
         Width           =   855
      End
      Begin VB.ComboBox ComboMonth 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5CDCD&
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "SolarPro.frx":0156
         Left            =   300
         List            =   "SolarPro.frx":0158
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   780
         Width           =   1515
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Month:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   300
         TabIndex        =   10
         Top             =   480
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Day:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   9
         Top             =   1320
         Width           =   390
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E5CDCD&
      Caption         =   "Location"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   3135
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   180
      Width           =   2535
      Begin VB.ComboBox ComboCountry 
         BackColor       =   &H00E5CDCD&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "SolarPro.frx":015A
         Left            =   360
         List            =   "SolarPro.frx":0164
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   780
         Width           =   1815
      End
      Begin VB.ComboBox ComboCity 
         BackColor       =   &H00E5CDCD&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "SolarPro.frx":0173
         Left            =   360
         List            =   "SolarPro.frx":019E
         TabIndex        =   3
         Top             =   1620
         Width           =   1815
      End
      Begin VB.TextBox TextLatitude 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5CDCD&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   2
         Text            =   "30.5"
         Top             =   2460
         Width           =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Country:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "City:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   5
         Top             =   1320
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Latitude:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   4
         Top             =   2160
         Width           =   750
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFEDED&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5880
      ScaleHeight     =   345
      ScaleWidth      =   2505
      TabIndex        =   82
      Top             =   3600
      Width           =   2535
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cloud Factor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   660
         TabIndex        =   83
         Top             =   60
         Width           =   1080
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E5CDCD&
      Caption         =   "Cloud Factor"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   4395
      Index           =   5
      Left            =   5880
      TabIndex        =   25
      Top             =   3600
      Width           =   2535
      Begin VB.TextBox TextCF 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5CDCD&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1200
         MouseIcon       =   "SolarPro.frx":020A
         MousePointer    =   99  'Custom
         TabIndex        =   28
         Text            =   "0.280"
         Top             =   1620
         Width           =   555
      End
      Begin VB.OptionButton OptionCF 
         BackColor       =   &H00E5CDCD&
         Caption         =   "Use Statistic Data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   300
         TabIndex        =   27
         Top             =   720
         Value           =   -1  'True
         Width           =   1995
      End
      Begin VB.OptionButton OptionCF 
         BackColor       =   &H00E5CDCD&
         Caption         =   "User Define Value"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   300
         TabIndex        =   26
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "It is Usual about .1 to .5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   262
         TabIndex        =   81
         Top             =   3780
         Width           =   2010
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "The ave. data are monthly"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   172
         TabIndex        =   51
         Top             =   3480
         Width           =   2190
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "And 1 for Full Cloudy Sky"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   247
         TabIndex        =   31
         Top             =   2940
         Width           =   2040
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CF is 0 for Ideal Clear Sky"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   30
         Top             =   2655
         Width           =   2055
      End
      Begin VB.Label LabelCF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CF  ="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   600
         TabIndex        =   29
         ToolTipText     =   "Cloud Factor"
         Top             =   1635
         Width           =   480
      End
   End
   Begin VB.Label LabelHourAng 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   10020
      TabIndex        =   100
      Top             =   2880
      Width           =   90
   End
   Begin VB.Label LabelTetaZenith 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "60"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   10020
      TabIndex        =   99
      Top             =   2280
      Width           =   180
   End
   Begin VB.Label LabelFi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   10020
      TabIndex        =   98
      Top             =   2580
      Width           =   90
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hour angle ="
      Height          =   195
      Left            =   9000
      TabIndex        =   69
      Top             =   2880
      Width           =   945
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "azimuth  ="
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   9180
      TabIndex        =   68
      Top             =   2580
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "j"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   5
      Left            =   9000
      TabIndex        =   67
      Top             =   2460
      Width           =   150
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "zenith   ="
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   9240
      TabIndex        =   66
      Top             =   2280
      Width           =   690
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "q"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   4
      Left            =   9000
      TabIndex        =   65
      Top             =   2220
      Width           =   105
   End
   Begin VB.Label LabelGTable 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "660.4"
      Height          =   195
      Index           =   2
      Left            =   11100
      TabIndex        =   64
      Top             =   1020
      Width           =   420
   End
   Begin VB.Label LabelGTable 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "660.4"
      Height          =   195
      Index           =   3
      Left            =   10050
      TabIndex        =   63
      Top             =   1860
      Width           =   420
   End
   Begin VB.Label LabelGTable 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "660.4"
      Height          =   195
      Index           =   8
      Left            =   11100
      TabIndex        =   62
      Top             =   1440
      Width           =   420
   End
   Begin VB.Label LabelGTable 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "660.4"
      Height          =   195
      Index           =   6
      Left            =   10050
      TabIndex        =   61
      Top             =   1440
      Width           =   420
   End
   Begin VB.Label LabelGTable 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "660.4"
      Height          =   195
      Index           =   5
      Left            =   11100
      TabIndex        =   60
      Top             =   1860
      Width           =   420
   End
   Begin VB.Label LabelGTable 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "660.4"
      Height          =   195
      Index           =   7
      Left            =   10605
      TabIndex        =   59
      Top             =   1440
      Width           =   420
   End
   Begin VB.Label LabelGTable 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "660.4"
      Height          =   195
      Index           =   4
      Left            =   10605
      TabIndex        =   58
      Top             =   1860
      Width           =   420
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00800000&
      X1              =   663
      X2              =   663
      Y1              =   36
      Y2              =   140
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00800000&
      X1              =   600
      X2              =   772
      Y1              =   60
      Y2              =   60
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NS.Collector"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   12
      Left            =   9000
      TabIndex        =   57
      Top             =   1440
      Width           =   885
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "G tot."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   11
      Left            =   11100
      TabIndex        =   56
      Top             =   600
      Width           =   465
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Horizental"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   10
      Left            =   9000
      TabIndex        =   55
      Top             =   1020
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Normal(max)"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   9
      Left            =   9000
      TabIndex        =   54
      Top             =   1830
      Width           =   915
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Direct"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   8
      Left            =   10020
      TabIndex        =   53
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "diff."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   7
      Left            =   10650
      TabIndex        =   52
      Top             =   600
      Width           =   330
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800000&
      X1              =   736
      X2              =   736
      Y1              =   36
      Y2              =   140
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   704
      X2              =   704
      Y1              =   36
      Y2              =   140
   End
   Begin VB.Label LabelGTable 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100.5"
      Height          =   195
      Index           =   0
      Left            =   10050
      TabIndex        =   50
      Top             =   1020
      Width           =   420
   End
   Begin VB.Label LabelGTable 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "660.4"
      Height          =   195
      Index           =   1
      Left            =   10605
      TabIndex        =   49
      Top             =   1020
      Width           =   420
   End
   Begin VB.Label LabelGaveDay 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Obj - w/m2"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   9000
      TabIndex        =   48
      Top             =   600
      Width           =   795
   End
End
Attribute VB_Name = "FormSolar1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefDbl A-Z
Dim StopComputeG As Boolean, LSTExact(2) As Single
Private Const kRad2Deg = 57.2957795130823 '180/pi


Sub PrintDailyTable()

 Dim s As SolarType, s2 As SolarType
 Dim m As String, ct As String
 Dim sRise As String, sSet As String, l As String
 Dim v As Vector3DType, Vh As Vector3DType
 Dim G As Double, Gave As Double, GDir As Double, Gdif As Double, daySec As Double
 
 'tehran:
 Dim Gtot As Variant, Gv1 As Variant, Gv2(11)
 Gtot = Array(210, 288, 352, 467, 554, 659, 645, 597, 494, 365, 252, 195)
 Gv1 = Array(128, 194, 230, 327, 403, 520, 508, 477, 395, 279, 177, 118)
 For i = 0 To 11
  Gv2(i) = (Gtot(i) - Gv1(i))
 Next
 
 ct = "shiraz"
 Vh.dx = 0: Vh.dy = 0: Vh.dz = 1

Print UCase(ct): Print

Dim TrNSCollector As Boolean
TrNSCollector = True

m = "10"
SetDate 1, m, s
s.CityName = ct
s.Month = m
SetCityData s

s.Beta = 5
s.h = FindHourByBeta(s)
's.h = 0
FindSunBetaFiByHour s
Print "Beta:"; Round(s.Beta, 1)
Print "Fi:"; Round(s.Fi, 1)
Print "h:"; Round(s.h, 1)
c = FindTetaCosineOpt_NSCollector(s)
Print "Cos(incidence):"; Round(c, 2)
G = GDirect(s.Beta / kRad2Deg, c, s.CF) + Gdiffuse(Beta / kRad2Deg, s.CF, 1)
Print "G:"; Round(G, 1)
Exit Sub

For i = 1 To 12

j = 1: m = CStr(i)
SetDate j, m, s
s.CityName = ct
s.Month = m
SetCityData s
G = SigmaGtByHour(s, Vh, 0, 0, 1, G1, G2, TrNSCollector)
Print Format(i, "00)"); "  Gnoon:"; Round(G, 0);
s.Beta = 4: h1 = FindHourByBeta(s)
G = SigmaGtByHour(s, Vh, h1, h1, 1, G1, G2, TrNSCollector)
Print "  /  "; "h4: "; Format(h1, "00"); "   GsR:"; Round(G, 0);

G = SigmaGtDaily(s, Vh, 15 * 60, GDir, Gdif, daySec, TrNSCollector, False): Print "   Gave day:"; Round(G / daySec, 1)

Next

'Exit Sub


SG = 0

For i = 1 To 12
 m = CStr(i)
 s.CityName = ct
 s.Month = m
 SetCityData s
 G = 0: GDir = 0: Gdif = 0
 For j = 1 To 30
  SetDate j, m, s
  G1 = SigmaGtDaily(s, Vh, 15 * 60, GDir, G2, daySec, TrNSCollector)
  G = G + G1 / daySec  '/41868 to cal/m2
  Gdif = Gdif + G2 / daySec ' 41868
 Next
 G = G / (j - 1): Gdif = Gdif / (j - 1)
 SG = SG + G
 Print Format(m, "  00] "); Round(G, 0); "   "; Round(Gdif)
Next
 
Print: Print
Print "Gt ave : "; Round(SG / (12), 0)
Print

s.CityName = ct
s.Month = m
SetCityData s
Vh.dx = 0: Vh.dy = 0: Vh.dz = 1

Print "  " + UCase(ct): Print
m = "1"
s.CityName = ct
s.Month = m
SetCityData s
SetDate 15, m, s
s.LST = 12 * 60 'noon
SetSolarParamsByLST s
'PrintDailyTable
TowerH = 90: TowerTeta = 0
G = TowerRadiation(s, TowerH, TowerTeta, GDir, Gdif, False)
Print "  Gtower = "; Round(G, 2)

End Sub


Private Sub CheckPersianDate_Click()
 SetComboMonth CheckPersianDate.Value
 OptionCF_Click (0)
 ReSetDateTimeParms
End Sub

Private Sub CheckRoGround_Click()
Dim v As Long, b As Boolean
v = CheckRoGround.Value
For i = 0 To 2: OptionRoGround(i).Enabled = v: Next
LabelGrRefTitle.Enabled = v

b = OptionRoGround(2).Enabled And OptionRoGround(2).Value
TextRoGround.Enabled = b
LabelRo.Enabled = b: LabelRo2.Enabled = b

If v = False Then
Else
End If
End Sub

Private Sub ComboCity_Click()
 sUser.CityName = ComboCity.Text
 SetCityData sUser, ComboMonth.Text, CheckPersianDate.Value
 OptionCF_Click (0)
 TextLatitude.Text = sUser.l
 ReSetDateTimeParms
End Sub

Private Sub ComboCountry_Click()
 SetForCountry
End Sub

Private Sub ComboDayOfMonth_Change()
 ComboDayOfMonth_Click
End Sub

Private Sub ComboDayOfMonth_Click()
 ReSetDateTimeParms
End Sub

Private Sub ComboMonth_Click()
 OptionCF_Click (0)
 ReSetDateTimeParms
End Sub

Sub ReSetDateTimeParms()
 
 UserSetDate Val(ComboDayOfMonth.Text), ComboMonth.Text, CheckPersianDate.Value
 Dim st(3) As String, i As Long
 FindSunRiseSunSetNoon sUser, st(0), st(1), st(2), st(3), LSTExact()
 For i = 0 To 3
  LabelSunTime(i).Caption = st(i)
 Next
 Call ComputeG
End Sub

Sub ComputeG(Optional UseLSTsUser As Boolean = False)
 
 Dim s(8) As Double, i As Long, h As Double

 If StopComputeG Then Exit Sub
 
 If Not UseLSTsUser Then
  sUser.LCTstring = TextTime.Text
  SetLSTasLCTstr sUser
 End If
 SetSolarParamsByLST sUser
 h = sUser.h
 s(2) = SigmaGtByHour(sUser, Vector3D(0, 0, 1), h, h, 15 * 60, s(0), s(1))
 s(5) = SigmaGtByHour(sUser, MakeSunVector(sUser.Beta, sUser.Fi), h, h, 15 * 60, s(3), s(4))
 s(8) = SigmaGtByHour(sUser, FindVnOpt_NSCollector(sUser.Beta, sUser.Fi), h, h, 15 * 60, s(6), s(7))
 For i = 0 To 8
  LabelGTable(i).Caption = Format(s(i), "00.0")
 Next
 
 LabelTetaZenith.Caption = Format(90 - sUser.Beta, "0.0")
 LabelFi.Caption = Format(sUser.Fi, "0.0")
 LabelHourAng.Caption = Format(sUser.h, "0.0")
End Sub

Private Sub Command1_Click()
 With FormGr
  .Show
  '.TestDraw
  .DrawGrShapes
 End With

End Sub


Private Sub CommandTime_Click(Index As Integer)
StopComputeG = True
TextTime.Text = LabelSunTime(Index).Caption
sUser.LST = LSTExact(Index)
StopComputeG = False
Call ComputeG(True)
End Sub

Private Sub Form_Load()
'Call PrintDailyTable: Exit Sub
'Dim s As SolarType
'SetDate 21, "mar", s
'Print s.nDay
Move 0, 0
Call InitiateForm

End Sub

Sub InitiateForm()
 Dim i As Long, a As Variant, t As String
 StopComputeG = True
 ComboCountry.ListIndex = 0
 'ComboCity.li
 For i = 1 To 31
  ComboDayOfMonth.AddItem CStr(i)
 Next
 For Each a In Controls
  If TypeOf a Is TextBox Then
   a.MousePointer = vbCustom
   a.MouseIcon = TextCF.MouseIcon
  End If
 Next
 OptionMethod_Click (0)
 OptionCF_Click (0)
 OptionRoGround_Click (0)
 StopComputeG = False
 ReSetDateTimeParms
End Sub

Private Sub SetForCountry()
 Dim t As String, i As Long, a As Variant, ct As Variant
 Dim jm As Long, jct As Long

 t = UCase(ComboCountry.Text)
 Select Case t
 Case "IRAN"
  'ct = Array("Shiraz", "Tehran", "Yazd", "Isfehan", "Bushehr", "BandarAbas", "Chabahar", "Ahvaz", "Tabriz", "Mashad", "Abadan", "Tabas", "Zahedan", "Kerman")
  ct = Array("Rasht", "Ramsar", "Babolsar", "Khoy", "Tabriz", "Rezaeyeh", "Saghez", _
             "Tehran", "Hamedan", "Kermanshah", "Kashan", "Khorramabad", "Shahrekord", "Shahroud", _
             "Mashhad", "Sabzevar", "Semnan", "Torbatheidarieh", "Tabas", "Birjand", _
             "Isfahan", "Yazd", "Ahwaz", "Zabol", "Abadan", "Kerman", "Shiraz", "Zahedan", _
             "Bam", "Boshehr", "Bandarabbas", "Iranshahr", "Chabahar")

  SetComboMonth True
  CheckPersianDate.Value = 1
  jm = 6: jct = 26
 Case "USA"
  ct = Array("NewYork")
  SetComboMonth False
  CheckPersianDate.Value = False
  jm = 9: jct = 0
 End Select
 
 ComboMonth.ListIndex = jm
 
 With ComboCity
  For i = 0 To .ListCount - 1
   .RemoveItem 0
  Next
  For i = 0 To UBound(ct)
   .AddItem ct(i)
  Next
  .ListIndex = jct
 End With
End Sub

Private Sub SetComboMonth(IranDate As Boolean)
 Dim a As Variant
 If IranDate Then
  a = Array("Farvardin", "Ordibehesht", "Khordad", "Tir", "Mordad", "Shahrivar", "Mehr", "Aban", "Azar", "Dey", "Bahman", "Esfand")
 Else
  a = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
 End If
 For i = 0 To 11
  ComboMonth.List(i) = a(i)
 Next

End Sub


Private Sub Form_Unload(Cancel As Integer)
 On Error Resume Next
 Unload FormGr
 End
End Sub

Private Sub OptionCF_Click(Index As Integer)
Dim i As Long
Select Case Index
Case 0
 i = ComboMonth.ListIndex + 1
 TextCF.Enabled = False
 'TextCF.Locked = True
 sUser.CityName = ComboCity.Text
 SetCityData sUser, CStr(i), CheckPersianDate.Value
 TextCF.Text = Format(sUser.CF, "0.000")
Case 1
 TextCF.Enabled = True
End Select
LabelCF.Enabled = TextCF.Enabled
End Sub

Private Sub OptionMethod_Click(Index As Integer)
MethodOfSolarCalc = Index + 1
Select Case MethodOfSolarCalc
Case 1

Case 2

Case 3

End Select

End Sub

Private Sub OptionRoGround_Click(Index As Integer)
With TextRoGround
Select Case Index
Case 0: .Text = "0.2": .Enabled = False
Case 1: .Text = "0.7": .Enabled = False
Case 2: .Enabled = True
End Select
LabelRo.Enabled = .Enabled: LabelRo2.Enabled = .Enabled
End With

End Sub

Private Sub TextCF_Change()
 sUser.CF = Val(TextCF.Text)
End Sub

Private Sub TextLatitude_Change()
sUser.l = Val(TextLatitude.Text)
ReSetDateTimeParms
End Sub

Private Sub TextRoGround_Change()
 RoGround = Val(TextRoGround.Text)
 If RoGround < 0 Or RoGround > 1 Then RoGround = 0.2: TextRoGround.Text = "0.2"
End Sub

Private Sub TextTime_Change()
 If BadTimeSyntax(TextTime.Text) Then Exit Sub
 
 Call ComputeG
End Sub


