VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FormWizard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Problem Definition Wizard ..."
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7650
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormSolarWizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PictureTopHide 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   385
      Left            =   6600
      ScaleHeight     =   390
      ScaleWidth      =   5895
      TabIndex        =   113
      Top             =   0
      Visible         =   0   'False
      Width           =   5895
      Begin VB.Label LabelDefineTop 
         AutoSize        =   -1  'True
         Caption         =   "Define"
         Height          =   195
         Left            =   120
         TabIndex        =   114
         Top             =   180
         Width           =   465
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4515
      Left            =   30
      TabIndex        =   37
      Top             =   30
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   7964
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      Tab             =   6
      TabsPerRow      =   7
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "                                "
      TabPicture(0)   =   "FormSolarWizard.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Method "
      TabPicture(1)   =   "FormSolarWizard.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture1(2)"
      Tab(1).Control(1)=   "Frame1(0)"
      Tab(1).Control(2)=   "Label2"
      Tab(1).Control(3)=   "Label1"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Location "
      TabPicture(2)   =   "FormSolarWizard.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "LabelLocation1"
      Tab(2).Control(1)=   "Picture1(1)"
      Tab(2).Control(2)=   "Frame2"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Date "
      TabPicture(3)   =   "FormSolarWizard.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label5"
      Tab(3).Control(1)=   "LabelDayBegin"
      Tab(3).Control(2)=   "LabelDayEnd"
      Tab(3).Control(3)=   "Label9"
      Tab(3).Control(4)=   "Label10"
      Tab(3).Control(5)=   "LabelDayStep"
      Tab(3).Control(6)=   "Picture1(0)"
      Tab(3).Control(7)=   "CheckIranDate"
      Tab(3).Control(8)=   "CheckDaylightSave"
      Tab(3).Control(9)=   "ComboMonth(0)"
      Tab(3).Control(10)=   "ComboDayOfMonth(0)"
      Tab(3).Control(11)=   "ComboDayOfMonth(1)"
      Tab(3).Control(12)=   "ComboMonth(1)"
      Tab(3).Control(13)=   "CheckOneDayPeriod"
      Tab(3).Control(14)=   "ComboDayStep"
      Tab(3).ControlCount=   15
      TabCaption(4)   =   "Time "
      TabPicture(4)   =   "FormSolarWizard.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "ComboTimeStep"
      Tab(4).Control(1)=   "CheckFindRate"
      Tab(4).Control(2)=   "Frame4"
      Tab(4).Control(3)=   "TextTime(0)"
      Tab(4).Control(4)=   "TextTime(1)"
      Tab(4).Control(5)=   "Picture1(3)"
      Tab(4).Control(6)=   "LabelTimeStepMin"
      Tab(4).Control(7)=   "LabelTimeStep"
      Tab(4).Control(8)=   "LabelTimeBegin"
      Tab(4).Control(9)=   "LabelTimeEnd"
      Tab(4).Control(10)=   "Label18"
      Tab(4).Control(11)=   "Label17"
      Tab(4).Control(12)=   "Label16"
      Tab(4).Control(13)=   "Label15"
      Tab(4).Control(14)=   "Label14"
      Tab(4).Control(15)=   "Label11"
      Tab(4).ControlCount=   16
      TabCaption(5)   =   "Options "
      TabPicture(5)   =   "FormSolarWizard.frx":0396
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "ComboTrace"
      Tab(5).Control(1)=   "CheckTracer"
      Tab(5).Control(2)=   "ComboConsiderShadow"
      Tab(5).Control(3)=   "CheckConsiderShadow"
      Tab(5).Control(4)=   "TextSpecificGh(0)"
      Tab(5).Control(5)=   "CheckSpecificGh"
      Tab(5).Control(6)=   "Frame3"
      Tab(5).Control(7)=   "Frame1(4)"
      Tab(5).Control(8)=   "Picture1(4)"
      Tab(5).Control(9)=   "LabelSpecificGh(0)"
      Tab(5).ControlCount=   10
      TabCaption(6)   =   "Object "
      TabPicture(6)   =   "FormSolarWizard.frx":03B2
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "Label28"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "Frame5"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "Picture1(5)"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).Control(3)=   "CommandDefineObj"
      Tab(6).Control(3).Enabled=   0   'False
      Tab(6).Control(4)=   "CheckWichObj(1)"
      Tab(6).Control(4).Enabled=   0   'False
      Tab(6).Control(5)=   "CommandSaveProblem"
      Tab(6).Control(5).Enabled=   0   'False
      Tab(6).Control(6)=   "CommandScriptEditor"
      Tab(6).Control(6).Enabled=   0   'False
      Tab(6).Control(7)=   "CheckWichObj(0)"
      Tab(6).Control(7).Enabled=   0   'False
      Tab(6).Control(8)=   "CommandLoadObject"
      Tab(6).Control(8).Enabled=   0   'False
      Tab(6).ControlCount=   9
      Begin VB.CommandButton CommandLoadObject 
         Caption         =   "Load Object"
         Height          =   375
         Left            =   4500
         TabIndex        =   117
         Top             =   1380
         Width           =   1185
      End
      Begin VB.ComboBox ComboTrace 
         Height          =   315
         ItemData        =   "FormSolarWizard.frx":03CE
         Left            =   -71640
         List            =   "FormSolarWizard.frx":03DB
         Style           =   2  'Dropdown List
         TabIndex        =   116
         Top             =   1830
         Width           =   2595
      End
      Begin VB.CheckBox CheckTracer 
         Caption         =   "Enable Tracer"
         Height          =   255
         Left            =   -73020
         TabIndex        =   115
         Top             =   1860
         Width           =   1515
      End
      Begin VB.ComboBox ComboDayStep 
         Height          =   315
         ItemData        =   "FormSolarWizard.frx":041A
         Left            =   -69240
         List            =   "FormSolarWizard.frx":0436
         TabIndex        =   108
         Text            =   "1"
         Top             =   3330
         Width           =   675
      End
      Begin VB.CheckBox CheckWichObj 
         Caption         =   "Use Current Defined Object(s)"
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   106
         Top             =   1980
         Width           =   3315
      End
      Begin VB.CommandButton CommandScriptEditor 
         Caption         =   "Script &Editor"
         Height          =   375
         Left            =   3180
         TabIndex        =   32
         Top             =   1380
         Width           =   1185
      End
      Begin VB.ComboBox ComboConsiderShadow 
         Height          =   315
         ItemData        =   "FormSolarWizard.frx":0455
         Left            =   -71580
         List            =   "FormSolarWizard.frx":0462
         Style           =   2  'Dropdown List
         TabIndex        =   104
         Top             =   600
         Width           =   3795
      End
      Begin VB.CheckBox CheckConsiderShadow 
         Caption         =   "Check Shadow"
         Height          =   195
         Left            =   -73020
         TabIndex        =   103
         Top             =   630
         Width           =   1395
      End
      Begin VB.TextBox TextSpecificGh 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   -70200
         MaxLength       =   6
         TabIndex        =   102
         Top             =   1380
         Width           =   705
      End
      Begin VB.CheckBox CheckSpecificGh 
         Caption         =   "Use specific magnitude of radiation. (measured by an instrument)"
         Height          =   255
         Left            =   -73020
         TabIndex        =   100
         Top             =   1110
         Width           =   5265
      End
      Begin VB.ComboBox ComboTimeStep 
         Height          =   315
         ItemData        =   "FormSolarWizard.frx":04DB
         Left            =   -72510
         List            =   "FormSolarWizard.frx":04F4
         TabIndex        =   17
         Text            =   "30"
         Top             =   3030
         Width           =   645
      End
      Begin VB.CheckBox CheckOneDayPeriod 
         Caption         =   "One Day Period"
         Height          =   285
         Left            =   -72990
         TabIndex        =   16
         Top             =   2430
         Value           =   1  'Checked
         Width           =   1425
      End
      Begin VB.CommandButton CommandSaveProblem 
         Caption         =   "&Save Problem"
         Height          =   375
         Left            =   5820
         TabIndex        =   33
         Top             =   1380
         Width           =   1185
      End
      Begin VB.CheckBox CheckWichObj 
         Caption         =   "Use Special Cases:"
         Height          =   315
         Index           =   1
         Left            =   1920
         TabIndex        =   30
         Top             =   2310
         Value           =   1  'Checked
         Width           =   1665
      End
      Begin VB.CommandButton CommandDefineObj 
         Caption         =   "&Create Object"
         Height          =   375
         Left            =   1860
         TabIndex        =   31
         Top             =   1380
         Width           =   1185
      End
      Begin VB.Frame Frame3 
         Caption         =   "Cloud Factor:"
         Height          =   2115
         Left            =   -69390
         TabIndex        =   83
         Top             =   2310
         Width           =   1875
         Begin VB.OptionButton OptionCF 
            Caption         =   "User Defined Value"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   28
            Top             =   630
            Width           =   1695
         End
         Begin VB.OptionButton OptionCF 
            Caption         =   "Use Statistic Data"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   27
            Top             =   300
            Value           =   -1  'True
            Width           =   1605
         End
         Begin VB.TextBox TextCF 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   840
            MouseIcon       =   "FormSolarWizard.frx":0513
            MousePointer    =   99  'Custom
            TabIndex        =   29
            Top             =   970
            Width           =   645
         End
         Begin VB.Label Label27 
            Caption         =   "The Cloud Factor is 0 for ideal clear sky and 1 for full cloudy sky."
            Height          =   675
            Left            =   150
            TabIndex        =   85
            Top             =   1380
            Width           =   1665
         End
         Begin VB.Label LabelCF 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CF ="
            Height          =   195
            Left            =   420
            TabIndex        =   84
            ToolTipText     =   "Cloud Factor"
            Top             =   975
            Width           =   360
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Constants:"
         Height          =   2115
         Index           =   4
         Left            =   -73080
         TabIndex        =   68
         Top             =   2310
         Width           =   3585
         Begin VB.TextBox TextdiffA 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   195
            Index           =   2
            Left            =   600
            MousePointer    =   99  'Custom
            TabIndex        =   72
            Top             =   1740
            Width           =   1155
         End
         Begin VB.TextBox TextdiffA 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   1
            Left            =   600
            MousePointer    =   99  'Custom
            TabIndex        =   71
            Top             =   1470
            Width           =   975
         End
         Begin VB.TextBox TextdiffA 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   0
            Left            =   600
            MousePointer    =   99  'Custom
            TabIndex        =   70
            Top             =   1215
            Width           =   855
         End
         Begin VB.TextBox TextGsc 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   540
            MouseIcon       =   "FormSolarWizard.frx":0665
            MousePointer    =   99  'Custom
            TabIndex        =   69
            ToolTipText     =   "Solar Constant Of Radiation"
            Top             =   570
            Width           =   525
         End
         Begin VB.OptionButton OptionRoGround 
            Caption         =   "Snowy"
            Height          =   255
            Index           =   1
            Left            =   1950
            TabIndex        =   24
            Top             =   870
            Width           =   795
         End
         Begin VB.OptionButton OptionRoGround 
            Caption         =   "Normal"
            Height          =   195
            Index           =   0
            Left            =   1950
            TabIndex        =   23
            Top             =   600
            Value           =   -1  'True
            Width           =   795
         End
         Begin VB.OptionButton OptionRoGround 
            Caption         =   "User Defined"
            Height          =   195
            Index           =   2
            Left            =   1950
            TabIndex        =   25
            Top             =   1200
            Width           =   1275
         End
         Begin VB.TextBox TextRoGround 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   2670
            MaxLength       =   6
            TabIndex        =   26
            Text            =   "0.2"
            Top             =   1470
            Width           =   615
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(w/m2)"
            Height          =   195
            Index           =   0
            Left            =   1110
            TabIndex        =   82
            Top             =   570
            Width           =   510
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A3 ="
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   81
            Top             =   1725
            Width           =   405
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Direct Beam Constant:"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   120
            TabIndex        =   80
            Top             =   300
            Width           =   1620
         End
         Begin VB.Label LabeldifTit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Diffuse Radiation:"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   120
            MousePointer    =   99  'Custom
            TabIndex        =   79
            Top             =   960
            Width           =   1290
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A2 ="
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   78
            Top             =   1470
            Width           =   405
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A1 ="
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   77
            Top             =   1200
            Width           =   405
         End
         Begin VB.Label LabelGsc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gc ="
            Height          =   195
            Left            =   120
            TabIndex        =   76
            ToolTipText     =   "Solar Constant Of Radiation"
            Top             =   570
            Width           =   345
         End
         Begin VB.Label LabelGrRefTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ground Reflectance:"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   1950
            MousePointer    =   99  'Custom
            TabIndex        =   75
            Top             =   300
            Width           =   1485
         End
         Begin VB.Label LabelRo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "r"
            BeginProperty Font 
               Name            =   "Symbol"
               Size            =   9.75
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2220
            TabIndex        =   74
            Top             =   1395
            Width           =   120
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
            Left            =   2340
            TabIndex        =   73
            Top             =   1470
            Width           =   255
         End
      End
      Begin VB.CheckBox CheckFindRate 
         Caption         =   "Find Rate at an instant (w/m2)"
         Height          =   255
         Left            =   -71310
         TabIndex        =   18
         Top             =   3030
         Value           =   1  'Checked
         Width           =   3285
      End
      Begin VB.Frame Frame4 
         Height          =   1005
         Left            =   -71340
         TabIndex        =   65
         Top             =   3330
         Width           =   3375
         Begin VB.OptionButton OptionLSTTime 
            Caption         =   "My values are as LST (local solar time)"
            Height          =   315
            Index           =   0
            Left            =   180
            TabIndex        =   21
            Top             =   180
            Value           =   -1  'True
            Width           =   3075
         End
         Begin VB.OptionButton OptionLSTTime 
            Caption         =   "i want to give civil time"
            Height          =   315
            Index           =   1
            Left            =   180
            TabIndex        =   22
            Top             =   600
            Width           =   1995
         End
      End
      Begin VB.TextBox TextTime 
         Height          =   285
         Index           =   0
         Left            =   -72510
         TabIndex        =   19
         Top             =   3510
         Width           =   885
      End
      Begin VB.TextBox TextTime 
         Height          =   285
         Index           =   1
         Left            =   -72510
         TabIndex        =   20
         Top             =   3960
         Width           =   885
      End
      Begin VB.ComboBox ComboMonth 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         ItemData        =   "FormSolarWizard.frx":07B7
         Left            =   -72300
         List            =   "FormSolarWizard.frx":07B9
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3810
         Width           =   1305
      End
      Begin VB.ComboBox ComboDayOfMonth 
         Height          =   315
         Index           =   1
         ItemData        =   "FormSolarWizard.frx":07BB
         Left            =   -70650
         List            =   "FormSolarWizard.frx":07BD
         TabIndex        =   13
         Top             =   3780
         Width           =   555
      End
      Begin VB.ComboBox ComboDayOfMonth 
         Height          =   315
         Index           =   0
         ItemData        =   "FormSolarWizard.frx":07BF
         Left            =   -70650
         List            =   "FormSolarWizard.frx":07C1
         TabIndex        =   11
         Top             =   3330
         Width           =   555
      End
      Begin VB.ComboBox ComboMonth 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         ItemData        =   "FormSolarWizard.frx":07C3
         Left            =   -72300
         List            =   "FormSolarWizard.frx":07C5
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   3330
         Width           =   1305
      End
      Begin VB.CheckBox CheckDaylightSave 
         Caption         =   "Assume central daylight saving time for spring && summer (so sub one hour to get real time)"
         Height          =   435
         Left            =   -72960
         TabIndex        =   14
         Top             =   1410
         Value           =   1  'Checked
         Width           =   4425
      End
      Begin VB.CheckBox CheckIranDate 
         Caption         =   "Persian Date Format"
         Height          =   285
         Left            =   -72990
         TabIndex        =   15
         Top             =   1980
         Value           =   1  'Checked
         Width           =   4215
      End
      Begin VB.Frame Frame2 
         Caption         =   "Location:"
         Height          =   3165
         Left            =   -73020
         TabIndex        =   49
         Top             =   1200
         Width           =   5085
         Begin VB.TextBox TextAltitude2 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1560
            MouseIcon       =   "FormSolarWizard.frx":07C7
            MousePointer    =   99  'Custom
            TabIndex        =   111
            Top             =   2700
            Width           =   705
         End
         Begin VB.TextBox TextAltitude 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            MouseIcon       =   "FormSolarWizard.frx":0919
            MousePointer    =   99  'Custom
            TabIndex        =   109
            Top             =   1560
            Width           =   705
         End
         Begin VB.ComboBox ComboCity 
            Height          =   315
            ItemData        =   "FormSolarWizard.frx":0A6B
            Left            =   2520
            List            =   "FormSolarWizard.frx":0A96
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   930
            Width           =   1635
         End
         Begin VB.ComboBox ComboCountry 
            Height          =   315
            ItemData        =   "FormSolarWizard.frx":0B02
            Left            =   600
            List            =   "FormSolarWizard.frx":0B0C
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   930
            Width           =   1635
         End
         Begin VB.OptionButton OptionLocation 
            Caption         =   "Specify a City"
            Height          =   315
            Index           =   0
            Left            =   270
            TabIndex        =   4
            Top             =   300
            Value           =   -1  'True
            Width           =   1755
         End
         Begin VB.OptionButton OptionLocation 
            Caption         =   "Arbitrary Location"
            Height          =   315
            Index           =   1
            Left            =   270
            TabIndex        =   8
            Top             =   2070
            Width           =   1815
         End
         Begin VB.TextBox TextLatitude1 
            Enabled         =   0   'False
            Height          =   285
            Left            =   630
            Locked          =   -1  'True
            MouseIcon       =   "FormSolarWizard.frx":0B1B
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   1560
            Width           =   705
         End
         Begin VB.TextBox TextLatitude 
            Height          =   285
            Left            =   630
            MaxLength       =   5
            MouseIcon       =   "FormSolarWizard.frx":0C6D
            MousePointer    =   99  'Custom
            TabIndex        =   9
            Top             =   2700
            Width           =   705
         End
         Begin VB.Label LabelAltitude2 
            AutoSize        =   -1  'True
            Caption         =   "Altitude:"
            Height          =   195
            Left            =   1560
            TabIndex        =   112
            Top             =   2460
            Width           =   615
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Altitude:"
            Height          =   195
            Left            =   1560
            TabIndex        =   110
            Top             =   1350
            Width           =   615
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "City:"
            Height          =   195
            Left            =   2520
            TabIndex        =   53
            Top             =   720
            Width           =   345
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Country:"
            Height          =   195
            Left            =   600
            TabIndex        =   52
            Top             =   720
            Width           =   645
         End
         Begin VB.Label LabelLatitude 
            AutoSize        =   -1  'True
            Caption         =   "Latitude:"
            Height          =   195
            Left            =   630
            TabIndex        =   51
            Top             =   2460
            Width           =   645
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Latitude:"
            Height          =   195
            Left            =   630
            TabIndex        =   50
            Top             =   1350
            Width           =   645
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00881C0E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4495
         Index           =   5
         Left            =   0
         Picture         =   "FormSolarWizard.frx":0DBF
         ScaleHeight     =   298
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   110
         TabIndex        =   43
         Top             =   30
         Width           =   1680
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00881C0E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4495
         Index           =   4
         Left            =   -75000
         Picture         =   "FormSolarWizard.frx":31E6
         ScaleHeight     =   298
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   110
         TabIndex        =   42
         Top             =   30
         Width           =   1680
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00881C0E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4495
         Index           =   3
         Left            =   -75000
         Picture         =   "FormSolarWizard.frx":5771
         ScaleHeight     =   298
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   110
         TabIndex        =   41
         Top             =   30
         Width           =   1680
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00881C0E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4495
         Index           =   0
         Left            =   -75000
         Picture         =   "FormSolarWizard.frx":763C
         ScaleHeight     =   298
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   110
         TabIndex        =   40
         Top             =   30
         Width           =   1680
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00881C0E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4495
         Index           =   2
         Left            =   -75000
         Picture         =   "FormSolarWizard.frx":957F
         ScaleHeight     =   298
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   110
         TabIndex        =   39
         Top             =   30
         Width           =   1680
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00881C0E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4495
         Index           =   1
         Left            =   -75000
         Picture         =   "FormSolarWizard.frx":B78B
         ScaleHeight     =   298
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   110
         TabIndex        =   38
         Top             =   30
         Width           =   1680
      End
      Begin VB.Frame Frame1 
         Caption         =   "Method:"
         Height          =   2565
         Index           =   0
         Left            =   -73050
         TabIndex        =   46
         Top             =   1800
         Width           =   5385
         Begin VB.OptionButton OptionMethod 
            Caption         =   "Dr. Daneshyar Note"
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
            Left            =   120
            TabIndex        =   1
            Top             =   270
            Value           =   -1  'True
            Width           =   2265
         End
         Begin VB.OptionButton OptionMethod 
            Caption         =   "ASHRAE Method"
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
            Left            =   120
            TabIndex        =   2
            Top             =   615
            Width           =   2505
         End
         Begin VB.OptionButton OptionMethod 
            Caption         =   "Hottel Formula"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   3
            Top             =   960
            Width           =   2205
         End
         Begin VB.Label LabelInfo 
            Caption         =   $"FormSolarWizard.frx":DEA6
            Height          =   1035
            Left            =   420
            TabIndex        =   47
            Top             =   1440
            Width           =   4335
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "        "
         Height          =   1815
         Left            =   1860
         TabIndex        =   89
         Top             =   2400
         Width           =   5325
         Begin VB.CommandButton CommandPreview 
            Caption         =   "Preview"
            Height          =   375
            Left            =   3960
            TabIndex        =   99
            Top             =   1260
            Width           =   1185
         End
         Begin VB.CommandButton CommandSettings 
            Caption         =   "Settings..."
            Height          =   375
            Left            =   3960
            TabIndex        =   98
            Top             =   780
            Width           =   1185
         End
         Begin VB.OptionButton OptionSpCase 
            Caption         =   "Four Objects"
            Height          =   285
            Index           =   2
            Left            =   240
            TabIndex        =   92
            Top             =   990
            Width           =   1545
         End
         Begin VB.OptionButton OptionSpCase 
            Caption         =   "Semi Sphere"
            Height          =   285
            Index           =   5
            Left            =   2460
            TabIndex        =   95
            Top             =   660
            Width           =   1245
         End
         Begin VB.OptionButton OptionSpCase 
            Caption         =   " Semi-Cylinder"
            Height          =   285
            Index           =   4
            Left            =   2460
            TabIndex        =   94
            Top             =   330
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.OptionButton OptionSpCase 
            Caption         =   "Horizental Plate"
            Height          =   285
            Index           =   0
            Left            =   240
            TabIndex        =   90
            Top             =   330
            Width           =   1545
         End
         Begin VB.OptionButton OptionSpCase 
            Caption         =   "Vertical Plate"
            Height          =   285
            Index           =   1
            Left            =   240
            TabIndex        =   91
            Top             =   660
            Width           =   1335
         End
         Begin VB.OptionButton OptionSpCase 
            Caption         =   "North - South Collector"
            Height          =   285
            Index           =   3
            Left            =   240
            TabIndex        =   93
            Top             =   1320
            Width           =   1965
         End
         Begin VB.OptionButton OptionSpCase 
            Caption         =   "Dome"
            Height          =   285
            Index           =   6
            Left            =   2460
            TabIndex        =   96
            Top             =   990
            Width           =   735
         End
         Begin VB.OptionButton OptionSpCase 
            Caption         =   "Cooling Tower"
            Height          =   285
            Index           =   7
            Left            =   2460
            TabIndex        =   97
            Top             =   1320
            Width           =   1335
         End
      End
      Begin VB.Label LabelDayStep 
         AutoSize        =   -1  'True
         Caption         =   "Day Step:"
         Height          =   195
         Left            =   -69240
         TabIndex        =   107
         Top             =   3030
         Width           =   720
      End
      Begin VB.Label LabelSpecificGh 
         AutoSize        =   -1  'True
         Caption         =   "Horizental total radiation (w/m2) :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   -72720
         TabIndex        =   101
         Top             =   1380
         Width           =   2430
      End
      Begin VB.Label LabelTimeStepMin 
         AutoSize        =   -1  'True
         Caption         =   "Min."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   -71820
         TabIndex        =   88
         Top             =   3120
         Width           =   270
      End
      Begin VB.Label LabelTimeStep 
         AutoSize        =   -1  'True
         Caption         =   "Step:"
         Height          =   195
         Left            =   -73020
         TabIndex        =   87
         Top             =   3090
         Width           =   390
      End
      Begin VB.Label Label28 
         Caption         =   $"FormSolarWizard.frx":DFAB
         Height          =   735
         Left            =   1950
         TabIndex        =   86
         Top             =   570
         Width           =   5235
      End
      Begin VB.Label LabelTimeBegin 
         AutoSize        =   -1  'True
         Caption         =   "Begin:"
         Height          =   195
         Left            =   -73020
         TabIndex        =   67
         Top             =   3540
         Width           =   450
      End
      Begin VB.Label LabelTimeEnd 
         AutoSize        =   -1  'True
         Caption         =   "End :"
         Height          =   195
         Left            =   -73020
         TabIndex        =   66
         Top             =   3990
         Width           =   375
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "SN: Solar Noon"
         Height          =   195
         Left            =   -73020
         TabIndex        =   64
         Top             =   2640
         Width           =   1080
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "SS: Sun Set"
         Height          =   195
         Left            =   -73020
         TabIndex        =   63
         Top             =   2370
         Width           =   840
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "SR: Sun Rise"
         Height          =   195
         Left            =   -73020
         TabIndex        =   62
         Top             =   2100
         Width           =   915
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "You can use some key words to mean special times as these:"
         Height          =   195
         Left            =   -73020
         TabIndex        =   61
         Top             =   1800
         Width           =   4365
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Special key words:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -73020
         TabIndex        =   60
         Top             =   1530
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   $"FormSolarWizard.frx":E052
         Height          =   825
         Left            =   -73020
         TabIndex        =   59
         Top             =   600
         Width           =   4785
      End
      Begin VB.Label Label10 
         Caption         =   "Day:"
         Height          =   255
         Left            =   -70650
         TabIndex        =   58
         Top             =   3030
         Width           =   495
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Month:"
         Height          =   195
         Left            =   -72270
         TabIndex        =   57
         Top             =   3030
         Width           =   510
      End
      Begin VB.Label LabelDayEnd 
         AutoSize        =   -1  'True
         Caption         =   "End  :"
         Height          =   195
         Left            =   -72930
         TabIndex        =   56
         Top             =   3870
         Width           =   420
      End
      Begin VB.Label LabelDayBegin 
         AutoSize        =   -1  'True
         Caption         =   "Date:"
         Height          =   195
         Left            =   -72930
         TabIndex        =   55
         Top             =   3360
         Width           =   405
      End
      Begin VB.Label Label5 
         Caption         =   $"FormSolarWizard.frx":E127
         Height          =   705
         Left            =   -73020
         TabIndex        =   54
         Top             =   630
         Width           =   4845
      End
      Begin VB.Label LabelLocation1 
         Caption         =   "You must determine your geographical location such as latitude , or you can specify a city from the list."
         Height          =   525
         Left            =   -73020
         TabIndex        =   48
         Top             =   630
         Width           =   4335
      End
      Begin VB.Label Label2 
         Caption         =   "In the first step you must choose Method of calculating radiation. you have three choises:"
         Height          =   555
         Left            =   -73020
         TabIndex        =   45
         Top             =   1230
         Width           =   4215
      End
      Begin VB.Label Label1 
         Caption         =   "This wizard helps you to define a complete solar radiation problem easily."
         Height          =   495
         Left            =   -73050
         TabIndex        =   44
         Top             =   630
         Width           =   4365
      End
   End
   Begin VB.CommandButton CommandFinish 
      Caption         =   "Finish"
      Height          =   315
      Left            =   6510
      TabIndex        =   34
      Top             =   4620
      Width           =   1095
   End
   Begin VB.CommandButton CommandCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2580
      TabIndex        =   36
      Top             =   4620
      Width           =   1095
   End
   Begin VB.CommandButton CommandBack 
      Caption         =   "Back"
      Height          =   315
      Left            =   3990
      TabIndex        =   35
      Top             =   4620
      Width           =   1095
   End
   Begin VB.CommandButton CommandNext 
      Caption         =   "Next"
      Default         =   -1  'True
      Height          =   315
      Left            =   5100
      TabIndex        =   0
      Top             =   4620
      Width           =   1095
   End
   Begin VB.CheckBox CheckShowOnStartUp 
      Caption         =   "Show this at start up"
      Height          =   225
      Left            =   90
      TabIndex        =   105
      Top             =   4680
      Value           =   1  'Checked
      Width           =   1815
   End
End
Attribute VB_Name = "FormWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CheckConsiderShadow_Click()
 Dim a As Integer
 a = CheckConsiderShadow.Value
 With ComboConsiderShadow
  .Enabled = a
  If .ListIndex = -1 Then .ListIndex = 0
 End With
End Sub

Private Sub CheckFindRate_Click()
If CheckFindRate.Value Then
 TextTime(1).Visible = False
 LabelTimeEnd.Visible = False
 LabelTimeBegin.Caption = "Time :"
Else
 TextTime(1).Visible = True
 LabelTimeEnd.Visible = True
 LabelTimeBegin.Caption = "Begin:"
End If

End Sub

Private Sub CheckIranDate_Click()
 SetComboMonth CheckIranDate.Value
End Sub

Private Sub CheckOneDayPeriod_Click()
If CheckOneDayPeriod.Value Then
 LabelDayBegin.Caption = "Date:"
 LabelDayEnd.Visible = False
 ComboMonth(1).Visible = False: ComboDayOfMonth(1).Visible = False
 'LabelDayStep.Visible = False: ComboDayStep.Visible = False
Else
 LabelDayBegin.Caption = "Begin:"
 LabelDayEnd.Visible = True
 ComboMonth(1).Visible = True: ComboDayOfMonth(1).Visible = True
 'LabelDayStep.Visible = True: ComboDayStep.Visible = True
End If
End Sub

Private Sub CheckShowOnStartUp_Click()
 SaveSetting "SolarRadiationSSerp", "Sec1", "ShowWizardAtStartUp", CStr(CheckShowOnStartUp.Value)
End Sub

Private Sub CheckSpecificGh_Click()
 setAsSpecificGH
End Sub

Sub setAsSpecificGH()
 Dim i As Integer, a As Integer
 a = CheckSpecificGh.Value
 LabelSpecificGh(0).Enabled = a
 TextSpecificGh(0).Enabled = a
End Sub

Private Sub CheckTracer_Click()
 ComboTrace.Enabled = IIf(CheckTracer.Value, 1, 0)
End Sub

Private Sub CheckwichObj_Click(Index As Integer)
 Static dontc As Integer
 If dontc Then Exit Sub
 Dim i As Integer, b As Integer
 If Index = 0 Then i = 1 Else i = 0
 dontc = 1
 CheckWichObj(Index).Value = 1: CheckWichObj(i).Value = 0
 dontc = 0
 b = CheckWichObj(1).Value <> 0
 For i = 0 To OptionSpCase.UBound
  OptionSpCase(i).Enabled = b
 Next

End Sub

Private Sub ComboCity_Click()
 Dim ct As String
 ct = ComboCity.Text
 sUser.CityName = ct
 sUser.Month = "1"
 SetCityData sUser  'only for get latitude, month is not important in this step
 TextLatitude1.Text = sUser.l
 TextAltitude.Text = sUser.Altitude
End Sub

Private Sub ComboConsiderShadow_Change()
 With ComboConsiderShadow
  UserShadowCode = .ListIndex + 1 '1,2,3
 End With
End Sub

Private Sub ComboCountry_Click()
 Call SetForCountry
End Sub

Private Sub ComboMonth_Click(Index As Integer)
 If ComboDayOfMonth(Index).ListIndex = -1 Then ComboDayOfMonth(Index).ListIndex = 0
End Sub

Private Sub CommandBack_Click()
With SSTab1

If .Tab > 1 Then
 .Tab = .Tab - 1
End If

End With

End Sub

Private Sub CommandCancel_Click()
 Me.Hide
End Sub

Private Sub CommandDefineObj_Click()
 
 Dim x1 As Single, y1 As Single, w1 As Long, n1 As Long
 x1 = ViewXshift: y1 = ViewYshift
 If FormGrIsVisible Then w1 = FormGr.WindowState: FormGr.WindowState = vbMinimized
 n1 = UserScene.NumObjects
 FormCreateObj.Show 1
 'SetViewScale3D ViewScale
 'SetView3D ViewTeta, ViewFi
 'SetViewShift2D x1, y1
 If FormGrIsVisible Then FormGr.WindowState = w1: DoEvents: Call FormGr.DoPaint
 
 'ShouldInitPic1 = 1
 If n1 <> UserScene.NumObjects Then
  CheckWichObj(0).Enabled = True
  CheckWichObj(0).Value = 1 'Call ReCalculate 'if add a new object
 End If

End Sub

Private Sub CommandNext_Click()
Dim c As Single, i As Integer, j As Integer, i2 As Integer, s As String, v As Variant

With SSTab1

 If CheckCorrectData(.Tab) <> 1 Then Exit Sub 'bad data

 If .Tab < .Tabs - 1 Then
  .Tab = .Tab + 1
 End If

End With

End Sub

Function CheckCorrectData(ByVal nTab As Integer) As Integer
'return 1 if no problem
Dim s As String, i As Long, j As Long, v As Variant, i2 As Long, c As String
Select Case nTab 'check for correct data
Case 3 'Date
 If CheckOneDayPeriod.Value Then j = 0 Else j = 1
 For i = 0 To j
  If ComboMonth(i).Text = "" Or ComboDayOfMonth(i).Text = "" Then
    MsgBox "Please Specify Date or Date-Range", vbExclamation, "Bad Date!"
    SSTab1.Tab = 3: ComboMonth(i).SetFocus: Exit Function
  End If
 Next
Case 4 'Time
 If CheckFindRate.Value = 1 Then j = 0 Else j = 1
 v = Array("SR", "SS", "SN")
 For i = 0 To j
  s = UCase(TextTime(i).Text)
  For i2 = 0 To 2: s = Replace(s, v(i2) + "+", ""): s = Replace(s, v(i2), "4:30"): s = Replace(s, "+" + v(i2), ""): Next
  If BadTimeSyntax(s) Then
    MsgBox "Bad Time Syntax or Value!" + vbCrLf + vbCrLf + "Please correct it to continue.", vbExclamation, "Bad Time!"
    SSTab1.Tab = 4: TextTime(i).SetFocus: Exit Function
  End If
 Next
Case 5 'Considerations
 If OptionCF(0).Value = False Then
  c = Val(TextCF.Text): If c < 0 Or c > 1 Then MsgBox "Cloud Factor must be between 0 and 1" + vbCrLf + vbCrLf + "Please correct it!": TextCF.SetFocus: Exit Function
 End If

End Select

CheckCorrectData = 1 'means that No Error

End Function

Sub SetForCountry()
 Dim t As String, i As Long, ct As Variant
 Dim jct As Long

 t = UCase(ComboCountry.Text)
 Select Case t
 Case "IRAN"
  jct = 26
  ct = Array("Rasht", "Ramsar", "Babolsar", "Khoy", "Tabriz", "Rezaeyeh", "Saghez", _
             "Tehran", "Hamedan", "Kermanshah", "Kashan", "Khorramabad", "Shahrekord", "Shahroud", _
             "Mashhad", "Sabzevar", "Semnan", "Torbatheidarieh", "Tabas", "Birjand", _
             "Isfahan", "Yazd", "Ahwaz", "Zabol", "Abadan", "Kerman", "Shiraz", "Zahedan", _
             "Bam", "Boshehr", "Bandarabbas", "Iranshahr", "Chabahar")
 Case "USA"
  jct = 0
  ct = Array("NewYork")
 End Select
 
 
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

Private Sub CommandSaveProblem_Click()
On Error Resume Next
With MDIFormMain.CD1
 .Filter = "*.Prb(Problem Files)|*.Prb|All Files|*.*"
 .ShowSave
 If Err Then Err.Clear: Exit Sub
 SaveProblem .FileName
End With
End Sub

Public Sub SaveProblem(FileStr As String)
'this sub save all wizard options in a file
Dim s As String, f As Long, i As Long, j As Long
On Error GoTo SaveProblem_Err
f = FreeFile
Open FileStr For Output As f
Print #f, "'Solar" + Space(1) + "Radiation" + " " + Format(1, "0.0") + " (c) " + "" + Chr(83) + ".S" + Chr(101) + "rpo" + Chr(111) + Chr(115) + "han " + CStr(20) + Format(2, "00") + " - Problem Definition"
'Method
Print #f, "'Method:"
Print #f, "MethodIndex ="; FindSelectedOption(OptionMethod)
'Location:
Print #f, "'Location:"
If OptionLocation(0).Value Then
 Print #f, "Specify a city = "; OptionLocation(0).Value
 Print #f, "CountryIndex ="; ComboCountry.ListIndex + 1
 Print #f, "CityIndex ="; ComboCity.ListIndex + 1
Else
 Print #f, "Arbitrary Location = "; OptionLocation(1).Value
 Print #f, "Latitude = "; TextLatitude.Text
 Print #f, "Altitude = "; TextAltitude2.Text
End If
'Date:
Print #f, "'Date:"
Print #f, "DayLightSavingTime = "; CBool(CheckDaylightSave.Value)
Print #f, "IranDate = "; CBool(CheckIranDate.Value)
Print #f, "OneDayPeriod = "; CBool(CheckOneDayPeriod.Value)
If CheckOneDayPeriod.Value Then j = 0 Else j = 1
For i = 0 To j
 Print #f, "MonthIndex" + CStr(i + 1) + " = " + CStr(ComboMonth(i).ListIndex + 1)
 Print #f, "Day" + CStr(i + 1) + " = " + ComboDayOfMonth(i).Text
Next
If CheckOneDayPeriod.Value = 0 Then Print #f, "DayStep = "; ComboDayStep.Text
'Time:
Print #f, "'Time:"
Print #f, "TimeStepMinutes = "; ComboTimeStep.Text
Print #f, "FindRate = "; CBool(CheckFindRate.Value)
For i = 0 To 1
 If CheckFindRate.Value Then j = 0 Else j = i
 Print #f, "Time" + CStr(i + 1) + " = "; TextTime(j).Text
Next
Print #f, "TimeIsAsLST = "; OptionLSTTime(0).Value
'Options:
Print #f, "ConsiderShadow = "; CBool(CheckConsiderShadow.Value)
If CheckConsiderShadow.Value Then
 Print #f, "CheckShadowIndex ="; ComboConsiderShadow.ListIndex + 1
End If
Print #f, "UseSpecificHorizontalG = "; CBool(CheckSpecificGh.Value)
If CheckSpecificGh.Value Then
 Print #f, "Specific horizontal G ="; TextSpecificGh(0)
End If
With ConstSolar
 Print #f, "GDirectConst ="; Val(TextGsc.Text)
 Print #f, "GdiffuseA1 ="; Val(TextdiffA(0).Text)
 Print #f, "GdiffuseA2 ="; Val(TextdiffA(1).Text)
 Print #f, "GdiffuseA3 ="; Val(TextdiffA(2).Text)
End With
Print #f, "OptionRoGroundIndex ="; FindSelectedOption(OptionRoGround)
If OptionRoGround(2).Value Then
 Print #f, "RoGround Value ="; Val(TextRoGround.Text)
End If
Print #f, "UseStatisticDataForCF = "; OptionCF(0).Value
If OptionCF(1).Value Then
 Print #f, "CF Value ="; Val(TextCF.Text)
End If
'Object:
Print #f, "'Object:"
Print #f, "UseSpecialCaseObject = "; CBool(CheckWichObj(1).Value)
If CheckWichObj(1).Value Then
 Print #f, "SpecialCaseObjectIndex ="; FindSelectedOption(OptionSpCase)
Else
 Print #f, "ObjectFile = ";
End If
Close f

Exit Sub

SaveProblem_Err:
s = Err.Description
Close f
MsgBox "Can't save to file." + vbCrLf + vbCrLf + s, vbCritical, "Save Error"
Err.Clear

End Sub

Public Function SetProblemParamsToUser() As Integer
'set sUser and other user vars by input data given in wizard
'Return True if be successful and False else
Dim i As Integer, j As Integer, c As SinCos3AngType, s As String, dz As Single
Dim CubeT1 As Object3DType, Obj1 As Object3DType, page1 As Page3DType, page2 As Page3DType
Dim k As Long, HorizPage As Object3DType, v As Vector3DType, dx As Single, Az As Single, TwoTower As Integer
Dim c3 As Variant

For i = 1 To 7
 If CheckCorrectData(i) <> 1 Then SetProblemParamsToUser = False: Exit Function
Next

'Method:
MethodOfSolarCalc = FindSelectedOption(OptionMethod)

'Location:
If OptionLocation(0).Value Then 'specify a city
 UserSetLocation ComboCity.Text
Else 'arbitrary location
 UserSetLocation "ArbitraryLocation"
 sUser.l = Val(TextLatitude.Text)
 sUser.Altitude = Val(TextAltitude2.Text)
 sUser.DifLong2Est = 0
End If

'Date:
DayLightSavingTime = -CheckDaylightSave.Value
j = CheckIranDate.Value
If CheckOneDayPeriod.Value Then
 ComboDayOfMonth(1).ListIndex = ComboDayOfMonth(0).ListIndex
 ComboMonth(1).ListIndex = ComboMonth(0).ListIndex
End If
For i = 0 To 1
 UserSetDate Val(ComboDayOfMonth(i).Text), ComboMonth(i).Text, j
 nDayBeginEnd(i) = sUser.nDay
Next
UserDayStep = Val(ComboDayStep.Text): If UserDayStep <= 1 Then UserDayStep = 1
If CheckOneDayPeriod.Value Then nDayBeginEnd(1) = nDayBeginEnd(0)

'Time:
If CheckFindRate.Value Then
 For i = 0 To 1: nTimeBeginEnd(i) = TextTime(0).Text: Next
Else
 For i = 0 To 1: nTimeBeginEnd(i) = TextTime(i).Text: Next
End If
TimeIsLST = OptionLSTTime(0).Value
UserdMinutes = Val(ComboTimeStep.Text)

'Options:
With ComboConsiderShadow
 If CheckConsiderShadow.Value = 0 Then UserShadowCode = 0 Else UserShadowCode = .ListIndex + 1
End With
 
UseSpecificHorizG = CheckSpecificGh.Value <> 0
SpecificHorizG = TextSpecificGh(0)

If CheckTracer.Value Then
 UserTraceAsCollector = ComboTrace.ListIndex + 1
Else
 UserTraceAsCollector = 0
End If

With ConstSolar
  .GDirectC = Val(TextGsc.Text)
  .GdiffuseA1 = Val(TextdiffA(0).Text)
  .GdiffuseA2 = Val(TextdiffA(1).Text)
  .GdiffuseA3 = Val(TextdiffA(2).Text)
End With
RoGround = Val(TextRoGround.Text)
UseStatisticDataForCF = OptionCF(0).Value
FixedCF = Val(TextCF.Text)


'Object:
If CheckWichObj(1).Value Then 'SpCase
 If SpCaseObjectChange = False Then GoTo SetProbAftObj
 ClearScene UserScene
 SpCaseObjectChange = False
 TargetObjIndex = -1
 SpCaseObjectCode = FindSelectedOption(OptionSpCase)
 UserScene.DefaultView = ViewParams(34, 75, 3, 200, 400, Point3D(0, 0, 0))
Else
 If UserScene.DefaultView.View3DScale = 0 Then UserScene.DefaultView = ViewParams(34, 75, 3, 200, 400, Point3D(0, 0, 0))
End If

HorizPageCol = GetRealNearestColor(hdc, &H406060)
HorizPage = ConvertPage2Obj(TransferPage3D(MakeHorizRectangle(300, 200, HorizPageCol), Vector3D(-175, -120, 0)), 0, &H6095C0, DefaultHorizontalPlateStr)
HorizPage.CenterPoint = Point3D(0, 0, -2000)
AddObjectToScene UserScene, HorizPage
SetColGradArray
SunDistance = 80

If CheckWichObj(1).Value Then
 
 With UserScene

 Select Case SpCaseObjectCode
 Case 1 'Horizontal Plate
   Obj1 = ConvertPage2Obj(EvalPage3D(MakeHorizRectangle(10, 10, &HFF5500), 2), 4, 0, "H-Plate")
   'RotateSelfObject3D Obj1, SinCos3Angles(0, 45, 0): TransferSelfObject3D Obj1, Vector3D(0, 0, 10): Obj1.NameStr = "Steep45-Plate"
   AddObjectToScene UserScene, Obj1
   UserScene.NameStr = "H-Plate"
   UserScene.DefaultView = ViewParams(34, 60, 6, 200, 300, Point3D(0, 0, 0))
 Case 2 'Vertical plate
   Obj1 = ConvertPage2Obj(RotatePage3D(MakeHorizRectangle(10, 10, &HFF5500), SinCos3Angles(0, 90, 0)), 4, 0, "V-Plate")
   TransferSelfObject3D Obj1, Vector3D(0, 0, 10 + 2)
   AddObjectToScene UserScene, Obj1
   UserScene.NameStr = "V-Plate"
   UserScene.DefaultView = ViewParams(34, 60, 6, 200, 300, Point3D(0, 0, 0))
 Case 3 '4Obj (Steep Roof)
   dz = 0
   v = Vector3D(38, -25, dz)
   AddObjectToScene UserScene, ConvertPage2Obj(TransferPage3D(MakeHorizRectangle(30, 50, &HFF5500), v), 4, 0, "H-Plate")
    '4Obj-SemiCylinder
    Obj1 = MakeSemiCylinder(90, Pi, 30 * 4, 21, 4, -1)
    Obj1.NameStr = "Semi-Cylinder"
    RotateSelfObject3D Obj1, SinCos3Angles(Pi2, 0, 0, True)
    'TransferSelfObject3D obj1, Vector3D(-120, 0, 0)
    AddObjectToScene UserScene, Obj1
    UserScene.DefaultView = ViewParams(34, 70, 3.8, 300, 300, Point3D(0, 0, 0))
   
   '4Obj-SemiSphere:
   Obj1 = MakeCylinderObj(16 + 4, 20, 25 * 4, &HFFFFFF, -1, "Semi Sphere")
   ScaleSelfObject3D Obj1, 0.7, Point3D(0, 0, 0)
   v = Vector3D(-60, 0, dz)
   AddObjectToScene UserScene, TransferObject3D(Obj1, v)
   
   '4Obj-Dome:
   Obj1 = MakeCylinderObj(32 + 4, 25, 25 * 4, &HFF5555, -1, "Dome")
   ScaleSelfObject3D Obj1, 0.7, Point3D(0, 0, 0)
   v = Vector3D(-115, 0, dz)
   AddObjectToScene UserScene, TransferObject3D(Obj1, v)
   
 Case 4 'North-South Collector
   'UserTraceAsCollector = True
   Obj1 = MakeYfXObj(-10, 10, 100, 1 / 20, 0, 60, 1, 128 + 4 + 8, &HFFFFFF, -1, 0, "NS.Collector")
   RotateSelfObject3D Obj1, SinCos3Angles(90, 0, -90) 'SinCos3Angles(90, 0, -90)
   'RotateSelfObject3D obj1, SinCos3Angles(20, 0, 0)
   AddObjectToScene UserScene, Obj1
   UserScene.NameStr = "NS.Collector"
   UserScene.DefaultView = ViewParams(34, 60, 4.5, 200, 300, Point3D(0, 0, 0))
   SunDistance = 60
   'UserShadowCode = 1 'shadow of itself
   
' Case 5 'Sphere(iCode:1)
'   AddObjectToScene UserScene, TransferObject3D(MakeCylinderObj(1 + 4, 40, 20 * 4, &HFFFFFF, -1, "Sphere"), Vector3D(0, 0, 40))
'   UserScene.NameStr = "Sphere"
'   UserScene.DefaultView = ViewParams(34, 70, 3.5, 300, 300, Point3D(0, 0, 0))
'   SunDistance = 180
 
 Case 5 'semi-Cylinder
    'MakeObjByRotateCurve Obj1, "30", -25, 25, 2, 0, Pi, 30 * 4, 4, "Semi-Cylinder", -1
    Obj1 = MakeSemiCylinder(90, Pi, 30 * 4, 21, 4, -1)
    Obj1.NameStr = "Semi-Cylinder"
    RotateSelfObject3D Obj1, SinCos3Angles(Pi2, 0, Pi2, True)
    AddObjectToScene UserScene, Obj1
    UserScene.NameStr = "Semi Cylinder"
    UserScene.DefaultView = ViewParams(34, 70, 3.8, 300, 300, Point3D(0, 0, 0))
    SunDistance = 80
    

 Case 6 'SemiSphere(iCode:16)       'R_Floor=30
   AddObjectToScene UserScene, TransferObject3D(MakeCylinderObj(16 + 4, 30, 30 * 4, &HFFFFFF, -1, "Semi Sphere"), Vector3D(0, 0, 0))
   UserScene.NameStr = "Semi Sphere"
   UserScene.DefaultView = ViewParams(34, 70, 3.8, 300, 300, Point3D(0, 0, 0))
   SunDistance = 60

 Case 7 'Dome(iCode:32)                'R_Floor=24.28516401
   AddObjectToScene UserScene, TransferObject3D(MakeCylinderObj(32 + 4, 20, 20 * 4, &HFF5555, -1, "Dome"), Vector3D(0, 0, 30))
   page1 = MakeRegularHorizPoly(6, 50, &H777790)
   page2 = EvalPage3D(page1, 30)
   AddObjectToScene UserScene, Join2Page(page1, page2, &H777777, &H550000, 2)
   UserScene.DefaultView = ViewParams(34, 60, 3.5, 300, 300, Point3D(0, 0, 0))
   UserScene.NameStr = "Dome"

 Case 8 'Cooling Tower
   SunDistance = 180
   dx = -127: Az = 35 '-127
   TwoTower = CheckConsiderShadow.Value
   If TwoTower Then 'first tower low quality
    'Obj1 = MakeTowerObj(15, 20 * 4, , -1, 2 + 4 + 8, "C.Tower") 'low quality
    Obj1 = MakeTowerObj(55, 40 * 4, , -1, 2 + 4 + 8, "C.Tower")   'high quality
   Else 'first tower high quality
    Obj1 = MakeTowerObj(32, 35 * 4, , -1, 2 + 4 + 8, "C.Tower")
   End If
   AddObjectToScene UserScene, Obj1
   If TwoTower Then 'if must add secound tower
    'Obj1 = MakeTowerObj(32, 35 * 4, , -1, 2 + 4 + 8, "C.Tower")
    Obj1 = MakeTowerObj(60, 45 * 4, , -1, 2 + 4, "C.Tower")
    TransferSelfObject3D Obj1, Vector3D(dx, 0, 0)
    RotateSelfObject3D Obj1, SinCos3Angles(0, 0, Az)
    Obj1.NameStr = "C.Tower2"
    Obj1.iCode = 2 + 4 'secound tower dont has shadow on first so (no +8)
    AddObjectToScene UserScene, Obj1
   End If
   'UserShadowCode = 0 ' 2
   
   'some extra objects
   'CubeT1 = TransferObject3D(MakeCube(30, 20, 10, &H660000, &HFF4400), Vector3D(40, 40, 0))
   'AddObjectToScene UserScene, CubeT1
   page1 = MakeRegularHorizPoly(12, 36, &H354260)
   page2 = RotatePage3D(EvalPage3D(page1, 19.3), SinCos3Angles(0, 0, 10))
   Obj1 = Join2Page(page1, page2, &H777777, 0, 2)
   AddObjectToScene UserScene, Obj1
   If TwoTower Then
    TransferSelfObject3D Obj1, Vector3D(dx, 0, 0)
    RotateSelfObject3D Obj1, SinCos3Angles(0, 0, Az)
    AddObjectToScene UserScene, Obj1
   End If
   UserScene.NameStr = "Cooling Tower"
   UserScene.DefaultView = ViewParams(34, 83, 2.5, 300, 300, Point3D(0, 0, 0))
 End Select
 
 End With
 
 If FormTransforms.CancelPressed = False Then
  'Apply Setting for object(s)
  For i = 0 To UserScene.NumObjects - 1
   If UserScene.Object(i).iCode And 4 Then
     ApplyTransformsToObj UserScene.Object(i)
   End If
  Next
 End If 'CancelPressed
 
Else 'no special case object
 SpCaseObjectCode = -1 'User-Defined Object

End If

SetProbAftObj:

SetProblemParamsToUser = True

End Function

Private Sub CommandFinish_Click()
 Dim c As Long
 If SetProblemParamsToUser = False Then Exit Sub
 Hide
 UserFindRadiationScene UserdMinutes, ComponentG, True, UserTraceAsCollector, 0
 
 'if is shown before must reload it to reset Grid1,ColTable,... with new calculated data
 'FormGr.ReLoadForm

 If FormGrIsVisible = False Then
  FormGr.Show
 Else
  FormGr.SetFocus
  Call FormGr.SetGMinMaxTot
  FormDataGrid.Show
  Call FormGr.DrawColTable
 End If
 
 Call FormGr.SetDataGridValues
 FormGr.DrawGrShapes
 
End Sub

Private Sub CommandScriptEditor_Click()
 Me.Hide
 FormScriptEditor.Show
End Sub

Private Sub CommandSettings_Click()
 FormTransforms.Show 1
End Sub

Sub Form_Activate()
 If UserScene.NumObjects = 0 Then
  CheckWichObj(0).Value = 0: CheckWichObj(1).Value = 1
  CheckWichObj(0).Enabled = False
 Else
  CheckWichObj(0).Enabled = True
 End If
 ComboTrace.Enabled = IIf(CheckTracer.Value, 1, 0)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If UnloadMode = 0 Then Me.Hide: Cancel = True
End Sub

Private Sub LabelCF_Click()
 TextCF.SetFocus
End Sub

Private Sub OptionSpCase_Click(Index As Integer)
 Call FormTransforms.SetAsDefault
 SpCaseObjectChange = True
End Sub

Private Sub TextCF_Change()
 If OptionCF(1).Value Then FixedCF = Val(TextCF.Text)
End Sub

Private Sub TextLatitude1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 If Button And 2 Then MsgBox "this is a locked text box"
End Sub

Private Sub TextLatitude1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
 TextLatitude1.Text = sUser.l
End Sub

Private Sub TextRoGround_Change()
 RoGround = Val(TextRoGround.Text)
 If RoGround < 0 Or RoGround > 1 Then RoGround = 0.2: TextRoGround.Text = "0.2"
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

Private Sub OptionCF_Click(Index As Integer)
Dim i As Long
Select Case Index
Case 0
 TextCF.Enabled = False
 TextCF.Text = "Auto" 'Format(sUser.CF, "0.000")
Case 1
 TextCF.Enabled = True
 TextCF.Text = ""
End Select
LabelCF.Enabled = TextCF.Enabled
End Sub

Private Sub OptionMethod_Click(Index As Integer)
Dim t As Variant
t = Array("This method use a solar constant value wich is the maximum amount of solar radiation in a clear sky on the normal  plate (950 w/m2), and a cloud factor 'CF' (0-1) to consider cloud's effect. also it use three coefficients  A1, A2 , A3 for diffuse radiation.", _
"This method is published by the 'ASHRAE' (American Society of Heating , Refrigerating and Air-conditioning Engineers) .    in this method  'normal direct rad.' is given by  the  formula: G(ND)=A/exp( B/sin(beta) )  where A, B obtain from a table and beta is solar altitude angle.", _
"This method uses a transmittanse factor for the atmosphere (Tb). it depends on altitude and the climate type.")
LabelInfo.Caption = t(Index)
End Sub

Private Sub OptionLocation_Click(Index As Integer)
 Call SetOptionLocation
End Sub

Sub SetOptionLocation()
 Dim i As Integer, j As Integer
 i = Abs(OptionLocation(0).Value <> 0): j = 1 - i
 ComboCountry.Enabled = i: ComboCity.Enabled = i
 TextLatitude1.Enabled = i: TextAltitude.Enabled = i
 LabelLatitude.Enabled = j: TextLatitude.Enabled = j
 LabelAltitude2.Enabled = j: TextAltitude2.Enabled = j
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

Dim i As Integer
With SSTab1
 If .Tab = 0 Then .Tab = 1
 i = (.Tab < .Tabs - 1)
 CommandNext.Visible = i
 If i Then CommandNext.Default = True
 LabelDefineTop.Caption = .Caption
End With

End Sub

Sub SetComboMonth(IranDate As Boolean)
 Dim a As Variant, i As Integer, j As Integer
 If IranDate Then
  a = Array("Farvardin", "Ordibehesht", "Khordad", "Tir", "Mordad", "Shahrivar", "Mehr", "Aban", "Azar", "Dey", "Bahman", "Esfand")
 Else
  a = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
 End If
 
 For i = 0 To 11
  For j = 0 To 1
   ComboMonth(j).List(i) = a(i)
  Next
 Next

End Sub


Private Sub Form_Load()
 
 Dim a As Object, j As Integer, i As Integer
 PictureTopHide.Move 1710, 0
 With ConstSolar
  TextGsc.Text = .GDirectC
  TextdiffA(0).Text = .GdiffuseA1
  TextdiffA(1).Text = .GdiffuseA2
  TextdiffA(2).Text = .GdiffuseA3
 End With
 TextSpecificGh(0).Text = "800"

 OptionLocation_Click (2)
 ComboCountry.ListIndex = 0 'Iran
 ComboTimeStep.ListIndex = 4
 Call SetOptionLocation
 Call setAsSpecificGH
 CheckFindRate_Click
 CheckConsiderShadow_Click
 OptionRoGround_Click (0)
 OptionCF_Click (0)
 
 For j = 0 To 1
   If ComboMonth(j).ListCount = 0 Then
    For i = 1 To 12: ComboMonth(j).AddItem " ": Next
   End If
   For i = 1 To 31
    ComboDayOfMonth(j).AddItem CStr(i)
   Next
 Next
  
 For Each a In Controls
  If TypeOf a Is TextBox Then
   a.MousePointer = vbCustom
   a.MouseIcon = TextCF.MouseIcon
  End If
 Next
 
 OptionCF_Click (0)
 CheckIranDate_Click
 CheckOneDayPeriod_Click
 ComboMonth(0).ListIndex = 6
 CheckFindRate.Value = 1: TextTime(0).Text = "12:00"
 
End Sub

