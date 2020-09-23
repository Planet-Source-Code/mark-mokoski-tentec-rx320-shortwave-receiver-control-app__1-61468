VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRX320 
   BackColor       =   &H009C7D2C&
   Caption         =   "WA1ZEK's RX-320 PC Control"
   ClientHeight    =   6690
   ClientLeft      =   165
   ClientTop       =   630
   ClientWidth     =   7095
   Icon            =   "frmRX320.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerTuneRepeat 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2040
      Top             =   6840
   End
   Begin VB.Timer TimerTuneDelay 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1560
      Top             =   6840
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2520
      Top             =   6840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton MEMtoVFO 
      BackColor       =   &H00C0C0C0&
      Caption         =   "M > VFO"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6060
      MouseIcon       =   "frmRX320.frx":030A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   1920
      Width           =   900
   End
   Begin VB.CommandButton VFOtoMEM 
      BackColor       =   &H00C0C0C0&
      Caption         =   "VFO > M"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      MouseIcon       =   "frmRX320.frx":0614
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   1920
      Width           =   900
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H009C7D2C&
      Caption         =   "Freq Step"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   3480
      TabIndex        =   87
      Top             =   2400
      Width           =   1095
      Begin VB.ComboBox ComboStep 
         Height          =   315
         Left            =   60
         TabIndex        =   88
         Text            =   "ComboStep"
         Top             =   225
         Width           =   975
      End
   End
   Begin VB.CommandButton FreqUp 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   3600
      MouseIcon       =   "frmRX320.frx":091E
      MousePointer    =   99  'Custom
      Picture         =   "frmRX320.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   86
      Top             =   1175
      Width           =   1300
   End
   Begin VB.CommandButton FreqDown 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   2280
      MouseIcon       =   "frmRX320.frx":106A
      MousePointer    =   99  'Custom
      Picture         =   "frmRX320.frx":1374
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   1175
      Width           =   1300
   End
   Begin VB.Frame FrameFilters 
      BackColor       =   &H009C7D2C&
      Caption         =   "Filter Width"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   3615
      Left            =   3480
      TabIndex        =   75
      Top             =   3000
      Width           =   1095
      Begin VB.CommandButton cmdFilter 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2700 Hz"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   120
         MouseIcon       =   "frmRX320.frx":17B6
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   84
         Tag             =   "2700"
         Top             =   3000
         Width           =   855
      End
      Begin VB.CommandButton cmdFilter 
         BackColor       =   &H00C0C0C0&
         Caption         =   "3300 Hz"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   120
         MouseIcon       =   "frmRX320.frx":1AC0
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   83
         Tag             =   "3300"
         Top             =   2520
         Width           =   855
      End
      Begin VB.ComboBox ComboFilter 
         Height          =   315
         ItemData        =   "frmRX320.frx":1DCA
         Left            =   60
         List            =   "frmRX320.frx":1DCC
         Style           =   2  'Dropdown List
         TabIndex        =   80
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdFilter 
         BackColor       =   &H00C0C0C0&
         Caption         =   "4800 Hz"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   120
         MouseIcon       =   "frmRX320.frx":1DCE
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   79
         Tag             =   "4800"
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton cmdFilter 
         BackColor       =   &H00C0C0C0&
         Caption         =   "5400 Hz"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   120
         MouseIcon       =   "frmRX320.frx":20D8
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   78
         Tag             =   "5400"
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton cmdFilter 
         BackColor       =   &H00C0C0C0&
         Caption         =   "5700 Hz"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   120
         MouseIcon       =   "frmRX320.frx":23E2
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   77
         Tag             =   "5700"
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdFilter 
         BackColor       =   &H00C0C0C0&
         Caption         =   "6000 Hz"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   120
         MouseIcon       =   "frmRX320.frx":26EC
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   76
         Tag             =   "6000"
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdAequalB 
      BackColor       =   &H00C0C0C0&
      Caption         =   "A = B"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      MouseIcon       =   "frmRX320.frx":29F6
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdBtoA 
      BackColor       =   &H00C0C0C0&
      Caption         =   "B -> A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      MouseIcon       =   "frmRX320.frx":2D00
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdAtoB 
      BackColor       =   &H00C0C0C0&
      Caption         =   "A -> B"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      MouseIcon       =   "frmRX320.frx":300A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdVFO 
      BackColor       =   &H00C0C0C0&
      Caption         =   "VFO B"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1080
      MouseIcon       =   "frmRX320.frx":3314
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdVFO 
      BackColor       =   &H00C0C0C0&
      Caption         =   "VFO A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      MouseIcon       =   "frmRX320.frx":361E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   1920
      Width           =   975
   End
   Begin VB.Frame FrameAGC 
      BackColor       =   &H009C7D2C&
      Caption         =   "AGC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1815
      Left            =   1800
      TabIndex        =   45
      Top             =   4800
      Width           =   1575
      Begin VB.PictureBox PicAGC 
         Appearance      =   0  'Flat
         BackColor       =   &H009C7D2C&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   2
         Left            =   60
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   62
         Top             =   1200
         Width           =   495
      End
      Begin VB.PictureBox PicAGC 
         Appearance      =   0  'Flat
         BackColor       =   &H009C7D2C&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   1
         Left            =   60
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   61
         Top             =   720
         Width           =   495
      End
      Begin VB.PictureBox PicAGC 
         Appearance      =   0  'Flat
         BackColor       =   &H009C7D2C&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   0
         Left            =   60
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   60
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton OptionAGC 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Slow"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   600
         MouseIcon       =   "frmRX320.frx":3928
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton OptionAGC 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Medium"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   600
         MouseIcon       =   "frmRX320.frx":3C32
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton OptionAGC 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fast"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   600
         MouseIcon       =   "frmRX320.frx":3F3C
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame FrameMode 
      BackColor       =   &H009C7D2C&
      Caption         =   "Mode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   2295
      Left            =   1800
      TabIndex        =   44
      Top             =   2400
      Width           =   1575
      Begin VB.CommandButton OptionMode 
         BackColor       =   &H00C0C0C0&
         Caption         =   "CW"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   600
         MouseIcon       =   "frmRX320.frx":4246
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton OptionMode 
         BackColor       =   &H00C0C0C0&
         Caption         =   "USB"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   600
         MouseIcon       =   "frmRX320.frx":4550
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton OptionMode 
         BackColor       =   &H00C0C0C0&
         Caption         =   "LSB"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   600
         MouseIcon       =   "frmRX320.frx":485A
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton OptionMode 
         BackColor       =   &H00C0C0C0&
         Caption         =   "AM"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   600
         MouseIcon       =   "frmRX320.frx":4B64
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   240
         Width           =   855
      End
      Begin VB.PictureBox PicMode 
         Appearance      =   0  'Flat
         BackColor       =   &H009C7D2C&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   3
         Left            =   60
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   59
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox PicMode 
         Appearance      =   0  'Flat
         BackColor       =   &H009C7D2C&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   2
         Left            =   60
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   58
         Top             =   1200
         Width           =   495
      End
      Begin VB.PictureBox PicMode 
         Appearance      =   0  'Flat
         BackColor       =   &H009C7D2C&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   1
         Left            =   60
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   57
         Top             =   720
         Width           =   495
      End
      Begin VB.PictureBox PicMode 
         Appearance      =   0  'Flat
         BackColor       =   &H009C7D2C&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   0
         Left            =   60
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   56
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Timer TimerSunit 
      Interval        =   500
      Left            =   1080
      Top             =   6840
   End
   Begin VB.Timer TimerClock 
      Interval        =   500
      Left            =   600
      Top             =   6840
   End
   Begin VB.Frame FrameKeyPad 
      BackColor       =   &H009C7D2C&
      Caption         =   "Frequency Entry"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   4215
      Left            =   4680
      TabIndex        =   6
      Top             =   2400
      Width           =   2295
      Begin VB.PictureBox picFreqEnter 
         Height          =   330
         Left            =   40
         ScaleHeight     =   270
         ScaleWidth      =   2145
         TabIndex        =   67
         Top             =   240
         Width           =   2200
         Begin VB.Label lblNewFreq 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Enter Frequency in MHz"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   300
            Left            =   20
            TabIndex        =   68
            Top             =   0
            Width           =   2125
         End
      End
      Begin VB.Label EntryClear 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CE"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1560
         TabIndex        =   48
         Top             =   2760
         Width           =   615
         WordWrap        =   -1  'True
      End
      Begin VB.Label DigitEnter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Enter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         TabIndex        =   18
         Top             =   3480
         Width           =   2055
      End
      Begin VB.Label Digit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   0
         Left            =   840
         TabIndex        =   17
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label DigitDP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         TabIndex        =   16
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Digit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   9
         Left            =   1560
         TabIndex        =   15
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Digit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   8
         Left            =   840
         TabIndex        =   14
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Digit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   7
         Left            =   120
         TabIndex        =   13
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Digit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   6
         Left            =   1560
         TabIndex        =   12
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Digit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   5
         Left            =   840
         TabIndex        =   11
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Digit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Digit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   3
         Left            =   1560
         TabIndex        =   9
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Digit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   2
         Left            =   840
         TabIndex        =   8
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Digit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.Frame FrameAudio 
      BackColor       =   &H009C7D2C&
      Caption         =   "Audio Levels"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   4215
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   1575
      Begin VB.VScrollBar SpkVol 
         Height          =   3375
         Left            =   240
         Max             =   63
         MouseIcon       =   "frmRX320.frx":4E6E
         MousePointer    =   99  'Custom
         TabIndex        =   47
         Top             =   480
         Width           =   255
      End
      Begin VB.VScrollBar LineVol 
         Height          =   3375
         Left            =   1080
         Max             =   63
         MouseIcon       =   "frmRX320.frx":5178
         MousePointer    =   99  'Custom
         TabIndex        =   46
         Top             =   480
         Width           =   255
      End
      Begin VB.CheckBox Mute 
         BackColor       =   &H009C7D2C&
         Caption         =   "Mute Audio"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   140
         TabIndex        =   5
         Top             =   3840
         Width           =   1300
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Line Out"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   880
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Speaker"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   -15
         TabIndex        =   3
         Top             =   240
         Width           =   825
         WordWrap        =   -1  'True
      End
   End
   Begin MSCommLib.MSComm RadioCOM 
      Left            =   0
      Top             =   6720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      BaudRate        =   1200
   End
   Begin VB.PictureBox DisplayPicture 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   7095
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      Begin VB.PictureBox sUnit 
         BackColor       =   &H000040C0&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   20
         Left            =   3000
         ScaleHeight     =   255
         ScaleWidth      =   105
         TabIndex        =   41
         Top             =   720
         Visible         =   0   'False
         Width           =   100
      End
      Begin VB.PictureBox sUnit 
         BackColor       =   &H000040C0&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   19
         Left            =   2880
         ScaleHeight     =   255
         ScaleWidth      =   105
         TabIndex        =   40
         Top             =   720
         Visible         =   0   'False
         Width           =   100
      End
      Begin VB.PictureBox sUnit 
         BackColor       =   &H000040C0&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   18
         Left            =   2760
         ScaleHeight     =   255
         ScaleWidth      =   105
         TabIndex        =   39
         Top             =   720
         Visible         =   0   'False
         Width           =   100
      End
      Begin VB.PictureBox sUnit 
         BackColor       =   &H000040C0&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   17
         Left            =   2640
         ScaleHeight     =   255
         ScaleWidth      =   105
         TabIndex        =   38
         Top             =   720
         Visible         =   0   'False
         Width           =   100
      End
      Begin VB.PictureBox sUnit 
         BackColor       =   &H000040C0&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   16
         Left            =   2520
         ScaleHeight     =   255
         ScaleWidth      =   105
         TabIndex        =   37
         Top             =   720
         Visible         =   0   'False
         Width           =   100
      End
      Begin VB.PictureBox sUnit 
         BackColor       =   &H000040C0&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   15
         Left            =   2400
         ScaleHeight     =   255
         ScaleWidth      =   105
         TabIndex        =   36
         Top             =   720
         Visible         =   0   'False
         Width           =   100
      End
      Begin VB.PictureBox sUnit 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   14
         Left            =   2280
         ScaleHeight     =   255
         ScaleWidth      =   105
         TabIndex        =   35
         Top             =   720
         Visible         =   0   'False
         Width           =   100
      End
      Begin VB.PictureBox sUnit 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   13
         Left            =   2160
         ScaleHeight     =   255
         ScaleWidth      =   105
         TabIndex        =   34
         Top             =   720
         Visible         =   0   'False
         Width           =   100
      End
      Begin VB.PictureBox sUnit 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   12
         Left            =   2040
         ScaleHeight     =   255
         ScaleWidth      =   105
         TabIndex        =   33
         Top             =   720
         Visible         =   0   'False
         Width           =   100
      End
      Begin VB.PictureBox sUnit 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   11
         Left            =   1920
         ScaleHeight     =   255
         ScaleWidth      =   105
         TabIndex        =   32
         Top             =   720
         Visible         =   0   'False
         Width           =   100
      End
      Begin VB.PictureBox sUnit 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   10
         Left            =   1800
         ScaleHeight     =   255
         ScaleWidth      =   105
         TabIndex        =   31
         Top             =   720
         Visible         =   0   'False
         Width           =   100
      End
      Begin VB.PictureBox sUnit 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   9
         Left            =   1680
         ScaleHeight     =   255
         ScaleWidth      =   105
         TabIndex        =   30
         Top             =   720
         Visible         =   0   'False
         Width           =   100
      End
      Begin VB.PictureBox sUnit 
         BackColor       =   &H0000C000&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   8
         Left            =   1560
         ScaleHeight     =   255
         ScaleWidth      =   105
         TabIndex        =   29
         Top             =   720
         Visible         =   0   'False
         Width           =   100
      End
      Begin VB.PictureBox sUnit 
         BackColor       =   &H0000C000&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   7
         Left            =   1440
         ScaleHeight     =   255
         ScaleWidth      =   105
         TabIndex        =   28
         Top             =   720
         Visible         =   0   'False
         Width           =   100
      End
      Begin VB.PictureBox sUnit 
         BackColor       =   &H0000C000&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   6
         Left            =   1320
         ScaleHeight     =   255
         ScaleWidth      =   105
         TabIndex        =   27
         Top             =   720
         Visible         =   0   'False
         Width           =   100
      End
      Begin VB.PictureBox sUnit 
         BackColor       =   &H0000C000&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   5
         Left            =   1200
         ScaleHeight     =   255
         ScaleWidth      =   105
         TabIndex        =   26
         Top             =   720
         Visible         =   0   'False
         Width           =   100
      End
      Begin VB.PictureBox sUnit 
         BackColor       =   &H0000C000&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   4
         Left            =   1080
         ScaleHeight     =   255
         ScaleWidth      =   105
         TabIndex        =   25
         Top             =   720
         Visible         =   0   'False
         Width           =   100
      End
      Begin VB.PictureBox sUnit 
         BackColor       =   &H0000C000&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   3
         Left            =   960
         ScaleHeight     =   255
         ScaleWidth      =   105
         TabIndex        =   23
         Top             =   720
         Visible         =   0   'False
         Width           =   100
      End
      Begin VB.PictureBox sUnit 
         BackColor       =   &H0000C000&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   2
         Left            =   840
         ScaleHeight     =   255
         ScaleWidth      =   105
         TabIndex        =   22
         Top             =   720
         Visible         =   0   'False
         Width           =   100
      End
      Begin VB.PictureBox sUnit 
         BackColor       =   &H0000C000&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   1
         Left            =   720
         ScaleHeight     =   255
         ScaleWidth      =   105
         TabIndex        =   21
         Top             =   720
         Visible         =   0   'False
         Width           =   100
      End
      Begin VB.PictureBox sUnit 
         BackColor       =   &H0000C000&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   0
         Left            =   600
         ScaleHeight     =   255
         ScaleWidth      =   105
         TabIndex        =   20
         Top             =   720
         Visible         =   0   'False
         Width           =   100
      End
      Begin VB.Label lblStep 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   3960
         TabIndex        =   102
         Top             =   825
         Width           =   975
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Step"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   3450
         TabIndex        =   101
         Top             =   825
         Width           =   420
      End
      Begin VB.Label lblLocal 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   255
         Left            =   5760
         TabIndex        =   100
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Local"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   255
         Left            =   4920
         TabIndex        =   99
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "UTC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   255
         Left            =   4920
         TabIndex        =   98
         Top             =   735
         Width           =   735
      End
      Begin VB.Label lblUTCtime 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   255
         Left            =   5760
         TabIndex        =   97
         Top             =   735
         Width           =   1215
      End
      Begin VB.Label lblAGC 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   3960
         TabIndex        =   96
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblFilter 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   3960
         TabIndex        =   95
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblMode 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   3960
         TabIndex        =   94
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "AGC"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   3450
         TabIndex        =   93
         Top             =   605
         Width           =   420
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Filter"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   3450
         TabIndex        =   92
         Top             =   355
         Width           =   420
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Mode"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   3450
         TabIndex        =   91
         Top             =   120
         Width           =   420
      End
      Begin VB.Label lblHz 
         BackStyle       =   0  'Transparent
         Caption         =   ".000"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   375
         Left            =   2040
         TabIndex        =   74
         Top             =   90
         Width           =   735
      End
      Begin VB.Label Label9 
         BackColor       =   &H00000000&
         Caption         =   "   +10      +20     +30     +40"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   135
         Left            =   1800
         TabIndex        =   66
         Top             =   570
         Width           =   1455
      End
      Begin VB.Label lblSubVFO 
         BackColor       =   &H00000000&
         Caption         =   "VFO B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   5040
         TabIndex        =   65
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lblMainVFO 
         BackColor       =   &H00000000&
         Caption         =   "VFO A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   100
         TabIndex        =   64
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "MHz"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   63
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "1  2  3  4  5  6  7  8  9"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   135
         Left            =   600
         TabIndex        =   42
         Top             =   570
         Width           =   1335
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "S Units"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   24
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblAltVFO 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "00.000.00 MHz"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   5640
         TabIndex        =   19
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label lblFreq 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "00.000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   600
         TabIndex        =   1
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H009C7D2C&
      Caption         =   "UP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0036FCFC&
      Height          =   255
      Left            =   4920
      TabIndex        =   82
      Top             =   1290
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H009C7D2C&
      Caption         =   "DOWN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0036FCFC&
      Height          =   255
      Left            =   1440
      TabIndex        =   81
      Top             =   1290
      Width           =   735
   End
   Begin VB.Label lblDSPver 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   0
      Width           =   3135
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSave 
         Caption         =   "Save Settings to File"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore Settings from File"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close RX320"
      End
   End
   Begin VB.Menu mnuMode 
      Caption         =   "&Mode"
      Begin VB.Menu mnuModeSel 
         Caption         =   "AM"
         Index           =   0
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuModeSel 
         Caption         =   "LSB"
         Index           =   1
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuModeSel 
         Caption         =   "USB"
         Index           =   2
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuModeSel 
         Caption         =   "CW"
         Index           =   3
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuFilter 
      Caption         =   "F&ilter"
      Begin VB.Menu mnuWidth 
         Caption         =   "IF Filter Width"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilterSel 
         Caption         =   "300"
         Index           =   0
      End
      Begin VB.Menu mnuFilterSel 
         Caption         =   "330"
         Index           =   1
      End
      Begin VB.Menu mnuFilterSel 
         Caption         =   "375"
         Index           =   2
      End
      Begin VB.Menu mnuFilterSel 
         Caption         =   "450"
         Index           =   3
      End
      Begin VB.Menu mnuFilterSel 
         Caption         =   "525"
         Index           =   4
      End
      Begin VB.Menu mnuFilterSel 
         Caption         =   "600"
         Index           =   5
      End
      Begin VB.Menu mnuFilterSel 
         Caption         =   "675"
         Index           =   6
      End
      Begin VB.Menu mnuFilterSel 
         Caption         =   "750"
         Index           =   7
      End
      Begin VB.Menu mnuFilterSel 
         Caption         =   "900"
         Index           =   8
      End
      Begin VB.Menu mnuFilterSel 
         Caption         =   "1050"
         Index           =   9
      End
      Begin VB.Menu mnuFilterSel 
         Caption         =   "1200"
         Index           =   10
      End
      Begin VB.Menu mnuFilterSel 
         Caption         =   "1350"
         Index           =   11
      End
      Begin VB.Menu mnuFilterSel 
         Caption         =   "1500"
         Index           =   12
      End
      Begin VB.Menu mnuFilterSel 
         Caption         =   "1650"
         Index           =   13
      End
      Begin VB.Menu mnuFilterSel 
         Caption         =   "1800"
         Index           =   14
      End
      Begin VB.Menu mnuFilterSel 
         Caption         =   "1950"
         Index           =   15
      End
      Begin VB.Menu mnuFilterSel 
         Caption         =   "2100"
         Index           =   16
      End
      Begin VB.Menu mnuFilterSel 
         Caption         =   "2250"
         Index           =   17
      End
      Begin VB.Menu mnuFilterSel 
         Caption         =   "2400"
         Index           =   18
      End
      Begin VB.Menu mnuFilterSel 
         Caption         =   "2550"
         Index           =   19
      End
      Begin VB.Menu mnuFilterSel 
         Caption         =   "2700"
         Index           =   20
      End
      Begin VB.Menu mnuFilterSel 
         Caption         =   "2850"
         Index           =   21
      End
      Begin VB.Menu mnuFilterSel 
         Caption         =   "3000"
         Index           =   22
      End
      Begin VB.Menu mnuFilterSel 
         Caption         =   "3300"
         Index           =   23
      End
      Begin VB.Menu mnuFilterSel 
         Caption         =   "3600"
         Index           =   24
      End
      Begin VB.Menu mnuFilterSel 
         Caption         =   "3900"
         Index           =   25
      End
      Begin VB.Menu mnuFilterSel 
         Caption         =   "4200"
         Index           =   26
      End
      Begin VB.Menu mnuFilterSel 
         Caption         =   "4500"
         Index           =   27
      End
      Begin VB.Menu mnuFilterSel 
         Caption         =   "4800"
         Index           =   28
      End
      Begin VB.Menu mnuFilterSel 
         Caption         =   "5100"
         Index           =   29
      End
      Begin VB.Menu mnuFilterSel 
         Caption         =   "5400"
         Index           =   30
      End
      Begin VB.Menu mnuFilterSel 
         Caption         =   "5700"
         Index           =   31
      End
      Begin VB.Menu mnuFilterSel 
         Caption         =   "6000"
         Index           =   32
      End
      Begin VB.Menu mnuFilterSel 
         Caption         =   "8000"
         Index           =   33
      End
   End
   Begin VB.Menu mnuProperties 
      Caption         =   "&Properties"
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Window"
      Begin VB.Menu mnuBasic 
         Caption         =   "Basic Window"
      End
      Begin VB.Menu mnuFull 
         Caption         =   "Full Window"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuShortKeys 
         Caption         =   "Shortcut Keys"
         Enabled         =   0   'False
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About RX320 Control"
         Shortcut        =   +{F1}
      End
   End
End
Attribute VB_Name = "frmRX320"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    '********************************************
    '
    '   TenTec RX320 Receiver Control Main Form
    '
    '   Mark Mokoski
    '   17-FEB-2005
    '
    '*********************************************

    Option Explicit
    
    Dim tuneUp              As Boolean 'Flag for repeat "UP" button
    Dim tuneDown            As Boolean 'Flag for repeat "DOWN" button
    

Private Sub cmdAequalB_Click()

    '
    'Copy contents of Active VFO into Sub VFO
    '

    DisplayPicture.SetFocus

End Sub

Private Sub cmdAtoB_Click()

    '
    'Copy contents of VFO A into VFO B
    '
    DisplayPicture.SetFocus

End Sub

Private Sub cmdBtoA_Click()

    '
    'Copy contents of VFO B into VFO A
    '

    DisplayPicture.SetFocus

End Sub

Private Sub cmdFilter_Click(Index As Integer)

    ComboFilter.Text = cmdFilter(Index).Tag
    DisplayPicture.SetFocus

    DoEvents

End Sub

Private Sub cmdVFO_Click(Index As Integer)

    '
    'Change active VFO (Swap)
    '
    
    'Get VFO to change to from button clicked (index)
    
        If Index = 0 Then
            'VFO A new active VFO
            Call MoveVFO(1, 0, 0)
        Else
            'VFO B new active VFO
            Call MoveVFO(0, 1, 0)
        End If
    
    SetDisplay
    DisplayPicture.SetFocus

End Sub

Private Sub ComboFilter_Change()

    SetFilter (Val(ComboFilter.Text))
    lblFilter.Caption = ComboFilter.Text & " Hz"
    DisplayPicture.SetFocus

    DoEvents

End Sub

Private Sub ComboFilter_Click()

    Dim x            As Integer

    SetFilter (Val(ComboFilter.Text))
    lblFilter.Caption = ComboFilter.Text & " Hz"

        For x = 0 To (ComboFilter.ListCount - 1)
            
                If ComboFilter.List(x) = ComboFilter.Text Then
                    'Find the matching menu item and check it
                    mnuFilterSel(x).Checked = True
                Else
                    'Clear non-matching menu items
                    mnuFilterSel(x).Checked = False
                End If

        Next x
        
    DisplayPicture.SetFocus

    DoEvents

End Sub

Private Sub ComboStep_Change()

    Call SetStep

End Sub

Private Sub ComboStep_Click()

    Call SetStep

End Sub

Private Sub Digit_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    Digit(Index).Appearance = 1
    Digit(Index).BackColor = &HFF00&
    AddFreqString (Index)

    DoEvents

End Sub

Private Sub Digit_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    Digit(Index).Appearance = 0
    Digit(Index).BackColor = &HC0C0C0

    DoEvents

End Sub

Private Sub DigitDP_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    DigitDP.Appearance = 1
    DigitDP.BackColor = &HFF00&
    strFreq = strFreq & "."
    lblNewFreq = strFreq

    DoEvents

End Sub

Private Sub DigitDP_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    DigitDP.Appearance = 0
    DigitDP.BackColor = &HC0C0C0

    DoEvents

End Sub

Private Sub DigitEnter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    DigitEnter.Appearance = 1
    DigitEnter.BackColor = &HFF00&

        If Val(strFreq) <= 30 And Val(strFreq) >= 0.5 Then
            valFreq = Val(strFreq) * 1000000
            SetFREQ (valFreq)
            lblNewFreq = "Enter Frequency in MHz"
            strFreq = ""
            SetDisplay
        Else
            lblNewFreq = "ERROR"
            strFreq = ""
        End If
    
    DoEvents
   
End Sub

Private Sub DigitEnter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    DigitEnter.Appearance = 0
    DigitEnter.BackColor = &HC0C0C0

    DoEvents

End Sub

Private Sub DisplayPicture_KeyDown(KeyCode As Integer, Shift As Integer)

    '
    '   Encode new frequency from key pad entry
    '

        Select Case KeyCode
            Case vbKeyNumpad0
                Digit(0).Appearance = 1
                Digit(0).BackColor = &HFF00&
                AddFreqString (0)
            Case vbKeyNumpad1
                Digit(1).Appearance = 1
                Digit(1).BackColor = &HFF00&
                AddFreqString (1)
            Case vbKeyNumpad2
                Digit(2).Appearance = 1
                Digit(2).BackColor = &HFF00&
                AddFreqString (2)
            Case vbKeyNumpad3
                Digit(3).Appearance = 1
                Digit(3).BackColor = &HFF00&
                AddFreqString (3)
            Case vbKeyNumpad4
                Digit(4).Appearance = 1
                Digit(4).BackColor = &HFF00&
                AddFreqString (4)
            Case vbKeyNumpad5
                Digit(5).Appearance = 1
                Digit(5).BackColor = &HFF00&
                AddFreqString (5)
            Case vbKeyNumpad6
                Digit(6).Appearance = 1
                Digit(6).BackColor = &HFF00&
                AddFreqString (6)
            Case vbKeyNumpad7
                Digit(7).Appearance = 1
                Digit(7).BackColor = &HFF00&
                AddFreqString (7)
            Case vbKeyNumpad8
                Digit(8).Appearance = 1
                Digit(8).BackColor = &HFF00&
                AddFreqString (8)
            Case vbKeyNumpad9
                Digit(9).Appearance = 1
                Digit(9).BackColor = &HFF00&
                AddFreqString (9)
            Case vbKeyDecimal
                DigitDP.Appearance = 1
                DigitDP.BackColor = &HFF00&
                strFreq = strFreq & "."
                lblNewFreq = strFreq
                
                'Enter Key = set new freq
            Case vbKeyReturn
                DigitEnter.Appearance = 1
                DigitEnter.BackColor = &HFF00&

                If Val(strFreq) <= 30 And Val(strFreq) >= 0.5 Then
                    valFreq = Val(strFreq) * 1000000
                    SetFREQ (valFreq)
                    lblNewFreq = "Enter Frequency in MHz"
                    strFreq = ""
                    SetDisplay
                Else
                    lblNewFreq = "ERROR"
                    strFreq = ""
                End If
                
            'Delete Key = clear current entry
            Case vbKeyDelete
                EntryClear.Appearance = 1
                EntryClear.BackColor = &HFF00&

                lblNewFreq = "Enter Frequency in MHz"
                strFreq = ""
                
                'Backspace Key = remove last didgit entered
            Case vbKeyBack

                If lblNewFreq <> "Enter Frequency in MHz" Then
                    lblNewFreq = Mid(lblNewFreq, 1, Len(lblNewFreq) - 1)

                        If lblNewFreq = "" Then
                            lblNewFreq = "Enter Frequency in MHz"
                            strFreq = ""
                        Else
                            strFreq = lblNewFreq
                        End If

                End If

            'Right arrow = freq up
            Case vbKeyRight
                valFreq = valFreq + valStep
                SetDisplay
                SetFREQ (valFreq)
                DoEvents
                'Left arrow = freq down
            Case vbKeyLeft
                valFreq = valFreq - valStep
                SetDisplay
                SetFREQ (valFreq)
                DoEvents
            
                'Up arrow = spk vol up
            Case vbKeyUp

                If volSPK > 1 Then
                    SpkVol.Value = volSPK - 1
                    SetSPKvol (SpkVol.Value)
                    volSPK = SpkVol.Value
                    DoEvents
                End If

            'Down arrow = spk vol down
            Case vbKeyDown

                If volSPK < 63 Then
                    SpkVol.Value = volSPK + 1
                    SetSPKvol (SpkVol.Value)
                    volSPK = SpkVol.Value
                    DoEvents
                End If

        End Select
    
    DoEvents

End Sub

Private Sub DisplayPicture_KeyUp(KeyCode As Integer, Shift As Integer)

    '
    '   Reset "Key Pad" background color
    '

    Dim x            As Integer

        For x = 0 To 9
            Digit(x).Appearance = 0
            Digit(x).BackColor = &HC0C0C0
        Next x

    DigitDP.Appearance = 0
    DigitDP.BackColor = &HC0C0C0
    DigitEnter.Appearance = 0
    DigitEnter.BackColor = &HC0C0C0
    EntryClear.Appearance = 0
    EntryClear.BackColor = &HC0C0C0

    DoEvents

End Sub

Private Sub EntryClear_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    '
    '   Clear current frequency entry
    '
    EntryClear.Appearance = 1
    EntryClear.BackColor = &HFF00&

    lblNewFreq = "Enter Frequency in MHz"
    strFreq = ""

    DoEvents

End Sub

Private Sub EntryClear_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    EntryClear.Appearance = 0
    EntryClear.BackColor = &HC0C0C0
    DisplayPicture.SetFocus
    
    DoEvents
  
End Sub

Private Sub Form_Load()

    '********* Set up the whole mess on form load **********
    
    'Last size of the main form, "Full (1)" or "Basic (0)"
    WindowSize = GetSetting("RX320", "General", "WindowSize", 1)
    
    Me.Visible = True

        If WindowSize = 1 Then
            Me.Height = 7380
            mnuFull.Checked = True
            mnuBasic.Checked = False
        Else
            frmRX320.Top = frmRX320.Top - 500
            Me.Height = 2520
            mnuFull.Checked = False
            mnuBasic.Checked = True
        End If

    ' Me.Visible = True
    Me.Caption = App.Title & "  -  Version " & App.Major & "." & App.Minor & "." & App.Revision
   
    'Set Volume and squelch levels
    SpkVol.Value = 63
    LineVol.Value = 63

    'Update UTC Clock display
    lblUTCtime.Caption = UTCtime & " UTC"

    'Start communcation to radio
    RadioComPort = GetSetting("RX320", "General", "RadioComPort", 1)
    RadioOK = False

    RadioCOM.CommPort = RadioComPort
    RadioCOM.RThreshold = 1
    RadioCOM.Settings = "1200,N,8,1"
    RadioCOM.PortOpen = True
    'Mute radio until all setup is done
    Mute.Value = 1
    'Get DSP Version
    RadioCOM.Output = "?" & vbCr
    'Get S Meter to start
    RadioCOM.Output = "X" & vbCr

    DoEvents

    'Start by setting up the form controls
    ComboFilter.AddItem "300"
    ComboFilter.AddItem "330"
    ComboFilter.AddItem "375"
    ComboFilter.AddItem "450"
    ComboFilter.AddItem "525"
    ComboFilter.AddItem "600"
    ComboFilter.AddItem "675"
    ComboFilter.AddItem "750"
    ComboFilter.AddItem "900"
    ComboFilter.AddItem "1050"
    ComboFilter.AddItem "1200"
    ComboFilter.AddItem "1350"
    ComboFilter.AddItem "1500"
    ComboFilter.AddItem "1650"
    ComboFilter.AddItem "1800"
    ComboFilter.AddItem "1950"
    ComboFilter.AddItem "2100"
    ComboFilter.AddItem "2250"
    ComboFilter.AddItem "2400"
    ComboFilter.AddItem "2550"
    ComboFilter.AddItem "2700"
    ComboFilter.AddItem "2850"
    ComboFilter.AddItem "3000"
    ComboFilter.AddItem "3300"
    ComboFilter.AddItem "3600"
    ComboFilter.AddItem "3900"
    ComboFilter.AddItem "4200"
    ComboFilter.AddItem "4500"
    ComboFilter.AddItem "4800"
    ComboFilter.AddItem "5100"
    ComboFilter.AddItem "5400"
    ComboFilter.AddItem "5700"
    ComboFilter.AddItem "6000"
    ComboFilter.AddItem "8000"
       
    DoEvents

    '**********************************
    'Load last radio stup from registry
    '**********************************

    'Last selected filter
    valFilter = GetSetting("RX320", "General", "valFilter", 6000)
    ComboFilter.Text = valFilter
    
    'Last selected AGC time constant
    valAGC = GetSetting("RX320", "General", "valAGC", 2)

        Select Case valAGC
            Case 1
                OptionAGC_Click (2)
            Case 2
                OptionAGC_Click (1)
            Case 3
                OptionAGC_Click (0)
        End Select
    
    'Get the stored Mode settings
    AMfilter = GetSetting("RX320", "AM", "AMfilter", 6000)      'Last selected AMfilter
    AMstep = GetSetting("RX320", "AM", "AMstep", "10 KHz")      'AM tuning step
    SSBfilter = GetSetting("RX320", "SSB", "SSBfilter", 3000)   'Last selected SSB filter
    SSBstep = GetSetting("RX320", "SSB", "SSBstep", "100 Hz")   'SSB tuning step
    LSBoffset = GetSetting("RX320", "SSB", "LSBoffset", 0)      'LSB/CW tuning correction
    USBoffset = GetSetting("RX320", "SSB", "USBoffset", 0)      'USB tuning correction
    CWfilter = GetSetting("RX320", "CW", "CWfilter", 1200)      'Last selected CW filter
    CWstep = GetSetting("RX320", "CW", "CWstep", "10 Hz")       'CW tuning step

    Dim x            As Integer

        For x = 0 To 5
            AMquickfilter(x) = GetSetting("RX320", "AM", "AMquickfilter_" & x, 6000)    'AM Quick filter selections
            SSBquickfilter(x) = GetSetting("RX320", "SSB", "SSBquickfilter_" & x, 3000) 'SSB Quick filter selections
            CWquickfilter(x) = GetSetting("RX320", "CW", "CWquickfilter_" & x, 1200)    'CW Quick filter selection
        Next x

    DoEvents
    
    'More current settings
    valModeCor = GetSetting("RX320", "General", "valModeCor", 0)
    valAdjFreq = GetSetting("RX320", "General", "valAdjFreq", 0)
    valCWOffSet = GetSetting("RX320", "General", "valCWOffSet", 800)
    SetCWOFFSET (valCWOffSet)
    
    'Last frequency
    valFreq = GetSetting("RX320", "General", "valFreq", 1080000)
    SetFREQ (valFreq)
    
    'Freq Step Settings
    show1Hz = GetSetting("RX320", "General", "show1Hz", False)

        If show1Hz = True Then
            ComboStep.AddItem "1 Hz"
        End If

    ComboStep.AddItem "10 Hz"
    ComboStep.AddItem "100 Hz"
    ComboStep.AddItem "1 KHz"
    ComboStep.AddItem "2.5 KHz"
    ComboStep.AddItem "5 KHz"
    ComboStep.AddItem "9 KHz"
    ComboStep.AddItem "10 KHz"
    ComboStep.AddItem "100 KHz"
    ComboStep.AddItem "1 MHz"
    
    'Last tuning step
    
    valStep = GetSetting("RX320", "General", "valStep", 10000)

        Select Case valStep
            Case 1
                ComboStep.Text = "1 Hz"
            Case 10
                ComboStep.Text = "10 Hz"
            Case 100
                ComboStep.Text = "100 Hz"
            Case 1000
                ComboStep.Text = "1 KHz"
            Case 2500
                ComboStep.Text = "2.5 KHz"
            Case 5000
                ComboStep.Text = "5 KHz"
            Case 9000
                ComboStep.Text = "9 KHz"
            Case 10000
                ComboStep.Text = "10 KHz"
            Case 100000
                ComboStep.Text = "100 KHz"
            Case 1000000
                ComboStep.Text = " 1 MHz"
        End Select
    
    'Last selected Mode

    valMode = GetSetting("RX320", "General", "valMode", "AM")

        Select Case valMode
            Case "AM"
                OptionMode_Click (0)
            Case "LSB"
                OptionMode_Click (1)
            Case "USB"
                OptionMode_Click (2)
            Case "CW"
                OptionMode_Click (3)
        End Select
    
    DoEvents
            
    'Mute On Exit flag
    muteOnExit = GetSetting("RX320", "General", "muteOnExit", False)

    'Get stored VFO settings and display "off line" VFO
    'VFO A = 0, VFO B = 1
    
    'Current "Active VFO
    CurrentVFO = GetSetting("RX320", "General", "CurrentVFO", 0)
    
    'VFO A
    freqVFO(0) = GetSetting("RX320", "VFOA", "freqVFO", 1080000)
    VFOFilter(0) = GetSetting("RX320", "VFOA", "VFOFilter", 6000)
    VFOAGC(0) = GetSetting("RX320", "VFOA", "VFOAGC", 2)
    VFOMode(0) = GetSetting("RX320", "VFOA", "VFOMode", "AM")
    VFOModeCor(0) = GetSetting("RX320", "VFOA", "VFOModeCor", 0)
    VFOFreq(0) = GetSetting("RX320", "VFOA", "VFOFreq", 1080000)
    VFOAdjFreq(0) = GetSetting("RX320", "VFOA", "VFOAdjFreq", 0)
    VFOCWOffSet(0) = GetSetting("RX320", "VFOA", "VFOCWOffSet", 800)
    VFOStep(0) = GetSetting("RX320", "VFOA", "VFOStep", 5000)
    VFOBFO(0) = GetSetting("RX320", "VFOA", "VFOBFO", 0)

    'VFO B
    freqVFO(1) = GetSetting("RX320", "VFOB", "freqVFO", 10000000)
    VFOFilter(1) = GetSetting("RX320", "VFOB", "VFOFilter", 6000)
    VFOAGC(1) = GetSetting("RX320", "VFOB", "VFOAGC", 2)
    VFOMode(1) = GetSetting("RX320", "VFOB", "VFOMode", "AM")
    VFOModeCor(1) = GetSetting("RX320", "VFOB", "VFOModeCor", 0)
    VFOFreq(1) = GetSetting("RX320", "VFOB", "VFOFreq", 10000000)
    VFOAdjFreq(1) = GetSetting("RX320", "VFOB", "VFOAdjFreq", 0)
    VFOCWOffSet(1) = GetSetting("RX320", "VFOB", "VFOCWOffSet", 800)
    VFOStep(1) = GetSetting("RX320", "VFOB", "VFOStep", 5000)
    VFOBFO(1) = GetSetting("RX320", "VFOB", "VFOBFO", 0)

    DoEvents
    
    SetDisplay
    
    'unMute radio now with setup complete
    FixLineLevel = GetSetting("RX320", "General", "FixLineLevel", False)
    valFixedLevel = GetSetting("RX320", "General", "valFixedLevel", 40)

    Mute.Value = 0
    volSPK = GetSetting("RX320", "General", "volSPK", 63)
    SpkVol.Value = volSPK
    volLINE = GetSetting("RX320", "General", "volLINE", 63)
    LineVol.Value = volLINE
    
    
    'If fixed line level, set it and disable control

        If FixLineLevel = True Then
            LineVol.Value = valFixedLevel
            LineVol.Enabled = False
        Else
            LineVol.Enabled = True
        End If


    'Set repeat tuning flags to "False"
    tuneUp = False
    tuneDown = False
    
    DoEvents

End Sub


Private Sub Form_Terminate()

    'Close Properties form if active
    Unload frmProperties
    'Close frmAbout if active
    Unload frmAbout

    Call SaveRegSettings

End Sub

Private Sub Form_Unload(Cancel As Integer)

    'Close Properties form if active
    Unload frmProperties
    'Close frmAbout if active
    Unload frmAbout

    'Mute RX320 audio on exit

    Dim responce            As Integer

        If muteOnExit = True Then
            SetSPKvol (63)
            SetLINEvol (63)
        Else
            responce = MsgBox("Program Ending...." & vbCrLf & "Mute RX320?", vbApplicationModal + vbInformation + vbYesNo, "Program Ending")

                If responce = vbYes Then
                    SetSPKvol (63)
                    SetLINEvol (63)
                End If

            DoEvents

        End If

    Call SaveRegSettings

End Sub

Private Sub FreqDown_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    valFreq = valFreq - valStep
    SetDisplay
    SetFREQ (valFreq)
    'Set tune repeat delay timer
    TimerTuneDelay.Enabled = True
    tuneDown = True

    DoEvents

End Sub

Private Sub FreqDown_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    tuneDown = False
    TimerTuneDelay.Enabled = False
    TimerTuneRepeat.Enabled = False
    DisplayPicture.SetFocus
    
    DoEvents

End Sub

Private Sub FreqUp_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    valFreq = valFreq + valStep
    SetDisplay
    SetFREQ (valFreq)
    'Set repeat tune delay timer
    TimerTuneDelay.Enabled = True
    tuneUp = True

    DoEvents

End Sub

Private Sub FreqUp_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    tuneUp = False
    TimerTuneDelay.Enabled = False
    TimerTuneRepeat.Enabled = False
    DisplayPicture.SetFocus
   
    DoEvents

End Sub

Private Sub LineVol_Change()

    SetLINEvol (LineVol.Value)
    volLINE = LineVol.Value
    
    DoEvents

End Sub

Private Sub LineVol_Scroll()

    SetLINEvol (LineVol.Value)
    volLINE = LineVol.Value
    DisplayPicture.SetFocus

    DoEvents

End Sub


Private Sub MEMtoVFO_Click()

    '
    'Copy contents of memory to the active VFO
    '

    DisplayPicture.SetFocus

End Sub

Private Sub mnuAbout_Click()

    'display About form
    frmAbout.Visible = True

End Sub

Private Sub mnuBasic_Click()

    'Set to "Basic Window" view of from
    WindowSize = 0
    mnuFull.Checked = False
    mnuBasic.Checked = True
    Me.Height = 2520

End Sub

Private Sub mnuClose_Click()

    Unload Me

End Sub

Private Sub mnuFilterSel_Click(Index As Integer)

    'Set the filter width
    ComboFilter.Text = mnuFilterSel(Index).Caption

End Sub

Private Sub mnuFull_Click()

    'Set to "Full Window" view of from
    WindowSize = 1
    mnuFull.Checked = True
    mnuBasic.Checked = False
    Me.Height = 7380

End Sub

Private Sub mnuModeSel_Click(Index As Integer)

    OptionMode_Click (Index)

End Sub

Private Sub mnuProperties_Click()

    Load frmProperties

End Sub

Private Sub Mute_Click()

    '
    '   Mute Speaker and Line Out audio
    '   Unchecked, Audio output enabled
    '   Checked, Audio Muted
    '

        Select Case Mute.Value
            Case 0  'Audio Enabled
                SetSPKvol (volSPK)
                SpkVol.Enabled = True
                SetLINEvol (volLINE)

                If FixLineLevel = True Then
                    LineVol.Enabled = False
                End If

            Case 1  'Audio Muted
                SetSPKvol (63)
                SpkVol.Enabled = False
                SetLINEvol (63)
                LineVol.Enabled = False
        End Select

    DisplayPicture.SetFocus
    DoEvents

End Sub

Private Sub OptionAGC_Click(Index As Integer)

    Dim x                   As Integer

    SetAGC (OptionAGC(Index).Caption)
    PicAGC(Index).Picture = LoadResPicture(101, 1)

        For x = 0 To 2

                If x <> Index Then
                    PicAGC(x).Picture = Nothing
                End If

        Next x

        Select Case Index
            Case 0
                lblAGC.Caption = "Fast"
            Case 1
                lblAGC.Caption = "Medium"
            Case 2
                lblAGC.Caption = "Slow"
        End Select

    DisplayPicture.SetFocus
    DoEvents

End Sub

Private Sub OptionMode_Click(Index As Integer)

    Dim x                   As Integer

    SetMode (OptionMode(Index).Caption)
    PicMode(Index).Picture = LoadResPicture(101, 1)
    mnuModeSel(Index).Checked = True

        For x = 0 To 3

                If x <> Index Then
                    PicMode(x).Picture = Nothing
                    mnuModeSel(x).Checked = False
                End If

        Next x

    DoEvents
 
        Select Case Index
            Case 0
                lblMode.Caption = "AM"
                valAdjFreq = 0
                ComboStep.Text = AMstep
                ComboFilter.Text = AMfilter
                cmdFilter(0).Tag = AMquickfilter(0)
                cmdFilter(0).Caption = cmdFilter(0).Tag & " Hz"
                cmdFilter(1).Tag = AMquickfilter(1)
                cmdFilter(1).Caption = cmdFilter(1).Tag & " Hz"
                cmdFilter(2).Tag = AMquickfilter(2)
                cmdFilter(2).Caption = cmdFilter(2).Tag & " Hz"
                cmdFilter(3).Tag = AMquickfilter(3)
                cmdFilter(3).Caption = cmdFilter(3).Tag & " Hz"
                cmdFilter(4).Tag = AMquickfilter(4)
                cmdFilter(4).Caption = cmdFilter(4).Tag & " Hz"
                cmdFilter(5).Tag = AMquickfilter(5)
                cmdFilter(5).Caption = cmdFilter(5).Tag & " Hz"
            Case 1
                lblMode.Caption = "LSB"
                valAdjFreq = LSBoffset
                ComboStep.Text = SSBstep
                ComboFilter.Text = SSBfilter
                cmdFilter(0).Tag = SSBquickfilter(0)
                cmdFilter(0).Caption = cmdFilter(0).Tag & " Hz"
                cmdFilter(1).Tag = SSBquickfilter(1)
                cmdFilter(1).Caption = cmdFilter(1).Tag & " Hz"
                cmdFilter(2).Tag = SSBquickfilter(2)
                cmdFilter(2).Caption = cmdFilter(2).Tag & " Hz"
                cmdFilter(3).Tag = SSBquickfilter(3)
                cmdFilter(3).Caption = cmdFilter(3).Tag & " Hz"
                cmdFilter(4).Tag = SSBquickfilter(4)
                cmdFilter(4).Caption = cmdFilter(4).Tag & " Hz"
                cmdFilter(5).Tag = SSBquickfilter(5)
                cmdFilter(5).Caption = cmdFilter(5).Tag & " Hz"

            Case 2
                lblMode.Caption = "USB"
                valAdjFreq = USBoffset
                ComboStep.Text = SSBstep
                ComboFilter.Text = SSBfilter
                cmdFilter(0).Tag = SSBquickfilter(0)
                cmdFilter(0).Caption = cmdFilter(0).Tag & " Hz"
                cmdFilter(1).Tag = SSBquickfilter(1)
                cmdFilter(1).Caption = cmdFilter(1).Tag & " Hz"
                cmdFilter(2).Tag = SSBquickfilter(2)
                cmdFilter(2).Caption = cmdFilter(2).Tag & " Hz"
                cmdFilter(3).Tag = SSBquickfilter(3)
                cmdFilter(3).Caption = cmdFilter(3).Tag & " Hz"
                cmdFilter(4).Tag = SSBquickfilter(4)
                cmdFilter(4).Caption = cmdFilter(4).Tag & " Hz"
                cmdFilter(5).Tag = SSBquickfilter(5)
                cmdFilter(5).Caption = cmdFilter(5).Tag & " Hz"

            Case 3
                lblMode.Caption = "CW"
                valAdjFreq = LSBoffset
                ComboStep.Text = CWstep
                ComboFilter.Text = CWfilter
                cmdFilter(0).Tag = CWquickfilter(0)
                cmdFilter(0).Caption = cmdFilter(0).Tag & " Hz"
                cmdFilter(1).Tag = CWquickfilter(1)
                cmdFilter(1).Caption = cmdFilter(1).Tag & " Hz"
                cmdFilter(2).Tag = CWquickfilter(2)
                cmdFilter(2).Caption = cmdFilter(2).Tag & " Hz"
                cmdFilter(3).Tag = CWquickfilter(3)
                cmdFilter(3).Caption = cmdFilter(3).Tag & " Hz"
                cmdFilter(4).Tag = CWquickfilter(4)
                cmdFilter(4).Caption = cmdFilter(4).Tag & " Hz"
                cmdFilter(5).Tag = CWquickfilter(5)
                cmdFilter(5).Caption = cmdFilter(5).Tag & " Hz"

        End Select

    DisplayPicture.SetFocus
    DoEvents
 
End Sub

Private Sub RadioCOM_OnComm()

    On Error Resume Next

    Dim inString            As String
    
    inString = RadioCOM.Input

        If inString <> "" Then

                Select Case Mid(inString, 1, 1)
                    Case "X"
                        SetSMeter (inString)
                    Case "V"
                        DSPver = Str(Val(Mid(inString, 4, Len(inString) - 1)) / 100)
                End Select

        End If
        
    'Start S Meter polling
    TimerSunit.Enabled = True
    DoEvents

End Sub

Private Sub SpkVol_Change()

    SetSPKvol (SpkVol.Value)
    volSPK = SpkVol.Value
    DisplayPicture.SetFocus

    DoEvents

End Sub

Private Sub SpkVol_Scroll()

    SetSPKvol (SpkVol.Value)
    volSPK = SpkVol.Value
    DisplayPicture.SetFocus

    DoEvents

End Sub

Private Sub TimerClock_Timer()

    '
    '   Update UTC Clock display based on timer polling
    '

    lblUTCtime.Caption = UTCtime
    lblLocal.Caption = Format(Time$, "HH:MM:SS")

End Sub

Private Sub TimerSunit_Timer()

    '
    '   Get "S" mete value based on timer polling
    '
    TimerSunit.Enabled = False
    RadioCOM.Output = "X" & vbCr
    DoEvents

End Sub

Public Function SetDisplay()

    '
    '   Sets the main display frequency in the format MM.KKK.HH
    '   where MM is MHz, KKK is KHz, HH is Hz x 10
    '

    Dim Mhz                   As String
    Dim KHz                   As String
    Dim Hz                    As String
    Dim strDisplay            As String
    Dim AltVFO                As Integer
    
    'Set Active VFO Display Frequency
    'Get a fixed length string to work with
    '**Code_Err: Unused variables
    '**Code_Err: AltVFO

    strDisplay = Format(Str(valFreq), "00000000")
    'Extract the frequency decades
    Mhz = Mid$(strDisplay, 1, 2)
    KHz = Mid$(strDisplay, 3, 3)
    
    'If "Show 1Hz" is true, main display frequency in the format MM.KKK.HHH

        If show1Hz = True Then
            Hz = Mid$(strDisplay, 6, 3)
        Else
            Hz = Mid$(strDisplay, 6, 2)
        End If

    'Now put it together in the display format
    strDisplay = Format(Mhz, "##") & "." & KHz
    lblFreq = strDisplay
    lblHz = "." & Hz
    
    'Set Inactive VFO Display Frequency
    'Get a fixed length string to work with
    
        If CurrentVFO = 0 Then
            AltVFO = 1 'VFO B
            lblMainVFO = "VFO A"
            lblSubVFO = "VFO B"
            cmdVFO(0).Enabled = False
            cmdVFO(1).Enabled = True
        Else
            AltVFO = 0 'VFO A
            lblMainVFO = "VFO B"
            lblSubVFO = "VFO A"
            cmdVFO(0).Enabled = True
            cmdVFO(1).Enabled = False
        End If
    
    strDisplay = Format(Str(VFOFreq(AltVFO)), "00000000")
    'Extract the frequency decades
    Mhz = Mid$(strDisplay, 1, 2)
    KHz = Mid$(strDisplay, 3, 3)
    Hz = Mid$(strDisplay, 6, 2)

    'Now put it together in the display format
    strDisplay = Format(Mhz, "##") & "." & KHz & "." & Hz & " MHz"
    lblAltVFO = strDisplay
    
    DoEvents
    
End Function

Private Function AddFreqString(Digit As Integer)

    strFreq = strFreq & Format(Str(Digit), "###0")
    lblNewFreq = strFreq

    DoEvents

End Function

Private Sub TimerTuneDelay_Timer()

    TimerTuneDelay.Enabled = False

        If valStep <= 5000 Then
            'Adjust repaet tune speed based on frequency step
            TimerTuneRepeat.Interval = (valStep / 10)
        Else
            'If set greater than 5 KHz, fis repeat tune to .5 sec
            TimerTuneRepeat.Interval = 500
        End If

        If TimerTuneRepeat.Interval = 0 Then TimerTuneRepeat.Interval = 1

    TimerTuneRepeat.Enabled = True

End Sub

Private Sub TimerTuneRepeat_Timer()

    'Repeat "UP" tune

        If tuneUp = True Then
            '"UP" frequency button clicked
            valFreq = valFreq + valStep
            SetDisplay
            SetFREQ (valFreq)

        End If
        
    'Repeat "DOWN" tune

        If tuneDown = True Then
            '"Down" frequency button clicked
            valFreq = valFreq - valStep
            SetDisplay
            SetFREQ (valFreq)
        End If

End Sub

Private Sub VFOtoMEM_Click()

    '
    'Copy active VFO contents into memory
    '

    DisplayPicture.SetFocus

End Sub

Private Sub SetStep()

    '
    'Set new step value and round off to nearest set value
    '

        Select Case ComboStep.Text
            Case "1 Hz"
                valStep = 1
            Case "10 Hz"
                valStep = 10
            Case "100 Hz"
                valStep = 100
            Case "1 KHz"
                valStep = 1000
            Case "2.5 KHz"
                valStep = 2500
            Case "5 KHz"
                valStep = 5000
            Case "9 KHz"
                valStep = 9000
            Case "10 KHz"
                valStep = 10000
            Case "100 KHz"
                valStep = 100000
            Case "1 MHz"
                valStep = 1000000
        End Select

    lblStep.Caption = ComboStep.Text
    'Change frequency by the step value,
    'Round up or down to the next step value
    'Set the radio frequency and display

    Dim tempFreq              As Long

    tempFreq = Round(valFreq / valStep)
    valFreq = tempFreq * valStep
    SetDisplay
    SetFREQ (valFreq)

    DisplayPicture.SetFocus
    
    DoEvents

End Sub

Private Sub MoveVFO(sourceVFO As Integer, destVFO As Integer, Operation As Integer, Optional MemCH As Integer)

    '
    'Move contents of VFO's
    '
    'VFO values
    '0 = VFO A
    '1 = VFO B
    '2 = Memory Location (Channel)
    '
    'Operation values
    '0 = Swap VFO's
    '1 = Copy VFO A to VFO B
    '2 = Copy VFO B to VFO A
    '
    'Memory Channel (Optional param need for memory operations)
    'Value Range = 0 to 99
    '
    
    'If Operation was "Swap", set new active VFO
    If Operation = 0 Then CurrentVFO = destVFO

End Sub
