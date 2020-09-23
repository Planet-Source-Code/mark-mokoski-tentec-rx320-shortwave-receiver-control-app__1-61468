VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmProperties 
   BackColor       =   &H00B18E54&
   Caption         =   "RX320 Properties"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6375
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00000000&
   Icon            =   "frmProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   6375
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab tabProperties 
      Height          =   3975
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   7011
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   11636308
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmProperties.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblUTCoffset"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblTZinfo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkMuteOnExit"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chkFixedLevel"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "chk1Hz"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "listFixedLevel"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "AM"
      TabPicture(1)   =   "frmProperties.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4(0)"
      Tab(1).Control(1)=   "Frame3(0)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "SSB"
      TabPicture(2)   =   "frmProperties.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4(1)"
      Tab(2).Control(1)=   "Frame3(1)"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "CW"
      TabPicture(3)   =   "frmProperties.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4(2)"
      Tab(3).Control(1)=   "Frame3(2)"
      Tab(3).ControlCount=   2
      Begin VB.ComboBox listFixedLevel 
         Height          =   315
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   80
         Top             =   2280
         Width           =   975
      End
      Begin VB.CheckBox chk1Hz 
         Caption         =   "Show 1Hz Display Digit"
         Height          =   255
         Left            =   3240
         TabIndex        =   48
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Frame Frame4 
         Caption         =   "Mode Parameters"
         Height          =   3375
         Index           =   2
         Left            =   -71760
         TabIndex        =   35
         Top             =   480
         Width           =   2655
         Begin VB.ComboBox comboCWoffset 
            Height          =   315
            Left            =   1440
            TabIndex        =   47
            Text            =   "Combo4"
            Top             =   1440
            Width           =   1095
         End
         Begin VB.ComboBox ComboDefaultStep 
            Height          =   315
            Index           =   2
            Left            =   1440
            TabIndex        =   37
            Text            =   "Combo2"
            Top             =   480
            Width           =   1095
         End
         Begin VB.ComboBox comboDefaultFilter 
            Height          =   315
            Index           =   2
            Left            =   1440
            TabIndex        =   36
            Text            =   "Combo3"
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "CW Off Set in Hz"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   1490
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Default Freq Sep"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   39
            Top             =   510
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Default Filter"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   38
            Top             =   1000
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Quick Filter Selection"
         Height          =   3375
         Index           =   2
         Left            =   -74880
         TabIndex        =   34
         Top             =   480
         Width           =   2895
         Begin VB.CommandButton cmdCWFilter 
            BackColor       =   &H00C0C0C0&
            Caption         =   "450 Hz"
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
            Left            =   240
            MouseIcon       =   "frmProperties.frx":037A
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   79
            Tag             =   "450"
            Top             =   2760
            Width           =   855
         End
         Begin VB.CommandButton cmdCWFilter 
            BackColor       =   &H00C0C0C0&
            Caption         =   "600 Hz"
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
            Left            =   240
            MouseIcon       =   "frmProperties.frx":0684
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   78
            Tag             =   "600"
            Top             =   2280
            Width           =   855
         End
         Begin VB.CommandButton cmdCWFilter 
            BackColor       =   &H00C0C0C0&
            Caption         =   "750 Hz"
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
            Left            =   240
            MouseIcon       =   "frmProperties.frx":098E
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   77
            Tag             =   "750"
            Top             =   1800
            Width           =   855
         End
         Begin VB.CommandButton cmdCWFilter 
            BackColor       =   &H00C0C0C0&
            Caption         =   "900 Hz"
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
            Left            =   240
            MouseIcon       =   "frmProperties.frx":0C98
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   76
            Tag             =   "900"
            Top             =   1320
            Width           =   855
         End
         Begin VB.CommandButton cmdCWFilter 
            BackColor       =   &H00C0C0C0&
            Caption         =   "1200 Hz"
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
            Left            =   240
            MouseIcon       =   "frmProperties.frx":0FA2
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   75
            Tag             =   "1200"
            Top             =   840
            Width           =   855
         End
         Begin VB.CommandButton cmdCWFilter 
            BackColor       =   &H00C0C0C0&
            Caption         =   "1500 Hz"
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
            Left            =   240
            MouseIcon       =   "frmProperties.frx":12AC
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   74
            Tag             =   "1500"
            Top             =   360
            Width           =   855
         End
         Begin VB.ComboBox ComboCWfilter 
            Height          =   315
            Index           =   0
            Left            =   1200
            TabIndex        =   66
            Text            =   "Combo1"
            Top             =   480
            Width           =   1335
         End
         Begin VB.ComboBox ComboCWfilter 
            Height          =   315
            Index           =   1
            Left            =   1200
            TabIndex        =   65
            Text            =   "Combo1"
            Top             =   960
            Width           =   1335
         End
         Begin VB.ComboBox ComboCWfilter 
            Height          =   315
            Index           =   2
            Left            =   1200
            TabIndex        =   64
            Text            =   "Combo1"
            Top             =   1440
            Width           =   1335
         End
         Begin VB.ComboBox ComboCWfilter 
            Height          =   315
            Index           =   3
            Left            =   1200
            TabIndex        =   63
            Text            =   "Combo1"
            Top             =   1920
            Width           =   1335
         End
         Begin VB.ComboBox ComboCWfilter 
            Height          =   315
            Index           =   4
            Left            =   1200
            TabIndex        =   62
            Text            =   "Combo1"
            Top             =   2400
            Width           =   1335
         End
         Begin VB.ComboBox ComboCWfilter 
            Height          =   315
            Index           =   5
            Left            =   1200
            TabIndex        =   61
            Text            =   "Combo1"
            Top             =   2880
            Width           =   1335
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Mode Parameters"
         Height          =   3375
         Index           =   1
         Left            =   -71760
         TabIndex        =   29
         Top             =   480
         Width           =   2655
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            ForeColor       =   &H000000C0&
            Height          =   1095
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   67
            TabStop         =   0   'False
            Text            =   "frmProperties.frx":15B6
            Top             =   2160
            Width           =   2415
         End
         Begin VB.TextBox txtLSBadjust 
            Height          =   285
            Left            =   1560
            TabIndex        =   50
            Top             =   1800
            Width           =   735
         End
         Begin VB.TextBox txtUSBadjust 
            Height          =   285
            Left            =   1560
            TabIndex        =   49
            Top             =   1440
            Width           =   735
         End
         Begin VB.ComboBox ComboDefaultStep 
            Height          =   315
            Index           =   1
            Left            =   1440
            TabIndex        =   31
            Text            =   "Combo2"
            Top             =   480
            Width           =   1095
         End
         Begin VB.ComboBox comboDefaultFilter 
            Height          =   315
            Index           =   1
            Left            =   1440
            TabIndex        =   30
            Text            =   "Combo3"
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label9 
            Caption         =   "LSB/CW Correction"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   1845
            Width           =   1455
         End
         Begin VB.Label Label8 
            Caption         =   "USB Correction"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   1480
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Hz"
            Height          =   255
            Left            =   2340
            TabIndex        =   52
            Top             =   1840
            Width           =   200
         End
         Begin VB.Label Label6 
            Caption         =   "Hz"
            Height          =   255
            Left            =   2340
            TabIndex        =   51
            Top             =   1480
            Width           =   200
         End
         Begin VB.Label Label2 
            Caption         =   "Default Freq Sep"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   33
            Top             =   510
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Default Filter"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   32
            Top             =   1000
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Quick Filter Selection"
         Height          =   3375
         Index           =   1
         Left            =   -74880
         TabIndex        =   28
         Top             =   480
         Width           =   2895
         Begin VB.CommandButton cmdSSBFilter 
            BackColor       =   &H00C0C0C0&
            Caption         =   "1800 Hz"
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
            Left            =   240
            MouseIcon       =   "frmProperties.frx":164C
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   73
            Tag             =   "1800"
            Top             =   2760
            Width           =   855
         End
         Begin VB.CommandButton cmdSSBFilter 
            BackColor       =   &H00C0C0C0&
            Caption         =   "2100 Hz"
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
            Left            =   240
            MouseIcon       =   "frmProperties.frx":1956
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   72
            Tag             =   "2100"
            Top             =   2280
            Width           =   855
         End
         Begin VB.CommandButton cmdSSBFilter 
            BackColor       =   &H00C0C0C0&
            Caption         =   "2400 Hz"
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
            Left            =   240
            MouseIcon       =   "frmProperties.frx":1C60
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   71
            Tag             =   "2400"
            Top             =   1800
            Width           =   855
         End
         Begin VB.CommandButton cmdSSBFilter 
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
            Index           =   2
            Left            =   240
            MouseIcon       =   "frmProperties.frx":1F6A
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   70
            Tag             =   "2700"
            Top             =   1320
            Width           =   855
         End
         Begin VB.CommandButton cmdSSBFilter 
            BackColor       =   &H00C0C0C0&
            Caption         =   "3000 Hz"
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
            Left            =   240
            MouseIcon       =   "frmProperties.frx":2274
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   69
            Tag             =   "3000"
            Top             =   840
            Width           =   855
         End
         Begin VB.CommandButton cmdSSBFilter 
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
            Index           =   0
            Left            =   240
            MouseIcon       =   "frmProperties.frx":257E
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   68
            Tag             =   "3300"
            Top             =   360
            Width           =   855
         End
         Begin VB.ComboBox ComboSSBfilter 
            Height          =   315
            Index           =   0
            Left            =   1200
            TabIndex        =   60
            Text            =   "Combo1"
            Top             =   480
            Width           =   1335
         End
         Begin VB.ComboBox ComboSSBfilter 
            Height          =   315
            Index           =   1
            Left            =   1200
            TabIndex        =   59
            Text            =   "Combo1"
            Top             =   960
            Width           =   1335
         End
         Begin VB.ComboBox ComboSSBfilter 
            Height          =   315
            Index           =   2
            Left            =   1200
            TabIndex        =   58
            Text            =   "Combo1"
            Top             =   1440
            Width           =   1335
         End
         Begin VB.ComboBox ComboSSBfilter 
            Height          =   315
            Index           =   3
            Left            =   1200
            TabIndex        =   57
            Text            =   "Combo1"
            Top             =   1920
            Width           =   1335
         End
         Begin VB.ComboBox ComboSSBfilter 
            Height          =   315
            Index           =   4
            Left            =   1200
            TabIndex        =   56
            Text            =   "Combo1"
            Top             =   2400
            Width           =   1335
         End
         Begin VB.ComboBox ComboSSBfilter 
            Height          =   315
            Index           =   5
            Left            =   1200
            TabIndex        =   55
            Text            =   "Combo1"
            Top             =   2880
            Width           =   1335
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Mode Parameters"
         Height          =   3375
         Index           =   0
         Left            =   -71760
         TabIndex        =   23
         Top             =   480
         Width           =   2655
         Begin VB.ComboBox comboDefaultFilter 
            Height          =   315
            Index           =   0
            Left            =   1440
            TabIndex        =   27
            Text            =   "Combo3"
            Top             =   960
            Width           =   1095
         End
         Begin VB.ComboBox ComboDefaultStep 
            Height          =   315
            Index           =   0
            Left            =   1440
            TabIndex        =   25
            Text            =   "Combo2"
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Default Filter"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   26
            Top             =   1000
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Default Freq Sep"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   24
            Top             =   510
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Quick Filter Selection"
         Height          =   3375
         Index           =   0
         Left            =   -74880
         TabIndex        =   16
         Top             =   480
         Width           =   2895
         Begin VB.CommandButton cmdAMFilter 
            BackColor       =   &H00C0C0C0&
            Caption         =   "3900 Hz"
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
            Left            =   240
            MouseIcon       =   "frmProperties.frx":2888
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   45
            Tag             =   "3900"
            Top             =   2760
            Width           =   855
         End
         Begin VB.CommandButton cmdAMFilter 
            BackColor       =   &H00C0C0C0&
            Caption         =   "4200 Hz"
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
            Left            =   240
            MouseIcon       =   "frmProperties.frx":2B92
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   44
            Tag             =   "4200"
            Top             =   2280
            Width           =   855
         End
         Begin VB.CommandButton cmdAMFilter 
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
            Left            =   240
            MouseIcon       =   "frmProperties.frx":2E9C
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   43
            Tag             =   "4800"
            Top             =   1800
            Width           =   855
         End
         Begin VB.CommandButton cmdAMFilter 
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
            Left            =   240
            MouseIcon       =   "frmProperties.frx":31A6
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   42
            Tag             =   "5400"
            Top             =   1320
            Width           =   855
         End
         Begin VB.CommandButton cmdAMFilter 
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
            Left            =   240
            MouseIcon       =   "frmProperties.frx":34B0
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   41
            Tag             =   "5700"
            Top             =   840
            Width           =   855
         End
         Begin VB.CommandButton cmdAMFilter 
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
            Left            =   240
            MouseIcon       =   "frmProperties.frx":37BA
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   40
            Tag             =   "6000"
            Top             =   360
            Width           =   855
         End
         Begin VB.ComboBox ComboAMfilter 
            Height          =   315
            Index           =   5
            Left            =   1200
            TabIndex        =   22
            Text            =   "Combo1"
            Top             =   2880
            Width           =   1335
         End
         Begin VB.ComboBox ComboAMfilter 
            Height          =   315
            Index           =   4
            Left            =   1200
            TabIndex        =   21
            Text            =   "Combo1"
            Top             =   2400
            Width           =   1335
         End
         Begin VB.ComboBox ComboAMfilter 
            Height          =   315
            Index           =   3
            Left            =   1200
            TabIndex        =   20
            Text            =   "Combo1"
            Top             =   1920
            Width           =   1335
         End
         Begin VB.ComboBox ComboAMfilter 
            Height          =   315
            Index           =   2
            Left            =   1200
            TabIndex        =   19
            Text            =   "Combo1"
            Top             =   1440
            Width           =   1335
         End
         Begin VB.ComboBox ComboAMfilter 
            Height          =   315
            Index           =   1
            Left            =   1200
            TabIndex        =   18
            Text            =   "Combo1"
            Top             =   960
            Width           =   1335
         End
         Begin VB.ComboBox ComboAMfilter 
            Height          =   315
            Index           =   0
            Left            =   1200
            TabIndex        =   17
            Text            =   "Combo1"
            Top             =   480
            Width           =   1335
         End
      End
      Begin VB.CheckBox chkFixedLevel 
         Caption         =   "Fixed Line Level"
         Height          =   255
         Left            =   3240
         TabIndex        =   11
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         Caption         =   "RX320 COM Port"
         ForeColor       =   &H000000C0&
         Height          =   2415
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   2775
         Begin VB.ListBox listComPorts 
            Height          =   1635
            ItemData        =   "frmProperties.frx":3AC4
            Left            =   170
            List            =   "frmProperties.frx":3ACB
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   8
            Top             =   600
            Width           =   2415
         End
         Begin VB.Label lblCOM 
            Caption         =   "Radio COM Port is"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   255
            Width           =   1400
         End
         Begin VB.Label lblRadioPort 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Left            =   1680
            TabIndex        =   9
            Top             =   250
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Software Information"
         ForeColor       =   &H000000C0&
         Height          =   855
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   5655
         Begin VB.Label lblDSPrev 
            Caption         =   "lblDSPrev"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   4695
         End
         Begin VB.Label lblProgVer 
            Caption         =   "lblProgVer"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   525
            Width           =   4815
         End
      End
      Begin VB.CheckBox chkMuteOnExit 
         Caption         =   "Mute Radio on Exit "
         Height          =   255
         Left            =   3240
         TabIndex        =   3
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label lblTZinfo 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label4"
         Height          =   255
         Left            =   3120
         TabIndex        =   15
         Top             =   3360
         Width           =   2775
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Time Zone Information"
         Height          =   255
         Left            =   3120
         TabIndex        =   14
         Top             =   3120
         Width           =   2775
      End
      Begin VB.Label lblUTCoffset 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         Height          =   255
         Left            =   4920
         TabIndex        =   13
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "UTC Offset in Minutes"
         Height          =   255
         Left            =   3120
         TabIndex        =   12
         Top             =   2760
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3600
      MouseIcon       =   "frmProperties.frx":3ADD
      MousePointer    =   99  'Custom
      Picture         =   "frmProperties.frx":3DE7
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0C0C0&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1200
      MouseIcon       =   "frmProperties.frx":40F1
      MousePointer    =   99  'Custom
      Picture         =   "frmProperties.frx":43FB
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4320
      Width           =   1575
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    '**************************************************
    '
    '   Propeties form for RX320 control program
    '
    '   Mark Mokoski
    '   20-APR-2005
    '
    '***************************************************
    Option Explicit

Private Sub chkFixedLevel_Click()

        If chkFixedLevel.Value = Checked Then
            listFixedLevel.Enabled = True
            listFixedLevel.Text = valFixedLevel
        Else
            listFixedLevel.Enabled = False
            listFixedLevel.Text = valFixedLevel
        End If

End Sub


Private Sub cmdCancel_Click()

    Unload Me

End Sub


Private Sub cmdOK_Click()

    Dim newCOM            As Integer

    'General Settings
    newCOM = Int(Mid(lblRadioPort.Caption, 5, 1))

        If newCOM <> RadioComPort Then
            'close COM port to radio, open new com port if changed

                With frmRX320
                    'Close COM Port
                    .RadioCOM.PortOpen = False
                    'Set new COM Port
                    RadioComPort = newCOM
                    'Open new COM Port
                    .RadioCOM.CommPort = RadioComPort
                    .RadioCOM.RThreshold = 1
                    .RadioCOM.Settings = "1200,N,8,1"
                    .RadioCOM.PortOpen = True
                End With

        End If

    'Set Mute on Exit

        If chkMuteOnExit.Value = Checked Then
            muteOnExit = True
        Else
            muteOnExit = False
        End If

    'Set Show 1 Hz digit

        If chk1Hz.Value = Checked Then
            show1Hz = True
            frmRX320.ComboStep.AddItem "1 Hz", 0
        Else
            show1Hz = False
            frmRX320.ComboStep.RemoveItem (0)
            'If step is 1 Hz, adjust step and display,
            'else, just set display

                If valStep = 1 Then
                    frmRX320.ComboStep.Text = "10 Hz"
                End If

        End If

    'Save fixed line level settings

        If chkFixedLevel.Value = Checked Then
            FixLineLevel = True
            valFixedLevel = listFixedLevel.Text
            frmRX320.LineVol.Enabled = False
            
        Else
            FixLineLevel = False
            valFixedLevel = listFixedLevel.Text
            frmRX320.LineVol.Enabled = True
        End If

    'Mode Settings
    AMfilter = Val(comboDefaultFilter(0).Text)       'Last selected AMfilter
    AMstep = ComboDefaultStep(0).Text                'AM tuning step
    SSBfilter = Val(comboDefaultFilter(1).Text)      'Last selected SSB filter
    SSBstep = ComboDefaultStep(1).Text               'SSB tuning step
    LSBoffset = Val(txtLSBadjust.Text)               'LSB/CW tuning correction
    USBoffset = Val(txtUSBadjust.Text)               'USB tuning correction
    CWfilter = Val(comboDefaultFilter(2).Text)       'Last selected CW filter
    CWstep = ComboDefaultStep(2).Text                'CW tuning step
    valCWOffSet = Val(comboCWoffset.Text)            'CW offset in Hz

    Dim x                 As Integer

        For x = 0 To 5
            AMquickfilter(x) = cmdAMFilter(x).Tag       'AM Quick filter selections
            SSBquickfilter(x) = cmdSSBFilter(x).Tag     'SSB Quick filter selections
            CWquickfilter(x) = cmdCWFilter(x).Tag       'CW Quick filter selection
        Next x

    DoEvents

    frmRX320.SetDisplay

    'Save all settings to registry
    Call SaveRegSettings
    'Close the Properties form
    frmRX320.DisplayPicture.SetFocus
    Unload Me

End Sub


Private Sub ComboAMfilter_Click(Index As Integer)

    'Set AM Quick filter value
    cmdAMFilter(Index).Tag = ComboAMfilter(Index).Text
    cmdAMFilter(Index).Caption = cmdAMFilter(Index).Tag & " Hz"

End Sub


Private Sub ComboCWfilter_Click(Index As Integer)

    'Set CW Quick filter value
    cmdCWFilter(Index).Tag = ComboCWfilter(Index).Text
    cmdCWFilter(Index).Caption = cmdCWFilter(Index).Tag & " Hz"

End Sub

Private Sub ComboSSBfilter_Click(Index As Integer)

    'Set SSB Quick filter value
    cmdSSBFilter(Index).Tag = ComboSSBfilter(Index).Text
    cmdSSBFilter(Index).Caption = cmdSSBFilter(Index).Tag & " Hz"

End Sub

Private Sub Form_Load()

    Dim x                 As Integer
    Dim i                 As Integer
    
    'Find installed COM ports
    FindPorts listComPorts
    
    'Display the radio firmware version
    lblDSPrev.Caption = "RX320 DSP Firmware Version " & DSPver
    
    'Display Progran version
    lblProgVer.Caption = "RX320 Control Program  -  Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    'Display the currently selected radio COM port
    lblRadioPort.Caption = "COM " & RadioComPort
    
    'Make the current COM port checked in the list


        For x = 0 To (listComPorts.ListCount - 1)

                If listComPorts.List(x) <> lblRadioPort.Caption Then
                    listComPorts.Selected(x) = False
                Else
                    listComPorts.Selected(x) = True
                End If

        Next x

    'Check if true, Mute On Exit

        If muteOnExit = True Then
            chkMuteOnExit.Value = Checked
        Else
            chkMuteOnExit = Unchecked
        End If

    'Check if true, Show 1 Hz
    
        If show1Hz = True Then
            chk1Hz.Value = Checked
        Else
            chk1Hz.Value = Unchecked
        End If

    'Populate listbox with audio levels

        For x = 0 To 63
            listFixedLevel.AddItem (x), x
        Next x

    'Populate Time Zone info labels
    x = (24 * 60) - UTCoffset
    lblUTCoffset = x - (24 * 60)

    lblTZinfo = longTZname
    
    'Set fixed line level controls
    
        If FixLineLevel = True Then
            chkFixedLevel.Value = Checked
            listFixedLevel.Enabled = True
            listFixedLevel.Text = valFixedLevel

        Else
            listFixedLevel.Text = valFixedLevel
            listFixedLevel.Enabled = False
            chkFixedLevel.Value = Unchecked
        End If

    'Populate filter selction combo boxes, same as frmrx320.ComboFilter.List

        For x = 0 To 5

                For i = 0 To frmRX320.ComboFilter.ListCount - 1
                    ComboAMfilter(x).AddItem frmRX320.ComboFilter.List(i)
                    ComboSSBfilter(x).AddItem frmRX320.ComboFilter.List(i)
                    ComboCWfilter(x).AddItem frmRX320.ComboFilter.List(i)
                    DoEvents
                Next i

            'Populate Quick Filter controls

            cmdAMFilter(x).Tag = AMquickfilter(x)
            cmdAMFilter(x).Caption = cmdAMFilter(x).Tag & " Hz"
            ComboAMfilter(x) = cmdAMFilter(x).Tag
            cmdSSBFilter(x).Tag = SSBquickfilter(x)
            cmdSSBFilter(x).Caption = cmdSSBFilter(x).Tag & " Hz"
            ComboSSBfilter(x) = cmdSSBFilter(x).Tag
            cmdCWFilter(x).Tag = CWquickfilter(x)
            cmdCWFilter(x).Caption = cmdCWFilter(x).Tag & " Hz"
            ComboCWfilter(x) = cmdCWFilter(x).Tag
            DoEvents
        Next x

        For i = 0 To frmRX320.ComboFilter.ListCount - 1
            comboDefaultFilter(0).AddItem frmRX320.ComboFilter.List(i)
            comboDefaultFilter(1).AddItem frmRX320.ComboFilter.List(i)
            comboDefaultFilter(2).AddItem frmRX320.ComboFilter.List(i)
            DoEvents
        Next i

    'Populate Step combo boxes

        For x = 0 To 2
            ComboDefaultStep(x).AddItem "10 Hz"
            ComboDefaultStep(x).AddItem "100 Hz"
            ComboDefaultStep(x).AddItem "1 KHz"
            ComboDefaultStep(x).AddItem "2.5 KHz"
            ComboDefaultStep(x).AddItem "5 KHz"
            ComboDefaultStep(x).AddItem "9 KHz"
            ComboDefaultStep(x).AddItem "10 KHz"
            ComboDefaultStep(x).AddItem "100 KHz"
            ComboDefaultStep(x).AddItem "1 MHz"
            DoEvents
        Next x

    'Populate cw offset combo box
    comboCWoffset.AddItem "400"
    comboCWoffset.AddItem "450"
    comboCWoffset.AddItem "500"
    comboCWoffset.AddItem "550"
    comboCWoffset.AddItem "600"
    comboCWoffset.AddItem "650"
    comboCWoffset.AddItem "700"
    comboCWoffset.AddItem "750"
    comboCWoffset.AddItem "800"
    comboCWoffset.AddItem "850"
    comboCWoffset.AddItem "900"
    
    comboCWoffset.Text = valCWOffSet
        
    'Populate default tuning step and default filter combo boxes
    
    ComboDefaultStep(0).Text = AMstep
    ComboDefaultStep(1).Text = SSBstep
    ComboDefaultStep(2).Text = CWstep
    comboDefaultFilter(0) = AMfilter
    comboDefaultFilter(1) = SSBfilter
    comboDefaultFilter(2) = CWfilter
    txtUSBadjust = USBoffset
    txtLSBadjust = LSBoffset

    Me.Visible = True

    tabProperties.SetFocus

End Sub

Private Sub listComPorts_ItemCheck(Item As Integer)

    Dim i            As Integer

    'Uncheck (deselect) old items in list

        For i = 0 To (listComPorts.ListCount - 1)

                If i <> Item Then
                    listComPorts.Selected(i) = False
                End If

        Next i

    lblRadioPort.Caption = listComPorts.List(Item)

End Sub
