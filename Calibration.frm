VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{1752FF26-D6C9-4BC8-BFE9-7D0CA26DED89}#1.0#0"; "BDaqOcx.dll"
Begin VB.Form FRMCAL 
   Caption         =   "HV_TEST Calibration"
   ClientHeight    =   9585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10290
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9585
   ScaleWidth      =   10290
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab Tab0 
      Height          =   3615
      Left            =   -30
      TabIndex        =   14
      Top             =   75
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   6376
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Gain Setting of Signal"
      TabPicture(0)   =   "Calibration.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FRAM1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Check Calibration Results"
      TabPicture(1)   =   "Calibration.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "optc(9)"
      Tab(1).Control(1)=   "optc(8)"
      Tab(1).Control(2)=   "optc(7)"
      Tab(1).Control(3)=   "optc(6)"
      Tab(1).Control(4)=   "optc(5)"
      Tab(1).Control(5)=   "optc(4)"
      Tab(1).Control(6)=   "optc(3)"
      Tab(1).Control(7)=   "optc(2)"
      Tab(1).Control(8)=   "optc(1)"
      Tab(1).Control(9)=   "optc(0)"
      Tab(1).ControlCount=   10
      Begin VB.OptionButton optc 
         Caption         =   "VOLTAGE (0-400 kV)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   9
         Left            =   -74715
         TabIndex        =   41
         Top             =   2895
         Width           =   2175
      End
      Begin VB.OptionButton optc 
         Caption         =   "VOLTAGE (0-232 kV)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   8
         Left            =   -74700
         TabIndex        =   40
         Top             =   2460
         Width           =   2175
      End
      Begin VB.OptionButton optc 
         Caption         =   "VOLTAGE (0-132 kV)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   7
         Left            =   -74670
         TabIndex        =   37
         Top             =   2040
         Width           =   2175
      End
      Begin VB.OptionButton optc 
         Caption         =   "CURRENT  (0 - 10 mA)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Index           =   6
         Left            =   -68235
         TabIndex        =   33
         Top             =   555
         Width           =   2775
      End
      Begin VB.OptionButton optc 
         Caption         =   "PD ( 0 - 20 PC)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   5
         Left            =   -71265
         TabIndex        =   32
         Top             =   1725
         Width           =   2175
      End
      Begin VB.OptionButton optc 
         Caption         =   "PD ( 0 - 5 PC)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   4
         Left            =   -71265
         TabIndex        =   31
         Top             =   1155
         Width           =   2175
      End
      Begin VB.OptionButton optc 
         Caption         =   "PD ( 0 - 2 PC)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   3
         Left            =   -71250
         TabIndex        =   30
         Top             =   600
         Width           =   2175
      End
      Begin VB.OptionButton optc 
         Caption         =   "VOLTAGE (0-100 kV)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   -74670
         TabIndex        =   29
         Top             =   1590
         Width           =   2175
      End
      Begin VB.OptionButton optc 
         Caption         =   "VOLTAGE (0-50 kV)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   -74685
         TabIndex        =   28
         Top             =   1095
         Width           =   2175
      End
      Begin VB.OptionButton optc 
         Caption         =   "VOLTAGE (0-10 kV)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   -74670
         TabIndex        =   27
         Top             =   615
         Width           =   2175
      End
      Begin VB.Frame FRAM1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3120
         Left            =   135
         TabIndex        =   15
         Top             =   315
         Width           =   9330
         Begin VB.OptionButton opt 
            Caption         =   "400 [KV]"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   9
            Left            =   5835
            TabIndex        =   42
            Top             =   660
            Width           =   1095
         End
         Begin VB.OptionButton opt 
            Caption         =   "230 [KV]"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   8
            Left            =   4695
            TabIndex        =   39
            Top             =   660
            Width           =   1095
         End
         Begin VB.OptionButton opt 
            Caption         =   "132 [KV]"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   7
            Left            =   3540
            TabIndex        =   38
            Top             =   660
            Width           =   1095
         End
         Begin VB.OptionButton opt 
            Caption         =   "10 [KV]"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   285
            TabIndex        =   36
            Top             =   660
            Width           =   975
         End
         Begin VB.OptionButton opt 
            Caption         =   " 50 [KV]"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   35
            Top             =   660
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton opt 
            Caption         =   "100 [KV]"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   2
            Left            =   2430
            TabIndex        =   34
            Top             =   660
            Width           =   1095
         End
         Begin VB.ComboBox cmbcrt 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   360
            ItemData        =   "Calibration.frx":0038
            Left            =   210
            List            =   "Calibration.frx":003A
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   2010
            Width           =   1455
         End
         Begin VB.OptionButton opt 
            Caption         =   "10 [mA]"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   375
            Index           =   6
            Left            =   1935
            TabIndex        =   20
            Top             =   2040
            Width           =   1215
         End
         Begin VB.OptionButton opt 
            Caption         =   "20   [Pc]"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Index           =   5
            Left            =   7785
            TabIndex        =   19
            Top             =   2040
            Width           =   1095
         End
         Begin VB.OptionButton opt 
            Caption         =   "10   [Pc]"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   240
            Index           =   4
            Left            =   7785
            TabIndex        =   18
            Top             =   1740
            Width           =   1095
         End
         Begin VB.OptionButton opt 
            Caption         =   "2   [Pc]"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Index           =   3
            Left            =   7785
            TabIndex        =   17
            Top             =   1335
            Width           =   1095
         End
         Begin VB.ComboBox cmbvlt 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   360
            ItemData        =   "Calibration.frx":003C
            Left            =   240
            List            =   "Calibration.frx":003E
            TabIndex        =   16
            Text            =   "Combo1"
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label lblvol 
            AutoSize        =   -1  'True
            Caption         =   "Voltage Selection"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   0
            Left            =   1200
            TabIndex        =   26
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label lblvol 
            AutoSize        =   -1  'True
            Caption         =   "PD "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   240
            Index           =   2
            Left            =   8145
            TabIndex        =   25
            Top             =   855
            Width           =   390
         End
         Begin VB.Label lblvol 
            AutoSize        =   -1  'True
            Caption         =   "CURRENT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   240
            Index           =   1
            Left            =   1980
            TabIndex        =   24
            Top             =   1665
            Width           =   1125
         End
         Begin VB.Label lblvol 
            AutoSize        =   -1  'True
            Caption         =   "RMS Value"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   3
            Left            =   255
            TabIndex        =   23
            Top             =   990
            Width           =   1050
         End
         Begin VB.Label lblvol 
            AutoSize        =   -1  'True
            Caption         =   "Peak Value"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   240
            Index           =   4
            Left            =   225
            TabIndex        =   22
            Top             =   1740
            Width           =   1095
         End
      End
   End
   Begin ComctlLib.Slider SLD 
      Height          =   4455
      Left            =   9450
      TabIndex        =   13
      Top             =   4920
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   7858
      _Version        =   327682
      BorderStyle     =   1
      Orientation     =   1
      Max             =   10000
   End
   Begin MSAdodcLib.Adodc Data 
      Height          =   375
      Left            =   7320
      Top             =   9195
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "data Gain"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.ComboBox cmbcon 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      ItemData        =   "Calibration.frx":0040
      Left            =   1200
      List            =   "Calibration.frx":0042
      TabIndex        =   10
      Text            =   "cmbcon1"
      Top             =   4035
      Width           =   1095
   End
   Begin VB.CommandButton CMDSET 
      Caption         =   " Gain Margin"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   1
      Top             =   3900
      Width           =   1575
   End
   Begin VB.TextBox TXTGAIN 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   4905
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   " "
      Top             =   4575
      Width           =   1335
   End
   Begin VB.TextBox TXTVOL 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   420
      Left            =   4920
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   " "
      Top             =   3915
      Width           =   1335
   End
   Begin VB.Timer t0 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2505
      Top             =   3780
   End
   Begin VB.CommandButton cmdend 
      Caption         =   "End"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7860
      TabIndex        =   3
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdstart 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6615
      TabIndex        =   0
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Timer t1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2040
      Top             =   3795
   End
   Begin VB.ComboBox cmbvd 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   1200
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   4515
      Width           =   1095
   End
   Begin VB.PictureBox pic 
      BackColor       =   &H80000005&
      FillColor       =   &H00FFFFFF&
      Height          =   4095
      Left            =   120
      ScaleHeight     =   4035
      ScaleWidth      =   9195
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4995
      Width           =   9255
      Begin VB.Line lv9 
         X1              =   8520
         X2              =   8400
         Y1              =   120
         Y2              =   4560
      End
      Begin VB.Line lv8 
         X1              =   7560
         X2              =   7560
         Y1              =   120
         Y2              =   4320
      End
      Begin VB.Line lv7 
         X1              =   6840
         X2              =   6960
         Y1              =   0
         Y2              =   4200
      End
      Begin VB.Line lv6 
         X1              =   6120
         X2              =   6120
         Y1              =   120
         Y2              =   4320
      End
      Begin VB.Line lv4 
         X1              =   3960
         X2              =   3960
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line lv3 
         X1              =   2280
         X2              =   2400
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line lv2 
         X1              =   1200
         X2              =   1200
         Y1              =   0
         Y2              =   4440
      End
      Begin VB.Line lv1 
         X1              =   360
         X2              =   480
         Y1              =   0
         Y2              =   4560
      End
      Begin VB.Line lv5 
         BorderWidth     =   2
         X1              =   5040
         X2              =   5040
         Y1              =   0
         Y2              =   4440
      End
      Begin VB.Line lh2 
         X1              =   120
         X2              =   9720
         Y1              =   360
         Y2              =   480
      End
      Begin VB.Line lh7 
         X1              =   0
         X2              =   9480
         Y1              =   3960
         Y2              =   3840
      End
      Begin VB.Line lh6 
         X1              =   0
         X2              =   9480
         Y1              =   3600
         Y2              =   3480
      End
      Begin VB.Line lh5 
         X1              =   0
         X2              =   9480
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line lh4 
         X1              =   0
         X2              =   9600
         Y1              =   2400
         Y2              =   2280
      End
      Begin VB.Line lh3 
         X1              =   0
         X2              =   9480
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line1 
         X1              =   -120
         X2              =   9480
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line lh1 
         X1              =   0
         X2              =   9600
         Y1              =   1920
         Y2              =   1920
      End
   End
   Begin BDaqOcxLibCtl.BufferedAiCtrl VIBuffer 
      Left            =   9150
      OleObjectBlob   =   "Calibration.frx":0044
      Top             =   4545
   End
   Begin VB.Label lblcon 
      AutoSize        =   -1  'True
      Caption         =   "DISPLAY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   240
      TabIndex        =   12
      Top             =   4035
      Width           =   735
   End
   Begin VB.Label lbl0 
      Caption         =   " [kv]  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   6360
      TabIndex        =   11
      Top             =   3915
      Width           =   1725
   End
   Begin VB.Label lblgain 
      AutoSize        =   -1  'True
      Caption         =   " SELECTED GAIN = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   2865
      TabIndex        =   8
      Top             =   4620
      Width           =   1845
   End
   Begin VB.Label LBLVOLT 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "VOLTAGE = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   3480
      TabIndex        =   7
      Top             =   4035
      Width           =   1170
   End
   Begin VB.Label lblvd 
      AutoSize        =   -1  'True
      Caption         =   "VOL/DIV"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   240
      TabIndex        =   6
      Top             =   4515
      Width           =   690
   End
End
Attribute VB_Name = "FRMCAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim sDataBuffer() As Double
Option Explicit
Dim dwDeviceNum, rtv As Integer
Dim gaincrt(0 To 14) As Single
Dim gainvlt(0 To 18) As Single 'added m.r
Dim gainpd(0 To 2) As Single
Dim maxv, vltdiv, cond As Double
Dim ppeak(1 To 4) As Double
Dim npeak(1 To 4) As Double
Dim ppk, npk, pdppk, pdnpk As Double
Dim draw As Boolean
Dim err As ErrorCode
Private Sub drawcurve()
Dim i, j, C, zr As Long

Dim preych1, preych2, p1, crtgain, vltgain, pdgain, m, BaseGain, cntPD As Double
Dim PDAvrg(1 To 100) As Double

pic.Cls
j = 0
zr = 0
ppeak(1) = 0
npeak(1) = 0
ppeak(2) = 0
npeak(2) = 0
ppeak(3) = 0
npeak(3) = 0
ppeak(4) = 0
npeak(4) = 0
ppk = 0
npk = 0
pdnpk = 0
pdppk = 0

'////////////////////////////////////////////////////////////////////////////

If ((opt(3).Value = False And opt(4).Value = False And opt(5).Value = False) And Tab0.Tab = 0) Or ((optc(3).Value = False And optc(4).Value = False And optc(5).Value = False) And Tab0.Tab = 1) Then '  voltage & current
    pic.Scale (0, 12)-(800, -12)
    
    For C = 3 To 400
        If sDataBuffer(C - 3) < 0 And sDataBuffer(C + 3) > 0 Then
            zr = C
            Exit For
        End If
    Next C
    
    For i = 3 To UBound(sDataBuffer) - 3
    
        p1 = sDataBuffer(i)
        For C = 1 To 3
            p1 = p1 + sDataBuffer(i - C) + sDataBuffer(i + C)
        Next C
        p1 = p1 / 7
    '///////////////////////////////////////////////////////////
        If i < 400 Then
            If ppeak(1) < p1 Then ppeak(1) = p1
            If npeak(1) > p1 Then npeak(1) = p1
            
        ElseIf i >= 400 And i < 800 Then
            If ppeak(2) < p1 Then ppeak(2) = p1
            If npeak(2) > p1 Then npeak(2) = p1
            
        ElseIf i >= 800 And i < 1200 Then
            If ppeak(3) < p1 Then ppeak(3) = p1
            If npeak(3) > p1 Then npeak(3) = p1
            
        ElseIf i >= 1200 Then
            If ppeak(4) < p1 Then ppeak(4) = p1
            If npeak(4) > p1 Then npeak(4) = p1
        End If
       If i >= zr And i < zr + 800 Then ' draw curve
            If j = 0 Then
                preych1 = p1
            Else
                pic.Line (j - 1, preych1 * 3 / vltdiv)-(j, p1 * 3 / vltdiv), vbRed
                preych1 = p1
            End If
            j = j + 1
       End If
    Next i
    
    ppk = (ppeak(1) + ppeak(2) + ppeak(3) + ppeak(4)) / 4
    npk = (npeak(1) + npeak(2) + npeak(3) + npeak(4)) / 4
    
    If ppk < Abs(npk) Then
        maxv = Abs(npk)
    Else
       maxv = ppk
    End If
    
 Else  '   ****************          PD Calculation     *****************************
 ' ///////////////////////////////////////////////////////////////////////////////
 
      pic.Scale (0, 12)-(UBound(sDataBuffer), -12)
      
        
'/////////////////////////////////////////////////////////////
    maxv = 0
    For C = 1 To 100
        cntPD = 0
        PDAvrg(C) = 0
        For i = (96 * (C - 1)) To (96 * C) - 1  ' last calculated arrey of PD is (96* 50)-1  (1 cycle)
            PDAvrg(C) = PDAvrg(C) + sDataBuffer(i)
            cntPD = cntPD + 1
        Next i
        PDAvrg(C) = PDAvrg(C) / cntPD
        
         ' Find Real Value of PD after removing DC offset
         
        For i = (96 * (C - 1)) To (96 * C) - 1
            If maxv < sDataBuffer(i) Then  ' find MAX value for I/O card gain
                maxv = sDataBuffer(i)
            End If
            
            sDataBuffer(i) = sDataBuffer(i) - PDAvrg(C)
        Next i
        
    Next C
      
      
'////////////////////////    DRaw PD ////////////////////////////////
    j = 0
    For i = 1 To UBound(sDataBuffer) - 2
        p1 = sDataBuffer(i)
            
        If pdppk < p1 Then pdppk = p1
        If pdnpk > p1 Then pdnpk = p1
          
        If j = 0 Then
            preych1 = p1
        Else
            pic.Line (j - 1, preych1 * 3 / vltdiv)-(j, p1 * 3 / vltdiv), vbRed
            preych1 = p1
        End If
        j = j + 1
               
    Next i
      
 End If


'//////////////////////////////////////////////////////////////////

If Tab0.Tab = 0 And (opt(0).Value = True Or opt(1).Value = True Or opt(2).Value = True Or opt(7).Value = True Or opt(8).Value = True) Then  ' voltage gain setting
    
    If ppk < Abs(npk) Then ' find absolute peak
        ppk = Abs(npk)
    End If
    
    TXTVOL.Text = Format((ppk / cond) * ((((SLD.Max + SLD.Min)) - SLD.Value) / 1000), "##0.0##")
    lbl0.Caption = "[KV] & .. [ I = V / 9.94] => I = " + Format((Val(TXTVOL.Text) / 9.94), "##0.0#")
    
ElseIf Tab0.Tab = 0 And opt(6).Value = True Then ' current gain setting

     If ppk < Abs(npk) Then ' find absolute peak
        ppk = Abs(npk)
     End If
   
     TXTVOL.Text = Format((ppk / cond) * ((((SLD.Max + SLD.Min)) - SLD.Value) / 1000), "#0.0##")
     lbl0.Caption = "[mA]"
     
ElseIf Tab0.Tab = 0 And (opt(3).Value = True Or opt(4).Value = True Or opt(5).Value = True) Then '  pd

    If pdppk < Abs(pdnpk) Then
        pdppk = Abs(pdnpk)
    End If
    
    TXTVOL.Text = Format(pdppk * ((((SLD.Max + SLD.Min)) - SLD.Value) / 1000), "00.00")
    lbl0.Caption = "[Pc]"

'******************************************************************************

ElseIf Tab0.Tab = 1 And optc(6).Value = True Then  ' current : check result
    lbl0.Caption = "[mA]"
    
    If ppk < Abs(npk) Then ' find absolute peak
        ppk = Abs(npk)
    End If
        
   
   If ppk * gaincrt(0) <= 0.1 Then
        crtgain = gaincrt(0)
            
    ElseIf ppk * gaincrt(0) > 0.1 And ppk * gaincrt(1) <= 0.15 Then
    
        m = (gaincrt(1) - gaincrt(0)) / (0.15 / gaincrt(1) - 0.1 / gaincrt(0))
        crtgain = gaincrt(0) + (m * (ppk - 0.1 / gaincrt(0)))
        
    ElseIf ppk * gaincrt(1) > 0.15 And ppk * gaincrt(2) <= 0.2 Then
        m = (gaincrt(2) - gaincrt(1)) / (0.2 / gaincrt(2) - 0.15 / gaincrt(1))
        crtgain = gaincrt(1) + (m * (ppk - 0.15 / gaincrt(1)))
            
    ElseIf ppk * gaincrt(2) > 0.2 And ppk * gaincrt(3) <= 0.5 Then
        m = (gaincrt(3) - gaincrt(2)) / (0.5 / gaincrt(3) - 0.2 / gaincrt(2))
        crtgain = gaincrt(2) + (m * (ppk - 0.2 / gaincrt(2)))
            
    ElseIf ppk * gaincrt(3) > 0.5 And ppk * gaincrt(4) <= 0.7 Then
        m = (gaincrt(4) - gaincrt(3)) / (0.7 / gaincrt(4) - 0.5 / gaincrt(3))
        crtgain = gaincrt(3) + (m * (ppk - 0.5 / gaincrt(3)))
            
    ElseIf ppk * gaincrt(4) > 0.7 And ppk * gaincrt(5) <= 1 Then
        m = (gaincrt(5) - gaincrt(4)) / (1 / gaincrt(5) - 0.7 / gaincrt(4))
        crtgain = gaincrt(4) + (m * (ppk - 0.7 / gaincrt(4)))
              
    ElseIf ppk * gaincrt(5) > 1 And ppk * gaincrt(6) <= 1.2 Then
        m = (gaincrt(6) - gaincrt(5)) / (1.2 / gaincrt(6) - 1 / gaincrt(5))
        crtgain = gaincrt(5) + (m * (ppk - 1 / gaincrt(5)))
        
    ElseIf ppk * gaincrt(6) > 1.2 And ppk * gaincrt(7) <= 1.5 Then
        m = (gaincrt(7) - gaincrt(6)) / (1.5 / gaincrt(7) - 1.2 / gaincrt(6))
        crtgain = gaincrt(6) + (m * (ppk - 1.2 / gaincrt(6)))
                                  
    ElseIf ppk * gaincrt(7) > 1.5 And ppk * gaincrt(8) <= 2 Then
        m = (gaincrt(8) - gaincrt(7)) / (2 / gaincrt(8) - 1.5 / gaincrt(7))
        crtgain = gaincrt(7) + (m * (ppk - 1.5 / gaincrt(7)))
         
    ElseIf ppk * gaincrt(8) > 2 And ppk * gaincrt(9) <= 2.5 Then
         m = (gaincrt(9) - gaincrt(8)) / (2.5 / gaincrt(9) - 2 / gaincrt(8))
        crtgain = gaincrt(8) + (m * (ppk - 2 / gaincrt(8)))
           
    ElseIf ppk * gaincrt(9) > 2.5 And ppk * gaincrt(10) <= 3 Then
        m = (gaincrt(10) - gaincrt(9)) / (3 / gaincrt(10) - 2.5 / gaincrt(9))
        crtgain = gaincrt(9) + (m * (ppk - 2.5 / gaincrt(9)))
                                                                                                       
    ElseIf ppk * gaincrt(10) > 3 And ppk * gaincrt(11) <= 3.5 Then
        m = (gaincrt(11) - gaincrt(10)) / (3.5 / gaincrt(11) - 3 / gaincrt(10))
        crtgain = gaincrt(10) + (m * (ppk - 3 / gaincrt(10)))
                                                                                                                                                                               
    ElseIf ppk * gaincrt(11) > 3.5 And ppk * gaincrt(12) <= 4 Then
        m = (gaincrt(12) - gaincrt(11)) / (4 / gaincrt(12) - 3.5 / gaincrt(11))
        crtgain = gaincrt(11) + (m * (ppk - 3.5 / gaincrt(11)))
                                                                                                                                                                                
    ElseIf ppk * gaincrt(12) > 4 And ppk * gaincrt(13) <= 4.5 Then
        m = (gaincrt(13) - gaincrt(12)) / (4.5 / gaincrt(13) - 4 / gaincrt(12))
        crtgain = gaincrt(12) + (m * (ppk - 4 / gaincrt(12)))
                                                                                                                                                                                                    
    ElseIf ppk * gaincrt(13) > 4.5 And ppk * gaincrt(14) <= 5 Then
        m = (gaincrt(14) - gaincrt(13)) / (5 / gaincrt(14) - 4.5 / gaincrt(13))
        crtgain = gaincrt(13) + (m * (ppk - 4.5 / gaincrt(13)))
        
    ElseIf ppk * gaincrt(14) > 5 Then
        crtgain = gaincrt(14)
    End If
     
    TXTVOL.Text = Format((ppk / cond) * crtgain, "#0.00#")

'//////////               Voltage

ElseIf Tab0.Tab = 1 And optc(0).Value = True Then   ' voltage calibration  10 KV
    
    If ppk < Abs(npk) Then ' find absolute peak
        ppk = Abs(npk)
    End If
    
    If ppk * gainvlt(0) <= 2 * 1.41 Then
        vltgain = gainvlt(0)
        
    ElseIf ppk * gainvlt(0) > 2 * 1.41 And ppk * gainvlt(1) <= 6 * 1.41 Then
    
        m = (gainvlt(1) - gainvlt(0)) / (6 * 1.4142 / gainvlt(1) - 2 * 1.4142 / gainvlt(0))
        vltgain = gainvlt(0) + (m * (ppk - 2 * 1.4141 / gainvlt(0)))
                        
    ElseIf ppk * gainvlt(1) > 6 * 1.41 And ppk * gainvlt(2) < 10 * 1.41 Then
        m = (gainvlt(2) - gainvlt(1)) / (10 * 1.4142 / gainvlt(2) - 6 * 1.4142 / gainvlt(1))
        vltgain = gainvlt(1) + (m * (ppk - 6 * 1.4141 / gainvlt(1)))
        
    ElseIf ppk * gainvlt(2) >= 10 * 1.41 Then
        vltgain = gainvlt(2)
        
    End If
    
    TXTVOL.Text = Format((ppk / cond) * vltgain, "00.000")
    lbl0.Caption = "[KV] "
    
ElseIf Tab0.Tab = 1 And optc(1).Value = True Then   ' voltage calibration  50 KV
    
    If ppk < Abs(npk) Then  ' find absolute peak
        ppk = Abs(npk)
    End If
    
    If ppk * gainvlt(3) <= 10 * 1.41 Then
         vltgain = gainvlt(3)
         
    ElseIf ppk * gainvlt(3) > 10 * 1.41 And ppk * gainvlt(4) <= 25 * 1.41 Then
        m = (gainvlt(4) - gainvlt(3)) / (25 * 1.4142 / gainvlt(4) - 10 * 1.4142 / gainvlt(3))
        vltgain = gainvlt(3) + (m * (ppk - 10 * 1.4141 / gainvlt(3)))
                        
     ElseIf ppk * gainvlt(4) > 25 * 1.41 And ppk * gainvlt(5) <= 50 * 1.41 Then
        m = (gainvlt(5) - gainvlt(4)) / (50 * 1.4142 / gainvlt(5) - 25 * 1.4142 / gainvlt(4))
        vltgain = gainvlt(4) + (m * (ppk - 25 * 1.4141 / gainvlt(4)))
            
    ElseIf ppk * gainvlt(5) >= 50 * 1.41 Then
    
        vltgain = gainvlt(5)
        
    End If
    
    TXTVOL.Text = Format((ppk / cond) * vltgain, "##0.0##")
    lbl0.Caption = "[KV]"
     
ElseIf Tab0.Tab = 1 And optc(2).Value = True Then   ' voltage calibration   ' voltage calibration  100 KV

    
    If ppk < Abs(npk) Then ' find absolute peak
        ppk = Abs(npk)
    End If
    
    If ppk * gainvlt(6) <= 30 * 1.41 Then '  gain6 (30)  gain7(60)  gain8(100)
        vltgain = gainvlt(6)
        
     ElseIf (ppk * gainvlt(6) > 30 * 1.41) And (ppk * gainvlt(7) <= 60 * 1.41) Then
        m = (gainvlt(7) - gainvlt(6)) / (60 * 1.4142 / gainvlt(7) - 30 * 1.4142 / gainvlt(6))
        vltgain = gainvlt(6) + (m * (ppk - 30 * 1.4141 / gainvlt(6)))
                        
    ElseIf ppk * gainvlt(7) > 60 * 1.41 And ppk * gainvlt(8) <= 100 * 1.41 Then
        m = (gainvlt(8) - gainvlt(7)) / (100 * 1.4142 / gainvlt(8) - 60 * 1.4142 / gainvlt(7))
        vltgain = gainvlt(7) + (m * (ppk - 60 * 1.4141 / gainvlt(7)))
            
    ElseIf ppk * gainvlt(8) >= 100 * 1.41 Then
    
        vltgain = gainvlt(8)
        
    End If
    
    TXTVOL.Text = Format((ppk / cond) * vltgain, "#0.0##")
    lbl0.Caption = "[KV] "
    
ElseIf Tab0.Tab = 1 And optc(7).Value = True Then   ' voltage calibration   ' voltage calibration  230 KV

    
    If ppk < Abs(npk) Then ' find absolute peak
        ppk = Abs(npk)
    End If
    
    If ppk * gainvlt(6) <= 30 * 1.41 Then '  gain6 (30)  gain7(60)  gain8(100)
        vltgain = gainvlt(6)
        
     ElseIf (ppk * gainvlt(6) > 30 * 1.41) And (ppk * gainvlt(7) <= 60 * 1.41) Then
        m = (gainvlt(7) - gainvlt(6)) / (60 * 1.4142 / gainvlt(7) - 30 * 1.4142 / gainvlt(6))
        vltgain = gainvlt(6) + (m * (ppk - 30 * 1.4141 / gainvlt(6)))
                        
    ElseIf ppk * gainvlt(7) > 60 * 1.41 And ppk * gainvlt(8) <= 100 * 1.41 Then
        m = (gainvlt(8) - gainvlt(7)) / (100 * 1.4142 / gainvlt(8) - 60 * 1.4142 / gainvlt(7))
        vltgain = gainvlt(7) + (m * (ppk - 60 * 1.4141 / gainvlt(7)))
            
    ElseIf ppk * gainvlt(8) >= 100 * 1.41 Then
    
        vltgain = gainvlt(8)
        
    End If
    
    TXTVOL.Text = Format((ppk / cond) * vltgain, "#0.0##")
    lbl0.Caption = "[KV] "
  ElseIf Tab0.Tab = 1 And optc(8).Value = True Then   ' voltage calibration   ' voltage calibration  400 KV

    
    If ppk < Abs(npk) Then ' find absolute peak
        ppk = Abs(npk)
    End If
    
    If ppk * gainvlt(6) <= 30 * 1.41 Then '  gain6 (30)  gain7(60)  gain8(100)
        vltgain = gainvlt(6)
        
     ElseIf (ppk * gainvlt(6) > 30 * 1.41) And (ppk * gainvlt(7) <= 60 * 1.41) Then
        m = (gainvlt(7) - gainvlt(6)) / (60 * 1.4142 / gainvlt(7) - 30 * 1.4142 / gainvlt(6))
        vltgain = gainvlt(6) + (m * (ppk - 30 * 1.4141 / gainvlt(6)))
                        
    ElseIf ppk * gainvlt(7) > 60 * 1.41 And ppk * gainvlt(8) <= 100 * 1.41 Then
        m = (gainvlt(8) - gainvlt(7)) / (100 * 1.4142 / gainvlt(8) - 60 * 1.4142 / gainvlt(7))
        vltgain = gainvlt(7) + (m * (ppk - 60 * 1.4141 / gainvlt(7)))
            
    ElseIf ppk * gainvlt(8) >= 100 * 1.41 Then
    
        vltgain = gainvlt(8)
        
    End If
    
    TXTVOL.Text = Format((ppk / cond) * vltgain, "#0.0##")
    lbl0.Caption = "[KV] "
'****************** PD calibtration

ElseIf Tab0.Tab = 1 And (optc(3).Value = True Or optc(4).Value = True Or optc(5).Value = True) Then '  pd

    lbl0.Caption = "[Pc]"
    If optc(3).Value = True Then
     pdgain = gainpd(0)
     
    ElseIf optc(4).Value = True Then
     pdgain = gainpd(1)
     
    ElseIf optc(4).Value = True Then
     pdgain = gainpd(2)
     
    End If
    
     If pdppk < Abs(pdnpk) Then
         pdppk = Abs(pdnpk)
     End If
     TXTVOL.Text = Format(pdppk * pdgain, "#0.0#")
 
End If

End Sub


Private Sub formscaling()
Dim hgh, wid As Double

FRMCAL.Width = Screen.Width * 0.85
FRMCAL.Height = Screen.Height * 0.8
hgh = FRMCAL.Height
wid = FRMCAL.Width

With Tab0
    .Width = wid * 0.75
    .Height = hgh * 0.3
    .Top = hgh * 0.03
    .Left = wid * 0.03
End With

With cmdstart
    .Top = hgh * 0.1
    .Left = wid * 0.85
     
End With

With cmdend
    .Top = hgh * 0.85
    .Left = wid * 0.85
     
End With

With CMDSET
    .Top = hgh * 0.23
    .Left = wid * 0.83
     
End With

With SLD
    .Top = hgh * 0.3
    .Left = wid * 0.86
    .Height = hgh * 0.52
End With

With pic
    .Width = wid * 0.75
    .Height = hgh * 0.45
    .Top = hgh * 0.47
    .Left = wid * 0.03
End With
'//////////////////////////////////////////////////////
With lblcon
    .Top = hgh * 0.365
    .Left = wid * 0.04
End With

With lblvd
    .Top = hgh * 0.415
    .Left = wid * 0.04
End With

With cmbcon
    .Top = hgh * 0.36
    .Left = wid * 0.12
End With

With cmbvd
    .Top = hgh * 0.41
    .Left = wid * 0.12
End With


With LBLVOLT
    .Top = hgh * 0.365
    .Left = wid * 0.3
End With

With lblgain
    .Top = hgh * 0.415
    .Left = wid * 0.25
End With

With TXTVOL
    .Top = hgh * 0.35
    .Left = wid * 0.42
End With

With lbl0
    .Top = hgh * 0.354
    .Left = wid * 0.54
End With


With TXTGAIN
    .Top = hgh * 0.41
    .Left = wid * 0.42
End With




'///////////////////////////////////////////////
With fram1
    .Height = Tab0.Height * 0.8
    .Width = Tab0.Width * 0.93
    .Top = Tab0.Height * 0.12
    .Left = Tab0.Width * 0.03
End With

With lblvol(0)
    .Top = fram1.Height * 0.15
    .Left = fram1.Width * 0.1
End With

With lblvol(1)
    .Top = fram1.Height * 0.15
    .Left = fram1.Width * 0.55
End With
With lblvol(2)
    .Top = fram1.Height * 0.15
    .Left = fram1.Width * 0.85
End With

With opt(0)
    .Top = fram1.Height * 0.4
    .Left = fram1.Width * 0.03
End With
With opt(1)
    .Top = fram1.Height * 0.4
    .Left = fram1.Width * 0.16
End With
With opt(2)
    .Top = fram1.Height * 0.4
    .Left = fram1.Width * 0.3
End With

With opt(3)
    .Top = fram1.Height * 0.4
    .Left = fram1.Width * 0.82
End With
With opt(4)
    .Top = fram1.Height * 0.57
    .Left = fram1.Width * 0.82
End With
With opt(5)
    .Top = fram1.Height * 0.7
    .Left = fram1.Width * 0.82
End With


With opt(6)
    .Top = fram1.Height * 0.4
    .Left = fram1.Width * 0.55
End With

With cmbvlt
    .Top = fram1.Height * 0.7
    .Left = fram1.Width * 0.1
End With
With lblvol(3)
    .Top = fram1.Height * 0.7
    .Left = fram1.Width * 0.25
End With

With cmbcrt
    .Top = fram1.Height * 0.7
    .Left = fram1.Width * 0.55
End With
With lblvol(4)
    .Top = fram1.Height * 0.7
    .Left = fram1.Width * 0.63
End With
'/////////////////////////////////////////////////////
With optc(0)
    .Top = fram1.Height * 0.35
    .Left = fram1.Width * 0.07
End With

With optc(1)
    .Top = fram1.Height * 0.62
    .Left = fram1.Width * 0.07
End With

With optc(2)
    .Top = fram1.Height * 0.86
    .Left = fram1.Width * 0.07
End With


With optc(3)
    .Top = fram1.Height * 0.35
    .Left = fram1.Width * 0.75
End With

With optc(4)
    .Top = fram1.Height * 0.62
    .Left = fram1.Width * 0.75
End With

With optc(5)
    .Top = fram1.Height * 0.62
    .Left = fram1.Width * 0.75
End With
With optc(5)
    .Top = fram1.Height * 0.86
    .Left = fram1.Width * 0.75
End With


With optc(6)
    .Top = fram1.Height * 0.35
    .Left = fram1.Width * 0.45
End With



End Sub

Private Sub selectgain()
Dim i, j As Integer
  
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
' 2 channel for data acquisition in one stage   [(0,1)=current (1,2)=voltage (3,4)=pd]


VIBuffer.ScanChannel.ChannelCount = 1   ' number of channel 1
VIBuffer.ScanChannel.IntervalCount = 1
 
If ((opt(0).Value = True Or opt(1).Value = True Or opt(2).Value = True) And Tab0.Tab = 0) Or ((optc(0).Value = True Or optc(1).Value = True Or optc(2).Value = True) And Tab0.Tab = 1) Then
    VIBuffer.ScanChannel.ChannelStart = 2   ' Voltage Signal
    VIBuffer.ScanChannel.Samples = 2000 ' 5 cycles
    VIBuffer.ConvertClock.Rate = 20000
    
    If maxv <= 0.62 Then
      VIBuffer.Channel(2).ValueRange = 12  ' gain for voltage = +/- 0.625
            
    ElseIf maxv > 0.62 And maxv <= 1.2 Then
        VIBuffer.Channel(2).ValueRange = 4 ' gain for voltage = +/- 1.25 v
                
    ElseIf maxv > 1.2 And maxv <= 2.4 Then
          VIBuffer.Channel(2).ValueRange = 3 ' gain for voltage = +/- 2.5 v
                
    ElseIf maxv > 2.4 And maxv <= 4.9 Then
          VIBuffer.Channel(2).ValueRange = 2 ' gain for voltage = +/- 5
                
    ElseIf maxv > 4.9 Then
         VIBuffer.Channel(2).ValueRange = 1 ' gain for voltage = +/- 10 v
    End If

   
   
'                    P 4, 5
ElseIf ((opt(3).Value = True Or opt(4).Value = True Or opt(5).Value = True) And Tab0.Tab = 0) Or ((optc(3).Value = True Or optc(4).Value = True Or optc(5).Value = True) And Tab0.Tab = 1) Then
    VIBuffer.ScanChannel.ChannelStart = 6   ' PD Signal
    VIBuffer.ScanChannel.Samples = 4800 * 2 ' 2 cycles
    VIBuffer.ConvertClock.Rate = 240000
    If maxv <= 0.62 Then
      VIBuffer.Channel(6).ValueRange = 12  ' gain for voltage = +/- 0.625
            
    ElseIf maxv > 0.62 And maxv <= 1.2 Then
        VIBuffer.Channel(6).ValueRange = 4 ' gain for voltage = +/- 1.25 v
                
    ElseIf maxv > 1.2 And maxv <= 2.4 Then
          VIBuffer.Channel(6).ValueRange = 3 ' gain for voltage = +/- 2.5 v
                
    ElseIf maxv > 2.4 And maxv <= 4.9 Then
          VIBuffer.Channel(6).ValueRange = 2 ' gain for voltage = +/- 5
                
    ElseIf maxv > 4.9 Then
         VIBuffer.Channel(6).ValueRange = 1 ' gain for voltage = +/- 10 v
    End If
 
    
    
    '                     Current Ch 0.1
    
ElseIf (opt(6).Value = True And Tab0.Tab = 0) Or (optc(6).Value = True And Tab0.Tab = 1) Then

    VIBuffer.ScanChannel.ChannelStart = 0   ' Current Signal
    VIBuffer.ScanChannel.Samples = 2000 ' 5 cycles
    VIBuffer.ConvertClock.Rate = 20000
    
    If maxv <= 0.62 Then
      VIBuffer.Channel(0).ValueRange = 12  ' gain for voltage = +/- 0.625
            
    ElseIf maxv > 0.62 And maxv <= 1.2 Then
        VIBuffer.Channel(0).ValueRange = 4 ' gain for voltage = +/- 1.25 v
                
    ElseIf maxv > 1.2 And maxv <= 2.4 Then
          VIBuffer.Channel(0).ValueRange = 3 ' gain for voltage = +/- 2.5 v
                
    ElseIf maxv > 2.4 And maxv <= 4.9 Then
          VIBuffer.Channel(0).ValueRange = 2 ' gain for voltage = +/- 5
                
    ElseIf maxv > 4.9 Then
         VIBuffer.Channel(0).ValueRange = 1 ' gain for voltage = +/- 10 v
    End If
   
End If



End Sub



Private Sub cmbcon_Change()
Select Case cmbcon.ListIndex
Case 0
    cond = 2 ^ 0.5
Case 1
    cond = 1
End Select
End Sub

Private Sub cmbcon_Click()
Select Case cmbcon.ListIndex
Case 0
    cond = 2 ^ 0.5
Case 1
    cond = 1
End Select
End Sub


Private Sub cmbcrt_Change()
FRMCAL.MousePointer = 13
SLD.Min = 0
SLD.Max = 100000
 Select Case cmbcrt.ListIndex
    Case 0
        SLD.Min = (gaincrt(cmbcrt.ListIndex) - 0.5) * 1000
        SLD.Max = (gaincrt(cmbcrt.ListIndex) + 0.5) * 1000
    Case 1
        SLD.Min = (gaincrt(cmbcrt.ListIndex) - 0.5) * 1000
        SLD.Max = (gaincrt(cmbcrt.ListIndex) + 0.5) * 1000
    Case 2
        SLD.Min = (gaincrt(cmbcrt.ListIndex) - 0.5) * 1000
        SLD.Max = (gaincrt(cmbcrt.ListIndex) + 0.5) * 1000
    Case 3
        SLD.Min = (gaincrt(cmbcrt.ListIndex) - 0.5) * 1000
        SLD.Max = (gaincrt(cmbcrt.ListIndex) + 0.5) * 1000
    Case 4
        SLD.Min = (gaincrt(cmbcrt.ListIndex) - 0.5) * 1000
        SLD.Max = (gaincrt(cmbcrt.ListIndex) + 0.5) * 1000
    Case 5
        SLD.Min = (gaincrt(cmbcrt.ListIndex) - 0.5) * 1000
        SLD.Max = (gaincrt(cmbcrt.ListIndex) + 0.5) * 1000
    Case 6
        SLD.Min = (gaincrt(cmbcrt.ListIndex) - 0.5) * 1000
        SLD.Max = (gaincrt(cmbcrt.ListIndex) + 0.5) * 1000
    Case 7
        SLD.Min = (gaincrt(cmbcrt.ListIndex) - 0.5) * 1000
        SLD.Max = (gaincrt(cmbcrt.ListIndex) + 0.5) * 1000
    Case 8
        SLD.Min = (gaincrt(cmbcrt.ListIndex) - 0.5) * 1000
        SLD.Max = (gaincrt(cmbcrt.ListIndex) + 0.5) * 1000
    Case 9
        SLD.Min = (gaincrt(cmbcrt.ListIndex) - 0.5) * 1000
        SLD.Max = (gaincrt(cmbcrt.ListIndex) + 0.5) * 1000
     Case 10
        SLD.Min = (gaincrt(cmbcrt.ListIndex) - 0.1) * 1000
        SLD.Max = (gaincrt(cmbcrt.ListIndex) + 0.1) * 1000
     Case 11
        SLD.Min = (gaincrt(cmbcrt.ListIndex) - 0.1) * 1000
        SLD.Max = (gaincrt(cmbcrt.ListIndex) + 0.1) * 1000
     Case 12
        SLD.Min = (gaincrt(cmbcrt.ListIndex) - 0.1) * 1000
        SLD.Max = (gaincrt(cmbcrt.ListIndex) + 0.1) * 1000
     Case 13
        SLD.Min = (gaincrt(cmbcrt.ListIndex) - 0.1) * 1000
        SLD.Max = (gaincrt(cmbcrt.ListIndex) + 0.1) * 1000
     Case 14
        SLD.Min = (gaincrt(cmbcrt.ListIndex) - 0.1) * 1000
        SLD.Max = (gaincrt(cmbcrt.ListIndex) + 0.1) * 1000
    End Select
SLD.Value = (SLD.Min + SLD.Max) / 2
TXTGAIN.Text = Str((((SLD.Max + SLD.Min)) - SLD.Value) / 1000)
FRMCAL.MousePointer = 0

End Sub

Private Sub cmbcrt_Click()
FRMCAL.MousePointer = 13
SLD.Min = 0
SLD.Max = 100000
Select Case cmbcrt.ListIndex
    Case 0
        SLD.Min = (gaincrt(cmbcrt.ListIndex) - 0.5) * 1000
        SLD.Max = (gaincrt(cmbcrt.ListIndex) + 0.5) * 1000
    Case 1
        SLD.Min = (gaincrt(cmbcrt.ListIndex) - 0.5) * 1000
        SLD.Max = (gaincrt(cmbcrt.ListIndex) + 0.5) * 1000
    Case 2
        SLD.Min = (gaincrt(cmbcrt.ListIndex) - 0.5) * 1000
        SLD.Max = (gaincrt(cmbcrt.ListIndex) + 0.5) * 1000
    Case 3
        SLD.Min = (gaincrt(cmbcrt.ListIndex) - 0.5) * 1000
        SLD.Max = (gaincrt(cmbcrt.ListIndex) + 0.5) * 1000
    Case 4
        SLD.Min = (gaincrt(cmbcrt.ListIndex) - 0.5) * 1000
        SLD.Max = (gaincrt(cmbcrt.ListIndex) + 0.5) * 1000
    Case 5
        SLD.Min = (gaincrt(cmbcrt.ListIndex) - 0.5) * 1000
        SLD.Max = (gaincrt(cmbcrt.ListIndex) + 0.5) * 1000
    Case 6
        SLD.Min = (gaincrt(cmbcrt.ListIndex) - 0.5) * 1000
        SLD.Max = (gaincrt(cmbcrt.ListIndex) + 0.5) * 1000
    Case 7
        SLD.Min = (gaincrt(cmbcrt.ListIndex) - 0.5) * 1000
        SLD.Max = (gaincrt(cmbcrt.ListIndex) + 0.5) * 1000
    Case 8
        SLD.Min = (gaincrt(cmbcrt.ListIndex) - 0.5) * 1000
        SLD.Max = (gaincrt(cmbcrt.ListIndex) + 0.5) * 1000
    Case 9
        SLD.Min = (gaincrt(cmbcrt.ListIndex) - 0.5) * 1000
        SLD.Max = (gaincrt(cmbcrt.ListIndex) + 0.5) * 1000
     Case 10
        SLD.Min = (gaincrt(cmbcrt.ListIndex) - 0.1) * 1000
        SLD.Max = (gaincrt(cmbcrt.ListIndex) + 0.1) * 1000
     Case 11
        SLD.Min = (gaincrt(cmbcrt.ListIndex) - 0.1) * 1000
        SLD.Max = (gaincrt(cmbcrt.ListIndex) + 0.1) * 1000
     Case 12
        SLD.Min = (gaincrt(cmbcrt.ListIndex) - 0.1) * 1000
        SLD.Max = (gaincrt(cmbcrt.ListIndex) + 0.1) * 1000
    Case 13
        SLD.Min = (gaincrt(cmbcrt.ListIndex) - 0.1) * 1000
        SLD.Max = (gaincrt(cmbcrt.ListIndex) + 0.1) * 1000
     Case 14
        SLD.Min = (gaincrt(cmbcrt.ListIndex) - 0.1) * 1000
        SLD.Max = (gaincrt(cmbcrt.ListIndex) + 0.1) * 1000
    End Select
    SLD.Value = (SLD.Min + SLD.Max) / 2
TXTGAIN.Text = Str((((SLD.Max + SLD.Min)) - SLD.Value) / 1000)
FRMCAL.MousePointer = 0
End Sub


Private Sub cmbvd_Change()
Select Case cmbvd.ListIndex
Case 0
    vltdiv = 0.005
Case 1
    vltdiv = 0.05
Case 2
    vltdiv = 0.1
Case 3
    vltdiv = 0.2
Case 4
    vltdiv = 0.5
Case 5
    vltdiv = 1
Case 6
    vltdiv = 2
Case 7
    vltdiv = 5
End Select
End Sub

Private Sub cmbvd_Click()
Select Case cmbvd.ListIndex
Case 0
    vltdiv = 0.005
Case 1
    vltdiv = 0.05
Case 2
    vltdiv = 0.1
Case 3
    vltdiv = 0.2
Case 4
    vltdiv = 0.5
Case 5
    vltdiv = 1
Case 6
    vltdiv = 2
Case 7
    vltdiv = 5
End Select
End Sub


Private Sub cmbvlt_Change()
FRMCAL.MousePointer = 13
If opt(0).Value = True Then
    Select Case cmbvlt.ListIndex
    Case 0
        SLD.Min = (gainvlt(0) - 1) * 1000
        SLD.Max = (gainvlt(0) + 1) * 1000
    Case 1
        SLD.Min = (gainvlt(1) - 1) * 1000
        SLD.Max = (gainvlt(1) + 1) * 1000
    
    Case 2
        SLD.Min = (gainvlt(2) - 1) * 1000
        SLD.Max = (gainvlt(2) + 1) * 1000
    End Select
ElseIf opt(1).Value = True Then
    Select Case cmbvlt.ListIndex
    Case 0
        SLD.Min = (gainvlt(3) - 1) * 1000
        SLD.Max = (gainvlt(3) + 1) * 1000
        
    Case 1
        SLD.Min = (gainvlt(4) - 1) * 1000
        SLD.Max = (gainvlt(4) + 1) * 1000
    
    Case 2
        SLD.Min = (gainvlt(5) - 1) * 1000
        SLD.Max = (gainvlt(5) + 1) * 1000
    End Select
ElseIf opt(2).Value = True Then
    Select Case cmbvlt.ListIndex
    Case 0
        SLD.Min = (gainvlt(6) - 1) * 1000
        SLD.Max = (gainvlt(6) + 1) * 1000
        
    Case 1
        SLD.Min = (gainvlt(7) - 1) * 1000
        SLD.Max = (gainvlt(7) + 1) * 1000
    
    Case 2
        SLD.Min = (gainvlt(8) - 1) * 1000
        SLD.Max = (gainvlt(8) + 1) * 1000
    End Select
End If
SLD.Value = (SLD.Min + SLD.Max) / 2
TXTGAIN.Text = Str((((SLD.Max + SLD.Min)) - SLD.Value) / 1000)
FRMCAL.MousePointer = 0
End Sub

Private Sub cmbvlt_Click()
FRMCAL.MousePointer = 13
SLD.Min = 0
SLD.Max = 300000
If opt(0).Value = True Then
    Select Case cmbvlt.ListIndex
    Case 0
        SLD.Min = (gainvlt(0) - 1) * 1000
        SLD.Max = (gainvlt(0) + 1) * 1000
    Case 1
        SLD.Min = (gainvlt(1) - 1) * 1000
        SLD.Max = (gainvlt(1) + 1) * 1000
    
    Case 2
        SLD.Min = (gainvlt(2) - 1) * 1000
        SLD.Max = (gainvlt(2) + 1) * 1000
    End Select
ElseIf opt(1).Value = True Then
    Select Case cmbvlt.ListIndex
    Case 0
        SLD.Min = (gainvlt(3) - 1) * 1000
        SLD.Max = (gainvlt(3) + 1) * 1000
        
    Case 1
        SLD.Min = (gainvlt(4) - 1) * 1000
        SLD.Max = (gainvlt(4) + 1) * 1000
    
    Case 2
        SLD.Min = (gainvlt(5) - 1) * 1000
        SLD.Max = (gainvlt(5) + 1) * 1000
    End Select
ElseIf opt(2).Value = True Then
    Select Case cmbvlt.ListIndex
    Case 0
        SLD.Min = (gainvlt(6) - 1) * 1000
        SLD.Max = (gainvlt(6) + 1) * 1000
        
    Case 1
        SLD.Min = (gainvlt(7) - 1) * 1000
        SLD.Max = (gainvlt(7) + 1) * 1000
    
    Case 2
        SLD.Min = (gainvlt(8) - 1) * 1000
        SLD.Max = (gainvlt(8) + 1) * 1000
    End Select
End If
SLD.Value = (SLD.Min + SLD.Max) / 2
TXTGAIN.Text = Str((((SLD.Max + SLD.Min)) - SLD.Value) / 1000)
FRMCAL.MousePointer = 0

End Sub


Private Sub cmdend_Click()
If t0.Enabled = False Then
    Unload Me
    frmtest.Show
End If
End Sub

Private Sub CMDSET_Click()
't1.Enabled = False
't0.Enabled = False
frmset.Show vbModal
End Sub

Private Sub cmdstart_Click()
'Dim ptEnableEvent As PT_EnableEvent
Dim i As Integer, j As Integer
Dim DataSize As Long
pic.Cls
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
If cmdstart.Caption = "START" Then
       
        fram1.Enabled = False
        Tab0.Enabled = False
        cmdstart.Caption = "STOP"
        t0.Enabled = True
        t0.Interval = 10
        draw = False
        maxv = 4.5
        
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
ElseIf cmdstart.Caption = "STOP" Then
    fram1.Enabled = True
    Tab0.Enabled = True
    cmdstart.Caption = "START"
    t0.Enabled = False
    If Tab0.Tab = 0 Then
        rtv = MsgBox("DO YOU WANT SAVE NEW GAIN FOR THIS SIGNAL ? ", vbYesNo + vbExclamation, "ATTENTION")
        If rtv = 6 Then
            If opt(0) Then ' Voltage 0-10 kV
                Select Case cmbvlt.ListIndex
                Case 0 ' 2 kv
                    gainvlt(0) = (((SLD.Max + SLD.Min)) - SLD.Value) / 1000
                     With Data.Recordset
                         .MoveFirst
                        
                         !Voltage_gain_10_2 = Round(gainvlt(0), 3)
                         .Update
                     End With
                 Case 1 ' 6 kv
                    gainvlt(1) = (((SLD.Max + SLD.Min)) - SLD.Value) / 1000
                     With Data.Recordset
                         .MoveFirst
                         
                         !Voltage_gain_10_6 = Round(gainvlt(1), 3)
                         .Update
                     End With
                    Case 2 ' 10 kv
                        gainvlt(2) = (((SLD.Max + SLD.Min)) - SLD.Value) / 1000
                         With Data.Recordset
                             .MoveFirst
                             
                             !Voltage_gain_10_10 = Round(gainvlt(2), 3)
                             .Update
                         End With
                    End Select
                
            ElseIf opt(1) Then ' Voltage 0-50 kV
                Select Case cmbvlt.ListIndex
                Case 0 ' 10 kV
                    gainvlt(3) = (((SLD.Max + SLD.Min)) - SLD.Value) / 1000
                     With Data.Recordset
                         .MoveFirst
                        
                         !Voltage_gain_50_10 = Round(gainvlt(3), 3)
                         .Update
                     End With
                 Case 1 ' 25 kv
                    gainvlt(4) = (((SLD.Max + SLD.Min)) - SLD.Value) / 1000
                     With Data.Recordset
                         .MoveFirst
                         
                         !Voltage_gain_50_25 = Round(gainvlt(4), 3)
                         .Update
                     End With
                    Case 2   ' 50 kV
                        gainvlt(5) = (((SLD.Max + SLD.Min)) - SLD.Value) / 1000
                         With Data.Recordset
                             .MoveFirst
                            
                             !Voltage_gain_50_50 = Round(gainvlt(5), 3)
                             .Update
                         End With
                    End Select
                
                
            ElseIf opt(2) Then ' Voltage 0-100 kV
                Select Case cmbvlt.ListIndex
                Case 0 '  30 kV
                    gainvlt(6) = (((SLD.Max + SLD.Min)) - SLD.Value) / 1000
                     With Data.Recordset
                         .MoveFirst
                         
                         !Voltage_gain_100_30 = Round(gainvlt(6), 3)
                         .Update
                     End With
                 Case 1  ' 60 kv
                    gainvlt(7) = (((SLD.Max + SLD.Min)) - SLD.Value) / 1000
                     With Data.Recordset
                         .MoveFirst
                         
                         !Voltage_gain_100_60 = Round(gainvlt(7), 3)
                         .Update
                     End With
                    Case 2  ' 100 kv
                        gainvlt(8) = (((SLD.Max + SLD.Min)) - SLD.Value) / 1000
                         With Data.Recordset
                             .MoveFirst
                            
                             !Voltage_gain_100_100 = Round(gainvlt(8), 3)
                             .Update
                         End With
                    End Select
                
                
            ElseIf opt(3) Then ' PD Gain 2 pC
                gainpd(0) = (((SLD.Max + SLD.Min)) - SLD.Value) / 1000
                With Data.Recordset
                    .MoveFirst
                   
                    !PD_gain_2 = Round(gainpd(0), 3)
                    .Update
                End With
                
            ElseIf opt(4) Then ' PD Gain 10 pC
                gainpd(1) = (((SLD.Max + SLD.Min)) - SLD.Value) / 1000
                With Data.Recordset
                    .MoveFirst
                    
                    !PD_gain_10 = Round(gainpd(1), 3)
                    .Update
                End With
                
            ElseIf opt(5) Then ' PD Gain 20 pC
                gainpd(2) = (((SLD.Max + SLD.Min)) - SLD.Value) / 1000
                With Data.Recordset
                    .MoveFirst
                    
                    !PD_gain_20 = Round(gainpd(2), 3)
                    .Update
                End With
                
            ElseIf opt(6) Then '........................... current
            
            Select Case cmbcrt.ListIndex
            Case 0
               gaincrt(0) = (((SLD.Max + SLD.Min)) - SLD.Value) / 1000
               With Data.Recordset
                    .MoveFirst
                    
                    !Current_gain_0_1 = Round(gaincrt(0), 3)
                    .Update
                End With
            Case 1
                gaincrt(1) = (((SLD.Max + SLD.Min)) - SLD.Value) / 1000
                With Data.Recordset
                    .MoveFirst
                    
                    !Current_gain_0_15 = Round(gaincrt(1), 3)
                    .Update
                End With
            Case 2
               gaincrt(2) = (((SLD.Max + SLD.Min)) - SLD.Value) / 1000
               With Data.Recordset
                    .MoveFirst
                    
                    !Current_gain_0_2 = Round(gaincrt(2), 3)
                    .Update
                End With
            Case 3
                gaincrt(3) = (((SLD.Max + SLD.Min)) - SLD.Value) / 1000
                With Data.Recordset
                    .MoveFirst
                    
                    !Current_gain_0_5 = Round(gaincrt(3), 3)
                    .Update
                End With
            Case 4
                gaincrt(4) = (((SLD.Max + SLD.Min)) - SLD.Value) / 1000
                With Data.Recordset
                    .MoveFirst
                    
                    !Current_gain_0_7 = Round(gaincrt(4), 3)
                    .Update
                End With
            Case 5
                gaincrt(5) = (((SLD.Max + SLD.Min)) - SLD.Value) / 1000
                With Data.Recordset
                    .MoveFirst

                    !Current_gain_1 = Round(gaincrt(5), 3)
                    .Update
                End With
            Case 6
                gaincrt(6) = (((SLD.Max + SLD.Min)) - SLD.Value) / 1000
                With Data.Recordset
                    .MoveFirst
                    
                    !Current_gain_1_2 = Round(gaincrt(6), 3)
                    .Update
                End With
            Case 7
                gaincrt(7) = (((SLD.Max + SLD.Min)) - SLD.Value) / 1000
                With Data.Recordset
                    .MoveFirst
                    
                    !Current_gain_1_5 = Round(gaincrt(7), 3)
                    .Update
                End With
            Case 8
                gaincrt(8) = (((SLD.Max + SLD.Min)) - SLD.Value) / 1000
                With Data.Recordset
                    .MoveFirst
                    
                    !Current_gain_2 = Round(gaincrt(8), 3)
                    .Update
                End With
            Case 9
                gaincrt(9) = (((SLD.Max + SLD.Min)) - SLD.Value) / 1000
                With Data.Recordset
                    .MoveFirst
                   
                    !Current_gain_2_5 = Round(gaincrt(9), 3)
                    .Update
                End With
            Case 10
                gaincrt(10) = (((SLD.Max + SLD.Min)) - SLD.Value) / 1000
                With Data.Recordset
                    .MoveFirst
                    
                    !Current_gain_3 = Round(gaincrt(10), 3)
                    .Update
                End With
            Case 11
                gaincrt(11) = (((SLD.Max + SLD.Min)) - SLD.Value) / 1000
                With Data.Recordset
                    .MoveFirst
                    
                    !Current_gain_3_5 = Round(gaincrt(11), 3)
                    .Update
                End With
             Case 12
                gaincrt(12) = (((SLD.Max + SLD.Min)) - SLD.Value) / 1000
                With Data.Recordset
                    .MoveFirst
                    
                    !Current_gain_4 = Round(gaincrt(12), 3)
                    .Update
                End With
              Case 13
                gaincrt(13) = (((SLD.Max + SLD.Min)) - SLD.Value) / 1000
                With Data.Recordset
                    .MoveFirst
                    
                    !Current_gain_4_5 = Round(gaincrt(13), 3)
                    .Update
                End With
              Case 14
                gaincrt(14) = (((SLD.Max + SLD.Min)) - SLD.Value) / 1000
                With Data.Recordset
                    .MoveFirst
                    
                    !Current_gain_5 = Round(gaincrt(14), 3)
                    .Update
                End With
            End Select
         End If
            
        End If
    
    End If
End If
End Sub



Private Sub Form_Activate()
TXTGAIN.Text = Str((((SLD.Max + SLD.Min)) - SLD.Value) / 1000)
If Tab0.Tab = 0 Then
    TXTGAIN.Visible = True
    lblgain.Visible = True
    SLD.Visible = True
    
Else
    TXTGAIN.Visible = False
    lblgain.Visible = False
    SLD.Visible = False
    
End If

End Sub

Private Sub Form_Load()
Dim gnNumOfDevices As Integer
Dim MaxEntries As Integer
Dim nOutEntries As Integer
Dim i, ii As Integer
Dim tt As Long
Dim tempStr As String
Dim Response
Dim tempNum As Integer
'Dim lpDEVCONFIG_AI As DEVCONFIG_AI


'*****************************************8
Call formscaling

With Data
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.acldb + ";Jet OLEDB:Database Password=" + mdl1.AclDBPas
    .CommandType = adCmdTable
    .RecordSource = "gain"
    .Refresh
    .Recordset.MoveFirst
End With


FRMCAL.icon = LoadPicture(mdl1.icon)
gainvlt(0) = Data.Recordset.Fields("Voltage_gain_10_2")
gainvlt(1) = Data.Recordset.Fields("Voltage_gain_10_6")
gainvlt(2) = Data.Recordset.Fields("Voltage_gain_10_10")
gainvlt(3) = Data.Recordset.Fields("Voltage_gain_50_10")
gainvlt(4) = Data.Recordset.Fields("Voltage_gain_50_25")
gainvlt(5) = Data.Recordset.Fields("Voltage_gain_50_50")
gainvlt(6) = Data.Recordset.Fields("Voltage_gain_100_30")
gainvlt(7) = Data.Recordset.Fields("Voltage_gain_100_60")
gainvlt(8) = Data.Recordset.Fields("Voltage_gain_100_100")

    gainvlt(9) = Data.Recordset.Fields("Voltage_gain_132_10") 'added m.r
    gainvlt(10) = Data.Recordset.Fields("Voltage_gain_132_66") 'added m.r
    gainvlt(11) = Data.Recordset.Fields("Voltage_gain_132_132") 'added m.r
            
    gainvlt(12) = Data.Recordset.Fields("Voltage_gain_230_10") 'added m.r
    gainvlt(13) = Data.Recordset.Fields("Voltage_gain_230_115") 'added m.r
    gainvlt(14) = Data.Recordset.Fields("Voltage_gain_230_230") 'added m.r
            
    gainvlt(15) = Data.Recordset.Fields("Voltage_gain_400_10") 'added m.r
    gainvlt(16) = Data.Recordset.Fields("Voltage_gain_400_200") 'added m.r
    gainvlt(17) = Data.Recordset.Fields("Voltage_gain_400_400") 'added m.r
    
gaincrt(0) = Data.Recordset.Fields("Current_gain_0_1")
gaincrt(1) = Data.Recordset.Fields("Current_gain_0_15")
gaincrt(2) = Data.Recordset.Fields("Current_gain_0_2")
gaincrt(3) = Data.Recordset.Fields("Current_gain_0_5")
gaincrt(4) = Data.Recordset.Fields("Current_gain_0_7")
gaincrt(5) = Data.Recordset.Fields("Current_gain_1")
gaincrt(6) = Data.Recordset.Fields("Current_gain_1_2")
gaincrt(7) = Data.Recordset.Fields("Current_gain_1_5")
gaincrt(8) = Data.Recordset.Fields("Current_gain_2")
gaincrt(9) = Data.Recordset.Fields("Current_gain_2_5")
gaincrt(10) = Data.Recordset.Fields("Current_gain_3")
gaincrt(11) = Data.Recordset.Fields("Current_gain_3_5")
gaincrt(12) = Data.Recordset.Fields("Current_gain_4")
gaincrt(13) = Data.Recordset.Fields("Current_gain_4_5")
gaincrt(14) = Data.Recordset.Fields("Current_gain_5")

gainpd(0) = Data.Recordset.Fields("PD_gain_2")
gainpd(1) = Data.Recordset.Fields("PD_gain_10")
gainpd(2) = Data.Recordset.Fields("PD_gain_20")



SLD.Min = (gainvlt(3) - 1) * 1000
SLD.Max = (gainvlt(3) + 1) * 1000
SLD.Value = (SLD.Min + SLD.Max) / 2
cmbcrt.Enabled = False
cmbvlt.Enabled = True

'////////////////            I/O CARD SETTING          //////////////

cmdstart.Caption = "START"
'CMDSAVE.Enabled = False
t0.Enabled = False
t1.Enabled = False
 MaxEntries = 100
  ' Add type of PC Laboratory Card

    '//////>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>.
'///////////////////////////////

With cmbvd
    .AddItem "5 mV", 0
    .AddItem "50 mV", 1
    .AddItem "0.1 V", 2
    .AddItem "0.2 V", 3
    .AddItem "0.5 V", 4
    .AddItem "1 V", 5
    .AddItem "2 V", 6
    .AddItem "5 V", 7
    .ListIndex = 4
End With
With cmbcon
    .AddItem "RMS", 0
    .AddItem "PEAK", 1
    .ListIndex = 0
End With
cond = 2 ^ 0.5
optc(1).Value = True
opt(1).Value = True
With cmbcrt
    .AddItem "0.1 mA", 0
    .AddItem "0.15 mA", 1
    .AddItem "0.2 mA", 2
    .AddItem "0.5 mA", 3
    .AddItem "0.7 mA", 4
    .AddItem "1 mA", 5
    .AddItem "1.2 mA", 6
    .AddItem "1.5 mA", 7
    .AddItem "2 mA", 8
    .AddItem "2.5 mA", 9
    .AddItem "3 mA", 10
    .AddItem "3.5 mA", 11
    .AddItem "4 mA", 12
    .AddItem "4.5 mA", 13
    .AddItem "5 mA", 14
    .ListIndex = 0
End With
opt(1).Value = True
With cmbvlt
    .AddItem "10 KV", 0
    .AddItem "25 KV", 1
    .AddItem "50 KV", 2
    .ListIndex = 0
End With
'///////////////////
With pic
    .ScaleHeight = .Height
    .ScaleWidth = .Width
End With
With lh1
    .x1 = 0
    .y1 = pic.ScaleHeight * 1 / 8
    .x2 = pic.ScaleWidth
    .y2 = pic.ScaleHeight * 1 / 8
End With

With lh2
    .x1 = 0
    .y1 = pic.ScaleHeight * 2 / 8
    .x2 = pic.ScaleWidth
    .y2 = pic.ScaleHeight * 2 / 8
End With
With lh3
    .x1 = 0
    .y1 = pic.ScaleHeight * 3 / 8
    .x2 = pic.ScaleWidth
    .y2 = pic.ScaleHeight * 3 / 8
End With
With lh4
    .x1 = 0
    .y1 = pic.ScaleHeight * 4 / 8
    .x2 = pic.ScaleWidth
    .y2 = pic.ScaleHeight * 4 / 8
    .BorderWidth = 2
End With
With lh5
    .x1 = 0
    .y1 = pic.ScaleHeight * 5 / 8
    .x2 = pic.ScaleWidth
    .y2 = pic.ScaleHeight * 5 / 8
End With
With lh6
    .x1 = 0
    .y1 = pic.ScaleHeight * 6 / 8
    .x2 = pic.ScaleWidth
    .y2 = pic.ScaleHeight * 6 / 8
End With
With lh7
    .x1 = 0
    .y1 = pic.ScaleHeight * 7 / 8
    .x2 = pic.ScaleWidth
    .y2 = pic.ScaleHeight * 7 / 8
End With
With lv1
    .x1 = pic.ScaleWidth * 1 / 10
    .x2 = pic.ScaleWidth * 1 / 10
    .y2 = pic.ScaleHeight
    .y1 = 0
End With
With lv2
    .x1 = pic.ScaleWidth * 2 / 10
    .x2 = pic.ScaleWidth * 2 / 10
    .y2 = pic.ScaleHeight
    .y1 = 0
End With
With lv3
    .x1 = pic.ScaleWidth * 3 / 10
    .x2 = pic.ScaleWidth * 3 / 10
    .y2 = pic.ScaleHeight
    .y1 = 0
End With

With lv4
    .x1 = pic.ScaleWidth * 4 / 10
    .x2 = pic.ScaleWidth * 4 / 10
    .y2 = pic.ScaleHeight
    .y1 = 0
End With

With lv5
   .x1 = pic.ScaleWidth * 5 / 10
    .x2 = pic.ScaleWidth * 5 / 10
    .y2 = pic.ScaleHeight
    .y1 = 0
    .BorderWidth = 2
End With

With lv6
    .x1 = pic.ScaleWidth * 6 / 10
    .x2 = pic.ScaleWidth * 6 / 10
    .y2 = pic.ScaleHeight
    .y1 = 0
End With

With lv7
    .x1 = pic.ScaleWidth * 7 / 10
    .x2 = pic.ScaleWidth * 7 / 10
    .y2 = pic.ScaleHeight
    .y1 = 0
End With
With lv8
    .x1 = pic.ScaleWidth * 8 / 10
    .x2 = pic.ScaleWidth * 8 / 10
    .y2 = pic.ScaleHeight
    .y1 = 0
End With
With lv9
    .x1 = pic.ScaleWidth * 9 / 10
    .x2 = pic.ScaleWidth * 9 / 10
    .y2 = pic.ScaleHeight
    .y1 = 0
End With
Tab0.Tab = 0
Call tab0_Click(0)
pic.AutoRedraw = True
TXTGAIN.Text = Str((((SLD.Max + SLD.Min)) - SLD.Value) / 1000)

End Sub

Private Sub opt_Click(Index As Integer)
FRMCAL.MousePointer = 13

If opt(0).Value = True Then 'voltage 0- 10 KV
cmbcon.Enabled = True
   With cmbvlt
        .Clear
        .AddItem "2 KV", 0
        .AddItem "6 KV", 1
        .AddItem "10 KV", 2
        .ListIndex = 0
    End With
    cmbcrt.Enabled = False
    cmbvlt.Enabled = True
    cmbcon.ListIndex = 0
    cmbcon_Click
     LBLVOLT.Caption = " VOLTAGE ="
     lbl0.Caption = "[kv]"
     
ElseIf opt(1).Value = True Then  ' VOLTAGE 0 _ 50 KV
cmbcon.Enabled = True
    With cmbvlt
        .Clear
        .AddItem "10 KV", 0
        .AddItem "25 KV", 1
        .AddItem "50 KV", 2
        .ListIndex = 0
    End With
    cmbcrt.Enabled = False
    cmbvlt.Enabled = True
    cmbcon.ListIndex = 0
    cmbcon_Click
    LBLVOLT.Caption = " VOLTAGE ="
    lbl0.Caption = "[kv]"
     
ElseIf opt(2).Value = True Then ' VOLTAGE 0_ 150 KV
cmbcon.Enabled = True
    With cmbvlt
        .Clear
        .AddItem "30 KV", 0
        .AddItem "60 KV", 1
        .AddItem "100 KV", 2
        .ListIndex = 0
    End With
    cmbcrt.Enabled = False
    cmbvlt.Enabled = True
    cmbcon.ListIndex = 0
    cmbcon_Click
    LBLVOLT.Caption = " VOLTAGE ="
    lbl0.Caption = "[kv]"
ElseIf opt(7).Value = True Then ' VOLTAGE 0_ 132 KV m.r added
cmbcon.Enabled = True
    With cmbvlt
        .Clear
        .AddItem "10 KV", 0
        .AddItem "66 KV", 1
        .AddItem "132 KV", 2
        .ListIndex = 0
    End With
    cmbcrt.Enabled = False
    cmbvlt.Enabled = True
    cmbcon.ListIndex = 0
    cmbcon_Click
    LBLVOLT.Caption = " VOLTAGE ="
    lbl0.Caption = "[kv]"
ElseIf opt(8).Value = True Then ' VOLTAGE 0_230 KV m.r added
cmbcon.Enabled = True
    With cmbvlt
        .Clear
        .AddItem "10 KV", 0
        .AddItem "115 KV", 1
        .AddItem "230 KV", 2
        .ListIndex = 0
    End With
    cmbcrt.Enabled = False
    cmbvlt.Enabled = True
    cmbcon.ListIndex = 0
    cmbcon_Click
    LBLVOLT.Caption = " VOLTAGE ="
    lbl0.Caption = "[kv]"
ElseIf opt(9).Value = True Then ' VOLTAGE 0_400 KV m.r added
cmbcon.Enabled = True
    With cmbvlt
        .Clear
        .AddItem "10 KV", 0
        .AddItem "200 KV", 1
        .AddItem "400 KV", 2
        .ListIndex = 0
    End With
    cmbcrt.Enabled = False
    cmbvlt.Enabled = True
    cmbcon.ListIndex = 0
    cmbcon_Click
    LBLVOLT.Caption = " VOLTAGE ="
    lbl0.Caption = "[kv]"
ElseIf opt(3).Value = True Then    ' PD 0-2 Pc
    cmbcon.ListIndex = 1
    cmbcon.Enabled = False
    cmbcrt.Enabled = False
    cmbvlt.Enabled = False
    SLD.Max = 500000
    SLD.Min = (gainpd(0) - 1) * 1000
    SLD.Max = (gainpd(0) + 1) * 1000
    cmbcon.ListIndex = 0
    cmbcon_Click
    LBLVOLT.Caption = " PD ="
    lbl0.Caption = "[Pc]"
     
ElseIf opt(4).Value = True Then  ' PD 0 - 5 Pc
cmbcon.Enabled = False
    cmbcon.ListIndex = 1
    cmbcrt.Enabled = False
    cmbvlt.Enabled = False
    SLD.Max = 500000
    SLD.Min = (gainpd(1) - 1) * 1000
    SLD.Max = (gainpd(1) + 1) * 1000
    cmbcon.ListIndex = 0
    cmbcon_Click
    LBLVOLT.Caption = " PD ="
    lbl0.Caption = "[Pc]"
     
ElseIf opt(5).Value = True Then 'PD 0 - 20 Pc
    cmbcon.ListIndex = 1
    cmbcon.Enabled = False
    cmbcrt.Enabled = False
    cmbvlt.Enabled = False
    SLD.Max = 500000
    SLD.Min = (gainpd(2) - 1) * 1000
    SLD.Max = (gainpd(2) + 1) * 1000
    cmbcon.ListIndex = 0
    cmbcon_Click
    LBLVOLT.Caption = " PD ="
    lbl0.Caption = "[Pc]"
    
ElseIf opt(6).Value = True Then   ' CURRENT 0 _10 mA

    cmbcon.Enabled = True
    cmbcrt.Enabled = True
    cmbcrt.ListIndex = 0
    cmbcrt_Click
    cmbvlt.Enabled = False
    cmbcon.ListIndex = 1
    cmbcon_Click
    LBLVOLT.Caption = " CURRENT ="
    lbl0.Caption = "[mA]"
End If
SLD.Value = (SLD.Min + SLD.Max) / 2
TXTGAIN.Text = Str((((SLD.Max + SLD.Min)) - SLD.Value) / 1000)
FRMCAL.MousePointer = 0

LBLVOLT.ForeColor = optc(Index).ForeColor
lbl0.ForeColor = optc(Index).ForeColor
TXTVOL.ForeColor = optc(Index).ForeColor

End Sub

Private Sub optc_Click(Index As Integer)
'If t1.Enabled = True Then
    ' ErrCde = DRV_FAIStop(DeviceHandle)  'Stop AI conversion
    'If (ChkErr(ErrCde) = 1) Then
       'DRV_DeviceClose (DeviceHandle)
       'Exit Sub
    'End If
    't1.Enabled = False
'End If
If optc(0).Value = True Or optc(1).Value = True Or optc(2).Value = True Or optc(7).Value = True Or optc(8).Value = True Or optc(9).Value = True Then
    cmbcon.Enabled = True
    cmbcon.ListIndex = 0
    cmbcon_Click
    LBLVOLT.Caption = " VOLTAGE = "
    lbl0.Caption = "[KV]"
ElseIf optc(6).Value = True Then
    
    cmbcon.Enabled = True
    cmbcon.ListIndex = 1
    cmbcon_Click
     LBLVOLT.Caption = " CURRENT ="
     lbl0.Caption = "[mA]"
Else
    cmbcon.ListIndex = 1
    cmbcon.Enabled = False
    LBLVOLT.Caption = " PD ="
    lbl0.Caption = "[Pc]"
    
End If

LBLVOLT.ForeColor = optc(Index).ForeColor
lbl0.ForeColor = optc(Index).ForeColor
TXTVOL.ForeColor = optc(Index).ForeColor

End Sub

Private Sub SLD_Scroll()
TXTGAIN.Text = Str((((SLD.Max + SLD.Min)) - SLD.Value) / 1000)
End Sub


Private Sub t0_Timer()
'3.2 Call advantech API function DRV_FAIIntScanStart
On Error GoTo errhand

    
err = VIBuffer.Prepare
If err <> 0 Then
    
    fram1.Enabled = True
    Tab0.Enabled = True
    cmdstart.Caption = "START"
    t0.Enabled = False
    Exit Sub
End If
   
'2. Set gain list and other setting
    
Call selectgain
   
ReDim sDataBuffer(0 To VIBuffer.ScanChannel.Samples)     'Allocate memory
  
       
err = VIBuffer.RunOnce ' start data acquisition
    
If err <> 0 Then
    t0.Enabled = False
    err = VIBuffer.Stop
    Exit Sub
End If
    

Exit Sub

errhand:
MsgBox "error in reading data", vbOKOnly, "err"

'//////////////////////////////////////////////
End Sub

Private Sub tab0_Click(PreviousTab As Integer)
Dim i As Integer
If Tab0.Tab = 0 Then
    TXTGAIN.Visible = True
    lblgain.Visible = True
    SLD.Enabled = True
    CMDSET.Enabled = True
    fram1.Visible = True
    For i = 0 To 9
        optc(i).Visible = False
    Next i
    For i = 0 To 9
        opt(i).Visible = True
    Next i
    opt(0).Value = True
    Call opt_Click(0)
Else

    TXTGAIN.Visible = False
    lblgain.Visible = False
    SLD.Enabled = False
    CMDSET.Enabled = False
    fram1.Visible = False
    
    For i = 0 To 9
        optc(i).Visible = True
    Next i
    For i = 0 To 9
        opt(i).Visible = False
    Next i
    optc(0).Value = True
    Call optc_Click(0)
End If

End Sub


Private Sub TXTGAIN_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub TXTVOL_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub VIBuffer_Stopped(ByVal Offset As Long, ByVal Count As Long)
Dim i As Long

'///////////     Check current status    //////////////////////

err = VIBuffer.GetData(Count, sDataBuffer())

If err <> 0 Then
    t0.Enabled = False
     MsgBox "Error In Reading Data From I/O card" + Chr(13) + "Please Check I/O card Setting And wirring.", vbExclamation, "Error"
    Exit Sub
End If

err = VIBuffer.Stop

If (err <> 0) Then
     MsgBox "Error in Stopping I/O card", vbOKOnly, "Error!!"
End If
    
    '*********    Draw data in image box to show screen and detect peak value  ******
Call drawcurve
    
t0.Enabled = True

End Sub


