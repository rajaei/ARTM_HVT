VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{1752FF26-D6C9-4BC8-BFE9-7D0CA26DED89}#1.0#0"; "BDaqOcx.dll"
Begin VB.Form frmtest 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HV _TEST "
   ClientHeight    =   10800
   ClientLeft      =   1365
   ClientTop       =   1185
   ClientWidth     =   20310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   16259.7
   ScaleMode       =   0  'User
   ScaleWidth      =   27720.39
   Begin VB.CheckBox chkIres 
      BackColor       =   &H8000000A&
      Caption         =   "Show Current Components"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   4680
      TabIndex        =   67
      Top             =   8880
      Width           =   4095
   End
   Begin VB.CommandButton CmdCalibration 
      Caption         =   "Calibration"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13680
      TabIndex        =   66
      Top             =   9240
      Width           =   1455
   End
   Begin VB.TextBox txtch4 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   435
      Left            =   11400
      TabIndex        =   65
      TabStop         =   0   'False
      Text            =   "0.255"
      ToolTipText     =   " Ê«‰  ·›« Ì »—ﬁêÌ—"
      Top             =   3240
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc DataArrtype 
      Height          =   330
      Left            =   13320
      Top             =   1320
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
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
      Caption         =   "DataArrtype"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc datatestprmt 
      Height          =   330
      Left            =   13440
      Top             =   720
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
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
      Caption         =   "datatestprmt"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc datatest 
      Height          =   330
      Left            =   13320
      Top             =   240
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
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
      Caption         =   "datatest"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtch3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   435
      Left            =   8040
      TabIndex        =   59
      TabStop         =   0   'False
      Text            =   "10.2"
      ToolTipText     =   "„Ê·›Â «Â„Ì Ã—Ì«‰ »—ﬁêÌ—"
      Top             =   3240
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   10440
      Top             =   9480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer timchck 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   19800
      Top             =   10440
   End
   Begin VB.CommandButton cmdstoptest 
      BackColor       =   &H8000000A&
      Caption         =   "&Stop Test"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13440
      TabIndex        =   8
      Top             =   3360
      Width           =   1095
   End
   Begin VB.PictureBox picright 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      FillColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   4335
      Left            =   11640
      ScaleHeight     =   4335
      ScaleWidth      =   375
      TabIndex        =   54
      Top             =   3960
      Width           =   375
   End
   Begin VB.PictureBox picleft 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      FillColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4215
      Left            =   840
      ScaleHeight     =   4215
      ScaleWidth      =   495
      TabIndex        =   53
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtch2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   5160
      TabIndex        =   52
      TabStop         =   0   'False
      Text            =   "1.25"
      ToolTipText     =   "ò· Ã—Ì«‰ ⁄»Ê—Ì «“ »—ﬁêÌ—"
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox txtch1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2160
      TabIndex        =   51
      TabStop         =   0   'False
      Text            =   "22.5"
      ToolTipText     =   "Ê· «é «⁄„«·Ì »Â »—ﬁêÌ—"
      Top             =   3360
      Width           =   975
   End
   Begin VB.Timer plottime 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   13920
      Top             =   2640
   End
   Begin VB.CommandButton cmdend 
      Caption         =   "END "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12000
      TabIndex        =   15
      Top             =   9240
      Width           =   1455
   End
   Begin VB.Frame framosc 
      BackColor       =   &H8000000A&
      Caption         =   "graph Setting"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   12360
      TabIndex        =   44
      Top             =   3960
      Width           =   2775
      Begin VB.ComboBox CmbPDSel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   450
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   4080
         Width           =   1095
      End
      Begin VB.ComboBox cmbosc2p 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1215
      End
      Begin VB.ComboBox cmbvsel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   450
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   4080
         Width           =   1095
      End
      Begin VB.CommandButton cmdclr1 
         BackColor       =   &H8000000A&
         Caption         =   "CH 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2400
         Width           =   615
      End
      Begin VB.CommandButton cmdclrback 
         Caption         =   "Back Ground"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CommandButton cmdclr2 
         BackColor       =   &H8000000A&
         Caption         =   "CH2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2400
         Width           =   615
      End
      Begin VB.ComboBox cmbosc2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox cmbosc1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label LblPDSel 
         BackColor       =   &H8000000A&
         Caption         =   "PD Level"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1560
         TabIndex        =   62
         Top             =   3600
         Width           =   975
      End
      Begin VB.Line ln1 
         X1              =   0
         X2              =   2640
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line ln2 
         X1              =   0
         X2              =   2640
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label lblosc2p 
         BackColor       =   &H8000000A&
         Caption         =   "PD/Div"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblvoltsel 
         BackColor       =   &H8000000A&
         Caption         =   "Voltage Level"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label lblclr 
         BackColor       =   &H8000000A&
         Caption         =   "Color Setting"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   47
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label lblosc2 
         BackColor       =   &H8000000A&
         Caption         =   "Crnt/Div"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   46
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblosc1 
         BackColor       =   &H8000000A&
         Caption         =   "Volt/Div"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.PictureBox pic1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      Height          =   4335
      Left            =   1320
      ScaleHeight     =   4275
      ScaleWidth      =   10155
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   4080
      Width           =   10215
      Begin VB.Line lv9 
         X1              =   9480
         X2              =   9480
         Y1              =   0
         Y2              =   5280
      End
      Begin VB.Line lv8 
         X1              =   8280
         X2              =   8280
         Y1              =   0
         Y2              =   5280
      End
      Begin VB.Line lv7 
         X1              =   7200
         X2              =   7200
         Y1              =   0
         Y2              =   5280
      End
      Begin VB.Line lv6 
         X1              =   6240
         X2              =   6240
         Y1              =   0
         Y2              =   5280
      End
      Begin VB.Line lv5 
         BorderWidth     =   2
         X1              =   5160
         X2              =   5160
         Y1              =   0
         Y2              =   5280
      End
      Begin VB.Line lv4 
         X1              =   3720
         X2              =   3720
         Y1              =   0
         Y2              =   5280
      End
      Begin VB.Line lh7 
         X1              =   0
         X2              =   10800
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line lh6 
         X1              =   0
         X2              =   10800
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line lh5 
         BorderColor     =   &H00000000&
         X1              =   0
         X2              =   10800
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line lh4 
         BorderColor     =   &H80000006&
         BorderWidth     =   2
         X1              =   0
         X2              =   10800
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line lv3 
         X1              =   3120
         X2              =   3120
         Y1              =   0
         Y2              =   5280
      End
      Begin VB.Line lv2 
         X1              =   2280
         X2              =   2280
         Y1              =   120
         Y2              =   5400
      End
      Begin VB.Line lv1 
         X1              =   960
         X2              =   960
         Y1              =   0
         Y2              =   5280
      End
      Begin VB.Line lh3 
         BorderColor     =   &H00000000&
         X1              =   0
         X2              =   10800
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line lh1 
         BorderColor     =   &H00000000&
         X1              =   0
         X2              =   10800
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line lh2 
         X1              =   0
         X2              =   10800
         Y1              =   1440
         Y2              =   1440
      End
   End
   Begin VB.Frame framtest2 
      BackColor       =   &H8000000A&
      Caption         =   "   PD  @  Ur"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2175
      Left            =   4920
      TabIndex        =   21
      Top             =   600
      Width           =   2055
      Begin VB.TextBox txtpd2 
         Alignment       =   2  'Center
         DataField       =   "PD@Ur"
         DataSource      =   " "
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
         Height          =   375
         Left            =   720
         TabIndex        =   38
         TabStop         =   0   'False
         Text            =   " "
         ToolTipText     =   " Œ·ÌÂ Ã“ÌÌ œ— Ê· «é ‰«„Ì"
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdtestok2 
         BackColor       =   &H8000000A&
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   5
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lblpdu2 
         BackColor       =   &H8000000A&
         Caption         =   "[pC]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   40
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblpdt2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000A&
         Caption         =   "PD :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblpdacp2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "Accp.PD  = 0.2  [Pc]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lblur2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "Ur=25.12 [KV]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame framtest3 
      BackColor       =   &H8000000A&
      Caption         =   "   PD  @ 1.05  *  Uc"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2295
      Left            =   7200
      TabIndex        =   20
      Top             =   600
      Width           =   2775
      Begin VB.TextBox txtpd3 
         Alignment       =   2  'Center
         DataField       =   "PD@Uc"
         DataSource      =   " "
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
         Height          =   375
         Left            =   840
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   " "
         ToolTipText     =   " Œ·ÌÂ Ã“ÌÌ œ— 1.05 Ê· «é ò«— œ«Ì„"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdtestok3 
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label lblpdu3 
         BackColor       =   &H8000000A&
         Caption         =   "[pC]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   35
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lblpdt3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000A&
         Caption         =   "PD :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label lblpdacp3 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "Accp.PD  = 0.2  [Pc]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lbluc3 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "1.05 * Uc =25 [KV]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   31
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame framtest4 
      BackColor       =   &H8000000A&
      Caption         =   "Leackage Current"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2535
      Left            =   10200
      TabIndex        =   19
      Top             =   360
      Width           =   3015
      Begin VB.TextBox txtic4 
         Alignment       =   2  'Center
         DataField       =   "Iref@Uref"
         DataSource      =   " "
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
         Height          =   285
         Left            =   600
         TabIndex        =   30
         TabStop         =   0   'False
         Text            =   " "
         ToolTipText     =   "Ã—Ì«‰ ‰‘ Ì »—ﬁêÌ—"
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdtestok4 
         BackColor       =   &H8000000A&
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   7
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblicu4 
         BackColor       =   &H8000000A&
         Caption         =   " [mA]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   29
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblict4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000A&
         Caption         =   "Ic :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lblic4 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   " Ic = 0.3  [mA]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label lbluref4 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "Uc =  24.5  [KV]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame framsrn 
      BackColor       =   &H8000000A&
      Caption         =   "Serial Number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   18
      Top             =   360
      Width           =   1935
      Begin VB.CommandButton cmdsnpre 
         BackColor       =   &H8000000A&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   1200
         Width           =   375
      End
      Begin VB.ComboBox cmbtype 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         TabIndex        =   0
         Text            =   "Combo1"
         ToolTipText     =   "„œ· »—ﬁêÌ— —« «‰ Œ«» ò‰Ìœ"
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton cmdsnnext 
         BackColor       =   &H8000000A&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton cmdsnreg 
         BackColor       =   &H8000000A&
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtsn 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   1
         ToolTipText     =   "‘„«—Â ”—Ì«· »—ﬁêÌ—  Õ   ” "
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblsp 
         BackColor       =   &H8000000A&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame framtest1 
      BackColor       =   &H8000000A&
      Caption         =   "Reference Voltage"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2295
      Left            =   2280
      TabIndex        =   17
      Top             =   600
      Width           =   2295
      Begin VB.TextBox txturef1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "Uref@Iref"
         DataSource      =   " "
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
         Height          =   375
         Left            =   600
         TabIndex        =   42
         TabStop         =   0   'False
         ToolTipText     =   "Ê· «é „—Ã⁄ À»  ‘œÂ »—ﬁêÌ—"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdtestok1 
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label lblurefu1 
         BackColor       =   &H8000000A&
         Caption         =   "[kV]"
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
         Left            =   1800
         TabIndex        =   25
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label lblureft1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000A&
         Caption         =   "Uref :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lbluref1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   " 24.5 < Uref <29.6  [KV]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lbliref1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   " iref="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   1455
      End
   End
   Begin BDaqOcxLibCtl.BufferedAiCtrl VIBuffer 
      Left            =   14520
      OleObjectBlob   =   "Test.frx":0000
      Top             =   2775
   End
   Begin VB.Label lblarrtype 
      BackColor       =   &H8000000A&
      Caption         =   "Type: ETM25-H10B"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   64
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label lblch4 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "P loss [mW] :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   9960
      TabIndex        =   60
      Top             =   3360
      Width           =   1500
   End
   Begin VB.Label lblch3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "Ires [uA] :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   6840
      TabIndex        =   58
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label lblcurpd 
      BackColor       =   &H8000000A&
      Caption         =   "Current [mA]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   9240
      TabIndex        =   56
      Top             =   8880
      Width           =   1815
   End
   Begin VB.Label lblvol 
      BackColor       =   &H8000000A&
      Caption         =   "Voltage  [kV]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1200
      TabIndex        =   55
      Top             =   8760
      Width           =   2535
   End
   Begin VB.Label lblch2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "current [mA] :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   3720
      TabIndex        =   50
      Top             =   3360
      Width           =   1560
   End
   Begin VB.Label lblch1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000A&
      Caption         =   "voltage [kV] :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   600
      TabIndex        =   49
      Top             =   3480
      Width           =   1500
   End
End
Attribute VB_Name = "frmtest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private arrcnt As Double
Private temselectedarr As String
Private rtv, timecnt As Double
Private voltgain, pdgain, currentgain, ELoss, PLoss, PLoss_mWatt, ArresterCap As Double
Private dV_dt(1 To 6), ArreyNo(1 To 6) As Single
Private Iresistive(0 To 402), Icapasitive(0 To 402) As Double
Private p1, p2 As Double
Private sampok, testapp1, testapp2, testapp3, testapp4, scopopen, testallow, draw As Boolean
Public iref1, vref1, vpd1, vpd2, vref2, ic2, icacp4, Unit_No As Double
Private prepd, precurrent, prevoltage As Double
Private preych1, prexch1, preych2, HVtest_Repetation_No As Double


Dim sDataBufferMain(0 To 4810) As Double '  data buffer for I/O card for voltage and Current in Test 1 and 4
Dim sDataBufferMain2(0 To 2410) As Double ' data buffer for I/O card for voltage in Test 2 and 3
Dim sDataBuffervlt(0 To 1210) As Double '  data buffer for I/O card for voltage
Dim sDataBuffercrt(0 To 1210) As Double '  data buffer for I/O card for current
Dim sDataBufferpd(0 To 14410) As Double '  data buffer for I/O card for PD
Dim sDataBuffer(0 To 200) As Double      '  data buffer for I/O card in reset Condition
Dim err  As ErrorCode
Dim IOCartState As Integer

'Private lFifoSize As Long
'Dim iRet As Integer
'Dim ptAICheck As PT_FAICheck
'Dim lStopped As Long
'Dim iRetrieved As Long
'Dim lOverrun As Long
'Dim lHalfReady As Long
'Dim lActiveBuf As Long
'Dim ptFAIData As PT_FAITransfer
'Dim gnNumOfDevices As Integer
'Dim MaxEntries As Integer
'Dim nOutEntries As Integer
Dim tt, ii As Long
'Dim tempStr As String
'Dim dwDeviceNum As Long

Dim DataTabletype() As String

Dim ch1color, ch2color As Double
Dim vltdiv, crtdiv, pddiv As Double
Dim i, j As Double
Dim timstr
Dim dtime
Dim maxv, maxc, maxp As Double
Dim gaincrt(0 To 14) As Double
Dim gainvlt(0 To 8) As Double
Dim gainpd(0 To 2) As Double
Dim ppeakv(1 To 3) As Double
Dim npeakv(1 To 3) As Double
Dim ppeakc(1 To 3) As Double
Dim npeakc(1 To 3) As Double
Dim pdppk, pdnpk, ppkv, npkv, ppkc, npkc As Double
Dim harm3 As Long '3th harmonic of current
Dim SelectedPDGain As Double ' final PD gain
Dim Response
Dim tempNum, state As Integer
Private reset As Boolean
Private tok1, tok2, tok3, tok4, endtest As Boolean
Private stptst As Boolean
Dim alwstrttst As Boolean
Private Function crtgain(ByVal ppk As Single) As Single
Dim m, BaseGain As Single


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

End Function


Private Sub stoptest()

rtv = MsgBox("Do you want to stop test?", vbYesNo + vbExclamation, "Attention")
If rtv = 6 Then
    plottime.Enabled = False
    alwstrttst = False
    mdl1.tested(arrcnt) = False
    framtest1.ForeColor = vbBlack
    lbliref1.ForeColor = vbBlack
    lbluref1.ForeColor = vbBlack
    framtest4.ForeColor = vbBlack
    lbluref4.ForeColor = vbBlack
    lblic4.ForeColor = vbBlack
    framtest3.ForeColor = vbBlack
    lbluc3.ForeColor = vbBlack
    lblpdacp3.ForeColor = vbBlack
    framtest2.ForeColor = vbBlack
    lblpdacp2.ForeColor = vbBlack
    lblur2.ForeColor = vbBlack
    cmdsnreg.Enabled = True
    
   
'/////////////////////////////////////////////////////////////////////////////////////////////
   
    
    cmdtestok1.Enabled = False
    cmdtestok2.Enabled = False
    cmdtestok3.Enabled = False
    cmdtestok4.Enabled = False
    plottime.Enabled = False
    timchck.Enabled = False
    txtsn.Enabled = True
    txtsn.SetFocus
    pic1.Cls
    cmdstoptest.Enabled = False
    stptst = False
    txtpd2.Text = ""
    txturef1.Text = ""
    txtpd3.Text = ""
    txtic4.Text = ""
    txtsn.Text = ""
    txtch1.Text = ""
    txtch2.Text = ""
    txtch3.Text = ""
    txtch4.Text = ""
    endtest = True
    IOCartState = VIBuffer.state
    err = VIBuffer.Stop
    IOCartState = VIBuffer.state
Else
    plottime.Enabled = True
End If
End Sub
Private Sub stoptest1()
    plottime.Enabled = False
    alwstrttst = False
    mdl1.tested(arrcnt) = False
    framtest1.ForeColor = vbBlack
    lbliref1.ForeColor = vbBlack
    lbluref1.ForeColor = vbBlack
    framtest4.ForeColor = vbBlack
    lbluref4.ForeColor = vbBlack
    lblic4.ForeColor = vbBlack
    framtest3.ForeColor = vbBlack
    lbluc3.ForeColor = vbBlack
    lblpdacp3.ForeColor = vbBlack
    framtest2.ForeColor = vbBlack
    lblpdacp2.ForeColor = vbBlack
    lblur2.ForeColor = vbBlack
    cmdsnreg.Enabled = True
    
    
'/////////////////////////////////////////////////////////////////////////////////////////////

    
    
    cmdtestok1.Enabled = False
    cmdtestok2.Enabled = False
    cmdtestok3.Enabled = False
    cmdtestok4.Enabled = False
    plottime.Enabled = False
    timchck.Enabled = False
    txtsn.Enabled = True
    pic1.Cls
    cmdstoptest.Enabled = False
    stptst = False
    txtpd2.Text = ""
    txturef1.Text = ""
    txtpd3.Text = ""
    txtic4.Text = ""
    'txtsn.Text = ""
    txtch1.Text = ""
    txtch2.Text = ""
    txtch3.Text = ""
    txtch4.Text = ""
    endtest = True
    IOCartState = VIBuffer.state
    err = VIBuffer.Stop
    IOCartState = VIBuffer.state
End Sub


Private Sub test2()
Dim tem As Double
Dim allow As Boolean
'Dim ptEnableEvent As PT_EnableEvent
Dim DataSize As Long
tok2 = False
'//////////////////////main cheke //////////////
If testapp2 = False Then
    MsgBox "Please Adjuste The Voltage In " + Str(vpd1) + "  (KV) And then Press OK Button." + Chr(13), vbExclamation, "Error"
    
ElseIf testapp2 = True Then

    tem = Val(txtpd2.Text)
    
    If mdl1.pdaccp < tem Then
        mdl1.trs2 = 2
        framtest2.ForeColor = vbRed
        lblpdacp2.ForeColor = vbRed
        lblur2.ForeColor = vbRed
        MsgBox "PD @ Urated is Higher than Acceptable value.(Upper Value)", vbOKOnly + vbExclamation, "ERROR"
    Else
        mdl1.trs2 = 1
        framtest2.ForeColor = vbGreen
        lblpdacp2.ForeColor = vbGreen
        lblur2.ForeColor = vbGreen
    End If
    
    framtest4.ForeColor = vbBlack
    lbluref4.ForeColor = vbBlack
    lblic4.ForeColor = vbBlack
    framtest3.ForeColor = vbBlue
    lbluc3.ForeColor = vbBlue
    lblpdacp3.ForeColor = vbBlue
   
'/////////////////////////////////////////////////////////////////
    cmdtestok2.Enabled = False
    cmdtestok3.Enabled = True
    cmdtestok3.SetFocus
    draw = False
    state = 1
End If
End Sub

Private Sub test1()
Dim col1, col2, tem As Double
Dim allow As Boolean
'Dim ptEnableEvent As PT_EnableEvent
Dim DataSize As Long
'?????????????????????????/

tok1 = False
allow = True
col1 = ch1color
col2 = ch2color

'//////////////////Read Input Value/////////////////

If testapp1 = False Then
    MsgBox "Please Adjuste The Current In " + Str(iref1) + "  (mA) Then Press OK Button." + Chr(13), vbExclamation, "Error"
    mdl1.trs1 = 0
ElseIf testapp1 = True Then

   
 
    tem = Val(txturef1.Text)
    If (datatestprmt.Recordset.Fields("Urefmin") / datatestprmt.Recordset.Fields("unit_number")) > tem Then
        mdl1.trs1 = 2
        framtest1.ForeColor = vbRed
        lbliref1.ForeColor = vbRed
        lbluref1.ForeColor = vbRed
        MsgBox "Reference Voltage is Lower than Acceptable Value. (Low Value).", vbOKOnly + vbExclamation, "ERROR"
    ElseIf datatestprmt.Recordset.Fields("Urefmax") / datatestprmt.Recordset.Fields("unit_number") < tem Then
            mdl1.trs1 = 3
            framtest1.ForeColor = vbRed
            lbliref1.ForeColor = vbRed
            lbluref1.ForeColor = vbRed
            MsgBox "Reference Voltage is Higher than Acceptable Value. (High Value).", vbOKOnly + vbExclamation, "ERROR"
    Else
            mdl1.trs1 = 1
            framtest1.ForeColor = vbGreen
            lbliref1.ForeColor = vbGreen
            lbluref1.ForeColor = vbGreen
    End If
    
    datatest.Refresh
    framtest4.ForeColor = vbBlack
    lbluref4.ForeColor = vbBlack
    lblic4.ForeColor = vbBlack
    framtest3.ForeColor = vbBlack
    lbluc3.ForeColor = vbBlack
    lblpdacp3.ForeColor = vbBlack
    framtest2.ForeColor = vbBlue
    lblpdacp2.ForeColor = vbBlue
    lblur2.ForeColor = vbBlue
    
  '////////////////////////////////////////   lable setting     ////////////////////////
    lblch2.Caption = "      PD : "
    lblcurpd.Caption = "     PD [pC]"
    lblosc2.Enabled = False
    lblosc2p.Enabled = True
    cmbosc2.Enabled = False
    cmbosc2p.Enabled = True
    cmbosc2p.Text = "0.5 Pc"
    Call mdl1.rightlable(Val(cmbosc2p.Text) * 4)
    cmdtestok1.Enabled = False
    cmdtestok2.Enabled = True
    cmdtestok2.SetFocus
    
    draw = False
    state = 1 ' voltage priority
    reset = False
    ch1color = col1
    ch2color = col2
    
    '**************************************************************************
    
 '///////////////////////////////////////////////////////////////
End If
End Sub

Private Sub test3()
Dim tem As Double
Dim allow As Boolean
tok3 = False
If testapp3 = False Then
    MsgBox "Please Adjuste The Voltage In " + Str(vpd2) + "  (KV) And then Press OK Button.", vbExclamation, "Error"
ElseIf testapp3 = True Then
 
    tem = Val(txtpd3.Text)
    If mdl1.pdaccp < tem Then
        mdl1.trs3 = 2
        framtest3.ForeColor = vbRed
        lbluc3.ForeColor = vbRed
        lblpdacp3.ForeColor = vbRed
        MsgBox "PD @ Ucon is Higher than Acceptable value.(Upper Value); (Upper Value)", vbOKOnly + vbExclamation, "ERROR"
    Else
        mdl1.trs3 = 1
        framtest3.ForeColor = vbGreen
        lbluc3.ForeColor = vbGreen
        lblpdacp3.ForeColor = vbGreen
    End If
    framtest4.ForeColor = vbBlue
    lbluref4.ForeColor = vbBlue
    lblic4.ForeColor = vbBlue
    lblch2.Caption = " Current : "
    lblcurpd.Caption = "Current [mA]"
    lblosc2.Enabled = True
    lblosc2p.Enabled = False
    cmbosc2.Enabled = True
    cmbosc2p.Enabled = False
    Call mdl1.rightlable(Val(cmbosc2.Text) * 4)
    cmdtestok3.Enabled = False
    cmdtestok4.Enabled = True
    cmdtestok4.SetFocus

    draw = False
    state = 1
'///////////////////////////////////////////////////////////
    If datatestprmt.Recordset.Fields("Icaccp") < 0.8 Then
        cmbosc2.Text = "0.2 mA"
      
        Call mdl1.rightlable(0.2 * 4)
    ElseIf datatestprmt.Recordset.Fields("Icaccp") >= 0.8 And datatestprmt.Recordset.Fields("Icaccp") < 1.6 Then
        cmbosc2.Text = "0.4 mA"
        Call mdl1.rightlable(0.4 * 4)
    ElseIf datatestprmt.Recordset.Fields("Icaccp") >= 1.6 And datatestprmt.Recordset.Fields("Icaccp") < 3.2 Then
        cmbosc2.Text = "0.8 mA"
        Call mdl1.rightlable(0.8 * 4)
    ElseIf datatestprmt.Recordset.Fields("Icaccp") >= 3.2 And datatestprmt.Recordset.Fields("Icaccp") < 4 Then
        cmbosc2.Text = "1 mA"
        Call mdl1.rightlable(1 * 4)
    ElseIf datatestprmt.Recordset.Fields("Icaccp") >= 4 And datatestprmt.Recordset.Fields("Icaccp") < 8 Then
        cmbosc2.Text = "2 mA"
         Call mdl1.rightlable(2 * 4)
    ElseIf datatestprmt.Recordset.Fields("Icaccp") >= 8 Then
        cmbosc2.Text = "3 mA"
        Call mdl1.rightlable(3 * 4)
    End If
    cmbosc2_Change
'/////////////////////////////////////////////////////////
End If
End Sub

Private Sub test4()
Dim temallow As Boolean
Dim rtv As Integer
Dim tem As Double
Dim allow As Boolean
pic1.Cls
tok4 = False

If testapp4 = False Then
    MsgBox "Please Adjuste Voltage In " + Str(vref2) + "  (KV) and Then Press OK Button.", vbExclamation, "Error"
    plottime.Enabled = True
ElseIf testapp4 = True Then
   
   endtest = True
   tem = Val(txtic4.Text)
   If datatestprmt.Recordset.Fields("Icaccp") < tem Then
        mdl1.trs4 = 2
        framtest4.ForeColor = vbRed
        lbluref4.ForeColor = vbRed
        lblic4.ForeColor = vbRed
        MsgBox "Leakage Current is Higher than Acceptable Value", vbOKOnly + vbExclamation, "ERROR"
    Else
        mdl1.trs4 = 1
        framtest4.ForeColor = vbGreen
        lbluref4.ForeColor = vbGreen
        lblic4.ForeColor = vbGreen
    End If
'******************* save information **********'''''

If mdl1.trs1 <> 0 And mdl1.trs2 <> 0 And mdl1.trs3 <> 0 And mdl1.trs4 <> 0 Then
    rtv = MsgBox("Do You Want Save This Result?", vbYesNo + vbExclamation, "Attention")
    If rtv = 6 Then
        datatest.Refresh
        With datatest.Recordset
            .MoveFirst
           .Find "Serial_Num=" + txtsn.Text, , adSearchForward, 1
           
        '//////////////////////////////////////////insert value
        
            If HVtest_Repetation_No = 1 Then  ' test 1 has been applied
            
                !IcUc1 = Val(txtic4.Text)
                !pduc1 = Val(txtpd3.Text)
                !pdur1 = Val(txtpd2.Text)
                !Urefiref1 = Val(txturef1.Text)
                '!harmonic=round(harm3,2)     '  3th harmonic in [uA]
                
                If mdl1.trs1 = 1 And mdl1.trs2 = 1 And mdl1.trs3 = 1 And mdl1.trs4 = 1 Then
                     !result1 = 3 ' test1 was OK
                     !Result2 = 2 ' test1 was OK and ready to test 2
                ElseIf mdl1.trs1 = 3 Then
                    !result1 = 4
                ElseIf mdl1.trs1 = 2 Then
                    !result1 = 5
                ElseIf mdl1.trs4 = 2 Then
                    !result1 = 6
                ElseIf mdl1.trs3 = 2 Then
                    !result1 = 7
                ElseIf mdl1.trs2 = 2 Then
                    !result1 = 8
                End If
                
                !Test1_Date = Calender.To_HejriChar(Date, 0) '   Format(Now, "yyyy/mm/dd")
                !Test1_Oprt = mdl1.optname
                
            ElseIf HVtest_Repetation_No = 2 Then  ' test 2 has been applied
            
                !IcUc2 = Val(txtic4.Text)
                !pduc2 = Val(txtpd3.Text)
                !pdur2 = Val(txtpd2.Text)
                !Urefiref2 = Val(txturef1.Text)
                '!harmonic=round(harm3,2)     '  3th harmonic in [uA]
                
                If mdl1.trs1 = 1 And mdl1.trs2 = 1 And mdl1.trs3 = 1 And mdl1.trs4 = 1 Then
                     !Result2 = 3 ' test2 was OK
                ElseIf mdl1.trs1 = 3 Then
                    !Result2 = 4
                ElseIf mdl1.trs1 = 2 Then
                    !Result2 = 5
                ElseIf mdl1.trs4 = 2 Then
                    !Result2 = 6
                ElseIf mdl1.trs3 = 2 Then
                    !Result2 = 7
                ElseIf mdl1.trs2 = 2 Then
                    !Result2 = 8
                End If
                
                !Test2_Date = Calender.To_HejriChar(Date, 0) '   Format(Now, "yyyy/mm/dd")
                !Test2_Oprt = mdl1.optname
                
            End If
                
            .Update
        
            '///////////////////////////////////////////////////
        End With
    End If
   End If
        mdl1.tested(arrcnt) = True
        cmdstoptest.Enabled = False
        framtest1.ForeColor = vbBlack
        lbliref1.ForeColor = vbBlack
        lbluref1.ForeColor = vbBlack
        framtest4.ForeColor = vbBlack
        lbluref4.ForeColor = vbBlack
        lblic4.ForeColor = vbBlack
        framtest3.ForeColor = vbBlack
        lbluc3.ForeColor = vbBlack
        lblpdacp3.ForeColor = vbBlack
        framtest2.ForeColor = vbBlack
        lblpdacp2.ForeColor = vbBlack
        lblur2.ForeColor = vbBlack
        cmdsnreg.Enabled = True
        plottime.Enabled = False
        cmdtestok4.Enabled = False
        txtsn.Enabled = True
        txtsn.Text = Str(Val(txtsn.Text) + 1)
        txtsn.SetFocus
        lbliref1.Caption = "Iref =... [mA]"
        lblarrtype.Caption = "Arrester Type : ....."
        lbluref1.Caption = "  " + " ...< Uref <... " + " [kV]"
        lbluref4.Caption = " Uc =... [kV]"
        lblic4.Caption = "Ic(accp)=... [mA]"
        lbluc3.Caption = "1.05*Uc=... [kV]"
        lblpdacp3.Caption = "Max. PD= ...[pC]"
        lblur2.Caption = "Ur =... [kV]"
        lblpdacp2.Caption = "Max. PD= ...[pC]"
        txturef1.Text = ""
        txtpd2.Text = ""
        txtpd3.Text = ""
        txtic4.Text = ""
        lblch1.Caption = "Voltage [kV]:"
        lblch2.Caption = "Current [mA]:"
        lblcurpd.Caption = "Current [mA]"
        txtch1.Text = ""
        txtch2.Text = ""
        txtch3.Text = ""
        txtch4.Text = ""
        HVtest_Repetation_No = 0
        pic1.Cls
    
    End If

End Sub


Private Function vltgain(ByVal ppk As Single) As Single
Dim m, BaseGain As Single

  
Select Case cmbvsel.ListIndex

Case 0  ' 10 kv

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

Case 1 ' 50 kv
        
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

Case 2 ' 100 kv
         
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

End Select

    

End Function

Private Sub excelsave()
'Dim C, zri As Double
' ////////////////////////////////////////////////////////////////////////////////
'                 Save Data on Excel File

    
    'Set mdl1.sht = Nothing
   ' Set mdl1.sht = CreateObject("Excel.Sheet")
   ' mdl1.sht.Application.Visible = False
    
   ' mdl1.sht.Worksheets(1).Cells(1, 1).value = "Arrester Test Result"
   ' mdl1.sht.Worksheets(1).Cells(2, 2).value = "Time [mS]"
   ' mdl1.sht.Worksheets(1).Cells(2, 3).value = "Voltage [kV]"
   ' mdl1.sht.Worksheets(1).Cells(2, 4).value = "Current [mA]"
    
   ' j = 3
   ' C = 0
    
   ' For i = zri To zri + 200
    
     '   mdl1.sht.Worksheets(1).Cells(j, 2).value = C * (20 / 200) ' msec
      '  mdl1.sht.Worksheets(1).Cells(j, 3).value = voltage(i) * voltgain / (2 ^ 0.5)
       ' mdl1.sht.Worksheets(1).Cells(j, 4).value = Current(i) * currentgain
   
       ' j = j + 1
    'Next i
    
  '  mdl1.sht.Worksheets(1).Cells(2, 5).value = "Voltage [kV] (RMS)"
  '  mdl1.sht.Worksheets(1).Cells(2, 6).value = ppkv * voltgain / (2 ^ 0.5)
    
  '  mdl1.sht.Worksheets(1).Cells(3, 5).value = "I total [mA] (peak)"
  '  mdl1.sht.Worksheets(1).Cells(3, 6).value = ppkc * currentgain
  '
  '  mdl1.sht.Worksheets(1).Cells(4, 5).value = "3th Harmonic [uA]"
  '  mdl1.sht.Worksheets(1).Cells(4, 6).value = har * currentgain * 1000
    
 '   mdl1.sht.Worksheets(1).Cells(5, 5).value = "I resistive [uA]( dv/dt =0)"
 '   mdl1.sht.Worksheets(1).Cells(5, 6).value = Ires * currentgain * 1000
    
    
    'cd1.FileName = "sheet1"
    'cd1.Filter = "*.xls|*.xls"
  '  cd1.ShowSave

    'mdl1.sht.Saveas (cd1.FileName)
  '  mdl1.sht.Application.Quit
    
    'Set mdl1.sht = Nothing

    'cansave = False

End Sub

Private Sub drawcurve()
Dim cnt, cntPD, C, jp, jc, avgcnt, zr, zrp, zrc, dt, IntegrationTime As Double
Dim PDAverage(1 To 50) As Double
Dim s1, s2, n, ak, bk, w0, har, phi As Long

rtv = 0
ppeakv(1) = 0
npeakv(1) = 0
ppeakv(2) = 0
npeakv(2) = 0
ppeakv(3) = 0
npeakv(3) = 0
ppeakc(1) = 0
npeakc(1) = 0
ppeakc(2) = 0
npeakc(2) = 0
ppeakc(3) = 0
npeakc(3) = 0
pdnpk = 0
pdppk = 0
j = 0
jp = 0
prexch1 = 0

For i = 1 To 6
    dV_dt(i) = 0
Next i
ArresterCap = 0

PLoss = 0
ELoss = 0
IntegrationTime = 0

s1 = 0
s2 = 0
ak = 0
bk = 0
n = 1
phi = 3.1416 ' pi number

w0 = 100 * phi * 50 * 10 ^ -6 ' (20 [ms]/400)= 50 [us] Sampling Time is 50 us for each chaneel

On Error GoTo errhandl

'/////////////////             voltage & current curve        /////////////////////////////////
'********************************************************************************
If cmdtestok1.Enabled = True Or cmdtestok4.Enabled = True Then
    
    ' ********         find zero crossing in voltage      ****************
    
    
    maxv = 0
    maxc = 0
    pic1.Cls
    zr = 0
    zrc = 0
    
    
   ' //////////   Data Separation   //////////////
   ' even value is Current sDataBuffer(0,2,4,6, ...)
    ' Odd value is Voltage sDataBuffer(1,3,5,7,
    
    ' For each cycle: convert 800 samples/Cycle to 400 sample/Cycle for each cycle by averaging 2 sample (1200 Sample for 3 Cycle)
    ' 1200 samples for each Channel
    
    cnt = 0
    For C = 0 To UBound(sDataBufferMain) - 5 Step 4 ' each 2 samples are converted to 1 sample by averaging
    
        sDataBuffercrt(cnt) = (sDataBufferMain(C) + sDataBufferMain(C + 2)) / 2
        sDataBuffervlt(cnt) = (sDataBufferMain(C + 1) + sDataBufferMain(C + 3)) / 2
        cnt = cnt + 1
        
    Next C
       
    
    For C = 2 To 403
        If sDataBuffervlt(C - 2) < 0 And sDataBuffervlt(C + 2) > 0 Then
            zr = C
            Exit For
        End If
    Next C
    
    
    For C = 2 To 403
        If sDataBuffercrt(C - 2) < 0 And sDataBuffercrt(C + 2) > 0 Then
            zrc = C
            Exit For
        End If
    Next C
        
    pic1.Scale (0, 12)-(400, -12)
    pic1.DrawWidth = 2
    
    
'**********     calculate  average and draw curve       **************************
     C = 1
     For i = 3 To UBound(sDataBuffervlt) - 3   ' averaging 3 samples (averaging window of 3 samples run durring whole cycle)
      
     
     '**********                 Calculate dV/dt (KV/uS)   ********
        
        If ((sDataBuffervlt(i - 3) < 0 And sDataBuffervlt(i + 3) > 0) Or (sDataBuffervlt(i - 3) > 0 And sDataBuffervlt(i + 3) < 0)) And C < 7 Then 'Zero cross + or -
        
            dV_dt(C) = (sDataBuffervlt(i + 1) - sDataBuffervlt(i - 1)) / 100 ' kV/uS (dt = 100 uS)for 3 samples
            ArreyNo(C) = i
            C = C + 1
            
        End If
     
     '**********                 Calculate Average Value  for 3 samples ********
     
        p1 = (sDataBuffervlt(i - 1) + sDataBuffervlt(i) + sDataBuffervlt(i + 1)) / 3 ' + sDataBuffervlt(i) + sDataBuffervlt(i + 1) + sDataBuffervlt(i + 2) + sDataBuffervlt(i + 3)) / 7     ' Voltage Value
        p2 = (sDataBuffercrt(i - 3) + sDataBuffercrt(i) + sDataBuffercrt(i + 1)) / 3 ' + sDataBuffercrt(i) + sDataBuffercrt(i + 1) + sDataBuffercrt(i + 2) + sDataBuffercrt(i + 3)) / 7    ' Current Value
     
     
    '**********                 Calculate Max Value   ********
        If i < 400 Then
            If ppeakv(1) < p1 Then ppeakv(1) = p1
            If npeakv(1) > p1 Then npeakv(1) = p1
            If ppeakc(1) < p2 Then ppeakc(1) = p2
            If npeakc(1) > p2 Then npeakc(1) = p2
            
        ElseIf i >= 400 And i < 800 Then
            If ppeakv(2) < p1 Then ppeakv(2) = p1
            If npeakv(2) > p1 Then npeakv(2) = p1
            If ppeakc(2) < p2 Then ppeakc(2) = p2
            If npeakc(2) > p2 Then npeakc(2) = p2
            
        ElseIf i >= 800 Then
            If ppeakv(3) < p1 Then ppeakv(3) = p1
            If npeakv(3) > p1 Then npeakv(3) = p1
            If ppeakc(3) < p2 Then ppeakc(3) = p2
            If npeakc(3) > p2 Then npeakc(3) = p2
            
        End If
'  ****************   Calculate Energy loss in 3 cycles *****************
        
        ELoss = ELoss + (p1 * p2 * 50 * (10 ^ -6)) ' Eloss = voltage * Current * time (50 us)
        IntegrationTime = IntegrationTime + (50 * (10 ^ -6))
        
'  ****************   Calculate 3th harmonic of current in one period *****************
        
        If i >= 5 And i < 405 Then
            s1 = s1 + (p2 * Cos(3 * n * w0))
            s2 = s2 + (p2 * Sin(3 * n * w0))
            n = n + 1
        End If
        
        ak = (2 / 400) * s1
        bk = (2 / 400) * s2
        har = (ak ^ 2 + bk ^ 2) ^ 0.5
     
        
'*********************      draw curve in picture*******************************

         If i >= zr And i < zr + 400 Then ' draw curve
         '      Draw Voltage
         
                If j = 0 Then
                    preych1 = p1
                Else
                    pic1.Line (j - 1, preych1 * 3 * vltdiv)-(j, p1 * 3 * vltdiv), ch1color
                    preych1 = p1
                End If
                j = j + 1
                
            '    Draw Current
                If jc = 0 Then
                    preych2 = p2
                Else
                    pic1.Line (jc - 1, preych2 * 3 * crtdiv)-(jc, p2 * 3 * crtdiv), ch2color
                    preych2 = p2
                End If
                jc = jc + 1
            End If
    Next i
'***************************       CALCULATE MAXIMUM VALUE    *******************************

        ppkv = (ppeakv(1) + ppeakv(2) + ppeakv(3)) / 3
        npkv = (npeakv(1) + npeakv(2) + npeakv(3)) / 3
        ppkc = (ppeakc(1) + ppeakc(2) + ppeakc(3)) / 3
        npkc = (npeakc(1) + npeakc(2) + npeakc(3)) / 3
        
        maxv = ppeakv(1)
        If maxv < ppeakv(2) Then
            maxv = ppeakv(2)
        ElseIf maxv < ppeakv(3) Then
            maxv = ppeakv(3)
        ElseIf maxv < Abs(npeakv(1)) Then
            maxv = Abs(npeakv(1))
        ElseIf maxv < Abs(npeakv(2)) Then
            maxv = Abs(npeakv(2))
        ElseIf maxv < Abs(npeakv(3)) Then
            maxv = Abs(npeakv(3))
        End If
        
        maxc = ppeakc(1)
        If maxc < ppeakc(2) Then
            maxc = ppeakc(2)
        ElseIf maxc < ppeakc(3) Then
            maxc = ppeakc(3)
        ElseIf maxc < Abs(npeakc(1)) Then
            maxc = Abs(npeakc(1))
        ElseIf maxc < Abs(npeakc(2)) Then
            maxc = Abs(npeakc(2))
        ElseIf maxc < Abs(npeakc(3)) Then
            maxc = Abs(npeakc(3))
        End If
        
       
        If ppkv < Abs(npkv) Then  ' absolute max value
            ppkv = Abs(npkv)
        End If
        
        If ppkc < Abs(npkc) Then
            ppkc = Abs(npkc)
        End If
        
        voltgain = vltgain(ppkv)
        currentgain = crtgain(ppkc)
        
'  ****************   Calculate Power and power loss in 3 cycles *****************
        
        PLoss = (ELoss * voltgain * currentgain) / IntegrationTime ' Ploss = Eloss / Integration time (3*20 ms = 60 ms) [Watt]
        PLoss_mWatt = PLoss * 1000   ' Ploss in mili watt
        
'  ****************   Draw Current components *****************



        If chkIres.Value = 1 Then ' show Current components Ir & Ic
        
'                 Calculate Capacitance of Arrester in PF

            ArresterCap = 0
            For i = 1 To 6
                 ArresterCap = ArresterCap + (sDataBuffercrt(ArreyNo(i)) * currentgain) / (dV_dt(i) * voltgain) ' Current : [mA] &, dV : [kV] & dt : [uS] & C = [pF]
            Next i
            ArresterCap = ArresterCap / 6 '[PF]

'  ****************   Draw Current components *****************

            jc = 0
            For i = zr + 1 To zr + 400 ' 1 cycle   Ic = C * dV/dt
                Icapasitive(jc) = ArresterCap * ((sDataBuffervlt(i + 1) - sDataBuffervlt(i - 1)) / 100) ' kV/uS (dt = 100 uS)for 3 samples
                Iresistive(jc) = sDataBuffercrt(i) - Icapasitive(jc)
                
        ' ***************        Draw Current components      ************************
                If jc = 0 Then
                    preych1 = Icapasitive(jc)
                    preych2 = Iresistive(jc)
                Else
                    pic1.Line (jc - 1, preych1 * 3 * crtdiv)-(jc, Icapasitive(jc) * 3 * crtdiv), vbGreen ' vbGreen
                    pic1.Line (jc - 1, preych2 * 3 * crtdiv)-(jc, Iresistive(jc) * 3 * crtdiv), vbMagenta
                    
                    preych1 = Icapasitive(jc)
                    preych2 = Iresistive(jc)
                End If
                jc = jc + 1
            Next i
            
        End If
    
   ' End If
'********************** DISPLAY CURRECT VALUE FOR VOLTAGE AND CURRENT  *************
       
        txtch1.Text = Format(ppkv * voltgain / (2 ^ 0.5), "##0.00") ' RMS Value   (KV)
        txtch2.Text = Format(ppkc * currentgain, "#0.00")          ' peak Value  (mA)
        txtch3.Text = Format(har * currentgain * 1000, "###00")   ' (uA) 3th harmonic of current
        txtch4.Text = Format(PLoss_mWatt, "##000")    ' (mWatt) Powel Loss
         
'**************************8      RESULT OF TEST 1       **************
        If cmdtestok1.Enabled = True And Val(txtch2.Text) < (1.1 * iref1) And Val(txtch2.Text) > (0.95 * iref1) Then

            testapp1 = True
            If txturef1.Text <> "" And txturef1.Text <> " " Then
                prevoltage = Val(txturef1.Text)
            Else
                prevoltage = Val(txtch1.Text)
            End If
            txturef1.Text = Format((Val(txtch1.Text) + prevoltage) / 2, "##0.00")

'**************************8      RESULT OF TEST 4       **************
        ElseIf cmdtestok4.Enabled = True And Val(txtch1.Text) < (1.04 * vref2) And Val(txtch1.Text) > (0.98 * vref2) Then

            testapp4 = True
            If txtic4.Text <> "" And txtic4.Text <> " " Then
                precurrent = Val(txtic4.Text)
                
            Else
                precurrent = Val(txtch2.Text)
                harm3 = Val(txtch3.Text)
            End If
            
            txtic4.Text = Format((Val(txtch2.Text) + precurrent) / 2, "#0.00#")
            harm3 = (harm3 + Val(txtch3.Text)) / 2
        End If
        
        
 '///////////////////////////////////////////////////////////////////////////////////
 
'//////////////////////     VOLTAGE   & PD CURVE       /////////////////////////////////

'**************************************************************************************

'////////////////////////////////////////////////////////////////////////////////////

 ElseIf cmdtestok2.Enabled = True Or cmdtestok3.Enabled = True Then
 
' ********         find zero crossing in voltage      ****************

'  voltage = sDataBuffervlt(0,1,2 ...)
'  PD      = sDataBufferpd(0,1,2 ...)
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    pic1.Cls
    maxv = 0
    maxp = 0
    zr = 0
    zrp = 0
    
    '     ////   Data Separation  ///////////////
    cnt = 0
    
    
    '  400 SAmple/ Cycle for Voltage and 4800 Sample/Cycle for PD (12 * 400 = 4800)
    '  1200 SAmples for Voltage and 14400 SAmple for PD (3 Cycle)
    
   ' jc = UBound(sDataBuffervlt)
    For C = 0 To (UBound(sDataBufferMain2) - 1) Step 2
        sDataBuffervlt(cnt) = (sDataBufferMain2(C) + sDataBufferMain2(C + 1)) / 2
        cnt = cnt + 1
    Next C
       
    
    For C = 2 To 402
        If sDataBuffervlt(C - 1) < 0 And sDataBuffervlt(C + 1) > 0 Then
            zr = C
            Exit For
        End If
    Next C
    
   
    '**********     calculate  average and draw curve       **************************
    
    '               Voltage Calculation  and Drawing
    pic1.Scale (0, 12)-(4800, -12)
    
    For i = 1 To (1200 - 1)         'Read Peak Value of Voltage
           
        p1 = (sDataBuffervlt(i) + sDataBuffervlt(i - 1) + sDataBuffervlt(i + 1)) / 3     ' Voltage Value
        If i < 400 Then
            If ppeakv(1) < p1 Then ppeakv(1) = p1
            If npeakv(1) > p1 Then npeakv(1) = p1
                    
        ElseIf i >= 400 And i < 800 Then
            If ppeakv(2) < p1 Then ppeakv(2) = p1
            If npeakv(2) > p1 Then npeakv(2) = p1
                   
        ElseIf i >= 800 Then
            If ppeakv(3) < p1 Then ppeakv(3) = p1
            If npeakv(3) > p1 Then npeakv(3) = p1
                    
        End If
            
           
        If i >= zr And i < zr + 400 Then          ' draw Voltage
            If j = 0 Then
                preych1 = p1
            Else
                pic1.DrawWidth = 2
                pic1.Line (j - 12, preych1 * 3 * vltdiv)-(j, p1 * 3 * vltdiv), ch1color
                preych1 = p1
    
            End If
            j = j + 12

        End If
    Next i
        
' *******************************   PD    ***************************

' ............  Find PD Offset value (DC VAlue)
    maxp = 0
    For C = 1 To 50
        cntPD = 0
        PDAverage(C) = 0
        For i = (96 * (C - 1)) To (96 * C) - 1  ' last calculated arrey of PD is (96* 50)-1  (1 cycle)
            PDAverage(C) = PDAverage(C) + sDataBufferpd(i)
            cntPD = cntPD + 1
        Next i
        PDAverage(C) = PDAverage(C) / cntPD
        
         ' Find Real Value of PD after removing DC offset
          
        For i = (96 * (C - 1)) To (96 * C) - 1
        
            If maxp < sDataBufferpd(i) Then  ' find MAX value for I/O card gain
                maxp = sDataBufferpd(i)
            End If
            
            sDataBufferpd(i) = sDataBufferpd(i) - PDAverage(C)
        Next i
        
    Next C
    
    
    For i = 0 To 4800 - 1 ' (120 * 120) - 1  ' last arrey of PD is (120* 120)-1
      
        p2 = sDataBufferpd(i)
        
        If pdppk < p2 Then pdppk = p2
        If pdnpk > p2 Then pdnpk = p2
        
        'If i >= 0 And i < 4800 Then ' Draw 1 Cycle of PD
        
        pic1.DrawWidth = 1
        If jp = 0 Then
            preych2 = p2
        Else
            pic1.DrawWidth = 1
            'pic1.Line ((jp - 1) / 3, preych2 * 3 * pddiv)-(jp / 3, p2 * 3 * pddiv), ch2color
            pic1.Line ((jp - 1), preych2 * 3 * pddiv)-(jp, p2 * 3 * pddiv), ch2color
            preych2 = p2
        End If
        jp = jp + 1
       'End If
       
    Next i
    
 '*************************** CALCULATE MAXIMUME VALUE    *****************************
    If pdppk < Abs(pdnpk) Then
        pdppk = Abs(pdnpk)
    End If
    
    
    ppkv = (ppeakv(1) + ppeakv(2) + ppeakv(3)) / 3
    npkv = (npeakv(1) + npeakv(2) + npeakv(3)) / 3
    
    If ppkv < Abs(npkv) Then
        ppkv = Abs(npkv)
    End If
    
    voltgain = vltgain(ppkv)
    
    maxv = ppeakv(1)
    If maxv < ppeakv(2) Then
        maxv = ppeakv(2)
    ElseIf maxv < ppeakv(3) Then
        maxv = ppeakv(3)
    ElseIf maxv < Abs(npeakv(1)) Then
        maxv = Abs(npeakv(1))
    ElseIf maxv < Abs(npeakv(2)) Then
        maxv = Abs(npeakv(2))
    ElseIf maxv < Abs(npeakv(3)) Then
        maxv = Abs(npeakv(3))
    End If
    
    
'**************************************************************
   
    txtch1.Text = Format(ppkv * voltgain / (2 ^ 0.5), "##0.00") ' RMS Value
    'txtch2.Text = Format(pdppk * SelectedPDGain * (1 + Rnd(0.5)), "0.0#") ' fake formula
    txtch2.Text = Format(pdppk * SelectedPDGain, "0.0#") '  real formula
    txtch3.Text = "No Signal"
    txtch4.Text = "No Signal"
    
    '**************************8      RESULT OF TEST 2       **************
    If cmdtestok2.Enabled = True And Val(txtch1.Text) < (1.04 * vpd1) And Val(txtch1.Text) > (0.96 * vpd1) Then
        testapp2 = True
        If txtpd2.Text <> "" And txtpd2.Text <> " " Then
            prepd = Val(txtpd2.Text)
        Else
            prepd = Val(txtch2.Text)
        End If
        txtpd2.Text = Format(Val(txtch2.Text), "0.00") ' Format(((Val(txtch2.Text) + prepd) / 2), "0.00")
    '**************************8      RESULT OF TEST 3       **************
    ElseIf cmdtestok3.Enabled = True And Val(txtch1.Text) < (1.04 * vpd2) And Val(txtch1.Text) > (0.96 * vpd2) Then
        testapp3 = True
        If txtpd3.Text <> "" And txtpd3.Text <> " " Then
            prepd = Val(txtpd3.Text)
        Else
            prepd = Val(txtch2.Text)
        End If
        txtpd3.Text = Format(Val(txtch2.Text), "0.00") 'Format(((Val(txtch2.Text) + prepd) / 2), "0.00")
    End If
 
'////////////////////////////////////////////////////////////////////////////
End If
Exit Sub

errhandl:
    pic1.Cls
    maxv = 2.5
    maxc = 1.5
    maxp = 0.5
End Sub
Private Sub selectgain()

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'                  Hardware Configuration

'  Channel 0 = Current   Channel 2 = Voltage     Channel 4 SC  Channel: 6 PD (Channel: 8 to 14 PD or open via PIN Switch)

' VIBuffer.channel(#).ValueRange :  12  =  +- 0.625 V
'                                    4  =  +- 1.25 V
'                                    3  =  +- 2.5 V
'                                    2  =  +- 5 V
'                                    1  =  +- 10 V


If cmdtestok1.Enabled = True Or cmdtestok4.Enabled = True Then 'voltage & current
       
        If reset = True Then ' Get ready to SC Condition
            VIBuffer.ScanChannel.ChannelStart = 4
            VIBuffer.ScanChannel.ChannelCount = 1
            VIBuffer.ScanChannel.Samples = 200
            VIBuffer.ConvertClock.Rate = 1000
            VIBuffer.ScanChannel.IntervalCount = 1
            VIBuffer.Channel(4).ValueRange = 12  ' gain for SC  = +/- 0.625
            
            
            
        Else ' Get ready to Read Voltage And Current signals simultaneously
        
            VIBuffer.ScanChannel.ChannelStart = 0   ' Current Signal
            VIBuffer.ScanChannel.ChannelCount = 2   ' number of channel ( 2 Channels, synchronous Sampling of Voltage & current)
            VIBuffer.ScanChannel.Samples = 2405
            VIBuffer.ConvertClock.Rate = 40000
            VIBuffer.ScanChannel.IntervalCount = 1
            
            '                           40 kHz  for each channel,  800 samp / cycle for each chanel
            '                           The Samples/ Cycle reduce to 400 samples in Drawcurve Sub Routine
            '                           2400 sample for each channel (3 cycles for each channel (60 mS))
           
            
            
           ' timstr = Timer
            
            If maxc <= 0.62 Then
                VIBuffer.Channel(0).ValueRange = 12  ' gain for current = +/- 0.625
                
            
            ElseIf maxc > 0.62 And maxc <= 1.2 Then
               VIBuffer.Channel(0).ValueRange = 4 ' gain for current = +/- 1.25 v
              
                
            ElseIf maxc > 1.2 And maxc <= 2.4 Then
                VIBuffer.Channel(0).ValueRange = 3 ' gain for current = +/- 2.5 v
               
               
            ElseIf maxc > 2.4 And maxc <= 4.9 Then
              VIBuffer.Channel(0).ValueRange = 2 ' gain for current = +/- 5 v
            
            
            ElseIf maxc > 4.9 Then
               
                VIBuffer.Channel(0).ValueRange = 1 ' gain for current = +/- 10 v
            
            End If
        
        
           ' dtime = (Timer - timstr) * 1000
           
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
            
        End If
       
ElseIf cmdtestok2.Enabled = True Or cmdtestok3.Enabled = True Then 'Voltage & SC & PD
     
    If reset = True Then ' Get ready to SC Condition
        
           VIBuffer.ScanChannel.ChannelStart = 4
            VIBuffer.ScanChannel.ChannelCount = 1
            VIBuffer.ScanChannel.Samples = 200
            VIBuffer.ConvertClock.Rate = 1000
            VIBuffer.ScanChannel.IntervalCount = 1
            VIBuffer.Channel(4).ValueRange = 12  ' gain for SC  = +/- 0.625
            
    Else ' Get ready to Read Voltage Or PD signals Asynchronously
            
        If state = 1 Then ' Get ready to read Voltage Signal
            VIBuffer.ScanChannel.ChannelStart = 2   ' Voltage Signal
            VIBuffer.ScanChannel.ChannelCount = 1   ' number of channel ( 1 Channels, Only Sampling of Voltage )
            VIBuffer.ScanChannel.Samples = 2410
            VIBuffer.ConvertClock.Rate = 40000
            VIBuffer.ScanChannel.IntervalCount = 1
            
            '                           40 kHz  for each channel,  800 samp / cycle for each chanel
            '                           The Samples/ Cycle reduce to 400 samples in Drawcurve Sub Routine
            '                           2400 sample for each channel (3 cycles for each channel (60 mS))
            
                 
                 'timstr = Timer
                    
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
                 
        ElseIf state = 0 Then ' Get ready to read PD Signal
                
                VIBuffer.ScanChannel.ChannelStart = 6   ' PD Signal
                VIBuffer.ScanChannel.ChannelCount = 1   ' number of channel ( 1 Channels, Only Sampling of PD )
                VIBuffer.ScanChannel.Samples = 14410 ' Number of Samples
                VIBuffer.ConvertClock.Rate = 240000  ' frequency in each chanel
                VIBuffer.ScanChannel.IntervalCount = 1
                '                               Sample Rtae 240 kHz (4800 Sam/Cycle for PD)
                '                               3 cycle = 14400 Samples for PD
                
                
                ' timstr = Timer
                   
               If maxp <= 0.62 Then
                      VIBuffer.Channel(6).ValueRange = 12  ' gain for PD = +/- 0.625
                  
                  ElseIf maxp > 0.62 And maxp <= 1.2 Then
                       VIBuffer.Channel(6).ValueRange = 4 ' gain for PD = +/- 1.25 v
                  
                  ElseIf maxp > 1.2 And maxp <= 2.4 Then
                      VIBuffer.Channel(6).ValueRange = 3 ' gain for PD = +/- 2.5 v
                  
                  ElseIf maxp > 2.4 And maxp <= 4.9 Then
                       VIBuffer.Channel(6).ValueRange = 2 ' gain for PD = +/- 5
                  
                  ElseIf maxp > 4.9 Then
                       VIBuffer.Channel(6).ValueRange = 1 ' gain for PD = +/- 10 v
                  
                  End If
               
        End If
            
    End If
        
End If

End Sub

Private Sub starttest()

'/*************************************

'/&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

alwstrttst = False
maxv = 2.5
maxc = 1.25
maxp = 0.5
draw = False
state = 1 ' Voltage signal has priority

tok1 = False
tok2 = False
tok3 = False
tok4 = False
endtest = False

reset = True ' start reading data with SC Condition in each test
stptst = False


'&&&&&&&&&&&&&&&&&                  DAQ CARD Setting            &&&&&&&&&&&&&&&&&&&&&&&&&&
'              Start Scan  Command

'gwFifoSize = 512
'gwEvtFlag = 1 '.......................... event enable(1)
'gwFifoEnable = 1 '........................Fifo  enable(1)
'gwFifoSize = 512 '........................Fifo size= 1024 /2 ==> 512
'ptEnableEvent.EventType = ADS_EVT_INTERRUPT Or ADS_EVT_BUFCHANGE Or ADS_EVT_OVERRUN Or ADS_EVT_TERMINATED
'ptEnableEvent.Enabled = 1
'ptEnableEvent.Count = 1 'with FIFO Function
'ptAiStart.NumChans = 1 ' number of channel
'ptAiStart.TrigSrc = 0     'Internal trig mode
'ptAiStart.cyclic = 0  ' without cycling
'ptAiStart.IntrCount = 1
'                     Check Command
'ptAICheck.stopped = DRV_GetAddress(lStopped)
'ptAICheck.retrieved = DRV_GetAddress(iRetrieved)
'ptAICheck.Overrun = DRV_GetAddress(lOverrun)
'ptAICheck.HalfReady = DRV_GetAddress(lHalfReady)
'ptAICheck.ActiveBuf = DRV_GetAddress(lActiveBuf)
'                 Transfer   Command
'ptFAIData.Overrun = DRV_GetAddress(lOverrun)
'ptFAIData.start = 0                    'Start from buffer begging
'ptFAIData.DataType = 1 '  voltage value   //want floating data
'ptFAIData.ActiveBuf = 0 'Single buffer

'*************************************************************************************
If datatestprmt.Recordset.Fields("Urefmax") < 0.3 Then
    
    cmbosc1.Text = "100 V"
    
ElseIf datatestprmt.Recordset.Fields("Urefmax") >= 0.3 And datatestprmt.Recordset.Fields("Urefmax") < 3 Then
   
    cmbosc1.Text = "1 KV"
    
ElseIf datatestprmt.Recordset.Fields("Urefmax") >= 3 And datatestprmt.Recordset.Fields("Urefmax") < 20 Then
    
    cmbosc1.Text = "8 KV"
    

ElseIf datatestprmt.Recordset.Fields("Urefmax") >= 20 And datatestprmt.Recordset.Fields("Urefmax") < 22 Then
    
    cmbosc1.Text = "8 KV"

ElseIf datatestprmt.Recordset.Fields("Urefmax") >= 22 And datatestprmt.Recordset.Fields("Urefmax") < 28 Then
   
    cmbosc1.Text = "10 KV"
    
ElseIf datatestprmt.Recordset.Fields("Urefmax") >= 28 And datatestprmt.Recordset.Fields("Urefmax") < 45 Then
    
    cmbosc1.Text = "20 KV"
    
 
ElseIf datatestprmt.Recordset.Fields("Urefmax") >= 45 And datatestprmt.Recordset.Fields("Urefmax") < 84 Then
    
    cmbosc1.Text = "30 KV"
    
ElseIf datatestprmt.Recordset.Fields("Urefmax") >= 84 And datatestprmt.Recordset.Fields("Urefmax") < 110 Then
    
    cmbosc1.Text = "40 KV"
    
ElseIf datatestprmt.Recordset.Fields("Urefmax") >= 110 And datatestprmt.Recordset.Fields("Urefmax") < 140 Then
   
    cmbosc1.Text = "50 KV"
    
ElseIf datatestprmt.Recordset.Fields("Urefmax") >= 140 And datatestprmt.Recordset.Fields("Urefmax") < 170 Then
  
    cmbosc1.Text = "60 KV"
    
ElseIf datatestprmt.Recordset.Fields("Urefmax") >= 170 And datatestprmt.Recordset.Fields("Urefmax") < 190 Then
   
    cmbosc1.Text = "70 KV"
ElseIf datatestprmt.Recordset.Fields("Urefmax") >= 190 Then
    
    cmbosc1.Text = "80 KV"
    
End If
'////////////////////////////////////////////////////////////////////////////
If datatestprmt.Recordset.Fields("Iref") < 0.8 Then
    cmbosc2.Text = "0.4 mA"
    
ElseIf datatestprmt.Recordset.Fields("Iref") >= 0.8 And datatestprmt.Recordset.Fields("Iref") < 1.6 Then
    cmbosc2.Text = "0.8 mA"
    
ElseIf datatestprmt.Recordset.Fields("Iref") >= 1.6 And datatestprmt.Recordset.Fields("Iref") < 3.2 Then
    cmbosc2.Text = "1 mA"
   
ElseIf datatestprmt.Recordset.Fields("Iref") >= 3.2 And datatestprmt.Recordset.Fields("Iref") < 4 Then
    cmbosc2.Text = "2 mA"
   
ElseIf datatestprmt.Recordset.Fields("Iref") >= 4 Then
    cmbosc2.Text = "3 mA"
   
End If

plottime.Enabled = True
plottime.Interval = 5

End Sub

Private Sub CmbPDSel_Change()

Select Case CmbPDSel.ListIndex

Case 0
    SelectedPDGain = gainpd(0)
Case 1
     SelectedPDGain = gainpd(1)
Case 2
     SelectedPDGain = gainpd(2)
End Select


End Sub

Private Sub CmbPDSel_Click()

Select Case CmbPDSel.ListIndex
Case 0
    SelectedPDGain = gainpd(0)
Case 1
     SelectedPDGain = gainpd(1)
Case 2
     SelectedPDGain = gainpd(2)
End Select


End Sub

Private Sub CmdCalibration_Click()
Me.Hide
FRMCAL.Show vbModal
End Sub

Private Sub cmdclr1_Click()
cd1.ShowColor
ch1color = cd1.Color
End Sub

Private Sub cmbosc1_Change()
Select Case cmbvsel.ListIndex

Case 0 ' 10 KV

    Select Case cmbosc1.ListIndex
    
    Case 0
        vltdiv = gainvlt(1) / 0.1
        Call mdl1.leftlable(0.1 * 4)
    Case 1
        vltdiv = gainvlt(1) / 1
        Call mdl1.leftlable(1 * 4)
    Case 2
       vltdiv = gainvlt(1) / 2
        Call mdl1.leftlable(2 * 4)
    Case 3
        vltdiv = gainvlt(1) / 5
        Call mdl1.leftlable(5 * 4)
    Case 4
        vltdiv = gainvlt(1) / 8
        Call mdl1.leftlable(8 * 4)
    Case 5
        vltdiv = gainvlt(1) / 10
        Call mdl1.leftlable(10 * 4)
    Case 6
        vltdiv = gainvlt(1) / 20
        Call mdl1.leftlable(20 * 4)
    Case 7
        vltdiv = gainvlt(1) / 30
        Call mdl1.leftlable(30 * 4)
    Case 8
        vltdiv = gainvlt(1) / 40
        Call mdl1.leftlable(40 * 4)
    Case 9
        vltdiv = gainvlt(1) / 50
        Call mdl1.leftlable(50 * 4)
    Case 10
        vltdiv = gainvlt(1) / 60
        Call mdl1.leftlable(60 * 4)
    Case 11
        vltdiv = gainvlt(1) / 70
        Call mdl1.leftlable(70 * 4)
    Case 12
        vltdiv = gainvlt(1) / 80
        Call mdl1.leftlable(80 * 4)
    End Select

Case 1 ' 50 kv

    Select Case cmbosc1.ListIndex
    
    Case 0
        vltdiv = gainvlt(4) / 0.1
        Call mdl1.leftlable(0.1 * 4)
    Case 1
        vltdiv = gainvlt(4) / 1
        Call mdl1.leftlable(1 * 4)
    Case 2
       vltdiv = gainvlt(4) / 2
        Call mdl1.leftlable(2 * 4)
    Case 3
        vltdiv = gainvlt(4) / 5
        Call mdl1.leftlable(5 * 4)
    Case 4
        vltdiv = gainvlt(4) / 8
        Call mdl1.leftlable(8 * 4)
    Case 5
        vltdiv = gainvlt(4) / 10
        Call mdl1.leftlable(10 * 4)
    Case 6
        vltdiv = gainvlt(4) / 20
        Call mdl1.leftlable(20 * 4)
    Case 7
        vltdiv = gainvlt(4) / 30
        Call mdl1.leftlable(30 * 4)
    Case 8
        vltdiv = gainvlt(4) / 40
        Call mdl1.leftlable(40 * 4)
    Case 9
        vltdiv = gainvlt(4) / 50
        Call mdl1.leftlable(50 * 4)
    Case 10
        vltdiv = gainvlt(4) / 60
        Call mdl1.leftlable(60 * 4)
    Case 11
        vltdiv = gainvlt(4) / 70
        Call mdl1.leftlable(70 * 4)
    Case 12
        vltdiv = gainvlt(4) / 80
        Call mdl1.leftlable(80 * 4)
    End Select
    
    
Case 2  ' 100 kv

    Select Case cmbosc1.ListIndex
    
    Case 0
        vltdiv = gainvlt(7) / 0.1
        Call mdl1.leftlable(0.1 * 4)
    Case 1
        vltdiv = gainvlt(7) / 1
        Call mdl1.leftlable(1 * 4)
    Case 2
       vltdiv = gainvlt(7) / 2
        Call mdl1.leftlable(2 * 4)
    Case 3
        vltdiv = gainvlt(7) / 5
        Call mdl1.leftlable(5 * 4)
    Case 4
        vltdiv = gainvlt(7) / 8
        Call mdl1.leftlable(8 * 4)
    Case 5
        vltdiv = gainvlt(7) / 10
        Call mdl1.leftlable(10 * 4)
    Case 6
        vltdiv = gainvlt(7) / 20
        Call mdl1.leftlable(20 * 4)
    Case 7
        vltdiv = gainvlt(7) / 30
        Call mdl1.leftlable(30 * 4)
    Case 8
        vltdiv = gainvlt(7) / 40
        Call mdl1.leftlable(40 * 4)
    Case 9
        vltdiv = gainvlt(7) / 50
        Call mdl1.leftlable(50 * 4)
    Case 10
        vltdiv = gainvlt(7) / 60
        Call mdl1.leftlable(60 * 4)
    Case 11
        vltdiv = gainvlt(7) / 70
        Call mdl1.leftlable(70 * 4)
    Case 12
        vltdiv = gainvlt(7) / 80
        Call mdl1.leftlable(80 * 4)
    End Select
End Select

End Sub

Private Sub cmbosc1_Click()
Select Case cmbvsel.ListIndex

Case 0

    Select Case cmbosc1.ListIndex
    
    Case 0
        vltdiv = gainvlt(1) / 0.1
        Call mdl1.leftlable(0.1 * 4)
    Case 1
        vltdiv = gainvlt(1) / 1
        Call mdl1.leftlable(1 * 4)
    Case 2
       vltdiv = gainvlt(1) / 2
        Call mdl1.leftlable(2 * 4)
    Case 3
        vltdiv = gainvlt(1) / 5
        Call mdl1.leftlable(5 * 4)
    Case 4
        vltdiv = gainvlt(1) / 8
        Call mdl1.leftlable(8 * 4)
    Case 5
        vltdiv = gainvlt(1) / 10
        Call mdl1.leftlable(10 * 4)
    Case 6
        vltdiv = gainvlt(1) / 20
        Call mdl1.leftlable(20 * 4)
    Case 7
        vltdiv = gainvlt(1) / 30
        Call mdl1.leftlable(30 * 4)
    Case 8
        vltdiv = gainvlt(1) / 40
        Call mdl1.leftlable(40 * 4)
    Case 9
        vltdiv = gainvlt(1) / 50
        Call mdl1.leftlable(50 * 4)
    Case 10
        vltdiv = gainvlt(1) / 60
        Call mdl1.leftlable(60 * 4)
    Case 11
        vltdiv = gainvlt(1) / 70
        Call mdl1.leftlable(70 * 4)
    Case 12
        vltdiv = gainvlt(1) / 80
        Call mdl1.leftlable(80 * 4)
    End Select

Case 1
    Select Case cmbosc1.ListIndex
    
    Case 0
        vltdiv = gainvlt(4) / 0.1
        Call mdl1.leftlable(0.1 * 4)
    Case 1
        vltdiv = gainvlt(4) / 1
        Call mdl1.leftlable(1 * 4)
    Case 2
       vltdiv = gainvlt(4) / 2
        Call mdl1.leftlable(2 * 4)
    Case 3
        vltdiv = gainvlt(4) / 5
        Call mdl1.leftlable(5 * 4)
    Case 4
        vltdiv = gainvlt(4) / 8
        Call mdl1.leftlable(8 * 4)
    Case 5
        vltdiv = gainvlt(4) / 10
        Call mdl1.leftlable(10 * 4)
    Case 6
        vltdiv = gainvlt(4) / 20
        Call mdl1.leftlable(20 * 4)
    Case 7
        vltdiv = gainvlt(4) / 30
        Call mdl1.leftlable(30 * 4)
    Case 8
        vltdiv = gainvlt(4) / 40
        Call mdl1.leftlable(40 * 4)
    Case 9
        vltdiv = gainvlt(4) / 50
        Call mdl1.leftlable(50 * 4)
    Case 10
        vltdiv = gainvlt(4) / 60
        Call mdl1.leftlable(60 * 4)
    Case 11
        vltdiv = gainvlt(4) / 70
        Call mdl1.leftlable(70 * 4)
    Case 12
        vltdiv = gainvlt(4) / 80
        Call mdl1.leftlable(80 * 4)
    End Select
Case 2
    Select Case cmbosc1.ListIndex
    
    Case 0
        vltdiv = gainvlt(7) / 0.1
        Call mdl1.leftlable(0.1 * 4)
    Case 1
        vltdiv = gainvlt(7) / 1
        Call mdl1.leftlable(1 * 4)
    Case 2
       vltdiv = gainvlt(7) / 2
        Call mdl1.leftlable(2 * 4)
    Case 3
        vltdiv = gainvlt(7) / 5
        Call mdl1.leftlable(5 * 4)
    Case 4
        vltdiv = gainvlt(7) / 8
        Call mdl1.leftlable(8 * 4)
    Case 5
        vltdiv = gainvlt(7) / 10
        Call mdl1.leftlable(10 * 4)
    Case 6
        vltdiv = gainvlt(7) / 20
        Call mdl1.leftlable(20 * 4)
    Case 7
        vltdiv = gainvlt(7) / 30
        Call mdl1.leftlable(30 * 4)
    Case 8
        vltdiv = gainvlt(7) / 40
        Call mdl1.leftlable(40 * 4)
    Case 9
        vltdiv = gainvlt(7) / 50
        Call mdl1.leftlable(50 * 4)
    Case 10
        vltdiv = gainvlt(7) / 60
        Call mdl1.leftlable(60 * 4)
    Case 11
        vltdiv = gainvlt(7) / 70
        Call mdl1.leftlable(70 * 4)
    Case 12
        vltdiv = gainvlt(7) / 80
        Call mdl1.leftlable(80 * 4)
    End Select
End Select

End Sub


Private Sub cmbosc2_Change()
Select Case cmbosc2.ListIndex
Case 0
    crtdiv = gaincrt(2) / 0.1
    Call mdl1.rightlable(0.1 * 4)
Case 1
    crtdiv = gaincrt(3) / 0.2
    Call mdl1.rightlable(0.2 * 4)
Case 2
    crtdiv = gaincrt(5) / 0.4
    Call mdl1.rightlable(0.4 * 4)
Case 3
    crtdiv = gaincrt(7) / 0.8
    Call mdl1.rightlable(0.8 * 4)
Case 4
    crtdiv = gaincrt(8) / 1
    Call mdl1.rightlable(1 * 4)
Case 5
    crtdiv = gaincrt(11) / 2
    Call mdl1.rightlable(2 * 4)
Case 6
    crtdiv = gaincrt(14) / 3
    Call mdl1.rightlable(3 * 4)
End Select
lblcurpd.Caption = "Current [mA]"
End Sub


Private Sub cmbosc2_Click()
Select Case cmbosc2.ListIndex
Case 0
    crtdiv = gaincrt(2) / 0.1
    Call mdl1.rightlable(0.1 * 4)
Case 1
    crtdiv = gaincrt(3) / 0.2
    Call mdl1.rightlable(0.2 * 4)
Case 2
    crtdiv = gaincrt(5) / 0.4
    Call mdl1.rightlable(0.4 * 4)
Case 3
    crtdiv = gaincrt(7) / 0.8
    Call mdl1.rightlable(0.8 * 4)
Case 4
    crtdiv = gaincrt(8) / 1
    Call mdl1.rightlable(1 * 4)
Case 5
    crtdiv = gaincrt(11) / 2
    Call mdl1.rightlable(2 * 4)
Case 6
    crtdiv = gaincrt(14) / 3
    Call mdl1.rightlable(3 * 4)
End Select
lblcurpd.Caption = "Current [mA]"
End Sub


Private Sub cmbosc2p_Change()
Select Case cmbosc2p.ListIndex

Case 0
    pddiv = SelectedPDGain / 0.2
    Call mdl1.rightlable(0.2 * 4)
Case 1
     pddiv = SelectedPDGain / 0.5
     Call mdl1.rightlable(0.5 * 4)
Case 2
    pddiv = SelectedPDGain / 1
    Call mdl1.rightlable(1 * 4)
Case 3
    pddiv = SelectedPDGain / 2
    Call mdl1.rightlable(2 * 4)
Case 4
    pddiv = SelectedPDGain / 4
    Call mdl1.rightlable(4 * 4)
End Select

End Sub

Private Sub cmbosc2p_Click()

Select Case cmbosc2p.ListIndex
Case 0
    pddiv = SelectedPDGain / 0.2
    Call mdl1.rightlable(0.2 * 4)
Case 1
     pddiv = SelectedPDGain / 0.5
     Call mdl1.rightlable(0.5 * 4)
Case 2
    pddiv = SelectedPDGain / 1
    Call mdl1.rightlable(1 * 4)
Case 3
    pddiv = SelectedPDGain / 2
    Call mdl1.rightlable(2 * 4)
Case 4
    pddiv = SelectedPDGain / 4
    Call mdl1.rightlable(4 * 4)
End Select

End Sub

Private Sub cmbtype_Change()
temselectedarr = cmbtype.List(cmbtype.ListIndex)

With datatest

    If DataTabletype(cmbtype.ListIndex) = "SUBST" Then  '  Station Class Arrester Coding
        .RecordSource = "SUBST"
        mdl1.ArrTableName = "SUBST"
        
    ElseIf DataTabletype(cmbtype.ListIndex) = "DIST" Then '  dISTRIBUTION Class Arrester Coding
       .RecordSource = "DIST"
        mdl1.ArrTableName = "DIST"
    ElseIf DataTabletype(cmbtype.ListIndex) = "CNT" Then '  Counter Coding
       .RecordSource = "DIST"
        mdl1.ArrTableName = "CNT"
        
    End If
    
    .Refresh
    
End With

End Sub

Private Sub cmbtype_Click()
temselectedarr = cmbtype.List(cmbtype.ListIndex)

With datatest

    If DataTabletype(cmbtype.ListIndex) = "SUBST" Then  '  Station Class Arrester Coding
        .RecordSource = "SUBST"
        mdl1.ArrTableName = "SUBST"
        
    ElseIf DataTabletype(cmbtype.ListIndex) = "DIST" Then '  dISTRIBUTION Class Arrester Coding
        .RecordSource = "DIST"
        mdl1.ArrTableName = "DIST"
        
    ElseIf DataTabletype(cmbtype.ListIndex) = "CNT" Then '  Counter Coding
       .RecordSource = "DIST"
        mdl1.ArrTableName = "CNT"
        
    End If
    
    .Refresh
    
End With

End Sub


Private Sub cmbvsel_Change()
Select Case cmbvsel.ListIndex

Case 0 ' 10 kV

    Select Case cmbosc1.ListIndex
    
    Case 0
        vltdiv = gainvlt(1) / 0.1
        Call mdl1.leftlable(0.1 * 4)
    Case 1
        vltdiv = gainvlt(1) / 1
        Call mdl1.leftlable(1 * 4)
    Case 2
       vltdiv = gainvlt(1) / 2
        Call mdl1.leftlable(2 * 4)
    Case 3
        vltdiv = gainvlt(1) / 5
        Call mdl1.leftlable(5 * 4)
    Case 4
        vltdiv = gainvlt(1) / 8
        Call mdl1.leftlable(8 * 4)
    Case 5
        vltdiv = gainvlt(1) / 10
        Call mdl1.leftlable(10 * 4)
    Case 6
        vltdiv = gainvlt(1) / 20
        Call mdl1.leftlable(20 * 4)
    Case 7
        vltdiv = gainvlt(1) / 30
        Call mdl1.leftlable(30 * 4)
    Case 8
        vltdiv = gainvlt(1) / 40
        Call mdl1.leftlable(40 * 4)
    Case 9
        vltdiv = gainvlt(1) / 50
        Call mdl1.leftlable(50 * 4)
    Case 10
        vltdiv = gainvlt(1) / 60
        Call mdl1.leftlable(60 * 4)
    Case 11
        vltdiv = gainvlt(1) / 70
        Call mdl1.leftlable(70 * 4)
    Case 12
        vltdiv = gainvlt(1) / 80
        Call mdl1.leftlable(80 * 4)
    End Select

Case 1 ' 50 kV
    Select Case cmbosc1.ListIndex
    
    Case 0
        vltdiv = gainvlt(4) / 0.1
        Call mdl1.leftlable(0.1 * 4)
    Case 1
        vltdiv = gainvlt(4) / 1
        Call mdl1.leftlable(1 * 4)
    Case 2
       vltdiv = gainvlt(4) / 2
        Call mdl1.leftlable(2 * 4)
    Case 3
        vltdiv = gainvlt(4) / 5
        Call mdl1.leftlable(5 * 4)
    Case 4
        vltdiv = gainvlt(4) / 8
        Call mdl1.leftlable(8 * 4)
    Case 5
        vltdiv = gainvlt(4) / 10
        Call mdl1.leftlable(10 * 4)
    Case 6
        vltdiv = gainvlt(4) / 20
        Call mdl1.leftlable(20 * 4)
    Case 7
        vltdiv = gainvlt(4) / 30
        Call mdl1.leftlable(30 * 4)
    Case 8
        vltdiv = gainvlt(4) / 40
        Call mdl1.leftlable(40 * 4)
    Case 9
        vltdiv = gainvlt(4) / 50
        Call mdl1.leftlable(50 * 4)
    Case 10
        vltdiv = gainvlt(4) / 60
        Call mdl1.leftlable(60 * 4)
    Case 11
        vltdiv = gainvlt(4) / 70
        Call mdl1.leftlable(70 * 4)
    Case 12
        vltdiv = gainvlt(4) / 80
        Call mdl1.leftlable(80 * 4)
    End Select
    
Case 2 ' 100 kv

    Select Case cmbosc1.ListIndex
    
    Case 0
        vltdiv = gainvlt(7) / 0.1
        Call mdl1.leftlable(0.1 * 4)
    Case 1
        vltdiv = gainvlt(7) / 1
        Call mdl1.leftlable(1 * 4)
    Case 2
       vltdiv = gainvlt(7) / 2
        Call mdl1.leftlable(2 * 4)
    Case 3
        vltdiv = gainvlt(7) / 5
        Call mdl1.leftlable(5 * 4)
    Case 4
        vltdiv = gainvlt(7) / 8
        Call mdl1.leftlable(8 * 4)
    Case 5
        vltdiv = gainvlt(7) / 10
        Call mdl1.leftlable(10 * 4)
    Case 6
        vltdiv = gainvlt(7) / 20
        Call mdl1.leftlable(20 * 4)
    Case 7
        vltdiv = gainvlt(7) / 30
        Call mdl1.leftlable(30 * 4)
    Case 8
        vltdiv = gainvlt(7) / 40
        Call mdl1.leftlable(40 * 4)
    Case 9
        vltdiv = gainvlt(7) / 50
        Call mdl1.leftlable(50 * 4)
    Case 10
        vltdiv = gainvlt(7) / 60
        Call mdl1.leftlable(60 * 4)
    Case 11
        vltdiv = gainvlt(7) / 70
        Call mdl1.leftlable(70 * 4)
    Case 12
        vltdiv = gainvlt(7) / 80
        Call mdl1.leftlable(80 * 4)
    End Select
End Select

End Sub

Private Sub cmbvsel_Click()
Select Case cmbvsel.ListIndex

Case 0 ' 10 kv

    Select Case cmbosc1.ListIndex
    
    Case 0
        vltdiv = gainvlt(1) / 0.1
        Call mdl1.leftlable(0.1 * 4)
    Case 1
        vltdiv = gainvlt(1) / 1
        Call mdl1.leftlable(1 * 4)
    Case 2
       vltdiv = gainvlt(1) / 2
        Call mdl1.leftlable(2 * 4)
    Case 3
        vltdiv = gainvlt(1) / 5
        Call mdl1.leftlable(5 * 4)
    Case 4
        vltdiv = gainvlt(1) / 8
        Call mdl1.leftlable(8 * 4)
    Case 5
        vltdiv = gainvlt(1) / 10
        Call mdl1.leftlable(10 * 4)
    Case 6
        vltdiv = gainvlt(1) / 20
        Call mdl1.leftlable(20 * 4)
    Case 7
        vltdiv = gainvlt(1) / 30
        Call mdl1.leftlable(30 * 4)
    Case 8
        vltdiv = gainvlt(1) / 40
        Call mdl1.leftlable(40 * 4)
    Case 9
        vltdiv = gainvlt(1) / 50
        Call mdl1.leftlable(50 * 4)
    Case 10
        vltdiv = gainvlt(1) / 60
        Call mdl1.leftlable(60 * 4)
    Case 11
        vltdiv = gainvlt(1) / 70
        Call mdl1.leftlable(70 * 4)
    Case 12
        vltdiv = gainvlt(1) / 80
        Call mdl1.leftlable(80 * 4)
    End Select

Case 1   ' 50 kv

    Select Case cmbosc1.ListIndex
    
    Case 0
        vltdiv = gainvlt(4) / 0.1
        Call mdl1.leftlable(0.1 * 4)
    Case 1
        vltdiv = gainvlt(4) / 1
        Call mdl1.leftlable(1 * 4)
    Case 2
       vltdiv = gainvlt(4) / 2
        Call mdl1.leftlable(2 * 4)
    Case 3
        vltdiv = gainvlt(4) / 5
        Call mdl1.leftlable(5 * 4)
    Case 4
        vltdiv = gainvlt(4) / 8
        Call mdl1.leftlable(8 * 4)
    Case 5
        vltdiv = gainvlt(4) / 10
        Call mdl1.leftlable(10 * 4)
    Case 6
        vltdiv = gainvlt(4) / 20
        Call mdl1.leftlable(20 * 4)
    Case 7
        vltdiv = gainvlt(4) / 30
        Call mdl1.leftlable(30 * 4)
    Case 8
        vltdiv = gainvlt(4) / 40
        Call mdl1.leftlable(40 * 4)
    Case 9
        vltdiv = gainvlt(4) / 50
        Call mdl1.leftlable(50 * 4)
    Case 10
        vltdiv = gainvlt(4) / 60
        Call mdl1.leftlable(60 * 4)
    Case 11
        vltdiv = gainvlt(4) / 70
        Call mdl1.leftlable(70 * 4)
    Case 12
        vltdiv = gainvlt(4) / 80
        Call mdl1.leftlable(80 * 4)
    End Select
    
Case 2  ' 100 kv

    Select Case cmbosc1.ListIndex
    
    Case 0
        vltdiv = gainvlt(7) / 0.1
        Call mdl1.leftlable(0.1 * 4)
    Case 1
        vltdiv = gainvlt(7) / 1
        Call mdl1.leftlable(1 * 4)
    Case 2
       vltdiv = gainvlt(7) / 2
        Call mdl1.leftlable(2 * 4)
    Case 3
        vltdiv = gainvlt(7) / 5
        Call mdl1.leftlable(5 * 4)
    Case 4
        vltdiv = gainvlt(7) / 8
        Call mdl1.leftlable(8 * 4)
    Case 5
        vltdiv = gainvlt(7) / 10
        Call mdl1.leftlable(10 * 4)
    Case 6
        vltdiv = gainvlt(7) / 20
        Call mdl1.leftlable(20 * 4)
    Case 7
        vltdiv = gainvlt(7) / 30
        Call mdl1.leftlable(30 * 4)
    Case 8
        vltdiv = gainvlt(7) / 40
        Call mdl1.leftlable(40 * 4)
    Case 9
        vltdiv = gainvlt(7) / 50
        Call mdl1.leftlable(50 * 4)
    Case 10
        vltdiv = gainvlt(7) / 60
        Call mdl1.leftlable(60 * 4)
    Case 11
        vltdiv = gainvlt(7) / 70
        Call mdl1.leftlable(70 * 4)
    Case 12
        vltdiv = gainvlt(7) / 80
        Call mdl1.leftlable(80 * 4)
    End Select
End Select

End Sub




Private Sub cmdclr2_Click()
cd1.ShowColor
ch2color = cd1.Color
End Sub

Private Sub cmdclrback_Click()
cd1.ShowColor
pic1.BackColor = cd1.Color
End Sub

Private Sub cmdend_Click()
Dim rtv As Integer
If cmdstoptest.Enabled = True Then
    rtv = MsgBox("If You Want Exit From Test; Please Stop Test Prccess . ", vbOKOnly + vbCritical + vbDefaultButton2, "Warnning")
Else

    rtv = MsgBox("Do you want to Exit Test", vbYesNo + vbCritical + vbDefaultButton2, "Warnning")
    If rtv = 6 Then
        
       Unload Me
       FrmInterface.Show
    End If
End If
End Sub








Private Sub cmdsnnext_Click()
If cmdsnreg.Enabled = True And txtsn.Enabled = True Then
    
    txtsn.Text = Str(Val(txtsn.Text) + 1)
    arrcnt = arrcnt + 1
    cmdsnreg.SetFocus
End If
If Val(txtsn.Text) = mdl1.srn(mdl1.arrnum) And txtsn.Enabled = True Then
    cmdsnnext.Enabled = False
End If
If Val(txtsn.Text) > mdl1.srn(1) And txtsn.Enabled = True Then
    cmdsnpre.Enabled = True
End If

End Sub

Private Sub cmdsnpre_Click()
If cmdsnreg.Enabled = True And txtsn.Enabled = True Then
    txtsn.Text = Str(Val(txtsn.Text) - 1)
    arrcnt = arrcnt - 1
     cmdsnreg.SetFocus
End If
If Val(txtsn.Text) = mdl1.srn(1) And txtsn.Enabled = True Then
    cmdsnpre.Enabled = False
End If
If Val(txtsn.Text) < mdl1.srn(mdl1.arrnum) And txtsn.Enabled = True Then
    cmdsnnext.Enabled = True
End If

End Sub

Private Sub cmdsnreg_Click()

If alwstrttst = True Then
    Call starttest
    cmdsnreg.Enabled = False
    txtsn.Enabled = False
    
    
    cmdtestok1.Enabled = True
    cmdstoptest.Enabled = True
    cmdtestok1.SetFocus
Else
    MsgBox "Enter serial number then click ,OK,", vbOKOnly + vbExclamation, "Error"
    txtsn.SetFocus
End If

End Sub

Private Sub cmdstoptest_Click()
stptst = True

 End Sub

Private Sub cmdtestok1_Click()
tok1 = True
End Sub


Private Sub cmdtestok1_KeyDown(KeyCode As Integer, Shift As Integer)
'Dim key As Integer
'Dim s As String
'key = 0
'If KeyAscii = 37 Or KeyAscii = 38 Or KeyAscii = 39 Or KeyAscii = 40 Then
'    key = KeyAscii
'    KeyAscii = 0
'End If
'If key = 37 Or ket = 40 Then
'    s = "up"
'ElseIf key = 38 Or key = 39 Then
'  s = "down"
'End If
End Sub


Private Sub cmdtestok2_Click()
tok2 = True
End Sub


Private Sub cmdtestok2_KeyDown(KeyCode As Integer, Shift As Integer)
'Dim key As Integer
'Dim s As String
'key = 0
'If KeyCode = 37 Or KeyCode = 38 Or KeyCode = 39 Or KeyCode = 40 Then
'    key = KeyCode
'    KeyCode = 0
'End If
'If key = 37 Or ket = 40 Then
'    s = "up"
'ElseIf key = 38 Or key = 39 Then
'  s = "down"
'End If

End Sub


Private Sub cmdtestok3_Click()
tok3 = True
End Sub


Private Sub cmdtestok3_KeyDown(KeyCode As Integer, Shift As Integer)
'Dim key As Integer
'Dim s As String
'key = 0
'If KeyAscii = 37 Or KeyAscii = 38 Or KeyAscii = 39 Or KeyAscii = 40 Then
'    key = KeyAscii
'    KeyAscii = 0
'End If
'If key = 37 Or ket = 40 Then
'    s = "up"
'ElseIf key = 38 Or key = 39 Then
'  s = "down"
'End If

End Sub


Private Sub cmdtestok4_Click()
tok4 = True
End Sub

Private Sub cmdtestok4_KeyDown(KeyCode As Integer, Shift As Integer)
'Dim key As Integer
'Dim s As String
'key = 0
'If KeyAscii = 37 Or KeyAscii = 38 Or KeyAscii = 39 Or KeyAscii = 40 Then
'    key = KeyAscii
'    KeyAscii = 0
'End If
'If key = 37 Or ket = 40 Then
'    s = "up"
'ElseIf key = 38 Or key = 39 Then
'  s = "down"
'End If

End Sub







Private Sub Form_Activate()
Dim code As Double
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

mdl1.trs1 = 0
mdl1.trs2 = 0
mdl1.trs3 = 0
mdl1.trs4 = 0
timchck.Enabled = False
cmdstoptest.Enabled = False
alwstrttst = False

'*************************** DATAMAIN configuration*******
'dbadresmain = "\\Pars_server\data1\SG\HVTEST"
With datatestprmt
    .RecordSource = "gain"
    .Refresh
End With

With datatest
    .Recordset.MoveFirst
    .Refresh
    
End With

With datatestprmt
    .Refresh
    .Recordset.MoveFirst
    
    gainvlt(0) = .Recordset.Fields("Voltage_gain_10_2")
    gainvlt(1) = .Recordset.Fields("Voltage_gain_10_6")
    gainvlt(2) = .Recordset.Fields("Voltage_gain_10_10")
    gainvlt(3) = .Recordset.Fields("Voltage_gain_50_10")
    gainvlt(4) = .Recordset.Fields("Voltage_gain_50_25")
    gainvlt(5) = .Recordset.Fields("Voltage_gain_50_50")
    gainvlt(6) = .Recordset.Fields("Voltage_gain_100_30")
    gainvlt(7) = .Recordset.Fields("Voltage_gain_100_60")
    gainvlt(8) = .Recordset.Fields("Voltage_gain_100_100")
    
    gaincrt(0) = .Recordset.Fields("Current_gain_0_1")
    gaincrt(1) = .Recordset.Fields("Current_gain_0_15")
    gaincrt(2) = .Recordset.Fields("Current_gain_0_2")
    gaincrt(3) = .Recordset.Fields("Current_gain_0_5")
    gaincrt(4) = .Recordset.Fields("Current_gain_0_7")
    gaincrt(5) = .Recordset.Fields("Current_gain_1")
    gaincrt(6) = .Recordset.Fields("Current_gain_1_2")
    gaincrt(7) = .Recordset.Fields("Current_gain_1_5")
    gaincrt(8) = .Recordset.Fields("Current_gain_2")
    gaincrt(9) = .Recordset.Fields("Current_gain_2_5")
    gaincrt(10) = .Recordset.Fields("Current_gain_3")
    gaincrt(11) = .Recordset.Fields("Current_gain_3_5")
    gaincrt(12) = .Recordset.Fields("Current_gain_4")
    gaincrt(13) = .Recordset.Fields("Current_gain_4_5")
    gaincrt(14) = .Recordset.Fields("Current_gain_5")
    
    gainpd(0) = .Recordset.Fields("PD_gain_2")
    gainpd(1) = .Recordset.Fields("PD_gain_10")
    gainpd(2) = .Recordset.Fields("PD_gain_20")
End With

With datatestprmt
    .RecordSource = "testprmt"
    .Refresh
    
End With

'//////////////////// common code for cheng arrester type  in cmbtype

mdl1.selectedarr = cmbtype.List(cmbtype.ListIndex)

cmdtestok1.Enabled = False
cmdtestok2.Enabled = False
cmdtestok3.Enabled = False
cmdtestok4.Enabled = False

framtest1.ForeColor = vbBlack
lbliref1.ForeColor = vbBlack
lblarrtype.ForeColor = vbBlack
lbluref1.ForeColor = vbBlack
framtest4.ForeColor = vbBlack
lbluref4.ForeColor = vbBlack
lblic4.ForeColor = vbBlack
framtest3.ForeColor = vbBlack
lbluc3.ForeColor = vbBlack
lblpdacp3.ForeColor = vbBlack
framtest2.ForeColor = vbBlack
lblpdacp2.ForeColor = vbBlack
lblur2.ForeColor = vbBlack
cmdsnreg.Enabled = True

',,,,,,,,,,,,,,,,,,,,,,,,,

lblsp.Caption = "ARRESTER TYPE"
txtsn.Enabled = True
txtsn.Text = ""

cmdsnnext.Enabled = True
cmdsnpre.Enabled = True
cmdsnnext.Visible = True
cmdsnpre.Visible = True

arrcnt = 0
lblsp.Caption = "ARRESTER TYPE"
arrcnt = 1
mdl1.fdim (1)
mdl1.tested(1) = True


'""""""""""""""""""""""""""""""""""""""""""""""
mdl1.arrnum = 1
''''''''''''''''''
lbliref1.Caption = "Iref =... [mA]"
lblarrtype.Caption = "Arrester Type :....."
lbluref1.Caption = "  " + " ...< Uref <... " + " [kV]"
lbluref4.Caption = " Uc =... [kV]"
lblic4.Caption = "Ic(accp)=... [mA]"
lbluc3.Caption = "1.05*Uc=... [kV]"
lblpdacp3.Caption = "Max. PD= ...[pC]"
lblur2.Caption = "Ur =... [kV]"
lblpdacp2.Caption = "Max. PD= ...[pC]"
HVtest_Repetation_No = 0
'///////////////////output control setting///////////////

ch1color = vbBlue
ch2color = vbRed
txtsn.Enabled = True
txtch1.Text = ""
txtch2.Text = ""
txtch3.Text = ""
txtch4.Text = ""
plottime.Enabled = False
timchck.Enabled = False
cmbosc1_Click
cmbosc2_Click
cmbosc2p_Click

If framsrn.Enabled = True Then
    txtsn.SetFocus
End If


End Sub



Private Sub Form_Load()
Dim hei, wid, h1, w1 As Long

Dim devNum As Long
Dim devDesc As String
Dim devMode As AccessMode
Dim modIndex As Long




'////////////  setting and get features of PCI CARD (ADVANTECH)  /////////////

VIBuffer.getSelectedDevice devNum, devDesc, devMode, modIndex
err = Success
devNum = VIBuffer.BoardId

err = VIBuffer.Prepare


If (err <> 0) Then
    Response = MsgBox("Error in Reading A/D Card", vbOKOnly, "Error!!")
End If
  
VIBuffer.ReleaseBuffer
IOCartState = VIBuffer.state

    
'*************************** DTAMAIN configuration*******
'//////////////////////////////////////////////////////////////////////////////////////

alwstrttst = False
HVtest_Repetation_No = 0

frmtest.icon = LoadPicture(mdl1.icon)

With datatest
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.maindb + ";Jet OLEDB:Database Password=" + mdl1.MainDBPas
    .CommandType = adCmdTable
    .RecordSource = "DIST"
    .Refresh
    
End With

With datatestprmt
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.acldb + ";Jet OLEDB:Database Password=" + mdl1.AclDBPas
    .CommandType = adCmdTable
    .RecordSource = "gain"
    .Refresh
End With

With DataArrtype
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.acldb + ";Jet OLEDB:Database Password=" + mdl1.AclDBPas
    .CommandType = adCmdTable
    .RecordSource = "Type"
    .Refresh
End With

With DataArrtype.Recordset
    ReDim DataTabletype(0 To .RecordCount - 1) As String   ' define Table name in main data base
    .MoveFirst
    If .RecordCount >= 1 Then
        For i = 1 To .RecordCount
            cmbtype.AddItem .Fields("arr_Type"), i - 1
             DataTabletype(i - 1) = .Fields("TableName")
            .MoveNext
        Next i
        cmbtype.ListIndex = 0
    End If
End With

If DataTabletype(cmbtype.ListIndex) = "SUBST" Then  '  Station Class Arrester Coding
    datatest.RecordSource = "SUBST"
    mdl1.ArrTableName = "SUBST"
ElseIf DataTabletype(cmbtype.ListIndex) = "DIST" Then '  dISTRIBUTION Class Arrester Coding
    datatest.RecordSource = "DIST"
    mdl1.ArrTableName = "DIST"
End If


With datatestprmt
    .Refresh
    .Recordset.MoveFirst
    
    gainvlt(0) = .Recordset.Fields("Voltage_gain_10_2")
    gainvlt(1) = .Recordset.Fields("Voltage_gain_10_6")
    gainvlt(2) = .Recordset.Fields("Voltage_gain_10_10")
    gainvlt(3) = .Recordset.Fields("Voltage_gain_50_10")
    gainvlt(4) = .Recordset.Fields("Voltage_gain_50_25")
    gainvlt(5) = .Recordset.Fields("Voltage_gain_50_50")
    gainvlt(6) = .Recordset.Fields("Voltage_gain_100_30")
    gainvlt(7) = .Recordset.Fields("Voltage_gain_100_60")
    gainvlt(8) = .Recordset.Fields("Voltage_gain_100_100")
    
    gaincrt(0) = .Recordset.Fields("Current_gain_0_1")
    gaincrt(1) = .Recordset.Fields("Current_gain_0_15")
    gaincrt(2) = .Recordset.Fields("Current_gain_0_2")
    gaincrt(3) = .Recordset.Fields("Current_gain_0_5")
    gaincrt(4) = .Recordset.Fields("Current_gain_0_7")
    gaincrt(5) = .Recordset.Fields("Current_gain_1")
    gaincrt(6) = .Recordset.Fields("Current_gain_1_2")
    gaincrt(7) = .Recordset.Fields("Current_gain_1_5")
    gaincrt(8) = .Recordset.Fields("Current_gain_2")
    gaincrt(9) = .Recordset.Fields("Current_gain_2_5")
    gaincrt(10) = .Recordset.Fields("Current_gain_3")
    gaincrt(11) = .Recordset.Fields("Current_gain_3_5")
    gaincrt(12) = .Recordset.Fields("Current_gain_4")
    gaincrt(13) = .Recordset.Fields("Current_gain_4_5")
    gaincrt(14) = .Recordset.Fields("Current_gain_5")
    
    gainpd(0) = .Recordset.Fields("PD_gain_2")
    gainpd(1) = .Recordset.Fields("PD_gain_10")
    gainpd(2) = .Recordset.Fields("PD_gain_20")
End With

With datatestprmt
    .RecordSource = "testprmt"
    .Refresh
End With





'//////////////////////////////////////////////////////////////////////////////////

With cmbosc1
    .AddItem "100 V", 0
    .AddItem "1 KV", 1
    .AddItem "2 KV", 2
    .AddItem "5 KV", 3
    .AddItem "8 KV", 4
    .AddItem "10 KV", 5
    .AddItem "20 KV", 6
    .AddItem "30 KV", 7
    .AddItem "40 KV", 8
    .AddItem "50 KV", 9
    .AddItem "60 KV", 10
    .AddItem "70 KV", 11
    .AddItem "80 KV", 12
    .Text = "8 KV"
End With

With cmbosc2p
    .AddItem "0.2 Pc", 0
    .AddItem "0.5 Pc", 1
    .AddItem "1 Pc", 2
    .AddItem "2 Pc", 3
    .AddItem "4 Pc", 4
    .Text = "0.5 Pc"
End With

With cmbosc2
    .AddItem "0.1 mA", 0
    .AddItem "0.2 mA", 1
    .AddItem "0.4 mA", 2
    .AddItem "0.8 mA", 3
    .AddItem "1 mA", 4
    .AddItem "2 mA", 5
    .AddItem "3 mA", 6
    .Text = "0.4 mA"
End With

With cmbvsel
    .AddItem "10 KV", 0
    .AddItem "50 KV", 1
    .AddItem "100 KV", 2
    .Text = "50 KV"
End With

With CmbPDSel
    .AddItem "2 Pc", 0
    .AddItem "10 Pc", 1
    .AddItem "20 Pc", 2
    .ListIndex = 1
    .Text = "10 Pc"
    
End With
SelectedPDGain = gainpd(1)

pic1.AutoRedraw = True
pic1.DrawStyle = 0

draw = False
txtsn.Enabled = True
txtsn.Text = ""

With frmtest
    .Height = Screen.Height - 400
    .Width = Screen.Width
    .Top = 0
    .Left = 0
    .ScaleHeight = .Height * 1.06
    .ScaleWidth = .Width * 1.06
End With

frmtest.ScaleHeight = frmtest.Height * 1.06
frmtest.ScaleWidth = frmtest.Width * 1.06
hei = frmtest.Height
wid = frmtest.Width


With framsrn
    .Width = wid * 0.18
    .Height = hei * 0.25
    .Left = wid * 0.012
    .Top = hei * 0.05
End With

With framtest1
    .Width = wid * 0.2
    .Height = hei * 0.25
    .Left = framsrn.Left + framsrn.Width + 0.016 * wid
    .Top = framsrn.Top
End With

With framtest2
    .Width = wid * 0.2
    .Height = hei * 0.25
    .Left = framtest1.Left + framtest1.Width + 0.014 * wid
    .Top = framsrn.Top
End With

With framtest3
    .Width = wid * 0.2
    .Height = hei * 0.25
    .Left = framtest2.Left + framtest2.Width + 0.014 * wid
    .Top = framsrn.Top
End With

With framtest4
    .Width = wid * 0.2
    .Height = hei * 0.25
    .Left = framtest3.Left + framtest3.Width + 0.014 * wid
    .Top = framsrn.Top
End With

'/////////////////////////////////////////////////////////////////////////
With lblarrtype
    .Top = framsrn.Top + framsrn.Height * 1.15
    .Left = framsrn.Left
    .Width = framsrn.Width
    .Height = framtest1.Height * 0.12
    .Caption = "Arrester Type:    "
End With

With lblch1
    .Top = framsrn.Top + framsrn.Height * 1.15
    .Left = framtest1.Left
   ' .Width = wid * 0.01
    .Height = framtest1.Height * 0.12
End With
With txtch1
    .Top = lblch1.Top * 0.95
    .Left = lblch1.Left + lblch1.Width + (wid * 0.003)
    .Width = wid * 0.07
    .Height = framtest1.Height * 0.14
End With

With lblch2
    .Top = lblch1.Top
    .Left = txtch1.Left + txtch1.Width + (wid * 0.02)
    '.Width = wid * 0.01
    .Height = framtest1.Height * 0.12
End With

With txtch2
    .Top = txtch1.Top
    .Left = lblch2.Left + lblch2.Width + (wid * 0.003)
    .Width = wid * 0.07
    .Height = framtest1.Height * 0.14
End With

With lblch3
    .Top = lblch1.Top
    .Left = txtch2.Left + txtch2.Width + (wid * 0.02)
    '.Width = wid * 0.08
    .Height = framtest1.Height * 0.12
End With

With txtch3
    .Top = txtch1.Top
    .Left = lblch3.Left + lblch3.Width + (wid * 0.003)
    .Width = wid * 0.07
    .Height = framtest1.Height * 0.14
End With

With lblch4
    .Top = lblch1.Top
    .Left = txtch3.Left + txtch3.Width + (wid * 0.02)
    '.Width = wid * 0.08
    .Height = framtest1.Height * 0.12
End With

With txtch4
    .Top = txtch1.Top
    .Left = lblch4.Left + lblch4.Width + (wid * 0.003)
    .Width = wid * 0.07
    .Height = framtest1.Height * 0.12
End With

With cmdstoptest
    .Top = txtch1.Top
     .Width = wid * 0.08
    .Left = framtest4.Left + (framtest4.Width / 2) - (.Width / 2)
    .Height = framtest1.Height * 0.16
End With


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>Picture Box Setting>>>>>>>>>>>>>>>>
With picleft
    .Height = (hei * 0.55) * 9 / 8
    .Width = wid * 0.034
    .Left = framsrn.Left * 1.5
    .Top = lblch1.Top + lblch1.Height * 1.9
    .ScaleHeight = .Height
    .ScaleWidth = .Width
End With

With pic1
    .Height = hei * 0.55
    .Width = wid * 0.738
    .Left = picleft.Left + picleft.Width
    .Top = lblch1.Top + lblch1.Height * 2
    .ScaleHeight = .Height
    .ScaleWidth = .Width
End With

With picright
    .Height = picleft.Height
    .Width = wid * 0.035
    .Left = (pic1.Left * 1.2 + pic1.Width)
    .Top = picleft.Top
    .ScaleHeight = .Height
    .ScaleWidth = .Width
End With

picleft.Scale (0, 0)-(10, 9)
picright.Scale (0, 0)-(10, 9)



With lblvol
    .Top = pic1.Top + pic1.Height * 1.05
    .Left = pic1.Left
    .Width = wid * 0.12
    .Height = framtest1.Height * 0.12
End With

With chkIres
    .Top = pic1.Top + pic1.Height * 1.05
    .Left = pic1.Left + (pic1.Width * 0.4)
    .Width = wid * 0.2
    .Height = framtest1.Height * 0.12
End With

With lblcurpd
    .Top = lblvol.Top
    .Width = wid * 0.12
    .Left = pic1.Left + pic1.Width - .Width * 1.01
    .Height = framtest1.Height * 0.12
End With


lblch1.Caption = "Voltage [kV]: "
lblch2.Caption = "Current [mA]: "
lblcurpd.Caption = "Current [mA]"
'????????????????????????????????????????????
'          ***********************fram seting***********************



'********************************fram Serial Number

With lblsp
    .Width = framsrn.Width * 0.7
    .Height = framsrn.Height * 0.15
    .Top = framsrn.Height * 0.15
    .Left = framsrn.Width * 0.18
End With

With cmbtype
    .Width = framsrn.Width * 0.6
    .Top = framsrn.Height * 0.3
    .Left = framsrn.Width * 0.18
End With

With txtsn
    .Width = framsrn.Width * 0.6
    .Height = framsrn.Height * 0.15
    .Top = framsrn.Height * 0.5
    .Left = framsrn.Width * 0.18
End With

With cmdsnpre
    .Width = framsrn.Width * 0.1
    .Height = framsrn.Height * 0.15
    .Top = txtsn.Top
    .Left = framsrn.Width * 0.05
End With

With cmdsnnext
    .Width = framsrn.Width * 0.1
    .Height = framsrn.Height * 0.15
    .Top = txtsn.Top
    .Left = framsrn.Width * 0.82
End With

With cmdsnreg
    .Width = framsrn.Width * 0.59
    .Height = framsrn.Height * 0.15
    .Top = framsrn.Height * 0.7
    .Left = framsrn.Width * 0.2
End With



'      ******************************FRAM TEST 1
With lbliref1
    .Width = framtest1.Width * 0.85
    .Height = framtest1.Height * 0.15
    .Left = framtest1.Width * 0.05
    .Top = framtest1.Height * 0.14
  
End With
With lbluref1
    .Width = framtest1.Width * 0.85
    .Height = framtest1.Height * 0.25
    .Left = framtest1.Width * 0.05
    .Top = framtest1.Height * 0.29
End With
With lblureft1
    .Width = framtest1.Width * 0.22
    .Height = framtest1.Height * 0.15
    .Left = framtest1.Width * 0.01
    .Top = framtest1.Height * 0.5
End With
With txturef1
    .Width = framtest1.Width * 0.5
    .Height = framtest1.Height * 0.15
    .Left = framtest1.Width * 0.26
    .Top = framtest1.Height * 0.46
End With
With lblurefu1
    .Width = framtest1.Width * 0.15
    .Height = framtest1.Height * 0.15
    .Left = txturef1.Left + txturef1.Width * 1.05
    .Top = framtest1.Height * 0.5
End With
With cmdtestok1
    .Width = framtest1.Width * 0.5
    .Height = framtest1.Height * 0.15
    .Left = framtest1.Width * 0.26
    .Top = framtest1.Height * 0.7
End With
'***************************************Fram TEST 4
With lbluref4
    .Width = framtest4.Width * 0.85
    .Height = framtest4.Height * 0.15
    .Left = framtest4.Width * 0.05
    .Top = framtest4.Height * 0.14
End With
With lblic4
    .Width = framtest4.Width * 0.85
    .Height = framtest4.Height * 0.15
    .Left = framtest4.Width * 0.05
    .Top = framtest4.Height * 0.29
End With
With lblict4
    .Width = framtest4.Width * 0.2
    .Height = framtest4.Height * 0.15
    .Left = framtest4.Width * 0.02
    .Top = framtest4.Height * 0.5
End With
With txtic4
    .Width = framtest4.Width * 0.5
    .Height = framtest4.Height * 0.15
    .Left = framtest4.Width * 0.22
    .Top = framtest4.Height * 0.47
End With
With lblicu4
    .Width = framtest4.Width * 0.15
    .Height = framtest4.Height * 0.2
    .Left = framtest4.Width * 0.75
    .Top = framtest4.Height * 0.5
End With
With cmdtestok4
    .Width = framtest4.Width * 0.5
    .Height = framtest4.Height * 0.15
    .Left = framtest4.Width * 0.22
    .Top = framtest4.Height * 0.7
End With
'***********************************FRAM TEST 3
With lbluc3
    .Width = framtest3.Width * 0.85
    .Height = framtest3.Height * 0.15
    .Left = framtest3.Width * 0.05
    .Top = framtest3.Height * 0.14
End With
With lblpdacp3
    .Width = framtest3.Width * 0.85
    .Height = framtest3.Height * 0.15
    .Left = framtest3.Width * 0.05
    .Top = framtest3.Height * 0.29
End With
With lblpdt3
    .Width = framtest3.Width * 0.2
    .Height = framtest3.Height * 0.15
    .Left = framtest3.Width * 0.02
    .Top = framtest3.Height * 0.5
End With
With txtpd3
    .Width = framtest3.Width * 0.5
    .Height = framtest3.Height * 0.15
    .Left = framtest3.Width * 0.22
    .Top = framtest3.Height * 0.47
End With
With lblpdu3
    .Width = framtest3.Width * 0.15
    .Height = framtest3.Height * 0.2
    .Left = framtest3.Width * 0.76
    .Top = framtest3.Height * 0.5
End With
With cmdtestok3
    .Width = framtest3.Width * 0.5
    .Height = framtest3.Height * 0.15
    .Left = framtest3.Width * 0.22
    .Top = framtest3.Height * 0.7
End With
'**************************************FRAM TEST 4
With lblur2
    .Width = framtest2.Width * 0.85
    .Height = framtest4.Height * 0.15
    .Left = framtest4.Width * 0.05
    .Top = framtest4.Height * 0.14
End With
With lblpdacp2
    .Width = framtest2.Width * 0.85
    .Height = framtest2.Height * 0.15
    .Left = framtest2.Width * 0.05
    .Top = framtest2.Height * 0.29
End With
With lblpdt2
    .Width = framtest2.Width * 0.18
    .Height = framtest2.Height * 0.15
    .Left = framtest2.Width * 0.04
    .Top = framtest2.Height * 0.5
End With
With txtpd2
    .Width = framtest2.Width * 0.5
    .Height = framtest2.Height * 0.15
    .Left = framtest2.Width * 0.22
    .Top = framtest2.Height * 0.47
End With
With lblpdu2
    .Width = framtest2.Width * 0.15
    .Height = framtest2.Height * 0.2
    .Left = framtest2.Width * 0.76
    .Top = framtest2.Height * 0.5
End With
With cmdtestok2
    .Width = framtest2.Width * 0.5
    .Height = framtest2.Height * 0.15
    .Left = framtest2.Width * 0.22
    .Top = framtest2.Height * 0.7
End With
hei = frmtest.Height
wid = frmtest.Width
'?????????????????????????????????????????????????
'               scope frame setting
'???????????????????????????????????????????????
With framosc
    .Height = hei * 0.51
    .Width = wid * 0.2
     h1 = .Height
     w1 = .Width
    .Left = framtest4.Left
    .Top = pic1.Top
End With
'*************************************************************************

With lblosc1
    .Height = h1 * 0.05
    .Width = w1 * 0.4
    .Left = w1 * 0.05
    .Top = h1 * 0.09
End With
With cmbosc1
    .Width = w1 * 0.35
    .Left = w1 * 0.5
    .Top = h1 * 0.08
End With
'////////////////////////
With lblosc2
    .Height = h1 * 0.05
    .Width = w1 * 0.4
    .Left = w1 * 0.05
    .Top = h1 * 0.19
End With

With cmbosc2
    .Width = w1 * 0.35
    .Left = w1 * 0.5
    .Top = h1 * 0.18
End With

With lblosc2p
     .Height = h1 * 0.05
    .Width = w1 * 0.4
    .Left = w1 * 0.05
    .Top = h1 * 0.285
    .Enabled = False
End With
With cmbosc2p
    .Width = w1 * 0.35
    .Left = w1 * 0.5
    .Top = h1 * 0.28
    .Enabled = False
End With

With ln1
    .x1 = 0
    .x2 = w1
    .y1 = h1 * 0.37
    .y2 = h1 * 0.37
End With
'*******************************************************************************8
With lblclr
    .Height = h1 * 0.05
    .Width = w1 * 0.4
    .Left = w1 * 0.3
    .Top = h1 * 0.39
End With
'/////////////////////////
With cmdclr1
    .Height = h1 * 0.06
    .Width = w1 * 0.3
    .Left = w1 * 0.11
    .Top = h1 * 0.467
End With
With cmdclr2
    .Height = h1 * 0.06
    .Width = w1 * 0.3
    .Left = w1 * 0.54
    .Top = h1 * 0.467
End With
'*****************************
With cmdclrback
    .Height = h1 * 0.07
    .Width = w1 * 0.5
    .Left = w1 * 0.242
    .Top = h1 * 0.58
End With
With ln2
    .x1 = 0
    .x2 = w1
    .y1 = h1 * 0.695
    .y2 = h1 * 0.695
End With

'/*****************************************************************
With lblvoltsel
    .Height = h1 * 0.05
    .Width = w1 * 0.4
    .Left = w1 * 0.05
    .Top = h1 * 0.7
End With

With cmbvsel
    .Width = w1 * 0.3
    .Left = w1 * 0.05
    .Top = h1 * 0.785
End With

With LblPDSel
    .Height = h1 * 0.05
    .Width = w1 * 0.3
    .Left = w1 * 0.55
    .Top = h1 * 0.7
End With

With CmbPDSel
    .Width = w1 * 0.3
    .Left = w1 * 0.5
    .Top = h1 * 0.785
End With


'***************************************************************
'//////////////////////////////////////////////////////////////

With cmdend
    .Height = hei * 0.05
    .Width = wid * 0.08
    .Left = framosc.Left '+ (framosc.Width / 2) - .Width / 2
    .Top = (framosc.Height + framosc.Top) * 1.025
End With

With CmdCalibration
    .Height = hei * 0.05
    .Width = wid * 0.08
    .Left = framosc.Left + framosc.Width - .Width
    .Top = (framosc.Height + framosc.Top) * 1.025
End With
'******************************************
With lh1
    .x1 = 0
    .y1 = pic1.ScaleHeight * 1 / 8
    .x2 = pic1.ScaleWidth
    .y2 = pic1.ScaleHeight * 1 / 8
End With

With lh2
    .x1 = 0
    .y1 = pic1.ScaleHeight * 2 / 8
    .x2 = pic1.ScaleWidth
    .y2 = pic1.ScaleHeight * 2 / 8
End With
With lh3
    .x1 = 0
    .y1 = pic1.ScaleHeight * 3 / 8
    .x2 = pic1.ScaleWidth
    .y2 = pic1.ScaleHeight * 3 / 8
End With
With lh4
    .x1 = 0
    .y1 = pic1.ScaleHeight * 4 / 8
    .x2 = pic1.ScaleWidth
    .y2 = pic1.ScaleHeight * 4 / 8
End With
With lh5
    .x1 = 0
    .y1 = pic1.ScaleHeight * 5 / 8
    .x2 = pic1.ScaleWidth
    .y2 = pic1.ScaleHeight * 5 / 8
End With
With lh6
    .x1 = 0
    .y1 = pic1.ScaleHeight * 6 / 8
    .x2 = pic1.ScaleWidth
    .y2 = pic1.ScaleHeight * 6 / 8
End With
With lh7
    .x1 = 0
    .y1 = pic1.ScaleHeight * 7 / 8
    .x2 = pic1.ScaleWidth
    .y2 = pic1.ScaleHeight * 7 / 8
End With
With lv1
    .x1 = pic1.ScaleWidth * 1 / 10
    .x2 = pic1.ScaleWidth * 1 / 10
    .y2 = pic1.ScaleHeight
    .y1 = 0
End With
With lv2
    .x1 = pic1.ScaleWidth * 2 / 10
    .x2 = pic1.ScaleWidth * 2 / 10
    .y2 = pic1.ScaleHeight
    .y1 = 0
End With
With lv3
    .x1 = pic1.ScaleWidth * 3 / 10
    .x2 = pic1.ScaleWidth * 3 / 10
    .y2 = pic1.ScaleHeight
    .y1 = 0
End With

With lv4
    .x1 = pic1.ScaleWidth * 4 / 10
    .x2 = pic1.ScaleWidth * 4 / 10
    .y2 = pic1.ScaleHeight
    .y1 = 0
End With

With lv5
   .x1 = pic1.ScaleWidth * 5 / 10
    .x2 = pic1.ScaleWidth * 5 / 10
    .y2 = pic1.ScaleHeight
    .y1 = 0
End With

With lv6
    .x1 = pic1.ScaleWidth * 6 / 10
    .x2 = pic1.ScaleWidth * 6 / 10
    .y2 = pic1.ScaleHeight
    .y1 = 0
End With

With lv7
    .x1 = pic1.ScaleWidth * 7 / 10
    .x2 = pic1.ScaleWidth * 7 / 10
    .y2 = pic1.ScaleHeight
    .y1 = 0
End With
With lv8
    .x1 = pic1.ScaleWidth * 8 / 10
    .x2 = pic1.ScaleWidth * 8 / 10
    .y2 = pic1.ScaleHeight
    .y1 = 0
End With
With lv9
    .x1 = pic1.ScaleWidth * 9 / 10
    .x2 = pic1.ScaleWidth * 9 / 10
    .y2 = pic1.ScaleHeight
    .y1 = 0
End With
cmbtype.Enabled = True

'//////////////////////   User Access Setting ///////////////////////////

If mdl1.Test_Admin = True Or mdl1.Technical_Manager = True Then
    CmdCalibration.Enabled = True
Else
     CmdCalibration.Enabled = False
End If

frmtest.Caption = "ARTM                        Current Year : " + Str(mdl1.FullYear) + "                         Operator : " + mdl1.optname


If mdl1.Test_Admin = True Or mdl1.HVTest_Operator = True Then
   
    framsrn.Enabled = True
Else
  
    framsrn.Enabled = False
End If

End Sub


Private Sub plottime_Timer()

Dim i As Integer, j As Integer

'On Error GoTo Err_Handller

'Set Channel for data acquisition    [CH(0,1)=Current CH(1,2)=Voltage CH(3,4)=SC  CH(5,6)= PD]
    
Call selectgain   ' gain & start channel setting
'VIBuffer.setSelectedDevice
IOCartState = VIBuffer.state

err = VIBuffer.RunOnce

If (err <> 0) Then
    Response = MsgBox("Error in Reading A/D Card", vbOKOnly, "Error!!")
End If
plottime.Enabled = False

IOCartState = VIBuffer.state
' timchck.Interval = 10
'timchck.Enabled = True
      
    
 'Exit Sub
 
 'Err_Handller
 
 'Call MsgBox(Err.Number & vbCrLf & Err.Description, vbExclamation, App.Title)
 'Call stoptest
 
 'Resume Next
'///////////////////////////////////////////////////////////////
End Sub



Private Sub txtch1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


Private Sub txtch2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


Private Sub txtch3_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


Private Sub txtch4_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


Private Sub txtic4_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


Private Sub txtpd2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


Private Sub txtpd3_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub





Private Sub txtsn_Change()

cmdtestok1.Enabled = False
cmdtestok2.Enabled = False
cmdtestok3.Enabled = False
cmdtestok4.Enabled = False
cmdstoptest.Enabled = False
cmdsnreg.Enabled = True
alwstrttst = False

lbliref1.Caption = "Iref =... [mA]"
lblarrtype.Caption = "Arrester Type :....."
lbluref1.Caption = "  " + " ...< Uref <... " + " [kV]"
lbluref4.Caption = " Uc =... [kV]"
lblic4.Caption = "Ic(accp)=... [mA]"
lbluc3.Caption = "1.05*Uc=... [kV]"
lblpdacp3.Caption = "Max. PD= ...[pC]"
lblur2.Caption = "Ur =... [kV]"
lblpdacp2.Caption = "Max. PD= ...[pC]"
HVtest_Repetation_No = 0


End Sub

Private Sub txtsn_KeyPress(KeyAscii As Integer)
KeyAscii = mdl1.inputtest(KeyAscii)
alwstrttst = False

If KeyAscii = 13 Then
   cmdsnreg.SetFocus
End If

End Sub


Private Sub txtsn_LostFocus()
Dim Arrestercode As String
If Val(txtsn.Text) > 0 Then
    alwstrttst = False
    pic1.Cls
    mdl1.trs1 = 0
    mdl1.trs2 = 0
    mdl1.trs3 = 0
    mdl1.trs4 = 0
    testapp1 = False
    testapp2 = False
    testapp3 = False
    testapp4 = False
    lblch1.Caption = "Voltage [kV]: "
    lblch2.Caption = "Current [mA]: "
    lblcurpd.Caption = "Current [mA]"
    lblosc2.Enabled = True
    lblosc2p.Enabled = False
    cmbosc2.Enabled = True
    cmbosc2p.Enabled = False
    
    Call mdl1.rightlable(Val(cmbosc2.Text) * 4)
        
    If (Not IsNumeric(txtsn.Text)) Then
        'cmdsnreg.SetFocus
    Else
        mdl1.arrnum = 1
        datatest.Refresh
        With datatest.Recordset
            
                .MoveFirst
                
                .Find "Serial_Num=" + txtsn.Text, , adSearchForward, 1
                If .EOF = False Then ' entered S.N has been registered
                   
                    'Datatestprmt.Recordset.Open
                    datatestprmt.Recordset.MoveFirst
                    datatestprmt.Recordset.Find ("ID=") + Str(.Fields("Test_Inf")), , adSearchForward, 1
                   
                    Arrestercode = .Fields("Arrester_Type")
                    If datatestprmt.Recordset.EOF = True Then
                        MsgBox "Error In Reading Test Parameter." + Chr(13) + "Check Parameter of this arrester in Report page ", vbOKOnly, "ERROR"
                        'txtsn.Text = ""
                        
                    Else
                        Select Case .Fields("Result1")
              
                        Case 1   ' Production process hasn't been registered
                            
                            MsgBox "Production process hasn't been registered" + Chr(13) + "You cann't perform HV Test" + Chr(13) + "«» œ« «ÿ·«⁄«  ›—«Ì‰œ  Ê·Ìœ —« À»  ‰„«ÌÌœ.", vbOKOnly + vbExclamation, "Error"
                            
                            cmdtestok1.Enabled = False
                            cmdstoptest.Enabled = False
                            cmdtestok2.Enabled = False
                            cmdtestok3.Enabled = False
                            cmdtestok4.Enabled = False
                
                            'txtsn.SetFocus
                            txtsn.Text = ""
                            lbliref1.Caption = "Iref =... [mA]"
                            lblarrtype.Caption = "Arrester Type :....."
                            lbluref1.Caption = "  " + " ...< Uref <... " + " [kV]"
                            lbluref4.Caption = " Uc =... [kV]"
                            lblic4.Caption = "Ic(accp)=... [mA]"
                            lbluc3.Caption = "1.05*Uc=... [kV]"
                            lblpdacp3.Caption = "Max. PD= ...[pC]"
                            lblur2.Caption = "Ur =... [kV]"
                            lblpdacp2.Caption = "Max. PD= ...[pC]"
                            mdl1.tested(1) = True
                            HVtest_Repetation_No = 0
                                               
                            Case 2
                                  
                                 'Apply Test1      **********************************
                                With datatestprmt.Recordset
                                    lbliref1.Caption = "Iref = " + Format((.Fields("Iref")), "##0.0") + "  [mA]"
                                    lblarrtype.Caption = "Arrester Type : " + Arrestercode
                                    lbluref1.Caption = " " + Format((.Fields("Urefmin") / .Fields("unit_number")), "##0.0") + " < Uref < " + Format((.Fields("Urefmax") / .Fields("unit_number")), "##0.0") + "  [kV]"
                                    lbluref4.Caption = " Uc = " + Format(.Fields("Ucon") / .Fields("unit_number"), "##0.0") + "  [kV]"
                                    lblic4.Caption = "Ic(accp)= " + Format(.Fields("Icaccp"), "##0.0") + "  [mA]"
                                    lbluc3.Caption = "1.05*Uc= " + Format(Str((.Fields("Ucon") / .Fields("unit_number")) * 1.05), "##0.0") + "  [kV]"
                                    
                                    mdl1.pdaccp = .Fields("accept_pd")
                                    
                                    lblpdacp3.Caption = "Max. PD= " + Format((mdl1.pdaccp), "#0.0") + "  [pC]"
                                    lblur2.Caption = "Ur = " + Format((.Fields("Urated") / .Fields("unit_number")), "##0.0") + "  [kV]"
                                    lblpdacp2.Caption = "Max. PD= " + Format((mdl1.pdaccp), "#0.0") + "  [pC]"
                                    
                                   
                                    iref1 = .Fields("Iref")
                                    vpd1 = .Fields("Urated") / .Fields("unit_number")
                                    vpd2 = (.Fields("Ucon") / .Fields("unit_number")) * 1.05
                                    vref2 = .Fields("Ucon") / .Fields("unit_number")
                                    icacp4 = .Fields("Icaccp")
                                    HVtest_Repetation_No = 1  ' Applying test 1
                                    alwstrttst = True
                                    
                                End With
                               
                               
                                cmdtestok2.Enabled = False
                                cmdtestok3.Enabled = False
                                cmdtestok4.Enabled = False
                                framtest1.ForeColor = vbBlue
                                lbliref1.ForeColor = vbBlue
                                lbluref1.ForeColor = vbBlue
                                framtest4.ForeColor = vbBlack
                                lbluref4.ForeColor = vbBlack
                                lblic4.ForeColor = vbBlack
                                framtest3.ForeColor = vbBlack
                                lbluc3.ForeColor = vbBlack
                                lblpdacp3.ForeColor = vbBlack
                                framtest2.ForeColor = vbBlack
                                lblpdacp2.ForeColor = vbBlack
                                lblur2.ForeColor = vbBlack
                                mdl1.tested(1) = False
                            
                        Case 3
                        
                            If .Fields("Result2") = 2 Or .Fields("Result2") = Null Then
                                rtv = MsgBox("HV Test1 has been Performed successfully." + Chr(13) + "Do you want to repeat HV Test? ", vbYesNo + vbExclamation, "Error")
                            
                            ElseIf .Fields("Result2") = 3 Then
                                rtv = MsgBox("HV Test 1 & Test 2 have been Performed successfully." + Chr(13) + "Do you want to repeat HV Test2? ", vbYesNo + vbExclamation, "Error")
                        
                            ElseIf .Fields("Result2") > 3 Then
                                rtv = MsgBox("HV Test1 has been Performed successfully. But test 2 failled." + Chr(13) + "Do you want to repeat HV Test2? ", vbYesNo + vbExclamation, "Error")
                            
                            End If
                            
                            If rtv = 6 Then
                                'Apply Test2      **********************************
                                HVtest_Repetation_No = 2  ' Applying test 2
                                cmdtestok2.Enabled = False
                                cmdtestok3.Enabled = False
                                cmdtestok4.Enabled = False
                                framtest1.ForeColor = vbBlue
                                lbliref1.ForeColor = vbBlue
                                lbluref1.ForeColor = vbBlue
                                framtest4.ForeColor = vbBlack
                                lbluref4.ForeColor = vbBlack
                                lblic4.ForeColor = vbBlack
                                framtest3.ForeColor = vbBlack
                                lbluc3.ForeColor = vbBlack
                                lblpdacp3.ForeColor = vbBlack
                                framtest2.ForeColor = vbBlack
                                lblpdacp2.ForeColor = vbBlack
                                lblur2.ForeColor = vbBlack
                                mdl1.tested(1) = False
                                
                                With datatestprmt.Recordset
                                    lbliref1.Caption = "Iref = " + Format((.Fields("Iref")), "##0.0") + "  [mA]"
                                    lblarrtype.Caption = "Arrester Type : " + Arrestercode
                                    lbluref1.Caption = " " + Format((.Fields("Urefmin") / .Fields("unit_number")), "##0.0") + " < Uref < " + Format((.Fields("Urefmax") / .Fields("unit_number")), "##0.0") + "  [kV]"
                                    lbluref4.Caption = " Uc = " + Format(.Fields("Ucon") / .Fields("unit_number"), "##0.0") + "  [kV]"
                                    lblic4.Caption = "Ic(accp)= " + Format(.Fields("Icaccp"), "##0.0") + "  [mA]"
                                    lbluc3.Caption = "1.05*Uc= " + Format((.Fields("Ucon") / .Fields("unit_number") * 1.05), "##0.0") + "  [kV]"
                                    mdl1.pdaccp = .Fields("accept_pd")
                                    
                                    lblpdacp3.Caption = "Max. PD= " + Format((mdl1.pdaccp), "##0.0") + "  [pC]"
                                    lblur2.Caption = "Ur = " + Format((.Fields("Urated") / .Fields("unit_number")), "##0.0") + "  [kV]"
                                    lblpdacp2.Caption = "Max. PD= " + Format((mdl1.pdaccp), "#0.0") + "  [pC]"
                                    
                                   
                                    iref1 = .Fields("Iref")
                                    vpd1 = .Fields("Urated") / .Fields("unit_number")
                                    vpd2 = (.Fields("Ucon") / .Fields("unit_number")) * 1.05
                                    vref2 = .Fields("Ucon") / .Fields("unit_number")
                                    icacp4 = .Fields("Icaccp")
                                    
                                     
                                    alwstrttst = True
                                End With
                            End If
                            
                        
                        Case Else
                        
                            If .Fields("Result2") = 2 Or .Fields("Result2") = Null Then
                                rtv = MsgBox("HV Test1 failed in previous test." + Chr(13) + "You should repeat HV test.", vbOKOnly + vbExclamation, "Error")
                            ElseIf .Fields("Result2") = 3 Then
                                rtv = MsgBox("HV Test 1 failled, but Test 2 have been Performed successfully." + Chr(13) + "You are trying to repeat HV Test2.", vbOKOnly + vbExclamation, "Error")
                        
                            ElseIf .Fields("Result2") > 3 Then
                                rtv = MsgBox("HV Test1 $ Test 2 failled." + Chr(13) + "You should repeat HV Test2.", vbOKOnly + vbExclamation, "Error")
                            
                            End If
                        
                            'Aplly Test2
                                HVtest_Repetation_No = 2  ' Applying test 2
                                cmdtestok2.Enabled = False
                                cmdtestok3.Enabled = False
                                cmdtestok4.Enabled = False
                                framtest1.ForeColor = vbBlue
                                lbliref1.ForeColor = vbBlue
                                lbluref1.ForeColor = vbBlue
                                framtest4.ForeColor = vbBlack
                                lbluref4.ForeColor = vbBlack
                                lblic4.ForeColor = vbBlack
                                framtest3.ForeColor = vbBlack
                                lbluc3.ForeColor = vbBlack
                                lblpdacp3.ForeColor = vbBlack
                                framtest2.ForeColor = vbBlack
                                lblpdacp2.ForeColor = vbBlack
                                lblur2.ForeColor = vbBlack
                                mdl1.tested(1) = False
                         
                            With datatestprmt.Recordset
                                    lbliref1.Caption = "Iref = " + Format((.Fields("Iref")), "##0.0") + "  [mA]"
                                    lblarrtype.Caption = "Arrester Type : " + Arrestercode
                                    lbluref1.Caption = " " + Format((.Fields("Urefmin") / .Fields("unit_number")), "##0.0") + " < Uref < " + Format((.Fields("Urefmax") / .Fields("unit_number")), "##0.0") + "  [kV]"
                                    lbluref4.Caption = " Uc = " + Format(.Fields("Ucon") / .Fields("unit_number"), "##0.0") + "  [kV]"
                                    lblic4.Caption = "Ic(accp)= " + Format(.Fields("Icaccp"), "##0.0") + "  [mA]"
                                    lbluc3.Caption = "1.05*Uc= " + Format(Str((.Fields("Ucon") / .Fields("unit_number")) * 1.05), "##0.0##") + "  [kV]"
                                    mdl1.pdaccp = .Fields("accept_pd")
                                    
                                    lblpdacp3.Caption = "Max. PD= " + Format((mdl1.pdaccp), "#0.0") + "  [pC]"
                                    lblur2.Caption = "Ur = " + Format((.Fields("Urated") / .Fields("unit_number")), "##0.0") + "  [kV]"
                                    lblpdacp2.Caption = "Max. PD= " + Format((mdl1.pdaccp), "#0.0") + "  [pC]"
                                    
                                   
                                    iref1 = .Fields("Iref")
                                    vpd1 = .Fields("Urated") / .Fields("unit_number")
                                    vpd2 = (.Fields("Ucon") / .Fields("unit_number")) * 1.05
                                    vref2 = .Fields("Ucon") / .Fields("unit_number")
                                    icacp4 = .Fields("Icaccp")
                                    
                                    alwstrttst = True
                                    
                                End With
                           
                            
                        End Select
                        
                    End If
                    'Datatestprmt.Recordset.Close
                Else
                        MsgBox "Selected serial number hasn't been registerred." + Chr(13) + "Enter registerred serial number", vbOKOnly + vbExclamation, "Attention"
                        
                        lbliref1.Caption = "Iref =... [mA]"
                        lblarrtype.Caption = "Arrester Type : ..... "
                        lbluref1.Caption = "  " + " ...< Uref <... " + " [kV]"
                        lbluref4.Caption = " Uc =... [kV]"
                        lblic4.Caption = "Ic(accp)=... [mA]"
                        lbluc3.Caption = "1.05*Uc=... [kV]"
                        lblpdacp3.Caption = "Max. PD= ...[pC]"
                        lblur2.Caption = "Ur =... [kV]"
                        lblpdacp2.Caption = "Max. PD= ...[pC]"
                        mdl1.tested(1) = True
                        HVtest_Repetation_No = 0
                        
                End If
                '.Close
            End With
    End If
    
    If alwstrttst = True Then
        cmdsnreg.SetFocus
    End If
End If
End Sub


Private Sub txturef1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


Private Sub VIBuffer_Stopped(ByVal Offset As Long, ByVal Count As Long)
Dim drwable As Boolean

draw = False


'On Error GoTo Err_Handller

    
    '////////////      Get the AI conversion data    ////////////
    
   IOCartState = VIBuffer.state
    
    If cmdtestok2.Enabled = True Or cmdtestok3.Enabled = True Then ' PD and Voltage
    '(Card Read Voltage, SC. , PD  Conditions consecutively and separately in three steps)
            
        If reset = True Then   ' I/O card Have Read Short Circuit Data
           
            err = VIBuffer.GetData(Count, sDataBuffer)
            
           If (err <> 0) Then
                Response = MsgBox("Error in Getting I/O card Info", vbOKOnly, "Error!!")
            End If
            
            IOCartState = VIBuffer.state
            reset = False
            draw = False ' draw curve after reading PD
                
        Else ' I/O card Have Read Main Data
            
            If state = 0 Then        'I/O card Have Read PD
                
               
                err = VIBuffer.GetData(Count, sDataBufferpd())
                
                If (err <> 0) Then
                    Response = MsgBox("Error in Reading A/D Card", vbOKOnly, "Error!!")
                End If
                                
                draw = True
                    
            ElseIf state = 1 Then         'I/O card Have Read Voltage buffer
                err = VIBuffer.GetData(Count, sDataBufferMain2())
                If (err <> 0) Then
                    Response = MsgBox("Error in Reading A/D Card", vbOKOnly, "Error!!")
                End If
                
                draw = False
                reset = True  ' Get ready to read SC Data in next step
                    
            End If
        
            If state = 0 Then  'I/O card Have Read PD
                state = 1
            ElseIf state = 1 Then 'I/O card Have Read Voltage
                state = 0
            End If
    
        End If
            
    ElseIf cmdtestok1.Enabled = True Or cmdtestok4.Enabled = True Then '(Card Read Voltage & Current Simultaneously then read SC Condition)
     
        If reset = True Then
           
            err = VIBuffer.GetData(Count, sDataBuffer())
            If (err <> 0) Then
                Response = MsgBox("Error in Reading A/D Card", vbOKOnly, "Error!!")
            End If
            
            reset = False
            draw = False
                
        Else    ' I/O card Have read Main Signals of Voltage and Current
            err = VIBuffer.GetData(Count, sDataBufferMain())
            If (err <> 0) Then
                Response = MsgBox("Error in Getting I/O card Info", vbOKOnly, "Error!!")
            End If
              
            draw = True
            reset = True
               
        End If
          
    End If
    
    
    
    '*********    Draw data in image box and detect peak value  ******
     If draw = True Then
        Call drawcurve    ' draw curve
        draw = False
    End If
    
    plottime.Enabled = True ' Ready for next Reading
    
    If tok1 = True Then
        Call test1
        plottime.Enabled = True
    ElseIf tok2 = True Then
        Call test2
        plottime.Enabled = True
    ElseIf tok3 = True Then
        Call test3
        plottime.Enabled = True
        
    ElseIf tok4 = True Then  ' test has been Compeleted
        plottime.Enabled = False
        Call test4
    End If
    
    If stptst = True Then
        plottime.Enabled = False
        Call stoptest
    End If
    
    If endtest = True Then
        plottime.Enabled = False
        Call stoptest1
    End If

Exit Sub

'Err_Handller

 'Call MsgBox(Err.Number & vbCrLf & Err.Description, vbExclamation, App.Title)
 'Call stoptest
 
 'Resume Next


End Sub
