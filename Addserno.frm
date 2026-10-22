VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmmain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Arrester Data"
   ClientHeight    =   9570
   ClientLeft      =   150
   ClientTop       =   330
   ClientWidth     =   15405
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   15405
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox CmbMonth 
      BeginProperty Font 
         Name            =   "B Zar"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4440
      TabIndex        =   24
      Text            =   "210200"
      Top             =   8760
      Width           =   2055
   End
   Begin VB.ComboBox CmbPrdcCode 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4440
      TabIndex        =   23
      Text            =   "210200"
      Top             =   8280
      Width           =   2055
   End
   Begin VB.ComboBox CmbGLN 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1680
      TabIndex        =   22
      Text            =   "36050"
      Top             =   8280
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DB Update"
      Height          =   375
      Left            =   10920
      TabIndex        =   71
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc datablock 
      Height          =   330
      Left            =   12360
      Top             =   8400
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
      Caption         =   "datablock"
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
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   12360
      Top             =   8040
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
      Caption         =   "Data1"
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
   Begin MSAdodcLib.Adodc dataarrtype 
      Height          =   330
      Left            =   12360
      Top             =   8760
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
      Caption         =   "dataarrtype"
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
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
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
      Left            =   9000
      TabIndex        =   33
      Top             =   8160
      Width           =   1335
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "End"
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
      Left            =   10680
      TabIndex        =   34
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton cmdapply 
      Caption         =   "Apply"
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
      Left            =   6720
      TabIndex        =   32
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Frame framblock 
      Caption         =   "ZnO Block Data"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   8400
      TabIndex        =   48
      Top             =   4920
      Width           =   5895
      Begin VB.ComboBox cmbblctype 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   15
         ToolTipText     =   "”«“‰œÂ ﬁ—’"
         Top             =   960
         Width           =   1335
      End
      Begin VB.ComboBox cmbnblc2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   19
         ToolTipText     =   " ⁄œ«œ ﬁ—’ 2"
         Top             =   2280
         Width           =   975
      End
      Begin VB.ComboBox cmbsblc2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   18
         ToolTipText     =   "”«Ì“ ﬁ—’ 2"
         Top             =   1800
         Width           =   855
      End
      Begin VB.OptionButton opt 
         Caption         =   "Duble Size"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   3840
         TabIndex        =   21
         ToolTipText     =   "Arrester With Two Size Block"
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton opt 
         Caption         =   "Single Size"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   600
         TabIndex        =   20
         ToolTipText     =   "Arrester With Single Size Block"
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.ComboBox cmbnblc1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   17
         ToolTipText     =   " ⁄œ«œ ﬁ—’ 1"
         Top             =   2400
         Width           =   975
      End
      Begin VB.ComboBox cmbsblc1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   16
         ToolTipText     =   "”«Ì“ ﬁ—’ 1"
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblinom 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   64
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblt2 
         Caption         =   "Size 2 Information............"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   63
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblt1 
         Caption         =   "Size 1 Information ............."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   62
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblblctype 
         Caption         =   "Type Of Blocks "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   53
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblnblc2 
         Caption         =   "Block no."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   52
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label lblnblc1 
         Caption         =   "Block no."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   51
         ToolTipText     =   " "
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label lblsblc2 
         Caption         =   "Block size"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3120
         TabIndex        =   50
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lblsblc1 
         Caption         =   "Block Size."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   49
         ToolTipText     =   "    "
         Top             =   1800
         Width           =   1215
      End
   End
   Begin VB.Frame framtprmt 
      Caption         =   "Technical Parameters "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   360
      TabIndex        =   39
      Top             =   4920
      Width           =   6855
      Begin VB.TextBox txturesx 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3960
         TabIndex        =   14
         Text            =   " "
         ToolTipText     =   "Õœ«òÀ— Ê· «é Å”„«‰œ"
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txturesn 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3960
         TabIndex        =   13
         Text            =   " "
         ToolTipText     =   "Õœ«ﬁ· Ê· «é Å”„«‰œ"
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txturefx 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4080
         TabIndex        =   12
         Text            =   " "
         ToolTipText     =   "Õœ«òÀ— Ê· «é „—Ã⁄"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txturefn 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3960
         TabIndex        =   11
         Text            =   " "
         ToolTipText     =   "Õœ«ﬁ· Ê· «é „—Ã⁄"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtPDMax 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   8
         Text            =   "2"
         ToolTipText     =   "Õœ«òÀ—  Œ·ÌÂ Ã“ÌÌ ﬁ«»· ﬁ»Ê·"
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txticaccp 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   10
         Text            =   " "
         ToolTipText     =   "Ã—Ì«‰ ‰‘ Ì ﬁ«»· ﬁ»Ê·"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtiref 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   9
         Text            =   " "
         ToolTipText     =   "Ã—Ì«‰ „—Ã⁄"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtuc 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   7
         Text            =   " "
         ToolTipText     =   "Ê· «é ò«— œ«Ì„"
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblu8 
         Caption         =   "[KV]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   61
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label lblu7 
         Caption         =   "[KV]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   60
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lblu6 
         Caption         =   "[KV]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   59
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblu5 
         Caption         =   "[KV]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   58
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblu2 
         Caption         =   "[pC]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   57
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label lblu4 
         Caption         =   "[mA]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   56
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lblu3 
         Caption         =   "[mA]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   55
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblu1 
         Caption         =   "[KV]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   54
         Top             =   840
         Width           =   375
      End
      Begin VB.Label lbluresn 
         Caption         =   "Ures min ="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   47
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblurefn 
         Caption         =   "Uref min ="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   46
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lbluc 
         Caption         =   "    Uc ="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   45
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblic 
         Caption         =   "Ic(acp)="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lbluresx 
         Caption         =   "Ures max ="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   43
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label lbliref 
         Caption         =   "   I ref ="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   42
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblurefx 
         Caption         =   "Uref max ="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   41
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblurt 
         Caption         =   "Max. PD"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   40
         Top             =   2280
         Width           =   615
      End
   End
   Begin VB.Frame framarrrate 
      Caption         =   "Arrester Coding and Order Data"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   240
      TabIndex        =   35
      Top             =   360
      Width           =   14175
      Begin VB.Frame framCNT 
         Caption         =   "Surge Counter info."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   8280
         TabIndex        =   75
         Top             =   1680
         Width           =   5055
         Begin VB.TextBox TXTCntModel 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   3000
            TabIndex        =   30
            Text            =   " "
            ToolTipText     =   "„œ· ò«‰ — „ò«‰ÌòÌ"
            Top             =   1200
            Width           =   1815
         End
         Begin VB.TextBox TXTCntMan 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   960
            TabIndex        =   28
            Text            =   " "
            ToolTipText     =   "‰«„ ”«“‰œÂ ò«‰ — „ò«‰ÌòÌ"
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox TXTMeterModel 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   3000
            TabIndex        =   29
            Text            =   " "
            ToolTipText     =   "„œ· „Ì·Ì Ê·  „ —"
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox TXTMeterMan 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   960
            TabIndex        =   27
            Text            =   " "
            ToolTipText     =   "‰«„ ”«“‰œÂ „Ì·Ì Ê·  „ —"
            Top             =   600
            Width           =   1335
         End
         Begin VB.ComboBox CmBProtection 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3000
            Style           =   2  'Dropdown List
            TabIndex        =   31
            ToolTipText     =   "‰Ê⁄ ”Ì” „ Õ›«Ÿ  »—œ œ— „ﬁ«»· Ã—Ì«‰ »«·«"
            Top             =   1920
            Width           =   1455
         End
         Begin VB.Label LblModel 
            AutoSize        =   -1  'True
            Caption         =   "Model"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   3720
            TabIndex        =   80
            Top             =   240
            Width           =   510
         End
         Begin VB.Label LBLMan 
            AutoSize        =   -1  'True
            Caption         =   "Manufacturer"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1080
            TabIndex        =   79
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label Lblprotection 
            AutoSize        =   -1  'True
            Caption         =   "High Current Protection Type"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   240
            TabIndex        =   78
            Top             =   2040
            Width           =   2445
         End
         Begin VB.Label LBLMeter 
            AutoSize        =   -1  'True
            Caption         =   "Meter"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   240
            TabIndex        =   77
            Top             =   720
            Width           =   495
         End
         Begin VB.Label LBLCnt 
            AutoSize        =   -1  'True
            Caption         =   "Counter"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   240
            TabIndex        =   76
            Top             =   1440
            Width           =   675
         End
      End
      Begin VB.Frame framHV 
         Caption         =   "HV Arrester Info."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   3480
         TabIndex        =   72
         Top             =   2040
         Width           =   3615
         Begin VB.ComboBox CmdNofloor 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   25
            ToolTipText     =   " ⁄œ«œ ÿ»ﬁ«  »—ﬁêÌ—"
            Top             =   240
            Width           =   1455
         End
         Begin VB.ComboBox cmbDiaph 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   26
            ToolTipText     =   "‰Ê⁄ œÌ«›—«ê„"
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label LblDiaph 
            AutoSize        =   -1  'True
            Caption         =   "Diaphragm"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   360
            TabIndex        =   74
            Top             =   840
            Width           =   930
         End
         Begin VB.Label lblnofloor 
            AutoSize        =   -1  'True
            Caption         =   " Arrester Units"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   73
            Top             =   360
            Width           =   1245
         End
      End
      Begin VB.ComboBox CmbArrModel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1830
         Width           =   1395
      End
      Begin VB.TextBox txtcustomer 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   9120
         TabIndex        =   6
         Text            =   " "
         ToolTipText     =   "‰«„ „‘ —Ì"
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox TXTARRNO 
         Alignment       =   2  'Center
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
         Left            =   5760
         TabIndex        =   4
         Text            =   " "
         ToolTipText     =   " ⁄œ«œ »—ﬁêÌ— œ— »—‰«„Â"
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtorderno 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7080
         TabIndex        =   5
         Text            =   " "
         ToolTipText     =   "‘„«—Â »—‰«„Â ° »—«Ì «’·«Õ  ⁄œ«œ »—ﬁêÌ—Â«Ì Ìò »—‰«„Â »œ·Ì· ò”—Ì ò«Å° ‘„«—Â »—‰«„Â «’·Ì —« »« Ìò ’›— œ— ¬Œ— À»  ò‰Ìœ "
         Top             =   1200
         Width           =   975
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   135
         Left            =   480
         TabIndex        =   66
         Top             =   1320
         Width           =   15
      End
      Begin VB.TextBox TxtCode 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3240
         TabIndex        =   3
         Text            =   " "
         ToolTipText     =   "»Œ‘  ò„Ì·Ì òœÌ‰ê »—ﬁêÌ—"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtUr1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1800
         TabIndex        =   2
         Text            =   " "
         ToolTipText     =   "Ê· «é ‰«„Ì »—ﬁêÌ—"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.ComboBox cmbarr1 
         DataField       =   "Arr_Type"
         DataSource      =   "Datatype"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   285
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "„œ· »—ﬁêÌ—"
         Top             =   1185
         Width           =   1335
      End
      Begin VB.Label lblcustomer 
         Alignment       =   2  'Center
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9120
         TabIndex        =   70
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lbltool2 
         Caption         =   "Selected Serial Number: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   69
         Top             =   2760
         Width           =   5535
      End
      Begin VB.Label LBLARRNO 
         Alignment       =   2  'Center
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5760
         TabIndex        =   68
         Top             =   720
         Width           =   885
      End
      Begin VB.Label lblorderno 
         Alignment       =   2  'Center
         Caption         =   "Order No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6960
         TabIndex        =   67
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblcode 
         Alignment       =   2  'Center
         Caption         =   "Complementary Coding"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   65
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label lblselarr 
         Caption         =   "Arrester Coding: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   2400
         Width           =   5415
      End
      Begin VB.Label lblarr2 
         Alignment       =   2  'Center
         Caption         =   "Rated Voltage"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   37
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lblarr1 
         Alignment       =   2  'Center
         Caption         =   "Product Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   720
         Width           =   1335
      End
   End
   Begin MSAdodcLib.Adodc datavac 
      Height          =   330
      Left            =   2280
      Top             =   6840
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "dataarrtype"
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
   Begin MSAdodcLib.Adodc DataGTIN 
      Height          =   330
      Left            =   9720
      Top             =   8760
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
      Caption         =   "dataarrtype"
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
   Begin VB.Label LblMonth 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "„«Â  Ê·Ìœ"
      BeginProperty Font 
         Name            =   "B Zar"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   83
      Top             =   8760
      Width           =   885
   End
   Begin VB.Label LblPrdcCode 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Product Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2955
      TabIndex        =   82
      Top             =   8280
      Width           =   1185
   End
   Begin VB.Label LblGLN 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Company GLN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   300
      TabIndex        =   81
      Top             =   8280
      Width           =   1215
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private hei, wid, maxsn As Double
Public dbadres As String
Public inoms1, inoms2, i, arrstate As Integer
Private strsn As String
Private sn1, sn2 As Double
Private Tabletype() As String
Dim SerialCode() As Integer
Dim ArrTypeCode() As String
Dim GTINCode As Integer
Dim AllGTINCodes() As Double

Private allow, SerialIsOK As Boolean
Dim rtv, j, cnt, C As Integer
Dim maxarrnum, maxarrnum1 As Double
Private Sub ApplySerialNumber()
SerialIsOK = False

If allow = True Then
  With frmmain.dataarrtype
      .RecordSource = "production"
      .Refresh
      If .Recordset.RecordCount > 0 Then
          .Recordset.MoveFirst
          .Recordset.Find "OrderNo=" + frmmain.txtorderno, , adSearchForward, 1
          If .Recordset.EOF = False Then
              MsgBox "Enterred order number is repetitive. You are not allowed to register repetitive order number?", vbYesNo, "Œÿ«Ì ‘„«—Â »—‰«„Â  ò—«—Ì"
              allow = False
              txtorderno.SetFocus
          End If
      End If
  End With
End If

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\



If allow = True And Tabletype(cmbarr1.ListIndex) <> "CNT" Then  ' Just for arrester

   If cmbsblc2.Visible = True Then
        mdl1.blockinf = cmbnblc1.Text + Mid(cmbsblc1.Text, 1, 2) + Right(cmbsblc1.Text, 2) + cmbnblc2.Text + Mid(cmbsblc2.Text, 1, 2) + Right(cmbsblc2.Text, 2)
    Else
         mdl1.blockinf = cmbnblc1.Text + Mid(cmbsblc1.Text, 1, 2) + Right(cmbsblc1.Text, 2) + "00000"
    End If
    
    mdl1.blockname = cmbblctype.Text
    mdl1.urefmin = Val(frmmain.txturefn.Text)
    mdl1.urefmax = Val(frmmain.txturefx.Text)
    mdl1.uresmin = Val(frmmain.txturesn.Text)
    mdl1.uresmax = Val(frmmain.txturesx.Text)
    mdl1.ucoun = Val(frmmain.txtuc.Text)
    mdl1.iref = Val(frmmain.txtiref.Text)
    mdl1.ilaccp = Val(frmmain.txticaccp.Text)
    mdl1.urt = Val(frmmain.txtUr1.Text)
    mdl1.unitnumber = Val(frmmain.CmdNofloor.Text)
    mdl1.ordernum = Val(txtorderno.Text)
    mdl1.Acpt_PD = Val(txtPDMax.Text)
    
    mdl1.sncode = SerialCode(cmbarr1.ListIndex)
    mdl1.ArresterType = cmbarr1.Text + "-" + txtUr1.Text + "-" + UCase(TxtCode.Text)
    
    mdl1.CompleteTavanirCode = (mdl1.GLNCode * 10 ^ 10) + (AllGTINCodes(CmbPrdcCode.ListIndex) * 10 ^ 4) + Format(mdl1.year, "##") * 100 + Format(CmbMonth.ListIndex + 1, "##")
    
    If Tabletype(cmbarr1.ListIndex) = "SUBST" Then '  Station Class Arrester Coding
        Data1.RecordSource = "SUBST"
        
        Data1.Refresh
        isHVArrester = True
        '//////////////////////////////////////find last serial number///////
         With Data1.Recordset
         
         ' Serial Number Coding of Station Class =  Year ## + Arresterv Code # + Unit Number # + Exclusive Number ###
         
            strsn = Str(mdl1.year) + Right(Str(mdl1.sncode), 1) ' the first 3 digits of serial Number
            maxsn = Val(strsn + "10109") ' The first serial Nymber is 10110. i.e: 1 (Unit 1) + 0110 (Exclisive Serial
            sn1 = Val(Right(Str(maxsn), 4)) ' the last exclusive serial Number
            
            If .RecordCount = 0 Then
               .addnew
               !Serial_Num = maxsn
               .Update
            End If
           .MoveFirst
           
           i = .RecordCount
           If .RecordCount > 0 Then
                Do While .EOF = False  ' Finding the last registered serial Number
                    If (.Fields("Serial_Num") - (.Fields("Serial_Num") Mod 100000)) / 100000 = Val(strsn) Then ' is the same Arrester Type?
                        If Val(Right(Str(.Fields("Serial_Num")), 4)) > Val(Right(Str(maxsn), 4)) Then ' is biger serial Number?
                            maxsn = .Fields("Serial_Num") ' in the last step, maxsn will be tha last registered Serial No. of arrester type
                            sn1 = Val(Right(Str(maxsn), 4)) ' the last exclusive serial Number
                        End If
                    End If
                     .MoveNext
                Loop
                maxsn = Val(strsn + Str(Right(Str(10000 + sn1), 5)))
                
           End If
        End With
       
         If Val(TXTARRNO.Text) > 0 Then
         
            maxarrnum = Val(TXTARRNO.Text)
            For i = 1 To Val(TXTARRNO.Text)
               If sn1 + i = 99999 Then
                    MsgBox "You can not allocate serial number for all entered quantity of arresters." + Chr(13) + " Serial number will be overflowed." + Chr(13) + "You can only add " + Str(i - 1) + " arrester(s).", vbExclamation, "Attention, Serial Number Overflow"
                    maxarrnum = i - 1
                    Exit For
              End If
            Next i
            
            maxarrnum1 = maxarrnum * Val(CmdNofloor.Text)
            ReDim mdl1.newsrn(1 To maxarrnum1) As Double
            mdl1.newarrnum = maxarrnum1
            
            cnt = 1
            For i = 1 To maxarrnum
                C = 0
                For j = 1 To Val(CmdNofloor.Text)
                    mdl1.newsrn(cnt) = (maxsn + i) + (C * 10000)
                    C = C + 1
                    cnt = cnt + 1
                Next j
            Next i
    
            frmmain.MousePointer = 0
            lbltool2.Caption = "New Serial Number:  From :" + Str(mdl1.newsrn(1)) + "  To: " + Str(mdl1.newsrn(mdl1.newarrnum))
            frmconfirm.lbl1.Caption = "Arrester Type :                 " + mdl1.ArresterType
            frmconfirm.lbl2.Caption = "Serial Number: From :" + Str(mdl1.newsrn(1))
            frmconfirm.lbl3.Caption = "                          TO :" + Str(mdl1.newsrn(mdl1.newarrnum))
            SerialIsOK = True
        End If
         
    
    Else            '   Distribution Class Arrester Coding
    
     '  Serial Number Coding of Distribution Class =  Year ## + Arresterv Code # +  Exclusive Number (5 digit)  ##### as following
     ' For ETM & ETI the last 5 digit for Exclusive Numberis is as following :
     
     'if ETM-25-H10B or "ETI-25-R10" : SR = [00110 up to 29999] total number is 29890 pieces
     'if ETM1-25-DH1-B2 or "ETI1-25-DH1-R2" : SR = [30110 up to 69999] total number is 39890 pieces
     
     'if ETM-36-K10B or "ETI-36-R10" : SR = [70110 up to 79999] total number is  9890 pieces
     'if ETM1-36-DH1-B3 or "ETI1-36-DH1-R3" : SR = [80110 up to 89999] total number is  9890 pieces
    
     'if Else type  : SR = [90110 up to 99999] total number is 9890 pieces
     
     'The first Serial Number Is 100 [YA ALI]
        isHVArrester = False
        Data1.RecordSource = "DIST"
        Data1.Refresh
        mdl1.ArrTableName = "DIST"
         '//////////////////////////////////////find last serial number///////
        
        With Data1.Recordset
            strsn = Str(mdl1.year) + Right(Str(mdl1.sncode), 1) ' the first 3 digits of serial Number
            
            If mdl1.ArresterType = "ETM-25-H10B" Or mdl1.ArresterType = "ETI-25-R10" Then
                maxsn = Val(strsn + "00109") '  The first Serial Number for 25 KV is ###00110 [YA ALI]
            ElseIf mdl1.ArresterType = "ETM1-25-DH1-B2" Or mdl1.ArresterType = "ETI1-25-DH1-R2" Then
                maxsn = Val(strsn + "30109") '  The first Serial Number for 25 KV is ###20110 [YA ALI]
                
            ElseIf mdl1.ArresterType = "ETM-36-K10B" Or mdl1.ArresterType = "ETI-36-R10" Then
                maxsn = Val(strsn + "70109") '  The first Serial Number for 36 KV is ###70110 [YA ALI]
            
            ElseIf mdl1.ArresterType = "ETM1-36-DH1-B3" Or mdl1.ArresterType = "ETI1-36-DH1-R3" Then
                maxsn = Val(strsn + "80109") '  The first Serial Number for 36 KV is ###80110 [YA ALI]
            
            Else
                maxsn = Val(strsn + "90109") '  The first Serial Number for others is ###90110 [YA ALI]
            End If
            
            sn1 = Val(Right(Str(maxsn), 5)) ' the last exclusive serial Number
            If .RecordCount = 0 Then
               .addnew
               !Serial_Num = maxsn
               .Update
            End If
           .MoveFirst
           
           i = .RecordCount
           If .RecordCount > 0 Then
                Do While .EOF = False  ' Finding the last registered serial Number
                
                    If (.Fields("Serial_Num") - (.Fields("Serial_Num") Mod 100000)) / 100000 = Val(strsn) Then ' is the same Arrester Type?
                        
                        
                        If mdl1.ArresterType = "ETM-25-H10B" Or mdl1.ArresterType = "ETI-25-R10" Then  ' SN =  [00110 up to 30109]
                        
                            If Val(Right(Str(.Fields("Serial_Num")), 5)) > Val(Right(Str(maxsn), 5)) And Val(Right(Str(.Fields("Serial_Num")), 5)) >= 109 And Val(Right(Str(.Fields("Serial_Num")), 5)) < 30109 Then ' IMP :20109 ' is biger serial Number?
                                maxsn = .Fields("Serial_Num") ' in the last step, maxsn will be tha last registered Serial No. of arrester type
                                sn1 = Val(Right(Str(maxsn), 5)) ' the last exclusive serial Number
                            End If
                        
                        
                        ElseIf mdl1.ArresterType = "ETM1-25-DH1-B2" Or mdl1.ArresterType = "ETI1-25-DH1-R2" Then ' SN =  [30110 up to 69999]
                        
                            If Val(Right(Str(.Fields("Serial_Num")), 5)) > Val(Right(Str(maxsn), 5)) And Val(Right(Str(.Fields("Serial_Num")), 5)) >= 30109 And Val(Right(Str(.Fields("Serial_Num")), 5)) < 69999 Then ' is biger serial Number?
                                maxsn = .Fields("Serial_Num") ' in the last step, maxsn will be tha last registered Serial No. of arrester type
                                sn1 = Val(Right(Str(maxsn), 5)) ' the last exclusive serial Number
                            End If
                        
                        
                        ElseIf mdl1.ArresterType = "ETM-36-K10B" Or mdl1.ArresterType = "ETI-36-R10" Then  ' SN= [70110 up to 79999]
                            
                            If Val(Right(Str(.Fields("Serial_Num")), 5)) > Val(Right(Str(maxsn), 5)) And Val(Right(Str(.Fields("Serial_Num")), 5)) >= 70109 And Val(Right(Str(.Fields("Serial_Num")), 5)) < 79999 Then ' is biger serial Number?
                                maxsn = .Fields("Serial_Num") ' in the last step, maxsn will be tha last registered Serial No. of arrester type
                                sn1 = Val(Right(Str(maxsn), 5)) ' the last exclusive serial Number
                            End If
                            
                       ElseIf mdl1.ArresterType = "ETM1-36-DH1-B3" Or mdl1.ArresterType = "ETI1-36-DH1-R3" Then ' SN =  [80110 up to 89999]
                        
                            If Val(Right(Str(.Fields("Serial_Num")), 5)) > Val(Right(Str(maxsn), 5)) And Val(Right(Str(.Fields("Serial_Num")), 5)) >= 80109 And Val(Right(Str(.Fields("Serial_Num")), 5)) < 89999 Then ' is biger serial Number?
                                maxsn = .Fields("Serial_Num") ' in the last step, maxsn will be tha last registered Serial No. of arrester type
                                sn1 = Val(Right(Str(maxsn), 5)) ' the last exclusive serial Number
                            End If
                            
                        Else  ' for  Else : SR = [90110 up to 99999]
                            
                            If Val(Right(Str(.Fields("Serial_Num")), 5)) > Val(Right(Str(maxsn), 5)) And Val(Right(Str(.Fields("Serial_Num")), 5)) >= 90109 And Val(Right(Str(.Fields("Serial_Num")), 5)) < 99999 Then ' is biger serial Number?
                                maxsn = .Fields("Serial_Num") ' in the last step, maxsn will be tha last registered Serial No. of arrester type
                                sn1 = Val(Right(Str(maxsn), 5)) ' the last exclusive serial Number
                            End If
                        
                        End If
                        
                    End If
                     .MoveNext
                Loop
           End If
         End With
        
        If Val(TXTARRNO.Text) > 0 Then
            maxarrnum = Val(TXTARRNO.Text)
            For i = 1 To Val(TXTARRNO.Text)
            
                
                If mdl1.ArresterType = "ETM-25-H10B" Or mdl1.ArresterType = "ETI-25-R10" Then
                    If sn1 + i = 30109 Then  ' IMP :20109
                        MsgBox "You can not allocate serial number for all entered quantity of arresters." + Chr(13) + " Serial number will be overflowed." + Chr(13) + "You can only add " + Str(i - 1) + " arrester(s).", vbExclamation, "Attention, Serial Number Overflow"
                        maxarrnum = i - 1
                        Exit For
                    End If
                    
               ElseIf mdl1.ArresterType = "ETM1-25-DH1-B2" Or mdl1.ArresterType = "ETI1-25-DH1-R2" Then
                    If sn1 + i = 69999 Then
                        MsgBox "You can not allocate serial number for all entered quantity of arresters." + Chr(13) + " Serial number will be overflowed." + Chr(13) + "You can only add " + Str(i - 1) + " arrester(s).", vbExclamation, "Attention, Serial Number Overflow"
                        maxarrnum = i - 1
                        Exit For
                    End If
                
                
                ElseIf mdl1.ArresterType = "ETM-36-K10B" Or mdl1.ArresterType = "ETI-36-R10" Then
                    If sn1 + i = 79999 Then
                        MsgBox "You can not allocate serial number for all entered quantity of arresters." + Chr(13) + " Serial number will be overflowed." + Chr(13) + "You can only add " + Str(i - 1) + " arrester(s).", vbExclamation, "Attention, Serial Number Overflow"
                        maxarrnum = i - 1
                        Exit For
                    End If
                    
                 ElseIf mdl1.ArresterType = "ETM1-36-DH1-B3" Or mdl1.ArresterType = "ETI1-36-DH1-R3" Then
                    If sn1 + i = 89999 Then
                        MsgBox "You can not allocate serial number for all entered quantity of arresters." + Chr(13) + " Serial number will be overflowed." + Chr(13) + "You can only add " + Str(i - 1) + " arrester(s).", vbExclamation, "Attention, Serial Number Overflow"
                        maxarrnum = i - 1
                        Exit For
                    End If
                    
                    
                Else
                    If sn1 + i = 99999 Then
                        MsgBox "You can not allocate serial number for all entered quantity of arresters." + Chr(13) + " Serial number will be overflowed." + Chr(13) + "You can only add " + Str(i - 1) + " arrester(s).", vbExclamation, "Attention, Serial Number Overflow"
                        maxarrnum = i - 1
                        Exit For
                    End If
                
                End If
                
            Next i
            
            If maxarrnum > 0 Then
                        
                ReDim mdl1.newsrn(1 To maxarrnum) As Double
                mdl1.newarrnum = maxarrnum
                For i = 1 To maxarrnum
                    mdl1.newsrn(i) = maxsn + i
                Next i
                
                frmmain.MousePointer = 0
                lbltool2.Caption = "New Serial Number: From :" + Str(mdl1.newsrn(1)) + " To: " + Str(mdl1.newsrn(i - 1))
                frmconfirm.lbl1.Caption = "Arrester Type : " + mdl1.ArresterType
                frmconfirm.lbl2.Caption = "Serial Number: From :" + Str(mdl1.newsrn(1))
                frmconfirm.lbl3.Caption = "                         TO  :" + Str(mdl1.newsrn(Val(TXTARRNO.Text)))
                SerialIsOK = True
            End If
        End If
        
   
        
    End If
    
'//////////////////////////////////////////////////////////////////////////////////////////////////////
'                                          Serial Number for Counter


ElseIf allow = True And Tabletype(cmbarr1.ListIndex) = "CNT" Then  '  for Counter

    mdl1.ordernum = Val(txtorderno.Text)
    mdl1.sncode = SerialCode(cmbarr1.ListIndex)
    mdl1.ArresterType = cmbarr1.Text
    Data1.RecordSource = "CNT"
    Data1.Refresh
        
    '//////////////////////////////////////find last serial number///////
    With Data1.Recordset
         
         ' Serial Number Coding of Counter =  Year ## + Product Code + Exclusive Number #####
         
            strsn = Str(mdl1.year) + Right(Str(mdl1.sncode), 1) ' the first 3 digits of serial Number
            maxsn = Val(strsn + "10109") ' The first serial Nymber is 10110. i.e: 1 (Unit 1) + 0110 (Exclisive Serial
            sn1 = Val(Right(Str(maxsn), 4)) ' the last exclusive serial Number
            
            If .RecordCount = 0 Then
               .addnew
               !Serial_Num = maxsn
               .Update
            End If
           .MoveFirst
           
           i = .RecordCount
           If .RecordCount > 0 Then
                Do While .EOF = False  ' Finding the last registered serial Number
                    If (.Fields("Serial_Num") - (.Fields("Serial_Num") Mod 100000)) / 100000 = Val(strsn) Then ' is the same Arrester Type?
                        If Val(Right(Str(.Fields("Serial_Num")), 4)) > Val(Right(Str(maxsn), 4)) Then ' is biger serial Number?
                            maxsn = .Fields("Serial_Num") ' in the last step, maxsn will be tha last registered Serial No. of arrester type
                            sn1 = Val(Right(Str(maxsn), 4)) ' the last exclusive serial Number
                        End If
                    End If
                     .MoveNext
                Loop
                maxsn = Val(strsn + Str(Right(Str(10000 + sn1), 5)))
                
           End If
        End With
       
         If Val(TXTARRNO.Text) > 0 Then
         
            maxarrnum = Val(TXTARRNO.Text)
            For i = 1 To Val(TXTARRNO.Text)
               If sn1 + i = 99999 Then
                    MsgBox "You can not allocate serial number for all entered quantity of arresters." + Chr(13) + " Serial number will be overflowed." + Chr(13) + "You can only add " + Str(i - 1) + " arrester(s).", vbExclamation, "Attention, Serial Number Overflow"
                    maxarrnum = i - 1
                    Exit For
              End If
            Next i
            
            maxarrnum1 = maxarrnum * Val(CmdNofloor.Text)
            ReDim mdl1.newsrn(1 To maxarrnum1) As Double
            mdl1.newarrnum = maxarrnum1
            
            cnt = 1
            For i = 1 To maxarrnum
                C = 0
                For j = 1 To Val(CmdNofloor.Text)
                    mdl1.newsrn(cnt) = (maxsn + i) + (C * 10000)
                    C = C + 1
                    cnt = cnt + 1
                Next j
            Next i
    
            frmmain.MousePointer = 0
            lbltool2.Caption = "New Serial Number:  From :" + Str(mdl1.newsrn(1)) + "  To: " + Str(mdl1.newsrn(mdl1.newarrnum))
            frmconfirm.lbl1.Caption = "Arrester Type :                 " + mdl1.ArresterType
            frmconfirm.lbl2.Caption = "Serial Number: From :" + Str(mdl1.newsrn(1))
            frmconfirm.lbl3.Caption = "                          TO :" + Str(mdl1.newsrn(mdl1.newarrnum))
            SerialIsOK = True
        End If
  
End If


frmmain.MousePointer = 0



End Sub


Private Function CheckInputData() As Boolean


allow = True

If Tabletype(cmbarr1.ListIndex) = "CNT" Then  ' Counter

    
    If (TXTARRNO.Text = "" Or TXTARRNO.Text = " " Or Val(TXTARRNO.Text) = 0) And allow = True Then
        allow = False
        rtv = MsgBox("Enter Quantity of Counters for Current Order", vbOKOnly + vbExclamation, "ERROR")
        TXTARRNO.SetFocus
    End If
    
    If (txtorderno.Text = "" Or txtorderno.Text = " " Or Val(txtorderno.Text) = 0) And allow = True Then
        allow = False
        rtv = MsgBox("Enter Order Number Of Current Batch", vbOKOnly + vbExclamation, "ERROR")
       txtorderno.SetFocus
    End If
    If (txtcustomer.Text = " " Or txtcustomer.Text = "") And allow = True Then
        allow = False
        rtv = MsgBox("Enter Customer Name", vbOKOnly + vbExclamation, "ERROR")
        txtcustomer.Text = ""
        txtcustomer.SetFocus
    End If
  If (TXTMeterMan.Text = " " Or TXTMeterMan.Text = "") And allow = True Then
        allow = False
        rtv = MsgBox("Enter name of meter manufacturer", vbOKOnly + vbExclamation, "ERROR")
        TXTMeterMan.Text = ""
        TXTMeterMan.SetFocus
    End If

   If (TXTMeterModel.Text = " " Or TXTMeterModel.Text = "") And allow = True Then
        allow = False
        rtv = MsgBox("Enter model of meter.", vbOKOnly + vbExclamation, "ERROR")
        TXTMeterModel.Text = ""
        TXTMeterModel.SetFocus
    End If
    
     If (TXTCntMan.Text = " " Or TXTCntMan.Text = "") And allow = True Then
        allow = False
        rtv = MsgBox("Enter name of counter manufacturer", vbOKOnly + vbExclamation, "ERROR")
        TXTCntMan.Text = ""
        TXTCntMan.SetFocus
    End If

   If (TXTCntModel.Text = " " Or TXTCntModel.Text = "") And allow = True Then
        allow = False
        rtv = MsgBox("Enter model of counter.", vbOKOnly + vbExclamation, "ERROR")
        TXTCntModel.Text = ""
        TXTCntModel.SetFocus
    End If
    

Else  ' Arrester

  If (txtUr1.Text = " " Or txtUr1.Text = "" Or Val(txtUr1.Text) = 0 Or Val(txtUr1.Text) > 1500) And allow = True Then
        CheckInputData = False
        rtv = MsgBox("Insert Currect Value For Rated Voltage", vbOKOnly + vbExclamation, "ERROR")
        txtUr1.Text = ""
        txtUr1.SetFocus
    End If
    If (TxtCode.Text = " " Or TxtCode.Text = "") And allow = True Then
        allow = False
        rtv = MsgBox("Insert Currect Value For Arrester Type", vbOKOnly + vbExclamation, "ERROR")
        TxtCode.Text = ""
        TxtCode.SetFocus
    End If
    
    If (txtcustomer.Text = " " Or txtcustomer.Text = "") And allow = True Then
        allow = False
        rtv = MsgBox("Insert Customer Name", vbOKOnly + vbExclamation, "ERROR")
        txtcustomer.Text = ""
        txtcustomer.SetFocus
    End If
    
    If (txtuc.Text = " " Or txtuc.Text = "" Or Val(txtuc.Text) = 0 Or Val(txtuc.Text) > 1000) And allow = True Then
        allow = False
        rtv = MsgBox("Insert Currect Value For Uc", vbOKOnly + vbExclamation, "ERROR")
        txtuc.Text = ""
        txtuc.SetFocus
    End If
    If (txtiref.Text = " " Or txtiref.Text = "" Or Val(txtiref.Text) = 0 Or Val(txtiref.Text) > 1000) And allow = True Then
        allow = False
        rtv = MsgBox("Insert Currect Value For Reference Current", vbOKOnly + vbExclamation, "ERROR")
        txtiref.Text = ""
        txtiref.SetFocus
    End If
    If (txticaccp.Text = " " Or txticaccp.Text = "" Or Val(txticaccp.Text) = 0 Or Val(txticaccp.Text) > 1000) And allow = True Then
        allow = False
        rtv = MsgBox("Insert Currect Value For Acceptable Leakage Current", vbOKOnly + vbExclamation, "ERROR")
        txticaccp.Text = ""
        txticaccp.SetFocus
    End If
    If (txtPDMax.Text = " " Or txtPDMax.Text = "" Or Val(txtPDMax.Text) <= 0 Or Val(txtPDMax.Text) > 10) And allow = True Then
        allow = False
        rtv = MsgBox("Insert Currect Value For Max. PD Level [2-10 pC]", vbOKOnly + vbExclamation, "ERROR")
        txtPDMax.Text = ""
        txtPDMax.SetFocus
    End If
    '/////////////////////////////////////////////////////////////////
    If (txturefn.Text = " " Or txturefn.Text = "" Or Val(txturefn.Text) = 0 Or Val(txturefn.Text) > 1200) And allow = True Then
        allow = False
        rtv = MsgBox("Insert Currect For Minimum Reference Voltage", vbOKOnly + vbExclamation, "ERROR")
        txturefn.Text = ""
        txturefn.SetFocus
    End If
    If (Val(txturefn.Text) >= Val(txturefx.Text)) And allow = True Then
        allow = False
        rtv = MsgBox("Insert Currect Value For Minimum Reference Voltage and Maximum Reference Voltage", vbOKOnly + vbExclamation, "ERROR")
        txturefn.Text = ""
        txturefx.Text = ""
        txturefn.SetFocus
    End If
    If (txturefx.Text = " " Or txturefx.Text = "" Or Val(txturefx.Text) = 0 Or Val(txturefn.Text) > 1200) And allow = True Then
        allow = False
        rtv = MsgBox("Insert Value For Maximum Reference Voltage", vbOKOnly + vbExclamation, "ERROR")
        txturefx.Text = ""
        txturefx.SetFocus
    End If
    '////////////////////////////////////////////////////////////
    If (txturesn.Text = " " Or txturesn.Text = "" Or Val(txturesn.Text) = 0 Or Val(txturesn.Text) > 1200) And allow = True Then
        allow = False
        rtv = MsgBox("Insert Currect Value For Minimum Residual Voltage", vbOKOnly + vbExclamation, "ERROR")
        txturesn.Text = ""
        txturesn.SetFocus
    End If
    If (Val(txturesn.Text) >= Val(txturesx.Text)) And allow = True Then
        allow = False
        rtv = MsgBox("Insert Currect Value For Minimum Residual Voltage and Maximum Residual Voltage", vbOKOnly + vbExclamation, "ERROR")
        txturesn.Text = ""
        txturesx.Text = ""
        txturesn.SetFocus
    End If
    If (Val(txtuc.Text) >= Val(txtUr1.Text)) And allow = True Then
        allow = False
        rtv = MsgBox("Insert Currect Value For Uc ", vbOKOnly + vbExclamation, "ERROR")
        txtuc.Text = ""
        txtuc.SetFocus
    End If
    If (txturesx.Text = " " Or txturesx.Text = "" Or Val(txturesx.Text) = 0 Or Val(txturesx.Text) > 1200) And allow = True Then
        allow = False
        rtv = MsgBox("Insert Value For Maximum Residual Voltage", vbOKOnly + vbExclamation, "ERROR")
        txturesx.Text = ""
        txturesx.SetFocus
    End If
    If (Val(txturefn.Text) >= Val(txturesn.Text)) And allow = True Then
        allow = False
        rtv = MsgBox("Insert Currect Value For Minimum Reference Voltage and Minimum Residual Voltage", vbOKOnly + vbExclamation, "ERROR")
        txturefn.Text = ""
        txturesn.Text = ""
        txturefn.SetFocus
    End If
    If (Val(txturefx.Text) >= Val(txturesx.Text)) And allow = True Then
        allow = False
        rtv = MsgBox("Insert Currect Value For Maximum Reference Voltage and Maximum residual Voltage", vbOKOnly + vbExclamation, "ERROR")
        txturefx.Text = ""
        txturesx.Text = ""
        txturefx.SetFocus
    End If
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    If (cmbblctype.Text = "" Or cmbblctype.Text = " ") And allow = True Then
        allow = False
        rtv = MsgBox("Select Block Type", vbOKOnly + vbExclamation, "ERROR")
        cmbblctype.SetFocus
    End If
    If (cmbsblc1.Text = "" Or cmbsblc1.Text = " ") And allow = True Then
        allow = False
        rtv = MsgBox("Select Size Of Block Number 1", vbOKOnly + vbExclamation, "ERROR")
        cmbsblc1.SetFocus
    End If
    If (cmbnblc1.Text = "" Or cmbnblc1.Text = " ") And allow = True Then
        allow = False
        rtv = MsgBox("Select Number Of Block Number 1", vbOKOnly + vbExclamation, "ERROR")
        cmbnblc1.SetFocus
    End If
    If (cmbsblc2.Text = "" Or cmbsblc2.Text = " ") And allow = True And cmbsblc2.Visible = True Then
        allow = False
        rtv = MsgBox("Select Size Of Block Number 2", vbOKOnly + vbExclamation, "ERROR")
        cmbsblc2.SetFocus
    End If
    If (cmbnblc2.Text = "" Or cmbnblc2.Text = " ") And allow = True And cmbsblc2.Visible = True Then
        allow = False
        rtv = MsgBox("Select Number Of Block Number 2", vbOKOnly + vbExclamation, "ERROR")
        cmbnblc2.SetFocus
    End If
    If (TXTARRNO.Text = "" Or TXTARRNO.Text = " " Or Val(TXTARRNO.Text) = 0) And allow = True Then
        allow = False
        rtv = MsgBox("Insert Quantity of Arresters for Current Order", vbOKOnly + vbExclamation, "ERROR")
        TXTARRNO.SetFocus
    End If
    If (txtorderno.Text = "" Or txtorderno.Text = " " Or Val(txtorderno.Text) = 0) And allow = True Then
        allow = False
        rtv = MsgBox("Insert Order Number Of Current Batch", vbOKOnly + vbExclamation, "ERROR")
       txtorderno.SetFocus
    End If


End If




CheckInputData = allow

End Function


Private Sub cmbarr1_Change()

Select Case Tabletype(cmbarr1.ListIndex)

Case "SUBST"    ' Multy unit arrester
   CmdNofloor.ListIndex = 0
    CmdNofloor.Enabled = True
    cmbDiaph.Enabled = True
    framHV.Enabled = True
    framCNT.Enabled = False
    framtprmt.Enabled = True
    framblock.Enabled = True
    txtUr1.Enabled = True
    
Case "DIST"                  ' Single  distribution

    CmdNofloor.ListIndex = 0
    CmdNofloor.Enabled = False
    cmbDiaph.Enabled = False
    framHV.Enabled = False
    framCNT.Enabled = False
    framtprmt.Enabled = True
    framblock.Enabled = True
    txtUr1.Enabled = True
Case "CNT" '  surge counter

  framCNT.Enabled = True
  TXTMeterMan.Text = ""
  TXTMeterModel.Text = ""
  TXTCntMan.Text = ""
  TXTCntModel.Text = ""
  framtprmt.Enabled = False
  framblock.Enabled = False
  framHV.Enabled = False
  txtUr1.Enabled = False
  
End Select


With CmbArrModel

    If cmbarr1.Text = "ETM" Then
        .Enabled = True
        .Clear
        .AddItem "ETM-25-H10B", 0
        .AddItem "ETM-36-K10B", 1
        .AddItem "Other types", 2
        .ListIndex = 0
        txtUr1.Text = ""
        TxtCode.Text = ""
        
    ElseIf cmbarr1.Text = "ETM1" Then
        .Enabled = True
        .Clear
        .AddItem "ETM1-25-DH1-B2", 0
        .AddItem "ETM1-36-DH1-B3", 1
        .AddItem "Other types", 2
        .ListIndex = 0
        txtUr1.Text = ""
        TxtCode.Text = ""
        
    ElseIf cmbarr1.Text = "ETI" Then
        .Enabled = True
        .Clear
        .AddItem "ETI-25-R10", 0
        .AddItem "ETI-36-R10", 1
        .AddItem "Other types", 2
        .ListIndex = 0
        txtUr1.Text = ""
        TxtCode.Text = ""
    ElseIf cmbarr1.Text = "ETI1" Then
        .Enabled = True
        .Clear
        .AddItem "ETI1-25-DH1-R2", 0
        .AddItem "ETI1-36-DH1-R3", 1
        .AddItem "Other types", 2
        .ListIndex = 0
        txtUr1.Text = ""
        TxtCode.Text = ""
        
    Else
        CmbArrModel.ListIndex = 2
        .Enabled = False
        txtUr1.Text = ""
        TxtCode.Text = ""
    End If
    
End With

 Select Case CmbArrModel.Text
    
    Case "ETM-25-H10B"
       
       txtUr1.Text = "25"
       txtuc.Text = Str(Round(Val(txtUr1.Text) * 0.8, 1))
       txtUr1.Enabled = False
       TxtCode.Text = "H10B"
       TxtCode.Enabled = False
         
   Case "ETM1-25-DH1-B2"
       
       txtUr1.Text = "25"
       txtuc.Text = Str(Round(Val(txtUr1.Text) * 0.8, 1))
       txtUr1.Enabled = False
       TxtCode.Text = "DH1-B2"
       TxtCode.Enabled = False
         
   Case "ETM-36-K10B"
   
      txtUr1.Text = "36"
      txtuc.Text = Str(Round(Val(txtUr1.Text) * 0.8, 1))
      txtUr1.Enabled = False
      TxtCode.Text = "K10B"
      TxtCode.Enabled = False
      
   Case "ETM1-36-DH1-B3"
       
       txtUr1.Text = "36"
       txtuc.Text = Str(Round(Val(txtUr1.Text) * 0.8, 1))
       txtUr1.Enabled = False
       TxtCode.Text = "DH1-B3"
       TxtCode.Enabled = False
   
   Case "ETI-25-R10"
       txtUr1.Text = "25"
       txtuc.Text = Str(Round(Val(txtUr1.Text) * 0.8, 1))
       txtUr1.Enabled = False
       TxtCode.Text = "R10"
       TxtCode.Enabled = False
    
   Case "ETI1-25-DH1-R2"
       
       txtUr1.Text = "25"
       txtuc.Text = Str(Round(Val(txtUr1.Text) * 0.8, 1))
       txtUr1.Enabled = False
       TxtCode.Text = "DH1-R2"
       TxtCode.Enabled = False
    
   Case "ETI-36-R10"
      txtUr1.Text = "36"
      txtuc.Text = Str(Round(Val(txtUr1.Text) * 0.8, 1))
      txtUr1.Enabled = False
      TxtCode.Text = "R10"
      TxtCode.Enabled = False
      
  Case "ETI1-36-DH1-R3"
       
       txtUr1.Text = "36"
       txtuc.Text = Str(Round(Val(txtUr1.Text) * 0.8, 1))
       txtUr1.Enabled = False
       TxtCode.Text = "DH1-R3"
       TxtCode.Enabled = False
      
   Case "Other types"
   
      txtUr1.Text = ""
      txtuc.Text = ""
      txtUr1.Enabled = True
      TxtCode.Text = ""
      TxtCode.Enabled = True
    
   End Select
   
If Tabletype(cmbarr1.ListIndex) = "CNT" Then
  lblselarr.Caption = "Arrester Coding:   " + cmbarr1.Text
Else
  lblselarr.Caption = "Arrester Coding:   " + cmbarr1.Text + "-" + txtUr1.Text + "-" + UCase(TxtCode.Text)
End If

End Sub

Private Sub cmbarr1_Click()

Select Case Tabletype(cmbarr1.ListIndex)

Case "SUBST"    ' Multy unit arrester
   CmdNofloor.ListIndex = 0
    CmdNofloor.Enabled = True
    cmbDiaph.Enabled = True
    framHV.Enabled = True
    framCNT.Enabled = False
    framtprmt.Enabled = True
    framblock.Enabled = True
    txtUr1.Enabled = True
Case "DIST"                  ' Single  distribution

    CmdNofloor.ListIndex = 0
    CmdNofloor.Enabled = False
    cmbDiaph.Enabled = False
    framHV.Enabled = False
    framCNT.Enabled = False
    framtprmt.Enabled = True
    framblock.Enabled = True
    txtUr1.Enabled = True
    
Case "CNT" '  surge counter

  framCNT.Enabled = True
  TXTMeterMan.Text = ""
  TXTMeterModel.Text = ""
  TXTCntMan.Text = ""
  TXTCntModel.Text = ""
  framtprmt.Enabled = False
  framblock.Enabled = False
  framHV.Enabled = False
  txtUr1.Enabled = False
  
End Select



With CmbArrModel

    If cmbarr1.Text = "ETM" Then
        .Enabled = True
        .Clear
        .AddItem "ETM-25-H10B", 0
        .AddItem "ETM-36-K10B", 1
        .AddItem "Other types", 2
        .ListIndex = 0
        txtUr1.Text = ""
        TxtCode.Text = ""
        
    ElseIf cmbarr1.Text = "ETM1" Then
        .Enabled = True
        .Clear
        .AddItem "ETM1-25-DH1-B2", 0
        .AddItem "ETM1-36-DH1-B3", 1
        .AddItem "Other types", 2
        .ListIndex = 0
        txtUr1.Text = ""
        TxtCode.Text = ""
        
    ElseIf cmbarr1.Text = "ETI" Then
        .Enabled = True
        .Clear
        .AddItem "ETI-25-R10", 0
        .AddItem "ETI-36-R10", 1
        .AddItem "Other types", 2
        .ListIndex = 0
        txtUr1.Text = ""
        TxtCode.Text = ""
    ElseIf cmbarr1.Text = "ETI1" Then
        .Enabled = True
        .Clear
        .AddItem "ETI1-25-DH1-R2", 0
        .AddItem "ETI1-36-DH1-R3", 1
        .AddItem "Other types", 2
        .ListIndex = 0
        txtUr1.Text = ""
        TxtCode.Text = ""
        
    Else
        CmbArrModel.ListIndex = 2
        .Enabled = False
        txtUr1.Text = ""
        TxtCode.Text = ""
    End If
    
End With

 Select Case CmbArrModel.Text
    
    Case "ETM-25-H10B"
       
       txtUr1.Text = "25"
       txtuc.Text = Str(Round(Val(txtUr1.Text) * 0.8, 1))
       txtUr1.Enabled = False
       TxtCode.Text = "H10B"
       TxtCode.Enabled = False
         
   Case "ETM1-25-DH1-B2"
       
       txtUr1.Text = "25"
       txtuc.Text = Str(Round(Val(txtUr1.Text) * 0.8, 1))
       txtUr1.Enabled = False
       TxtCode.Text = "DH1-B2"
       TxtCode.Enabled = False
         
   Case "ETM-36-K10B"
   
      txtUr1.Text = "36"
      txtuc.Text = Str(Round(Val(txtUr1.Text) * 0.8, 1))
      txtUr1.Enabled = False
      TxtCode.Text = "K10B"
      TxtCode.Enabled = False
      
   Case "ETM1-36-DH1-B3"
       
       txtUr1.Text = "36"
       txtuc.Text = Str(Round(Val(txtUr1.Text) * 0.8, 1))
       txtUr1.Enabled = False
       TxtCode.Text = "DH1-B3"
       TxtCode.Enabled = False
   
   Case "ETI-25-R10"
       txtUr1.Text = "25"
       txtuc.Text = Str(Round(Val(txtUr1.Text) * 0.8, 1))
       txtUr1.Enabled = False
       TxtCode.Text = "R10"
       TxtCode.Enabled = False
    
   Case "ETI1-25-DH1-R2"
       
       txtUr1.Text = "25"
       txtuc.Text = Str(Round(Val(txtUr1.Text) * 0.8, 1))
       txtUr1.Enabled = False
       TxtCode.Text = "DH1-R2"
       TxtCode.Enabled = False
    
   Case "ETI-36-R10"
      txtUr1.Text = "36"
      txtuc.Text = Str(Round(Val(txtUr1.Text) * 0.8, 1))
      txtUr1.Enabled = False
      TxtCode.Text = "R10"
      TxtCode.Enabled = False
      
  Case "ETI1-36-DH1-R3"
       
       txtUr1.Text = "36"
       txtuc.Text = Str(Round(Val(txtUr1.Text) * 0.8, 1))
       txtUr1.Enabled = False
       TxtCode.Text = "DH1-R3"
       TxtCode.Enabled = False
      
   Case "Other types"
   
      txtUr1.Text = ""
      txtuc.Text = ""
      txtUr1.Enabled = True
      TxtCode.Text = ""
      TxtCode.Enabled = True
    
   End Select

If Tabletype(cmbarr1.ListIndex) = "CNT" Then
  lblselarr.Caption = "Arrester Coding:   " + cmbarr1.Text
  txtUr1.Enabled = False
Else
  txtUr1.Enabled = True
  lblselarr.Caption = "Arrester Coding:   " + cmbarr1.Text + "-" + txtUr1.Text + "-" + UCase(TxtCode.Text)
End If


End Sub




Private Sub cmbarr1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   txtUr1.SetFocus
End If


End Sub

Private Sub CmbArrModel_Click()

   Select Case CmbArrModel.Text
    
    Case "ETM-25-H10B"
       
       txtUr1.Text = "25"
       txtuc.Text = Str(Round(Val(txtUr1.Text) * 0.8, 1))
       txtUr1.Enabled = False
       TxtCode.Text = "H10B"
       TxtCode.Enabled = False
         
   Case "ETM1-25-DH1-B2"
       
       txtUr1.Text = "25"
       txtuc.Text = Str(Round(Val(txtUr1.Text) * 0.8, 1))
       txtUr1.Enabled = False
       TxtCode.Text = "DH1-B2"
       TxtCode.Enabled = False
         
   Case "ETM-36-K10B"
   
      txtUr1.Text = "36"
      txtuc.Text = Str(Round(Val(txtUr1.Text) * 0.8, 1))
      txtUr1.Enabled = False
      TxtCode.Text = "K10B"
      TxtCode.Enabled = False
      
   Case "ETM1-36-DH1-B3"
       
       txtUr1.Text = "36"
       txtuc.Text = Str(Round(Val(txtUr1.Text) * 0.8, 1))
       txtUr1.Enabled = False
       TxtCode.Text = "DH1-B3"
       TxtCode.Enabled = False
   
   Case "ETI-25-R10"
       txtUr1.Text = "25"
       txtuc.Text = Str(Round(Val(txtUr1.Text) * 0.8, 1))
       txtUr1.Enabled = False
       TxtCode.Text = "R10"
       TxtCode.Enabled = False
    
   Case "ETI1-25-DH1-R2"
       
       txtUr1.Text = "25"
       txtuc.Text = Str(Round(Val(txtUr1.Text) * 0.8, 1))
       txtUr1.Enabled = False
       TxtCode.Text = "DH1-R2"
       TxtCode.Enabled = False
    
   Case "ETI-36-R10"
      txtUr1.Text = "36"
      txtuc.Text = Str(Round(Val(txtUr1.Text) * 0.8, 1))
      txtUr1.Enabled = False
      TxtCode.Text = "R10"
      TxtCode.Enabled = False
      
  Case "ETI1-36-DH1-R3"
       
       txtUr1.Text = "36"
       txtuc.Text = Str(Round(Val(txtUr1.Text) * 0.8, 1))
       txtUr1.Enabled = False
       TxtCode.Text = "DH1-R3"
       TxtCode.Enabled = False
      
   Case "Other types"
   
      txtUr1.Text = ""
      txtuc.Text = ""
      txtUr1.Enabled = True
      TxtCode.Text = ""
      TxtCode.Enabled = True
    
   End Select
    

End Sub


Private Sub cmbblctype_Click()
Dim s1 As String
Dim i As Integer
cmdapply.Enabled = True
lbltool2.Caption = "New Serial Number: #########" + " To: #########"

'///////////
i = 0
cmbsblc1.Clear
cmbsblc2.Clear
With datablock.Recordset
    .MoveFirst
    Do While .EOF = False
        If .Fields("name") = cmbblctype.Text Then
            cmbsblc1.AddItem Str(.Fields("diameter")) + "*" + Str(.Fields("thickness")), i
            cmbsblc2.AddItem Str(.Fields("diameter")) + "*" + Str(.Fields("thickness")), i
            i = i + 1
        End If
        .MoveNext
    Loop
End With

'//////////

With cmbnblc1
    .Clear
    .AddItem "1", 0
    .AddItem "2", 1
    .AddItem "3", 2
    .AddItem "4", 3
    .AddItem "5", 4
    .AddItem "6", 5
    .AddItem "7", 6
    .AddItem "8", 7
    .AddItem "9", 8
    .AddItem "10", 9
    .AddItem "11", 10
    .AddItem "12", 11
    .AddItem "13", 12
    .AddItem "14", 13
    .AddItem "15", 14
    .AddItem "16", 15
    .AddItem "17", 16
    .AddItem "18", 17
    .AddItem "19", 18
    .AddItem "20", 19
    .AddItem "21", 20
    .AddItem "22", 21
    .AddItem "23", 22
    .AddItem "24", 23
    .AddItem "25", 24
    .AddItem "26", 25
    .AddItem "27", 26
    .AddItem "28", 27
    .AddItem "29", 28
    .AddItem "30", 29
    .AddItem "31", 30
    .AddItem "32", 31
    .AddItem "33", 32
    .AddItem "34", 33
    .AddItem "35", 34
    .AddItem "36", 35
    .AddItem "37", 36
    .AddItem "38", 37
    .AddItem "39", 38
    .AddItem "40", 39
    .AddItem "41", 40
    .AddItem "42", 41
    .AddItem "43", 42
    .AddItem "44", 43
    .AddItem "45", 44
    .AddItem "46", 45
    .AddItem "47", 46
    .AddItem "48", 47
    .AddItem "49", 48
    .AddItem "50", 49
    .AddItem "51", 50
    .AddItem "52", 51
    .AddItem "53", 52
    .AddItem "54", 53
    .AddItem "55", 54
    .AddItem "56", 55
    .AddItem "57", 56
    .AddItem "58", 57
    .AddItem "59", 58
    .AddItem "60", 59
    .AddItem "61", 60
    .AddItem "62", 61
    .AddItem "63", 62
    .AddItem "64", 63
    .AddItem "65", 64
    .AddItem "66", 65
    .AddItem "67", 66
    .AddItem "68", 67
    .AddItem "69", 68
    .AddItem "70", 69
    .AddItem "71", 70
    .AddItem "72", 71
    .AddItem "73", 72
    .AddItem "74", 73
    .AddItem "75", 74
    .AddItem "76", 75
    .AddItem "77", 76
    .AddItem "78", 77
    .AddItem "79", 78
    .AddItem "80", 79
End With

With cmbnblc2
    .Clear
    .AddItem "1", 0
    .AddItem "2", 1
    .AddItem "3", 2
    .AddItem "4", 3
    .AddItem "5", 4
    .AddItem "6", 5
    .AddItem "7", 6
    .AddItem "8", 7
    .AddItem "9", 8
    .AddItem "10", 9
    .AddItem "11", 10
    .AddItem "12", 11
    .AddItem "13", 12
    .AddItem "14", 13
    .AddItem "15", 14
    .AddItem "16", 15
    .AddItem "17", 16
    .AddItem "18", 17
    .AddItem "19", 18
    .AddItem "20", 19
    .AddItem "21", 20
    .AddItem "22", 21
    .AddItem "23", 22
    .AddItem "24", 23
    .AddItem "25", 24
    .AddItem "26", 25
    .AddItem "27", 26
    .AddItem "28", 27
    .AddItem "29", 28
    .AddItem "30", 29
    .AddItem "31", 30
    .AddItem "32", 31
    .AddItem "33", 32
    .AddItem "34", 33
    .AddItem "35", 34
    .AddItem "36", 35
    .AddItem "37", 36
    .AddItem "38", 37
    .AddItem "39", 38
    .AddItem "40", 39
    .AddItem "41", 40
    .AddItem "42", 41
    .AddItem "43", 42
End With

lblinom.Caption = "inom =.....[KA]"
End Sub



Private Sub cmbblctype_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmbsblc1.SetFocus
End If
End Sub


Private Sub cmbnblc1_Click()
cmdapply.Enabled = True
lbltool2.Caption = "New Serial Number: #########" + " To: #########"

End Sub


Private Sub cmbnblc1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdapply.SetFocus
End If
End Sub


Private Sub cmbnblc2_Change()
cmdapply.Enabled = True
lbltool2.Caption = "New Serial Number: #########" + " To: #########"

End Sub

Private Sub cmbsblc1_Click()
Dim s1, s2 As String
Dim rtv As Integer
cmdapply.Enabled = True
lbltool2.Caption = "New Serial Number: #########" + " To: #########"

s1 = cmbblctype.Text + Mid(cmbsblc1.Text, 2, 2) + Right(cmbsblc1.Text, 2)
With datablock.Recordset
    .MoveFirst
    .Find "code='" + s1 + "'", , adSearchForward, 1
    inoms1 = .Fields("inom")
End With
If cmbsblc2.Visible = True And cmbsblc2.Text <> "" Then
    s1 = cmbblctype.Text + Mid(cmbsblc2.Text, 2, 2) + Right(cmbsblc2.Text, 2)
    With datablock.Recordset
        .MoveFirst
       .Find "code='" + s1 + "'", , adSearchForward, 1
       inoms2 = .Fields("inom")
    End With
    'inoms2 = Val(frmaddblc.lblinom.Caption)
    If inoms1 <> inoms2 Then
        rtv = MsgBox("Please Select Blocks With Same Nominal Current.", vbOKOnly + vbExclamation, "ERROR")
       lblinom.Caption = "inom =...." + "  [KA]"
       i = 0
        cmbsblc1.Clear
        cmbsblc2.Clear
        mdl1.blockid = 0
        With datablock.Recordset
            .MoveFirst
            Do While .EOF = False
                If .Fields("name") = cmbblctype.Text Then
                    cmbsblc1.AddItem Str(.Fields("diameter")) + "*" + Str(.Fields("thickness")), i
                    cmbsblc2.AddItem Str(.Fields("diameter")) + "*" + Str(.Fields("thickness")), i
                    i = i + 1
                End If
                .MoveNext
            Loop
        End With

       '/////
    Else
        s1 = cmbblctype.Text + Mid(cmbsblc1.Text, 2, 2) + Right(cmbsblc1.Text, 2)
        With datablock.Recordset
            .MoveFirst
            .Find "code='" + s1 + "'", , adSearchForward, 1
            mdl1.inom = .Fields("inom")
            mdl1.blockid = .Fields("ID")
            lblinom.Caption = "inom =" + Str(.Fields("inom")) + "  [KA]"
        End With
    End If
Else
     s1 = cmbblctype.Text + Mid(cmbsblc1.Text, 2, 2) + Right(cmbsblc1.Text, 2)
     With datablock.Recordset
         .MoveFirst
         .Find "code='" + s1 + "'", , adSearchForward, 1
         mdl1.inom = .Fields("inom")
         mdl1.blockid = .Fields("ID")
         lblinom.Caption = "inom =" + Str(.Fields("inom")) + "  [KA]"
     End With
End If
End Sub


Private Sub cmbsblc1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmbnblc1.SetFocus
End If
End Sub


Private Sub cmbsblc2_Click()
Dim s1, s2 As String
Dim rtv As Integer
cmdapply.Enabled = True
lbltool2.Caption = "New Serial Number: #########" + " To: #########"

s1 = cmbblctype.Text + Mid(cmbsblc2.Text, 2, 2) + Right(cmbsblc2.Text, 2)
With datablock.Recordset
    .MoveFirst
    .Find "code='" + s1 + "'", , adSearchForward, 1
    inoms2 = .Fields("inom")
End With
If cmbsblc1.Text <> "" Then
    s1 = cmbblctype.Text + Mid(cmbsblc1.Text, 2, 2) + Right(cmbsblc1.Text, 2)
    With datablock.Recordset
        .MoveFirst
        .Find "code='" + s1 + "'", , adSearchForward, 1
        inoms1 = .Fields("inom")
    End With
    If inoms1 <> inoms2 And cmbsblc1.Text <> "" Then
        rtv = MsgBox("Please Select Blocks With Same Nominal Current.", vbOKOnly + vbExclamation, "ERROR")
        lblinom.Caption = "inom =...." + "  [KA]"
        i = 0
        cmbsblc1.Clear
        cmbsblc2.Clear
        mdl1.blockid = 0
        With datablock.Recordset
            .MoveFirst
            Do While .EOF = False
                If .Fields("name") = cmbblctype.Text Then
                    cmbsblc1.AddItem Str(.Fields("diameter")) + "*" + Str(.Fields("thickness")), i
                    cmbsblc2.AddItem Str(.Fields("diameter")) + "*" + Str(.Fields("thickness")), i
                    i = i + 1
                End If
                .MoveNext
            Loop
        End With

    Else
        s1 = cmbblctype.Text + Mid(cmbsblc2.Text, 2, 2) + Right(cmbsblc2.Text, 2)
        With datablock.Recordset
            .MoveFirst
            .Find "code='" + s1 + "'", , adSearchForward, 1
            lblinom.Caption = "inom =" + Str(.Fields("inom")) + "  [KA]"
            mdl1.inom = .Fields("inom")
            mdl1.blockid = .Fields("ID")
        End With
        
    End If
Else
    s1 = cmbblctype.Text + Mid(cmbsblc2.Text, 2, 2) + Right(cmbsblc2.Text, 2)
    With datablock.Recordset
        .MoveFirst
        .Find "code='" + s1 + "'", , adSearchForward, 1
        lblinom.Caption = "inom =" + Str(.Fields("inom")) + "  [KA]"
        mdl1.inom = .Fields("inom")
        mdl1.blockid = .Fields("ID")
    End With
End If
End Sub




Private Sub cmdapply_Click()


allow = True

frmmain.MousePointer = 11

allow = CheckInputData()
Call ApplySerialNumber

If SerialIsOK Then
  cmdapply.Enabled = False
  cmdok.SetFocus
End If

frmmain.MousePointer = 0
End Sub

Private Sub CMDCANCEL_Click()
dataarrtype.Refresh
Data1.Refresh
datablock.Refresh

dataarrtype.Recordset.Close
Data1.Recordset.Close
datablock.Recordset.Close

Unload Me
FrmInterface.Show
End Sub

Private Sub CmdNofloor_Change()

lbltool2.Caption = "New Serial Number: #########" + " To: #########"


End Sub

Private Sub CMDOK_Click()


frmmain.MousePointer = 11

allow = CheckInputData()
Call ApplySerialNumber

If SerialIsOK Then
  cmdapply.Enabled = False
  frmconfirm.Show vbModal
End If

frmmain.MousePointer = 0

End Sub












Private Sub Command1_Click()

Data1.RecordSource = "DIST"
Data1.Refresh
mdl1.ArrTableName = "DIST"
   
With Data1.Recordset
           
    .MoveFirst
    Do While .EOF = False  ' Finding the last registered serial Number
                  
         Randomize
         'cnt = Abs(Int((200000 - 999999 + 1) * Rnd + 200000))
         
        !Tracing_Code = Abs(Int(Abs((100000 - 999999 + 1)) * Rnd + 100000))
       !Production_Completed = True
        .Update
        .MoveNext
    Loop
End With

End Sub


Private Sub Form_Activate()
Dim i, cnt As Integer
Dim typestr As String
Dim allow As Boolean
arrstate = 1
isHVArrester = False
txticaccp.Text = ""
txtuc.Text = ""
txtiref.Text = ""
txturefx.Text = ""
txturefn.Text = ""
txturesx.Text = ""
txturesn.Text = ""
TXTARRNO.Text = ""
txtorderno.Text = ""
txtUr1.Text = ""
TxtCode.Text = ""
CmdNofloor.ListIndex = 0
txtcustomer.Text = ""
Call cmbarr1_Click

If datablock.Recordset.RecordCount >= 1 Then
        mdl1.blcnum = 1
        ReDim mdl1.blocks(1 To 1) As String
        With datablock
            .Recordset.MoveFirst
             mdl1.blocks(mdl1.blcnum) = Mid(.Recordset.Fields("name"), 1, Len(.Recordset.Fields("name")))
                
            .Recordset.MoveNext
            Do While .Recordset.EOF = False
                allow = True
                For i = 1 To mdl1.blcnum
                    If .Recordset.Fields("name") = mdl1.blocks(i) Then
                    allow = False
                    Exit For
                    End If
                Next i
                If allow = True Then
                    mdl1.blcnum = mdl1.blcnum + 1
                    ReDim Preserve mdl1.blocks(1 To mdl1.blcnum) As String
                    mdl1.blocks(mdl1.blcnum) = .Recordset.Fields("name")
                End If
                .Recordset.MoveNext
            Loop
        End With
        '///////////////////
        mdl1.blockid = 0
        cmbblctype.Clear
        cmbsblc1.Clear
        cmbsblc2.Clear
        cmbnblc1.Clear
        cmbnblc2.Clear
        For i = 1 To mdl1.blcnum
            cmbblctype.AddItem mdl1.blocks(i), i - 1
        Next i
         '//////////////////
    Else
        MsgBox "Block Information Not Fond", vbOKOnly, "Error"
    End If

End Sub

'        ****************Initial Setting For Form And Control*************
'        *****************************************************************
Private Sub Form_Load()
Dim frhe, frwi, w1, h1, hb, wb, cnt, TableCnt As Long
Dim s1, s2, typestr As String

'     *********data control configuration*************
With Data1 '   connected to main data base
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.maindb + ";Jet OLEDB:Database Password=" + mdl1.MainDBPas
    .CommandType = adCmdTable
    .RecordSource = "SUBST"
    .Refresh
End With

With datavac '   connected to Vaccum data base
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.VacDBAdress + ";Jet OLEDB:Database Password=" + mdl1.VacDBpas
    .CommandType = adCmdTable
    .RecordSource = "STATION_CLASS"
    .Refresh
    .Recordset.Close
End With

With datablock
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.acldb + ";Jet OLEDB:Database Password=" + mdl1.AclDBPas
    .CommandType = adCmdTable
    .RecordSource = "block"
    .Refresh
End With


With dataarrtype
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.acldb + ";Jet OLEDB:Database Password=" + mdl1.AclDBPas
    .CommandType = adCmdTable
    .RecordSource = "Type"
    .Refresh
    TableCnt = .Recordset.RecordCount
End With

ReDim Tabletype(0 To TableCnt - 1) As String  ' define Table name in main data base
ReDim SerialCode(0 To TableCnt - 1) As Integer ' define arrester coding in serial number , 1 digit
ReDim ArrTypeCode(0 To TableCnt - 1) As String ' define All Arrestre types


With DataGTIN
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.acldb + ";Jet OLEDB:Database Password=" + mdl1.AclDBPas
    .CommandType = adCmdTable
    .RecordSource = "GTINCoding"
    .Refresh
    TableCnt = .Recordset.RecordCount
End With

ReDim AllGTINCodes(0 To TableCnt - 1) As Double   ' define All Arrester coding in Tavanir system

CmbPrdcCode.Clear
cnt = 0
With DataGTIN
    .Recordset.MoveFirst
    Do While DataGTIN.Recordset.EOF = False
        AllGTINCodes(cnt) = .Recordset.Fields("Coding") ' define all arrester coding in Tavanir
        s2 = .Recordset.Fields("ProductDescription") + " : " + Str(.Recordset.Fields("Coding"))
        CmbPrdcCode.AddItem s2, cnt
        cnt = cnt + 1
        .Recordset.MoveNext
    Loop
End With
CmbPrdcCode.ListIndex = 1

CmbGLN.Clear
CmbGLN.AddItem Str(mdl1.GLNCode), 0
CmbGLN.ListIndex = 0

frmmain.icon = LoadPicture(mdl1.icon)
cmbarr1.Clear
cnt = 0
With dataarrtype
    .Recordset.MoveFirst
    Do While dataarrtype.Recordset.EOF = False
        typestr = Mid(.Recordset.Fields("Arr_Type"), 1, 3) 'Arrester Type
        cmbarr1.AddItem .Recordset.Fields("Arr_Type"), cnt
        Tabletype(cnt) = .Recordset.Fields("TableName")
        SerialCode(cnt) = .Recordset.Fields("Code")
        ArrTypeCode(cnt) = .Recordset.Fields("Arr_Type")
        cnt = cnt + 1
        .Recordset.MoveNext
    Loop
End With



With CmbArrModel
  .Clear
  .AddItem "ETM-25-H10B", 0
  
  .AddItem "ETM-36-K10B", 1
  
  .AddItem "Other types", 2
  .ListIndex = 0
End With





'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\Form configuration\\\\\\\\\\\\\\\\\\\\\\\\\
With frmmain
    .Height = Screen.Height * 0.9
    .Width = Screen.Width * 0.8
    .ScaleHeight = .Height * 1.005
    .ScaleWidth = .Width * 1.005
End With

hei = frmmain.Height
wid = frmmain.Width


'               ********* FRAME Configuration***********
With framarrrate
    .Top = hei * 0.05
    .Width = wid * 0.9
    .Height = hei * 0.5
    .Left = wid * 0.05
End With
frhe = framarrrate.Height
frwi = framarrrate.Width

With framtprmt
    .Top = framarrrate.Height * 1.05 + framarrrate.Top
    .Left = wid * 0.05
    .Height = hei * 0.3
    .Width = wid * 0.425
    .Enabled = True
End With

With framblock
    .Top = framtprmt.Top
    .Left = wid * 0.05 + framtprmt.Width + framtprmt.Left
    .Height = framtprmt.Height
    .Width = wid * 0.425
    .Enabled = True
End With

With LblGLN
    .Left = framtprmt.Left
    .Height = hei * 0.05
    .Top = framtprmt.Height + framtprmt.Top * 1.05
End With

With CmbGLN
    .Left = LblGLN.Left + LblGLN.Width * 1.1
    .Width = framtprmt.Width * 0.2
    
    .Top = framtprmt.Height + framtprmt.Top * 1.05
End With

With LblPrdcCode
    .Left = CmbGLN.Left + CmbGLN.Width * 1.1
    .Height = hei * 0.05
    .Top = framtprmt.Height + framtprmt.Top * 1.05
End With

With CmbPrdcCode
    .Left = LblPrdcCode.Left + LblPrdcCode.Width * 1.1
    .Width = framtprmt.Width * 0.4
    .Top = framtprmt.Height + framtprmt.Top * 1.05
End With


With LblMonth
    .Left = LblPrdcCode.Left
    .Height = LblPrdcCode.Height
    .Top = CmbGLN.Top + CmbGLN.Height * 1.1
End With

With CmbMonth
    .Left = CmbPrdcCode.Left
    .Width = CmbPrdcCode.Width
    .Top = CmbGLN.Top + CmbGLN.Height * 1.1
    .Clear
    .AddItem "›—Ê—œÌ‰", 0
    .AddItem "«—œÌ»Â‘ ", 1
    .AddItem "Œ—œ«œ", 2
    .AddItem " Ì—", 3
    .AddItem "„—œ«œ", 4
    .AddItem "‘Â—ÌÊ—", 5
    .AddItem "„Â—", 6
    .AddItem "¬»«‰", 7
    .AddItem "¬–—", 8
    .AddItem "œÌ", 9
    .AddItem "»Â„‰", 10
    .AddItem "«”›‰œ", 11
    .ListIndex = 0
End With









With cmdapply
    .Left = framblock.Left + framblock.Width * 0.1
    .Width = framblock.Width * 0.2
    .Height = hei * 0.05
    .Top = framblock.Height + framblock.Top * 1.05
End With

With cmdok
    .Left = framblock.Left + framblock.Width * 0.4
    .Width = cmdapply.Width
    .Height = hei * 0.05
    .Top = cmdapply.Top
End With

With cmdcancel
    .Left = framblock.Left + framblock.Width * 0.7
    .Width = cmdapply.Width
    .Height = hei * 0.05
    .Top = cmdapply.Top
End With



' //////////////////////////////////////////////////////////////////////////////////////////
'  ..............           Frame setting 1

With lblarr1
    .Top = frhe * 0.1
    .Left = frwi * 0.03
    .Width = frwi * 0.12
    .Height = frhe * 0.15
End With
With cmbarr1
    .Top = frhe * 0.21
    .Width = lblarr1.Width
    .Left = lblarr1.Left
End With

With CmbArrModel
    .Top = frhe * 0.35
    .Width = lblarr1.Width
    .Left = lblarr1.Left
End With

With lblarr2
    .Top = lblarr1.Top
    .Left = lblarr1.Left + lblarr1.Width + frwi * 0.03
    .Width = lblarr1.Width
    .Height = lblarr1.Height
End With
With txtUr1
    .Top = cmbarr1.Top
    .Width = lblarr2.Width
    .Left = lblarr2.Left
    .Height = cmbarr1.Height
End With

With lblcode
    .Top = lblarr1.Top
    .Left = lblarr2.Left + lblarr2.Width + frwi * 0.03
    .Width = lblarr1.Width * 1.2
    .Height = lblarr1.Height
End With
With TxtCode
    .Height = cmbarr1.Height
    .Top = cmbarr1.Top
    .Width = frwi * 0.12
    .Left = lblcode.Left
End With


With LBLARRNO
    .Left = lblcode.Left + lblcode.Width + frwi * 0.03
    .Width = frwi * 0.1
    .Height = lblarr1.Height
    .Top = lblarr1.Top
End With
With TXTARRNO
     .Left = LBLARRNO.Left
    .Width = LBLARRNO.Width
    .Height = cmbarr1.Height
    .Top = cmbarr1.Top
End With

With lblorderno
    .Left = LBLARRNO.Left + LBLARRNO.Width + frwi * 0.03
    .Width = lblcode.Width
    .Height = lblarr1.Height
    .Top = lblarr1.Top
End With

With txtorderno
    .Left = lblorderno.Left
    .Width = lblcode.Width
    .Height = cmbarr1.Height
    .Top = cmbarr1.Top
End With

With lblcustomer
    .Left = lblorderno.Left + lblorderno.Width + frwi * 0.03
    .Width = lblcode.Width
    .Height = lblarr1.Height
    .Top = lblarr1.Top
End With

With txtcustomer
    .Left = lblcustomer.Left
    .Width = lblcustomer.Width
    .Height = cmbarr1.Height
    .Top = cmbarr1.Top
End With



With framHV
    .Left = TxtCode.Left
    .Width = txtorderno.Width + TXTARRNO.Width + (txtorderno.Left - TXTARRNO.Left - TXTARRNO.Width)
    .Height = frhe * 0.5
    .Top = CmbArrModel.Top
    .Enabled = False
End With

With framCNT
    .Left = txtorderno.Left
    .Width = txtorderno.Width + txtcustomer.Width + (txtcustomer.Left - txtorderno.Left - txtorderno.Width)
    .Height = frhe * 0.5
    .Top = framHV.Top
    .Enabled = False
    
End With


With lblnofloor
    .Left = framHV.Width * 0.05
    .Width = framHV.Width * 0.3
    .Height = framHV.Height * 0.3
    .Top = framHV.Height * 0.3
    .Caption = "No. of Arrester Unit"
End With

With CmdNofloor
    .Left = lblnofloor.Left + (lblnofloor.Width * 1.1)
    .Width = framHV.Width * 0.15
    '.Height = cmbarr1.Height
   .Top = lblnofloor.Top
   .Clear
   .AddItem "1", 0
   .AddItem "2", 1
   .AddItem "3", 2
   .AddItem "4", 3
   .AddItem "5", 4
   .ListIndex = 0
   
End With

If Tabletype(0) = "SUBST" Then ' Multy unit arrester
    CmdNofloor.ListIndex = 0
    CmdNofloor.Enabled = True
Else   ' Single
    CmdNofloor.ListIndex = 0 '= "1"
    CmdNofloor.Enabled = False
End If

With LblDiaph
    .Left = lblnofloor.Left
    .Width = lblnofloor.Width
    .Height = lblnofloor.Height
    .Top = framHV.Height * 0.6
End With

With cmbDiaph
    .Clear
    .AddItem "Steel", 0
    .AddItem "Nickel", 1
    .ListIndex = 0
        
    .Left = LblDiaph.Left + LblDiaph.Width
    .Width = LblDiaph.Width * 1.1
    .Top = LblDiaph.Top
End With





hb = framCNT.Height
wb = framCNT.Width

With LBLMan
    .Left = wb * 0.23
    .Width = wb * 0.3
    .Height = hb * 0.1
    .Top = hb * 0.18
    .Caption = "Manufacturer"
End With

With LblModel
    .Left = wb * 0.67
    .Width = wb * 0.2
    .Height = hb * 0.1
    .Top = hb * 0.18
    .Caption = "Model"
End With

With LBLMeter
    .Left = wb * 0.02
    .Width = wb * 0.3
    .Height = hb * 0.1
    .Top = hb * 0.3
    .Caption = "  Meter"
End With

With LBLCnt
    .Left = LBLMeter.Left
    .Width = LBLMeter.Width * 1.4
    .Height = hb * 0.1
    .Top = hb * 0.5
    .Caption = "Counter"
End With

With TXTMeterMan
    .Left = wb * 0.23
    .Width = wb * 0.27
    .Height = hb * 0.1
    .Top = LBLMeter.Top
End With

With TXTMeterModel
    .Left = wb * 0.55
    .Width = wb * 0.4
    .Height = hb * 0.1
    .Top = TXTMeterMan.Top
End With

With TXTCntMan
    .Left = TXTMeterMan.Left
    .Width = TXTMeterMan.Width
    .Height = TXTMeterMan.Height
    .Top = LBLCnt.Top
End With

With TXTCntModel
    .Left = TXTMeterModel.Left
    .Width = TXTMeterModel.Width
    .Height = TXTMeterModel.Height
    .Top = LBLCnt.Top
End With

With Lblprotection
    .Left = LBLCnt.Left
    .Width = wb * 0.5
    .Height = LBLCnt.Height
    .Top = hb * 0.75
    .Caption = "High Current Protection Type"
End With

With CmBProtection
    .Left = TXTMeterModel.Left
    .Width = TXTMeterModel.Width
    .Top = Lblprotection.Top
    .Clear
    .AddItem "Air Gap,0.2 mm", 0
    .AddItem "Spark Gap Module", 1
    .AddItem "Not Applied", 2
    .AddItem "Other Design", 3
    .ListIndex = 0
End With


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
With lblselarr
      .Top = frhe * 0.5
      .Width = frwi * 0.42
      .Left = lblarr1.Left * 0.8
      .Height = lblarr1.Height
End With
With lbltool2
      .Top = frhe * 0.6
      .Width = lblselarr.Width
      .Left = lblarr1.Left * 0.8
      .Height = lblarr1.Height
End With

hb = framblock.Height
wb = framblock.Width


lblnblc2.Visible = False
lblsblc2.Visible = False
cmbnblc2.Visible = False
cmbsblc2.Visible = False
lblt2.Visible = False

'*******************************

'/////////////////////////////////////////////
w1 = framtprmt.Width
h1 = framtprmt.Height
With lbluc
    .Top = h1 * 1 / 6 + (1 / 8) * h1
    .Left = w1 / 24
    .Height = h1 / 6
    .Width = w1 / 6
End With
With lblurt
    .Top = h1 * 2 / 6 + (1 / 8) * h1
    .Left = w1 / 18
    .Height = h1 / 6
    .Width = w1 * 0.15
End With
With lbliref
    .Top = h1 * 3 / 6 + (1 / 8) * h1
    .Left = w1 / 24
    .Height = h1 / 6
    .Width = w1 / 6
End With
With lblic
    .Top = h1 * 4 / 6 + (1 / 8) * h1
    .Left = w1 / 35
    .Height = h1 * 0.1
    .Width = w1 * 0.1
End With
With txtuc
    .Top = h1 * 1 / 6 + (1 / 8) * h1
    .Left = w1 / 6.5
    .Height = h1 / 10
    .Width = w1 / 6
End With
With txtPDMax
    .Top = h1 * 2 / 6 + (1 / 8) * h1
    .Left = w1 / 6.5
    .Height = h1 / 10
    .Width = w1 / 6
End With
With txtiref
    .Top = h1 * 3 / 6 + (1 / 8) * h1
    .Left = w1 / 6.5
    .Height = h1 / 10
    .Width = w1 / 6
End With
With txticaccp
    .Top = h1 * 4 / 6 + (1 / 8) * h1
    .Left = w1 / 6.5
    .Height = h1 / 10
    .Width = w1 / 6
End With
With lblu1
    .Top = h1 * 1 / 6 + (1 / 8) * h1
    .Left = w1 * 0.35
    .Height = h1 / 10
    .Width = w1 / 12
End With
With lblu2
    .Top = h1 * 2 / 6 + (1 / 8) * h1
    .Left = w1 * 0.35
    .Height = h1 / 10
    .Width = w1 / 12
End With
With lblu3
    .Top = h1 * 3 / 6 + (1 / 8) * h1
    .Left = w1 * 0.35
    .Height = h1 / 10
    .Width = w1 / 12
End With
With lblu4
    .Top = h1 * 4 / 6 + (1 / 8) * h1
    .Left = w1 * 0.35
    .Height = h1 / 10
    .Width = w1 / 12
End With
With lblurefn
    .Top = h1 * 1 / 6 + (1 / 8) * h1
    .Left = w1 * 0.52
    .Height = h1 / 10
    .Width = w1 / 7
End With
With lblurefx
    .Top = h1 * 2 / 6 + (1 / 8) * h1
    .Left = w1 * 0.52
    .Height = h1 / 10
    .Width = w1 / 7
End With
With lbluresn
    .Top = h1 * 3 / 6 + (1 / 8) * h1
    .Left = w1 * 0.52
    .Height = h1 / 10
    .Width = w1 / 7
End With
With lbluresx
    .Top = h1 * 4 / 6 + (1 / 8) * h1
    .Left = w1 * 0.52
    .Height = h1 / 10
    .Width = w1 / 6
End With

With txturefn
    .Top = h1 * 1 / 6 + (1 / 8) * h1
    .Left = w1 * 0.68
    .Height = h1 / 10
    .Width = w1 / 7
End With
With txturefx
    .Top = h1 * 2 / 6 + (1 / 8) * h1
    .Left = w1 * 0.68
    .Height = h1 / 10
    .Width = w1 / 7
End With
With txturesn
    .Top = h1 * 3 / 6 + (1 / 8) * h1
    .Left = w1 * 0.68
    .Height = h1 / 10
    .Width = w1 / 7
End With
With txturesx
    .Top = h1 * 4 / 6 + (1 / 8) * h1
    .Left = w1 * 0.68
    .Height = h1 / 10
    .Width = w1 / 7
End With

With lblu5
    .Top = h1 * 1 / 6 + (1 / 8) * h1
    .Left = w1 * 0.85
    .Height = h1 / 10
    .Width = w1 * 0.12
End With
With lblu6
    .Top = h1 * 2 / 6 + (1 / 8) * h1
    .Left = w1 * 0.85
    .Height = h1 / 10
    .Width = w1 * 0.12
End With
With lblu7
    .Top = h1 * 3 / 6 + (1 / 8) * h1
    .Left = w1 * 0.85
    .Height = h1 / 10
    .Width = w1 * 0.12
End With
With lblu8
    .Top = h1 * 4 / 6 + (1 / 8) * h1
    .Left = w1 * 0.85
    .Height = h1 / 10
    .Width = w1 * 0.12
End With
'////////////////////////////////////////////

With lblt1
    .Top = hb * 0.45
    .Left = wb * 0.05
    .Height = hb * 0.1
    .Width = wb * 0.4
End With
With lblt2
    .Top = hb * 0.45
    .Left = wb * 0.55
    .Height = hb * 0.1
    .Width = wb * 0.4
End With

With lblblctype
    .Top = hb * 0.3
    .Left = wb * 0.05
    .Height = hb * 0.1
    .Width = wb * 0.25
End With
With cmbblctype
    .Top = hb * 0.3
    .Left = wb * 0.3
    .Width = wb * 0.25
End With

With lblsblc1
    .Top = hb * 0.6
    .Left = wb * 0.05
    .Height = hb * 0.1
    .Width = wb * 0.2
End With
With lblnblc1
    .Top = hb * 0.8
    .Left = wb * 0.05
    .Height = hb * 0.1
    .Width = wb * 0.2
End With
With cmbsblc1
    .Top = hb * 0.6
    .Left = wb * 0.2
    .Width = wb * 0.2
End With
With cmbnblc1
    .Top = hb * 0.8
    .Left = wb * 0.2
    .Width = wb * 0.2
End With
With lblsblc2
    .Top = hb * 0.6
    .Left = wb * 0.55
    .Height = hb * 0.1
    .Width = wb * 0.2
End With
With lblnblc2
    .Top = hb * 0.8
    .Left = wb * 0.55
    .Height = hb * 0.1
    .Width = wb * 0.2
End With
With cmbsblc2
    .Top = hb * 0.6
    .Left = wb * 0.7
    .Width = wb * 0.2
End With
With cmbnblc2
    .Top = hb * 0.8
    .Left = wb * 0.7
    .Width = wb * 0.2
End With

With lblinom
    .Top = hb * 0.31
    .Left = wb * 0.6
    .Width = wb * 0.2
    .Caption = "inom =.....[KA]"
End With
cmbarr1.ListIndex = 0
Call cmbarr1_Click
lbltool2.Caption = "New Serial Number: #########" + " To: #########"

End Sub










Private Sub txtcode_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then
    KeyAscii = 0
End If

If KeyAscii = 13 Then
   TXTARRNO.SetFocus
End If
End Sub

Private Sub txtcustomer_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 And CmdNofloor.Enabled = True Then
   CmdNofloor.SetFocus
ElseIf KeyAscii = 13 And CmdNofloor.Enabled = False Then
   txtuc.SetFocus
End If

End Sub


Private Sub txtUr1_Change()

lblselarr.Caption = "Arrester Coding:   " + cmbarr1.Text + "-" + txtUr1.Text + "-" + UCase(TxtCode.Text)
lbltool2.Caption = "New Serial Number: #########" + " To: #########"
End Sub


Private Sub txtUr1_KeyPress(KeyAscii As Integer)

If KeyAscii = 32 Then
    KeyAscii = 0
End If

KeyAscii = mdl1.inputtest(KeyAscii)

If KeyAscii = 13 Then
   TxtCode.SetFocus
End If

End Sub


Private Sub TXTARRNO_KeyPress(KeyAscii As Integer)
KeyAscii = mdl1.inputtest(KeyAscii)
cmdapply.Enabled = True
lbltool2.Caption = "New Serial Number: #########" + " To: #########"
If KeyAscii = 13 Then
   txtorderno.SetFocus
End If

End Sub



Private Sub txtcode_Change()
lblselarr.Caption = "Arrester Coding:   " + cmbarr1.Text + "-" + txtUr1.Text + "-" + UCase(TxtCode.Text)
lbltool2.Caption = "New Serial Number: #########" + " To: #########"
End Sub

Private Sub txticaccp_KeyPress(KeyAscii As Integer)
KeyAscii = mdl1.inputtest(KeyAscii)
cmdapply.Enabled = True
lbltool2.Caption = "New Serial Number: #########" + " To: #########"

If KeyAscii = 13 Then
   txturefn.SetFocus
End If

End Sub
Private Sub txtiref_KeyPress(KeyAscii As Integer)
KeyAscii = mdl1.inputtest(KeyAscii)
cmdapply.Enabled = True
lbltool2.Caption = "New Serial Number: #########" + " To: #########"

If KeyAscii = 13 Then
   txticaccp.SetFocus
End If
End Sub



Private Sub txtorderNo_KeyPress(KeyAscii As Integer)
cmdapply.Enabled = True
lbltool2.Caption = "New Serial Number: #########" + " To: #########"

If KeyAscii = 13 Then
   txtcustomer.SetFocus
End If

End Sub


Private Sub txtuc_KeyPress(KeyAscii As Integer)

KeyAscii = mdl1.inputtest(KeyAscii)
cmdapply.Enabled = True
lbltool2.Caption = "New Serial Number: #########" + " To: #########"

If KeyAscii = 13 Then
   txtPDMax.SetFocus
End If


End Sub

Private Sub txtUr1_LostFocus()
txtuc.Text = Str(Round(Val(txtUr1.Text) * 0.8, 1))
End Sub

Private Sub txturefn_KeyPress(KeyAscii As Integer)
KeyAscii = mdl1.inputtest(KeyAscii)
cmdapply.Enabled = True
lbltool2.Caption = "New Serial Number: #########" + " To: #########"

If KeyAscii = 13 Then
   txturefx.SetFocus
End If
End Sub

Private Sub txturefx_KeyPress(KeyAscii As Integer)
KeyAscii = mdl1.inputtest(KeyAscii)
cmdapply.Enabled = True
lbltool2.Caption = "New Serial Number: #########" + " To: #########"
If KeyAscii = 13 Then
   txturesn.SetFocus
End If
End Sub

Private Sub txturesn_KeyPress(KeyAscii As Integer)
KeyAscii = mdl1.inputtest(KeyAscii)
cmdapply.Enabled = True
lbltool2.Caption = "New Serial Number: #########" + " To: #########"
If KeyAscii = 13 Then
   txturesx.SetFocus
End If
End Sub


Private Sub txturesx_KeyPress(KeyAscii As Integer)
KeyAscii = mdl1.inputtest(KeyAscii)
cmdapply.Enabled = True
lbltool2.Caption = "New Serial Number: #########" + " To: #########"


If KeyAscii = 13 Then
   cmbblctype.SetFocus
End If

End Sub


Private Sub txtPDMax_KeyPress(KeyAscii As Integer)


lbltool2.Caption = "New Serial Number: #########" + " To: #########"


If KeyAscii = 13 Then
   txtiref.SetFocus
End If

End Sub


