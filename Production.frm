VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmProdData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "›—„ À»  «ÿ·«⁄«  ‘‰«”«ÌÌ Ê —œÌ«»Ì  Ê·Ìœ »—ﬁêÌ—"
   ClientHeight    =   9720
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16590
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "B Lotus"
      Size            =   15.75
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9720
   ScaleWidth      =   16590
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame framdata 
      Caption         =   "5-  “—Ìﬁ ”Ì·ÌòÊ‰"
      BeginProperty Font 
         Name            =   "B Lotus"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Index           =   4
      Left            =   12480
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   5500
      Width           =   3735
      Begin VB.TextBox cmbdate 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   4
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   72
         Text            =   " "
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CheckBox chkmodify 
         Alignment       =   1  'Right Justify
         Caption         =   "«’·«Õ «ÿ·«⁄« "
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   1920
         TabIndex        =   32
         Top             =   2040
         Width           =   1575
      End
      Begin VB.ComboBox cmbname 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   240
         TabIndex        =   6
         Text            =   "À»  ‰‘œÂ"
         Top             =   720
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPicke 
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   65
         Top             =   1200
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   130547713
         CurrentDate     =   44285
      End
      Begin VB.Label lblname 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "  ‰«„ „Ã—Ì"
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   4
         Left            =   2640
         TabIndex        =   49
         Top             =   600
         Width           =   840
      End
      Begin VB.Label lbldate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ «‰Ã«„"
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   4
         Left            =   2625
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   1200
         Width           =   855
      End
   End
   Begin VB.Frame framdata 
      Caption         =   "8- »” Â »‰œÌ"
      BeginProperty Font 
         Name            =   "B Lotus"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Index           =   7
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   5500
      Width           =   3735
      Begin VB.TextBox cmbdate 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   7
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   75
         Text            =   " "
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CheckBox chkmodify 
         Alignment       =   1  'Right Justify
         Caption         =   "«’·«Õ «ÿ·«⁄« "
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   1920
         TabIndex        =   35
         Top             =   2040
         Width           =   1575
      End
      Begin VB.ComboBox cmbname 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   240
         TabIndex        =   9
         Text            =   "À»  ‰‘œÂ"
         Top             =   720
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPicke 
         Height          =   375
         Index           =   7
         Left            =   240
         TabIndex        =   68
         Top             =   1320
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   130613249
         CurrentDate     =   44285
      End
      Begin VB.Label lblname 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "  ‰«„ „Ã—Ì"
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   7
         Left            =   2640
         TabIndex        =   46
         Top             =   600
         Width           =   840
      End
      Begin VB.Label lbldate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ «‰Ã«„"
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   7
         Left            =   2625
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   1200
         Width           =   855
      End
   End
   Begin VB.Frame framdata 
      Caption         =   "7-  «ÌÌœ ò‰ —· òÌ›Ì "
      BeginProperty Font 
         Name            =   "B Lotus"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Index           =   6
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   5500
      Width           =   3735
      Begin VB.TextBox cmbdate 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   6
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   74
         Text            =   " "
         Top             =   1920
         Width           =   1575
      End
      Begin VB.ComboBox cmbname 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   120
         TabIndex        =   8
         Text            =   "À»  ‰‘œÂ"
         Top             =   720
         Width           =   1935
      End
      Begin VB.CheckBox chkmodify 
         Alignment       =   1  'Right Justify
         Caption         =   "«’·«Õ «ÿ·«⁄« "
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   1920
         TabIndex        =   34
         Top             =   2040
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPicke 
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   67
         Top             =   1320
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   130613249
         CurrentDate     =   44285
      End
      Begin VB.Label lbldate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ «‰Ã«„"
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   6
         Left            =   2745
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblname 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "  ‰«„ „Ã—Ì"
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   6
         Left            =   2745
         TabIndex        =   42
         Top             =   720
         Width           =   840
      End
   End
   Begin VB.Frame framdata 
      Caption         =   "6- ò‰ —· Ÿ«Â—Ì"
      BeginProperty Font 
         Name            =   "B Lotus"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Index           =   5
      Left            =   8400
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   5500
      Width           =   3735
      Begin VB.TextBox cmbdate 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   5
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Text            =   " "
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CheckBox chkmodify 
         Alignment       =   1  'Right Justify
         Caption         =   "«’·«Õ «ÿ·«⁄« "
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   1920
         TabIndex        =   33
         Top             =   2040
         Width           =   1575
      End
      Begin VB.ComboBox cmbname 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   240
         TabIndex        =   7
         Text            =   "À»  ‰‘œÂ"
         Top             =   720
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPicke 
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   66
         Top             =   1200
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   130613249
         CurrentDate     =   44285
      End
      Begin VB.Label lblname 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "  ‰«„ „Ã—Ì"
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   5
         Left            =   2640
         TabIndex        =   40
         Top             =   600
         Width           =   840
      End
      Begin VB.Label lbldate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ «‰Ã«„"
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   5
         Left            =   2625
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   1200
         Width           =   855
      End
   End
   Begin VB.Frame framdata 
      Caption         =   "4-  “—Ìﬁ —“Ì‰"
      BeginProperty Font 
         Name            =   "B Lotus"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Index           =   3
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   2600
      Width           =   3735
      Begin VB.TextBox cmbdate 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   3
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   71
         Text            =   " "
         Top             =   1920
         Width           =   1575
      End
      Begin VB.ComboBox cmbname 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   240
         TabIndex        =   5
         Text            =   "À»  ‰‘œÂ"
         Top             =   720
         Width           =   1935
      End
      Begin VB.CheckBox chkmodify 
         Alignment       =   1  'Right Justify
         Caption         =   "«’·«Õ «ÿ·«⁄« "
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1920
         TabIndex        =   31
         Top             =   2040
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPicke 
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   64
         Top             =   1320
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   130613249
         CurrentDate     =   44285
      End
      Begin VB.Label lbldate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ «‰Ã«„"
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   3
         Left            =   2625
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblname 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "  ‰«„ „Ã—Ì"
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   3
         Left            =   2640
         TabIndex        =   36
         Top             =   600
         Width           =   840
      End
   End
   Begin VB.Frame framdata 
      Caption         =   "3- ‰’» „Ì·Â Å«· —Êé‰ («·Ì«› ‘Ì‘Â)"
      BeginProperty Font 
         Name            =   "B Lotus"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Index           =   2
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   2600
      Width           =   3735
      Begin VB.ComboBox cmbname 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   240
         TabIndex        =   77
         Text            =   "À»  ‰‘œÂ"
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox cmbdate 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   2
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Text            =   " "
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CheckBox chkmodify 
         Alignment       =   1  'Right Justify
         Caption         =   "«’·«Õ «ÿ·«⁄« "
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   27
         Top             =   2040
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPicke 
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   63
         Top             =   1320
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   130613249
         CurrentDate     =   44285
      End
      Begin VB.Label lblname 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "  ‰«„ „Ã—Ì"
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   2
         Left            =   2640
         TabIndex        =   29
         Top             =   600
         Width           =   840
      End
      Begin VB.Label lbldate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ «‰Ã«„"
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   2
         Left            =   2625
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   1200
         Width           =   855
      End
   End
   Begin VB.Frame framdata 
      Caption         =   "2- ¬„«œÂ ”«“Ì ” Ê‰ ﬁ—’"
      BeginProperty Font 
         Name            =   "B Lotus"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Index           =   1
      Left            =   8400
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   2600
      Width           =   3735
      Begin VB.TextBox cmbdate 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   69
         Text            =   " "
         Top             =   1920
         Width           =   1575
      End
      Begin VB.ComboBox cmbname 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Text            =   "À»  ‰‘œÂ"
         Top             =   720
         Width           =   1935
      End
      Begin VB.CheckBox chkmodify 
         Alignment       =   1  'Right Justify
         Caption         =   "«’·«Õ «ÿ·«⁄« "
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   23
         Top             =   2040
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPicke 
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   62
         Top             =   1320
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   130678785
         CurrentDate     =   44285
      End
      Begin VB.Label lbldate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ «‰Ã«„"
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   2625
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblname 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "  ‰«„ „Ã—Ì"
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   2640
         TabIndex        =   24
         Top             =   600
         Width           =   840
      End
   End
   Begin VB.Frame framserial 
      Caption         =   "  ‘„«—Â ”—Ì«· »—ﬁêÌ—"
      BeginProperty Font 
         Name            =   "B Lotus"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   240
      Width           =   15975
      Begin VB.ComboBox CmbArrType 
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   14280
         TabIndex        =   58
         Text            =   "ETM"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtsubquantity 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7800
         RightToLeft     =   -1  'True
         TabIndex        =   57
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txttotalquantity 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   7800
         RightToLeft     =   -1  'True
         TabIndex        =   56
         TabStop         =   0   'False
         Text            =   " "
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtlastSNselected 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   200
         RightToLeft     =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtfirstSNselected 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtlastSN 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   200
         RightToLeft     =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   " "
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtfirstSN 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   " "
         Top             =   480
         Width           =   1935
      End
      Begin VB.ComboBox CmbSubOrder 
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   11040
         TabIndex        =   2
         Text            =   "1"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtorderNo 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11040
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Text            =   " "
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblmodel 
         Caption         =   "Arrester type  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7680
         TabIndex        =   76
         Top             =   120
         Width           =   6255
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "‰Ê⁄ »—ﬁêÌ—"
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   14400
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   720
         Width           =   960
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " ⁄œ«œ œ— “Ì— »—‰«„Â:"
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   9000
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   1320
         Width           =   1785
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " ⁄œ«œ ò· »—‰«„Â:"
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   9000
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   600
         Width           =   1485
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " « :"
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   2300
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   1320
         Width           =   270
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "‘„«—Â ”—Ì«· «‰ Œ«» ‘œÂ: "
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   1
         Left            =   5000
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   1320
         Width           =   2490
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " « :"
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   2300
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   480
         Width           =   270
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "‘„«—Â ”—Ì«· ò· »—‰«„Â: "
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   0
         Left            =   5000
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   480
         Width           =   2265
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "‘„«—Â “Ì—»—‰«„Â"
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   12360
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   1440
         Width           =   1425
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "‘„«—Â »—‰«„Â"
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   12480
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   720
         Width           =   1125
      End
   End
   Begin VB.CommandButton cmdend 
      Caption         =   "Œ—ÊÃ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9120
      TabIndex        =   13
      Top             =   8520
      Width           =   1815
   End
   Begin VB.CommandButton cmdapply 
      Caption         =   "À» "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13440
      TabIndex        =   12
      Top             =   8520
      Width           =   1815
   End
   Begin VB.Frame framdata 
      Caption         =   "1- ÿ—«ÕÌ »—ﬁêÌ—"
      BeginProperty Font 
         Name            =   "B Lotus"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Index           =   0
      Left            =   12480
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   2600
      Width           =   3735
      Begin VB.TextBox cmbdate 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   0
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Text            =   " "
         Top             =   1920
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPicke 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   60
         Top             =   1320
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   130678785
         CurrentDate     =   44285
      End
      Begin VB.CheckBox chkmodify 
         Alignment       =   1  'Right Justify
         Caption         =   "«’·«Õ «ÿ·«⁄« "
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   14
         Top             =   2040
         Width           =   1575
      End
      Begin VB.ComboBox cmbname 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Text            =   "À»  ‰‘œÂ"
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label lblname 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "  ‰«„ „Ã—Ì"
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   0
         Left            =   2640
         TabIndex        =   11
         Top             =   600
         Width           =   840
      End
      Begin VB.Label lbldate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ «‰Ã«„"
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   0
         Left            =   2625
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1200
         Width           =   855
      End
   End
   Begin MSAdodcLib.Adodc DataPerson 
      Height          =   375
      Left            =   2400
      Top             =   8280
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Caption         =   "dataperson"
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
   Begin MSAdodcLib.Adodc DataArrester 
      Height          =   375
      Left            =   0
      Top             =   8280
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "DataArrester"
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
   Begin MSAdodcLib.Adodc DataProduction 
      Height          =   375
      Left            =   2520
      Top             =   8760
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Caption         =   "DataProduction"
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
   Begin MSAdodcLib.Adodc DataArrType 
      Height          =   375
      Left            =   120
      Top             =   8880
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Caption         =   "DataProduction"
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
End
Attribute VB_Name = "FrmProdData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sn, cnt, cnt0, cnt1, cnt2, cnt3, cnt4, cnt5, cnt6, cnt7, temarrnum, fc As Double
Dim sernum() As Double
Dim today1, today2 As String
Dim NoOfSubprog, totalquantity As Integer
Dim FirstSelectedSN, LastSelectedSN As Double
Dim Tabletype() As String
Dim SelectedSN() As Double
Dim Rooz_Sel(0 To 7) As Single
Dim Mah_Sel(0 To 7) As Single
Dim Sal_Sel(0 To 7) As Single
Dim Rooz_as_sal(0 To 7) As Single

Private Sub CmbArrType_Change()

With DataArrester

    If Tabletype(CmbArrType.ListIndex) = "SUBST" Then   '  Station Class Arrester Coding
        .RecordSource = "SUBST"
       ' mdl1.ArrTableName = "SUBST"
        
    ElseIf Tabletype(CmbArrType.ListIndex) = "DIST" Then '  dISTRIBUTION Class Arrester Coding
       .RecordSource = "DIST"
        'mdl1.ArrTableName = "DIST"
   ElseIf Tabletype(CmbArrType.ListIndex) = "CNT" Then '  Counter
       .RecordSource = "CNT"
        'mdl1.ArrTableName = "DIST"
        
    End If
    
    .Refresh
End With
txtorderNo.Text = ""

With DataArrType.Recordset
        
        If .RecordCount > 0 Then
            .MoveFirst
           .Find ("Arr_Type='") + CmbArrType.Text + "'", , adSearchForward, 1
              
            If .EOF = False Then
               framdata(0).Caption = "  1- " + .Fields("Production_Phase0")
               framdata(1).Caption = "  2- " + .Fields("Production_Phase1")
               framdata(2).Caption = "  3- " + .Fields("Production_Phase2")
               framdata(3).Caption = "  4- " + .Fields("Production_Phase3")
               framdata(4).Caption = "  5- " + .Fields("Production_Phase4")
               framdata(5).Caption = "  6- " + .Fields("Production_Phase5")
               framdata(6).Caption = "  7- " + .Fields("Production_Phase6")
               framdata(7).Caption = "  8- " + .Fields("Production_Phase7")
                
            End If
            
         End If

End With




End Sub

Private Sub CmbArrType_Click()
With DataArrester

    If Tabletype(CmbArrType.ListIndex) = "SUBST" Then   '  Station Class Arrester Coding
        .RecordSource = "SUBST"
       ' mdl1.ArrTableName = "SUBST"
        
    ElseIf Tabletype(CmbArrType.ListIndex) = "DIST" Then '  dISTRIBUTION Class Arrester Coding
       .RecordSource = "DIST"
        'mdl1.ArrTableName = "DIST"
   ElseIf Tabletype(CmbArrType.ListIndex) = "CNT" Then '  Counter
       .RecordSource = "CNT"
        'mdl1.ArrTableName = "DIST"
        
    End If
    
    .Refresh
    
End With
txtorderNo.Text = ""
With DataArrType.Recordset
        
        If .RecordCount > 0 Then
            .MoveFirst
            .Find ("Arr_Type='") + CmbArrType.Text + "'", , adSearchForward, 1
             
            If .EOF = False Then
               framdata(0).Caption = "  1- " + .Fields("Production_Phase0")
               framdata(1).Caption = "  2- " + .Fields("Production_Phase1")
               framdata(2).Caption = "  3- " + .Fields("Production_Phase2")
               framdata(3).Caption = "  4- " + .Fields("Production_Phase3")
               framdata(4).Caption = "  5- " + .Fields("Production_Phase4")
               framdata(5).Caption = "  6- " + .Fields("Production_Phase5")
               framdata(6).Caption = "  7- " + .Fields("Production_Phase6")
               framdata(7).Caption = "  8- " + .Fields("Production_Phase7")
                
            End If
            
         End If

End With

End Sub




Private Sub cmbdate_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub CmbSubOrder_Change()
txtfirstSNselected.Text = Str(Val(txtfirstSN.Text) + (40 * CmbSubOrder.ListIndex))

If Val(txtfirstSNselected.Text) + 39 < Val(txtlastSN.Text) Then
    txtlastSNselected.Text = Str(Val(txtfirstSNselected.Text) + 39)
    LastSelectedSN = Val(txtlastSNselected.Text)
Else
    txtlastSNselected.Text = txtlastSN.Text
    LastSelectedSN = Val(txtlastSNselected.Text)
End If
txtsubquantity.Text = Str(Val(txtlastSNselected.Text) - Val(txtfirstSNselected.Text) + 1)

For i = 0 To 7
    cmbname(i).ListIndex = 0
    cmbdate(i).Text = To_HejriChar(Date, 0) ' set current date as selected date of calender
Next i


ReDim SelectedSN(1 To Val(txtsubquantity.Text)) As Double
For i = 1 To Val(txtsubquantity.Text)
    SelectedSN(i) = Val(txtfirstSNselected.Text) + i - 1
Next i

'//////////////////////////////////////////////////////////////////////////////////
                ' ‰„«Ì‘ «ÿ·«⁄«  ﬁ»·« À»  ‘œÂ
                
                With DataArrester.Recordset
                    If .RecordCount > 0 Then
        
                        .MoveFirst
                        .Find ("Serial_Num=") + txtfirstSNselected.Text, , adSearchForward, 1
                        ' Add Data
                        If .EOF = False Then
                        
                             '  «ÿ·«⁄«  ÿ—«ÕÌ »—ﬁêÌ—Ì
                            For i = 0 To cmbname(0).ListCount - 1
                                If .Fields("Production_phase0_Oprt") = cmbname(0).List(i) Then
                                    cmbname(0).ListIndex = i
                                    cmbname(0).BackColor = vbGreen
                                    Exit For
                                Else
                                    cmbname(0).BackColor = vbRed
                                End If
                            Next i
                            If .Fields("Production_phase0_Date") <> "1300/01/01" Then
                                 cmbdate(0).Text = .Fields("Production_phase0_Date")
                                 cmbdate(0).BackColor = vbGreen
                            Else
                                 cmbdate(0).BackColor = vbRed
                            End If
                            
                             '  «ÿ·«⁄«  ﬁ—’ ÅÌçÌ
                            For i = 0 To cmbname(1).ListCount - 1
                                If .Fields("Production_phase1_Oprt") = cmbname(1).List(i) Then
                                    cmbname(1).ListIndex = i
                                    cmbname(1).BackColor = vbGreen
                                    Exit For
                                Else
                                    cmbname(1).BackColor = vbRed
                                End If
                            Next i
                            If .Fields("Production_phase1_Date") <> "1300/01/01" Then
                                 cmbdate(1).Text = .Fields("Production_phase1_Date")
                                 cmbdate(1).BackColor = vbGreen
                            Else
                                 cmbdate(1).BackColor = vbRed
                            End If
                            
                            '  «ÿ·«⁄«  «·Ì«› ÅÌç
                            For i = 0 To cmbname(2).ListCount - 1
                                If .Fields("Production_phase2_Oprt") = cmbname(2).List(i) Then
                                    cmbname(2).ListIndex = i
                                    cmbname(2).BackColor = vbGreen
                                    Exit For
                                Else
                                    cmbname(2).BackColor = vbRed
                                End If
                            Next i
                            If .Fields("Production_phase2_Date") <> "1300/01/01" Then
                                 cmbdate(2).Text = .Fields("Production_phase2_Date")
                                 cmbdate(2).BackColor = vbGreen
                            Else
                                 cmbdate(2).BackColor = vbRed
                            End If
                            
                            '  «ÿ·«⁄«   “—Ìﬁ —“Ì‰
                            For i = 0 To cmbname(3).ListCount - 1
                                If .Fields("Production_phase3_Oprt") = cmbname(3).List(i) Then
                                    cmbname(3).ListIndex = i
                                    cmbname(3).BackColor = vbGreen
                                    Exit For
                                Else
                                    cmbname(3).BackColor = vbRed
                                End If
                            Next i
                            If .Fields("Production_phase3_Date") <> "1300/01/01" Then
                                 cmbdate(3).Text = .Fields("Production_phase3_Date")
                                 cmbdate(3).BackColor = vbGreen
                            Else
                                 cmbdate(3).BackColor = vbRed
                            End If
                            
                            ' «ÿ·«⁄«   “—Ìﬁ ”Ì·ÌòÊ‰
                            For i = 0 To cmbname(4).ListCount - 1
                                If .Fields("Production_phase4_Oprt") = cmbname(4).List(i) Then
                                    cmbname(4).ListIndex = i
                                    cmbname(4).BackColor = vbGreen
                                    Exit For
                                Else
                                    cmbname(4).BackColor = vbRed
                                End If
                            Next i
                            If .Fields("Production_phase4_Date") <> "1300/01/01" Then
                                 cmbdate(4).Text = .Fields("Production_phase4_Date")
                                 cmbdate(4).BackColor = vbGreen
                            Else
                                 cmbdate(4).BackColor = vbRed
                            End If
                
                            
                            ' «ÿ·«⁄«  ò‰ —· Ÿ«Â—Ì
                            For i = 0 To cmbname(5).ListCount - 1
                                If .Fields("Production_phase5_Oprt") = cmbname(5).List(i) Then
                                    cmbname(5).ListIndex = i
                                    cmbname(5).BackColor = vbGreen
                                    Exit For
                                Else
                                    cmbname(5).BackColor = vbRed
                                End If
                            Next i
                            If .Fields("Production_phase5_Date") <> "1300/01/01" Then
                                 cmbdate(5).Text = .Fields("Production_phase5_Date")
                                 cmbdate(5).BackColor = vbGreen
                            Else
                                 cmbdate(5).BackColor = vbRed
                            End If
                            
                            
                            ' «ÿ·«⁄«   «ÌÌœ ò‰ —· òÌ›Ì
                            For i = 0 To cmbname(6).ListCount - 1
                                If .Fields("Production_phase6_Oprt") = cmbname(6).List(i) Then
                                    cmbname(6).ListIndex = i
                                    cmbname(6).BackColor = vbGreen
                                    Exit For
                                Else
                                    cmbname(6).BackColor = vbRed
                                End If
                            Next i
                            If .Fields("Production_phase6_Date") <> "1300/01/01" Then
                                 cmbdate(6).Text = .Fields("Production_phase6_Date")
                                 cmbdate(6).BackColor = vbGreen
                            Else
                                 cmbdate(6).BackColor = vbRed
                            End If
                            
                            ' «ÿ·«⁄«  »” Â »‰œÌ
                            For i = 0 To cmbname(7).ListCount - 1
                                If .Fields("Production_phase7_Oprt") = cmbname(7).List(i) Then
                                    cmbname(7).ListIndex = i
                                    cmbname(7).BackColor = vbGreen
                                    Exit For
                                Else
                                    cmbname(7).BackColor = vbRed
                                End If
                            Next i
                            If .Fields("Production_phase7_Date") <> "1300/01/01" Then
                                 cmbdate(7).Text = .Fields("Production_phase7_Date")
                                 cmbdate(7).BackColor = vbGreen
                            Else
                                 cmbdate(7).BackColor = vbRed
                            End If
                            
                
                        End If '.EOF = False Then
                    End If '.RecordCount > 0 Then
                End With
                  
          '////////////////////////////////////////////////////////////////////////////


End Sub

Private Sub CmbSubOrder_Click()
' Calculate Serial Number of Sub program

txtfirstSNselected.Text = Str(Val(txtfirstSN.Text) + (40 * CmbSubOrder.ListIndex))

If Val(txtfirstSNselected.Text) + 39 < Val(txtlastSN.Text) Then
    txtlastSNselected.Text = Str(Val(txtfirstSNselected.Text) + 39)
    LastSelectedSN = Val(txtlastSNselected.Text)
Else
    txtlastSNselected.Text = txtlastSN.Text
    LastSelectedSN = Val(txtlastSNselected.Text)
End If
txtsubquantity.Text = Str(Val(txtlastSNselected.Text) - Val(txtfirstSNselected.Text) + 1)

For i = 0 To 7
    cmbname(i).ListIndex = 0
    cmbdate(i).Text = To_HejriChar(Date, 0) ' set current date as selected date of calender
Next i

' Show registered data

ReDim SelectedSN(1 To Val(txtsubquantity.Text)) As Double
For i = 1 To Val(txtsubquantity.Text)
    SelectedSN(i) = Val(txtfirstSNselected.Text) + i - 1
Next i

'//////////////////////////////////////////////////////////////////////////////////
                ' ‰„«Ì‘ «ÿ·«⁄«  ﬁ»·« À»  ‘œÂ
                
                With DataArrester.Recordset
                    If .RecordCount > 0 Then
        
                        .MoveFirst
                        .Find ("Serial_Num=") + txtfirstSNselected.Text, , adSearchForward, 1
                        ' Add Data
                        If .EOF = False Then
                        
                             '  «ÿ·«⁄«  ÿ—«ÕÌ »—ﬁêÌ—Ì
                            For i = 0 To cmbname(0).ListCount - 1
                                If .Fields("Production_phase0_Oprt") = cmbname(0).List(i) Then
                                    cmbname(0).ListIndex = i
                                    cmbname(0).BackColor = vbGreen
                                    Exit For
                                Else
                                    cmbname(0).BackColor = vbRed
                                End If
                            Next i
                            If .Fields("Production_phase0_Date") <> "1300/01/01" Then
                                 cmbdate(0).Text = .Fields("Production_phase0_Date")
                                 cmbdate(0).BackColor = vbGreen
                            Else
                                 cmbdate(0).BackColor = vbRed
                            End If
                            
                             '  «ÿ·«⁄«  ﬁ—’ ÅÌçÌ
                            For i = 0 To cmbname(1).ListCount - 1
                                If .Fields("Production_phase1_Oprt") = cmbname(1).List(i) Then
                                    cmbname(1).ListIndex = i
                                    cmbname(1).BackColor = vbGreen
                                    Exit For
                                Else
                                    cmbname(1).BackColor = vbRed
                                End If
                            Next i
                            If .Fields("Production_phase1_Date") <> "1300/01/01" Then
                                 cmbdate(1).Text = .Fields("Production_phase1_Date")
                                 cmbdate(1).BackColor = vbGreen
                            Else
                                 cmbdate(1).BackColor = vbRed
                            End If
                            
                            '  «ÿ·«⁄«  «·Ì«› ÅÌç
                            For i = 0 To cmbname(2).ListCount - 1
                                If .Fields("Production_phase2_Oprt") = cmbname(2).List(i) Then
                                    cmbname(2).ListIndex = i
                                    cmbname(2).BackColor = vbGreen
                                    Exit For
                                Else
                                    cmbname(2).BackColor = vbRed
                                End If
                            Next i
                            If .Fields("Production_phase2_Date") <> "1300/01/01" Then
                                 cmbdate(2).Text = .Fields("Production_phase2_Date")
                                 cmbdate(2).BackColor = vbGreen
                            Else
                                 cmbdate(2).BackColor = vbRed
                            End If
                            
                            '  «ÿ·«⁄«   “—Ìﬁ —“Ì‰
                            For i = 0 To cmbname(3).ListCount - 1
                                If .Fields("Production_phase3_Oprt") = cmbname(3).List(i) Then
                                    cmbname(3).ListIndex = i
                                    cmbname(3).BackColor = vbGreen
                                    Exit For
                                Else
                                    cmbname(3).BackColor = vbRed
                                End If
                            Next i
                            If .Fields("Production_phase3_Date") <> "1300/01/01" Then
                                 cmbdate(3).Text = .Fields("Production_phase3_Date")
                                 cmbdate(3).BackColor = vbGreen
                            Else
                                 cmbdate(3).BackColor = vbRed
                            End If
                            
                            ' «ÿ·«⁄«   “—Ìﬁ ”Ì·ÌòÊ‰
                            For i = 0 To cmbname(4).ListCount - 1
                                If .Fields("Production_phase4_Oprt") = cmbname(4).List(i) Then
                                    cmbname(4).ListIndex = i
                                    cmbname(4).BackColor = vbGreen
                                    Exit For
                                Else
                                    cmbname(4).BackColor = vbRed
                                End If
                            Next i
                            If .Fields("Production_phase4_Date") <> "1300/01/01" Then
                                 cmbdate(4).Text = .Fields("Production_phase4_Date")
                                 cmbdate(4).BackColor = vbGreen
                            Else
                                 cmbdate(4).BackColor = vbRed
                            End If
                
                            
                            ' «ÿ·«⁄«  ò‰ —· Ÿ«Â—Ì
                            For i = 0 To cmbname(5).ListCount - 1
                                If .Fields("Production_phase5_Oprt") = cmbname(5).List(i) Then
                                    cmbname(5).ListIndex = i
                                    cmbname(5).BackColor = vbGreen
                                    Exit For
                                Else
                                    cmbname(5).BackColor = vbRed
                                End If
                            Next i
                            If .Fields("Production_phase5_Date") <> "1300/01/01" Then
                                 cmbdate(5).Text = .Fields("Production_phase5_Date")
                                 cmbdate(5).BackColor = vbGreen
                            Else
                                 cmbdate(5).BackColor = vbRed
                            End If
                            
                            
                            ' «ÿ·«⁄«   «ÌÌœ ò‰ —· òÌ›Ì
                            For i = 0 To cmbname(6).ListCount - 1
                                If .Fields("Production_phase6_Oprt") = cmbname(6).List(i) Then
                                    cmbname(6).ListIndex = i
                                    cmbname(6).BackColor = vbGreen
                                    Exit For
                                Else
                                    cmbname(6).BackColor = vbRed
                                End If
                            Next i
                            If .Fields("Production_phase6_Date") <> "1300/01/01" Then
                                 cmbdate(6).Text = .Fields("Production_phase6_Date")
                                 cmbdate(6).BackColor = vbGreen
                            Else
                                 cmbdate(6).BackColor = vbRed
                            End If
                            
                            ' «ÿ·«⁄«  »” Â »‰œÌ
                            For i = 0 To cmbname(7).ListCount - 1
                                If .Fields("Production_phase7_Oprt") = cmbname(7).List(i) Then
                                    cmbname(7).ListIndex = i
                                    cmbname(7).BackColor = vbGreen
                                    Exit For
                                Else
                                    cmbname(7).BackColor = vbRed
                                End If
                            Next i
                            If .Fields("Production_phase7_Date") <> "1300/01/01" Then
                                 cmbdate(7).Text = .Fields("Production_phase7_Date")
                                 cmbdate(7).BackColor = vbGreen
                            Else
                                 cmbdate(7).BackColor = vbRed
                            End If
                            
                
                        End If '.EOF = False Then
                    End If '.RecordCount > 0 Then
                End With
                  
          '////////////////////////////////////////////////////////////

End Sub


Private Sub cmdapply_Click()
Dim CheckData As Boolean
Dim ReadyToTest(0 To 5) As Boolean

For i = 0 To 5
    ReadyToTest(i) = False
Next i

CheckData = True
' check inserted date to be in correct sequence

For i = 0 To 7
    Rooz_Sel(i) = Right(cmbdate(i).Text, 2)
    Mah_Sel(i) = Mid(cmbdate(i).Text, 6, 2)
    Sal_Sel(i) = Mid(cmbdate(i).Text, 1, 4)
    
    If Mah_Sel(i) = 1 Then
        Rooz_as_sal(i) = ((Sal_Sel(i) - 1300) * 1000) + Rooz_Sel(i)
    ElseIf Mah_Sel(i) >= 2 And Mah_Sel(i) <= 7 Then
        Rooz_as_sal(i) = ((Sal_Sel(i) - 1300) * 1000) + ((Mah_Sel(i) - 1) * 31) + Rooz_Sel(i)
    Else
        Rooz_as_sal(i) = ((Sal_Sel(i) - 1300) * 1000) + ((Mah_Sel(i) - 1) * 30) + 6 + Rooz_Sel(i)
    End If
    
Next i

For i = 1 To 7
    If Rooz_as_sal(i) < Rooz_as_sal(i - 1) And cmbname(i).ListIndex > 0 Then
        cmbdate(i - 1).BackColor = vbYellow
        cmbdate(i).BackColor = vbYellow
        CheckData = False
        
    End If
Next i

If CheckData = False Then
     MsgBox " «—ÌŒ «‰ Œ«»Ì ›—«Ì‰œÂ« »«Ìœ »Â  — Ì» «‰Ã«„ ›—«Ì‰œ »«‘œ." + Chr(13) + "  «—ÌŒ —« «’·«Õ ò‰Ìœ.", vbOKOnly, "Date Sequence Error"
End If

'/////////////////////////////////////////////////////

With DataArrester.Recordset
    If .RecordCount > 0 And Val(txtsubquantity.Text) > 0 And CheckData Then
        .MoveFirst
        For cnt = 1 To Val(txtsubquantity.Text)
            .Find ("Serial_Num=") + Str(SelectedSN(cnt)), , adSearchForward, 1
                If .EOF = False Then
                    For i = 0 To 5
                        ReadyToTest(i) = False
                    Next i
                    
                 ' «ÿ·«⁄«  ÿ—«ÕÌ »—ﬁêÌ—
                    If cmbname(0).ListIndex > 0 Then
                        !Production_phase0_Oprt = cmbname(0).Text
                        !Production_phase0_Date = cmbdate(0).Text
                         ReadyToTest(0) = True
                    End If
                    
                    '  «ÿ·«⁄«  ﬁ—’ ÅÌçÌ
                    If cmbname(1).ListIndex > 0 Then
                        !Production_phase1_Oprt = cmbname(1).Text
                        !Production_phase1_Date = cmbdate(1).Text
                        ReadyToTest(1) = True
                    End If
                    
                     '  «ÿ·«⁄«  «·Ì«› ÅÌçÌ
                    If cmbname(2).ListIndex > 0 Then
                        !Production_phase2_Oprt = cmbname(2).Text
                        !Production_phase2_Date = cmbdate(2).Text
                        ReadyToTest(2) = True
                    End If
                    
                     '  «ÿ·«⁄«   “—Ìﬁ —“Ì‰
                    If cmbname(3).ListIndex > 0 Then
                        !Production_phase3_Oprt = cmbname(3).Text
                        !Production_phase3_Date = cmbdate(3).Text
                        ReadyToTest(3) = True
                    End If
                    
                      '  «ÿ·«⁄«   “—Ìﬁ ”Ì·ÌòÊ‰
                     If cmbname(4).ListIndex > 0 Then
                        !Production_phase4_Oprt = cmbname(4).Text
                        !Production_phase4_Date = cmbdate(4).Text
                        ReadyToTest(4) = True
                    End If
                    
                     '  «ÿ·«⁄«  ò‰ —· Ÿ«Â—Ì
                    If cmbname(5).ListIndex > 0 Then
                        !Production_phase5_Oprt = cmbname(5).Text
                        !Production_phase5_Date = cmbdate(5).Text
                        ReadyToTest(5) = True
                    End If
                    
                     '  «ÿ·«⁄«   «ÌÌœ ò‰ —· òÌ›Ì
                   If cmbname(6).ListIndex > 0 Then
                   
                        If .Fields("result1") = 1 Or .Fields("result1") = 2 Then
                           rtv = MsgBox("»—«Ì À»  «ÿ·«⁄«  ò‰ —· òÌ›Ì »‰œÌ°«» œ« »«Ìœ  ”  «·ò —ÌòÌ «‰Ã«„ ê—œœ" + Chr(13) + Chr(13) + " ”  «·ò —ÌòÌ »—ﬁêÌ— »« ‘„«—Â ”—Ì«·  " + Str(.Fields("Serial_Num")) + "  «‰Ã«„ ‰‘œÂ «” " + Chr(13) + Chr(13) + "¬Ì« „«Ì· »Â «œ«„Â ›—«Ì‰œ „Ì »«‘Ìœø", vbYesNo, "Date Sequence Error")
                            If rtv <> 6 Then
                                Exit For
                            End If
                            
                        ElseIf .Fields("result1") > 3 And .Fields("result2") <> 3 Then
                         
                          rtv = MsgBox("»—«Ì À»  «ÿ·«⁄«  ò‰ —· òÌ›Ì°«» œ« »«Ìœ  ”  «·ò —ÌòÌ «‰Ã«„ ê—œœ" + Chr(13) + Chr(13) + " ”  «·ò —ÌòÌ »—ﬁêÌ— »« ‘„«—Â ”—Ì«·  " + Str(.Fields("Serial_Num")) + "  ﬁ»Ê· ‰‘œÂ «” " + Chr(13) + Chr(13) + "¬Ì« „«Ì· »Â «œ«„Â ›—«Ì‰œ „Ì »«‘Ìœø", vbYesNo, "Date Sequence Error")
                            If rtv <> 6 Then
                                Exit For
                            End If
                            
                        ElseIf .Fields("result1") = 3 Or .Fields("result2") = 3 Then
                            !Production_phase6_Oprt = cmbname(6).Text
                            !Production_phase6_Date = cmbdate(6).Text
                        End If
                        
                    End If
                    
                      '  «ÿ·«⁄«  »” Â »‰œÌ
                    If cmbname(7).ListIndex > 0 Then
                        
                        If .Fields("result1") = 1 Or .Fields("result1") = 2 Then
                           rtv = MsgBox("»—«Ì À»  «ÿ·«⁄«  »” Â »‰œÌ°«» œ« »«Ìœ  ”  «·ò —ÌòÌ «‰Ã«„ ê—œœ" + Chr(13) + Chr(13) + " ”  «·ò —ÌòÌ »—ﬁêÌ— »« ‘„«—Â ”—Ì«·  " + Str(.Fields("Serial_Num")) + "  «‰Ã«„ ‰‘œÂ «” " + Chr(13) + Chr(13) + "¬Ì« „«Ì· »Â «œ«„Â ›—«Ì‰œ „Ì »«‘Ìœø", vbYesNo, "Date Sequence Error")
                            If rtv <> 6 Then
                                Exit For
                            End If
                            
                        ElseIf .Fields("result1") > 3 And .Fields("result2") <> 3 Then
                         
                            rtv = MsgBox("»—«Ì À»  «ÿ·«⁄«  »” Â »‰œÌ°«» œ« »«Ìœ  ”  «·ò —ÌòÌ «‰Ã«„ ê—œœ" + Chr(13) + Chr(13) + " ”  «·ò —ÌòÌ »—ﬁêÌ— »« ‘„«—Â ”—Ì«·  " + Str(.Fields("Serial_Num")) + "  ﬁ»Ê· ‰‘œÂ «” " + Chr(13) + Chr(13) + "¬Ì« „«Ì· »Â «œ«„Â ›—«Ì‰œ „Ì »«‘Ìœø", vbYesNo, "Date Sequence Error")
                            If rtv <> 6 Then
                                Exit For
                            End If
                            
                        ElseIf .Fields("result1") = 3 Or .Fields("result2") = 3 Then
                            !Production_phase7_Oprt = cmbname(7).Text
                            !Production_phase7_Date = cmbdate(7).Text
                        End If
                      
                    End If
                    
                    If ReadyToTest(0) And ReadyToTest(1) And ReadyToTest(2) And ReadyToTest(3) And ReadyToTest(4) And ReadyToTest(5) And .Fields("result1") = 1 Then
                        !result1 = 2
                    End If
                     
                    .Update
                    
                    '/////////////////////////////////////////
                    ' »——”Ì « „«„ ›—«Ì‰œ  Ê·œ »—ﬁêÌ—
                    
                    If .Fields("Result1") = 3 Or .Fields("Result2") = 3 Then
                        If .Fields("Production_phase0_Oprt") <> "No One" And .Fields("Production_phase0_Date") <> "1300/01/01" Then
                            If .Fields("Production_phase1_Oprt") <> "No One" And .Fields("Production_phase1_Date") <> "1300/01/01" Then
                                If .Fields("Production_phase2_Oprt") <> "No One" And .Fields("Production_phase2_Date") <> "1300/01/01" Then
                                    If .Fields("Production_phase3_Oprt") <> "No One" And .Fields("Production_phase3_Date") <> "1300/01/01" Then
                                        If .Fields("Production_phase4_Oprt") <> "No One" And .Fields("Production_phase4_Date") <> "1300/01/01" Then
                                            If .Fields("Production_phase5_Oprt") <> "No One" And .Fields("Production_phase5_Date") <> "1300/01/01" Then
                                                If .Fields("Production_phase7_Oprt") <> "No One" And .Fields("Production_phase7_Date") <> "1300/01/01" Then
                                                    If .Fields("Production_phase6_Oprt") <> "No One" And .Fields("Production_phase6_Date") <> "1300/01/01" Then
                                                        !Production_Completed = True
                                                        .Update
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    
                    End If
                    '////////////////////////////////////////////
                    
                End If
           Next cnt
       End If
End With

Call CmbSubOrder_Click


End Sub

Private Sub cmdend_Click()
Unload Me
FrmInterface.Show
End Sub

Private Sub DTPicke_Change(Index As Integer)
cmbdate(Index).Text = To_HejriChar(DTPicke(Index).value, 0)
End Sub

Private Sub Form_Activate()
MsgBox " ÃÂ  ’Õ  ‰„«Ì‘ «ÿ·«⁄«  ° «» œ« ‰Ê⁄ „Õ’Ê· —« «‰ Œ«» Ê ”Å” ‘„«—Â »—‰«„Â —« À»  ‰„«ÌÌœ." + Chr(13) + "«ÿ·«⁄«  ‘„«—Â »—‰«„Â Ê ‰Ê⁄ „Õ’Ê· »«Ìœ œ—”  «‰ Œ«» ê—œœ. ", vbOKOnly, " ÊÃÂ"

End Sub

Private Sub Form_Load()
Dim w1, h1 As Single
Dim testdate As String
'////////////////////////////////////////////
frmchnginf.icon = LoadPicture(mdl1.icon)

lblmodel.Caption = "Arrester type: "

With DataArrester
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.maindb + ";Jet OLEDB:Database Password=13581837"
    .CommandType = adCmdTable
    .RecordSource = "DIST"
    .Refresh
    
End With

With DataPerson
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.acldb + ";Jet OLEDB:Database Password=" + mdl1.AclDBPas
    .CommandType = adCmdTable
    .RecordSource = "personel"
    .Refresh
End With
With DataProduction
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.acldb + ";Jet OLEDB:Database Password=" + mdl1.AclDBPas
    .CommandType = adCmdTable
    .RecordSource = "production"
    .Refresh
End With

With DataArrType
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.acldb + ";Jet OLEDB:Database Password=" + mdl1.AclDBPas
    .CommandType = adCmdTable
    .RecordSource = "Type"
    .Refresh
    
End With


For i = 0 To 7
    cmbname(i).Clear
    cmbname(i).AddItem "À»  ‰‘œÂ", 0
Next i
    
With DataPerson.Recordset
    
    cnt0 = 1
    cnt1 = 1
    cnt2 = 1
    cnt3 = 1
    cnt4 = 1
    cnt5 = 1
    cnt6 = 1
    cnt7 = 1
    
    If .RecordCount >= 1 Then
        .MoveFirst
        For i = 1 To .RecordCount
            'Add   persons  «”«„Ì «›—«œ
            For j = 0 To 7
                cmbname(j).AddItem .Fields("firstname") + " " + .Fields("lastname"), i
                cnt0 = cnt0 + 1
                cnt1 = cnt1 + 1
                cnt2 = cnt2 + 1
                cnt3 = cnt3 + 1
                cnt4 = cnt4 + 1
                cnt5 = cnt5 + 1
                cnt6 = cnt6 + 1
                cnt7 = cnt7 + 1
            Next j
 
        .MoveNext
        Next i
        'cnt = cmbname(0).ListIndex
    End If
End With

For i = 0 To 7
    
    cmbdate(i).Text = To_HejriChar(Date, 0) ' set current date as selected date of calender
    cmbdate(i).BackColor = vbWhite
    cmbname(i).BackColor = vbWhite
Next i

With DataArrType.Recordset
    ReDim Tabletype(0 To .RecordCount - 1) As String   ' define Table name in main data base
    .MoveFirst
    If .RecordCount >= 1 Then
        For i = 1 To .RecordCount
            CmbArrType.AddItem .Fields("arr_Type"), i - 1
             Tabletype(i - 1) = .Fields("TableName")
            .MoveNext
        Next i
        CmbArrType.ListIndex = 0
        
        
    End If
End With

If mdl1.Test_Admin = True Then
    For i = 0 To 6
        chkmodify(i).Enabled = True
    Next i
Else
    For i = 0 To 6
        chkmodify(i).Enabled = False
    Next i
End If

'framdata(1).Caption = "2- ‘„«—Â ”—Ì«·"

End Sub


Private Sub Text1_Change()

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = mdl1.inputtest(KeyAscii)
End Sub

Private Sub txtfirstSN_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


Private Sub txtfirstSNselected_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


Private Sub txtlastSN_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


Private Sub txtlastSNselected_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


Private Sub txtorderNo_Change()

CmbSubOrder.Clear
txtfirstSN.Text = ""
txtlastSN.Text = ""
txtfirstSNselected.Text = ""
txtlastSNselected.Text = ""
txtsubquantity.Text = ""
txttotalquantity.Text = ""
lblmodel.Caption = "Arrester type: "

For i = 0 To 7
    cmbname(i).ListIndex = 0
    cmbdate(i).Text = To_HejriChar(Date, 0) ' set current date as selected date of calender
    cmbdate(i).BackColor = vbWhite
    cmbname(i).BackColor = vbWhite
    
Next i


End Sub

Private Sub txtorderNo_KeyPress(KeyAscii As Integer)
KeyAscii = mdl1.inputtest(KeyAscii)

If KeyAscii = 13 Then
    Call txtOrderNo_LostFocus
End If

End Sub

Private Sub txtOrderNo_LostFocus()
Dim findlast As Boolean
Dim firstSN, LastSN As Double
lblmodel.Caption = "Arrester type: "

If Val(txtorderNo.Text) > 0 Then
    findlast = False
    firstSN = 0
    LastSN = 0
    CmbSubOrder.Clear
    DataProduction.Refresh
    totalquantity = 0
    NoOfSubprog = 0
    FirstSelectedSN = 0
    LastSelectedSN = 0
    
    With DataProduction.Recordset
        
        If .RecordCount > 0 Then
            .MoveFirst
            .Find ("OrderNo=") + Str(Val(txtorderNo.Text)), , adSearchForward, 1
             
            If .EOF = False Then
                firstSN = .Fields("FirstSN")
                txtfirstSN.Text = Str(firstSN)
                
                LastSN = .Fields("lastSN")
                txtlastSN.Text = Str(LastSN)
                
                NoOfSubprog = .Fields("NoOfSubDivision")
                
                totalquantity = .Fields("Quantity")
                txttotalquantity.Text = Str(totalquantity)
                
                 lblmodel.Caption = "Arrester type :    " + .Fields("Type")
                 
                 
                For cnt = 1 To NoOfSubprog
                    CmbSubOrder.AddItem cnt, cnt - 1
                Next cnt
                CmbSubOrder.ListIndex = 0
                
                txtfirstSNselected.Text = Str(.Fields("FirstSN"))
                FirstSelectedSN = .Fields("FirstSN")
                
                If totalquantity <= 40 Then
                    txtlastSNselected.Text = Str(.Fields("lastSN"))
                    LastSelectedSN = Val(txtlastSNselected.Text)
                Else
                     txtlastSNselected.Text = Str(.Fields("FirstSN") + 39)
                     LastSelectedSN = Val(txtlastSNselected.Text)
                End If
                txtsubquantity.Text = Str(Val(txtlastSNselected.Text) - Val(txtfirstSNselected.Text) + 1)
                
             '//////////////////////////////////////////////////////////////////////////////////
                ' ‰„«Ì‘ «ÿ·«⁄«  ﬁ»·« À»  ‘œÂ
                
                With DataArrester.Recordset
                    If .RecordCount > 0 Then
        
                        .MoveFirst
                        .Find ("Serial_Num=") + txtfirstSNselected.Text, , adSearchForward, 1
                        ' Add Data
                        If .EOF = False Then
                        
                             '  «ÿ·«⁄«  ÿ—«ÕÌ »—ﬁêÌ—Ì
                            For i = 0 To cmbname(0).ListCount - 1
                                If .Fields("Production_phase0_Oprt") = cmbname(0).List(i) Then
                                    cmbname(0).ListIndex = i
                                    cmbname(0).BackColor = vbGreen
                                    Exit For
                                Else
                                    cmbname(0).BackColor = vbRed
                                End If
                            Next i
                            
                           
                            
                            If .Fields("Production_phase0_Date") <> "1300/01/01" Then
                                 cmbdate(0).Text = .Fields("Production_phase0_Date")
                                 cmbdate(0).BackColor = vbGreen
                                 
                            Else
                                 cmbdate(0).BackColor = vbRed
                            End If
                            
                             '  «ÿ·«⁄«  ﬁ—’ ÅÌçÌ
                            For i = 0 To cmbname(1).ListCount - 1
                                If .Fields("Production_phase1_Oprt") = cmbname(1).List(i) Then
                                    cmbname(1).ListIndex = i
                                    cmbname(1).BackColor = vbGreen
                                    Exit For
                                Else
                                    cmbname(1).BackColor = vbRed
                                End If
                            Next i
                            If .Fields("Production_phase1_Date") <> "1300/01/01" Then
                                 cmbdate(1).Text = .Fields("Production_phase1_Date")
                                 cmbdate(1).BackColor = vbGreen
                            Else
                                 cmbdate(1).BackColor = vbRed
                            End If
                            
                            '  «ÿ·«⁄«  «·Ì«› ÅÌç
                            For i = 0 To cmbname(2).ListCount - 1
                                If .Fields("Production_phase2_Oprt") = cmbname(2).List(i) Then
                                    cmbname(2).ListIndex = i
                                    cmbname(2).BackColor = vbGreen
                                    Exit For
                                Else
                                    cmbname(2).BackColor = vbRed
                                End If
                            Next i
                            If .Fields("Production_phase2_Date") <> "1300/01/01" Then
                                 cmbdate(2).Text = .Fields("Production_phase2_Date")
                                 cmbdate(2).BackColor = vbGreen
                            Else
                                 cmbdate(2).BackColor = vbRed
                            End If
                            
                            '  «ÿ·«⁄«   “—Ìﬁ —“Ì‰
                            For i = 0 To cmbname(3).ListCount - 1
                                If .Fields("Production_phase3_Oprt") = cmbname(3).List(i) Then
                                    cmbname(3).ListIndex = i
                                    cmbname(3).BackColor = vbGreen
                                    Exit For
                                Else
                                    cmbname(3).BackColor = vbRed
                                End If
                            Next i
                            If .Fields("Production_phase3_Date") <> "1300/01/01" Then
                                 cmbdate(3).Text = .Fields("Production_phase3_Date")
                                 cmbdate(3).BackColor = vbGreen
                            Else
                                 cmbdate(3).BackColor = vbRed
                            End If
                            
                            ' «ÿ·«⁄«   “—Ìﬁ ”Ì·ÌòÊ‰
                            For i = 0 To cmbname(4).ListCount - 1
                                If .Fields("Production_phase4_Oprt") = cmbname(4).List(i) Then
                                    cmbname(4).ListIndex = i
                                    cmbname(4).BackColor = vbGreen
                                    Exit For
                                Else
                                    cmbname(4).BackColor = vbRed
                                End If
                            Next i
                            If .Fields("Production_phase4_Date") <> "1300/01/01" Then
                                 cmbdate(4).Text = .Fields("Production_phase4_Date")
                                 cmbdate(4).BackColor = vbGreen
                            Else
                                 cmbdate(4).BackColor = vbRed
                            End If
                
                            
                            ' «ÿ·«⁄«  ò‰ —· Ÿ«Â—Ì
                            For i = 0 To cmbname(5).ListCount - 1
                                If .Fields("Production_phase5_Oprt") = cmbname(5).List(i) Then
                                    cmbname(5).ListIndex = i
                                    cmbname(5).BackColor = vbGreen
                                    Exit For
                                Else
                                    cmbname(5).BackColor = vbRed
                                End If
                            Next i
                            If .Fields("Production_phase5_Date") <> "1300/01/01" Then
                                 cmbdate(5).Text = .Fields("Production_phase5_Date")
                                 cmbdate(5).BackColor = vbGreen
                            Else
                                 cmbdate(5).BackColor = vbRed
                            End If
                            
                            
                            ' «ÿ·«⁄«   «ÌÌœ ò‰ —· òÌ›Ì
                            For i = 0 To cmbname(6).ListCount - 1
                                If .Fields("Production_phase6_Oprt") = cmbname(6).List(i) Then
                                    cmbname(6).ListIndex = i
                                    cmbname(6).BackColor = vbGreen
                                    Exit For
                                Else
                                    cmbname(6).BackColor = vbRed
                                End If
                            Next i
                            If .Fields("Production_phase6_Date") <> "1300/01/01" Then
                                 cmbdate(6).Text = .Fields("Production_phase6_Date")
                                 cmbdate(6).BackColor = vbGreen
                            Else
                                 cmbdate(6).BackColor = vbRed
                            End If
                            
                            ' «ÿ·«⁄«  »” Â »‰œÌ
                            For i = 0 To cmbname(7).ListCount - 1
                                If .Fields("Production_phase7_Oprt") = cmbname(7).List(i) Then
                                    cmbname(7).ListIndex = i
                                    cmbname(7).BackColor = vbGreen
                                    Exit For
                                Else
                                    cmbname(7).BackColor = vbRed
                                End If
                            Next i
                            If .Fields("Production_phase7_Date") <> "1300/01/01" Then
                                 cmbdate(7).Text = .Fields("Production_phase7_Date")
                                 cmbdate(7).BackColor = vbGreen
                            Else
                                 cmbdate(7).BackColor = vbRed
                            End If
                            
                
                        End If '.EOF = False Then
                    End If '.RecordCount > 0 Then
                End With
                  
          '////////////////////////////////////////////////////////////
          
            Else  ' .EOF = False Then
             
                MsgBox " ‰Ê⁄ „Õ’Ê· Ê ‘„«—Â »—‰«„Â —« œ—”  «‰ Œ«» ‰„«ÌÌœ." + Chr(13) + "«ÿ·«⁄«  ‘„«—Â »—‰«„Â ›Êﬁ ÿ»ﬁ ‰Ê⁄ „Õ’Ê· «‰ Œ«»Ì –ŒÌ—Â ‰‘œÂ «” .", vbOKOnly, "ERROR"
                txtfirstSN.Text = ""
                txtlastSN.Text = ""
                txtfirstSNselected.Text = ""
                txtlastSNselected.Text = ""
                txtorderNo.Text = ""
                txtsubquantity.Text = ""
                txttotalquantity.Text = ""
                lblmodel.Caption = "Arrester type: "
                CmbSubOrder.Clear
                
             
            End If ' .EOF = False Then
            
        Else  '.RecordCount > 0 Then
            MsgBox " ‰Ê⁄ „Õ’Ê· Ê ‘„«—Â »—‰«„Â —« œ—”  «‰ Œ«» ‰„«ÌÌœ." + Chr(13) + "«ÿ·«⁄«  ‘„«—Â »—‰«„Â ›Êﬁ ÿ»ﬁ ‰Ê⁄ „Õ’Ê· «‰ Œ«»Ì –ŒÌ—Â ‰‘œÂ «” .", vbOKOnly, "ERROR"
            txtfirstSN.Text = ""
            txtlastSN.Text = ""
            txtorderNo.Text = ""
            txtfirstSNselected.Text = ""
            txtlastSNselected.Text = ""
            txtsubquantity.Text = ""
            txttotalquantity.Text = ""
            lblmodel.Caption = "Arrester type: "
            CmbSubOrder.Clear
            
            
        End If  '.RecordCount > 0 Then
        
    End With  'DataProduction.Recordset
End If    'Val(txtorderNo.Text) > 0 Then
   
End Sub


Private Sub Text2_Change()

End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = mdl1.inputtest(KeyAscii)
End Sub


Private Sub Text3_Change()

End Sub


Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = mdl1.inputtest(KeyAscii)
End Sub


Private Sub txtsubquantity_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


Private Sub txttotalquantity_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


