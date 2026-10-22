VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmInterface 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ARTM_QC"
   ClientHeight    =   9375
   ClientLeft      =   1365
   ClientTop       =   375
   ClientWidth     =   15285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   14114.33
   ScaleMode       =   0  'User
   ScaleWidth      =   20861.94
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdDefectRgst 
      Caption         =   "À»  „‘ò·«   Ê·Ìœ „Õ’Ê·"
      BeginProperty Font 
         Name            =   "B Lotus"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   664
      Index           =   1
      Left            =   11160
      TabIndex        =   11
      Top             =   3240
      Width           =   3855
   End
   Begin VB.CommandButton Cmdtest 
      Caption         =   "«‰Ã«„  ”  «·ò —ÌòÌ"
      BeginProperty Font 
         Name            =   "B Lotus"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   664
      Left            =   11160
      TabIndex        =   2
      Top             =   1800
      Width           =   3855
   End
   Begin VB.CommandButton CmdYearset 
      Caption         =   " €ÌÌ— ”«· Ã«—Ì"
      BeginProperty Font 
         Name            =   "B Lotus"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   664
      Left            =   11160
      TabIndex        =   9
      Top             =   7560
      Width           =   3855
   End
   Begin VB.CommandButton CmdUserEdit 
      Caption         =   "«’·«Õ «ÿ·«⁄«  Å—”‰·Ì"
      BeginProperty Font 
         Name            =   "B Lotus"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   664
      Left            =   11160
      TabIndex        =   8
      Top             =   6840
      Width           =   3855
   End
   Begin VB.CommandButton CmdNewuser 
      Caption         =   "À»  «ÿ·«⁄«  Å—”‰· ÃœÌœ"
      BeginProperty Font 
         Name            =   "B Lotus"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   664
      Left            =   11160
      TabIndex        =   7
      Top             =   6120
      Width           =   3855
   End
   Begin VB.CommandButton CmdNewBlock 
      Caption         =   "À»  «ÿ·«⁄«  ﬁ—’ ÃœÌœ"
      BeginProperty Font 
         Name            =   "B Lotus"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   664
      Left            =   11160
      TabIndex        =   6
      Top             =   5400
      Width           =   3855
   End
   Begin VB.CommandButton CmdNewArr 
      Caption         =   "À»  „œ· »—ﬁêÌ— ÃœÌœ"
      BeginProperty Font 
         Name            =   "B Lotus"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   664
      Left            =   11160
      TabIndex        =   5
      Top             =   4680
      Width           =   3855
   End
   Begin VB.CommandButton CmdModify 
      Caption         =   "«’·«Õ «ÿ·«⁄«  Å«ÌÂ »—ﬁêÌ—"
      BeginProperty Font 
         Name            =   "B Lotus"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   664
      Left            =   11160
      TabIndex        =   4
      Top             =   3960
      Width           =   3855
   End
   Begin VB.CommandButton CmdReport 
      Caption         =   "ê“«—‘ êÌ—Ì Ê ç«Å ò«—  ê«—«‰ Ì"
      BeginProperty Font 
         Name            =   "B Lotus"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   664
      Left            =   11160
      TabIndex        =   3
      Top             =   2520
      Width           =   3855
   End
   Begin VB.CommandButton CmdprodData 
      Caption         =   "Ê—Êœ «ÿ·«⁄«  ›—«Ì‰œ  Ê·Ìœ"
      BeginProperty Font 
         Name            =   "B Lotus"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   664
      Left            =   11160
      TabIndex        =   1
      Top             =   1080
      Width           =   3855
   End
   Begin VB.CommandButton CmdNewSer 
      Caption         =   "«›“Êœ‰ ‘„«—Â ”—Ì«· ÃœÌœ"
      BeginProperty Font 
         Name            =   "B Lotus"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   664
      Left            =   11160
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
   Begin MSAdodcLib.Adodc DataArrtype 
      Height          =   330
      Left            =   840
      Top             =   9000
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
      Left            =   6960
      Top             =   9000
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
      Left            =   3720
      Top             =   9000
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
   Begin MSComDlg.CommonDialog cd1 
      Left            =   2160
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdend 
      Caption         =   "Œ—ÊÃ «“ »—‰«„Â"
      BeginProperty Font 
         Name            =   "B Lotus"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   664
      Left            =   11160
      TabIndex        =   10
      Top             =   8280
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   8595
      Left            =   240
      Stretch         =   -1  'True
      Top             =   360
      Width           =   10335
   End
End
Attribute VB_Name = "FrmInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private arrcnt As Double
Private temselectedarr As String
Private rtv, timecnt As Double

Private Sub cmdend_Click()
    Unload Me
    End
End Sub


Private Sub CmdNewSer_Click()
Me.Hide
frmmain.Show vbModal

End Sub

Private Sub CmdprodData_Click()
Me.Hide
FrmProdData.Show vbModal

End Sub

Private Sub CmdReport_Click()
Me.Hide
frmrpt.Show vbModal
End Sub

Private Sub CmdModify_Click()
Me.Hide
frmchnginf.Show vbModal
End Sub

Private Sub CmdNewArr_Click()
Me.Hide
frmarrtype.Show vbModal

End Sub

Private Sub CmdNewBlock_Click()
Me.Hide
frmaddblc.Show vbModal
End Sub

Private Sub CmdNewuser_Click()
frmnewper.Show vbModal
End Sub

Private Sub Cmdtest_Click()
Me.Hide
frmtest.Show vbModal
End Sub

Private Sub CmdUserEdit_Click()
frmchpas.Show vbModal

End Sub

Private Sub CmdYearset_Click()
Call mdl1.YearChng
Call Form_Load
End Sub

Private Sub Form_Load()
Dim hei, wid, h1, w1 As Long


'*************************** DTAMAIN configuration*******
'//////////////////////////////////////////////////////////////////////////////////////
FrmInterface.icon = LoadPicture(mdl1.icon)
Image1.Picture = LoadPicture(mdl1.ARTMQC_Logo)

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

'//////////////////////   User Access Setting ///////////////////////////

If mdl1.Test_Admin = True Or mdl1.HVTest_Operator = True Or mdl1.QC_Operator = True Or mdl1.QC_Manager = True Or mdl1.Technical_Manager = True Then
    CmdReport.Enabled = True
    CmdYearset.Enabled = True
    CmdNewSer.Enabled = True
Else
    CmdReport.Enabled = False
    CmdYearset.Enabled = False
    CmdNewSer.Enabled = False
End If

If mdl1.QC_Operator = True Or mdl1.QC_Manager = True Then
    CmdprodData.Enabled = True
Else
    CmdprodData.Enabled = False
End If

If mdl1.Test_Admin = True Or mdl1.Technical_Manager = True Then
    CmdNewBlock.Enabled = True
    CmdNewArr.Enabled = True
Else
    CmdNewBlock.Enabled = False
    CmdNewArr.Enabled = False
End If

If mdl1.Test_Admin = True Or mdl1.Technical_Manager = True Or mdl1.HVTest_Operator = True Then
    CmdModify.Enabled = True
Else
    CmdModify.Enabled = False
End If

If (mdl1.Test_Admin = True Or mdl1.Technical_Manager = True Or mdl1.HVTest_Operator = True) And mdl1.ProgramType = "HV Test Program" Then
    Cmdtest.Enabled = True
Else
     Cmdtest.Enabled = False
End If

If mdl1.Test_Admin = True Or mdl1.Technical_Manager = True Then
    CmdNewuser.Enabled = True
Else
     CmdNewuser.Enabled = False
End If

CmdUserEdit.Enabled = True
'/////////////////////////////////////////////////////////////////////////////////////////////

FrmInterface.Caption = "ARTM                        Current Year : " + Str(mdl1.FullYear) + "                         Operator : " + mdl1.optname



End Sub


