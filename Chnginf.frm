VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmchnginf 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Arrester Parameters Modification"
   ClientHeight    =   8610
   ClientLeft      =   3075
   ClientTop       =   1575
   ClientWidth     =   10200
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   10200
   Begin VB.TextBox txtOrderNo 
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
      Left            =   1200
      TabIndex        =   34
      Text            =   " "
      Top             =   1680
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc DataArrtype 
      Height          =   495
      Left            =   7800
      Top             =   6120
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
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
   Begin MSAdodcLib.Adodc datatestprmt 
      Height          =   375
      Left            =   7800
      Top             =   6840
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
   Begin MSAdodcLib.Adodc Datatest 
      Height          =   375
      Left            =   7920
      Top             =   7680
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
      Caption         =   "Datatest"
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
   Begin VB.TextBox txtto 
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
      Left            =   7080
      TabIndex        =   2
      Text            =   " "
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox txtfrom 
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
      Left            =   7080
      TabIndex        =   1
      Top             =   1920
      Width           =   1455
   End
   Begin VB.ComboBox cmbarrtype 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   480
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdchng 
      Caption         =   "Chenge Information"
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
      Left            =   5040
      TabIndex        =   11
      Top             =   6960
      Width           =   1455
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
      Height          =   615
      Left            =   1560
      TabIndex        =   12
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Frame framtprmt 
      Caption         =   "arrester Information "
      Height          =   2895
      Left            =   360
      TabIndex        =   13
      Top             =   2880
      Width           =   5655
      Begin VB.TextBox txtuc 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   3
         Text            =   " "
         ToolTipText     =   "Countinous Voltage "
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtiref 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   5
         Text            =   " "
         ToolTipText     =   "Reference Current "
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txticaccp 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   6
         Text            =   " "
         ToolTipText     =   "Acceptable Leakage Current "
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txturt 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   4
         Text            =   " "
         ToolTipText     =   "Rated Voltage "
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txturefn 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3960
         TabIndex        =   7
         Text            =   " "
         ToolTipText     =   "Minimum Reference Voltage"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txturefx 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3960
         TabIndex        =   8
         Text            =   " "
         ToolTipText     =   "Maximum Reference Voltage"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txturesn 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3960
         TabIndex        =   9
         Text            =   " "
         ToolTipText     =   "Minimum Residual  Voltage"
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txturesx 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3960
         TabIndex        =   10
         Text            =   " "
         ToolTipText     =   "Maximum Residual  Voltage"
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label lblurt 
         Caption         =   "   Ur ="
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label lblurefx 
         Caption         =   "Uref max ="
         Height          =   255
         Left            =   3120
         TabIndex        =   28
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lbliref 
         Caption         =   "   I ref ="
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lbluresx 
         Caption         =   "Ures max ="
         Height          =   255
         Left            =   3000
         TabIndex        =   26
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label lblic 
         Caption         =   "Ic(acp)="
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lbluc 
         Caption         =   "    Uc ="
         Height          =   375
         Left            =   0
         TabIndex        =   24
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblurefn 
         Caption         =   "Uref min ="
         Height          =   255
         Left            =   3120
         TabIndex        =   23
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lbluresn 
         Caption         =   "Ures min ="
         Height          =   255
         Left            =   3000
         TabIndex        =   22
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblu1 
         Caption         =   "[KV]"
         Height          =   255
         Left            =   2160
         TabIndex        =   21
         Top             =   840
         Width           =   375
      End
      Begin VB.Label lblu3 
         Caption         =   "[mA]"
         Height          =   375
         Left            =   2040
         TabIndex        =   20
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblu4 
         Caption         =   "[mA]"
         Height          =   255
         Left            =   2040
         TabIndex        =   19
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lblu2 
         Caption         =   "[KV]"
         Height          =   255
         Left            =   2160
         TabIndex        =   18
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label lblu5 
         Caption         =   "[KV]"
         Height          =   255
         Left            =   5040
         TabIndex        =   17
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblu6 
         Caption         =   "[KV]"
         Height          =   255
         Left            =   5040
         TabIndex        =   16
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblu7 
         Caption         =   "[KV]"
         Height          =   255
         Left            =   5040
         TabIndex        =   15
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lblu8 
         Caption         =   "[KV]"
         Height          =   255
         Left            =   5040
         TabIndex        =   14
         Top             =   2280
         Width           =   495
      End
   End
   Begin VB.Label lblOrderNo 
      Caption         =   "Order Number:"
      Height          =   375
      Left            =   360
      TabIndex        =   35
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label lbl2 
      Caption         =   "Selected Serial Numbers: "
      Height          =   375
      Left            =   6960
      TabIndex        =   33
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label lblto 
      Caption         =   "To : "
      Height          =   255
      Left            =   6000
      TabIndex        =   32
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblfrom 
      Caption         =   "From : "
      Height          =   255
      Left            =   6000
      TabIndex        =   31
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lblarrtype 
      Caption         =   "Arrester Type :"
      Height          =   375
      Left            =   1320
      TabIndex        =   30
      Top             =   2640
      Width           =   1455
   End
End
Attribute VB_Name = "frmchnginf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sn, cnt, temarrnum, fc As Double
Dim sernum() As Double
Dim blc1 As String
Dim blc2 As String
Dim Tabletype() As String
Private Sub CmbArrType_Change()

With datatest

    If Tabletype(CmbArrType.ListIndex) = "SUBST" Then  '  Station Class Arrester Coding
        .RecordSource = "SUBST"
        mdl1.ArrTableName = "SUBST"
        
    ElseIf Tabletype(CmbArrType.ListIndex) = "DIST" Then '  dISTRIBUTION Class Arrester Coding
       .RecordSource = "DIST"
        mdl1.ArrTableName = "DIST"
        
    End If
    
    .Refresh
    
End With
cmdchng.Enabled = False
framtprmt.Enabled = False
txtuc.Text = ""
txtiref.Text = ""
txticaccp.Text = ""
txturt.Text = ""
txturefn.Text = ""
txturefx.Text = ""
txturesn.Text = ""
txturesx.Text = ""


cmdchng.Enabled = False
End Sub

Private Sub CmbArrType_Click()


With datatest

    If Tabletype(CmbArrType.ListIndex) = "SUBST" Then  '  Station Class Arrester Coding
        .RecordSource = "SUBST"
        mdl1.ArrTableName = "SUBST"
        
    ElseIf Tabletype(CmbArrType.ListIndex) = "DIST" Then '  dISTRIBUTION Class Arrester Coding
       .RecordSource = "DIST"
        mdl1.ArrTableName = "DIST"
        
    End If
    
    .Refresh
    
End With
cmdchng.Enabled = False
framtprmt.Enabled = False
txtuc.Text = ""
txtiref.Text = ""
txticaccp.Text = ""
txturt.Text = ""
txturefn.Text = ""
txturefx.Text = ""
txturesn.Text = ""
txturesx.Text = ""


End Sub

Private Sub cmdchng_Click()
Dim exist, allow As Boolean
Dim testcode As Double
frmchnginf.MousePointer = 11
exist = False
allow = True
'//////////////////////////////

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
If (txturt.Text = " " Or txturt.Text = "" Or Val(txturt.Text) = 0 Or Val(txturt.Text) > 1200) And allow = True Then
    allow = False
    rtv = MsgBox("Insert Currect Value For Urated", vbOKOnly + vbExclamation, "ERROR")
    txturt.Text = ""
    txturt.SetFocus
End If
If allow = True And Val(txturt.Text) <= Val(txtuc.Text) Then
     allow = False
    rtv = MsgBox("Insert Currect Value For Uc", vbOKOnly + vbExclamation, "ERROR")
    txtuc.Text = ""
    txtuc.SetFocus
End If
'/////////////////////////////////////////////////////////////////
If (txturefn.Text = " " Or txturefn.Text = "" Or Val(txturefn.Text) = 0 Or Val(txturefn.Text) > 1200) And allow = True Then
    allow = False
    rtv = MsgBox("Insert Currect For Minimum Reference Voltage", vbOKOnly + vbExclamation, "ERROR")
    txturefn.Text = ""
    txturefn.SetFocus
End If
If (txturefx.Text = " " Or txturefx.Text = "" Or Val(txturefx.Text) = 0 Or Val(txturefn.Text) > 1200) And allow = True Then
    allow = False
    rtv = MsgBox("Insert Value For Maximum Reference Voltage", vbOKOnly + vbExclamation, "ERROR")
    txturefx.Text = ""
    txturefx.SetFocus
End If
If (Val(txturefn.Text) >= Val(txturefx.Text)) And allow = True Then
    allow = False
    rtv = MsgBox("Insert Currect Value For Minimum Reference Voltage and Maximum Reference Voltage", vbOKOnly + vbExclamation, "ERROR")
    txturefn.Text = ""
    txturefx.Text = ""
    txturefn.SetFocus
End If


If (txturesn.Text = " " Or txturesn.Text = "" Or Val(txturesn.Text) = 0 Or Val(txturesn.Text) > 1200) And allow = True Then
    allow = False
    rtv = MsgBox("Insert Currect Value For Minimum Residual Voltage", vbOKOnly + vbExclamation, "ERROR")
    txturesn.Text = ""
    txturesn.SetFocus
End If
If (txturesx.Text = " " Or txturesx.Text = "" Or Val(txturesx.Text) = 0 Or Val(txturesx.Text) > 1200) And allow = True Then
    allow = False
    rtv = MsgBox("Insert Value For Maximum Residual Voltage", vbOKOnly + vbExclamation, "ERROR")
    txturesx.Text = ""
    txturesx.SetFocus
End If
If Val(txturesn.Text) >= Val(txturesx.Text) And allow = True Then
    allow = False
    rtv = MsgBox("Insert Currect Value For Minimum Residual Voltage and Maximum Residual Voltage", vbOKOnly + vbExclamation, "ERROR")
    txturesn.Text = ""
    txturesx.Text = ""
    txturesn.SetFocus
End If

If (Val(txturefn.Text) >= Val(txturesn.Text)) And allow = True Then
    allow = False
    rtv = MsgBox("Insert Currect Value For Minimum Reference Voltage and Minimum Residual Voltage", vbOKOnly + vbExclamation, "ERROR")
    txturefn.Text = ""
    txturesn.Text = ""
    txturefn.SetFocus
End If
If (Val(txturefx.Text) >= Val(txturesn.Text)) And allow = True Then
    allow = False
    rtv = MsgBox("Insert Currect Value For Maximum Reference Voltage and Maximum residual Voltage", vbOKOnly + vbExclamation, "ERROR")
    txturefx.Text = ""
    txturesx.Text = ""
    txturefx.SetFocus
End If
'///////////////////////////////////
If allow = True Then
    With datatestprmt.Recordset
    If .RecordCount >= 1 Then
        .MoveFirst
        Do While .EOF = False
           If .Fields("Ucon") = Val(txtuc.Text) And .Fields("Urated") = Val(txturt.Text) And .Fields("Iref") = Val(txtiref.Text) And .Fields("Icaccp") = Val(txticaccp.Text) And .Fields("Urefmin") = Val(txturefn.Text) And .Fields("Urefmax") = Val(txturefx.Text) And .Fields("Uresmin") = Val(txturesn.Text) And .Fields("Uresmax") = Val(txturesx.Text) And .Fields("Block_Size1") = blc1 And .Fields("Block_Size2") = blc2 Then
              exist = True
              testcode = .Fields("Id")
              Exit Do
           End If
           .MoveNext
        Loop
    End If
    End With
    If exist = False Then
     With datatestprmt.Recordset
           .addnew
           !Ucon = Val(txtuc.Text)
           !Urated = Val(txturt.Text)
           !iref = Val(txtiref.Text)
           !Icaccp = Val(txticaccp.Text)
           !urefmin = Val(txturefn.Text)
           !urefmax = Val(txturefx.Text)
           !uresmin = Val(txturesn.Text)
           !uresmax = Val(txturesx.Text)
           !Block_Size1 = blc1
           !Block_Size2 = blc2
           .Update
           testcode = .Fields("Id")
     End With
     
    End If
    
    With datatest.Recordset
       For i = 1 To cnt
            .MoveFirst
           .Find "serial_Num=" + Str(sernum(i)), , adSearchForward, 1
           
           !Test_Inf = testcode
           '!Result1 = 2
           '!IcUc1 = 0
           '!pduc1 = 0
           '!pdur1 = 0
           '!Urefiref1 = 0
           .Update
       Next i
    End With
    
    txtfrom.Text = ""
    txtto.Text = ""
    txtfrom.Text = ""
    txtiref.Text = ""
    txtuc.Text = ""
    txticaccp.Text = ""
    txturt.Text = ""
    txturefn.Text = ""
    txturefx.Text = ""
    txturesn.Text = ""
    txturesx.Text = ""
    framtprmt.Enabled = False
    cmdchng.Enabled = False
End If
frmchnginf.MousePointer = 0
End Sub

Private Sub cmdend_Click()
Unload Me
FrmInterface.Show
End Sub

Private Sub Form_Activate()
txtfrom.Text = ""
txtto.Text = ""
txtfrom.Text = ""
txtiref.Text = ""
txtuc.Text = ""
txticaccp.Text = ""
txturt.Text = ""
txturefn.Text = ""
txturefx.Text = ""
txturesn.Text = ""
txturesx.Text = ""
framtprmt.Enabled = False
cmdchng.Enabled = False
End Sub

Private Sub Form_Load()
Dim w1, h1 As Single
'////////////////////////////////////////////
frmchnginf.icon = LoadPicture(mdl1.icon)
With datatest
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.maindb + ";Jet OLEDB:Database Password=" + mdl1.MainDBPas
    .CommandType = adCmdTable
    .RecordSource = "DIST"
    .Refresh
    
End With

With datatestprmt
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.acldb + ";Jet OLEDB:Database Password=" + mdl1.AclDBPas
    .CommandType = adCmdTable
    .RecordSource = "testprmt"
    .Refresh
End With

With DataArrtype
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.acldb + ";Jet OLEDB:Database Password=" + mdl1.AclDBPas
    .CommandType = adCmdTable
    .RecordSource = "Type"
    .Refresh
End With

With DataArrtype.Recordset
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

If Tabletype(CmbArrType.ListIndex) = "SUBST" Then  '  Station Class Arrester Coding
    datatest.RecordSource = "SUBST"
    mdl1.ArrTableName = "SUBST"
ElseIf Tabletype(CmbArrType.ListIndex) = "DIST" Then '  dISTRIBUTION Class Arrester Coding
    datatest.RecordSource = "DIST"
    mdl1.ArrTableName = "DIST"
End If









cmdchng.Enabled = False
framtprmt.Enabled = False


'/////////////////////////////////////
With frmchnginf
    .Height = Screen.Height * 0.7
    .Width = Screen.Width * 0.6
    .Top = Screen.Height * 0.14
    .Left = Screen.Width * 0.14
    .ScaleHeight = .Height * 1.06
    .ScaleWidth = .Width * 1.06
    w1 = .Width
    h1 = .Height
End With
'With lbl1
    '.Top = h1 * 0.07
    '.Left = w1 * 0.11
    '.Height = w1 * 0.08
    '.Width = w1 * 0.45
'End With

With lbl2
    .Top = h1 * 0.11
    .Left = w1 * 0.68
    .Height = w1 * 0.04
    .Width = w1 * 0.35
End With
With lblfrom
    .Top = h1 * 0.19
    .Left = w1 * 0.68
    .Height = w1 * 0.03
    .Width = w1 * 0.15
End With
With txtfrom
    .Top = h1 * 0.17
    .Left = w1 * 0.74
    .Height = w1 * 0.03
    .Width = w1 * 0.13
End With

With lblto
    .Top = h1 * 0.26
    .Left = w1 * 0.68
    .Height = w1 * 0.03
    .Width = w1 * 0.15
End With
With txtto
    .Top = h1 * 0.25
    .Left = w1 * 0.74
    .Height = w1 * 0.03
    .Width = w1 * 0.13
End With

With lblarrtype
    .Top = h1 * 0.18
    .Left = w1 * 0.12
    .Height = w1 * 0.04
    .Width = w1 * 0.15
End With
With CmbArrType
    .Top = h1 * 0.18
    .Left = w1 * 0.25
    .Width = w1 * 0.15
End With

With lblOrderNo
    .Top = h1 * 0.27
    .Left = w1 * 0.12
    .Height = w1 * 0.04
    .Width = w1 * 0.15
    
End With
With txtorderNo
    .Top = h1 * 0.27
    .Left = w1 * 0.25
    .Width = w1 * 0.15
    .Height = w1 * 0.03
End With




With cmdchng
    .Top = h1 * 0.92
    .Left = w1 * 0.2
    .Height = w1 * 0.06
    .Width = w1 * 0.18
End With

With cmdend
    .Top = h1 * 0.92
    .Left = w1 * 0.6
    .Height = w1 * 0.06
    .Width = w1 * 0.18
End With

With framtprmt
    .Top = frmchnginf.Height * 0.35
    .Left = frmchnginf.Width * 0.11
    .Height = frmchnginf.Height * 0.5
    .Width = frmchnginf.Width * 0.85
    .Enabled = False
End With
w1 = framtprmt.Width
h1 = framtprmt.Height
With lbluc
    .Top = h1 * 1 / 6 '+ (1 / 8) * h1
    .Left = w1 / 24
    .Height = h1 / 6
    .Width = w1 / 6
End With
With lblurt
    .Top = h1 * 2 / 6 '+ (1 / 8) * h1
    .Left = w1 / 18
    .Height = h1 / 6
    .Width = w1 * 0.15
End With
With lbliref
    .Top = h1 * 3 / 6 '+ (1 / 8) * h1
    .Left = w1 / 24
    .Height = h1 / 6
    .Width = w1 / 6
End With
With lblic
    .Top = h1 * 4 / 6 '+ (1 / 8) * h1
    .Left = w1 / 35
    .Height = h1 * 0.1
    .Width = w1 * 0.1
End With
With txtuc
    .Top = h1 * 1 / 6 '+ (1 / 8) * h1
    .Left = w1 / 6.5
    .Height = h1 / 10
    .Width = w1 / 6
End With
With txturt
    .Top = h1 * 2 / 6 '+ (1 / 8) * h1
    .Left = w1 / 6.5
    .Height = h1 / 10
    .Width = w1 / 6
End With
With txtiref
    .Top = h1 * 3 / 6 '+ (1 / 8) * h1
    .Left = w1 / 6.5
    .Height = h1 / 10
    .Width = w1 / 6
End With
With txticaccp
    .Top = h1 * 4 / 6 '+ (1 / 8) * h1
    .Left = w1 / 6.5
    .Height = h1 / 10
    .Width = w1 / 6
End With
With lblu1
    .Top = h1 * 1 / 6 '+ (1 / 8) * h1
    .Left = w1 * 0.35
    .Height = h1 / 10
    .Width = w1 / 12
End With
With lblu2
    .Top = h1 * 2 / 6 '+ (1 / 8) * h1
    .Left = w1 * 0.35
    .Height = h1 / 10
    .Width = w1 / 12
End With
With lblu3
    .Top = h1 * 3 / 6 '+ (1 / 8) * h1
    .Left = w1 * 0.35
    .Height = h1 / 10
    .Width = w1 / 12
End With
With lblu4
    .Top = h1 * 4 / 6 ' + (1 / 8) * h1
    .Left = w1 * 0.35
    .Height = h1 / 10
    .Width = w1 / 12
End With
With lblurefn
    .Top = h1 * 1 / 6 '+ (1 / 8) * h1
    .Left = w1 * 0.52
    .Height = h1 / 10
    .Width = w1 / 7
End With
With lblurefx
    .Top = h1 * 2 / 6 '+ (1 / 8) * h1
    .Left = w1 * 0.52
    .Height = h1 / 10
    .Width = w1 / 7
End With
With lbluresn
    .Top = h1 * 3 / 6 '+ (1 / 8) * h1
    .Left = w1 * 0.52
    .Height = h1 / 10
    .Width = w1 / 7
End With
With lbluresx
    .Top = h1 * 4 / 6 '+ (1 / 8) * h1
    .Left = w1 * 0.52
    .Height = h1 / 10
    .Width = w1 / 6
End With

With txturefn
    .Top = h1 * 1 / 6 '+ (1 / 8) * h1
    .Left = w1 * 0.68
    .Height = h1 / 10
    .Width = w1 / 7
End With
With txturefx
    .Top = h1 * 2 / 6 '+ (1 / 8) * h1
    .Left = w1 * 0.68
    .Height = h1 / 10
    .Width = w1 / 7
End With
With txturesn
    .Top = h1 * 3 / 6 ' + (1 / 8) * h1
    .Left = w1 * 0.68
    .Height = h1 / 10
    .Width = w1 / 7
End With
With txturesx
    .Top = h1 * 4 / 6 '+ (1 / 8) * h1
    .Left = w1 * 0.68
    .Height = h1 / 10
    .Width = w1 / 7
End With

With lblu5
    .Top = h1 * 1 / 6 '+ (1 / 8) * h1
    .Left = w1 * 0.85
    .Height = h1 / 10
    .Width = w1 * 0.07
End With
With lblu6
    .Top = h1 * 2 / 6 ' + (1 / 8) * h1
    .Left = w1 * 0.85
    .Height = h1 / 10
    .Width = w1 * 0.07
End With
With lblu7
    .Top = h1 * 3 / 6 '+ (1 / 8) * h1
    .Left = w1 * 0.85
    .Height = h1 / 10
    .Width = w1 * 0.07
End With
With lblu8
    .Top = h1 * 4 / 6 '+ (1 / 8) * h1
    .Left = w1 * 0.85
    .Height = h1 / 10
    .Width = w1 * 0.07
End With
txtfrom.Text = ""
txtiref.Text = ""
txtuc.Text = ""
txticaccp.Text = ""
txturt.Text = ""
txturefn.Text = ""
txturefx.Text = ""
txturesn.Text = ""
txturesx.Text = ""

End Sub




Private Sub txtfrom_Change()
cmdchng.Enabled = False
framtprmt.Enabled = False
txtuc.Text = ""
txtiref.Text = ""
txtuc.Text = ""
txticaccp.Text = ""
txturt.Text = ""
txturefn.Text = ""
txturefx.Text = ""
txturesn.Text = ""
txturesx.Text = ""

End Sub

Private Sub txtfrom_KeyPress(KeyAscii As Integer)
KeyAscii = 0
MsgBox "Insert Order Number.", vbExclamation, "ERROR...."
txtorderNo.SetFocus
  
End Sub


Private Sub txticaccp_KeyPress(KeyAscii As Integer)
KeyAscii = mdl1.inputtest(KeyAscii)
End Sub


Private Sub txtiref_KeyPress(KeyAscii As Integer)
KeyAscii = mdl1.inputtest(KeyAscii)
End Sub


Private Sub txtorderNo_Change()
cmdchng.Enabled = False
framtprmt.Enabled = False
txtuc.Text = ""
txtiref.Text = ""
txtuc.Text = ""
txticaccp.Text = ""
txturt.Text = ""
txturefn.Text = ""
txturefx.Text = ""
txturesn.Text = ""
txturesx.Text = ""

End Sub

Private Sub txtorderNo_KeyPress(KeyAscii As Integer)
KeyAscii = mdl1.inputtest(KeyAscii)
End Sub


Private Sub txtOrderNo_LostFocus()
Dim allow, ended, exiti As Boolean
Dim prmtcode As Long

' Find firs & Last Serial Number
With datatestprmt
    .RecordSource = "production"
    .Refresh
End With


With datatestprmt.Recordset
    .Find "OrderNo=" + Str(Val(txtorderNo.Text)), , adSearchForward, 1
    If .EOF = True Then
        MsgBox "Selected Order Number hasn't been found." + Chr(13) + "Check Arrester Type and Order Number Again.", vbExclamation, "ERROR...."
        txtorderNo.Text = ""
    Else
        txtfrom.Text = Str(.Fields("FirstSN"))
        txtto.Text = Str(.Fields("LastSN"))
    End If

End With



' check serial Numbers

allow = True
cnt = 0
cmdchng.Enabled = False
frmchnginf.MousePointer = 11

With datatestprmt
    .RecordSource = "testprmt"
    .Refresh
End With

If (txtfrom.Text = " " Or txtfrom.Text = "" Or Val(txtfrom.Text) = 0) And allow = True Then
    allow = False
    rtv = MsgBox("Insert First Serial Number.", vbOKOnly + vbExclamation, "ERROR")
    txtfrom.Text = ""
Else
   
    With datatest.Recordset
        If datatest.RecordSource = "SUBST" Then
            .MoveFirst
            .Find "Serial_Num=" + Str(Val(txtfrom.Text)), , adSearchForward, 1
        Else
            .MoveFirst
            .Find "Serial_Num=" + txtfrom.Text, , adSearchForward, 1
        End If
        
        If .EOF = True Then
            framtprmt.Enabled = False
            allow = False
            txtfrom.Text = ""
            CmbArrType.SetFocus
            MsgBox "Selected Serial Numbers havn't been found." + Chr(13) + "Check Arrester Type and Serial Number Again.", vbExclamation, "ERROR...."
        Else
            prmtcode = .Fields("Test_Inf")
        End If
    End With
End If

If (txtto.Text = " " Or txtto.Text = "" Or Val(txtto.Text) = 0) And allow = True Then
    allow = False
    rtv = MsgBox("Insert Final Serial Number.", vbOKOnly + vbExclamation, "ERROR")
    framtprmt.Enabled = False
    txtto.Text = ""
End If
If Val(txtfrom.Text) > Val(txtto.Text) Then
    MsgBox "Insert Currect Value For First And final Arrester." + Chr(13) + " Must Be (First Value < Final Value) .", vbExclamation, "Error"
    allow = False
    txtfrom.Text = ""
    txtto.Text = ""
End If
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

If allow = True And datatest.RecordSource <> "SUBST" Then
    cnt = 0
    ReDim sernum(1 To ((Val(txtto.Text) - Val(txtfrom.Text)) + 1)) As Double
    For sn = Val(txtfrom.Text) To Val(txtto.Text)
         '            Find Serial number
        If sn = Val(txtfrom.Text) Then
            datatest.Recordset.MoveFirst
            datatest.Recordset.Find "Serial_Num=" + Str(sn), , adSearchForward, 1
        Else
            datatest.Recordset.MoveFirst
            datatest.Recordset.Find "Serial_Num=" + Str(sn), , adSearchForward, 1
        End If
   '                   Check Arrester information
        If datatest.Recordset.EOF = True Then
            MsgBox " Serial Number" + Str(sn) + " didn't Found." + Chr(13) + "You cann't chenge parameters of serial numbers : " + Str(Val(txtfrom.Text)) + " TO " + Str(sn - 1), vbOKOnly + vbExclamation, "ERROR"
            txtto.Text = Str(sn - 1)
            Exit For
           
        Else
            If prmtcode <> datatest.Recordset.Fields("Test_Inf") Then
                MsgBox " Serial Number Of " + Str(sn) + " has a different test Parameters." + Chr(13) + "You can chenge parameters of arresters from: " + Str(Val(txtfrom.Text)) + " TO " + Str(sn - 1) + " .", vbExclamation, "Error"
                txtto.Text = Str(sn - 1)
                Exit For
             Else
                cnt = cnt + 1
                sernum(cnt) = sn
            End If
        End If
    Next sn
    datatest.Recordset.MoveFirst
    datatest.Recordset.Find "Serial_Num=" + Str(Val(txtfrom.Text)), , adSearchForward, 1
    
    ReDim Preserve sernum(1 To cnt) As Double
    
    With datatestprmt.Recordset
        .MoveFirst
        .Find "ID=" + Str(datatest.Recordset.Fields("Test_Inf")), , adSearchForward, 1
        If .EOF = False Then
            framtprmt.Enabled = True
            txtuc.Text = Str(.Fields("Ucon"))
            txturt.Text = Str(.Fields("Urated"))
            txtiref.Text = Str(.Fields("iref"))
            txticaccp.Text = Str(.Fields("Icaccp"))
            txturefn.Text = Str(.Fields("urefmin"))
            txturefx.Text = Str(.Fields("urefmax"))
            txturesn.Text = Str(.Fields("uresmin"))
            txturesx.Text = Str(.Fields("uresmax"))
            blc1 = .Fields("Block_Size1")
            blc2 = .Fields("Block_Size2")
        Else
            MsgBox "Parameters of selected serial numbers havn't been found." + Chr(13) + "Check arrester type and serial number again.", vbExclamation, "ERROR...."
        End If
      End With
      cmdchng.Enabled = True
'////////////////////////////////////////////////////////////////////////////
      
ElseIf allow = True And datatest.RecordSource = "SUBST" Then
   cnt = 0
   exiti = False
    For sn = Val(txtfrom.Text) To Val(txtto.Text)
        ended = False
        fc = 0
        Do While ended = False
            If fc = 0 Then
                datatest.Recordset.MoveFirst
                datatest.Recordset.Find "Serial_Num=" + Str(sn), , adSearchForward, 1
                If datatest.Recordset.EOF = True Then
                    exiti = True
                    MsgBox " Serial Number" + Str(sn) + " didn't Found." + Chr(13) + "You can chenge parameters of arresters from: " + Str(Val(txtfrom.Text)) + " TO " + Str(sn - 1), vbOKOnly + vbExclamation, "ERROR"
                    txtto.Text = Str(sn - 1)
                    ended = True
                End If
            Else
                datatest.Recordset.MoveFirst
                datatest.Recordset.Find "Serial_Num=" + Str(sn + 1000 * fc), , adSearchForward, 1
                If datatest.Recordset.EOF = True Then ended = True
                  
            End If
            
            If ended = False Then
                If prmtcode <> datatest.Recordset.Fields("Test_Inf") Then
                    MsgBox " Serial Number Of " + Str(sn) + " has a different Test Parameter." + Chr(13) + "You Can chenge Information Of Arrester " + Str(Val(txtfrom.Text)) + " TO " + Str(sn - 1) + " .", vbExclamation, "Error"
                    txtto.Text = Str(sn - 1)
                    exiti = True
                Else
                    cnt = cnt + 1
                    ReDim Preserve sernum(1 To cnt) As Double
                    sernum(cnt) = sn + 1000 * fc
               End If
            End If
        
        fc = fc + 1
        Loop
        If exiti = True Then Exit For
    Next sn
    datatest.Recordset.MoveFirst
    datatest.Recordset.Find "Serial_Num=" + Str(sernum(1)), , adSearchForward, 1
      With datatestprmt.Recordset
        .MoveFirst
        .Find "ID=" + Str(datatest.Recordset.Fields("Test_Inf")), , adSearchForward, 1
        framtprmt.Enabled = True
        txtuc.Text = Str(.Fields("Ucon"))
        txturt.Text = Str(.Fields("Urated"))
        txtiref.Text = Str(.Fields("iref"))
        txticaccp.Text = Str(.Fields("Icaccp"))
        txturefn.Text = Str(.Fields("urefmin"))
        txturefx.Text = Str(.Fields("urefmax"))
        txturesn.Text = Str(.Fields("uresmin"))
        txturesx.Text = Str(.Fields("uresmax"))
        blc1 = .Fields("Block_Size1")
        blc2 = .Fields("Block_Size2")
      End With
      cmdchng.Enabled = True

End If
frmchnginf.MousePointer = 0

End Sub

Private Sub txtto_Change()
cmdchng.Enabled = False
framtprmt.Enabled = False
txtuc.Text = ""
txtiref.Text = ""
txticaccp.Text = ""
txturt.Text = ""
txturefn.Text = ""
txturefx.Text = ""
txturesn.Text = ""
txturesx.Text = ""
End Sub

Private Sub txtto_KeyPress(KeyAscii As Integer)
KeyAscii = 0
MsgBox "Insert Order Number.", vbExclamation, "ERROR...."
txtorderNo.SetFocus
  
End Sub


Private Sub txtuc_KeyPress(KeyAscii As Integer)
KeyAscii = mdl1.inputtest(KeyAscii)
End Sub


Private Sub txturefn_KeyPress(KeyAscii As Integer)
KeyAscii = mdl1.inputtest(KeyAscii)
End Sub


Private Sub txturefx_KeyPress(KeyAscii As Integer)
KeyAscii = mdl1.inputtest(KeyAscii)
End Sub


Private Sub txturesn_KeyPress(KeyAscii As Integer)
KeyAscii = mdl1.inputtest(KeyAscii)
End Sub


Private Sub txturesx_KeyPress(KeyAscii As Integer)
KeyAscii = mdl1.inputtest(KeyAscii)
End Sub


Private Sub txturt_KeyPress(KeyAscii As Integer)
KeyAscii = mdl1.inputtest(KeyAscii)
End Sub


