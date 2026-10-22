VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmaddblc 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADD NEW BLOCK INFORMATION"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12555
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9255
   ScaleWidth      =   12555
   Begin MSAdodcLib.Adodc datablock 
      Height          =   375
      Left            =   4320
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
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
      Caption         =   "Adodc1"
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
   Begin MSFlexGridLib.MSFlexGrid fgblock 
      Height          =   4455
      Left            =   240
      TabIndex        =   32
      Top             =   720
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   7858
      _Version        =   393216
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "Add"
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
      Left            =   10800
      TabIndex        =   10
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Frame fram1 
      Caption         =   "New Block Information"
      Height          =   3135
      Left            =   240
      TabIndex        =   12
      Top             =   5160
      Width           =   9975
      Begin VB.TextBox txtname 
         Height          =   285
         Left            =   6720
         TabIndex        =   9
         Text            =   " "
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtpc 
         Height          =   285
         Left            =   6720
         TabIndex        =   8
         Text            =   " "
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtthic 
         Height          =   285
         Left            =   3600
         TabIndex        =   7
         Text            =   " "
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtdim 
         Height          =   285
         Left            =   3720
         TabIndex        =   6
         Text            =   " "
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtiref 
         Height          =   285
         Left            =   3720
         TabIndex        =   5
         Text            =   " "
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtinom 
         Height          =   285
         Left            =   3600
         TabIndex        =   4
         Text            =   " "
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtures 
         Height          =   285
         Left            =   840
         TabIndex        =   3
         Text            =   " "
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txturef 
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Text            =   " "
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtur 
         Height          =   285
         Left            =   840
         TabIndex        =   1
         Text            =   " "
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtuc 
         Height          =   285
         Left            =   720
         TabIndex        =   0
         Text            =   " "
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblunit 
         Caption         =   "mW"
         Height          =   255
         Index           =   8
         Left            =   8400
         TabIndex        =   31
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblunit 
         Caption         =   "mm"
         Height          =   255
         Index           =   7
         Left            =   5280
         TabIndex        =   30
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label lblunit 
         Caption         =   "mm"
         Height          =   255
         Index           =   6
         Left            =   5280
         TabIndex        =   29
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label lblunit 
         Caption         =   "mA"
         Height          =   255
         Index           =   5
         Left            =   5280
         TabIndex        =   28
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblunit 
         Caption         =   "KA"
         Height          =   255
         Index           =   4
         Left            =   5160
         TabIndex        =   27
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblunit 
         Caption         =   "KV"
         Height          =   255
         Index           =   3
         Left            =   2280
         TabIndex        =   26
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblunit 
         Caption         =   "KV"
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   25
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label lblunit 
         Caption         =   "KV"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   24
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblunit 
         Caption         =   "KV"
         Height          =   255
         Index           =   0
         Left            =   2400
         TabIndex        =   23
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblname 
         Caption         =   "Manufacturer"
         Height          =   255
         Left            =   5640
         TabIndex        =   22
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblpc 
         Caption         =   "Power Diss"
         Height          =   255
         Left            =   5760
         TabIndex        =   21
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblthic 
         Caption         =   "Thickness"
         Height          =   255
         Left            =   2760
         TabIndex        =   20
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lbldim 
         Caption         =   "Diameter"
         Height          =   255
         Left            =   3000
         TabIndex        =   19
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label lbliref 
         Caption         =   "Iref"
         Height          =   255
         Left            =   3000
         TabIndex        =   18
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblinom 
         Caption         =   "Inom"
         Height          =   255
         Left            =   3120
         TabIndex        =   17
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblures 
         Caption         =   "Ures"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lbluref 
         Caption         =   "Uref"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblur 
         Caption         =   "Urated"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lbluc 
         Caption         =   "Ucon"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   735
      End
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
      Height          =   375
      Left            =   10800
      TabIndex        =   11
      Top             =   7560
      Width           =   1215
   End
End
Attribute VB_Name = "frmaddblc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim norow As Long
Private Sub cmdadd_Click()
Dim s1 As String
Dim allow As Boolean
Dim rtv As Integer
allow = True
If allow = True And (txtuc.Text = " " Or txtuc.Text = "" Or Val(txtuc.Text) = 0 Or Val(txtuc.Text) > 100) Then
    allow = False
    rtv = MsgBox("Insert Currect Value For Ucon", vbOKOnly + vbExclamation, "ERROR")
    txtuc.Text = ""
    txtuc.SetFocus
    allow = False
End If
If allow = True And (txtur.Text = " " Or txtur.Text = "" Or Val(txtur.Text) = 0 Or Val(txtur.Text) > 100) Then
    allow = False
    rtv = MsgBox("Insert Currect Value For Urated", vbOKOnly + vbExclamation, "ERROR")
    txtur.Text = ""
    txtur.SetFocus
    allow = False
End If
If allow = True And (txturef.Text = " " Or txturef.Text = "" Or Val(txturef.Text) = 0 Or Val(txturef.Text) > 100) Then
    allow = False
    rtv = MsgBox("Insert Currect Value For Uref", vbOKOnly + vbExclamation, "ERROR")
    txturef.Text = ""
    txturef.SetFocus
    allow = False
End If
If allow = True And (txtures.Text = " " Or txtures.Text = "" Or Val(txtures.Text) = 0 Or Val(txtures.Text) > 100 Or Val(txtures.Text) < Val(txturef.Text)) Then
    allow = False
    rtv = MsgBox("Insert Currect Value For Ures", vbOKOnly + vbExclamation, "ERROR")
    txtures.Text = ""
    txtures.SetFocus
    allow = False
End If
If allow = True And (txtinom.Text = " " Or txtinom.Text = "" Or Val(txtinom.Text) = 0 Or Val(txtinom.Text) > 50) Then
    allow = False
    rtv = MsgBox("Insert Currect Value For Inom", vbOKOnly + vbExclamation, "ERROR")
    txtinom.Text = ""
    txtinom.SetFocus
    allow = False
End If
If allow = True And (txtiref.Text = " " Or txtiref.Text = "" Or Val(txtiref.Text) = 0 Or Val(txtiref.Text) > 100) Then
    allow = False
    rtv = MsgBox("Insert Currect Value For Iref", vbOKOnly + vbExclamation, "ERROR")
    txtiref.Text = ""
    txtiref.SetFocus
    allow = False
End If
If allow = True And (txtdim.Text = " " Or txtdim.Text = "" Or Val(txtdim.Text) = 0 Or Val(txtdim.Text) > 100) Then
    allow = False
    rtv = MsgBox("Insert Currect Value For Diameter", vbOKOnly + vbExclamation, "ERROR")
    txtdim.Text = ""
    txtdim.SetFocus
    allow = False
End If
If allow = True And (txtthic.Text = " " Or txtthic.Text = "" Or Val(txtthic.Text) = 0 Or Val(txtthic.Text) > 100) Then
    allow = False
    rtv = MsgBox("Insert Currect Value For Thickness", vbOKOnly + vbExclamation, "ERROR")
    txtthic.Text = ""
    txtthic.SetFocus
    allow = False
End If
If allow = True And (txtpc.Text = " " Or txtpc.Text = "" Or Val(txtpc.Text) = 0 Or Val(txtpc.Text) > 1000) Then
    allow = False
    rtv = MsgBox("Insert Currect Value For Power Dissipation", vbOKOnly + vbExclamation, "ERROR")
    txtpc.Text = ""
    txtpc.SetFocus
    allow = False
End If
If allow = True And (txtname.Text = " " Or txtname.Text = "") Then
    allow = False
    rtv = MsgBox("Insert Block Name", vbOKOnly + vbExclamation, "ERROR")
    txtname.Text = ""
    txtname.SetFocus
End If
If allow = True Then
  rtv = MsgBox("Do You Want To Add New Block data to Database?", vbYesNo + vbDefaultButton2, "Attention")
  If rtv = 6 Then
    With datablock.Recordset
        '.MoveFirst
        .addnew
        !Name = txtname.Text
        !ures = Val(txtures.Text)
        !inom = Val(txtinom.Text)
        !iref = Val(txtiref.Text)
        !uc = Val(txtuc.Text)
        !ur = Val(txtur.Text)
        !uref = Val(txturef.Text)
        !diameter = Val(txtdim.Text)
        !thickness = Val(txtthic.Text)
        !pc = Val(txtpc.Text)
        !code = Left(txtname.Text, Len(txtname.Text)) + Right(Str(Val(txtdim.Text)), 2) + Right(Str(Val(txtthic.Text)), 2)
        .Update
    End With
    datablock.Refresh
    Form_Activate
    txtuc.Text = ""
    txtur.Text = ""
    txturef.Text = ""
    txtures.Text = ""
    txtinom.Text = ""
    txtiref.Text = ""
    txtdim.Text = ""
    txtthic.Text = ""
    txtpc.Text = ""
    txtname.Text = ""
    Call Form_Activate
 End If
End If
End Sub



Private Sub cmdend_Click()
Unload Me
FrmInterface.Show
End Sub

Private Sub Form_Activate()

With datablock
     .Refresh
     norow = .Recordset.RecordCount
End With

txtuc.Text = ""
txtur.Text = ""
txturef.Text = ""
txtures.Text = ""
txtinom.Text = ""
txtiref.Text = ""
txtthic.Text = ""
txtname.Text = ""


With fgblock
     .RowHeightMin = 100
    .Clear
    .Cols = 12
    .ColWidth(0) = .Width * 0.072
    .ColWidth(1) = .Width * 0.072
    .ColWidth(2) = .Width * 0.072
    .ColWidth(3) = .Width * 0.072
    .ColWidth(4) = .Width * 0.072
    .ColWidth(5) = .Width * 0.072
    .ColWidth(6) = .Width * 0.072
    .ColWidth(7) = .Width * 0.072
    .ColWidth(8) = .Width * 0.072
    .ColWidth(9) = .Width * 0.072
    .ColWidth(10) = .Width * 0.12
    .ColWidth(11) = .Width * 0.13
    If norow >= 1 Then
        .Rows = norow + 1
    Else
        .Rows = 2
    End If
    .Row = 0
    .Col = 0
    .Text = "Block ID"
    .Col = 1
    .Text = "Ucon"
    .Col = 2
    .Text = "Urated"
    .Col = 3
    .Text = "Uref"
    .Col = 4
    .Text = "Ures"
    .Col = 5
    .Text = "Inom"
    .Col = 6
    .Text = "Iref"
    .Col = 7
    .Text = "Power Diss"
    .Col = 8
    .Text = "Diameter"
    .Col = 9
    .Text = "Thickness"
    .Col = 10
    .Text = "Name"
    .Col = 11
    .Text = "Code"
End With

With datablock.Recordset

If norow >= 1 Then
    .MoveFirst
    For i = 1 To norow
        fgblock.Row = i
        fgblock.Col = 0
        fgblock.Text = Str(.Fields("ID"))
         fgblock.Col = 1
        fgblock.Text = Str(.Fields("uc"))
        fgblock.Col = 2
        fgblock.Text = Str(.Fields("ur"))
        fgblock.Col = 3
        fgblock.Text = Str(.Fields("uref"))
        fgblock.Col = 4
        fgblock.Text = Str(.Fields("ures"))
        fgblock.Col = 5
        fgblock.Text = Str(.Fields("inom"))
        fgblock.Col = 6
        fgblock.Text = Str(.Fields("iref"))
        fgblock.Col = 7
        fgblock.Text = Str(.Fields("pc"))
        fgblock.Col = 8
        fgblock.Text = Str(.Fields("diameter"))
        fgblock.Col = 9
        fgblock.Text = Str(.Fields("thickness"))
        fgblock.Col = 10
        fgblock.Text = .Fields("name")
        fgblock.Col = 11
        fgblock.Text = .Fields("code")
        .MoveNext
    Next i
End If
End With
'
End Sub

Private Sub Form_Load()
Dim w, h As Double

With datablock
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.acldb + ";Jet OLEDB:Database Password=" + mdl1.AclDBPas
    .CommandType = adCmdTable
    .RecordSource = "block"
    .Refresh
     norow = .Recordset.RecordCount
End With
With frmaddblc
    .Width = Screen.Width * 0.9
    .Height = Screen.Height * 0.95
    .Left = Screen.Width * 0.05
    .Top = 10
End With

With fgblock
    .Width = frmaddblc.Width * 0.9
    .Height = frmaddblc.Height * 0.42
    .Top = frmaddblc.Height * 0.08
    .Left = frmaddblc.Width * 0.05
    '.DataSource = datablock.Recordset
End With
With fram1
    .Width = frmaddblc.Width * 0.7
    .Height = frmaddblc.Height * 0.3
    .Top = frmaddblc.Height * 0.57
    .Left = frmaddblc.Width * 0.05
    w = .Width
    h = .Height
End With
With cmdadd
    .Width = frmaddblc.Width * 0.08
    .Height = frmaddblc.Height * 0.05
    .Top = frmaddblc.Height * 0.63
    .Left = frmaddblc.Width * 0.83
End With
With cmdend
    .Width = frmaddblc.Width * 0.08
    .Height = frmaddblc.Height * 0.05
    .Top = frmaddblc.Height * 0.8
    .Left = frmaddblc.Width * 0.83
End With

frmaddblc.icon = LoadPicture(mdl1.icon)
'////////////////////////////////////fram''''''''
With lbluc
    .Width = w * 0.06
    .Height = h * 0.1
    .Top = h * 0.158
    .Left = w * 0.02
End With
With lblur
    .Width = w * 0.06
    .Height = h * 0.1
    .Top = h * 0.358
    .Left = w * 0.02
End With
With lbluref
    .Width = w * 0.06
    .Height = h * 0.1
    .Top = h * 0.558
    .Left = w * 0.025
End With
With lblures
    .Width = w * 0.06
    .Height = h * 0.1
    .Top = h * 0.758
    .Left = w * 0.025
End With
With txtuc
    .Width = w * 0.13
    .Height = h * 0.1
    .Top = h * 0.15
    .Left = w * 0.08
End With
With txtur
    .Width = w * 0.13
    .Height = h * 0.1
    .Top = h * 0.35
    .Left = w * 0.08
End With
With txturef
    .Width = w * 0.13
    .Height = h * 0.1
    .Top = h * 0.55
    .Left = w * 0.08
End With
With txtures
    .Width = w * 0.13
    .Height = h * 0.1
    .Top = h * 0.75
    .Left = w * 0.08
End With
''''''''''''''''''''''''''''''''''''''''
With lblinom
    .Width = w * 0.06
    .Height = h * 0.1
    .Top = h * 0.158
    .Left = w * 0.35
End With
With lbliref
    .Width = w * 0.06
    .Height = h * 0.1
    .Top = h * 0.358
    .Left = w * 0.358
End With
With lbldim
    .Width = w * 0.06
    .Height = h * 0.1
    .Top = h * 0.558
    .Left = w * 0.33
End With
With lblthic
    .Width = w * 0.06
    .Height = h * 0.1
    .Top = h * 0.758
    .Left = w * 0.33
End With
With txtinom
    .Width = w * 0.13
    .Height = h * 0.1
    .Top = h * 0.15
    .Left = w * 0.4
End With
With txtiref
    .Width = w * 0.13
    .Height = h * 0.1
    .Top = h * 0.35
    .Left = w * 0.4
End With
With txtdim
    .Width = w * 0.13
    .Height = h * 0.1
    .Top = h * 0.55
    .Left = w * 0.4
End With
With txtthic
    .Width = w * 0.13
    .Height = h * 0.1
    .Top = h * 0.75
    .Left = w * 0.4
End With
''''''''''''''''''''''''''''''''''''''''''''
With lblpc
    .Width = w * 0.1
    .Height = h * 0.1
    .Top = h * 0.158
    .Left = w * 0.68
End With
With lblname
    .Width = w * 0.1
    .Height = h * 0.1
    .Top = h * 0.358
    .Left = w * 0.68
End With
With txtpc
    .Width = w * 0.13
    .Height = h * 0.1
    .Top = h * 0.15
    .Left = w * 0.8
End With
With txtname
    .Width = w * 0.13
    .Height = h * 0.1
    .Top = h * 0.35
    .Left = w * 0.8
End With
'/////////////////////////////////////////
With lblunit(0)
    .Width = w * 0.04
    .Height = h * 0.1
    .Top = h * 0.158
    .Left = w * 0.22
End With
With lblunit(1)
    .Width = w * 0.04
    .Height = h * 0.1
    .Top = h * 0.358
    .Left = w * 0.22
End With
With lblunit(2)
    .Width = w * 0.04
    .Height = h * 0.1
    .Top = h * 0.558
    .Left = w * 0.22
End With
With lblunit(3)
    .Width = w * 0.04
    .Height = h * 0.1
    .Top = h * 0.758
    .Left = w * 0.22
End With
''''''''''''''''''''''''
With lblunit(4)
    .Width = w * 0.04
    .Height = h * 0.1
    .Top = h * 0.158
    .Left = w * 0.55
End With
With lblunit(5)
    .Width = w * 0.04
    .Height = h * 0.1
    .Top = h * 0.358
    .Left = w * 0.55
End With
With lblunit(6)
    .Width = w * 0.04
    .Height = h * 0.1
    .Top = h * 0.558
    .Left = w * 0.55
End With
With lblunit(7)
    .Width = w * 0.04
    .Height = h * 0.1
    .Top = h * 0.758
    .Left = w * 0.55
End With
With lblunit(8)
    .Width = w * 0.04
    .Height = h * 0.1
    .Top = h * 0.158
    .Left = w * 0.95
End With

txtuc.Text = ""
txtur.Text = ""
txturef.Text = ""
txtures.Text = ""
txtinom.Text = ""
txtiref.Text = ""
txtthic.Text = ""
txtname.Text = ""

With fgblock
     .RowHeightMin = 100
    .Clear
    .Cols = 12
    .ColWidth(0) = .Width * 0.072
    .ColWidth(1) = .Width * 0.072
    .ColWidth(2) = .Width * 0.072
    .ColWidth(3) = .Width * 0.072
    .ColWidth(4) = .Width * 0.072
    .ColWidth(5) = .Width * 0.072
    .ColWidth(6) = .Width * 0.072
    .ColWidth(7) = .Width * 0.072
    .ColWidth(8) = .Width * 0.072
    .ColWidth(9) = .Width * 0.072
    .ColWidth(10) = .Width * 0.12
    .ColWidth(11) = .Width * 0.13
    If norow >= 1 Then
        .Rows = norow + 1
    Else
        .Rows = 2
    End If
    .Row = 0
    .Col = 0
    .Text = "Block ID"
    .Col = 1
    .Text = "Ucon"
    .Col = 2
    .Text = "Urated"
    .Col = 3
    .Text = "Uref"
    .Col = 4
    .Text = "Ures"
    .Col = 5
    .Text = "Inom"
    .Col = 6
    .Text = "Iref"
    .Col = 7
    .Text = "Power Diss"
    .Col = 8
    .Text = "Diameter"
    .Col = 9
    .Text = "Thickness"
    .Col = 10
    .Text = "Name"
    .Col = 11
    .Text = "Code"
End With

With datablock.Recordset
If norow >= 1 Then
    .MoveFirst

    For i = 1 To norow
        fgblock.Row = i
        fgblock.Col = 0
        fgblock.Text = Str(.Fields("ID"))
        fgblock.Col = 1
        fgblock.Text = Str(.Fields("uc"))
        fgblock.Col = 2
        fgblock.Text = Str(.Fields("ur"))
        fgblock.Col = 3
        fgblock.Text = Str(.Fields("uref"))
        fgblock.Col = 4
        fgblock.Text = Str(.Fields("ures"))
        fgblock.Col = 5
        fgblock.Text = Str(.Fields("inom"))
        fgblock.Col = 6
        fgblock.Text = Str(.Fields("iref"))
        fgblock.Col = 7
        fgblock.Text = Str(.Fields("pc"))
        fgblock.Col = 8
        fgblock.Text = Str(.Fields("diameter"))
        fgblock.Col = 9
        fgblock.Text = Str(.Fields("thickness"))
        fgblock.Col = 10
        fgblock.Text = .Fields("name")
        fgblock.Col = 11
        fgblock.Text = .Fields("code")
        .MoveNext
    Next i
End If
End With

End Sub



Private Sub txtdim_KeyPress(KeyAscii As Integer)
KeyAscii = mdl1.inputtest(KeyAscii)
End Sub


Private Sub txtinom_KeyPress(KeyAscii As Integer)
KeyAscii = mdl1.inputtest(KeyAscii)
End Sub


Private Sub txtiref_KeyPress(KeyAscii As Integer)
KeyAscii = mdl1.inputtest(KeyAscii)
End Sub



Private Sub txtname_LostFocus()
txtname.Text = UCase(txtname.Text)
End Sub

Private Sub txtpc_KeyPress(KeyAscii As Integer)
KeyAscii = mdl1.inputtest(KeyAscii)
End Sub


Private Sub txtthic_KeyPress(KeyAscii As Integer)
KeyAscii = mdl1.inputtest(KeyAscii)
End Sub


Private Sub txtuc_KeyPress(KeyAscii As Integer)
KeyAscii = mdl1.inputtest(KeyAscii)
End Sub


Private Sub txtur_KeyPress(KeyAscii As Integer)
KeyAscii = mdl1.inputtest(KeyAscii)
End Sub


Private Sub txturef_KeyPress(KeyAscii As Integer)
KeyAscii = mdl1.inputtest(KeyAscii)
End Sub


Private Sub txtures_KeyPress(KeyAscii As Integer)
KeyAscii = mdl1.inputtest(KeyAscii)
End Sub


