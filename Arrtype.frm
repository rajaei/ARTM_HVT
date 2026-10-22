VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmarrtype 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADD NEW BLOCK INFORMATION"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12555
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   12555
   Begin MSAdodcLib.Adodc datablock 
      Height          =   375
      Left            =   9480
      Top             =   2880
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
   Begin MSFlexGridLib.MSFlexGrid fgblock 
      Height          =   2415
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   4260
      _Version        =   393216
      BackColor       =   16777215
      BackColorFixed  =   -2147483637
      ForeColorFixed  =   -2147483640
      BackColorSel    =   -2147483638
      ForeColorSel    =   -2147483640
      BackColorBkg    =   12632256
      GridColor       =   0
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
      Left            =   9840
      TabIndex        =   2
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Frame fram1 
      Caption         =   "New  Arrester  Type  And  Code"
      Height          =   2895
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Width           =   8895
      Begin VB.TextBox txtdes 
         Height          =   975
         Left            =   3840
         TabIndex        =   12
         Text            =   "   "
         Top             =   840
         Width           =   4695
      End
      Begin VB.OptionButton optsubst 
         Caption         =   "SubStation Class"
         Height          =   255
         Left            =   1800
         TabIndex        =   10
         Top             =   2280
         Width           =   1815
      End
      Begin VB.OptionButton optdist 
         Caption         =   "Distributaion Class"
         Height          =   375
         Left            =   1800
         TabIndex        =   9
         Top             =   1800
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.TextBox txtarrcode 
         Height          =   285
         Left            =   1920
         TabIndex        =   1
         Text            =   " "
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtarrtype 
         Height          =   285
         Left            =   1920
         TabIndex        =   0
         Text            =   " "
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label LBLDES 
         Caption         =   "Description"
         Height          =   375
         Left            =   3960
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lbltable 
         Caption         =   "Arrester Type"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblur 
         Caption         =   "Arrester Code"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lbluc 
         Caption         =   "Arrester Name"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1215
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
      Left            =   9840
      TabIndex        =   3
      Top             =   5400
      Width           =   1215
   End
End
Attribute VB_Name = "frmarrtype"
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
If allow = True And (txtarrtype.Text = " " Or txtarrtype.Text = "") Then
    allow = False
    rtv = MsgBox("Insert Currect Value For Arrester Type", vbOKOnly + vbExclamation, "ERROR")
    txtarrtype.Text = ""
    txtarrtpye.SetFocus
    allow = False
End If

If allow = True And (txtarrcode.Text = " " Or txtarrcode.Text = "" Or Val(txtarrcode.Text) > 9 Or Val(txtarrcode.Text) < 1) Then
    allow = False
    rtv = MsgBox("Insert Currect Value For Arrester Code. [1-9]", vbOKOnly + vbExclamation, "ERROR")
    txtarrcode.Text = ""
    txtarrcode.SetFocus
    allow = False
End If

If allow = True And (txtdes.Text = " " Or txtdes.Text = "" Or txtdes.Text = Null) Then
    allow = False
    rtv = MsgBox("Insert brief description for arrester type", vbOKOnly + vbExclamation, "ERROR")
    txtdes.Text = " "
    txtdes.SetFocus
    allow = False
End If

If allow = True Then
    datablock.Recordset.Find "Code=" + txtarrcode.Text, , adSearchForward, 1
    If datablock.Recordset.EOF = False Then
        rtv = MsgBox("Selected code has been registerd in dataBase." + Chr(13) + "Insert new value Fhor arrester code", vbOKOnly + vbExclamation, "ERROR")
        allow = False
        txtarrcode.Text = ""
        txtarrcode.SetFocus
    End If
    datablock.Recordset.Find "arr_Type='" + txtarrtype.Text + "'", , adSearchForward, 1
    If datablock.Recordset.EOF = False Then
        rtv = MsgBox("Selected arrester Type Existing In DataBase." + Chr(13) + "Insert Other Value For Arrester Type", vbOKOnly + vbExclamation, "ERROR")
        allow = False
        txtarrtype.Text = ""
        txtarrtype.SetFocus
    End If
End If
If allow = True Then
  rtv = MsgBox("Do You Want To Add New Arrester Type Information to Database?", vbYesNo + vbDefaultButton2, "Attention")
  If rtv = 6 Then
    With datablock.Recordset
        .MoveFirst
        .addnew
        !arr_Type = txtarrtype.Text
        !code = Val(txtarrcode.Text)
        
        If optdist.value = True Then
            !TableName = "DIST"
            !Description = "  " + txtdes.Text
        ElseIf optsubst.value = True Then
             !TableName = "SUBST"
              !Description = "  " + txtdes.Text
        End If
        .Update
    End With
    datablock.Refresh
    Form_Activate
    txtarrtype.Text = ""
    txtarrcode.Text = ""
    txtdes.Text = " "
 End If
End If
End Sub
Private Sub cmdend_Click()
Unload Me
FrmInterface.Show
End Sub

Private Sub Form_Activate()

With frmarrtype.datablock
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.acldb + ";Jet OLEDB:Database Password=" + mdl1.AclDBPas
    .CommandType = adCmdTable
    .RecordSource = "type"
    .Refresh
    .Recordset.MoveFirst
    norow = .Recordset.RecordCount
End With


With fgblock
     .RowHeightMin = 100
    .Clear
    .Cols = 4
    .ColWidth(0) = .Width * 0.15
    .ColWidth(1) = .Width * 0.15
    .ColWidth(2) = .Width * 0.15
    .ColWidth(3) = .Width * 0.55
    If norow >= 1 Then
        .Rows = norow + 1
    Else
        .Rows = 2
    End If
    .Row = 0
    .Col = 0
    .Text = "Arrestr Name"
    .Col = 1
    .Text = "Arrester Code"
    .Col = 2
    .Text = "Arrester Type"
    .Col = 3
    .Text = "Arrester Description"
    
End With

With datablock.Recordset
.MoveFirst
If norow >= 1 Then
    For i = 1 To norow
        fgblock.Row = i
        fgblock.Col = 0
        fgblock.Text = .Fields("arr_Type")
        fgblock.Col = 1
        fgblock.Text = Str(.Fields("Code"))
        fgblock.Col = 2
        fgblock.Text = .Fields("TableName")
        fgblock.Col = 3
        fgblock.Text = "  " + .Fields("Description")
        .MoveNext
    Next i
End If
End With
'
End Sub

Private Sub Form_Load()
Dim w, h As Double



With frmarrtype.datablock
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.acldb + ";Jet OLEDB:Database Password=" + mdl1.AclDBPas
    .CommandType = adCmdTable
    .RecordSource = "type"
    .Refresh
    .Recordset.MoveFirst
    norow = .Recordset.RecordCount
End With


With frmarrtype
    .Width = Screen.Width * 0.7
    .Height = Screen.Height * 0.7
    .Left = Screen.Width * 0.15
    .Top = Screen.Height * 0.1
    .ScaleHeight = .Height
    .ScaleWidth = .Width
    
End With

With fgblock
    .Width = frmarrtype.Width * 0.8
    .Height = frmarrtype.Height * 0.5
    .Top = frmarrtype.Height * 0.05
    .Left = frmarrtype.Width * 0.1
End With
With fram1
    .Width = frmarrtype.Width * 0.55
    .Height = frmarrtype.Height * 0.35
    .Top = frmarrtype.Height * 0.6
    .Left = frmarrtype.Width * 0.1
    w = .Width
    h = .Height
End With
With cmdadd
    .Width = frmarrtype.Width * 0.08
    .Height = frmarrtype.Height * 0.05
    .Top = fram1.Top + fram1.Height * 0.35
    .Left = frmarrtype.Width * 0.75
End With
With cmdend
    .Width = frmarrtype.Width * 0.08
    .Height = frmarrtype.Height * 0.05
    .Top = fram1.Top + fram1.Height * 0.65
    .Left = cmdadd.Left
End With

frmarrtype.icon = LoadPicture(mdl1.icon)
'////////////////////////////////////fram''''''''
With lbluc
    .Width = w * 0.2
    .Height = h * 0.1
    .Top = h * 0.2
    .Left = w * 0.1
End With

With LBLDES
    .Width = w * 0.2
    .Height = h * 0.1
    .Top = h * 0.2
    .Left = w * 0.6
End With
With txtdes
    .Width = w * 0.38
    .Height = h * 0.5
    .Top = h * 0.37
    .Left = w * 0.6
End With

With lblur
    .Width = w * 0.2
    .Height = h * 0.1
    .Top = h * 0.4
    .Left = w * 0.1
End With
With lbltable
    .Width = w * 0.2
    .Height = h * 0.1
    .Top = h * 0.6
    .Left = w * 0.1
End With

With txtarrtype
    .Width = w * 0.2
    .Height = h * 0.1
    .Top = h * 0.17
    .Left = w * 0.33
End With
With txtarrcode
    .Width = w * 0.2
    .Height = h * 0.1
    .Top = h * 0.37
    .Left = w * 0.33
End With

With optdist
    .Width = w * 0.3
    .Height = h * 0.1
    .Top = h * 0.57
    .Left = w * 0.33
    .value = True
End With
With optsubst
    .Width = w * 0.3
    .Height = h * 0.1
    .Top = h * 0.7
    .Left = w * 0.33
End With

''''''''''''''''''''''''''''''''''''''''


With fgblock
     .RowHeightMin = 100
    .Clear
    .Cols = 4
    .ColWidth(0) = .Width * 0.15
    .ColWidth(1) = .Width * 0.15
    .ColWidth(2) = .Width * 0.15
    .ColWidth(3) = .Width * 0.55
    If norow >= 1 Then
        .Rows = norow + 1
    Else
        .Rows = 2
    End If
    .Row = 0
    .Col = 0
    .Text = "Arrestr Name"
    .Col = 1
    .Text = "Arrester Code"
    .Col = 2
    .Text = "Arrester Type"
    .Col = 3
    .Text = "Arrester Description"
    
End With

With datablock.Recordset
.MoveFirst
If norow >= 1 Then
    For i = 1 To norow
        fgblock.Row = i
        fgblock.Col = 0
        fgblock.Text = .Fields("arr_Type")
        fgblock.Col = 1
        fgblock.Text = Str(.Fields("Code"))
        fgblock.Col = 2
        fgblock.Text = .Fields("TableName")
        fgblock.Col = 3
        fgblock.Text = .Fields("Description")
        .MoveNext
    Next i
End If
End With

End Sub

Private Sub txtarrcode_KeyPress(KeyAscii As Integer)
KeyAscii = mdl1.inputtest(KeyAscii)
End Sub
Private Sub txtarrtype_LostFocus()
txtarrtype.Text = UCase(txtarrtype.Text)
End Sub
