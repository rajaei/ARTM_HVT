VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmnewper 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADD NEW PERSONAL INFORMATION"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9915
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   9915
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fram3 
      Caption         =   "User Access "
      Height          =   4935
      Left            =   4440
      TabIndex        =   13
      Top             =   480
      Width           =   5055
      Begin VB.CommandButton cmdapply 
         Caption         =   "Apply"
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
         Left            =   480
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   4080
         Width           =   1575
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
         Left            =   3240
         TabIndex        =   20
         Top             =   4080
         Width           =   1455
      End
      Begin VB.CheckBox chkAccess 
         Caption         =   "Production Personnel (Perform Production Process)"
         Height          =   255
         Index           =   6
         Left            =   150
         TabIndex        =   19
         Top             =   3000
         Width           =   4575
      End
      Begin VB.CheckBox chkAccess 
         Caption         =   "HV Test Operator ( Perform HV Test)"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   18
         Top             =   2520
         Width           =   3015
      End
      Begin VB.CheckBox chkAccess 
         Caption         =   "QC Operator (Register Arrester and production Data)"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   17
         Top             =   2040
         Width           =   4455
      End
      Begin VB.CheckBox chkAccess 
         Caption         =   "QC Manager (Confirm QC Process)"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   3015
      End
      Begin VB.CheckBox chkAccess 
         Caption         =   "Technical Manager ( Desin Arrester)"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   3015
      End
      Begin VB.CheckBox chkAccess 
         Caption         =   "Test Administrator (Perform Calibration)"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Value           =   1  'Checked
         Width           =   3015
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3480
      Top             =   5640
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin VB.Frame fram1 
      Caption         =   "Select Item:"
      Height          =   1335
      Left            =   360
      TabIndex        =   10
      Top             =   480
      Width           =   3735
      Begin VB.OptionButton optchng 
         Caption         =   "Cheng Personel Information"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   840
         Width           =   3015
      End
      Begin VB.OptionButton optnew 
         Caption         =   "Add New personel"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   480
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.Frame fram2 
      Caption         =   "Personal Information"
      Height          =   3495
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   3735
      Begin VB.CommandButton cmdtest 
         Caption         =   "Test data"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1320
         TabIndex        =   4
         Top             =   2640
         Width           =   1695
      End
      Begin VB.TextBox txtpass 
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Text            =   "1"
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox txtcode 
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Text            =   " "
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtlname 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Text            =   " "
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox txtfname 
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Text            =   " "
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Password"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Personal Code"
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Last Name"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "First Name"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   1080
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmnewper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private tested As Boolean
Private addnew As Boolean
Dim percode As String
Private Sub cmdapply_Click()
Dim allow As Boolean
Dim i As Integer
allow = True
percode = ""
If txtfname.Text = "" Or txtfname.Text = "  " Then
    allow = False
    MsgBox " Enter First Name Then Test Input Information.", vbOKOnly, "Error"
End If
If txtlname.Text = "" Or txtlname.Text = "  " And allow = True Then
    allow = False
    MsgBox " Enter Last Name Then Test Input Information.", vbOKOnly, "Error"
End If
If txtcode.Text = "" Or txtcode.Text = "  " And allow = True Then
    allow = False
    MsgBox " Enter Personal Code Then Test Input Information.", vbOKOnly, "Error"
End If
If txtpass.Text = "" Or txtpass.Text = "  " And allow = True Then
    allow = False
    MsgBox " Enter Password  Then Test Input Data.", vbOKOnly, "Error"
End If
If chkAccess(1).value = 0 And chkAccess(2).value = 0 And chkAccess(3).value = 0 And chkAccess(4).value = 0 And chkAccess(5).value = 0 And chkAccess(6).value = 0 Then
    allow = False
    MsgBox "You should select at least one access options", vbOKOnly, "Error"
End If

'***********************      apply chenge    **************
If allow = True Then
     
    percode = ""
    
    If chkAccess(1).value = 1 Then  'Program  Admin
        percode = "admin"
    Else                            ' Limited user
        percode = "user"
    End If
   
End If

If optnew.value = True And addnew = True And allow = True Then
    With Data1.Recordset
        .addnew
        !code = Val(txtcode.Text)
        !firstname = txtfname.Text
        !lastname = txtlname.Text
        !asccode = percode
        !pass = Val(txtpass.Text)
        !access = 1
        
        If chkAccess(1).value = 1 Then
            !Test_Admin = True
        Else
            !Test_Admin = False
        End If
        
        If chkAccess(2).value = 1 Then
            !Technical_Manager = True
        Else
            !Technical_Manager = False
        End If
        
        If chkAccess(3).value = 1 Then
            !QC_Manager = True
        Else
            !QC_Manager = False
        End If
        
        If chkAccess(4).value = 1 Then
            !QC_Oprt = True
        Else
            !QC_Oprt = False
        End If
        
        If chkAccess(5).value = 1 Then
            !Test_Oprt = True
        Else
            !Test_Oprt = False
        End If
        
       If chkAccess(6).value = 1 Then
            !Production_Worker = True
        Else
            !Production_Worker = False
        End If
        
        .Update
        
    End With
    
ElseIf optchng.value And tested = True And allow = True Then
    With Data1.Recordset
        .MoveFirst
        .Find "Code=" + txtcode.Text, , adSearchForward, 1
        .Update
        !code = Val(txtcode.Text)
        !firstname = txtfname.Text
        !lastname = txtlname.Text
        !asccode = percode
        !pass = Val(txtpass.Text)
        !access = 1
        
       If chkAccess(1).value = 1 Then
            !Test_Admin = True
        Else
            !Test_Admin = False
        End If
        
        If chkAccess(2).value = 1 Then
            !Technical_Manager = True
        Else
            !Technical_Manager = False
        End If
        
        If chkAccess(3).value = 1 Then
            !QC_Manager = True
        Else
            !QC_Manager = False
        End If
        
        If chkAccess(4).value = 1 Then
            !QC_Oprt = True
        Else
            !QC_Oprt = False
        End If
        
        If chkAccess(5).value = 1 Then
            !Test_Oprt = True
        Else
            !Test_Oprt = False
        End If
        
       If chkAccess(6).value = 1 Then
            !Production_Worker = True
        Else
            !Production_Worker = False
        End If
        
        
        .Update
                
    End With
   Data1.Refresh
End If
'//////////////////////////////       ready to   new work       //////////////////
If allow = True Then
    addnew = False
    tested = False
    cmdapply.Enabled = False
    If optchng.value = True Then
        txtcode.Enabled = True
        txtcode.Text = ""
        txtfname.Enabled = False
        txtlname.Enabled = False
        txtpass.Enabled = False
        txtfname.Text = ""
        txtlname.Text = ""
        txtpass.Text = ""
        For i = 1 To 6
            chkAccess(i).value = 0
        Next i
    Else
        txtcode.Enabled = True
        txtcode.Text = ""
        txtfname.Enabled = True
        txtlname.Enabled = True
        txtpass.Enabled = True
        txtfname.Text = ""
        txtlname.Text = ""
        txtpass.Text = ""
        
        For i = 1 To 6
            chkAccess(i).value = 0
        Next i
        
    End If
    
End If
End Sub

Private Sub cmdend_Click()
Data1.Recordset.Close
Unload Me
FrmInterface.Show
End Sub

Private Sub Cmdtest_Click()
Dim allow As Boolean
Dim A As Integer
allow = True
addnew = False
tested = False
'********************************************************************8
If optnew.value = True And allow = True Then
    If txtfname.Text = "" Or txtfname.Text = " " Then
        allow = False
        MsgBox " Enter First Name Then Test Input Information.", vbOKOnly, "Error'"
    End If
    If txtlname.Text = "" Or txtlname.Text = " " And allow = True Then
        allow = False
        MsgBox " Enter Last Name Then Test Input Information.", vbOKOnly, "Error'"
    End If
    If txtcode.Text = "" Or txtcode.Text = " " And allow = True Then
        allow = False
        MsgBox " Enter Personal Code Then Test Input Information.", vbOKOnly, "Error'"
    End If
    If txtpass.Text = "" Or txtpass.Text = " " And allow = True Then
        allow = False
        MsgBox " Enter Password  Then Test Input Information.", vbOKOnly, "Error'"
    End If
    If allow = True Then
        With Data1.Recordset
           .MoveFirst
            .Find "Code=" + txtcode.Text, , adSearchForward, 1
            If .EOF = False Then
                 MsgBox "Personal Code Is repetitious ." + Chr(13) + "Chenge Personel Code.", vbOKOnly, "Error"
                 txtcode.Text = ""
                 allow = False
                 cmdapply.Enabled = False
            Else
                 cmdapply.Enabled = True
                 addnew = True
            End If
        End With
    End If
'????????????????????????????????????????????????????????????????????????????
ElseIf optchng.value = True And allow = True Then
    If txtcode.Text = "" Or txtcode.Text = "  " And allow = True Then
        allow = False
        MsgBox " Enter Personal Code Then Test Input Information.", vbOKOnly, "Error'"
    Else
    With Data1.Recordset
       .MoveFirst
        .Find "Code=" + txtcode.Text, , adSearchForward, 1
        If .EOF = True Then
             tested = True
             MsgBox "Inserted ID Code was not Registered." + Chr(13) + "Enter Registered Personal Code.", vbOKOnly, "Error"
             txtcode.Text = ""
             cmdapply.Enabled = False
        Else
             txtfname.Text = .Fields("firstname")
             txtlname.Text = .Fields("lastname")
             txtpass.Text = .Fields("pass")
            
            If .Fields("asccode") = "admin" Then
                chkAccess(1).value = 1
            Else
                chkAccess(1).value = 0
            End If
            
             cmdapply.Enabled = True
             tested = True
        End If
    End With
    End If
End If
End Sub

Private Sub Form_Activate()
Dim constr As String
tested = False
addnew = False
cmdapply.Enabled = False
percodee = ""
On Error GoTo errorhand1

With Data1
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.acldb + ";Jet OLEDB:Database Password=" + mdl1.AclDBPas
    .CommandType = adCmdTable
    .RecordSource = "personel"
    .Refresh
    .Recordset.MoveFirst
End With


Exit Sub

errorhand1:
MsgBox "error in opening database.", vbOKOnly, "Error"
Cmdtest.Enabled = False
cmdapply.Enabled = False
End Sub

Private Sub Form_Load()
With Data1
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.acldb + ";Jet OLEDB:Database Password=" + mdl1.AclDBPas
    .CommandType = adCmdTable
    .RecordSource = "personel"
    .Refresh
    .Recordset.MoveFirst
End With
chkAccess(1).value = 1
frmnewper.icon = LoadPicture(mdl1.icon)
End Sub

Private Sub optchng_Click()
addnew = False
tested = False
cmdapply.Enabled = False
If optchng.value = True Then
    txtcode.Enabled = True
    txtcode.Text = ""
    txtfname.Enabled = False
    txtlname.Enabled = False
    txtpass.Enabled = False
    txtfname.Text = ""
    txtlname.Text = ""
    txtpass.Text = ""
Else
    txtcode.Enabled = True
    txtcode.Text = ""
    txtfname.Enabled = True
    txtlname.Enabled = True
    txtpass.Enabled = True
    txtfname.Text = ""
    txtlname.Text = ""
    txtpass.Text = ""
End If

End Sub

Private Sub optnew_Click()
addnew = False
tested = False
cmdapply.Enabled = False
If optnew.value = False Then
    txtcode.Enabled = True
    txtcode.Text = ""
    txtfname.Enabled = False
    txtlname.Enabled = False
    txtpass.Enabled = False
    txtfname.Text = ""
    txtlname.Text = ""
    txtpass.Text = ""
Else
    txtcode.Enabled = True
    txtcode.Text = ""
    txtfname.Enabled = True
    txtlname.Enabled = True
    txtpass.Enabled = True
    txtfname.Text = ""
    txtlname.Text = ""
    txtpass.Text = ""
End If
End Sub

Private Sub txtcode_Change()
addnew = False
tested = False
cmdapply.Enabled = False
End Sub

Private Sub txtcode_KeyPress(KeyAscii As Integer)
KeyAscii = mdl1.inputtest(KeyAscii)
End Sub

Private Sub txtfname_Change()
addnew = False
tested = False
cmdapply.Enabled = False
End Sub

Private Sub txtlname_Change()
addnew = False
tested = False
cmdapply.Enabled = False
End Sub

Private Sub txtpass_Change()
addnew = False
tested = False
cmdapply.Enabled = False
End Sub

Private Sub txtpass_KeyPress(KeyAscii As Integer)
KeyAscii = mdl1.inputtest(KeyAscii)
End Sub
