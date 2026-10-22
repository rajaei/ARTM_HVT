VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLogin 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LOGIN"
   ClientHeight    =   4080
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   8010
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2410.598
   ScaleMode       =   0  'User
   ScaleWidth      =   7520.957
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Ado1 
      Height          =   495
      Left            =   1680
      Top             =   3240
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Operator Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   3480
      TabIndex        =   8
      Top             =   600
      Width           =   2655
      Begin VB.Label lblln 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " "
         DataField       =   "lastname"
         DataSource      =   "Datalogin"
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
         Left            =   2040
         TabIndex        =   12
         Top             =   960
         Width           =   60
      End
      Begin VB.Label lblfn 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " "
         DataField       =   "firstname"
         DataSource      =   "Datalogin"
         BeginProperty Font 
            Name            =   "B Lotus"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         TabIndex        =   11
         Top             =   480
         Width           =   60
      End
      Begin VB.Label lbltln 
         Caption         =   "Last Name :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lbltfn 
         Caption         =   "First Name :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.TextBox txtcode 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1440
      TabIndex        =   1
      Top             =   720
      Width           =   1725
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   1440
      TabIndex        =   4
      Top             =   2640
      Width           =   1260
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   4200
      TabIndex        =   5
      Top             =   2640
      Width           =   1260
   End
   Begin VB.TextBox txtpass 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1800
      Width           =   1725
   End
   Begin VB.Label l1 
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Insert Personal Code  And  Password."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label lblpass 
      Caption         =   " "
      DataField       =   "pass"
      DataSource      =   "Datalogin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblLabels 
      Caption         =   "Personal Code :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1200
   End
   Begin VB.Label lblLabels 
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub CMDCANCEL_Click()
   
    LoginSucceeded = False
    Unload Me
    End
End Sub

Private Sub CMDOK_Click()
If txtcode.Text <> "" And txtcode.Text <> "  " Then
    Ado1.Recordset.MoveFirst
    Ado1.Recordset.Find "Code=" + txtcode.Text, , adSearchForward, 1
    
    If Ado1.Recordset.EOF = True Then
        MsgBox "Invalid Personal Code, try again!", , "Login"
        txtcode.SetFocus
    Else
        lblfn.Visible = True
        lblln.Visible = True
        If txtpass.Text = Ado1.Recordset.Fields("pass") Then
            If Ado1.Recordset.Fields("access") = 0 Then
                MsgBox "You haven't been permitted to use of program", vbExclamation, "ERROR"
            Else
                LoginSucceeded = True
                Me.Hide
                mdl1.optcode = Val(txtcode.Text)
                mdl1.optname = lblfn.Caption + " " + lblln.Caption
                mdl1.access = Ado1.Recordset.Fields("asccode")
                
                With Ado1.Recordset
                    
                    If .Fields("Test_Admin") = True Then
                        mdl1.Test_Admin = True
                    Else
                        mdl1.Test_Admin = False
                    End If
                    
                    If .Fields("Technical_Manager") = True Then
                        mdl1.Technical_Manager = True
                    Else
                        mdl1.Technical_Manager = False
                    End If
                    
                    If .Fields("QC_Manager") = True Then
                        mdl1.QC_Manager = True
                    Else
                        mdl1.QC_Manager = False
                    End If
                    
                   If .Fields("QC_Oprt") = True Then
                        mdl1.QC_Operator = True
                    Else
                        mdl1.QC_Operator = False
                    End If
                    
                   If .Fields("Test_Oprt") = True Then
                        mdl1.HVTest_Operator = True
                    Else
                        mdl1.HVTest_Operator = False
                    End If
                    
                   If .Fields("Production_Worker") = True Then
                        mdl1.Production_Worker = True
                    Else
                        mdl1.Production_Worker = False
                    End If
                    
                    .Close
                    
                End With
                
                FrmInterface.Show
               ' frmarrtype.Show
            End If
        Else
            MsgBox "Invalid Password, try again!", , "Login"
            txtpass.SetFocus
            SendKeys "{Home}+{End}"
        End If
    End If
End If
End Sub

Private Sub Form_Activate()

lblfn.Visible = False
lblln.Visible = False
End Sub

Private Sub Form_Load()
frmLogin.icon = LoadPicture(mdl1.icon)

With frmLogin.Ado1
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.acldb + ";Jet OLEDB:Database Password=" + mdl1.AclDBPas
    .CommandType = adCmdTable
    .RecordSource = "personel"
    .Refresh
End With
lblfn.Top = lbltfn.Top * 0.8
lblln.Top = lbltln.Top * 0.85
lblln.Left = lblfn.Left

lblfn.Visible = False
lblln.Visible = False
End Sub

Private Sub txtpass_GotFocus()
With frmLogin.Ado1 ' .Datalogin
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.acldb + ";Jet OLEDB:Database Password=" + mdl1.AclDBPas
    .CommandType = adCmdTable
    .RecordSource = "personel"
    .Recordset.MoveFirst
End With
If txtcode.Text <> "" And txtcode.Text <> "  " Then
    Ado1.Recordset.MoveFirst
    Ado1.Recordset.Find "Code=" + txtcode.Text, , adSearchForward, 1
    If Ado1.Recordset.EOF = True Then
        MsgBox "Invalid Personal Code, try again!", , "Login"
        txtcode.SetFocus
         lblfn.Visible = False
         lblln.Visible = False
    Else
        lblfn.Visible = True
        lblfn.Caption = Ado1.Recordset.Fields("firstname")
        
        lblln.Visible = True
        lblln.Caption = Ado1.Recordset.Fields("lastname")
    End If
End If
End Sub
