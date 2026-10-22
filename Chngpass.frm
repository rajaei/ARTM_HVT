VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmchpas 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CHANGE PASSWORD"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10815
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   10815
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3120
      Top             =   4440
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
   Begin VB.CommandButton cmdend 
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
      Height          =   375
      Left            =   7680
      TabIndex        =   8
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Frame Fram2 
      Caption         =   "Chenge Password"
      Height          =   3375
      Left            =   480
      TabIndex        =   12
      Top             =   600
      Width           =   5175
      Begin VB.TextBox txtcode 
         Height          =   375
         Left            =   1920
         TabIndex        =   0
         Text            =   " "
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdchn 
         Caption         =   "Chenge"
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
         Left            =   3840
         TabIndex        =   4
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox txtnew2 
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox txtnew1 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txtold 
         DataSource      =   "Data1"
         Height          =   375
         Left            =   1920
         TabIndex        =   1
         Text            =   " "
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblcod1 
         Caption         =   "Personel Code"
         Height          =   375
         Left            =   720
         TabIndex        =   18
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblnew2 
         Caption         =   "Confirm New Password"
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label lblnew1 
         Caption         =   "New Password"
         Height          =   255
         Left            =   720
         TabIndex        =   14
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lblold 
         Caption         =   "Old Password"
         Height          =   255
         Left            =   840
         TabIndex        =   13
         Top             =   1200
         Width           =   1095
      End
   End
   Begin VB.Frame fram1 
      Caption         =   "View Password"
      Height          =   3375
      Left            =   6120
      TabIndex        =   9
      Top             =   600
      Width           =   4335
      Begin VB.CommandButton cmdclear 
         Caption         =   "Clear"
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
         Left            =   2640
         TabIndex        =   7
         Top             =   2640
         Width           =   975
      End
      Begin VB.CommandButton txtview 
         Caption         =   "View"
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
         TabIndex        =   6
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox txtcode1 
         Height          =   405
         Left            =   1320
         TabIndex        =   5
         Text            =   " "
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lbllname 
         Caption         =   "Last Name : "
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   1800
         Width           =   3735
      End
      Begin VB.Label lblfname 
         Caption         =   "First Name :"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   1440
         Width           =   3855
      End
      Begin VB.Label lblpass 
         Caption         =   "Password :"
         Height          =   255
         Left            =   390
         TabIndex        =   11
         Top             =   1080
         Width           =   3495
      End
      Begin VB.Label lblcode 
         Caption         =   "Personel Code"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmchpas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private allow1, allow2, allow3 As Boolean
Dim dataopen As Boolean
Private Sub cmdchn_Click()
Dim rt As Integer
If txtcode.Text <> "" And txtcode.Text <> " " Then
    With Data1.Recordset
        .MoveFirst
        .Find ("code=") + Str(Val(txtcode.Text)), , adSearchForward, 1
        If .EOF = False Then
            dataopen = True
            allow1 = True
        Else
           MsgBox "Invavid Personel Code.", vbOKOnly + vbExclamation, "Error"
          txtcode.SetFocus
          allow1 = False
          dataopen = True
        End If
    End With
Else
    MsgBox "Insert Personel Code Then Password.", vbOKOnly + vbExclamation, "Error"
    txtcode.SetFocus
    allow1 = False
End If
'''''''''''''''''''''''''''''''''''
If allow1 = True And Data1.Recordset.Fields("pass") = Val(txtold.Text) Then
   allow2 = True
ElseIf allow1 = True And Data1.Recordset.Fields("pass") <> Val(txtold.Text) Then
    MsgBox "Invalid Password,Try Again.", vbOKOnly + vbExclamation, "Error"
    txtold.Text = ""
    allow2 = False
    txtold.SetFocus
End If
''''''''''''''''''''''''''''''
If allow1 = True And allow2 = True Then
    If Val(txtnew1.Text) <> Val(txtnew2.Text) Then
        MsgBox "Insert Same Password .", vbExclamation + vbOKOnly, "Error"
        txtnew1.Text = ""
        txtnew2.Text = ""
        txtnew1.SetFocus
        allow3 = False
    Else
        allow3 = True
    End If
End If
If allow1 = True And allow2 = True And allow3 = True Then
    With Data1.Recordset
        rt = MsgBox("First Name : " + .Fields("firstname") + Chr(13) + "Last Name : " + .Fields("lastname") + Chr(13) + "Do you want to chenge password?", vbYesNo + vbExclamation, "")
        If rt = 6 Then
            !pass = txtnew1.Text
            .Update
            txtcode.Text = ""
            txtold.Text = ""
            txtnew1.Text = ""
            txtnew2.Text = ""
            allow1 = False
            allow2 = False
            allow3 = False
        Else
            txtcode.Text = ""
            txtold.Text = ""
            txtnew1.Text = ""
            txtnew2.Text = ""
            allow1 = False
            allow2 = False
            allow3 = False
        End If
    End With
  End If

End Sub

Private Sub cmdclear_Click()
lblpass.Caption = "Password : "
lblfname.Caption = "First Name : "
lbllname.Caption = "Last Name : "
txtcode1.Text = ""
txtcode1.SetFocus
End Sub

Private Sub cmdend_Click()
If dataopen = True Then
    Data1.Recordset.Close
End If

Unload Me
FrmInterface.Show

End Sub

Private Sub Form_Activate()
allow1 = False
allow2 = False
allow3 = False
dataopen = False

With frmchpas.Data1
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.acldb + ";Jet OLEDB:Database Password=" + mdl1.AclDBPas
    .CommandType = adCmdTable
    .RecordSource = "personel"
    .Refresh
End With

If mdl1.access = "admin" Then
    fram1.Enabled = True
Else
    fram1.Enabled = False
End If
lblpass.Caption = "Password : "
lblfname.Caption = "First Name : "
lbllname.Caption = "Last Name : "
txtcode.SetFocus
End Sub

Private Sub Form_Load()
Dim access As Integer

With frmchpas.Data1
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.acldb + ";Jet OLEDB:Database Password=" + mdl1.AclDBPas
    .CommandType = adCmdTable
    .RecordSource = "personel"
    .Refresh
End With

frmchpas.icon = LoadPicture(mdl1.icon)
End Sub


Private Sub txtcode1_GotFocus()
lblpass.Caption = "Password : "
lblfname.Caption = "First Name : "
lbllname.Caption = "Last Name : "
txtcode1.Text = ""
End Sub

Private Sub txtview_Click()
If txtcode1.Text <> "" And txtcode1.Text <> " " And fram1.Enabled = True Then
    With Data1.Recordset
        .MoveFirst
        .Find ("code=") + Str(Val(txtcode1.Text)), , adSearchForward, 1
        If .EOF = False Then
            lblpass.Caption = "Password : " + Str(.Fields("pass"))
            lblfname.Caption = "First Name : " + .Fields("firstname")
            lbllname.Caption = "Last Name : " + .Fields("lastname")
        Else
           MsgBox "Insert Personel Code ", vbOKOnly + vbExclamation, "Error"
          lblpass.Caption = "Password : "
          lblfname.Caption = "First Name : "
          lbllname.Caption = "Last Name : "
          txtcode1.Text = ""
          txtcode1.SetFocus
        End If
    End With
Else
    MsgBox "Insert Personel Code ", vbOKOnly + vbExclamation, "Error"
    txtcode1.SetFocus
End If
End Sub
