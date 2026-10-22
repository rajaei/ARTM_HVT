VERSION 5.00
Begin VB.Form frmset 
   Caption         =   "SLIDER SETTING "
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CMDCANCEL 
      Caption         =   "CANCEL"
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton CMDOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtmax 
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
      Left            =   2400
      TabIndex        =   1
      Text            =   " "
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox txtmin 
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
      Left            =   2400
      TabIndex        =   0
      Text            =   " "
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "MAXIMUM VALUE  ="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   1200
      Width           =   1770
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "MINIMUME VALUE ="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   1785
   End
End
Attribute VB_Name = "frmset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDCANCEL_Click()
Me.Hide
End Sub

Private Sub CMDOK_Click()
If Val(txtmin.Text) < Val(txtmax.Text) And Val(txtmin.Text) > 0 Then
    FRMCAL.SLD.Min = 0
    FRMCAL.SLD.Max = Val(txtmax.Text) * 1000
    FRMCAL.SLD.Min = Val(txtmin.Text) * 1000
    FRMCAL.SLD.value = (FRMCAL.SLD.Min + FRMCAL.SLD.Max) / 2
    Me.Hide
Else
    MsgBox "Please insert currect value for minimume and maximum value ", vbExclamation, "ERROR"
    txtmin.Text = Str(FRMCAL.SLD.Min / 1000)
    txtmax.Text = Str(FRMCAL.SLD.Max / 1000)
End If
End Sub

Private Sub Form_Load()
txtmin.Text = Str(FRMCAL.SLD.Min / 1000)
txtmax.Text = Str(FRMCAL.SLD.Max / 1000)
End Sub
