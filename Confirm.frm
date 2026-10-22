VERSION 5.00
Begin VB.Form frmconfirm 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONFIRM NEW ARRESTER SERIAL NUMBER"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   6435
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdno 
      Caption         =   "No"
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
      Left            =   3240
      TabIndex        =   6
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdyes 
      Caption         =   "Yes"
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
      Left            =   960
      TabIndex        =   5
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Frame fram1 
      Caption         =   "Arrester Information "
      Height          =   2295
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   4935
      Begin VB.Label lbl3 
         Caption         =   "etyesrytsertser"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label lbl2 
         Caption         =   "tttryrtyrtyhrr"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label lbl1 
         Caption         =   " tesf[jkm,m"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Label lbl4 
      Caption         =   "tyrtyertyrty"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   2760
      Width           =   3495
   End
End
Attribute VB_Name = "frmconfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdno_Click()
Unload Me
frmmain.lbltool2.Caption = "New Serial Number:  From: " + Str(mdl1.newsrn(1)) + "  To :" + Str(mdl1.newsrn(mdl1.newarrnum))
End Sub

Private Sub cmdyes_Click()
Dim blc1 As String
Dim blc2 As String
Dim exist As Boolean
Dim testcod As Double
Dim allowpd As Boolean
allowpd = False

frmconfirm.MousePointer = 11

If frmmain.Data1.RecordSource <> "CNT" Then    ' Arresters

    blc1 = Str(Val(frmmain.cmbnblc1.Text)) + "@(" + Mid(frmmain.cmbsblc1, 2, 3) + Right(frmmain.cmbsblc1, 2) + ")"
    If frmmain.cmbsblc2.Visible = True Then
       blc2 = Str(Val(frmmain.cmbnblc2.Text)) + "@(" + Mid(frmmain.cmbsblc2, 2, 3) + Right(frmmain.cmbsblc2, 2) + ")"
    Else
       blc2 = "0@(00*00)"
    End If
    
    
    With frmmain.DataArrtype
        .RecordSource = "Testprmt"
        .Refresh
    End With
    
    exist = False
    
    With frmmain.DataArrtype.Recordset
        If .RecordCount >= 1 Then
            Do While .EOF = False
               If .Fields("accept_pd") = Val(frmmain.txtPDMax.Text) And .Fields("Ucon") = Val(frmmain.txtuc.Text) And .Fields("Urated") = Val(frmmain.txtUr1.Text) And .Fields("unit_number") = Val(frmmain.CmdNofloor.Text) And .Fields("Iref") = Val(frmmain.txtiref.Text) And .Fields("Icaccp") = Val(frmmain.txticaccp.Text) And .Fields("Urefmin") = Val(frmmain.txturefn.Text) And .Fields("Urefmax") = Val(frmmain.txturefx.Text) And .Fields("Uresmin") = Val(frmmain.txturesn.Text) And .Fields("Uresmax") = Val(frmmain.txturesx.Text) And .Fields("Block_Size1") = blc1 And .Fields("Block_Size2") = blc2 Then
                  exist = True
                  testcode = .Fields("Id")
                  Exit Do
               End If
               .MoveNext
            Loop
        End If
    End With
    
    If exist = False Then
        With frmmain.DataArrtype.Recordset
              .addnew
              !Urated = Val(frmmain.txtUr1.Text)
              !Ucon = Val(frmmain.txtuc.Text)
              !iref = Val(frmmain.txtiref.Text)
              !Icaccp = Val(frmmain.txticaccp.Text)
              !urefmin = Val(frmmain.txturefn.Text)
              !urefmax = Val(frmmain.txturefx.Text)
              !uresmin = Val(frmmain.txturesn.Text)
              !uresmax = Val(frmmain.txturesx.Text)
              !unit_number = Val(frmmain.CmdNofloor.Text)
              !accept_pd = Val(frmmain.txtPDMax.Text)
             
              !Block_Size1 = blc1
              !Block_Size2 = blc2
              .Update
              testcode = .Fields("Id")
        End With
    End If
End If

' add order information to production table


With frmmain.DataArrtype
    .RecordSource = "production"
    .Refresh
End With

With frmmain.DataArrtype.Recordset
    If .RecordCount > 0 Then
        .Find ("OrderNo=") + frmmain.txtorderNo.Text, , adSearchForward, 1
        If .EOF = True Then  ' new order code
            .MoveLast
            .addnew
            !OrderNo = Val(frmmain.txtorderNo.Text)
            !Quantity = Val(frmmain.TXTARRNO.Text)
            !NoOfSubDivision = Int((Val(frmmain.TXTARRNO.Text) / 40) + 0.999)
            !firstSN = mdl1.newsrn(1)
            !LastSN = mdl1.newsrn(mdl1.newarrnum)
            'blockinf = To_HejriChar(Date, 0)
            !Date = To_HejriChar(Date, 0)
            !Customer = frmmain.txtcustomer.Text
            !Type = mdl1.ArresterType
            
            If frmmain.Data1.RecordSource <> "CNT" Then
              !MOVType = frmmain.cmbblctype
              !Test_Inf = testcode
            End If
            
            .Update
            
        Else               ' repetitive order code
        
            !OrderNo = Val(frmmain.txtorderNo.Text)
            !Quantity = Val(frmmain.TXTARRNO.Text) + .Fields("Quantity")
            !NoOfSubDivision = Int((.Fields("Quantity") / 40) + 0.999)
            '!firstSN = mdl1.newsrn(1)
            !LastSN = mdl1.newsrn(mdl1.newarrnum)
            !Date = To_HejriChar(Date, 0)
            '!Customer = frmmain.txtcustomer.Text
            '!Type = mdl1.arrestertype
            '!MOVType = frmmain.cmbblctype
            '!Test_Inf = testcode
            .Update
            
            
        End If
    Else
        .addnew
        !OrderNo = Val(frmmain.txtorderNo.Text)
        !Quantity = Val(frmmain.TXTARRNO.Text)
        !NoOfSubDivision = Int((Val(frmmain.TXTARRNO.Text) / 40) + 0.999)
        !firstSN = mdl1.newsrn(1)
        !LastSN = mdl1.newsrn(mdl1.newarrnum)
        !Date = To_HejriChar(Date, 0)
        !Customer = frmmain.txtcustomer.Text
        !Type = mdl1.ArresterType
        
         If frmmain.Data1.RecordSource <> "CNT" Then
            !MOVType = frmmain.cmbblctype
            !Test_Inf = testcode
         End If
        .Update
    End If
End With

If frmmain.Data1.RecordSource <> "CNT" Then

    With frmmain.Data1.Recordset
        For i = 1 To mdl1.newarrnum
            .MoveLast
            .addnew
            !Serial_Num = mdl1.newsrn(i)
            !Arrester_Type = mdl1.ArresterType
            !Test_Inf = testcode
            !Order_num = Val(frmmain.txtorderNo)
            !Test1_Oprt = "Not Tested"
            !Register_Date = To_HejriChar(Date, 0)
            !Customer = frmmain.txtcustomer.Text
            !Production_Completed = False
            
            !CompleteTavanirCode = mdl1.CompleteTavanirCode  ' GLN 5 + GTIN 6 + Year 2 + Month 2 digit = 15 digit
            
            
            ' Create random number for each arrester to trace arrester manufactured by company
            Randomize
            !Tracing_Code = Abs(Int(Abs((100000 - 999999 + 1)) * Rnd + 100000))
            
            '!Test1_Date = Format(Now, "yyyy/mm/dd")
            'Result1 Code description: 1= Production is not completed, 2 = Production is completed & ready to HV test,  3 = Test1 is OK, 4 = Uref1 Is High,  5 = Uref1 is low,   6 = Leakage current1 is High,   7 = PD @ Uc1 is high,     8 = PD @ Ur1 is High
            'Result2 Code description: 1= Test2 not applied 2 = Test 1 is Ok & test2 not applied, 3 = Test2 is OK, 4 = Uref2 Is High,  5 = Uref2 is low, 6 = Ic2 is High,   7 = PD @ Uc2 is high,     8 = PD @ Ur2 is High
             
            If mdl1.ArrTableName = "DIST" Then
                 !result1 = 1 ' Arester is registered
                 !Result2 = 1
             Else
                 !result1 = 1
                 !Result2 = 1
             End If
                     
            !Block_Type = frmmain.cmbblctype
            .Update
        Next i
    End With
    
    
ElseIf frmmain.Data1.RecordSource = "CNT" Then  '  Surge Counter

   With frmmain.Data1.Recordset
        For i = 1 To mdl1.newarrnum
            .MoveLast
            .addnew
            !Serial_Num = mdl1.newsrn(i)
            !Surge_Counter_Type = mdl1.ArresterType
            !Order_num = Val(frmmain.txtorderNo)
            !Register_Date = To_HejriChar(Date, 0)
            !Customer = frmmain.txtcustomer.Text
            !Production_Completed = False
            !Meter_Manufacturer = frmmain.TXTMeterMan.Text
            !Meter_Model = frmmain.TXTMeterModel.Text
            !Counter_Manufacturer = frmmain.TXTCntMan.Text
            !Counter_Model = frmmain.TXTCntModel.Text
            !Protection_Type = frmmain.CmBProtection.Text
            
            !CompleteTavanirCode = mdl1.CompleteTavanirCode  ' GLN 5 + GTIN 6 + Year 2 + Month 2 digit = 15 digit
            
            
            ' Create random number for each arrester to trace counter manufactured by company
            Randomize
            !Tracing_Code = Abs(Int(Abs((100000 - 999999 + 1)) * Rnd + 100000))

            .Update
        Next i
    End With

End If


'                 add new arrster in vacuum database for HV arresters

If isHVArrester Then

    frmmain.datavac.Recordset.Open
    With frmmain.datavac.Recordset
      For i = 1 To mdl1.newarrnum
           '.MoveLast
           .addnew
           !serial_number = mdl1.newsrn(i)
           !Type = mdl1.ArresterType
           !unit = Val(frmmain.CmdNofloor.Text)
           !Test_Inf = testcode
           !Order_num = Val(frmmain.txtorderNo)
           !diaphragm = frmmain.cmbDiaph.Text
           !client = frmmain.txtcustomer.Text
           !Test_Oprt = "Not Tested"
           !vactest_date = To_HejriChar(Date, 0) ' Format(Now, "yyyy/mm/dd")
           !result = 0
           .Update
       Next i
    End With
    
    frmmain.datavac.Refresh
    frmmain.datavac.Recordset.Close
    
End If

frmmain.Data1.Refresh
frmconfirm.MousePointer = 0
frmmain.MousePointer = 0
Unload Me

End Sub

Private Sub Form_Load()
Dim he, wi As Double
With frmconfirm
    .Width = Screen.Width * 0.4
    .Height = Screen.Height * 0.4
    he = .Height
    wi = .Width
    .icon = LoadPicture(mdl1.icon)
End With
With FRAM1
    .Width = wi * 0.8
    .Height = he * 0.4
    .Left = wi * 0.1
    .Top = he * 0.078
End With
With lbl1
    .Width = FRAM1.Width * 0.8
    .Height = FRAM1.Height * 0.2
    .Left = FRAM1.Width * 0.1
    .Top = FRAM1.Height * 0.25
End With
With lbl2
    .Width = FRAM1.Width * 0.8
    .Height = FRAM1.Height * 0.2
    .Left = FRAM1.Width * 0.1
    .Top = FRAM1.Height * 0.48
End With
With lbl3
    .Width = FRAM1.Width * 0.8
    .Height = FRAM1.Height * 0.2
    .Left = FRAM1.Width * 0.1
    .Top = FRAM1.Height * 0.7
End With

With lbl4
    .Width = wi * 0.8
    .Height = he * 0.4
    .Left = wi * 0.12
    .Top = he * 0.53
    .Caption = "Are you sure to add the new serial numbers to database?"
End With

With cmdyes
    .Width = wi * 0.2
    .Height = he * 0.1
    .Left = wi * 0.2
    .Top = he * 0.76
End With
With cmdno
    .Width = wi * 0.2
    .Height = he * 0.1
    .Left = wi * 0.6
    .Top = he * 0.76
End With

End Sub

