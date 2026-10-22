VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmrpt 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REPORT & SEARCH "
   ClientHeight    =   8805
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   15195
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   15195
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.ComboBox cmbPrtList 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   7200
      Width           =   3375
   End
   Begin VB.CommandButton CmdprintGaranti 
      Caption         =   "Print Guarantee Form"
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
      Left            =   5040
      TabIndex        =   21
      Top             =   7080
      Width           =   2655
   End
   Begin MSComDlg.CommonDialog cdlg1 
      Left            =   1920
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog prt1 
      Left            =   1080
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame fram2 
      Caption         =   "Select Report Type"
      Height          =   1455
      Left            =   9480
      TabIndex        =   16
      Top             =   240
      Width           =   4815
      Begin VB.OptionButton opt 
         Caption         =   "Test Parameter"
         Height          =   255
         Index           =   3
         Left            =   2400
         TabIndex        =   9
         Top             =   960
         Width           =   2055
      End
      Begin VB.OptionButton opt 
         Caption         =   "Product Informatin"
         Height          =   195
         Index           =   2
         Left            =   2280
         TabIndex        =   8
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton opt 
         Caption         =   "MOV Block Instruction"
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton opt 
         Caption         =   "Test report"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.Frame fram1 
      Caption         =   "«» œ« ‰Ê⁄ „Õ’Ê· Ê »⁄œ «“ ¬‰ ‘„«—Â »—‰«„Â —« À»  ‰„«ÌÌœ."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   13
      Top             =   120
      Width           =   8895
      Begin VB.ComboBox CmbSubOrder 
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
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox TxtOrderNo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         TabIndex        =   1
         Text            =   " "
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmbtype 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   960
         Width           =   1335
      End
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
         Left            =   7080
         TabIndex        =   5
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmdapply 
         Caption         =   "Add To List"
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
         Left            =   7320
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtsn2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5280
         TabIndex        =   3
         Text            =   " "
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtsn1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5280
         TabIndex        =   2
         Text            =   " "
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label LblSubOrder 
         Caption         =   "Sub Order:"
         Height          =   375
         Left            =   2040
         TabIndex        =   19
         Top             =   840
         Width           =   975
      End
      Begin VB.Label LblOrderNo 
         Caption         =   "Order No."
         Height          =   375
         Left            =   2040
         TabIndex        =   18
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblArType 
         Caption         =   "‰Ê⁄ „Õ’Ê·"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   17
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblsn2 
         Caption         =   "TO :"
         Height          =   375
         Left            =   4800
         TabIndex        =   15
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblsn1 
         Caption         =   "FROM :"
         Height          =   255
         Left            =   4440
         TabIndex        =   14
         Top             =   480
         Width           =   735
      End
   End
   Begin MSFlexGridLib.MSFlexGrid fgrpt 
      Height          =   4935
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   8705
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      BackColor       =   -2147483637
      GridColor       =   8421504
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdclose 
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
      Left            =   10440
      TabIndex        =   11
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "Print Test Report"
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
      Left            =   8160
      TabIndex        =   10
      Top             =   7080
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc datatestprmt 
      Height          =   330
      Left            =   11760
      Top             =   7080
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
   Begin MSAdodcLib.Adodc datamain 
      Height          =   330
      Left            =   11520
      Top             =   7560
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
      Caption         =   "datamain"
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
      Height          =   330
      Left            =   11760
      Top             =   8160
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
Attribute VB_Name = "frmrpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private colnum As Integer
Private seisrn() As Double
Private selarrnum As Double
Private applysn, mp3 As Boolean
Private PrintBarcode As Boolean

Private arrtype As String  ' Type Of Selected Arrester
Private srn1, srn2, sn, TotalArrNum, LastArrNumber As Double
Private srnrpt() As Double
Private TotalRegsSerNo() As Double
Private address As String
Private aryprt() As Double
Private aryprmt() As Double
Private arytype() As String
Private aryGrntAllowed() As Boolean
Private aryGrntDate() As String

Private AryTavanirCode() As Double
Private aryTraceCode() As Double


Private selarrtype As String

Dim Tabletype() As String
Dim NoOfSubprog As Double
Dim FirstSNofOrder, LastSNofOrder As Double


Private Function GetInstalledPrinter() As Boolean
    
    Dim ssPrinters As String * 4096
    Dim sPrinterList() As String
    Dim iIndex As Long
    Dim sTemp As String
    
    
    cmbPrtList.Clear
    If GetBIXOLON_PrinterList(ssPrinters) <= 0 Then
        GetInstalledPrinter = False
    Else
        sTemp = CStr(ssPrinters)
        sPrinterList() = Split(sTemp, "^")
        
        If UBound(sPrinterList) >= 0 Then
                        
            For iIndex = 0 To UBound(sPrinterList)
                cmbPrtList.AddItem sPrinterList(iIndex)
            Next
            
            cmbPrtList.ListIndex = 0
            GetInstalledPrinter = True
            PrintBarcode = True
        Else
            GetInstalledPrinter = False
            PrintBarcode = False
        End If
        
    End If
    
End Function

Private Sub InitControls()

    Dim i As Long
    Dim bResult As Boolean
    Dim sDllVersion As String * 256
    Dim dValue As Double
        
    'If GetDllVersion(sDllVersion) Then
        'lblDllVersion.Caption = "Version " + CStr(sDllVersion)
    'Else
        'lblDllVersion.Caption = "Unknown"
    'End If
    
    cmbPrtList.Enabled = GetInstalledPrinter()
        
    'cmbDensity.Clear
    'For i = 0 To 20 - 1
        'cmbDensity.AddItem CStr(i + 1)
    'Next
    'cmbDensity.ListIndex = (14 - 1)
    
    'cmbSpeed.Clear
    'cmbSpeed.AddItem CStr("Printer Setting")
    'cmbSpeed.ListIndex = 0
    
    '--------------------------------------------------------
    ' Paper Size
    ' Default 4x6 inch -> (4*25.4) x (6*25.4) mm
   ' dValue = 75
    'txtP_Width.Text = CStr(dValue)
    'dValue = 40
    'txtP_Height.Text = CStr(dValue)
    
    'txtMargin_X.Text = "0"
    'txtMargin_Y.Text = "0"

End Sub

Private Sub printsub()
Dim x1, x2, x3, x4, x5, x6, x7, x8, x9, x10, x11, x12, x13, x14, x15, x16, x17, x18 As Long
Dim y1, y2, y3, y4, y5, y6, y7, y8 As Long
Dim xx1, xx2, xx3 As Long
Dim colno, rowno, arrno As Double
Dim printend, newarr As Boolean
Dim prtno As Long

    x1 = 0
    x2 = 42.64
    x3 = 155.9
    x4 = 210
    x5 = 50.102
    x6 = 103.402
    x7 = 155.636
    x8 = 41.574
    x9 = 57.564
    x10 = 73.556
    x11 = 89.544
    x12 = 105.534
    x13 = x12 + x8
    x14 = x12 + x9
    x15 = x12 + x10
    x16 = x12 + x11
    x17 = 65.026
    x18 = 116.194
    x19 = 164.164
    '//////////////////////''''''''''''''''''''''''
    y1 = 0 + 5
    y2 = 21 + 5
    y3 = 29.179 + 5
    y4 = 59 + 5
    y5 = 72.9477 + 5
    y6 = 255.317 + 5
    y7 = 285.58 + 5
    y8 = 295 + 5
    
    selarrnum = TotalArrNum
    
    '"""""""""""""""""""
    arrno = 0
    printend = False
   'fr prtno = 1 To prt1.Copies
    
        Printer.FontSize = 12                   ' chech printer error
        Printer.FontName = "Times New Roman"
        Printer.Scale (0, 0)-(211.2, 301.2)  ' A4 size = 210 * 297 [cm]
        Do Until printend = True
        newarr = False
          For colno = 1 To 2
             For rowno = 1 To 24
                arrno = arrno + 1
                ''""""""""""""print data"""""""""""
                Select Case colno
                Case 1
                    Printer.FontName = "Times New Roman"
                    Printer.FontSize = 15
                    Printer.CurrentX = x1 + 2
                    Printer.CurrentY = y5 + ((rowno - 1) * (y6 - y5) / 24) + 0.7
                    Printer.Print Str(aryprt(1, arrno))
                 '""""""""""""""""""""""""""""""""""""""""
                    Printer.FontSize = 12
                    Printer.CurrentX = x8 + 1
                    Printer.CurrentY = y5 + ((rowno - 1) * (y6 - y5) / 24) + 0.8
                    Printer.Print Format(aryprt(2, arrno), "##0.00")
                 '"""""""""""""""""""""""""""""""""""""""""
                    Printer.CurrentX = x9 + 1
                    Printer.CurrentY = y5 + ((rowno - 1) * (y6 - y5) / 24) + 0.8
                    Printer.Print Format(aryprt(3, arrno), "#0.000")
                 '""""""""""""""""""""""""""""""""""""""""""
                    Printer.CurrentX = x10 + 1
                    Printer.CurrentY = y5 + ((rowno - 1) * (y6 - y5) / 24) + 0.8
                    Printer.Print Format(aryprt(4, arrno), "#0.00")
                 '"""""""""""""""""""""""""""""""""""""""""
                    Printer.CurrentX = x11 + 1
                    Printer.CurrentY = y5 + ((rowno - 1) * (y6 - y5) / 24) + 0.8
                    Printer.Print Format(aryprt(5, arrno), "#0.00")
                
                Case 2
                    Printer.FontName = "Times New Roman"
                    Printer.FontSize = 15
                    Printer.CurrentX = x12 + 2
                    Printer.CurrentY = y5 + ((rowno - 1) * (y6 - y5) / 24) + 0.7
                    Printer.Print Str(aryprt(1, arrno))
                 '""""""""""""""""""""""""""""""""""""""""
                    Printer.FontSize = 12
                    Printer.CurrentX = x13 + 1
                    Printer.CurrentY = y5 + ((rowno - 1) * (y6 - y5) / 24) + 0.8
                    Printer.Print Format(aryprt(2, arrno), "##0.00")
                 '"""""""""""""""""""""""""""""""""""""""""
                    Printer.CurrentX = x14 + 1
                    Printer.CurrentY = y5 + ((rowno - 1) * (y6 - y5) / 24) + 0.8
                    Printer.Print Format(aryprt(3, arrno), "#0.000")
                 '""""""""""""""""""""""""""""""""""""""""""
                    Printer.CurrentX = x15 + 1
                    Printer.CurrentY = y5 + ((rowno - 1) * (y6 - y5) / 24) + 0.8
                    Printer.Print Format(aryprt(4, arrno), "#0.00")
                 '"""""""""""""""""""""""""""""""""""""""""
                    Printer.CurrentX = x16 + 1
                    Printer.CurrentY = y5 + ((rowno - 1) * (y6 - y5) / 24) + 0.8
                    Printer.Print Format(aryprt(5, arrno), "#0.00")
                
                End Select
                '""""""""""""""""""""""""" print test parameter """""""
                
                If arrno < selarrnum Then
                    If aryprmt(9, arrno) <> aryprmt(9, arrno + 1) Then
                        newarr = True
                    End If
               End If
                If arrno = selarrnum Then printend = True
                If (printend = True Or newarr = True) Then Exit For
                
             Next rowno
             If (printend = True Or newarr = True) Then Exit For
          Next colno
        '//////////////////form defult/////////
        
        '//////////////////////////////////////////////
        Printer.CurrentX = 0
        Printer.CurrentY = 0
        Printer.PaintPicture LoadPicture(mdl1.pic2), 1, 6, 40, 20
        ';;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
        Printer.FontSize = 12
        Printer.FontName = "Times New Roman"
        Printer.CurrentX = x2 + (x3 - x2) * 0.05
        Printer.CurrentY = y1 + (y2 - y1) * 0.5 - Printer.TextHeight("ELEC") * 0.5
        Printer.FontBold = True
        Printer.Print ("ELECTRICAL ROUTINE TEST RESULTS FORM ")
        Printer.FontBold = False
        '.......................................
        Printer.FontSize = 10
        Printer.CurrentX = x1 + (x2 - x1) * 0.2
        Printer.CurrentY = y2 + ((y3 - y2) - Printer.TextHeight("ELEC")) * 0.5
        Printer.Print "Temperature : .......... C"
        '***************************************************************************
        Printer.FontSize = 10
        Printer.FontBold = False
        Printer.CurrentX = x2 * 1.7
        Printer.CurrentY = y2 + ((y3 - y2) - Printer.TextHeight("œ„«")) * 0.5
        Printer.Print "Test Date: ..................."
        '*************************************************
        Printer.FontSize = 10
        Printer.CurrentX = x3
        Printer.CurrentY = y2 + ((y3 - y2) - Printer.TextHeight("œ„«")) * 0.5
        Printer.Print "Test Sheet No.  .........."
        
        ',,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
        Printer.FontName = "Times New Roman"
        Printer.FontSize = 16
        Printer.FontBold = True
        Printer.CurrentX = x3 + (x4 - x3) * 0.33
        Printer.CurrentY = y1 + ((y2 - y1) * 0.25) - Printer.TextHeight("ELEC") * 0.5
        Printer.Print "ETP Co."
        '...............................................
        Printer.FontSize = 16
        Printer.FontBold = True
        Printer.CurrentX = x3 + (x4 - x3) * 0.05
        Printer.CurrentY = y1 + (y2 - y1) * 0.74 - Printer.TextHeight("ELEC") * 0.5
        Printer.Print "PAYA ARRESTER"
        
       
        '''''''''''''''''''''''''/////////////////
        Printer.FontSize = 12
        Printer.CurrentX = x7 + (x4 - x7) * 0.1
        Printer.CurrentY = y3 + (y4 - y3) * 0.33
        Printer.Print "Residual Voltage Limits"
        ''''''''''''''''''''''
        Printer.CurrentX = x7 + (x4 - x7) * 0.8
        Printer.CurrentY = y3
       
        '/////////////////////////////
        
        
        Printer.FontName = "Times New Roman"
        Printer.FontSize = 10
        Printer.CurrentX = x1 + (x2 - x1) * 0.1
        Printer.CurrentY = y6 * 1.01
        Printer.Print "Test Operator :"
        
        '''''''''''''''''''''
        Printer.CurrentX = x17 + (x18 - x17) * 0.08
        Printer.CurrentY = y6 * 1.01
        Printer.Print "Production Manager :"
           
        '///////////////////
        Printer.CurrentX = x18 + (x19 - x18) * 0.08
        Printer.CurrentY = y6 * 1.01
        Printer.Print "QC Manager : "
        '''''''''''''''''''''''''''''''
        
        '///////////////////////
        Printer.CurrentX = x19 + (x4 - x19) * 0.08
        Printer.CurrentY = y6 * 1.01
        Printer.Print "Notes : "
        '''''''''''''''''''''''''''''''
        
        Printer.FontSize = 12
        Printer.FontName = "Times New Roman"
        ''''''''''''''''''''''''''''''''''''''''
        Printer.CurrentX = x1 + (x8 - x1) * 0.3
        Printer.CurrentY = y4 + (y5 - y4) * 0.3
        Printer.Print "Serial No."
        ''''''''''''''''''''''''''''''''
        Printer.CurrentX = x8 + (x9 - x8) * 0.25
        Printer.CurrentY = y4 + (y5 - y4) * 0.055
        Printer.Print "Uref"
        '''''''''''''''''''''''''
        Printer.CurrentX = x8 + (x9 - x8) * 0.2
        Printer.CurrentY = y4 + (y5 - y4) * 0.55
        Printer.Print "(KV)"
        '''''''''''''''''''''''''''''
        Printer.CurrentX = x9 + (x10 - x9) * 0.4
        Printer.CurrentY = y4 + (y5 - y4) * 0.05
        Printer.Print "Ic"
        '''''''''''''''''''''''''
        Printer.CurrentX = x9 + (x10 - x9) * 0.22
        Printer.CurrentY = y4 + (y5 - y4) * 0.55
        Printer.Print "(mA)"
        ''''''''''''''''''''''''
        Printer.CurrentX = x10 + (x12 - x10) * 0.25
        Printer.CurrentY = y4 + (y5 - y4) * 0.05
        Printer.Print "PD (Pc)"
        '''''''''''''''''''''''''
        Printer.CurrentX = x10 + (x11 - x10) * 0.005
        Printer.CurrentY = y4 + (y5 - y4) * 0.55
        Printer.Print "1.05 Uc"
        ''''''''''''''''''''''
        Printer.CurrentX = x11 + (x12 - x11) * 0.3
        Printer.CurrentY = y4 + (y5 - y4) * 0.55
        Printer.Print "Ur"
        '/////////////////////////
        Printer.CurrentX = x12 + (x13 - x12) * 0.3
        Printer.CurrentY = y4 + (y5 - y4) * 0.3
        Printer.Print "Serial No."
        ''''''''''''''''''''''''''''''''
        Printer.CurrentX = x13 + (x14 - x13) * 0.25
        Printer.CurrentY = y4 + (y5 - y4) * 0.055
        Printer.Print "Uref"
        '''''''''''''''''''''''''
        Printer.CurrentX = x13 + (x14 - x13) * 0.2
        Printer.CurrentY = y4 + (y5 - y4) * 0.55
        Printer.Print "(KV)"
        '''''''''''''''''''''''''''''
        Printer.CurrentX = x14 + (x15 - x14) * 0.4
        Printer.CurrentY = y4 + (y5 - y4) * 0.05
        Printer.Print "Ic"
        '''''''''''''''''''''''''
        Printer.CurrentX = x14 + (x15 - x14) * 0.22
        Printer.CurrentY = y4 + (y5 - y4) * 0.55
        Printer.Print "(mA)"
        ''''''''''''''''''''''''
        Printer.CurrentX = x15 + (x4 - x15) * 0.25
        Printer.CurrentY = y4 + (y5 - y4) * 0.05
        Printer.Print "PD (Pc)"
        '''''''''''''''''''''''''
        Printer.CurrentX = x15 + (x16 - x15) * 0.005
        Printer.CurrentY = y4 + (y5 - y4) * 0.55
        Printer.Print "1.05 Uc"
        ''''''''''''''''''''''
        Printer.CurrentX = x16 + (x4 - x16) * 0.3
        Printer.CurrentY = y4 + (y5 - y4) * 0.55
        Printer.Print "Ur"
        '''//////////////////////////////////////////
        Printer.FontSize = 16
        Printer.CurrentX = x5 * 0.95
        Printer.CurrentY = y3 * 1.005
        Printer.Print "ACCEPTANCE CIRITERIA "
        '''''''''''''''''''''''''''''''''''
        Printer.FontSize = 12
        Printer.CurrentX = x7 * 1.02
        Printer.CurrentY = y3 * 1.01
        Printer.Print "Type : " + arytype(arrno)
        '''''''''''''''''''''''''''''''''''
        Printer.CurrentX = x1 + (x5 - x1) * 0.1
        Printer.CurrentY = y3 * 1.22
        Printer.Print "Reference Voltage"
        '''''''''''''''''''''''''''''
        Printer.CurrentX = x6 + (x7 - x6) * 0.2
        Printer.CurrentY = y3 * 1.22
        Printer.Print "Partial Discharge"
        '''''''''''''''''''''''''''''''''''
        Printer.CurrentX = x5 + (x6 - x5) * 0.2
        Printer.CurrentY = y3 * 1.22
        Printer.Print " Leakage Current"
        
        ''''''''''''''''''''''''''''
        Printer.CurrentX = x1 + (x5 - x1) * 0.1
        Printer.CurrentY = y3 * 1.45
        Printer.Print "   Iref = " + Format(aryprmt(1, arrno), "#0.0") + " (mA)"
        ''''''''''''''''''''''''''''''''''
        Printer.CurrentX = x5 + (x6 - x5) * 0.1
        Printer.CurrentY = y3 * 1.45
        Printer.Print "    Uc = " + Format(aryprmt(4, arrno), "##0.0") + " (KV)"
        ''''''''''''''''''''''''''''''''''
        Printer.CurrentX = x6 + (x7 - x6) * 0.03
        Printer.CurrentY = y3 * 1.45
        Printer.Print "     1.05 Uc = " + Format(aryprmt(4, arrno) * 1.05, "##0.0") + " (KV)"
        ''''''''''''''''''''''''''''''''''
        
        Printer.CurrentX = x1 + (x5 - x1) * 0.05
        Printer.CurrentY = y3 * 1.65
        Printer.Print Format(aryprmt(2, arrno), "##0.0") + " < Uref < " + Format(aryprmt(3, arrno), "##0.0") + " (KV)"
        ''''''''''''''''''''''''''''''''''
        Printer.CurrentX = x7 + (x4 - x7) * 0.1
        Printer.CurrentY = y3 * 1.55
        Printer.Print Format(aryprmt(6, arrno), "##0.0") + " < Ures < " + Format(aryprmt(7, arrno), "##0.0") + " (KV)"
        ''''''''''''''''''''''''''''''''''''
        Printer.CurrentX = x5 + (x6 - x5) * 0.2
        Printer.CurrentY = y3 * 1.65
        Printer.Print " Ic <  " + Format(aryprmt(5, arrno), "0.0#") + " (mA)"
        ''''''''''''''''''''''''''''''''''
        Printer.CurrentX = x6 + (x7 - x6) * 0.3
        Printer.CurrentY = y3 * 1.65
        Printer.Print "PD < 2 (Pc)"
        ''''''''''''''''''''''''''''''''''
        '................................draw line and grid.......
        '''''''''''''''''''''''''''''''''fram 1//////////////////
        Printer.DrawWidth = 17
        Printer.Line (x1, y1)-(x4, y1)
        Printer.DrawWidth = 11
        Printer.Line -(x4, y2)
        Printer.Line -(x1, y2)
        Printer.DrawWidth = 20
        Printer.Line (x1, y2)-(x1, y1)
        Printer.DrawWidth = 11
        Printer.Line (x2, y1)-(x2, y2)
        Printer.Line (x3, y1)-(x3, y2)
        '////////////////////////////fram 2///////////////
        Printer.Line (x1, y3)-(x4, y3)
        Printer.Line -(x4, y4)
        Printer.Line -(x1, y4)
        Printer.DrawWidth = 20
        Printer.Line -(x1, y3)
        Printer.DrawWidth = 11
        '//////////////////fram 3,4,5//////////////
        Printer.Line (x1, y4)-(x4, y4)
        Printer.Line -(x4, y7)
        Printer.Line -(x1, y7)
        Printer.DrawWidth = 20
        Printer.Line (x1, y7)-(x1, y4)
        Printer.DrawWidth = 11
        Printer.Line (x12, y4)-(x12, y6)
        '//////////////////////////////fram 3,5//////////
        Printer.Line (x1, y5)-(x4, y5)
        Printer.Line (x1, y6)-(x4, y6)
        '///////////////////////////////fram 6 ////////////
        Printer.Line (x1, y7)-(x1, y8)
        Printer.Line (x4, y7)-(x4, y8)
        Printer.Line (x1, y8)-(x4, y8)
        '///////////////////////dividing grid in fram 4//////////
        Printer.DrawWidth = 5
        For i = 1 To 23
           Printer.Line (x1, y5 + ((i / 24) * (y6 - y5)))-(x4, y5 + ((i / 24) * (y6 - y5)))
        Next i
        '////////////////////////////grid in fram 2/////////
        Printer.Line (x1, y3 + 0.25 * (y4 - y3))-(x4, y3 + 0.25 * (y4 - y3))
        Printer.Line (x1, y3 + 0.45 * (y4 - y3))-(x7, y3 + 0.45 * (y4 - y3))
        Printer.Line (x5, y3 + 0.25 * (y4 - y3))-(x5, y4)
        Printer.Line (x6, y3 + 0.25 * (y4 - y3))-(x6, y4)
        Printer.Line (x7, y3)-(x7, y4)
        '////////////////////grid in fram 3///////////
        Printer.Line (x10, y4 + (y5 - y4) * 0.5)-(x12, y4 + (y5 - y4) * 0.5)
        Printer.Line (x15, y4 + (y5 - y4) * 0.5)-(x4, y4 + (y5 - y4) * 0.5)
        '////////////////grid in  fram 4''''''''''''
        Printer.Line (x8, y4)-(x8, y6)
        Printer.Line (x9, y4)-(x9, y6)
        Printer.Line (x10, y4)-(x10, y6)
        Printer.Line (x11, y4 + ((y5 - y4) * 0.5))-(x11, y6)
        Printer.Line (x13, y4)-(x13, y6)
        Printer.Line (x14, y4)-(x14, y6)
        Printer.Line (x15, y4)-(x15, y6)
        Printer.Line (x16, y4 + ((y5 - y4) * 0.5))-(x16, y6)
        
        '///////////////grid in fram 5'////////////
        Printer.Line (x17, y6)-(x17, y7)
        Printer.Line (x18, y6)-(x18, y7)
        Printer.Line (x19, y6)-(x19, y7)
        
        '////////////////////////////grid in fram 6'///////////////////
        xx1 = x1 + ((x4 - x1) * 0.25)
        xx2 = x1 + ((x4 - x1) * 0.5)
        xx3 = x1 + ((x4 - x1) * 0.75)
        
        Printer.Line (xx1, y8)-(xx1, y7)
        Printer.Line (xx2, y8)-(xx2, y7)
        Printer.Line (xx3, y8)-(xx3, y7)
       '//////////////////////////////////////////////
        
        Printer.FontSize = 10
        Printer.CurrentX = x1 + (xx1 - x1) * 0.1
        Printer.CurrentY = y7 + 1
        Printer.Print " D:ETP\F\F-15-05-01.doc"
        '''''''''''''''''''''''''''''''''''''''''''''
    
        Printer.CurrentX = xx1 + (xx2 - xx1) * 0.1
        Printer.CurrentY = y7 + 1
        Printer.Print "Doc. Code: F-05-01"
        ''''''''''''''''''''''''''''''''''''''''''''''
        Printer.CurrentX = xx2 + (xx3 - xx2) * 0.08
        Printer.CurrentY = y7 + 1
        Printer.Print "Rev. 01"
        '''''''''''''''''''''''''''''''''''''''''''''''
        Printer.CurrentX = xx3 + (x4 - xx3) * 0.1
        Printer.CurrentY = y7 + 1
        Printer.FontName = "Times New Roman"
        Printer.Print "Doc. Distribution: QC"
        
        Printer.NewPage
       
        '///////////////////////////////////////
        
        Loop
    Printer.EndDoc
End Sub
Private Sub PrintGuaranteeSub()

Dim arrno As Double
Dim prtno As Long
Dim cntyear, cntmonth, cntsr As String
Dim NewSerial As Double

'////////////////////////////////////////////////////////////
Dim bResult As Boolean

Dim strPrinterName As String
    
Dim dotsPer1mm As Long
Dim multiplier As Long
Dim resolution As Long
Dim nPaper_Width As Long
Dim nPaper_Height As Long
Dim nMarginX As Long
Dim nMarginY As Long

Dim nSpeed As Long
Dim nDensity As Long

Dim bAutoCut As Boolean
Dim bReverseFeeding As Boolean
Dim nSensorType As Long
Dim xString As Long
Dim yString As Long
    
Dim Direct_MaxiCode_2DBarcode_Cmd As String
Dim Direct_PDF_2DBarcocde_Cmd As String
Dim Direct_DataMatrix_2DBarcocde_Cmd As String
Dim QRCode_data As String
Dim Maxicode_data As String
Dim DataMatrix_data As String
Dim PDF417_data As String
    
Dim ArresterType, SerialNumber, SerialNumber1, SerialNumber2, SerialNumber3, TracingCode As String




'///////////////////////////////////////////

selarrnum = TotalArrNum
    '"""""""""""""""""""
arrno = 0
printend = False

strPrinterName = cmbPrtList.Text

If ConnectPrinter(strPrinterName) = False Then
   MsgBox "Can't connect to Barcode printer of Bixolon" + Chr(13), vbOKOnly + vbExclamation, "ERROR"
  Exit Sub
End If

' 203 DPI : 1mm = 8 dots
' 300 DPI : 1mm = 12 dots
' 600 DPI : 1mm = 24 dots
resolution = GetPrinterResolution()

If resolution <= 203 Then
    dotsPer1mm = 8
    multiplier = 1
ElseIf resolution <= 300 Then
    dotsPer1mm = 12
    multiplier = 1
ElseIf resolution <= 600 Then
    dotsPer1mm = 24
    multiplier = 3
End If


nPaper_Width = CInt(75 * dotsPer1mm)  ' Width of label set on 75 mm
nPaper_Height = CInt(40 * dotsPer1mm)  ' Height of label set on 40 mm
nMarginX = CInt(1 * dotsPer1mm)   ' X Margin of label set on 1 mm
nMarginY = CInt(1 * dotsPer1mm)   ' Y Margin of label set on 1 mm
    
nSpeed = SLCS_PRINT_SPEED.PRINTER_SETTING_SPEED
    
nDensity = 14 ' CInt(cmbDensity.Text)

nSensorType = GAP

        
'   Set the label start
bResult = StartLabel()

'   Clear Buffer of Printer
bResult = ClearBuffer()
    
    

' Select international character set and code table.To
bResult = SetCharacterset(SLCS_INTERNATIONAL_CHARSET.ICS_USA, SLCS_CODEPAGE.FCP_CP1252)
    
'Set Label and Printer
'SetConfigOfPrinter(SPEED_50, 17, TOP, false, 0, true);
bResult = SetConfigOfPrinter(nSpeed, nDensity, Top, False, 0, True)

'1 Inch : 25.4mm
'1 mm   :  8 Dots (MX4000, TX400, DX420, DX220, DL410, T400, D420, D220, SRP-770/770II/770III)
'1 mm   : 12 Dots (MX4030, TX403, DX423, DX223, DL413, T403, D423, D223)
'1 mm   : 24 Dots (MX4060)

bResult = SetPaper(nMarginX, nMarginY, nPaper_Width, nPaper_Height, nSensorType, 0, 2)  ' 75 mm (Width) * 40 mm (Hiehgt)
  
bResult = PrintDirect("STd", True)

'If rdoDt.Value Then ' Direct thermal
    'bResult = PrintDirect("STd", True)
'Else ' Thermal transfer
    'bResult = PrintDirect("STt", True)
'End If
    

    ' Prints string using TrueFont
    '  P1 : Horizontal position (X) [dot]
    '  P2 : Vertical position (Y) [dot]
    '  P3 : Font Name
    '  P4 : Font Size
    '  P5 : Rotation : (0 : 0 degree , 1 : 90 degree, 2 : 180 degree, 3 : 270 degree)
    '  P6 : Italic
    '  P7 : Bold
    '  P8 : Underline
    '  P9 : RLE (Run-length encoding)


'////////////////////////////////////////////////////////////////////////////////////////////////////

' œ—Ì«›  „‘Œ’«  „«Â  Ê·Ìœ Ê«ﬁ⁄Ì Õò ‘œÂ —ÊÌ »—ﬁêÌ—
    
    cntmonth = Mid(aryGrntDate(1), 6, 2)
    cntmonth = InputBox("·ÿ›« „«Â  Ê·Ìœ ﬁÌœ ‘œÂ —ÊÌ  —„Ì‰«· —« Ê«—œ ‰„«ÌÌœ", "„«Â  Ê·Ìœ", cntmonth, Screen.Width / 3, Screen.Height / 3)
    
    If cntmonth = "" Then
        cntmonth = Mid(aryGrntDate(1), 6, 2)
    End If
    
    
'  Print  Barcode


'For prtno = 1 To prt1.Copies  ' nomber of copies
    
    For i = 1 To TotalArrNum  ' number of arresters
    
        If aryGrntAllowed(i) = True Then
        
          
          cntsr = Right(Str(aryprt(1, i)), 6)                       ' Serial Number
          cntyear = Format((aryprt(1, i) - cntsr) / 10 ^ 6, "0#")   ' Year
          cntmonth = Mid(aryGrntDate(i), 6, 2)                      ' month
                     
          SerialNumber = Right(AryTavanirCode(i), 15) + cntsr '  GLN Code + GTIN Code + Year + Month + Serial No.
          
          SerialNumber1 = Left(SerialNumber, 5) '  GLN part
          SerialNumber2 = Mid(SerialNumber, 6, 6) ' GTIN Part
          SerialNumber3 = Mid(SerialNumber, 12, 4) ' Year + Month Part
          
        
          '   ç«Å ‰«„ ›«”Ì ‘—ò 
          bResult = PrintTrueFont(45 * dotsPer1mm, 0.5 * dotsPer1mm, "B zar", 14, 0, True, True, False, "*»—ﬁêÌ— Å«Ì«*", False)
         
         '   ç«Å  „œ· »—ﬁêÌ—
          bResult = PrintTrueFont(3 * dotsPer1mm, 2 * dotsPer1mm, "Arial", 10, 0, False, True, False, "Type: " + arytype(i), False)
           
           '   ç«Å »«— òœ
          bResult = Print1DBarcode(4 * dotsPer1mm, 8 * dotsPer1mm, 1, 3 * multiplier, 5 * multiplier, 100 * multiplier, ROTATE_0, 1, SerialNumber)
          
          ' ç«Å «’«·  ò«·«
          bResult = PrintTrueFont(1 * dotsPer1mm, 24 * dotsPer1mm, "Arial", 8, 0, False, True, False, "ÃÂ  «ÿ„Ì‰«‰ «“ «’«·  ò«·«° òœ —ÂêÌ—Ì Ê ”—Ì«· —« »Â ‘—ò  «⁄·«„ ‰„«ÌÌœ", False)
          
          '  ç«Å ‘„«—Â ”—Ì«· ⁄œœÌ Ê  ›òÌòÌ
          bResult = PrintTrueFont(2 * dotsPer1mm, 29 * dotsPer1mm, "Arial", 10, 0, False, True, False, "Serial No.: " + SerialNumber1 + " - " + SerialNumber2 + " - " + SerialNumber3 + " - " + cntsr, False)
          
          ' ç«Å òœ —ÂêÌ—Ì
          bResult = PrintTrueFont(2 * dotsPer1mm, 34 * dotsPer1mm, "Arial", 9, 0, False, True, False, "Tracing Code:" + Str(aryTraceCode(i)), False)
          
          ' ç«Å  «—ÌŒ  Ê·Ìœ
          bResult = PrintTrueFont(35 * dotsPer1mm, 34 * dotsPer1mm, "Arial", 9, 0, False, True, False, "- Date of Prod.:" + aryGrntDate(i), False)
          
            
         '   Print Command
          bResult = Prints(1, 2)
          '   Set the Label End
          'EndLabel
          
          Printer.NewPage
        End If
     
    Next i

'Next prtno

'   Set the Label End
      EndLabel
                
'   Disconnect Printer Driver
    bResult = DisconnectPrinter()

End Sub





Private Sub CmbSubOrder_Change()

If CmbSubOrder.ListIndex < CmbSubOrder.ListCount - 1 Then
    txtsn1.Text = Str(FirstSNofOrder) + (40 * CmbSubOrder.ListIndex)
    If Val(txtsn1.Text) + 39 < LastSNofOrder Then
        txtsn2.Text = Str(Val(txtsn1.Text) + 39)
    Else
        txtsn2.Text = Str(LastSNofOrder)
        
    End If
Else
    txtsn1.Text = Str(FirstSNofOrder)
    txtsn2.Text = Str(LastSNofOrder)
End If

End Sub

Private Sub CmbSubOrder_Click()

If CmbSubOrder.ListIndex < CmbSubOrder.ListCount - 1 Then
    txtsn1.Text = Str(FirstSNofOrder) + (40 * CmbSubOrder.ListIndex)
    If Val(txtsn1.Text) + 39 < LastSNofOrder Then
        txtsn2.Text = Str(Val(txtsn1.Text) + 39)
    Else
        txtsn2.Text = Str(LastSNofOrder)
        
    End If
Else
    txtsn1.Text = Str(FirstSNofOrder)
    txtsn2.Text = Str(LastSNofOrder)
End If

End Sub


Private Sub cmbtype_Click()

With datamain

    If Tabletype(cmbtype.ListIndex) = "SUBST" Then  '  Station Class Arrester Coding
        .RecordSource = "SUBST"
              
    ElseIf Tabletype(cmbtype.ListIndex) = "DIST" Then '  dISTRIBUTION Class Arrester Coding
        .RecordSource = "DIST"
    ElseIf Tabletype(cmbtype.ListIndex) = "CNT" Then '  Counter Coding
        .RecordSource = "CNT"
                
            
    End If
    
    .Refresh
    
End With

TotalArrNum = 0
ReDim aryprt(1 To 6, 1 To 1)
ReDim aryprmt(1 To 9, 1 To 1)
ReDim arytype(1 To 1)
ReDim aryGrntDate(1 To 1)
ReDim AryTavanirCode(1 To 1)

ReDim aryTraceCode(1 To 1)

ReDim aryGrntAllowed(1 To 1)
ReDim TotalRegsSerNo(1 To 1)

  
    
applysn = False
cmdprint.Enabled = False
CmdprintGaranti.Enabled = False

'If PrintBarcode = True Then CmdprintGaranti.Enabled = True

With fgrpt
    .RowHeightMin = 100
    .Rows = 1
    .Cols = 12
    colnum = 12
    .ColWidth(0) = .Width * 0.08
    .ColWidth(1) = .Width * 0.08
    .ColWidth(2) = .Width * 0.08
    .ColWidth(3) = .Width * 0.08
    .ColWidth(4) = .Width * 0.08
    .ColWidth(5) = .Width * 0.08
    .ColWidth(6) = .Width * 0.15
    .ColWidth(7) = .Width * 0.08
    .ColWidth(8) = .Width * 0.08
    .ColWidth(9) = .Width * 0.08
    .ColWidth(10) = .Width * 0.08
    .ColWidth(11) = .Width * 0.15
                            
'.AllowUserResizing = flexResizeNone
    .Row = 0
    .Col = 0
    .Text = "Serial.No"
    .CellAlignment = 0
                            
    .Col = 1
    .Text = "Arrester Type"
    .Col = 2
    .Text = "Uref @ Iref (T.1)"
    .Col = 3
    .Text = "Ic @ Uc (T.1)"
    .Col = 4
    .Text = "PD @ 1.05*Uc (T.1)"
    .Col = 5
    .Text = "PD @ Ur (T.1)"
    .Col = 6
    .Text = "Result of Test 1"
    .Col = 7
    .Text = "Uref @ Iref (T.2)"
    .Col = 8
    .Text = "Ic @ Uc (T.2)"
    .Col = 9
    .Text = "PD @ 1.05*Uc (Test 2)"
    .Col = 10
    .Text = "PD @ Ur (T.2)"
    .Col = 11
    .Text = "Result of Test 2"
                            
                            
    For i = 0 To 11
        .Col = i
        .CellBackColor = &H8000000A
    Next i
End With

opt(0).Value = True
txtsn1.Text = ""
txtsn2.Text = ""
txtorderNo.Text = ""

opt(0).Value = True
fram2.Enabled = False
CmbSubOrder.Clear
CmbSubOrder.Enabled = False


End Sub


Private Sub cmdapply_Click()
    Dim allow, ended As Boolean
    Dim txtsrn1, txtsrn2, temstr As String
    Dim state, Sal, A As Integer
    Dim arrcnt, fc, cnti As Long
    Dim ProductionCompleted As Boolean
    
    On Error GoTo errhdnl
    
    frmrpt.MousePointer = 11
    fgrpt.Visible = False
    allow = True
    applysn = False
    cmdprint.Enabled = False
    CmdprintGaranti.Enabled = False
        
    srn1 = Val(txtsn1.Text)
    srn2 = Val(txtsn2.Text)
        
   '///////////////////////////////////type of report ///////////
   
        opt(0).Value = True
        fram2.Enabled = False
        If opt(0).Value = True Then
          state = 0
        ElseIf opt(1).Value = True Then
          state = 1
        ElseIf opt(2).Value = True Then
          state = 2
        ElseIf opt(3).Value = True Then
          state = 3
        End If
    
        '/////////////////////////////////input test///////
        
        selarrnum = 1   ' number of selected arrester
        
        If txtsn1.Text = "" Or txtsn1.Text = " " Then
           MsgBox "Enter Serial Number" + Chr(13), vbOKOnly + vbExclamation, "ERROR"
           allow = False
           selarrnum = 0
           txtsn1.SetFocus
        End If
        
        If allow = True And (txtsn2.Text = "" Or txtsn2.Text = " ") Then
            selarrnum = 1   ' number of selected arrester
            srn1 = Val(txtsn1.Text)
            srn2 = srn1
        End If
        If allow = True And Val(txtsn1.Text) > Val(txtsn2.Text) And (txtsn2.Text <> "" And txtsn2.Text <> " ") Then
            MsgBox " — Ì» ‘„«—Â ”—Ì«·Â« ’ÕÌÕ ‰„Ì »«‘œ.", vbOKOnly, "ERROR"
            txtsn2.Text = ""
            txtsn2.SetFocus
            allow = False
        End If
                         
         If allow = True Then
             selarrnum = srn2 - srn1 + 1 ' number of selected arrester
         End If
         
         LastArrNumber = TotalArrNum
         
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$4
        If allow = True Then
                   
            With datatestprmt
                  .RecordSource = "Testprmt"
                  .Refresh
            End With
            
            selarrnum = 0
            For sn = srn1 To srn2
                 
                datamain.Recordset.Find "Serial_Num=" + Str(sn), , adSearchForward, 1
                  
                If datamain.Recordset.EOF = True Then
                      
                    MsgBox "Serial number  " + Str(sn) + " has not been registered ", vbOKOnly + vbExclamation, "ERROR"
                     MsgBox " ‰Ê⁄ „Õ’Ê· Ê ‘„«—Â »—‰«„Â —« œ—”  «‰ Œ«» ‰„«ÌÌœ." + Chr(13) + "«ÿ·«⁄«  ‘„«—Â »—‰«„Â ›Êﬁ ÿ»ﬁ ‰Ê⁄ „Õ’Ê· «‰ Œ«»Ì –ŒÌ—Â ‰‘œÂ «” .", vbOKOnly, "ERROR"
                    allow = False
                    txtsn1.Text = ""
                    txtsn2.Text = ""
                    txtsn1.SetFocus
                    opt(0).Value = True
                    fram2.Enabled = False
                    Exit For
                           
                Else
                      'selarrnum = selarrnum + 1
                    TotalArrNum = TotalArrNum + 1
                    ReDim Preserve srnrpt(1 To TotalArrNum)
                    srnrpt(TotalArrNum) = datamain.Recordset.Fields("Serial_Num")
                End If
             Next sn
        End If
        
           
        '/////////////////////////////////////insert data in grid ////////////////////
        
        If allow = True Then
                applysn = True   '  apply sn for report
                
                fram2.Enabled = True
                 
                '   *****  define arry for printing information***********
                
                'LastArrNumber = TotalArrNum
                'TotalArrNum = TotalArrNum + selarrnum
                
                ReDim Preserve TotalRegsSerNo(1 To TotalArrNum)
                ReDim Preserve arytype(1 To TotalArrNum)
                ReDim Preserve aryprt(1 To 6, 1 To TotalArrNum)
                ReDim Preserve aryprmt(1 To 9, 1 To TotalArrNum)
                ReDim Preserve aryGrntDate(1 To TotalArrNum)
                
                ReDim Preserve AryTavanirCode(1 To TotalArrNum)
                
                ReDim Preserve aryTraceCode(1 To TotalArrNum)
                ReDim Preserve aryGrntAllowed(1 To TotalArrNum)
                
                
                Select Case state
                Case 0 '///////////Test Result////////
                
                  cmdprint.Enabled = True
                  If cmbPrtList.Enabled Then CmdprintGaranti.Enabled = True
                    
                  fgrpt.RIGHTTOLEFT = False
                  fgrpt.Font = Arial
                  
                  
                  If LastArrNumber = 0 Then
                        With fgrpt
                            .RowHeightMin = 100
                            .Rows = 1
                            .Cols = 12
                            colnum = 12
                            .ColWidth(0) = .Width * 0.07
                            .ColWidth(1) = .Width * 0.08
                            .ColWidth(2) = .Width * 0.07
                            .ColWidth(3) = .Width * 0.07
                            .ColWidth(4) = .Width * 0.08
                            .ColWidth(5) = .Width * 0.06
                            .ColWidth(6) = .Width * 0.15
                            .ColWidth(7) = .Width * 0.07
                            .ColWidth(8) = .Width * 0.07
                            .ColWidth(9) = .Width * 0.08
                            .ColWidth(10) = .Width * 0.06
                            .ColWidth(11) = .Width * 0.17
                            
                            
                            
                            '.AllowUserResizing = flexResizeNone
                            .Row = 0
                            .Col = 0
                            .Text = "Serial.No"
                            .CellAlignment = 0
                            
                            .Col = 1
                            .Text = "Arrester Type"
                             .Col = 2
                            .Text = "Uref @ Iref (T.1)"
                            .Col = 3
                            .Text = "Ic @ Uc (T.1)"
                            .Col = 4
                            .Text = "PD @ 1.05*Uc (T.1)"
                            .Col = 5
                            .Text = "PD @ Ur (T.1)"
                            .Col = 6
                            .Text = "Result of Test 1"
                            .Col = 7
                            .Text = "Uref @ Iref (T.2)"
                            .Col = 8
                            .Text = "Ic @ Uc (T.2)"
                            .Col = 9
                            .Text = "PD @ 1.05*Uc (Test 2)"
                            .Col = 10
                            .Text = "PD @ Ur (T.2)"
                            .Col = 11
                            .Text = "Result of Test 2"
                            
                            
                            For i = 0 To 11
                                .Col = i
                                .CellBackColor = &H8000000A
                            Next i
                        End With
                  End If
                  
                  
                    
                    For cnti = LastArrNumber + 1 To TotalArrNum
                    
                        sn = srnrpt(cnti)
                        datamain.Recordset.Find "Serial_Num=" + Str(sn), , adSearchForward, 1
                        
                        With fgrpt
                              .Rows = .Rows + 1
                              .Row = .Rows - 1
                              .Col = 0
                              .Text = datamain.Recordset.Fields("Serial_Num")
                              .CellBackColor = &H8000000E
                              .CellAlignment = 0
                              
                              TotalRegsSerNo(cnti) = Val(.Text)
                              
                              aryprt(1, cnti) = datamain.Recordset.Fields("Serial_Num")
                              
                              arytype(cnti) = datamain.Recordset.Fields("Arrester_Type")
                              
                              .Col = 1
                              .Text = datamain.Recordset.Fields("Arrester_Type")
                              .CellBackColor = &H8000000E
                                         
                              .Col = 2
                              .Text = Str(datamain.Recordset.Fields("Urefiref1"))
                              .CellBackColor = &H8000000E
                              
                              .Col = 3
                              .Text = Str(datamain.Recordset.Fields("IcUc1"))
                              .CellBackColor = &H8000000E
                              
                              .Col = 4
                              .Text = Str(datamain.Recordset.Fields("pduc1"))
                              .CellBackColor = &H8000000E
                              
                              .Col = 5
                              .Text = Str(datamain.Recordset.Fields("pdur1"))
                              .CellBackColor = &H8000000E
                              
                               '  ‰ŸÌ„  «—ÌŒ ‘—Ê⁄ ê«—«‰ Ì «“ “„«‰ »” Â »‰œÌ
                              aryGrntDate(cnti) = datamain.Recordset.Fields("Production_phase7_Date")
                               '  ‰ŸÌ„ « „«„  Ê·Ìœ
                              aryGrntAllowed(cnti) = datamain.Recordset.Fields("Production_Completed")
                              'òœ «‰Õ’«—Ì Â— »—ﬁêÌ—
                              aryTraceCode(cnti) = datamain.Recordset.Fields("Tracing_Code")
                              
                              ' òœ  Ê«‰Ì— „Õ’Ê·
                              AryTavanirCode(cnti) = datamain.Recordset.Fields("CompleteTavanirCode")
                              
                              If datamain.Recordset.Fields("Result2") < 3 Then  'Test2 not applied & test1 is reference
                                
                                aryprt(2, cnti) = Str(datamain.Recordset.Fields("Urefiref1"))
                                aryprt(3, cnti) = Str(datamain.Recordset.Fields("IcUc1"))
                                aryprt(4, cnti) = Str(datamain.Recordset.Fields("pduc1"))
                                aryprt(5, cnti) = Str(datamain.Recordset.Fields("pdur1"))
                                
                                
                                
                              Else    ' test 2 is reference
                              
                                aryprt(2, cnti) = Str(datamain.Recordset.Fields("Urefiref2"))
                                aryprt(3, cnti) = Str(datamain.Recordset.Fields("IcUc2"))
                                aryprt(4, cnti) = Str(datamain.Recordset.Fields("pduc2"))
                                aryprt(5, cnti) = Str(datamain.Recordset.Fields("pdur2"))
                              
                              End If
                              
                              .Col = 6
                              .CellBackColor = &H8000000E
                              
                              Select Case Str(datamain.Recordset.Fields("Result1"))
                              Case 0
                                .Text = "Not Applied"
                              Case 1
                                  .Text = "Production Process isn't Completed"
                              Case 2
                                  .Text = "Ready for HV Test"
                              Case 3
                                  .Text = "Test 1 is OK"
                              Case 4
                                  .Text = "Uref 1 is High"
                              Case 5
                                  .Text = "Uref 1 is Low"
                              Case 6
                                  .Text = "Iref 1 is Failed"
                              Case 7
                                  .Text = "PD@Uc 1 is failed"
                              Case 8
                                  .Text = "PD@Ur 1 is failed"
                              End Select
                              
                              .Col = 7
                              .Text = Str(datamain.Recordset.Fields("Urefiref2"))
                              
                              .CellBackColor = &H80000003
                              
                              .Col = 8
                              .Text = Str(datamain.Recordset.Fields("IcUc2"))
                              
                              .CellBackColor = &H80000003
                              
                              .Col = 9
                              .Text = Str(datamain.Recordset.Fields("pduc2"))
                             
                              .CellBackColor = &H80000003
                              
                              .Col = 10
                              .Text = Str(datamain.Recordset.Fields("pdur2"))
                              
                              .CellBackColor = &H80000003
                              
                              .Col = 11
                              .CellBackColor = &H80000003
                              Select Case Str(datamain.Recordset.Fields("Result2"))
                              Case 0
                                .Text = "Not Applied"
                              Case 1
                                  .Text = "Not Applied"
                              Case 2
                                  .Text = "Test1 is OK ,Test2 don't Required"
                              Case 3
                                  .Text = "Test 2 is OK"
                              Case 4
                                  .Text = "Uref 2 is High"
                              Case 5
                                  .Text = "Uref 2 is Low"
                              Case 6
                                  .Text = "Iref 2 is Failed"
                              Case 7
                                  .Text = "PD@Uc 2 is failed"
                              Case 8
                                  .Text = "PD@Ur 2 is failed"
                              End Select
                                                     
                              
                          End With
                              
                      
                          '"""""""""""""""'test parameter   ''''''
                             
                              datatestprmt.RecordSource = "Testprmt"
                              datatestprmt.Refresh
                              datatestprmt.Recordset.Find "ID=" + Str(datamain.Recordset.Fields("Test_Inf")), , adSearchForward, 1
                              
                              If datatestprmt.Recordset.EOF = False Then
                                  aryprmt(1, cnti) = Val(Format(datatestprmt.Recordset.Fields("Iref"), "##.##"))
                                  aryprmt(2, cnti) = datatestprmt.Recordset.Fields("Urefmin")
                                  aryprmt(3, cnti) = datatestprmt.Recordset.Fields("Urefmax")
                                  aryprmt(4, cnti) = datatestprmt.Recordset.Fields("Ucon")
                                  aryprmt(5, cnti) = datatestprmt.Recordset.Fields("Icaccp")
                                  aryprmt(6, cnti) = datatestprmt.Recordset.Fields("Uresmin")
                                  aryprmt(7, cnti) = datatestprmt.Recordset.Fields("Uresmax")
                                  aryprmt(8, cnti) = datatestprmt.Recordset.Fields("Urated")
                                  aryprmt(9, cnti) = datatestprmt.Recordset.Fields("ID")
                              End If
                              
                    Next cnti
                 
                 ',,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
                Case 1              '///////////BLock Information////////////////
                
                    fgrpt.RIGHTTOLEFT = False
                    CmdprintGaranti.Enabled = False
                    cmdprint.Enabled = False
                    
                    fgrpt.Font = Arial
                    
                    With datatestprmt
                        .RecordSource = "Testprmt"
                        .Refresh
                     End With
                     If LastArrNumber = 0 Then
                         With fgrpt
                             
                             .Rows = 1
                             .RowHeightMin = 100
                             .Cols = 6
                             colnum = 6
                             
                             .ColWidth(0) = .Width * 0.12
                             .ColWidth(1) = .Width * 0.12
                             .ColWidth(2) = .Width * 0.12
                             .ColWidth(3) = .Width * 0.12
                             .ColWidth(4) = .Width * 0.12
                             .ColWidth(5) = .Width * 0.12
                             '.AllowUserResizing = flexResizeNone
                             .Row = 0
                             .Col = 0
                             .Text = "Serial Number of Arrester"
                             .CellAlignment = 2
                             .Col = 1
                             .Text = "Manufacturer of MOV"
                             .CellAlignment = 2
                              .Col = 2
                             .Text = "Quantity of Type1"
                             .CellAlignment = 2
                             .Col = 3
                             .Text = "Dimension of Type1 (D*H)"
                             .CellAlignment = 2
                             .Col = 4
                             .Text = "Quantity of Type2"
                             .CellAlignment = 2
                             .Col = 5
                             .Text = "Dimension of Type2 (D*H)"
                             .CellAlignment = 2
                            
                            For i = 0 To 5
                                .Col = i
                                .CellBackColor = &H8000000A
                            Next i
                            
                        End With
                    End If
                   
                   For cnti = LastArrNumber + 1 To TotalArrNum
                        sn = srnrpt(cnti)
                        datamain.Recordset.Find "Serial_Num=" + Str(sn), , adSearchForward, 1
                        With fgrpt
                            .Rows = .Rows + 1
                            .Row = .Rows - 1
                            .Col = 0
                            .Text = datamain.Recordset.Fields("Serial_Num")
                            .CellAlignment = 2
                            
                            TotalRegsSerNo(cnti) = Val(.Text)
                            datatestprmt.Recordset.Find "ID=" + Str(datamain.Recordset.Fields("Test_Inf")), , adSearchForward, 1
                            If datatestprmt.Recordset.EOF = True Then
                                .Col = 1
                                .Text = "Not Found "
                                .Col = 3
                                .Text = "Test "
                                .Col = 2
                                .Text = "Information"
                                .Col = 5
                                .Text = ""
                                .Col = 4
                                .Text = ""
                                
                            Else
                                .Col = 1
                                .Text = datamain.Recordset.Fields("Block_Type")  ' manufacturer
                                .CellAlignment = 2
                                .Col = 2
                                .Text = Val(datatestprmt.Recordset.Fields("Block_Size1"))
                                .CellAlignment = 2
                                .Col = 3
                                .Text = Right(datatestprmt.Recordset.Fields("Block_Size1"), 7)
                                .CellAlignment = 2
                                .Col = 4
                                .Text = Val(datatestprmt.Recordset.Fields("Block_Size2"))
                                .CellAlignment = 2
                                .Col = 5
                                .Text = Right(datatestprmt.Recordset.Fields("Block_Size2"), 7)
                                .CellAlignment = 2
                               ' .CellBackColor = &H80000003
                            End If
                        End With
                        
                   Next cnti
                   
                ';;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
                
                Case 2             '/////////////Product Information///////////
                
                    fgrpt.RIGHTTOLEFT = True
                    CmdprintGaranti.Enabled = False
                    cmdprint.Enabled = False
                    
                    If LastArrNumber = 0 Then
                          With fgrpt
                            
                            .RowHeightMin = 100
                            .Rows = 1
                            .Cols = 22
                            colnum = 22
                            fgrpt.Font = "B Koodak"
                            For i = 0 To 1
                                .ColWidth(i) = .Width * 0.06
                            Next i
                            
                            .ColWidth(2) = .Width * 0.15
                            
                            For i = 3 To 21
                                .ColWidth(i) = .Width * 0.08
                            Next i
                            
                            For i = 0 To 21
                                .Col = i
                                .CellBackColor = &H8000000A
                            Next i
                            
                            .AllowUserResizing = flexResizeNone
                            
                            ' Add Title of production process name
                    
                            datatestprmt.RecordSource = "Type"
                            datatestprmt.Refresh
                        
                            datatestprmt.Recordset.MoveFirst
                            datatestprmt.Recordset.Find ("Arr_Type='") + cmbtype.Text + "'", , adSearchForward, 1
                            
                            
                            .Row = 0
                            
                            .Col = 0
                            .Text = "‘„«—Â ”—Ì«·"
                            .Col = 1
                            .Text = "‘„«—Â »—‰«„Â"
                             .Col = 2
                            .Text = "‰«„ „‘ —Ì"
                            .Col = 3
                            .Text = "«Å—« Ê— " + datatestprmt.Recordset.Fields("Production_Phase0")
                            .Col = 4
                            .Text = " «—ÌŒ " + datatestprmt.Recordset.Fields("Production_Phase0")
                            .Col = 5
                            .Text = "«Å—« Ê— " + datatestprmt.Recordset.Fields("Production_Phase1")
                            .Col = 6
                            .Text = " «—ÌŒ " + datatestprmt.Recordset.Fields("Production_Phase1")
                            .Col = 7
                            .Text = "«Å—« Ê— " + datatestprmt.Recordset.Fields("Production_Phase2")
                             .Col = 8
                            .Text = " «—ÌŒ " + datatestprmt.Recordset.Fields("Production_Phase2")
                            .Col = 9
                            .Text = "«Å—« Ê— " + datatestprmt.Recordset.Fields("Production_Phase3")
                            .Col = 10
                            .Text = " «—ÌŒ " + datatestprmt.Recordset.Fields("Production_Phase3")
                            .Col = 11
                            .Text = "«Å—« Ê— " + datatestprmt.Recordset.Fields("Production_Phase4")
                            .Col = 12
                            .Text = " «—ÌŒ " + datatestprmt.Recordset.Fields("Production_Phase4")
                            .Col = 13
                            .Text = "«Å—« Ê— " + datatestprmt.Recordset.Fields("Production_Phase5")
                             .Col = 14
                            .Text = " «—ÌŒ " + datatestprmt.Recordset.Fields("Production_Phase5")
                            .Col = 15
                            .Text = "«Å—« Ê—  ”  «·ò —ÌòÌ"
                            .Col = 16
                            .Text = " «—ÌŒ  ”  «·ò —ÌòÌ"
                             .Col = 17
                            .Text = "«Å—« Ê— " + datatestprmt.Recordset.Fields("Production_Phase7")
                            .Col = 18
                            .Text = " «—ÌŒ " + datatestprmt.Recordset.Fields("Production_Phase7")
                            .Col = 19
                            .Text = "«Å—« Ê— " + datatestprmt.Recordset.Fields("Production_Phase6")
                            .Col = 20
                            .Text = " «—ÌŒ " + datatestprmt.Recordset.Fields("Production_Phase6")
                            .Col = 21
                            .Text = "Ê÷⁄Ì   Ê·Ìœ"
                            
                        End With
                  End If
                  
                 
                  
                  For cnti = LastArrNumber + 1 To TotalArrNum
                        sn = srnrpt(cnti)
                        datamain.Recordset.Find "Serial_Num=" + Str(sn), , adSearchForward, 1
                        
                        
                        With fgrpt
                            .Rows = .Rows + 1
                            .Row = .Rows - 1
                            
                            If datamain.Recordset.Fields("Production_Completed") = True Then
                                .Text = " ò„Ì· ›—«Ì‰œ  Ê·Ìœ"
                                For i = 0 To 21
                                    .Col = i
                                    .CellBackColor = vbGreen
                                Next i
                            Else
                                .Text = "œ— Õ«·  Ê·Ìœ"
                                For i = 0 To 21
                                    .Col = i
                                    .CellBackColor = vbYellow
                                Next i
                            End If
                            
                            
                            
                            .Col = 0
                            .Text = datamain.Recordset.Fields("Serial_Num")
                            .Col = 1
                            .Text = datamain.Recordset.Fields("Order_Num")
                            .Col = 2
                            .Text = Format(datamain.Recordset.Fields("Customer"))
                            .Col = 3
                            .Text = datamain.Recordset.Fields("Production_phase0_Oprt")
                            .Col = 4
                            .Text = datamain.Recordset.Fields("Production_phase0_Date")  ', "yyyy/mm/dd")
                            .Col = 5
                            .Text = datamain.Recordset.Fields("Production_phase1_Oprt")
                            .Col = 6
                            .Text = datamain.Recordset.Fields("Production_phase1_Date")
                            .Col = 7
                            .Text = datamain.Recordset.Fields("Production_phase2_Oprt")
                            .Col = 8
                            .Text = datamain.Recordset.Fields("Production_phase2_Date")
                            .Col = 9
                            .Text = datamain.Recordset.Fields("Production_phase3_Oprt")
                            .Col = 10
                            .Text = datamain.Recordset.Fields("Production_phase3_Date")  ', "yyyy/mm/dd")
                            .Col = 11
                            .Text = datamain.Recordset.Fields("Production_phase4_Oprt")
                            .Col = 12
                            .Text = datamain.Recordset.Fields("Production_phase4_Date")
                            .Col = 13
                            .Text = datamain.Recordset.Fields("Production_phase5_Oprt")
                            .Col = 14
                            .Text = datamain.Recordset.Fields("Production_phase5_Date")
                            
                            .Col = 15
                            .Text = datamain.Recordset.Fields("Test1_Oprt")
                            .Col = 16
                            .Text = datamain.Recordset.Fields("Test1_Date")
                            .Col = 17
                            .Text = datamain.Recordset.Fields("Production_phase7_Oprt")  ', "yyyy/mm/dd")
                            .Col = 18
                            .Text = datamain.Recordset.Fields("Production_phase7_Date")
                            .Col = 19
                            .Text = datamain.Recordset.Fields("Production_phase6_Oprt")
                            .Col = 20
                            .Text = datamain.Recordset.Fields("Production_phase6_Date")
                            
                            .Col = 21
                            
                            If datamain.Recordset.Fields("Production_Completed") = True Then
                                .Text = " ò„Ì· ›—«Ì‰œ  Ê·Ìœ"
                                .CellBackColor = vbGreen
                            Else
                                .Text = "œ— Õ«·  Ê·Ìœ"
                                .CellBackColor = vbYellow
                           End If
                                 
                        End With
                        
                  Next cnti
                
                  '....................................................................................
                '........................////////////////Test Parameter ////////////.............................................
                Case 3
                    CmdprintGaranti.Enabled = False
                    cmdprint.Enabled = False
                    fgrpt.RIGHTTOLEFT = False
                    fgrpt.Font = Arial
                    If LastArrNumber = 0 Then
                            With fgrpt
                            .RowHeightMin = 100
                            .Rows = 1
                            .Cols = 9
                            colnum = 9
                            .ColWidth(0) = .Width / 10
                            .ColWidth(1) = .Width / 10
                            .ColWidth(2) = .Width / 10
                            .ColWidth(3) = .Width / 10
                            .ColWidth(4) = .Width / 10
                            .ColWidth(5) = .Width / 10
                            .ColWidth(6) = .Width / 10
                            .ColWidth(7) = .Width / 10
                            .ColWidth(8) = .Width / 10
                            .AllowUserResizing = flexResizeNone
                            .Row = 0
                            .Col = 0
                            .Text = "Serial No."
                            .Col = 1
                            .Text = "Uc"
                             .Col = 2
                            .Text = "Ur"
                            .Col = 3
                            .Text = "Iref"
                            .Col = 4
                            .Text = "Ic acceptable"
                            .Col = 5
                            .Text = "Uref min"
                            .Col = 6
                            .Text = "Uref max"
                             .Col = 7
                            .Text = "Ures min"
                            .Col = 8
                            .Text = "Ures max"
                        End With
                  End If
                  
                  For cnti = LastArrNumber + 1 To TotalArrNum
                        sn = srnrpt(cnti)
                        datamain.Recordset.Find "Serial_Num=" + Str(sn), , adSearchForward, 1
                        With fgrpt
                            .Rows = .Rows + 1
                            .Row = .Rows - 1
                            .Col = 0
                            .Text = datamain.Recordset.Fields("Serial_Num")
                            
                            datatestprmt.RecordSource = "Testprmt"
                            datatestprmt.Refresh
                            datatestprmt.Recordset.Find "ID=" + Str(datamain.Recordset.Fields("Test_Inf")), , adSearchForward, 1
                            If datatestprmt.Recordset.EOF = True Then
                                .Col = 1
                                .Text = "Test"
                                .Col = 2
                                .Text = "Information"
                                .Col = 3
                                .Text = "Not"
                                .Col = 4
                                .Text = "Found"
                                .Col = 5
                                .Text = ""
                                .Col = 6
                                .Text = ""
                                .Col = 7
                                .Text = ""
                                .Col = 8
                                .Text = ""
                            Else
                                .Col = 1
                                .Text = datatestprmt.Recordset.Fields("Ucon")
                                .Col = 2
                                .Text = datatestprmt.Recordset.Fields("Urated")
                                .Col = 3
                                .Text = datatestprmt.Recordset.Fields("Iref")
                                .Col = 4
                                .Text = datatestprmt.Recordset.Fields("Icaccp")
                                .Col = 5
                                .Text = datatestprmt.Recordset.Fields("Urefmin")
                                .Col = 6
                                .Text = datatestprmt.Recordset.Fields("Urefmax")
                                .Col = 7
                                .Text = datatestprmt.Recordset.Fields("Uresmin")
                                .Col = 8
                                .Text = datatestprmt.Recordset.Fields("Uresmax")
                            End If
                        End With
                        
                  Next cnti
                  
                  
            End Select
        End If
        
    fgrpt.Visible = True
    frmrpt.MousePointer = 0
        
    
Exit Sub
    
errhdnl:
    fgrpt.Visible = True
    frmrpt.MousePointer = 0
    MsgBox "Error : " + Str(err.Number) + Chr(13) + Chr(13) + err.Description, vbOKOnly, "Error"
    txtsn1.Text = ""
    txtsn2.Text = ""
    opt(0).Value = True
    fram2.Enabled = False
End Sub

Private Sub cmdclear_Click()

TotalArrNum = 0
ReDim aryprt(1 To 6, 1 To 1)
ReDim aryprmt(1 To 9, 1 To 1)
ReDim arytype(1 To 1)
ReDim aryGrntDate(1 To 1)
ReDim AryTavanirCode(1 To 1)
ReDim aryTraceCode(1 To 1)
ReDim aryGrntAllowed(1 To 1)
ReDim TotalRegsSerNo(1 To 1)
  
    
applysn = False
cmdprint.Enabled = False
CmdprintGaranti.Enabled = False

With fgrpt
    .RowHeightMin = 100
    .Rows = 1
    .Cols = 12
    colnum = 12
    .ColWidth(0) = .Width * 0.08
    .ColWidth(1) = .Width * 0.08
    .ColWidth(2) = .Width * 0.08
    .ColWidth(3) = .Width * 0.08
    .ColWidth(4) = .Width * 0.08
    .ColWidth(5) = .Width * 0.08
    .ColWidth(6) = .Width * 0.15
    .ColWidth(7) = .Width * 0.08
    .ColWidth(8) = .Width * 0.08
    .ColWidth(9) = .Width * 0.08
    .ColWidth(10) = .Width * 0.08
    .ColWidth(11) = .Width * 0.15
                            
'.AllowUserResizing = flexResizeNone
    .Row = 0
    .Col = 0
    .Text = "Serial.No"
    .CellAlignment = 0
                            
    .Col = 1
    .Text = "Arrester Type"
    .Col = 2
    .Text = "Uref @ Iref (T.1)"
    .Col = 3
    .Text = "Ic @ Uc (T.1)"
    .Col = 4
    .Text = "PD @ 1.05*Uc (T.1)"
    .Col = 5
    .Text = "PD @ Ur (T.1)"
    .Col = 6
    .Text = "Result of Test 1"
    .Col = 7
    .Text = "Uref @ Iref (T.2)"
    .Col = 8
    .Text = "Ic @ Uc (T.2)"
    .Col = 9
    .Text = "PD @ 1.05*Uc (Test 2)"
    .Col = 10
    .Text = "PD @ Ur (T.2)"
    .Col = 11
    .Text = "Result of Test 2"
                            
                            
    For i = 0 To 11
        .Col = i
        .CellBackColor = &H8000000A
    Next i
End With

'opt(0).value = True
txtsn1.Text = ""
txtsn2.Text = ""
txtorderNo.Text = ""
opt(0).Value = True
fram2.Enabled = False
CmbSubOrder.Clear
CmbSubOrder.Enabled = False

End Sub

Private Sub cmdclose_Click()
Unload Me
FrmInterface.Show
End Sub

Private Sub cmdprint_Click()
'Dim x1, x2, x3, x4, x5, x6, x7, x8, x9, x10, x11, x12, x13, x14, x15x16, x17, x18 As Long
'Dim y1, y2, y3, y4, y5, y6, y7 As Long
'Dim colno, rowno, arrno As Double
Dim newarr As Boolean
Dim cop As Long

frmrpt.MousePointer = 11
newarr = False
'////////////////////////////main information//////////
If applysn = False Then
    
    MsgBox "Insert Serial Number Then Press 'ADD To List' to Able  Printing", vbExclamation + vbOKOnly, "Attention"
    frmrpt.MousePointer = 0
Else

   On Error GoTo errhand1
    
    prt1.ShowPrinter
    cop = prt1.Copies
    
    For i = 1 To cop
        Call printsub
    Next i

    frmrpt.MousePointer = 0
End If

Exit Sub

errhand1:

MsgBox "There was a problem with your printer." + Chr(13) + "Please install printer , then Try Again.", vbExclamation + vbOKOnly, "Error"
frmrpt.MousePointer = 0

End Sub




Private Sub CmdprintGaranti_Click()
Dim newarr As Boolean
Dim cop As Long

frmrpt.MousePointer = 11
newarr = False
'////////////////////////////main information//////////
'If applysn = False Then
    
    'MsgBox "Insert Serial Number Then Press 'ADD To List' to Able  Printing", vbExclamation + vbOKOnly, "Attention"
    'frmrpt.MousePointer = 0
'Else

   On Error GoTo errhand1
    
    prt1.ShowPrinter
    cop = prt1.Copies
    
    'For i = 1 To cop
        Call PrintGuaranteeSub
    'Next i

    frmrpt.MousePointer = 0
'End If

Exit Sub

errhand1:

MsgBox "There was a problem with your printer." + Chr(13) + "Please install printer , then Try Again.", vbExclamation + vbOKOnly, "Error"
frmrpt.MousePointer = 0

End Sub



Private Sub Form_Activate()
applysn = False
cmdprint.Enabled = False
CmdprintGaranti.Enabled = False
txtsn1.Visible = True
lblsn1.Caption = "FROM :"
lblsn2.Caption = "     TO :"
txtsn1.Text = ""
txtsn2.Text = ""
opt(0).Value = True
fram2.Enabled = False
TotalArrNum = 0
MsgBox " ÃÂ  ’Õ  ‰„«Ì‘ «ÿ·«⁄«  ° «» œ« ‰Ê⁄ „Õ’Ê· —« «‰ Œ«» Ê ”Å” ‘„«—Â »—‰«„Â —« À»  ‰„«ÌÌœ." + Chr(13) + "«ÿ·«⁄«  ‘„«—Â »—‰«„Â Ê ‰Ê⁄ „Õ’Ê· »«Ìœ œ—”  «‰ Œ«» ê—œœ. ", vbOKOnly, " ÊÃÂ"

End Sub

Private Sub Form_Load()
Dim he, wi As Long

PrintBarcode = False
applysn = False
cmdprint.Enabled = False
CmdprintGaranti.Enabled = False

' Bixolon Printer control
 
 InitControls


addnum = True
With frmrpt
 .Height = Screen.Height * 0.97
 .Width = Screen.Width * 0.9997
 .Left = 0
 .Top = 0
 .ScaleHeight = .Height
 .ScaleWidth = .Width
 he = .Height
 wi = .Width
 .icon = LoadPicture(mdl1.icon)
End With

With datamain
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.maindb + ";Jet OLEDB:Database Password=" + mdl1.MainDBPas
    .CommandType = adCmdTable
    .RecordSource = "DIST"
    .Refresh
    
End With

With datatestprmt
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.acldb + ";Jet OLEDB:Database Password=" + mdl1.AclDBPas
    .CommandType = adCmdTable
    .RecordSource = "Type"
    .Refresh
End With

With DataProduction
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.acldb + ";Jet OLEDB:Database Password=" + mdl1.AclDBPas
    .CommandType = adCmdTable
    .RecordSource = "production"
    .Refresh
End With

With DataProduction.Recordset
    If .RecordCount = 0 Then
        .addnew
        !OrderNo = 9600
        .Update
    End If
End With

With datatestprmt.Recordset
    ReDim Tabletype(0 To .RecordCount - 1) As String   ' define Table name in main data base
    .MoveFirst
    If .RecordCount >= 1 Then
        
        For i = 1 To .RecordCount
            cmbtype.AddItem .Fields("arr_Type"), i - 1
             Tabletype(i - 1) = .Fields("TableName")
            .MoveNext
        Next i
        cmbtype.ListIndex = 0
    End If
End With

If Tabletype(cmbtype.ListIndex) = "SUBST" Then  '  Station Class Arrester Coding
    datamain.RecordSource = "SUBST"
  
ElseIf Tabletype(cmbtype.ListIndex) = "DIST" Then '  dISTRIBUTION Class Arrester Coding
    datamain.RecordSource = "DIST"
   
End If


'///////////////////////fram1//////////////////////
'???????????????????????????????????

'//////////////////////////////////////////////////

With fram1
    .Width = wi * 0.6
    .Height = he * 0.2
    .Left = wi * 0.05
    .Top = he * 0.035
    
End With

With lblArType
    .Width = fram1.Width * 0.2
    .Left = fram1.Width * 0.05
    .Top = fram1.Height * 0.3
    .Height = fram1.Height * 0.2
End With
With cmbtype
    .Width = fram1.Width * 0.15
    .Left = fram1.Width * 0.05
    .Top = fram1.Height * 0.6
    .Visible = True
End With


With lblOrderNo
    .Width = fram1.Width * 0.08
    .Height = fram1.Height * 0.2
    .Left = fram1.Width * 0.25
    .Top = fram1.Height * 0.3
    .Visible = True
    .Caption = "Order No."
End With
With txtorderNo
    .Width = fram1.Width * 0.16
    .Height = fram1.Height * 0.2
    .Left = fram1.Width * 0.35
    .Top = fram1.Height * 0.3
    .Visible = True
End With

With LblSubOrder
    .Width = fram1.Width * 0.08
    .Height = fram1.Height * 0.2
    .Left = fram1.Width * 0.25
    .Top = fram1.Height * 0.6
    .Visible = True
    .Visible = True
End With
With CmbSubOrder
    .Width = fram1.Width * 0.1
    '.Height = fram1.Height * 0.2
    .Left = fram1.Width * 0.35
    .Top = fram1.Height * 0.6
    .Visible = True
    .Clear
    .Enabled = False
End With


With lblsn1
    .Width = fram1.Width * 0.08
    .Height = fram1.Height * 0.2
    .Left = fram1.Width * 0.55
    .Top = fram1.Height * 0.3
    .Visible = True
    .Caption = "FROM:"
End With
With txtsn1
    .Width = fram1.Width * 0.16
    .Height = fram1.Height * 0.2
    .Left = fram1.Width * 0.6
    .Top = fram1.Height * 0.28
    .Visible = True
End With

With lblsn2
    .Width = fram1.Width * 0.08
    .Height = fram1.Height * 0.2
    .Left = fram1.Width * 0.55
    .Top = fram1.Height * 0.6
    .Visible = True
    .Caption = "     TO:"
End With
With txtsn2
    .Width = fram1.Width * 0.16
    .Height = fram1.Height * 0.15
    .Left = fram1.Width * 0.6
    .Top = fram1.Height * 0.58
    .Visible = True
End With
With cmdapply
    .Width = fram1.Width * 0.12
    .Height = fram1.Height * 0.2
    .Left = fram1.Width * 0.85
    .Top = fram1.Height * 0.27
End With
With cmdclear
    .Width = fram1.Width * 0.12
    .Height = fram1.Height * 0.2
    .Left = fram1.Width * 0.85
    .Top = fram1.Height * 0.59
    .Enabled = True
End With
'////////////////////fram2////////////////////////
With fram2
    .Width = wi * 0.25
    .Height = he * 0.2
    .Left = wi * 0.7
    .Top = he * 0.035
End With
With opt(0)
    .Width = fram2.Width * 0.4
    .Height = fram2.Height * 0.2
    .Left = fram2.Width * 0.08
    .Top = fram2.Height * 0.2
    .Value = True
End With
With opt(1)
    .Width = fram2.Width * 0.4
    .Height = fram2.Height * 0.2
    .Left = fram2.Width * 0.08
    .Top = fram2.Height * 0.6
End With
With opt(2)
    .Width = fram2.Width * 0.4
    .Height = fram2.Height * 0.2
    .Left = fram2.Width * 0.5
    .Top = fram2.Height * 0.2
End With
With opt(3)
    .Width = fram2.Width * 0.4
    .Height = fram2.Height * 0.2
    .Left = fram2.Width * 0.5
    .Top = fram2.Height * 0.6
End With



'////////////////////////////////////////

With cmbPrtList
    .Width = wi * 0.18
    '.Height = he * 0.06
    .Left = wi * 0.15
    .Top = he * 0.89
End With

With cmdprint
    .Width = wi * 0.12
    .Height = he * 0.06
    .Left = wi * 0.6
    .Top = he * 0.89
End With

With CmdprintGaranti
    .Width = wi * 0.12
    .Height = he * 0.06
    .Left = wi * 0.4
    .Top = he * 0.89
End With


With cmdclose
    .Width = wi * 0.12
    .Height = he * 0.06
    .Left = wi * 0.8
    .Top = he * 0.89
End With
'/////////////////////////////////////grid setting///////////////////
With fgrpt
    .Width = frmrpt.ScaleWidth * 0.9
    .Left = frmrpt.ScaleWidth * 0.05
    .Top = frmrpt.ScaleHeight * 0.28
    .Height = frmrpt.ScaleHeight * 0.6
End With
',,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,setting for test report/////////
NoOfSubprog = 0
With fgrpt
    .RowHeightMin = 100
    .Rows = 1
    .Cols = 12
    colnum = 12
    .ColWidth(0) = .Width * 0.07
    .ColWidth(1) = .Width * 0.07
    .ColWidth(2) = .Width * 0.07
    .ColWidth(3) = .Width * 0.07
    .ColWidth(4) = .Width * 0.07
    .ColWidth(5) = .Width * 0.07
    .ColWidth(6) = .Width * 0.12
    .ColWidth(7) = .Width * 0.07
    .ColWidth(8) = .Width * 0.07
    .ColWidth(9) = .Width * 0.07
    .ColWidth(10) = .Width * 0.07
    .ColWidth(11) = .Width * 0.12
                            
'.AllowUserResizing = flexResizeNone
    .Row = 0
    .Col = 0
    .Text = "Serial.No"
    .CellAlignment = 0
                            
    .Col = 1
    .Text = "Arrester Type"
    .Col = 2
    .Text = "Uref @ Iref (T.1)"
    .Col = 3
    .Text = "Ic @ Uc (T.1)"
    .Col = 4
    .Text = "PD @ 1.05*Uc (T.1)"
    .Col = 5
    .Text = "PD @ Ur (T.1)"
    .Col = 6
    .Text = "Result of Test 1"
    .Col = 7
    .Text = "Uref @ Iref (T.2)"
    .Col = 8
    .Text = "Ic @ Uc (T.2)"
    .Col = 9
    .Text = "PD @ 1.05*Uc (Test 2)"
    .Col = 10
    .Text = "PD @ Ur (T.2)"
    .Col = 11
    .Text = "Result of Test 2"
                            
                            
    For i = 0 To 11
        .Col = i
        .CellBackColor = &H8000000A
    Next i
End With
opt(0).Value = True
fram2.Enabled = False
End Sub



Private Sub opt_Click(Index As Integer)
    
Dim allow, ended As Boolean
Dim txtsrn1, txtsrn2, temstr As String
Dim state, Sal, A As Integer
Dim arrcnt, fc, cnti As Long
Dim ProductionCompleted As Boolean

On Error GoTo errhdnl
    
fgrpt.Visible = False
fgrpt.Height = frmrpt.ScaleHeight * 0.01
frmrpt.MousePointer = 11
        
allow = True
applysn = False
cmdprint.Enabled = False
CmdprintGaranti.Enabled = False
        
       ' ///////////////////////////////////type of report ///////////
        'opt(0).value = True
       
        If opt(0).Value = True Then
          state = 0
        ElseIf opt(1).Value = True Then
          state = 1
        ElseIf opt(2).Value = True Then
          state = 2
        ElseIf opt(3).Value = True Then
          state = 3
        End If
    
        
                 
        fgrpt.Clear
        fgrpt.Rows = 1
        fgrpt.Cols = 1
                
        Select Case state
        Case 0 '///////////Test Result////////
                
            cmdprint.Enabled = True
            If cmbPrtList.Enabled Then CmdprintGaranti.Enabled = True
                    
            fgrpt.RIGHTTOLEFT = False
            
            fgrpt.Font = Arial
            fgrpt.Visible = False
                  
                  
                With fgrpt
                            .RowHeightMin = 100
                            .Rows = 1
                            .Cols = 12
                            colnum = 12
                            
                            .ColWidth(0) = .Width * 0.07
                            .ColWidth(1) = .Width * 0.08
                            .ColWidth(2) = .Width * 0.07
                            .ColWidth(3) = .Width * 0.07
                            .ColWidth(4) = .Width * 0.08
                            .ColWidth(5) = .Width * 0.06
                            .ColWidth(6) = .Width * 0.15
                            .ColWidth(7) = .Width * 0.07
                            .ColWidth(8) = .Width * 0.07
                            .ColWidth(9) = .Width * 0.08
                            .ColWidth(10) = .Width * 0.06
                            .ColWidth(11) = .Width * 0.17
                            
                            
                            
                            '.AllowUserResizing = flexResizeNone
                            .Row = 0
                            .Col = 0
                            .Text = "Serial.No"
                            .CellAlignment = 0
                            
                            .Col = 1
                            .Text = "Arrester Type"
                             .Col = 2
                            .Text = "Uref @ Iref (T.1)"
                            .Col = 3
                            .Text = "Ic @ Uc (T.1)"
                            .Col = 4
                            .Text = "PD @ 1.05*Uc (T.1)"
                            .Col = 5
                            .Text = "PD @ Ur (T.1)"
                            .Col = 6
                            .Text = "Result of Test 1"
                            .Col = 7
                            .Text = "Uref @ Iref (T.2)"
                            .Col = 8
                            .Text = "Ic @ Uc (T.2)"
                            .Col = 9
                            .Text = "PD @ 1.05*Uc (Test 2)"
                            .Col = 10
                            .Text = "PD @ Ur (T.2)"
                            .Col = 11
                            .Text = "Result of Test 2"
                            
                            
                            For i = 0 To 11
                                .Col = i
                                .CellBackColor = &H8000000A
                            Next i
                        End With
                                    
                                    
                    For cnti = 1 To TotalArrNum
                    
                        sn = srnrpt(cnti)
                        datamain.Recordset.Find "Serial_Num=" + Str(sn), , adSearchForward, 1
                        
                        With fgrpt
                              .Rows = .Rows + 1
                              .Row = .Rows - 1
                              .Col = 0
                              .Text = datamain.Recordset.Fields("Serial_Num")
                              .CellBackColor = &H8000000E
                              .CellAlignment = 0
                              
                              TotalRegsSerNo(cnti) = Val(.Text)
                              
                              aryprt(1, cnti) = datamain.Recordset.Fields("Serial_Num")
                              
                              arytype(cnti) = datamain.Recordset.Fields("Arrester_Type")
                              
                              .Col = 1
                              .Text = datamain.Recordset.Fields("Arrester_Type")
                              .CellBackColor = &H8000000E
                                         
                              .Col = 2
                              .Text = Str(datamain.Recordset.Fields("Urefiref1"))
                              .CellBackColor = &H8000000E
                              
                              .Col = 3
                              .Text = Str(datamain.Recordset.Fields("IcUc1"))
                              .CellBackColor = &H8000000E
                              
                              .Col = 4
                              .Text = Str(datamain.Recordset.Fields("pduc1"))
                              .CellBackColor = &H8000000E
                              
                              .Col = 5
                              .Text = Str(datamain.Recordset.Fields("pdur1"))
                              .CellBackColor = &H8000000E
                              
                               '  ‰ŸÌ„  «—ÌŒ ‘—Ê⁄ ê«—«‰ Ì «“ “„«‰ »” Â »‰œÌ
                              aryGrntDate(cnti) = datamain.Recordset.Fields("Production_phase7_Date")
                               '  ‰ŸÌ„ « „«„  Ê·Ìœ
                              aryGrntAllowed(cnti) = datamain.Recordset.Fields("Production_Completed")
                              
                              aryTraceCode(cnti) = datamain.Recordset.Fields("Tracing_Code")
                              
                              ' òœ  Ê«‰Ì— „Õ’Ê·
                              AryTavanirCode(cnti) = datamain.Recordset.Fields("CompleteTavanirCode")
                              
                              If datamain.Recordset.Fields("Result2") <= 3 Then  'Test2 not applied & test1 is reference
                                
                                aryprt(2, cnti) = Str(datamain.Recordset.Fields("Urefiref1"))
                                aryprt(3, cnti) = Str(datamain.Recordset.Fields("IcUc1"))
                                aryprt(4, cnti) = Str(datamain.Recordset.Fields("pduc1"))
                                aryprt(5, cnti) = Str(datamain.Recordset.Fields("pdur1"))
                                
                              Else    ' test 2 is reference
                              
                                aryprt(2, cnti) = Str(datamain.Recordset.Fields("Urefiref2"))
                                aryprt(3, cnti) = Str(datamain.Recordset.Fields("IcUc2"))
                                aryprt(4, cnti) = Str(datamain.Recordset.Fields("pduc2"))
                                aryprt(5, cnti) = Str(datamain.Recordset.Fields("pdur2"))
                              
                              End If
                              
                              .Col = 6
                              .CellBackColor = &H8000000E
                              
                              Select Case Str(datamain.Recordset.Fields("Result1"))
                              Case 0
                                .Text = "Not Applied"
                              Case 1
                                  .Text = "Production Process isn't Completed"
                              Case 2
                                  .Text = "Ready for HV Test"
                              Case 3
                                  .Text = "Test 1 is OK"
                              Case 4
                                  .Text = "Uref 1 is High"
                              Case 5
                                  .Text = "Uref 1 is Low"
                              Case 6
                                  .Text = "Iref 1 is Failed"
                              Case 7
                                  .Text = "PD@Uc 1 is failed"
                              Case 8
                                  .Text = "PD@Ur 1 is failed"
                              End Select
                              
                              .Col = 7
                              .Text = Str(datamain.Recordset.Fields("Urefiref2"))
                              
                              .CellBackColor = &H80000003
                              
                              .Col = 8
                              .Text = Str(datamain.Recordset.Fields("IcUc2"))
                              
                              .CellBackColor = &H80000003
                              
                              .Col = 9
                              .Text = Str(datamain.Recordset.Fields("pduc2"))
                             
                              .CellBackColor = &H80000003
                              
                              .Col = 10
                              .Text = Str(datamain.Recordset.Fields("pdur2"))
                              
                              .CellBackColor = &H80000003
                              
                              .Col = 11
                              .CellBackColor = &H80000003
                              Select Case Str(datamain.Recordset.Fields("Result2"))
                              Case 0
                                .Text = "Not Applied"
                              Case 1
                                  .Text = "Not Applied"
                              Case 2
                                  .Text = "Test1 is OK ,Test2 don't Required"
                              Case 3
                                  .Text = "Test 2 is OK"
                              Case 4
                                  .Text = "Uref 2 is High"
                              Case 5
                                  .Text = "Uref 2 is Low"
                              Case 6
                                  .Text = "Iref 2 is Failed"
                              Case 7
                                  .Text = "PD@Uc 2 is failed"
                              Case 8
                                  .Text = "PD@Ur 2 is failed"
                              End Select
                                                     
                              
                          End With
                              
                      
                          '"""""""""""""""'test parameter   ''''''
                             
                              datatestprmt.RecordSource = "Testprmt"
                              datatestprmt.Refresh
                              datatestprmt.Recordset.Find "ID=" + Str(datamain.Recordset.Fields("Test_Inf")), , adSearchForward, 1
                              
                              If datatestprmt.Recordset.EOF = False Then
                                  aryprmt(1, cnti) = Val(Format(datatestprmt.Recordset.Fields("Iref"), "##.##"))
                                  aryprmt(2, cnti) = datatestprmt.Recordset.Fields("Urefmin")
                                  aryprmt(3, cnti) = datatestprmt.Recordset.Fields("Urefmax")
                                  aryprmt(4, cnti) = datatestprmt.Recordset.Fields("Ucon")
                                  aryprmt(5, cnti) = datatestprmt.Recordset.Fields("Icaccp")
                                  aryprmt(6, cnti) = datatestprmt.Recordset.Fields("Uresmin")
                                  aryprmt(7, cnti) = datatestprmt.Recordset.Fields("Uresmax")
                                  aryprmt(8, cnti) = datatestprmt.Recordset.Fields("Urated")
                                  aryprmt(9, cnti) = datatestprmt.Recordset.Fields("ID")
                              End If
                              
                Next cnti
                 
                 ',,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
                Case 1              '///////////BLock Information////////////////
                
                    fgrpt.RIGHTTOLEFT = False
                    CmdprintGaranti.Enabled = False
                    cmdprint.Enabled = False
                    
                    fgrpt.Font = Arial
                    fgrpt.Visible = False
                    
                    With datatestprmt
                        .RecordSource = "Testprmt"
                        .Refresh
                     End With
                    
                    With fgrpt
                             
                             .Rows = 1
                             .RowHeightMin = 100
                             .Cols = 6
                             colnum = 6
                             
                             .ColWidth(0) = .Width * 0.12
                             .ColWidth(1) = .Width * 0.12
                             .ColWidth(2) = .Width * 0.12
                             .ColWidth(3) = .Width * 0.12
                             .ColWidth(4) = .Width * 0.12
                             .ColWidth(5) = .Width * 0.12
                             '.AllowUserResizing = flexResizeNone
                             .Row = 0
                             .Col = 0
                             .Text = "Serial Number of Arrester"
                             .CellAlignment = 2
                             .Col = 1
                             .Text = "Manufacturer of MOV"
                             .CellAlignment = 2
                              .Col = 2
                             .Text = "Quantity of Type1"
                             .CellAlignment = 2
                             .Col = 3
                             .Text = "Dimension of Type1 (D*H)"
                             .CellAlignment = 2
                             .Col = 4
                             .Text = "Quantity of Type2"
                             .CellAlignment = 2
                             .Col = 5
                             .Text = "Dimension of Type2 (D*H)"
                             .CellAlignment = 2
                            
                            For i = 0 To 5
                                .Col = i
                                .CellBackColor = &H8000000A
                            Next i
                            
                    End With
                    
                    For cnti = 1 To TotalArrNum
                        sn = srnrpt(cnti)
                        datamain.Recordset.Find "Serial_Num=" + Str(sn), , adSearchForward, 1
                        With fgrpt
                            .Rows = .Rows + 1
                            .Row = .Rows - 1
                            .Col = 0
                            .Text = datamain.Recordset.Fields("Serial_Num")
                            .CellAlignment = 2
                            
                            TotalRegsSerNo(cnti) = Val(.Text)
                            datatestprmt.Recordset.Find "ID=" + Str(datamain.Recordset.Fields("Test_Inf")), , adSearchForward, 1
                            If datatestprmt.Recordset.EOF = True Then
                                .Col = 1
                                .Text = "Not Found "
                                .Col = 3
                                .Text = "Test "
                                .Col = 2
                                .Text = "Information"
                                .Col = 5
                                .Text = ""
                                .Col = 4
                                .Text = ""
                                
                            Else
                                .Col = 1
                                .Text = datamain.Recordset.Fields("Block_Type")  ' manufacturer
                                .CellAlignment = 2
                                .Col = 2
                                .Text = Val(datatestprmt.Recordset.Fields("Block_Size1"))
                                .CellAlignment = 2
                                .Col = 3
                                .Text = Right(datatestprmt.Recordset.Fields("Block_Size1"), 7)
                                .CellAlignment = 2
                                .Col = 4
                                .Text = Val(datatestprmt.Recordset.Fields("Block_Size2"))
                                .CellAlignment = 2
                                .Col = 5
                                .Text = Right(datatestprmt.Recordset.Fields("Block_Size2"), 7)
                                .CellAlignment = 2
                               ' .CellBackColor = &H80000003
                            End If
                        End With
                        
                   Next cnti
                   
                ';;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
                
                Case 2             '/////////////Product Information///////////
                
                    fgrpt.RIGHTTOLEFT = True
                    CmdprintGaranti.Enabled = False
                    cmdprint.Enabled = False
                    fgrpt.Visible = False
                    
                    ' Add Title of production process name
                    
                    datatestprmt.RecordSource = "Type"
                    datatestprmt.Refresh
                
                    datatestprmt.Recordset.MoveFirst
                    datatestprmt.Recordset.Find ("Arr_Type='") + cmbtype.Text + "'", , adSearchForward, 1
                              
                                                
                          With fgrpt
                            
                            .RowHeightMin = 100
                            .Rows = 1
                            .Cols = 22
                            colnum = 22
                            
                            fgrpt.Font = "B Koodak"
                            For i = 0 To 1
                                .ColWidth(i) = .Width * 0.06
                            Next i
                            
                            .ColWidth(2) = .Width * 0.15
                            
                            For i = 3 To 21
                                .ColWidth(i) = .Width * 0.08
                            Next i
                            
                            For i = 0 To 21
                                .Col = i
                                .CellBackColor = &H8000000A
                            Next i
                            
                            .AllowUserResizing = flexResizeNone
                            .Row = 0
                            
                            .Col = 0
                            .Text = "‘„«—Â ”—Ì«·"
                            .Col = 1
                            .Text = "‘„«—Â »—‰«„Â"
                             .Col = 2
                            .Text = "‰«„ „‘ —Ì"
                            .Col = 3
                            .Text = "«Å—« Ê— " + datatestprmt.Recordset.Fields("Production_Phase0")
                            .Col = 4
                            .Text = " «—ÌŒ " + datatestprmt.Recordset.Fields("Production_Phase0")
                            .Col = 5
                            .Text = "«Å—« Ê— " + datatestprmt.Recordset.Fields("Production_Phase1")
                            .Col = 6
                            .Text = " «—ÌŒ " + datatestprmt.Recordset.Fields("Production_Phase1")
                            .Col = 7
                            .Text = "«Å—« Ê— " + datatestprmt.Recordset.Fields("Production_Phase2")
                             .Col = 8
                            .Text = " «—ÌŒ " + datatestprmt.Recordset.Fields("Production_Phase2")
                            .Col = 9
                            .Text = "«Å—« Ê— " + datatestprmt.Recordset.Fields("Production_Phase3")
                            .Col = 10
                            .Text = " «—ÌŒ " + datatestprmt.Recordset.Fields("Production_Phase3")
                            .Col = 11
                            .Text = "«Å—« Ê— " + datatestprmt.Recordset.Fields("Production_Phase4")
                            .Col = 12
                            .Text = " «—ÌŒ " + datatestprmt.Recordset.Fields("Production_Phase4")
                            .Col = 13
                            .Text = "«Å—« Ê— " + datatestprmt.Recordset.Fields("Production_Phase5")
                             .Col = 14
                            .Text = " «—ÌŒ " + datatestprmt.Recordset.Fields("Production_Phase5")
                            .Col = 15
                            .Text = "«Å—« Ê—  ”  «·ò —ÌòÌ"
                            .Col = 16
                            .Text = " «—ÌŒ  ”  «·ò —ÌòÌ"
                             .Col = 17
                            .Text = "«Å—« Ê— " + datatestprmt.Recordset.Fields("Production_Phase7")
                            .Col = 18
                            .Text = " «—ÌŒ " + datatestprmt.Recordset.Fields("Production_Phase7")
                            .Col = 19
                            .Text = "«Å—« Ê— " + datatestprmt.Recordset.Fields("Production_Phase6")
                            .Col = 20
                            .Text = " «—ÌŒ " + datatestprmt.Recordset.Fields("Production_Phase6")
                            .Col = 21
                            .Text = "Ê÷⁄Ì   Ê·Ìœ"
                            
                        End With
                 
                 
                 
                 
                  
                  For cnti = 1 To TotalArrNum
                        sn = srnrpt(cnti)
                        datamain.Recordset.Find "Serial_Num=" + Str(sn), , adSearchForward, 1
                        
                        
                        With fgrpt
                            .Rows = .Rows + 1
                            .Row = .Rows - 1
                            
                            If datamain.Recordset.Fields("Production_Completed") = True Then
                                .Text = " ò„Ì· ›—«Ì‰œ  Ê·Ìœ"
                                For i = 0 To 21
                                    .Col = i
                                    .CellBackColor = vbGreen
                                Next i
                            Else
                                .Text = "œ— Õ«·  Ê·Ìœ"
                                For i = 0 To 21
                                    .Col = i
                                    .CellBackColor = vbYellow
                                Next i
                            End If
                            
                            
                            
                            .Col = 0
                            .Text = datamain.Recordset.Fields("Serial_Num")
                            .Col = 1
                            .Text = datamain.Recordset.Fields("Order_Num")
                            .Col = 2
                            .Text = Format(datamain.Recordset.Fields("Customer"))
                            .Col = 3
                            .Text = datamain.Recordset.Fields("Production_phase0_Oprt")
                            .Col = 4
                            .Text = datamain.Recordset.Fields("Production_phase0_Date")  ', "yyyy/mm/dd")
                            .Col = 5
                            .Text = datamain.Recordset.Fields("Production_phase1_Oprt")
                            .Col = 6
                            .Text = datamain.Recordset.Fields("Production_phase1_Date")
                            .Col = 7
                            .Text = datamain.Recordset.Fields("Production_phase2_Oprt")
                            .Col = 8
                            .Text = datamain.Recordset.Fields("Production_phase2_Date")
                            .Col = 9
                            .Text = datamain.Recordset.Fields("Production_phase3_Oprt")
                            .Col = 10
                            .Text = datamain.Recordset.Fields("Production_phase3_Date")  ', "yyyy/mm/dd")
                            .Col = 11
                            .Text = datamain.Recordset.Fields("Production_phase4_Oprt")
                            .Col = 12
                            .Text = datamain.Recordset.Fields("Production_phase4_Date")
                            .Col = 13
                            .Text = datamain.Recordset.Fields("Production_phase5_Oprt")
                            .Col = 14
                            .Text = datamain.Recordset.Fields("Production_phase5_Date")
                            
                            .Col = 15
                            .Text = datamain.Recordset.Fields("Test1_Oprt")
                            .Col = 16
                            .Text = datamain.Recordset.Fields("Test1_Date")
                            .Col = 17
                            .Text = datamain.Recordset.Fields("Production_phase7_Oprt")  ', "yyyy/mm/dd")
                            .Col = 18
                            .Text = datamain.Recordset.Fields("Production_phase7_Date")
                            .Col = 19
                            .Text = datamain.Recordset.Fields("Production_phase6_Oprt")
                            .Col = 20
                            .Text = datamain.Recordset.Fields("Production_phase6_Date")
                            
                            .Col = 21
                            
                            If datamain.Recordset.Fields("Production_Completed") = True Then
                                .Text = " ò„Ì· ›—«Ì‰œ  Ê·Ìœ"
                                .CellBackColor = vbGreen
                            Else
                                .Text = "œ— Õ«·  Ê·Ìœ"
                                .CellBackColor = vbYellow
                           End If
                                 
                        End With
                        
                  Next cnti
                
                  '....................................................................................
                '........................////////////////Test Parameter ////////////.............................................
                Case 3
                    CmdprintGaranti.Enabled = False
                    cmdprint.Enabled = False
                    fgrpt.RIGHTTOLEFT = False
                    fgrpt.Font = Arial
                    fgrpt.Visible = False
                    'If LastArrNumber = 0 Then
                            With fgrpt
                            .RowHeightMin = 100
                            .Rows = 1
                            .Cols = 9
                            colnum = 9
                            .ColWidth(0) = .Width / 10
                            .ColWidth(1) = .Width / 10
                            .ColWidth(2) = .Width / 10
                            .ColWidth(3) = .Width / 10
                            .ColWidth(4) = .Width / 10
                            .ColWidth(5) = .Width / 10
                            .ColWidth(6) = .Width / 10
                            .ColWidth(7) = .Width / 10
                            .ColWidth(8) = .Width / 10
                            .AllowUserResizing = flexResizeNone
                            .Row = 0
                            .Col = 0
                            .Text = "Serial No."
                            .Col = 1
                            .Text = "Uc"
                             .Col = 2
                            .Text = "Ur"
                            .Col = 3
                            .Text = "Iref"
                            .Col = 4
                            .Text = "Ic acceptable"
                            .Col = 5
                            .Text = "Uref min"
                            .Col = 6
                            .Text = "Uref max"
                             .Col = 7
                            .Text = "Ures min"
                            .Col = 8
                            .Text = "Ures max"
                        End With
                  
                 For cnti = 1 To TotalArrNum
                        sn = srnrpt(cnti)
                        datamain.Recordset.Find "Serial_Num=" + Str(sn), , adSearchForward, 1
                        With fgrpt
                            .Rows = .Rows + 1
                            .Row = .Rows - 1
                            .Col = 0
                            .Text = datamain.Recordset.Fields("Serial_Num")
                            
                            datatestprmt.RecordSource = "Testprmt"
                            datatestprmt.Refresh
                            datatestprmt.Recordset.Find "ID=" + Str(datamain.Recordset.Fields("Test_Inf")), , adSearchForward, 1
                            If datatestprmt.Recordset.EOF = True Then
                                .Col = 1
                                .Text = "Test"
                                .Col = 2
                                .Text = "Information"
                                .Col = 3
                                .Text = "Not"
                                .Col = 4
                                .Text = "Found"
                                .Col = 5
                                .Text = ""
                                .Col = 6
                                .Text = ""
                                .Col = 7
                                .Text = ""
                                .Col = 8
                                .Text = ""
                            Else
                                .Col = 1
                                .Text = datatestprmt.Recordset.Fields("Ucon")
                                .Col = 2
                                .Text = datatestprmt.Recordset.Fields("Urated")
                                .Col = 3
                                .Text = datatestprmt.Recordset.Fields("Iref")
                                .Col = 4
                                .Text = datatestprmt.Recordset.Fields("Icaccp")
                                .Col = 5
                                .Text = datatestprmt.Recordset.Fields("Urefmin")
                                .Col = 6
                                .Text = datatestprmt.Recordset.Fields("Urefmax")
                                .Col = 7
                                .Text = datatestprmt.Recordset.Fields("Uresmin")
                                .Col = 8
                                .Text = datatestprmt.Recordset.Fields("Uresmax")
                            End If
                        End With
                        
                  Next cnti
                  
        End Select
        
    
    frmrpt.MousePointer = 0
     fgrpt.Visible = True
     fgrpt.Height = frmrpt.ScaleHeight * 0.6
    Exit Sub
    
errhdnl:
     fgrpt.Visible = True
     fgrpt.Height = frmrpt.ScaleHeight * 0.6
    frmrpt.MousePointer = 0
    MsgBox "Error : " + Str(err.Number) + Chr(13) + Chr(13) + err.Description, vbOKOnly, "Error"
    txtsn1.Text = ""
    txtsn2.Text = ""
       

End Sub

Private Sub txtorderNo_Change()
txtsn1.Text = ""
txtsn2.Text = ""
opt(0).Value = True
fram2.Enabled = False



End Sub

Private Sub txtorderNo_KeyPress(KeyAscii As Integer)

KeyAscii = mdl1.inputtest1(KeyAscii)

If KeyAscii = 13 Then
    cmdapply.SetFocus
End If

End Sub


Private Sub txtOrderNo_LostFocus()

If Val(txtorderNo.Text) > 0 Or txtorderNo.Text <> Null Then
    With DataProduction.Recordset
       .Find "OrderNo=" + txtorderNo.Text, , adSearchForward, 1
        If .EOF = False Then
            txtsn1.Text = .Fields("FirstSN")
            txtsn2.Text = .Fields("LastSN")
            
            FirstSNofOrder = Val(txtsn1.Text)
            LastSNofOrder = Val(txtsn2.Text)
            
            NoOfSubprog = .Fields("NoOfSubDivision")
            'totalquantity = .Fields("Quantity")
            'txttotalquantity.Text = Str(totalquantity)
            CmbSubOrder.Enabled = True
            CmbSubOrder.Clear
            For cnt = 1 To NoOfSubprog
                CmbSubOrder.AddItem Str(cnt), cnt - 1
            Next cnt
            CmbSubOrder.AddItem "All", cnt - 1
            
            CmbSubOrder.ListIndex = 0
            
            If NoOfSubprog > 1 Then
                CmbSubOrder.Enabled = True
            Else
                CmbSubOrder.Enabled = False
            End If
            
            cmdapply.SetFocus
        End If
    End With
Else
    txtsn1.SetFocus
    FirstSNofOrder = 0
    LastSNofOrder = 0

End If

End Sub


Private Sub txtsn1_KeyPress(KeyAscii As Integer)
KeyAscii = mdl1.inputtest1(KeyAscii)
If KeyAscii = 13 Then
    txtsn2.SetFocus
End If
End Sub

Private Sub txtsn2_Change()
applysn = False

End Sub


Private Sub txtsn2_KeyPress(KeyAscii As Integer)

KeyAscii = mdl1.inputtest1(KeyAscii)

If KeyAscii = 13 Then
    cmdapply.SetFocus
End If


End Sub
