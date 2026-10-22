Attribute VB_Name = "mdl1"
Option Explicit
Public uref, ucoun, urt, iref, ic, unitnumber, ilaccp, inom, pduc1, pdur1, pdaccp, uresmax, uresmin, urefmin, Acpt_PD, urefmax, tem, optcode, blockid, ordernum As Double
Public ArresterType, ArrTableName, selectedarr, optname, blockinf, blockname, access As String
Public Test_Admin, Technical_Manager, QC_Manager, QC_Operator, HVTest_Operator, Production_Worker, isHVArrester As Boolean ' set access of user
Public maindb, acldb, icon, pic1, pic2, ARTMQC_Logo, MainDBPas, AclDBPas, VacDBAdress, VacDBpas As String
Public year, FullYear, blcnum, sncode, svr1, svr2 As Integer
Public trs1, trs2, trs3, trs4 As Integer
Public sealcon() As Integer
Public srn() As Double
Public newsrn() As Double
Public newarrnum, NewSubOrders, GLNCode, CompleteTavanirCode As Double
Public tested() As Boolean
Public arrnum As Double
Public testrun, testprog As Boolean
Public blocks() As String
Public preform As String
Public i, rtv, j, ini As Integer
Dim errnum As Long

Public ProgramType As String

Public sht As Object  ' excel file object use in Loading test
Public cansave As Boolean


Public Function inputtest(key As Integer) As Integer
Dim rtv As Integer
'&&&&&&&&&&&&&&&&&&&    accept (.) value
If key > 32 And key <> 48 And key <> 49 And key <> 50 And key <> 51 And key <> 52 And key <> 53 And key <> 54 And key <> 55 And key <> 56 And key <> 57 And key <> 13 And key <> 46 Then
    rtv = MsgBox("PLease Enter Integer Value ", vbOKOnly, "ERROR")
    inputtest = 0
Else
    inputtest = key
End If
End Function
Public Function inputtest1(key As Integer) As Integer
Dim rtv As Integer
'&&&&&&&&&&&&&&&&&&&    donot accept (.) value
If key > 32 And key <> 48 And key <> 49 And key <> 50 And key <> 51 And key <> 52 And key <> 53 And key <> 54 And key <> 55 And key <> 56 And key <> 57 And key <> 13 Then
    rtv = MsgBox("PLease Enter Integer Value ", vbOKOnly, "ERROR")
    inputtest1 = 0
Else
    inputtest1 = key
End If
End Function
Public Sub Main()
Dim S, spd As String
Dim ytest As Boolean
Dim y, pd As Integer

'year = number of year For identify  of  main database database "HVtest##"
'maindb=addres of main database "HVtest##"
'acldb=addres of accelary database "hvtestdb" consist of (block & personel  & testprmt)  Tables
'pic# = addres of picture for img1 in maimform

mdl1.ProgramType = "HV Test Program"  ' for HV test Application
'mdl1.ProgramType = "QC Program"       'for QC Applicastion


mdl1.testprog = True
mdl1.pic1 = App.Path + "\ETP1.bmp"
mdl1.pic2 = App.Path + "\ETP2.bmp" '  For Print in Print Form
mdl1.icon = App.Path + "\ETPicon.ico"
mdl1.ARTMQC_Logo = App.Path + "\ARTM_QC Logo.jpg" '  For Print in Print Form

pdaccp = 10
'***************************************************************
With frmSplash.ado
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\Hvadres.mdb" + ";Persist Security Info=False"
            'describe Provider and file Address of main data base
    .CommandType = adCmdTable
    .RecordSource = "dbadres"  'Table name of data bade
    .Refresh
End With

With frmSplash.ado.Recordset

    If .RecordCount = 0 Then
        .addnew
        !mainadr = "E:\HVtest\Hvtest93.mdb"
        !acladr = "E:\HVtest\Hvtestdb.mdb"
        !AclDBPassword = "1837"
        !MainDBPassword = "13581837"
        !Vacuumaddress = "E:\HVtest\Hvtestdb.mdb"
        !VacummPassword = "13581837"
        !year = 1401
        !CompanyGLN = "36050"
        .Update
     End If
     .MoveFirst
     If .Fields("year") < 1394 Then
        ytest = False
        Do While ytest = False
            S = InputBox("Please Enter Operating Year." + Chr(13) + "Year  must be >= 1394", "Enter Year ", "1398", Screen.Width / 3, Screen.Height / 3)
            y = Val(S)
            If y < 1393 Or y > 1499 Then
                MsgBox "Enter Integer Number From [1394 - 1500] Margin.", vbOKOnly, "Error"
            Else
                rtv = MsgBox("Are you sure that Operating Year is :" + S + " ?", vbYesNo, "Confirm Year ")
                If rtv = 6 Then
                    ytest = True
                    !year = y
                    .Update
                End If
            End If
        Loop
        
        ytest = False
        Do While ytest = False
            spd = InputBox("Acceptable PD is 2 [Pc]." + Chr(13) + "If you Want to change Acceptance Criteria, Enter New Value." + Chr(13) + "Max Acceptance PD is : ", "Acceptance PD", "2", Screen.Width / 3, Screen.Height / 3)
            pd = Val(spd)
            If pd < 1 Or pd > 10 Then
                MsgBox "Enter Integer Number From the margin of [1 - 10].", vbOKOnly, "Error"
            Else
                rtv = MsgBox("Are you sure that Max Acceptance PD is : " + pd + "  [Pc] ?", vbYesNo, "Confirm Year Number")
                If rtv = 6 Then
                    ytest = True
                    !accept_pd = pd
                    .Update
                End If
            End If
        Loop
        !mainadr = "E:\HVtest\Hvtest" + S + ".mdb"
        !acladr = "E:\HVtest\Hvtestdb.mdb"
        !AclDBPassword = "1837"
        !MainDBPassword = "13581837"
        !Vacuumaddress = "E:\HVtest\Hvtestdb.mdb"
        !VacummPassword = "13581837"
        !CompanyGLN = "36050"
        .Update
    End If
    
    S = Right(Str(.Fields("year")), 2)
    mdl1.maindb = .Fields("mainadr")
    mdl1.acldb = .Fields("acladr")
    mdl1.MainDBPas = .Fields("MainDBPassword")
    mdl1.AclDBPas = .Fields("AclDBPassword")
    mdl1.VacDBAdress = .Fields("Vacuumaddress")
    mdl1.VacDBpas = .Fields("VacummPassword")
    
    mdl1.GLNCode = .Fields("CompanyGLN")
    
    mdl1.year = Val(Right(Str(.Fields("year")), 2))
    mdl1.FullYear = Val(Right(Str(.Fields("year")), 4))
    mdl1.pdaccp = .Fields("accept_pd")
    .Close
    
End With


'*****************************************************************
errnum = 0

Call opendb1

If errnum <> 0 Then
    Call opendb1
End If

If errnum = 0 Then
    With frmSplash.Ado1
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.acldb + ";Jet OLEDB:Database Password=" + mdl1.AclDBPas
        .CommandType = adCmdTable
        .RecordSource = "personel"
        .Refresh
        .Recordset.Close
    End With
    errnum = err.Number
End If

If errnum = 0 Then

    Call opendb2
    
    If errnum <> 0 Then
         Call opendb2
    End If
    
    If errnum = 0 Then
        With frmSplash.Ado2
            .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.maindb + ";Jet OLEDB:Database Password=" + mdl1.MainDBPas
            .CommandType = adCmdTable
            .RecordSource = "DIST"
            .Refresh
            .Recordset.Close
        End With
        errnum = err.Number
     End If
End If


If errnum = 0 Then

    Call opendbVac
    
    If errnum <> 0 Then
         Call opendbVac
    End If
    
    If errnum = 0 Then
        With frmSplash.Ado3
            .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.VacDBAdress + ";Jet OLEDB:Database Password=" + mdl1.VacDBpas
            .CommandType = adCmdTable
            .RecordSource = "STATION_CLASS"
            .Refresh
            .Recordset.Close
        End With
        errnum = err.Number
     End If
End If



If errnum = 0 Then
    
    
   With frmSplash.ado.Recordset
        .Open
        .MoveFirst
        !mainadr = mdl1.maindb
        !acladr = mdl1.acldb
        !Vacuumaddress = mdl1.VacDBAdress
        
        .Update
        .Close
    End With

    frmSplash.Timer1.Enabled = True
    frmSplash.Show

Else
    
    End
    
End If

End Sub
Public Sub YearChng()
Dim S, TemAdres As String
Dim ytest As Boolean
Dim y, TemYear As Integer

'************************************

ytest = False
Do While ytest = False
    S = InputBox("Please enter desired year for database." + Chr(13) + "Year  must be in the ragne of [1394-1500]", "Enter Year ", "1394", Screen.Width / 3, Screen.Height / 3)
    y = Val(S)
    If y < 1394 Or y > 1500 Then
        MsgBox "Enter Integer Number From [1394 - 1500] Margin.", vbOKOnly, "Error"
    Else
         ytest = True
         TemAdres = mdl1.maindb
         TemYear = mdl1.year
         mdl1.maindb = "E:\HVtest\Hvtest" + Right(S, 2) + ".mdb"
         mdl1.year = Val(Right(S, 2))
    End If
Loop
ytest = False
   
   
'*****************************************************************
errnum = 0

Call opendb2
    
If errnum <> 0 Then
        Call opendb2
End If
    
If errnum = 0 Then
    With frmSplash.Ado2
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.maindb + ";Jet OLEDB:Database Password=" + mdl1.MainDBPas
        .CommandType = adCmdTable
        .RecordSource = "DIST"
        .Refresh
        .Recordset.Close
    End With
    
    With frmSplash.ado.Recordset
        rtv = MsgBox("Are you sure that operating year is:  " + S + " ?", vbYesNo, "Confirm Year ")
            If rtv = 6 Then
                .Open
                .MoveFirst
                !mainadr = mdl1.maindb
                !year = y
                .Update
                S = Right(Str(.Fields("year")), 2)
                mdl1.maindb = .Fields("mainadr")
                mdl1.acldb = .Fields("acladr")
                mdl1.MainDBPas = .Fields("MainDBPassword")
                mdl1.AclDBPas = .Fields("AclDBPassword")
                mdl1.year = Right(Str(.Fields("year")), 2)
                mdl1.FullYear = Right(Str(.Fields("year")), 4)
               .Close
                
                'frmtest.LblYear(1).Enabled = True
                'frmtest.LblYear(1).Caption = "  Year :" + Str(mdl1.FullYear)
                'frmtest.LblYear(1).Enabled = False
                frmtest.Caption = "ARTM                        Current Year : " + Str(mdl1.FullYear) + "                         Operator : " + mdl1.optname
 
                MsgBox "The Year has been changed to " + Str(y) + ".", vbInformation, "Error"
                
            Else
            
                MsgBox "The Year has not been changed.", vbInformation, "Error"
                mdl1.maindb = TemAdres
                mdl1.year = TemYear
                
            End If
    End With
Else

    MsgBox "There is no required database file." + Chr(13) + "The Year has not been changed.", vbInformation, "Error"
    mdl1.maindb = TemAdres
    mdl1.year = TemYear
    errnum = err.Number
    
End If
    
End Sub

Public Sub fdim(num As Integer)
ReDim srn(1 To num) As Double
ReDim sealcon(1 To num) As Integer
ReDim tested(1 To num) As Boolean
arrnum = num
testrun = True
End Sub
Public Sub leftlable(maxscale As Double)
frmtest.picleft.Cls
frmtest.picleft.CurrentX = 0
frmtest.picleft.CurrentY = 0
frmtest.picleft.Print Format(Str(maxscale), "0.#")
frmtest.picleft.CurrentX = 0
frmtest.picleft.CurrentY = 1
frmtest.picleft.Print Format(Str(maxscale * 3 / 4), "0.#")
frmtest.picleft.CurrentX = 0
frmtest.picleft.CurrentY = 2
frmtest.picleft.Print Format(Str(maxscale * 2 / 4), "0.#");
frmtest.picleft.CurrentX = 0
frmtest.picleft.CurrentY = 2.99
frmtest.picleft.Print Format(Str(maxscale * 1 / 4), "0.#")
frmtest.picleft.CurrentX = 0
frmtest.picleft.CurrentY = 3.98
frmtest.picleft.Print Format(Str(maxscale * 0), "0.#")
frmtest.picleft.CurrentX = 0
frmtest.picleft.CurrentY = 4.98
frmtest.picleft.Print Format(Str(-maxscale * 1 / 4), "0.#")
frmtest.picleft.CurrentX = 0
frmtest.picleft.CurrentY = 5.98
frmtest.picleft.Print Format(Str(-maxscale * 2 / 4), "0.#")
frmtest.picleft.CurrentX = 0
frmtest.picleft.CurrentY = 6.97
frmtest.picleft.Print Format(Str(-maxscale * 3 / 4), "0.#")
frmtest.picleft.CurrentX = 0
frmtest.picleft.CurrentY = 7.945
frmtest.picleft.Print Format(Str(-maxscale), "0.#")

End Sub

Public Sub opendb1()
Dim errdesc As String

On Error GoTo errhndl1


With frmSplash.Ado1
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.acldb + ";Jet OLEDB:Database Password=" + mdl1.AclDBPas
    .CommandType = adCmdTable
    .RecordSource = "personel"
    .Refresh
    .Recordset.Close
End With
    
errhndl1:
    errnum = err.Number
    errdesc = err.Description
    
    If errnum <> 0 Then
        MsgBox "Error : " + Str(errnum) + Chr(13) + Chr(13) + errdesc + Chr(13) + Chr(13) + "Please Select Location Of : Hvtestdb.mdb  ", vbOKOnly, "Error"
        frmSplash.cdlg1.FileName = "Hvtestdb"
        frmSplash.cdlg1.Filter = "*.mdb|*.mdb"
        frmSplash.cdlg1.ShowOpen
        mdl1.acldb = frmSplash.cdlg1.FileName
    End If

End Sub
Public Sub opendb2()
Dim S, errdesc As String

S = Right(Str(mdl1.year), 2)


On Error GoTo errhndl1

With frmSplash.Ado2
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.maindb + ";Jet OLEDB:Database Password=" + mdl1.MainDBPas
    .CommandType = adCmdTable
    .RecordSource = "DIST"
    .Refresh
    .Recordset.Close
   
End With
    
errhndl1:
    errnum = err.Number
    errdesc = err.Description
    If errnum <> 0 Then
        MsgBox "Error : " + Str(errnum) + Chr(13) + Chr(13) + errdesc + Chr(13) + Chr(13) + "Please Select Location Of : Hvtest" + Str(mdl1.year) + ".mdb  ", vbOKOnly, "Error"
        frmSplash.cdlg1.FileName = "Hvtest" + S + ".mdb"
        frmSplash.cdlg1.Filter = "*.mdb|*.mdb"
        frmSplash.cdlg1.ShowOpen
        mdl1.maindb = frmSplash.cdlg1.FileName
   
    End If
    
End Sub


Public Sub opendbVac()
Dim S, errdesc As String

S = Right(Str(mdl1.year), 2)


On Error GoTo errhndl1

With frmSplash.Ado3
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.VacDBAdress + ";Jet OLEDB:Database Password=" + mdl1.VacDBpas
    .CommandType = adCmdTable
    .RecordSource = "STATION_CLASS"
    .Refresh
    .Recordset.Close
   
End With
    
errhndl1:
    errnum = err.Number
    errdesc = err.Description
    If errnum <> 0 Then
        MsgBox "Error : " + Str(errnum) + Chr(13) + Chr(13) + errdesc + Chr(13) + Chr(13) + "Please Select Location Of : Vacuum" + Str(mdl1.year) + ".mdb  ", vbOKOnly, "Error"
        frmSplash.cdlg1.FileName = "Vacuum" + S + ".mdb"
        frmSplash.cdlg1.Filter = "*.mdb|*.mdb"
        frmSplash.cdlg1.ShowOpen
        mdl1.VacDBAdress = frmSplash.cdlg1.FileName
   
    End If
    
End Sub
Public Sub rightlable(maxscale As Double)
frmtest.picright.Cls
frmtest.picright.CurrentX = 0
frmtest.picright.CurrentY = 0
frmtest.picright.Print Format(Str(maxscale), "0.#")
frmtest.picright.CurrentX = 0
frmtest.picright.CurrentY = 1
frmtest.picright.Print Format(Str(maxscale * 3 / 4), "0.#")
frmtest.picright.CurrentX = 0
frmtest.picright.CurrentY = 2
frmtest.picright.Print Format(Str(maxscale * 2 / 4), "0.#")
frmtest.picright.CurrentX = 0
frmtest.picright.CurrentY = 2.99
frmtest.picright.Print Format(Str(maxscale * 1 / 4), "0.#")
frmtest.picright.CurrentX = 0
frmtest.picright.CurrentY = 3.98
frmtest.picright.Print Format(Str(maxscale * 0), "0.#")
frmtest.picright.CurrentX = 0
frmtest.picright.CurrentY = 4.98
frmtest.picright.Print Format(Str(-maxscale * 1 / 4), "0.#")
frmtest.picright.CurrentX = 0
frmtest.picright.CurrentY = 5.98
frmtest.picright.Print Format(Str(-maxscale * 2 / 4), "0.#")
frmtest.picright.CurrentX = 0
frmtest.picright.CurrentY = 6.97
frmtest.picright.Print Format(Str(-maxscale * 3 / 4), "0.#")
frmtest.picright.CurrentX = 0
frmtest.picright.CurrentY = 7.945
frmtest.picright.Print Format(Str(-maxscale), "0.#")

End Sub



