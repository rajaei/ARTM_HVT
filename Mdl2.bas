Attribute VB_Name = "mdl1"
Option Explicit
Public uref, ucoun, urt, iref, ic, ilaccp, inom, pduc, pdur, pdaccp, uresmax, uresmin, urefmin, urefmax, tem, optcode, blockid, ordernum As Double
Public arrestertype, selectedarr, optname, blockinf, blockname, access As String
Public maindb, acldb, icon, pic1, pic2 As String
Public year, blcnum, sncode, svr1, svr2 As Integer
Public trs1, trs2, trs3, trs4 As Integer
Public sealcon() As Integer
Public srn() As Double
Public newsrn() As Double
Public newarrnum As Double
Public tested() As Boolean
Public arrnum As Double
Public testrun, testprog As Boolean
Public blocks() As String
Public preform As String
Public i, rtv, j, ini As Integer
Dim errnum As Long
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
Dim s, spd As String
Dim ytest As Boolean
Dim y, pd As Integer
'year = number of year For identify  of  main database database "HVtest##"
'maindb=addres of main database "HVtest##"
'acldb=addres of accelary database "hvtestdb" consist of (block & personel  & testprmt)  Tables
'pic# = addres of picture for img1 in maimform

mdl1.testprog = False
mdl1.pic1 = App.Path + "\ETP1.bmp"
mdl1.pic2 = App.Path + "\ETP2.bmp" '  For Print in Print Form
mdl1.icon = App.Path + "\ETPicon.ico"

pdaccp = 2
'***************************************************************
With frmSplash.ado
    'ConnectionString : determine Provider and Address of data Base file
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\Hvadres.mdb" + ";Persist Security Info=False"
    .CommandType = adCmdTable
    .RecordSource = "dbadres" 'RecordSource : Table Name of Data base file
    .Refresh
End With

With frmSplash.ado.Recordset ' determine commands to set the fields of table mentioned in recordsourse

    If .RecordCount = 0 Then
        .addnew
        !mainadr = "E:\HVtest\Hvtest93.mdb"
        !acladr = "E:\HVtest\Hvtestdb.mdb"
        !year = 93
        .Update
     End If
     .MoveFirst
     If .Fields("year") < 93 Then
        ytest = False
        Do While ytest = False
            s = InputBox("Please Enter Current Year." + Chr(13) + "Year Most be > 1393", "Enter Year:", "1393", Screen.Width / 3, Screen.Height / 3)
            y = Val(s)
            If y < 1393 Or y > 1500 Then
                MsgBox "Enter Integer Number From 1393 to 1500.", vbOKOnly, "Error"
            Else
                rtv = MsgBox("Are you sure that operating year is : " + s + " ?", vbYesNo, "Confirm Year")
                If rtv = 6 Then
                    ytest = True
                    !year = y
                    .Update
                End If
            End If
        Loop
        
        ytest = False
        Do While ytest = False
            spd = InputBox("Acceptable PD is 2 [Pc]." + Chr(13) + "If you Want to change Acceptance Criteria, enter new value." + Chr(13) + "Max Acceptance PD is :", "Acceptance PD", "2", Screen.Width / 3, Screen.Height / 3)
            pd = Val(spd)
            If pd < 1 Or pd > 10 Then
                MsgBox "Enter integer number from the range of [1 - 10].", vbOKOnly, "Error"
            Else
                rtv = MsgBox("Are you sure that Max Acceptance PD is : " + spd + "  [Pc] ?", vbYesNo, "Confirm PD Value")
                If rtv = 6 Then
                    ytest = True
                    !accept_pd = pd
                    .Update
                End If
            End If
        Loop
        !mainadr = "E:\HVtest\Hvtest" + s + ".mdb"
        !acladr = "E:\HVtest\Hvtestdb.mdb"
        .Update
    End If
    
    s = Right(Str(.Fields("year")), 2)
    mdl1.maindb = .Fields("mainadr")
    mdl1.acldb = .Fields("acladr")
    mdl1.year = Right(Str(.Fields("year")), 2)
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
        'ConnectionString : determine Provider and Address of data Base file
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.acldb + ";Persist Security Info=False"
        .CommandType = adCmdTable
        .RecordSource = "personel"   'RecordSource : Table Name of Data base file
        .Refresh
        .Recordset.Close
    End With

    
    errnum = Err.Number
End If

If errnum = 0 Then

    Call opendb2
    
    If errnum <> 0 Then
         Call opendb2
    End If
    
    If errnum = 0 Then
        With frmSplash.Ado2
        'ConnectionString : determine Provider and Address of data Base file
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.maindb + ";Persist Security Info=False"
        .CommandType = adCmdTable
        .RecordSource = "DIST"   'RecordSource : Table Name of Data base file
        .Refresh
        .Recordset.Close
    End With

        errnum = Err.Number
     End If
    
    
End If

If errnum = 0 Then
    
    
   With frmSplash.ado.Recordset
        .Open
        .MoveFirst
        !mainadr = mdl1.maindb
        !acladr = mdl1.acldb
        .Update
        .Close
    End With

    frmSplash.Timer1.Enabled = True
    frmSplash.Show

Else
    
    End
    
End If

End Sub

Public Sub fdim(num As Integer)
ReDim srn(1 To num) As Double
ReDim sealcon(1 To num) As Integer
ReDim tested(1 To num) As Boolean
arrnum = num
testrun = True
End Sub


Public Sub opendb1()
Dim errdesc As String

On Error GoTo errhndl1

With frmSplash.Ado1
        'ConnectionString : determine Provider and Address of data Base file
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.acldb + ";Persist Security Info=False"
    .CommandType = adCmdTable
    .RecordSource = "personel"   'RecordSource : Table Name of Data base file
    .Refresh
    .Recordset.Close
End With



errhndl1:
    errnum = Err.Number
    errdesc = Err.Description
    
    If errnum <> 0 Then
        MsgBox "Error : " + Str(errnum) + Chr(13) + Chr(13) + errdesc + Chr(13) + Chr(13) + "Please Select Location Of : Hvtestdb.mdb  ", vbOKOnly, "Error"
        frmSplash.cdlg1.FileName = "Hvtestdb"
        frmSplash.cdlg1.Filter = "*.mdb"
        frmSplash.cdlg1.ShowOpen
        mdl1.acldb = frmSplash.cdlg1.FileName
    End If

End Sub
Public Sub opendb2()
Dim s, errdesc As String

s = Right(Str(mdl1.year), 2)


On Error GoTo errhndl1


With frmSplash.Ado2
        'ConnectionString : determine Provider and Address of data Base file
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdl1.maindb + ";Persist Security Info=False"
    .CommandType = adCmdTable
    .RecordSource = "DIST"   'RecordSource : Table Name of Data base file
    .Refresh
    .Recordset.Close
End With


errhndl1:
    errnum = Err.Number
    errdesc = Err.Description
    If errnum <> 0 Then
        MsgBox "Error : " + Str(errnum) + Chr(13) + Chr(13) + errdesc + Chr(13) + Chr(13) + "Please Select Location Of : Hvtest" + Str(mdl1.year) + ".mdb  ", vbOKOnly, "Error"
        frmSplash.cdlg1.FileName = "Hvtest" + s + ".mdb"
        frmSplash.cdlg1.Filter = "*.mdb"
        frmSplash.cdlg1.ShowOpen
        mdl1.maindb = frmSplash.cdlg1.FileName
   
    End If
    
End Sub







