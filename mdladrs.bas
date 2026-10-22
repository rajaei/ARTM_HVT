Attribute VB_Name = "Module3"
Public acldb As String
Public iconadd As String
Private errnum As Long
Option Explicit
Sub Main()
Dim s As String
Dim ytest As Boolean
Dim y, rtv As Integer
acldb = "E:\Hvtest\Hvtestdb.mdb"
iconadd = App.Path + "\parsicon.ico"

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&7
With frmLogin.ado
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\Hvadres.mdb" + ";Persist Security Info=False"
    .CommandType = adCmdTable
    .RecordSource = "dbadres"
    .Refresh
End With

With frmLogin.ado.Recordset

    If .RecordCount = 0 Then
        .AddNew
        !mainadr = "E:\HVtest\Hvtest83.mdb"
        !acladr = "E:\HVtest\Hvtestdb.mdb"
        !Year = 83
        .Update
     End If
     .MoveFirst
     If .Fields("year") = 83 Then
        ytest = False
        Do While ytest = False
            s = InputBox("Please Enter Operating Year." + Chr(13) + "Year Number Most be > 83", "Enter Year Number", "83", Screen.Width / 3, Screen.Height / 3)
            y = Val(s)
            If y < 83 Or y > 99 Then
                MsgBox "Enter Integer Number From [83 - 99] Margin.", vbOKOnly, "Error"
            Else
                rtv = MsgBox("Do you sure that Operating Year is : 13" + s + " ?", vbYesNo, "Confirm Year Number")
                If rtv = 6 Then
                    ytest = True
                    !Year = y
                    .Update
                End If
            End If
        Loop
        
        !mainadr = "E:\HVtest\Hvtest" + s + ".mdb"
        !acladr = "E:\HVtest\Hvtestdb.mdb"
        .Update
    End If
    
    acldb = .Fields("acladr")
    .Close
    
End With

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

errnum = 0

Call opendb1

If errnum <> 0 Then
    Call opendb1
End If

If errnum = 0 Then
    With frmLogin.Data1
        .Connect = "Access"
        .DatabaseName = acldb
        .RecordSource = "personel"
        .Refresh
        .Recordset.Close
    End With
    errnum = Err.Number
End If

If errnum = 0 Then

    With frmLogin.Datalogin
        .Connect = "Access"
        .DatabaseName = acldb
        .RecordSource = "personel"
        .Refresh
    End With
    
    With frmLogin.ado.Recordset
        .Open
        .MoveFirst
        !acladr = acldb
        .Update
        .Close
    End With
    
    
    frmLogin.lblfn.Visible = False
    frmLogin.lblln.Visible = False
    frmLogin.Show
Else

    End
    
End If

End Sub
Public Sub opendb1()
Dim errdesc As String

On Error GoTo errhndl1


With frmLogin.Data1
    .Connect = "Access"
    .DatabaseName = acldb
    .RecordSource = "personel"
    .Refresh
    .Recordset.Close
End With
    
errhndl1:
    errnum = Err.Number
    errdesc = Err.Description
    
    If errnum <> 0 Then
        MsgBox "Error : " + Str(errnum) + Chr(13) + Chr(13) + errdesc + Chr(13) + Chr(13) + "Please Select Location Of : Hvtestdb.mdb  ", vbOKOnly, "Error"
        frmLogin.cdlg1.FileName = "Hvtestdb"
        frmLogin.cdlg1.Filter = "*.mdb"
        frmLogin.cdlg1.ShowOpen
        acldb = frmLogin.cdlg1.FileName
    End If

End Sub
Public Sub free()
Dim iRet As Integer
Dim ptAICheck As PT_FAICheck
Dim lStopped As Long
Dim iRetrieved As Long
Dim lOverrun As Long
Dim lHalfReady As Long
Dim lActiveBuf As Long
Dim i As Long
Dim ptFAIData As PT_FAITransfer
'//////////////////////////////////////////////////////////////////////
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
ptAICheck.stopped = DRV_GetAddress(lStopped)
ptAICheck.retrieved = DRV_GetAddress(iRetrieved)
ptAICheck.Overrun = DRV_GetAddress(lOverrun)
ptAICheck.HalfReady = DRV_GetAddress(lHalfReady)
ptAICheck.ActiveBuf = DRV_GetAddress(lActiveBuf)

'///////////     Check current status    //////////////////////

ErrCde = DRV_FAICheck(DeviceHandle, ptAICheck)
If (ChkErr(ErrCde)) Then
    FRMCAL.t0.Enabled = False
     MsgBox "Error In Reading Data From I/O card" + Chr(13) + "Please Check I/O card Setting And wirring.", vbExclamation, "Error"
    Exit Sub
End If
If (iRetrieved = 0) Then
    iRetrieved = 1
End If
       
'////////////////////////////////////////////////////////////////////
'//////////       transfer data from pci card ///////////////////
    
If (lStopped = 1) Then   ' if data ready for transfer from card
      
    ErrCde = DRV_FAIStop(DeviceHandle)  'Stop AI conversion
    If (ChkErr(ErrCde) = 1) Then
       DRV_DeviceClose (DeviceHandle)
       Exit Sub
    End If
    '////////////      Get the AI conversion data    ////////////
    
    ptFAIData.Overrun = DRV_GetAddress(lOverrun)
    ptFAIData.Count = ptAiStart.Count
    ptFAIData.start = 0                    'Start from buffer begging
    ptFAIData.DataType = 1 '  voltage value   //want floating data
    ptFAIData.ActiveBuf = 0                'Single buffer
    ReDim sDataBuffer(ptFAIData.Count)
    ptFAIData.DataBuffer = DRV_GetAddress(sDataBuffer(0))
    
    '***********   Read data from driver   ******************
    
    ErrCde = DRV_FAITransfer(DeviceHandle, ptFAIData)
    If (ChkErr(ErrCde) = 1) Then
        ReDim sDataBuffer(0)
        DRV_DeviceClose (DeviceHandle)
        Exit Sub
    End If
    '*********    Draw data in image box to show screen and detect peak value  ******
    DRV_DeviceClose (DeviceHandle)  'Close device
    
End If

End Sub

