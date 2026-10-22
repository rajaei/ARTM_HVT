Attribute VB_Name = "Module2"
Option Explicit
'Const MaxEntries = 255
'Global DeviceHandle As Long
'Global ptDevGetFeatures As PT_DeviceGetFeatures
'Global lpDevFeatures As DEVFEATURES
'Global devicelist(0 To MaxEntries) As PT_DEVLIST
'Global SubDevicelist(0 To MaxEntries) As PT_DEVLIST
'Global ErrCde As Long
'Global szErrMsg As String * 80
'Global lpAIConfig As PT_AIConfig
'Global AiVolIn As PT_AIVoltageIn
'Global ptAIGetConfig As PT_AIGetConfig
'Global AiCtrMode As Integer
'Global wGainCode(0 To 15) As Integer
'Global iGainIndex(0 To 15) As Integer
'Global dfCurDevice As DEVCONFIG_AI
'Global lpDEVCONFIG_AI As DEVCONFIG_AI
'Global lpAIGetConfig As PT_AIGetConfig
'Global ptEnableEvent As PT_EnableEvent          ' Enable event
'Global gnNumOfSubdevices As Integer
'Global gwEvtFlag As Integer                      'event enable(0)
'Global gwFifoSize As Integer           'Fifo Size
'Global gwFifoEnable As Integer                   ' Fifo
'Global Response As Integer
'Global lBuffer() As Long
'Global ptAiStart As PT_FAIIntScanStart
'Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
'Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
'Public Const GMEM_FIXED = &H0
'Public Const GMEM_ZEROINIT = &H40
'Public Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Public Function ChkErr(ByVal lErrCde As Long) As Long


    If (lErrCde <> 0) Then
        'DRV_GetErrorMessage lErrCde, szErrMsg
        'Response = MsgBox(szErrMsg, vbOKOnly, "Error!!")
        ChkErr = 1
    Else
        'ChkErr = 0
    End If
End Function


