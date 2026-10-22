Attribute VB_Name = "mdlBXLLIB"

'------------------------------------------------------------------------------------------------------
'  Constant Variables
'------------------------------------------------------------------------------------------------------

    Public Enum SLCS_HRI
        HRI_NOT_PRINT = 0
        HRI_BELOW_SIZE1 = 1
        HRI_ABOVE_SIZE1 = 2
        HRI_BELOW_SIZE2 = 3
        HRI_ABOVE_SIZE2 = 4
        HRI_BELOW_SIZE3 = 5
        HRI_ABOVE_SIZE3 = 6
        HRI_BELOW_SIZE4 = 7
        HRI_ABOVE_SIZE4 = 8
    End Enum

    Public Enum SLCS_PDF417_ECC_LEVEL
        PDF417_ECC_LEVEL0 = 0
        PDF417_ECC_LEVEL1 = 1
        PDF417_ECC_LEVEL2 = 2
        PDF417_ECC_LEVEL3 = 3
        PDF417_ECC_LEVEL4 = 4
        PDF417_ECC_LEVEL5 = 5
        PDF417_ECC_LEVEL6 = 6
        PDF417_ECC_LEVEL7 = 7
        PDF417_ECC_LEVEL8 = 8
    End Enum

    Public Enum SLCS_CIRCLE_SIZE
        CIRCLE_SIZE_1 = 1 ' 40 x 40 in dot unit
        CIRCLE_SIZE_2 = 2 ' 56 x 56 in dot unit
        CIRCLE_SIZE_3 = 3 ' 72 x 72 in dot unit
        CIRCLE_SIZE_4 = 4 ' 88 x 88 in dot unit
        CIRCLE_SIZE_5 = 5 ' 104 x 104 in dot unit
        CIRCLE_SIZE_6 = 6     ' 168 x 168 in dot unit
    End Enum

    Enum SLCS_PDF417_DATA_TYPE
        PDF417_TEXT_TYPE = 0
        PDF417_NUMERIC_TYPE = 1
        PDF417_BINARY_TYPE = 2
    End Enum
    
    Public Enum SLCS_INTERNATIONAL_CHARSET
        ICS_USA = 0
        ICS_FRANCE = 1
        ICS_GERMANY = 2
        ICS_UK = 3
        ICS_DENMARK_I = 4
        ICS_SWEDEN = 5
        ICS_ITALY = 6
        ICS_SPAIN_I = 7
        ICS_NORWAY = 8
        ICS_DENMARK_II = 9
        ICS_JAPAN = 10
        ICS_SPAIN_II = 11
        ICS_LATIN_AMERICA = 12
        ICS_KOREA = 13
        ICS_SLOVENIA_CROATIA = 14
        ICS_CHINA = 15
    End Enum

    Public Enum SLCS_CODEPAGE
        FCP_CP437 = 0
        FCP_CP850 = 1 ' OEM Multilingual Latin 1; Western European (DOS)
        FCP_CP852 = 2 ' OEM Latin 2; Central European (DOS)
        FCP_CP860 = 3 ' OEM Portuguese; Portuguese (DOS)
        FCP_CP863 = 4 ' OEM French Canadian; French Canadian (DOS)
        FCP_CP865 = 5 ' OEM Nordic; Nordic (DOS)
        FCP_CP1252 = 6 'ANSI Latin 1; Western European (Windows)
        FCP_CP857 = 8 ' OEM Turkish; Turkish (DOS)
        FCP_CP737 = 9 ' OEM Greek (formerly 437G); Greek (DOS)
        FCP_CP1250 = 10 ' ANSI Central European; Central European (Windows)// MPOS_CODEPAGE_CP1250 ??
        FCP_CP1253 = 11 ' ANSI Greek; Greek (Windows)
        FCP_CP1254 = 12 ' ANSI Turkish; Turkish (Windows)
        FCP_CP855 = 13 ' OEM Cyrillic (primarily Russian)
        FCP_CP862 = 14 ' OEM Hebrew; Hebrew (DOS)
        FCP_CP866 = 15 ' OEM Russian; Cyrillic (DOS)
        FCP_CP1251 = 16 ' ANSI Cyrillic; Cyrillic (Windows)
        FCP_CP1255 = 17 ' ANSI Hebrew; Hebrew (Windows)
        FCP_CP928 = 18 ' Greek
        FCP_CP775 = 20 ' OEM Baltic; Baltic (DOS)
        FCP_CP1257 = 21 ' ANSI Baltic; Baltic (Windows)
        FCP_CP858 = 22 '  OEM Multilingual Latin 1 + Euro symbol
    End Enum

    Public Enum SLCS_ROTATION
        ROTATE_0 = 0
        ROTATE_90 = 1
        ROTATE_180 = 2
        ROTATE_270 = 3
    End Enum

    '   Device Fonts
    Public Enum SLCS_DEVICE_FONT
        ENG_9X15 = 0
        ENG_12X20 = 1
        ENG_16X25 = 2
        ENG_19X30 = 3
        ENG_24X38 = 4
        ENG_32X50 = 5
        ENG_48X76 = 6
        ENG_22X34 = 7
        ENG_28X44 = 8
        ENG_37X58 = 9
        KOR_16X16 = &H61
        KOR_24X24 = &H62
        KOR_20X20 = &H63
        KOR_20X26 = &H65
        KOR_38X38 = &H66
        JPN_ShiftJIS = &H6A
        CHN_GB2312 = &H6D
        CHN_BIG5 = &H6E
    End Enum

    Public Enum SLCS_BARCODE
        CODE39 = 0
        CODE128 = 1
        I2OF5 = 2
        CODABAR = 3
        CODE93 = 4
        UPC_A = 5
        UPC_E = 6
        EAN13 = 7
        EAN8 = 8
        UCC_EAN128 = 9
    End Enum

    Public Enum SLCS_BLOCK_OPTION
        LINE_OVER_WRITING = 0
        LINE_EXCLUSIVE_OR = 1
        LINE_DELETE = 2
        SLOPE = 3
        BOX = 4
    End Enum

    Public Enum SLCS_DATAMATRIX_SIZE
        DATAMATRIX_SIZE_1 = 1
        DATAMATRIX_SIZE_2 = 2
        DATAMATRIX_SIZE_3 = 3
        DATAMATRIX_SIZE_4 = 4
    End Enum

    Public Enum SLCS_QRCODE_MODEL
        QRMODEL_1 = 1
        QRMODEL_2 = 2
    End Enum


    Public Enum SLCS_QRCODE_SIZE
        QRSIZE_1 = 1
        QRSIZE_2 = 2
        QRSIZE_3 = 3
        QRSIZE_4 = 4
        QRSIZE_5 = 5
        QRSIZE_6 = 6
        QRSIZE_7 = 7
        QRSIZE_8 = 8
        QRSIZE_9 = 9
    End Enum

    Public Enum SLCS_QRCODE_ECC_LEVEL
        QRECCLEVEL_L = 1 ' 7%
        QRECCLEVEL_M = 2 ' 15%
        QRECCLEVEL_Q = 3 ' 25%
        QRECCLEVEL_H = 4 ' 30%
    End Enum

    Public Enum SLCS_ORIENTATION
        TOP2BOTTOM = 0
        BOTTOM2TOP = 1
    End Enum


    Public Enum SLCS_MEDIA_TYPE
        GAP = 0
        CONTINUOUS = 1
        BLACKMARK = 2
    End Enum


    Public Enum SLCS_PRINT_SPEED
        PRINTER_SETTING_SPEED = -1
        PRINTER_SPEED_0 = 0
        PRINTER_SPEED_1 = 1
        PRINTER_SPEED_2 = 2
        PRINTER_SPEED_3 = 3
        PRINTER_SPEED_4 = 4
        PRINTER_SPEED_5 = 5
        PRINTER_SPEED_6 = 6
        PRINTER_SPEED_7 = 7
        PRINTER_SPEED_8 = 8
        PRINTER_SPEED_9 = 9
        PRINTER_SPEED_10 = 10
        PRINTER_SPEED_11 = 11
        PRINTER_SPEED_12 = 12
    End Enum

    Public Enum SLCS_DITHER_OPTION
        DITHER_NONE = -1
        DITHER_1 = 0
        DITHER_2 = 1
        DITHER_3 = 6
        DITHER_4 = 7
    End Enum

    ' Font Selection
    Public Const ASCII As String = "U"
    Public Const KS5601 As String = "K"
    Public Const BIG5 As String = "B"
    Public Const GB2312 As String = "G"
    Public Const ShiftJIS As String = "J"


    ' Font Alignment
    Public Const LEFTALIGN As String = "L"
    Public Const RIGHTALIGN As String = "R"
    Public Const CENTERALIGN As String = "C"

    ' Font Direction
    Public Const LEFTTORIGHT As Integer = 0
    Public Const RIGHTTOLEFT As Integer = 1

'------------------------------------------------------------------------------------------------------
' DLL API Function
'------------------------------------------------------------------------------------------------------

Public Declare Function ConnectPrinter Lib "BXLLIB.dll" (ByVal szPrinterName As String) As Boolean

Public Declare Function DisconnectPrinter Lib "BXLLIB.dll" () As Boolean

Public Declare Function StartLabel Lib "BXLLIB.dll" () As Boolean

Public Declare Sub EndLabel Lib "BXLLIB.dll" ()

Public Declare Function GetBIXOLON_PrinterList Lib "BXLLIB.dll" (ByVal strInstalledPrinter As Any) As Long

Public Declare Function GetDllVersion Lib "BXLLIB.dll" (ByVal strBxlPrtList As Any) As Boolean


Public Declare Function GetPrinterResolution Lib "BXLLIB.dll" () As Long


Public Declare Function Print1DBarcode Lib "BXLLIB.dll" (ByVal nHorizontalPos As Long, _
                                             ByVal nVerticalPos As Long, _
                                             ByVal nBarcodeType As Long, _
                                             ByVal nNarrowBarWidth As Long, _
                                             ByVal nWideBarWidth As Long, _
                                             ByVal nBarcodeHeight As Long, _
                                             ByVal nRotation As Long, _
                                             ByVal nHRI As Long, _
                                             ByVal pData As String) _
                                             As Boolean


Public Declare Function PrintDeviceFont Lib "BXLLIB.dll" (ByVal nHorizontalPos As Long, _
                                              ByVal nVerticalPos As Long, _
                                              ByVal nFontName As Long, _
                                              ByVal nHorizontalMulti As Long, _
                                              ByVal nVerticalMulti As Long, _
                                              ByVal nRotation As Long, _
                                              ByVal bBold As Boolean, _
                                              ByVal szText As String) _
                                              As Boolean


Public Declare Function SetConfigOfPrinter Lib "BXLLIB.dll" (ByVal nSpeed As Long, _
                                                 ByVal nDensity As Long, _
                                                 ByVal nOrientation As Long, _
                                                 ByVal bAutoCut As Long, _
                                                 ByVal nCuttingPeriod As Long, _
                                                 ByVal bBackFeeding As Boolean) _
                                                 As Boolean


Public Declare Function Prints Lib "BXLLIB.dll" (ByVal nLabelSet As Long, _
                                     ByVal nCopiesOfEachLabel As Long) _
                                     As Boolean


Public Declare Function SetCharacterset Lib "BXLLIB.dll" (ByVal nInternationalCharacterSet As Long, _
                                       ByVal nCodepage As Long) _
                                       As Boolean
                                       
Public Declare Function SetPaper Lib "BXLLIB.dll" (ByVal nHorizontalMagin As Long, _
                                       ByVal nVerticalMargin As Long, _
                                       ByVal nPaperWidth As Long, _
                                       ByVal nPaperLength As Long, _
                                       ByVal nMediaType As Long, _
                                       ByVal nOffSet As Long, _
                                       ByVal nGapLengthORThicknessOfBlackLine As Long) _
                                       As Boolean


Public Declare Function ClearBuffer Lib "BXLLIB.dll" () As Boolean


Public Declare Function PrintBlock Lib "BXLLIB.dll" (ByVal nHorizontalStartPos As Long, _
                                         ByVal nVerticalStartPos As Long, _
                                         ByVal nHorizontalEndPos As Long, _
                                         ByVal nVerticalEndPos As Long, _
                                         ByVal nOption As Long, _
                                         ByVal nThickness As Long) _
                                         As Boolean

Public Declare Function PrintCircle Lib "BXLLIB.dll" (ByVal nHorizontalStartPos As Long, _
                                         ByVal nVerticalStartPos As Long, _
                                         ByVal nDiameter As Long, _
                                         ByVal nMulti As Long) _
                                         As Boolean


Public Declare Function PrintDirect Lib "BXLLIB.dll" (ByVal pDirectData As String, ByVal bAddCrLf As Boolean) As Boolean

Public Declare Function WriteBuffer Lib "BXLLIB.dll" (ByVal pDirectData As String, ByVal nLength As Long) As Boolean


Public Declare Function PrintTrueFont Lib "BXLLIB.dll" (ByVal nXPos As Long, _
                                ByVal nYPos As Long, _
                                ByVal strFontName As String, _
                                ByVal nFontSize As Long, _
                                ByVal nRotaion As Long, _
                                ByVal bItalic As Boolean, _
                                ByVal bBold As Boolean, _
                                ByVal bUnderline As Boolean, _
                                ByVal strText As String, _
                                ByVal bDataCompression As Boolean) _
                                As Boolean
                                
Public Declare Function PrintTrueFontLib Lib "BXLLIB.dll" (ByVal nXPos As Long, _
                                ByVal nYPos As Long, _
                                ByVal strFontName As String, _
                                ByVal nFontSize As Long, _
                                ByVal nRotaion As Long, _
                                ByVal bItalic As Boolean, _
                                ByVal bBold As Boolean, _
                                ByVal bUnderline As Boolean, _
                                ByVal strText As String, _
                                ByVal bDataCompression As Boolean) _
                                As Boolean

Public Declare Function PrintTrueFontLibWithAlign Lib "BXLLIB.dll" (ByVal nXPos As Long, _
                                ByVal nYPos As Long, _
                                ByVal strFontName As String, _
                                ByVal nFontSize As Long, _
                                ByVal nRotaion As Long, _
                                ByVal bItalic As Boolean, _
                                ByVal bBold As Boolean, _
                                ByVal bUnderline As Boolean, _
                                ByVal strText As String, _
                                ByVal nPrintWidth As String, _
                                ByVal nAlignment As Long, _
                                ByVal bDataCompression As Boolean) _
                                As Boolean


Public Declare Function PrintVectorFont Lib "BXLLIB.dll" ( _
                                        ByVal nXPos As Long, _
                                        ByVal nYPos As Long, _
                                        ByVal strFontSelection As String, _
                                        ByVal nFontWidth As Long, _
                                        ByVal nFontHeight As Long, _
                                        ByVal strRightSideCharSpacing As String, _
                                        ByVal bBold As Boolean, _
                                        ByVal bReversePrinting As Boolean, _
                                        ByVal bTextStyle As Boolean, _
                                        ByVal nRotation As Long, _
                                        ByVal strTextAlignment As String, _
                                        ByVal nTextDirection As Long, _
                                        ByVal pData As String) _
                                        As Boolean


Public Declare Function PrintPDF417 Lib "BXLLIB.dll" ( _
                                ByVal nXPos As Integer, _
                                ByVal nYPos As Integer, _
                                ByVal nMaxRow As Integer, _
                                ByVal nMaxCol As Integer, _
                                ByVal nECLevel As Integer, _
                                ByVal nDataType As Integer, _
                                ByVal bHRI As Boolean, _
                                ByVal nOriginPoint As Integer, _
                                ByVal nModuleWidth As Integer, _
                                ByVal nBarHeight As Integer, _
                                ByVal nRotation As Integer, _
                                ByVal pData As String) _
                                As Boolean

Public Declare Function PrintMaxiCode Lib "BXLLIB.dll" ( _
                                    ByVal nXPos As Integer, _
                                    ByVal nYPos As Integer, _
                                    ByVal nMode As Integer, _
                                    ByVal pData As String) _
                                    As Boolean

Public Declare Function PrintDataMatrix Lib "BXLLIB.dll" ( _
                                    ByVal nXPos As Integer, _
                                    ByVal nYPos As Integer, _
                                    ByVal nSize As Integer, _
                                    ByVal bReverse As Boolean, _
                                    ByVal nRotation As Integer, _
                                    ByVal pData As String) _
                                    As Boolean
                                    
Public Declare Function PrintQRCode Lib "BXLLIB.dll" (ByVal nXPos As Long, _
                                                      ByVal nYPos As Long, _
                                                      ByVal nModel As Long, _
                                                      ByVal nECCLevel As Long, _
                                                      ByVal nSize As Long, _
                                                      ByVal nRotation As Long, _
                                                      ByVal pData As String) _
                                                      As Boolean

Public Declare Function PrintImageLib Lib "BXLLIB.dll" (ByVal nHorizontalStartPos As Long, _
                                                        ByVal nVerticalStartPos As Long, _
                                                        ByVal pBitmapFilename As String, _
                                                        ByVal nDither As Long, _
                                                        ByVal bDataCompression As Boolean) _
                                                        As Boolean
                                                        
Public Declare Function PrintImageLibWithSize Lib "BXLLIB.dll" (ByVal nHorizontalStartPos As Long, _
                                                        ByVal nVerticalStartPos As Long, _
                                                        ByVal nNewWidth As Long, _
                                                        ByVal nNewHeight As Long, _
                                                        ByVal pBitmapFilename As String, _
                                                        ByVal nDither As Long, _
                                                        ByVal bDataCompression As Boolean) _
                                                        As Boolean

Public Declare Function RFIDSetup Lib "BXLLIB.dll" ( _
                                    ByVal RFIDType As Long, _
                                    ByVal NumberOfRetries As Long, _
                                    ByVal NumberOfLabel As Long, _
                                    ByVal RadioPower As Long) _
                                    As Boolean
                                    
Public Declare Function RFIDWrite Lib "BXLLIB.dll" ( _
                                    ByVal DataType As Long, _
                                    ByVal StartingBlockNumber As Long, _
                                    ByVal WriteByte As Long, _
                                    ByVal Data As String) _
                                    As Boolean
                                    
Public Declare Function RFIDPassword Lib "BXLLIB.dll" ( _
                                    ByVal OldAccessPwd As String, _
                                    ByVal OldKillPwd As String, _
                                    ByVal NewAccessPwd As String, _
                                    ByVal NewKillPwd As String) _
                                    As Boolean
                                    
Public Declare Sub RFIDLock Lib "BXLLIB.dll" ()
Public Declare Sub RFIDCalibration Lib "BXLLIB.dll" ()

Public Declare Function SetShowMsgBox Lib "BXLLIB.dll" (ByVal bShowMsgBox As Boolean) As Boolean
Public Declare Sub DEBUGLOG_ON Lib "BXLLIB.dll" ()
Public Declare Sub DEBUGLOG_OFF Lib "BXLLIB.dll" ()
