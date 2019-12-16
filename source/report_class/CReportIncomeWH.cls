VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportIncomeWH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportIncomeWH"
Public MODULE_DESC As String

Private WithEvents Vsp As VSPrinter
Attribute Vsp.VB_VarHelpID = -1
Private mstrErrMsg As String
Private mcolParam As Collection
Private mblnNewPage As Boolean
Private mblnHeader As Boolean
Private mblnEndPage As Boolean
Private mdblWid   As Double
Private mdteDate As Date
Private mdY As Double
Private Rs As ADODB.Recordset

Private Const TITLE_SIZE = 14
Private Const HEADER_SIZE = 12
Private Const DETAIL_SIZE = 10

Private Const TITLE_FORMAT = "<32|>3|<5;"

Private m_TempCol As Collection
Private m_HeaderFormat1 As String
Private m_HeaderFormat2 As String
Private m_HeaderText1 As String
Private m_HeaderText2 As String

Private m_Headers0 As CFieldList
Private m_Headers0_1 As CFieldList
Private m_Headers1 As CFieldList
Private m_Details1 As CFieldList
Private m_Left As Double
Private m_PartTxtypes As Collection
Public ID As Long
Private DOCUMENT_NO As String
Private CUSTOMER_NAME As String
Private TRUCK_NO As String
Private DOCUMENT_DATE As String
Private CONDITION1 As Long
Private CONDITION2 As Long
Private CONDITION3 As Long
Private EMP_CHECK_CAR_ID As Long
Private EMP_CHECK_CAR_NAME As String
Private EMP_CHECK_CAR_LNAME As String
Private EMP_CHECK_PRODUCT_NAME As String
Private EMP_CHECK_PRODUCT_LNAME As String
Private EXIT_WEIGHT As Double
Private ENTRY_WEIGHT As Double
Private TOTAL_WEIGHT As Double
Private isNewPage As Boolean
Private Total1(100) As Double

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_TempCol = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers0_1 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Details1 = New CFieldList
   Set m_PartTxtypes = New Collection
   
   Set Rs = New ADODB.Recordset
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers0_1 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_PartTxtypes = Nothing
   
'   If Rs.State = adStateOpen Then
'      Rs.Close
'   End If
   Set Rs = Nothing
End Sub

Private Function CReportInterface_AddParam(varItem As Variant, strIndex As String) As Boolean
   Call mcolParam.add(varItem, strIndex)
   CReportInterface_AddParam = True
End Function

Private Sub CReportInterface_ClearParam()
   Call ClearParam
End Sub

Private Property Get CReportInterface_ErrorMsg() As String
   CReportInterface_ErrorMsg = mstrErrMsg
End Property

Private Function CReportInterface_Preview() As Boolean
   CReportInterface_Preview = genDoc(True)
End Function

Private Function CReportInterface_PrintDoc() As Boolean
   CReportInterface_PrintDoc = genDoc(False)
End Function

Private Property Set CReportInterface_VsPrint(RHS As VSPrinter7LibCtl.IVSPrinter)
   Set Vsp = RHS
End Property

Private Sub ClearParam()
   Dim I As Long
   
   For I = 1 To mcolParam.Count
      mcolParam.Remove 1
   Next I

End Sub
'============================= Codes above used in every report ==========================

Private Sub printHeader()
Dim strFormat As String
Dim strPrint As String
Dim tbdT As TableBorderSettings
Dim blnBold As Boolean
Dim iSize As Integer
Dim alngX() As Long
Dim Sgi As CSGroupItem
Dim TempSgi As CSGroupItem
Dim PrevID As Long
Dim TempStr1 As String
Dim TempStr2 As String
Dim Cf As CReportField
   
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbNone
   Vsp.TextAlign = taLeftBottom
   
   Vsp.MarginLeft = "1.7cm"
   Vsp.CurrentY = 1900
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 1
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcRowHeight, 1) = 1 * 567
   Vsp.TableCell(tcColWidth, 1, 1) = "3.5cm"
   Vsp.TableCell(tcAlign, 1, 1) = taLeftBottom
   Vsp.TableCell(tcText, 1, 1) = Year(mcolParam("PRINT_DATE")) + 543 & "/" & Format(Month(mcolParam("PRINT_DATE")), "00")
   Vsp.TableBorder = tbNone
   Vsp.EndTable
   
   Vsp.MarginLeft = "3cm"
   Vsp.CurrentY = 2400
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 1
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcRowHeight, 1) = 1 * 567
   Vsp.TableCell(tcColWidth, 1, 1) = "3.5cm"
   Vsp.TableCell(tcAlign, 1, 1) = taLeftBottom
   Vsp.TableCell(tcText, 1, 1) = mcolParam("HEAD_PACK_NAME")
   Vsp.TableBorder = tbNone
   Vsp.EndTable
   
      Vsp.MarginLeft = "25.5cm"
      Vsp.CurrentY = 2400
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 3
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight, 1) = 1 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "1.0cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "1.2cm"
      Vsp.TableCell(tcColWidth, 1, 3) = "1.3cm"
      Vsp.TableCell(tcAlign, 1, 1) = taCenterBottom
      Vsp.TableCell(tcAlign, 1, 2) = taCenterBottom
      Vsp.TableCell(tcAlign, 1, 3) = taCenterBottom
      Vsp.TableCell(tcText, 1, 1) = Format(Day(mcolParam("PRINT_DATE")), "00")
      Vsp.TableCell(tcText, 1, 2) = Format(Month(mcolParam("PRINT_DATE")), "00")
      Vsp.TableCell(tcText, 1, 3) = Year(mcolParam("PRINT_DATE")) + 543
      Vsp.TableBorder = tbNone
      Vsp.EndTable
            
      Vsp.TableBorder = tbdT
      Vsp.FontBold = blnBold
      Vsp.FontSize = iSize
End Sub

Private Function initDoc() As Boolean
Dim strDate As String
Dim TempStr1 As String
Dim TempStr2 As String
Dim TempStr3 As String
Dim FromDate As String
Dim ToDate As String
   
   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
   Vsp.ORIENTATION = orLandscape
   Vsp.MarginBottom = 1440
   Vsp.MarginFooter = 1440
   Vsp.MarginHeader = 1440
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 1440
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
'   Vsp.DrawPicture mcolParam("BACK_GROUND"), "0.0cm", "0.4cm", "29.7cm", "20.5cm", 10
   
'   FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
'   ToDate = "ถึง " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")
'
'   MODULE_DESC = vbCrLf & glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & vbCrLf & _
'                                       mcolParam("REPORT_NAME") & vbCrLf & _
'                                        FromDate & " " & ToDate & " " & mcolParam("LOCATION_NAME")
   MODULE_DESC = ""
   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))
   
   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
      mstrErrMsg = "Error GetDateTime Error."
      Exit Function
   End If
   mdteDate = InternalDateToDate(strDate)
   initDoc = True
End Function

Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim RName As String
Dim I As Long
Dim J As Long
Dim num As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Pi As CPartItem
Dim iCount As Long
Dim TempStr As String
Dim Amt As Double
Dim m_InventoryWHDoc As CInventoryWHDoc
Dim LotItemWh As CLotItemWH
Dim C_IWD As Collection

Dim LD As CLotDoc
Dim PD As CPalletDoc
Dim Count1 As Double
Dim TempY As Double
Dim TempX As Double
Dim CountLT As Long
Dim CountPD As Long
Dim Arr(10) As String
Arr(0) = "1.5cm"
Arr(1) = "1.4cm"
Arr(2) = "1.35cm"
Arr(3) = "1.4cm"
Arr(4) = "1.3cm"
Arr(5) = "1.4cm"
Arr(6) = "1.4cm"
Arr(7) = "1.3cm"
Arr(8) = "1.4cm"
Arr(9) = "1.3cm"
Dim ArrH(10) As Long
ArrH(0) = 4550
ArrH(1) = 5000
ArrH(2) = 5500
ArrH(3) = 6000
ArrH(4) = 6450
ArrH(5) = 6950
ArrH(6) = 7450
ArrH(7) = 7900
ArrH(8) = 8400
ArrH(9) = 8900
Dim TempCount As Long
Dim CountRow As Long

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   If Not initDoc Then Exit Function

   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------

   For J = 1 To UBound(Total1)
      Total1(J) = 0
   Next J
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
   
   Vsp.FontSize = 15
   I = 0
   num = 0
   Vsp.TextAlign = taLeftBottom
   Vsp.MarginLeft = 1090 'm_ReportConfig.DUMMY_OFFSET
   
   If mcolParam("NOT_SHOW") = 0 Then
    Set C_IWD = New Collection
    Set m_InventoryWHDoc = New CInventoryWHDoc
    m_InventoryWHDoc.INVENTORY_WH_DOC_ID = -1
    m_InventoryWHDoc.VERIFY_FLAG = "Y" 'จะดึงเฉพาะตัวที่ผ่านการตรวจสอบแล้วเท่านั้น
    m_InventoryWHDoc.PACK_DATE = mcolParam("PRINT_DATE")
    m_InventoryWHDoc.HEAD_PACK_NO = mcolParam("HEAD_PACK_NO")
    m_InventoryWHDoc.DOCUMENT_TYPE = mcolParam("DOCUMENT_TYPE")
    m_InventoryWHDoc.OrderBy = mcolParam("ORDER_BY")
    m_InventoryWHDoc.OrderType = mcolParam("ORDER_TYPE")
    m_InventoryWHDoc.QueryFlag = 1 'ถ้าเป็น 0 ยังไม่ต้องเอาลูก หลาน แต่ถ้าเป็น 1 เอา ลูกหลานด้วย
   
    If Not glbDaily.QueryInventoryWhDocForIG(m_InventoryWHDoc, C_IWD, Rs, iCount, IsOK, glbErrorLog) Then
       glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
       Call EnableForm(Me, True)
       Exit Function
    End If
    
'     Call printHeader
    
    CountRow = 0
    For Each m_InventoryWHDoc In C_IWD
         For Each LotItemWh In m_InventoryWHDoc.C_LotItemsWH
            I = I + 1
            J = 0
            num = 0
            CountLT = 0
            
            CountRow = CountRow + 1
            If CountRow > 10 Then
               CountRow = 1
              Call GenerateFooter
              
               For J = 1 To UBound(Total1)
                Total1(J) = 0
               Next J
   
              Vsp.NewPage
              Vsp.MarginLeft = 1090
            End If

            
            Vsp.CurrentY = ArrH(CountRow - 1)
            Vsp.FontSize = 12
            Vsp.StartTable
            Vsp.TableCell(tcCols) = 15
            Vsp.TableCell(tcRows) = 1
            Vsp.TableCell(tcRowHeight, 1) = 1 * 567
            Vsp.TableCell(tcColWidth, 1, 1) = "1.9cm"
            Vsp.TableCell(tcColWidth, 1, 2) = "1.4cm"
            Vsp.TableCell(tcColWidth, 1, 3) = "1.7cm"
            Vsp.TableCell(tcColWidth, 1, 4) = "1.7cm"
            Vsp.TableCell(tcColWidth, 1, 5) = "1.3cm"
            Vsp.TableCell(tcColWidth, 1, 6) = "1.7cm"
            Vsp.TableCell(tcColWidth, 1, 7) = "1.7cm"
            Vsp.TableCell(tcColWidth, 1, 8) = "1.5cm"
            Vsp.TableCell(tcColWidth, 1, 9) = "2.1cm"
            Vsp.TableCell(tcColWidth, 1, 10) = "1.8cm"
            Vsp.TableCell(tcColWidth, 1, 11) = "1.8cm"
            Vsp.TableCell(tcColWidth, 1, 12) = "1.5cm"
            Vsp.TableCell(tcColWidth, 1, 13) = "1.5cm"
            Vsp.TableCell(tcColWidth, 1, 14) = "2.1cm"
            Vsp.TableCell(tcColWidth, 1, 15) = "3.5cm"
            
            Vsp.TableCell(tcAlign, 1, 1) = taLeftMiddle
            Vsp.TableCell(tcAlign, 1, 2) = taCenterMiddle
            Vsp.TableCell(tcAlign, 1, 3) = taLeftMiddle
            Vsp.TableCell(tcAlign, 1, 4) = taLeftMiddle
            Vsp.TableCell(tcAlign, 1, 5) = taCenterMiddle
            Vsp.TableCell(tcAlign, 1, 6) = taCenterMiddle
            Vsp.TableCell(tcAlign, 1, 7) = taCenterMiddle
            Vsp.TableCell(tcAlign, 1, 8) = taCenterMiddle
            Vsp.TableCell(tcAlign, 1, 9) = taCenterMiddle
            Vsp.TableCell(tcAlign, 1, 10) = taCenterMiddle
            Vsp.TableCell(tcAlign, 1, 11) = taCenterMiddle
            Vsp.TableCell(tcAlign, 1, 12) = taCenterMiddle
            Vsp.TableCell(tcAlign, 1, 13) = taCenterMiddle
            Vsp.TableCell(tcAlign, 1, 14) = taCenterMiddle
            Vsp.TableCell(tcAlign, 1, 15) = taLeftMiddle
      
            
            Vsp.TableCell(tcText, 1, 1) = "" 'LotItemWh.PART_NO
            Vsp.TableCell(tcText, 1, 2) = LotItemWh.PRODUCT_TYPE_NAME
            Vsp.TableCell(tcText, 1, 3) = ""
            Vsp.TableCell(tcText, 1, 4) = DateToStringExtEx2(m_InventoryWHDoc.START_DATE)
            Vsp.TableCell(tcText, 1, 5) = m_InventoryWHDoc.BATCH_NO
            Total1(5) = Total1(5) + IIf(Val(m_InventoryWHDoc.BATCH_NO) > 0, m_InventoryWHDoc.BATCH_NO, 0)
            Vsp.TableCell(tcText, 1, 6) = FormatNumber(LotItemWh.GOOD_AMOUNT, 0)
            Total1(6) = Total1(6) + LotItemWh.GOOD_AMOUNT
            Vsp.TableCell(tcText, 1, 7) = FormatNumber(LotItemWh.LOSE_AMOUNT, 0)
            Total1(7) = Total1(7) + LotItemWh.LOSE_AMOUNT
            If Val(mcolParam("DOCUMENT_TYPE")) = 13 Then
               Vsp.TableCell(tcText, 1, 8) = "-"
               Vsp.TableCell(tcText, 1, 9) = "-"
            Else
               Vsp.TableCell(tcText, 1, 8) = LotItemWh.WEIGHT_PER_PACK
               Vsp.TableCell(tcText, 1, 9) = FormatNumber((LotItemWh.GOOD_AMOUNT + LotItemWh.LOSE_AMOUNT) * LotItemWh.WEIGHT_PER_PACK, 0)
               Total1(9) = Total1(9) + (LotItemWh.GOOD_AMOUNT + LotItemWh.LOSE_AMOUNT) * LotItemWh.WEIGHT_PER_PACK
            End If
            Vsp.TableCell(tcText, 1, 10) = FormatNumber(LotItemWh.REST_AMOUNT, 0)
             Total1(10) = Total1(10) + LotItemWh.REST_AMOUNT
            Vsp.TableCell(tcText, 1, 11) = LotItemWh.BIN_NAME
            If Val(mcolParam("DOCUMENT_TYPE")) = 13 Then
               Vsp.TableCell(tcText, 1, 12) = "-"
               Vsp.TableCell(tcText, 1, 13) = "-"
            Else
               Vsp.TableCell(tcText, 1, 12) = Format(LotItemWh.TIME_PACK_BEGIN, "HH:mm")
               Vsp.TableCell(tcText, 1, 13) = Format(LotItemWh.TIME_PACK_END, "HH:mm")
            End If
            Vsp.TableCell(tcText, 1, 14) = FormatNumber((LotItemWh.GOOD_AMOUNT + LotItemWh.LOSE_AMOUNT), 0)
            Total1(14) = Total1(14) + (LotItemWh.GOOD_AMOUNT + LotItemWh.LOSE_AMOUNT)
            Vsp.TableCell(tcText, 1, 15) = LotItemWh.NOTE
            Vsp.TableBorder = tbNone
            Vsp.EndTable
            
            Vsp.CurrentY = ArrH(CountRow - 1)
            Vsp.FontSize = 10
            Vsp.StartTable
            Vsp.TableCell(tcCols) = 3
            Vsp.TableCell(tcRows) = 1
            Vsp.TableCell(tcRowHeight, 1) = 1 * 567
            Vsp.TableCell(tcColWidth, 1, 1) = "1.9cm"
            Vsp.TableCell(tcColWidth, 1, 2) = "1.4cm"
            Vsp.TableCell(tcColWidth, 1, 3) = "1.8cm"
            Vsp.TableCell(tcAlign, 1, 1) = taLeftMiddle
            Vsp.TableCell(tcAlign, 1, 2) = taCenterMiddle
            Vsp.TableCell(tcAlign, 1, 3) = taCenterMiddle
            Vsp.TableCell(tcText, 1, 1) = LotItemWh.PART_NO
            Vsp.TableCell(tcText, 1, 2) = ""
            Vsp.TableCell(tcText, 1, 3) = LotItemWh.LOT_NO
            Vsp.TableBorder = tbNone
            Vsp.EndTable
         Next LotItemWh
   Next m_InventoryWHDoc
   Call GenerateFooter
   End If
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   Set m_InventoryWHDoc = Nothing
   genDoc = True
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
   Set Rs = Nothing
End Function

Private Sub GenerateFooter()
    Vsp.CurrentY = 9400
     Vsp.FontSize = 12
     Vsp.StartTable
      Vsp.TableCell(tcCols) = 15
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight, 1) = 1 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "1.9cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "1.5cm"
      Vsp.TableCell(tcColWidth, 1, 3) = "1.6cm"
      Vsp.TableCell(tcColWidth, 1, 4) = "1.5cm"
      Vsp.TableCell(tcColWidth, 1, 5) = "1.5cm"
      Vsp.TableCell(tcColWidth, 1, 6) = "1.7cm"
      Vsp.TableCell(tcColWidth, 1, 7) = "1.7cm"
      Vsp.TableCell(tcColWidth, 1, 8) = "1.5cm"
      Vsp.TableCell(tcColWidth, 1, 9) = "2.1cm"
      Vsp.TableCell(tcColWidth, 1, 10) = "1.8cm"
      Vsp.TableCell(tcColWidth, 1, 11) = "1.8cm"
      Vsp.TableCell(tcColWidth, 1, 12) = "1.5cm"
      Vsp.TableCell(tcColWidth, 1, 13) = "1.5cm"
      Vsp.TableCell(tcColWidth, 1, 14) = "2.1cm"
      Vsp.TableCell(tcColWidth, 1, 15) = "3.5cm"
      
      Vsp.TableCell(tcAlign, 1, 1) = taLeftMiddle
      Vsp.TableCell(tcAlign, 1, 2) = taCenterMiddle
      Vsp.TableCell(tcAlign, 1, 3) = taCenterMiddle
      Vsp.TableCell(tcAlign, 1, 4) = taLeftMiddle
      Vsp.TableCell(tcAlign, 1, 5) = taCenterMiddle
      Vsp.TableCell(tcAlign, 1, 6) = taCenterMiddle
      Vsp.TableCell(tcAlign, 1, 7) = taCenterMiddle
      Vsp.TableCell(tcAlign, 1, 8) = taCenterMiddle
      Vsp.TableCell(tcAlign, 1, 9) = taCenterMiddle
      Vsp.TableCell(tcAlign, 1, 10) = taCenterMiddle
      Vsp.TableCell(tcAlign, 1, 11) = taCenterMiddle
      Vsp.TableCell(tcAlign, 1, 12) = taCenterMiddle
      Vsp.TableCell(tcAlign, 1, 13) = taCenterMiddle
      Vsp.TableCell(tcAlign, 1, 14) = taCenterMiddle
      Vsp.TableCell(tcAlign, 1, 15) = taLeftMiddle
   
      Vsp.TableCell(tcText, 1, 5) = Total1(5) 'FormatNumber(Total1(5), 2)
      Vsp.TableCell(tcText, 1, 6) = FormatNumber(Total1(6), 0)
      Vsp.TableCell(tcText, 1, 7) = FormatNumber(Total1(7), 0)
      Vsp.TableCell(tcText, 1, 9) = FormatNumber(Total1(9), 0)
      Vsp.TableCell(tcText, 1, 10) = FormatNumber(Total1(10), 0)
      Vsp.TableCell(tcText, 1, 14) = FormatNumber(Total1(14), 0)

      Vsp.TableBorder = tbNone
      Vsp.EndTable
End Sub

Private Sub VSP_EndDoc()
'This event occur when VSPrinter.EndDoc is used
End Sub

Private Sub VSP_EndPage()
   If Not mblnEndPage Then Exit Sub
'   Vsp.DrawLine Vsp.MarginLeft, mdY, Vsp.MarginLeft, Vsp.CurrentY
'   Vsp.DrawLine Vsp.PageWidth - Vsp.MarginRight, mdY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY
'   Vsp.DrawLine Vsp.MarginLeft, Vsp.CurrentY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY
End Sub

Private Sub VSP_Error()
'Error in runtime occur here press F1 in VSP.ErrorDescription to see more information
   mstrErrMsg = Vsp.ErrorDescription
End Sub

Private Sub VSP_NewPage()
Dim talnT As TextAlignSettings
Dim tbdT As TableBorderSettings
Dim blnBold As Boolean
Dim blnUnder As Boolean
Dim blnItalic As Boolean
Dim iSize As Integer
Dim sName As String
Dim strFormat As String
Dim dY(0 To 1) As Double
Dim alngX() As Long

   If Not mblnNewPage Then Exit Sub
   isNewPage = True
   talnT = Vsp.TextAlign
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   blnUnder = Vsp.FontUnderline
   blnItalic = Vsp.FontItalic
   iSize = Vsp.FontSize
   sName = Vsp.FontName
  
  Vsp.DrawPicture mcolParam("BACK_GROUND"), "0.0cm", "0.3cm", "29.7cm", "20.5cm", 10
   Vsp.FontSize = 8
   Vsp.CurrentY = 400
   Vsp.MarginLeft = 2835
   dY(0) = Vsp.CurrentY
   Vsp.TextAlign = taLeftBottom
   strFormat = VSP_CalTable(TITLE_FORMAT, mdblWid, alngX)
   Vsp.TableBorder = tbNone
   Call Vsp.AddTable(strFormat, "", "|หน้าที่|: " & Vsp.CurrentPage)
'   Call Vsp.AddTable(strFormat, "", "|วันที่พิมพ์|: " & DateToStringExt(mdteDate))
'   Call Vsp.AddTable(strFormat, "", "|เวลา|: " & Right$(DateToStringExtEx(mdteDate), 8))
'   Call Vsp.AddTable(strFormat, "", "|ผู้ใช้|: " & glbUser.USER_NAME)
   Vsp.CurrentY = dY(0)

   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   Vsp.TextAlign = taCenterBottom
   Vsp.Paragraph = MODULE_DESC
   Vsp.Paragraph = ""
   If mblnHeader Then
      Call printHeader
   End If
   mdY = Vsp.CurrentY

   Vsp.TextAlign = talnT
   Vsp.TableBorder = tbdT
   Vsp.FontBold = blnBold
   Vsp.FontUnderline = blnUnder
   Vsp.FontItalic = blnItalic
   Vsp.FontSize = iSize
   Vsp.FontName = sName
End Sub

Private Sub VSP_StartDoc()
'This event occur when VSPrinter.StartDoc is used and used to initialize some information before generating printed document
End Sub


