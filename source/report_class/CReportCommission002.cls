VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportCommission002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportSell001_0"
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
Private Const DETAIL_SIZE = 12

Private Const TITLE_FORMAT = "<32|<3|<5;"

Private m_TempCol As Collection
Private m_HeaderFormat1 As String
Private m_HeaderFormat2 As String
Private m_HeaderText1 As String
Private m_HeaderText2 As String

Private m_Headers0 As CFieldList
Private m_Headers1 As CFieldList
Private m_Headers1_1 As CFieldList
Private m_Headers1_2 As CFieldList
Private m_Headers1_3 As CFieldList
Private m_Headers1_4 As CFieldList
Private m_Details1 As CFieldList
Private m_Details2 As CFieldList

Private collTargets As Collection
Private CommissionChartColl As Collection
Private CommissionChartEmpColl As Collection

Private CnBySale As Collection
Private DnBySale As Collection

Private CommissionSaleColl As Collection
Private CommissionCostColl As Collection

Private SumTotal As Collection
Private CollSaleYYYYMM As Collection
Private collEmployee As Collection

Private CurrentHeader As Long

Private collCommissionSubtract As Collection
Private collCommissionCredit As Collection

Private m_PaidAmounts As Collection
Private m_BillTotalPrices As Collection
Private m_DnItemsByBill As Collection
Private m_CnItemsByBill As Collection
Private m_RtItemsByBill As Collection
Private m_BillingDiscounts As Collection

Private tempCollPartItem As Collection
Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_TempCol = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Headers1_1 = New CFieldList
   Set m_Headers1_2 = New CFieldList
   Set m_Headers1_3 = New CFieldList
   Set m_Headers1_4 = New CFieldList
   Set m_Details1 = New CFieldList
   Set m_Details2 = New CFieldList
   
   Set collTargets = New Collection
   Set CommissionChartColl = New Collection
   Set CommissionChartEmpColl = New Collection
   
   Set CnBySale = New Collection
   Set DnBySale = New Collection
   Set CommissionSaleColl = New Collection
   Set CommissionCostColl = New Collection
   
   Set SumTotal = New Collection
   Set CollSaleYYYYMM = New Collection
   Set collEmployee = New Collection
   
   Set collCommissionSubtract = New Collection
   Set collCommissionCredit = New Collection
   
   Set m_PaidAmounts = New Collection
   Set m_BillTotalPrices = New Collection
   Set m_DnItemsByBill = New Collection
   Set m_CnItemsByBill = New Collection
   Set m_RtItemsByBill = New Collection
   Set m_BillingDiscounts = New Collection
   Set tempCollPartItem = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
   
   Set collTargets = Nothing
   Set CommissionChartColl = Nothing
   Set CommissionChartEmpColl = Nothing
   
   Set CnBySale = Nothing
   Set DnBySale = Nothing
   Set CommissionSaleColl = Nothing
   Set CommissionCostColl = Nothing
   
   Set SumTotal = Nothing
   Set CollSaleYYYYMM = Nothing
   Set collEmployee = Nothing
   
   Set collCommissionSubtract = Nothing
   Set collCommissionCredit = Nothing
   
   Set m_PaidAmounts = Nothing
   Set m_BillTotalPrices = Nothing
   Set m_DnItemsByBill = Nothing
   Set m_CnItemsByBill = Nothing
   Set m_RtItemsByBill = Nothing
   Set m_BillingDiscounts = Nothing
   Set tempCollPartItem = Nothing
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
Dim TempStr1_1 As String
Dim Cf As CReportField
Dim ExportWidth As Long
Dim HeadCf As CReportField
Dim J As Long
Dim TempStr As String
Dim Inv As CInventoryDoc
Dim TempRs As ADODB.Recordset
Dim ImportItems As Collection
Dim Amt As Double
Dim iCount As Long

   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField

   Call Cf.SetFieldValue(1, "^", "No", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "เดือน-ปี", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "จำนวนถุง", "^", "SUM_FLAG2")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "น้ำหนัก", ">", "SUM_FLAG2")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "สุทธิ", ">", "SUM_FLAG1")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ราคา/น.น.", ">", "AVG_FLAG")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ลดหนี้", ">", "SUM_FLAG1")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "เพิ่มหนี้", ">", "SUM_FLAG1")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "สุทธิ", ">", "SUM_FLAG1")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "เป้า", ">", "SUM_FLAG2")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "%", ">", "SUM_FLAG1")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "คอม รับ/หัก", ">", "SUM_FLAG1")
   Call m_Headers1.AddField(Cf)
   
   Set Cf = Nothing
            
  Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
            
   Set TempRs = Nothing
   
   Vsp.TableBorder = tbdT
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
End Sub
Private Sub printHeader2()
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
Dim TempStr1_1 As String
Dim Cf As CReportField
Dim ExportWidth As Long
Dim HeadCf As CReportField
Dim J As Long
Dim TempStr As String
Dim Inv As CInventoryDoc
Dim TempRs As ADODB.Recordset
Dim ImportItems As Collection
Dim Amt As Double
Dim iCount As Long

   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField

   Call Cf.SetFieldValue(4, "^", "MANAGER", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "ยอดขาย", ">", "SUM_FLAG1")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "น้ำหนัก", ">", "SUM_FLAG1")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "เป้า", ">", "SUM_FLAG2")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "%", ">", "SUM_FLAG1")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "คอม รับ/หัก", ">", "SUM_FLAG1")
   Call m_Headers1.AddField(Cf)
   
   Set Cf = Nothing
            
  Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
            
   Set TempRs = Nothing
   
   Vsp.TableBorder = tbdT
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
End Sub
Private Sub printHeader3()
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
Dim TempStr1_1 As String
Dim Cf As CReportField
Dim ExportWidth As Long
Dim HeadCf As CReportField
Dim J As Long
Dim TempStr As String
Dim Inv As CInventoryDoc
Dim TempRs As ADODB.Recordset
Dim ImportItems As Collection
Dim Amt As Double
Dim iCount As Long

   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField
   
   Call Cf.SetFieldValue(45, "^", "ใบเสร็จ", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(40, "^", "INVOICE", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(35, "^", "", "^")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(25, "^", "เลขที่", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(20, "^", "วันที่", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(20, "^", "เลขที่", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(20, "^", "วันที่", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(35, "^", "ลูกค้า", "<")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(20, "^", "%", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(20, "^", "รับ-ขาย-Com", ">", "KEY_PAID")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(155, "^", "รายละเอียด", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(25, "^", "เบอร์", ">")
   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(15, "^", "จำนวนถุง", ">", "SUM_FLAG2")
'   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(15, "^", "กก", ">", "SUM_FLAG2")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(20, "^", "มูลค่า", ">", "SUM_FLAG1")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(15, "^", "ราคาขาย", ">", "AVG_FLAG1")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(15, "^", "ต้นทุน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(15, "^", "ผลต่าง", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(20, "^", "GP", ">", "SUM_FLAG1-GP")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(15, "^", "ดบ.ล่าช้า", ">", "SUM_FLAG1-INTEREST")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(15, "^", "CMS", ">", "SUM_FLAG1-COM")
   Call m_Headers1.AddField(Cf)
   
   Set Cf = Nothing
            
   Call m_Headers0.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
  Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
            
   Set TempRs = Nothing
   
   Vsp.TableBorder = tbdT
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
End Sub
Private Sub printHeader4()
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
Dim TempStr1_1 As String
Dim Cf As CReportField
Dim ExportWidth As Long
Dim HeadCf As CReportField
Dim J As Long
Dim TempStr As String
Dim Inv As CInventoryDoc
Dim TempRs As ADODB.Recordset
Dim ImportItems As Collection
Dim Amt As Double
Dim iCount As Long

   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField

   Call Cf.SetFieldValue(4, "^", "พนักงานขาย", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ยอดขายสุทธิ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "เป้า", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "%", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "คอม รับ/หัก ทันที", ">", "SUM_FLAG1")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ยอดรับเงิน", ">", "SUM_FLAG1")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "คอม รับชำระ", ">", "SUM_FLAG1")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ดอกเบี้ยล่าช้า", ">", "SUM_FLAG1")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "คอม สุทธิ", ">", "SUM_FLAG1")
   Call m_Headers1.AddField(Cf)
   
   
   Set Cf = Nothing
            
  Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
            
   Set TempRs = Nothing
   
   Vsp.TableBorder = tbdT
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
End Sub
Private Sub printHeader5()
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
Dim TempStr1_1 As String
Dim Cf As CReportField
Dim ExportWidth As Long
Dim HeadCf As CReportField
Dim J As Long
Dim TempStr As String
Dim Inv As CInventoryDoc
Dim TempRs As ADODB.Recordset
Dim ImportItems As Collection
Dim Amt As Double
Dim iCount As Long

   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField
   
   Call Cf.SetFieldValue(20, "^", "เลขที่", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(20, "^", "วันที่", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(50, "^", "ลูกค้า", "<")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(20, "^", "เครดิต+3", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(20, "^", "ครบกำหนด", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(20, "^", "เกินกำหนด(วัน)", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(20, "^", "ยอดบิล", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(20, "^", "ยอดคงเหลือ", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(20, "^", "เบี้ยปรับ", ">", "SUM_FLAG1")
   Call m_Headers1.AddField(Cf)
   
   Set Cf = Nothing
            
  Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
            
   Set TempRs = Nothing
   
   Vsp.TableBorder = tbdT
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
End Sub

Private Function initDoc() As Boolean
Dim strDate As String
Dim TempStr1 As String
Dim TempStr2 As String
Dim TempStr3 As String
Dim FromDate As Date
Dim ToDate As Date
   
   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
   Vsp.ORIENTATION = orLandscape
   Vsp.MarginBottom = 300
   Vsp.MarginFooter = 300
   Vsp.MarginHeader = 300
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 300
   Vsp.MarginTop = 300
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   Call GetFirstLastDate(DateSerial(mcolParam("YEAR_NO") - 543, mcolParam("MONTH_ID"), 1), FromDate, ToDate)
   FromDate = DateAdd("M", -5, FromDate)
   
   Call LoadEmployee(Nothing, collEmployee)
   
   Call LoadTargetEmp(collTargets)
   
   Call LoadCommissionChartValidFromTo(CommissionChartColl, FromDate, FromDate)
   Call LoadCommissionChartValidFromToEmp(CommissionChartEmpColl, FromDate, FromDate)
     
   Call LoadDnCnAmountBySale(Nothing, CnBySale, FromDate, ToDate, 3)   'ลดหนี้
   Call LoadDnCnAmountBySale(Nothing, DnBySale, FromDate, ToDate, 4)  'เพิ่มหนี้
   
   Call LoadMasterValidCommissionSale(CommissionSaleColl)
   
   Call LoadMasterValidCommissionCost(CommissionCostColl, FromDate, ToDate)
   
   Call LoadCommissionSubtractAmount(Nothing, collCommissionSubtract, Val(mcolParam("YEAR_NO")) - 543, mcolParam("MONTH_ID"))
   
   Call LoadCommissionCredit(Nothing, collCommissionCredit)
   
   
   Call LoadTotalPriceByBill(Nothing, m_BillTotalPrices, , ToDate)
   Call LoadBillingDiscountByBill(Nothing, m_BillingDiscounts, , ToDate)
   Call LoadPaidAmountByBill(Nothing, m_PaidAmounts, , ToDate)
   Call LoadDnCnAmountByBill(Nothing, m_DnItemsByBill, , ToDate, 4, 2)
   Call LoadDnCnAmountByBill(Nothing, m_CnItemsByBill, , ToDate, 3, 2)
   Call LoadDnCnAmountByBill(Nothing, m_RtItemsByBill, , ToDate, 18, 2)
   
   
   MODULE_DESC = vbCrLf & glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & vbCrLf & _
                                       mcolParam("REPORT_NAME") & vbCrLf & _
                                        "จากวันที่ " & EmptyToString(DateToStringExtEx2(FromDate), "N/A") & " " & "ถึง " & EmptyToString(DateToStringExtEx2(ToDate), "N/A") & " วันที่เริ่มคำนวณ " & EmptyToString(DateToStringExtEx2(mcolParam("LIMIT_DATE")), "N/A")
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
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Amt As Double
Dim Di As CDoItem
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(100) As Double
Dim Total2(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim PrevKey2_1 As String
Dim Cm As CCustomer
Dim Pi As CPartItem
Dim ProductNo As String
Dim ProductName As String
Dim IncDiscount As Double
Dim Emp As CEmployee
Dim strPrint As String

Dim tempTargetAmount As Double
Dim PrevKeyID As Long
Dim TempCmsbgc As CCommissionBgChart
Dim SkipEmpFlag As Boolean
Dim ComNowAmount As Double
Dim TempRcp As CReceiptItem
   
Dim FromDate As Date
Dim ToDate As Date

Dim TempTotalSale As CTotalCommission
Dim Rcp As CReceiptItem
Dim PrevInv As String
Dim PrevRcv As String
Dim TempAvgComCost As Double
Dim TotalGp As Double
Dim ComRcpAmount As Double
Dim comSubtract As CCommissionSubtract
Dim comSubtractSub As CCommissionSubtract
Dim RcpDateAmount As Long
Dim RcpDateFromFirstDate As Long
Dim SumComSub As Double
Dim LimitDate As Date
Dim InterestAmount As Double
Dim RcpPercent As Double

Dim comCredit As CCommissionCredit
Dim RealCredit As Long

Dim mInterest As CBillingDoc

Dim Ari As CReceiptItem
Dim Ri1_1 As CReceiptItem
Dim Ri1_2 As CReceiptItem
Dim Ri1_3 As CReceiptItem
Dim Bdc As CBillingDiscount
Dim InCompleteFlag As Boolean
Dim NetAmt As Double
Dim ExceedDay As Long
Dim ExceedDayMonth As Long

Dim tempPart As CPartItem
   
   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set Rs = New ADODB.Recordset
   
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------

   For J = 1 To UBound(Total1)
      Total1(J) = 0
      Total2(J) = 0
   Next J
   
   CurrentHeader = 1
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
   
   I = 0
   Set BodyCf = New CReportField
   
   If CommissionChartColl.Count <= 0 Then
      glbErrorLog.LocalErrorMsg = "ไม่พบช่วงแผนภูมิ COMMISSION"
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Exit Function
   End If
   
   PrevKeyID = 0
   For Each TempCmsbgc In CommissionChartColl
      If TempCmsbgc.MASTER_VALID_ID <> PrevKeyID And PrevKeyID > 0 Then
         glbErrorLog.LocalErrorMsg = vbCrLf & vbCrLf & vbCrLf & "พบช่วงแผนภูมิ COMMISSION มากกว่า 1 ช่วง กรุณาตรวจสอบว่ามีการตั้งแผนภูมิ COMMISSION ทับซ้อนกันหรือไม่" & vbCrLf & vbCrLf & vbCrLf
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Exit Function
      End If
      PrevKeyID = TempCmsbgc.MASTER_VALID_ID
   Next TempCmsbgc
   
   Call GetFirstLastDate(DateSerial(mcolParam("YEAR_NO") - 543, mcolParam("MONTH_ID"), 1), FromDate, ToDate)
   FromDate = DateAdd("M", -5, FromDate)
      
   Set Di = New CDoItem
   Di.DO_ITEM_ID = -1
   
   Di.FROM_DATE = FromDate
   Di.TO_DATE = ToDate
   Di.DocTypeSet = "(1, 2)"
   Call Di.QueryData(35, Rs, iCount)
   
   I = 0
   PrevKey1 = ""
   
   While Not Rs.EOF
      Call Di.PopulateFromRS(35, Rs)
      SkipEmpFlag = False
      
      Set TempCmsbgc = GetObject("CCommissionBgChart", CommissionChartEmpColl, Trim(str(Di.ACCEPT_BY)), False)
      If TempCmsbgc Is Nothing Then
         SkipEmpFlag = True
      End If
      
      If Not SkipEmpFlag Then
         If PrevKey1 <> Di.EMP_CODE And I > 0 Then
            Call GenerateFooter("รวม", "", "", Total1, m_Details1, 0)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
            
            For J = 1 To UBound(Total1)
               Total1(J) = 0
            Next J
                  
            strFormat = VSP_CalTable("10;", mdblWid, alngX)
            strPrint = Di.SALE_NAME & " " & Di.SALE_LNAME & ";"
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", strPrint)
            Vsp.FontBold = False
         ElseIf I = 0 Then
            strFormat = VSP_CalTable("10;", mdblWid, alngX)
            strPrint = Di.SALE_NAME & " " & Di.SALE_LNAME & ";"
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", strPrint)
            Vsp.FontBold = False
         End If
                  
         PrevKey1 = Di.EMP_CODE
         
         Call m_Details1.ClearField
         
         I = I + 1
         J = 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = I
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = (Right(Di.YYYYMM, 2)) & "-" & (Val(Left(Di.YYYYMM, 4)) + 543)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = FormatNumber(Di.PACK_AMOUNT, 0)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(J) = Total1(J) + Di.PACK_AMOUNT
         Total2(J) = Total2(J) + Di.PACK_AMOUNT
   
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = FormatNumber(Di.ITEM_AMOUNT / 1000, 0)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(J) = Total1(J) + (Di.ITEM_AMOUNT / 1000)
         Total2(J) = Total2(J) + (Di.ITEM_AMOUNT / 1000)
         
         IncDiscount = Di.TOTAL_PRICE - Di.EXTRA_DISCOUNT
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = FormatNumber(IncDiscount)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(J) = Total1(J) + IncDiscount
         Total2(J) = Total2(J) + IncDiscount
         
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = FormatNumber(MyDiffEx(IncDiscount, Di.ITEM_AMOUNT))
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set TempRcp = GetObject("CReceiptItem", CnBySale, Trim(Di.ACCEPT_BY & "-" & Di.YYYYMM))
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = FormatNumber(TempRcp.DEBIT_CREDIT_AMOUNT)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(J) = Total1(J) + TempRcp.DEBIT_CREDIT_AMOUNT
         Total2(J) = Total2(J) + TempRcp.DEBIT_CREDIT_AMOUNT
         IncDiscount = IncDiscount - TempRcp.DEBIT_CREDIT_AMOUNT
         
         
         Set TempRcp = GetObject("CReceiptItem", DnBySale, Trim(Di.ACCEPT_BY & "-" & Di.YYYYMM))
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = FormatNumber(TempRcp.DEBIT_CREDIT_AMOUNT)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(J) = Total1(J) + TempRcp.DEBIT_CREDIT_AMOUNT
         Total2(J) = Total2(J) + TempRcp.DEBIT_CREDIT_AMOUNT
         IncDiscount = IncDiscount + TempRcp.DEBIT_CREDIT_AMOUNT
         
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = FormatNumber(IncDiscount)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(J) = Total1(J) + IncDiscount
         Total2(J) = Total2(J) + IncDiscount
         
         tempTargetAmount = GetTarget(Di.ACCEPT_BY, Right(Di.YYYYMM, 2), Val(Left(Di.YYYYMM, 4)))
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = FormatNumber(tempTargetAmount, 0)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(J) = Total1(J) + tempTargetAmount
         Total2(J) = Total2(J) + tempTargetAmount
         
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = FormatNumber(MyDiffEx((Di.ITEM_AMOUNT / 1000) * 100, tempTargetAmount))
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(J) = Total1(J) + 0
         Total2(J) = Total2(J) + 0
         
         
        ComNowAmount = GetSaleNow(MyDiffEx((Di.ITEM_AMOUNT / 1000) * 100, tempTargetAmount), DateSerial(Val(Left(Di.YYYYMM, 4)), Val(Right(Di.YYYYMM, 2)), 1))
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = FormatNumber(ComNowAmount)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(J) = Total1(J) + ComNowAmount
         Total2(J) = Total2(J) + ComNowAmount
         
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
             
         Set TempTotalSale = New CTotalCommission
         TempTotalSale.EMP_ID = Di.ACCEPT_BY
         TempTotalSale.SALE_CODE = Di.EMP_CODE
         TempTotalSale.SALE_NAME = Di.SALE_NAME & " " & Di.SALE_LNAME
         TempTotalSale.YYYYMM = Di.YYYYMM
         
         TempTotalSale.TOTAL_SALE_AMOUNT = (Di.ITEM_AMOUNT / 1000)
         TempTotalSale.TOTAL_SALE = IncDiscount
         
         TempTotalSale.TOTAL_TARGET = tempTargetAmount
         TempTotalSale.SALE_PERCENT = MyDiffEx((Di.ITEM_AMOUNT / 1000) * 100, tempTargetAmount)
         TempTotalSale.COM_NOW = ComNowAmount
         
        Call CollSaleYYYYMM.add(TempTotalSale, TempTotalSale.GetKey)
      End If
      Rs.MoveNext
   Wend

   Call GenerateFooter("รวม", "", "", Total1, m_Details1, 0)
   Call m_Details1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False
               
   Call GenerateFooter("T", "", "", Total2, m_Details1, 0)
   Call m_Details1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False
   
    
    For J = 1 To UBound(Total1)
      Total1(J) = 0
      Total2(J) = 0
   Next J
   CurrentHeader = 2
   Call SumChart
   Call GetFirstLastDate(ToDate, FromDate, ToDate)
   MODULE_DESC = "COM MANAGER" & vbCrLf & _
                                    "จากวันที่ " & EmptyToString(DateToStringExtEx2(FromDate), "N/A") & " " & "ถึง " & EmptyToString(DateToStringExtEx2(ToDate), "N/A") & " วันที่เริ่มคำนวณ " & EmptyToString(DateToStringExtEx2(mcolParam("LIMIT_DATE")), "N/A")
   mblnHeader = False
   Vsp.NewPage
   Call printHeader2
   mblnHeader = True
   
   
   For Each TempTotalSale In SumTotal
      If TempTotalSale.PARENT_ID < 0 Then
         Call m_Details1.ClearField
         J = 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = TempTotalSale.SALE_NAME & "(" & TempTotalSale.SALE_CODE & ")"
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = FormatNumber(TempTotalSale.TOTAL_SALE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(J) = Total1(J) + TempTotalSale.TOTAL_SALE
         Total2(J) = Total2(J) + TempTotalSale.TOTAL_SALE

         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = FormatNumber(TempTotalSale.TOTAL_SALE_AMOUNT)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(J) = Total1(J) + TempTotalSale.TOTAL_SALE_AMOUNT
         Total2(J) = Total2(J) + TempTotalSale.TOTAL_SALE_AMOUNT
         
         tempTargetAmount = GetTarget(TempTotalSale.EMP_ID, mcolParam("MONTH_ID"), mcolParam("YEAR_NO") - 543)
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = FormatNumber(tempTargetAmount, 0)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(J) = Total1(J) + tempTargetAmount
         Total2(J) = Total2(J) + tempTargetAmount
            
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = FormatNumber(MyDiffEx(TempTotalSale.TOTAL_SALE_AMOUNT * 100, tempTargetAmount))
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(J) = Total1(J) + 0
         Total2(J) = Total2(J) + 0
         
         ComNowAmount = GetSaleNowManager(MyDiffEx(TempTotalSale.TOTAL_SALE_AMOUNT * 100, tempTargetAmount), FromDate)
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = FormatNumber(ComNowAmount)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(J) = Total1(J) + ComNowAmount
         Total2(J) = Total2(J) + ComNowAmount
         TempTotalSale.COM_NOW = ComNowAmount
         
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
      End If
   Next TempTotalSale
   
   For J = 1 To UBound(Total1)
      Total1(J) = 0
      Total2(J) = 0
   Next J
   
   'LimitDate = DateSerial(2017, 3, 31)
   
   CurrentHeader = 3
   Call GetFirstLastDate(ToDate, FromDate, ToDate)
   MODULE_DESC = "รายงานการรับชำระ GP COMMISION ตามเลขที่เอกสาร" & vbCrLf & _
                                    "จากวันที่ " & EmptyToString(DateToStringExtEx2(FromDate), "N/A") & " " & "ถึง " & EmptyToString(DateToStringExtEx2(ToDate), "N/A") & " วันที่เริ่มคำนวณ " & EmptyToString(DateToStringExtEx2(mcolParam("LIMIT_DATE")), "N/A")
   mblnHeader = False
   Vsp.NewPage
   Call printHeader3
   mblnHeader = True
   
   '-------------------------------------------------------------------------------------------------
   Set Rcp = New CReceiptItem
   Rcp.RECEIPT_ITEM_ID = -1
   
   Rcp.FROM_DOC_DATE = FromDate                        'จากวันที่ใบเสร็จ
   Rcp.TO_DOC_DATE = ToDate                                  'ถึงวันที่ใบเสร็จ
   Call Rcp.QueryData(120, Rs, iCount)
   
   I = 0
   PrevKey1 = ""
   
   While Not Rs.EOF
      Call Rcp.PopulateFromRS(120, Rs)
      SkipEmpFlag = False
      
      Set TempCmsbgc = GetObject("CCommissionBgChart", CommissionChartEmpColl, Trim(str(Rcp.ACCEPT_BY)), False)
      If TempCmsbgc Is Nothing Then
         SkipEmpFlag = True
      End If
      
      If Not SkipEmpFlag Then
         If PrevKey1 <> Rcp.EMP_CODE And I > 0 Then
            Call GenerateFooter3("รวม", "", "", Total1, m_Details1, 0)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
            
            SumComSub = 0
            Set comSubtract = GetObject("CCommissionSubtract", collCommissionSubtract, Trim(PrevKey2), False)
            If Not comSubtract Is Nothing Then
               For Each comSubtractSub In comSubtract.collCommissionSubTractSub
                  SumComSub = SumComSub + comSubtractSub.COMMISSION_SUBTRACT_AMOUNT
                  Call GenerateFooter3_1(comSubtractSub.CUSTOMER_NAME & " ( " & comSubtractSub.COMMISSION_SUBTRACT_DESC & ") ", FormatNumber(comSubtractSub.COMMISSION_SUBTRACT_AMOUNT), "", m_Details1)
                  Call m_Details1.GetString(1, TempStr1, TempStr2)
                  strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                  Vsp.FontBold = True
                  Call Vsp.AddTable(strFormat, "", TempStr2)
                  Vsp.FontBold = False
               Next comSubtractSub
            End If
            
            Call GenerateFooter3_2("GP สุทธิ", "", "", Total1, SumComSub, PrevKey2, m_Details1, Total2)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
            
            For J = 1 To UBound(Total1)
               Total1(J) = 0
            Next J
                  
            strFormat = VSP_CalTable("10;", mdblWid, alngX)
            strPrint = Rcp.SALE_NAME & " " & Rcp.SALE_LNAME & ";"
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", strPrint)
            Vsp.FontBold = False
         ElseIf I = 0 Then
            strFormat = VSP_CalTable("10;", mdblWid, alngX)
            strPrint = Rcp.SALE_NAME & " " & Rcp.SALE_LNAME & ";"
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", strPrint)
            Vsp.FontBold = False
         End If
                  
         PrevKey1 = Rcp.EMP_CODE
         PrevKey2 = Rcp.ACCEPT_BY
         
         Call m_Details1.ClearField
         
         I = I + 1
         
         J = 1
         Set HeadCf = m_Headers1.Fields(J)
         If PrevRcv <> Rcp.RECEIPT_NO Then
            TempStr = Rcp.RECEIPT_NO
         Else
            TempStr = ""
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         If PrevRcv <> Rcp.RECEIPT_NO Then
            TempStr = DateToStringExtEx2(Rcp.RECEIPT_DATE)
         Else
            TempStr = ""
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         PrevRcv = Rcp.RECEIPT_NO
         
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         If PrevInv <> Rcp.BILL_NO Then
            TempStr = Rcp.BILL_NO
         Else
            TempStr = ""
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         If PrevInv <> Rcp.BILL_NO Then
            TempStr = DateToStringExtEx2(Rcp.BILL_DATE)
         Else
            TempStr = ""
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
          
          J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         If PrevInv <> Rcp.BILL_NO Then
            TempStr = Left(Rcp.CUSTOMER_NAME, 20)
         Else
            TempStr = ""
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         PrevInv = Rcp.BILL_NO
         
         J = J + 1
         RcpPercent = MyDiff(Rcp.PAID_AMOUNT * 100, Rcp.RECEIPT_ITEM_AMOUNT)
         TempStr = FormatNumber(RcpPercent, 0)
         Set TempTotalSale = GetObject("CollSaleYYYYMM", CollSaleYYYYMM, Trim(Rcp.ACCEPT_BY & "-" & Year(Rcp.BILL_DATE) & "-" & Format(Month(Rcp.BILL_DATE), "00")))
         If TempTotalSale Is Nothing Then
            TempStr = TempStr & "-" & "Ivd"
         Else
            TempStr = TempStr & "-" & FormatNumber(TempTotalSale.SALE_PERCENT, 0)
         End If
         If TempTotalSale Is Nothing Then
            ComNowAmount = 0
         Else
            ComNowAmount = GetSaleRcp(TempTotalSale.SALE_PERCENT, Rcp.BILL_DATE)
        End If
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = TempStr & "-" & FormatNumber(ComNowAmount, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(J) = Total1(J) + Rcp.PAID_AMOUNT
         Total2(J) = Total2(J) + Rcp.PAID_AMOUNT
         
         
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = Left(Rcp.PART_NO, 15)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
'         j = j + 1
'         Set HeadCf = m_Headers1.Fields(j)
'         TempStr = FormatNumber(Rcp.PACK_AMOUNT, 0)
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'         Call m_Details1.AddField(BodyCf)
'         Total1(j) = Total1(j) + Rcp.PACK_AMOUNT
'         Total2(j) = Total2(j) + Rcp.PACK_AMOUNT
   
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = FormatNumber(Rcp.ITEM_AMOUNT, 0)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(J) = Total1(J) + Rcp.ITEM_AMOUNT
         Total2(J) = Total2(J) + Rcp.ITEM_AMOUNT
         
         IncDiscount = Rcp.TOTAL_PRICE - Rcp.EXTRA_DISCOUNT
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = FormatNumber(IncDiscount)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(J) = Total1(J) + IncDiscount
         Total2(J) = Total2(J) + IncDiscount
         
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = FormatNumber(MyDiffEx(IncDiscount, Rcp.ITEM_AMOUNT))
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         TempAvgComCost = GetCost(Rcp.PART_ITEM_ID, Rcp.BILL_DATE)
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = FormatNumber(TempAvgComCost)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = FormatNumber(MyDiffEx(IncDiscount, Rcp.ITEM_AMOUNT) - TempAvgComCost)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         If Rcp.BILL_DATE >= mcolParam("LIMIT_DATE") Then
            TotalGp = (MyDiffEx(IncDiscount, Rcp.ITEM_AMOUNT) - TempAvgComCost) * Rcp.ITEM_AMOUNT
            TotalGp = TotalGp * RcpPercent / 100
         Else
            TotalGp = 0
         End If
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = FormatNumber(TotalGp)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(J) = Total1(J) + TotalGp
         Total2(J) = Total2(J) + TotalGp
         
         Set comCredit = GetObject("CCommissionCredit", collCommissionCredit, Trim(str(Rcp.CUSTOMER_ID)), False)
         If Not comCredit Is Nothing Then
            RealCredit = comCredit.COMMISSION_CREDIT_AMOUNT
         Else
            RealCredit = Rcp.Credit + 3
         End If
         
         InterestAmount = 0
         
         
         ' ดอกเบี้ยกรณีเกิน 30 วัน
         ComRcpAmount = TotalGp * ComNowAmount / 100
         RcpDateAmount = DateDiff("D", Rcp.BILL_DATE, Rcp.RECEIPT_DATE) + 1        'จำนวนวันที่รับชำระทั้งหมด
         If RcpDateAmount > (RealCredit) Then 'เช็คว่าเกินกว่า CREDIT ของลูกค้าไหม ตามที่ประชุมให้ T+3
            RcpDateFromFirstDate = DateDiff("D", DateSerial(Year(Rcp.RECEIPT_DATE), Month(Rcp.RECEIPT_DATE), 1), Rcp.RECEIPT_DATE) + 1 'วันที่นับจากต้นเดือน
            If (RcpDateAmount - Rcp.Credit) >= (RcpDateFromFirstDate) Then
               InterestAmount = ComRcpAmount * ((RcpDateFromFirstDate) / 365) * (Val(mcolParam("INTEREST_RATE")) / 100)
            Else
               InterestAmount = ComRcpAmount * ((RcpDateAmount - RealCredit) / 365) * (Val(mcolParam("INTEREST_RATE")) / 100)
            End If
         End If
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = FormatNumber(InterestAmount)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(J) = Total1(J) + InterestAmount
         Total2(J) = Total2(J) + InterestAmount
         
         ComRcpAmount = ComRcpAmount - InterestAmount
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = FormatNumber(ComRcpAmount)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(J) = Total1(J) + ComRcpAmount
         
         'ถ้าค่าคอมเป็น + แต่ไม่มีต้นทุนให้แสดงสีแดง
         If (ComRcpAmount > 0) And (TempAvgComCost <= 0) Then
            Vsp.TextColor = GLB_ALERT_COLOR
            Set tempPart = GetObject("CPartItem", tempCollPartItem, Trim(Rcp.PART_NO & "-" & Format(Month(Rcp.BILL_DATE), "00") & "-" & (Year(Rcp.BILL_DATE) + 543)), False)
            If tempPart Is Nothing Then
               Set tempPart = New CPartItem
               tempPart.PART_NO = Rcp.PART_NO
               tempPart.PART_DESC = Rcp.PART_DESC & " เดือนปี " & Format(Month(Rcp.BILL_DATE), "00") & "-" & Format(Year(Rcp.BILL_DATE) + 543, "0000")
               Call tempCollPartItem.add(tempPart, Trim(Rcp.PART_NO & "-" & Format(Month(Rcp.BILL_DATE), "00") & "-" & (Year(Rcp.BILL_DATE) + 543)))
            End If
         End If
         
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
      End If
      Vsp.TextColor = GLB_NORMAL_COLOR
      Rs.MoveNext
   Wend

   Call GenerateFooter3("รวม", "", "", Total1, m_Details1, 0)
   Call m_Details1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False
   
   
   SumComSub = 0
   Set comSubtract = GetObject("CCommissionSubtract", collCommissionSubtract, Trim(PrevKey2), False)
   If Not comSubtract Is Nothing Then
      For Each comSubtractSub In comSubtract.collCommissionSubTractSub
         SumComSub = SumComSub + comSubtractSub.COMMISSION_SUBTRACT_AMOUNT
         Call GenerateFooter3_1(comSubtractSub.CUSTOMER_NAME & " ( " & comSubtractSub.COMMISSION_SUBTRACT_DESC & ") ", FormatNumber(comSubtractSub.COMMISSION_SUBTRACT_AMOUNT), "", m_Details1)
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
      Next comSubtractSub
   End If
   
   Call GenerateFooter3_2("GP สุทธิ", "", "", Total1, SumComSub, PrevKey2, m_Details1, Total2)
   Call m_Details1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False
   
   Call GenerateFooter3("T", "", "", Total2, m_Details1, 0)
   Call m_Details1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False
   
   Vsp = "* ดอกเบี้ยร้อยละ 15% ต่อปี กรณี เลย 30 วัน บิลคิดย้อนหลัง 6 เดือนเท่านั้น"
   Vsp = "** บิลคิดย้อนหลัง 6 เดือนเท่านั้น"
   J = 0
   If tempCollPartItem.Count > 0 Then
      Vsp = "รายการสินค้าที่ยังไม่มีการใส่ต้นทุน"
      For Each tempPart In tempCollPartItem
         J = J + 1
         Vsp = J & "  " & tempPart.PART_NO & "   " & tempPart.PART_DESC
      Next tempPart
   End If
   
   '------------------------------------------------------------------------------------------------
   For J = 1 To UBound(Total1)
      Total1(J) = 0
      Total2(J) = 0
   Next J
   CurrentHeader = 5
   Call GetFirstLastDate(ToDate, FromDate, ToDate)
   MODULE_DESC = "รายงานดอกเบี้ยปรับ" & vbCrLf & _
                                    "จากวันที่ " & EmptyToString(DateToStringExtEx2(FromDate), "N/A") & " " & "ถึง " & EmptyToString(DateToStringExtEx2(ToDate), "N/A") & " วันที่เริ่มคำนวณ " & EmptyToString(DateToStringExtEx2(mcolParam("LIMIT_DATE")), "N/A") & " ระดับลูกค้า " & EmptyToString(mcolParam("CUSTOMER_GRADE_NAME"), "N/A") & " อัตราดอกเบี้ยล่าช้า " & Val(mcolParam("INTEREST_RATE")) & "% ต่อปี"
   mblnHeader = False
   Vsp.NewPage
   Call printHeader5
   mblnHeader = True
   ExceedDayMonth = DateDiff("D", FromDate, ToDate) + 1
   
   Set mInterest = New CBillingDoc
   mInterest.TO_DATE = ToDate
   mInterest.DOCUMENT_TYPE = 1
   mInterest.CUSTOMER_GRADE = mcolParam("CUSTOMER_GRADE")
   mInterest.OrderType = 1
   Call mInterest.QueryData(111, Rs, iCount)

   
   I = 0
   PrevKey1 = ""
   
   While Not Rs.EOF
      Call mInterest.PopulateFromRS(112, Rs)
      SkipEmpFlag = False
      
      Set TempCmsbgc = GetObject("CCommissionBgChart", CommissionChartEmpColl, Trim(str(mInterest.ACCEPT_BY)), False)
      If TempCmsbgc Is Nothing Then
         SkipEmpFlag = True
      End If
      
      If Not SkipEmpFlag Then
         Call mInterest.PopulateFromRS(111, Rs)
         
         If PrevKey1 <> mInterest.EMP_CODE And I > 0 Then
            Call GenerateFooter5("รวม", Total1, m_Details1, 0)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
            
            For J = 1 To UBound(Total1)
               Total1(J) = 0
            Next J
                  
            strFormat = VSP_CalTable("10;", mdblWid, alngX)
            strPrint = mInterest.EMP_NAME & " " & mInterest.EMP_LNAME & ";"
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", strPrint)
            Vsp.FontBold = False
         End If
         
         Set Di = GetDoItem(m_BillTotalPrices, mInterest.BILLING_DOC_ID)
         Set Bdc = GetBillingDiscount(m_BillingDiscounts, mInterest.BILLING_DOC_ID)
         Set Ari = GetReceiptItem(m_PaidAmounts, mInterest.BILLING_DOC_ID)
         Set Ri1_1 = GetReceiptItem(m_DnItemsByBill, mInterest.BILLING_DOC_ID)
         Set Ri1_2 = GetReceiptItem(m_CnItemsByBill, mInterest.BILLING_DOC_ID)
         Set Ri1_3 = GetReceiptItem(m_RtItemsByBill, mInterest.BILLING_DOC_ID)
         
         InCompleteFlag = ROUND(((Di.TOTAL_PRICE - Bdc.DISCOUNT_AMOUNT) + (Ri1_1.DEBIT_CREDIT_AMOUNT - Ri1_2.DEBIT_CREDIT_AMOUNT - Ri1_3.DEBIT_CREDIT_AMOUNT) - (Ari.PAID_AMOUNT)), 2) <> 0
   
         PrevKey1 = mInterest.EMP_CODE
         If InCompleteFlag Then
            If I = 0 Then
               strFormat = VSP_CalTable("10;", mdblWid, alngX)
               strPrint = mInterest.EMP_NAME & " " & mInterest.EMP_LNAME & ";"
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", strPrint)
               Vsp.FontBold = False
            End If
            
            Call m_Details1.ClearField
            I = I + 1
            
            J = 0
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = mInterest.DOCUMENT_NO
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = DateToStringExtEx2(mInterest.DOCUMENT_DATE)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = mInterest.CUSTOMER_NAME
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
               
            Set comCredit = GetObject("CCommissionCredit", collCommissionCredit, Trim(str(mInterest.CUSTOMER_ID)), False)
            If Not comCredit Is Nothing Then
               RealCredit = comCredit.COMMISSION_CREDIT_AMOUNT
            Else
               RealCredit = mInterest.Credit + 3
            End If
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = RealCredit
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = DateToStringExtEx2(DateAdd("D", RealCredit, mInterest.DOCUMENT_DATE))
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
                        
            ExceedDay = DateDiff("D", DateAdd("D", RealCredit, mInterest.DOCUMENT_DATE), ToDate)
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = ExceedDay
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
                                    
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = FormatNumber(Di.TOTAL_PRICE)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            NetAmt = ROUND(((Di.TOTAL_PRICE - Bdc.DISCOUNT_AMOUNT) + (Ri1_1.DEBIT_CREDIT_AMOUNT - Ri1_2.DEBIT_CREDIT_AMOUNT - Ri1_3.DEBIT_CREDIT_AMOUNT) - (Ari.PAID_AMOUNT)), 2)
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = FormatNumber(NetAmt)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            If mInterest.DOCUMENT_DATE >= mcolParam("LIMIT_DATE") Then
               If ExceedDay > ExceedDayMonth Then
                  InterestAmount = NetAmt * Val(mcolParam("INTEREST_RATE")) / 100 * (ExceedDayMonth / 365)
               Else
                  If ExceedDay < 0 Then
                     InterestAmount = 0
                  Else
                     InterestAmount = NetAmt * Val(mcolParam("INTEREST_RATE")) / 100 * (ExceedDay / 365)
                  End If
               End If
            Else
               InterestAmount = 0
            End If
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = FormatNumber(InterestAmount)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
                                                
            Set TempTotalSale = GetObject("CTotalCommission", SumTotal, Trim(str(mInterest.ACCEPT_BY)))
            TempTotalSale.COM_INTEREST = TempTotalSale.COM_INTEREST + InterestAmount
            
            If InterestAmount > 0 Then
               Total1(J) = Total1(J) + InterestAmount
               
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
            End If
         End If
      End If
      Rs.MoveNext
   Wend
   
   Call GenerateFooter5("รวม", Total1, m_Details1, 0)
   Call m_Details1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False
               
   
   '------------------------------------------------------------------------------------------------
   For J = 1 To UBound(Total1)
      Total1(J) = 0
      Total2(J) = 0
   Next J
   CurrentHeader = 4
   Call GetFirstLastDate(ToDate, FromDate, ToDate)
   MODULE_DESC = "รายงานสรุปค่า COMMISSION" & vbCrLf & _
                                    "จากวันที่ " & EmptyToString(DateToStringExtEx2(FromDate), "N/A") & " " & "ถึง " & EmptyToString(DateToStringExtEx2(ToDate), "N/A") & " วันที่เริ่มคำนวณ " & EmptyToString(DateToStringExtEx2(mcolParam("LIMIT_DATE")), "N/A")
   mblnHeader = False
   Vsp.NewPage
   Call printHeader4
   mblnHeader = True
   
   For Each TempTotalSale In SumTotal
         
      Call m_Details1.ClearField
      J = 1
      Set HeadCf = m_Headers1.Fields(J)
      If TempTotalSale.PARENT_ID > 0 Then
         TempStr = "    " & TempTotalSale.SALE_NAME & "(" & TempTotalSale.SALE_CODE & ")"
      Else
         TempStr = TempTotalSale.SALE_NAME & "(" & TempTotalSale.SALE_CODE & ")"
      End If
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      J = J + 1
      Set HeadCf = m_Headers1.Fields(J)
      TempStr = FormatNumber(TempTotalSale.TOTAL_SALE)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      Total1(J) = Total1(J) + TempTotalSale.TOTAL_SALE
      Total2(J) = Total2(J) + TempTotalSale.TOTAL_SALE
      
      tempTargetAmount = GetTarget(TempTotalSale.EMP_ID, mcolParam("MONTH_ID"), mcolParam("YEAR_NO") - 543)
      J = J + 1
      Set HeadCf = m_Headers1.Fields(J)
      TempStr = FormatNumber(tempTargetAmount, 0)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      Total1(J) = Total1(J) + tempTargetAmount
      Total2(J) = Total2(J) + tempTargetAmount
         
      J = J + 1
      Set HeadCf = m_Headers1.Fields(J)
      TempStr = FormatNumber(MyDiffEx(TempTotalSale.TOTAL_SALE_AMOUNT * 100, tempTargetAmount))
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      Total1(J) = Total1(J) + 0
      Total2(J) = Total2(J) + 0
      
      J = J + 1
      Set HeadCf = m_Headers1.Fields(J)
      TempStr = FormatNumber(TempTotalSale.COM_NOW)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      Total1(J) = Total1(J) + TempTotalSale.COM_NOW
      Total2(J) = Total2(J) + TempTotalSale.COM_NOW
      
      J = J + 1
      Set HeadCf = m_Headers1.Fields(J)
      TempStr = FormatNumber(TempTotalSale.TOTAL_RCP)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      Total1(J) = Total1(J) + TempTotalSale.TOTAL_RCP
      Total2(J) = Total2(J) + TempTotalSale.TOTAL_RCP
      
      J = J + 1
      Set HeadCf = m_Headers1.Fields(J)
      TempStr = FormatNumber(TempTotalSale.COM_RCP)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      Total1(J) = Total1(J) + TempTotalSale.COM_RCP
      Total2(J) = Total2(J) + TempTotalSale.COM_RCP
      
      J = J + 1
      Set HeadCf = m_Headers1.Fields(J)
      TempStr = FormatNumber(TempTotalSale.COM_INTEREST)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      Total1(J) = Total1(J) + TempTotalSale.COM_INTEREST
      Total2(J) = Total2(J) + TempTotalSale.COM_INTEREST
      
      J = J + 1
      Set HeadCf = m_Headers1.Fields(J)
      TempStr = FormatNumber(TempTotalSale.COM_RCP + TempTotalSale.COM_NOW - TempTotalSale.COM_INTEREST)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      Total1(J) = Total1(J) + TempTotalSale.COM_RCP + TempTotalSale.COM_NOW - TempTotalSale.COM_INTEREST
      Total2(J) = Total2(J) + TempTotalSale.COM_RCP + TempTotalSale.COM_NOW - TempTotalSale.COM_INTEREST
      
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Call Vsp.AddTable(strFormat, "", TempStr2)
      
   Next TempTotalSale
   
   Call GenerateFooter4("T", Total1, m_Details1, 0)
   Call m_Details1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False
   
   Set BodyCf = Nothing
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   genDoc = True
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
   Set Rs = Nothing
End Function

Private Sub GenerateFooter(TXT1 As String, TXT2 As String, TXT3 As String, Tot() As Double, Details As CFieldList, DateCount As Long)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim J As Long
Dim TempStr As String
   
   Call Details.ClearField
   Set BodyCf = New CReportField
   
   J = 0
   For Each HeadCf In m_Headers1.Fields
      J = J + 1
      If J = 1 Then
         TempStr = TXT1
      ElseIf J = 2 Then
         TempStr = TXT2
      ElseIf J = 3 Then
         TempStr = TXT3
      ElseIf HeadCf.TempKey = "SUM_FLAG2" Then
         TempStr = FormatNumber(Tot(J), 0)
      ElseIf HeadCf.TempKey = "AVG_FLAG" Then
         TempStr = FormatNumber(MyDiffEx(Tot(6), Tot(5)))
      ElseIf HeadCf.TempKey = "SUM_FLAG1" Then
         TempStr = FormatNumber(Tot(J))
      Else
         TempStr = ""
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.Align)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub
Private Sub GenerateFooter3(TXT1 As String, TXT2 As String, TXT3 As String, Tot() As Double, Details As CFieldList, DateCount As Long)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim J As Long
Dim TempStr As String
   
   Call Details.ClearField
   Set BodyCf = New CReportField
   
   J = 0
   For Each HeadCf In m_Headers1.Fields
      J = J + 1
      If J = 1 Then
         TempStr = TXT1
      ElseIf J = 2 Then
         TempStr = TXT2
      ElseIf J = 3 Then
         TempStr = TXT3
      ElseIf HeadCf.TempKey = "SUM_FLAG2" Then
         TempStr = FormatNumber(Tot(J), 0)
      ElseIf HeadCf.TempKey = "AVG_FLAG" Then
         TempStr = FormatNumber(MyDiffEx(Tot(6), Tot(5)))
      ElseIf HeadCf.TempKey = "SUM_FLAG1" Or HeadCf.TempKey = "SUM_FLAG1-GP" Or HeadCf.TempKey = "SUM_FLAG1-INTEREST" Or HeadCf.TempKey = "SUM_FLAG1-COM" Then
         TempStr = FormatNumber(Tot(J))
      Else
         TempStr = ""
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.Align)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub
Private Sub GenerateFooter4(TXT1 As String, Tot() As Double, Details As CFieldList, DateCount As Long)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim J As Long
Dim TempStr As String
   
   Call Details.ClearField
   Set BodyCf = New CReportField
   
   J = 0
   For Each HeadCf In m_Headers1.Fields
      J = J + 1
      If J = 1 Then
         TempStr = TXT1
      ElseIf HeadCf.TempKey = "SUM_FLAG1" Then
         TempStr = FormatNumber(Tot(J))
      Else
         TempStr = ""
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.Align)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub
Private Sub GenerateFooter5(TXT1 As String, Tot() As Double, Details As CFieldList, DateCount As Long)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim J As Long
Dim TempStr As String
   
   Call Details.ClearField
   Set BodyCf = New CReportField
   
   J = 0
   For Each HeadCf In m_Headers1.Fields
      J = J + 1
      If J = 1 Then
         TempStr = TXT1
      ElseIf HeadCf.TempKey = "SUM_FLAG1" Then
         TempStr = FormatNumber(Tot(J))
      Else
         TempStr = ""
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.Align)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
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
Dim tTextColor As Long

   If Not mblnNewPage Then Exit Sub
   talnT = Vsp.TextAlign
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   blnUnder = Vsp.FontUnderline
   blnItalic = Vsp.FontItalic
   iSize = Vsp.FontSize
   sName = Vsp.FontName
   tTextColor = Vsp.TextColor
   
   Vsp.TextColor = GLB_NORMAL_COLOR
   Vsp.FontSize = 8
   dY(0) = Vsp.CurrentY
   Vsp.TextAlign = taLeftBottom
   strFormat = VSP_CalTable(TITLE_FORMAT, mdblWid, alngX)
   Vsp.TableBorder = tbNone
   Call Vsp.AddTable(strFormat, "", "|หน้าที่|: " & Vsp.CurrentPage)
   Call Vsp.AddTable(strFormat, "", "|วันที่พิมพ์|: " & DateToStringExt(mdteDate))
   Call Vsp.AddTable(strFormat, "", "|เวลา|: " & Right$(DateToStringExtEx(mdteDate), 8))
   Call Vsp.AddTable(strFormat, "", "|ผู้ใช้|: " & glbUser.USER_NAME)
   Vsp.CurrentY = dY(0)
   
   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   Vsp.TextAlign = taCenterBottom
   Vsp.Paragraph = MODULE_DESC
   Vsp.Paragraph = ""
   If mblnHeader Then
      If CurrentHeader = 1 Then
         Call printHeader
      ElseIf CurrentHeader = 2 Then
         Call printHeader2
      ElseIf CurrentHeader = 3 Then
         Call printHeader3
      ElseIf CurrentHeader = 4 Then
         Call printHeader4
      ElseIf CurrentHeader = 5 Then
         Call printHeader5
      End If
   End If
   mdY = Vsp.CurrentY
   
   Vsp.TextAlign = talnT
   Vsp.TableBorder = tbdT
   Vsp.FontBold = blnBold
   Vsp.FontUnderline = blnUnder
   Vsp.FontItalic = blnItalic
   Vsp.FontSize = iSize
   Vsp.FontName = sName
   Vsp.TextColor = tTextColor
End Sub

Private Sub VSP_StartDoc()
'This event occur when VSPrinter.StartDoc is used and used to initialize some information before generating printed document
End Sub
Public Function GetTarget(id As Long, MonthID As Long, YearNo As String) As Double
Dim tempTgdt As CTargetDetail
   Set tempTgdt = GetObject("CTargetdetail", collTargets, Trim(id & "-" & YearNo), False)
   
   If tempTgdt Is Nothing Then
      GetTarget = 0
   Else
      If MonthID = 1 Then
         GetTarget = tempTgdt.TARGET_PRICE1
      ElseIf MonthID = 2 Then
         GetTarget = tempTgdt.TARGET_PRICE2
      ElseIf MonthID = 3 Then
         GetTarget = tempTgdt.TARGET_PRICE3
      ElseIf MonthID = 4 Then
         GetTarget = tempTgdt.TARGET_PRICE4
      ElseIf MonthID = 5 Then
         GetTarget = tempTgdt.TARGET_PRICE5
      ElseIf MonthID = 6 Then
         GetTarget = tempTgdt.TARGET_PRICE6
      ElseIf MonthID = 7 Then
         GetTarget = tempTgdt.TARGET_PRICE7
      ElseIf MonthID = 8 Then
         GetTarget = tempTgdt.TARGET_PRICE8
      ElseIf MonthID = 9 Then
         GetTarget = tempTgdt.TARGET_PRICE9
      ElseIf MonthID = 10 Then
         GetTarget = tempTgdt.TARGET_PRICE10
      ElseIf MonthID = 11 Then
         GetTarget = tempTgdt.TARGET_PRICE11
      ElseIf MonthID = 12 Then
         GetTarget = tempTgdt.TARGET_PRICE12
      End If
   End If
End Function
Private Function GetSaleNow(Percent As Double, InvDate As Date) As Double
Dim tempSale As CCommissionSale
Dim TempValid As CMasterValid
Dim TempComNow As Double

   TempComNow = 0
   
   For Each TempValid In CommissionSaleColl
      If TempValid.VALID_FROM <= InvDate And InvDate <= TempValid.VALID_TO Then
         Exit For
      End If
   Next TempValid
   If TempValid Is Nothing Then
      TempComNow = 0
      Exit Function
   End If
   
   
   For Each tempSale In TempValid.CollSaleNow
      If (Percent >= tempSale.SELL_FROM And Percent < tempSale.SELL_TO) Then
         If tempSale.COMMISSION_SALE_TYPE = 2 Then 'รับหักเป็นบาท
            TempComNow = tempSale.COMMISSION_SALE_AMOUNT
            Exit For
         End If
      End If
   Next tempSale
   
   GetSaleNow = TempComNow
End Function
Private Function GetSaleNowManager(Percent As Double, InvDate As Date) As Double
Dim tempSale As CCommissionSale
Dim TempValid As CMasterValid
Dim TempComNow As Double

   TempComNow = 0
   
   For Each TempValid In CommissionSaleColl
      If TempValid.VALID_FROM <= InvDate And InvDate <= TempValid.VALID_TO Then
         Exit For
      End If
   Next TempValid
   If TempValid Is Nothing Then
      TempComNow = 0
      Exit Function
   End If
   
   
   For Each tempSale In TempValid.CollSaleManagerNow
      If (Percent >= tempSale.SELL_FROM And Percent < tempSale.SELL_TO) Then
         If tempSale.COMMISSION_SALE_TYPE = 2 Then 'รับหักเป็นบาท
            TempComNow = tempSale.COMMISSION_SALE_AMOUNT
            Exit For
         End If
      End If
   Next tempSale
   
   GetSaleNowManager = TempComNow
End Function
Private Function GetSaleRcp(Percent As Double, InvDate As Date) As Double
Dim tempSale As CCommissionSale
Dim TempValid As CMasterValid
Dim TempComNow As Double

   TempComNow = 0
   
   For Each TempValid In CommissionSaleColl
      If TempValid.VALID_FROM <= InvDate And InvDate <= TempValid.VALID_TO Then
         Exit For
      End If
   Next TempValid
   If TempValid Is Nothing Then
      TempComNow = 0
      Exit Function
   End If
   
   
   For Each tempSale In TempValid.CollSaleRcp
      If (Percent >= tempSale.SELL_FROM And Percent < tempSale.SELL_TO) Then
         If tempSale.COMMISSION_SALE_TYPE = 2 Then 'รับหักเป็นบาท
            TempComNow = tempSale.COMMISSION_SALE_AMOUNT
            Exit For
         Else  'เป็น % GP
            TempComNow = tempSale.COMMISSION_SALE_AMOUNT
            Exit For
         End If
      End If
   Next tempSale
   
   GetSaleRcp = TempComNow
End Function
Private Function GetCost(PartItemID As Long, InvDate As Date) As Double
Dim TempComCost As CCommissionCost
Dim TempValid As CMasterValid
Dim TempCost As Double

   GetCost = 0
   
   For Each TempValid In CommissionCostColl
      If TempValid.VALID_FROM <= InvDate And InvDate <= TempValid.VALID_TO Then
         Exit For
      End If
   Next TempValid
   If TempValid Is Nothing Then
      TempCost = 0
      Exit Function
   End If
   
   Set TempComCost = GetObject("CCommissionCost", TempValid.CollCommissionCost, Trim(str(PartItemID)))
   TempCost = TempComCost.COST_AMOUNT
   
   GetCost = TempCost
End Function

Private Sub SumChart()
Dim Cm As CCommissionBgChart
Dim ParentID As Long
   For Each Cm In CommissionChartColl
      Call Recuresive(Cm, GetParent(Cm.COMMISSION_BUDGET_CHART_ID), GetSale(Cm.COMMISSION_BUDGET_CHART_ID), 0)
   Next Cm
End Sub
Public Sub Recuresive(Cm As CCommissionBgChart, ParentID As Long, OwnId As Long, Optional Level As Long)
On Error Resume Next
Dim TotalCommission As CTotalCommission
Dim ComStockBySale  As CTotalCommission
Dim tempSale As CEmployee
   
   If ParentID > 0 Then
      Call Recuresive(Cm, GetParent(ParentID), GetSale(ParentID), Level + 1)
   End If
   
   Set tempSale = GetObject("CEmployee", collEmployee, Trim(str(OwnId)))
   Set ComStockBySale = GetObject("CTotalCommission", CollSaleYYYYMM, Trim(Cm.EMP_ID & "-" & (Val(mcolParam("YEAR_NO") - 543) & "-" & Format(mcolParam("MONTH_ID"), "00"))))    'ดึงยอดขายของ Sale แต่ละรหัสมา
   
   Set TotalCommission = GetObject("CTotalCommission", SumTotal, Trim(OwnId), False)
   If TotalCommission Is Nothing Then
      Set TotalCommission = New CTotalCommission
      TotalCommission.EMP_ID = OwnId
      TotalCommission.PARENT_ID = ParentID
      TotalCommission.SALE_CODE = tempSale.EMP_CODE
      TotalCommission.SALE_NAME = tempSale.NAME & " " & tempSale.LASTNAME
      Call SumTotal.add(TotalCommission, Trim(TotalCommission.EMP_ID))
      
   End If
   
   TotalCommission.TOTAL_SALE_AMOUNT = TotalCommission.TOTAL_SALE_AMOUNT + ComStockBySale.TOTAL_SALE_AMOUNT
   TotalCommission.TOTAL_SALE = TotalCommission.TOTAL_SALE + ComStockBySale.TOTAL_SALE
   TotalCommission.COM_NOW = TotalCommission.COM_NOW + ComStockBySale.COM_NOW
End Sub
Private Function GetParent(id As Long) As Long
Dim Cm As CCommissionBgChart
   Set Cm = GetObject("CCommissionBgChart", CommissionChartColl, Trim(str(id)))
   GetParent = Cm.PARENT_ID
End Function
Private Function GetSale(id As Long) As Long
Dim Cm As CCommissionBgChart
   Set Cm = GetObject("CCommissionBgChart", CommissionChartColl, Trim(str(id)))
   GetSale = Cm.EMP_ID
End Function
Private Sub GenerateFooter3_1(TXT1 As String, TXT2 As String, TXT3 As String, Details As CFieldList)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim J As Long
Dim TempStr As String
Dim TempWidth1 As Double
Dim TempWidth2 As Double
Dim TempWidth3 As Double
Dim TempWidth4 As Double

   Call Details.ClearField
   Set BodyCf = New CReportField
   
    J = 0
   For Each HeadCf In m_Headers1.Fields
      J = J + 1
      If HeadCf.TempKey = "SUM_FLAG1-COM" Then
         TempWidth4 = TempWidth4 + HeadCf.Width
      ElseIf HeadCf.TempKey = "SUM_FLAG1-INTEREST" Then
         TempWidth3 = TempWidth3 + HeadCf.Width
      ElseIf HeadCf.TempKey = "SUM_FLAG1-GP" Then
         TempWidth2 = TempWidth2 + HeadCf.Width
      Else
         TempWidth1 = TempWidth1 + HeadCf.Width
      End If
   Next HeadCf
   
   TempStr = TXT1
   Call BodyCf.SetFieldValue(TempWidth1, ">", TempStr)
   Call Details.AddField(BodyCf)
      
   TempStr = TXT2
   Call BodyCf.SetFieldValue(TempWidth2, ">", TempStr)
   Call Details.AddField(BodyCf)
   
   TempStr = TXT3
   Call BodyCf.SetFieldValue(TempWidth3, ">", TempStr)
   Call Details.AddField(BodyCf)
      
   TempStr = ""
   Call BodyCf.SetFieldValue(TempWidth4, ">", TempStr)
   Call Details.AddField(BodyCf)
      
   Set BodyCf = Nothing
End Sub
Private Sub GenerateFooter3_2(TXT1 As String, TXT2 As String, TXT3 As String, Tot() As Double, SubTractAmount As Double, EmpID As String, Details As CFieldList, Tot2() As Double)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim J As Long
Dim TempStr As String
Dim TempWidth1 As Double
Dim TempWidth2 As Double
Dim TempWidth3 As Double
Dim TempWidth4 As Double
Dim ComRcpAmount As Double
Dim TempTotalSale As CTotalCommission
Dim SumPaidAmount As Double
Dim SumInterest As Double

   Call Details.ClearField
   Set BodyCf = New CReportField
   
    J = 0
   For Each HeadCf In m_Headers1.Fields
      J = J + 1
      If HeadCf.TempKey = "SUM_FLAG1-COM" Then
         TempWidth4 = TempWidth4 + HeadCf.Width
         ComRcpAmount = Tot(J) * (1 - (MyDiff(SubTractAmount, Tot(J - 2))))
         Tot2(J) = Tot2(J) + ComRcpAmount
      ElseIf HeadCf.TempKey = "SUM_FLAG1-INTEREST" Then
         TempWidth3 = TempWidth3 + HeadCf.Width
         SumInterest = Tot(J)
      ElseIf HeadCf.TempKey = "SUM_FLAG1-GP" Then
         TempWidth2 = TempWidth2 + HeadCf.Width
      Else
         TempWidth1 = TempWidth1 + HeadCf.Width
      End If
      
      If HeadCf.TempKey = "KEY_PAID" Then
         SumPaidAmount = Tot(J)
      End If
   Next HeadCf
   
   TempStr = TXT1
   Call BodyCf.SetFieldValue(TempWidth1, ">", TempStr)
   Call Details.AddField(BodyCf)
      
   TempStr = TXT2
   Call BodyCf.SetFieldValue(TempWidth2, ">", TempStr)
   Call Details.AddField(BodyCf)
   
   TempStr = FormatNumber(SumInterest)
   Call BodyCf.SetFieldValue(TempWidth3, ">", TempStr)
   Call Details.AddField(BodyCf)
   
   TempStr = FormatNumber(ComRcpAmount)
   Call BodyCf.SetFieldValue(TempWidth4, ">", TempStr)
   Call Details.AddField(BodyCf)
   
   Set TempTotalSale = GetObject("CTotalCommission", SumTotal, Trim(EmpID))
   TempTotalSale.COM_RCP = ComRcpAmount
   TempTotalSale.TOTAL_RCP = SumPaidAmount
   
   Set BodyCf = Nothing
End Sub

