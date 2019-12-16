VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportFrelance001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportFrelance001"
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
Private tempCollIncentive As Collection
Private tempCollIncentiveCusPD As Collection
Private tempCollIncentiveComExtra As Collection
Private tempCollIncentiveIncExtra As Collection

Private m_PaidAmounts As Collection
Private m_BillTotalPrices As Collection
Private m_DnItemsByBill As Collection
Private m_CnItemsByBill As Collection
Private m_RtItemsByBill As Collection
Private m_BillingDiscounts As Collection
Private tempRcpNoPaidColl As Collection
Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_TempCol = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Details1 = New CFieldList
   Set m_Details2 = New CFieldList

   Set tempCollIncentive = New Collection
   Set tempCollIncentiveCusPD = New Collection
   Set tempCollIncentiveComExtra = New Collection
   Set tempCollIncentiveIncExtra = New Collection
   
   Set m_PaidAmounts = New Collection
   Set m_BillTotalPrices = New Collection
   Set m_DnItemsByBill = New Collection
   Set m_CnItemsByBill = New Collection
   Set m_RtItemsByBill = New Collection
   Set m_BillingDiscounts = New Collection
   Set tempRcpNoPaidColl = New Collection
   
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_Details2 = Nothing

   Set tempCollIncentive = Nothing
   Set tempCollIncentiveCusPD = Nothing
   Set tempCollIncentiveComExtra = Nothing
   Set tempCollIncentiveIncExtra = Nothing
   
   Set m_PaidAmounts = Nothing
   Set m_BillTotalPrices = Nothing
   Set m_DnItemsByBill = Nothing
   Set m_CnItemsByBill = Nothing
   Set m_RtItemsByBill = Nothing
   Set m_BillingDiscounts = Nothing
   Set tempRcpNoPaidColl = Nothing
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
      
   Call Cf.SetFieldValue(10, "^", "No.", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(20, "^", "รหัสลูกค้า/PC", "<")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(40, "^", "ชื่อลูกค้า/PC", "<")
   Call m_Headers1.AddField(Cf)
   
      Call Cf.SetFieldValue(15, "^", "ค่า Com/ถุง", ">")
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
   
      Call LoadCommissionIncentive(Nothing, tempCollIncentive, 1, mcolParam("FREELANCE_CODE"), mcolParam("FROM_FREELANCE_CODE"), mcolParam("TO_FREELANCE_CODE"))
      Call LoadCommissionIncentive(Nothing, tempCollIncentiveCusPD, 2, mcolParam("FREELANCE_CODE"), mcolParam("FROM_FREELANCE_CODE"), mcolParam("TO_FREELANCE_CODE"))
      Call LoadCommissionIncentive(Nothing, tempCollIncentiveComExtra, 3, mcolParam("FREELANCE_CODE"), mcolParam("FROM_FREELANCE_CODE"), mcolParam("TO_FREELANCE_CODE"))
      Call LoadCommissionIncentive(Nothing, tempCollIncentiveIncExtra, 4, mcolParam("FREELANCE_CODE"), mcolParam("FROM_FREELANCE_CODE"), mcolParam("TO_FREELANCE_CODE"))
   
'      Call GetFirstLastDate(DateSerial(mcolParam("YEAR_NO") - 543, mcolParam("MONTH_ID"), 1), FromDate, ToDate)
      FromDate = mcolParam("FROM_DATE")
      ToDate = mcolParam("TO_DATE")
      
      Call LoadPaidAmountByBill(Nothing, m_PaidAmounts, -1, ToDate)
      Call LoadTotalPriceByBill(Nothing, m_BillTotalPrices, -1, ToDate)
      Call LoadBillingDiscountByBill(Nothing, m_BillingDiscounts, -1, ToDate)
      Call LoadDnCnAmountByBill(Nothing, m_DnItemsByBill, -1, ToDate, 4, 2)
      Call LoadDnCnAmountByBill(Nothing, m_CnItemsByBill, -1, ToDate, 3, 2)
      Call LoadDnCnAmountByBill(Nothing, m_RtItemsByBill, -1, ToDate, 18, 2)
   
   MODULE_DESC = vbCrLf & glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & vbCrLf & _
                                       mcolParam("REPORT_NAME") & vbCrLf & _
                                       "จากวันที่ " & EmptyToString(DateToStringExtEx2(FromDate), "N/A") & " ถึง " & EmptyToString(DateToStringExtEx2(ToDate), "N/A")
   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))
   
   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
      mstrErrMsg = "Error GetDateTime Error."
      Exit Function
   End If
   mdteDate = InternalDateToDate(strDate)
   initDoc = True
End Function
Function CalComExtra(Cl As Collection, FrelanceId As Long, Optional PackOver As Double = 0, Optional WeightOver As Double = 0, Optional PackTrue As Double = 0, Optional ByRef A As Double) As Double
Dim D As CCommissionIncentive
Dim TempData As CCommissionIncentive
Dim TempValue1 As Double
Dim TempValue2 As Double
Dim AmountOver As Double
CalComExtra = 0
A = -1
For Each TempData In Cl
   If TempData.UNIT_TYPE = 1 Then 'ยอดเป็นถุง
      AmountOver = PackOver
   Else
      AmountOver = WeightOver \ 1000
   End If
   If TempData.FREELANCE_ID = FrelanceId Then
      If (AmountOver >= TempData.FROM_AMOUNT) And (AmountOver <= TempData.TO_AMOUNT) Then
         TempValue1 = TempData.INCENTIVE_PER_PACK * AmountOver
         A = TempData.INCENTIVE_PER_PACK
      End If
      If TempData.AMOUNT_OVER_FLAG = "Y" Then
         If (AmountOver - TempData.TO_AMOUNT) > 0 Then
            TempValue1 = TempData.INCENTIVE_PER_PACK * (AmountOver - (AmountOver - TempData.TO_AMOUNT))
            A = TempData.INCENTIVE_PER_PACK
         End If
         
         If (AmountOver - TempData.TO_AMOUNT) > 0 Then
           TempValue2 = TempData.RATE_OVER * (AmountOver - TempData.TO_AMOUNT)
           
           If (AmountOver - PackTrue) > TempData.TO_AMOUNT Then
             A = TempData.RATE_OVER
           Else
             A = (((TempData.TO_AMOUNT - (AmountOver - PackTrue)) * TempData.INCENTIVE_PER_PACK) + ((AmountOver - TempData.TO_AMOUNT) * TempData.RATE_OVER)) / PackTrue
           End If
         End If
      End If
      CalComExtra = CalComExtra + TempValue1 + TempValue2
      TempValue1 = 0
      TempValue2 = 0
   End If
Next TempData
If CalComExtra = 0 Then
   CalComExtra = -1
End If
End Function
Function CalIncExtra(Cl As Collection, FrelanceId As Long, PackOver As Double, WeightOver As Double) As Double
Dim D As CCommissionIncentive
Dim TempData As CCommissionIncentive
Dim TempValue1 As Double
Dim TempValue2 As Double
Dim AmountOver As Double
CalIncExtra = 0
For Each TempData In Cl
   If TempData.UNIT_TYPE = 1 Then 'ยอดเป็นถุง
      AmountOver = PackOver
   Else
      AmountOver = WeightOver \ 1000
   End If
   If TempData.FREELANCE_ID = FrelanceId Then
      If (AmountOver >= TempData.FROM_AMOUNT) And (AmountOver <= TempData.TO_AMOUNT) Then
         TempValue1 = TempData.INCENTIVE_PER_PACK
      End If
      If TempData.AMOUNT_OVER_FLAG = "Y" Then
         If (AmountOver - TempData.TO_AMOUNT) > 0 Then
            TempValue1 = TempData.INCENTIVE_PER_PACK
         End If
         If (AmountOver - TempData.TO_AMOUNT) > 0 Then
           TempValue2 = TempData.RATE_OVER * (AmountOver - TempData.TO_AMOUNT)
         End If
      End If
      CalIncExtra = CalIncExtra + TempValue1 + TempValue2
      TempValue1 = 0
      TempValue2 = 0
   End If
Next TempData
If CalIncExtra = 0 Then
   CalIncExtra = -1
End If
End Function
Function FindCollFrelanceCode(Cl As Collection, FrelanceId As Long) As String
Dim TempData As CCommissionIncentive
FindCollFrelanceCode = ""
For Each TempData In Cl
   If TempData.FREELANCE_ID = FrelanceId Then
      FindCollFrelanceCode = TempData.EMP_CODE
      Exit For
   End If
Next TempData
End Function
Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim RName As String
Dim I As Long
Dim J As Long
Dim J2 As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Amt As Double
Dim Amt2 As Double
Dim TempIncentive As Double
Dim Di As CDoItem
Dim TempColl As Collection
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(100) As Double
Dim Total2(100) As Double
Dim Total3(100) As Double
Dim Total4(100) As Double

Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim PrevKey3 As String
Dim PrevKey4 As String
Dim PrevKey2_1 As String
Dim PrevKey5 As String
Dim PrevKey6 As String
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
Dim TempCCommissionIncentive As CCommissionIncentive
Dim SkipEmpFlag As Boolean
Dim ComNowAmount As Double
Dim TempRcp As CReceiptItem
Dim TempRcp2 As CReceiptItem
   
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
Dim tempRcpPaid As CReceiptItem
Dim PaidFlag As Boolean
Dim AlertNoData As String

Dim tempCollReceiptItem As Collection
Dim tempReceiptItem  As CReceiptItem
Dim TempCollection As Collection
Dim EmpCode As String
Dim tempDoItem As CDoItem

Set tempCollReceiptItem = New Collection
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
      Total3(J) = 0
      Total4(J) = 0
   Next J
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True

   I = 0
   Set BodyCf = New CReportField
   '-------------------------------------------------------------------------------------------------
   Set Rcp = New CReceiptItem
   Rcp.RECEIPT_ITEM_ID = -1
   Rcp.FROM_DOC_DATE = mcolParam("FROM_DATE") 'FromDate                        'จากวันที่ใบเสร็จ
   Rcp.TO_DOC_DATE = mcolParam("TO_DATE") 'ToDate                                  'ถึงวันที่ใบเสร็จ
   Rcp.FREELANCE_CODE = mcolParam("FREELANCE_CODE")
   Rcp.FROM_FREELANCE_CODE = mcolParam("FROM_FREELANCE_CODE")
   Rcp.TO_FREELANCE_CODE = mcolParam("TO_FREELANCE_CODE")
   Rcp.CUSTOMER_CODE = mcolParam("CUSTOMER_CODE")
   Rcp.FROM_CUSTOMER_CODE = mcolParam("FROM_CUSTOMER_CODE")
   Rcp.TO_CUSTOMER_CODE = mcolParam("TO_CUSTOMER_CODE")
   Call Rcp.QueryData(121, Rs, iCount)

   I = 0
   PrevKey1 = ""
   
   Set TempColl = New Collection
   While Not Rs.EOF
      Call Rcp.PopulateFromRS(121, Rs)
      
      PrevKey3 = Rcp.CUSTOMER_CODE
      PrevKey5 = Rcp.FREELANCE_ID

      Set Di = GetDoItem(m_BillTotalPrices, Rcp.BILLING_DOC_ID)
      Set Bdc = GetBillingDiscount(m_BillingDiscounts, Rcp.BILLING_DOC_ID)
      Set Ari = GetReceiptItem(m_PaidAmounts, Rcp.BILLING_DOC_ID)
      Set Ri1_1 = GetReceiptItem(m_DnItemsByBill, Rcp.BILLING_DOC_ID)
      Set Ri1_2 = GetReceiptItem(m_CnItemsByBill, Rcp.BILLING_DOC_ID)
      Set Ri1_3 = GetReceiptItem(m_RtItemsByBill, Rcp.BILLING_DOC_ID)
      If ROUND((Di.TOTAL_PRICE - Bdc.DISCOUNT_AMOUNT) + (Ri1_1.DEBIT_CREDIT_AMOUNT - Ri1_2.DEBIT_CREDIT_AMOUNT - Ri1_3.DEBIT_CREDIT_AMOUNT), 2) <> ROUND(Ari.PAID_AMOUNT, 2) Then    'จ่ายยังไม่ครบ
         SkipEmpFlag = True
      Else
         SkipEmpFlag = False
      End If
      
      If Not SkipEmpFlag Then
      Set TempRcp = GetObject("CReceiptItem", tempRcpNoPaidColl, Trim(Rcp.FREELANCE_ID & "-" & Rcp.BILL_NO), False)        'ถ้าบิลไหนจ่ายไปแล้วไม่ต้องจ่ายแล้วในกรณีที่เป็นใบเสร็จอื่นๆ
         If TempRcp Is Nothing Then
         Set TempRcp = New CReceiptItem
         TempRcp.RECEIPT_NO = Rcp.RECEIPT_NO
         TempRcp.BILL_NO = Rcp.BILL_NO
         TempRcp.FREELANCE_ID = Rcp.FREELANCE_ID
         Call tempRcpNoPaidColl.add(TempRcp, Trim(Rcp.FREELANCE_ID & "-" & Rcp.BILL_NO))
         PaidFlag = False                                            'ยังไม่ได้จ่ายให้ FREELANCE คนนี้
         Else
            If Rcp.RECEIPT_NO = TempRcp.RECEIPT_NO Then
            PaidFlag = False
            Else
            PaidFlag = True                                          'จ่ายให้ FREELANCE คนนี้แล้ว
            End If
         End If
      End If


  
      Set TempCCommissionIncentive = GetObject("CCommissionIncentive", tempCollIncentiveCusPD, Trim(Rcp.FREELANCE_ID & "-" & Rcp.CUSTOMER_ID & "-" & Rcp.PART_ITEM_ID), False)
      If TempCCommissionIncentive Is Nothing Then
         Set TempCCommissionIncentive = GetObject("CCommissionIncentive", tempCollIncentive, Trim(Rcp.FREELANCE_ID & "-" & Rcp.PART_ITEM_ID))
      End If
      If SkipEmpFlag Then
         Amt = 0
      ElseIf PaidFlag Then
         Amt = 0
      Else
      Set tempReceiptItem = GetObject("CReceiptItem", tempCollReceiptItem, Trim(PrevKey5), False)
      If tempReceiptItem Is Nothing Then
         Set tempReceiptItem = New CReceiptItem
         tempReceiptItem.PACK_AMOUNT = Rcp.PACK_AMOUNT
         Call tempCollReceiptItem.add(tempReceiptItem, Trim(PrevKey5))
      Else
         tempReceiptItem.PACK_AMOUNT = tempReceiptItem.PACK_AMOUNT + Rcp.PACK_AMOUNT
      End If
      
      Amt = CalComExtra(tempCollIncentiveComExtra, Val(PrevKey5), tempReceiptItem.PACK_AMOUNT, , Rcp.PACK_AMOUNT, TempIncentive)
      If TempIncentive = -1 Then
         TempIncentive = TempCCommissionIncentive.INCENTIVE_PER_PACK
      End If
         PrevKey6 = Trim(str(TempIncentive))
      End If

      If PrevKey1 <> Rcp.FREELANCE_CODE And I > 0 Then
         Set TempRcp = GetObject("CReceiptItem", TempColl, Trim(PrevKey1 & "-" & PrevKey3 & "-" & PrevKey6), False)
         If TempRcp Is Nothing And TempIncentive <> 0 Then
           PrevKey1 = Rcp.FREELANCE_CODE
            I = 0
            strFormat = VSP_CalTable("10;", mdblWid, alngX)
            strPrint = "P/C :" & Rcp.FREELANCE_NAME & " " & Rcp.FREELANCE_LASTNAME & " (" & Rcp.FREELANCE_CODE & ");"
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", strPrint)
            Vsp.FontBold = False
         End If
      ElseIf I = 0 Then
         PrevKey1 = Rcp.FREELANCE_CODE
         strFormat = VSP_CalTable("10;", mdblWid, alngX)
         strPrint = "P/C :" & Rcp.FREELANCE_NAME & " " & Rcp.FREELANCE_LASTNAME & " (" & Rcp.FREELANCE_CODE & ");"
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", strPrint)
         Vsp.FontBold = False
      End If

      Call m_Details1.ClearField
      I = I + 1
      J = 1

      Set HeadCf = m_Headers1.Fields(J)
      TempStr = I
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

      J = J + 1
      Set HeadCf = m_Headers1.Fields(J)
      TempStr = Rcp.CUSTOMER_CODE
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

       J = J + 1
      Set HeadCf = m_Headers1.Fields(J)
      TempStr = Rcp.CUSTOMER_NAME
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      J = J + 1
      Set HeadCf = m_Headers1.Fields(J)
      TempStr = IsDouble(TempIncentive)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

      Set TempRcp = GetObject("CReceiptItem", TempColl, Trim(PrevKey1 & "-" & PrevKey3 & "-" & PrevKey6), False)
      If TempRcp Is Nothing And TempIncentive <> 0 Then
         Call TempColl.add(Rcp, Trim(PrevKey1 & "-" & PrevKey3 & "-" & PrevKey6))
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
      Else
         I = I - 1
      End If
      Vsp.TextColor = GLB_NORMAL_COLOR
      Rs.MoveNext
   Wend
   
   
'   Set TempColl = New Collection
'   While Not Rs.EOF
'      Call Rcp.PopulateFromRS(121, Rs)
'
'      Set Di = GetDoItem(m_BillTotalPrices, Rcp.BILLING_DOC_ID)
'      Set Bdc = GetBillingDiscount(m_BillingDiscounts, Rcp.BILLING_DOC_ID)
'      Set Ari = GetReceiptItem(m_PaidAmounts, Rcp.BILLING_DOC_ID)
'      Set Ri1_1 = GetReceiptItem(m_DnItemsByBill, Rcp.BILLING_DOC_ID)
'      Set Ri1_2 = GetReceiptItem(m_CnItemsByBill, Rcp.BILLING_DOC_ID)
'      Set Ri1_3 = GetReceiptItem(m_RtItemsByBill, Rcp.BILLING_DOC_ID)
'      If ROUND((Di.TOTAL_PRICE - Bdc.DISCOUNT_AMOUNT) + (Ri1_1.DEBIT_CREDIT_AMOUNT - Ri1_2.DEBIT_CREDIT_AMOUNT - Ri1_3.DEBIT_CREDIT_AMOUNT), 2) <> ROUND(Ari.PAID_AMOUNT, 2) Then    'จ่ายยังไม่ครบ
'         SkipEmpFlag = True
'      Else
'         SkipEmpFlag = False
'      End If
'
'
'      If PrevKey1 <> Rcp.FREELANCE_CODE And I > 0 Then
'               I = 0
'               strFormat = VSP_CalTable("10;", mdblWid, alngX)
'               strPrint = "P/C :" & Rcp.FREELANCE_NAME & " " & Rcp.FREELANCE_LASTNAME & ";"
'               Vsp.FontBold = True
'               Call Vsp.AddTable(strFormat, "", strPrint)
'               Vsp.FontBold = False
'      ElseIf I = 0 Then
'            strFormat = VSP_CalTable("10;", mdblWid, alngX)
'            strPrint = "P/C :" & Rcp.FREELANCE_NAME & " " & Rcp.FREELANCE_LASTNAME & ";"
'            Vsp.FontBold = True
'            Call Vsp.AddTable(strFormat, "", strPrint)
'            Vsp.FontBold = False
'      End If
'
'      PrevKey2 = Rcp.ACCEPT_BY
'      PrevKey2_1 = Rcp.FREELANCE_NAME & " " & Rcp.FREELANCE_LASTNAME
'
'      PrevKey4 = Rcp.CUSTOMER_NAME
'      PrevKey5 = Rcp.FREELANCE_ID
'
'      Call m_Details1.ClearField
'
'      I = I + 1
'       J = 1
'
'        If Not SkipEmpFlag Then
'            Set TempRcp = GetObject("CReceiptItem", tempRcpNoPaidColl, Trim(Rcp.FREELANCE_ID & "-" & Rcp.BILL_NO), False)        'ถ้าบิลไหนจ่ายไปแล้วไม่ต้องจ่ายแล้วในกรณีที่เป็นใบเสร็จอื่นๆ
'            If TempRcp Is Nothing Then
'               Set TempRcp = New CReceiptItem
'               TempRcp.RECEIPT_NO = Rcp.RECEIPT_NO
'               TempRcp.BILL_NO = Rcp.BILL_NO
'               TempRcp.FREELANCE_ID = Rcp.FREELANCE_ID
'               Call tempRcpNoPaidColl.add(TempRcp, Trim(Rcp.FREELANCE_ID & "-" & Rcp.BILL_NO))
'
'               PaidFlag = False                                            'ยังไม่ได้จ่ายให้ FREELANCE คนนี้
'            Else
'               If Rcp.RECEIPT_NO = TempRcp.RECEIPT_NO Then
'                  PaidFlag = False
'               Else
'                  PaidFlag = True                                          'จ่ายให้ FREELANCE คนนี้แล้ว
'               End If
'            End If
'        End If
'
'            Set HeadCf = m_Headers1.Fields(J)
'            TempStr = I
'            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'            Call m_Details1.AddField(BodyCf)
'
'            J = J + 1
'            Set HeadCf = m_Headers1.Fields(J)
'            TempStr = Rcp.CUSTOMER_CODE
'            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'            Call m_Details1.AddField(BodyCf)
'
'             J = J + 1
'            Set HeadCf = m_Headers1.Fields(J)
'            TempStr = Rcp.CUSTOMER_NAME
'            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'            Call m_Details1.AddField(BodyCf)
'
'            Set TempCCommissionIncentive = GetObject("CCommissionIncentive", tempCollIncentiveCusPD, Trim(Rcp.FREELANCE_ID & "-" & Rcp.CUSTOMER_ID & "-" & Rcp.PART_ITEM_ID), False)
'            If TempCCommissionIncentive Is Nothing Then
'               Set TempCCommissionIncentive = GetObject("CCommissionIncentive", tempCollIncentive, Trim(Rcp.FREELANCE_ID & "-" & Rcp.PART_ITEM_ID))
'            End If
'
'            J = J + 1
'            Set HeadCf = m_Headers1.Fields(J)
'            If SkipEmpFlag Then
'               Amt = 0
'               TempStr = ""
'            ElseIf PaidFlag Then
'               Amt = 0
'               TempStr = ""
'            Else
'
'            Set tempReceiptItem = GetObject("CReceiptItem", tempCollReceiptItem, Trim(PrevKey5), False)
'            If tempReceiptItem Is Nothing Then
'               Set tempReceiptItem = New CReceiptItem
'               tempReceiptItem.PACK_AMOUNT = Rcp.PACK_AMOUNT
'               Call tempCollReceiptItem.add(tempReceiptItem, Trim(PrevKey5))
'            Else
'               tempReceiptItem.PACK_AMOUNT = tempReceiptItem.PACK_AMOUNT + Rcp.PACK_AMOUNT
'            End If
'
'            Amt = CalComExtra(tempCollIncentiveComExtra, Val(PrevKey5), tempReceiptItem.PACK_AMOUNT, , Rcp.PACK_AMOUNT, Amt2)
'            If Amt2 = -1 Then
'               Amt2 = TempCCommissionIncentive.INCENTIVE_PER_PACK
'            End If
'            TempIncentive = Amt2
'            TempStr = IsDouble(Amt2)
'         End If
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'         Call m_Details1.AddField(BodyCf)
'
'            PrevKey1 = Rcp.FREELANCE_CODE
'            PrevKey3 = Rcp.CUSTOMER_CODE
'            PrevKey6 = Trim(str(TempIncentive))
'
'            Set TempRcp = GetObject("CReceiptItem", TempColl, Trim(PrevKey1 & "-" & PrevKey3 & "-" & PrevKey6), False)
'            If TempRcp Is Nothing And TempIncentive <> 0 Then
'
'               Call TempColl.add(Rcp, Trim(PrevKey1 & "-" & PrevKey3 & "-" & PrevKey6))
'
'               Call m_Details1.GetString(1, TempStr1, TempStr2)
'               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'               Call Vsp.AddTable(strFormat, "", TempStr2)
'            Else
'               I = I - 1
'            End If
'         Vsp.TextColor = GLB_NORMAL_COLOR
'      Rs.MoveNext
'   Wend
   
   Set BodyCf = Nothing
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   genDoc = True
   
   Set tempCollIncentive = Nothing
   Set tempCollIncentiveCusPD = Nothing
   
   Set m_PaidAmounts = Nothing
'   Set m_BillTotalPrices = Nothing
   Set m_DnItemsByBill = Nothing
   Set m_CnItemsByBill = Nothing
   Set m_RtItemsByBill = Nothing
   Set m_BillingDiscounts = Nothing
   Set tempRcpNoPaidColl = Nothing
   
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
   Set Rs = Nothing
End Function
Function IsDouble(ByVal varValue As Variant) As String
   Dim Arr() As String
   Arr = Split(varValue, ".")
   If UBound(Arr) > 0 Then '
      IsDouble = FormatNumber(varValue, 2)
   Else
      IsDouble = str(varValue)
   End If
End Function
Private Sub GenerateFooter(TXT1 As String, TXT2 As String, TXT3 As String, Tot() As Double, Details As CFieldList, DateCount As Long, Optional FrelanceId As Long)
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
      ElseIf J = 5 Then
         TempStr = TXT2
      ElseIf J = 3 Then
         TempStr = TXT3
      ElseIf HeadCf.TempKey = "SUM_FLAG2" Then
         TempStr = FormatNumber(Tot(J), 0)
      ElseIf HeadCf.TempKey = "AVG_FLAG" Then
         TempStr = FormatNumber(MyDiffEx(Tot(6), Tot(5)))
      ElseIf HeadCf.TempKey = "SUM_FLAG1" Or HeadCf.TempKey = "SUM_FLAG1-GP" Or HeadCf.TempKey = "SUM_FLAG1-INTEREST" Or HeadCf.TempKey = "SUM_FLAG1-COM" Then
         TempStr = FormatNumber(Tot(J))
      ElseIf HeadCf.TempKey = "COM_EXTRA_FLAG" Then
         TempStr = FormatNumber(Tot(J))
      ElseIf HeadCf.TempKey = "INC_EXTRA_FLAG" Then
         TempStr = FormatNumber(Tot(J))
      ElseIf HeadCf.TempKey = "SUM_NET_FLAG1" Then
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
   Vsp.TextColor = tTextColor
End Sub

Private Sub VSP_StartDoc()
'This event occur when VSPrinter.StartDoc is used and used to initialize some information before generating printed document
End Sub
'





