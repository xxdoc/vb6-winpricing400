VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAR021"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAR009"
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
Private m_SupplierTypes As Collection
Private m_PaidAmounts As Collection
Private m_MaxMins As Collection
Private m_ApSupAmounts As Collection
Private m_ApSupPaids As Collection
Private m_BillTotalPrices As Collection
Private m_DnItemsByBill As Collection
Private m_CnItemsByBill As Collection
Private m_RtItemsByBill As Collection
Private m_BillingDiscounts As Collection
Private m_CashTranAmounts As Collection

Private m_PaidAmountByDocDate As Collection
Private m_BalanceDueDate As Collection
Private m_BalanceTotal As Collection
Private m_BalanceTotalLastDue As Collection

Private Hd As Long

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Headers1_1 = New CFieldList
   Set m_Headers1_2 = New CFieldList
   Set m_Headers1_3 = New CFieldList
   Set m_Headers1_4 = New CFieldList
   Set m_Details1 = New CFieldList
   Set m_Details2 = New CFieldList
   
   Set m_SupplierTypes = New Collection
   Set m_PaidAmounts = New Collection
   Set m_MaxMins = New Collection
   Set m_ApSupAmounts = New Collection
   Set m_ApSupPaids = New Collection
   Set m_BillTotalPrices = New Collection
   Set m_DnItemsByBill = New Collection
   Set m_CnItemsByBill = New Collection
   Set m_RtItemsByBill = New Collection
   Set m_BillingDiscounts = New Collection
   Set m_CashTranAmounts = New Collection
   
   Set m_PaidAmountByDocDate = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
   
   Set m_SupplierTypes = Nothing
   Set m_PaidAmounts = Nothing
   Set m_MaxMins = Nothing
   Set m_BillTotalPrices = Nothing
   Set m_DnItemsByBill = Nothing
   Set m_CnItemsByBill = Nothing
   Set m_RtItemsByBill = Nothing
   Set m_BillingDiscounts = Nothing
   Set m_CashTranAmounts = Nothing
   
   Set m_PaidAmountByDocDate = Nothing
   Set m_BalanceDueDate = Nothing
   Set m_BalanceTotal = Nothing
   Set m_BalanceTotalLastDue = Nothing
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
Dim PrevID As Long
Dim TempStr1 As String
Dim TempStr1_1 As String
Dim Cf As CReportField
Dim ExportWidth As Long
Dim HeadCf As CReportField
Dim J As Long
Dim TempStr As String
Dim TempRs As ADODB.Recordset
Dim Amt As Double
Dim iCount As Long
Dim TempDate As Date
Dim TempWidth As Long

   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField

   Set Cf = New CReportField

   Call Cf.SetFieldValue(2, "^", "No", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "รหัส", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(9, "^", "ชื่อ", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "Cr", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(5, "^", "วงเงิน", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(5, "^", "หนี้", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(5, "^", "DUE", ">")
   Call m_Headers0.AddField(Cf)


   TempDate = mcolParam("FROM_DUE_DATE")
   While TempDate <= mcolParam("TO_DUE_DATE")
      TempWidth = 0
      Call Cf.SetFieldValue(6, "^", DateToStringExtEx2(TempDate), ">")
      Call m_Headers0.AddField(Cf)
      TempWidth = TempWidth + 1
      
      TempDate = DateAdd("D", 1, TempDate)
   Wend
   
   Call Cf.SetFieldValue(5, "^", "สะสม", ">")
   Call m_Headers0.AddField(Cf)
   
   If mcolParam("COMPARE_TARGET_ACTUAL_FLAG") = 1 Then
      Call Cf.SetFieldValue(5, "^", "สะสม", ">")
      Call m_Headers0.AddField(Cf)
      Call Cf.SetFieldValue(5, "^", "(ขาด)", ">")
      Call m_Headers0.AddField(Cf)
   End If
'      ===
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "ลูกค้า", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(9, "^", "ลูกค้า", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(5, "^", "", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(5, "^", "คงเหลือ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(5, "^", "ยกมา", ">")
   Call m_Headers1.AddField(Cf)
   
   TempDate = mcolParam("FROM_DUE_DATE")
   While TempDate <= mcolParam("TO_DUE_DATE")
      If mcolParam("COMPARE_TARGET_ACTUAL_FLAG") = 1 Then
         Call Cf.SetFieldValue(3, "^", "TG", ">")
         Call m_Headers1.AddField(Cf)
         
         Call Cf.SetFieldValue(3, "^", "AT", ">")
         Call m_Headers1.AddField(Cf)
      Else
         Call Cf.SetFieldValue(6, "^", "TG", ">")
         Call m_Headers1.AddField(Cf)
      End If
      TempDate = DateAdd("D", 1, TempDate)
   Wend
   
   Call Cf.SetFieldValue(5, "^", "TG", ">")
   Call m_Headers1.AddField(Cf)
   
   If mcolParam("COMPARE_TARGET_ACTUAL_FLAG") = 1 Then
      Call Cf.SetFieldValue(5, "^", "AT", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(5, "^", "เกิน", ">")
      Call m_Headers1.AddField(Cf)
   End If
   
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
Private Sub printHeader2()
Dim strFormat As String
Dim strPrint As String
Dim tbdT As TableBorderSettings
Dim blnBold As Boolean
Dim iSize As Integer
Dim alngX() As Long
Dim PrevID As Long
Dim TempStr1 As String
Dim TempStr1_1 As String
Dim Cf As CReportField
Dim ExportWidth As Long
Dim HeadCf As CReportField
Dim J As Long
Dim TempStr As String
Dim TempRs As ADODB.Recordset
Dim Amt As Double
Dim iCount As Long
Dim TempDate As Date
Dim TempWidth As Long

   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField

   Set Cf = New CReportField

   Call Cf.SetFieldValue(2, "^", "No", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "รหัส", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(9, "^", "ชื่อ", "<")
   Call m_Headers0.AddField(Cf)

   TempDate = mcolParam("FROM_DUE_DATE")
   While TempDate <= mcolParam("TO_DUE_DATE")
      TempWidth = 0
      Call Cf.SetFieldValue(6, "^", DateToStringExtEx2(TempDate), ">")
      Call m_Headers0.AddField(Cf)
      TempWidth = TempWidth + 1
      
      TempDate = DateAdd("D", 1, TempDate)
   Wend
   

'      ===
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "ลูกค้า", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(9, "^", "ลูกค้า", "<")
   Call m_Headers1.AddField(Cf)
   
   TempDate = mcolParam("FROM_DUE_DATE")
   While TempDate <= mcolParam("TO_DUE_DATE")
      If mcolParam("COMPARE_TARGET_ACTUAL_FLAG") = 1 Then
         Call Cf.SetFieldValue(6, "^", "AT", ">")
         Call m_Headers1.AddField(Cf)

      End If
      TempDate = DateAdd("D", 1, TempDate)
   Wend
      
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
Private Function initDoc() As Boolean
Dim strDate As String
Dim TempStr1 As String
Dim TempStr2 As String
Dim TempStr3 As String
Dim FromDate As String
Dim ToDate As String
Dim ToPaidDate As String
Dim NewDate As Date
Dim NewFromDate As Date
Dim NewToDate As Date

   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
   Vsp.ORIENTATION = orLandscape
   Vsp.MarginBottom = 700
   Vsp.MarginFooter = 700
   Vsp.MarginHeader = 700
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 700
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
  
   Call LoadPaidAmountByBill(Nothing, m_PaidAmounts, mcolParam("FROM_DOC_DATE"), mcolParam("TO_PAY_DATE"))
   Call LoadTotalPriceByBill(Nothing, m_BillTotalPrices, mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"))
   Call LoadBillingDiscountByBill(Nothing, m_BillingDiscounts, mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"))
'   Call LoadBillingFreeAmountByBill(Nothing, m_BillingDiscounts, mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"), "I")
   Call LoadTotalPriceByCustomer(Nothing, m_ApSupAmounts, mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"), 1)
   Call LoadPaidAmountByCustomer(Nothing, m_ApSupPaids, mcolParam("FROM_DOC_DATE"), mcolParam("TO_PAY_DATE"))
   Call LoadDnCnAmountByBill(Nothing, m_DnItemsByBill, mcolParam("FROM_DOC_DATE"), mcolParam("TO_PAY_DATE"), 4, 2)
   Call LoadDnCnAmountByBill(Nothing, m_CnItemsByBill, mcolParam("FROM_DOC_DATE"), mcolParam("TO_PAY_DATE"), 3, 2)
   Call LoadDnCnAmountByBill(Nothing, m_RtItemsByBill, mcolParam("FROM_DOC_DATE"), mcolParam("TO_PAY_DATE"), 18, 2)

    Call LoadPaidAmountByCustomerDocDate(Nothing, m_PaidAmountByDocDate, mcolParam("FROM_DUE_DATE"), mcolParam("TO_DUE_DATE"))
   Call LoadCashTrnAmountByCustomerDocDate(Nothing, m_CashTranAmounts, mcolParam("FROM_DUE_DATE"), mcolParam("TO_DUE_DATE"))
    
   
   FromDate = "จากวันที่บิล " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DUE_DATE")), "N/A")
   ToDate = " ถึง วันที่บิล " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DUE_DATE")), "N/A")
   
   MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & vbCrLf & mcolParam("REPORT_NAME") & _
                                       vbCrLf & FromDate & ToDate & "  " & ToPaidDate
   
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
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(100) As Double
Dim Total2(100) As Double
Dim Total3(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim Width As Double
Dim TempRs As ADODB.Recordset
Dim Ap As CCustomer
Dim Hc As CReportField
Dim Pi1 As CPaymentItem
Dim CashAmount As Double
Dim CheckAmount As Double
Dim TrnAmount As Double
Dim FeeAmount As Double
Dim ReceiptAmount As Double
Dim Bnk As CBank
Dim Bbr As CBankBranch
Dim TempDate As Date
Dim Dept As Double
Dim DeptBalance As Double
Dim Ri1_0 As CReceiptItem
Dim Ri1_1 As CReceiptItem
Dim Ri1_2 As CReceiptItem
Dim Bdc As CBillingDiscount

Dim TempApt As CDoItem
Dim Apt As CBillingDoc
Dim TempApRcIt As CReceiptItem
Dim TempPaidAmountByDocDate As CReceiptItem
Dim TempCashTrnAmountByDocDate As CCashTran
Dim temp_Apt As CBillingDoc
Dim Cal As Double
Dim SUM_TARGET  As Double
Dim SUM_ACTUAL  As Double
Dim temp_Actual As Double
Dim count_Actual As Long
Dim isRun As Boolean
Dim isRun2 As Boolean

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   'เช็คว่าเรียนได้ไม่เกิน 7 วัน
   If DateDiff("D", mcolParam("FROM_DUE_DATE"), mcolParam("TO_DUE_DATE")) > 7 Then
      mstrErrMsg = "ไม่สามารถใส่ช่วงวันที่ เกิน 7 วันได้"
      Exit Function
   End If
   
   Set Rs = New ADODB.Recordset
   Set TempRs = New ADODB.Recordset
   
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
   
   For J = 1 To UBound(Total1)
      Total1(J) = 0
      Total2(J) = 0
      Total3(J) = 0
   Next J
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
   
       Set BodyCf = New CReportField
      Set Ap = New CCustomer
      Ap.CUSTOMER_CODE = mcolParam("CUSTOMER_CODE")
      
      Ap.FROM_CUSTOMER_CODE = mcolParam("FROM_CUSTOMER_CODE")
      Ap.TO_CUSTOMER_CODE = mcolParam("TO_CUSTOMER_CODE")
      Ap.CUSTOMER_TYPE = mcolParam("CUSTOMER_TYPE")
      Ap.CUSTOMER_GRADE = mcolParam("CUSTOMER_GRADE")
      
      Ap.OrderType = mcolParam("ORDER_TYPE")
      Ap.OrderBy = mcolParam("ORDER_BY")

      Call GenerateBill
      
      Call glbDaily.QueryCustomer(Ap, Rs, iCount, IsOK, glbErrorLog)

   
      J = 0
      While Not Rs.EOF
         I = 0
         temp_Actual = 0
         
         Call m_Details1.ClearField
         Call Ap.PopulateFromRS(1, Rs)
         Cal = 1
         
            If mcolParam("BALANCE_SUMMIT_FLAG") = 1 Then
               Set temp_Apt = GetObject("CBillingDoc", m_BalanceTotal, Trim(str(Ap.CUSTOMER_ID)), False)
               If Not temp_Apt Is Nothing Then
                   Cal = temp_Apt.TOTAL_BALANCE_AMOUNT
               Else
                  Cal = 0
               End If
            End If
            
            If ROUND(Cal, 2) > 0 Then
               isRun = True
            Else
               isRun = False
            End If
            
          If mcolParam("BALANCE_ACTUAL_FLAG") = 1 And isRun Then
             FeeAmount = 0
             ReceiptAmount = 0
             temp_Actual = 0
             isRun2 = False
            TempDate = mcolParam("FROM_DUE_DATE")
               While TempDate <= mcolParam("TO_DUE_DATE") And Not isRun2
                   Set TempCashTrnAmountByDocDate = GetObject("CReceiptItem", m_CashTranAmounts, Trim(str(Ap.CUSTOMER_ID)) & "-" & Trim(TempDate), False)
                   If Not TempCashTrnAmountByDocDate Is Nothing Then
                     FeeAmount = TempCashTrnAmountByDocDate.GetFieldValue("FEE_AMOUNT")
                     ReceiptAmount = TempCashTrnAmountByDocDate.GetFieldValue("AMOUNT")
                   End If
                   
                  Set TempPaidAmountByDocDate = GetObject("CReceiptItem", m_PaidAmountByDocDate, Trim(str(Ap.CUSTOMER_ID)) & "-" & Trim(TempDate), False)
                  If Not TempPaidAmountByDocDate Is Nothing Then
                     temp_Actual = ROUND(ReceiptAmount - TempPaidAmountByDocDate.DISCOUNT_AMOUNT - FeeAmount, 2)
                  End If
                  
                     If temp_Actual > 0 Then
                         isRun = True
                     Else
                        isRun = False
                     End If
                     isRun2 = isRun
                     TempDate = DateAdd("D", 1, TempDate)
               Wend
         End If
           
           
           If isRun Then
               I = I + 1
               J = J + 1
               
                Set Hc = m_Headers1.Fields(I)
               TempStr = J
               Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
      
                I = I + 1
                Set Hc = m_Headers1.Fields(I)
               TempStr = Ap.CUSTOMER_CODE
               Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
                I = I + 1
                Set Hc = m_Headers1.Fields(I)
               TempStr = Ap.CUSTOMER_NAME
               Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
                I = I + 1
                Set Hc = m_Headers1.Fields(I) 'เครดิต
               TempStr = Ap.MAX_CREDIT
               Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
                I = I + 1
                Set Hc = m_Headers1.Fields(I) 'วงเงิน
               TempStr = FormatNumber(Ap.CREDIT_LIMIT)
               Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Total1(I) = Total1(I) + Ap.CREDIT_LIMIT
               Total2(I) = Total2(I) + Ap.CREDIT_LIMIT
               
                 I = I + 1
                Set Hc = m_Headers1.Fields(I) 'หนี้คงเหลือ
               Set temp_Apt = GetObject("CBillingDoc", m_BalanceTotal, Trim(str(Ap.CUSTOMER_ID)), False)
               If Not temp_Apt Is Nothing Then
                   TempStr = IIf(ROUND(temp_Apt.TOTAL_BALANCE_AMOUNT, 2) = 0, "", FormatNumber(ROUND(temp_Apt.TOTAL_BALANCE_AMOUNT, 2)))
                   Total1(I) = Total1(I) + ROUND(temp_Apt.TOTAL_BALANCE_AMOUNT, 2)
                   Total2(I) = Total2(I) + ROUND(temp_Apt.TOTAL_BALANCE_AMOUNT, 2)
               Else
                  TempStr = ""
               End If
               Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
                I = I + 1
                Set Hc = m_Headers1.Fields(I) 'ครบกำหนดยกมา
               Set temp_Apt = GetObject("CBillingDoc", m_BalanceTotalLastDue, Trim(str(Ap.CUSTOMER_ID)), False)
               If Not temp_Apt Is Nothing Then
                  TempStr = IIf(ROUND(temp_Apt.TOTAL_BALANCE_AMOUNT_LAST_DUE, 2) = 0, "", FormatNumber(ROUND(temp_Apt.TOTAL_BALANCE_AMOUNT_LAST_DUE, 2)))
                   Total1(I) = Total1(I) + ROUND(temp_Apt.TOTAL_BALANCE_AMOUNT_LAST_DUE, 2)
                   Total2(I) = Total2(I) + ROUND(temp_Apt.TOTAL_BALANCE_AMOUNT_LAST_DUE, 2)
               Else
                  TempStr = ""
               End If
               Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
            temp_Actual = 0
            TempDate = mcolParam("FROM_DUE_DATE")
            While TempDate <= mcolParam("TO_DUE_DATE")
               If mcolParam("COMPARE_TARGET_ACTUAL_FLAG") = 1 Then
                   I = I + 1
                   Set Hc = m_Headers1.Fields(I) 'TARGET
                  Set temp_Apt = GetObject("CBillingDoc", m_BalanceDueDate, Trim(str(Ap.CUSTOMER_ID)) & "-" & Trim(TempDate), False)
                  If Not temp_Apt Is Nothing Then
                     TempStr = IIf(ROUND(temp_Apt.DUE_AMOUNT, 2) = 0, "", FormatNumber(ROUND(temp_Apt.DUE_AMOUNT, 2)))
                      Total1(I) = Total1(I) + ROUND(temp_Apt.DUE_AMOUNT, 2)
                     Total2(I) = Total2(I) + ROUND(temp_Apt.DUE_AMOUNT, 2)
                  Else
                     TempStr = ""
                  End If
                  Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
                    
                   I = I + 1
                   
                   FeeAmount = 0
                   ReceiptAmount = 0
                   Set TempCashTrnAmountByDocDate = GetObject("CReceiptItem", m_CashTranAmounts, Trim(str(Ap.CUSTOMER_ID)) & "-" & Trim(TempDate), False)
                   If Not TempCashTrnAmountByDocDate Is Nothing Then
                     FeeAmount = TempCashTrnAmountByDocDate.GetFieldValue("FEE_AMOUNT")
                     ReceiptAmount = TempCashTrnAmountByDocDate.GetFieldValue("AMOUNT")
                   End If
                   
                   Set TempPaidAmountByDocDate = GetObject("CReceiptItem", m_PaidAmountByDocDate, Trim(str(Ap.CUSTOMER_ID)) & "-" & Trim(TempDate), False)
                  If Not TempPaidAmountByDocDate Is Nothing Then
'                     TempStr = IIf(ROUND(TempPaidAmountByDocDate.PAID_AMOUNT - TempPaidAmountByDocDate.DISCOUNT_AMOUNT - FeeAmount, 2) = 0, "", FormatNumber(ROUND(TempPaidAmountByDocDate.PAID_AMOUNT - TempPaidAmountByDocDate.DISCOUNT_AMOUNT - FeeAmount, 2)))
'                      Total1(I) = Total1(I) + ROUND(TempPaidAmountByDocDate.PAID_AMOUNT - TempPaidAmountByDocDate.DISCOUNT_AMOUNT - FeeAmount, 2)
'                      Total2(I) = Total2(I) + ROUND(TempPaidAmountByDocDate.PAID_AMOUNT - TempPaidAmountByDocDate.DISCOUNT_AMOUNT - FeeAmount, 2)
'                      temp_Actual = ROUND(TempPaidAmountByDocDate.PAID_AMOUNT - TempPaidAmountByDocDate.DISCOUNT_AMOUNT - FeeAmount, 2)
                     TempStr = IIf(ROUND(ReceiptAmount - TempPaidAmountByDocDate.DISCOUNT_AMOUNT - FeeAmount, 2) = 0, "", FormatNumber(ROUND(ReceiptAmount - TempPaidAmountByDocDate.DISCOUNT_AMOUNT - FeeAmount, 2)))
                      Total1(I) = Total1(I) + ROUND(ReceiptAmount - TempPaidAmountByDocDate.DISCOUNT_AMOUNT - FeeAmount, 2)
                      Total2(I) = Total2(I) + ROUND(ReceiptAmount - TempPaidAmountByDocDate.DISCOUNT_AMOUNT - FeeAmount, 2)
'                      temp_Actual = ROUND(TempPaidAmountByDocDate.TOTAL_RCP - TempPaidAmountByDocDate.DISCOUNT_AMOUNT - FeeAmount, 2)
                   Else
                     TempStr = ""
                  End If
                   Set Hc = m_Headers1.Fields(I) 'ACTUAL
                  Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               Else
                   I = I + 1
                   Set Hc = m_Headers1.Fields(I) 'TARGET
                  Set temp_Apt = GetObject("CBillingDoc", m_BalanceDueDate, Trim(str(Ap.CUSTOMER_ID)) & "-" & Trim(TempDate), False)
                  If Not temp_Apt Is Nothing Then
                     TempStr = IIf(ROUND(temp_Apt.DUE_AMOUNT, 2) = 0, "", FormatNumber(ROUND(temp_Apt.DUE_AMOUNT, 2)))
                      Total1(I) = Total1(I) + ROUND(temp_Apt.DUE_AMOUNT, 2)
                     Total2(I) = Total2(I) + ROUND(temp_Apt.DUE_AMOUNT, 2)
                  Else
                     TempStr = ""
                  End If
                  Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               TempDate = DateAdd("D", 1, TempDate)
            Wend
            
            'คำนวนหา ยอดรวม Sum_Actual,Sum_Target
               SUM_TARGET = 0
               SUM_ACTUAL = 0
               TempDate = mcolParam("FROM_DUE_DATE")
               While TempDate <= mcolParam("TO_DUE_DATE")
                  Set temp_Apt = GetObject("CBillingDoc", m_BalanceDueDate, Trim(str(Ap.CUSTOMER_ID)) & "-" & Trim(TempDate), False)
                  If Not temp_Apt Is Nothing Then
                     SUM_TARGET = SUM_TARGET + temp_Apt.DUE_AMOUNT
                  End If
                  
                   FeeAmount = 0
                   ReceiptAmount = 0
                   Set TempCashTrnAmountByDocDate = GetObject("CReceiptItem", m_CashTranAmounts, Trim(str(Ap.CUSTOMER_ID)) & "-" & Trim(TempDate), False)
                   If Not TempCashTrnAmountByDocDate Is Nothing Then
                     FeeAmount = TempCashTrnAmountByDocDate.GetFieldValue("FEE_AMOUNT")
                     ReceiptAmount = TempCashTrnAmountByDocDate.GetFieldValue("AMOUNT")
                   End If
                   
                   Set TempPaidAmountByDocDate = GetObject("CReceiptItem", m_PaidAmountByDocDate, Trim(str(Ap.CUSTOMER_ID)) & "-" & Trim(TempDate), False)
                  If Not TempPaidAmountByDocDate Is Nothing Then
                     SUM_ACTUAL = SUM_ACTUAL + (ReceiptAmount - TempPaidAmountByDocDate.DISCOUNT_AMOUNT - FeeAmount)
                  End If
               TempDate = DateAdd("D", 1, TempDate)
               Wend
               
               
                I = I + 1
                Set Hc = m_Headers1.Fields(I) 'สะสม TARGET
               TempStr = FormatNumber(SUM_TARGET)
               Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Total1(I) = Total1(I) + SUM_TARGET
               Total2(I) = Total2(I) + SUM_TARGET
               
               If mcolParam("COMPARE_TARGET_ACTUAL_FLAG") = 1 Then
                   I = I + 1
                   Set Hc = m_Headers1.Fields(I) 'สะสม ACTUAL
                  TempStr = FormatNumber(SUM_ACTUAL)
                  Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
                   Total1(I) = Total1(I) + SUM_ACTUAL
                   Total2(I) = Total2(I) + SUM_ACTUAL
                  
                    I = I + 1
                   Set Hc = m_Headers1.Fields(I) 'ขาด เกิน
                   If Val(SUM_ACTUAL - SUM_TARGET) < 0 Then
                     TempStr = "(" & FormatNumber(Abs(SUM_ACTUAL - SUM_TARGET)) & ")"
                  Else
                     TempStr = FormatNumber(SUM_ACTUAL - SUM_TARGET)
                  End If
                  Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
                  Total1(I) = Total1(I) + SUM_TARGET - SUM_ACTUAL
                   Total2(I) = Total2(I) + SUM_TARGET - SUM_ACTUAL
               End If
      
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
           
            End If ' If isRun Then
         Rs.MoveNext
      Wend

      Call GenerateFooter("", "", "รวมขายเชื่อ", Total1, m_Details1, 0)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
'
      Call GenerateFooter("", "", "รวมสุทธิ", Total2, m_Details1, 0)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      
''      If mcolParam("DIFF_ACTUAL_FLAG") = 1 Then
''      Hd = 2
''      Vsp = ""
''      Vsp = "สรุปส่วนต่างยอดรับ ACTUAL"
''      Call printHeader2
''        Rs.MoveFirst
''
''        J = 0
''      While Not Rs.EOF
''         I = 0
''         Call m_Details1.ClearField
''         Call Ap.PopulateFromRS(1, Rs)
''         Cal = 1
''
'''               If Trim(Ap.CUSTOMER_CODE) = "ง-0016" Then
'''      Debug.Print
'''      End If
''            If mcolParam("BALANCE_SUMMIT_FLAG") = 1 Then
''               Set temp_Apt = GetObject("CBillingDoc", m_BalanceTotal, Trim(str(Ap.CUSTOMER_ID)), False)
''               If Not temp_Apt Is Nothing Then
''                   Cal = temp_Apt.TOTAL_BALANCE_AMOUNT
''               Else
''                  Cal = 0
''               End If
''            End If
''            If ROUND(Cal, 2) > 0 Then
''               I = I + 1
''               J = J + 1
''
''                Set Hc = m_Headers1.Fields(I)
''               TempStr = J
''               Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
''               Call m_Details1.AddField(BodyCf)
''
''
''                I = I + 1
''                Set Hc = m_Headers1.Fields(I)
''               TempStr = Ap.CUSTOMER_CODE
''               Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
''               Call m_Details1.AddField(BodyCf)
''
''                I = I + 1
''                Set Hc = m_Headers1.Fields(I)
''               TempStr = Ap.CUSTOMER_NAME
''               Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
''               Call m_Details1.AddField(BodyCf)
''
''            temp_Actual = 0
''            TempDate = mcolParam("FROM_DUE_DATE")
''            While TempDate <= mcolParam("TO_DUE_DATE")
'''               If mcolParam("COMPARE_TARGET_ACTUAL_FLAG") = 1 Then
''                      I = I + 1
''                      FeeAmount = 0
''                      ReceiptAmount = 0
''                      Set TempCashTrnAmountByDocDate = GetObject("CReceiptItem", m_CashTranAmounts, Trim(str(Ap.CUSTOMER_ID)) & "-" & Trim(TempDate), False)
''                      If Not TempCashTrnAmountByDocDate Is Nothing Then
''                        FeeAmount = TempCashTrnAmountByDocDate.GetFieldValue("FEE_AMOUNT")
''                        ReceiptAmount = TempCashTrnAmountByDocDate.GetFieldValue("AMOUNT")
''                      End If
''
''                      Set TempPaidAmountByDocDate = GetObject("CReceiptItem", m_PaidAmountByDocDate, Trim(str(Ap.CUSTOMER_ID)) & "-" & Trim(TempDate), False)
''                     If Not TempPaidAmountByDocDate Is Nothing Then
''                        TempStr = IIf(ROUND(ReceiptAmount, 2) = 0, "", FormatNumber(ROUND(ReceiptAmount, 2)))
''                         If temp_Actual = 0 Then 'temp_Actual มีเพียงแค่วันเดียว ก็ไม่ต้องเข้าหาอีกแล้ว
''                           temp_Actual = ROUND(ReceiptAmount, 2)
''                         End If
''                      Else
''                        TempStr = ""
''                     End If
''                      Set Hc = m_Headers1.Fields(I) 'ACTUAL
''                     Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
''                     Call m_Details1.AddField(BodyCf)
'''                  Else
'''                     TempStr = ""
'''                  End If
'''                  Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
'''                  Call m_Details1.AddField(BodyCf)
''
''               TempDate = DateAdd("D", 1, TempDate)
''            Wend
''
''               If temp_Actual <> 0 Then
''                 Call m_Details1.GetString(1, TempStr1, TempStr2)
''                 strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
''                 Call Vsp.AddTable(strFormat, "", TempStr2)
''               End If
''            End If 'if val(Cal)>0
''         Rs.MoveNext
''      Wend
''
'''      Call GenerateFooter("", "", "รวมขายเชื่อ", Total1, m_Details1, 0)
'''      Call m_Details1.GetString(2, TempStr1, TempStr2)
'''      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'''      Vsp.FontBold = True
'''      Call Vsp.AddTable(strFormat, "", TempStr2)
'''      Vsp.FontBold = False
''''
'''      Call GenerateFooter("", "", "รวมสุทธิ", Total2, m_Details1, 0)
'''      Call m_Details1.GetString(2, TempStr1, TempStr2)
'''      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'''      Vsp.FontBold = True
'''      Call Vsp.AddTable(strFormat, "", TempStr2)
'''      Vsp.FontBold = False
''      End If
      
      Set BodyCf = Nothing
    Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Set m_SupplierTypes = Nothing
   Set m_PaidAmounts = Nothing
   Set m_MaxMins = Nothing
   Set m_BillTotalPrices = Nothing
   Set m_DnItemsByBill = Nothing
   Set m_CnItemsByBill = Nothing
   Set m_RtItemsByBill = Nothing
   Set m_BillingDiscounts = Nothing
   
   Set m_PaidAmountByDocDate = Nothing
   Set m_BalanceDueDate = Nothing
   Set m_BalanceTotal = Nothing
   Set m_BalanceTotalLastDue = Nothing
   Set m_CashTranAmounts = Nothing
   
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
      Else
         TempStr = FormatNumber(Tot(J))
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
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

   If Not mblnNewPage Then Exit Sub
   talnT = Vsp.TextAlign
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   blnUnder = Vsp.FontUnderline
   blnItalic = Vsp.FontItalic
   iSize = Vsp.FontSize
   sName = Vsp.FontName
   
   Vsp.FontSize = 8
   dY(0) = Vsp.CurrentY
   Vsp.TextAlign = taLeftBottom
   strFormat = VSP_CalTable(TITLE_FORMAT, mdblWid, alngX)
   Vsp.TableBorder = tbNone
   Call Vsp.AddTable(strFormat, "", "|หน้าที่|: " & Vsp.CurrentPage)
   Call Vsp.AddTable(strFormat, "", "|วัน|: " & DateToStringExt(mdteDate))
   Call Vsp.AddTable(strFormat, "", "|เวลา|: " & Right$(DateToStringExtEx(Now), 8))
   Vsp.CurrentY = dY(0)
   
   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   Vsp.TextAlign = taCenterBottom
   Vsp.Paragraph = MODULE_DESC
   Vsp.Paragraph = ""
   If mblnHeader Then
      If Hd = 2 Then
         Call printHeader2
      Else
         Call printHeader
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
End Sub

Private Sub VSP_StartDoc()
'This event occur when VSPrinter.StartDoc is used and used to initialize some information before generating printed document
End Sub

Private Sub GenerateBill()
Dim Ari As CReceiptItem
Dim DueCount As Long
Dim MM As CMaxMin
Dim J As Long
Dim iCount As Long
Dim IsOK As Boolean
Dim Sum As Double
Dim PrintDate As Date
Dim Di As CDoItem
Dim Ri1_1 As CReceiptItem
Dim Ri1_2 As CReceiptItem
Dim Ri1_3 As CReceiptItem
Dim Bdc As CBillingDiscount
Dim Cash_Trn  As CCashTran
Dim NetAmt As Double
Dim toa As Double
Dim Apt As CBillingDoc
Dim AptN As CBillingDoc
Dim AptTotalBalance As CBillingDoc
Dim temp_Apt As CBillingDoc

Set m_BalanceDueDate = New Collection
Set m_BalanceTotal = New Collection
Set m_BalanceTotalLastDue = New Collection


      Set Apt = New CBillingDoc
      Apt.CUSTOMER_CODE = mcolParam("CUSTOMER_CODE")
      Apt.FROM_CUSTOMER_CODE = mcolParam("FROM_CUSTOMER_CODE")
      Apt.TO_CUSTOMER_CODE = mcolParam("TO_CUSTOMER_CODE")
      Apt.FROM_DATE = mcolParam("FROM_DOC_DATE")
      Apt.TO_DATE = mcolParam("TO_DOC_DATE")
      Apt.DOCUMENT_TYPE = 1
      Apt.CUSTOMER_TYPE = mcolParam("CUSTOMER_TYPE")
      Apt.CUSTOMER_GRADE = mcolParam("CUSTOMER_GRADE")
      Apt.OrderBy = mcolParam("ORDER_BY")
      Apt.OrderType = 1
      Apt.ItemSumFlag = True
      Call Apt.QueryData(3, Rs, iCount)

      While Not Rs.EOF
          Set Apt = Nothing
          Set Apt = New CBillingDoc
         Call Apt.PopulateFromRS(3, Rs)
         Set Di = GetDoItem(m_BillTotalPrices, Apt.BILLING_DOC_ID)
         Set Bdc = GetBillingDiscount(m_BillingDiscounts, Apt.BILLING_DOC_ID)
         
         DueCount = DateDiff("D", Apt.DUE_DATE, PrintDate)
         Set Ari = GetReceiptItem(m_PaidAmounts, Apt.BILLING_DOC_ID)
         Set Ri1_1 = GetReceiptItem(m_DnItemsByBill, Apt.BILLING_DOC_ID)
         Set Ri1_2 = GetReceiptItem(m_CnItemsByBill, Apt.BILLING_DOC_ID)
         Set Ri1_3 = GetReceiptItem(m_RtItemsByBill, Apt.BILLING_DOC_ID)
       
         'ปัดก่อนแล้วค่อยรวม
         Di.TOTAL_PRICE = Val(Format(Di.TOTAL_PRICE, "0.00"))
         Bdc.DISCOUNT_AMOUNT = Val(Format(Bdc.DISCOUNT_AMOUNT, "0.00"))
         Ri1_1.DEBIT_CREDIT_AMOUNT = Val(Format(Ri1_1.DEBIT_CREDIT_AMOUNT, "0.00"))
         Ri1_2.DEBIT_CREDIT_AMOUNT = Val(Format(Ri1_2.DEBIT_CREDIT_AMOUNT, "0.00"))
         Ri1_3.DEBIT_CREDIT_AMOUNT = Val(Format(Ri1_3.DEBIT_CREDIT_AMOUNT, "0.00"))
         Ari.CASH_DISCOUNT = Val(Format(Ari.CASH_DISCOUNT, "0.00"))
   
         NetAmt = (Di.TOTAL_PRICE - Bdc.DISCOUNT_AMOUNT) + (Ri1_1.DEBIT_CREDIT_AMOUNT - Ri1_2.DEBIT_CREDIT_AMOUNT - Ri1_3.DEBIT_CREDIT_AMOUNT - Ari.CASH_DISCOUNT) - Ari.PAID_AMOUNT
   
         Set temp_Apt = GetObject("CBillingDoc", m_BalanceDueDate, Trim(str(Apt.CUSTOMER_ID)) & "-" & Trim(Apt.DUE_DATE), False) 'หายอดหนี้ Target ของลูกค้าแต่ละคนแต่ละวันดิล
         If Not temp_Apt Is Nothing Then
            temp_Apt.DUE_AMOUNT = temp_Apt.DUE_AMOUNT + NetAmt
         Else
           Apt.DUE_AMOUNT = NetAmt
            Call m_BalanceDueDate.add(Apt, Trim(str(Apt.CUSTOMER_ID)) & "-" & Trim(Apt.DUE_DATE))
         End If
         
         Set temp_Apt = GetObject("CBillingDoc", m_BalanceTotal, Trim(str(Apt.CUSTOMER_ID)), False) 'หายอดหนี้คงเหลือทั้งหมด
         If Not temp_Apt Is Nothing Then
            temp_Apt.TOTAL_BALANCE_AMOUNT = temp_Apt.TOTAL_BALANCE_AMOUNT + NetAmt
         Else
           Apt.TOTAL_BALANCE_AMOUNT = NetAmt
            Call m_BalanceTotal.add(Apt, Trim(str(Apt.CUSTOMER_ID)))
         End If
         
          If Apt.DUE_DATE < mcolParam("FROM_DUE_DATE") Then
            Set temp_Apt = GetObject("CBillingDoc", m_BalanceTotalLastDue, Trim(str(Apt.CUSTOMER_ID)), False) 'หายอดหนี้คงเหลือที่เกินดิวมาจนถึงดิวที่เลือก
            If Not temp_Apt Is Nothing Then
               temp_Apt.TOTAL_BALANCE_AMOUNT_LAST_DUE = temp_Apt.TOTAL_BALANCE_AMOUNT_LAST_DUE + NetAmt
            Else
               Apt.TOTAL_BALANCE_AMOUNT_LAST_DUE = NetAmt
               Call m_BalanceTotalLastDue.add(Apt, Trim(str(Apt.CUSTOMER_ID)))
            End If
          End If
   
         Rs.MoveNext
      Wend
End Sub



