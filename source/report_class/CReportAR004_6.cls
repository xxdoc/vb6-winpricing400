VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAR004_6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAR004_6"
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
Private m_Headers0_1 As CFieldList
Private m_Headers0_2 As CFieldList
Private m_Headers1 As CFieldList
Private m_Headers1_1 As CFieldList
Private m_Headers1_2 As CFieldList
Private m_Headers1_3 As CFieldList
Private m_Headers1_4 As CFieldList
Private m_Details1 As CFieldList
Private m_Details2 As CFieldList
Private m_Left As Double
Private m_YearSeqs As Collection

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
Private m_AccountFollowCancelFlag_N  As Collection
Private m_MarketingFollowCancelFlag_N  As Collection




Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_TempCol = New Collection
  
   
   Set m_Headers0 = New CFieldList
   Set m_Headers0_1 = New CFieldList
   Set m_Headers0_2 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Headers1_1 = New CFieldList
   Set m_Headers1_2 = New CFieldList
   Set m_Headers1_3 = New CFieldList
   Set m_Headers1_4 = New CFieldList
   Set m_Details1 = New CFieldList
   Set m_Details2 = New CFieldList
   Set m_YearSeqs = New Collection
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
   Set m_AccountFollowCancelFlag_N = New Collection
   Set m_MarketingFollowCancelFlag_N = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers0_1 = Nothing
   Set m_Headers0_2 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_YearSeqs = Nothing
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
   Set m_AccountFollowCancelFlag_N = Nothing
   Set m_MarketingFollowCancelFlag_N = Nothing
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

   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   'Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers0_1.ClearField
   Call m_Headers0_2.ClearField
   
   Call m_Headers1.ClearField
   Call m_Headers1_1.ClearField
   Call m_Headers1_2.ClearField
   
   Set Cf = New CReportField
   
   Call Cf.SetFieldValue(22, "<", "พนักงานขาย", "^")
   Call m_Headers0_1.AddField(Cf)
         
    Call Cf.SetFieldValue(20, "^", "รหัส", "^")
    Call m_Headers0_2.AddField(Cf)
    Call Cf.SetFieldValue(50, "^", "ชื่อลูกค้า", "<")
    Call m_Headers0_2.AddField(Cf)
    If Check2Flag(mcolParam("SHOW_DETAIL")) = "Y" Then
      Call Cf.SetFieldValue(15, "^", "เครดิต", "^")
    Else
      Call Cf.SetFieldValue(15, "^", "เครดิต (วัน)", "^")
    End If
    Call m_Headers0_2.AddField(Cf)
    Call Cf.SetFieldValue(20, "^", "วงเงินอนุมัติ", ">")
    Call m_Headers0_2.AddField(Cf)
    Call Cf.SetFieldValue(20, "^", "หนี้คงเหลือ", ">")
    Call m_Headers0_2.AddField(Cf)
    Call Cf.SetFieldValue(20, "^", "วงเงินคงเหลือ", ">", , False)
    Call m_Headers0_2.AddField(Cf)
    If Check2Flag(mcolParam("SHOW_DETAIL")) = "Y" Then
      Call Cf.SetFieldValue(20, "^", "จำนวนเงิน", ">")
    Else
      Call Cf.SetFieldValue(20, "^", "จำนวนเงิน เกินวงวัน", ">")
    End If
    Call m_Headers0_2.AddField(Cf)
    If Check2Flag(mcolParam("SHOW_DETAIL")) = "Y" Then
      Call Cf.SetFieldValue(20, "^", "จำนวนเงิน", ">")
   Else
      Call Cf.SetFieldValue(20, "^", "จำนวนเงิน ในDUE", ">")
   End If
   
'   If Check2Flag(mcolParam("SHOW_DETAIL")) = "Y" Then
'    Call m_Headers0_2.AddField(Cf)
'    Call Cf.SetFieldValue(0, "^", "การติดตาม", "<", , False) 'การติดตามผลฝ่ายบัญชี
'    Call m_Headers0_2.AddField(Cf)
'     Call Cf.SetFieldValue(0, "^", "การติดตามผลฝ่ายขาย", "<", , False)
'    Call m_Headers0_2.AddField(Cf)
'   Else
     Call m_Headers0_2.AddField(Cf)
    Call Cf.SetFieldValue(50, "^", "การติดตาม", "<", , False) 'การติดตามผลฝ่ายบัญชี
    Call m_Headers0_2.AddField(Cf)
     Call Cf.SetFieldValue(0, "^", "การติดตามผลฝ่ายขาย", "<", , False)
    Call m_Headers0_2.AddField(Cf)
'   End If
    
    
    
    Call m_Headers0_2.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
    Call Cf.SetFieldValue(20, "^", "เลขที่บิล", "^")
    Call m_Headers1.AddField(Cf)
    Call Cf.SetFieldValue(25, "^", "วันที่เอกสาร", "<")
    Call m_Headers1.AddField(Cf)
    Call Cf.SetFieldValue(25, "^", "วันที่ครบกำหนด", "<")
    Call m_Headers1.AddField(Cf)
    Call Cf.SetFieldValue(15, "^", "เกิน(วัน)", "^")
    Call m_Headers1.AddField(Cf)
    Call Cf.SetFieldValue(20, "^", "ยอดเต็มบิล", ">")
    Call m_Headers1.AddField(Cf)
    Call Cf.SetFieldValue(20, "^", "ยอดคงเหลือ", ">")
    Call m_Headers1.AddField(Cf)
    Call Cf.SetFieldValue(20, "^", "", ">", , False)
    Call m_Headers1.AddField(Cf)
    Call Cf.SetFieldValue(20, "^", "เกินวงวัน", ">")
    Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(20, "^", "ในDUE", ">")
    Call m_Headers1.AddField(Cf)
    
'    If Check2Flag(mcolParam("SHOW_DETAIL")) = "Y" Then
'    Call Cf.SetFieldValue(0, "^", "", "<", , False)
'    Call m_Headers1.AddField(Cf)
'     Call Cf.SetFieldValue(0, "^", "", "<", , False)
'    Call m_Headers1.AddField(Cf)
'    Else
      Call Cf.SetFieldValue(50, "^", "", "<", , False)
    Call m_Headers1.AddField(Cf)
     Call Cf.SetFieldValue(0, "^", "", "<", , False)
    Call m_Headers1.AddField(Cf)
'    End If
    
    If Check2Flag(mcolParam("SHOW_DETAIL")) = "Y" Then
      Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      strPrint = TempStr1_1
      Call Vsp.AddTable(strFormat, "", strPrint)
   End If
   Set Cf = Nothing
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
   
   Call LoadTotalPriceByBill(Nothing, m_BillTotalPrices, mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"))
   Call LoadBillingDiscountByBill(Nothing, m_BillingDiscounts, mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"))
   Call LoadPaidAmountByBill(Nothing, m_PaidAmounts, mcolParam("FROM_DOC_DATE"), mcolParam("TO_PAY_DATE"))
   Call LoadDnCnAmountByBill(Nothing, m_DnItemsByBill, mcolParam("FROM_DOC_DATE"), mcolParam("TO_PAY_DATE"), 4, 2)
   Call LoadDnCnAmountByBill(Nothing, m_CnItemsByBill, mcolParam("FROM_DOC_DATE"), mcolParam("TO_PAY_DATE"), 3, 2)
   Call LoadDnCnAmountByBill(Nothing, m_RtItemsByBill, mcolParam("FROM_DOC_DATE"), mcolParam("TO_PAY_DATE"), 18, 2)
  
'   If mcolParam("INTERVAL_TYPE") = 1 Then
'  Call LoadDueDateInterval(Nothing, m_MaxMins, mcolParam("SUM_INDUE_FLAG"))
'   Else
'      Call LoadDueDateInterval2(Nothing, m_MaxMins, mcolParam("SUM_INDUE_FLAG"))
'   End If
   
   'CONCLUSION_FLAG
   'Call LoadTotalPriceByCustomer(Nothing, m_ApSupAmounts, mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"), 1)
   'Call LoadPaidAmountByCustomer(Nothing, m_ApSupPaids, mcolParam("FROM_DOC_DATE"), mcolParam("TO_PAY_DATE"))
   
  
  'Call GetAccountFollowCancelFlag_N(m_AccountFollowCancelFlag_N, mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"), 1)
  Call GetAccountFollow(m_AccountFollowCancelFlag_N)
  
  Call GetMKTFollow(m_MarketingFollowCancelFlag_N)

   If Check2Flag(mcolParam("OVERDUE_FLAG")) = "Y" Then
      TempStr1 = " เฉพาะเกินวงเงินหรือวงวัน"
   End If
   

   FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DOC_DATE")), "N/A")
   ToDate = " ถึง วันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DOC_DATE")), "N/A")
   ToPaidDate = "    วันที่ชำระเงินถึงวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_PAY_DATE")), "N/A") & "    วันที่พิมพ์ " & EmptyToString(DateToStringExtEx2(mcolParam("PRINT_DATE")), "N/A")
'   ToPaidDate = ToPaidDate & " เรียงตาม " & EmptyToString(mcolParam("ORDER_BY_NAME"), "N/A")

   MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & vbCrLf & mcolParam("REPORT_NAME") & _
                                       vbCrLf & FromDate & ToDate & "  " & ToPaidDate & TempStr1
   
   
   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))

   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
      mstrErrMsg = "Error GetDateTime Error."
      Exit Function
   End If
   mdteDate = InternalDateToDate(strDate)
   initDoc = True
End Function

Private Function GetItemFromListIndex(TempCol As Collection, Ind As Long) As Object
   Set GetItemFromListIndex = TempCol(Ind)
End Function

Private Sub GenerateDetail(alngX() As Long, I As Long, Apt As CBillingDoc, Tot1() As Double, Tot2() As Double, Indue As Double, Overdue As Double, NewCust As Boolean, HeaderExist As Boolean)
Dim HeadCf As CReportField
Dim TempStr As String
Dim TempStr1 As String
Dim TempStr2 As String
Dim BodyCf As CReportField
Dim strFormat As String
Dim Ari As CReceiptItem
Dim DueCount As Long
Dim MM As CMaxMin
Dim J As Long
Dim Sum As Double
Dim PrintDate As Date
Dim Di As CDoItem
Dim Ri1_1 As CReceiptItem
Dim Ri1_2 As CReceiptItem
Dim Ri1_3 As CReceiptItem
Dim Bdc As CBillingDiscount
Dim NetAmt As Double
Dim InCompleteFlag As Boolean
Dim Credit As Double


   Set BodyCf = New CReportField
   

   If mcolParam("PRINT_DATE") <= 0 Then
      PrintDate = Now
   Else
      PrintDate = mcolParam("PRINT_DATE")
   End If
   
   Set Di = GetDoItem(m_BillTotalPrices, Apt.BILLING_DOC_ID)
   Set Bdc = GetBillingDiscount(m_BillingDiscounts, Apt.BILLING_DOC_ID)
   
   DueCount = DateDiff("D", Apt.DUE_DATE, PrintDate)
   Set Ari = GetReceiptItem(m_PaidAmounts, Apt.BILLING_DOC_ID)
   Set Ri1_1 = GetReceiptItem(m_DnItemsByBill, Apt.BILLING_DOC_ID)
   Set Ri1_2 = GetReceiptItem(m_CnItemsByBill, Apt.BILLING_DOC_ID)
   Set Ri1_3 = GetReceiptItem(m_RtItemsByBill, Apt.BILLING_DOC_ID)
   
   InCompleteFlag = ROUND(((Di.TOTAL_PRICE - Bdc.DISCOUNT_AMOUNT) + (Ri1_1.DEBIT_CREDIT_AMOUNT - Ri1_2.DEBIT_CREDIT_AMOUNT - Ri1_3.DEBIT_CREDIT_AMOUNT) - (Ari.PAID_AMOUNT)), 2) <> 0
  
   If InCompleteFlag Then          'จ่ายยังไม่ครบ
      If NewCust Then 'ให้มั่นใจว่ามีอย่างน้อย 1 record ถึงค่อยสร้างหัว
          
'             Call GenerateHeader2(Apt, m_Details1)
        
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
         
         If Check2Flag(mcolParam("SHOW_DETAIL")) = "Y" Then
            Call GenerateHeader(Apt, m_Details1)
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
        End If
         Indue = 0
         Overdue = 0
         HeaderExist = True
         NewCust = False
         
      End If
            
      'If mcolParam("CONCLUSION_FLAG") = 1 Then
'         Call m_Details1.ClearField
'         Set HeadCf = m_Headers0_2.Fields(1)
'         TempStr = Apt.CUSTOMER_CODE
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'         Call m_Details1.AddField(BodyCf)
'
'         Di.TOTAL_PRICE = Val(Format(Di.TOTAL_PRICE, "0.00"))
'         Bdc.DISCOUNT_AMOUNT = Val(Format(Bdc.DISCOUNT_AMOUNT, "0.00"))
'         Ri1_1.DEBIT_CREDIT_AMOUNT = Val(Format(Ri1_1.DEBIT_CREDIT_AMOUNT, "0.00"))
'         Ri1_2.DEBIT_CREDIT_AMOUNT = Val(Format(Ri1_2.DEBIT_CREDIT_AMOUNT, "0.00"))
'         Ri1_3.DEBIT_CREDIT_AMOUNT = Val(Format(Ri1_3.DEBIT_CREDIT_AMOUNT, "0.00"))
'         Ari.CASH_DISCOUNT = Val(Format(Ari.CASH_DISCOUNT, "0.00"))
'
'         NetAmt = (Di.TOTAL_PRICE - Bdc.DISCOUNT_AMOUNT) + (Ri1_1.DEBIT_CREDIT_AMOUNT - Ri1_2.DEBIT_CREDIT_AMOUNT - Ri1_3.DEBIT_CREDIT_AMOUNT - Ari.CASH_DISCOUNT) - Ari.PAID_AMOUNT
'         j = 6
'
'         If DueCount <= 0 Then
'            Indue = Indue + NetAmt
'         Else
'            Overdue = Overdue + NetAmt
'         End If
'
'      Set HeadCf = m_Headers0_2.Fields(j)
'      TempStr = FormatNumber(Sum)
'      Tot1(j) = Tot1(j) + NetAmt
'      Tot2(j) = Tot2(j) + NetAmt
''      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
''      Call m_Details1.AddField(BodyCf)
'
'       Tot1(5) = Tot1(j)
'       Tot2(5) = Tot2(j)
'      Set HeadCf = m_Headers0_2.Fields(5)
'      TempStr = FormatNumber(Di.TOTAL_PRICE)
'      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'      Call m_Details1.AddField(BodyCf)
     
'    Else
      Call m_Details2.ClearField
      Set HeadCf = m_Headers1.Fields(1)
      TempStr = Apt.DOCUMENT_NO
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details2.AddField(BodyCf)

      Set HeadCf = m_Headers1.Fields(2)
      TempStr = DateToStringExtEx2(Apt.DOCUMENT_DATE)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details2.AddField(BodyCf)
      
      Set HeadCf = m_Headers1.Fields(3)
      TempStr = DateToStringExtEx2(Apt.DUE_DATE)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details2.AddField(BodyCf)

      Set HeadCf = m_Headers1.Fields(4)
      TempStr = -DueCount
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details2.AddField(BodyCf)

      Set HeadCf = m_Headers1.Fields(5)
      TempStr = FormatNumber(Di.TOTAL_PRICE)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details2.AddField(BodyCf)
      
      'ปัดก่อนแล้วค่อยรวม
      Di.TOTAL_PRICE = Val(Format(Di.TOTAL_PRICE, "0.00"))
      Bdc.DISCOUNT_AMOUNT = Val(Format(Bdc.DISCOUNT_AMOUNT, "0.00"))
      Ri1_1.DEBIT_CREDIT_AMOUNT = Val(Format(Ri1_1.DEBIT_CREDIT_AMOUNT, "0.00"))
      Ri1_2.DEBIT_CREDIT_AMOUNT = Val(Format(Ri1_2.DEBIT_CREDIT_AMOUNT, "0.00"))
      Ri1_3.DEBIT_CREDIT_AMOUNT = Val(Format(Ri1_3.DEBIT_CREDIT_AMOUNT, "0.00"))
      Ari.CASH_DISCOUNT = Val(Format(Ari.CASH_DISCOUNT, "0.00"))

      NetAmt = (Di.TOTAL_PRICE - Bdc.DISCOUNT_AMOUNT) + (Ri1_1.DEBIT_CREDIT_AMOUNT - Ri1_2.DEBIT_CREDIT_AMOUNT - Ri1_3.DEBIT_CREDIT_AMOUNT - Ari.CASH_DISCOUNT) - Ari.PAID_AMOUNT
      'j = 6
      'Sum = 0
      
      Tot1(5) = Tot1(5) + NetAmt
      Tot2(5) = Tot2(5) + NetAmt
      
      If DueCount <= 0 Then
         Indue = Indue + NetAmt
         
         Tot1(8) = Tot1(8) + NetAmt
         Tot2(8) = Tot2(8) + NetAmt
      Else
         Overdue = Overdue + NetAmt
         
         Tot1(7) = Tot1(7) + NetAmt
         Tot2(7) = Tot2(7) + NetAmt
      End If
         
      Set HeadCf = m_Headers1.Fields(6)
      TempStr = FormatNumber(NetAmt)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details2.AddField(BodyCf)
      
      Set HeadCf = m_Headers1.Fields(7)
      TempStr = ""
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details2.AddField(BodyCf)
      
      If DueCount <= 0 Then
         Set HeadCf = m_Headers1.Fields(8)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details2.AddField(BodyCf)
      
         Set HeadCf = m_Headers1.Fields(9)
         TempStr = FormatNumber(NetAmt)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details2.AddField(BodyCf)
      Else
         Set HeadCf = m_Headers1.Fields(8)
         TempStr = FormatNumber(NetAmt)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details2.AddField(BodyCf)
      
         Set HeadCf = m_Headers1.Fields(9)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details2.AddField(BodyCf)
      End If

      Set HeadCf = m_Headers1.Fields(10)
      TempStr = ""
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details2.AddField(BodyCf)
      
      If Check2Flag(mcolParam("SHOW_DETAIL")) = "Y" Then
         Call m_Details2.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
      End If
   Else
      I = I - 1
   End If
   Set BodyCf = Nothing
End Sub

Private Sub GenerateHeader(Cm As CBillingDoc, Details As CFieldList)
Dim BodyCf As CReportField
Dim HeadCf As CReportField
Dim TempStr As String
Dim TempStr1 As String
Dim TempStr2 As String
Dim TempStr3 As String
Dim TempWidth1 As Double
Dim TempWidth2 As Double

Dim J As Long

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   TempStr1 = Cm.CUSTOMER_CODE
   If Cm.SUSPEND_SALES = "Y" Then
      TempStr2 = Cm.CUSTOMER_NAME & "  " & ConvertFlag(Cm.SUSPEND_SALES)
   Else
      TempStr2 = Cm.CUSTOMER_NAME
   End If
   
   TempWidth1 = m_Headers1.Fields(1).Width + m_Headers1.Fields(2).Width + m_Headers1.Fields(3).Width
   TempWidth2 = m_Headers1.Fields(4).Width + m_Headers1.Fields(5).Width + m_Headers1.Fields(6).Width + m_Headers1.Fields(7).Width + m_Headers1.Fields(8).Width + m_Headers1.Fields(9).Width + m_Headers1.Fields(10).Width
   
   Call BodyCf.SetFieldValue(TempWidth1, "<", TempStr1, "<")
   Call Details.AddField(BodyCf)
      
   Call BodyCf.SetFieldValue(TempWidth2, "<", TempStr2, "<")
   Call Details.AddField(BodyCf)
   
   Set BodyCf = Nothing
End Sub
Private Sub GenerateHeader2(Cm As CBillingDoc, Details As CFieldList)
Dim BodyCf As CReportField
Dim HeadCf As CReportField
Dim TempStr As String
Dim TempStr1 As String
Dim TempStr2 As String
Dim TempStr3 As String
Dim J As Long
       
   Call Details.ClearField
   Set BodyCf = New CReportField
   
      TempStr1 = "พนักงานขาย : " & Cm.EMP_NAME & " " & Cm.EMP_LNAME
'   End If
   
   J = 0
   For Each HeadCf In m_Headers0_1.Fields
      J = J + 1
      If J = 1 Then
         TempStr = TempStr1
'      ElseIf j = 2 Then
'         TempStr = TempStr2
'      ElseIf j = 3 Then
'         TempStr = ""
      Else
         TempStr = ""
      End If
      
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub
Private Function GetKey(Mode As Long, Di As CBillingDoc) As String
   If mcolParam("ORDER_BY") = 2 Then
      GetKey = Di.CUSTOMER_CODE
   ElseIf mcolParam("ORDER_BY") = 3 Then
      GetKey = Di.EMP_NAME & "-" & Di.EMP_LNAME
   ElseIf mcolParam("ORDER_BY") = 4 Then
      GetKey = Di.CUSTOMER_CODE & "-" & Di.EMP_NAME & "-" & Di.EMP_LNAME
   ElseIf mcolParam("ORDER_BY") = 5 Then
      GetKey = Di.EMP_NAME & "-" & Di.EMP_LNAME & "-" & Di.CUSTOMER_CODE
   Else
      GetKey = Di.CUSTOMER_CODE
   End If
End Function
Private Function GetKey2(Mode As Long, Di As CBillingDoc) As String
      GetKey2 = Di.CUSTOMER_CODE
End Function
Private Function GetKey3(Mode As Long, Di As CBillingDoc) As String
      GetKey3 = Di.EMP_CODE
End Function
Private Function GetKey4(Mode As Long, Di As CBillingDoc) As String
   If Di.SUSPEND_SALES = "Y" Then
      GetKey4 = Di.CUSTOMER_NAME & "  " & ConvertFlag(Di.SUSPEND_SALES)
   Else
      GetKey4 = Di.CUSTOMER_NAME
   End If
End Function
Private Function GetKey5(Mode As Long, Di As CBillingDoc) As String
      GetKey5 = Di.Credit
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
Dim BillCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim PrevKey4 As String
Dim Width As Double
Dim TempRs As ADODB.Recordset
Dim Ap As CCustomer
Dim TempApt As CDoItem
Dim Apt As CBillingDoc
Dim TempApRcIt As CReceiptItem
Dim Overdue As Double
Dim DebtAmountRecent  As Double
Dim Indue As Double
Dim NewCust As Boolean
Dim HeaderExist As Boolean
Dim CreditLimit As Double
Dim Credit  As Double
'Dim m_AccFol As CAccFol
'Dim m_AccFol2 As CAccFol
'Dim AccFolStr As String


Dim m_AccFol  As CAccFol
Dim m_MKTFol  As CMKTFol
'AccFolStr = ""

Dim Key3 As Long


   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
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
    I = 0
    Set BodyCf = New CReportField

    '=== Detail
    Set Apt = New CBillingDoc
    Apt.CUSTOMER_CODE = mcolParam("CUSTOMER_CODE")
    Apt.FROM_CUSTOMER_CODE = mcolParam("FROM_CUSTOMER_CODE")
    Apt.TO_CUSTOMER_CODE = mcolParam("TO_CUSTOMER_CODE")
    Apt.FROM_DATE = mcolParam("FROM_DOC_DATE")
    Apt.TO_DATE = mcolParam("TO_DOC_DATE")
    Apt.DOCUMENT_TYPE = 1
    Apt.CUSTOMER_TYPE = mcolParam("CUSTOMER_TYPE")
    Apt.CUSTOMER_GRADE = mcolParam("CUSTOMER_GRADE")
    Apt.EMP_CODE = mcolParam("EMP_CODE")
    'Apt.OrderBy = mcolParam("ORDER_BY")
    Apt.OrderType = 1
   'Apt.ItemSumFlag = True
   Call Apt.QueryData(102, Rs, BillCount)

      Indue = 0
      Overdue = 0
        
      If Not Rs.EOF Then
         Call Apt.PopulateFromRS(102, Rs)
         PrevKey1 = GetKey2(2, Apt)
         PrevKey4 = GetKey4(2, Apt)
   
         NewCust = True
     
      End If
    
      While Not Rs.EOF
         Call Apt.PopulateFromRS(102, Rs)
       
         If PrevKey1 <> GetKey2(2, Apt) Then
            
            If HeaderExist Then
               Call GenerateFooter3(CreditLimit, m_Details1, Indue, Overdue, Total1, PrevKey1, PrevKey4, Credit, Key3)
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               If Check2Flag(mcolParam("OVERDUE_FLAG")) = "Y" And ((Overdue > 0) Or ((Indue + Overdue) - CreditLimit > 0)) Then
                  Call Vsp.AddTable(strFormat, "", TempStr2)
               ElseIf Check2Flag(mcolParam("OVERDUE_FLAG")) = "Y" Then
                  Total2(5) = Total2(5) - (Indue + Overdue)
                  Total2(7) = Total2(7) - (Overdue)
                  Total2(8) = Total2(8) - (Indue)
               ElseIf Check2Flag(mcolParam("OVERDUE_FLAG")) = "N" Then
                  Call Vsp.AddTable(strFormat, "", TempStr2)
               End If
               
               
'         If Check2Flag(mcolParam("SHOW_DETAIL")) = "Y" Then
'            Set m_AccFol = GetObject("CAccFol", m_AccountFollowCancelFlag_N, Trim(Str(Key3)), False)
'            If m_AccFol Is Nothing Then
'             Vsp = ""
'
'           Else
'           Vsp.FontUnderline = True
'            Vsp = "การติดตามผลด้านบัญชี :  " & PrevKey4 & "    (" & PrevKey1 & ")"
'             Vsp.FontUnderline = False
'            Vsp = m_AccFol.FOL_NOTE
'
'             Vsp = ""
'
'            End If
'         End If
              
              
              
               Vsp.FontBold = False
               NewCust = True
               HeaderExist = False
            End If
            For J = 1 To UBound(Total1)
               Total1(J) = 0
            Next J
         End If
         PrevKey1 = GetKey2(2, Apt)
        PrevKey4 = GetKey4(2, Apt)
         Key3 = Apt.CUSTOMER_ID
'         If Apt.CUSTOMER_CODE = "ก-0007" Then
'            'Debug.Print
'         End If
         Call m_Details1.ClearField
         I = I + 1
         If PrevKey2 <> GetKey3(2, Apt) Or I = 1 Then
            Call GenerateHeader2(Apt, m_Details1)
         End If
         Call GenerateDetail(alngX(), I, Apt, Total1, Total2, Indue, Overdue, NewCust, HeaderExist)
         CreditLimit = Apt.CREDIT_LIMIT
         Credit = Apt.Credit
          PrevKey2 = GetKey3(2, Apt)
         PrevKey4 = GetKey4(2, Apt)
         Rs.MoveNext
      Wend

      If HeaderExist Then
'         If mcolParam("CONCLUSION_FLAG") = 1 Then
            Key3 = Apt.CUSTOMER_ID
            PrevKey4 = GetKey4(2, Apt)
            Call GenerateFooter3(CreditLimit, m_Details1, Indue, Overdue, Total1, PrevKey1, PrevKey4, Credit, Key3)
            
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            
            
            If Check2Flag(mcolParam("OVERDUE_FLAG")) = "Y" And ((Overdue > 0) Or ((Indue + Overdue) - CreditLimit > 0)) Then
               Call Vsp.AddTable(strFormat, "", TempStr2)
            ElseIf Check2Flag(mcolParam("OVERDUE_FLAG")) = "Y" Then
               Total2(5) = Total2(5) - (Indue + Overdue)
               Total2(7) = Total2(7) - (Overdue)
               Total2(8) = Total2(8) - (Indue)
            ElseIf Check2Flag(mcolParam("OVERDUE_FLAG")) = "N" Then
               Call Vsp.AddTable(strFormat, "", TempStr2)
            End If
            Vsp.FontBold = False
'         End If
             
      End If
      
      Call GenerateFooter4("สุทธิ", "", "", Total2, m_Details1, 0)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      
      

         
'         Else
'            Call GenerateFooter("รวม", "", "", Total1, m_Details1, 0)
'            Call m_Details1.GetString(2, TempStr1, TempStr2)
'            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'            Vsp.FontBold = True
'            Call Vsp.AddTable(strFormat, "", TempStr2)
'            Vsp.FontBold = False
'
'            Call GenerateFooter2(CreditLimit, m_Details1, Indue, Overdue)
'            Call m_Details1.GetString(2, TempStr1, TempStr2)
'            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'            Vsp.FontBold = True
'            Call Vsp.AddTable(strFormat, "", TempStr2)
'            Vsp.FontBold = False
         
      Set Ap = Nothing
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
      ElseIf (J = 4) Then
         TempStr = ""
      ElseIf J = m_Headers1.Fields.Count Then
         TempStr = ""
      Else
         TempStr = FormatNumber(Tot(J))
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub
Private Sub GenerateFooter2(CreditLimit As Double, Details As CFieldList, Indue As Double, Overdue As Double)
Dim BodyCf As CReportField
Dim HeadCf As CReportField
Dim TempStr As String
Dim J As Long

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   J = 0
   For Each HeadCf In m_Headers1_2.Fields
      J = J + 1
      If J = 1 Then
         TempStr = "#"
      ElseIf J = 2 Then
         TempStr = "วงเงิน  = " & FormatNumber(CreditLimit)
      ElseIf J = 3 Then
         TempStr = "วงเงินคงเหลือ = " & FormatNumber(CreditLimit - (Indue + Overdue))
      ElseIf J = 4 Then
'         If mcolParam("SUM_INDUE_FLAG") = 0 Then
'            TempStr = "ในดิว  = " & FormatNumber(Indue)
'         Else
            TempStr = FormatNumber(Indue)
'         End If
      ElseIf J = 5 Then
         TempStr = "เกินดิว  = " & FormatNumber(Overdue)
      Else
         TempStr = ""
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub

Private Sub GenerateFooter3(CreditLimit As Double, Details As CFieldList, Indue As Double, Overdue As Double, Tot() As Double, PrevKey As String, customerName As String, Credit As Double, CustomerID As Long)
Dim BodyCf As CReportField
Dim HeadCf As CReportField
Dim TempStr As String
Dim J As Long
Dim m_AccFol  As CAccFol
Dim m_MKTFol  As CMKTFol
Dim AccFolStr As String
Dim MKTFolStr As String
Dim CountNoteList As Long

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   J = 0
   For Each HeadCf In m_Headers0_2.Fields
      J = J + 1
      If J = 1 Then
         TempStr = PrevKey 'รหัส
     ElseIf J = 2 Then
         TempStr = customerName    'ชื่อลูกค้า
         'TempStr = FormatNumber(Tot(j))
      ElseIf J = 3 Then
         TempStr = Credit   'วงวัน  credit
      ElseIf J = 4 Then
        TempStr = FormatNumber(CreditLimit) 'วงเงิน
      ElseIf J = 5 Then
         TempStr = FormatNumber(Tot(J))   'หนี้ปัจจุบัน
      ElseIf J = 6 Then
         TempStr = FormatNumberToNullMinus(CreditLimit - (Indue + Overdue)) 'วงเงินคงเหลือ
      ElseIf J = 7 Then
         TempStr = FormatNumber(Overdue)  'เกินวงวัน
      ElseIf J = 8 Then
         TempStr = FormatNumber(Indue)  'ใน Due
      ElseIf J = 9 Then
         Set m_AccFol = GetObject("CAccFol", m_AccountFollowCancelFlag_N, Trim(str(CustomerID)), False)
         If m_AccFol Is Nothing Then
            TempStr = ""
         Else
            TempStr = m_AccFol.FOL_NOTE
         End If
       ElseIf J = 10 Then
        Set m_MKTFol = GetObject("CMKTFol", m_MarketingFollowCancelFlag_N, Trim(str(CustomerID)), False)
         If m_MKTFol Is Nothing Then
            TempStr = ""
         Else
            TempStr = m_MKTFol.FOL_NOTE
         End If
      Else
         TempStr = ""
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub
Private Sub GenerateFooter4(TXT1 As String, TXT2 As String, TXT3 As String, Tot() As Double, Details As CFieldList, DateCount As Long)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim J As Long
Dim TempStr As String

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   J = 0
   For Each HeadCf In m_Headers0_2.Fields
      J = J + 1
      If J = 1 Then
         TempStr = TXT1
      ElseIf J = 2 Then
         TempStr = TXT2
      ElseIf J = 3 Then
         TempStr = TXT3
      ElseIf (J = 4) Then
         TempStr = ""
      ElseIf HeadCf.SumFlag Then
         TempStr = FormatNumber(Tot(J))
      Else
         TempStr = ""
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
   
   mdY = Vsp.CurrentY
   
   Vsp.TextAlign = talnT
   Vsp.TableBorder = tbdT
   Vsp.FontBold = blnBold
   Vsp.FontUnderline = blnUnder
   Vsp.FontItalic = blnItalic
   Vsp.FontSize = iSize
   Vsp.FontName = sName

   If mblnHeader Then
      Call printHeader
   End If
End Sub

Private Sub VSP_StartDoc()
'This event occur when VSPrinter.StartDoc is used and used to initialize some information before generating printed document
End Sub




