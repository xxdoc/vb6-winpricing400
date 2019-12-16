VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAR004_11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface
'config from class "CReportAR004_6"
Private Const MODULE_NAME = "CReportAR004_11"
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

Private CollCusGrade As Collection
Private CollCus As Collection




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
'   Set m_SupplierTypes = Nothing
'   Set m_PaidAmounts = Nothing
'   Set m_MaxMins = Nothing
'   Set m_BillTotalPrices = Nothing
'   Set m_DnItemsByBill = Nothing
'   Set m_CnItemsByBill = Nothing
'   Set m_RtItemsByBill = Nothing
'   Set m_BillingDiscounts = Nothing
'   Set m_AccountFollowCancelFlag_N = Nothing
'   Set m_MarketingFollowCancelFlag_N = Nothing
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
         
    If Check2Flag(mcolParam("SUM_FLAG")) = "Y" Then
      Call Cf.SetFieldValue(0, "^", "รหัส", "^")
      Call m_Headers0_2.AddField(Cf)
      Call Cf.SetFieldValue(40, "^", "กลุ่มระดับลูกค้า", "<")
      Call m_Headers0_2.AddField(Cf)
    Else
      Call Cf.SetFieldValue(20, "^", "รหัส", "^")
      Call m_Headers0_2.AddField(Cf)
      Call Cf.SetFieldValue(50, "^", "ชื่อลูกค้า", "<")
      Call m_Headers0_2.AddField(Cf)
    End If
    If Check2Flag(mcolParam("SHOW_DETAIL")) = "Y" Then
      Call Cf.SetFieldValue(15, "^", "เครดิต", "^")
    Else
      Call Cf.SetFieldValue(0, "^", "เครดิต (วัน)", "^")
    End If
    Call m_Headers0_2.AddField(Cf)
    If Check2Flag(mcolParam("SUM_FLAG")) = "Y" Then
      Call Cf.SetFieldValue(0, "^", "วงเงินอนุมัติ", ">")
      Call m_Headers0_2.AddField(Cf)
    Else
      Call Cf.SetFieldValue(20, "^", "วงเงินอนุมัติ", ">")
      Call m_Headers0_2.AddField(Cf)
    End If
    Call Cf.SetFieldValue(20, "^", "หนี้คงเหลือ", ">")
    Call m_Headers0_2.AddField(Cf)
    Call Cf.SetFieldValue(0, "^", "วงเงินคงเหลือ", ">", , False)
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
   Call m_Headers0_2.AddField(Cf)
    
    
'   If Check2Flag(mcolParam("SHOW_COLUMN_FLAG")) = "Y" Then
'     Call Cf.SetFieldValue(50, "^", "การติดตาม", "<", , False) 'การติดตามผลฝ่ายบัญชี
'     Call m_Headers0_2.AddField(Cf)
'     Call Cf.SetFieldValue(0, "^", "การติดตามผลฝ่ายขาย", "<", , False)
'     Call m_Headers0_2.AddField(Cf)
'   Else
      Call Cf.SetFieldValue(0, "^", "การติดตาม", "<", , False) 'การติดตามผลฝ่ายบัญชี
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
    Call Cf.SetFieldValue(0, "^", "", ">", , False)
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
''  If Check2Flag(mcolParam("SHOW_DETAIL")) = "Y" And Check2Flag(mcolParam("SHOW_COLUMN_FLAG")) = "Y" And Check2Flag(mcolParam("SHOW_ROW_FLAG")) = "N" Then
''      Call Cf.SetFieldValue(50, "^", "", "<", , False)
''    Call m_Headers1.AddField(Cf)
''     Call Cf.SetFieldValue(0, "^", "", "<", , False)
''    Call m_Headers1.AddField(Cf)
''    Else
    Call Cf.SetFieldValue(0, "^", "", "<", , False)
    Call m_Headers1.AddField(Cf)
     Call Cf.SetFieldValue(0, "^", "", "<", , False)
    Call m_Headers1.AddField(Cf)
''    End If
    
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

Private Sub GenerateDetail(alngX() As Long, I As Long, Apt As CBillingDoc, Tot1() As Double, Tot2() As Double, Tot3() As Double, Indue As Double, Overdue As Double, NewCust As Boolean, HeaderExist As Boolean)
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
   
'   If Apt.BILLING_DOC_ID = 184776 Then
'      'Debug.Print
'   End If
   
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
      Tot3(5) = Tot3(5) + NetAmt
      
      If DueCount <= 0 Then
         Indue = Indue + NetAmt
         
         Tot1(8) = Tot1(8) + NetAmt
         Tot2(8) = Tot2(8) + NetAmt
         Tot3(8) = Tot3(8) + NetAmt
      Else
         Overdue = Overdue + NetAmt
         
         Tot1(7) = Tot1(7) + NetAmt
         Tot2(7) = Tot2(7) + NetAmt
         Tot3(7) = Tot3(7) + NetAmt
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
   
'      TempStr1 = "พนักงานขาย : " & Cm.EMP_NAME & " " & Cm.EMP_LNAME & " ( " & Cm.EMP_CODE & " )"
      TempStr1 = "ระดับลูกหนี้ : " & Cm.CSTGRADE_NAME & " ( " & Cm.CSTGRADE_NO & " )"
      
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
Private Function GetCusCodeKey(Di As CBillingDoc) As String
      GetCusCodeKey = Di.CUSTOMER_CODE
End Function
Private Function GetSaleCodeKey(Di As CBillingDoc) As String
      GetSaleCodeKey = Di.EMP_CODE
End Function
Private Function GetCusGradeKey(Di As CBillingDoc) As String
      GetCusGradeKey = str(Di.CUSTOMER_GRADE)
End Function
Private Function GetCusGradeNameKey(Di As CBillingDoc) As String
    GetCusGradeNameKey = Di.CSTGRADE_NAME & " ( " & Di.CSTGRADE_NO & " )"
End Function
Private Function GetCusIdKey(Di As CBillingDoc) As Long
    GetCusIdKey = Di.CUSTOMER_ID
End Function
Private Function GetCusNameKey(Di As CBillingDoc) As String
   If Di.SUSPEND_SALES = "Y" Then
      GetCusNameKey = Di.CUSTOMER_NAME & "  " & ConvertFlag(Di.SUSPEND_SALES)
   Else
      GetCusNameKey = Di.CUSTOMER_NAME
   End If
End Function
Private Function GetSaleNameKey(Di As CBillingDoc) As String
      GetSaleNameKey = Di.EMP_NAME & "-" & Di.EMP_LNAME
End Function
Private Function GetHasSelCusGrade(Key As Long) As Boolean
Dim tempCusGrade As CCustomerGrade
   GetHasSelCusGrade = False
   Set tempCusGrade = GetObject("CCustomerGrade", CollCusGrade, Trim(str(Key)), False)
   If Not tempCusGrade Is Nothing Then
      If tempCusGrade.SELECT_FLAG = "Y" Then
         GetHasSelCusGrade = True
      End If
   End If
   Set tempCusGrade = Nothing
End Function
Private Function GetHasSelCus(Key As Long) As Boolean
Dim tempCus As CCustomer
   GetHasSelCus = False
   Set tempCus = GetObject("CCustomer", CollCus, Trim(str(Key)), False)
   If Not tempCus Is Nothing Then
     If tempCus.SELECT_FLAG = "Y" Then
         GetHasSelCus = True
      End If
   End If
   Set tempCus = Nothing
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
Dim Total4(100) As Double 'Sum อื่นๆ ตามกลุ่มระดับ
Dim Total5(100) As Double 'Sum อื่นๆ ตามกลุ่ม
Dim iCount As Long
Dim BillCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim PrevKey3 As String
Dim PrevKey4 As String
Dim PrevKey5 As Long
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


Dim CountCus As Long


Dim m_AccFol  As CAccFol
Dim m_MKTFol  As CMKTFol

'Dim Key3 As Long


   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set Rs = New ADODB.Recordset
   Set TempRs = New ADODB.Recordset
   
Set CollCusGrade = mcolParam("CollCusGrade")
  Set CollCus = mcolParam("CollCus")

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

    '=== Detail
   Set Apt = New CBillingDoc
'   Apt.CUSTOMER_CODE = mcolParam("CUSTOMER_CODE")
'   Apt.FROM_CUSTOMER_CODE = mcolParam("FROM_CUSTOMER_CODE")
'   Apt.TO_CUSTOMER_CODE = mcolParam("TO_CUSTOMER_CODE")
   Apt.FROM_DATE = mcolParam("FROM_DOC_DATE")
   Apt.TO_DATE = mcolParam("TO_DOC_DATE")
   Apt.DOCUMENT_TYPE = 1
'   Apt.CUSTOMER_TYPE = mcolParam("CUSTOMER_TYPE")
'   Apt.CUSTOMER_GRADE = mcolParam("CUSTOMER_GRADE")
'   Apt.EMP_CODE = mcolParam("EMP_CODE")
'   Apt.EMP_RESIGN_FLAG = Check2FlagConvert2(mcolParam("SHOW_EMP_RESIGN_FLAG"))

   Apt.OrderType = 1
   Call Apt.QueryData(116, Rs, BillCount)

   Indue = 0
   Overdue = 0
   NewCust = True
   CountCus = 0 'ไว้นับจำนวนลูกค้าในแต่ละกลุ่ม
   While Not Rs.EOF
      Call Apt.PopulateFromRS(116, Rs)
         
       If PrevKey2 <> GetCusGradeKey(Apt) And I > 0 Then 'GetCusGradeKey
         If HeaderExist And GetHasSelCusGrade(Val(PrevKey2)) And GetHasSelCus(Val(PrevKey5)) Then
            Call GenerateFooter3(CreditLimit, m_Details1, Indue, Overdue, Total1, PrevKey1, PrevKey4, Credit, PrevKey5)
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            If Check2Flag(mcolParam("OVERDUE_FLAG")) = "Y" And ((Overdue > 0) Or ((Indue + Overdue) - CreditLimit > 0)) Then
              If GetHasSelCusGrade(Val(PrevKey2)) And GetHasSelCus(Val(PrevKey5)) And Total1(5) > 0 Then
                     Call Vsp.AddTable(strFormat, "", TempStr2)
                     CountCus = CountCus + 1
               End If
            ElseIf Check2Flag(mcolParam("OVERDUE_FLAG")) = "Y" Then
               Total2(5) = Total2(5) - (Indue + Overdue)
               Total2(7) = Total2(7) - (Overdue)
               Total2(8) = Total2(8) - (Indue)
               
               Total3(5) = Total3(5) - (Indue + Overdue)
               Total3(7) = Total3(7) - (Overdue)
               Total3(8) = Total3(8) - (Indue)
               
            ElseIf Check2Flag(mcolParam("OVERDUE_FLAG")) = "N" Then
               If GetHasSelCusGrade(Val(PrevKey2)) And GetHasSelCus(Val(PrevKey5)) And Total1(5) > 0 Then
                     Call Vsp.AddTable(strFormat, "", TempStr2)
                     CountCus = CountCus + 1
               End If
            End If
               
               
''            If Check2Flag(mcolParam("SHOW_ROW_FLAG")) = "Y" Then
''               Set m_AccFol = GetObject("CAccFol", m_AccountFollowCancelFlag_N, Trim(str(Key3)), False)
''               If m_AccFol Is Nothing Then
''              Else
''                  Vsp.FontUnderline = True
''                  Vsp = "การติดตามผลด้านบัญชี :  " & PrevKey4 & "    (" & PrevKey1 & ")"
''                  Vsp.FontUnderline = False
''                  Vsp = m_AccFol.FOL_NOTE
''
''                  Vsp = ""
''               End If
''            End If
              
            Vsp.FontBold = False
            NewCust = True
            HeaderExist = False
         End If
         
     If GetHasSelCusGrade(Val(PrevKey2)) Then
     
       If Check2Flag(mcolParam("SUM_FLAG")) = "N" And CountCus > 0 Then
            Call GenerateFooter4("", "ลูกค้า อื่นๆ", "", Total4, m_Details1, 0)
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
         End If

            
         If Check2Flag(mcolParam("SUM_FLAG")) = "Y" Then
            Call GenerateFooter4("", PrevKey3, "", Total2, m_Details1, 0)
         Else
            Call GenerateFooter4("", "รวม " & PrevKey3, "", Total2, m_Details1, 0)
         End If
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         If Total2(5) > 0 Then
            Call Vsp.AddTable(strFormat, "", TempStr2)
         End If
         Vsp.FontBold = False
      Else
         Total5(5) = Total5(5) + Total2(5)
         Total5(7) = Total5(7) + Total2(7)
         Total5(8) = Total5(8) + Total2(8)
       End If
       
       If Check2Flag(mcolParam("SUM_FLAG")) = "N" Then
         If GetHasSelCusGrade(Val(Apt.CUSTOMER_GRADE)) Then
              Call GenerateHeader2(Apt, m_Details1)
              Call m_Details1.GetString(2, TempStr1, TempStr2)                  'แสดงหัวของ SALE
              strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
              Vsp.FontBold = True
              Call Vsp.AddTable(strFormat, "", TempStr2)
              Vsp.FontBold = False
           End If
         End If
         
         For J = 1 To UBound(Total1)
            Total1(J) = 0
            Total2(J) = 0
            Total4(J) = 0
         Next J
      CountCus = 0

      ElseIf PrevKey1 <> GetCusCodeKey(Apt) And I > 0 Then
         If HeaderExist Then
            Call GenerateFooter3(CreditLimit, m_Details1, Indue, Overdue, Total1, PrevKey1, PrevKey4, Credit, PrevKey5)
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            
                  If Check2Flag(mcolParam("OVERDUE_FLAG")) = "Y" And ((Overdue > 0) Or ((Indue + Overdue) - CreditLimit > 0)) Then
                    If GetHasSelCusGrade(Val(PrevKey2)) And GetHasSelCus(Val(PrevKey5)) And Total1(5) > 0 Then
                     Call Vsp.AddTable(strFormat, "", TempStr2)
                     CountCus = CountCus + 1
                     Else
                           Total4(5) = Total4(5) + Total1(5)
                           Total4(7) = Total4(7) + Total1(7)
                           Total4(8) = Total4(8) + Total1(8)
                     End If
                  ElseIf Check2Flag(mcolParam("OVERDUE_FLAG")) = "Y" Then
                     Total2(5) = Total2(5) - (Indue + Overdue)
                     Total2(7) = Total2(7) - (Overdue)
                     Total2(8) = Total2(8) - (Indue)
                  
                     Total3(5) = Total3(5) - (Indue + Overdue)
                     Total3(7) = Total3(7) - (Overdue)
                     Total3(8) = Total3(8) - (Indue)
                  
                  ElseIf Check2Flag(mcolParam("OVERDUE_FLAG")) = "N" Then
                     If GetHasSelCusGrade(Val(PrevKey2)) And GetHasSelCus(Val(PrevKey5)) And Total1(5) > 0 Then
                           Call Vsp.AddTable(strFormat, "", TempStr2)
                           CountCus = CountCus + 1
                     Else
                           Total4(5) = Total4(5) + Total1(5)
                           Total4(7) = Total4(7) + Total1(7)
                           Total4(8) = Total4(8) + Total1(8)
                     End If
                  End If
               
'            If Check2Flag(mcolParam("SHOW_ROW_FLAG")) = "Y" Then
'               Set m_AccFol = GetObject("CAccFol", m_AccountFollowCancelFlag_N, Trim(str(Key3)), False)
'               If m_AccFol Is Nothing Then
'              Else
'                  Vsp.FontUnderline = True
'                  Vsp = "การติดตามผลด้านบัญชี :  " & PrevKey4 & "    (" & PrevKey1 & ")"
'                  Vsp.FontUnderline = False
'                  Vsp = m_AccFol.FOL_NOTE
'
'                  Vsp = ""
'               End If
'            End If
'         ElseIf HeaderExist Then
'           Total4(5) = Total4(5) + Total1(5)
         End If
                
            Vsp.FontBold = False
            NewCust = True
            HeaderExist = False
            
         For J = 1 To UBound(Total1)
            Total1(J) = 0
         Next J
      ElseIf I = 0 Then
        If Check2Flag(mcolParam("SUM_FLAG")) = "N" Then
            If GetHasSelCusGrade(Val(Apt.CUSTOMER_GRADE)) Then
               Call GenerateHeader2(Apt, m_Details1)
               Call m_Details1.GetString(2, TempStr1, TempStr2)                  'แสดงหัวของ SALE
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
            End If
         End If
      End If
         
         PrevKey1 = GetCusCodeKey(Apt)
         PrevKey2 = GetCusGradeKey(Apt) 'GetSaleCodeKey(Apt)
         PrevKey3 = GetCusGradeNameKey(Apt) '
         PrevKey4 = GetCusNameKey(Apt)
         PrevKey5 = GetCusIdKey(Apt)

         Call m_Details1.ClearField
         I = I + 1
         
         Call GenerateDetail(alngX(), I, Apt, Total1, Total2, Total3, Indue, Overdue, NewCust, HeaderExist)
         CreditLimit = Apt.CREDIT_LIMIT
         Credit = Apt.Credit
         
         Rs.MoveNext
      Wend

         If HeaderExist And GetHasSelCusGrade(Val(PrevKey2)) And GetHasSelCus(Val(PrevKey5)) Then
            Call GenerateFooter3(CreditLimit, m_Details1, Indue, Overdue, Total1, PrevKey1, PrevKey4, Credit, PrevKey5)
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            If Check2Flag(mcolParam("OVERDUE_FLAG")) = "Y" And ((Overdue > 0) Or ((Indue + Overdue) - CreditLimit > 0)) Then
               If GetHasSelCusGrade(Val(PrevKey2)) And GetHasSelCus(Val(PrevKey5)) And Total1(5) > 0 Then
                  Call Vsp.AddTable(strFormat, "", TempStr2)
               End If
            ElseIf Check2Flag(mcolParam("OVERDUE_FLAG")) = "Y" Then
               Total2(5) = Total2(5) - (Indue + Overdue)
               Total2(7) = Total2(7) - (Overdue)
               Total2(8) = Total2(8) - (Indue)
               
               Total3(5) = Total3(5) - (Indue + Overdue)
               Total3(7) = Total3(7) - (Overdue)
               Total3(8) = Total3(8) - (Indue)
               
            ElseIf Check2Flag(mcolParam("OVERDUE_FLAG")) = "N" Then
               If GetHasSelCusGrade(Val(PrevKey2)) And GetHasSelCus(Val(PrevKey5)) And Total1(5) > 0 Then
                  Call Vsp.AddTable(strFormat, "", TempStr2)
               End If
            End If


'         If Check2Flag(mcolParam("SHOW_ROW_FLAG")) = "Y" Then
'            Set m_AccFol = GetObject("CAccFol", m_AccountFollowCancelFlag_N, Trim(str(Key3)), False)
'            If m_AccFol Is Nothing Then
'           Else
'               Vsp.FontUnderline = True
'               Vsp = "การติดตามผลด้านบัญชี :  " & PrevKey4 & "    (" & PrevKey1 & ")"
'               Vsp.FontUnderline = False
'               Vsp = m_AccFol.FOL_NOTE
'
'               Vsp = ""
'            End If
       Else
         Vsp.FontBold = False
         NewCust = True
         HeaderExist = False
         
           End If

''
     If GetHasSelCusGrade(Val(PrevKey2)) Then
            
            If Check2Flag(mcolParam("SUM_FLAG")) = "N" And CountCus > 0 Then
               Call GenerateFooter4("", "ลูกค้า อื่นๆ", "", Total4, m_Details1, 0)
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
            End If
         
         If Check2Flag(mcolParam("SUM_FLAG")) = "Y" Then
            Call GenerateFooter4("", PrevKey3, "", Total2, m_Details1, 0)
         Else
            Call GenerateFooter4("", "รวม " & PrevKey3, "", Total2, m_Details1, 0)
         End If
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         If Total2(5) > 0 Then
            Call Vsp.AddTable(strFormat, "", TempStr2)
         End If
         Vsp.FontBold = False
      Else
         Total5(5) = Total5(5) + Total2(5)
         Total5(7) = Total5(7) + Total2(7)
         Total5(8) = Total5(8) + Total2(8)
            
         Call GenerateFooter4("รวมกลุ่ม อื่นๆ", "", "", Total5, m_Details1, 0)
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
       End If
     If Check2Flag(mcolParam("SUM_FLAG")) = "Y" Then
         Call GenerateFooter4("", "สุทธิ", "", Total3, m_Details1, 0)
      Else
         Call GenerateFooter4("สุทธิ", "", "", Total3, m_Details1, 0)
      End If
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
               
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
         TempStr = "วงเงินคงเกลือ = " & FormatNumber(CreditLimit - (Indue + Overdue))
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
         TempStr = FormatNumberToNullMinus(CreditLimit - (Indue + Overdue)) 'วงเงินคงเกลือ
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
