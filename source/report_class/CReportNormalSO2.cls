VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportNormalSO2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportNormalSO2"
Private MODULE_DESC As String

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

Private Const TITLE_SIZE = 18
Private Const HEADER_SIZE = 14
Private Const DETAIL_SIZE = 14

Private Const TITLE_FORMAT = "<32|<3|<5;"
Private Const HEADER_FORMAT = "^1|^6|^2|^2|^3;"
Private Const DETAIL_FORMAT = "^1|<6|>2|>2|>3;"
Private Const FOOTER_FORMAT = "^1|>6|>2|>2|>3;"

Private m_Enterprise As CEnterprise
Private m_CustAddress As CAddress
Private m_EnpAddress As CAddress
Private m_EnpPerson As cName
Private Sp As CSystemParam
Private m_OwnerAddress As CAddress
Private m_cData As CBillingDoc

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_Enterprise = New CEnterprise
   Set m_CustAddress = New CAddress
   Set m_EnpAddress = New CAddress
   Set m_EnpPerson = New cName
   Set m_OwnerAddress = New CAddress
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_Enterprise = Nothing
   Set m_CustAddress = Nothing
   Set m_EnpAddress = Nothing
   Set m_EnpPerson = Nothing
   Set m_OwnerAddress = Nothing
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

Private Sub PrintPOFooter(alngX() As Long)
Dim strFormat As String
Dim strPrint As String

   Vsp.TableBorder = tbNone
   Vsp.FontBold = False
   
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("^10|^10|^10;", mdblWid, alngX)
   strPrint = MapText("วันที่รับงาน") & EmptyToString(DateToStringExtEx2(m_cData.DOCUMENT_DATE), "..................................................") & " |ผู้รับเงิน .........................................................| ผู้รับงาน .......................................................;"
   Call Vsp.AddTable(strFormat, "", strPrint)
End Sub

Private Sub GetARAmount(CustomerID As Long, Credit As Long, Overdue As Double, NotOverDue As Double)
Dim Bd As CBillingDoc
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim TempBd As CBillingDoc
Dim DueCount As Long
Dim Amt As Double
Dim Di As CDoItem
Dim Bdc As CBillingDiscount
Dim Ari As CReceiptItem
Dim Ri1_1 As CReceiptItem
Dim Ri1_2 As CReceiptItem
Dim Ri1_3 As CReceiptItem
Dim m_BillTotalPrices As Collection
Dim m_BillingDiscounts As Collection
Dim m_PaidAmounts As Collection
Dim m_DnItemsByBill As Collection
Dim m_CnItemsByBill As Collection
Dim m_RtItemsByBill As Collection

   Set TempRs = New ADODB.Recordset
   Set Bd = New CBillingDoc
   
   Set m_BillTotalPrices = New Collection
   Set m_BillingDiscounts = New Collection
   Set m_PaidAmounts = New Collection
   Set m_DnItemsByBill = New Collection
   Set m_CnItemsByBill = New Collection
   Set m_RtItemsByBill = New Collection
   
   Call LoadPaidAmountByBill(Nothing, m_PaidAmounts, -1, -1)
   Call LoadTotalPriceByBill(Nothing, m_BillTotalPrices, -1, -1)
   Call LoadBillingDiscountByBill(Nothing, m_BillingDiscounts, -1, -1)
   Call LoadDnCnAmountByBill(Nothing, m_DnItemsByBill, -1, -1, 4)
   Call LoadDnCnAmountByBill(Nothing, m_CnItemsByBill, -1, -1, 3)
   Call LoadDnCnAmountByBill(Nothing, m_RtItemsByBill, -1, -1, 18)
   
   Bd.CUSTOMER_ID = CustomerID
   Bd.FROM_DATE = -1
   Bd.TO_DATE = -1
   Bd.DOCUMENT_TYPE = 1
   Bd.OrderBy = 7
   Bd.OrderType = 1
   Bd.ItemSumFlag = True
   Call Bd.QueryData(1, TempRs, iCount)
   
   Overdue = 0
   NotOverDue = 0
   While Not TempRs.EOF
      Call Bd.PopulateFromRS(1, TempRs)
      
      Set Di = GetDoItem(m_BillTotalPrices, Bd.BILLING_DOC_ID)
      Set Bdc = GetBillingDiscount(m_BillingDiscounts, Bd.BILLING_DOC_ID)
      Set Ari = GetReceiptItem(m_PaidAmounts, Bd.BILLING_DOC_ID)
      Set Ri1_1 = GetReceiptItem(m_DnItemsByBill, Bd.BILLING_DOC_ID)
      Set Ri1_2 = GetReceiptItem(m_CnItemsByBill, Bd.BILLING_DOC_ID)
      Set Ri1_3 = GetReceiptItem(m_RtItemsByBill, Bd.BILLING_DOC_ID)
      
      DueCount = DateDiff("D", Bd.DUE_DATE, m_cData.DOCUMENT_DATE)
      Amt = (Di.TOTAL_PRICE - Bdc.DISCOUNT_AMOUNT) + (Ri1_1.DEBIT_CREDIT_AMOUNT - Ri1_2.DEBIT_CREDIT_AMOUNT - Ri1_3.DEBIT_CREDIT_AMOUNT) - Ari.PAID_AMOUNT
      If DueCount > 0 Then
         Overdue = Overdue + Amt
      Else
         NotOverDue = NotOverDue + Amt
      End If
      
      Call TempRs.MoveNext
   Wend
   
   If TempRs.State <> adStateOpen Then
      Call TempRs.Close
   End If
   Set Bd = Nothing

   Set m_BillTotalPrices = Nothing
   Set m_BillingDiscounts = Nothing
   Set m_PaidAmounts = Nothing
   Set m_DnItemsByBill = Nothing
   Set m_CnItemsByBill = Nothing
   Set m_RtItemsByBill = Nothing
End Sub

Private Sub PrintPOHeader(alngX() As Long)
Dim strFormat As String
Dim strPrint As String
Static I As Long
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim OwnerFlag As String
Dim TempStr As String
Dim TempOwnerAddress As String

   If I = 0 Then
      I = 1
      Set TempRs = New ADODB.Recordset
      
      m_Enterprise.ENTERPRISE_ID = -1
      Call m_Enterprise.QueryData(TempRs, iCount)
      If Not TempRs.EOF Then
         Call m_Enterprise.PopulateFromRS(1, TempRs)
      End If
      
      m_EnpAddress.ADDRESS_ID = m_cData.ENTERPRISE_ADDRESS_ID
      Call m_EnpAddress.QueryData(TempRs, iCount)
      If Not TempRs.EOF Then
         Call m_EnpAddress.PopulateFromRS(TempRs)
      End If
      
      m_CustAddress.ADDRESS_ID = -1
      m_CustAddress.CUSTOMER_ID = m_cData.CUSTOMER_ID
      Call m_CustAddress.QueryData3(TempRs, iCount)
      If Not TempRs.EOF Then
         Call m_CustAddress.PopulateFromRS(TempRs)
      End If
      
      m_EnpPerson.NAME_ID = -1
      m_EnpPerson.ENTERPRISE_ID = -1
      m_EnpPerson.MASTER_FLAG = "Y"
      Call m_EnpPerson.QueryData2(TempRs, iCount)
      If Not TempRs.EOF Then
         Call m_EnpPerson.PopulateFromRS2(TempRs)
      End If
      
      TempOwnerAddress = ""
      If m_EnpPerson.ADDRESS_ID > 0 Then
         m_OwnerAddress.ADDRESS_ID = m_EnpPerson.ADDRESS_ID
         m_OwnerAddress.CUSTOMER_ID = -1
         Call m_OwnerAddress.QueryData(TempRs, iCount)
         If Not TempRs.EOF Then
            Call m_OwnerAddress.PopulateFromRS(TempRs)
            TempOwnerAddress = vbCrLf & m_OwnerAddress.PackAddress
         End If
      End If
      
      If TempRs.State = adStateOpen Then
         TempRs.Close
      End If
      Set TempRs = Nothing
   
'       Set Sp = GetSystemParam(glbSystemParam, "PRINT_OWNER_FLAG")
'       If Sp.PARAM_VALUE = "" Then
'          glbErrorLog.LocalErrorMsg = GetTextMessage("TEXT-KEY559")
'          glbErrorLog.ShowUserError
'       Else
'          OwnerFlag = Sp.PARAM_VALUE
'       End If
   End If
   
   Vsp.TableBorder = tbNone
   Vsp.FontSize = 14
   Vsp.FontBold = False
   
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = m_Enterprise.ENTERPRISE_NAME & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)

   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = MapText("ที่อยู่") & m_EnpAddress.PackAddress & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)

   If OwnerFlag = "Y" Then
      TempStr = m_EnpPerson.LONG_NAME & " " & m_EnpPerson.LAST_NAME
      If Trim(TempStr) <> "" Then
         TempStr = TempStr & ", "
      End If
   End If

   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = MapText("โทรศัพท์ ") & EmptyToString(m_EnpAddress.PHONE1, "N/A") & " แฟกซ์ " & EmptyToString(m_EnpAddress.FAX1, "N/A") & "  (" & TempStr & MapText("เลขประจำตัวผู้เสียภาษี ") & m_Enterprise.TAX_ID & TempOwnerAddress & ") " & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
      
   strFormat = VSP_CalTable("<10|>10;", mdblWid, alngX)
   strPrint = MapText("วันที่เอกสาร ") & DateToStringExtEx2(m_cData.DOCUMENT_DATE) & "| เลขที่เอกสาร " & m_cData.DOCUMENT_NO & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("<10|>10;", mdblWid, alngX)
   strPrint = MapText("ชื่อลูกค้า ") & m_cData.CUSTOMER_NAME & "|" & MapText("รหัสลูกค้า ") & m_cData.CUSTOMER_CODE & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)

   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = MapText("โทรศัพท์ ") & EmptyToString(m_CustAddress.PHONE1, "N/A") & " แฟกซ์ " & EmptyToString(m_CustAddress.FAX1, "N/A") & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
End Sub

Private Sub printHeader()
   Dim strFormat As String
   Dim strPrint As String
   Dim tbdT As TableBorderSettings
   Dim blnBold As Boolean
   Dim iSize As Integer
   Dim alngX() As Long
   
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Call PrintPOHeader(alngX)
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   strFormat = VSP_CalTable(HEADER_FORMAT, mdblWid, alngX)
   strPrint = "ลำดับ|รายการ|จำนวน|ราคาต่อหน่วย|จำนวนเงิน;"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Vsp.TableBorder = tbdT
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
End Sub

Private Function initDoc() As Boolean
   Dim strDate As String
   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprLetter
   Vsp.ORIENTATION = orPortrait
   Vsp.MarginBottom = 1440
   Vsp.MarginFooter = 1440
   Vsp.MarginHeader = 1440
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 1440
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   MODULE_DESC = MapText("แบบฟอร์มขออนุมัติขายเกินวงเงิน")
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
Dim iCount As Long
Dim I As Long
Dim J As Long
Dim strFormat As String
Dim strPrint As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Sum1 As Double
Dim Sum2 As Double
Dim Overdue As Double
Dim NotOverDue As Double
Dim Cq As CCheque
Dim Tempsum As Double

   RName = "genDoc"
   
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   Set Rs = New ADODB.Recordset
   Set m_cData = New CBillingDoc
   m_cData.BILLING_DOC_ID = mcolParam("BILLING_DOC_ID")
   m_cData.QueryFlag = -1
   If Not glbDaily.QueryBillingDoc(m_cData, Rs, iCount, IsOK, glbErrorLog) Then
      mstrErrMsg = "Query Error."
      Exit Function
   End If
   Call m_cData.PopulateFromRS(1, Rs)
   
   If (Rs Is Nothing) Or (iCount < 1) Then
      mstrErrMsg = "ไม่พบข้อมูลที่ต้องการในฐานข้อมูล"
      Exit Function
   End If
   
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
   Sum1 = 0
   Sum2 = 0
   
   'ไม่ต้องแสดงหัว
   Call StartExportFile(Vsp)
   Vsp.StartDoc

      Vsp.TableBorder = tbNone
      
      strFormat = VSP_CalTable("<10;", mdblWid, alngX)
      strPrint = "วันที่ " & DateToStringExt(m_cData.DOCUMENT_DATE) & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strFormat = VSP_CalTable("<10|<10|>10|>10|>10|>10;", mdblWid, alngX)
      strPrint = "ชื่อลูกค้า|" & m_cData.CUSTOMER_NAME & "|วงเงิน|" & FormatNumber(m_cData.CREDIT_LIMIT) & "|บาท|เครดิต   " & m_cData.CUSTOMER_CREDIT & "  วัน;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      Call GetARAmount(m_cData.CUSTOMER_ID, m_cData.Credit, Overdue, NotOverDue)
      
      Sum1 = Overdue + NotOverDue
      strFormat = VSP_CalTable("<10|<10|>10|>10|>10|>10;", mdblWid, alngX)
      strPrint = "|ยอดหนี้ปัจจุบัน||" & FormatNumber(Sum1) & "|บาท|;"
      Call Vsp.AddTable(strFormat, "", strPrint)
         
      Sum2 = Format(m_cData.TOTAL_PRICE, "0.00") - Format(m_cData.DISCOUNT_AMOUNT, "0.00")
      strFormat = VSP_CalTable("<10|<10|>10|>10|>10|>10;", mdblWid, alngX)
      strPrint = "|ต้องการซื้อเพิ่มเติม||" & FormatNumber(Sum2) & "|บาท|;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strFormat = VSP_CalTable("<10|<10|>10|>10|>10|>10;", mdblWid, alngX)
      strPrint = "|หนี้ปัจจุบัน+ซื้อเพิ่ม||" & FormatNumber(Sum1 + Sum2) & "|บาท|;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strFormat = VSP_CalTable("<10|<10|>10|>10|>10|>10;", mdblWid, alngX)
      strPrint = "|คงเหลือวงเงิน||" & FormatNumber(m_cData.CREDIT_LIMIT - (Sum1 + Sum2)) & "|บาท|;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      
      Vsp.FontBold = True
      Vsp.Paragraph = "ความเห็นฝ่ายข้อมูล"
      Vsp.FontBold = False
      strFormat = VSP_CalTable("<10|<10|>10|>10|>10|>10;", mdblWid, alngX)
      strPrint = "|ยอดหนี้ที่เกินดิว||" & FormatNumber(Overdue) & "|บาท|;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      strPrint = "|ยอดหนี้ที่ไม่เกินดิว||" & FormatNumber(NotOverDue) & "|บาท|;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      Set Cq = New CCheque
      Call Cq.SetFieldValue("APAR_ID", m_cData.CUSTOMER_ID)
      Call Cq.SetFieldValue("POST_FLAG", "N")
      Call Cq.SetFieldValue("BANK_FLAG", "")
      Call Cq.QueryData(1, Rs, iCount)
      If iCount > 0 Then
'         Vsp.Paragraph = ""
         Vsp.TableBorder = tbAll
         Vsp.FontBold = True
         strFormat = VSP_CalTable("^1|^3|^3|^2;", mdblWid, alngX)
         strPrint = "NO.|เลขที่เช็ค|วันที่ดิว|จำนวนเงิน;"
         Call Vsp.AddTable(strFormat, "", strPrint)
         Vsp.FontBold = False
         
         J = 0
         Tempsum = 0
         While Not Rs.EOF
            J = J + 1
            Call Cq.PopulateFromRS(1, Rs)
            
            strFormat = VSP_CalTable("^1|^3|^3|>2;", mdblWid, alngX)
            strPrint = J & "|" & Cq.GetFieldValue("CHEQUE_NO") & "|" & DateToStringExtEx2(Cq.GetFieldValue("EFFECTIVE_DATE")) & "|" & FormatNumber(Cq.GetFieldValue("CHEQUE_AMOUNT")) & ";"
            Call Vsp.AddTable(strFormat, "", strPrint)
            Tempsum = Tempsum + Cq.GetFieldValue("CHEQUE_AMOUNT")
            
            Rs.MoveNext
         Wend
         
         strFormat = VSP_CalTable("^1|^3|^3|>2;", mdblWid, alngX)
         strPrint = "รวม|||" & FormatNumber(Tempsum) & ";"
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", strPrint)
         Vsp.FontBold = False
         
         Vsp.TableBorder = tbNone
      End If
      Set Cq = Nothing
      
      'หา DUE ที่เกิน
      
      Dim TempX As Long
      Dim TempY As Long
      
      'ความเห็นฝ่ายการตลาด
      '============
      Vsp.Paragraph = ""
      Vsp.FontBold = True
      Vsp.Paragraph = "ความเห็นฝ่ายการตลาด"
      Vsp.FontBold = False
      Vsp.Paragraph = ""
      
'      Vsp.Paragraph = ""
      Call Vsp.DrawLine(Vsp.CurrentX + 100, Vsp.CurrentY, Vsp.PageWidth - 1000, Vsp.CurrentY)
      Vsp.Paragraph = ""
      
'      Vsp.Paragraph = ""
      Call Vsp.DrawLine(Vsp.CurrentX + 100, Vsp.CurrentY, Vsp.PageWidth - 1000, Vsp.CurrentY)
      Vsp.Paragraph = ""
      
      Vsp.FontBold = True
      Call Vsp.DrawLine(Vsp.CurrentX + 500, Vsp.CurrentY, Vsp.CurrentX + 1000, Vsp.CurrentY)
      TempX = Vsp.CurrentX
      TempY = Vsp.CurrentY
      Vsp.CurrentX = Vsp.CurrentX + 1000
      Vsp.CurrentY = Vsp.CurrentY - 100
      Vsp = "    เห็นด้วย"
      Vsp.CurrentX = TempX
      Vsp.CurrentY = TempY
      Call Vsp.DrawLine(Vsp.PageWidth - 4500, Vsp.CurrentY, Vsp.PageWidth - 4000, Vsp.CurrentY)
      TempX = Vsp.CurrentX
      Vsp.CurrentX = Vsp.PageWidth - 4000
      Vsp.CurrentY = Vsp.CurrentY - 100
      Vsp = "    ไม่เห็นด้วย"
      Vsp.CurrentX = TempX
      Vsp.CurrentY = TempY
      
      Call Vsp.DrawLine(Vsp.CurrentX + 500, Vsp.CurrentY, Vsp.CurrentX + 500, Vsp.CurrentY + 250)
      Call Vsp.DrawLine(Vsp.CurrentX + 1000, Vsp.CurrentY, Vsp.CurrentX + 1000, Vsp.CurrentY + 250)
      Call Vsp.DrawLine(Vsp.PageWidth - 4500, Vsp.CurrentY, Vsp.PageWidth - 4500, Vsp.CurrentY + 250)
      Call Vsp.DrawLine(Vsp.PageWidth - 4000, Vsp.CurrentY, Vsp.PageWidth - 4000, Vsp.CurrentY + 250)
      
      Call Vsp.DrawLine(Vsp.CurrentX + 500, Vsp.CurrentY + 250, Vsp.CurrentX + 1000, Vsp.CurrentY + 250)
      Call Vsp.DrawLine(Vsp.PageWidth - 4500, Vsp.CurrentY + 250, Vsp.PageWidth - 4000, Vsp.CurrentY + 250)
      Vsp.FontBold = False
      
'      Vsp.Paragraph = ""
      Vsp.Paragraph = ""
      Vsp = "  ลงชื่อ    _________________________________________________"
      '============
      If m_cData.AGREEMENT_DATA <> "" Then
         Vsp.Paragraph = m_cData.AGREEMENT_DATA
         Vsp.Paragraph = ""
      Else
         Vsp.Paragraph = ""
      End If
      Vsp.FontBold = True
      Vsp.Paragraph = "ความเห็นฝ่ายสินชื่อ"
      Vsp.FontBold = False
      Vsp.Paragraph = m_cData.AGREEMENT_FINANCE
      
'      Vsp.Paragraph = ""
      Call Vsp.DrawLine(Vsp.CurrentX + 100, Vsp.CurrentY, Vsp.PageWidth - 1000, Vsp.CurrentY)
      Vsp.Paragraph = ""
      
'      Vsp.Paragraph = ""
      Call Vsp.DrawLine(Vsp.CurrentX + 100, Vsp.CurrentY, Vsp.PageWidth - 1000, Vsp.CurrentY)
      Vsp.Paragraph = ""
    
      Call Vsp.DrawLine(Vsp.CurrentX + 100, Vsp.CurrentY, Vsp.PageWidth - 1000, Vsp.CurrentY)
      Vsp.Paragraph = ""
      
      Call Vsp.DrawLine(Vsp.CurrentX + 100, Vsp.CurrentY, Vsp.PageWidth - 1000, Vsp.CurrentY)
      Vsp.Paragraph = ""
      
      Vsp.FontBold = True
      Call Vsp.DrawLine(Vsp.CurrentX + 500, Vsp.CurrentY, Vsp.CurrentX + 1000, Vsp.CurrentY)
      TempX = Vsp.CurrentX
      TempY = Vsp.CurrentY
      Vsp.CurrentX = Vsp.CurrentX + 1000
      Vsp.CurrentY = Vsp.CurrentY - 100
      Vsp = "    เห็นด้วย"
      Vsp.CurrentX = TempX
      Vsp.CurrentY = TempY
      Call Vsp.DrawLine(Vsp.PageWidth - 4500, Vsp.CurrentY, Vsp.PageWidth - 4000, Vsp.CurrentY)
      TempX = Vsp.CurrentX
      Vsp.CurrentX = Vsp.PageWidth - 4000
      Vsp.CurrentY = Vsp.CurrentY - 100
      Vsp = "    ไม่เห็นด้วย"
      Vsp.CurrentX = TempX
      Vsp.CurrentY = TempY
      
      Call Vsp.DrawLine(Vsp.CurrentX + 500, Vsp.CurrentY, Vsp.CurrentX + 500, Vsp.CurrentY + 250)
      Call Vsp.DrawLine(Vsp.CurrentX + 1000, Vsp.CurrentY, Vsp.CurrentX + 1000, Vsp.CurrentY + 250)

      Call Vsp.DrawLine(Vsp.PageWidth - 4500, Vsp.CurrentY, Vsp.PageWidth - 4500, Vsp.CurrentY + 250)
      Call Vsp.DrawLine(Vsp.PageWidth - 4000, Vsp.CurrentY, Vsp.PageWidth - 4000, Vsp.CurrentY + 250)
      
      Call Vsp.DrawLine(Vsp.CurrentX + 500, Vsp.CurrentY + 250, Vsp.CurrentX + 1000, Vsp.CurrentY + 250)
      Call Vsp.DrawLine(Vsp.PageWidth - 4500, Vsp.CurrentY + 250, Vsp.PageWidth - 4000, Vsp.CurrentY + 250)
      Vsp.FontBold = False
      
      Vsp.Paragraph = ""
      
'      Call Vsp.DrawLine(Vsp.CurrentX + 100, Vsp.CurrentY, Vsp.PageWidth - 1000, Vsp.CurrentY)
'      Vsp.Paragraph = ""
'      Call Vsp.DrawLine(Vsp.CurrentX + 100, Vsp.CurrentY, Vsp.PageWidth - 1000, Vsp.CurrentY)
'      Vsp.Paragraph = ""
      Vsp = "  ลงชื่อ    _________________________________________________"
      Vsp.Paragraph = ""
      Vsp.Paragraph = ""
      Call Vsp.DrawLine(Vsp.CurrentX + 100, Vsp.CurrentY, Vsp.PageWidth - 1000, Vsp.CurrentY)
      Vsp.Paragraph = ""
      Call Vsp.DrawLine(Vsp.CurrentX + 100, Vsp.CurrentY, Vsp.PageWidth - 1000, Vsp.CurrentY)
      
      '------------------------------------------------------------------>
      Vsp.FontBold = True
      Vsp.Paragraph = ""
      Call Vsp.DrawLine(Vsp.CurrentX + 500, Vsp.CurrentY, Vsp.CurrentX + 1000, Vsp.CurrentY)
      TempX = Vsp.CurrentX
      TempY = Vsp.CurrentY
      Vsp.CurrentX = Vsp.CurrentX + 1000
      Vsp.CurrentY = Vsp.CurrentY - 100
      Vsp = "    อนุมัติ"
      Vsp.CurrentX = TempX
      Vsp.CurrentY = TempY
      Call Vsp.DrawLine(Vsp.PageWidth - 4500, Vsp.CurrentY, Vsp.PageWidth - 4000, Vsp.CurrentY)
      TempX = Vsp.CurrentX
      Vsp.CurrentX = Vsp.PageWidth - 4000
      Vsp.CurrentY = Vsp.CurrentY - 100
      Vsp = "    ไม่อนุมัติ"
      Vsp.CurrentX = TempX
      Vsp.CurrentY = TempY
      
      Call Vsp.DrawLine(Vsp.CurrentX + 500, Vsp.CurrentY, Vsp.CurrentX + 500, Vsp.CurrentY + 250)
      Call Vsp.DrawLine(Vsp.CurrentX + 1000, Vsp.CurrentY, Vsp.CurrentX + 1000, Vsp.CurrentY + 250)
      
      Call Vsp.DrawLine(Vsp.PageWidth - 4500, Vsp.CurrentY, Vsp.PageWidth - 4500, Vsp.CurrentY + 250)
      Call Vsp.DrawLine(Vsp.PageWidth - 4000, Vsp.CurrentY, Vsp.PageWidth - 4000, Vsp.CurrentY + 250)
      
      Call Vsp.DrawLine(Vsp.CurrentX + 500, Vsp.CurrentY + 250, Vsp.CurrentX + 1000, Vsp.CurrentY + 250)
      Call Vsp.DrawLine(Vsp.PageWidth - 4500, Vsp.CurrentY + 250, Vsp.PageWidth - 4000, Vsp.CurrentY + 250)
      Vsp.FontBold = False
      
      Vsp.Paragraph = ""
      Vsp = "  ลงชื่อ    _________________________________________________  ผู้มีอำนาจลงนาม"
      
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   Rs.Close
   Set Rs = Nothing
   Set m_cData = Nothing
   
   genDoc = True
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
'   ErrorObj.RoutineName = RName
'   ErrorObj.ModuleName = MODULE_NAME
'   ErrorObj.LocalErrorMsg = "Run time error."
'   ErrorObj.SystemErrorMsg = mstrErrMsg
'   Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
   Set Rs = Nothing
End Function

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
   
'   Vsp.DrawPicture mcolParam("PICTURE"), "0.75in", "0.75in", "100%", "100%"
   
   Vsp.FontSize = 8
   dY(0) = Vsp.CurrentY
   Vsp.TextAlign = taLeftBottom
   strFormat = VSP_CalTable(TITLE_FORMAT, mdblWid, alngX)
   Vsp.TableBorder = tbNone
'   Call Vsp.AddTable(strFormat, "", "|หน้าที่|: " & Vsp.CurrentPage)
'   Call Vsp.AddTable(strFormat, "", "|วันที่พิมพ์|: " & DateToStringExt(mdteDate))
'   Call Vsp.AddTable(strFormat, "", "|เวลา|: " & Right$(DateToStringExtEx(mdteDate), 8))
   Vsp.CurrentY = dY(0)
   
   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   Vsp.TextAlign = taCenterBottom
   Vsp.Paragraph = MODULE_DESC
   Vsp.Paragraph = ""
   If mblnHeader Then
'      Call printHeader
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

'Private Function CheckCreditLimit() As Double
'Dim Doc As CDoItem
'Dim Rcp As CReceiptItem
'Dim Cn As CReceiptItem
'Dim Dn As CReceiptItem
'Dim RT As CReceiptItem
'Dim BLD As CBillingDiscount
'Dim m_Rs  As ADODB.Recordset
'Dim ItemCount As Long
'
'   CheckCreditLimit = 0
'
'   Set Doc = New CDoItem
'   Set m_Rs = New ADODB.Recordset
'   Doc.CUSTOMER_ID = m_cData.CUSTOMER_ID
'   Call Doc.QueryData(9, m_Rs, ItemCount)
'   If ItemCount > 0 Then
'      Call Doc.PopulateFromRS(9, m_Rs)
'   End If
'   CheckCreditLimit = CheckCreditLimit + Doc.TOTAL_PRICE
'   Set Doc = Nothing
'   Set m_Rs = Nothing
'
'   Set Rcp = New CReceiptItem
'   Set m_Rs = New ADODB.Recordset
'   Rcp.CUSTOMER_ID = m_cData.CUSTOMER_ID
'   Call Rcp.QueryData(3, m_Rs, ItemCount)
'   If ItemCount > 0 Then
'      Call Rcp.PopulateFromRS(3, m_Rs)
'   End If
'   CheckCreditLimit = CheckCreditLimit - Rcp.PAID_AMOUNT - Rcp.CASH_DISCOUNT
'   Set Rcp = Nothing
'   Set m_Rs = Nothing
'
'   Set Dn = New CReceiptItem
'   Set m_Rs = New ADODB.Recordset
'   Dn.CUSTOMER_ID = m_cData.CUSTOMER_ID
'   Dn.DOCUMENT_TYPE = 4
'   Call Dn.QueryData(6, m_Rs, ItemCount)
'   If ItemCount > 0 Then
'      Call Dn.PopulateFromRS(6, m_Rs)
'   End If
'   CheckCreditLimit = CheckCreditLimit + Dn.DEBIT_CREDIT_AMOUNT
'   Set Dn = Nothing
'   Set m_Rs = Nothing
'
'   Set Cn = New CReceiptItem
'   Set m_Rs = New ADODB.Recordset
'   Cn.CUSTOMER_ID = m_cData.CUSTOMER_ID
'   Cn.DOCUMENT_TYPE = 3
'   Call Cn.QueryData(6, m_Rs, ItemCount)
'   If ItemCount > 0 Then
'      Call Cn.PopulateFromRS(6, m_Rs)
'   End If
'   CheckCreditLimit = CheckCreditLimit - Cn.DEBIT_CREDIT_AMOUNT
'   Set Cn = Nothing
'   Set m_Rs = Nothing
'
'   Set RT = New CReceiptItem
'   Set m_Rs = New ADODB.Recordset
'   RT.CUSTOMER_ID = m_cData.CUSTOMER_ID
'   RT.DOCUMENT_TYPE = 18
'   Call RT.QueryData(6, m_Rs, ItemCount)
'   If ItemCount > 0 Then
'      Call RT.PopulateFromRS(6, m_Rs)
'   End If
'   CheckCreditLimit = CheckCreditLimit - RT.DEBIT_CREDIT_AMOUNT
'   Set RT = Nothing
'   Set m_Rs = Nothing
'
'   Set BLD = New CBillingDiscount
'   Set m_Rs = New ADODB.Recordset
'   BLD.CUSTOMER_ID = m_cData.CUSTOMER_ID
'   Call BLD.QueryData(4, m_Rs, ItemCount)
'   If ItemCount > 0 Then
'      Call BLD.PopulateFromRS(4, m_Rs)
'   End If
'   CheckCreditLimit = CheckCreditLimit - BLD.DISCOUNT_AMOUNT
'   Set BLD = Nothing
'   Set m_Rs = Nothing
'
'End Function

