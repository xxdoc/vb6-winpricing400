VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAR003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAR003"
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
Private TempDate As Date
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
Private m_Headers2 As CFieldList
Private m_Details1 As CFieldList
Private m_Details2 As CFieldList
Private m_Details3 As CFieldList
Private m_Left As Double
Private m_YearSeqs As Collection

Private m_ExpenseTypes As Collection
Private m_PartGroup As Collection

Private Total4(100) As Double
Private Total5(100) As Double
Private Total6(100) As Double

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_TempCol = New Collection
   
   Set m_ExpenseTypes = New Collection
   Set m_PartGroup = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Headers2 = New CFieldList
   Set m_Details1 = New CFieldList
   Set m_Details2 = New CFieldList
   Set m_Details3 = New CFieldList
   Set m_YearSeqs = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_ExpenseTypes = Nothing
   Set m_PartGroup = Nothing
      
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Headers2 = Nothing
   Set m_Details1 = Nothing
   Set m_YearSeqs = Nothing
   Set m_Details2 = Nothing
   Set m_Details3 = Nothing
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

   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   Call m_Headers2.ClearField
   
   Set Cf = New CReportField

   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "หนี้เพิ่ม", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "หนี้ลด", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "คงเหลือ", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "DUE", "^")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "วันที่", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "เลขที่เอกสาร", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ขายเชื่อ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "เพิ่มหนี้", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "รับชำระ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ลดหนี้", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "คงเหลือ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จ่ายจริง", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "รวมเดือน", "^")
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "ขายเชื่อ", ">")
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "รับชำระ", ">")
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "คงเหลือ", ">")
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "AVG DUE", "^")
   Call m_Headers2.AddField(Cf)

   Set Cf = Nothing
 
 If mcolParam("SUM_MONTH") = 1 Then
   Call m_Headers2.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
 Else
   Call m_Headers0.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
         
   Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
  End If
   
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
   Vsp.ORIENTATION = orPortrait
   Vsp.MarginBottom = 1440
   Vsp.MarginFooter = 1440
   Vsp.MarginHeader = 1440
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 1440
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   FromDate = "จาก วันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
   ToDate = " ถึง วันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")
   
   MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & _
                                       vbCrLf & mcolParam("REPORT_NAME") & _
                                       vbCrLf & FromDate & ToDate
   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))
   
   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
      mstrErrMsg = "Error GetDateTime Error."
      Exit Function
   End If
   mdteDate = InternalDateToDate(strDate)
   initDoc = True
End Function

Private Sub GenerateBalanceItem(TXT1 As String, Total() As Double, Details As CFieldList, Bd As CBillingDoc, SellItems As Collection, PaidItems As Collection, DnBalances As Collection, CnBalances As Collection, RtBalances As Collection, DiscountBalances As Collection, BalAmount As Double, Total3() As Double)
Dim Ri As CReceiptItem
Dim Di As CDoItem
Dim J As Long
Dim KeyStr As String
Dim BodyCf As CReportField
Dim HeadCf As CReportField
Dim TempStr As String
Dim LeftAmount As Double
Dim TempRi As CReceiptItem
Dim TempRi1 As CReceiptItem
Dim Bdc As CBillingDiscount

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   J = 0
   LeftAmount = 0
   For Each HeadCf In m_Headers1.Fields
      J = J + 1
      If J = 1 Then
         TempStr = ""
      ElseIf J = 2 Then
         TempStr = TXT1
      ElseIf J = 3 Then
         TempStr = ""
      ElseIf (J = 4) Then 'ซื้อเชื่อหมู
         Set Bdc = GetBillingDiscount(DiscountBalances, Trim(str(Bd.ACCOUNT_ID)))
         
         KeyStr = Bd.ACCOUNT_ID & "-1-0-0"
         Set Di = GetDoItem(SellItems, KeyStr)
'         TempStr = FormatNumber(Di.TOTAL_PRICE - Bdc.DISCOUNT_AMOUNT)
'         Total(4) = Total(4) + (Di.TOTAL_PRICE - Bdc.DISCOUNT_AMOUNT)
         TempStr = FormatNumber(0)
         Total(4) = Total(4) + (0)
         LeftAmount = LeftAmount + (Di.TOTAL_PRICE - Bdc.DISCOUNT_AMOUNT)
      ElseIf (J = 5) Then 'เพิ่มหนี้
         Set TempRi = GetReceiptItem(DnBalances, Trim(str(Bd.ACCOUNT_ID)))
'         TempStr = FormatNumber(TempRi.DEBIT_CREDIT_AMOUNT)
'         Total(5) = Total(5) + TempRi.DEBIT_CREDIT_AMOUNT
         TempStr = FormatNumber(0)
         Total(5) = Total(5) + 0
         LeftAmount = LeftAmount + TempRi.DEBIT_CREDIT_AMOUNT
      ElseIf (J = 6) Then 'รับชำระหนี้
         KeyStr = Bd.ACCOUNT_ID & "-2-0-3"
         Set Ri = GetReceiptItem(PaidItems, KeyStr)
'         TempStr = FormatNumber(Ri.PAID_AMOUNT)
'         Total(6) = Total(6) + Ri.PAID_AMOUNT
         TempStr = FormatNumber(0)
         Total(6) = Total(6) + 0
         LeftAmount = LeftAmount - Ri.PAID_AMOUNT
      ElseIf (J = 7) Then 'ลดหนี้
         Set TempRi = GetReceiptItem(CnBalances, Trim(str(Bd.ACCOUNT_ID)))
         Set TempRi1 = GetReceiptItem(RtBalances, Trim(str(Bd.ACCOUNT_ID)))
         
         KeyStr = Bd.ACCOUNT_ID & "-2-0-3"
         Set Ri = GetReceiptItem(PaidItems, KeyStr)
         
'         TempStr = FormatNumber(TempRi.DEBIT_CREDIT_AMOUNT + Ri.CASH_DISCOUNT)
'         Total(7) = Total(7) + TempRi.DEBIT_CREDIT_AMOUNT + Ri.CASH_DISCOUNT
         TempStr = FormatNumber(0)
         Total(7) = Total(7) + 0
         LeftAmount = LeftAmount - (TempRi.DEBIT_CREDIT_AMOUNT + Ri.CASH_DISCOUNT + TempRi1.DEBIT_CREDIT_AMOUNT)
      ElseIf (J = 8) Then 'คงเหลือ
         TempStr = FormatNumber(LeftAmount)
         Total(8) = Total(8) + LeftAmount
         Total3(8) = Total3(8) + LeftAmount
         Total6(8) = LeftAmount
      ElseIf J = 9 Then
         TempStr = ""
      Else
         TempStr = ""
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
   BalAmount = LeftAmount
End Sub
Private Sub GenerateCustomerDetail(Details As CFieldList, Ac As CAccount)
Dim J As Long
Dim KeyStr As String
Dim BodyCf As CReportField
Dim HeadCf As CReportField
Dim TempStr As String
Dim SumWidth As Double
Dim temp_Headers As CFieldList

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   J = 0
   
   Set temp_Headers = New CFieldList
   If mcolParam("SUM_MONTH") = 1 Then
     Set temp_Headers = m_Headers2
   Else
     Set temp_Headers = m_Headers1
   End If
   For Each HeadCf In temp_Headers.Fields
         J = J + 1
         If J = 2 Then
             SumWidth = SumWidth + HeadCf.Width
             
            TempStr = Ac.CUSTOMER_CODE
            Call BodyCf.SetFieldValue(SumWidth, "<", TempStr, "<")
            Call m_Details1.AddField(BodyCf)
            SumWidth = 0
         ElseIf J = 4 Then
             SumWidth = SumWidth + HeadCf.Width
             
            TempStr = Ac.CUSTOMER_NAME
            Call BodyCf.SetFieldValue(SumWidth, "<", TempStr, "<")
            Call m_Details1.AddField(BodyCf)
            SumWidth = 0
         ElseIf J = 6 Then
             SumWidth = SumWidth + HeadCf.Width
             
             If mcolParam("SUM_MONTH") = 1 Then
               TempStr = "วงเงิน" & FormatNumber(Ac.CREDIT_LIMIT) & " บาท " & "   เครดิต " & Ac.CREDIT_DATE & " วัน "
               Call BodyCf.SetFieldValue(SumWidth, "<", TempStr, "<")
               Call m_Details1.AddField(BodyCf)
               SumWidth = 0
            End If
         ElseIf J = 9 Then
             SumWidth = SumWidth + HeadCf.Width
             
             If mcolParam("SUM_MONTH") <> 1 Then
               TempStr = "วงเงิน" & FormatNumber(Ac.CREDIT_LIMIT) & " บาท " & "   เครดิต " & Ac.CREDIT_DATE & " วัน "
               Call BodyCf.SetFieldValue(SumWidth, "<", TempStr, "<")
               Call m_Details1.AddField(BodyCf)
               SumWidth = 0
            End If
         Else
             SumWidth = SumWidth + HeadCf.Width
         End If
         
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call Details.AddField(BodyCf)
      Next HeadCf

   Set BodyCf = Nothing
End Sub
Private Sub GenerateBalanceItemSum(TXT1 As String, Total() As Double, Details As CFieldList, Bd As CBillingDoc, SellItems As Collection, PaidItems As Collection, DnBalances As Collection, CnBalances As Collection, RtBalances As Collection, DiscountBalances As Collection, BalAmount As Double, Total3() As Double)
Dim Ri As CReceiptItem
Dim Di As CDoItem
Dim J As Long
Dim KeyStr As String
Dim BodyCf As CReportField
Dim HeadCf As CReportField
Dim TempStr As String
Dim LeftAmount As Double
Dim TempRi As CReceiptItem
Dim TempRi1 As CReceiptItem
Dim Bdc As CBillingDiscount

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   J = 0
   LeftAmount = 0
   For Each HeadCf In m_Headers2.Fields
      J = J + 1
      If J = 1 Then
         TempStr = ""
      ElseIf J = 2 Then
         TempStr = TXT1
      ElseIf (J = 3) Then 'ซื้อเชื่อหมู
         Set Bdc = GetBillingDiscount(DiscountBalances, Trim(str(Bd.ACCOUNT_ID)))
         
         KeyStr = Bd.ACCOUNT_ID & "-1-0-0"
         Set Di = GetDoItem(SellItems, KeyStr)
         TempStr = FormatNumber(0)
         Total(4) = Total(4) + (0)
         LeftAmount = LeftAmount + (Di.TOTAL_PRICE - Bdc.DISCOUNT_AMOUNT)
         
         'เพิ่มหนี้
         Set TempRi = GetReceiptItem(DnBalances, Trim(str(Bd.ACCOUNT_ID)))
         LeftAmount = LeftAmount + TempRi.DEBIT_CREDIT_AMOUNT
      
      ElseIf (J = 4) Then 'รับชำระหนี้
         KeyStr = Bd.ACCOUNT_ID & "-2-0-3"
         Set Ri = GetReceiptItem(PaidItems, KeyStr)
         TempStr = FormatNumber(0)
         Total(6) = Total(6) + 0
         LeftAmount = LeftAmount - Ri.PAID_AMOUNT
         
         'ลดหนี้
         Set TempRi = GetReceiptItem(CnBalances, Trim(str(Bd.ACCOUNT_ID)))
         Set TempRi1 = GetReceiptItem(RtBalances, Trim(str(Bd.ACCOUNT_ID)))
         KeyStr = Bd.ACCOUNT_ID & "-2-0-3"
         Set Ri = GetReceiptItem(PaidItems, KeyStr)
         LeftAmount = LeftAmount - (TempRi.DEBIT_CREDIT_AMOUNT + Ri.CASH_DISCOUNT + TempRi1.DEBIT_CREDIT_AMOUNT)
      ElseIf (J = 5) Then 'คงเหลือ
         TempStr = FormatNumber(LeftAmount)
         Total(8) = Total(8) + LeftAmount
         Total3(8) = Total3(8) + LeftAmount
         Total6(8) = LeftAmount
      Else
         TempStr = ""
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
   BalAmount = LeftAmount
End Sub
Private Sub GenerateDetail(TXT1 As String, Total() As Double, Details As CFieldList, Bd As CBillingDoc, iCount As Long, BalAmount As Double)
Dim J As Long
Dim KeyStr As String
Dim BodyCf As CReportField
Dim HeadCf As CReportField
Dim TempStr As String
Dim LeftAmount As Double

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   'ปัดก่อนแล้วค่อยรวม
   Bd.DO_TOTAL_PRICE = Val(Format(Bd.DO_TOTAL_PRICE, "0.00"))
   Bd.REVENUE_TOTAL_PRICE = Val(Format(Bd.REVENUE_TOTAL_PRICE, "0.00"))
   Bd.SUM_DISCOUNT_AMOUNT = Val(Format(Bd.SUM_DISCOUNT_AMOUNT, "0.00"))
   Bd.TOTAL_DEBIT_AMOUNT = Val(Format(Bd.TOTAL_DEBIT_AMOUNT, "0.00"))
   Bd.SUM_PAID_AMOUNT = Val(Format(Bd.SUM_PAID_AMOUNT, "0.00"))
   Bd.TOTAL_CREDIT_AMOUNT = Val(Format(Bd.TOTAL_CREDIT_AMOUNT, "0.00"))
   Bd.SUM_CASH_DISCOUNT = Val(Format(Bd.SUM_CASH_DISCOUNT, "0.00"))
   Bd.RETURN_AMOUNT = Val(Format(Bd.RETURN_AMOUNT, "0.00"))
   
   J = 0
   LeftAmount = BalAmount
   For Each HeadCf In m_Headers1.Fields
      J = J + 1
      If J = 1 Then
         TempStr = iCount
      ElseIf J = 2 Then
         TempStr = DateToStringExtEx2(Bd.DOCUMENT_DATE)
         TempDate = Bd.DOCUMENT_DATE
      ElseIf J = 3 Then
         TempStr = Bd.DOCUMENT_NO
      ElseIf (J = 4) Then
         If (Bd.DOCUMENT_TYPE = 1) Then
            TempStr = FormatNumber(Bd.DO_TOTAL_PRICE + Bd.REVENUE_TOTAL_PRICE - Bd.SUM_DISCOUNT_AMOUNT)
            Total(4) = Total(4) + Bd.DO_TOTAL_PRICE + Bd.REVENUE_TOTAL_PRICE - Bd.SUM_DISCOUNT_AMOUNT
            Total4(4) = Total4(4) + Bd.DO_TOTAL_PRICE + Bd.REVENUE_TOTAL_PRICE - Bd.SUM_DISCOUNT_AMOUNT
            LeftAmount = LeftAmount + Bd.DO_TOTAL_PRICE + Bd.REVENUE_TOTAL_PRICE - Bd.SUM_DISCOUNT_AMOUNT
         Else
            TempStr = FormatNumber(0)
            Total(4) = Total(4) + 0
             Total4(4) = Total4(4) + 0
            LeftAmount = LeftAmount + 0
         End If
      ElseIf (J = 5) Then 'เพิ่มหนี้
         If Bd.DOCUMENT_TYPE = 4 Then
            TempStr = FormatNumber(Bd.TOTAL_DEBIT_AMOUNT)
            Total(5) = Total(5) + Bd.TOTAL_DEBIT_AMOUNT
            Total4(5) = Total4(5) + Bd.TOTAL_DEBIT_AMOUNT
            LeftAmount = LeftAmount + Bd.TOTAL_DEBIT_AMOUNT
         Else
            TempStr = FormatNumber(0)
            Total(5) = Total(5) + 0
            LeftAmount = LeftAmount + 0
         End If
      ElseIf (J = 6) Then 'รับชำระหนี้
         If (Bd.DOCUMENT_TYPE = 2) And (Bd.RECEIPT_TYPE = 3) Then
            TempStr = FormatNumber(Bd.SUM_PAID_AMOUNT)
            Total(6) = Total(6) + Bd.SUM_PAID_AMOUNT
            Total4(6) = Total4(6) + Bd.SUM_PAID_AMOUNT
            LeftAmount = LeftAmount - Bd.SUM_PAID_AMOUNT
         Else
            TempStr = FormatNumber(0)
            Total(6) = Total(6) + 0
            Total4(6) = Total4(6) + 0
            LeftAmount = LeftAmount + 0
         End If
      ElseIf (J = 7) Then 'ลดหนี้
         If Bd.DOCUMENT_TYPE = 3 Then
            TempStr = FormatNumber(Bd.TOTAL_CREDIT_AMOUNT)   'Bd.CREDIT_AMOUNT
            Total(7) = Total(7) + Bd.TOTAL_CREDIT_AMOUNT
            Total4(7) = Total4(7) + Bd.TOTAL_CREDIT_AMOUNT
            LeftAmount = LeftAmount - Bd.TOTAL_CREDIT_AMOUNT
         ElseIf Bd.DOCUMENT_TYPE = 18 Then
            TempStr = FormatNumber(Bd.RETURN_AMOUNT)   'Bd.RETURN_AMOUNT
            Total(7) = Total(7) + Bd.RETURN_AMOUNT
            LeftAmount = LeftAmount - Bd.RETURN_AMOUNT
         ElseIf (Bd.DOCUMENT_TYPE = 2) And (Bd.RECEIPT_TYPE = 3) Then
            TempStr = FormatNumber(Bd.SUM_CASH_DISCOUNT)
            Total(7) = Total(7) + Bd.SUM_CASH_DISCOUNT
            LeftAmount = LeftAmount - Bd.SUM_CASH_DISCOUNT
         Else
            TempStr = FormatNumber(0)
            Total(7) = Total(7) + 0
            LeftAmount = LeftAmount + 0
         End If
      ElseIf (J = 8) Then 'คงเหลือ
         TempStr = FormatNumber(LeftAmount)
         Total(8) = Total(8) + LeftAmount
      ElseIf J = 9 Then
         TempStr = ""
      Else
         TempStr = ""
      End If
      
       If mcolParam("SUM_MONTH") <> 1 Then
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call Details.AddField(BodyCf)
      End If
   Next HeadCf
      
   Set BodyCf = Nothing
   BalAmount = LeftAmount
End Sub

Private Function FilterDocumentType(Bd As CBillingDoc) As Boolean
   If Bd.DOCUMENT_TYPE = 1 Then
      FilterDocumentType = True
   ElseIf (Bd.DOCUMENT_TYPE = 2) And (Bd.RECEIPT_TYPE = 3) Then
      FilterDocumentType = True
   ElseIf (Bd.DOCUMENT_TYPE = 3) Then
      FilterDocumentType = True
   ElseIf (Bd.DOCUMENT_TYPE = 4) Then
      FilterDocumentType = True
   ElseIf (Bd.DOCUMENT_TYPE = 18) Then
      FilterDocumentType = True
   Else
      FilterDocumentType = False
   End If
End Function

Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim O As Object
Dim RName As String
Dim I As Long
Dim I2 As Long
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
Dim TempID As Long
Dim NewDate As Date
Dim Accounts As Collection
Dim DistinctAccounts As Collection
Dim Bd As CBillingDoc
Dim Ac As CAccount
Dim SumWidth As Double
Dim SellItems As Collection
Dim PaidItems As Collection
Dim DiscountBalances As Collection
Dim TempBd As CBillingDoc
Dim BalanceAmount As Double
Dim DnBalances As Collection
Dim CnBalances As Collection
Dim RtBalances As Collection
Dim HaveValue As Boolean
Dim k As Long
Dim LastMonth As Long
Dim lastDate As Date
   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set Accounts = New Collection
   Call LoadAccount(Nothing, Accounts)
   
   Set DistinctAccounts = New Collection
   Call LoadBillingDocDistinctAccount(Nothing, DistinctAccounts, mcolParam("CUSTOMER_CODE"), mcolParam("FROM_CUSTOMER_CODE"), mcolParam("TO_CUSTOMER_CODE"))
   
   Set SellItems = New Collection
   Set PaidItems = New Collection
   Set CnBalances = New Collection
   Set DnBalances = New Collection
   Set RtBalances = New Collection
   Set DiscountBalances = New Collection
   
   NewDate = DateAdd("D", -1, mcolParam("FROM_DATE"))
   If NewDate > 0 Then
      Call LoadDnCnAmountByAccount(Nothing, DnBalances, -1, NewDate, 4, 2) '1 ของเดิมใช้ date type = 1
      Call LoadDnCnAmountByAccount(Nothing, CnBalances, -1, NewDate, 3, 2) '1
      Call LoadDnCnAmountByAccount(Nothing, RtBalances, -1, NewDate, 18, 2) '1
      Call LoadDiscountAmountByAccount(Nothing, DiscountBalances, -1, NewDate)
      Call LoadSellPriceByDocTypeSubTypeReceiptTypeAcc(Nothing, SellItems, -1, NewDate)
      Call LoadReceiptByDocTypeSubTypeReceiptTypeAcc(Nothing, PaidItems, -1, NewDate)
   End If
   
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
      Total5(J) = 0
   Next J
   
   Set TempBd = New CBillingDoc
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
   
      I = 0
      I2 = 0
      Set BodyCf = New CReportField
      For Each Bd In DistinctAccounts
         'Show customer
         Set Ac = Accounts(Trim(str(Bd.ACCOUNT_ID)))
         
         Call m_Details1.ClearField
         Call GenerateCustomerDetail(m_Details2, Ac)
         HaveValue = False
         
         'Show customer
         
         For J = 1 To UBound(Total1)
            Total1(J) = 0
            Total6(J) = 0
         Next J
         
         If mcolParam("SUM_MONTH") = 1 Then
            Call GenerateBalanceItemSum("***ยกมา***", Total1, m_Details2, Bd, SellItems, PaidItems, DnBalances, CnBalances, RtBalances, DiscountBalances, BalanceAmount, Total3)
         Else
            Call GenerateBalanceItem("***ยกมา***", Total1, m_Details2, Bd, SellItems, PaidItems, DnBalances, CnBalances, RtBalances, DiscountBalances, BalanceAmount, Total3)
         End If

         If BalanceAmount <> 0 Then
            HaveValue = True
         End If
         
         k = 0
         lastDate = -1
         Set TempBd = Nothing
         Set TempBd = New CBillingDoc
         TempBd.BILLING_DOC_ID = -1
         TempBd.ACCOUNT_ID = Bd.ACCOUNT_ID
         TempBd.FROM_DATE = mcolParam("FROM_DATE")
         TempBd.TO_DATE = mcolParam("TO_DATE")
         TempBd.CUSTOMER_TYPE = mcolParam("CUSTOMER_TYPE")
         TempBd.CUSTOMER_GRADE = mcolParam("CUSTOMER_GRADE")
         TempBd.OrderBy = 6
         TempBd.OrderType = mcolParam("ORDER_TYPE")
         TempBd.ItemSumFlag = True
         Call glbDaily.QueryBillingDoc(TempBd, Rs, iCount, IsOK, glbErrorLog)
         While Not Rs.EOF
            Call TempBd.PopulateFromRS(1, Rs)
            If FilterDocumentType(TempBd) Then
               I = I + 1
               k = k + 1
               
               If mcolParam("SUM_MONTH") <> 1 And mcolParam("SUM_MONTH_DETAIL") <> 1 Then 'แบบปรกติ
                  Call GenerateDetail("", Total1, m_Details3, TempBd, I, BalanceAmount)
               ElseIf mcolParam("SUM_MONTH") = 1 Then
                  If I = 1 Then
                     LastMonth = Month(TempBd.DOCUMENT_DATE)
                     lastDate = TempBd.DOCUMENT_DATE
                  End If
                  If LastMonth <> Month(TempBd.DOCUMENT_DATE) Then
                     I2 = I2 + 1
                     Call GenerateSumMonth(m_Details3, lastDate, I2)
                     LastMonth = Month(TempBd.DOCUMENT_DATE)
                     lastDate = TempBd.DOCUMENT_DATE
                     For J = 1 To UBound(Total4)
                        Total4(J) = 0
                        Total5(J) = 0
                     Next J
                     Call GenerateDetail("", Total1, m_Details3, TempBd, I, BalanceAmount)
                  Else
                      If mcolParam("SUM_MONTH_DETAIL") <> 1 Then
                        Call GenerateDetail("", Total1, m_Details3, TempBd, I, BalanceAmount)
                     End If
                  End If
               ElseIf mcolParam("SUM_MONTH_DETAIL") = 1 Then
                  
                  If I = 1 Then
                     LastMonth = Month(TempBd.DOCUMENT_DATE)
                     lastDate = TempBd.DOCUMENT_DATE
                  End If
                  If LastMonth <> Month(TempBd.DOCUMENT_DATE) Then
                     I2 = I2 + 1
                     Call GenerateSumMonth(m_Details3, lastDate, I2)
                     LastMonth = Month(TempBd.DOCUMENT_DATE)
                     lastDate = TempBd.DOCUMENT_DATE
                     For J = 1 To UBound(Total4)
                        Total4(J) = 0
                        Total5(J) = 0
                     Next J
                     Call GenerateDetail("", Total1, m_Details3, TempBd, I, BalanceAmount)
                  Else
                     Call GenerateDetail("", Total1, m_Details3, TempBd, I, BalanceAmount)
                  End If
               End If
                                 
               If BalanceAmount <> 0 Then
                  HaveValue = True
               End If
               
               If k = 1 And HaveValue Then
                  Call m_Details1.GetString(2, TempStr1, TempStr2)
                  strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                  Call Vsp.AddTable(strFormat, "", TempStr2)
                  
                  Call m_Details2.GetString(2, TempStr1, TempStr2)
                  strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                  Call Vsp.AddTable(strFormat, "", TempStr2)
               End If
               
               Call m_Details3.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
               
                If TempBd.DOCUMENT_TYPE = 2 Then
                     Call GenerateBillNoRcp(m_Details3, TempBd.BILLING_DOC_ID)
                  End If
            End If
            Rs.MoveNext
         Wend

         If k = 0 And HaveValue Then
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
            
            Call m_Details2.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
         End If
         
         For J = 1 To UBound(Total1)
            If J <> 8 Then
               Total2(J) = Total2(J) + Total1(J)
            Else
               Total2(J) = Total2(J) + BalanceAmount
            End If
         Next J
         If HaveValue Then
            If mcolParam("SUM_MONTH") = 1 Then
               I2 = I2 + 1
               Call GenerateSumMonth(m_Details3, lastDate, I2)
               Call GenerateFooter2("รวม", "", "", Total1, m_Details3, BalanceAmount)
               For J = 1 To UBound(Total4)
                  Total4(J) = 0
                  Total5(J) = 0
               Next J
               I = 0
            ElseIf mcolParam("SUM_MONTH_DETAIL") = 1 Then
                 I2 = I2 + 1
               Call GenerateSumMonth(m_Details3, lastDate, I2)
                Call GenerateFooter("รวม", "", "", Total1, m_Details3, BalanceAmount)
               For J = 1 To UBound(Total4)
                  Total4(J) = 0
                  Total5(J) = 0
               Next J
               I = 0
            Else
               Call GenerateFooter("รวม", "", "", Total1, m_Details3, BalanceAmount)
            End If
            
            Call m_Details3.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
         End If
      Next Bd
      Set TempBd = Nothing

 If mcolParam("SUM_MONTH") <> 1 Then
      Call GenerateFooter("ยกมาสุทธิ", "", "", Total3, m_Details1, Total3(8))
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False

      Call GenerateFooter("สุทธิ", "", "", Total2, m_Details1, Total2(8))
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
   End If

      Set BodyCf = Nothing
   Vsp.EndDoc
   Call CloseExportFile(Vsp)

   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Set CnBalances = Nothing
   Set DnBalances = Nothing
   Set RtBalances = Nothing
   Set Accounts = Nothing
   Set DistinctAccounts = Nothing
   Set SellItems = Nothing
   Set PaidItems = Nothing
   
   genDoc = True
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
   Set Rs = Nothing
End Function

Private Sub GenerateFooter(TXT1 As String, TXT2 As String, TXT3 As String, Tot() As Double, Details As CFieldList, BalAmount As Double)
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
      ElseIf J = 8 Then
         TempStr = FormatNumber(BalAmount)
      ElseIf J = 9 Then
         TempStr = ""
      Else
         TempStr = FormatNumber(Tot(J))
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub
Private Sub GenerateFooter2(TXT1 As String, TXT2 As String, TXT3 As String, Tot() As Double, Details As CFieldList, BalAmount As Double)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim J As Long
Dim TempStr As String

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   J = 0
   For Each HeadCf In m_Headers2.Fields
      J = J + 1
      If J = 1 Then
         TempStr = ""
      ElseIf J = 2 Then
         TempStr = TXT1
      ElseIf J = 3 Then
         TempStr = FormatNumber(Tot(4))
      ElseIf J = 4 Then
         TempStr = FormatNumber(Tot(6))
      ElseIf J = 5 Then
         TempStr = FormatNumber(Total6(8))
      ElseIf J = 6 Then
         TempStr = "" 'FormatNumber(0)
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
End Sub

Private Sub VSP_StartDoc()
'This event occur when VSPrinter.StartDoc is used and used to initialize some information before generating printed document
End Sub
Private Sub GenerateBillNoRcp(Details As CFieldList, RcpID As Long)
Dim Ri As CReceiptItem
Dim J As Long
Dim KeyStr As String
Dim BodyCf As CReportField
Dim HeadCf As CReportField
Dim TempStr As String
Dim TempStr1 As String
Dim TempStr2 As String
Dim alngX() As Long
Dim strFormat As String
   
   Dim m_Rs As ADODB.Recordset
   Dim ItemCount As Long
   Set m_Rs = New ADODB.Recordset
   Set Ri = New CReceiptItem
   
   Ri.BILLING_DOC_ID = RcpID
   Call Ri.QueryData(15, m_Rs, ItemCount)
   
   While Not (m_Rs.EOF)
      Call Ri.PopulateFromRS(15, m_Rs)
      Call Details.ClearField
      Set BodyCf = New CReportField
   
      J = 0
      For Each HeadCf In m_Headers1.Fields
         J = J + 1
         If J = 1 Then
            TempStr = ""
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call Details.AddField(BodyCf)
         ElseIf J = 2 Then
            TempStr = "***" & DateToStringExtEx2(Ri.BILL_DATE)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, "<")
            Call Details.AddField(BodyCf)
         ElseIf J = 3 Then
            TempStr = "***" & Ri.BILL_NO
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, ">")
            Call Details.AddField(BodyCf)
         ElseIf J = 6 Then
            TempStr = "***" & FormatNumber(Ri.PAID_AMOUNT)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, ">")
            Call Details.AddField(BodyCf)
         ElseIf J = 9 Then
            TempStr = DateDiff("D", Ri.BILL_DATE, TempDate)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, "^")
            Call Details.AddField(BodyCf)
            Total5(9) = Total5(9) + Val(TempStr)
            Total5(1) = Total5(1) + 1 'นับจำนวนที่ ถูกคำนวน
         Else
            TempStr = ""
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call Details.AddField(BodyCf)
         End If
      
         
      Next HeadCf
         
      Set BodyCf = Nothing
      
      If mcolParam("SUM_MONTH") <> 1 Then
         Call Details.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
      End If
      
      m_Rs.MoveNext
   Wend
   
   Set m_Rs = Nothing
   Set Ri = Nothing
End Sub

Private Sub GenerateSumMonth(Details As CFieldList, Dt As Date, iCount As Long)
Dim Ri As CReceiptItem
Dim J As Long
Dim KeyStr As String
Dim BodyCf As CReportField
Dim HeadCf As CReportField
Dim TempStr As String
Dim TempStr1 As String
Dim TempStr2 As String
Dim alngX() As Long
Dim strFormat As String
   
      Call Details.ClearField
      Set BodyCf = New CReportField
       'For Each HeadCf In m_Headers1.Fields
      J = 0
      If mcolParam("SUM_MONTH") = 1 Then
         For Each HeadCf In m_Headers2.Fields
            J = J + 1
            If J = 1 Then
               TempStr = iCount
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
               Call Details.AddField(BodyCf)
            ElseIf J = 2 Then
               TempStr = Format(Month(Dt), "00") & "/" & Format(Year(Dt), "0000")
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
               Call Details.AddField(BodyCf)
            ElseIf J = 3 Then
               TempStr = FormatNumber(Total4(4), 2)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, ">")
               Call Details.AddField(BodyCf)
            ElseIf J = 4 Then
               TempStr = FormatNumber(Total4(6), 2)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, ">")
               Call Details.AddField(BodyCf)
            ElseIf J = 5 Then
               TempStr = FormatNumber(Total6(8) + Total4(4) - Total4(6), 2)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, ">")
               Call Details.AddField(BodyCf)
               Total6(8) = Total6(8) + Total4(4) - Total4(6)
            ElseIf J = 6 Then
               TempStr = FormatNumber(MyDiffEx(Total5(9), Total5(1)), 0)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, "^")
               Call Details.AddField(BodyCf)
            End If
         Next HeadCf
         
      Call Details.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Call Vsp.AddTable(strFormat, "", TempStr2)

   ElseIf mcolParam("SUM_MONTH_DETAIL") = 1 Then
         For Each HeadCf In m_Headers1.Fields
            J = J + 1
            If J = 1 Then
               TempStr = ""
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
               Call Details.AddField(BodyCf)
            ElseIf J = 2 Then
               TempStr = "รวมเดือน " & Format(Month(Dt), "00") & "/" & Format(Year(Dt), "0000")
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
               Call Details.AddField(BodyCf)
            ElseIf J = 4 Then
               TempStr = FormatNumber(Total4(4), 2)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, ">")
               Call Details.AddField(BodyCf)
            ElseIf J = 5 Then
               TempStr = FormatNumber(Total4(5), 2)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, ">")
               Call Details.AddField(BodyCf)
            ElseIf J = 6 Then
               TempStr = FormatNumber(Total4(6), 2)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, ">")
               Call Details.AddField(BodyCf)
            ElseIf J = 7 Then
               TempStr = FormatNumber(Total4(7), 2)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, ">")
               Call Details.AddField(BodyCf)
            ElseIf J = 8 Then
               TempStr = FormatNumber(Total6(8) + Total4(4) - Total4(6), 2)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, ">")
               Call Details.AddField(BodyCf)
               Total6(8) = Total6(8) + Total4(4) - Total4(6)
            ElseIf J = 9 Then
               TempStr = FormatNumber(MyDiffEx(Total5(9), Total5(1)), 0)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, "^")
               Call Details.AddField(BodyCf)
            Else
               TempStr = ""
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, ">")
               Call Details.AddField(BodyCf)
            End If
         Next HeadCf
         
      Vsp.FontBold = True
      Call Details.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False

End If
      Set BodyCf = Nothing
      

      
End Sub
