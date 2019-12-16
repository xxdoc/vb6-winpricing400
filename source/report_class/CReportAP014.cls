VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAP014"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAP014"
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
Private m_Headers2 As CFieldList
Private m_Details1 As CFieldList
Private m_Details2 As CFieldList
Private m_Left As Double
Private m_YearSeqs As Collection

Private m_ExpenseTypes As Collection
Private m_PartGroup As Collection
Private mRecWhenRecCheque As Collection
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
   Set m_YearSeqs = New Collection
   Set mRecWhenRecCheque = New Collection
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
   Set mRecWhenRecCheque = Nothing
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
   Call Cf.SetFieldValue(2, "^", "ยกมา", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "หนี้เพิ่ม", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "หนี้ลด", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "คงเหลือ", "^")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "วันที่", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "เลขที่เอกสาร", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ยกมา", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ซื้อเชื่อ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "เพิ่มหนี้", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จ่ายชำระ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ลดหนี้", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "คงเหลือ", ">")
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
                                       vbCrLf & FromDate & ToDate & " ประเภทผู้ขาย " & EmptyToString(mcolParam("SUPPLIER_TYPE_NAME"), "N/A")
   
   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))
   
   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
      mstrErrMsg = "Error GetDateTime Error."
      Exit Function
   End If
   mdteDate = InternalDateToDate(strDate)
   initDoc = True
End Function

Private Sub GenerateBalanceItem(TXT1 As String, Total() As Double, Details As CFieldList, Bd As CBillingDoc, SellItems As Collection, PaidItems As Collection, DnBalances As Collection, CnBalances As Collection, RtBalances As Collection, DiscountBalances As Collection, mRecWhenRecCheque As Collection, BalAmount As Double, Total3() As Double)
Dim Ri As CReceiptItem
Dim Di As CSupItem
Dim J As Long
Dim KeyStr As String
Dim BodyCf As CReportField
Dim HeadCf As CReportField
Dim TempStr As String
Dim LeftAmount As Double
Dim TempRi As CReceiptItem
Dim TempRi1 As CReceiptItem
Dim Bdc As CBillingDiscount
Dim TempBd  As CBillingDoc
   
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
      ElseIf J = 4 Then
         KeyStr = Bd.SUPPLIER_ID & "-100-0-0"
         Set Di = GetSupItem(SellItems, KeyStr) 'ซื้อเชื่อ
         LeftAmount = LeftAmount + Di.TOTAL_PRICE
         
         KeyStr = Bd.SUPPLIER_ID & "-101-0-0"
         Set Di = GetSupItem(SellItems, KeyStr) 'ซื้อเชื่อ
         LeftAmount = LeftAmount + Di.TOTAL_PRICE
         
         KeyStr = Bd.SUPPLIER_ID & "-102-0-0"
         Set Di = GetSupItem(SellItems, KeyStr) 'ซื้อเชื่อ
         LeftAmount = LeftAmount + Di.TOTAL_PRICE
         
         KeyStr = Bd.SUPPLIER_ID & "-103-0-0"
         Set Di = GetSupItem(SellItems, KeyStr) 'ซื้อเชื่อ
         LeftAmount = LeftAmount + Di.TOTAL_PRICE
         
         Set TempRi = GetReceiptItem(DnBalances, Trim(str(Bd.SUPPLIER_ID)))  'เพิ่มหนี้
         LeftAmount = LeftAmount + TempRi.DEBIT_CREDIT_AMOUNT
         
         KeyStr = Bd.SUPPLIER_ID & "-8-0-3"
         Set Ri = GetReceiptItem(PaidItems, KeyStr) 'รับชำระ
         LeftAmount = LeftAmount - Ri.PAID_AMOUNT
         
         KeyStr = Trim(str(Bd.SUPPLIER_ID))
         Set TempBd = GetObject("CBillingDoc", mRecWhenRecCheque, KeyStr) 'รับชำระคิดตอนเมื่อรับบเช็ค
         LeftAmount = LeftAmount - TempBd.PAID_AMOUNT
         
         Set TempRi = GetReceiptItem(CnBalances, Trim(str(Bd.SUPPLIER_ID))) 'ลดหนี้
         LeftAmount = LeftAmount - TempRi.DEBIT_CREDIT_AMOUNT
         
         Set TempRi1 = GetReceiptItem(RtBalances, Trim(str(Bd.SUPPLIER_ID))) 'รับคืน
         LeftAmount = LeftAmount - TempRi1.DEBIT_CREDIT_AMOUNT
         
         TempStr = FormatNumber(LeftAmount)
         Total(4) = Total(4) + (LeftAmount)
      ElseIf (J = 5) Then 'ซื้อเชื่อ
         TempStr = FormatNumber(0)
         Total(5) = Total(5) + (0)
      ElseIf (J = 6) Then 'เพิ่มหนี้
         TempStr = FormatNumber(0)
         Total(6) = Total(6) + 0
      ElseIf (J = 7) Then 'รับชำระหนี้
         TempStr = FormatNumber(0)
         Total(7) = Total(7) + 0
      ElseIf (J = 8) Then 'ลดหนี้
         TempStr = FormatNumber(0)
         Total(8) = Total(8) + 0
      ElseIf (J = 9) Then 'คงเหลือ
         TempStr = FormatNumber(LeftAmount)
         Total(9) = Total(9) + LeftAmount
         Total3(9) = Total3(9) + LeftAmount
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
      ElseIf J = 3 Then
         TempStr = Bd.DOCUMENT_NO
      ElseIf J = 4 Then
         TempStr = FormatNumber(0)
      ElseIf (J = 5) Then
         If (Bd.DOCUMENT_TYPE = 100) Or (Bd.DOCUMENT_TYPE = 101) Or (Bd.DOCUMENT_TYPE = 102) Or (Bd.DOCUMENT_TYPE = 103) Then
            TempStr = FormatNumber(Bd.DO_TOTAL_PRICE + Bd.REVENUE_TOTAL_PRICE - Bd.SUM_DISCOUNT_AMOUNT)
            Total(5) = Total(5) + Bd.DO_TOTAL_PRICE + Bd.REVENUE_TOTAL_PRICE - Bd.SUM_DISCOUNT_AMOUNT
            LeftAmount = LeftAmount + Bd.DO_TOTAL_PRICE + Bd.REVENUE_TOTAL_PRICE - Bd.SUM_DISCOUNT_AMOUNT
         Else
            TempStr = FormatNumber(0)
            Total(5) = Total(5) + 0
            LeftAmount = LeftAmount + 0
         End If
      ElseIf (J = 6) Then 'เพิ่มหนี้
         If Bd.DOCUMENT_TYPE = 10 Then
            TempStr = FormatNumber(Bd.TOTAL_DEBIT_AMOUNT)
            Total(6) = Total(6) + Bd.TOTAL_DEBIT_AMOUNT
            LeftAmount = LeftAmount + Bd.TOTAL_DEBIT_AMOUNT
         Else
            TempStr = FormatNumber(0)
            Total(6) = Total(6) + 0
            LeftAmount = LeftAmount + 0
         End If
      ElseIf (J = 7) Then 'รับชำระหนี้
         If (Bd.DOCUMENT_TYPE = 8) Then
            TempStr = FormatNumber(Bd.SUM_PAID_AMOUNT)
            Total(7) = Total(7) + Bd.SUM_PAID_AMOUNT
            LeftAmount = LeftAmount - Bd.SUM_PAID_AMOUNT
         ElseIf Bd.DOCUMENT_TYPE = 15000 Then
            TempStr = FormatNumber(Bd.PAID_AMOUNT)
            Total(7) = Total(7) + Bd.PAID_AMOUNT
            LeftAmount = LeftAmount - Bd.PAID_AMOUNT
         Else
            TempStr = FormatNumber(0)
            Total(7) = Total(7) + 0
            LeftAmount = LeftAmount + 0
         End If
      ElseIf (J = 8) Then 'ลดหนี้
         If Bd.DOCUMENT_TYPE = 9 Then
            TempStr = FormatNumber(Bd.TOTAL_CREDIT_AMOUNT)   'Bd.CREDIT_AMOUNT
            Total(8) = Total(8) + Bd.TOTAL_CREDIT_AMOUNT
            LeftAmount = LeftAmount - Bd.TOTAL_CREDIT_AMOUNT
         ElseIf Bd.DOCUMENT_TYPE = 110 Then
            TempStr = FormatNumber(Bd.RETURN_AMOUNT)   'Bd.RETURN_AMOUNT
            Total(8) = Total(8) + Bd.RETURN_AMOUNT
            LeftAmount = LeftAmount - Bd.RETURN_AMOUNT
         Else
            TempStr = FormatNumber(0)
            Total(8) = Total(8) + 0
            LeftAmount = LeftAmount + 0
         End If
      ElseIf (J = 9) Then 'คงเหลือ
         TempStr = FormatNumber(LeftAmount)
         Total(9) = Total(9) + LeftAmount
      Else
         TempStr = ""
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
   BalAmount = LeftAmount
End Sub

Private Function FilterDocumentType(Bd As CBillingDoc) As Boolean
Dim TmpFlag As String

   FilterDocumentType = True
   TmpFlag = NoFlag2Null(Check2Flag(mcolParam("EFFECTIVE_FLAG")))

   If (Bd.DOCUMENT_TYPE = 8 And Not (Bd.CT_PAYMENT_TYPE = 1 Or Bd.CT_PAYMENT_TYPE = 2)) Then
      FilterDocumentType = False
   End If
End Function
Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim O As Object
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

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set Accounts = New Collection
   Call LoadAccount(Nothing, Accounts)
   
   Set DistinctAccounts = New Collection
   Call LoadBillingDocDistinctSup(Nothing, DistinctAccounts, mcolParam("SUPPLIER_CODE"), mcolParam("SUPPLIER_TYPE"), mcolParam("SUPPLIER_GRADE"))
   
   Set SellItems = New Collection
   Set PaidItems = New Collection
   Set CnBalances = New Collection
   Set DnBalances = New Collection
   Set RtBalances = New Collection
   Set DiscountBalances = New Collection
   
   NewDate = DateAdd("D", -1, mcolParam("FROM_DATE"))
   If NewDate > 0 Then
      Call LoadDnCnAmountBySupplier(Nothing, DnBalances, -1, NewDate, 10, 2) '1 ของเดิมใช้ date type = 1
      Call LoadDnCnAmountBySupplier(Nothing, CnBalances, -1, NewDate, 9, 2) '1
      Call LoadDnCnAmountBySupplier(Nothing, RtBalances, -1, NewDate, 110, 2) '1
      
      Call LoadSellPriceByDocTypeSubTypeReceiptTypeSup(Nothing, SellItems, -1, NewDate)
      Call LoadReceiptByDocTypeSubTypeReceiptTypeSupEx(Nothing, PaidItems, -1, NewDate)
      Call LoadReceiptWhenRecCheque(Nothing, mRecWhenRecCheque, -1, NewDate, 15000)        'ใบรับเมื่อได้รับเช็คแล้ว
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
   Next J
   
   Set TempBd = New CBillingDoc
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
      I = 0
      Set BodyCf = New CReportField
      For Each Bd In DistinctAccounts
         'Show supplier
         Call m_Details1.ClearField

         Set HeadCf = m_Headers1.Fields(1)
         SumWidth = HeadCf.Width
         Set HeadCf = m_Headers1.Fields(2)
         SumWidth = SumWidth + HeadCf.Width
         TempStr = Bd.SUPPLIER_CODE
         Call BodyCf.SetFieldValue(SumWidth, "^", TempStr, "^")
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(3)
         SumWidth = HeadCf.Width
         Set HeadCf = m_Headers1.Fields(4)
         SumWidth = SumWidth + HeadCf.Width
         TempStr = Bd.SUPPLIER_NAME
         Call BodyCf.SetFieldValue(SumWidth, "<", TempStr, "<")
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(5)
         SumWidth = HeadCf.Width
         Set HeadCf = m_Headers1.Fields(6)
         SumWidth = SumWidth + HeadCf.Width
         Set HeadCf = m_Headers1.Fields(7)
         SumWidth = SumWidth + HeadCf.Width
         Set HeadCf = m_Headers1.Fields(8)
         SumWidth = SumWidth + HeadCf.Width
         Set HeadCf = m_Headers1.Fields(9)
         SumWidth = SumWidth + HeadCf.Width
         TempStr = ""
         Call BodyCf.SetFieldValue(SumWidth, "<", TempStr, "<")
         Call m_Details1.AddField(BodyCf)

         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         'Show supplier

         For J = 1 To UBound(Total1)
            Total1(J) = 0
         Next J

         Call GenerateBalanceItem("***ยกมา***", Total1, m_Details1, Bd, SellItems, PaidItems, DnBalances, CnBalances, RtBalances, DiscountBalances, mRecWhenRecCheque, BalanceAmount, Total3)
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)

         Set TempBd = Nothing
         Set TempBd = New CBillingDoc
         TempBd.BILLING_DOC_ID = -1
         TempBd.SUPPLIER_ID = Bd.SUPPLIER_ID
         TempBd.FROM_DATE = mcolParam("FROM_DATE")
         TempBd.TO_DATE = mcolParam("TO_DATE")
         TempBd.OrderBy = 6
         TempBd.OrderType = mcolParam("ORDER_TYPE")
         TempBd.ItemSumFlagSup = True
         TempBd.FindPayType = True
         Call glbDaily.QueryBillingDoc(TempBd, Rs, iCount, IsOK, glbErrorLog)
         While Not Rs.EOF
            Call TempBd.PopulateFromRS(1, Rs)
            If FilterDocumentType(TempBd) Then
               I = I + 1

               Call GenerateDetail("", Total1, m_Details1, TempBd, I, BalanceAmount)
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
            End If
            Rs.MoveNext
         Wend

         For J = 1 To UBound(Total1)
            If J <> 9 Then
               Total2(J) = Total2(J) + Total1(J)
            Else
               Total2(J) = Total2(J) + BalanceAmount
            End If
         Next J

         Call GenerateFooter("รวม", "", "", Total1, m_Details1, BalanceAmount)
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
      Next Bd
      Set TempBd = Nothing

      Call GenerateFooter("สุทธิ", "", "", Total2, m_Details1, Total2(9))
      Call m_Details1.GetString(2, TempStr1, TempStr2)
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
      ElseIf J = 9 Then
         TempStr = FormatNumber(BalAmount)
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


