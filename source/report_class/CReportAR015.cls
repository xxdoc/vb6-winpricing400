VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAR015"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAR015"
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

Private Const TITLE_SIZE = 12
Private Const HEADER_SIZE = 10
Private Const DETAIL_SIZE = 10

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
Private m_Left As Double
Private m_BankAccounts As Collection

Private m_CashTranAmounts As Collection
Private m_CashTranAmount2s As Collection
Private m_CashTranAmount3s As Collection
Private m_Banks As Collection
Private m_BankBranchs As Collection
Private m_ReceiptItemAmounts As Collection

Private m_BalanceCashCheque As Collection
Private m_Payin As Collection

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
   Set m_BankAccounts = New Collection
   Set m_CashTranAmounts = New Collection
   Set m_CashTranAmount2s = New Collection
   Set m_CashTranAmount3s = New Collection
   Set m_Banks = New Collection
   Set m_BankBranchs = New Collection
   Set m_ReceiptItemAmounts = New Collection
   
   Set m_BalanceCashCheque = New Collection
   Set m_Payin = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_BankAccounts = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
   Set m_CashTranAmounts = Nothing
   Set m_CashTranAmount2s = Nothing
   Set m_CashTranAmount3s = Nothing
   Set m_Banks = Nothing
   Set m_BankBranchs = Nothing
   Set m_ReceiptItemAmounts = Nothing
   
   Set m_BalanceCashCheque = Nothing
   Set m_Payin = Nothing
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
Dim ct As CCashTran

   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField

   Set Cf = New CReportField
   
   '===
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "วันที่", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "รหัสลูกค้า", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(5, "^", "ชื่อลูกค้า", "<")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "รับชำระ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "เงินสด ยกมา", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "เงินสด + เช็ค", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "นำเข้า ธนาคาร", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "โอน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "เงินสด ยกไป", ">")
   Call m_Headers1.AddField(Cf)
   
   For Each ct In m_BankAccounts
      Call Cf.SetFieldValue(2, "^", ct.GetFieldValue("ACCOUNT_NAME"), ">")
      Call m_Headers1.AddField(Cf)
   Next ct
   
   Call Cf.SetFieldValue(2, "^", "เพิ่มหนี้ ลดหนี้", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "ส่วนต่าง รับชำระ", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "ค่าธรรมเนียม", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "รวม", ">")
   Call m_Headers1.AddField(Cf)
   
   
   Call Cf.SetFieldValue(12, "^", "", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "เงินสด ยกมา", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "เงินสด + เช็ค", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "นำเข้า ธนาคาร", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "โอน", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "เงินสด ยกไป", ">")
   Call m_Headers0.AddField(Cf)
   
   For Each ct In m_BankAccounts
      Call Cf.SetFieldValue(2, "^", ct.GetFieldValue("ACCOUNT_NAME"), ">")
      Call m_Headers0.AddField(Cf)
   Next ct
   
   Call Cf.SetFieldValue(2, "^", "เพิ่มหนี้ ลดหนี้", ">")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "ส่วนต่าง รับชำระ", ">")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "ค่าธรรมเนียม", ">")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "รวม", ">")
   Call m_Headers0.AddField(Cf)
   
   Set Cf = Nothing

'  Call m_Headers0.GetString(1, TempStr1, TempStr1_1)
'   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'   strPrint = TempStr1_1
'   Call Vsp.AddTable(strFormat, "", strPrint)

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
   Vsp.MarginLeft = 300
   Vsp.MarginRight = 300
   Vsp.MarginTop = 700
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   Call LoadBalanceCashCheque(Nothing, m_BalanceCashCheque, , DateAdd("D", -1, mcolParam("FROM_DATE")))        ' ยอดยกมาของ เงินสดในมือ
   
   Call LoadDistinctAccountInCashTran(Nothing, m_BankAccounts, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), "I")          ' BankAccount
   Call LoadSumCashTranAmount(Nothing, m_CashTranAmounts, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), "I")            '
   Call LoadReceiptByCustomerDate(Nothing, m_ReceiptItemAmounts, mcolParam("FROM_DATE"), mcolParam("TO_DATE"))          ' จำนวนเงินหักหนี้
   
   Call LoadSumCashTranAmountByCustDate(Nothing, m_CashTranAmount2s, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), "I")
   Call LoadSumCashTranAmountByCustDate2(Nothing, m_CashTranAmount3s, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), "I")
   
   Call LoadPayinByCustDateAccount(Nothing, m_Payin, mcolParam("FROM_DATE"), mcolParam("TO_DATE"))
   
   
   FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
   ToDate = " ถึง วันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")
   
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

Private Function GetItemFromListIndex(TempCol As Collection, Ind As Long) As Object
   Set GetItemFromListIndex = TempCol(Ind)
End Function

Private Function GetKey(ct As CCashTran) As String
   GetKey = DateToStringInt(ct.GetFieldValue("TX_DATE"))
End Function

Private Sub GenerateHeader(TXT1 As String, TXT2 As String, TXT3 As String, Details As CFieldList)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim J As Long
Dim TempStr As String

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   J = 0
   For Each HeadCf In m_Headers0.Fields
      J = J + 1
      If J = 1 Then
         TempStr = TXT1
      ElseIf J = 2 Then
         TempStr = TXT2
      ElseIf J = 3 Then
         TempStr = TXT3
      Else
         TempStr = ""
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.Align)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub

Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim RName As String
Dim I As Long
Dim J As Long
Dim k As Long
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
Dim Width As Double
Dim TempRs As ADODB.Recordset
Dim Hc As CReportField
Dim BalanceCt As CCashTran
Dim Payin As CCashTran
Dim ct As CCashTran
Dim Ct2 As CCashTran
Dim Ct3 As CCashTran
Dim Ct4 As CCashTran
Dim TempCt As CCashTran
Dim Ba As CCashTran
Dim Sum1 As Double
Dim Ri As CReceiptItem
Dim Diff  As Double
Dim FeeAmount As Double
Dim BalanceCashAmount As Double
 
'   m_ApSupAmounts
'   m_ApSupPaids
   
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
   
       Set BodyCf = New CReportField
      Set ct = New CCashTran
      Call ct.SetFieldValue("CASH_TRAN_ID", -1)
      Call ct.SetFieldValue("FROM_DATE", mcolParam("FROM_DATE"))
      Call ct.SetFieldValue("TO_DATE", mcolParam("TO_DATE"))
      Call ct.SetFieldValue("TX_TYPE", "I")
      Call ct.SetFieldValue("CUSTOMER_CODE", mcolParam("CUSTOMER_CODE"))
      Call ct.SetFieldValue("FROM_CUSTOMER_CODE", mcolParam("FROM_CUSTOMER_CODE"))
      Call ct.SetFieldValue("TO_CUSTOMER_CODE", mcolParam("TO_CUSTOMER_CODE"))
      Call ct.SetFieldValue("ORDER_BY", mcolParam("ORDER_BY"))
      Call ct.SetFieldValue("ORDER_TYPE", mcolParam("ORDER_TYPE"))
      Call ct.QueryData(4, Rs, iCount)
      
      If Not Rs.EOF Then
         Call ct.PopulateFromRS(4, Rs)
         PrevKey1 = GetKey(ct)
      End If
      
      While Not Rs.EOF
         I = I + 1
         Call ct.PopulateFromRS(4, Rs)
         If PrevKey1 <> GetKey(ct) Then
            Call GenerateFooter("รวม", "", "", Total1, m_Details1, 0)
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False

            For J = 1 To UBound(Total1)
               Total1(J) = 0
            Next J
         End If
         PrevKey1 = GetKey(ct)
         
         Call m_Details1.ClearField
         k = 0
         Sum1 = 0
         FeeAmount = 0
         
         k = k + 1
          Set Hc = m_Headers1.Fields(k)
         TempStr = I
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         k = k + 1
          Set Hc = m_Headers1.Fields(k)
         TempStr = DateToStringExtEx2(ct.GetFieldValue("TX_DATE"))
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         k = k + 1
          Set Hc = m_Headers1.Fields(k)
         TempStr = ct.GetFieldValue("CUSTOMER_CODE")
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         k = k + 1
          Set Hc = m_Headers1.Fields(k)
         TempStr = ct.GetFieldValue("CUSTOMER_NAME")
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set Ri = GetReceiptItem(m_ReceiptItemAmounts, DateToStringInt(ct.GetFieldValue("TX_DATE")) & "-" & ct.GetFieldValue("CUSTOMER_ID"))
         k = k + 1
          Set Hc = m_Headers1.Fields(k)
         TempStr = FormatNumber(Ri.PAID_AMOUNT)
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + (Ri.PAID_AMOUNT)
         Total2(k) = Total2(k) + (Ri.PAID_AMOUNT)
         

         Set BalanceCt = GetObject("CCashTran", m_BalanceCashCheque, Trim(str(I)), False)
         If Not (BalanceCt Is Nothing) Then
            BalanceCashAmount = BalanceCt.GetFieldValue("AMOUNT")
         End If
         k = k + 1
         Set Hc = m_Headers1.Fields(k)
         TempStr = FormatNumber(BalanceCashAmount)
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set Ct2 = GetCashTran(m_CashTranAmount2s, ct.GetFieldValue("CUSTOMER_ID") & "-3-" & DateToStringInt(ct.GetFieldValue("TX_DATE"))) 'เช็ค
         Set Ct3 = GetCashTran(m_CashTranAmount2s, ct.GetFieldValue("CUSTOMER_ID") & "-1-" & DateToStringInt(ct.GetFieldValue("TX_DATE"))) 'เงินสด
         Set Ct4 = GetCashTran(m_CashTranAmount3s, ct.GetFieldValue("CUSTOMER_ID") & "-2-" & DateToStringInt(ct.GetFieldValue("TX_DATE"))) 'เงินโอน
         
         Diff = Ri.PAID_AMOUNT - (Ct2.GetFieldValue("AMOUNT") + Ct3.GetFieldValue("AMOUNT") + Ct4.GetFieldValue("AMOUNT"))
         
         FeeAmount = FeeAmount + Ct3.GetFieldValue("FEE_AMOUNT") + Ct2.GetFieldValue("FEE_AMOUNT")
         
         k = k + 1                                                                                                   ' เงินสด
          Set Hc = m_Headers1.Fields(k)
         TempStr = FormatNumber(Ct3.GetFieldValue("AMOUNT") + Ct2.GetFieldValue("AMOUNT") - (Ct3.GetFieldValue("FEE_AMOUNT") + Ct2.GetFieldValue("FEE_AMOUNT")))
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + (Ct3.GetFieldValue("AMOUNT") + Ct2.GetFieldValue("AMOUNT") - (Ct3.GetFieldValue("FEE_AMOUNT") + Ct2.GetFieldValue("FEE_AMOUNT")))
         Total2(k) = Total2(k) + (Ct3.GetFieldValue("AMOUNT") + Ct2.GetFieldValue("AMOUNT") - (Ct3.GetFieldValue("FEE_AMOUNT") + Ct2.GetFieldValue("FEE_AMOUNT")))
         Sum1 = Sum1 + Ct3.GetFieldValue("AMOUNT") + Ct2.GetFieldValue("AMOUNT") - (Ct3.GetFieldValue("FEE_AMOUNT") + Ct2.GetFieldValue("FEE_AMOUNT"))
         BalanceCashAmount = BalanceCashAmount + Ct3.GetFieldValue("AMOUNT") + Ct2.GetFieldValue("AMOUNT")
         
         'Set BalanceCt = GetCashTran(m_BalanceCashCheque, Trim(Str(I)))
         k = k + 1
         Set Hc = m_Headers1.Fields(k)                                                 ' PAY IN  เข้าธนาคาร
         TempStr = FormatNumber(0)
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + (0)
         Total2(k) = Total2(k) + (0)
         BalanceCashAmount = BalanceCashAmount - 0
         
         k = k + 1
          Set Hc = m_Headers1.Fields(k)                                                ' เงินโอน
         TempStr = FormatNumber(Ct4.GetFieldValue("AMOUNT") - Ct4.GetFieldValue("FEE_AMOUNT"))
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + (Ct4.GetFieldValue("AMOUNT") - Ct4.GetFieldValue("FEE_AMOUNT"))
         Total2(k) = Total2(k) + (Ct4.GetFieldValue("AMOUNT") - Ct4.GetFieldValue("FEE_AMOUNT"))
         Sum1 = Sum1 + Ct4.GetFieldValue("AMOUNT") - Ct4.GetFieldValue("FEE_AMOUNT")
         FeeAmount = FeeAmount + Ct4.GetFieldValue("FEE_AMOUNT")
         
         k = k + 1
         Set Hc = m_Headers1.Fields(k)
         TempStr = FormatNumber(BalanceCashAmount)
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + (BalanceCashAmount)
         Total2(k) = Total2(k) + (BalanceCashAmount)
         
         For Each Ba In m_BankAccounts
            Set TempCt = GetCashTran(m_CashTranAmounts, ct.GetFieldValue("CUSTOMER_ID") & "-" & Ba.GetFieldValue("BANK_ACCOUNT") & "-" & DateToStringInt(ct.GetFieldValue("TX_DATE")))
            
            k = k + 1
             Set Hc = m_Headers1.Fields(k)
            TempStr = FormatNumber(TempCt.GetFieldValue("NET_AMOUNT"))
            Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(k) = Total1(k) + (TempCt.GetFieldValue("NET_AMOUNT"))
            Total2(k) = Total2(k) + (TempCt.GetFieldValue("NET_AMOUNT"))
            
         Next Ba
         
         k = k + 1                                 ' เพิ่มหนี้ลดหนี้
         Set Hc = m_Headers1.Fields(k)
         TempStr = FormatNumber(0)
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + (0)
         Total2(k) = Total2(k) + (0)
         
         k = k + 1
         Set Hc = m_Headers1.Fields(k)
         TempStr = FormatNumber(Diff)
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + (Diff)
         Total2(k) = Total2(k) + (Diff)
         
         k = k + 1
         Set Hc = m_Headers1.Fields(k)
         TempStr = FormatNumber(FeeAmount)
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + (FeeAmount)
         Total2(k) = Total2(k) + (FeeAmount)
         
         k = k + 1
         Set Hc = m_Headers1.Fields(k)
         TempStr = FormatNumber(Sum1)
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + (Sum1)
         Total2(k) = Total2(k) + (Sum1)

         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         
         ' Generate PayIn
         Call GeneratePayin(ct.GetFieldValue("TX_DATE"), ct.GetFieldValue("CUSTOMER_ID"), BalanceCashAmount, Total1, Total2)
         'BalanceCashAmount
         
         Rs.MoveNext
      Wend

      Call GenerateFooter("รวม", "", "", Total1, m_Details1, 0)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False

      Call GenerateFooter("สุทธิ", "", "", Total2, m_Details1, 0)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False

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
      ElseIf J = 3 Or J = 4 Then
         TempStr = ""
      ElseIf J = 6 Or J = 10 Then
         TempStr = ""
      Else
         TempStr = FormatNumber(Tot(J))
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub

Private Sub GenerateBkTr(TXT1 As String, TXT2 As String, TXT3 As String, Tot() As Double, Details As CFieldList, DateCount As Long)
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
Private Sub GeneratePayin(TxDate As Date, CustomerID As Long, BalanceCashAmount As Double, Total1() As Double, Total2() As Double)
Dim ct As CCashTran
Dim k As Long
Dim TempCt As CCashTran
Dim Ba As CCashTran
Dim Use As Long
Dim Hc As CReportField
Dim TempStr As String
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim alngX() As Long
Dim strFormat As String
Dim I As Long
   Set BodyCf = New CReportField
   
   Use = 0
   For Each ct In m_Payin
      
      If ct.GetFieldValue("TX_DATE") = TxDate And ct.GetFieldValue("CUSTOMER_ID") = CustomerID Then
         k = 0
         Call m_Details1.ClearField
         Use = Use + 1
         k = k + 1
         Set Hc = m_Headers0.Fields(k)
         TempStr = ""
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         k = 5
         k = k + 1
         Set Hc = m_Headers0.Fields(k - 4)
         TempStr = FormatNumber(BalanceCashAmount)
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         k = k + 1                                                                                                   ' เงินสด
         Set Hc = m_Headers0.Fields(k - 4)
         TempStr = FormatNumber(0)
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         k = k + 1
         Set Hc = m_Headers0.Fields(k - 4)                                                 ' PAY IN  เข้าธนาคาร
         TempStr = FormatNumber(ct.GetFieldValue("NET_AMOUNT"))
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + ct.GetFieldValue("NET_AMOUNT")
         Total2(k) = Total2(k) + ct.GetFieldValue("NET_AMOUNT")
         BalanceCashAmount = BalanceCashAmount - ct.GetFieldValue("NET_AMOUNT")
         
         k = k + 1
          Set Hc = m_Headers0.Fields(k - 4)                                                ' เงินโอน
         TempStr = FormatNumber(0)
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         k = k + 1
         Set Hc = m_Headers0.Fields(k - 4)
         TempStr = FormatNumber(BalanceCashAmount)
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         For Each Ba In m_BankAccounts
            k = k + 1
            Set Hc = m_Headers0.Fields(k - 4)
            If ct.GetFieldValue("BANK_ACCOUNT") = Ba.GetFieldValue("BANK_ACCOUNT") Then
               TempStr = FormatNumber(ct.GetFieldValue("NET_AMOUNT"))
               Total1(k) = Total1(k) + (ct.GetFieldValue("NET_AMOUNT"))
               Total2(k) = Total2(k) + (ct.GetFieldValue("NET_AMOUNT"))
            Else
               TempStr = FormatNumber(0)
            End If
            Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
         Next Ba
         
         k = k + 1                                 ' เพิ่มหนี้ลดหนี้
         Set Hc = m_Headers0.Fields(k - 4)
         TempStr = FormatNumber(0)
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         k = k + 1
         Set Hc = m_Headers0.Fields(k - 4)
         TempStr = FormatNumber(0)
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         k = k + 1
         Set Hc = m_Headers0.Fields(k - 4)
         TempStr = FormatNumber(0)
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         k = k + 1
         Set Hc = m_Headers0.Fields(k - 4)
         TempStr = FormatNumber(ct.GetFieldValue("NET_AMOUNT"))
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) - (ct.GetFieldValue("NET_AMOUNT"))
         Total2(k) = Total2(k) - (ct.GetFieldValue("NET_AMOUNT"))
   
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
      Else
         For I = 1 To Use
            Call m_Payin.Remove(1)
         Next I
         Exit Sub
      End If
   Next ct
 
End Sub
