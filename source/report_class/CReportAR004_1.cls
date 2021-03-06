VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAR004_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAR004_1"
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
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
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
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   Call m_Headers1_1.ClearField
   Call m_Headers1_2.ClearField
   
   Set Cf = New CReportField
      
   '===
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "�Фú��˹�", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(8, "^", "�Թ��˹�", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�Թ��˹�", "^")
   Call m_Headers0.AddField(Cf)
   '===
   
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "�����١���", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "�����١���", "<")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "�ʹ���", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "< 30 �ѹ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "30-60 �ѹ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "> 60 �ѹ", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "1-15 �ѹ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "16-30 �ѹ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "31-60 �ѹ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "> 60 �ѹ", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "���˹���ҧ", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "�Թ��˹�", "^")
   Call m_Headers1.AddField(Cf)

  Call m_Headers0.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
'   Call Vsp.AddTable(strFormat, "", strPrint)
            
  Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
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

   Call LoadPaidAmountByBill(Nothing, m_PaidAmounts, mcolParam("FROM_DOC_DATE"), mcolParam("TO_PAY_DATE"))
   Call LoadTotalPriceByBill(Nothing, m_BillTotalPrices, mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"))
   Call LoadBillingDiscountByBill(Nothing, m_BillingDiscounts, mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"))
   Call LoadDueDateInterval(Nothing, m_MaxMins)
   Call LoadTotalPriceByCustomer(Nothing, m_ApSupAmounts, mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"), 1)
   Call LoadPaidAmountByCustomer(Nothing, m_ApSupPaids, mcolParam("FROM_DOC_DATE"), mcolParam("TO_PAY_DATE"))
   Call LoadDnCnAmountByBill(Nothing, m_DnItemsByBill, mcolParam("FROM_DOC_DATE"), mcolParam("TO_PAY_DATE"), 4)
   Call LoadDnCnAmountByBill(Nothing, m_CnItemsByBill, mcolParam("FROM_DOC_DATE"), mcolParam("TO_PAY_DATE"), 3)
   Call LoadDnCnAmountByBill(Nothing, m_RtItemsByBill, mcolParam("FROM_DOC_DATE"), mcolParam("TO_PAY_DATE"), 18)
   
   FromDate = "�ҡ�ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DOC_DATE")), "N/A")
   ToDate = " �֧ �ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DOC_DATE")), "N/A")
   ToPaidDate = "    �ѹ�������Թ�֧�ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("TO_PAY_DATE")), "N/A") & "    �ѹ������� " & EmptyToString(DateToStringExtEx2(mcolParam("PRINT_DATE")), "N/A")

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

Private Sub GenerateDetail(alngX() As Long, I As Long, Apt As CBillingDoc, Tot1() As Double, Tot2() As Double, Indue As Double, Overdue As Double)
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
   
   If (Di.TOTAL_PRICE - Bdc.DISCOUNT_AMOUNT) + (Ri1_1.DEBIT_CREDIT_AMOUNT - Ri1_2.DEBIT_CREDIT_AMOUNT - Ri1_3.DEBIT_CREDIT_AMOUNT) <> Ari.PAID_AMOUNT Then      '�����ѧ���ú
      Set HeadCf = m_Headers1.Fields(1)
      TempStr = I
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

      Set HeadCf = m_Headers1.Fields(2)
      TempStr = DateToStringExtEx2(Apt.DUE_DATE)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

      Set HeadCf = m_Headers1.Fields(3)
      TempStr = Apt.DOCUMENT_NO
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

'      Set HeadCf = m_Headers1.Fields(4)
'      TempStr = DateToStringExtEx2(Apt.DOCUMENT_DATE)
'      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'      Call m_Details1.AddField(BodyCf)

      Set HeadCf = m_Headers1.Fields(4)
      TempStr = FormatNumber(Di.TOTAL_PRICE)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      Tot1(4) = Tot1(4) + Di.TOTAL_PRICE
      Tot2(4) = Tot2(4) + Di.TOTAL_PRICE

      '�Ѵ��͹���Ǥ������
      Di.TOTAL_PRICE = Val(Format(Di.TOTAL_PRICE, "0.00"))
      Bdc.DISCOUNT_AMOUNT = Val(Format(Bdc.DISCOUNT_AMOUNT, "0.00"))
      Ri1_1.DEBIT_CREDIT_AMOUNT = Val(Format(Ri1_1.DEBIT_CREDIT_AMOUNT, "0.00"))
      Ri1_2.DEBIT_CREDIT_AMOUNT = Val(Format(Ri1_2.DEBIT_CREDIT_AMOUNT, "0.00"))
      Ri1_3.DEBIT_CREDIT_AMOUNT = Val(Format(Ri1_3.DEBIT_CREDIT_AMOUNT, "0.00"))
      Ari.CASH_DISCOUNT = Val(Format(Ari.CASH_DISCOUNT, "0.00"))

      NetAmt = (Di.TOTAL_PRICE - Bdc.DISCOUNT_AMOUNT) + (Ri1_1.DEBIT_CREDIT_AMOUNT - Ri1_2.DEBIT_CREDIT_AMOUNT - Ri1_3.DEBIT_CREDIT_AMOUNT - Ari.CASH_DISCOUNT) - Ari.PAID_AMOUNT
      J = 5
      Sum = 0
      
      If DueCount <= Apt.Credit Then
         Indue = Indue + NetAmt
      Else
         Overdue = Overdue + NetAmt
      End If
         
      For Each MM In m_MaxMins
         Set HeadCf = m_Headers1.Fields(J)
         If (MM.MIN < DueCount) And (DueCount <= MM.MAX) Then
            TempStr = FormatNumber(NetAmt)

            Tot1(J) = Tot1(J) + NetAmt
            Tot2(J) = Tot2(J) + NetAmt
            Sum = Sum + NetAmt
         Else
            TempStr = FormatNumber(0)
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         J = J + 1
      Next MM

      Set HeadCf = m_Headers1.Fields(J)
      TempStr = FormatNumber(Sum)
      Tot1(J) = Tot1(J) + NetAmt
      Tot2(J) = Tot2(J) + NetAmt
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

      J = J + 1
      Set HeadCf = m_Headers1.Fields(J)
      TempStr = FormatNumber(DueCount, 0)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'      Call Vsp.AddTable(strFormat, "", TempStr2)
   Else
'      i = i - 1
   End If
   Set BodyCf = Nothing
End Sub

Private Sub GenerateHeader(Cm As CCustomer, Details As CFieldList)
Dim BodyCf As CReportField
Dim HeadCf As CReportField
Dim TempStr As String
Dim J As Long

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   J = 0
   For Each HeadCf In m_Headers1_1.Fields
      J = J + 1
      If J = 1 Then
         TempStr = Cm.CUSTOMER_CODE
      ElseIf J = 2 Then
         TempStr = Cm.CUSTOMER_NAME
'      ElseIf j = 3 Then
'         TempStr = "ǧ�Թ  = " & FormatNumber(Cm.CREDIT_LIMIT)
'      ElseIf j = 4 Then
'         TempStr = "㹴��  = " & FormatNumber(0)
'      ElseIf j = 5 Then
'         TempStr = "�Թ���  = " & FormatNumber(0)
'      ElseIf j = 6 Then
'         TempStr = "�������  = " & FormatNumber(0)
      Else
         TempStr = ""
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub

Private Sub GenerateFooter2(Cm As CCustomer, Details As CFieldList, Indue As Double, Overdue As Double)
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
         TempStr = "��ػ"
      ElseIf J = 2 Then
         TempStr = "ǧ�Թ  = " & FormatNumber(Cm.CREDIT_LIMIT)
      ElseIf J = 3 Then
         TempStr = "�Թǧ�Թ = " & FormatNumber((Indue + Overdue) - Cm.CREDIT_LIMIT)
      ElseIf J = 4 Then
         TempStr = "㹴��  = " & FormatNumber(Indue)
      ElseIf J = 5 Then
         TempStr = "�Թ���  = " & FormatNumber(Overdue)
      Else
         TempStr = ""
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub

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
Dim Width As Double
Dim TempRs As ADODB.Recordset
Dim Ap As CCustomer
Dim TempApt As CDoItem
Dim Apt As CBillingDoc
Dim TempApRcIt As CReceiptItem
Dim Overdue As Double
Dim Indue As Double

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

      Set Ap = New CCustomer
      Ap.CUSTOMER_CODE = mcolParam("CUSTOMER_CODE")
      
      Ap.FROM_CUSTOMER_CODE = mcolParam("FROM_CUSTOMER_CODE")
      Ap.TO_CUSTOMER_CODE = mcolParam("TO_CUSTOMER_CODE")
      
      Ap.CUSTOMER_TYPE = mcolParam("CUSTOMER_TYPE")
      Ap.CUSTOMER_GRADE = mcolParam("CUSTOMER_GRADE")
      Ap.OrderBy = mcolParam("ORDER_TYPE")
      Ap.OrderType = mcolParam("ORDER_BY")
      Call glbDaily.QueryCustomer(Ap, Rs, iCount, IsOK, glbErrorLog)

      While Not Rs.EOF
         For J = 1 To UBound(Total1)
            Total1(J) = 0
         Next J

         Call m_Details1.ClearField
         Call Ap.PopulateFromRS(1, Rs)

         '=== Detail
         Set Apt = New CBillingDoc
         Apt.CUSTOMER_ID = Ap.CUSTOMER_ID
         Apt.FROM_DATE = mcolParam("FROM_DOC_DATE")
         Apt.TO_DATE = mcolParam("TO_DOC_DATE")
         Apt.DOCUMENT_TYPE = 1
         Apt.OrderBy = 7
         Apt.OrderType = 1
         Apt.ItemSumFlag = True
         Call Apt.QueryData(1, TempRs, BillCount)

         If BillCount > 0 Then
            Set TempApt = GetDoItem(m_ApSupAmounts, Ap.CUSTOMER_ID)
            Set TempApRcIt = GetReceiptItem(m_ApSupPaids, Ap.CUSTOMER_ID)
            TempApRcIt.PAID_AMOUNT = 0
            If TempApt.TOTAL_PRICE <> TempApRcIt.PAID_AMOUNT Then    '���ʹ��ҧ���� ���� �Դź
               '===Head
               Call GenerateHeader(Ap, m_Details1)
   
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               'Call Vsp.AddTable(strFormat, "", TempStr2)
               '===
   
               Overdue = 0
               Indue = 0
               While Not TempRs.EOF
'                  i = i + 1
                  Call m_Details1.ClearField
                  Call Apt.PopulateFromRS(1, TempRs)
   
                  Call GenerateDetail(alngX(), I, Apt, Total1, Total2, Indue, Overdue)
   
                  TempRs.MoveNext
               Wend
               Set Apt = Nothing
   
               I = I + 1
               Call GenerateFooter(Trim(str(I)), Ap.CUSTOMER_CODE, Ap.CUSTOMER_NAME, Total1, m_Details1, 0)
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
'               Vsp.FontBold = False
            End If
          End If

         Rs.MoveNext
      Wend

      Call GenerateFooter("���", "", "", Total2, m_Details1, 0)
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
   Call Vsp.AddTable(strFormat, "", "|˹�ҷ��|: " & Vsp.CurrentPage)
   Call Vsp.AddTable(strFormat, "", "|�ѹ|: " & DateToStringExt(mdteDate))
   Call Vsp.AddTable(strFormat, "", "|����|: " & Right$(DateToStringExtEx(Now), 8))
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




