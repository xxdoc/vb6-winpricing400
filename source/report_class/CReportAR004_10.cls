VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAR004_10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAR004_10"
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
Private m_PayByChequeDoc As Collection
Private m_ReceiptPayByChequeDoc As Collection
Private m_SumReceiptPayByChequeDoc As Collection



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
   Set m_PayByChequeDoc = New Collection
   Set m_ReceiptPayByChequeDoc = New Collection
   Set m_SumReceiptPayByChequeDoc = New Collection
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
   Set m_PayByChequeDoc = Nothing
   Set m_ReceiptPayByChequeDoc = Nothing
   Set m_SumReceiptPayByChequeDoc = Nothing
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


   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "�ú��˹�", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "�Ţ�պ��", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "�ѹ�����", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "�ʹ���", ">")
   Call m_Headers1.AddField(Cf)

   Call Cf.SetFieldValue(2, "^", "�Ţ�����", "^")
   Call m_Headers1.AddField(Cf)
   
   
  Call Cf.SetFieldValue(2, "^", "�ѹ������", "^")
   Call m_Headers1.AddField(Cf)
   
  Call Cf.SetFieldValue(2, "^", "�ӹǹ�Թ���", ">") '�������¡��
    Call m_Headers1.AddField(Cf)
   

  Call Cf.SetFieldValue(2, "^", "��Ҥ��", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�Ң�", "<")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(0, "^", "˹�餧�����", ">")
   Call m_Headers1.AddField(Cf)
   
   

            
  Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)

'   '===
   Call Cf.SetFieldValue(3, "^", "�����١���", "^")
   Call m_Headers1_1.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "�����١���", "<")
   Call m_Headers1_1.AddField(Cf)
   If mcolParam("SUM_INDUE_FLAG") = 0 Then
      Call Cf.SetFieldValue(12, "^", "��ѡ�ҹ���", "<")
      Call m_Headers1_1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(12, "^", "��ѡ�ҹ���", "<")
      Call m_Headers1_1.AddField(Cf)
   End If

   '===
   Call Cf.SetFieldValue(1, "^", "��ػ", "^")
   Call m_Headers1_2.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "ǧ�Թ", "^")
   Call m_Headers1_2.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "�Թǧ�Թ", "<")
   Call m_Headers1_2.AddField(Cf)
   If mcolParam("SUM_INDUE_FLAG") = 0 Then
      Call Cf.SetFieldValue(6, "^", "㹴��", "<")
      Call m_Headers1_2.AddField(Cf)
   Else
      Call Cf.SetFieldValue(2, "^", "㹴��", "<")
      Call m_Headers1_2.AddField(Cf)
   End If
   Call Cf.SetFieldValue(8, "^", "㹴��", "<")
   Call m_Headers1_2.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "", "<")
   Call m_Headers1_2.AddField(Cf)
'   '===
   
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

   Call LoadPaidAmountByBill(Nothing, m_PaidAmounts, , mcolParam("TO_PAY_DATE"))
   Call LoadTotalPriceByBill(Nothing, m_BillTotalPrices, , mcolParam("TO_DOC_DATE"))
   Call LoadBillingDiscountByBill(Nothing, m_BillingDiscounts, , mcolParam("TO_DOC_DATE"))

      Call LoadDueDateInterval2(Nothing, m_MaxMins, mcolParam("SUM_INDUE_FLAG"))
    

   Call LoadTotalPriceByCustomer(Nothing, m_ApSupAmounts, , mcolParam("TO_DOC_DATE"), 1)
   Call LoadPaidAmountByCustomer(Nothing, m_ApSupPaids, , mcolParam("TO_PAY_DATE"))
   Call LoadDnCnAmountByBill(Nothing, m_DnItemsByBill, , mcolParam("TO_PAY_DATE"), 4, 2)
   Call LoadDnCnAmountByBill(Nothing, m_CnItemsByBill, , mcolParam("TO_PAY_DATE"), 3, 2)
   Call LoadDnCnAmountByBill(Nothing, m_RtItemsByBill, , mcolParam("TO_PAY_DATE"), 18, 2)
   Call LoadChequeDocWaitForBill(m_PayByChequeDoc, mcolParam("CUSTOMER_CODE"), mcolParam("FROM_CUSTOMER_CODE"), mcolParam("TO_CUSTOMER_CODE"), mcolParam("CUSTOMER_TYPE"), mcolParam("CUSTOMER_GRADE")) 'Load �����Cheque ʶҹм�ҹ 2 Ẻ    �����¡�� ��� �Ѵ�������
'   Call LoadReceiptPayByChequeDoc(m_ReceiptPayByChequeDoc, mcolParam("CUSTOMER_CODE"), mcolParam("FROM_CUSTOMER_CODE"), mcolParam("TO_CUSTOMER_CODE"), mcolParam("CUSTOMER_TYPE"), mcolParam("CUSTOMER_GRADE")) '����� ���Ѵ��ҹ���� �����¡㹵��ҧ report �������ʶҹ� �����¡ ���� �Ѵ�������
'   Call LoadSumReceiptPayByChequeDoc(m_SumReceiptPayByChequeDoc, mcolParam("CUSTOMER_CODE"), mcolParam("FROM_CUSTOMER_CODE"), mcolParam("TO_CUSTOMER_CODE"), mcolParam("CUSTOMER_TYPE"), mcolParam("CUSTOMER_GRADE")) '����� ���Ѵ��ҹ���� �����¡㹵��ҧ report �������ʶҹ� �����¡ ���� �Ѵ�������
   
   'LoadSumReceiptPayByChequeDoc
   
'   FromDate = "�ҡ�ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DOC_DATE")), "N/A")
   ToDate = " �֧ �ѹ����� " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DOC_DATE")), "N/A")
   ToPaidDate = "    �ѹ�������Թ�֧�ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("TO_PAY_DATE")), "N/A") & "    �ѹ������� " & EmptyToString(DateToStringExtEx2(mcolParam("PRINT_DATE")), "N/A")
   ToPaidDate = ToPaidDate & " ���§��� " & EmptyToString(mcolParam("ORDER_BY_NAME"), "N/A")
   
   MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & vbCrLf & mcolParam("REPORT_NAME") & _
                                       vbCrLf & ToDate & "  " & ToPaidDate
'FromDate &
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
Dim ShowOvrdueFlag As Boolean
Dim m_PayCheqDoc As CChequeDoc
Dim m_ReceiptCheqDoc As CChequeDoc

   Set BodyCf = New CReportField

   If mcolParam("PRINT_DATE") <= 0 Then
      PrintDate = Now
   Else
      PrintDate = mcolParam("PRINT_DATE")
   End If
   
'   If I = 79 And Apt.DOCUMENT_NO = "03-00629-07" Then
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
   
   ShowOvrdueFlag = False
   If Check2Flag(mcolParam("OVERDUE_FLAG")) = "Y" Then
      If DueCount > 0 Then
         ShowOvrdueFlag = True
      End If
   Else 'N
      ShowOvrdueFlag = True
   End If
   
   If InCompleteFlag And ShowOvrdueFlag Then          '�����ѧ���ú
      If NewCust Then '���������������ҧ���� 1 record �֧�������ҧ���
         Call GenerateHeader(Apt, m_Details1)
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
         
         Indue = 0
         Overdue = 0
         HeaderExist = True
         NewCust = False
      End If
            
'      If I = 79 Then
'         'Debug.Print
'      End If
      Call m_Details1.ClearField
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

      Set HeadCf = m_Headers1.Fields(4)
      TempStr = DateToStringExtEx2(Apt.DOCUMENT_DATE)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

      Set HeadCf = m_Headers1.Fields(5)
      TempStr = FormatNumber(Di.TOTAL_PRICE)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      Tot1(5) = Tot1(5) + Di.TOTAL_PRICE
      Tot2(5) = Tot2(5) + Di.TOTAL_PRICE
   
      '�Ѵ��͹���Ǥ������
      Di.TOTAL_PRICE = Val(Format(Di.TOTAL_PRICE, "0.00"))
      Bdc.DISCOUNT_AMOUNT = Val(Format(Bdc.DISCOUNT_AMOUNT, "0.00"))
      Ri1_1.DEBIT_CREDIT_AMOUNT = Val(Format(Ri1_1.DEBIT_CREDIT_AMOUNT, "0.00"))
      Ri1_2.DEBIT_CREDIT_AMOUNT = Val(Format(Ri1_2.DEBIT_CREDIT_AMOUNT, "0.00"))
      Ri1_3.DEBIT_CREDIT_AMOUNT = Val(Format(Ri1_3.DEBIT_CREDIT_AMOUNT, "0.00"))
      Ari.CASH_DISCOUNT = Val(Format(Ari.CASH_DISCOUNT, "0.00"))

      NetAmt = (Di.TOTAL_PRICE - Bdc.DISCOUNT_AMOUNT) + (Ri1_1.DEBIT_CREDIT_AMOUNT - Ri1_2.DEBIT_CREDIT_AMOUNT - Ri1_3.DEBIT_CREDIT_AMOUNT - Ari.CASH_DISCOUNT) - Ari.PAID_AMOUNT
      
      J = 6
      Sum = 0
      
      If DueCount <= 0 Then
         Indue = Indue + NetAmt
      Else
         Overdue = Overdue + NetAmt
      End If
         
      For Each MM In m_MaxMins
         Set HeadCf = m_Headers1.Fields(J)
         If (MM.MIN < DueCount) And (DueCount <= MM.MAX) Then
      
            TempStr = FormatNumber(NetAmt)
    
'            Tot1(j) = Tot1(j) + NetAmt
'            Tot2(j) = Tot2(j) + NetAmt
            Sum = Sum + NetAmt
         Else
         
            TempStr = FormatNumber(0)
 
         End If

      Next MM

     
      Set m_PayCheqDoc = GetObject("CChequeDoc", m_PayByChequeDoc, Apt.CUSTOMER_ID & "-" & Apt.DOCUMENT_NO & "-" & Apt.RECEIPT_CHEQUE_DOC_ID, False)

      If m_PayCheqDoc Is Nothing Then
      
          Set HeadCf = m_Headers1.Fields(J)
          TempStr = ""
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      
          J = J + 1
          Set HeadCf = m_Headers1.Fields(J)
          TempStr = ""
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
            J = J + 1
          Set HeadCf = m_Headers1.Fields(J)
         TempStr = ""
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          
              J = J + 1
          Set HeadCf = m_Headers1.Fields(J)
         TempStr = ""
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
              J = J + 1
          Set HeadCf = m_Headers1.Fields(J)
         TempStr = ""
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
       
       Else
          Set HeadCf = m_Headers1.Fields(J)
          TempStr = m_PayCheqDoc.CHEQUE_DOC_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      
          J = J + 1
       Set HeadCf = m_Headers1.Fields(J)
      TempStr = m_PayCheqDoc.CHEQUE_DOC_DATE
       Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
       Call m_Details1.AddField(BodyCf)
            
            J = J + 1
          Set HeadCf = m_Headers1.Fields(J)
          TempStr = FormatNumber(m_PayCheqDoc.PAID_AMOUNT)
          Tot1(J) = Tot1(J) + m_PayCheqDoc.PAID_AMOUNT
          Tot2(J) = Tot2(J) + m_PayCheqDoc.PAID_AMOUNT
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         
              J = J + 1
          Set HeadCf = m_Headers1.Fields(J)
         TempStr = m_PayCheqDoc.BANK_NAME
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
         
         
            J = J + 1
          Set HeadCf = m_Headers1.Fields(J)
         TempStr = m_PayCheqDoc.BANK_BRANCH_NAME
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
        
      End If
      

        J = J + 1
      Set HeadCf = m_Headers1.Fields(J)
      TempStr = FormatNumber(Sum)
      Tot1(J) = Tot1(J) + NetAmt
      Tot2(J) = Tot2(J) + NetAmt

      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Call Vsp.AddTable(strFormat, "", TempStr2)
   Else
      I = I - 1
   End If
   Set BodyCf = Nothing
End Sub
Private Sub GenerateDetail2(alngX() As Long, I As Long, Apt As CBillingDoc, Tot1() As Double, Tot2() As Double, Indue As Double, Overdue As Double, NewCust As Boolean, HeaderExist As Boolean)
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
Dim ShowOvrdueFlag As Boolean
Dim m_PayCheqDoc As CChequeDoc
Dim m_ReceiptCheqDoc As CChequeDoc

   Set BodyCf = New CReportField

   If mcolParam("PRINT_DATE") <= 0 Then
      PrintDate = Now
   Else
      PrintDate = mcolParam("PRINT_DATE")
   End If
   
'   If I = 79 And Apt.DOCUMENT_NO = "03-00629-07" Then
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
   
   ShowOvrdueFlag = False
   If Check2Flag(mcolParam("OVERDUE_FLAG")) = "Y" Then
      If DueCount > 0 Then
         ShowOvrdueFlag = True
      End If
   Else 'N
      ShowOvrdueFlag = True
   End If
   
   If InCompleteFlag And ShowOvrdueFlag Then          '�����ѧ���ú
      If NewCust Then '���������������ҧ���� 1 record �֧�������ҧ���
         Call GenerateHeader(Apt, m_Details1)
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
         
         Indue = 0
         Overdue = 0
         HeaderExist = True
         NewCust = False
      End If
            
'      If I = 79 Then
'         'Debug.Print
'      End If
      Call m_Details1.ClearField
      Set HeadCf = m_Headers1.Fields(1)
      TempStr = ""
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

      Set HeadCf = m_Headers1.Fields(2)
      TempStr = ""
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

      Set HeadCf = m_Headers1.Fields(3)
      TempStr = ""
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

      Set HeadCf = m_Headers1.Fields(4)
      TempStr = ""
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

      Set HeadCf = m_Headers1.Fields(5)
      TempStr = ""
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      Tot1(5) = Tot1(5) + 0
      Tot2(5) = Tot2(5) + 0
   
      '�Ѵ��͹���Ǥ������
      Di.TOTAL_PRICE = Val(Format(Di.TOTAL_PRICE, "0.00"))
      Bdc.DISCOUNT_AMOUNT = Val(Format(Bdc.DISCOUNT_AMOUNT, "0.00"))
      Ri1_1.DEBIT_CREDIT_AMOUNT = Val(Format(Ri1_1.DEBIT_CREDIT_AMOUNT, "0.00"))
      Ri1_2.DEBIT_CREDIT_AMOUNT = Val(Format(Ri1_2.DEBIT_CREDIT_AMOUNT, "0.00"))
      Ri1_3.DEBIT_CREDIT_AMOUNT = Val(Format(Ri1_3.DEBIT_CREDIT_AMOUNT, "0.00"))
      Ari.CASH_DISCOUNT = Val(Format(Ari.CASH_DISCOUNT, "0.00"))

      NetAmt = (Di.TOTAL_PRICE - Bdc.DISCOUNT_AMOUNT) + (Ri1_1.DEBIT_CREDIT_AMOUNT - Ri1_2.DEBIT_CREDIT_AMOUNT - Ri1_3.DEBIT_CREDIT_AMOUNT - Ari.CASH_DISCOUNT) - Ari.PAID_AMOUNT
      
      J = 6
      Sum = 0
      
      If DueCount <= 0 Then
         Indue = Indue + 0
      Else
         Overdue = Overdue + 0
      End If
         
      For Each MM In m_MaxMins
         Set HeadCf = m_Headers1.Fields(J)
         If (MM.MIN < DueCount) And (DueCount <= MM.MAX) Then
      
            TempStr = FormatNumber(NetAmt)
    

           Sum = Sum + NetAmt
         Else
         
            TempStr = FormatNumber(0)
 
         End If
      Next MM

     

      Set HeadCf = m_Headers1.Fields(J)
      Set m_PayCheqDoc = GetObject("CChequeDoc", m_PayByChequeDoc, Apt.CUSTOMER_ID & "-" & Apt.DOCUMENT_NO & "-" & Apt.RECEIPT_CHEQUE_DOC_ID, False)
     
   
      If m_PayCheqDoc Is Nothing Then
      
          Set HeadCf = m_Headers1.Fields(J)
          TempStr = ""
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      
          J = J + 1
          Set HeadCf = m_Headers1.Fields(J)
          TempStr = ""
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
            J = J + 1
          Set HeadCf = m_Headers1.Fields(J)
         TempStr = ""
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          
              J = J + 1
          Set HeadCf = m_Headers1.Fields(J)
          TempStr = ""
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
            J = J + 1
          Set HeadCf = m_Headers1.Fields(J)
         TempStr = ""
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
       
       Else
          Set HeadCf = m_Headers1.Fields(J)
          TempStr = m_PayCheqDoc.CHEQUE_DOC_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      
         J = J + 1
       Set HeadCf = m_Headers1.Fields(J)
      TempStr = m_PayCheqDoc.CHEQUE_DOC_DATE
       Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
       Call m_Details1.AddField(BodyCf)
      
            J = J + 1
          Set HeadCf = m_Headers1.Fields(J)
          TempStr = FormatNumber(m_PayCheqDoc.PAID_AMOUNT)
          Tot1(J) = Tot1(J) + m_PayCheqDoc.PAID_AMOUNT
          Tot2(J) = Tot2(J) + m_PayCheqDoc.PAID_AMOUNT
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         
         
            J = J + 1
          Set HeadCf = m_Headers1.Fields(J)
         TempStr = m_PayCheqDoc.BANK_NAME
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
         
         
            J = J + 1
          Set HeadCf = m_Headers1.Fields(J)
         TempStr = m_PayCheqDoc.BANK_BRANCH_NAME
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)

      
      End If
      
      

       J = J + 1
       Set HeadCf = m_Headers1.Fields(J)
       TempStr = FormatNumber(Sum)
      Tot1(J) = Tot1(J) + 0
      Tot2(J) = Tot2(J) + 0
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      
'If I = 3 Then
''Debug.Print
'End If
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Call Vsp.AddTable(strFormat, "", TempStr2)
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
Dim J As Long

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   If mcolParam("ORDER_BY") = 2 Then
      TempStr1 = Cm.CUSTOMER_CODE
      TempStr2 = Cm.CUSTOMER_NAME
      TempStr3 = "��ѡ�ҹ��� : " & Cm.EMP_NAME & " " & Cm.EMP_LNAME
   ElseIf mcolParam("ORDER_BY") = 3 Then
      TempStr1 = Cm.EMP_CODE
      TempStr2 = "��ѡ�ҹ��� : " & Cm.EMP_NAME & " " & Cm.EMP_LNAME
      TempStr3 = ""
   ElseIf mcolParam("ORDER_BY") = 4 Then
      TempStr1 = Cm.CUSTOMER_CODE
      TempStr2 = Cm.CUSTOMER_NAME
      TempStr3 = "��ѡ�ҹ��� : " & Cm.EMP_NAME & " " & Cm.EMP_LNAME
   ElseIf mcolParam("ORDER_BY") = 5 Then
      TempStr1 = Cm.EMP_CODE
      TempStr2 = "��ѡ�ҹ��� : " & Cm.EMP_NAME & " " & Cm.EMP_LNAME
      TempStr3 = Cm.CUSTOMER_NAME & " (" & Cm.CUSTOMER_CODE & ")"
   Else
      TempStr1 = Cm.CUSTOMER_CODE
      TempStr2 = Cm.CUSTOMER_NAME
      TempStr3 = "��ѡ�ҹ��� : " & Cm.EMP_NAME & " " & Cm.EMP_LNAME
   End If
   
   J = 0
   For Each HeadCf In m_Headers1_1.Fields
      J = J + 1
      If J = 1 Then
         TempStr = TempStr1
      ElseIf J = 2 Then
         TempStr = TempStr2
      ElseIf J = 3 Then
         TempStr = TempStr3
      Else
         TempStr = ""
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
         TempStr = "ǧ�Թ  = " & FormatNumber(CreditLimit)
      ElseIf J = 3 Then
         TempStr = "�Թǧ�Թ = " & FormatNumber((Indue + Overdue) - CreditLimit)
      ElseIf J = 4 Then
         If mcolParam("SUM_INDUE_FLAG") = 0 Then
            TempStr = "㹴��  = " & FormatNumber(Indue)
         Else
            TempStr = FormatNumber(Indue)
         End If
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

Dim PrevKey3 As String
Dim PrevKey4 As String
Dim Width As Double
Dim TempRs As ADODB.Recordset
Dim Ap As CCustomer
Dim TempApt As CDoItem
Dim Apt As CBillingDoc
Dim TempApRcIt As CReceiptItem
Dim Overdue As Double
Dim Indue As Double
Dim NewCust As Boolean
Dim HeaderExist As Boolean
Dim CreditLimit As Double
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
'      Apt.FROM_DATE = mcolParam("FROM_DOC_DATE")
      Apt.TO_DATE = mcolParam("TO_DOC_DATE")
      Apt.DOCUMENT_TYPE = 1
      Apt.CUSTOMER_TYPE = mcolParam("CUSTOMER_TYPE")
      Apt.CUSTOMER_GRADE = mcolParam("CUSTOMER_GRADE")
      Apt.OrderBy = mcolParam("ORDER_BY")
      Apt.OrderType = 1
      Apt.ItemSumFlag = True
      Call Apt.QueryData(108, Rs, BillCount)

      Indue = 0
      Overdue = 0
      
      If Not Rs.EOF Then
         Call Apt.PopulateFromRS(108, Rs)
         PrevKey1 = GetKey(2, Apt)
         NewCust = True
      End If

      While Not Rs.EOF
         Call Apt.PopulateFromRS(108, Rs)

         If PrevKey1 <> GetKey(2, Apt) Then
          
            If HeaderExist Then
            
               Call GenerateFooter("���", "", "", Total1, m_Details1, 0)
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
               
               NewCust = True
               HeaderExist = False
            End If
            For J = 1 To UBound(Total1)
               Total1(J) = 0
            Next J
         End If
         PrevKey1 = GetKey(2, Apt)
         PrevKey3 = Apt.DOCUMENT_NO
'         If Apt.CUSTOMER_CODE = "�-0007" Then
'            'Debug.Print
'         End If
         Call m_Details1.ClearField
       
         
         If I = 0 Or PrevKey4 <> PrevKey3 Then
          
           I = I + 1
           PrevKey4 = PrevKey3
            Call GenerateDetail(alngX(), I, Apt, Total1, Total2, Indue, Overdue, NewCust, HeaderExist)
         Else
            Call GenerateDetail2(alngX(), I, Apt, Total1, Total2, Indue, Overdue, NewCust, HeaderExist)
         End If
         
         CreditLimit = Apt.CREDIT_LIMIT
         Rs.MoveNext
      Wend

      If HeaderExist Then
         Call GenerateFooter("���", "", "", Total1, m_Details1, 0)
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
      

      End If
      
      Call GenerateFooter("�ط��", "", "", Total2, m_Details1, 0)
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
      ElseIf (J = 4) Then
         TempStr = ""
      ElseIf (J = 6) Then
         TempStr = ""
      ElseIf (J = 7) Then
         TempStr = ""
      ElseIf (J = 9) Then
       TempStr = ""
      ElseIf (J = 10) Then
       TempStr = ""
      ElseIf J = m_Headers1.Fields.Count Then
         TempStr = FormatNumber(Tot(J))
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




