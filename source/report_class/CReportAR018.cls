VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAR018"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAR018"
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
Private m_PaidAmounts As Collection
Private m_CnAmounts As Collection
Private m_DnAmounts As Collection
Private m_RtAmounts As Collection
Private m_DiscountAmounts As Collection
Private EnablePrintHeader As Boolean
Dim alngX() As Long
Dim MaxPage As Long
Dim MaxRow As Long
Dim BD As CBillingDoc
Dim Ri1_0 As CReceiptItem
Dim Ri1_1 As CReceiptItem
Dim Ri1_2 As CReceiptItem
Dim Ri1_3 As CReceiptItem
Dim Bdc As CBillingDiscount
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
   Set m_PaidAmounts = New Collection
   Set m_CnAmounts = New Collection
   Set m_DnAmounts = New Collection
   Set m_RtAmounts = New Collection
   Set m_DiscountAmounts = New Collection
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
   Set m_PaidAmounts = Nothing
   Set m_CnAmounts = Nothing
   Set m_DnAmounts = Nothing
   Set m_RtAmounts = Nothing
   Set m_DiscountAmounts = Nothing
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
   Vsp.FontBold = False
   Vsp.TableBorder = tbNone
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Vsp.FontSize = 16
   
   Set Cf = New CReportField

   Call Cf.SetFieldValue(10, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(20, "^", "วันที่เอกสาร", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(20, "^", "เลขที่เอกสาร", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(20, "^", "วันที่ครบกำหนดชำระ", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(22, "^", "คงค้างชำระ (บาท)", ">")
   Call m_Headers1.AddField(Cf)

   Set Cf = Nothing
   Vsp.TableBorder = tbAll
  Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   Vsp.TableBorder = tbNone
            
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
   
   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprLetter
   Vsp.ORIENTATION = orPortrait
   Vsp.MarginBottom = 500
   Vsp.MarginFooter = 500
   Vsp.MarginHeader = 500
   Vsp.MarginLeft = 500
   Vsp.MarginRight = 700
   Vsp.MarginTop = 500
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE

   Call LoadPaidAmountByBill(Nothing, m_PaidAmounts, -1, mcolParam("TO_DATE"), , , , mcolParam("CUSTOMER_CODE"), mcolParam("CUSTOMER_CODE"))
   Call LoadDnCnAmountByBill(Nothing, m_DnAmounts, -1, mcolParam("TO_DATE"), 4, 2, mcolParam("CUSTOMER_CODE"), mcolParam("CUSTOMER_CODE"))
   Call LoadDnCnAmountByBill(Nothing, m_CnAmounts, -1, mcolParam("TO_DATE"), 3, 2, mcolParam("CUSTOMER_CODE"), mcolParam("CUSTOMER_CODE"))
   Call LoadDnCnAmountByBill(Nothing, m_RtAmounts, -1, mcolParam("TO_DATE"), 18, 2, mcolParam("CUSTOMER_CODE"), mcolParam("CUSTOMER_CODE"))
   Call LoadBillingDiscountByBill(Nothing, m_DiscountAmounts, -1, mcolParam("TO_DATE"), , , mcolParam("CUSTOMER_CODE"), mcolParam("CUSTOMER_CODE"))
   
   ToDate = " ณ วันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")
   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))

   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
      mstrErrMsg = "Error GetDateTime Error."
      Exit Function
   End If
   mdteDate = InternalDateToDate(strDate)
   initDoc = True
End Function
Private Function printHeader2()
Dim strFormat As String '
Dim strPrint As String

Vsp.TableBorder = tbNone

  Vsp.FontSize = 18
   strFormat = VSP_CalTable("^1;", mdblWid, alngX)
   Vsp.TextAlign = taCenterMiddle
    Vsp.FontBold = True
    strPrint = "เอกสารเพื่อสอบทานจำนวนหนี้ทางบัญชี" & vbNewLine & "(ไม่ใช่เอกสารเพื่อทวงถามหนี้)"
    Call Vsp.AddTable(strFormat, "", strPrint)

    Vsp.FontSize = 14
    Vsp.FontBold = False
    strFormat = VSP_CalTable("<2|<15|>3|<3;", mdblWid, alngX)
   Call Vsp.AddTable(strFormat, "", "||หน้าที่ | " & Vsp.CurrentPage & IIf(MaxPage = 0, "/1", "/" & MaxPage))
    strFormat = VSP_CalTable("<2|<15|>3|<3;", mdblWid, alngX)
    strPrint = "||วันที่ |" & DateToStringExtEx2(mdteDate) & ";"
    Call Vsp.AddTable(strFormat, "", strPrint)
    
   Vsp.FontSize = 16
   strFormat = VSP_CalTable("<1;", mdblWid, alngX)
    Vsp.FontBold = False
    strPrint = "บริษัท มิตรภาพอาหารสัตว์ จำกัด;"
    Call Vsp.AddTable(strFormat, "", strPrint)
    
    Vsp.FontSize = 16
    strFormat = VSP_CalTable("<1;", mdblWid, alngX)
    strPrint = "11 ซ.รามคำแหง 81/1 แขวงหัวหมาก เขตบางกะปิ กทม 10240;"
    Call Vsp.AddTable(strFormat, "", strPrint)
    
    strFormat = VSP_CalTable("<1;", mdblWid, alngX)
    strPrint = "โทร. (02)7322448-57 โทรสาร 7329132;"
    Call Vsp.AddTable(strFormat, "", strPrint)
    Vsp.FontBold = False
    
    strFormat = VSP_CalTable("<2|<20;", mdblWid, alngX)
    strPrint = "เรียน|ผู้จัดการฝ่ายบัญชี;"
    Call Vsp.AddTable(strFormat, "", strPrint)
    Vsp.FontBold = False
    
    Dim m_Customer As CCustomer
    Dim ItemCount  As Long
    Dim m_Rs As ADODB.Recordset
    Set m_Rs = New ADODB.Recordset
    Set m_Customer = New CCustomer
    m_Customer.CUSTOMER_CODE = mcolParam("CUSTOMER_CODE")
   Call m_Customer.QueryData1(m_Rs, ItemCount)
   
   While Not (m_Rs.EOF)
    Call m_Customer.PopulateFromRS(2, m_Rs)
    strFormat = VSP_CalTable("<2|<15|>3|<3;", mdblWid, alngX) '<2|<15|<5
    strPrint = "|" & m_Customer.CUSTOMER_NAME & "|รหัสลูกค้า|" & m_Customer.CUSTOMER_CODE & ";"
    Call Vsp.AddTable(strFormat, "", strPrint)
    
      strFormat = VSP_CalTable("<2|<20;", mdblWid, alngX)
      strPrint = "|" & m_Customer.PackAddressEx1(False) & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
      m_Rs.MoveNext
   Wend
    
    
     strFormat = VSP_CalTable("<0|<20;", mdblWid, alngX)
      strPrint = "|" & "      เพื่อประโยชน์ในการตรวจสอบจำนวนรายการสินค้าที่ซื้อขายต่อกัน ซึ่งผู้ซื้อได้รับสินค้าที่ซื้อไว้แล้วและเพื่อเป็นการตรวจสอบ" & vbNewLine & "ทางบัญชี บริษัทจึงเรียนมาเพื่อแจ้งยอดคงเหลือทางบัญชี ณ.วันที่ " & Format(mcolParam("TO_DATE"), "DD") & " เดือน " & IntToThaiMonth(Month(mcolParam("TO_DATE"))) & " พ.ศ. " & Format(mcolParam("TO_DATE"), "YYYY") & " ตามรายการดังนี้;"
      Call Vsp.AddTable(strFormat, "", strPrint)
End Function
Private Sub findDataRow()
Dim iCount As Long
Dim IsOK As Boolean
Dim I As Long
Dim II As Long
Dim isHave As Boolean



     Set BD = New CBillingDoc
      BD.BILLING_DOC_ID = -1
      BD.TO_DATE = mcolParam("TO_DATE")
      BD.CUSTOMER_CODE = mcolParam("CUSTOMER_CODE")
      BD.OrderBy = 3
      BD.OrderType = 1
      BD.DOCUMENT_TYPE = 1
      BD.ItemSumFlag = True
      Call glbDaily.QueryBillingDoc(BD, Rs, iCount, IsOK, glbErrorLog)
      I = 0

      While Not Rs.EOF
         I = I + 1
         
         Call m_Details1.ClearField
         Call BD.PopulateFromRS(1, Rs)

         Set Ri1_0 = GetReceiptItem(m_PaidAmounts, BD.BILLING_DOC_ID) 'รับชำระ
         Set Ri1_1 = GetReceiptItem(m_DnAmounts, BD.BILLING_DOC_ID) 'เพิ่มหนี้
         Set Ri1_2 = GetReceiptItem(m_CnAmounts, BD.BILLING_DOC_ID) 'ลดหนี้
         Set Ri1_3 = GetReceiptItem(m_RtAmounts, BD.BILLING_DOC_ID) 'รับคืน
         Set Bdc = GetBillingDiscount(m_DiscountAmounts, BD.BILLING_DOC_ID) 'ส่วนลด
         
         BD.PAID_AMOUNT = Ri1_0.PAID_AMOUNT
         BD.CASH_DISCOUNT = Ri1_0.CASH_DISCOUNT
         BD.DEBIT_AMOUNT = Ri1_1.DEBIT_CREDIT_AMOUNT
         BD.CREDIT_AMOUNT = Ri1_2.DEBIT_CREDIT_AMOUNT
         BD.RETURN_AMOUNT = Ri1_3.DEBIT_CREDIT_AMOUNT
         BD.DISCOUNT_AMOUNT = Bdc.DISCOUNT_AMOUNT
                  
         'ทำให้เป็น 2 ตำแหน่งก่อนแล้วค่อยรวม
         BD.PAID_AMOUNT = Val(Format(BD.PAID_AMOUNT, "0.00"))
         BD.CASH_DISCOUNT = Val(Format(BD.CASH_DISCOUNT, "0.00"))
         BD.DEBIT_AMOUNT = Val(Format(BD.DEBIT_AMOUNT, "0.00"))
         BD.CREDIT_AMOUNT = Val(Format(BD.CREDIT_AMOUNT, "0.00"))
         BD.RETURN_AMOUNT = Val(Format(BD.RETURN_AMOUNT, "0.00"))
         BD.DISCOUNT_AMOUNT = Val(Format(BD.DISCOUNT_AMOUNT, "0.00"))
         
         BD.DO_TOTAL_PRICE = Val(Format(BD.DO_TOTAL_PRICE, "0.00"))
         BD.REVENUE_TOTAL_PRICE = Val(Format(BD.REVENUE_TOTAL_PRICE, "0.00"))


         If (ROUND(BD.DO_TOTAL_PRICE + BD.REVENUE_TOTAL_PRICE - BD.DISCOUNT_AMOUNT + (BD.DEBIT_AMOUNT - BD.CREDIT_AMOUNT) - BD.PAID_AMOUNT - BD.CASH_DISCOUNT, 2) <= 0) Then
            I = I - 1
         Else
            If (BD.DO_TOTAL_PRICE + BD.REVENUE_TOTAL_PRICE - BD.DISCOUNT_AMOUNT + (BD.DEBIT_AMOUNT - BD.CREDIT_AMOUNT) - BD.PAID_AMOUNT - BD.CASH_DISCOUNT) > 0 Then
            End If
         End If
         Rs.MoveNext
      Wend
      
      If (I >= 8) Then
        If (I Mod 8) = 0 Then
            MaxPage = (I \ 8)
        Else
         MaxPage = (I \ 8) + 1
        End If
      Else
         MaxPage = 0
      End If
      MaxRow = I
End Sub



Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim RName As String
Dim I As Long
Dim J As Long
Dim R As Long
Dim strFormat As String
Dim IsOK As Boolean
Dim Amt As Double
'Dim BD As CBillingDoc
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

Dim strPrint  As String
Dim isHave As Boolean
   
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
   Next J
'   EnablePrintHeader = False
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   mblnHeader = True
     Call findDataRow
      Call printHeader2
      Call printHeader
      I = 0
      Set BodyCf = New CReportField

      I = 0
      If Rs.RecordCount = 0 Then
         glbErrorLog.LocalErrorMsg = MapText("ไม่มีข้อมูลของลูกค้ารหัส " & mcolParam("CUSTOMER_CODE"))
         glbErrorLog.ShowUserError
         Exit Function
      End If
      
       Rs.MoveFirst
      If Not Rs.EOF Then
         PrevKey1 = NVLS(Rs("CUSTOMER_CODE"), "")
      End If
      
       
      While Not Rs.EOF
         I = I + 1
         
         Call m_Details1.ClearField
         Call BD.PopulateFromRS(1, Rs)

         Set Ri1_0 = GetReceiptItem(m_PaidAmounts, BD.BILLING_DOC_ID) 'รับชำระ
         Set Ri1_1 = GetReceiptItem(m_DnAmounts, BD.BILLING_DOC_ID) 'เพิ่มหนี้
         Set Ri1_2 = GetReceiptItem(m_CnAmounts, BD.BILLING_DOC_ID) 'ลดหนี้
         Set Ri1_3 = GetReceiptItem(m_RtAmounts, BD.BILLING_DOC_ID) 'รับคืน
         Set Bdc = GetBillingDiscount(m_DiscountAmounts, BD.BILLING_DOC_ID) 'ส่วนลด
         
         BD.PAID_AMOUNT = Ri1_0.PAID_AMOUNT
         BD.CASH_DISCOUNT = Ri1_0.CASH_DISCOUNT
         BD.DEBIT_AMOUNT = Ri1_1.DEBIT_CREDIT_AMOUNT
         BD.CREDIT_AMOUNT = Ri1_2.DEBIT_CREDIT_AMOUNT
         BD.RETURN_AMOUNT = Ri1_3.DEBIT_CREDIT_AMOUNT
         BD.DISCOUNT_AMOUNT = Bdc.DISCOUNT_AMOUNT
                  
         'ทำให้เป็น 2 ตำแหน่งก่อนแล้วค่อยรวม
         BD.PAID_AMOUNT = Val(Format(BD.PAID_AMOUNT, "0.00"))
         BD.CASH_DISCOUNT = Val(Format(BD.CASH_DISCOUNT, "0.00"))
         BD.DEBIT_AMOUNT = Val(Format(BD.DEBIT_AMOUNT, "0.00"))
         BD.CREDIT_AMOUNT = Val(Format(BD.CREDIT_AMOUNT, "0.00"))
         BD.RETURN_AMOUNT = Val(Format(BD.RETURN_AMOUNT, "0.00"))
         BD.DISCOUNT_AMOUNT = Val(Format(BD.DISCOUNT_AMOUNT, "0.00"))
         
         BD.DO_TOTAL_PRICE = Val(Format(BD.DO_TOTAL_PRICE, "0.00"))
         BD.REVENUE_TOTAL_PRICE = Val(Format(BD.REVENUE_TOTAL_PRICE, "0.00"))
         'Bd.PAID_AMOUNT = Val(Format(Bd.PAID_AMOUNT, "0.00"))
           
         Set HeadCf = m_Headers1.Fields(1)
         TempStr = I
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(2)
         TempStr = DateToStringExtEx2(BD.DOCUMENT_DATE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(3)
         TempStr = BD.DOCUMENT_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(4)
         TempStr = DateToStringExtEx2(BD.DUE_DATE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(5)
         TempStr = FormatNumber(BD.DO_TOTAL_PRICE + BD.REVENUE_TOTAL_PRICE - BD.DISCOUNT_AMOUNT + (BD.DEBIT_AMOUNT - BD.CREDIT_AMOUNT - BD.RETURN_AMOUNT) - BD.PAID_AMOUNT - BD.CASH_DISCOUNT)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         If (ROUND(BD.DO_TOTAL_PRICE + BD.REVENUE_TOTAL_PRICE - BD.DISCOUNT_AMOUNT + (BD.DEBIT_AMOUNT - BD.CREDIT_AMOUNT) - BD.PAID_AMOUNT - BD.CASH_DISCOUNT, 2) <= 0) Then
            I = I - 1
            isHave = False
         Else
            If (BD.DO_TOTAL_PRICE + BD.REVENUE_TOTAL_PRICE - BD.DISCOUNT_AMOUNT + (BD.DEBIT_AMOUNT - BD.CREDIT_AMOUNT) - BD.PAID_AMOUNT - BD.CASH_DISCOUNT) > 0 Then
               Total1(5) = Total1(5) + (BD.DO_TOTAL_PRICE + BD.REVENUE_TOTAL_PRICE - BD.DISCOUNT_AMOUNT + (BD.DEBIT_AMOUNT - BD.CREDIT_AMOUNT - BD.RETURN_AMOUNT) - BD.PAID_AMOUNT - BD.CASH_DISCOUNT)
               
               Total2(5) = Total2(5) + (BD.DO_TOTAL_PRICE + BD.REVENUE_TOTAL_PRICE - BD.DISCOUNT_AMOUNT + (BD.DEBIT_AMOUNT - BD.CREDIT_AMOUNT - BD.RETURN_AMOUNT) - BD.PAID_AMOUNT - BD.CASH_DISCOUNT)
              
              isHave = True
              Vsp.TableBorder = tbAll
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
            End If
         End If
      If I = 128 Then
      Debug.Print
      End If
           If I Mod 8 = 0 And I > 0 And isHave Then

            
               
               If MaxRow > I Then 'กรณีที่ หมดรอบ ใน loop พอดี
                  Call GenerateFooter("รวม", "", "", Total1, m_Details1, 1)
                  Call m_Details1.GetString(2, TempStr1, TempStr2)
                  strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                  Vsp.FontBold = True
                  Call Vsp.AddTable(strFormat, "", TempStr2)
                  Vsp.FontBold = False
                  
                  For J = 1 To UBound(Total1)
                  Total1(J) = 0
                  Next J
                  
                  Call PrintFooter
                  Vsp.NewPage
               End If
         End If
         Rs.MoveNext
         
      Wend

   Call GenerateFooter("รวม", "", "", Total1, m_Details1, 1)
   Call m_Details1.GetString(2, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False

If MaxPage > 0 Then 'ถ้าเป็นหน้าเดียว ไม่ต้องแสดงยอด สุทธิ
   Call GenerateFooter("รวมสุทธิ", "", "", Total2, m_Details1, 1)
   Call m_Details1.GetString(2, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False
End If

   Vsp.TableBorder = tbNone
   
'   If AdjustPage(Vsp, strFormat, strPrint, 6151) Then
'      Vsp = ""
'      Vsp = ""
'   End If

   Call PrintFooter

   Set BodyCf = Nothing
   Vsp.EndDoc
   Call CloseExportFile(Vsp)

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
Private Function PrintFooter()
Dim strFormat As String '
Dim strPrint As String

Vsp.TableBorder = tbNone

   Vsp = ""
   strFormat = VSP_CalTable("<13|^12;", mdblWid, alngX)
   strPrint = "|ลงชื่อ" & " . . . . . . . . . . . . . . . . . . . . . . . . . . . บริษัท หรือผู้ทำการแทน;"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("<9|^15;", mdblWid, alngX)
   strPrint = "|" & "( นายสุชาย ไหลวณิชชาเศรษฐ์ );"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("^15;", mdblWid, alngX)
   strPrint = "_ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Vsp.FontSize = 16
   strFormat = VSP_CalTable("<1;", mdblWid, alngX)
   Vsp.FontBold = True
   strPrint = "[ส่วนนี้ลูกค้าลงรายการ];"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Vsp.FontBold = False
   strFormat = VSP_CalTable("<0|<20;", mdblWid, alngX)
   strPrint = "|" & "      ข้าพเจ้าในฐานะผู้ซื้อสินค้าได้ตรวจสอบรายการข้างต้นทั้งหมดเป็นรายการที่ [   ] ถูกต้อง [   ] ไม่ถูกต้อง จำนวนยอดหนี้ที่ถูกต้อง " & vbNewLine & "จำนวน . . . . . . . . . . . . . .  บาท (. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . ) ตามเอกสารหลักฐานแนบการสอบทานยอดการชำระค่าสินค้า;"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Vsp = ""
      
   strFormat = VSP_CalTable("<9|^15;", mdblWid, alngX)
   strPrint = "|ลงชื่อ" & " . . . . . . . . . . . . . . . . . . . . . . . . . . . ผู้ซื้อ;"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("<9|^15;", mdblWid, alngX)
   strPrint = "|" & "( . . . . . . . . . . . . . . . . . . . . . . . . . . . );"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("<9|^15;", mdblWid, alngX)
   strPrint = "|" & " . . . . . . ./ . . . . . . . . . . /. . . . . . . . . . ;"
   Call Vsp.AddTable(strFormat, "", strPrint)
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
      ElseIf (J > 4) Then
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
'Dim talnT As TextAlignSettings
'Dim tbdT As TableBorderSettings
'Dim blnBold As Boolean
'Dim blnUnder As Boolean
'Dim blnItalic As Boolean
'Dim iSize As Integer
'Dim sName As String
'Dim strFormat As String
'Dim dY(0 To 1) As Double
'Dim alngX() As Long
'
'   If Not mblnNewPage Then Exit Sub
'   talnT = Vsp.TextAlign
'   tbdT = Vsp.TableBorder
'   blnBold = Vsp.FontBold
'   blnUnder = Vsp.FontUnderline
'   blnItalic = Vsp.FontItalic
'   iSize = Vsp.FontSize
'   sName = Vsp.FontName
'
'   Vsp.FontSize = 12
'   dY(0) = Vsp.CurrentY
'   Vsp.TextAlign = taLeftBottom
'   strFormat = VSP_CalTable(TITLE_FORMAT, mdblWid, alngX)
'   Vsp.TableBorder = tbNone
'   Call Vsp.AddTable(strFormat, "", "|หน้าที่|: " & Vsp.CurrentPage)
'   Call Vsp.AddTable(strFormat, "", "|วัน|: " & DateToStringExt(mdteDate))
'   Vsp.CurrentY = dY(0)
'
'   Vsp.FontSize = TITLE_SIZE
'   Vsp.FontBold = True
'   Vsp.TextAlign = taCenterBottom
'   Vsp.Paragraph = MODULE_DESC
'   Vsp.Paragraph = ""
   If mblnHeader Then
      Call printHeader2
      Call printHeader
   End If
'   mdY = Vsp.CurrentY
'
'   Vsp.TextAlign = talnT
'   Vsp.TableBorder = tbdT
'   Vsp.FontBold = blnBold
'   Vsp.FontUnderline = blnUnder
'   Vsp.FontItalic = blnItalic
'   Vsp.FontSize = iSize
'   Vsp.FontName = sName
End Sub

Private Sub VSP_StartDoc()
'This event occur when VSPrinter.StartDoc is used and used to initialize some information before generating printed document
End Sub



'Option Explicit
''User Group Report
'Implements CReportInterface
'
'Private Const MODULE_NAME = "CReportAR018"
'Public MODULE_DESC As String
'
'Private WithEvents Vsp As VSPrinter
'Private mstrErrMsg As String
'Private mcolParam As Collection
'Private mblnNewPage As Boolean
'Private mblnHeader As Boolean
'Private mblnEndPage As Boolean
'Private mdblWid   As Double
'Private mdteDate As Date
'Private mdY As Double
'Private Rs As ADODB.Recordset
'
'Private Const TITLE_SIZE = 14
'Private Const HEADER_SIZE = 12
'Private Const DETAIL_SIZE = 12
'
'Private Const TITLE_FORMAT = "<32|<3|<5;"
'
'Private m_TempCol As Collection
'Private m_HeaderFormat1 As String
'Private m_HeaderFormat2 As String
'Private m_HeaderText1 As String
'Private m_HeaderText2 As String
'
'Private m_Headers0 As CFieldList
'Private m_Headers1 As CFieldList
'Private m_Headers1_1 As CFieldList
'Private m_Headers1_2 As CFieldList
'Private m_Headers1_3 As CFieldList
'Private m_Headers1_4 As CFieldList
'Private m_Details1 As CFieldList
'Private m_Details2 As CFieldList
'Private m_Left As Double
'Private m_YearSeqs As Collection
'Private m_PaidAmounts As Collection
'Private m_CnAmounts As Collection
'Private m_DnAmounts As Collection
'Private m_RtAmounts As Collection
'Private m_DiscountAmounts As Collection
'Private EnablePrintHeader As Boolean
'Private Sub Class_Initialize()
'   Set mcolParam = New Collection
'   Set m_TempCol = New Collection
'
'   Set m_Headers0 = New CFieldList
'   Set m_Headers1 = New CFieldList
'   Set m_Headers1_1 = New CFieldList
'   Set m_Headers1_2 = New CFieldList
'   Set m_Headers1_3 = New CFieldList
'   Set m_Headers1_4 = New CFieldList
'   Set m_Details1 = New CFieldList
'   Set m_Details2 = New CFieldList
'   Set m_YearSeqs = New Collection
'   Set m_PaidAmounts = New Collection
'   Set m_CnAmounts = New Collection
'   Set m_DnAmounts = New Collection
'   Set m_RtAmounts = New Collection
'   Set m_DiscountAmounts = New Collection
'End Sub
'
'Private Sub Class_Terminate()
'   Call ClearParam
'   Set mcolParam = Nothing
'   Set Vsp = Nothing
'   Set m_TempCol = Nothing
'
'   Set m_Headers0 = Nothing
'   Set m_Headers1 = Nothing
'   Set m_Details1 = Nothing
'   Set m_YearSeqs = Nothing
'   Set m_Headers1_1 = Nothing
'   Set m_Headers1_2 = Nothing
'   Set m_Headers1_3 = Nothing
'   Set m_Headers1_4 = Nothing
'   Set m_Details2 = Nothing
'   Set m_PaidAmounts = Nothing
'   Set m_CnAmounts = Nothing
'   Set m_DnAmounts = Nothing
'   Set m_RtAmounts = Nothing
'   Set m_DiscountAmounts = Nothing
'End Sub
'
'Private Function CReportInterface_AddParam(varItem As Variant, strIndex As String) As Boolean
'   Call mcolParam.add(varItem, strIndex)
'   CReportInterface_AddParam = True
'End Function
'
'Private Sub CReportInterface_ClearParam()
'   Call ClearParam
'End Sub
'
'Private Property Get CReportInterface_ErrorMsg() As String
'   CReportInterface_ErrorMsg = mstrErrMsg
'End Property
'
'Private Function CReportInterface_Preview() As Boolean
'   CReportInterface_Preview = genDoc(True)
'End Function
'
'Private Function CReportInterface_PrintDoc() As Boolean
'   CReportInterface_PrintDoc = genDoc(False)
'End Function
'
'Private Property Set CReportInterface_VsPrint(RHS As VSPrinter7LibCtl.IVSPrinter)
'   Set Vsp = RHS
'End Property
'
'Private Sub ClearParam()
'   Dim I As Long
'
'   For I = 1 To mcolParam.Count
'      mcolParam.Remove 1
'   Next I
'
'End Sub
''============================= Codes above used in every report ==========================
'
'Private Sub printHeader()
'Dim strFormat As String
'Dim strPrint As String
'Dim tbdT As TableBorderSettings
'Dim blnBold As Boolean
'Dim iSize As Integer
'Dim alngX() As Long
'Dim Sgi As CSGroupItem
'Dim TempSgi As CSGroupItem
'Dim PrevID As Long
'Dim TempStr1 As String
'Dim TempStr1_1 As String
'Dim Cf As CReportField
'Dim ExportWidth As Long
'Dim HeadCf As CReportField
'Dim J As Long
'Dim TempStr As String
'Dim Inv As CInventoryDoc
'Dim TempRs As ADODB.Recordset
'Dim ImportItems As Collection
'Dim Amt As Double
'Dim iCount As Long
'
'   tbdT = Vsp.TableBorder
'   blnBold = Vsp.FontBold
'   iSize = Vsp.FontSize
'
'   Vsp.FontSize = HEADER_SIZE
'   Vsp.FontBold = True
'   Vsp.TableBorder = tbNone
'
'   Call m_Headers0.ClearField
'   Call m_Headers1.ClearField
'
'   Set Cf = New CReportField
'
'   Call Cf.SetFieldValue(1, "^", "NO.", "^")
'   Call m_Headers1.AddField(Cf)
'
'   Call Cf.SetFieldValue(3, "^", "วันที่ครบกำหนด", "^")
'   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(3, "^", "เลขที่เอกสาร", "^")
'   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(3, "^", "วันที่เอกสาร", "^")
'   Call m_Headers1.AddField(Cf)
'
'   Call Cf.SetFieldValue(2, "^", "ยอดคงค้าง", ">")
'   Call m_Headers1.AddField(Cf)
'
'   Set Cf = Nothing
'
'  Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
'   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'   strPrint = TempStr1_1
'   Call Vsp.AddTable(strFormat, "", strPrint)
'
'   Set TempRs = Nothing
'
'   Vsp.TableBorder = tbdT
'   Vsp.FontBold = blnBold
'   Vsp.FontSize = iSize
'End Sub
'
'Private Function initDoc() As Boolean
'Dim strDate As String
'Dim TempStr1 As String
'Dim TempStr2 As String
'Dim TempStr3 As String
'Dim FromDate As String
'Dim ToDate As String
'
'   mstrErrMsg = ""
'   mblnHeader = True
'   mblnNewPage = True
'   mblnEndPage = True
'   Vsp.PaperSize = pprLetter
'   Vsp.ORIENTATION = orPortrait
'   Vsp.MarginBottom = 500
'   Vsp.MarginFooter = 500
'   Vsp.MarginHeader = 500
'   Vsp.MarginLeft = 500
'   Vsp.MarginRight = 700
'   Vsp.MarginTop = 500
'   Vsp.FontName = "AngsanaUPC"
'   Vsp.FontSize = DETAIL_SIZE
'
'   Call LoadPaidAmountByBill(Nothing, m_PaidAmounts, -1, mcolParam("TO_DATE"), , , , mcolParam("CUSTOMER_CODE"), mcolParam("CUSTOMER_CODE"))
'   Call LoadDnCnAmountByBill(Nothing, m_DnAmounts, -1, mcolParam("TO_DATE"), 4, 2, mcolParam("CUSTOMER_CODE"), mcolParam("CUSTOMER_CODE"))
'   Call LoadDnCnAmountByBill(Nothing, m_CnAmounts, -1, mcolParam("TO_DATE"), 3, 2, mcolParam("CUSTOMER_CODE"), mcolParam("CUSTOMER_CODE"))
'   Call LoadDnCnAmountByBill(Nothing, m_RtAmounts, -1, mcolParam("TO_DATE"), 18, 2, mcolParam("CUSTOMER_CODE"), mcolParam("CUSTOMER_CODE"))
'   Call LoadBillingDiscountByBill(Nothing, m_DiscountAmounts, -1, mcolParam("TO_DATE"), , , mcolParam("CUSTOMER_CODE"), mcolParam("CUSTOMER_CODE"))
'
'   ToDate = " ณ วันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")
'   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))
'
'   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
'   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
'      mstrErrMsg = "Error GetDateTime Error."
'      Exit Function
'   End If
'   mdteDate = InternalDateToDate(strDate)
'   initDoc = True
'End Function
'
'Private Function genDoc(isPreview As Boolean) As Boolean
'On Error GoTo ErrHandler
'Dim RName As String
'Dim I As Long
'Dim J As Long
'Dim strFormat As String
'Dim alngX() As Long
'Dim IsOK As Boolean
'Dim Amt As Double
'Dim BD As CBillingDoc
'Dim HeadCf As CReportField
'Dim BodyCf As CReportField
'Dim TempStr1 As String
'Dim TempStr2 As String
'Dim Total1(100) As Double
'Dim Total2(100) As Double
'Dim Total3(100) As Double
'Dim iCount As Long
'Dim TempStr As String
'Dim PrevKey1 As String
'Dim PrevKey2 As String
'Dim Ri1_0 As CReceiptItem
'Dim Ri1_1 As CReceiptItem
'Dim Ri1_2 As CReceiptItem
'Dim Ri1_3 As CReceiptItem
'Dim Bdc As CBillingDiscount
'Dim strPrint  As String
'
'   RName = "genDoc"
''-----------------------------------------------------------------------------------------------------
''                                             Query Here
''-----------------------------------------------------------------------------------------------------
'
'   Set Rs = New ADODB.Recordset
'
'   If Not initDoc Then Exit Function
'
'   Vsp.Preview = isPreview
''-----------------------------------------------------------------------------------------------------
''                                         Main Operation Here
''-----------------------------------------------------------------------------------------------------
'
'   For J = 1 To UBound(Total1)
'      Total1(J) = 0
'      Total2(J) = 0
'      Total3(J) = 0
'   Next J
'   EnablePrintHeader = False
'
'   mblnHeader = False
'   Call StartExportFile(Vsp)
'   Vsp.StartDoc
'   Call printHeader
'   mblnHeader = True
'
'   EnablePrintHeader = True
'   Vsp.TableBorder = tbNone
'
'    Vsp.FontSize = 18
'   strFormat = VSP_CalTable("<1;", mdblWid, alngX)
'    Vsp.FontBold = True
'    strPrint = "บริษัท มิตรภาพอาหารสัตว์ จำกัด;"
'    Call Vsp.AddTable(strFormat, "", strPrint)
'
'    Vsp.FontSize = 16
'    strFormat = VSP_CalTable("<1;", mdblWid, alngX)
'    strPrint = "11 ซ.รามคำแหง 81/1 แขวงหัวหมาก เขตบางกะปิ กทม 10240;"
'    Call Vsp.AddTable(strFormat, "", strPrint)
'
'    strFormat = VSP_CalTable("<1;", mdblWid, alngX)
'    strPrint = "โทร. (02)7322448-57 โทรสาร 7329132;"
'    Call Vsp.AddTable(strFormat, "", strPrint)
'    Vsp.FontBold = False
'
'    Vsp.FontSize = 14
'    strFormat = VSP_CalTable("<20|<2|<4;", mdblWid, alngX)
'   Call Vsp.AddTable(strFormat, "", "|หน้าที่|: " & Vsp.CurrentPage)
'   Call Vsp.AddTable(strFormat, "", "|วัน|: " & DateToStringExtEx2(mdteDate))
'
'    Vsp.FontBold = True
'
'    strFormat = VSP_CalTable("^10;", mdblWid, alngX)
'    strPrint = "รายการบัญชีสอบทาน (ไม่ใช้ใบทวงหนี้);"
'    Call Vsp.AddTable(strFormat, "", strPrint)
'
'    Vsp = ""
'
'    strFormat = VSP_CalTable("<2|<20;", mdblWid, alngX)
'    strPrint = "เรียน|ผู้จัดการฝ่ายบัญชี;"
'    Call Vsp.AddTable(strFormat, "", strPrint)
'    Vsp.FontBold = False
'
'    Dim m_Customer As CCustomer
'    Dim ItemCount  As Long
'    Dim m_Rs As ADODB.Recordset
'    Set m_Rs = New ADODB.Recordset
'    Set m_Customer = New CCustomer
'    m_Customer.CUSTOMER_CODE = mcolParam("CUSTOMER_CODE")
'   Call m_Customer.QueryData1(m_Rs, ItemCount)
'
'    Call m_Customer.PopulateFromRS(2, m_Rs)
'
'    strFormat = VSP_CalTable("<2|<15|<5;", mdblWid, alngX)
'    strPrint = "|" & m_Customer.CUSTOMER_NAME & "|รหัสลูกค้า" & m_Customer.CUSTOMER_CODE & ";"
'    Call Vsp.AddTable(strFormat, "", strPrint)
'
'      strFormat = VSP_CalTable("<2|<20;", mdblWid, alngX)
'      strPrint = "|" & m_Customer.PackAddressEx1(False) & ";"
'      Call Vsp.AddTable(strFormat, "", strPrint)
'
'    Vsp = ""
'
'     strFormat = VSP_CalTable("<2|<20;", mdblWid, alngX)
'      strPrint = "|" & "เพื่อประโยชน์ในการตรวจสอบบัญชี บริษัทฯจึงได้ส่งยอดคงเหลือบัญชี ณ วันที่ " & EmptyToString(DateToStringExt(mcolParam("TO_DATE")), "N/A") & " แสดงว่า ;"
'      Call Vsp.AddTable(strFormat, "", strPrint)
'
'      strFormat = VSP_CalTable("<10;", mdblWid, alngX)
'      strPrint = "เป็นลูกหนี้ของบริษัทฯดังรายการต่อไปนี้(การซื้อหรือการชำระเงินหลังจากวันที่ตรวจสอบไม่รวมอยู่ในรายการนี้);"
'      Call Vsp.AddTable(strFormat, "", strPrint)
'
'      Vsp = ""
'
'      Call printHeader
'      I = 0
'      Set BodyCf = New CReportField
'
'      Set BD = New CBillingDoc
'      BD.BILLING_DOC_ID = -1
'      BD.TO_DATE = mcolParam("TO_DATE")
'      BD.CUSTOMER_CODE = mcolParam("CUSTOMER_CODE")
'      BD.OrderBy = 3 'mcolParam("ORDER_BY")
'      BD.OrderType = 1
'      BD.DOCUMENT_TYPE = 1
'      BD.ItemSumFlag = True
'      Call glbDaily.QueryBillingDoc(BD, Rs, iCount, IsOK, glbErrorLog)
'      I = 0
'
'      If Not Rs.EOF Then
'         PrevKey1 = NVLS(Rs("CUSTOMER_CODE"), "")
'      End If
'
'      While Not Rs.EOF
'         I = I + 1
'
'         If I = 14 Then
'            Vsp.NewPage
'         End If
''         If (I Mod 13 = 0) And (Vsp.CurrentPage = 1) Then
''            Vsp.NewPage
''         ElseIf ((I - 13) Mod 26 = 0) And (Vsp.CurrentPage > 1) Then
''            Vsp.NewPage
''         End If
'
'         Call m_Details1.ClearField
'         Call BD.PopulateFromRS(1, Rs)
'
'         Set Ri1_0 = GetReceiptItem(m_PaidAmounts, BD.BILLING_DOC_ID) 'รับชำระ
'         Set Ri1_1 = GetReceiptItem(m_DnAmounts, BD.BILLING_DOC_ID) 'เพิ่มหนี้
'         Set Ri1_2 = GetReceiptItem(m_CnAmounts, BD.BILLING_DOC_ID) 'ลดหนี้
'         Set Ri1_3 = GetReceiptItem(m_RtAmounts, BD.BILLING_DOC_ID) 'รับคืน
'         Set Bdc = GetBillingDiscount(m_DiscountAmounts, BD.BILLING_DOC_ID) 'ส่วนลด
'
'         BD.PAID_AMOUNT = Ri1_0.PAID_AMOUNT
'         BD.CASH_DISCOUNT = Ri1_0.CASH_DISCOUNT
'         BD.DEBIT_AMOUNT = Ri1_1.DEBIT_CREDIT_AMOUNT
'         BD.CREDIT_AMOUNT = Ri1_2.DEBIT_CREDIT_AMOUNT
'         BD.RETURN_AMOUNT = Ri1_3.DEBIT_CREDIT_AMOUNT
'         BD.DISCOUNT_AMOUNT = Bdc.DISCOUNT_AMOUNT
'
'         'ทำให้เป็น 2 ตำแหน่งก่อนแล้วค่อยรวม
'         BD.PAID_AMOUNT = Val(Format(BD.PAID_AMOUNT, "0.00"))
'         BD.CASH_DISCOUNT = Val(Format(BD.CASH_DISCOUNT, "0.00"))
'         BD.DEBIT_AMOUNT = Val(Format(BD.DEBIT_AMOUNT, "0.00"))
'         BD.CREDIT_AMOUNT = Val(Format(BD.CREDIT_AMOUNT, "0.00"))
'         BD.RETURN_AMOUNT = Val(Format(BD.RETURN_AMOUNT, "0.00"))
'         BD.DISCOUNT_AMOUNT = Val(Format(BD.DISCOUNT_AMOUNT, "0.00"))
'
'         BD.DO_TOTAL_PRICE = Val(Format(BD.DO_TOTAL_PRICE, "0.00"))
'         BD.REVENUE_TOTAL_PRICE = Val(Format(BD.REVENUE_TOTAL_PRICE, "0.00"))
'         'Bd.PAID_AMOUNT = Val(Format(Bd.PAID_AMOUNT, "0.00"))
'
'         Set HeadCf = m_Headers1.Fields(1)
'         TempStr = I
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'         Call m_Details1.AddField(BodyCf)
'
'         Set HeadCf = m_Headers1.Fields(2)
'         TempStr = DateToStringExtEx2(BD.DUE_DATE)
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'         Call m_Details1.AddField(BodyCf)
'
'         Set HeadCf = m_Headers1.Fields(3)
'         TempStr = BD.DOCUMENT_NO
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'         Call m_Details1.AddField(BodyCf)
'
'         Set HeadCf = m_Headers1.Fields(4)
'         TempStr = DateToStringExtEx2(BD.DOCUMENT_DATE)
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'         Call m_Details1.AddField(BodyCf)
'
'         Set HeadCf = m_Headers1.Fields(5)
'         TempStr = FormatNumber(BD.DO_TOTAL_PRICE + BD.REVENUE_TOTAL_PRICE - BD.DISCOUNT_AMOUNT + (BD.DEBIT_AMOUNT - BD.CREDIT_AMOUNT - BD.RETURN_AMOUNT) - BD.PAID_AMOUNT - BD.CASH_DISCOUNT)
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'         Call m_Details1.AddField(BodyCf)
'
'         If (ROUND(BD.DO_TOTAL_PRICE + BD.REVENUE_TOTAL_PRICE - BD.DISCOUNT_AMOUNT + (BD.DEBIT_AMOUNT - BD.CREDIT_AMOUNT) - BD.PAID_AMOUNT - BD.CASH_DISCOUNT, 2) <= 0) Then
'            I = I - 1
'         Else
'            If (BD.DO_TOTAL_PRICE + BD.REVENUE_TOTAL_PRICE - BD.DISCOUNT_AMOUNT + (BD.DEBIT_AMOUNT - BD.CREDIT_AMOUNT) - BD.PAID_AMOUNT - BD.CASH_DISCOUNT) > 0 Then
'               Total1(5) = Total1(5) + (BD.DO_TOTAL_PRICE + BD.REVENUE_TOTAL_PRICE - BD.DISCOUNT_AMOUNT + (BD.DEBIT_AMOUNT - BD.CREDIT_AMOUNT - BD.RETURN_AMOUNT) - BD.PAID_AMOUNT - BD.CASH_DISCOUNT)
'
'               Total2(5) = Total2(5) + (BD.DO_TOTAL_PRICE + BD.REVENUE_TOTAL_PRICE - BD.DISCOUNT_AMOUNT + (BD.DEBIT_AMOUNT - BD.CREDIT_AMOUNT - BD.RETURN_AMOUNT) - BD.PAID_AMOUNT - BD.CASH_DISCOUNT)
'
'               Call m_Details1.GetString(1, TempStr1, TempStr2)
'               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'               Call Vsp.AddTable(strFormat, "", TempStr2)
'            End If
'         End If
'         Rs.MoveNext
'      Wend
'
'      If PrevKey1 <> NVLS(Rs("CUSTOMER_CODE"), "") Then
'         If I > 1 Then
'            PrevKey1 = NVLS(Rs("CUSTOMER_CODE"), "")
'            If Total1(5) > 0 Then
'               Call GenerateFooter("รวม", "", "", Total1, m_Details1, 1)
'               Call m_Details1.GetString(2, TempStr1, TempStr2)
'               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'               Vsp.FontBold = True
'               Call Vsp.AddTable(strFormat, "", TempStr2)
'               Vsp.FontBold = False
'            End If
'            For J = 1 To UBound(Total1)
'               Total1(J) = 0
'            Next J
'         End If
'      End If
'
'   ''Debug.Print Vsp.CurrentY
'   EnablePrintHeader = False
'   If AdjustPage(Vsp, strFormat, strPrint, 3151) Then
'      Vsp = ""
'      Vsp = ""
'   End If
'   strFormat = VSP_CalTable("<2|<20;", mdblWid, alngX)
'   strPrint = "หมายเหตุ|" & "หากท่านเห็นว่ายอดบัญชีที่แจ้งมาไม่ถูกต้อง กรุณาแสดงรายละเอียดที่ปรากฏอยู่ หรือหลักฐานที่ท่านมีอยู่หรือคำชี้แจงใดๆด้วย" & vbCrLf & "แบบฟอร์มนี้มิใช้การเตือนหรือเรียกร้องให้ชำระหนี้สินแต่อย่างใด;"
'   Call Vsp.AddTable(strFormat, "", strPrint)
'
'   strFormat = VSP_CalTable("<15|^8;", mdblWid, alngX)
'   strPrint = "|" & "ขอแสดงความนับถือ;"
'   Call Vsp.AddTable(strFormat, "", strPrint)
'
'   Vsp = ""
'
'   strFormat = VSP_CalTable("<15|^8;", mdblWid, alngX)
'   strPrint = "|" & "_ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ ;"
'   Call Vsp.AddTable(strFormat, "", strPrint)
'
'   strFormat = VSP_CalTable("<15|^8;", mdblWid, alngX)
'   strPrint = "|" & "( นายสุชาย ไหลวณิชชาเศรษฐ์ );"
'   Call Vsp.AddTable(strFormat, "", strPrint)
'
'   strFormat = VSP_CalTable("<15|^8;", mdblWid, alngX)
'   strPrint = "|" & "ในนามของ บริษัทมิตรภาพอาหารสัตว์ จำกัด;"
'   Call Vsp.AddTable(strFormat, "", strPrint)
'
'   Vsp.FontSize = 8
'   strFormat = VSP_CalTable("^5|^15;", mdblWid, alngX)
'   strPrint = "ประทับตราบริษัท/ห้างร้าน|;"
'   Call Vsp.AddTable(strFormat, "", strPrint)
'
'   strFormat = VSP_CalTable("^5|^15;", mdblWid, alngX)
'   strPrint = "(ถ้ามี)|;"
'   Call Vsp.AddTable(strFormat, "", strPrint)
'   ''Debug.Print Vsp.CurrentY
'   Vsp.FontSize = 14
'
'   EnablePrintHeader = False
'   Vsp.NewPage
'
'   Vsp = ""
'   Vsp = ""
'
'   Vsp.FontSize = 18
'   Vsp.FontBold = True
'   strFormat = VSP_CalTable("^5;", mdblWid, alngX)
'   strPrint = "หนังสือตอบยืนยันยอด;"
'   Call Vsp.AddTable(strFormat, "", strPrint)
'   Vsp.FontSize = 14
'   Vsp.FontBold = False
'   Vsp = ""
'   Vsp = ""
'
'   strFormat = VSP_CalTable("<2|<20;", mdblWid, alngX)
'    strPrint = "เรียน|ผู้จัดการฝ่ายบัญชี;"
'    Call Vsp.AddTable(strFormat, "", strPrint)
'
'    strFormat = VSP_CalTable("<2|<20;", mdblWid, alngX)
'    strPrint = "|บริษัท มิตรภาพอาหารสัตว์ จำกัด;"
'    Call Vsp.AddTable(strFormat, "", strPrint)
'
'    Vsp = ""
'
'    strFormat = VSP_CalTable("<20;", mdblWid, alngX)
'    strPrint = "           ก.ข้าพเจ้าฯขอยืนยันว่า ณ วันที่ " & EmptyToString(DateToStringExt(mcolParam("TO_DATE")), "N/A") & " เป็นลูกหนี้ บริษัทฯท่านตามที่แจ้งนั้นถูกต้องแล้ว" & ";"
'    Call Vsp.AddTable(strFormat, "", strPrint)
'
'    strFormat = VSP_CalTable("<20;", mdblWid, alngX)
'    strPrint = "เป็นจำนวนเงิน ................................................................(ตัวอักษร ..................................................................................................);"
'    Call Vsp.AddTable(strFormat, "", strPrint)
'
'    strFormat = VSP_CalTable("<20;", mdblWid, alngX)
'    strPrint = "           ข.ข้าพเจ้าฯขอเรียนว่า ยอดหนี้ที่แจ้งข้างต้นไม่ถูกต้อง สิ้นสุด ณ วันที่ " & EmptyToString(DateToStringExt(mcolParam("TO_DATE")), "N/A") & " จำนวน" & ";"
'    Call Vsp.AddTable(strFormat, "", strPrint)
'
'    strFormat = VSP_CalTable("<20;", mdblWid, alngX)
'    strPrint = "เงินที่ถูกต้องคือ ................................................................ บาท (ตัวอักษร ..................................................................................................);"
'    Call Vsp.AddTable(strFormat, "", strPrint)
'
'    strFormat = VSP_CalTable("<20;", mdblWid, alngX)
'    strPrint = "รายการที่แตกต่าง ..................................................................................................................................................................................;"
'    Call Vsp.AddTable(strFormat, "", strPrint)
'
'    strFormat = VSP_CalTable("<20;", mdblWid, alngX)
'    strPrint = "..............................................................................................................................................................................................................;"
'    Call Vsp.AddTable(strFormat, "", strPrint)
'
'    strFormat = VSP_CalTable("<20;", mdblWid, alngX)
'    strPrint = "..............................................................................................................................................................................................................;"
'    Call Vsp.AddTable(strFormat, "", strPrint)
'
'    Vsp = ""
'    Vsp = ""
'
'    strFormat = VSP_CalTable("<15|^8;", mdblWid, alngX)
'   strPrint = "|" & "ขอแสดงความนับถือ;"
'   Call Vsp.AddTable(strFormat, "", strPrint)
'
'   Vsp = ""
'
'   strFormat = VSP_CalTable("<15|^8;", mdblWid, alngX)
'   strPrint = "|" & "_ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ ;"
'   Call Vsp.AddTable(strFormat, "", strPrint)
'
'   strFormat = VSP_CalTable("<15|^8;", mdblWid, alngX)
'   strPrint = "|" & "(                                                              );"
'   Call Vsp.AddTable(strFormat, "", strPrint)
'
'   Vsp.FontSize = 8
'   strFormat = VSP_CalTable("^5|^15;", mdblWid, alngX)
'   strPrint = "ประทับตราบริษัท/ห้างร้าน|;"
'   Call Vsp.AddTable(strFormat, "", strPrint)
'
'   strFormat = VSP_CalTable("^5|^15;", mdblWid, alngX)
'   strPrint = "(ถ้ามี)|;"
'   Call Vsp.AddTable(strFormat, "", strPrint)
'
'
'   Set BodyCf = Nothing
'   Vsp.EndDoc
'   Call CloseExportFile(Vsp)
'
'   If Rs.State = adStateOpen Then
'      Rs.Close
'   End If
'   Set Rs = Nothing
'
'   genDoc = True
'   Exit Function
'
'ErrHandler:
'   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
'   Set Rs = Nothing
'End Function
'
'Private Sub GenerateFooter(TXT1 As String, TXT2 As String, TXT3 As String, Tot() As Double, Details As CFieldList, DateCount As Long)
'Dim HeadCf As CReportField
'Dim BodyCf As CReportField
'Dim J As Long
'Dim TempStr As String
'
'   Call Details.ClearField
'   Set BodyCf = New CReportField
'
'   J = 0
'   For Each HeadCf In m_Headers1.Fields
'      J = J + 1
'      If J = 1 Then
'         TempStr = TXT1
'      ElseIf J = 2 Then
'         TempStr = TXT2
'      ElseIf J = 3 Then
'         TempStr = TXT3
'      ElseIf (J > 4) Then
'         TempStr = FormatNumber(Tot(J))
'      Else
'         TempStr = ""
'      End If
'
'      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
'      Call Details.AddField(BodyCf)
'   Next HeadCf
'
'   Set BodyCf = Nothing
'End Sub
'
'Private Sub VSP_EndDoc()
''This event occur when VSPrinter.EndDoc is used
'End Sub
'
'Private Sub VSP_EndPage()
'   If Not mblnEndPage Then Exit Sub
''   Vsp.DrawLine Vsp.MarginLeft, mdY, Vsp.MarginLeft, Vsp.CurrentY
''   Vsp.DrawLine Vsp.PageWidth - Vsp.MarginRight, mdY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY
''   Vsp.DrawLine Vsp.MarginLeft, Vsp.CurrentY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY
'End Sub
'
'Private Sub VSP_Error()
''Error in runtime occur here press F1 in VSP.ErrorDescription to see more information
'   mstrErrMsg = Vsp.ErrorDescription
'End Sub
'
'Private Sub VSP_NewPage()
''Dim talnT As TextAlignSettings
''Dim tbdT As TableBorderSettings
''Dim blnBold As Boolean
''Dim blnUnder As Boolean
''Dim blnItalic As Boolean
''Dim iSize As Integer
''Dim sName As String
''Dim strFormat As String
''Dim dY(0 To 1) As Double
''Dim alngX() As Long
''
''   If Not mblnNewPage Then Exit Sub
''   talnT = Vsp.TextAlign
''   tbdT = Vsp.TableBorder
''   blnBold = Vsp.FontBold
''   blnUnder = Vsp.FontUnderline
''   blnItalic = Vsp.FontItalic
''   iSize = Vsp.FontSize
''   sName = Vsp.FontName
''
''   Vsp.FontSize = 12
''   dY(0) = Vsp.CurrentY
''   Vsp.TextAlign = taLeftBottom
''   strFormat = VSP_CalTable(TITLE_FORMAT, mdblWid, alngX)
''   Vsp.TableBorder = tbNone
''   Call Vsp.AddTable(strFormat, "", "|หน้าที่|: " & Vsp.CurrentPage)
''   Call Vsp.AddTable(strFormat, "", "|วัน|: " & DateToStringExt(mdteDate))
''   Vsp.CurrentY = dY(0)
''
''   Vsp.FontSize = TITLE_SIZE
''   Vsp.FontBold = True
''   Vsp.TextAlign = taCenterBottom
''   Vsp.Paragraph = MODULE_DESC
''   Vsp.Paragraph = ""
'   If EnablePrintHeader Then
'      Vsp = ""
'      Vsp = ""
'      Call printHeader
'   End If
''   mdY = Vsp.CurrentY
''
''   Vsp.TextAlign = talnT
''   Vsp.TableBorder = tbdT
''   Vsp.FontBold = blnBold
''   Vsp.FontUnderline = blnUnder
''   Vsp.FontItalic = blnItalic
''   Vsp.FontSize = iSize
''   Vsp.FontName = sName
'End Sub
'
'Private Sub VSP_StartDoc()
''This event occur when VSPrinter.StartDoc is used and used to initialize some information before generating printed document
'End Sub
'
'
