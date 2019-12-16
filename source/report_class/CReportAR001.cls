VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAR001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAR001"
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
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField

   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "รหัสลูกค้า", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(5, "^", "ชื่อลูกค้า", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "วันที่เอกสาร", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "เลขที่เอกสาร", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "เพิ่ม/ลดหนี้", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ชำระแล้ว", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ยอดค้างชำระ", ">")
   Call m_Headers1.AddField(Cf)

   Set Cf = Nothing
            
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

   Call LoadPaidAmountByBill(Nothing, m_PaidAmounts, mcolParam("FROM_DATE"), mcolParam("TO_PAY_DATE"))
   Call LoadDnCnAmountByBill(Nothing, m_DnAmounts, mcolParam("FROM_DATE"), mcolParam("TO_PAY_DATE"), 4, 2)
   Call LoadDnCnAmountByBill(Nothing, m_CnAmounts, mcolParam("FROM_DATE"), mcolParam("TO_PAY_DATE"), 3, 2)
   Call LoadDnCnAmountByBill(Nothing, m_RtAmounts, mcolParam("FROM_DATE"), mcolParam("TO_PAY_DATE"), 18, 2)
   Call LoadBillingDiscountByBill(Nothing, m_DiscountAmounts, mcolParam("FROM_DATE"), mcolParam("TO_DATE"))

   FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
   ToDate = " ถึง วันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A") & " ถึง วันที่ชำระเงิน " & EmptyToString(DateToStringExtEx2(mcolParam("TO_PAY_DATE")), "N/A")
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

Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim RName As String
Dim I As Long
Dim J As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Amt As Double
Dim Bd As CBillingDoc
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
Dim Ri1_0 As CReceiptItem
Dim Ri1_1 As CReceiptItem
Dim Ri1_2 As CReceiptItem
Dim Ri1_3 As CReceiptItem
Dim Bdc As CBillingDiscount

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
      
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
      I = 0
      Set BodyCf = New CReportField

      Set Bd = New CBillingDoc
      Bd.BILLING_DOC_ID = -1
      Bd.FROM_DATE = mcolParam("FROM_DATE")
      Bd.TO_DATE = mcolParam("TO_DATE")
      Bd.CUSTOMER_CODE = mcolParam("CUSTOMER_CODE")
      Bd.FROM_CUSTOMER_CODE = mcolParam("FROM_CUSTOMER_CODE")
      Bd.TO_CUSTOMER_CODE = mcolParam("TO_CUSTOMER_CODE")
      Bd.OrderBy = 3 'mcolParam("ORDER_BY")
      Bd.OrderType = mcolParam("ORDER_TYPE")
      If Bd.OrderType <= 0 Then
         Bd.OrderType = 1
      End If
      Bd.DOCUMENT_TYPE = 1
      Bd.ItemSumFlag = True
      Call glbDaily.QueryBillingDoc(Bd, Rs, iCount, IsOK, glbErrorLog)
      I = 0

      If Not Rs.EOF Then
         PrevKey1 = NVLS(Rs("CUSTOMER_CODE"), "")
      End If
      
      While Not Rs.EOF
         I = I + 1
         
         If PrevKey1 <> NVLS(Rs("CUSTOMER_CODE"), "") Then
            PrevKey1 = NVLS(Rs("CUSTOMER_CODE"), "")
            
'            If Total1(9) > 0 Then
               Call GenerateFooter("รวม", "", "", Total1, m_Details1, 1)
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
'            End If
            
            For J = 1 To UBound(Total1)
               Total1(J) = 0
            Next J
         End If
         
         Call m_Details1.ClearField
         Call Bd.PopulateFromRS(1, Rs)

         Set Ri1_0 = GetReceiptItem(m_PaidAmounts, Bd.BILLING_DOC_ID) 'รับชำระ
         Set Ri1_1 = GetReceiptItem(m_DnAmounts, Bd.BILLING_DOC_ID) 'เพิ่มหนี้
         Set Ri1_2 = GetReceiptItem(m_CnAmounts, Bd.BILLING_DOC_ID) 'ลดหนี้
         Set Ri1_3 = GetReceiptItem(m_RtAmounts, Bd.BILLING_DOC_ID) 'รับคืน
         Set Bdc = GetBillingDiscount(m_DiscountAmounts, Bd.BILLING_DOC_ID) 'ส่วนลด
         
         Bd.PAID_AMOUNT = Ri1_0.PAID_AMOUNT
         Bd.CASH_DISCOUNT = Ri1_0.CASH_DISCOUNT
         Bd.DEBIT_AMOUNT = Ri1_1.DEBIT_CREDIT_AMOUNT
         Bd.CREDIT_AMOUNT = Ri1_2.DEBIT_CREDIT_AMOUNT
         Bd.RETURN_AMOUNT = Ri1_3.DEBIT_CREDIT_AMOUNT
         Bd.DISCOUNT_AMOUNT = Bdc.DISCOUNT_AMOUNT
                  
         'ทำให้เป็น 2 ตำแหน่งก่อนแล้วค่อยรวม
         Bd.PAID_AMOUNT = Val(Format(Bd.PAID_AMOUNT, "0.00"))
         Bd.CASH_DISCOUNT = Val(Format(Bd.CASH_DISCOUNT, "0.00"))
         Bd.DEBIT_AMOUNT = Val(Format(Bd.DEBIT_AMOUNT, "0.00"))
         Bd.CREDIT_AMOUNT = Val(Format(Bd.CREDIT_AMOUNT, "0.00"))
         Bd.RETURN_AMOUNT = Val(Format(Bd.RETURN_AMOUNT, "0.00"))
         Bd.DISCOUNT_AMOUNT = Val(Format(Bd.DISCOUNT_AMOUNT, "0.00"))
         
         Bd.DO_TOTAL_PRICE = Val(Format(Bd.DO_TOTAL_PRICE, "0.00"))
         Bd.REVENUE_TOTAL_PRICE = Val(Format(Bd.REVENUE_TOTAL_PRICE, "0.00"))
         Bd.PAID_AMOUNT = Val(Format(Bd.PAID_AMOUNT, "0.00"))
         
            
         Set HeadCf = m_Headers1.Fields(1)
         TempStr = I
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(2)
         TempStr = Bd.CUSTOMER_CODE
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(3)
         TempStr = Bd.CUSTOMER_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(4)
         TempStr = DateToStringExtEx2(Bd.DOCUMENT_DATE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(5)
         TempStr = Bd.DOCUMENT_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(6)
         TempStr = FormatNumber(Bd.DO_TOTAL_PRICE + Bd.REVENUE_TOTAL_PRICE - Bd.DISCOUNT_AMOUNT)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(7)
         TempStr = FormatNumber(Bd.DEBIT_AMOUNT - Bd.CREDIT_AMOUNT - Bd.CASH_DISCOUNT - Bd.RETURN_AMOUNT)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(8)
         TempStr = FormatNumber(Bd.PAID_AMOUNT)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(9)
         TempStr = FormatNumber(Bd.DO_TOTAL_PRICE + Bd.REVENUE_TOTAL_PRICE - Bd.DISCOUNT_AMOUNT + (Bd.DEBIT_AMOUNT - Bd.CREDIT_AMOUNT - Bd.RETURN_AMOUNT) - Bd.PAID_AMOUNT - Bd.CASH_DISCOUNT)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         If (ROUND(Bd.DO_TOTAL_PRICE + Bd.REVENUE_TOTAL_PRICE - Bd.DISCOUNT_AMOUNT + (Bd.DEBIT_AMOUNT - Bd.CREDIT_AMOUNT) - Bd.PAID_AMOUNT - Bd.CASH_DISCOUNT, 2) <= 0) And (mcolParam("INCLUDE_FLAG") = "0") Then
            I = I - 1
         Else
            If (Bd.DO_TOTAL_PRICE + Bd.REVENUE_TOTAL_PRICE - Bd.DISCOUNT_AMOUNT + (Bd.DEBIT_AMOUNT - Bd.CREDIT_AMOUNT) - Bd.PAID_AMOUNT - Bd.CASH_DISCOUNT) > 0 Then
               Total1(6) = Total1(6) + Bd.DO_TOTAL_PRICE + Bd.REVENUE_TOTAL_PRICE - Bd.DISCOUNT_AMOUNT
               Total1(7) = Total1(7) + (Bd.DEBIT_AMOUNT - Bd.CREDIT_AMOUNT - Bd.RETURN_AMOUNT) - Bd.CASH_DISCOUNT
               Total1(8) = Total1(8) + Bd.PAID_AMOUNT
               Total1(9) = Total1(9) + (Bd.DO_TOTAL_PRICE + Bd.REVENUE_TOTAL_PRICE - Bd.DISCOUNT_AMOUNT + (Bd.DEBIT_AMOUNT - Bd.CREDIT_AMOUNT - Bd.RETURN_AMOUNT) - Bd.PAID_AMOUNT - Bd.CASH_DISCOUNT)
               
               Total2(6) = Total2(6) + Bd.DO_TOTAL_PRICE + Bd.REVENUE_TOTAL_PRICE - Bd.DISCOUNT_AMOUNT
               Total2(7) = Total2(7) + (Bd.DEBIT_AMOUNT - Bd.CREDIT_AMOUNT - Bd.RETURN_AMOUNT) - Bd.CASH_DISCOUNT
               Total2(8) = Total2(8) + Bd.PAID_AMOUNT
               Total2(9) = Total2(9) + (Bd.DO_TOTAL_PRICE + Bd.REVENUE_TOTAL_PRICE - Bd.DISCOUNT_AMOUNT + (Bd.DEBIT_AMOUNT - Bd.CREDIT_AMOUNT - Bd.RETURN_AMOUNT) - Bd.PAID_AMOUNT - Bd.CASH_DISCOUNT)
            
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
            End If
         End If
         Rs.MoveNext
      Wend

      If PrevKey1 <> NVLS(Rs("CUSTOMER_CODE"), "") Then
         If I > 1 Then
            PrevKey1 = NVLS(Rs("CUSTOMER_CODE"), "")
            If Total1(9) > 0 Then
               Call GenerateFooter("รวม", "", "", Total1, m_Details1, 1)
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
            End If
            For J = 1 To UBound(Total1)
               Total1(J) = 0
            Next J
         End If
      End If

      Call GenerateFooter("สุทธิ", "", "", Total2, m_Details1, 0)
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
      ElseIf (J = 6) Or (J = 7) Or (J = 8) Or (J = 9) Then
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


