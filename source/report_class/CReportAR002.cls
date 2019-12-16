VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAR002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAR002"
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
Private m_PaidBalanceAmount As Collection
Private m_MaxMins As Collection
Private m_ApSupAmounts As Collection
Private m_ApSupPaids As Collection
Private m_ARBalanceAmount As Collection
Private m_ARDiscountAmount As Collection

Private m_DnItemsByCust As Collection
Private m_CnItemsByCust As Collection
Private m_RtItemsByCust As Collection

Private m_DnBalancesByCust As Collection
Private m_CnBalanceByCust As Collection
Private m_RtBalanceByCust As Collection

Private m_DiscountByCust As Collection

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
   Set m_PaidBalanceAmount = New Collection
   Set m_MaxMins = New Collection
   Set m_ApSupAmounts = New Collection
   Set m_ApSupPaids = New Collection
   Set m_ARBalanceAmount = New Collection
   Set m_DnItemsByCust = New Collection
   Set m_CnItemsByCust = New Collection
   Set m_RtItemsByCust = New Collection
   
   Set m_DnBalancesByCust = New Collection
   Set m_CnBalanceByCust = New Collection
   Set m_RtBalanceByCust = New Collection
   
   Set m_ARDiscountAmount = New Collection
   Set m_DiscountByCust = New Collection
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
   Set m_PaidBalanceAmount = Nothing
   Set m_MaxMins = Nothing
   Set m_ARBalanceAmount = Nothing
   Set m_DnItemsByCust = Nothing
   Set m_CnItemsByCust = Nothing
   Set m_RtItemsByCust = Nothing
   Set m_DnBalancesByCust = Nothing
   Set m_CnBalanceByCust = Nothing
   Set m_RtBalanceByCust = Nothing
   Set m_ARDiscountAmount = Nothing
   Set m_DiscountByCust = Nothing
   
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
   
   Set Cf = New CReportField
      
   '===
   '===
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "รหัสลูกค้า", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(7, "^", "ชื่อลูกค้า", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "ยอดยกมา", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "ยอดหนี้", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "รวม", ">")
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
Dim ToPaidDate As String
Dim NewDate As Date

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
   
   NewDate = DateAdd("D", -1, mcolParam("FROM_DOC_DATE"))
   If mcolParam("FROM_DOC_DATE") > 0 Then
      Call LoadPaidAmountByCustomer(Nothing, m_PaidBalanceAmount, -1, NewDate)
      Call LoadDiscountAmountByCustomer(Nothing, m_ARDiscountAmount, -1, NewDate)
      Call LoadTotalPriceByCustomer(Nothing, m_ARBalanceAmount, -1, NewDate, 1)
   
      Call LoadDnCnAmountByCustomer(Nothing, m_CnBalanceByCust, -1, NewDate, 3, 2) '1 ของเดิมใช้ date type = 1
      Call LoadDnCnAmountByCustomer(Nothing, m_DnBalancesByCust, -1, NewDate, 4, 2) '1
      Call LoadDnCnAmountByCustomer(Nothing, m_RtBalanceByCust, -1, NewDate, 18, 2) '1
   End If
   
   Call LoadTotalPriceByCustomer(Nothing, m_ApSupAmounts, mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"), 1)
   Call LoadPaidAmountByCustomer(Nothing, m_ApSupPaids, mcolParam("FROM_DOC_DATE"), mcolParam("TO_PAY_DATE"))
   Call LoadDiscountAmountByCustomer(Nothing, m_DiscountByCust, mcolParam("FROM_DOC_DATE"), mcolParam("TO_DOC_DATE"))
   Call LoadDnCnAmountByCustomer(Nothing, m_CnItemsByCust, mcolParam("FROM_DOC_DATE"), mcolParam("TO_PAY_DATE"), 3, 2)
   Call LoadDnCnAmountByCustomer(Nothing, m_DnItemsByCust, mcolParam("FROM_DOC_DATE"), mcolParam("TO_PAY_DATE"), 4, 2)
   Call LoadDnCnAmountByCustomer(Nothing, m_RtItemsByCust, mcolParam("FROM_DOC_DATE"), mcolParam("TO_PAY_DATE"), 18, 2)
   

   
   FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DOC_DATE")), "N/A")
   ToDate = " ถึง วันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DOC_DATE")), "N/A")
   
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
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim Width As Double
Dim TempRs As ADODB.Recordset
Dim Ap As CCustomer
Dim TempApt As CDoItem
Dim Apt As CBillingDoc
Dim TempApRcIt As CReceiptItem
Dim Hc As CReportField
Dim Dept As Double
Dim DeptBalance As Double
Dim Ri1_0 As CReceiptItem
Dim Ri1_1 As CReceiptItem
Dim Ri1_2 As CReceiptItem
Dim Bdc As CBillingDiscount

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
   
      I = 0
      Set BodyCf = New CReportField

      Set Ap = New CCustomer
      Ap.CUSTOMER_CODE = mcolParam("CUSTOMER_CODE")
      
      Ap.FROM_CUSTOMER_CODE = mcolParam("FROM_CUSTOMER_CODE")
      Ap.TO_CUSTOMER_CODE = mcolParam("TO_CUSTOMER_CODE")
      
      Ap.OrderBy = mcolParam("ORDER_TYPE")
      Ap.OrderBy = 1                   'บังคับเลยเนื่องจากพอทำแล้วของเดิมมันเป็น Value 3 และ 4 ซึ่งไม่ตรงกันรัหว่างหน้า SummaryReport และ CCustomer
      Ap.OrderType = mcolParam("ORDER_BY")
      Ap.OrderType = 1                   'บังคับเลยเนื่องจากพอทำแล้วของเดิมมันเป็น Value 3 และ 4 ซึ่งไม่ตรงกันรัหว่างหน้า SummaryReport และ CCustomer
      Call glbDaily.QueryCustomer(Ap, Rs, iCount, IsOK, glbErrorLog)

      I = 0
      While Not Rs.EOF
'         For j = 1 To UBound(Total1)
'            Total1(j) = 0
'         Next j
         
         Call m_Details1.ClearField
         Call Ap.PopulateFromRS(1, Rs)

         'ยกมา
         Set TempApt = GetDoItem(m_ARBalanceAmount, Ap.CUSTOMER_ID)
         Set TempApRcIt = GetReceiptItem(m_PaidBalanceAmount, Ap.CUSTOMER_ID)
         Set Ri1_0 = GetReceiptItem(m_CnBalanceByCust, Ap.CUSTOMER_ID)
         Set Ri1_1 = GetReceiptItem(m_DnBalancesByCust, Ap.CUSTOMER_ID)
         Set Ri1_2 = GetReceiptItem(m_RtBalanceByCust, Ap.CUSTOMER_ID)                                                             'ถึงตรงนี้
         Set Bdc = GetBillingDiscount(m_ARDiscountAmount, Ap.CUSTOMER_ID)
         
         If Ri1_2.DEBIT_CREDIT_AMOUNT > 0 Then
            ''Debug.Print ("")
         End If
          
         DeptBalance = TempApt.TOTAL_PRICE - Bdc.DISCOUNT_AMOUNT + (Ri1_1.DEBIT_CREDIT_AMOUNT - Ri1_0.DEBIT_CREDIT_AMOUNT - Ri1_2.DEBIT_CREDIT_AMOUNT) - TempApRcIt.PAID_AMOUNT - TempApRcIt.CASH_DISCOUNT
         DeptBalance = Val(Format(DeptBalance, "0.00"))
         
         Set TempApt = GetDoItem(m_ApSupAmounts, Ap.CUSTOMER_ID)
         Set TempApRcIt = GetReceiptItem(m_ApSupPaids, Ap.CUSTOMER_ID)
         Set Ri1_0 = GetReceiptItem(m_CnItemsByCust, Ap.CUSTOMER_ID)
         Set Ri1_1 = GetReceiptItem(m_DnItemsByCust, Ap.CUSTOMER_ID)
         Set Ri1_2 = GetReceiptItem(m_RtItemsByCust, Ap.CUSTOMER_ID)
         Set Bdc = GetBillingDiscount(m_DiscountByCust, Ap.CUSTOMER_ID)
         
         Dept = TempApt.TOTAL_PRICE - Bdc.DISCOUNT_AMOUNT + (Ri1_1.DEBIT_CREDIT_AMOUNT - Ri1_0.DEBIT_CREDIT_AMOUNT - Ri1_2.DEBIT_CREDIT_AMOUNT) - TempApRcIt.PAID_AMOUNT - TempApRcIt.CASH_DISCOUNT
         Dept = Val(Format(Dept, "0.00"))

         If Dept + DeptBalance <> 0 Then
            I = I + 1
            Set Hc = m_Headers1.Fields(1)
            TempStr = I
            Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
   
            Set Hc = m_Headers1.Fields(2)
            TempStr = Ap.CUSTOMER_CODE
            Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
   
             Set Hc = m_Headers1.Fields(3)
            TempStr = Ap.CUSTOMER_NAME
            Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
   
             Set Hc = m_Headers1.Fields(4)
            TempStr = FormatNumber(DeptBalance)
            Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(4) = Total1(4) + DeptBalance
   
             Set Hc = m_Headers1.Fields(5)
            TempStr = FormatNumber(Dept)
            Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(5) = Total1(5) + Dept
            
             Set Hc = m_Headers1.Fields(6)
            TempStr = FormatNumber(Dept + DeptBalance)
            Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(6) = Total1(6) + Dept + DeptBalance
   
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
         End If
         '===
         Rs.MoveNext
      Wend

      Call GenerateFooter("รวม", "", "", Total1, m_Details1, 0)
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




