VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportVoucherPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportVoucherPay"
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

Private Const TITLE_SIZE = 14
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
Private m_ReportConfig As CReportConfig
Private m_cData As CBillingDoc

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_Enterprise = New CEnterprise
   Set m_CustAddress = New CAddress
   Set m_EnpAddress = New CAddress
   Set m_EnpPerson = New cName
   Set m_ReportConfig = New CReportConfig
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_Enterprise = Nothing
   Set m_CustAddress = Nothing
   Set m_EnpAddress = Nothing
   Set m_EnpPerson = Nothing
   Set m_ReportConfig = Nothing
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
   
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
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
   
   Vsp.PaperSize = pprA4
   Vsp.ORIENTATION = orPortrait
   Vsp.MarginBottom = "1cm"
   Vsp.MarginFooter = "1cm"
   Vsp.MarginHeader = "1cm"
   Vsp.MarginLeft = "1cm"
   Vsp.MarginRight = "1cm"
   Vsp.MarginTop = "0cm"
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   MODULE_DESC = MapText("ใบสำคัญจ่าย")
   
   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"), m_ReportConfig, False)
      
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
Dim Poi As CReceiptItem
Dim TempBorder As TableBorderSettings
Dim Addr As CAddress
Dim SumDoc As String
Dim Ct As CCashTran
Dim TotalReceive As Double
Dim TotalFee As Double
Dim TempStr As String
Dim TempY As Double
Dim TempY2 As Double

Dim Gl As CGLDetail
Dim SumDr As Double
Dim SumCr As Double
Dim SumTotalPrice As Double

   RName = "genDoc"
   
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   Set Rs = New ADODB.Recordset
   Set m_cData = New CBillingDoc
   m_cData.BILLING_DOC_ID = mcolParam("BILLING_DOC_ID")
   m_cData.QueryFlag = 1
   If Not glbDaily.QueryBillingDoc(m_cData, Rs, iCount, IsOK, glbErrorLog) Then
      mstrErrMsg = "Query Error."
      Exit Function
   End If
      
   Call m_cData.PopulateFromRS(1, Rs)
   Call mcolParam.add(m_cData.DOCUMENT_NO, "DOCUMENT_NO")
   
   Call mcolParam.add(m_cData.PV_NO, "PV_NO")
   Call mcolParam.add(m_cData.DOCUMENT_DATE, "DOCUMENT_DATE")
   Call mcolParam.add(m_cData.SUPPLIER_NAME, "SUPPLIER_NAME")
   Call mcolParam.add(m_cData.NOTE, "NOTE")
   Call mcolParam.add(m_cData.DUE_DATE, "DUE_DATE")
   Call mcolParam.add(m_cData.TOTAL_PRICE, "TOTAL_PRICE")
   
'   If m_cData.BILLING_ADDRESS_ID > 0 Then
'      Set Addr = New CAddress
'      Addr.ADDRESS_ID = m_cData.BILLING_ADDRESS_ID
'      Call Addr.QueryData(Rs, iCount)
'      If Not Rs.EOF Then
'         Call Addr.PopulateFromRS(Rs)
'         Call mcolParam.add(Addr.PackAddress, "CUSTOMER_ADDRESS")
'      End If
'      Set Addr = Nothing
'   Else
'      Call mcolParam.add("", "CUSTOMER_ADDRESS")
'   End If
   
   If (Rs Is Nothing) Or (iCount < 1) Then
      mstrErrMsg = "ไม่พบข้อมูลที่ต้องการในฐานข้อมูล"
      Exit Function
   End If
   
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
   Vsp.MarginLeft = 0
   
   strFormat = VSP_CalTable(DETAIL_FORMAT, mdblWid, alngX)
   TempBorder = tbNone
   
   Vsp.TableBorder = tbNone
   Vsp.MarginLeft = 1 * 567 - m_ReportConfig.DUMMY_OFFSET
      
   
   For Each Ct In m_cData.Payments
      TotalReceive = TotalReceive + Ct.GetFieldValue("AMOUNT")
      TotalFee = TotalFee + Ct.GetFieldValue("FEE_AMOUNT")
   Next Ct
   
   '--------------------------------------------------------------------------------------------------------------------------------------------
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 1
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcRowHeight) = 1.4 * 567 + m_ReportConfig.HEAD_OFFSET '
   Vsp.TableCell(tcColWidth, , 1) = "19.1cm"
   Vsp.EndTable
   '--------------------------------------------------------------------------------------------------------------------------------------------
   
   '--------------------------------------------------------------------------------------------------------------------------------------------
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcRowHeight) = "0.8cm"
   Vsp.TableCell(tcColWidth, , 1) = "17.5cm"
   Vsp.TableCell(tcColWidth, , 2) = "1.6cm"
   Vsp.TableCell(tcAlign, 1, 2) = taRightMiddle
   Vsp.TableCell(tcText, 1, 2) = m_cData.PV_NO
   Vsp.EndTable
   '--------------------------------------------------------------------------------------------------------------------------------------------
      
      '--------------------------------------------------------------------------------------------------------------------------------------------
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcRowHeight) = "0.5cm"
   Vsp.TableCell(tcColWidth, , 1) = "14.1cm"
   Vsp.TableCell(tcColWidth, , 2) = "5cm"
   Vsp.EndTable
   '--------------------------------------------------------------------------------------------------------------------------------------------
   
   '--------------------------------------------------------------------------------------------------------------------------------------------
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcRowHeight) = "0.7cm"
   Vsp.TableCell(tcColWidth, , 1) = "14.1cm"
   Vsp.TableCell(tcColWidth, , 2) = "5cm"
   Vsp.TableCell(tcAlign, 1, 2) = taCenterMiddle
   Vsp.TableCell(tcText, 1, 2) = DateToStringExtEx2(m_cData.DOCUMENT_DATE)
   Vsp.EndTable
   '--------------------------------------------------------------------------------------------------------------------------------------------
   
   '--------------------------------------------------------------------------------------------------------------------------------------------
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcRowHeight) = "0.8cm"
   Vsp.TableCell(tcColWidth, , 1) = "4cm"
   Vsp.TableCell(tcColWidth, , 2) = "15.1cm"
   Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
   Vsp.TableCell(tcText, 1, 2) = mcolParam("SUPPLIER_NAME")
   Vsp.EndTable
   '--------------------------------------------------------------------------------------------------------------------------------------------
   
   '--------------------------------------------------------------------------------------------------------------------------------------------
   I = 0
   Vsp.FontSize = 14
   For Each Poi In m_cData.ReceiptItems
   I = I + 1
     If I > 1 Then
       SumDoc = SumDoc & " , " & Poi.DOCUMENT_NO
     ElseIf I = 1 Then
         SumDoc = Poi.DOCUMENT_NO
     End If
   Next Poi
   
   Dim strArr() As String
   Dim strArr2(4) As String
   Dim R As Long
   R = 0
   J = 0
    TempY = Vsp.CurrentY
    TempY2 = Vsp.CurrentY

   If m_cData.ReceiptItems.Count > 6 Then
       strArr = Split(SumDoc, ",")
       SumDoc = ""
       For I = 1 To UBound(strArr) + 1
       J = J + 1
         If J Mod 8 = 0 Then
            strArr2(R) = SumDoc
            Vsp.StartTable
            Vsp.TableCell(tcCols) = 2
            Vsp.TableCell(tcColWidth, 1, 1) = "4.8cm"
            Vsp.TableCell(tcColWidth, 1, 2) = "14.3cm"
            Vsp.TableCell(tcAlign, 1, 2) = taLeftTop
            Vsp.FontSize = 12
            Vsp.TableCell(tcRowHeight) = "0.5cm"
            If R = 0 Then
               Vsp.CurrentY = TempY2 - 200
            Else
               Vsp.CurrentY = TempY2 - 170
            End If
            Vsp.TableCell(tcRows) = 1
            Vsp.TableCell(tcText, 1, 2) = "" & strArr2(R)
            Vsp.TableBorder = tbNone
            Vsp.EndTable
            TempY2 = Vsp.CurrentY
   
            R = R + 1
            J = 1
            SumDoc = ""
         End If
         If J = 1 Then
             SumDoc = strArr(I - 1)
         ElseIf J > 1 Then
           SumDoc = SumDoc & " , " & strArr(I - 1)
         End If
     Next I
   strArr2(R) = SumDoc
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcColWidth, 1, 1) = "4.8cm"
   Vsp.TableCell(tcColWidth, 1, 2) = "14.3cm"
   Vsp.TableCell(tcAlign, 1, 2) = taLeftTop
   Vsp.FontSize = 12
   Vsp.TableCell(tcRowHeight) = "0.5cm"
   If R = 0 Then
      Vsp.CurrentY = TempY2 - 200
   Else
      Vsp.CurrentY = TempY2 - 170
   End If
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcText, 1, 2) = strArr2(R)
   Vsp.TableBorder = tbNone
   Vsp.EndTable
    
   Else 'ถ้ามีแค่บรรทัดเดียว
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 2
      Vsp.TableCell(tcColWidth, 1, 1) = "4.8cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "14.3cm"
      Vsp.TableCell(tcAlign, 1, 2) = taLeftTop
      Vsp.FontSize = 14
      Vsp.TableCell(tcRowHeight) = "0.9cm"
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcText, 1, 2) = SumDoc
      Vsp.TableBorder = tbNone
      Vsp.EndTable
   End If
   
   '--------------------------------------------------------------------------------------------------------------------------------------------
   
   '--------------------------------------------------------------------------------------------------------------------------------------------

   Vsp.CurrentY = TempY + 512
   Vsp.FontSize = 14
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcRowHeight) = "1.6cm"
   Vsp.TableCell(tcColWidth, , 1) = "3.5cm"
   Vsp.TableCell(tcColWidth, , 2) = "15.6cm"
   Vsp.TableCell(tcAlign, 1, 2) = taLeftTop
   Vsp.TableCell(tcText, 1, 2) = mcolParam("NOTE")
   Vsp.EndTable
   '--------------------------------------------------------------------------------------------------------------------------------------------
      
   '--------------------------------------------------------------------------------------------------------------------------------------------
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcRowHeight) = "0.8cm"
   Vsp.TableCell(tcColWidth, , 1) = "3.5cm"
   Vsp.TableCell(tcColWidth, , 2) = "15.6cm"
   Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
   Vsp.TableCell(tcText, 1, 2) = DateToStringExtEx2(m_cData.DUE_DATE)
   Vsp.EndTable
   '--------------------------------------------------------------------------------------------------------------------------------------------
   
   '--------------------------------------------------------------------------------------------------------------------------------------------
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 1
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcRowHeight) = "0.6cm"
   Vsp.EndTable
   '--------------------------------------------------------------------------------------------------------------------------------------------
   
   '--------------------------------------------------------------------------------------------------------------------------------------------
   SumTotalPrice = 0
   For Each Gl In m_cData.GlDetails
      If Gl.GetFieldValue("SUM_FLAG") = "Y" And Gl.GetFieldValue("GL_TYPE") = 2 Then
         SumTotalPrice = SumTotalPrice + Gl.GetFieldValue("GL_AMOUNT")
      End If
   Next Gl
   
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 4
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcRowHeight) = "0.75cm" '
   Vsp.TableCell(tcColWidth, , 1) = "2cm"
   Vsp.TableCell(tcColWidth, , 2) = "5cm"
   Vsp.TableCell(tcColWidth, , 3) = "1cm"
   Vsp.TableCell(tcColWidth, , 4) = "11cm"
   
   Vsp.TableCell(tcAlign, 1, 2) = taCenterMiddle
   Vsp.TableCell(tcText, 1, 2) = FormatNumber(SumTotalPrice)
   
   Vsp.TableCell(tcAlign, 1, 4) = taLeftMiddle
   Vsp.TableCell(tcText, 1, 4) = ThaiBaht(SumTotalPrice)
   
   Vsp.EndTable
   '--------------------------------------------------------------------------------------------------------------------------------------------
   
   Vsp.FontSize = 12

   If m_cData.Payments.Count = 1 Then
      Set Ct = m_cData.Payments(1)
      If Ct.GetFieldValue("PAYMENT_TYPE") = 1 Then
         Vsp.StartTable
         Vsp.TableCell(tcCols) = 1
         Vsp.TableCell(tcRows) = 1
         Vsp.TableCell(tcRowHeight) = "0.15cm" '
         Vsp.TableCell(tcColWidth, , 1) = "19cm"
         Vsp.EndTable
         
         Vsp.StartTable
         Vsp.TableCell(tcCols) = 3
         Vsp.TableCell(tcRows) = 1
         Vsp.TableCell(tcRowHeight) = "0.4cm" '
         Vsp.TableCell(tcColWidth, , 1) = "1.5cm"
         Vsp.TableCell(tcColWidth, , 2) = "0.5cm"
         Vsp.TableCell(tcColWidth, , 3) = "17cm"
         Vsp.TableCell(tcAlign, 1, 2) = taCenterTop
         Vsp.TableCell(tcText, 1, 2) = "X"
         Vsp.EndTable
         
         
         Vsp.StartTable
         Vsp.TableCell(tcCols) = 1
         Vsp.TableCell(tcRows) = 1
         Vsp.TableCell(tcRowHeight) = "3.15cm" '
         Vsp.TableCell(tcColWidth, , 1) = "19cm"
         Vsp.EndTable

'      ElseIf Ct.GetFieldValue("PAYMENT_TYPE") = 2 Then
'         Vsp.StartTable
'         Vsp.TableCell(tcCols) = 1
'         Vsp.TableCell(tcRows) = 1
'         Vsp.TableCell(tcRowHeight) = "1.2cm" '
'         Vsp.TableCell(tcColWidth, , 1) = "19cm"
'         Vsp.EndTable
'
'         Vsp.StartTable
'         Vsp.TableCell(tcCols) = 3
'         Vsp.TableCell(tcRows) = 1
'         Vsp.TableCell(tcRowHeight) = "0.45cm" '
'         Vsp.TableCell(tcColWidth, , 1) = "1.5cm"
'         Vsp.TableCell(tcColWidth, , 2) = "0.5cm"
'         Vsp.TableCell(tcColWidth, , 3) = "17cm"
'         Vsp.TableCell(tcAlign, 1, 2) = taCenterMiddle
'         Vsp.TableCell(tcText, 1, 2) = "X"
'         Vsp.EndTable
'
'         Vsp.StartTable
'         Vsp.TableCell(tcCols) = 1
'         Vsp.TableCell(tcRows) = 1
'         Vsp.TableCell(tcRowHeight) = "3.8cm" '
'         Vsp.TableCell(tcColWidth, , 1) = "19cm"
'         Vsp.EndTable
      ElseIf Ct.GetFieldValue("PAYMENT_TYPE") = 3 Then
         Vsp.StartTable
         Vsp.TableCell(tcCols) = 1
         Vsp.TableCell(tcRows) = 1
         Vsp.TableCell(tcRowHeight) = "1.05cm" '
         Vsp.TableCell(tcColWidth, , 1) = "19cm"
         Vsp.EndTable
         
         Vsp.StartTable
         Vsp.TableCell(tcCols) = 6
         Vsp.TableCell(tcRows) = 2
         Vsp.TableCell(tcRowHeight, 1) = "0.4cm" '
         Vsp.TableCell(tcRowHeight, 2) = "1.25cm" '
         Vsp.TableCell(tcColWidth, , 1) = "2.05cm"
         Vsp.TableCell(tcColWidth, , 2) = "0.45cm"
         Vsp.TableCell(tcColWidth, , 3) = "1.5cm"
         Vsp.TableCell(tcColWidth, , 4) = "3.5cm"
         Vsp.TableCell(tcColWidth, , 5) = "1.5cm"
         Vsp.TableCell(tcColWidth, , 6) = "10.1cm"
         Vsp.TableCell(tcAlign, 1, 2) = taCenterTop
         Vsp.TableCell(tcText, 1, 2) = "X"
         Vsp.TableCell(tcAlign, 1, 4) = taCenterTop
         Vsp.TableCell(tcText, 1, 4) = Ct.GetFieldValue("CHEQUE_NO")
         Vsp.TableCell(tcAlign, 1, 6) = taLeftTop
         Vsp.TableCell(tcText, 1, 6) = DateToStringExtEx2(Ct.GetFieldValue("CHEQUE_DATE"))
         Vsp.TableCell(tcAlign, 2, 4) = taCenterMiddle
         Vsp.TableCell(tcText, 2, 4) = Ct.GetFieldValue("BANK_NAME")
         Vsp.TableCell(tcAlign, 2, 6) = taLeftMiddle
         Vsp.TableCell(tcText, 2, 6) = Ct.GetFieldValue("BRANCH_NAME")
         Vsp.EndTable
         
         Vsp.StartTable
         Vsp.TableCell(tcCols) = 1
         Vsp.TableCell(tcRows) = 1
         Vsp.TableCell(tcRowHeight) = "1.6cm" '
         Vsp.TableCell(tcColWidth, , 1) = "19cm"
         Vsp.EndTable
      End If
   Else
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 1
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight) = "4.2cm" '
      Vsp.TableCell(tcColWidth, , 1) = "19cm"
      Vsp.EndTable
   End If
   
   Vsp.FontSize = DETAIL_SIZE
      
   J = 0
   
   '--------------------------------------------------------------------------------------------------------------------------------------------
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 1
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcRowHeight) = "0.4cm"
   Vsp.EndTable
   '--------------------------------------
   
   For Each Gl In m_cData.GlDetails
      J = J + 1
      If J > 11 Then
         Call Vsp.NewPage
      End If
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 6
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight, 1) = "0.75cm" '
      Vsp.TableCell(tcColWidth, , 1) = "8.2cm"
      Vsp.TableCell(tcColWidth, , 2) = "2.9cm"
      Vsp.TableCell(tcColWidth, , 3) = "3cm"
      Vsp.TableCell(tcColWidth, , 4) = "1cm"
      Vsp.TableCell(tcColWidth, , 5) = "3cm"
      Vsp.TableCell(tcColWidth, , 6) = "1cm"
      
      Vsp.TableCell(tcAlign, 1, 1) = taLeftMiddle
      Vsp.TableCell(tcText, 1, 1) = Gl.GetFieldValue("GL_NAME") & " " & Gl.GetFieldValue("GL_DESC")
      Vsp.TableCell(tcAlign, 1, 2) = taCenterMiddle
      Vsp.TableCell(tcText, 1, 2) = Gl.GetFieldValue("GL_NO")
      
      If Gl.GetFieldValue("GL_TYPE") = 1 Then
         Vsp.TableCell(tcAlign, 1, 3) = taRightMiddle
         Vsp.TableCell(tcText, 1, 3) = Left(FormatNumber(Gl.GetFieldValue("GL_AMOUNT")), Len(FormatNumber(Gl.GetFieldValue("GL_AMOUNT"))) - 3)
         Vsp.TableCell(tcAlign, 1, 4) = taRightMiddle
         Vsp.TableCell(tcText, 1, 4) = Right(FormatNumber(Gl.GetFieldValue("GL_AMOUNT")), 2)
         
         SumDr = SumDr + Gl.GetFieldValue("GL_AMOUNT")
      ElseIf Gl.GetFieldValue("GL_TYPE") = 2 Then
         Vsp.TableCell(tcAlign, 1, 5) = taRightMiddle
         Vsp.TableCell(tcText, 1, 5) = Left(FormatNumber(Gl.GetFieldValue("GL_AMOUNT")), Len(FormatNumber(Gl.GetFieldValue("GL_AMOUNT"))) - 3)
         Vsp.TableCell(tcAlign, 1, 6) = taRightMiddle
         Vsp.TableCell(tcText, 1, 6) = Right(FormatNumber(Gl.GetFieldValue("GL_AMOUNT")), 2)
         SumCr = SumCr + Gl.GetFieldValue("GL_AMOUNT")
      End If
      Vsp.EndTable
   Next Gl
   
   For I = 1 To (11 - J)
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 6
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight, 1) = "0.75cm" '
      Vsp.TableCell(tcColWidth, , 1) = "8.2cm"
      Vsp.TableCell(tcColWidth, , 2) = "2.9cm"
      Vsp.TableCell(tcColWidth, , 3) = "3cm"
      Vsp.TableCell(tcColWidth, , 4) = "1cm"
      Vsp.TableCell(tcColWidth, , 5) = "3cm"
      Vsp.TableCell(tcColWidth, , 6) = "1cm"
      
      Vsp.EndTable
   Next I
   
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 6
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcRowHeight, 1) = "1.6cm" '
   Vsp.TableCell(tcColWidth, , 1) = "10.2cm"
   Vsp.TableCell(tcColWidth, , 2) = "0.9cm"
   Vsp.TableCell(tcColWidth, , 3) = "3cm"
   Vsp.TableCell(tcColWidth, , 4) = "1cm"
   Vsp.TableCell(tcColWidth, , 5) = "3cm"
   Vsp.TableCell(tcColWidth, , 6) = "1cm"
   
   Vsp.TableCell(tcAlign, 1, 1) = taLeftMiddle
   Vsp.TableCell(tcText, 1, 1) = "   " & ThaiBaht(SumCr)
   Vsp.TableCell(tcAlign, 1, 2) = taCenterMiddle
   Vsp.TableCell(tcText, 1, 2) = ""
   
   Vsp.TableCell(tcAlign, 1, 3) = taRightMiddle
   Vsp.TableCell(tcText, 1, 3) = Left(FormatNumber(SumCr), Len(FormatNumber(SumCr)) - 3)
   Vsp.TableCell(tcAlign, 1, 4) = taRightMiddle
   Vsp.TableCell(tcText, 1, 4) = Right(FormatNumber(SumCr), 2)
   
   Vsp.TableCell(tcAlign, 1, 5) = taRightMiddle
   Vsp.TableCell(tcText, 1, 5) = Left(FormatNumber(SumCr), Len(FormatNumber(SumCr)) - 3)
   Vsp.TableCell(tcAlign, 1, 6) = taRightMiddle
   Vsp.TableCell(tcText, 1, 6) = Right(FormatNumber(SumCr), 2)
   SumCr = SumCr + SumCr
   
   Vsp.EndTable
   
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
'   Vsp.Paragraph = MODULE_DESC
'   Vsp.Paragraph = ""
   If mcolParam("REPORT_TYPE") = 4 Then
      Vsp.DrawPicture mcolParam("BACK_GROUND"), "0.05cm", "0.0cm", "21.55cm", "15.5cm", 10
   End If
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


