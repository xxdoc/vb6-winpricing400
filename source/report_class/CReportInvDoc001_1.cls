VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportInvDoc001_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportInvDoc001_1"
Private MODULE_DESC  As String

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

Private Const TITLE_SIZE = 16
Private Const HEADER_SIZE = 14
Private Const DETAIL_SIZE = 14

Private Const TITLE_FORMAT = "<32|<3|<5;"
Private Const HEADER_FORMAT = "^1|^2|^5|^2|^2|^2|^2;"
Private Const DETAIL_FORMAT = "^1|^2|<5|>2|>2|>2|<2;"
Private Const FOOTER_FORMAT_2 = "^1|>2|^5|^2|>2|^2|<2;"
Private Const FOOTER_FORMAT_1 = "^1|>2|^5|^2|>2|^2|<2;"

Private m_Enterprise As CEnterprise
Private m_CustAddress As CAddress
Private m_EnpAddress As CAddress
Private m_EnpPerson As CName
Private m_Addr As CAddress
Private m_cData As CInventoryDoc
Private m_OwnerAddress As CAddress

Private m_Headers0 As CFieldList
Private m_Headers1 As CFieldList
Private m_Details1 As CFieldList
Private m_Align(100) As String

Private Sp As CSystemParam

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_Enterprise = New CEnterprise
   Set m_CustAddress = New CAddress
   Set m_EnpAddress = New CAddress
   Set m_EnpPerson = New CName
   Set m_Addr = New CAddress
   Set m_OwnerAddress = New CAddress
   
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Details1 = New CFieldList
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_Enterprise = Nothing
   Set m_CustAddress = Nothing
   Set m_EnpAddress = Nothing
   Set m_EnpPerson = Nothing
   Set m_Addr = Nothing
   Set m_OwnerAddress = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
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

Private Sub PrintBillFooter(alngX() As Long)
Dim strFormat As String
Dim strPrint As String
Dim TimeInStr As String
Dim TimeOutStr As String

   Vsp.TableBorder = tbNone
   Vsp.FontBold = False
   
'   Vsp = ""
   
   strFormat = VSP_CalTable("^10|^10|^10;", mdblWid, alngX)
   strPrint = "ผู้ส่งของ .....................................................|ผู้ตรวจรับ .....................................................|ผู้อนุมัติ .....................................................;"
   Call Vsp.AddTable(strFormat, "", strPrint)
      
   TimeInStr = Format(HOUR(m_cData.ENTRY_DATE), "00") & ":" & Format(Minute(m_cData.ENTRY_DATE), "00")
   TimeOutStr = Format(HOUR(m_cData.EXIT_DATE), "00") & ":" & Format(Minute(m_cData.EXIT_DATE), "00")
      
   strFormat = VSP_CalTable("<10|<24|^10;", mdblWid, alngX)
   strPrint = "||FM-PN-34 Rev. 00;"
   Call Vsp.AddTable(strFormat, "", strPrint)
End Sub

Private Sub PrintBillHeader(alngX() As Long)
'Dim strFormat As String
'Dim strPrint As String
'Static I As Long
'Dim iCount As Long
'Dim TempRs As ADODB.Recordset
'Dim OwnerFlag As String
'Dim TempStr As String
'Dim TempAddr As String

End Sub

Private Sub SetReportAlign(F1 As String, F2 As String, F3 As String, F4 As String, F5 As String, F6 As String, F7 As String)
   m_Align(1) = F1
   m_Align(2) = F2
   m_Align(3) = F3
   m_Align(4) = F4
   m_Align(5) = F5
   m_Align(6) = F6
   m_Align(7) = F7
End Sub

Private Sub PrintReportBody(F1 As String, F2 As String, F3 As String, F4 As String, F5 As String, F6 As String, F7 As String, FontBold As Boolean)
Dim Cf As CReportField
Dim strFormat As String
Dim strPrint As String
Dim TempStr1 As String
Dim TempStr2 As String
Dim alngX() As Long
Dim TempFontBold As Boolean

   TempFontBold = Vsp.FontBold
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField

   Set Cf = New CReportField

   Call Cf.SetFieldValue(11, m_Align(1), F1, "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(16, m_Align(2), F2, "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(12, m_Align(3), F3, "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(28, m_Align(4), F4, "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(14, m_Align(5), F5, "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(12, m_Align(6), F6, "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(5, m_Align(7), F7, "^")
   Call m_Headers1.AddField(Cf)
   
  Call m_Headers1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr2
   Vsp.FontBold = FontBold
'Vsp.TableBorder = tbAll
   Call Vsp.AddTable(strFormat, "", strPrint)
   Vsp.FontBold = TempFontBold
End Sub

Private Sub PrintReportNote(F1 As String, F2 As String, F3 As String, F4 As String)
Dim Cf As CReportField
Dim strFormat As String
Dim strPrint As String
Dim TempStr1 As String
Dim TempStr2 As String
Dim alngX() As Long

   Call m_Headers0.ClearField
   Call m_Headers1.ClearField

   Set Cf = New CReportField

   Call Cf.SetFieldValue(10, m_Align(1), F1, "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(12.5 + 12.5 + 25 + 5, m_Align(2), F2, "^")
   Call m_Headers1.AddField(Cf)

   Call Cf.SetFieldValue(15, m_Align(3), F3, "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(10 + 5, m_Align(4), F4, "^")
   Call m_Headers1.AddField(Cf)
   
  Call m_Headers1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr2
'Vsp.TableBorder = tbAll
   Call Vsp.AddTable(strFormat, "", strPrint)
End Sub

Private Sub printHeader()
Dim strFormat As String
Dim strPrint As String
Dim tbdT As TableBorderSettings
Dim blnBold As Boolean
Dim iSize As Integer
Dim alngX() As Long
Dim Li As CLotItem
Dim ActualWeight As Double
Dim TotalPrice As Double
Dim Cf As CReportField
Dim TotalWeight As Double
Dim NetWeight As Double
Dim Exp1 As Double
Dim Exp2 As Double
Dim ExtraName As String

   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
      
   Vsp.FontSize = DETAIL_SIZE
   Vsp.FontBold = False
   Vsp.TableBorder = tbNone
   
   Call SetReportAlign("<", "<", "<", "<", "<", "<", "<")
   Call PrintReportBody("เลขที่ใบสั่งซื้อ", m_cData.DO_NO, "วันที่เอกสาร", DateToStringExtEx2(m_cData.DOCUMENT_DATE), "เลขที่รับ", m_cData.DOCUMENT_NO, "", True)
   If mcolParam("REPORT_TYPE") = 8 Or mcolParam("REPORT_TYPE") = 9 Then
      Call PrintReportBody("รหัสผู้ขาย", m_cData.SUPPLIER_CODE_TRUE, "ชื่อผู้ขาย", m_cData.SUPPLIER_NAME_TRUE, "", "", "", False)
   Else
      Call PrintReportBody("รหัสผู้ขาย", m_cData.SUPPLIER_CODE, "ชื่อผู้ขาย", m_cData.SUPPLIER_NAME, "", "", "", False)
   End If

   Set Li = m_cData.ImportExports(1)
   
   
   Call PrintReportBody("รหัส R/M", Li.PART_NO, "ชื่อ R/M", Li.PART_DESC, "", "", "", False)
   
   
   Call SetReportAlign("<", "<", "<", "<", "<", ">", "<")
   Call PrintReportBody("รหัสโกดัง", Li.LOCATION_NO, "ชื่อโกดัง", Li.LOCATION_NAME, "น้ำหนักเข้า", FormatNumber(Li.ENTRY_WEIGHT), "Kg", False)
   Call PrintReportBody("เวลารถเข้า", Right$(DateToStringExtEx(m_cData.ENTRY_DATE), 8), "เวลารถออก", Right$(DateToStringExtEx(m_cData.EXIT_DATE), 8), "น้ำหนักออก", FormatNumber(Li.EXIT_WEIGHT), "Kg", False)
   TotalWeight = Li.WEIGHT_AMOUNT   'Li.ENTRY_WEIGHT - Li.EXIT_WEIGHT
   Call PrintReportBody("ทะเบียนรถ", m_cData.TRUCK_NO, "คิวที่", m_cData.QUE_NO, "น้ำหนักรวม", FormatNumber(TotalWeight), "Kg", True)
   'ถ้าคืนถุงน้ำหนักสุทธิ = น้ำหนักรวม
   'ถ้าไม่คืนถุงน้ำหนักสุทธิ = น้ำหนักรวม - น้ำหนักถุง
   If Li.BAG_RETURN = "N" Then
      TotalWeight = TotalWeight - Li.PACKAGE_WEIGHT
   End If
   Call PrintReportBody("จำนวนบรรจุ", Li.PACKAGE_AMOUNT, "จำนวนรับจริง", Li.ACTUAL_PKG_AMOUNT & "                [" & Li.PACKAGING_CODE & "]", "ตัด น.น. บรรจุ", FormatNumber(Li.PACKAGE_WEIGHT), "Kg", False)
   Call PrintReportBody("% ความชื้น", FormatNumber(Li.PERCENT_HUMID), "", "", "ตัด น.น. Moist", FormatNumber(Li.HUMID_WEIGHT), "Kg", False)
   If Li.OTHER_WEIGHT > 0 Then
      If Len(Trim(Li.EXTRA_NAME)) <= 0 Then
         ExtraName = "ตัด น.น. อื่น ๆ"
      Else
         ExtraName = Li.EXTRA_NAME
      End If
      Call PrintReportBody("ตัดอื่น ๆ", "", "", "", ExtraName, FormatNumber(Li.OTHER_WEIGHT), "Kg", False)
   End If
   NetWeight = TotalWeight - Li.OTHER_WEIGHT - Li.HUMID_WEIGHT
   Call PrintReportBody("", "", "", "", "** น.น. สุทธิ **", FormatNumber(NetWeight), "Kg", True)
   If mcolParam("REPORT_TYPE") = 3 Or mcolParam("REPORT_TYPE") = 9 Then
      Call PrintReportBody("หน่วย R/M", "", "ราคา ", "" & " บาท", "คิดเป็นเงิน", "", "บาท", True)
   ElseIf mcolParam("REPORT_TYPE") = 8 Then
      Call PrintReportBody("หน่วย R/M", "", "ราคา ", FormatNumber(isDiff(Li.ACTUAL_UNIT_PRICE), 2) & " บาท", "คิดเป็นเงิน", FormatNumber(calCalulateType(Li, FormatNumber(isDiff(Li.ACTUAL_UNIT_PRICE), 2))), "บาท", True)
   Else
      Call PrintReportBody("หน่วย R/M", "", "ราคา ", FormatNumber(Li.ACTUAL_UNIT_PRICE, 2) & " บาท", "คิดเป็นเงิน", FormatNumber(Li.TOTAL_ACTUAL_PRICE), "บาท", True)
   End If
   Exp1 = 0
   If Li.PUREXP_ID1 > 0 Then
      Exp1 = Li.EXPENSE1
      Call PrintReportBody("คชจ. จัดซื้อ", "", "", "", Li.PUREXP_NAME1, FormatNumber(Exp1), "บาท", False)
   End If
   
   Exp2 = 0
   If Li.PUREXP_ID2 > 0 Then
      Exp2 = Li.EXPENSE2
      Call PrintReportBody("คชจ. จัดซื้อ", "", "", "", Li.PUREXP_NAME2, FormatNumber(Exp2), "บาท", False)
   End If
   
   Call SetReportAlign("<", ">", "<", "<", "<", ">", "<")
   If mcolParam("REPORT_TYPE") = 3 Or mcolParam("REPORT_TYPE") = 9 Then
      Call PrintReportBody("น.น. ผู้ขาย", FormatNumber(Li.SUPPLIER_WEIGHT), "Kg", "", "เป็นเงินสุทธิ", "", "บาท", True)
   ElseIf mcolParam("REPORT_TYPE") = 8 Then
      Call PrintReportBody("น.น. ผู้ขาย", FormatNumber(Li.SUPPLIER_WEIGHT), "Kg", "", "เป็นเงินสุทธิ", FormatNumber(calCalulateType(Li, FormatNumber(isDiff(Li.ACTUAL_UNIT_PRICE), 2)) + Exp1 + Exp2), "บาท", True)
   Else
      Call PrintReportBody("น.น. ผู้ขาย", FormatNumber(Li.SUPPLIER_WEIGHT), "Kg", "", "เป็นเงินสุทธิ", FormatNumber(Li.TOTAL_ACTUAL_PRICE + Exp1 + Exp2), "บาท", True)
   End If
   Call SetReportAlign("<", "<", "<", "<", "", "", "")
   If (mcolParam("REPORT_TYPE") = 1 Or mcolParam("REPORT_TYPE") = 3) And Len(m_cData.SUPPLIER_CODE_TRUE) > 0 Then
      Call PrintReportNote("หมายเหตุ", m_cData.DOCUMENT_DESC & " [" & m_cData.SUPPLIER_CODE_TRUE & ":" & m_cData.SUPPLIER_NAME_TRUE & "]", "[" & Li.CALCULATE_TYPE & "]", CalculateTypeToText(Li.CALCULATE_TYPE))
   Else
      Call PrintReportNote("หมายเหตุ", m_cData.DOCUMENT_DESC, "[" & Li.CALCULATE_TYPE & "]", CalculateTypeToText(Li.CALCULATE_TYPE))
   End If
   
   Vsp.TableBorder = tbdT
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
End Sub
Private Function calCalulateType(m_Li As CLotItem, ActualUnitPrice As Double) As Double
Dim TempData As Double
Dim TotalWeight As Double
Dim NetWeight As Double
   If m_Li.CALCULATE_TYPE = 1 Then
       TempData = ActualUnitPrice * m_Li.SUPPLIER_WEIGHT
   ElseIf m_Li.CALCULATE_TYPE = 2 Then
      TempData = ActualUnitPrice * m_Li.WEIGHT_AMOUNT
   ElseIf m_Li.CALCULATE_TYPE = 3 Then
      TotalWeight = m_Li.WEIGHT_AMOUNT
      If m_Li.BAG_RETURN = "N" Then
         TotalWeight = TotalWeight - m_Li.PACKAGE_WEIGHT
      End If
     NetWeight = TotalWeight - m_Li.OTHER_WEIGHT - m_Li.HUMID_WEIGHT
     TempData = ActualUnitPrice * NetWeight
   End If
   calCalulateType = TempData
End Function
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
   
   MODULE_DESC = MapText("ใบรายงานรับวัตถุดิบ")
   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))
   
   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
      mstrErrMsg = "Error GetDateTime Error."
      Exit Function
   End If
   mdteDate = InternalDateToDate(strDate)
   initDoc = True
End Function

Private Function GetSpace() As Long

End Function

Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim RName As String
Dim iCount As Long
Dim I As Long
Dim strFormat As String
Dim strPrint As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Bi As CLotItem
Dim Sum1 As Double
Dim Sum2 As Double
Dim Sum3 As Double
Dim DiscountAmount As Double
Dim VatAmount As Double
Dim IncludeDiscountAmount As Double
Dim IncludeVatAmount As Double
Dim TempVatAmount As Double
Dim TempDeposit As Double
Dim TempDiscount As Double
Dim TempCol As Collection
Dim SpaceCount As Long
Dim Total(10) As Double
Dim J As Long

   RName = "genDoc"
   
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   Set Rs = New ADODB.Recordset
   Set m_cData = New CInventoryDoc
   m_cData.INVENTORY_DOC_ID = mcolParam("INVENTORY_DOC_ID")
   m_cData.QueryFlag = 1
   If Not glbDaily.QueryInventoryDoc(m_cData, Rs, iCount, IsOK, glbErrorLog) Then
      mstrErrMsg = "Query Error."
      Exit Function
   End If
   
   If (Rs Is Nothing) Or (iCount < 1) Then
      mstrErrMsg = "ไม่พบข้อมูลที่ต้องการในฐานข้อมูล"
      Exit Function
   End If
   
   Call m_cData.PopulateFromRS(1, Rs)
   
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
   Sum1 = 0
   Sum2 = 0
   Sum3 = 0
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
   
      Call PrintBillFooter(alngX)
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
   Call Vsp.AddTable(strFormat, "", "|หน้าที่|: " & Vsp.CurrentPage)
   Call Vsp.AddTable(strFormat, "", "|วันที่พิมพ์|: " & DateToStringExt(mdteDate))
   Call Vsp.AddTable(strFormat, "", "|เวลา|: " & Right$(DateToStringExtEx(mdteDate), 8))
   Vsp.CurrentY = dY(0)
   
   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   Vsp.TextAlign = taLeftBottom
   
   If m_cData.EXCEPTION_FLAG = "N" Then
      If mcolParam("REPORT_TYPE") = 1 Or mcolParam("REPORT_TYPE") = 3 Then
         Vsp.Paragraph = "     " & glbEnterPrise.ENTERPRISE_NAME & "         ( " & MODULE_DESC & " )"
      ElseIf mcolParam("REPORT_TYPE") = 2 Or mcolParam("REPORT_TYPE") = 8 Or mcolParam("REPORT_TYPE") = 9 Then
         Vsp.Paragraph = "     " & "บริษัทเอ็มจีฟาร์มา จำกัด" & "         ( " & MODULE_DESC & " )"
      End If
   Else
      Vsp.Paragraph = ""
   End If
   If mcolParam("REPORT_TYPE") = 1 Or mcolParam("REPORT_TYPE") = 3 Then
      Vsp.Paragraph = "     " & FindFactoryAddress
   Else
      Vsp.Paragraph = "     " & "11 ซ.รามคำแหง 81/1 ถ.รามคำแหง แขวงหัวหมาก เขตบางกระปิ กทม. 10240"
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



