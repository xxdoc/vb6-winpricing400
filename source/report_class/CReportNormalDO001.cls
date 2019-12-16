VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportNormalDO001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportNormalDO001"
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
Private m_FactAddress As CAddress
Private m_cData As CBillingDoc
Private m_CustAddr As CAddress
Private m_Addresses As Collection

Private m_Headers0 As CFieldList
Private m_Headers1 As CFieldList
Private m_Details1 As CFieldList
Private m_Align(100) As String

Private Sp As CSystemParam
Private TEMP_FORMAT As String
Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_Enterprise = New CEnterprise
   Set m_CustAddress = New CAddress
   Set m_FactAddress = New CAddress
   Set m_EnpPerson = New CName
   Set m_EnpAddress = New CAddress
   Set m_CustAddr = New CAddress
   Set m_Addresses = New Collection
   
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
   Set m_FactAddress = Nothing
   Set m_EnpPerson = Nothing
   Set m_EnpAddress = Nothing
   Set m_CustAddr = Nothing
   Set m_Addresses = Nothing
   
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

Private Sub PrintBillFooter1(alngX() As Long)
Dim strFormat As String
Dim strPrint As String
Dim TempStr1 As String
Dim TempStr2 As String

   Vsp.TableBorder = tbNone
   Vsp.FontBold = False
   
   Vsp = ""
   Vsp.TableBorder = tbAll
   
   TempStr1 = "ใบเสร็จรับเงิน" & vbCrLf
   TempStr1 = TempStr1 & "ได้รับเงิน/เช็ค จำนวน ............................................................" & vbCrLf
   TempStr1 = TempStr1 & "เงินสด" & vbCrLf
   TempStr1 = TempStr1 & "เช็คเลขที่ .......................................... ลงวันที่ ............/............./..........." & vbCrLf
   TempStr1 = TempStr1 & "ธนาคาร   ..........................................    สาขา ......................................."
   
   TempStr2 = "ข้าพเจ้าได้รับสินค้าข้างต้นจำนวนถูกต้องและสภาพเรียบร้อย" & vbCrLf & vbCrLf & vbCrLf
   TempStr2 = TempStr2 & "ผู้ขนส่ง ........................................ ผู้รับของ ......................................." & vbCrLf
   TempStr2 = TempStr2 & "     วันที่ ............/............./...........        วันที่ ............/............./..........."
   
   strFormat = VSP_CalTable("<10|<10;", mdblWid, alngX)
   strPrint = TempStr1 & "|" & TempStr2 & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)

   TempStr1 = "" & vbCrLf
   TempStr1 = TempStr1 & "ผู้รับเงิน ........................................" & vbCrLf
   TempStr1 = TempStr1 & "     วันที่ ............/............./..........."

   TempStr2 = "" & vbCrLf
   TempStr2 = TempStr2 & "ผู้ตรวจสอบ ................................. ผู้รับมอบอำนาจ ................................" & vbCrLf
   TempStr2 = TempStr2 & "            วันที่ ........./.........../..........                  วันที่ ......../.........../........"

   strFormat = VSP_CalTable("<10|<10;", mdblWid, alngX)
   strPrint = TempStr1 & "|" & TempStr2 & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
End Sub

Private Sub PrintBillFooter2(alngX() As Long)
Dim strFormat As String
Dim strPrint As String
Dim TempStr1 As String
Dim TempStr2 As String
Dim TempStr3 As String
Dim Bh As CBulkHole
Dim I As Long
Dim EndStr As String

   Vsp.TableBorder = tbNone
   Vsp.FontBold = False
   
   'Vsp = ""
   Vsp.TableBorder = tbAll

   TempStr1 = ""
   I = 0
   For Each Bh In m_cData.BulkHoles
      I = I + 1
      If I = m_cData.BulkHoles.Count Then
         EndStr = ""
      Else
         EndStr = vbCrLf
      End If
      TempStr1 = TempStr1 & "ช่องที่ " & I & " " & Bh.PART_NO & " = " & FormatNumber(Bh.HOLE_AMOUNT) & " " & Bh.NOTE & EndStr
   Next Bh
   
   TempStr2 = ""
   I = 0
   For Each Bh In m_cData.BulkHoles
      I = I + 1
      If I = m_cData.BulkHoles.Count Then
         EndStr = ""
      Else
         EndStr = vbCrLf
      End If
      TempStr2 = TempStr2 & "ลงเรือน ............. จำนวน ................." & EndStr
   Next Bh

   TempStr3 = ""
   TempStr3 = TempStr3 & "ทำลายชีลตะกั่ว        .............." & vbCrLf
   TempStr3 = TempStr3 & "ชั่ง น.น. รถเข้า         .............." & vbCrLf
   TempStr3 = TempStr3 & "รับอาหารเข้า            .............." & vbCrLf
   TempStr3 = TempStr3 & "ชั่ง น.น. รถออก       .............." & vbCrLf
   TempStr3 = TempStr3 & "ตรวจสอบรถก่อนออก .........."
   ''Debug.Print Vsp.CurrentY
   strFormat = VSP_CalTable("<20|<11|<9;", mdblWid, alngX)
   strPrint = TempStr1 & "|" & TempStr2 & "|" & TempStr3 & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   ''Debug.Print Vsp.CurrentY
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

Private Sub SetHeaderBodyAlign(F1 As String, F2 As String, F3 As String, F4 As String, F5 As String, F6 As String, F7 As String, F8 As String, F9 As String, F10 As String)
   m_Align(1) = F1
   m_Align(2) = F2
   m_Align(3) = F3
   m_Align(4) = F4
   m_Align(5) = F5
   m_Align(6) = F6
   m_Align(7) = F7
   m_Align(8) = F8
   m_Align(9) = F9
   m_Align(10) = F10
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

   Call Cf.SetFieldValue(15, m_Align(1), F1, "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(45.5, m_Align(2), F2, "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(17, m_Align(3), F3, "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(25, m_Align(4), F4, "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(15, m_Align(5), F5, "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(20, m_Align(6), F6, "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(0, m_Align(7), F7, "^")
   Call m_Headers1.AddField(Cf)
   
  Call m_Headers1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr2
   Vsp.FontBold = FontBold
'Vsp.TableBorder = tbAll
   Call Vsp.AddTable(strFormat, "", strPrint)
   Vsp.FontBold = TempFontBold
End Sub
Private Sub PrintReportBody2(F1 As String, F2 As String, F3 As String, F4 As String, F5 As String, F6 As String, F7 As String, FontBold As Boolean)
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

   Call Cf.SetFieldValue(15, m_Align(1), F1, "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(20, m_Align(2), F2, "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(17, m_Align(3), F3, "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(50, m_Align(4), F4, "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(15, m_Align(5), F5, "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(20, m_Align(6), F6, "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(0, m_Align(7), F7, "^")
   Call m_Headers1.AddField(Cf)
   
  Call m_Headers1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr2
   Vsp.FontBold = FontBold
'Vsp.TableBorder = tbAll
   Call Vsp.AddTable(strFormat, "", strPrint)
   Vsp.FontBold = TempFontBold
End Sub

Private Sub PrintReportAddress(F1 As String, F2 As String, FontBold As Boolean)
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

   Call Cf.SetFieldValue(15, m_Align(1), F1, "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(125.5, m_Align(2), F2, "^")
   Call m_Headers1.AddField(Cf)
      
  Call m_Headers1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr2
   Vsp.FontBold = FontBold
'Vsp.TableBorder = tbAll
   Call Vsp.AddTable(strFormat, "", strPrint)
   Vsp.FontBold = TempFontBold
End Sub

Private Sub PrintReportHeader(F1 As String, F2 As String, F3 As String, F4 As String, F5 As String, F6 As String, F7 As String, F8 As String, F9 As String, F10 As String, FontBold As Boolean, Border As TableBorderSettings)
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

   Call Cf.SetFieldValue(18, m_Align(1), F1, "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(40.5, m_Align(2), F2, "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(12, m_Align(3), F3, "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(16, m_Align(4), F4, "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(15, m_Align(5), F5, "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(20, m_Align(6), F6, "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(0, m_Align(7), F7, "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(20, m_Align(8), F8, "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(19, m_Align(9), F9, "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(20, m_Align(10), F10, "^")
   Call m_Headers1.AddField(Cf)
'strFormat = VSP_CalTable(DETAIL_FORMAT, mdblWid, alngX)
   Call m_Headers1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr2
   Vsp.FontBold = FontBold
'Vsp.TableBorder = tbAll
  Vsp.TableBorder = Border
   Call Vsp.AddTable(strFormat, "", strPrint)
   Vsp.FontBold = TempFontBold
   
   TEMP_FORMAT = strFormat
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

   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
      
   Vsp.FontSize = DETAIL_SIZE
   Vsp.FontBold = False
   Vsp.TableBorder = tbNone

   Call SetReportAlign("<", "<", "<", "<", "<", "<", "<")
   Call PrintReportAddress("ชื่อบริษัท", glbEnterPrise.ENTERPRISE_NAME, False)
   If m_Addresses.Count >= 1 Then
      Set m_EnpAddress = m_Addresses(1)
   Else
      Set m_EnpAddress = New CAddress
   End If
   If m_Addresses.Count >= 2 Then
      Set m_FactAddress = m_Addresses(2)
   Else
      Set m_FactAddress = New CAddress
   End If
   
   Call PrintReportAddress("สำนักงาน", m_EnpAddress.PackAddressEx, False)
   Call PrintReportAddress("โรงงาน", m_FactAddress.PackAddressEx, False)
   Call PrintReportBody("ใบสั่งซื้อ", m_cData.REF, "วันที่", DateToStringExtEx2(m_cData.DOCUMENT_DATE), "เลขที่", m_cData.DOCUMENT_NO, "", True)
   If mcolParam("notShowDlc") Then
      Call PrintReportBody2("รหัสลูกค้า", m_cData.CUSTOMER_CODE, "สถานที่จัดส่ง", "", "ใบส่งของ", m_cData.TEMP_DO_NO, "", True)
   Else
       Call PrintReportBody2("รหัสลูกค้า", m_cData.CUSTOMER_CODE, "สถานที่จัดส่ง", m_cData.PAYMENT_DESC, "ใบส่งของ", m_cData.TEMP_DO_NO, "", True)
   End If
'   Call PrintReportBody("ชื่อลูกค้า", m_cData.CUSTOMER_NAME, "ทะเบียนรถ", m_cData.NOTE, "วันที่ส่งสินค้า", IIf(DateToStringExtEx2(m_cData.SHIPMENT) <> "", DateToStringExtEx2(m_cData.SHIPMENT), DateToStringExtEx2(m_cData.DOCUMENT_DATE)), "", True)
    Call PrintReportBody("ชื่อลูกค้า", m_cData.CUSTOMER_NAME, "ทะเบียนรถ", m_cData.NOTE, "วันที่ส่งสินค้า", DateToStringExtEx2(m_cData.DOCUMENT_DATE), "", True)
   Call PrintReportAddress("ที่อยู่ลูกค้า", m_CustAddr.PackAddressEx, True)
   Call PrintReportBody("พนักงานขาย", m_cData.EMP_CODE & " " & m_cData.EMP_NAME & " " & m_cData.EMP_LNAME, "กำหนดชำระ", m_cData.Credit & " วัน", "วันนัดชำระ", DateToStringExtEx2(m_cData.DUE_DATE), "", True)

   Call SetHeaderBodyAlign("^", "^", "^", "^", "^", "^", "^", "^", "^", "^")
   Call PrintReportHeader("เบอร์สินค้า", "รายการ", "กก./ถุง", "จำนวนถุง", "น.น.รวม" & vbNewLine & "(กก.)", "ราคา/หน่วย" & vbNewLine & "(บาท)", "จำนวนเงิน" & vbNewLine & "(บาท)", "ส่วนลด/ถุง" & vbNewLine & "(บาท)", "รวมส่วนลด" & vbNewLine & "(บาท)", "ยอดรวมสุทธิ" & vbNewLine & "(บาท)", True, tbAll)
'   Call PrintReportHeader("รหัส", "รายการ", "ขนาด", "จำนวน", "น.น.", "ราคา/หน่วย", "จำนวนเงิน (บาท)", "ส่วนลด/ถุง", "ส่วนลด", "สุทธิ", True, tbAll)
   
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
   
   If mcolParam("DOCUMENT_TYPE") = 1 Then
      MODULE_DESC = MapText("ใบส่งสินค้า/ใบแจ้งหนี้")
   ElseIf mcolParam("DOCUMENT_TYPE") = 2 Then
      MODULE_DESC = MapText("ใบส่งสินค้า/ใบเสร็จรับเงิน")
   End If
   
   
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

Private Function CountBulkHole() As Long
Dim Bh As CBulkHole
Dim TempCount As Long

   TempCount = 0
   For Each Bh In m_cData.BulkHoles
      If Bh.PART_ITEM_ID > 0 Then
         TempCount = TempCount + 1
      End If
   Next Bh
   
   CountBulkHole = TempCount
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
Dim Bi As CDoItem
Dim TempBi As CDoItem
Dim Bi2 As Collection
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
Dim Total1(20) As Double
Dim J As Long
Dim MaxRow As Long
Dim Bsc As CBillingDiscount
Dim TimeInStr As String
Dim TimeOutStr As String
Dim tempOffset As Long

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
   
   If (Rs Is Nothing) Or (iCount < 1) Then
      mstrErrMsg = "ไม่พบข้อมูลที่ต้องการในฐานข้อมูล"
      Exit Function
   End If
   
   Call m_cData.PopulateFromRS(1, Rs)
   
   If m_cData.BILLING_ADDRESS_ID > 0 Then
      m_CustAddr.ADDRESS_ID = m_cData.BILLING_ADDRESS_ID
      Call m_CustAddr.QueryData(Rs, iCount)
      Call m_CustAddr.PopulateFromRS(Rs)
   End If
   
   m_EnpAddress.ENTERPRISE_ID = glbEnterPrise.ENTERPRISE_ID
   Call m_EnpAddress.QueryData2(Rs, iCount)
   Set m_EnpAddress = Nothing
   While Not Rs.EOF
      Set m_EnpAddress = New CAddress
      Call m_EnpAddress.PopulateFromRS(Rs)
      Call m_Addresses.add(m_EnpAddress)
      Set m_EnpAddress = Nothing
      Rs.MoveNext
   Wend
   
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
   Sum1 = 0
   Sum2 = 0
   Sum3 = 0
   
   For I = 1 To UBound(Total1)
      Total1(I) = 0
   Next I
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
      Call SetHeaderBodyAlign("^", "<", "^", ">", ">", ">", ">", ">", ">", ">")
      J = 0
      
      For Each Bi In m_cData.DoItems
         J = J + 1
         If (Bi.FEATURE_ID > 0) And (Bi.SERVICE_FLAG = "Y") Then
            Bi.PACK_AMOUNT = 0
            Bi.ITEM_AMOUNT = 0
         End If

         For Each TempBi In m_cData.DoItems
            If TempBi.SUM_WITH_DO_ITEM_ID = Bi.DO_ITEM_ID Then
               Bi.PACK_AMOUNT = Bi.PACK_AMOUNT + TempBi.PACK_AMOUNT
               Bi.ITEM_AMOUNT = Bi.ITEM_AMOUNT + TempBi.ITEM_AMOUNT
               Bi.DISCOUNT_AMOUNT = Bi.PACK_AMOUNT * Bi.DISCOUNT_PER_PACK
'               Bi.DISCOUNT_AMOUNT = Bi.DISCOUNT_AMOUNT + TempBi.DISCOUNT_AMOUNT
'               Bi.DISCOUNT_PER_PACK = MyDiff(Bi.DISCOUNT_AMOUNT, Bi.PACK_AMOUNT)
               Bi.TOTAL_PRICE = Bi.PACK_AMOUNT * (Bi.PRICE_PER_PACK - Bi.DISCOUNT_PER_PACK)
            End If
         Next TempBi

         If Not Val(Bi.SUM_WITH_DO_ITEM_ID) > -1 Then
              If mcolParam("PrintNotPrice") Then
                  Bi.PRICE_PER_PACK = 0
                  Bi.DISCOUNT_PER_PACK = 0
                  Bi.DISCOUNT_AMOUNT = 0
                  Bi.TOTAL_PRICE = 0
             End If
   
            Call PrintReportHeader(Bi.ShowCodeText, Bi.ShowDescText, FormatNumber(Bi.WEIGHT_PER_PACK), FormatNumber(Bi.PACK_AMOUNT), FormatNumber(Bi.ITEM_AMOUNT), FormatNumber(Bi.PRICE_PER_PACK, 3), 0, FormatNumber(Bi.DISCOUNT_PER_PACK), FormatNumber(Bi.DISCOUNT_AMOUNT), FormatNumber(Bi.TOTAL_PRICE), False, tbAll)
            Total1(4) = Total1(4) + Bi.PACK_AMOUNT
            Total1(5) = Total1(5) + Bi.ITEM_AMOUNT
            Total1(9) = Total1(9) + Bi.DISCOUNT_AMOUNT
            Total1(10) = Total1(10) + Bi.TOTAL_PRICE
         End If
      Next Bi
      
      
      
'      tempOffset = 5000
'      tempOffset = tempOffset + (m_cData.BillingDiscounts.Count * 386.4)
'      If CountBulkHole > 0 Then
'         tempOffset = tempOffset + 1932
'      End If
      tempOffset = 5600
      tempOffset = tempOffset + (m_cData.BillingDiscounts.Count * 386.4)
      If CountBulkHole > 0 Then
         tempOffset = tempOffset + 1932
      End If
      
      Call PatchTable(Vsp, TEMP_FORMAT, "||||||||;", tempOffset)
      
      'Call PrintReportHeader("", "", "", "", "", "", "", "", "", "", False, tbAll)
      
      For Each Bsc In m_cData.BillingDiscounts
         J = J + 1
         Call PrintReportHeader("", Bsc.DISCOUNT_NAME, "", "", "", "", "", "", "", FormatNumber(Bsc.DISCOUNT_AMOUNT), False, tbAll)
         Total1(10) = Total1(10) - Bsc.DISCOUNT_AMOUNT
      Next Bsc
      Call PrintReportHeader("รวม", "", "", FormatNumber(Total1(4)), FormatNumber(Total1(5)), "", "", "", FormatNumber(Total1(9)), FormatNumber(Total1(10)), True, tbAll)

      Call PrintBillFooter1(alngX)
      If CountBulkHole > 0 Then
         Call PrintBillFooter2(alngX)
      End If
   
   TimeInStr = Format(HOUR(m_cData.ENTRY_DATE), "00") & ":" & Format(Minute(m_cData.ENTRY_DATE), "00")
   TimeOutStr = Format(HOUR(m_cData.EXIT_DATE), "00") & ":" & Format(Minute(m_cData.EXIT_DATE), "00")
         
   Vsp.TableBorder = tbNone
   strFormat = VSP_CalTable("<10|^10|^10;", mdblWid, alngX)
   strPrint = "FM-PN-35 Rev. 00|เวลาเข้า " & TimeInStr & "|เวลาออก " & TimeOutStr & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Vsp.TableBorder = tbNone
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = "      ในกรณีที่ผู้ซื้อผิดนัดการชำระค่าสินค้าตาม ใบส่งสินค้า/ใบแจ้งหนี้ ที่กำหนดไว้ บริษัทฯ มีสิทธิ์คิดดอกเบี้ยกับผู้ซื้อในอัตรา 15% ต่อปี;"
   Call Vsp.AddTable(strFormat, "", strPrint)
      
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

   If mcolParam("ExampleDoc") Then
      Vsp.DrawPicture mcolParam("BACK_GROUND2"), "0.05cm", "0.0cm", "21.55cm", "15.5cm", 10
   End If
   
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

