VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAP021"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAP021"
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
Private m_Suppliers As Collection
Private m_SuppliersAcc As Collection
Private m_PartItems As Collection
Private m_CustomerTrans As Collection
Private m_DriverTrans As Collection
Private m_SaleTrans As Collection
Private m_FacTrans As Collection
Private m_PriceTrans As Collection
Private m_TruckTrans As Collection
Private m_DocDetail As Collection
Private Sp As CSupplierTranSport
Private SA As CSupplierAccount
Private DocFlag As Long
Dim strTruck As String
Dim NetTotal As Double
Private AddTrnDriver As Double
Private ReduceTrnDriver  As Double
Private TAX As Double

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
   Set m_Suppliers = New Collection
   Set m_SuppliersAcc = New Collection
   Set m_PartItems = New Collection
   Set m_CustomerTrans = New Collection
   Set m_DriverTrans = New Collection
   Set m_SaleTrans = New Collection
   Set m_FacTrans = New Collection
   Set m_PriceTrans = New Collection
   Set m_TruckTrans = New Collection
   Set m_DocDetail = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_Suppliers = Nothing
   Set m_SuppliersAcc = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
   Set m_PartItems = Nothing
   Set m_CustomerTrans = Nothing
   Set m_DriverTrans = Nothing
   Set m_SaleTrans = Nothing
   Set m_FacTrans = Nothing
   Set m_PriceTrans = Nothing
   Set m_TruckTrans = Nothing
   Set m_DocDetail = Nothing
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

   Call Cf.SetFieldValue(12, "^", "วันที่", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(13, "^", "เลขที่", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(27, "^", "สถานที่", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(10, "^", "เลข", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "จ.น.", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(27, "^", "จำนวน(ถุง)", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(16, "^", "ราคา/หน่วย", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(20, "^", "ราคารวม", ">")
   Call m_Headers0.AddField(Cf)
   
   If mcolParam("CUSTOMER_ONLY_FLAG") = 1 Then
      If mcolParam("COLUMN_DEDUC_FLAG") = 1 Then
         Call Cf.SetFieldValue(14, "^", "หักค่าขนส่ง", ">")
         Call m_Headers0.AddField(Cf)
      Else
         Call Cf.SetFieldValue(0, "^", "หักค่าขนส่ง", ">")
         Call m_Headers0.AddField(Cf)
      End If
      If mcolParam("ACCOUNT_ADD_FLAG") = 1 Then
         Call Cf.SetFieldValue(14, "^", "หักค่าขนส่ง", ">")
         Call m_Headers0.AddField(Cf)
       Else
         Call Cf.SetFieldValue(0, "^", "หักค่าขนส่ง", ">")
         Call m_Headers0.AddField(Cf)
      End If
      Call Cf.SetFieldValue(13, "^", "หมาย", "<")
      Call m_Headers0.AddField(Cf)
   Else
      Call Cf.SetFieldValue(8, "^", "กำไร", ">")
      Call m_Headers0.AddField(Cf)
      Call Cf.SetFieldValue(11, "^", "ขนส่งเก็บ", ">")
      Call m_Headers0.AddField(Cf)
      Call Cf.SetFieldValue(11, "^", "ขนส่งเก็บ", ">")
      Call m_Headers0.AddField(Cf)
      If mcolParam("COLUMN_DEDUC_FLAG") = 1 Then
         Call Cf.SetFieldValue(18, "^", "หักค่าขนส่ง", ">")
         Call m_Headers0.AddField(Cf)
      Else
          Call Cf.SetFieldValue(0, "^", "หักค่าขนส่ง", ">")
         Call m_Headers0.AddField(Cf)
      End If
      If mcolParam("ACCOUNT_ADD_FLAG") = 1 Then
         Call Cf.SetFieldValue(18, "^", "เพิ่มค่าขนส่ง", ">")
         Call m_Headers0.AddField(Cf)
      Else
        Call Cf.SetFieldValue(0, "^", "เพิ่มค่าขนส่ง", ">")
         Call m_Headers0.AddField(Cf)
      End If
      Call Cf.SetFieldValue(14, "^", "หมาย", "<")
      Call m_Headers0.AddField(Cf)
   End If

   Call Cf.SetFieldValue(12, "^", "", "^") '1
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(13, "^", "เอกสาร", "^") '2
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(27, "^", "จัดส่ง", "<") '3
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(10, "^", "ทะเบียน", "^") '4
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "รอบ", "^") '5
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(5, "^", "10", ">") '6
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(5, "^", "20", ">") '7
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(5, "^", "30", ">") '8
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(5, "^", "50", ">") '9
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(7, "^", "B", ">") '10
   Call m_Headers1.AddField(Cf)
   If mcolParam("CUSTOMER_ONLY_FLAG") = 1 Then
      Call Cf.SetFieldValue(16, "^", "รับจ้าง", ">") '11
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(20, "^", "รับจ้าง", ">") '12
      Call m_Headers1.AddField(Cf)
      If mcolParam("COLUMN_DEDUC_FLAG") = 1 Then
         Call Cf.SetFieldValue(7, "^", "รับจ้าง", ">") '18
         Call m_Headers1.AddField(Cf)
         Call Cf.SetFieldValue(7, "^", "ลูกค้า", ">") '18
         Call m_Headers1.AddField(Cf)
      Else
         Call Cf.SetFieldValue(0, "^", "รับจ้าง", ">") '18
         Call m_Headers1.AddField(Cf)
         Call Cf.SetFieldValue(0, "^", "ลูกค้า", ">") '18
         Call m_Headers1.AddField(Cf)
      End If
      If mcolParam("ACCOUNT_ADD_FLAG") = 1 Then
         Call Cf.SetFieldValue(7, "^", "รับจ้าง", ">") '18
         Call m_Headers1.AddField(Cf)
         Call Cf.SetFieldValue(7, "^", "ลูกค้า", ">") '18
         Call m_Headers1.AddField(Cf)
      Else
         Call Cf.SetFieldValue(0, "^", "รับจ้าง", ">") '18
         Call m_Headers1.AddField(Cf)
         Call Cf.SetFieldValue(0, "^", "ลูกค้า", ">") '18
         Call m_Headers1.AddField(Cf)
      End If
      Call Cf.SetFieldValue(13, "^", "เหตุ", "<") '13
      Call m_Headers1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(8, "^", "รับจ้าง", ">") '11
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(8, "^", "ลูกค้า", ">") '12
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(10, "^", "รับจ้าง", ">") '13
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(10, "^", "ลูกค้า", ">") '14
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(8, "^", "ขนส่ง", ">") '15
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(11, "^", "การตลาด", ">") '16
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(11, "^", "โรงงาน", ">") '17
      Call m_Headers1.AddField(Cf)
      If mcolParam("COLUMN_DEDUC_FLAG") = 1 Then
         Call Cf.SetFieldValue(9, "^", "รับจ้าง", ">") '18
         Call m_Headers1.AddField(Cf)
         Call Cf.SetFieldValue(9, "^", "ลูกค้า", ">") '18
         Call m_Headers1.AddField(Cf)
      Else
         Call Cf.SetFieldValue(0, "^", "รับจ้าง", ">") '18
         Call m_Headers1.AddField(Cf)
         Call Cf.SetFieldValue(0, "^", "ลูกค้า", ">") '18
         Call m_Headers1.AddField(Cf)
      End If
      If mcolParam("ACCOUNT_ADD_FLAG") = 1 Then
         Call Cf.SetFieldValue(9, "^", "รับจ้าง", ">") '18
         Call m_Headers1.AddField(Cf)
         Call Cf.SetFieldValue(9, "^", "ลูกค้า", ">") '18
         Call m_Headers1.AddField(Cf)
      Else
         Call Cf.SetFieldValue(0, "^", "รับจ้าง", ">") '18
         Call m_Headers1.AddField(Cf)
         Call Cf.SetFieldValue(0, "^", "ลูกค้า", ">") '18
         Call m_Headers1.AddField(Cf)
      End If
      Call Cf.SetFieldValue(14, "^", "เหตุ", "<") '19
      Call m_Headers1.AddField(Cf)
   End If
 
   Set Cf = Nothing
        
    If mcolParam("INVOICE_FLAG") = 1 Then
       Call m_Headers0.GetString(1, TempStr1, TempStr1_1)
       strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
       strPrint = TempStr1_1
       Call Vsp.AddTable(strFormat, "", strPrint)
       
      Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
       strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
       strPrint = TempStr1_1
       Call Vsp.AddTable(strFormat, "", strPrint)
   End If
            
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
   
   TAX = 0.01 'ภาษีที่หัก
   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
   Vsp.ORIENTATION = orLandscape
   Vsp.MarginBottom = 1440
   Vsp.MarginFooter = 1440
   Vsp.MarginHeader = 1440
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 1440
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
   ToDate = "ถึง " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")
   
   MODULE_DESC = vbCrLf & glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & vbCrLf & _
                                       "ใบวางบิล" & vbCrLf & mcolParam("REPORT_NAME") & "ประจำงวด" & IntToThaiMonth(Month(mcolParam("FROM_DATE"))) & Year(mcolParam("FROM_DATE")) + 543 & vbCrLf & _
                                        FromDate & " " & ToDate
                                        
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
Dim R As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Amt As Double
Dim BT As CBillTransport
Dim TempDI As CBillTransport
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
Dim Cm As CCustomer
Dim Pi As CPartItem
Dim TempCal As Double
Dim TempCal2 As Double
Dim FIX_RATE_DETAIL As String
Dim NOT_CAL_VAT As String
Dim IsMain As Boolean


   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------

   Set Rs = New ADODB.Recordset

  Call LoadSupplierTransport(Nothing, m_Suppliers, Trim(mcolParam("TRUCK_NO")), Trim(mcolParam("SUPPLIER_CODE_TRANSPORT")))
  If mcolParam("TRUCK_NO") <> "" Then
      Set Sp = GetObject("", m_Suppliers, Trim(mcolParam("TRUCK_NO")))
  ElseIf mcolParam("SUPPLIER_CODE_TRANSPORT") <> "" Then
      Set Sp = GetObject("", m_Suppliers, Trim(mcolParam("SUPPLIER_CODE_TRANSPORT")))
  Else
      Set Sp = GetObject("", m_Suppliers, Trim(mcolParam("SUPPLIER_CODE_TRANSPORT")))
  End If
  If Sp Is Nothing Then
      Set Sp = New CSupplierTranSport
  End If
  
  Call LoadSupplierAccount(Nothing, m_SuppliersAcc, , Trim(Sp.SUPPLIER_CODE))
  Set SA = GetObject("", m_SuppliersAcc, Trim(Sp.SUPPLIER_CODE))
  If SA Is Nothing Then
      Set SA = New CSupplierAccount
  End If
   
  Call LoadBillTransportByTypePrice(Nothing, m_PriceTrans, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), 1) 'ค่าขน KEY BY PRICE
  Call LoadBillTransportByTypePrice(Nothing, m_TruckTrans, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), 2) 'ค่าขน KEY BY WEIGHT
  
  ' ดึงข้อมูลทั้งหมดที่อยู่ในบิลค่าขนส่ง
     Dim m_BillDetail As CBillTransport
     Dim m_BillDetail2 As CBillTransport
      Set BT = New CBillTransport
      BT.BILL_TRANSPORT_ID = -1
      BT.SUPPLIER_TRANSPORT_CODE = mcolParam("TRUCK_NO")
      BT.SUPPLIER_CODE = mcolParam("SUPPLIER_CODE_TRANSPORT")
      BT.CUSTOMER_CODE = mcolParam("CUSTOMER_CODE")
      BT.DOCUMENT_NO = mcolParam("DOCUMENT_NO")
      BT.FROM_DATE = mcolParam("FROM_DATE")
      BT.TO_DATE = mcolParam("TO_DATE")
      BT.OrderBy = mcolParam("ORDER_BY")
      BT.OrderType = mcolParam("ORDER_TYPE")
      I = 0

     Call BT.QueryData(2, Rs, iCount)
     While Not Rs.EOF
         Call BT.PopulateFromRS(2, Rs)
         
         I = I + 1
         
         Set m_BillDetail = GetObject("", m_DocDetail, Trim(BT.DOCUMENT_NO) & "-" & Trim(str(BT.WEIGHT_PER_UNIT)))
         If (m_BillDetail Is Nothing) Or Val(BT.KEY_CODE) > 2 Then
            Set m_BillDetail = New CBillTransport
         End If
         
            m_BillDetail.DOCUMENT_NO = BT.DOCUMENT_NO
            m_BillDetail.DOCUMENT_DATE = BT.DOCUMENT_DATE
            m_BillDetail.SHIPPING = BT.SHIPPING
            m_BillDetail.SHIPPING_DATE = BT.SHIPPING_DATE
            m_BillDetail.SUPPLIER_TRANSPORT_CODE = BT.SUPPLIER_TRANSPORT_CODE
            m_BillDetail.AROUND = IIf(BT.AROUND > 0, BT.AROUND, 0)
            m_BillDetail.NOTE = BT.NOTE
            m_BillDetail.NOTE_ITEM = BT.NOTE_ITEM
            m_BillDetail.KEY_CODE = BT.KEY_CODE
            m_BillDetail.CUSTOMER_CODE = BT.CUSTOMER_CODE
            m_BillDetail.CUSTOMER_NAME = BT.CUSTOMER_NAME
            m_BillDetail.WEIGHT_PER_UNIT = BT.WEIGHT_PER_UNIT
            m_BillDetail.CAL_DRIVER = BT.CAL_DRIVER
            m_BillDetail.CAL_CUSTOMER = BT.CAL_CUSTOMER
            m_BillDetail.CAL_PRICE_IN_PRODUCT = BT.CAL_PRICE_IN_PRODUCT
            m_BillDetail.FIX_RATE_DETAIL = BT.FIX_RATE_DETAIL
            FIX_RATE_DETAIL = BT.FIX_RATE_DETAIL
            m_BillDetail.NOT_CAL_VAT = BT.NOT_CAL_VAT
            
            Set TempDI = GetObject("", m_TruckTrans, Trim(str(BT.BILLING_DOC_ID)) & "-" & Trim(str(BT.BILL_TRANSPORT_ITEM_ID)) & "-" & Trim("10"))
            If Not TempDI Is Nothing Then
                  m_BillDetail.C10 = TempDI.PACK_AMOUNT
            End If
            Set TempDI = GetObject("", m_TruckTrans, Trim(str(BT.BILLING_DOC_ID)) & "-" & Trim(str(BT.BILL_TRANSPORT_ITEM_ID)) & "-" & Trim("20"))
            If Not TempDI Is Nothing Then
                  m_BillDetail.C20 = TempDI.PACK_AMOUNT
            End If
            Set TempDI = GetObject("", m_TruckTrans, Trim(str(BT.BILLING_DOC_ID)) & "-" & Trim(str(BT.BILL_TRANSPORT_ITEM_ID)) & "-" & Trim("30"))
            If Not TempDI Is Nothing Then
                  m_BillDetail.C30 = TempDI.PACK_AMOUNT
            End If
            Set TempDI = GetObject("", m_TruckTrans, Trim(str(BT.BILLING_DOC_ID)) & "-" & Trim(str(BT.BILL_TRANSPORT_ITEM_ID)) & "-" & Trim("50"))
            If Not TempDI Is Nothing Then
                  m_BillDetail.C50 = TempDI.PACK_AMOUNT
            End If
            Set TempDI = GetObject("", m_TruckTrans, Trim(str(BT.BILLING_DOC_ID)) & "-" & Trim(str(BT.BILL_TRANSPORT_ITEM_ID)) & "-" & Trim("1"))
            If Not TempDI Is Nothing Then
                  m_BillDetail.Cb = TempDI.PACK_AMOUNT
            End If
            
            Set TempDI = GetObject("", m_PriceTrans, Trim(str(BT.BILLING_DOC_ID)) & "-" & Trim(str(BT.BILL_TRANSPORT_ITEM_ID)) & "-" & Trim("01"))
            If Not TempDI Is Nothing Then
                  m_BillDetail.RATE_DRIVER = TempDI.RATE_PER_UNIT
                  m_BillDetail.PRICE_DRIVER = TempDI.TOTAL_PRICE
            End If
            Set TempDI = GetObject("", m_PriceTrans, Trim(str(BT.BILLING_DOC_ID)) & "-" & Trim(str(BT.BILL_TRANSPORT_ITEM_ID)) & "-" & Trim("02"))
            If Not TempDI Is Nothing Then
               m_BillDetail.RATE_CUSTOMER = TempDI.RATE_PER_UNIT
               If BT.CAL_PRICE_IN_PRODUCT = "Y" Then
                   m_BillDetail.PRICE_SALE = TempDI.TOTAL_PRICE
               Else
                  m_BillDetail.PRICE_CUSTOMER = TempDI.TOTAL_PRICE
               End If
            End If
           Set TempDI = GetObject("", m_PriceTrans, Trim(str(BT.BILLING_DOC_ID)) & "-" & Trim(str(BT.BILL_TRANSPORT_ITEM_ID)) & "-" & Trim("03"))
            If Not TempDI Is Nothing Then
                   m_BillDetail.RATE_DRIVER = TempDI.RATE_PER_UNIT
                  m_BillDetail.PRICE_DRIVER = TempDI.TOTAL_PRICE
                  
                  m_BillDetail.RATE_SALE = TempDI.RATE_PER_UNIT
                  m_BillDetail.PRICE_SALE = TempDI.TOTAL_PRICE
            End If
           Set TempDI = GetObject("", m_PriceTrans, Trim(str(BT.BILLING_DOC_ID)) & "-" & Trim(str(BT.BILL_TRANSPORT_ITEM_ID)) & "-" & Trim("04"))
            If Not TempDI Is Nothing Then
                   m_BillDetail.RATE_DRIVER = TempDI.RATE_PER_UNIT
                  m_BillDetail.PRICE_DRIVER = TempDI.TOTAL_PRICE
                  
                  m_BillDetail.RATE_FAC = TempDI.RATE_PER_UNIT
                  m_BillDetail.PRICE_FAC = TempDI.TOTAL_PRICE
            End If
           Set TempDI = GetObject("", m_PriceTrans, Trim(str(BT.BILLING_DOC_ID)) & "-" & Trim(str(BT.BILL_TRANSPORT_ITEM_ID)) & "-" & Trim("05")) 'รายได้
            If Not TempDI Is Nothing Then
                 If BT.CAL_DRIVER = "Y" Then
                     m_BillDetail.RATE_DRIVER = TempDI.RATE_PER_UNIT
                     m_BillDetail.PRICE_DRIVER = TempDI.TOTAL_PRICE * -1
                 ElseIf BT.CAL_CUSTOMER = "Y" Then
                     m_BillDetail.RATE_CUSTOMER = TempDI.RATE_PER_UNIT
                     m_BillDetail.PRICE_CUSTOMER = TempDI.TOTAL_PRICE * -1
                 End If
                  m_BillDetail.PRICE_IMCOME_OTHER = TempDI.TOTAL_PRICE
            End If
            Set TempDI = GetObject("", m_PriceTrans, Trim(str(BT.BILLING_DOC_ID)) & "-" & Trim(str(BT.BILL_TRANSPORT_ITEM_ID)) & "-" & Trim("06")) 'รายจ่าย
            If Not TempDI Is Nothing Then
                 If BT.CAL_DRIVER = "Y" Then
                     m_BillDetail.RATE_DRIVER = TempDI.RATE_PER_UNIT
                     m_BillDetail.PRICE_DRIVER = TempDI.TOTAL_PRICE
                 ElseIf BT.CAL_CUSTOMER = "Y" Then
                     m_BillDetail.RATE_CUSTOMER = TempDI.RATE_PER_UNIT
                     m_BillDetail.PRICE_CUSTOMER = TempDI.TOTAL_PRICE
                 End If
                 m_BillDetail.PRICE_EXPEND_OTHER = TempDI.TOTAL_PRICE
            End If
            
            Set TempDI = GetObject("", m_PriceTrans, Trim(str(BT.BILLING_DOC_ID)) & "-" & Trim(str(BT.BILL_TRANSPORT_ITEM_ID)) & "-" & Trim("07")) 'รายละเอียดอื่นๆ
            If Not TempDI Is Nothing Then
                 m_BillDetail.NOTE_ITEM = TempDI.NOTE_ITEM
            End If
            
            Set TempDI = GetObject("", m_PriceTrans, Trim(str(BT.BILLING_DOC_ID)) & "-" & Trim(str(BT.BILL_TRANSPORT_ITEM_ID)) & "-" & Trim("08")) 'ส่วนลดลูกค้าคิดการตลาด
            If Not TempDI Is Nothing Then
                  m_BillDetail.RATE_CUSTOMER = TempDI.RATE_PER_UNIT
                  m_BillDetail.PRICE_CUSTOMER = TempDI.TOTAL_PRICE
                     
                  m_BillDetail.RATE_SALE = TempDI.RATE_PER_UNIT
                  m_BillDetail.PRICE_SALE = TempDI.TOTAL_PRICE
            End If
            
            Set TempDI = GetObject("", m_PriceTrans, Trim(str(BT.BILLING_DOC_ID)) & "-" & Trim(str(BT.BILL_TRANSPORT_ITEM_ID)) & "-" & Trim("09")) 'รายได้เรียกคืนค่าขนส่ง
            If Not TempDI Is Nothing Then
                 If BT.CAL_DRIVER = "Y" Then
                     m_BillDetail.PRICE_DRIVER = TempDI.TOTAL_PRICE * -1
                     m_BillDetail.PRICE_CUSTOMER = 0
                 ElseIf BT.CAL_CUSTOMER = "Y" Then
                     m_BillDetail.PRICE_DRIVER = 0
                     m_BillDetail.PRICE_CUSTOMER = TempDI.TOTAL_PRICE
                 End If
            End If
            
            Set TempDI = GetObject("", m_PriceTrans, Trim(str(BT.BILLING_DOC_ID)) & "-" & Trim(str(BT.BILL_TRANSPORT_ITEM_ID)) & "-" & Trim("10")) 'รายจ่ายเพิ่มเติมค่าขนส่ง
            If Not TempDI Is Nothing Then
                 If BT.CAL_DRIVER = "Y" Then
                     m_BillDetail.PRICE_DRIVER = TempDI.TOTAL_PRICE
                 ElseIf BT.CAL_CUSTOMER = "Y" Then
                     m_BillDetail.PRICE_CUSTOMER = TempDI.TOTAL_PRICE * -1
                 End If
            End If
            
            If (PrevKey1 <> Trim(m_BillDetail.DOCUMENT_NO) & "-" & Trim(str(m_BillDetail.WEIGHT_PER_UNIT))) Or (BT.KEY_CODE = "03" Or BT.KEY_CODE = "04" Or BT.KEY_CODE = "05" Or BT.KEY_CODE = "06" Or BT.KEY_CODE = "07" Or BT.KEY_CODE = "08" Or BT.KEY_CODE = "09" Or BT.KEY_CODE = "10") Then
               If Val(BT.KEY_CODE) > 2 Then
                  Call m_DocDetail.add(m_BillDetail, Trim(m_BillDetail.DOCUMENT_NO) & "-" & Trim(BT.BILL_TRANSPORT_ITEM_ID))
                  PrevKey1 = BT.DOCUMENT_NO
               Else
                  Set m_BillDetail2 = GetObject("", m_DocDetail, Trim(m_BillDetail.DOCUMENT_NO) & "-" & Trim(str(m_BillDetail.WEIGHT_PER_UNIT)))
                  If (m_BillDetail2 Is Nothing) Then
                     Call m_DocDetail.add(m_BillDetail, Trim(m_BillDetail.DOCUMENT_NO) & "-" & Trim(str(m_BillDetail.WEIGHT_PER_UNIT)))
                     PrevKey1 = Trim(m_BillDetail.DOCUMENT_NO) & "-" & Trim(str(m_BillDetail.WEIGHT_PER_UNIT))
                  End If
               End If
            End If
         Rs.MoveNext
      Wend
      
    Vsp.Preview = isPreview
    
   mblnHeader = False
   Call StartExportFile(Vsp)
   Call printHeader
   mblnHeader = True
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------

   For J = 1 To UBound(Total1)
      Total1(J) = 0
      Total2(J) = 0
      Total3(J) = 0
   Next J
   
      I = 0
      Set BodyCf = New CReportField
   
   If mcolParam("INVOICE_FLAG") = 1 Then
      If Not initDoc Then Exit Function

      DocFlag = 1
      Vsp.StartDoc
       I = 0
       R = 0
       PrevKey1 = ""
     For Each m_BillDetail In m_DocDetail
            I = 0
            R = R + 1
               Call m_Details1.ClearField

            If PrevKey1 <> m_BillDetail.DOCUMENT_NO Then
                IsMain = True 'แจ้งว่า เป็นบรรทัดหลักที่แสดง
            
               PrevKey1 = m_BillDetail.DOCUMENT_NO
               
                I = I + 1
               Set HeadCf = m_Headers1.Fields(I)
               TempStr = DateToStringExtEx2(m_BillDetail.DOCUMENT_DATE)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
      
               I = I + 1
               Set HeadCf = m_Headers1.Fields(I)
               If m_BillDetail.NOT_CAL_VAT = "Y" Then 'ถ้า bill ใบนี้คิดภาษีให้เข้าคำนวณ
                  TempStr = "*" & m_BillDetail.DOCUMENT_NO
               Else
                  TempStr = m_BillDetail.DOCUMENT_NO
               End If
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               I = I + 1
               Set HeadCf = m_Headers1.Fields(I)
               If Trim(m_BillDetail.SHIPPING) = "" Or Trim(m_BillDetail.SHIPPING) = "-" Then
                  If mcolParam("SHOW_CUS_NAME_FLAG") = 1 Then
                     TempStr = m_BillDetail.SHIPPING & " (" & m_BillDetail.CUSTOMER_CODE & ") " & " " & m_BillDetail.CUSTOMER_NAME
                  Else
                     TempStr = m_BillDetail.SHIPPING & " (" & m_BillDetail.CUSTOMER_CODE & ") "
                  End If
               Else
                  If mcolParam("SHOW_CUS_NAME_FLAG") = 1 Then
                     TempStr = m_BillDetail.SHIPPING & "[" & "(" & m_BillDetail.CUSTOMER_CODE & ") " & m_BillDetail.CUSTOMER_NAME & "]"
                  Else
                     TempStr = m_BillDetail.SHIPPING & "[" & m_BillDetail.CUSTOMER_CODE & "]"
                  End If
               End If
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               I = I + 1
               Set HeadCf = m_Headers1.Fields(I)
               TempStr = m_BillDetail.SUPPLIER_TRANSPORT_CODE
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               I = I + 1
               Set HeadCf = m_Headers1.Fields(I)
               TempStr = IIf(m_BillDetail.AROUND > 0, m_BillDetail.AROUND, "")
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Total1(I) = Total1(I) + m_BillDetail.AROUND
               Total2(I) = Total2(I) + m_BillDetail.AROUND
            Else
            
               IsMain = False 'แจ้งว่า เป็นบรรทัดหลักที่ไม่แสดง
                 I = I + 1
               Set HeadCf = m_Headers1.Fields(I)
               TempStr = ""
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
      
               I = I + 1
               Set HeadCf = m_Headers1.Fields(I)
               TempStr = ""
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               I = I + 1
               Set HeadCf = m_Headers1.Fields(I)
               TempStr = "-"
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               I = I + 1
               Set HeadCf = m_Headers1.Fields(I)
               If Not Val(m_BillDetail.KEY_CODE) > 2 Then
                  TempStr = m_BillDetail.SUPPLIER_TRANSPORT_CODE
               Else
                  TempStr = ""
               End If
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               I = I + 1
               Set HeadCf = m_Headers1.Fields(I)
               TempStr = ""
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
            End If
      
                I = I + 1
                Set HeadCf = m_Headers1.Fields(I)
               TempStr = IIf(m_BillDetail.C10 = 0, "", FormatNumber(m_BillDetail.C10, 0))
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               If IsMain Then
               Total1(I) = Total1(I) + m_BillDetail.C10
               Total2(I) = Total2(I) + m_BillDetail.C10
               End If
               
                I = I + 1
                Set HeadCf = m_Headers1.Fields(I)
               TempStr = IIf(m_BillDetail.C20 = 0, "", FormatNumber(m_BillDetail.C20, 0))
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               If IsMain Then
               Total1(I) = Total1(I) + m_BillDetail.C20
               Total2(I) = Total2(I) + m_BillDetail.C20
               End If
               
                I = I + 1
                Set HeadCf = m_Headers1.Fields(I)
                TempStr = IIf(m_BillDetail.C30 = 0, "", FormatNumber(m_BillDetail.C30, 0))
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               If IsMain Then
               Total1(I) = Total1(I) + m_BillDetail.C30
               Total2(I) = Total2(I) + m_BillDetail.C30
               End If
               
                I = I + 1
                Set HeadCf = m_Headers1.Fields(I)
               TempStr = IIf(m_BillDetail.C50 = 0, "", FormatNumber(m_BillDetail.C50, 0))
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               If IsMain Then
               Total1(I) = Total1(I) + m_BillDetail.C50
               Total2(I) = Total2(I) + m_BillDetail.C50
               End If
               
                I = I + 1
                Set HeadCf = m_Headers1.Fields(I)
               TempStr = IIf(m_BillDetail.Cb = 0, "", FormatNumber(m_BillDetail.Cb, 2))
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               If IsMain Then
               Total1(I) = Total1(I) + m_BillDetail.Cb
               Total2(I) = Total2(I) + m_BillDetail.Cb
               End If
      
                I = I + 1
                Set HeadCf = m_Headers1.Fields(I)
                If Val(m_BillDetail.Cb) > 0 Then
                  TempStr = IIf(m_BillDetail.RATE_DRIVER = 0, "", FormatNumber(m_BillDetail.RATE_DRIVER, 3))
               Else
                  TempStr = IIf(m_BillDetail.RATE_DRIVER = 0, "", FormatNumber(m_BillDetail.RATE_DRIVER, 2))
               End If
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Total1(I) = Total1(I) + m_BillDetail.RATE_DRIVER
               Total2(I) = Total2(I) + m_BillDetail.RATE_DRIVER
               
               If mcolParam("CUSTOMER_ONLY_FLAG") = 0 Then
                   I = I + 1
                   Set HeadCf = m_Headers1.Fields(I)
                   If Val(m_BillDetail.Cb) > 0 Then
                     TempStr = IIf(m_BillDetail.RATE_CUSTOMER = 0, "", FormatNumber(m_BillDetail.RATE_CUSTOMER, 3))
                  Else
                     TempStr = IIf(m_BillDetail.RATE_CUSTOMER = 0, "", FormatNumber(m_BillDetail.RATE_CUSTOMER, 2))
                  End If
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
                  Total1(I) = Total1(I) + m_BillDetail.RATE_CUSTOMER
                  Total2(I) = Total2(I) + m_BillDetail.RATE_CUSTOMER
               End If
               
                   I = I + 1
                   Set HeadCf = m_Headers1.Fields(I)
                   If Val(m_BillDetail.KEY_CODE) > 4 And Val(m_BillDetail.KEY_CODE) <> 9 And Val(m_BillDetail.KEY_CODE) <> 10 Then   'ต้องเป็น 4
                     TempStr = ""
                   Else
                     TempStr = IIf(m_BillDetail.PRICE_DRIVER = 0, "", FormatNumber(m_BillDetail.PRICE_DRIVER))
                     Total1(I) = Total1(I) + m_BillDetail.PRICE_DRIVER
                     Total2(I) = Total2(I) + m_BillDetail.PRICE_DRIVER
                     If m_BillDetail.NOT_CAL_VAT = "N" Then 'ถ้า bill ใบนี้คิดภาษีให้เข้าคำนวณ
                        Total3(I) = Total3(I) + (m_BillDetail.PRICE_DRIVER * TAX)
                     End If
                   End If
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               
               If Not mcolParam("CUSTOMER_ONLY_FLAG") = 1 Then
                   I = I + 1
                   
                   Set HeadCf = m_Headers1.Fields(I)
                   If Val(m_BillDetail.KEY_CODE) > 2 And Val(m_BillDetail.KEY_CODE) <> 8 And Val(m_BillDetail.KEY_CODE) <> 9 And Val(m_BillDetail.KEY_CODE) <> 10 Then
                     TempStr = ""
                   ElseIf Val(m_BillDetail.KEY_CODE) = 8 Then 'ส่วนลดลูกค้าคิดการตลาด
                     TempStr = IIf(m_BillDetail.PRICE_CUSTOMER = 0, "", "(" & FormatNumber(m_BillDetail.PRICE_CUSTOMER) & ")")
                     Total1(I) = Total1(I) - m_BillDetail.PRICE_CUSTOMER
                     Total2(I) = Total2(I) - m_BillDetail.PRICE_CUSTOMER
                   Else
                     TempStr = IIf(m_BillDetail.PRICE_CUSTOMER = 0, "", FormatNumber(m_BillDetail.PRICE_CUSTOMER))
                     Total1(I) = Total1(I) + m_BillDetail.PRICE_CUSTOMER
                     Total2(I) = Total2(I) + m_BillDetail.PRICE_CUSTOMER
                   End If
                  
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
   
                  I = I + 1 'กำไรขนส่ง
                  Set HeadCf = m_Headers1.Fields(I)
                  If Val(m_BillDetail.KEY_CODE) > 2 And Val(m_BillDetail.KEY_CODE) <> 9 And Val(m_BillDetail.KEY_CODE) <> 10 Then
                     m_BillDetail.PRICE_CUSTOMER = 0
                     m_BillDetail.PRICE_DRIVER = 0
                  End If
                  TempStr = ""
                  If m_BillDetail.CAL_PRICE_IN_PRODUCT = "Y" Then
                     If Val(m_BillDetail.KEY_CODE) <> 7 Then
                        TempStr = IIf((m_BillDetail.PRICE_SALE - m_BillDetail.PRICE_DRIVER) = 0, "0.00", FormatNumber(m_BillDetail.PRICE_SALE - m_BillDetail.PRICE_DRIVER))
                     End If
                     Total1(I) = Total1(I) + (m_BillDetail.PRICE_SALE - m_BillDetail.PRICE_DRIVER)
                     Total2(I) = Total2(I) + (m_BillDetail.PRICE_SALE - m_BillDetail.PRICE_DRIVER)
                  Else
                     If Val(m_BillDetail.KEY_CODE) <> 7 Then
                        TempStr = IIf((m_BillDetail.PRICE_CUSTOMER - m_BillDetail.PRICE_DRIVER) = 0, "0.00", FormatNumber(m_BillDetail.PRICE_CUSTOMER - m_BillDetail.PRICE_DRIVER))
                     End If
                      Total1(I) = Total1(I) + (m_BillDetail.PRICE_CUSTOMER - m_BillDetail.PRICE_DRIVER)
                      Total2(I) = Total2(I) + (m_BillDetail.PRICE_CUSTOMER - m_BillDetail.PRICE_DRIVER)
                   End If
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)

                  
                   I = I + 1
                   Set HeadCf = m_Headers1.Fields(I)
                  TempStr = IIf(m_BillDetail.PRICE_SALE = 0, "", FormatNumber(m_BillDetail.PRICE_SALE))
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
                  Total1(I) = Total1(I) + m_BillDetail.PRICE_SALE
                  Total2(I) = Total2(I) + m_BillDetail.PRICE_SALE
      
                   I = I + 1
                   Set HeadCf = m_Headers1.Fields(I)
                  TempStr = IIf(m_BillDetail.PRICE_FAC = 0, "", FormatNumber(m_BillDetail.PRICE_FAC))
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
                  Total1(I) = Total1(I) + m_BillDetail.PRICE_FAC
                  Total2(I) = Total2(I) + m_BillDetail.PRICE_FAC
               End If
      
                I = I + 1
                TempStr = ""
                Set HeadCf = m_Headers1.Fields(I) 'รายได้รับจ้าง
                 If m_BillDetail.CAL_DRIVER = "Y" Then
                     TempStr = IIf(m_BillDetail.PRICE_IMCOME_OTHER = 0, "", FormatNumber(m_BillDetail.PRICE_IMCOME_OTHER))
                     Total1(I) = Total1(I) + m_BillDetail.PRICE_IMCOME_OTHER
                     Total2(I) = Total2(I) + m_BillDetail.PRICE_IMCOME_OTHER
                 ElseIf m_BillDetail.CAL_CUSTOMER = "Y" Then
                     TempStr = ""
                 End If
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
                  I = I + 1
                  TempStr = ""
                Set HeadCf = m_Headers1.Fields(I) 'รายได้ลูกค้า
                 If m_BillDetail.CAL_DRIVER = "Y" Then
                     TempStr = ""
                 ElseIf m_BillDetail.CAL_CUSTOMER = "Y" Then
                     TempStr = IIf(m_BillDetail.PRICE_IMCOME_OTHER = 0, "", FormatNumber(m_BillDetail.PRICE_IMCOME_OTHER))
                     Total1(I) = Total1(I) + m_BillDetail.PRICE_IMCOME_OTHER
                     Total2(I) = Total2(I) + m_BillDetail.PRICE_IMCOME_OTHER
                 End If
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
            
                I = I + 1 'รายจ่ายรับจ้าง
                TempStr = ""
                Set HeadCf = m_Headers1.Fields(I)
                If m_BillDetail.CAL_DRIVER = "Y" Then
                     TempStr = IIf(m_BillDetail.PRICE_EXPEND_OTHER = 0, "", FormatNumber(m_BillDetail.PRICE_EXPEND_OTHER))
                     Total1(I) = Total1(I) + m_BillDetail.PRICE_EXPEND_OTHER
                     Total2(I) = Total2(I) + m_BillDetail.PRICE_EXPEND_OTHER
                 ElseIf m_BillDetail.CAL_CUSTOMER = "Y" Then
                     TempStr = ""
                 End If
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               
                I = I + 1 'รายจ่ายลูกค้า
                TempStr = ""
                Set HeadCf = m_Headers1.Fields(I)
                 If m_BillDetail.CAL_DRIVER = "Y" Then
                     TempStr = ""
                 ElseIf m_BillDetail.CAL_CUSTOMER = "Y" Then
                     TempStr = IIf(m_BillDetail.PRICE_EXPEND_OTHER = 0, "", FormatNumber(m_BillDetail.PRICE_EXPEND_OTHER))
                     Total1(I) = Total1(I) + m_BillDetail.PRICE_EXPEND_OTHER
                     Total2(I) = Total2(I) + m_BillDetail.PRICE_EXPEND_OTHER
                 End If
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               I = I + 1
               Set HeadCf = m_Headers1.Fields(I)
               If Val(m_BillDetail.KEY_CODE) > 2 Then
                  TempStr = m_BillDetail.NOTE_ITEM
               ElseIf Len(Trim(m_BillDetail.NOTE_ITEM)) > 0 Then
                  If Len(Trim(m_BillDetail.NOTE)) > 0 Then
                     TempStr = m_BillDetail.NOTE & " (" & m_BillDetail.NOTE_ITEM & ")"
                  Else
                     TempStr = m_BillDetail.NOTE_ITEM
                  End If
               Else
                  TempStr = m_BillDetail.NOTE
               End If
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
      
         Next m_BillDetail
      
         Call GenerateFooter("รวมยอด", "", "", Total1, m_Details1, 0)
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
         
         Call GenerateFooter("รวมสุทธิ", "", "", Total2, m_Details1, 0)
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
          
           If mcolParam("CUSTOMER_ONLY_FLAG") = 1 Then
'            NetTotal = Total2(12) - Total3(12) - (Total2(5) * 60) - Total2(13) + Total2(15)AddTrnDriver
            AddTrnDriver = Total2(15) * TAX
            ReduceTrnDriver = Total2(13) * TAX
             NetTotal = FormatNumber(Total2(12)) - FormatNumber(Total3(12)) - FormatNumber(Total2(5) * 60) - FormatNumber(Total2(13)) + FormatNumber(Total2(15)) - FormatNumber(AddTrnDriver) + FormatNumber(ReduceTrnDriver)
           Else
'            NetTotal = Total2(13) - Total3(13) - (Total2(5) * 60) - Total2(18) + Total2(20)
            AddTrnDriver = Total2(20) * TAX
            ReduceTrnDriver = Total2(18) * TAX
            NetTotal = FormatNumber(Total2(13)) - FormatNumber(Total3(13)) - FormatNumber(Total2(5) * 60) - FormatNumber(Total2(18)) + FormatNumber(Total2(20)) - FormatNumber(AddTrnDriver) + FormatNumber(ReduceTrnDriver)
          End If
          
         If Vsp.CurrentY > 8150 Then
                   mblnHeader = False
                  Vsp.NewPage
         End If
         Call GenerateFooter2(Total2, Total3)
   Else
       PrevKey1 = ""
        For Each m_BillDetail In m_DocDetail
                I = 5
                If PrevKey1 <> m_BillDetail.DOCUMENT_NO Then
                  PrevKey1 = m_BillDetail.DOCUMENT_NO
                   Total1(I) = Total1(I) + m_BillDetail.AROUND
                  Total2(I) = Total2(I) + m_BillDetail.AROUND
               End If
               
               If IsMain Then
               I = I + 1
               Total1(I) = Total1(I) + m_BillDetail.C10
               Total2(I) = Total2(I) + m_BillDetail.C10
               
                I = I + 1
               Total1(I) = Total1(I) + m_BillDetail.C20
               Total2(I) = Total2(I) + m_BillDetail.C20
               
                I = I + 1
               Total1(I) = Total1(I) + m_BillDetail.C30
               Total2(I) = Total2(I) + m_BillDetail.C30
               
                I = I + 1
               Total1(I) = Total1(I) + m_BillDetail.C50
               Total2(I) = Total2(I) + m_BillDetail.C50
               
                I = I + 1
               Total1(I) = Total1(I) + m_BillDetail.Cb
               Total2(I) = Total2(I) + m_BillDetail.Cb
               Else
                  I = I + 5
               End If
      
                I = I + 1
               Total1(I) = Total1(I) + m_BillDetail.RATE_DRIVER
               Total2(I) = Total2(I) + m_BillDetail.RATE_DRIVER
               
               If Not mcolParam("CUSTOMER_ONLY_FLAG") = 1 Then
                   I = I + 1
                  Total1(I) = Total1(I) + m_BillDetail.RATE_CUSTOMER
                  Total2(I) = Total2(I) + m_BillDetail.RATE_CUSTOMER
               End If
               
                   I = I + 1
                    If Not Val(m_BillDetail.KEY_CODE) > 4 Or Val(m_BillDetail.KEY_CODE) = 9 Or Val(m_BillDetail.KEY_CODE) = 10 Then
                     Total1(I) = Total1(I) + m_BillDetail.PRICE_DRIVER
                     Total2(I) = Total2(I) + m_BillDetail.PRICE_DRIVER
                     
                      If m_BillDetail.NOT_CAL_VAT = "N" Then 'ถ้า bill ใบนี้คิดภาษีให้เข้าคำนวณ
                        Total3(I) = Total3(I) + (m_BillDetail.PRICE_DRIVER * TAX)
                      End If
                   End If
               
               If Not mcolParam("CUSTOMER_ONLY_FLAG") = 1 Then
                   I = I + 1
'                   If Not (Val(m_BillDetail.KEY_CODE) > 2) Then
                   If Val(m_BillDetail.KEY_CODE) > 2 And Val(m_BillDetail.KEY_CODE) <> 8 Then
                   ElseIf Val(m_BillDetail.KEY_CODE) = 8 Then 'ส่วนลดลูกค้าคิดการตลาด
                     TempStr = IIf(m_BillDetail.PRICE_CUSTOMER = 0, "", "(" & FormatNumber(m_BillDetail.PRICE_CUSTOMER) & ")")
                     Total1(I) = Total1(I) - m_BillDetail.PRICE_CUSTOMER
                     Total2(I) = Total2(I) - m_BillDetail.PRICE_CUSTOMER
                   Else
                     TempStr = IIf(m_BillDetail.PRICE_CUSTOMER = 0, "", FormatNumber(m_BillDetail.PRICE_CUSTOMER))
                     Total1(I) = Total1(I) + m_BillDetail.PRICE_CUSTOMER
                     Total2(I) = Total2(I) + m_BillDetail.PRICE_CUSTOMER
                   End If
   
                  I = I + 1 'กำไรขนส่ง
                   If Val(m_BillDetail.KEY_CODE) > 2 And Val(m_BillDetail.KEY_CODE) <> 9 And Val(m_BillDetail.KEY_CODE) <> 10 Then
                     m_BillDetail.PRICE_CUSTOMER = 0
                     m_BillDetail.PRICE_DRIVER = 0
                  End If
                  Total1(I) = Total1(I) + (m_BillDetail.PRICE_CUSTOMER - m_BillDetail.PRICE_DRIVER)
                  Total2(I) = Total2(I) + (m_BillDetail.PRICE_CUSTOMER - m_BillDetail.PRICE_DRIVER)
                  
                   I = I + 1
                  Total1(I) = Total1(I) + m_BillDetail.PRICE_SALE
                  Total2(I) = Total2(I) + m_BillDetail.PRICE_SALE
      
                   I = I + 1
                  Total1(I) = Total1(I) + m_BillDetail.PRICE_FAC
                  Total2(I) = Total2(I) + m_BillDetail.PRICE_FAC
               End If
               
                   I = I + 1 'รายได้รับจ้าง
                    If m_BillDetail.CAL_DRIVER = "Y" Then
                        Total1(I) = Total1(I) + m_BillDetail.PRICE_IMCOME_OTHER
                        Total2(I) = Total2(I) + m_BillDetail.PRICE_IMCOME_OTHER
                    End If
                  
                     I = I + 1 'รายได้ลูกค้า
                  If m_BillDetail.CAL_CUSTOMER = "Y" Then
                        Total1(I) = Total1(I) + m_BillDetail.PRICE_IMCOME_OTHER
                        Total2(I) = Total2(I) + m_BillDetail.PRICE_IMCOME_OTHER
                  End If
               
                I = I + 1 'รายจ่ายรับจ้าง
                If m_BillDetail.CAL_DRIVER = "Y" Then
                     Total1(I) = Total1(I) + m_BillDetail.PRICE_EXPEND_OTHER
                     Total2(I) = Total2(I) + m_BillDetail.PRICE_EXPEND_OTHER
                 End If
               
                I = I + 1 'รายจ่ายลูกค้า
                 If m_BillDetail.CAL_CUSTOMER = "Y" Then
                     Total1(I) = Total1(I) + m_BillDetail.PRICE_EXPEND_OTHER
                     Total2(I) = Total2(I) + m_BillDetail.PRICE_EXPEND_OTHER
                 End If
         
         Next m_BillDetail
                
         If mcolParam("CUSTOMER_ONLY_FLAG") = 1 Then
'            NetTotal = Total2(12) - Total3(12) - (Total2(5) * 60) - Total2(13) + Total2(15)g
            NetTotal = FormatNumber(Total2(12)) - FormatNumber(Total3(12)) - FormatNumber(Total2(5) * 60) - FormatNumber(Total2(13)) + FormatNumber(Total2(15)) + FormatNumber(AddTrnDriver) - FormatNumber(ReduceTrnDriver)
           Else
'            NetTotal = Total2(13) - Total3(13) - (Total2(5) * 60) - Total2(18) + Total2(20)
            NetTotal = FormatNumber(Total2(13)) - FormatNumber(Total3(13)) - FormatNumber((Total2(5) * 60)) - FormatNumber(Total2(18)) + FormatNumber(Total2(20)) + FormatNumber(AddTrnDriver) - FormatNumber(ReduceTrnDriver)
          End If
   End If
      
      If mcolParam("SUMMARY_DOC_FLAG") = 1 Then
         Vsp.ORIENTATION = orPortrait
      If DocFlag = 0 Then
            DocFlag = 2
            Vsp.StartDoc
      Else
            DocFlag = 2
            Vsp.NewPage
        End If
         Vsp.MarginBottom = 1440
         Vsp.MarginFooter = 1440
         Vsp.MarginHeader = 1440
         Vsp.MarginLeft = 700
         Vsp.MarginRight = 700
         Vsp.MarginTop = 1440
         Vsp.FontName = "AngsanaUPC"
         mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
         Call GenerateFooter3(Total2, Total3)
      End If
      
      If mcolParam("PAYMENT_VOUCHER_FLAG") = 1 Then
        Vsp.ORIENTATION = orPortrait
      If DocFlag = 0 Then
             DocFlag = 3
            Vsp.StartDoc
      Else
            DocFlag = 3
            Vsp.NewPage
        End If

         Vsp.MarginBottom = 1440
         Vsp.MarginFooter = 1440
         Vsp.MarginHeader = 1440
         Vsp.MarginLeft = 700
         Vsp.MarginRight = 700
         Vsp.MarginTop = 1440
         Vsp.FontName = "AngsanaUPC"
         mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
         Vsp.TextAlign = taLeftBottom
         
         If mcolParam("USE_FORM_FLAG") = 1 Then
            Vsp.DrawPicture mcolParam("BACK_GROUND"), "0.0cm", "0.0cm", Vsp.PageWidth, Vsp.PageHeight, 10
         End If
         Call GenerateFooter4(Total2, Total3)
      End If
      
      If mcolParam("ACCOUNT_TRANSFER_FLAG") = 1 Then
        Vsp.ORIENTATION = orPortrait
      If DocFlag = 0 Then
            DocFlag = 4
            Vsp.StartDoc
      Else
            DocFlag = 4
            Vsp.NewPage
        End If

         Vsp.MarginBottom = 1440
         Vsp.MarginFooter = 1440
         Vsp.MarginHeader = 1440
         Vsp.MarginLeft = 700
         Vsp.MarginRight = 700
         Vsp.MarginTop = 1440
         Vsp.FontName = "AngsanaUPC"
         mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
         Vsp.TextAlign = taLeftBottom
          If mcolParam("USE_FORM_FLAG") = 1 Then
            Vsp.DrawPicture mcolParam("BACK_GROUND2"), "0.0cm", "0.0cm", Vsp.PageWidth, Vsp.PageHeight, 10
         End If
         Call GenerateFooter5(Total2, Total3)
      End If
      
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
   If mcolParam("CUSTOMER_ONLY_FLAG") = 1 Then
    For Each HeadCf In m_Headers1.Fields
      J = J + 1
      If J = 1 Then
         TempStr = ""
      ElseIf J = 2 Then
         TempStr = TXT2
      ElseIf J = 3 Then
         TempStr = TXT3
      ElseIf (J = 4) Then
         TempStr = TXT1
      ElseIf (J = 5) Or (J = 6) Or (J = 7) Or (J = 8) Or (J = 9) Then
         TempStr = FormatNumber(Tot(J), 0)
      ElseIf (J = 11) Or (J = 17) Then
         TempStr = ""
      ElseIf (J = m_Headers1.Fields.Count) Then 'ช่องหมายเหตุ
         TempStr = ""
      Else
         TempStr = FormatNumber(Tot(J))
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.Align)
      Call Details.AddField(BodyCf)
   Next HeadCf
   Else
      For Each HeadCf In m_Headers1.Fields
         J = J + 1
         If J = 1 Then
            TempStr = ""
         ElseIf J = 2 Then
            TempStr = TXT2
         ElseIf J = 3 Then
            TempStr = TXT3
         ElseIf (J = 4) Then
            TempStr = TXT1
         ElseIf (J = 5) Or (J = 6) Or (J = 7) Or (J = 8) Or (J = 9) Then
            TempStr = FormatNumber(Tot(J), 0)
         ElseIf (J = 11) Or (J = 12) Or (J = 22) Then
            TempStr = ""
         ElseIf (J = m_Headers1.Fields.Count) Then 'ช่องหมายเหตุ
           TempStr = ""
         Else
            TempStr = FormatNumber(Tot(J))
         End If
         
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.Align)
         Call Details.AddField(BodyCf)
      Next HeadCf
   End If
      
   Set BodyCf = Nothing
End Sub
Private Sub GenerateTotal(Tot() As Double, Tot3() As Double)
Dim alngX() As Long
Dim strFormat As String
Dim strPrint As String
Dim tempCurentY As Double
Dim HaveNet As Boolean
Dim TempY  As Double
Dim TempY2 As Double

      Vsp.TableBorder = tbNone
      Vsp.FontSize = 14
      
      
       
         If Vsp.ORIENTATION = orPortrait Then
            strFormat = VSP_CalTable("<17|<12|>10|<10|<8;", mdblWid, alngX)
         Else
            strFormat = VSP_CalTable("<32|<12|>10|<10|<23;", mdblWid, alngX)
         End If
      
      If mcolParam("CUSTOMER_ONLY_FLAG") = 1 Then
          Vsp.StartTable
          strPrint = "|รวมขนส่ง|" & FormatNumber(Tot(12)) & "|บาท|;"
         Call Vsp.AddTable(strFormat, "", strPrint)
         Vsp.EndTable
         
          If Val(Tot(13)) > 0 Then
            TempY = Vsp.CurrentY
            Call AddLine(TempY, strFormat, strPrint)
            Vsp.CurrentY = TempY
            
           Vsp.StartTable
            strPrint = "|หักค่าขนส่ง|(" & FormatNumber(Tot(13)) & ")|บาท|;"
            Call Vsp.AddTable(strFormat, "", strPrint)
            Vsp.EndTable
         End If
         
         If Val(Tot(15)) > 0 Then
            TempY = Vsp.CurrentY
            Call AddLine(TempY, strFormat, strPrint)
            Vsp.CurrentY = TempY
            
            Vsp.StartTable
            strPrint = "|เพิ่มค่าขนส่ง|" & FormatNumber(Tot(15)) & "|บาท|;"
            Call Vsp.AddTable(strFormat, "", strPrint)
            Vsp.EndTable
         End If
         
         TempY = Vsp.CurrentY
         Call AddLine(TempY, strFormat, strPrint)
         Call AddLine(TempY + 40, strFormat, strPrint)
         Vsp.CurrentY = TempY
         
         Vsp.StartTable
         strPrint = "|รวมค่าขนส่ง|" & FormatNumber(Tot(12) - Tot(13) + Tot(15)) & "|บาท|;"
         Call Vsp.AddTable(strFormat, "", strPrint)
         Vsp.EndTable

         
         If Val(Tot3(12)) > 0 Then
            Vsp.StartTable
            strPrint = "|หักภาษี  1%|(" & FormatNumber(Tot3(12) + AddTrnDriver - ReduceTrnDriver) & ")|บาท|;" 'เพิ่มค่าขนส่ง บวกเพิ่มไปด้วย
            Call Vsp.AddTable(strFormat, "", strPrint)
            Vsp.EndTable
         End If
   
         If Val(Tot(5)) * 60 > 0 Then
            TempY = Vsp.CurrentY
            Call AddLine(TempY, strFormat, strPrint)
            Call AddLine(TempY + 40, strFormat, strPrint)
            Vsp.CurrentY = TempY
         
           Vsp.StartTable
            strPrint = "|หักค่าดำเนินการ(เที่ยว)|(" & FormatNumber(Tot(5) * 60) & ")|บาท|;"
            Call Vsp.AddTable(strFormat, "", strPrint)
            Vsp.EndTable
         End If

      Else
         Vsp.StartTable
         strPrint = "|รวมขนส่ง|" & FormatNumber(Tot(13)) & "|บาท|;"
         Call Vsp.AddTable(strFormat, "", strPrint)
         Vsp.EndTable
         
         If Val(Tot(18)) > 0 Then
            TempY = Vsp.CurrentY
            Call AddLine(TempY, strFormat, strPrint)
            Vsp.CurrentY = TempY
            
           Vsp.StartTable
            strPrint = "|หักค่าขนส่ง|(" & FormatNumber(Tot(18)) & ")|บาท|;"
            Call Vsp.AddTable(strFormat, "", strPrint)
            Vsp.EndTable
         End If
         
         If Val(Tot(20)) > 0 Then
            TempY = Vsp.CurrentY
            Call AddLine(TempY, strFormat, strPrint)
            Vsp.CurrentY = TempY
            
           Vsp.StartTable
            strPrint = "|เพิ่มค่าขนส่ง|" & FormatNumber(Tot(20)) & "|บาท|;"
            Call Vsp.AddTable(strFormat, "", strPrint)
            Vsp.EndTable
         End If

         TempY = Vsp.CurrentY
         Call AddLine(TempY, strFormat, strPrint)
         Call AddLine(TempY + 40, strFormat, strPrint)
         Vsp.CurrentY = TempY
         
         Vsp.StartTable
         strPrint = "|รวมค่าขนส่ง|" & FormatNumber(Tot(13) - Tot(18) + Tot(20)) & "|บาท|;"
         Call Vsp.AddTable(strFormat, "", strPrint)
         Vsp.EndTable
         
         If Val(Tot3(13)) > 0 Then
            Vsp.StartTable
            strPrint = "|หักภาษี  1%|(" & FormatNumber(Tot3(13) + AddTrnDriver - ReduceTrnDriver) & ")|บาท|;" 'เอา เพิ่มค่าขนส่ง ไปบวกเพิ่มด้วย
            Call Vsp.AddTable(strFormat, "", strPrint)
            Vsp.EndTable
         End If
   
         If Val(Tot(5)) * 60 > 0 Then
            TempY = Vsp.CurrentY
            Call AddLine(TempY, strFormat, strPrint)
            Call AddLine(TempY + 40, strFormat, strPrint)
            Vsp.CurrentY = TempY
            
            Vsp.StartTable
            strPrint = "|หักค่าดำเนินการ(เที่ยว)|(" & FormatNumber(Tot(5) * 60) & ")|บาท|;"
            Call Vsp.AddTable(strFormat, "", strPrint)
            Vsp.EndTable
         End If
         
      End If
'      Vsp.EndTable

      
'      tempCurentY = Vsp.CurrentY
     
      
      TempY = Vsp.CurrentY
      Call AddLine(TempY, strFormat, strPrint, True)
      Call AddLine(TempY + 40, strFormat, strPrint, True)
      Vsp.CurrentY = TempY
      
      Vsp.FontBold = True
      Vsp.StartTable
      strPrint = "|ค่าขนส่งสุทธิ|" & FormatNumber(NetTotal) & "|บาท|;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      Vsp.EndTable
      Vsp.FontBold = False

'       Vsp.CurrentY = tempCurentY 'ปรับให้เป็น CurrentY ปัจจุบัน
      Vsp.TableBorder = tbNone
End Sub
Private Sub AddLine(CurrenY As Double, strFormat As String, strPrint As String, Optional FBold As Boolean = False)
      Vsp.CurrentY = CurrenY
      Vsp.FontBold = FBold
      Vsp.StartTable
      strPrint = "||_____________||;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      Vsp.EndTable
      Vsp.FontBold = False
End Sub
Private Sub GenerateFooter2(Tot() As Double, Tot3() As Double)
Dim I As Long
Dim alngX() As Long
Dim strFormat As String
Dim strPrint As String
       Vsp = "* หมายถึง ไม่หักภาษี 1 %"
       Call GenerateTotal(Tot, Tot3)
   
       Vsp.StartTable
      Vsp.TableBorder = tbNone
      Vsp.CurrentY = 10000
      strFormat = VSP_CalTable("^15|^15|^15;", mdblWid, alngX)
      strPrint = ".....................................................|.....................................................|.....................................................;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      strPrint = "ผู้จัดทำ|ผู้ขนส่ง|ผู้ตรวจสอบ;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      strPrint = "......../......../........|......../......../........|......../......../........;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
       Vsp.EndTable
End Sub
Private Sub GenerateFooter3(Tot() As Double, Tot3() As Double)
Dim I As Long
Dim alngX() As Long
Dim strFormat As String
Dim strPrint As String
Dim strDes1 As String
Dim strDes2 As String
Dim strDes3 As String
Dim TempY As Double

Dim strDate As String

         If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
            mstrErrMsg = "Error GetDateTime Error."
            Exit Sub
         End If
          mdteDate = InternalDateToDate(strDate)
            
            Vsp.FontSize = 8
            Vsp.CurrentY = 400
            TempY = Vsp.CurrentY
            Vsp.TextAlign = taLeftBottom
            strFormat = VSP_CalTable(TITLE_FORMAT, mdblWid, alngX)
            Vsp.TableBorder = tbNone
            Vsp.FontBold = False
            Call Vsp.AddTable(strFormat, "", "|วันที่พิมพ์|: " & DateToStringExt(mdteDate))
            Call Vsp.AddTable(strFormat, "", "|เวลา|: " & Right$(DateToStringExtEx(mdteDate), 8))
            Call Vsp.AddTable(strFormat, "", "|ผู้ใช้|: " & glbUser.USER_NAME)
            Vsp.CurrentY = TempY
            
            
      Vsp.FontSize = 24
      Vsp.TextAlign = taLeftBottom
      strFormat = VSP_CalTable(TITLE_FORMAT, mdblWid, alngX)
      Vsp.TableBorder = tbNone
      Vsp.CurrentY = 100
      Vsp.FontBold = True
      Vsp.TextAlign = taCenterBottom
      Vsp.Paragraph = vbCrLf & glbEnterPrise.ENTERPRISE_NAME
      
      Vsp.FontSize = 14
      Vsp.FontBold = False
      Vsp = ""
      Vsp.TableBorder = tbNone
      strFormat = VSP_CalTable("^7|<3;", mdblWid, alngX)
      strDes1 = DateToStringExt(mcolParam("TO_DATE"))
      strPrint = "|" & strDes1 & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      Vsp = ""
      Vsp.TableBorder = tbNone
      strFormat = VSP_CalTable("^3|<25;", mdblWid, alngX)
      strDes1 = "เรื่อง เรียกเก็บเงินค่าขนส่ง"
      strDes2 = "เรียน ผู้จัดการฝ่ายบัญชี"
      strPrint = "|" & strDes1 & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
      strPrint = "|" & strDes2 & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      Vsp = ""
      Vsp.TableBorder = tbNone
      strFormat = VSP_CalTable("^5|<20;", mdblWid, alngX)
      strDes1 = "ตั้งแต่วันที่ " & DateToStringExtEx2(mcolParam("FROM_DATE")) & " - " & DateToStringExtEx2(mcolParam("TO_DATE")) & " ทางบริษัท มิตรภาพอาหารสัตว์ จำกัด ได้จัดรถบรรทุก "
      strDes2 = strTruck & " ส่งอาหาร ให้กับทาง บริษัท มิตรภาพ จำกัด เป็นจำนวนเงินสุทธิ " & FormatNumber(NetTotal) & " บาท"
      strDes3 = "โดยขอให้สั่งจ่ายในนามของ " & Sp.SUPPLIER_NAME & " ตามรายการดังนี้"
      strPrint = "|" & strDes1 & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
      strPrint = "|" & strDes2 & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
      strPrint = "|" & strDes3 & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)

      Call GenerateTotal(Tot, Tot3)
  
      Vsp = ""
      Vsp = ""
      Vsp.TableBorder = tbNone
      strFormat = VSP_CalTable("^7|<14.5;", mdblWid, alngX)
      strDes1 = ThaiBaht(NetTotal)
      strPrint = "|" & strDes1 & ";"
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", strPrint)
      Vsp.FontBold = False
       
      Vsp = ""
      Vsp.TableBorder = tbNone
      strFormat = VSP_CalTable("^5|<20;", mdblWid, alngX)
      If SA.USE_TRANSPORT_FLAG = "Y" Then
         strDes1 = "ทั้งนี้ขอให้จ่ายเช็ค โอนเข้าบัญชี ในนาม " & SA.SUPPLIER_NAME & " ดังรายละเอียด ดังนี้" & vbNewLine
         strDes1 = strDes1 & "เลขที่บัญชี " & SA.SUPPLIER_ACCOUNT_NO & vbNewLine
         strDes1 = strDes1 & "ชื่อบัญชี " & SA.SUPPLIER_ACCOUNT_BANK & " สาขา " & SA.SUPPLIER_ACCOUNT_BRANCH
      Else
         strDes1 = "ทั้งนี้ขอให้จ่ายเช็ค ให้แก่ผู้รับเงินที่โรงงานอาหารสัตว์ (MM)"
      End If
      strPrint = "|" & strDes1 & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)

      Vsp.CurrentY = Vsp.CurrentY + ((Vsp.PageHeight - Vsp.CurrentY) / 2) '10550
      strFormat = VSP_CalTable("^1|^1|^1;", mdblWid, alngX)
      strPrint = ".....................................................|.....................................................|.....................................................;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strPrint = "ผู้จัดทำ|ผู้ตรวจสอบ|ผู้อนุมัติ;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      strPrint = "......../......../........|......../......../........|......../......../........;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
   End Sub
Private Sub GenerateFooter4(Tot() As Double, Tot3() As Double)
Dim I As Long
Dim J As Long
Dim alngX() As Long
Dim strFormat As String
Dim strPrint As String
Dim strDes1 As String
Dim strDes2 As String
Dim strDes3 As String
Dim TempTot(100) As Double
Dim TempTot2(100) As Double
Dim CerrentY As Double

      For J = 1 To UBound(TempTot)
         TempTot(J) = 0
         TempTot2(J) = 0
      Next J
      Vsp.CurrentY = 11160
      Vsp.FontSize = 14.25
      Vsp.CurrentY = 2300
      Vsp.TableBorder = tbNone
      
      strFormat = VSP_CalTable("^4|<25;", mdblWid, alngX)
      strDes1 = Sp.SUPPLIER_NAME
      strPrint = "|" & strDes1 & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      Vsp.CurrentY = Vsp.CurrentY + 450
      strFormat = VSP_CalTable("^4|<20;", mdblWid, alngX)
      strDes1 = "ค่าขนส่งอาหารระหว่างวันที่  " & DateToStringExtEx2(mcolParam("FROM_DATE")) & " - " & DateToStringExtEx2(mcolParam("TO_DATE"))
      strPrint = "|" & strDes1 & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      Vsp.CurrentY = Vsp.CurrentY + 490
      strFormat = VSP_CalTable("^4|<20;", mdblWid, alngX)
      
      Dim strDate As String
      Dim cDates As Date
      If Day(mcolParam("FROM_DATE")) >= 1 And Day(mcolParam("FROM_DATE")) <= 15 Then
         cDates = DateSerial(Year(mcolParam("FROM_DATE")), Month(mcolParam("FROM_DATE")), 25)
         If DatePart("w", cDates) = vbSaturday Then
            strDes1 = DateToStringExtEx2(DateSerial(Year(mcolParam("FROM_DATE")), Month(mcolParam("FROM_DATE")), 24))
         ElseIf DatePart("w", cDates) = vbSunday Then
            strDes1 = DateToStringExtEx2(DateSerial(Year(mcolParam("FROM_DATE")), Month(mcolParam("FROM_DATE")), 23))
         Else
            strDes1 = DateToStringExtEx2(cDates)
         End If
      Else
         cDates = DateAdd("M", 1, mcolParam("FROM_DATE"))
         cDates = DateSerial(Year(cDates), Month(cDates), 10)
         If DatePart("w", cDates) = vbSaturday Then
            strDes1 = DateToStringExtEx2(DateSerial(Year(mcolParam("FROM_DATE")), Month(mcolParam("FROM_DATE")) + 1, 9))
         ElseIf DatePart("w", cDates) = vbSunday Then
            strDes1 = DateToStringExtEx2(DateSerial(Year(mcolParam("FROM_DATE")), Month(mcolParam("FROM_DATE")) + 1, 8))
         Else
            strDes1 = DateToStringExtEx2(cDates)
         End If
      End If
      strDate = strDes1
      strPrint = "|" & strDate & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      Vsp.CurrentY = Vsp.CurrentY + 450
      strFormat = VSP_CalTable("^4|<5|<15;", mdblWid, alngX)
      strDes1 = FormatNumber(NetTotal)
      strDes2 = ThaiBaht(NetTotal)
      strPrint = "|" & strDes1 & "|" & strDes2 & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      Vsp.CurrentY = Vsp.CurrentY + 450
      strFormat = VSP_CalTable("^10|<10;", mdblWid, alngX)
      strDes1 = strDate
      strPrint = "|" & strDes1 & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      If mcolParam("ACCOUNT_TRANSFER_FLAG") = 1 Then 'ถ้ามีการเรียกใบโอนด้วย
         Vsp.TableBorder = tbNone
         Vsp.FontSize = 14
         Vsp.StartTable
         Vsp.TableCell(tcCols) = 5
         Vsp.TableCell(tcRows) = 1
          Vsp.TableCell(tcRowHeight, 1) = 1 * 627
          Vsp.CurrentY = Vsp.CurrentY + 1530
   
            For I = 1 To Vsp.TableCell(tcRows)
               Vsp.TableCell(tcColWidth, I, 1) = ".1cm"
               Vsp.TableCell(tcColWidth, I, 2) = "7.9cm"
               Vsp.TableCell(tcColWidth, I, 3) = "3cm"
               Vsp.TableCell(tcColWidth, I, 4) = "3.1cm"
               Vsp.TableCell(tcColWidth, I, 5) = "4cm"
                
                Vsp.TableCell(tcAlign, I, 1) = taLeftBottom
                Vsp.TableCell(tcAlign, I, 2) = taLeftBottom
                Vsp.TableCell(tcAlign, I, 3) = taLeftBottom
                Vsp.TableCell(tcAlign, I, 4) = taRightBottom
                Vsp.TableCell(tcAlign, I, 5) = taRightBottom
                For J = 1 To 20
                     If J = 13 And Val(Tot(J)) > 0 And TempTot(J) = 0 Then
                        Vsp.TableCell(tcText, I, 2) = "ค่าใช้จ่ายค้างจ่าย"
                        Vsp.TableCell(tcText, I, 3) = "217-3999"
                        Vsp.TableCell(tcText, I, 4) = FormatNumber(NetTotal + Tot3(J) + AddTrnDriver - ReduceTrnDriver) '(Tot(20) * 0.01)+ AddTrnDriver
                        Vsp.TableCell(tcText, I, 5) = ""
                        TempTot(J) = -1
                        TempTot2(4) = TempTot2(4) + NetTotal + Tot3(J) + AddTrnDriver - ReduceTrnDriver
                        Exit For
                     End If
                Next J
          Next I
          Vsp.EndTable
          
          Vsp.StartTable
         Vsp.TableCell(tcCols) = 5
         Vsp.TableCell(tcRows) = 6
          Vsp.TableCell(tcRowHeight, 1) = 1 * 597 '567
          Vsp.CurrentY = Vsp.CurrentY - 140
   
            For I = 1 To Vsp.TableCell(tcRows)
               Vsp.TableCell(tcColWidth, I, 1) = ".5cm"
               Vsp.TableCell(tcColWidth, I, 2) = "7.8cm"
               Vsp.TableCell(tcColWidth, I, 3) = "2.7cm"
               Vsp.TableCell(tcColWidth, I, 4) = "3.1cm"
               Vsp.TableCell(tcColWidth, I, 5) = "4cm"
   
                Vsp.TableCell(tcAlign, I, 1) = taLeftBottom
                Vsp.TableCell(tcAlign, I, 2) = taLeftBottom
                Vsp.TableCell(tcAlign, I, 3) = taLeftBottom
                Vsp.TableCell(tcAlign, I, 4) = taRightBottom
                Vsp.TableCell(tcAlign, I, 5) = taRightBottom
                For J = 1 To 20
                      If J = 13 And Val(Tot(J)) > 0 And TempTot(J) = -1 Then
                        Vsp.TableCell(tcText, I, 2) = "ภาษี หัก ณ ที่จ่าย"
                        Vsp.TableCell(tcText, I, 3) = "217-5100"
                        Vsp.TableCell(tcText, I, 4) = ""
'                        Vsp.TableCell(tcText, I, 5) = FormatNumber(Tot3(J))
                         Vsp.TableCell(tcText, I, 5) = FormatNumber(Tot3(J) + AddTrnDriver - ReduceTrnDriver) '(Tot(20) * 0.01)
                        TempTot(J) = -2
                        TempTot2(5) = TempTot2(5) + Tot3(J) + AddTrnDriver - ReduceTrnDriver
                        Exit For
                     ElseIf J = 13 And Val(Tot(J)) > 0 And TempTot(J) = -2 Then
                        Vsp.TableCell(tcText, I, 2) = "เช็คจ่ายลงวันที่ล่วงหน้า"
                        Vsp.TableCell(tcText, I, 3) = "212-2200"
                        Vsp.TableCell(tcText, I, 4) = ""
                        Vsp.TableCell(tcText, I, 5) = FormatNumber(NetTotal)
                        TempTot(J) = -3
                        TempTot2(5) = TempTot2(5) + NetTotal
                         Exit For
                     End If
                Next J
          Next I
          Vsp.EndTable

      Else '***********************************
         Vsp.TableBorder = tbNone
         Vsp.FontSize = 14
         Vsp.StartTable
         Vsp.TableCell(tcCols) = 5
         Vsp.TableCell(tcRows) = 3
          Vsp.TableCell(tcRowHeight, 1) = 1 * 627
          Vsp.CurrentY = Vsp.CurrentY + 1530
   
            For I = 1 To Vsp.TableCell(tcRows)
               Vsp.TableCell(tcColWidth, I, 1) = ".1cm"
               Vsp.TableCell(tcColWidth, I, 2) = "7.9cm"
               Vsp.TableCell(tcColWidth, I, 3) = "3cm"
               Vsp.TableCell(tcColWidth, I, 4) = "3.1cm"
               Vsp.TableCell(tcColWidth, I, 5) = "4cm"
                
                Vsp.TableCell(tcAlign, I, 1) = taLeftBottom
                Vsp.TableCell(tcAlign, I, 2) = taLeftBottom
                Vsp.TableCell(tcAlign, I, 3) = taLeftBottom
                Vsp.TableCell(tcAlign, I, 4) = taRightBottom
                Vsp.TableCell(tcAlign, I, 5) = taRightBottom
                For J = 1 To 20
                     If J = 13 And Val(Tot(J)) > 0 And TempTot(J) = 0 Then
                        Vsp.TableCell(tcText, I, 2) = "ค่าขนส่งอาหาร"
                        Vsp.TableCell(tcText, I, 3) = "570-5007"
                        Vsp.TableCell(tcText, I, 4) = FormatNumber(Tot(J))
                        Vsp.TableCell(tcText, I, 5) = ""
                        TempTot(J) = -1
                        TempTot2(4) = TempTot2(4) + Tot(J)
                        Exit For
                     ElseIf J = 16 And Val(Tot(J)) > 0 And TempTot(J) = 0 Then
                        Vsp.TableCell(tcText, I, 2) = "รายได้จากการขายอาหาร-บุคคลอื่น(เก็บการตลาด)"
                        Vsp.TableCell(tcText, I, 3) = "411-2006"
                        Vsp.TableCell(tcText, I, 4) = FormatNumber(Tot(J))
                        Vsp.TableCell(tcText, I, 5) = ""
                        TempTot(J) = -1
                        TempTot2(4) = TempTot2(4) + Tot(J)
                        Exit For
                      ElseIf J = 17 And Val(Tot(J)) > 0 And TempTot(J) = 0 Then
                        Vsp.TableCell(tcText, I, 2) = "ขนส่งเก็บโรงงาน"
                        Vsp.TableCell(tcText, I, 3) = ""
                        Vsp.TableCell(tcText, I, 4) = FormatNumber(Tot(J))
                        Vsp.TableCell(tcText, I, 5) = ""
                        TempTot(J) = -1
                        TempTot2(4) = TempTot2(4) + Tot(J)
                         Exit For
                       ElseIf J = 19 And Val(Tot(J)) > 0 And TempTot(J) = 0 Then
                        Vsp.TableCell(tcText, I, 2) = "รายจ่ายอื่นๆ"
                        Vsp.TableCell(tcText, I, 3) = "610-4020"
                        Vsp.TableCell(tcText, I, 4) = FormatNumber(Tot(J))
                        Vsp.TableCell(tcText, I, 5) = ""
                        TempTot(J) = -1
                        TempTot2(4) = TempTot2(4) + (Tot(J))
                         Exit For
                     End If
                Next J
          Next I
          Vsp.EndTable
          
          Vsp.StartTable
         Vsp.TableCell(tcCols) = 5
         Vsp.TableCell(tcRows) = 6
          Vsp.TableCell(tcRowHeight, 1) = 1 * 597 '567
          Vsp.CurrentY = Vsp.CurrentY - 140
   
            For I = 1 To Vsp.TableCell(tcRows)
               Vsp.TableCell(tcColWidth, I, 1) = ".5cm"
               Vsp.TableCell(tcColWidth, I, 2) = "7.8cm"
               Vsp.TableCell(tcColWidth, I, 3) = "2.7cm"
               Vsp.TableCell(tcColWidth, I, 4) = "3.1cm"
               Vsp.TableCell(tcColWidth, I, 5) = "4cm"
   
                Vsp.TableCell(tcAlign, I, 1) = taLeftBottom
                Vsp.TableCell(tcAlign, I, 2) = taLeftBottom
                Vsp.TableCell(tcAlign, I, 3) = taLeftBottom
                Vsp.TableCell(tcAlign, I, 4) = taRightBottom
                Vsp.TableCell(tcAlign, I, 5) = taRightBottom
                For J = 1 To 20
                     If J = 16 And Val(Tot(J)) > 0 And TempTot(J) = -1 Then
                        Vsp.TableCell(tcText, I, 2) = "รายได้ค่าขนส่งบุคคลอื่นๆ"
                        Vsp.TableCell(tcText, I, 3) = "413-2006"
                        Vsp.TableCell(tcText, I, 4) = ""
                        Vsp.TableCell(tcText, I, 5) = FormatNumber(Tot(J))
                        TempTot(J) = -2
                        TempTot2(5) = TempTot2(5) + Tot(J)
                         Exit For
                       ElseIf J = 5 And Val(Tot(J)) > 0 And TempTot(J) = 0 Then
                        Vsp.TableCell(tcText, I, 2) = "รายได้อื่นๆ(ค่าดำเนินการ)"
                        Vsp.TableCell(tcText, I, 3) = "420-9999"
                        Vsp.TableCell(tcText, I, 4) = ""
                        Vsp.TableCell(tcText, I, 5) = FormatNumber(Tot(J) * 60)
                        TempTot(J) = -1
                        TempTot2(5) = TempTot2(5) + (Tot(J) * 60)
                         Exit For
                       ElseIf J = 18 And Val(Tot(J)) > 0 And TempTot(J) = 0 Then
                        Vsp.TableCell(tcText, I, 2) = "รายได้หักอื่นๆ"
                        Vsp.TableCell(tcText, I, 3) = "420-9999"
                        Vsp.TableCell(tcText, I, 4) = ""
                        Vsp.TableCell(tcText, I, 5) = FormatNumber(Tot(J))
                        TempTot(J) = -1
                        TempTot2(5) = TempTot2(5) + (Tot(J))
                         Exit For
                      ElseIf J = 13 And Val(Tot(J)) > 0 And TempTot(J) = -1 Then
                        Vsp.TableCell(tcText, I, 2) = "ภาษี หัก ณ ที่จ่าย"
                        Vsp.TableCell(tcText, I, 3) = "217-5100"
                        Vsp.TableCell(tcText, I, 4) = ""
'                        Vsp.TableCell(tcText, I, 5) = FormatNumber(Tot3(J))
                        Vsp.TableCell(tcText, I, 5) = FormatNumber(Tot3(J) + AddTrnDriver - ReduceTrnDriver) '(Tot(20) * 0.01)
                        TempTot(J) = -2
                        TempTot2(5) = TempTot2(5) + (Tot3(J))
                        Exit For
                     ElseIf J = 17 And Val(Tot(J)) > 0 And TempTot(J) = -1 Then
                        Vsp.TableCell(tcText, I, 2) = "ค่าขนส่งอาหาร(เก็บโรงงาน)"
                        Vsp.TableCell(tcText, I, 3) = "570-5007"
                        Vsp.TableCell(tcText, I, 4) = ""
                        Vsp.TableCell(tcText, I, 5) = FormatNumber(Tot(J))
                        TempTot(J) = -2
                        TempTot2(5) = TempTot2(5) + Tot(J)
                         Exit For
                     ElseIf J = 13 And Val(Tot(J)) > 0 And TempTot(J) = -2 Then
                        Vsp.TableCell(tcText, I, 2) = "เช็คจ่ายลงวันที่ล่วงหน้า"
                        Vsp.TableCell(tcText, I, 3) = "212-2200"
                        Vsp.TableCell(tcText, I, 4) = ""
                        Vsp.TableCell(tcText, I, 5) = FormatNumber(NetTotal)
                        TempTot(J) = -3
                        TempTot2(5) = TempTot2(5) + NetTotal
                        Exit For
                     End If
                Next J
          Next I
          Vsp.EndTable
       End If
       
       Vsp.StartTable
      Vsp.TableCell(tcCols) = 3
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight, 1) = 1 * 567
      Vsp.CurrentY = 12519 'Vsp.CurrentY + 2000

      Vsp.TableCell(tcColWidth, 1, 1) = "11cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "3.1cm"
      Vsp.TableCell(tcColWidth, 1, 3) = "4cm"
       Vsp.TableCell(tcAlign, 1, 1) = taLeftBottom
       Vsp.TableCell(tcAlign, 1, 2) = taRightBottom
       Vsp.TableCell(tcAlign, 1, 3) = taRightBottom
'      If Val(TempTot2(4)) = Val(TempTot2(5)) Then
       If FormatNumber(TempTot2(4)) = FormatNumber(TempTot2(5)) Then
         Vsp.TableCell(tcText, 1, 1) = ThaiBaht(TempTot2(4))
      Else
         Vsp.TableCell(tcText, 1, 1) = "* ผลรวม Debit กับ Credit ไม่เท่ากัน"
      End If
      Vsp.TableCell(tcText, 1, 2) = FormatNumber(TempTot2(4))
      Vsp.TableCell(tcText, 1, 3) = FormatNumber(TempTot2(5))
      
         
      Vsp.EndTable
            
   End Sub
Private Sub GenerateFooter5(Tot() As Double, Tot3() As Double)
Dim I As Long
Dim J As Long
Dim alngX() As Long
Dim strFormat As String
Dim strPrint As String
Dim strDes1 As String
Dim strDes2 As String
Dim strDes3 As String
Dim TempTot(100) As Double
Dim TempTot2(100) As Double
Dim CerrentY As Double

      For J = 1 To UBound(TempTot)
         TempTot(J) = 0
         TempTot2(J) = 0
      Next J
'      Vsp.CurrentY = 14560 ' 11160
      Vsp.FontSize = 14.25
      Vsp.CurrentY = 2400
      Vsp.TableBorder = tbNone
      
      Vsp.TableBorder = tbNone
      Vsp.FontSize = 14
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 5
      Vsp.TableCell(tcRows) = 4
       Vsp.TableCell(tcRowHeight, 1) = 1 * 627
       Vsp.CurrentY = Vsp.CurrentY + 1530

         For I = 1 To Vsp.TableCell(tcRows)
            Vsp.TableCell(tcColWidth, I, 1) = ".2cm"
            Vsp.TableCell(tcColWidth, I, 2) = "2.8cm"
            Vsp.TableCell(tcColWidth, I, 3) = "7.7cm"
            Vsp.TableCell(tcColWidth, I, 4) = "2.9cm"
            Vsp.TableCell(tcColWidth, I, 5) = "4cm"
             
             Vsp.TableCell(tcAlign, I, 1) = taLeftBottom
             Vsp.TableCell(tcAlign, I, 2) = taLeftBottom
             Vsp.TableCell(tcAlign, I, 3) = taLeftBottom
             Vsp.TableCell(tcAlign, I, 4) = taRightBottom
             Vsp.TableCell(tcAlign, I, 5) = taRightBottom
             For J = 1 To 20
                  If J = 13 And Val(Tot(J)) > 0 And TempTot(J) = 0 Then
                     Vsp.TableCell(tcText, I, 2) = "570-5007"
                     Vsp.TableCell(tcText, I, 3) = "ค่าขนส่งอาหาร"
                     Vsp.TableCell(tcText, I, 4) = FormatNumber(Tot(J))
                     Vsp.TableCell(tcText, I, 5) = ""
                     TempTot(J) = -1
                     TempTot2(4) = TempTot2(4) + Tot(J)
                     Exit For
                  ElseIf J = 16 And Val(Tot(J)) > 0 And TempTot(J) = 0 Then
                     Vsp.TableCell(tcText, I, 2) = "411-2006"
                     Vsp.TableCell(tcText, I, 3) = "รายได้จากการขายอาหาร-บุคคลอื่น(เก็บการตลาด)"
                     Vsp.TableCell(tcText, I, 4) = FormatNumber(Tot(J))
                     Vsp.TableCell(tcText, I, 5) = ""
                     TempTot(J) = -1
                     TempTot2(4) = TempTot2(4) + Tot(J)
                     Exit For
                   ElseIf J = 17 And Val(Tot(J)) > 0 And TempTot(J) = 0 Then
                     Vsp.TableCell(tcText, I, 2) = ""
                     Vsp.TableCell(tcText, I, 3) = "ขนส่งเก็บโรงงาน"
                     Vsp.TableCell(tcText, I, 4) = FormatNumber(Tot(J))
                     Vsp.TableCell(tcText, I, 5) = ""
                     TempTot(J) = -1
                     TempTot2(4) = TempTot2(4) + Tot(J)
                      Exit For
                     ElseIf J = 19 And Val(Tot(J)) > 0 And TempTot(J) = 0 Then
                        Vsp.TableCell(tcText, I, 2) = "610-4020"
                        Vsp.TableCell(tcText, I, 3) = "รายจ่ายอื่นๆ"
                        Vsp.TableCell(tcText, I, 4) = FormatNumber(Tot(J))
                        Vsp.TableCell(tcText, I, 5) = ""
                        TempTot(J) = -1
                        TempTot2(4) = TempTot2(4) + (Tot(J))
                         Exit For
                     ElseIf J = 20 And Val(Tot(J)) > 0 And TempTot(J) = 0 Then
                        Vsp.TableCell(tcText, I, 2) = "570-5007"
                        Vsp.TableCell(tcText, I, 3) = "ค่าขนส่งอาหาร(เพิ่ม)"
                        Vsp.TableCell(tcText, I, 4) = FormatNumber(Tot(J))
                        Vsp.TableCell(tcText, I, 5) = ""
                        TempTot(J) = -1
                        TempTot2(4) = TempTot2(4) + (Tot(J))
                         Exit For
                  End If
             Next J
       Next I
       Vsp.EndTable
       
       Vsp.StartTable
      Vsp.TableCell(tcCols) = 5
      Vsp.TableCell(tcRows) = 4
       Vsp.TableCell(tcRowHeight, 1) = 1 * 597 '567
       Vsp.CurrentY = Vsp.CurrentY - 140

         For I = 1 To Vsp.TableCell(tcRows)
            Vsp.TableCell(tcColWidth, I, 1) = ".5cm"
            Vsp.TableCell(tcColWidth, I, 2) = "3cm"
            Vsp.TableCell(tcColWidth, I, 3) = "7.4cm"
            Vsp.TableCell(tcColWidth, I, 4) = "3cm"
            Vsp.TableCell(tcColWidth, I, 5) = "3.9cm"

             Vsp.TableCell(tcAlign, I, 1) = taLeftBottom
             Vsp.TableCell(tcAlign, I, 2) = taLeftBottom
             Vsp.TableCell(tcAlign, I, 3) = taLeftBottom
             Vsp.TableCell(tcAlign, I, 4) = taRightBottom
             Vsp.TableCell(tcAlign, I, 5) = taRightBottom
             For J = 1 To 21
                  If J = 15 And Val(Tot(J)) > 0 And TempTot(J) = -1 Then
                     Vsp.TableCell(tcText, I, 2) = "413-2006"
                     Vsp.TableCell(tcText, I, 3) = "รายได้ค่าขนส่งบุคคลอื่นๆ(เก็บการตลาด)"
                     Vsp.TableCell(tcText, I, 4) = ""
                     Vsp.TableCell(tcText, I, 5) = FormatNumber(Tot(J))
                     TempTot(J) = -2
'                     TempTot2(5) = TempTot2(5) + Tot(J)
                      Exit For
                  ElseIf J = 5 And Val(Tot(J)) > 0 And TempTot(J) = 0 Then
                     Vsp.TableCell(tcText, I, 2) = "420-9999"
                     Vsp.TableCell(tcText, I, 3) = "รายได้อื่นๆ(ค่าดำเนินการ)"
                     Vsp.TableCell(tcText, I, 4) = ""
                     Vsp.TableCell(tcText, I, 5) = FormatNumber(Tot(J) * 60)
                     TempTot(J) = -1
                      Exit For
                  ElseIf J = 16 And Val(Tot(J)) > 0 And TempTot(J) = -1 Then
'                     Vsp.TableCell(tcText, I, 2) = ""
'                     Vsp.TableCell(tcText, I, 3) = "ขนส่งเก็บโรงงาน"
                     Vsp.TableCell(tcText, I, 2) = "413-2006"
                     Vsp.TableCell(tcText, I, 3) = "รายได้ค่าขนส่งบุคคลอื่น"
                     Vsp.TableCell(tcText, I, 4) = ""
                     Vsp.TableCell(tcText, I, 5) = FormatNumber(Tot(J))
                     TempTot(J) = -2
                      Exit For
                  ElseIf J = 17 And Val(Tot(J)) > 0 And TempTot(J) = 0 Then
                     Vsp.TableCell(tcText, I, 2) = "420-9999"
                     Vsp.TableCell(tcText, I, 3) = "รายได้อื่นๆ"
                     Vsp.TableCell(tcText, I, 4) = ""
                     Vsp.TableCell(tcText, I, 5) = FormatNumber(Tot(J))
                     TempTot(J) = -1
                      Exit For
                  ElseIf J = 18 And Val(Tot(J)) > 0 And TempTot(J) = 0 Then
                     Vsp.TableCell(tcText, I, 2) = "570-5007"
                     Vsp.TableCell(tcText, I, 3) = "ค่าขนส่งอาหาร(ลด)"
                     Vsp.TableCell(tcText, I, 4) = ""
                     Vsp.TableCell(tcText, I, 5) = FormatNumber(Tot(J))
                     TempTot(J) = -1
                      Exit For
                  ElseIf J = 21 And TempTot(J) = 0 Then
                     Vsp.TableCell(tcText, I, 2) = "217-3999"
                     Vsp.TableCell(tcText, I, 3) = "ค่าใช้จ่ายค้างจ่าย"
                     Vsp.TableCell(tcText, I, 4) = ""
                     Vsp.TableCell(tcText, I, 5) = FormatNumber(NetTotal + Tot3(13) + AddTrnDriver - ReduceTrnDriver) '+ AddTrnDriver
                     TempTot(J) = -1
                      Exit For
                  End If
             Next J
       Next I
       Vsp.EndTable
       
      
      Vsp.CurrentY = 12500 'Vsp.CurrentY + 100
      strFormat = VSP_CalTable("^4|<20;", mdblWid, alngX)
      strDes1 = "ค่าขนส่งอาหารระหว่างวันที่  " & DateToStringExtEx2(mcolParam("FROM_DATE")) & " - " & DateToStringExtEx2(mcolParam("TO_DATE")) & " (" & Sp.SUPPLIER_NAME & ")"
      strPrint = "|" & strDes1 & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
            
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
   
   If DocFlag = 1 Then
      If mcolParam("INVOICE_FLAG") = 1 Then
      Vsp.FontSize = 8
      dY(0) = Vsp.CurrentY
      Vsp.TextAlign = taLeftBottom
      strFormat = VSP_CalTable(TITLE_FORMAT, mdblWid, alngX)
      Vsp.TableBorder = tbNone
      Vsp.FontBold = False
      Call Vsp.AddTable(strFormat, "", "|หน้าที่|: " & Vsp.CurrentPage)
      Call Vsp.AddTable(strFormat, "", "|วันที่พิมพ์|: " & DateToStringExt(mdteDate))
      Call Vsp.AddTable(strFormat, "", "|เวลา|: " & Right$(DateToStringExtEx(mdteDate), 8))
      Call Vsp.AddTable(strFormat, "", "|ผู้ใช้|: " & glbUser.USER_NAME)
      Vsp.CurrentY = dY(0)
      
      Vsp.FontSize = TITLE_SIZE
      Vsp.FontBold = True
      Vsp.TextAlign = taCenterBottom
      Vsp.Paragraph = MODULE_DESC
      
      Vsp.TextAlign = taLeftBottom
      Vsp.Paragraph = "ส่งโดย " & Sp.SUPPLIER_NAME & " (" & Sp.SUPPLIER_CODE & ")"

      End If
      If mblnHeader Then
         Call printHeader
      End If
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



