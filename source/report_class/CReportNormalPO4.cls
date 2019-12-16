VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportNormalPO4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportNormalPO4"
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

Private Const TITLE_SIZE = 24
Private Const HEADER_SIZE = 16
Private Const HEADER_SIZE2 = 13
Private Const ADDRESS_SIZE = 15
Private Const DETAIL_SIZE = 15
Private Const DETAIL_SIZE2 = 11
Private Const APPROVE_SIZE = 15
Private Const FOOTER_SIZE = 12

Private Const TITLE_FORMAT = "<32|<3|<5;"
Private Const HEADER_FORMAT = "^1|^7|^2|^2|^2;"
Private Const HEADER_FORMAT2 = "^1|^2|^2|^3|^2|^2|^2;"
Private Const HEADER_FORMAT3 = "^1|^7|^2|^2|^2;"
Private Const DETAIL_FORMAT = "^1|<7|>2|>2|>2;"
Private Const DETAIL_FORMAT2 = "^1|^2|^2|<3|>2|>2|>2;"
Private Const DETAIL_FORMAT3 = "^1|<7|>2|>2|>2;"
Private Const FOOTER_FORMAT = ">9|>2|>2;"
Private Const FOOTER_FORMAT2 = "^8|>2|>2|>2;"

Private m_Enterprise As CEnterprise
Private m_CustAddress As CAddress
Private m_EnpAddress As CAddress
Private m_EnpPerson As CName
Private Sp As CSystemParam
Private m_OwnerAddress As CAddress
Private m_cData As CBillingDoc
Private m_Addresses As Collection

Private X1 As Long
Private X2  As Long
Private Y1  As Long
Private Y2  As Long
Private AllEditEndPage As Boolean

Private collUserAccount As Collection

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_Enterprise = New CEnterprise
   Set m_CustAddress = New CAddress
   Set m_EnpAddress = New CAddress
   Set m_EnpPerson = New CName
   Set m_OwnerAddress = New CAddress
   Set m_Addresses = New Collection
   Set collUserAccount = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_Enterprise = Nothing
   Set m_CustAddress = Nothing
   Set m_EnpAddress = Nothing
   Set m_EnpPerson = Nothing
   Set m_OwnerAddress = Nothing
   Set m_Addresses = Nothing
   Set collUserAccount = Nothing
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

Private Sub PrintPOFooter(alngX() As Long)
Dim strFormat As String
Dim strPrint As String
Dim tempUserAccount As CUserAccount
Dim createRealName As String
Dim approveRealName As String

   Vsp.TableBorder = tbNone
   Vsp.FontBold = False
   
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Set tempUserAccount = GetObject("CUserAccount", collUserAccount, Trim(m_cData.CREATE_NAME))
   createRealName = tempUserAccount.REAL_NAME
   If Len(createRealName) > 0 Then
      createRealName = "(  " & createRealName & "  )"
   End If
   Set tempUserAccount = GetObject("CUserAccount", collUserAccount, Trim(m_cData.APPROVE_NAME))
   approveRealName = tempUserAccount.REAL_NAME
   If Len(approveRealName) > 0 Then
      approveRealName = "( " & approveRealName & " )"
   End If
   
  Vsp.FontSize = APPROVE_SIZE
   Vsp.FontBold = False
   strFormat = VSP_CalTable("^15|^15;", mdblWid, alngX)
   strPrint = createRealName & "|" & approveRealName & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)

   Vsp.FontSize = FOOTER_SIZE
   Vsp.FontBold = True
  strFormat = VSP_CalTable("^15|^15;", mdblWid, alngX)
   strPrint = " ผู้ขอสั่งซื้อ|ผู้อนุมัติสั่งซื้อ;"
   Call Vsp.AddTable(strFormat, "", strPrint)
End Sub

Private Sub PrintPOHeader(alngX() As Long)
Dim strFormat As String
Dim strPrint As String
Static I As Long
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim OwnerFlag As String
Dim TempStr As String
Dim TempOwnerAddress As String

If I = 0 Then
   I = 1
   Set TempRs = New ADODB.Recordset
   m_Enterprise.ENTERPRISE_ID = -1
   Call m_Enterprise.QueryData(TempRs, iCount)
      If Not TempRs.EOF Then
         Call m_Enterprise.PopulateFromRS(1, TempRs)
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
   If m_Addresses.Count >= 0 Then
      Set m_EnpAddress = m_Addresses(1)
   Else
      Set m_EnpAddress = New CAddress
   End If
m_CustAddress.ADDRESS_ID = -1
m_CustAddress.CUSTOMER_ID = m_cData.CUSTOMER_ID
Call m_CustAddress.QueryData3(TempRs, iCount)
   If Not TempRs.EOF Then
      Call m_CustAddress.PopulateFromRS(TempRs)
   End If
m_EnpPerson.NAME_ID = -1
m_EnpPerson.ENTERPRISE_ID = -1
m_EnpPerson.MASTER_FLAG = "Y"
Call m_EnpPerson.QueryData2(TempRs, iCount)
   If Not TempRs.EOF Then
      Call m_EnpPerson.PopulateFromRS2(TempRs)
   End If

TempOwnerAddress = ""
   If m_EnpPerson.ADDRESS_ID > 0 Then
      m_OwnerAddress.ADDRESS_ID = m_EnpPerson.ADDRESS_ID
      m_OwnerAddress.CUSTOMER_ID = -1
      Call m_OwnerAddress.QueryData(TempRs, iCount)
      If Not TempRs.EOF Then
         Call m_OwnerAddress.PopulateFromRS(TempRs)
         TempOwnerAddress = vbCrLf & m_OwnerAddress.PackAddress
      End If
   End If
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
Set TempRs = Nothing
End If
   
   Vsp.TableBorder = tbNone
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = m_Enterprise.ENTERPRISE_NAME & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)

   Vsp.FontSize = ADDRESS_SIZE
   Vsp.FontBold = False
   
   strFormat = VSP_CalTable("<12|>6|<3;", mdblWid, alngX)
   strPrint = MapText("ที่อยู่") & "  " & m_EnpAddress.PackAddressEx & "|" & MapText("วันที่เอกสาร ") & "|" & DateToStringExtEx2(m_cData.DOCUMENT_DATE) & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)

   If OwnerFlag = "Y" Then
      TempStr = m_EnpPerson.LONG_NAME & "  " & m_EnpPerson.LAST_NAME
      If Trim(TempStr) <> "" Then
         TempStr = TempStr & ", "
      End If
   End If
   
   strFormat = VSP_CalTable("<12|>6|<3;", mdblWid, alngX)
   strPrint = MapText("เลขประจำตัวผู้เสียภาษี ") & "  " & m_Enterprise.TAX_ID & TempOwnerAddress & "|" & MapText("เลขที่เอกสาร") & "|" & m_cData.DOCUMENT_NO & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("<12|>6|<3;", mdblWid, alngX)
   If m_cData.DOCUMENT_TYPE = 1000 Then
      strPrint = MapText("ชื่อซัพพลายเออร์") & "  " & m_cData.SUPPLIER_NAME & "|" & MapText("รหัสซัพพลายเออร์") & "|" & m_cData.SUPPLIER_CODE & ";"
   Else
      strPrint = MapText("ชื่อซัพพลายเออร์") & "  " & m_cData.SUPPLIER_NAME & "(" & m_cData.SUPPLIER_CODE & ")|" & MapText("เลขที่ PR") & "|" & m_cData.PR_NO & ";"
   End If
   Call Vsp.AddTable(strFormat, "", strPrint)

   strFormat = VSP_CalTable("<12|>6|<3;", mdblWid, alngX)
   strPrint = MapText("กำหนดวันส่งมอบ") & "  " & DateToStringExtEx2(m_cData.DUE_DATE) & "|" & MapText("เครดิต") & "|" & m_cData.Credit & " วัน ;"
   Call Vsp.AddTable(strFormat, "", strPrint)

   strFormat = VSP_CalTable("<12|>6|<3;", mdblWid, alngX)
   strPrint = MapText("เงื่อนไขหลังรับวัสดุ") & "  " & m_cData.CONDITION_NAME & "|" & MapText("การชำระ") & "|" & m_cData.PAID_TYPE_NAME & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
End Sub

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
   
   If mblnHeader Then
      Call PrintPOHeader(alngX)
      mblnHeader = False
   End If
   
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
  Vsp.TableBorder = tbAll
   strFormat = VSP_CalTable(HEADER_FORMAT, mdblWid, alngX)
   strPrint = "ลำดับ|รายการ|จำนวน|ราคาต่อหน่วย|มูลค่า;"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
  Vsp.TableBorder = tbdT
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
     
   X1 = Vsp.PageWidth - Vsp.MarginRight
   Y1 = Vsp.CurrentY
End Sub
Private Sub printHeader2()
   Dim strFormat As String
   Dim strPrint As String
   Dim tbdT As TableBorderSettings
   Dim blnBold As Boolean
   Dim iSize As Integer
   Dim alngX() As Long
   
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   'Call PrintPOHeader(alngX)
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
  Vsp.TableBorder = tbAll
   strFormat = VSP_CalTable(HEADER_FORMAT2, mdblWid, alngX)
   strPrint = "ลำดับ|วันที่รับของ|เลขที่รับของ|รายการ|จำนวน|@|มูลค่า;"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
  Vsp.TableBorder = tbdT
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
     
   X1 = Vsp.PageWidth - Vsp.MarginRight
   Y1 = Vsp.CurrentY
End Sub
Private Sub printHeader3()
   Dim strFormat As String
   Dim strPrint As String
   Dim tbdT As TableBorderSettings
   Dim blnBold As Boolean
   Dim iSize As Integer
   Dim alngX() As Long
   
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   'Call PrintPOHeader(alngX)
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
  Vsp.TableBorder = tbAll
   strFormat = VSP_CalTable(HEADER_FORMAT3, mdblWid, alngX)
   strPrint = "ลำดับ|รายการ|คงเหลือจำนวน|ราคาต่อหน่วย|มูลค่า;"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
  Vsp.TableBorder = tbdT
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
     
   X1 = Vsp.PageWidth - Vsp.MarginRight
   Y1 = Vsp.CurrentY
End Sub
Private Function initDoc() As Boolean
   Dim strDate As String
   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   
   'Vsp.PaperSize = 37
'   Dim I As Long
'   ''Debug.Print I & "------>" & Vsp.PageWidth & " " & Vsp.PageHeight
   
'   For I = 1 To 256
'      Vsp.PaperSize = I
'      'If (Vsp.PageWidth <= 6000 Or Vsp.PageHeight <= 6000)  Then
'      If Vsp.PaperSize = I Then                                     'เช็คว่าขนาดไหนบ้างที่ PRINTER รองรับ
'         ''Debug.Print I & "------>" & FormatNumber(Vsp.PageWidth / 566.8, 0) & "X" & FormatNumber(Vsp.PageHeight / 566.8)
'      End If
''      End If
'   Next I
  'vsp.PaperSize = 11
   
   Vsp.PaperSize = pprA4
'   Vsp.ORIENTATION = orPortrait
'   Vsp.MarginBottom = 1440
'   Vsp.MarginFooter = 1440
'   Vsp.MarginHeader = 1440
'   Vsp.MarginLeft = 700
'   Vsp.MarginRight = 700
'   Vsp.MarginTop = 1440
   
   Call LoadUserAccountByName(Nothing, collUserAccount)
   
   Vsp.ORIENTATION = orPortrait
   Vsp.MarginBottom = 200
   Vsp.MarginFooter = 500
   Vsp.MarginHeader = 500
   Vsp.MarginLeft = 500
   Vsp.MarginRight = 500
   Vsp.MarginTop = 600
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
     Vsp.PenStyle = psDash
   MODULE_DESC = MapText("ใบสั่งซื้อ(PO)")
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
Dim iCount As Long
Dim I As Long
Dim strFormat As String
Dim strPrint As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Poi As CDoItem
Dim Sup As CSupItem
Dim Sup2 As CSupItem
Dim Sum1 As Double
Dim Sum2 As Double
Dim Sum3 As Double
Dim Total(10) As Double
Dim SumDiscount As Double
Dim SumDeposit As Double

Dim SumTxAmountTotal As Double 'Sum ค่า TX_AMOUNT ของรายการรับเหลือทั้งหมด
Dim SumTxAmountEachPart As Double 'Sum ค่า TX_AMOUNT ของรายการรับ
Dim Tmp As Double
Dim Tmp2 As Double
Dim SumUnitPriceSupReceive As Double 'Sum มูลค่าของรายการรับ
Dim SumPriceAmountTotal As Double 'ยอดมูลค่าคงเหลือทั้งหมด
   
   RName = "genDoc"
   
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   Set Rs = New ADODB.Recordset
   Set m_cData = New CBillingDoc
   m_cData.BILLING_DOC_ID = mcolParam("BILLING_DOC_ID")
   m_cData.QueryFlag = 1
   m_cData.PrintReportNormalPO2 = True
   If Not glbDaily.QueryBillingDoc(m_cData, Rs, iCount, IsOK, glbErrorLog) Then
      mstrErrMsg = "Query Error."
      Exit Function
   End If
   Call m_cData.PopulateFromRS(1, Rs)
   
   If (Rs Is Nothing) Or (iCount < 1) Then
      mstrErrMsg = "ไม่พบข้อมูลที่ต้องการในฐานข้อมูล"
      Exit Function
   End If
   
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
   Sum1 = 0
   Sum2 = 0
   Sum3 = 0
   SumDiscount = 0
   SumDeposit = 0
   
'   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
'   Call printHeader
   mblnHeader = True
   
   AllEditEndPage = True
   strFormat = VSP_CalTable(DETAIL_FORMAT, mdblWid, alngX)
   Vsp.TableBorder = tbAll
   Vsp.TableBorder = tbColumns
      
      I = 0
       For Each Sup In m_cData.SupItems
         I = I + 1
         strPrint = I & "|"
         If Len(Sup.PART_SUPPLIER_NAME) > 0 Then
            strPrint = strPrint & Sup.PART_SUPPLIER_NAME
         Else
            strPrint = strPrint & Sup.PART_DESC
         End If
         
         If Len(Sup.PART_NOTE) > 0 Then
            strPrint = strPrint & " (" & Sup.PART_NOTE & ")" & "|"
         Else
            strPrint = strPrint & "|"
         End If
         
         strPrint = strPrint & FormatNumber(Sup.TX_AMOUNT) & "|"
'         If Sup.SUPPLIER_CODE = "อ-0012" Or Sup.SUPPLIER_CODE = "ค-1051" Then
'            strPrint = strPrint & FormatNumber(Sup.ACTUAL_UNIT_PRICE, 3) & "|"
'         Else
'            strPrint = strPrint & FormatNumber(Sup.ACTUAL_UNIT_PRICE) & "|"
'         End If
         strPrint = strPrint & FormatNumber(Sup.ACTUAL_UNIT_PRICE) & "|"
         strPrint = strPrint & FormatNumber(Sup.TX_AMOUNT * Sup.ACTUAL_UNIT_PRICE) & ";"
         Sum2 = Sum2 + Sup.TX_AMOUNT
         Sum3 = Sum3 + (Sup.TX_AMOUNT * Sup.ACTUAL_UNIT_PRICE)
         
         If I >= m_cData.DoItems.Count Then
            Call AdjustPage(Vsp, strFormat, strPrint, 4000)
         End If
         Call Vsp.AddTable(strFormat, "", strPrint)
      Next Sup
   
      X2 = Vsp.PageWidth - Vsp.MarginRight
      Y2 = Vsp.CurrentY
      
      Call Vsp.DrawLine(Vsp.MarginLeft, Y1, Vsp.MarginLeft, Y2)
      Call Vsp.DrawLine(X1, Y1, X2, Y2)
      Call Vsp.DrawLine(Vsp.MarginLeft, Vsp.CurrentY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY)
      Vsp.TableBorder = tbNone
      Y1 = Vsp.CurrentY
      AllEditEndPage = False
      
      strFormat = VSP_CalTable(FOOTER_FORMAT, mdblWid, alngX)
      Vsp.FontBold = True
      strPrint = "|รวม|" & FormatNumber(Sum3) & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)

      If m_cData.DELIVERY_FEE > 0 Then
      strFormat = VSP_CalTable(FOOTER_FORMAT, mdblWid, alngX)
      strPrint = "|ภาษี 7 %|" & FormatNumber(m_cData.DELIVERY_FEE) & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
      End If
      
      strFormat = VSP_CalTable(FOOTER_FORMAT, mdblWid, alngX)
      strPrint = "(" & ThaiBaht(Sum3 + m_cData.DELIVERY_FEE) & ")|รวมสุทธิ|" & FormatNumber(Sum3 + m_cData.DELIVERY_FEE) & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
      Vsp.FontBold = False
         
      Call Vsp.DrawLine(Vsp.MarginLeft, Y1, Vsp.MarginLeft, Vsp.CurrentY)
      Call Vsp.DrawLine(Vsp.PageWidth - Vsp.MarginRight, Y1, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY)
      Call Vsp.DrawLine(Vsp.MarginLeft, Vsp.CurrentY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY)
      
        If Len(m_cData.NOTE) > 0 Then
          Vsp = "หมายเหตุ : " & m_cData.NOTE
       End If
       
       Vsp.TableBorder = tbAll
       Vsp.TableBorder = tbColumns
       Vsp.FontBold = True
       Vsp.TextAlign = taLeftMiddle
       Vsp = "ยอดรับ"
       Call printHeader2
       strFormat = VSP_CalTable(DETAIL_FORMAT2, mdblWid, alngX)
       Vsp.FontBold = False
        For I = 0 To UBound(Total)
          Total(I) = 0
        Next I
        I = 0
        Dim Temp  As Double
       For Each Sup2 In m_cData.SupReceiptItems
         I = I + 1
         strPrint = I & "|"
         strPrint = strPrint & DateToStringExtEx2(Sup2.DOCUMENT_DATE) & "|"
         strPrint = strPrint & Sup2.DOCUMENT_NO & "|"
         strPrint = strPrint & Left(Sup2.PART_DESC, 22) & "|"
         If Sup2.CALCULATE_TYPE = 0 Then
            strPrint = strPrint & FormatNumber(Sup2.TX_AMOUNT) & "|"
            Temp = Sup2.TX_AMOUNT
         ElseIf Sup2.CALCULATE_TYPE = 1 Then
            strPrint = strPrint & FormatNumber(Sup2.SUPPLIER_WEIGHT) & "|"
            Temp = Sup2.SUPPLIER_WEIGHT
         ElseIf Sup2.CALCULATE_TYPE = 2 Then
            strPrint = strPrint & FormatNumber(Sup2.WEIGHT_AMOUNT) & "|"
            Temp = Sup2.WEIGHT_AMOUNT
         ElseIf Sup2.CALCULATE_TYPE = 3 Then
            strPrint = strPrint & FormatNumber(Sup2.TOTAL_WEIGHT) & "|"
            Temp = Sup2.TOTAL_WEIGHT
         End If
         
         strPrint = strPrint & FormatNumber(Sup2.ACTUAL_UNIT_PRICE) & "|"
         strPrint = strPrint & FormatNumber(Sup2.TOTAL_ACTUAL_PRICE) & ";"
         Total(4) = Total(4) + Temp
         Total(6) = Total(6) + Sup2.TOTAL_ACTUAL_PRICE
         
         Call Vsp.AddTable(strFormat, "", strPrint)
         Vsp.FontBold = False
      Next Sup2
      Call Vsp.DrawLine(Vsp.MarginLeft, Y1, Vsp.MarginLeft, Vsp.CurrentY)
      Call Vsp.DrawLine(Vsp.PageWidth - Vsp.MarginRight, Y1, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY)
      Call Vsp.DrawLine(Vsp.MarginLeft, Vsp.CurrentY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY)
      
      strFormat = VSP_CalTable(FOOTER_FORMAT2, mdblWid, alngX)
      Vsp.TableBorder = tbNone
      Vsp.TableBorder = tbNone
      Vsp.FontBold = True
      strPrint = "รวมยอดรับ|" & FormatNumber(Total(4)) & "|" & FormatNumber(MyDiffEx(Total(6), Total(4)), 2) & "|" & FormatNumber(Total(6)) & "|;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      Vsp.FontBold = False
      
      Call Vsp.DrawLine(Vsp.MarginLeft, Y1, Vsp.MarginLeft, Vsp.CurrentY)
      Call Vsp.DrawLine(Vsp.PageWidth - Vsp.MarginRight, Y1, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY)
      Call Vsp.DrawLine(Vsp.MarginLeft, Vsp.CurrentY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY)
      
      Vsp.TableBorder = tbAll
      Vsp.TableBorder = tbColumns
      Vsp.FontBold = True
      Vsp.TextAlign = taLeftMiddle
      Vsp = "ยอดคงเหลือ"
      Call printHeader3
      Vsp.FontSize = DETAIL_SIZE
      strFormat = VSP_CalTable(DETAIL_FORMAT3, mdblWid, alngX)
      Vsp.FontBold = False
      
      I = 0

   
      SumTxAmountTotal = 0
      SumTxAmountEachPart = 0
      SumUnitPriceSupReceive = 0
      SumPriceAmountTotal = 0
      For Each Sup In m_cData.SupItems
      I = I + 1
      SumTxAmountEachPart = 0
      SumUnitPriceSupReceive = 0
      Temp = 0
         For Each Sup2 In m_cData.SupReceiptItems
            If Sup.PART_NO = Sup2.PART_NO Then
               If Sup2.CALCULATE_TYPE = 0 Then
                  Temp = Sup2.TX_AMOUNT
               ElseIf Sup2.CALCULATE_TYPE = 1 Then
                  Temp = Sup2.SUPPLIER_WEIGHT
               ElseIf Sup2.CALCULATE_TYPE = 2 Then
                  Temp = Sup2.WEIGHT_AMOUNT
               ElseIf Sup2.CALCULATE_TYPE = 3 Then
                  Temp = Sup2.TOTAL_WEIGHT
               End If
               SumTxAmountEachPart = SumTxAmountEachPart + Temp
            End If
         Next Sup2
         Tmp = Sup.TX_AMOUNT - SumTxAmountEachPart
         SumTxAmountTotal = SumTxAmountTotal + Tmp

         SumPriceAmountTotal = SumPriceAmountTotal + (Tmp * Sup.ACTUAL_UNIT_PRICE)
      
         strPrint = I & "|"
         strPrint = strPrint & Left(Sup.PART_DESC, 100) & "|"
         strPrint = strPrint & IIf(Tmp < 0, "(เกิน)" & FormatNumber(Abs(Tmp), 2), FormatNumber(Tmp, 2)) & "|"
         strPrint = strPrint & FormatNumber(Sup.ACTUAL_UNIT_PRICE, 2) & "|"
         strPrint = strPrint & IIf(Tmp * Sup.ACTUAL_UNIT_PRICE < 0, "(เกิน)" & FormatNumber(Abs(Tmp * Sup.ACTUAL_UNIT_PRICE), 2), FormatNumber(Tmp * Sup.ACTUAL_UNIT_PRICE, 2)) & ";"
         Call Vsp.AddTable(strFormat, "", strPrint)
         Vsp.FontBold = False
      Next Sup
      
      Call Vsp.DrawLine(Vsp.MarginLeft, Y1, Vsp.MarginLeft, Vsp.CurrentY)
      Call Vsp.DrawLine(Vsp.PageWidth - Vsp.MarginRight, Y1, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY)
      Call Vsp.DrawLine(Vsp.MarginLeft, Vsp.CurrentY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY)
      Vsp.FontSize = DETAIL_SIZE
      strFormat = VSP_CalTable(FOOTER_FORMAT2, mdblWid, alngX)
     Vsp.TableBorder = tbNone
     Vsp.TableBorder = tbNone
      Vsp.FontBold = True '
      strPrint = "รวมยอดคงเหลือทั้งหมด|"
      strPrint = strPrint & IIf(SumTxAmountTotal < 0, "(เกิน )" & FormatNumber(Abs(SumTxAmountTotal)), FormatNumber(SumTxAmountTotal)) & "|"
      strPrint = strPrint & FormatNumber(MyDiffEx(SumPriceAmountTotal, SumTxAmountTotal), 2) & "|"
      strPrint = strPrint & IIf((SumPriceAmountTotal) < 0, "(เกิน )" & FormatNumber(Abs(SumPriceAmountTotal)), FormatNumber(SumPriceAmountTotal)) & "|;"
      Call Vsp.AddTable(strFormat, "", strPrint)
       
      Call Vsp.DrawLine(Vsp.MarginLeft, Y1, Vsp.MarginLeft, Vsp.CurrentY)
      Call Vsp.DrawLine(Vsp.PageWidth - Vsp.MarginRight, Y1, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY)
      Call Vsp.DrawLine(Vsp.MarginLeft, Vsp.CurrentY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY)
      
      Call PrintPOFooter(alngX)
      
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
    Vsp.FontBold = False
   Vsp.TextAlign = taRightBottom
'  If m_cData.DOCUMENT_TYPE = 1000 Then 'การเปลี่ยนแปลงรูปแบบเอกสารแต่ล่ะครั้งให้เปลี่ยนเวอร์ชั่น ISO ของเอกสารด้วยทุกครั้ง
'       Vsp = "FM-PC-02 Rev.01(07/04/59)"
'  ElseIf m_cData.DOCUMENT_TYPE = 1001 Or m_cData.DOCUMENT_TYPE = 1003 Then
'      Vsp = "FM-PC-03 Rev.01(07/04/59)"
'   End If
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
   
   'Vsp.DrawPicture mcolParam("PICTURE"), "0.75in", "0.75in", "100%", "100%"
   
   Vsp.FontSize = 8
   dY(0) = Vsp.CurrentY
   Vsp.TextAlign = taLeftBottom
   strFormat = VSP_CalTable(TITLE_FORMAT, mdblWid, alngX)
   Vsp.TableBorder = tbNone
   Call Vsp.AddTable(strFormat, "", "|หน้าที่|: " & Vsp.CurrentPage)
'   Call Vsp.AddTable(strFormat, "", "|วันที่พิมพ์|: " & DateToStringExt(mdteDate))
'   Call Vsp.AddTable(strFormat, "", "|เวลา|: " & Right$(DateToStringExtEx(mdteDate), 8))
   Vsp.CurrentY = dY(0)
   
   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   Vsp.TextAlign = taCenterBottom
   Vsp.Paragraph = MODULE_DESC
  ' If mblnHeader Then
      Call printHeader
   'End If
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


