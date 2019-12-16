VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAgePartWh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
''User Group Report
'Implements CReportInterface
'
'Private Const MODULE_NAME = "CReportInventoryWh"
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
'Public m_TempCol As Collection
'Private m_HeaderFormat1 As String
'Private m_HeaderFormat2 As String
'Private m_HeaderText1 As String
'Private m_HeaderText2 As String
'
'Private m_Headers0 As CFieldList
'Private m_Headers1 As CFieldList
'Private m_Details1 As CFieldList
'Private m_Left As Double
'Private m_Suppliers As Collection
'Private m_Locations As Collection
'Private m_CollLotItemWh As Collection
'Private m_LotItemWh As CLotItemWH
'
'Private Sub Class_Initialize()
'   Set mcolParam = New Collection
'   Set m_TempCol = New Collection
'
'   Set m_Headers0 = New CFieldList
'   Set m_Headers1 = New CFieldList
'   Set m_Details1 = New CFieldList
'   Set m_Suppliers = New Collection
'   Set m_Locations = New Collection
'
'
'   Set m_CollLotItemWh = New Collection
'   Set m_LotItemWh = New CLotItemWH
'
'   Set Rs = New ADODB.Recordset
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
'   Set m_Suppliers = Nothing
'   Set m_Locations = Nothing
'   Set m_CollLotItemWh = Nothing
'   Set m_LotItemWh = Nothing
'
'   If Rs.State = adStateOpen Then
'      Rs.Close
'   End If
'   Set Rs = Nothing
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
'Dim TempStr2 As String
'Dim Cf As CReportField
'
'   tbdT = Vsp.TableBorder
'   blnBold = Vsp.FontBold
'   iSize = Vsp.FontSize
'
'   Vsp.FontSize = HEADER_SIZE
'   Vsp.FontBold = True
'   Vsp.TableBorder = tbAll
'
'   Call m_Headers0.ClearField
'   Call m_Headers1.ClearField
'
'   Set Cf = New CReportField
'
'   Call Cf.SetFieldValue(1, "^", "NO.", "^")
'   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(2.5, "^", "เบอร์อาหาร", "<")
'   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(1.5, "^", "รหัสขาย", "<")
'   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(2, "^", "ชนิด", "<")
'   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(1.5, "^", "LOT", "^")
'   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(1.5, "^", "วันผลิต", "^")
'   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(1, "^", "ถัง", "^")
'   Call m_Headers1.AddField(Cf)
'   If Val(mcolParam("DOCUMENT_TYPE")) = 14 Then
'      Call Cf.SetFieldValue(1, "^", "ล๊อค", "^")
'      Call m_Headers1.AddField(Cf)
'   End If
'   Call Cf.SetFieldValue(1, "^", "ยกมา", ">")
'   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(1, "^", "รับเข้า", ">")
'   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(1, "^", "ป.ย.", ">")
'   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(1, "^", "จ่ายออก", ">")
'   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(1, "^", "คงเหลือ", ">")
'   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(1, "^", "ยอดตัน", ">")
'   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(3, "^", "ประเภท", "<")
'   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(1, "^", "ม.ห.", "<")
'   Call m_Headers1.AddField(Cf)
'
'   Set Cf = Nothing
'
'  Call m_Headers1.GetString(1, TempStr1, TempStr2)
'   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'   strPrint = TempStr2
'   Call Vsp.AddTable(strFormat, "", strPrint)
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
'   Vsp.PaperSize = pprA4 '
'   Vsp.ORIENTATION = orLandscape
'   Vsp.MarginBottom = 1440
'   Vsp.MarginFooter = 1440
'   Vsp.MarginHeader = 1440
'   Vsp.MarginLeft = 700
'   Vsp.MarginRight = 700
'   Vsp.MarginTop = 1440
'   Vsp.FontName = "AngsanaUPC"
'   Vsp.FontSize = DETAIL_SIZE
'
'   FromDate = "วันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
'   'ToDate = "ถึง " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")
'
'   MODULE_DESC = vbCrLf & glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & vbCrLf & _
'                                      "รายงานสินค้าคงเหลือ" & vbCrLf & FromDate
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
'Dim LWH As CLotItemWH
'Dim I As Long
'Dim j As Long
'Dim strFormat As String
'Dim alngX() As Long
'Dim IsOK As Boolean
'Dim HeadCf As CReportField
'Dim BodyCf As CReportField
'Dim TempStr1 As String
'Dim TempStr2 As String
'Dim Total1(100) As Double
'Dim iCount As Long
'Dim TempStr As String
'Dim Amt As Double
'Dim Sum1 As Double
'Dim Count1 As Double
'Dim DOCUMENT_TYPE As Long
'
'
'   RName = "genDoc"
''-----------------------------------------------------------------------------------------------------
''                                             Query Here
''-----------------------------------------------------------------------------------------------------
'
'   If Not initDoc Then Exit Function
'
'   Vsp.Preview = isPreview
''-----------------------------------------------------------------------------------------------------
''                                         Main Operation Here
''-----------------------------------------------------------------------------------------------------
'
'   For j = 1 To UBound(Total1)
'      Total1(j) = 0
'   Next j
''   Set m_TempCol = mcolParam("LOT_ITEM_WH")
'
'
''Dim IsOK As Boolean
''Dim ItemCount As Long
''Dim Temp As Long
''Dim LWH As CLotItemWH
''Dim TempLotItemWh As CLotItemWH
''Dim I As Long
''Dim Data1 As Long
''Dim Data2 As Long
''Dim Data3 As Long
''Dim Data4 As Long
''Dim Data5 As Long
''Dim Key As String
''Dim PartTypeID As Long
''
''   If Flag Then
''      Call EnableForm(Me, False)
'
''      m_LotItemWh.PART_NO = PatchWildCard(txtPartNo.Text)
''      m_LotItemWh.PART_DESC = PatchWildCard(txtPartName.Text)
''      m_LotItemWh.OrderBy = cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex))
''      m_LotItemWh.OrderType = cboOrdertype.ItemData(Minus2Zero(cboOrdertype.ListIndex))
'      DOCUMENT_TYPE = mcolParam("DOCUMENT_TYPE")
''      m_LotItemWh.DOCUMENT_DATE = uctlDateStock.ShowDate
'      m_LotItemWh.BALANCE_FLAG = "N"
'      m_LotItemWh.TX_TYPE = "I"
'
'      If DOCUMENT_TYPE = 13 Then 'bulk
''            Call InitGridBulk
'         m_LotItemWh.DOCUMENT_TYPE_SET = "(13,16)"
'      ElseIf DOCUMENT_TYPE = 14 Then 'bag
''            Call InitGridBag
'         m_LotItemWh.DOCUMENT_TYPE_SET = "(14,15)"
'      End If
'
''
'      If glbDaily.QueryLotItemWh(m_LotItemWh, m_Rs, ItemCount, IsOK, glbErrorLog) Then
'           PartTypeID = cboPartType2.ItemData(Minus2Zero(cboPartType2.ListIndex))
'            Call LoadPartStockFromLotItemWh(Nothing, m_CollLotItemWhImportBal, -1, uctlDateStock.ShowDate - 1, PartTypeID, , "I", 1, 3, DOCUMENT_TYPE) 'ยอดรับเข้ายกมา
'            Call LoadPartStockFromLotItemWh(Nothing, m_CollLotItemWhExportBal, -1, uctlDateStock.ShowDate - 1, PartTypeID, , "E", 1, 6, DOCUMENT_TYPE) 'ยอดจ่ายออกยกมา
'            'เอา m_CollLotItemWhImportBal-m_CollLotItemWhExportBal ก็จะได้ยอดยกมา
'
'            Call LoadPartStockFromLotItemWh(Nothing, m_CollLotItemWhImport, uctlDateStock.ShowDate, uctlDateStock.ShowDate, PartTypeID, , "I", 1, 3, DOCUMENT_TYPE) 'ยอดรับเข้าวันนี้
'            Call LoadPartStockFromLotItemWh(Nothing, m_CollLotItemWhExport, uctlDateStock.ShowDate, uctlDateStock.ShowDate, PartTypeID, , "E", 1, 6, DOCUMENT_TYPE) 'ยอดจ่ายออกวันนี้
''
'      I = 0
'         Set m_LotItemWh = Nothing
'         While Not m_Rs.EOF
'            I = I + 1
'             Set m_LotItemWh = New CLotItemWH
'            Call m_LotItemWh.PopulateFromRS(2, m_Rs)
''
''            Set TempLotItemWh = New CLotItemWH
''            TempLotItemWh.AddEditMode = SHOW_VIEW
''            TempLotItemWh.PART_ITEM_ID = m_LotItemWh.PART_ITEM_ID
''            TempLotItemWh.PART_NO = m_LotItemWh.PART_NO
''            TempLotItemWh.BARCODE_NO = m_LotItemWh.BARCODE_NO
''            TempLotItemWh.PART_DESC = m_LotItemWh.PART_DESC
''            TempLotItemWh.DOCUMENT_TYPE = m_LotItemWh.DOCUMENT_TYPE
''
''            Set LWH = Nothing
''            If DOCUMENT_TYPE = 14 Then 'Bag
''               Key = Trim(str(m_LotItemWh.PART_ITEM_ID) & "-" & str(m_LotItemWh.WEIGHT_PER_PACK) & "-" & str(m_LotItemWh.LOT_DOC_ID) & "-" & str(m_LotItemWh.LOT_ID) & "-" & str(m_LotItemWh.BIN_NO) & "-" & str(m_LotItemWh.LOCK_NO) & "-" & "I")
''               '''Debug.Print "Key2=" & Key
''            ElseIf DOCUMENT_TYPE = 13 Then 'Bulk
''              Key = Trim(str(m_LotItemWh.PART_ITEM_ID) & "-" & str(m_LotItemWh.LOT_ID) & "-" & "I") '& "-" & str(m_LotItemWh.LOT_DOC_ID)
''            End If
''
''            Set LWH = GetObject("CLotItemWH", m_CollLotItemWhImportBal, Key, False)
''            If Not LWH Is Nothing Then
''               Data1 = LWH.CAPACITY_AMOUNT  'รับเข้ายกมา
''            Else
''               Data1 = 0
''            End If
''
''            Set LWH = Nothing
''            If DOCUMENT_TYPE = 14 Then 'Bag
''                 Key = Trim(str(m_LotItemWh.PART_ITEM_ID) & "-" & str(m_LotItemWh.WEIGHT_PER_PACK) & "-" & str(m_LotItemWh.LOT_DOC_ID) & "-" & str(m_LotItemWh.LOT_ID) & "-" & str(m_LotItemWh.BIN_NO) & "-" & str(m_LotItemWh.LOCK_NO) & "-" & "E")
''            ElseIf DOCUMENT_TYPE = 13 Then  'Bulk
''                  Key = Trim(str(m_LotItemWh.PART_ITEM_ID) & "-" & str(m_LotItemWh.LOT_ID) & "-" & "E")
''            End If
''
''            Set LWH = GetObject("CLotItemWH", m_CollLotItemWhExportBal, Key, False)
''            If Not LWH Is Nothing Then
''               Data2 = LWH.CAPACITY_AMOUNT 'จ่ายออกยกมา
''            Else
''               Data2 = 0
''            End If
''            TempLotItemWh.BALANCE_AMOUNT = Abs(Data1 - Data2)
''
''            Set LWH = Nothing
''            If DOCUMENT_TYPE = 14 Then  'Bag
''                 Key = Trim(str(m_LotItemWh.PART_ITEM_ID) & "-" & str(m_LotItemWh.WEIGHT_PER_PACK) & "-" & str(m_LotItemWh.LOT_DOC_ID) & "-" & str(m_LotItemWh.LOT_ID) & "-" & str(m_LotItemWh.BIN_NO) & "-" & str(m_LotItemWh.LOCK_NO) & "-" & "I")
''            ElseIf DOCUMENT_TYPE = 13 Then 'Bulk
''               Key = Trim(str(m_LotItemWh.PART_ITEM_ID) & "-" & str(m_LotItemWh.LOT_ID) & "-" & "I")  '& "-" & str(m_LotItemWh.LOT_DOC_ID)
''            End If
''
''            Set LWH = GetObject("CLotItemWH", m_CollLotItemWhImport, Key, False)
''            If Not LWH Is Nothing Then
''               Data3 = LWH.CAPACITY_AMOUNT
''            Else
''               Data3 = 0
''            End If
''            TempLotItemWh.IMPORT_AMOUNT = Data3 'รับเข้า
''
''            Set LWH = Nothing
''            If DOCUMENT_TYPE = 14 Then 'Bag
''                 Key = Trim(str(m_LotItemWh.PART_ITEM_ID) & "-" & str(m_LotItemWh.WEIGHT_PER_PACK) & "-" & str(m_LotItemWh.LOT_DOC_ID) & "-" & str(m_LotItemWh.LOT_ID) & "-" & str(m_LotItemWh.BIN_NO) & "-" & str(m_LotItemWh.LOCK_NO) & "-" & "E")
''            ElseIf DOCUMENT_TYPE = 13 Then 'Bulk
''               Key = Trim(str(m_LotItemWh.PART_ITEM_ID) & "-" & str(m_LotItemWh.LOT_ID) & "-" & "E")
'''               ''Debug.Print Key
''            End If
''
''            Set LWH = GetObject("CLotItemWH", m_CollLotItemWhExport, Key, False)
''            If Not LWH Is Nothing Then
''               Data4 = LWH.CAPACITY_AMOUNT
''            Else
''               Data4 = 0
''            End If
''            TempLotItemWh.EXPORT_AMOUNT = Data4 'จ่ายออก
''
''            Data5 = ((Data1 - Data2) + Data3) - Data4
''            TempLotItemWh.ACTUAL_AMOUNT = Data5 'คงเหลือ
''            TempLotItemWh.TOTAL_WEIGHT = Data5 * m_LotItemWh.WEIGHT_PER_PACK 'ยอดตัน
''            TempLotItemWh.BILL_DESC = m_LotItemWh.BILL_DESC  'ประเภท
''            TempLotItemWh.START_DATE = m_LotItemWh.START_DATE     'วันที่ผลิต
''            TempLotItemWh.LOT_NO = m_LotItemWh.LOT_NO   'Lot
''            TempLotItemWh.BIN_NAME = m_LotItemWh.BIN_NAME   'ถังบรรจุ
''            TempLotItemWh.LOCK_NAME = m_LotItemWh.LOCK_NAME    'ล๊อค
''            TempLotItemWh.NOTE = m_LotItemWh.NOTE  'หมายเหตุ
''            TempLotItemWh.WEIGHT_PER_PACK = m_LotItemWh.WEIGHT_PER_PACK
''
''            If TempLotItemWh.ACTUAL_AMOUNT > 0 Then
''               Call m_CollLotItemWh.add(TempLotItemWh)
''            End If
'            Set TempLotItemWh = Nothing
'            m_Rs.MoveNext
'         Wend
'   End If
''   End If
'
'
'
'
'
'
'
'
'
'   Call StartExportFile(Vsp)
'   Vsp.StartDoc
'      I = 0
'
'
'      Set BodyCf = New CReportField
'      Set Li = New CLotItem
'      For Each LWH In m_TempCol
'         I = I + 1
'          j = 0
'         Call m_Details1.ClearField
'
'         j = j + 1
'         Set HeadCf = m_Headers1.Fields(j)
'         TempStr = I
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
'         Call m_Details1.AddField(BodyCf)
'
'         j = j + 1
'         Set HeadCf = m_Headers1.Fields(j)
'         TempStr = LWH.PART_NO
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
'         Call m_Details1.AddField(BodyCf)
'
'         j = j + 1
'         Set HeadCf = m_Headers1.Fields(j)
'         TempStr = LWH.BARCODE_NO
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
'         Call m_Details1.AddField(BodyCf)
'
'         j = j + 1
'         Set HeadCf = m_Headers1.Fields(j)
'         TempStr = LWH.PART_DESC
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
'         Call m_Details1.AddField(BodyCf)
'         Total1(j) = Total1(j) + Li.TOTAL_ACTUAL_PRICE
'
'         j = j + 1
'         Set HeadCf = m_Headers1.Fields(j)
'         TempStr = LWH.LOT_NO   'Lot
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
'         Call m_Details1.AddField(BodyCf)
'
'         j = j + 1
'         Set HeadCf = m_Headers1.Fields(j)
'         TempStr = LWH.START_DATE   'วันที่ผลิต
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
'         Call m_Details1.AddField(BodyCf)
'
'         j = j + 1
'         Set HeadCf = m_Headers1.Fields(j)
'         TempStr = LWH.BIN_NAME   'ถังบรรจุ
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
'         Call m_Details1.AddField(BodyCf)
'
'         If Val(mcolParam("DOCUMENT_TYPE")) = 14 Then
'             j = j + 1
'            Set HeadCf = m_Headers1.Fields(j)
'            TempStr = LWH.LOCK_NAME  'ล๊อค
'            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
'            Call m_Details1.AddField(BodyCf)
'         End If
'
'         j = j + 1
'         Set HeadCf = m_Headers1.Fields(j)
'         TempStr = FormatNumberToNull(LWH.BALANCE_AMOUNT, 0)
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
'         Call m_Details1.AddField(BodyCf)
'         Total1(j) = Total1(j) + (LWH.BALANCE_AMOUNT)
'
'         If LWH.DOCUMENT_TYPE = 15 Or LWH.DOCUMENT_TYPE = 16 Then
'            j = j + 1
'            Set HeadCf = m_Headers1.Fields(j)
'            TempStr = ""  'รับเข้า
'            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
'            Call m_Details1.AddField(BodyCf)
''            Total1(j) = Total1(j) + LWH.IMPORT_AMOUNT
'
'            j = j + 1
'            Set HeadCf = m_Headers1.Fields(j)
'            TempStr = FormatNumberToNull(LWH.IMPORT_AMOUNT, 0)  'ปรับยอด
'            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
'            Call m_Details1.AddField(BodyCf)
'            Total1(j) = Total1(j) + LWH.IMPORT_AMOUNT
'         Else
'             j = j + 1
'            Set HeadCf = m_Headers1.Fields(j)
'            TempStr = ""  'รับเข้า
'            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
'            Call m_Details1.AddField(BodyCf)
'            Total1(j) = Total1(j) + LWH.IMPORT_AMOUNT
'
'            j = j + 1
'            Set HeadCf = m_Headers1.Fields(j)
'            TempStr = ""  'ปรับยอด
'            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
'            Call m_Details1.AddField(BodyCf)
''            Total1(j) = Total1(j) + LWH.IMPORT_AMOUNT
'
'         End If
'
'          j = j + 1
'         Set HeadCf = m_Headers1.Fields(j)
'         TempStr = FormatNumberToNull(LWH.EXPORT_AMOUNT, 0) 'จ่ายออก
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
'         Call m_Details1.AddField(BodyCf)
'         Total1(j) = Total1(j) + LWH.EXPORT_AMOUNT
'
'         j = j + 1
'         Set HeadCf = m_Headers1.Fields(j)
'         TempStr = FormatNumberToNull(LWH.ACTUAL_AMOUNT, 0)  '"" 'คงเหลือ
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
'         Call m_Details1.AddField(BodyCf)
'          Total1(j) = Total1(j) + LWH.ACTUAL_AMOUNT
'
'         If Val(mcolParam("DOCUMENT_TYPE")) = 14 Then 'Bag
'            j = j + 1
'            Set HeadCf = m_Headers1.Fields(j)
'            TempStr = MyDiffEx(LWH.ACTUAL_AMOUNT * LWH.WEIGHT_PER_PACK, 1000)  '"" 'ยอดตัน
'            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
'            Call m_Details1.AddField(BodyCf)
'             Total1(j) = Total1(j) + MyDiffEx(LWH.ACTUAL_AMOUNT * LWH.WEIGHT_PER_PACK, 1000)
'          ElseIf Val(mcolParam("DOCUMENT_TYPE")) = 13 Then 'Bulk
'            j = j + 1
'            Set HeadCf = m_Headers1.Fields(j)
'            TempStr = MyDiffEx(LWH.ACTUAL_AMOUNT, 1000)   '"" 'ยอดตัน
'            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
'            Call m_Details1.AddField(BodyCf)
'            Total1(j) = Total1(j) + MyDiffEx(LWH.ACTUAL_AMOUNT, 1000)
'          End If
'
'         j = j + 1
'         Set HeadCf = m_Headers1.Fields(j)
'         TempStr = LWH.BILL_DESC  'ประเภท
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
'         Call m_Details1.AddField(BodyCf)
'
'         j = j + 1
'         Set HeadCf = m_Headers1.Fields(j)
'         TempStr = LWH.NOTE 'หมายเหตุ
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
'         Call m_Details1.AddField(BodyCf)
'
'         Call m_Details1.GetString(2, TempStr1, TempStr2)
'         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'         Call Vsp.AddTable(strFormat, "", TempStr2)
'         '=====
'      Next LWH
'
'      Call GenerateFooter("รวม", "", "", Total1, m_Details1)
'      Call m_Details1.GetString(2, TempStr1, TempStr2)
'      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'      Vsp.FontBold = True
'      Call Vsp.AddTable(strFormat, "", TempStr2)
'      Vsp.FontBold = False
'
'      Set BodyCf = Nothing
'      Set Li = Nothing
'
'   Vsp.EndDoc
'   Call CloseExportFile(Vsp)
'
'   Set LWH = Nothing
'
'   genDoc = True
'   Exit Function
'
'ErrHandler:
'   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
'   Set Rs = Nothing
'End Function
'
'Private Sub GenerateFooter(TXT1 As String, TXT2 As String, TXT3 As String, Tot() As Double, Details As CFieldList)
'Dim HeadCf As CReportField
'Dim BodyCf As CReportField
'Dim j As Long
'Dim TempStr As String
'
'   Call Details.ClearField
'   Set BodyCf = New CReportField
'
'   j = 0
'
'   For Each HeadCf In m_Headers1.Fields
'       j = j + 1
'      If Val(mcolParam("DOCUMENT_TYPE")) = 13 Then
'      Select Case j
'      Case 1
'         TempStr = TXT1
'      Case 2, 3, 4, 5, 6, 7, 14, 15
'          TempStr = ""
'      Case 13 'ยอดตัน
'         TempStr = FormatNumber(Tot(j), 3)
'      Case Else
'          TempStr = FormatNumber(Tot(j), 0)
'      End Select
'   ElseIf Val(mcolParam("DOCUMENT_TYPE")) = 14 Then
'      Select Case j
'      Case 1
'         TempStr = TXT1
'      Case 2, 3, 4, 5, 6, 7, 8, 15, 16
'         TempStr = ""
'      Case 14 'ยอดตัน
'         TempStr = FormatNumber(Tot(j), 3)
'      Case Else
'         TempStr = FormatNumber(Tot(j), 0)
'      End Select
'   End If
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
'   Vsp.FontSize = 8
'   dY(0) = Vsp.CurrentY
'   Vsp.TextAlign = taLeftBottom
'   strFormat = VSP_CalTable(TITLE_FORMAT, mdblWid, alngX)
'   Vsp.TableBorder = tbNone
'   Call Vsp.AddTable(strFormat, "", "|หน้าที่|: " & Vsp.CurrentPage)
'   Call Vsp.AddTable(strFormat, "", "|วันที่พิมพ์|: " & DateToStringExt(mdteDate))
'   Call Vsp.AddTable(strFormat, "", "|เวลา|: " & Right$(DateToStringExtEx(mdteDate), 8))
'   Call Vsp.AddTable(strFormat, "", "|ผู้ใช้|: " & glbUser.USER_NAME)
'   Vsp.CurrentY = dY(0)
'
'   Vsp.FontSize = TITLE_SIZE
'   Vsp.FontBold = True
'   Vsp.TextAlign = taCenterBottom
'   Vsp.Paragraph = MODULE_DESC
'   Vsp.Paragraph = ""
'   If mblnHeader Then
'       Call printHeader
'   End If
'   mdY = Vsp.CurrentY
'
'   Vsp.TextAlign = talnT
'   Vsp.TableBorder = tbdT
'   Vsp.FontBold = blnBold
'   Vsp.FontUnderline = blnUnder
'   Vsp.FontItalic = blnItalic
'   Vsp.FontSize = iSize
'   Vsp.FontName = sName
'End Sub
'
'Private Sub VSP_StartDoc()
''This event occur when VSPrinter.StartDoc is used and used to initialize some information before generating printed document
'End Sub
'
'
'
'
