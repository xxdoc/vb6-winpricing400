VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportInventory026_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportProduct020"
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

Private Const TITLE_SIZE = 12
Private Const HEADER_SIZE = 8
Private Const DETAIL_SIZE = 6

Private Const TITLE_FORMAT = "<32|<3|<5;"

Private m_TempCol As Collection
Private m_HeaderFormat1 As String
Private m_HeaderFormat2 As String
Private m_HeaderText1 As String
Private m_HeaderText2 As String

Private m_Headers0 As CFieldList
Private m_Headers0_1 As CFieldList
Private m_Headers1 As CFieldList
Private m_Details1 As CFieldList
Private m_Left As Double
Private m_PartTxtypes As Collection
Private m_PartTxtypeExs As Collection
Private m_CusTypes As Collection
Private m_PartTxtypeBas As Collection


Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_TempCol = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers0_1 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Details1 = New CFieldList
   Set m_PartTxtypes = New Collection
   Set m_PartTxtypeExs = New Collection
   Set m_CusTypes = New Collection
   Set m_PartTxtypeBas = New Collection
   
   Set Rs = New ADODB.Recordset
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers0_1 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_PartTxtypes = Nothing
   Set m_PartTxtypeExs = Nothing
   Set m_CusTypes = Nothing
   Set m_PartTxtypeBas = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
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
Dim TempStr2 As String
Dim Cf As CReportField
Dim CusType As CLotItem

   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll

   Call m_Headers0.ClearField
   Call m_Headers0_1.ClearField
   Call m_Headers1.ClearField

   Set Cf = New CReportField

   Call Cf.SetFieldValue(5, "^", "", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "ยกมา", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "รับเข้า", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(6 + (3 * m_CusTypes.Count), "^", "จ่ายออก", "^")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(9, "^", "คงเหลือ", "^")
   Call m_Headers0.AddField(Cf)
   
   '===========
   Call Cf.SetFieldValue(5, "^", "", "^")
   Call m_Headers0_1.AddField(Cf)
   
   Call Cf.SetFieldValue(3, "^", "ยกมา", "^")
   Call m_Headers0_1.AddField(Cf)
   
'   Call Cf.SetFieldValue(3, "^", "ซื้อ/รับ/คืน", "^")
'   Call m_Headers0_1.AddField(Cf)
'   Call Cf.SetFieldValue(3, "^", "ผลิต", "^")
'   Call m_Headers0_1.AddField(Cf)
'   Call Cf.SetFieldValue(3, "^", "โอน", "^")
'   Call m_Headers0_1.AddField(Cf)
'   Call Cf.SetFieldValue(3, "^", "ปรับยอด", "^")
'   Call m_Headers0_1.AddField(Cf)
'   Call Cf.SetFieldValue(0, "^", "ตรวจนับ", "^")
'   Call m_Headers0_1.AddField(Cf)
   
   Call Cf.SetFieldValue(3, "^", "รับเข้า", "^")
   Call m_Headers0_1.AddField(Cf)
   
   Call Cf.SetFieldValue(3, "^", "ใช้/ผลิต/โอน", "^")
   Call m_Headers0_1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "โอนเปลี่ยนประเภท", "^")
   Call m_Headers0_1.AddField(Cf)
   For Each CusType In m_CusTypes
      Call Cf.SetFieldValue(3, "^", CusType.CSTTYPE_NAME, "^")
      Call m_Headers0_1.AddField(Cf)
   Next CusType
   
   Call Cf.SetFieldValue(3, "^", "บัญชี", "^")
   Call m_Headers0_1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "PHYSICAL", "^")
   Call m_Headers0_1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "DIFF", "^")
   Call m_Headers0_1.AddField(Cf)
   
   '====
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "รหัส", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "ชื่อ", "<")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(1, "^", "จ.น.", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "@", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "มูลค่า", ">")
   Call m_Headers1.AddField(Cf)
   
'   Call Cf.SetFieldValue(1, "^", "จ.น.", ">")
'   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(1, "^", "@", ">")
'   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(1, "^", "มูลค่า", ">")
'   Call m_Headers1.AddField(Cf)
'
'   Call Cf.SetFieldValue(1, "^", "จ.น.", ">")
'   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(1, "^", "@", ">")
'   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(1, "^", "มูลค่า", ">")
'   Call m_Headers1.AddField(Cf)
'
'   Call Cf.SetFieldValue(1, "^", "จ.น.", ">")
'   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(1, "^", "@", ">")
'   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(1, "^", "มูลค่า", ">")
'   Call m_Headers1.AddField(Cf)
'
'   Call Cf.SetFieldValue(1, "^", "จ.น.", ">")
'   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(1, "^", "@", ">")
'   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(1, "^", "มูลค่า", ">")
'   Call m_Headers1.AddField(Cf)
'
'   Call Cf.SetFieldValue(0, "^", "จ.น.", ">")
'   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(0, "^", "@", ">")
'   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(0, "^", "มูลค่า", ">")
'   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(1, "^", "จ.น.", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "@", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "มูลค่า", ">")
   Call m_Headers1.AddField(Cf)
   
   '===
   Call Cf.SetFieldValue(1, "^", "จ.น.", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "@", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "มูลค่า", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(1, "^", "จ.น.", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "@", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "มูลค่า", ">")
   Call m_Headers1.AddField(Cf)
   
   For Each CusType In m_CusTypes
      Call Cf.SetFieldValue(1, "^", "จ.น.", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(1, "^", "@", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(1, "^", "มูลค่า", ">")
      Call m_Headers1.AddField(Cf)
   Next CusType
   
   '====
   Call Cf.SetFieldValue(1, "^", "จ.น.", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "@", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "มูลค่า", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(1, "^", "จ.น.", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "@", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "มูลค่า", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(1, "^", "จ.น.", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "@", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "มูลค่า", ">")
   Call m_Headers1.AddField(Cf)
   
   Set Cf = Nothing
   
   Call m_Headers0.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr2
   Call Vsp.AddTable(strFormat, "", strPrint)
   
  Call m_Headers0_1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr2
   Call Vsp.AddTable(strFormat, "", strPrint)
   
  Call m_Headers1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr2
   Call Vsp.AddTable(strFormat, "", strPrint)
            
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
   Vsp.ORIENTATION = orLandscape
   Vsp.MarginBottom = 300
   Vsp.MarginFooter = 300
   Vsp.MarginHeader = 300
   Vsp.MarginLeft = 300
   Vsp.MarginRight = 500
   Vsp.MarginTop = 300
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
   ToDate = "ถึง " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")
   
   MODULE_DESC = vbCrLf & glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & vbCrLf & _
                                       mcolParam("REPORT_NAME") & vbCrLf & _
                                        FromDate & " " & ToDate & " " & mcolParam("LOCATION_NAME")
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
Dim cData As CLotItem
Dim I As Long
Dim J As Long
Dim k  As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Pi As CPartItem
Dim Total1(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim Amt As Double
Dim InventoryBals1 As Collection
Dim Li As CLotItem
Dim TempLi1 As CLotItem
Dim TempLi2 As CLotItem
Dim Sum1 As Double
Dim Count1 As Double
Dim BalanceAccums As Collection
Dim BalanceLi As CLotItem
Dim LeftAmount1 As Double
Dim LeftAmount2 As Double
Dim TempLi As CLotItem
Dim Tot1 As Double
Dim Tot2 As Double
Dim TxValue As Double
'Dim Pi As CPartItem
Dim TempAmt As Double
Dim TempValue As Double
Dim Ct As CLotItem
Dim TempLiBa1 As CLotItem
Dim TempLiBa2 As CLotItem
   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Call LoadDistinctCusType(Nothing, m_CusTypes, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), , mcolParam("LOCATION_ID"), mcolParam("PART_TYPE"), mcolParam("PART_GROUP"))
   
   Call LoadPartTxTypeDocTypeAmount(Nothing, m_PartTxtypes, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), , mcolParam("LOCATION_ID"), mcolParam("PART_TYPE"), mcolParam("PART_GROUP"))
   
   
   Call LoadPartTxTypeDocTypeCusTypeAmount(Nothing, m_PartTxtypeExs, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), , mcolParam("LOCATION_ID"), mcolParam("PART_TYPE"), mcolParam("PART_GROUP"))
      
   Set BalanceAccums = New Collection
   Set InventoryBals1 = New Collection
    If mcolParam("FROM_DATE") > 0 Then
      Dim MonthlyAccums  As Collection
      Dim YYYYMM As String
      Dim firstDate As Date
      Dim lastDate As Date
      Set MonthlyAccums = New Collection
      Call GetFirstLastDate(mcolParam("FROM_DATE"), firstDate, lastDate)
      YYYYMM = Format(Year(DateAdd("D", -1, firstDate)), "0000") & "-" & Format(Month(DateAdd("D", -1, firstDate)), "00")
      Call LoadMonthlyBalancePartItem(Nothing, MonthlyAccums, YYYYMM, mcolParam("LOCATION_ID"))
      Call glbDaily.CopyMonthlyAccumPartItem(MonthlyAccums, InventoryBals1)
      Set MonthlyAccums = Nothing
      
      If (firstDate <> mcolParam("FROM_DATE")) Then
         Call LoadPartTxTypeAmount(Nothing, m_PartTxtypeBas, firstDate, DateAdd("D", -1, mcolParam("FROM_DATE")), , mcolParam("LOCATION_ID"))
      End If
   End If
   Set BalanceAccums = Nothing

   If Not initDoc Then Exit Function

   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------

   For J = 1 To UBound(Total1)
      Total1(J) = 0
   Next J

   Dim Ma As CMonthlyAccum
   Set Ma = New CMonthlyAccum
   Ma.PART_ITEM_ID = -1
   YYYYMM = Format(Year(mcolParam("TO_DATE")), "0000") & "-" & Format(Month(mcolParam("TO_DATE")), "00")
   Ma.YYYYMM = YYYYMM
   Ma.LOCATION_ID = mcolParam("LOCATION_ID")
   Ma.PART_TYPE = mcolParam("PART_TYPE")
   Ma.PART_GROUP = mcolParam("PART_GROUP")
   Ma.OrderBy = 1
   Call Ma.QueryData(4, Rs, iCount, False)
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
      I = 0

      Set BodyCf = New CReportField
      Set Ma = New CMonthlyAccum
      While Not Rs.EOF
         I = I + 1
         Call Ma.PopulateFromRS(4, Rs)
         
         Tot1 = 0
         Tot2 = 0
         TxValue = 0
         
         Set BalanceLi = GetLotItem(InventoryBals1, Trim(str(Ma.PART_ITEM_ID)))
         Set TempLiBa1 = GetLotItem(m_PartTxtypeBas, Ma.PART_ITEM_ID & "-" & "I")
         Set TempLiBa2 = GetLotItem(m_PartTxtypeBas, Ma.PART_ITEM_ID & "-" & "E")

         Call m_Details1.ClearField
         
         Set HeadCf = m_Headers1.Fields(1)
         TempStr = I
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(2)
         TempStr = Ma.PART_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(3)
         TempStr = Ma.PART_DESC
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         ' ยกมา ====
         LeftAmount1 = BalanceLi.NEW_AMOUNT + TempLiBa1.TX_AMOUNT - TempLiBa2.TX_AMOUNT
         Set HeadCf = m_Headers1.Fields(4)
         TempStr = FormatNumber(LeftAmount1, , "-")
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(4) = Total1(4) + LeftAmount1
         Amt = LeftAmount1
         TxValue = TxValue + Amt
         Tot1 = Tot1 + Amt
         Tot2 = Tot2 + Sum1
         
         Sum1 = BalanceLi.TOTAL_INCLUDE_PRICE + TempLiBa1.TOTAL_INCLUDE_PRICE - TempLiBa2.TOTAL_INCLUDE_PRICE
         Set HeadCf = m_Headers1.Fields(5)
         TempStr = FormatNumber(MyDiff(Sum1, LeftAmount1), , "-")
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(5) = Total1(5) + 0
         
         Set HeadCf = m_Headers1.Fields(6)
         TempStr = FormatNumber(Sum1, , "-")
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(6) = Total1(6) + (Sum1)
         Tot2 = Tot2 + Sum1
         '====

         '=== ซื้อเข้า รับเข้าทั่วไป
         Set TempLi = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-I-1")  'ซื้อเข้า
         Set TempLi2 = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-I-23")  'รับเข้าทั่วไป
         TempLi.TX_AMOUNT = TempLi.TX_AMOUNT + TempLi2.TX_AMOUNT
         TempLi.TOTAL_INCLUDE_PRICE = TempLi.TOTAL_INCLUDE_PRICE + TempLi2.TOTAL_INCLUDE_PRICE
         Set TempLi2 = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-I-18")  'รับคืน
         TempLi.TX_AMOUNT = TempLi.TX_AMOUNT + TempLi2.TX_AMOUNT
         TempLi.TOTAL_INCLUDE_PRICE = TempLi.TOTAL_INCLUDE_PRICE + TempLi2.TOTAL_INCLUDE_PRICE
         Set TempLi2 = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-I-20")  'รับเข้าจ่ายออก
         TempLi.TX_AMOUNT = TempLi.TX_AMOUNT + TempLi2.TX_AMOUNT
         TempLi.TOTAL_INCLUDE_PRICE = TempLi.TOTAL_INCLUDE_PRICE + TempLi2.TOTAL_INCLUDE_PRICE
         Amt = TempLi.TX_AMOUNT
         Sum1 = TempLi.TOTAL_INCLUDE_PRICE

         '==== รับจากการผลิต
         Set TempLi = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-I-11")  'รับจากการผลิต
         Amt = Amt + TempLi.TX_AMOUNT
         Sum1 = Sum1 + TempLi.TOTAL_INCLUDE_PRICE
         '==== รับจากการผลิต
         
         '=== รับจากการโอน
         Set TempLi = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-I-3")  'โอนเข้า
         Set TempLi2 = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-I-22")  'โอนเข้า
         TempLi.TX_AMOUNT = TempLi.TX_AMOUNT + TempLi2.TX_AMOUNT
         TempLi.TOTAL_INCLUDE_PRICE = TempLi.TOTAL_INCLUDE_PRICE + TempLi2.TOTAL_INCLUDE_PRICE
         Amt = Amt + TempLi.TX_AMOUNT
         Sum1 = Sum1 + TempLi.TOTAL_INCLUDE_PRICE
         '=== รับจากการโอน
         
         '=== รับเข้าจากการปรับยอดจากการชั่งตวง
         Set TempLi = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-I-5")  'ปรับยอด
         Amt = Amt + TempLi.TX_AMOUNT
         Sum1 = Sum1 + TempLi.TOTAL_INCLUDE_PRICE
         TxValue = TxValue + Amt
         Tot1 = Tot1 + Amt
         Tot2 = Tot2 + Sum1
         '=== รับเข้าจากการปรับยอด

         '=== รับเข้าจากการตรวจนับ
'         Set TempLi = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-I-4")  'ปรับยอด
'         Amt = TempLi.TX_AMOUNT
'         Sum1 = TempLi.TOTAL_INCLUDE_PRICE
'         TxValue = TxValue + Amt
         'Tot1 = Tot1 + Amt
         'Tot2 = Tot2 + Sum1
         '=== รับเข้าจากการตรวจนับ
         
         Set HeadCf = m_Headers1.Fields(7)
         TempStr = FormatNumber(Amt, , "-")
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(7) = Total1(7) + (Amt)
         
         Set HeadCf = m_Headers1.Fields(8)
         TempStr = FormatNumber(MyDiffEx(Sum1, Amt), , "-")
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(8) = Total1(8) + 0
         
         Set HeadCf = m_Headers1.Fields(9)
         TempStr = FormatNumber(Sum1, , "-")
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(9) = Total1(9) + Sum1
         '===
         
         '=== เบิก ผลิต
         Set TempLi = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-E-2")  'เบิก
         Amt = TempLi.TX_AMOUNT
         Sum1 = TempLi.TOTAL_INCLUDE_PRICE
         Set TempLi = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-E-11")  'ผลิต
         Amt = Amt + TempLi.TX_AMOUNT
         Sum1 = Sum1 + TempLi.TOTAL_INCLUDE_PRICE
         Set TempLi = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-E-20")  'รับเข้าจ่ายออก
         Amt = Amt + TempLi.TX_AMOUNT
         Sum1 = Sum1 + TempLi.TOTAL_INCLUDE_PRICE
         Set TempLi = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-E-3")  'โอนระหว่างคลัง
         Amt = Amt + TempLi.TX_AMOUNT
         Sum1 = Sum1 + TempLi.TOTAL_INCLUDE_PRICE
         Set TempLi = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-E-5")  'ปรับยอด '=== ปรับยอดจากการชั่งตวง
         Amt = Amt + TempLi.TX_AMOUNT
         Sum1 = Sum1 + TempLi.TOTAL_INCLUDE_PRICE
         
         TxValue = TxValue + Amt
         Tot1 = Tot1 - Amt
         Tot2 = Tot2 - Sum1

         Set HeadCf = m_Headers1.Fields(10)
         TempStr = FormatNumber(Amt, , "-")
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(10) = Total1(10) + (Amt)
         
         Set HeadCf = m_Headers1.Fields(11)
         TempStr = FormatNumber(MyDiffEx(Sum1, Amt), , "-")
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(11) = Total1(11) + 0
         
         Set HeadCf = m_Headers1.Fields(12)
         TempStr = FormatNumber(Sum1, , "-")
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(12) = Total1(12) + (Sum1)
         '=== เบิก ผลิต
         
         '=== โอนเปลี่ยนวัตถุดิบ
         
         Set TempLi2 = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-E-22")  'โอนเปลี่ยนวัตถุดิบ
         Amt = TempLi2.TX_AMOUNT
         Sum1 = TempLi2.TOTAL_INCLUDE_PRICE
         TxValue = TxValue + Amt
         Tot1 = Tot1 - Amt
         Tot2 = Tot2 - Sum1

         Set HeadCf = m_Headers1.Fields(13)
         TempStr = FormatNumber(Amt, , "-")
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(13) = Total1(13) + (Amt)

         Set HeadCf = m_Headers1.Fields(14)
         TempStr = FormatNumber(MyDiffEx(Sum1, Amt), , "-")
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(14) = Total1(14) + 0
         
         Set HeadCf = m_Headers1.Fields(15)
         TempStr = FormatNumber(Sum1, , "-")
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(15) = Total1(15) + (Sum1)
         '=== โอน
         
         k = 15
         For Each Ct In m_CusTypes
            k = k + 1
            Set TempLi = GetLotItem(m_PartTxtypeExs, Ma.PART_ITEM_ID & "-E-" & Ct.CUSTOMER_TYPE)
            Amt = TempLi.TX_AMOUNT
            Sum1 = TempLi.TOTAL_INCLUDE_PRICE
            TxValue = TxValue + Amt
            Tot1 = Tot1 - Amt
            Tot2 = Tot2 - Sum1
            
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumber(Amt, , "-")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            Total1(k) = Total1(k) + (Amt)
            
            k = k + 1
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumber(MyDiffEx(Sum1, Amt), , "-")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            Total1(k) = Total1(k) + 0
            
            k = k + 1
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumber(Sum1, , "-")
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            Total1(k) = Total1(k) + (Sum1)
            '=== ขาย
         
         Next Ct
         '=== ขายเชื่อ ขายสด
         
         '=== ปรับยอดจากการตรวจนับ
'         Set TempLi = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-E-4")  'ปรับยอด
'         Amt = TempLi.TX_AMOUNT
'         Sum1 = TempLi.TOTAL_INCLUDE_PRICE
'         TxValue = TxValue + Amt
         'Tot1 = Tot1 - Amt
         'Tot2 = Tot2 - Sum1
         
         k = k + 1
         '=== คงเหลือบัญชี
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(Tot1, , "-")
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + Tot1
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(MyDiffEx(Tot2, Tot1), , "-")
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + 0
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(Tot2, , "-")
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + Tot2
         '===
         
         '=== คงเหลือ Physical
         
         TempAmt = 0
         TempValue = 0
         
         '=== รับเข้าจากการตรวจนับ
         Set TempLi = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-I-4")  'ปรับยอด
         Amt = TempLi.TX_AMOUNT
         Sum1 = TempLi.TOTAL_INCLUDE_PRICE
         TxValue = TxValue + Amt
         Tot1 = Tot1 + Amt
         Tot2 = Tot2 + Sum1
         
         TempAmt = TempAmt + Amt
         TempValue = TempValue + Sum1
         
         
         '=== ปรับยอดจากการตรวจนับ
         Set TempLi = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-E-4")  'ปรับยอด
         Amt = TempLi.TX_AMOUNT
         Sum1 = TempLi.TOTAL_INCLUDE_PRICE
         TxValue = TxValue + Amt
         Tot1 = Tot1 - Amt
         Tot2 = Tot2 - Sum1
         
         TempAmt = TempAmt - Amt
         TempValue = TempValue - Sum1
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(Tot1, , "-")
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + Tot1
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(MyDiffEx(Tot2, Tot1), , "-")
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + 0
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(Tot2, , "-")
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + Tot2
         '===
         
         '=== Diff
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(TempAmt, , "-")
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + (TempAmt)
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(MyDiffEx(TempValue, TempAmt), , "-")
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + 0
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(TempValue, , "-")
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + (TempValue)
         '=== Diff
         
         If TxValue <> 0 Then
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
         Else
            I = I - 1
         End If
'         =====

         Rs.MoveNext
      Wend
      
      Call GenerateFooter("รวม", "", "", Total1, m_Details1)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False

      Set BodyCf = Nothing
      Set Li = Nothing
      Set Pi = Nothing
      
   Vsp.EndDoc
   Call CloseExportFile(Vsp)

   Set cData = Nothing
   Set InventoryBals1 = Nothing
   
   genDoc = True
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
   Set Rs = Nothing
End Function

Private Sub GenerateFooter(TXT1 As String, TXT2 As String, TXT3 As String, Tot() As Double, Details As CFieldList)
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
         TempStr = ""
      ElseIf J = 3 Then
         TempStr = ""
      ElseIf (J = 4) Or (J = 6) Or (J = 7) Or (J = 9) Or (J = 10) Or (J = 12) Or (J = 13) Or (J = 15) Or (J = 16) Or (J = 18) Or _
                  (J = 19) Or (J = 21) Or (J = 22) Or (J = 24) Or (J = 25) Or (J = 27) Or (J = 28) Or (J = 30) Or (J = 31) Or (J = 33) Or _
                  (J = 34) Or (J = 36) Or (J = 37) Or (J = 39) Or (J = 40) Or (J = 42) Or (J = 43) Or (J = 45) Then
            TempStr = FormatNumber(Tot(J), , "-")
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
   Call Vsp.AddTable(strFormat, "", "|วันที่พิมพ์|: " & DateToStringExt(mdteDate))
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




