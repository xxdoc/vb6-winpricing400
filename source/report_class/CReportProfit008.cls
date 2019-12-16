VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportProfit008"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportProfit008"
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
Private m_Customers As Collection
Private m_PartItems As Collection
Private m_Features As Collection

Private m_ProcessParams As Collection
Private m_ExpenseAmounts As Collection
Private m_SumParamByPartParams As Collection
Private m_ExWorkPricesItem As Collection
Private m_CostPerPart As Collection

Private m_OrderByPart As Collection

Private TempD As CExWorksPrice

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
   Set m_Customers = New Collection
   Set m_PartItems = New Collection
   Set m_Features = New Collection

   Set m_ProcessParams = New Collection
   Set m_ExpenseAmounts = New Collection
   Set m_SumParamByPartParams = New Collection
   
   Set m_ExWorkPricesItem = New Collection
   Set m_CostPerPart = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing

   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_Customers = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
   Set m_PartItems = Nothing
   Set m_Features = Nothing
   Set m_ProcessParams = Nothing
   Set m_ExpenseAmounts = Nothing
   Set m_SumParamByPartParams = Nothing
   Set m_ExWorkPricesItem = Nothing
   Set m_CostPerPart = Nothing
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
Dim PP As CParameterProcess

   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize

   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll

   Call m_Headers0.ClearField
   Call m_Headers1.ClearField

   Set Cf = New CReportField

If mcolParam("SHOW_ORDER_BILL") = 1 Then
   Call Cf.SetFieldValue(3, "^", "ประเภทสินค้า.", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(18, "^", "รายได้จาการขาย.", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(14, "^", "ต้นทุนขาย.", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(8, "^", "กำไรขั้นต้น", "^")
   Call m_Headers0.AddField(Cf)

   Call Cf.SetFieldValue(3, "^", "", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ขาย(ตัน)", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(0, "^", "จำนวนถุง", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ราคาประกาศ/ถุง", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ค่าใช้จ่าย/ถุง", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "โปรโมชั่น/ถุง", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ราคารวม/ถุง", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ขายจริง/ถุง", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ราคาต่าง", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ราคาขาย/กก.", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ขาย(บาท)", ">")
   Call m_Headers1.AddField(Cf)

   Call Cf.SetFieldValue(2, "^", "วัตถุดิบ/ตัน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "Shrink/ตัน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ค่าถุง/ตัน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "OH./ตัน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "รวมต้นทุน/ตัน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ต้นทุน/ถุง", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ต้นทุน(บาท)", ">")
   Call m_Headers1.AddField(Cf)

   Call Cf.SetFieldValue(2, "^", "GP/Ton", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "GP/Bag", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "GP(B)", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "%GP", ">")
   Call m_Headers1.AddField(Cf)
End If


   Set Cf = Nothing

   Call m_Headers0.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)

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
   Vsp.ORIENTATION = orLandscape
   Vsp.MarginBottom = 300
   Vsp.MarginFooter = 300
   Vsp.MarginHeader = 300
   Vsp.MarginLeft = 300
   Vsp.MarginRight = 500
   Vsp.MarginTop = 300
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   Call LoadExWorksPriceItem(Nothing, m_ExWorkPricesItem, , 2, mcolParam("FROM_DATE"), mcolParam("TO_DATE"))
'   Call LoadLotItemFindCostByBilling(Nothing, m_CostPerPart)
'   Call LoadCostPerUnitByPartNo(Nothing, m_CostPerPart, mcolParam("FROM_DATE"), mcolParam("TO_DATE"))
   'LoadCostPerUnitByPartNo

   FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
   ToDate = "ถึง " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")

   MODULE_DESC = vbCrLf & glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & vbCrLf & _
                                       mcolParam("REPORT_NAME") & vbCrLf & _
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
Private Function OrderDoc(Di As CDoItem)
Dim TempDI As CDoItem
Dim Key As String
      While Not Rs.EOF
         Call Di.PopulateFromRS(43, Rs)
            Key = Trim(str(Di.CUSTOMER_TYPE) & "-" & Di.ANIMAL_NAME & "-" & Di.PART_MASTER_NO)
            Set TempDI = GetObject("CDoItem", m_OrderByPart, Key, False)
         If Not TempDI Is Nothing Then
            'Amt = MyDiff(Lt.TOTAL_INCLUDE_PRICE, Lt.TX_AMOUNT) * 1000
         Else
           Call m_OrderByPart.add(Di, Key)
         End If
'          m_OrderByPart
            Rs.MoveNext
      Wend
'SHOW_ORDER_PART
End Function
Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim RName As String
Dim I As Long
Dim J As Long
Dim k As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Amt As Double
Dim Amt2 As Double
Dim Di As CDoItem
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
Dim PrevKey3 As String
Dim PrevKey4 As String
Dim Pi As CPartItem
Dim Ft As CFeature
Dim GP As Double
Dim PP As CParameterProcess
Dim Jp As CJobParameter
Dim Lt As CLotItem
Dim Cost As Double
Dim Cost2
Dim Jip As CJobInput
Dim AMOUNT As Double


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

      Set Di = New CDoItem
      Di.DO_ITEM_ID = -1
      Di.CUSTOMER_CODE = mcolParam("CUSTOMER_CODE")
      Di.FROM_CUSTOMER_CODE = mcolParam("FROM_CUSTOMER_CODE")
      Di.TO_CUSTOMER_CODE = mcolParam("TO_CUSTOMER_CODE")
      Di.PART_NO = mcolParam("PART_NO")
      Di.FROM_DATE = mcolParam("FROM_DATE")
      Di.TO_DATE = mcolParam("TO_DATE")
      Di.CUSTOMER_TYPE = mcolParam("CUSTOMER_TYPE")
      Di.PART_TYPE = mcolParam("PART_TYPE")
      Di.DocTypeSet = BillingDocType2Set(mcolParam("BILL_SUBTYPE"))
      Di.OrderBy = mcolParam("ORDER_BY")
      If Di.OrderBy <= 0 Then
         Di.OrderBy = 3
      End If
      Di.OrderType = mcolParam("ORDER_TYPE")
      Call Di.QueryData(43, Rs, iCount)

      I = 0

      While Not Rs.EOF
         Call Di.PopulateFromRS(43, Rs)
         I = I + 1
         k = 0
          If PrevKey1 <> Di.CSTTYPE_NAME And I > 1 Then
            Call GenerateFooter("รวม", "", "", Total1, m_Details1, 0, Total3)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False

            For J = 1 To UBound(Total1)
               Total1(J) = 0
            Next J

            Vsp = "ประเภทลูกค้า  :  " & Di.CSTTYPE_NAME
'            Vsp = "ชนิดอาหาร  :  " & Di.ANIMAL_NAME

            If mcolParam("SHOW_CUS_NAME") = 1 Then
               Vsp = "ชื่อลูกค้า  :  " & Di.CUSTOMER_NAME
            End If
            If mcolParam("SHOW_ORDER_BILL") = 1 Then
               Vsp = "เลขที่เอกสาร  :  " & Di.DOCUMENT_NO
            End If
'         ElseIf PrevKey2 <> Di.ANIMAL_NAME And I > 1 Then
'            Vsp = "ชนิดอาหาร  :  " & Di.ANIMAL_NAME
'            If mcolParam("SHOW_CUS_NAME") = 1 Then
'               Vsp = "ชื่อลูกค้า  :  " & Di.CUSTOMER_NAME
'            End If
         ElseIf PrevKey3 <> Di.CUSTOMER_NAME And I > 1 Then
            If mcolParam("SHOW_CUS_NAME") = 1 Then
               Vsp = "ชื่อลูกค้า  :  " & Di.CUSTOMER_NAME
            End If
            If mcolParam("SHOW_ORDER_BILL") = 1 Then
               Vsp = "เลขที่เอกสาร  :  " & Di.DOCUMENT_NO
            End If
         ElseIf PrevKey4 <> Di.DOCUMENT_NO And I > 1 Then
            If mcolParam("SHOW_ORDER_BILL") = 1 Then
               Vsp = "เลขที่เอกสาร  :  " & Di.DOCUMENT_NO
            End If
         ElseIf I = 1 Then
            Vsp = "ประเภทลูกค้า  :  " & Di.CSTTYPE_NAME
'            Vsp = "ชนิดอาหาร  :  " & Di.ANIMAL_NAME
            If mcolParam("SHOW_CUS_NAME") = 1 Then
               Vsp = "ชื่อลูกค้า  :  " & Di.CUSTOMER_NAME
            End If
            If mcolParam("SHOW_ORDER_BILL") = 1 Then
               Vsp = "เลขที่เอกสาร  :  " & Di.DOCUMENT_NO
            End If
         End If

         PrevKey1 = Di.CSTTYPE_NAME
         PrevKey2 = Di.ANIMAL_NAME
         PrevKey3 = Di.CUSTOMER_NAME
         PrevKey4 = Di.DOCUMENT_NO
         
         k = k + 1
         Call m_Details1.ClearField
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = Di.PART_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         k = k + 1                                    ' ขาย(ตัน)
         Amt = Di.ITEM_AMOUNT / 1000     ' ตัน
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(Amt, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + (Amt)
         Total2(k) = Total2(k) + (Amt)
         Total3(k) = -2
         
         k = k + 1
         Amt = Di.PACK_AMOUNT
         Set HeadCf = m_Headers1.Fields(k)  'จำนวนถุง
         TempStr = FormatNumber(Amt)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + Amt
         Total2(k) = Total2(k) + Amt
         Total3(k) = -2
         AMOUNT = Amt
        

         k = k + 1
        Amt = Di.PACKAGE_RATE
        Cost2 = Amt
         Set HeadCf = m_Headers1.Fields(k)  'ราคาประกาศ/ถุง
         TempStr = FormatNumber(Di.PACKAGE_RATE, -1)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + (AMOUNT * Amt)
         Total2(k) = Total2(k) + (AMOUNT * Amt)
         Total3(k) = -1
        
        k = k + 1
        Amt = SumOther(Di, Di.PART_TYPE)
        Cost2 = Cost2 + Amt
         Set HeadCf = m_Headers1.Fields(k)  'ค่าใช้จ่าย/ถุง
         TempStr = FormatNumber(Amt, -1)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + (AMOUNT * Amt)
         Total2(k) = Total2(k) + (AMOUNT * Amt)
         Total3(k) = -1

         k = k + 1
         Amt = Di.DISCOUNT_PER_PACK
         Cost2 = Cost2 - Amt
         Set HeadCf = m_Headers1.Fields(k)  'โปรโมชั่น/ถุง
         TempStr = IIf(Amt > 0, "(" & FormatNumber(Amt, -1) & ")", 0)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + (AMOUNT * Amt)
         Total2(k) = Total2(k) + (AMOUNT * Amt)
         Total3(k) = -1

         k = k + 1
         Amt = Cost2
         Set HeadCf = m_Headers1.Fields(k)  'ราคารวม/ถุง
         TempStr = FormatNumber(Amt, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + (AMOUNT * Amt)
         Total2(k) = Total2(k) + (AMOUNT * Amt)
         Total3(k) = -1
         

         k = k + 1   'ขายจริง/ถุง
         Amt = Di.PRICE_PER_PACK
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(Amt, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + (AMOUNT * Amt)
         Total2(k) = Total2(k) + (AMOUNT * Amt)
         Total3(k) = -1
        
         k = k + 1 'ส่วนต่าง
         Amt = Di.PRICE_PER_PACK - Amt
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = IIf(Amt >= 0, FormatNumber(Amt), "(" & Abs(Amt) & ")")
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + (AMOUNT * Amt)
         Total2(k) = Total2(k) + (AMOUNT * Amt)
         Total3(k) = -1
         
         Amt = MyDiff(Di.PRICE_PER_PACK, Di.WEIGHT_PER_PACK)         'ราคาขาย/กก.
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(Amt, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + (AMOUNT * Amt)
         Total2(k) = Total2(k) + (AMOUNT * Amt)
         Total3(k) = -1

         Amt = Di.ITEM_AMOUNT * Amt   'ราคาขาย(บาท)
         Cost = Amt
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(Amt, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + (Amt)
         Total2(k) = Total2(k) + (Amt)
         Total3(k) = -2
      
         Amt = 0
         Call LoadLotItemFindCostByBilling(Nothing, m_CostPerPart, Di.DO_ID)
         Set Lt = GetObject("CLotItem", m_CostPerPart, Trim(str(Di.DO_ID) & "-" & str(Di.PART_ITEM_ID) & "-" & "21"), False)     ' IVN ต้นทุนต่อตันเฉพาะ raw mat
         If Not Lt Is Nothing Then
            Amt = MyDiff(Lt.TOTAL_INCLUDE_PRICE, Lt.TX_AMOUNT) * 1000
         End If
         
         Cost2 = Amt
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(Amt, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + (AMOUNT * Amt)
         Total2(k) = Total2(k) + (AMOUNT * Amt)
         Total3(k) = -1
   
         Amt = Amt * MyDiff(Val(mcolParam("SHRINK")), 100)   'Shrink
         Cost2 = Cost2 + Amt
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(Amt, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + (AMOUNT * Amt)
         Total2(k) = Total2(k) + (AMOUNT * Amt)
         Total3(k) = -1
         
         Amt = 0
         Set Lt = GetObject("CLotItem", m_CostPerPart, Trim(str(Di.DO_ID) & "-" & str(Di.PART_ITEM_ID) & "-" & "21"), False)    'ค่าถุง/ตัน
         If Not Lt Is Nothing Then
             Amt = Lt.TX_AMOUNT
         End If
         
         Set Lt = GetObject("CLotItem", m_CostPerPart, Trim(str(Di.DO_ID) & "-" & str(Di.PART_ITEM_ID) & "-" & "B"), False)    'ค่าถุง/ตัน
         If Not Lt Is Nothing Then
             Amt = MyDiff(1000, Amt) * Lt.TOTAL_INCLUDE_PRICE
         End If
         Cost2 = Cost2 + Amt
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(Amt, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + (AMOUNT * Amt)
         Total2(k) = Total2(k) + (AMOUNT * Amt)
         Total3(k) = -1
   
       Amt = 0
       If Di.CUSTOMER_ID = 1191 Then 'เฉพาะนำโชค
            If Di.PRODUCT_TYPE_ID = 221 Then 'ผง
                Amt = Val(mcolParam("OH_NC_POWDER"))
             ElseIf Di.PRODUCT_TYPE_ID = 222 Or Di.PRODUCT_TYPE_ID = 227 Then      'เม็ด หรือ ครัม
                Amt = Val(mcolParam("OH_NC_GRAIN"))
            End If
       Else
         If Di.PRODUCT_TYPE_ID = 221 Then 'ผง
             Amt = Val(mcolParam("OH_POWDER"))
          ElseIf Di.PRODUCT_TYPE_ID = 222 Or Di.PRODUCT_TYPE_ID = 227 Then      'เม็ด หรือ ครัม
             Amt = Val(mcolParam("OH_GRAIN"))
         End If
      End If
         
         Cost2 = Cost2 + Amt
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(Amt, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + (AMOUNT * Amt)
         Total2(k) = Total2(k) + (AMOUNT * Amt)
         Total3(k) = -1
         
         Amt = Cost2 ' ต้นทุนรวม
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(Amt, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + (AMOUNT * Amt)
         Total2(k) = Total2(k) + (AMOUNT * Amt)
         Total3(k) = -1
         
         Amt = MyDiff(Cost2, 1000) * Di.WEIGHT_PER_PACK  ' ต้นทุน/ถุง
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(Amt, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + (AMOUNT * Amt)
         Total2(k) = Total2(k) + (AMOUNT * Amt)
         Total3(k) = -1
         
         Amt = (Di.ITEM_AMOUNT / 1000) * Cost2 'ต้นทุน(บาท)
         Cost2 = Amt
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(Amt, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + (AMOUNT * Amt)
         Total2(k) = Total2(k) + (AMOUNT * Amt)
         Total3(k) = -1

         GP = (Cost - Cost2)
         Amt = MyDiff(GP, (Di.ITEM_AMOUNT / 1000))  ' GP/TONS
          k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(Amt, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + (AMOUNT * Amt)
         Total2(k) = Total2(k) + (AMOUNT * Amt)
         Total3(k) = -1
         
         Amt = MyDiff(GP, Di.PACK_AMOUNT)   ' GP/Bag
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(Amt, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + (AMOUNT * Amt)
         Total2(k) = Total2(k) + (AMOUNT * Amt)
         Total3(k) = -1

         Amt = GP  'GP(B)
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(Amt, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + (AMOUNT * Amt)
         Total2(k) = Total2(k) + (AMOUNT * Amt)
         Total3(k) = -1
         
         Amt = MyDiff(GP, Cost) * 100
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(Amt, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + (AMOUNT * Amt)
         Total2(k) = Total2(k) + (AMOUNT * Amt)
         Total3(k) = -1

'         If mcolParam("SHOW_ORDER_PART") = 1 Then
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
'         End If

         Rs.MoveNext
      Wend
      
   Call GenerateFooter("รวม", "", "", Total1, m_Details1, 0, Total3)
   Call m_Details1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False

      Call GenerateFooter("สุทธิ", "", "", Total2, m_Details1, 0, Total3)
      Call m_Details1.GetString(1, TempStr1, TempStr2)
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

Private Sub GenerateFooter(TXT1 As String, TXT2 As String, TXT3 As String, Tot() As Double, Details As CFieldList, DateCount As Long, TotS() As Double)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim J As Long
Dim TempStr As String
Dim Amt As Double

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   J = 0
   For Each HeadCf In m_Headers1.Fields
      J = J + 1
      If J = 1 Then
         TempStr = TXT1
'      ElseIf J = 2 Then
'         TempStr = TXT2
'      ElseIf J = 3 Then
'         TempStr = TXT3
      ElseIf TotS(J) = -1 Then 'ให้ทำการหารกับ ราคารวมต่อหน่วย
         TempStr = FormatNumber(MyDiffEx(Tot(J), Tot(3)))
      ElseIf TotS(J) = -2 Then
          TempStr = FormatNumber(Tot(J))
      ElseIf TotS(J) = -3 Then
         Amt = MyDiffEx(Tot(J), Tot(6))
         TempStr = IIf(Amt > 0, "(" & FormatNumber(Amt) & ")", FormatNumber(Amt))
      Else
         TempStr = ""
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.Align)
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

Private Function SumOther(Di As CDoItem, Pt As Long) As Double
Dim RateOther As Double
  If Pt = 10 Then
       RateOther = Di.PRO_COMMISSION_BAG + Di.PRO_CHEER_BAG + Di.PRO_DST_BAG + Di.PRO_OTHER1_BAG + Di.PRO_OTHER2_BAG + Di.PRO_OTHER3_BAG
   ElseIf Pt = 21 Then
       RateOther = Di.PRO_COMMISSION_KG + Di.PRO_CHEER_KG + Di.PRO_DST_KG + Di.PRO_OTHER1_KG + Di.PRO_OTHER2_KG + Di.PRO_OTHER3_KG
   End If
   SumOther = RateOther
End Function



