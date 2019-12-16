VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportSell002_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportSell002_3"
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
Private m_Employees As Collection
Private m_PartItems As Collection

Private FromDate As Date
Private ToDate As Date

Private FromDate2 As Date
Private ToDate2 As Date

Private TempFromDate As Date
Private TempToDate As Date

Private TempCollection As Collection

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
   Set m_Employees = New Collection
   Set m_PartItems = New Collection
   Set TempCollection = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_Employees = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
   Set m_PartItems = Nothing
   Set TempCollection = Nothing
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
Dim TempDate As Date
Dim TempWidth As Double

   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField
   
   Call Cf.SetFieldValue(6, "^", "", "<")
   Call m_Headers0.AddField(Cf)

   TempFromDate = DateSerial(mcolParam("FROM_YEAR") - 543, mcolParam("FROM_MONTH"), 1)
   TempToDate = DateSerial(mcolParam("TO_YEAR") - 543, mcolParam("TO_MONTH"), 1)
   '------------------------------------------------------------------------------------------>          ยอดขาย
   While TempFromDate <= TempToDate
          Call Cf.SetFieldValue(6, "^", Month(TempFromDate) & "/" & Year(TempFromDate) + 543, ">")
          Call m_Headers0.AddField(Cf)
      TempFromDate = DateAdd("M", 1, TempFromDate)
   Wend
   
   Call Cf.SetFieldValue(6, "^", "รวม", "<")
   Call m_Headers0.AddField(Cf)
   
   
   
   Call Cf.SetFieldValue(3, "^", "รหัสสินค้า", "<")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(3, "^", "ชื่อสินค้า", "<")
   Call m_Headers1.AddField(Cf)
   
   TempFromDate = DateSerial(mcolParam("FROM_YEAR") - 543, mcolParam("FROM_MONTH"), 1)
   TempToDate = DateSerial(mcolParam("TO_YEAR") - 543, mcolParam("TO_MONTH"), 1)
   '------------------------------------------------------------------------------------------>          ยอดขาย
   While TempFromDate <= TempToDate
      Call Cf.SetFieldValue(3, "^", "จำนวน", ">")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(3, "^", "มูลค่า", ">")
      Call m_Headers1.AddField(Cf)
         
         
'      If mcolParam("SHOW_AMOUNT") > 0 Then
'         Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
'         Call m_Headers1.AddField(Cf)
'      End If
'      If mcolParam("SHOW_PRICE") > 0 Then
'         Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
'         Call m_Headers1.AddField(Cf)
'      End If
      TempFromDate = DateAdd("M", 1, TempFromDate)
   Wend
   
   Call Cf.SetFieldValue(3, "^", "จำนวน", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(3, "^", "มูลค่า", ">")
   Call m_Headers1.AddField(Cf)
   
'   Call Cf.SetFieldValue(1, "^", "NO.", "^")
'   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(3, "^", "รหัสพนักงาน", "^")
'   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(4, "^", "ชื่อพนักงาน", "<")
'   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(3, "^", "รหัสสินค้า", "^")
'   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(6, "^", "ชื่อสินค้า", "<")
'   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(2, "^", "จำนวนถุง", "^")
'   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(2, "^", "น้ำหนัก", ">")
'   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(3, "^", "มูลค่า", ">")
'   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(3, "^", "ส่วนลด", ">")
'   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(3, "^", "ส่วนลดเพิ่ม", ">")
'   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(3, "^", "สุทธิ", ">")
'   Call m_Headers1.AddField(Cf)
   
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
Dim tFromDate As String
Dim tToDate As String
   
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
   
   Call GetFirstLastDate(DateSerial(mcolParam("FROM_YEAR") - 543, mcolParam("FROM_MONTH"), 1), FromDate, ToDate)
   Call GetFirstLastDate(DateSerial(mcolParam("TO_YEAR") - 543, mcolParam("TO_MONTH"), 1), FromDate2, ToDate2)
'
   tFromDate = "จากเดือน " & IntToThaiMonth(mcolParam("FROM_MONTH")) & " " & mcolParam("FROM_YEAR")
   tToDate = "ถึงเดือน " & EmptyToString(IntToThaiMonth(Val(mcolParam("TO_MONTH"))) & " " & mcolParam("TO_YEAR"), "N/A")
   
   MODULE_DESC = vbCrLf & glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & vbCrLf & _
                                       mcolParam("REPORT_NAME") & vbCrLf & _
                                        tFromDate & " " & tToDate
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
Dim Di As CDoItem
Dim TempDI As CDoItem
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(100) As Double
Dim Total2(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim PrevKey3 As String
Dim Cm As CEmployee
Dim Pi As CPartItem
Dim TotalItemAmount As Double
Dim TotalBeforDiscount As Double
Dim TempKey As String


   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set Rs = New ADODB.Recordset
   
   Call LoadEmployee(Nothing, m_Employees)
'   Call LoadPartItem(Nothing, m_PartItems, mcolParam("PART_TYPE"))
   
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------

   For J = 1 To UBound(Total1)
      Total1(J) = 0
      Total2(J) = 0
   Next J
   
   Call GetFirstLastDate(DateSerial(mcolParam("TO_YEAR_NO") - 543, mcolParam("TO_MONTH_ID"), 1), TempFromDate, TempToDate)
   TempFromDate = DateSerial(mcolParam("FROM_YEAR_NO") - 543, mcolParam("FROM_MONTH_ID"), 1)
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
   I = 0
   Set BodyCf = New CReportField
   
   Set Di = New CDoItem
   Di.DO_ITEM_ID = -1
   Di.EMP_CODE = mcolParam("EMP_CODE")
   Di.PART_NO = mcolParam("PART_NO")
   Di.FROM_DATE = TempFromDate
   Di.TO_DATE = TempToDate
   Di.CURRENT_POSITION = mcolParam("POSITION_ID")
   Di.PART_TYPE = mcolParam("PART_TYPE")
   Di.CUSTOMER_TYPE = mcolParam("CUSTOMER_TYPE")
   Di.DocTypeSet = BillingDocType2Set(mcolParam("BILL_SUBTYPE"))
   Di.CUSTOMER_CODE = mcolParam("CUSTOMER_CODE")
   Di.FROM_CUSTOMER_CODE = mcolParam("FROM_CUSTOMER_CODE")
   Di.TO_CUSTOMER_CODE = mcolParam("TO_CUSTOMER_CODE")
   Di.OrderBy = mcolParam("ORDER_BY")
   If Di.OrderBy <= 0 Then
      Di.OrderBy = 1
   End If
   Di.OrderType = mcolParam("ORDER_TYPE")
   Call Di.QueryData(40, Rs, iCount)
      
      While Not Rs.EOF
         Set Di = Nothing
         Set Di = New CDoItem
         Call Di.PopulateFromRS(40, Rs)
'         If Di.PART_NO <> "" Then
              TempKey = Trim(Di.EMP_CODE) & "-" & Trim(Di.PART_NO) & "-" & Di.YYYYMM
              Set TempDI = GetObject("CDoItem", TempCollection, TempKey, False)
             If TempDI Is Nothing Then
               Call TempCollection.add(Di, TempKey)
             End If
'          End If
          Rs.MoveNext
      Wend
      
      I = 0
      PrevKey1 = ""
      Rs.MoveFirst
      If Not Rs.EOF Then
         Set Di = Nothing
         Set Di = New CDoItem
         Call Di.PopulateFromRS(40, Rs)
         PrevKey1 = Di.ACCEPT_BY
      End If
      
'      Set Di = Nothing
'      Set Di = New CDoItem
      While Not Rs.EOF
          Set Di = Nothing
         Set Di = New CDoItem
         Call Di.PopulateFromRS(40, Rs)
         If Di.PARCEL_TYPE = 2 Then
            Di.PACK_AMOUNT = 0
            
         End If
     If PrevKey3 <> Di.PART_ITEM_ID & "-" & Di.PART_ITEM_ID Then
       PrevKey3 = Di.PART_ITEM_ID & "-" & Di.PART_ITEM_ID
         If Di.PART_ITEM_ID > 0 Then
            If PrevKey1 <> Di.ACCEPT_BY Then
               Call GenerateFooter("รวม", "", "", Total1, m_Details1, 0)
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False

               For J = 1 To UBound(Total1)
                  Total1(J) = 0
               Next J
         ElseIf I = 0 Then
             Set Cm = GetEmployee(m_Employees, Trim(str(Di.ACCEPT_BY)))
            Vsp = "     พนักงานขาย    : (" & Cm.EMP_CODE & ")" & Cm.NAME
            PrevKey2 = Cm.EMP_CODE
         End If
         
         If PrevKey2 <> Cm.EMP_CODE Then
            Set Cm = GetEmployee(m_Employees, Trim(str(Di.ACCEPT_BY)))
            Vsp = "     พนักงานขาย    : (" & Cm.EMP_CODE & ")" & Cm.NAME
            PrevKey2 = Cm.EMP_CODE
         End If
      
         
           
            PrevKey1 = Di.ACCEPT_BY
            
            Call m_Details1.ClearField
            Set Cm = GetEmployee(m_Employees, Trim(str(Di.ACCEPT_BY)))
'            Set Pi = GetPartItem(m_PartItems, Trim(Str(Di.PART_ITEM_ID)))

             I = 1
            Set HeadCf = m_Headers1.Fields(I)
            TempStr = Di.PART_NO
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            I = I + 1
            Set HeadCf = m_Headers1.Fields(I)
            TempStr = Di.PART_DESC
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
               
            TempFromDate = DateSerial(mcolParam("FROM_YEAR") - 543, mcolParam("FROM_MONTH"), 1)
            TempToDate = DateSerial(mcolParam("TO_YEAR") - 543, mcolParam("TO_MONTH"), 1)
         '------------------------------------------------------------------------------------------>          ยอดขาย
            While TempFromDate <= TempToDate
             TempKey = Trim(Di.EMP_CODE) & "-" & Trim(Di.PART_NO) & "-" & Year(TempFromDate) & "-" & Format(Month(TempFromDate), "00") 'Trim(Di.EMP_CODE) & "-" & Trim(Di.PART_NO) & "-" & Di.YYYYMM
            Set TempDI = GetObject("CDoItem", TempCollection, TempKey)
             If Not TempDI Is Nothing Then
          
               I = I + 1
               Set HeadCf = m_Headers1.Fields(I)
               TempStr = FormatNumber(TempDI.ITEM_AMOUNT)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Total1(I) = Total1(I) + TempDI.ITEM_AMOUNT
               Total2(I) = Total2(I) + TempDI.ITEM_AMOUNT
               TotalItemAmount = TotalItemAmount + TempDI.ITEM_AMOUNT
               
               I = I + 1
               Set HeadCf = m_Headers1.Fields(I)
               TempStr = FormatNumber(TempDI.BEFORE_DISCOUNT_PRICE)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Total1(I) = Total1(I) + (TempDI.BEFORE_DISCOUNT_PRICE)
               Total2(I) = Total2(I) + (TempDI.BEFORE_DISCOUNT_PRICE)
               TotalBeforDiscount = TotalBeforDiscount + TempDI.BEFORE_DISCOUNT_PRICE
               Else
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
               End If
                  
               TempFromDate = DateAdd("M", 1, TempFromDate)
            Wend
      
            I = I + 1
            Set HeadCf = m_Headers1.Fields(I)
            TempStr = FormatNumber(TotalItemAmount)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(I) = Total1(I) + TotalItemAmount
            Total2(I) = Total2(I) + TotalItemAmount
            TotalItemAmount = 0

            I = I + 1
            Set HeadCf = m_Headers1.Fields(I)
            TempStr = FormatNumber(TotalBeforDiscount)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(I) = Total1(I) + TotalBeforDiscount
            Total2(I) = Total2(I) + TotalBeforDiscount
            TotalBeforDiscount = 0

            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
            End If
            
            End If

         Rs.MoveNext
      Wend
      
'      While Not Rs.EOF
'         Call Di.PopulateFromRS(4, Rs)
'         If Di.PARCEL_TYPE = 2 Then
'            Di.PACK_AMOUNT = 0
'         End If
'
'         If Di.PART_ITEM_ID > 0 Then
'            If PrevKey1 <> Di.ACCEPT_BY Then
'               Call GenerateFooter("รวม", "", "", Total1, m_Details1, 0)
'               Call m_Details1.GetString(1, TempStr1, TempStr2)
'               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'               Vsp.FontBold = True
'               Call Vsp.AddTable(strFormat, "", TempStr2)
'               Vsp.FontBold = False
'
'               For J = 1 To UBound(Total1)
'                  Total1(J) = 0
'               Next J
'            End If
'            PrevKey1 = Di.ACCEPT_BY
'
'            Call m_Details1.ClearField
'            Set Cm = GetEmployee(m_Employees, Trim(str(Di.ACCEPT_BY)))
''            Set Pi = GetPartItem(m_PartItems, Trim(Str(Di.PART_ITEM_ID)))
'
'            I = I + 1
'            Set HeadCf = m_Headers1.Fields(1)
'            TempStr = I
'            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'            Call m_Details1.AddField(BodyCf)
'
'            Set HeadCf = m_Headers1.Fields(2)
'            TempStr = Cm.EMP_CODE
'            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'            Call m_Details1.AddField(BodyCf)
'
'            Set HeadCf = m_Headers1.Fields(3)
'            TempStr = Cm.NAME
'            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'            Call m_Details1.AddField(BodyCf)
'
'            Set HeadCf = m_Headers1.Fields(4)
'            TempStr = Di.PART_NO
'            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'            Call m_Details1.AddField(BodyCf)
'
'            Set HeadCf = m_Headers1.Fields(5)
'            TempStr = Di.PART_DESC
'            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'            Call m_Details1.AddField(BodyCf)
'
'            Set HeadCf = m_Headers1.Fields(6)
'            TempStr = FormatNumber(Di.PACK_AMOUNT, 0)
'            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'            Call m_Details1.AddField(BodyCf)
'            Total1(6) = Total1(6) + Di.PACK_AMOUNT
'            Total2(6) = Total2(6) + Di.PACK_AMOUNT
'
'            Set HeadCf = m_Headers1.Fields(7)
'            TempStr = FormatNumber(Di.ITEM_AMOUNT)
'            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'            Call m_Details1.AddField(BodyCf)
'            Total1(7) = Total1(7) + Di.ITEM_AMOUNT
'            Total2(7) = Total2(7) + Di.ITEM_AMOUNT
'
'            Set HeadCf = m_Headers1.Fields(8)
'            TempStr = FormatNumber(Di.BEFORE_DISCOUNT_PRICE)
'            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'            Call m_Details1.AddField(BodyCf)
'            Total1(8) = Total1(8) + (Di.BEFORE_DISCOUNT_PRICE)
'            Total2(8) = Total2(8) + (Di.BEFORE_DISCOUNT_PRICE)
'
'            Set HeadCf = m_Headers1.Fields(9)
'            TempStr = FormatNumber(Di.DISCOUNT_AMOUNT)
'            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'            Call m_Details1.AddField(BodyCf)
'            Total1(9) = Total1(9) + (Di.DISCOUNT_AMOUNT)
'            Total2(9) = Total2(9) + (Di.DISCOUNT_AMOUNT)
'
'            Set HeadCf = m_Headers1.Fields(10)
'            TempStr = FormatNumber(Di.EXTRA_DISCOUNT)
'            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'            Call m_Details1.AddField(BodyCf)
'            Total1(10) = Total1(10) + (Di.EXTRA_DISCOUNT)
'            Total2(10) = Total2(10) + (Di.EXTRA_DISCOUNT)
'
'            Set HeadCf = m_Headers1.Fields(11)
'            TempStr = FormatNumber(Di.TOTAL_PRICE - Di.EXTRA_DISCOUNT)
'            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'            Call m_Details1.AddField(BodyCf)
'            Total1(11) = Total1(11) + (Di.TOTAL_PRICE - Di.EXTRA_DISCOUNT)
'            Total2(11) = Total2(11) + (Di.TOTAL_PRICE - Di.EXTRA_DISCOUNT)
'
'            Call m_Details1.GetString(1, TempStr1, TempStr2)
'            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'            Call Vsp.AddTable(strFormat, "", TempStr2)
'         End If
'
'         Rs.MoveNext
'      Wend

      Call GenerateFooter("รวม", "", "", Total1, m_Details1, 0)
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
               
      Call GenerateFooter("สุทธิ", "", "", Total2, m_Details1, 0)
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
      Else
         TempStr = FormatNumber(Tot(J))
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


