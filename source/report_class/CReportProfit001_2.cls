VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportProfit001_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportProfit001_2"
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

Private TempFromDate As Date
Private TempToDate As Date

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
   
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ปี", "^")
   Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "เดือน", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "รหัสลูกค้า", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "ชื่อลูกค้า", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "รหัสพนักงานขาย", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "ชื่อพนักงานขาย", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "รหัสสินค้า", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "ชื่อสินค้า", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จำนวนถุง", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ปริมาณ", ">")
   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(2, "^", "PROMO", ">")
'   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ราคา/Kg", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "มูลค่า", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ส่วนลด", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "ขายสุทธิ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ต/ท@", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "ต้นทุนรวม", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "กำไร", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "%GP", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "รหัสขาย", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "ชื่ออาหารที่ขาย", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "ประเภทการบรรจุ", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "ประเภทลูกค้า", "<")
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
   
   FromDate = "จากเดือน " & mcolParam("FROM_MONTH") & "/" & mcolParam("FROM_YEAR")
   ToDate = "ถึงเดือน " & mcolParam("TO_MONTH") & "/" & mcolParam("TO_YEAR")
   
      Call GetFirstLastDate(DateSerial(mcolParam("TO_YEAR") - 543, mcolParam("TO_MONTH"), 1), TempFromDate, TempToDate)
      TempFromDate = DateSerial(mcolParam("FROM_YEAR") - 543, mcolParam("FROM_MONTH"), 1)
   
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
Dim Di As CDoItem
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(100) As Double
Dim Total2(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim Cm As CCustomer
Dim Pi As CPartItem
Dim Ft As CFeature

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set Rs = New ADODB.Recordset
'
   Call LoadCustomer(Nothing, m_Customers)
'   Call LoadPartItem(Nothing, m_PartItems, mcolParam("PART_TYPE"))
'   Call LoadFeature(Nothing, m_Features)
   
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------

   For J = 1 To UBound(Total1)
      Total1(J) = 0
      Total2(J) = 0
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
      Di.FROM_DATE = TempFromDate
      Di.TO_DATE = TempToDate
      Di.CUSTOMER_TYPE = mcolParam("CUSTOMER_TYPE")
      Di.PART_TYPE = mcolParam("PART_TYPE")
      Di.DocTypeSet = BillingDocType2Set(mcolParam("BILL_SUBTYPE"))
      Di.OrderBy = 1
      Di.OrderType = 1
      Call Di.QueryData(53, Rs, iCount)
      
      I = 0
      
      While Not Rs.EOF
         Call Di.PopulateFromRS(53, Rs)
         If Di.PARCEL_TYPE = 2 Then
            Di.PACK_AMOUNT = 0
         End If
         Set Cm = GetCustomer(m_Customers, Trim(str(Di.CUSTOMER_ID)))
         I = I + 1
         k = 0
''         If PrevKey1 <> Cm.CUSTOMER_CODE And I > 1 Then
''            Call GenerateFooter("รวม", "", "", Total1, m_Details1, 0)
''            Call m_Details1.GetString(1, TempStr1, TempStr2)
''            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
''            Vsp.FontBold = True
''            Call Vsp.AddTable(strFormat, "", TempStr2)
''            Vsp.FontBold = False
''
''            For J = 1 To UBound(Total1)
''               Total1(J) = 0
''            Next J
''
''            Vsp = "ลูกค้า  :  " & Cm.CUSTOMER_NAME & " / " & Cm.CUSTOMER_CODE
''         ElseIf I = 1 Then
''            Vsp = "ลูกค้า  :  " & Cm.CUSTOMER_NAME & " / " & Cm.CUSTOMER_CODE
''         End If
''
''         PrevKey1 = Cm.CUSTOMER_CODE
         
         
         Call m_Details1.ClearField
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = I
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = Left(Di.YYYYMM, 4)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = Format(Right(Di.YYYYMM, 2), "00")
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = Di.CUSTOMER_CODE
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = Di.CUSTOMER_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = Di.EMP_CODE
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = Di.SALE_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         If Di.PART_ITEM_ID > 0 Then
            TempStr1 = Di.PART_NO   'Pi.PART_NO
            TempStr2 = Di.PART_DESC '& " : " & Di.MANUAL_NAME
         ElseIf Di.FEATURE_ID > 0 Then
            TempStr1 = Di.FEATURE_CODE   'Ft.FEATURE_CODE
            TempStr2 = Di.FEATURE_DESC   'Ft.FEATURE_DESC
            Di.PACK_AMOUNT = 0
            Di.PRICE_PER_PACK = 0
            Di.ITEM_AMOUNT = 0
            Di.AVG_PRICE = 0
         End If
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr1)
         Call m_Details1.AddField(BodyCf)
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr2)
         Call m_Details1.AddField(BodyCf)
         
         Amt = Di.PACK_AMOUNT
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(Amt, 0)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + Amt
         Total2(k) = Total2(k) + Amt
         
         Amt = Di.ITEM_AMOUNT
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(Amt, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + Amt
         Total2(k) = Total2(k) + Amt
         
'         k = k + 1
'         Set HeadCf = m_Headers1.Fields(k)
'         TempStr = "0"
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'         Call m_Details1.AddField(BodyCf)
'         Total1(k) = Total1(k) + 0
'         Total2(k) = Total2(k) + 0
         
         Amt = Di.AVG_PRICE
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(Amt, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + Amt
         Total2(k) = Total2(k) + Amt
         
         Amt = Di.TOTAL_PRICE + Di.DISCOUNT_AMOUNT
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(Amt, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + (Amt)
         Total2(k) = Total2(k) + (Amt)
         
         Amt = Di.DISCOUNT_AMOUNT + Di.EXTRA_DISCOUNT
          k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(Amt, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + (Amt)
         Total2(k) = Total2(k) + (Amt)
         
         Amt = Di.TOTAL_PRICE - Di.EXTRA_DISCOUNT
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(Amt, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + (Amt)
         Total2(k) = Total2(k) + (Amt)
         
         Amt = MyDiff(Di.TOTAL_INCLUDE_PRICE, Di.ITEM_AMOUNT)
         k = k + 1                                    ' ต้นทุน @
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(Amt, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + Amt
         Total2(k) = Total2(k) + Amt
         
         Amt = Di.TOTAL_INCLUDE_PRICE
         k = k + 1                                    ' ต้นทุน
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(Amt, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + Amt
         Total2(k) = Total2(k) + Amt
          
         Amt = Di.TOTAL_PRICE - Di.EXTRA_DISCOUNT - Di.TOTAL_INCLUDE_PRICE
         k = k + 1                                    ' กำไร
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(Amt, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + Amt
         Total2(k) = Total2(k) + Amt
         
         Amt = MyDiff((Di.TOTAL_PRICE - Di.EXTRA_DISCOUNT - Di.TOTAL_INCLUDE_PRICE), (Di.TOTAL_PRICE - Di.EXTRA_DISCOUNT)) * 100
         k = k + 1                                    ' ต้นทุน @
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(Amt, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + 0
         Total2(k) = Total2(k) + 0
         
         k = k + 1 'ชนิดอาหาร
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = Di.BARCODE_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         k = k + 1 'ชื่ออาหารที่ขาย
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = Di.BILL_DESC
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         k = k + 1 'ประเภทการบรรจุอาหาร
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = Di.PART_TYPE_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         k = k + 1 'ประเภทลูกค้า
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = Di.CSTTYPE_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         
         Rs.MoveNext
      Wend
      
      Call GenerateFooter("รวม", "", "", Total1, m_Details1, 0)
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
            
'      Call GenerateFooter("สุทธิ", "", "", Total2, m_Details1, 0)
'      Call m_Details1.GetString(1, TempStr1, TempStr2)
'      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'      Vsp.FontBold = True
'      Call Vsp.AddTable(strFormat, "", TempStr2)
'      Vsp.FontBold = False

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
      ElseIf (J = 3) Then
         TempStr = TXT2
      ElseIf (J = 8) Or (J = 9) Or (J = 11) Or (J = 12) Or (J = 13) Or (J = 15) Or (J = 16) Then
         TempStr = FormatNumber(Tot(J))
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




