VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportSellCT001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportSellCT001"
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
   
   Call Cf.SetFieldValue(11, "^", "ข้อมูลเอกสาร", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(8, "^", "ข้อมูลราคาสินค้า/หน่วย", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "ข้อมูลราคาค่าขนส่ง/หน่วย", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(12, "^", "ข้อมูลอื่นๆที่คิดเพิ่มจากราคาประกาศ/หน่วย", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(11, "^", "ข้อมูลราคาขายจริง", "^")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "เลขที่", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "วันที่", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "ชื่อลูกค้า", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ประกาศ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ลดหน้าบิล", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ลดหลังบิล", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ประกาศ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ส่วนลด", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ลดเพิ่ม", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ค่าคอม", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ค่าเชียร์", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ค่าทอย", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "อื่นๆ 1", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "อื่นๆ 2", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "อื่นๆ 3", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "หน่วย", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "คิดแบบ", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ขายจริง", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "มูลค่าขาย", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "อนุมัติ", "^")
   Call m_Headers1.AddField(Cf)
   
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
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Amt As Double
Dim Di As CDoItem
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim TempStr3 As String
Dim TempStr4 As String
Dim Total1(100) As Double
Dim Total2(100) As Double
Dim Total3(100) As Double
Dim Total4(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim Cm As CCustomer
Dim Pi As CPartItem
Dim Ft As CFeature
Dim CountDoc As Long
Dim CountDoc2 As Long
Dim AMOUNT As Long
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
      Total4(J) = 0
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
         Di.OrderBy = 2
      End If
      Di.OrderType = mcolParam("ORDER_TYPE")
      Call Di.QueryData(44, Rs, iCount)

      I = 0

      While Not Rs.EOF
        J = 0
         Call Di.PopulateFromRS(44, Rs)
         If Di.PART_ITEM_ID > 0 Then
            If Di.PARCEL_TYPE = 2 Then
               Di.PACK_AMOUNT = 0
            End If
         If Di.PART_ITEM_ID > 0 Then
            TempStr3 = Di.PART_NO   'Pi.PART_NO
            TempStr4 = Di.PART_DESC & " : " & Di.MANUAL_NAME
         ElseIf Di.FEATURE_ID > 0 Then
            TempStr3 = Di.FEATURE_CODE   'Ft.FEATURE_CODE
            TempStr4 = Di.FEATURE_DESC   'Ft.FEATURE_DESC
            Di.PACK_AMOUNT = 0
            Di.PRICE_PER_PACK = 0
            Di.ITEM_AMOUNT = 0
            Di.AVG_PRICE = 0
         End If
         
       If PrevKey1 <> TempStr3 And I > 0 Then
            Call GenerateFooter("รวม", "", "", Total1, m_Details1, CountDoc, Total3)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
            
            For J = 1 To UBound(Total1)
               Total1(J) = 0
            Next J
            J = 0
            Vsp = "สินค้า : " & TempStr3 & " (" & TempStr4 & ")"
         ElseIf I = 0 Then
             Vsp = "สินค้า : " & TempStr3 & " (" & TempStr4 & ")"
         End If
         
         PrevKey1 = TempStr3
            
         Call m_Details1.ClearField
         
         I = I + 1
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = I
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)


          J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = Di.DOCUMENT_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

          J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = DateToStringExtEx2(Di.DOCUMENT_DATE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
          J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = "(" & Di.CUSTOMER_CODE & ") " & Left(Di.CUSTOMER_NAME, 15)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      
         If Di.PART_TYPE = 21 Then 'ถ้าเป็น Bulk
             J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = FormatNumber(Di.ITEM_AMOUNT, -1)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(J) = Total1(J) + Di.ITEM_AMOUNT
            Total2(J) = Total2(J) + Di.ITEM_AMOUNT
             Total3(J) = -2
            AMOUNT = Di.ITEM_AMOUNT
         Else
             J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = FormatNumber(Di.PACK_AMOUNT, -1)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(J) = Total1(J) + Di.PACK_AMOUNT
            Total2(J) = Total2(J) + Di.PACK_AMOUNT
            Total3(J) = -2
            AMOUNT = Di.PACK_AMOUNT
         End If

          J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = FormatNumber(Di.PACKAGE_RATE, -1)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(J) = Total1(J) + (AMOUNT * Di.PACKAGE_RATE)
         Total2(J) = Total2(J) + (AMOUNT * Di.PACKAGE_RATE)
        Total3(J) = -1
         
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = IIf(Di.DISCOUNT_PER_PACK_PART > 0, "(" & FormatNumber(Di.DISCOUNT_PER_PACK_PART, -1) & ")", Di.DISCOUNT_PER_PACK_PART)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(J) = Total1(J) + (AMOUNT * Di.DISCOUNT_PER_PACK_PART)
            Total2(J) = Total2(J) + (AMOUNT * Di.DISCOUNT_PER_PACK_PART)
            Total3(J) = -3
 
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            If Di.DISCOUNT_AMOUNT_PART > 0 Then
               Amt = FormatNumber(Di.DISCOUNT_AMOUNT_PART - Di.DISCOUNT_PER_PACK_PART, -1)
            Else
               Amt = 0
            End If
           Amt = Amt + FormatNumber(Di.DISCOUNT_EXTRA_AMOUNT_PART, -1)
            
            TempStr = IIf(Amt > 0, "(" & FormatNumber(Amt, -1) & ")", Amt)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(J) = Total1(J) + (AMOUNT * Amt)
            Total2(J) = Total2(J) + (AMOUNT * Amt)
            Total3(J) = -3
'             J = J + 1
'            Set HeadCf = m_Headers1.Fields(J)
'            If Di.DISCOUNT_AMOUNT_PART > 0 Then
'               Amt = FormatNumber(Di.DISCOUNT_AMOUNT_PART - Di.DISCOUNT_PER_PACK_PART, -1)
'            Else
'               Amt = 0
'            End If
'            TempStr = IIf(Amt > 0, "(" & FormatNumber(Amt, -1) & ")", Amt)
'            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'            Call m_Details1.AddField(BodyCf)
'            Total1(J) = Total1(J) + (AMOUNT * Amt)
'            Total2(J) = Total2(J) + (AMOUNT * Amt)
'            Total3(J) = -3
         
          J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = FormatNumber(Di.RATE_CUSTOMER, -1)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(J) = Total1(J) + (AMOUNT * Di.RATE_CUSTOMER)
         Total2(J) = Total2(J) + (AMOUNT * Di.RATE_CUSTOMER)
         Total3(J) = -1
         
          J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = IIf(Di.DISCOUNT_PER_PACK_DLC > 0, "(" & FormatNumber(Di.DISCOUNT_PER_PACK_DLC, -1) & ")", Di.DISCOUNT_PER_PACK_DLC)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(J) = Total1(J) + (AMOUNT * Di.DISCOUNT_PER_PACK_DLC)
         Total2(J) = Total2(J) + (AMOUNT * Di.DISCOUNT_PER_PACK_DLC)
         Total3(J) = -3
          
          J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         If Di.DISCOUNT_AMOUNT_DLC > 0 Then
            Amt = FormatNumber(Di.DISCOUNT_AMOUNT_DLC - Di.DISCOUNT_PER_PACK_DLC, -1)
         Else
            Amt = 0
         End If
         TempStr = IIf(Amt > 0, "(" & FormatNumber(Amt, -1) & ")", Amt)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(J) = Total1(J) + (AMOUNT * Amt)
         Total2(J) = Total2(J) + (AMOUNT * Amt)
         Total3(J) = -3

         If Di.PART_TYPE = 21 Then 'ถ้าเป็น Bulk
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = FormatNumber(Di.PRO_COMMISSION_KG, -1)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(J) = Total1(J) + (AMOUNT * Di.PRO_COMMISSION_KG)
            Total2(J) = Total2(J) + (AMOUNT * Di.PRO_COMMISSION_KG)
            Total3(J) = -1
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = FormatNumber(Di.PRO_CHEER_KG, -1)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(J) = Total1(J) + (AMOUNT * Di.PRO_CHEER_KG)
            Total2(J) = Total2(J) + (AMOUNT * Di.PRO_CHEER_KG)
            Total3(J) = -1
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = FormatNumber(Di.PRO_DST_KG, -1)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(J) = Total1(J) + (AMOUNT * Di.PRO_DST_KG)
            Total2(J) = Total2(J) + (AMOUNT * Di.PRO_DST_KG)
            Total3(J) = -1
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = FormatNumber(Di.PRO_OTHER1_KG, -1)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(J) = Total1(J) + (AMOUNT * Di.PRO_OTHER1_KG)
            Total2(J) = Total2(J) + (AMOUNT * Di.PRO_OTHER1_KG)
            Total3(J) = -1
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = FormatNumber(Di.PRO_OTHER2_KG, -1)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(J) = Total1(J) + (AMOUNT * Di.PRO_OTHER2_KG)
            Total2(J) = Total2(J) + (AMOUNT * Di.PRO_OTHER2_KG)
            Total3(J) = -1
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = FormatNumber(Di.PRO_OTHER3_KG, -1)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(J) = Total1(J) + (AMOUNT * Di.PRO_OTHER3_KG)
            Total2(J) = Total2(J) + (AMOUNT * Di.PRO_OTHER3_KG)
            Total3(J) = -1
            
            Amt = Di.PRO_COMMISSION_KG + Di.PRO_CHEER_KG + Di.PRO_DST_KG + Di.PRO_OTHER1_KG + Di.PRO_OTHER2_KG + Di.PRO_OTHER3_KG
         Else
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = FormatNumber(Di.PRO_COMMISSION_BAG, -1)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(J) = Total1(J) + (AMOUNT * Di.PRO_COMMISSION_BAG)
            Total2(J) = Total2(J) + (AMOUNT * Di.PRO_COMMISSION_BAG)
            Total3(J) = -1

            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = FormatNumber(Di.PRO_CHEER_BAG, -1)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(J) = Total1(J) + (AMOUNT * Di.PRO_CHEER_BAG)
            Total2(J) = Total2(J) + (AMOUNT * Di.PRO_CHEER_BAG)
            Total3(J) = -1
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = FormatNumber(Di.PRO_DST_BAG, -1)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(J) = Total1(J) + (AMOUNT * Di.PRO_DST_BAG)
            Total2(J) = Total2(J) + (AMOUNT * Di.PRO_DST_BAG)
            Total3(J) = -1

            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = FormatNumber(Di.PRO_OTHER1_BAG, -1)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(J) = Total1(J) + (AMOUNT * Di.PRO_OTHER1_BAG)
            Total2(J) = Total2(J) + (AMOUNT * Di.PRO_OTHER1_BAG)
            Total3(J) = -1

            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = FormatNumber(Di.PRO_OTHER2_BAG, -1)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(J) = Total1(J) + (AMOUNT * Di.PRO_OTHER2_BAG)
            Total2(J) = Total2(J) + (AMOUNT * Di.PRO_OTHER2_BAG)
            Total3(J) = -1

            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = FormatNumber(Di.PRO_OTHER3_BAG, -1)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(J) = Total1(J) + (AMOUNT * Di.PRO_OTHER3_BAG)
            Total2(J) = Total2(J) + (AMOUNT * Di.PRO_OTHER3_BAG)
            Total3(J) = -1

            
            Amt = Di.PRO_COMMISSION_BAG + Di.PRO_CHEER_BAG + Di.PRO_DST_BAG + Di.PRO_OTHER1_BAG + Di.PRO_OTHER2_BAG + Di.PRO_OTHER3_BAG
         End If
            
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = DeliveryUnit(Di.PART_TYPE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = Di.PRICE_THINK_TYPE
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         
         Amt = Di.PACKAGE_RATE + Di.RATE_CUSTOMER + Amt - Di.DISCOUNT_AMOUNT_PART - Di.DISCOUNT_EXTRA_AMOUNT_PART - Di.DISCOUNT_AMOUNT_DLC
         
          J = J + 1
          Set HeadCf = m_Headers1.Fields(J)
         TempStr = FormatNumber(Amt, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(J) = Total1(J) + (AMOUNT * Amt)
         Total2(J) = Total2(J) + (AMOUNT * Amt)
         Total3(J) = -1
         
         If Di.PART_TYPE = 21 Then 'ถ้าเป็น Bag
             J = J + 1
             Set HeadCf = m_Headers1.Fields(J)
            TempStr = FormatNumber(Di.ITEM_AMOUNT * Amt, 2)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(J) = Total1(J) + Di.ITEM_AMOUNT * Amt
            Total2(J) = Total2(J) + Di.ITEM_AMOUNT * Amt
            Total3(J) = -2
         Else
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = FormatNumber(Di.PACK_AMOUNT * Amt, 2)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(J) = Total1(J) + Di.PACK_AMOUNT * Amt
            Total2(J) = Total2(J) + Di.PACK_AMOUNT * Amt
            Total3(J) = -2
         End If
         
             J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = Di.USER_APPLOVE_PRICE
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
         
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         
         End If
         Rs.MoveNext
      Wend

      Call GenerateFooter("รวม", "", "", Total1, m_Details1, CountDoc, Total3)
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False

      Vsp = ""

      Call GenerateFooter("สุทธิ", "", "", Total2, m_Details1, CountDoc2, Total3)
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      
      Vsp = "หมายเหตุ  1 = มารับเอง, 2 = รวมค่าขนส่ง, 3 = แยกค่าขนส่ง"

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

Private Sub GenerateFooter(TXT1 As String, TXT2 As String, TXT3 As String, Tot() As Double, Details As CFieldList, DocCount As Long, TotS() As Double)
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
      ElseIf J = 2 Then
         TempStr = TXT2
      ElseIf J = 3 Then
         TempStr = TXT3
      ElseIf TotS(J) = -1 Then
         TempStr = FormatNumber(MyDiffEx(Tot(J), Tot(5)))
      ElseIf TotS(J) = -2 Then
          TempStr = FormatNumber(Tot(J))
      ElseIf TotS(J) = -3 Then
         Amt = MyDiffEx(Tot(J), Tot(5))
         TempStr = IIf(Amt > 0, "(" & FormatNumber(MyDiffEx(Tot(J), Tot(5))) & ")", FormatNumber(Amt))
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




