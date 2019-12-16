VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportInventory004_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportInventory004_4"
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
Private m_Details1 As CFieldList
Private m_Left As Double
Private m_PartLocationTxtypes As Collection
Private m_PartLocations1 As Collection
Private m_PartLocations2 As Collection
Private m_PartTxtypeBas As Collection

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_TempCol = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Details1 = New CFieldList
   Set m_PartLocationTxtypes = New Collection
   Set m_PartLocations1 = New Collection
   Set m_PartLocations2 = New Collection
   Set m_PartTxtypeBas = New Collection

   Set Rs = New ADODB.Recordset
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_PartLocationTxtypes = Nothing
   Set m_PartLocations1 = Nothing
   Set m_PartLocations2 = Nothing
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
   
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField

   Call Cf.SetFieldValue(13, "^", "", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(8, "^", "จำนวน", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(10, "^", "มูลค่า", "^")
   Call m_Headers0.AddField(Cf)

   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "รหัส", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "ชื่อ", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "สถานที่จัดเก็บ", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ยกมา", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "รับเข้า", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จ่ายออก", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "คงเหลือ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ยกมา", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "รับเข้า", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จ่ายออก", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "คงเหลือ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ราคา @", ">")
   Call m_Headers1.AddField(Cf)
      
   Set Cf = Nothing
   
  Call m_Headers0.GetString(1, TempStr1, TempStr2)
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
   Vsp.ORIENTATION = orPortrait
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
Dim cData As CLotItem
Dim I As Long
Dim J As Long
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
Dim TempLi1 As CLotItem
Dim TempLi2 As CLotItem
Dim Sum1 As Double
Dim Sum2 As Double
Dim Count1 As Double
Dim BalanceAccums As Collection
Dim BalanceLi As CLotItem
Dim LeftAmount1 As Double
Dim LeftAmount2 As Double
Dim LeftAmount3 As Double
Dim InventoryBals2 As Collection
Dim BalanceLi2 As CLotItem
Dim TempLiBa1 As CLotItem
Dim TempLiBa2 As CLotItem
Dim EndDate As Date

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Call LoadPartLocationTxTypeAmount(Nothing, m_PartLocationTxtypes, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), , mcolParam("LOCATION_ID"), mcolParam("PART_TYPE"))
   
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
      Call LoadMonthlyBalance(Nothing, MonthlyAccums, YYYYMM, mcolParam("LOCATION_ID"))
      Call glbDaily.CopyMonthlyAccum(MonthlyAccums, InventoryBals1)
      Set MonthlyAccums = Nothing
      If (firstDate <> mcolParam("FROM_DATE")) Then
         Call LoadPartLocationTxTypeAmount(Nothing, m_PartTxtypeBas, firstDate, DateAdd("D", -1, mcolParam("FROM_DATE")), , mcolParam("LOCATION_ID"))
      End If
   End If
   Set BalanceAccums = Nothing
   
   Set BalanceAccums = New Collection
   Set InventoryBals2 = New Collection
   If mcolParam("TO_DATE") > 0 Then
      Set MonthlyAccums = New Collection
      YYYYMM = Format(Year(mcolParam("TO_DATE")), "0000") & "-" & Format(Month(mcolParam("TO_DATE")), "00")
      Call LoadMonthlyBalance(Nothing, MonthlyAccums, YYYYMM, mcolParam("LOCATION_ID"))
      Call glbDaily.CopyMonthlyAccum(MonthlyAccums, InventoryBals2)
      Set MonthlyAccums = Nothing
   End If
   Set BalanceAccums = Nothing
   
   Dim Li As CMonthlyAccum
   
   Set Li = New CMonthlyAccum
   Li.PART_ITEM_ID = -1
   Li.YYYYMM = YYYYMM
   Li.LOCATION_ID = mcolParam("LOCATION_ID")
   Li.PART_TYPE = mcolParam("PART_TYPE")
   Li.PART_GROUP = mcolParam("PART_GROUP")
   Li.FROM_PART_NO = mcolParam("FROM_PART_NO")
   Li.TO_PART_NO = mcolParam("TO_PART_NO")
   Li.OrderBy = 1
   Call Li.QueryData(3, Rs, iCount, False)
   
   If Not initDoc Then Exit Function

   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------

   For J = 1 To UBound(Total1)
      Total1(J) = 0
   Next J
   

   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
   
      I = 0

      Set BodyCf = New CReportField
      While Not (Rs.EOF)
         Call Li.PopulateFromRS(3, Rs)
         I = I + 1
         Sum1 = 0
         Sum2 = 0
         
         Set TempLi1 = GetLotItem(m_PartLocationTxtypes, Li.LOCATION_ID & "-" & Li.PART_ITEM_ID & "-" & "I")
         Set TempLi2 = GetLotItem(m_PartLocationTxtypes, Li.LOCATION_ID & "-" & Li.PART_ITEM_ID & "-" & "E")
         Set BalanceLi = GetLotItem(InventoryBals1, Li.LOCATION_ID & "-" & Li.PART_ITEM_ID)
         Set BalanceLi2 = GetLotItem(InventoryBals2, Li.LOCATION_ID & "-" & Li.PART_ITEM_ID)
         
         Set TempLiBa1 = GetLotItem(m_PartTxtypeBas, Li.LOCATION_ID & "-" & Li.PART_ITEM_ID & "-" & "I")
         Set TempLiBa2 = GetLotItem(m_PartTxtypeBas, Li.LOCATION_ID & "-" & Li.PART_ITEM_ID & "-" & "E")
         
         Call m_Details1.ClearField

         Set HeadCf = m_Headers1.Fields(1)
         TempStr = I
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(2)
         TempStr = Li.PART_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(3)
         TempStr = Li.PART_DESC
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(4)
         'TempStr = Li.LOCATION_NAME
         TempStr = Li.LOCATION_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)

         LeftAmount1 = BalanceLi.NEW_AMOUNT + TempLiBa1.TX_AMOUNT - TempLiBa2.TX_AMOUNT
         Set HeadCf = m_Headers1.Fields(5)
         TempStr = FormatNumber(LeftAmount1)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(5) = Total1(5) + LeftAmount1
         Sum1 = Sum1 + Abs(LeftAmount1)
         
         Set HeadCf = m_Headers1.Fields(6)
         TempStr = FormatNumber(TempLi1.TX_AMOUNT)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(6) = Total1(6) + TempLi1.TX_AMOUNT
         Sum1 = Sum1 + Abs(TempLi1.TX_AMOUNT)

         Set HeadCf = m_Headers1.Fields(7)
         TempStr = FormatNumber(TempLi2.TX_AMOUNT)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(7) = Total1(7) + TempLi2.TX_AMOUNT
         Sum1 = Sum1 + Abs(TempLi2.TX_AMOUNT)

         LeftAmount3 = (LeftAmount1 + TempLi1.TX_AMOUNT) - TempLi2.TX_AMOUNT
         Set HeadCf = m_Headers1.Fields(8)
         TempStr = FormatNumber(LeftAmount3)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(8) = Total1(8) + LeftAmount3
         Sum1 = Sum1 + Abs(LeftAmount3)
         
         Sum1 = BalanceLi.TOTAL_INCLUDE_PRICE + TempLiBa1.TOTAL_INCLUDE_PRICE - TempLiBa2.TOTAL_INCLUDE_PRICE
         Set HeadCf = m_Headers1.Fields(9)
        TempStr = FormatNumber(Sum1)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(9) = Total1(9) + Sum1
         Sum2 = Sum2 + Abs(FormatNumber(Sum1))

         Set HeadCf = m_Headers1.Fields(10)
         TempStr = FormatNumber(TempLi1.TOTAL_INCLUDE_PRICE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(10) = Total1(10) + TempLi1.TOTAL_INCLUDE_PRICE
         Sum2 = Sum2 + Abs(TempLi1.TOTAL_INCLUDE_PRICE)

         Set HeadCf = m_Headers1.Fields(11)
         TempStr = FormatNumber(TempLi2.TOTAL_INCLUDE_PRICE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(11) = Total1(11) + TempLi2.TOTAL_INCLUDE_PRICE
         Sum2 = Sum2 + Abs(TempLi2.TOTAL_INCLUDE_PRICE)

         LeftAmount2 = (Sum1 + TempLi1.TOTAL_INCLUDE_PRICE) - TempLi2.TOTAL_INCLUDE_PRICE
         Set HeadCf = m_Headers1.Fields(12)
         TempStr = FormatNumber(LeftAmount2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(12) = Total1(12) + LeftAmount2
         Sum2 = Sum2 + Abs(LeftAmount2)

         Set HeadCf = m_Headers1.Fields(13)
         TempStr = FormatNumber(BalanceLi2.NEW_PRICE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(13) = Total1(13) + 0

         If (Sum1 <> 0) Or (Sum2 <> 0) Then
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
      ElseIf J = 4 Then
         TempStr = ""
      ElseIf J = 13 Then
         TempStr = ""
      Else
         TempStr = FormatNumber(Tot(J))
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


