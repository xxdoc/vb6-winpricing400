VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportPlanning001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportPlanning001"
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


Private m_PartTxtypes As Collection
Private m_InventoryActItem As Collection
Private m_PartTxtypeBas As Collection

Private m_PlanMonth As Collection
Private m_PlanUsePartWeek As Collection
Private m_PlanUsePartWeek2 As Collection
Private m_PlanUsePartDay As Collection
Private m_PlanGetPartDay As Collection
Private RoColl As Collection

Private FromDate As Date
Private ToDate As Date

Private FromDate2 As Date
Private ToDate2 As Date

Private ComeInColl As Collection  'Query จำนวนวัตถุดิบที่เข้ามาก่อนหน้าวันที่ต้องการออก report 1 วัน  เพื่อทำยอดยกมา และ ยอดสะสม

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_TempCol = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Details1 = New CFieldList
   Set m_PartTxtypes = New Collection
   Set m_InventoryActItem = New Collection
   Set m_PartTxtypeBas = New Collection
   Set Rs = New ADODB.Recordset
   
   Set m_PlanMonth = New Collection
   Set m_PlanUsePartWeek = New Collection
   Set m_PlanUsePartWeek2 = New Collection
   Set m_PlanUsePartDay = New Collection
   Set m_PlanGetPartDay = New Collection
   Set RoColl = New Collection
   Set ComeInColl = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_PartTxtypes = Nothing
   Set m_InventoryActItem = Nothing
   Set m_PartTxtypeBas = Nothing
   
   Set m_PlanMonth = Nothing
   Set m_PlanUsePartWeek = Nothing
   Set m_PlanUsePartWeek2 = Nothing
   Set m_PlanUsePartDay = Nothing
   Set m_PlanGetPartDay = Nothing
   Set RoColl = Nothing
   Set ComeInColl = Nothing
   
   If Not Rs Is Nothing Then
      If Rs.State = adStateOpen Then
         Rs.Close
      End If
      Set Rs = Nothing
   End If
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
Dim TempDate As Date
Dim TempWidth As Long
Dim TempSize As Double
   
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll

   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField

   Call Cf.SetFieldValue(65, "^", "", "^")
   Call m_Headers0.AddField(Cf)
   
'   Call Cf.SetFieldValue(20, "^", "Plan ใช้", "^")
'   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(15, "^", "C", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(50, "^", "R/M", "<")
   Call m_Headers1.AddField(Cf)
   
   If mcolParam("SHOW_PLAN_MONTH") = 1 Then
      TempSize = 0
      TempDate = FromDate
      While TempDate <= ToDate2
         Call Cf.SetFieldValue(20, "^", Month(TempDate) & "/" & Year(TempDate), ">")
         Call m_Headers1.AddField(Cf)
         TempSize = TempSize + 20
         TempDate = DateAdd("M", 1, TempDate)
      Wend
         
      Call Cf.SetFieldValue(TempSize, "^", "Plan เดือน", "^")
      Call m_Headers0.AddField(Cf)
      TempSize = 0
   Else
      Call Cf.SetFieldValue(0, "^", "Plan เดือน", "^")
      Call m_Headers0.AddField(Cf)
      
      Call Cf.SetFieldValue(0, "^", "", ">")
      Call m_Headers1.AddField(Cf)
      
   End If
   
   Call Cf.SetFieldValue(20, "^", "Plan ใช้", "^")
   Call m_Headers0.AddField(Cf)
   
   
   
   Call Cf.SetFieldValue(20, "^", Day(mcolParam("FROM_WEEK")) & "-" & Day(DateAdd("D", 6, mcolParam("FROM_WEEK"))), ">")
   Call m_Headers1.AddField(Cf)
   
   TempDate = mcolParam("FROM_WEEK_DATE")
   While TempDate <= mcolParam("TO_WEEK_DATE")
      TempWidth = 0
      Call Cf.SetFieldValue(20, "^", "Inv", ">")
      Call m_Headers1.AddField(Cf)
      TempWidth = TempWidth + 1
      

        If mcolParam("SHOW_INV_ACTUAL") = 1 Then
         Call Cf.SetFieldValue(20, "^", "InvA", ">")
         Call m_Headers1.AddField(Cf)
         TempWidth = TempWidth + 1
          
            If mcolParam("SHOW_DIFF_INV") = 1 Then
               Call Cf.SetFieldValue(20, "^", "Diff", ">")
               Call m_Headers1.AddField(Cf)
               TempWidth = TempWidth + 1
          End If
       End If
       
      If mcolParam("SHOW_PLAN_DAILY") = 1 Then
         Call Cf.SetFieldValue(20, "^", "Plan", ">")
         Call m_Headers1.AddField(Cf)
         TempWidth = TempWidth + 1
      End If
      
      If mcolParam("SHOW_ACT_DAILY") = 1 Then
         If (mcolParam("PRINT_DATE") > TempDate) Then
            Call Cf.SetFieldValue(20, "^", "Act", ">")
            Call m_Headers1.AddField(Cf)
            TempWidth = TempWidth + 1
         End If
      End If
      
      Call Cf.SetFieldValue(20 * TempWidth, "^", Day(TempDate), ">")
      Call m_Headers0.AddField(Cf)
      
      TempDate = DateAdd("D", 1, TempDate)
   Wend
   
    If mcolParam("SHOW_SUM_ACT_DAILY") = 1 Then
        ' If (mcolParam("PRINT_DATE") > TempDate) Then
            Call Cf.SetFieldValue(40, "^", Day(mcolParam("FROM_WEEK")) & "-" & Day(TempDate - 1), ">")
            Call m_Headers0.AddField(Cf)
            TempWidth = TempWidth + 1
         'End If
   End If
   
   If mcolParam("SHOW_SUM_ACT_DAILY") = 1 Then
      Call Cf.SetFieldValue(20, "^", "SAct", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(20, "^", "Diff", ">")
      Call m_Headers1.AddField(Cf)
   End If
   
   Call Cf.SetFieldValue(20, "^", Day(TempDate), ">")
   Call m_Headers0.AddField(Cf)
      
   Call Cf.SetFieldValue(20, "^", "Inv", ">")
   Call m_Headers1.AddField(Cf)
   
   If mcolParam("SHOW_USE_PLAN") = 1 Or mcolParam("SHOW_RX_ACTUAL") = 1 Then
      TempDate = mcolParam("FROM_WEEK_DATE")
      While TempDate <= mcolParam("TO_SUP_DATE")
         If mcolParam("SHOW_USE_PLAN") = 1 Then
            Call Cf.SetFieldValue(20, "^", "ปมก", ">")
            Call m_Headers1.AddField(Cf)
            TempSize = 20
         End If
          If mcolParam("SHOW_RX_ACTUAL") = 1 Then
            Call Cf.SetFieldValue(20, "^", "รจ", ">")
            Call m_Headers1.AddField(Cf)
            TempSize = TempSize + 20
          End If
         
         Call Cf.SetFieldValue(TempSize, "^", Day(TempDate) & " " & WeekdayName(Weekday(TempDate, vbMonday), True), ">")
         Call m_Headers0.AddField(Cf)
         TempSize = 0
         TempDate = DateAdd("D", 1, TempDate)
      Wend
      
      Call Cf.SetFieldValue(20, "^", "Plan ใช้", ">")
      Call m_Headers0.AddField(Cf)
         
      Call Cf.SetFieldValue(20, "^", Day(DateAdd("D", 7, mcolParam("FROM_WEEK"))) & "-" & Day(DateAdd("D", 13, mcolParam("FROM_WEEK"))), ">")
      Call m_Headers1.AddField(Cf)
   
   End If
      
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
   Vsp.ORIENTATION = orLandscape
   Vsp.MarginBottom = 300
   Vsp.MarginFooter = 300
   Vsp.MarginHeader = 100
   Vsp.MarginLeft = 300
   Vsp.MarginRight = 500
   Vsp.MarginTop = 100
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_WEEK_DATE")), "N/A")
   ToDate = "ถึง " & EmptyToString(DateToStringExtEx2(mcolParam("TO_WEEK_DATE")), "N/A")
   
   MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & vbCrLf & _
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

Private Sub AmtToWriteFlag(Amt As Double, TmpFlag As Boolean)
   If Not TmpFlag Then
      If Amt <> 0 Then
         TmpFlag = True
      Else
         TmpFlag = False
      End If
   End If
End Sub

Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim RName As String
Dim I As Long
Dim J As Long
Dim strForm_PlanningItemt As String
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
Dim Count1 As Double
Dim BalanceAccums As Collection
Dim BalanceLi As CLotItem
Dim LeftAmount1 As Double
Dim LeftAmount2 As Double
Dim TmpFlag As Boolean
Dim TempLiBa1 As CLotItem
Dim TempLiBa2 As CLotItem
Dim TempDate As Date
Dim TempPlanningItem As CPlanningItem
Dim m_PlanningItem As CPlanningItem
Dim TempInventoryActItem As CInventoryActItem
Dim TempRo As CSupItem
Dim TempComeIn  As CSupItem
Dim SumAct As Double
Dim PlanUseOfWeek As Double
Dim planVersion As Long




   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
If Not mcolParam("SHOW_DATE_OVER") = 1 Then
   'เช็คว่าเรียนได้ไม่เกิน 7 วัน
   If DateDiff("D", mcolParam("FROM_WEEK_DATE"), mcolParam("TO_WEEK_DATE")) > 7 Then
      mstrErrMsg = "ไม่สามารถใส่ช่วงวันที่ เกิน 7 วันได้"
      Exit Function
   End If
End If

If mcolParam("SHOW_PLAN_MONTH") = 1 Then
   Call GetFirstLastDate(DateSerial(mcolParam("FROM_YEAR") - 543, mcolParam("FROM_MONTH"), 1), FromDate, ToDate)
   Call GetFirstLastDate(DateSerial(mcolParam("TO_YEAR") - 543, mcolParam("TO_MONTH"), 1), FromDate2, ToDate2)

   Call LoadPlanMountAmount(Nothing, m_PlanMonth, FromDate, ToDate2, 4)
End If

   'ถ้าวันที่พิมพ์มากกว่าวันที่เคลื่อนไหว
   If (mcolParam("PRINT_DATE") > mcolParam("FROM_WEEK_DATE")) Then
      'ต้อง Load ยอดใช้วัตถุดิบมา
      Call LoadSumJobInputViaLotItem(Nothing, m_PartTxtypes, mcolParam("FROM_WEEK_DATE"), DateAdd("D", -1, mcolParam("TO_WEEK_DATE")), "LoadSumJobInputViaLotItem")
   End If
   Call LoadInventoryActual(Nothing, m_InventoryActItem, mcolParam("FROM_WEEK_DATE"), mcolParam("TO_WEEK_DATE"), "LoadInventoryActual")
   
   Call LoadPlanningItemDateAmount(m_PlanUsePartDay, mcolParam("FROM_WEEK_DATE"), mcolParam("TO_WEEK_DATE"), 1, 1, "LoadPlanningItemDateAmount")
   Call LoadPlanningItemDateAmount(m_PlanUsePartWeek, mcolParam("FROM_WEEK"), mcolParam("FROM_WEEK"), 2, 1, "LoadPlanningItemDateAmount")
   
   Call LoadPlanningItemDateAmount(m_PlanUsePartWeek2, DateAdd("D", 7, mcolParam("FROM_WEEK")), DateAdd("D", 7, mcolParam("FROM_WEEK")), 2, 1, "LoadPlanningItemDateAmount")
   
   Call LoadPlanningItemDateAmount(m_PlanGetPartDay, mcolParam("FROM_WEEK_DATE"), mcolParam("TO_SUP_DATE"), 3, 3, "LoadPlanningItemDateAmount")

  Call LoadSupItemComeIn(Nothing, RoColl, mcolParam("FROM_WEEK_DATE"), mcolParam("TO_WEEK_DATE"), , 1, "LoadSupItemComeIn") 'RoColl : ข้อมูล วันรับเข้าที่ต้องการ
  'planVersion = LoadPlanVersion(uctlPlanningDate.ShowDate, Trim(str(PlanningArea)))
   
   Set BalanceAccums = New Collection
   Set InventoryBals1 = New Collection
   If mcolParam("FROM_WEEK_DATE") > 0 Then        'รูปแบบคำนวณยอดยกมาในอนาคต
      Dim MonthlyAccums  As Collection
      Dim YYYYMM As String
      Dim firstDate As Date
      Dim lastDate As Date
      Set MonthlyAccums = New Collection
      Call GetFirstLastDate(mcolParam("FROM_WEEK_DATE"), firstDate, lastDate)
      YYYYMM = Format(Year(DateAdd("D", -1, firstDate)), "0000") & "-" & Format(Month(DateAdd("D", -1, firstDate)), "00")
      Call LoadMonthlyBalancePartItem(Nothing, MonthlyAccums, YYYYMM)
      Call glbDaily.CopyMonthlyAccumPartItem(MonthlyAccums, InventoryBals1)
      Set MonthlyAccums = Nothing
      
      If (firstDate <> mcolParam("FROM_WEEK_DATE")) Then
         Call LoadPartTxTypeAmount(Nothing, m_PartTxtypeBas, firstDate, DateAdd("D", -1, mcolParam("FROM_WEEK_DATE")))
      End If
   End If
   Set BalanceAccums = Nothing
   
   If Not initDoc Then Exit Function

   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         m_PlanningItemin Operation Here
'-----------------------------------------------------------------------------------------------------

   For J = 1 To UBound(Total1)
      Total1(J) = 0
   Next J

   Set m_PlanningItem = New CPlanningItem
   m_PlanningItem.FROM_DATE = mcolParam("FROM_WEEK")
   m_PlanningItem.TO_DATE = DateAdd("D", 7, mcolParam("FROM_WEEK"))

   If Not mcolParam("SHOW_RM_ALL") = 1 Then
      Call m_PlanningItem.QueryData(3, Rs, iCount)
   Else
      m_PlanningItem.FROM_DATE = -1
      m_PlanningItem.TO_DATE = -1
      m_PlanningItem.OrderType = 1
   Call m_PlanningItem.QueryData(4, Rs, iCount)
   End If
   
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
      I = 0
      
      Set BodyCf = New CReportField
      Set m_PlanningItem = New CPlanningItem
      While Not Rs.EOF
         TmpFlag = False
         I = I + 1
         PlanUseOfWeek = 0
         If Not mcolParam("SHOW_RM_ALL") = 1 Then
            Call m_PlanningItem.PopulateFromRS(3, Rs)
         Else
            Call m_PlanningItem.PopulateFromRS(4, Rs)
         End If
         Call m_Details1.ClearField
         
         Set HeadCf = m_Headers1.Fields(1)
         TempStr = m_PlanningItem.PART_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(2)
         If Val(mcolParam("LIMIT_PART_DESC")) > 0 Then
            TempStr = Left(m_PlanningItem.PART_DESC, Val(mcolParam("LIMIT_PART_DESC")))
         Else
            TempStr = Left(m_PlanningItem.PART_DESC, 15)
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(3)
         If mcolParam("SHOW_PLAN_MONTH") = 1 Then
            Dim SearchItemNo As CPlanningItem
            TempDate = FromDate
            While TempDate <= ToDate2
            TempStr = ""
            Set SearchItemNo = GetObject("CPlanningItem", m_PlanMonth, Trim(m_PlanningItem.PART_ITEM_ID & "-" & TempDate), False)
            If Not SearchItemNo Is Nothing Then
               TempStr = FormatNumber(SearchItemNo.PLAN_AMOUNT)
            End If
            'm_PlanMonth
               
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
               Call m_Details1.AddField(BodyCf)
                     
               TempDate = DateAdd("M", 1, TempDate)
            Wend
         Else
               TempStr = ""
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
               Call m_Details1.AddField(BodyCf)
         End If
         
         Set TempPlanningItem = GetObject("CPlanningItem", m_PlanUsePartWeek, m_PlanningItem.PART_ITEM_ID & "-" & mcolParam("FROM_WEEK"))
         Set HeadCf = m_Headers1.Fields(4)
         PlanUseOfWeek = TempPlanningItem.PLAN_AMOUNT / 1000
         If (Abs(TempPlanningItem.PLAN_AMOUNT / 1000) >= 1) Then
            TempStr = FormatNumberToNull(PlanUseOfWeek, 2)
         Else
            TempStr = FormatNumberToNull(PlanUseOfWeek, 3)
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(4) = Total1(4) + TempPlanningItem.PLAN_AMOUNT
         
         Set BalanceLi = GetLotItem(InventoryBals1, Trim(str(m_PlanningItem.PART_ITEM_ID)))
         Set TempLiBa1 = GetLotItem(m_PartTxtypeBas, m_PlanningItem.PART_ITEM_ID & "-" & "I")
         Set TempLiBa2 = GetLotItem(m_PartTxtypeBas, m_PlanningItem.PART_ITEM_ID & "-" & "E")
         
         
         SumAct = 0
         J = 4
         TempDate = mcolParam("FROM_WEEK_DATE")
         While TempDate <= mcolParam("TO_WEEK_DATE")
            J = J + 1
            
            Set HeadCf = m_Headers1.Fields(J)
            If J = 5 Then 'รอบแรก
               LeftAmount1 = BalanceLi.NEW_AMOUNT + TempLiBa1.TX_AMOUNT - TempLiBa2.TX_AMOUNT
            End If
            TempStr = FormatNumberToNull(LeftAmount1, 0)
            If ((LeftAmount1 / 1000) >= 1) Then
               TempStr = FormatNumberToNull(LeftAmount1 / 1000, 2)
            Else
               TempStr = FormatNumberToNull(LeftAmount1 / 1000, 3)
            End If
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            Total1(J) = Total1(J) + LeftAmount1
            Call AmtToWriteFlag(LeftAmount1, TmpFlag)
            
           
            If mcolParam("SHOW_INV_ACTUAL") = 1 Then
                Set TempInventoryActItem = GetObject("CInventoryActItem", m_InventoryActItem, m_PlanningItem.PART_ITEM_ID & "-" & TempDate)
                J = J + 1
                Set HeadCf = m_Headers1.Fields(J)
                If (Abs(TempInventoryActItem.INVENTORY_ACT_AMOUNT / 1000) >= 1) Then
                   TempStr = FormatNumberToNull(TempInventoryActItem.INVENTORY_ACT_AMOUNT / 1000, 2)
                Else
                   TempStr = FormatNumberToNull(TempInventoryActItem.INVENTORY_ACT_AMOUNT / 1000, 3)
                End If
                Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
                Call m_Details1.AddField(BodyCf)
               Total1(J) = Total1(J) + TempInventoryActItem.INVENTORY_ACT_AMOUNT
               
              If mcolParam("SHOW_DIFF_INV") = 1 Then
                  J = J + 1
                  Set HeadCf = m_Headers1.Fields(J)
                  TempStr = FormatNumberToNull((LeftAmount1 - TempInventoryActItem.INVENTORY_ACT_AMOUNT) / 1000, 2)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
                  Call m_Details1.AddField(BodyCf)
                  Total1(J) = Total1(J) + (LeftAmount1 - TempInventoryActItem.INVENTORY_ACT_AMOUNT)
              End If
           End If
               
          
            
            If mcolParam("SHOW_PLAN_DAILY") = 1 Then
               Set TempPlanningItem = GetObject("CPlanningItem", m_PlanUsePartDay, m_PlanningItem.PART_ITEM_ID & "-" & TempDate)
               J = J + 1
               If (Abs(TempPlanningItem.PLAN_AMOUNT / 1000) >= 1) Then
                  TempStr = FormatNumberToNull(TempPlanningItem.PLAN_AMOUNT / 1000, 2)
               Else
                  TempStr = FormatNumberToNull(TempPlanningItem.PLAN_AMOUNT / 1000, 3)
               End If
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
               Call m_Details1.AddField(BodyCf)
               Total1(J) = Total1(J) + TempPlanningItem.PLAN_AMOUNT
               Call AmtToWriteFlag(TempPlanningItem.PLAN_AMOUNT, TmpFlag)
           End If
            
            If (mcolParam("PRINT_DATE") > TempDate) Then
               Set TempLi2 = GetLotItem(m_PartTxtypes, m_PlanningItem.PART_ITEM_ID & "-" & TempDate)
'                  j = j + 1
               If ((TempLi2.TX_AMOUNT / 1000) >= 1) Then
                  TempStr = FormatNumberToNull(TempLi2.TX_AMOUNT / 1000, 2)
               Else
                  TempStr = FormatNumberToNull(TempLi2.TX_AMOUNT / 1000, 3)
               End If
               If mcolParam("SHOW_ACT_DAILY") = 1 Then
                   J = J + 1
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
                  Call m_Details1.AddField(BodyCf)
                  Total1(J) = Total1(J) + TempLi2.TX_AMOUNT
               End If
               SumAct = SumAct + TempLi2.TX_AMOUNT
               Call AmtToWriteFlag(TempLi2.TX_AMOUNT, TmpFlag)

               Set TempRo = GetObject("CSupItem", RoColl, Trim(m_PlanningItem.PART_ITEM_ID & "-" & TempDate))
               If TempRo.CALCULATE_TYPE = 1 Then
                  LeftAmount1 = LeftAmount1 + TempRo.SUPPLIER_WEIGHT - TempLi2.TX_AMOUNT
               ElseIf TempRo.CALCULATE_TYPE = 2 Then
                  LeftAmount1 = LeftAmount1 + TempRo.WEIGHT_AMOUNT - TempLi2.TX_AMOUNT
               ElseIf TempRo.CALCULATE_TYPE = 3 Then
                  LeftAmount1 = LeftAmount1 + TempRo.TOTAL_WEIGHT - TempLi2.TX_AMOUNT
               ElseIf TempRo.CALCULATE_TYPE = 0 Then
                  LeftAmount1 = LeftAmount1 - TempLi2.TX_AMOUNT
               Else
                   LeftAmount1 = LeftAmount1 - TempLi2.TX_AMOUNT
               End If
            Else
               LeftAmount1 = LeftAmount1 - TempPlanningItem.PLAN_AMOUNT
               Set TempPlanningItem = GetObject("CPlanningItem", m_PlanGetPartDay, m_PlanningItem.PART_ITEM_ID & "-" & TempDate)
               LeftAmount1 = LeftAmount1 + TempPlanningItem.PLAN_AMOUNT
            End If
            TempDate = DateAdd("D", 1, TempDate)
         Wend
         
           If mcolParam("SHOW_SUM_ACT_DAILY") = 1 Then
                   J = J + 1
                  TempStr = FormatNumberToNull(SumAct / 1000, 2)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
                  Call m_Details1.AddField(BodyCf)
                  Total1(J) = Total1(J) + SumAct
                  
                  J = J + 1
                  TempStr = FormatNumberToNull(PlanUseOfWeek - (SumAct / 1000), 2)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
                  Call m_Details1.AddField(BodyCf)
                  Total1(J) = Total1(J) + (PlanUseOfWeek - (SumAct / 1000)) 'Sum Diff
          End If
         
         
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         If ((LeftAmount1 / 1000) >= 1) Then
            TempStr = FormatNumberToNull(LeftAmount1 / 1000, 2)
         Else
            TempStr = FormatNumberToNull(LeftAmount1 / 1000, 3)
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(J) = Total1(J) + LeftAmount1
         Call AmtToWriteFlag(LeftAmount1, TmpFlag)
         
         If mcolParam("SHOW_USE_PLAN") = 1 Or mcolParam("SHOW_RX_ACTUAL") = 1 Then
            TempDate = mcolParam("FROM_WEEK_DATE")
              While TempDate <= mcolParam("TO_SUP_DATE")
                  If mcolParam("SHOW_USE_PLAN") = 1 Then
                     Set TempPlanningItem = GetObject("CPlanningItem", m_PlanGetPartDay, m_PlanningItem.PART_ITEM_ID & "-" & TempDate)
                     J = J + 1
                     If (Abs(TempPlanningItem.PLAN_AMOUNT / 1000) >= 1) Then
                        TempStr = FormatNumberToNull(TempPlanningItem.PLAN_AMOUNT / 1000, 2)
                     Else
                        TempStr = FormatNumberToNull(TempPlanningItem.PLAN_AMOUNT / 1000, 3)
                     End If
                     Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
                     Call m_Details1.AddField(BodyCf)
                     Total1(J) = Total1(J) + TempPlanningItem.PLAN_AMOUNT
                     Call AmtToWriteFlag(TempPlanningItem.PLAN_AMOUNT, TmpFlag)
                  End If
                  
               'แสดงยอดรับจริง
               If mcolParam("SHOW_RX_ACTUAL") = 1 Then
               Set TempRo = GetObject("CSupItem", RoColl, Trim(m_PlanningItem.PART_ITEM_ID & "-" & TempDate))
                J = J + 1
               If TempRo.CALCULATE_TYPE = 1 Then
                  LeftAmount1 = TempRo.SUPPLIER_WEIGHT
               ElseIf TempRo.CALCULATE_TYPE = 2 Then
                  LeftAmount1 = TempRo.WEIGHT_AMOUNT
               ElseIf TempRo.CALCULATE_TYPE = 3 Then
                  LeftAmount1 = TempRo.TOTAL_WEIGHT
               ElseIf TempRo.CALCULATE_TYPE = 0 Then
                  LeftAmount1 = 0
               Else
                   LeftAmount1 = 0
               End If

               TempStr = FormatNumberToNull(LeftAmount1 / 1000, 2)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
               Call m_Details1.AddField(BodyCf)
               Total1(J) = Total1(J) + LeftAmount1
               Call AmtToWriteFlag(LeftAmount1, TmpFlag)
               End If
               'จบการแสดงยอดรับจริง
               
               
               TempDate = DateAdd("D", 1, TempDate)
            Wend
            
            J = J + 1
            Set TempPlanningItem = GetObject("CPlanningItem", m_PlanUsePartWeek2, m_PlanningItem.PART_ITEM_ID & "-" & DateAdd("D", 7, mcolParam("FROM_WEEK")))
            Set HeadCf = m_Headers1.Fields(J)
            If (Abs(TempPlanningItem.PLAN_AMOUNT / 1000) >= 1) Then
               TempStr = FormatNumberToNull(TempPlanningItem.PLAN_AMOUNT / 1000, 2)
            Else
               TempStr = FormatNumberToNull(TempPlanningItem.PLAN_AMOUNT / 1000, 3)
            End If
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            Total1(J) = Total1(J) + TempPlanningItem.PLAN_AMOUNT
         End If
         If Not mcolParam("SHOW_RM_ALL") = 1 Then
            If TmpFlag Then
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strForm_PlanningItemt = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strForm_PlanningItemt, "", TempStr2)
            Else
               I = I - 1
            End If
         Else
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strForm_PlanningItemt = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strForm_PlanningItemt, "", TempStr2)
        End If
'         =====
         
         Rs.MoveNext
      Wend

      Call GenerateFooter("", "รวม", "", Total1, m_Details1)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strForm_PlanningItemt = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strForm_PlanningItemt, "", TempStr2)
      Vsp.FontBold = False

      Set BodyCf = Nothing
      Set m_PlanningItem = Nothing
      
   Vsp.EndDoc
   Call CloseExportFile(Vsp)

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
'         TempStr = TXT1
      ElseIf J = 2 Then
         TempStr = TXT1
      Else
         If ((Tot(J) / 1000) >= 1) Then
            TempStr = FormatNumberToNull(Tot(J) / 1000, 2)
         Else
            TempStr = FormatNumberToNull(Tot(J) / 1000, 3)
         End If
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


