VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportLD001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportLD001"
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
Private m_ReportConfig As CReportConfig

Private Const TITLE_SIZE = 14
Private Const HEADER_SIZE = 12
Private Const DETAIL_SIZE = 10

Private Const TITLE_FORMAT = "<30|>3|<5;"

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
Public ID As Long
Private DOCUMENT_NO As String
Private CUSTOMER_NAME As String
Private TRUCK_NO As String
Private DOCUMENT_DATE As String
Private DOCUMENT_DESC As String
Private CONDITION1 As Long
Private CONDITION2 As Long
Private CONDITION3 As Long
Private EMP_CHECK_CAR_ID As Long
Private EMP_CHECK_CAR_NAME As String
Private EMP_CHECK_CAR_LNAME As String
Private EMP_CHECK_PRODUCT_NAME As String
Private EMP_CHECK_PRODUCT_LNAME As String
Private EXIT_WEIGHT As Double
Private ENTRY_WEIGHT As Double
Private TOTAL_WEIGHT As Double
Private WEIGHT_ID As String
Private ENTRY_WEIGHT_DATE As String
Private ENTRY_WEIGHT_TIME As String
Private EXIT_WEIGHT_DATE As String
Private EXIT_WEIGHT_TIME As String

Private TotalPack As Double
Private TotalWeight As Double
Private isNewPage As Boolean
Private TpCY As Double

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_TempCol = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers0_1 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Details1 = New CFieldList
   Set m_PartTxtypes = New Collection
   Set m_ReportConfig = New CReportConfig
   
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
   Set m_ReportConfig = Nothing
   
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
Dim TempC As Double
   
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = False
   Vsp.TableBorder = tbNone
   Vsp.TextAlign = taLeftBottom

      TempC = Vsp.CurrentY + 150
      
     Vsp.CurrentY = TempC
     Vsp.FontSize = 18
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 2
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcColWidth, 1, 1) = "22.3cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "3.5cm"
      Vsp.TableCell(tcAlign, 1, 2) = taLeftBottom
      Vsp.TableCell(tcText, 1, 2) = DOCUMENT_NO
      Vsp.TableBorder = tbNone
      Vsp.EndTable
      
      Vsp.CurrentY = Vsp.CurrentY - 110
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 4
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight, 1) = 1 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "2.4cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "11.8cm"
      Vsp.TableCell(tcColWidth, 1, 3) = "8.1cm"
      Vsp.TableCell(tcColWidth, 1, 4) = "3cm"
      Vsp.TableCell(tcAlign, 1, 2) = taLeftBottom
      Vsp.TableCell(tcAlign, 1, 3) = taLeftBottom
      Vsp.TableCell(tcAlign, 1, 4) = taLeftBottom
      Vsp.TableCell(tcText, 1, 2) = CUSTOMER_NAME
      Vsp.TableCell(tcText, 1, 3) = TRUCK_NO
      Vsp.TableCell(tcText, 1, 4) = DOCUMENT_DATE
      Vsp.TableBorder = tbNone
      Vsp.EndTable
      
      If Len(WEIGHT_ID) > 0 And Not WEIGHT_ID = "TW001" Then
       TpCY = Vsp.CurrentY
      Vsp.FontBold = True
       Vsp.CurrentY = TempC + 50
      Vsp.FontSize = 15
       Vsp.StartTable
       Vsp.TableCell(tcCols) = 2
       Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcColWidth, 1, 1) = "19.2cm"
       Vsp.TableCell(tcColWidth, 1, 2) = "3.5cm"
       Vsp.TableCell(tcAlign, 1, 1) = taRightBottom
       Vsp.TableCell(tcAlign, 1, 2) = taLeftBottom
       Vsp.TableCell(tcText, 1, 1) = "ใบชั่ง :"
       Vsp.TableCell(tcText, 1, 2) = Trim(WEIGHT_ID)
       Vsp.TableBorder = tbNone
       Vsp.EndTable
       Vsp.FontBold = False
       Vsp.CurrentY = TpCY
      End If
      
    If Len(DOCUMENT_DESC) > 0 Then
    TpCY = Vsp.CurrentY
      Vsp.FontSize = 16
      Vsp.FontBold = True
      Vsp.CurrentY = TempC - 700
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 2
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight, 1) = 800 '1 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "2.4cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "22.9cm"
      Vsp.TableCell(tcAlign, 1, 1) = taRightTop
      Vsp.TableCell(tcAlign, 1, 2) = taLeftTop
      Vsp.TableCell(tcText, 1, 1) = "* หมายเหตุ : "
      Vsp.TableCell(tcText, 1, 2) = DOCUMENT_DESC
      Vsp.TableBorder = tbNone
      Vsp.EndTable
      Vsp.FontBold = False
      Vsp.CurrentY = TpCY
   End If
      
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
   Vsp.PaperSize = pprLetter
   Vsp.ORIENTATION = orLandscape
   Vsp.MarginBottom = 0
   Vsp.MarginFooter = 0
   Vsp.MarginHeader = 0
   Vsp.MarginLeft = 0
   Vsp.MarginRight = 0
   Vsp.MarginTop = 0
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE

   MODULE_DESC = ""
   'Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))
   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"), m_ReportConfig, False)
   
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
Dim num As Long
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
Dim m_InventoryWHDoc As CInventoryWHDoc
Dim LotItemWh As CLotItemWH
Dim LD As CLotDoc
Dim PD As CPalletDoc
Dim Count1 As Double
Dim TempY As Double
Dim TempX As Double
Dim CountLT As Long
Dim CountPD As Long
Dim TempCY As Double
Dim Arr(10) As String

Arr(0) = "1.5cm"
Arr(1) = "1.36cm"
Arr(2) = "1.32cm"
Arr(3) = "1.32cm"
Arr(4) = "1.32cm"
Arr(5) = "1.31cm"
Arr(6) = "1.31cm"
Arr(7) = "1.30cm"
Arr(8) = "1.32cm"
Arr(9) = "1.30cm"
Dim ArrH(10) As Long
ArrH(0) = 660 '660
ArrH(1) = 550
ArrH(2) = 570
ArrH(3) = 570
ArrH(4) = 570
ArrH(5) = 570
ArrH(6) = 570
ArrH(7) = 570
ArrH(8) = 570
ArrH(9) = 570
Dim TempCount As Long
Dim CountRow As Long

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   If Not initDoc Then Exit Function

   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
   TotalPack = 0
   TotalWeight = 0
   For J = 1 To UBound(Total1)
      Total1(J) = 0
   Next J

    Set m_InventoryWHDoc = New CInventoryWHDoc
    m_InventoryWHDoc.INVENTORY_WH_DOC_ID = mcolParam("INVENTORY_WH_DOC_ID")
    m_InventoryWHDoc.QueryFlag = 1 'ถ้าเป็น 0 ยังไม่ต้องเอาลูก หลาน แต่ถ้าเป็น 1 เอา ลูกหลานด้วย
   
    If Not glbDaily.QueryInventoryWhDocForLG(m_InventoryWHDoc, Rs, iCount, IsOK, glbErrorLog) Then
       glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
       Call EnableForm(Me, True)
       Exit Function
    End If
    
    Call m_InventoryWHDoc.PopulateFromRS(1, Rs)
    
    DOCUMENT_NO = m_InventoryWHDoc.DOCUMENT_NO
    CUSTOMER_NAME = m_InventoryWHDoc.CUSTOMER_NAME
    TRUCK_NO = m_InventoryWHDoc.TRUCK_NO
    DOCUMENT_DATE = DateToStringExtEx2(m_InventoryWHDoc.DOCUMENT_DATE)
    DOCUMENT_DESC = m_InventoryWHDoc.NOTE
    CONDITION1 = m_InventoryWHDoc.CONDITION1
    CONDITION2 = m_InventoryWHDoc.CONDITION2
    CONDITION3 = m_InventoryWHDoc.CONDITION3
    EMP_CHECK_CAR_NAME = m_InventoryWHDoc.EMP_CHECK_CAR_NAME
    EMP_CHECK_CAR_LNAME = m_InventoryWHDoc.EMP_CHECK_CAR_LNAME
    EMP_CHECK_PRODUCT_NAME = m_InventoryWHDoc.EMP_CHECK_PRODUCT_NAME
    EMP_CHECK_PRODUCT_LNAME = m_InventoryWHDoc.EMP_CHECK_PRODUCT_LNAME
    EXIT_WEIGHT = m_InventoryWHDoc.EXIT_WEIGHT
    ENTRY_WEIGHT = m_InventoryWHDoc.ENTRY_WEIGHT
    TOTAL_WEIGHT = m_InventoryWHDoc.TOTAL_WEIGHT
    WEIGHT_ID = m_InventoryWHDoc.WEIGHT_ID
    ENTRY_WEIGHT_DATE = m_InventoryWHDoc.ENTRY_WEIGHT_DATE
    ENTRY_WEIGHT_TIME = m_InventoryWHDoc.ENTRY_WEIGHT_TIME
    EXIT_WEIGHT_DATE = m_InventoryWHDoc.EXIT_WEIGHT_DATE
    EXIT_WEIGHT_TIME = m_InventoryWHDoc.EXIT_WEIGHT_TIME

   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
   
   I = 0
   num = 0
   Vsp.TextAlign = taLeftBottom
  
   TempCY = Vsp.CurrentY
   TempY = Vsp.CurrentY
   CountRow = 0
   For Each LotItemWh In m_InventoryWHDoc.C_LotItemsWH
      I = I + 1
      J = 0
      num = 0
      CountLT = 0
      CountRow = CountRow + 1
      If CountRow = 1 Then
         TempCY = TempY + ArrH(CountRow - 1)
      Else
         TempCY = TempCY + ArrH(CountRow - 1)
      End If
      If CountRow > 10 Then
         CountRow = 1
        Call GenerateFooter
        Vsp.NewPage
        TempCY = TempY + ArrH(CountRow - 1)
      End If
         
      Call m_Details1.ClearField
      TempX = Vsp.CurrentX
   
      Vsp.FontSize = 16
      Vsp.FontBold = False

      Vsp.CurrentY = TempCY
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 1
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight, 1) = 1 * 650
      Vsp.TableCell(tcColWidth, 1, 1) = "1.1cm"
      Vsp.TableCell(tcAlign, 1, 1) = taCenterMiddle
      Vsp.TableCell(tcText, 1, 1) = I
      Vsp.TableBorder = tbNone
      Vsp.EndTable
      
      If Len(LotItemWh.PART_NO) <= 23 Then
         Vsp.FontSize = 20
      ElseIf Len(LotItemWh.PART_NO) <= 27 Then
         Vsp.FontSize = 18
      ElseIf Len(LotItemWh.PART_NO) > 27 Then
         Vsp.FontSize = 17
      End If
      Vsp.CurrentY = TempCY
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 2
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight, 1) = 1 * 650
      Vsp.TableCell(tcColWidth, 1, 1) = "0.9cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "6.3cm"
      Vsp.TableCell(tcAlign, 1, 1) = taCenterMiddle
      Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
      If LotItemWh.PART_ITEM_ID = -1 Then
         Vsp.TableCell(tcText, 1, 2) = LotItemWh.FEATURE_DESC
      Else
         Vsp.TableCell(tcText, 1, 2) = LotItemWh.PART_NO
      End If
      Vsp.TableBorder = tbNone
      Vsp.EndTable
      
      Vsp.FontSize = 18
      Vsp.CurrentY = TempCY
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 2
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight, 1) = 1 * 650
      Vsp.TableCell(tcColWidth, 1, 1) = "6.8cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "1.15cm"
      Vsp.TableCell(tcAlign, 1, 2) = taRightMiddle
      Vsp.TableCell(tcText, 1, 2) = LotItemWh.PACK_AMOUNT
      Vsp.TableBorder = tbNone
      Vsp.EndTable
      TotalPack = TotalPack + LotItemWh.PACK_AMOUNT
      TotalWeight = TotalWeight + (LotItemWh.PACK_AMOUNT * LotItemWh.WEIGHT_PER_PACK)
      
         For Each LD In LotItemWh.C_LotDoc
            CountLT = CountLT + 1
            If CountLT > 1 Then
               CountLT = 1
               CountRow = CountRow + 1
               TempCY = TempCY + ArrH(CountRow - 1)
               If CountRow > 10 Then
                  CountRow = 1
                  Call GenerateFooter
                  Vsp.NewPage
                  TempCY = TempY + ArrH(CountRow - 1)
               End If
            End If
            
            Vsp.FontSize = 16
            Vsp.CurrentY = TempCY - 60
             Vsp.StartTable
            Vsp.TableCell(tcCols) = 2
            Vsp.TableCell(tcRows) = 1
            Vsp.TableCell(tcRowHeight, 1) = 0.6 * 490 '567
            Vsp.TableCell(tcColWidth, 1, 1) = "7.9cm"
            Vsp.TableCell(tcColWidth, 1, 2) = "2.2cm"
            Vsp.TableCell(tcAlign, 1, 2) = taCenterMiddle
            If LD.DOCUMENT_TYPE = 15 Or LD.DOCUMENT_TYPE = 16 Then
               Vsp.TableCell(tcText, 1, 2) = DateToStringExtEx2(LD.BL_START_DATE)
            Else
               Vsp.TableCell(tcText, 1, 2) = DateToStringExtEx2(LD.START_DATE)
            End If
            Vsp.TableBorder = tbNone
            Vsp.EndTable
  
            Vsp.FontSize = 26
            Vsp.CurrentY = TempCY + 50
             Vsp.StartTable
            Vsp.TableCell(tcCols) = 2
            Vsp.TableCell(tcRows) = 1
            Vsp.TableCell(tcRowHeight, 1) = 0.6 * 567 + 300
            Vsp.TableCell(tcColWidth, 1, 1) = "7.67cm"
            Vsp.TableCell(tcColWidth, 1, 2) = "2.5cm"
            Vsp.TableCell(tcAlign, 1, 2) = taCenterMiddle
            Vsp.TableCell(tcText, 1, 2) = Right(LD.LOT_NO, 3)
            Vsp.TableBorder = tbNone
            Vsp.EndTable
            
            Vsp.CurrentY = TempCY
            Vsp.FontSize = 12
            Vsp.StartTable
            Vsp.TableCell(tcCols) = 2
            Vsp.TableCell(tcRows) = 1
            Vsp.TableCell(tcRowHeight, 1) = 1 * 567
            Vsp.TableCell(tcColWidth, 1, 1) = "9.8cm"
            Vsp.TableCell(tcColWidth, 1, 2) = "1.5cm"
            Vsp.TableCell(tcAlign, 1, 2) = taCenterMiddle
            Vsp.TableCell(tcText, 1, 2) = "" ' LD.LOCK_NAME
            Vsp.TableBorder = tbNone
            Vsp.EndTable
            '***************
          If CountItem(LD.C_PalletDoc) > 0 Then
            Vsp.StartTable
            Vsp.TableCell(tcCols) = 11
            Vsp.TableCell(tcRows) = 2
            Vsp.TableCell(tcRowHeight, 1) = 0.5 * 567
            Vsp.TableCell(tcRowHeight, 2) = 0.5 * 567
            Vsp.TableCell(tcColWidth, 1, 1) = "10.9cm"
            Vsp.TableCell(tcColWidth, 2, 1) = "10.9cm"

            num = 2
            CountPD = 1
            For Each PD In LD.C_PalletDoc
            Vsp.FontSize = 13
            Vsp.CurrentY = TempCY - 15
               If num > 11 Then
                  num = 2
                  Vsp.TableBorder = tbNone
                  Vsp.EndTable

                  CountRow = CountRow + 1
                  TempCY = TempCY + ArrH(CountRow - 1)
                  If CountRow > 10 Then
                     CountRow = 1
                     Call GenerateFooter
                     Vsp.NewPage
                     TempCY = TempY + ArrH(CountRow - 1)
                  End If
                  
                  Vsp.CurrentY = TempCY
                  Vsp.StartTable
                  Vsp.TableCell(tcCols) = 11
                  Vsp.TableCell(tcRows) = 2
                  Vsp.TableCell(tcRowHeight, 1) = 0.5 * 567
                  Vsp.TableCell(tcRowHeight, 2) = 0.5 * 567
                  Vsp.TableCell(tcColWidth, 1, 1) = "10.9cm"
                  Vsp.TableCell(tcColWidth, 2, 1) = "10.9cm"
               End If

                  Vsp.TableCell(tcColWidth, 1, num) = Arr(num - 2)
                  Vsp.TableCell(tcAlign, 1, num) = taCenterMiddle
                  Vsp.TableCell(tcText, 1, num) = LD.BIN_NAME

                  Vsp.TableCell(tcColWidth, 2, num) = Arr(num - 2)
                  Vsp.TableCell(tcAlign, 2, num) = taCenterMiddle
                  Vsp.TableCell(tcText, 2, num) = PD.PALLET_DOC_NO & "/" & PD.CAPACITY_AMOUNT

                  If CountPD >= CountItem(LD.C_PalletDoc) Then
                     Vsp.TableBorder = tbNone
                     Vsp.EndTable
                     isNewPage = True
                     If isNewPage Then
                        CountRow = CountRow - 1
                        isNewPage = False
                     End If
                     CountRow = CountRow + 1
                     If CountRow > 10 Then
                     CountRow = 1
                     Call GenerateFooter
                     Vsp.NewPage
                     TempCY = TempY + ArrH(CountRow - 1)

                     End If
                  End If

                  num = num + 1
                  CountPD = CountPD + 1
            Next PD
            End If
            Vsp.FontSize = 13
         Next LD
    
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Call Vsp.AddTable(strFormat, "", TempStr2)
   Next LotItemWh
      
   Call GenerateFooter
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   Set m_InventoryWHDoc = Nothing
   genDoc = True
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
   Set Rs = Nothing
End Function

Private Sub GenerateFooter()
Dim TempCY As Double

   Vsp.CurrentY = m_ReportConfig.HEAD_OFFSET + 8580
   Vsp.FontBold = True
    Vsp.FontSize = 15
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 3
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcRowHeight, 1) = 1 * 567
   Vsp.TableCell(tcColWidth, 1, 1) = "6.9cm"
   Vsp.TableCell(tcColWidth, 1, 2) = "1.1cm"
   Vsp.TableCell(tcColWidth, 1, 3) = "2cm"
   Vsp.TableCell(tcAlign, 1, 1) = taRightTop
   Vsp.TableCell(tcAlign, 1, 2) = taRightTop
   Vsp.TableCell(tcAlign, 1, 3) = taRightTop
   Vsp.TableCell(tcText, 1, 1) = "รวม"
   Vsp.TableCell(tcText, 1, 2) = FormatNumber(TotalPack, 0)
   Vsp.TableCell(tcText, 1, 3) = FormatNumber(TotalWeight, 0)
   Vsp.TableBorder = tbNone
   Vsp.EndTable
   

  Vsp.CurrentY = m_ReportConfig.HEAD_OFFSET + 2690
   Vsp.FontBold = False
    Vsp.FontSize = 13
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcRowHeight, 1) = 1.5 * 567
   Vsp.TableCell(tcColWidth, 1, 1) = "23.9cm"
   Vsp.TableCell(tcColWidth, 1, 2) = "2.30cm"
   Vsp.TableCell(tcAlign, 1, 2) = taCenterMiddle
   Vsp.TableCell(tcText, 1, 2) = EMP_CHECK_PRODUCT_NAME & vbNewLine & EMP_CHECK_PRODUCT_LNAME
   Vsp.TableBorder = tbNone
   Vsp.EndTable
   
   Vsp.CurrentY = m_ReportConfig.HEAD_OFFSET + 8850
   Vsp.FontSize = 13
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 3
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcRowHeight, 1) = 0.5 * 567
   Vsp.TableCell(tcColWidth, 1, 1) = "7.08cm"
   Vsp.TableCell(tcColWidth, 1, 2) = "4.85cm"
   Vsp.TableCell(tcColWidth, 1, 3) = "1cm"
   If CONDITION1 = 1 Then
   Vsp.TableCell(tcAlign, 1, 2) = taLeftBottom
   Vsp.TableCell(tcText, 1, 2) = "X"
   ElseIf CONDITION1 = 2 Then
   Vsp.TableCell(tcAlign, 1, 3) = taLeftBottom
   Vsp.TableCell(tcText, 1, 3) = "X"
   End If
   Vsp.TableBorder = tbNone
   Vsp.EndTable
   
   Vsp.CurrentY = Vsp.CurrentY - 60
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 3
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcRowHeight, 1) = 0.5 * 567
   Vsp.TableCell(tcColWidth, 1, 1) = "7.08cm"
   Vsp.TableCell(tcColWidth, 1, 2) = "4.85cm"
   Vsp.TableCell(tcColWidth, 1, 3) = "1cm"
   If CONDITION2 = 1 Then
   Vsp.TableCell(tcAlign, 1, 2) = taLeftBottom
   Vsp.TableCell(tcText, 1, 2) = "X"
   ElseIf CONDITION2 = 2 Then
   Vsp.TableCell(tcAlign, 1, 3) = taLeftBottom
   Vsp.TableCell(tcText, 1, 3) = "X"
   End If
   Vsp.TableBorder = tbNone
   Vsp.EndTable
   
   Vsp.CurrentY = Vsp.CurrentY - 70
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 3
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcRowHeight, 1) = 0.5 * 567
   Vsp.TableCell(tcColWidth, 1, 1) = "7.08cm"
   Vsp.TableCell(tcColWidth, 1, 2) = "4.85cm"
   Vsp.TableCell(tcColWidth, 1, 3) = "1cm"
   If CONDITION3 = 1 Then
   Vsp.TableCell(tcAlign, 1, 2) = taLeftBottom
   Vsp.TableCell(tcText, 1, 2) = "X"
   ElseIf CONDITION3 = 2 Then
   Vsp.TableCell(tcAlign, 1, 3) = taLeftBottom
   Vsp.TableCell(tcText, 1, 3) = "X"
   End If
   Vsp.TableBorder = tbNone
   Vsp.EndTable
   
   Vsp.CurrentY = Vsp.CurrentY + 100
   Vsp.FontSize = 14
   Vsp.FontBold = False
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcRowHeight, 1) = 0.7 * 567
   Vsp.TableCell(tcColWidth, 1, 1) = "2.5cm"
   Vsp.TableCell(tcColWidth, 1, 2) = "7.99cm"
   Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
   Vsp.TableCell(tcText, 1, 2) = EMP_CHECK_CAR_NAME & " " & EMP_CHECK_CAR_LNAME
   Vsp.TableBorder = tbNone
   Vsp.EndTable
   
   TempCY = Vsp.CurrentY
      
   Vsp.CurrentY = TempCY - 1420
   Vsp.FontSize = 20
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 3
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcRowHeight, 1) = 0.7 * 567
   Vsp.TableCell(tcColWidth, 1, 1) = "20cm"
   Vsp.TableCell(tcColWidth, 1, 2) = "3cm"
   Vsp.TableCell(tcColWidth, 1, 3) = "3cm"
   Vsp.TableCell(tcAlign, 1, 2) = taRightMiddle
   Vsp.TableCell(tcAlign, 1, 3) = taRightMiddle
   Vsp.TableCell(tcText, 1, 2) = IIf(ENTRY_WEIGHT = 0 Or ENTRY_WEIGHT = 9999999, "", FormatNumber(ENTRY_WEIGHT, 0))
   Vsp.TableCell(tcText, 1, 3) = Format(ENTRY_WEIGHT_TIME, "HH:mm")
   Vsp.TableBorder = tbNone
   Vsp.EndTable
   
   Vsp.CurrentY = Vsp.CurrentY + 30
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 3
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcRowHeight, 1) = 0.7 * 567
   Vsp.TableCell(tcColWidth, 1, 1) = "20cm"
   Vsp.TableCell(tcColWidth, 1, 2) = "3cm"
   Vsp.TableCell(tcColWidth, 1, 3) = "3cm"
   Vsp.TableCell(tcAlign, 1, 2) = taRightMiddle
   Vsp.TableCell(tcAlign, 1, 3) = taRightMiddle
   Vsp.TableCell(tcText, 1, 2) = IIf(EXIT_WEIGHT = 0 Or EXIT_WEIGHT = 9999999, "", FormatNumber(EXIT_WEIGHT, 0))
   Vsp.TableCell(tcText, 1, 3) = Format(EXIT_WEIGHT_TIME, "HH:mm")
   Vsp.TableBorder = tbNone
   Vsp.EndTable
   
   Vsp.CurrentY = Vsp.CurrentY + 20
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcRowHeight, 1) = 0.7 * 567
   Vsp.TableCell(tcColWidth, 1, 1) = "20cm"
   Vsp.TableCell(tcColWidth, 1, 2) = "3cm"
   Vsp.TableCell(tcAlign, 1, 2) = taRightMiddle
   Vsp.TableCell(tcText, 1, 2) = IIf(TOTAL_WEIGHT = 0 Or TOTAL_WEIGHT = 9999999, "", FormatNumber(TOTAL_WEIGHT, 0))
   Vsp.TableBorder = tbNone
   Vsp.EndTable
    Vsp.FontBold = False
    
'    If Len(DOCUMENT_DESC) > 0 Then
'      Vsp.FontSize = 16
'      Vsp.FontBold = True
'      Vsp.CurrentY = Vsp.CurrentY + 1200
'      Vsp.StartTable
'      Vsp.TableCell(tcCols) = 2
'      Vsp.TableCell(tcRows) = 1
'      Vsp.TableCell(tcRowHeight, 1) = 800 '1 * 567
'      Vsp.TableCell(tcColWidth, 1, 1) = "2.4cm"
'      Vsp.TableCell(tcColWidth, 1, 2) = "22.9cm"
'      Vsp.TableCell(tcAlign, 1, 1) = taRightTop
'      Vsp.TableCell(tcAlign, 1, 2) = taLeftTop
'      Vsp.TableCell(tcText, 1, 1) = "* หมายเหตุ : "
'      Vsp.TableCell(tcText, 1, 2) = DOCUMENT_DESC
'      Vsp.TableBorder = tbNone
'      Vsp.EndTable
'      Vsp.FontBold = False
'   End If
      
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
Dim TempLeft As Double

   If Not mblnNewPage Then Exit Sub
   isNewPage = True
   talnT = Vsp.TextAlign
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   blnUnder = Vsp.FontUnderline
   blnItalic = Vsp.FontItalic
   iSize = Vsp.FontSize
   sName = Vsp.FontName

'  Vsp.DrawPicture mcolParam("BACK_GROUND"), "0.5cm", "0.5cm", "27cm", "21cm", "LeftTop"
If mcolParam("FULL_SHOW") Then
' Vsp.DrawPicture mcolParam("BACK_GROUND"), "0.5cm", "0.3cm", "27cm", "20.5cm", "LeftTop"
  Vsp.DrawPicture mcolParam("BACK_GROUND"), "-.1cm", "0.85cm", "27.99cm", "20.6cm"
End If

  
  Vsp.MarginLeft = (0.4 * 567) + m_ReportConfig.DUMMY_OFFSET + 270
  
  Vsp.MarginTop = 0
  
   Vsp.StartTable
   Vsp.TableBorder = tbNone
   Vsp.TableCell(tcCols) = 1
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcRowHeight) = (1 * 567) + m_ReportConfig.HEAD_OFFSET + 520 ' '' (0.1 * 567) + m_ReportConfig.HEAD_OFFSET '
   Vsp.TableCell(tcColWidth, , 1) = "10cm"
   Vsp.EndTable
  
   Vsp.FontSize = 8
'
   dY(0) = Vsp.CurrentY
   Vsp.CurrentY = Vsp.CurrentY - 200
   TempLeft = Vsp.MarginLeft

   Vsp.MarginBottom = 0
   Vsp.TextAlign = taLeftBottom
   strFormat = VSP_CalTable(TITLE_FORMAT, mdblWid, alngX)
   Vsp.TableBorder = tbNone

   Call Vsp.AddTable(strFormat, "", "|หน้าที่|: " & Vsp.CurrentPage)
'   Call Vsp.AddTable(strFormat, "", "|วันที่พิมพ์|: " & DateToStringExt(mdteDate))
'   Call Vsp.AddTable(strFormat, "", "|เวลา|: " & Right$(DateToStringExtEx(mdteDate), 8))
'   Call Vsp.AddTable(strFormat, "", "|ผู้ใช้|: " & glbUser.USER_NAME)
   Vsp.CurrentY = dY(0)
'
'   Vsp.MarginLeft = TempLeft
'
'   Vsp.FontSize = TITLE_SIZE
'   Vsp.FontBold = True
'   Vsp.TextAlign = taCenterBottom
'   Vsp.Paragraph = MODULE_DESC
'   Vsp.Paragraph = ""
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

