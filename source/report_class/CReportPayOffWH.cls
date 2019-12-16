VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportPayOffWH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportIncomeWH"
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
Private Const DETAIL_SIZE = 10

Private Const TITLE_FORMAT = "<32|>3|<5;"

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
Private isNewPage As Boolean
Private Total1(100) As Double
Private TempCusName As Collection
Private m_CusLotItemWh As CLotItemWH

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_TempCol = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers0_1 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Details1 = New CFieldList
   Set m_PartTxtypes = New Collection
   Set TempCusName = New Collection
   
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
   Set TempCusName = Nothing
   
'   If Rs.State = adStateOpen Then
'      Rs.Close
'   End If
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
   
   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbNone
   Vsp.TextAlign = taLeftBottom
   
'   Vsp.MarginLeft = "3cm"
'   Vsp.CurrentY = 2400
'   Vsp.StartTable
'   Vsp.TableCell(tcCols) = 1
'   Vsp.TableCell(tcRows) = 1
'   Vsp.TableCell(tcRowHeight, 1) = 1 * 567
'   Vsp.TableCell(tcColWidth, 1, 1) = "3.5cm"
'   Vsp.TableCell(tcAlign, 1, 1) = taLeftBottom
'   Vsp.TableCell(tcText, 1, 1) = mcolParam("HEAD_PACK_NAME")
'   Vsp.TableBorder = tbNone
'   Vsp.EndTable
'
      Vsp.MarginLeft = "24cm"
      Vsp.CurrentY = 1000
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 1
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight, 1) = 1 * 567
      Vsp.TableCell(tcColWidth, 1, 1) = "3.5cm"
      Vsp.TableCell(tcAlign, 1, 1) = taLeftBottom
      Vsp.TableCell(tcText, 1, 1) = DateToStringExtEx2(mcolParam("PRINT_DATE"))
      Vsp.TableBorder = tbNone
      Vsp.EndTable
            
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
'   Vsp.DrawPicture mcolParam("BACK_GROUND"), "0.0cm", "0.4cm", "29.7cm", "20.5cm", 10
   
'   FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
'   ToDate = "ถึง " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")
'
'   MODULE_DESC = vbCrLf & glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & vbCrLf & _
'                                       mcolParam("REPORT_NAME") & vbCrLf & _
'                                        FromDate & " " & ToDate & " " & mcolParam("LOCATION_NAME")
   MODULE_DESC = ""
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
Dim num As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Pi As CPartItem
Dim iCount As Long
Dim TempStr As String
Dim Amt As Double
Dim m_InventoryWHDoc As CInventoryWHDoc
Dim LotItemWh As CLotItemWH
Dim C_IWD As Collection
Dim LTD As CLotDoc

Dim LD As CLotDoc
Dim PD As CPalletDoc
Dim Count1 As Double
Dim TempY As Double
Dim TempX As Double
Dim CountLT As Long
Dim CountPD As Long
Dim Arr(10) As String
Arr(0) = "1.5cm"
Arr(1) = "1.4cm"
Arr(2) = "1.35cm"
Arr(3) = "1.4cm"
Arr(4) = "1.3cm"
Arr(5) = "1.4cm"
Arr(6) = "1.4cm"
Arr(7) = "1.3cm"
Arr(8) = "1.4cm"
Arr(9) = "1.3cm"
Dim ArrH(12) As Long
ArrH(0) = 2400
ArrH(1) = 2850
ArrH(2) = 3270
ArrH(3) = 3690
ArrH(4) = 4120
ArrH(5) = 4550
ArrH(6) = 4970
ArrH(7) = 5400
ArrH(8) = 5820
ArrH(9) = 6230
ArrH(10) = 6670
ArrH(11) = 7090
Dim TempCount As Long
Dim CountRow As Long
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim SearchCusName As CLotItemWH

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
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
   Vsp.FontSize = 15
   I = 0
   num = 0
   Vsp.TextAlign = taLeftBottom
'   Vsp.MarginLeft = 472 'm_ReportConfig.DUMMY_OFFSET
   
   If mcolParam("NOT_SHOW") = 0 Then
    Set C_IWD = New Collection
    Set m_InventoryWHDoc = New CInventoryWHDoc
    m_InventoryWHDoc.INVENTORY_WH_DOC_ID = -1
    m_InventoryWHDoc.DOCUMENT_DATE = mcolParam("PRINT_DATE")
    If Val(mcolParam("DOCUMENT_TYPE")) > 0 Then
      m_InventoryWHDoc.DOCUMENT_TYPE = mcolParam("DOCUMENT_TYPE")
    Else
      m_InventoryWHDoc.DOCUMENT_TYPE_SET = "(2000,2001)"
    End If

    m_InventoryWHDoc.LOAD_FLAG = Check2FlagConvert3(mcolParam("LOAD_FLAG"))
    m_InventoryWHDoc.PART_NO = mcolParam("PART_NO")
    m_InventoryWHDoc.LOT_NO = mcolParam("LOT_NO")
    m_InventoryWHDoc.OrderBy = mcolParam("ORDER_BY")
    m_InventoryWHDoc.OrderType = mcolParam("ORDER_TYPE")
    m_InventoryWHDoc.QueryFlag = 1 'ถ้าเป็น 0 ยังไม่ต้องเอาลูก หลาน แต่ถ้าเป็น 1 เอา ลูกหลานด้วย
   
    If Not glbDaily.QueryInventoryWhDocForPG(m_InventoryWHDoc, C_IWD, Rs, iCount, IsOK, glbErrorLog) Then
       glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
       Call EnableForm(Me, True)
       Exit Function
    End If
       
    CountRow = 0
    For Each m_InventoryWHDoc In C_IWD
         For Each LotItemWh In m_InventoryWHDoc.C_LotItemsWH
         
             For Each LTD In LotItemWh.C_LotDoc
             
            Vsp.MarginLeft = 472
            Vsp.FontSize = 10
            I = I + 1
            J = 0
            num = 0
            CountLT = 0
            
            CountRow = CountRow + 1
            If CountRow > 12 Then
               CountRow = 1
              Call GenerateFooter
              
              Set TempCusName = Nothing
              Set TempCusName = New Collection
               For J = 1 To UBound(Total1)
                Total1(J) = 0
               Next J
   
              Vsp.NewPage
              Vsp.MarginLeft = 472
            End If
            
            Vsp.CurrentY = ArrH(CountRow - 1)
      
            Vsp.StartTable
            Vsp.TableCell(tcCols) = 15
            Vsp.TableCell(tcRows) = 1
            Vsp.TableCell(tcRowHeight, 1) = 1 * 567
            Vsp.TableCell(tcColWidth, 1, 1) = "1.2cm" 'ok
            Vsp.TableCell(tcColWidth, 1, 2) = "2.85cm" 'ok
            Vsp.TableCell(tcColWidth, 1, 3) = "3.2cm" 'ok
            Vsp.TableCell(tcColWidth, 1, 4) = "2.1cm" 'ok
            Vsp.TableCell(tcColWidth, 1, 5) = "2.27cm" 'ok
            Vsp.TableCell(tcColWidth, 1, 6) = "1.02cm" 'ok
            Vsp.TableCell(tcColWidth, 1, 7) = "1.7cm" 'ok
            Vsp.TableCell(tcColWidth, 1, 8) = "1.65cm" 'ok
            Vsp.TableCell(tcColWidth, 1, 9) = "2.25cm" 'ok
            Vsp.TableCell(tcColWidth, 1, 10) = "2.05cm" 'ok
            Vsp.TableCell(tcColWidth, 1, 11) = "1.75cm" 'ok
            Vsp.TableCell(tcColWidth, 1, 12) = "1.75cm" 'ok
            Vsp.TableCell(tcColWidth, 1, 13) = "2.02cm" 'ok
            Vsp.TableCell(tcColWidth, 1, 14) = "2.2cm" 'ok
            
            Vsp.TableCell(tcAlign, 1, 1) = taCenterMiddle
            Vsp.TableCell(tcAlign, 1, 2) = taCenterMiddle
            Vsp.TableCell(tcAlign, 1, 3) = taLeftMiddle
            Vsp.TableCell(tcAlign, 1, 4) = taLeftMiddle
            Vsp.TableCell(tcAlign, 1, 5) = taLeftMiddle
            Vsp.TableCell(tcAlign, 1, 6) = taRightMiddle
            Vsp.TableCell(tcAlign, 1, 7) = taRightMiddle
            Vsp.TableCell(tcAlign, 1, 8) = taRightMiddle
            Vsp.TableCell(tcAlign, 1, 9) = taRightMiddle
            Vsp.TableCell(tcAlign, 1, 10) = taCenterMiddle
            Vsp.TableCell(tcAlign, 1, 11) = taCenterMiddle
            Vsp.TableCell(tcAlign, 1, 12) = taCenterMiddle
            Vsp.TableCell(tcAlign, 1, 13) = taCenterMiddle
            Vsp.TableCell(tcAlign, 1, 14) = taLeftMiddle
      
           If PrevKey1 <> m_InventoryWHDoc.DOCUMENT_NO Or CountRow = 1 Then
               PrevKey1 = m_InventoryWHDoc.DOCUMENT_NO
               Vsp.TableCell(tcText, 1, 1) = I
               Vsp.TableCell(tcText, 1, 2) = m_InventoryWHDoc.DOCUMENT_NO
               Vsp.TableCell(tcText, 1, 3) = Left(m_InventoryWHDoc.CUSTOMER_NAME, IIf(Val(mcolParam("LIMIT_CUS_NAME")) = 0, 15, Val(mcolParam("LIMIT_CUS_NAME"))))
               Vsp.TableCell(tcText, 1, 4) = m_InventoryWHDoc.TRUCK_NO
            
            Else
               Vsp.TableCell(tcText, 1, 1) = I
               Vsp.TableCell(tcText, 1, 2) = ""
               Vsp.TableCell(tcText, 1, 3) = ""
               Vsp.TableCell(tcText, 1, 4) = ""
         End If

         If PrevKey2 <> m_InventoryWHDoc.CUSTOMER_NAME Or CountRow = 1 Then
               PrevKey2 = m_InventoryWHDoc.CUSTOMER_NAME
               Set m_CusLotItemWh = New CLotItemWH
                m_CusLotItemWh.CUSTOMER_NAME = PrevKey2
              
               Set SearchCusName = GetObject("CLotItemWH", TempCusName, Trim(PrevKey2), False)
               If SearchCusName Is Nothing Then   '
                  Call TempCusName.add(m_CusLotItemWh, Trim(PrevKey2))
               End If
         End If
         
      Vsp.TableCell(tcText, 1, 5) = Left(LotItemWh.PART_NO, IIf(Val(mcolParam("LIMIT_PART_NAME")) = 0, 25, Val(mcolParam("LIMIT_PART_NAME"))))
         
      Dim TempValue As Double
      TempValue = LTD.CAPACITY_AMOUNT
        If mcolParam("DOCUMENT_TYPE") = 2000 Or mcolParam("DOCUMENT_TYPE") = 2002 Then
            Vsp.TableCell(tcText, 1, 6) = IIf(Val(TempValue) > 0, FormatNumber(TempValue, 0), "-")
            Vsp.TableCell(tcAlign, 1, 7) = taCenterMiddle
            Vsp.TableCell(tcAlign, 1, 8) = taCenterMiddle
            Vsp.TableCell(tcText, 1, 7) = "-"
            Vsp.TableCell(tcText, 1, 8) = "-"
        ElseIf mcolParam("DOCUMENT_TYPE") = 2001 Or mcolParam("DOCUMENT_TYPE") = 2004 Then
            Vsp.TableCell(tcAlign, 1, 6) = taCenterMiddle
            Vsp.TableCell(tcAlign, 1, 7) = taCenterMiddle
            Vsp.TableCell(tcText, 1, 6) = "-"
            Vsp.TableCell(tcText, 1, 7) = "-"
            Vsp.TableCell(tcText, 1, 8) = IIf(Val(TempValue) > 0, FormatNumber(TempValue, 0), "-")
         Else
            Vsp.TableCell(tcText, 1, 6) = "-"
            Vsp.TableCell(tcText, 1, 7) = "-"
            Vsp.TableCell(tcText, 1, 8) = "-"
         End If
      
      Vsp.TableCell(tcText, 1, 9) = IIf(Val(LTD.CAPACITY_AMOUNT * LotItemWh.WEIGHT_PER_PACK) > 0, FormatNumber(LTD.CAPACITY_AMOUNT * LotItemWh.WEIGHT_PER_PACK, 0), "-") 'FormatNumber(LTD.CAPACITY_AMOUNT * LotItemWh.WEIGHT_PER_PACK, 0)
      
      Set SearchCusName = GetObject("CLotItemWH", TempCusName, Trim(m_InventoryWHDoc.CUSTOMER_NAME), False)
      If Not SearchCusName Is Nothing Then  '
         SearchCusName.PACK_AMOUNT = SearchCusName.PACK_AMOUNT + TempValue
         SearchCusName.TOTAL_WEIGHT = SearchCusName.TOTAL_WEIGHT + (LTD.CAPACITY_AMOUNT * LotItemWh.WEIGHT_PER_PACK)
      End If
      
         Vsp.TableCell(tcText, 1, 10) = IIf(LTD.LOCK_NAME = "", "-", LTD.LOCK_NAME) 'TempStr
         If mcolParam("DOCUMENT_TYPE") = 2000 Or mcolParam("DOCUMENT_TYPE") = 2002 Then
            Vsp.TableCell(tcText, 1, 11) = "X"
            Vsp.TableCell(tcText, 1, 12) = "-"
        ElseIf mcolParam("DOCUMENT_TYPE") = 2001 Or mcolParam("DOCUMENT_TYPE") = 2004 Then
            Vsp.TableCell(tcText, 1, 11) = "-"
            Vsp.TableCell(tcText, 1, 12) = "X"
         Else
            Vsp.TableCell(tcText, 1, 11) = "-"
            Vsp.TableCell(tcText, 1, 12) = "-"
         End If
         
               Vsp.TableCell(tcText, 1, 13) = ""
               Vsp.TableCell(tcText, 1, 14) = m_InventoryWHDoc.NOTE
               Vsp.TableBorder = tbNone
               Vsp.EndTable
               
               Vsp.MarginLeft = 14000
               Vsp.FontSize = 9
               Vsp.CurrentY = ArrH(CountRow - 1) + 60
               Vsp.StartTable
               Vsp.TableCell(tcCols) = 1
               Vsp.TableCell(tcRows) = 2
               Vsp.TableCell(tcRowHeight, 1) = 1 * 200
               Vsp.TableCell(tcRowHeight, 2) = 1 * 200
               Vsp.TableCell(tcColWidth, 1, 1) = "2.02cm" 'ok
               Vsp.TableCell(tcColWidth, 2, 1) = "2.02cm" 'ok
               Vsp.TableCell(tcAlign, 1, 1) = taCenterMiddle
               Vsp.TableCell(tcAlign, 2, 1) = taCenterMiddle
               
               If LTD.DOCUMENT_TYPE = 15 Or LTD.DOCUMENT_TYPE = 16 Then
                  Vsp.TableCell(tcText, 1, 1) = DateToStringExtEx2(LTD.BL_START_DATE)
               Else
                  Vsp.TableCell(tcText, 1, 1) = DateToStringExtEx2(LTD.START_DATE)
               End If
               Vsp.TableCell(tcText, 2, 1) = LTD.LOT_NO
               Vsp.TableBorder = tbNone
               Vsp.EndTable
               Vsp.CurrentY = ArrH(CountRow - 1) - 60
         
              Next LTD
            Set m_CusLotItemWh = Nothing
         Next LotItemWh
   Next m_InventoryWHDoc
   Call GenerateFooter
   End If
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   Set m_InventoryWHDoc = Nothing
   genDoc = True
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
   Set Rs = Nothing
End Function
'Private Function genDoc(isPreview As Boolean) As Boolean
'On Error GoTo ErrHandler
'Dim RName As String
'Dim I As Long
'Dim j As Long
'Dim num As Long
'Dim strFormat As String
'Dim alngX() As Long
'Dim IsOK As Boolean
'Dim HeadCf As CReportField
'Dim BodyCf As CReportField
'Dim TempStr1 As String
'Dim TempStr2 As String
'Dim Pi As CPartItem
'Dim iCount As Long
'Dim TempStr As String
'Dim Amt As Double
'Dim m_InventoryWHDoc As CInventoryWHDoc
'Dim LotItemWh As CLotItemWH
'Dim C_IWD As Collection
'Dim LTD As CLotDoc
'
'Dim LD As CLotDoc
'Dim PD As CPalletDoc
'Dim Count1 As Double
'Dim TempY As Double
'Dim TempX As Double
'Dim CountLT As Long
'Dim CountPD As Long
'Dim Arr(10) As String
'Arr(0) = "1.5cm"
'Arr(1) = "1.4cm"
'Arr(2) = "1.35cm"
'Arr(3) = "1.4cm"
'Arr(4) = "1.3cm"
'Arr(5) = "1.4cm"
'Arr(6) = "1.4cm"
'Arr(7) = "1.3cm"
'Arr(8) = "1.4cm"
'Arr(9) = "1.3cm"
'Dim ArrH(12) As Long
'ArrH(0) = 2400
'ArrH(1) = 2850
'ArrH(2) = 3270
'ArrH(3) = 3690
'ArrH(4) = 4120
'ArrH(5) = 4550
'ArrH(6) = 4970
'ArrH(7) = 5400
'ArrH(8) = 5820
'ArrH(9) = 6230
'ArrH(10) = 6670
'ArrH(11) = 7090
'Dim TempCount As Long
'Dim CountRow As Long
'Dim PrevKey1 As String
'Dim PrevKey2 As String
'Dim SearchCusName As CLotItemWH
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
'
'   Call StartExportFile(Vsp)
'   Vsp.StartDoc
'   Vsp.FontSize = 15
'   I = 0
'   num = 0
'   Vsp.TextAlign = taLeftBottom
'   Vsp.MarginLeft = 472 'm_ReportConfig.DUMMY_OFFSET
'
'
'   Vsp.FontSize = 11
'
'   If mcolParam("NOT_SHOW") = 0 Then
'    Set C_IWD = New Collection
'    Set m_InventoryWHDoc = New CInventoryWHDoc
'    m_InventoryWHDoc.INVENTORY_WH_DOC_ID = -1
'    m_InventoryWHDoc.DOCUMENT_DATE = mcolParam("PRINT_DATE")
'    m_InventoryWHDoc.DOCUMENT_TYPE = 2000 'mcolParam("DOCUMENT_TYPE")
'    m_InventoryWHDoc.OrderBy = mcolParam("ORDER_BY")
'    m_InventoryWHDoc.OrderType = mcolParam("ORDER_TYPE")
'    m_InventoryWHDoc.QueryFlag = 1 'ถ้าเป็น 0 ยังไม่ต้องเอาลูก หลาน แต่ถ้าเป็น 1 เอา ลูกหลานด้วย
'
'    If Not glbDaily.QueryInventoryWhDocForPG(m_InventoryWHDoc, C_IWD, Rs, iCount, IsOK, glbErrorLog) Then
'       glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'       Call EnableForm(Me, True)
'       Exit Function
'    End If
'
'    CountRow = 0
'    For Each m_InventoryWHDoc In C_IWD
'         For Each LotItemWh In m_InventoryWHDoc.C_LotItemsWH
'            I = I + 1
'            j = 0
'            num = 0
'            CountLT = 0
'
'            CountRow = CountRow + 1
'            If CountRow > 12 Then
'               CountRow = 1
'              Call GenerateFooter
'
'              Set TempCusName = Nothing
'              Set TempCusName = New Collection
'               For j = 1 To UBound(Total1)
'                Total1(j) = 0
'               Next j
'
'              Vsp.NewPage
'              Vsp.MarginLeft = 472
'            End If
'
'            Vsp.CurrentY = ArrH(CountRow - 1)
'
'            Vsp.StartTable
'            Vsp.TableCell(tcCols) = 15
'            Vsp.TableCell(tcRows) = 1
'            Vsp.TableCell(tcRowHeight, 1) = 1 * 567
'            Vsp.TableCell(tcColWidth, 1, 1) = "1.2cm" 'ok
'            Vsp.TableCell(tcColWidth, 1, 2) = "2.85cm" 'ok
'            Vsp.TableCell(tcColWidth, 1, 3) = "3.2cm" 'ok
'            Vsp.TableCell(tcColWidth, 1, 4) = "2.1cm" 'ok
'            Vsp.TableCell(tcColWidth, 1, 5) = "2.57cm" 'ok
'            Vsp.TableCell(tcColWidth, 1, 6) = "0.72cm" 'ok
'            Vsp.TableCell(tcColWidth, 1, 7) = "1.7cm" 'ok
'            Vsp.TableCell(tcColWidth, 1, 8) = "1.65cm" 'ok
'            Vsp.TableCell(tcColWidth, 1, 9) = "2.25cm" 'ok
'            Vsp.TableCell(tcColWidth, 1, 10) = "2.05cm" 'ok
'            Vsp.TableCell(tcColWidth, 1, 11) = "1.75cm" 'ok
'            Vsp.TableCell(tcColWidth, 1, 12) = "1.75cm" 'ok
'            Vsp.TableCell(tcColWidth, 1, 13) = "2.02cm" 'ok
'            Vsp.TableCell(tcColWidth, 1, 14) = "2.2cm" 'ok
'            Vsp.TableCell(tcColWidth, 1, 15) = "3.5cm" 'ok
'
'            Vsp.TableCell(tcAlign, 1, 1) = taCenterMiddle
'            Vsp.TableCell(tcAlign, 1, 2) = taCenterMiddle
'            Vsp.TableCell(tcAlign, 1, 3) = taLeftMiddle
'            Vsp.TableCell(tcAlign, 1, 4) = taLeftMiddle
'            Vsp.TableCell(tcAlign, 1, 5) = taLeftMiddle
'            Vsp.TableCell(tcAlign, 1, 6) = taRightMiddle
'            Vsp.TableCell(tcAlign, 1, 7) = taRightMiddle
'            Vsp.TableCell(tcAlign, 1, 8) = taRightMiddle
'            Vsp.TableCell(tcAlign, 1, 9) = taRightMiddle
'            Vsp.TableCell(tcAlign, 1, 10) = taCenterMiddle
'            Vsp.TableCell(tcAlign, 1, 11) = taCenterMiddle
'            Vsp.TableCell(tcAlign, 1, 12) = taCenterMiddle
'            Vsp.TableCell(tcAlign, 1, 13) = taCenterMiddle
'            Vsp.TableCell(tcAlign, 1, 14) = taLeftMiddle
'
'           If PrevKey1 <> m_InventoryWHDoc.DOCUMENT_NO Or CountRow = 1 Then
'               PrevKey1 = m_InventoryWHDoc.DOCUMENT_NO
'               Vsp.TableCell(tcText, 1, 1) = I
'               Vsp.TableCell(tcText, 1, 2) = m_InventoryWHDoc.DOCUMENT_NO
'               Vsp.TableCell(tcText, 1, 3) = Left(m_InventoryWHDoc.CUSTOMER_NAME, Val(mcolParam("LIMIT_CUS_NAME")))
'               Vsp.TableCell(tcText, 1, 4) = m_InventoryWHDoc.TRUCK_NO
'
'            Else
'               Vsp.TableCell(tcText, 1, 1) = I
'               Vsp.TableCell(tcText, 1, 2) = ""
'               Vsp.TableCell(tcText, 1, 3) = ""
'               Vsp.TableCell(tcText, 1, 4) = ""
'         End If
'
'         If PrevKey2 <> m_InventoryWHDoc.CUSTOMER_NAME Or CountRow = 1 Then
'               PrevKey2 = m_InventoryWHDoc.CUSTOMER_NAME
'               Set m_CusLotItemWh = New CLotItemWH
'                m_CusLotItemWh.CUSTOMER_NAME = PrevKey2
'                Call TempCusName.add(m_CusLotItemWh, Trim(PrevKey2))
'         End If
'
'      Vsp.TableCell(tcText, 1, 5) = LotItemWh.PART_NO
'
'      Dim TempValue As Double
'        If LotItemWh.PART_TYPE = 10 Then
'            TempValue = LotItemWh.PACK_AMOUNT
'            Vsp.TableCell(tcText, 1, 6) = IIf(Val(TempValue) > 0, FormatNumber(TempValue, 0), "-")
'            Vsp.TableCell(tcAlign, 1, 7) = taCenterMiddle
'            Vsp.TableCell(tcAlign, 1, 8) = taCenterMiddle
'            Vsp.TableCell(tcText, 1, 7) = "-"
'            Vsp.TableCell(tcText, 1, 8) = "-"
'        ElseIf LotItemWh.PART_TYPE = 22 Then
'             TempValue = LotItemWh.TX_AMOUNT
'            Vsp.TableCell(tcAlign, 1, 6) = taCenterMiddle
'            Vsp.TableCell(tcAlign, 1, 7) = taCenterMiddle
'            Vsp.TableCell(tcText, 1, 6) = "-"
'            Vsp.TableCell(tcText, 1, 7) = "-"
'            Vsp.TableCell(tcText, 1, 8) = IIf(Val(TempValue) > 0, FormatNumber(TempValue, 0), "-")
'         Else
'            Vsp.TableCell(tcText, 1, 6) = "-"
'            Vsp.TableCell(tcText, 1, 7) = "-"
'            Vsp.TableCell(tcText, 1, 8) = "-"
'         End If
'
'      Vsp.TableCell(tcText, 1, 9) = FormatNumber(LotItemWh.PACK_AMOUNT * LotItemWh.WEIGHT_PER_PACK, 0)
'
'      Set SearchCusName = GetObject("CLotItemWH", TempCusName, Trim(m_InventoryWHDoc.CUSTOMER_NAME), False)
'      If Not SearchCusName Is Nothing Then  '
'         SearchCusName.PACK_AMOUNT = SearchCusName.PACK_AMOUNT + TempValue
'         SearchCusName.TOTAL_WEIGHT = SearchCusName.TOTAL_WEIGHT + (LotItemWh.PACK_AMOUNT * LotItemWh.WEIGHT_PER_PACK)
'      End If
'
'            TempStr = ""
'            TempStr1 = ""
'            j = 0
'             For Each LTD In LotItemWh.C_LotDoc
'                  j = j + 1
'                  If j = 1 Then
'                      TempStr = LTD.LOCK_NAME
'                      TempStr1 = LTD.LOCK_NAME
'                  Else
'                     If TempStr1 <> LTD.LOCK_NAME Then
'                        TempStr = TempStr & "," & LTD.LOCK_NAME
'                        TempStr1 = LTD.LOCK_NAME
'                     End If
'                  End If
'               Next LTD
'''             For Each LTD In LotItemWh.C_LotDoc
'''                  j = j + 1
'''                  If j = 1 Then
'''                      TempStr = LTD.LOCK_NAME
'''                      TempStr1 = LTD.LOCK_NAME
'''                  Else
'''                    ' If TempStr1 <> LTD.LOCK_NAME Then
'''                        TempStr = TempStr & "," & LTD.LOCK_NAME
'''                        TempStr1 = LTD.LOCK_NAME
'''                     'End If
'''                  End If
'''               Next LTD
'            Vsp.TableCell(tcText, 1, 10) = TempStr
'         If LotItemWh.PART_TYPE = 10 Then
'            Vsp.TableCell(tcText, 1, 11) = "X"
'            Vsp.TableCell(tcText, 1, 12) = "-"
'        ElseIf LotItemWh.PART_TYPE = 22 Then
'            Vsp.TableCell(tcText, 1, 11) = "-"
'            Vsp.TableCell(tcText, 1, 12) = "X"
'         Else
'            Vsp.TableCell(tcText, 1, 11) = "-"
'            Vsp.TableCell(tcText, 1, 12) = "-"
'         End If
'
'         TempStr = ""
'         TempStr1 = ""
'         j = 0
'          For Each LTD In LotItemWh.C_LotDoc
'               j = j + 1
'               If j = 1 Then
'                   TempStr = LTD.START_DATE & "/" & LTD.LOT_NO
'                   TempStr1 = LTD.START_DATE
'               Else
'                  If TempStr1 <> LTD.START_DATE Then
'                     TempStr = TempStr & vbNewLine & LTD.START_DATE & "/" & LTD.LOT_NO
'                     TempStr1 = LTD.START_DATE
'                  End If
'               End If
'            Next LTD
'''''''          For Each LTD In LotItemWh.C_LotDoc
'''''''               j = j + 1
'''''''               If j = 1 Then
'''''''                   TempStr = LTD.START_DATE '& "/" & LTD.LOT_NO
'''''''                   TempStr1 = LTD.START_DATE
'''''''               Else
'''''''                 ' If TempStr1 <> LTD.START_DATE Then
''''''''                     tempStr = tempStr & vbNewLine & LTD.START_DATE '& "/" & LTD.LOT_NO
'''''''                     TempStr = TempStr & "," & LTD.START_DATE '& "/" & LTD.LOT_NO
'''''''                     TempStr1 = LTD.START_DATE
'''''''                 ' End If
'''''''               End If
'''''''            Next LTD
'
'            Vsp.TableCell(tcText, 1, 13) = TempStr
'            Vsp.TableCell(tcText, 1, 14) = m_InventoryWHDoc.NOTE
'
'            Vsp.TableBorder = tbNone
'            Vsp.EndTable
'
'            Set m_CusLotItemWh = Nothing
'         Next LotItemWh
'   Next m_InventoryWHDoc
'   Call GenerateFooter
'   End If
'   Vsp.EndDoc
'   Call CloseExportFile(Vsp)
'   Set m_InventoryWHDoc = Nothing
'   genDoc = True
'   Exit Function
'
'ErrHandler:
'   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
'   Set Rs = Nothing
'End Function

Private Sub GenerateFooter()
Dim CountRow As Long
Dim CountTotal As Long
Dim ArrH(5) As Long
Dim LWH As CLotItemWH
Dim Total(4) As Double
Dim J As Long
ArrH(0) = 9230
ArrH(1) = 9670
ArrH(2) = 10070
ArrH(3) = 10500
ArrH(4) = 10920
CountRow = 0
CountTotal = 0
Dim CurrentLeft As Double

For J = 1 To UBound(Total)
   Total(J) = 0
Next J

CurrentLeft = 1210
Vsp.FontSize = 10
For Each LWH In TempCusName
   CountRow = CountRow + 1
   CountTotal = CountTotal + 1
   If CountRow > 5 Then
      CountRow = 1
      CurrentLeft = 7280
      Vsp.MarginLeft = CurrentLeft
   Else
     Vsp.MarginLeft = CurrentLeft
   End If
   
   Vsp.CurrentY = ArrH(CountRow - 1)
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 3
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcRowHeight, 1) = 1 * 567
   
   Vsp.TableCell(tcColWidth, 1, 1) = "2.85cm" 'ok
   Vsp.TableCell(tcColWidth, 1, 2) = "3.2cm" 'ok
   Vsp.TableCell(tcColWidth, 1, 3) = "2.23cm" 'ok
   
   Vsp.TableCell(tcAlign, 1, 1) = taLeftMiddle
   Vsp.TableCell(tcAlign, 1, 2) = taCenterMiddle
   Vsp.TableCell(tcAlign, 1, 3) = taCenterMiddle
   
   Vsp.TableCell(tcText, 1, 1) = Left(LWH.CUSTOMER_NAME, IIf(Val(mcolParam("LIMIT_CUS_NAME")) = 0, 15, Val(mcolParam("LIMIT_CUS_NAME"))))
   Vsp.TableCell(tcText, 1, 2) = FormatNumber(LWH.PACK_AMOUNT, 0)
   Vsp.TableCell(tcText, 1, 3) = FormatNumber(LWH.TOTAL_WEIGHT, 0)
   
  Total(2) = Total(2) + LWH.PACK_AMOUNT
  Total(3) = Total(3) + LWH.TOTAL_WEIGHT
   
   Vsp.TableBorder = tbNone
   Vsp.EndTable
Next LWH

'ยอดรวม
   Vsp.MarginLeft = 7280
   Vsp.CurrentY = ArrH(4)
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 3
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcRowHeight, 1) = 1 * 567
   
   Vsp.TableCell(tcColWidth, 1, 1) = "2.85cm" 'ok
   Vsp.TableCell(tcColWidth, 1, 2) = "3.2cm" 'ok
   Vsp.TableCell(tcColWidth, 1, 3) = "2.23cm" 'ok
   
   Vsp.TableCell(tcAlign, 1, 1) = taLeftMiddle
   Vsp.TableCell(tcAlign, 1, 2) = taCenterMiddle
   Vsp.TableCell(tcAlign, 1, 3) = taCenterMiddle
   
   Vsp.TableCell(tcText, 1, 1) = ""
   Vsp.TableCell(tcText, 1, 2) = FormatNumber(Total(2), 0)
   Vsp.TableCell(tcText, 1, 3) = FormatNumber(Total(3), 0)
      
   Vsp.TableBorder = tbNone
   Vsp.EndTable

Vsp.MarginLeft = 472
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
   isNewPage = True
   talnT = Vsp.TextAlign
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   blnUnder = Vsp.FontUnderline
   blnItalic = Vsp.FontItalic
   iSize = Vsp.FontSize
   sName = Vsp.FontName
  
  Vsp.DrawPicture mcolParam("BACK_GROUND"), "0.0cm", "0.3cm", "29.7cm", "20.5cm", 10
   Vsp.FontSize = 8
   Vsp.CurrentY = 400
   Vsp.MarginLeft = 2535
   Vsp.MarginBottom = 0
   dY(0) = Vsp.CurrentY
   Vsp.TextAlign = taLeftBottom
   strFormat = VSP_CalTable(TITLE_FORMAT, mdblWid, alngX)
   Vsp.TableBorder = tbNone
   Call Vsp.AddTable(strFormat, "", "|หน้าที่|: " & Vsp.CurrentPage)
'   Call Vsp.AddTable(strFormat, "", "|วันที่พิมพ์|: " & DateToStringExt(mdteDate))
'   Call Vsp.AddTable(strFormat, "", "|เวลา|: " & Right$(DateToStringExtEx(mdteDate), 8))
'   Call Vsp.AddTable(strFormat, "", "|ผู้ใช้|: " & glbUser.USER_NAME)
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


