VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportPlanning002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportPlanning002"
Private MODULE_DESC  As String

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
Public ErrorObj As clsErrorLog

Private Const TITLE_SIZE = 20
Private Const HEADER_SIZE = 14
Private Const DETAIL_SIZE = 14

Private Const TITLE_FORMAT = "<32|<3|<5;"
Private Const HEADER_FORMAT = "^1|^2|^5|^2|^2|^2;"
Private Const DETAIL_FORMAT = "^1|^2|<5|>2|>2|>2;"
Private Const FOOTER_FORMAT_2 = "^1|>2|^5|^2|>2|^2;"
Private Const FOOTER_FORMAT_1 = "^1|>2|^5|^2|>2|^2;"

Private m_CPlanningItem As CPlanning
Private checkGetOrUse As String
Private Sp As CSystemParam
Private PlanningArea As Long

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set ErrorObj = New clsErrorLog
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set ErrorObj = Nothing
   Set Vsp = Nothing
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

Private Sub PrintBillFooter(alngX() As Long)
Dim strFormat As String
Dim strPrint As String

   Vsp.TableBorder = tbNone
   Vsp.FontBold = False
   
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
End Sub
Private Sub printHeaderDetail()
   Dim strFormat As String
   Dim strFormat2 As String
   Dim strPrint As String
   Dim tbdT As TableBorderSettings
   Dim blnBold As Boolean
   Dim iSize As Integer
   Dim alngX() As Long
   
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = DETAIL_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbNone
   strFormat = VSP_CalTable("<10|<10;", mdblWid, alngX)
   strFormat2 = VSP_CalTable("<20|;", mdblWid, alngX)
   
   If Len(m_CPlanningItem.PLANNING_DESC) > 0 Then
         strPrint = "รายละเอียด : " & m_CPlanningItem.PLANNING_DESC & "|;"
         Call Vsp.AddTable(strFormat, "", strPrint)
   End If
   If PlanningArea = 1 Or PlanningArea = 3 Then
          strPrint = "วันที่ประมาณการ : " & DateToStringExtEx2(m_CPlanningItem.PLANNING_DATE) & ";"
          Call Vsp.AddTable(strFormat2, "", strPrint)
   ElseIf PlanningArea = 2 Or PlanningArea = 4 Then
          strPrint = "จากวันที่ประมาณการ : " & DateToStringExtEx2(m_CPlanningItem.PLANNING_FROM) & " ถึงวันที่ประมาณการ : " & DateToStringExtEx2(m_CPlanningItem.PLANNING_TO) & ";"
          Call Vsp.AddTable(strFormat2, "", strPrint)
   End If
   
   Vsp.TableBorder = tbdT
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
End Sub

Private Sub printHeader()
   Dim strFormat As String
   Dim strPrint As String
   Dim tbdT As TableBorderSettings
   Dim blnBold As Boolean
   Dim iSize As Integer
   Dim alngX() As Long
   
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Vsp.TextAlign = taLeftBottom
   If checkGetOrUse = "PartUse" Then
         Vsp = "วัตถุดิบที่ใช้ผลิต"
   ElseIf checkGetOrUse = "ProductGet" Then
         Vsp = "สินค้าที่ผลิตได้"
   ElseIf checkGetOrUse = "PartSup" Then
         Vsp = "วัตถุดิบที่ใช้ผลิตรับจากซัพพลายเออร์"
   End If
   
   strFormat = VSP_CalTable(HEADER_FORMAT, mdblWid, alngX)
   strPrint = "ลำดับ|รหัส|ชื่อ|จำนวน|หน่วย|หมายเหตุ;"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Vsp.TableBorder = tbdT
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
End Sub

Private Function initDoc() As Boolean
Dim strDate As String
   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprLetter
   
   Vsp.ORIENTATION = orPortrait
   Vsp.MarginBottom = 1440
   Vsp.MarginFooter = 1440
   Vsp.MarginHeader = 1440
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 300
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
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

If PlanningArea = 1 Then
MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & _
                                       vbCrLf & MapText("รายงานประมาณการ การใช้วัตถุดิบและผลิตสินค้ารายวัน")
ElseIf PlanningArea = 2 Then
MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & _
                                       vbCrLf & MapText("รายงานประมาณการ การใช้วัตถุดิบและผลิตสินค้ารายสัปดาห์")
ElseIf PlanningArea = 3 Then
MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & _
                                       vbCrLf & MapText("รายงานประมาณการ รับเข้าวัตถุดิบรายวันจากซัพพลายเออร์")
ElseIf PlanningArea = 4 Then
MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & _
                                       vbCrLf & MapText("รายงานประมาณการ การใช้วัตถุดิบและผลิตสินค้ารายเดือน")
End If
   
   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))
   

   
   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
      mstrErrMsg = "Error GetDateTime Error."
      Exit Function
   End If
   mdteDate = InternalDateToDate(strDate)
   initDoc = True
End Function

Private Function GetSpace() As Long

End Function

Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim RName As String
Dim iCount As Long
Dim I As Long
Dim strFormat As String
Dim strPrint As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim m_PLN As CPlanningItem
Dim Sum1 As Double
Dim Sum2 As Double
Dim Sum3 As Double
Dim DiscountAmount As Double
Dim VatAmount As Double
Dim IncludeDiscountAmount As Double
Dim IncludeVatAmount As Double
Dim TempVatAmount As Double
Dim TempDeposit As Double
Dim TempDiscount As Double
Dim TempCol As Collection
Dim SpaceCount As Long
Dim Total(10) As Double
Dim J As Long

   RName = "genDoc"
   
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   Set Rs = New ADODB.Recordset
   Set m_CPlanningItem = New CPlanning
   m_CPlanningItem.PLANNING_ID = mcolParam("PLANNING_ID")
   m_CPlanningItem.QueryFlag = 1
   If Not glbPlanning.QueryPlanning(m_CPlanningItem, Rs, iCount, IsOK, glbErrorLog) Then
      mstrErrMsg = "Query Error."
      Exit Function
   End If
   If (Rs Is Nothing) Or (iCount < 1) Then
      mstrErrMsg = "ไม่พบข้อมูลที่ต้องการในฐานข้อมูล"
      Exit Function
   End If
   
   Call m_CPlanningItem.PopulateFromRS(1, Rs)
   PlanningArea = m_CPlanningItem.PLANNING_AREA
   If PlanningArea = 1 Or PlanningArea = 2 Or PlanningArea = 4 Then
         checkGetOrUse = "PartUse"
   ElseIf PlanningArea = 3 Then
         checkGetOrUse = "PartSup"
   End If
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
   Sum1 = 0
   Sum2 = 0
   Sum3 = 0
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
      strFormat = VSP_CalTable(DETAIL_FORMAT, mdblWid, alngX)
      Vsp.TableBorder = tbAll
      I = 0

      TempDiscount = 0
      TempVatAmount = 0

      For J = 1 To 10
         Total(J) = 0
      Next J
   
   If PlanningArea = 1 Or PlanningArea = 2 Or PlanningArea = 4 Then
      Set TempCol = m_CPlanningItem.CollPartUse
      For Each m_PLN In TempCol
         I = I + 1
         strPrint = I & "|"
         strPrint = strPrint & m_PLN.PART_NO & "|"
         strPrint = strPrint & m_PLN.PART_DESC & "|"
         strPrint = strPrint & FormatNumber(m_PLN.PLAN_AMOUNT) & "|"
         Total(1) = Total(1) + m_PLN.PLAN_AMOUNT
         strPrint = strPrint & m_PLN.UNIT_NAME & "|"
         strPrint = strPrint & m_PLN.NOTE & ";"

         Call Vsp.AddTable(strFormat, "", strPrint)
      Next m_PLN

      Vsp.FontBold = True
      strPrint = "รวม|||" & FormatNumber(Total(1)) & "||;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      checkGetOrUse = "ProductGet"
      Vsp.NewPage
      Set TempCol = Nothing
      I = 0
      Set TempCol = m_CPlanningItem.CollProductGet
      For Each m_PLN In TempCol
         I = I + 1
         strPrint = I & "|"
         strPrint = strPrint & m_PLN.PART_NO & "|"
         strPrint = strPrint & m_PLN.PART_DESC & "|"
         strPrint = strPrint & FormatNumber(m_PLN.PLAN_AMOUNT) & "|"
         Total(2) = Total(2) + m_PLN.PLAN_AMOUNT
         strPrint = strPrint & m_PLN.UNIT_NAME & "|"
         strPrint = strPrint & m_PLN.NOTE & ";"

         Call Vsp.AddTable(strFormat, "", strPrint)
      Next m_PLN

      Vsp.FontBold = True
      strPrint = "รวม|||" & FormatNumber(Total(2)) & "" & "||;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      Vsp.FontBold = False
      
      ElseIf PlanningArea = 3 Then
         Set TempCol = m_CPlanningItem.CollPartSup
         For Each m_PLN In TempCol
            I = I + 1
            strPrint = I & "|"
            strPrint = strPrint & m_PLN.PART_NO & "|"
            strPrint = strPrint & m_PLN.PART_DESC & "|"
            strPrint = strPrint & FormatNumber(m_PLN.PLAN_AMOUNT) & "|"
            Total(1) = Total(1) + m_PLN.PLAN_AMOUNT
            strPrint = strPrint & m_PLN.UNIT_NAME & "|"
            strPrint = strPrint & m_PLN.NOTE & ";"
   
            Call Vsp.AddTable(strFormat, "", strPrint)
         Next m_PLN
   
         Vsp.FontBold = True
         strPrint = "รวม|||" & FormatNumber(Total(1)) & "||;"
         Call Vsp.AddTable(strFormat, "", strPrint)
         End If
      'Call PrintBillFooter(alngX)
   Vsp.EndDoc
   Call CloseExportFile(Vsp)

   Rs.Close
   Set Rs = Nothing
   Set m_CPlanningItem = Nothing
   
   genDoc = True
   Exit Function

ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = MODULE_NAME
   ErrorObj.LocalErrorMsg = "Run time error."
   ErrorObj.SystemErrorMsg = mstrErrMsg
   Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
   Set Rs = Nothing
End Function

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
   Vsp.CurrentY = dY(0)
   
   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   Vsp.TextAlign = taCenterBottom
   Vsp.Paragraph = MODULE_DESC
'   Vsp.Paragraph = ""
   Call printHeaderDetail
   
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






