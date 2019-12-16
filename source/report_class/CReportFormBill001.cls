VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportFormBill001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportFormDO001"
Private MODULE_DESC As String

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
Private Const HEADER_SIZE = 14
Private Const DETAIL_SIZE = 14

Private Const TITLE_FORMAT = "<32|<3|<5;"
Private Const HEADER_FORMAT = "^1|^6|^2|^2|^3;"
Private Const DETAIL_FORMAT = "^1|<6|>2|>2|>3;"
Private Const FOOTER_FORMAT = "^1|>6|>2|>2|>3;"

Private m_Enterprise As CEnterprise
Private m_CustAddress As CAddress
Private m_EnpAddress As CAddress
Private m_EnpPerson As cName
Private Sp As CSystemParam
Private m_ReportConfig As CReportConfig
Private m_cData As CBillingDoc
Private m_ImageWidth As Double
Private m_ImageHeight As Double

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_Enterprise = New CEnterprise
   Set m_CustAddress = New CAddress
   Set m_EnpAddress = New CAddress
   Set m_EnpPerson = New cName
   Set m_ReportConfig = New CReportConfig
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_Enterprise = Nothing
   Set m_CustAddress = Nothing
   Set m_EnpAddress = Nothing
   Set m_EnpPerson = Nothing
   Set m_ReportConfig = Nothing
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
   Dim i As Long
   
   For i = 1 To mcolParam.Count
      mcolParam.Remove 1
   Next i

End Sub
'============================= Codes above used in every report ==========================

Private Sub PrintPOFooter(alngX() As Long, Bd As CBillingDoc)
   If glbParameterObj.Programowner = XEROX_OWNER Then
      Call GenerateDocFooterXerox(2, Vsp, mcolParam, tbNone, Bd)
   ElseIf glbParameterObj.Programowner = PLAZA_OWNER Then
      Call GenerateDocFooterPlaza(2, Vsp, mcolParam, tbNone, Bd)
  ElseIf glbParameterObj.Programowner = FACTORY_OWNER Then
      Call GenerateDocFooterFACTORY2(2, Vsp, mcolParam, tbNone, Bd)
   End If
End Sub

Private Sub PrintPOHeader(alngX() As Long)
Dim strFormat As String
Dim strPrint As String
Static i As Long
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim OwnerFlag As String
Dim TempStr As String
Dim TempBorder As TableBorderSettings
Dim TempMarginLeft As Double

   TempMarginLeft = Vsp.MarginLeft
   Vsp.MarginLeft = 0 * 567 - m_ReportConfig.DUMMY_OFFSET
'   Vsp.MarginLeft = 0
   
   If glbParameterObj.Programowner = XEROX_OWNER Then
      Call GenerateDocHeaderXerox(1, Vsp, mcolParam, tbNone, m_ReportConfig.HEAD_OFFSET)
   ElseIf glbParameterObj.Programowner = PLAZA_OWNER Then
      Vsp.MarginLeft = 576 * 5 - m_ReportConfig.DUMMY_OFFSET
      Call GenerateDocHeaderPlaza(1, Vsp, mcolParam, tbNone, m_ReportConfig.HEAD_OFFSET)
    ElseIf glbParameterObj.Programowner = FACTORY_OWNER Then
      Vsp.MarginLeft = 576 * 6.8 - m_ReportConfig.DUMMY_OFFSET
      Call GenerateDocHeaderFactory2(1, Vsp, mcolParam, tbNone, m_ReportConfig.HEAD_OFFSET, m_cData)
   Else
      Call GenerateDocHeaderXerox(1, Vsp, mcolParam, tbNone, m_ReportConfig.HEAD_OFFSET)
    End If
    
    Vsp.MarginLeft = TempMarginLeft
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
   
   Call PrintPOHeader(alngX)
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbNone
   strFormat = VSP_CalTable(HEADER_FORMAT, mdblWid, alngX)
   strPrint = "ลำดับ|สินค้าและบริการ|จำนวน|ราคาต่อหน่วย|จำนวนเงิน;"
'   Call Vsp.AddTable(strFormat, "", strPrint)
   
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
      
   If glbParameterObj.Programowner = XEROX_OWNER Then
         Vsp.MarginLeft = "1cm"
         Vsp.MarginRight = "1cm"
   Else
      m_ImageWidth = 16.3 * 567
      m_ImageHeight = 25.4 * 567
      Vsp.MarginLeft = (Vsp.PageWidth - m_ImageWidth) \ 2
      Vsp.MarginRight = (Vsp.PageHeight - m_ImageWidth) \ 2
   End If
   
   Vsp.PaperSize = pprA4
'   Vsp.PaperWidth = 6.1 * 1440
'   Vsp.PaperHeight = 8.5 * 1440
   Vsp.ORIENTATION = orPortrait
   Vsp.MarginBottom = "1cm"
   Vsp.MarginFooter = "1cm"
   Vsp.MarginHeader = "1cm"
   Vsp.MarginTop = "0cm"
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   MODULE_DESC = MapText("ใบรับงาน/สั่งงาน (PO)")
   
   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"), m_ReportConfig, False)
'   If mcolParam("REPORT_TYPE") = 4 Then
'      m_ReportConfig.HEAD_OFFSET = 0
'      m_ReportConfig.DeleteData = 0
'   End If
   
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
Dim iCount As Long
Dim i As Long
Dim strFormat As String
Dim strPrint As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Poi As CDoItem
Dim Sum1 As Double
Dim Sum2 As Double
Dim Sum3 As Double
Dim j As Long
Dim TempBorder As TableBorderSettings
Dim Addr As CAddress
Dim SumDiscount As Double

   RName = "genDoc"
   
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   Set Rs = New ADODB.Recordset
   Set m_cData = New CBillingDoc
   m_cData.BILLING_DOC_ID = mcolParam("BILLING_DOC_ID")
   m_cData.QueryFlag = 1
   If Not glbDaily.QueryBillingDoc(m_cData, Rs, iCount, IsOK, glbErrorLog) Then
      mstrErrMsg = "Query Error."
      Exit Function
   End If
      
   Call m_cData.PopulateFromRS(1, Rs)
   Call mcolParam.add(m_cData.DOCUMENT_NO, "DOCUMENT_NO")
   Call mcolParam.add(m_cData.DOCUMENT_DATE, "DOCUMENT_DATE")
   Call mcolParam.add(m_cData.ESTIMATE_DATE, "ESTIMATE_DATE")
   Call mcolParam.add(m_cData.APPROVE_DATE, "APPROVE_DATE")
   Call mcolParam.add(m_cData.CUSTOMER_CODE, "CUSTOMER_CODE")
   Call mcolParam.add(m_cData.CUSTOMER_NAME, "CUSTOMER_NAME")
   Call mcolParam.add(m_cData.DEPOSIT_AMOUNT, "DEPOSIT_AMOUNT")
   Call mcolParam.add(m_cData.RESOURCE_NO, "RESOURCE_NO")
   Call mcolParam.add(m_cData.PAYMENT, "PAYMENT_DATE")
   If m_cData.BILLING_ADDRESS_ID > 0 Then
      Set Addr = New CAddress
      Addr.ADDRESS_ID = m_cData.BILLING_ADDRESS_ID
      Call Addr.QueryData(Rs, iCount)
      If Not Rs.EOF Then
         Call Addr.PopulateFromRS(Rs)
         Call mcolParam.add(Addr.PackAddress, "CUSTOMER_ADDRESS")
          Call mcolParam.add(Addr.PHONE1, "PHONE")
          Call mcolParam.add(Addr.FAX1, "FAX")
      End If
      Set Addr = Nothing
   Else
      Call mcolParam.add("", "CUSTOMER_ADDRESS")
   End If
   
   If (Rs Is Nothing) Or (iCount < 1) Then
      mstrErrMsg = "ไม่พบข้อมูลที่ต้องการในฐานข้อมูล"
      Exit Function
   End If
   
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
   Sum1 = 0
   Sum2 = 0
   Sum3 = 0
   
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Vsp.MarginLeft = 0
   
      strFormat = VSP_CalTable(DETAIL_FORMAT, mdblWid, alngX)
      TempBorder = tbNone
      If glbParameterObj.Programowner = XEROX_OWNER Then
         Vsp.MarginLeft = 0 * 567 - m_ReportConfig.DUMMY_OFFSET
      ElseIf glbParameterObj.Programowner = PLAZA_OWNER Then
         Vsp.MarginLeft = 576 * 0.72 - m_ReportConfig.DUMMY_OFFSET
        ElseIf glbParameterObj.Programowner = FACTORY_OWNER Then
         Vsp.MarginLeft = 576 * 3.2 - m_ReportConfig.DUMMY_OFFSET
      End If
      
      i = 0
      j = 0
      SumDiscount = 0

      
      If glbParameterObj.Programowner = XEROX_OWNER Then
         Call GenerateReceiptBodyXerox(1, Vsp, mcolParam, m_cData, tbNone)
      ElseIf glbParameterObj.Programowner = FACTORY_OWNER Then
         Call GenerateReceiptBodyFACTORY2(1, Vsp, mcolParam, m_cData, tbNone)
      Else
         Call GenerateReceiptBodyPlaza(1, Vsp, mcolParam, m_cData, tbNone)
      End If
      Call PrintPOFooter(alngX, m_cData)
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   Rs.Close
   Set Rs = Nothing
   Set m_cData = Nothing
   
   genDoc = True
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
'   ErrorObj.RoutineName = RName
'   ErrorObj.ModuleName = MODULE_NAME
'   ErrorObj.LocalErrorMsg = "Run time error."
'   ErrorObj.SystemErrorMsg = mstrErrMsg
'   Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
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
   
'   Vsp.DrawPicture mcolParam("PICTURE"), "0.75in", "0.75in", "100%", "100%"
   
   Vsp.FontSize = 8
   dY(0) = Vsp.CurrentY
   Vsp.TextAlign = taLeftBottom
   strFormat = VSP_CalTable(TITLE_FORMAT, mdblWid, alngX)
   Vsp.TableBorder = tbNone
'   Call Vsp.AddTable(strFormat, "", "|หน้าที่|: " & Vsp.CurrentPage)
'   Call Vsp.AddTable(strFormat, "", "|วันที่พิมพ์|: " & DateToStringExt(mdteDate))
'   Call Vsp.AddTable(strFormat, "", "|เวลา|: " & Right$(DateToStringExtEx(mdteDate), 8))
   Vsp.CurrentY = dY(0)
   
   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   Vsp.TextAlign = taCenterBottom
'   Vsp.Paragraph = MODULE_DESC
'   Vsp.Paragraph = ""
   If mcolParam("REPORT_TYPE") = 4 Then
   If glbParameterObj.Programowner = XEROX_OWNER Then
      Vsp.DrawPicture mcolParam("BACK_GROUND"), "0.05cm", "0.0cm", "21.55cm", "15.5cm", 10
   ElseIf glbParameterObj.Programowner = FACTORY_OWNER Then
      Vsp.DrawPicture mcolParam("BACK_GROUND"), Vsp.MarginLeft, Vsp.MarginTop, m_ImageWidth, m_ImageHeight, 10
   Else
      Vsp.DrawPicture mcolParam("BACK_GROUND"), Vsp.MarginLeft, Vsp.MarginTop, m_ImageWidth, m_ImageHeight, 10
   End If
   End If
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


