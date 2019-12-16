VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportVoucherPay2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportVoucherPay2"
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

Private Const TITLE_FORMAT = "<32|>3|<2;"
Private Const HEADER_FORMAT = "^1|^6|^2|^2|^3;"
Private Const DETAIL_FORMAT = "^1|<6|>2|>2|>3;"
Private Const FOOTER_FORMAT = "^1|>6|>2|>2|>3;"

Private m_Enterprise As CEnterprise
Private m_CustAddress As CAddress
Private m_EnpAddress As CAddress
Private m_EnpPerson As cName
Private Sp As CSystemParam
Private m_ReportConfig As CReportConfig
Private m_cData As CBillingPayment
Private insRow As Long
Dim Total1(5) As Double
Dim DocumentType As Long

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
   
   Dim TempY As Long
   Dim TempSize As Long
   Dim Pc As Long
   
    tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = False
   Vsp.TableBorder = tbNone
   Vsp.TextAlign = taLeftBottom
   
   mblnHeader = True
   Vsp.MarginLeft = 0
   
   
   Vsp.DrawPicture mcolParam("BACK_GROUND"), "0.0cm", "0.0cm", Vsp.PageWidth, Vsp.PageHeight, 10
   
   If DocumentType = 111 Then
      insRow = 11
      Vsp.MarginLeft = 1 * 567 - m_ReportConfig.DUMMY_OFFSET - 95
      Vsp.TableBorder = tbNone
      '--------------------------------------------------------------------------------------------------------------------------------------------
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 1
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight) = 1.75 * 567 + m_ReportConfig.HEAD_OFFSET
      Vsp.TableCell(tcColWidth, , 1) = "19.1cm"
      Vsp.EndTable
      '--------------------------------------------------------------------------------------------------------------------------------------------
   
      TempSize = Vsp.FontSize
      Vsp.FontSize = 8
      TempY = Vsp.CurrentY
      Vsp.CurrentY = Vsp.CurrentY - 400
      Vsp.TextAlign = taLeftBottom
      strFormat = VSP_CalTable(TITLE_FORMAT, mdblWid, alngX)
     
      Pc = getPage(m_cData.GlDetails.Count, insRow)
      If Pc > 1 Then
         Call Vsp.AddTable(strFormat, "", "|หน้าที่|: " & Vsp.CurrentPage & "/" & Pc)
      Else
         Call Vsp.AddTable(strFormat, "", "")
      End If
      
      
      Vsp.CurrentY = TempY + 30
      Vsp.FontSize = TempSize
      '--------------------------------------------------------------------------------------------------------------------------------------------
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 2
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight) = "0.8cm"
      Vsp.TableCell(tcColWidth, , 1) = "15.55cm"
      Vsp.TableCell(tcColWidth, , 2) = "3.6cm"
      Vsp.TableCell(tcAlign, 1, 2) = taRightTop
      Vsp.TableCell(tcText, 1, 2) = m_cData.DOCUMENT_NO
      Vsp.EndTable
      '--------------------------------------------------------------------------------------------------------------------------------------------
   
         '--------------------------------------------------------------------------------------------------------------------------------------------
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 2
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight) = "0.8cm"
      Vsp.TableCell(tcColWidth, , 1) = "14.35cm"
      Vsp.TableCell(tcColWidth, , 2) = "5cm"
      Vsp.EndTable
      '--------------------------------------------------------------------------------------------------------------------------------------------
     Vsp.CurrentY = Vsp.CurrentY - 100
      '--------------------------------------------------------------------------------------------------------------------------------------------
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 2
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight) = "0.8cm"
      Vsp.TableCell(tcColWidth, , 1) = "12.75cm"
      Vsp.TableCell(tcColWidth, , 2) = "5cm"
      Vsp.TableCell(tcAlign, 1, 2) = taCenterTop
      Vsp.TableCell(tcText, 1, 2) = DateToStringExtEx2(m_cData.DOCUMENT_DATE)
      Vsp.EndTable
      '--------------------------------------------------------------------------------------------------------------------------------------------
      Vsp.CurrentY = Vsp.CurrentY + 20
      '--------------------------------------------------------------------------------------------------------------------------------------------
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 2
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight) = "0.8cm"
      Vsp.TableCell(tcColWidth, , 1) = "2.35cm"
      Vsp.TableCell(tcColWidth, , 2) = "26.1cm"
      Vsp.TableCell(tcAlign, 1, 2) = taLeftTop
      Vsp.TableCell(tcText, 1, 2) = m_cData.PAYMENT_TO
      Vsp.EndTable
      '--------------------------------------------------------------------------------------------------------------------------------------------

         '--------------------------------------------------------------------------------------------------------------------------------------------
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 2
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight) = "0.8cm"
      Vsp.TableCell(tcColWidth, , 1) = "4.85cm"
      Vsp.TableCell(tcColWidth, , 2) = "21.1cm"
      Vsp.TableCell(tcAlign, 1, 2) = taLeftTop
      Vsp.TableCell(tcText, 1, 2) = m_cData.DOC_ASSEMBLE
      Vsp.EndTable
      
      '--------------------------------------------------------------------------------------------------------------------------------------------
      Vsp.CurrentY = Vsp.CurrentY - 30
            '--------------------------------------------------------------------------------------------------------------------------------------------
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 2
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight) = "1.4cm"
      Vsp.TableCell(tcColWidth, , 1) = "3.25cm"
      Vsp.TableCell(tcColWidth, , 2) = "15.5cm"
      Vsp.TableCell(tcAlign, 1, 2) = taLeftTop
      Vsp.TableCell(tcText, 1, 2) = m_cData.PAYMENT_COST
      Vsp.EndTable
      '--------------------------------------------------------------------------------------------------------------------------------------------
      Vsp.CurrentY = Vsp.CurrentY + 60
      '--------------------------------------------------------------------------------------------------------------------------------------------
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 2
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight) = "0.8cm"
      Vsp.TableCell(tcColWidth, , 1) = "3.25cm"
      Vsp.TableCell(tcColWidth, , 2) = "23.1cm"
      Vsp.TableCell(tcAlign, 1, 2) = taLeftTop
      Vsp.TableCell(tcText, 1, 2) = DateToStringExtEx2(m_cData.PAYMENT_DUE)
      Vsp.EndTable
      '--------------------------------------------------------------------------------------------------------------------------------------------
      Vsp.CurrentY = Vsp.CurrentY - 30
      '--------------------------------------------------------------------------------------------------------------------------------------------
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 4
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight) = "0.8cm"
      Vsp.TableCell(tcColWidth, , 1) = "4.25cm"
      Vsp.TableCell(tcColWidth, , 2) = "5.2cm"
      Vsp.TableCell(tcColWidth, , 3) = "4.8cm"
      Vsp.TableCell(tcColWidth, , 4) = "4.8cm"
      Vsp.TableCell(tcAlign, 1, 2) = taLeftTop
      Vsp.TableCell(tcAlign, 1, 3) = taLeftTop
      Vsp.TableCell(tcAlign, 1, 4) = taLeftTop
      Vsp.TableCell(tcText, 1, 2) = m_cData.PAYMENT_PART
      Vsp.TableCell(tcText, 1, 3) = m_cData.PAYMENT_DEPT
      Vsp.TableCell(tcText, 1, 4) = m_cData.PAYMENT_DEPT2
      Vsp.EndTable
      '--------------------------------------------------------------------------------------------------------------------------------------------
      Vsp.CurrentY = Vsp.CurrentY - 30
      '--------------------------------------------------------------------------------------------------------------------------------------------
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 3
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight) = "0.8cm"
      Vsp.TableCell(tcColWidth, , 1) = "2.55cm"
      Vsp.TableCell(tcColWidth, , 2) = "4.5cm"
      Vsp.TableCell(tcColWidth, , 3) = "11cm"
      Vsp.TableCell(tcAlign, 1, 2) = taLeftTop
      Vsp.TableCell(tcAlign, 1, 3) = taLeftTop
      Vsp.TableCell(tcText, 1, 2) = FormatNumber(m_cData.PAYMENT_AMOUNT)
      Vsp.TableCell(tcText, 1, 3) = ThaiBaht(m_cData.PAYMENT_AMOUNT)
      Vsp.EndTable
      '--------------------------------------------------------------------------------------------------------------------------------------------

      '--------------------------------------------------------------------------------------------------------------------------------------------
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 2
      Vsp.TableCell(tcRows) = 2
      Vsp.TableCell(tcRowHeight) = "0.8cm"
'      Vsp.TableCell(tcRowHeight) = "1cm"
      Vsp.TableCell(tcColWidth, , 1) = "2.2cm"
      Vsp.TableCell(tcColWidth, , 2) = "4.5cm"
      Vsp.TableCell(tcAlign, 1, 2) = taLeftTop
      If m_cData.PAYMENT_BY = 1 Then
         Vsp.TableCell(tcText, 1, 2) = "X"
      ElseIf m_cData.PAYMENT_BY = 2 Then
        Vsp.CurrentY = Vsp.CurrentY - 30
         Vsp.TableCell(tcText, 2, 2) = "X"
      End If
      Vsp.EndTable
      '--------------------------------------------------------------------------------------------------------------------------------------------
   ElseIf DocumentType = 112 Then
      insRow = 19
     Vsp.MarginLeft = 1 * 567 - m_ReportConfig.DUMMY_OFFSET + 180
      Vsp.TableBorder = tbNone
      '--------------------------------------------------------------------------------------------------------------------------------------------
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 1
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight) = 1.75 * 567 + m_ReportConfig.HEAD_OFFSET + 920
      Vsp.TableCell(tcColWidth, , 1) = "19.1cm"
      Vsp.EndTable
      '--------------------------------------------------------------------------------------------------------------------------------------------
   
      TempSize = Vsp.FontSize
      Vsp.FontSize = 8
      TempY = Vsp.CurrentY
      Vsp.CurrentY = Vsp.CurrentY - 1300
      Vsp.TextAlign = taLeftBottom
      strFormat = VSP_CalTable(TITLE_FORMAT, mdblWid, alngX)
     
      Pc = getPage(m_cData.GlDetails.Count, insRow)
      If Pc > 1 Then
         Call Vsp.AddTable(strFormat, "", "|หน้าที่|: " & Vsp.CurrentPage & "/" & Pc)
      Else
         Call Vsp.AddTable(strFormat, "", "")
      End If
      
      
      Vsp.CurrentY = TempY
      Vsp.FontSize = TempSize
      '--------------------------------------------------------------------------------------------------------------------------------------------
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 2
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight) = "0.8cm"
      Vsp.TableCell(tcColWidth, , 1) = "14cm"
      Vsp.TableCell(tcColWidth, , 2) = "3.6cm"
      Vsp.TableCell(tcAlign, 1, 2) = taRightTop
      Vsp.TableCell(tcText, 1, 2) = m_cData.DOCUMENT_NO_JV
      Vsp.EndTable
      '--------------------------------------------------------------------------------------------------------------------------------------------
   
         '--------------------------------------------------------------------------------------------------------------------------------------------
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 2
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight) = "0.8cm"
      Vsp.TableCell(tcColWidth, , 1) = "15.35cm"
      Vsp.TableCell(tcColWidth, , 2) = "5cm"
      Vsp.EndTable
      '--------------------------------------------------------------------------------------------------------------------------------------------
     Vsp.CurrentY = Vsp.CurrentY - 140
      '--------------------------------------------------------------------------------------------------------------------------------------------
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 2
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight) = "1.0cm"
      Vsp.TableCell(tcColWidth, , 1) = "14cm"
      Vsp.TableCell(tcColWidth, , 2) = "5cm"
      Vsp.TableCell(tcAlign, 1, 2) = taLeftTop
      Vsp.TableCell(tcText, 1, 2) = DateToStringExtEx2(m_cData.DOCUMENT_DATE)
      Vsp.EndTable
      '--------------------------------------------------------------------------------------------------------------------------------------------
   End If
   
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

   Vsp.PaperSize = pprA4
   Vsp.ORIENTATION = orPortrait
   Vsp.MarginBottom = "1cm"
   Vsp.MarginFooter = "1cm"
   Vsp.MarginHeader = "1cm"
   Vsp.MarginLeft = "1cm"
   Vsp.MarginRight = "1cm"
   Vsp.MarginTop = "0cm"
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE

   

'   MODULE_DESC = MapText("ใบสำคัญจ่าย")
   
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
Dim iCount As Long
Dim I As Long
Dim J As Long
Dim strFormat As String
Dim strPrint As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Poi As CReceiptItem
Dim TempBorder As TableBorderSettings
Dim Addr As CAddress
Dim SumDoc As String
Dim Ct As CCashTran
Dim TotalReceive As Double
Dim TotalFee As Double
Dim TempStr As String
Dim TempY As Double
Dim TempY2 As Double

Dim Gl As CGLDetail
Dim SumDr As Double
Dim SumCr As Double
Dim SumTotalPrice As Double

   RName = "genDoc"

'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   Set Rs = New ADODB.Recordset
'   Set m_cData = New CBillingDoc
   Set m_cData = mcolParam("m_BillingPayment")
   DocumentType = mcolParam("DOCUMENT_TYPE")
   If Not initDoc Then Exit Function
   Vsp.Preview = isPreview
''-----------------------------------------------------------------------------------------------------
''                                         Main Operation Here
''-----------------------------------------------------------------------------------------------------
   For J = 1 To UBound(Total1)
      Total1(J) = 0
   Next J
  
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   
   Call printHeader
   Call GenerateDetail

   Vsp.EndDoc
   Call CloseExportFile(Vsp)

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
   
   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   Vsp.TextAlign = taCenterBottom
'   Vsp.Paragraph = MODULE_DESC
'   Vsp.Paragraph = ""

   If mblnHeader Then
      Vsp.TableBorder = tbdT
      Vsp.FontBold = blnBold
      Vsp.FontSize = iSize
      
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
Private Sub GenerateDetail()
Dim I As Long
Dim J As Long
Dim alngX() As Long
Dim strFormat As String
Dim glDetail1 As CGLDetail
Dim glDetail2 As CGLDetail

If DocumentType = 111 Then
         Vsp.TableBorder = tbNone
         Vsp.FontSize = 14
         Vsp.StartTable
         Vsp.TableCell(tcCols) = 4
         Vsp.TableCell(tcRows) = insRow
          
          Vsp.CurrentY = Vsp.CurrentY + 1240
            I = 0
            For Each glDetail1 In m_cData.GlDetails
               If I = insRow Then
                 Vsp.EndTable
                 Vsp.NewPage
                 
                  Vsp.FontSize = 14
                  Vsp.StartTable
                  Vsp.TableCell(tcCols) = 4
                  Vsp.TableCell(tcRows) = insRow
               
                  Vsp.CurrentY = Vsp.CurrentY + 1240
                  I = 0
               End If
               I = I + 1
               Vsp.TableCell(tcRowHeight, I) = 433
               
               Vsp.TableCell(tcColWidth, I, 1) = "8.25cm"
               Vsp.TableCell(tcColWidth, I, 2) = "2.75cm"
               Vsp.TableCell(tcColWidth, I, 3) = "3.5cm"
               Vsp.TableCell(tcColWidth, I, 4) = "4cm"

                Vsp.TableCell(tcAlign, I, 1) = taLeftTop
                Vsp.TableCell(tcAlign, I, 2) = taLeftTop
                Vsp.TableCell(tcAlign, I, 3) = taRightTop
                Vsp.TableCell(tcAlign, I, 4) = taRightTop
         
               If glDetail1.GetFieldValue("GL_TYPE") = 1 Then
                  Vsp.TableCell(tcText, I, 1) = glDetail1.GetFieldValue("GL_NAME")
                  Vsp.TableCell(tcText, I, 2) = IIf(glDetail1.GetFieldValue("SHOW_FLAG") = "Y", "", glDetail1.GetFieldValue("GL_NO"))
                  Vsp.TableCell(tcText, I, 3) = FormatNumber(glDetail1.GetFieldValue("GL_AMOUNT"))
                  Vsp.TableCell(tcText, I, 4) = ""
                  Total1(3) = Total1(3) + glDetail1.GetFieldValue("GL_AMOUNT")
               ElseIf glDetail1.GetFieldValue("GL_TYPE") = 2 Then
                  Vsp.TableCell(tcText, I, 1) = "     " & glDetail1.GetFieldValue("GL_NAME")
                  Vsp.TableCell(tcText, I, 2) = "     " & IIf(glDetail1.GetFieldValue("SHOW_FLAG") = "Y", "", glDetail1.GetFieldValue("GL_NO"))
                  Vsp.TableCell(tcText, I, 3) = ""
                  Vsp.TableCell(tcText, I, 4) = FormatNumber(glDetail1.GetFieldValue("GL_AMOUNT"))
                  Total1(4) = Total1(4) + glDetail1.GetFieldValue("GL_AMOUNT")
               End If
               Vsp.CurrentY = Vsp.CurrentY + 30
          Next glDetail1
          If I < insRow Then
            J = I + 1
            For I = J To insRow
               Vsp.TableCell(tcRowHeight, I) = 433
               
               Vsp.TableCell(tcColWidth, I, 1) = "8.25cm"
               Vsp.TableCell(tcColWidth, I, 2) = "2.75cm"
               Vsp.TableCell(tcColWidth, I, 3) = "3.5cm"
               Vsp.TableCell(tcColWidth, I, 4) = "4cm"

                Vsp.TableCell(tcAlign, I, 1) = taLeftTop
                Vsp.TableCell(tcAlign, I, 2) = taLeftTop
                Vsp.TableCell(tcAlign, I, 3) = taRightTop
                Vsp.TableCell(tcAlign, I, 4) = taRightTop
         
                  Vsp.TableCell(tcText, I, 1) = ""
                  Vsp.TableCell(tcText, I, 2) = ""
                  Vsp.TableCell(tcText, I, 3) = ""
                  Vsp.TableCell(tcText, I, 4) = ""
               Vsp.CurrentY = Vsp.CurrentY + 30
            Next I
          End If
          Vsp.EndTable
          
  '--------------------------------------------------------------------------------------------------------------------------------------------
  'Sum footter
   Vsp.CurrentY = Vsp.CurrentY + 170
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 3
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcRowHeight) = "0.8cm"
   Vsp.TableCell(tcColWidth, 1, 1) = "11cm"
   Vsp.TableCell(tcColWidth, 1, 2) = "3.5cm"
   Vsp.TableCell(tcColWidth, 1, 3) = "4cm"
   
   Vsp.TableCell(tcAlign, 1, 1) = taLeftTop
   Vsp.TableCell(tcAlign, 1, 2) = taRightTop
   Vsp.TableCell(tcAlign, 1, 3) = taRightTop
   
   If Total1(3) = Total1(4) Then
      Vsp.TableCell(tcText, 1, 1) = "         " & ThaiBaht(Total1(3))
   Else
      Vsp.TableCell(tcText, 1, 1) = "         *** Debit ไม่เท่ากับ Credit ***"
   End If
   Vsp.TableCell(tcText, 1, 2) = FormatNumber(Total1(3))
   Vsp.TableCell(tcText, 1, 3) = FormatNumber(Total1(4))
   
   Vsp.EndTable
   '--------------------------------------------------------------------------------------------------------------------------------------------
   ElseIf DocumentType = 112 Then
         Vsp.TableBorder = tbNone
         Vsp.FontSize = 14
         Vsp.StartTable
         Vsp.TableCell(tcCols) = 4
         Vsp.TableCell(tcRows) = insRow
          
          Vsp.CurrentY = Vsp.CurrentY + 350
          
            I = 0
            For Each glDetail1 In m_cData.GlDetails2
               If I = insRow Then
                 Vsp.EndTable
                 Vsp.NewPage
                 
                  Vsp.FontSize = 14
                  Vsp.StartTable
                  Vsp.TableCell(tcCols) = 4
                  Vsp.TableCell(tcRows) = insRow
               
                  Vsp.CurrentY = Vsp.CurrentY + 350
                  I = 0
               End If
               I = I + 1
               Vsp.TableCell(tcRowHeight) = 412
               
               Vsp.TableCell(tcColWidth, I, 1) = "2.75cm"
               Vsp.TableCell(tcColWidth, I, 2) = "7.25cm"
               Vsp.TableCell(tcColWidth, I, 3) = "3.5cm"
               Vsp.TableCell(tcColWidth, I, 4) = "4.2cm"

                Vsp.TableCell(tcAlign, I, 1) = taLeftTop
                Vsp.TableCell(tcAlign, I, 2) = taLeftTop
                Vsp.TableCell(tcAlign, I, 3) = taRightTop
                Vsp.TableCell(tcAlign, I, 4) = taRightTop
         
               If glDetail1.GetFieldValue("GL_TYPE") = 1 Then
                  Vsp.TableCell(tcText, I, 1) = IIf(glDetail1.GetFieldValue("SHOW_FLAG") = "Y", "", glDetail1.GetFieldValue("GL_NO"))
                  Vsp.TableCell(tcText, I, 2) = glDetail1.GetFieldValue("GL_NAME")
                  Vsp.TableCell(tcText, I, 3) = FormatNumber(glDetail1.GetFieldValue("GL_AMOUNT"))
                  Vsp.TableCell(tcText, I, 4) = ""
                  Total1(3) = Total1(3) + glDetail1.GetFieldValue("GL_AMOUNT")
               ElseIf glDetail1.GetFieldValue("GL_TYPE") = 2 Then
                  Vsp.TableCell(tcText, I, 1) = "     " & IIf(glDetail1.GetFieldValue("SHOW_FLAG") = "Y", "", glDetail1.GetFieldValue("GL_NO"))
                  Vsp.TableCell(tcText, I, 2) = "     " & glDetail1.GetFieldValue("GL_NAME")
                  Vsp.TableCell(tcText, I, 3) = ""
                  Vsp.TableCell(tcText, I, 4) = FormatNumber(glDetail1.GetFieldValue("GL_AMOUNT"))
                  Total1(4) = Total1(4) + glDetail1.GetFieldValue("GL_AMOUNT")
               End If
               Vsp.CurrentY = Vsp.CurrentY + 30
          Next glDetail1
          If I < insRow Then
            J = I + 1
            For I = J To insRow
               Vsp.TableCell(tcRowHeight, I) = 412
               
               Vsp.TableCell(tcColWidth, I, 1) = "2.75cm"
               Vsp.TableCell(tcColWidth, I, 2) = "7.25cm"
               Vsp.TableCell(tcColWidth, I, 3) = "3.5cm"
               Vsp.TableCell(tcColWidth, I, 4) = "4.2cm"

                Vsp.TableCell(tcAlign, I, 1) = taLeftTop
                Vsp.TableCell(tcAlign, I, 2) = taLeftTop
                Vsp.TableCell(tcAlign, I, 3) = taRightTop
                Vsp.TableCell(tcAlign, I, 4) = taRightTop

                  Vsp.TableCell(tcText, I, 1) = ""
                  Vsp.TableCell(tcText, I, 2) = ""
                  Vsp.TableCell(tcText, I, 3) = ""
                  Vsp.TableCell(tcText, I, 4) = ""
               Vsp.CurrentY = Vsp.CurrentY + 30
            Next I
          End If
          Vsp.EndTable
          
  '--------------------------------------------------------------------------------------------------------------------------------------------
  'Sum footter
   Vsp.CurrentY = Vsp.CurrentY
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 3
   Vsp.TableCell(tcRows) = 1

   Vsp.TableCell(tcRowHeight, I) = 413 '433
   Vsp.TableCell(tcColWidth, 1, 1) = "10cm"
   Vsp.TableCell(tcColWidth, 1, 2) = "3.5cm"
   Vsp.TableCell(tcColWidth, 1, 3) = "4.2cm"
   
   Vsp.TableCell(tcAlign, 1, 1) = taLeftTop
   Vsp.TableCell(tcAlign, 1, 2) = taRightTop
   Vsp.TableCell(tcAlign, 1, 3) = taRightTop
   
   Vsp.TableCell(tcText, 1, 1) = ""
   Vsp.TableCell(tcText, 1, 2) = FormatNumber(Total1(3))
   Vsp.TableCell(tcText, 1, 3) = FormatNumber(Total1(4))
   
   Vsp.EndTable
   '--------------------------------------------------------------------------------------------------------------------------------------------
   
     '--------------------------------------------------------------------------------------------------------------------------------------------
     'Note
   Vsp.CurrentY = Vsp.CurrentY + 110
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 1
   
   Vsp.TableCell(tcRowHeight, 1) = 2000 '433
   Vsp.TableCell(tcColWidth, , 1) = "2.7cm"
   Vsp.TableCell(tcColWidth, , 2) = "15.5cm"
   Vsp.TableCell(tcAlign, , 1) = taLeftTop
   Vsp.TableCell(tcAlign, , 2) = taLeftTop
   Vsp.TableCell(tcText, 1, 2) = m_cData.NOTE
   Vsp.EndTable
   '--------------------------------------------------------------------------------------------------------------------------------------------
   End If
   End Sub
   Private Function getPage(mRow As Long, insRow As Long) As Long
      If mRow Mod insRow <> 0 Then
         getPage = Fix(mRow / insRow) + 1
      Else
         getPage = Fix(mRow / insRow)
      End If
   End Function
