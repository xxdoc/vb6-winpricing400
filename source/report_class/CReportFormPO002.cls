VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportFormPO002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportFormPO002"
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
Private m_EnpPerson As CName
Private Sp As CSystemParam
Private m_ReportConfig As CReportConfig
Private m_cData As CBillingDoc

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_Enterprise = New CEnterprise
   Set m_CustAddress = New CAddress
   Set m_EnpAddress = New CAddress
   Set m_EnpPerson = New CName
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

Private Sub PrintPOHeader(alngX() As Long)
Dim strFormat As String
Dim strPrint As String
Static I As Long
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim OwnerFlag As String
Dim TempStr As String
Dim TempBorder As TableBorderSettings
Dim TempMarginLeft As Double

   TempMarginLeft = Vsp.MarginLeft
   Vsp.MarginLeft = 0 * 567 - m_ReportConfig.DUMMY_OFFSET
'   Vsp.MarginLeft = 0
   
   Call GenerateDocHeader(1, Vsp, mcolParam, tbNone, m_ReportConfig.HEAD_OFFSET)
    
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
   
   Vsp.PaperSize = pprA4
   Vsp.ORIENTATION = orPortrait
   Vsp.MarginBottom = "1cm"
   Vsp.MarginFooter = "1cm"
   Vsp.MarginHeader = "1cm"
   Vsp.MarginLeft = "1cm"
   Vsp.MarginRight = "1cm"
   Vsp.MarginTop = "1cm"
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
Dim I As Long
Dim strFormat As String
Dim strPrint As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Poi As CSupItem
Dim Sum1 As Double
Dim Sum2 As Double
Dim Sum3 As Double
Dim J As Long
Dim TempBorder As TableBorderSettings
Dim Addr As CAddress

Dim Amt As Double

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
   Call mcolParam.add(m_cData.DUE_DATE, "DUE_DATE")
   Call mcolParam.add(m_cData.SUPPLIER_CODE, "SUPPLIER_CODE")
   Call mcolParam.add(m_cData.SUPPLIER_NAME, "SUPPLIER_NAME")
   Call mcolParam.add(m_cData.Credit, "CREDIT")
   
   Call mcolParam.add(m_cData.SUPPLIER_CODE_TRUE, "SUPPLIER_CODE_TRUE")
   Call mcolParam.add(m_cData.SUPPLIER_NAME_TRUE, "SUPPLIER_NAME_TRUE")
   Call mcolParam.add(m_cData.CreditTrue, "CREDIT_TRUE")
   
   Call mcolParam.add(m_cData.DUE_AMOUNT, "DUE_AMOUNT")
   
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
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
   
   Vsp.MarginLeft = 0
   
      strFormat = VSP_CalTable(DETAIL_FORMAT, mdblWid, alngX)
      TempBorder = tbNone
      Vsp.MarginLeft = 0 * 567 - m_ReportConfig.DUMMY_OFFSET
      
      I = 0
      J = 0
      For Each Poi In m_cData.SupItems
         I = I + 1
         J = J + 1
         
         If J > 4 Then
            J = 1
            Vsp.NewPage
         End If
         
         Vsp.StartTable
         Vsp.TableCell(tcCols) = 7
         Vsp.TableCell(tcRows) = 1
          Vsp.TableCell(tcRowHeight) = "0.7cm" '95
          Vsp.TableCell(tcColWidth, , 1) = "1.2cm"
          Vsp.TableCell(tcColWidth, , 2) = "7.8cm"
          Vsp.TableCell(tcColWidth, , 3) = "2.8cm"
          Vsp.TableCell(tcColWidth, , 4) = "1.8cm"
          Vsp.TableCell(tcColWidth, , 5) = "2.3cm"
          Vsp.TableCell(tcColWidth, , 6) = "2.7cm"
          Vsp.TableCell(tcColWidth, , 7) = "0.8cm"
          
          Vsp.TableCell(tcAlign, 1, 1) = taCenterMiddle
          Vsp.TableCell(tcText, 1, 1) = I
          Vsp.TableBorder = TempBorder
          
          Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
          Vsp.TableCell(tcText, 1, 2) = Poi.PART_DESC & " ( " & Poi.PART_DESC & " ) "
          Vsp.TableBorder = TempBorder
          
          Vsp.TableCell(tcAlign, 1, 3) = taLeftMiddle
          
         If m_cData.DUE_AMOUNT > 0 Then
            If Month(m_cData.DUE_DATE) = Month(DateAdd("D", m_cData.DUE_AMOUNT - 1, m_cData.DUE_DATE)) Then
               If m_cData.DUE_DATE > 0 And DateAdd("D", m_cData.DUE_AMOUNT - 1, m_cData.DUE_DATE) > 0 Then
                  Vsp.TableCell(tcText, 1, 3) = Day(m_cData.DUE_DATE) & "-" & DateToStringExtEx4(DateAdd("D", m_cData.DUE_AMOUNT - 1, m_cData.DUE_DATE))
               Else
                  Vsp.TableCell(tcText, 1, 3) = ""
               End If
            Else
               Vsp.TableCell(tcText, 1, 3) = DateToStringExtEx4(m_cData.DUE_DATE) & "-" & DateToStringExtEx4(DateAdd("D", m_cData.DUE_AMOUNT - 1, m_cData.DUE_DATE))
            End If
         Else
            Vsp.TableCell(tcText, 1, 3) = DateToStringExtEx4(mcolParam("DUE_DATE"))
          End If
          
          Vsp.TableBorder = TempBorder
         
          Vsp.TableCell(tcAlign, 1, 4) = taRightMiddle
          Vsp.TableCell(tcText, 1, 4) = Poi.TX_AMOUNT
          Vsp.TableBorder = TempBorder
          
          If mcolParam("REPORT_TYPE") = 10 Then
            Vsp.TableCell(tcAlign, 1, 5) = taRightMiddle
            Vsp.TableCell(tcText, 1, 5) = FormatNumber(MyDiff(Poi.INCLUDE_UNIT_PRICE * 100, 102))
            Vsp.TableBorder = TempBorder
            
            Amt = MyDiff(Poi.TOTAL_INCLUDE_PRICE * 100, 102)
            
            Vsp.TableCell(tcAlign, 1, 6) = taRightMiddle
            Vsp.TableCell(tcText, 1, 6) = Left(FormatNumber(Amt), Len(FormatNumber(Amt)) - 3)
            
            Vsp.TableCell(tcAlign, 1, 7) = taCenterMiddle
            Vsp.TableCell(tcText, 1, 7) = Right(FormatNumber(Amt), 2)
            Vsp.TableBorder = TempBorder
            
            Sum3 = Sum3 + (Amt)
          Else
            Vsp.TableCell(tcAlign, 1, 5) = taRightMiddle
            Vsp.TableCell(tcText, 1, 5) = FormatNumber(Poi.INCLUDE_UNIT_PRICE)
            Vsp.TableBorder = TempBorder
            
            Vsp.TableCell(tcAlign, 1, 6) = taRightMiddle
            Vsp.TableCell(tcText, 1, 6) = Left(FormatNumber((Poi.TOTAL_INCLUDE_PRICE)), Len(FormatNumber((Poi.TOTAL_INCLUDE_PRICE))) - 3)
            
            Vsp.TableCell(tcAlign, 1, 7) = taCenterMiddle
            Vsp.TableCell(tcText, 1, 7) = Right(FormatNumber(Poi.TOTAL_INCLUDE_PRICE), 2)
            Vsp.TableBorder = TempBorder
            
            Sum3 = Sum3 + (Poi.TOTAL_INCLUDE_PRICE)
          End If
         Vsp.EndTable
      Next Poi
      
      For I = 1 To (4 - J)
         Vsp.StartTable
         Vsp.TableCell(tcCols) = 1
         Vsp.TableCell(tcRows) = 1
          Vsp.TableCell(tcRowHeight) = "0.7cm" '0.95
          Vsp.TableCell(tcColWidth, , 1) = "19.5cm"
         Vsp.EndTable
      Next I
      
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 3
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight) = "1.2cm" '95
       Vsp.TableCell(tcColWidth, , 1) = "16cm"
       Vsp.TableCell(tcColWidth, , 2) = "2.7cm"
       Vsp.TableCell(tcColWidth, , 3) = "0.8cm"
       
       Vsp.TableCell(tcAlign, 1, 1) = taLeftMiddle
       Vsp.TableCell(tcText, 1, 1) = "       " & ThaiBaht(Sum3)
       Vsp.TableBorder = TempBorder
       
       Vsp.TableCell(tcAlign, 1, 2) = taRightMiddle
       Vsp.TableCell(tcText, 1, 2) = Left(FormatNumber((Sum3)), Len(FormatNumber((Sum3))) - 3)
       Vsp.TableBorder = TempBorder
          
      Vsp.TableCell(tcAlign, 1, 3) = taCenterMiddle
      Vsp.TableCell(tcText, 1, 3) = Right(FormatNumber(Sum3), 2)
      Vsp.TableBorder = TempBorder
      Vsp.EndTable
      
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
Public Function GenerateDocHeader(Ind As Long, Vsp As VSPrinter, mcolParam As Collection, TempBorder As TableBorderSettings, HeaderOffset As Double)
   '============
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 1
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 2.3 * 567 + HeaderOffset
    Vsp.TableCell(tcColWidth, 1) = "19.5cm"
    Vsp.TableBorder = TempBorder
    Vsp.EndTable
   '============
   
    Vsp.StartTable
   Vsp.TableCell(tcCols) = 5
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight) = "0.8cm"
    Vsp.TableCell(tcColWidth, , 1) = "1.7cm"
    Vsp.TableCell(tcColWidth, , 2) = "4.8cm"
    Vsp.TableCell(tcColWidth, , 3) = "1.2cm"
    Vsp.TableCell(tcColWidth, , 4) = "1.3cm"
    Vsp.TableCell(tcColWidth, , 5) = "10.4cm"

If mcolParam("REPORT_TYPE") = 10 Then
Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
    Vsp.TableCell(tcFontBold, 1, 2) = True
    Vsp.TableCell(tcText, 1, 2) = mcolParam("SUPPLIER_NAME_TRUE")
    Vsp.TableBorder = TempBorder
    
    Vsp.TableCell(tcAlign, 1, 4) = taCenterMiddle
    Vsp.TableCell(tcFontBold, 1, 4) = True
    Vsp.TableCell(tcText, 1, 4) = mcolParam("CREDIT_TRUE")
    Vsp.TableBorder = TempBorder
Else
    Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
    Vsp.TableCell(tcFontBold, 1, 2) = True
    Vsp.TableCell(tcText, 1, 2) = mcolParam("SUPPLIER_NAME")
    Vsp.TableBorder = TempBorder
    
    Vsp.TableCell(tcAlign, 1, 4) = taCenterMiddle
    Vsp.TableCell(tcFontBold, 1, 4) = True
    Vsp.TableCell(tcText, 1, 4) = mcolParam("CREDIT")
    Vsp.TableBorder = TempBorder
   End If
    
    
    
    Vsp.EndTable

   '============
    Vsp.StartTable
   Vsp.TableCell(tcCols) = 1
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight) = "0.8cm"
    Vsp.TableCell(tcColWidth, , 1) = "19.5cm"
    Vsp.TableBorder = TempBorder
    Vsp.EndTable
    
   '============
    Vsp.StartTable
   Vsp.TableCell(tcCols) = 3
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight) = "0.8cm"
    Vsp.TableCell(tcColWidth, , 1) = "6cm"
    Vsp.TableCell(tcColWidth, , 2) = "3.5cm"
    Vsp.TableCell(tcColWidth, , 3) = "10cm"
    
    Vsp.TableCell(tcAlign, 1, 2) = taCenterMiddle
    Vsp.TableCell(tcFontBold, 1, 2) = False
    Vsp.TableCell(tcText, 1, 2) = "" & DateToStringExtEx2(mcolParam("DOCUMENT_DATE"))
    
    Vsp.TableBorder = TempBorder
    Vsp.EndTable

   '============================
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 1
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = "1.8cm"
    Vsp.TableCell(tcColWidth, 1) = "19.5cm"
    Vsp.TableBorder = TempBorder
    Vsp.EndTable
End Function
