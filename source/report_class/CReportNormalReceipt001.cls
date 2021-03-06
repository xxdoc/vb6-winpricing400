VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportNormalRcp001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportNormalRcp001"
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

Private Const TITLE_SIZE = 22
Private Const HEADER_SIZE = 12
Private Const DETAIL_SIZE = 12

Private Const TITLE_FORMAT = "<32|<3|<5;"
Private Const HEADER_FORMAT = "^1|^7|^2|^2|^2;"
Private Const DETAIL_FORMAT = "^1|<7|>2|>2|>2;"
Private Const FOOTER_FORMAT_2 = "^1|>7|>2|>2|>2;"
Private Const FOOTER_FORMAT_1 = "^1|>7|>2|^2|>2;"

Private m_Enterprise As CEnterprise
Private m_CustAddress As CAddress
Private m_EnpAddress As CAddress
Private m_EnpPerson As cName
Private m_Addr As CAddress
Private m_cData As CBillingDoc
Private m_OwnerAddress As CAddress

Private Sp As CSystemParam

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_Enterprise = New CEnterprise
   Set m_CustAddress = New CAddress
   Set m_EnpAddress = New CAddress
   Set m_EnpPerson = New cName
   Set m_Addr = New CAddress
   Set m_OwnerAddress = New CAddress
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_Enterprise = Nothing
   Set m_CustAddress = Nothing
   Set m_EnpAddress = Nothing
   Set m_EnpPerson = Nothing
   Set m_Addr = Nothing
   Set m_OwnerAddress = Nothing
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
Dim TempCol As CCashTran
   
   Vsp.TableBorder = tbNone
   Vsp.FontBold = False
   
   For Each TempCol In m_cData.Payments
        strFormat = VSP_CalTable("<10;", mdblWid, alngX)
        If TempCol.GetFieldValue("PAYMENT_TYPE") = 1 Then
            strPrint = "������    : " & TempCol.GetFieldValue("PAYMENT_TYPE_NAME")
        ElseIf TempCol.GetFieldValue("PAYMENT_TYPE") = 2 Then
            strPrint = "������    : " & TempCol.GetFieldValue("PAYMENT_TYPE_NAME") & " ��Ҥ�� " & TempCol.GetFieldValue("BANK_NAME") & " �Ң� " & TempCol.GetFieldValue("BRANCH_NAME") & " �Ţ���ѭ�� " & TempCol.GetFieldValue("ACCOUNT_NAME")
        ElseIf TempCol.GetFieldValue("PAYMENT_TYPE") = 3 Then
            strPrint = "������    : " & TempCol.GetFieldValue("PAYMENT_TYPE_NAME") & " ��Ҥ�� " & TempCol.GetFieldValue("BANK_NAME") & " �Ң� " & TempCol.GetFieldValue("BRANCH_NAME") & " �Ţ����� " & TempCol.GetFieldValue("CHEQUE_NO") & " ŧ�ѹ��� " & DateToStringExtEx2(TempCol.GetFieldValue("CHEQUE_DATE"))
        End If
        Call Vsp.AddTable(strFormat, "", strPrint)
   Next TempCol
   
   Vsp.FontSize = 6
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Vsp.FontSize = 12
  strFormat = VSP_CalTable("<10|>10;", mdblWid, alngX)
   strPrint = "����Ѻ�ͺ�ӹҨ ..................................................................................|����Ѻ�Թ ....................................................................................;"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Vsp.FontSize = 10
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = "������Ѻ�Թ��������ó��������������繼���Ѻ�ͺ�ӹҨ ��м���Ѻ�Թ������ҹ�� �ҡ������ ����稨�����ó����������礹�� �����¡���Թ�ҡ��Ҥ�����º��������;"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Vsp.FontSize = 6
   
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Vsp.FontSize = 8
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = "�鹩�Ѻ-�١��� , ���ҷ��1-�ѭ�ա���Թ , ���ҷ���ͧ-�Դ���� ;"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
End Sub

Private Sub PrintBillHeader(alngX() As Long)
Dim strFormat As String
Dim strPrint As String
Static I As Long
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim OwnerFlag As String
Dim TempStr As String
Dim TempAddr As String

   If I = 0 Then
      I = 1
      Set TempRs = New ADODB.Recordset
      
      m_Enterprise.ENTERPRISE_ID = -1
      Call m_Enterprise.QueryData(TempRs, iCount)
      If Not TempRs.EOF Then
         Call m_Enterprise.PopulateFromRS(1, TempRs)
      End If
      
      m_EnpAddress.ADDRESS_ID = m_cData.ENTERPRISE_ADDRESS_ID
      Call m_EnpAddress.QueryData(TempRs, iCount)
      If Not TempRs.EOF Then
         Call m_EnpAddress.PopulateFromRS(TempRs)
      End If
      
      m_CustAddress.ADDRESS_ID = -1
      m_CustAddress.CUSTOMER_ID = m_cData.CUSTOMER_ID
      Call m_CustAddress.QueryData3(TempRs, iCount)
      If Not TempRs.EOF Then
         Call m_CustAddress.PopulateFromRS(TempRs)
      End If
      
      If TempRs.State = adStateOpen Then
         TempRs.Close
      End If
      Set TempRs = Nothing
    End If
   
   Vsp.TableBorder = tbNone
   Vsp.FontSize = 14
   Vsp.FontBold = False
       
    Vsp.CurrentY = Vsp.CurrentY - 300
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = m_Enterprise.ENTERPRISE_NAME & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
    
    Vsp.TableBorder = tbNone
   Vsp.FontSize = 12
   Vsp.FontBold = False
   
   strFormat = VSP_CalTable("<10|>5;", mdblWid, alngX)
   strPrint = MapText("������� ") & m_EnpAddress.PackAddress & "|�Ţ����͡��� " & m_cData.DOCUMENT_NO & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)

   strFormat = VSP_CalTable("<15|>10;", mdblWid, alngX)
   strPrint = "���Ѿ�� " & EmptyToString(m_EnpAddress.PHONE1, "N/A") & " ῡ�� " & EmptyToString(m_EnpAddress.FAX1, "N/A") & " (" & TempStr & "�Ţ������������� " & m_Enterprise.TAX_ID & ")" & "|" & " �ѹ��� " & EmptyToString(DateToStringExtEx2((m_cData.DOCUMENT_DATE)), "N/A") & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
    
    Vsp.FontSize = 8
    
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
    
    Vsp.FontSize = 12
    
   strFormat = VSP_CalTable("<10|>10;", mdblWid, alngX)
   strPrint = "�����١��� " & m_cData.CUSTOMER_NAME & "|" & "�����١���" & m_cData.CUSTOMER_CODE & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)

   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = "������� " & EmptyToString(m_CustAddress.PackAddress, "N/A") & " " & "���Ѿ�� " & EmptyToString(m_CustAddress.PHONE1, "N/A") & " ῡ�� " & EmptyToString(m_CustAddress.FAX1, "N/A") & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
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
   
   Call PrintBillHeader(alngX)
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   If m_cData.RECEIPT_TYPE = 1 Then
        strFormat = VSP_CalTable(HEADER_FORMAT, mdblWid, alngX)
        strPrint = "�ӴѺ|��¡��|�ӹǹ|�Ҥҵ��˹���|�ӹǹ�Թ;"
   Else
        strFormat = VSP_CalTable("^1|^7|^2;", mdblWid, alngX)
        strPrint = "�ӴѺ|��¡��|�ӹǹ�Թ;"
   End If
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
   Vsp.MarginBottom = 300
   Vsp.MarginFooter = 300
   Vsp.MarginHeader = 300
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 300
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   MODULE_DESC = MapText("������Ѻ�Թ")
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
Dim Bi As CDoItem
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
Dim Ri As CReceiptItem

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
   
   If (Rs Is Nothing) Or (iCount < 1) Then
      mstrErrMsg = "��辺�����ŷ���ͧ���㹰ҹ������"
      Exit Function
   End If
   
   Call m_cData.PopulateFromRS(1, Rs)
   
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
   
      
      Vsp.TableBorder = tbAll
      I = 0
                     
      TempDiscount = 0
      TempVatAmount = 0
      If m_cData.RECEIPT_TYPE = 1 Then
        strFormat = VSP_CalTable(DETAIL_FORMAT, mdblWid, alngX)
         Set TempCol = m_cData.DoItems
         For Each Bi In TempCol
            I = I + 1
            strPrint = I & "|"
            strPrint = strPrint & Bi.ShowDescText & "|"
            strPrint = strPrint & FormatNumber(Bi.ITEM_AMOUNT) & "|"
            strPrint = strPrint & FormatNumber(Bi.AVG_PRICE) & "|"
            strPrint = strPrint & FormatNumber(Bi.TOTAL_PRICE) & ";"
            Sum3 = Sum3 + Bi.TOTAL_PRICE
            TempDiscount = TempDiscount + Bi.DISCOUNT_AMOUNT
   
            If I >= TempCol.Count Then
               Call AdjustPage(Vsp, strFormat, strPrint, 4000)
            End If
            Call Vsp.AddTable(strFormat, "", strPrint)
         Next Bi
      ElseIf (m_cData.RECEIPT_TYPE = 3) Or (m_cData.RECEIPT_TYPE = 5) Then
        strFormat = VSP_CalTable("^1|<7|>2;", mdblWid, alngX)
         Set TempCol = m_cData.ReceiptItems
         For Each Ri In TempCol
            I = I + 1
            strPrint = I & "|"
            If m_cData.RECEIPT_TYPE = 3 Then
               strPrint = strPrint & "��ҧ�֧�͡��� " & Ri.DOCUMENT_NO & " (" & DateToStringExtEx2(Ri.DOCUMENT_DATE) & ")|"
            ElseIf m_cData.RECEIPT_TYPE = 5 Then
               strPrint = strPrint & "��ҧ�֧�͡��� " & Ri.DOCUMENT_NO & " (" & DateToStringExtEx2(Ri.DOCUMENT_DATE) & ")|"
            End If
            'strPrint = strPrint & FormatNumber(Ri.RECEIPT_ITEM_AMOUNT) & "|"
            'strPrint = strPrint & FormatNumber(Ri.RECEIPT_ITEM_AMOUNT - Ri.PAID_AMOUNT) & ";"
            strPrint = strPrint & FormatNumber(Ri.PAID_AMOUNT) & ";"
            
            Sum3 = Sum3 + Ri.PAID_AMOUNT
            TempDiscount = TempDiscount + Ri.DISCOUNT_AMOUNT
            
            If I <= TempCol.Count Then
               Call AdjustPage(Vsp, strFormat, strPrint, 10500)
            End If
            Call Vsp.AddTable(strFormat, "", strPrint)
         Next Ri
      End If
      Call PatchTable(Vsp, strFormat, "||||;", 10500) '10
      
      Vsp.FontBold = True
      If m_cData.RECEIPT_TYPE = 1 Then
        strPrint = "���|" & "|||" & FormatNumber(Sum3) & ";"
      Else
        strPrint = "���|" & "|" & FormatNumber(Sum3) & ";"
      End If
      Call Vsp.AddTable(strFormat, "", strPrint)
   
      If TempDiscount > 0 Then
        If m_cData.RECEIPT_TYPE = 1 Then
            strFormat = VSP_CalTable(FOOTER_FORMAT_1, mdblWid, alngX)
            strPrint = "|||��ǹŴ|" & FormatNumber(TempDiscount) & ";"
         Else
            strFormat = VSP_CalTable("^1|>7|>2;", mdblWid, alngX)
            strPrint = "|��ǹŴ|" & FormatNumber(TempDiscount) & ";"
         End If
         Call Vsp.AddTable(strFormat, "", strPrint)
      End If
         
      If m_cData.VAT_AMOUNT > 0 Then
        If m_cData.RECEIPT_TYPE = 1 Then
            strFormat = VSP_CalTable(FOOTER_FORMAT_1, mdblWid, alngX)
            strPrint = "|||����|" & FormatNumber(m_cData.VAT_AMOUNT) & ";"
         Else
            strFormat = VSP_CalTable("^1|>7|>2;", mdblWid, alngX)
            strPrint = "|����|" & FormatNumber(m_cData.VAT_AMOUNT) & ";"
         End If
         Call Vsp.AddTable(strFormat, "", strPrint)
      End If
      
      
      If m_cData.RECEIPT_TYPE = 1 Then
            strFormat = VSP_CalTable(FOOTER_FORMAT_2, mdblWid, alngX)
            strPrint = "�ط��|(" & ThaiBaht(Sum3 - TempDiscount + m_cData.VAT_AMOUNT) & ")|||" & FormatNumber(Sum3 - TempDiscount + m_cData.VAT_AMOUNT) & ";"
      Else
            strFormat = VSP_CalTable("^1|>7|>2;", mdblWid, alngX)
            strPrint = "�ط��|(" & ThaiBaht(Sum3 - TempDiscount + m_cData.VAT_AMOUNT) & ")|" & FormatNumber(Sum3 - TempDiscount + m_cData.VAT_AMOUNT) & ";"
      End If
      Call Vsp.AddTable(strFormat, "", strPrint)
   
      Vsp.FontBold = False
      Call PrintBillFooter(alngX)
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
'   Call Vsp.AddTable(strFormat, "", "|˹�ҷ��|: " & Vsp.CurrentPage)
'   Call Vsp.AddTable(strFormat, "", "|�ѹ�������|: " & DateToStringExt(mdteDate))
'   Call Vsp.AddTable(strFormat, "", "|����|: " & Right$(DateToStringExtEx(mdteDate), 8))
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


