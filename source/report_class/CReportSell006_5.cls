VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportSell006_5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportSell006_5"
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
Private m_FeatureTotalPrices As Collection
Private m_PartItems As Collection
Private m_CustAddr As Collection

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
   Set m_FeatureTotalPrices = New Collection
   Set m_PartItems = New Collection
   Set m_CustAddr = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_FeatureTotalPrices = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
   Set m_PartItems = Nothing
   Set m_CustAddr = Nothing
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
      
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   If mcolParam("FLAG_SHOW_PART_NO") = 1 Then
      Call Cf.SetFieldValue(4, "^", "��¡���Թ���", "<")
      Call m_Headers1.AddField(Cf)
   End If
   Call Cf.SetFieldValue(2, "^", "�Ţ����͡���", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�ѹ����͡���", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�ӹǹ(�ا)", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�Ҥ�/˹���", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "���˹ѡ(��.)", ">")
   Call m_Headers1.AddField(Cf)
   If mcolParam("FLAG_DELIVERY_CUS") = 1 Then
      Call Cf.SetFieldValue(3, "^", "ʶҹ���Ѵ��", "<")
      Call m_Headers1.AddField(Cf)
   End If
   Call Cf.SetFieldValue(3, "^", "�ӹǹ�Թ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "��� 30 �ѹ", "^")
   Call m_Headers1.AddField(Cf)
   
   Set Cf = Nothing
   
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
   Vsp.ORIENTATION = orPortrait
   Vsp.MarginBottom = 1440
   Vsp.MarginFooter = 1440
   Vsp.MarginHeader = 1440
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 1440
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   Call LoadFeatureTotalPriceByBill(Nothing, m_FeatureTotalPrices, mcolParam("FROM_DATE"), mcolParam("TO_DATE"))
   Call LoadCustomerAddress(Nothing, m_CustAddr, , False, 2)
   
'   FromDate = "�ҡ�ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
'   ToDate = "�֧ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")
   
   MODULE_DESC = vbCrLf & glbEnterPrise.ENTERPRISE_NAME & vbCrLf & _
                                       "��ҧ���"
   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))
   
   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
      mstrErrMsg = "Error GetDateTime Error."
      Exit Function
   End If
   mdteDate = InternalDateToDate(strDate)
   initDoc = True
End Function

Private Function GetKey(Di As CDoItem) As String
Dim OrderBy As Long

   OrderBy = mcolParam("ORDER_BY")
   
   If OrderBy = 1 Then
      GetKey = Di.CUSTOMER_ID
   ElseIf OrderBy = 2 Then
      GetKey = Di.CUSTOMER_ID
   Else
      GetKey = Di.CUSTOMER_ID
   End If
End Function

Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim RName As String
Dim I As Long
Dim J As Long
Dim k As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Amt As Double
Dim Di As CDoItem
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(100) As Double
Dim Total2(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim Cm As CCustomer
Dim Pi As CPartItem
Dim Ft As CFeature
Dim Ri As CReceiptItem
Dim PrevCode As String
Dim PrevName As String
Dim IncDiscount As Double
Dim Di2 As CDoItem
Dim PackAmount As Double
Dim WeightAmount As Double
Dim tmpStr1 As String
Dim tmpStr2 As String
Dim m_CustAddress As CAddress

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set Rs = New ADODB.Recordset
'
'   Call LoadPartItem(Nothing, m_PartItems, mcolParam("PART_TYPE"))
   
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------

   For J = 1 To UBound(Total1)
      Total1(J) = 0
      Total2(J) = 0
   Next J
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
'   Call printHeader
'   mblnHeader = True
   
      I = 0
      k = 0
      Set BodyCf = New CReportField

      Set Di = New CDoItem
      Di.DO_ITEM_ID = -1
      Di.CUSTOMER_CODE = mcolParam("CUSTOMER_CODE")
      Di.FROM_CUSTOMER_CODE = mcolParam("FROM_CUSTOMER_CODE")
      Di.TO_CUSTOMER_CODE = mcolParam("TO_CUSTOMER_CODE")
      Di.FROM_DATE = mcolParam("FROM_DATE")
      Di.TO_DATE = mcolParam("TO_DATE")
      Di.CUSTOMER_TYPE = mcolParam("CUSTOMER_TYPE")
      Di.DocTypeSet = BillingDocType2Set(mcolParam("BILL_SUBTYPE"))
      Di.OrderBy = 1
      Di.OrderType = 1
      Call Di.QueryData(50, Rs, iCount)

      I = 0
      While Not Rs.EOF
         Call Di.PopulateFromRS(50, Rs)
          If Di.PART_ITEM_ID > 0 Then
          
           If PrevKey1 <> Di.CUSTOMER_CODE And I > 0 Then
               Call GenerateFooter("", "���������", "", Total1, m_Details1, 0)
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False

               Vsp = ""

               Vsp = "�����˵� : �������� " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A") & "-" & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")
                Vsp.FontBold = True
               Vsp = "��سҪ����Թ���͹��Һѭ��   " & glbEnterPrise.ENTERPRISE_NAME & " �ѭ���Թ�ҡ���������ѹ ��Ҥ��" & mcolParam("BANKS_NAME") & " �Ң� " & mcolParam("BANK_BRANCHS_NAME") & " �Ţ��� " & mcolParam("ACCOUNT_ID_NAME")
               Vsp.FontBold = False
               
               Vsp = ""
               
               Vsp.TableBorder = tbNone
               strFormat = VSP_CalTable("<15|<15;", mdblWid, alngX)
                TempStr1 = " ����Ѻ�ҧ���...........................................................................|����ҧ���...........................................................................;"
                Call Vsp.AddTable(strFormat, "", TempStr1)
                TempStr1 = " �ѹ���......................................................................................|�ѹ��� " & DateToStringExt(mcolParam("PRINT_DATE")) & " ;"
                Call Vsp.AddTable(strFormat, "", TempStr1)
                Vsp.TableBorder = tbAll
               
               Vsp.NewPage
               I = 0
               
               Vsp.TableBorder = tbNone
               strFormat = VSP_CalTable("<15|>15;", mdblWid, alngX)
                TempStr1 = "|�ѹ��� " & DateToStringExt(mcolParam("PRINT_DATE")) & " ;"
                Call Vsp.AddTable(strFormat, "", TempStr1)
                Vsp.TableBorder = tbAll
              
               Vsp = "�����١��� : " & Di.CUSTOMER_CODE
               Vsp = Di.CUSTOMER_NAME

               Set m_CustAddress = GetObject("CAddress", m_CustAddr, Trim(str(Di.BILLING_ADDRESS_ID)), False)
               If Not m_CustAddress Is Nothing Then
                   Vsp = "��������١��� : " & m_CustAddress.PackAddress
               Else
'                  Vsp = ""
               End If

               Call printHeader
               
               For J = 1 To UBound(Total1)
                  Total1(J) = 0
               Next J
            ElseIf I = 0 Then
               Vsp.TableBorder = tbNone
               strFormat = VSP_CalTable("<15|>15;", mdblWid, alngX)
                TempStr1 = "|�ѹ��� " & DateToStringExt(mcolParam("PRINT_DATE")) & " ;"
                Call Vsp.AddTable(strFormat, "", TempStr1)
                Vsp.TableBorder = tbAll
                
               Vsp = "�����١��� : " & Di.CUSTOMER_CODE
               Vsp = Di.CUSTOMER_NAME
               
               Set m_CustAddress = GetObject("CAddress", m_CustAddr, Trim(str(Di.BILLING_ADDRESS_ID)), False)
               If Not m_CustAddress Is Nothing Then
                   Vsp = "��������١��� : " & m_CustAddress.PackAddress
               Else
'                  Vsp = ""
               End If
               
               Call printHeader
            End If
            PrevKey1 = Di.CUSTOMER_CODE
            
         Call m_Details1.ClearField
         Set Di2 = GetDoItem(m_FeatureTotalPrices, Trim(str(Di.DO_ID)))
         
         k = 1
         TempStr = ""
         Set HeadCf = m_Headers1.Fields(k)
         If PrevKey2 <> Di.DOCUMENT_NO Then
            I = I + 1
            TempStr = I
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         If mcolParam("FLAG_SHOW_PART_NO") = 1 Then
            k = k + 1
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = Di.BARCODE_NO & " " & Di.BILL_DESC
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
         End If
         
         tmpStr1 = ""
         tmpStr2 = ""
         If PrevKey2 <> Di.DOCUMENT_NO Then
             tmpStr1 = Di.DOCUMENT_NO
             tmpStr2 = DateToStringExtEx2(Di.DOCUMENT_DATE)
         End If
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = tmpStr1
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = tmpStr2
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(Di.PACK_AMOUNT, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + (Di.PACK_AMOUNT)

         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(Di.PRICE_PER_PACK, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(Di.ITEM_AMOUNT, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + (Di.ITEM_AMOUNT)
         
         If mcolParam("FLAG_DELIVERY_CUS") = 1 Then
               k = k + 1
               Set HeadCf = m_Headers1.Fields(k)
               TempStr = ""
               If PrevKey2 <> Di.DOCUMENT_NO Then
                  If Len(Di.SHIPPING) > 0 Then
                     TempStr = Trim(Replace(Di.SHIPPING, vbNewLine, ""))
                  Else
                     TempStr = Trim(Replace(Di.PAYMENT_DESC, vbNewLine, ""))
                  End If
               End If
               
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
         
         End If
         PrevKey2 = Di.DOCUMENT_NO
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(Di.TOTAL_PRICE, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + (Di.TOTAL_PRICE)
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = DateToStringExtEx2(Di.DUE_DATE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)

          End If
         Rs.MoveNext
      Wend
   Call GenerateFooter("", "���������", "", Total1, m_Details1, 0)
   Call m_Details1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False
   
   Vsp = ""
   
   Vsp = "�����˵� : �������� " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A") & "-" & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")
   Vsp.FontBold = True
   Vsp = "��سҪ����Թ���͹��Һѭ��   " & glbEnterPrise.ENTERPRISE_NAME & " �ѭ���Թ�ҡ���������ѹ ��Ҥ��" & mcolParam("BANKS_NAME") & " �Ң� " & mcolParam("BANK_BRANCHS_NAME") & " �Ţ��� " & mcolParam("ACCOUNT_ID_NAME")
   Vsp.FontBold = False
   
   Vsp = ""
   
   Vsp.TableBorder = tbNone
   strFormat = VSP_CalTable("<15|<15;", mdblWid, alngX)
   TempStr1 = " ����Ѻ�ҧ���...........................................................................|����ҧ���...........................................................................;"
   Call Vsp.AddTable(strFormat, "", TempStr1)
   TempStr1 = " �ѹ���......................................................................................|�ѹ��� " & DateToStringExt(mcolParam("PRINT_DATE")) & " ;"
   Call Vsp.AddTable(strFormat, "", TempStr1)
   Vsp.TableBorder = tbAll

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

Private Sub GenerateFooter(TXT1 As String, TXT2 As String, TXT3 As String, Tot() As Double, Details As CFieldList, DateCount As Long)
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
         TempStr = TXT2
      ElseIf J = 3 Then
         TempStr = TXT3
      ElseIf J = 4 Or J = 6 Or J = 8 Or J = 10 Then
         TempStr = ""
      Else
         TempStr = FormatNumber(Tot(J))
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.Align)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub

Private Sub GenerateHeader(TXT1 As String, TXT2 As String, TXT3 As String, Details As CFieldList)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim J As Long
Dim TempStr As String

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   J = 0
   For Each HeadCf In m_Headers0.Fields
      J = J + 1
      If J = 1 Then
         TempStr = TXT1
      ElseIf J = 2 Then
         TempStr = TXT2
      ElseIf J = 3 Then
         TempStr = TXT3
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
   
'   Vsp.FontSize = 8
'   dY(0) = Vsp.CurrentY
'   Vsp.TextAlign = taLeftBottom
'   strFormat = VSP_CalTable(TITLE_FORMAT, mdblWid, alngX)
'   Vsp.TableBorder = tbNone
'   Call Vsp.AddTable(strFormat, "", "|˹�ҷ��|: " & Vsp.CurrentPage)
'   Call Vsp.AddTable(strFormat, "", "|�ѹ�������|: " & DateToStringExt(mdteDate))
'   Call Vsp.AddTable(strFormat, "", "|����|: " & Right$(DateToStringExtEx(mdteDate), 8))
'   Call Vsp.AddTable(strFormat, "", "|�����|: " & glbUser.USER_NAME)
'   Vsp.CurrentY = dY(0)
   
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



