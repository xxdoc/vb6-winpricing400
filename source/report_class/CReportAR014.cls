VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAR014"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAR014"
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
Private m_YearSeqs As Collection

Private m_CustTotalPrices As Collection
Private m_CustTotalPrices2 As Collection
Private m_CashTrans As Collection
Private m_CustomerPaids As Collection

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
   Set m_YearSeqs = New Collection
   Set m_CustTotalPrices = New Collection
   Set m_CashTrans = New Collection
   Set m_CustomerPaids = New Collection
   Set m_CustTotalPrices2 = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_YearSeqs = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
   Set m_CustTotalPrices = Nothing
   Set m_CashTrans = Nothing
   Set m_CustomerPaids = Nothing
   Set m_CustTotalPrices2 = Nothing
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
Dim PrevID As Long
Dim TempStr1 As String
Dim TempStr1_1 As String
Dim Cf As CReportField
Dim ExportWidth As Long
Dim HeadCf As CReportField
Dim J As Long
Dim TempStr As String
Dim TempRs As ADODB.Recordset
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

'   Call Cf.SetFieldValue(4, "^", "�����١���", "^")
'   Call m_Headers0.AddField(Cf)
'   Call Cf.SetFieldValue(6, "^", "�����١���", "<")
'   Call m_Headers0.AddField(Cf)
'   Call Cf.SetFieldValue(7, "^", "��ѡ�ҹ���", "<")
'   Call m_Headers0.AddField(Cf)

   '===
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�����١���", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(5, "^", "�����١���", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�ӹǹ���", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "��Ť�Ң������", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�Ѻ����", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(0, "^", "�Ѻ�Թ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(0, "^", "��ǹ��ҧ�Ѻ����", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(0, "^", "˹�餧�����", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�ŵ�ҧ��������Ѻ����", ">")
   Call m_Headers1.AddField(Cf)
   
   Set Cf = Nothing

'  Call m_Headers0.GetString(1, TempStr1, TempStr1_1)
'   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'   strPrint = TempStr1_1
'   Call Vsp.AddTable(strFormat, "", strPrint)

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
Dim ToPaidDate As String
Dim NewDate As Date
Dim NewFromDate As Date
Dim NewToDate As Date

   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
   Vsp.ORIENTATION = orLandscape
   Vsp.MarginBottom = 700
   Vsp.MarginFooter = 700
   Vsp.MarginHeader = 700
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 700
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   Call LoadTotalPriceByCustomer(Nothing, m_CustTotalPrices, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), 1, 8)
   Call LoadTotalPriceByCustomer(Nothing, m_CustTotalPrices2, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), 1, 18)
   Call LoadPaidAmountByCustomer(Nothing, m_CustomerPaids, mcolParam("FROM_DATE"), mcolParam("TO_DATE"))
   Call LoadCashTranAmountByCust(Nothing, m_CashTrans, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), 2, 3)
   
   FromDate = "�ҡ�ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
   ToDate = " �֧ �ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")
   
   MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & vbCrLf & mcolParam("REPORT_NAME") & _
                                       vbCrLf & FromDate & ToDate & "  " & ToPaidDate
   
   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))
   
   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
      mstrErrMsg = "Error GetDateTime Error."
      Exit Function
   End If
   mdteDate = InternalDateToDate(strDate)
   initDoc = True
End Function

Private Function GetItemFromListIndex(TempCol As Collection, Ind As Long) As Object
   Set GetItemFromListIndex = TempCol(Ind)
End Function

Private Function GetKey(Mode As Long, Di As CReceiptItem) As String
   If Mode = 1 Then
      GetKey = Di.RECEIPT_NO
   ElseIf Mode = 2 Then
      GetKey = Di.CUSTOMER_CODE
   End If
End Function

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

Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim RName As String
Dim I As Long
Dim J As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Amt As Double
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(100) As Double
Dim Total2(100) As Double
Dim Total3(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim Width As Double
Dim TempRs As ADODB.Recordset
Dim Hc As CReportField
Dim Bd As CBillingDoc
Dim Di As CDoItem
Dim Di2 As CDoItem
Dim Ri As CReceiptItem
Dim Ct As CCashTran
Dim NetPrice As Double

'   m_ApSupAmounts
'   m_ApSupPaids
   
   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set Rs = New ADODB.Recordset
   Set TempRs = New ADODB.Recordset
   
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
   
   For J = 1 To UBound(Total1)
      Total1(J) = 0
      Total2(J) = 0
      Total3(J) = 0
   Next J
   
     mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
   
       Set BodyCf = New CReportField
      Set Bd = New CBillingDoc
      Bd.BILLING_DOC_ID = -1
      Bd.FROM_DATE = mcolParam("FROM_DATE")
      Bd.TO_DATE = mcolParam("TO_DATE")
      Bd.CUSTOMER_CODE = mcolParam("CUSTOMER_CODE")
      Bd.FROM_CUSTOMER_CODE = mcolParam("FROM_CUSTOMER_CODE")
      Bd.TO_CUSTOMER_CODE = mcolParam("TO_CUSTOMER_CODE")
      Bd.CUSTOMER_TYPE = mcolParam("CUSTOMER_TYPE")
      Bd.CUSTOMER_GRADE = mcolParam("CUSTOMER_GRADE")
      Bd.OrderBy = mcolParam("ORDER_BY")
      Bd.OrderType = 1
      Call Bd.QueryData(5, Rs, iCount)

'      If Not Rs.EOF Then
'         Call BD.PopulateFromRS(5, Rs)
'         PrevKey1 = GetKey(2, Ri)
'         PrevKey2 = GetKey(1, Ri)
'
'         Call GenerateHeader(Ri.CUSTOMER_CODE, Ri.CUSTOMER_NAME, Ri.SALE_NAME, m_Details1)
'         Call m_Details1.GetString(1, TempStr1, TempStr2)
'         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'         Vsp.FontBold = True
'         Call Vsp.AddTable(strFormat, "", TempStr2)
'         Vsp.FontBold = False
'      End If
      
      While Not Rs.EOF
         Call Bd.PopulateFromRS(5, Rs)
         
         Set Di = GetDoItem(m_CustTotalPrices, Trim(str(Bd.CUSTOMER_ID)))
         Set Di2 = GetDoItem(m_CustTotalPrices2, Trim(str(Bd.CUSTOMER_ID)))
         Set Ri = GetReceiptItem(m_CustomerPaids, Trim(str(Bd.CUSTOMER_ID)))
         Set Ct = GetCashTran(m_CashTrans, Trim(str(Bd.CUSTOMER_ID)))
         NetPrice = Di.TOTAL_PRICE - Di.EXTRA_DISCOUNT '��ǹŴ�������� TOTAL_PRICE ����

         I = I + 1
         Call m_Details1.ClearField
         
          Set Hc = m_Headers1.Fields(1)
         TempStr = I
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

          Set Hc = m_Headers1.Fields(2)
         TempStr = Bd.CUSTOMER_CODE
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

          Set Hc = m_Headers1.Fields(3)
         TempStr = Bd.CUSTOMER_NAME
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

          Set Hc = m_Headers1.Fields(4)
         TempStr = FormatNumber(Di2.ITEM_AMOUNT)
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(4) = Total1(4) + (Di2.ITEM_AMOUNT)
         Total2(4) = Total2(4) + (Di2.ITEM_AMOUNT)
         
          Set Hc = m_Headers1.Fields(5)
         TempStr = FormatNumber(NetPrice)
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(5) = Total1(5) + (NetPrice)
         Total2(5) = Total2(5) + (NetPrice)
         
          Set Hc = m_Headers1.Fields(6)
         TempStr = FormatNumber(Ri.PAID_AMOUNT)
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(6) = Total1(6) + (Ri.PAID_AMOUNT)
         Total2(6) = Total2(6) + (Ri.PAID_AMOUNT)

          Set Hc = m_Headers1.Fields(7)
         TempStr = FormatNumber(Ct.GetFieldValue("AMOUNT"))
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(7) = Total1(7) + (Ct.GetFieldValue("AMOUNT"))
         Total2(7) = Total2(7) + (Ct.GetFieldValue("AMOUNT"))

          Set Hc = m_Headers1.Fields(8)
         TempStr = FormatNumber(Ri.PAID_AMOUNT - Ct.GetFieldValue("AMOUNT"))
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(8) = Total1(8) + (Ri.PAID_AMOUNT - Ct.GetFieldValue("AMOUNT"))
         Total2(8) = Total2(8) + (Ri.PAID_AMOUNT - Ct.GetFieldValue("AMOUNT"))

          Set Hc = m_Headers1.Fields(9)
         TempStr = FormatNumber(0)
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(9) = Total1(9) + (0)
         Total2(9) = Total2(9) + (0)
         
         Set Hc = m_Headers1.Fields(10)
         TempStr = FormatNumber(Ri.PAID_AMOUNT - NetPrice)
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(10) = Total1(10) + (Ri.PAID_AMOUNT - NetPrice)
         Total2(10) = Total2(10) + (Ri.PAID_AMOUNT - NetPrice)
         
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)

         Rs.MoveNext
      Wend

'      Call GenerateFooter("���", "�����", "", Total3, m_Details1, 0)
'      Call m_Details1.GetString(2, TempStr1, TempStr2)
'      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'      Vsp.FontBold = True
'      Call Vsp.AddTable(strFormat, "", TempStr2)
'      Vsp.FontBold = False
      
      Call GenerateFooter("���", "", "", Total1, m_Details1, 0)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False

'      Call GenerateFooter("�ط��", "", "", Total2, m_Details1, 0)
'      Call m_Details1.GetString(2, TempStr1, TempStr2)
'      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'      Vsp.FontBold = True
'      Call Vsp.AddTable(strFormat, "", TempStr2)
'      Vsp.FontBold = False

      Set BodyCf = Nothing
    Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
   
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
         TempStr = ""
      Else
         TempStr = FormatNumber(Tot(J))
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub

Private Sub GenerateBkTr(TXT1 As String, TXT2 As String, TXT3 As String, Tot() As Double, Details As CFieldList, DateCount As Long)
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
   Call Vsp.AddTable(strFormat, "", "|˹�ҷ��|: " & Vsp.CurrentPage)
   Call Vsp.AddTable(strFormat, "", "|�ѹ|: " & DateToStringExt(mdteDate))
   Call Vsp.AddTable(strFormat, "", "|����|: " & Right$(DateToStringExtEx(Now), 8))
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




