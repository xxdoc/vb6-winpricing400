VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAP017"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAP017"
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
Private m_Headers2 As CFieldList
Private m_Details1 As CFieldList
Private m_Details2 As CFieldList
Private m_Left As Double
Private m_YearSeqs As Collection

Private m_ExpenseTypes As Collection
Private m_PartGroup As Collection
Private collSumSupItemByRo  As Collection
Dim TmpFlag As String

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_TempCol = New Collection

   Set m_ExpenseTypes = New Collection
   Set m_PartGroup = New Collection

   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Headers2 = New CFieldList
   Set m_Details1 = New CFieldList
   Set m_Details2 = New CFieldList
   Set m_YearSeqs = New Collection
   Set collSumSupItemByRo = New Collection
End Sub
Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing

   Set m_ExpenseTypes = Nothing
   Set m_PartGroup = Nothing

   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Headers2 = Nothing
   Set m_Details1 = Nothing
   Set m_YearSeqs = Nothing
   Set m_Details2 = Nothing
   Set collSumSupItemByRo = Nothing
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

   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize

   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   

   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   Call m_Headers2.ClearField

   Set Cf = New CReportField
   
   Call Cf.SetFieldValue(20, "^", "�ѹ���", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(25, "^", "P/O", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(50, "^", "SUPPLIER", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(50, "^", "R/M", "<")
   Call m_Headers1.AddField(Cf)
     Call Cf.SetFieldValue(20, "^", "@", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(20, "^", "�ӹǹ", ">", "SUM_FLAG") '�ӹǹ PO
   Call m_Headers1.AddField(Cf)
   If mcolParam("SHOW_UNIT_NAME_FLAG") = 1 Then
      Call Cf.SetFieldValue(20, "^", "˹���", ">")
   Else
      Call Cf.SetFieldValue(0, "^", "˹���", ">")
   End If
    Call m_Headers1.AddField(Cf)

   Call Cf.SetFieldValue(20, "^", "¡�ʹ��", ">", "SUM_FLAG")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(20, "^", "��.�ط��", ">", "SUM_FLAG") '�ӹǹ�Ѻ
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(20, "^", "��.���͢��", ">", "SUM_FLAG") '�ӹǹ�Ѻ
   Call m_Headers1.AddField(Cf)
    Call Cf.SetFieldValue(20, "^", "��.��Ҥ�ѧ", ">", "SUM_FLAG") '�ӹǹ�Ѻ�ط��
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(20, "^", "�������", ">", "SUM_FLAG")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(20, "^", "��ҧ��", ">", "SUM_FLAG") '�������
   Call m_Headers1.AddField(Cf)
   
    Call Cf.SetFieldValue(20, "^", "�ӹǹ�Թ", ">", "SUM_FLAG_DEC")
   Call m_Headers1.AddField(Cf)
    Call Cf.SetFieldValue(20, "^", "�����˵�", "<")
   Call m_Headers1.AddField(Cf)
   
   If mcolParam("SHOW_STATUS_PO_FLAG") = 1 Then
      Call Cf.SetFieldValue(20, "^", "ʶҹ�", ">")
   Else
      Call Cf.SetFieldValue(0, "^", "ʶҹ�", ">")
   End If
   Call m_Headers1.AddField(Cf)

   Set Cf = Nothing
   
  Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
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
Dim ComeInDate As String

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
   
    ComeInDate = "�ѹ����Ѻ�ͧ   " & EmptyToString(DateToStringExtEx2(mcolParam("PRINT_DATE")), "N/A")

   MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & _
                                       vbCrLf & mcolParam("REPORT_NAME") & _
                                       vbCrLf & ComeInDate
                                       '& _vbCrLf & "���§���" & mcolParam("OrderType") & "," & "���§�ҡ" & mcolParam("ORDER_BY")
                                       
                                       '& FromDate & ToDate & " , "OrderType
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
Dim O As Object
Dim RName As String
Dim I As Long
Dim J As Long
Dim k As Long
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
Dim iCount As Long
Dim TempStr As String
Dim TempID As Long

Dim Sup As CSupItem
Dim PrevKey1 As String

Dim TempRo As CSupItem
Dim TempComeIn  As CSupItem
Dim RoColl As Collection
Dim ComeInColl As Collection  'Query �ӹǹ�ѵ�شԺ�������ҡ�͹˹���ѹ����ͧ����͡ report 1 �ѹ  ���ͷ��ʹ¡�� ��� �ʹ����
Dim CarryForward As Double

Dim Sum1 As Double
Dim Sum2 As Double
Dim A As Double '�ӹǹ�ѵ�ط�������
Dim B As Double '¡�ʹ��
Dim C As Double  '���
Dim D As Double ' �������
Dim E As Double '��ҧ��
Dim PrintDate As Date
Dim TMP_WEIGHT As Double
'Dim Mark1 As String
'Dim Mark2 As String
'Dim Mark3 As String

Dim tempSupItem As CSupItem

Dim TempLimit As Long
'
   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------

   Set RoColl = New Collection
    Set ComeInColl = New Collection
    

    
   If mcolParam("PRINT_DATE") <= 0 Then
      PrintDate = Now
   Else
      PrintDate = mcolParam("PRINT_DATE")
   End If
   
  'RoColl : ������ �ѹ�Ѻ��ҷ���ͧ���
  'ComeInColl : �������ѹ�Ѻ��� ��͹�ѹ����ͧ��� 1 �ѹ  ���� ���ʹ¡��  ���� �ʹ���� ��͹˹�ҹ��
  Call LoadSupItemComeIn(Nothing, RoColl, PrintDate, PrintDate, mcolParam("SUPPLIER_CODE"))
  Call LoadSupItemComeIn(Nothing, ComeInColl, -1, DateAdd("d", -1, PrintDate), mcolParam("SUPPLIER_CODE"))
   
   Call LoadSupItemPartItemByRo(Nothing, collSumSupItemByRo, -1, DateAdd("D", -1, PrintDate), mcolParam("SUPPLIER_CODE"), "(100,101,102,103)")
   Set Rs = New ADODB.Recordset

   If Not initDoc Then Exit Function

   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
   
   For J = 1 To UBound(Total1)
      Total1(J) = 0
      Total2(J) = 0
   Next J
   
   Set Sup = New CSupItem
   
   Sup.SUP_ITEM_ID = -1
   Sup.SUPPLIER_CODE = mcolParam("SUPPLIER_CODE")
   Sup.SUPPLIER_GRADE = mcolParam("SUPPLIER_GRADE")
   Sup.SUPPLIER_TYPE = mcolParam("SUPPLIER_TYPE")
   Sup.OrderBy = mcolParam("ORDER_BY")
   Sup.OrderType = mcolParam("ORDER_TYPE")
   If mcolParam("PO_TYPE") > 0 Then
      Sup.DOCUMENT_TYPE = mcolParam("PO_TYPE")
    Else
      Sup.DOCUMENT_TYPE_SET = "(1000,1001,1002,1003)"
   End If
   Sup.CLOSE_FLAG = CFLAG(mcolParam("PO_CLOSE"))
     
   Sup.PO_APPROVED_FLAG = "Y"
   Sup.TO_DATE = PrintDate
   Sup.PART_NO = mcolParam("PART_NO")
   Call Sup.QueryData(112, Rs, iCount)
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
   
   I = 0
   k = 0
   Set BodyCf = New CReportField
   
   While Not Rs.EOF
      Call m_Details1.ClearField
      Call Sup.PopulateFromRS(112, Rs)
      
      If Sup.CLOSE_FLAG = "" Or Sup.CLOSE_FLAG = "Y" Then
          I = I + 1
      
         Call m_Details1.ClearField
         
         PrevKey1 = Sup.PART_NO
         
         Set HeadCf = m_Headers1.Fields(1)
         TempStr = DateToStringExtEx2(Sup.DOCUMENT_DATE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(2)
         TempStr = Sup.DOCUMENT_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         TempLimit = 25
         If Val(mcolParam("LIMIT_SUPPLIER_NAME")) > 0 Then
            TempLimit = Val(mcolParam("LIMIT_SUPPLIER_NAME"))
         End If
         Set HeadCf = m_Headers1.Fields(3)
         TempStr = Left(Sup.SUPPLIER_NAME, TempLimit)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         TempLimit = 25
         If Val(mcolParam("LIMIT_PART_DESC")) > 0 Then
            TempLimit = Val(mcolParam("LIMIT_PART_DESC"))
         End If
         Set HeadCf = m_Headers1.Fields(4)
         TempStr = Left(Sup.PART_DESC, TempLimit)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
               
         Set HeadCf = m_Headers1.Fields(5)
'         If Sup.SUPPLIER_CODE = "�-0012" Or Sup.SUPPLIER_CODE = "�-1051" Then
'            TempStr = FormatNumberToNull(Sup.ACTUAL_UNIT_PRICE, 3)
'         Else
'           TempStr = FormatNumberToNull(Sup.ACTUAL_UNIT_PRICE, 2)
'         End If
         TempStr = FormatNumberToNull(Sup.ACTUAL_UNIT_PRICE, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(6)
         TempStr = FormatNumberToNull(Sup.TX_AMOUNT, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(6) = Total1(6) + Sup.TX_AMOUNT
         
         Set HeadCf = m_Headers1.Fields(7)
         TempStr = Sup.UNIT_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         Set TempRo = GetObject("CSupItem", RoColl, Trim(Sup.SUPPLIER_ID & "-" & Sup.BILLING_DOC_ID & "-" & Sup.PART_ITEM_ID))
         Set TempComeIn = GetObject("CSupItem", ComeInColl, Trim(Sup.SUPPLIER_ID & "-" & Sup.BILLING_DOC_ID & "-" & Sup.PART_ITEM_ID))
         
          If Sup.CALCULATE_TYPE = 0 Then
            Set HeadCf = m_Headers1.Fields(8) '���ҧ column ��ҧ
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, "", HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
         End If
         
          If Sup.CALCULATE_TYPE = 1 Then
               Set HeadCf = m_Headers1.Fields(8) '���˹ѡ���͢�� > �����
               TempStr = FormatNumberToNull(TempComeIn.SUPPLIER_WEIGHT, 2)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
               Call m_Details1.AddField(BodyCf)
               Total1(8) = Total1(8) + TempRo.SUPPLIER_WEIGHT
         ElseIf Sup.CALCULATE_TYPE = 2 Then
               Set HeadCf = m_Headers1.Fields(8) '���˹ѡ���͢�� > ���
               TempStr = FormatNumberToNull(TempComeIn.WEIGHT_AMOUNT, 2)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
               Call m_Details1.AddField(BodyCf)
               Total1(8) = Total1(8) + TempRo.WEIGHT_AMOUNT
         ElseIf Sup.CALCULATE_TYPE = 3 Then
               Set HeadCf = m_Headers1.Fields(8) '���˹ѡ���͢�� > �ط��
               TempStr = FormatNumberToNull(TempComeIn.TOTAL_WEIGHT, 2)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
               Call m_Details1.AddField(BodyCf)
               Total1(8) = Total1(8) + TempRo.TOTAL_WEIGHT
         End If
         
         Set HeadCf = m_Headers1.Fields(9) '���˹ѡ�ط��
         TempStr = FormatNumberToNull(TempRo.TOTAL_WEIGHT, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(9) = Total1(9) + TempRo.TOTAL_WEIGHT
         
         If Sup.CALCULATE_TYPE = 1 Then
               Set HeadCf = m_Headers1.Fields(10) '���˹ѡ���͢�� > �����
               TempStr = FormatNumberToNull(TempRo.SUPPLIER_WEIGHT, 2)
               TMP_WEIGHT = TempRo.SUPPLIER_WEIGHT
               Total1(10) = Total1(10) + TempRo.SUPPLIER_WEIGHT
         ElseIf Sup.CALCULATE_TYPE = 2 Then
               Set HeadCf = m_Headers1.Fields(10) '���˹ѡ���͢�� > ���
               TempStr = FormatNumberToNull(TempRo.WEIGHT_AMOUNT, 2)
               TMP_WEIGHT = TempRo.WEIGHT_AMOUNT
               Total1(10) = Total1(10) + TempRo.WEIGHT_AMOUNT
         ElseIf Sup.CALCULATE_TYPE = 3 Then
               Set HeadCf = m_Headers1.Fields(10) '���˹ѡ���͢�� > �ط��
               TempStr = FormatNumberToNull(TempRo.TOTAL_WEIGHT, 2)
               TMP_WEIGHT = TempRo.TOTAL_WEIGHT
               Total1(10) = Total1(10) + TempRo.TOTAL_WEIGHT
         End If
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
          Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(11) '���˹ѡ��Ҥ�ѧ
         TempStr = FormatNumberToNull(TempRo.TX_AMOUNT, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(11) = Total1(11) + TempRo.TX_AMOUNT
         
         Set HeadCf = m_Headers1.Fields(12) '�������
         If Sup.CALCULATE_TYPE = 1 Then
               Set HeadCf = m_Headers1.Fields(12) '�������
               TempStr = FormatNumberToNull(TempComeIn.SUPPLIER_WEIGHT + TempRo.SUPPLIER_WEIGHT, 2)
               Total1(12) = Total1(12) + TempComeIn.TX_AMOUNT + TempRo.SUPPLIER_WEIGHT
         ElseIf Sup.CALCULATE_TYPE = 2 Then
                Set HeadCf = m_Headers1.Fields(12) '�������
               TempStr = FormatNumberToNull(TempComeIn.WEIGHT_AMOUNT + TempRo.WEIGHT_AMOUNT, 2)
               Total1(12) = Total1(12) + TempComeIn.TX_AMOUNT + TempRo.WEIGHT_AMOUNT
         ElseIf Sup.CALCULATE_TYPE = 3 Then
                Set HeadCf = m_Headers1.Fields(12) '�������
               TempStr = FormatNumberToNull(TempComeIn.TOTAL_WEIGHT + TempRo.TOTAL_WEIGHT, 2)
               Total1(12) = Total1(12) + TempComeIn.TX_AMOUNT + TempRo.TOTAL_WEIGHT
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(13) ' ��ҧ��
         If Sup.CALCULATE_TYPE = 1 Then
               Set HeadCf = m_Headers1.Fields(13) '���˹ѡ���͢�� > �����
               TempStr = FormatNumberToNull(Sup.TX_AMOUNT - (TempComeIn.SUPPLIER_WEIGHT + TempRo.SUPPLIER_WEIGHT))
               Total1(13) = Total1(13) + (Sup.TX_AMOUNT - (TempComeIn.SUPPLIER_WEIGHT + TempRo.SUPPLIER_WEIGHT))
         ElseIf Sup.CALCULATE_TYPE = 2 Then
               Set HeadCf = m_Headers1.Fields(13) '���˹ѡ���͢�� > ���
               TempStr = FormatNumberToNull(Sup.TX_AMOUNT - (TempComeIn.WEIGHT_AMOUNT + TempRo.WEIGHT_AMOUNT))
               Total1(13) = Total1(13) + (Sup.TX_AMOUNT - (TempComeIn.WEIGHT_AMOUNT + TempRo.WEIGHT_AMOUNT))
         ElseIf Sup.CALCULATE_TYPE = 3 Then
               Set HeadCf = m_Headers1.Fields(13) '���˹ѡ���͢�� > �ط��
               TempStr = FormatNumberToNull(Sup.TX_AMOUNT - (TempComeIn.TOTAL_WEIGHT + TempRo.TOTAL_WEIGHT))
               Total1(13) = Total1(13) + (Sup.TX_AMOUNT - (TempComeIn.TOTAL_WEIGHT + TempRo.TOTAL_WEIGHT))
         End If
         
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(14) ' �ӹǹ�Թ  ���ͧ���
         TempStr = FormatNumberToNull(TempRo.TOTAL_INCLUDE_PRICE)
         Total1(14) = Total1(14) + (TempRo.TOTAL_INCLUDE_PRICE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         
         Set HeadCf = m_Headers1.Fields(15)
         TempStr = Sup.NOTE
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
   
         Set HeadCf = m_Headers1.Fields(16)
         TempStr = IIf(Sup.CLOSE_FLAG = "N", "�ѧ���Դ", "�Դ����")
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
      ElseIf Sup.CLOSE_FLAG = "N" Then
         Set tempSupItem = GetObject("CSupItem", collSumSupItemByRo, Trim(Sup.BILLING_DOC_ID & "-" & Sup.PART_ITEM_ID)) '�ʹ�Ѻ�Թ���
         If (Sup.TX_AMOUNT > tempSupItem.TX_AMOUNT) Then
            I = I + 1
         
            Call m_Details1.ClearField
            
            PrevKey1 = Sup.PART_NO
            
            Set HeadCf = m_Headers1.Fields(1)
            TempStr = DateToStringExtEx2(Sup.DOCUMENT_DATE)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(2)
            TempStr = Sup.DOCUMENT_NO
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            TempLimit = 25
            If Val(mcolParam("LIMIT_SUPPLIER_NAME")) > 0 Then
               TempLimit = Val(mcolParam("LIMIT_SUPPLIER_NAME"))
            End If
            Set HeadCf = m_Headers1.Fields(3)
            TempStr = Left(Sup.SUPPLIER_NAME, TempLimit)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            TempLimit = 25
            If Val(mcolParam("LIMIT_PART_DESC")) > 0 Then
               TempLimit = Val(mcolParam("LIMIT_PART_DESC"))
            End If
            Set HeadCf = m_Headers1.Fields(4)
            TempStr = Left(Sup.PART_DESC, TempLimit)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
                  
            Set HeadCf = m_Headers1.Fields(5)
'            If Sup.SUPPLIER_CODE = "�-0012" Or Sup.SUPPLIER_CODE = "�-1051" Then
'               TempStr = FormatNumberToNull(Sup.ACTUAL_UNIT_PRICE, 3)
'            Else
'               TempStr = FormatNumberToNull(Sup.ACTUAL_UNIT_PRICE, 2)
'            End If
            TempStr = FormatNumberToNull(Sup.ACTUAL_UNIT_PRICE, 2)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
                 
'            Set HeadCf = m_Headers1.Fields(5)
'            TempStr = FormatNumberToNull(Sup.ACTUAL_UNIT_PRICE, 2)
'            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
'            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(6)
            TempStr = FormatNumberToNull(Sup.TX_AMOUNT, 2)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            Total1(6) = Total1(6) + Sup.TX_AMOUNT
            
            Set HeadCf = m_Headers1.Fields(7)
            TempStr = Sup.UNIT_NAME
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            Set TempRo = GetObject("CSupItem", RoColl, Trim(Sup.SUPPLIER_ID & "-" & Sup.BILLING_DOC_ID & "-" & Sup.PART_ITEM_ID))
            Set TempComeIn = GetObject("CSupItem", ComeInColl, Trim(Sup.SUPPLIER_ID & "-" & Sup.BILLING_DOC_ID & "-" & Sup.PART_ITEM_ID))
            
             If Sup.CALCULATE_TYPE = 0 Then
               Set HeadCf = m_Headers1.Fields(8) '���ҧ column ��ҧ
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, "", HeadCf.BodyAlign)
               Call m_Details1.AddField(BodyCf)
            End If
            
             If Sup.CALCULATE_TYPE = 1 Then
                  Set HeadCf = m_Headers1.Fields(8) '���˹ѡ���͢�� > �����
                  TempStr = FormatNumberToNull(TempComeIn.SUPPLIER_WEIGHT, 2)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
                  Call m_Details1.AddField(BodyCf)
                  Total1(8) = Total1(8) + TempRo.SUPPLIER_WEIGHT
            ElseIf Sup.CALCULATE_TYPE = 2 Then
                  Set HeadCf = m_Headers1.Fields(8) '���˹ѡ���͢�� > ���
                  TempStr = FormatNumberToNull(TempComeIn.WEIGHT_AMOUNT, 2)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
                  Call m_Details1.AddField(BodyCf)
                  Total1(8) = Total1(8) + TempRo.WEIGHT_AMOUNT
            ElseIf Sup.CALCULATE_TYPE = 3 Then
                  Set HeadCf = m_Headers1.Fields(8) '���˹ѡ���͢�� > �ط��
                  TempStr = FormatNumberToNull(TempComeIn.TOTAL_WEIGHT, 2)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
                  Call m_Details1.AddField(BodyCf)
                  Total1(8) = Total1(8) + TempRo.TOTAL_WEIGHT
            End If
            
            Set HeadCf = m_Headers1.Fields(9) '���˹ѡ�ط��
            TempStr = FormatNumberToNull(TempRo.TOTAL_WEIGHT, 2)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            Total1(9) = Total1(9) + TempRo.TOTAL_WEIGHT
            
            If Sup.CALCULATE_TYPE = 1 Then
                  Set HeadCf = m_Headers1.Fields(10) '���˹ѡ���͢�� > �����
                  TempStr = FormatNumberToNull(TempRo.SUPPLIER_WEIGHT, 2)
                  TMP_WEIGHT = TempRo.SUPPLIER_WEIGHT
                  Total1(10) = Total1(10) + TempRo.SUPPLIER_WEIGHT
            ElseIf Sup.CALCULATE_TYPE = 2 Then
                  Set HeadCf = m_Headers1.Fields(10) '���˹ѡ���͢�� > ���
                  TempStr = FormatNumberToNull(TempRo.WEIGHT_AMOUNT, 2)
                  TMP_WEIGHT = TempRo.WEIGHT_AMOUNT
                  Total1(10) = Total1(10) + TempRo.WEIGHT_AMOUNT
            ElseIf Sup.CALCULATE_TYPE = 3 Then
                  Set HeadCf = m_Headers1.Fields(10) '���˹ѡ���͢�� > �ط��
                  TempStr = FormatNumberToNull(TempRo.TOTAL_WEIGHT, 2)
                  TMP_WEIGHT = TempRo.TOTAL_WEIGHT
                  Total1(10) = Total1(10) + TempRo.TOTAL_WEIGHT
            End If
             Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
             Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(11) '���˹ѡ��Ҥ�ѧ
            TempStr = FormatNumberToNull(TempRo.TX_AMOUNT, 2)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            Total1(11) = Total1(11) + TempRo.TX_AMOUNT
            
            Set HeadCf = m_Headers1.Fields(12) '�������
            If Sup.CALCULATE_TYPE = 1 Then
                  Set HeadCf = m_Headers1.Fields(12) '�������
                  TempStr = FormatNumberToNull(TempComeIn.SUPPLIER_WEIGHT + TempRo.SUPPLIER_WEIGHT, 2)
                  Total1(12) = Total1(12) + TempComeIn.TX_AMOUNT + TempRo.SUPPLIER_WEIGHT
            ElseIf Sup.CALCULATE_TYPE = 2 Then
                   Set HeadCf = m_Headers1.Fields(12) '�������
                  TempStr = FormatNumberToNull(TempComeIn.WEIGHT_AMOUNT + TempRo.WEIGHT_AMOUNT, 2)
                  Total1(12) = Total1(12) + TempComeIn.TX_AMOUNT + TempRo.WEIGHT_AMOUNT
            ElseIf Sup.CALCULATE_TYPE = 3 Then
                   Set HeadCf = m_Headers1.Fields(12) '�������
                  TempStr = FormatNumberToNull(TempComeIn.TOTAL_WEIGHT + TempRo.TOTAL_WEIGHT, 2)
                  Total1(12) = Total1(12) + TempComeIn.TX_AMOUNT + TempRo.TOTAL_WEIGHT
            End If
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(13) ' ��ҧ��
            If Sup.CALCULATE_TYPE = 1 Then
                  Set HeadCf = m_Headers1.Fields(13) '���˹ѡ���͢�� > �����
                  TempStr = FormatNumberToNull(Sup.TX_AMOUNT - (TempComeIn.SUPPLIER_WEIGHT + TempRo.SUPPLIER_WEIGHT))
                  Total1(13) = Total1(13) + (Sup.TX_AMOUNT - (TempComeIn.SUPPLIER_WEIGHT + TempRo.SUPPLIER_WEIGHT))
            ElseIf Sup.CALCULATE_TYPE = 2 Then
                  Set HeadCf = m_Headers1.Fields(13) '���˹ѡ���͢�� > ���
                  TempStr = FormatNumberToNull(Sup.TX_AMOUNT - (TempComeIn.WEIGHT_AMOUNT + TempRo.WEIGHT_AMOUNT))
                  Total1(13) = Total1(13) + (Sup.TX_AMOUNT - (TempComeIn.WEIGHT_AMOUNT + TempRo.WEIGHT_AMOUNT))
            ElseIf Sup.CALCULATE_TYPE = 3 Then
                  Set HeadCf = m_Headers1.Fields(13) '���˹ѡ���͢�� > �ط��
                  TempStr = FormatNumberToNull(Sup.TX_AMOUNT - (TempComeIn.TOTAL_WEIGHT + TempRo.TOTAL_WEIGHT))
                  Total1(13) = Total1(13) + (Sup.TX_AMOUNT - (TempComeIn.TOTAL_WEIGHT + TempRo.TOTAL_WEIGHT))
            End If
            
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(14) ' �ӹǹ�Թ  ���ͧ���
            TempStr = FormatNumberToNull(TempRo.TOTAL_INCLUDE_PRICE)
            Total1(14) = Total1(14) + (TempRo.TOTAL_INCLUDE_PRICE)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            
            Set HeadCf = m_Headers1.Fields(15)
            TempStr = Sup.NOTE
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            If mcolParam("SHOW_STATUS_PO_FLAG") = 1 Then
               Set HeadCf = m_Headers1.Fields(16)
               TempStr = IIf(Sup.CLOSE_FLAG = "N", "�ѧ���Դ", "�Դ����")
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
               Call m_Details1.AddField(BodyCf)
            End If
            
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
         End If
      End If
      
      Rs.MoveNext
   Wend
   
   Call GenerateFooter("���", "", "", Total1, m_Details1, 0)
   Call m_Details1.GetString(2, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False
   
   Set BodyCf = Nothing
   Vsp.EndDoc
   Call CloseExportFile(Vsp)

   Set RoColl = Nothing
   
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

Private Sub GenerateFooter(TXT1 As String, TXT2 As String, TXT3 As String, Tot() As Double, Details As CFieldList, BalAmount As Double)
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
      ElseIf HeadCf.TempKey = "SUM_FLAG" Then
         TempStr = FormatNumberToNull(Tot(J), 0)
      ElseIf HeadCf.TempKey = "SUM_FLAG_DEC" Then
         TempStr = FormatNumberToNull(Tot(J))
      Else
          TempStr = ""
      End If

      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf

   Set BodyCf = Nothing
End Sub

Private Sub VSP_EndDoc()
'This event occur when VSPrinter.EndDoc is used
Vsp = "*dfdfsffdf"
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
   Call Vsp.AddTable(strFormat, "", "|����|: " & Right$(DateToStringExtEx(mdteDate), 8))
   Call Vsp.AddTable(strFormat, "", "|�����|: " & glbUser.USER_NAME)
   Vsp.CurrentY = dY(0)

   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   Vsp.TextAlign = taCenterBottom
   Vsp.Paragraph = MODULE_DESC
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



