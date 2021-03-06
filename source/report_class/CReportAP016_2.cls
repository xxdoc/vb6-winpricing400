VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAP016_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAP016_2"
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

   'Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll

   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   Call m_Headers2.ClearField

   Set Cf = New CReportField

   Call Cf.SetFieldValue(3, "^", "�ѹ����͡���", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "RO", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "PO", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "RM", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "@", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "�ӹǹ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "��Ť��", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "��è�", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�Ѵ��", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "�ú��˹�", "^")
   Call m_Headers1.AddField(Cf)
   Set Cf = Nothing
   
   Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Call m_Headers0.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Vsp.TableBorder = tbdT
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
   Vsp.TableBorder = tbNone
'   Vsp.TableBorder = tbAll
   
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
   Vsp.MarginBottom = 300
   Vsp.MarginFooter = 300
   Vsp.MarginHeader = 300
   Vsp.MarginLeft = 300
   Vsp.MarginRight = 500
   Vsp.MarginTop = 300
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   FromDate = "�ҡ �ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
   ToDate = " �֧ �ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")

   MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & _
                                       vbCrLf & mcolParam("REPORT_NAME") & _
                                       vbCrLf & FromDate & ToDate
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
Dim currentKey1 As String

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
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
   Sup.FROM_DATE = mcolParam("FROM_DATE")
   Sup.TO_DATE = mcolParam("TO_DATE")
   Sup.SUPPLIER_CODE = mcolParam("SUPPLIER_CODE")
   Sup.PART_NO = mcolParam("PART_NO")
   Sup.SUPPLIER_GRADE = mcolParam("SUPPLIER_GRADE")
   Sup.SUPPLIER_TYPE = mcolParam("SUPPLIER_TYPE")
   Sup.DOCUMENT_TYPE_SET = "(100,101,102,103)"
   Dim DocType As String
   DocType = mcolParam("RO_TYPE")
   If DocType = "0" Then
      Sup.DOCUMENT_TYPE_SET = "(100,101,102,103)"
   Else
      Sup.DOCUMENT_TYPE_SET = "(" & DocType & ")"
   End If
   
   Call Sup.QueryData(114, Rs, iCount)
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
   
   I = 0
   k = 0
   Set BodyCf = New CReportField
'   Dim PrevKey1 As String
   PrevKey1 = ""
   While Not Rs.EOF
      Call m_Details1.ClearField
      Call Sup.PopulateFromRS(114, Rs)
         currentKey1 = Sup.SUPPLIER_CODE
         If PrevKey1 <> currentKey1 And I > 0 Then
             Vsp.FontBold = True
            Call GenerateFooter("���", "", PrevKey1, Total1, m_Details1, 0)
             
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False

            For J = 1 To UBound(Total1)
                Total1(J) = 0
            Next J
            
            Vsp = Sup.SUPPLIER_NAME & "(" & Sup.SUPPLIER_CODE & ")"
         ElseIf I = 0 Then
             Vsp = Sup.SUPPLIER_NAME & "(" & Sup.SUPPLIER_CODE & ")"
         End If
         
         I = I + 1
         PrevKey1 = Sup.SUPPLIER_CODE
         Call m_Details1.ClearField
'         i = i + 1
         
         Set HeadCf = m_Headers1.Fields(1) '�ѹ����͡���
         TempStr = DateToStringExtEx2(Sup.DOCUMENT_DATE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(2) 'RO
         TempStr = Sup.RO_DOCUMENT_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(3) 'PO
         TempStr = Sup.PO_DOCUMENT_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(4) 'RM
         TempStr = Sup.PART_DESC
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
                  
         Set HeadCf = m_Headers1.Fields(5) '@
         TempStr = FormatNumberToNull(Sup.ACTUAL_UNIT_PRICE, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(5) = Total1(5) + Sup.ACTUAL_UNIT_PRICE
         
         If Sup.CALCULATE_TYPE = 1 Then
               Set HeadCf = m_Headers1.Fields(6) '���˹ѡ���͢�� > �����
               TempStr = FormatNumberToNull(Sup.SUPPLIER_WEIGHT, 2)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
               Call m_Details1.AddField(BodyCf)
               Total1(6) = Total1(6) + Sup.SUPPLIER_WEIGHT
       ElseIf Sup.CALCULATE_TYPE = 2 Then
            Set HeadCf = m_Headers1.Fields(6) '���˹ѡ���͢�� > ���
            TempStr = FormatNumberToNull(Sup.WEIGHT_AMOUNT, 2)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            Total1(6) = Total1(6) + Sup.WEIGHT_AMOUNT
      ElseIf Sup.CALCULATE_TYPE = 3 Then
            Set HeadCf = m_Headers1.Fields(6) '���˹ѡ���͢�� > �ط��
            TempStr = FormatNumberToNull(Sup.TOTAL_WEIGHT, 2)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            Total1(6) = Total1(6) + Sup.TOTAL_WEIGHT
     ElseIf Sup.CALCULATE_TYPE = 0 Then
            Set HeadCf = m_Headers1.Fields(6) '���˹ѡ���͢�� > �ط��
            TempStr = ""
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            Total1(6) = Total1(6)
      End If
         
'         Set HeadCf = m_Headers1.Fields(6) '�ӹǹ
'         TempStr = FormatNumberToNull(Sup.TX_AMOUNT, 2)
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
'         Call m_Details1.AddField(BodyCf)
'         Total1(6) = Total1(6) + Sup.TX_AMOUNT
            
         Set HeadCf = m_Headers1.Fields(7) '��Ť��
         TempStr = FormatNumber(Sup.TOTAL_PRICE - Sup.DISCOUNT_AMOUNT)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(7) = Total1(7) + (Sup.TOTAL_PRICE - Sup.DISCOUNT_AMOUNT)
         
         Set HeadCf = m_Headers1.Fields(8) '�Ҫ�к�è�
         TempStr = Sup.PACKAGING_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(9) 'ʶҹ���Ѵ��
         TempStr = Sup.LOCATION_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(10) '�ú��˹�
         TempStr = DateToStringExtEx2(Sup.DUE_DATE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         
      Rs.MoveNext
   Wend
   
   Call GenerateFooter("���", "", PrevKey1, Total1, m_Details1, 0)
     
    Call m_Details1.GetString(2, TempStr1, TempStr2)
    strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
    Call Vsp.AddTable(strFormat, "", TempStr2)
    Vsp.FontBold = False
            
   Set BodyCf = Nothing
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   'Set RoColl = Nothing
   
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
      ElseIf J = 4 Then
         TempStr = TXT3
      ElseIf J = 5 Then
         TempStr = FormatNumber(Tot(J))
      ElseIf J = 6 Then
         TempStr = FormatNumber(Tot(J))
      ElseIf J = 7 Then
         TempStr = FormatNumber(Tot(J))
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
   Vsp.Paragraph = ""
   Vsp.FontSize = iSize
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
   Vsp.TableBorder = tbAll
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

