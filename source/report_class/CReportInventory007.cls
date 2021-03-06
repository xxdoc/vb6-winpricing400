VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportInventory007"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportInventory007"
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
Private m_Details1 As CFieldList
Private m_Left As Double
Private m_YearSeqs As Collection

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_TempCol = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Details1 = New CFieldList
   Set m_YearSeqs = New Collection
   
   Set Rs = New ADODB.Recordset
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
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
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
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField

   Call Cf.SetFieldValue(1.5, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "�Ţ����Ѻ", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "�ѹ����Ѻ", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "���� R/M", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(5, "^", "���� R/M", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "���ʼ����", "^")
   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(5, "^", "���ͼ����", "<")
'   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(2, "^", "�Ҫ�� ��è�", "^")
'   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�ӹǹ �Ѻ��ԧ", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3.5, "^", "�.�. ���", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�.�. ��è�", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "% Moist", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�.�. Moist", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�Ѵ��� �", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3.5, "^", "�.�. �ط��", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3.5, "^", "�.�. ���͢��", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3.5, "^", "�.�. ��Ҥ�ѧ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "��Ť��", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "�.�.�. �Ѵ����", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "�ط��", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�Ҥ�", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "��ѧ", "^")
   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(2, "^", "�Դ �.�.Ẻ", "^")
'   Call m_Headers1.AddField(Cf)
   
   Set Cf = Nothing
   
  Call m_Headers1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr2
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
   
   FromDate = "�ҡ�ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
   ToDate = "�֧ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")
   
   MODULE_DESC = vbCrLf & glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & vbCrLf & _
                                       mcolParam("REPORT_NAME") & vbCrLf & _
                                        FromDate & " " & ToDate
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
Dim cData As CLotItem
Dim I As Long
Dim J As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim Amt As Double
Dim Li As CLotItem
Dim TempLi As CLotItem
Dim Sum1 As Double
Dim Count1 As Double
Dim CalculateWeight As Double

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

   Set cData = New CLotItem
   cData.LOT_ITEM_ID = -1
   cData.PART_NO = mcolParam("PART_NO")
   cData.PART_GROUP_ID = mcolParam("PART_GROUP")
   cData.PART_TYPE = mcolParam("PART_TYPE")
   cData.LOCATION_ID = mcolParam("LOCATION_ID")
   cData.FROM_DATE = mcolParam("FROM_DATE")
   cData.TO_DATE = mcolParam("TO_DATE")
   cData.FROM_TX_TYPE = "I"
   cData.COMMIT_FLAG = ""
   cData.DOCUMENT_TYPE = 1
   cData.PART_NO = mcolParam("PART_NO")
   cData.SUPPLIER_CODE = mcolParam("SUPPLIER_CODE")
   cData.OrderBy = mcolParam("ORDER_BY")
   cData.OrderType = mcolParam("ORDER_TYPE")
   Call cData.QueryData(1, Rs, iCount)
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
   
      I = 0

      Set BodyCf = New CReportField
      Set Li = New CLotItem
      While Not Rs.EOF
         I = I + 1
         Call Li.PopulateFromRS(1, Rs)
         Call m_Details1.ClearField

         Set HeadCf = m_Headers1.Fields(1)
         TempStr = I
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(2)
         TempStr = Li.DOCUMENT_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(3)
         TempStr = DateToStringExtEx2(Li.DOCUMENT_DATE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(4)
         TempStr = Li.PART_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(5)
         TempStr = Li.PART_DESC
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
                  
         Set HeadCf = m_Headers1.Fields(6)
         TempStr = Li.SUPPLIER_CODE
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(7)
         TempStr = FormatNumber(Li.ACTUAL_PKG_AMOUNT, 0)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(7) = Total1(7) + Li.ACTUAL_PKG_AMOUNT
         
         Set HeadCf = m_Headers1.Fields(8)
         TempStr = ""
         If Li.CALCULATE_TYPE = 2 Then
         TempStr = "* "
         End If
         TempStr = TempStr & FormatNumber(Li.WEIGHT_AMOUNT, 3)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(8) = Total1(8) + Li.WEIGHT_AMOUNT
         
         Set HeadCf = m_Headers1.Fields(9)
         TempStr = FormatNumber(Li.PACKAGE_WEIGHT, 3)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(9) = Total1(9) + Li.PACKAGE_WEIGHT
         
         Set HeadCf = m_Headers1.Fields(10)
         TempStr = FormatNumber(Li.PERCENT_HUMID)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(10) = Total1(10) + Li.PERCENT_HUMID
         
         Set HeadCf = m_Headers1.Fields(11)
         TempStr = FormatNumber(Li.HUMID_WEIGHT, 3)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(11) = Total1(11) + Li.HUMID_WEIGHT
         
         Set HeadCf = m_Headers1.Fields(12)
         TempStr = FormatNumber(Li.OTHER_WEIGHT, 3)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(12) = Total1(12) + Li.OTHER_WEIGHT
         
         Set HeadCf = m_Headers1.Fields(13)
         TempStr = ""
         If Li.CALCULATE_TYPE = 3 Then
           TempStr = "* "
         End If
         TempStr = TempStr & FormatNumber(Li.TOTAL_WEIGHT, 3)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(13) = Total1(13) + Li.TOTAL_WEIGHT
         
         CalculateWeight = Li.CALCULATE_WEIGHT

         Set HeadCf = m_Headers1.Fields(14)
         TempStr = ""
         If Li.CALCULATE_TYPE = 1 Then
            TempStr = "* "
         End If
         TempStr = TempStr & FormatNumber(CalculateWeight, 3)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(14) = Total1(14) + CalculateWeight
         
         Set HeadCf = m_Headers1.Fields(15)
         TempStr = FormatNumber(Li.TX_AMOUNT, 3)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(15) = Total1(15) + Li.TX_AMOUNT
         
         Set HeadCf = m_Headers1.Fields(16)
         TempStr = FormatNumber(Li.TOTAL_ACTUAL_PRICE, 3)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(16) = Total1(16) + Li.TOTAL_ACTUAL_PRICE
         
         Set HeadCf = m_Headers1.Fields(17)
         TempStr = FormatNumber(Li.EXPENSE1 + Li.EXPENSE2, 3)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(17) = Total1(17) + (Li.EXPENSE1 + Li.EXPENSE2)
                  
         Set HeadCf = m_Headers1.Fields(18)
         TempStr = FormatNumber(Li.TOTAL_INCLUDE_PRICE, 3)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(18) = Total1(18) + Li.TOTAL_INCLUDE_PRICE
                  
         Set HeadCf = m_Headers1.Fields(19)
         TempStr = FormatNumber(MyDiffEx(Li.TOTAL_INCLUDE_PRICE, Li.TX_AMOUNT), 3)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(19) = Total1(19) + 0
                  
         Set HeadCf = m_Headers1.Fields(20)
         TempStr = Li.LOCATION_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(20) = Total1(20) + 0
         
'         Set HeadCf = m_Headers1.Fields(21)
'         TempStr = CalculateTypeToText2(Li.CALCULATE_TYPE)
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
'         Call m_Details1.AddField(BodyCf)
                  
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         '=====

         Rs.MoveNext
      Wend

      Call GenerateFooter("���", "", "", Total1, m_Details1)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      
'      Vsp = "* 1=" & CalculateTypeToText(1) & ",2=" & CalculateTypeToText(2) & ",3=" & CalculateTypeToText(3)

      Set BodyCf = Nothing
      Set Li = Nothing

   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   Set cData = Nothing
   
   genDoc = True
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
   Set Rs = Nothing
End Function

Private Sub GenerateFooter(TXT1 As String, TXT2 As String, TXT3 As String, Tot() As Double, Details As CFieldList)
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
      ElseIf (J = 7) Then
         TempStr = FormatNumber(Tot(J), 0)
      ElseIf (J = 8) Then
         TempStr = FormatNumber(Tot(J), 3)
      ElseIf (J = 9) Then
         TempStr = FormatNumber(Tot(J), 3)
      ElseIf (J = 10) Then
         TempStr = FormatNumber(Tot(J), 2)
      ElseIf (J >= 11) And (J <= 18) Then
         TempStr = FormatNumber(Tot(J), 3)
      ElseIf (J = 19) Then
         TempStr = FormatNumber(MyDiffEx(Tot(18), Tot(15)), 3)
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
   Call Vsp.AddTable(strFormat, "", "|�ѹ�������|: " & DateToStringExt(mdteDate))
   Call Vsp.AddTable(strFormat, "", "|����|: " & Right$(DateToStringExtEx(mdteDate), 8))
   Call Vsp.AddTable(strFormat, "", "|�����|: " & glbUser.USER_NAME)
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

