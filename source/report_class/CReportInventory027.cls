VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportInventory001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportInventory027"
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
Private Const DETAIL_SIZE = 10

Private Const TITLE_FORMAT = "<32|<3|<5;"
Private Const HEADER_FORMAT1 = "^1|^3|^2|^2|^2|%P^2|^2;"
Private Const HEADER_FORMAT2 = "^1|^3|^2|^2|^2|%P^2|^2;"
Private Const DETAIL_FORMAT1 = "<4|<4|%N;"
Private Const DETAIL_FORMAT2 = "^1|^3|>2|>2|>2|%P>2|>2;"
Private Const FOOTER_FORMAT1 = "^1|^3|>2|>2|%P>2|>2;"
Private Const FOOTER_FORMAT2 = "^1|^3|>2|>2|>2|%P>2|>2;"
'NO.|�ѹ���|�Ţ����͡���|�Դ|�͹���|%P|�͹�͡|�������;"

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
End Sub

Private Function CReportInterface_AddParam(varItem As Variant, strIndex As String) As Boolean
   Call mcolParam.Add(varItem, strIndex)
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
Dim Yw As CYearSeq

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
   Call m_Headers0.AddField(Cf)
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "�ѻ�����Դ", "^")
   Call m_Headers0.AddField(Cf)
   Call m_Headers1.AddField(Cf)
                                                                                                                                                                                                                                                   
   ExportWidth = 0
   For Each Yw In m_YearSeqs
      Call Cf.SetFieldValue(2, "^", Yw.YEAR_NO, "^", Trim(Str(Yw.YEAR_NO)))
      Call m_Headers1.AddField(Cf)
      
      ExportWidth = ExportWidth + 2
   Next Yw
   
   Call Cf.SetFieldValue(CDbl(ExportWidth), "^", "�ѻ�����Դ", "^")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "���", "^")
   Call m_Headers1.AddField(Cf)
   Call m_Headers0.AddField(Cf)
   
   Set Cf = Nothing
   
  Call m_Headers0.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
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

   Call LoadYearSeq(Nothing, m_YearSeqs)
   
   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprFanfoldUS 'pprA4 '
   Vsp.ORIENTATION = orPortrait
   Vsp.MarginBottom = 1440
   Vsp.MarginFooter = 1440
   Vsp.MarginHeader = 1440
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 1440
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   FromDate = "� �ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")
'   ToDate = "�֧ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")
   
   MODULE_DESC = vbCrLf & glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & vbCrLf & _
                                       mcolParam("REPORT_NAME") & _
                                       vbCrLf & FromDate & " ������ " & mcolParam("PIG_TYPE_NAME")
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
Dim InventoryBals As Collection
Dim RName As String
Dim cData As CPartLocation
Dim i As Long
Dim j As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Amt As Double
Dim Ei As CExportItem
Dim Ii As CLotItem
Dim Rs1 As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim TxCode As String
Dim PartItems As Collection
Dim Houses As Collection
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Hs As CLocation
Dim Lc As CLocation
Dim Pi As CPartItem
Dim Ib As CLotItem
Dim Total1(100) As Double
Dim StatusMatch As Boolean
Dim PiCount As Long
Dim iCount As Long
Dim TempStr As String
Dim Locations As Collection
Dim Sum1 As Double
Dim PrevPI As CPartItem
Dim Yw As CYearWeek
Dim YearWeeks As Collection
Dim PartItemID As Long
Dim NewDate As Date
Dim PigTypeCode As String
Dim ImportItems As Collection
Dim ExportItems As Collection

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set YearWeeks = New Collection
   Call LoadDistinctYearWeek(Nothing, YearWeeks)
   
   Set PartItems = New Collection
   Call LoadPigItem(Nothing, PartItems)
   
   Set ImportItems = New Collection
   Call LoadPigImportAmount(Nothing, ImportItems, -1, mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")))
   
   Set ExportItems = New Collection
   Call LoadPigExportAmount(Nothing, ExportItems, -1, mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")))
   
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------

   For j = 1 To UBound(Total1)
      Total1(j) = 0
   Next j
   
   Vsp.StartDoc
      i = 0
      
      For Each Yw In YearWeeks
         Set BodyCf = New CReportField
         
         i = i + 1
         j = 0
         Call m_Details1.ClearField
         
         Sum1 = 0
         For Each HeadCf In m_Headers1.Fields
            j = j + 1
            
            PigTypeCode = PigTypeToCode(mcolParam("PIG_TYPE_ID"))
            Set Pi = GetPartItem(PartItems, HeadCf.TempKey & Format(Yw.WEEK_NO, "00") & "-" & PigTypeCode)
            
            Set Ii = GetImportItem(ImportItems, Trim(Str(Pi.PART_ITEM_ID)))
            Set Ei = GetExportItem(ExportItems, Trim(Str(Pi.PART_ITEM_ID)))

            If j = 1 Then
               TempStr = Trim(Str(i))
            ElseIf j = 2 Then
               TempStr = Format(Yw.WEEK_NO, "00")
            ElseIf j = m_Headers1.Fields.Count Then
               Amt = Sum1
               TempStr = Trim(Str(Amt))
               Total1(j) = Total1(j) + Amt
            Else
               Amt = Ii.TX_AMOUNT - Ei.EXPORT_AMOUNT
               TempStr = Trim(Str(Amt))
               Sum1 = Sum1 + Amt
               
               Total1(j) = Total1(j) + Amt
            End If
            
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
         Next HeadCf
         
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         
         Set BodyCf = Nothing
      Next Yw
      
      Call GenerateFooter("���", "", "", Total1, m_Details1)
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.EndDoc
   
   Set Rs2 = Nothing
         
   Set YearWeeks = Nothing
   Set PartItems = Nothing
   Set ImportItems = Nothing
   Set ExportItems = Nothing
   
   genDoc = True
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
   Set Rs = Nothing
End Function

Private Sub GenerateFooter(Txt1 As String, Txt2 As String, Txt3 As String, Tot() As Double, Details As CFieldList)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   j = 0
   For Each HeadCf In m_Headers1.Fields
      j = j + 1
      If j = 1 Then
         TempStr = Txt1
      ElseIf j = 2 Then
         TempStr = Txt2
      Else
         TempStr = Trim(Str((Tot(j))))
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
   Call Vsp.AddTable(strFormat, "", "|�Ţ��Шӵ�Ǽ����������|: " & glbEnterPrise.TAX_ID)
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


