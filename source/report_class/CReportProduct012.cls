VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportProduct012"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportProduct012"
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

Private m_ProcessParams As Collection

Private m_RmAmounts As Collection
Private m_ExpenseAmounts As Collection

Private m_SumParamByPartParams As Collection

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
   Set m_RmAmounts = New Collection
   Set m_ProcessParams = New Collection
   Set m_ExpenseAmounts = New Collection
   Set m_SumParamByPartParams = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_RmAmounts = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
   Set m_ProcessParams = Nothing
   Set m_ExpenseAmounts = Nothing
   Set m_SumParamByPartParams = Nothing
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
Dim PP As CParameterProcess

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
   Call Cf.SetFieldValue(3, "^", "รหัสผลิตภัณฑ์", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "ชื่อผลิตภัณฑ์", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ปริมาณยกมา", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ราคา/หน่วย", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "มูลค่ายกมา", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "จำนวนผลิต BAG", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จำนวนผลิต BULK", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "ต้นทุน RM", ">")
   Call m_Headers1.AddField(Cf)
   For Each PP In m_ProcessParams
      Call Cf.SetFieldValue(2, "^", PP.PARAMETER_PROCESS_NAME, ">")
      Call m_Headers1.AddField(Cf)
   Next PP
   Call Cf.SetFieldValue(2, "^", "ต้นทุนรวม", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ต้นทุน/หน่วย", ">")
   Call m_Headers1.AddField(Cf)
            
  Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
            
   Set Cf = Nothing
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
      
   Call LoadSumJobInputPartParam(Nothing, m_SumParamByPartParams, mcolParam("FROM_DATE"), mcolParam("TO_DATE"))
   
   FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
   ToDate = "ถึง " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")
   
   MODULE_DESC = vbCrLf & glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & vbCrLf & _
                                       mcolParam("REPORT_NAME") & vbCrLf & _
                                        FromDate & " " & ToDate & " โปรเซส " & EmptyToString(mcolParam("PROCESS_NAME"), "N/A")
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
Dim Total2(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim Jb As CJob
Dim Ji As CJobInput
Dim PP As CParameterProcess
Dim Jp As CJobParameter
Dim Cost As Double
Dim NewDate As Date
Dim BalanceAccums As Collection
Dim InventoryBals As Collection
Dim Li As CLotItem
Dim Amt As Double
Dim Jip As CJobInput

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set Rs = New ADODB.Recordset
   
   Call LoadParameterProcess(Nothing, m_ProcessParams)
   Call LoadJobProductRMAmount(Nothing, m_RmAmounts, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), "", mcolParam("PROCESS_ID"))
   Call LoadJobProductExpenseAmount(Nothing, m_ExpenseAmounts, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), "", mcolParam("PROCESS_ID"))
   
'   Set BalanceAccums = New Collection
'   Set InventoryBals = New Collection
'   If mcolParam("FROM_DATE") > 0 Then
'      Call LoadInventoryBalanceByPartItem(Nothing, BalanceAccums, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), "")
'      Call glbDaily.CopyBalanceAccumByPartItem(BalanceAccums, InventoryBals)
'   End If
   
   Set InventoryBals = New Collection
   If mcolParam("FROM_DATE") > 0 Then
      Dim MonthlyAccums  As Collection
      Dim YYYYMM As String
      Set MonthlyAccums = New Collection
      YYYYMM = Format(Year(DateAdd("D", -1, mcolParam("FROM_DATE"))), "0000") & "-" & Format(Month(DateAdd("D", -1, mcolParam("FROM_DATE"))), "00")
      Call LoadMonthlyBalancePartItem(Nothing, MonthlyAccums, YYYYMM)
      Call glbDaily.CopyMonthlyAccumPartItem(MonthlyAccums, InventoryBals)
      Set MonthlyAccums = Nothing
   End If

   Set BalanceAccums = Nothing
   
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
   For J = 1 To UBound(Total1)
      Total1(J) = 0
      Total2(J) = 0
   Next J
      
   Set Jb = New CJob
   Jb.JOB_ID = -1
   Jb.PROCESS_ID = mcolParam("PROCESS_ID")
   Jb.FROM_DATE = mcolParam("FROM_DATE")
   Jb.TO_DATE = mcolParam("TO_DATE")
   Call Jb.QueryData(3, Rs, iCount)
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
   
      I = 0
      
      Set BodyCf = New CReportField
      While Not Rs.EOF
         I = I + 1
         Call Jb.PopulateFromRS(3, Rs)
         Call m_Details1.ClearField
         
         Set Ji = GetJobInOut(m_RmAmounts, Jb.PART_ITEM_ID & "-I")
         Set Li = GetLotItem(InventoryBals, Trim(str(Jb.PART_ITEM_ID)))
                  
         Set HeadCf = m_Headers1.Fields(1)
         TempStr = I
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(2)
         TempStr = Jb.PART_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(3)
         TempStr = Jb.PART_DESC
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(4)
         TempStr = FormatNumber(Li.NEW_AMOUNT)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(4) = Total1(4) + Li.NEW_AMOUNT
         
         Set HeadCf = m_Headers1.Fields(5)
         TempStr = FormatNumber(Li.NEW_PRICE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(5) = Total1(5) + Li.NEW_PRICE

         Set HeadCf = m_Headers1.Fields(6)
         TempStr = FormatNumber(Li.NEW_PRICE * Li.NEW_AMOUNT)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(6) = Total1(6) + (Li.NEW_PRICE * Li.NEW_AMOUNT)
         
         If Jb.PARCEL_TYPE = PARCEL_BAG Then
            Amt = Jb.ACTUAL_AMOUNT
         Else
            Amt = 0
         End If
         Set HeadCf = m_Headers1.Fields(7)
         TempStr = FormatNumber(Amt, 3)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(7) = Total1(7) + Amt
         
         If Jb.PARCEL_TYPE = PARCEL_BULK Then
            Amt = Jb.ACTUAL_AMOUNT
         Else
            Amt = 0
         End If
         Set HeadCf = m_Headers1.Fields(8)
         TempStr = FormatNumber(Amt, 3)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(8) = Total1(8) + Amt
         
         Set HeadCf = m_Headers1.Fields(9)
         TempStr = FormatNumber(Ji.RAW_COST)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(9) = Total1(9) + Ji.RAW_COST
         Cost = Ji.RAW_COST
         
         J = 9
         For Each PP In m_ProcessParams
            Set Jp = GetJobParameter(m_ExpenseAmounts, Jb.PART_ITEM_ID & "-" & PP.PARAMETER_PROCESS_ID)
            Set Jip = GetJobInOut(m_SumParamByPartParams, Jb.PART_ITEM_ID & "-" & PP.PARAMETER_PROCESS_ID)
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = FormatNumber(Jp.PARAM_AMOUNT + Jip.TOTAL_INCLUDE_PRICE)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total1(J) = Total1(J) + Jp.PARAM_AMOUNT + Jip.TOTAL_INCLUDE_PRICE
            Cost = Cost + Jp.PARAM_AMOUNT + Jip.TOTAL_INCLUDE_PRICE
         Next PP
         
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = FormatNumber(Cost)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(J) = Total1(J) + Cost
         
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = FormatNumber(MyDiffEx(Cost, Jb.ACTUAL_AMOUNT))
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(J) = Total1(J) + 0
         
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)

         Rs.MoveNext
      Wend

      Call GenerateFooter("รวม", "", "", Total1, m_Details1)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False

      Set BodyCf = Nothing
      
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   Set Jb = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Set InventoryBals = Nothing
   
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
Dim PercentDiff As Double

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
      ElseIf J = 4 Then
         TempStr = FormatNumber(Tot(J))
      ElseIf J = 7 Then
         TempStr = FormatNumber(Tot(J), 3)
      ElseIf J = m_Headers1.Fields.Count Then
         TempStr = FormatNumber(MyDiffEx(Tot(J - 1), Tot(7)))
      Else
         TempStr = FormatNumber(Tot(J))
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub

Private Sub GenerateFooterEx(TXT1 As String, TXT2 As String, TXT3 As String, Tot() As Double, Details As CFieldList, Jb As CJob)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim J As Long
Dim TempStr As String
Dim PercentDiff As Double
Dim Diff As Double

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   For J = 1 To 5
      If J = 1 Then
         TempStr = TXT1
         Set HeadCf = m_Headers0.Fields(1)
      ElseIf J = 2 Then
         TempStr = TXT2
         Set HeadCf = m_Headers0.Fields(2)
      ElseIf J = 3 Then
         TempStr = TXT3
         Set HeadCf = m_Headers0.Fields(3)
      ElseIf J = 4 Then
         Diff = Jb.ACTUAL_AMOUNT - Tot(6)
         TempStr = "ยอดผลิต - ยอดใช้รวม = " & FormatNumber(Diff, 3)
         Set HeadCf = New CReportField
         HeadCf.Width = 8
         HeadCf.Align = "<"
      ElseIf J = 5 Then
         PercentDiff = MyDiffEx(Diff, Tot(6)) * 100
         TempStr = "%DIFF = " & FormatNumber(PercentDiff, 2)
         Set HeadCf = New CReportField
         HeadCf.Width = 6
         HeadCf.Align = "<"
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
      
      Set HeadCf = Nothing
   Next J
      
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
   Call Vsp.AddTable(strFormat, "", "|หน้าที่|: " & Vsp.CurrentPage)
   Call Vsp.AddTable(strFormat, "", "|วันที่พิมพ์|: " & DateToStringExt(mdteDate))
   Call Vsp.AddTable(strFormat, "", "|เวลา|: " & Right$(DateToStringExtEx(mdteDate), 8))
   Call Vsp.AddTable(strFormat, "", "|ผู้ใช้|: " & glbUser.USER_NAME)
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


