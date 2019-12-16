VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportProduct020"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportProduct013"
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

Private Const TITLE_SIZE = 12
Private Const HEADER_SIZE = 10
Private Const DETAIL_SIZE = 8

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
Private m_SumInput As Collection
Private m_SumYYYYMMColl As Collection
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
   Set m_SumInput = New Collection
   Set m_SumYYYYMMColl = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_SumInput = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
   Set m_SumYYYYMMColl = Nothing
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

Dim TempDate  As Date
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
   Call Cf.SetFieldValue(5, "^", "ชื่อวัตถุดิบ", "<")
   Call m_Headers0.AddField(Cf)
   
   TempDate = mcolParam("FROM_DATE")
   
   While TempDate <= mcolParam("TO_DATE")
      Call Cf.SetFieldValue(6, "^", "เดือน " & Month(TempDate) & " ปี " & Year(TempDate) + 543, ">", TempStr)
      Call m_Headers0.AddField(Cf)
      TempDate = DateAdd("M", 1, TempDate)
   Wend
   
   Call Cf.SetFieldValue(6, "^", "รวม", "<")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(5, "^", "ชื่อวัตถุดิบ", "<")
   Call m_Headers1.AddField(Cf)
   TempDate = mcolParam("FROM_DATE")
   
   While TempDate <= mcolParam("TO_DATE")
      Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(1, "^", "@", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(1, "^", "%", ">")
      Call m_Headers1.AddField(Cf)
      
      TempDate = DateAdd("M", 1, TempDate)
   Wend
   
   Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "@", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "%", ">")
   Call m_Headers1.AddField(Cf)
   
   Call m_Headers0.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)

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
   Vsp.ORIENTATION = orLandscape
   Vsp.MarginBottom = 300
   Vsp.MarginFooter = 300
   Vsp.MarginHeader = 300
   Vsp.MarginLeft = 300
   Vsp.MarginRight = 500
   Vsp.MarginTop = 300
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   Call LoadJobInputByPartItemYYYYMM(Nothing, m_SumInput, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), mcolParam("PROCESS_ID"), mcolParam("PART_GROUP"), mcolParam("PART_TYPE"), mcolParam("LOCATION_ID"))
   Call SetSumMonthItem
   
   FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
   ToDate = "ถึง " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")
   
   MODULE_DESC = vbCrLf & glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & vbCrLf & _
                                       mcolParam("REPORT_NAME") & " " & FromDate & " " & ToDate & vbCrLf & _
                                        "โปรเซส : " & " " & EmptyToString(mcolParam("PROCESS_NAME"), "N/A") & " กลุ่ม : " & EmptyToString(mcolParam("PART_GROUP_NAME"), "N/A") & " ประเภท : " & EmptyToString(mcolParam("PART_TYPE_NAME"), "N/A") & " สถานที่เก็บ : " & EmptyToString(mcolParam("LOCATION_NAME"), "N/A")
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
Dim k As Long
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
Dim Ji As CJobInput
Dim TempJi As CJobInput
Dim TempDate As Date
Dim TempSumJi As CJobInput

Dim AvgPrice As Double
Dim TotalAmount As Double
Dim TotalPrice As Double
Dim SumTotalAmount As Double
Dim SumTotalPrice As Double

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
   
   Set Ji = New CJobInput
   Ji.JOB_INOUT_ID = -1
   Ji.PROCESS_ID = mcolParam("PROCESS_ID")
   Ji.FROM_DATE = mcolParam("FROM_DATE")
   Ji.TO_DATE = mcolParam("TO_DATE")
   Ji.PROCESS_ID = mcolParam("PROCESS_ID")
   Ji.PART_GROUP_ID = mcolParam("PART_GROUP")
   Ji.PART_TYPE_ID = mcolParam("PART_TYPE")
   Ji.TX_TYPE = "E"
   Ji.LOCATION_ID = mcolParam("LOCATION_ID")
   Call Ji.QueryData(11, Rs, iCount)
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
   
      I = 0
      
      Set BodyCf = New CReportField
      While Not Rs.EOF
         I = I + 1
         Call Ji.PopulateFromRS(11, Rs)
         Call m_Details1.ClearField
         
         Set HeadCf = m_Headers1.Fields(1)
         TempStr = I
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(2)
         TempStr = Ji.PART_DESC
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         k = 2
         
         TempDate = mcolParam("FROM_DATE")
         
         SumTotalAmount = 0
         SumTotalPrice = 0
            
         While TempDate <= mcolParam("TO_DATE")
            
            k = k + 1
            
            Set TempJi = GetJobInOut(m_SumInput, Trim(Ji.PART_ITEM_ID & "-" & Year(TempDate) & "-" & Format(Month(TempDate), "00")))
            
            AvgPrice = MyDiffEx(TempJi.TOTAL_INCLUDE_PRICE, TempJi.TX_AMOUNT)
            If mcolParam("MODE") = 1 Then
               TotalAmount = TempJi.TX_AMOUNT
               TotalPrice = TempJi.TOTAL_INCLUDE_PRICE
            ElseIf mcolParam("MODE") = 2 Then
               TotalAmount = TempJi.STD_AMOUNT
               TotalPrice = TempJi.STD_AMOUNT * AvgPrice
            End If
            
            SumTotalAmount = SumTotalAmount + TotalAmount
            SumTotalPrice = SumTotalPrice + TotalPrice
            
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumber(TotalAmount)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(k) = Total1(k) + TotalAmount
                  
            k = k + 1
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumber(AvgPrice)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            k = k + 1
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumber(TotalPrice)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(k) = Total1(k) + TotalPrice
            
            k = k + 1
            Set HeadCf = m_Headers1.Fields(k)
            Set TempSumJi = GetJobInOut(m_SumYYYYMMColl, Trim(Year(TempDate) & "-" & Format(Month(TempDate), "00")))
            TempStr = FormatNumber(MyDiff(TotalPrice * 100, TempSumJi.TOTAL_INCLUDE_PRICE))
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(k) = Total1(k) + MyDiff(TotalPrice * 100, TempSumJi.TOTAL_INCLUDE_PRICE)
            
            If TotalAmount <> 0 Or TotalPrice <> 0 Then
               Call m_SumInput.Remove(Trim(Ji.PART_ITEM_ID & "-" & Year(TempDate) & "-" & Format(Month(TempDate), "00")))
            End If
            
            TempDate = DateAdd("M", 1, TempDate)
         Wend
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(SumTotalAmount)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + SumTotalAmount
               
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(MyDiffEx(SumTotalPrice, SumTotalAmount))
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumber(SumTotalPrice)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + SumTotalPrice
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         Set TempSumJi = GetJobInOut(m_SumYYYYMMColl, "9999-99")
         TempStr = FormatNumber(MyDiff(SumTotalPrice * 100, TempSumJi.TOTAL_INCLUDE_PRICE))
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + MyDiff(SumTotalPrice * 100, TempSumJi.TOTAL_INCLUDE_PRICE)
         
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
   
   Set Ji = Nothing
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
      ElseIf J Mod 4 = 0 Then
         TempStr = FormatNumber(MyDiffEx(Tot(J + 1), Tot(J - 1)))
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
Private Sub SetSumMonthItem()
Dim I As Double
Dim TempSuminput As CJobInput
Dim TempIn As CJobInput
Dim TotalSum As Double
   For Each TempSuminput In m_SumInput
      Set TempIn = GetObject("CJobInput", m_SumYYYYMMColl, TempSuminput.YYYYMM, False)
      If TempIn Is Nothing Then
         Set TempIn = New CJobInput
         TempIn.YYYYMM = TempSuminput.YYYYMM
         TempIn.TX_AMOUNT = TempSuminput.TX_AMOUNT
         TempIn.TOTAL_INCLUDE_PRICE = TempSuminput.TOTAL_INCLUDE_PRICE
         Call m_SumYYYYMMColl.add(TempIn, TempIn.YYYYMM)
         Set TempIn = Nothing
      Else
         TempIn.TX_AMOUNT = TempIn.TX_AMOUNT + TempSuminput.TX_AMOUNT
         TempIn.TOTAL_INCLUDE_PRICE = TempIn.TOTAL_INCLUDE_PRICE + TempSuminput.TOTAL_INCLUDE_PRICE
      End If
      TotalSum = TotalSum + TempSuminput.TOTAL_INCLUDE_PRICE
   Next TempSuminput
   
   Set TempIn = New CJobInput
   TempIn.YYYYMM = "9999-99"
   TempIn.TX_AMOUNT = 0
   TempIn.TOTAL_INCLUDE_PRICE = TotalSum
   Call m_SumYYYYMMColl.add(TempIn, TempIn.YYYYMM)
   Set TempIn = Nothing
End Sub

