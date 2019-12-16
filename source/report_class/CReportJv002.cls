VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportJv002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportJV002"
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

Private Const TITLE_SIZE = 16
Private Const HEADER_SIZE = 14
Private Const DETAIL_SIZE = 14

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

Private m_BankAccounts As Collection
Private m_ReportConfig As CReportConfig


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
   Set m_BankAccounts = New Collection
   Set m_ReportConfig = New CReportConfig
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_BankAccounts = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
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
Dim Ct As CCashTran

   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbNone
   
   Call GenerateDocHeader(1, Vsp, mcolParam, tbNone, m_ReportConfig.HEAD_OFFSET)
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField

   Set Cf = New CReportField

   '===
   Call Cf.SetFieldValue(3, "^", "รหัสบัญชี", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "ชื่อบัญชี", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "เดบิต", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "เครดิต", ">")
   Call m_Headers1.AddField(Cf)
   
   
   Set Cf = Nothing

'  Call m_Headers0.GetString(1, TempStr1, TempStr1_1)
'   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'   strPrint = TempStr1_1
'   Call Vsp.AddTable(strFormat, "", strPrint)
'
'  Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
'   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'   strPrint = TempStr1_1
'   Call Vsp.AddTable(strFormat, "", strPrint)
            
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
   Vsp.ORIENTATION = orPortrait
'   Vsp.MarginBottom = 700
'   Vsp.MarginFooter = 700
'   Vsp.MarginHeader = 700
'   Vsp.MarginLeft = 700
'   Vsp.MarginRight = 700
   Vsp.MarginTop = 300
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   Call LoadSumCashTranAccountCusAccList(Nothing, m_BankAccounts, mcolParam("FROM_DATE"), mcolParam("FROM_DATE"), "I")
   
   FromDate = "ณ วันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
   
   MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & vbCrLf & mcolParam("REPORT_NAME") & _
                                       vbCrLf & FromDate & ToDate & "  " & ToPaidDate
   
   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"), m_ReportConfig, False)
   
   Vsp.MarginLeft = 576 * 1.4 - m_ReportConfig.DUMMY_OFFSET
   
   'mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
      mstrErrMsg = "Error GetDateTime Error."
      Exit Function
   End If
   mdteDate = InternalDateToDate(strDate)
   initDoc = True
End Function
Private Function GetKey(Ct As CCashTran) As String
   GetKey = DateToStringInt(Ct.GetFieldValue("TX_DATE"))
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
Dim Width As Double
Dim TempRs As ADODB.Recordset
Dim Hc As CReportField
Dim Ct As CCashTran
Dim Ct2 As CCashTran
Dim Ct3 As CCashTran
Dim Ct4 As CCashTran
Dim TempCt As CCashTran
Dim Ba As CCashTran
Dim Sum1 As Double
Dim Ri As CReceiptItem

Dim MainNo  As String
Dim RemainNo As String

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
   
   Vsp.TableBorder = tbNone
   
   k = 0
   
      For Each Ct In m_BankAccounts
         k = k + 1
         
         If Len(Ct.GetFieldValue("ACC_NAME1") & " รับชำระหนี้จาก " & Ct.GetFieldValue("CUSTOMER_NAME")) > 84 Then
            If k > 18 Then
               k = 1
               Vsp.NewPage
            End If
         ElseIf Len(Ct.GetFieldValue("ACC_NAME1") & " รับชำระหนี้จาก " & Ct.GetFieldValue("CUSTOMER_NAME")) > 42 Then
            If k > 19 Then
               k = 1
               Vsp.NewPage
            End If
         Else
            If k > 20 Then
               k = 1
               Vsp.NewPage
            End If
         End If
         
         '============
         Vsp.StartTable
         Vsp.TableCell(tcCols) = 6
         Vsp.TableCell(tcRows) = 1
         If Len(Ct.GetFieldValue("ACC_NAME1") & " รับชำระหนี้จาก " & Ct.GetFieldValue("CUSTOMER_NAME")) > 84 Then
            Vsp.TableCell(tcRowHeight, 1) = 2.4 * 567
            k = k + 2
         ElseIf Len(Ct.GetFieldValue("ACC_NAME1") & " รับชำระหนี้จาก " & Ct.GetFieldValue("CUSTOMER_NAME")) > 42 Then
            Vsp.TableCell(tcRowHeight, 1) = 1.6 * 567
            k = k + 1
         Else
            Vsp.TableCell(tcRowHeight, 1) = 0.8 * 567
         End If
        Vsp.TableCell(tcColWidth, , 1) = "2.8cm"
        Vsp.TableCell(tcColWidth, , 2) = "7.2cm"
        Vsp.TableCell(tcColWidth, , 3) = "3cm"
        Vsp.TableCell(tcColWidth, , 4) = "1cm"
         Vsp.TableCell(tcColWidth, , 5) = "3.2cm"
         Vsp.TableCell(tcColWidth, , 6) = "1cm"
         
        Vsp.TableCell(tcAlign, 1, 1) = taCenterTop
        Vsp.TableCell(tcText, 1, 1) = Ct.GetFieldValue("ACC_CODE1")
         
         Vsp.TableCell(tcAlign, 1, 2) = taLeftTop
         Vsp.TableCell(tcText, 1, 2) = Ct.GetFieldValue("ACC_NAME1") & " รับชำระหนี้จาก " & Ct.GetFieldValue("CUSTOMER_NAME")
        
        MainNo = Left(FormatNumber(Ct.GetFieldValue("NET_AMOUNT")), Len(FormatNumber(Ct.GetFieldValue("NET_AMOUNT"))) - 3)
        RemainNo = Right(FormatNumber(Ct.GetFieldValue("NET_AMOUNT")), 2)
        
         Vsp.TableCell(tcAlign, 1, 3) = taRightTop
         Vsp.TableCell(tcText, 1, 3) = MainNo
         Total1(3) = Total1(3) + (Ct.GetFieldValue("NET_AMOUNT"))
        
         Vsp.TableCell(tcAlign, 1, 4) = taRightTop
         Vsp.TableCell(tcText, 1, 4) = RemainNo
         
         Vsp.TableCell(tcAlign, 1, 5) = taRightTop
         Vsp.TableCell(tcText, 1, 5) = ""
        
         Vsp.TableCell(tcAlign, 1, 6) = taRightTop
         Vsp.TableCell(tcText, 1, 6) = ""
        
         Vsp.EndTable
         '---------------------------------------------------------------------------------------------------------------------------->
         
         If Ct.GetFieldValue("FEE_AMOUNT") > 0 Then
            
            k = k + 1
            
            If Len("ค่าธรรมเนียมธนาคาร" & " รับชำระหนี้จาก " & Ct.GetFieldValue("CUSTOMER_NAME")) > 84 Then
               If k > 18 Then
                  k = 1
                  Vsp.NewPage
               End If
            ElseIf Len("ค่าธรรมเนียมธนาคาร" & " รับชำระหนี้จาก " & Ct.GetFieldValue("CUSTOMER_NAME")) > 42 Then
               If k > 19 Then
                  k = 1
                  Vsp.NewPage
               End If
            Else
               If k > 20 Then
                  k = 1
                  Vsp.NewPage
               End If
            End If
            
            '============
            Vsp.StartTable
            Vsp.TableCell(tcCols) = 6
            Vsp.TableCell(tcRows) = 1
            If Len("ค่าธรรมเนียมธนาคาร" & " รับชำระหนี้จาก " & Ct.GetFieldValue("CUSTOMER_NAME")) > 84 Then
               Vsp.TableCell(tcRowHeight, 1) = 2.4 * 567
               k = k + 2
            ElseIf Len("ค่าธรรมเนียมธนาคาร" & " รับชำระหนี้จาก " & Ct.GetFieldValue("CUSTOMER_NAME")) > 42 Then
               Vsp.TableCell(tcRowHeight, 1) = 1.6 * 567
               k = k + 1
            Else
               Vsp.TableCell(tcRowHeight, 1) = 0.8 * 567
            End If
            Vsp.TableCell(tcColWidth, , 1) = "2.8cm"
            Vsp.TableCell(tcColWidth, , 2) = "7.2cm"
            Vsp.TableCell(tcColWidth, , 3) = "3cm"
            Vsp.TableCell(tcColWidth, , 4) = "1cm"
             Vsp.TableCell(tcColWidth, , 5) = "3.2cm"
             Vsp.TableCell(tcColWidth, , 6) = "1cm"
             
            Vsp.TableCell(tcAlign, 1, 1) = taCenterTop
            Vsp.TableCell(tcText, 1, 1) = "740-4012"
             
             Vsp.TableCell(tcAlign, 1, 2) = taLeftTop
             Vsp.TableCell(tcText, 1, 2) = "ค่าธรรมเนียมธนาคาร" & " รับชำระหนี้จาก " & Ct.GetFieldValue("CUSTOMER_NAME")
            
            MainNo = Left(FormatNumber(Ct.GetFieldValue("FEE_AMOUNT")), Len(FormatNumber(Ct.GetFieldValue("FEE_AMOUNT"))) - 3)
            RemainNo = Right(FormatNumber(Ct.GetFieldValue("FEE_AMOUNT")), 2)
            
             Vsp.TableCell(tcAlign, 1, 3) = taRightTop
             Vsp.TableCell(tcText, 1, 3) = MainNo
             Total1(3) = Total1(3) + (Ct.GetFieldValue("FEE_AMOUNT"))
            
             Vsp.TableCell(tcAlign, 1, 4) = taRightTop
             Vsp.TableCell(tcText, 1, 4) = RemainNo
             
             Vsp.TableCell(tcAlign, 1, 5) = taRightTop
             Vsp.TableCell(tcText, 1, 5) = ""
            
             Vsp.TableCell(tcAlign, 1, 6) = taRightTop
             Vsp.TableCell(tcText, 1, 6) = ""
            
             Vsp.EndTable
         End If
         '---------------------------------------------------------------------------------------------------------------------------->
            
            k = k + 1
                     
            If Len(Ct.GetFieldValue("ACC_NAME2") & " รับชำระหนี้จาก " & Ct.GetFieldValue("CUSTOMER_NAME")) > 84 Then
               If k > 18 Then
                  k = 1
                  Vsp.NewPage
               End If
            ElseIf Len(Ct.GetFieldValue("ACC_NAME2") & " รับชำระหนี้จาก " & Ct.GetFieldValue("CUSTOMER_NAME")) > 42 Then
               If k > 19 Then
                  k = 1
                  Vsp.NewPage
               End If
            Else
               If k > 20 Then
                  k = 1
                  Vsp.NewPage
               End If
            End If
            
            '============
            Vsp.StartTable
            Vsp.TableCell(tcCols) = 6
            Vsp.TableCell(tcRows) = 1
            If Len(Ct.GetFieldValue("ACC_NAME2") & " รับชำระหนี้จาก " & Ct.GetFieldValue("CUSTOMER_NAME")) > 84 Then
               Vsp.TableCell(tcRowHeight, 1) = 2.4 * 567
               k = k + 2
            ElseIf Len(Ct.GetFieldValue("ACC_NAME2") & " รับชำระหนี้จาก " & Ct.GetFieldValue("CUSTOMER_NAME")) > 42 Then
               Vsp.TableCell(tcRowHeight, 1) = 1.6 * 567
               k = k + 1
            Else
               Vsp.TableCell(tcRowHeight, 1) = 0.8 * 567
            End If
            Vsp.TableCell(tcColWidth, , 1) = "2.8cm"
            Vsp.TableCell(tcColWidth, , 2) = "7.2cm"
            Vsp.TableCell(tcColWidth, , 3) = "3cm"
            Vsp.TableCell(tcColWidth, , 4) = "1cm"
             Vsp.TableCell(tcColWidth, , 5) = "3.2cm"
             Vsp.TableCell(tcColWidth, , 6) = "1cm"
             
            Vsp.TableCell(tcAlign, 1, 1) = taCenterTop
            Vsp.TableCell(tcText, 1, 1) = Ct.GetFieldValue("ACC_CODE2")
             
             Vsp.TableCell(tcAlign, 1, 2) = taLeftTop
             Vsp.TableCell(tcText, 1, 2) = Ct.GetFieldValue("ACC_NAME2") & " รับชำระหนี้จาก " & Ct.GetFieldValue("CUSTOMER_NAME")
            
            MainNo = Left(FormatNumber(Ct.GetFieldValue("AMOUNT")), Len(FormatNumber(Ct.GetFieldValue("AMOUNT"))) - 3)
            RemainNo = Right(FormatNumber(Ct.GetFieldValue("AMOUNT")), 2)
            
             Vsp.TableCell(tcAlign, 1, 3) = taRightTop
             Vsp.TableCell(tcText, 1, 3) = ""
             
             Vsp.TableCell(tcAlign, 1, 4) = taRightTop
             Vsp.TableCell(tcText, 1, 4) = ""
             
             Vsp.TableCell(tcAlign, 1, 5) = taRightTop
             Vsp.TableCell(tcText, 1, 5) = MainNo
            Total1(5) = Total1(5) + (Ct.GetFieldValue("AMOUNT"))
            
             Vsp.TableCell(tcAlign, 1, 6) = taRightTop
             Vsp.TableCell(tcText, 1, 6) = RemainNo
            
             Vsp.EndTable
            
      Next Ct

      Vsp.FontBold = True
      
      'k = k + 1
                     
      If k > 20 Then
         k = 1
         Vsp.NewPage
      End If
                     
      '============
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 6
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight, 1) = 0.8 * 567
      Vsp.TableCell(tcColWidth, , 1) = "2.8cm"
      Vsp.TableCell(tcColWidth, , 2) = "7.2cm"
      Vsp.TableCell(tcColWidth, , 3) = "3cm"
      Vsp.TableCell(tcColWidth, , 4) = "1cm"
       Vsp.TableCell(tcColWidth, , 5) = "3.2cm"
       Vsp.TableCell(tcColWidth, , 6) = "1cm"
       
      Vsp.TableCell(tcAlign, 1, 1) = taCenterMiddle
      Vsp.TableCell(tcText, 1, 1) = "สุทธิ"
       
      Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
      Vsp.TableCell(tcText, 1, 2) = ""
      
      MainNo = Left(FormatNumber(Total1(3)), Len(FormatNumber(Total1(3))) - 3)
      RemainNo = Right(FormatNumber(Total1(3)), 2)
      
       Vsp.TableCell(tcAlign, 1, 3) = taRightMiddle
       Vsp.TableCell(tcText, 1, 3) = MainNo
       
       Vsp.TableCell(tcAlign, 1, 4) = taRightMiddle
       Vsp.TableCell(tcText, 1, 4) = RemainNo
       
       MainNo = Left(FormatNumber(Total1(5)), Len(FormatNumber(Total1(5))) - 3)
      RemainNo = Right(FormatNumber(Total1(5)), 2)
       
       Vsp.TableCell(tcAlign, 1, 5) = taRightMiddle
       Vsp.TableCell(tcText, 1, 5) = MainNo
      
       Vsp.TableCell(tcAlign, 1, 6) = taRightMiddle
       Vsp.TableCell(tcText, 1, 6) = RemainNo
      
       Vsp.EndTable
             
      Vsp.FontBold = False

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

Private Sub GenerateFooter(TXT1 As String, TXT2 As String, Tot() As Double, Details As CFieldList, DateCount As Long)
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
   'Vsp.TextAlign = taLeftBottom
   strFormat = VSP_CalTable(TITLE_FORMAT, mdblWid, alngX)
   Vsp.TableBorder = tbNone
'   Call Vsp.AddTable(strFormat, "", "|หน้าที่|: " & Vsp.CurrentPage)
'   Call Vsp.AddTable(strFormat, "", "|วัน|: " & DateToStringExt(mdteDate))
'   Call Vsp.AddTable(strFormat, "", "|เวลา|: " & Right$(DateToStringExtEx(Now), 8))
   Vsp.CurrentY = dY(0)
   
   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   'Vsp.TextAlign = taCenterBottom
   'Vsp.Paragraph = MODULE_DESC
   'Vsp.Paragraph = ""
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
Private Sub GenerateDocHeader(Ind As Long, Vsp As VSPrinter, mcolParam As Collection, TempBorder As TableBorderSettings, HeaderOffset As Double)
      
   '============
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 1
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 3.5 * 567 + HeaderOffset
    Vsp.TableCell(tcColWidth, 1, 1) = "18.2cm"
    Vsp.TableBorder = TempBorder
    Vsp.EndTable
   
   '============
   
   '============
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 1 * 567
    Vsp.TableCell(tcColWidth, , 1) = "15.7cm"
    Vsp.TableCell(tcColWidth, , 2) = "2.5cm"
           
    Vsp.TableCell(tcAlign, 1, 2) = taCenterMiddle
    Vsp.TableCell(tcText, 1, 2) = mcolParam("JV_NO")
    
    Vsp.TableBorder = TempBorder
    Vsp.EndTable
   '============
   '

   '============
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 1.2 * 567
    Vsp.TableCell(tcColWidth, , 1) = "14cm"
    Vsp.TableCell(tcColWidth, , 2) = "4.2cm"
           
    Vsp.TableCell(tcAlign, 1, 2) = taCenterMiddle
    Vsp.TableCell(tcText, 1, 2) = DateToStringExtEx2(mcolParam("FROM_DATE"))
    
    Vsp.TableBorder = TempBorder
    Vsp.EndTable
   '============
   
   '============
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 1
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 1.8 * 567
    Vsp.TableCell(tcColWidth, 1, 1) = "18.2cm"
    Vsp.TableBorder = TempBorder
    Vsp.EndTable
   '============
End Sub

