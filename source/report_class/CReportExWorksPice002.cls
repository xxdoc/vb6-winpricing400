VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportExWorksPice002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportExWorksPice002"
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
   
   If mcolParam("FOR_IMPORT_FLAG") = 1 Then
      Call Cf.SetFieldValue(5, "^", "รหัสลูกค้า", "<")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(14, "^", "ชื่อลูกค้า", "<")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(5, "^", "รหัสสถานที", "<")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(14, "^", "สถานทีจัดส่ง", "<")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(6, "^", "ค่าขนส่ง", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(3, "^", "หน่วย", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(3, "^", "น้ำหนัก", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(4, "^", "คิดลูกค้า", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(3, "^", "หน่วย", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(3, "^", "น้ำหนัก", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(3, "^", "ใช้งาน", ">")
      Call m_Headers1.AddField(Cf)
      Set Cf = Nothing
   Else
      Call Cf.SetFieldValue(3, "^", "NO.", "^")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(5, "^", "รหัสลูกค้า", "<")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(14, "^", "ชื่อลูกค้า", "<")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(5, "^", "รหัสสถานที", "<")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(14, "^", "สถานทีจัดส่ง", "<")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(6, "^", "ค่าขนส่ง", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(3, "^", "หน่วย", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "กก.", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(4, "^", "ประเภท", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(4, "^", "คิดลูกค้า", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(3, "^", "หน่วย", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "กก.", ">")
      Call m_Headers1.AddField(Cf)

   End If
            
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
      
   MODULE_DESC = vbCrLf & glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & vbCrLf & _
                                       mcolParam("REPORT_NAME")
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
Dim Amt As Double
Dim EWP As CExWorksPrice
Dim Cm As CFeature
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim Total1(100) As Double
Dim Total2(100) As Double

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

      mblnHeader = False
      Call StartExportFile(Vsp)
      Vsp.StartDoc
      Call printHeader
      mblnHeader = True
      I = 0
      Set BodyCf = New CReportField

      Set EWP = New CExWorksPrice
      EWP.EX_WORKS_PRICE_ID = -1
      EWP.EX_WORKS_PRICE_DATE = mcolParam("DOC_DATE")
      EWP.EX_WORKS_PRICE_CODE = mcolParam("WORKS_PRICE_CODE")
      EWP.BETWEEN_DATE = mcolParam("BETWEEN_DATE")
'      EWP.PART_NO = mcolParam("PART_NO")
      EWP.EX_WORKS_PRICE_TYPE = 2

      EWP.OrderBy = mcolParam("ORDER_BY")
      EWP.OrderType = mcolParam("ORDER_TYPE")
      EWP.QueryFlag = 1
      Call EWP.QueryData(3, Rs, iCount)
      
      I = 0
      k = 0
      PrevKey1 = ""

      While Not Rs.EOF
         Call m_Details1.ClearField
         Call EWP.PopulateFromRS(3, Rs)
         
         If PrevKey1 <> EWP.EX_WORKS_PRICE_CODE And I > 0 Then
         'สรุป
'            Call GenerateFooter("รวม", "", "", Total1, m_Details1, 0)
'            Call m_Details1.GetString(2, TempStr1, TempStr2)
'            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'            Vsp.FontBold = True
'            Call Vsp.AddTable(strFormat, "", TempStr2)
'            Vsp.FontBold = False
'
            For J = 1 To UBound(Total1)
               Total1(J) = 0
            Next J
               
            Vsp = " แพคแกจ : " & EWP.EX_WORKS_PRICE_CODE & " วันที่ประกาศ " & EWP.EX_WORKS_PRICE_DATE & " มีผลระหว่างวันที่ " & EWP.FROM_ACTIVE_DATE & " - " & EWP.TO_VALID_DATE
            k = 0
         ElseIf I = 0 Then
            Vsp = " แพคแกจ : " & EWP.EX_WORKS_PRICE_CODE & " วันที่ประกาศ " & EWP.EX_WORKS_PRICE_DATE & " มีผลระหว่างวันที่ " & EWP.FROM_ACTIVE_DATE & " - " & EWP.TO_VALID_DATE
         End If
         
         Call m_Details1.ClearField
         PrevKey1 = EWP.EX_WORKS_PRICE_CODE
         I = I + 1
         k = k + 1
         J = 0
         
         If mcolParam("FOR_IMPORT_FLAG") = 1 Then

            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = EWP.CUSTOMER_CODE
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = EWP.CUSTOMER_NAME
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = EWP.DELIVERY_CUS_ITEM_CODE
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(4) = Total1(4) + EWP.PACKAGE_RATE
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = EWP.DELIVERY_CUS_ITEM_NAME
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = FormatNumber(EWP.RATE_DELIVERY, 3)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = EWP.RATE_TYPE
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = EWP.WEIGHT_PER_PACK
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = FormatNumber(EWP.RATE_CUSTOMER, 3)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = EWP.RATE_TYPE_CUS
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = EWP.WEIGHT_PER_PACK_CUS
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
         
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = "1"
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)

         
         Else
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = k
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
   
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = EWP.CUSTOMER_CODE
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
   
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = EWP.CUSTOMER_NAME
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
   
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = EWP.DELIVERY_CUS_ITEM_CODE
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(4) = Total1(4) + EWP.PACKAGE_RATE
   
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = EWP.DELIVERY_CUS_ITEM_NAME
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = FormatNumber(EWP.RATE_DELIVERY, 3)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = DeliveryUnit(EWP.RATE_TYPE)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = EWP.WEIGHT_PER_PACK
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
             J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            If EWP.RATE_TYPE = 1 Then
               TempStr = "BAG"
            ElseIf EWP.RATE_TYPE = 2 Then
               TempStr = "BULK"
            ElseIf EWP.RATE_TYPE = 3 Then
               TempStr = "เหมาเที่ยว"
            End If
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = FormatNumber(EWP.RATE_CUSTOMER, 3)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = DeliveryUnit(EWP.RATE_TYPE_CUS)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = EWP.WEIGHT_PER_PACK_CUS
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            

         
      End If
         
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Rs.MoveNext
      Wend

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
      ElseIf (J = 4) Then
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


