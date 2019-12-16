VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportProduct009"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportProduct009"
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
Private Const DETAIL_SIZE = 13                                       '  ปรับจากขนาดปกติ ของเดิม 12

Private Const TITLE_FORMAT = "<32|<3|<5;"

Private m_TempCol As Collection
Private m_HeaderFormat1 As String
Private m_HeaderFormat2 As String
Private m_HeaderText1 As String
Private m_HeaderText2 As String

Private m_Headers0 As CFieldList
Private m_Headers1 As CFieldList
Private m_Headers1_1 As CFieldList
Private m_Details1 As CFieldList
Private m_Left As Double
Private m_YearSeqs As Collection

Private m_Headers1_3 As CFieldList
Private m_Headers1_4 As CFieldList
Private PageAll As Long



Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_TempCol = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Headers1_1 = New CFieldList
   Set m_Details1 = New CFieldList
   Set m_YearSeqs = New Collection
   
   Set Rs = New ADODB.Recordset
   
   Set m_Headers1_3 = New CFieldList
   Set m_Headers1_4 = New CFieldList
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Details1 = Nothing
   Set m_YearSeqs = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   
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
   Call m_Headers1_1.ClearField
   
    
   
   Set Cf = New CReportField
   
   Call Cf.SetFieldValue(1, "<", "....กะ A          ....กะ B              ....กะ C", "<")
   Call m_Headers1_1.AddField(Cf)
   
   Set Cf = Nothing
   
  Call m_Headers1_1.GetString(1, TempStr1, TempStr2)
   Vsp.TableBorder = 0
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr2
   Call Vsp.AddTable(strFormat, "", strPrint)
    Vsp.TableBorder = tbAll


 Set Cf = New CReportField
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "รหัสวัตถุดิบ ( ยา )", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "ชื่อวัตถุดิบ ( ยา )", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "ประเภทวัตถุดิบ", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "หน่วยนับ", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "วันที่ / เวลา / เทยา", "^")
   Call m_Headers1.AddField(Cf)
    Call Cf.SetFieldValue(2, "^", "จากสูตรอาหาร", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "หมายเหตุ", "<")
   Call m_Headers1.AddField(Cf)
   
   Set Cf = Nothing
   
  Call m_Headers1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr2
   Call Vsp.AddTable(strFormat, "", strPrint)
            
   Vsp.TableBorder = tbdT
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
   
   Call Vsp.DrawRectangle(100, 100, 10, 10)
   
End Sub

Private Function initDoc() As Boolean
Dim strDate As String
Dim TempStr1 As String
Dim TempStr2 As String
Dim TempStr3 As String
Dim FromDate As String
Dim ToDate As Date
   
   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprFanfoldUS 'pprA4 '
   Vsp.ORIENTATION = orLandscape
   Vsp.MarginBottom = 1440
   Vsp.MarginFooter = 1440
   Vsp.MarginHeader = 1440
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 1440
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   ToDate = mcolParam("TO_DATE")
   FromDate = "    ประจำวันที่  " & EmptyToString(DateToStringExt(mcolParam("TO_DATE")), "N/A")

   MODULE_DESC = vbCrLf & glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & vbCrLf & _
                                       mcolParam("REPORT_NAME") & FromDate
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
Dim cData As CPartItem
Dim I, I2 As Long
Dim J As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Pi As CPartItem
Dim Total1(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim Amt As Double
'Dim InventoryBals1 As Collection
'Dim InventoryBals2 As Collection
Dim Li1 As CLotItem
Dim Li2 As CLotItem
Dim TempLi As CLotItem
Dim Sum1 As Double
Dim Count1 As Double

Dim RCount As String

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------

'   Set InventoryBals1 = New Collection
'   Call LoadLeftAmount1(Nothing, InventoryBals1, -1, mcolParam("TO_DATE"), "Y", mcolParam("LOCATION_ID"))
'
'   Set InventoryBals2 = New Collection
'   Call LoadLeftAmount2(Nothing, InventoryBals2, -1, mcolParam("TO_DATE"), "", mcolParam("LOCATION_ID"))
   
   If Not initDoc Then Exit Function

   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------

   For J = 1 To UBound(Total1)
      Total1(J) = 0
   Next J

   Set cData = New CPartItem
   cData.PART_ITEM_ID = -1
   cData.PART_NO = mcolParam("PART_NO")
   cData.PART_DESC = mcolParam("PART_DESC")
   cData.UNIT_COUNT = mcolParam("UNIT_COUNT")
   cData.PART_TYPE = mcolParam("PART_TYPE")
   cData.LOCATION_ID = mcolParam("LOCATION_ID")
   cData.TO_DATE = mcolParam("TO_DATE")
   cData.OrderBy = mcolParam("ORDER_BY")
   cData.OrderType = mcolParam("ORDER_TYPE")
  cData.PART_TYPE_NO = "07"                            '    วัตถุดิบประเถทยา
 
   Call cData.QueryData(3, Rs, iCount)
   If (Rs Is Nothing) Or (iCount < 1) Then
      mstrErrMsg = "ไม่พบข้อมูลที่ต้องการในฐานข้อมูล"
      Exit Function
    End If
    
    PageAll = (Rs.RecordCount \ 13)                                                                  '       จำนวนหน้าทั้งหมด
    If (Rs.RecordCount Mod 13) > 0 Then PageAll = PageAll + 1
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
   
      I = 0
      
      Set Pi = New CPartItem
      Set BodyCf = New CReportField
      
      RCount = I2
      While Not Rs.EOF
         I = I + 1
         I2 = I2 + 1
         Call Pi.PopulateFromRS(3, Rs)
          Call m_Details1.ClearField
          
'          Set Li1 = GetLotItem(InventoryBals1, Trim(Str(Pi.PART_ITEM_ID)))
'          If mcolParam("LOCATION_ID") > 0 Then
'            Set Li2 = GetLotItem(InventoryBals2, mcolParam("LOCATION_ID") & "-" & Pi.PART_ITEM_ID)
'          Else
'            Sum1 = 0
'            Count1 = 0
'            For Each TempLi In InventoryBals2
'               If Pi.PART_ITEM_ID = TempLi.PART_ITEM_ID Then
'                  Count1 = Count1 + TempLi.NEW_AMOUNT
'                  Sum1 = Sum1 + TempLi.NEW_AMOUNT * TempLi.NEW_PRICE
'               End If
'               Set Li2 = Nothing
'               Set Li2 = New CLotItem
'               Li2.NEW_PRICE = MyDiff(Sum1, Count1)
'            Next TempLi
'          End If
         '=====
         Set HeadCf = m_Headers1.Fields(1)
         TempStr = I
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(2)
         TempStr = Pi.PART_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
          
         Set HeadCf = m_Headers1.Fields(3)
         TempStr = Pi.PART_DESC
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(4)
         TempStr = Pi.PART_TYPE_NAME & " (" & Pi.PART_TYPE_NO & ")"
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(5)
         Amt = Pi.TX_AMOUNT
         TempStr = FormatNumber(Amt)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(6)
         TempStr = Pi.UNIT_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         'Total1(6) = Total1(6) + Amt

         Set HeadCf = m_Headers1.Fields(7)
          TempStr = DateToStringExtEx3(Pi.MODIFY_DATE)                                                                               ' วันที่เทยา
         'TempStr = FormatNumber(Amt)
         
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         'Total1(7) = Total1(7) + Amt
         
         Set HeadCf = m_Headers1.Fields(8)                                                                      ' ใบสั่งผลิต
        TempStr = Pi.JOB_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         'Total1(8) = Total1(8) + Amt
         
         Set HeadCf = m_Headers1.Fields(9)                                                                      ' หมายเหตุ
        TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         '=====
         
         If (I Mod 16) = 0 Then
            If Rs.RecordCount - I > 0 Then   'Still left
                Vsp.NewPage
            End If
          I2 = 0
        End If
         If Not Rs.EOF Then Rs.MoveNext

    Wend

 
    
    If I2 > 0 Then
        RCount = I2
      For I2 = 0 To 16 - RCount
         I = I + 1
         Call m_Details1.ClearField

         Set HeadCf = m_Headers1.Fields(1)
         TempStr = I
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(2)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(3)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(4)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(5)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(6)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         

         Set HeadCf = m_Headers1.Fields(7)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(8)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
        Set HeadCf = m_Headers1.Fields(9)
        TempStr = ""
        Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
        Call m_Details1.AddField(BodyCf)

         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
     
   Next I2
Else
End If
                                                                            
                                                                                                 'Footer Report

      Set BodyCf = Nothing
      Set Pi = Nothing
      
   Vsp.EndDoc
   Call CloseExportFile(Vsp)

   Set cData = Nothing
'   Set InventoryBals1 = Nothing
   
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
      ElseIf J = 6 Then
         TempStr = FormatNumber(Tot(J))
      ElseIf J = 8 Then
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
Dim ImportItems As Collection
Dim Amt As Double
Dim iCount As Long

On Error Resume Next

Call m_Headers1_3.ClearField
Call m_Headers1_4.ClearField
   
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = 12
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
 
If (Vsp.PageCount = PageAll) Or (PageAll = 0) Then
    strFormat = VSP_CalTable("^10|^10;", mdblWid, alngX)
    Vsp.TableBorder = 0
    
    strPrint = "" & vbCrLf
   Call Vsp.AddTable(strFormat, "", strPrint)
   
    strPrint = "ลงชื่อ .....................................................ผู้เทยา | ลงชื่อ.....................................................หัวหน้ากะ;"
    Call Vsp.AddTable(strFormat, "", strPrint)
    strPrint = "(...................................................................)|(...........................................................................);"
    Call Vsp.AddTable(strFormat, "", strPrint)
   
    Vsp.TableBorder = tbAll
   End If
   
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
















