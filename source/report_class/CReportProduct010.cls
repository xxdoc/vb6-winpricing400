VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportProduct010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportProduct011"
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

Private Const TITLE_SIZE = 11
Private Const HEADER_SIZE = 11
Private Const DETAIL_SIZE = 11

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

Private Total1(100) As Double
Private Total2(100) As Double
Private Total3(100) As Double

Private PageAll As Long
Private alngX As Long

Private RCount As Long
Private i, I2 As Long
Private iCount As Long
Private IsOK As Boolean

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
Dim TempStr2 As String
Dim Cf As CReportField
Dim jobOp As CJobOutputEn
Dim Rs2 As ADODB.Recordset
Dim assetStr As String
Dim TITLE_FORMAT As String
Dim COLLUM_FORMAT  As String
Dim AllTX As Double

TITLE_FORMAT = "<5|<5|<5|<5|<5;"
COLLUM_FORMAT = "<5|^5|>5|>5|<5;"
AllTX = 0

strFormat = VSP_CalTable(COLLUM_FORMAT, mdblWid, alngX)
Call Vsp.AddTable(strFormat, "", "ประเภทอาหารสัตว์ | จำนวนผลิต ( ตัน ) |จำนวนจำหน่าย ( ตัน ) | จำนวนเงินจำหน่าย( พันบาท )|")

Set Rs2 = New ADODB.Recordset
Set jobOp = New CJobOutputEn

Call jobOp.QueryData(1, Rs2, iCount)
While Not Rs2.EOF
 Call jobOp.PopulateFromRS(2, Rs2)
        AllTX = AllTX + jobOp.TX_AMOUNT
        strFormat = VSP_CalTable(COLLUM_FORMAT, mdblWid, alngX)
        Call Vsp.AddTable(strFormat, "", jobOp.PART_TYPE_NAME & " | " & jobOp.TX_AMOUNT & "|N/A|N/A")
        
    If Not Rs2.EOF Then Rs2.MoveNext
Wend

        strFormat = VSP_CalTable(COLLUM_FORMAT, mdblWid, alngX)
        Call Vsp.AddTable(strFormat, "", "รวม | " & AllTX & "|N/A|N/A")
        
If Rs2.State = adStateOpen Then
      Rs2.Close
   End If
   Set Rs2 = Nothing
strFormat = VSP_CalTable(COLLUM_FORMAT, mdblWid, alngX)
Call Vsp.AddTable("<10;", "", vbCrLf)
   
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   Call m_Headers1_2.ClearField
   Call m_Headers1_3.ClearField
   
   

   
   Set Cf = New CReportField

   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ชื่อวัตถุดิบ", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "จำนวน", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "หน่วย", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "มูลค่า", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "หมายเหตุ", ">")
   Call m_Headers1.AddField(Cf)

   
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
   
   
'   FromDate = "ณ วันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")
'   ToDate = "ถึง " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")
   
   MODULE_DESC = "รายงานสรุปยอดการผลิต" & vbCrLf & glbEnterPrise.ENTERPRISE_NAME & "   " & glbEnterPrise.BRANCH_NAME '& vbCrLf & _
                                       'mcolParam("REPORT_NAME") & vbCrLf & _
                                       'FromDate & " สถานที่จัดเก็บ " & EmptyToString(mcolParam("LOCATION_NAME"), "N/A")
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
Dim j As Long

Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Amt As Double
Dim Ug  As CUserGroup

Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim strPrint As String
Dim TempStr1 As String
Dim TempStr2 As String

Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim ID As Long
Dim RsLine As Integer
Dim E_NAME As String
Dim jobOp As CJobOutput

  
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------

   For j = 1 To UBound(Total1)
      Total1(j) = 0
      Total2(j) = 0
      Total3(j) = 0
   Next j
   
  
Set Rs = New ADODB.Recordset
Set jobOp = New CJobOutput

      jobOp.FROM_DATE = mcolParam("FROM_DATE")
      jobOp.TO_DATE = mcolParam("TO_DATE")
     jobOp.OrderBy = mcolParam("ORDER_BY")
     jobOp.OrderType = mcolParam("ORDER_TYPE")
      'cData.QueryFlag = -1
Call jobOp.QueryData(1, Rs, iCount)

PageAll = (iCount \ 25)                                                                 ' จำนวนหน้าทั้งหมด
If (iCount Mod 20) > 0 Then PageAll = PageAll + 1

Call StartExportFile(Vsp)
 Vsp.StartDoc
    i = 0
    I2 = 0
    
Set BodyCf = New CReportField

While Not Rs.EOF

         i = i + 1
         I2 = I2 + 1
         Call jobOp.PopulateFromRS(2, Rs)
         Call m_Details1.ClearField

         Set HeadCf = m_Headers1.Fields(1)
         TempStr = i
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(2)
         TempStr = jobOp.PART_DESC
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(3)
         TempStr = jobOp.TX_AMOUNT
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(4)
         TempStr = jobOp.UNIT_NAME
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
         'Total1(6) = Total1(6) + CDbl(Replace(FormatNumber(Ti.PAID_AMOUNT, "0"), ",", ""))

'         Set HeadCf = m_Headers1.Fields(7)
'         TempStr = ""
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'         Call m_Details1.AddField(BodyCf)
'         'Total1(7) = Total1(7) + CDbl(Replace(FormatNumber(Ti.WH_AMOUNT, "0"), ",", "")) 'Ti.WH_AMOUNT

'         Set HeadCf = m_Headers1.Fields(8)
'         TempStr = ""
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'         Call m_Details1.AddField(BodyCf)

         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)

        If (i Mod 25) = 0 Then
            If Rs.RecordCount - i > 0 Then    'Still left
               Vsp.NewPage
            End If
            I2 = 0
        End If
        If Not Rs.EOF Then Rs.MoveNext
Wend

 
   
If I2 < 7 Then
   'Call PatchTable(Vsp, strFormat, "|||||||||", 500)
   '===
    RCount = I2
     
    For I2 = 0 To 24 - RCount
         i = i + 1
         Call m_Details1.ClearField

         Set HeadCf = m_Headers1.Fields(1)
         TempStr = i
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
         'Total1(6) = Total1(6) + CDbl(Replace(FormatNumber(Ti.PAID_AMOUNT, "0"), ",", ""))

'         Set HeadCf = m_Headers1.Fields(7)
'         TempStr = ""
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'         Call m_Details1.AddField(BodyCf)
'         'Total1(7) = Total1(7) + CDbl(Replace(FormatNumber(Ti.WH_AMOUNT, "0"), ",", "")) 'Ti.WH_AMOUNT

'         Set HeadCf = m_Headers1.Fields(8)
'         TempStr = ""
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'         Call m_Details1.AddField(BodyCf)


         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
     
      Next I2
    
End If
      
    
                                                                                                                        'Footer Report

    If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing

   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", strPrint)
   Vsp.FontBold = False

   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
      
   genDoc = True
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
   Set Rs = Nothing
End Function

Private Sub GenerateFooter(TXT1 As String, TXT2 As String, TXT3 As String, Tot() As Double, Details As CFieldList, DateCount As Long)
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
         TempStr = TXT1
      ElseIf j = 2 Then
         TempStr = TXT2
      ElseIf j = 3 Then
         TempStr = TXT3
      ElseIf (j = 4) Or (j = 5) Then
         TempStr = ""
      ElseIf j = 6 Then
         'TempStr = Tot(j)
         TempStr = FormatNumber(Tot(j))
      ElseIf j = 7 Then
         TempStr = FormatNumber(Tot(j))
      ElseIf j = 8 Then
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

'Private Sub VSP_EndPage()
'Dim strFormat As String
'Dim strPrint As String
'Dim tbdT As TableBorderSettings
'Dim blnBold As Boolean
'Dim iSize As Integer
'Dim alngX() As Long
'Dim PrevID As Long
'Dim TempStr1 As String
'Dim TempStr1_1 As String
'Dim Cf As CReportField
'Dim ExportWidth As Long
'Dim HeadCf As CReportField
'Dim j As Long
'Dim TempStr As String
'Dim TempRs As ADODB.Recordset
'Dim ImportItems As Collection
'Dim Amt As Double
'Dim iCount As Long
'
'On Error Resume Next
'
'Call m_Headers1_3.ClearField
'Call m_Headers1_4.ClearField
'
'   tbdT = Vsp.TableBorder
'   blnBold = Vsp.FontBold
'   iSize = Vsp.FontSize
'
'   Vsp.FontSize = 12
'   Vsp.FontBold = True
'   Vsp.TableBorder = tbAll
'
'If Vsp.PageCount = PageAll Then
'    strFormat = VSP_CalTable("<", mdblWid, alngX)
'    Vsp.TableBorder = 0
'
'    strPrint = "" & vbCrLf
'   Call Vsp.AddTable(strFormat, "", strPrint)
'
'   strPrint = "คนงานฝ่ายผลิต : " & vbCrLf & "จำนวนช่วโมงการทำงาน :" & vbCrLf & "จำนวนวัน :"
'   Call Vsp.AddTable(strFormat, "", strPrint)
'   Vsp.TableBorder = tbAll
'   End If
'
'End Sub
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
Dim j As Long
Dim TempStr As String
Dim TempRs As ADODB.Recordset
Dim ImportItems As Collection
Dim Amt As Double
Dim iCount As Long
Dim jobOp As CJobOutputEn
Dim jobRS As CJobResource
Dim Rs2 As ADODB.Recordset
Dim assetStr As String
Dim TITLE_FORMAT As String
Dim COLLUM_FORMAT  As String

COLLUM_FORMAT = "<10;"


Set Rs2 = New ADODB.Recordset
Set jobRS = New CJobResource

 Vsp.TableBorder = 0
Call jobRS.QueryDataEmp(1, Rs2, iCount)

While Not Rs2.EOF
 Call jobRS.PopulateFromRS(3, Rs2)
strFormat = VSP_CalTable(COLLUM_FORMAT, mdblWid, alngX)
Call Vsp.AddTable(strFormat, "", "คนงานฝ่ายผลิต :       " & jobRS.PEOPLE & "      คน" & vbCrLf & "จำนวนชั่วโมงการทำงาน :      " & jobRS.OCCUPY_INTERVAL & "   ชั่วโมง :      " & vbCrLf & " จำนวนวัน :    " & (jobRS.FINISH_DATE) - (jobRS.START_DATE) & "    วัน ")

    If Not Rs2.EOF Then Rs2.MoveNext
Wend
If Rs2.State = adStateOpen Then
      Rs2.Close
   End If
   Set Rs2 = Nothing

    Vsp.TableBorder = tbAll
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

Private Sub PrintBillFooter(alngX() As Long)
Dim strFormat As String
Dim strPrint As String

   Vsp.TableBorder = tbNone
   Vsp.FontBold = False
   
   Vsp.FontBold = True
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = "RMC : " & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = "PMC : " & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   Vsp.FontBold = False
   
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("^10|^10;", mdblWid, alngX)
   strPrint = "ผู้สร้างสูตร .....................................................|ผู้อนุมัติ .....................................................;"
   Call Vsp.AddTable(strFormat, "", strPrint)
End Sub

'       jobOp.FROM_DATE = mcolParam("FROM_DATE")
'       jobOp.TO_DATE = mcolParam("TO_DATE")
'      jobOp.OrderBy = mcolParam("ORDER_BY")
'      jobOp.OrderType = mcolParam("ORDER_TYPE")
'     If mcolParam("SUCCESS_FLAG") = 1 Then
'         jobOp.SUCCESS_FLAG = "Y"
'    Else
'         jobOp.SUCCESS_FLAG = "N"
'     End If





