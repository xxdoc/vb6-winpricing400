VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportInventory023"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportInventory023"
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
Private m_RcpCndnItems As Collection
Private m_RcpCndnBals As Collection
Private m_InvoiceBals As Collection
Private m_PartItems As Collection
Private m_PartItemTxs As Collection

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
   Set m_RcpCndnItems = New Collection
   Set m_RcpCndnBals = New Collection
   Set m_InvoiceBals = New Collection
   Set m_PartItems = New Collection
   Set m_PartItemTxs = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_RcpCndnItems = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
   Set m_RcpCndnBals = Nothing
   Set m_InvoiceBals = Nothing
   Set m_PartItems = Nothing
   Set m_PartItemTxs = Nothing
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

   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "วันที่เอกสาร", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "เลขที่เอกสาร", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "รับเข้า", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "จ่ายออก", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "คงเหลือ", ">")
   Call m_Headers0.AddField(Cf)

   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "วันที่เอกสาร", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "เลขที่เอกสาร", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ราคา", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ราคา", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ราคา", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
   Call m_Headers1.AddField(Cf)

   Set Cf = Nothing
   
  Call m_Headers0.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr2
   Call Vsp.AddTable(strFormat, "", strPrint)
            
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
Dim Li As CLotItem
Dim NewDate As Date

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
   
   Call LoadDistinctPartItemTx(Nothing, m_PartItemTxs, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), "", mcolParam("LOCATION_ID"), mcolParam("PART_TYPE"), mcolParam("PART_NO"))
   
   Call LoadDistinctPartItemTx(Nothing, m_PartItems, , , "", mcolParam("LOCATION_ID"), mcolParam("PART_TYPE"), mcolParam("PART_NO"))
      
   FromDate = "จาก " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
   ToDate = " ถึง " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")
                                          
   MODULE_DESC = vbCrLf & glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & _
                                       vbCrLf & mcolParam("REPORT_NAME") & _
                                       vbCrLf & FromDate & " - " & ToDate & _
                                       vbCrLf & "สถานที่จัดเก็บ " & EmptyToString(mcolParam("LOCATION_NAME"), "N/A")
   
   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))
   
   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
      mstrErrMsg = "Error GetDateTime Error."
      Exit Function
   End If
   mdteDate = InternalDateToDate(strDate)
   initDoc = True
End Function

Private Sub GenerateSubHeader(TXT1 As String, TXT2 As String, TXT3 As String, Tot() As Double, Details As CFieldList, DateCount As Long)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim J As Long
Dim TempStr As String
Dim TempWidth As Double

   Call Details.ClearField
   Set BodyCf = New CReportField

   TempStr = TXT1
   TempWidth = m_Headers1.Fields(1).Width + m_Headers1.Fields(2).Width
   Call BodyCf.SetFieldValue(TempWidth, "^", TempStr, "^")
   Call Details.AddField(BodyCf)

   TempStr = TXT2
   TempWidth = 0
   For J = 3 To m_Headers1.Fields.Count
      TempWidth = TempWidth + m_Headers1.Fields(J).Width
   Next J
   Call BodyCf.SetFieldValue(TempWidth, "<", TempStr, "^")
   Call Details.AddField(BodyCf)

   Set BodyCf = Nothing
End Sub

Private Function GetBalance(CustID As Long, Col As Collection) As CLotItem
   Set GetBalance = GetLotItem(Col, mcolParam("LOCATION_ID") & "-" & CustID)
End Function

Private Function IsIn(TempID As Long, Col As Collection) As Boolean
Dim Li As CLotItem

   IsIn = False
   For Each Li In Col
      If Li.PART_ITEM_ID = TempID Then
         IsIn = True
         Exit Function
      End If
   Next Li
'IsIn = True
End Function

Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim RName As String
Dim I As Long
Dim J As Long
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
Dim OldID As Long
Dim Li As CLotItem
Dim TempLi As CLotItem
Dim k As Long

Dim BalanceAccums As Collection
Dim InventoryBals2 As Collection
Dim BalLi As CLotItem
Dim ShowFlag As Boolean

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set Rs = New ADODB.Recordset
   Set InventoryBals2 = New Collection
   Set BalanceAccums = New Collection
   
   If mcolParam("FROM_DATE") > 0 Then
      Dim MonthlyAccums  As Collection
      Dim YYYYMM As String
      Set MonthlyAccums = New Collection
      YYYYMM = Format(Year(DateAdd("D", -1, mcolParam("FROM_DATE"))), "0000") & "-" & Format(Month(DateAdd("D", -1, mcolParam("FROM_DATE"))), "00")
      Call LoadMonthlyBalance(Nothing, MonthlyAccums, YYYYMM, mcolParam("LOCATION_ID"))
      Call glbDaily.CopyMonthlyAccum(MonthlyAccums, InventoryBals2)
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

   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
   
      I = 0
      k = 0
      Set BodyCf = New CReportField

      Set Li = New CLotItem

      Li.LOT_ITEM_ID = -1
      Li.FROM_DATE = mcolParam("FROM_DATE")
      Li.TO_DATE = mcolParam("TO_DATE")
      Li.LOCATION_ID = mcolParam("LOCATION_ID")
      Li.PART_TYPE = mcolParam("PART_TYPE")
      Li.PART_GROUP_ID = mcolParam("PART_GROUP")
      Li.PART_NO = mcolParam("PART_NO")
      Li.OrderBy = 14
      Li.OrderType = 1
       Call Li.QueryData(1, Rs, iCount, False)
      
      For Each TempLi In m_PartItems
         For J = 1 To UBound(Total1)
            Total1(J) = 0
         Next J
         
         Set BalLi = GetBalance(TempLi.PART_ITEM_ID, InventoryBals2)
         ShowFlag = False
         If (IsIn(TempLi.PART_ITEM_ID, m_PartItemTxs)) Or (BalLi.NEW_AMOUNT <> 0) Then
            ShowFlag = True
         End If
         
         If ShowFlag Then
            Call GenerateSubHeader(TempLi.PART_TYPE_NAME & " (" & TempLi.PART_TYPE_NO & ")", TempLi.PART_DESC & " (" & TempLi.PART_NO & ")", "", Total1, m_Details1, 0)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
            
            Call GenerateBalance("ยกมา", "", "", BalLi, m_Details1, 0)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
            
            Total1(10) = BalLi.NEW_AMOUNT
            Total1(12) = BalLi.TOTAL_INCLUDE_PRICE
         End If

         If IsIn(TempLi.PART_ITEM_ID, m_PartItemTxs) Then
            If Not Rs.EOF Then
               Call Li.PopulateFromRS(1, Rs)
            End If
            While (Not Rs.EOF) And (TempLi.PART_NO = Li.PART_NO)
               I = I + 1

               '===
               Call m_Details1.ClearField
               For k = 1 To m_Headers1.Fields.Count
                  Set HeadCf = m_Headers1.Fields(k)
                  If k = 1 Then
                     TempStr = I
                  ElseIf k = 2 Then
                     TempStr = DateToStringExtEx2(Li.DOCUMENT_DATE)
                  ElseIf k = 3 Then
                     TempStr = Li.DOCUMENT_NO
                  ElseIf k = 4 Then
                     If Li.TX_TYPE = "I" Then
                        Amt = Li.TX_AMOUNT
                     Else
                        Amt = 0
                     End If
                     TempStr = FormatNumber(Amt)
                     Total1(k) = Total1(k) + Amt
                     Total2(k) = Total2(k) + Amt
                  ElseIf k = 5 Then
                     If Li.TX_TYPE = "I" Then
                        Amt = Li.INCLUDE_UNIT_PRICE
                     Else
                        Amt = 0
                     End If
                     TempStr = FormatNumber(Amt)
                     Total1(k) = Total1(k) + Amt
                     Total2(k) = Total2(k) + Amt
                  ElseIf k = 6 Then
                     If Li.TX_TYPE = "I" Then
                        Amt = Li.TOTAL_INCLUDE_PRICE
                     Else
                        Amt = 0
                     End If
                     TempStr = FormatNumber(Amt)
                     Total1(k) = Total1(k) + Amt
                     Total2(k) = Total2(k) + Amt
                  ElseIf k = 7 Then
                     If Li.TX_TYPE = "E" Then
                        Amt = Li.TX_AMOUNT
                     Else
                        Amt = 0
                     End If
                     TempStr = FormatNumber(Amt)
                     Total1(k) = Total1(k) + Amt
                     Total2(k) = Total2(k) + Amt
                  ElseIf k = 8 Then
                     If Li.TX_TYPE = "E" Then
                        Amt = Li.PREVIOUS_PRICE
                     Else
                        Amt = 0
                     End If
                     TempStr = FormatNumber(Amt)
                     Total1(k) = Total1(k) + Amt
                     Total2(k) = Total2(k) + Amt
                  ElseIf k = 9 Then
                     If Li.TX_TYPE = "E" Then
                        Amt = Li.TOTAL_INCLUDE_PRICE  'Li.TX_AMOUNT * Li.PREVIOUS_PRICE
                     Else
                        Amt = 0
                     End If
                     TempStr = FormatNumber(Amt)
                     Total1(k) = Total1(k) + Amt
                     Total2(k) = Total2(k) + Amt
                  ElseIf k = 10 Then
                     Amt = Li.NEW_AMOUNT
                     TempStr = FormatNumber(Amt)
                     Total1(k) = Amt
                  ElseIf k = 11 Then
                     Amt = Li.NEW_PRICE
                     TempStr = FormatNumber(Amt)
                     Total1(k) = Amt
                  ElseIf k = 12 Then
                     If Li.TOTAL_NEW_PRICE <= 0 Then
                        Amt = Li.NEW_AMOUNT * Li.NEW_PRICE
                     Else
                        Amt = Li.TOTAL_NEW_PRICE
                     End If
                     TempStr = FormatNumber(Amt)
                     Total1(k) = Amt
                  Else
                     TempStr = ""
                     Amt = 0
                  End If
                  
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               Next k
               
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
               '===

               Call Rs.MoveNext
               Call Li.PopulateFromRS(1, Rs)
            Wend
            
            Total2(10) = Total2(10) + Total1(10)
            Total2(12) = Total2(12) + Total1(12)
         End If

'         Total1(11) = CurAmount
'         Total2(11) = Total2(11) + CurAmount

         If ShowFlag Then
            Call GenerateFooter("รวม", "", "", Total1, m_Details1)
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
         End If
      Next TempLi

      Call GenerateFooter("สุทธิ", "", "", Total2, m_Details1)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False

      Set Li = Nothing
      Set BodyCf = Nothing
      Vsp.EndDoc
      Call CloseExportFile(Vsp)
      If Rs.State = adStateOpen Then
         Rs.Close
      End If
      Set Rs = Nothing
      Set InventoryBals2 = Nothing
      
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
      ElseIf J = 2 Then
         TempStr = TXT2
      ElseIf J = 3 Then
         TempStr = TXT3
      ElseIf J = 4 Then
         TempStr = FormatNumber(Tot(J))
      ElseIf J = 6 Then
         TempStr = FormatNumber(Tot(J))
      ElseIf J = 7 Then
         TempStr = FormatNumber(Tot(J))
      ElseIf J = 9 Then
         TempStr = FormatNumber(Tot(J))
      ElseIf J = 10 Then
         TempStr = FormatNumber(Tot(J))
      ElseIf J = 12 Then
         TempStr = FormatNumber(Tot(J))
      Else
         TempStr = ""
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub

Private Sub GenerateBalance(TXT1 As String, TXT2 As String, TXT3 As String, Bal As CLotItem, Details As CFieldList, DateCount As Long)
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
         TempStr = ""
      ElseIf J = 4 Then
         TempStr = FormatNumber(0)
      ElseIf J = 5 Then
         TempStr = FormatNumber(0)
      ElseIf J = 6 Then
         TempStr = FormatNumber(0)
      ElseIf J = 7 Then
         TempStr = FormatNumber(0)
      ElseIf J = 8 Then
         TempStr = FormatNumber(0)
      ElseIf J = 9 Then
         TempStr = FormatNumber(0)
      ElseIf J = 10 Then
         TempStr = FormatNumber(Bal.NEW_AMOUNT)
      ElseIf J = 11 Then
         TempStr = FormatNumber(Bal.NEW_PRICE)
      ElseIf J = 12 Then
         TempStr = FormatNumber(Bal.TOTAL_INCLUDE_PRICE)
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
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
   Call Vsp.AddTable(strFormat, "", "|วัน|: " & DateToStringExtEx2(mdteDate))
   Call Vsp.AddTable(strFormat, "", "|เวลา|: " & Right$(DateToStringExtEx2(mdteDate), 8))
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



