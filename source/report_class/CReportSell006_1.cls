VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportSell006_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportSell006_1"
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
Private m_Features As Collection
Private m_PartItems As Collection
Private m_Credit As Collection

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
   Set m_Features = New Collection
   Set m_PartItems = New Collection
   Set m_Credit = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_Features = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
   Set m_PartItems = Nothing
   Set m_Credit = Nothing
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
   
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "�����١���", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(28, "^", "", "<")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "�����١���", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(8, "^", "�����١���", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�ӹǹ�ا", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "���˹ѡ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(0, "^", "�Ҥ�/˹���", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "��Ť��", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "��ǹŴ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "��ǹŴ����", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "�ط��", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�Ҥ�/�.�.", ">")
   Call m_Headers1.AddField(Cf)
   If mcolParam("CREDIT") = 1 Then
      Call Cf.SetFieldValue(3, "^", "Ŵ˹��/����˹��", ">")
      Call m_Headers1.AddField(Cf)
   End If
   Set Cf = Nothing
            
'  Call m_Headers0.GetString(1, TempStr1, TempStr1_1)
'   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'   strPrint = TempStr1_1
'   Call Vsp.AddTable(strFormat, "", strPrint)
   
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
   
   FromDate = "�ҡ�ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
   ToDate = "�֧ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")
   
   Call LoadCredit(Nothing, m_Credit, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), mcolParam("CUSTOMER_CODE"), mcolParam("FROM_CUSTOMER_CODE"), mcolParam("TO_CUSTOMER_CODE"))
   
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

Private Function GetKey(Di As CDoItem) As String
Dim OrderBy As Long

   OrderBy = mcolParam("ORDER_BY")
   
   If OrderBy = 1 Then
      GetKey = Di.CUSTOMER_ID
   ElseIf OrderBy = 2 Then
      GetKey = Di.CUSTOMER_ID
   Else
      GetKey = Di.CUSTOMER_ID
   End If
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
Dim Di As CDoItem
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(100) As Double
Dim Total2(100) As Double
Dim TempData As CReceiptItem
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim PrevKey3 As String
Dim Cm As CCustomer
Dim Pi As CPartItem
Dim Ft As CFeature
Dim PrevCode As String
Dim PrevName As String
Dim IncDiscount As Double
Dim SumCredit As Double

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set Rs = New ADODB.Recordset
'
'   Call LoadPartItem(Nothing, m_PartItems, mcolParam("PART_TYPE"))
   Call LoadFeature(Nothing, m_Features)
   
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

      Set Di = New CDoItem
      Di.DO_ITEM_ID = -1
      Di.CUSTOMER_CODE = mcolParam("CUSTOMER_CODE")
      Di.FROM_CUSTOMER_CODE = mcolParam("FROM_CUSTOMER_CODE")
      Di.TO_CUSTOMER_CODE = mcolParam("TO_CUSTOMER_CODE")
      Di.PART_NO = mcolParam("PART_NO")
      Di.FROM_DATE = mcolParam("FROM_DATE")
      Di.TO_DATE = mcolParam("TO_DATE")
      Di.CUSTOMER_TYPE = mcolParam("CUSTOMER_TYPE")
      Di.PART_TYPE = mcolParam("PART_TYPE")
      Di.DocTypeSet = BillingDocType2Set(mcolParam("BILL_SUBTYPE"))
      Di.OrderBy = mcolParam("ORDER_BY")
      Di.OrderType = mcolParam("ORDER_TYPE")
      Call Di.QueryData(7, Rs, iCount)

      I = 0
      PrevKey1 = ""
      If Not Rs.EOF Then
         Call Di.PopulateFromRS(7, Rs)
         PrevKey1 = GetKey(Di)
         PrevCode = Di.CUSTOMER_CODE
         PrevName = Di.CUSTOMER_NAME
         
         Call GenerateHeader("", Di.CUSTOMER_CODE, Di.CUSTOMER_NAME, m_Details1)
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
'         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
      End If

      While Not Rs.EOF
         PrevCode = Di.CUSTOMER_CODE
         PrevName = Di.CUSTOMER_NAME
         Call Di.PopulateFromRS(7, Rs)
         If Di.PARCEL_TYPE = 2 Then
            Di.PACK_AMOUNT = 0
         End If
         
         If PrevKey1 <> GetKey(Di) Then
            k = k + 1
            Call GenerateFooter(Trim(str(k)), PrevCode, PrevName, Total1, m_Details1, 0)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
'            Vsp.FontBold = False

            Call GenerateHeader("", Di.CUSTOMER_CODE, Di.CUSTOMER_NAME, m_Details1)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
'            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False

            For J = 1 To UBound(Total1)
               Total1(J) = 0
            Next J
         End If
         PrevKey1 = GetKey(Di)

         Call m_Details1.ClearField
'         Set Pi = GetPartItem(m_PartItems, Trim(Str(Di.PART_ITEM_ID)))
'         Set Ft = GetFeature(m_Features, Trim(Str(Di.FEATURE_ID)))

         I = I + 1
         Set HeadCf = m_Headers1.Fields(1)
         TempStr = I
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         If Di.PART_ITEM_ID > 0 Then
            TempStr1 = Di.PART_NO   'Pi.PART_NO
            TempStr2 = Di.PART_DESC   'Pi.PART_DESC
         ElseIf Di.FEATURE_ID > 0 Then
            TempStr1 = Di.FEATURE_CODE  'Ft.FEATURE_CODE
            TempStr2 = Di.FEATURE_DESC  'Ft.FEATURE_DESC
            Di.PACK_AMOUNT = 0
            Di.ITEM_AMOUNT = 0
         End If

         Set HeadCf = m_Headers1.Fields(2)
         TempStr = Di.DOCUMENT_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(3)
         TempStr = ID2PackageType(Di.PACKAGE_TYPE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(4)
         TempStr = FormatNumber(Di.PACK_AMOUNT, 0)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(4) = Total1(4) + (Di.PACK_AMOUNT)
         Total2(4) = Total2(4) + (Di.PACK_AMOUNT)

         Set HeadCf = m_Headers1.Fields(5)
         TempStr = FormatNumber(Di.ITEM_AMOUNT, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(5) = Total1(5) + (Di.ITEM_AMOUNT)
         Total2(5) = Total2(5) + (Di.ITEM_AMOUNT)

         Set HeadCf = m_Headers1.Fields(6)
         TempStr = FormatNumber(Di.AVG_PRICE, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(6) = Total1(6) + (0)
         Total2(6) = Total2(6) + (0)

         Set HeadCf = m_Headers1.Fields(7)
         TempStr = FormatNumber(Di.TOTAL_PRICE + Di.DISCOUNT_AMOUNT)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(7) = Total1(7) + (Di.TOTAL_PRICE + Di.DISCOUNT_AMOUNT)
         Total2(7) = Total2(7) + (Di.TOTAL_PRICE + Di.DISCOUNT_AMOUNT)

         Set HeadCf = m_Headers1.Fields(8)
         TempStr = FormatNumber(Di.DISCOUNT_AMOUNT, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(8) = Total1(8) + Di.DISCOUNT_AMOUNT
         Total2(8) = Total2(8) + Di.DISCOUNT_AMOUNT

         Set HeadCf = m_Headers1.Fields(9)
         TempStr = FormatNumber(Di.EXTRA_DISCOUNT)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(9) = Total1(9) + Di.EXTRA_DISCOUNT
         Total2(9) = Total2(9) + Di.EXTRA_DISCOUNT

         IncDiscount = Di.TOTAL_PRICE - Di.EXTRA_DISCOUNT
         Set HeadCf = m_Headers1.Fields(10)
         TempStr = FormatNumber(IncDiscount)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(10) = Total1(10) + (IncDiscount)
         Total2(10) = Total2(10) + (IncDiscount)

         Set HeadCf = m_Headers1.Fields(11)
         TempStr = FormatNumber(MyDiffEx(IncDiscount, Di.ITEM_AMOUNT))
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(11) = Total1(11) + 0
         Total2(11) = Total2(11) + 0
         
         If mcolParam("CREDIT") = 1 Then
            SumCredit = 0
            
            Rs.MoveNext
            Call Di.PopulateFromRS(7, Rs)
            PrevKey2 = Di.CUSTOMER_CODE
            Rs.MovePrevious
            Call Di.PopulateFromRS(7, Rs)
            PrevKey3 = Di.CUSTOMER_CODE
            
            If PrevKey2 <> PrevKey3 Then
               For Each TempData In m_Credit
                  If Di.CUSTOMER_CODE = TempData.CUSTOMER_CODE Then
                     If TempData.DOCUMENT_TYPE = 3 Then
                        SumCredit = SumCredit - TempData.DEBIT_CREDIT_AMOUNT
                     Else
                        SumCredit = SumCredit + TempData.DEBIT_CREDIT_AMOUNT
                     End If
                     TempData.Flag = "Y"
                  End If
               Next TempData
               Set HeadCf = m_Headers1.Fields(12)
               TempStr = FormatNumber(SumCredit, 2)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Total1(12) = Total1(12) + SumCredit
               Total2(12) = Total2(12) + SumCredit
            End If
         End If
         
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'         Call Vsp.AddTable(strFormat, "", TempStr2)

         Rs.MoveNext
      Wend

      PrevCode = Di.CUSTOMER_CODE
      PrevName = Di.CUSTOMER_NAME
      k = k + 1
      Call GenerateFooter(Trim(str(k)), PrevCode, PrevName, Total1, m_Details1, 0)
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
'      Vsp.FontBold = False

      Call GenerateFooter("���", "", "", Total2, m_Details1, 0)
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False

      If mcolParam("CREDIT") = 1 Then
            SumCredit = 0
            PrevKey2 = ""
            For Each TempData In m_Credit
               If PrevKey2 <> TempData.CUSTOMER_CODE And PrevKey2 <> "" And SumCredit <> 0 Then
                  Call GenerateFooter2("" & PrevKey2, "" & PrevKey3, "" & SumCredit, m_Details1, 0)
                  Call m_Details1.GetString(1, TempStr1, TempStr2)
                  strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                  Vsp.TextColor = RGB(255, 0, 0)
                  Call Vsp.AddTable(strFormat, "", TempStr2)
                  Vsp.TextColor = RGB(0, 0, 0)
                  SumCredit = 0
               End If
               If TempData.Flag <> "Y" Then
                  If TempData.DOCUMENT_TYPE = 3 Then
                     SumCredit = SumCredit - TempData.DEBIT_CREDIT_AMOUNT
                  Else
                     SumCredit = SumCredit + TempData.DEBIT_CREDIT_AMOUNT
                  End If
               End If
               PrevKey2 = TempData.CUSTOMER_CODE
               PrevKey3 = TempData.CUSTOMER_NAME
            Next TempData
         End If

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
      ElseIf J = 4 Then
         TempStr = FormatNumber(Tot(4), 0)
      ElseIf J = 6 Then
         TempStr = FormatNumber(MyDiffEx(Tot(7), Tot(5)))
      ElseIf J = 11 Then
         TempStr = FormatNumber(MyDiffEx(Tot(10), Tot(5)))
      Else
         TempStr = FormatNumber(Tot(J))
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.Align)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub
Private Sub GenerateFooter2(TXT1 As String, TXT2 As String, TXT3 As Double, Details As CFieldList, DateCount As Long)
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
         TempStr = "*"
      ElseIf J = 2 Then
         TempStr = TXT1
      ElseIf J = 3 Then
         TempStr = TXT2
      ElseIf J = 12 Then
         TempStr = FormatNumber(TXT3, 2)
      Else
         TempStr = ""
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.Align)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
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


