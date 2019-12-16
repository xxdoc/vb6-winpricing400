VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportInventory018_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportInventory018_1"
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

Private Const HEADER_FORMAT1 = "^1|^3|^3|^6|^6|^6|^6|^4;"
Private Const HEADER_FORMAT2 = "^1|^3|^3|^2|^2|^2|^2|^2|^2|^2|^2|^2|^2|^2|^2|^2|^2;"
Private Const DETAIL_FORMAT1 = "<7|<6|<22;"
Private Const DETAIL_FORMAT2 = "^1|^3|^3|>2|>2|>2|>2|>2|>2|>2|>2|>2|>2|>2|>2|>2|>2;"
Private Const FOOTER_FORMAT1 = "^1|^6|>2|>2|>2|>2|>2|>2|>2|>2|>2|>2|>2|>2|>2|>2;"
Private Const FOOTER_FORMAT2 = "^1|^3|^3|>2|>2|>2|>2|>2|>2|>2|>2|>2|>2|>2|>2|>2|>2;"

Private m_Headers0 As CFieldList
Private m_Headers1 As CFieldList
Private m_Headers2 As CFieldList
Private m_Details1 As CFieldList
Private m_Left As Double
Private m_ProductStatus As Collection

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Headers2 = New CFieldList
   Set m_Details1 = New CFieldList
   Set m_ProductStatus = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Headers2 = Nothing
   Set m_Details1 = Nothing
   Set m_ProductStatus = Nothing
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
   
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   Call m_Headers2.ClearField
   
   Set Cf = New CReportField
   
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "วัตถุดิบ", "<")
   Call m_Headers1.AddField(Cf)
   Call m_Headers2.AddField(Cf)

   Call Cf.SetFieldValue(7, "^", "", "^")
   Call m_Headers0.AddField(Cf)

   Call Cf.SetFieldValue(4, "^", "ยกมา", "^")
   Call m_Headers1.AddField(Cf)

   Call Cf.SetFieldValue(4, "^", "รับเข้า (โอน+ซื้อ+ปรับเพิ่ม)", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
   Call m_Headers2.AddField(Cf)
      
   Call Cf.SetFieldValue(8, "^", "รับเข้า", "^")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(4, "^", "จ่ายออก (โอน+ปรับลด)", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "เบิกใช้", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "ขาย", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(12, "^", "จ่ายออก", "^")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(4, "^", "คงเหลือ", "^")
   Call m_Headers1.AddField(Cf)
                                                
   Call Cf.SetFieldValue(4, "^", "คงเหลือ", "^")
   Call m_Headers0.AddField(Cf)
                                                
  Call m_Headers0.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
                                                
  Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
      
  Call m_Headers2.GetString(1, TempStr1, TempStr1_1)
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
   
   FromDate = "จาก " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
   ToDate = "ถึง " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")
   
   MODULE_DESC = vbCrLf & glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & vbCrLf & _
                                       mcolParam("REPORT_NAME") & _
                                       vbCrLf & FromDate & " - " & ToDate & " (" & EmptyToString(mcolParam("LOCATION_NAME"), "รวม") & ")"

   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))
   
   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
      mstrErrMsg = "Error GetDateTime Error."
      Exit Function
   End If
   mdteDate = InternalDateToDate(strDate)
   initDoc = True
End Function

Private Function GetPartItem(TempCol As Collection, TempID As Long) As CPartItem
On Error Resume Next
Dim Ei As CPartItem

   Set Ei = TempCol(Trim(Str(TempID)))
   If Ei Is Nothing Then
      Set GetPartItem = Nothing
   Else
      Set GetPartItem = Ei
   End If
End Function

Private Function GetImportItem(TempCol As Collection, TempKey As String) As CLotItem
On Error Resume Next
Dim Ei As CLotItem
Static TempEi As CLotItem

   Set Ei = TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CLotItem
      End If
      Set GetImportItem = TempEi
   Else
      Set GetImportItem = Ei
   End If
End Function

Private Function CompareKey(Key1 As String, Key2 As String, CompareType As Long) As Boolean
      If CompareType = 1 Then
         CompareKey = Key1 < Key2
      Else
         CompareKey = Key2 < Key1
      End If
End Function

Private Function GetNextTransaction(Rs1 As ADODB.Recordset, Rs2 As ADODB.Recordset, Ii As CLotItem, Ei As CExportItem) As String
Dim EofFlag1 As Boolean
Dim EofFlag2 As Boolean
   
   'Export
   EofFlag1 = Rs1.EOF
   If Not Rs1.EOF Then
      Call Ei.PopulateFromRS(13, Rs1)
   End If
   
   'Import
   EofFlag2 = Rs2.EOF
   If Not Rs2.EOF Then
      Call Ii.PopulateFromRS(7, Rs2)
   End If
   
   If (EofFlag1 And EofFlag2) Then
      GetNextTransaction = ""
   ElseIf (EofFlag1 And (Not EofFlag2)) Then
      GetNextTransaction = "I"
      Rs2.MoveNext
   ElseIf ((Not EofFlag1) And EofFlag2) Then
      GetNextTransaction = "E"
      Rs1.MoveNext
   Else
      '===
      'การเรียงลำดับมีผลอย่างมาก
      If Ei.PART_TYPE_NO = Ii.PART_TYPE_NO Then
         If Ei.PART_NO = Ii.PART_NO Then
            If DateToStringInt(Ei.DOCUMENT_DATE) = DateToStringInt(Ii.DOCUMENT_DATE) Then
               If Ei.TRANSACTION_SEQ < Ii.TRANSACTION_SEQ Then
                  GetNextTransaction = "E"
               Else
                  GetNextTransaction = "I"
               End If
            ElseIf DateToStringInt(Ei.DOCUMENT_DATE) < DateToStringInt(Ii.DOCUMENT_DATE) Then
               GetNextTransaction = "E"
            Else
               GetNextTransaction = "I"
            End If 'Document date
         ElseIf Ei.PART_NO < Ii.PART_NO Then
            GetNextTransaction = "E"
         Else
            GetNextTransaction = "I"
         End If 'Part no
      ElseIf CompareKey(Ei.PART_TYPE_NO, Ii.PART_TYPE_NO, mcolParam("ORDER_TYPE")) Then
         GetNextTransaction = "E"
      Else
         GetNextTransaction = "I"
      End If 'Part type no
      '===
      If GetNextTransaction = "I" Then
'Debug.Print "I" & ";" & Ii.PART_TYPE_NO & ";" & Ii.PART_NO & ";" & DateToStringInt(Ii.DOCUMENT_DATE) & ";" & Ii.TRANSACTION_SEQ & ";" & Ii.DOCUMENT_NO
         Rs2.MoveNext
      ElseIf GetNextTransaction = "E" Then
'Debug.Print "E" & ";" & Ei.PART_TYPE_NO & ";" & Ei.PART_NO & ";" & DateToStringInt(Ei.DOCUMENT_DATE) & ";" & Ei.TRANSACTION_SEQ & ";" & Ei.DOCUMENT_NO
         Rs1.MoveNext
      End If
   End If 'Eof flag
   
End Function

Private Function ItemToTable(Ei As CExportItem, Ii As CLotItem, TxCode As String, Totals1() As Double, Totals2() As Double) As String
Dim TempStr As String
Static i As Long
'NO.|วันที่|เลขที่เอกสาร|จำนวน|ราคา|มูลค่า|จำนวน|ราคา|มูลค่า| จำนวน|ราคา|มูลค่า|จำนวน|ราคา|มูลค่า|จำนวน|ราคา|มูลค่า|จำนวน|มูลค่า;
'NO.|วันที่|เลขที่เอกสาร|ยกมา|รับเข้า (โอน+ซื้อ+ปรับเพิ่ม)|จ่ายออก (โอน+ปรับลด)|เบิกใช้|ขาย|คงเหลือ;"

   TempStr = ""
   i = i + 1

   If TxCode = "I" Then
      TempStr = i & "|"
      TempStr = TempStr & InternalDateToStringEx4(DateToStringInt(Ii.DOCUMENT_DATE)) & "|"
      TempStr = TempStr & Ii.DOCUMENT_NO & "|"
      If (Ii.DOCUMENT_TYPE = 1) Or (Ii.DOCUMENT_TYPE = 3) Or (Ii.DOCUMENT_TYPE = 4) Then
         TempStr = TempStr & "" & FormatNumber(Ii.TX_AMOUNT) & "|" & Format(Ii.TOTAL_INCLUDE_PRICE / Ii.TX_AMOUNT, "0.00") & "|" & FormatNumber(Ii.TOTAL_INCLUDE_PRICE) & "|0.00|0.00|0.00|0.00|0.00|0.00|0.00|0.00|0.00|" & FormatNumber(Ii.NEW_AMOUNT) & "|" & FormatNumber(Ii.NEW_AMOUNT * Ii.NEW_PRICE) & ";"
         
         Totals1(4) = Totals1(4) + Ii.TX_AMOUNT
         Totals2(4) = Totals2(4) + Ii.TX_AMOUNT
         
         Totals1(5) = Totals1(5) + 0
         Totals2(5) = Totals2(5) + 0
      
         Totals1(6) = Totals1(6) + (Ii.TOTAL_INCLUDE_PRICE)
         Totals2(6) = Totals2(6) + (Ii.TOTAL_INCLUDE_PRICE)
      End If
      
      Totals1(16) = Ii.NEW_AMOUNT
      Totals1(17) = Ii.NEW_AMOUNT * Ii.NEW_PRICE
   ElseIf TxCode = "E" Then
      TempStr = i & "|"
      TempStr = TempStr & InternalDateToStringEx4(DateToStringInt(Ei.DOCUMENT_DATE)) & "|"
      TempStr = TempStr & Ei.DOCUMENT_NO & "|"
      If Ei.DOCUMENT_TYPE = 2 Then
         If Ei.SALE_FLAG = "Y" Then
            TempStr = TempStr & "0.00|0.00|0.00|0.00|0.00|0.00|0.00|0.00|0.00|" & FormatNumber(Ei.EXPORT_AMOUNT) & "|" & FormatNumber(Ei.EXPORT_AVG_PRICE) & "|" & FormatNumber(Ei.EXPORT_TOTAL_PRICE) & "|" & FormatNumber(Ei.CURRENT_AMOUNT) & "|" & FormatNumber(Ei.CURRENT_AMOUNT * Ei.EXPORT_AVG_PRICE) & ";"
            Totals1(13) = Totals1(13) + Ei.EXPORT_AMOUNT
            Totals2(13) = Totals2(13) + Ei.EXPORT_AMOUNT
            
            Totals1(14) = Totals1(14) + 0
            Totals2(14) = Totals2(14) + 0
         
            Totals1(15) = Totals1(15) + (Ei.EXPORT_TOTAL_PRICE)
            Totals2(15) = Totals2(15) + (Ei.EXPORT_TOTAL_PRICE)
         Else
            TempStr = TempStr & "0.00|0.00|0.00|0.00|0.00|0.00|" & FormatNumber(Ei.EXPORT_AMOUNT) & "|" & FormatNumber(Ei.EXPORT_AVG_PRICE) & "|" & FormatNumber(Ei.EXPORT_TOTAL_PRICE) & "|0.00|0.00|0.00|" & FormatNumber(Ei.CURRENT_AMOUNT) & "|" & FormatNumber(Ei.CURRENT_AMOUNT * Ei.EXPORT_AVG_PRICE) & ";"
            Totals1(10) = Totals1(10) + Ei.EXPORT_AMOUNT
            Totals2(10) = Totals2(10) + Ei.EXPORT_AMOUNT
            
            Totals1(11) = Totals1(11) + 0
            Totals2(11) = Totals2(11) + 0
         
            Totals1(12) = Totals1(12) + (Ei.EXPORT_TOTAL_PRICE)
            Totals2(12) = Totals2(12) + (Ei.EXPORT_TOTAL_PRICE)
         End If
      ElseIf (Ei.DOCUMENT_TYPE = 3) Or (Ei.DOCUMENT_TYPE = 4) Then
         TempStr = TempStr & "0.00|0.00|0.00|" & FormatNumber(Ei.EXPORT_AMOUNT) & "|" & FormatNumber(Ei.EXPORT_AVG_PRICE) & "|" & FormatNumber(Ei.EXPORT_TOTAL_PRICE) & "|0.00|0.00|0.00|0.00|0.00|0.00|" & FormatNumber(Ei.CURRENT_AMOUNT) & "|" & FormatNumber(Ei.CURRENT_AMOUNT * Ei.EXPORT_AVG_PRICE) & ";"
      
         Totals1(7) = Totals1(7) + Ei.EXPORT_AMOUNT
         Totals2(7) = Totals2(7) + Ei.EXPORT_AMOUNT
         
         Totals1(8) = Totals1(8) + 0
         Totals2(8) = Totals2(8) + 0
      
         Totals1(9) = Totals1(9) + (Ei.EXPORT_TOTAL_PRICE)
         Totals2(9) = Totals2(9) + (Ei.EXPORT_TOTAL_PRICE)
      End If
      
      Totals1(16) = Ei.CURRENT_AMOUNT
      Totals1(17) = Ei.CURRENT_AMOUNT * Ei.EXPORT_AVG_PRICE
   End If
   
   ItemToTable = TempStr
End Function

Private Sub GenerateDetail(i As Long, O As Object, Details As CFieldList, Tot() As Double, Tot2() As Double, Tot3() As Double)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim Amt1 As Double
Dim Amt2 As Double
Dim Amt3 As Double
Dim Buf(30) As Double
Dim j As Long

   Call Details.ClearField
   
   For j = 1 To UBound(Buf)
      Buf(j) = 0
   Next j
   
   Set BodyCf = New CReportField
   
   Set HeadCf = m_Headers2.Fields(1)
   Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, Trim(Str(i)), HeadCf.BodyAlign)
   Call Details.AddField(BodyCf)
   
   Set HeadCf = m_Headers2.Fields(2)
   Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, DateToStringExtEx2(O.DOCUMENT_DATE), HeadCf.BodyAlign)
   Call Details.AddField(BodyCf)
   
'   Set HeadCf = m_Headers2.Fields(3)
'   Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, O.DOCUMENT_NO, HeadCf.BodyAlign)
'   Call Details.AddField(BodyCf)
   
   If O.TX_TYPE = "I" Then
      Amt1 = O.IMPORT_AMOUNT
      If O.IMPORT_AMOUNT = 0 Then
         Amt2 = 0
      Else
         Amt2 = O.TOTAL_INCLUDE_PRICE / O.IMPORT_AMOUNT
      End If
      Amt3 = O.TOTAL_INCLUDE_PRICE
      
      Buf(3) = Amt1
      Buf(4) = Amt3
   
      Buf(11) = O.CURRENT_AMOUNT
      Buf(12) = O.CURRENT_AMOUNT * O.NEW_PRICE
   Else
      Amt1 = O.EXPORT_AMOUNT
      Amt2 = O.EXPORT_AVG_PRICE
      Amt3 = O.EXPORT_TOTAL_PRICE
            
      If (O.DOCUMENT_TYPE = 3) Or (O.DOCUMENT_TYPE = 4) Then
         Buf(5) = Amt1
         Buf(6) = Amt3
      ElseIf (O.DOCUMENT_TYPE = 2) And (O.SALE_FLAG = "N") Then
         Buf(7) = Amt1
         Buf(8) = Amt3
      ElseIf (O.DOCUMENT_TYPE = 2) And (O.SALE_FLAG = "Y") Then
         Buf(9) = Amt1
         Buf(10) = Amt3
      End If
      
      Buf(11) = O.CURRENT_AMOUNT
      Buf(12) = O.CURRENT_AMOUNT * O.EXPORT_AVG_PRICE
   End If
   
   Set HeadCf = m_Headers2.Fields(4)
   
   For j = 1 To 12
      Tot(j) = Tot(j) + Buf(j)
      If j < 11 Then
         Tot2(j) = Tot2(j) + Buf(j)
         Tot3(j) = Tot3(j) + Buf(j)
      End If
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, FormatNumber(Buf(j)), HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next j
   
   Tot(11) = Buf(11)
   Tot(12) = Buf(12)
   
   Set BodyCf = Nothing
End Sub

Private Sub GenerateBalanceDetail(i As Long, O As Object, Details As CFieldList, Tot() As Double, Tot2() As Double, Tot3() As Double)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim Amt1 As Double
Dim Amt2 As Double
Dim Amt3 As Double
Dim Buf(30) As Double
Dim j As Long

   Call Details.ClearField
   
   For j = 1 To UBound(Buf)
      Buf(j) = 0
   Next j
   
   Set BodyCf = New CReportField
   
   Set HeadCf = m_Headers2.Fields(1)
   Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, "", HeadCf.BodyAlign)
   Call Details.AddField(BodyCf)
   
   Set HeadCf = m_Headers2.Fields(2)
   Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, "***ยกมา***", HeadCf.BodyAlign)
   Call Details.AddField(BodyCf)
   
'   Set HeadCf = m_Headers2.Fields(3)
'   Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, "", HeadCf.BodyAlign)
'   Call Details.AddField(BodyCf)
   
   If O.TX_TYPE = "I" Then
      Amt1 = O.CURRENT_AMOUNT
      Amt2 = O.NEW_PRICE
      Amt3 = O.NEW_PRICE * O.CURRENT_AMOUNT
      
      Buf(1) = Amt1
      Buf(2) = Amt3
      
      Buf(11) = Amt1
      Buf(12) = Amt3
   Else
      Amt1 = O.CURRENT_AMOUNT
      Amt2 = O.EXPORT_AVG_PRICE
      Amt3 = O.EXPORT_AVG_PRICE * O.CURRENT_AMOUNT
      
      Buf(1) = Amt1
      Buf(2) = Amt3
      
      Buf(11) = Amt1
      Buf(12) = Amt3
   End If
   
   Set HeadCf = m_Headers2.Fields(4)
   
   For j = 1 To 12
      Tot(j) = Tot(j) + Buf(j)
      If j < 11 Then
         Tot2(j) = Tot2(j) + Buf(j)
         Tot3(j) = Tot3(j) + Buf(j)
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, FormatNumber(Buf(j)), HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next j
   
   Tot(11) = Buf(11)
   Tot(12) = Buf(12)
   
   Set BodyCf = Nothing
End Sub

Private Sub GenerateFooter(Txt1 As String, Txt2 As String, Txt3 As String, Tot() As Double, Details As CFieldList)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim Amt1 As Double
Dim Amt2 As Double
Dim Amt3 As Double
Dim Buf(30) As Double
Dim j As Long

   Call Details.ClearField
   
   For j = 1 To UBound(Buf)
      Buf(j) = 0
   Next j
   
   Set BodyCf = New CReportField
   
   Set HeadCf = m_Headers2.Fields(1)
   Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, Txt1, HeadCf.BodyAlign)
   Call Details.AddField(BodyCf)
   
   Set HeadCf = m_Headers2.Fields(2)
   Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, Txt2, HeadCf.BodyAlign)
   Call Details.AddField(BodyCf)
   
'   Set HeadCf = m_Headers2.Fields(3)
'   Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, Txt3, HeadCf.BodyAlign)
'   Call Details.AddField(BodyCf)
   
   Set HeadCf = m_Headers2.Fields(4)
   For j = 1 To 12
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, FormatNumber(Tot(j)), HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next j
      
   Set BodyCf = Nothing
End Sub

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
Dim PartTypes As Collection
Dim Houses As Collection
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Hs As CLocation
Dim Pi As CPartItem
Dim Ib As CLotItem
Dim Total1(30) As Double
Dim Total2(30) As Double
Dim Total3(30) As Double
Dim PiCount As Long
Dim iCount As Long
Dim OldColor As Double
Dim Pt As CPartType

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   Set Rs1 = New ADODB.Recordset
   Set Rs2 = New ADODB.Recordset
   
   Set PartItems = New Collection
   
   Set PartTypes = New Collection
   Call LoadPartType(Nothing, PartTypes)
   
   Set InventoryBals = New Collection
   Call LoadInventoryBalance(Nothing, InventoryBals, InternalDateToDate(DateToStringIntLow(mcolParam("FROM_DATE"))), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")), mcolParam("LOCATION_ID"), "N")
   
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
   
   Set BodyCf = New CReportField
   
   Vsp.StartDoc
      For Each Pt In PartTypes
         For j = 1 To UBound(Total2)
            Total2(j) = 0
         Next j
         
         Call LoadImportPartItem(Nothing, PartItems, Pt.PART_TYPE_ID, mcolParam("LOCATION_ID"), mcolParam("FROM_DATE"), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")))
         PiCount = 0
         For Each Pi In PartItems
            PiCount = PiCount + 1
            For j = 1 To UBound(Total1)
               Total1(j) = 0
            Next j
            'Head
            Call m_Details1.ClearField
      
            Call BodyCf.SetFieldValue(7, "<", Pt.PART_TYPE_NAME & " (" & Pt.PART_TYPE_NO & ")", "^")
            Call m_Details1.AddField(BodyCf)
            Call BodyCf.SetFieldValue(12, "<", Pi.PART_DESC & " (" & Pi.PART_NO & ")", "^")
            Call m_Details1.AddField(BodyCf)
            Call BodyCf.SetFieldValue(18, "<", Pi.UNIT_NAME, "^")
            Call m_Details1.AddField(BodyCf)
            
            Call m_Details1.GetString(1, TempStr1, TempStr2)
             strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'             Call Vsp.AddTable(strFormat, "", TempStr2)
            'Head

            '=== ยกมา
            Set O = GetBalanceItem(InventoryBals, Pi.PART_ITEM_ID, mcolParam("LOCATION_ID"), InternalDateToDate(DateToStringIntLow(mcolParam("FROM_DATE"))))
            Call GenerateBalanceDetail(0, O, m_Details1, Total1, Total2, Total3)
            Call m_Details1.GetString(2, TempStr1, TempStr2)
             strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'             Call Vsp.AddTable(strFormat, "", TempStr2)
            '=== ยกมา
            
            '=== Detail
            Set Ei = New CExportItem
            Ei.EXPORT_ITEM_ID = -1
            Ei.FROM_DATE = mcolParam("FROM_DATE")
            Ei.TO_DATE = mcolParam("TO_DATE")
            Ei.COMMIT_FLAG = CommitTypeToFlag(mcolParam("COMMIT_TYPE"))
            Ei.PIG_FLAG = "N"
            Ei.PART_ITEM_ID = Pi.PART_ITEM_ID
            Ei.LOCATION_ID = mcolParam("LOCATION_ID")
            Ei.OrderBy = mcolParam("ORDER_BY")
            Ei.OrderType = mcolParam("ORDER_TYPE")
            Call Ei.QueryData(13, Rs1, iCount)
            
            Set Ii = New CLotItem
            Ii.LOT_ITEM_ID = -1
            Ii.FROM_DATE = mcolParam("FROM_DATE")
            Ii.TO_DATE = mcolParam("TO_DATE")
            Ii.COMMIT_FLAG = CommitTypeToFlag(mcolParam("COMMIT_TYPE"))
            Ii.PIG_FLAG = "N"
            Ii.PART_ITEM_ID = Pi.PART_ITEM_ID
            Ii.LOCATION_ID = mcolParam("LOCATION_ID")
            Ii.OrderBy = mcolParam("ORDER_BY")
            Ii.OrderType = mcolParam("ORDER_TYPE")
            Call Ii.QueryData(7, Rs2, iCount)
            '== Detail
            
            TxCode = "X"
            PiCount = 0
            While TxCode <> ""
               PiCount = PiCount + 1
               TxCode = GetNextTransaction(Rs1, Rs2, Ii, Ei)
               If TxCode <> "" Then
                  If TxCode = "I" Then
                     Set O = Ii
                  ElseIf TxCode = "E" Then
                     Set O = Ei
                  End If
                                       
                  m_Details1.ClearField
                  Call GenerateDetail(i, O, m_Details1, Total1, Total2, Total3)
                  Call m_Details1.GetString(2, TempStr1, TempStr2)
                  Vsp.FontBold = False
                  strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'                  Call Vsp.AddTable(strFormat, "", TempStr2)
               End If 'Tx code
            Wend
            
         Total2(11) = Total2(11) + Total1(11)
         Total2(12) = Total2(12) + Total1(12)
            
           i = i + 1
           
            m_Details1.ClearField
            Call GenerateFooter(Trim(Str(i)), Pt.PART_TYPE_NAME & " (" & Pt.PART_TYPE_NO & ")", Pi.PART_DESC & " (" & Pi.PART_NO & ")", Total1, m_Details1)
            Call m_Details1.GetString(2, TempStr1, TempStr2)
'            Vsp.FontBold = True
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
'            Vsp.FontBold = False
         Next Pi
         
         '=== Summary for each type
         If PiCount > 0 Then
            Total3(11) = Total3(11) + Total2(11)
            Total3(12) = Total3(12) + Total2(12)
            
            m_Details1.ClearField
            Call GenerateFooter("รวม", "*** " & Pt.PART_TYPE_NAME & " (" & Pt.PART_TYPE_NO & ") ***", "", Total2, m_Details1)
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            Vsp.FontBold = True
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
         End If
         '=== Summary for each type
      Next Pt
   
      m_Details1.ClearField
      Call GenerateFooter("สุทธิ", "", "", Total3, m_Details1)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      Vsp.FontBold = True
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
   
   Vsp.EndDoc
   
   Set BodyCf = Nothing
   
   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing
      
   If Rs2.State = adStateOpen Then
      Rs2.Close
   End If
   Set Rs2 = Nothing
      
   Set Ei = Nothing
   Set Ii = Nothing
   Set InventoryBals = Nothing
   Set PartTypes = Nothing
   Set PartItems = Nothing
   
   genDoc = True
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
   Set Rs = Nothing
End Function

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
   Call Vsp.AddTable(strFormat, "", "|วัน|: " & DateToStringExt(mdteDate))
   Call Vsp.AddTable(strFormat, "", "|เวลา|: " & Right$(DateToStringExtEx(mdteDate), 8))
   Call Vsp.AddTable(strFormat, "", "|เลขประจำตัวผู้เสียภาษี|: " & glbEnterPrise.TAX_ID)
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


