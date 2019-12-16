VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportFormReceipt002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportInvDoc001"
Private MODULE_DESC  As String


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
Private m_Rs1 As ADODB.Recordset
Private COLL As Collection

Private Const TITLE_SIZE = 12
Private Const HEADER_SIZE = 12
Private Const DETAIL_SIZE = 12

Private Const TITLE_FORMAT = "<32|<3|<5;"
Private Const HEADER_FORMAT = "^1|^6|^2|^2|^3;"
Private Const DETAIL_FORMAT = "^1|<6|>2|>2|>3;"
Private Const FOOTER_FORMAT = "^1|>6|>2|>2|>3;"

Private m_Enterprise As CEnterprise
Private m_CustAddress As CAddress
Private m_EnpAddress As CAddress
Private m_EnpPerson As cName
Private Sp As CSystemParam
Private m_ReportConfig As CReportConfig
Private m_cData As CBillingDoc
Private m_ImageWidth As Double
Private m_ImageHeight As Double
Private TYPEDO As Long
Private countLine As Integer
Private totalPage As Integer


Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_Enterprise = New CEnterprise
   Set m_CustAddress = New CAddress
   Set m_EnpAddress = New CAddress
   Set m_EnpPerson = New cName
   Set m_ReportConfig = New CReportConfig
   Set COLL = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_Enterprise = Nothing
   Set m_CustAddress = Nothing
   Set m_EnpAddress = Nothing
   Set m_EnpPerson = Nothing
   Set m_ReportConfig = Nothing
   Set COLL = Nothing
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
Private Sub PrintPOHeader(alngX() As Long)
   Vsp.MarginLeft = 0.2 * 567 - m_ReportConfig.DUMMY_OFFSET
   Vsp.MarginRight = 0.2 * 567 - m_ReportConfig.DUMMY_OFFSET
   
   'Vsp.TableBorder = tbAll
'   Vsp.TableBorder =tbNone
   Call GenerateReceiptHeader(Vsp, mcolParam, tbNone, m_ReportConfig.HEAD_OFFSET)
End Sub

Private Sub printHeader()
   Dim strFormat As String
   Dim strPrint As String
   Dim tbdT As TableBorderSettings
   Dim blnBold As Boolean
   Dim iSize As Integer
   Dim alngX() As Long
   
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Call PrintPOHeader(alngX)
   
   Vsp.FontSize = DETAIL_SIZE
'  Vsp.FontBold = True
'   Vsp.TableBorder = tbAll
'   strFormat = VSP_CalTable(HEADER_FORMAT, mdblWid, alngX)
'   strPrint = "ลำดับ|สินค้าและบริการ|จำนวน|ราคาต่อหน่วย|จำนวนเงิน;"
'   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Vsp.TableBorder = tbdT
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
End Sub

Public Sub GenerateReceiptHeader(Vsp As VSPrinter, mcolParam As Collection, TempBorder As TableBorderSettings, Offset As Long)
Dim Amt As Double
Dim VatAmt As Double
Dim NetAmt As Double
Dim Bd As CBillingDoc
 Dim numLinePerPage As Integer
Dim countPage As Collection
  
  numLinePerPage = 7
  Set countPage = m_cData.DoItems
  countLine = countPage.Count
    totalPage = Int((countLine / 7) + 1)
    
   Vsp.FontBold = True
   Vsp.TableBorder = TempBorder
   
    Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 3
    Vsp.TableCell(tcRowHeight) = Offset + (0.7 * 567)

    Vsp.TableCell(tcColWidth, , 1) = "16.7cm"
    Vsp.TableCell(tcColWidth, , 2) = "3.5cm"
    
   Vsp.TableCell(tcAlign, 2, 2) = taLeftBottom
   Vsp.TableCell(tcText, 2, 2) = "     " & mcolParam("DOCUMENT_NO") & "     " & " (" & "" & Vsp.CurrentPage & "/" & totalPage & "" & ")"
   
   'totalPage
   Vsp.TableCell(tcAlign, 3, 2) = taLeftBottom
   Vsp.TableCell(tcText, 3, 2) = "     " & DateToStringExtEx2(mcolParam("DOCUMENT_DATE"))
     
   Vsp.EndTable
'------------------------------------------------------------------------------------------------------------> OFFSET

   Vsp.StartTable
   Vsp.TableCell(tcCols) = 4
   Vsp.TableCell(tcRows) = 4
'   Vsp.TableCell(tcRowHeight) = "0.6cm"
   Vsp.TableCell(tcRowHeight) = (0.6 * 567)
   Vsp.TableBorder = TempBorder
   Vsp.TableCell(tcColWidth, , 1) = "1.8cm"
   Vsp.TableCell(tcColWidth, , 2) = "10.2cm"
   Vsp.TableCell(tcColWidth, , 3) = "4.1cm"
   Vsp.TableCell(tcColWidth, , 4) = "4.1cm"
   
   Vsp.TableCell(tcAlign, 1, 2) = taLeftBottom
   Vsp.TableCell(tcText, 1, 2) = mcolParam("CUSTOMER_NAME") & "    รุ่น  " & mcolParam("GENERATION")
   
   Vsp.TableCell(tcAlign, 2, 2) = taLeftBottom
   Vsp.TableCell(tcText, 2, 2) = mcolParam("CUSTOMER_ADDRESS")
   
     Vsp.TableCell(tcAlign, 3, 2) = taLeftBottom
    Vsp.TableCell(tcText, 3, 2) = "Ref.  " & mcolParam("REFERENCE")
     
      Vsp.TableCell(tcAlign, 4, 2) = taLeftBottom
     Vsp.TableCell(tcText, 4, 2) = mcolParam("FARM_NAME")
   
   Vsp.TableCell(tcAlign, 2, 3) = taCenterBottom
   Vsp.TableCell(tcText, 2, 3) = mcolParam("REF")
   Vsp.TableCell(tcAlign, 4, 3) = taCenterBottom
    Vsp.TableCell(tcText, 2, 4) = mcolParam("Credit")
    Vsp.TableCell(tcAlign, 2, 4) = taCenterBottom
    Vsp.TableCell(tcText, 4, 3) = mcolParam("DUE_DATE")
   
   Vsp.TableCell(tcAlign, 4, 4) = taCenterBottom
   Vsp.TableCell(tcText, 4, 4) = mcolParam("NOTE")
   
   
   Vsp.EndTable
   
   
   
  
    Vsp.StartTable
      Vsp.TableCell(tcCols) = 9
      Vsp.TableCell(tcRows) = 1
    '  Vsp.TableCell(tcRowHeight) = "0.8cm"
      Vsp.TableCell(tcRowHeight) = (0.8 * 567)
      Vsp.TableBorder = TempBorder

     Vsp.TableCell(tcColWidth, , 1) = "6.5cm"
      Vsp.TableCell(tcColWidth, , 2) = "1.5cm"
      Vsp.TableCell(tcColWidth, , 3) = "1.5cm"
      Vsp.TableCell(tcColWidth, , 4) = "1cm"
      Vsp.TableCell(tcColWidth, , 5) = "1.9cm"
      Vsp.TableCell(tcColWidth, , 6) = "1.9cm"
      Vsp.TableCell(tcColWidth, , 7) = "1.5cm"
      Vsp.TableCell(tcColWidth, , 8) = "1.5cm"
      Vsp.TableCell(tcColWidth, , 9) = "2.9cm"
  Vsp.EndTable
End Sub
Public Function GenerateReceiptBody(Vsp As VSPrinter, mcolParam As Collection, Bd As CBillingDoc, TempBorder As TableBorderSettings)
Dim Poi As CReceiptItem
Dim Bi As CDoItem
Dim I As Long
Dim J As Long
Dim Sum As Double
Dim Offset As Long

   I = -1
   J = 0
   Sum = 0
   Offset = 0
  Vsp.FontBold = True
   For Each Bi In m_cData.DoItems

    I = I + 1
       
      If I > 6 Then
        I = 1
        
         Vsp.NewPage
         
      End If
   
     Vsp.StartTable
      Vsp.TableCell(tcCols) = 9
      Vsp.TableCell(tcRows) = 1
      'Vsp.TableCell(tcRowHeight) = Offset + (0.7 * 567)
      Vsp.TableCell(tcRowHeight) = "0.7cm"
'      Offset = Offset + (0.7 * 567)
      Vsp.TableBorder = TempBorder

     Vsp.TableCell(tcColWidth, , 1) = "6.3cm"
      Vsp.TableCell(tcColWidth, , 2) = "1.5cm"
      Vsp.TableCell(tcColWidth, , 3) = "2.5cm"
      'Vsp.TableCell(tcColWidth, , 4) = "1cm"
      Vsp.TableCell(tcColWidth, , 4) = "1.9cm"
      Vsp.TableCell(tcColWidth, , 5) = "1.9cm"
      Vsp.TableCell(tcColWidth, , 6) = "1.5cm"
      Vsp.TableCell(tcColWidth, , 7) = "1.5cm"
      Vsp.TableCell(tcColWidth, , 8) = "2.9cm"
     Vsp.TableCell(tcAlign, 1, 1) = taLeftMiddle
      Vsp.TableCell(tcText, 1, 1) = Bi.ShowCodeText & " " & Bi.ShowDescText

     'Bi.ShowCodeText, Bi.ShowDescText,
     
     Vsp.TableCell(tcAlign, 1, 2) = taRightMiddle
      Vsp.TableCell(tcText, 1, 2) = FormatNumber(Bi.WEIGHT_PER_PACK)

     Vsp.TableCell(tcAlign, 1, 3) = taRightMiddle
      Vsp.TableCell(tcText, 1, 3) = FormatNumber(Bi.PACK_AMOUNT) & " " & Bi.UNIT_NAME
      
'      Vsp.TableCell(tcAlign, 1, 4) = taCenterMiddle
'      Vsp.TableCell(tcText, 1, 4) = Bi.UNIT_NAME
      
'      Vsp.TableCell(tcAlign, 1, 4) = taRightMiddle
'      Vsp.TableCell(tcText, 1, 4) = FormatNumber(Bi.ITEM_AMOUNT)
      
      Vsp.TableCell(tcAlign, 1, 5) = taRightMiddle
      Vsp.TableCell(tcText, 1, 5) = FormatNumber(Bi.PRICE_PER_PACK)
      
      Vsp.TableCell(tcAlign, 1, 6) = taRightMiddle
      Vsp.TableCell(tcText, 1, 6) = FormatNumber(Bi.DISCOUNT_PERCENT)
      
      Vsp.TableCell(tcAlign, 1, 7) = taRightMiddle
      Vsp.TableCell(tcText, 1, 7) = FormatNumber(Bi.DISCOUNT_AMOUNT)
      
      
      Vsp.TableCell(tcAlign, 1, 8) = taRightMiddle
      Vsp.TableCell(tcText, 1, 8) = FormatNumber(Bi.TOTAL_PRICE)
     
     Sum = Sum + Bi.TOTAL_PRICE
     
     Vsp.EndTable
      

   Next Bi
' Private countLine As Integer
'Private totalPage As Integer
'  Set countPage = m_cData.DoItems
'  countLine = countPage.Count
'    totalPage = Int((countLine / 7) + 1)
   
     If I < 7 Then
       Vsp.StartTable
       Vsp.TableCell(tcCols) = 1
       Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight) = (0.7 * 567) * (6 - I)
       Vsp.TableBorder = TempBorder
        Vsp.TableCell(tcColWidth, , 1) = "20cm"
       Vsp.EndTable
  
      End If

   
'   For Each Poi In Bd.ReceiptItems
'      I = I + 1
'      j = j + 1
'
'      If j > 7 Then
'         j = 1
'         Vsp.NewPage
'      End If
''
'      Vsp.StartTable
'      Vsp.TableCell(tcCols) = 9
'      Vsp.TableCell(tcRows) = 7
'      Vsp.TableCell(tcRowHeight) = "0.8cm"
'      Vsp.TableBorder = TempBorder
'
'      Vsp.TableCell(tcColWidth, , 1) = "6.5cm"
'      Vsp.TableCell(tcColWidth, , 2) = "1.5cm"
'      Vsp.TableCell(tcColWidth, , 3) = "1.5cm"
'      Vsp.TableCell(tcColWidth, , 4) = "1cm"
'      Vsp.TableCell(tcColWidth, , 5) = "1.9cm"
'      Vsp.TableCell(tcColWidth, , 6) = "1.9cm"
'      Vsp.TableCell(tcColWidth, , 7) = "1.5cm"
'      Vsp.TableCell(tcColWidth, , 8) = "1.5cm"
'      Vsp.TableCell(tcColWidth, , 9) = "2.9cm"
'
'
''
''      Vsp.TableCell(tcAlign, 2, 2) = taLeftMiddle
''      Vsp.TableCell(tcText, 2, 2) = DateToStringExtEx2(Poi.DOCUMENT_DATE)
'       Vsp.TableCell(tcAlign, 2, 2) = taLeftMiddle
'       Vsp.TableCell(tcText, 2, 2) = Poi.PART_DESC
'
'      Vsp.TableCell(tcAlign, 2, 3) = taRightMiddle
'      Vsp.TableCell(tcText, 2, 3) = Left(FormatNumber(Poi.PAID_AMOUNT), Len(FormatNumber(Poi.PAID_AMOUNT)) - 3)
'
'      Vsp.TableCell(tcAlign, 2, 4) = taCenterMiddle
'      Vsp.TableCell(tcText, 2, 4) = Right(FormatNumber(Poi.PAID_AMOUNT), 2)
''
'      Sum = Sum + Poi.PAID_AMOUNT
'
'      Vsp.EndTable
'   Next Poi
'
''     Vsp.StartTable
''      Vsp.TableCell(tcCols) = 9
''      Vsp.TableCell(tcRows) = 7
''      Vsp.TableCell(tcRowHeight) = "0.8cm"
''      Vsp.TableBorder = TempBorder
''
''      Vsp.TableCell(tcColWidth, , 1) = "6.5cm"
''      Vsp.TableCell(tcColWidth, , 2) = "1.5cm"
''      Vsp.TableCell(tcColWidth, , 3) = "1.5cm"
''      Vsp.TableCell(tcColWidth, , 4) = "1cm"
''      Vsp.TableCell(tcColWidth, , 5) = "1.9cm"
''      Vsp.TableCell(tcColWidth, , 6) = "1.9cm"
''      Vsp.TableCell(tcColWidth, , 7) = "1.5cm"
''      Vsp.TableCell(tcColWidth, , 8) = "1.5cm"
''      Vsp.TableCell(tcColWidth, , 9) = "2.9cm"
''      Vsp.EndTable
'
'  For I = 1 To (6 - j)
'      Vsp.StartTable
'      Vsp.TableCell(tcCols) = 9
'      Vsp.TableCell(tcRows) = 7
'      Vsp.TableCell(tcRowHeight) = "0.8cm"
'      Vsp.TableBorder = TempBorder
'
'        Vsp.TableCell(tcColWidth, , 1) = "6.5cm"
'         Vsp.TableCell(tcColWidth, , 1) = "6.5cm"
'      Vsp.TableCell(tcColWidth, , 2) = "1.5cm"
'      Vsp.TableCell(tcColWidth, , 3) = "1.5cm"
'      Vsp.TableCell(tcColWidth, , 4) = "1cm"
'      Vsp.TableCell(tcColWidth, , 5) = "1.9cm"
'      Vsp.TableCell(tcColWidth, , 6) = "1.9cm"
'      Vsp.TableCell(tcColWidth, , 7) = "1.5cm"
'      Vsp.TableCell(tcColWidth, , 8) = "1.5cm"
'      Vsp.TableCell(tcColWidth, , 9) = "2.9cm"
''
''      Vsp.TableCell(tcAlign, 2, 1) = taCenterMiddle
''      Vsp.TableCell(tcText, 2, 1) = I
''      Vsp.TableCell(tcAlign, 2, 2) = taCenterMiddle
''      Vsp.TableCell(tcText, 2, 2) =
'     Vsp.TableCell(tcAlign, 2, 1) = taLeftMiddle
'      Vsp.TableCell(tcText, 2, 1) = I
'
'
'      Vsp.EndTable
'  Next I

   Vsp.StartTable
      Vsp.TableCell(tcCols) = 4
      Vsp.TableCell(tcRows) = 2
      Vsp.TableCell(tcRowHeight) = "0.65cm"
      Vsp.TableBorder = TempBorder

      Vsp.TableCell(tcColWidth, , 1) = "10.5cm"
      Vsp.TableCell(tcColWidth, , 2) = "1.9cm"
      Vsp.TableCell(tcColWidth, , 3) = "4.9cm"
      Vsp.TableCell(tcColWidth, , 4) = "2.7cm"



      Vsp.TableCell(tcAlign, 1, 4) = taRightMiddle
      Vsp.TableCell(tcText, 1, 4) = FormatNumber(Sum)

      Vsp.EndTable
  
      Vsp.FontSize = TITLE_SIZE
      Vsp.FontBold = True
      
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 3
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight) = "0.8cm"
      Vsp.TableBorder = TempBorder

      Vsp.TableCell(tcColWidth, , 1) = "13cm"
      Vsp.TableCell(tcColWidth, , 2) = "3.6cm"
      Vsp.TableCell(tcColWidth, , 3) = "3.4cm"
 
      Vsp.TableCell(tcAlign, 1, 1) = taLeftMiddle
      Vsp.TableCell(tcText, 1, 1) = "(" & ThaiBaht(Sum) & ")"
      
       Vsp.TableCell(tcAlign, 1, 3) = taRightMiddle
      Vsp.TableCell(tcText, 1, 3) = FormatNumber(Sum)

      Vsp.EndTable
      
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 3
      Vsp.TableCell(tcRows) = 2
      Vsp.TableCell(tcRowHeight) = "1cm"
      Vsp.TableBorder = TempBorder

      Vsp.TableCell(tcColWidth, , 1) = "13cm"
      Vsp.TableCell(tcColWidth, , 2) = "3.6cm"
      Vsp.TableCell(tcColWidth, , 3) = "3.6cm"
 
      Vsp.TableCell(tcAlign, 1, 3) = taCenterMiddle
      Vsp.TableCell(tcText, 1, 3) = Bd.EMP_NAME & " " & Bd.EMP_LNAME

      Vsp.EndTable
End Function

'Public Sub GenerateReceiptFooter(Vsp As VSPrinter, mcolParam As Collection, TempBorder As TableBorderSettings, Bd As CBillingDoc)
'Dim Amt As Double
'Dim VatAmt As Double
'Dim NetAmt As Double
'Dim Ct As CCashTran
'
'   Vsp.FontBold = True
'   Vsp.TableBorder = TempBorder
'
'   Vsp.FontSize = 14
'
'   For Each Ct In Bd.Payments
'      If Ct.GetFieldValue("PAYMENT_TYPE") = 1 Then   'เงินสด
'         Vsp.StartTable
'         Vsp.TableCell(tcCols) = 4
'         Vsp.TableCell(tcRows) = 1
'         Vsp.TableCell(tcRowHeight) = "0.6cm"
'         Vsp.TableBorder = TempBorder
'         Vsp.TableCell(tcColWidth, , 1) = "2cm"
'         Vsp.TableCell(tcColWidth, , 2) = "0.4cm"
'         Vsp.TableCell(tcColWidth, , 3) = "2cm"
'         Vsp.TableCell(tcColWidth, , 4) = "14.6cm"
'
'         Vsp.TableCell(tcAlign, 1, 2) = taCenterMiddle
'         Vsp.TableCell(tcText, 1, 2) = "X"
'
'         Vsp.EndTable
'      ElseIf Ct.GetFieldValue("PAYMENT_TYPE") = 2 Then   'เงินโอน
'         Vsp.StartTable
'         Vsp.TableCell(tcCols) = 4
'         Vsp.TableCell(tcRows) = 1
'         Vsp.TableCell(tcRowHeight) = "0.6cm"
'         Vsp.TableBorder = TempBorder
'         Vsp.TableCell(tcColWidth, , 1) = "2cm"
'         Vsp.TableCell(tcColWidth, , 2) = "0.4cm"
'         Vsp.TableCell(tcColWidth, , 3) = "2cm"
'         Vsp.TableCell(tcColWidth, , 4) = "14.6cm"
'
'         Vsp.TableCell(tcAlign, 1, 4) = taLeftBottom
'         Vsp.TableCell(tcText, 1, 4) = "( X )  เงินโอน   " & "ธนาคาร " & Ct.GetFieldValue("BANK_NAME") & " " & Ct.GetFieldValue("BRANCH_NAME") & "  เลขที่บัญชี   " & Ct.GetFieldValue("ACCOUNT_NAME")
'         Vsp.EndTable
'
'      ElseIf Ct.GetFieldValue("PAYMENT_TYPE") = 3 Then   'เช็ค
'         Vsp.StartTable
'         Vsp.TableCell(tcCols) = 4
'         Vsp.TableCell(tcRows) = 1
'         Vsp.TableCell(tcRowHeight) = "0.6cm"
'         Vsp.TableBorder = TempBorder
'         Vsp.TableCell(tcColWidth, , 1) = "2cm"
'         Vsp.TableCell(tcColWidth, , 2) = "0.4cm"
'         Vsp.TableCell(tcColWidth, , 3) = "2cm"
'         Vsp.TableCell(tcColWidth, , 4) = "14.6cm"
'         Vsp.EndTable
'
'         Vsp.StartTable
'         Vsp.TableCell(tcCols) = 6
'         Vsp.TableCell(tcRows) = 1
'         Vsp.TableCell(tcRowHeight) = "0.6cm"
'         Vsp.TableBorder = TempBorder
'         Vsp.TableCell(tcColWidth, , 1) = "2cm"
'         Vsp.TableCell(tcColWidth, , 2) = "0.4cm"
'         Vsp.TableCell(tcColWidth, , 3) = "3.5cm"
'         Vsp.TableCell(tcColWidth, , 4) = "5cm"
'         Vsp.TableCell(tcColWidth, , 5) = "2.5cm"
'         Vsp.TableCell(tcColWidth, , 6) = "5.4cm"
'
'         Vsp.TableCell(tcAlign, 1, 2) = taCenterMiddle
'         Vsp.TableCell(tcText, 1, 2) = "X"
'
'         Vsp.TableCell(tcAlign, 1, 4) = taLeftBottom
'         Vsp.TableCell(tcText, 1, 4) = Ct.GetFieldValue("BANK_NAME")
'
'         Vsp.TableCell(tcAlign, 1, 6) = taLeftBottom
'         Vsp.TableCell(tcText, 1, 6) = Ct.GetFieldValue("BRANCH_NAME")
'
'         Vsp.EndTable
'
'         Vsp.StartTable
'         Vsp.TableCell(tcCols) = 6
'         Vsp.TableCell(tcRows) = 1
'         Vsp.TableCell(tcRowHeight) = "0.6cm"
'         Vsp.TableBorder = TempBorder
'         Vsp.TableCell(tcColWidth, , 1) = "2cm"
'         Vsp.TableCell(tcColWidth, , 2) = "0.4cm"
'         Vsp.TableCell(tcColWidth, , 3) = "3.5cm"
'         Vsp.TableCell(tcColWidth, , 4) = "5cm"
'         Vsp.TableCell(tcColWidth, , 5) = "2.5cm"
'         Vsp.TableCell(tcColWidth, , 6) = "5.4cm"
'
'         Vsp.TableCell(tcAlign, 1, 2) = taCenterMiddle
'         Vsp.TableCell(tcText, 1, 2) = ""
'
'         Vsp.TableCell(tcAlign, 1, 4) = taLeftBottom
'         Vsp.TableCell(tcText, 1, 4) = Ct.GetFieldValue("CHEQUE_NO")
'
'         Vsp.TableCell(tcAlign, 1, 6) = taLeftBottom
'         Vsp.TableCell(tcText, 1, 6) = DateToStringExtEx2(Ct.GetFieldValue("CHEQUE_DATE"))
'
'         Vsp.EndTable
'
'
'      End If
'   Next Ct
'
'   Vsp.FontBold = False
'
'End Sub
Private Function initDoc() As Boolean
   Dim strDate As String
   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
      
   Vsp.MarginLeft = 0
   Vsp.MarginRight = 0
    Vsp.MarginTop = 0
  
   Vsp.PaperSize = pprLetter
   Vsp.ORIENTATION = orPortrait
   
   Vsp.MarginBottom = "0cm"
   Vsp.MarginFooter = "1cm"
   Vsp.MarginHeader = "0cm"
   Vsp.MarginTop = "0.1cm"
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   MODULE_DESC = MapText("ใบเสร็จรับเงิน")
   
   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"), m_ReportConfig, False)
   
   'mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
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
Dim iCount As Long
Dim I As Long
Dim strFormat As String
Dim strPrint As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Poi As CDoItem
Dim Sum1 As Double
Dim Sum2 As Double
Dim Sum3 As Double
Dim J As Long
Dim TempBorder As TableBorderSettings
Dim Addr As CAddress
Dim SumDiscount As Double

   RName = "genDoc"
   
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   Set Rs = New ADODB.Recordset
   Set m_cData = New CBillingDoc
   m_cData.BILLING_DOC_ID = mcolParam("BILLING_DOC_ID")
   m_cData.QueryFlag = 1
   If Not glbDaily.QueryBillingDoc(m_cData, Rs, iCount, IsOK, glbErrorLog) Then
      mstrErrMsg = "Query Error."
      Exit Function
   End If
      
   Call m_cData.PopulateFromRS(1, Rs)
   Call mcolParam.add(m_cData.DOCUMENT_NO, "DOCUMENT_NO")
   Call mcolParam.add(m_cData.DOCUMENT_DATE, "DOCUMENT_DATE")
  Call mcolParam.add(m_cData.CUSTOMER_CODE, "CUSTOMER_CODE")
  Call mcolParam.add(m_cData.CUSTOMER_NAME, "CUSTOMER_NAME")
  Call mcolParam.add(m_cData.REF, "REF")
  Call mcolParam.add(m_cData.NOTE, "NOTE")
  Call mcolParam.add(m_cData.Credit, "Credit")
  Call mcolParam.add(DateToStringExtEx2(m_cData.DUE_DATE), "DUE_DATE")
  Call mcolParam.add(m_cData.GENERATION, "GENERATION")
 Call mcolParam.add(m_cData.REFERENCE, "REFERENCE")
  Call mcolParam.add(m_cData.FARM_NAME, "FARM_NAME")
       
     ' mcolParam("GENERATION")
    
     'DateToStringExtEx2(m_cData.DUE_DATE)
   'mcolParam("BILLING_DOC_ID")
   If m_cData.BILLING_ADDRESS_ID > 0 Then
      Set Addr = New CAddress
      Addr.ADDRESS_ID = m_cData.BILLING_ADDRESS_ID
      Call Addr.QueryData(Rs, iCount)
      If Not Rs.EOF Then
         Call Addr.PopulateFromRS(Rs)
         Call mcolParam.add(Addr.PackAddress, "CUSTOMER_ADDRESS")
      End If
      Set Addr = Nothing
   Else
      Call mcolParam.add("", "CUSTOMER_ADDRESS")
   End If
   
   If (Rs Is Nothing) Or (iCount < 1) Then
      mstrErrMsg = "ไม่พบข้อมูลที่ต้องการในฐานข้อมูล"
      Exit Function
   End If
   
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
   Sum1 = 0
   Sum2 = 0
   Sum3 = 0
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
   
   strFormat = VSP_CalTable(DETAIL_FORMAT, mdblWid, alngX)
  TempBorder = tbNone
'   TempBorder = tbAll
   
   Call GenerateReceiptBody(Vsp, mcolParam, m_cData, TempBorder)
              
   'Call GenerateReceiptFooter(Vsp, mcolParam, TempBorder, m_cData)
           
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   Rs.Close
   Set Rs = Nothing
   Set m_cData = Nothing
   
   genDoc = True
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
'   ErrorObj.RoutineName = RName
'   ErrorObj.ModuleName = MODULE_NAME
'   ErrorObj.LocalErrorMsg = "Run time error."
'   ErrorObj.SystemErrorMsg = mstrErrMsg
'   Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
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
   
'   Vsp.DrawPicture mcolParam("PICTURE"), "0.75in", "0.75in", "100%", "100%"
   
   Vsp.FontSize = 8
   dY(0) = Vsp.CurrentY
   Vsp.TextAlign = taLeftBottom
   strFormat = VSP_CalTable(TITLE_FORMAT, mdblWid, alngX)
   Vsp.TableBorder = tbNone
'   Call Vsp.AddTable(strFormat, "", "|หน้าที่|: " & Vsp.CurrentPage)
'   Call Vsp.AddTable(strFormat, "", "|วันที่พิมพ์|: " & DateToStringExt(mdteDate))
'   Call Vsp.AddTable(strFormat, "", "|เวลา|: " & Right$(DateToStringExtEx(mdteDate), 8))
   Vsp.CurrentY = dY(0)
   
   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
'   Vsp.TextAlign = taCenterBottom
'   Vsp.Paragraph = MODULE_DESC
'   Vsp.Paragraph = ""
'   If mcolParam("REPORT_TYPE") = 4 Or mcolParam("REPORT_TYPE") = 6 Then
'      Vsp.DrawPicture mcolParam("BACK_GROUND"), Vsp.MarginLeft, Vsp.MarginTop, m_ImageWidth, m_ImageHeight, 10
'   End If
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


