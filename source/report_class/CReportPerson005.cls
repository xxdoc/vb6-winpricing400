VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportPerson005"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAdmin005"
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
Private Rs1 As ADODB.Recordset
Private Rs2 As ADODB.Recordset
Private Rs3 As ADODB.Recordset
Private Rs5 As ADODB.Recordset

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
      
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(5, "^", "ชื่อ-นามสกุล", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "เดือน", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ปี", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(5, "^", "รายการ", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "รับ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "หัก", ">")
   Call m_Headers1.AddField(Cf)
   
   Set Cf = Nothing
            
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
      Dim Em As CEmployee
      
   MODULE_DESC = vbCrLf & glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & vbCrLf & _
                                                                          mcolParam("REPORT_NAME") & vbCrLf & _
                                       "  จาก เดือน " & frmSummaryReport.cboGeneric(2) & "  ปี  " & frmSummaryReport.cboGeneric(3) & _
                                               "  ถึง  เดือน  " & frmSummaryReport.cboGeneric(4) & "  ปี  " & frmSummaryReport.cboGeneric(5)
   
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
Dim Amt As Double
Dim Cm As CEmployee
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
Dim PrevKey2 As String
Dim Sl As CSliptSalary
Dim Sum1 As String
Dim Sum2 As String
Dim Sum3 As String
Dim Sum4 As String
Dim countM As Long
Dim SumSalary As Double
Dim SumAdd As Double
Dim SumSub As Double
Dim Sumtot As Double
   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set Rs = New ADODB.Recordset
   Set Rs1 = New ADODB.Recordset
   Set Rs2 = New ADODB.Recordset
   Set Rs3 = New ADODB.Recordset
   Set Rs5 = New ADODB.Recordset
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
      I = 0
      Set BodyCf = New CReportField

   Set Cm = New CEmployee
      Cm.EMP_CODE = mcolParam("EMP_NO")
      Cm.NAME = mcolParam("EMP_NAME")
      Cm.LASTNAME = mcolParam("EMP_LAST_NAME")
      Cm.CURRENT_POSITION = mcolParam("EMP_POSITION")
      Cm.QueryFlag = 0
    Cm.OrderBy = mcolParam("ORDER_BY")
    Cm.OrderType = mcolParam("ORDER_TYPE")
      
      Call glbDaily.QueryEmployee(Cm, Rs, iCount, IsOK, glbErrorLog)
   I = 0
Sum4 = 0
      While Not Rs.EOF
         
         Call m_Details1.ClearField
         Call Cm.PopulateFromRS(1, Rs)
    Set Sl = New CSliptSalary
    Sl.FROM_MONTH = mcolParam("FROM_MONTH")
    Sl.FROM_YEAR = frmSummaryReport.cboGeneric(3).Text
    Sl.TO_MONTH = mcolParam("TO_MONTH")
    Sl.TO_YEAR = frmSummaryReport.cboGeneric(5).Text
    Sl.EMP_ID = Cm.EMP_ID
    
    Sl.OrderBy = 1
    Sl.OrderType = 1
    Sl.QueryFlag = 1
Call glbDaily.QuerySliptSalary(Sl, Rs1, iCount, IsOK, glbErrorLog)
      
If iCount <> 0 Then
          While Not Rs1.EOF
    Call Sl.PopulateFromRS(Rs1)
 '''''Debug.Print (Sl.EMP_ID)
Sum1 = 0
Sum2 = 0
Sum3 = 0
Call glbDaily.QuerySliptSalary2(Sl, Rs1, iCount, IsOK, glbErrorLog)
              Vsp.TableBorder = tbAll
     Call m_Details1.ClearField
        I = I + 1
          Set HeadCf = m_Headers1.Fields(1)
         TempStr = I
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
    
    Set HeadCf = m_Headers1.Fields(2)
         TempStr = Replace(Cm.NAME, vbCrLf, "") & " " & Cm.LASTNAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(3)
         TempStr = Sl.MONTH_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Set HeadCf = m_Headers1.Fields(4)
         TempStr = Sl.YEAR_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

    Set HeadCf = m_Headers1.Fields(5)
         TempStr = "เงินเดือนประจำ"
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(6)
         TempStr = FormatNumber(Sl.SALARY)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         SumSalary = SumSalary + Sl.SALARY
         Set HeadCf = m_Headers1.Fields(7)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
    
            Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)

    Dim SA As CSliptAdd
    For Each SA In Sl.SliptAdd
         Call m_Details1.ClearField
          Set HeadCf = m_Headers1.Fields(1)
         TempStr = ""
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
         TempStr = SA.MONTHLY_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(6)
         TempStr = FormatNumber(SA.MONTHLY_AMOUNT)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Sum1 = Sum1 + SA.MONTHLY_AMOUNT
         Set HeadCf = m_Headers1.Fields(7)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
    
            Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)

    Next SA
    Vsp.FontBold = True
Call m_Details1.ClearField
          
          Set HeadCf = m_Headers1.Fields(1)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

          Set HeadCf = m_Headers1.Fields(2)
         TempStr = "รวม"
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Set HeadCf = m_Headers1.Fields(3)
         TempStr = Sl.MONTH_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Set HeadCf = m_Headers1.Fields(4)
         TempStr = Sl.YEAR_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(5)
         TempStr = "รายการรับ"
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Set HeadCf = m_Headers1.Fields(6)
         TempStr = FormatNumber(Sum1 + Sl.SALARY)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Set HeadCf = m_Headers1.Fields(7)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
        
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
Vsp.FontBold = False
    Dim SB As CSliptSub
    For Each SB In Sl.SliptSub
            Call m_Details1.ClearField
          Set HeadCf = m_Headers1.Fields(1)
         TempStr = ""
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
         TempStr = SB.MONTHLY_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(6)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(7)
         TempStr = FormatNumber(SB.MONTHLY_AMOUNT)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      Sum2 = Sum2 + SB.MONTHLY_AMOUNT
        Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
    
    Next SB
    Vsp.FontBold = True
 Call m_Details1.ClearField
        
        Set HeadCf = m_Headers1.Fields(1)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(2)
         TempStr = "รวม"
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(3)
         TempStr = Sl.MONTH_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Set HeadCf = m_Headers1.Fields(4)
         TempStr = Sl.YEAR_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(5)
         TempStr = "รายการหัก"
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Set HeadCf = m_Headers1.Fields(6)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
  Set HeadCf = m_Headers1.Fields(7)
         TempStr = FormatNumber(Sum2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
        
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
        
        Sum3 = Sl.SALARY + Val(Sum1) - Val(Sum2)
                
        Call m_Details1.ClearField
        
        Set HeadCf = m_Headers1.Fields(1)
         TempStr = ""
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
         TempStr = "รวม"
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Set HeadCf = m_Headers1.Fields(5)
         TempStr = "เงินเดือนสุทธิ"
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Set HeadCf = m_Headers1.Fields(6)
         TempStr = FormatNumber(Sum3)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
        Set HeadCf = m_Headers1.Fields(7)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
        
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
      SumAdd = SumAdd + Val(Sum1)
      SumSub = SumSub + Val(Sum2)
  Sumtot = Sumtot + Val(Sum3)
  Vsp.FontBold = False
   Rs1.MoveNext
      Wend
                           End If ' Icount
         Rs.MoveNext
      Wend
       Vsp = ""
         Vsp = ""
         Vsp = ""
       Call m_Details1.ClearField
        Vsp.FontBold = True
        Set HeadCf = m_Headers1.Fields(1)
         TempStr = ""
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
         TempStr = "รวม"
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Set HeadCf = m_Headers1.Fields(5)
         TempStr = "เงินเดือนประจำทั้งหมด"
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Set HeadCf = m_Headers1.Fields(6)
         TempStr = FormatNumber(SumSalary)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
        Set HeadCf = m_Headers1.Fields(7)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
        
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
         
   '  COUNT SUM


                 Dim a2_6 As CMonthlyAdd
      Set a2_6 = New CMonthlyAdd
      Dim ItemCount As Long
      Dim Status As Boolean
      a2_6.MONTHLY_ADD_ID = -1
      Status = a2_6.QueryData(Rs2, ItemCount)
     Dim Sum_A() As Long
     ReDim Sum_A(ItemCount)
    J = 1
          While Not Rs2.EOF
    Call a2_6.PopulateFromRS(Rs2)
    Set Sl = New CSliptSalary
Sl.MONTHLY_ADD_ID = a2_6.MONTHLY_ADD_ID
    
    If J <= ItemCount Then
       Set Cm = New CEmployee
      Cm.EMP_CODE = mcolParam("EMP_NO")
      Cm.NAME = mcolParam("EMP_NAME")
      Cm.LASTNAME = mcolParam("EMP_LAST_NAME")
      Cm.CURRENT_POSITION = mcolParam("EMP_POSITION")
      Cm.QueryFlag = 0
    Cm.OrderBy = mcolParam("ORDER_BY")
    Cm.OrderType = mcolParam("ORDER_TYPE")
      Call glbDaily.QueryEmployee(Cm, Rs, iCount, IsOK, glbErrorLog)
      While Not Rs.EOF
         
         Call m_Details1.ClearField
         Call Cm.PopulateFromRS(1, Rs)

    Sl.FROM_MONTH = mcolParam("FROM_MONTH")
    Sl.FROM_YEAR = frmSummaryReport.cboGeneric(3).Text
    Sl.TO_MONTH = mcolParam("TO_MONTH")
    Sl.TO_YEAR = frmSummaryReport.cboGeneric(5).Text
    Sl.EMP_ID = Cm.EMP_ID
Call glbDaily.QuerySliptSalary1(Sl, Rs1, iCount, IsOK, glbErrorLog)
     
       Call Sl.PopulateFromRS1(Rs1)
     Sum_A(J) = Sum_A(J) + Sl.SUM_ADD
   
   Rs.MoveNext
      Wend
      J = J + 1
      End If
      Rs2.MoveNext
      Wend
      
      Set Rs2 = Nothing
      Set Rs2 = New ADODB.Recordset
      Set a2_6 = New CMonthlyAdd
      a2_6.MONTHLY_ADD_ID = -1
      Status = a2_6.QueryData(Rs2, ItemCount)
        J = 1
    
    While Not Rs2.EOF
    
    
    Call a2_6.PopulateFromRS(Rs2)
'        For j = 1 To j = ItemCount
        
        Vsp.FontBold = True
        Call m_Details1.ClearField
        Set HeadCf = m_Headers1.Fields(1)
         TempStr = ""
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
         TempStr = "รวม"
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Set HeadCf = m_Headers1.Fields(5)
         TempStr = a2_6.MONTHLY_ADD_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(6)
         TempStr = FormatNumber(Sum_A(J))
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
        Set HeadCf = m_Headers1.Fields(7)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
        
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
 '     Next j
   J = J + 1
      
      Rs2.MoveNext
      Wend
      Vsp.FontBold = False
      Vsp.FontBold = True
      Call m_Details1.ClearField
        
        Set HeadCf = m_Headers1.Fields(1)
         TempStr = ""
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
         TempStr = "รวม"
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Set HeadCf = m_Headers1.Fields(5)
         TempStr = "ส่วนบวกเงินเดือนทั้งหมด"
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Set HeadCf = m_Headers1.Fields(6)
         TempStr = FormatNumber(SumAdd + SumSalary)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
        Set HeadCf = m_Headers1.Fields(7)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
        
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
       Vsp.FontBold = False
   ' Sum Detail sub
              Dim a2_7 As CMonthlySub
      Set a2_7 = New CMonthlySub
      
      a2_7.MONTHLY_SUB_ID = -1
      Status = a2_7.QueryData(Rs2, ItemCount)
     Dim Sum_S() As Long
     ReDim Sum_S(ItemCount)
    J = 1
          While Not Rs2.EOF
    Call a2_7.PopulateFromRS(Rs2)
    Set Sl = New CSliptSalary
Sl.MONTHLY_SUB_ID = a2_7.MONTHLY_SUB_ID
    
    If J <= ItemCount Then
       Set Cm = New CEmployee
      Cm.EMP_CODE = mcolParam("EMP_NO")
      Cm.NAME = mcolParam("EMP_NAME")
      Cm.LASTNAME = mcolParam("EMP_LAST_NAME")
      Cm.CURRENT_POSITION = mcolParam("EMP_POSITION")
      Cm.QueryFlag = 0
    Cm.OrderBy = mcolParam("ORDER_BY")
    Cm.OrderType = mcolParam("ORDER_TYPE")
      Call glbDaily.QueryEmployee(Cm, Rs, iCount, IsOK, glbErrorLog)
      While Not Rs.EOF
         
         Call m_Details1.ClearField
         Call Cm.PopulateFromRS(1, Rs)

    Sl.FROM_MONTH = mcolParam("FROM_MONTH")
    Sl.FROM_YEAR = frmSummaryReport.cboGeneric(3).Text
    Sl.TO_MONTH = mcolParam("TO_MONTH")
    Sl.TO_YEAR = frmSummaryReport.cboGeneric(5).Text
    Sl.EMP_ID = Cm.EMP_ID
Call glbDaily.QuerySliptSalary1_1(Sl, Rs1, iCount, IsOK, glbErrorLog)
     
       Call Sl.PopulateFromRS2(Rs1)
     Sum_S(J) = Sum_S(J) + Sl.SUM_SUB

   Rs.MoveNext
      Wend
      J = J + 1
      End If
      Rs2.MoveNext
      Wend
      
      Set Rs2 = Nothing
      Set Rs2 = New ADODB.Recordset
      Set a2_7 = New CMonthlySub
      a2_7.MONTHLY_SUB_ID = -1
      Status = a2_7.QueryData(Rs2, ItemCount)
        J = 1
    Vsp.FontBold = True
    While Not Rs2.EOF
    
    
    Call a2_7.PopulateFromRS(Rs2)
'        For j = 1 To j = ItemCount
        
        
        Call m_Details1.ClearField
        Set HeadCf = m_Headers1.Fields(1)
         TempStr = ""
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
         TempStr = "รวม"
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Set HeadCf = m_Headers1.Fields(5)
         TempStr = a2_7.MONTHLY_SUB_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(6)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
        Set HeadCf = m_Headers1.Fields(7)
         TempStr = FormatNumber(Sum_S(J))
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
        
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
 '     Next j
   J = J + 1
      
      Rs2.MoveNext
      Wend
       
      Call m_Details1.ClearField
        
        Set HeadCf = m_Headers1.Fields(1)
         TempStr = ""
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
         TempStr = "รวม"
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Set HeadCf = m_Headers1.Fields(5)
         TempStr = "ส่วนหักเงินเดือนทั้งหมด"
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Set HeadCf = m_Headers1.Fields(6)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
        Set HeadCf = m_Headers1.Fields(7)
         TempStr = FormatNumber(SumSub)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
        
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
      
      Call m_Details1.ClearField
      
        Set HeadCf = m_Headers1.Fields(1)
         TempStr = ""
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
         TempStr = "รวม"
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Set HeadCf = m_Headers1.Fields(5)
         TempStr = "เงินเดือนสุทธิทั้งหมด"
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Set HeadCf = m_Headers1.Fields(6)
         TempStr = FormatNumber(Sumtot)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
        Set HeadCf = m_Headers1.Fields(7)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
        
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = True
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
      ElseIf (J = 4) Or (J = 5) Then
         TempStr = ""
      ElseIf J = 6 Then
         TempStr = Tot(J)
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
   Vsp.TextAlign = taLeftBottom
   strFormat = VSP_CalTable(TITLE_FORMAT, mdblWid, alngX)
   Vsp.TableBorder = tbNone
   Call Vsp.AddTable(strFormat, "", "|หน้าที่|: " & Vsp.CurrentPage)
   Call Vsp.AddTable(strFormat, "", "|วันที่พิมพ์|: " & DateToStringExt(mdteDate))
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




