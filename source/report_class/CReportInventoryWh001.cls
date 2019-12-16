VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportInventoryWh001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportInventoryWh001"
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

Private m_PartItemTxs As Collection

Private m_CollLotItemWhImportBal As Collection
Private m_CollLotItemWhExportBal As Collection

Private DOCUMENT_TYPE As Long

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
   
   Set m_CollLotItemWhImportBal = New Collection
   Set m_CollLotItemWhExportBal = New Collection
   
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
   
   Set m_CollLotItemWhImportBal = Nothing
   Set m_CollLotItemWhExportBal = Nothing
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

   Call Cf.SetFieldValue(7, "^", "", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(11, "^", "รับเข้า", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(11, "^", "จ่ายออก", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "คงเหลือ", ">")
   Call m_Headers0.AddField(Cf)

   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "วันที่เอกสาร", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "เลขที่เอกสาร", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "LOT", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ถัง", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ล๊อค", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "พาเลท", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "LOT", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ถัง", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ล๊อค", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "พาเลท", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
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
Dim MaxDate  As Date

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

   If mcolParam("DOCUMENT_TYPE") = 2000 Then
      DOCUMENT_TYPE = 14
      Call LoadInventoryDocMaxDateBalance(Nothing, Nothing, 15, MaxDate)
   ElseIf mcolParam("DOCUMENT_TYPE") = 2001 Then
      DOCUMENT_TYPE = 13
      Call LoadInventoryDocMaxDateBalance(Nothing, Nothing, 16, MaxDate)
   ElseIf mcolParam("DOCUMENT_TYPE") = 0 Then
     DOCUMENT_TYPE = -1
   End If
   
   Call LoadPartStockFromLotItemWh(Nothing, m_CollLotItemWhImportBal, MaxDate, mcolParam("FROM_DATE") - 1, mcolParam("PART_TYPE"), , "I", 3, 18, DOCUMENT_TYPE, 2, mcolParam("LOCATION_ID")) 'ยอดรับเข้ายกมา
   Call LoadPartStockFromLotItemWh(Nothing, m_CollLotItemWhExportBal, MaxDate, mcolParam("FROM_DATE") - 1, mcolParam("PART_TYPE"), , "E", 3, 18, DOCUMENT_TYPE, , mcolParam("LOCATION_ID")) ''ยอดจ่ายออกยกมา
      
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
Private Sub GetBalance(PartNo As String, WeightPerPack As Long, ByRef ImportBal As Double, ByRef ExportBal As Double, ByRef TotalBal As Double)
Dim Bi As CLotItemWH
Dim bE As CLotItemWH
Dim Bl As Double
      Set Bi = GetObject("CLotItemWH", m_CollLotItemWhImportBal, Trim(PartNo) & "-" & Trim(str(WeightPerPack)) & "-" & "I", False)
      If Not Bi Is Nothing Then
         ImportBal = Bi.CAPACITY_AMOUNT
      End If
      
      Set bE = GetObject("CLotItemWH", m_CollLotItemWhExportBal, Trim(PartNo) & "-" & Trim(str(WeightPerPack)) & "-" & "E", False)
      If Not bE Is Nothing Then
         ExportBal = bE.CAPACITY_AMOUNT
         Bl = ImportBal - bE.CAPACITY_AMOUNT
      End If
      
      TotalBal = Bl
End Sub
Private Function setNewLine(IsNew As Boolean, data As String) As String
If IsNew Then
   setNewLine = data
Else
   setNewLine = ""
End If
End Function
Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim RName As String
Dim I As Long
Dim J As Long
Dim J2 As Long
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
Dim OldID As Long
Dim LIW As CLotItemWH
Dim k As Long

Dim PrevKey1 As String
Dim PrevKey2  As String
Dim PrevKey3  As String
Dim PrevKey4  As String
Dim TempTxType As String
Dim TempDate As String
Dim TempDate2 As String
Dim TempKey   As String
Dim ImportBal As Double
Dim ExportBal As Double
Dim TotalBal As Double

Dim ShowFlag As Boolean
Dim IsNewLine As Boolean
Dim IsNewLine2 As Boolean
Dim IsNewLine3 As Boolean

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
      k = 0
      Set BodyCf = New CReportField

      Set LIW = New CLotItemWH
      LIW.LOT_ITEM_WH_ID = -1
      LIW.FROM_DATE = mcolParam("FROM_DATE")
      LIW.TO_DATE = mcolParam("TO_DATE")
      LIW.LOCATION_ID = mcolParam("LOCATION_ID")
      LIW.PART_TYPE = mcolParam("PART_TYPE")
      LIW.PART_GROUP_ID = mcolParam("PART_GROUP")
      LIW.PART_NO = mcolParam("PART_NO")
      LIW.LOT_NO = mcolParam("LOT_NO")
'      LIW.VERIFY_FLAG = "Y" 'จะดึงเฉพาะตัวที่ผ่านการตรวจสอบแล้วเท่านั้น
      LIW.OrderBy = 1
      LIW.OrderType = 1
      Call LIW.QueryData(17, Rs, iCount, False)
       
      I = 0
      Set BodyCf = New CReportField
      Set LIW = New CLotItemWH
      Dim Tot1 As Double
      Dim Tot2 As Double
      While Not Rs.EOF
         Call LIW.PopulateFromRS(17, Rs)
        
         Tot1 = 0
         Tot2 = 0
         J = 0

         Call m_Details1.ClearField
         If PrevKey1 <> Trim(LIW.PART_NO) Then

            If TempTxType = "I" Then
               Call GenerateFooter("", "", "", "รวม Lot รับเข้า", "", Total1, m_Details1, TempTxType, DOCUMENT_TYPE)
            ElseIf TempTxType = "E" Then
               Call GenerateFooter("", "", "", "", "รวม Lot จ่ายออก", Total1, m_Details1, TempTxType, DOCUMENT_TYPE)
            End If
            
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
            
            
            For J2 = 1 To UBound(Total1)
               Total1(J2) = 0
            Next J2
            
            If TempTxType = "I" Then
               Call GenerateFooter("", "รวม รับเข้า", "", "", "", Total2, m_Details1, TempTxType, DOCUMENT_TYPE)
            ElseIf TempTxType = "E" Then
               Call GenerateFooter("", "รวม จ่ายออก", "", "", "", Total2, m_Details1, TempTxType, DOCUMENT_TYPE)
            End If
         
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False

         TempTxType = LIW.TX_TYPE
            
         For J2 = 1 To UBound(Total1)
            Total2(J2) = 0
         Next J2
         
            Call GenerateSubHeader(LIW.PART_TYPE_NAME & " (" & LIW.PART_TYPE_NO & ")", LIW.PART_DESC & " (" & LIW.PART_NO & ")", "", Total1, m_Details1, 0)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
         
            Call GetBalance(LIW.PART_NO, LIW.WEIGHT_PER_PACK, ImportBal, ExportBal, TotalBal)
            Call GenerateBalance("", "ยกมา", "", m_Details1, 0, ImportBal, ExportBal, TotalBal)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
         
            TempTxType = LIW.TX_TYPE
            I = 0
        End If
      
      TempKey = DateToStringExtEx2(LIW.DOCUMENT_DATE) & "-" & LIW.DOCUMENT_NO & "-" & LIW.LOT_NO & "-" & LIW.BIN_NAME & "-" & LIW.LOCK_NAME
      If PrevKey3 <> TempKey Then
         IsNewLine2 = True
      Else
        IsNewLine2 = False
      End If
      PrevKey3 = DateToStringExtEx2(LIW.DOCUMENT_DATE) & "-" & LIW.DOCUMENT_NO & "-" & LIW.LOT_NO & "-" & LIW.BIN_NAME & "-" & LIW.LOCK_NAME
   
   
     If IsNewLine2 And I > 0 Then
            If TempTxType = "I" Then
               Call GenerateFooter("", "", "", "รวม Lot รับเข้า", "", Total1, m_Details1, TempTxType, DOCUMENT_TYPE)
            ElseIf TempTxType = "E" Then
               Call GenerateFooter("", "", "", "", "รวม Lot จ่ายออก", Total1, m_Details1, TempTxType, DOCUMENT_TYPE)
            End If
         
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
            
         For J2 = 1 To UBound(Total1)
            Total1(J2) = 0
         Next J2
      End If
           
      TempKey = DateToStringExtEx2(LIW.DOCUMENT_DATE) & "-" & LIW.DOCUMENT_NO
         If PrevKey2 <> TempKey Then
            IsNewLine = True
            I = I + 1
         Else
           IsNewLine = False
         End If
         PrevKey2 = DateToStringExtEx2(LIW.DOCUMENT_DATE) & "-" & LIW.DOCUMENT_NO
         
         If IsNewLine And I > 1 Then
            If TempTxType = "I" Then
               Call GenerateFooter("", "รวม รับเข้า", "", "", "", Total2, m_Details1, TempTxType, DOCUMENT_TYPE)
            ElseIf TempTxType = "E" Then
               Call GenerateFooter("", "รวม จ่ายออก", "", "", "", Total2, m_Details1, TempTxType, DOCUMENT_TYPE)
            End If
         
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False

         TempTxType = LIW.TX_TYPE
            
         For J2 = 1 To UBound(Total1)
            Total2(J2) = 0
         Next J2
      End If
      
         TempKey = DateToStringExtEx2(LIW.DOCUMENT_DATE)
         If PrevKey4 <> TempKey Then
            IsNewLine3 = True
            TempDate2 = TempDate
            TempDate = DateToStringExtEx2(LIW.DOCUMENT_DATE)
'            I = I + 1
         Else
           IsNewLine3 = False
         End If
         PrevKey4 = DateToStringExtEx2(LIW.DOCUMENT_DATE)
         
         
         If IsNewLine3 And I > 1 Then
'            If TempTxType = "I" Then
'               Call GenerateFooter("", "รวม วันที่", TempDate2, "", "", Total3, m_Details1, "A", DOCUMENT_TYPE)
'            ElseIf TempTxType = "E" Then
'               Call GenerateFooter("", "รวม วันที่", TempDate2, "", "", Total3, m_Details1, "A", DOCUMENT_TYPE)
'            End If
          Call GenerateFooter("", "รวม วันที่", TempDate2, "", "", Total3, m_Details1, "A", DOCUMENT_TYPE)
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
         
         Vsp = ""
            
         For J2 = 1 To UBound(Total1)
            Total3(J2) = 0
         Next J2
      End If
   
      Call m_Details1.ClearField
      PrevKey1 = Trim(LIW.PART_NO)
      
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = setNewLine(IsNewLine, str(I))
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)

         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = setNewLine(IsNewLine, DateToStringExtEx2(LIW.DOCUMENT_DATE))
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)

         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = setNewLine(IsNewLine, LIW.DOCUMENT_NO)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)

         If LIW.TX_TYPE = "I" Then
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
           TempStr = setNewLine(IsNewLine2, LIW.LOT_NO)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = setNewLine(IsNewLine2, LIW.BIN_NAME)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = setNewLine(IsNewLine2, LIW.LOCK_NAME)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
   
             J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = LIW.PALLET_NAME
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = LIW.CAPACITY_AMOUNT
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            Total1(J) = Total1(J) + LIW.CAPACITY_AMOUNT
            Total2(J) = Total2(J) + LIW.CAPACITY_AMOUNT
            Total3(J) = Total3(J) + LIW.CAPACITY_AMOUNT
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = ""
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = ""
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = ""
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
             J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = ""
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = ""
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
         End If
      
       If LIW.TX_TYPE = "E" Then
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = setNewLine(IsNewLine2, LIW.LOT_NO)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = setNewLine(IsNewLine2, LIW.BIN_NAME)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = setNewLine(IsNewLine2, LIW.LOCK_NAME)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = LIW.PALLET_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = LIW.CAPACITY_AMOUNT
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         Total1(J) = Total1(J) + LIW.CAPACITY_AMOUNT
         Total2(J) = Total2(J) + LIW.CAPACITY_AMOUNT
         Total3(J) = Total3(J) + LIW.CAPACITY_AMOUNT
       End If
      
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         If LIW.TX_TYPE = "I" Then
            TotalBal = TotalBal + LIW.CAPACITY_AMOUNT
         ElseIf LIW.TX_TYPE = "E" Then
            TotalBal = TotalBal - LIW.CAPACITY_AMOUNT
         End If
         TempStr = FormatNumber(TotalBal, 0)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
      
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         
         Rs.MoveNext
      Wend

            If TempTxType = "I" Then
               Call GenerateFooter("", "", "", "รวม Lot รับเข้า", "", Total1, m_Details1, TempTxType, DOCUMENT_TYPE)
            ElseIf TempTxType = "E" Then
               Call GenerateFooter("", "", "", "", "รวม Lot จ่ายออก", Total1, m_Details1, TempTxType, DOCUMENT_TYPE)
            End If
            
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
            
            
            If TempTxType = "I" Then
               Call GenerateFooter("", "รวม รับเข้า", "", "", "", Total2, m_Details1, TempTxType, DOCUMENT_TYPE)
            ElseIf TempTxType = "E" Then
               Call GenerateFooter("", "รวม จ่ายออก", "", "", "", Total2, m_Details1, TempTxType, DOCUMENT_TYPE)
            End If
         
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
         

      Call GenerateFooter("", "รวม วันที่", TempDate2, "", "", Total3, m_Details1, "A", DOCUMENT_TYPE)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
            
            
            
      Set LIW = Nothing
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

Private Sub GenerateFooter(TXT1 As String, TXT2 As String, TXT3 As String, TXT4 As String, TXT5 As String, Tot() As Double, Details As CFieldList, TX_TYPE As String, DOCUMENT_TYPE As Long)
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
         TempStr = TXT4
      ElseIf J = 9 Then
         TempStr = TXT5
      ElseIf J = 8 Then
         If TX_TYPE = "I" Or TX_TYPE = "A" Then
            If DOCUMENT_TYPE = 14 Then
               TempStr = FormatNumber(Tot(J), 0)
            Else
               TempStr = FormatNumber(Tot(J), -1)
            End If
         Else
            TempStr = ""
         End If
      ElseIf J = 13 Then
         If TX_TYPE = "E" Or TX_TYPE = "A" Then
             If DOCUMENT_TYPE = 14 Then
               TempStr = FormatNumber(Tot(J), 0)
            Else
               TempStr = FormatNumber(Tot(J), -1)
            End If
         Else
            TempStr = ""
         End If
      Else
         TempStr = ""
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub

Private Sub GenerateBalance(TXT1 As String, TXT2 As String, TXT3 As String, Details As CFieldList, DateCount As Long, ImportBal As Double, ExportBal As Double, TotalBal As Double)
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
      ElseIf J = 8 Then
         TempStr = FormatNumber(ImportBal, 0)
      ElseIf J = 13 Then
         TempStr = FormatNumber(ExportBal, 0)
      ElseIf J = 14 Then
         TempStr = FormatNumber(TotalBal, 0)
      Else
         TempStr = ""
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



