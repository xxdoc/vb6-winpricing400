VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportComIncentive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportComIncentive"
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

Private m_HeaderFormat1 As String
Private m_HeaderFormat2 As String
Private m_HeaderText1 As String
Private m_HeaderText2 As String

Private m_Headers0 As CFieldList
Private m_Headers1 As CFieldList
Private m_Details1 As CFieldList


Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Details1 = New CFieldList
   Set Rs = New ADODB.Recordset
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing

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
Dim DocType As Long

   
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField
   
   DocType = Val(mcolParam("DOCUMENT_TYPE"))
   
   Call Cf.SetFieldValue(3, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   
   If DocType = 1 Then
      Call Cf.SetFieldValue(6, "^", "รหัสฟรีแลนซ์", "^")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(13, "^", "ชื่อฟรีแลนซ์", "<")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(15, "^", "รหัสสินค้า", "<")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(15, "^", "ชื่อสินค้า", "<")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(5, "^", "บาท/หน่วย", ">")
      Call m_Headers1.AddField(Cf)
   ElseIf DocType = 2 Then
      Call Cf.SetFieldValue(6, "^", "รหัสฟรีแลนซ์", "^")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(13, "^", "ชื่อฟรีแลนซ์", "<")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(20, "^", "ชื่อลูกค้า", "<")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(15, "^", "รหัสสินค้า", "<")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(15, "^", "ชื่อสินค้า", "<")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(5, "^", "บาท/หน่วย", ">")
      Call m_Headers1.AddField(Cf)
   ElseIf DocType = 3 Or DocType = 4 Then
      Call Cf.SetFieldValue(5, "^", "รหัสฟรีแลนซ์", "^")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(13, "^", "ชื่อฟรีแลนซ์", "<")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(5, "^", "บาท/หน่วย", ">")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(5, "^", "ยอดตั้งแต่", ">")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(5, "^", "ถึงยอด", ">")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(3, "^", "หน่วย", "^")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(5, "^", "ยอดเกิน", ">")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(7, "^", "คิดหน่วยล่ะ(บาท)", ">")
      Call m_Headers1.AddField(Cf)
   End If
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
   Vsp.PaperSize = pprA4 '
   Vsp.ORIENTATION = orLandscape
   Vsp.MarginBottom = 1440
   Vsp.MarginFooter = 1440
   Vsp.MarginHeader = 1440
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 1440
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
'   FromDate = "วันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
   'ToDate = "ถึง " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")

   MODULE_DESC = vbCrLf & glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & vbCrLf & _
                                      mcolParam("HEADER_TEXT") & vbCrLf & FromDate
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
Dim Ci As CCommissionIncentive
Dim Temp_CI As Collection
Dim I As Long
Dim J As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim Amt As Double
Dim Sum1 As Double
Dim Count1 As Double
Dim CalculateWeight As Double
Dim Sp As CSupplier
Dim DOCUMENT_TYPE As Long

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   If Not initDoc Then Exit Function

   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
   
   For J = 1 To UBound(Total1)
      Total1(J) = 0
   Next J

   Set Temp_CI = mcolParam("COMMISSION_INCENTIVE")
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
      I = 0
     
      Set BodyCf = New CReportField
      DOCUMENT_TYPE = mcolParam("DOCUMENT_TYPE")

      Set Ci = New CCommissionIncentive
      If DOCUMENT_TYPE = 1 Then
         For Each Ci In Temp_CI
            I = I + 1
             J = 0
            Call m_Details1.ClearField
   
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = I
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
   '
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = Ci.FREELANCE_CODE
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = Ci.FREELANCE_NAME & " " & Ci.FREELANCE_LASTNAME
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = Ci.PART_ITEM_CODE
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = Ci.PART_ITEM_NAME
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = Ci.INCENTIVE_PER_PACK
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
            '=====
         Next Ci
      ElseIf DOCUMENT_TYPE = 2 Then
         For Each Ci In Temp_CI
            I = I + 1
             J = 0
            Call m_Details1.ClearField
   
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = I
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
   '
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = Ci.FREELANCE_CODE
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = Ci.FREELANCE_NAME & " " & Ci.FREELANCE_LASTNAME
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = Ci.CUSTOMER_NAME
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = Ci.PART_ITEM_CODE
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = Ci.PART_ITEM_NAME
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = Ci.INCENTIVE_PER_PACK
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
            '=====
         Next Ci
      ElseIf DOCUMENT_TYPE = 3 Or DOCUMENT_TYPE = 4 Then
         For Each Ci In Temp_CI
            I = I + 1
             J = 0
            Call m_Details1.ClearField
   
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = I
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
   '
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = Ci.FREELANCE_CODE
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = Ci.FREELANCE_NAME & " " & Ci.FREELANCE_LASTNAME
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = FormatNumber(Ci.INCENTIVE_PER_PACK)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = FormatNumber(Ci.FROM_AMOUNT)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = FormatNumber(Ci.TO_AMOUNT)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = TypeUnitFlag(Ci.UNIT_TYPE)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            

            If Ci.AMOUNT_OVER_FLAG = "Y" Then
               J = J + 1
               Set HeadCf = m_Headers1.Fields(J)
               TempStr = FormatNumber(Ci.TO_AMOUNT)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
               Call m_Details1.AddField(BodyCf)
               
                J = J + 1
               Set HeadCf = m_Headers1.Fields(J)
               TempStr = Ci.RATE_OVER
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
               Call m_Details1.AddField(BodyCf)
           Else
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

            
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
            '=====
         Next Ci
      End If

'      Call GenerateFooter("รวม", "", "", Total1, m_Details1)
'      Call m_Details1.GetString(2, TempStr1, TempStr2)
'      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'      Vsp.FontBold = True
'      Call Vsp.AddTable(strFormat, "", TempStr2)
'      Vsp.FontBold = False

      Set BodyCf = Nothing

   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   Set Ci = Nothing
   
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
      If Val(mcolParam("DOCUMENT_TYPE")) = 13 Then
      Select Case J
      Case 1
         TempStr = TXT1
      Case 2, 3, 4, 5, 6, 7, 14, 15
          TempStr = ""
      Case 13 'ยอดตัน
         TempStr = FormatNumber(Tot(J), 3)
      Case Else
          TempStr = FormatNumber(Tot(J), 3)
      End Select
   ElseIf Val(mcolParam("DOCUMENT_TYPE")) = 14 Then
      Select Case J
      Case 1
         TempStr = TXT1
      Case 2, 3, 4, 5, 6, 7, 8, 15, 16
         TempStr = ""
      Case 14 'ยอดตัน
         TempStr = FormatNumber(Tot(J), 2)
      Case Else
         TempStr = FormatNumber(Tot(J), 0)
      End Select
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
'   Vsp.Paragraph = ""
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




