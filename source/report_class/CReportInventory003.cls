VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportInventory003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportInventory003"
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
Private m_Details1 As CFieldList
Private m_Left As Double
Private m_YearSeqs As Collection

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_TempCol = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Details1 = New CFieldList
   Set m_YearSeqs = New Collection
   
   Set Rs = New ADODB.Recordset
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
   
   Set Cf = New CReportField

   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "�ѹ����͡���", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "�Ţ����͡���", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "�Ѻ���", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "�����͡", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "�������", ">")
   Call m_Headers0.AddField(Cf)

   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "�ѹ����͡���", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "�Ţ����͡���", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�ӹǹ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�Ҥ�", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "��Ť��", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�ӹǹ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�Ҥ�", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "��Ť��", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�ӹǹ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�Ҥ�", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "��Ť��", ">")
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
   
   FromDate = "� �ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
   ToDate = " �֧ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")
   
   MODULE_DESC = vbCrLf & glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & vbCrLf & _
                                       mcolParam("REPORT_NAME") & vbCrLf & _
                                       FromDate & "" & ToDate & " ʶҹ���Ѵ�� " & EmptyToString(mcolParam("LOCATION_NAME"), "N/A")
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
Dim I As Long
Dim J As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Pi As CPartItem
Dim Pt As CPartType
Dim Total1(100) As Double
Dim Total2(100) As Double
Dim Total3(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim Amt As Double
Dim PartTypes As Collection
Dim PartItems As Collection
Dim Li As CLotItem
Dim TempLi As CLotItem
Dim Sum1 As Double
Dim Count1 As Double
Dim InventoryBals2 As Collection
Dim NewDate As Date
Dim BalanceAccums As Collection

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
'
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

   Set Li = New CLotItem
   Set PartItems = New Collection
   
   Set InventoryBals2 = New Collection
   NewDate = DateAdd("D", -1, mcolParam("FROM_DATE"))

   Set BalanceAccums = New Collection
   If mcolParam("FROM_DATE") > 0 Then
      Call LoadInventoryBalanceEx(Nothing, BalanceAccums, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), "")
      Call glbDaily.CopyBalanceAccum(BalanceAccums, InventoryBals2)
   End If
   Set BalanceAccums = Nothing

   Set PartTypes = New Collection
   Call LoadDistinctPartType(Nothing, PartTypes, -1, -1, "", mcolParam("LOCATION_ID"), mcolParam("PART_TYPE"), mcolParam("PART_GROUP"))

   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
   
      I = 0
      Set BodyCf = New CReportField
      
      For Each Pt In PartTypes
         For J = 1 To UBound(Total2)
            Total2(J) = 0
         Next J

         Call LoadDistinctPartItem(Nothing, PartItems, -1, -1, "", mcolParam("LOCATION_ID"), Pt.PART_TYPE_ID)
         For Each Pi In PartItems
            For J = 1 To UBound(Total1)
               Total1(J) = 0
            Next J
   
            Call m_Details1.ClearField
            Call BodyCf.SetFieldValue(7, "^", Pt.PART_TYPE_NAME & " (" & Pt.PART_TYPE_NO & ")")
            Call m_Details1.AddField(BodyCf)
            Call BodyCf.SetFieldValue(18, "^", Pi.PART_DESC & " (" & Pi.PART_NO & ")" & "    ˹��¹Ѻ�� " & Pi.UNIT_NAME)
            Call m_Details1.AddField(BodyCf)
            
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
            
            '=== �ʹ¡��
            Call m_Details1.ClearField
            J = 0
            For Each HeadCf In m_Headers1.Fields
               J = J + 1
               If J = 2 Then
                  TempStr = "*** �ʹ¡�� ***"
               ElseIf J = 10 Then
                  Set Li = GetLotItem(InventoryBals2, mcolParam("LOCATION_ID") & "-" & Pi.PART_ITEM_ID)
                  Amt = Li.NEW_AMOUNT
                  TempStr = FormatNumber(Amt)
                  Total1(J) = Total1(J) + Amt
               ElseIf J = 11 Then
                  Amt = Li.NEW_PRICE
                  TempStr = FormatNumber(Amt)
                  Total1(J) = Total1(J) + Amt
               ElseIf J = 12 Then
                  Amt = Li.TOTAL_INCLUDE_PRICE  'Li.NEW_PRICE * Li.NEW_AMOUNT
                  TempStr = FormatNumber(Amt)
                  Total1(J) = Total1(J) + Amt
               Else
                  TempStr = ""
               End If
               
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
               Call m_Details1.AddField(BodyCf)
            Next HeadCf
            
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
            '=== �ʹ¡��
            
            '=== ��¡��
            Set Li = New CLotItem
            Li.LOT_ITEM_ID = -1
            Li.FROM_DATE = mcolParam("FROM_DATE")
            Li.TO_DATE = mcolParam("TO_DATE")
            Li.LOCATION_ID = mcolParam("LOCATION_ID")
            Li.PART_ITEM_ID = Pi.PART_ITEM_ID
            Li.COMMIT_FLAG = ""
            Li.OrderBy = 11
            Li.OrderType = 1
            Call Li.QueryData(1, Rs, iCount)
            Set Li = Nothing
            While Not Rs.EOF
               Call m_Details1.ClearField
               
               I = I + 1
               Set Li = New CLotItem
               Call Li.PopulateFromRS(1, Rs)
               
               For J = 1 To m_Headers1.Fields.Count
                  Set HeadCf = m_Headers1.Fields(J)
                  
                  If J = 1 Then
                     TempStr = I
                  ElseIf J = 2 Then
                     TempStr = DateToStringExtEx2(Li.DOCUMENT_DATE)
                  ElseIf J = 3 Then
                     TempStr = Li.DOCUMENT_NO
                  ElseIf J = 4 Then
                     If Li.TX_TYPE = "I" Then
                        Amt = Li.TX_AMOUNT
                     Else
                        Amt = 0
                     End If
                     TempStr = FormatNumber(Amt)
                     Total1(J) = Total1(J) + Amt
                  ElseIf J = 5 Then
                     If Li.TX_TYPE = "I" Then
                        Amt = Li.INCLUDE_UNIT_PRICE
                     Else
                        Amt = 0
                     End If
                     TempStr = FormatNumber(Amt)
                     Total1(J) = Total1(J) + Amt
                  ElseIf J = 6 Then
                     If Li.TX_TYPE = "I" Then
                        Amt = Li.TOTAL_INCLUDE_PRICE
                     Else
                        Amt = 0
                     End If
                     TempStr = FormatNumber(Amt)
                     Total1(J) = Total1(J) + Amt
                  ElseIf J = 7 Then
                     If Li.TX_TYPE = "E" Then
                        Amt = Li.TX_AMOUNT
                     Else
                        Amt = 0
                     End If
                     TempStr = FormatNumber(Amt)
                     Total1(J) = Total1(J) + Amt
                  ElseIf J = 8 Then
                     If Li.TX_TYPE = "E" Then
                        Amt = Li.PREVIOUS_PRICE
                     Else
                        Amt = 0
                     End If
                     TempStr = FormatNumber(Amt)
                     Total1(J) = Total1(J) + Amt
                  ElseIf J = 9 Then
                     If Li.TX_TYPE = "E" Then
                        Amt = Li.TOTAL_INCLUDE_PRICE  'Li.TX_AMOUNT * Li.PREVIOUS_PRICE
                     Else
                        Amt = 0
                     End If
                     TempStr = FormatNumber(Amt)
                     Total1(J) = Total1(J) + Amt
                  ElseIf J = 10 Then
                     Amt = Li.NEW_AMOUNT
                     TempStr = FormatNumber(Amt)
                     Total1(J) = Amt
                  ElseIf J = 11 Then
                     Amt = Li.NEW_PRICE
                     TempStr = FormatNumber(Amt)
                     Total1(J) = Amt
                  ElseIf J = 12 Then
                     If Li.TOTAL_NEW_PRICE <= 0 Then
                        Amt = Li.NEW_AMOUNT * Li.NEW_PRICE
                     Else
                        Amt = Li.TOTAL_NEW_PRICE
                     End If
                     TempStr = FormatNumber(Amt)
                     Total1(J) = Amt
                  Else
                     TempStr = ""
                  End If
               
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
                  Call m_Details1.AddField(BodyCf)
               Next J
               
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
               
               Set Li = Nothing
               Rs.MoveNext
            Wend
            '=== ��¡��
            
            For J = 1 To UBound(Total1)
               Total2(J) = Total2(J) + Total1(J)
            Next J
            
            Call GenerateFooter("���", Pt.PART_TYPE_NAME, Pi.PART_DESC, Total1, m_Details1)
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
         Next Pi
      
         For J = 1 To UBound(Total3)
            Total3(J) = Total3(J) + Total2(J)
         Next J
            
         Call GenerateFooter("���", Pt.PART_TYPE_NAME, "", Total2, m_Details1)
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
      Next Pt
      
      Call GenerateFooter("�ط��", "", "", Total3, m_Details1)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      
      Set BodyCf = Nothing
   Vsp.EndDoc
   Call CloseExportFile(Vsp)

   Set InventoryBals2 = Nothing
   Set Li = Nothing
   Set cData = Nothing
   Set PartTypes = Nothing
   Set PartItems = Nothing
   
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


