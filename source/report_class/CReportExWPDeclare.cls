VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportExWPDeclare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportExWPDeclare"
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

Private collUserAccount  As Collection
Private collUser As Collection
Public m_TempCol As Collection
Private m_HeaderFormat1 As String
Private m_HeaderFormat2 As String
Private m_HeaderText1 As String
Private m_HeaderText2 As String

Private m_Headers0 As CFieldList
Private m_Headers1 As CFieldList
Private m_Details1 As CFieldList
Private m_Left As Double

Private DocNo As String
Private DocDesc As String
Private DocDate As String
Private FromDate As String
Private ToDate As String

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_TempCol = New Collection
   Set collUserAccount = New Collection
   Set collUser = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Details1 = New CFieldList
   
   Set Rs = New ADODB.Recordset
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   Set collUserAccount = Nothing
   Set collUser = Nothing
   
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
   
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField

If mcolParam("AREA") = 1 Then
   Call Cf.SetFieldValue(7, "^", "�ӴѺ���", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(10, "^", "�����Թ���", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(20, "^", "�����Թ���", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(11, "^", "�Ҥ�/�ا", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(20, "^", "�������Թ���", "^")
   Call m_Headers1.AddField(Cf)
ElseIf mcolParam("AREA") = 2 Then
   Call Cf.SetFieldValue(7, "^", "�ӴѺ���", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(25, "^", "�����١���", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(25, "^", "�����١���", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(25, "^", "����ʶҹ���", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(25, "^", "ʶҹ���Ѵ��", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(10, "^", "��Ң���/˹���", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(10, "^", "˹���", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(10, "^", "���˹ѡ(��.)", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(10, "^", "������ö����", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(10, "^", "�Դ�١���/˹���", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(10, "^", "˹���", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(10, "^", "���˹ѡ(��.)", ">")
   Call m_Headers1.AddField(Cf)
ElseIf mcolParam("AREA") = 3 Or mcolParam("AREA") = 5 Then
   Call Cf.SetFieldValue(7, "^", "�ӴѺ���", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(10, "^", "�����١���", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(25, "^", "�����١���", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(25, "^", "�����Թ���", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(25, "^", "�����Թ���", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(10, "^", "�.�./˹���", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(9, "^", "˹���", ">")
   Call m_Headers1.AddField(Cf)
ElseIf mcolParam("AREA") = 4 Then
   Call Cf.SetFieldValue(7, "^", "�ӴѺ���", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(10, "^", "�����١���", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(25, "^", "�����١���", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(25, "^", "����ʶҹ���", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(25, "^", "ʶҹ���Ѵ��", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(10, "^", "�.�./˹���", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(9, "^", "˹���", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(9, "^", "���˹ѡ(��.)", ">")
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
Dim ReportHeader As String
Dim DeclareCountStr As String
   
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
   
   Call LoadUserAccountByName(Nothing, collUserAccount)
   
   
   DocNo = "�Ţ����͡��� : " & EmptyToString(mcolParam("DOCUMENT_NO"), "N/A")
   DocDesc = "��������´�͡��� : " & EmptyToString(mcolParam("DOCUMENT_DESC"), "N/A")
   DocDate = "��С�� � �ѹ��� : " & EmptyToString(DateToStringExtEx2(mcolParam("DOCUMENT_DATE")), "N/A")
   FromDate = "�ռ� �ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
   ToDate = "����ش �ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")

If Val(mcolParam("DECLARE_COUNT")) > 0 Then
   DeclareCountStr = " ���駷�� " & Val(mcolParam("DECLARE_COUNT"))
   If Val(mcolParam("DECLARE_COUNT")) > 1 Then
      DeclareCountStr = DeclareCountStr & " (�������)"
   End If
   
   DocDate = DocDate & DeclareCountStr
End If
If mcolParam("AREA") = 1 Then
   ReportHeader = "㺻�С���Ҥ�������ѵ��������Ե��Ҿ"
ElseIf mcolParam("AREA") = 2 Then
   ReportHeader = "㺻�С���ҤҤ�Ң���������Ե��Ҿ"
ElseIf mcolParam("AREA") = 3 Then
   ReportHeader = "㺻�С����ǹŴ�Ҥ�������ѵ��������Ե��Ҿ"
ElseIf mcolParam("AREA") = 4 Then
   ReportHeader = "㺻�С����ǹŴ�ҤҤ�Ң���������Ե��Ҿ"
ElseIf mcolParam("AREA") = 5 Then
   ReportHeader = "㺻�С����ǹŴ������Ҥ�������ѵ��������Ե��Ҿ"
End If

   MODULE_DESC = vbCrLf & ReportHeader
   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))
   
   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
      mstrErrMsg = "Error GetDateTime Error."
      Exit Function
   End If
   mdteDate = InternalDateToDate(strDate)
   initDoc = True
End Function

Private Sub getUser(Key As String, UserID As String)
Dim tmpUser As CUser
Dim Keys As String
Keys = Key & "-" & Trim(UserID)
 Set tmpUser = GetObject("CUser", collUser, Trim(Keys), False)
      If tmpUser Is Nothing Then
          Set tmpUser = New CUser
          tmpUser.GROUP_ID = Key
          tmpUser.USER_NAME = UserID
          
         Call collUser.add(tmpUser, Trim(Keys))
         Set tmpUser = Nothing
      End If


'   If User <> UserCompare And Len(UserCompare) > 0 Then
'     If User = "" Then
'       User = UserCompare
'     Else
'       User = User & "," & UserCompare
'     End If
'   End If
'   getUser = collTemp
End Sub

Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim RName As String
Dim EWPI As CExWorksPriceItem
Dim EDCI As CExDeliveryCostItem
Dim EPPI As CExPromotionPartItem
Dim EPDI As CExPromotionDlcItem
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
Dim DOCUMENT_TYPE As Long
Dim MODIFY_BY As Long
Dim CREATE_BY As String
Dim userApproved As String
Dim userVerify As String
Dim PrevKey As String

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
'   If mcolParam("AREA") <> 1 Then
'      Exit Function
'   End If
   
   If Not initDoc Then Exit Function

   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
   
   For J = 1 To UBound(Total1)
      Total1(J) = 0
   Next J
   Set m_TempCol = mcolParam("EX_WORK_PRICE_APPROVED")
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
   
      I = 0
     

   Set BodyCf = New CReportField
   DOCUMENT_TYPE = mcolParam("DOCUMENT_TYPE")
   If mcolParam("AREA") = 1 Then
      For Each EWPI In m_TempCol
         I = I + 1
          J = 0
         Call m_Details1.ClearField
         
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = I
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)

         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = EWPI.PART_MASTER_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)

         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = EWPI.PART_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)

         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = FormatNumber(EWPI.PACKAGE_RATE, -1)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)

         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = ConvertPerPack(EWPI.PART_TYPE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)

         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         '=====
         
         Set Temp_LTK = GetObject("CLoginTracking", m_LoginTracking, Trim(str(EWPI.CREATE_BY)), False)
         If Not Temp_LTK Is Nothing Then
             Call getUser("1", Temp_LTK.USER_NAME)
         End If
         Call getUser("2", EWPI.VERIFY_NAME)
         Call getUser("3", EWPI.APPROVED_NAME)
      Next EWPI
   ElseIf mcolParam("AREA") = 2 Then
     For Each EDCI In m_TempCol
            I = I + 1
            J = 0
            Call m_Details1.ClearField
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = I
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = EDCI.CUSTOMER_CODE
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = EDCI.CUSTOMER_NAME
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = EDCI.DELIVERY_CUS_ITEM_CODE
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = EDCI.DELIVERY_CUS_ITEM_NAME
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = FormatNumber(EDCI.RATE_DELIVERY, 3)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = DeliveryUnit(EDCI.RATE_TYPE)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = FormatNumber(EDCI.WEIGHT_PER_PACK, 0)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = DeliveryType(EDCI.RATE_TYPE)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = FormatNumber(EDCI.RATE_CUSTOMER, 3)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = DeliveryUnit(EDCI.RATE_TYPE_CUS)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
            
            J = J + 1
            Set HeadCf = m_Headers1.Fields(J)
            TempStr = FormatNumber(EDCI.WEIGHT_PER_PACK_CUS, 0)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
   
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
            '=====
            Set Temp_LTK = GetObject("CLoginTracking", m_LoginTracking, Trim(str(EDCI.CREATE_BY)), False)
            If Not Temp_LTK Is Nothing Then
                Call getUser("1", Temp_LTK.USER_NAME)
            End If
            Call getUser("2", EDCI.VERIFY_NAME)
            Call getUser("3", EDCI.APPROVED_NAME)
         Next EDCI
   ElseIf mcolParam("AREA") = 3 Or mcolParam("AREA") = 5 Then
      For Each EPPI In m_TempCol
         I = I + 1
          J = 0
         Call m_Details1.ClearField
         
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = I
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)

         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = EPPI.CUSTOMER_CODE
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)

         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = EPPI.CUSTOMER_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = EPPI.PART_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)

         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = EPPI.PART_DESC
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)

         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = FormatNumber(EPPI.DISCOUNT_AMOUNT, 0)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)

         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = ConvertPerPack(EPPI.PART_TYPE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)

         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         '=====
         Set Temp_LTK = GetObject("CLoginTracking", m_LoginTracking, Trim(str(EPPI.CREATE_BY)), False)
         If Not Temp_LTK Is Nothing Then
             Call getUser("1", Temp_LTK.USER_NAME)
         End If
         Call getUser("2", EPPI.VERIFY_NAME)
         Call getUser("3", EPPI.APPROVED_NAME)
      Next EPPI
   ElseIf mcolParam("AREA") = 4 Then
    For Each EPDI In m_TempCol
         I = I + 1
          J = 0
         Call m_Details1.ClearField
         
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = I
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)

         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = EPDI.CUSTOMER_CODE
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)

         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = EPDI.CUSTOMER_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = EPDI.DELIVERY_CUS_ITEM_CODE
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)

         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = EPDI.DELIVERY_CUS_ITEM_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)

         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = FormatNumber(EPDI.DISCOUNT_AMOUNT, 3)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)

         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = DeliveryUnit(EPDI.RATE_TYPE_CUS)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         J = J + 1
         Set HeadCf = m_Headers1.Fields(J)
         TempStr = FormatNumber(EPDI.WEIGHT_PER_PACK_CUS, 0)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)

         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         '=====
         Set Temp_LTK = GetObject("CLoginTracking", m_LoginTracking, Trim(str(EPDI.CREATE_BY)), False)
         If Not Temp_LTK Is Nothing Then
             Call getUser("1", Temp_LTK.USER_NAME)
         End If
         Call getUser("2", EPDI.VERIFY_NAME)
         Call getUser("3", EPDI.APPROVED_NAME)
      Next EPDI
   End If
      
   Vsp = ""
      
   If Len(mcolParam("DOCUMENT_NOTE")) > 0 Then
      Vsp.FontBold = True
       Vsp = "�����˵� : "
       Vsp.FontBold = False
      Vsp = EmptyToString(mcolParam("DOCUMENT_NOTE"), "")
   End If
         
   Call PrintFooter(alngX, userVerify, userApproved, CREATE_BY)

      Set BodyCf = Nothing

   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   Set EWPI = Nothing
   
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
      Case 13 '�ʹ�ѹ
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
      Case 14 '�ʹ�ѹ
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
   Call Vsp.AddTable(strFormat, "", "|˹�ҷ��|: " & Vsp.CurrentPage)
   Call Vsp.AddTable(strFormat, "", "|�ѹ�������|: " & DateToStringExt(mdteDate))
   Call Vsp.AddTable(strFormat, "", "|����|: " & Right$(DateToStringExtEx(mdteDate), 8))
'   Call Vsp.AddTable(strFormat, "", "|�����|: " & glbUser.USER_NAME)
   Vsp.CurrentY = dY(0)
   
   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   Vsp.TextAlign = taCenterBottom
   Vsp.Paragraph = MODULE_DESC
   Vsp.Paragraph = DocDate
   Vsp = ""
   ' vbCrLf & DocDate & vbCrLf & FromDate & vbCrLf & ToDate
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
Private Sub PrintFooter(alngX() As Long, VerifyName As String, ApprovedName As String, CREATE_BY As String)
Dim strFormat As String
Dim strPrint As String
Dim tempUserAccount As CUserAccount
Dim createRealName As String
Dim approveRealName As String
Dim modifyRealName As String
Dim strUser() As String
Dim I As Long
Dim tempUsers As CUser

   Vsp.TableBorder = tbNone
   Vsp.FontBold = False
   
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   For Each tempUsers In collUser
     If tempUsers.GROUP_ID = 1 Then
         Set tempUserAccount = GetObject("CUserAccount", collUserAccount, Trim(tempUsers.USER_NAME))
         If Len(modifyRealName) = 0 Then
          modifyRealName = "(  " & tempUserAccount.REAL_NAME & "  )"
        Else
         modifyRealName = modifyRealName & vbCrLf & "(  " & tempUserAccount.REAL_NAME & "  )"
        End If
      End If
   Next tempUsers
   
   For Each tempUsers In collUser
     If tempUsers.GROUP_ID = 2 Then
         Set tempUserAccount = GetObject("CUserAccount", collUserAccount, Trim(tempUsers.USER_NAME))
         If Len(createRealName) = 0 Then
          createRealName = "(  " & tempUserAccount.REAL_NAME & "  )"
        Else
         createRealName = createRealName & vbCrLf & "(  " & tempUserAccount.REAL_NAME & "  )"
        End If
      End If
   Next tempUsers
   
   For Each tempUsers In collUser
     If tempUsers.GROUP_ID = 3 Then
         Set tempUserAccount = GetObject("CUserAccount", collUserAccount, Trim(tempUsers.USER_NAME))
         If Len(approveRealName) = 0 Then
          approveRealName = "(  " & tempUserAccount.REAL_NAME & "  )"
        Else
         approveRealName = approveRealName & vbCrLf & "(  " & tempUserAccount.REAL_NAME & "  )"
        End If
      End If
   Next tempUsers
   
   Vsp.FontSize = 15
   Vsp.FontBold = False
   strFormat = VSP_CalTable("^15|^15|^15;", mdblWid, alngX)
   strPrint = modifyRealName & "|" & createRealName & "|" & approveRealName & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Vsp.FontSize = 12
   Vsp.FontBold = True
   strFormat = VSP_CalTable("^15|^15|^15;", mdblWid, alngX)
   strPrint = "���Ѵ��|����Ǩ�ͺ|���͹��ѵ�;"
   Call Vsp.AddTable(strFormat, "", strPrint)
End Sub





