VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportProduct008"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportProduct004"
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
Private m_Rs1 As ADODB.Recordset
Private m_Rs2 As ADODB.Recordset

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
   Call Cf.SetFieldValue(2.7, "^", "�����Ţ㺻����Թ", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "��������´", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2.7, "^", "�����Ţặ", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "�ѹ�������Թ", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "�ѹ���������ҹ", "^")
   Call m_Headers1.AddField(Cf)
    Call Cf.SetFieldValue(3, "^", "�ѹ������稧ҹ", "^")
   Call m_Headers1.AddField(Cf)
  Call Cf.SetFieldValue(3, "^", "͹��ѵ���", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "�Ѻ�Դ�ͺ��", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(5, "^", "����", "^")
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
Private Sub printHeader1()
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
   Call Cf.SetFieldValue(2.7, "^", "�����Ţ�ѵ�شԺ", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "�����ѵ�شԺ", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2.7, "^", "�������ѵ�شԺ", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�ӹǹ", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "˹���", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�Ҥ�", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�Ҥ����", ">")
   Call m_Headers1.AddField(Cf)
    Call Cf.SetFieldValue(3, "^", "ʶҹ����ԡ", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "�����Թ��Ң��", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "�����Ţ��ҧ�ԧ", "^")
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
Private Sub printHeader2()
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
   Call Cf.SetFieldValue(2.7, "^", "�����Ţ��Ե�ѳ��", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "���ͼ�Ե�ѳ��", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2.7, "^", "��������Ե�ѳ��", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�ӹǹ", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "˹���", "<")
   Call m_Headers1.AddField(Cf)
  Call Cf.SetFieldValue(2, "^", "�Ҥ�", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�Ҥ����", ">")
   Call m_Headers1.AddField(Cf)
    Call Cf.SetFieldValue(3, "^", "ʶҹ�����", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "�����Թ��Ң��", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "�����Ţ��ҧ�ԧ", "^")
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
Private Sub printHeader3()
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
   Call Cf.SetFieldValue(4, "^", "���;�ѡ�ҹ", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "���˹�", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "�ӹǹ���ҷ����(�������/�ҷ�)", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�ѹ����", "<")
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

Private Sub printHeader4()
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
   Call Cf.SetFieldValue(2.7, "^", "�����Ţ����ͧ�ѡ�", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "��������ͧ�ѡ�", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "�ӹǹ���ҷ����(�������/�ҷ�)", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�ѹ�����", "<")
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
Private Sub printHeader5()
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
   Call Cf.SetFieldValue(3, "^", "�����Ե����", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(7, "^", "��������´", "<")
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
      
   MODULE_DESC = mcolParam("REPORT_NAME") & vbCrLf

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
Dim II As Long
Dim J As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Amt As Double
Dim Cm As CJob
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set Rs = New ADODB.Recordset
   Set m_Rs1 = New ADODB.Recordset
   Set m_Rs2 = New ADODB.Recordset
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------

   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
      II = 0
      Set BodyCf = New CReportField

   Set Cm = New CJob
            Cm.JOB_ID = -1
      Cm.JOB_NO = mcolParam("JOB_NO")
      Cm.JOB_DESC = mcolParam("JOB_DESC")
      Cm.BATCH_NO = mcolParam("BATCH_NO")
      Cm.JOB_DATE = mcolParam("JOB_DATE")
      Cm.START_DATE = mcolParam("START_JOB")
      Cm.FINISH_DATE = mcolParam("FINISH_JOB")
      Cm.APPROVED_BY = mcolParam("APPROVED_BY")
      Cm.RESPONSE_BY = mcolParam("RESPONSE_BY")
      Cm.PROCESS_ID = mcolParam("PROCESS_ID")
      If mcolParam("JOB_STATUS") < 1 Then
      Cm.COMMIT_FLAG = ""
      ElseIf mcolParam("JOB_STATUS") = 1 Then
      Cm.COMMIT_FLAG = "Y"
       ElseIf mcolParam("JOB_STATUS") = 2 Then
       Cm.COMMIT_FLAG = "N"
      End If
      Cm.OrderBy = mcolParam("ORDER_BY")
      Cm.OrderType = mcolParam("ORDER_TYPE")
      Cm.QueryFlag = -1
      Cm.JOB_DOC_TYPE = 2
      Call glbProduction.QueryJob(Cm, Rs, iCount, IsOK, glbErrorLog)
    If (Rs Is Nothing) Or (iCount < 1) Then
      mstrErrMsg = "��辺�����ŷ���ͧ���㹰ҹ������"
      Exit Function
   End If

   II = 0
   Dim counti As Long
     counti = 0
     
      While Not Rs.EOF
         II = II + 1
         Call m_Details1.ClearField
          Call Cm.PopulateFromRS(1, Rs)
         

     If counti > 0 Then
      Vsp.NewPage
      End If
            Call printHeader
    'INPUT ++++++++++++++++++++++++++
      Dim Inp As CJobInput
      Set Inp = New CJobInput
      Inp.JOB_INOUT_ID = -1
      Inp.JOB_ID = Cm.JOB_ID
      Call Inp.QueryData(1, m_Rs2, iCount)
      Set Inp = Nothing
      
      Set Cm.Inputs = Nothing
      Set Cm.Inputs = New Collection
      While Not m_Rs2.EOF
       Set Inp = New CJobInput
         Call Inp.PopulateFromRS(1, m_Rs2)
                 Inp.Flag = "I"
         If Inp.TX_TYPE = "E" Then
            Call Cm.Inputs.add(Inp)
            End If
         Set Inp = Nothing
                 m_Rs2.MoveNext
      Wend
'INPUT ++++++++++++++++++++++++++

'OUTPUT ++++++++++++++++++++++++++
      Dim Op As CJobInput
      Set Op = New CJobInput
      Op.JOB_INOUT_ID = -1
      Op.JOB_ID = Cm.JOB_ID
      Call Op.QueryData(1, m_Rs2, iCount)
      Set Op = Nothing
      
      Set Cm.Outputs = Nothing
      Set Cm.Outputs = New Collection
      While Not m_Rs2.EOF
       Set Op = New CJobInput
         Call Op.PopulateFromRS(1, m_Rs2)
                  Op.Flag = "I"
         If Op.TX_TYPE = "I" Then
            Call Cm.Outputs.add(Op)
            End If
         Set Op = Nothing
                 m_Rs2.MoveNext
      Wend
'OUTPUT ++++++++++++++++++++++++++

'MACHINE USED TIME ++++++++++++++++++++++++++
      Dim EH As CJobResource
     Set EH = New CJobResource
      EH.JOB_ID = Cm.JOB_ID
      Call EH.QueryData(m_Rs1, iCount)
      Set EH = Nothing
      
      Set Cm.Machines = Nothing
      Set Cm.Machines = New Collection
      While Not m_Rs1.EOF
         Set EH = New CJobResource
         Call EH.PopulateFromRS(1, m_Rs1)
      
         EH.Flag = "I"
         If EH.MACHINE_NO <> "" Then
         Call Cm.Machines.add(EH)
         End If
         Set EH = Nothing
         m_Rs1.MoveNext
      Wend
      'MACHINE TIME USED ++++++++++++++++++++++++++
        
'PERSON USED TIME ++++++++++++++++++++++++++
      Dim Ep As CJobResource
     Set Ep = New CJobResource
      Ep.JOB_ID = Cm.JOB_ID
      Call Ep.QueryData(m_Rs1, iCount)
      Set Ep = Nothing
      
      Set Cm.Peoples = Nothing
      Set Cm.Peoples = New Collection
      While Not m_Rs1.EOF
         Set Ep = New CJobResource
         Call Ep.PopulateFromRS(2, m_Rs1)
      
         Ep.Flag = "I"
         If Ep.EMP_ID > 0 Then
         Call Cm.Peoples.add(Ep)
         End If
         Set Ep = Nothing
         m_Rs1.MoveNext
      Wend
      'PERSON TIME USED ++++++++++++++++++++++++++
     'PARAMETER TIME ++++++++++++++++++++++++++
      Dim PP As CJobParameter
     Set PP = New CJobParameter
      PP.JOB_ID = Cm.JOB_ID
      Call PP.QueryData(1, m_Rs1, iCount)
      Set PP = Nothing
      
      Set Cm.Parameters = Nothing
      Set Cm.Parameters = New Collection
      While Not m_Rs1.EOF
         Set PP = New CJobParameter
         Call PP.PopulateFromRS(1, m_Rs1)
      
         PP.Flag = "I"
         Call Cm.Parameters.add(PP)
         Set PP = Nothing
         m_Rs1.MoveNext
      Wend
      'PARAMETER USED ++++++++++++++++++++++++++

  Set HeadCf = m_Headers1.Fields(1)
         TempStr = II
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(2)
         TempStr = Cm.JOB_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(3)
         TempStr = Cm.JOB_DESC
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(4)
         TempStr = Cm.BATCH_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(5)
         TempStr = DateToStringExt(Cm.JOB_DATE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(6)
         TempStr = DateToStringExt(Cm.START_DATE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(7)
         TempStr = DateToStringExt(Cm.FINISH_DATE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         
         Set HeadCf = m_Headers1.Fields(8)
         TempStr = Replace(Cm.LONG_NAMEA, vbCrLf, "") & " " & Cm.LAST_NAMEA
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(9)
         TempStr = Replace(Cm.LONG_NAMER, vbCrLf, "") & " " & Cm.LAST_NAMER
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(10)
         TempStr = Cm.PROCESS_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         
         
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
    
    If Cm.Inputs.Count > 0 Then
         Vsp = ""
         Vsp = "�ѵ�شԺ�����"
        I = 0
         Call printHeader1
         Dim SA As CJobInput
         For Each SA In Cm.Inputs
         
         I = I + 1
         Call m_Details1.ClearField
                  
         Set HeadCf = m_Headers1.Fields(1)
         TempStr = I
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(2)
         TempStr = SA.PART_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
               
         Set HeadCf = m_Headers1.Fields(3)
         TempStr = SA.PART_DESC
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
              
        Set HeadCf = m_Headers1.Fields(4)
         TempStr = SA.PART_TYPE_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
                            
        Set HeadCf = m_Headers1.Fields(5)
         TempStr = SA.TX_AMOUNT
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
              
        
        Set HeadCf = m_Headers1.Fields(6)
         TempStr = FormatNumber(SA.AVG_PRICE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
        
         Set HeadCf = m_Headers1.Fields(7)
         TempStr = SA.UNIT_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

        Set HeadCf = m_Headers1.Fields(8)
         TempStr = FormatNumber(SA.AVG_PRICE * SA.TX_AMOUNT)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
       
        Set HeadCf = m_Headers1.Fields(9)
         TempStr = SA.LOCATION_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      
      
      
        Set HeadCf = m_Headers1.Fields(10)
         TempStr = SA.SERIAL_NUMBER
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      
        Set HeadCf = m_Headers1.Fields(11)
         TempStr = SA.INOUT_REF
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
                 
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         
         Next SA
    End If
    If Cm.Outputs.Count > 0 Then
     
         Vsp = ""
         Vsp = "��Ե�ѳ������"
        I = 0
         Call printHeader2
         Dim SB As CJobInput
         For Each SB In Cm.Outputs
         
         I = I + 1
         Call m_Details1.ClearField
                  
         Set HeadCf = m_Headers1.Fields(1)
         TempStr = I
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(2)
         TempStr = SB.PART_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
               
         Set HeadCf = m_Headers1.Fields(3)
         TempStr = SB.PART_DESC
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
              
        Set HeadCf = m_Headers1.Fields(4)
         TempStr = SB.PART_TYPE_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
                            
        Set HeadCf = m_Headers1.Fields(5)
         TempStr = SB.TX_AMOUNT
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
              
        
        Set HeadCf = m_Headers1.Fields(6)
         TempStr = FormatNumber(SB.AVG_PRICE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
       
       
         Set HeadCf = m_Headers1.Fields(7)
         TempStr = SB.UNIT_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
        Set HeadCf = m_Headers1.Fields(8)
         TempStr = FormatNumber(SB.AVG_PRICE * SB.TX_AMOUNT)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
       
        Set HeadCf = m_Headers1.Fields(9)
         TempStr = SB.LOCATION_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
        Set HeadCf = m_Headers1.Fields(10)
         TempStr = SB.SERIAL_NUMBER
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      
        Set HeadCf = m_Headers1.Fields(11)
         TempStr = SB.INOUT_REF
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
                 
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         
         Next SB
    End If
    If Cm.Peoples.Count > 0 Then
         Vsp = ""
         Vsp = "�ç�ҹ�����"
        I = 0
        Call printHeader3
         Dim SC As CJobResource
         For Each SC In Cm.Peoples
         I = I + 1
         Call m_Details1.ClearField
                  
         Set HeadCf = m_Headers1.Fields(1)
         TempStr = I
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(2)
         TempStr = SC.LONG_NAME & " " & SC.LAST_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
               
         Set HeadCf = m_Headers1.Fields(3)
         TempStr = SC.POSITION_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
              
        Set HeadCf = m_Headers1.Fields(4)
         TempStr = FormatNumber(SC.OCCUPY_INTERVAL)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
                            
        Set HeadCf = m_Headers1.Fields(5)
         TempStr = DateToStringExt(SC.OCCUPY_DATE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
              
                 
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         
         Next SC
    End If
    If Cm.Machines.Count > 0 Then
         Vsp = ""
         Vsp = "����ͧ�ѡ÷����"
        I = 0
        Call printHeader4
         Dim SD As CJobResource
         For Each SD In Cm.Machines
         I = I + 1
         Call m_Details1.ClearField
                  
         Set HeadCf = m_Headers1.Fields(1)
         TempStr = I
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(2)
         TempStr = SD.MACHINE_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
               
         Set HeadCf = m_Headers1.Fields(3)
         TempStr = SD.MACHINE_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
              
        Set HeadCf = m_Headers1.Fields(4)
         TempStr = FormatNumber(SD.OCCUPY_INTERVAL)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
                            
        Set HeadCf = m_Headers1.Fields(5)
         TempStr = DateToStringExt(SD.OCCUPY_DATE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
              
                 
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         
         Next SD
    End If
    If Cm.Parameters.Count > 0 Then
         Vsp = ""
         Vsp = "��������������"
        I = 0
        Call printHeader5
         Dim SDD As CJobParameter
         For Each SDD In Cm.Parameters
         I = I + 1
         Call m_Details1.ClearField
                  
         Set HeadCf = m_Headers1.Fields(1)
         TempStr = I
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(2)
         TempStr = SDD.PARAMETER_PROCESS_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
               
         Set HeadCf = m_Headers1.Fields(3)
         TempStr = SDD.JOB_PARAMETER_DESC
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
              
              
                 
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         
         Next SDD
    End If
         counti = counti + 1
         Rs.MoveNext
      Wend

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
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESDRIPTION
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
   Call Vsp.AddTable(strFormat, "", "|˹�ҷ��|: " & Vsp.CurrentPage)
   Call Vsp.AddTable(strFormat, "", "|�ѹ�������|: " & DateToStringExt(mdteDate))
   Call Vsp.AddTable(strFormat, "", "|����|: " & Right$(DateToStringExtEx(mdteDate), 8))
   Vsp.CurrentY = dY(0)
   
   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   Vsp.TextAlign = taCenterBottom
   Vsp.Paragraph = MODULE_DESC
   Vsp.Paragraph = ""
   If mblnHeader Then
      'Call printHeader
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


