VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportGraph001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportGraph001"
Private Const MODULE_DESC = ""


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
Private Const HEADER_FORMAT = "^2|^8|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1;"
Private Const DETAIL_FORMAT = "^2|^8|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1|^1;"
Private Const FOOTER_FORMAT = "^2|^8|>50;"

Private m_Points As Collection

Private Sub Class_Initialize()
   Set mcolParam = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
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

Private Function GenerateHeaderStr(TempStr2 As String) As String
Dim XStart As Long
Dim XWindows As Long
Dim TempStr As String
Dim I As Long
Dim p As CPoint

   TempStr = "NO.|กลุ่มตัวเลข|"
   TempStr2 = "||"

   XStart = mcolParam("X_START")
   XWindows = mcolParam("X_WINDOW")
   
   If XStart <= 0 Then
      XStart = 1
   End If
   
   For I = XStart To (XStart + XWindows - 1)
      If I <= m_Points.Count Then
         Set p = m_Points(I)
         TempStr = TempStr & p.Value
         TempStr2 = TempStr2 & I
      End If
      
      TempStr = TempStr & "|"
      TempStr2 = TempStr2 & "|"
   Next I
   
   TempStr = TempStr & ";"
   TempStr2 = TempStr2 & ";"
   
   GenerateHeaderStr = TempStr
End Function

Private Sub printHeader()
Dim strFormat As String
Dim strPrint As String
Dim tbdT As TableBorderSettings
Dim blnBold As Boolean
Dim iSize As Integer
Dim alngX() As Long
Dim TempStr As String

   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   strFormat = VSP_CalTable(HEADER_FORMAT, mdblWid, alngX)
   strPrint = GenerateHeaderStr(TempStr)
   Vsp.FontSize = 8
   Call Vsp.AddTable(strFormat, "", TempStr)
   Vsp.FontSize = iSize
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Vsp.TableBorder = tbdT
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
End Sub

Private Function initDoc() As Boolean
   Dim strDate As String
   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4
   Vsp.ORIENTATION = orLandscape
   Vsp.MarginBottom = 1440
   Vsp.MarginFooter = 1440
   Vsp.MarginHeader = 1440
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 1440
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))
   
   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
      mstrErrMsg = "Error GetDateTime Error."
      Exit Function
   End If
   mdteDate = InternalDateToDate(strDate)
   initDoc = True
   Set m_Points = mcolParam("POINTS")
End Function

Private Sub PlotGraph(Points As Collection)
Dim p As CPoint
Dim CurrentPoint As CPoint
Dim XStart As Double
Dim YStart As Double
Dim XOffset As Double
Dim YOffset As Double
Dim I As Long
Dim J As Long
Dim Start As Long
Dim Window As Long

   XOffset = 22.57 / 50 * 567
   YOffset = 0.58 * 567
   XStart = 5.55 * 567
   YStart = 5.4 * 567
   
   Set CurrentPoint = Nothing
   Vsp.PenWidth = 25
   I = 0
   
   Start = mcolParam("X_START")
   Window = mcolParam("X_WINDOW")
   If Start <= 0 Then
      Start = 0
   End If
   
   'Vsp.DrawLine (XStart + CurrentPoint.X * XOffset), (YStart + CurrentPoint.Y * YOffset), (XStart + P.X * XOffset), (YStart + P.Y * YOffset)
   J = 0
   For Each p In Points
      I = I + 1
      If (I >= Start) And (I <= Start + Window - 1) Then
         J = J + 1
         p.X = J
         If CurrentPoint Is Nothing Then
            Set CurrentPoint = New CPoint
            CurrentPoint.X = p.X
            CurrentPoint.Y = p.Y
         End If
         Vsp.DrawLine (XStart + CurrentPoint.X * XOffset), (YStart + CurrentPoint.Y * YOffset), (XStart + p.X * XOffset), (YStart + p.Y * YOffset)
         Set CurrentPoint = p
      End If
   Next p
   
   Set CurrentPoint = Nothing
End Sub

Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim RName As String
Dim iCount As Long
Dim I As Long
Dim strFormat As String
Dim strPrint As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Groups As Collection
Dim Cg As CYGroup

   RName = "genDoc"
   
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   If Not initDoc Then Exit Function

   Vsp.Preview = isPreview
''-----------------------------------------------------------------------------------------------------
''                                         Main Operation Here
''-----------------------------------------------------------------------------------------------------
''   ลำดับ รหัสลูกค้า ชื่อลูกค้า ประเภทลูกค้า ระดับลูกค้า ประเภทธุรกิจ เครดิต อีเมลล์ เว็บไซต์ รายละเอียด
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
      Set Groups = mcolParam("GROUPS")
      strFormat = VSP_CalTable(DETAIL_FORMAT, mdblWid, alngX)
      Vsp.TableBorder = tbAll
      
      I = 0
      For Each Cg In Groups
         I = I + 1
         
         strPrint = I & "|"
         strPrint = strPrint & Cg.Y_GROUP & "|"
         strPrint = strPrint & "||||||||||||||||||||||||||||||||||||||||||||||||;"
         Call Vsp.AddTable(strFormat, "", strPrint)
      Next Cg
      
      Call PlotGraph(m_Points)
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
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
   
   Vsp.FontSize = 8
   dY(0) = Vsp.CurrentY
   Vsp.TextAlign = taLeftBottom
   strFormat = VSP_CalTable(TITLE_FORMAT, mdblWid, alngX)
   Vsp.TableBorder = tbNone
   Call Vsp.AddTable(strFormat, "", "|หน้าที่|: " & Vsp.CurrentPage)
   Call Vsp.AddTable(strFormat, "", "|วันที่พิมพ์|: " & DateToStringExt(mdteDate))
   Call Vsp.AddTable(strFormat, "", "|เวลา|: " & Right$(DateToStringExtEx(mdteDate), 8))
   Vsp.CurrentY = dY(0)
   
   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   Vsp.TextAlign = taCenterBottom
   Vsp.Paragraph = "ชื่อสูตร : " & mcolParam("FORMULA_NAME") & "      รายละเอียด : " & mcolParam("FORMULA_DESC")
   Vsp.Paragraph = "ตัวเลข : " & mcolParam("X_COLLECTION_NAME") & "      กลุ่มตัวเลข : " & mcolParam("Y_COLLECTION_NAME")
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

