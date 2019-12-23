VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmWeightPreview 
   ClientHeight    =   9420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18315
   Icon            =   "frmWeightPreview.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   18315
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   11040
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   18375
      _ExtentX        =   32411
      _ExtentY        =   19473
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboOrderBy 
         Height          =   315
         Left            =   7200
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1440
         Width           =   2985
      End
      Begin VB.ComboBox cboOrderType 
         Height          =   315
         Left            =   7200
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1920
         Width           =   2985
      End
      Begin VB.ComboBox cboSupType 
         Height          =   315
         Left            =   7200
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   960
         Width           =   2955
      End
      Begin prjFarmManagement.uctlTextBox txtSupplierCode 
         Height          =   495
         Left            =   1680
         TabIndex        =   13
         Top             =   1440
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   873
      End
      Begin prjFarmManagement.uctlDate uctlWeightDate 
         Height          =   405
         Left            =   1680
         TabIndex        =   10
         Top             =   960
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   120
         Top             =   720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   6015
         Left            =   120
         TabIndex        =   1
         Top             =   2520
         Width           =   18135
         _ExtentX        =   31988
         _ExtentY        =   10610
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         TabKeyBehavior  =   1
         MethodHoldFields=   -1  'True
         AllowColumnDrag =   0   'False
         AllowEdit       =   0   'False
         BorderStyle     =   3
         GroupByBoxVisible=   0   'False
         DataMode        =   99
         HeaderFontName  =   "AngsanaUPC"
         HeaderFontBold  =   -1  'True
         HeaderFontSize  =   14.25
         HeaderFontWeight=   700
         FontSize        =   9.75
         BackColorBkg    =   16777215
         ColumnHeaderHeight=   480
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmWeightPreview.frx":27A2
         Column(2)       =   "frmWeightPreview.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmWeightPreview.frx":290E
         FormatStyle(2)  =   "frmWeightPreview.frx":2A6A
         FormatStyle(3)  =   "frmWeightPreview.frx":2B1A
         FormatStyle(4)  =   "frmWeightPreview.frx":2BCE
         FormatStyle(5)  =   "frmWeightPreview.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmWeightPreview.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   15
         TabIndex        =   6
         Top             =   0
         Width           =   18285
         _ExtentX        =   32253
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtTruckNo 
         Height          =   495
         Left            =   1680
         TabIndex        =   14
         Top             =   1920
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   873
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   5400
         TabIndex        =   22
         Top             =   1560
         Width           =   1755
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   5400
         TabIndex        =   21
         Top             =   1980
         Width           =   1755
      End
      Begin VB.Label lblSupType 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5880
         TabIndex        =   18
         Top             =   960
         Width           =   1155
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   525
         Left            =   14760
         TabIndex        =   16
         Top             =   8760
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   11040
         TabIndex        =   15
         Top             =   1560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmWeightPreview.frx":2F36
         ButtonStyle     =   3
      End
      Begin VB.Label lblTruckNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   360
         TabIndex        =   12
         Top             =   2040
         Width           =   1155
      End
      Begin VB.Label lblSupplierCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   360
         TabIndex        =   11
         Top             =   1560
         Width           =   1155
      End
      Begin Threed.SSCommand cmdSelectAll 
         Height          =   525
         Left            =   2640
         TabIndex        =   3
         Top             =   4440
         Visible         =   0   'False
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmWeightPreview.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSelect 
         Height          =   525
         Left            =   1680
         TabIndex        =   2
         Top             =   4440
         Visible         =   0   'False
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmWeightPreview.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   11040
         TabIndex        =   0
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmWeightPreview.frx":3884
         ButtonStyle     =   3
      End
      Begin VB.Label lblWeightDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   360
         TabIndex        =   9
         Top             =   960
         Width           =   1155
      End
      Begin VB.Label Label4 
         Height          =   315
         Left            =   10920
         TabIndex        =   8
         Top             =   3420
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   570
         TabIndex        =   7
         Top             =   1320
         Visible         =   0   'False
         Width           =   1155
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   16560
         TabIndex        =   4
         Top             =   8760
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmWeightPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Rs2 As ADODB.Recordset
Private m_Weight As CWeight

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public CancelWeigth As Boolean
Public TempWeigth As Boolean
Public EditTempWeigth As Boolean
Public id As Long
Public TempCollection As Collection
Public SumLotAmount As Double

Private FileName As String
Private m_SumUnit As Double
Private m_TempCol1 As Collection
Public m_TempCol2 As Collection
Public m_Suppliers As Collection
Public m_SuppliersType As Collection

Public SupplierID As Long
Public SupplierCode As String
Public DocumentType As Long
Public DocumentDate As Date
Public TruckNo As String

Private Sub PopulateDestColl()
Dim Ri As CSupItem
Dim D As CSupItem

   If TempCollection Is Nothing Then
      Exit Sub
   End If
   
   For Each Ri In TempCollection
      Set D = New CSupItem
      
      If Ri.Flag <> "D" Then
         Call D.CopyObject(1, Ri)
         Call m_TempCol2.add(D)
      End If
      
      Set D = Nothing
   Next Ri
End Sub

Private Function IsIn(TempCol As Collection, TempID As Long) As Boolean
Dim D As CSupItem
Dim Found As Boolean
   
   Found = False
   For Each D In TempCol
      If D.PO_ID = TempID Then
         Found = True
      End If
   Next D
   
   IsIn = Found
End Function

Private Sub GenerateSourceItem(Rs As ADODB.Recordset, Rs2 As ADODB.Recordset, TempCol As Collection)
Dim CW As CWeight
Dim TempD As CSupplier
   Set m_TempCol1 = Nothing
   Set m_TempCol1 = New Collection
   While Not Rs2.EOF
      Set CW = New CWeight
      Call CW.PopulateFromRS(2, Rs2)
      Call TempCol.add(CW)
      Set CW = Nothing
      Rs2.MoveNext
   Wend
   
   While Not Rs.EOF
      Set CW = New CWeight
      Call CW.PopulateFromRS(1, Rs)
      Set TempD = GetObject("CSupplier", m_Suppliers, Trim(CW.CUST_ID), False)
      If Not TempD Is Nothing Then
            Call TempCol.add(CW)
      End If
      Set CW = Nothing
      Rs.MoveNext
   Wend
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
'
'   If Not VerifyDate(lblWeightDate, uctlWeightDate, False) Then
'      Exit Sub
'   End If
   
   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
      m_Weight.Date1 = uctlWeightDate.ShowDate
      m_Weight.CUST_ID = Trim(txtSupplierCode.Text) 'SupplierCode
      m_Weight.TRUCK_ID = Trim(txtTruckNo.Text) 'TruckNo
      m_Weight.OrderBy = cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex))
      m_Weight.OrderType = cboOrderType.ItemData(Minus2Zero(cboOrderType.ListIndex))
   
      Call glbDaily.QueryLegacyWeight(m_Weight, m_Rs, m_Rs2, ItemCount, IsOK, glbErrorLog)
      
   End If

   If ItemCount > 0 Then
      Call GenerateSourceItem(m_Rs, m_Rs2, m_TempCol1)
      GridEX1.ItemCount = m_TempCol1.Count
      GridEX1.Rebind
   Else
      GridEX1.ItemCount = 0
      GridEX1.Rebind
   End If

   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cboGroup_Click()
   m_HasModify = True
End Sub

Private Sub chkEnable_Click()
   m_HasModify = True
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
      
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Call PopulateTempColl(SumLotAmount)

   Call EnableForm(Me, True)
   SaveData = True
End Function


Private Sub chkSaleFlag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub


Private Sub cmdClear_Click()
   uctlWeightDate.ShowDate = -1
   txtSupplierCode.Text = ""
   txtTruckNo.Text = ""
   cboSupType.ListIndex = 0
End Sub





Private Sub cmdPictureAdd_Click()
On Error Resume Next
Dim strDescription As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Picture Files (*.jpg, *.gif)|*.jpg;*.gif"
   dlgAdd.DialogTitle = "Select Picture to Add to Database"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
   m_HasModify = True
End Sub

Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub

Private Sub cmdPrint_Click()
Dim Report As CReportInterface
Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
Dim ReportKey As String
Dim ReportFlag As Boolean
Dim Rc As CReportConfig
Dim iCount As Long
Dim EditMode As SHOW_MODE_TYPE
Dim ReportMode As Long

   ReportMode = 1
   
   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("ใบรายงานชั่งประจำวัน", "ปรับค่าหน้ากระดาษ")
   If lMenuChosen = 0 Then
      Exit Sub
   End If

   Call EnableForm(Me, False)

   If lMenuChosen = 1 Then
      ReportKey = "CReportWeight"
      Set Report = New CReportWeight
      ReportFlag = True
      Call Report.AddParam(1, "PREVIEW_TYPE")
      Call Report.AddParam("ใบรายงานชั่งประจำวัน", "REPORT_NAME")
   End If

   If Not Report Is Nothing Then
      Call Report.AddParam(m_TempCol1, "WEIGHT_REPORT")
      Call Report.AddParam(uctlWeightDate.ShowDate, "FROM_DATE")
      Call Report.AddParam(lMenuChosen, "REPORT_TYPE")
      Call Report.AddParam(ReportKey, "REPORT_KEY")
   End If

   If ReportFlag Then
    frmReport.ClassName = ReportKey
      Set frmReport.ReportObject = Report

      frmReport.HeaderText = pnlHeader.Caption
      Load frmReport
      frmReport.Show 1

      Unload frmReport
      Set frmReport = Nothing
      Set Report = Nothing

   Else
   
   ReportKey = "CReportWeight"
   ReportMode = 1
   
   Set Rc = New CReportConfig
   Rc.REPORT_KEY = ReportKey
   Call Rc.QueryData(m_Rs, iCount)
   
   If Not m_Rs.EOF Then
      Call Rc.PopulateFromRS(1, m_Rs)
      frmReportConfig.ShowMode = SHOW_EDIT
      frmReportConfig.id = Rc.REPORT_CONFIG_ID
   Else
      frmReportConfig.ShowMode = SHOW_ADD
   End If
   
   frmReportConfig.ReportMode = ReportMode
   frmReportConfig.ReportKey = ReportKey
   frmReportConfig.HeaderText = HeaderText
   Load frmReportConfig
   frmReportConfig.Show 1
   
   Unload frmReportConfig
   Set frmReportConfig = Nothing
   
   Set Rc = Nothing
   End If

   Call EnableForm(Me, True)
End Sub

Private Sub cmdSearch_Click()
   Call LoadSupplier(Nothing, m_Suppliers, 2, , cboSupType.ItemData(Minus2Zero(cboSupType.ListIndex)))
   Call QueryData(True)
End Sub

Public Sub CopyItem(TempCol1 As Collection, TempCol2 As Collection, id As Long)
Dim L As CWeight

   If id > 0 Then
      Set L = TempCol1(id)
      
      L.Flag = "A"
      Call TempCol2.add(L)
      TempCol1.Remove (id)
   End If
End Sub
Public Sub CopyAllItem(TempCol1 As Collection, TempCol2 As Collection)
Dim J As Long

   For J = 1 To TempCol1.Count
      TempCol1(J).Flag = "A"
      TempCol1(J).IncludeFlag = True
      Call TempCol2.add(TempCol1(J))
   Next J
   Set TempCol1 = Nothing
   Set TempCol1 = New Collection
End Sub


Public Sub PopulateTempColl(Tempsum As Double)
Dim D As CWeight
Dim Ri As CWeight
Dim Sum As Double
   Sum = 0
   For Each D In m_TempCol2
      Set Ri = New CWeight '
            Ri.WEIGHT_ID = D.WEIGHT_ID
            Ri.WEIGHT1 = D.WEIGHT1
            Ri.WEIGHT2 = D.WEIGHT2
            Ri.NetWeight = D.NetWeight
            Ri.TRUCK_ID = D.TRUCK_ID
            Ri.Date1 = D.Date1
            Ri.Date2 = D.Date2
            Ri.DateShow1 = D.DateShow1
            Ri.DateShow2 = D.DateShow2
            Ri.Time1 = D.Time1
            Ri.Time2 = D.Time2
            Ri.REMARK = D.REMARK
            Ri.Flag = "A"
            Call TempCollection.add(Ri, "1")
      Set Ri = Nothing
   Next D
   
   Tempsum = Sum
End Sub





Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      uctlWeightDate.ShowDate = Now
      
        Set glbDatabaseMngr2 = New clsDatabaseMngr
       If Not glbDatabaseMngr2.ConnectLegacyDatabase(glbParameterObj.DBFileACCESSS, glbParameterObj.UserNameAccess, glbParameterObj.PasswordAccess, glbErrorLog) Then
       End If
       
       Call LoadSupplierType(cboSupType, m_SuppliersType)
      
      Call InitWeightOrderBy(cboOrderBy)
      Call InitOrderType(cboOrderType)
       
      uctlWeightDate.ShowDate = Now
      txtSupplierCode.Text = SupplierCode
      txtTruckNo.Text = TruckNo
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_Weight.QueryFlag = 1
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         m_Weight.QueryFlag = 0
         Call QueryData(True)
      End If
      
      Call EnableForm(Me, True)
      m_HasModify = False
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.NAME
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 116 Then
'      Call cmdSearch_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 118 Then
'      Call cmdAdd_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
'      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
'      Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 123 Then
      Call AddMemoNote
      KeyCode = 0
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   If m_Rs2.State = adStateOpen Then
      m_Rs2.Close
   End If
   Set m_Rs2 = Nothing
   
   Set m_Weight = Nothing
   Set m_TempCol1 = Nothing
   Set m_TempCol2 = Nothing
   Set glbDatabaseMngr2 = Nothing
   Set m_Suppliers = Nothing
   Set m_SuppliersType = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub GridEX2_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid1()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.NAME = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 1600
   Col.Caption = "เลขที่ใบชั่ง"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 1200
   Col.Caption = "รหัสซัพฯ"
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 3000
   Col.Caption = "ชื่อซัพฯ"
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 1200
   Col.Caption = "รหัสสินค้า"
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 3500
   Col.Caption = "ชื่อสินค้า"

'   Set Col = GridEX1.Columns.add '3
'   Col.Width = 1500
'   Col.Caption = MapText("วันที่เข้า")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 1200
   Col.Caption = MapText("น้ำหนักเข้า")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 1200
   Col.Caption = MapText("น้ำหนักออก")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.add '6
   Col.Width = 1200
   Col.Caption = MapText("น้ำหนักสุทธิ")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = 2300
   Col.Caption = MapText("เวลาเข้า")
   Col.TextAlignment = jgexAlignCenter
   
   Set Col = GridEX1.Columns.add '8
   Col.Width = 2300
   Col.Caption = MapText("เวลาออก")
   Col.TextAlignment = jgexAlignCenter
   
   Set Col = GridEX1.Columns.add '9
   Col.Width = 2500
   Col.Caption = MapText("ทะเบียนรถ")
   
   Set Col = GridEX1.Columns.add '10
   Col.Width = 3500
   Col.Caption = MapText("หมายเหตุ")
End Sub

Private Sub GetTotalPrice()
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdPrint.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdClear.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitCombo(cboSupType)
   Call InitCombo(cboOrderBy)
   Call InitCombo(cboOrderType)
   
   Call txtSupplierCode.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtTruckNo.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   
   Call InitMainButton(cmdPrint, MapText("พิมพ์"))
   Call InitMainButton(cmdOK, MapText("ตกลง (ESC)"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา"))
   Call InitMainButton(cmdClear, MapText("เคลียร์ (F4)"))
   
   Call InitNormalLabel(lblWeightDate, MapText("วันที่ชั่ง"))
   Call InitNormalLabel(lblSupplierCode, MapText("รหัสซัพฯ"))
   Call InitNormalLabel(lblTruckNo, MapText("ทะเบียนรถ"))
   Call InitNormalLabel(lblSupType, MapText("ประเภทซัพฯ"))
   Call InitNormalLabel(lblOrderBy, MapText("เรียงตาม"))
   Call InitNormalLabel(lblOrderType, MapText("เรียงจาก"))
   
   
   Call InitGrid1
End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   OKClick = False
   Call InitFormLayout
      
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_Rs2 = New ADODB.Recordset
   Set m_Weight = New CWeight
   Set m_TempCol1 = New Collection
   Set m_TempCol2 = New Collection
   Set TempCollection = New Collection
   Set m_Suppliers = New Collection
   Set m_SuppliersType = New Collection
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub



Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"


   If m_TempCol1 Is Nothing Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If

   Dim CW As CWeight
   If m_TempCol1.Count <= 0 Then
      Exit Sub
   End If
   Set CW = GetItem(m_TempCol1, RowIndex, RealIndex)
   If CW Is Nothing Then
      Exit Sub
   End If

   Values(1) = CW.WEIGHT_ID
   Values(2) = RealIndex
   Values(3) = CW.CUST_ID
   Values(4) = CW.CUST_NAME
   Values(5) = CW.PRODUCT_ID
   Values(6) = CW.PRODUCT_NAME
'   Values(7) = CW.DateShow1
   Values(7) = FormatNumber(CW.WEIGHT1, 0)
   Values(8) = FormatNumber(CW.WEIGHT2, 0)
   Values(9) = FormatNumber(CW.NetWeight, 0)
   Values(10) = CW.DateShow1 & " " & CW.Time1
   Values(11) = IIf(CW.DateShow2 = "", "", CW.DateShow2 & " " & CW.Time2)
   Values(12) = CW.TRUCK_ID
   Values(13) = CW.REMARK
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub txtTruckNo_Change()
   m_HasModify = True
End Sub

