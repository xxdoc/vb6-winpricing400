VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmCommissionIncentive 
   BackColor       =   &H80000000&
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   11910
   Icon            =   "frmCommissionIncentive.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboOrderBy 
         Height          =   315
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1090
         Width           =   2955
      End
      Begin VB.ComboBox cboOrderType 
         Height          =   315
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1570
         Width           =   2955
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   5295
         Left            =   180
         TabIndex        =   2
         Top             =   2400
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   9340
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
         Column(1)       =   "frmCommissionIncentive.frx":27A2
         Column(2)       =   "frmCommissionIncentive.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmCommissionIncentive.frx":290E
         FormatStyle(2)  =   "frmCommissionIncentive.frx":2A6A
         FormatStyle(3)  =   "frmCommissionIncentive.frx":2B1A
         FormatStyle(4)  =   "frmCommissionIncentive.frx":2BCE
         FormatStyle(5)  =   "frmCommissionIncentive.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmCommissionIncentive.frx":2D5E
      End
      Begin prjFarmManagement.uctlTextBox txtFreelance 
         Height          =   435
         Left            =   6150
         TabIndex        =   12
         Top             =   1080
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtStockCode 
         Height          =   435
         Left            =   6150
         TabIndex        =   13
         Top             =   1560
         Width           =   2985
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   525
         Left            =   6800
         TabIndex        =   18
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmCommissionIncentive.frx":2F36
         ButtonStyle     =   3
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   360
         TabIndex        =   17
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   435
         TabIndex        =   16
         Top             =   1650
         Width           =   1365
      End
      Begin VB.Label lblStockCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4620
         TabIndex        =   11
         Top             =   1710
         Width           =   1455
      End
      Begin VB.Label lblFreelance 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4620
         TabIndex        =   10
         Top             =   1230
         Width           =   1455
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   10110
         TabIndex        =   0
         Top             =   930
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmCommissionIncentive.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   10110
         TabIndex        =   1
         Top             =   1500
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   5
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmCommissionIncentive.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   3
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmCommissionIncentive.frx":3884
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   4
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10095
         TabIndex        =   7
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8445
         TabIndex        =   6
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmCommissionIncentive.frx":3B9E
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmCommissionIncentive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_CommissionIncentive As CCommissionIncentive
Private m_TempCommissionIncentive As CCommissionIncentive
Private m_CollCI As Collection
Private m_Rs As ADODB.Recordset
Private m_TableName As String

Public OKClick As Boolean
Public HeaderText As String
Public DocumentType As Long
'DocumentType = 1 คือ ข้อมูล INCENTIVE รายสินค้า
'DocumentType = 2 คือ ข้อมูล INCENTIVE รายลูกค้า สินค้า
'DocumentType = 3 คือ เงื่อนไขข้อมูล COMMISSION
'DocumentType = 4 คือ เงื่อนไขข้อมูล INCENTIVE

Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean
Dim HeaderText As String
If DocumentType = 1 Then
   If Not VerifyAccessRight("COMMISSION_INCENTIVE_ADD") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   HeaderText = "เพิ่มข้อมูล INCENTIVE รายสินค้า"
ElseIf DocumentType = 2 Then
   If Not VerifyAccessRight("COMMISSION_INCENTIVE-CUS-PD_ADD") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   HeaderText = "เพิ่มข้อมูล INCENTIVE รายลูกค้า สินค้า"
 ElseIf DocumentType = 3 Then
   If Not VerifyAccessRight("COMMISSION_INCENTIVE-COM-EXTRA_ADD") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   HeaderText = "เพิ่มข้อมูล COMMISTION ของฟรีแลนซ์"
 ElseIf DocumentType = 4 Then
   If Not VerifyAccessRight("COMMISSION_INCENTIVE-INC-EXTRA_ADD") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   HeaderText = "เพิ่มข้อมูล INCENTIVE ของฟรีแลนซ์"
End If
   frmAddEditCommissionIncentive.HeaderText = MapText(HeaderText)
   Set frmAddEditCommissionIncentive.ParentForm = Me
   frmAddEditCommissionIncentive.ShowMode = SHOW_ADD
   frmAddEditCommissionIncentive.DocumentType = DocumentType
   Load frmAddEditCommissionIncentive
   frmAddEditCommissionIncentive.Show 1
   
   OKClick = frmAddEditCommissionIncentive.OKClick
   
   Unload frmAddEditCommissionIncentive
   Set frmAddEditCommissionIncentive = Nothing
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub

Private Sub cmdClear_Click()
   txtStockCode.Text = ""
   txtFreelance.Text = ""
   cboOrderBy.ListIndex = -1
   cboOrderType.ListIndex = -1
End Sub

Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim id As Long
If DocumentType = 1 Then
   If Not VerifyAccessRight("COMMISSION_INCENTIVE_DELETE") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
ElseIf DocumentType = 2 Then
   If Not VerifyAccessRight("COMMISSION_INCENTIVE-CUS-PD_DELETE") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
 ElseIf DocumentType = 3 Then
   If Not VerifyAccessRight("COMMISSION_INCENTIVE-COM-EXTRA_DELETE") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
 ElseIf DocumentType = 4 Then
   If Not VerifyAccessRight("COMMISSION_INCENTIVE-INC-EXTRA_DELETE") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
End If

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   id = GridEX1.Value(1)

   Call glbDatabaseMngr.LockTable(m_TableName, id, IsCanLock, glbErrorLog)
   If Not ConfirmDelete(GridEX1.Value(2)) Then
      Call glbDatabaseMngr.UnLockTable(m_TableName, id, IsCanLock, glbErrorLog)
      Exit Sub
   End If

   Call EnableForm(Me, False)
   If Not glbDaily.DeleteCommissionIncentive(id, IsOK, True, glbErrorLog) Then
      m_CommissionIncentive.INCENTIVE_ID = -1
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call glbDatabaseMngr.UnLockTable(m_TableName, id, IsCanLock, glbErrorLog)
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call QueryData(True)
   
   Call glbDatabaseMngr.UnLockTable(m_TableName, id, IsCanLock, glbErrorLog)
   Call EnableForm(Me, True)
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim id As Long
Dim OKClick As Boolean
Dim HeaderText As String
If DocumentType = 1 Then
   If Not VerifyAccessRight("COMMISSION_INCENTIVE_EDIT") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   HeaderText = "แก้ไขข้อมูล INCENTIVE รายสินค้า"
ElseIf DocumentType = 2 Then
   If Not VerifyAccessRight("COMMISSION_INCENTIVE-CUS-PD_EDIT") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   HeaderText = "แก้ไขข้อมูล INCENTIVE รายลูกค้า สินค้า"
 ElseIf DocumentType = 3 Then
   If Not VerifyAccessRight("COMMISSION_INCENTIVE-COM-EXTRA_EDIT") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   HeaderText = "แก้ไขข้อมูล COMMISTION ของฟรีแลนซ์"
 ElseIf DocumentType = 4 Then
   If Not VerifyAccessRight("COMMISSION_INCENTIVE-INC-EXTRA_EDIT") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   HeaderText = "แก้ไขข้อมูล INCENTIVE ของฟรีแลนซ์"
End If

If Not VerifyGrid(GridEX1.Value(1)) Then
   Exit Sub
End If

   id = Val(GridEX1.Value(1))
   Call glbDatabaseMngr.LockTable(m_TableName, id, IsCanLock, glbErrorLog)

   frmAddEditCommissionIncentive.id = id
   Set frmAddEditCommissionIncentive.ParentForm = Me
   frmAddEditCommissionIncentive.HeaderText = MapText(HeaderText)
   frmAddEditCommissionIncentive.ShowMode = SHOW_EDIT
   frmAddEditCommissionIncentive.DocumentType = DocumentType
   Load frmAddEditCommissionIncentive
   frmAddEditCommissionIncentive.Show 1

   OKClick = frmAddEditCommissionIncentive.OKClick

   Unload frmAddEditCommissionIncentive
   Set frmAddEditCommissionIncentive = Nothing

   If OKClick Then
      Call QueryData(True)
   End If
   Call glbDatabaseMngr.UnLockTable(m_TableName, id, IsCanLock, glbErrorLog)

End Sub

Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub

Private Sub cmdSearch_Click()
   Call QueryData(True)
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
Dim HeaderText As String

   ReportMode = 1
  Set oMenu = New cPopupMenu
  
   
   If DocumentType = 1 Then
      HeaderText = "ใบรายงานข้อมูล INCENTIVE รายสินค้า"
   ElseIf DocumentType = 2 Then
      HeaderText = "ใบรายงานข้อมูล INCENTIVE รายลูกค้า สินค้า"
   ElseIf DocumentType = 3 Then
      HeaderText = "ใบรายงานข้อมูล COMMISTION ของฟรีแลนซ์"
   ElseIf DocumentType = 4 Then
      HeaderText = "ใบรายงานข้อมูล INCENTIVE ของฟรีแลนซ์"
   End If
    lMenuChosen = oMenu.Popup(HeaderText, "ปรับค่าหน้ากระดาษ")
   If lMenuChosen = 0 Then
      Exit Sub
   End If

   Call EnableForm(Me, False)

   If lMenuChosen = 1 Then
      ReportKey = "CReportComIncentive"
      Set Report = New CReportComIncentive
      ReportFlag = True
      Call Report.AddParam(1, "PREVIEW_TYPE")
   End If

   If Not Report Is Nothing Then
      Call Report.AddParam(DocumentType, "DOCUMENT_TYPE")
      Call Report.AddParam(m_CollCI, "COMMISSION_INCENTIVE")
      Call Report.AddParam(HeaderText, "HEADER_TEXT")
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
   
   ReportKey = "CReportComIncentive"
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

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      
      Call InitIncentiveOrderBy(cboOrderBy)
      Call InitOrderType(cboOrderType)
      
      Call QueryData(True)
   End If
End Sub

Public Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long

   If Flag Then
      Call EnableForm(Me, False)

      m_CommissionIncentive.FREELANCE_CODE = txtFreelance.Text
      m_CommissionIncentive.PART_NO = txtStockCode.Text
      m_CommissionIncentive.OrderBy = cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex))
      m_CommissionIncentive.OrderType = cboOrderType.ItemData(Minus2Zero(cboOrderType.ListIndex))
      m_CommissionIncentive.DOCUMENT_TYPE = DocumentType
      If Not glbDaily.QueryCommissionIncentive(m_CommissionIncentive, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   Set m_CollCI = Nothing
   Set m_CollCI = New Collection
   While Not m_Rs.EOF
      Set m_TempCommissionIncentive = Nothing
     Set m_TempCommissionIncentive = New CCommissionIncentive
      Call m_TempCommissionIncentive.PopulateFromRS(1, m_Rs)
      Call m_CollCI.add(m_TempCommissionIncentive)
      m_Rs.MoveNext
   Wend
   m_Rs.MoveFirst
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If

   GridEX1.ItemCount = ItemCount
   GridEX1.Rebind

   Call EnableForm(Me, True)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.NAME
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 116 Then
      Call cmdSearch_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
      Call cmdClear_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 118 Then
      Call cmdAdd_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
      Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 123 Then
      Call AddMemoNote
      KeyCode = 0
   End If
End Sub

Private Sub InitGrid()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   
   GridEX1.FormatStyles.Clear
   Set fmsTemp = GridEX1.FormatStyles.add("N")
   fmsTemp.ForeColor = GLB_ALERT_COLOR
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 2000
   Col.Caption = MapText("รหัสฟรีแลนซ์")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5000
   Col.Caption = MapText("ชื่อฟรีแลนซ์")
   
   If DocumentType = 1 Then
      Set Col = GridEX1.Columns.add '4
      Col.Width = 0
      Col.Caption = MapText("ชื่อลูกค้า")
      
      Set Col = GridEX1.Columns.add '5
      Col.Width = 6000
      Col.Caption = MapText("รหัสสินค้า")
      
      Set Col = GridEX1.Columns.add '5
      Col.Width = 5000
      Col.Caption = MapText("ชื่อสินค้า")
      
      Set Col = GridEX1.Columns.add '6
      Col.Width = 1500
      Col.Caption = MapText("บาท/หน่วย")
         
      Set Col = GridEX1.Columns.add '7
      Col.Width = 0
      Col.Caption = MapText("ยอดตั้งแต่")
      
      Set Col = GridEX1.Columns.add '8
      Col.Width = 0
      Col.Caption = MapText("ถึงยอด")
      
      Set Col = GridEX1.Columns.add '9
      Col.Width = 0
      Col.Caption = MapText("ชนิดหน่วย")
      
      Set Col = GridEX1.Columns.add '10
      Col.Width = 0
      Col.Caption = MapText("ยอดเกิน")
      
      Set Col = GridEX1.Columns.add '11
      Col.Width = 0
      Col.Caption = MapText("คิดหน่วยล่ะ(บาท)")
   ElseIf DocumentType = 2 Then
       Set Col = GridEX1.Columns.add '4
      Col.Width = 5000
      Col.Caption = MapText("ชื่อลูกค้า")
      
      Set Col = GridEX1.Columns.add '5
      Col.Width = 6000
      Col.Caption = MapText("รหัสสินค้า")
      
      Set Col = GridEX1.Columns.add '5
      Col.Width = 5000
      Col.Caption = MapText("ชื่อสินค้า")
      
      Set Col = GridEX1.Columns.add '6
      Col.Width = 3000
      Col.Caption = MapText("บาท/หน่วย")
      
      Set Col = GridEX1.Columns.add '7
      Col.Width = 0
      Col.Caption = MapText("ยอดตั้งแต่")
      
      Set Col = GridEX1.Columns.add '8
      Col.Width = 0
      Col.Caption = MapText("ถึงยอด")
      
      Set Col = GridEX1.Columns.add '9
      Col.Width = 0
      Col.Caption = MapText("ชนิดหน่วย")
      
      Set Col = GridEX1.Columns.add '10
      Col.Width = 0
      Col.Caption = MapText("ยอดเกิน")
      
      Set Col = GridEX1.Columns.add '11
      Col.Width = 0
      Col.Caption = MapText("คิดหน่วยล่ะ(บาท)")
   ElseIf DocumentType = 3 Or DocumentType = 4 Then
       Set Col = GridEX1.Columns.add '4
      Col.Width = 0
      Col.Caption = MapText("ชื่อลูกค้า")
      
      Set Col = GridEX1.Columns.add '5
      Col.Width = 0
      Col.Caption = MapText("รหัสสินค้า")
      
      Set Col = GridEX1.Columns.add '5
      Col.Width = 0
      Col.Caption = MapText("ชื่อสินค้า")
      
      Set Col = GridEX1.Columns.add '6
      Col.Width = 3000
      Col.Caption = MapText("บาท/หน่วย")
      
      Set Col = GridEX1.Columns.add '7
      Col.Width = 3000
      Col.Caption = MapText("ยอดตั้งแต่")
      
      Set Col = GridEX1.Columns.add '8
      Col.Width = 3000
      Col.Caption = MapText("ถึงยอด")
      
      Set Col = GridEX1.Columns.add '9
      Col.Width = 1500
      Col.Caption = MapText("ชนิดหน่วย")
      
      Set Col = GridEX1.Columns.add '10
      Col.Width = 1500
      Col.Caption = MapText("ยอดเกิน")
      
      Set Col = GridEX1.Columns.add '11
      Col.Width = 2000
      Col.Caption = MapText("คิดหน่วยล่ะ(บาท)")
      
   End If
   
   
   GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitGrid
   
   Call InitNormalLabel(lblStockCode, MapText("สินค้า"))
   Call InitNormalLabel(lblFreelance, MapText("PC"))
   Call InitNormalLabel(lblOrderBy, MapText("เรียงตาม"))
   Call InitNormalLabel(lblOrderType, MapText("เรียงจาก"))
   
  txtStockCode.SetKeySearch ("FREELANCE_CODE")
  txtFreelance.SetKeySearch ("PART_NO")

   Call InitCombo(cboOrderBy)
   Call InitCombo(cboOrderType)
   
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPrint.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdClear.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdPrint, MapText("พิมพ์"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา (F5)"))
   Call InitMainButton(cmdClear, MapText("เคลียร์ (F4)"))
End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   m_TableName = "USER_GROUP"
   
   Set m_CommissionIncentive = New CCommissionIncentive
   Set m_TempCommissionIncentive = New CCommissionIncentive
   Set m_CollCI = New Collection
   Set m_Rs = New ADODB.Recordset

   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_CollCI = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   RowBuffer.RowStyle = RowBuffer.Value(4)
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"

   If m_Rs Is Nothing Then
      Exit Sub
   End If

   If m_Rs.State <> adStateOpen Then
      Exit Sub
   End If

   If m_Rs.EOF Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If
   
   Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
   Call m_TempCommissionIncentive.PopulateFromRS(1, m_Rs)

   Values(1) = m_TempCommissionIncentive.INCENTIVE_ID
   Values(2) = m_TempCommissionIncentive.FREELANCE_CODE
   Values(3) = m_TempCommissionIncentive.FREELANCE_NAME & " " & m_TempCommissionIncentive.FREELANCE_LASTNAME
   Values(4) = m_TempCommissionIncentive.CUSTOMER_NAME & " " & m_TempCommissionIncentive.CUSTOMER_LASTNAME
   Values(5) = m_TempCommissionIncentive.PART_ITEM_CODE
   Values(6) = m_TempCommissionIncentive.PART_ITEM_NAME
   Values(7) = m_TempCommissionIncentive.INCENTIVE_PER_PACK
   Values(8) = FormatNumber(m_TempCommissionIncentive.FROM_AMOUNT)
   Values(9) = FormatNumber(m_TempCommissionIncentive.TO_AMOUNT)
    Values(10) = TypeUnitFlag(m_TempCommissionIncentive.UNIT_TYPE)

   If m_TempCommissionIncentive.AMOUNT_OVER_FLAG = "Y" Then
      Values(11) = Values(8)
      Values(12) = m_TempCommissionIncentive.RATE_OVER
   Else
     Values(11) = ""
     Values(12) = ""
   End If
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.HEIGHT = ScaleHeight
   
   pnlHeader.Width = ScaleWidth
   GridEX1.Width = ScaleWidth - 2 * GridEX1.Left
   GridEX1.HEIGHT = ScaleHeight - GridEX1.Top - 620
   cmdAdd.Top = ScaleHeight - 580
   cmdEdit.Top = ScaleHeight - 580
   cmdDelete.Top = ScaleHeight - 580
   cmdPrint.Top = ScaleHeight - 580
   cmdOK.Top = ScaleHeight - 580
   cmdExit.Top = ScaleHeight - 580
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
   cmdPrint.Left = cmdOK.Left - cmdPrint.Width - 50
End Sub
Public Sub RefreshGrid()
   Call QueryData(True)
End Sub

