VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmWinPricingMain 
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   11910
   Icon            =   "frmWinPricingMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3300
      Top             =   1110
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":11A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":1A7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":2358
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":24B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":2D8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":3666
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":3980
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":425A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":4B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":540E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":5CE8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   795
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   1402
      _Version        =   131073
      BackStyle       =   1
      Begin VB.Label lblCompany 
         Alignment       =   2  'Center
         Caption         =   "lblCompany"
         Height          =   465
         Left            =   2880
         TabIndex        =   12
         Top             =   480
         Width           =   6765
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   555
         Left            =   9660
         TabIndex        =   8
         Top             =   6390
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   979
         _Version        =   131073
         PictureFrames   =   1
         Picture         =   "frmWinPricingMain.frx":5FF7
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin VB.Label lblDateTime 
         Alignment       =   2  'Center
         Caption         =   "Label1"
         Height          =   315
         Left            =   9390
         TabIndex        =   7
         Top             =   30
         Width           =   2505
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   7755
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   13679
      _Version        =   131073
      BackStyle       =   1
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   2040
         ScaleHeight     =   1215
         ScaleWidth      =   1185
         TabIndex        =   11
         Top             =   4920
         Width           =   1185
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   0
         Top             =   0
      End
      Begin MSComctlLib.TreeView trvMain 
         Height          =   3645
         Left            =   0
         TabIndex        =   1
         Top             =   1230
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   6429
         _Version        =   393217
         Indentation     =   882
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "JasmineUPC"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblLastVersion2 
         Caption         =   "Label1"
         Height          =   465
         Left            =   360
         TabIndex        =   17
         Top             =   6480
         Width           =   1365
      End
      Begin VB.Label lblLastVersion 
         Caption         =   "Label1"
         Height          =   465
         Left            =   1800
         TabIndex        =   16
         Top             =   6480
         Width           =   2445
      End
      Begin Threed.SSCommand cmdPasswd 
         Height          =   465
         Left            =   810
         TabIndex        =   10
         Top             =   7170
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   820
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   465
         Left            =   2400
         TabIndex        =   9
         Top             =   7170
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   820
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblVersion 
         Caption         =   "Label1"
         Height          =   465
         Left            =   360
         TabIndex        =   6
         Top             =   6000
         Width           =   4005
      End
      Begin VB.Label lblUserGroup 
         Caption         =   "Label1"
         Height          =   465
         Left            =   360
         TabIndex        =   5
         Top             =   5490
         Width           =   3045
      End
      Begin VB.Label lblUserName 
         Caption         =   "Label1"
         Height          =   465
         Left            =   360
         TabIndex        =   4
         Top             =   4980
         Width           =   3045
      End
   End
   Begin Threed.SSPanel pnlHeader 
      Height          =   795
      Left            =   4560
      TabIndex        =   3
      Top             =   840
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   1402
      _Version        =   131073
      BackStyle       =   1
   End
   Begin Threed.SSFrame fraGeneric 
      Height          =   1455
      Left            =   4800
      TabIndex        =   13
      Top             =   1920
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   2566
      _Version        =   131073
      BackStyle       =   1
      Begin Threed.SSCommand cmdGeneric 
         Height          =   885
         Index           =   0
         Left            =   720
         TabIndex        =   14
         Top             =   300
         Visible         =   0   'False
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   1561
         _Version        =   131073
         Caption         =   "SSCommand2"
      End
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   915
      Left            =   4320
      TabIndex        =   15
      Top             =   6870
      Visible         =   0   'False
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   1614
      _Version        =   131073
      Caption         =   "SSCommand2"
   End
End
Attribute VB_Name = "frmWinPricingMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const ROOT_TREE = "Root"
Private m_Sp As CSystemParam
Private MustAsk As Boolean
Private m_HasActivate As Boolean
Private m_Rs  As ADODB.Recordset

Private m_TableName As String

Public HeaderText As String
Private m_XCollection As CXCollection
Private m_MustAsk As Boolean

Private m_PartGroupMenus As Collection
Private m_JobProcessMenus As Collection

'*********************************
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_SHOWNOACTIVATE = 4

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

'*************************************

Private Sub InitMainTreeview()
Dim Node As Node
Dim NewNodeID As String
   
   trvMain.Nodes.Clear
   trvMain.Font.NAME = GLB_FONT
   trvMain.Font.Size = 14
   trvMain.Font.Bold = False
   
   
   Set Node = trvMain.Nodes.add(, tvwFirst, ROOT_TREE, MapText("ระบบงานทั้งหมด"), 1)
   Node.Expanded = True
   Node.Selected = True
   
   #If LIMIT_AREA <> 1 Then
      '==
      Set Node = trvMain.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-0", MapText("ระบบข้อมูลผู้ใช้งาน"), 4, 4)
      Node.Expanded = False
      '==
      
      Set Node = trvMain.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-1", MapText("ระบบข้อมูลหลัก"), 2, 2)
      Node.Expanded = False
   
      Set Node = trvMain.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-2", MapText("ระบบข้อมูลส่วนกลาง"), 6, 6)
      Node.Expanded = False
   
      If glbGuiConfigs.VerifyGuiConfig("HR_VIEW") Then
         Set Node = trvMain.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-8", MapText("ระบบบริหารฝ่ายบุคคล"), 12, 12)
         Node.Expanded = False
      End If
   
      If glbGuiConfigs.VerifyGuiConfig("PACKAGE_VIEW") Then
         Set Node = trvMain.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-7", MapText("ระบบแพคเกจสินค้า/บริการ"), 5, 5)
         Node.Expanded = False
      End If
   
      If glbGuiConfigs.VerifyGuiConfig("INVENTORY_VIEW") Then
         Set Node = trvMain.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-3", MapText("ระบบบริหารคลัง"), 3, 3)
         Node.Expanded = False
      End If
      
      If glbGuiConfigs.VerifyGuiConfig("INVENTORY-WH_VIEW") Then
         Set Node = trvMain.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-11", MapText("ระบบบริหารคลังสินค้า"), 3, 3)
         Node.Expanded = False
      End If
   
      If glbGuiConfigs.VerifyGuiConfig("PRODUCTION_VIEW") Then
         Set Node = trvMain.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-9", MapText("ระบบการผลิต"), 10, 10)
         Node.Expanded = False
      End If
   #End If
   
   If glbGuiConfigs.VerifyGuiConfig("LEDGER_VIEW") Then
      Set Node = trvMain.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-5", MapText("ระบบบริหารบัญชี"), 8, 8)
      Node.Expanded = False
   End If
      
   If glbGuiConfigs.VerifyGuiConfig("PLAN_VIEW") Then
      Set Node = trvMain.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-4", MapText("ระบบประมาณการ/วางแผน"), 12, 12)
      Node.Expanded = False
   End If
   
   If glbGuiConfigs.VerifyGuiConfig("COMMISSION_VIEW") Then
      Set Node = trvMain.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-10", MapText("ระบบคอมมิตชั่น"), 11, 11)
      Node.Expanded = False
   End If
End Sub

Private Sub InitFormLayout()
   Call InitNormalLabel(lblUsername, MapText("ผู้ใช้ : "), RGB(0, 0, 255))
   Call InitNormalLabel(lblUserGroup, MapText("กลุ่มผู้ใช้ : "), RGB(0, 0, 255))
   Call InitNormalLabel(lblVersion, MapText("เวอร์ชั่นปัจจุบัน:") & glbParameterObj.Version & " (" & glbParameterObj.Programowner & ") ", RGB(0, 0, 255))
   Dim LVP As String
   LVP = CheckLastVersionProgram(glbParameterObj.Version)
    Call InitNormalLabel(lblLastVersion2, MapText("เวอร์ชั่นใหม่:"), RGB(0, 0, 255))
   If LVP > glbParameterObj.Version Then
      Call InitNormalLabel(lblLastVersion, LVP & " (" & glbParameterObj.Programowner & ") ", RGB(255, 0, 0))
   Else
      Call InitNormalLabel(lblLastVersion, LVP & " (" & glbParameterObj.Programowner & ") ", RGB(0, 0, 255))
   End If
   Call InitNormalLabel(lblDateTime, "", RGB(0, 0, 255))
   lblDateTime.BackStyle = 1
   lblDateTime.BackColor = RGB(255, 255, 255)
   Call InitNormalLabel(lblCompany, MapText(glbEnterPrise.ENTERPRISE_NAME & "  " & glbEnterPrise.BRANCH_NAME))
'   Me.Picture = LoadPicture(glbParameterObj.NormalForm1)
   Me.BackColor = RGB(210, 240, 250)
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPasswd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   
   Me.Caption = glbGuiConfigs.ShowWindowCaption(glbUser.USER_NAME & " " & glbParameterObj.Programowner)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19

'   Call InitMainButton(cmdUserGroup, MapText("ข้อมูลกลุ่มผู้ใช้งาน"))
'   Call InitMainButton(cmdUser, MapText("ข้อมูลผู้ใช้งาน"))
'   Call InitMainButton(cmdAdminReport, MapText("รายงานข้อมูลผู้ใช้งาน"))
'
'   Call InitMainButton(cmdMaster1, MapText("ข้อมูลหลักส่วนกลาง"))
'   Call InitMainButton(cmdMaster2, MapText("ข้อมูลหลักระบบคลัง"))
'   Call InitMainButton(cmdMaster3, MapText("ข้อมูลหลักระบบฝ่ายบุคคล"))
'   Call InitMainButton(cmdMaster4, MapText("ข้อมูลหลักระบบบริหารบัญชี"))
'   Call InitMainButton(cmdMaster5, MapText("ข้อมูลหลักงานขาย"))
'   cmdMaster5.Visible = False
'   Call InitMainButton(cmdPricePlanMaster, MapText("ข้อมูลหลักแพคเกจสินค้า/บริการ"))
'   Call InitMainButton(cmdMaster6, MapText("ข้อมูลหลักระบบการผลิต"))
'
'   Call InitMainButton(cmdMainEnterprise, MapText("ข้อมูลบริษัท"))
'   Call InitMainButton(cmdMainCustomer, MapText("ข้อมูลลูกค้า"))
'   Call InitMainButton(cmdMainSupplier, MapText("ข้อมูลซัพพลายเออร์"))
'   Call InitMainButton(cmdMainEmployee, MapText("ข้อมูลพนักงาน"))
'   Call InitMainButton(cmdMainReport, MapText("รายงานข้อมูลกลาง"))
'
'   Call InitMainButton(cmdRawMatterial, MapText("ข้อมูลสินค้าและวัตถุดิบ"))
'   Call InitMainButton(cmdImport, MapText("ข้อมูลการรับเข้าวัตถุดิบ"))
'   Call InitMainButton(cmdExport, MapText("ข้อมูลการเบิกวัตถุดิบ"))
'   Call InitMainButton(cmdTransfer, MapText("ข้อมูลการโอนย้ายวัตถุดิบ"))
'   Call InitMainButton(cmdAdjust, MapText("ข้อมูลการปรับยอดคลัง"))
'   Call InitMainButton(cmdInventoryReport, MapText("รายงานระบบคลัง"))
'
'   Call InitMainButton(cmdPigWeek, MapText("ข้อมูลรหัสสัปดาห์เกิดสุกร"))
'   Call InitMainButton(cmdPigBirth, MapText("ข้อมูลสุกรคลอด"))
'   Call InitMainButton(cmdPigTransfer, MapText("ข้อมูลการโอนย้ายสุกร"))
'   Call InitMainButton(cmdPigAdjustment, MapText("ข้อมูลการปรับยอดสุกร"))
'   Call InitMainButton(cmdPigReport, MapText("รายงานระบบบริหารสุกร"))
'
'   Call InitMainButton(cmdCurrencyExchange, MapText("ข้อมูลอัตราการแลกเปลี่ยนเงินตรา"))
'   Call InitMainButton(cmdBuy, MapText("ระบบงานซื้อ (รายจ่าย)"))
'   Call InitMainButton(cmdSell, MapText("ระบบงานขาย"))
'   Call InitMainButton(cmdPayment, MapText("ระบบข้อมูลเงินสด"))
'   Call InitMainButton(cmdLedgerReport, MapText("รายงานระบบบัญชี"))
'
'   Call InitMainButton(cmdGldDailyPrice, MapText("ราคาทองประจำวัน"))
'   Call InitMainButton(cmdGoldWage, MapText("ข้อมูลค่าแรงช่างทอง"))
'   Call InitMainButton(cmdGldSaleBuy, MapText("ระบบซื้อขายทอง"))
'   Call InitMainButton(cmdGldReport, MapText("รายงานระบบร้านทอง"))
'
'   Call InitMainButton(cmdFeature, MapText("ข้อมูลสินค้า/บริการ"))
'   Call InitMainButton(cmdSoc, MapText("ข้อมูลแพคเกจสินค้า/บริการ"))
'   Call InitMainButton(cmdPackageReport, MapText("รายงานแพคเกจสินค้า/บริการ"))
'
'   Call InitMainButton(cmdDataPerson, MapText("ข้อมูลพนักงาน"))
'   Call InitMainButton(cmdMoneyPerson, MapText("เงินยืมส่วนบุคคล"))
'   Call InitMainButton(cmdSalarySlipt, MapText("สลิปเงินเดือน"))
'   Call InitMainButton(cmdReportPerson, MapText("รายงานระบบฝ่ายบุคคล"))
'
'   Call InitMainButton(cmdProductionFormula, MapText("ข้อมูลสูตรการผลิต"))
'   Call InitMainButton(cmdProductionJob, MapText("ข้อมูลใบสั่งผลิต"))
'   Call InitMainButton(cmdProductionEstimate, MapText("ข้อมูลใบประเมินราคา"))
'   Call InitMainButton(cmdProductionReport, MapText("รายงานระบบการผลิต"))

   Call InitMainButton(cmdExit, MapText("ออก"))
   Call InitMainButton(cmdPasswd, MapText("โปรแกรม"))
   
   Picture1.Visible = glbGuiConfigs.VerifyGuiConfig("LOGO_VIEW")
   If glbGuiConfigs.VerifyGuiConfig("LOGO_VIEW") Then
      Picture1.Picture = LoadPicture(glbParameterObj.CompanyLogo)
   End If
   
   Call InitMainTreeview
End Sub
Private Sub cmdExit_Click()
   Unload Me
End Sub

''Public Sub GeneratePartGroupMenu(Col As Collection)
''Dim G As CPartGroup
''Dim D As CMenuItem
''Dim TempRs As ADODB.Recordset
''Dim iCount As Long
''
''   Set G = New CPartGroup
''   Set TempRs = New ADODB.Recordset
''
''   G.PART_GROUP_ID = -1
''   Call G.QueryData(TempRs, iCount)
''
''   While Not TempRs.EOF
''      Call G.PopulateFromRS(1, TempRs)
''
''      Set D = New CMenuItem
''      D.KEYWORD = G.PART_GROUP_NAME
''      D.KEY_ID = G.PART_GROUP_ID
''      Call Col.add(D)
''      Set D = Nothing
''
''      TempRs.MoveNext
''   Wend
''
''   If TempRs.State = adStateOpen Then
''      Call TempRs.Close
''   End If
''   Set TempRs = Nothing
''   Set G = Nothing
''End Sub
''
''Private Sub GenerateJobProcessMenu(Col As Collection)
''Dim G As CProcess
''Dim D As CMenuItem
''Dim TempRs As ADODB.Recordset
''Dim iCount As Long
''
''   Set G = New CProcess
''   Set TempRs = New ADODB.Recordset
''
''   G.PROCESS_ID = -1
''   Call G.QueryData(TempRs, iCount)
''
''   While Not TempRs.EOF
''      Call G.PopulateFromRS(1, TempRs)
''
''      Set D = New CMenuItem
''      D.KEYWORD = G.PROCESS_NAME
''      D.KEY_ID = G.PROCESS_ID
''      Call Col.add(D)
''      Set D = Nothing
''
''      TempRs.MoveNext
''   Wend
''
''   If TempRs.State = adStateOpen Then
''      Call TempRs.Close
''   End If
''   Set TempRs = Nothing
''   Set G = Nothing
''End Sub

Private Sub cmdGeneric_Click(Index As Integer)
Dim Key As String
Dim Caption As String
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu
Dim lNewmenu As Long
Dim DocumentType As Long
Dim D As CMenuItem
Dim DocumentTypeDesc As String
   
   Set oMenu = New cPopupMenu
   
   Key = cmdGeneric(Index).Tag
   Caption = cmdGeneric(Index).Caption
   If Key = "ADMIN_GROUP" Then
    If Not VerifyAccessRight("ADMIN_GROUP") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      Load frmUserGroup
      frmUserGroup.Show 1
      Unload frmUserGroup
      Set frmUserGroup = Nothing
   ElseIf Key = "ADMIN_USER" Then
      If Not VerifyAccessRight("ADMIN_USER") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      Load frmUser
      frmUser.Show 1
      Unload frmUser
      Set frmUser = Nothing
   ElseIf Key = "ADMIN_REPORT" Then
      If Not VerifyAccessRight("ADMIN_REPORT") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      frmSummaryReport.MasterMode = 1
      frmSummaryReport.HeaderText = Caption
      Load frmSummaryReport
      frmSummaryReport.Show 1
      Unload frmSummaryReport
      Set frmSummaryReport = Nothing
   ElseIf Key = "PACKAGE_FEATURE" Then
     If Not VerifyAccessRight("PACKAGE_FEATURE") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   
      Load frmFeature
      frmFeature.Show 1
      Unload frmFeature
      Set frmFeature = Nothing
   ElseIf Key = "PACKAGE_SOC" Then
      If Not VerifyAccessRight("PACKAGE_SOC") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      Load frmSoc
      frmSoc.Show 1
      Unload frmSoc
      Set frmSoc = Nothing
   ElseIf Key = "PACKAGE_REPORT" Then
      If Not VerifyAccessRight("PACKAGE_REPORT") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      frmSummaryReport.HeaderText = Caption
      frmSummaryReport.MasterMode = 2
      Load frmSummaryReport
      frmSummaryReport.Show 1
      Unload frmSummaryReport
      Set frmSummaryReport = Nothing
   ElseIf Key = "INVENTORY_PART-MASTER" Then
    If Not VerifyAccessRight("INVENTORY_PART-MASTER") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.Popup("สินค้าสำเร็จรูป", "-", "สินค้าผลิตเสร็จ")
      If lMenuChosen <= 0 Then
         Call EnableForm(Me, True)
         Exit Sub
      Else
         If lMenuChosen = 1 Then
            DocumentType = 10
         ElseIf lMenuChosen = 3 Then
            DocumentType = 21
         End If
      End If
      
      frmPartMaster.DocumentType = DocumentType
      Load frmPartMaster
      frmPartMaster.Show 1

      Unload frmPartMaster
      Set frmPartMaster = Nothing
   ElseIf Key = "INVENTORY_PART" Then
      If Not VerifyAccessRight("INVENTORY_PART") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If

      lMenuChosen = oMenu.AddMenu(m_PartGroupMenus)
      If lMenuChosen = 0 Then
         Exit Sub
      End If

      frmPartItem.PartGroupID = lMenuChosen
      Load frmPartItem
      frmPartItem.Show 1
      
      Unload frmPartItem
      Set frmPartItem = Nothing
   ElseIf Key = "INVENTORY_IMPORT" Then
      If Not VerifyAccessRight("INVENTORY_IMPORT") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      Set m_Sp = GetSystemParam(glbSystemParams, "SHOW_IMPORT_INVENTORY")
      
      If m_Sp.PARAM_VALUE = "N" Then
         glbErrorLog.LocalErrorMsg = "โปรแกรมไม่สนับสนุนฟังก์ชันนี้ในเวอร์ชันนี้"
         glbErrorLog.ShowUserError
         Exit Sub
      End If
      
      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.Popup("ใบรับเข้าวัตถุดิบ", "-", "ใบรับเข้าวัสดุอุปกรณ์", "-", "ใบรับเข้าจ่ายออกวัสดุอุปกรณ์", "-", "ใบรับเข้าทั่วไป")
      If lMenuChosen <= 0 Then
         Call EnableForm(Me, True)
         Exit Sub
      Else
         If lMenuChosen = 1 Then
            DocumentType = 1
         ElseIf lMenuChosen = 3 Then
            DocumentType = 19
         ElseIf lMenuChosen = 5 Then
            DocumentType = 20
         ElseIf lMenuChosen = 7 Then
            DocumentType = 23
         End If
      End If
      
      frmInventoryDoc1.DocumentType = DocumentType
      Load frmInventoryDoc1
      frmInventoryDoc1.Show 1
      
      Unload frmInventoryDoc1
      Set frmInventoryDoc1 = Nothing
      
      Set oMenu = Nothing
   ElseIf Key = "INVENTORY_EXPORT" Then
      If Not VerifyAccessRight("INVENTORY_EXPORT") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      Load frmInventoryDoc2
      frmInventoryDoc2.Show 1
      
      Unload frmInventoryDoc2
      Set frmInventoryDoc2 = Nothing
   ElseIf Key = "INVENTORY_TRANSFER" Then
      If Not VerifyAccessRight("INVENTORY_TRANSFER") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.Popup("ใบโอนวัตถุดิบระหว่างคลัง", "-", "ใบโอนเปลี่ยนวัตถุดิบ")
      If lMenuChosen <= 0 Then
         Call EnableForm(Me, True)
         Exit Sub
      Else
         If lMenuChosen = 1 Then
            DocumentType = 3
         ElseIf lMenuChosen = 3 Then
            DocumentType = 22
         End If
      End If
      
      frmInventoryDoc3.DocumentType = DocumentType
      Load frmInventoryDoc3
      frmInventoryDoc3.Show 1
   
      Unload frmInventoryDoc3
      Set frmInventoryDoc3 = Nothing
   
   
   ElseIf Key = "INVENTORY_ADJUST" Then
      If Not VerifyAccessRight("INVENTORY_ADJUST") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.Popup("ใบปรับยอดเนื่องจากการตรวจนับ", "-", "ใบปรับยอดเนื่องจากการชั่งตวง")
      If lMenuChosen <= 0 Then
         Call EnableForm(Me, True)
         Exit Sub
      End If

      If lMenuChosen = 1 Then
         frmInventoryDoc4.DocumentType = 4
         Load frmInventoryDoc4
         frmInventoryDoc4.Show 1
      
         Unload frmInventoryDoc4
         Set frmInventoryDoc4 = Nothing
      ElseIf lMenuChosen = 3 Then
         frmInventoryDoc4.DocumentType = 5
         Load frmInventoryDoc4
         frmInventoryDoc4.Show 1
      
         Unload frmInventoryDoc4
         Set frmInventoryDoc4 = Nothing
      End If
   ElseIf Key = "INVENTORY_ACTUAL" Then
      If Not VerifyAccessRight("INVENTORY_ACTUAL") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   
      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.Popup("คลังวัตถุดิบ", "-", "ห้องยา", "-", "ไซโล")
      If lMenuChosen <= 0 Then
         Call EnableForm(Me, True)
         Exit Sub
      Else
         If lMenuChosen = 1 Then 'คลังวัตถุดิบ
            DocumentType = 1
         ElseIf lMenuChosen = 3 Then 'ห้องยา
            DocumentType = 2
         ElseIf lMenuChosen = 5 Then  'ไซโล
            DocumentType = 3
         ElseIf lMenuChosen = 7 Then  'ไซโล
            DocumentType = 3
         End If
      End If 'raw material
   
   If Not VerifyAccessRight("INVENTORY_ACTUAL_" & InventoryActArea2Text2(DocumentType), InventoryActArea2Text(DocumentType)) Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
    
      
      frmInventoryAct.HeaderText = InventoryActArea2Text(DocumentType)
      frmInventoryAct.InventoryActArea = DocumentType
      Load frmInventoryAct
      frmInventoryAct.Show 1
   
      Unload frmInventoryAct
      Set frmInventoryAct = Nothing
   ElseIf Key = "INVENTORY-WH_ACTUAL" Then
      If Not VerifyAccessRight("INVENTORY-WH_ACTUAL") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   
      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.Popup("คลังสินค้า BAG", "-", "คลังสินค้า BULK")
      If lMenuChosen <= 0 Then
         Call EnableForm(Me, True)
         Exit Sub
      Else
         If lMenuChosen = 1 Then 'คลังสินค้า(BAG)
            DocumentType = 14
            DocumentTypeDesc = "คลังสินค้า BAG"
         ElseIf lMenuChosen = 3 Then 'คลังสินค้า(BULK)
            DocumentType = 13
            DocumentTypeDesc = "คลังสินค้า BULK"
         End If
      End If
   
   If Not VerifyAccessRight("INVENTORY-WH_ACTUAL_" & InventoryWhActArea2Text2(DocumentType), InventoryWhActArea2Text(DocumentType)) Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
    
      
      frmInventoryWhAct.HeaderText = InventoryWhActArea2Text(DocumentType)
      frmInventoryWhAct.InventoryWhActArea = DocumentType
      frmInventoryWhAct.HeaderText = DocumentTypeDesc
      Load frmInventoryWhAct
      frmInventoryWhAct.Show 1
   
      Unload frmInventoryWhAct
      Set frmInventoryWhAct = Nothing
   ElseIf Key = "INVENTORY_REPORT" Then
      If Not VerifyAccessRight("INVENTORY_REPORT") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      frmSummaryReport.HeaderText = Caption
      frmSummaryReport.MasterMode = 4
      Load frmSummaryReport
      frmSummaryReport.Show 1
      
      Unload frmSummaryReport
       Set frmSummaryReport = Nothing
ElseIf Key = "INVENTORY-WH_REPORT" Then
      If Not VerifyAccessRight("INVENTORY-WH_REPORT", "ระบบรายงานคลังสินค้า") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      frmSummaryReport.HeaderText = Caption
      frmSummaryReport.MasterMode = 9
      Load frmSummaryReport
      frmSummaryReport.Show 1
      
      Unload frmSummaryReport
       Set frmSummaryReport = Nothing
   
   ElseIf Key = "INVENTORY-WH_IMPORT" Then
      If Not VerifyAccessRight("INVENTORY-WH_IMPORT", "ข้อมูลการรับเข้า") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      Dim strNameMenu As String
      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.Popup("ใบบรรจุอาหาร BAG", "-", "ใบบรรจุอาหาร BULK", "-", "ใบบรรจุอาหาร RE-BAG -> BAG", "-", "ใบบรรจุอาหาร RE-BAG -> BULK", "-", "ใบบรรจุอาหาร RE-BAG -> RM and Other")
      If lMenuChosen <= 0 Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
    
      If lMenuChosen = 1 Then
         DocumentType = 14 'บรรจุBAG
         strNameMenu = "ใบบรรจุอาหาร BAG"
         If Not VerifyAccessRight("INVENTORY-WH_IMPORT" & "_" & DocumentType, strNameMenu) Then
            Call EnableForm(Me, True)
            Exit Sub
          End If
      
         frmJob.HeaderText = strNameMenu
         frmJob.mainText = strNameMenu
         frmJob.ProcessID = 2
         frmJob.JobDocType = 1
         frmJob.DOCUMENT_TYPE = DocumentType
         Load frmJob
         frmJob.Show 1
         
         Unload frmJob
         Set frmJob = Nothing
      ElseIf lMenuChosen = 3 Then
         DocumentType = 13 'บรรจุ Bulk
         strNameMenu = "ใบบรรจุอาหาร BULK"
          If Not VerifyAccessRight("INVENTORY-WH_IMPORT" & "_" & DocumentType, strNameMenu) Then
            Call EnableForm(Me, True)
            Exit Sub
          End If
         frmJob.HeaderText = strNameMenu
         frmJob.mainText = strNameMenu
         frmJob.ProcessID = 4
         frmJob.JobDocType = 1
         frmJob.DOCUMENT_TYPE = DocumentType
         Load frmJob
         frmJob.Show 1
         
         Unload frmJob
         Set frmJob = Nothing
      ElseIf lMenuChosen = 5 Then
         DocumentType = 17 'บรรจุ RE-BAG -> BAG
         strNameMenu = "ใบบรรจุอาหาร RE-BAG -> BAG"
          If Not VerifyAccessRight("INVENTORY-WH_IMPORT" & "_" & DocumentType, strNameMenu) Then
            Call EnableForm(Me, True)
            Exit Sub
          End If
         frmJob.DOCUMENT_TYPE = DocumentType
         frmJob.HeaderText = strNameMenu
         frmJob.mainText = strNameMenu
         frmJob.ProcessID = 6 'บรรจุ RE-BAG ใน job
         frmJob.JobDocType = 1
         Load frmJob
         frmJob.Show 1
         
         Unload frmJob
         Set frmJob = Nothing
       ElseIf lMenuChosen = 7 Then
         DocumentType = 18 'บรรจุ RE-BAG -> BULK
         strNameMenu = "ใบบรรจุอาหาร RE-BAG -> BULK"
         If Not VerifyAccessRight("INVENTORY-WH_IMPORT" & "_" & DocumentType, strNameMenu) Then
            Call EnableForm(Me, True)
            Exit Sub
          End If
         frmJob.DOCUMENT_TYPE = DocumentType
         frmJob.HeaderText = strNameMenu
         frmJob.mainText = strNameMenu
         frmJob.ProcessID = 7 'บรรจุ RE-BAG ใน job
         frmJob.JobDocType = 1
         Load frmJob
         frmJob.Show 1
         
         Unload frmJob
         Set frmJob = Nothing
      ElseIf lMenuChosen = 9 Then
         DocumentType = 19 'บรรจุ RE-BAG -> RM and Other
         strNameMenu = "ใบบรรจุอาหาร RE-BAG -> RM and Other"
         If Not VerifyAccessRight("INVENTORY-WH_IMPORT" & "_" & DocumentType, strNameMenu) Then
            Call EnableForm(Me, True)
            Exit Sub
          End If
         frmJob.DOCUMENT_TYPE = DocumentType
         frmJob.HeaderText = strNameMenu
         frmJob.mainText = strNameMenu
         frmJob.ProcessID = 8 'บรรจุ RE-BAG ใน job
         frmJob.JobDocType = 1
         Load frmJob
         frmJob.Show 1
         
         Unload frmJob
         Set frmJob = Nothing
      End If
   ElseIf Key = "INVENTORY-WH_EXPORT" Then
      If Not VerifyAccessRight("INVENTORY-WH_EXPORT") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
'
      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.Popup("ใบขึ้นอาหาร BAG", "-", "ใบขึ้นอาหาร BULK", "-", "ใบขึ้นอาหาร อื่นๆ")
      If lMenuChosen <= 0 Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
'
      If lMenuChosen = 1 Then
         If Not VerifyAccessRight("INVENTORY-WH_EXPORT_2000", "ใบขึ้นอาหาร BAG") Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
         frmInventoryDocWH.DocumentType = 2000
         frmInventoryDocWH.HeaderText = "ใบขึ้นอาหาร BAG"
         Load frmInventoryDocWH
         frmInventoryDocWH.Show 1

         Unload frmInventoryDocWH
         Set frmInventoryDocWH = Nothing
      ElseIf lMenuChosen = 3 Then
         If Not VerifyAccessRight("INVENTORY-WH_EXPORT_2001", "ใบขึ้นอาหาร BULK") Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
         frmInventoryDocWH.DocumentType = 2001
         frmInventoryDocWH.HeaderText = "ใบขึ้นอาหาร BULK"
         Load frmInventoryDocWH
         frmInventoryDocWH.Show 1

         Unload frmInventoryDocWH
         Set frmInventoryDocWH = Nothing
      ElseIf lMenuChosen = 5 Then
         If Not VerifyAccessRight("INVENTORY-WH_EXPORT_2004", "ใบขึ้นอาหาร อื่นๆ") Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
         frmInventoryDocWH.DocumentType = 2004
         frmInventoryDocWH.HeaderText = "ใบขึ้นอาหาร อื่นๆ"
         Load frmInventoryDocWH
         frmInventoryDocWH.Show 1

         Unload frmInventoryDocWH
         Set frmInventoryDocWH = Nothing
      End If
   ElseIf Key = "INVENTORY-WH_TRANSFER" Then
      If Not VerifyAccessRight("INVENTORY-WH_TRANSFER", "ใบโอนย้ายสินค้าโกดังอาหาร") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If

      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.Popup("ใบโอนเปลี่ยนสินค้า BAG", "-", "ใบโอนเปลี่ยนสินค้า BULK")
      If lMenuChosen <= 0 Then
         Call EnableForm(Me, True)
         Exit Sub
      Else
         If lMenuChosen = 1 Then
            DocumentType = 20
            frmInventoryDocWH.HeaderText = "ใบโอนเปลี่ยนสินค้า BAG"
            
            If Not VerifyAccessRight("INVENTORY-WH_TRANSFER" & "_BAG", "ใบโอนเปลี่ยนสินค้า BAG") Then
               Call EnableForm(Me, True)
               Exit Sub
            End If
         ElseIf lMenuChosen = 3 Then
            DocumentType = 21
            frmInventoryDocWH.HeaderText = "ใบโอนเปลี่ยนสินค้า BULK"
            
             If Not VerifyAccessRight("INVENTORY-WH_TRANSFER" & "_BULK", "ใบโอนเปลี่ยนสินค้า BULK") Then
               Call EnableForm(Me, True)
               Exit Sub
            End If
         End If
      End If
      
      frmInventoryDocWH.DocumentType = DocumentType
      Load frmInventoryDocWH
      frmInventoryDocWH.Show 1

      Unload frmInventoryDocWH
      Set frmInventoryDocWH = Nothing
   ElseIf Key = "INVENTORY-WH_ADJUST" Then
      If Not VerifyAccessRight("INVENTORY-WH_ADJUST", "ข้อมูลการปรับยอดคงเหลือ") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.Popup("ใบปรับยอดคงเหลือ BAG", "-", "ใบปรับยอดคงเหลือ BULK")
      If lMenuChosen <= 0 Then
         Call EnableForm(Me, True)
         Exit Sub
      End If

      If lMenuChosen = 1 Then
         frmInventoryDoc5.DocumentType = 15 'ปรับยอด BAG
         frmInventoryDoc5.HeaderText = "ใบปรับยอดคงเหลือ BAG"
         Load frmInventoryDoc5
         frmInventoryDoc5.Show 1
      
         Unload frmInventoryDoc5
         Set frmInventoryDoc5 = Nothing
      ElseIf lMenuChosen = 3 Then
         frmInventoryDoc5.DocumentType = 16 'ปรับยอด Bulk
         frmInventoryDoc5.HeaderText = "ใบปรับยอดคงเหลือ BULK"
         Load frmInventoryDoc5
         frmInventoryDoc5.Show 1
      
         Unload frmInventoryDoc5
         Set frmInventoryDoc5 = Nothing
      End If
        
    ElseIf Key = "INVENTORY-WH_STOCK" Then
      If Not VerifyAccessRight("INVENTORY-WH_STOCK", "ข้อมูลสินค้าคงเหลือ") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.Popup("สินค้าคงเหลือรายสินค้า", "สินค้าคงเหลือรายวัน")
      If lMenuChosen <= 0 Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      If lMenuChosen = 1 Then
         If Not VerifyAccessRight("INVENTORY-WH_STOCK_PRODUCT", "สินค้าคงเหลือรายสินค้า") Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
         DocumentType = 14
         frmInventoryDocWHPart.DocumentType = DocumentType
         frmInventoryDocWHPart.HeaderText = "สินค้าคงเหลือรายสินค้า"
         Load frmInventoryDocWHPart
         frmInventoryDocWHPart.Show 1

         Unload frmInventoryDocWHPart
         Set frmInventoryDocWHPart = Nothing
      ElseIf lMenuChosen = 2 Then
         If Not VerifyAccessRight("INVENTORY-WH_STOCK_LOCATION", "สินค้าคงเหลือรายวัน") Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
         
         frmInventoryWH.JobDocType = 1
         frmInventoryWH.HeaderText = "สินค้าคงเหลือรายวัน"
         Load frmInventoryWH
         frmInventoryWH.Show 1
         Unload frmInventoryWH
         Set frmInventoryWH = Nothing
      End If

   
   ElseIf Key = "MAIN_ENTERPRISE" Then
      If Not VerifyAccessRight("MAIN_ENTERPRISE") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      frmAddEditEnterprise.ShowMode = SHOW_ADD
      Load frmAddEditEnterprise
      frmAddEditEnterprise.Show 1
      
      Unload frmAddEditEnterprise
      Set frmAddEditEnterprise = Nothing
   ElseIf Key = "MAIN_CUSTOMER" Then
      If Not VerifyAccessRight("MAIN_CUSTOMER") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      Load frmCustomer
      frmCustomer.Show 1
      
      Unload frmCustomer
      Set frmCustomer = Nothing
   ElseIf Key = "MAIN_SUPPLIER" Then
   If Not VerifyAccessRight("MAIN_SUPPLIER") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If

      Load frmSupplier
      frmSupplier.Show 1
      
      Unload frmSupplier
      Set frmSupplier = Nothing
   ElseIf Key = "MAIN_EMPLOYEE" Then
   If Not VerifyAccessRight("MAIN_EMPLOYEE") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If

      Load frmEmployee
      frmEmployee.Show 1
      
      Unload frmEmployee
      Set frmEmployee = Nothing
   ElseIf Key = "MAIN_FREELANCE" Then
   If Not VerifyAccessRight("MAIN_FREELANCE") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If

      Load frmFreelance
      frmFreelance.Show 1
      
      Unload frmFreelance
      Set frmFreelance = Nothing
   ElseIf Key = "MAIN_REPORT" Then
   If Not VerifyAccessRight("MAIN_REPORT") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If

      frmSummaryReport.HeaderText = Caption
      frmSummaryReport.MasterMode = 3
      Load frmSummaryReport
      frmSummaryReport.Show 1
      
      Unload frmSummaryReport
       Set frmSummaryReport = Nothing
   ElseIf Key = "PRODUCT_FORMULA" Then
   If Not VerifyAccessRight("PRODUCT_FORMULA") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      Load frmFormula
      frmFormula.Show 1
      
      Unload frmFormula
      Set frmFormula = Nothing
   ElseIf Key = "PRODUCT_JOB" Then
      If Not VerifyAccessRight("PRODUCT_JOB") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      lMenuChosen = oMenu.AddMenu(m_JobProcessMenus)
      If lMenuChosen = 0 Then
         Exit Sub
      End If
      
      If lMenuChosen = 2 Then
         DocumentType = 14 'บรรจุBAG
         frmJob.HeaderText = "ใบบรรจุอาหาร BAG"
         frmJob.ProcessID = 2
         frmJob.JobDocType = 1
         frmJob.DOCUMENT_TYPE = DocumentType
         Load frmJob
         frmJob.Show 1
         
         Unload frmJob
         Set frmJob = Nothing
      
      ElseIf lMenuChosen = 4 Then
          DocumentType = 13 'บรรจุ Bulk
         frmJob.HeaderText = "ใบบรรจุอาหาร BULK"
         frmJob.ProcessID = 4
         frmJob.JobDocType = 1
         frmJob.DOCUMENT_TYPE = DocumentType
         Load frmJob
         frmJob.Show 1
         
         Unload frmJob
         Set frmJob = Nothing
         ElseIf lMenuChosen = 6 Then
         DocumentType = 17 'บรรจุ RE-BAG -> BAG
         frmJob.DOCUMENT_TYPE = DocumentType
         frmJob.HeaderText = "ใบบรรจุอาหาร RE-BAG -> BAG"
         frmJob.ProcessID = 6 'บรรจุ RE-BAG ใน job
         frmJob.JobDocType = 1
         Load frmJob
         frmJob.Show 1
         
         Unload frmJob
         Set frmJob = Nothing
       ElseIf lMenuChosen = 7 Then
         DocumentType = 18 'บรรจุ RE-BAG -> BULK
         frmJob.DOCUMENT_TYPE = DocumentType
         frmJob.HeaderText = "ใบบรรจุอาหาร RE-BAG -> BULK"
         frmJob.ProcessID = 7 'บรรจุ RE-BAG ใน job
         frmJob.JobDocType = 1
         Load frmJob
         frmJob.Show 1
         
         Unload frmJob
         Set frmJob = Nothing
       Else
         frmJob.ProcessID = lMenuChosen
         frmJob.JobDocType = 1
         Load frmJob
         frmJob.Show 1
         
         Unload frmJob
         Set frmJob = Nothing
      End If
      

      
   ElseIf Key = "PRODUCT_PACK" Then
      If Not VerifyAccessRight("PRODUCT_PACK") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.Popup("ใบสั่งแพ็คอาหาร")
      If lMenuChosen <= 0 Then
         Call EnableForm(Me, True)
         Exit Sub
      End If

'      frmPackProduction.ProcessID = lMenuChosen
'      frmPackProduction.JobDocType = 1
      Load frmPackProduction
      frmPackProduction.Show 1
      
      Unload frmPackProduction
      Set frmPackProduction = Nothing
   ElseIf Key = "PRODUCT_ESTIMATE" Then
      If Not VerifyAccessRight("PRODUCT_ESTIMATE") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.Popup("ปันค่าใช้จ่าย", "-", "รายการค่าใช้จ่าย")
      If lMenuChosen <= 0 Then
         Call EnableForm(Me, True)
         Exit Sub
      End If

      If lMenuChosen = 1 Then
         Load frmCostProduction
         frmCostProduction.Show 1
      
         Unload frmCostProduction
         Set frmCostProduction = Nothing
      ElseIf lMenuChosen = 3 Then
         Load frmExpense
         frmExpense.Show 1
      
         Unload frmExpense
         Set frmExpense = Nothing
      End If
   ElseIf Key = "PRODUCT_PLAN" Then
      If Not VerifyAccessRight("PRODUCT_PLAN") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.Popup("วางแผนการผลิตประจำวัน", "-", "ประมาณการผลิตรายเดือน")
      If lMenuChosen <= 0 Then
         Call EnableForm(Me, True)
         Exit Sub
      End If

      If lMenuChosen = 1 Then
         Load frmCostProduction
         frmCostProduction.Show 1
      
         Unload frmCostProduction
         Set frmCostProduction = Nothing
'      ElseIf lMenuChosen = 3 Then
'         Load frmExpense
'         frmExpense.Show 1
'
'         Unload frmExpense
'         Set frmExpense = Nothing
      End If
      
   ElseIf Key = "PRODUCT_REPORT" Then
   If Not VerifyAccessRight("PRODUCT_REPORT") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
     frmSummaryReport.HeaderText = Caption
      frmSummaryReport.MasterMode = 8
      Load frmSummaryReport
      frmSummaryReport.Show 1
      
      Unload frmSummaryReport
       Set frmSummaryReport = Nothing
   ElseIf Key = "LEDGER_CURRENCY" Then
      Load frmCurrency
      frmCurrency.Show 1
      
      Unload frmCurrency
      Set frmCurrency = Nothing
   ElseIf Key = "LEDGER_CURRENCYEX" Then
      Load frmCurrencyEx
      frmCurrencyEx.Show 1
      
      Unload frmCurrencyEx
      Set frmCurrencyEx = Nothing
   ElseIf Key = "LEDGER_BUY" Then
      If Not VerifyAccessRight("LEDGER_BUY") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If

      lMenuChosen = oMenu.AddMenu(glbGuiConfigs.BuyMenuItems)
      If lMenuChosen = 0 Then
         Exit Sub
      End If
      Set D = GetObject("CMenuItem", glbGuiConfigs.BuyMenuItems, Trim(str(lMenuChosen)), False)
      If D Is Nothing Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      If lMenuChosen = 17 Then
         If Not VerifyAccessRight("LEDGER_BUY" & "_" & CashDocType2Text(WAITING_CHEQUE), CashDocType2Text(WAITING_CHEQUE)) Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
         frmCashDoc.Area = WAITING_CHEQUE                   'เป็นเอกสารที่ทำเมื่อได้จ่ายเช็คให้ซัพพลายเออร์แล้ว -> เช็ครอเรียกเก็บ
         frmCashDoc.DocumentType = WAITING_CHEQUE
         frmCashDoc.HeaderText = CashDocType2Text(WAITING_CHEQUE)
         Load frmCashDoc
         frmCashDoc.Show 1
         
         Unload frmCashDoc
         Set frmCashDoc = Nothing
         Exit Sub
      ElseIf lMenuChosen = 19 Then
         If Not VerifyAccessRight("LEDGER_BUY" & "_" & CashDocType2Text(PASSED_CHEQUE), CashDocType2Text(PASSED_CHEQUE)) Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
         frmCashDoc.Area = PASSED_CHEQUE                   'เป็นเอกสารที่ทำเมื่อได้จ่ายเช็คให้ซัพพลายเออร์แล้ว -> เช็ครอเรียกเก็บ
         frmCashDoc.DocumentType = PASSED_CHEQUE
         frmCashDoc.HeaderText = CashDocType2Text(PASSED_CHEQUE)
         Load frmCashDoc
         frmCashDoc.Show 1
         
         Unload frmCashDoc
         Set frmCashDoc = Nothing
         Exit Sub
      ElseIf lMenuChosen = 23 Then
         Load frmEvaluatePay
         frmEvaluatePay.Show 1
         
         Unload frmEvaluatePay
         Set frmEvaluatePay = Nothing
         Exit Sub
      ElseIf lMenuChosen = 1 Then
         frmBillingDoc1.DocumentType = 13
      ElseIf lMenuChosen = 3 Then
         frmBillingDoc1.DocumentType = 7
      ElseIf lMenuChosen = 5 Then
         frmBillingDoc1.DocumentType = 11
      ElseIf lMenuChosen = 7 Then
         frmBillingDoc1.DocumentType = 8                                   'ใบเสร็จรับเงิน
         frmBillingDoc1.ReceiptType = 3
      ElseIf lMenuChosen = 9 Then
         frmBillingDoc1.DocumentType = 10                                  'ใบเพิ่มหนี้
      ElseIf lMenuChosen = 11 Then
         frmBillingDoc1.DocumentType = 9                                   'ใบลดหนี้
       ElseIf lMenuChosen = 13 Then
         frmBillingDoc1.DocumentType = 15
      ElseIf lMenuChosen = 21 Then
         frmBillingDoc1.DocumentType = 110
       ElseIf lMenuChosen = 25 Then
         frmBillingDocPayment.DocumentType = 111
      End If
  
      
      If lMenuChosen = 25 Then
         If Not VerifyAccessRight("LEDGER_BUY" & "_" & frmBillingDocPayment.DocumentType, D.KEYWORD) Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
         
         Load frmBillingDocPayment
         frmBillingDocPayment.Show 1
         
         Unload frmBillingDocPayment
         Set frmBillingDocPayment = Nothing
      Else
          If Not VerifyAccessRight("LEDGER_BUY" & "_" & frmBillingDoc1.DocumentType, D.KEYWORD) Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
      
         frmBillingDoc1.Area = 2
         Load frmBillingDoc1
         frmBillingDoc1.Show 1
         
         Unload frmBillingDoc1
         Set frmBillingDoc1 = Nothing
      End If
   ElseIf Key = "LEDGER_STOCK_BUY" Then
      If Not VerifyAccessRight("LEDGER_STOCKBUY") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.Popup("PO สั่งซื้อวัตถุดิบ", "-", "PO สั่งซื้อวัสดุอุปกรณ์", "-", "PO สั่งซื้อ รับเข้าจ่ายออกวัสดุอุปกรณ์", "-", "PO สั่งซื้อทั่วไป", "-", "ใบรับเข้าวัตถุดิบ", "-", "ใบรับเข้าวัสดุอุปกรณ์", "-", "ใบรับเข้าจ่ายออกวัสดุอุปกรณ์", "-", "ใบรับเข้าทั่วไป")
      If lMenuChosen <= 0 Then
         Call EnableForm(Me, True)
         Exit Sub
      Else
         If lMenuChosen = 1 Then
            DocumentType = 1000
            DocumentTypeDesc = "PO สั่งซื้อวัตถุดิบ"
         ElseIf lMenuChosen = 3 Then
            DocumentType = 1001
            DocumentTypeDesc = "PO สั่งซื้อวัสดุอุปกรณ์"
         ElseIf lMenuChosen = 5 Then
            DocumentType = 1002
            DocumentTypeDesc = "PO สั่งซื้อรับเข้าจ่ายออกวัสดุอุปกรณ์"
         ElseIf lMenuChosen = 7 Then
            DocumentType = 1003
            DocumentTypeDesc = "PO สั่งซื้อทั่วไป"
         ElseIf lMenuChosen = 9 Then
            DocumentType = 100
            DocumentTypeDesc = "ใบรับเข้าวัตถุดิบ"
         ElseIf lMenuChosen = 11 Then
            DocumentType = 101
            DocumentTypeDesc = "ใบรับเข้าวัสดุอุปกรณ์"
         ElseIf lMenuChosen = 13 Then
            DocumentType = 102
            DocumentTypeDesc = "ใบรับเข้าจ่ายออกวัสดุอุปกรณ์"
         ElseIf lMenuChosen = 15 Then
            DocumentType = 103
            DocumentTypeDesc = "ใบรับเข้าทั่วไป"
         End If
      End If
      If Not VerifyAccessRight("LEDGER_STOCKBUY" & "_" & DocumentType, DocumentTypeDesc) Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
            
      frmBillingDoc1.DocumentType = DocumentType
      frmBillingDoc1.Area = 2
      Load frmBillingDoc1
      frmBillingDoc1.Show 1
      
      Unload frmBillingDoc1
      Set frmBillingDoc1 = Nothing
      
      Set oMenu = Nothing
      
   ElseIf Key = "LEDGER_SELL" Then
      If Not VerifyAccessRight("LEDGER_SELL") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If

      lMenuChosen = oMenu.AddMenu(glbGuiConfigs.SellMenuItems)
      If lMenuChosen = 0 Then
         Exit Sub
      End If
      Set D = GetObject("CMenuItem", glbGuiConfigs.SellMenuItems, Trim(str(lMenuChosen)), False)
      If D Is Nothing Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      If (lMenuChosen = 23) Then
         frmPayment.Direction = "O"
         Load frmPayment
         frmPayment.Show 1
         
         Unload frmPayment
         Set frmPayment = Nothing
         Exit Sub
      ElseIf (lMenuChosen = 25) Then
         If Not VerifyAccessRight("LEDGER_SELL" & "_" & CashDocType2Text(CASH_DEPOSIT), CashDocType2Text(CASH_DEPOSIT)) Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
         frmCashDoc.Area = CASH_DEPOSIT                   'ใบนำฝาก เงินสด/เช็คในมือ
         frmCashDoc.DocumentType = CASH_DEPOSIT
         frmCashDoc.HeaderText = CashDocType2Text(CASH_DEPOSIT)
         Load frmCashDoc
         frmCashDoc.Show 1
         
         Unload frmCashDoc
         Set frmCashDoc = Nothing
         Exit Sub
      ElseIf (lMenuChosen = 27) Then
         If Not VerifyAccessRight("LEDGER_SELL" & "_" & CashDocType2Text(POST_CHEQUE), CashDocType2Text(POST_CHEQUE)) Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
         frmCashDoc.Area = POST_CHEQUE                   'ใบยืนยันการได้รับเงิน
         frmCashDoc.DocumentType = POST_CHEQUE
         frmCashDoc.HeaderText = CashDocType2Text(POST_CHEQUE)
         Load frmCashDoc
         frmCashDoc.Show 1
         
         Unload frmCashDoc
         Set frmCashDoc = Nothing
         Exit Sub
         
      ElseIf (lMenuChosen = 33) Then
           If Not VerifyAccessRight("INVENTORY-WH_EXPORT_2000", "ใบขึ้นอาหาร BAG") Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
         frmInventoryDocWH.DocumentType = 2000
         frmInventoryDocWH.HeaderText = "ใบขึ้นอาหาร BAG"
         Load frmInventoryDocWH
        frmInventoryDocWH.Show 1

         Unload frmInventoryDocWH
         Set frmInventoryDocWH = Nothing
   
      ElseIf (lMenuChosen = 35) Then
         If Not VerifyAccessRight("INVENTORY-WH_EXPORT_2001", "ใบขึ้นอาหาร BULK") Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
         frmInventoryDocWH.DocumentType = 2001
         frmInventoryDocWH.HeaderText = "ใบขึ้นอาหาร BULK"
         Load frmInventoryDocWH
         frmInventoryDocWH.Show 1

         Unload frmInventoryDocWH
         Set frmInventoryDocWH = Nothing
      ElseIf (lMenuChosen = 37) Then
         If Not VerifyAccessRight("INVENTORY-WH_EXPORT_2004", "ใบขึ้นอาหาร อื่นๆ") Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
         frmInventoryDocWH.DocumentType = 2004
         frmInventoryDocWH.HeaderText = "ใบขึ้นอาหาร อื่นๆ"
         Load frmInventoryDocWH
         frmInventoryDocWH.Show 1

         Unload frmInventoryDocWH
         Set frmInventoryDocWH = Nothing
       ElseIf (lMenuChosen = 31) Then
         If Not VerifyAccessRight("LEDGER_SELL" & "_" & ChequeDocType2Text(CHECK_CHEQUE), ChequeDocType2Text(CHECK_CHEQUE)) Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
         frmCheckChequeDoc.Area = CHECK_CHEQUE                'ใบตรวจสอบการรับ/คืนเช็ค
         frmCheckChequeDoc.DocumentType = CHECK_CHEQUE
         frmCheckChequeDoc.HeaderText = ChequeDocType2Text(CHECK_CHEQUE)
         Load frmCheckChequeDoc
         frmCheckChequeDoc.Show 1
         
         Unload frmCheckChequeDoc
         Set frmCheckChequeDoc = Nothing
         Exit Sub
      ElseIf lMenuChosen = 1 Then
         frmBillingDoc1.DocumentType = 14
      ElseIf lMenuChosen = 3 Then
         frmBillingDoc1.DocumentType = 12
      ElseIf lMenuChosen = 5 Then
         frmBillingDoc1.DocumentType = 1
         frmBillingDoc1.DoReceiptFlag = "Y"
      ElseIf lMenuChosen = 7 Then
         frmBillingDoc1.DocumentType = 5
      ElseIf lMenuChosen = 9 Then
         frmBillingDoc1.DoReceiptFlag = "Y"
         frmBillingDoc1.DocumentType = 2
         frmBillingDoc1.ReceiptType = 1
      ElseIf lMenuChosen = 11 Then
         frmBillingDoc1.DocumentType = 2
         frmBillingDoc1.ReceiptType = 3
      ElseIf lMenuChosen = 13 Then
         frmBillingDoc1.DocumentType = 4
      ElseIf lMenuChosen = 15 Then
         frmBillingDoc1.DocumentType = 3
      ElseIf lMenuChosen = 17 Then
         frmBillingDoc1.DocumentType = 6
      ElseIf lMenuChosen = 19 Then
         frmBillingDoc1.DocumentType = 17
      ElseIf lMenuChosen = 21 Then
         frmBillingDoc1.DocumentType = 18
      ElseIf lMenuChosen = 29 Then
         frmBillingDoc1.DocumentType = 19                   'ใบสั่งขาย SO
       ElseIf lMenuChosen = 31 Then
         frmBillingDoc1.DocumentType = 20
      End If
      If lMenuChosen <> 33 And lMenuChosen <> 35 And lMenuChosen <> 37 Then
         If Not VerifyAccessRight("LEDGER_SELL" & "_" & frmBillingDoc1.DocumentType, D.KEYWORD) Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
      Else
            Call EnableForm(Me, True)
            Exit Sub
      End If
               
      frmBillingDoc1.Area = 1
      Load frmBillingDoc1
      frmBillingDoc1.Show 1
      
      Unload frmBillingDoc1
      Set frmBillingDoc1 = Nothing
   ElseIf Key = "LEDGER_CASH" Then
      lMenuChosen = oMenu.Popup("ใบนำฝากธนาคาร")
      If lMenuChosen = 0 Then
         Exit Sub
      End If
   
      If lMenuChosen = 1 Then
         frmPayment.Direction = "O"
      ElseIf lMenuChosen = 3 Then
         frmPayment.Direction = "I"
      End If
      
      Load frmPayment
      frmPayment.Show 1
   
      Unload frmPayment
      Set frmPayment = Nothing
   ElseIf Key = "LEDGER_REPORT" Then
      If Not VerifyAccessRight("LEDGER_REPORT") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      frmSummaryReport.HeaderText = Caption
      frmSummaryReport.MasterMode = 5
      Load frmSummaryReport
      frmSummaryReport.Show 1
      
      Unload frmSummaryReport
       Set frmSummaryReport = Nothing
   ElseIf Key = "MASTER_MAIN" Then
      If Not VerifyAccessRight("MASTER_MAIN") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      frmMasterMain.HeaderText = Caption
      frmMasterMain.MasterMode = 3
      Load frmMasterMain
      frmMasterMain.Show 1
      
      Unload frmMasterMain
      Set frmMasterMain = Nothing
   ElseIf Key = "MASTER_INVENTORY" Then
      If Not VerifyAccessRight("MASTER_INVENTORY") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If

      frmMasterMain.HeaderText = Caption
      frmMasterMain.MasterMode = 1
      Load frmMasterMain
      frmMasterMain.Show 1
      
      Unload frmMasterMain
      Set frmMasterMain = Nothing
   ElseIf Key = "MASTER_HR" Then
      frmMasterMain.HeaderText = Caption
      frmMasterMain.MasterMode = 2
      Load frmMasterMain
      frmMasterMain.Show 1
      
      Unload frmMasterMain
      Set frmMasterMain = Nothing
   ElseIf Key = "MASTER_LEDGER" Then
    If Not VerifyAccessRight("MASTER_LEDGER") Then
         Call EnableForm(Me, True)
         Exit Sub
     End If

      frmMasterMain.HeaderText = Caption
      frmMasterMain.MasterMode = 7
      Load frmMasterMain
      frmMasterMain.Show 1

      Unload frmMasterMain
      Set frmMasterMain = Nothing
   ElseIf Key = "MASTER_PACKAGE" Then
   If Not VerifyAccessRight("MASTER_PACKAGE") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If

      frmMasterMain.HeaderText = Caption
      frmMasterMain.MasterMode = 6
      Load frmMasterMain
      frmMasterMain.Show 1
      
      Unload frmMasterMain
      Set frmMasterMain = Nothing
   ElseIf Key = "MASTER_PRODUCTION" Then
    If Not VerifyAccessRight("MASTER_PRODUCTION") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If

      frmMasterMain.HeaderText = Caption
      frmMasterMain.MasterMode = 8
      Load frmMasterMain
      frmMasterMain.Show 1
   
      Unload frmMasterMain
      Set frmMasterMain = Nothing
      
   ElseIf Key = "PLANNING_1" Then
      If Not VerifyAccessRight("PLANNING_1", Caption) Then
         Call EnableForm(Me, True)
         Exit Sub
      End If

      frmPlanning.HeaderText = Caption
      frmPlanning.PlanningArea = 1
      Load frmPlanning
      frmPlanning.Show 1
   
      Unload frmPlanning
      Set frmPlanning = Nothing
   ElseIf Key = "PLANNING_2" Then
      If Not VerifyAccessRight("PLANNING_2", Caption) Then
         Call EnableForm(Me, True)
         Exit Sub
      End If

      frmPlanning.HeaderText = Caption
      frmPlanning.PlanningArea = 2
      Load frmPlanning
      frmPlanning.Show 1
   
      Unload frmPlanning
      Set frmPlanning = Nothing
   ElseIf Key = "PLANNING_3" Then
      If Not VerifyAccessRight("PLANNING_3", Caption) Then
         Call EnableForm(Me, True)
         Exit Sub
      End If

      frmPlanning.HeaderText = Caption
      frmPlanning.PlanningArea = 3
      Load frmPlanning
      frmPlanning.Show 1
   
      Unload frmPlanning
      Set frmPlanning = Nothing
   ElseIf Key = "PLANNING_4" Then
      If Not VerifyAccessRight("PLANNING_4", Caption) Then
         Call EnableForm(Me, True)
         Exit Sub
      End If

      frmPlanning.HeaderText = Caption
      frmPlanning.PlanningArea = 4
      Load frmPlanning
      frmPlanning.Show 1
   
      Unload frmPlanning
      Set frmPlanning = Nothing
   ElseIf Key = "PLAN_REPORT" Then
      If Not VerifyAccessRight("PLANNING_REPORT") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      frmSummaryReport.HeaderText = Caption
      frmSummaryReport.MasterMode = 6
      Load frmSummaryReport
      frmSummaryReport.Show 1
      
      Unload frmSummaryReport
      Set frmSummaryReport = Nothing
   
   
   ElseIf Key = "COMMISSION_TARGET" Then
      If Not VerifyAccessRight("COMMISSION_TARGET") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If

      frmtarget.HeaderText = Caption
      Load frmtarget
      frmtarget.Show 1
   
      Unload frmtarget
      Set frmtarget = Nothing
   ElseIf Key = "COMMISSION_ORGANIZE" Then
      If Not VerifyAccessRight("COMMISSION_ORGANIZE") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      frmMasterValid.HeaderText = "แผนภูมิพนักงาน"
      frmMasterValid.DocumentType = COMMISSION_BUDGET_CHART
      Load frmMasterValid
      frmMasterValid.Show 1

      Unload frmMasterValid
      Set frmMasterValid = Nothing
   ElseIf Key = "COMMISSION_CONDITION" Then
      If Not VerifyAccessRight("COMMISSION_CONDITION") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   
      frmMasterValid.HeaderText = "เงื่อนไข COMMISSION"
      frmMasterValid.DocumentType = COMMISSION_CONDITION
      Load frmMasterValid
      frmMasterValid.Show 1

      Unload frmMasterValid
      Set frmMasterValid = Nothing
   ElseIf Key = "COMMISSION_COST" Then
      If Not VerifyAccessRight("COMMISSION_COST") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   
      frmMasterValid.HeaderText = "ต้นทุนรายเบอร์"
      frmMasterValid.DocumentType = COMMISSION_COST
      Load frmMasterValid
      frmMasterValid.Show 1

      Unload frmMasterValid
      Set frmMasterValid = Nothing
   ElseIf Key = "COMMISSION_SUBTRACT" Then
      If Not VerifyAccessRight("COMMISSION_SUBTRACT") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If

      frmCommissionSubtract.HeaderText = Caption
      Load frmCommissionSubtract
      frmCommissionSubtract.Show 1
   
      Unload frmCommissionSubtract
      Set frmCommissionSubtract = Nothing
   ElseIf Key = "COMMISSION_INCENTIVE" Then
      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.Popup("ข้อมูล INCENTIVE รายสินค้า", "-", "ข้อมูล INCENTIVE รายลูกค้า สินค้า", "-", "ข้อมูล COMMISSION พิเศษ", "-", "ข้อมูล INCENTIVE พิเศษ")
      If lMenuChosen <= 0 Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      If lMenuChosen = 1 Then
         If Not VerifyAccessRight("COMMISSION_INCENTIVE") Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
      
         frmCommissionIncentive.HeaderText = Caption
         frmCommissionIncentive.DocumentType = 1
         Load frmCommissionIncentive
         frmCommissionIncentive.Show 1
      
         Unload frmCommissionIncentive
         Set frmCommissionIncentive = Nothing
      ElseIf lMenuChosen = 3 Then
         If Not VerifyAccessRight("COMMISSION_INCENTIVE-CUS-PD") Then
            Call EnableForm(Me, True)
            Exit Sub
         End If

         frmCommissionIncentive.HeaderText = Caption
         frmCommissionIncentive.DocumentType = 2
         Load frmCommissionIncentive
         frmCommissionIncentive.Show 1
      
         Unload frmCommissionIncentive
         Set frmCommissionIncentive = Nothing
     ElseIf lMenuChosen = 5 Then
          If Not VerifyAccessRight("COMMISSION_INCENTIVE-COM-EXTRA") Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
         
         frmCommissionIncentive.HeaderText = Caption
         frmCommissionIncentive.DocumentType = 3
         Load frmCommissionIncentive
         frmCommissionIncentive.Show 1
      
         Unload frmCommissionIncentive
         Set frmCommissionIncentive = Nothing
      ElseIf lMenuChosen = 7 Then
          If Not VerifyAccessRight("COMMISSION_INCENTIVE-INC-EXTRA") Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
         
         frmCommissionIncentive.HeaderText = Caption
         frmCommissionIncentive.DocumentType = 4
         Load frmCommissionIncentive
         frmCommissionIncentive.Show 1
      
         Unload frmCommissionIncentive
         Set frmCommissionIncentive = Nothing
      End If
      
     
'   ElseIf Key = "COMMISSION_INCENTIVE-CUS-PD" Then
'      If Not VerifyAccessRight("COMMISSION_INCENTIVE-CUS-PD") Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If
'
'      frmCommissionIncentive.HeaderText = Caption
'      frmCommissionIncentive.DocumentType = 2
'      Load frmCommissionIncentive
'      frmCommissionIncentive.Show 1
'
'      Unload frmCommissionIncentive
'      Set frmCommissionIncentive = Nothing
   ElseIf Key = "COMMISSION_CREDIT" Then
      If Not VerifyAccessRight("COMMISSION_CREDIT") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If

      frmCommissionCredit.HeaderText = Caption
      Load frmCommissionCredit
      frmCommissionCredit.Show 1
   
      Unload frmCommissionCredit
      Set frmCommissionCredit = Nothing
   ElseIf Key = "COMMISSION_REPORT" Then
      If Not VerifyAccessRight("COMMISSION_REPORT") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      frmSummaryReport.HeaderText = Caption
      frmSummaryReport.MasterMode = 10
      Load frmSummaryReport
      frmSummaryReport.Show 1
      
      Unload frmSummaryReport
      Set frmSummaryReport = Nothing
      
       If Not VerifyAccessRight("PRODUCT_ESTIMATE") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf Key = "PACKAGE-CENTER" Then
      If Not VerifyAccessRight("PACKAGE-CENTER") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If

      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.Popup("ราคาสินค้าหน้าโรง", "-", "ราคาค่าขนส่ง", "-", "โปรโมชั่นราคาค่าสินค้า", "-", "โปรโมชั่นราคาค่าขนส่ง")
      If lMenuChosen <= 0 Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
'
      If lMenuChosen = 1 Then
        If Not VerifyAccessRight("PACKAGE-CENTER_EX-WORKS-PRICE") Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
         
         frmExWorksPrice.Area = 1
         Load frmExWorksPrice
         frmExWorksPrice.Show 1
         Unload frmExWorksPrice
         Set frmExWorksPrice = Nothing
      ElseIf lMenuChosen = 3 Then
         If Not VerifyAccessRight("PACKAGE-CENTER_DELIVERY-COST") Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
         
         frmExWorksPrice.Area = 2
         Load frmExWorksPrice
         frmExWorksPrice.Show 1
         Unload frmExWorksPrice
         Set frmExWorksPrice = Nothing
     ElseIf lMenuChosen = 5 Then
         If Not VerifyAccessRight("PACKAGE-CENTER_PROMOTION-PART") Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
         
         frmExWorksPrice.Area = 3
         Load frmExWorksPrice
         frmExWorksPrice.Show 1
         Unload frmExWorksPrice
         Set frmExWorksPrice = Nothing
     ElseIf lMenuChosen = 7 Then
         If Not VerifyAccessRight("PACKAGE-CENTER_PROMOTION-DELIVERY") Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
         
         frmExWorksPrice.Area = 4
         Load frmExWorksPrice
         frmExWorksPrice.Show 1
         Unload frmExWorksPrice
         Set frmExWorksPrice = Nothing
      End If
      
   End If
   Set oMenu = Nothing
End Sub

Private Sub cmdPasswd_Click()
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu

   Set oMenu = New cPopupMenu
   #If LIMIT_AREA <> 1 Then
       lMenuChosen = oMenu.Popup("เปลี่ยนรหัสผ่าน", "-", "ปรับราคาเฉลี่ย", "-", "อิมพอร์ตข้อมูล", "-", "คอนฟิคเลขที่เอกสาร", "-", "โน๊ต MEMO", "-", "EXPORT ข้อมูล AP", "-", "IMPORT ข้อมูล AP", "-", "IMPORT ผังบัญชี mapping", "-", "แพตข้อมูล", "-", "EXPORT ข้อมูลเช็ค", "-", "IMPORT ข้อมูลเช็ค", "-", "ประมวลผลประจำปี", "-", "ระบบปรับราคาย้อนหลัง", "-", "ระบบแจ้งเตือน", "-", "ตั้งค่าสิทธิ์อนุมัติการสั่งซื้อ", "-", "กำหนดวันที่เอกสาร", "-", "IMPORT ข้อมูลวัตถุดิบที่ต้องการ", "-", "ข้อมูลการชั่งน้ำหนัก")
      'lMenuChosen = oMenu.Popup("เปลี่ยนรหัสผ่าน", "-", "ปรับราคาเฉลี่ย", "-", "อิมพอร์ตข้อมูล", "-", "คอนฟิคเลขที่เอกสาร", "-", "โน๊ต MEMO", "-", "EXPORT ข้อมูล AP", "-", "IMPORT ข้อมูล AP", "-", "IMPORT ผังบัญชี mapping", "-", "แพตข้อมูล", "-", "EXPORT ข้อมูลเช็ค", "-", "IMPORT ข้อมูลเช็ค", "-", "ประมวลผลประจำปี", "-", "ระบบปรับราคาย้อนหลัง", "-", "ระบบแจ้งเตือน")
   #End If
   
   If lMenuChosen = 0 Then
      Exit Sub
   End If

   If lMenuChosen = 1 Then
      If Not VerifyAccessRight("PROGRAM_" & "เปลี่ยนรหัสผ่าน", "เปลี่ยนรหัสผ่าน") Then
         Call EnableForm(Me, True)
         Exit Sub
     End If
      Load frmChangePassword
      frmChangePassword.Show 1
      
      Unload frmChangePassword
      Set frmChangePassword = Nothing
   ElseIf lMenuChosen = 3 Then
     If Not VerifyAccessRight("PROGRAM_" & "ปรับราคาเฉลี่ย", "ปรับราคาเฉลี่ย") Then
         Call EnableForm(Me, True)
         Exit Sub
     End If
      Load frmReArrangeDoc
      frmReArrangeDoc.Show 1
      
      Unload frmReArrangeDoc
      Set frmReArrangeDoc = Nothing
   ElseIf lMenuChosen = 5 Then
      If Not VerifyAccessRight("PROGRAM_" & "อิมพอร์ตข้อมูล", "อิมพอร์ตข้อมูล") Then
         Call EnableForm(Me, True)
         Exit Sub
     End If
     'Exit Sub
      Load frmImportDoc
      frmImportDoc.Show 1
      
      Unload frmImportDoc
      Set frmImportDoc = Nothing
   ElseIf lMenuChosen = 7 Then
      If Not VerifyAccessRight("PROGRAM_" & "คอนฟิคเลขที่เอกสาร", "คอนฟิคเลขที่เอกสาร") Then
         Call EnableForm(Me, True)
         Exit Sub
     End If
      frmConfigDoc.HeaderText = "คอนฟิคเลขที่เอกสาร"
      Load frmConfigDoc
      frmConfigDoc.Show 1
      
      Unload frmConfigDoc
      Set frmConfigDoc = Nothing
   ElseIf lMenuChosen = 9 Then
      If Not VerifyAccessRight("PROGRAM_" & "โน๊ต MEMO", "โน๊ต MEMO") Then
         Call EnableForm(Me, True)
         Exit Sub
     End If
      Load frmMemoNote
      frmMemoNote.Show 1
      
      Unload frmMemoNote
      Set frmMemoNote = Nothing
   ElseIf lMenuChosen = 11 Then
      If Not VerifyAccessRight("PROGRAM_" & "EXPORT ข้อมูล AP", "EXPORT ข้อมูล AP") Then
         Call EnableForm(Me, True)
         Exit Sub
     End If
      Load frmExportSupItem
      frmExportSupItem.Show 1
      
      Unload frmExportSupItem
      Set frmExportSupItem = Nothing
   ElseIf lMenuChosen = 13 Then
      If Not VerifyAccessRight("PROGRAM_" & "IMPORT ข้อมูล AP", "IMPORT ข้อมูล AP") Then
         Call EnableForm(Me, True)
         Exit Sub
     End If
      Load frmImportSupItem
      frmImportSupItem.Show 1
      
      Unload frmImportSupItem
      Set frmImportSupItem = Nothing
   ElseIf lMenuChosen = 15 Then
      If Not VerifyAccessRight("PROGRAM_" & "IMPORT ผังบัญชี mapping", "IMPORT ผังบัญชี mapping") Then
         Call EnableForm(Me, True)
         Exit Sub
     End If
      Exit Sub
      Load frmImportDoc2
      frmImportDoc2.Show 1
      
      Unload frmImportDoc2
      Set frmImportDoc2 = Nothing
   ElseIf lMenuChosen = 17 Then
      If Not VerifyAccessRight("PROGRAM_" & "แพตข้อมูล", "แพตข้อมูล") Then
         Call EnableForm(Me, True)
         Exit Sub
     End If
     Exit Sub
      Load frmImportDoc3
      frmImportDoc3.Show 1
      
      Unload frmImportDoc3
      Set frmImportDoc3 = Nothing
   ElseIf lMenuChosen = 19 Then
      If Not VerifyAccessRight("PROGRAM_" & "EXPORT ข้อมูลเช็ค", "EXPORT ข้อมูลเช็ค") Then
         Call EnableForm(Me, True)
         Exit Sub
     End If
      Load frmExportPostItem
      frmExportPostItem.Show 1
      
      Unload frmExportPostItem
      Set frmExportPostItem = Nothing
   ElseIf lMenuChosen = 21 Then
      If Not VerifyAccessRight("PROGRAM_" & "IMPORT ข้อมูลเช็ค", "IMPORT ข้อมูลเช็ค") Then
         Call EnableForm(Me, True)
         Exit Sub
     End If
      Load frmImportPostItem
      frmImportPostItem.Show 1
      
      Unload frmImportPostItem
      Set frmImportPostItem = Nothing
   ElseIf lMenuChosen = 23 Then
      If Not VerifyAccessRight("PROGRAM_" & "ประมวลผลประจำปี", "ประมวลผลประจำปี") Then
         Call EnableForm(Me, True)
         Exit Sub
     End If
     
      Load frmInitBalance
      frmInitBalance.Show 1
      
      Unload frmInitBalance
      Set frmInitBalance = Nothing
   ElseIf lMenuChosen = 25 Then
      If Not VerifyAccessRight("PROGRAM_" & "ระบบปรับราคาย้อนหลัง", "ระบบปรับราคาย้อนหลัง") Then
         Call EnableForm(Me, True)
         Exit Sub
     End If
     
      Load frmAdjustSellPrice
      frmAdjustSellPrice.Show 1
      
      Unload frmAdjustSellPrice
      Set frmAdjustSellPrice = Nothing
   ElseIf lMenuChosen = 27 Then
      If Not VerifyAccessRight("PROGRAM_" & "ระบบแจ้งเตือน", "ระบบแจ้งเตือน") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      Load frmAlertBox
      frmAlertBox.Show 1
      
      Unload frmAlertBox
      Set frmAlertBox = Nothing
   ElseIf lMenuChosen = 29 Then
      If Not VerifyAccessRight("PROGRAM_APPROVE-PO", "ระบบตั้งค่าสิทธิ์อนุมัติใบสั่งซื้อ") Then
         Call EnableForm(Me, True)
         Exit Sub
     End If
     
     frmAuthenPO.HeaderText = "สิทธิ์การอนุมัติใบสั่งซื้อ"
     frmAuthenPO.ShowMode = SHOW_VIEW_ONLY
      Load frmAuthenPO
      frmAuthenPO.Show 1

      Unload frmAuthenPO
      Set frmAuthenPO = Nothing
    ElseIf lMenuChosen = 31 Then
        If Not VerifyAccessRight("PROGRAM_LOCK-DATE", "กำหนดวันที่เอกสาร") Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
      frmLockDate.HeaderText = "กำหนดวันที่เอกสาร"
      Load frmLockDate
      frmLockDate.Show 1
      
      Unload frmLockDate
      Set frmLockDate = Nothing
   ElseIf lMenuChosen = 33 Then
      If Not VerifyAccessRight("PROGRAM_" & "IMPORT รหัสวัตถุดิบที่ต้องการ", "IMPORT รหัสวัตถุดิบที่ต้องการ") Then
         Call EnableForm(Me, True)
         Exit Sub
     End If

      Load frmImportDoc4
      frmImportDoc4.Show 1
      
      Unload frmImportDoc4
      Set frmImportDoc4 = Nothing
   ElseIf lMenuChosen = 35 Then
        If Not VerifyAccessRight("PROGRAM_WEIGHT-PREVIEW", "ข้อมูลการชั่งน้ำหนัก") Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
      frmWeightPreview.HeaderText = "รายงานข้อมูลการชั่งน้ำหนัก"
      Load frmWeightPreview
      frmWeightPreview.Show 1
      
      Unload frmWeightPreview
      Set frmWeightPreview = Nothing
   End If
End Sub

Private Sub Form_Activate()
Dim OKClick As Boolean
Dim iCount As Long
Dim TempDB As String
Dim m_AlertBox As CAlertBox
Dim m_BillingDoc As CBillingDoc
Dim massageAlert As String
Dim documentTypeString As String
   
   If Not m_HasActivate Then
   
      If Command = "1" Or Command = "" Then
         TempDB = glbParameterObj.DBFile
      ElseIf Command = "2" Then
         TempDB = glbParameterObj.DBFileAP
      Else
         TempDB = glbParameterObj.DBFileAPX
      End If
      Me.Caption = Me.Caption & "  " & TempDB
   
      m_HasActivate = True
      Call PatchDB
       Call InitFormLayout
      Load frmLogin
      frmLogin.Show 1
      
      OKClick = frmLogin.OKClick
      
      Unload frmLogin
      Set frmLogin = Nothing
      glbEnterPrise.ENTERPRISE_ID = -1
      Call glbEnterPrise.QueryData(m_Rs, iCount)
      If Not m_Rs.EOF Then
         Call glbEnterPrise.PopulateFromRS(1, m_Rs)
         Call InitNormalLabel(lblCompany, MapText(glbEnterPrise.ENTERPRISE_NAME & "  " & glbEnterPrise.BRANCH_NAME))
      End If
       
      If Not (CheckTask) Then
         trvMain.Refresh
      End If
         
      If Not OKClick Then
         m_MustAsk = False
         Unload Me
         Exit Sub
      Else
         Call LoadAccessRight(Nothing, glbAccessRight, glbUser.GROUP_ID)
         Call GeneratePartGroupMenu(m_PartGroupMenus)
         Call GenerateJobProcessMenu(m_JobProcessMenus)
      End If
      
      Me.Caption = glbUser.USER_NAME & " " & glbGuiConfigs.ShowWindowCaption(glbParameterObj.Programowner)
   
      IP_ADDRESS = glbDatabaseMngr.m_Winsock.LocalIP  'getIPAddress()
      If ((glbUser.LOGON_STATUS = 1) And (IP_ADDRESS <> glbUser.IP_ADDRESS) And (Not (glbUser.USER_NAME = "ADMIN"))) Then
               glbErrorLog.LocalErrorMsg = "มีการเข้าใช้งานด้วยผู้ใช้งานนี้แล้วในเครื่องอื่นๆ หรือ ท่านลืม LOGOUT ออกจากระบบ"
               glbErrorLog.ShowUserError
               
               m_MustAsk = False
               Unload Me
               Exit Sub
      End If
      
      Call glbDaily.UpdateLogonStatus(1, IP_ADDRESS)
      
      'เช็คแจ้งเตือน
      Set m_AlertBox = New CAlertBox
      m_AlertBox.USER_NAME = glbUser.USER_NAME
      m_AlertBox.ALERT_DATE_SEARCH = Now
      m_AlertBox.ALERT_CANCEL_FLAG = "N"
      m_AlertBox.ALERT_BOX_TYPE = 1
      Call m_AlertBox.QueryData(1, m_Rs, iCount)
      While Not m_Rs.EOF
         Call m_AlertBox.PopulateFromRS(1, m_Rs)
         glbErrorLog.LocalErrorMsg = "จากวันที่ : " & DateToStringExtEx2(m_AlertBox.ALERT_BOX_FROM) & " ถึงวันที่ : " & DateToStringExtEx2(m_AlertBox.ALERT_BOX_TO) & " : " & m_AlertBox.ALERT_BOX_DESC
         glbErrorLog.ShowUserError
         m_Rs.MoveNext
      Wend
      
      Set m_AlertBox = New CAlertBox
      m_AlertBox.USER_NAME = glbUser.USER_NAME
      m_AlertBox.ALERT_DATE_SEARCH = Now
      m_AlertBox.ALERT_CANCEL_FLAG = "N"
      m_AlertBox.ALERT_BOX_TYPE = 2
      Call m_AlertBox.QueryData(1, m_Rs, iCount)
      If Not m_Rs.EOF Then
         Call m_AlertBox.PopulateFromRS(1, m_Rs)
         
        Set m_BillingDoc = New CBillingDoc
        m_BillingDoc.AUTO_GEN_FLAG = "Y"
        m_BillingDoc.GEN_COMMIT_FLAG = "N"
        Call m_BillingDoc.QueryData(109, m_Rs, iCount)
        
        massageAlert = m_AlertBox.ALERT_BOX_DESC & vbCrLf
        
       If Not m_Rs.EOF Then
         While Not m_Rs.EOF
             Call m_BillingDoc.PopulateFromRS(109, m_Rs)
             massageAlert = massageAlert & "เลขที่เอกสาร : " & m_BillingDoc.DOCUMENT_NO & "   วันที่ : " & DateToStringExtEx2(m_BillingDoc.DOCUMENT_DATE) & " ประเภท : " & PoRoTypeToString(m_BillingDoc.DOCUMENT_TYPE)
             massageAlert = massageAlert & vbCrLf
             m_Rs.MoveNext
         Wend
          glbErrorLog.LocalErrorMsg = massageAlert
          glbErrorLog.ShowUserError
         End If
      End If
      Call getLockDate
      Call LoadLoginTracking(Nothing, m_LoginTracking)
   End If
End Sub

Private Sub Form_Load()
   m_MustAsk = True
'   Call InitFormLayout
   Set m_PartGroupMenus = New Collection
   Set m_JobProcessMenus = New Collection
   Set m_Rs = New ADODB.Recordset
   Set m_LoginTracking = New Collection
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

   If m_MustAsk Then
      glbErrorLog.LocalErrorMsg = MapText("ท่านต้องการออกจากโปรแกรมใช่หรือไม่")
      If glbErrorLog.AskMessage = vbYes Then
         Cancel = False
         Call glbDaily.UpdateLogonStatus(2, IP_ADDRESS)
      Else
         Cancel = True
      End If
   End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
   
   SSPanel1.Width = ScaleWidth
   pnlHeader.Width = ScaleWidth - SSFrame1.Width
   lblDateTime.Left = ScaleWidth - lblDateTime.Width - 50
   
   SSFrame1.Top = SSPanel1.HEIGHT
   pnlHeader.Top = SSPanel1.HEIGHT
   SSFrame1.HEIGHT = ScaleHeight - SSPanel1.HEIGHT
   
   cmdExit.Top = SSFrame1.HEIGHT - cmdExit.HEIGHT - 50
   cmdPasswd.Top = cmdExit.Top
   
   lblUsername.Top = SSFrame1.HEIGHT - 2900
   lblUserGroup.Top = SSFrame1.HEIGHT - 2300
   lblVersion.Top = SSFrame1.HEIGHT - 1700
   lblLastVersion.Top = SSFrame1.HEIGHT - 1100
   lblLastVersion2.Top = SSFrame1.HEIGHT - 1100
   
   trvMain.HEIGHT = SSFrame1.HEIGHT - 4500
   
'   lblUsername.Top = SSFrame1.HEIGHT - 2200
'   lblUserGroup.Top = SSFrame1.HEIGHT - 1600
'   lblVersion.Top = SSFrame1.HEIGHT - 1000
'   lblLastVersion.Top = SSFrame1.HEIGHT - 400
'    trvMain.HEIGHT = SSFrame1.HEIGHT - 3500
   
   lblCompany.Width = ScaleWidth
         
   fraGeneric.Width = pnlHeader.Width * 4 / 5
   fraGeneric.Left = pnlHeader.Left + ((pnlHeader.Width - fraGeneric.Width) / 2)
   
   cmdGeneric(0).Width = fraGeneric.Width * 9 / 10
   cmdGeneric(0).Left = (fraGeneric.Width - cmdGeneric(0).Width) / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_PartGroupMenus = Nothing
   Set m_JobProcessMenus = Nothing
   Set m_Rs = Nothing
   Call ReleaseAll
   Set m_LoginTracking = Nothing
End Sub

Private Sub SSCommand2_Click()
Dim IsOK As Boolean
   Call glbDaily.PatchBankAccount(IsOK, True, glbErrorLog)
End Sub

Private Sub Timer1_Timer()
   Timer1.Enabled = False
   lblDateTime.Caption = "                                                    "
   lblDateTime.Caption = DateToStringExtEx3(Now)
   lblUsername.Caption = MapText("ผู้ใช้ : ") & " " & glbUser.USER_NAME
   lblUserGroup.Caption = MapText("กลุ่มผู้ใช้ : ") & " " & glbUser.GROUP_NAME
   
   Timer1.Enabled = True
End Sub
Private Sub trvMain_NodeClick(ByVal Node As MSComctlLib.Node)
   If Node Is Nothing Then
      Exit Sub
   End If
   pnlHeader.Caption = Node.Text
   If Node.Key = ROOT_TREE & " 1-0" Then
      Call InitCommandLayout(glbGuiConfigs.AdminCommandMenuItems)
   ElseIf Node.Key = ROOT_TREE & " 1-1" Then
      Call InitCommandLayout(glbGuiConfigs.MasterCommandMenuItems)
   ElseIf Node.Key = ROOT_TREE & " 1-2" Then
      Call InitCommandLayout(glbGuiConfigs.MainCommandMenuItems)
   ElseIf Node.Key = ROOT_TREE & " 1-3" Then
      Call InitCommandLayout(glbGuiConfigs.StockCommandMenuItems)
   ElseIf Node.Key = ROOT_TREE & " 1-4" Then
      Call InitCommandLayout(glbGuiConfigs.PlanCommandMenuItems)
   ElseIf Node.Key = ROOT_TREE & " 1-5" Then
      Call InitCommandLayout(glbGuiConfigs.LedgerCommandMenuItems)
   ElseIf Node.Key = ROOT_TREE & " 1-6" Then
      
   ElseIf Node.Key = ROOT_TREE & " 1-7" Then
      Call InitCommandLayout(glbGuiConfigs.PackageCommandMenuItems)
   ElseIf Node.Key = ROOT_TREE & " 1-9" Then
      Call InitCommandLayout(glbGuiConfigs.ProdCommandMenuItems)
   ElseIf Node.Key = ROOT_TREE & " 1-10" Then
      Call InitCommandLayout(glbGuiConfigs.CommissionCommandMenuItems)
   ElseIf Node.Key = ROOT_TREE & " 1-11" Then
      Call InitCommandLayout(glbGuiConfigs.StockWHCommandMenuItems)
   End If
End Sub

Private Sub InitCommandLayout(Col As Collection)
Dim D As CMenuItem
Dim Top As Long
Dim Left As Long
Dim I As Long
Dim hight As Long
   Top = cmdGeneric(0).Top
   Left = cmdGeneric(0).Left
   fraGeneric.HEIGHT = 1450
   hight = fraGeneric.HEIGHT
   For I = 1 To (cmdGeneric.Count - 1)
      cmdGeneric(I).Visible = False
      Unload cmdGeneric(I)
      fraGeneric.Visible = False
   Next I
   
   I = 0
   For Each D In Col
      I = I + 1
      
      Load cmdGeneric(I)
      cmdGeneric(I).Visible = False
      cmdGeneric(I).Picture = LoadPicture(glbParameterObj.MainButton)
      cmdGeneric(I).PictureAlignment = ssLeftMiddle
      cmdGeneric(I).Left = Left
      cmdGeneric(I).Top = Top
      cmdGeneric(I).Tag = D.KEYWORD
      Call InitMainButton(cmdGeneric(I), D.MENU_TEXT)
      cmdGeneric(I).Visible = True
      fraGeneric.HEIGHT = hight
      fraGeneric.Visible = True
      hight = hight + cmdGeneric(0).HEIGHT + 10
      Top = Top + cmdGeneric(0).HEIGHT + 10
   Next D
     
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.NAME
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 123 Then
      Call AddMemoNote
      KeyCode = 0
   End If
End Sub

