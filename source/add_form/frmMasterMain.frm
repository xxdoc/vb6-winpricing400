VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmMasterMain 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmMasterMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   15690
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel pnlFooter 
         Height          =   705
         Left            =   30
         TabIndex        =   2
         Top             =   7800
         Width           =   11850
         _ExtentX        =   20902
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin Threed.SSCommand cmdOK 
            Height          =   525
            Left            =   8445
            TabIndex        =   8
            Top             =   120
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   926
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmMasterMain.frx":27A2
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdExit 
            Cancel          =   -1  'True
            Height          =   525
            Left            =   10095
            TabIndex        =   7
            Top             =   120
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   926
            _Version        =   131073
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdEdit 
            Height          =   525
            Left            =   1770
            TabIndex        =   6
            Top             =   120
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   926
            _Version        =   131073
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdAdd 
            Height          =   525
            Left            =   150
            TabIndex        =   5
            Top             =   120
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   926
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmMasterMain.frx":2ABC
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdDelete 
            Height          =   525
            Left            =   3420
            TabIndex        =   4
            Top             =   120
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   926
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmMasterMain.frx":2DD6
            ButtonStyle     =   3
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   855
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1508
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   0
            Top             =   30
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   2
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMasterMain.frx":30F0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMasterMain.frx":39CC
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ImageList ImageList2 
            Left            =   2850
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   32
            ImageHeight     =   32
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   1
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMasterMain.frx":3CE8
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
      Begin MSComctlLib.TreeView trvMaster 
         Height          =   6945
         Left            =   0
         TabIndex        =   3
         Top             =   870
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   12250
         _Version        =   393217
         Indentation     =   882
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "JasmineUPC"
            Size            =   15.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   6915
         Left            =   4500
         TabIndex        =   9
         Top             =   900
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   12197
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
         Column(1)       =   "frmMasterMain.frx":4002
         Column(2)       =   "frmMasterMain.frx":40CA
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmMasterMain.frx":416E
         FormatStyle(2)  =   "frmMasterMain.frx":42CA
         FormatStyle(3)  =   "frmMasterMain.frx":437A
         FormatStyle(4)  =   "frmMasterMain.frx":442E
         FormatStyle(5)  =   "frmMasterMain.frx":4506
         ImageCount      =   0
         PrinterProperties=   "frmMasterMain.frx":45BE
      End
   End
End
Attribute VB_Name = "frmMasterMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Rs As ADODB.Recordset
Private m_HasActivate As Boolean
Private m_TableName As String
Private m_PartType As CPartType
Private m_Location As CLocation
Private m_ProductType As CProductType
Private m_ProductStatus As CProductStatus
Private m_House As CHouse
Private m_Country As CCountry
Private m_CustomerType As CCustomerType
Private m_CustomerGrade As CCustomerGrade
Private m_SupplierType As CSupplierType
Private m_SupplierGrade As CSupplierGrade
Private m_SupplierStatus As CSupplierStatus
Private m_Position As CEmpPosition
Private m_Unit As CUnit
Private m_PartGroup As CPartGroup
Private m_HouseGroup As CHouseGroup
Private m_StatusGroup As CStatusGroup
Private m_AgeRange As CAgeRange
Private m_FormulaType As CFormulaType
Private m_Reason As CReason
Private m_Layout As CLayout
Private m_SellType As CSellType
Private m_DoType As CDoType
Private m_FeatureType As CFeatureType
Private m_Resource As CResource
Private m_Process As CProcess
Private m_Machine As CMachine
Private m_Money_family As CMoneyFamily
Private m_ParameterProcess As CParameterProcess
Private m_Bank As CBank
Private m_BankBranch As CBankBranch
Private m_Packaging As CPackaging
Private m_PurchaseExpense As CPurchaseExpense
Private m_MasterRef As CMasterRef

Private m_Sp As CSystemParam
Private m_Work As CWorkStatus
Private m_Religious As CReligious
Private m_Resign As CResignReason
Private m_BankAccount As CBankAccount
Private m_DocumentType As CDocumentType
Private m_MonthlyAdd As CMonthlyAdd
Private m_MonthlySub As CMonthlySub

Public HeaderText As String
Public MasterMode As Long

Private Sub cmdAdd_Click()
Dim OKClick As Boolean

   If trvMaster.SelectedItem Is Nothing Then
      Exit Sub
   End If
   
   If trvMaster.SelectedItem.Key = "" Then
      Exit Sub
   End If
   
   If trvMaster.SelectedItem.Key = ROOT_TREE Then
      glbErrorLog.LocalErrorMsg = ""
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-1" Then
      frmAddEditParameter.ShowMode = SHOW_ADD
      frmAddEditParameter.HeaderText = MapText("ข้อมูลหลัก") & " " & trvMaster.SelectedItem.Text
      Load frmAddEditParameter
      frmAddEditParameter.Show 1
      
      OKClick = frmAddEditParameter.OKClick
      
      Unload frmAddEditParameter
      Set frmAddEditParameter = Nothing
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 1-4" Then
      frmAddEditHouseGroup.ShowMode = SHOW_ADD
      frmAddEditHouseGroup.HeaderText = MapText("ข้อมูลหลัก") & " " & trvMaster.SelectedItem.Text
      Load frmAddEditHouseGroup
      frmAddEditHouseGroup.Show 1
      
      OKClick = frmAddEditHouseGroup.OKClick
      
      Unload frmAddEditHouseGroup
      Set frmAddEditHouseGroup = Nothing
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 7-6" Then
      frmAddEditMaster2.MasterMode = MasterMode
      frmAddEditMaster2.MasterKey = trvMaster.SelectedItem.Key
      frmAddEditMaster2.ShowMode = SHOW_ADD
      frmAddEditMaster2.HeaderText = MapText("ข้อมูลหลัก") & " " & trvMaster.SelectedItem.Text
      Load frmAddEditMaster2
      frmAddEditMaster2.Show 1
      
      OKClick = frmAddEditMaster2.OKClick
      
      Unload frmAddEditMaster2
      Set frmAddEditMaster2 = Nothing
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 1-11" Then
      frmAddEditPrtItemSet.ShowMode = SHOW_ADD
      frmAddEditPrtItemSet.HeaderText = MapText("ข้อมูลหลัก") & " " & trvMaster.SelectedItem.Text
      Load frmAddEditPrtItemSet
      frmAddEditPrtItemSet.Show 1
      
      OKClick = frmAddEditPrtItemSet.OKClick
      
      Unload frmAddEditPrtItemSet
      Set frmAddEditPrtItemSet = Nothing
   Else
      frmAddEditMaster1.MasterMode = MasterMode
      frmAddEditMaster1.MasterKey = trvMaster.SelectedItem.Key
      frmAddEditMaster1.ShowMode = SHOW_ADD
      frmAddEditMaster1.HeaderText = MapText("ข้อมูลหลัก") & " " & trvMaster.SelectedItem.Text
      Load frmAddEditMaster1
      frmAddEditMaster1.Show 1
      
      OKClick = frmAddEditMaster1.OKClick
      
      Unload frmAddEditMaster1
      Set frmAddEditMaster1 = Nothing
   End If
   
   If OKClick Then
      Call trvMaster_NodeClick(trvMaster.SelectedItem)
   End If
End Sub


Private Sub InitTreeView()
Dim Node As Node

   trvMaster.Font.NAME = GLB_FONT
   trvMaster.Font.Size = 14
   
   If MasterMode = 1 Then
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-13", MapText("SET สินค้า"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-4", MapText("กลุ่มวัตถุดิบ"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-1", MapText("ประเภทวัตถุดิบ"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-2", MapText("สถานที่จัดเก็บ"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-3", MapText("หน่วยวัด"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-5", MapText("สาเหตุการเบิก"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-6", MapText("สาเหตุการปรับยอด"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-7", MapText("หน่วยงาน/แผนก"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-8", MapText("ภาชนะบรรจุ"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-9", MapText("ค่าใช้จ่ายจัดซื้อ"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-10", MapText("ค่าใช้จ่ายการเบิก"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-11", MapText("เซตของอาหาร/วัตถุดิบ"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-12", MapText("รายละเอียดการเบิก"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-14", MapText("รายละเอียดโครงการ"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-15", MapText("ประเภทที่จัดเก็บ"), 1, 2)
      Node.Expanded = False
   

      
   ElseIf MasterMode = 2 Then
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True

      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 2-1", MapText("สถานะการทำงาน"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 2-2", MapText("ศาสนา"), 1, 2)
      Node.Expanded = False

      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 2-3", MapText("สาเหตุที่ออก"), 1, 2)
      Node.Expanded = False

      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 2-4", MapText("ธนาคาร"), 1, 2)
      Node.Expanded = False

      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 2-5", MapText("ประเภทบัตร"), 1, 2)
      Node.Expanded = False

      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 2-6", MapText("ส่วนบวกเงินเดือน"), 1, 2)
      Node.Expanded = False

      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 2-7", MapText("ส่วนหักเงินเดือน"), 1, 2) '
    Node.Expanded = False


   ElseIf MasterMode = 3 Then
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-1", MapText("ประเทศ"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-2", MapText("ระดับลูกค้า"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-3", MapText("ประเภทลูกค้า"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-4", MapText("ระดับซัพพลายเออร์"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-5", MapText("ประเภทซัพพลายเออร์"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-6", MapText("สถานะซัพพลายเออร์"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-7", MapText("ตำแหน่ง"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-8", MapText("ประเภท MEMO"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-9", MapText("สถานะ MEMO"), 1, 2)
      Node.Expanded = False
   ElseIf MasterMode = 4 Then
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-1", MapText("ประเภทราคาทอง"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-2", MapText("ประเภทบิลซื้อขาย"), 1, 2)
      Node.Expanded = False
      
   ElseIf MasterMode = 5 Then
   ElseIf MasterMode = 6 Then
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-1", MapText("หน่วยวัด"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-3", MapText("กลุ่มสินค้า/บริการ"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-2", MapText("ประเภทสินค้า/บริการ"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-4", MapText("ชนิดสินค้าก่อนบรรจุ"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-5", MapText("ประเภทลูกค้า"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-6", MapText("ชนิดสัตว์"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-7", MapText("รายการค่าขนส่ง"), 1, 2)
      Node.Expanded = False
      
'      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-8", MapText("รายการส่งเสริมการขาย"), 1, 2)
'      Node.Expanded = False
   ElseIf MasterMode = 7 Then
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 7-1", MapText("ทรัพยากร"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 7-2", MapText("สกุลเงิน"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 7-3", MapText("ธนาคาร"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 7-4", MapText("สาขาธนาคาร"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 7-5", MapText("สาเหตุการเพิ่ม/ลดหนี้"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 7-6", MapText("เลขที่บัญชีธนาคาร"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 7-7", MapText("ประเภทเช็ค"), 1, 2) '
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 7-8", MapText("รายการบัญชี"), 1, 2) '
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 7-9", MapText("จ่ายให้กับ"), 1, 2) '
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 7-10", MapText("เงื่อนไขหลังรับวัตถุดิบ"), 1, 2) '
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 7-11", MapText("การชำระในPO"), 1, 2) '
      Node.Expanded = False
      
   ElseIf MasterMode = 8 Then
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-0", MapText("ค่าใช้จ่ายผลิต"), 1, 2)
      Node.Expanded = False
      
      ' Comment เพราะว่าไม่ต้องการให้สร้าง แก้ไข ลบ เพราะว่าจะต้องเอา รหัส process ไปฝังกับเลขที่เอกสารใน inventory_doc
      'มีการ hard code ในโปรแกรมเกิดขึ้น
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-1", MapText("โปรเซส"), 1, 2)
      Node.Expanded = False

      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-2", MapText("ประเภทสูตร"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-3", MapText("เครื่องจักร"), 1, 2)
      Node.Expanded = False
      
'      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-4", MapText("ประเภทอาหารสัตว์"), 1, 2)
'      Node.Expanded = False
   End If
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long
   
   Call EnableForm(Me, True)
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrorHandler
Dim Status As Boolean
Dim IsOK As Boolean
Dim TempID As Long

   If trvMaster.SelectedItem.Key = "" Then
      Exit Sub
   End If
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   TempID = GridEX1.Value(1)
   
   If Not ConfirmDelete(GridEX1.Value(3)) Then
      Exit Sub
   End If
   
   If trvMaster.SelectedItem.Key = ROOT_TREE & " 1-1" Then
      Status = glbMaster.DeletePartType(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 1-2" Then
      Status = glbMaster.DeleteLocation(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 1-3" Then
      Status = glbMaster.DeleteUnit(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 1-4" Then
      Status = glbMaster.DeletePartGroup(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 1-5" Then
      Status = glbMaster.DeleteReason(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 1-6" Then
      Status = glbMaster.DeleteReason(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 1-7" Then
      Status = glbMaster.DeleteLayout(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 1-8" Then
      Status = glbMaster.DeletePackaging(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 1-9" Then
      Status = glbMaster.DeletePurchaseExpense(TempID, IsOK, glbErrorLog)
  ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 1-10" Then
      Status = glbMaster.DeleteMasterRef(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 1-11" Then
      Status = glbMaster.DeleteMasterRef(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 2-1" Then
      Status = glbMaster.DeleteWorkStatus(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 2-2" Then
      Status = glbMaster.DeleteReligious(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 2-3" Then
      Status = glbMaster.DeleteResign(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 2-4" Then
      Status = glbMaster.DeleteBankAccount(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 2-5" Then
      Status = glbMaster.DeleteDocumentType(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 2-6" Then
      Status = glbMaster.DeleteMonthlyAdd(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 2-7" Then
      If TempID = 1 Then
      glbErrorLog.LocalErrorMsg = MapText("ไม่สามารถลบข้อมูลเงินยืมได้")
      glbErrorLog.ShowUserError
      Exit Sub
    End If
      Status = glbMaster.DeleteMonthlySub(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 3-1" Then
      Status = glbMaster.DeleteCountry(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 3-2" Then
      Status = glbMaster.DeleteCustomerGrade(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 3-3" Then
      Status = glbMaster.DeleteCustomerType(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 3-4" Then
      Status = glbMaster.DeleteSupplierGrade(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 3-5" Then
      Status = glbMaster.DeleteSupplierType(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 3-6" Then
      Status = glbMaster.DeleteSupplierStatus(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 3-7" Then
      Status = glbMaster.DeletePosition(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 4-1" Then
      Status = glbMaster.DeleteSellType(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 4-2" Then
      Status = glbMaster.DeleteDoType(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 6-1" Then
      Status = glbMaster.DeleteUnit(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 6-2" Then
      Status = glbMaster.DeleteFeatureType(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 7-1" Then
      Status = glbMaster.DeleteResource(TempID, IsOK, glbErrorLog)
  ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 7-2" Then
      Status = glbMaster.DeleteMoneyFamily(TempID, IsOK, glbErrorLog)
  ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 7-3" Then
      Status = glbMaster.DeleteBank(TempID, IsOK, glbErrorLog)
  ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 7-4" Then
      Status = glbMaster.DeleteBankBranch(TempID, IsOK, glbErrorLog)
  ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 7-5" Then
      Status = glbMaster.DeleteMasterRef(TempID, IsOK, glbErrorLog)
  ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 7-6" Then
      Status = glbMaster.DeleteMasterRef(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 7-7" Then
      Status = glbMaster.DeleteMasterRef(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 8-0" Then
      Status = glbMaster.DeleteParameterProcess(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 8-1" Then
      Status = glbMaster.DeleteProcess(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 8-2" Then
      Status = glbMaster.DeleteFormulaType(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 8-3" Then
      Status = glbMaster.DeleteMachine(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 8-4" Then
      Status = glbMaster.DeleteMachine(TempID, IsOK, glbErrorLog)
   Else
      Status = glbMaster.DeleteMasterRef(TempID, IsOK, glbErrorLog)
   End If

   If Status Then
      Call trvMaster_NodeClick(trvMaster.SelectedItem)
   Else
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   Exit Sub
   
ErrorHandler:
End Sub

Private Sub cmdEdit_Click()
Dim OKClick As Boolean
Dim TempID As Long

   If trvMaster.SelectedItem.Key = "" Then
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   TempID = GridEX1.Value(1)
   
   If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-1" Then
      frmAddEditParameter.id = TempID
      frmAddEditParameter.ShowMode = SHOW_EDIT
      frmAddEditParameter.HeaderText = MapText("ข้อมูลหลัก") & " " & trvMaster.SelectedItem.Text
      Load frmAddEditParameter
      frmAddEditParameter.Show 1
      
      OKClick = frmAddEditParameter.OKClick
      
      Unload frmAddEditParameter
      Set frmAddEditParameter = Nothing
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 1-4" Then
      frmAddEditHouseGroup.id = TempID
      frmAddEditHouseGroup.ShowMode = SHOW_EDIT
      frmAddEditHouseGroup.HeaderText = MapText("ข้อมูลหลัก") & " " & trvMaster.SelectedItem.Text
      Load frmAddEditHouseGroup
      frmAddEditHouseGroup.Show 1
      
      OKClick = frmAddEditHouseGroup.OKClick
      
      Unload frmAddEditHouseGroup
      Set frmAddEditHouseGroup = Nothing
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 7-6" Then
      frmAddEditMaster2.MasterMode = MasterMode
      frmAddEditMaster2.MasterKey = trvMaster.SelectedItem.Key
      frmAddEditMaster2.id = TempID
      frmAddEditMaster2.ShowMode = SHOW_EDIT
      frmAddEditMaster2.HeaderText = MapText("ข้อมูลหลัก") & " " & trvMaster.SelectedItem.Text
      Load frmAddEditMaster2
      frmAddEditMaster2.Show 1
      
      OKClick = frmAddEditMaster2.OKClick
      
      Unload frmAddEditMaster2
      Set frmAddEditMaster2 = Nothing
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 1-11" Then
      frmAddEditPrtItemSet.ShowMode = SHOW_EDIT
      frmAddEditPrtItemSet.HeaderText = MapText("ข้อมูลหลัก") & " " & trvMaster.SelectedItem.Text
      frmAddEditPrtItemSet.id = TempID
      Load frmAddEditPrtItemSet
      frmAddEditPrtItemSet.Show 1
      
      OKClick = frmAddEditPrtItemSet.OKClick
      
      Unload frmAddEditPrtItemSet
      Set frmAddEditPrtItemSet = Nothing
   Else
      frmAddEditMaster1.MasterMode = MasterMode
      frmAddEditMaster1.id = TempID
      frmAddEditMaster1.MasterKey = trvMaster.SelectedItem.Key
      frmAddEditMaster1.ShowMode = SHOW_EDIT
      frmAddEditMaster1.HeaderText = MapText("ข้อมูลหลัก") & " " & trvMaster.SelectedItem.Text
      Load frmAddEditMaster1
      frmAddEditMaster1.Show 1
      
      OKClick = frmAddEditMaster1.OKClick
      
      Unload frmAddEditMaster1
      Set frmAddEditMaster1 = Nothing
   End If
   
   If OKClick Then
      Call trvMaster_NodeClick(trvMaster.SelectedItem)
   End If
End Sub
Private Sub Form_Activate()
Dim ItemCount As Long

   If Not m_HasActivate Then
      Me.Refresh
      DoEvents
      
      Call QueryData(True)
      m_HasActivate = True
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
      Call cmdAdd_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
'      Call cmdOK_Click
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

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_PartType = Nothing
   Set m_Location = Nothing
   Set m_ProductType = Nothing
   Set m_ProductStatus = Nothing
   Set m_House = Nothing
   Set m_Country = Nothing
   Set m_CustomerGrade = Nothing
   Set m_CustomerType = Nothing
   Set m_SupplierGrade = Nothing
   Set m_SupplierType = Nothing
   Set m_SupplierStatus = Nothing
   Set m_Position = Nothing
   Set m_Unit = Nothing
   Set m_PartGroup = Nothing
   Set m_HouseGroup = Nothing
   Set m_StatusGroup = Nothing
   Set m_AgeRange = Nothing
   Set m_FormulaType = Nothing
   Set m_Reason = Nothing
   Set m_Layout = Nothing
   Set m_SellType = Nothing
   Set m_DoType = Nothing
   Set m_FeatureType = Nothing
   Set m_Resource = Nothing
   Set m_Work = Nothing
   Set m_Religious = Nothing
   Set m_Resign = Nothing
   Set m_BankAccount = Nothing
   Set m_DocumentType = Nothing
   Set m_MonthlyAdd = Nothing
   Set m_MonthlySub = Nothing
   Set m_Process = Nothing
   Set m_Machine = Nothing
   Set m_Bank = Nothing
   Set m_BankBranch = Nothing
   Set m_Packaging = Nothing
   Set m_PurchaseExpense = Nothing
   Set m_MasterRef = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid0()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   
   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid1()
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
   Col.Width = 1110
   Col.Caption = MapText("รหัสวัตถุดิบ")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 6225
   Col.Caption = MapText("วัตถุดิบ")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid2()
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
   Col.Width = 1620
   Col.Caption = MapText("รหัสสถานที่จัดเก็บ")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2200
   Col.Caption = MapText("สถานที่จัดเก็บ")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 1520
   Col.Caption = MapText("ประเภทที่จัดเก็บ")
   
   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid1_3()
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
   Col.Width = 1620
   Col.Caption = MapText("รหัสหน่วยวัด")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5715
   Col.Caption = MapText("หน่วยวัด")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid1_4()
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
   Col.Width = 1620
   Col.Caption = MapText("รหัสกลุ่มวัตถุดิบ")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5715
   Col.Caption = MapText("กลุ่มวัตถุดิบ")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid1_5()
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
   Col.Width = 1620
   Col.Caption = MapText("รหัสสาเหตุการเบิก")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5715
   Col.Caption = MapText("สาเหตุการเบิก")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid1_6()
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
   Col.Width = 1620
   Col.Caption = MapText("รหัสสาเหตุการปรับยอด")

   Set Col = GridEX1.Columns.add '3
   Col.Width = 5715
   Col.Caption = MapText("สาเหตุการปรับยอด")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid1_14()
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
   Col.Width = 1620
   Col.Caption = MapText("รหัสโครงการ")

   Set Col = GridEX1.Columns.add '3
   Col.Width = 5715
   Col.Caption = MapText("ชื่อโครงการ")

   GridEX1.ItemCount = 0
End Sub
Private Sub InitGrid1_15()
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
   Col.Width = 1820
   Col.Caption = MapText("รหัสประเภทที่จัดเก็บ")

   Set Col = GridEX1.Columns.add '3
   Col.Width = 5515
   Col.Caption = MapText("ชื่อประเภทที่จัดเก็บ")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid1_8()
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
   Col.Width = 1620
   Col.Caption = MapText("รหัสภาชนะ")

   Set Col = GridEX1.Columns.add '3
   Col.Width = 5715
   Col.Caption = MapText("ภาชนะบรรจุ")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid1_9()
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
   Col.Width = 1620
   Col.Caption = MapText("รหัสค่าใช้จ่าย")

   Set Col = GridEX1.Columns.add '3
   Col.Width = 5715
   Col.Caption = MapText("ค่าใช้จ่ายจัดซื้อ")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid1_10()
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
   Col.Width = 1620
   Col.Caption = MapText("รหัสค่าใช้จ่าย")

   Set Col = GridEX1.Columns.add '3
   Col.Width = 5715
   Col.Caption = MapText("ค่าใช้จ่ายการเบิก")

   GridEX1.ItemCount = 0
End Sub
Private Sub InitGrid1_11()
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
   Col.Width = 1620
   Col.Caption = MapText("รหัสเซตข้อมูล")

   Set Col = GridEX1.Columns.add '3
   Col.Width = 5715
   Col.Caption = MapText("เซตข้อมูล")

   GridEX1.ItemCount = 0
End Sub
Private Sub InitGrid1_12()
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
   Col.Width = 1620
   Col.Caption = MapText("รหัสรายละเอียดการเบิก")

   Set Col = GridEX1.Columns.add '3
   Col.Width = 5715
   Col.Caption = MapText("รายละเอียดการเบิก")

   GridEX1.ItemCount = 0
End Sub
Private Sub InitGrid1_7()
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
   Col.Width = 1620
   Col.Caption = MapText("รหัสหน่วงาน/แผนก")

   Set Col = GridEX1.Columns.add '3
   Col.Width = 5715
   Col.Caption = MapText("หน่วยงาน/แผนก")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid3_1()
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
   Col.Width = 1620
   Col.Caption = MapText("รหัสประเทศ")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5715
   Col.Caption = MapText("ประเทศ")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid3_2()
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
   Col.Width = 1620
   Col.Caption = MapText("รหัสระดับลูกค้า")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5715
   Col.Caption = MapText("ระดับลูกค้า")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid3_3()
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
   Col.Width = 1620
   Col.Caption = MapText("รหัสประเภทลูกค้า")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5715
   Col.Caption = MapText("ประเภทลูกค้า")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid3_4()
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
   Col.Width = 2220
   Col.Caption = MapText("รหัสระดับซัพพลายเออร์")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5115
   Col.Caption = MapText("ระดับซับพลายเออร์")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid3_5()
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
   Col.Width = 2220
   Col.Caption = MapText("รหัสประเภทซัพพลายเออร์")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5115
   Col.Caption = MapText("ประเภทซับพลายเออร์")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid3_6()
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
   Col.Width = 2220
   Col.Caption = MapText("รหัสสถานะซัพพลายเออร์")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5115
   Col.Caption = MapText("สถานะซับพลายเออร์")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid3_7()
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
   Col.Width = 2220
   Col.Caption = MapText("รหัสตำแหน่ง")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5115
   Col.Caption = MapText("ตำแหน่ง")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid4_1()
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
   Col.Width = 2220
   Col.Caption = MapText("รหัสประเภทราคาทอง")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5115
   Col.Caption = MapText("ชื่อประเภทราคาทอง")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid4_2()
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
   Col.Width = 2220
   Col.Caption = MapText("รหัสประเภทบิล")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5115
   Col.Caption = MapText("ชื่อประเภทบิล")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid6_2()
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
   Col.Width = 2220
   Col.Caption = MapText("รหัสสินค้า/บริการ")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5115
   Col.Caption = MapText("ชื่อสินค้า/บริการ")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid7_1()
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
   Col.Width = 2220
   Col.Caption = MapText("หมายเลขทรัพยากร")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5115
   Col.Caption = MapText("รายละเอียด")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid7_2()
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
   Col.Width = 2220
  Col.Caption = MapText("หมายเลขสกุลเงิน")
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5115
   Col.Caption = MapText("ชื่อสกุลเงิน")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid7_3()
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
   Col.Width = 2220
  Col.Caption = MapText("รหัสธนาคาร")
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5115
   Col.Caption = MapText("ธนาคาร")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid7_4()
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
   Col.Width = 2220
  Col.Caption = MapText("รหัสสาขาธนาคาร")
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5115
   Col.Caption = MapText("สาขาธนาคาร")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid7_5()
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
   Col.Width = 2220
  Col.Caption = MapText("รหัสสาเหตุ")
  
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5115
   Col.Caption = MapText("สาเหตุการเพิ่ม/ลดหนี้")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid7_6()
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
   Col.Width = 2220
  Col.Caption = MapText("รหัสบัญชี")
  
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5115
   Col.Caption = MapText("เลขที่บัญชีธนาคาร")

   GridEX1.ItemCount = 0
End Sub
Private Sub InitGrid7_7()
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
   Col.Width = 2220
  Col.Caption = MapText("รหัสประเภทเช็ค")
  
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5115
   Col.Caption = MapText("ประเภทเช็ค")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid7_8()
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
   Col.Width = 2220
  Col.Caption = MapText("รหัสรายการบัญชี")
  
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5115
   Col.Caption = MapText("รายการบัญชี")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid7_9()
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
   Col.Width = 2220
  Col.Caption = MapText("รหัส")
  
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5115
   Col.Caption = MapText("ชื่อ")

   GridEX1.ItemCount = 0
End Sub
Private Sub InitGrid7_10()
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
   Col.Width = 2220
  Col.Caption = MapText("รหัส")
  
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5115
   Col.Caption = MapText("เงื่อนไขหลังรับของ")

   GridEX1.ItemCount = 0
End Sub
Private Sub InitGrid7_11()
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
   Col.Width = 2220
  Col.Caption = MapText("รหัส")
  
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5115
   Col.Caption = MapText("การชำระในPO")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid8_0()
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
   Col.Width = 2220
  Col.Caption = MapText("รหัสช่าใช้จ่าย")
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5115
   Col.Caption = MapText("ค่าใช้จ่ายผลิต")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid8_1()
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
   Col.Width = 2220
   
   Col.Caption = MapText("หมายเลขโปรเซส")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5115
   Col.Caption = MapText("ชื่อโปรเซส")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid8_2()
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
   Col.Width = 2220
   
   Col.Caption = MapText("รหัสประเภทสูตร")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5115
   Col.Caption = MapText("ประเภทสูตร")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid8_3()
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
   Col.Width = 2220
   Col.Caption = MapText("รหัสเครื่องจักร")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5115
   Col.Caption = MapText("ชื่อเครื่องจักร")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid2_1()
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
   Col.Width = 1620
   Col.Caption = MapText("รหัส")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5715
   Col.Caption = MapText("สถานะการทำงาน")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid2_2()
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
   Col.Width = 1620
   Col.Caption = MapText("รหัส")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5715
   Col.Caption = MapText("ศาสนา")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid2_3()
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
   Col.Width = 1620
   Col.Caption = MapText("รหัส")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5715
   Col.Caption = MapText("สาเหตุที่ออก")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid2_4()
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
   Col.Width = 1620
   Col.Caption = MapText("รหัส")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5715
   Col.Caption = MapText("ชื่อธนาคาร")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid2_5()
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
   Col.Width = 1620
   Col.Caption = MapText("รหัส")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5715
   Col.Caption = MapText("ประเภทบัตร")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid2_6()
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
   Col.Width = 1620
   Col.Caption = MapText("รหัส")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5715
   Col.Caption = MapText("ส่วนบวกเงินเดือน")

   GridEX1.ItemCount = 0
End Sub
Private Sub InitGrid2_7()
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
   Col.Width = 1620
   Col.Caption = MapText("รหัส")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5715
   Col.Caption = MapText("ส่วนหักเงินเดือน")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
   Me.KeyPreview = True
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   
   Me.BackColor = GLB_FORM_COLOR
   SSFrame1.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_HEAD_COLOR
   pnlFooter.BackColor = GLB_HEAD_COLOR
   Call InitHeaderFooter(pnlHeader, pnlFooter)

   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdExit, MapText("ออก (ESC)"))
   
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   pnlFooter.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitTreeView
   Call InitGrid0
   
'   lsvMaster.Font.NAME = GLB_FONT
'   lsvMaster.Font.Size = 14
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Call InitFormLayout
   
   m_HasActivate = False
   m_TableName = "SYSTEM_PARAM"
   Set m_Rs = New ADODB.Recordset
   
   Set m_PartType = New CPartType
   Set m_Location = New CLocation
   Set m_ProductType = New CProductType
   Set m_ProductStatus = New CProductStatus
   Set m_House = New CHouse
   Set m_Country = New CCountry
   Set m_CustomerGrade = New CCustomerGrade
   Set m_CustomerType = New CCustomerType
   Set m_SupplierGrade = New CSupplierGrade
   Set m_SupplierType = New CSupplierType
   Set m_SupplierStatus = New CSupplierStatus
   Set m_Position = New CEmpPosition
   Set m_Unit = New CUnit
   Set m_PartGroup = New CPartGroup
   Set m_HouseGroup = New CHouseGroup
   Set m_StatusGroup = New CStatusGroup
   Set m_AgeRange = New CAgeRange
   Set m_FormulaType = New CFormulaType
   Set m_Reason = New CReason
   Set m_Layout = New CLayout
   Set m_SellType = New CSellType
   Set m_DoType = New CDoType
   Set m_FeatureType = New CFeatureType
   Set m_Resource = New CResource
   Set m_Work = New CWorkStatus
   Set m_Religious = New CReligious
   Set m_Resign = New CResignReason
   Set m_BankAccount = New CBankAccount
   Set m_DocumentType = New CDocumentType
   Set m_MonthlyAdd = New CMonthlyAdd
   Set m_MonthlySub = New CMonthlySub
   Set m_Process = New CProcess
   Set m_Machine = New CMachine
   Set m_Money_family = New CMoneyFamily
   Set m_ParameterProcess = New CParameterProcess
   Set m_Bank = New CBank
   Set m_BankBranch = New CBankBranch
   Set m_Packaging = New CPackaging
   Set m_PurchaseExpense = New CPurchaseExpense
   Set m_MasterRef = New CMasterRef
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
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
   
   If MasterMode = 1 Then
      If trvMaster.SelectedItem.Key = "Root 1-1" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_PartType.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_PartType.PART_TYPE_ID
         Values(2) = m_PartType.PART_TYPE_NO
         Values(3) = m_PartType.PART_TYPE_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 1-2" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_Location.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_Location.LOCATION_ID
         Values(2) = m_Location.LOCATION_NO
         Values(3) = m_Location.LOCATION_NAME
         Values(4) = m_Location.LOCATION_GROUP_NAME
'         Values(5) = m_Location.MAX_AMOUNT
'         Values(6) = m_Location.UNIT_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 1-3" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_Unit.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_Unit.UNIT_ID
         Values(2) = m_Unit.UNIT_NO
         Values(3) = m_Unit.UNIT_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 1-4" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_PartGroup.PopulateFromRS(1, m_Rs)
      
         Values(1) = m_PartGroup.PART_GROUP_ID
         Values(2) = m_PartGroup.PART_GROUP_NO
         Values(3) = m_PartGroup.PART_GROUP_NAME
      
      ElseIf trvMaster.SelectedItem.Key = "Root 1-5" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_Reason.PopulateFromRS(1, m_Rs)
      
         Values(1) = m_Reason.REASON_ID
         Values(2) = m_Reason.REASON_NO
         Values(3) = m_Reason.REASON_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 1-6" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_Reason.PopulateFromRS(1, m_Rs)
      
         Values(1) = m_Reason.REASON_ID
         Values(2) = m_Reason.REASON_NO
         Values(3) = m_Reason.REASON_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 1-7" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_Layout.PopulateFromRS(1, m_Rs)
      
         Values(1) = m_Layout.LAY_OUT_ID
         Values(2) = m_Layout.LAY_OUT_NO
         Values(3) = m_Layout.LAY_OUT_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 1-8" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_Packaging.PopulateFromRS(1, m_Rs)
      
         Values(1) = m_Packaging.PACKAGING_ID
         Values(2) = m_Packaging.PACKAGING_NO
         Values(3) = m_Packaging.PACKAGING_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 1-9" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_PurchaseExpense.PopulateFromRS(1, m_Rs)
      
         Values(1) = m_PurchaseExpense.PUREXP_ID
         Values(2) = m_PurchaseExpense.PUREXP_NO
         Values(3) = m_PurchaseExpense.PUREXP_NAME
     ElseIf trvMaster.SelectedItem.Key = "Root 1-10" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_MasterRef.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_MasterRef.KEY_ID
         Values(2) = m_MasterRef.KEY_CODE
         Values(3) = m_MasterRef.KEY_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 1-11" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_MasterRef.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_MasterRef.KEY_ID
         Values(2) = m_MasterRef.KEY_CODE
         Values(3) = m_MasterRef.KEY_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 1-12" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_MasterRef.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_MasterRef.KEY_ID
         Values(2) = m_MasterRef.KEY_CODE
         Values(3) = m_MasterRef.KEY_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 1-13" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_MasterRef.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_MasterRef.KEY_ID
         Values(2) = m_MasterRef.KEY_CODE
         Values(3) = m_MasterRef.KEY_NAME
         
       ElseIf trvMaster.SelectedItem.Key = "Root 1-14" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_MasterRef.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_MasterRef.KEY_ID
         Values(2) = m_MasterRef.KEY_CODE
         Values(3) = m_MasterRef.KEY_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 1-15" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_MasterRef.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_MasterRef.KEY_ID
         Values(2) = m_MasterRef.KEY_CODE
         Values(3) = m_MasterRef.KEY_NAME
      End If
   ElseIf MasterMode = 2 Then
      If trvMaster.SelectedItem.Key = "Root 2-1" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_Work.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_Work.WORK_ID
         Values(2) = m_Work.WORK_NO
         Values(3) = m_Work.WORK_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 2-2" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_Religious.PopulateFromRS(m_Rs)
         
         Values(1) = m_Religious.RELIGIOUS_ID
         Values(2) = m_Religious.RELIGIOUS_NO
         Values(3) = m_Religious.RELIGIOUS_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 2-3" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_Resign.PopulateFromRS(m_Rs)
         
         Values(1) = m_Resign.RSGRESON_ID
         Values(2) = m_Resign.RSGRESON_NO
         Values(3) = m_Resign.RSGRESON_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 2-4" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_BankAccount.PopulateFromRS(m_Rs)
         
         Values(1) = m_BankAccount.BANK_ID
         Values(2) = m_BankAccount.BANK_NO
         Values(3) = m_BankAccount.BANK_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 2-5" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_DocumentType.PopulateFromRS(m_Rs)
         
         Values(1) = m_DocumentType.DOCTYPE_ID
         Values(2) = m_DocumentType.DOCTYPE_NO
         Values(3) = m_DocumentType.DOCTYPE_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 2-6" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_MonthlyAdd.PopulateFromRS(m_Rs)
         
         Values(1) = m_MonthlyAdd.MONTHLY_ADD_ID
         Values(2) = m_MonthlyAdd.MONTHLY_ADD_NO
         Values(3) = m_MonthlyAdd.MONTHLY_ADD_NAME
            ElseIf trvMaster.SelectedItem.Key = "Root 2-7" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_MonthlySub.PopulateFromRS(m_Rs)
         
         Values(1) = m_MonthlySub.MONTHLY_SUB_ID
         Values(2) = m_MonthlySub.MONTHLY_SUB_NO
         Values(3) = m_MonthlySub.MONTHLY_SUB_NAME
    End If
   ElseIf MasterMode = 3 Then
      If trvMaster.SelectedItem.Key = "Root 3-1" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_Country.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_Country.COUNTRY_ID
         Values(2) = m_Country.COUNTRY_NO
         Values(3) = m_Country.COUNTRY_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 3-2" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_CustomerGrade.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_CustomerGrade.CSTGRADE_ID
         Values(2) = m_CustomerGrade.CSTGRADE_NO
         Values(3) = m_CustomerGrade.CSTGRADE_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 3-3" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_CustomerType.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_CustomerType.CSTTYPE_ID
         Values(2) = m_CustomerType.CSTTYPE_NO
         Values(3) = m_CustomerType.CSTTYPE_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 3-4" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_SupplierGrade.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_SupplierGrade.SUPPLIER_GRADE_ID
         Values(2) = m_SupplierGrade.SUPPLIER_GRADE_NO
         Values(3) = m_SupplierGrade.SUPPLIER_GRADE_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 3-5" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_SupplierType.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_SupplierType.SUPPLIER_TYPE_ID
         Values(2) = m_SupplierType.SUPPLIER_TYPE_NO
         Values(3) = m_SupplierType.SUPPLIER_TYPE_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 3-6" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_SupplierStatus.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_SupplierStatus.SUPPLIER_STATUS_ID
         Values(2) = m_SupplierStatus.SUPPLIER_STATUS_NO
         Values(3) = m_SupplierStatus.SUPPLIER_STATUS_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 3-7" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_Position.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_Position.POSITION_ID
         Values(2) = m_Position.POSITION_NAME
         Values(3) = m_Position.POSITION_DESC
      ElseIf trvMaster.SelectedItem.Key = "Root 3-8" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_MasterRef.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_MasterRef.KEY_ID
         Values(2) = m_MasterRef.KEY_CODE
         Values(3) = m_MasterRef.KEY_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 3-9" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_MasterRef.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_MasterRef.KEY_ID
         Values(2) = m_MasterRef.KEY_CODE
         Values(3) = m_MasterRef.KEY_NAME
      End If
   ElseIf MasterMode = 4 Then
      If trvMaster.SelectedItem.Key = "Root 4-1" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_SellType.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_SellType.SELL_TYPE_ID
         Values(2) = m_SellType.SELL_TYPE_NO
         Values(3) = m_SellType.SELL_TYPE_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 4-2" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_DoType.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_DoType.DO_TYPE_ID
         Values(2) = m_DoType.DO_TYPE_NO
         Values(3) = m_DoType.DO_TYPE_NAME
      End If
   ElseIf MasterMode = 6 Then
      If trvMaster.SelectedItem.Key = "Root 6-1" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_Unit.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_Unit.UNIT_ID
         Values(2) = m_Unit.UNIT_NO
         Values(3) = m_Unit.UNIT_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 6-2" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_FeatureType.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_FeatureType.FEATURE_TYPE_ID
         Values(2) = m_FeatureType.FEATURE_TYPE_NO
         Values(3) = m_FeatureType.FEATURE_TYPE_NAME
      Else
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_MasterRef.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_MasterRef.KEY_ID
         Values(2) = m_MasterRef.KEY_CODE
         Values(3) = m_MasterRef.KEY_NAME
      End If
   ElseIf MasterMode = 7 Then
      If trvMaster.SelectedItem.Key = "Root 7-1" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_Resource.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_Resource.RESOURCE_ID
         Values(2) = m_Resource.RESOURCE_NO
         Values(3) = m_Resource.RESOURCE_NAME
     ElseIf trvMaster.SelectedItem.Key = "Root 7-2" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_Money_family.PopulateFromRS(m_Rs)
         
         Values(1) = m_Money_family.MONEY_FAMILY_ID
         Values(2) = m_Money_family.MONEY_FAMILY_NO
         Values(3) = m_Money_family.MONEY_FAMILY_NAME
     ElseIf trvMaster.SelectedItem.Key = "Root 7-3" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_Bank.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_Bank.BANK_ID
         Values(2) = m_Bank.BANK_NO
         Values(3) = m_Bank.BANK_NAME
     ElseIf trvMaster.SelectedItem.Key = "Root 7-4" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_BankBranch.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_BankBranch.BBRANCH_ID
         Values(2) = m_BankBranch.BBRANCH_NO
         Values(3) = m_BankBranch.BBRANCH_NAME
     ElseIf trvMaster.SelectedItem.Key = "Root 7-5" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_MasterRef.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_MasterRef.KEY_ID
         Values(2) = m_MasterRef.KEY_CODE
         Values(3) = m_MasterRef.KEY_NAME
     ElseIf trvMaster.SelectedItem.Key = "Root 7-7" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_MasterRef.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_MasterRef.KEY_ID
         Values(2) = m_MasterRef.KEY_CODE
         Values(3) = m_MasterRef.KEY_NAME
     
     ElseIf trvMaster.SelectedItem.Key = "Root 7-6" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_MasterRef.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_MasterRef.KEY_ID
         Values(2) = m_MasterRef.KEY_CODE
         Values(3) = m_MasterRef.KEY_NAME
      Else
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_MasterRef.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_MasterRef.KEY_ID
         Values(2) = m_MasterRef.KEY_CODE
         Values(3) = m_MasterRef.KEY_NAME
      End If
   ElseIf MasterMode = 8 Then
      If trvMaster.SelectedItem.Key = "Root 8-0" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_ParameterProcess.PopulateFromRS(m_Rs)
         
         Values(1) = m_ParameterProcess.PARAMETER_PROCESS_ID
         Values(2) = m_ParameterProcess.PARAMETER_PROCESS_NO
         Values(3) = m_ParameterProcess.PARAMETER_PROCESS_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 8-1" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_Process.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_Process.PROCESS_ID
         Values(2) = m_Process.PROCESS_NO
         Values(3) = m_Process.PROCESS_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 8-2" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_FormulaType.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_FormulaType.FORMULA_TYPE_ID
         Values(2) = m_FormulaType.FORMULA_TYPE_NO
         Values(3) = m_FormulaType.FORMULA_TYPE_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 8-3" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_Machine.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_Machine.MACHINE_ID
         Values(2) = m_Machine.MACHINE_NO
         Values(3) = m_Machine.MACHINE_NAME
      End If
   End If
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

'Private Sub LoadListView(Rs As ADODB.Recordset, FieldName As String, IDName As String)
'Dim Lst As ListItem
'
'   While Not Rs.EOF
'      Set Lst = lsvMaster.ListItems.Add(, , NVLS(Rs(FieldName), ""), 1, 1)
'      Lst.Tag = NVLI(Rs(IDName), 0)
'      Rs.MoveNext
'   Wend
'End Sub

Private Sub trvMaster_NodeClick(ByVal Node As MSComctlLib.Node)
Static LastKey As String
Dim Status As Boolean
Dim ItemCount As Long
Dim QueryFlag As Boolean

   Set m_Sp = GetSystemParam(glbSystemParams, "PROGRAM_OWNER")
   
   If LastKey = Node.Key Then
      Exit Sub
   End If

   Status = True
   QueryFlag = False
   
   If Node.Key = ROOT_TREE & " 1-1" Then
      Call InitGrid1
      Dim a1_1 As CPartType

      Set a1_1 = New CPartType
      a1_1.PART_TYPE_ID = -1
      Status = a1_1.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a1_1 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 1-2" Then
      Call InitGrid2
      Dim a1_2 As CLocation

      Set a1_2 = New CLocation
      a1_2.LOCATION_ID = -1
      a1_2.LOCATION_TYPE = 2
      Status = a1_2.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a1_2 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 1-3" Then
      Call InitGrid1_3
      Dim a1_3 As CUnit

      Set a1_3 = New CUnit
      a1_3.UNIT_ID = -1
      Status = a1_3.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a1_3 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 1-4" Then
      Call InitGrid1_4
      Dim a1_4 As CPartGroup

      Set a1_4 = New CPartGroup
      a1_4.PART_GROUP_ID = -1
      Status = a1_4.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a1_4 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 1-5" Then
      Call InitGrid1_5
      Dim a1_5 As CReason

      Set a1_5 = New CReason
      a1_5.REASON_ID = -1
      a1_5.Area = 1
      Status = a1_5.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a1_5 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 1-6" Then
      Call InitGrid1_6
      Dim a1_6 As CReason

      Set a1_6 = New CReason
      a1_6.REASON_ID = -1
      a1_6.Area = 2
      Status = a1_6.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a1_6 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 1-7" Then
      Call InitGrid1_7
      Dim a1_7 As CLayout

      Set a1_7 = New CLayout
      a1_7.LAY_OUT_ID = -1
      a1_7.LOCATION_ID = -1
      Status = a1_7.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a1_7 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 1-8" Then
      Call InitGrid1_8
      Dim a1_8 As CPackaging

      Set a1_8 = New CPackaging
      a1_8.PACKAGING_ID = -1
      Status = a1_8.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a1_8 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 1-9" Then
      Call InitGrid1_9
      Dim a1_9 As CPurchaseExpense

      Set a1_9 = New CPurchaseExpense
      a1_9.PUREXP_ID = -1
      Status = a1_9.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a1_9 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 1-10" Then
      Call InitGrid1_10
      Dim a1_10 As CMasterRef

      Set a1_10 = New CMasterRef
      a1_10.KEY_ID = -1
      a1_10.MASTER_AREA = EXPENSE_TYPE
      Status = a1_10.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind

      Set a1_10 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 1-11" Then
      Call InitGrid1_11
      Dim a1_11 As CMasterRef

      Set a1_11 = New CMasterRef
      a1_11.KEY_ID = -1
      a1_11.MASTER_AREA = PRTITEM_SET
      Status = a1_11.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind

      Set a1_11 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 1-12" Then
      Call InitGrid1_11
      Dim a1_12 As CMasterRef

      Set a1_12 = New CMasterRef
      a1_12.KEY_ID = -1
      a1_12.MASTER_AREA = EXPORT_DESC
      Status = a1_12.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind

      Set a1_12 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 1-13" Then
      Call InitGridMaster
      Dim a1_13 As CMasterRef
      
      Set a1_13 = New CMasterRef
      a1_13.KEY_ID = -1
      a1_13.MASTER_AREA = SET_PRODUCT
      Status = a1_13.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      
      
      Set a1_13 = Nothing

   ElseIf Node.Key = ROOT_TREE & " 1-14" Then
      Call InitGrid1_14
     Dim a1_14 As CMasterRef
     Set a1_14 = New CMasterRef
      a1_14.KEY_ID = -1
      a1_14.MASTER_AREA = SET_PROJECT
      Status = a1_14.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind

      Set a1_14 = Nothing
      
   ElseIf Node.Key = ROOT_TREE & " 1-15" Then
      Call InitGrid1_15
     Dim a1_15 As CMasterRef
     Set a1_15 = New CMasterRef
      a1_15.KEY_ID = -1
      a1_15.MASTER_AREA = LOCATION_GROUP
      Status = a1_15.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind

      Set a1_15 = Nothing
      
      
   ElseIf Node.Key = ROOT_TREE & " 2-1" Then
      Call InitGrid2_1
      Dim a2_1 As CWorkStatus

      Set a2_1 = New CWorkStatus
      a2_1.WORK_ID = -1
      Status = a2_1.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a2_1 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 2-2" Then
      Call InitGrid2_2
      Dim a2_2 As CReligious

      Set a2_2 = New CReligious
      a2_2.RELIGIOUS_ID = -1
      Status = a2_2.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a2_2 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 2-3" Then
      Call InitGrid2_3
      Dim a2_3 As CResignReason

      Set a2_3 = New CResignReason
      a2_3.RSGRESON_ID = -1
      Status = a2_3.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a2_3 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 2-4" Then
      Call InitGrid2_4
      Dim a2_4 As CBankAccount

      Set a2_4 = New CBankAccount
      a2_4.BANK_ID = -1
      Status = a2_4.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a2_4 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 2-5" Then
      Call InitGrid2_5
      Dim a2_5 As CDocumentType

      Set a2_5 = New CDocumentType
      a2_5.DOCTYPE_ID = -1
      Status = a2_5.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a2_5 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 2-6" Then
      Call InitGrid2_6
      Dim a2_6 As CMonthlyAdd

      Set a2_6 = New CMonthlyAdd
      a2_6.MONTHLY_ADD_ID = -1
      Status = a2_6.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a2_6 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 2-7" Then
      Call InitGrid2_7
      Dim a2_7 As CMonthlySub

      Set a2_7 = New CMonthlySub
      a2_7.MONTHLY_SUB_ID = -1
      Status = a2_7.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a2_7 = Nothing
      ElseIf Node.Key = ROOT_TREE & " 3-1" Then
      Call InitGrid3_1
      Dim a3_1 As CCountry

      Set a3_1 = New CCountry
      a3_1.COUNTRY_ID = -1
      Status = a3_1.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a3_1 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 3-2" Then
      Call InitGrid3_2
      Dim a3_2 As CCustomerGrade

      Set a3_2 = New CCustomerGrade
      a3_2.CSTGRADE_ID = -1
      Status = a3_2.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a3_2 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 3-3" Then
      Call InitGrid3_3
      Dim a3_3 As CCustomerType

      Set a3_3 = New CCustomerType
      a3_3.CSTTYPE_ID = -1
      Status = a3_3.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a3_3 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 3-4" Then
      Call InitGrid3_4
      Dim a3_4 As CSupplierGrade

      Set a3_4 = New CSupplierGrade
      a3_4.SUPPLIER_GRADE_ID = -1
      Status = a3_4.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a3_4 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 3-5" Then
      Call InitGrid3_5
      Dim a3_5 As CSupplierType

      Set a3_5 = New CSupplierType
      a3_5.SUPPLIER_TYPE_ID = -1
      Status = a3_5.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a3_5 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 3-6" Then
      Call InitGrid3_6
      Dim a3_6 As CSupplierStatus

      Set a3_6 = New CSupplierStatus
      a3_6.SUPPLIER_STATUS_ID = -1
      Status = a3_6.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a3_6 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 3-7" Then
      Call InitGrid3_7
      Dim a3_7 As CEmpPosition

      Set a3_7 = New CEmpPosition
      a3_7.POSITION_ID = -1
      Status = a3_7.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a3_7 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 3-8" Then
      Call InitGrid3_8
      Dim A3_8 As CMasterRef

      Set A3_8 = New CMasterRef
      A3_8.KEY_ID = -1
      A3_8.MASTER_AREA = MEMO_TYPE
      Status = A3_8.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind

      Set A3_8 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 3-9" Then
      Call InitGrid3_9
      Dim A3_9 As CMasterRef

      Set A3_9 = New CMasterRef
      A3_9.KEY_ID = -1
      A3_9.MASTER_AREA = MEMO_STATUS
      Status = A3_9.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind

      Set A3_9 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 4-1" Then
      Call InitGrid4_1
      Dim a4_1 As CSellType

      Set a4_1 = New CSellType
      a4_1.SELL_TYPE_ID = -1
      Status = a4_1.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a4_1 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 4-2" Then
      Call InitGrid4_2
      Dim a4_2 As CDoType

      Set a4_2 = New CDoType
      a4_2.DO_TYPE_ID = -1
      Status = a4_2.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a4_2 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 6-1" Then
      Call InitGrid1_3
      Dim a6_1 As CUnit

      Set a6_1 = New CUnit
      a6_1.UNIT_ID = -1
      Status = a6_1.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a6_1 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 6-2" Then
      Call InitGrid6_2
      Dim a6_2 As CFeatureType
      
      Set a6_2 = New CFeatureType
      a6_2.FEATURE_TYPE_ID = -1
      Status = a6_2.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a6_2 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 6-3" Then
      Call InitGridMaster
      Dim a6_3 As CMasterRef

      Set a6_3 = New CMasterRef
      a6_3.KEY_ID = -1
      a6_3.MASTER_AREA = FEATURE_GROUP
      Status = a6_3.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a6_3 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 6-4" Then
      Call InitGridMaster
      Dim a6_4 As CMasterRef

      Set a6_4 = New CMasterRef
      a6_4.KEY_ID = -1
      a6_4.MASTER_AREA = PRODUCT_TYPE
      Status = a6_4.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a6_4 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 6-5" Then
      Call InitGridMaster
      Dim a6_5 As CMasterRef

      Set a6_5 = New CMasterRef
      a6_5.KEY_ID = -1
      a6_5.MASTER_AREA = CUSTOMER_SALE_TYPE
      Status = a6_5.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a6_5 = Nothing
 ElseIf Node.Key = ROOT_TREE & " 6-6" Then
      Call InitGridMaster
      Dim a6_6 As CMasterRef

      Set a6_6 = New CMasterRef
      a6_6.KEY_ID = -1
      a6_6.MASTER_AREA = ANIMAL_TYPE
      Status = a6_6.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a6_6 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 6-7" Then
      If Not VerifyAccessRight("MASTER_PACKAGE_DELIVERY-COST", Node.Text) Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      Call InitGridMaster
      Dim a6_7 As CMasterRef

      Set a6_7 = New CMasterRef
      a6_7.KEY_ID = -1
      a6_7.MASTER_AREA = TRANSPORT_DETAIL
      Status = a6_7.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a6_7 = Nothing
  ElseIf Node.Key = ROOT_TREE & " 6-8" Then
      If Not VerifyAccessRight("MASTER_PACKAGE_PROMOTIONAL", Node.Text) Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      Call InitGridMaster
      Dim a6_8 As CMasterRef

      Set a6_8 = New CMasterRef
      a6_8.KEY_ID = -1
      a6_8.MASTER_AREA = PROMOTIONAL_DETAIL
      Status = a6_8.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a6_8 = Nothing
      
   ElseIf Node.Key = ROOT_TREE & " 7-1" Then
      Call InitGrid7_1
      Dim a7_1 As CResource

      Set a7_1 = New CResource
      a7_1.RESOURCE_ID = -1
      Status = a7_1.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a7_1 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 7-2" Then
      Call InitGrid7_2
      Dim a7_2 As CMoneyFamily

      Set a7_2 = New CMoneyFamily
      a7_2.MONEY_FAMILY_ID = -1
      Status = a7_2.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind

      Set a7_2 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 7-3" Then
      Call InitGrid7_3
      Dim a7_3 As CBank

      Set a7_3 = New CBank
      a7_3.BANK_ID = -1
      Status = a7_3.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind

      Set a7_3 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 7-4" Then
      Call InitGrid7_4
      Dim a7_4 As CBankBranch

      Set a7_4 = New CBankBranch
      a7_4.BBRANCH_ID = -1
      Status = a7_4.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind

      Set a7_4 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 7-5" Then
      Call InitGrid7_5
      Dim a7_5 As CMasterRef

      Set a7_5 = New CMasterRef
      a7_5.KEY_ID = -1
      a7_5.MASTER_AREA = DRCR_REASON
      Status = a7_5.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind

      Set a7_5 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 7-6" Then
      Call InitGrid7_6
      Dim a7_6 As CMasterRef

      Set a7_6 = New CMasterRef
      a7_6.KEY_ID = -1
      a7_6.MASTER_AREA = BANK_ACCOUNT
      Status = a7_6.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind

      Set a7_6 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 7-7" Then
      Call InitGrid7_7
      Dim a7_7 As CMasterRef

      Set a7_7 = New CMasterRef
      a7_7.KEY_ID = -1
      a7_7.MASTER_AREA = CHEQUE_TYPE
      Status = a7_7.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind

      Set a7_7 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 7-8" Then
      Call InitGrid7_8
      Dim a7_8 As CMasterRef
      
      Set a7_8 = New CMasterRef
      a7_8.KEY_ID = -1
      a7_8.MASTER_AREA = ACCOUNT_LIST
      Status = a7_8.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a7_8 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 7-9" Then
      Call InitGrid7_9
      Dim a7_9 As CMasterRef
      
      Set a7_9 = New CMasterRef
      a7_9.KEY_ID = -1
      a7_9.MASTER_AREA = PAY_TO
      Status = a7_9.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a7_9 = Nothing
     ElseIf Node.Key = ROOT_TREE & " 7-10" Then
      Call InitGrid7_10
      Dim a7_10 As CMasterRef
      
      Set a7_10 = New CMasterRef
      a7_10.KEY_ID = -1
      a7_10.MASTER_AREA = CONDITION
      Status = a7_10.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a7_10 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 7-11" Then
      Call InitGrid7_11
      Dim a7_11 As CMasterRef
      
      Set a7_11 = New CMasterRef
      a7_11.KEY_ID = -1
      a7_11.MASTER_AREA = PAID_TYPE
      Status = a7_11.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a7_11 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 8-0" Then
      Call InitGrid8_0
      Dim a8_0 As CParameterProcess
      Set a8_0 = New CParameterProcess
      a8_0.PARAMETER_PROCESS_ID = -1
      Status = a8_0.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind

      Set a8_0 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 8-1" Then
      If Not VerifyAccessRight("MASTER_PRODUCTION_PROCESS", "PROCESS") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      Call InitGrid8_1
      Dim a8_1 As CProcess

      Set a8_1 = New CProcess
      a8_1.PROCESS_ID = -1
      Status = a8_1.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a8_1 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 8-2" Then
      Call InitGrid8_2
      Dim a8_2 As CFormulaType

      Set a8_2 = New CFormulaType
      a8_2.FORMULA_TYPE_ID = -1
      Status = a8_2.QueryData(1, m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a8_2 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 8-3" Then
      Call InitGrid8_3
      Dim a8_3 As CMachine

      Set a8_3 = New CMachine
      a8_3.MACHINE_ID = -1
      Status = a8_3.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a8_3 = Nothing
   Else
      Call InitGrid0
   End If
End Sub
Private Sub InitGrid3_8()
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
   Col.Width = 1620
   Col.Caption = MapText("รหัส")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5715
   Col.Caption = MapText("ประเภท MEMO")

   GridEX1.ItemCount = 0
End Sub
Private Sub InitGrid3_9()
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
   Col.Width = 1620
   Col.Caption = MapText("รหัส")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5715
   Col.Caption = MapText("สถานะ MEMO")

   GridEX1.ItemCount = 0
End Sub
Private Sub InitGridMaster()
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
   Col.Width = 2220
  Col.Caption = MapText("รหัส")
  
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5115
   Col.Caption = MapText("ชื่อ")

   GridEX1.ItemCount = 0
End Sub

