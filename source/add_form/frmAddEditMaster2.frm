VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditMaster2 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8550
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditMaster2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   8550
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame Frame1 
      Height          =   2415
      Left            =   -30
      TabIndex        =   6
      Top             =   600
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   4260
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboBankBranch 
         Height          =   510
         Left            =   2250
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1620
         Visible         =   0   'False
         Width           =   3435
      End
      Begin VB.ComboBox cboBank 
         Height          =   510
         Left            =   2250
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1200
         Visible         =   0   'False
         Width           =   3435
      End
      Begin prjFarmManagement.uctlTextBox txtCode 
         Height          =   435
         Left            =   2250
         TabIndex        =   0
         Top             =   270
         Width           =   1845
         _ExtentX        =   4683
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtName 
         Height          =   435
         Left            =   2250
         TabIndex        =   1
         Top             =   720
         Width           =   5745
         _ExtentX        =   4683
         _ExtentY        =   767
      End
      Begin VB.Label lblBankBranch 
         Alignment       =   1  'Right Justify
         Height          =   465
         Left            =   90
         TabIndex        =   14
         Top             =   1650
         Width           =   2055
      End
      Begin VB.Label lblBank 
         Alignment       =   1  'Right Justify
         Height          =   465
         Left            =   90
         TabIndex        =   13
         Top             =   1230
         Width           =   2055
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Height          =   465
         Left            =   60
         TabIndex        =   12
         Top             =   750
         Width           =   2055
      End
      Begin VB.Label lblCode 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   150
         TabIndex        =   7
         Top             =   300
         Width           =   1965
      End
   End
   Begin Threed.SSPanel pnlFooter 
      Height          =   705
      Left            =   0
      TabIndex        =   8
      Top             =   3000
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   1244
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSCommand cmdSave 
         Height          =   525
         Left            =   2670
         TabIndex        =   4
         Top             =   90
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdCancel 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4298
         TabIndex        =   5
         Top             =   90
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   615
         Index           =   0
         Left            =   11130
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   60
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   615
         Left            =   13230
         TabIndex        =   9
         Top             =   60
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditMaster2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Header As String
Public ShowMode As SHOW_MODE_TYPE
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public MasterKey As String

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
Private m_FormulaType As CFormulaType
Private m_Reason As CReason
Private m_Layout As CLayout
Private m_SellType As CSellType
Private m_DoType As CDoType
Private m_FeatureType As CFeatureType
Private m_Resource As CResource
Private m_Work As CWorkStatus
Private m_Religious As CReligious
Private m_Resign As CResignReason
Private m_BankAccount As CBankAccount
Private m_DocumentType As CDocumentType
Private m_MonthlyAdd As CMonthlyAdd
Private m_MonthlySub As CMonthlySub
Private m_Process As CProcess
Private m_Machine As CMachine
Private m_Money_family As CMoneyFamily
Private m_ParameterProcess As CParameterProcess
Private m_Bank As CBank
Private m_BankBranch As CBankBranch
Private m_Packaging As CPackaging
Private m_PurchaseExpense As CPurchaseExpense
Private m_MasterRef As CMasterRef

Public MasterMode As Long

Private Sub cboTextType_Click()
   m_HasModify = True
End Sub

Private Sub cboGroup_Click()
   m_HasModify = True
End Sub

Private Sub chkFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cboBank_Click()
Dim TempID As Long
   
   TempID = cboBank.ItemData(Minus2Zero(cboBank.ListIndex))
   If TempID > 0 Then
      Call LoadBankBranch(cboBankBranch, , TempID)
   End If
   m_HasModify = True
End Sub

Private Sub cboBankBranch_Click()
   m_HasModify = True
End Sub

Private Sub cmdCancel_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub InitFormLayout()
   Me.KeyPreview = True
   pnlHeader.Caption = HeaderText
   Me.BackColor = GLB_FORM_COLOR
   Frame1.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_HEAD_COLOR
   pnlFooter.BackColor = GLB_HEAD_COLOR
   Call InitHeaderFooter(pnlHeader, pnlFooter)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblCode, "")
   Call InitNormalLabel(lblName, "")
   
   If MasterKey = ROOT_TREE & " 1-1" Then
'      Call InitCombo(cboGroup)
'      Call LoadPartGroup(cboGroup)
'      cboGroup.Visible = True
      Call InitNormalLabel(lblCode, MapText("รหัสประเภทวัตถุดิบ"))
      Call InitNormalLabel(lblName, MapText("ประเภทวัตถุดิบ"))
   ElseIf MasterKey = ROOT_TREE & " 1-2" Then
'      Call InitCombo(cboGroup)
'      Call LoadPartGroup(cboGroup)
'      cboGroup.Visible = True
      Call InitNormalLabel(lblCode, MapText("รหัสสถานที่จัดเก็บ"))
      Call InitNormalLabel(lblName, MapText("สถานที่จัดเก็บ"))
   ElseIf MasterKey = ROOT_TREE & " 1-3" Then
'      Call InitCombo(cboGroup)
'      Call InitPeriodType(cboGroup)
'      cboGroup.Visible = True
      Call InitNormalLabel(lblCode, MapText("รหัสหน่วยวัด"))
      Call InitNormalLabel(lblName, MapText("หน่วยวัด"))
   ElseIf MasterKey = ROOT_TREE & " 1-4" Then
      Call InitNormalLabel(lblCode, MapText("รหัสกลุ่มวัตถุดิบ"))
      Call InitNormalLabel(lblName, MapText("กลุ่มวัตถุดิบ"))
   ElseIf MasterKey = ROOT_TREE & " 1-5" Then
      Call InitNormalLabel(lblCode, MapText("รหัสสาเหตุการเบิก"))
      Call InitNormalLabel(lblName, MapText("สาเหตุการเบิก"))
   ElseIf MasterKey = ROOT_TREE & " 1-6" Then
      Call InitNormalLabel(lblCode, MapText("รหัสสาเหตุการปรับยอด"))
      Call InitNormalLabel(lblName, MapText("สาเหตุการปรับยอด"))
   ElseIf MasterKey = ROOT_TREE & " 1-7" Then
      Call InitNormalLabel(lblCode, MapText("รหัสหน่วงาน/แผนก"))
      Call InitNormalLabel(lblName, MapText("หน่วงาน/แผนก"))
   ElseIf MasterKey = ROOT_TREE & " 1-10" Then
      Call InitNormalLabel(lblCode, MapText("รหัสรายจ่าย"))
      Call InitNormalLabel(lblName, MapText("รายจ่ายการเบิก"))
   ElseIf MasterKey = ROOT_TREE & " 2-1" Then
      Call InitNormalLabel(lblCode, MapText("รหัส"))
      Call InitNormalLabel(lblName, MapText("สถานะการทำงาน"))
      
    ElseIf MasterKey = ROOT_TREE & " 1-14" Then
      Call InitNormalLabel(lblCode, MapText("รหัสโครงการ"))
      Call InitNormalLabel(lblName, MapText("ชื่อโครงการ"))
   ElseIf MasterKey = ROOT_TREE & " 1-8" Then
'      Call txtWeightRate.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
'      lblWeightRate.Visible = True
'      txtWeightRate.Visible = True
'      Call InitNormalLabel(lblWeightRate, MapText("ก.ก./หน่วย"))
'      Call InitNormalLabel(lblCode, MapText("รหัสภาชนะ"))
'      Call InitNormalLabel(lblName, MapText("ภาชนะบรรจุ"))
   ElseIf MasterKey = ROOT_TREE & " 1-9" Then
'      Call txtWeightRate.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
'      lblWeightRate.Visible = True
'      txtWeightRate.Visible = True
'      Call InitNormalLabel(lblWeightRate, MapText("บาท/หน่วย"))
'      Call InitNormalLabel(lblCode, MapText("รหัสค่าใช้จ่าย"))
'      Call InitNormalLabel(lblName, MapText("ค่าใช้จ่ายผลิต"))
   ElseIf MasterKey = ROOT_TREE & " 2-1" Then
      Call InitNormalLabel(lblCode, MapText("รหัส"))
      Call InitNormalLabel(lblName, MapText("สถานะการทำงาน"))
   ElseIf MasterKey = ROOT_TREE & " 2-2" Then
      Call InitNormalLabel(lblCode, MapText("รหัส"))
      Call InitNormalLabel(lblName, MapText("ศาสนา"))
   ElseIf MasterKey = ROOT_TREE & " 2-3" Then
      Call InitNormalLabel(lblCode, MapText("รหัส"))
      Call InitNormalLabel(lblName, MapText("สาเหตุที่ออก"))
      ElseIf MasterKey = ROOT_TREE & " 2-4" Then
      Call InitNormalLabel(lblCode, MapText("รหัส"))
      Call InitNormalLabel(lblName, MapText("ชื่อธนาคาร"))
      ElseIf MasterKey = ROOT_TREE & " 2-5" Then
      Call InitNormalLabel(lblCode, MapText("รหัส"))
      Call InitNormalLabel(lblName, MapText("ประเภทบัตร"))
       ElseIf MasterKey = ROOT_TREE & " 2-6" Then
      Call InitNormalLabel(lblCode, MapText("รหัส"))
      Call InitNormalLabel(lblName, MapText("ส่วนบวกเงินเดือน"))
       ElseIf MasterKey = ROOT_TREE & " 2-7" Then
      Call InitNormalLabel(lblCode, MapText("รหัส"))
      Call InitNormalLabel(lblName, MapText("ส่วนหักเงินเดือน"))
   
   ElseIf MasterKey = ROOT_TREE & " 3-1" Then
      Call InitNormalLabel(lblCode, MapText("รหัสประเทศ"))
      Call InitNormalLabel(lblName, MapText("ประเทศ"))
   ElseIf MasterKey = ROOT_TREE & " 3-2" Then
      Call InitNormalLabel(lblCode, MapText("รหัสระดับลูกค้า"))
      Call InitNormalLabel(lblName, MapText("ระดับลูกค้า"))
   ElseIf MasterKey = ROOT_TREE & " 3-3" Then
      Call InitNormalLabel(lblCode, MapText("รหัสประเภทลูกค้า"))
      Call InitNormalLabel(lblName, MapText("ประเภทลูกค้า"))
   ElseIf MasterKey = ROOT_TREE & " 3-4" Then
      Call InitNormalLabel(lblCode, MapText("รหัสระดับซัพ ฯ"))
      Call InitNormalLabel(lblName, MapText("ระดับซัพ ฯ"))
   ElseIf MasterKey = ROOT_TREE & " 3-5" Then
      Call InitNormalLabel(lblCode, MapText("รหัสประเภทซัพ ฯ"))
      Call InitNormalLabel(lblName, MapText("ประเภทซัพ ฯ"))
   ElseIf MasterKey = ROOT_TREE & " 3-6" Then
      Call InitNormalLabel(lblCode, MapText("รหัสสถานะซัพ ฯ"))
      Call InitNormalLabel(lblName, MapText("สถานะซัพ ฯ"))
   ElseIf MasterKey = ROOT_TREE & " 3-7" Then
      Call InitNormalLabel(lblCode, MapText("รหัสตำแหน่ง"))
      Call InitNormalLabel(lblName, MapText("ตำแหน่ง"))
   
   ElseIf MasterKey = ROOT_TREE & " 4-1" Then
      Call InitNormalLabel(lblCode, MapText("รหัสราคาทอง"))
      Call InitNormalLabel(lblName, MapText("ชื่อราคาทอง"))
   ElseIf MasterKey = ROOT_TREE & " 4-2" Then
      Call InitNormalLabel(lblCode, MapText("รหัสประเภทบิล"))
      Call InitNormalLabel(lblName, MapText("ชื่อประเภทบิล"))
   ElseIf MasterKey = ROOT_TREE & " 6-1" Then
'      Call InitCombo(cboGroup)
'      Call InitPeriodType(cboGroup)
'      cboGroup.Visible = True
'      Call InitNormalLabel(lblCode, MapText("รหัสหน่วยวัด"))
'      Call InitNormalLabel(lblName, MapText("หน่วยวัด"))
   ElseIf MasterKey = ROOT_TREE & " 6-2" Then
'      Call InitCheckBox(chkFlag, "ค่าขนส่ง")
'      chkFlag.Visible = True
'      Call InitNormalLabel(lblCode, MapText("รหัสประเภท"))
'      Call InitNormalLabel(lblName, MapText("ประเภท"))
   ElseIf MasterKey = ROOT_TREE & " 7-1" Then
      Call InitNormalLabel(lblCode, MapText("หมายเลขทรัพยากร"))
      Call InitNormalLabel(lblName, MapText("รายละเอียด"))
    ElseIf MasterKey = ROOT_TREE & " 7-2" Then
      Call InitNormalLabel(lblCode, MapText("หมายเลขสกุลเงิน"))
      Call InitNormalLabel(lblName, MapText("ชื่อสกุลเงิน"))
   ElseIf MasterKey = ROOT_TREE & " 7-3" Then
      Call InitNormalLabel(lblCode, MapText("รหัสธนาคาร"))
      Call InitNormalLabel(lblName, MapText("ธนาคาร"))
   ElseIf MasterKey = ROOT_TREE & " 7-4" Then
'      cboGroup.Visible = True
'      Call InitCombo(cboGroup)
'      Call LoadBank(cboGroup)
'
'      Call InitNormalLabel(lblCode, MapText("รหัสสาขาธนาคาร"))
'      Call InitNormalLabel(lblName, MapText("สาขาธนาคาร"))
   ElseIf MasterKey = ROOT_TREE & " 7-5" Then
      Call InitNormalLabel(lblCode, MapText("รหัสสาเหตุ"))
      Call InitNormalLabel(lblName, MapText("สาเหตุการเพิ่ม/ลดหนี้"))
   ElseIf MasterKey = ROOT_TREE & " 7-6" Then
      Call InitCombo(cboBank)
      Call InitCombo(cboBankBranch)
      
      Call LoadBank(cboBank)
      cboBank.Visible = True
      cboBankBranch.Visible = True
      
      Call InitNormalLabel(lblCode, MapText("รหัสบัญชี"))
      Call InitNormalLabel(lblName, MapText("เลขที่บัญชีธนาคาร"))
      Call InitNormalLabel(lblBank, MapText("ธนาคาร"))
      Call InitNormalLabel(lblBankBranch, MapText("สาขา"))
    ElseIf MasterKey = ROOT_TREE & " 8-0" Then
      Call InitNormalLabel(lblCode, MapText("รหัสค่าใช้จ่าย"))
      Call InitNormalLabel(lblName, MapText("ค่าใช้จ่ายผลิต"))
   ElseIf MasterKey = ROOT_TREE & " 8-1" Then
      Call InitNormalLabel(lblCode, MapText("หมายเลขโปรเซส"))
      Call InitNormalLabel(lblName, MapText("ชื่อโปรเซส"))
   ElseIf MasterKey = ROOT_TREE & " 8-2" Then
'      Call InitNormalLabel(lblCode, MapText("รหัสประเภทสูตร"))
'      Call InitNormalLabel(lblName, MapText("ประเภทสูตร"))
'      Call InitCheckBox(chkFlag, "Intermediat")
'      chkFlag.Visible = True
   ElseIf MasterKey = ROOT_TREE & " 8-3" Then
      Call InitNormalLabel(lblCode, MapText("หมายเลขเครื่องจักร"))
      Call InitNormalLabel(lblName, MapText("เครื่องจักร"))
   End If

   Call txtCode.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   Call txtName.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)

   Call InitMainButton(cmdSave, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdCancel, MapText("ยกเลิก (ESC)"))
      
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   pnlFooter.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   Frame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   cmdCancel.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSave.Picture = LoadPicture(glbParameterObj.NormalButton1)
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If MasterKey = ROOT_TREE & " 1-1" Then
         m_PartType.PART_TYPE_ID = ID
         Call m_PartType.QueryData(m_Rs, itemcount)
'         If ItemCount > 0 Then
'            Call m_PartType.PopulateFromRS(1, m_Rs)
'            txtCode.Text = m_PartType.PART_TYPE_NO
'            txtName.Text = m_PartType.PART_TYPE_NAME
'            cboGroup.ListIndex = IDToListIndex(cboGroup, m_PartType.PART_GROUP_ID)
'         End If
      ElseIf MasterKey = ROOT_TREE & " 1-2" Then
         m_Location.LOCATION_ID = ID
         m_Location.LOCATION_TYPE = 2
         Call m_Location.QueryData(m_Rs, itemcount)
         If itemcount > 0 Then
            Call m_Location.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_Location.LOCATION_NO
            txtName.Text = m_Location.LOCATION_NAME
'            cboGroup.ListIndex = IDToListIndex(cboGroup, m_Location.PART_GROUP_ID)
         End If
      ElseIf MasterKey = ROOT_TREE & " 1-3" Then
         m_Unit.UNIT_ID = ID
         Call m_Unit.QueryData(m_Rs, itemcount)
         If itemcount > 0 Then
            Call m_Unit.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_Unit.UNIT_NO
            txtName.Text = m_Unit.UNIT_NAME
'            cboGroup.ListIndex = IDToListIndex(cboGroup, m_Unit.PERIOD_TYPE)
         End If
      ElseIf MasterKey = ROOT_TREE & " 1-4" Then
         m_PartGroup.PART_GROUP_ID = ID
         Call m_PartGroup.QueryData(m_Rs, itemcount)
         If itemcount > 0 Then
            Call m_PartGroup.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_PartGroup.PART_GROUP_NO
            txtName.Text = m_PartGroup.PART_GROUP_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 1-5" Then
         m_Reason.REASON_ID = ID
         m_Reason.Area = 1
         Call m_Reason.QueryData(m_Rs, itemcount)
         If itemcount > 0 Then
            Call m_Reason.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_Reason.REASON_NO
            txtName.Text = m_Reason.REASON_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 1-6" Then
         m_Reason.REASON_ID = ID
         m_Reason.Area = 2
         Call m_Reason.QueryData(m_Rs, itemcount)
         If itemcount > 0 Then
            Call m_Reason.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_Reason.REASON_NO
            txtName.Text = m_Reason.REASON_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 1-7" Then
         m_Layout.LAY_OUT_ID = ID
         m_Layout.LOCATION_ID = -1
         Call m_Layout.QueryData(m_Rs, itemcount)
         If itemcount > 0 Then
            Call m_Layout.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_Layout.LAY_OUT_NO
            txtName.Text = m_Layout.LAY_OUT_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 1-8" Then
         m_Packaging.PACKAGING_ID = ID
         Call m_Packaging.QueryData(m_Rs, itemcount)
         If itemcount > 0 Then
            Call m_Packaging.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_Packaging.PACKAGING_NO
            txtName.Text = m_Packaging.PACKAGING_NAME
'            txtWeightRate.Text = m_Packaging.WEIGHT_RATE
         End If
      ElseIf MasterKey = ROOT_TREE & " 1-9" Then
         m_PurchaseExpense.PUREXP_ID = ID
         Call m_PurchaseExpense.QueryData(m_Rs, itemcount)
         If itemcount > 0 Then
            Call m_PurchaseExpense.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_PurchaseExpense.PUREXP_NO
            txtName.Text = m_PurchaseExpense.PUREXP_NAME
'            txtWeightRate.Text = m_PurchaseExpense.EXPENSE_RATE
         End If
      ElseIf MasterKey = ROOT_TREE & " 1-10" Then
         m_MasterRef.KEY_ID = ID
         Call m_MasterRef.QueryData(m_Rs, itemcount)
         If itemcount > 0 Then
            Call m_MasterRef.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_MasterRef.KEY_CODE
            txtName.Text = m_MasterRef.KEY_NAME
         End If
     ElseIf MasterKey = ROOT_TREE & " 1-14" Then
         m_MasterRef.KEY_ID = ID
         Call m_MasterRef.QueryData(m_Rs, itemcount)
         If itemcount > 0 Then
            Call m_MasterRef.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_MasterRef.KEY_CODE
            txtName.Text = m_MasterRef.KEY_NAME
         End If
ElseIf MasterKey = ROOT_TREE & " 2-1" Then
         m_Work.WORK_ID = ID
         Call m_Work.QueryData(m_Rs, itemcount)
         If itemcount > 0 Then
            Call m_Work.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_Work.WORK_NO
            txtName.Text = m_Work.WORK_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 2-2" Then
         m_Religious.RELIGIOUS_ID = ID
         Call m_Religious.QueryData(m_Rs, itemcount)
         If itemcount > 0 Then
            Call m_Religious.PopulateFromRS(m_Rs)
            txtCode.Text = m_Religious.RELIGIOUS_NO
            txtName.Text = m_Religious.RELIGIOUS_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 2-3" Then
         m_Resign.RSGRESON_ID = ID
         Call m_Resign.QueryData(m_Rs, itemcount)
         If itemcount > 0 Then
            Call m_Resign.PopulateFromRS(m_Rs)
            txtCode.Text = m_Resign.RSGRESON_NO
            txtName.Text = m_Resign.RSGRESON_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 2-4" Then
         m_BankAccount.BANK_ID = ID
         Call m_BankAccount.QueryData(m_Rs, itemcount)
         If itemcount > 0 Then
            Call m_BankAccount.PopulateFromRS(m_Rs)
            txtCode.Text = m_BankAccount.BANK_NO
            txtName.Text = m_BankAccount.BANK_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 2-5" Then
         m_DocumentType.DOCTYPE_ID = ID
         Call m_DocumentType.QueryData(m_Rs, itemcount)
         If itemcount > 0 Then
            Call m_DocumentType.PopulateFromRS(m_Rs)
            txtCode.Text = m_DocumentType.DOCTYPE_NO
            txtName.Text = m_DocumentType.DOCTYPE_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 2-6" Then
         m_MonthlyAdd.MONTHLY_ADD_ID = ID
         Call m_MonthlyAdd.QueryData(m_Rs, itemcount)
         If itemcount > 0 Then
            Call m_MonthlyAdd.PopulateFromRS(m_Rs)
            txtCode.Text = m_MonthlyAdd.MONTHLY_ADD_NO
            txtName.Text = m_MonthlyAdd.MONTHLY_ADD_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 2-7" Then
         m_MonthlySub.MONTHLY_SUB_ID = ID
         Call m_MonthlySub.QueryData(m_Rs, itemcount)
         If itemcount > 0 Then
            Call m_MonthlySub.PopulateFromRS(m_Rs)
            txtCode.Text = m_MonthlySub.MONTHLY_SUB_NO
            txtName.Text = m_MonthlySub.MONTHLY_SUB_NAME
             txtName.Enabled = True
            If ID = 1 Then
            txtName.Enabled = False
            End If
         End If
            ElseIf MasterKey = ROOT_TREE & " 3-1" Then
         m_Country.COUNTRY_ID = ID
         Call m_Country.QueryData(m_Rs, itemcount)
         If itemcount > 0 Then
            Call m_Country.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_Country.COUNTRY_NO
            txtName.Text = m_Country.COUNTRY_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 3-2" Then
         m_CustomerGrade.CSTGRADE_ID = ID
         Call m_CustomerGrade.QueryData(m_Rs, itemcount)
         If itemcount > 0 Then
            Call m_CustomerGrade.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_CustomerGrade.CSTGRADE_NO
            txtName.Text = m_CustomerGrade.CSTGRADE_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 3-3" Then
         m_CustomerType.CSTTYPE_ID = ID
         Call m_CustomerType.QueryData(m_Rs, itemcount)
         If itemcount > 0 Then
            Call m_CustomerType.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_CustomerType.CSTTYPE_NO
            txtName.Text = m_CustomerType.CSTTYPE_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 3-4" Then
         m_SupplierGrade.SUPPLIER_GRADE_ID = ID
         Call m_SupplierGrade.QueryData(m_Rs, itemcount)
         If itemcount > 0 Then
            Call m_SupplierGrade.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_SupplierGrade.SUPPLIER_GRADE_NO
            txtName.Text = m_SupplierGrade.SUPPLIER_GRADE_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 3-5" Then
         m_SupplierType.SUPPLIER_TYPE_ID = ID
         Call m_SupplierType.QueryData(m_Rs, itemcount)
         If itemcount > 0 Then
            Call m_SupplierType.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_SupplierType.SUPPLIER_TYPE_NO
            txtName.Text = m_SupplierType.SUPPLIER_TYPE_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 3-6" Then
         m_SupplierStatus.SUPPLIER_STATUS_ID = ID
         Call m_SupplierStatus.QueryData(m_Rs, itemcount)
         If itemcount > 0 Then
            Call m_SupplierStatus.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_SupplierStatus.SUPPLIER_STATUS_NO
            txtName.Text = m_SupplierStatus.SUPPLIER_STATUS_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 3-7" Then
         m_Position.POSITION_ID = ID
         Call m_Position.QueryData(m_Rs, itemcount)
         If itemcount > 0 Then
            Call m_Position.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_Position.POSITION_NAME
            txtName.Text = m_Position.POSITION_DESC
         End If
      ElseIf MasterKey = ROOT_TREE & " 4-1" Then
         m_SellType.SELL_TYPE_ID = ID
         Call m_SellType.QueryData(m_Rs, itemcount)
         If itemcount > 0 Then
            Call m_SellType.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_SellType.SELL_TYPE_NO
            txtName.Text = m_SellType.SELL_TYPE_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 4-2" Then
         m_DoType.DO_TYPE_ID = ID
         Call m_DoType.QueryData(m_Rs, itemcount)
         If itemcount > 0 Then
            Call m_DoType.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_DoType.DO_TYPE_NO
            txtName.Text = m_DoType.DO_TYPE_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 6-1" Then
         m_Unit.UNIT_ID = ID
         Call m_Unit.QueryData(m_Rs, itemcount)
         If itemcount > 0 Then
            Call m_Unit.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_Unit.UNIT_NO
            txtName.Text = m_Unit.UNIT_NAME
'            cboGroup.ListIndex = IDToListIndex(cboGroup, m_Unit.PERIOD_TYPE)
         End If
      ElseIf MasterKey = ROOT_TREE & " 6-2" Then
         m_FeatureType.FEATURE_TYPE_ID = ID
         Call m_FeatureType.QueryData(m_Rs, itemcount)
         If itemcount > 0 Then
            Call m_FeatureType.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_FeatureType.FEATURE_TYPE_NO
            txtName.Text = m_FeatureType.FEATURE_TYPE_NAME
'            chkFlag.Value = FlagToCheck(m_FeatureType.LOGISTIC_FLAG)
         End If
      ElseIf MasterKey = ROOT_TREE & " 7-1" Then
         m_Resource.RESOURCE_ID = ID
         Call m_Resource.QueryData(m_Rs, itemcount)
         If itemcount > 0 Then
            Call m_Resource.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_Resource.RESOURCE_NO
            txtName.Text = m_Resource.RESOURCE_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 7-2" Then
         m_Money_family.MONEY_FAMILY_ID = ID
         Call m_Money_family.QueryData(m_Rs, itemcount)
         If itemcount > 0 Then
            Call m_Money_family.PopulateFromRS(m_Rs)
            txtCode.Text = m_Money_family.MONEY_FAMILY_NO
            txtName.Text = m_Money_family.MONEY_FAMILY_NAME
         End If
      
      ElseIf MasterKey = ROOT_TREE & " 7-3" Then
         m_Bank.BANK_ID = ID
         Call m_Bank.QueryData(m_Rs, itemcount)
         If itemcount > 0 Then
            Call m_Bank.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_Bank.BANK_NO
            txtName.Text = m_Bank.BANK_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 7-4" Then
         m_BankBranch.BBRANCH_ID = ID
         Call m_BankBranch.QueryData(m_Rs, itemcount)
         If itemcount > 0 Then
            Call m_BankBranch.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_BankBranch.BBRANCH_NO
            txtName.Text = m_BankBranch.BBRANCH_NAME
'            cboGroup.ListIndex = IDToListIndex(cboGroup, m_BankBranch.BANK_ID)
         End If
      ElseIf MasterKey = ROOT_TREE & " 7-5" Then
         m_MasterRef.KEY_ID = ID
         Call m_MasterRef.QueryData(m_Rs, itemcount)
         If itemcount > 0 Then
            Call m_MasterRef.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_MasterRef.KEY_CODE
            txtName.Text = m_MasterRef.KEY_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 7-6" Then
         m_MasterRef.KEY_ID = ID
         Call m_MasterRef.QueryData(m_Rs, itemcount)
         If itemcount > 0 Then
            Call m_MasterRef.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_MasterRef.KEY_CODE
            txtName.Text = m_MasterRef.KEY_NAME
            cboBank.ListIndex = IDToListIndex(cboBank, m_MasterRef.TEMP_ID1)
            cboBankBranch.ListIndex = IDToListIndex(cboBankBranch, m_MasterRef.TEMP_ID2)
         End If
      ElseIf MasterKey = ROOT_TREE & " 8-0" Then
         m_ParameterProcess.PARAMETER_PROCESS_ID = ID
         Call m_ParameterProcess.QueryData(m_Rs, itemcount)
         If itemcount > 0 Then
            Call m_ParameterProcess.PopulateFromRS(m_Rs)
            txtCode.Text = m_ParameterProcess.PARAMETER_PROCESS_NO
            txtName.Text = m_ParameterProcess.PARAMETER_PROCESS_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 8-1" Then
         m_Process.PROCESS_ID = ID
         Call m_Process.QueryData(m_Rs, itemcount)
         If itemcount > 0 Then
            Call m_Process.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_Process.PROCESS_NO
            txtName.Text = m_Process.PROCESS_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 8-2" Then
         m_FormulaType.FORMULA_TYPE_ID = ID
         Call m_FormulaType.QueryData(1, m_Rs, itemcount)
         If itemcount > 0 Then
            Call m_FormulaType.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_FormulaType.FORMULA_TYPE_NO
            txtName.Text = m_FormulaType.FORMULA_TYPE_NAME
'            chkFlag.Value = FlagToCheck(m_FormulaType.INTERMEDIAT_FLAG)
         End If
      ElseIf MasterKey = ROOT_TREE & " 8-3" Then
         m_Machine.MACHINE_ID = ID
         Call m_Machine.QueryData(m_Rs, itemcount)
         If itemcount > 0 Then
            Call m_Machine.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_Machine.MACHINE_NO
            txtName.Text = m_Machine.MACHINE_NAME
         End If
      End If
   
      Call EnableForm(Me, True)
   End If
   
   IsOK = True
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cmdSave_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Function SaveData() As Boolean
On Error GoTo ErrorHandler
Dim IsOK As Boolean
   
   If Not VerifyTextControl(lblCode, txtCode, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblName, txtName, Not txtName.Visible) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Call EnableForm(Me, False)
      
   If MasterKey = ROOT_TREE & " 1-1" Then
'      If Not VerifyCombo(lblCode, cboGroup, False) Then
'         Call EnableForm(Me, True)
'         Exit Function
'      End If
      
      If Not CheckUniqueNs(PARTTYPE_NO, txtCode.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
         Call EnableForm(Me, True)
         txtCode.SetFocus
         Exit Function
      End If
   
      If Not CheckUniqueNs(PARTTYPE_NAME, txtName.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtName.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
         Call EnableForm(Me, True)
         txtName.SetFocus
         Exit Function
      End If
   
      m_PartType.AddEditMode = ShowMode
      m_PartType.PART_TYPE_NAME = txtName.Text
      m_PartType.RAW_FLAG = "Y"
      m_PartType.PART_TYPE_NO = txtCode.Text
'      m_PartType.PART_GROUP_ID = cboGroup.ItemData(Minus2Zero(cboGroup.ListIndex))
      Call glbMaster.AddEditPartType(m_PartType, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 1-2" Then
      If Not CheckUniqueNs(LOCATION_NO, txtCode.Text & "2", ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
         Call EnableForm(Me, True)
         txtCode.SetFocus
         Exit Function
      End If
   
      If Not CheckUniqueNs(LOCATION_NAME, txtName.Text & "2", ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtName.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtName.SetFocus
         Exit Function
      End If
      
      m_Location.AddEditMode = ShowMode
      m_Location.LOCATION_NAME = txtName.Text
      m_Location.LOCATION_NO = txtCode.Text
      m_Location.LOCATION_TYPE = 2 'คลัง
      m_Location.SALE_FLAG = "N"
'      m_Location.PART_GROUP_ID = cboGroup.ItemData(Minus2Zero(cboGroup.ListIndex))
      Call glbMaster.AddEditLocation(m_Location, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 1-3" Then
      If Not CheckUniqueNs(UNIT_NO, txtCode.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtCode.SetFocus
         Exit Function
      End If
   
      If Not CheckUniqueNs(UNIT_NAME, txtName.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtName.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtName.SetFocus
         Exit Function
      End If
      
      m_Unit.AddEditMode = ShowMode
      m_Unit.UNIT_NAME = txtName.Text
      m_Unit.UNIT_NO = txtCode.Text
'      m_Unit.PERIOD_TYPE = cboGroup.ItemData(Minus2Zero(cboGroup.ListIndex))
      Call glbMaster.AddEditUnit(m_Unit, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 1-4" Then
      If Not CheckUniqueNs(PARTGROUP_NO, txtCode.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtCode.SetFocus
         Exit Function
      End If
   
      If Not CheckUniqueNs(PARTGROUP_NAME, txtName.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtName.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtName.SetFocus
         Exit Function
      End If
      
      m_PartGroup.AddEditMode = ShowMode
      m_PartGroup.PART_GROUP_NAME = txtName.Text
      m_PartGroup.PART_GROUP_NO = txtCode.Text
      Call glbMaster.AddEditPartGroup(m_PartGroup, IsOK, glbErrorLog)
   
   ElseIf MasterKey = ROOT_TREE & " 1-5" Then
      m_Reason.AddEditMode = ShowMode
      m_Reason.REASON_NAME = txtName.Text
      m_Reason.REASON_NO = txtCode.Text
      m_Reason.Area = 1
      Call glbMaster.AddEditReason(m_Reason, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 1-6" Then
      m_Reason.AddEditMode = ShowMode
      m_Reason.REASON_NAME = txtName.Text
      m_Reason.REASON_NO = txtCode.Text
      m_Reason.Area = 2
      Call glbMaster.AddEditReason(m_Reason, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 1-7" Then
      m_Layout.AddEditMode = ShowMode
      m_Layout.LAY_OUT_NAME = txtName.Text
      m_Layout.LAY_OUT_NO = txtCode.Text
      m_Layout.LOCATION_ID = -1
      Call glbMaster.AddEditLayout(m_Layout, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 1-8" Then
      m_Packaging.AddEditMode = ShowMode
      m_Packaging.PACKAGING_NAME = txtName.Text
      m_Packaging.PACKAGING_NO = txtCode.Text
'      m_Packaging.WEIGHT_RATE = Val(txtWeightRate.Text)
      Call glbMaster.AddEditPackaging(m_Packaging, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 1-9" Then
      m_PurchaseExpense.AddEditMode = ShowMode
      m_PurchaseExpense.PUREXP_NAME = txtName.Text
      m_PurchaseExpense.PUREXP_NO = txtCode.Text
'      m_PurchaseExpense.EXPENSE_RATE = Val(txtWeightRate.Text)
      Call glbMaster.AddEditPurchaseExpense(m_PurchaseExpense, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 1-10" Then
      m_MasterRef.AddEditMode = ShowMode
      m_MasterRef.MASTER_AREA = EXPENSE_TYPE
      m_MasterRef.KEY_NAME = txtName.Text
      m_MasterRef.KEY_CODE = txtCode.Text
      Call glbMaster.AddEditMasterRef(m_MasterRef, IsOK, True, glbErrorLog)
    ElseIf MasterKey = ROOT_TREE & " 1-14" Then
      m_MasterRef.AddEditMode = ShowMode
      m_MasterRef.MASTER_AREA = SET_PROJECT
      m_MasterRef.KEY_NAME = txtName.Text
      m_MasterRef.KEY_CODE = txtCode.Text
      Call glbMaster.AddEditMasterRef(m_MasterRef, IsOK, True, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 2-1" Then
      If Not CheckUniqueNs(WORK_NO, txtCode.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtCode.SetFocus
         Exit Function
      End If
   
      If Not CheckUniqueNs(WORK_NAME, txtName.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtName.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtName.SetFocus
         Exit Function
      End If
       m_Work.AddEditMode = ShowMode
      m_Work.WORK_NAME = txtName.Text
      m_Work.WORK_NO = txtCode.Text
      Call glbMaster.AddEditWorkStatus(m_Work, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 2-2" Then
      If Not CheckUniqueNs(RELIGIOUS_NO, txtCode.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtCode.SetFocus
         Exit Function
      End If
   
      If Not CheckUniqueNs(RELIGIOUS_NAME, txtName.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtName.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtName.SetFocus
         Exit Function
      End If
   
      m_Religious.AddEditMode = ShowMode
      m_Religious.RELIGIOUS_NAME = txtName.Text
      m_Religious.RELIGIOUS_NO = txtCode.Text
      Call glbMaster.AddEditReligious(m_Religious, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 2-3" Then
      If Not CheckUniqueNs(RESIGN_NO, txtCode.Text & "1", ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtCode.SetFocus
         Exit Function
      End If
   
      If Not CheckUniqueNs(RESIGN_NAME, txtName.Text & "1", ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtName.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtName.SetFocus
         Exit Function
      End If
   
      m_Resign.AddEditMode = ShowMode
      m_Resign.RSGRESON_NAME = txtName.Text
      m_Resign.RSGRESON_NO = txtCode.Text
      Call glbMaster.AddEditResign(m_Resign, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 2-4" Then
      If Not CheckUniqueNs(BANK_NO, txtCode.Text & "1", ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtCode.SetFocus
         Exit Function
      End If
   
      If Not CheckUniqueNs(BANK_NAME, txtName.Text & "1", ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtName.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtName.SetFocus
         Exit Function
      End If
   
      m_BankAccount.AddEditMode = ShowMode
      m_BankAccount.BANK_NAME = txtName.Text
      m_BankAccount.BANK_NO = txtCode.Text
      Call glbMaster.AddEditBankAccount(m_BankAccount, IsOK, glbErrorLog)
      ElseIf MasterKey = ROOT_TREE & " 2-5" Then
      If Not CheckUniqueNs(DOC_NO, txtCode.Text & "1", ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtCode.SetFocus
         Exit Function
      End If
   
      If Not CheckUniqueNs(DOC_NAME, txtName.Text & "1", ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtName.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtName.SetFocus
         Exit Function
      End If
   
      m_DocumentType.AddEditMode = ShowMode
      m_DocumentType.DOCTYPE_NAME = txtName.Text
      m_DocumentType.DOCTYPE_NO = txtCode.Text
      Call glbMaster.AddEditDocumentType(m_DocumentType, IsOK, glbErrorLog)

   ElseIf MasterKey = ROOT_TREE & " 2-6" Then
      If Not CheckUniqueNs(MONTHLY_ADD_NO, txtCode.Text & "1", ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtCode.SetFocus
         Exit Function
      End If
   
      If Not CheckUniqueNs(MONTHLY_ADD_NAME, txtName.Text & "1", ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtName.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtName.SetFocus
         Exit Function
      End If
   
      m_MonthlyAdd.AddEditMode = ShowMode
      m_MonthlyAdd.MONTHLY_ADD_NAME = txtName.Text
      m_MonthlyAdd.MONTHLY_ADD_NO = txtCode.Text
      Call glbMaster.AddEditMonthlyAdd(m_MonthlyAdd, IsOK, glbErrorLog)

   ElseIf MasterKey = ROOT_TREE & " 2-7" Then
      If Not CheckUniqueNs(MONTHLY_SUB_NO, txtCode.Text & "1", ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtCode.SetFocus
         Exit Function
      End If
   
      If Not CheckUniqueNs(MONTHLY_SUB_NAME, txtName.Text & "1", ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtName.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtName.SetFocus
         Exit Function
      End If
   
      m_MonthlySub.AddEditMode = ShowMode
      m_MonthlySub.MONTHLY_SUB_NAME = txtName.Text
      m_MonthlySub.MONTHLY_SUB_NO = txtCode.Text
      Call glbMaster.AddEditMonthlySub(m_MonthlySub, IsOK, glbErrorLog)
   
   ElseIf MasterKey = ROOT_TREE & " 3-1" Then
      If Not CheckUniqueNs(COUNTRY_NO, txtCode.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtCode.SetFocus
         Exit Function
      End If
   
      If Not CheckUniqueNs(COUNTRY_NAME, txtName.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtName.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtName.SetFocus
         Exit Function
      End If
   
      m_Country.AddEditMode = ShowMode
      m_Country.COUNTRY_NAME = txtName.Text
      m_Country.COUNTRY_NO = txtCode.Text
      Call glbMaster.AddEditCountry(m_Country, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 3-2" Then
      If Not CheckUniqueNs(CSTGRADE_NO, txtCode.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtCode.SetFocus
         Exit Function
      End If
   
      If Not CheckUniqueNs(CSTGRADE_NAME, txtName.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtName.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtName.SetFocus
         Exit Function
      End If
   
      m_CustomerGrade.AddEditMode = ShowMode
      m_CustomerGrade.CSTGRADE_NAME = txtName.Text
      m_CustomerGrade.CSTGRADE_NO = txtCode.Text
      Call glbMaster.AddEditCustomerGrade(m_CustomerGrade, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 3-3" Then
      If Not CheckUniqueNs(CSTTYPE_NO, txtCode.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtCode.SetFocus
         Exit Function
      End If
   
      If Not CheckUniqueNs(CSTTYPE_NAME, txtName.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtName.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtName.SetFocus
         Exit Function
      End If
   
      m_CustomerType.AddEditMode = ShowMode
      m_CustomerType.CSTTYPE_NAME = txtName.Text
      m_CustomerType.CSTTYPE_NO = txtCode.Text
      Call glbMaster.AddEditCustomerType(m_CustomerType, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 3-4" Then
      If Not CheckUniqueNs(SUPPLIERGRADE_NO, txtCode.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtCode.SetFocus
         Exit Function
      End If
   
      If Not CheckUniqueNs(SUPPLIERGRADE_NAME, txtName.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtName.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtName.SetFocus
         Exit Function
      End If
   
      m_SupplierGrade.AddEditMode = ShowMode
      m_SupplierGrade.SUPPLIER_GRADE_NAME = txtName.Text
      m_SupplierGrade.SUPPLIER_GRADE_NO = txtCode.Text
      Call glbMaster.AddEditSupplierGrade(m_SupplierGrade, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 3-5" Then
      If Not CheckUniqueNs(SUPPLIERTYPE_NO, txtCode.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtCode.SetFocus
         Exit Function
      End If
   
      If Not CheckUniqueNs(SUPPLIERYPE_NAME, txtName.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtName.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtName.SetFocus
         Exit Function
      End If
   
      m_SupplierType.AddEditMode = ShowMode
      m_SupplierType.SUPPLIER_TYPE_NAME = txtName.Text
      m_SupplierType.SUPPLIER_TYPE_NO = txtCode.Text
      Call glbMaster.AddEditSupplierType(m_SupplierType, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 3-6" Then
      If Not CheckUniqueNs(SUPPLIERSTATUS_NO, txtCode.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtCode.SetFocus
         Exit Function
      End If
   
      If Not CheckUniqueNs(SUPPLIERSTATUS_NAME, txtName.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtName.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtName.SetFocus
         Exit Function
      End If
   
      m_SupplierStatus.AddEditMode = ShowMode
      m_SupplierStatus.SUPPLIER_STATUS_NAME = txtName.Text
      m_SupplierStatus.SUPPLIER_STATUS_NO = txtCode.Text
      Call glbMaster.AddEditSupplierStatus(m_SupplierStatus, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 3-7" Then
      If Not CheckUniqueNs(POSITION_NO, txtCode.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtCode.SetFocus
         Exit Function
      End If
      
      m_Position.AddEditMode = ShowMode
      m_Position.POSITION_DESC = txtName.Text
      m_Position.POSITION_NAME = txtCode.Text
      Call glbMaster.AddEditPosition(m_Position, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 4-1" Then
      m_SellType.AddEditMode = ShowMode
      m_SellType.SELL_TYPE_NAME = txtName.Text
      m_SellType.SELL_TYPE_NO = txtCode.Text
      Call glbMaster.AddEditSellType(m_SellType, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 4-2" Then
      m_DoType.AddEditMode = ShowMode
      m_DoType.DO_TYPE_NAME = txtName.Text
      m_DoType.DO_TYPE_NO = txtCode.Text
      Call glbMaster.AddEditDoType(m_DoType, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 6-1" Then
      If Not CheckUniqueNs(UNIT_NO, txtCode.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtCode.SetFocus
         Exit Function
      End If
   
      If Not CheckUniqueNs(UNIT_NAME, txtName.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtName.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtName.SetFocus
         Exit Function
      End If
      
      m_Unit.AddEditMode = ShowMode
      m_Unit.UNIT_NAME = txtName.Text
      m_Unit.UNIT_NO = txtCode.Text
'      m_Unit.PERIOD_TYPE = cboGroup.ItemData(Minus2Zero(cboGroup.ListIndex))
      Call glbMaster.AddEditUnit(m_Unit, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 6-2" Then
      m_FeatureType.AddEditMode = ShowMode
      m_FeatureType.FEATURE_TYPE_NAME = txtName.Text
      m_FeatureType.FEATURE_TYPE_NO = txtCode.Text
'      m_FeatureType.LOGISTIC_FLAG = Check2Flag(chkFlag.Value)
      Call glbMaster.AddEditFeatureType(m_FeatureType, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 7-1" Then
      m_Resource.AddEditMode = ShowMode
      m_Resource.RESOURCE_NAME = txtName.Text
      m_Resource.RESOURCE_NO = txtCode.Text
      Call glbMaster.AddEditResource(m_Resource, IsOK, glbErrorLog)
  ElseIf MasterKey = ROOT_TREE & " 7-2" Then
'      If Not CheckUniqueNs(MONEY_FAMILY_NO, txtCode.Text & "2", ID) Then
'         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
'         glbErrorLog.ShowUserError
'
'         Call EnableForm(Me, True)
'         txtCode.SetFocus
'         Exit Function
'      End If

      m_Money_family.AddEditMode = ShowMode
      m_Money_family.MONEY_FAMILY_NAME = txtName.Text
      m_Money_family.MONEY_FAMILY_NO = txtCode.Text
      Call glbMaster.AddEditMoneyFamily(m_Money_family, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 7-3" Then
      m_Bank.AddEditMode = ShowMode
      m_Bank.BANK_NAME = txtName.Text
      m_Bank.BANK_NO = txtCode.Text
      Call glbMaster.AddEditBank(m_Bank, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 7-4" Then
      m_BankBranch.AddEditMode = ShowMode
      m_BankBranch.BBRANCH_NAME = txtName.Text
      m_BankBranch.BBRANCH_NO = txtCode.Text
'      m_BankBranch.BANK_ID = cboGroup.ItemData(Minus2Zero(cboGroup.ListIndex))
      Call glbMaster.AddEditBankBranch(m_BankBranch, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 7-5" Then
      m_MasterRef.AddEditMode = ShowMode
      m_MasterRef.MASTER_AREA = DRCR_REASON
      m_MasterRef.KEY_NAME = txtName.Text
      m_MasterRef.KEY_CODE = txtCode.Text
      Call glbMaster.AddEditMasterRef(m_MasterRef, IsOK, True, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 7-6" Then
      m_MasterRef.AddEditMode = ShowMode
      m_MasterRef.MASTER_AREA = BANK_ACCOUNT
      m_MasterRef.KEY_NAME = txtName.Text
      m_MasterRef.KEY_CODE = txtCode.Text
      m_MasterRef.TEMP_ID1 = cboBank.ItemData(Minus2Zero(cboBank.ListIndex))
      m_MasterRef.TEMP_ID2 = cboBankBranch.ItemData(Minus2Zero(cboBankBranch.ListIndex))
      Call glbMaster.AddEditMasterRef(m_MasterRef, IsOK, True, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 8-0" Then
      m_ParameterProcess.AddEditMode = ShowMode
      m_ParameterProcess.PARAMETER_PROCESS_NAME = txtName.Text
      m_ParameterProcess.PARAMETER_PROCESS_NO = txtCode.Text
      Call glbMaster.AddEditParameterProcess(m_ParameterProcess, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 8-1" Then
'      m_Process.AddEditMode = ShowMode
'      m_Process.PROCESS_NAME = txtName.Text
'      m_Process.PROCESS_NO = txtCode.Text
'      Call glbMaster.AddEditProcess(m_Process, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 8-2" Then
      m_FormulaType.AddEditMode = ShowMode
      m_FormulaType.FORMULA_TYPE_NAME = txtName.Text
      m_FormulaType.FORMULA_TYPE_NO = txtCode.Text
'      m_FormulaType.INTERMEDIAT_FLAG = Check2Flag(chkFlag.Value)
      Call glbMaster.AddEditFormulaType(m_FormulaType, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 8-3" Then
      m_Machine.AddEditMode = ShowMode
      m_Machine.MACHINE_NAME = txtName.Text
      m_Machine.MACHINE_NO = txtCode.Text
      Call glbMaster.AddEditMachine(m_Machine, IsOK, glbErrorLog)
   End If
   
   IsOK = True
   Call EnableForm(Me, True)
   
   If Not IsOK Then
      Call EnableForm(Me, True)
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   Call EnableForm(Me, True)
   SaveData = True
   Exit Function
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
   Call EnableForm(Me, True)
   SaveData = False
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
      End If
      
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
      Call cmdSave_Click
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

Private Sub Form_Load()
   
   OKClick = False
   Call InitFormLayout
   
   m_HasActivate = False
   m_HasModify = False
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

Private Sub SSCommand2_Click()

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
   Set m_Money_family = Nothing
   Set m_ParameterProcess = Nothing
   Set m_Bank = Nothing
   Set m_BankBranch = Nothing
   Set m_Packaging = Nothing
   Set m_PurchaseExpense = Nothing
   Set m_MasterRef = Nothing
End Sub

Private Sub txtDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtCode_Change()
   m_HasModify = True
End Sub

Private Sub txtThaiMsg_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox1_Change()
   m_HasModify = True
End Sub

Private Sub txtName_Change()
   m_HasModify = True
End Sub

Private Sub txtWeightRate_Change()
   m_HasModify = True
End Sub
