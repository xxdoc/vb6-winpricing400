VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditJobInput 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8880
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditJobInput.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   6225
      Left            =   0
      TabIndex        =   16
      Top             =   600
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   10980
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTime uctlMixDate 
         Height          =   375
         Left            =   1800
         TabIndex        =   10
         Top             =   4320
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
      End
      Begin prjFarmManagement.uctlTextLookup uctlProductLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   1
         Top             =   720
         Width           =   6825
         _ExtentX        =   12039
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtAmount 
         Height          =   435
         Left            =   1800
         TabIndex        =   3
         Top             =   1560
         Width           =   1485
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPlaceLookup 
         Height          =   405
         Left            =   1800
         TabIndex        =   2
         Top             =   1140
         Width           =   6825
         _ExtentX        =   12039
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtRef 
         Height          =   435
         Left            =   1800
         TabIndex        =   6
         Top             =   2520
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtSerialNo 
         Height          =   435
         Left            =   1800
         TabIndex        =   5
         Top             =   2070
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPartTypeLookup 
         Height          =   465
         Left            =   1800
         TabIndex        =   0
         Top             =   270
         Width           =   6825
         _ExtentX        =   12039
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextLookup uctlFromFormulaLookup 
         Height          =   405
         Left            =   1800
         TabIndex        =   7
         Top             =   2970
         Width           =   6825
         _ExtentX        =   12039
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtAvgPrice 
         Height          =   435
         Left            =   1800
         TabIndex        =   8
         Top             =   3420
         Width           =   1485
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtGroupNo 
         Height          =   435
         Left            =   1800
         TabIndex        =   9
         Top             =   3870
         Width           =   1485
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtStdAmount 
         Height          =   435
         Left            =   6540
         TabIndex        =   4
         Top             =   1560
         Width           =   1485
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlCostTypeLookup 
         Height          =   405
         Left            =   1800
         TabIndex        =   11
         Top             =   4740
         Width           =   6825
         _ExtentX        =   12039
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtPackAmount 
         Height          =   435
         Left            =   6885
         TabIndex        =   32
         Top             =   2040
         Visible         =   0   'False
         Width           =   1125
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin VB.Label lblPackAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPackAmount"
         Height          =   375
         Left            =   5280
         TabIndex        =   31
         Top             =   2100
         Visible         =   0   'False
         Width           =   1545
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   405
         Left            =   8160
         TabIndex        =   30
         Top             =   1560
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJobInput.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdLotSelect 
         Height          =   405
         Left            =   3360
         TabIndex        =   29
         Top             =   1560
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJobInput.frx":0BE4
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   2160
         TabIndex        =   12
         Top             =   5460
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblCostType 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPlace"
         Height          =   315
         Left            =   240
         TabIndex        =   28
         Top             =   4770
         Width           =   1455
      End
      Begin VB.Label lblStdAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblAmount"
         Height          =   375
         Left            =   3960
         TabIndex        =   27
         Top             =   1590
         Width           =   2505
      End
      Begin VB.Label lblMixDate 
         Alignment       =   1  'Right Justify
         Caption         =   "lblAmount"
         Height          =   375
         Left            =   480
         TabIndex        =   26
         Top             =   4320
         Width           =   1245
      End
      Begin VB.Label lblGroupNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblAmount"
         Height          =   375
         Left            =   480
         TabIndex        =   25
         Top             =   3930
         Width           =   1245
      End
      Begin VB.Label lblAvgPrice 
         Alignment       =   1  'Right Justify
         Caption         =   "lblAmount"
         Height          =   375
         Left            =   480
         TabIndex        =   24
         Top             =   3480
         Width           =   1245
      End
      Begin VB.Label lblFromFormula 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPlace"
         Height          =   315
         Left            =   240
         TabIndex        =   23
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label lblSerialNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblSerialNo"
         Height          =   375
         Left            =   60
         TabIndex        =   22
         Top             =   2100
         Width           =   1665
      End
      Begin VB.Label lblRef 
         Alignment       =   1  'Right Justify
         Caption         =   "lblRef"
         Height          =   345
         Left            =   30
         TabIndex        =   21
         Top             =   2550
         Width           =   1695
      End
      Begin VB.Label lblPlace 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPlace"
         Height          =   315
         Left            =   240
         TabIndex        =   20
         Top             =   1170
         Width           =   1455
      End
      Begin VB.Label lblAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblAmount"
         Height          =   375
         Left            =   0
         TabIndex        =   19
         Top             =   1590
         Width           =   1725
      End
      Begin VB.Label lblType 
         Alignment       =   1  'Right Justify
         Caption         =   "lblType"
         Height          =   315
         Left            =   270
         TabIndex        =   18
         Top             =   270
         Width           =   1455
      End
      Begin VB.Label lblProduct 
         Alignment       =   1  'Right Justify
         Caption         =   "lblProduct"
         Height          =   315
         Left            =   270
         TabIndex        =   17
         Top             =   720
         Width           =   1455
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3810
         TabIndex        =   13
         Top             =   5460
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJobInput.frx":0EFE
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5460
         TabIndex        =   14
         Top             =   5460
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditJobInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Header As String
Public ShowMode As SHOW_MODE_TYPE
Public ParentShowMode As SHOW_MODE_TYPE
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Input_combo As Collection
Private m_Input1_combo As Collection
Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public TempCollection As Collection
Public COMMIT_FLAG As String
Public ParentForm As Form
Public HaveData As Long
Public Area As Long
Public ProcessID As Long
Private m_PartTypes As Collection
Private m_PartItems As Collection
Private m_Locations As Collection
Private m_Formulas As Collection
Private m_CostType As Collection
Public PartType As Long
Private WeightPerPack As Long
Public m_InventoryWhDocInput As CInventoryWHDoc
Public t_InventoryWhDocInput As CInventoryWHDoc
Public DOCUMENT_TYPE As Long
Public TYPE_LIST_RM As Long
Public DOCUMENT_DATE As Date
Public typeForm As Long
Private LIW As CLotItemWH
Private LTD As CLotDoc
Public LOCATION_ID As Long
Dim tempLW As CLotItemWH

Private Sub cmdEdit_Click()
   If CheckIwdAmount > 0 Then
      txtAmount.Enabled = True
   End If
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub InitFormLayout()
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)

   Me.KeyPreview = True
   pnlHeader.Caption = HeaderText
   Me.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_HEAD_COLOR
   SSFrame1.BackColor = GLB_FORM_COLOR
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   
   If DOCUMENT_TYPE = 14 And TYPE_LIST_RM = 1 Then
      Call InitNormalLabel(lblType, MapText("ประเภทถุงบรรจุ"))
      Call InitNormalLabel(lblProduct, MapText("ถุงบรรจุ"))
      Call InitNormalLabel(lblAmount, MapText("จำนวนถุงใช้จริง"))
      Call InitNormalLabel(lblStdAmount, MapText("จำนวนถุงมาตรฐาน"))
   ElseIf DOCUMENT_TYPE = 18 And TYPE_LIST_RM = 5 Then
      Call InitNormalLabel(lblType, MapText("ประเภทอาหาร"))
      Call InitNormalLabel(lblProduct, MapText("เบอร์อาหาร"))
      Call InitNormalLabel(lblAmount, MapText("น้ำหนักใช้จริง"))
      Call InitNormalLabel(lblStdAmount, MapText("น้ำหนักมาตรฐาน"))
   Else
     Call InitNormalLabel(lblType, MapText("ประเภทวัตถุดิบ"))
      Call InitNormalLabel(lblProduct, MapText("วัตถุดิบ"))
      Call InitNormalLabel(lblAmount, MapText("น้ำหนักใช้จริง"))
      Call InitNormalLabel(lblStdAmount, MapText("น้ำหนักมาตรฐาน"))
   End If
   
   Call InitNormalLabel(lblPlace, MapText("สถานที่จัดเก็บ"))
   Call InitNormalLabel(lblSerialNo, MapText("ซีเรียล"))
   Call InitNormalLabel(lblRef, MapText("หมายเลขอ้างอิง"))
   Call InitNormalLabel(lblFromFormula, MapText("จากสูตร"))
   Call InitNormalLabel(lblAvgPrice, MapText("ราคาเฉลี่ย"))
   Call InitNormalLabel(lblMixDate, MapText("เวลาผสม"))
   Call InitNormalLabel(lblGroupNo, MapText("กลุ่ม"))
   
   Call InitNormalLabel(lblCostType, MapText("ต้นทุนผลิต"))
   Call InitNormalLabel(lblPackAmount, MapText("จำนวนถุง"))
      
   lblPackAmount.Enabled = False
   
   Call txtAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   Call txtSerialNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtRef.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtAvgPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtAvgPrice.Enabled = False
   Call txtGroupNo.SetTextLenType(TEXT_INTEGER, glbSetting.MONEY_TYPE)
   Call txtStdAmount.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtPackAmount.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   txtPackAmount.Enabled = False
   
   Call uctlProductLookup.MyTextBox.SetKeySearch("PART_NO")
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdLotSelect.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdNext, MapText("ถัดไป"))
   Call InitMainButton(cmdLotSelect, MapText("..."))
   Call InitMainButton(cmdEdit, MapText("E"))
   
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Ma As CJobInput
Dim MaWarehouse As CJobInputWarehouse
If Area = 2 Then
      
      If Flag Then
         Call EnableForm(Me, False)
         
         If ShowMode = SHOW_EDIT Then
            Set MaWarehouse = TempCollection.Item(ID)
           uctlPartTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPartTypeLookup.MyCombo, MaWarehouse.PART_TYPE_ID)
           uctlProductLookup.MyCombo.ListIndex = IDToListIndex(uctlProductLookup.MyCombo, MaWarehouse.PART_ITEM_ID)
           txtAmount.Text = MaWarehouse.TX_AMOUNT
           uctlPlaceLookup.MyCombo.ListIndex = IDToListIndex(uctlPlaceLookup.MyCombo, MaWarehouse.LOCATION_ID)
           uctlFromFormulaLookup.MyCombo.ListIndex = IDToListIndex(uctlFromFormulaLookup.MyCombo, MaWarehouse.FROM_FORMULA)
           txtSerialNo.Text = MaWarehouse.SERIAL_NUMBER
           txtRef.Text = MaWarehouse.INOUT_REF
           txtAvgPrice.Text = MaWarehouse.INCLUDE_UNIT_PRICE
           txtGroupNo.Text = MaWarehouse.GROUP_NO
            uctlMixDate.HR = HOUR(MaWarehouse.MIX_DATE)
            uctlMixDate.MI = Minute(MaWarehouse.MIX_DATE)
           txtStdAmount.Text = MaWarehouse.STD_AMOUNT
           uctlCostTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlCostTypeLookup.MyCombo, MaWarehouse.PARAM_ID)
           
           cmdOK.Enabled = (COMMIT_FLAG <> "Y")
         ElseIf HaveData > 0 And ShowMode = SHOW_ADD Then
            Set MaWarehouse = TempCollection.Item(HaveData)
            uctlPartTypeLookup.MyCombo.ListIndex = -1
            uctlProductLookup.MyCombo.ListIndex = -1
            txtAmount.Text = ""
            uctlPlaceLookup.MyCombo.ListIndex = -1
            uctlFromFormulaLookup.MyCombo.ListIndex = -1
            txtSerialNo.Text = ""
            txtRef.Text = ""
            txtAvgPrice.Text = ""
            txtGroupNo.Text = "0"
             uctlMixDate.HR = HOUR(Now)
             uctlMixDate.MI = Minute(Now)
            txtStdAmount.Text = ""
            uctlCostTypeLookup.MyCombo.ListIndex = -1
            cmdOK.Enabled = (COMMIT_FLAG <> "Y")
         End If
      End If
Else
      If Flag Then
         Call EnableForm(Me, False)
         
         If ShowMode = SHOW_EDIT Then
            Set Ma = TempCollection.Item(ID)
           uctlPartTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPartTypeLookup.MyCombo, Ma.PART_TYPE_ID)
            If (ProcessID = 2 And Ma.PART_TYPE_ID = 21) Or (ProcessID = 6 And Ma.PART_TYPE_ID = 10) Or (ProcessID = 7 And Ma.PART_TYPE_ID = 10) Or (ProcessID = 8 And Ma.PART_TYPE_ID = 10) Then
               PartType = Ma.PART_TYPE_ID
               uctlPartTypeLookup.Enabled = False
               cmdLotSelect.Visible = True
'               cmdEdit.Visible = True
               txtPackAmount.Visible = True
               lblPackAmount.Visible = True
               txtAmount.Enabled = False
               txtStdAmount.Enabled = False
               txtPackAmount.Enabled = False
               lblPackAmount.Enabled = False
            End If
            cmdEdit.Visible = False
            If Not ((ProcessID = 6 And Ma.PART_TYPE_ID = 10) Or (ProcessID = 7 And Ma.PART_TYPE_ID = 10) Or (ProcessID = 8 And Ma.PART_TYPE_ID = 10)) Then
               cmdEdit.Visible = True
            End If
           uctlProductLookup.MyCombo.ListIndex = IDToListIndex(uctlProductLookup.MyCombo, Ma.PART_ITEM_ID)
           txtAmount.Text = Ma.TX_AMOUNT
           uctlPlaceLookup.MyCombo.ListIndex = IDToListIndex(uctlPlaceLookup.MyCombo, Ma.LOCATION_ID)
           uctlFromFormulaLookup.MyCombo.ListIndex = IDToListIndex(uctlFromFormulaLookup.MyCombo, Ma.FROM_FORMULA)
           txtSerialNo.Text = Ma.SERIAL_NUMBER
           txtRef.Text = Ma.INOUT_REF
           txtAvgPrice.Text = Ma.INCLUDE_UNIT_PRICE
           txtGroupNo.Text = Ma.GROUP_NO
           uctlMixDate.HR = HOUR(Ma.MIX_DATE)
           uctlMixDate.MI = Minute(Ma.MIX_DATE)
           txtStdAmount.Text = Ma.STD_AMOUNT
           txtPackAmount.Text = MyDiff(Ma.TX_AMOUNT, WeightPerPack)
           uctlCostTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlCostTypeLookup.MyCombo, Ma.PARAM_ID)
           
           cmdOK.Enabled = (COMMIT_FLAG <> "Y")
         ElseIf HaveData > 0 And ShowMode = SHOW_ADD Then
            Set Ma = TempCollection.Item(HaveData)
            uctlPartTypeLookup.MyCombo.ListIndex = -1
            uctlProductLookup.MyCombo.ListIndex = -1
            txtAmount.Text = ""
            uctlPlaceLookup.MyCombo.ListIndex = -1
            uctlFromFormulaLookup.MyCombo.ListIndex = -1
            txtSerialNo.Text = ""
            txtRef.Text = ""
            txtAvgPrice.Text = ""
            txtGroupNo.Text = "0"
             uctlMixDate.HR = HOUR(Now)
             uctlMixDate.MI = Minute(Now)
            txtStdAmount.Text = ""
            uctlCostTypeLookup.MyCombo.ListIndex = -1
            cmdOK.Enabled = (COMMIT_FLAG <> "Y")
         Else
                
         End If
      End If
      
   End If
   
   Call EnableForm(Me, True)
End Sub
Private Sub cmdLotSelect_Click()
Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
Dim OKClick As Boolean
Dim LotAmount As Long
Dim I As Long
Dim t_Liw As CLotItemWH


   If Not VerifyCombo(lblType, uctlPartTypeLookup.MyCombo, False) Then
      Exit Sub
   End If
   If Not VerifyCombo(lblProduct, uctlProductLookup.MyCombo, False) Then
      Exit Sub
   End If
   
   If Not VerifyCombo(lblPlace, uctlPlaceLookup.MyCombo, False) Then
      Exit Sub
   End If

 If Not m_InventoryWhDocInput Is Nothing Then
   frmAddEditTrnGoods.ShowMode = m_InventoryWhDocInput.AddEditMode
 End If
 
 
Set oMenu = New cPopupMenu
'lMenuChosen = oMenu.Popup("เพิ่ม", "-", "แก้ไข", "-", "ลบ")
lMenuChosen = oMenu.Popup("เพิ่ม", "-", "แก้ไข")
If lMenuChosen = 0 Then
   Exit Sub
End If

If lMenuChosen = 0 Then
   Exit Sub
End If

   Set LIW = New CLotItemWH
   LIW.PART_ITEM_ID = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
   LIW.LOCATION_ID = uctlPlaceLookup.MyCombo.ItemData(Minus2Zero(uctlPlaceLookup.MyCombo.ListIndex))
   LIW.PART_NO = uctlProductLookup.MyTextBox.Text
   LIW.PART_DESC = uctlProductLookup.MyCombo.Text
   LIW.WEIGHT_PER_PACK = WeightPerPack
   
If lMenuChosen = 1 Then
   frmAddEditTrnGoods.ShowMode = SHOW_ADD

   If CountItem(m_InventoryWhDocInput.C_LotItemsWH) = 0 Then 'ให้มีแค่ lot item เดียว
        Set m_InventoryWhDocInput.C_LotItemsWH = Nothing
      Set m_InventoryWhDocInput.C_LotItemsWH = New Collection
      LIW.AddEditMode = SHOW_ADD
      LIW.Flag = "A"
      Call m_InventoryWhDocInput.C_LotItemsWH.add(LIW, str(LIW.PART_ITEM_ID))
   Else
      For Each t_Liw In m_InventoryWhDocInput.C_LotItemsWH
         If LIW.PART_ITEM_ID <> t_Liw.PART_ITEM_ID Then
                  Set m_InventoryWhDocInput.C_LotItemsWH = Nothing
                  Set m_InventoryWhDocInput.C_LotItemsWH = New Collection
                  LIW.AddEditMode = SHOW_ADD
                  LIW.Flag = "A"
                  Call m_InventoryWhDocInput.C_LotItemsWH.add(LIW, str(LIW.PART_ITEM_ID))
           Else
                  If t_Liw.Flag <> "A" Then
'                     frmAddEditTrnGoods.ShowMode = SHOW_EDIT
                     t_Liw.AddEditMode = SHOW_EDIT
                     t_Liw.Flag = "E"
                  End If
                  frmAddEditTrnGoods.ShowMode = SHOW_EDIT
         End If
      Next t_Liw
   End If

   frmAddEditTrnGoods.DOCUMENT_TYPE = DOCUMENT_TYPE '17,19
   frmAddEditTrnGoods.ProcessID = ProcessID
   
   frmAddEditTrnGoods.ID = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
    Set frmAddEditTrnGoods.m_InventoryWHDoc = m_InventoryWhDocInput
   Set frmAddEditTrnGoods.ParentForm = Me
   frmAddEditTrnGoods.DOCUMENT_DATE = DOCUMENT_DATE
   Load frmAddEditTrnGoods
   frmAddEditTrnGoods.Show 1
   
   OKClick = frmAddEditTrnGoods.OKClick
   
   Unload frmAddEditTrnGoods
   Set frmAddEditTrnGoods = Nothing
   
   If OKClick Then
      m_HasModify = True
   End If
ElseIf lMenuChosen = 3 Then
   If m_InventoryWhDocInput.C_LotItemsWH.Count = 0 Then
      glbErrorLog.LocalErrorMsg = "ยังไม่มีข้อมูล"
      glbErrorLog.ShowUserError
      Exit Sub
   End If

   For Each t_Liw In m_InventoryWhDocInput.C_LotItemsWH
            If t_Liw.Flag <> "A" Then
               t_Liw.Flag = "E"
            End If
   Next t_Liw

   frmAddEditTrnGoods.DOCUMENT_TYPE = DOCUMENT_TYPE '17,19
   frmAddEditTrnGoods.ProcessID = ProcessID
   frmAddEditTrnGoods.ShowMode = SHOW_EDIT
   frmAddEditTrnGoods.ID = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
   frmAddEditTrnGoods.DOCUMENT_DATE = DOCUMENT_DATE
    Set frmAddEditTrnGoods.m_InventoryWHDoc = m_InventoryWhDocInput
   Set frmAddEditTrnGoods.ParentForm = Me
   Load frmAddEditTrnGoods
   frmAddEditTrnGoods.Show 1
   
   OKClick = frmAddEditTrnGoods.OKClick
   
   Unload frmAddEditTrnGoods
   Set frmAddEditTrnGoods = Nothing
   
   If OKClick Then
      m_HasModify = True
   End If
ElseIf lMenuChosen = 5 Then
   Dim IWHD As CInventoryWHDoc
   Dim LIW2 As CLotItemWH
   Dim LTD As CLotDoc
   
   If m_InventoryWhDocInput.C_LotItemsWH.Count = 0 Then
      glbErrorLog.LocalErrorMsg = "ยังไม่มีข้อมูล"
      glbErrorLog.ShowUserError
      Exit Sub
   End If

   For Each LIW2 In m_InventoryWhDocInput.C_LotItemsWH
      For Each LTD In LIW2.C_LotDoc
         LTD.Flag = "D" 'ให้ลบถึง lotdoc เพราะต้องให้เครียร์ยอด update ด้วย
      Next LTD
'      LIW2.Flag = "D"
   Next LIW2
   
    txtPackAmount.Text = ""
    txtAmount.Text = "0"
    txtStdAmount.Text = "0"
End If


      
      
 
''I = 0
''Set t_InventoryWhDocInput = New CInventoryWHDoc
'' For Each LIW In m_InventoryWhDocInput.C_LotItemsWH
''   If LIW.Flag = "I" Then
''       LIW.Flag2 = LIW.Flag
''   End If
''   Call t_InventoryWhDocInput.C_LotItemsWH.add(LIW, str(LIW.PART_ITEM_ID))
'' Next LIW
''For Each LIW In m_InventoryWhDocInput.C_LotItemsWH
''  I = 1
''   Call m_InventoryWhDocInput.C_LotItemsWH.Remove(I)
'' Next LIW
''For Each LIW In t_InventoryWhDocInput.C_LotItemsWH
''    Call m_InventoryWhDocInput.C_LotItemsWH.add(LIW, str(LIW.PART_ITEM_ID))
'' Next LIW
''
''  Set LIW = GetObject("CLotItemWH", m_InventoryWhDocInput.C_LotItemsWH, str(uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))))
'' If Not LIW Is Nothing Then
''   LIW.LOCATION_ID = uctlPlaceLookup.MyCombo.ItemData(Minus2Zero(uctlPlaceLookup.MyCombo.ListIndex))
''   LIW.PART_NO = uctlProductLookup.MyTextBox.Text
''   LIW.PART_DESC = uctlProductLookup.MyCombo.Text
''   LIW.WEIGHT_PER_PACK = WeightPerPack
''
''   If LIW.Flag2 = "I" Then 'ถ้าเป็น ข้อมูลที่เคยถูกบันทึกไปแล้ว และ แก้ใหม่เป็นไม่เลือก แล้วแก้ใหม่อีกอยากกลับมาเลือก
''      LIW.Flag = "E"
''      For Each LTD In LIW.C_LotDoc
''         If LTD.Flag = "D" Then
''            LTD.Flag = "E"
''         End If
''      Next LTD
''   End If
''
''   If LIW.Flag <> "A" Then
''      LIW.Flag = "E"
''   End If
''
''
'''   If LIW.Flag = "D" Then 'ถ้าเป็น ข้อมูลที่เคยถูกบันทึกไปแล้ว และ แก้ใหม่เป็นไม่เลือก แล้วแก้ใหม่อีกอยากกลับมาเลือก
'''      LIW.Flag = "I"
'''      For Each LTD In LIW.C_LotDoc
'''         If LTD.Flag = "D" Then
'''            LTD.Flag = "I"
'''         End If
'''      Next LTD
'''   End If
'''   If LIW.Flag <> "A" And LIW.Flag <> "I" Then
'''      LIW.Flag = "E"
'''   End If
'' Else
''   Set LIW = New CLotItemWH
''   LIW.PART_ITEM_ID = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
''   LIW.LOCATION_ID = uctlPlaceLookup.MyCombo.ItemData(Minus2Zero(uctlPlaceLookup.MyCombo.ListIndex))
''   LIW.PART_NO = uctlProductLookup.MyTextBox.Text
''   LIW.PART_DESC = uctlProductLookup.MyCombo.Text
''   LIW.WEIGHT_PER_PACK = WeightPerPack
''   LIW.Flag = "A"
''   Call m_InventoryWhDocInput.C_LotItemsWH.add(LIW, str(LIW.PART_ITEM_ID))
'' End If
''' Set LIW = GetObject("CLotItemWH", m_InventoryWhDocInput.C_LotItemsWH, str(uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))))
''' If Not LIW Is Nothing Then
'''   LIW.LOCATION_ID = uctlPlaceLookup.MyCombo.ItemData(Minus2Zero(uctlPlaceLookup.MyCombo.ListIndex))
'''   LIW.PART_NO = uctlProductLookup.MyTextBox.Text
'''   LIW.PART_DESC = uctlProductLookup.MyCombo.Text
'''   LIW.WEIGHT_PER_PACK = WeightPerPack
'''   If LIW.Flag = "D" Then 'ถ้าเป็น ข้อมูลที่เคยถูกบันทึกไปแล้ว และ แก้ใหม่เป็นไม่เลือก แล้วแก้ใหม่อีกอยากกลับมาเลือก
'''      LIW.Flag = "I"
'''      For Each LTD In LIW.C_LotDoc
'''         If LTD.Flag = "D" Then
'''            LTD.Flag = "I"
'''         End If
'''      Next LTD
'''   End If
'''   If LIW.Flag <> "A" And LIW.Flag <> "I" Then
'''      LIW.Flag = "E"
'''   End If
''' Else
'''   Set LIW = New CLotItemWH
'''   LIW.PART_ITEM_ID = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
'''   LIW.LOCATION_ID = uctlPlaceLookup.MyCombo.ItemData(Minus2Zero(uctlPlaceLookup.MyCombo.ListIndex))
'''   LIW.PART_NO = uctlProductLookup.MyTextBox.Text
'''   LIW.PART_DESC = uctlProductLookup.MyCombo.Text
'''   LIW.WEIGHT_PER_PACK = WeightPerPack
'''   LIW.Flag = "A"
'''   Call m_InventoryWhDocInput.C_LotItemsWH.add(LIW, str(LIW.PART_ITEM_ID))
''' End If
''
'''   frmAddEditTrnGoods.ID = 1
''   frmAddEditTrnGoods.DOCUMENT_TYPE = DOCUMENT_TYPE '17,19
''   frmAddEditTrnGoods.ProcessID = ProcessID
''   frmAddEditTrnGoods.id = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
''    Set frmAddEditTrnGoods.m_InventoryWHDoc = m_InventoryWhDocInput
''   Set frmAddEditTrnGoods.ParentForm = Me
''   Load frmAddEditTrnGoods
''   frmAddEditTrnGoods.Show 1
''
''   OKClick = frmAddEditTrnGoods.OKClick
''
''   Unload frmAddEditTrnGoods
''   Set frmAddEditTrnGoods = Nothing
''
''   If OKClick Then
''      m_HasModify = True
''   End If
   
   Set oMenu = Nothing
End Sub
Function CheckIwdAmount() As Double
Dim IWD As CInventoryWHDoc
Dim LIW As CLotItemWH
Dim LTD As CLotDoc
Dim PD As CPalletDoc
Dim Sum As Double
Dim cExit As Boolean

cExit = True
If Not m_InventoryWhDocInput Is Nothing Then
  Set IWD = m_InventoryWhDocInput
  cExit = False
End If

If cExit Then
   CheckIwdAmount = -1
   Exit Function
End If
For Each LIW In IWD.C_LotItemsWH
      For Each LTD In LIW.C_LotDoc
        If LTD.Flag <> "D" Then
            For Each PD In LTD.C_PalletDoc
               Sum = Sum + PD.CAPACITY_AMOUNT
            Next PD
         End If
      Next LTD
Next LIW

CheckIwdAmount = Sum
End Function
Private Sub cmdNext_Click()
Dim NewID As Long
   m_HasModify = True
   If Not SaveData Then
      Exit Sub
   End If

   If ShowMode = SHOW_EDIT Then
      NewID = GetNextID(ID, TempCollection)
      If ID = NewID Then
         glbErrorLog.LocalErrorMsg = "ถึงเรคคอร์ดสุดท้ายแล้ว"
         glbErrorLog.ShowUserError

         ParentForm.GridEX1.ItemCount = CountItem(TempCollection)
         Call ParentForm.GridEX1.Rebind
         Exit Sub
      End If

      ID = NewID
   ElseIf ShowMode = SHOW_ADD Then
      HaveData = TempCollection.Count
   End If

   ParentForm.GridEX1.ItemCount = CountItem(TempCollection)
   Call ParentForm.GridEX1.Rebind
   
   Call QueryData(True)

'   Call uctlPartTypeLookup.SetFocus

   m_HasModify = False
End Sub
Private Sub cmdOK_Click()
   If Not SaveData Then
        Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long

   
   If Not VerifyCombo(lblType, uctlPartTypeLookup.MyCombo, False) Then
      Exit Function
   End If

   If Not VerifyCombo(lblProduct, uctlProductLookup.MyCombo, False) Then
      Exit Function
   End If
     
   If Not VerifyTextControl(lblAmount, txtAmount, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblStdAmount, txtStdAmount, False) Then
      Exit Function
   End If
   
  If Not VerifyCombo(lblPlace, uctlPlaceLookup.MyCombo, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblGroupNo, txtGroupNo, False) Then
      Exit Function
   End If
   
''''   If Not (m_InventoryWhDocInput Is Nothing) Then
''''      If CountItem(m_InventoryWhDocInput.C_LotItemsWH) > 0 Then
''''         For Each LIW In m_InventoryWhDocInput.C_LotItemsWH
''''            If LIW.PART_ITEM_ID <> uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex)) Then 'ถ้าเลือกเบอร์อาหารเบอร์ใหม่ก็ให้เครียร์  Lot เดิมด้วย
''''               If LIW.Flag2 = "I" Then 'ถ้าเป็น lot ที่โดน Query มา ให้ทำการ ลบ lot นั้นออก
''''                  LIW.Flag = "D"
''''                  For Each LTD In LIW.C_LotDoc
''''                     LTD.Flag = "D"
''''                  Next LTD
''''               Else 'ถ้าเป็น Lot ที่เพิ่งเลือกทั่วไปก็ ให้เปลี่ยน Flag เป็นว่าง ไม่ต้องทำอะไร
''''                  LIW.Flag = ""
''''               End If
''''
''''   '           If LIW.Flag = "I" Or LIW.Flag = "D" Then  'ถ้าเป็น Lot ที่เคยถูกบันทึกแล้ว ให้เปลี่ยน Flag เป็น D
''''   '             LIW.Flag = "D"
''''   '            For Each LTD In LIW.C_LotDoc
''''   '               LTD.Flag = "D"
''''   '            Next LTD
''''   '           Else 'ถ้าเป็น Lot ที่เพิ่งเลือกทั่วไปก็ ให้เปลี่ยน Flag เป็นว่าง ไม่ต้องทำอะไร
''''   '            LIW.Flag = ""
''''   '          End If
''''   '               glbErrorLog.LocalErrorMsg = MapText("กรุณาระบุ LOT ที่ตัดจ่ายของอาหารเบอร์ " & uctlProductLookup.MyTextBox.Text)
''''   '               glbErrorLog.ShowUserError
''''   '               Exit Function
''''               Else 'ถ้าเป็น เบอร์ที่เลือก ให้เข้าทำ
''''               If LIW.Flag <> "A" Then
''''                  LIW.Flag = "E"
''''               ElseIf LIW.Flag2 = "I" Then
''''                  LIW.Flag = "E"
''''                  For Each LTD In LIW.C_LotDoc
''''                     If LTD.Flag <> "A" Then
''''                        LTD.Flag = "E"
''''                     End If
''''                  Next LTD
''''               End If
''''            End If
''''         Next LIW
''''      End If
''''   End If

   
 If ((ProcessID = 2 And DOCUMENT_TYPE = 18) Or (ProcessID = 6 And DOCUMENT_TYPE = 17) Or (ProcessID = 7 And DOCUMENT_TYPE = 18) Or (ProcessID = 7 And DOCUMENT_TYPE = 17) Or (ProcessID = 8 And DOCUMENT_TYPE = 19)) And typeForm <> 1 Then
Dim dataCheck As Double
Dim PartType As Long
PartType = uctlPartTypeLookup.MyTextBox.Text
   If PartType = 22 Or PartType = 10 Then
      dataCheck = CheckIwdAmount
      If dataCheck > 0 Then
         If Val(txtAmount.Text) < Val(dataCheck) Then
            glbErrorLog.LocalErrorMsg = MapText("จำนวนใช้จริงต้องมากกว่าหรือเท่ากับจำนวนที่จ่ายออกจากโกดังอาหาร")
            glbErrorLog.ShowUserError
            Exit Function
         End If
      End If
   End If
End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim Ma As CJobInput
   If ShowMode = SHOW_ADD Then
      Set Ma = New CJobInput
   Else
      Set Ma = TempCollection.Item(ID)
   End If
   
   Ma.PART_ITEM_ID = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
   Ma.PART_DESC = uctlProductLookup.MyCombo.Text
   Ma.PART_NO = uctlProductLookup.MyTextBox.Text
   Ma.PART_TYPE_ID = uctlPartTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPartTypeLookup.MyCombo.ListIndex))
   Ma.PART_TYPE_NAME = uctlPartTypeLookup.MyCombo.Text
   Ma.TX_AMOUNT = txtAmount.Text
   Ma.LOCATION_ID = uctlPlaceLookup.MyCombo.ItemData(Minus2Zero(uctlPlaceLookup.MyCombo.ListIndex))
   Ma.LOCATION_NO = uctlPlaceLookup.MyTextBox.Text
   Ma.LOCATION_NAME = uctlPlaceLookup.MyCombo.Text
   Ma.SERIAL_NUMBER = txtSerialNo.Text
   Ma.INOUT_REF = txtRef.Text
   Ma.FROM_FORMULA = uctlFromFormulaLookup.MyCombo.ItemData(Minus2Zero(uctlFromFormulaLookup.MyCombo.ListIndex))
   Ma.TX_TYPE = "E"
   Ma.AVG_PRICE = Val(txtAvgPrice.Text)
   Ma.GROUP_NO = Val(txtGroupNo.Text)
   Ma.MIX_DATE = Now
   Ma.MIX_DATE = DateAdd("h", uctlMixDate.HR, Ma.MIX_DATE)
   Ma.MIX_DATE = DateAdd("n", uctlMixDate.MI, Ma.MIX_DATE)
   Ma.STD_AMOUNT = Val(txtStdAmount.Text)
   Ma.PARAM_ID = uctlCostTypeLookup.MyCombo.ItemData(Minus2Zero(uctlCostTypeLookup.MyCombo.ListIndex))
   
   If ShowMode = SHOW_ADD Then
      Ma.Flag = "A"
      Call TempCollection.add(Ma)
   Else
      If Ma.Flag <> "A" Then
         Ma.Flag = "E"
      End If
   End If
   
   If (ProcessID = 2 Or ProcessID = 6 Or ProcessID = 7 Or ProcessID = 8) And Not (m_InventoryWhDocInput Is Nothing) Then
   If ShowMode = SHOW_ADD Then
      If ProcessID = 2 And DOCUMENT_TYPE = 18 Then
         m_InventoryWhDocInput.DOCUMENT_TYPE = 2003 'จ่ายออก BULK to Pack bag
      Else
         m_InventoryWhDocInput.DOCUMENT_TYPE = 2000 ''จ่ายออก BAG
      End If
   Else
      If Ma.Flag <> "A" Then
         Ma.Flag = "E"
      End If
   End If
   End If
   
   SaveData = True
End Function
Private Function SaveData2() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long
   If Not VerifyCombo(lblType, uctlPartTypeLookup.MyCombo, False) Then
      Exit Function
   End If

   If Not VerifyCombo(lblProduct, uctlProductLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblStdAmount, txtStdAmount, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblAmount, txtAmount, False) Then
      Exit Function
   End If
   
  If Not VerifyCombo(lblPlace, uctlPlaceLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblGroupNo, txtGroupNo, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData2 = True
      Exit Function
   End If
   
   Dim Ma As CJobInputWarehouse
   If ShowMode = SHOW_ADD Then
      Set Ma = New CJobInputWarehouse
   Else
      Set Ma = TempCollection.Item(ID)
   End If
   
   Ma.PART_ITEM_ID = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
   Ma.PART_DESC = uctlProductLookup.MyCombo.Text
   Ma.PART_NO = uctlProductLookup.MyTextBox.Text
   Ma.PART_TYPE_ID = uctlPartTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPartTypeLookup.MyCombo.ListIndex))
   Ma.PART_TYPE_NAME = uctlPartTypeLookup.MyCombo.Text
   Ma.TX_AMOUNT = txtAmount.Text
   Ma.LOCATION_ID = uctlPlaceLookup.MyCombo.ItemData(Minus2Zero(uctlPlaceLookup.MyCombo.ListIndex))
   Ma.LOCATION_NO = uctlPlaceLookup.MyTextBox.Text
   Ma.LOCATION_NAME = uctlPlaceLookup.MyCombo.Text
   Ma.SERIAL_NUMBER = txtSerialNo.Text
   Ma.INOUT_REF = txtRef.Text
   Ma.FROM_FORMULA = uctlFromFormulaLookup.MyCombo.ItemData(Minus2Zero(uctlFromFormulaLookup.MyCombo.ListIndex))
   Ma.TX_TYPE = "E"
   Ma.AVG_PRICE = Val(txtAvgPrice.Text)
   Ma.GROUP_NO = Val(txtGroupNo.Text)
   Ma.MIX_DATE = Now
   Ma.MIX_DATE = DateAdd("h", uctlMixDate.HR, Ma.MIX_DATE)
   Ma.MIX_DATE = DateAdd("n", uctlMixDate.MI, Ma.MIX_DATE)
   Ma.STD_AMOUNT = Val(txtStdAmount.Text)
   Ma.PARAM_ID = uctlCostTypeLookup.MyCombo.ItemData(Minus2Zero(uctlCostTypeLookup.MyCombo.ListIndex))
   
   If ShowMode = SHOW_ADD Then
      Ma.Flag = "A"
      Call TempCollection.add(Ma)
   Else
      If Ma.Flag <> "A" Then
         Ma.Flag = "E"
      End If
   End If
   
   SaveData2 = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadPartType(uctlPartTypeLookup.MyCombo, m_PartTypes)
      Set uctlPartTypeLookup.MyCollection = m_PartTypes
    
      Call LoadLocation(uctlPlaceLookup.MyCombo, m_Locations, 2)
      Set uctlPlaceLookup.MyCollection = m_Locations

      If ProcessID <> 2 And ProcessID <> 4 And ProcessID <> 6 And ProcessID <> 7 And ProcessID <> 8 Then
         Call LoadFormula(uctlFromFormulaLookup.MyCombo, m_Formulas)
         Set uctlFromFormulaLookup.MyCollection = m_Formulas
     End If

      Call LoadParameterProcess(uctlCostTypeLookup.MyCombo, m_CostType)
      Set uctlCostTypeLookup.MyCollection = m_CostType
       
       txtGroupNo.Text = "0"
       
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
         
          If ((ProcessID = 2 And DOCUMENT_TYPE = 18) Or (ProcessID = 6 And DOCUMENT_TYPE = 17) Or (ProcessID = 7 And DOCUMENT_TYPE = 18) Or (ProcessID = 7 And DOCUMENT_TYPE = 17) Or (ProcessID = 8 And DOCUMENT_TYPE = 19)) And typeForm <> 1 Then
            uctlProductLookup.Enabled = False
            uctlPlaceLookup.Enabled = False
          End If
      ElseIf ShowMode = SHOW_ADD Then
         If ((ProcessID = 2 And DOCUMENT_TYPE = 18) Or (ProcessID = 6 And DOCUMENT_TYPE = 17) Or (ProcessID = 7 And DOCUMENT_TYPE = 18) Or (ProcessID = 7 And DOCUMENT_TYPE = 17) Or (ProcessID = 8 And DOCUMENT_TYPE = 19)) And typeForm <> 1 Then
            uctlPartTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPartTypeLookup.MyCombo, PartType)
            uctlPlaceLookup.MyCombo.ListIndex = IDToListIndex(uctlPlaceLookup.MyCombo, LOCATION_ID)
            uctlPartTypeLookup.Enabled = False
            cmdLotSelect.Visible = True
            cmdEdit.Visible = True
            txtPackAmount.Visible = True
            lblPackAmount.Visible = True
            txtAmount.Enabled = False
            txtStdAmount.Enabled = False
            cmdNext.Enabled = False
         End If
         ID = 0
         uctlMixDate.HR = HOUR(Now)
         uctlMixDate.MI = Minute(Now)
         Call QueryData(False)
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
      Call cmdOK_Click
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
   Set m_Input_combo = New Collection
   Set m_Input1_combo = New Collection
   Set m_Rs = New ADODB.Recordset
   
   Set m_PartTypes = New Collection
   Set m_PartItems = New Collection
   Set m_Locations = New Collection
   Set m_Formulas = New Collection
   Set m_CostType = New Collection

End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_PartTypes = Nothing
   Set m_PartItems = Nothing
   Set m_Locations = Nothing
   Set m_Formulas = Nothing
   Set m_CostType = Nothing

End Sub
Private Sub txtAvgPrice_Change()
   m_HasModify = True
End Sub

Private Sub txtGroupNo_Change()
   m_HasModify = True
End Sub

Private Sub txtStdAmount_Change()
   m_HasModify = True
End Sub

Private Sub uctlCostTypeLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlFromFormulaLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlMixDate_HasChange()
   m_HasModify = True
End Sub
Private Sub uctlPartTypeLookup_Change()
Dim PartTypeID As Long
Dim Pt As CPartType

   PartTypeID = uctlPartTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPartTypeLookup.MyCombo.ListIndex))
      
   If ShowMode = SHOW_ADD Then
      If PartTypeID = 2 And ProcessID = 8 And DOCUMENT_TYPE <> 19 Then
         glbErrorLog.LocalErrorMsg = MapText("การเพิ่มวัตถุดิบที่ใช้เป็น BULK ให้เลือกจากเมนู เพิ่มข้อมูลใหม่ BULK เท่านั้น")
         glbErrorLog.ShowUserError
         Exit Sub
      End If
      
      If PartTypeID = 21 And ProcessID = 2 And DOCUMENT_TYPE <> 18 Then
         glbErrorLog.LocalErrorMsg = MapText("การเพิ่มวัตถุดิบที่ใช้เป็น BULK ให้เลือกจากเมนู เพิ่มข้อมูลใหม่ BULK เท่านั้น")
         glbErrorLog.ShowUserError
         Exit Sub
      End If
      
      If PartTypeID = 10 And ProcessID = 6 And DOCUMENT_TYPE <> 17 Then
         glbErrorLog.LocalErrorMsg = MapText("การเพิ่มวัตถุดิบที่ใช้เป็น BULK ให้เลือกจากเมนู เพิ่มข้อมูลใหม่ RE-BAG -> BAG เท่านั้น")
         glbErrorLog.ShowUserError
         Exit Sub
      End If
      
     If PartTypeID = 10 And ProcessID = 7 And DOCUMENT_TYPE <> 17 Then
         glbErrorLog.LocalErrorMsg = MapText("การเพิ่มวัตถุดิบที่ใช้เป็น BAG ให้เลือกจากเมนู เพิ่มข้อมูลใหม่ RE-BAG -> BAG เท่านั้น")
         glbErrorLog.ShowUserError
         Exit Sub
      End If
   End If
   
   If PartTypeID > 0 Then
      Set Pt = GetPartType(m_PartTypes, Trim(str(PartTypeID)))
      Call LoadPartItem(uctlProductLookup.MyCombo, m_PartItems, PartTypeID, "N", , , "N")
      Set uctlProductLookup.MyCollection = m_PartItems
   
      Call LoadLocation(uctlPlaceLookup.MyCombo, m_Locations, 2, , , Pt.PART_GROUP_ID)
      Set uctlPlaceLookup.MyCollection = m_Locations
   End If
   
   m_HasModify = True
End Sub

Private Sub txtAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtLink_Change()
   m_HasModify = True
End Sub

Private Sub txtRef_Change()
   m_HasModify = True
End Sub

Private Sub txtSerialNo_Change()
   m_HasModify = True
End Sub

Private Sub uctlPlaceLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlProductLookup_Change()
Dim ID As Long


Dim Pi As CPartItem

   ID = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
   If ID <> 0 Then
       If (ProcessID = 6 And PartType = 10) Or (ProcessID = 7 And PartType = 10) Or (ProcessID = 8 And PartType = 10) Then
          Set Pi = GetPartItem(m_PartItems, Trim(str(ID)))
          If Pi.WEIGHT_PER_PACK <= 0 Then
            glbErrorLog.LocalErrorMsg = MapText("สินค้าเบอร์ ") & " " & uctlProductLookup.MyCombo.Text & " " & MapText("ยังไม่ระบุน้ำหนัก กรุณาระบุน้ำหนักที่ ข้อมูลหลักให้เรียบร้อย")
            glbErrorLog.ShowUserError
            Exit Sub
          End If
          WeightPerPack = Pi.WEIGHT_PER_PACK
      End If
     
      Call LoadFormula(uctlFromFormulaLookup.MyCombo, m_Formulas, , ID)
      Set uctlFromFormulaLookup.MyCollection = m_Formulas
   End If
   m_HasModify = True
End Sub

Public Sub setQuantity(Value As Double)
  If (ProcessID = 2 And DOCUMENT_TYPE = 18) Then  'ถ้าเป็น โปรแซส การบรรจุ Bulk'Or (ProcessID = 8 And DOCUMENT_TYPE = 19)
      txtAmount.Text = Value
   Else
      txtAmount.Text = Value * WeightPerPack
      txtPackAmount.Text = ""
      txtPackAmount.Text = Value
   End If
   txtStdAmount.Text = txtAmount.Text
End Sub
