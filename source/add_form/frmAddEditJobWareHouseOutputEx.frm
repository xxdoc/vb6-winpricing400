VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditJobWHOutputEx 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7665
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditJobWareHouseOutputEx.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   7665
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   7515
      Left            =   0
      TabIndex        =   22
      Top             =   600
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   13256
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboLockNo 
         Height          =   510
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   2640
         Width           =   1725
      End
      Begin VB.ComboBox cboPalletNo 
         Height          =   510
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   2160
         Width           =   1725
      End
      Begin VB.ComboBox cboBinNo 
         Height          =   510
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   2160
         Width           =   1725
      End
      Begin prjFarmManagement.uctlDate uctlPackDate 
         Height          =   375
         Left            =   1800
         TabIndex        =   12
         Top             =   4680
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin prjFarmManagement.uctlTime txtTimePackBegin 
         Height          =   435
         Left            =   1800
         TabIndex        =   13
         Top             =   5160
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtWeightPerPack 
         Height          =   435
         Left            =   5520
         TabIndex        =   7
         Top             =   3120
         Width           =   1725
         _ExtentX        =   2990
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlProductLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   1
         Top             =   750
         Width           =   5470
         _ExtentX        =   9657
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtAmount 
         Height          =   435
         Left            =   1800
         TabIndex        =   8
         Top             =   3600
         Width           =   1725
         _ExtentX        =   2619
         _ExtentY        =   556
      End
      Begin prjFarmManagement.uctlTextLookup uctlPlaceLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   15
         Top             =   5640
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtRef 
         Height          =   435
         Left            =   6600
         TabIndex        =   20
         Top             =   7560
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtSerialNo 
         Height          =   435
         Left            =   1320
         TabIndex        =   19
         Top             =   7560
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPartTypeLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   0
         Top             =   300
         Width           =   5470
         _ExtentX        =   9657
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtStdAmount 
         Height          =   435
         Left            =   5520
         TabIndex        =   11
         Top             =   4080
         Width           =   1725
         _ExtentX        =   2619
         _ExtentY        =   979
      End
      Begin prjFarmManagement.uctlTextBox txtPackAmount 
         Height          =   435
         Left            =   1800
         TabIndex        =   10
         Top             =   4050
         Width           =   1725
         _ExtentX        =   2619
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlProductTypeLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   2
         Top             =   1200
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtLotNo 
         Height          =   435
         Left            =   1800
         TabIndex        =   3
         Top             =   1680
         Width           =   1725
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtBathAmount 
         Height          =   435
         Left            =   5520
         TabIndex        =   4
         Top             =   1680
         Width           =   1725
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtRestAmount 
         Height          =   435
         Left            =   5520
         TabIndex        =   9
         Top             =   3600
         Width           =   1725
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTime txtTimePackEnd 
         Height          =   435
         Left            =   6120
         TabIndex        =   14
         Top             =   5160
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtNote 
         Height          =   435
         Left            =   1800
         TabIndex        =   16
         Top             =   6120
         Width           =   5445
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtGoodAmount 
         Height          =   435
         Left            =   5520
         TabIndex        =   5
         Top             =   2640
         Width           =   1725
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtLoseAmount 
         Height          =   435
         Left            =   1800
         TabIndex        =   6
         Top             =   3120
         Width           =   1725
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin VB.Label lblLockNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblLockNo"
         Height          =   345
         Left            =   0
         TabIndex        =   44
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label lblPalletNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPalletNo"
         Height          =   345
         Left            =   3720
         TabIndex        =   43
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label lblLoseAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblLoseAmount"
         Height          =   345
         Left            =   600
         TabIndex        =   42
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label lblGoodAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblGoodAmount"
         Height          =   345
         Left            =   4440
         TabIndex        =   41
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label lblPackDate 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPackDate"
         Height          =   375
         Left            =   -240
         TabIndex        =   40
         Top             =   4680
         Width           =   1905
      End
      Begin VB.Label lblNote 
         Alignment       =   1  'Right Justify
         Caption         =   "lblNote"
         Height          =   375
         Left            =   0
         TabIndex        =   39
         Top             =   6120
         Width           =   1665
      End
      Begin VB.Label lblTimePackEnd 
         Alignment       =   1  'Right Justify
         Caption         =   "lblTimePackEnd"
         Height          =   375
         Left            =   3480
         TabIndex        =   38
         Top             =   5160
         Width           =   1905
      End
      Begin VB.Label lblTimePackBegin 
         Alignment       =   1  'Right Justify
         Caption         =   "lblTimePackBegin"
         Height          =   375
         Left            =   -240
         TabIndex        =   37
         Top             =   5160
         Width           =   1905
      End
      Begin VB.Label lblBinNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblBinNo"
         Height          =   375
         Left            =   0
         TabIndex        =   36
         Top             =   2160
         Width           =   1665
      End
      Begin VB.Label lblRestAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblRestAmount"
         Height          =   375
         Left            =   3720
         TabIndex        =   35
         Top             =   3600
         Width           =   1665
      End
      Begin VB.Label lblWeightPerPack 
         Alignment       =   1  'Right Justify
         Caption         =   "lblWeightPerPack"
         Height          =   375
         Left            =   3720
         TabIndex        =   34
         Top             =   3120
         Width           =   1665
      End
      Begin VB.Label lblBathAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblBathAmount"
         Height          =   345
         Left            =   3720
         TabIndex        =   33
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lblLotNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblLotNo"
         Height          =   345
         Left            =   0
         TabIndex        =   32
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lblProductType 
         Alignment       =   1  'Right Justify
         Caption         =   "lblProductType"
         Height          =   315
         Left            =   240
         TabIndex        =   31
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblPackAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPackAmount"
         Height          =   375
         Left            =   -240
         TabIndex        =   30
         Top             =   4110
         Width           =   1905
      End
      Begin VB.Label lblStdAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblStdAmount"
         Height          =   375
         Left            =   3480
         TabIndex        =   29
         Top             =   4080
         Width           =   1785
      End
      Begin VB.Label lblSerialNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblSerialNo"
         Height          =   375
         Left            =   -120
         TabIndex        =   28
         Top             =   7440
         Width           =   1665
      End
      Begin VB.Label lblRef 
         Alignment       =   1  'Right Justify
         Caption         =   "lblRef"
         Height          =   345
         Left            =   4800
         TabIndex        =   27
         Top             =   7560
         Width           =   1695
      End
      Begin VB.Label lblPlace 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPlace"
         Height          =   315
         Left            =   240
         TabIndex        =   26
         Top             =   5640
         Width           =   1455
      End
      Begin VB.Label lblAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblAmount"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   3600
         Width           =   1545
      End
      Begin VB.Label lblType 
         Alignment       =   1  'Right Justify
         Caption         =   "lblType"
         Height          =   315
         Left            =   240
         TabIndex        =   24
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label lblProduct 
         Alignment       =   1  'Right Justify
         Caption         =   "lblProduct"
         Height          =   315
         Left            =   240
         TabIndex        =   23
         Top             =   720
         Width           =   1455
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2280
         TabIndex        =   17
         Top             =   6840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJobWareHouseOutputEx.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4200
         TabIndex        =   18
         Top             =   6840
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditJobWHOutputEx"
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

Private m_PartTypes As Collection
Private m_PartItems As Collection
Private m_Locations As Collection
Private m_Units As Collection

Private Sub cboBinNo_Change()
   m_HasModify = True
End Sub

Private Sub cboBinNo_Click()
   m_HasModify = True
End Sub

Private Sub cboLockNo_Change()
   m_HasModify = True
End Sub

Private Sub cboLockNo_Click()
   m_HasModify = True
End Sub

Private Sub cboPalletNo_Change()
   m_HasModify = True
End Sub

Private Sub cboPalletNo_Click()
   m_HasModify = True
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
      
   Call InitNormalLabel(lblType, MapText("ประเภทสินค้า"))
   Call InitNormalLabel(lblProduct, MapText("เบอร์สินค้า"))
   Call InitNormalLabel(lblProductType, MapText("ชนิดสินค้า"))
   Call InitNormalLabel(lblLotNo, MapText("Lot. No."))
   Call InitNormalLabel(lblBathAmount, MapText("จำนวนแบท (B)"))
   Call InitNormalLabel(lblGoodAmount, MapText("ของดี"))
   Call InitNormalLabel(lblLoseAmount, MapText("ของเสีย"))
   Call InitNormalLabel(lblWeightPerPack, MapText("ขนาดถุง"))
   Call InitNormalLabel(lblAmount, MapText("น้ำหนักรวม"))
   Call InitNormalLabel(lblStdAmount, MapText("จำนวนมาตรฐาน"))
   Call InitNormalLabel(lblRestAmount, MapText("จำนวนเศษ (กก.)"))
   Call InitNormalLabel(lblBinNo, MapText("เบอร์ถังอาหาร"))
   Call InitNormalLabel(lblPalletNo, MapText("พาเลทที่วาง"))
   Call InitNormalLabel(lblLockNo, MapText("ล๊อค"))
   Call InitNormalLabel(lblPackDate, MapText("วันที่บรรจุ"))
   Call InitNormalLabel(lblTimePackBegin, MapText("เวลาเริ่มบรรจุ"))
   Call InitNormalLabel(lblTimePackEnd, MapText("เวลาหลังบรรจุ"))
   Call InitNormalLabel(lblPlace, MapText("สถานที่จัดเก็บ"))
   Call InitNormalLabel(lblPackAmount, MapText("จำนวนบรรจุ (ถุง)"))
   Call InitNormalLabel(lblNote, MapText("หมายเหตุ"))
'   Call InitNormalLabel(lblSerialNo, MapText("ซีเรียล"))
'   Call InitNormalLabel(lblRef, MapText("หมายเลขอ้างอิง"))
   
   Call txtLotNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtBathAmount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtGoodAmount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtLoseAmount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtWeightPerPack.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtAmount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtStdAmount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtRestAmount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtPackAmount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtNote.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   '   Call txtSerialNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
'   Call txtRef.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call InitCombo(cboBinNo)
   Call InitCombo(cboPalletNo)
   Call InitCombo(cboLockNo)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Dim Ma As CJobInputWarehouse
         Set Ma = TempCollection.Item(ID)

        uctlPartTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPartTypeLookup.MyCombo, Ma.PART_TYPE_ID)
        uctlProductLookup.MyCombo.ListIndex = IDToListIndex(uctlProductLookup.MyCombo, Ma.PART_ITEM_ID)
        uctlProductTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlProductTypeLookup.MyCombo, Ma.PRODUCT_TYPE_ID)
        txtLotNo.Text = Ma.LOT_NO
        txtBathAmount.Text = Ma.BATCH_NO
        cboBinNo.ListIndex = IDToListIndex(cboBinNo, Ma.BIN_NO)
        cboPalletNo.ListIndex = IDToListIndex(cboPalletNo, Ma.PALLET_NO)
        cboLockNo.ListIndex = IDToListIndex(cboLockNo, Ma.LOCK_NO)
        txtGoodAmount.Text = Ma.GOOD_AMOUNT
        txtLoseAmount.Text = Ma.LOSE_AMOUNT
        txtWeightPerPack.Text = Ma.WEIGHT_PER_PACK
        txtAmount.Text = Ma.TX_AMOUNT
        txtRestAmount.Text = Ma.REST_AMOUNT
        txtPackAmount.Text = Ma.PACK_AMOUNT
        txtStdAmount.Text = Ma.STD_AMOUNT
        uctlPackDate.ShowDate = Ma.PACK_DATE
        uctlPlaceLookup.MyCombo.ListIndex = IDToListIndex(uctlPlaceLookup.MyCombo, Ma.LOCATION_ID)
        txtNote.Text = Ma.NOTE
        uctlPackDate.ShowDate = Ma.PACK_DATE
        txtTimePackBegin.HR = HOUR(Ma.TIME_PACK_BEGIN)
        txtTimePackBegin.MI = Minute(Ma.TIME_PACK_BEGIN)
        txtTimePackEnd.HR = HOUR(Ma.TIME_PACK_END)
        txtTimePackEnd.MI = Minute(Ma.TIME_PACK_END)
        
'        txtSerialNo.Text = Ma.SERIAL_NUMBER
'        txtRef.Text = Ma.INOUT_REF
        
        
         cmdOK.Enabled = (COMMIT_FLAG <> "Y")
      End If
   End If
   
   Call EnableForm(Me, True)
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
         
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim Ma As CJobInputWarehouse
   If ShowMode = SHOW_ADD Then
      Set Ma = New CJobInputWarehouse
   Else
      Set Ma = TempCollection.Item(ID)
   End If
   
   Ma.PART_TYPE_ID = uctlPartTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPartTypeLookup.MyCombo.ListIndex))
   Ma.PART_TYPE_NAME = uctlPartTypeLookup.MyCombo.Text
   Ma.PART_DESC = uctlProductLookup.MyCombo.Text
   Ma.PART_NO = uctlProductLookup.MyTextBox.Text
   Ma.PART_ITEM_ID = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
   Ma.PRODUCT_TYPE_ID = uctlProductTypeLookup.MyCombo.ItemData(Minus2Zero(uctlProductTypeLookup.MyCombo.ListIndex))
   Ma.LOT_NO = Val(txtLotNo.Text)
   Ma.BATCH_NO = Trim(txtBathAmount.Text)
   Ma.BIN_NO = cboBinNo.ItemData(Minus2Zero(cboBinNo.ListIndex)) 'Trim(txtBinNo.Text)
   Ma.PALLET_NO = cboPalletNo.ItemData(Minus2Zero(cboPalletNo.ListIndex)) 'Trim(txtPalletNo.Text)
   Ma.LOCK_NO = cboLockNo.ItemData(Minus2Zero(cboLockNo.ListIndex)) 'Trim(txtLockNo.Text)
   Ma.GOOD_AMOUNT = Val(txtGoodAmount.Text)
   Ma.LOSE_AMOUNT = Val(txtLoseAmount.Text)
   Ma.WEIGHT_PER_PACK = Val(txtWeightPerPack.Text)
   Ma.PACK_AMOUNT = Val(txtPackAmount.Text)
   Ma.REST_AMOUNT = Val(txtRestAmount.Text)
   Ma.PACK_DATE = uctlPackDate.ShowDate
   Ma.TIME_PACK_BEGIN = txtTimePackBegin.HR & ":" & txtTimePackBegin.MI
   Ma.TIME_PACK_END = txtTimePackEnd.HR & ":" & txtTimePackEnd.MI
   Ma.TX_AMOUNT = Val(txtAmount.Text)
   Ma.NOTE = txtNote.Text
   Ma.LOCATION_ID = uctlPlaceLookup.MyCombo.ItemData(Minus2Zero(uctlPlaceLookup.MyCombo.ListIndex))
   Ma.LOCATION_NO = uctlPlaceLookup.MyTextBox.Text
   Ma.LOCATION_NAME = uctlPlaceLookup.MyCombo.Text
   Ma.TX_TYPE = "I"
   Ma.STD_AMOUNT = Val(txtStdAmount.Text)
'   Ma.SERIAL_NUMBER = txtSerialNo.Text
'   Ma.INOUT_REF = txtRef.Text
   If ShowMode = SHOW_ADD Then
      Ma.Flag = "A"
      Call TempCollection.add(Ma)
   Else
      If Ma.Flag <> "A" Then
         Ma.Flag = "E"
      End If
   End If
   
   SaveData = True
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
      
     Call LoadMaster(uctlProductTypeLookup.MyCombo, m_Units, PRODUCT_TYPE)
     Set uctlProductTypeLookup.MyCollection = m_Units
     
     Call LoadLocation(cboBinNo, Nothing, 2, , , , , "BIN")
     Call LoadLocation(cboPalletNo, Nothing, 2, , , , , "PALLET")
     Call LoadLocation(cboLockNo, Nothing, 2, , , , , "LOCK")
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
         uctlPackDate.ShowDate = Now
         txtTimePackBegin.HR = HOUR(Now)
         txtTimePackBegin.MI = Minute(Now)
         txtTimePackEnd.HR = HOUR(Now)
         txtTimePackEnd.MI = Minute(Now)
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
   Set m_Units = New Collection
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_PartTypes = Nothing
   Set m_PartItems = Nothing
   Set m_Locations = Nothing
   Set m_Units = Nothing
End Sub





Private Sub txtBathAmount_Change()
   m_HasModify = True
End Sub



Private Sub txtGoodAmount_Change()
On Error Resume Next
   m_HasModify = True
   txtAmount.Text = (Val(txtGoodAmount.Text) + Val(txtLoseAmount.Text)) * Val(txtWeightPerPack.Text)
   txtPackAmount.Text = Val(txtGoodAmount.Text) + Val(txtLoseAmount.Text)
End Sub

Private Sub txtLockNo_Change()
   m_HasModify = True
End Sub

Private Sub txtLoseAmount_Change()
On Error Resume Next
   m_HasModify = True
   txtAmount.Text = (Val(txtGoodAmount.Text) + Val(txtLoseAmount.Text)) * Val(txtWeightPerPack.Text)
   txtPackAmount.Text = Val(txtGoodAmount.Text) + Val(txtLoseAmount.Text)
End Sub

Private Sub txtLotNo_Change()
   m_HasModify = True
End Sub

Private Sub txtNote_Change()
   m_HasModify = True
End Sub

Private Sub txtPackAmount_Change()
   m_HasModify = True
   txtAmount.Text = Val(txtPackAmount.Text) * Val(txtWeightPerPack.Text)
End Sub

Private Sub txtPalletNo_Change()
   m_HasModify = True
End Sub

Private Sub txtRestAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtStdAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtTimePackBegin_HasChange()
   m_HasModify = True
End Sub

Private Sub txtTimePackEnd_HasChange()
   m_HasModify = True
End Sub

Private Sub txtWeightPerPack_Change()
On Error Resume Next
   m_HasModify = True
   txtAmount.Text = (Val(txtGoodAmount.Text) + Val(txtLoseAmount.Text)) * Val(txtWeightPerPack.Text)
End Sub

Private Sub uctlPackDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlPartTypeLookup_Change()
Dim PartTypeID As Long
Dim Pt As CPartType

   PartTypeID = uctlPartTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPartTypeLookup.MyCombo.ListIndex))
   
   If PartTypeID > 0 Then
      Set Pt = GetPartType(m_PartTypes, Trim(str(PartTypeID)))
      Call LoadPartItem(uctlProductLookup.MyCombo, m_PartItems, PartTypeID, "N")
      Set uctlProductLookup.MyCollection = m_PartItems
   
      Call LoadLocation(uctlPlaceLookup.MyCombo, m_Locations, 2, , , Pt.PART_GROUP_ID)
      Set uctlPlaceLookup.MyCollection = m_Locations
   End If
   
   m_HasModify = True
End Sub

Private Sub txtAmount_Change()
   m_HasModify = True
   txtStdAmount.Text = txtAmount.Text
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
Dim Pi As CPartItem
Dim PartItemID As Long

   PartItemID = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
   If PartItemID > 0 Then
      Set Pi = GetPartItem(m_PartItems, Trim(str(PartItemID)))
      txtWeightPerPack.Text = Pi.WEIGHT_PER_PACK
   End If
   
   m_HasModify = True
End Sub

Private Sub uctlProductTypeLookup_Change()
   m_HasModify = True
End Sub
