VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditJobOutputEx3 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditJobOutputEx3.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   7275
      Left            =   0
      TabIndex        =   19
      Top             =   600
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   12832
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboHead 
         Height          =   510
         Left            =   8760
         TabIndex        =   20
         Top             =   810
         Visible         =   0   'False
         Width           =   800
      End
      Begin VB.ComboBox cboLotNo 
         Height          =   510
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1680
         Width           =   2085
      End
      Begin prjFarmManagement.uctlDate uctlPackDate 
         Height          =   495
         Left            =   1800
         TabIndex        =   13
         Top             =   4200
         Width           =   3855
         _extentx        =   6800
         _extenty        =   873
      End
      Begin VB.ComboBox cboLockNo 
         Height          =   510
         Left            =   5445
         TabIndex        =   6
         Top             =   2220
         Width           =   1725
      End
      Begin VB.ComboBox cboBinNo 
         Height          =   510
         Left            =   1800
         TabIndex        =   5
         Top             =   2220
         Width           =   2085
      End
      Begin prjFarmManagement.uctlTextLookup uctlProductLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   1
         Top             =   720
         Width           =   5385
         _extentx        =   9499
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtAmount 
         Height          =   435
         Left            =   1800
         TabIndex        =   11
         Top             =   3720
         Width           =   2085
         _extentx        =   4524
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPlaceLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   15
         Top             =   5160
         Width           =   5385
         _extentx        =   9499
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPartTypeLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   0
         Top             =   240
         Width           =   5385
         _extentx        =   9499
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtWeightPerPack 
         Height          =   435
         Left            =   1800
         TabIndex        =   9
         Top             =   3240
         Width           =   2085
         _extentx        =   4524
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlProductTypeLookup 
         Height          =   435
         Left            =   10800
         TabIndex        =   21
         Top             =   810
         Visible         =   0   'False
         Width           =   3345
         _extentx        =   5900
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtGoodAmount 
         Height          =   435
         Left            =   5445
         TabIndex        =   10
         Top             =   3240
         Width           =   1725
         _extentx        =   4524
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPackAmount 
         Height          =   435
         Left            =   5445
         TabIndex        =   12
         Top             =   3720
         Width           =   1725
         _extentx        =   4524
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTime uctlTime1 
         Height          =   375
         Left            =   60500
         TabIndex        =   36
         Top             =   5040
         Width           =   1335
         _extentx        =   2355
         _extenty        =   661
      End
      Begin prjFarmManagement.uctlTime txtTimePackBegin 
         Height          =   375
         Left            =   1800
         TabIndex        =   14
         Top             =   4680
         Width           =   1335
         _extentx        =   2355
         _extenty        =   661
      End
      Begin prjFarmManagement.uctlTextBox txtNote 
         Height          =   435
         Left            =   1800
         TabIndex        =   16
         Top             =   5640
         Width           =   5325
         _extentx        =   9393
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPalletFrom 
         Height          =   435
         Left            =   1800
         TabIndex        =   7
         Top             =   2760
         Width           =   2085
         _extentx        =   4524
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPalletPerUnit 
         Height          =   435
         Left            =   5445
         TabIndex        =   8
         Top             =   2760
         Width           =   1725
         _extentx        =   4524
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtLotNoNew 
         Height          =   435
         Left            =   5460
         TabIndex        =   4
         Top             =   1680
         Width           =   1695
         _extentx        =   2990
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlDate uctlStartDate 
         Height          =   495
         Left            =   1800
         TabIndex        =   2
         Top             =   1200
         Width           =   3855
         _extentx        =   6800
         _extenty        =   873
      End
      Begin VB.Label lblStartDate 
         Alignment       =   1  'Right Justify
         Caption         =   "lblStartDate"
         Height          =   375
         Left            =   -240
         TabIndex        =   45
         Top             =   1200
         Width           =   1905
      End
      Begin VB.Label lblUnit2 
         Alignment       =   1  'Right Justify
         Caption         =   "lblUnit2"
         Height          =   345
         Left            =   7200
         TabIndex        =   44
         Top             =   3720
         Width           =   375
      End
      Begin VB.Label lblHead 
         Alignment       =   1  'Right Justify
         Caption         =   "lblHead"
         Height          =   345
         Left            =   7680
         TabIndex        =   43
         Top             =   840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblLotNo2 
         Alignment       =   1  'Right Justify
         Caption         =   "lblLotNo2"
         Height          =   315
         Left            =   4440
         TabIndex        =   42
         Top             =   1680
         Width           =   795
      End
      Begin VB.Label lblLotNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblJobNo"
         Height          =   315
         Left            =   960
         TabIndex        =   41
         Top             =   1680
         Width           =   675
      End
      Begin Threed.SSCommand cmdAuto2 
         Height          =   405
         Left            =   3960
         TabIndex        =   40
         Top             =   1725
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJobOutputEx3.frx":08CA
         ButtonStyle     =   3
      End
      Begin VB.Label lblUnit1 
         Alignment       =   1  'Right Justify
         Caption         =   "lblUnit1"
         Height          =   345
         Left            =   7200
         TabIndex        =   39
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label lblPalletPerUnit 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPalletPerUnit"
         Height          =   345
         Left            =   3960
         TabIndex        =   38
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label lblNote 
         Alignment       =   1  'Right Justify
         Caption         =   "lblRef"
         Height          =   345
         Left            =   0
         TabIndex        =   37
         Top             =   5640
         Width           =   1695
      End
      Begin VB.Label lblTimePackBegin 
         Alignment       =   1  'Right Justify
         Caption         =   "lblTimePackEnd"
         Height          =   375
         Left            =   -240
         TabIndex        =   29
         Top             =   4680
         Width           =   1905
      End
      Begin VB.Label lblPackDate 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPackAmount"
         Height          =   375
         Left            =   -240
         TabIndex        =   35
         Top             =   4200
         Width           =   1905
      End
      Begin VB.Label lblGoodAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblGoodAmount"
         Height          =   345
         Left            =   4320
         TabIndex        =   34
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label lblBinNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblBinNo"
         Height          =   375
         Left            =   480
         TabIndex        =   33
         Top             =   2160
         Width           =   1185
      End
      Begin VB.Label lblPalletFrom 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPalletFrom"
         Height          =   345
         Left            =   0
         TabIndex        =   32
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label lblLockNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblLockNo"
         Height          =   345
         Left            =   3600
         TabIndex        =   31
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label lblProductType 
         Alignment       =   1  'Right Justify
         Caption         =   "lblProductType"
         Height          =   315
         Left            =   9720
         TabIndex        =   30
         Top             =   840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblWeightPerPack 
         Alignment       =   1  'Right Justify
         Caption         =   "lblWeightPerPack"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   3240
         Width           =   1545
      End
      Begin VB.Label lblPackAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPackAmount"
         Height          =   375
         Left            =   3360
         TabIndex        =   27
         Top             =   3720
         Width           =   1905
      End
      Begin VB.Label lblPlace 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPlace"
         Height          =   315
         Left            =   240
         TabIndex        =   26
         Top             =   5160
         Width           =   1455
      End
      Begin VB.Label lblAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblAmount"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   3720
         Width           =   1545
      End
      Begin VB.Label lblType 
         Alignment       =   1  'Right Justify
         Caption         =   "lblType"
         Height          =   315
         Left            =   270
         TabIndex        =   24
         Top             =   240
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
         Top             =   6240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJobOutputEx3.frx":0BE4
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4080
         TabIndex        =   18
         Top             =   6240
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditJobOutputEx3"
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
Private m_HasModify3 As Boolean 'flag พิเศษไว้บังคับให้เปลี่ยนแปลงค่าตัวที่ต้องการ อย่าง อัตโนมัติ
Private m_Rs As ADODB.Recordset
Private m_Input_combo As Collection
Private m_Input1_combo As Collection
Public HeaderText As String
Public ID As Long
Public JOB_INOUT_ID As Long
Public OKClick As Boolean
Public TempCollection As Collection
Public TempCollection2 As Collection
Public Temp_IWD As CInventoryWHDoc
Public COMMIT_FLAG As String
Public StartJob As Date
Public StopJob As Date
Public PartType As Long
Private PartItemID As Long
Private m_CollLotItemWh As Collection
Private m_PartTypes As Collection
Private m_PartItems As Collection
Private m_Locations As Collection
Private m_Units As Collection
Private Lt As cLot
Public m_JobInOut As Collection
Private IWD As CInventoryWHDoc
Private LWH As CLotItemWH

Private OldLotNo As Long
Private NewLotNo As Long
Public DocumentType As Long


Private Sub cboBinNo_Change()
   m_HasModify = True
End Sub

Private Sub cboBinNo_Click()
   m_HasModify = True
End Sub

Private Sub cboBinNo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
KeyAscii = 0
End Sub

Private Sub cboHead_Change()
m_HasModify = True
End Sub

Private Sub cboHead_Click()
   If cboHead.ListIndex = 1 Then 'head pack 1
      Call LoadLocation(cboBinNo, Nothing, 2, , -2, , , "BIN")
   ElseIf cboHead.ListIndex = 2 Then 'head pack 2
      Call LoadLocation(cboBinNo, Nothing, 2, , -3, , , "BIN")
   End If
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

Private Sub cboLockNo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
  KeyAscii = 0
End Sub

Private Sub cboLotNo_Change()
   m_HasModify = True
End Sub

Private Sub cboLotNo_Click()
m_HasModify = True
End Sub

Private Sub cboLotNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys ("{TAB}")
End If
KeyAscii = 0
End Sub

Private Sub cmdAuto2_Click()
Dim No As String
Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
Dim IsOK As Boolean

  Set oMenu = New cPopupMenu
  lMenuChosen = oMenu.Popup("เพิ่ม LOT NO ใหม่", "-", "บันทึก", "-", "ลบ", "-", "LOT NO อื่นๆ")
  If lMenuChosen = 0 Then
      Exit Sub
  ElseIf lMenuChosen = 1 Then
      lblLotNo2.Enabled = True
      txtLotNoNew.Enabled = True
      txtLotNoNew.SetFocus
   ElseIf lMenuChosen = 3 Then
      If Not VerifyTextControl(lblLotNo2, txtLotNoNew, False) Then
        Exit Sub
      End If
   
      If Not VerifyDate(lblStartDate, uctlStartDate, False) Then
        Exit Sub
      End If
      
      Set Lt = New cLot
      Lt.AddEditMode = SHOW_ADD
      No = "LG" & Right(Format(Year(uctlStartDate.ShowDate) + 543, "0000"), 2) & Format(uctlStartDate.ShowDate, "mm") & Format(uctlStartDate.ShowDate, "dd")
      Lt.LOT_NO = No & Format(Val(txtLotNoNew.Text), "000")
      Lt.LOT_DATE = uctlStartDate.ShowDate
      
      If Not CheckUniqueNs(LOT_UNIQUE, Lt.LOT_NO, ID) Then
          glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & Lt.LOT_NO & " " & MapText("อยู่ในระบบแล้ว")
          glbErrorLog.ShowUserError
          Call LoadLotFromLot(cboLotNo, Nothing, , uctlStartDate.ShowDate, uctlStartDate.ShowDate, , , 1, , 2, , Lt.LOT_NO)
          Call EnableForm(Me, True)
          Exit Sub
       End If
   
      Call Lt.AddEditData
      Call LoadLotIdByPartItem(cboLotNo, m_CollLotItemWh, , uctlStartDate.ShowDate, uctlStartDate.ShowDate, , PartItemID, 5, 1, 1, "I", TempCollection2, 1, Lt)
      lblLotNo2.Enabled = False
      txtLotNoNew.Enabled = False
   ElseIf lMenuChosen = 5 Then
      If Not VerifyCombo(lblLotNo, cboLotNo, False) Then
         Exit Sub
      End If
      
      Call EnableForm(Me, False)
      If Not glbDaily.DeleteLot(cboLotNo.ItemData(Minus2Zero(cboLotNo.ListIndex)), IsOK, True, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
      Call LoadLotFromLot(cboLotNo, Nothing, , uctlStartDate.ShowDate, uctlStartDate.ShowDate, , , 1, , 2)
      Call EnableForm(Me, True)

    ElseIf lMenuChosen = 7 Then
      Call LoadLotFromLot(cboLotNo, Nothing, , , , , , 1, , 2)
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
   cmdAuto2.Picture = LoadPicture(glbParameterObj.NormalButton1)
     
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
   Call InitNormalLabel(lblLotNo, MapText("Lot การผลิต"))
   Call InitNormalLabel(lblLotNo2, MapText("Lot"))
   Call InitNormalLabel(lblGoodAmount, MapText("จำนวนในโกดัง"))
   Call InitNormalLabel(lblWeightPerPack, MapText("ขนาดถุง"))
   Call InitNormalLabel(lblAmount, MapText("น้ำหนักรวม"))
   Call InitNormalLabel(lblBinNo, MapText("เบอร์ถัง"))
   Call InitNormalLabel(lblPalletFrom, MapText("พาเลทยกมา"))
   Call InitNormalLabel(lblPalletPerUnit, MapText("ยอดคงเหลือ"))

   Call InitNormalLabel(lblHead, MapText("หัวแพ็ค"))
   
   If DocumentType = 15 Then
       Call InitNormalLabel(lblUnit1, MapText("ถุง"))
       Call InitNormalLabel(lblUnit2, MapText("ถุง"))
   ElseIf DocumentType = 16 Then
      Call InitNormalLabel(lblUnit1, MapText("ก.ก."))
      Call InitNormalLabel(lblUnit2, MapText("ก.ก."))
   End If
   
   Call InitNormalLabel(lblStartDate, MapText("วันที่ผลิต"))
   Call InitNormalLabel(lblLockNo, MapText("ล๊อค"))
   Call InitNormalLabel(lblPackDate, MapText("วันที่นับสต๊อก"))
   Call InitNormalLabel(lblTimePackBegin, MapText("เวลานับสต๊อกเสร็จ"))
   Call InitNormalLabel(lblPlace, MapText("สถานที่จัดเก็บ"))
   Call InitNormalLabel(lblPackAmount, MapText("จำนวนบรรจุ"))
   Call InitNormalLabel(lblNote, MapText("หมายเหตุ"))

   Call txtLotNoNew.SetTextLenType(TEXT_STRING, glbSetting.LOT_NO)
   lblLotNo2.Enabled = False
   txtLotNoNew.Enabled = False
   Call txtPalletFrom.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   txtPalletFrom.Enabled = False
   Call txtPalletPerUnit.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   Call txtGoodAmount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtWeightPerPack.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtAmount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtAmount.Enabled = False
   Call txtPackAmount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPackAmount.Enabled = False
   Call txtNote.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtAmount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   
   Call uctlProductLookup.MyTextBox.SetKeySearch("PART_NO")

   Call InitCombo(cboLotNo)
   Call InitCombo(cboBinNo)
   Call InitCombo(cboLockNo)
   Call InitCombo(cboHead)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdAuto2, MapText("A"))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim ID2 As Long
If Not Temp_IWD Is Nothing Then
   Set TempCollection = Temp_IWD.C_LotItemsWH
End If
   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         If CountItem(TempCollection) <= 0 Then
            Call EnableForm(Me, True)
            Exit Sub
         End If

         Set LWH = TempCollection.Item(ID)
         
         If CountItem(LWH.C_LotDoc) <= 0 Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
         
         ID2 = 1
         If LWH.C_LotDoc.Item(ID2) Is Nothing Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
         
         Dim LTD As CLotDoc
         Set LTD = LWH.C_LotDoc.Item(ID2)
         
         uctlPartTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPartTypeLookup.MyCombo, LWH.PART_TYPE_ID)
         uctlProductLookup.MyCombo.ListIndex = IDToListIndex(uctlProductLookup.MyCombo, LWH.PART_ITEM_ID)
         txtAmount.Text = LWH.TX_AMOUNT
         uctlPlaceLookup.MyCombo.ListIndex = IDToListIndex(uctlPlaceLookup.MyCombo, LWH.LOCATION_ID)
         uctlProductTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlProductTypeLookup.MyCombo, LWH.PRODUCT_TYPE_ID)
         cboLotNo.ListIndex = IDToListIndex(cboLotNo, LTD.LOT_ID)

         cboLockNo.ListIndex = IDToListIndex(cboLockNo, LWH.LOCK_NO)
         cboHead.ListIndex = IDToListIndex(cboHead, LWH.HEAD_PACK_NO)
         
         If cboHead.ListIndex = 1 Then 'head pack 1
            Call LoadLocation(cboBinNo, Nothing, 2, , -2, , , "BIN")
         ElseIf cboHead.ListIndex = 2 Then 'head pack 2
            Call LoadLocation(cboBinNo, Nothing, 2, , -3, , , "BIN")
         End If
'         uctlStartDate.ShowDate = LWH.BL_START_DATE

         If LWH.BL_START_DATE > 0 Then
            uctlStartDate.ShowDate = LWH.BL_START_DATE
         Else
            uctlStartDate.ShowDate = LWH.START_DATE
            m_HasModify3 = True
         End If
         cboBinNo.ListIndex = IDToListIndex(cboBinNo, LWH.BIN_NO)
         txtGoodAmount.Text = LWH.GOOD_AMOUNT
         txtWeightPerPack.Text = LWH.WEIGHT_PER_PACK
         txtAmount.Text = LWH.TX_AMOUNT
         txtPackAmount.Text = LWH.PACK_AMOUNT
         uctlPlaceLookup.MyCombo.ListIndex = IDToListIndex(uctlPlaceLookup.MyCombo, LWH.LOCATION_ID)
         uctlPackDate.ShowDate = LWH.PACK_DATE
         txtTimePackBegin.HR = HOUR(LWH.TIME_PACK_BEGIN)
         txtTimePackBegin.MI = Minute(LWH.TIME_PACK_BEGIN)
         txtNote.Text = LWH.NOTE
         'เป็นการ Gen ตอนสร้าง เพราะฉะนั้นตอนแก้ไขก็ไม่ต้องให้แสดง
         txtPalletFrom.Text = LWH.FULL_PALLET_FROM
         txtPalletPerUnit.Text = LWH.FULL_UNIT_PER_PALLET
         
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
Dim I As Long
Dim tempPallet As Collection

   If (Val(uctlPartTypeLookup.MyTextBox.Text) <> 10) And (Val(uctlPartTypeLookup.MyTextBox.Text) <> 22) Then
        glbErrorLog.LocalErrorMsg = MapText("กรุณาเลือก อาหารผลิตเสร็จ หรือ อาหารสำเร็จรูป เท่านั้น")
      glbErrorLog.ShowUserError
      Exit Function
   End If

   If Not VerifyCombo(lblType, uctlPartTypeLookup.MyCombo, False) Then
      Exit Function
   End If

   If Not VerifyCombo(lblProduct, uctlProductLookup.MyCombo, False) Then
      Exit Function
   End If
   
   If Not VerifyCombo(lblLotNo, cboLotNo, False) Then
      Exit Function
   End If
   
   If Not VerifyCombo(lblBinNo, cboBinNo, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblAmount, txtAmount, False) Then
      Exit Function
   End If
   
  If Not VerifyCombo(lblPlace, uctlPlaceLookup.MyCombo, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblPalletPerUnit, txtPalletPerUnit, False) Then
      Exit Function
   End If

'   If Val(txtAmount.Text) = 0 Then
'        Call MsgBox(lblAmount.Caption & "ต้องไม่เท่ากับ 0 ", vbOKOnly, PROJECT_NAME)
'        Exit Function
'   End If

  If (txtTimePackBegin.HR) = "24" Then
        txtTimePackBegin.HR = "00"
        txtTimePackBegin.MI = "00"
  End If
        
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If

   If ShowMode = SHOW_ADD Then
      Set LWH = New CLotItemWH
   Else
      Set TempCollection = Temp_IWD.C_LotItemsWH
      Set LWH = TempCollection.Item(ID)
   End If
   
   LWH.BL_START_DATE = uctlStartDate.ShowDate
   LWH.PART_ITEM_ID = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
   LWH.PART_DESC = uctlProductLookup.MyCombo.Text
   LWH.PART_NO = uctlProductLookup.MyTextBox.Text
   LWH.PART_TYPE_ID = uctlPartTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPartTypeLookup.MyCombo.ListIndex))
   LWH.PRODUCT_TYPE_ID = uctlProductTypeLookup.MyCombo.ItemData(Minus2Zero(uctlProductTypeLookup.MyCombo.ListIndex))
   LWH.BIN_NO = cboBinNo.ItemData(Minus2Zero(cboBinNo.ListIndex))
   LWH.LOCK_NO = cboLockNo.ItemData(Minus2Zero(cboLockNo.ListIndex))
   LWH.HEAD_PACK_NO = cboHead.ItemData(Minus2Zero(cboHead.ListIndex))
   LWH.GOOD_AMOUNT = FormatNumber(txtGoodAmount.Text)
   LWH.BALANCE_AMOUNT = FormatNumber(txtGoodAmount.Text) 'ไว้ตอนปรับยอด
   LWH.WEIGHT_PER_PACK = Val(txtWeightPerPack.Text)
   LWH.PACK_AMOUNT = Val(txtPackAmount.Text)
   LWH.PACK_DATE = uctlPackDate.ShowDate
   LWH.TIME_PACK_BEGIN = txtTimePackBegin.HR & ":" & txtTimePackBegin.MI
   LWH.TX_AMOUNT = Val(txtAmount.Text)
   LWH.NOTE = txtNote.Text
   LWH.LOCATION_ID = uctlPlaceLookup.MyCombo.ItemData(Minus2Zero(uctlPlaceLookup.MyCombo.ListIndex))
   LWH.LOCATION_NO = uctlPlaceLookup.MyTextBox.Text
   LWH.LOCATION_NAME = uctlPlaceLookup.MyCombo.Text
   LWH.TX_TYPE = "I" 'รับเข้า
   
   LWH.FULL_PALLET_FROM = Val(txtPalletFrom.Text)
   LWH.FULL_PALLET_TO = -1
   LWH.FULL_UNIT_PER_PALLET = Val(txtPalletPerUnit.Text)
   LWH.SCRAP_PALLET = -1
   LWH.SCRAP_UNIT_PER_PALLET = -1
   LWH.AddEditMode = ShowMode
   
   Dim LTD As CLotDoc
   Dim PD As CPalletDoc
   Set LTD = New CLotDoc
   If ShowMode = SHOW_ADD Then
         Set PD = New CPalletDoc
         PD.Flag = "A"
         PD.PALLET_DOC_NO = txtPalletFrom.Text
         PD.CAPACITY_AMOUNT = Val(txtPalletPerUnit.Text)
         PD.TX_TYPE = "I"
         PD.AddEditMode = ShowMode
         Call LTD.C_PalletDoc.add(PD)
         Set PD = Nothing
      
         LTD.Flag = "A"
         LTD.LOT_ID = cboLotNo.ItemData(Minus2Zero(cboLotNo.ListIndex))
         LTD.BIN_NO = cboBinNo.ItemData(Minus2Zero(cboBinNo.ListIndex))
         LTD.AddEditMode = ShowMode
         Call LWH.C_LotDoc.add(LTD)
         LWH.Flag = "A"
         Call TempCollection.add(LWH)
   Else
      LWH.Flag = "E"
      Set LTD = LWH.C_LotDoc.Item(1)
      LTD.LOT_ID = cboLotNo.ItemData(Minus2Zero(cboLotNo.ListIndex))
      LTD.BIN_NO = cboBinNo.ItemData(Minus2Zero(cboBinNo.ListIndex))
      LTD.AddEditMode = ShowMode
      LTD.Flag = "E"
      
      Set PD = LTD.C_PalletDoc(1)
         PD.CAPACITY_AMOUNT = Val(txtPalletPerUnit.Text)
         PD.Flag = "E"
         PD.AddEditMode = ShowMode
      Set PD = Nothing
      Set LTD = Nothing
   End If
   SaveData = True
End Function
Function SumPalletAmount(Cl As Collection) As Long
   Dim PD As CPalletDoc
   For Each PD In Cl
      If Not PD.Flag = "D" Then
         SumPalletAmount = SumPalletAmount + PD.CAPACITY_AMOUNT
      End If
   Next PD
End Function
Private Sub CallPalletAmount()
   Dim PD As CPalletDoc
   Dim C As Long
   C = C + 1
   For Each PD In LWH.C_LotDoc.Item(ID).C_PalletDoc
      If Not LWH.C_LotDoc.Item(ID).C_PalletDoc(C).Flag = "D" And C < LWH.C_LotDoc.Item(ID).C_PalletDoc.Count Then
        LWH.C_LotDoc.Item(ID).C_PalletDoc(C).CAPACITY_AMOUNT = getFormat(uctlProductTypeLookup.MyCombo.ItemData(Minus2Zero(uctlProductTypeLookup.MyCombo.ListIndex)), Val(txtWeightPerPack.Text))
        LWH.C_LotDoc.Item(ID).C_PalletDoc(C).Flag = "E"
      Else
        LWH.C_LotDoc.Item(ID).C_PalletDoc(C).Flag = "E"
      End If
      C = C + 1
   Next PD
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadPartType(uctlPartTypeLookup.MyCombo, m_PartTypes)
      Set uctlPartTypeLookup.MyCollection = m_PartTypes
      uctlPartTypeLookup.Enabled = True
    
      Call LoadLocation(uctlPlaceLookup.MyCombo, m_Locations, 2)
      Set uctlPlaceLookup.MyCollection = m_Locations
      
     Call LoadMaster(uctlProductTypeLookup.MyCombo, m_Units, PRODUCT_TYPE)
     Set uctlProductTypeLookup.MyCollection = m_Units
   
     Call LoadLocation(cboHead, Nothing, 2, , , , , "HEAD")
     Call LoadLotFromLot(cboLotNo, Nothing, , , , , , 1, , 2)
     Call LoadLocation(cboLockNo, Nothing, 2, , , , , "LOCK")
     Call LoadLocation(cboBinNo, Nothing, 2, , , , , "BIN")
     
      txtPalletFrom.Text = "999"
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
         uctlStartDate.ShowDate = Now
         uctlPackDate.ShowDate = Now
         txtTimePackBegin.HR = HOUR(Now)
         txtTimePackBegin.MI = Minute(Now)
         uctlPartTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPartTypeLookup.MyCombo, PartType)
         uctlPartTypeLookup.Enabled = False
         Call QueryData(False)
      End If
      
      If DocumentType = 16 Then
         cboLockNo.Enabled = False
         txtWeightPerPack.Enabled = False
         txtGoodAmount.Enabled = False
      End If
      
      If Not m_HasModify3 Then
         m_HasModify = False
      End If
'      m_HasModify = False
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
   Set m_CollLotItemWh = New Collection
   Set TempCollection2 = New Collection
   Set m_JobInOut = New Collection
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
   Set m_CollLotItemWh = Nothing
   Set TempCollection2 = Nothing
   Set Lt = Nothing
   Set m_JobInOut = Nothing
End Sub

Private Sub txtGoodAmount_Change()
On Error Resume Next
   m_HasModify = True
   If DocumentType = 15 Then
      txtAmount.Text = Val(txtGoodAmount.Text) * Val(txtWeightPerPack.Text) 'น้ำหนักรวม=ดี * น้ำหนัก
   Else
      txtAmount.Text = txtGoodAmount.Text 'น้ำหนักรวม=ดี
   End If
   
   txtPackAmount.Text = Val(txtGoodAmount.Text)
End Sub

Private Sub txtLoseAmount_Change()
On Error Resume Next
   m_HasModify = True
End Sub

Private Sub txtLotNo_Change()
   m_HasModify = True
End Sub

Private Sub txtLotNoNew_KeyPress(KeyAscii As Integer)
  KeyAscii = CheckIntAscii(KeyAscii)
End Sub

Private Sub txtNote_Change()
   m_HasModify = True
End Sub

Private Sub txtPackAmount_Change()
   m_HasModify = True
   'txtAmount.Text = Val(txtPackAmount.Text) * Val(txtWeightPerPack.Text)
End Sub

Private Sub txtPalletFrom_Change()
   m_HasModify = True
End Sub

Private Sub txtPalletPerUnit_Change()
   m_HasModify = True
   txtGoodAmount.Text = Val(txtPalletPerUnit.Text)
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

Private Sub txtWeightPerPack_Change()
On Error Resume Next
   m_HasModify = True
   If DocumentType = 15 Then
      txtAmount.Text = Val(txtGoodAmount.Text) * Val(txtWeightPerPack.Text)
   Else
      txtAmount.Text = txtGoodAmount.Text
   End If
End Sub
Function getFormat(ProductType As Long, WEIGHT As Long) As Long
Dim data As Long
   If ProductType = 221 And WEIGHT = 30 Then 'ผง
      data = 48
   ElseIf ProductType = 221 And WEIGHT = 50 Then 'ผง
      data = 30
   ElseIf ProductType = 222 And WEIGHT = 30 Then 'เม็ด
      data = 60
   ElseIf ProductType = 222 And WEIGHT = 50 Then 'เม็ด
      data = 35
   ElseIf ProductType = 227 And WEIGHT = 30 Then  'ครัม
      data = 60
   Else
      data = 0
   End If
   getFormat = data
End Function
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
   PartItemID = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
   If PartItemID > 0 Then
      Set Pi = GetPartItem(m_PartItems, Trim(str(PartItemID)))
      txtWeightPerPack.Text = Pi.WEIGHT_PER_PACK
   End If
'   If ShowMode = SHOW_ADD Then
    Call LoadLotIdByPartItem(cboLotNo, m_CollLotItemWh, , , , , PartItemID, 5, 1, 1, "I", TempCollection2, 1, Lt)
'   End If
   m_HasModify = True
End Sub

Private Sub uctlProductTypeLookup_Change()
   m_HasModify = True
'      If ShowMode = SHOW_ADD Then
         txtPalletPerUnit.Text = getFormat(uctlProductTypeLookup.MyCombo.ItemData(Minus2Zero(uctlProductTypeLookup.MyCombo.ListIndex)), Val(txtWeightPerPack.Text))
'      End If
End Sub

Private Sub uctlStartDate_HasChange()
   m_HasModify = True
End Sub
