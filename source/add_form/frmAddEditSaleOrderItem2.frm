VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditSaleOrderItem2 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9435
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditSaleOrderItem2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   9435
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   7755
      Left            =   0
      TabIndex        =   27
      Top             =   600
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   13679
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSFrame SSFrame2 
         Height          =   555
         Left            =   1800
         TabIndex        =   41
         Top             =   360
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   979
         _Version        =   131073
         CaptionStyle    =   1
         Begin Threed.SSOption radCustom 
            Height          =   375
            Left            =   3960
            TabIndex        =   2
            Top             =   90
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   661
            _Version        =   131073
            Caption         =   "SSOption4"
         End
         Begin Threed.SSOption radStock 
            Height          =   375
            Left            =   1950
            TabIndex        =   1
            Top             =   90
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   661
            _Version        =   131073
            Caption         =   "SSOption4"
         End
         Begin Threed.SSOption radFeature 
            Height          =   375
            Left            =   30
            TabIndex        =   0
            Top             =   90
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   131073
            Caption         =   "SSOption4"
         End
      End
      Begin prjFarmManagement.uctlTextBox txtQuantity 
         Height          =   435
         Left            =   1800
         TabIndex        =   14
         Top             =   4050
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPartLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   9
         Top             =   2280
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlToLocationLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   10
         Top             =   2730
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPigTypeLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   7
         Top             =   1830
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTotalPrice 
         Height          =   435
         Left            =   6780
         TabIndex        =   18
         Top             =   4950
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtAvgPrice 
         Height          =   435
         Left            =   7305
         TabIndex        =   15
         Top             =   4050
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlFeatureLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   5
         Top             =   1380
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtManual 
         Height          =   465
         Left            =   1800
         TabIndex        =   3
         Top             =   900
         Width           =   5355
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtDiscount 
         Height          =   435
         Left            =   7290
         TabIndex        =   17
         Top             =   4500
         Width           =   1485
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtWeightPerPack 
         Height          =   435
         Left            =   1800
         TabIndex        =   11
         Top             =   3600
         Width           =   1455
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPackAmount 
         Height          =   435
         Left            =   4620
         TabIndex        =   12
         Top             =   3600
         Width           =   1455
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPricePerPack 
         Height          =   435
         Left            =   7320
         TabIndex        =   13
         Top             =   3600
         Width           =   1455
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtDiscountPerPack 
         Height          =   435
         Left            =   1800
         TabIndex        =   16
         Top             =   4500
         Width           =   1455
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtExcludeDiscount 
         Height          =   435
         Left            =   1800
         TabIndex        =   46
         Top             =   4950
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtManualCode 
         Height          =   465
         Left            =   1800
         TabIndex        =   20
         Top             =   5850
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtManualName 
         Height          =   465
         Left            =   3780
         TabIndex        =   21
         Top             =   5850
         Width           =   4995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTransferWage 
         Height          =   435
         Left            =   1800
         TabIndex        =   22
         Top             =   6330
         Width           =   1695
         _ExtentX        =   2514
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtStdTrfCharge 
         Height          =   435
         Left            =   7080
         TabIndex        =   23
         Top             =   6330
         Width           =   1695
         _ExtentX        =   2514
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtReturnAvg 
         Height          =   435
         Left            =   4620
         TabIndex        =   55
         Top             =   4500
         Width           =   1455
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin Threed.SSCommand cmdLastBill 
         Height          =   405
         Left            =   8880
         TabIndex        =   60
         Top             =   3600
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditSaleOrderItem2.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCheck chkUpdatePrice 
         Height          =   435
         Left            =   6720
         TabIndex        =   59
         Top             =   6960
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   767
         _Version        =   131073
         Enabled         =   0   'False
         Caption         =   "chkUpdatePrice"
      End
      Begin Threed.SSCommand cmdUpdatePrice 
         Height          =   405
         Left            =   7320
         TabIndex        =   58
         Top             =   2280
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditSaleOrderItem2.frx":0BE4
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEditCon 
         Height          =   405
         Left            =   7320
         TabIndex        =   57
         Top             =   3120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditSaleOrderItem2.frx":0EFE
         ButtonStyle     =   3
      End
      Begin VB.Label lblReturnAvg 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   3240
         TabIndex        =   56
         Top             =   4560
         Width           =   1275
      End
      Begin VB.Label lblStdTrfCharge 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   5190
         TabIndex        =   54
         Top             =   6390
         Width           =   1785
      End
      Begin VB.Label Label6 
         Height          =   345
         Left            =   8865
         TabIndex        =   53
         Top             =   6390
         Width           =   405
      End
      Begin VB.Label lblTransferWage 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   180
         TabIndex        =   52
         Top             =   6390
         Width           =   1485
      End
      Begin VB.Label Label3 
         Height          =   345
         Left            =   3555
         TabIndex        =   51
         Top             =   6390
         Width           =   1215
      End
      Begin VB.Label lblShowDetail 
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   240
         TabIndex        =   50
         Top             =   3240
         Width           =   6915
      End
      Begin Threed.SSCheck chkManualName 
         Height          =   435
         Left            =   1800
         TabIndex        =   19
         Top             =   5400
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblManualName 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   49
         Top             =   5880
         Width           =   1575
      End
      Begin VB.Label lblExcludeDiscount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   90
         TabIndex        =   48
         Top             =   5070
         Width           =   1575
      End
      Begin VB.Label Label1 
         Height          =   345
         Left            =   3900
         TabIndex        =   47
         Top             =   5010
         Width           =   495
      End
      Begin VB.Label lblDiscountPerPack 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   45
         Top             =   4560
         Width           =   1485
      End
      Begin VB.Label lblPricePerPack 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   6150
         TabIndex        =   44
         Top             =   3660
         Width           =   1095
      End
      Begin VB.Label lblPackAmount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   3240
         TabIndex        =   43
         Top             =   3660
         Width           =   1275
      End
      Begin VB.Label lblWeightPerPack 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   42
         Top             =   3660
         Width           =   1485
      End
      Begin Threed.SSOption SSOption3 
         Height          =   405
         Left            =   7260
         TabIndex        =   8
         Top             =   1830
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   714
         _Version        =   131073
         Caption         =   "SSOption1"
      End
      Begin Threed.SSOption SSOption2 
         Height          =   405
         Left            =   7260
         TabIndex        =   6
         Top             =   1380
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   714
         _Version        =   131073
         Caption         =   "SSOption1"
      End
      Begin Threed.SSOption SSOption1 
         Height          =   405
         Left            =   7260
         TabIndex        =   4
         Top             =   900
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   714
         _Version        =   131073
         Caption         =   "SSOption1"
      End
      Begin VB.Label Label7 
         Height          =   345
         Left            =   8880
         TabIndex        =   40
         Top             =   5010
         Width           =   495
      End
      Begin VB.Label lblUnit 
         Height          =   345
         Left            =   3300
         TabIndex        =   39
         Top             =   4110
         Width           =   1215
      End
      Begin VB.Label Label4 
         Height          =   345
         Left            =   8850
         TabIndex        =   38
         Top             =   4620
         Width           =   495
      End
      Begin VB.Label lblDiscount 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   6030
         TabIndex        =   37
         Top             =   4590
         Width           =   1185
      End
      Begin VB.Label lblManual 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   36
         Top             =   960
         Width           =   1485
      End
      Begin VB.Label lblFeatureCode 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   35
         Top             =   1440
         Width           =   1485
      End
      Begin VB.Label lblTotalPrice 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   5370
         TabIndex        =   34
         Top             =   5070
         Width           =   1305
      End
      Begin VB.Label lblAvgPrice 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   5985
         TabIndex        =   33
         Top             =   4110
         Width           =   1215
      End
      Begin VB.Label Label2 
         Height          =   345
         Left            =   8865
         TabIndex        =   32
         Top             =   4110
         Width           =   435
      End
      Begin VB.Label lblPigType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   180
         TabIndex        =   31
         Top             =   1860
         Width           =   1485
      End
      Begin VB.Label lblToLocation 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   180
         TabIndex        =   30
         Top             =   2790
         Width           =   1485
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3090
         TabIndex        =   24
         Top             =   6930
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4740
         TabIndex        =   25
         Top             =   6930
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblPart 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   29
         Top             =   2340
         Width           =   1485
      End
      Begin VB.Label lblQuantity 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   28
         Top             =   4110
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmAddEditSaleOrderItem2"
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

Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public TempCollection As Collection
Public TempCollection2 As Collection
Public COMMIT_FLAG As String
Public Area As Long

Private m_Sp As CSystemParam
Private m_Features As Collection
Private m_PartTypes As Collection
Private m_Parts As Collection
Private m_Locations As Collection
Private m_Houses As Collection
Private m_Pigs As Collection
Private m_PigStatuss As Collection
Private m_SubLotItems As Collection
Private m_ManualFlag As Boolean

Public m_ExWorkPricesItem As Collection
Public m_ExDeliveryCostItem As Collection
Public m_ExPromotionPartItem As Collection
Public m_ExPromotionDlcItem As Collection
Public m_Customers As Collection
Public m_DeliveryCus As Collection

Private m_SocID As Long
Public AccountID As Long
Public SubscriberID As Long
Public UsageDate As Date
Public DayInMonth As Long
Public DocumentDate As Date
Private DOLLAR As Double
Private Dollar1 As Long
Private COUNTRY_CURRENCY1 As Long
Private COUNTRY_CURRENCY2 As Long
Public DocumentType As Long
Private UNIT As String

Public CUSTOMER_ID As Long
Public DELIVERY_CUS_ITEM_ID As Long
Public PRICE_THINK_TYPE As Long
Public CAL_RATE_DELIVERY_TYPE As Long
Public CAL_PRICE_PART_CENTER_FLAG As String
Public CAL_PRICE_DLC_CENTER_FLAG As String
Public ISuctlDeliveryCusLookup As Boolean
Public PartItemID As Long
Public OldPartItemId As Long
Public PartType As Long
Public WeigthPerPack As Double
Public TypeSale As Long
Public NewUpdatePrice As Boolean

Public SuccessFlag As String
Private TempData As Collection
Private Pi As CPartItem
Private CurrentPricePerPack As Double
Private NewPricePerPack As Double

Private TempD As CExWorksPrice
Private TempD2 As CCustomer
Private TempD3 As CDoItem
Private EditPriceFlag As Boolean
Private TempUserName As String
Private TempUserName2 As String

Private Sub cboTextType_Click()
   m_HasModify = True
End Sub

Private Sub cboCountry_Click()
   m_HasModify = True
End Sub

Private Sub chkBangkok_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkFeature_Click(Value As Integer)
   m_HasModify = True
   Call ShowGui
End Sub

Private Sub chkManual_Click(Value As Integer)
   m_HasModify = True
   Call ShowGui
End Sub

Private Sub chkStock_Click(Value As Integer)
   m_HasModify = True
   Call ShowGui
End Sub

Private Sub chkManual_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub chkManualName_Click(Value As Integer)
Dim Pi As CPartItem
Dim PartItemID As Long
   
   If Check2Flag(chkManualName.Value) = "Y" Then
      If radFeature.Value Then
         txtManualCode.Text = uctlFeatureLookup.MyTextBox.Text
         txtManualName.Text = uctlFeatureLookup.MyCombo.Text
      ElseIf radStock.Value Then
         If uctlPartLookup.MyCombo.ListIndex <= 0 Then
            Exit Sub
         End If
         
         PartItemID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
         If PartItemID > 0 Then
            Set Pi = GetPartItem(m_Pigs, Trim(str(PartItemID)))
            Call InitNormalLabel(lblUnit, Pi.UNIT_NAME)
            txtManualCode.Text = Pi.BARCODE_NO
            txtManualName.Text = Pi.BILL_DESC
         End If
      End If
   Else
      txtManualCode.Enabled = False
      txtManualName.Enabled = False
   End If
   
   m_HasModify = True
End Sub

Private Sub chkManualName_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub chkUpdatePrice_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cmdEditCon_Click()
   EditPriceFlag = False
   If Val(txtPricePerPack.Text) = 0 Then
      glbErrorLog.LocalErrorMsg = "ไม่สามารถปรับลดราคาได้เนื่องจากยังไม่มีการกำหนดราคา กรุณากำหนดราคา ของรายการนี้ก่อน"
      glbErrorLog.ShowUserError
      Exit Sub
  End If


   frmVerifyAccRight.AccName = "CREDIT_PROMOTIONAL"
   frmVerifyAccRight.AccDesc = "สามารถเปลี่ยนแปลงเงื่อนไขส่งเสริมการขาย"
   Load frmVerifyAccRight
   frmVerifyAccRight.Show 1

   If frmVerifyAccRight.GrantRight Then
      TempUserName = frmVerifyAccRight.UserName
      Unload frmVerifyAccRight
      Set frmVerifyAccRight = Nothing
   Else
      Unload frmVerifyAccRight
      Set frmVerifyAccRight = Nothing
      Exit Sub
   End If
   EditPriceFlag = True
   txtPricePerPack.Enabled = True
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub ShowGui()
   If radFeature.Value Then
      uctlFeatureLookup.Enabled = True
      SSOption2.Enabled = True
      SSOption2.Value = True
         
      uctlPigTypeLookup.Enabled = False
      uctlPartLookup.Enabled = False
      uctlToLocationLookup.Enabled = False
   Else
      uctlFeatureLookup.Enabled = False
      uctlFeatureLookup.MyCombo.ListIndex = -1
      SSOption2.Enabled = False
   End If
   
   If radStock.Value Then
      uctlToLocationLookup.Enabled = True
      If m_Sp.PARAM_VALUE <> "Y" Then
         uctlPigTypeLookup.Enabled = True
      End If
      uctlPartLookup.Enabled = True
      SSOption3.Enabled = True
      SSOption3.Value = True
   Else
      uctlToLocationLookup.Enabled = False
      uctlPigTypeLookup.Enabled = False
      uctlPartLookup.Enabled = False
      
      uctlPigTypeLookup.MyCombo.ListIndex = -1
      uctlToLocationLookup.MyCombo.ListIndex = -1
      uctlPartLookup.MyCombo.ListIndex = -1
      
      SSOption3.Enabled = False
   End If

   If radCustom.Value Then
      txtManual.Enabled = True
      SSOption1.Enabled = True
      SSOption1.Value = True
'      cmdLotSelect.Visible = False
   Else
      txtManual.Enabled = False
      txtManual.Text = ""
      SSOption1.Enabled = False
   End If
End Sub

Private Function CreateConfigFlag() As String
Dim Flag1 As String
Dim Flag2 As String
Dim Flag3 As String

   Flag1 = "N"
   If radFeature.Value Then
      Flag1 = "Y"
   End If
   
   Flag2 = "N"
   If radStock.Value Then
      Flag2 = "Y"
   End If
   
   Flag3 = "N"
   If radCustom.Value Then
      Flag3 = "Y"
   End If
   
   CreateConfigFlag = Flag1 & Flag2 & Flag3
End Function

Private Sub ShowConfigFlag(ConfigFlag As String)
Dim Flag1 As String
Dim Flag2 As String
Dim Flag3 As String

   Flag1 = Mid(ConfigFlag, 1, 1)
   Flag2 = Mid(ConfigFlag, 2, 1)
   Flag3 = Mid(ConfigFlag, 3, 1)
   
   radFeature.Value = (Flag1 = "Y")
   radStock.Value = (Flag2 = "Y")
   radCustom.Value = (Flag3 = "Y")
End Sub

Private Sub InitFormLayout()
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame2.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)

   Me.KeyPreview = True
   pnlHeader.Caption = HeaderText
   Me.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_HEAD_COLOR
   SSFrame1.BackColor = GLB_FORM_COLOR
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblPart, MapText("สินค้า"))
   Call InitNormalLabel(lblQuantity, MapText("ปริมาณ"))
   Call InitNormalLabel(lblToLocation, MapText("สถานที่จัดเก็บ"))
   Call InitNormalLabel(Label2, MapText("บาท"))
   Call InitNormalLabel(lblPigType, MapText("ประเภทสินค้า"))
   Call InitNormalLabel(lblTotalPrice, MapText("ราคารวม"))
   Call InitNormalLabel(lblAvgPrice, MapText("ราคา/หน่วย"))
   Call InitNormalLabel(lblFeatureCode, MapText("สินค้า/บริการ"))
   Call InitNormalLabel(lblManual, MapText("กำหนดเอง"))
   Call InitNormalLabel(lblDiscount, MapText("ส่วนลด"))
   Call InitNormalLabel(Label4, MapText("บาท"))
   Call InitNormalLabel(lblUnit, "")
   Call InitNormalLabel(Label7, MapText("บาท"))
   Call InitNormalLabel(lblWeightPerPack, MapText("น้ำหนัก/ถุง"))
   Call InitNormalLabel(lblPackAmount, MapText("จำนวนถุง"))
   Call InitNormalLabel(lblPricePerPack, MapText("ราคา/ถุง"))
   Call InitNormalLabel(lblDiscountPerPack, MapText("ส่วนลด/ถุง"))
   Call InitNormalLabel(lblExcludeDiscount, MapText("ราคาก่อนส่วนลด"))
   Call InitNormalLabel(Label1, MapText("บาท"))
   Call InitNormalLabel(lblManualName, MapText("รหัส/ชื่อ"))
    Call InitNormalLabel(lblTransferWage, MapText("ค่าจ้างขนส่ง"))
    Call InitNormalLabel(lblStdTrfCharge, MapText("ค่าขนส่งมาตรฐาน"))
    Call InitNormalLabel(Label3, MapText("บาท"))
    Call InitNormalLabel(Label6, MapText("บาท"))
    Call InitNormalLabel(lblReturnAvg, MapText("ทุนคืน"))
    Call InitNormalLabel(lblShowDetail, MapText(""))
    
   Call InitOptionEx(radFeature, "สินค้า/บริการ")
   Call InitOptionEx(radStock, "สินค้า/วัตถุดิบ")
   Call InitOptionEx(radCustom, "กำหนดเอง")
      
   Call InitOptionEx(SSOption1, "เลือกแสดง")
   Call InitOptionEx(SSOption2, "เลือกแสดง")
   Call InitOptionEx(SSOption3, "เลือกแสดง")
  
   Call txtQuantity.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtAvgPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   txtAvgPrice.Enabled = False
   txtTotalPrice.Enabled = False
   Call txtReturnAvg.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   Call txtTotalPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   Call txtManual.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtDiscount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   Call txtWeightPerPack.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtPackAmount.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtPricePerPack.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   Call txtDiscountPerPack.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   Call txtExcludeDiscount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   txtExcludeDiscount.Enabled = False
   Call txtManualCode.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtManualName.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtManualCode.Enabled = False
   txtManualName.Enabled = False
   Call txtTransferWage.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtStdTrfCharge.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtTransferWage.Enabled = False
   txtStdTrfCharge.Enabled = False
   
   Call InitCheckBox(chkManualName, "กำหนด รหัส/ชื่อ เอง")
   Call InitCheckBox(chkUpdatePrice, "แก้ไขราคาใหม่")

'   Call InitCombo(cboRateType)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEditCon.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdUpdatePrice.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdLastBill.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdEditCon, MapText("ปรับลดราคา"))
   Call InitMainButton(cmdUpdatePrice, MapText("ปรับปรุงราคา"))
   Call InitMainButton(cmdLastBill, MapText("....."))
   
   txtManualCode.Enabled = False
   txtManualName.Enabled = False
   chkManualName.Enabled = False
   

   
   If TypeSale = 1 Then
'       txtPricePerPack.Enabled = True
'       txtDiscountPerPack.Enabled = False
       cmdUpdatePrice.Enabled = True
       cmdEditCon.Enabled = True
   ElseIf TypeSale = 2 Then
'       txtPricePerPack.Enabled = True
'       txtDiscountPerPack.Enabled = True
       cmdUpdatePrice.Enabled = False
       cmdEditCon.Enabled = False
   End If
   
   If Area = 1 Then
      radFeature.Enabled = True
   Else
      radFeature.Enabled = False
   End If

   If m_Sp.PARAM_VALUE = "Y" Then
      uctlPigTypeLookup.Enabled = False
   Else
      uctlPigTypeLookup.Enabled = True
      uctlPartLookup.Enabled = False
      uctlToLocationLookup.Enabled = False
   End If
   
   Call ShowGui
End Sub


Private Sub ShowDisplayID(Did As Long)
   If Did = 1 Then
      SSOption1.Value = True
   ElseIf Did = 2 Then
      SSOption2.Value = True
   ElseIf Did = 3 Then
      SSOption3.Value = True
   End If
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Ivd As CInventoryDoc
Dim iCount As Long
Dim Ei As CLotItem

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         If DocumentType = 18 Then
            Dim Ri As CReceiptItem
           
            Set Ri = TempCollection.Item(ID)
                    
            Call ShowConfigFlag(Ri.CONFIG_CODE)
             
            uctlPigTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPigTypeLookup.MyCombo, Ri.PART_TYPE)
            uctlPartLookup.MyCombo.ListIndex = IDToListIndex(uctlPartLookup.MyCombo, Ri.PART_ITEM_ID)
            uctlFeatureLookup.MyCombo.ListIndex = IDToListIndex(uctlFeatureLookup.MyCombo, Ri.FEATURE_ID)
            txtWeightPerPack.Text = Ri.WEIGHT_PER_PACK
            txtPackAmount.Text = Ri.PACK_AMOUNT
            txtPricePerPack.Text = Ri.PRICE_PER_PACK
            txtDiscountPerPack.Text = Ri.DISCOUNT_PER_PACK
            txtQuantity.Text = Ri.RETURN_AMOUNT
            uctlToLocationLookup.MyCombo.ListIndex = IDToListIndex(uctlToLocationLookup.MyCombo, Ri.LOCATION_ID)
            txtTotalPrice.Text = Ri.RETURN_TOTAL_PRICE
            txtReturnAvg.Text = Ri.RETURN_AVG_PRICE
            txtAvgPrice.Text = Ri.AVG_PRICE
            txtDiscount.Text = Ri.RETURN_DISCOUNT_AMOUNT
            txtManual.Text = Ri.ITEM_DESC
            Call ShowDisplayID(Ri.DISPLAY_ID)
            chkManualName.Value = FlagToCheck(Ri.MANUAL_FLAG)
            txtManualCode.Text = Ri.MANUAL_CODE
            txtManualName.Text = Ri.MANUAL_NAME
            txtTransferWage.Text = Ri.TRANSFER_WAGE
            txtStdTrfCharge.Text = Ri.STD_TRANSFER_CHARGE
           
            cmdOK.Enabled = (COMMIT_FLAG <> "Y")
            Call CalculateCurrentBath
         Else
            Dim Di As CSaleOrder
           
            Set Di = TempCollection.Item(ID)
                    
            Call ShowConfigFlag(Di.CONFIG_CODE)
            TempUserName2 = Di.USER_APPLOVE_PRICE
            OldPartItemId = Di.PART_ITEM_ID
            txtWeightPerPack.Text = Di.WEIGHT_PER_PACK
            PartType = Di.PART_TYPE
            uctlPigTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPigTypeLookup.MyCombo, Di.PART_TYPE)
            uctlPartLookup.MyCombo.ListIndex = IDToListIndex(uctlPartLookup.MyCombo, Di.PART_ITEM_ID)
            uctlFeatureLookup.MyCombo.ListIndex = IDToListIndex(uctlFeatureLookup.MyCombo, Di.FEATURE_ID)
            txtPackAmount.Text = Di.PACK_AMOUNT
            txtPricePerPack.Text = Di.PRICE_PER_PACK
            txtDiscountPerPack.Text = Di.DISCOUNT_PER_PACK
            txtQuantity.Text = Di.ITEM_AMOUNT
            uctlToLocationLookup.MyCombo.ListIndex = IDToListIndex(uctlToLocationLookup.MyCombo, Di.LOCATION_ID)
            txtTotalPrice.Text = Di.TOTAL_PRICE
            txtAvgPrice.Text = Di.AVG_PRICE
            txtDiscount.Text = Di.DISCOUNT_AMOUNT
            txtManual.Text = Di.ITEM_DESC
            Call ShowDisplayID(Di.DISPLAY_ID)
            chkManualName.Value = FlagToCheck(Di.MANUAL_FLAG)
            txtManualCode.Text = Di.MANUAL_CODE
            txtManualName.Text = Di.MANUAL_NAME
            txtTransferWage.Text = Di.TRANSFER_WAGE
            txtStdTrfCharge.Text = Di.STD_TRANSFER_CHARGE
            
            cmdOK.Enabled = (COMMIT_FLAG <> "Y")
            Call CalculateCurrentBath
            
            Dim Tdi As CDoItem
            Dim RateOther As Double
            Set Tdi = New CDoItem
            Tdi.PACKAGE_RATE = Di.PACKAGE_RATE
            Tdi.RATE_CUSTOMER = Di.RATE_CUSTOMER
            Tdi.DISCOUNT_AMOUNT_PART = Di.DISCOUNT_AMOUNT_PART
            Tdi.DISCOUNT_AMOUNT_DLC = Di.DISCOUNT_AMOUNT_DLC
         
            Tdi.PRO_COMMISSION_BAG = Di.PRO_COMMISSION_BAG
            Tdi.PRO_CHEER_BAG = Di.PRO_CHEER_BAG
            Tdi.PRO_DST_BAG = Di.PRO_DST_BAG
            Tdi.PRO_OTHER1_BAG = Di.PRO_OTHER1_BAG
            Tdi.PRO_OTHER2_BAG = Di.PRO_OTHER2_BAG
            Tdi.PRO_OTHER3_BAG = Di.PRO_OTHER3_BAG
            Tdi.PRO_COMMISSION_KG = Di.PRO_COMMISSION_KG
            Tdi.PRO_CHEER_KG = Di.PRO_CHEER_KG
            Tdi.PRO_DST_KG = Di.PRO_DST_KG
            Tdi.PRO_OTHER1_KG = Di.PRO_OTHER1_KG
            Tdi.PRO_OTHER2_KG = Di.PRO_OTHER2_KG
            Tdi.PRO_OTHER3_KG = Di.PRO_OTHER3_KG
            
            
            
            Set TempData = New Collection
            Call TempData.add(Tdi, Trim(str(CUSTOMER_ID)))
            Call showDetail(TempData)
         End If
      End If
   End If
   
   Call EnableForm(Me, True)
End Sub

'Private Sub cmdLotSelect_Click()
'Dim OKClick As Boolean
'Dim LotAmount As Long
'
'   If Not VerifyCombo(lblPart, uctlPartLookup.MyCombo, False) Then
'      Exit Sub
'   End If
'   If Not VerifyCombo(lblToLocation, uctlToLocationLookup.MyCombo, False) Then
'      Exit Sub
'   End If
'
'   frmAddSubLotItem.HeaderText = "เลือกหมายเลขล็อต"
'   frmAddSubLotItem.ShowMode = ShowMode
''   If ShowMode = SHOW_ADD Then
'      Set frmAddSubLotItem.TempCollection = m_SubLotItems
''   Else
''      Call CopySubLotItem(TempCollection(ID).SubLotItems, m_SubLotItems)
''      Set frmAddSubLotItem.TempCollection = m_SubLotItems
''   End If
'   frmAddSubLotItem.PartItemID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
'   frmAddSubLotItem.LocationID = uctlToLocationLookup.MyCombo.ItemData(Minus2Zero(uctlToLocationLookup.MyCombo.ListIndex))
'   Load frmAddSubLotItem
'   frmAddSubLotItem.Show 1
'
'   OKClick = frmAddSubLotItem.OKClick
'   LotAmount = frmAddSubLotItem.SumLotAmount
'
'   Unload frmAddSubLotItem
'   Set frmAddSubLotItem = Nothing
'
'   If OKClick Then
'      m_ManualFlag = True
'      txtQuantity.Text = LotAmount
'      m_HasModify = True
'   End If
'End Sub

Private Sub cmdLastBill_Click()
Dim m_DoItemLastBill As Collection
Call EnableForm(Me, False)
Set m_DoItemLastBill = New Collection
   Call LoadDoItemByLastBilling(Nothing, m_DoItemLastBill, CUSTOMER_ID, PartItemID)
   If Not m_DoItemLastBill Is Nothing Then
      Set frmShowDoc.TempCollection = m_DoItemLastBill
      Load frmShowDoc
      frmShowDoc.Show 1
      OKClick = frmShowDoc.OKClick
      
      Unload frmShowDoc
      Set frmShowDoc = Nothing
   End If
   Call EnableForm(Me, True)
End Sub

Private Sub cmdOK_Click()
   If Not cmdOK.Enabled Then
      Exit Sub
   End If
   
   If Not SaveData Then
      Exit Sub
   End If
   
   If ShowMode = SHOW_EDIT Then 'ถ้ามีการเปลี่ยนแปลง สถานที่จัดส่งที่ form แม่ ให้ทำการ update ราคาใหม่อีกครั้ง
         If ISuctlDeliveryCusLookup Then
            ISuctlDeliveryCusLookup = False
         End If
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Function GetDisplayID() As Long
   If SSOption1.Value Then
      GetDisplayID = 1
   ElseIf SSOption2.Value Then
      GetDisplayID = 2
   ElseIf SSOption3.Value Then
      GetDisplayID = 3
   End If
End Function

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long
Dim OldPackAmount As Double

   If (Not radFeature.Value) And (Not radStock.Value) And (Not radCustom.Value) Then
      glbErrorLog.LocalErrorMsg = "กรุณากำหนดตัวเลือกอย่างใดอย่างหนึ่ง"
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   If Not VerifyTextControl(lblManual, txtManual, Not radCustom.Value) Then
      Exit Function
   End If
   If Not VerifyCombo(lblFeatureCode, uctlFeatureLookup.MyCombo, Not radFeature.Value) Then
      Exit Function
   End If
   If Not VerifyCombo(lblToLocation, uctlToLocationLookup.MyCombo, Not radStock.Value) Then
      Exit Function
   End If
   If Not VerifyCombo(lblPart, uctlPartLookup.MyCombo, Not radStock.Value) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblQuantity, txtQuantity, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblTotalPrice, txtTotalPrice, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblDiscount, txtDiscount, True) Then
      Exit Function
   End If
   
      Set TempD2 = GetObject("CCustomer", m_Customers, Trim(str(CUSTOMER_ID)), False) 'ตรวจสอบก่อนว่า ลูกค้าคนนี้สามารถขายสินค้าที่ตั้งราคาเป็น 0 ได้หรือไม่
      If Not TempD2 Is Nothing Then
         If TempD2.FREE_PRICE_FLAG = "N" Then
            If Val(txtPricePerPack.Text) <= 0 Then
               glbErrorLog.LocalErrorMsg = "ราคาสินค้าต้องไม่มีค่าเป็น 0"
               glbErrorLog.ShowUserError
               Exit Function
            End If
         End If
      End If


   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   If DocumentType = 18 Then
'      Dim Ri As CReceiptItem
'      If ShowMode = SHOW_ADD Then
'         Set Ri = New CReceiptItem
'
'         Ri.Flag = "A"
'         Call TempCollection.add(Ri)
'      Else
'         Set Ri = TempCollection.Item(ID)
'         If Ri.Flag <> "A" Then
'            Ri.Flag = "E"
'         End If
'      End If
'
'      If radStock.Value Then
'         Ri.PART_ITEM_ID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
'      Else
'         Ri.PART_ITEM_ID = -1
'      End If
'      Ri.PART_NO = uctlPartLookup.MyTextBox.Text
'      Ri.PART_DESC = uctlPartLookup.MyCombo.Text
'      Ri.RETURN_AMOUNT = txtQuantity.Text
'      Ri.LOCATION_ID = uctlToLocationLookup.MyCombo.ItemData(Minus2Zero(uctlToLocationLookup.MyCombo.ListIndex))
'      Ri.LOCATION_NAME = uctlToLocationLookup.MyCombo.Text
'      If m_Sp.PARAM_VALUE = "Y" Then
'         Ri.PART_TYPE = -1
'      Else
'         Ri.PART_TYPE = uctlPigTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPigTypeLookup.MyCombo.ListIndex))
'      End If
'      Ri.RETURN_TOTAL_PRICE = Val(Format(Val(txtTotalPrice.Text), "0.00")) 'แปลงให้เป็น 2 ตำแหน่งด้วย
'      Ri.AVG_PRICE = Val(txtAvgPrice.Text)
'      Ri.RETURN_AVG_PRICE = Val(txtReturnAvg.Text)
'      Ri.AVG_WEIGHT = 0
'      Ri.FEATURE_ID = uctlFeatureLookup.MyCombo.ItemData(Minus2Zero(uctlFeatureLookup.MyCombo.ListIndex))
'      Ri.FEATURE_CODE = uctlFeatureLookup.MyTextBox.Text
'      Ri.FEATURE_DESC = uctlFeatureLookup.MyCombo.Text
'      Ri.RETURN_DISCOUNT_AMOUNT = Val(Format(Val(txtDiscount.Text), "0.00"))  'แปลงให้เป็น 2 ตำแหน่งด้วย
'      Ri.CONFIG_CODE = CreateConfigFlag()
'      Ri.ITEM_DESC = txtManual.Text
'      Ri.DISPLAY_ID = GetDisplayID
'      Ri.COUNTRY_CURRENCY1 = COUNTRY_CURRENCY1
'      Ri.COUNTRY_CURRENCY2 = COUNTRY_CURRENCY2
'      Ri.WEIGHT_PER_PACK = Val(txtWeightPerPack.Text)
'      Ri.PACK_AMOUNT = Val(txtPackAmount.Text)
'      Ri.PRICE_PER_PACK = Val(txtPricePerPack.Text)
'      Ri.DISCOUNT_PER_PACK = Val(txtDiscountPerPack.Text)
'      Ri.MANUAL_FLAG = Check2Flag(chkManualName.Value)
'      Ri.MANUAL_CODE = txtManualCode.Text
'      Ri.MANUAL_NAME = txtManualName.Text
'      Ri.TRANSFER_WAGE = Val(txtTransferWage.Text)
'      Ri.STD_TRANSFER_CHARGE = Val(Format(Val(txtStdTrfCharge.Text), "0.00")) 'แปลงให้เป็น 2 ตำแหน่งด้วย

   
   Else
      Dim Di As CSaleOrder
      If ShowMode = SHOW_ADD Then
         Set Di = New CSaleOrder
   
         Di.Flag = "A"
         Call TempCollection.add(Di)
      Else
         Set Di = TempCollection.Item(ID)
         If Di.Flag <> "A" Then
            Di.Flag = "E"
         End If
      End If
   
      If radStock.Value Then
         Di.PART_ITEM_ID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
      Else
         Di.PART_ITEM_ID = -1
      End If
      Di.PART_NO = uctlPartLookup.MyTextBox.Text
      Di.PART_DESC = uctlPartLookup.MyCombo.Text
      Di.ITEM_AMOUNT = txtQuantity.Text
      Di.LOCATION_ID = uctlToLocationLookup.MyCombo.ItemData(Minus2Zero(uctlToLocationLookup.MyCombo.ListIndex))
      Di.LOCATION_NAME = uctlToLocationLookup.MyCombo.Text
      Di.PART_TYPE_NAME = uctlPigTypeLookup.MyCombo.Text
      If m_Sp.PARAM_VALUE = "Y" Then
         Di.PART_TYPE = -1
      Else
         Di.PART_TYPE = uctlPigTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPigTypeLookup.MyCombo.ListIndex))
      End If
      Di.TOTAL_PRICE = Val(Format(Val(txtTotalPrice.Text), "0.00")) 'แปลงให้เป็น 2 ตำแหน่งด้วย
      Di.AVG_PRICE = Val(txtAvgPrice.Text)
      Di.AVG_WEIGHT = 0
      Di.FEATURE_ID = uctlFeatureLookup.MyCombo.ItemData(Minus2Zero(uctlFeatureLookup.MyCombo.ListIndex))
      Di.FEATURE_CODE = uctlFeatureLookup.MyTextBox.Text
      Di.FEATURE_DESC = uctlFeatureLookup.MyCombo.Text
      Di.DISCOUNT_AMOUNT = Val(Format(Val(txtDiscount.Text), "0.00"))  'แปลงให้เป็น 2 ตำแหน่งด้วย
      Di.CONFIG_CODE = CreateConfigFlag()
      Di.ITEM_DESC = txtManual.Text
      Di.DISPLAY_ID = GetDisplayID
      Di.COUNTRY_CURRENCY1 = COUNTRY_CURRENCY1
      Di.COUNTRY_CURRENCY2 = COUNTRY_CURRENCY2
      Di.WEIGHT_PER_PACK = Val(txtWeightPerPack.Text)
      OldPackAmount = Di.PACK_AMOUNT
      Di.PACK_AMOUNT = Val(txtPackAmount.Text)
      Di.PRICE_PER_PACK = Val(txtPricePerPack.Text)
      Di.DISCOUNT_PER_PACK = Val(txtDiscountPerPack.Text)
      Di.MANUAL_FLAG = Check2Flag(chkManualName.Value)
      Di.MANUAL_CODE = txtManualCode.Text
      Di.MANUAL_NAME = txtManualName.Text
      Di.TRANSFER_WAGE = Val(txtTransferWage.Text)
      Di.STD_TRANSFER_CHARGE = Val(Format(Val(txtStdTrfCharge.Text), "0.00")) 'แปลงให้เป็น 2 ตำแหน่งด้วย
      NewUpdatePrice = chkUpdatePrice.Value
      
    Set TempD3 = GetObject("CDoItem", TempData, Trim(str(CUSTOMER_ID)), False)
      If Not TempD3 Is Nothing Then
         Di.EX_WORKS_PRICE_ITEM_ID = TempD3.EX_WORKS_PRICE_ITEM_ID
         Di.PACKAGE_RATE = TempD3.PACKAGE_RATE
         Di.EX_DELIVERY_COST_ITEM_ID = TempD3.EX_DELIVERY_COST_ITEM_ID
         Di.RATE_CUSTOMER = TempD3.RATE_CUSTOMER
         Di.PRO_COMMISSION_BAG = TempD3.PRO_COMMISSION_BAG
         Di.PRO_CHEER_BAG = TempD3.PRO_CHEER_BAG
         Di.PRO_DST_BAG = TempD3.PRO_DST_BAG
         Di.PRO_OTHER1_BAG = TempD3.PRO_OTHER1_BAG
         Di.PRO_OTHER2_BAG = TempD3.PRO_OTHER2_BAG
         Di.PRO_OTHER3_BAG = TempD3.PRO_OTHER3_BAG
         Di.PRO_COMMISSION_KG = TempD3.PRO_COMMISSION_KG
         Di.PRO_CHEER_KG = TempD3.PRO_CHEER_KG
         Di.PRO_DST_KG = TempD3.PRO_DST_KG
         Di.PRO_OTHER1_KG = TempD3.PRO_OTHER1_KG
         Di.PRO_OTHER2_KG = TempD3.PRO_OTHER2_KG
         Di.PRO_OTHER3_KG = TempD3.PRO_OTHER3_KG
         
         Di.EX_PROMOTION_PART_ITEM_ID = TempD3.EX_PROMOTION_PART_ITEM_ID
         Di.EX_PROMOTION_DLC_ITEM_ID = TempD3.EX_PROMOTION_DLC_ITEM_ID
         
         TempD3.SUM_RATE_OTHER_BAG = TempD3.PRO_COMMISSION_BAG + TempD3.PRO_CHEER_BAG + TempD3.PRO_DST_BAG + TempD3.PRO_OTHER1_BAG + TempD3.PRO_OTHER2_BAG + TempD3.PRO_OTHER3_BAG
         TempD3.SUM_RATE_OTHER_KG = TempD3.PRO_COMMISSION_KG + TempD3.PRO_CHEER_KG + TempD3.PRO_DST_KG + TempD3.PRO_OTHER1_KG + TempD3.PRO_OTHER2_KG + TempD3.PRO_OTHER3_KG
      Dim SumOther As Double
      Dim tRate As Double
      If Di.PART_TYPE = 10 Then
          SumOther = TempD3.SUM_RATE_OTHER_BAG
      ElseIf Di.PART_TYPE = 21 Then
         SumOther = TempD3.SUM_RATE_OTHER_KG
      End If
      If PRICE_THINK_TYPE = 1 Then 'มารับเอง
         tRate = TempD3.PACKAGE_RATE + SumOther
            If Val(txtPricePerPack.Text) >= tRate Then
                Di.DISCOUNT_AMOUNT_PART = Di.DISCOUNT_PER_PACK
            Else
                  Di.DISCOUNT_AMOUNT_PART = Di.DISCOUNT_PER_PACK + (tRate - Val(txtPricePerPack.Text))
            End If
      ElseIf PRICE_THINK_TYPE = 2 Then 'รวมค่าขนส่ง
          tRate = TempD3.PACKAGE_RATE + TempD3.RATE_CUSTOMER + SumOther - TempD3.DISCOUNT_AMOUNT_DLC
            If radStock.Value Then
               If Val(txtPricePerPack.Text) >= tRate Then
                   Di.DISCOUNT_AMOUNT_PART = Di.DISCOUNT_PER_PACK
                   Di.DISCOUNT_AMOUNT_DLC = TempD3.DISCOUNT_AMOUNT_DLC
               Else
                  Di.DISCOUNT_AMOUNT_PART = Di.DISCOUNT_PER_PACK + (tRate - Val(txtPricePerPack.Text))
                  Di.DISCOUNT_AMOUNT_DLC = TempD3.DISCOUNT_AMOUNT_DLC
               End If
            End If
         ElseIf PRICE_THINK_TYPE = 3 Then 'คิดแยกค่าขนส่ง
            If radFeature.Value Then
               If Val(txtPricePerPack.Text) >= Di.RATE_CUSTOMER Then
                   Di.DISCOUNT_AMOUNT_DLC = Di.DISCOUNT_PER_PACK
               Else
                  Di.DISCOUNT_AMOUNT_DLC = Di.DISCOUNT_PER_PACK + (Di.RATE_CUSTOMER - Val(txtPricePerPack.Text))
               End If
               'เอาส่วนลดค่าขนส่งไปเก็บที่รายการของสินค้าด้วย
               Dim Tdi As CSaleOrder
               For Each Tdi In TempCollection
                 If Tdi.PART_ITEM_ID > 0 And (Tdi.WEIGHT_PER_PACK = Di.WEIGHT_PER_PACK) Then
                     If Tdi.Flag <> "A" Then
                        Tdi.Flag = "E"
                     End If
                     Tdi.DISCOUNT_AMOUNT_DLC = Di.DISCOUNT_AMOUNT_DLC
                 End If
               Next Tdi
            ElseIf radStock.Value Then
             tRate = TempD3.PACKAGE_RATE + SumOther
               If Val(txtPricePerPack.Text) >= tRate Then
                   Di.DISCOUNT_AMOUNT_PART = Di.DISCOUNT_PER_PACK
               Else
                     Di.DISCOUNT_AMOUNT_PART = Di.DISCOUNT_PER_PACK + (tRate - Val(txtPricePerPack.Text))
               End If
            End If
         End If
         
'          If CurrentPricePerPack > NewPricePerPack And EditPriceFlag Then
'            If radFeature.Value Then
'               Di.DISCOUNT_AMOUNT_DLC = TempD3.DISCOUNT_AMOUNT_DLC + (Di.RATE_CUSTOMER - NewPricePerPack)
'               Dim Tdi As CSaleOrder
'               For Each Tdi In TempCollection
'                 If Tdi.PART_ITEM_ID > 0 And (Tdi.WEIGHT_PER_PACK = Di.WEIGHT_PER_PACK) Then
'                     If Tdi.Flag <> "A" Then
'                        Tdi.Flag = "E"
'                     End If
'                     Tdi.DISCOUNT_AMOUNT_DLC = Di.DISCOUNT_AMOUNT_DLC
'                 End If
'               Next Tdi
'            ElseIf radStock.Value Then
'               Di.DISCOUNT_AMOUNT_PART = Di.DISCOUNT_PER_PACK + (Di.PACKAGE_RATE - NewPricePerPack) 'TempD3.DISCOUNT_AMOUNT_PART +
'               Di.DISCOUNT_AMOUNT_DLC = TempD3.DISCOUNT_AMOUNT_DLC
'            End If
'         Else
'            If NewPricePerPack >= Di.PACKAGE_RATE Then
'               Di.DISCOUNT_AMOUNT_PART = 0
'               Di.DISCOUNT_AMOUNT_DLC = 0
'            Else
'               'Di.DISCOUNT_AMOUNT_PART = TempD3.DISCOUNT_AMOUNT_PART
'                Di.DISCOUNT_AMOUNT_PART = Di.DISCOUNT_PER_PACK + (Di.PACKAGE_RATE - Val(txtPricePerPack.Text))
'               Di.DISCOUNT_AMOUNT_DLC = TempD3.DISCOUNT_AMOUNT_DLC
'            End If
'
'         End If

         If (PRICE_THINK_TYPE = 1) Then
            Di.EX_DELIVERY_COST_ITEM_ID = -1
            Di.RATE_CUSTOMER = 0
            Di.EX_PROMOTION_DLC_ITEM_ID = -1
            Di.DISCOUNT_AMOUNT_DLC = 0
         End If
      End If

      If EditPriceFlag Then
         Di.USER_APPLOVE_PRICE = TempUserName
      End If
      
      Dim Di2 As CSaleOrder
      Dim TempDI As CSaleOrder
      
'      If TypeSale = 1 Or TypeSale = 3 Then 'ถ้าเป็นการแก้ไข ให้เข้าไปสร้าง key ของ ค่าขนส่งก่อน
      If TypeSale = 1 Then  'เข้าไปสร้าง key ของ ค่าขนส่งก่อน
         Dim TempCollection2 As Collection
            Set TempCollection2 = New Collection
            For Each TempDI In TempCollection
               Call TempCollection2.add(TempDI)
            Next TempDI
            
            For Each TempDI In TempCollection
               Call TempCollection.Remove(1)
            Next TempDI
            
            For Each TempDI In TempCollection2
               If TempDI.PART_ITEM_ID = -1 Then
                  Call TempCollection.add(TempDI, Trim(TempDI.PART_ITEM_ID) & "-" & Trim(str(TempDI.WEIGHT_PER_PACK)))
               Else
                   Call TempCollection.add(TempDI)
               End If
            Next TempDI
      End If
      
'      If (PRICE_THINK_TYPE = 3 And TypeSale = 1) Or (PRICE_THINK_TYPE = 3 And TypeSale = 3) Then 'สร้างค่าขนส่งอัตโนมัติ เมื่อเป็นการคิดแบบแยกค่าขนส่ง
      If (PRICE_THINK_TYPE = 3 And TypeSale = 1) Then  'สร้างค่าขนส่งอัตโนมัติ เมื่อเป็นการคิดแบบแยกค่าขนส่ง
          If Di.PART_ITEM_ID > 0 Then '
               Set Di2 = GetObject("CSaleOrder", TempCollection, Trim("-1") & "-" & Trim(str(Di.WEIGHT_PER_PACK)), False)
                  If Di2 Is Nothing Then
                     Set Di2 = New CSaleOrder
                     Di2.Flag = "A"
                     Call TempCollection.add(Di2, Trim("-1") & "-" & Trim(str(Di.WEIGHT_PER_PACK)))
                  Else 'ถ้ามีแล้ว
                      If Di2.Flag <> "A" Then
                        Di2.Flag = "E"
                      End If
                  End If
      
               Di2.PART_ITEM_ID = -1
               Di2.PART_NO = ""
               Di2.PART_DESC = ""
               
               Di2.LOCATION_ID = 0
               Di2.LOCATION_NAME = ""
               Di2.PART_TYPE_NAME = ""
               Di2.PART_TYPE = 0

               Di2.FEATURE_ID = 44
               Di2.FEATURE_CODE = "TR-0001"
               Di2.FEATURE_DESC = "ค่าขนส่งสินค้า"
               
               Di2.CONFIG_CODE = "YNN"
               Di2.ITEM_DESC = ""
               Di2.DISPLAY_ID = 2
               Di2.COUNTRY_CURRENCY1 = 0
               Di2.COUNTRY_CURRENCY2 = 0
               Di2.WEIGHT_PER_PACK = Val(txtWeightPerPack.Text)
               Di2.PACK_AMOUNT = Di2.PACK_AMOUNT + (Val(txtPackAmount.Text) - OldPackAmount)


         Dim EX_DELIVERY_COST_ITEM_ID As Long
         Dim EX_PROMOTION_DLC_ITEM_ID As Long
         Dim RateDeliveryCostFlag As Boolean
         Dim RatePromotionDlcFlag As Boolean
         Dim RateDeliveryCost As Double
         Dim RatePromotionDlc As Double
         Dim PartType As Long
         Set TempData = New Collection
         If Val(txtWeightPerPack.Text) = 1 Then 'set type เป็น Bulk
            PartType = 21
         ElseIf Val(txtWeightPerPack.Text) > 0 Then 'set type ถุง
           PartType = 10
         End If
         Call calExDeliveryCost(DELIVERY_CUS_ITEM_ID, CUSTOMER_ID, Val(txtWeightPerPack.Text), PartType, PRICE_THINK_TYPE, m_ExDeliveryCostItem, m_Customers, RateDeliveryCost, RateDeliveryCostFlag, EX_DELIVERY_COST_ITEM_ID)
         Call calExPromotionDls(DELIVERY_CUS_ITEM_ID, CUSTOMER_ID, Val(txtWeightPerPack.Text), PartType, PRICE_THINK_TYPE, m_ExPromotionDlcItem, m_Customers, RatePromotionDlc, RatePromotionDlcFlag, EX_PROMOTION_DLC_ITEM_ID)
               
               Di2.DISCOUNT_PER_PACK = RatePromotionDlc 'Val(txtDiscountPerPack.Text)
               Di2.PRICE_PER_PACK = RateDeliveryCost '- RatePromotionDlc '- Di2.DISCOUNT_PER_PACK
              Di2.DISCOUNT_AMOUNT = RatePromotionDlc * Di2.PACK_AMOUNT
                
               Di2.ITEM_AMOUNT = Di2.WEIGHT_PER_PACK * Di2.PACK_AMOUNT
               If Di.PART_TYPE = 10 Then
                  Di2.TOTAL_PRICE = Val(Format(Val(((Di2.PRICE_PER_PACK) * Di2.PACK_AMOUNT) - Di2.DISCOUNT_AMOUNT), "0.00"))  'แปลงให้เป็น 2 ตำแหน่งด้วย
               ElseIf Di.PART_TYPE = 21 Then
                   Di2.TOTAL_PRICE = Val(Format(Val(((Di2.PRICE_PER_PACK) * Di2.ITEM_AMOUNT) - Di2.DISCOUNT_AMOUNT), "0.00"))   'แปลงให้เป็น 2 ตำแหน่งด้วย
               End If

               Di2.AVG_PRICE = Val(MyDiffEx(Di2.PACK_AMOUNT * Di2.PRICE_PER_PACK, Di2.PACK_AMOUNT * Di2.WEIGHT_PER_PACK))
               Di2.AVG_WEIGHT = 0
               Di2.MANUAL_FLAG = "N"
               Di2.MANUAL_CODE = ""
               Di2.MANUAL_NAME = ""
               Di2.TRANSFER_WAGE = 0
               Di2.STD_TRANSFER_CHARGE = 0

               Di2.EX_WORKS_PRICE_ITEM_ID = -1
               Di2.PACKAGE_RATE = 0
               Di2.EX_DELIVERY_COST_ITEM_ID = EX_DELIVERY_COST_ITEM_ID
               Di2.RATE_CUSTOMER = RateDeliveryCost
               Di2.PRO_COMMISSION_BAG = 0
               Di2.PRO_CHEER_BAG = 0
               Di2.PRO_DST_BAG = 0
               Di2.PRO_OTHER1_BAG = 0
               Di2.PRO_OTHER2_BAG = 0
               Di2.PRO_OTHER3_BAG = 0
               Di2.PRO_COMMISSION_KG = 0
               Di2.PRO_CHEER_KG = 0
               Di2.PRO_DST_KG = 0
               Di2.PRO_OTHER1_KG = 0
               Di2.PRO_OTHER2_KG = 0
               Di2.PRO_OTHER3_KG = 0
               
               Di2.EX_PROMOTION_PART_ITEM_ID = -1
               Di2.EX_PROMOTION_DLC_ITEM_ID = EX_PROMOTION_DLC_ITEM_ID
               Di2.DISCOUNT_AMOUNT_PART = 0
               Di2.DISCOUNT_AMOUNT_DLC = RatePromotionDlc
            End If
      ElseIf PRICE_THINK_TYPE = 2 Then 'ตรวจสอบดู หากมีเหตุการณ์ เปลี่ยนวิธีการคิด เป็นแบบ รวมค่าขนส่ง หากเดิม ได้มีการสร้างค่าขนส่ง เอาไว้ ก็ให้สั่งลบด้วย
          Set Di2 = GetObject("CSaleOrder", TempCollection, Trim("-1") & "-" & Trim(str(Di.WEIGHT_PER_PACK)), False)
                  If Not Di2 Is Nothing Then
                      If Di2.Flag <> "A" Then
                        Di2.Flag = "D"
                      End If
                  End If
      End If
   End If
   
   SaveData = True
End Function

Private Sub cmdUpdatePrice_Click()
Dim ExPromotionPart As Double
Dim ExPromotionDlc As Double
If TypeSale = 1 Then
  If radStock.Value Then
   If Not VerifyCombo(lblPart, uctlPartLookup.MyCombo, False) Then
      Exit Sub
   End If
End If
If radFeature.Value Then
   Exit Sub
End If
 txtPricePerPack.Text = calExWorksPrice(Pi, DELIVERY_CUS_ITEM_ID, CUSTOMER_ID, PRICE_THINK_TYPE, m_ExWorkPricesItem, m_ExDeliveryCostItem, m_Customers, TempData, m_ExPromotionPartItem, m_ExPromotionDlcItem, ExPromotionPart, ExPromotionDlc)
If radFeature.Value Then
  txtDiscountPerPack.Text = ExPromotionDlc
ElseIf radStock.Value Then
  txtDiscountPerPack.Text = ExPromotionPart
End If
Call showDetail(TempData)
  ' เครียร์สถานะการอนุมัติออกด้วยหากมีการคิดราคาอาหารจากราคาประกาศ
  TempUserName = ""
   EditPriceFlag = True
End If
    m_HasModify = True
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
'      Call InitDoRateType(cboRateType)
      
      Call LoadFeature(uctlFeatureLookup.MyCombo, m_Features, , , "Y")
      Set uctlFeatureLookup.MyCollection = m_Features
      
      If m_Sp.PARAM_VALUE = "Y" Then
         Call LoadPartItem(uctlPartLookup.MyCombo, m_Pigs)
         Set uctlPartLookup.MyCollection = m_Pigs
      Else
         Call LoadPartType(uctlPigTypeLookup.MyCombo, m_PartTypes)
         Set uctlPigTypeLookup.MyCollection = m_PartTypes
      End If
      
      Call LoadLocation(uctlToLocationLookup.MyCombo, m_Houses, 2)
      Set uctlToLocationLookup.MyCollection = m_Houses

      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
         radFeature.Value = True
         Call QueryData(True)
         
         Call DisableForReturn
         Call ShowGui
      End If
      
      m_HasModify = False
      
      
       If ShowMode = SHOW_EDIT Then 'ถ้ามีการเปลี่ยนแปลง สถานที่จัดส่งที่ form แม่ ให้ทำการ update ราคาใหม่อีกครั้ง
         If ISuctlDeliveryCusLookup Then
            Call cmdUpdatePrice_Click
         End If
      End If
      
      CurrentPricePerPack = Val(txtPricePerPack.Text)
      
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
   Set m_Sp = GetSystemParam(glbSystemParams, "BARCODE_FLAG")
m_Sp.PARAM_VALUE = "N"
   OKClick = False
   Call InitFormLayout
   
   m_ManualFlag = False
   m_HasActivate = False
   m_HasModify = False
   
   Set m_Rs = New ADODB.Recordset
   Set m_PartTypes = New Collection
   Set m_Parts = New Collection
   Set m_Locations = New Collection
   Set m_Houses = New Collection
   Set m_Pigs = New Collection
   Set m_PigStatuss = New Collection
   Set m_PartTypes = New Collection
   Set m_SubLotItems = New Collection
   Set m_Features = New Collection
End Sub

Private Sub SSCommand2_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   
   Set m_Rs = Nothing
   Set m_PartTypes = Nothing
   Set m_Parts = Nothing
   Set m_Locations = Nothing
   Set m_Houses = Nothing
   Set m_Pigs = Nothing
   Set m_PigStatuss = Nothing
   Set m_PartTypes = Nothing
   Set m_SubLotItems = Nothing
   Set m_Features = Nothing
End Sub

Private Sub txtDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtKeyName_Change()
   m_HasModify = True
End Sub

Private Sub txtThaiMsg_Change()
   m_HasModify = True
End Sub

Private Sub txtAmphur_Change()
   m_HasModify = True
End Sub

Private Sub txtDistrict_Change()
   m_HasModify = True
End Sub

Private Sub txtFax_Change()
   m_HasModify = True
End Sub

Private Sub txtHomeNo_Change()
   m_HasModify = True
End Sub

Private Sub radCustom_Click(Value As Integer)
   m_HasModify = True
   Call ShowGui
   chkManualName.Value = FlagToCheck("N")
   Call QueryData(True)
   
   txtWeightPerPack.Enabled = True
   txtPackAmount.Enabled = True
   txtPricePerPack.Enabled = True
   txtDiscountPerPack.Enabled = True
End Sub

Private Sub radFeature_Click(Value As Integer)
   m_HasModify = True
   Call ShowGui
   chkManualName.Value = FlagToCheck("N")
   Call QueryData(True)
End Sub

Private Sub radFeature_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub radStock_Click(Value As Integer)
   m_HasModify = True
   Call ShowGui
   chkManualName.Value = FlagToCheck("Y")
   Call QueryData(True)
End Sub

Private Sub radStock_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub



Private Sub SSOption1_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub SSOption1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub SSOption2_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub SSOption2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub SSOption3_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub SSOption3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub txtAvgPrice_Change()
   m_HasModify = True
   txtExcludeDiscount.Text = Val(txtAvgPrice.Text) * Val(txtQuantity.Text)
End Sub

Private Sub txtAvgPrice_GotFocus()
'Dim Ug As CDoItem
'Dim IsOK As Boolean
'Dim iCount As Long
'
'   If Area <> 1 Then
'      Exit Sub
'   End If
'
'   If radCustom.Value Then
'      Exit Sub
'   End If
'
''   If uctlFeatureLookup.MyCombo.ListIndex <= 0 Then
''      Exit Sub
''   End If
'
'   Call EnableForm(Me, False)
'
'   Set Ug = New CDoItem
'   m_SocID = -1
'   If True Then
'      If txtAvgPrice.Text = "" Then
'         Ug.ITEM_AMOUNT = Val(txtQuantity.Text)
'         If uctlFeatureLookup.MyCombo.ListIndex > 0 Then
'            Ug.FEATURE_ID = uctlFeatureLookup.MyCombo.ItemData(Minus2Zero(uctlFeatureLookup.MyCombo.ListIndex))
'         Else
'            Ug.FEATURE_ID = -1
'         End If
'         If uctlPartLookup.MyCombo.ListIndex > 0 Then
'            Ug.PART_ITEM_ID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
'         Else
'            Ug.PART_ITEM_ID = -1
'         End If
'
'         Ug.ACCOUNT_ID = AccountID
'         Ug.SUBSCRIBER_ID = SubscriberID
'         Ug.USAGE_DATE = UsageDate
'
'         If glbDaily.CalculatePrice(Ug, Nothing, False, 2, IsOK, glbErrorLog) Then
'            txtAvgPrice.Text = Ug.AVG_PRICE
'            txtTotalPrice.Text = Ug.UC_AMOUNT + Ug.AC_AMOUNT
'
'            m_SocID = Ug.SOC_ID
'         End If
'      End If
'   End If
'   Set Ug = Nothing
'   Call EnableForm(Me, True)
End Sub

Private Sub txtDeposit_Change()
   m_HasModify = True
   
End Sub

Private Sub txtDiscount_Change()
   m_HasModify = True
   txtTotalPrice.Text = Val(txtExcludeDiscount.Text) - Val(txtDiscount.Text)
End Sub

Private Sub txtLeft_Change()
   m_HasModify = True
End Sub

Private Sub txtDiscountPerPack_Change()
   m_HasModify = True
   txtDiscount.Text = Val(txtDiscountPerPack.Text) * Val(txtPackAmount.Text)
End Sub

Private Sub txtDiscountPerPack_GotFocus()
Dim Ug As CDoItem
Dim IsOK As Boolean
Dim iCount As Long
Dim RateType As DO_RATE_TYPE
Dim Ft As CFeature
Dim TempID As Long

   If Area <> 1 Then
      Exit Sub
   End If

   If radCustom.Value Then
      Exit Sub
   End If

'   RateType = cboRateType.ItemData(Minus2Zero(cboRateType.ListIndex))
'   If RateType = RATE_CUSTOM Then
'      Exit Sub
'   End If
   

   Call EnableForm(Me, False)

   Set Ug = New CDoItem
   m_SocID = -1
   If True Then
      If txtDiscountPerPack.Text = "" Then
         Ug.ITEM_AMOUNT = Val(txtPackAmount.Text)
         If uctlFeatureLookup.MyCombo.ListIndex > 0 Then
            Ug.FEATURE_ID = uctlFeatureLookup.MyCombo.ItemData(Minus2Zero(uctlFeatureLookup.MyCombo.ListIndex))
         Else
            Ug.FEATURE_ID = -1
         End If
         If uctlPartLookup.MyCombo.ListIndex > 0 Then
            Ug.PART_ITEM_ID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
         Else
            Ug.PART_ITEM_ID = -1
         End If
         
         Ug.ACCOUNT_ID = AccountID
         Ug.SUBSCRIBER_ID = SubscriberID
         Ug.USAGE_DATE = UsageDate
         
         If glbDaily.CalculatePrice(Ug, Nothing, False, 1, "N", IsOK, glbErrorLog) Then
            txtDiscountPerPack.Text = Val(txtPricePerPack.Text) - Ug.AVG_PRICE

            m_SocID = Ug.SOC_ID
         End If
      End If
   End If
   
   TempID = uctlFeatureLookup.MyCombo.ItemData(Minus2Zero(uctlFeatureLookup.MyCombo.ListIndex))
   Set Ft = GetFeature(m_Features, Trim(str(TempID)))
   If Ft.LOGISTIC_FLAG = "Y" Then
      txtStdTrfCharge.Text = Val(txtExcludeDiscount.Text)
   End If
   
   Set Ug = Nothing
   Call EnableForm(Me, True)
End Sub

Private Sub txtExcludeDiscount_Change()
   m_HasModify = True
   txtTotalPrice.Text = Val(txtExcludeDiscount.Text) - Val(txtDiscount.Text)
End Sub

Private Sub txtManual_Change()
   m_HasModify = True
End Sub

Private Sub txtNetTotal_Change()
   m_HasModify = True

End Sub

Private Sub txtPack_Change()
   m_HasModify = True
End Sub

Private Sub txtQtyp_Change()
   m_HasModify = True
End Sub

Private Sub txtManualCode_Change()
   m_HasModify = True
End Sub

Private Sub txtManualName_Change()
   m_HasModify = True
End Sub

Private Sub txtPackAmount_Change()
   m_HasModify = True
   txtQuantity.Text = Val(txtWeightPerPack.Text) * Val(txtPackAmount.Text)
   txtDiscount.Text = Val(txtDiscountPerPack.Text) * Val(txtPackAmount.Text)
End Sub

Private Sub txtPricePerPack_Change()
   m_HasModify = True
   txtAvgPrice.Text = MyDiffEx(Val(txtPackAmount.Text) * Val(txtPricePerPack.Text), Val(txtQuantity.Text))
End Sub

Private Sub txtPricePerPack_GotFocus()
Dim Ug As CDoItem
Dim IsOK As Boolean
Dim iCount As Long
Dim SocLevel As String
Dim RateType As DO_RATE_TYPE
Dim TempID As Long
Dim Ft As CFeature

   If Area <> 1 Then
      Exit Sub
   End If

   If radCustom.Value Then
      Exit Sub
   End If

   Call EnableForm(Me, False)

   Set Ug = New CDoItem
   m_SocID = -1
   If True Then
      If txtPricePerPack.Text = "" Then
         Ug.ITEM_AMOUNT = Val(txtPackAmount.Text)
         If uctlFeatureLookup.MyCombo.ListIndex > 0 Then
            Ug.FEATURE_ID = uctlFeatureLookup.MyCombo.ItemData(Minus2Zero(uctlFeatureLookup.MyCombo.ListIndex))
         Else
            Ug.FEATURE_ID = -1
         End If
         If uctlPartLookup.MyCombo.ListIndex > 0 Then
            Ug.PART_ITEM_ID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
         Else
            Ug.PART_ITEM_ID = -1
         End If
         
         Ug.ACCOUNT_ID = AccountID
         Ug.SUBSCRIBER_ID = SubscriberID
         Ug.USAGE_DATE = UsageDate
         
         If glbDaily.CalculatePrice(Ug, Nothing, False, 1, SocLevel, IsOK, glbErrorLog) Then
            txtPricePerPack.Text = Ug.AVG_PRICE
            txtTotalPrice.Text = Ug.UC_AMOUNT + Ug.AC_AMOUNT

            m_SocID = Ug.SOC_ID
         End If
      End If
   End If
   
   TempID = uctlFeatureLookup.MyCombo.ItemData(Minus2Zero(uctlFeatureLookup.MyCombo.ListIndex))
   Set Ft = GetFeature(m_Features, Trim(str(TempID)))
   If Ft.LOGISTIC_FLAG = "Y" Then
      txtStdTrfCharge.Text = Val(txtExcludeDiscount.Text)
   End If
   
   Set Ug = Nothing
   Call EnableForm(Me, True)
End Sub

Private Sub txtPricePerPack_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      If Not EditPriceFlag Then
         Call CheckPricePerPack
     End If
  End If
End Sub
Private Sub CheckPricePerPack()
   If radFeature.Value And CAL_PRICE_DLC_CENTER_FLAG = "N" Then  'ถ้าลูกค้าคนนี้ ไม่คิดราคาค่าขนส่งจากส่วนกลาง ก็ให้ออกจากการตรวจสอบไปเลย
      Exit Sub
   End If
   
   If radStock.Value And CAL_PRICE_PART_CENTER_FLAG = "N" Then  'ถ้าลูกค้าคนนี้ ไม่คิดราคาสินค้าจากส่วนกลาง ก็ให้ออกจากการตรวจสอบไปเลย
      Exit Sub
   End If
   
   If Val(txtPricePerPack.Text) < CurrentPricePerPack Then
       txtPricePerPack.Text = CurrentPricePerPack
      glbErrorLog.LocalErrorMsg = "ราคาใหม่ที่ระบุต้องมากกว่าราคาเดิม"
      glbErrorLog.ShowUserError
       Exit Sub
    ElseIf CurrentPricePerPack <= 0 And PRICE_THINK_TYPE > 0 Then
       txtPricePerPack.Text = CurrentPricePerPack
      glbErrorLog.LocalErrorMsg = "ยังไม่มีการกำหนดราคา กรุณากำหนดราคา ของรายการนี้ก่อน"
      glbErrorLog.ShowUserError
       Exit Sub
    Else
     CurrentPricePerPack = NewPricePerPack 'Val(txtPricePerPack.Text)
    End If
End Sub
Private Sub txtPricePerPack_LostFocus()
Dim Temp As Double
   If TypeSale <> 1 Then
     Exit Sub
   End If
   
  If Not EditPriceFlag Then
    Call CheckPricePerPack
  End If
  If SuccessFlag = "C" Then
    chkUpdatePrice.Value = ssCBChecked
  End If
  NewPricePerPack = Val(txtPricePerPack.Text)
End Sub

Private Sub txtQuantity_Change()
   m_HasModify = True
   txtAvgPrice.Text = MyDiffEx(Val(txtPackAmount.Text) * Val(txtPricePerPack.Text), Val(txtQuantity.Text))
   txtExcludeDiscount.Text = Val(txtAvgPrice.Text) * Val(txtQuantity.Text)
End Sub

Private Sub txtPhone_Change()
   m_HasModify = True
End Sub

Private Sub txtProvince_Change()
   m_HasModify = True
End Sub

Private Sub txtRoad_Change()
   m_HasModify = True
End Sub

Private Sub txtSoi_Change()
   m_HasModify = True
End Sub

Private Sub txtPrice_Change()
   m_HasModify = True
End Sub

Private Sub txtZipcode_Change()
   m_HasModify = True
End Sub
Private Sub txtReturnAvg_Change()
   m_HasModify = True
End Sub

Private Sub txtStdTrfCharge_Change()
   m_HasModify = True
End Sub

Private Sub uctlPigStatusLookup_Change()
   m_HasModify = True
End Sub

Private Sub txtTransferWage_Change()
   m_HasModify = True
End Sub

Private Sub txtWeightPerPack_Change()
   m_HasModify = True
   txtQuantity.Text = Val(txtWeightPerPack.Text) * Val(txtPackAmount.Text)
End Sub

Private Sub uctlFromPeriod_HasChange()
   m_HasModify = True
End Sub
Private Sub verifyDelivery(TempID As Long)
Dim RateDeliveryCost As Double
Dim RatePromotionDlc As Double

Dim tempRate As CDoItem
Set tempRate = New CDoItem
     If TempID = 44 Then
       If TypeSale = 1 Then
         If Not PRICE_THINK_TYPE = 3 Then
            glbErrorLog.LocalErrorMsg = "ลูกค้ารายนี้ไม่แยก ค่าขนส่ง"
            glbErrorLog.ShowUserError
            Exit Sub
          End If
         End If
          
          If Len(txtWeightPerPack.Text) = 0 Then
            txtWeightPerPack.Text = "1"
          End If
   
         If Not VerifyTextControl(lblWeightPerPack, txtWeightPerPack, False) Then
            Exit Sub
        End If
        
         Dim EX_DELIVERY_COST_ITEM_ID As Long
         Dim EX_PROMOTION_DLC_ITEM_ID As Long
         Dim RateDeliveryCostFlag As Boolean
         Dim RatePromotionDlcFlag As Boolean
         Dim PartType As Long
         Set TempData = New Collection
         
         If Val(txtWeightPerPack.Text) = 1 Then 'set type เป็น Bulk
            PartType = 21
         ElseIf Val(txtWeightPerPack.Text) > 0 Then 'set type ถุง
           PartType = 10
         End If
         Call calExDeliveryCost(DELIVERY_CUS_ITEM_ID, CUSTOMER_ID, Val(txtWeightPerPack.Text), PartType, PRICE_THINK_TYPE, m_ExDeliveryCostItem, m_Customers, RateDeliveryCost, RateDeliveryCostFlag, EX_DELIVERY_COST_ITEM_ID)
         Call calExPromotionDls(DELIVERY_CUS_ITEM_ID, CUSTOMER_ID, Val(txtWeightPerPack.Text), PartType, PRICE_THINK_TYPE, m_ExPromotionDlcItem, m_Customers, RatePromotionDlc, RatePromotionDlcFlag, EX_PROMOTION_DLC_ITEM_ID)
          
         txtPricePerPack.Text = RateDeliveryCost - RatePromotionDlc
'         txtDiscountPerPack.Text = RatePromotionDlc
         tempRate.EX_DELIVERY_COST_ITEM_ID = EX_DELIVERY_COST_ITEM_ID
         tempRate.RATE_CUSTOMER = RateDeliveryCost
         tempRate.EX_PROMOTION_DLC_ITEM_ID = EX_PROMOTION_DLC_ITEM_ID
         tempRate.DISCOUNT_AMOUNT_DLC = RatePromotionDlc
         If EX_DELIVERY_COST_ITEM_ID > 0 Then
         Call TempData.add(tempRate, Trim(str(CUSTOMER_ID)))
         End If
    End If
End Sub

Private Sub txtWeightPerPack_LostFocus()
Dim TempID As Long
  If TypeSale = 1 Then
'    If TypeSale = 2 Then
   TempID = uctlFeatureLookup.MyCombo.ItemData(Minus2Zero(uctlFeatureLookup.MyCombo.ListIndex))
    Call verifyDelivery(TempID)
 End If
End Sub

Private Sub uctlFeatureLookup_Change()
Dim TempID As Long
Dim Ft As CFeature
  
   m_HasModify = True
   
   txtPackAmount.Text = ""
 TempID = uctlFeatureLookup.MyCombo.ItemData(Minus2Zero(uctlFeatureLookup.MyCombo.ListIndex))
   If TempID > 0 Then
   Call DisableForReturn
   If TypeSale = 1 Then 'ขายอาหาร
         If TempID = 44 Then  'ถ้าเป็นค่าขนส่ง และขายสินค้า
            If CAL_RATE_DELIVERY_TYPE = 2 Then 'คิดตามปริมาณ
               txtWeightPerPack.Text = "1"
               txtPackAmount.Text = "1"
            End If
            uctlFeatureLookup.Enabled = True
            txtWeightPerPack.Enabled = True
            txtPackAmount.Enabled = True
            txtPricePerPack.Enabled = True
            txtDiscountPerPack.Enabled = True
         
         ElseIf TempID <> 44 Then
            uctlFeatureLookup.Enabled = True
         End If
   
   ElseIf TypeSale = 2 Then  'ขายอาหาร 'ขายอื่นๆ
      If (TempID = 44) Then
         uctlFeatureLookup.Enabled = True
      ElseIf TempID <> 44 Then
         txtPricePerPack.Text = ""
         uctlFeatureLookup.Enabled = True
         txtWeightPerPack.Enabled = True
         txtPackAmount.Enabled = True
         txtPricePerPack.Enabled = True
         txtDiscountPerPack.Enabled = True
      End If
      
   End If
   
'      If (TempID = 44) And TypeSale = 1 Then 'ถ้าเป็นค่าขนส่ง และขายสินค้า
'         txtPricePerPack.Enabled = True
'         If CAL_RATE_DELIVERY_TYPE = 2 Then 'คิดตามปริมาณ
'            txtWeightPerPack.Text = "1"
'            txtPackAmount.Text = "1"
'         End If
'      ElseIf (TempID <> 44) And TypeSale = 1 Then 'ถ้าไม่ใช่ค่าขนส่ง และขายสินค้า
'         uctlFeatureLookup.Enabled = True
'      ElseIf (TempID = 44) And TypeSale = 2 Then 'ถ้าเป็นค่าขนส่ง และขายสินค้าอื่นๆ
'         uctlFeatureLookup.Enabled = True
'      ElseIf (TempID <> 44) And TypeSale = 2 Then 'ถ้าไม่ใช่ค่าขนส่ง และขายสินค้าอื่นๆ
'         txtPricePerPack.Text = ""
'         uctlFeatureLookup.Enabled = True
'         txtWeightPerPack.Enabled = True
'         txtPackAmount.Enabled = True
'         txtPricePerPack.Enabled = True
'         txtDiscountPerPack.Enabled = True
''         Exit Sub
'      ElseIf (TempID = 44) And TypeSale = 3 Then 'กรณีการเข้าแก้ไข
'         uctlFeatureLookup.Enabled = True
'         txtWeightPerPack.Enabled = True
'         txtPackAmount.Enabled = True
'         txtPricePerPack.Enabled = True
'      ElseIf (TempID <> 44) And TypeSale = 3 Then  'กรณีการเข้าแก้ไขที่ไม่ใช่ค่าขนส่ง
'         uctlFeatureLookup.Enabled = True
'         txtWeightPerPack.Enabled = True
'         txtPackAmount.Enabled = True
'         txtPricePerPack.Enabled = True
'         txtDiscountPerPack.Enabled = True
''      Else 'ขายอื่นๆ
''         txtPricePerPack.Text = ""
''         uctlFeatureLookup.Enabled = True
''         txtWeightPerPack.Enabled = True
''         txtPackAmount.Enabled = True
''         txtPricePerPack.Enabled = True
''         txtDiscountPerPack.Enabled = True
'      End If
      
''''       Call verifyDelivery(TempID)
      Set Ft = GetFeature(m_Features, Trim(str(TempID)))
      txtTransferWage.Enabled = (Ft.LOGISTIC_FLAG = "Y")
      txtStdTrfCharge.Enabled = (Ft.LOGISTIC_FLAG = "Y")
   End If
End Sub

Private Sub uctlFeatureLookup_LostFocus()
Dim TempID As Long
If TypeSale = 1 Then
 TempID = uctlFeatureLookup.MyCombo.ItemData(Minus2Zero(uctlFeatureLookup.MyCombo.ListIndex))
 Call verifyDelivery(TempID)
End If
End Sub

Private Sub uctlPartLookup_LostFocus()
 If uctlPartLookup.MyCombo.ListIndex < 0 Then
      Exit Sub
   End If
   
   If TypeSale = 1 Then
      PartItemID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
      If PartItemID > 0 Then
         Set Pi = GetPartItem(m_Parts, Trim(str(PartItemID)))
'         Call InitNormalLabel(lblUnit, Pi.UNIT_NAME)
'         txtManualCode.Text = Pi.BARCODE_NO
'         txtManualName.Text = Pi.BILL_DESC
'         txtWeightPerPack.Text = Pi.WEIGHT_PER_PACK
'         PartType = Pi.PART_TYPE
'         WeigthPerPack = Pi.WEIGHT_PER_PACK
   
         If Len(TempUserName2) = 0 Or (PartItemID <> OldPartItemId And OldPartItemId <> 0) Then
               'ส่วนของคิดราคาแบบใหม่
               Dim ExPromotionPart As Double
               Dim ExPromotionDlc As Double
              txtPricePerPack.Text = calExWorksPrice(Pi, DELIVERY_CUS_ITEM_ID, CUSTOMER_ID, PRICE_THINK_TYPE, m_ExWorkPricesItem, m_ExDeliveryCostItem, m_Customers, TempData, m_ExPromotionPartItem, m_ExPromotionDlcItem, ExPromotionPart, ExPromotionDlc)
              CurrentPricePerPack = Val(txtPricePerPack.Text)
              NewPricePerPack = Val(txtPricePerPack.Text)
              If radFeature.Value Then
                txtDiscountPerPack.Text = ExPromotionDlc
              ElseIf radStock.Value Then
                txtDiscountPerPack.Text = ExPromotionPart
              End If
              Call showDetail(TempData)
              If Len(TempUserName2) > 0 Then
                 ' เครียร์สถานะการอนุมัติออกด้วยหากมีการคิดราคาอาหารจากราคาประกาศ
                  TempUserName = ""
                   EditPriceFlag = True
              End If
          End If
      Else
         txtPricePerPack.Text = "0"
      End If
   End If
   
   m_HasModify = True
End Sub

Private Sub uctlPigTypeLookup_Change()
Dim PartTypeID As Long
Dim Pt As CPartType
   PartTypeID = uctlPigTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPigTypeLookup.MyCombo.ListIndex))
   If PartTypeID > 0 Then
      Call DisableForReturn
      uctlPigTypeLookup.Enabled = True
      uctlPartLookup.Enabled = True
      uctlToLocationLookup.Enabled = True
      If TypeSale = 1 Then
         If (PartTypeID = 10 Or PartTypeID = 21) Then
            txtWeightPerPack.Enabled = True
            txtPackAmount.Enabled = True
            txtPricePerPack.Enabled = True
         Else
             txtPricePerPack.Enabled = True
             txtDiscountPerPack.Enabled = True
         End If
      ElseIf TypeSale = 2 Then
          If (PartTypeID = 10 Or PartTypeID = 21) Then
              uctlPartLookup.Enabled = False
              uctlToLocationLookup.Enabled = False
         Else
               txtWeightPerPack.Enabled = True
               txtPackAmount.Enabled = True
               txtPricePerPack.Enabled = True
               txtDiscountPerPack.Enabled = True
          End If
      End If
            
      Set Pt = GetPartType(m_PartTypes, Trim(str(PartTypeID)))
      Call LoadPartItem(uctlPartLookup.MyCombo, m_Parts, PartTypeID, "N")
      Set uctlPartLookup.MyCollection = m_Parts

      Call LoadLocation(uctlToLocationLookup.MyCombo, m_Locations, 2, , , Pt.PART_GROUP_ID)
      Set uctlToLocationLookup.MyCollection = m_Locations
      
   End If
   
   m_HasModify = True
'Dim PartTypeID As Long
'Dim Pt As CPartType
'   PartTypeID = uctlPigTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPigTypeLookup.MyCombo.ListIndex))
'   If PartTypeID > 0 Then
'      Call DisableForReturn
'      uctlPigTypeLookup.Enabled = True
'      uctlPartLookup.Enabled = True
'      uctlToLocationLookup.Enabled = True
'      If (PartTypeID = 10 Or PartTypeID = 21) And TypeSale = 1 Then
'         txtWeightPerPack.Enabled = True
'         txtPackAmount.Enabled = True
'         txtPricePerPack.Enabled = True
'      ElseIf (PartTypeID <> 10 Or PartTypeID <> 21) And TypeSale = 1 Then
'         txtPricePerPack.Enabled = True
'         txtDiscountPerPack.Enabled = True
'      ElseIf (PartTypeID = 10 Or PartTypeID = 21) And TypeSale = 2 Then
'         uctlPartLookup.Enabled = False
'         uctlToLocationLookup.Enabled = False
''      ElseIf (PartTypeID = 10 Or PartTypeID = 21) And TypeSale = 3 Then 'กรณีการเข้าแก้ไข
''         txtWeightPerPack.Enabled = True
''         txtPackAmount.Enabled = True
''         txtPricePerPack.Enabled = True
'      Else
'         txtWeightPerPack.Enabled = True
'         txtPackAmount.Enabled = True
'         txtPricePerPack.Enabled = True
'         txtDiscountPerPack.Enabled = True
'      End If
'
'      Set Pt = GetPartType(m_PartTypes, Trim(str(PartTypeID)))
'      Call LoadPartItem(uctlPartLookup.MyCombo, m_Parts, PartTypeID, "N")
'      Set uctlPartLookup.MyCollection = m_Parts
'
'      Call LoadLocation(uctlToLocationLookup.MyCombo, m_Locations, 2, , , Pt.PART_GROUP_ID)
'      Set uctlToLocationLookup.MyCollection = m_Locations
'
'   End If
'
'   m_HasModify = True
End Sub

Private Sub uctlToLocationLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlLocationLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlPartLookup_Change()
   If uctlPartLookup.MyCombo.ListIndex < 0 Then
      Exit Sub
   End If
   
'   PartItemID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
'   If PartItemID > 0 Then
'      Set Pi = GetPartItem(m_Parts, Trim(str(PartItemID)))
'      PartType = Pi.PART_TYPE
'   End If
      PartItemID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
      If PartItemID > 0 Then
         Set Pi = GetPartItem(m_Parts, Trim(str(PartItemID)))
         Call InitNormalLabel(lblUnit, Pi.UNIT_NAME)
         txtManualCode.Text = Pi.BARCODE_NO
         txtManualName.Text = Pi.BILL_DESC
         txtWeightPerPack.Text = Pi.WEIGHT_PER_PACK
         PartType = Pi.PART_TYPE
         WeigthPerPack = Pi.WEIGHT_PER_PACK
      End If
      
   m_HasModify = True
End Sub
Function showDetail(cTempData As Collection)
 Dim SumOther As Double
      Set TempD3 = GetObject("CDoItem", cTempData, Trim(str(CUSTOMER_ID)), False)
      If Not TempD3 Is Nothing Then
      If radStock.Value Then
         If PartType = 10 Then
            SumOther = TempD3.PRO_COMMISSION_BAG + TempD3.PRO_CHEER_BAG + TempD3.PRO_DST_BAG + TempD3.PRO_OTHER1_BAG + TempD3.PRO_OTHER2_BAG + TempD3.PRO_OTHER3_BAG
         ElseIf PartType = 21 Then
             SumOther = TempD3.PRO_COMMISSION_KG + TempD3.PRO_CHEER_KG + TempD3.PRO_DST_KG + TempD3.PRO_OTHER1_KG + TempD3.PRO_OTHER2_KG + TempD3.PRO_OTHER3_KG
         End If
      End If
         
         If PRICE_THINK_TYPE = 1 Then
            lblShowDetail.Caption = "ราคาสินค้า=" & TempD3.PACKAGE_RATE & ",ส่วนลดสินค้า=" & TempD3.DISCOUNT_AMOUNT_PART & ",อื่นๆ=" & SumOther
         ElseIf PRICE_THINK_TYPE = 2 Then
            lblShowDetail.Caption = "ราคาสินค้า=" & TempD3.PACKAGE_RATE & ",ส่วนลดสินค้า=" & TempD3.DISCOUNT_AMOUNT_PART & ",ค่าขนส่ง=" & TempD3.RATE_CUSTOMER & ",ส่วนลดขนส่ง=" & TempD3.DISCOUNT_AMOUNT_DLC & ",อื่นๆ=" & SumOther
         ElseIf PRICE_THINK_TYPE = 3 Then
            If radFeature.Value Then
                 lblShowDetail.Caption = "ค่าขนส่ง=" & TempD3.RATE_CUSTOMER & ",ส่วนลดขนส่ง=" & TempD3.DISCOUNT_AMOUNT_DLC & ""
            ElseIf radStock.Value Then
                lblShowDetail.Caption = "ราคาสินค้า=" & TempD3.PACKAGE_RATE & ",ส่วนลดสินค้า=" & TempD3.DISCOUNT_AMOUNT_PART & ",อื่นๆ=" & SumOther
            End If
        End If
      End If
End Function
Private Sub uctlPartTypeLookup_Change()
Dim PartTypeID As Long
   
   Call LoadPartItem(uctlPartLookup.MyCombo, m_Parts, PartTypeID, "")
   Set uctlPartLookup.MyCollection = m_Parts
   
   m_HasModify = True
End Sub

Private Sub uctlPigWeekLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlToPeriod_HasChange()
   m_HasModify = True
End Sub

Public Sub CalculateCurrentBath()
End Sub
Private Sub ClearData()
      txtReturnAvg.Text = ""
      txtManual.Text = ""
      txtQuantity.Text = ""
      txtPackAmount.Text = ""
      txtPricePerPack.Text = ""
      txtWeightPerPack.Text = ""
      
      txtAvgPrice.Text = ""
      txtDiscountPerPack.Text = ""
      txtDiscount.Text = ""
      txtExcludeDiscount.Text = ""
      
      txtTotalPrice.Text = ""
      
      txtManualCode.Text = ""
      txtTransferWage.Text = ""
      txtStdTrfCharge.Text = ""
End Sub
Private Sub DisableForReturn()
    Call ClearData

      txtReturnAvg.Enabled = False
      txtManual.Enabled = False
   
      SSOption1.Enabled = False
      SSOption2.Enabled = False
      SSOption3.Enabled = False
      
      txtQuantity.Enabled = False
      txtPackAmount.Enabled = False
      txtPricePerPack.Enabled = False
      txtWeightPerPack.Enabled = False
      
      txtAvgPrice.Enabled = False
      txtDiscountPerPack.Enabled = False
      txtDiscount.Enabled = False
      txtExcludeDiscount.Enabled = False
      
      txtTotalPrice.Enabled = False
      chkManualName.Enabled = False
      
      txtManualName.Enabled = False
      txtManualCode.Enabled = False
      txtTransferWage.Enabled = False
      txtStdTrfCharge.Enabled = False
End Sub
