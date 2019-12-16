VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditDoItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9450
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditDOItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   9450
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   8835
      Left            =   0
      TabIndex        =   31
      Top             =   600
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   15584
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboSumWithDoItemId 
         Height          =   510
         Left            =   1770
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   6480
         Width           =   7005
      End
      Begin VB.ComboBox cboRateType 
         Height          =   510
         Left            =   1770
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   3180
         Width           =   2445
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   555
         Left            =   1800
         TabIndex        =   45
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
         Left            =   1770
         TabIndex        =   15
         Top             =   4050
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPartLookup 
         Height          =   435
         Left            =   1785
         TabIndex        =   9
         Top             =   2280
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlToLocationLookup 
         Height          =   435
         Left            =   1770
         TabIndex        =   10
         Top             =   2730
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPigTypeLookup 
         Height          =   435
         Left            =   1785
         TabIndex        =   7
         Top             =   1830
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTotalPrice 
         Height          =   435
         Left            =   6780
         TabIndex        =   21
         Top             =   4950
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtAvgPrice 
         Height          =   435
         Left            =   7305
         TabIndex        =   16
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
         Left            =   7305
         TabIndex        =   19
         Top             =   4500
         Width           =   1485
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtWeightPerPack 
         Height          =   435
         Left            =   1770
         TabIndex        =   12
         Top             =   3600
         Width           =   1455
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPackAmount 
         Height          =   435
         Left            =   4620
         TabIndex        =   13
         Top             =   3600
         Width           =   1455
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPricePerPack 
         Height          =   435
         Left            =   7305
         TabIndex        =   14
         Top             =   3600
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtDiscountPerPack 
         Height          =   435
         Left            =   1770
         TabIndex        =   17
         Top             =   4500
         Width           =   1455
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtExcludeDiscount 
         Height          =   435
         Left            =   1770
         TabIndex        =   20
         Top             =   4950
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtManualCode 
         Height          =   465
         Left            =   1770
         TabIndex        =   23
         Top             =   6000
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtManualName 
         Height          =   465
         Left            =   3840
         TabIndex        =   29
         Top             =   6000
         Width           =   4995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTransferWage 
         Height          =   435
         Left            =   1770
         TabIndex        =   25
         Top             =   7170
         Width           =   1695
         _ExtentX        =   2514
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtStdTrfCharge 
         Height          =   435
         Left            =   7080
         TabIndex        =   26
         Top             =   7170
         Width           =   1695
         _ExtentX        =   2514
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtReturnAvg 
         Height          =   435
         Left            =   4620
         TabIndex        =   18
         Top             =   4500
         Width           =   1455
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin VB.Label lblSumWithDoItemId 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   59
         Top             =   6480
         Width           =   1485
      End
      Begin VB.Label lblReturnAvg 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   3240
         TabIndex        =   58
         Top             =   4560
         Width           =   1275
      End
      Begin VB.Label lblStdTrfCharge 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   5190
         TabIndex        =   57
         Top             =   7230
         Width           =   1785
      End
      Begin VB.Label Label6 
         Height          =   345
         Left            =   8865
         TabIndex        =   56
         Top             =   7200
         Width           =   405
      End
      Begin VB.Label lblTransferWage 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   180
         TabIndex        =   55
         Top             =   7230
         Width           =   1485
      End
      Begin VB.Label Label3 
         Height          =   345
         Left            =   3480
         TabIndex        =   54
         Top             =   7200
         Width           =   1215
      End
      Begin VB.Label lblRateType 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   270
         TabIndex        =   53
         Top             =   3270
         Width           =   1395
      End
      Begin Threed.SSCheck chkManualName 
         Height          =   435
         Left            =   1770
         TabIndex        =   22
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
         TabIndex        =   52
         Top             =   6000
         Width           =   1575
      End
      Begin VB.Label lblExcludeDiscount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   90
         TabIndex        =   51
         Top             =   5070
         Width           =   1575
      End
      Begin VB.Label Label1 
         Height          =   345
         Left            =   3900
         TabIndex        =   50
         Top             =   5010
         Width           =   495
      End
      Begin VB.Label lblDiscountPerPack 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   49
         Top             =   4560
         Width           =   1485
      End
      Begin VB.Label lblPricePerPack 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   6150
         TabIndex        =   48
         Top             =   3660
         Width           =   1095
      End
      Begin VB.Label lblPackAmount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   3240
         TabIndex        =   47
         Top             =   3660
         Width           =   1275
      End
      Begin VB.Label lblWeightPerPack 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   46
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
         TabIndex        =   44
         Top             =   5010
         Width           =   495
      End
      Begin VB.Label lblUnit 
         Height          =   345
         Left            =   3300
         TabIndex        =   43
         Top             =   4110
         Width           =   1215
      End
      Begin VB.Label Label4 
         Height          =   345
         Left            =   8850
         TabIndex        =   42
         Top             =   4620
         Width           =   495
      End
      Begin VB.Label lblDiscount 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   6030
         TabIndex        =   41
         Top             =   4590
         Width           =   1185
      End
      Begin VB.Label lblManual 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   40
         Top             =   960
         Width           =   1485
      End
      Begin VB.Label lblFeatureCode 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   39
         Top             =   1440
         Width           =   1485
      End
      Begin VB.Label lblTotalPrice 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   5370
         TabIndex        =   38
         Top             =   5070
         Width           =   1305
      End
      Begin VB.Label lblAvgPrice 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   5985
         TabIndex        =   37
         Top             =   4110
         Width           =   1215
      End
      Begin VB.Label Label2 
         Height          =   345
         Left            =   8865
         TabIndex        =   36
         Top             =   4110
         Width           =   435
      End
      Begin VB.Label lblPigType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   180
         TabIndex        =   35
         Top             =   1860
         Width           =   1485
      End
      Begin VB.Label lblToLocation 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   180
         TabIndex        =   34
         Top             =   2790
         Width           =   1485
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3120
         TabIndex        =   27
         Top             =   8040
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
         Left            =   4800
         TabIndex        =   28
         Top             =   8040
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
         TabIndex        =   33
         Top             =   2340
         Width           =   1485
      End
      Begin VB.Label lblQuantity 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   32
         Top             =   4110
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmAddEditDoItem"
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
Public id As Long
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
Private m_Suppliers As Collection
Private m_SuppliersTransport As Collection

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
Private m_WagePrice As Double
Public DeliveryCostFlag As Boolean

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

Private Sub cboRateType_Click()
   m_HasModify = True
End Sub

Private Sub cboRateType_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cboSumWithDoItemId_Change()
   m_HasModify = True
End Sub

Private Sub cboSumWithDoItemId_Click()
   m_HasModify = True
End Sub

Private Sub cboSumWithDoItemId_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
'      If Not DeliveryCostFlag Then
'         'SendKeys ("{TAB}")
'        cmdOK.SetFocus
'      Else
'         SendKeys ("{TAB}")
'      End If
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
'      cmdLotSelect.Visible = False
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
'      cmdLotSelect.Visible = True
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
    Call InitNormalLabel(lblRateType, MapText("คิดราคาแบบ"))
    Call InitNormalLabel(lblTransferWage, MapText("ค่าจ้างขนส่ง"))
    Call InitNormalLabel(lblStdTrfCharge, MapText("ค่าขนส่งมาตรฐาน"))

    Call InitNormalLabel(Label3, MapText("บาท"))
    Call InitNormalLabel(lblReturnAvg, MapText("ทุนคืน"))
    Call InitNormalLabel(lblSumWithDoItemId, MapText("รวมยอดกับ"))
    
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
   Call txtTransferWage.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   Call txtStdTrfCharge.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)

   txtStdTrfCharge.Enabled = False
   
   Call InitCheckBox(chkManualName, "กำหนด รหัส/ชื่อ เอง")
   Call InitCombo(cboRateType)
   Call InitCombo(cboSumWithDoItemId)
   
   txtManualCode.Enabled = False
   txtManualName.Enabled = False
   chkManualName.Enabled = False
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   
   If Area = 1 Then
'      cmdLotSelect.Visible = True
      radFeature.Enabled = True
   Else
'      cmdLotSelect.Visible = False
      radFeature.Enabled = False
   End If

   'Barcode flag
   If m_Sp.PARAM_VALUE = "Y" Then
      uctlPigTypeLookup.Enabled = False
   Else
      uctlPigTypeLookup.Enabled = True
   End If
   
   Call ShowGui
End Sub

Private Sub CalculatePrice()
'   txtLeft.Text = Format(Val(txtNetTotal.Text) - Val(txtDeposit.Text), "0.00")
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
           
            Set Ri = TempCollection.Item(id)
                    
            Call ShowConfigFlag(Ri.CONFIG_CODE)
            ' radFeature.Value
            
            
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
            cboRateType.ListIndex = IDToListIndex(cboRateType, Ri.RATE_TYPE)
            txtTransferWage.Text = Ri.TRANSFER_WAGE
            txtStdTrfCharge.Text = Ri.STD_TRANSFER_CHARGE
            
            cmdOK.Enabled = (COMMIT_FLAG <> "Y")
            Call CalculateCurrentBath
         Else
         
            Dim Di As CDoItem
           
            Set Di = TempCollection.Item(id)
                    
            Call ShowConfigFlag(Di.CONFIG_CODE)
            
'            Call ForDeliveryCost(radFeature.Value)
            
            uctlPigTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPigTypeLookup.MyCombo, Di.PART_TYPE)
            uctlPartLookup.MyCombo.ListIndex = IDToListIndex(uctlPartLookup.MyCombo, Di.PART_ITEM_ID)
            uctlFeatureLookup.MyCombo.ListIndex = IDToListIndex(uctlFeatureLookup.MyCombo, Di.FEATURE_ID)
            txtWeightPerPack.Text = Di.WEIGHT_PER_PACK
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
            cboRateType.ListIndex = IDToListIndex(cboRateType, Di.RATE_TYPE)
            txtTransferWage.Text = Di.TRANSFER_WAGE
            txtStdTrfCharge.Text = Di.STD_TRANSFER_CHARGE
            cboSumWithDoItemId.ListIndex = IDToListIndex(cboSumWithDoItemId, Di.SUM_WITH_DO_ITEM_ID)
            
            cmdOK.Enabled = (COMMIT_FLAG <> "Y")
            Call CalculateCurrentBath
         End If
      End If
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cmdLotSelect_Click()
Dim OKClick As Boolean
Dim LotAmount As Long

   If Not VerifyCombo(lblPart, uctlPartLookup.MyCombo, False) Then
      Exit Sub
   End If
   If Not VerifyCombo(lblToLocation, uctlToLocationLookup.MyCombo, False) Then
      Exit Sub
   End If
   
   frmAddSubLotItem.HeaderText = "เลือกหมายเลขล็อต"
   frmAddSubLotItem.ShowMode = ShowMode
'   If ShowMode = SHOW_ADD Then
      Set frmAddSubLotItem.TempCollection = m_SubLotItems
'   Else
'      Call CopySubLotItem(TempCollection(ID).SubLotItems, m_SubLotItems)
'      Set frmAddSubLotItem.TempCollection = m_SubLotItems
'   End If
   frmAddSubLotItem.PartItemID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
   frmAddSubLotItem.LocationID = uctlToLocationLookup.MyCombo.ItemData(Minus2Zero(uctlToLocationLookup.MyCombo.ListIndex))
   Load frmAddSubLotItem
   frmAddSubLotItem.Show 1
   
   OKClick = frmAddSubLotItem.OKClick
   LotAmount = frmAddSubLotItem.SumLotAmount
   
   Unload frmAddSubLotItem
   Set frmAddSubLotItem = Nothing
   
   If OKClick Then
      m_ManualFlag = True
      txtQuantity.Text = LotAmount
      m_HasModify = True
   End If
End Sub

Private Sub cmdOK_Click()
   If Not cmdOK.Enabled Then
      Exit Sub
   End If
   
   If Not SaveData Then
      Exit Sub
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
'   If Not VerifyTextControl(lblTransferWage, txtTransferWage, Not txtTransferWage.Enabled) Then
'      Exit Function
'   End If
'   If Not VerifyTextControl(lblStdTrfCharge, txtStdTrfCharge, Not txtStdTrfCharge.Enabled) Then
'      Exit Function
'   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   If DocumentType = 18 Then
      Dim Ri As CReceiptItem
      If ShowMode = SHOW_ADD Then
         Set Ri = New CReceiptItem
         
         Ri.Flag = "A"
         Call TempCollection.add(Ri)
      Else
         Set Ri = TempCollection.Item(id)
         If Ri.Flag <> "A" Then
            Ri.Flag = "E"
         End If
      End If
   
      If radStock.Value Then
         Ri.PART_ITEM_ID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
      Else
         Ri.PART_ITEM_ID = -1
      End If
      Ri.PART_NO = uctlPartLookup.MyTextBox.Text
      Ri.PART_DESC = uctlPartLookup.MyCombo.Text
      Ri.RETURN_AMOUNT = txtQuantity.Text
      Ri.LOCATION_ID = uctlToLocationLookup.MyCombo.ItemData(Minus2Zero(uctlToLocationLookup.MyCombo.ListIndex))
      Ri.LOCATION_NAME = uctlToLocationLookup.MyCombo.Text
      If m_Sp.PARAM_VALUE = "Y" Then
         Ri.PART_TYPE = -1
      Else
         Ri.PART_TYPE = uctlPigTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPigTypeLookup.MyCombo.ListIndex))
      End If
      Ri.RETURN_TOTAL_PRICE = Val(Format(Val(txtTotalPrice.Text), "0.00")) 'แปลงให้เป็น 2 ตำแหน่งด้วย
      Ri.AVG_PRICE = Val(txtAvgPrice.Text)
      Ri.RETURN_AVG_PRICE = Val(txtReturnAvg.Text)
      Ri.AVG_WEIGHT = 0
      Ri.FEATURE_ID = uctlFeatureLookup.MyCombo.ItemData(Minus2Zero(uctlFeatureLookup.MyCombo.ListIndex))
      Ri.FEATURE_CODE = uctlFeatureLookup.MyTextBox.Text
      Ri.FEATURE_DESC = uctlFeatureLookup.MyCombo.Text
      Ri.RETURN_DISCOUNT_AMOUNT = Val(Format(Val(txtDiscount.Text), "0.00"))  'แปลงให้เป็น 2 ตำแหน่งด้วย
      Ri.CONFIG_CODE = CreateConfigFlag()
      Ri.ITEM_DESC = txtManual.Text
      Ri.DISPLAY_ID = GetDisplayID
      Ri.COUNTRY_CURRENCY1 = COUNTRY_CURRENCY1
      Ri.COUNTRY_CURRENCY2 = COUNTRY_CURRENCY2
      Ri.WEIGHT_PER_PACK = Val(txtWeightPerPack.Text)
      Ri.PACK_AMOUNT = Val(txtPackAmount.Text)
      Ri.PRICE_PER_PACK = Val(txtPricePerPack.Text)
      Ri.DISCOUNT_PER_PACK = Val(txtDiscountPerPack.Text)
      Ri.MANUAL_FLAG = Check2Flag(chkManualName.Value)
      Ri.MANUAL_CODE = txtManualCode.Text
      Ri.MANUAL_NAME = txtManualName.Text
      Ri.RATE_TYPE = cboRateType.ItemData(Minus2Zero(cboRateType.ListIndex))
      Ri.TRANSFER_WAGE = Val(txtTransferWage.Text)
      Ri.STD_TRANSFER_CHARGE = Val(Format(Val(txtStdTrfCharge.Text), "0.00")) 'แปลงให้เป็น 2 ตำแหน่งด้วย
      
      Ri.DEBIT_CREDIT_AMOUNT = Ri.RETURN_TOTAL_PRICE - Ri.RETURN_DISCOUNT_AMOUNT
   
   Else
      Dim Di As CDoItem
      If ShowMode = SHOW_ADD Then
         Set Di = New CDoItem
   
         Di.Flag = "A"
         Call TempCollection.add(Di)
      Else
         Set Di = TempCollection.Item(id)
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
      Di.PACK_AMOUNT = Val(txtPackAmount.Text)
      Di.PRICE_PER_PACK = Val(txtPricePerPack.Text)
      Di.DISCOUNT_PER_PACK = Val(txtDiscountPerPack.Text)
      Di.MANUAL_FLAG = Check2Flag(chkManualName.Value)
      Di.MANUAL_CODE = txtManualCode.Text
      Di.MANUAL_NAME = txtManualName.Text
      Di.RATE_TYPE = cboRateType.ItemData(Minus2Zero(cboRateType.ListIndex))
      Di.TRANSFER_WAGE = Val(txtTransferWage.Text)
      Di.STD_TRANSFER_CHARGE = Val(Format(Val(txtStdTrfCharge.Text), "0.00")) 'แปลงให้เป็น 2 ตำแหน่งด้วย
   
      If cboSumWithDoItemId.ListIndex > -1 Then
         Di.SUM_WITH_DO_ITEM_ID = cboSumWithDoItemId.ItemData(cboSumWithDoItemId.ListIndex)
      End If
   End If
   
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call InitDoRateType(cboRateType)
      
      Call LoadFeature(uctlFeatureLookup.MyCombo, m_Features)
      Set uctlFeatureLookup.MyCollection = m_Features
         
      If m_Sp.PARAM_VALUE = "Y" Then
         Call LoadPartItem(uctlPartLookup.MyCombo, m_Pigs, , , , , "N")
         Set uctlPartLookup.MyCollection = m_Pigs
      Else
         Call LoadPartType(uctlPigTypeLookup.MyCombo, m_PartTypes)
         Set uctlPigTypeLookup.MyCollection = m_PartTypes
      End If
      
      Call LoadLocation(uctlToLocationLookup.MyCombo, m_Houses, 2)
      Set uctlToLocationLookup.MyCollection = m_Houses
      
      Call LoadSumWithDoItemId(cboSumWithDoItemId, TempCollection)
'      Set uctlSumWithDoItemId.MyCollection = m_SumWith

      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         id = 0
         radFeature.Value = True
         Call QueryData(True)
      End If
      m_HasModify = False
   
   End If
   Call DisableForReturn
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
   Set m_Suppliers = New Collection
   Set m_SuppliersTransport = New Collection
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
   Set m_Suppliers = Nothing
   Set m_SuppliersTransport = Nothing
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
End Sub

Private Sub radFeature_Click(Value As Integer)
   m_HasModify = True
   Call ShowGui
      chkManualName.Value = FlagToCheck("N")
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
   Call CalculatePrice
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

   RateType = cboRateType.ItemData(Minus2Zero(cboRateType.ListIndex))
   If RateType = RATE_CUSTOM Then
      Exit Sub
   End If
   
'   If uctlFeatureLookup.MyCombo.ListIndex <= 0 Then
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
   Call CalculatePrice
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

Private Sub txtNoteTransport_Change()
   m_HasModify = True
End Sub

Private Sub txtPackAmount_Change()
   m_HasModify = True
   txtQuantity.Text = Val(txtWeightPerPack.Text) * Val(txtPackAmount.Text)
End Sub

Private Sub txtPricePerPack_Change()
   m_HasModify = True
   txtAvgPrice.Text = MyDiff(Val(txtPackAmount.Text) * Val(txtPricePerPack.Text), Val(txtQuantity.Text))
   txtTransferWage.Text = Val(txtPackAmount.Text) * m_WagePrice
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

   RateType = cboRateType.ItemData(Minus2Zero(cboRateType.ListIndex))
   If RateType = RATE_CUSTOM Then
      SocLevel = "N"
   ElseIf RateType = RATE_MASTER Then
      SocLevel = "Y"
   Else
      SocLevel = "Y"
   End If
   
   If radCustom.Value Then
      Exit Sub
   End If

'   If uctlFeatureLookup.MyCombo.ListIndex <= 0 Then
'      Exit Sub
'   End If

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
            m_WagePrice = Ug.LOGISTIC_PRICE
            m_SocID = Ug.SOC_ID
         End If
      End If
   End If
   
   TempID = uctlFeatureLookup.MyCombo.ItemData(Minus2Zero(uctlFeatureLookup.MyCombo.ListIndex))
   Set Ft = GetFeature(m_Features, Trim(str(TempID)))
   If Ft.LOGISTIC_FLAG = "Y" Then
      txtStdTrfCharge.Text = Val(txtExcludeDiscount.Text)
   End If
   txtTransferWage.Text = m_WagePrice * Val(txtPackAmount.Text)
   Set Ug = Nothing
   Call EnableForm(Me, True)
End Sub

Private Sub txtQuantity_Change()
   m_HasModify = True
   'txtAvgPrice.Text = Round(MyDiffEx(Val(txtPackAmount.Text) * Val(txtPricePerPack.Text), Val(txtQuantity.Text)), 6)
   txtAvgPrice.Text = MyDiff(Val(txtPackAmount.Text) * Val(txtPricePerPack.Text), Val(txtQuantity.Text))
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

Private Sub txtRateCustomerTransport_Change()
   m_HasModify = True
End Sub

Private Sub txtRateDriverTransport_Change()
   m_HasModify = True
End Sub

Private Sub txtRateFacTransport_Change()
   m_HasModify = True
End Sub

Private Sub txtRateSaleTransport_Change()
   m_HasModify = True
End Sub

Private Sub txtReturnAvg_Change()
   m_HasModify = True
End Sub

Private Sub txtStdTrfCharge_Change()
   m_HasModify = True
End Sub

Private Sub txtTotalPrice_Change()
'   m_HasModify = True
'   If Val(txtQuantity.Text) > 0 Then
'      txtAvgPrice.Text = Val(txtTotalPrice.Text) / Val(txtQuantity.Text)
'   Else
'      txtAvgPrice.Text = "0.00"
'   End If
   Call CalculatePrice
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

Private Sub uctlFeatureLookup_Change()
Dim TempID As Long
Dim Ft As CFeature

   TempID = uctlFeatureLookup.MyCombo.ItemData(Minus2Zero(uctlFeatureLookup.MyCombo.ListIndex))
   Set Ft = GetFeature(m_Features, Trim(str(TempID)))
'   txtTransferWage.Enabled = (Ft.LOGISTIC_FLAG = "Y")
If Not DeliveryCostFlag Then
   txtStdTrfCharge.Enabled = (Ft.LOGISTIC_FLAG = "Y")
End If
End Sub

Private Sub uctlPigTypeLookup_Change()
Dim PartTypeID As Long
Dim Pt As CPartType

PartTypeID = uctlPigTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPigTypeLookup.MyCombo.ListIndex))


'--------------------------------------------------------------เล็กเพิ่ม ถ้าเป็น Bag หรือ Bulk จะไม่ยอมให้แสดง ให้ไปเลือกอีกเมนู
If ShowMode = SHOW_ADD Then
   If PartTypeID = 10 Or PartTypeID = 21 Then
      uctlPartLookup.MyTextBox.Text = ""
      uctlPartLookup.MyCombo.Clear
      Exit Sub
   End If
End If

If ShowMode = SHOW_EDIT Then
    If PartTypeID = 10 Or PartTypeID = 21 Then
      txtWeightPerPack.Enabled = False
      txtPackAmount.Enabled = False
      txtQuantity.Enabled = False
   End If
End If
'--------------------------------------------------------------

   If PartTypeID > 0 Then
      Set Pt = GetPartType(m_PartTypes, Trim(str(PartTypeID)))
      Call LoadPartItem(uctlPartLookup.MyCombo, m_Parts, PartTypeID, "N")
      Set uctlPartLookup.MyCollection = m_Parts
   
         Call LoadLocation(uctlToLocationLookup.MyCombo, m_Locations, 2, , , Pt.PART_GROUP_ID)
         Set uctlToLocationLookup.MyCollection = m_Locations
   End If
   
   m_HasModify = True
End Sub

Private Sub uctlToLocationLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlLocationLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlPartLookup_Change()
Dim Pi As CPartItem
Dim PartItemID As Long

   If uctlPartLookup.MyCombo.ListIndex < 0 Then
      Exit Sub
   End If
   
   PartItemID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
   If PartItemID > 0 Then
      Set Pi = GetPartItem(m_Parts, Trim(str(PartItemID)))
      Call InitNormalLabel(lblUnit, Pi.UNIT_NAME)
      txtManualCode.Text = Pi.BARCODE_NO
      txtManualName.Text = Pi.BILL_DESC
      txtWeightPerPack.Text = Pi.WEIGHT_PER_PACK
   End If
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
Private Sub DisableForReturn()
   If DocumentType <> 18 Then
      txtReturnAvg.Enabled = False
   Else
      SSFrame2.Enabled = False
      
      txtManual.Enabled = False
      uctlFeatureLookup.Enabled = False
      uctlPigTypeLookup.Enabled = False
      uctlPartLookup.Enabled = False
      uctlToLocationLookup.Enabled = False
   
      SSOption1.Enabled = False
      SSOption2.Enabled = False
      SSOption3.Enabled = False
      
      cboRateType.Enabled = False
      txtQuantity.Enabled = False
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
   End If
End Sub
