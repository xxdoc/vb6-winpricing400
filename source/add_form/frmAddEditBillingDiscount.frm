VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditBillingDiscount 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5295
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
   Icon            =   "frmAddEditBillingDiscount.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   9435
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   4725
      Left            =   0
      TabIndex        =   12
      Top             =   600
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   8334
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboPartItem 
         Height          =   510
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1380
         Width           =   5385
      End
      Begin VB.ComboBox cboPartFeature 
         Height          =   510
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1830
         Width           =   5385
      End
      Begin VB.ComboBox cboDiscountType 
         Height          =   510
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2280
         Width           =   2985
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   555
         Left            =   1800
         TabIndex        =   16
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
      Begin prjFarmManagement.uctlTextBox txtManual 
         Height          =   435
         Left            =   1800
         TabIndex        =   3
         Top             =   930
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtWeightPerPack 
         Height          =   435
         Left            =   1800
         TabIndex        =   7
         Top             =   2700
         Width           =   1455
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtExcludeDiscount 
         Height          =   465
         Left            =   1800
         TabIndex        =   8
         Top             =   3150
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   820
      End
      Begin VB.Label lblPartItem 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   22
         Top             =   1440
         Width           =   1485
      End
      Begin VB.Label lblExcludeDiscount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   3270
         Width           =   1575
      End
      Begin VB.Label Label1 
         Height          =   345
         Left            =   3870
         TabIndex        =   20
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label lblPackAmount 
         Height          =   375
         Left            =   3300
         TabIndex        =   19
         Top             =   2760
         Width           =   2535
      End
      Begin VB.Label lblWeightPerPack 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   2760
         Width           =   1485
      End
      Begin VB.Label lblPercent 
         Height          =   345
         Left            =   5040
         TabIndex        =   17
         Top             =   5040
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblManual 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   15
         Top             =   990
         Width           =   1485
      End
      Begin VB.Label lblFeatureCode 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   14
         Top             =   1890
         Width           =   1485
      End
      Begin VB.Label lblToLocation 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   13
         Top             =   2340
         Width           =   1485
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3090
         TabIndex        =   9
         Top             =   3840
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
         TabIndex        =   10
         Top             =   3840
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditBillingDiscount"
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

Private Sub cboDiscountType_Click()
Dim TempID As Long

   TempID = cboDiscountType.ItemData(Minus2Zero(cboDiscountType.ListIndex))
   If TempID <= 0 Then
      Call InitNormalLabel(lblPackAmount, "")
   Else
      If TempID = 1 Then
         Call InitNormalLabel(lblPackAmount, "บาท/ถุง")
      ElseIf TempID = 2 Then
         Call InitNormalLabel(lblPackAmount, "บาท/ก.ก.")
      ElseIf TempID = 3 Then
         Call InitNormalLabel(lblPackAmount, "% ของจำนวนเงิน")
      End If
   End If
   
   m_HasModify = True
End Sub

Private Sub cboDiscountType_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cboPartFeature_Click()
   m_HasModify = True
End Sub

Private Sub cboPartFeature_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cboPartItem_Click()
   m_HasModify = True
End Sub

Private Sub LoadDoFeature()
Dim Di As CDoItem
Dim I As Long

   cboPartFeature.Clear
   I = 0
   cboPartFeature.AddItem ("")
   
   I = 0
   For Each Di In TempCollection2
      If (Di.Flag <> "D") And (Di.FEATURE_ID > 0) Then
         I = I + 1
         cboPartFeature.AddItem (Di.FEATURE_DESC & " (" & Di.FEATURE_CODE & ")")
         cboPartFeature.ItemData(I) = Di.FEATURE_ID
      End If
   Next Di
End Sub

Private Sub LoadDoPartItem()
Dim Di As CDoItem
Dim I As Long

   cboPartItem.Clear
   I = 0
   cboPartItem.AddItem ("")
   
   I = 0
   For Each Di In TempCollection2
      If (Di.Flag <> "D") And (Di.PART_ITEM_ID > 0) Then
         I = I + 1
         cboPartItem.AddItem (Di.PART_DESC & " (" & Di.PART_NO & ")")
         cboPartItem.ItemData(I) = Di.PART_ITEM_ID
      End If
   Next Di
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub ShowGui()
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
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText

   Call InitNormalLabel(lblToLocation, MapText("ประเภทส่วนลด"))
   Call InitNormalLabel(lblFeatureCode, MapText("บริการ"))
   Call InitNormalLabel(lblPartItem, MapText("สินค้า"))
   Call InitNormalLabel(lblManual, MapText("ชื่อส่วนลด"))
   Call InitNormalLabel(lblWeightPerPack, MapText("ส่วนลด"))
   Call InitNormalLabel(lblPackAmount, MapText(""))
   Call InitNormalLabel(lblExcludeDiscount, MapText("มูลค่าส่วนลด"))
   Call InitNormalLabel(Label1, MapText("บาท"))
    
   Call InitOptionEx(radFeature, "บริการ")
   Call InitOptionEx(radStock, "สินค้า")
   Call InitOptionEx(radCustom, "ทั้งหมด")

   Call txtManual.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtWeightPerPack.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   Call txtExcludeDiscount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
'   txtExcludeDiscount.Enabled = False
   
   Call InitCombo(cboDiscountType)
   Call InitCombo(cboPartFeature)
   Call InitCombo(cboPartItem)
   
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
   
   Call ShowGui
End Sub

Private Sub CalculatePrice()
'   txtLeft.Text = Format(Val(txtNetTotal.Text) - Val(txtDeposit.Text), "0.00")
End Sub

Private Sub ShowDisplayID(Did As Long)

End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long
Dim Ivd As CInventoryDoc
Dim iCount As Long
Dim Ei As CLotItem

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Dim Di As CBillingDiscount

         Set Di = TempCollection.Item(ID)

         If Di.DISCOUNT_TYPE = 1 Then
            radStock.Value = True
            radFeature.Value = False
            radCustom.Value = False
         ElseIf Di.DISCOUNT_TYPE = 2 Then
            radStock.Value = False
            radFeature.Value = True
            radCustom.Value = False
         Else
            radFeature.Value = False
            radStock.Value = False
            radCustom.Value = True
         End If
         
         cboPartItem.ListIndex = IDToListIndex(cboPartItem, Di.PART_ITEM_ID)
         cboPartFeature.ListIndex = IDToListIndex(cboPartFeature, Di.FEATURE_ID)
         
         txtManual.Text = Di.DISCOUNT_NAME
         txtExcludeDiscount.Text = Di.DISCOUNT_AMOUNT
         
         If Di.DSCN_PER_PACK > 0 Then
            cboDiscountType.ListIndex = IDToListIndex(cboDiscountType, 1)
            txtWeightPerPack.Text = Di.DSCN_PER_PACK
         ElseIf Di.DSCN_PER_WEIGHT > 0 Then
            cboDiscountType.ListIndex = IDToListIndex(cboDiscountType, 2)
            txtWeightPerPack.Text = Di.DSCN_PER_WEIGHT
         ElseIf Di.DSCN_PER_MONEY > 0 Then
            cboDiscountType.ListIndex = IDToListIndex(cboDiscountType, 3)
            txtWeightPerPack.Text = Di.DSCN_PER_MONEY
         End If
         
         cmdOK.Enabled = (COMMIT_FLAG <> "Y")
      End If
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
   
   OKClick = True
   Unload Me
End Sub

Private Function GetDisplayID() As Long

End Function

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long

   If (Not radFeature.Value) And (Not radStock.Value) And (Not radCustom.Value) Then
      glbErrorLog.LocalErrorMsg = "กรุณากำหนดตัวเลือกอย่างใดอย่างหนึ่ง"
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   If Not VerifyTextControl(lblManual, txtManual, False) Then
      Exit Function
   End If
   
   If Not VerifyCombo(lblToLocation, cboDiscountType, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If

   Dim Di As CBillingDiscount
   If ShowMode = SHOW_ADD Then
      Set Di = New CBillingDiscount

      Di.Flag = "A"
      Call TempCollection.add(Di)
   Else
      Set Di = TempCollection.Item(ID)
      If Di.Flag <> "A" Then
         Di.Flag = "E"
      End If
   End If

   If radFeature.Value Then
      Di.ITEM_DESC = cboPartFeature.Text
   ElseIf radStock.Value Then
      Di.ITEM_DESC = cboPartItem.Text
   Else
      Di.ITEM_DESC = ""
   End If
   Di.DISCOUNT_NAME = txtManual.Text
   Di.PART_ITEM_ID = cboPartItem.ItemData(Minus2Zero(cboPartItem.ListIndex))
   Di.FEATURE_ID = cboPartFeature.ItemData(Minus2Zero(cboPartFeature.ListIndex))
   Di.DISCOUNT_AMOUNT = Val(txtExcludeDiscount.Text)
   
   Dim TempID As Long
   TempID = cboDiscountType.ItemData(Minus2Zero(cboDiscountType.ListIndex))
   If TempID = 1 Then
      Di.DSCN_PER_PACK = Val(txtWeightPerPack.Text)
      Di.DSCN_PER_WEIGHT = 0
      Di.DSCN_PER_MONEY = 0
   ElseIf TempID = 2 Then
      Di.DSCN_PER_PACK = 0
      Di.DSCN_PER_WEIGHT = Val(txtWeightPerPack.Text)
      Di.DSCN_PER_MONEY = 0
   ElseIf TempID = 3 Then
      Di.DSCN_PER_PACK = 0
      Di.DSCN_PER_WEIGHT = 0
      Di.DSCN_PER_MONEY = Val(txtWeightPerPack.Text)
   End If
   
   If radStock.Value Then
      Di.DISCOUNT_TYPE = 1
   ElseIf radFeature.Value Then
      Di.DISCOUNT_TYPE = 2
   ElseIf radCustom.Value Then
      Di.DISCOUNT_TYPE = 3
   End If
   
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents

      Call InitDiscountType(cboDiscountType)
      Call LoadDoPartItem
      Call LoadDoFeature
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
         radFeature.Value = True
         Call QueryData(True)
      End If
      
      m_HasModify = False
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
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
   cboPartFeature.ListIndex = -1
   cboPartItem.ListIndex = -1
   Call SetEnableDisableComboBox(cboPartFeature, False)
   Call SetEnableDisableComboBox(cboPartItem, False)
End Sub

Private Sub radFeature_Click(Value As Integer)
   m_HasModify = True
   Call SetEnableDisableComboBox(cboPartFeature, True)
   cboPartItem.ListIndex = -1
   Call SetEnableDisableComboBox(cboPartItem, False)
End Sub

Private Sub radFeature_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub radStock_Click(Value As Integer)
   m_HasModify = True
   cboPartFeature.ListIndex = -1
   Call SetEnableDisableComboBox(cboPartFeature, False)
   Call SetEnableDisableComboBox(cboPartItem, True)
End Sub

Private Sub radStock_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub txtExcludeDiscount_Change()
   m_HasModify = True
End Sub

Private Sub CalculateDiscount()
Dim Di As CDoItem
Dim SumPack As Double
Dim SumWeight As Double
Dim SumPrice As Double
Dim FeatureID As Long
Dim PartItemID As Long
Dim DiscountAmount As Double
Dim TempID As Long

   FeatureID = cboPartFeature.ItemData(Minus2Zero(cboPartFeature.ListIndex))
   PartItemID = cboPartItem.ItemData(Minus2Zero(cboPartItem.ListIndex))
   TempID = cboDiscountType.ItemData(Minus2Zero(cboDiscountType.ListIndex))
   
   SumPack = 0
   SumWeight = 0
   SumPrice = 0
   
   For Each Di In TempCollection2
      If Di.Flag <> "D" Then
         If radFeature.Value And (FeatureID <= 0) And (Di.FEATURE_ID > 0) Then
            SumPack = SumPack + Di.PACK_AMOUNT
            SumWeight = SumWeight + Di.ITEM_AMOUNT
            SumPrice = SumPrice + Di.TOTAL_PRICE
         ElseIf radFeature.Value And (FeatureID > 0) And (Di.FEATURE_ID = FeatureID) Then
            SumPack = SumPack + Di.PACK_AMOUNT
            SumWeight = SumWeight + Di.ITEM_AMOUNT
            SumPrice = SumPrice + Di.TOTAL_PRICE
         End If
         
         If radStock.Value And (PartItemID <= 0) And (Di.PART_ITEM_ID > 0) Then
            SumPack = SumPack + Di.PACK_AMOUNT
            SumWeight = SumWeight + Di.ITEM_AMOUNT
            SumPrice = SumPrice + Di.TOTAL_PRICE
         ElseIf radStock.Value And (PartItemID > 0) And (Di.PART_ITEM_ID = PartItemID) Then
            SumPack = SumPack + Di.PACK_AMOUNT
            SumWeight = SumWeight + Di.ITEM_AMOUNT
            SumPrice = SumPrice + Di.TOTAL_PRICE
         End If
         
         If (radCustom.Value) Then
            SumPack = SumPack + Di.PACK_AMOUNT
            SumWeight = SumWeight + Di.ITEM_AMOUNT
            SumPrice = SumPrice + Di.TOTAL_PRICE
         End If
      End If
   Next Di
   
   If TempID = 1 Then
      DiscountAmount = SumPack * Val(txtWeightPerPack.Text)
   ElseIf TempID = 2 Then
      DiscountAmount = SumWeight * Val(txtWeightPerPack.Text)
   ElseIf TempID = 3 Then
      DiscountAmount = SumPrice * Val(txtWeightPerPack.Text) / 100
   End If
   
   txtExcludeDiscount.Text = Format(DiscountAmount, "0.00")
End Sub

Private Sub txtExcludeDiscount_GotFocus()
   If Len(Trim(txtExcludeDiscount.Text)) > 0 Then
      Exit Sub
   End If
   
   Call CalculateDiscount
End Sub

Private Sub txtManual_Change()
   m_HasModify = True
End Sub

Private Sub uctlFeatureLookup_Change()

End Sub

Private Sub txtWeightPerPack_Change()
   m_HasModify = True
End Sub
