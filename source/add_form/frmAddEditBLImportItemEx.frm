VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditBLImportItemEx 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8610
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditBLImportItemEx.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   8610
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   4785
      Left            =   0
      TabIndex        =   12
      Top             =   600
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   8440
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextLookup uctlPartTypeLookup 
         Height          =   435
         Left            =   2040
         TabIndex        =   0
         Top             =   240
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPrice 
         Height          =   435
         Left            =   2025
         TabIndex        =   3
         Top             =   1650
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtQuantity 
         Height          =   435
         Left            =   2040
         TabIndex        =   2
         Top             =   1200
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPartLookup 
         Height          =   435
         Left            =   2040
         TabIndex        =   1
         Top             =   720
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlLocationLookup 
         Height          =   435
         Left            =   2055
         TabIndex        =   7
         Top             =   3030
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTotalPrice 
         Height          =   435
         Left            =   2040
         TabIndex        =   4
         Top             =   2100
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtDiscount 
         Height          =   435
         Left            =   5940
         TabIndex        =   5
         Top             =   2100
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtNetTotalPrice 
         Height          =   435
         Left            =   2040
         TabIndex        =   6
         Top             =   2550
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtNote 
         Height          =   435
         Left            =   2040
         TabIndex        =   8
         Top             =   3480
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin VB.Label lblNote 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   480
         TabIndex        =   25
         Top             =   3480
         Width           =   1485
      End
      Begin VB.Label Label3 
         Height          =   375
         Left            =   4200
         TabIndex        =   24
         Top             =   2580
         Width           =   465
      End
      Begin VB.Label lblNetTotalPrice 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   450
         TabIndex        =   23
         Top             =   2610
         Width           =   1485
      End
      Begin VB.Label Label4 
         Height          =   375
         Left            =   7440
         TabIndex        =   22
         Top             =   2130
         Width           =   465
      End
      Begin VB.Label lblDiscount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4800
         TabIndex        =   21
         Top             =   2160
         Width           =   1005
      End
      Begin VB.Label lblUnit 
         Height          =   375
         Left            =   4080
         TabIndex        =   20
         Top             =   1260
         Width           =   1005
      End
      Begin VB.Label lblTotalPrice 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   450
         TabIndex        =   19
         Top             =   2160
         Width           =   1485
      End
      Begin VB.Label Label1 
         Height          =   375
         Left            =   4200
         TabIndex        =   18
         Top             =   2130
         Width           =   465
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2760
         TabIndex        =   9
         Top             =   4080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditBLImportItemEx.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4440
         TabIndex        =   10
         Top             =   4080
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblPrice 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   435
         TabIndex        =   17
         Top             =   1710
         Width           =   1485
      End
      Begin VB.Label lblPartType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   75
         TabIndex        =   16
         Top             =   330
         Width           =   1845
      End
      Begin VB.Label lblPart 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   15
         Top             =   810
         Width           =   1725
      End
      Begin VB.Label lblQuantity 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   435
         TabIndex        =   14
         Top             =   1260
         Width           =   1485
      End
      Begin VB.Label lblLocation 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   435
         TabIndex        =   13
         Top             =   3090
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmAddEditBLImportItemEx"
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
Public COMMIT_FLAG As String
Public SupplierID As Long

Private m_PartTypes As Collection
Private m_Parts As Collection
Private m_Locations As Collection
Private m_Packagings As Collection
Private m_PartItemSpecs As Collection
Private m_PurchaseExpenses As Collection

Private TotalPrice As Double
Private Price As Double

Private Sub cboTextType_Click()
   m_HasModify = True
End Sub

Private Sub cboCountry_Click()
   m_HasModify = True
End Sub

Private Sub chkBangkok_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cboCalculateType_KeyPress(KeyAscii As Integer)
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

Private Sub InitFormLayout()
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)

   Me.KeyPreview = True
   pnlHeader.Caption = HeaderText
   Me.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_HEAD_COLOR
   SSFrame1.BackColor = GLB_FORM_COLOR
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
      
   Call InitNormalLabel(lblPartType, MapText("ประเภทวัสดุอุปกรณ์"))
   Call InitNormalLabel(lblPart, MapText("วัสดุอุปกรณ์"))
   Call InitNormalLabel(lblQuantity, MapText("ปริมาณนำเข้า"))
   Call InitNormalLabel(lblPrice, MapText("ราคา/หน่วย"))
   Call InitNormalLabel(lblTotalPrice, MapText("ราคารวม"))
   Call InitNormalLabel(lblLocation, MapText("สถานที่จัดเก็บ"))
   Call InitNormalLabel(Label1, MapText("บาท"))
   Call InitNormalLabel(Label4, MapText("บาท"))
   Call InitNormalLabel(Label3, MapText("บาท"))
   Call InitNormalLabel(lblUnit, MapText(""))
   Call InitNormalLabel(lblDiscount, MapText("ส่วนลด"))
   Call InitNormalLabel(lblNetTotalPrice, MapText("มูลค่าสุทธิ"))
   Call InitNormalLabel(lblNote, MapText("หมายเหตุ"))
   
   Call txtQuantity.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
  Call txtPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtPrice.Enabled = True
 Call txtTotalPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtDiscount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtNetTotalPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtNetTotalPrice.Enabled = False
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Dim EnpAddr As CSupItem
         
         Set EnpAddr = TempCollection.Item(id)
         
         uctlPartTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPartTypeLookup.MyCombo, EnpAddr.PART_TYPE)
         uctlPartLookup.MyCombo.ListIndex = IDToListIndex(uctlPartLookup.MyCombo, EnpAddr.PART_ITEM_ID)
         uctlLocationLookup.MyCombo.ListIndex = IDToListIndex(uctlLocationLookup.MyCombo, EnpAddr.LOCATION_ID)
            
         txtQuantity.Text = EnpAddr.TX_AMOUNT
         txtPrice.Text = EnpAddr.RAW_PRICE
         txtTotalPrice.Text = EnpAddr.RAW_TOT_PRICE
         txtDiscount.Text = EnpAddr.DISCOUNT_AMT
         txtNetTotalPrice.Text = EnpAddr.TOTAL_ACTUAL_PRICE
         txtNote.Text = EnpAddr.PART_NOTE

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

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long
Dim Pi As CPartItem

   If Not VerifyCombo(lblPartType, uctlPartTypeLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblPart, uctlPartLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblQuantity, txtQuantity, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblTotalPrice, txtTotalPrice, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblPrice, txtPrice, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblLocation, uctlLocationLookup.MyCombo, False) Then
      Exit Function
   End If

   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim EnpAddress As CSupItem
   If ShowMode = SHOW_ADD Then
      Set EnpAddress = New CSupItem
      EnpAddress.Flag = "A"
      Call TempCollection.add(EnpAddress)
   Else
      Set EnpAddress = TempCollection.Item(id)
      If EnpAddress.Flag <> "A" Then
         EnpAddress.Flag = "E"
      End If
   End If

   EnpAddress.PART_TYPE = uctlPartTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPartTypeLookup.MyCombo.ListIndex))
   EnpAddress.PART_ITEM_ID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
   EnpAddress.LOCATION_ID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
   EnpAddress.TX_AMOUNT = Val(txtQuantity.Text)
   EnpAddress.DISCOUNT_AMT = Val(txtDiscount.Text)
   EnpAddress.RAW_PRICE = Val(txtPrice.Text)
   EnpAddress.ACTUAL_UNIT_PRICE = MyDiffEx(Val(txtNetTotalPrice.Text), Val(txtQuantity.Text))
   EnpAddress.PART_TYPE_NAME = uctlPartLookup.MyCombo.Text
   EnpAddress.PART_NOTE = txtNote.Text
   EnpAddress.LOCATION_NAME = uctlLocationLookup.MyCombo.Text
   EnpAddress.PART_NO = uctlPartLookup.MyTextBox.Text
   EnpAddress.PART_DESC = uctlPartLookup.MyCombo.Text
   EnpAddress.CALCULATE_FLAG = "N"
   EnpAddress.TOTAL_ACTUAL_PRICE = Val(txtNetTotalPrice.Text)
   EnpAddress.RAW_TOT_PRICE = Val(txtTotalPrice.Text)
   EnpAddress.TX_TYPE = "I"
   Set Pi = GetPartItem(m_Parts, Trim(str(EnpAddress.PART_ITEM_ID)))
   EnpAddress.PIG_FLAG = Pi.PIG_FLAG
   'EnpAddress.PROJECT_NAME_ID = -1
   Set EnpAddress = Nothing
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadPartType(uctlPartTypeLookup.MyCombo, m_PartTypes)
      Set uctlPartTypeLookup.MyCollection = m_PartTypes
      
      Call LoadLocation(uctlLocationLookup.MyCombo, m_Locations, 2)
      Set uctlLocationLookup.MyCollection = m_Locations
            
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         id = 0
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

   OKClick = False
   Call InitFormLayout
   
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_PartTypes = New Collection
   Set m_Parts = New Collection
   Set m_Locations = New Collection
   Set m_Packagings = New Collection
   Set m_PartItemSpecs = New Collection
   Set m_PurchaseExpenses = New Collection
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
   Set m_Packagings = Nothing
   Set m_PartItemSpecs = Nothing
   Set m_PurchaseExpenses = Nothing
End Sub



Private Sub txtDiscount_Change()
   m_HasModify = True
   txtNetTotalPrice.Text = Val(txtTotalPrice.Text) - Val(txtDiscount.Text)
End Sub



Private Sub txtNote_Change()
m_HasModify = True
End Sub





Private Sub txtPrice_KeyPress(KeyAscii As Integer)
m_HasModify = True
If KeyAscii = 13 Then
      txtTotalPrice.Text = Val(txtPrice.Text) * Val(txtQuantity.Text)
      txtNetTotalPrice.Text = Val(txtTotalPrice.Text) - Val(txtDiscount.Text)
End If
End Sub


Private Sub txtPrice_LostFocus()
      txtTotalPrice.Text = Val(txtPrice.Text) * Val(txtQuantity.Text)
      txtNetTotalPrice.Text = Val(txtTotalPrice.Text) - Val(txtDiscount.Text)
End Sub

Private Sub txtQuantity_Change()
   m_HasModify = True
   txtTotalPrice.Text = Val(txtQuantity.Text) * Val(txtPrice.Text)
End Sub

Private Sub txtTotalPrice_KeyPress(KeyAscii As Integer)
m_HasModify = True
If KeyAscii = 13 Then
      txtPrice.Text = MyDiffEx(Val(txtTotalPrice.Text), Val(txtQuantity.Text))
      txtNetTotalPrice.Text = Val(txtTotalPrice.Text) - Val(txtDiscount.Text)
End If
End Sub

Private Sub txtTotalPrice_LostFocus()
      txtPrice.Text = MyDiffEx(Val(txtTotalPrice.Text), Val(txtQuantity.Text))
      txtNetTotalPrice.Text = Val(txtTotalPrice.Text) - Val(txtDiscount.Text)
End Sub

Private Sub uctlLocationLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlPartLookup_Change()
Dim Pi As CPartItem
Dim PartItemID As Long

   PartItemID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
   If PartItemID > 0 Then
      Set Pi = GetPartItem(m_Parts, Trim(str(PartItemID)))
      Call InitNormalLabel(lblUnit, Pi.UNIT_NAME)
   End If
   
   m_HasModify = True
End Sub

Private Sub uctlPartTypeLookup_Change()
Dim PartTypeID As Long
Dim Pt As CPartType

   PartTypeID = uctlPartTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPartTypeLookup.MyCombo.ListIndex))
   
   If PartTypeID > 0 Then
      Set Pt = GetPartType(m_PartTypes, Trim(str(PartTypeID)))
      Call LoadPartItem(uctlPartLookup.MyCombo, m_Parts, PartTypeID, "N")
      Set uctlPartLookup.MyCollection = m_Parts
   
      Call LoadLocation(uctlLocationLookup.MyCombo, m_Locations, 2, , , Pt.PART_GROUP_ID)
      Set uctlLocationLookup.MyCollection = m_Locations
   End If
   
   m_HasModify = True
End Sub
