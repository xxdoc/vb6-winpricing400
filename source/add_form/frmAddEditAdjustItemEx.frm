VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditAdjustItemEx 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5040
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
   Icon            =   "frmAddEditAdjustItemEx.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   9435
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   4455
      Left            =   0
      TabIndex        =   13
      Top             =   600
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   7858
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextLookup uctlPartTypeLookup 
         Height          =   435
         Left            =   1785
         TabIndex        =   0
         Top             =   300
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPrice 
         Height          =   435
         Left            =   1785
         TabIndex        =   8
         Top             =   3000
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtQuantity 
         Height          =   435
         Left            =   1785
         TabIndex        =   7
         Top             =   2550
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPartLookup 
         Height          =   435
         Left            =   1785
         TabIndex        =   1
         Top             =   750
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlLocationLookup 
         Height          =   435
         Left            =   1785
         TabIndex        =   2
         Top             =   1200
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtActualPrice 
         Height          =   435
         Left            =   1770
         TabIndex        =   5
         Top             =   2100
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtActualAmount 
         Height          =   435
         Left            =   1770
         TabIndex        =   3
         Top             =   1650
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtCurrentPrice 
         Height          =   435
         Left            =   6330
         TabIndex        =   6
         Top             =   2100
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtCurrentAmount 
         Height          =   435
         Left            =   6330
         TabIndex        =   4
         Top             =   1650
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   767
      End
      Begin VB.Label lblCurrentAmount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4740
         TabIndex        =   25
         Top             =   1710
         Width           =   1485
      End
      Begin VB.Label lblCurrentPrice 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4740
         TabIndex        =   24
         Top             =   2160
         Width           =   1485
      End
      Begin VB.Label Label5 
         Height          =   375
         Left            =   8835
         TabIndex        =   23
         Top             =   2100
         Width           =   465
      End
      Begin VB.Label lblActualAmount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   180
         TabIndex        =   22
         Top             =   1710
         Width           =   1485
      End
      Begin VB.Label lblActualPrice 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   180
         TabIndex        =   21
         Top             =   2160
         Width           =   1485
      End
      Begin VB.Label Label2 
         Height          =   375
         Left            =   3855
         TabIndex        =   20
         Top             =   2100
         Width           =   1005
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   2265
         TabIndex        =   9
         Top             =   3630
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditAdjustItemEx.frx":08CA
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Height          =   375
         Left            =   3870
         TabIndex        =   19
         Top             =   3000
         Width           =   1005
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3915
         TabIndex        =   10
         Top             =   3630
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditAdjustItemEx.frx":0BE4
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5565
         TabIndex        =   11
         Top             =   3630
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblPrice 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   18
         Top             =   3060
         Width           =   1485
      End
      Begin VB.Label lblPartType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   17
         Top             =   330
         Width           =   1485
      End
      Begin VB.Label lblPart 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   16
         Top             =   810
         Width           =   1485
      End
      Begin VB.Label lblQuantity 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   15
         Top             =   2610
         Width           =   1485
      End
      Begin VB.Label lblLocation 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   14
         Top             =   1260
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmAddEditAdjustItemEx"
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
Public TxSeq As Long
Public COMMIT_FLAG As String
Public ParentForm As Form
Public FROM_DATE As Date

Private m_BalanceItems As Collection
Private m_PartTypes As Collection
Private m_Parts As Collection
Private m_Locations As Collection
Private m_Layout As Collection

Private Sub cboTextType_Click()
   m_HasModify = True
End Sub

Private Sub cboCountry_Click()
   m_HasModify = True
End Sub

Private Sub chkBangkok_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkAvg_Click(Value As Integer)
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
      
   Call InitNormalLabel(lblPartType, MapText("ประเภทวัตถุดิบ"))
   Call InitNormalLabel(lblPart, MapText("วัตถุดิบ"))
   Call InitNormalLabel(lblQuantity, MapText("ปริมาณ"))
   Call InitNormalLabel(lblPrice, MapText("มูลค่า"))
   Call InitNormalLabel(lblLocation, MapText("สถานที่จัดเก็บ"))
   Call InitNormalLabel(Label1, MapText("บาท"))
   Call InitNormalLabel(Label2, MapText("บาท"))
   Call InitNormalLabel(Label5, MapText("บาท"))
   Call InitNormalLabel(lblActualAmount, MapText("จำนวนตรวจนับ"))
   Call InitNormalLabel(lblActualPrice, MapText("มูลค่าตรวจนับ"))
   Call InitNormalLabel(lblCurrentAmount, MapText("จำนวนในระบบ"))
   Call InitNormalLabel(lblCurrentPrice, MapText("มูลค่าในระบบ"))
   
   Call txtQuantity.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtActualAmount.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtActualPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtCurrentAmount.SetTextLenType(TEXT_FLOAT, glbSetting.CODE_TYPE)
   txtCurrentAmount.Enabled = False
   Call txtCurrentPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.CODE_TYPE)
   txtCurrentPrice.Enabled = False
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdNext, MapText("ถัดไป"))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Dim EnpAddr As CLotItem
         
         Set EnpAddr = TempCollection.Item(ID)
         
         uctlPartTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPartTypeLookup.MyCombo, EnpAddr.PART_TYPE)
         uctlPartLookup.MyCombo.ListIndex = IDToListIndex(uctlPartLookup.MyCombo, EnpAddr.PART_ITEM_ID)
         uctlLocationLookup.MyCombo.ListIndex = IDToListIndex(uctlLocationLookup.MyCombo, EnpAddr.LOCATION_ID)
         
         txtQuantity.Text = EnpAddr.TX_AMOUNT
         txtPrice.Text = EnpAddr.ACTUAL_UNIT_PRICE
         txtCurrentAmount.Text = EnpAddr.CURRENT_AMOUNT
         txtCurrentPrice.Text = EnpAddr.CURRENT_PRICE
         txtActualAmount.Text = EnpAddr.ACTUAL_AMOUNT
         txtActualPrice.Text = EnpAddr.ACTUAL_PRICE
         
         cmdOK.Enabled = (COMMIT_FLAG <> "Y")
      End If
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cmdLayout_Click()
Dim OKClick As Boolean
Dim LayoutID As Long

   If Not VerifyCombo(lblPart, uctlPartLookup.MyCombo, False) Then
      Exit Sub
   End If
   If Not VerifyCombo(lblLocation, uctlLocationLookup.MyCombo, False) Then
      Exit Sub
   End If
   
   frmLayoutSearch.PartItemID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
   frmLayoutSearch.LocationID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
   Load frmLayoutSearch
   frmLayoutSearch.Show 1
   
   OKClick = frmLayoutSearch.OKClick
   LayoutID = frmLayoutSearch.LayoutID
   
   Unload frmLayoutSearch
   Set frmLayoutSearch = Nothing
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Function GetNextID(OldID As Long) As Long
Dim Ei As CExtractItem
Dim TempIndex As Long
Dim J As Long

   If OldID >= TempCollection.Count Then
      J = TempCollection.Count
      glbErrorLog.LocalErrorMsg = "ถึงเรคคอร์ดสุดท้ายแล้ว"
      glbErrorLog.ShowUserError
   Else
      J = OldID + 1
   End If
   
   GetNextID = J
End Function

Private Sub cmdNext_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   Call ParentForm.ShowGridItem
   
   txtCurrentAmount.Text = ""
   txtCurrentPrice.Text = ""
   txtActualAmount.Text = ""
   txtActualPrice.Text = ""
   txtQuantity.Text = ""
   txtPrice.Text = ""
   
   uctlLocationLookup.SetFocus
   
   Call QueryData(True)
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

   If Not VerifyCombo(lblPartType, uctlPartTypeLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblPart, uctlPartLookup.MyCombo, False) Then
      Exit Function
   End If
'   If Not VerifyTextControl(lblQuantity, txtQuantity, False) Then
'      Exit Function
'   End If
'   If Not VerifyTextControl(lblPrice, txtPrice, False) Then
'      Exit Function
'   End If
   If Not VerifyCombo(lblLocation, uctlLocationLookup.MyCombo, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim EnpAddress As CLotItem
   If ShowMode = SHOW_ADD Then
      Set EnpAddress = New CLotItem
      EnpAddress.Flag = "A"
      EnpAddress.TRANSACTION_SEQ = TxSeq
      If Val(txtQuantity.Text) >= 0 Then
         Call TempCollection.add(EnpAddress)
      Else
         Call TempCollection2.add(EnpAddress)
      End If
   Else
'      Set EnpAddress = TempCollection.Item(ID)
'      If EnpAddress.Flag <> "A" Then
'         EnpAddress.Flag = "E"
'      End If
   End If

   EnpAddress.PART_TYPE = uctlPartTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPartTypeLookup.MyCombo.ListIndex))
   EnpAddress.PART_ITEM_ID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
   EnpAddress.LOCATION_ID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
   EnpAddress.PART_TYPE_NAME = uctlPartLookup.MyCombo.Text
   EnpAddress.LOCATION_NAME = uctlLocationLookup.MyCombo.Text
   EnpAddress.PART_NO = uctlPartLookup.MyTextBox.Text
   EnpAddress.PART_DESC = uctlPartLookup.MyCombo.Text
   EnpAddress.TX_AMOUNT = Abs(Val(txtQuantity.Text))
   If Val(txtQuantity.Text) > 0 Then
      EnpAddress.TX_TYPE = "I"
      EnpAddress.TOTAL_INCLUDE_PRICE = Val(txtPrice.Text)
   Else
      EnpAddress.TX_TYPE = "E"
      EnpAddress.TOTAL_INCLUDE_PRICE = -1 * Val(txtPrice.Text)
   End If
   EnpAddress.ACTUAL_UNIT_PRICE = MyDiffEx(EnpAddress.TOTAL_INCLUDE_PRICE, EnpAddress.TX_AMOUNT)

'   If Val(txtQuantity.Text) >= 0 Then
'      EnpAddress.TX_TYPE = "I"
'      If Val(txtPrice.Text) >= 0 Then
'         EnpAddress.ACTUAL_UNIT_PRICE = MyDiffEx(Abs(Val(txtPrice.Text)), Abs(Val(txtQuantity.Text)))
'         EnpAddress.TOTAL_ACTUAL_PRICE = Abs(Val(txtPrice.Text))
'      Else
'         EnpAddress.ACTUAL_UNIT_PRICE = MyDiffEx(Abs(Val(txtPrice.Text)), Abs(Val(txtQuantity.Text)))
'         EnpAddress.TOTAL_ACTUAL_PRICE = -1 * Abs(Val(txtPrice.Text))
'      End If
'   Else
'      EnpAddress.TX_TYPE = "E"
'      If Val(txtPrice.Text) >= 0 Then
'         EnpAddress.ACTUAL_UNIT_PRICE = MyDiffEx(Abs(Val(txtPrice.Text)), Abs(Val(txtQuantity.Text)))
'         EnpAddress.TOTAL_ACTUAL_PRICE = -1 * Abs(Val(txtPrice.Text))
'      Else
'         EnpAddress.ACTUAL_UNIT_PRICE = MyDiffEx(Abs(Val(txtPrice.Text)), Abs(Val(txtQuantity.Text)))
'         EnpAddress.TOTAL_ACTUAL_PRICE = Abs(Val(txtPrice.Text))
'      End If
'   End If
   EnpAddress.INCLUDE_UNIT_PRICE = EnpAddress.ACTUAL_UNIT_PRICE
   EnpAddress.TOTAL_ACTUAL_PRICE = EnpAddress.TOTAL_INCLUDE_PRICE
   EnpAddress.CALCULATE_FLAG = "N"
   EnpAddress.CURRENT_AMOUNT = Val(txtCurrentAmount.Text)
   EnpAddress.CURRENT_PRICE = Val(txtCurrentPrice.Text)
   EnpAddress.ACTUAL_AMOUNT = Val(txtActualAmount.Text)
   EnpAddress.ACTUAL_PRICE = Val(txtActualPrice.Text)
   
   Set EnpAddress = Nothing
   SaveData = True
End Function

Private Sub Form_Activate()
Dim BalanceAccums As Collection
Static BalanceItems As Collection

   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadPartType(uctlPartTypeLookup.MyCombo, m_PartTypes)
      Set uctlPartTypeLookup.MyCollection = m_PartTypes
      
      Call LoadLocation(uctlLocationLookup.MyCombo, m_Locations, 2)
      Set uctlLocationLookup.MyCollection = m_Locations
      
      Set BalanceAccums = New Collection
      If FROM_DATE > 0 Then
         If BalanceItems Is Nothing Then
            Set BalanceItems = New Collection
            FROM_DATE = DateAdd("D", 1, FROM_DATE)
            Call LoadInventoryBalanceEx(Nothing, BalanceAccums, FROM_DATE, -1, "")
            Call glbDaily.CopyBalanceAccum(BalanceAccums, BalanceItems)
            Set m_BalanceItems = BalanceItems
         End If
      End If
      Set BalanceAccums = Nothing
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
         Call QueryData(True)
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
   Set m_Rs = New ADODB.Recordset
   Set m_PartTypes = New Collection
   Set m_Parts = New Collection
   Set m_Locations = New Collection
   Set m_Layout = New Collection
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
   Set m_Layout = Nothing
   Set m_BalanceItems = Nothing
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

Private Sub SSCommand1_Click()

End Sub

Private Sub txtActualAmount_Change()
   m_HasModify = True
   txtQuantity.Text = Val(txtActualAmount.Text) - Val(txtCurrentAmount.Text)
End Sub

Private Sub txtActualPrice_Change()
   m_HasModify = True
   txtPrice.Text = Val(txtActualPrice.Text) - Val(txtCurrentPrice.Text)
End Sub

Private Sub txtCurrentAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtCurrentPrice_Change()
   m_HasModify = True
End Sub

Private Sub txtQuantity_Change()
   m_HasModify = True
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

Private Sub uctlLayoutLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlLocationLookup_Change()
Dim Li As CLotItem
Dim LocationID As Long
Dim PartItemID As Long

   If uctlLocationLookup.MyCombo.ListIndex <= 0 Then
      Exit Sub
   End If
   If uctlPartLookup.MyCombo.ListIndex <= 0 Then
      Exit Sub
   End If
   
   LocationID = uctlLocationLookup.MyCombo.ItemData(uctlLocationLookup.MyCombo.ListIndex)
   PartItemID = uctlPartLookup.MyCombo.ItemData(uctlPartLookup.MyCombo.ListIndex)
   
   Set Li = GetLotItem(m_BalanceItems, LocationID & "-" & PartItemID)
   txtCurrentAmount.Text = Li.NEW_AMOUNT
   txtCurrentPrice.Text = Li.NEW_AMOUNT * Li.NEW_PRICE
   
   m_HasModify = True
End Sub

Private Sub uctlPartLookup_Change()
Dim Li As CLotItem
Dim LocationID As Long
Dim PartItemID As Long

   If uctlLocationLookup.MyCombo.ListIndex <= 0 Then
      Exit Sub
   End If
   If uctlPartLookup.MyCombo.ListIndex <= 0 Then
      Exit Sub
   End If
   
   LocationID = uctlLocationLookup.MyCombo.ItemData(uctlLocationLookup.MyCombo.ListIndex)
   PartItemID = uctlPartLookup.MyCombo.ItemData(uctlPartLookup.MyCombo.ListIndex)
   
   Set Li = GetLotItem(m_BalanceItems, LocationID & "-" & PartItemID)
   txtCurrentAmount.Text = Li.NEW_AMOUNT
   txtCurrentPrice.Text = Li.NEW_AMOUNT * Li.NEW_PRICE
   
   m_HasModify = True
End Sub

Private Sub uctlPartTypeLookup_Change()
Dim PartTypeID As Long

   PartTypeID = uctlPartTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPartTypeLookup.MyCombo.ListIndex))
   
   Call LoadPartItem(uctlPartLookup.MyCombo, m_Parts, PartTypeID, "")
   Set uctlPartLookup.MyCollection = m_Parts
   
   m_HasModify = True
End Sub
