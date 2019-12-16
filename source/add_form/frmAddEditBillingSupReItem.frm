VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditBillingSupReItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11280
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditBillingSupReItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   11280
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   4275
      Left            =   0
      TabIndex        =   11
      Top             =   600
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   7541
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextLookup uctlPartTypeLookup 
         Height          =   435
         Left            =   1785
         TabIndex        =   0
         Top             =   270
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPrice 
         Height          =   435
         Left            =   1785
         TabIndex        =   5
         Top             =   2540
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtQuantity 
         Height          =   435
         Left            =   1785
         TabIndex        =   3
         Top             =   1650
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPartLookup 
         Height          =   435
         Left            =   1785
         TabIndex        =   1
         Top             =   720
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlLocationLookup 
         Height          =   435
         Left            =   1785
         TabIndex        =   2
         Top             =   1170
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtActualPrice 
         Height          =   435
         Left            =   1770
         TabIndex        =   6
         Top             =   2100
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtDesc 
         Height          =   435
         Left            =   1770
         TabIndex        =   7
         Top             =   2970
         Width           =   7395
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   180
         TabIndex        =   21
         Top             =   3030
         Width           =   1485
      End
      Begin VB.Label lblUnit 
         Height          =   375
         Left            =   3870
         TabIndex        =   20
         Top             =   1710
         Width           =   1275
      End
      Begin Threed.SSCommand cmdLotSelect 
         Height          =   405
         Left            =   5790
         TabIndex        =   4
         Top             =   1680
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditBillingSupReItem.frx":08CA
         ButtonStyle     =   3
      End
      Begin VB.Label lblActualPrice 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   180
         TabIndex        =   19
         Top             =   2130
         Width           =   1485
      End
      Begin VB.Label Label2 
         Height          =   375
         Left            =   3825
         TabIndex        =   18
         Top             =   2580
         Width           =   1005
      End
      Begin VB.Label Label1 
         Height          =   375
         Left            =   3840
         TabIndex        =   17
         Top             =   2130
         Width           =   1005
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3090
         TabIndex        =   8
         Top             =   3480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditBillingSupReItem.frx":0BE4
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4740
         TabIndex        =   9
         Top             =   3480
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
         TabIndex        =   16
         Top             =   2640
         Width           =   1485
      End
      Begin VB.Label lblPartType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   15
         Top             =   300
         Width           =   1485
      End
      Begin VB.Label lblPart 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   14
         Top             =   780
         Width           =   1485
      End
      Begin VB.Label lblQuantity 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   13
         Top             =   1710
         Width           =   1485
      End
      Begin VB.Label lblLocation 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   180
         TabIndex        =   12
         Top             =   1200
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmAddEditBillingSupReItem"
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
Public COMMIT_FLAG As String

Private m_PartTypes As Collection
Private m_Parts As Collection
Private m_Locations As Collection
Private m_Houses As Collection
Private m_Pigs As Collection
Private m_PigTypes As Collection
Private m_SubLotItems As Collection
Private m_ManualFlag As Boolean

Private m_TempCollection As Collection

Private Sub cboTextType_Click()
   m_HasModify = True
End Sub

Private Sub cboCountry_Click()
   m_HasModify = True
End Sub

Private Sub chkBangkok_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cboDepartment_Click()
   m_HasModify = True
End Sub

Private Sub cboExpenseType_Click()
   m_HasModify = True
End Sub
Private Sub cboItemDesc_Click()
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
   Call InitNormalLabel(lblPrice, MapText("ราคา"))
   Call InitNormalLabel(lblActualPrice, MapText("ราคาเฉลี่ย"))
   Call InitNormalLabel(lblLocation, MapText("จากคลัง"))
   Call InitNormalLabel(Label1, MapText("บาท"))
   Call InitNormalLabel(Label2, MapText("บาท"))
   Call InitNormalLabel(lblUnit, MapText(""))
   Call InitNormalLabel(lblDesc, MapText("รายละเอียด"))
   
   Call txtQuantity.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtPrice.Enabled = False
   Call txtActualPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtActualPrice.Enabled = True
   Call txtActualPrice.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdLotSelect.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdLotSelect, MapText("..."))
End Sub

Private Sub CopyCollection(SourceCol As Collection, DestCol As Collection)
Dim D As Object

   Set DestCol = Nothing
   Set DestCol = New Collection
   
   For Each D In SourceCol
      Call DestCol.add(D)
   Next D
End Sub
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Dim EnpAddr As CReceiptItem

         Set EnpAddr = TempCollection.Item(ID)
         
         uctlPartTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPartTypeLookup.MyCombo, EnpAddr.PART_TYPE)
         uctlPartLookup.MyCombo.ListIndex = IDToListIndex(uctlPartLookup.MyCombo, EnpAddr.PART_ITEM_ID)
         uctlLocationLookup.MyCombo.ListIndex = IDToListIndex(uctlLocationLookup.MyCombo, EnpAddr.LOCATION_ID)

         txtQuantity.Text = EnpAddr.RETURN_AMOUNT
         txtPrice.Text = EnpAddr.RETURN_TOTAL_PRICE
         txtActualPrice.Text = EnpAddr.RETURN_AVG_PRICE
         txtDesc.Text = EnpAddr.ITEM_DESC
         'cboDepartment.ListIndex = IDToListIndex(cboDepartment, EnpAddr.TO_DEPARTMENT)
         'cboExpenseType.ListIndex = IDToListIndex(cboExpenseType, EnpAddr.EXPENSE_TYPE)
         
         'cboItemDesc.ListIndex = IDToListIndex(cboItemDesc, EnpAddr.ITEM_DESC_ID)
         
         'm_ManualFlag = (EnpAddr.LOT_MANUAL = "Y")
         cmdOK.Enabled = (COMMIT_FLAG <> "Y")
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
   If Not VerifyCombo(lblLocation, uctlLocationLookup.MyCombo, False) Then
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
   frmAddSubLotItem.LocationID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
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
Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long
Dim Pi As CPartItem

   If Not VerifyCombo(lblLocation, uctlLocationLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblPartType, uctlPartTypeLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblPart, uctlPartLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblQuantity, txtQuantity, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblPrice, txtPrice, True) Then
      Exit Function
   End If

   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim EnpAddress As CReceiptItem
   If ShowMode = SHOW_ADD Then
      Set EnpAddress = New CReceiptItem
      EnpAddress.Flag = "A"
      Call TempCollection.add(EnpAddress)
   Else
      Set EnpAddress = TempCollection.Item(ID)
      If EnpAddress.Flag <> "A" Then
         EnpAddress.Flag = "E"
      End If
   End If

   EnpAddress.PART_TYPE = uctlPartTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPartTypeLookup.MyCombo.ListIndex))
   EnpAddress.PART_ITEM_ID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
   EnpAddress.LOCATION_ID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
   EnpAddress.RETURN_AMOUNT = txtQuantity.Text
   EnpAddress.RETURN_TOTAL_PRICE = Val(txtPrice.Text)
   EnpAddress.RETURN_AVG_PRICE = Val(txtActualPrice.Text)
   'EnpAddress.PART_TYPE_NAME = uctlPartLookup.MyCombo.Text
   EnpAddress.LOCATION_NAME = uctlLocationLookup.MyCombo.Text
   EnpAddress.PART_NO = uctlPartLookup.MyTextBox.Text
   EnpAddress.PART_DESC = uctlPartLookup.MyCombo.Text
   'EnpAddress.TO_DEPARTMENT = cboDepartment.ItemData(Minus2Zero(cboDepartment.ListIndex))
   EnpAddress.ITEM_DESC = txtDesc.Text
   'EnpAddress.CALCULATE_FLAG = "Y"
   'EnpAddress.TX_TYPE = "E"
   'EnpAddress.EXPENSE_TYPE = cboExpenseType.ItemData(Minus2Zero(cboExpenseType.ListIndex))
   Set Pi = GetPartItem(m_Parts, Trim(str(EnpAddress.PART_ITEM_ID)))
   'EnpAddress.PIG_FLAG = Pi.PIG_FLAG
   
   'EnpAddress.ITEM_DESC_ID = cboItemDesc.ItemData(Minus2Zero(cboItemDesc.ListIndex))
   'EnpAddress.ITEM_DESC_NAME = cboItemDesc.Text
   
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
   
   m_ManualFlag = False
   
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_PartTypes = New Collection
   Set m_Parts = New Collection
   Set m_Locations = New Collection
   Set m_Houses = New Collection
   Set m_Pigs = New Collection
   Set m_PigTypes = New Collection
   Set m_SubLotItems = New Collection
   Set m_TempCollection = New Collection
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
   Set m_PigTypes = Nothing
   Set m_SubLotItems = Nothing
   Set m_TempCollection = Nothing
End Sub

Private Sub txtDepartment_Click()
   m_HasModify = True
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

Private Sub txtActualPrice_Change()
   txtPrice.Text = Val(txtActualPrice.Text) * Val(txtQuantity.Text)
   m_HasModify = True
End Sub

Private Sub txtQuantity_Change()
   txtPrice.Text = Val(txtActualPrice.Text) * Val(txtQuantity.Text)
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

Private Sub uctlHouseLookup_Change()
   m_HasModify = True
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

Private Sub uctlPigWeekLookup_Change()
   m_HasModify = True
End Sub
