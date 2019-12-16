VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditTransferItem2 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6015
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
   Icon            =   "frmAddEditTransferItem2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   9435
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   5445
      Left            =   0
      TabIndex        =   15
      Top             =   600
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   9604
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
         Top             =   2100
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
      Begin prjFarmManagement.uctlTextLookup uctlToLocationLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   9
         Top             =   3480
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPackaging 
         Height          =   435
         Left            =   5820
         TabIndex        =   6
         Top             =   2100
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlToPartTypeLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   7
         Top             =   2580
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlToPartItemLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   8
         Top             =   3030
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtManualPrice 
         Height          =   435
         Left            =   1800
         TabIndex        =   11
         Top             =   3930
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin Threed.SSCheck chkManualPrice 
         Height          =   435
         Left            =   7350
         TabIndex        =   10
         Top             =   3480
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblManualPrice 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   27
         Top             =   3990
         Width           =   1485
      End
      Begin VB.Label Label4 
         Height          =   375
         Left            =   3855
         TabIndex        =   26
         Top             =   3960
         Width           =   1005
      End
      Begin VB.Label lblToPartItem 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   25
         Top             =   3090
         Width           =   1485
      End
      Begin VB.Label lblToPartType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   24
         Top             =   2610
         Width           =   1485
      End
      Begin VB.Label lblPackaging 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4530
         TabIndex        =   23
         Top             =   2160
         Width           =   1185
      End
      Begin Threed.SSCommand cmdLotSelect 
         Height          =   405
         Left            =   3780
         TabIndex        =   4
         Top             =   1680
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditTransferItem2.frx":08CA
         ButtonStyle     =   3
      End
      Begin VB.Label lblToLocation 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   22
         Top             =   3540
         Width           =   1485
      End
      Begin VB.Label Label1 
         Height          =   375
         Left            =   3840
         TabIndex        =   21
         Top             =   2130
         Width           =   1005
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3090
         TabIndex        =   12
         Top             =   4620
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditTransferItem2.frx":0BE4
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4740
         TabIndex        =   13
         Top             =   4620
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
         TabIndex        =   20
         Top             =   2160
         Width           =   1485
      End
      Begin VB.Label lblPartType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   19
         Top             =   300
         Width           =   1485
      End
      Begin VB.Label lblPart 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   18
         Top             =   780
         Width           =   1485
      End
      Begin VB.Label lblQuantity 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   17
         Top             =   1710
         Width           =   1485
      End
      Begin VB.Label lblLocation 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   180
         TabIndex        =   16
         Top             =   1200
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmAddEditTransferItem2"
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

Private m_PartTypes As Collection
Private m_Parts As Collection
Private m_Locations As Collection
Private m_Houses As Collection
Private m_Pigs As Collection
Private m_Layout As Collection
Private m_SubLotItems As Collection
Private m_ManualFlag As Boolean
Private m_ToPartTypes As Collection
Private m_ToParts As Collection
Private m_ToLocations As Collection

Private Sub cboTextType_Click()
   m_HasModify = True
End Sub

Private Sub cboCountry_Click()
   m_HasModify = True
End Sub

Private Sub chkBangkok_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkManualPrice_Click(Value As Integer)
   txtManualPrice.Enabled = (Check2Flag(CLng(Value)) = "Y")
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

   Call InitNormalLabel(lblPartType, MapText("�������ѵ�شԺ"))
   Call InitNormalLabel(lblPart, MapText("�ѵ�شԺ"))
   Call InitNormalLabel(lblQuantity, MapText("����ҳ"))
   Call InitNormalLabel(lblPrice, MapText("�Ҥ�/˹���"))
   Call InitNormalLabel(lblLocation, MapText("�ҡ��ѧ"))
   Call InitNormalLabel(lblToLocation, MapText("��Ҥ�ѧ"))
   Call InitNormalLabel(Label1, MapText("�ҷ"))
   Call InitNormalLabel(lblPackaging, MapText("�ӹǹ�պ���"))
   Call InitNormalLabel(lblToPartType, MapText("�������ѵ�شԺ"))
   Call InitNormalLabel(lblToPartItem, MapText("�ѵ�شԺ"))
   Call InitNormalLabel(lblManualPrice, MapText("�Ҥ�/˹���"))
   Call InitNormalLabel(Label4, MapText("�ҷ"))
   
   Call txtQuantity.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtPrice.Enabled = False
   Call txtPackaging.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtManualPrice.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtManualPrice.Enabled = False
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdLotSelect.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitCheckBox(chkManualPrice, "��˹��Ҥ��ͧ")
   
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
'   Call InitMainButton(cmdLayout, MapText("..."))
   Call InitMainButton(cmdLotSelect, MapText("..."))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Dim EnpAddr As CTransferItem

         Set EnpAddr = TempCollection.Item(ID)
         Call CopySubLotItem(EnpAddr.ExportItem.SubLotItems, m_SubLotItems)

         uctlPartTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPartTypeLookup.MyCombo, EnpAddr.ExportItem.PART_TYPE)
         uctlPartLookup.MyCombo.ListIndex = IDToListIndex(uctlPartLookup.MyCombo, EnpAddr.ExportItem.PART_ITEM_ID)
         uctlLocationLookup.MyCombo.ListIndex = IDToListIndex(uctlLocationLookup.MyCombo, EnpAddr.ExportItem.LOCATION_ID)
         
         uctlToPartTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlToPartTypeLookup.MyCombo, EnpAddr.ImportItem.PART_TYPE)
         uctlToPartItemLookup.MyCombo.ListIndex = IDToListIndex(uctlToPartItemLookup.MyCombo, EnpAddr.ImportItem.PART_ITEM_ID)
         uctlToLocationLookup.MyCombo.ListIndex = IDToListIndex(uctlToLocationLookup.MyCombo, EnpAddr.ImportItem.LOCATION_ID)

         txtQuantity.Text = EnpAddr.ExportItem.TX_AMOUNT
         txtPrice.Text = EnpAddr.ExportItem.INCLUDE_UNIT_PRICE
         txtPackaging.Text = EnpAddr.ExportItem.PACKAGING_AMT
         txtManualPrice.Text = EnpAddr.ImportItem.INCLUDE_UNIT_PRICE
         
         chkManualPrice.Value = FlagToCheck(EnpAddr.ImportItem.MANUAL_PRICE)
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
'      uctlLayoutLookup.MyCombo.ListIndex = IDToListIndex(uctlLayoutLookup.MyCombo, LayoutID)
      m_HasModify = True
   End If
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
   
   frmAddSubLotItem.HeaderText = "���͡�����Ţ��͵"
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
   If Not VerifyCombo(lblLocation, uctlLocationLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblToLocation, uctlToLocationLookup.MyCombo, False) Then
      Exit Function
   End If
   
'   If uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex)) = _
'       uctlToLocationLookup.MyCombo.ItemData(Minus2Zero(uctlToLocationLookup.MyCombo.ListIndex)) Then
'         glbErrorLog.LocalErrorMsg = "�ç���͹��ҡѺ�ç���͹�͡�е�ͧᵡ��ҧ�ѹ"
'         glbErrorLog.ShowUserError
'
'         Exit Function
'   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim EnpAddress As CTransferItem
   Dim Ei As CLotItem
   Dim II As CLotItem
   If ShowMode = SHOW_ADD Then
      Set Ei = New CLotItem
      Set II = New CLotItem
      Set EnpAddress = New CTransferItem

      Ei.Flag = "A"
      Ei.CALCULATE_FLAG = "Y"
      II.Flag = "A"
      II.CALCULATE_FLAG = "Y"
      EnpAddress.Flag = "A"

      Set EnpAddress.ExportItem = Ei
      Set EnpAddress.ImportItem = II

      Call TempCollection.add(EnpAddress)
   Else
      Set EnpAddress = TempCollection.Item(ID)
      If EnpAddress.Flag <> "A" Then
         EnpAddress.Flag = "E"
         EnpAddress.ExportItem.Flag = "E"
         EnpAddress.ImportItem.Flag = "E"
      End If
   End If

   EnpAddress.ExportItem.PART_TYPE = uctlPartTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPartTypeLookup.MyCombo.ListIndex))
   EnpAddress.ExportItem.PART_ITEM_ID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
   EnpAddress.ExportItem.LOCATION_ID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
   EnpAddress.ExportItem.TX_AMOUNT = txtQuantity.Text
   EnpAddress.ExportItem.INCLUDE_UNIT_PRICE = Val(txtPrice.Text)
   EnpAddress.ExportItem.PART_TYPE_NAME = uctlPartLookup.MyCombo.Text
   EnpAddress.ExportItem.LOCATION_NAME = uctlLocationLookup.MyCombo.Text
   EnpAddress.ExportItem.PART_NO = uctlPartLookup.MyTextBox.Text
   EnpAddress.ExportItem.PART_DESC = uctlPartLookup.MyCombo.Text
   EnpAddress.ExportItem.TX_TYPE = "E"
   EnpAddress.ExportItem.PACKAGING_AMT = Val(txtPackaging.Text)
   EnpAddress.ExportItem.MANUAL_PRICE = Check2Flag(chkManualPrice.Value)
'   Call glbDaily.GenerateSubLot(EnpAddress.ExportItem, Nothing)

   EnpAddress.ImportItem.PART_TYPE = uctlToPartTypeLookup.MyCombo.ItemData(Minus2Zero(uctlToPartTypeLookup.MyCombo.ListIndex))
   EnpAddress.ImportItem.PART_ITEM_ID = uctlToPartItemLookup.MyCombo.ItemData(Minus2Zero(uctlToPartItemLookup.MyCombo.ListIndex))
   EnpAddress.ImportItem.LOCATION_ID = uctlToLocationLookup.MyCombo.ItemData(Minus2Zero(uctlToLocationLookup.MyCombo.ListIndex))
   EnpAddress.ImportItem.TX_AMOUNT = txtQuantity.Text
   EnpAddress.ImportItem.ACTUAL_UNIT_PRICE = Val(txtManualPrice.Text)
   EnpAddress.ImportItem.TOTAL_ACTUAL_PRICE = (txtQuantity.Text) * Val(txtManualPrice.Text)
   EnpAddress.ImportItem.INCLUDE_UNIT_PRICE = EnpAddress.ImportItem.ACTUAL_UNIT_PRICE
   EnpAddress.ImportItem.TOTAL_INCLUDE_PRICE = EnpAddress.ImportItem.TOTAL_ACTUAL_PRICE
   EnpAddress.ImportItem.PART_TYPE_NAME = uctlToPartTypeLookup.MyCombo.Text
   EnpAddress.ImportItem.LOCATION_NAME = uctlToLocationLookup.MyCombo.Text
   EnpAddress.ImportItem.PART_NO = uctlToPartItemLookup.MyTextBox.Text
   EnpAddress.ImportItem.PART_DESC = uctlToPartItemLookup.MyCombo.Text
   EnpAddress.ImportItem.TX_TYPE = "I"
   EnpAddress.ImportItem.PACKAGING_AMT = Val(txtPackaging.Text)
   EnpAddress.ImportItem.MANUAL_PRICE = Check2Flag(chkManualPrice.Value)

   Set EnpAddress = Nothing
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
'      Call LoadLayout(uctlLayoutLookup.MyCombo, m_Layout)
'      Set uctlLayoutLookup.MyCollection = m_Layout
      
      Call LoadPartType(uctlPartTypeLookup.MyCombo, m_PartTypes)
      Set uctlPartTypeLookup.MyCollection = m_PartTypes
      
      Call LoadLocation(uctlLocationLookup.MyCombo, m_Locations, 2)
      Set uctlLocationLookup.MyCollection = m_Locations

      Call LoadPartType(uctlToPartTypeLookup.MyCombo, m_ToPartTypes)
      Set uctlToPartTypeLookup.MyCollection = m_ToPartTypes
      
      Call LoadLocation(uctlToLocationLookup.MyCombo, m_Houses, 2)
      Set uctlToLocationLookup.MyCollection = m_Houses

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
   Set m_Layout = New Collection
   Set m_SubLotItems = New Collection
   Set m_ToPartTypes = New Collection
   Set m_ToParts = New Collection
   Set m_ToLocations = New Collection
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
   Set m_Layout = Nothing
   Set m_SubLotItems = Nothing
   Set m_ToPartTypes = Nothing
   Set m_ToParts = Nothing
   Set m_ToLocations = Nothing
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

Private Sub txtManualPrice_Change()
   m_HasModify = True
End Sub

Private Sub txtPackaging_Change()
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

Private Sub uctlToLocationLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlLocationLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlPartLookup_Change()
'Dim PartItemID As Long
'Dim LocationID As Long
'Dim PL As CPartLocation
'Dim iCount As Long
'
   m_HasModify = True
'   PartItemID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
'   LocationID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
'
'   If (PartItemID <= 0) Or (LocationID <= 0) Then
'      Exit Sub
'   End If
'
'   Set PL = New CPartLocation
'   PL.PART_LOCATION_ID = -1
'   PL.PART_ITEM_ID = PartItemID
'   PL.LOCATION_ID = LocationID
'   Call PL.QueryData(1, m_Rs, iCount)
'
'   If Not m_Rs.EOF Then
'      Call PL.PopulateFromRS(m_Rs)
'      txtPrice.Text = Format(PL.AVG_PRICE, "0.00")
'   Else
'      txtPrice.Text = Format(0, "0.00")
'   End If
'
'   Set PL = Nothing
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

Private Sub uctlToPartItemLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlToPartTypeLookup_Change()
Dim PartTypeID As Long
Dim Pt As CPartType

   PartTypeID = uctlToPartTypeLookup.MyCombo.ItemData(Minus2Zero(uctlToPartTypeLookup.MyCombo.ListIndex))

   If PartTypeID > 0 Then
      Set Pt = GetPartType(m_ToPartTypes, Trim(str(PartTypeID)))
      Call LoadPartItem(uctlToPartItemLookup.MyCombo, m_ToParts, PartTypeID, "N")
      Set uctlToPartItemLookup.MyCollection = m_ToParts
      
      Call LoadLocation(uctlToLocationLookup.MyCombo, m_Houses, 2, , , Pt.PART_GROUP_ID)
      Set uctlToLocationLookup.MyCollection = m_Houses
   End If
   
   m_HasModify = True

End Sub
