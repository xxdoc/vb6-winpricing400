VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditTransferItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5190
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
   Icon            =   "frmAddEditTransferItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
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
      Height          =   4605
      Left            =   0
      TabIndex        =   13
      Top             =   600
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   8123
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextLookup uctlPartTypeLookup 
         Height          =   435
         Left            =   1785
         TabIndex        =   0
         Top             =   270
         Width           =   5355
         _extentx        =   9446
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPrice 
         Height          =   435
         Left            =   1785
         TabIndex        =   5
         Top             =   2100
         Width           =   1995
         _extentx        =   3519
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtQuantity 
         Height          =   435
         Left            =   1785
         TabIndex        =   3
         Top             =   1650
         Width           =   1995
         _extentx        =   3519
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPartLookup 
         Height          =   435
         Left            =   1785
         TabIndex        =   1
         Top             =   720
         Width           =   5355
         _extentx        =   9446
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlLocationLookup 
         Height          =   435
         Left            =   1785
         TabIndex        =   2
         Top             =   1170
         Width           =   5355
         _extentx        =   9446
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlToLocationLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   7
         Top             =   2550
         Width           =   5355
         _extentx        =   9446
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlLayoutLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   8
         Top             =   3000
         Width           =   5355
         _extentx        =   9446
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPackaging 
         Height          =   435
         Left            =   5820
         TabIndex        =   6
         Top             =   2100
         Width           =   1365
         _extentx        =   2408
         _extenty        =   767
      End
      Begin VB.Label lblPackaging 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4530
         TabIndex        =   22
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
         MouseIcon       =   "frmAddEditTransferItem.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdLayout 
         Height          =   405
         Left            =   7170
         TabIndex        =   9
         Top             =   3000
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditTransferItem.frx":0BE4
         ButtonStyle     =   3
      End
      Begin VB.Label lblLayout 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   21
         Top             =   3060
         Width           =   1485
      End
      Begin VB.Label lblToLocation 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   20
         Top             =   2610
         Width           =   1485
      End
      Begin VB.Label Label1 
         Height          =   375
         Left            =   3840
         TabIndex        =   19
         Top             =   2130
         Width           =   1005
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3090
         TabIndex        =   10
         Top             =   3750
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditTransferItem.frx":0EFE
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4740
         TabIndex        =   11
         Top             =   3750
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
         Top             =   2160
         Width           =   1485
      End
      Begin VB.Label lblPartType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   17
         Top             =   300
         Width           =   1485
      End
      Begin VB.Label lblPart 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   16
         Top             =   780
         Width           =   1485
      End
      Begin VB.Label lblQuantity 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   15
         Top             =   1710
         Width           =   1485
      End
      Begin VB.Label lblLocation 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   180
         TabIndex        =   14
         Top             =   1200
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmAddEditTransferItem"
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
Public DocPartType As Long

Private m_PartTypes As Collection
Private m_Parts As Collection
Private m_Locations As Collection
Private m_Houses As Collection
Private m_Pigs As Collection
Private m_Layout As Collection
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
   Call InitNormalLabel(lblPrice, MapText("�Ҥ�"))
   Call InitNormalLabel(lblLocation, MapText("�ҡ��ѧ"))
   Call InitNormalLabel(lblToLocation, MapText("��Ҥ�ѧ"))
   Call InitNormalLabel(Label1, MapText("�ҷ"))
   Call InitNormalLabel(lblLayout, MapText("������ҵ�"))
   Call InitNormalLabel(lblPackaging, MapText("�ӹǹ�պ���"))
   
   Call txtQuantity.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtPrice.Enabled = False
   Call txtPackaging.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdLayout.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdLotSelect.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdLayout, MapText("..."))
   Call InitMainButton(cmdLotSelect, MapText("..."))
   
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Dim EnpAddr As CTransferItem

         Set EnpAddr = TempCollection.Item(id)
         Call CopySubLotItem(EnpAddr.ExportItem.SubLotItems, m_SubLotItems)

         uctlPartTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPartTypeLookup.MyCombo, EnpAddr.ExportItem.PART_TYPE)
         uctlPartLookup.MyCombo.ListIndex = IDToListIndex(uctlPartLookup.MyCombo, EnpAddr.ExportItem.PART_ITEM_ID)
         uctlLocationLookup.MyCombo.ListIndex = IDToListIndex(uctlLocationLookup.MyCombo, EnpAddr.ExportItem.LOCATION_ID)
         uctlToLocationLookup.MyCombo.ListIndex = IDToListIndex(uctlToLocationLookup.MyCombo, EnpAddr.ImportItem.LOCATION_ID)
         uctlLayoutLookup.MyCombo.ListIndex = IDToListIndex(uctlLayoutLookup.MyCombo, EnpAddr.ImportItem.LAYOUT_ID)

         txtQuantity.Text = EnpAddr.ExportItem.TX_AMOUNT
         txtPrice.Text = EnpAddr.ExportItem.INCLUDE_UNIT_PRICE
         txtPackaging.Text = EnpAddr.ExportItem.PACKAGING_AMT

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
      uctlLayoutLookup.MyCombo.ListIndex = IDToListIndex(uctlLayoutLookup.MyCombo, LayoutID)
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
   
   If uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex)) = _
       uctlToLocationLookup.MyCombo.ItemData(Minus2Zero(uctlToLocationLookup.MyCombo.ListIndex)) Then
         glbErrorLog.LocalErrorMsg = "�ç���͹��ҡѺ�ç���͹�͡�е�ͧᵡ��ҧ�ѹ"
         glbErrorLog.ShowUserError
         
         Exit Function
   End If
   
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
      Set EnpAddress = TempCollection.Item(id)
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
'   Call glbDaily.GenerateSubLot(EnpAddress.ExportItem, Nothing)

   EnpAddress.ImportItem.PART_TYPE = uctlPartTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPartTypeLookup.MyCombo.ListIndex))
   EnpAddress.ImportItem.PART_ITEM_ID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
   EnpAddress.ImportItem.LOCATION_ID = uctlToLocationLookup.MyCombo.ItemData(Minus2Zero(uctlToLocationLookup.MyCombo.ListIndex))
   EnpAddress.ImportItem.TX_AMOUNT = txtQuantity.Text
   EnpAddress.ImportItem.ACTUAL_UNIT_PRICE = Val(txtPrice.Text)
   EnpAddress.ImportItem.TOTAL_ACTUAL_PRICE = (txtQuantity.Text) * Val(txtPrice.Text)
   EnpAddress.ImportItem.INCLUDE_UNIT_PRICE = Val(txtPrice.Text)
   EnpAddress.ImportItem.TOTAL_INCLUDE_PRICE = EnpAddress.ImportItem.TOTAL_ACTUAL_PRICE
   EnpAddress.ImportItem.PART_TYPE_NAME = uctlPartLookup.MyCombo.Text
   EnpAddress.ImportItem.LOCATION_NAME = uctlToLocationLookup.MyCombo.Text
   EnpAddress.ImportItem.PART_NO = uctlPartLookup.MyTextBox.Text
   EnpAddress.ImportItem.PART_DESC = uctlPartLookup.MyCombo.Text
   EnpAddress.ImportItem.TX_TYPE = "I"
   EnpAddress.ImportItem.LAYOUT_ID = uctlLayoutLookup.MyCombo.ItemData(Minus2Zero(uctlLayoutLookup.MyCombo.ListIndex))
   EnpAddress.ImportItem.PACKAGING_AMT = Val(txtPackaging.Text)

'   If Not m_ManualFlag Then
'      EnpAddress.ExportItem.LOT_MANUAL = "N"
'      Call glbDaily.GenerateSubLot(EnpAddress.ExportItem, m_SubLotItems)
'   Else
'      EnpAddress.ExportItem.LOT_MANUAL = "Y"
'      Call glbDaily.GenerateSubLot(EnpAddress.ExportItem, m_SubLotItems)
'   End If

'   If Not glbDaily.VerifySubLot(EnpAddress.ExportItem) Then
'      EnpAddress.Flag = "D"
'      SaveData = False
'      Exit Function
'   End If

   Set EnpAddress = Nothing
   SaveData = True
End Function

Private Sub cmdWhExport_Click()
Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
Dim OKClick As Boolean
Dim LotAmount As Long
Dim I As Long
Dim t_Liw As CLotItemWH


   If Not VerifyCombo(lblPartType, uctlPartTypeLookup.MyCombo, False) Then
      Exit Sub
   End If
   If Not VerifyCombo(lblPart, uctlPartLookup.MyCombo, False) Then
      Exit Sub
   End If

   If Not VerifyCombo(lblLocation, uctlLocationLookup.MyCombo, False) Then
      Exit Sub
   End If

' If Not m_InventoryWhDocInput Is Nothing Then
'   frmAddEditTrnGoods.ShowMode = m_InventoryWhDocInput.AddEditMode
' End If


Set oMenu = New cPopupMenu
lMenuChosen = oMenu.Popup("����", "-", "���")
If lMenuChosen = 0 Then
   Exit Sub
End If

If lMenuChosen = 0 Then
   Exit Sub
End If

Dim IWD As CInventoryWHDoc
Set IWD = TempCollection2.Item(id)

   Dim LIW As CLotItemWH
   Set LIW = New CLotItemWH
   LIW.PART_ITEM_ID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
   LIW.LOCATION_ID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
   LIW.PART_NO = uctlPartLookup.MyTextBox.Text
   LIW.PART_DESC = uctlPartLookup.MyCombo.Text
   LIW.WEIGHT_PER_PACK = 30 ' WeightPerPack
'
If lMenuChosen = 1 Then
'   frmAddEditTrnGoods.ShowMode = SHOW_ADD
'
'   If CountItem(m_InventoryWhDocInput.C_LotItemsWH) = 0 Then '������� lot item ����
'        Set m_InventoryWhDocInput.C_LotItemsWH = Nothing
'      Set m_InventoryWhDocInput.C_LotItemsWH = New Collection
'      liw.AddEditMode = SHOW_ADD
'      liw.Flag = "A"
'      Call m_InventoryWhDocInput.C_LotItemsWH.add(liw, str(liw.PART_ITEM_ID))
'   Else
'      For Each t_Liw In m_InventoryWhDocInput.C_LotItemsWH
'         If liw.PART_ITEM_ID <> t_Liw.PART_ITEM_ID Then
'                  Set m_InventoryWhDocInput.C_LotItemsWH = Nothing
'                  Set m_InventoryWhDocInput.C_LotItemsWH = New Collection
'                  liw.AddEditMode = SHOW_ADD
'                  liw.Flag = "A"
'                  Call m_InventoryWhDocInput.C_LotItemsWH.add(liw, str(liw.PART_ITEM_ID))
'           Else
'                  If t_Liw.Flag <> "A" Then
''                     frmAddEditTrnGoods.ShowMode = SHOW_EDIT
'                     t_Liw.AddEditMode = SHOW_EDIT
'                     t_Liw.Flag = "E"
'                  End If
'                  frmAddEditTrnGoods.ShowMode = SHOW_EDIT
'         End If
'      Next t_Liw
'   End If
'
'   frmAddEditTrnGoods.DOCUMENT_TYPE = DOCUMENT_TYPE '17,19
'   frmAddEditTrnGoods.ProcessID = ProcessID
'
'   frmAddEditTrnGoods.id = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
'    Set frmAddEditTrnGoods.m_InventoryWHDoc = m_InventoryWhDocInput
'   Set frmAddEditTrnGoods.ParentForm = Me
'   frmAddEditTrnGoods.DOCUMENT_DATE = DOCUMENT_DATE
'   Load frmAddEditTrnGoods
'   frmAddEditTrnGoods.Show 1
'
'   OKClick = frmAddEditTrnGoods.OKClick
'
'   Unload frmAddEditTrnGoods
'   Set frmAddEditTrnGoods = Nothing
'
'   If OKClick Then
'      m_HasModify = True
'   End If
ElseIf lMenuChosen = 3 Then
'   If m_InventoryWhDocInput.C_LotItemsWH.Count = 0 Then
'      glbErrorLog.LocalErrorMsg = "�ѧ����բ�����"
'      glbErrorLog.ShowUserError
'      Exit Sub
'   End If

'   For Each t_Liw In m_InventoryWhDocInput.C_LotItemsWH
'            If t_Liw.Flag <> "A" Then
'               t_Liw.Flag = "E"
'            End If
'   Next t_Liw

'   frmAddEditTrnGoods.DOCUMENT_TYPE = DOCUMENT_TYPE '17,19
'   frmAddEditTrnGoods.ProcessID = ProcessID
'   frmAddEditTrnGoods.ShowMode = SHOW_EDIT
'   frmAddEditTrnGoods.id = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
'   frmAddEditTrnGoods.DOCUMENT_DATE = DOCUMENT_DATE
'    Set frmAddEditTrnGoods.m_InventoryWHDoc = m_InventoryWhDocInput
'   Set frmAddEditTrnGoods.ParentForm = Me
'   Load frmAddEditTrnGoods
'   frmAddEditTrnGoods.Show 1
'
'   OKClick = frmAddEditTrnGoods.OKClick
'
'   Unload frmAddEditTrnGoods
'   Set frmAddEditTrnGoods = Nothing
'
'   If OKClick Then
'      m_HasModify = True
'   End If
'ElseIf lMenuChosen = 5 Then
'   Dim IWHD As CInventoryWHDoc
'   Dim LIW2 As CLotItemWH
'   Dim LTD As CLotDoc
'
'   If m_InventoryWhDocInput.C_LotItemsWH.Count = 0 Then
'      glbErrorLog.LocalErrorMsg = "�ѧ����բ�����"
'      glbErrorLog.ShowUserError
'      Exit Sub
'   End If
'
'   For Each LIW2 In m_InventoryWhDocInput.C_LotItemsWH
'      For Each LTD In LIW2.C_LotDoc
'         LTD.Flag = "D" '���ź�֧ lotdoc ���е�ͧ����������ʹ update ����
'      Next LTD
''      LIW2.Flag = "D"
'   Next LIW2
'
'    txtPackAmount.Text = ""
'    txtAmount.Text = "0"
'    txtStdAmount.Text = "0"
End If
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadLayout(uctlLayoutLookup.MyCombo, m_Layout)
      Set uctlLayoutLookup.MyCollection = m_Layout
      
      Call LoadPartType(uctlPartTypeLookup.MyCombo, m_PartTypes)
      Set uctlPartTypeLookup.MyCollection = m_PartTypes
      
      Call LoadLocation(uctlLocationLookup.MyCombo, m_Locations, 2)
      Set uctlLocationLookup.MyCollection = m_Locations
      Call LoadLocation(uctlToLocationLookup.MyCombo, m_Houses, 2)
      Set uctlToLocationLookup.MyCollection = m_Houses

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
Dim PartItemID As Long
Dim LocationID As Long
Dim PL As CPartLocation
Dim iCount As Long

   m_HasModify = True
   PartItemID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
   LocationID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
   
   If (PartItemID <= 0) Or (LocationID <= 0) Then
      Exit Sub
   End If
   
   Set PL = New CPartLocation
   PL.PART_LOCATION_ID = -1
   PL.PART_ITEM_ID = PartItemID
   PL.LOCATION_ID = LocationID
   Call PL.QueryData(1, m_Rs, iCount)
   
   If Not m_Rs.EOF Then
      Call PL.PopulateFromRS(m_Rs)
      txtPrice.Text = Format(PL.AVG_PRICE, "0.00")
   Else
      txtPrice.Text = Format(0, "0.00")
   End If
   
   Set PL = Nothing
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
      Call LoadLocation(uctlToLocationLookup.MyCombo, m_Houses, 2, , , Pt.PART_GROUP_ID)
      Set uctlToLocationLookup.MyCollection = m_Houses
   End If
   
   m_HasModify = True
End Sub

Private Sub uctlPigWeekLookup_Change()
   m_HasModify = True
End Sub
