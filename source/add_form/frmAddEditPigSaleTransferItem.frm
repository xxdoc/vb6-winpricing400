VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditPigSaleTransferItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4635
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
   Icon            =   "frmAddEditPigSaleTransferItem.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   9435
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   4065
      Left            =   0
      TabIndex        =   10
      Top             =   600
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   7170
      _Version        =   131073
      PictureBackgroundStyle=   2
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
         TabIndex        =   2
         Top             =   1200
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlLocationLookup 
         Height          =   435
         Left            =   1785
         TabIndex        =   0
         Top             =   300
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlToLocationLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   5
         Top             =   2100
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPigStatusLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   6
         Top             =   2550
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtWeight 
         Height          =   435
         Left            =   5190
         TabIndex        =   4
         Top             =   1650
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPigTypeLookup 
         Height          =   435
         Left            =   1785
         TabIndex        =   1
         Top             =   750
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin VB.Label lblPigType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   180
         TabIndex        =   18
         Top             =   780
         Width           =   1485
      End
      Begin VB.Label Label1 
         Height          =   345
         Left            =   7200
         TabIndex        =   17
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblWeight 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   3960
         TabIndex        =   16
         Top             =   1710
         Width           =   1125
      End
      Begin VB.Label lblPigStatus 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   15
         Top             =   2610
         Width           =   1485
      End
      Begin VB.Label lblToLocation 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   14
         Top             =   2160
         Width           =   1485
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3090
         TabIndex        =   7
         Top             =   3240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPigSaleTransferItem.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4740
         TabIndex        =   8
         Top             =   3240
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
         TabIndex        =   13
         Top             =   1260
         Width           =   1485
      End
      Begin VB.Label lblQuantity 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   12
         Top             =   1710
         Width           =   1485
      End
      Begin VB.Label lblLocation 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   180
         TabIndex        =   11
         Top             =   330
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmAddEditPigSaleTransferItem"
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
Private m_PigStatuss As Collection
Private m_PigTypes As Collection

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
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblPart, MapText("�ѻ�����Դ"))
   Call InitNormalLabel(lblQuantity, MapText("����ҳ"))
   Call InitNormalLabel(lblLocation, MapText("�ҡ�ç���͹"))
   Call InitNormalLabel(lblToLocation, MapText("����ç���͹"))
   Call InitNormalLabel(lblPigStatus, MapText("ʶҹ����"))
   Call InitNormalLabel(lblWeight, MapText("���˹ѡ"))
   Call InitNormalLabel(Label1, MapText("�.�."))
   Call InitNormalLabel(lblPigType, MapText("�������ء�"))
   
   Call txtQuantity.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtWeight.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
'      If ShowMode = SHOW_EDIT Then
'         Dim EnpAddr As CTransferItem
'
'         Set EnpAddr = TempCollection.Item(ID)
'
'         uctlPigTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPigTypeLookup.MyCombo, PigCodeToID(EnpAddr.ImportItem.PIG_TYPE))
'         uctlPartLookup.MyCombo.ListIndex = IDToListIndex(uctlPartLookup.MyCombo, EnpAddr.ExportItem.PART_ITEM_ID)
'         uctlLocationLookup.MyCombo.ListIndex = IDToListIndex(uctlLocationLookup.MyCombo, EnpAddr.ExportItem.LOCATION_ID)
'         uctlToLocationLookup.MyCombo.ListIndex = IDToListIndex(uctlToLocationLookup.MyCombo, EnpAddr.ImportItem.LOCATION_ID)
'         uctlPigStatusLookup.MyCombo.ListIndex = IDToListIndex(uctlPigStatusLookup.MyCombo, EnpAddr.ExportItem.PIG_STATUS)
'
'         txtQuantity.Text = EnpAddr.ExportItem.EXPORT_AMOUNT
'         txtWeight.Text = EnpAddr.ExportItem.TOTAL_WEIGHT
'         cmdOK.Enabled = (COMMIT_FLAG <> "Y")
'      End If
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

   If Not VerifyCombo(lblPart, uctlPartLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblQuantity, txtQuantity, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblWeight, txtWeight, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblLocation, uctlLocationLookup.MyCombo, False) Then
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
   
'   Dim EnpAddress As CTransferItem
'   Dim Ei As CExportItem
'   Dim Ii As CLotItem
'   If ShowMode = SHOW_ADD Then
'      Set Ei = New CExportItem
'      Set Ii = New CLotItem
'      Set EnpAddress = New CTransferItem
'
'      Ei.Flag = "A"
'      Ei.CALCULATE_FLAG = "N"
'      Ii.Flag = "A"
'      Ii.CALCULATE_FLAG = "N"
'      EnpAddress.Flag = "A"
'
'      Set EnpAddress.ExportItem = Ei
'      Set EnpAddress.ImportItem = Ii
'
'      Call TempCollection.Add(EnpAddress)
'   Else
'      Set EnpAddress = TempCollection.Item(ID)
'      If EnpAddress.Flag <> "A" Then
'         EnpAddress.Flag = "E"
'         EnpAddress.ExportItem.Flag = "E"
'         EnpAddress.ImportItem.Flag = "E"
'      End If
'   End If
'
'   EnpAddress.ExportItem.PART_ITEM_ID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
'   EnpAddress.ExportItem.LOCATION_ID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
'   EnpAddress.ExportItem.EXPORT_AMOUNT = txtQuantity.Text
'   EnpAddress.ExportItem.EXPORT_AVG_PRICE = 0
'   EnpAddress.ExportItem.PART_TYPE_NAME = uctlPartLookup.MyCombo.Text
'   EnpAddress.ExportItem.LOCATION_NAME = uctlLocationLookup.MyCombo.Text
'   EnpAddress.ExportItem.PART_NO = uctlPartLookup.MyTextBox.Text
'   EnpAddress.ExportItem.PART_DESC = uctlPartLookup.MyCombo.Text
'   EnpAddress.ExportItem.HOUSE_ID = -1
'   EnpAddress.ExportItem.PIG_TYPE = PigTypeToCode(uctlPigTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPigTypeLookup.MyCombo.ListIndex)))
'   EnpAddress.ExportItem.PIG_STATUS = uctlPigStatusLookup.MyCombo.ItemData(Minus2Zero(uctlPigStatusLookup.MyCombo.ListIndex))
'   EnpAddress.ExportItem.TOTAL_WEIGHT = Val(txtWeight.Text)
'
'   EnpAddress.ImportItem.PART_ITEM_ID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
'   EnpAddress.ImportItem.LOCATION_ID = uctlToLocationLookup.MyCombo.ItemData(Minus2Zero(uctlToLocationLookup.MyCombo.ListIndex))
'   EnpAddress.ImportItem.TX_AMOUNT = txtQuantity.Text
'   EnpAddress.ImportItem.ACTUAL_UNIT_PRICE = 0
'   EnpAddress.ImportItem.INCLUDE_UNIT_PRICE = 0
'   EnpAddress.ImportItem.PART_TYPE_NAME = uctlPartLookup.MyCombo.Text
'   EnpAddress.ImportItem.LOCATION_NAME = uctlToLocationLookup.MyCombo.Text
'   EnpAddress.ImportItem.PART_NO = uctlPartLookup.MyTextBox.Text
'   EnpAddress.ImportItem.PART_DESC = uctlPartLookup.MyCombo.Text
'   EnpAddress.ImportItem.PIG_TYPE = PigTypeToCode(uctlPigTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPigTypeLookup.MyCombo.ListIndex)))
'   EnpAddress.ImportItem.TOTAL_WEIGHT = Val(txtWeight.Text)
'
'   Set EnpAddress = Nothing
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadPigType(uctlPigTypeLookup.MyCombo, m_PigTypes)
      Set uctlPigTypeLookup.MyCollection = m_PigTypes
      
      Call LoadLocation(uctlLocationLookup.MyCombo, m_Locations, 1)
      Set uctlLocationLookup.MyCollection = m_Locations
      Call LoadLocation(uctlToLocationLookup.MyCombo, m_Houses, 1, "Y")
      Set uctlToLocationLookup.MyCollection = m_Houses

      Call LoadProductStatus(uctlPigStatusLookup.MyCombo, m_PigStatuss)
      Set uctlPigStatusLookup.MyCollection = m_PigStatuss

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
   Set m_Houses = New Collection
   Set m_Pigs = New Collection
   Set m_PigStatuss = New Collection
   Set m_PigTypes = New Collection
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
   Set m_PigTypes = Nothing
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

Private Sub txtWeight_Change()
   m_HasModify = True
End Sub

Private Sub uctlPigStatusLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlPigTypeLookup_Change()
Dim PigTypeCode As String

   m_HasModify = True

   PigTypeCode = PigTypeToCode(uctlPigTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPigTypeLookup.MyCombo.ListIndex)))
   If PigTypeCode <> "" Then
      Call LoadPartItem(uctlPartLookup.MyCombo, m_Pigs, -1, "Y", PigTypeCode)
      Set uctlPartLookup.MyCollection = m_Pigs
   End If
End Sub

Private Sub uctlToLocationLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlLocationLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlPartLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlPartTypeLookup_Change()
Dim PartTypeID As Long
   
   Call LoadPartItem(uctlPartLookup.MyCombo, m_Parts, PartTypeID, "")
   Set uctlPartLookup.MyCollection = m_Parts
   
   m_HasModify = True
End Sub

Private Sub uctlPigWeekLookup_Change()
   m_HasModify = True
End Sub
