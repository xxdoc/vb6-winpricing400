VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditPigAdjustItem1 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3780
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
   Icon            =   "frmAddEditPigAdjustItem1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   9435
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   3195
      Left            =   0
      TabIndex        =   7
      Top             =   600
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   5636
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextBox txtQuantity 
         Height          =   435
         Left            =   1785
         TabIndex        =   2
         Top             =   1230
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPartLookup 
         Height          =   435
         Left            =   1785
         TabIndex        =   1
         Top             =   780
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlLocationLookup 
         Height          =   435
         Left            =   1785
         TabIndex        =   3
         Top             =   1680
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPigTypeLookup 
         Height          =   435
         Left            =   1770
         TabIndex        =   0
         Top             =   330
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin VB.Label lblPigType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   180
         TabIndex        =   12
         Top             =   390
         Width           =   1485
      End
      Begin VB.Label Label1 
         Height          =   375
         Left            =   3870
         TabIndex        =   11
         Top             =   1650
         Width           =   1005
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3090
         TabIndex        =   4
         Top             =   2370
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPigAdjustItem1.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4740
         TabIndex        =   5
         Top             =   2370
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
         TabIndex        =   10
         Top             =   840
         Width           =   1485
      End
      Begin VB.Label lblQuantity 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   9
         Top             =   1290
         Width           =   1485
      End
      Begin VB.Label lblLocation 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   8
         Top             =   1740
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmAddEditPigAdjustItem1"
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
Public TxSeq As Long
Public COMMIT_FLAG As String

Private m_PartTypes As Collection
Private m_Pigs As Collection
Private m_PigTypes As Collection
Private m_Locations As Collection

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
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblPart, MapText("�ѻ�����Դ"))
   Call InitNormalLabel(lblQuantity, MapText("�ӹǹ"))
   Call InitNormalLabel(lblLocation, MapText("�ç���͹"))
   Call InitNormalLabel(Label1, MapText("�ҷ"))
   Call InitNormalLabel(lblPigType, MapText("�������ء�"))
   
   Call txtQuantity.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   
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
      
      If ShowMode = SHOW_EDIT Then
         Dim EnpAddr As CLotItem
         
         Set EnpAddr = TempCollection.Item(ID)
         
         uctlPigTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPigTypeLookup.MyCombo, PigCodeToID(EnpAddr.PIG_TYPE))
         uctlPartLookup.MyCombo.ListIndex = IDToListIndex(uctlPartLookup.MyCombo, EnpAddr.PART_ITEM_ID)
         uctlLocationLookup.MyCombo.ListIndex = IDToListIndex(uctlLocationLookup.MyCombo, EnpAddr.LOCATION_ID)
         
         txtQuantity.Text = EnpAddr.TX_AMOUNT
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

   If Not VerifyCombo(lblPart, uctlPartLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblQuantity, txtQuantity, False) Then
      Exit Function
   End If
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
      Call TempCollection.Add(EnpAddress)
   Else
      Set EnpAddress = TempCollection.Item(ID)
      If EnpAddress.Flag <> "A" Then
         EnpAddress.Flag = "E"
      End If
   End If

   EnpAddress.PART_ITEM_ID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
   EnpAddress.LOCATION_ID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
   EnpAddress.TX_AMOUNT = txtQuantity.Text
   EnpAddress.INCLUDE_UNIT_PRICE = EnpAddress.ACTUAL_UNIT_PRICE
   EnpAddress.PART_TYPE_NAME = uctlPartLookup.MyCombo.Text
   EnpAddress.LOCATION_NAME = uctlLocationLookup.MyCombo.Text
   EnpAddress.PART_NO = uctlPartLookup.MyTextBox.Text
   EnpAddress.PART_DESC = uctlPartLookup.MyCombo.Text
   EnpAddress.PIG_TYPE = PigTypeToCode(uctlPigTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPigTypeLookup.MyCombo.ListIndex)))
   EnpAddress.CALCULATE_FLAG = "N"
   
   Set EnpAddress = Nothing
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadPigType(uctlPigTypeLookup.MyCombo, m_PigTypes)
      Set uctlPigTypeLookup.MyCollection = m_PigTypes
      
      Call LoadLocation(uctlLocationLookup.MyCombo, m_Locations, 1, "")
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
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
   ElseIf Shift = 0 And KeyCode = 116 Then
'      Call cmdSearch_Click
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
   ElseIf Shift = 0 And KeyCode = 118 Then
'      Call cmdAdd_Click
'   ElseIf Shift = 0 And KeyCode = 117 Then
'      Call cmdDelete_Click
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
   ElseIf Shift = 0 And KeyCode = 114 Then
'      Call cmdEdit_Click
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
   End If
End Sub

Private Sub Form_Load()

   OKClick = False
   Call InitFormLayout
   
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_PartTypes = New Collection
   Set m_Pigs = New Collection
   Set m_Locations = New Collection
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
   Set m_Pigs = Nothing
   Set m_Locations = Nothing
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

Private Sub uctlLocationLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlPartLookup_Change()
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
