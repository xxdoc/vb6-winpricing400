VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditPaymentItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4605
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
   Icon            =   "frmAddEditPaymentItem.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
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
      Height          =   4035
      Left            =   0
      TabIndex        =   10
      Top             =   600
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   7117
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboPaymentType 
         Height          =   510
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   300
         Width           =   2985
      End
      Begin prjFarmManagement.uctlTextBox txtQuantity 
         Height          =   435
         Left            =   1785
         TabIndex        =   3
         Top             =   1620
         Width           =   3645
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlBankLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   1
         Top             =   750
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtWeight 
         Height          =   435
         Left            =   1770
         TabIndex        =   4
         Top             =   2070
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlBankBranchLookup 
         Height          =   435
         Left            =   1785
         TabIndex        =   2
         Top             =   1200
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtFeeAmount 
         Height          =   435
         Left            =   1770
         TabIndex        =   5
         Top             =   2520
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   767
      End
      Begin Threed.SSCheck chkPostFlag 
         Height          =   375
         Left            =   5640
         TabIndex        =   6
         Top             =   2550
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   661
         _Version        =   131073
         TripleState     =   -1  'True
      End
      Begin VB.Label lblFeeAmount 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   120
         TabIndex        =   18
         Top             =   2580
         Width           =   1515
      End
      Begin VB.Label Label1 
         Height          =   345
         Left            =   4065
         TabIndex        =   17
         Top             =   2580
         Width           =   1245
      End
      Begin VB.Label lblPaymentType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   60
         TabIndex        =   16
         Top             =   360
         Width           =   1665
      End
      Begin VB.Label lblAvgPrice 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   4065
         TabIndex        =   15
         Top             =   2130
         Width           =   1245
      End
      Begin VB.Label lblBankBranch 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   180
         TabIndex        =   14
         Top             =   1230
         Width           =   1485
      End
      Begin VB.Label lblWeight 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   540
         TabIndex        =   13
         Top             =   2130
         Width           =   1125
      End
      Begin VB.Label lblBank 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   12
         Top             =   810
         Width           =   1485
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3090
         TabIndex        =   7
         Top             =   3180
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPaymentItem.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4740
         TabIndex        =   8
         Top             =   3180
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblQuantity 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   11
         Top             =   1680
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmAddEditPaymentItem"
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
Private m_BankBranchs As Collection
Private m_Pigs As Collection
Private m_PigStatuss As Collection

Public Area As Long

Private Sub cboTextType_Click()
   m_HasModify = True
End Sub

Private Sub cboCountry_Click()
   m_HasModify = True
End Sub

Private Sub chkBangkok_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cboPaymentType_Click()
   m_HasModify = True
End Sub

Private Sub cboPaymentType_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub chkPostFlag_Click(Value As Integer)
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
   
   Call InitNormalLabel(lblQuantity, MapText("�Ţ�����"))
   Call InitNormalLabel(lblBank, MapText("��Ҥ��"))
   Call InitNormalLabel(lblWeight, MapText("�ӹǹ�Թ"))
   Call InitNormalLabel(lblBankBranch, MapText("�Ң�"))
   Call InitNormalLabel(lblAvgPrice, MapText(""))
   Call InitNormalLabel(lblPaymentType, MapText("��ê����Թ"))
   Call InitNormalLabel(lblFeeAmount, MapText("��Ҹ�������"))
   Call InitNormalLabel(Label1, MapText("�ҷ"))
   
   Call txtQuantity.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtWeight.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtFeeAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   
   Call InitCombo(cboPaymentType)
   Call InitCheckBox(chkPostFlag, "����Թ�����º����")
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Dim Di As CPaymentItem
         
         Set Di = TempCollection.Item(ID)
         
         cboPaymentType.ListIndex = IDToListIndex(cboPaymentType, Di.PAYMENT_TYPE)
         uctlBankLookup.MyCombo.ListIndex = IDToListIndex(uctlBankLookup.MyCombo, Di.BANK_ID)
         uctlBankBranchLookup.MyCombo.ListIndex = IDToListIndex(uctlBankBranchLookup.MyCombo, Di.BANK_BRANCH)
         txtQuantity.Text = Di.CHECK_NO
         txtWeight.Text = Di.PAY_AMOUNT
         txtFeeAmount.Text = Di.FEE_AMOUNT
         chkPostFlag.Value = FlagToCheck(Di.POST_FLAG)
         
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

   If Not VerifyCombo(lblPaymentType, cboPaymentType, False) Then
      Exit Function
   End If
      
   If Not VerifyTextControl(lblWeight, txtWeight, False) Then
      Exit Function
   End If
      
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim Di As CPaymentItem
   If ShowMode = SHOW_ADD Then
      Set Di = New CPaymentItem
      
      Di.Flag = "A"
      Call TempCollection.add(Di)
   Else
      Set Di = TempCollection.Item(ID)
      If Di.Flag <> "A" Then
         Di.Flag = "E"
      End If
   End If

   Di.PAYMENT_TYPE = cboPaymentType.ItemData(Minus2Zero(cboPaymentType.ListIndex))
   Di.CHECK_NO = txtQuantity.Text
   Di.BANK_ID = uctlBankLookup.MyCombo.ItemData(Minus2Zero(uctlBankLookup.MyCombo.ListIndex))
   Di.BANK_NAME = uctlBankLookup.MyCombo.Text
   If uctlBankBranchLookup.MyCombo.ListIndex > 0 Then
      Di.BANK_BRANCH = uctlBankBranchLookup.MyCombo.ItemData(Minus2Zero(uctlBankBranchLookup.MyCombo.ListIndex))
   Else
      Di.BANK_BRANCH = -1
   End If
   Di.BANK_BRANCH_NAME = uctlBankBranchLookup.MyCombo.Text
   Di.PAY_AMOUNT = Val(txtWeight.Text)
   Di.FEE_AMOUNT = Val(txtFeeAmount.Text)
   Di.POST_FLAG = Check2Flag(chkPostFlag.Value)
   
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call InitPaymentType(cboPaymentType)
      
      Call LoadBank(uctlBankLookup.MyCombo, m_PartTypes)
      Set uctlBankLookup.MyCollection = m_PartTypes

      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
         chkPostFlag.Value = FlagToCheck("Y")
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
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 118 Then
'      Call cmdAdd_Click
      KeyCode = 0
'   ElseIf Shift = 0 And KeyCode = 117 Then
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
   Set m_BankBranchs = New Collection
   Set m_Pigs = New Collection
   Set m_PigStatuss = New Collection
   Set m_PartTypes = New Collection
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
   Set m_BankBranchs = Nothing
   Set m_Pigs = Nothing
   Set m_PigStatuss = Nothing
   Set m_PartTypes = Nothing
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

Private Sub txtAvgPrice_Change()
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

Private Sub uctlPigStatusLookup_Change()
   m_HasModify = True
End Sub

Private Sub txtFeeAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtQuantity_Change()
   m_HasModify = True
End Sub

Private Sub txtWeight_Change()
   m_HasModify = True
End Sub

Private Sub uctlBankBranchLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlBankLookup_Change()
Dim BankID As Long

   BankID = uctlBankLookup.MyCombo.ItemData(Minus2Zero(uctlBankLookup.MyCombo.ListIndex))
   If BankID > 0 Then
      Call LoadBankBranch(uctlBankBranchLookup.MyCombo, m_BankBranchs, BankID)
      Set uctlBankBranchLookup.MyCollection = m_BankBranchs
   End If
   m_HasModify = True
End Sub

Private Sub uctlLocationLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlPartLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlPigWeekLookup_Change()
   m_HasModify = True
End Sub
