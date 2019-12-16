VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditBLImportItemEx2 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9075
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditBLImportItemEx2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   9075
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   6765
      Left            =   0
      TabIndex        =   16
      Top             =   600
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   11933
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboProjectName 
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   9
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   4830
         Width           =   3615
      End
      Begin VB.ComboBox cboExpenseType 
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   9
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   4410
         Width           =   3615
      End
      Begin VB.ComboBox cboDepartment 
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   9
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   3480
         Width           =   2955
      End
      Begin prjFarmManagement.uctlTextLookup uctlPartTypeLookup 
         Height          =   435
         Left            =   2400
         TabIndex        =   0
         Top             =   300
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPrice 
         Height          =   435
         Left            =   2400
         TabIndex        =   3
         Top             =   1650
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtQuantity 
         Height          =   435
         Left            =   2400
         TabIndex        =   2
         Top             =   1200
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPartLookup 
         Height          =   435
         Left            =   2400
         TabIndex        =   1
         Top             =   750
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlLocationLookup 
         Height          =   435
         Left            =   2415
         TabIndex        =   7
         Top             =   3030
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTotalPrice 
         Height          =   435
         Left            =   2400
         TabIndex        =   4
         Top             =   2100
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtDiscount 
         Height          =   435
         Left            =   6420
         TabIndex        =   5
         Top             =   2100
         Width           =   1365
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtNetTotalPrice 
         Height          =   435
         Left            =   2400
         TabIndex        =   6
         Top             =   2550
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtDesc 
         Height          =   435
         Left            =   2400
         TabIndex        =   9
         Top             =   3930
         Width           =   5415
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtNote 
         Height          =   435
         Left            =   2400
         TabIndex        =   12
         Top             =   5280
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin VB.Label lblNote 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   840
         TabIndex        =   34
         Top             =   5280
         Width           =   1485
      End
      Begin VB.Label lblProjectName 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   840
         TabIndex        =   33
         Top             =   4800
         Width           =   1485
      End
      Begin VB.Label lblExpenseType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   660
         TabIndex        =   32
         Top             =   4470
         Width           =   1605
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   810
         TabIndex        =   31
         Top             =   3960
         Width           =   1485
      End
      Begin VB.Label lblDepartment 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   660
         TabIndex        =   30
         Top             =   3540
         Width           =   1605
      End
      Begin VB.Label Label6 
         Height          =   375
         Left            =   4470
         TabIndex        =   29
         Top             =   2550
         Width           =   495
      End
      Begin VB.Label lblNetTotalPrice 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   810
         TabIndex        =   28
         Top             =   2520
         Width           =   1485
      End
      Begin VB.Label Label4 
         Height          =   375
         Left            =   7890
         TabIndex        =   27
         Top             =   2160
         Width           =   705
      End
      Begin VB.Label lblDiscount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   5130
         TabIndex        =   26
         Top             =   2160
         Width           =   1185
      End
      Begin VB.Label lblUnit 
         Height          =   375
         Left            =   4440
         TabIndex        =   25
         Top             =   1320
         Width           =   1005
      End
      Begin VB.Label lblTotalPrice 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   810
         TabIndex        =   24
         Top             =   2160
         Width           =   1485
      End
      Begin VB.Label Label2 
         Height          =   375
         Left            =   5400
         TabIndex        =   23
         Top             =   2190
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label1 
         Height          =   375
         Left            =   4470
         TabIndex        =   22
         Top             =   2100
         Width           =   495
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3240
         TabIndex        =   13
         Top             =   6000
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditBLImportItemEx2.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4920
         TabIndex        =   14
         Top             =   6000
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblPrice 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   795
         TabIndex        =   21
         Top             =   1710
         Width           =   1485
      End
      Begin VB.Label lblPartType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   75
         TabIndex        =   20
         Top             =   330
         Width           =   2205
      End
      Begin VB.Label lblPart 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   19
         Top             =   810
         Width           =   2085
      End
      Begin VB.Label lblQuantity 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   840
         TabIndex        =   18
         Top             =   1260
         Width           =   1485
      End
      Begin VB.Label lblLocation 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   795
         TabIndex        =   17
         Top             =   3060
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmAddEditBLImportItemEx2"
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
Public SupplierID As Long

Private m_PartTypes As Collection
Private m_Parts As Collection
Private m_Locations As Collection
Private m_Packagings As Collection
Private m_PartItemSpecs As Collection
Private m_PurchaseExpenses As Collection

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

Private Sub cboDepartment_Click()
   m_HasModify = True
End Sub

Private Sub cboExpenseType_Click()
   m_HasModify = True
End Sub

Private Sub cboProjectName_Click()
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
      
   Call InitNormalLabel(lblPartType, MapText("��������ʴ��ػ�ó�"))
   Call InitNormalLabel(lblPart, MapText("��ʴ��ػ�ó�"))
   Call InitNormalLabel(lblQuantity, MapText("����ҳ�����"))
   Call InitNormalLabel(lblPrice, MapText("�Ҥ�/˹���"))
   Call InitNormalLabel(lblTotalPrice, MapText("�Ҥ����"))
   Call InitNormalLabel(lblLocation, MapText("ʶҹ���Ѵ��"))
   Call InitNormalLabel(Label1, MapText("�ҷ"))
   Call InitNormalLabel(Label2, MapText("�ҷ"))
   Call InitNormalLabel(Label4, MapText("�ҷ"))
   Call InitNormalLabel(Label6, MapText("�ҷ"))
   Call InitNormalLabel(lblUnit, MapText(""))
   Call InitNormalLabel(lblDiscount, MapText("��ǹŴ"))
   Call InitNormalLabel(lblNetTotalPrice, MapText("��Ť���ط��"))
   Call InitNormalLabel(lblDepartment, MapText("˹��§ҹ/Ἱ�"))
   Call InitNormalLabel(lblDesc, MapText("��������´"))
   Call InitNormalLabel(lblExpenseType, MapText("�������¡���ԡ"))
   Call InitNormalLabel(lblProjectName, MapText("�����ç���"))
   Call InitNormalLabel(lblNote, MapText("�����˵�"))
   
   Call txtQuantity.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtPrice.Enabled = True
   Call txtTotalPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtNetTotalPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtNetTotalPrice.Enabled = False
   Call txtDiscount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtDesc.SetTextLenType(TEXT_STRING, glbSetting.NOTE_LEN)
   
   Call InitCombo(cboDepartment)
   Call InitCombo(cboExpenseType)
   Call InitCombo(cboProjectName)
   
   
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
         Dim EnpAddr As CSupItem

         Set EnpAddr = TempCollection.Item(ID)
         
         uctlPartTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPartTypeLookup.MyCombo, EnpAddr.PART_TYPE)
         uctlPartLookup.MyCombo.ListIndex = IDToListIndex(uctlPartLookup.MyCombo, EnpAddr.PART_ITEM_ID)
         uctlLocationLookup.MyCombo.ListIndex = IDToListIndex(uctlLocationLookup.MyCombo, EnpAddr.LOCATION_ID)

         txtQuantity.Text = EnpAddr.TX_AMOUNT
         txtPrice.Text = EnpAddr.RAW_PRICE
         txtTotalPrice.Text = EnpAddr.RAW_TOT_PRICE
         txtDiscount.Text = EnpAddr.DISCOUNT_AMT
         txtNetTotalPrice.Text = EnpAddr.TOTAL_ACTUAL_PRICE
         txtDesc.Text = EnpAddr.ITEM_DESC
         cboDepartment.ListIndex = IDToListIndex(cboDepartment, EnpAddr.TO_DEPARTMENT)
         cboExpenseType.ListIndex = IDToListIndex(cboExpenseType, EnpAddr.EXPENSE_TYPE)
         cboProjectName.ListIndex = IDToListIndex(cboProjectName, EnpAddr.PROJECT_NAME_ID)
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
      
      EnpAddress.CALCULATE_FLAG = "N"
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
   EnpAddress.TX_AMOUNT = txtQuantity.Text
   EnpAddress.ACTUAL_UNIT_PRICE = MyDiffEx(Val(txtNetTotalPrice.Text), Val(txtQuantity.Text))
   EnpAddress.INCLUDE_UNIT_PRICE = MyDiffEx(Val(txtNetTotalPrice.Text), Val(txtQuantity.Text))
   EnpAddress.TOTAL_ACTUAL_PRICE = Val(txtNetTotalPrice.Text)
   EnpAddress.RAW_PRICE = Val(txtPrice.Text)
   EnpAddress.RAW_TOT_PRICE = Val(txtTotalPrice.Text)
   EnpAddress.DISCOUNT_AMT = Val(txtDiscount.Text)
   EnpAddress.PART_TYPE_NAME = uctlPartLookup.MyCombo.Text
   EnpAddress.PART_NOTE = txtNote.Text
   EnpAddress.LOCATION_NAME = uctlLocationLookup.MyCombo.Text
   EnpAddress.PART_NO = uctlPartLookup.MyTextBox.Text
   EnpAddress.PART_DESC = uctlPartLookup.MyCombo.Text
   EnpAddress.TX_TYPE = "I"
   EnpAddress.TO_DEPARTMENT = cboDepartment.ItemData(Minus2Zero(cboDepartment.ListIndex))
   EnpAddress.ITEM_DESC = txtDesc.Text
   Set Pi = GetPartItem(m_Parts, Trim(Str(EnpAddress.PART_ITEM_ID)))
   EnpAddress.PIG_FLAG = Pi.PIG_FLAG
   EnpAddress.EXPENSE_TYPE = cboExpenseType.ItemData(Minus2Zero(cboExpenseType.ListIndex))
    EnpAddress.PROJECT_NAME_ID = cboProjectName.ItemData(Minus2Zero(cboProjectName.ListIndex))
    
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
         
      Call LoadLayout(cboDepartment)
      Call LoadMaster(cboExpenseType, , EXPENSE_TYPE)
      Call LoadMaster(cboProjectName, , SET_PROJECT)
      
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
Private Sub txtDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtDiscount_Change()
   m_HasModify = True
   txtNetTotalPrice.Text = Val(txtTotalPrice.Text) - Val(txtDiscount.Text)
End Sub

Private Sub txtNetTotalPrice_Change()
   m_HasModify = True
   txtNetTotalPrice.Text = Val(txtTotalPrice.Text) - Val(txtDiscount.Text)
End Sub

Private Sub txtNote_Change()
m_HasModify = True
End Sub

'Private Sub txtPrice_Change()
'   m_HasModify = True
''   txtTotalPrice.Text = Val(txtPrice.Text) * Val(txtQuantity.Text)
''   txtNetTotalPrice.Text = Val(txtTotalPrice.Text) - Val(txtDiscount.Text)
'End Sub

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

'Private Sub txtTotalPrice_Change()
'   m_HasModify = True
'   txtPrice.Text = MyDiffEx(Val(txtTotalPrice.Text), Val(txtQuantity.Text))
'   txtNetTotalPrice.Text = Val(txtTotalPrice.Text) - Val(txtDiscount.Text)
'End Sub

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
      Set Pi = GetPartItem(m_Parts, Trim(Str(PartItemID)))
      Call InitNormalLabel(lblUnit, Pi.UNIT_NAME)
   End If
   
   m_HasModify = True
End Sub

Private Sub uctlPartTypeLookup_Change()
Dim PartTypeID As Long
Dim Pt As CPartType

   PartTypeID = uctlPartTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPartTypeLookup.MyCombo.ListIndex))
   
   If PartTypeID > 0 Then
      Set Pt = GetPartType(m_PartTypes, Trim(Str(PartTypeID)))
      Call LoadPartItem(uctlPartLookup.MyCombo, m_Parts, PartTypeID, "N")
      Set uctlPartLookup.MyCollection = m_Parts
   
      Call LoadLocation(uctlLocationLookup.MyCombo, m_Locations, 2, , , Pt.PART_GROUP_ID)
      Set uctlLocationLookup.MyCollection = m_Locations
   End If
   
   m_HasModify = True
End Sub

Private Sub uctlTextBox1_Change()

End Sub
