VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditPackProductItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7995
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditPackProductItem.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   7995
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   4215
      Left            =   0
      TabIndex        =   12
      Top             =   600
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   7435
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboSewingThread 
         Height          =   510
         Left            =   5880
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   1560
         Width           =   1455
      End
      Begin VB.ComboBox cboWeightPerPack 
         Height          =   510
         Left            =   1920
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   1560
         Width           =   1455
      End
      Begin prjFarmManagement.uctlTextBox txtTxAmount 
         Height          =   495
         Left            =   1920
         TabIndex        =   2
         Top             =   1080
         Width           =   1455
         _ExtentX        =   3413
         _ExtentY        =   873
      End
      Begin prjFarmManagement.uctlTextLookup uctlProductLookup 
         Height          =   435
         Left            =   1920
         TabIndex        =   1
         Top             =   600
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPartTypeLookup 
         Height          =   435
         Left            =   1920
         TabIndex        =   0
         Top             =   120
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPackAmount 
         Height          =   495
         Left            =   5880
         TabIndex        =   3
         Top             =   1080
         Width           =   1455
         _ExtentX        =   3413
         _ExtentY        =   873
      End
      Begin prjFarmManagement.uctlTextBox txtPalletLabelYellow 
         Height          =   495
         Left            =   1920
         TabIndex        =   6
         Top             =   2160
         Width           =   1455
         _ExtentX        =   3413
         _ExtentY        =   873
      End
      Begin prjFarmManagement.uctlTextBox txtPalletLabelGreen 
         Height          =   495
         Left            =   5880
         TabIndex        =   7
         Top             =   2160
         Width           =   1455
         _ExtentX        =   3413
         _ExtentY        =   873
      End
      Begin prjFarmManagement.uctlTextBox txtNote 
         Height          =   495
         Left            =   1920
         TabIndex        =   8
         Top             =   2760
         Width           =   5415
         _ExtentX        =   3413
         _ExtentY        =   873
      End
      Begin VB.Label lblNote 
         Alignment       =   1  'Right Justify
         Caption         =   "lblNote"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   2760
         Width           =   1665
      End
      Begin VB.Label lblPalletLabelGreen 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPalletLabelGreen"
         Height          =   375
         Left            =   4080
         TabIndex        =   20
         Top             =   2160
         Width           =   1665
      End
      Begin VB.Label lblPalletLabelYellow 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPalletLabelYellow"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   2160
         Width           =   1665
      End
      Begin VB.Label lblSewing_Thread 
         Alignment       =   1  'Right Justify
         Caption         =   "lblSewing_Thread"
         Height          =   375
         Left            =   4080
         TabIndex        =   18
         Top             =   1560
         Width           =   1665
      End
      Begin VB.Label lblPackAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPackAmount"
         Height          =   375
         Left            =   4080
         TabIndex        =   17
         Top             =   1080
         Width           =   1665
      End
      Begin VB.Label lblWeightPerPack 
         Alignment       =   1  'Right Justify
         Caption         =   "lblWeightPerPack"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   1665
      End
      Begin VB.Label lblProduct 
         Alignment       =   1  'Right Justify
         Caption         =   "lblProduct"
         Height          =   315
         Left            =   360
         TabIndex        =   15
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblType 
         Alignment       =   1  'Right Justify
         Caption         =   "lblType"
         Height          =   315
         Left            =   360
         TabIndex        =   14
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label lblTxAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblTxAmount"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1665
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2400
         TabIndex        =   9
         Top             =   3360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPackProductItem.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4080
         TabIndex        =   10
         Top             =   3360
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditPackProductItem"
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
Private m_Input_combo As Collection
Private m_Input1_combo As Collection
Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public TempCollection As Collection
Public COMMIT_FLAG As String
Public FromDate As Date
Public ToDate As Date

Public PartType As Long
Private m_PartTypes As Collection
Private m_PartItems As Collection

Private m_TempColl As Collection

Private Sub cboParcelType_Click()
   m_HasModify = True
End Sub

Private Sub cboSewingThread_Change()
   m_HasModify = True
End Sub

Private Sub cboSewingThread_Click()
m_HasModify = True
End Sub

Private Sub cboWeightPerPack_Change()
   m_HasModify = True
End Sub

Private Sub cboWeightPerPack_Click()
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
   
   Call InitNormalLabel(lblType, MapText("ประเภทวัตถุดิบ"))
   Call InitNormalLabel(lblProduct, MapText("วัตถุดิบ"))
   Call InitNormalLabel(lblTxAmount, MapText("จำนวน"))
   Call InitNormalLabel(lblPackAmount, MapText("จำนวนถุง"))
   Call InitNormalLabel(lblWeightPerPack, MapText("ขนาดบรรจุ"))
   Call InitNormalLabel(lblSewing_Thread, MapText("ด้ายที่เย็บ"))
   Call InitNormalLabel(lblPalletLabelYellow, MapText("ป้ายชี้บ่งสีเหลือง"))
   Call InitNormalLabel(lblPalletLabelGreen, MapText("ป้ายชี้บ่งสีเขียว"))
   Call InitNormalLabel(lblNote, MapText("หมายเหตุ"))
   
   Call txtNote.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtPalletLabelYellow.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtPalletLabelYellow.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtTxAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtPackAmount.SetTextLenType(TEXT_INTEGER, glbSetting.MONEY_TYPE)
   
   
   Call InitCombo(cboWeightPerPack)
   Call InitCombo(cboSewingThread)
      
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)

   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
End Sub

Private Function GetCostItem(TempCol As Collection, Ind As Long) As CCostItem
Dim Ci As CCostItem

   For Each Ci In TempCol
      If Ci.PARAM_PROCESS_ID = Ind Then
         Set GetCostItem = Ci
         Exit Function
      End If
   Next Ci
   
   Set GetCostItem = Nothing
End Function

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long

   If Flag Then
      Call EnableForm(Me, False)
         Dim Ma As CPackProductionItem
         Set Ma = TempCollection.Item(ID)
         
      If ShowMode = SHOW_EDIT Then
         uctlProductLookup.MyCombo.ListIndex = IDToListIndex(uctlProductLookup.MyCombo, Ma.PART_ITEM_ID)
         txtTxAmount.Text = Ma.TX_AMOUNT
         txtPackAmount.Text = Ma.PACK_AMOUNT
         cboWeightPerPack.ListIndex = Ma.WEIGHT_PER_PACK
         cboSewingThread.ListIndex = Ma.SEWING_THREAD
         txtPalletLabelYellow.Text = Ma.PALLET_LABEL_YELLOW
         txtPalletLabelGreen.Text = Ma.PALLET_LABEL_GREEN
         txtNote.Text = Ma.NOTE
      End If
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long
         
'   If Not VerifyCombo(lblExpenseType, uctlExpenseTypeLookup.MyCombo, False) Then
'      Exit Function
'   End If
'   If Not VerifyCombo(lblParcelType, cboParcelType, False) Then
'      Exit Function
'   End If
'   If Not VerifyCombo(lblRatioType, cboRatioType, False) Then
'      Exit Function
'   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim Ma As CPackProductionItem
   If ShowMode = SHOW_ADD Then
      Set Ma = New CPackProductionItem
   Else
      Set Ma = TempCollection.Item(ID)
   End If
   
   If ShowMode = SHOW_ADD Then
      Ma.Flag = "A"
      Call TempCollection.add(Ma)
   Else
      If Ma.Flag <> "A" Then
         Ma.Flag = "E"
      End If
   End If
   Ma.PART_ITEM_ID = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
   Ma.TX_AMOUNT = Val(txtTxAmount.Text)
   Ma.WEIGHT_PER_PACK = cboWeightPerPack.ListIndex
   Ma.PACK_AMOUNT = Val(txtPackAmount.Text)
   Ma.PALLET_LABEL_YELLOW = txtPalletLabelYellow.Text
   Ma.PALLET_LABEL_GREEN = txtPalletLabelGreen.Text
   Ma.SEWING_THREAD = cboSewingThread.ListIndex
   Ma.NOTE = txtNote.Text

   SaveData = True
End Function



Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadPartType(uctlPartTypeLookup.MyCombo, m_PartTypes)
      Set uctlPartTypeLookup.MyCollection = m_PartTypes
      
      Call initWeightPerPack(cboWeightPerPack)
      Call initSewingThread(cboSewingThread)

      uctlPartTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPartTypeLookup.MyCombo, PartType)
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
         Call QueryData(False)
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
   Set m_Input_combo = New Collection
   Set m_Input1_combo = New Collection
   Set m_Rs = New ADODB.Recordset
   
   Set m_PartTypes = New Collection
   Set m_PartItems = New Collection

   Set m_TempColl = New Collection
'   Set m_ExpenseType = New Collection
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_PartTypes = Nothing
   Set m_PartItems = Nothing
   
   Set m_TempColl = Nothing
'   Set m_ExpenseType = Nothing
End Sub

Private Sub txtAvgPrice_Change()
   m_HasModify = True
End Sub

Private Sub txtGroupNo_Change()
   m_HasModify = True
End Sub

Private Sub txtStdAmount_Change()
   m_HasModify = True
End Sub

Private Sub uctlFromFormulaLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlMixDate_HasChange()
   m_HasModify = True
End Sub

Private Sub txtAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtLink_Change()
   m_HasModify = True
End Sub

Private Sub txtCostAmount_Change(Index As Integer)
   m_HasModify = True
End Sub

Private Sub txtRef_Change()
   m_HasModify = True
End Sub

Private Sub txtSerialNo_Change()
   m_HasModify = True
End Sub

Private Sub uctlPlaceLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox1_Change()

End Sub

Private Sub uctlTextLookup1_Change()

End Sub

Private Sub txtExpenseAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtNote_Change()
   m_HasModify = True
End Sub

Private Sub txtPackAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtPalletLabelGreen_Change()
   m_HasModify = True
End Sub

Private Sub txtPalletLabelYellow_Change()
   m_HasModify = True
End Sub

Private Sub txtTxAmount_Change()
   m_HasModify = True
End Sub

'Private Sub uctlExpenseTypeLookup_Change()
'Dim ExpenseDetail As CExpenseDetail
'Dim ID As Long
'   ID = uctlExpenseTypeLookup.MyCombo.ItemData(Minus2Zero(uctlExpenseTypeLookup.MyCombo.ListIndex))
'
'   If ID > 0 Then
'      Set ExpenseDetail = GetObject("CExpenseDetail", m_ExpenseType, Trim(str(ID)))
'      txtExpenseAmount.Text = ExpenseDetail.GetFieldValue("EXPENSE_DETAIL_PRICE")
'
'      Set ExpenseDetail = Nothing
'   End If
'   m_HasModify = True
'End Sub
Private Sub uctlPartTypeLookup_Change()
Dim PartTypeID As Long
Dim Pt As CPartType

   PartTypeID = uctlPartTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPartTypeLookup.MyCombo.ListIndex))
   
   If PartTypeID > 0 Then
      Set Pt = GetPartType(m_PartTypes, Trim(str(PartTypeID)))
      Call LoadPartItem(uctlProductLookup.MyCombo, m_PartItems, PartTypeID, "N")
      Set uctlProductLookup.MyCollection = m_PartItems
   
'      Call LoadLocation(uctlPlaceLookup.MyCombo, m_Locations, 2, , , Pt.PART_GROUP_ID)
'      Set uctlPlaceLookup.MyCollection = m_Locations
   End If
   
   m_HasModify = True
End Sub

Private Sub uctlProductLookup_Change()
   m_HasModify = True
End Sub
