VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditCostExpense 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditCostExpense.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   3135
      Left            =   0
      TabIndex        =   8
      Top             =   600
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   5530
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboRatioType 
         Height          =   510
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1200
         Width           =   4920
      End
      Begin VB.ComboBox cboParcelType 
         Height          =   510
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   750
         Width           =   2670
      End
      Begin prjFarmManagement.uctlTextLookup uctlExpenseTypeLookup 
         Height          =   405
         Left            =   1860
         TabIndex        =   0
         Top             =   300
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtExpenseAmount 
         Height          =   495
         Left            =   1860
         TabIndex        =   4
         Top             =   1620
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
      End
      Begin Threed.SSCommand cmdSelect 
         Height          =   405
         Left            =   6780
         TabIndex        =   3
         Top             =   1260
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCostExpense.frx":08CA
         ButtonStyle     =   3
      End
      Begin VB.Label lblExpenseType 
         Alignment       =   1  'Right Justify
         Caption         =   "lblCountry1"
         Height          =   315
         Left            =   30
         TabIndex        =   12
         Top             =   330
         Width           =   1725
      End
      Begin VB.Label lblRatioType 
         Alignment       =   1  'Right Justify
         Caption         =   "lblCountry1"
         Height          =   315
         Left            =   30
         TabIndex        =   11
         Top             =   1260
         Width           =   1755
      End
      Begin VB.Label lblParcelType 
         Alignment       =   1  'Right Justify
         Caption         =   "lblCountry1"
         Height          =   315
         Left            =   30
         TabIndex        =   10
         Top             =   810
         Width           =   1725
      End
      Begin VB.Label lblExpenseAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblSerialNo"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   1665
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2370
         TabIndex        =   5
         Top             =   2310
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCostExpense.frx":0BE4
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4020
         TabIndex        =   6
         Top             =   2310
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditCostExpense"
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

Private m_ProcessParams As Collection
Private m_PartItems As Collection
Private m_Locations As Collection
Private m_Formulas As Collection
Private m_ExpenseType As Collection

Private m_TempColl As Collection

Private Sub cboParcelType_Click()
   m_HasModify = True
End Sub

Private Sub cboParcelType_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cboRatioType_Click()
Dim TempID As Long
   
   TempID = cboRatioType.ItemData(Minus2Zero(cboRatioType.ListIndex))
   If TempID <= 0 Then
      cmdSelect.Enabled = False
      Exit Sub
   End If
   
   cmdSelect.Enabled = (TempID = RATIO_RAW)
   If (TempID = RATIO_RAW) Or (TempID = RATIO_VARY) Then
      Call InitNormalLabel(lblExpenseAmount, "ต้นทุนต่อตัน")
   ElseIf (TempID = RATIO_PERCENT) Then
      Call InitNormalLabel(lblExpenseAmount, "เปอร์เซ็นต์")
   Else
      Call InitNormalLabel(lblExpenseAmount, "มูลค่าต้นทุน")
   End If
   
   m_HasModify = True
End Sub

Private Sub cboRatioType_KeyPress(KeyAscii As Integer)
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
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
      
   Call InitNormalLabel(lblExpenseType, MapText("ต้นทุนผลิต"))
   Call InitNormalLabel(lblParcelType, MapText("ปันให้กับ"))
   Call InitNormalLabel(lblRatioType, MapText("อัตราส่วนตาม"))
   Call InitNormalLabel(lblExpenseAmount, MapText("มูลค่าต้นทุน"))
   
   Call txtExpenseAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSelect.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitCombo(cboParcelType)
   Call InitCombo(cboRatioType)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdSelect, MapText("..."))
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
'      Call EnableForm(Me, False)
'         Dim Ma As CPackProductionItem
'         Set Ma = TempCollection.Item(ID)
'
'      If ShowMode = SHOW_EDIT Then
'         uctlExpenseTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlExpenseTypeLookup.MyCombo, Ma.PACK)
'         cboParcelType.ListIndex = IDToListIndex(cboParcelType, Ma.PACKAGE_TYPE)
'         cboRatioType.ListIndex = IDToListIndex(cboRatioType, Ma.RATIO_TYPE)
'         txtExpenseAmount.Text = Ma.EXPENSE_AMOUNT
'
'         Set m_TempColl = Ma.CostRaws
'      End If
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
         
   If Not VerifyCombo(lblExpenseType, uctlExpenseTypeLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblParcelType, cboParcelType, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblRatioType, cboRatioType, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim Ma As CCostExpense
   If ShowMode = SHOW_ADD Then
      Set Ma = New CCostExpense
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
   
   Ma.EXPENSE_AMOUNT = Val(txtExpenseAmount.Text)
   Ma.EXPENSE_TYPE = uctlExpenseTypeLookup.MyCombo.ItemData(Minus2Zero(uctlExpenseTypeLookup.MyCombo.ListIndex))
   Ma.RATIO_TYPE = cboRatioType.ItemData(Minus2Zero(cboRatioType.ListIndex))
   Ma.PACKAGE_TYPE = cboParcelType.ItemData(Minus2Zero(cboParcelType.ListIndex))
   Ma.PARAMETER_PROCESS_NAME = uctlExpenseTypeLookup.MyCombo.Text
   
   Set Ma.CostRaws = m_TempColl
   
   SaveData = True
End Function

Private Sub cmdSelect_Click()
   frmAddEditCostRaw.ShowMode = SHOW_ADD
   
   Set frmAddEditCostRaw.TempCollection = m_TempColl
   
   frmAddEditCostRaw.HeaderText = MapText("")
   Load frmAddEditCostRaw
   frmAddEditCostRaw.Show 1
      
   OKClick = frmAddEditCostRaw.OKClick
      
   Unload frmAddEditCostRaw
   Set frmAddEditCostRaw = Nothing
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadParameterProcess(uctlExpenseTypeLookup.MyCombo, m_ProcessParams)
      Set uctlExpenseTypeLookup.MyCollection = m_ProcessParams
      
      Call InitParcelType(cboParcelType)
      Call InitRatioType(cboRatioType)
      
      Call LoadSumExpenseDetail(Nothing, m_ExpenseType, FromDate, ToDate)
      
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
   
   Set m_ProcessParams = New Collection
   Set m_PartItems = New Collection
   Set m_Locations = New Collection
   Set m_Formulas = New Collection
   
   Set m_TempColl = New Collection
   Set m_ExpenseType = New Collection
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_ProcessParams = Nothing
   Set m_PartItems = Nothing
   Set m_Locations = Nothing
   Set m_Formulas = Nothing
   
   Set m_TempColl = Nothing
   Set m_ExpenseType = Nothing
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

Private Sub uctlPartTypeLookup_Change()

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

Private Sub uctlProductLookup_Change()

End Sub

Private Sub uctlTextBox1_Change()

End Sub

Private Sub uctlTextLookup1_Change()

End Sub

Private Sub txtExpenseAmount_Change()
   m_HasModify = True
End Sub
Private Sub uctlExpenseTypeLookup_Change()
Dim ExpenseDetail As CExpenseDetail
Dim ID As Long
   ID = uctlExpenseTypeLookup.MyCombo.ItemData(Minus2Zero(uctlExpenseTypeLookup.MyCombo.ListIndex))
   
   If ID > 0 Then
      Set ExpenseDetail = GetObject("CExpenseDetail", m_ExpenseType, Trim(str(ID)))
      txtExpenseAmount.Text = ExpenseDetail.GetFieldValue("EXPENSE_DETAIL_PRICE")
      
      Set ExpenseDetail = Nothing
   End If
   m_HasModify = True
End Sub
