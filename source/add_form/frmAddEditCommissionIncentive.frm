VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditCommissionIncentive 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7860
   Icon            =   "frmAddEditCommissionIncentive.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   7860
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   5925
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   10451
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboUnitType 
         Height          =   315
         Left            =   1860
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   3960
         Width           =   2895
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   7875
         _ExtentX        =   13891
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextLookup uctlStockCodeLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   2
         Top             =   2040
         Width           =   5385
         _extentx        =   9499
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlFreelanceLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   1080
         Width           =   5385
         _extentx        =   9499
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtAmount 
         Height          =   435
         Left            =   1860
         TabIndex        =   3
         Top             =   2520
         Width           =   2955
         _extentx        =   5212
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlCustomerLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   1
         Top             =   1560
         Width           =   5385
         _extentx        =   9499
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtFromAmount 
         Height          =   435
         Left            =   1860
         TabIndex        =   4
         Top             =   3000
         Width           =   2955
         _extentx        =   5212
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtToAmount 
         Height          =   435
         Left            =   1860
         TabIndex        =   5
         Top             =   3480
         Width           =   2955
         _extentx        =   5212
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtAmountOver 
         Height          =   435
         Left            =   5160
         TabIndex        =   8
         Top             =   4440
         Width           =   1275
         _extentx        =   5212
         _extenty        =   767
      End
      Begin VB.Label Label2 
         Caption         =   "Label1"
         Height          =   315
         Left            =   6480
         TabIndex        =   22
         Top             =   4560
         Width           =   735
      End
      Begin VB.Label lblAmountOver 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   255
         Left            =   3720
         TabIndex        =   21
         Top             =   4560
         Width           =   1335
      End
      Begin Threed.SSCheck sscAmountOver 
         Height          =   375
         Left            =   1860
         TabIndex        =   7
         Top             =   4500
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblUnitType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   20
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label lblToAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   19
         Top             =   3570
         Width           =   1575
      End
      Begin VB.Label lblFromAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   18
         Top             =   3120
         Width           =   1575
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   1440
         TabIndex        =   9
         Top             =   5160
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCommissionIncentive.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label lblCustomerLookup 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   17
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label lblAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   16
         Top             =   2610
         Width           =   1575
      End
      Begin VB.Label lblFreelance 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   15
         Top             =   1230
         Width           =   1575
      End
      Begin VB.Label lblStockCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   14
         Top             =   2160
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5040
         TabIndex        =   11
         Top             =   5160
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3240
         TabIndex        =   10
         Top             =   5160
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCommissionIncentive.frx":2ABC
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditCommissionIncentive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_CommissionIncentive As CCommissionIncentive

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Public DocumentType As Long

Private m_Freelances As Collection
Private m_StockCodes As Collection
Private m_Customers As Collection
Public ParentForm As Object

Private Sub cboUnitType_Change()
   m_HasModify = True
End Sub

Private Sub cboUnitType_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cmdNext_Click()
If Not SaveData Then
      Exit Sub
End If

uctlStockCodeLookup.MyCombo.ListIndex = IDToListIndex(uctlStockCodeLookup.MyCombo, -1)
Call ParentForm.RefreshGrid

   m_HasModify = False
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)

      m_CommissionIncentive.INCENTIVE_ID = ID
      m_CommissionIncentive.QueryFlag = 1
      If Not glbDaily.QueryCommissionIncentive(m_CommissionIncentive, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_CommissionIncentive.PopulateFromRS(1, m_Rs)
      uctlFreelanceLookup.MyCombo.ListIndex = IDToListIndex(uctlFreelanceLookup.MyCombo, m_CommissionIncentive.FREELANCE_ID)
      uctlStockCodeLookup.MyCombo.ListIndex = IDToListIndex(uctlStockCodeLookup.MyCombo, m_CommissionIncentive.PART_ITEM_ID)
      uctlCustomerLookup.MyCombo.ListIndex = IDToListIndex(uctlCustomerLookup.MyCombo, m_CommissionIncentive.CUSTOMER_ID)
      txtAmount.Text = m_CommissionIncentive.INCENTIVE_PER_PACK
      txtFromAmount.Text = m_CommissionIncentive.FROM_AMOUNT
      txtToAmount.Text = m_CommissionIncentive.TO_AMOUNT
      cboUnitType.ListIndex = m_CommissionIncentive.UNIT_TYPE
      sscAmountOver.value = FlagToCheck(m_CommissionIncentive.AMOUNT_OVER_FLAG)
      txtAmountOver.Text = m_CommissionIncentive.RATE_OVER
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean

   If Not VerifyCombo(lblFreelance, uctlFreelanceLookup.MyCombo, False) Then
      Exit Function
   End If
   

   If DocumentType = 1 Then
      If Not VerifyCombo(lblStockCode, uctlStockCodeLookup.MyCombo, False) Then
         Exit Function
      End If
   End If
   
   If DocumentType = 2 Then
      If Not VerifyCombo(lblStockCode, uctlStockCodeLookup.MyCombo, False) Then
         Exit Function
      End If
    
      If Not VerifyCombo(lblCustomerLookup, uctlCustomerLookup.MyCombo, False) Then
         Exit Function
      End If
   End If
   
   If DocumentType = 3 Or DocumentType = 4 Then
      If Not VerifyTextControl(lblFromAmount, txtFromAmount, False) Then
         Exit Function
      End If
      
      If Not VerifyTextControl(lblToAmount, txtToAmount, False) Then
         Exit Function
      End If
      
      If Not VerifyCombo(lblUnitType, cboUnitType, False) Then
         Exit Function
      End If
   End If
   
   If Not VerifyTextControl(lblAmount, txtAmount, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If

   If DocumentType = 1 Then
      If Not CheckUniqueNs(INCENTIVE_PD_UNIQUE, Trim(uctlFreelanceLookup.MyCombo.ItemData(Minus2Zero(uctlFreelanceLookup.MyCombo.ListIndex))), ID, Trim(uctlStockCodeLookup.MyCombo.ItemData(Minus2Zero(uctlStockCodeLookup.MyCombo.ListIndex)))) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & uctlStockCodeLookup.MyCombo.Text & " ของ " & uctlFreelanceLookup.MyCombo.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
         Exit Function
      End If
   ElseIf DocumentType = 2 Then
   If Not CheckUniqueNs(INCENTIVE_PD_CUS_UNIQUE, Trim(uctlFreelanceLookup.MyCombo.ItemData(Minus2Zero(uctlFreelanceLookup.MyCombo.ListIndex))), ID, Trim(uctlStockCodeLookup.MyCombo.ItemData(Minus2Zero(uctlStockCodeLookup.MyCombo.ListIndex))), 3, Trim(uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex)))) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & uctlStockCodeLookup.MyCombo.Text & " ของ " & uctlFreelanceLookup.MyCombo.Text & " ของลูกค้า " & uctlCustomerLookup.MyCombo.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
         Exit Function
      End If
   End If

   
   m_CommissionIncentive.INCENTIVE_ID = ID
   m_CommissionIncentive.AddEditMode = ShowMode
   m_CommissionIncentive.FREELANCE_ID = uctlFreelanceLookup.MyCombo.ItemData(Minus2Zero(uctlFreelanceLookup.MyCombo.ListIndex))
   m_CommissionIncentive.PART_ITEM_ID = uctlStockCodeLookup.MyCombo.ItemData(Minus2Zero(uctlStockCodeLookup.MyCombo.ListIndex))
   m_CommissionIncentive.CUSTOMER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
   m_CommissionIncentive.INCENTIVE_PER_PACK = Val(txtAmount.Text)
   m_CommissionIncentive.FROM_AMOUNT = Val(txtFromAmount.Text)
   m_CommissionIncentive.TO_AMOUNT = Val(txtToAmount.Text)
   m_CommissionIncentive.UNIT_TYPE = cboUnitType.ListIndex
   m_CommissionIncentive.AMOUNT_OVER_FLAG = Check2Flag(sscAmountOver.value)
   m_CommissionIncentive.RATE_OVER = Val(txtAmountOver.Text)

   m_CommissionIncentive.DOCUMENT_TYPE = DocumentType
      
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditCommissionIncentive(m_CommissionIncentive, IsOK, True, glbErrorLog) Then
      SaveData = False
      Call EnableForm(Me, True)
      Exit Function
   End If
   If Not IsOK Then
      Call EnableForm(Me, True)
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   Call EnableForm(Me, True)
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadFreelance(uctlFreelanceLookup.MyCombo, m_Freelances)
      Set uctlFreelanceLookup.MyCollection = m_Freelances
      
      Call LoadStockPartItem(uctlStockCodeLookup.MyCombo, m_StockCodes)
      Set uctlStockCodeLookup.MyCollection = m_StockCodes
      
      Call LoadCustomer(uctlCustomerLookup.MyCombo, m_Customers)
      Set uctlCustomerLookup.MyCollection = m_Customers
      
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
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

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblFreelance, MapText("ชื่อฟรีแลนซ์"))
   Call InitNormalLabel(lblStockCode, MapText("สินค้า"))
   Call InitNormalLabel(lblAmount, MapText("บาท/หน่วย"))
   Call InitNormalLabel(lblCustomerLookup, MapText("ชื่อลูกค้า"))
   Call InitNormalLabel(lblFromAmount, MapText("ยอดขายตั้งแต่"))
   Call InitNormalLabel(lblToAmount, MapText("จนถึงยอดขาย"))
   Call InitNormalLabel(lblUnitType, MapText("หน่วย"))
    Call InitCheckBox(sscAmountOver, MapText("เกินยอดที่กำหนด"))
   Call InitNormalLabel(lblAmountOver, MapText("คิดหน่วยละ"))
   Call InitNormalLabel(Label2, MapText("บาท"))
   
   Call txtAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtFromAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtToAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtAmountOver.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   
   Call InitCombo(cboUnitType)
   
   If DocumentType = 1 Then
      lblCustomerLookup.Enabled = False
      uctlCustomerLookup.Enabled = False
      lblFromAmount.Enabled = False
      txtFromAmount.Enabled = False
      lblToAmount.Enabled = False
      txtToAmount.Enabled = False
      lblUnitType.Enabled = False
      cboUnitType.Enabled = False
      sscAmountOver.Enabled = False
      lblAmountOver.Enabled = False
      txtAmountOver.Enabled = False
      Label2.Enabled = False
   ElseIf DocumentType = 2 Then
      lblFromAmount.Enabled = False
      txtFromAmount.Enabled = False
      lblToAmount.Enabled = False
      txtToAmount.Enabled = False
      lblUnitType.Enabled = False
      cboUnitType.Enabled = False
      sscAmountOver.Enabled = False
      lblAmountOver.Enabled = False
      txtAmountOver.Enabled = False
      Label2.Enabled = False
   ElseIf DocumentType = 3 Or DocumentType = 4 Then
      lblCustomerLookup.Enabled = False
      uctlCustomerLookup.Enabled = False
      lblStockCode.Enabled = False
      uctlStockCodeLookup.Enabled = False
   End If
   Call InitUnitType(cboUnitType)
   
   uctlFreelanceLookup.MyTextBox.SetKeySearch ("FREELANCE_CODE")
   uctlStockCodeLookup.MyTextBox.SetKeySearch ("PART_NO")
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   
   Call InitMainButton(cmdNext, MapText("ถัดไป"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   Call EnableForm(Me, False)
   m_HasActivate = False
   
   Set m_CommissionIncentive = New CCommissionIncentive
   Set m_Rs = New ADODB.Recordset
   
   Set m_Freelances = New Collection
   Set m_StockCodes = New Collection
   Set m_Customers = New Collection
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_Freelances = Nothing
   Set m_StockCodes = Nothing
   Set m_Customers = Nothing
End Sub

Private Sub sscAmountOver_Click(value As Integer)
   m_HasModify = True
End Sub

Private Sub sscAmountOver_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub txtAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtAmountOver_Change()
   m_HasModify = True
End Sub

Private Sub txtFromAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtToAmount_Change()
   m_HasModify = True
End Sub

Private Sub uctlCustomerLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlFreelanceLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlStockCodeLookup_Change()
   m_HasModify = True
End Sub
