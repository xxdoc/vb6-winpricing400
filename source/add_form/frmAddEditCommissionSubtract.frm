VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditCommissionSubtract 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Icon            =   "frmAddEditCommissionSubtract.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   10410
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   4845
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   8546
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboMonthID 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1440
         Width           =   2955
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   10
         Top             =   0
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtYearNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   990
         Width           =   2955
         _extentx        =   5212
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlCustomerLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   3
         Top             =   2400
         Width           =   5385
         _extentx        =   9499
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlEmployeeLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   2
         Top             =   1920
         Width           =   5385
         _extentx        =   9499
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtAmount 
         Height          =   435
         Left            =   1860
         TabIndex        =   4
         Top             =   2880
         Width           =   2955
         _extentx        =   5212
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtDesc 
         Height          =   435
         Left            =   1860
         TabIndex        =   5
         Top             =   3360
         Width           =   7515
         _extentx        =   5212
         _extenty        =   767
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   1920
         TabIndex        =   6
         Top             =   4110
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCommissionSubtract.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   16
         Top             =   3450
         Width           =   1575
      End
      Begin VB.Label lblAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   15
         Top             =   2970
         Width           =   1575
      End
      Begin VB.Label lblEmployee 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   14
         Top             =   2070
         Width           =   1575
      End
      Begin VB.Label lblCustomer 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   13
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label lblMonthID 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   12
         Top             =   1500
         Width           =   1575
      End
      Begin VB.Label lblYearNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   11
         Top             =   1080
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5235
         TabIndex        =   8
         Top             =   4110
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3585
         TabIndex        =   7
         Top             =   4110
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCommissionSubtract.frx":2ABC
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditCommissionSubtract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_CommissionSubtract As CCommissionSubtract

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Private m_Customers As Collection
Private m_Employees As Collection
Public ParentForm As Object
Private Sub cboMonthID_Click()
   m_HasModify = True
End Sub

Private Sub cmdNext_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   uctlCustomerLookup.MyCombo.ListIndex = -1
   txtAmount.Text = ""
   txtDesc.Text = ""
   Call ParentForm.QueryData(True)
   Call uctlCustomerLookup.MyTextBox.SetFocus
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
      
      m_CommissionSubtract.COMMISSION_SUBTRACT_ID = ID
      m_CommissionSubtract.QueryFlag = 1
      If Not glbDaily.QueryCommissionSubtract(m_CommissionSubtract, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_CommissionSubtract.PopulateFromRS(1, m_Rs)
      
      txtYearNo.Text = m_CommissionSubtract.YEAR_NO + 543
      cboMonthID.ListIndex = IDToListIndex(cboMonthID, m_CommissionSubtract.MONTH_ID)
      uctlEmployeeLookup.MyCombo.ListIndex = IDToListIndex(uctlEmployeeLookup.MyCombo, m_CommissionSubtract.EMP_ID)
      uctlCustomerLookup.MyCombo.ListIndex = IDToListIndex(uctlCustomerLookup.MyCombo, m_CommissionSubtract.CUSTOMER_ID)
      txtAmount.Text = m_CommissionSubtract.COMMISSION_SUBTRACT_AMOUNT
      txtDesc.Text = m_CommissionSubtract.COMMISSION_SUBTRACT_DESC
      
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
   If ShowMode = SHOW_EDIT Then
      If Not VerifyAccessRight("COMMISSION_SUBTRACT_EDIT") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   End If
   
   If Not VerifyTextControl(lblYearNo, txtYearNo, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblMonthID, cboMonthID, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblEmployee, uctlEmployeeLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblCustomer, uctlCustomerLookup.MyCombo, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_CommissionSubtract.COMMISSION_SUBTRACT_ID = ID
   m_CommissionSubtract.AddEditMode = ShowMode
   m_CommissionSubtract.YEAR_NO = Val(txtYearNo.Text) - 543
   m_CommissionSubtract.MONTH_ID = cboMonthID.ItemData(Minus2Zero(cboMonthID.ListIndex))
   m_CommissionSubtract.EMP_ID = uctlEmployeeLookup.MyCombo.ItemData(Minus2Zero(uctlEmployeeLookup.MyCombo.ListIndex))
   m_CommissionSubtract.CUSTOMER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
   m_CommissionSubtract.COMMISSION_SUBTRACT_AMOUNT = Val(txtAmount.Text)
   m_CommissionSubtract.COMMISSION_SUBTRACT_DESC = txtDesc.Text
      
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditCommissionSubtract(m_CommissionSubtract, IsOK, True, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
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
      
      Call InitThaiMonth(cboMonthID)
      
      Call LoadCustomer(uctlCustomerLookup.MyCombo, m_Customers)
      Set uctlCustomerLookup.MyCollection = m_Customers
      
      Call LoadEmployee(uctlEmployeeLookup.MyCombo, m_Employees)
      Set uctlEmployeeLookup.MyCollection = m_Employees
      
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
   
   Call InitNormalLabel(lblYearNo, MapText("ปี"))
   Call InitNormalLabel(lblMonthID, MapText("เดือน"))
   Call InitNormalLabel(lblEmployee, MapText("พนักงาน"))
   Call InitNormalLabel(lblCustomer, MapText("ลูกค้า"))
   Call InitNormalLabel(lblAmount, MapText("ยอดหัก"))
   Call InitNormalLabel(lblDesc, MapText("ราละเอียด"))
   
   Call txtYearNo.SetTextLenType(TEXT_INTEGER, glbSetting.AMOUNT_LEN)
   
   uctlEmployeeLookup.MyTextBox.SetKeySearch ("EMP_CODE")
  uctlCustomerLookup.MyTextBox.SetKeySearch ("CUSTOMER_CODE")
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitCombo(cboMonthID)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdNext, MapText("ถัดไป"))
   
   If ShowMode = SHOW_EDIT Then
      cmdNext.Enabled = False
   End If
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
   
   Set m_CommissionSubtract = New CCommissionSubtract
   Set m_Rs = New ADODB.Recordset
   Set m_Customers = New Collection
   Set m_Employees = New Collection
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_Customers = Nothing
   Set m_Employees = Nothing
End Sub

Private Sub txtAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtYearNo_Change()
   m_HasModify = True
End Sub

Private Sub uctlCustomerLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlEmployeeLookup_Change()
   m_HasModify = True
End Sub
