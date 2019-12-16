VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditCommissionCredit 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Icon            =   "frmAddEditCommissionCredit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   10410
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   4845
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   8546
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   7
         Top             =   0
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextLookup uctlCustomerLookup 
         Height          =   435
         Left            =   2100
         TabIndex        =   0
         Top             =   1440
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtAmount 
         Height          =   435
         Left            =   2100
         TabIndex        =   1
         Top             =   1920
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtDesc 
         Height          =   435
         Left            =   2100
         TabIndex        =   2
         Top             =   2400
         Width           =   7515
         _ExtentX        =   5212
         _ExtentY        =   767
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   2280
         TabIndex        =   3
         Top             =   3390
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCommissionCredit.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   450
         TabIndex        =   10
         Top             =   2490
         Width           =   1575
      End
      Begin VB.Label lblAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   450
         TabIndex        =   9
         Top             =   2010
         Width           =   1575
      End
      Begin VB.Label lblCustomer 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   450
         TabIndex        =   8
         Top             =   1560
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5595
         TabIndex        =   5
         Top             =   3390
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3945
         TabIndex        =   4
         Top             =   3390
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCommissionCredit.frx":2ABC
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditCommissionCredit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_CommissionCredit As CCommissionCredit

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Private m_Customers As Collection
Public ParentForm As Object
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
Dim itemcount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      m_CommissionCredit.COMMISSION_CREDIT_ID = ID
      m_CommissionCredit.QueryFlag = 1
      If Not glbDaily.QueryCommissionCredit(m_CommissionCredit, m_Rs, itemcount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If itemcount > 0 Then
      Call m_CommissionCredit.PopulateFromRS(1, m_Rs)
      
      uctlCustomerLookup.MyCombo.ListIndex = IDToListIndex(uctlCustomerLookup.MyCombo, m_CommissionCredit.CUSTOMER_ID)
      txtAmount.Text = m_CommissionCredit.COMMISSION_CREDIT_AMOUNT
      txtDesc.Text = m_CommissionCredit.COMMISSION_CREDIT_DESC
      
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
      If Not VerifyAccessRight("COMMISSION_CREDIT_EDIT") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   End If
   
   If Not VerifyCombo(lblCustomer, uctlCustomerLookup.MyCombo, False) Then
      Exit Function
   End If
   
   If Not CheckUniqueNs(MASTER_COMMISSION_CREDIT, uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex)), ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & uctlCustomerLookup.MyTextBox.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_CommissionCredit.COMMISSION_CREDIT_ID = ID
   m_CommissionCredit.AddEditMode = ShowMode
   m_CommissionCredit.CUSTOMER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
   m_CommissionCredit.COMMISSION_CREDIT_AMOUNT = Val(txtAmount.Text)
   m_CommissionCredit.COMMISSION_CREDIT_DESC = txtDesc.Text
      
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditCommissionCredit(m_CommissionCredit, IsOK, True, glbErrorLog) Then
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
   
   Call InitNormalLabel(lblCustomer, MapText("ลูกค้า"))
   Call InitNormalLabel(lblAmount, MapText("ยอดเครดิตใหม่"))
   Call InitNormalLabel(lblDesc, MapText("ราละเอียด"))
   
   uctlCustomerLookup.MyTextBox.SetKeySearch ("CUSTOMER_CODE")
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
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
   
   Set m_CommissionCredit = New CCommissionCredit
   Set m_Rs = New ADODB.Recordset
   Set m_Customers = New Collection
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_Customers = Nothing
End Sub

Private Sub txtAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtDesc_Change()
   m_HasModify = True
End Sub
Private Sub uctlCustomerLookup_Change()
   m_HasModify = True
End Sub
