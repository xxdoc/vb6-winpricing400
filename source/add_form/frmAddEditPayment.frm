VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddEditPayment 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditPayment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8520
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboBankAccount 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   930
         Width           =   3045
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   2805
         Left            =   150
         TabIndex        =   13
         Top             =   4950
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   4948
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         TabKeyBehavior  =   1
         MethodHoldFields=   -1  'True
         AllowColumnDrag =   0   'False
         AllowEdit       =   0   'False
         BorderStyle     =   3
         GroupByBoxVisible=   0   'False
         DataMode        =   99
         HeaderFontName  =   "AngsanaUPC"
         HeaderFontBold  =   -1  'True
         HeaderFontSize  =   14.25
         HeaderFontWeight=   700
         FontSize        =   9.75
         BackColorBkg    =   16777215
         ColumnHeaderHeight=   480
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmAddEditPayment.frx":27A2
         Column(2)       =   "frmAddEditPayment.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditPayment.frx":290E
         FormatStyle(2)  =   "frmAddEditPayment.frx":2A6A
         FormatStyle(3)  =   "frmAddEditPayment.frx":2B1A
         FormatStyle(4)  =   "frmAddEditPayment.frx":2BCE
         FormatStyle(5)  =   "frmAddEditPayment.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditPayment.frx":2D5E
      End
      Begin prjFarmManagement.uctlTextLookup uctlBankLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   4
         Top             =   1350
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlDocumentDate 
         Height          =   405
         Left            =   6570
         TabIndex        =   2
         Top             =   900
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   12
         Top             =   4410
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   979
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "AngsanaUPC"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin prjFarmManagement.uctlTextBox txtAccountNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   1
         Top             =   1350
         Visible         =   0   'False
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtTotalAmount 
         Height          =   435
         Left            =   1860
         TabIndex        =   6
         Top             =   2250
         Width           =   1845
         _ExtentX        =   2831
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlSellByLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   7
         Top             =   2730
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlBankBranchLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   5
         Top             =   1800
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlCustomerLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   8
         Top             =   3180
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtReceiptNo 
         Height          =   435
         Left            =   9120
         TabIndex        =   9
         Top             =   3180
         Width           =   2535
         _ExtentX        =   2831
         _ExtentY        =   767
      End
      Begin VB.Label lblReceiptNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7290
         TabIndex        =   29
         Top             =   3300
         Width           =   1695
      End
      Begin VB.Label lblCustomer 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   150
         TabIndex        =   28
         Top             =   3240
         Width           =   1635
      End
      Begin VB.Label lblBankBranch 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   150
         TabIndex        =   27
         Top             =   1860
         Width           =   1635
      End
      Begin Threed.SSCheck chkCommit 
         Height          =   435
         Left            =   10500
         TabIndex        =   3
         Top             =   900
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblSellBy 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   150
         TabIndex        =   26
         Top             =   2790
         Width           =   1635
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   525
         Left            =   1860
         TabIndex        =   10
         Top             =   3660
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPayment.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   525
         Left            =   3480
         TabIndex        =   11
         Top             =   3660
         Visible         =   0   'False
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblBank 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   150
         TabIndex        =   25
         Top             =   1410
         Width           =   1635
      End
      Begin VB.Label Label1 
         Height          =   315
         Left            =   3810
         TabIndex        =   24
         Top             =   2340
         Width           =   405
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5040
         TabIndex        =   23
         Top             =   930
         Width           =   1395
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   17
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPayment.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   18
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   15
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   14
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPayment.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   16
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPayment.frx":3884
         ButtonStyle     =   3
      End
      Begin VB.Label lblTotalAmount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         TabIndex        =   21
         Top             =   2370
         Width           =   1695
      End
      Begin VB.Label lblAccountNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   90
         TabIndex        =   20
         Top             =   960
         Width           =   1665
      End
   End
End
Attribute VB_Name = "frmAddEditPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Payment As CPayment
Private m_Customers As Collection
Private m_Employees As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public id As Long

Private FileName As String
Private m_SumUnit As Double
Public DocumentSubType As Long
Public DocumentType As Long
Public Direction As String
Public m_BankBranchs As Collection
Public m_BankAccounts As Collection

Private Sub ShowButton(Ind As Long)
   If ShowMode = SHOW_ADD Then
      Exit Sub
   End If
   
   If Ind = 1 Then
      cmdAdd.Enabled = (m_Payment.COMMIT_FLAG = "N")
      cmdDelete.Enabled = (m_Payment.COMMIT_FLAG = "N")
      cmdEdit.Enabled = True
   ElseIf Ind = 2 Then
      cmdAdd.Enabled = False
      cmdEdit.Enabled = False

      cmdDelete.Enabled = False
   End If
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
      m_Payment.PAYMENT_ID = id
      If Not glbDaily.QueryPayment(m_Payment, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_Payment.PopulateFromRS(1, m_Rs)
      
      cboBankAccount.ListIndex = IDToListIndex(cboBankAccount, m_Payment.BANK_ACCOUNT)
      txtAccountNo.Text = m_Payment.ACCOUNT_NO
      uctlDocumentDate.ShowDate = m_Payment.PAYMENT_DATE
      uctlBankLookup.MyCombo.ListIndex = IDToListIndex(uctlBankLookup.MyCombo, m_Payment.TO_BANK_ID)
      uctlBankBranchLookup.MyCombo.ListIndex = IDToListIndex(uctlBankBranchLookup.MyCombo, m_Payment.TO_BANK_BRANCH)
      txtTotalAmount.Text = Format(m_Payment.TOTAL_AMOUNT, "0.00")
      uctlSellByLookup.MyCombo.ListIndex = IDToListIndex(uctlSellByLookup.MyCombo, m_Payment.ACCEPT_BY)
      uctlCustomerLookup.MyCombo.ListIndex = IDToListIndex(uctlCustomerLookup.MyCombo, m_Payment.CUSTOMER_ID)
      txtReceiptNo.Text = m_Payment.RECEIPT_NO
      
      chkCommit.Value = FlagToCheck(m_Payment.COMMIT_FLAG)
      chkCommit.Enabled = (m_Payment.COMMIT_FLAG = "N")
      
      Call ShowButton(1)
   End If
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call TabStrip1_Click
   
   Call EnableForm(Me, True)
End Sub

Private Sub cboGroup_Click()
   m_HasModify = True
End Sub

Private Sub chkEnable_Click()
   m_HasModify = True
End Sub


Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim Ivd As CInventoryDoc
Dim Pm As CPayment
   
   If Not VerifyCombo(lblAccountNo, cboBankAccount, False) Then
      Exit Function
   End If

   If Not VerifyDate(lblDocumentDate, uctlDocumentDate, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblTotalAmount, txtTotalAmount, True) Then
      Exit Function
   End If
         
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_Payment.BANK_ACCOUNT = cboBankAccount.ItemData(Minus2Zero(cboBankAccount.ListIndex))
   m_Payment.ACCOUNT_NO = txtAccountNo.Text
   m_Payment.AddEditMode = ShowMode
   m_Payment.PAYMENT_ID = id
    m_Payment.PAYMENT_DATE = uctlDocumentDate.ShowDate
   m_Payment.TO_BANK_BRANCH = uctlBankBranchLookup.MyCombo.ItemData(Minus2Zero(uctlBankBranchLookup.MyCombo.ListIndex))
   m_Payment.TO_BANK_ID = uctlBankLookup.MyCombo.ItemData(Minus2Zero(uctlBankLookup.MyCombo.ListIndex))
   m_Payment.TOTAL_AMOUNT = Val(txtTotalAmount.Text)
   m_Payment.TX_TYPE = "O"
   m_Payment.ACCEPT_BY = uctlSellByLookup.MyCombo.ItemData(Minus2Zero(uctlSellByLookup.MyCombo.ListIndex))
   m_Payment.COMMIT_FLAG = Check2Flag(chkCommit.Value)
   m_Payment.INTERNAL_FLAG = "N"
   m_Payment.CUSTOMER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
   m_Payment.RECEIPT_NO = txtReceiptNo.Text
   
   Call EnableForm(Me, False)
   
   Call glbDaily.StartTransaction

   If Not glbDaily.AddEditPayment(m_Payment, IsOK, False, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call glbDaily.RollbackTransaction
      Call EnableForm(Me, True)
      Exit Function
   End If
   
   If Not IsOK Then
      Call EnableForm(Me, True)
      glbErrorLog.ShowUserError
      Call glbDaily.RollbackTransaction
      Exit Function
   End If
   
   Call glbDaily.CommitTransaction
   Call EnableForm(Me, True)
   SaveData = True
End Function

Private Sub cboBusinessGroup_Click()
   m_HasModify = True
End Sub

Private Sub cboBusinessType_Click()
   m_HasModify = True
End Sub

Private Sub cboEnterpriseType_Click()
   m_HasModify = True
End Sub

Private Sub cboAccount_Click()
   m_HasModify = True
End Sub

Private Sub cboAccount_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cboBank_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cboBankBranch_Click()
   m_HasModify = True
End Sub

Private Sub cboBankBranch_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cboCustomerAddress_Click()
   m_HasModify = True
End Sub

Private Sub cboCustomerAddress_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cboEnpAddress_Click()
   m_HasModify = True
End Sub

Private Sub cboEnpAddress_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cboPaymentType_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cboBankAccount_Click()
Dim TempID As Long
Dim Mr As CMasterRef

   m_HasModify = True
   
   TempID = cboBankAccount.ItemData(Minus2Zero(cboBankAccount.ListIndex))
   If TempID > 0 Then
      Set Mr = GetMasterRef(m_BankAccounts, Trim(str(TempID)))
      
      txtAccountNo.Text = cboBankAccount.Text
      uctlBankLookup.MyCombo.ListIndex = IDToListIndex(uctlBankLookup.MyCombo, Mr.TEMP_ID1)
      uctlBankBranchLookup.MyCombo.ListIndex = IDToListIndex(uctlBankBranchLookup.MyCombo, Mr.TEMP_ID2)
   Else
      txtAccountNo.Text = ""
      uctlBankLookup.MyCombo.ListIndex = -1
      uctlBankBranchLookup.MyCombo.ListIndex = -1
   End If
End Sub

Private Sub chkCommit_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkCommit_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub chkExtraFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu

   If Not cmdAdd.Enabled Then
      Exit Sub
   End If
   
   OKClick = False
   If TabStrip1.SelectedItem.Index = 1 Then
      frmAddEditPaymentItem.COMMIT_FLAG = m_Payment.COMMIT_FLAG
      Set frmAddEditPaymentItem.TempCollection = m_Payment.PaymentItems
      frmAddEditPaymentItem.ParentShowMode = ShowMode
      frmAddEditPaymentItem.ShowMode = SHOW_ADD
      frmAddEditPaymentItem.HeaderText = MapText("������¡�ùӽҡ��Ҥ��")
      Load frmAddEditPaymentItem
      frmAddEditPaymentItem.Show 1

      OKClick = frmAddEditPaymentItem.OKClick

      Unload frmAddEditPaymentItem
      Set frmAddEditPaymentItem = Nothing
      
      If OKClick Then
         Call GetTotalPrice

         GridEX1.ItemCount = CountItem(m_Payment.PaymentItems)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdDelete_Click()
Dim ID1 As Long
Dim ID2 As Long

   If Not cmdDelete.Enabled Then
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
   If Not ConfirmDelete(GridEX1.Value(3)) Then
      Exit Sub
   End If
   
   ID2 = GridEX1.Value(2)
   ID1 = GridEX1.Value(1)
   
   If TabStrip1.SelectedItem.Index = 1 Then
      If ID1 <= 0 Then
         m_Payment.PaymentItems.Remove (ID2)
      Else
         m_Payment.PaymentItems.Item(ID2).Flag = "D"
      End If

      Call GetTotalPrice
      GridEX1.ItemCount = CountItem(m_Payment.PaymentItems)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim id As Long
Dim OKClick As Boolean
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   id = Val(GridEX1.Value(2))
   OKClick = False

   If TabStrip1.SelectedItem.Index = 1 Then
      frmAddEditPaymentItem.id = id
      frmAddEditPaymentItem.COMMIT_FLAG = m_Payment.COMMIT_FLAG
      Set frmAddEditPaymentItem.TempCollection = m_Payment.PaymentItems
      frmAddEditPaymentItem.HeaderText = MapText("�����¡�ùӽҡ��Ҥ��")
      frmAddEditPaymentItem.ParentShowMode = ShowMode
      frmAddEditPaymentItem.ShowMode = SHOW_EDIT
      Load frmAddEditPaymentItem
      frmAddEditPaymentItem.Show 1

      OKClick = frmAddEditPaymentItem.OKClick

      Unload frmAddEditPaymentItem
      Set frmAddEditPaymentItem = Nothing
      
      If OKClick Then
         Call GetTotalPrice
         GridEX1.ItemCount = CountItem(m_Payment.PaymentItems)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub CalculateIncludePrice()
'Dim II As CImportItem
'Dim AvgFee As Double
'
'   If m_SumUnit > 0 Then
'      AvgFee = Val(txtTotalAmount.Text) / m_SumUnit
'   Else
'      AvgFee = 0
'   End If
'
'   For Each II In m_Payment.DoItems
'      If II.Flag <> "D" Then
'         II.INCLUDE_UNIT_PRICE = II.ACTUAL_UNIT_PRICE + AvgFee
'      End If
'   Next II
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Sub cmdPictureAdd_Click()
On Error Resume Next
Dim strDescription As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Picture Files (*.jpg, *.gif)|*.jpg;*.gif"
   dlgAdd.DialogTitle = "Select Picture to Add to Database"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
   m_HasModify = True
End Sub

Private Sub cmdSave_Click()
Dim Result As Boolean
   If Not SaveData Then
      Exit Sub
   End If
   
   ShowMode = SHOW_EDIT
   id = m_Payment.PAYMENT_ID
   m_Payment.QueryFlag = 1
   QueryData (True)
   m_HasModify = False
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      
      Call LoadCustomer(uctlCustomerLookup.MyCombo, m_Customers)
      Set uctlCustomerLookup.MyCollection = m_Customers
      
      Call LoadMaster(cboBankAccount, m_BankAccounts, BANK_ACCOUNT)
      
      Call LoadBank(uctlBankLookup.MyCombo, m_Customers)
      Set uctlBankLookup.MyCollection = m_Customers
      
      Call LoadEmployee(uctlSellByLookup.MyCombo, m_Employees)
      Set uctlSellByLookup.MyCollection = m_Employees
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_Payment.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         m_Payment.QueryFlag = 0
         Call QueryData(False)
      End If
      
      Call EnableForm(Me, True)
      m_HasModify = False
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Static InUsed As Long

   If InUsed = 1 Then
      Exit Sub
   End If
   
   InUsed = 1
   
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
      Call cmdAdd_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
      Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 123 Then
      Call AddMemoNote
      KeyCode = 0
   End If
   
   InUsed = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_Payment = Nothing
   Set m_Customers = Nothing
   Set m_Employees = Nothing
   Set m_BankBranchs = Nothing
   Set m_BankAccounts = Nothing
   Set m_Customers = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid1()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.NAME = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   Set Col = GridEX1.Columns.add '3
   Col.Width = 2220
   Col.Caption = MapText("������")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 2415
   Col.Caption = MapText("�Ţ�����")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 4740
   Col.Caption = MapText("��Ҥ��-�Ң�")

   Set Col = GridEX1.Columns.add '6
   Col.Width = 2190
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("�ӹǹ�Թ")
End Sub

Private Sub InitGrid2()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.NAME = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   Set Col = GridEX1.Columns.add '3
   Col.Width = 2400
   Col.Caption = MapText("�������ѵ�شԺ")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 1725
   Col.Caption = MapText("�����ѵ�شԺ")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 3420
   Col.Caption = MapText("�����ѵ�شԺ")

   Set Col = GridEX1.Columns.add '6
   Col.Width = 1770
   Col.Caption = MapText("�ӹǹ")
      
   Set Col = GridEX1.Columns.add '7
   Col.TextAlignment = jgexAlignRight
   Col.Width = 1890
   Col.Caption = MapText("�Ҥ����")
   
   Set Col = GridEX1.Columns.add '8
   Col.TextAlignment = jgexAlignRight
   Col.Width = 1290
   Col.Caption = MapText("�Ҥ�/˹���")

   Set Col = GridEX1.Columns.add '9
   Col.Width = 2235
   Col.Caption = MapText("ʶҹ���Ѵ��")
End Sub

Private Sub GetTotalPrice()
Dim II As CPaymentItem
Dim Sum1 As Double
Dim Sum2 As Double
Dim Sum3 As Double

   Sum1 = 0
   For Each II In m_Payment.PaymentItems
      If II.Flag <> "D" Then
         Sum1 = Sum1 + II.PAY_AMOUNT
      End If
   Next II

   txtTotalAmount.Text = Format(Sum1, "0.00")
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblAccountNo, MapText("�Ţ���ѭ��"))
   Call InitNormalLabel(lblTotalAmount, MapText("�ӹǹ���"))
   Call InitNormalLabel(lblDocumentDate, MapText("�ѹ���ӽҡ"))
   Call InitNormalLabel(lblBank, MapText("��Ҥ��"))
   Call InitNormalLabel(lblCustomer, MapText("�����١���"))
   Call InitNormalLabel(lblReceiptNo, MapText("��ҧ�֧�����"))

   Call InitNormalLabel(lblBankBranch, MapText("�Ң�"))
   
   Call InitNormalLabel(Label1, MapText("�ҷ"))
   Call InitNormalLabel(lblSellBy, MapText("������¡��"))
   
   Call InitCombo(cboBankAccount)
   Call InitCheckBox(chkCommit, "�ӹǳ")
   
   Call txtTotalAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtTotalAmount.Enabled = False
   Call txtReceiptNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   uctlBankLookup.Enabled = False
   uctlBankBranchLookup.Enabled = False
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
      
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPrint.Picture = LoadPicture(glbParameterObj.NormalButton1)
'   cmdPrint.Enabled = False
   cmdSave.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
   Call InitMainButton(cmdAdd, MapText("���� (F7)"))
   Call InitMainButton(cmdEdit, MapText("��� (F3)"))
   Call InitMainButton(cmdDelete, MapText("ź (F6)"))
   Call InitMainButton(cmdSave, MapText("�ѹ�֡"))
   Call InitMainButton(cmdPrint, MapText("�����"))
   
   If DocumentSubType = 1 Then
      Call InitGrid1
   ElseIf DocumentSubType = 2 Then
      Call InitGrid2
   End If
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.NAME = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   If Direction = "O" Then
      TabStrip1.Tabs.add().Caption = MapText("��¡��㺹ӽҡ")
   ElseIf Direction = "I" Then
      TabStrip1.Tabs.add().Caption = MapText("��¡��㺶͹�Թ")
   End If
   Call InitGrid1
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   OKClick = False
   Call InitFormLayout
      
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_Payment = New CPayment
   Set m_Customers = New Collection
   Set m_Employees = New Collection
   Set m_BankBranchs = New Collection
   Set m_BankAccounts = New Collection
   Set m_Customers = New Collection
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   If TabStrip1.SelectedItem.Index = 5 Then
      RowBuffer.RowStyle = RowBuffer.Value(7)
   End If
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Index = 1 Then
      If m_Payment.PaymentItems Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim CR As CPaymentItem
      If m_Payment.PaymentItems.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_Payment.PaymentItems, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If

      Values(1) = CR.PAYMENT_ID
      Values(2) = RealIndex
      Values(3) = PaymentTypeToText(CR.PAYMENT_TYPE)
      Values(4) = CR.CHECK_NO
      Values(5) = CR.BANK_NAME & "-" & CR.BANK_BRANCH_NAME
      Values(6) = FormatNumber(CR.PAY_AMOUNT)
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
      
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub TabStrip1_Click()
   If TabStrip1.SelectedItem.Index = 1 Then
      If DocumentSubType = 1 Then
         Call InitGrid1
      ElseIf DocumentSubType = 2 Then
         Call InitGrid2
      End If
      
      Call GetTotalPrice
      GridEX1.ItemCount = CountItem(m_Payment.PaymentItems)
      GridEX1.Rebind
      GridEX1.Visible = True
      
      Call ShowButton(TabStrip1.SelectedItem.Index)
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub txtCheckNo_Change()
   m_HasModify = True
End Sub

Private Sub txtNote_Change()
   m_HasModify = True
End Sub

Private Sub txtDoNo_Change()
   m_HasModify = True
End Sub

Private Sub txtDeliveryNo_Change()
   m_HasModify = True
End Sub

Private Sub txtSellBy_Change()
   m_HasModify = True
End Sub

Private Sub txtDocumentNo_Change()
   m_HasModify = True
End Sub

Private Sub txtPassword_Change()
   m_HasModify = True
End Sub

Private Sub txtSender_Change()
   m_HasModify = True
End Sub

Private Sub txtTotal_Change()
   m_HasModify = True
End Sub

Private Sub txtTruckNo_Change()
   m_HasModify = True
End Sub

Private Sub uctlSetupDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlTextBox1_Change()
   m_HasModify = True
End Sub

Private Sub uctlDeliveryLookup_Change()
   m_HasModify = True
End Sub

Private Sub txtAccountNo_Change()
   m_HasModify = True
End Sub

Private Sub txtReceiptNo_Change()
   m_HasModify = True
End Sub

Private Sub uctlBankLookup_Change()
Dim BankID As Long

   BankID = uctlBankLookup.MyCombo.ItemData(Minus2Zero(uctlBankLookup.MyCombo.ListIndex))
   If BankID > 0 Then
      Call LoadBankBranch(uctlBankBranchLookup.MyCombo, m_BankBranchs, BankID)
      Set uctlBankBranchLookup.MyCollection = m_BankBranchs
   End If
End Sub

Private Sub uctlCustomerLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlDocumentDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlSellByLookup_Change()
   m_HasModify = True
End Sub
