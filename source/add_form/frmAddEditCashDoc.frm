VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddEditCashDoc 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditCashDoc.frx":0000
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
      TabIndex        =   16
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextLookup uctlBankLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   3
         Top             =   1350
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlDocumentDate 
         Height          =   405
         Left            =   6300
         TabIndex        =   2
         Top             =   930
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   8
         Top             =   3720
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
      Begin prjFarmManagement.uctlTextBox txtDocumentNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   900
         Width           =   2385
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
      Begin GridEX20.GridEX GridEX1 
         Height          =   3495
         Left            =   150
         TabIndex        =   9
         Top             =   4230
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   6165
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
         Column(1)       =   "frmAddEditCashDoc.frx":27A2
         Column(2)       =   "frmAddEditCashDoc.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditCashDoc.frx":290E
         FormatStyle(2)  =   "frmAddEditCashDoc.frx":2A6A
         FormatStyle(3)  =   "frmAddEditCashDoc.frx":2B1A
         FormatStyle(4)  =   "frmAddEditCashDoc.frx":2BCE
         FormatStyle(5)  =   "frmAddEditCashDoc.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditCashDoc.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextLookup uctlBankBranchLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   4
         Top             =   1800
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlBankAccountLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   5
         Top             =   2250
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlEmployeeLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   6
         Top             =   2700
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtOut 
         Height          =   435
         Left            =   8760
         TabIndex        =   25
         Top             =   1770
         Width           =   1545
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtIn 
         Height          =   435
         Left            =   8760
         TabIndex        =   27
         Top             =   2250
         Width           =   1545
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtLeft 
         Height          =   435
         Left            =   8760
         TabIndex        =   29
         Top             =   2730
         Width           =   1545
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlCustomerLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   31
         Top             =   3150
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin VB.Label lblCustomer 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   360
         TabIndex        =   32
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label lblLeft 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7260
         TabIndex        =   30
         Top             =   2790
         Width           =   1395
      End
      Begin VB.Label lblIn 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7260
         TabIndex        =   28
         Top             =   2310
         Width           =   1395
      End
      Begin VB.Label lblOut 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7260
         TabIndex        =   26
         Top             =   1830
         Width           =   1395
      End
      Begin VB.Label lblEmployee 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   330
         TabIndex        =   24
         Top             =   2730
         Width           =   1455
      End
      Begin VB.Label lblBankAccount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   330
         TabIndex        =   23
         Top             =   2310
         Width           =   1455
      End
      Begin VB.Label lblBankBranch 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   330
         TabIndex        =   22
         Top             =   1860
         Width           =   1455
      End
      Begin Threed.SSCommand cmdAuto 
         Height          =   405
         Left            =   4260
         TabIndex        =   1
         Top             =   900
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCashDoc.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   525
         Left            =   6840
         TabIndex        =   13
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCashDoc.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCheck chkCommit 
         Height          =   435
         Left            =   7350
         TabIndex        =   7
         Top             =   1350
         Visible         =   0   'False
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblBank 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   330
         TabIndex        =   21
         Top             =   1410
         Width           =   1455
      End
      Begin VB.Label Label4 
         Height          =   315
         Left            =   11250
         TabIndex        =   20
         Top             =   3420
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5010
         TabIndex        =   19
         Top             =   960
         Width           =   1155
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   14
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCashDoc.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   15
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCashDoc.frx":3884
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   12
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCashDoc.frx":3B9E
         ButtonStyle     =   3
      End
      Begin VB.Label lblDocumentNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   90
         TabIndex        =   17
         Top             =   960
         Width           =   1665
      End
   End
End
Attribute VB_Name = "frmAddEditCashDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_CashDoc As CCashDoc
Private m_Employees As Collection
Private m_Customers  As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public id As Long
Public DocumentType As CASH_DOC_TYPE

Private FileName As String
Private m_SumUnit As Double
Private m_Employee As CEmployee
Private m_Mr As CMasterRef
Private m_Banks As Collection
Private m_BankBranchs As Collection
Private m_BankAccounts As Collection

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
      Call m_CashDoc.SetFieldValue("CASH_DOC_ID", id)
      m_CashDoc.QueryFlag = 1
'      Call m_CashDoc.SetFieldValue("COMMIT_FLAG", "")
      If Not glbDaily.QueryCashDoc(m_CashDoc, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_CashDoc.PopulateFromRS(1, m_Rs)
      
      uctlDocumentDate.ShowDate = m_CashDoc.GetFieldValue("DOCUMENT_DATE")
      txtDocumentNo.Text = m_CashDoc.GetFieldValue("DOCUMENT_NO")
      uctlBankLookup.MyCombo.ListIndex = IDToListIndex(uctlBankLookup.MyCombo, m_CashDoc.GetFieldValue("BANK_ID"))
      uctlBankBranchLookup.MyCombo.ListIndex = IDToListIndex(uctlBankBranchLookup.MyCombo, m_CashDoc.GetFieldValue("BANK_BRANCH"))
      uctlBankAccountLookup.MyCombo.ListIndex = IDToListIndex(uctlBankAccountLookup.MyCombo, m_CashDoc.GetFieldValue("BANK_ACCOUNT"))
      uctlEmployeeLookup.MyCombo.ListIndex = IDToListIndex(uctlEmployeeLookup.MyCombo, m_CashDoc.GetFieldValue("EMP_ID"))
      uctlCustomerLookup.MyCombo.ListIndex = IDToListIndex(uctlCustomerLookup.MyCombo, m_CashDoc.GetFieldValue("CUSTOMER_ID"))
      
      If (DocumentType = CASH_WITHDRAW) Or (DocumentType = CASH_TRANSFER) Or (DocumentType = CASH_DEPOSIT) Then
         Call glbDaily.CreateCashTransferItems(m_CashDoc)
      End If
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
Private Sub CalculateSumPrice()
'Dim Li As CLotItem
'Dim Ti As CTransferItem
'Dim Sum As Double
'
'   Sum = 0
'   If DocumentType = TRANSFER_DOCTYPE Then
'      For Each Ti In m_CashDoc.TransferItems
'         If Ti.Flag <> "D" Then
'            Sum = Sum + Ti.ImportItem.GetFieldValue("TOTAL_INCLUDE_PRICE")
'         End If
'      Next Ti
'   Else
'      For Each Li In m_CashDoc.CashTranItems
'         If Li.Flag <> "D" Then
'            Sum = Sum + Li.GetFieldValue("TOTAL_INCLUDE_PRICE")
'         End If
'      Next Li
'   End If
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim BD As CBillingDoc

   If Not VerifyTextControl(lblDocumentNo, txtDocumentNo, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblDocumentDate, uctlDocumentDate, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblBank, uctlBankLookup.MyCombo, Not uctlBankLookup.Enabled) Then
      Exit Function
   End If
   If Not VerifyCombo(lblBankBranch, uctlBankBranchLookup.MyCombo, Not uctlBankBranchLookup.Enabled) Then
      Exit Function
   End If
   If Not VerifyCombo(lblBankAccount, uctlBankAccountLookup.MyCombo, Not uctlBankAccountLookup.Enabled) Then
      Exit Function
   End If
   If Not VerifyCombo(lblEmployee, uctlEmployeeLookup.MyCombo, True) Then
      Exit Function
   End If
   If Not (DocumentType = POST_CHEQUE Or DocumentType = WAITING_CHEQUE Or DocumentType = PASSED_CHEQUE) Then
      If Not VerifyCombo(lblCustomer, uctlCustomerLookup.MyCombo, False) Then
         Exit Function
      End If
   End If
   
'   If Not glbDaily.VerifyDrCr(m_CashDoc.JournalItems) Then
'      glbErrorLog.LocalErrorMsg = "������ͧ ഺԵ �е�ͧ��ҡѺ������ͧ �ôԵ"
'      glbErrorLog.ShowUserError
'      Exit Function
'   End If

'   If Not CheckUniqueNs(EXPORT_UNIQUE, txtDocumentNo.Text, ID) Then
'      glbErrorLog.LocalErrorMsg = MapText("�բ�����") & " " & txtDocumentNo.Text & " " & MapText("������к�����")
'      glbErrorLog.ShowUserError
'      Exit Function
'   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_CashDoc.ShowMode = ShowMode
   Call m_CashDoc.SetFieldValue("CASH_DOC_ID", id)
    Call m_CashDoc.SetFieldValue("DOCUMENT_DATE", uctlDocumentDate.ShowDate)
   Call m_CashDoc.SetFieldValue("DOCUMENT_NO", txtDocumentNo.Text)
   Call m_CashDoc.SetFieldValue("DOCUMENT_TYPE", DocumentType)
   Call m_CashDoc.SetFieldValue("BANK_ID", uctlBankLookup.MyCombo.ItemData(Minus2Zero(uctlBankLookup.MyCombo.ListIndex)))
   Call m_CashDoc.SetFieldValue("BANK_BRANCH", uctlBankBranchLookup.MyCombo.ItemData(Minus2Zero(uctlBankBranchLookup.MyCombo.ListIndex)))
   Call m_CashDoc.SetFieldValue("BANK_ACCOUNT", uctlBankAccountLookup.MyCombo.ItemData(Minus2Zero(uctlBankAccountLookup.MyCombo.ListIndex)))
   Call m_CashDoc.SetFieldValue("EMP_ID", uctlEmployeeLookup.MyCombo.ItemData(Minus2Zero(uctlEmployeeLookup.MyCombo.ListIndex)))
   Call m_CashDoc.SetFieldValue("CUSTOMER_ID", uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex)))
      
   Call EnableForm(Me, False)
   If (DocumentType = CASH_WITHDRAW) Or (DocumentType = CASH_TRANSFER) Or (DocumentType = CASH_DEPOSIT) Then
      Call CreateCashTranItems
   End If
   
   Call glbDaily.StartTransaction
   
   If DocumentType = WAITING_CHEQUE Then
      Call CashDocPost2BillingDoc(m_CashDoc, BD, 15000)        ' ����ҧ�ҡ����ͨ���
   End If
   
   If Not glbDaily.AddEditCashDoc(m_CashDoc, IsOK, False, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call glbDaily.RollbackTransaction
      Call EnableForm(Me, True)
      Exit Function
   End If
   
   
   
   If Not IsOK Then
      Call EnableForm(Me, True)
      Call glbDaily.RollbackTransaction
      glbErrorLog.ShowUserError
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

Private Sub cboReason_Click()
   m_HasModify = True
End Sub

Private Sub cboDepartment_Click()
   m_HasModify = True
End Sub

Private Sub cboDepartMent_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
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

Private Sub chkSaleFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub GetTotalPrice()
Dim D As CCashTran
Dim Sum1 As Double
Dim Sum2 As Double

   Sum1 = 0
   For Each D In m_CashDoc.CashTranItems
      If D.Flag <> "D" Then
         If D.GetFieldValue("TX_TYPE") = "E" Then
            Sum1 = Sum1 + D.GetFieldValue("AMOUNT")
         End If
      End If
   Next D
   
   txtOut.Text = FormatNumber(Sum1)
End Sub

Private Sub chkSaleFlag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean
Dim oMenu As cPopupMenu
Dim lMenuChoosen As Long

   If Not cmdAdd.Enabled Then
      Exit Sub
   End If
   
   OKClick = False
   If TabStrip1.SelectedItem.Tag = DocumentType & "-1" Then
      If DocumentType = CASH_WITHDRAW Then
'         frmAddEditCashTran2.DocumentType = DocumentType
'         Set frmAddEditCashTran2.ParentForm = Me
'         Set frmAddEditCashTran2.TempCollection = m_CashDoc.TransferItems
'         frmAddEditCashTran2.ShowMode = SHOW_ADD
'         frmAddEditCashTran2.HeaderText = MapText("����" & "��¡�ö͹�Թ")
'         Load frmAddEditCashTran2
'         frmAddEditCashTran2.Show 1
'
'         OKClick = frmAddEditCashTran2.OKClick
'
'         Unload frmAddEditCashTran2
'         Set frmAddEditCashTran2 = Nothing
'
'         If OKClick Then
'            m_HasModify = True
'            GridEX1.ItemCount = CountItem(m_CashDoc.TransferItems)
'            GridEX1.Rebind
'         End If
'      ElseIf DocumentType = CASH_WHTHDRAW2 Then
'         frmAddEditCashTran3.DocumentType = DocumentType
'         Set frmAddEditCashTran3.ParentForm = Me
'         Set frmAddEditCashTran3.TempCollection = m_CashDoc.CashTranItems
'         frmAddEditCashTran3.ShowMode = SHOW_ADD
'         frmAddEditCashTran3.HeaderText = MapText("����" & "��¡�ö͹�Թ/�͹�Թ")
'         Load frmAddEditCashTran3
'         frmAddEditCashTran3.Show 1
'
'         OKClick = frmAddEditCashTran3.OKClick
'
'         Unload frmAddEditCashTran3
'         Set frmAddEditCashTran3 = Nothing
'
'         If OKClick Then
'            m_HasModify = True
'            GridEX1.ItemCount = CountItem(m_CashDoc.CashTranItems)
'            GridEX1.Rebind
'         End If
'      ElseIf DocumentType = CASH_TRANSFER Then
'         frmAddEditCashTran4.DocumentType = DocumentType
'         Set frmAddEditCashTran4.ParentForm = Me
'         Set frmAddEditCashTran4.TempCollection = m_CashDoc.TransferItems
'         frmAddEditCashTran4.ShowMode = SHOW_ADD
'         frmAddEditCashTran4.HeaderText = MapText("����" & "��¡���͹�Թ�����ҧ�ѭ��")
'         Load frmAddEditCashTran4
'         frmAddEditCashTran4.Show 1
'
'         OKClick = frmAddEditCashTran4.OKClick
'
'         Unload frmAddEditCashTran4
'         Set frmAddEditCashTran4 = Nothing
'
'         If OKClick Then
'            m_HasModify = True
'            GridEX1.ItemCount = CountItem(m_CashDoc.TransferItems)
'            GridEX1.Rebind
'         End If
      ElseIf DocumentType = CASH_DEPOSIT Then
         Set oMenu = New cPopupMenu
         lMenuChoosen = oMenu.Popup("�Թʴ����", "-", "������")
         Set oMenu = Nothing
         
         If lMenuChoosen = 1 Then
            frmAddEditCashTran5.DocumentType = DocumentType
            Set frmAddEditCashTran5.ParentForm = Me
            Set frmAddEditCashTran5.TempCollection = m_CashDoc.TransferItems
            frmAddEditCashTran5.ShowMode = SHOW_ADD
            frmAddEditCashTran5.HeaderText = MapText("����" & "��¡�ùӽҡ�Թ")
            Load frmAddEditCashTran5
            frmAddEditCashTran5.Show 1
   
            OKClick = frmAddEditCashTran5.OKClick
   
            Unload frmAddEditCashTran5
            Set frmAddEditCashTran5 = Nothing
         ElseIf lMenuChoosen = 3 Then
            
            If Not VerifyCombo(lblCustomer, uctlCustomerLookup.MyCombo, False) Then
               Exit Sub
            End If
            
            frmAddChequeItem.DocumentType = DocumentType
            frmAddChequeItem.ApArID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
            Set frmAddChequeItem.TempCollection = m_CashDoc.TransferItems
            frmAddChequeItem.ShowMode = SHOW_ADD
            frmAddChequeItem.HeaderText = MapText("����" & "��¡�ùӽҡ��")
            Load frmAddChequeItem
            frmAddChequeItem.Show 1
   
            OKClick = frmAddChequeItem.OKClick
   
            Unload frmAddChequeItem
            Set frmAddChequeItem = Nothing
         End If
         
         If OKClick Then
            m_HasModify = True
            GridEX1.ItemCount = CountItem(m_CashDoc.TransferItems)
            GridEX1.Rebind
         End If
      ElseIf DocumentType = POST_CHEQUE Then
         If Not VerifyCombo(lblBankAccount, uctlBankAccountLookup.MyCombo, False) Then
            Exit Sub
         End If
         
         frmAddChequeItemEx.DocumentType = DocumentType
         frmAddChequeItemEx.PostType = POST_CLEAR
         frmAddChequeItemEx.AccountID = uctlBankAccountLookup.MyCombo.ItemData(Minus2Zero(uctlBankAccountLookup.MyCombo.ListIndex))
         Set frmAddChequeItemEx.TempCollection = m_CashDoc.PostItems
         frmAddChequeItemEx.ShowMode = SHOW_ADD
         frmAddChequeItemEx.HeaderText = MapText("����" & "��¡���礷�����Թ����")
         Load frmAddChequeItemEx
         frmAddChequeItemEx.Show 1

         OKClick = frmAddChequeItemEx.OKClick

         Unload frmAddChequeItemEx
         Set frmAddChequeItemEx = Nothing
         
      If OKClick Then
         m_HasModify = True
         GridEX1.ItemCount = CountItem(m_CashDoc.PostItems)
         GridEX1.Rebind
      End If
   ElseIf DocumentType = WAITING_CHEQUE Then
      frmAddChequeItemEx.DocumentType = DocumentType
      frmAddChequeItemEx.PostType = WAITING_CLEAR
      Set frmAddChequeItemEx.TempCollection = m_CashDoc.PostItems
      frmAddChequeItemEx.ShowMode = SHOW_ADD
      frmAddChequeItemEx.HeaderText = MapText("����" & "��¡���礷�����Թ����")
      Load frmAddChequeItemEx
      frmAddChequeItemEx.Show 1

      OKClick = frmAddChequeItemEx.OKClick

      Unload frmAddChequeItemEx
      Set frmAddChequeItemEx = Nothing
      
      If OKClick Then
         m_HasModify = True
         GridEX1.ItemCount = CountItem(m_CashDoc.PostItems)
         GridEX1.Rebind
      End If
   ElseIf DocumentType = PASSED_CHEQUE Then
      frmAddChequeItemEx.DocumentType = DocumentType
      frmAddChequeItemEx.PostType = PASSED_CLEAR
      Set frmAddChequeItemEx.TempCollection = m_CashDoc.PostItems
      frmAddChequeItemEx.ShowMode = SHOW_ADD
      frmAddChequeItemEx.HeaderText = MapText("����" & "��¡���礷�����Թ����")
      Load frmAddChequeItemEx
      frmAddChequeItemEx.Show 1

      OKClick = frmAddChequeItemEx.OKClick

      Unload frmAddChequeItemEx
      Set frmAddChequeItemEx = Nothing
      
      If OKClick Then
         m_HasModify = True
         GridEX1.ItemCount = CountItem(m_CashDoc.PostItems)
         GridEX1.Rebind
      End If


'      ElseIf DocumentType = CASH_DEPOSIT2 Then
'         frmAddEditCashTran3.DocumentType = DocumentType
'         Set frmAddEditCashTran3.ParentForm = Me
'         Set frmAddEditCashTran3.TempCollection = m_CashDoc.CashTranItems
'         frmAddEditCashTran3.ShowMode = SHOW_ADD
'         frmAddEditCashTran3.HeaderText = MapText("����" & "��¡�ýҡ�Թ/�͹�Թ")
'         Load frmAddEditCashTran3
'         frmAddEditCashTran3.Show 1
'
'         OKClick = frmAddEditCashTran3.OKClick
'
'         Unload frmAddEditCashTran3
'         Set frmAddEditCashTran3 = Nothing
'
'         If OKClick Then
'            m_HasModify = True
'            GridEX1.ItemCount = CountItem(m_CashDoc.CashTranItems)
'            GridEX1.Rebind
'         End If
'      ElseIf DocumentType = CASH_PITTYCASH Then
'         frmAddEditCashTran6.DocumentType = DocumentType
'         Set frmAddEditCashTran6.ParentForm = Me
'         Set frmAddEditCashTran6.TempCollection = m_CashDoc.CashTranItems
'         frmAddEditCashTran6.ShowMode = SHOW_ADD
'         frmAddEditCashTran6.HeaderText = MapText("����" & "��¡���ԡ�Թʴ����")
'         Load frmAddEditCashTran6
'         frmAddEditCashTran6.Show 1
'
'         OKClick = frmAddEditCashTran6.OKClick
'
'         Unload frmAddEditCashTran6
'         Set frmAddEditCashTran6 = Nothing
'
'         If OKClick Then
'            m_HasModify = True
'            GridEX1.ItemCount = CountItem(m_CashDoc.CashTranItems)
'            GridEX1.Rebind
'         End If
      End If
'   ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-2" Then
'      Set frmAddEditJournalItem.ParentForm = Me
'      frmAddEditJournalItem.HeaderText = "������¡����ش����ѹ"
'      frmAddEditJournalItem.ShowMode = SHOW_ADD
'      Set frmAddEditJournalItem.TempCollection = m_CashDoc.JournalItems
'      Load frmAddEditJournalItem
'      frmAddEditJournalItem.Show 1
'
'      OKClick = frmAddEditJournalItem.OKClick
'
'      Unload frmAddEditJournalItem
'      Set frmAddEditJournalItem = Nothing
'
'      If OKClick Then
'         m_HasModify = True
'         GridEX1.ItemCount = CountItem(m_CashDoc.JournalItems)
'         GridEX1.Rebind
'      End If
   End If
   
End Sub

Private Sub cmdAuto_Click()
Dim No As String
'
'   If Trim(txtDocumentNo.Text) = "" Then
'      Call glbDatabaseMngr.GenerateNumber(EXPORT_NUMBER, No, glbErrorLog)
'      txtDocumentNo.Text = No
'   End If
End Sub
Private Sub CreateCashTranItems()
Dim Ti As CCashTransferItem
Dim Ei As CCashTran
Dim II As CCashTran
   
   Set m_CashDoc.CashTranItems = Nothing
   Set m_CashDoc.CashTranItems = New Collection
   
   For Each Ti In m_CashDoc.TransferItems
      Set Ei = Ti.ExportItem
      Set II = Ti.ImportItem
      
      Ei.Flag = Ti.Flag
      II.Flag = Ti.Flag
      
      Call m_CashDoc.CashTranItems.add(Ei)
      Call m_CashDoc.CashTranItems.add(II)
   Next Ti
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
   
   If TabStrip1.SelectedItem.Tag = DocumentType & "-1" Then
      If DocumentType = CASH_TRANSFER Then
         If ID1 <= 0 Then
            m_CashDoc.TransferItems.Remove (ID2)
         Else
            m_CashDoc.TransferItems.Item(ID2).Flag = "D"
         End If
      ElseIf DocumentType = CASH_WHTHDRAW2 Then
         If ID1 <= 0 Then
            m_CashDoc.CashTranItems.Remove (ID2)
         Else
            m_CashDoc.CashTranItems.Item(ID2).Flag = "D"
         End If
      ElseIf DocumentType = CASH_DEPOSIT Then
         If ID1 <= 0 Then
            m_CashDoc.TransferItems.Remove (ID2)
         Else
            m_CashDoc.TransferItems.Item(ID2).Flag = "D"
         End If
      ElseIf DocumentType = CASH_DEPOSIT2 Then
         If ID1 <= 0 Then
            m_CashDoc.CashTranItems.Remove (ID2)
         Else
            m_CashDoc.CashTranItems.Item(ID2).Flag = "D"
         End If
      ElseIf DocumentType = CASH_PITTYCASH Then
         If ID1 <= 0 Then
            m_CashDoc.CashTranItems.Remove (ID2)
         Else
            m_CashDoc.CashTranItems.Item(ID2).Flag = "D"
         End If
      ElseIf DocumentType = POST_CHEQUE Or DocumentType = WAITING_CHEQUE Or DocumentType = PASSED_CHEQUE Then
         If ID1 <= 0 Then
            m_CashDoc.PostItems.Remove (ID2)
         Else
            m_CashDoc.PostItems.Item(ID2).Flag = "D"
         End If
      End If
      
      Call RefreshGrid(DocumentType, True)
   ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-2" Then
      If ID1 <= 0 Then
         m_CashDoc.JournalItems.Remove (ID2)
      Else
         m_CashDoc.JournalItems.Item(ID2).Flag = "D"
      End If
      Call RefreshGrid(DocumentType, True)
   End If
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim id As Long
Dim OKClick As Boolean
Dim PaymentType As Long
         
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   id = Val(GridEX1.Value(2))
   OKClick = False
   
   If TabStrip1.SelectedItem.Tag = DocumentType & "-1" Then
      If (DocumentType = CASH_WITHDRAW) Then
'         frmAddEditCashTran2.DocumentType = DocumentType
'         Set frmAddEditCashTran2.ParentForm = Me
'         frmAddEditCashTran2.ID = ID
'         Set frmAddEditCashTran2.TempCollection = m_CashDoc.TransferItems
'         frmAddEditCashTran2.HeaderText = MapText("���" & "��¡�ö͹")
'         frmAddEditCashTran2.ShowMode = SHOW_EDIT
'         Load frmAddEditCashTran2
'         frmAddEditCashTran2.Show 1
'
'         OKClick = frmAddEditCashTran2.OKClick
'
'         Unload frmAddEditCashTran2
'         Set frmAddEditCashTran2 = Nothing
'
'         If OKClick Then
'            m_HasModify = True
'            GridEX1.ItemCount = CountItem(m_CashDoc.TransferItems)
'            GridEX1.Rebind
'         End If
'      ElseIf (DocumentType = CASH_WHTHDRAW2) Then
'         frmAddEditCashTran3.DocumentType = DocumentType
'         Set frmAddEditCashTran3.ParentForm = Me
'         frmAddEditCashTran3.ID = ID
'         Set frmAddEditCashTran3.TempCollection = m_CashDoc.CashTranItems
'         frmAddEditCashTran3.HeaderText = MapText("���" & "��¡�ö͹�Թ/�͹�Թ")
'         frmAddEditCashTran3.ShowMode = SHOW_EDIT
'         Load frmAddEditCashTran3
'         frmAddEditCashTran3.Show 1
'
'         OKClick = frmAddEditCashTran3.OKClick
'
'         Unload frmAddEditCashTran3
'         Set frmAddEditCashTran3 = Nothing
'
'         If OKClick Then
'            m_HasModify = True
'            GridEX1.ItemCount = CountItem(m_CashDoc.CashTranItems)
'            GridEX1.Rebind
'         End If
'      ElseIf (DocumentType = CASH_TRANSFER) Then
'         frmAddEditCashTran4.DocumentType = DocumentType
'         Set frmAddEditCashTran4.ParentForm = Me
'         frmAddEditCashTran4.ID = ID
'         Set frmAddEditCashTran4.TempCollection = m_CashDoc.TransferItems
'         frmAddEditCashTran4.HeaderText = MapText("���" & "��¡���͹�����ҧ�ѭ��")
'         frmAddEditCashTran4.ShowMode = SHOW_EDIT
'         Load frmAddEditCashTran4
'         frmAddEditCashTran4.Show 1
'
'         OKClick = frmAddEditCashTran4.OKClick
'
'         Unload frmAddEditCashTran4
'         Set frmAddEditCashTran4 = Nothing
'
'         If OKClick Then
'            m_HasModify = True
'            GridEX1.ItemCount = CountItem(m_CashDoc.TransferItems)
'            GridEX1.Rebind
'         End If
      ElseIf (DocumentType = POST_CHEQUE) Or (DocumentType = WAITING_CHEQUE) Or (DocumentType = PASSED_CHEQUE) Then
         glbErrorLog.LocalErrorMsg = "�������ö�����Դ��������س�ź�������ҧ����"
         glbErrorLog.ShowUserError
         Exit Sub
      ElseIf (DocumentType = CASH_DEPOSIT) Then
         PaymentType = GridEX1.Value(8)
         If PaymentType = 3 Then  '��
            glbErrorLog.LocalErrorMsg = "��¡�ùӽҡ���������ö�����Դ�������"
            glbErrorLog.ShowUserError
            Exit Sub
         End If
         
         frmAddEditCashTran5.DocumentType = DocumentType
         Set frmAddEditCashTran5.ParentForm = Me
         frmAddEditCashTran5.id = id
         Set frmAddEditCashTran5.TempCollection = m_CashDoc.TransferItems
         frmAddEditCashTran5.HeaderText = MapText("���" & "��¡�ùӽҡ�Թ")
         frmAddEditCashTran5.ShowMode = SHOW_EDIT
         Load frmAddEditCashTran5
         frmAddEditCashTran5.Show 1
   
         OKClick = frmAddEditCashTran5.OKClick
   
         Unload frmAddEditCashTran5
         Set frmAddEditCashTran5 = Nothing
      
         If OKClick Then
            m_HasModify = True
            GridEX1.ItemCount = CountItem(m_CashDoc.TransferItems)
            GridEX1.Rebind
         End If
'      ElseIf (DocumentType = CASH_DEPOSIT2) Then
'         frmAddEditCashTran3.DocumentType = DocumentType
'         Set frmAddEditCashTran3.ParentForm = Me
'         frmAddEditCashTran3.ID = ID
'         Set frmAddEditCashTran3.TempCollection = m_CashDoc.CashTranItems
'         frmAddEditCashTran3.HeaderText = MapText("���" & "��¡�ýҡ�Թ/�͹�Թ")
'         frmAddEditCashTran3.ShowMode = SHOW_EDIT
'         Load frmAddEditCashTran3
'         frmAddEditCashTran3.Show 1
'
'         OKClick = frmAddEditCashTran3.OKClick
'
'         Unload frmAddEditCashTran3
'         Set frmAddEditCashTran3 = Nothing
'
'         If OKClick Then
'            m_HasModify = True
'            GridEX1.ItemCount = CountItem(m_CashDoc.CashTranItems)
'            GridEX1.Rebind
'         End If
'      ElseIf (DocumentType = CASH_PITTYCASH) Then
'         frmAddEditCashTran6.DocumentType = DocumentType
'         Set frmAddEditCashTran6.ParentForm = Me
'         frmAddEditCashTran6.ID = ID
'         Set frmAddEditCashTran6.TempCollection = m_CashDoc.CashTranItems
'         frmAddEditCashTran6.HeaderText = MapText("���" & "��¡���ԡ�Թʴ����")
'         frmAddEditCashTran6.ShowMode = SHOW_EDIT
'         Load frmAddEditCashTran6
'         frmAddEditCashTran6.Show 1
'
'         OKClick = frmAddEditCashTran6.OKClick
'
'         Unload frmAddEditCashTran6
'         Set frmAddEditCashTran6 = Nothing
'
'         If OKClick Then
'            m_HasModify = True
'            GridEX1.ItemCount = CountItem(m_CashDoc.CashTranItems)
'            GridEX1.Rebind
'         End If
      End If
   ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-2" Then
'      Set frmAddEditJournalItem.ParentForm = Me
'      frmAddEditJournalItem.ID = ID
'      frmAddEditJournalItem.HeaderText = "�����¡����ش����ѹ"
'      frmAddEditJournalItem.ShowMode = SHOW_EDIT
'      Set frmAddEditJournalItem.TempCollection = m_CashDoc.JournalItems
'      Load frmAddEditJournalItem
'      frmAddEditJournalItem.Show 1
'
'      OKClick = frmAddEditJournalItem.OKClick
'
'      Unload frmAddEditJournalItem
'      Set frmAddEditJournalItem = Nothing
'
'      If OKClick Then
'         m_HasModify = True
'         GridEX1.ItemCount = CountItem(m_CashDoc.CashTranItems)
'         GridEX1.Rebind
'      End If
   End If
End Sub
Private Sub cmdOK_Click()
Dim oMenu As cPopupMenu
Dim lMenuChosen  As Long

   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("�ѹ�֡", "-", "�ѹ�֡����͡�ҡ˹�Ҩ�")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
   If lMenuChosen = 1 Then
      If Not SaveData Then
         Exit Sub
      End If
      
      ShowMode = SHOW_EDIT
      id = m_CashDoc.GetFieldValue("CASH_DOC_ID")
      m_CashDoc.QueryFlag = 1
      QueryData (True)
      m_HasModify = False
   ElseIf lMenuChosen = 3 Then
      If Not SaveData Then
         Exit Sub
      End If
      
      OKClick = True
      Unload Me
   End If
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

Private Sub cmdPrint_Click()
'Dim Report As CReportInterface
'Dim oMenu As CPopupMenu
'Dim lMenuChosen As Long
'Dim ReportKey As String
'Dim ReportFlag As Boolean
'Dim Rc As CReportConfig
'Dim iCount As Long
'Dim EditMode As SHOW_MODE_TYPE
'Dim ReportMode As Long
'
'   If m_HasModify Or ((Not m_HasModify) And (ShowMode = SHOW_ADD)) Then
'      glbErrorLog.LocalErrorMsg = "��سҷӡ�úѹ�֡������������º���¡�͹"
'      glbErrorLog.ShowUserError
'      Exit Sub
'   End If
'
'   Set oMenu = New CPopupMenu
'   lMenuChosen = oMenu.Popup("��ԡ�Թ���/�ѵ�شԺ", "��Ѻ���˹�ҡ�д��")
'   If lMenuChosen = 0 Then
'      Exit Sub
'   End If
'
'   Call EnableForm(Me, False)
'
'   If lMenuChosen = 1 Then
'      ReportKey = "CReportInvDoc002"
'
'      Set Report = New CReportInvDoc002
'      ReportFlag = True
'   ElseIf lMenuChosen = 2 Then
'      ReportKey = "CReportInvDoc002"
'
'      Set Rc = New CReportConfig
'      Rc.REPORT_KEY = ReportKey
'      Call Rc.QueryData(m_Rs, iCount)
'      HeaderText = MapText("��ԡ�Թ���/�ѵ�شԺ")
'      If Not m_Rs.EOF Then
'         Call Rc.PopulateFromRs(1, m_Rs)
'         EditMode = SHOW_EDIT
'      Else
'         EditMode = SHOW_ADD
'      End If
'   End If
'
'   If Not Report Is Nothing Then
'      Call Report.AddParam(m_CashDoc.CASH_DOC_ID, "CASH_DOC_ID")
'      Call Report.AddParam(ReportKey, "REPORT_KEY")
'   End If
'
'   If ReportFlag Then
'      Set frmReport.ReportObject = Report
'      frmReport.HeaderText = pnlHeader.Caption
'      Load frmReport
'      frmReport.Show 1
'
'      Unload frmReport
'      Set frmReport = Nothing
'      Set Report = Nothing
'   Else
'      frmReportConfig.ReportMode = 1
'      frmReportConfig.ShowMode = EditMode
'      frmReportConfig.ID = Rc.REPORT_CONFIG_ID
'      frmReportConfig.ReportKey = ReportKey
'      frmReportConfig.HeaderText = HeaderText
'      Load frmReportConfig
'      frmReportConfig.Show 1
'
'      Unload frmReportConfig
'      Set frmReportConfig = Nothing
'   End If
'
'   Call EnableForm(Me, True)
End Sub
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      
      Call LoadCustomer(uctlCustomerLookup.MyCombo, m_Customers)
      Set uctlCustomerLookup.MyCollection = m_Customers
      
      Call LoadBank(uctlBankLookup.MyCombo, m_Banks)
      Set uctlBankLookup.MyCollection = m_Banks
      
      Call LoadBankBranch(uctlBankBranchLookup.MyCombo, m_BankBranchs)
      Set uctlBankBranchLookup.MyCollection = m_BankBranchs
            
      Call LoadMaster(uctlBankAccountLookup.MyCombo, m_BankAccounts, BANK_ACCOUNT)
      Set uctlBankAccountLookup.MyCollection = m_BankAccounts
            
      Call LoadEmployee(uctlEmployeeLookup.MyCombo, m_Employees)
      Set uctlEmployeeLookup.MyCollection = m_Employees
            
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_CashDoc.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         uctlDocumentDate.ShowDate = Now
         m_CashDoc.QueryFlag = 0
         Call QueryData(False)
      End If
      
      Call EnableForm(Me, True)
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_CashDoc = Nothing
   Set m_Employees = Nothing
   Set m_Employee = Nothing
   Set m_Banks = Nothing
   Set m_BankBranchs = Nothing
   Set m_BankAccounts = Nothing
   Set m_Employee = Nothing
   Set m_Customers = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid1(Ind As CASH_DOC_TYPE)
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

   If TabStrip1.SelectedItem.Tag = DocumentType & "-2" Then
      Set Col = GridEX1.Columns.add '3
      Col.Width = 1965
      Col.Caption = MapText("���ʺѭ��")
   
      Set Col = GridEX1.Columns.add '4
      Col.Width = 5100
      Col.Caption = MapText("��������´")
   
      Set Col = GridEX1.Columns.add '5
      Col.Width = 2025
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ഺԵ")
   
      Set Col = GridEX1.Columns.add '6
      Col.Width = 2160
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("�ôԵ")
   ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-CLR" Then
      Set Col = GridEX1.Columns.add '3
      Col.Width = 5175
      Col.Caption = MapText("��������´")
   
      Set Col = GridEX1.Columns.add '4
      Col.Width = 1980
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("�Ҥҵ��˹���")
   
      Set Col = GridEX1.Columns.add '5
      Col.Width = 2025
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("�ӹǹ")
   
      Set Col = GridEX1.Columns.add '6
      Col.Width = 2160
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("�ӹǹ�Թ")
   ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-1" Then
      If (Ind = CASH_WITHDRAW) Or (Ind = CASH_WHTHDRAW2) Or (Ind = CASH_DEPOSIT) Or (Ind = CASH_DEPOSIT2) Or (Ind = CASH_PITTYCASH) Then
         Set Col = GridEX1.Columns.add '3
         Col.Width = 2400
         Col.Caption = MapText("������")
      
         Set Col = GridEX1.Columns.add '4
         Col.Width = 2415
         Col.TextAlignment = jgexAlignRight
         Col.Caption = MapText("�ӹǹ�Թ")
         
         Set Col = GridEX1.Columns.add '5
         Col.Width = 2745
         Col.Caption = MapText("�Ţ�����")
         
         Set Col = GridEX1.Columns.add '6
         Col.Width = 2820
         Col.Caption = MapText("�ѹ�����")
      
         Set Col = GridEX1.Columns.add '7
         Col.Width = 3570
         Col.Caption = MapText("�ѹ������Թ")
      
         Set Col = GridEX1.Columns.add '8
         Col.Width = 0
         Col.Visible = False
         Col.Caption = MapText("PAYMENT_TYPE")
      ElseIf (Ind = CASH_TRANSFER) Then
         Set Col = GridEX1.Columns.add '3
         Col.Width = 2400
         Col.Caption = MapText("������")
      
         Set Col = GridEX1.Columns.add '4
         Col.Width = 2415
         Col.TextAlignment = jgexAlignRight
         Col.Caption = MapText("�ӹǹ�Թ")
         
         Set Col = GridEX1.Columns.add '5
         Col.Width = 2745
         Col.Caption = MapText("�Ţ���ѭ��")
         
         Set Col = GridEX1.Columns.add '6
         Col.Width = 2820
         Col.Caption = MapText("��Ҥ��")
      
         Set Col = GridEX1.Columns.add '7
         Col.Width = 3570
         Col.Caption = MapText("�ҢҸ�Ҥ��")
      ElseIf (Ind = POST_CHEQUE) Or (Ind = WAITING_CHEQUE) Or (Ind = PASSED_CHEQUE) Then
      
         Set Col = GridEX1.Columns.add '3
         Col.Width = 1500
         Col.Caption = MapText("�Ţ�����")
         
         Set Col = GridEX1.Columns.add '5
         Col.Width = 1500
         Col.Caption = MapText("�ѹ�����")
         
         Set Col = GridEX1.Columns.add '4
         Col.Width = 1500
         Col.TextAlignment = jgexAlignRight
         Col.Caption = MapText("�ӹǹ�Թ")
         
         Set Col = GridEX1.Columns.add '5
         Col.Width = 2820
         Col.Caption = MapText("��Ҥ��")
      
         Set Col = GridEX1.Columns.add '6
         Col.Width = 3570
         Col.Caption = MapText("�ҢҸ�Ҥ��")
      End If
   End If
End Sub

Private Sub InitFormLayout()

   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)

   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblDocumentNo, MapText("�Ţ����͡���"))
   Call InitNormalLabel(lblDocumentDate, MapText("�ѹ����͡���"))
   Call InitNormalLabel(Label4, MapText("�ҷ"))
   Call InitNormalLabel(lblBank, MapText("��Ҥ��"))
   Call InitNormalLabel(lblBankBranch, MapText("�ҢҸ�Ҥ��"))
   Call InitNormalLabel(lblBankAccount, MapText("�Ţ���ѭ��"))
   Call InitNormalLabel(lblCustomer, MapText("�����١���"))
   
   If DocumentType = CASH_PITTYCASH Then
      Call InitNormalLabel(lblEmployee, MapText("����ԡ"))
      uctlBankLookup.Enabled = False
      uctlBankBranchLookup.Enabled = False
      uctlBankAccountLookup.Enabled = False
      lblIn.Visible = True
      lblOut.Visible = True
      lblLeft.Visible = True
      txtIn.Visible = True
      txtOut.Visible = True
      txtLeft.Visible = True
      
      Call InitNormalLabel(lblIn, MapText("������ԧ"))
      Call InitNormalLabel(lblOut, MapText("����ԡ"))
      Call InitNormalLabel(lblLeft, MapText("������ͤ׹"))
   ElseIf DocumentType = POST_CHEQUE Then
      Call InitNormalLabel(lblEmployee, MapText("����Ǩ�ͺ"))
      uctlCustomerLookup.Enabled = False
      lblIn.Visible = False
      lblOut.Visible = False
      lblLeft.Visible = False
      txtIn.Visible = False
      txtOut.Visible = False
      txtLeft.Visible = False
      Label4.Visible = False
   ElseIf DocumentType = WAITING_CHEQUE Or DocumentType = PASSED_CHEQUE Then
      Call InitNormalLabel(lblEmployee, MapText("��������"))
      uctlCustomerLookup.Enabled = False
      uctlBankLookup.Enabled = False
      uctlBankBranchLookup.Enabled = False
      uctlBankAccountLookup.Enabled = False
      
      lblIn.Visible = False
      lblOut.Visible = False
      lblLeft.Visible = False
      txtIn.Visible = False
      txtOut.Visible = False
      txtLeft.Visible = False
      Label4.Visible = False
   Else
      lblIn.Visible = False
      lblOut.Visible = False
      lblLeft.Visible = False
      txtIn.Visible = False
      txtOut.Visible = False
      txtLeft.Visible = False
      
      Call InitNormalLabel(lblEmployee, MapText("��ѡ�ҹ"))
   End If
   
   Call txtIn.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtIn.Enabled = False
   Call txtOut.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtOut.Enabled = False
   Call txtLeft.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtLeft.Enabled = False
   
   Call txtDocumentNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
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
   cmdAuto.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
   Call InitMainButton(cmdAdd, MapText("���� (F7)"))
   Call InitMainButton(cmdEdit, MapText("��� (F3)"))
   Call InitMainButton(cmdDelete, MapText("ź (F6)"))
   Call InitMainButton(cmdPrint, MapText("�����"))
   Call InitMainButton(cmdAuto, MapText("A"))

   Call InitCheckBox(chkCommit, MapText("�������"))
   
   Call InitGrid1(DocumentType)
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.NAME = GLB_FONT
   TabStrip1.Font.Size = 16
   
   Dim T As Object
   TabStrip1.Tabs.Clear
   
   If DocumentType = CASH_DEPOSIT Then
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText("��¡�ùӽҡ")
      T.Tag = DocumentType & "-1"
   ElseIf DocumentType = CASH_PITTYCASH Then
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText("��¡���ԡ")
      T.Tag = DocumentType & "-1"
   
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText("��¡��������")
      T.Tag = DocumentType & "-CLR"
   ElseIf DocumentType = CASH_TRANSFER Then
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText("��¡���͹�Թ")
      T.Tag = DocumentType & "-1"
   ElseIf DocumentType = CASH_WITHDRAW Then
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText("��¡�ö͹�Թ")
      T.Tag = DocumentType & "-1"
   ElseIf DocumentType = CASH_WHTHDRAW2 Then
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText("��¡�ö͹�Թ/�͹�Թ")
      T.Tag = DocumentType & "-1"
   ElseIf DocumentType = CASH_DEPOSIT2 Then
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText("��¡�ýҡ�Թ/�͹�Թ")
      T.Tag = DocumentType & "-1"
   ElseIf DocumentType = POST_CHEQUE Then
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText("��¡���礷�����Թ������")
      T.Tag = DocumentType & "-1"
   ElseIf DocumentType = WAITING_CHEQUE Then
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText("��¡�����ͨ���")
      T.Tag = DocumentType & "-1"
   ElseIf DocumentType = PASSED_CHEQUE Then
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText("��¡�÷���ҹ����")
      T.Tag = DocumentType & "-1"
   End If

'   Set T = TabStrip1.Tabs.add()
'   T.Caption = MapText("��¡����ش����ѹ")
'   T.Tag = DocumentType & "-2"
End Sub
Private Function Doctype2Text(TempID As INVENTORY_DOCTYPE) As String
   If TempID = IMPORT_DOCTYPE Then
      Doctype2Text = "��¡�ù����"
   ElseIf TempID = EXPORT_DOCTYPE Then
      Doctype2Text = "��¡���ԡ����"
   ElseIf TempID = TRANSFER_DOCTYPE Then
      Doctype2Text = "��¡���͹ʵ�ͤ"
   ElseIf TempID = ADJUST_DOCTYPE Then
      Doctype2Text = "��¡�û�Ѻ�ʹʵ�ͤ"
   End If
End Function

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
   Set m_CashDoc = New CCashDoc
   Set m_Employee = New CEmployee
   Set m_Employees = New Collection
   Set m_Customers = New Collection
   Set m_Mr = New CMasterRef
   Set m_Banks = New Collection
   Set m_BankBranchs = New Collection
   Set m_BankAccounts = New Collection
   Set m_Employee = New CEmployee
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
Dim Tr As CCashTransferItem
Dim Ct1 As CCashTran
Dim Pos As CCashDocPost

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Tag = DocumentType & "-1" Then
      If m_CashDoc.CashTranItems Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      If DocumentType = CASH_WITHDRAW Then
         If m_CashDoc.TransferItems.Count <= 0 Then
            Exit Sub
         End If
         Set Tr = GetItem(m_CashDoc.TransferItems, RowIndex, RealIndex)
         If Tr Is Nothing Then
            Exit Sub
         End If
   
         Values(1) = Tr.ImportItem.GetFieldValue("CASH_TRAN_ID")
         Values(2) = RealIndex
         Values(3) = Tr.ExportItem.GetFieldValue("PAYMENT_TYPE_NAME")
         Values(4) = FormatNumber(Tr.ImportItem.GetFieldValue("AMOUNT"))
         Values(5) = Tr.ExportItem.Cheque.GetFieldValue("CHEQUE_NO")
         Values(6) = DateToStringExtEx2(Tr.ExportItem.Cheque.GetFieldValue("CHEQUE_DATE"))
         Values(7) = DateToStringExtEx2(Tr.ExportItem.Cheque.GetFieldValue("EFFECTIVE_DATE"))
      ElseIf DocumentType = CASH_WHTHDRAW2 Then
         If m_CashDoc.CashTranItems.Count <= 0 Then
            Exit Sub
         End If
         Set Ct1 = GetItem(m_CashDoc.CashTranItems, RowIndex, RealIndex)
         If Ct1 Is Nothing Then
            Exit Sub
         End If
   
         Values(1) = Ct1.GetFieldValue("CASH_TRAN_ID")
         Values(2) = RealIndex
         Values(3) = Ct1.GetFieldValue("PAYMENT_TYPE_NAME")
         Values(4) = FormatNumber(Ct1.GetFieldValue("AMOUNT"))
         Values(5) = Ct1.Cheque.GetFieldValue("CHEQUE_NO")
         Values(6) = DateToStringExtEx2(Ct1.Cheque.GetFieldValue("CHEQUE_DATE"))
         Values(7) = DateToStringExtEx2(Ct1.Cheque.GetFieldValue("EFFECTIVE_DATE"))
      ElseIf DocumentType = POST_CHEQUE Then
         If m_CashDoc.PostItems.Count <= 0 Then
            Exit Sub
         End If
         Set Pos = GetItem(m_CashDoc.PostItems, RowIndex, RealIndex)
         If Pos Is Nothing Then
            Exit Sub
         End If
   
         Values(1) = Pos.GetFieldValue("CASH_DOC_POST_ID")
         Values(2) = RealIndex
         Values(3) = Pos.GetFieldValue("CHEQUE_NO")
         Values(4) = DateToStringExtEx2(Pos.GetFieldValue("CHEQUE_DATE"))
         Values(5) = FormatNumber(Pos.GetFieldValue("CHEQUE_AMOUNT"))
         Values(6) = Pos.GetFieldValue("BANK_NAME")
         Values(7) = Pos.GetFieldValue("BRANCH_NAME")
      ElseIf DocumentType = WAITING_CHEQUE Then
         If m_CashDoc.PostItems.Count <= 0 Then
            Exit Sub
         End If
         Set Pos = GetItem(m_CashDoc.PostItems, RowIndex, RealIndex)
         If Pos Is Nothing Then
            Exit Sub
         End If
   
         Values(1) = Pos.GetFieldValue("CASH_DOC_POST_ID")
         Values(2) = RealIndex
         Values(3) = Pos.GetFieldValue("CHEQUE_NO")
         Values(4) = DateToStringExtEx2(Pos.GetFieldValue("CHEQUE_DATE"))
         Values(5) = FormatNumber(Pos.GetFieldValue("CHEQUE_AMOUNT"))
         Values(6) = Pos.GetFieldValue("BANK_NAME")
         Values(7) = Pos.GetFieldValue("BRANCH_NAME")
      ElseIf DocumentType = PASSED_CHEQUE Then
         If m_CashDoc.PostItems.Count <= 0 Then
            Exit Sub
         End If
         Set Pos = GetItem(m_CashDoc.PostItems, RowIndex, RealIndex)
         If Pos Is Nothing Then
            Exit Sub
         End If
   
         Values(1) = Pos.GetFieldValue("CASH_DOC_POST_ID")
         Values(2) = RealIndex
         Values(3) = Pos.GetFieldValue("CHEQUE_NO")
         Values(4) = DateToStringExtEx2(Pos.GetFieldValue("CHEQUE_DATE"))
         Values(5) = FormatNumber(Pos.GetFieldValue("CHEQUE_AMOUNT"))
         Values(6) = Pos.GetFieldValue("BANK_NAME")
         Values(7) = Pos.GetFieldValue("BRANCH_NAME")
      ElseIf DocumentType = CASH_TRANSFER Then
         If m_CashDoc.TransferItems.Count <= 0 Then
            Exit Sub
         End If
         Set Tr = GetItem(m_CashDoc.TransferItems, RowIndex, RealIndex)
         If Tr Is Nothing Then
            Exit Sub
         End If
   
         Values(1) = Tr.ImportItem.GetFieldValue("CASH_TRAN_ID")
         Values(2) = RealIndex
         Values(3) = Tr.ImportItem.GetFieldValue("PAYMENT_TYPE_NAME")
         Values(4) = FormatNumber(Tr.ImportItem.GetFieldValue("AMOUNT"))
         Values(5) = Tr.ImportItem.GetFieldValue("ACCOUNT_NAME")
         Values(6) = Tr.ImportItem.GetFieldValue("BANK_NAME")
         Values(7) = Tr.ImportItem.GetFieldValue("BRANCH_NAME")
      ElseIf DocumentType = CASH_DEPOSIT Then
         If m_CashDoc.TransferItems.Count <= 0 Then
            Exit Sub
         End If
         Set Tr = GetItem(m_CashDoc.TransferItems, RowIndex, RealIndex)
         If Tr Is Nothing Then
            Exit Sub
         End If
   
         Values(1) = Tr.ImportItem.GetFieldValue("CASH_TRAN_ID")
         Values(2) = RealIndex
         Values(3) = Tr.ExportItem.GetFieldValue("PAYMENT_TYPE_NAME")
         Values(4) = FormatNumber(Tr.ImportItem.GetFieldValue("AMOUNT"))
         Values(5) = Tr.ExportItem.Cheque.GetFieldValue("CHEQUE_NO")
         Values(6) = DateToStringExtEx2(Tr.ExportItem.Cheque.GetFieldValue("CHEQUE_DATE"))
         Values(7) = DateToStringExtEx2(Tr.ExportItem.Cheque.GetFieldValue("EFFECTIVE_DATE"))
         Values(8) = Tr.ExportItem.GetFieldValue("PAYMENT_TYPE")
         
      ElseIf DocumentType = CASH_DEPOSIT2 Then
         If m_CashDoc.CashTranItems.Count <= 0 Then
            Exit Sub
         End If
         Set Ct1 = GetItem(m_CashDoc.CashTranItems, RowIndex, RealIndex)
         If Ct1 Is Nothing Then
            Exit Sub
         End If
   
         Values(1) = Ct1.GetFieldValue("CASH_TRAN_ID")
         Values(2) = RealIndex
         Values(3) = Ct1.GetFieldValue("PAYMENT_TYPE_NAME")
         Values(4) = FormatNumber(Ct1.GetFieldValue("AMOUNT"))
         Values(5) = Ct1.Cheque.GetFieldValue("CHEQUE_NO")
         Values(6) = DateToStringExtEx2(Ct1.Cheque.GetFieldValue("CHEQUE_DATE"))
         Values(7) = DateToStringExtEx2(Ct1.Cheque.GetFieldValue("EFFECTIVE_DATE"))
      ElseIf DocumentType = CASH_PITTYCASH Then
         If m_CashDoc.CashTranItems.Count <= 0 Then
            Exit Sub
         End If
         Set Ct1 = GetItem(m_CashDoc.CashTranItems, RowIndex, RealIndex)
         If Ct1 Is Nothing Then
            Exit Sub
         End If
   
         Values(1) = Ct1.GetFieldValue("CASH_TRAN_ID")
         Values(2) = RealIndex
         Values(3) = Ct1.GetFieldValue("PAYMENT_TYPE_NAME")
         Values(4) = FormatNumber(Ct1.GetFieldValue("AMOUNT"))
         Values(5) = Ct1.Cheque.GetFieldValue("CHEQUE_NO")
         Values(6) = DateToStringExtEx2(Ct1.Cheque.GetFieldValue("CHEQUE_DATE"))
         Values(7) = DateToStringExtEx2(Ct1.Cheque.GetFieldValue("EFFECTIVE_DATE"))
      End If
   ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-2" Then
'      If m_CashDoc.JournalItems Is Nothing Then
'         Exit Sub
'      End If
'
'      If RowIndex <= 0 Then
'         Exit Sub
'      End If
'
'      Dim Ji As CJournalItem
'      If m_CashDoc.JournalItems.Count <= 0 Then
'         Exit Sub
'      End If
'      Set Ji = GetItem(m_CashDoc.JournalItems, RowIndex, RealIndex)
'      If Ji Is Nothing Then
'         Exit Sub
'      End If
'
'      Values(1) = Ji.GetFieldValue("JOURNAL_ITEM_ID")
'      Values(2) = RealIndex
'      Values(3) = Ji.GetFieldValue("ACC_CODE")
'      Values(4) = Ji.GetFieldValue("ITEM_DESC")
'      If Ji.GetFieldValue("DBCR_TYPE") = 1 Then
'         Values(5) = FormatNumber(Ji.GetFieldValue("DBCR_AMOUNT"))
'         Values(6) = FormatNumber(0)
'      ElseIf Ji.GetFieldValue("DBCR_TYPE") = 2 Then
'         Values(5) = FormatNumber(0)
'         Values(6) = FormatNumber(Ji.GetFieldValue("DBCR_AMOUNT"))
'      End If
   End If
      
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub SSCommand1_Click()

End Sub

Public Sub RefreshGrid(Ind As CASH_DOC_TYPE, Flag As Boolean)
   If TabStrip1.SelectedItem.Tag = DocumentType & "-1" Then
      If (Ind = CASH_WITHDRAW) Or (Ind = CASH_TRANSFER) Or (Ind = CASH_DEPOSIT) Then
         GridEX1.ItemCount = CountItem(m_CashDoc.TransferItems)
         GridEX1.Rebind
      ElseIf (Ind = CASH_WHTHDRAW2) Or (Ind = CASH_DEPOSIT2) Or (Ind = CASH_PITTYCASH) Then
         GridEX1.ItemCount = CountItem(m_CashDoc.CashTranItems)
         GridEX1.Rebind
      ElseIf (Ind = POST_CHEQUE) Or (Ind = WAITING_CHEQUE) Or (Ind = PASSED_CHEQUE) Then
         GridEX1.ItemCount = CountItem(m_CashDoc.PostItems)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-2" Then
         GridEX1.ItemCount = CountItem(m_CashDoc.JournalItems)
         GridEX1.Rebind
   End If
   
   Call GetTotalPrice
   If Flag Then
      m_HasModify = Flag
   End If
End Sub

Private Sub TabStrip1_Click()
   If TabStrip1.SelectedItem Is Nothing Then
      Exit Sub
   End If
   
   Call InitGrid1(DocumentType)
   Call RefreshGrid(DocumentType, False)
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

Private Sub txtNote_Change()
   m_HasModify = True
End Sub
Private Sub uctlBankLookup_Change()
Dim TempID As Long
Dim BB As CBankBranch
   TempID = uctlBankLookup.MyCombo.ItemData(Minus2Zero(uctlBankLookup.MyCombo.ListIndex))
   
   If TempID > 0 Then
      Call LoadBankBranch(uctlBankBranchLookup.MyCombo, m_BankBranchs, TempID)
      Set uctlBankBranchLookup.MyCollection = m_BankBranchs
   End If
   
   m_HasModify = True
End Sub
Private Sub uctlBankAccountLookup_Change()
   m_HasModify = True
End Sub
Private Sub uctlBankBranchLookup_Change()
Dim TempID1 As Long
Dim TempID2 As Long
   
   TempID1 = uctlBankLookup.MyCombo.ItemData(Minus2Zero(uctlBankLookup.MyCombo.ListIndex))
   TempID2 = uctlBankBranchLookup.MyCombo.ItemData(Minus2Zero(uctlBankBranchLookup.MyCombo.ListIndex))
   
   If TempID2 > 0 Then
      Call LoadMaster(uctlBankAccountLookup.MyCombo, m_BankAccounts, BANK_ACCOUNT, TempID1, TempID2)
      Set uctlBankAccountLookup.MyCollection = m_BankAccounts
   End If
   
   m_HasModify = True
End Sub


Private Sub uctlCustomerLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlDocumentDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlEmployeeLookup_Change()
   m_HasModify = True
End Sub
Public Function CashDocPost2BillingDoc(Cd As CCashDoc, BD As CBillingDoc, IvdDocType As Long) As Boolean
Dim IsOK As Boolean
Dim CP As CCashDocPost
   
   For Each CP In Cd.PostItems
      If CP.Post2BD.Count > 0 Then
         If CP.Flag = "" Then
            CP.Flag = "E"
         End If
         Set BD = CP.Post2BD(1)
         BD.Flag = "E"
      Else
         If CP.Flag = "" Then
            CP.Flag = "E"
         End If
         Set BD = New CBillingDoc
         BD.Flag = "A"
         Call CP.Post2BD.add(BD)
      End If
      
      BD.DOCUMENT_NO = CP.GetFieldValue("BILLING_DOC_NO")         '�����Ţ PV NO
      BD.DOCUMENT_DATE = Cd.GetFieldValue("DOCUMENT_DATE")
      BD.SUPPLIER_ID = CP.GetFieldValue("CHEQUE_SUPPLIER_ID")
      BD.PAID_AMOUNT = CP.GetFieldValue("CHEQUE_AMOUNT") + CP.GetFieldValue("WH_AMOUNT") - CP.GetFieldValue("INTERREST_AMOUNT")
      BD.DOCUMENT_TYPE = IvdDocType
      BD.COMMIT_FLAG = "N"
      BD.EXCEPTION_FLAG = "N"
   Next CP
End Function

