VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddEditChequeDoc 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditChequeDoc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000009&
      Height          =   1275
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   1575
      TabIndex        =   19
      Top             =   240
      Visible         =   0   'False
      Width           =   1635
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   8520
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlDate uctlPassChequeDate 
         Height          =   375
         Left            =   1800
         TabIndex        =   42
         Top             =   1920
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   661
      End
      Begin prjFarmManagement.uctlDate uctlBadChequeDate 
         Height          =   495
         Left            =   1800
         TabIndex        =   41
         Top             =   2400
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   873
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   4200
         Left            =   120
         TabIndex        =   8
         Top             =   3600
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   7408
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
         Column(1)       =   "frmAddEditChequeDoc.frx":27A2
         Column(2)       =   "frmAddEditChequeDoc.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditChequeDoc.frx":290E
         FormatStyle(2)  =   "frmAddEditChequeDoc.frx":2A6A
         FormatStyle(3)  =   "frmAddEditChequeDoc.frx":2B1A
         FormatStyle(4)  =   "frmAddEditChequeDoc.frx":2BCE
         FormatStyle(5)  =   "frmAddEditChequeDoc.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditChequeDoc.frx":2D5E
      End
      Begin Threed.SSFrame SSFrame3 
         Height          =   2085
         Left            =   11910
         TabIndex        =   26
         Top             =   5700
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   3678
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin VB.ComboBox cboBankBranch 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   1170
            Width           =   4035
         End
         Begin VB.ComboBox cboBank 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   720
            Width           =   4035
         End
         Begin VB.ComboBox cboPaymentType 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   270
            Width           =   2325
         End
         Begin prjFarmManagement.uctlTextBox txtCheckNo 
            Height          =   435
            Left            =   7470
            TabIndex        =   28
            Top             =   210
            Width           =   2625
            _ExtentX        =   5001
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlDate uctlCheckDate 
            Height          =   405
            Left            =   7470
            TabIndex        =   30
            Top             =   660
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   714
         End
         Begin VB.Label lblCheckDate 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5940
            TabIndex        =   36
            Top             =   690
            Width           =   1395
         End
         Begin VB.Label lblBankBranch 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   150
            TabIndex        =   35
            Top             =   1260
            Width           =   1275
         End
         Begin VB.Label lblBank 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   150
            TabIndex        =   34
            Top             =   810
            Width           =   1275
         End
         Begin VB.Label lblCheckNo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5700
            TabIndex        =   33
            Top             =   270
            Width           =   1665
         End
         Begin VB.Label lblPaymentType 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   150
            TabIndex        =   32
            Top             =   360
            Width           =   1275
         End
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   2085
         Left            =   270
         TabIndex        =   20
         Top             =   5730
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   3678
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin prjFarmManagement.uctlTextLookup uctlResource 
            Height          =   435
            Left            =   1740
            TabIndex        =   21
            Top             =   120
            Width           =   5385
            _ExtentX        =   9499
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtPaidFor 
            Height          =   435
            Left            =   1770
            TabIndex        =   24
            Top             =   1080
            Width           =   9585
            _ExtentX        =   16907
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlDate uctlPaidDate 
            Height          =   405
            Left            =   7530
            TabIndex        =   23
            Top             =   570
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   714
         End
         Begin prjFarmManagement.uctlTextBox txtPvNo 
            Height          =   435
            Left            =   1770
            TabIndex        =   22
            Top             =   600
            Width           =   2625
            _ExtentX        =   5001
            _ExtentY        =   767
         End
         Begin Threed.SSCommand cmdPvNo 
            Height          =   405
            Left            =   4440
            TabIndex        =   40
            Top             =   600
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   714
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmAddEditChequeDoc.frx":2F36
            ButtonStyle     =   3
         End
         Begin VB.Label lblPvNo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            TabIndex        =   39
            Top             =   660
            Width           =   1545
         End
         Begin VB.Label lblPaidFor 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            TabIndex        =   38
            Top             =   1140
            Width           =   1515
         End
         Begin VB.Label lblPaidDate 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6000
            TabIndex        =   37
            Top             =   600
            Width           =   1395
         End
         Begin VB.Label lblResource 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   30
            TabIndex        =   25
            Top             =   180
            Width           =   1635
         End
      End
      Begin prjFarmManagement.uctlTextLookup uctlCustomerLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   4
         Top             =   1380
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlChequeDocDate 
         Height          =   405
         Left            =   6600
         TabIndex        =   2
         Top             =   960
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   120
         TabIndex        =   7
         Top             =   3120
         Width           =   11475
         _ExtentX        =   20241
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
      Begin prjFarmManagement.uctlTextBox txtChequeDocNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   930
         Width           =   2535
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   1080
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin Threed.SSCheck chkPassCheque 
         Height          =   435
         Left            =   6000
         TabIndex        =   6
         Top             =   1920
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblPassChequeDate 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label lblBadChequeDate 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   1920
         Width           =   1575
      End
      Begin Threed.SSCommand cmdCustomer 
         Height          =   405
         Left            =   7260
         TabIndex        =   5
         Top             =   1380
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditChequeDoc.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAuto 
         Height          =   405
         Left            =   4410
         TabIndex        =   1
         Top             =   930
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditChequeDoc.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCheck chkBadCheque 
         Height          =   435
         Left            =   6000
         TabIndex        =   3
         Top             =   2400
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblCustomer 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   150
         TabIndex        =   18
         Top             =   1440
         Width           =   1635
      End
      Begin VB.Label lblChequeDocDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5040
         TabIndex        =   17
         Top             =   990
         Width           =   1395
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   12
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditChequeDoc.frx":3884
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   13
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditChequeDoc.frx":3B9E
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   11
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditChequeDoc.frx":3EB8
         ButtonStyle     =   3
      End
      Begin VB.Label lblChequeDocNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   90
         TabIndex        =   15
         Top             =   990
         Width           =   1665
      End
   End
End
Attribute VB_Name = "frmAddEditChequeDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_BillingDoc As CBillingDoc
Private m_ChequeDoc As CChequeDoc
Private m_Customers As Collection
Private m_Employees As Collection
Private m_Resources As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public ReceiptType As Long
Public Area As Long
Public DocumentType As Long
Public CUSTOMER_ID  As String

Private FileName As String
Private m_SumUnit As Double

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
'      m_BillingDoc.BILLING_DOC_ID = ID
      m_ChequeDoc.CHEQUE_DOC_ID = ID
      If Not glbDaily.QueryChequeDoc(m_ChequeDoc, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_ChequeDoc.PopulateFromRS(1, m_Rs)
      
      uctlChequeDocDate.ShowDate = m_ChequeDoc.CHEQUE_DOC_DATE
      txtChequeDocNo.Text = m_ChequeDoc.CHEQUE_DOC_NO
       uctlCustomerLookup.MyCombo.ListIndex = IDToListIndex(uctlCustomerLookup.MyCombo, m_ChequeDoc.CUSTOMER_ID)
       uctlBadChequeDate.ShowDate = m_ChequeDoc.BADCHEQUE_DATE
       uctlPassChequeDate.ShowDate = m_ChequeDoc.PASSCHEQUE_DATE
       chkPassCheque.Value = FlagToCheck(m_ChequeDoc.PASSCHEQUE_FLAG)
       chkBadCheque.Value = FlagToCheck(m_ChequeDoc.BADCHEQUE_FLAG)
'      If Area = 1 Then
'         uctlCustomerLookup.MyCombo.ListIndex = IDToListIndex(uctlCustomerLookup.MyCombo, m_BillingDoc.CUSTOMER_ID)
'         cboAccount.ListIndex = IDToListIndex(cboAccount, m_BillingDoc.ACCOUNT_ID)
'         cmdPrint.Enabled = True
'      ElseIf Area = 2 Then
'         uctlCustomerLookup.MyCombo.ListIndex = IDToListIndex(uctlCustomerLookup.MyCombo, m_BillingDoc.SUPPLIER_ID)
'         cboAccount.ListIndex = -1
'      End If
''''''      cboCustomerAddress.ListIndex = IDToListIndex(cboCustomerAddress, m_BillingDoc.BILLING_ADDRESS_ID)
''''''      cboEnpAddress.ListIndex = IDToListIndex(cboEnpAddress, m_BillingDoc.ENTERPRISE_ADDRESS_ID)

'      uctlSellByLookup.MyCombo.ListIndex = IDToListIndex(uctlSellByLookup.MyCombo, m_BillingDoc.ACCEPT_BY)
'      txtVatPercent.Text = m_BillingDoc.VAT_PERCENT
'      txtWH.Text = m_BillingDoc.WH_PERCENT
'      uctlResource.MyCombo.ListIndex = IDToListIndex(uctlResource.MyCombo, m_BillingDoc.RESOURCE_ID)
'      txtTotalRcp.Text = m_BillingDoc.TOTAL_RCP
      
'      txtPvNo.Text = m_BillingDoc.PV_NO
'      uctlPaidDate.ShowDate = m_BillingDoc.DUE_DATE
'      txtPaidFor.Text = m_BillingDoc.NOTE
      
'      chkCommit.Value = FlagToCheck(m_BillingDoc.OLD_COMMIT_FLAG)
'      chkCommit.Enabled = (m_BillingDoc.OLD_COMMIT_FLAG = "N")
'      chkPayFlag.Value = FlagToCheck(m_BillingDoc.PAY_FLAG)
      
'      txtCheckNo.Text = m_BillingDoc.CHECK_NO
'      uctlCheckDate.ShowDate = m_BillingDoc.CHECK_DATE
'      cboPaymentType.ListIndex = IDToListIndex(cboPaymentType, m_BillingDoc.PAYMENT_TYPE)
'      cboBank.ListIndex = IDToListIndex(cboBank, m_BillingDoc.BANK_ID)
'      cboBankBranch.ListIndex = IDToListIndex(cboBankBranch, m_BillingDoc.BBRANCH_ID)
'
      Call EnableDisableButton(True)
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

Private Sub PopulateGuiID(Bd As CBillingDoc)
Dim Di As CDoItem

   For Each Di In Bd.DoItems
      If Di.Flag = "A" Then
         Di.LINK_ID = GetNextGuiID(Bd)
      End If
   Next Di
End Sub

Private Function GetNextGuiID(Bd As CBillingDoc) As Long
Dim Di As CDoItem
Dim MaxId As Long

   MaxId = 0
   For Each Di In Bd.DoItems
      If Di.LINK_ID > MaxId Then
         MaxId = Di.LINK_ID
      End If
   Next Di
   
   GetNextGuiID = MaxId + 1
End Function

Public Function GetExportItem(Ivd As CInventoryDoc, GuiID As Long) As CLotItem
Dim Ei As CLotItem

      For Each Ei In Ivd.ImportExports
         If Ei.LINK_ID = GuiID Then
            Set GetExportItem = Ei
            Exit Function
         End If
      Next Ei
End Function

Private Function VerifyJournalItem(Bd As CBillingDoc) As Boolean
Dim Gl As CGLDetail
Dim SumDr As Double
Dim SumCr As Double

   SumDr = 0
   SumCr = 0
   For Each Gl In Bd.GlDetails
      If Gl.Flag <> "D" Then
         If Gl.GetFieldValue("GL_TYPE") = 1 Then
            SumDr = SumDr + Gl.GetFieldValue("GL_AMOUNT")
         ElseIf Gl.GetFieldValue("GL_TYPE") = 2 Then
            SumCr = SumCr + Gl.GetFieldValue("GL_AMOUNT")
         End If
      End If
   Next Gl
   
   If FormatNumber(SumDr) <> FormatNumber(SumCr) Then
      VerifyJournalItem = False
   Else
      VerifyJournalItem = True
   End If
End Function

Private Function MyCountItem(Col As Collection) As Long
Dim I As Long
Dim Count As Long
Dim Ji As CCashTran

   Count = 0
   For I = 1 To Col.Count
      Set Ji = Col.Item(I)
      If (Ji.Flag <> "D") And (Ji.Cheque.GetFieldValue("EFFECTIVE_DATE") > 0) Then
         Count = Count + 1
      End If
   Next I
   
   MyCountItem = Count
End Function

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim Ivd As CInventoryDoc
Dim Pm As CPayment
Dim Ct As CCashTran
   If ShowMode = SHOW_EDIT Then
      If Area = 1 Then
         If Not VerifyAccessRight("LEDGER_SELL" & "_" & DocumentType & "_" & "EDIT", "แก้ไข") Then
            Call EnableForm(Me, True)
            Exit Function
         End If
      ElseIf Area = 2 Then
         If Not VerifyAccessRight("LEDGER_BUY" & "_" & DocumentType & "_" & "EDIT", "แก้ไข") Then
            Call EnableForm(Me, True)
            Exit Function
         End If
      End If
   End If
   If Not VerifyTextControl(lblChequeDocNo, txtChequeDocNo, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblChequeDocDate, uctlChequeDocDate, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblCustomer, uctlCustomerLookup.MyCombo, True) Then
      Exit Function
   End If
'   If Not VerifyTextControl(lblVat, txtVatPercent, True) Then
'      Exit Function
'   End If
'   If Not VerifyTextControl(lblWH, txtWH, True) Then
'      Exit Function
'   End If
'   If Not VerifyCombo(lblPaymentType, cboPaymentType, False) Then
'      Exit Function
'   End If
   
'   If Not (txtDocumentNo.Text = txtPvNo.Text) And DocumentType = 8 Then
'      glbErrorLog.LocalErrorMsg = " ! หมายเลขเอกสาร กับหมายเลข PV ไม่ตรงกัน "
'      glbErrorLog.ShowUserError
'      Exit Function
'   End If
'
'   If Not CheckUniqueNs(DO_PLAN_UNIQUE, txtDocumentNo.Text, ID) Then
'      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtDocumentNo.Text & " " & MapText("อยู่ในระบบแล้ว")
'      glbErrorLog.ShowUserError
'      Exit Function
'   End If
'
'   If CountItem(m_BillingDoc.Payments) <= 0 And Area = 1 Then
'      glbErrorLog.LocalErrorMsg = "กรุณาใส่การชำระเงินใหถูกต้องครบถ้วน"
'      glbErrorLog.ShowUserError
'      Exit Function
'   End If
'
'   If Not VerifyJournalItem(m_BillingDoc) Then
'      glbErrorLog.LocalErrorMsg = "ยอดรวมเดบิตต้องเท่ากับยอดรวมเครดิต"
'      glbErrorLog.ShowUserError
'      Exit Function
'   End If
'
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
'   m_BillingDoc.AddEditMode = ShowMode
'   m_BillingDoc.BILLING_DOC_ID = ID
'    m_BillingDoc.DOCUMENT_DATE = uctlDocumentDate.ShowDate
'   m_BillingDoc.DOCUMENT_NO = txtDocumentNo.Text
   m_ChequeDoc.AddEditMode = ShowMode
   m_ChequeDoc.CHEQUE_DOC_ID = ID
   m_ChequeDoc.CHEQUE_DOC_DATE = uctlChequeDocDate.ShowDate
   m_ChequeDoc.CHEQUE_DOC_NO = txtChequeDocNo.Text
   m_ChequeDoc.BADCHEQUE_FLAG = Check2Flag(chkBadCheque.Value)
   m_ChequeDoc.PASSCHEQUE_FLAG = Check2Flag(chkPassCheque.Value)
  m_ChequeDoc.BADCHEQUE_DATE = uctlBadChequeDate.ShowDate
  m_ChequeDoc.PASSCHEQUE_DATE = uctlPassChequeDate.ShowDate
  
 m_ChequeDoc.CUSTOMER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
'    uctlBadChequeDate
'   If Area = 1 Then
'      m_BillingDoc.DOCUMENT_TYPE = 2 'ใบเสร็จรับเงิน
'      m_BillingDoc.CUSTOMER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
''      m_BillingDoc.ACCOUNT_ID = cboAccount.ItemData(Minus2Zero(cboAccount.ListIndex))
'   ElseIf Area = 2 Then
'      m_BillingDoc.DOCUMENT_TYPE = 8 'ใบเสร็จรับเงิน
'      m_BillingDoc.SUPPLIER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
'      m_BillingDoc.ACCOUNT_ID = -1
'   End If
''   m_BillingDoc.BILLING_ADDRESS_ID = cboCustomerAddress.ItemData(Minus2Zero(cboCustomerAddress.ListIndex))
''   m_BillingDoc.ENTERPRISE_ADDRESS_ID = cboEnpAddress.ItemData(Minus2Zero(cboEnpAddress.ListIndex))
'   m_BillingDoc.EXCEPTION_FLAG = "N"
''   m_BillingDoc.WH_PERCENT = Val(txtWH.Text)
''   m_BillingDoc.WH_AMOUNT = Val(txtWHAmount.Text)
''   m_BillingDoc.VAT_PERCENT = Val(txtVatPercent.Text)
''   m_BillingDoc.VAT_AMOUNT = Val(txtVatAmount.Text)
''   m_BillingDoc.ACCEPT_BY = uctlSellByLookup.MyCombo.ItemData(Minus2Zero(uctlSellByLookup.MyCombo.ListIndex))
'   m_BillingDoc.RECEIPT_TYPE = ReceiptType
''   m_BillingDoc.DISCOUNT_AMOUNT = Val(txtDiscount.Text)
''   m_BillingDoc.TOTAL_RCP = Val(txtTotalRcp.Text)
'   If MyCountItem(m_BillingDoc.Payments) <= 0 Then 'effective เมื่อมีการ คีย์เช็คจ่ายแล้ว
'      m_BillingDoc.EFFECTIVE_FLAG = "N"
'   Else
'      m_BillingDoc.EFFECTIVE_FLAG = "Y"
'   End If
''   m_BillingDoc.TOTAL_AMOUNT = Val(txtTotalAmount.Text)
''   m_BillingDoc.TOTAL_PRICE = Val(txtNetTotal.Text)
'   m_BillingDoc.RESOURCE_ID = uctlResource.MyCombo.ItemData(Minus2Zero(uctlResource.MyCombo.ListIndex))
''   m_BillingDoc.PAY_FLAG = Check2Flag(chkPayFlag.Value)
''   m_BillingDoc.COMMIT_FLAG = Check2Flag(chkCommit.Value)
   
'   m_BillingDoc.CHECK_NO = txtCheckNo.Text
'   m_BillingDoc.CHECK_DATE = uctlCheckDate.ShowDate
'   m_BillingDoc.PAYMENT_TYPE = cboPaymentType.ItemData(Minus2Zero(cboPaymentType.ListIndex))
'   m_BillingDoc.BANK_ID = cboBank.ItemData(Minus2Zero(cboBank.ListIndex))
'   If cboBankBranch.ListIndex > 0 Then
'      m_BillingDoc.BBRANCH_ID = cboBankBranch.ItemData(Minus2Zero(cboBankBranch.ListIndex))
'   End If
'   m_BillingDoc.PV_NO = txtPvNo.Text
'   m_BillingDoc.DUE_DATE = uctlPaidDate.ShowDate
'   m_BillingDoc.NOTE = txtPaidFor.Text
   
'   Call PopulateGuiID(m_BillingDoc)
'
'   Call EnableForm(Me, False)
'
'   Call glbDaily.DO2InventoryDoc(m_BillingDoc, Ivd, Area, 21)
''   Call glbDaily.DO2Payment(m_BillingDoc, Pm)
'
'   If (m_BillingDoc.COMMIT_FLAG = "Y") Then
'      If m_BillingDoc.OLD_COMMIT_FLAG <> "Y" Then
'         Call glbDaily.TriggerCommit(Ivd.ImportExports)
'         If Not glbDaily.VerifyStockBalance(Ivd.ImportExports, glbErrorLog) Then
'            m_BillingDoc.COMMIT_FLAG = "N"
'            Call EnableForm(Me, True)
'            Exit Function
'         End If
'      End If
'   End If
   
'   Call glbDaily.StartTransaction
'   If Not glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog) Then
'      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'      SaveData = False
'      Call glbDaily.RollbackTransaction
'      Call EnableForm(Me, True)
'      Exit Function
'   End If
   
'   If Not glbDaily.AddEditPayment(Pm, IsOK, False, glbErrorLog) Then
'      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'      SaveData = False
'      Call glbDaily.RollbackTransaction
'      Call EnableForm(Me, True)
'      Exit Function
'   End If
'   m_BillingDoc.PAYMENT_ID = Pm.PAYMENT_ID
      
'   m_BillingDoc.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID
'   If Not glbDaily.AddEditBillingDoc(m_BillingDoc, IsOK, False, glbErrorLog) Then
'      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'      SaveData = False
'      Call glbDaily.RollbackTransaction
'      Call EnableForm(Me, True)
'      Exit Function
'   End If
   
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditChequeDoc(m_ChequeDoc, IsOK, True, glbErrorLog) Then
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
      KeyAscii = 0
   End If
End Sub

Private Sub cboBank_Click()
Dim BankID As Long

   BankID = cboBank.ItemData(Minus2Zero(cboBank.ListIndex))
   If BankID > 0 Then
      Call LoadBankBranch(cboBankBranch, , BankID)
   End If

   m_HasModify = True
End Sub

Private Sub cboBankBranch_Click()
   m_HasModify = True
End Sub

Private Sub cboCustomerAddress_Click()
   m_HasModify = True
End Sub

Private Sub cboCustomerAddress_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
      KeyAscii = 0
   End If
End Sub

Private Sub cboEnpAddress_Click()
   m_HasModify = True
End Sub

Private Sub cboEnpAddress_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
      KeyAscii = 0
   End If
End Sub

Private Sub cboPaymentType_Click()
   m_HasModify = True
End Sub

Private Sub chkCommit_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkCommit_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
      KeyAscii = 0
   End If
End Sub

Public Sub RefreshGrid()
   Call GetTotalPrice

   GridEX1.ItemCount = CountItem(m_BillingDoc.Payments)
   GridEX1.Rebind
End Sub

Private Sub chkPayFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkBadCheque_Click(Value As Integer)
  m_HasModify = True

End Sub

Private Sub chkPassCheque_Click(Value As Integer)
  m_HasModify = True
End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu

   If Not cmdAdd.Enabled Then
      Exit Sub
   End If
   
'   If Area = 1 Then
''      If Not VerifyCombo(lblAccountNo, cboAccount) Then
''         Exit Sub
''      End If
'      If Not VerifyDate(lblDocumentDate, uctlDocumentDate) Then
'         Exit Sub
'      End If
'   ElseIf Area = 2 Then
      If Not VerifyCombo(lblCustomer, uctlCustomerLookup.MyCombo) Then
         Exit Sub
      End If
'   End If
'
'   OKClick = False
'  If TabStrip1.SelectedItem.Index = 1 Then
'      If (ReceiptType = 3) Or (ReceiptType = 5) Then
'         frmAddReceiptItem.Area = Area
'         frmAddReceiptItem.ReceiptType = ReceiptType
'         If Area = 1 Then
''            frmAddReceiptItem.AccountID = cboAccount.ItemData(Minus2Zero(cboAccount.ListIndex))
'         ElseIf Area = 2 Then
'            frmAddReceiptItem.AccountID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
'         End If
'         Set frmAddEditReceiptDoc.TempCollection = m_BillingDoc.ReceiptItems
'         frmAddEditReceiptDoc.ShowMode = SHOW_ADD
'        frmAddEditReceiptDoc.HeaderText = MapText("เพิ่มรายการเช็ค")
'         Load frmAddEditReceiptDoc
'         frmAddEditReceiptDoc.Show 1
'
'         OKClick = frmAddEditReceiptDoc.OKClick
'
'         Unload frmAddEditReceiptDoc
'         Set frmAddEditReceiptDoc = Nothing
'
'         If OKClick Then
'            Call GetTotalPrice
'
'            GridEX1.ItemCount = CountItem(m_BillingDoc.ReceiptItems)
'            GridEX1.Rebind
'         End If
'      ElseIf ReceiptType = 1 Then
'         If Area = 1 Then
'            Set oMenu = New cPopupMenu
'            'Same as DO
'            lMenuChosen = oMenu.AddMenu(glbGuiConfigs.DOAddMenuItems)
'            Set oMenu = Nothing
'            If lMenuChosen = 0 Then
'               Exit Sub
'            End If
'         Else
'            lMenuChosen = 1
'         End If
'
'         If lMenuChosen = 1 Then
''            If Area = 1 Then
''               frmAddEditDoItem.AccountID = cboAccount.ItemData(cboAccount.ListIndex)
''            End If
'
'            frmAddEditDoItem.DocumentDate = uctlDocumentDate.ShowDate
'            frmAddEditDoItem.SubscriberID = -1
'            frmAddEditDoItem.Area = Area
'            frmAddEditDoItem.COMMIT_FLAG = m_BillingDoc.OLD_COMMIT_FLAG
'            Set frmAddEditDoItem.TempCollection = m_BillingDoc.DoItems
'            frmAddEditDoItem.ParentShowMode = ShowMode
'            frmAddEditDoItem.ShowMode = SHOW_ADD
'            frmAddEditDoItem.HeaderText = MapText("เพิ่มรายการใบเสร็จ")
'            Load frmAddEditDoItem
'            frmAddEditDoItem.Show 1
'
'            OKClick = frmAddEditDoItem.OKClick
'
'            Unload frmAddEditDoItem
'            Set frmAddEditDoItem = Nothing
'
'            If OKClick Then
'               Call GetTotalPriceEx
'
'               GridEX1.ItemCount = CountItem(m_BillingDoc.DoItems)
'               GridEX1.Rebind
'            End If
'         ElseIf lMenuChosen = 2 Then
''            If Area = 1 Then
''               frmAddEditDoItemEx.AccountID = cboAccount.ItemData(cboAccount.ListIndex)
''            Else
''               glbErrorLog.LocalErrorMsg = "ฟังก์ชันนี้ไม่สนับสนุนในส่วนงานซื้อ"
''               glbErrorLog.ShowUserError
''               Exit Sub
''            End If
'            Set frmAddEditDoItemEx.ParentForm = Me
'            frmAddEditDoItemEx.SubscriberID = -1
'            frmAddEditDoItemEx.Area = Area
'            frmAddEditDoItemEx.COMMIT_FLAG = m_BillingDoc.OLD_COMMIT_FLAG
'            Set frmAddEditDoItemEx.TempCollection = m_BillingDoc.DoItems
'            frmAddEditDoItemEx.ParentShowMode = ShowMode
'            frmAddEditDoItemEx.ShowMode = SHOW_ADD
'            frmAddEditDoItemEx.HeaderText = MapText("เพิ่มรายการใบเสร็จ")
'            Load frmAddEditDoItemEx
'            frmAddEditDoItemEx.Show 1
'
'            OKClick = frmAddEditDoItemEx.OKClick
'
'            Unload frmAddEditDoItemEx
'            Set frmAddEditDoItemEx = Nothing
'
'            If OKClick Then
'               Call GetTotalPriceEx
'
'               GridEX1.ItemCount = CountItem(m_BillingDoc.DoItems)
'               GridEX1.Rebind
'            End If
'         ElseIf lMenuChosen = 4 Then
''            frmAddPOItem.AccountID = cboAccount.ItemData(Minus2Zero(cboAccount.ListIndex))
'            Set frmAddPOItem.TempCollection = m_BillingDoc.DoItems
'            frmAddPOItem.ShowMode = SHOW_ADD
'            frmAddPOItem.HeaderText = MapText("เพิ่มรายการใบเสร็จ จากใบ PO")
'            Load frmAddPOItem
'            frmAddPOItem.Show 1
'
'            OKClick = frmAddPOItem.OKClick
'
'            Unload frmAddPOItem
'            Set frmAddPOItem = Nothing
'
'            If OKClick Then
'               Call GetTotalPrice
'
'               GridEX1.ItemCount = CountItem(m_BillingDoc.DoItems)
'               GridEX1.Rebind
'            End If
'         ElseIf lMenuChosen = 5 Then
''            frmAddQuoatationItem.AccountID = cboAccount.ItemData(Minus2Zero(cboAccount.ListIndex))
'            Set frmAddQuoatationItem.TempCollection = m_BillingDoc.DoItems
'            frmAddQuoatationItem.ShowMode = SHOW_ADD
'            frmAddQuoatationItem.HeaderText = MapText("เพิ่มรายการใบเสร็จจากใบเสนอราคา")
'            Load frmAddQuoatationItem
'            frmAddQuoatationItem.Show 1
'
'            OKClick = frmAddQuoatationItem.OKClick
'
'            Unload frmAddQuoatationItem
'            Set frmAddQuoatationItem = Nothing
'
'            If OKClick Then
'               Call GetTotalPrice
'
'               GridEX1.ItemCount = CountItem(m_BillingDoc.DoItems)
'               GridEX1.Rebind
'            End If
'         End If
'      End If
'   ElseIf TabStrip1.SelectedItem.Index = 2 Then
'   ElseIf TabStrip1.SelectedItem.Index = 3 Then
'      Set frmAddEditCashTran.GnlItem = m_BillingDoc.GlDetails
'      frmAddEditCashTran.Area = Area
'      Set frmAddEditCashTran.ParentForm = Me
'      frmAddEditCashTran.HeaderText = "เพิ่มรายการการชำระเงิน"
'      frmAddEditCashTran.ShowMode = SHOW_ADD
'      Set frmAddEditCashTran.TempCollection = m_BillingDoc.Payments
'      Load frmAddEditCashTran
'      frmAddEditCashTran.Show 1
'
'      OKClick = frmAddEditCashTran.OKClick
'
'      Unload frmAddEditCashTran
'      Set frmAddEditCashTran = Nothing
'
'      If OKClick Then
'         m_HasModify = True
'
'         GridEX1.ItemCount = CountItem(m_BillingDoc.Payments)
'         Call GridEX1.Rebind
'
'         Call GetTotalPrice
'      End If
'   ElseIf TabStrip1.SelectedItem.Index = 4 Then
'      Set frmAddEditGlDetail.ParentForm = Me
'      frmAddEditGlDetail.HeaderText = "เพิ่มรายการสมุดรายวัน"
'      frmAddEditGlDetail.ShowMode = SHOW_ADD
'      Set frmAddEditGlDetail.TempCollection = m_BillingDoc.GlDetails
'      Load frmAddEditGlDetail
'      frmAddEditGlDetail.Show 1
'
'      OKClick = frmAddEditGlDetail.OKClick
'
'      Unload frmAddEditGlDetail
'      Set frmAddEditGlDetail = Nothing
'
'      If OKClick Then
'         m_HasModify = True
'
'         GridEX1.ItemCount = CountItem(m_BillingDoc.GlDetails)
'         Call GridEX1.Rebind
'
'         Call GetTotalPrice
'      End If
'
'   ElseIf TabStrip1.SelectedItem.Index = 5 Then
'   End If

'        frmAddReceiptItem.Area = Area
'        frmAddReceipt
'         frmAddReceiptItem.ReceiptType = 1

'         If Area = 1 Then
''            frmAddReceiptItem.AccountID = cboAccount.ItemData(Minus2Zero(cboAccount.ListIndex))
'         ElseIf Area = 2 Then
'            frmAddReceiptItem.AccountID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
'         End If

        frmAddEditReceiptDoc.Area = Area
        frmAddEditReceiptDoc.CustomerID = CUSTOMER_ID
'       uctlCustomerLookup.MyCombo.ListIndex = IDToListIndex(uctlCustomerLookup.MyCombo, m_ChequeDoc.CUSTOMER_ID)
         Set frmAddEditReceiptDoc.TempCollection = m_BillingDoc.ReceiptItems
         
         frmAddEditReceiptDoc.ShowMode = SHOW_ADD
        frmAddEditReceiptDoc.HeaderText = MapText("เพิ่มรายการเช็ค")
         Load frmAddEditReceiptDoc
         frmAddEditReceiptDoc.Show 1
   
         OKClick = frmAddEditReceiptDoc.OKClick
   
         Unload frmAddEditReceiptDoc
         Set frmAddEditReceiptDoc = Nothing
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub
Private Sub cmdAuto_Click()
Dim No As String

'   If Trim(txtChequeDocNo.Text) = "" Then
'      Call glbDatabaseMngr.GenerateNumber(CHEQUE_DOC_NUMBER, No, glbErrorLog)
'      txtDocumentNo.Text = No
'   End If
End Sub

Private Sub cmdCustomer_Click()
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu
Dim OKClick As Boolean
Dim TempCol As Collection
Dim Cs As CCustomer

   Set TempCol = New Collection
   
   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("ค้นหา", "-", "เพิ่มข้อมูลใหม่")
   If lMenuChosen = 0 Then
      Exit Sub
   End If

   If lMenuChosen = 1 Then
      Set frmQueryCustomer.TempCollection = TempCol
      frmQueryCustomer.ShowMode = SHOW_ADD
      Load frmQueryCustomer
      frmQueryCustomer.Show 1
      
      OKClick = frmQueryCustomer.OKClick
      
      Unload frmQueryCustomer
      Set frmQueryCustomer = Nothing
      
      If OKClick Then
         Set Cs = TempCol(1)
         uctlCustomerLookup.MyCombo.ListIndex = IDToListIndex(uctlCustomerLookup.MyCombo, Cs.CUSTOMER_ID)
         m_HasModify = True
      End If
   ElseIf lMenuChosen = 3 Then
      frmAddEditCustomer.ShowMode = SHOW_ADD
      frmAddEditCustomer.HeaderText = MapText("เพิ่มลูกค้า")
      Load frmAddEditCustomer
      frmAddEditCustomer.Show 1
      
      OKClick = frmAddEditCustomer.OKClick
      Call EnableForm(Me, False)
      If Area = 1 Then
         Call LoadCustomer(uctlCustomerLookup.MyCombo, m_Customers)
         Set uctlCustomerLookup.MyCollection = m_Customers
      ElseIf Area = 2 Then
         Call LoadSupplier(uctlCustomerLookup.MyCombo, m_Customers)
         Set uctlCustomerLookup.MyCollection = m_Customers
      End If
      Call EnableForm(Me, True)
      
      Unload frmAddEditCustomer
      Set frmAddEditCustomer = Nothing
   End If
   
   Set TempCol = Nothing

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
      If (ReceiptType = 3) Or (ReceiptType = 5) Then
         If ID1 <= 0 Then
            m_BillingDoc.ReceiptItems.Remove (ID2)
         Else
            m_BillingDoc.ReceiptItems.Item(ID2).Flag = "D"
         End If
   
         Call GetTotalPrice
         GridEX1.ItemCount = CountItem(m_BillingDoc.ReceiptItems)
         GridEX1.Rebind
         m_HasModify = True
      ElseIf ReceiptType = 1 Then
         If ID1 <= 0 Then
            m_BillingDoc.DoItems.Remove (ID2)
         Else
            m_BillingDoc.DoItems.Item(ID2).Flag = "D"
         End If
   
         Call GetTotalPriceEx
         GridEX1.ItemCount = CountItem(m_BillingDoc.DoItems)
         GridEX1.Rebind
         m_HasModify = True
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
      If ID1 <= 0 Then
         m_BillingDoc.Payments.Remove (ID2)
      Else
         m_BillingDoc.Payments.Item(ID2).Flag = "D"
      End If
      
      Call GetTotalPrice
      GridEX1.ItemCount = CountItem(m_BillingDoc.Payments)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
      If ID1 <= 0 Then
         m_BillingDoc.GlDetails.Remove (ID2)
      Else
         m_BillingDoc.GlDetails.Item(ID2).Flag = "D"
      End If
      
      'Call GetTotalPrice
      GridEX1.ItemCount = CountItem(m_BillingDoc.GlDetails)
      GridEX1.Rebind
      m_HasModify = True

   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean

   If Not cmdEdit.Enabled Then
      Exit Sub
   End If
         
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   If Area = 1 Then
'      If Not VerifyCombo(lblAccountNo, cboAccount) Then
'         Exit Sub
'      End If
      If Not VerifyDate(lblChequeDocDate, uctlChequeDocDate) Then
         Exit Sub
      End If
   End If
   
   ID = Val(GridEX1.Value(2))
   OKClick = False
   
   If TabStrip1.SelectedItem.Index = 1 Then
      If ReceiptType = 1 Then
'         If Area = 1 Then
'            frmAddEditDoItem.AccountID = cboAccount.ItemData(cboAccount.ListIndex)
'         End If

         frmAddEditDoItem.DocumentDate = uctlChequeDocDate.ShowDate
         frmAddEditDoItem.SubscriberID = -1
         frmAddEditDoItem.Area = Area
         frmAddEditDoItem.ID = ID
         frmAddEditDoItem.COMMIT_FLAG = m_BillingDoc.OLD_COMMIT_FLAG
         Set frmAddEditDoItem.TempCollection = m_BillingDoc.DoItems
         frmAddEditDoItem.HeaderText = MapText("แก้ไขรายการใบเสร็จ")
         frmAddEditDoItem.ParentShowMode = ShowMode
         frmAddEditDoItem.ShowMode = SHOW_EDIT
         Load frmAddEditDoItem
         frmAddEditDoItem.Show 1
   
         OKClick = frmAddEditDoItem.OKClick
   
         Unload frmAddEditDoItem
         Set frmAddEditDoItem = Nothing
   
         If OKClick Then
            Call GetTotalPriceEx
            GridEX1.ItemCount = CountItem(m_BillingDoc.DoItems)
            GridEX1.Rebind
         End If
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
      frmAddEditCashTran.Area = Area
      Set frmAddEditCashTran.ParentForm = Me
      frmAddEditCashTran.ID = ID
      frmAddEditCashTran.HeaderText = "แก้ไขรายการการชำระเงิน"
      frmAddEditCashTran.ShowMode = SHOW_EDIT
      Set frmAddEditCashTran.TempCollection = m_BillingDoc.Payments
      Load frmAddEditCashTran
      frmAddEditCashTran.Show 1
      
      OKClick = frmAddEditCashTran.OKClick
      
      Unload frmAddEditCashTran
      Set frmAddEditCashTran = Nothing
   
      If OKClick Then
         m_HasModify = True
         
         GridEX1.ItemCount = CountItem(m_BillingDoc.Payments)
         Call GridEX1.Rebind
         
         Call GetTotalPrice
      End If
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
      Set frmAddEditGlDetail.ParentForm = Me
      frmAddEditGlDetail.ID = ID
      frmAddEditGlDetail.HeaderText = "แก้ไขรายการสมุดรายวัน"
      frmAddEditGlDetail.ShowMode = SHOW_EDIT
      Set frmAddEditGlDetail.TempCollection = m_BillingDoc.GlDetails
      Load frmAddEditGlDetail
      frmAddEditGlDetail.Show 1
      
      OKClick = frmAddEditGlDetail.OKClick
      
      Unload frmAddEditGlDetail
      Set frmAddEditGlDetail = Nothing
   
      If OKClick Then
         m_HasModify = True
         
         GridEX1.ItemCount = CountItem(m_BillingDoc.GlDetails)
         Call GridEX1.Rebind
         
         Call GetTotalPrice
      End If
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub CalculateIncludePrice()
Dim II As CLotItem
Dim AvgFee As Double

'   If m_SumUnit > 0 Then
'      AvgFee = Val(txtTotalAmount.Text) / m_SumUnit
'   Else
'      AvgFee = 0
'   End If
'
'   For Each II In m_BillingDoc.DoItems
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

Private Function VerifyOnwerVersionMenu(Menu As Long, Owner As String) As Boolean
   VerifyOnwerVersionMenu = True
   
   If (Menu <> 1) And (Menu <> 2) Then
      If Not glbGuiConfigs.VerifyGuiConfig("SELL_RECEIPT_PREFORM_PRINT", True) Then
         VerifyOnwerVersionMenu = False
         Exit Function
      End If
   End If
End Function

Private Sub cmdPrint_Click()
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu
Dim ReportFlag As Boolean
Dim ReportKey As String
Dim Report As CReportInterface
Dim Rc As CReportConfig
Dim iCount As Long
Dim EditMode As SHOW_MODE_TYPE
Dim ReportMode As Long
Dim Programowner As String
   Programowner = glbParameterObj.Programowner

   ReportMode = 1
   
   If m_HasModify Or (m_BillingDoc.BILLING_DOC_ID <= 0) Then
      glbErrorLog.LocalErrorMsg = MapText("กรุณาทำการบันทึกข้อมูลให้เรียบร้อยก่อน")
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   ReportFlag = False

   Call LoadPictureFromFile(glbParameterObj.ReceiptPicture1, Picture2)
   
   Set oMenu = New cPopupMenu
   If DocumentType = 8 Then
      lMenuChosen = oMenu.AddMenu(glbGuiConfigs.RCPrintMenuItemsSpacialBuy)
   Else
      If ReceiptType = 1 Then
         lMenuChosen = oMenu.AddMenu(glbGuiConfigs.RCPrintMenuItems)
      Else
         lMenuChosen = oMenu.AddMenu(glbGuiConfigs.RCPrintMenuItemsSpacial)
      End If
   End If
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   Set oMenu = Nothing

'   If Not VerifyOnwerVersionMenu(lMenuChosen, glbParameterObj.Programowner) Then
'      Exit Sub
'   End If
   
'   If lMenuChosen = 1 Then
'      ReportKey = "CReportNormalRcp001"
'
'      Set Report = New CReportNormalRcp001
'      ReportFlag = True
'   ElseIf lMenuChosen = 2 Then
'      ReportKey = "CReportNormalRcp001"
'
'      Set Rc = New CReportConfig
'      Rc.REPORT_KEY = ReportKey
'      Call Rc.QueryData(m_Rs, iCount)
'      HeaderText = MapText("ใบเสร็จรับเงิน")
'      If Not m_Rs.EOF Then
'         Call Rc.PopulateFromRS(1, m_Rs)
'         EditMode = SHOW_EDIT
'      Else
'         EditMode = SHOW_ADD
'      End If
'   ElseIf lMenuChosen = 4 Then
'      ReportKey = "CReportFormReceipt001"
'
'      Set Report = New CReportFormReceipt001
'      ReportFlag = True
'   ElseIf lMenuChosen = 5 Then
'      ReportKey = "CReportFormReceipt001"
'
'      Set Report = New CReportFormReceipt001
'      ReportFlag = True
'   ElseIf lMenuChosen = 6 Then
'      ReportKey = "CReportFormReceipt001"
'      ReportMode = 2
'
'      Set Rc = New CReportConfig
'      Rc.REPORT_KEY = ReportKey
'      Call Rc.QueryData(m_Rs, iCount)
'      HeaderText = MapText("ใบเสร็จรับเงิน")
'      If Not m_Rs.EOF Then
'         Call Rc.PopulateFromRS(1, m_Rs)
'         EditMode = SHOW_EDIT
'      Else
'         EditMode = SHOW_ADD
'      End If
'   ElseIf lMenuChosen = 8 Then
'      ReportKey = "CReportFormPO001"
'
'      Set Report = New CReportFormPO001
'      ReportFlag = True
'   ElseIf lMenuChosen = 9 Then
'      ReportKey = "CReportFormPO001"
'
'      Set Report = New CReportFormPO001
'      ReportFlag = True
'   ElseIf lMenuChosen = 10 Then
'      ReportKey = "CReportFormPO001"
'      ReportMode = 2
'
'      Set Rc = New CReportConfig
'      Rc.REPORT_KEY = ReportKey
'      Call Rc.QueryData(m_Rs, iCount)
'      HeaderText = MapText("ใบเสร็จรับเงิน")
'      If Not m_Rs.EOF Then
'         Call Rc.PopulateFromRS(1, m_Rs)
'         EditMode = SHOW_EDIT
'      Else
'         EditMode = SHOW_ADD
'      End If
'   End If
      If lMenuChosen = 1 Then
      ReportKey = "CReportNormalRcp001"
      
      Set Report = New CReportNormalRcp001
      ReportFlag = True
   ElseIf lMenuChosen = 2 Then
      ReportKey = "CReportNormalRcp001"
      
      Set Rc = New CReportConfig
      Rc.REPORT_KEY = ReportKey
      Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
      Call Rc.QueryData(m_Rs, iCount)
      HeaderText = MapText("ใบเสร็จรับเงิน")
      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
   ElseIf lMenuChosen = 4 Then
      ReportKey = "CReportFormReceipt001"
      
      Set Report = New CReportFormReceipt001
      Call Report.AddParam(1, "DO_TYPE")
 
      ReportFlag = True
   ElseIf lMenuChosen = 5 Then
      ReportKey = "CReportFormReceipt001"
      
      Set Report = New CReportFormReceipt001
    Call Report.AddParam(1, "DO_TYPE")
      ReportFlag = True
      ElseIf lMenuChosen = 6 Then
      ReportKey = "CReportFormReceipt001"
      
      Set Report = New CReportFormReceipt001
      Call Report.AddParam(0, "DO_TYPE")
      ReportFlag = True
   ElseIf lMenuChosen = 7 Then
      ReportKey = "CReportFormReceipt001"
      
      Set Report = New CReportFormReceipt001
      Call Report.AddParam(0, "DO_TYPE")
      ReportFlag = True
   ElseIf lMenuChosen = 8 Then
      ReportKey = "CReportFormReceipt001"
      ReportMode = 2
      
      Set Rc = New CReportConfig
      Rc.REPORT_KEY = ReportKey
      Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
      Call Rc.QueryData(m_Rs, iCount)
      HeaderText = MapText("ใบเสร็จรับเงิน")
      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
   ElseIf lMenuChosen = 10 Then
      ReportKey = "CReportNormalRcpHead"
      Set Report = New CReportNormalRcpHead
      ReportFlag = True
   ElseIf lMenuChosen = 11 Then
      ReportKey = "CReportNormalRcpHead"
      Set Rc = New CReportConfig
      Rc.REPORT_KEY = ReportKey
      Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
      Call Rc.QueryData(m_Rs, iCount)
      HeaderText = MapText("ใบเสร็จรับเงิน")
      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
   ElseIf lMenuChosen = 22 Then
      ReportKey = "CReportVoucherReceive"
      Set Report = New CReportVoucherReceive
      ReportFlag = True
   ElseIf lMenuChosen = 23 Then
      ReportKey = "CReportVoucherReceive"
      ReportMode = 2
      Set Rc = New CReportConfig
      Rc.REPORT_KEY = ReportKey
      Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
      Call Rc.QueryData(m_Rs, iCount)
      HeaderText = MapText("ใบสำคัญรับ")
      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
   ElseIf lMenuChosen = 25 Then
      ReportKey = "CReportVoucherPay"
      Set Report = New CReportVoucherPay
      ReportFlag = True
   ElseIf lMenuChosen = 26 Then
      ReportKey = "CReportVoucherPay"
      ReportMode = 2
      Set Rc = New CReportConfig
      Rc.REPORT_KEY = ReportKey
      Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
      Call Rc.QueryData(m_Rs, iCount)
      HeaderText = MapText("ใบสำคัญรับ")
      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
   End If

   If Not Report Is Nothing Then
      Call Report.AddParam(lMenuChosen, "REPORT_TYPE")
      Call Report.AddParam(m_BillingDoc.BILLING_DOC_ID, "BILLING_DOC_ID")
      Call Report.AddParam(ReportKey, "REPORT_KEY")
      Call Report.AddParam(MapText("ใบเสร็จรับเงิน"), "REPORT_HEADER")
      Call Report.AddParam(Picture2.Picture, "BACK_GROUND")
'      Call Report.AddParam(uctlSellByLookup.MyCombo.Text, "RECEIVE_NAME")
      Call Report.AddParam("", "ACCEPT_NAME")
   ElseIf lMenuChosen = 5 Then
      ReportKey = "CReportFormPO001"
      
      Set Rc = New CReportConfig
      Rc.REPORT_KEY = ReportKey
      Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
      Call Rc.QueryData(m_Rs, iCount)
      HeaderText = MapText("ใบเสร็จรับเงิน")
      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
   End If
   
   Call EnableForm(Me, False)
   If ReportFlag Then
      Set frmReport.ReportObject = Report
      frmReport.HeaderText = ""
      Load frmReport
      frmReport.Show 1
         
      Unload frmReport
      Set frmReport = Nothing
      Set Report = Nothing
   Else
      frmReportConfig.ReportMode = ReportMode
      frmReportConfig.ShowMode = EditMode
      frmReportConfig.ID = Rc.REPORT_CONFIG_ID
      frmReportConfig.ReportKey = ReportKey
      frmReportConfig.HeaderText = HeaderText
      Load frmReportConfig
      frmReportConfig.Show 1
      
      Unload frmReportConfig
      Set frmReportConfig = Nothing
   End If
   Call EnableForm(Me, True)
End Sub

Private Sub cmdPvNo_Click()
   'debug.print
   
End Sub

Private Sub cmdSave_Click()
Dim Result As Boolean
   If Not SaveData Then
      Exit Sub
   End If
   
   ShowMode = SHOW_EDIT
   ID = m_ChequeDoc.CHEQUE_DOC_ID
   m_ChequeDoc.QueryFlag = 1
   QueryData (True)
   m_HasModify = False
End Sub
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
'      Call LoadEnterpriseAddress(cboEnpAddress, , , True)
      
'      Call InitPaymentType(cboPaymentType)
'      Call LoadBank(cboBank)
      
'      If Area = 1 Then
         Call LoadCustomer(uctlCustomerLookup.MyCombo, m_Customers)
         Set uctlCustomerLookup.MyCollection = m_Customers
'      ElseIf Area = 2 Then
'         Call LoadSupplier(uctlCustomerLookup.MyCombo, m_Customers)
'         Set uctlCustomerLookup.MyCollection = m_Customers
'      End If
      
'      Call LoadEmployee(uctlSellByLookup.MyCombo, m_Employees)
'      Set uctlSellByLookup.MyCollection = m_Employees
'
'      Call LoadResource(uctlResource.MyCombo, m_Resources)
'      Set uctlResource.MyCollection = m_Resources
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_ChequeDoc.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         uctlChequeDocDate.ShowDate = Now
         m_ChequeDoc.QueryFlag = 0
         Call QueryData(False)
      End If
      
      Call EnableForm(Me, True)
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
   
   Set m_BillingDoc = Nothing
   Set m_Customers = Nothing
   Set m_Employees = Nothing
   Set m_Resources = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   Debug.Print ColIndex & " " & NewColWidth
End Sub

'Private Sub InitGrid2()
'Dim Col As JSColumn
'
'   GridEX1.Columns.Clear
'   GridEX1.BackColor = GLB_GRID_COLOR
'   GridEX1.ItemCount = 0
'   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
'   GridEX1.ColumnHeaderFont.Bold = True
'   GridEX1.ColumnHeaderFont.Name = GLB_FONT
'   GridEX1.TabKeyBehavior = jgexControlNavigation
'
'   Set Col = GridEX1.Columns.add '1
'   Col.Width = 0
'   Col.Caption = "ID"
'
'   Set Col = GridEX1.Columns.add '2
'   Col.Width = 0
'   Col.Caption = "Real ID"
'
'   Set Col = GridEX1.Columns.add '3
'   Col.Width = 2325 + 2055 + 2235
'   Col.Caption = MapText("รายละเอียด")
'
'   Set Col = GridEX1.Columns.add '4
'   Col.Width = 1620
'   Col.Caption = MapText("จำนวน")
'
'   Set Col = GridEX1.Columns.add '5
'   Col.TextAlignment = jgexAlignRight
'   Col.Width = 1575
'   Col.Caption = MapText("ราคารวม")
'
'   Set Col = GridEX1.Columns.add '6
'   Col.TextAlignment = jgexAlignRight
'   Col.Width = 1755
'   Col.Caption = MapText("ราคา/หน่วย")
'End Sub

Private Sub InitGrid1()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.Name = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   Set Col = GridEX1.Columns.add '3
   Col.Width = 2415
   Col.Caption = MapText("เลขที่เอกสาร")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 2250
   Col.Caption = MapText("วันที่เอกสาร")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 2460
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("จำนวนเงิน")
   
   Set Col = GridEX1.Columns.add '6
   Col.Width = 2250
   Col.Caption = MapText("ประเภทเอกสาร")
   
   Set Col = GridEX1.Columns.add '7
   Col.Visible = False
   Col.Caption = MapText("DO_ID")

   Set Col = GridEX1.Columns.add '8
   Col.Width = 1920
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ส่วนลดเงินสด")
End Sub

Private Sub GetTotalPrice()
Dim II As CReceiptItem
Dim Sum1 As Double
Dim Sum2 As Double
Dim Sum3 As Double
Dim Sum4 As Double
Dim Sum7 As Double
Dim Pm As CCashTran

   Sum1 = 0
   Sum2 = 0
   Sum3 = 0
   Sum4 = 0
   For Each II In m_BillingDoc.ReceiptItems
      If II.Flag <> "D" Then
         Sum1 = Sum1 + II.PAID_AMOUNT
         Sum2 = Sum2 + II.VAT_AMOUNT
         Sum3 = Sum3 + II.DISCOUNT_AMOUNT
         Sum4 = Sum4 + II.DEPOSIT_AMOUNT
      End If
   Next II
   
   Sum7 = 0
   For Each Pm In m_BillingDoc.Payments
      Sum7 = Sum7 + Pm.GetFieldValue("AMOUNT") - Pm.GetFieldValue("INTERREST_PAY") + Pm.GetFieldValue("WH_PAY")
   Next Pm
   
'   txtNetTotal.Text = Format(Sum1, "0.00")
'   txtVatAmount.Text = Format(Sum2, "0.00")
'   txtDiscount.Text = Format(Sum3, "0.00")
'   txtTotalRcp.Text = Format(Sum7, "0.00")
End Sub

Private Sub GetTotalPriceEx()
Dim II As CDoItem
Dim Sum1 As Double
Dim Sum2 As Double
Dim Sum3 As Double

   Sum2 = 0
   Sum1 = 0
   For Each II In m_BillingDoc.DoItems
      If II.Flag <> "D" Then
         Sum1 = Sum1 + II.TOTAL_PRICE
         Sum2 = Sum2 + II.DISCOUNT_AMOUNT
         Sum3 = Sum3 + II.DEPOSIT_AMOUNT
      End If
   Next II

'   txtNetTotal.Text = Format(Sum1, "0.00")
'   txtDiscount.Text = Format(Sum2, "0.00")
   
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame2.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame3.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption

   Call InitNormalLabel(lblChequeDocNo, MapText("เลขที่ใบเช็ค"))
'   Call InitNormalLabel(lblAccountNo, MapText("เลขที่บัญชี"))
'   If Area = 1 Then
'      Call InitNormalLabel(lblCustomerAddress, MapText("ที่อยู่ลูกค้า"))
      Call InitNormalLabel(lblCustomer, MapText("รหัสลูกค้า"))
'      Call InitNormalLabel(lblEnpAddress, MapText("ที่อยู่ออกเอกสาร"))
'      Call InitNormalLabel(lblSellBy, MapText("ผู้ออกใบเสร็จ"))
'   ElseIf Area = 2 Then
'      Call InitNormalLabel(lblCustomerAddress, MapText("ที่อยู่ซัพ ฯ"))
'      Call InitNormalLabel(lblCustomer, MapText("รหัสซัพ ฯ"))
'      Call InitNormalLabel(lblEnpAddress, MapText("ที่อยู่รับเอกสาร"))
'      Call InitNormalLabel(lblSellBy, MapText("ผู้รับเอกสาร"))
      cmdAuto.Visible = False
      cmdCustomer.Visible = True
'      If DocumentType = 8 Then
'         cmdPrint.Enabled = True
'      Else
'         cmdPrint.Enabled = False
'      End If
'   End If
   Call InitNormalLabel(lblChequeDocDate, MapText("วันที่เช็ค"))
   Call InitNormalLabel(lblBadChequeDate, MapText("วันที่เช็คไม่ผ่าน"))
    Call InitNormalLabel(lblPassChequeDate, MapText("วันที่เช็คผ่าน"))
   
'   Call InitNormalLabel(Label4, MapText("บาท"))
'   Call InitNormalLabel(lblNetTotal, MapText("ราคารวม"))
'   Call InitNormalLabel(lblVat, MapText("ภาษีมูลค่าเพิ่ม"))
'   Call InitNormalLabel(lblWH, MapText("ภาษีหัก ณ ที่จ่าย"))
'   Call InitNormalLabel(lblWHAmount, MapText("มูลค่า WH"))
'   Call InitNormalLabel(lblVatAmount, MapText("มูลค่า VAT"))
'   Call InitNormalLabel(lblIncludeVat, MapText("ยอดรวม VAT"))
'   Call InitNormalLabel(lblIncludeWH, MapText("ยอดรวม WH"))
'   Call InitNormalLabel(Label1, MapText("%"))
'   Call InitNormalLabel(Label3, MapText("%"))
'   Call InitNormalLabel(Label2, MapText("บาท"))
'   Call InitNormalLabel(Label5, MapText("บาท"))
'   Call InitNormalLabel(Label10, MapText("บาท"))
'   Call InitNormalLabel(Label12, MapText("บาท"))
'   Call InitNormalLabel(lblDiscount, MapText("ส่วนลด"))
   
'   Call InitNormalLabel(lblIncludeDiscount, MapText("รวมส่วนลด"))
   
'   Call InitNormalLabel(lblPvNo, MapText("PV NO"))
'   Call InitNormalLabel(lblPaidDate, MapText("กำหนดจ่าย"))
'   Call InitNormalLabel(lblPaidFor, MapText("เป็นการชำระค่า"))
   
'   Call InitNormalLabel(Label6, MapText("บาท"))
'   Call InitNormalLabel(Label8, MapText("บาท"))
'   Call InitNormalLabel(Label11, MapText("บาท"))
   
'   Call InitNormalLabel(Label14, MapText("บาท"))
      
'   Call InitNormalLabel(lblPaymentType, MapText("การชำระเงิน"))
'   Call InitNormalLabel(lblCheckNo, MapText("เลขที่เช็ค"))
'   Call InitNormalLabel(lblCheckDate, MapText("วันที่เช็ค"))
'   Call InitNormalLabel(lblBank, MapText("ธนาคาร"))
'   Call InitNormalLabel(lblBankBranch, MapText("สาขา"))
'   Call InitNormalLabel(lblTotalRcp, MapText("ยอดชำะจริง"))
'   Call InitNormalLabel(lblDipRcp, MapText("ส่วนต่างชำระ"))
   
'   Call InitNormalLabel(lblResource, MapText("ทรัพยากร"))
   
   Call InitCheckBox(chkBadCheque, "เช็คผ่าน")
   Call InitCheckBox(chkPassCheque, "เช็คไม่ผ่าน")
   
   If Area = 1 Then
'      lblAccountNo.Visible = True
'      cboAccount.Visible = True
      chkPassCheque.Visible = True
      chkBadCheque.Visible = True
   ElseIf Area = 2 Then
'      lblAccountNo.Visible = False
'      cboAccount.Visible = False
'      chkPayFlag.Visible = True
   End If
   
   Call txtChequeDocNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
'   Call txtNetTotal.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
'   txtNetTotal.Enabled = False
'   Call txtVatPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
'   If ReceiptType = 5 Then
''      txtVatPercent.Enabled = False
'   End If
'   Call txtWH.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
'   Call txtVatAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
'   txtVatAmount.Enabled = False
'   Call txtWHAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
'   txtWHAmount.Enabled = False
'   Call txtIncludeVat.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
'   txtIncludeVat.Enabled = False
'   Call txtIncludeWH.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
'   txtIncludeWH.Enabled = False
'   Call txtTotalRcp.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
'   txtTotalRcp.Enabled = False
'   Call txtDipRcp.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
'   txtDipRcp.Enabled = False
   
'   Call txtIncludeDiscount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
'   txtIncludeDiscount.Enabled = False
'   Call txtDiscount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
'   txtDiscount.Enabled = False
   
'   Call txtCheckNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   GridEX1.Visible = True
   SSFrame2.Visible = False
   SSFrame3.Visible = False
   
'   Call InitCombo(cboAccount)
'   Call InitCombo(cboCustomerAddress)
'   Call InitCombo(cboEnpAddress)
   Call InitCombo(cboPaymentType)
   Call InitCombo(cboBank)
   Call InitCombo(cboBankBranch)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
'   cmdPrint.Picture = LoadPicture(glbParameterObj.NormalButton1)
'   cmdSave.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAuto.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdCustomer.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPvNo.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
'   Call InitMainButton(cmdSave, MapText("บันทึก"))
'   Call InitMainButton(cmdPrint, MapText("พิมพ์"))
   Call InitMainButton(cmdAuto, MapText("A"))
   Call InitMainButton(cmdCustomer, MapText("F"))
   Call InitMainButton(cmdPvNo, MapText("P"))
   
'   If (ReceiptType = 3) Or (ReceiptType = 5) Then
'      Call InitGrid1
'   ElseIf ReceiptType = 1 Then
'      Call InitGrid2
'   End If
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.Name = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.add().Caption = MapText("รายการใบเสร็จ")
'   TabStrip1.Tabs.add().Caption = MapText("รายละเอียดทั่วไป")
'   TabStrip1.Tabs.add().Caption = MapText("การชำระเงิน")
'   TabStrip1.Tabs.add().Caption = MapText("สมุดรายวัน")

 Call InitGrid1

   
'   If (ReceiptType = 3) Or (ReceiptType = 5) Then
'      cmdEdit.Enabled = False
'   End If
'   If DocumentType = 8 Then
''      cmdPrint.Enabled = True
'   End If
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
   Set m_BillingDoc = New CBillingDoc
  Set m_ChequeDoc = New CChequeDoc
   Set m_Customers = New Collection
   Set m_Employees = New Collection
   Set m_Resources = New Collection
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

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Index = 1 Then
      If m_BillingDoc.ReceiptItems Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      If (ReceiptType = 3) Or (ReceiptType = 5) Then
         Dim CR As CReceiptItem
         If m_BillingDoc.ReceiptItems.Count <= 0 Then
            Exit Sub
         End If
         Set CR = GetItem(m_BillingDoc.ReceiptItems, RowIndex, RealIndex)
         If CR Is Nothing Then
            Exit Sub
         End If
   
         Values(1) = CR.RECEIPT_ITEM_ID
         Values(2) = RealIndex
         Values(3) = CR.DOCUMENT_NO
         Values(4) = DateToStringExtEx2(CR.DOCUMENT_DATE)
         Values(5) = FormatNumber(CR.PAID_AMOUNT)
         If ReceiptType = 3 Then
            Values(6) = "ใบส่งสินค้า"
         ElseIf ReceiptType = 5 Then
            Values(6) = "ใบกำกับภาษี"
         End If
         Values(7) = CR.DO_ID
         Values(8) = FormatNumber(CR.CASH_DISCOUNT)
      ElseIf ReceiptType = 1 Then
         Dim Di As CDoItem
         If m_BillingDoc.DoItems.Count <= 0 Then
            Exit Sub
         End If
         Set Di = GetItem(m_BillingDoc.DoItems, RowIndex, RealIndex)
         If Di Is Nothing Then
            Exit Sub
         End If
   
         Values(1) = Di.DO_ITEM_ID
         Values(2) = RealIndex
         Values(3) = Di.ShowDescText
         Values(4) = FormatNumber(Di.ITEM_AMOUNT)
         Values(5) = FormatNumber(Di.TOTAL_PRICE)
         Values(6) = FormatNumber(Di.AVG_PRICE)
      End If 'ReceiptType
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
      If m_BillingDoc.Payments Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim Ct As CCashTran
      If m_BillingDoc.Payments.Count <= 0 Then
         Exit Sub
      End If
      Set Ct = GetItem(m_BillingDoc.Payments, RowIndex, RealIndex)
      If Ct Is Nothing Then
         Exit Sub
      End If

      Values(1) = Ct.GetFieldValue("CASH_TRAN_ID")
      Values(2) = RealIndex
      Values(3) = Ct.GetFieldValue("PAYMENT_TYPE_NAME")
      If Ct.GetFieldValue("PAYMENT_TYPE") = CASH_PMT Then
         Values(7) = FormatNumber(Ct.GetFieldValue("AMOUNT"))
      ElseIf Ct.GetFieldValue("PAYMENT_TYPE") = CREDITCRD_PMT Then
         Values(4) = Ct.GetFieldValue("ACCOUNT_NAME")
         Values(5) = Ct.GetFieldValue("BANK_NAME")
         Values(6) = Ct.GetFieldValue("BRANCH_NAME")
         Values(7) = FormatNumber(Ct.GetFieldValue("AMOUNT"))
      ElseIf Ct.GetFieldValue("PAYMENT_TYPE") = CHECK_PMT Then
         Values(4) = Ct.Cheque.GetFieldValue("CHEQUE_NO")
         Values(5) = Ct.Cheque.GetFieldValue("BANK_NAME")
         Values(6) = Ct.Cheque.GetFieldValue("BRANCH_NAME")
         Values(7) = FormatNumber(Ct.GetFieldValue("AMOUNT"))
      End If
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
      If m_BillingDoc.GlDetails Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim Gl As CGLDetail
      If m_BillingDoc.GlDetails.Count <= 0 Then
         Exit Sub
      End If
      Set Gl = GetItem(m_BillingDoc.GlDetails, RowIndex, RealIndex)
      If Gl Is Nothing Then
         Exit Sub
      End If

      Values(1) = Gl.GetFieldValue("GL_DETAIL_ID")
      Values(2) = RealIndex
      Values(3) = Gl.GetFieldValue("GL_NO")
      Values(4) = Gl.GetFieldValue("GL_NAME")
      Values(5) = Gl.GetFieldValue("GL_DESC")
      If Gl.GetFieldValue("GL_TYPE") = 1 Then
         Values(6) = FormatNumber(Gl.GetFieldValue("GL_AMOUNT"))
         Values(7) = ""
      Else
         Values(6) = ""
         Values(7) = FormatNumber(Gl.GetFieldValue("GL_AMOUNT"))
      End If
      
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
      
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub EnableDisableButton(En As Boolean)
   If En Then
'      If ShowMode = SHOW_EDIT Then
'         cmdAdd.Enabled = (m_BillingDoc.OLD_COMMIT_FLAG = "N")
'         cmdEdit.Enabled = (m_BillingDoc.COMMIT_FLAG = "N")
'         cmdDelete.Enabled = (m_BillingDoc.OLD_COMMIT_FLAG = "N")
'      Else
'         cmdAdd.Enabled = True
'         cmdDelete.Enabled = True
'      End If
'      If ((ReceiptType = 3) Or (ReceiptType = 5)) And Not (TabStrip1.SelectedItem.Index = 3 Or TabStrip1.SelectedItem.Index = 4) Then
'         cmdEdit.Enabled = False
'      Else
'         cmdEdit.Enabled = True
'      End If
'   Else
'      cmdAdd.Enabled = En
'      cmdDelete.Enabled = En
'      cmdEdit.Enabled = En
'
      cmdAdd.Enabled = En
      cmdDelete.Enabled = En
      cmdEdit.Enabled = En
   End If
End Sub

Private Sub TabStrip1_Click()
'   GridEX1.Top = 5670
'   GridEX1.Left = 150
'   GridEX1.Visible = False
'
'   SSFrame2.Top = 5670
'   SSFrame2.Left = 150
'   SSFrame2.Visible = False
'
'   SSFrame3.Top = 5670
'   SSFrame3.Left = 150
'   SSFrame3.Visible = False
   
'   If TabStrip1.SelectedItem.Index = 1 Then
'      Call EnableDisableButton(True)
'      GridEX1.Visible = True
'      If (ReceiptType = 3) Or (ReceiptType = 5) Then
'         Call GetTotalPrice
'         Call InitGrid1
'         GridEX1.ItemCount = CountItem(m_BillingDoc.ReceiptItems)
'         GridEX1.Rebind
'      ElseIf ReceiptType = 1 Then
'         Call GetTotalPriceEx
'         Call InitGrid2
'         GridEX1.ItemCount = CountItem(m_BillingDoc.DoItems)
'         GridEX1.Rebind
'      End If
'   ElseIf TabStrip1.SelectedItem.Index = 2 Then
'      Call EnableDisableButton(False)
'      SSFrame2.Visible = True
'   ElseIf TabStrip1.SelectedItem.Index = 3 Then
'      Call EnableDisableButton(True)
'      Call InitGrid3
'      GridEX1.Visible = True
'
'      Call GetTotalPrice
'      GridEX1.ItemCount = CountItem(m_BillingDoc.Payments)
'      GridEX1.Rebind
'   ElseIf TabStrip1.SelectedItem.Index = 4 Then
'      Call EnableDisableButton(True)
'      Call InitGrid4
'      GridEX1.Visible = True
'
'      'Call GetTotalPrice
'      GridEX1.ItemCount = CountItem(m_BillingDoc.GlDetails)
'      GridEX1.Rebind
'
'   ElseIf TabStrip1.SelectedItem.Index = 5 Then
'   End If
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

Private Sub txtCheckNo_Change()
   m_HasModify = True
End Sub

Private Sub txtDeposit_Change()
   m_HasModify = True
   Call CalculateAmount
End Sub

Private Sub txtDiscount_Change()
   m_HasModify = True
   Call CalculateAmount
End Sub

Private Sub txtDocumentNo_Change()
'   txtPvNo.Text = txtDocumentNo.Text
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

Private Sub txtIncludeDiscount_Change()
   m_HasModify = True
   Call CalculateAmount
End Sub

Private Sub txtIncludeVat_Change()
   m_HasModify = True
   Call CalculateAmount
End Sub

Private Sub txtIncludeWH_Change()
   m_HasModify = True
   Call CalculateAmount
End Sub

Private Sub txtLeft_Change()
   m_HasModify = True
End Sub

Private Sub txtNetTotal_Change()
   m_HasModify = True
   Call CalculateAmount
End Sub

Private Sub txtChequeDocNo_Change()
  m_HasModify = True
End Sub

Private Sub txtPaidFor_Change()
   m_HasModify = True
End Sub

Private Sub txtPvNo_Change()
   m_HasModify = True
End Sub

Private Sub txtTotalRcp_Change()
   m_HasModify = True
   Call CalculateAmount
End Sub

Private Sub txtVatAmount_Change()
   m_HasModify = True
   Call CalculateAmount
End Sub

Private Sub CalculateAmount()
'   txtIncludeDiscount.Text = Val(txtNetTotal.Text) - Val(Replace(txtDiscount.Text, ",", ""))
   If (ReceiptType <> 5) Then
'      txtVatAmount.Text = Val(txtVatPercent.Text) * Val(Replace(txtIncludeDiscount.Text, ",", "")) / 100
   End If
'   txtIncludeVat.Text = Val(Replace(txtIncludeDiscount.Text, ",", "")) + Val(txtVatAmount.Text)
'   txtWHAmount.Text = Val(txtWH.Text) * Val(txtIncludeDiscount.Text) / 100
'   txtIncludeWH.Text = Val(txtIncludeVat.Text) - txtWHAmount.Text
'   txtDipRcp.Text = Val(txtIncludeWH.Text) - Val(txtTotalRcp.Text)
End Sub

Private Sub txtVatPercent_Change()
   m_HasModify = True
   Call CalculateAmount
End Sub

Private Sub txtWH_Change()
   m_HasModify = True
   Call CalculateAmount
End Sub

Private Sub txtWHAmount_Change()
   m_HasModify = True
End Sub

Private Sub uctlBadChequeDate_HasChange()
  m_HasModify = True
End Sub

Private Sub uctlCheckDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlDocumentDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlChequeDocDate_HasChange()
  m_HasModify = True
  
End Sub

Private Sub uctlCustomerLookup_Change()
Dim CustomerID As Long
Dim Customer As CCustomer

   CustomerID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
   CUSTOMER_ID = CustomerID
   If CustomerID > 0 Then
'      If Area = 1 Then
         Set Customer = m_Customers(Trim(Str(CustomerID)))
'         Call LoadAccount(cboAccount, , CustomerID)
'         cboAccount.ListIndex = 1
   
'         Call LoadCustomerAddress(cboCustomerAddress, , CustomerID, True)
'         If Customer.RESPONSE_BY > 0 Then
'            uctlSellByLookup.MyCombo.ListIndex = IDToListIndex(uctlSellByLookup.MyCombo, Customer.RESPONSE_BY)
'         Else
'            uctlSellByLookup.MyCombo.ListIndex = -1
'         End If
'      ElseIf Area = 2 Then
''         Call LoadAccount(cboAccount, , CustomerID)
''         cboAccount.ListIndex = -1
'
'         Call LoadSupplierAddress(cboCustomerAddress, , CustomerID, True)
'      End If
   Else
'      cboAccount.ListIndex = -1
'      cboCustomerAddress.ListIndex = -1
   End If
   m_HasModify = True
End Sub

Private Sub uctlPaidDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlPassChequeDate_HasChange()
  m_HasModify = True
End Sub

Private Sub uctlResource_Change()
   m_HasModify = True
End Sub

Private Sub uctlSellByLookup_Change()
   m_HasModify = True
End Sub
'Private Sub InitGrid3()
'Dim Col As JSColumn
'
'   GridEX1.Columns.Clear
'   GridEX1.BackColor = GLB_GRID_COLOR
'   GridEX1.ItemCount = 0
'   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
'   GridEX1.ColumnHeaderFont.Bold = True
'   GridEX1.ColumnHeaderFont.Name = GLB_FONT
'   GridEX1.TabKeyBehavior = jgexControlNavigation
'
'   Set Col = GridEX1.Columns.add '1
'   Col.Width = 0
'   Col.Caption = "ID"
'
'   Set Col = GridEX1.Columns.add '2
'   Col.Width = 0
'   Col.Caption = "Real ID"
'
'      Set Col = GridEX1.Columns.add '3
'      Col.Width = 1965
'      Col.Caption = MapText("ประเภทการชำระเงิน")
'
'      Set Col = GridEX1.Columns.add '4
'      Col.Width = 2625
'      Col.Caption = MapText("เลขที่เช็ค/บัญชี")
'
'      Set Col = GridEX1.Columns.add '5
'      Col.Width = 2160
'      Col.TextAlignment = jgexAlignLeft
'      Col.Caption = MapText("ธนาคาร")
'
'      Set Col = GridEX1.Columns.add '6
'      Col.Width = 2565
'      Col.TextAlignment = jgexAlignLeft
'      Col.Caption = MapText("สาขาธนาคาร")
'
'      Set Col = GridEX1.Columns.add '7
'      Col.Width = 2000
'      Col.TextAlignment = jgexAlignRight
'      Col.Caption = MapText("จำนวนเงิน")
'
'End Sub
'Private Sub InitGrid4()
'Dim Col As JSColumn
'
'   GridEX1.Columns.Clear
'   GridEX1.BackColor = GLB_GRID_COLOR
'   GridEX1.ItemCount = 0
'   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
'   GridEX1.ColumnHeaderFont.Bold = True
'   GridEX1.ColumnHeaderFont.Name = GLB_FONT
'   GridEX1.TabKeyBehavior = jgexControlNavigation
'
'   Set Col = GridEX1.Columns.add '1
'   Col.Width = 0
'   Col.Caption = "ID"
'
'   Set Col = GridEX1.Columns.add '2
'   Col.Width = 0
'   Col.Caption = "Real ID"
'
'      Set Col = GridEX1.Columns.add '3
'      Col.Width = 1500
'      Col.Caption = MapText("เลขที่บัญชี")
'
'      Set Col = GridEX1.Columns.add '4
'      Col.Width = 3000
'      Col.Caption = MapText("ชื่อบัญชี")
'
'      Set Col = GridEX1.Columns.add '5
'      Col.Width = 3000
'      Col.Caption = MapText("รายละเอียด")
'
'      Set Col = GridEX1.Columns.add '6
'      Col.Width = 2000
'      Col.TextAlignment = jgexAlignRight
'      Col.Caption = MapText("Dr.")
'
'      Set Col = GridEX1.Columns.add '7
'      Col.Width = 2000
'      Col.TextAlignment = jgexAlignRight
'      Col.Caption = MapText("Cr.")
'
'End Sub

