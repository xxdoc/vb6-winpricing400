VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAddEditBillingPayment 
   ClientHeight    =   9570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11775
   ForeColor       =   &H00000000&
   Icon            =   "frmAddEditBillingPayment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   11775
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   9840
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   17357
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000009&
         Height          =   1275
         Left            =   360
         ScaleHeight     =   1215
         ScaleWidth      =   1575
         TabIndex        =   53
         Top             =   -120
         Visible         =   0   'False
         Width           =   1635
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   120
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   4800
         Width           =   13755
         _ExtentX        =   24262
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
      Begin prjFarmManagement.uctlDate uctlDocumentDate 
         Height          =   405
         Left            =   7440
         TabIndex        =   2
         Top             =   840
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtPaymentTo 
         Height          =   435
         Left            =   3000
         TabIndex        =   6
         Top             =   1770
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtDocumentNo 
         Height          =   435
         Left            =   3000
         TabIndex        =   0
         Top             =   840
         Width           =   2295
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPaymentCost 
         Height          =   435
         Left            =   3000
         TabIndex        =   9
         Top             =   2700
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
      Begin GridEX20.GridEX GridEX1 
         Height          =   3375
         Left            =   120
         TabIndex        =   18
         Top             =   5280
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   5953
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
         Column(1)       =   "frmAddEditBillingPayment.frx":27A2
         Column(2)       =   "frmAddEditBillingPayment.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditBillingPayment.frx":290E
         FormatStyle(2)  =   "frmAddEditBillingPayment.frx":2A6A
         FormatStyle(3)  =   "frmAddEditBillingPayment.frx":2B1A
         FormatStyle(4)  =   "frmAddEditBillingPayment.frx":2BCE
         FormatStyle(5)  =   "frmAddEditBillingPayment.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditBillingPayment.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Width           =   14085
         _ExtentX        =   24844
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtPaymentDept2 
         Height          =   435
         Left            =   9600
         TabIndex        =   14
         Top             =   3180
         Width           =   1995
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPaymentDept 
         Height          =   435
         Left            =   6600
         TabIndex        =   13
         Top             =   3180
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtNote 
         Height          =   435
         Left            =   3000
         TabIndex        =   16
         Top             =   4080
         Width           =   9825
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPaymentPart 
         Height          =   435
         Left            =   3000
         TabIndex        =   12
         Top             =   3180
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlPaymentDue 
         Height          =   405
         Left            =   7440
         TabIndex        =   8
         Top             =   2280
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin Threed.SSFrame SSFrame4 
         Height          =   3255
         Left            =   120
         TabIndex        =   44
         Top             =   5400
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   5741
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin VB.ComboBox cboCondition 
            Height          =   315
            Left            =   2760
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   240
            Width           =   4035
         End
         Begin VB.ComboBox cboPaidType 
            Height          =   315
            Left            =   2760
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Top             =   720
            Width           =   4005
         End
         Begin VB.Label lblPaidType 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   840
            TabIndex        =   47
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label lblCondition 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   360
            TabIndex        =   48
            Top             =   240
            Width           =   2295
         End
      End
      Begin prjFarmManagement.uctlTextBox txtDocAssembly 
         Height          =   435
         Left            =   3000
         TabIndex        =   7
         Top             =   2235
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPaymentAmount 
         Height          =   435
         Left            =   3000
         TabIndex        =   15
         Top             =   3600
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlDocumentDate2 
         Height          =   405
         Left            =   7440
         TabIndex        =   5
         Top             =   1320
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtDocumentNo2 
         Height          =   435
         Left            =   3000
         TabIndex        =   3
         Top             =   1320
         Width           =   2295
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin Threed.SSOption sspCheque 
         Height          =   255
         Left            =   9600
         TabIndex        =   11
         Top             =   2760
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   131073
         Caption         =   "sspCheque"
      End
      Begin Threed.SSOption sspCash 
         Height          =   255
         Left            =   7440
         TabIndex        =   10
         Top             =   2760
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   131073
         Caption         =   "sspCash"
      End
      Begin VB.Label lblDocumentNo2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   450
         TabIndex        =   52
         Top             =   1380
         Width           =   2505
      End
      Begin VB.Label lblDocumentDate2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6150
         TabIndex        =   51
         Top             =   1350
         Width           =   1155
      End
      Begin Threed.SSCommand cmdAuto2 
         Height          =   405
         Left            =   5310
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1320
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditBillingPayment.frx":2F36
         ButtonStyle     =   3
      End
      Begin VB.Label lblPaymentAmount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   240
         TabIndex        =   50
         Top             =   3600
         Width           =   2655
      End
      Begin Threed.SSCheck SSCheck1 
         Height          =   375
         Left            =   7560
         TabIndex        =   25
         Top             =   2280
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "chkCash"
      End
      Begin VB.Label lblPaymentBy 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6120
         TabIndex        =   49
         Top             =   2760
         Width           =   1125
      End
      Begin VB.Label Label5 
         Height          =   315
         Left            =   10440
         TabIndex        =   43
         Top             =   3120
         Width           =   585
      End
      Begin VB.Label lblPaymentPart 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   42
         Top             =   3180
         Width           =   2685
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   525
         Left            =   6720
         TabIndex        =   22
         Top             =   8880
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditBillingPayment.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAuto 
         Height          =   405
         Left            =   5310
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   840
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditBillingPayment.frx":356A
         ButtonStyle     =   3
      End
      Begin VB.Label lblNote 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   360
         TabIndex        =   41
         Top             =   4080
         Width           =   2535
      End
      Begin VB.Label Label4 
         Height          =   315
         Left            =   10950
         TabIndex        =   40
         Top             =   3660
         Width           =   585
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8280
         TabIndex        =   39
         Top             =   3660
         Width           =   1125
      End
      Begin VB.Label lblBath 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5880
         TabIndex        =   38
         Top             =   3690
         Width           =   855
      End
      Begin VB.Label Label2 
         Height          =   315
         Left            =   7440
         TabIndex        =   37
         Top             =   3660
         Width           =   765
      End
      Begin VB.Label Label1 
         Height          =   315
         Left            =   3900
         TabIndex        =   36
         Top             =   3660
         Width           =   765
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6150
         TabIndex        =   35
         Top             =   870
         Width           =   1155
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8400
         TabIndex        =   23
         Top             =   8880
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditBillingPayment.frx":3884
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10080
         TabIndex        =   24
         Top             =   8880
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1800
         TabIndex        =   20
         Top             =   8880
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   120
         TabIndex        =   19
         Top             =   8880
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditBillingPayment.frx":3B9E
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3480
         TabIndex        =   21
         Top             =   8880
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditBillingPayment.frx":3EB8
         ButtonStyle     =   3
      End
      Begin VB.Label lblPaymentDept 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5910
         TabIndex        =   33
         Top             =   3180
         Width           =   735
      End
      Begin VB.Label lblPaymentCost 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   420
         TabIndex        =   32
         Top             =   2730
         Width           =   2535
      End
      Begin VB.Label lblDocAssembly 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   60
         TabIndex        =   31
         Top             =   2310
         Width           =   2895
      End
      Begin VB.Label lblDocumentNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1290
         TabIndex        =   30
         Top             =   900
         Width           =   1665
      End
      Begin VB.Label lblPaymentDept2 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   8640
         TabIndex        =   29
         Top             =   3180
         Width           =   885
      End
      Begin VB.Label lblPaymentDue 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   5910
         TabIndex        =   28
         Top             =   2310
         Width           =   1485
      End
      Begin VB.Label lblPaymentTo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   300
         TabIndex        =   27
         Top             =   1860
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmAddEditBillingPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_BillingPayment As CBillingPayment
Private m_Suppliers As Collection
Private m_Cd As Collection
Private DocAdd As Long

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public AutoGenPo As Boolean

Public ID As Long
Public ID2 As Long
Public DocumentType As Long

Private FileName As String
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
   
   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
      
      m_BillingPayment.BILLING_PAYMENT_ID = ID
      m_BillingPayment.BILLING_PAYMENT_ID_REF = ID2
      If Not glbDaily.QueryBillingPayment(m_BillingPayment, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
        Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_BillingPayment.PopulateFromRS(1, m_Rs)
  
      txtDocumentNo.Text = m_BillingPayment.DOCUMENT_NO
      uctlDocumentDate.ShowDate = m_BillingPayment.DOCUMENT_DATE
      txtDocumentNo2.Text = m_BillingPayment.DOCUMENT_NO_JV
      uctlDocumentDate2.ShowDate = m_BillingPayment.DOCUMENT_DATE_JV
      txtPaymentTo.Text = m_BillingPayment.PAYMENT_TO
      txtDocAssembly.Text = m_BillingPayment.DOC_ASSEMBLE
      txtPaymentCost.Text = m_BillingPayment.PAYMENT_COST
       uctlPaymentDue.ShowDate = m_BillingPayment.PAYMENT_DUE
       txtPaymentPart.Text = m_BillingPayment.PAYMENT_PART
       txtPaymentDept.Text = m_BillingPayment.PAYMENT_DEPT
       txtPaymentDept2.Text = m_BillingPayment.PAYMENT_DEPT2
       txtPaymentAmount.Text = Val(m_BillingPayment.PAYMENT_AMOUNT)
       If m_BillingPayment.PAYMENT_BY = 1 Then
         sspCash.Value = 1
         sspCheque.Value = 0
       ElseIf m_BillingPayment.PAYMENT_BY = 2 Then
          sspCash.Value = 0
          sspCheque.Value = 1
       End If
       txtNote.Text = m_BillingPayment.NOTE
   End If
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
'''   Call TabStrip1_Click
   
   Call EnableForm(Me, True)
End Sub


Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim Ivd As CInventoryDoc
Dim Sp As CSupItem
Dim Lt As CLotItem
Dim StrStockAmount As String
Dim TempDocNo As String
Dim firstDate As Date
Dim lastDate As Date
Dim MonthlyAccums  As Collection
Dim YYYYMM As String
Dim BalanceLi As CLotItem
Dim TempLi1 As CLotItem
Dim TempLi2 As CLotItem
Dim InventoryBals1  As Collection
               
   If Not VerifyTextControl(lblDocumentNo, txtDocumentNo, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblDocumentDate, uctlDocumentDate, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblPaymentCost, txtPaymentCost, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblPaymentAmount, txtPaymentAmount, True) Then
      Exit Function
   End If
   
If ShowMode = SHOW_ADD Then
      If Not CheckUniqueNs(BILLING_PAYMENT_NO_UNIQUE, txtDocumentNo.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtDocumentNo.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
         txtDocumentNo.Text = ""
         Call LoadConfigDoc(Nothing, m_Cd)
         Call cmdAuto_Click
         Exit Function
      End If
   
      
      If Not CheckUniqueNs(BILLING_PAYMENT_NO_UNIQUE, txtDocumentNo2.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtDocumentNo2.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
         txtDocumentNo2.Text = ""
         Call LoadConfigDoc(Nothing, m_Cd)
         Call cmdAuto2_Click
         Exit Function
      End If
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   

   
   
   m_BillingPayment.AddEditMode = ShowMode
   m_BillingPayment.BILLING_PAYMENT_ID = ID
   m_BillingPayment.DOCUMENT_NO = txtDocumentNo.Text
   m_BillingPayment.DOCUMENT_DATE = uctlDocumentDate.ShowDate
   m_BillingPayment.DOCUMENT_TYPE = DocumentType
   m_BillingPayment.PAYMENT_TO = txtPaymentTo.Text
   m_BillingPayment.DOC_ASSEMBLE = txtDocAssembly.Text
   m_BillingPayment.PAYMENT_COST = txtPaymentCost.Text
   m_BillingPayment.PAYMENT_DUE = uctlPaymentDue.ShowDate
   m_BillingPayment.PAYMENT_PART = txtPaymentPart.Text
   m_BillingPayment.PAYMENT_DEPT = txtPaymentDept.Text
   m_BillingPayment.PAYMENT_DEPT2 = txtPaymentDept2.Text
'   m_BillingPayment.PAYMENT_AMOUNT = Val(txtPaymentAmount.Text)
   m_BillingPayment.PAYMENT_AMOUNT = calSumAmount(m_BillingPayment)
   If Not m_BillingPayment.PAYMENT_AMOUNT > -1 Then
       glbErrorLog.LocalErrorMsg = MapText("ยอดรวม Debit และ Credit ไม่เท่ากัน")
      glbErrorLog.ShowUserError
      SaveData = False
      Exit Function
   End If
   
   If sspCash.Value Then
      m_BillingPayment.PAYMENT_BY = 1
   ElseIf sspCheque.Value Then
      m_BillingPayment.PAYMENT_BY = 2
   End If
   m_BillingPayment.NOTE = txtNote.Text
   
   Call EnableForm(Me, False)
   
   Call glbDaily.StartTransaction
   
   If Len(txtDocumentNo2.Text) Then
       m_BillingPayment.BILLING_PAYMENT_ID = ID2
       m_BillingPayment.BILLING_PAYMENT_ID_REF = -1
       m_BillingPayment.DOCUMENT_NO = txtDocumentNo2.Text
       m_BillingPayment.DOCUMENT_DATE = uctlDocumentDate2.ShowDate
       m_BillingPayment.DOCUMENT_TYPE = 112
      If Not glbDaily.AddEditBillingPayment(m_BillingPayment, IsOK, False, glbErrorLog, "JV") Then
            glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
            SaveData = False
            Call glbDaily.RollbackTransaction
            Call EnableForm(Me, True)
            Exit Function
      End If
      m_BillingPayment.BILLING_PAYMENT_ID_REF = m_BillingPayment.BILLING_PAYMENT_ID
   End If
   
   m_BillingPayment.BILLING_PAYMENT_ID = ID
   m_BillingPayment.DOCUMENT_NO = txtDocumentNo.Text
   m_BillingPayment.DOCUMENT_DATE = uctlDocumentDate.ShowDate
   m_BillingPayment.DOCUMENT_TYPE = DocumentType
   If Not glbDaily.AddEditBillingPayment(m_BillingPayment, IsOK, False, glbErrorLog, "PV") Then
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

Private Sub chkCash_Click(Value As Integer)

End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean
Dim oMenu  As cPopupMenu
Dim lMenuChosen As Long

   If Not cmdAdd.Enabled Then
      Exit Sub
   End If

   OKClick = False

   If TabStrip1.SelectedItem.Index = 1 Then
         Set frmAddEditGlDetail.ParentForm = Me
         frmAddEditGlDetail.ListType = TabStrip1.SelectedItem.Index
         frmAddEditGlDetail.HeaderText = "เพิ่มรายการใบสำคัญจ่าย (PV)"
         frmAddEditGlDetail.ShowMode = SHOW_ADD
         Set frmAddEditGlDetail.TempCollection = m_BillingPayment.GlDetails
         Load frmAddEditGlDetail
         frmAddEditGlDetail.Show 1
         
         OKClick = frmAddEditGlDetail.OKClick
         
         Unload frmAddEditGlDetail
         Set frmAddEditGlDetail = Nothing
         If OKClick Then
            m_HasModify = True
            
            GridEX1.ItemCount = CountItem(m_BillingPayment.GlDetails)
            Call GridEX1.Rebind
         End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
         Set frmAddEditGlDetail.ParentForm = Me
         frmAddEditGlDetail.ListType = TabStrip1.SelectedItem.Index
         frmAddEditGlDetail.HeaderText = "เพิ่มรายการใบสำคัญโอนบัญชี (JV)"
         frmAddEditGlDetail.ShowMode = SHOW_ADD
         Set frmAddEditGlDetail.TempCollection = m_BillingPayment.GlDetails2
         Load frmAddEditGlDetail
         frmAddEditGlDetail.Show 1
         
         OKClick = frmAddEditGlDetail.OKClick
         
         Unload frmAddEditGlDetail
         Set frmAddEditGlDetail = Nothing
         If OKClick Then
            m_HasModify = True
            
            GridEX1.ItemCount = CountItem(m_BillingPayment.GlDetails2)
            Call GridEX1.Rebind
         End If
   End If

   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdAuto_Click()
   If Trim(txtDocumentNo.Text) = "" And ShowMode = SHOW_ADD Then '
      txtDocumentNo.Text = GetDocumentNo(DocumentType)
   End If
End Sub

Private Function GetDocumentNo(DocNoType As Long) As String
Dim No As String
Dim DOC_ID As Long
Dim Cd As CConfigDoc
Dim TempStr As String
Dim I As Long
Dim ServerDateTime As String

   If DocNoType = 111 Then
      DOC_ID = PAYMENT_VOUCHER
   ElseIf DocNoType = 112 Then
      DOC_ID = TRANSFER_VOUCHER
   End If

    If DOC_ID > 0 Then
       Set Cd = GetObject("CConfigDoc", m_Cd, Trim(str(DOC_ID)), False)
       If Not (Cd Is Nothing) Then
          GetDocumentNo = Cd.GetFieldValue("PREFIX") & Cd.GetFieldValue("CODE1")
          TempStr = ""
          If Cd.GetFieldValue("YEAR_TYPE") = 1 Then
             TempStr = Right(Format(Year(Now) + 543, "0000"), 2)
          ElseIf Cd.GetFieldValue("YEAR_TYPE") = 2 Then
             TempStr = Format(Year(Now) + 543, "0000")
          ElseIf Cd.GetFieldValue("YEAR_TYPE") = 3 Then
             TempStr = Right(Format(Year(Now), "0000"), 2)
          ElseIf Cd.GetFieldValue("YEAR_TYPE") = 4 Then
             TempStr = Format(Year(Now), "0000")
          End If
          GetDocumentNo = GetDocumentNo & TempStr & Cd.GetFieldValue("CODE2")
          TempStr = ""
          If Cd.GetFieldValue("MONTH_TYPE") = 1 Then
             TempStr = Format(Month(Now), "00")
          End If
          GetDocumentNo = GetDocumentNo & TempStr & Cd.GetFieldValue("CODE3")
          TempStr = ""
          For I = 1 To Cd.GetFieldValue("DIGIT_AMOUNT")
             TempStr = TempStr & "0"
          Next I
           If Cd.GetFieldValue("AUTO_BEGIN_FLAG") = "Y" Then
               If CheckNewMounth And CheckUniqueNs(BILLING_PAYMENT_NO_UNIQUE, GetDocumentNo & Format(1, TempStr), ID) Then
                  GetDocumentNo = GetDocumentNo & Format(1, TempStr) 'เริ่มจาก 1 เสมอ
                  m_BillingPayment.RUNNING_NO = 1
               Else
                  GetDocumentNo = GetDocumentNo & Format(Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd, TempStr)
                  If DocNoType = 111 Then
                     m_BillingPayment.RUNNING_NO = Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd
                  ElseIf DocNoType = 112 Then
                     m_BillingPayment.RUNNING_NO2 = Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd
                  End If
               End If
          Else
               GetDocumentNo = GetDocumentNo & Format(Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd, TempStr)
                If DocNoType = 111 Then
                     m_BillingPayment.RUNNING_NO = Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd
                  ElseIf DocNoType = 112 Then
                     m_BillingPayment.RUNNING_NO2 = Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd
                  End If
          End If
         If DocNoType = 111 Then
            m_BillingPayment.CONFIG_DOC_TYPE = DOC_ID
         ElseIf DocNoType = 112 Then
            m_BillingPayment.CONFIG_DOC_TYPE2 = DOC_ID
         End If
          
       Else
          GetDocumentNo = ""
       End If
    End If
      
End Function

Private Sub cmdAuto2_Click()
   If Trim(txtDocumentNo2.Text) = "" And Trim(txtDocumentNo.Text) <> "" And ShowMode = SHOW_ADD Then  '
      txtDocumentNo2.Text = Replace(txtDocumentNo.Text, "PV", "JV")
      'txtDocumentNo2.Text = GetDocumentNo(112)
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
         m_BillingPayment.GlDetails.Remove (ID2)
      Else
         m_BillingPayment.GlDetails.Item(ID2).Flag = "D"
      End If
      
      GridEX1.ItemCount = CountItem(m_BillingPayment.GlDetails)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      If ID1 <= 0 Then
         m_BillingPayment.GlDetails2.Remove (ID2)
      Else
         m_BillingPayment.GlDetails2.Item(ID2).Flag = "D"
      End If
      
      GridEX1.ItemCount = CountItem(m_BillingPayment.GlDetails2)
      GridEX1.Rebind
      m_HasModify = True
   End If
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean

'   If Not VerifyAccessRight("LEDGER_BUY" & "_" & DocumentType & "_" & "EDIT", "แก้ไข") Then
'      Call EnableForm(Me, True)
'      Exit Sub
'   End If
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(2))
   OKClick = False
'
   If TabStrip1.SelectedItem.Index = 1 Then
         Set frmAddEditGlDetail.ParentForm = Me
         frmAddEditGlDetail.ID = ID
         frmAddEditGlDetail.HeaderText = "เพิ่มรายการใบสำคัญจ่าย (PV)"
         frmAddEditGlDetail.ListType = TabStrip1.SelectedItem.Index
         frmAddEditGlDetail.ShowMode = SHOW_EDIT
         Set frmAddEditGlDetail.TempCollection = m_BillingPayment.GlDetails
         Load frmAddEditGlDetail
         frmAddEditGlDetail.Show 1

         OKClick = frmAddEditGlDetail.OKClick
         
         Unload frmAddEditGlDetail
         Set frmAddEditGlDetail = Nothing
         If OKClick Then
            m_HasModify = True
            
            GridEX1.ItemCount = CountItem(m_BillingPayment.GlDetails)
            Call GridEX1.Rebind
         End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
         Set frmAddEditGlDetail.ParentForm = Me
         frmAddEditGlDetail.ID = ID
         frmAddEditGlDetail.HeaderText = "แก้ไขรายการใบสำคัญจ่าย (JV)"
         frmAddEditGlDetail.ListType = TabStrip1.SelectedItem.Index
         frmAddEditGlDetail.ShowMode = SHOW_EDIT
         Set frmAddEditGlDetail.TempCollection = m_BillingPayment.GlDetails2
         Load frmAddEditGlDetail
         frmAddEditGlDetail.Show 1
         
         OKClick = frmAddEditGlDetail.OKClick
         
         Unload frmAddEditGlDetail
         Set frmAddEditGlDetail = Nothing
         If OKClick Then
            m_HasModify = True
            
            GridEX1.ItemCount = CountItem(m_BillingPayment.GlDetails2)
            Call GridEX1.Rebind
         End If
   End If
End Sub

Private Sub cmdOK_Click()
Dim oMenu As cPopupMenu
Dim lMenuChosen  As Long

   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("บันทึก", "-", "บันทึกและออกจากหน้าจอ")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
   If lMenuChosen = 1 Then
      If Not SaveData Then
         Exit Sub
      End If
      
      ShowMode = SHOW_EDIT
      ID = m_BillingPayment.BILLING_PAYMENT_ID
      ID2 = m_BillingPayment.BILLING_PAYMENT_ID_REF
      Set m_BillingPayment = Nothing
      Set m_BillingPayment = New CBillingPayment
      m_BillingPayment.QueryFlag = 1
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
Dim DocumentType2 As Long

   Programowner = glbParameterObj.Programowner
   ReportMode = 1
   
   If m_HasModify Or (m_BillingPayment.BILLING_PAYMENT_ID <= 0) Then
      glbErrorLog.LocalErrorMsg = MapText("กรุณาทำการบันทึกข้อมูลให้เรียบร้อยก่อน")
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   ReportFlag = False
   Call LoadPictureFromFile(glbParameterObj.ReceiptPicture1, Picture2)
   
   Set oMenu = New cPopupMenu
    lMenuChosen = oMenu.Popup("พิมพ์ใบสำคัญจ่าย (PV)", "พิมพ์ใบสำคัญจ่าย (PV) มีพื้นหลัง", "-", "ตั้งค่าหน้ากระดาษ", "-", "พิมพ์ใบสำคัญโอนบัญชี (JV)", "พิมพ์ใบสำคัญโอนบัญชี (JV) มีพื้นหลัง", "-", "ตั้งค่าหน้ากระดาษ")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   Set oMenu = Nothing
    If lMenuChosen = 1 Then 'พิมพ์ใบสำคัญจ่าย (PV)
      ReportKey = "CReportVoucherPay2"
      Set Report = New CReportVoucherPay2
      ReportFlag = True
      DocumentType2 = 111
   ElseIf lMenuChosen = 2 Then 'พิมพ์ใบสำคัญจ่าย (PV) มีพื้นหลัง
      ReportKey = "CReportVoucherPay2"
      Set Report = New CReportVoucherPay2
      Picture2.Picture = LoadPicture(glbParameterObj.PaymentVoucher)
      ReportFlag = True
      DocumentType2 = 111
   ElseIf lMenuChosen = 4 Then
      ReportKey = "CReportVoucherPay2"
      DocumentType2 = 111
   ElseIf lMenuChosen = 6 Then
      ReportKey = "CReportVoucherPay2_JV"
      Set Report = New CReportVoucherPay2
      ReportFlag = True
      DocumentType2 = 112
   ElseIf lMenuChosen = 7 Then
      ReportKey = "CReportVoucherPay2_JV"
      Set Report = New CReportVoucherPay2
      Picture2.Picture = LoadPicture(glbParameterObj.AccountTransfer)
      ReportFlag = True
      DocumentType2 = 112
   ElseIf lMenuChosen = 9 Then
      ReportKey = "CReportVoucherPay2_JV"
      DocumentType2 = 112
   End If

   If Not Report Is Nothing Then
      Call Report.AddParam(lMenuChosen, "REPORT_TYPE")
      Call Report.AddParam(m_BillingPayment, "m_BillingPayment")
      Call Report.AddParam(ReportKey, "REPORT_KEY")
      Call Report.AddParam(DocumentType2, "DOCUMENT_TYPE")
      Call Report.AddParam("", "ACCEPT_NAME")
      Call Report.AddParam(Picture2.Picture, "BACK_GROUND")
   End If
   
   Call EnableForm(Me, False)
   
   If ReportFlag Then
    frmReport.ClassName = ReportKey
      Set frmReport.ReportObject = Report

      frmReport.HeaderText = pnlHeader.Caption
      Load frmReport
      frmReport.Show 1

      Unload frmReport
      Set frmReport = Nothing
      Set Report = Nothing

   Else
      If lMenuChosen = 4 Or lMenuChosen = 9 Then
         ReportMode = 2
      End If
         
         Set Rc = New CReportConfig
         Rc.REPORT_KEY = ReportKey
         Call Rc.QueryData(m_Rs, iCount)
         
         If Not m_Rs.EOF Then
            Call Rc.PopulateFromRS(1, m_Rs)
            frmReportConfig.ShowMode = SHOW_EDIT
            frmReportConfig.ID = Rc.REPORT_CONFIG_ID
         Else
            frmReportConfig.ShowMode = SHOW_ADD
         End If
         
         frmReportConfig.ReportMode = ReportMode
         frmReportConfig.ReportKey = ReportKey
         frmReportConfig.HeaderText = HeaderText
         Load frmReportConfig
         frmReportConfig.Show 1
         
         Unload frmReportConfig
         Set frmReportConfig = Nothing
         
         Set Rc = Nothing
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      Call LoadConfigDoc(Nothing, m_Cd)
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_BillingPayment.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
         
      ElseIf ShowMode = SHOW_ADD Then
         uctlDocumentDate.ShowDate = Now
         uctlDocumentDate2.ShowDate = Now
'         uctlPaymentDue.ShowDate = Now
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
Private Sub Form_Resize()
On Error Resume Next

   SSFrame1.Top = 0
   SSFrame1.Width = ScaleWidth
   SSFrame1.HEIGHT = ScaleHeight
   
   pnlHeader.Width = ScaleWidth
   
   GridEX1.Width = ScaleWidth - 300
   GridEX1.HEIGHT = ScaleHeight - GridEX1.Top - 640
   
   TabStrip1.Width = GridEX1.Width
   SSFrame4.Top = GridEX1.Top
   SSFrame4.Width = GridEX1.Width
   SSFrame4.HEIGHT = GridEX1.HEIGHT
   
   cmdAdd.Top = ScaleHeight - 580
   cmdEdit.Top = ScaleHeight - 580
   cmdDelete.Top = ScaleHeight - 580
   
   cmdOK.Top = ScaleHeight - 580
   cmdExit.Top = ScaleHeight - 580
  cmdPrint.Top = ScaleHeight - 580
  
  
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
    cmdPrint.Left = cmdOK.Left - cmdPrint.Width - 50
End Sub
Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_BillingPayment = Nothing
   Set m_Cd = Nothing
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
   Col.Width = 10
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 10
   Col.Caption = "Real ID"
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 6000
   Col.Caption = MapText("ชื่อบัญชี")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 2100
   Col.Caption = MapText("รหัสบัญชี")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 2100
   Col.Caption = MapText("รายละเอียด")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 2100
   Col.Caption = MapText("Dr.")
   
   Set Col = GridEX1.Columns.add '6
   Col.Width = 2100
   Col.Caption = MapText("Cr.")
   
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
   Col.Width = 10
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 10
   Col.Caption = "Real ID"
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 2100
   Col.Caption = MapText("รหัสบัญชี")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 6000
   Col.Caption = MapText("ชื่อบัญชี")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 2100
   Col.Caption = MapText("รายละเอียด")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 2100
   Col.Caption = MapText("เดบิต")
   
   Set Col = GridEX1.Columns.add '6
   Col.Width = 2100
   Col.Caption = MapText("เครดิต")
   
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame4.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblDocumentNo, MapText("เลขที่ใบ PV"))
   Call InitNormalLabel(lblDocumentNo2, MapText("เลขที่ใบ JV"))
   Call InitNormalLabel(lblDocumentDate, MapText("วันที่ใบ PV"))
   Call InitNormalLabel(lblDocumentDate2, MapText("วันที่ใบ JV"))
   Call InitNormalLabel(lblPaymentTo, MapText("เพื่อจ่ายให้"))
   Call InitNormalLabel(lblDocAssembly, MapText("เลขที่เอกสารประกอบรายการ"))
   Call InitNormalLabel(lblPaymentCost, MapText("เป็นการชำระค่า"))
   Call InitNormalLabel(lblPaymentDue, MapText("กำหนดจ่าย"))
   Call InitNormalLabel(lblPaymentPart, MapText("เพื่อเป็นค่าใช้จ่ายของส่วน"))
   Call InitNormalLabel(lblPaymentDept, MapText("แผนก"))
   Call InitNormalLabel(lblPaymentDept2, MapText("ฝ่าย"))
   Call InitNormalLabel(lblPaymentAmount, MapText("จำนวนเงิน"))
   Call InitNormalLabel(lblBath, MapText("บาท"))
   Call InitNormalLabel(lblPaymentBy, MapText("ชำระโดย"))
   Call InitOptionEx(sspCash, "เงินสด")
   Call InitOptionEx(sspCheque, "เช็ค")
   Call InitNormalLabel(lblNote, MapText("คำอธิบายรายการ"))

   Call txtDocumentNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtPaymentTo.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtPaymentCost.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtPaymentAmount.SetTextLenType(TEXT_INTEGER, glbSetting.MONEY_TYPE)
   
   txtPaymentTo.Text = "คุณ นันท์นภัส ช่างทอง หรือ คุณ สุพิน งามมาก"
   sspCash.Value = True
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   SSFrame4.Visible = False
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAuto.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAuto2.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPrint.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdPrint, MapText("พิมพ์"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdAuto, MapText("A"))
   Call InitMainButton(cmdAuto2, MapText("A"))
   
   Call InitGrid1
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.NAME = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear

   TabStrip1.Tabs.add().Caption = MapText("รายละเอียดใบสำคัญจ่าย (PV)")
   TabStrip1.Tabs.add().Caption = MapText("รายละเอียดใบสำคัญโอนบัญชี (JV)")
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
   Set m_BillingPayment = New CBillingPayment
   Set m_Cd = New Collection
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
Dim Gl As CGLDetail

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Index = 1 Then
         If m_BillingPayment.GlDetails Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      If m_BillingPayment.GlDetails.Count <= 0 Then
         Exit Sub
      End If
      Set Gl = GetItem(m_BillingPayment.GlDetails, RowIndex, RealIndex)
      If Gl Is Nothing Then
         Exit Sub
      End If

      Values(1) = Gl.GetFieldValue("GL_DETAIL_ID")
      Values(2) = RealIndex
      Values(3) = Gl.GetFieldValue("GL_NAME")
      Values(4) = Gl.GetFieldValue("GL_NO")
      Values(5) = Gl.GetFieldValue("GL_DESC")
      If Gl.GetFieldValue("GL_TYPE") = 1 Then
         Values(6) = FormatNumber(Gl.GetFieldValue("GL_AMOUNT"))
         Values(7) = ""
      Else
         Values(6) = ""
         Values(7) = FormatNumber(Gl.GetFieldValue("GL_AMOUNT"))
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
         If m_BillingPayment.GlDetails2 Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      If m_BillingPayment.GlDetails2.Count <= 0 Then
         Exit Sub
      End If
      Set Gl = GetItem(m_BillingPayment.GlDetails2, RowIndex, RealIndex)
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
   
   End If
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub


Private Sub sspCash_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub sspCash_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub sspCheque_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub sspCheque_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub TabStrip1_Click()
   If TabStrip1.SelectedItem.Index = 1 Then
      If Len(txtDocumentNo.Text) > 0 Then
         cmdAdd.Enabled = True
         cmdEdit.Enabled = True
         cmdDelete.Enabled = True
         

      Else
         cmdAdd.Enabled = False
         cmdEdit.Enabled = False
         cmdDelete.Enabled = False
      End If
      
         Call InitGrid1
         GridEX1.ItemCount = CountItem(m_BillingPayment.GlDetails)
         GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      If Len(txtDocumentNo2.Text) > 0 Then
         cmdAdd.Enabled = True
         cmdEdit.Enabled = True
         cmdDelete.Enabled = True
      Else
         cmdAdd.Enabled = False
         cmdEdit.Enabled = False
         cmdDelete.Enabled = False
      End If
      Call InitGrid2
      GridEX1.ItemCount = CountItem(m_BillingPayment.GlDetails2)
      GridEX1.Rebind
   End If
End Sub

Private Sub txtDocAssembly_Change()
   m_HasModify = True
End Sub

Private Sub txtDocumentNo_Change()
   m_HasModify = True
End Sub

Private Sub txtDocumentNo2_Change()
m_HasModify = True

End Sub

Private Sub txtNote_Change()
   m_HasModify = True
End Sub

Private Sub txtPaymentAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtPaymentCost_Change()
   m_HasModify = True
End Sub

Private Sub txtPaymentDept_Change()
   m_HasModify = True
End Sub

Private Sub txtPaymentDept2_Change()
   m_HasModify = True
End Sub

Private Sub txtPaymentPart_Change()
   m_HasModify = True
End Sub

Private Sub txtPaymentTo_Change()
   m_HasModify = True
End Sub

Private Sub uctlDocumentDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlDocumentDate2_HasChange()
m_HasModify = True
End Sub

Private Sub uctlPaymentDue_HasChange()
   m_HasModify = True
End Sub
Public Sub RefreshGridAccountList(AccountListType As Long)
   If AccountListType = 1 Then
      GridEX1.ItemCount = CountItem(m_BillingPayment.GlDetails)
   ElseIf AccountListType = 2 Then
      GridEX1.ItemCount = CountItem(m_BillingPayment.GlDetails2)
   End If
   GridEX1.Rebind
End Sub
Public Function calSumAmount(mData As CBillingPayment) As Double
Dim TotalDr As Double
Dim TotalCr As Double
Dim TotalPv As Double
Dim TotalJv As Double
Dim Gl As CGLDetail
Dim CountPV As Long
Dim CountJV As Long
calSumAmount = -1
CountPV = CountItem(mData.GlDetails)
CountJV = CountItem(mData.GlDetails2)
If CountPV > 0 Then
   For Each Gl In mData.GlDetails
       If Gl.Flag <> "D" Then
        If Gl.GetFieldValue("GL_TYPE") = 1 Then
         TotalDr = TotalDr + Gl.GetFieldValue("GL_AMOUNT")
        ElseIf Gl.GetFieldValue("GL_TYPE") = 2 Then
         TotalCr = TotalCr + Gl.GetFieldValue("GL_AMOUNT")
        End If
      End If
   Next Gl
   If TotalDr = TotalCr Then
      TotalPv = TotalDr
      calSumAmount = TotalPv
   Else
      calSumAmount = -1
   End If
End If

TotalDr = 0
TotalCr = 0
If CountJV > 0 Then
   For Each Gl In mData.GlDetails2
      If Gl.Flag <> "D" Then
         If Gl.GetFieldValue("GL_TYPE") = 1 Then
          TotalDr = TotalDr + Gl.GetFieldValue("GL_AMOUNT")
         ElseIf Gl.GetFieldValue("GL_TYPE") = 2 Then
          TotalCr = TotalCr + Gl.GetFieldValue("GL_AMOUNT")
         End If
     End If
   Next Gl
   If TotalDr = TotalCr And CountJV > 0 Then
      TotalJv = TotalDr
      calSumAmount = TotalJv
   Else
      calSumAmount = -1
   End If
End If

If CountPV > 0 And CountJV = 0 Then 'กรณีที่มี PV อย่างเดียว
   If calSumAmount > -1 Then
      calSumAmount = TotalPv
   Else
      calSumAmount = -1
   End If
ElseIf CountPV > 0 And CountJV > 0 Then 'กรณีที่มี PV  และ JV อย่างเดียว
   If TotalPv = TotalJv And calSumAmount > -1 Then
      calSumAmount = TotalPv
   Else
      calSumAmount = -1
   End If
ElseIf CountPV = 0 And CountJV > 0 Then 'กรณีที่มี JV อย่างเดียว
   calSumAmount = -1
ElseIf CountPV = 0 And CountJV = 0 Then 'กรณีที่เปิดหัวอย่างเดียวแต่ยังไม่ลงรายละเอียด
   calSumAmount = 0
End If

End Function
