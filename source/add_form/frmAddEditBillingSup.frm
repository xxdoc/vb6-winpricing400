VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAddEditBillingSup 
   ClientHeight    =   9570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11775
   ForeColor       =   &H00000000&
   Icon            =   "frmAddEditBillingSup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   11775
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   9840
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   17357
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   120
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   4960
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
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000009&
         Height          =   1275
         Left            =   0
         ScaleHeight     =   1215
         ScaleWidth      =   1575
         TabIndex        =   53
         Top             =   0
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.ComboBox cboDepartMent 
         Height          =   315
         Left            =   9510
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2160
         Width           =   1905
      End
      Begin prjFarmManagement.uctlTime uctlEntryTime 
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   1740
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
      End
      Begin prjFarmManagement.uctlTextLookup uctlSupplierLookup 
         Height          =   435
         Left            =   6000
         TabIndex        =   5
         Top             =   1260
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlDocumentDate 
         Height          =   405
         Left            =   6000
         TabIndex        =   2
         Top             =   840
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtDoNo 
         Height          =   435
         Left            =   1560
         TabIndex        =   4
         Top             =   1290
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtDocumentNo 
         Height          =   435
         Left            =   1560
         TabIndex        =   0
         Top             =   840
         Width           =   2295
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTruckNo 
         Height          =   435
         Left            =   1560
         TabIndex        =   9
         Top             =   2160
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
         Height          =   3135
         Left            =   120
         TabIndex        =   25
         Top             =   5520
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   5530
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
         Column(1)       =   "frmAddEditBillingSup.frx":27A2
         Column(2)       =   "frmAddEditBillingSup.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditBillingSup.frx":290E
         FormatStyle(2)  =   "frmAddEditBillingSup.frx":2A6A
         FormatStyle(3)  =   "frmAddEditBillingSup.frx":2B1A
         FormatStyle(4)  =   "frmAddEditBillingSup.frx":2BCE
         FormatStyle(5)  =   "frmAddEditBillingSup.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditBillingSup.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtSender 
         Height          =   435
         Left            =   1560
         TabIndex        =   13
         Top             =   2610
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtReceiver 
         Height          =   435
         Left            =   6000
         TabIndex        =   14
         Top             =   2640
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtDeliveryFee 
         Height          =   435
         Left            =   1560
         TabIndex        =   19
         Top             =   3540
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtMaterialPrice 
         Height          =   435
         Left            =   6000
         TabIndex        =   20
         Top             =   3600
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTotal 
         Height          =   435
         Left            =   9480
         TabIndex        =   21
         Top             =   3540
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtQueNo 
         Height          =   435
         Left            =   9510
         TabIndex        =   15
         Top             =   2610
         Width           =   1365
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtDesc 
         Height          =   435
         Left            =   6000
         TabIndex        =   8
         Top             =   1710
         Width           =   5385
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTime uctlExitTime 
         Height          =   375
         Left            =   3270
         TabIndex        =   7
         Top             =   1770
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
      End
      Begin prjFarmManagement.uctlTextBox txtCredit 
         Height          =   435
         Left            =   10830
         TabIndex        =   3
         Top             =   840
         Width           =   525
         _ExtentX        =   2408
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPrNo 
         Height          =   435
         Left            =   1560
         TabIndex        =   16
         Top             =   3080
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlDueDate 
         Height          =   405
         Left            =   6000
         TabIndex        =   17
         Top             =   3120
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtDueAmount 
         Height          =   435
         Left            =   9840
         TabIndex        =   18
         Top             =   3120
         Width           =   525
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin Threed.SSFrame SSFrame4 
         Height          =   3135
         Left            =   120
         TabIndex        =   59
         Top             =   5520
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   5530
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin VB.ComboBox cboCondition 
            Height          =   315
            Left            =   2760
            Style           =   2  'Dropdown List
            TabIndex        =   61
            Top             =   240
            Width           =   4035
         End
         Begin VB.ComboBox cboPaidType 
            Height          =   315
            Left            =   2760
            Style           =   2  'Dropdown List
            TabIndex        =   60
            Top             =   720
            Width           =   4005
         End
         Begin VB.Label lblPaidType 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   840
            TabIndex        =   62
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label lblCondition 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   360
            TabIndex        =   63
            Top             =   240
            Width           =   2295
         End
      End
      Begin prjFarmManagement.uctlTextBox txtTotalAmount 
         Height          =   435
         Left            =   9480
         TabIndex        =   24
         Top             =   3960
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlSupplierTrueLookup 
         Height          =   435
         Left            =   6000
         TabIndex        =   67
         Top             =   4440
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin VB.Label lblSupplierTrueNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4470
         TabIndex        =   68
         Top             =   4500
         Width           =   1455
      End
      Begin VB.Label lblVolume 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8280
         TabIndex        =   66
         Top             =   4080
         Width           =   1125
      End
      Begin Threed.SSCheck chkDeliveryFee 
         Height          =   375
         Left            =   1560
         TabIndex        =   65
         Top             =   3960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSCheck2"
         Value           =   1
      End
      Begin Threed.SSCommand cmdOther 
         Height          =   525
         Left            =   5100
         TabIndex        =   29
         Top             =   8880
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditBillingSup.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCheck chkGenCommitFlag 
         Height          =   375
         Left            =   1560
         TabIndex        =   22
         Top             =   4320
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSCheck2"
      End
      Begin Threed.SSCheck chkClose 
         Height          =   375
         Left            =   6000
         TabIndex        =   23
         Top             =   3960
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSCheck2"
      End
      Begin VB.Label Label5 
         Height          =   315
         Left            =   10440
         TabIndex        =   58
         Top             =   3120
         Width           =   585
      End
      Begin VB.Label lblDueDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4440
         TabIndex        =   57
         Top             =   3240
         Width           =   1515
      End
      Begin VB.Label lblPrNo 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   0
         TabIndex        =   56
         Top             =   3180
         Width           =   1485
      End
      Begin VB.Label Label6 
         Height          =   315
         Left            =   11400
         TabIndex        =   55
         Top             =   870
         Width           =   405
      End
      Begin VB.Label lblCredit 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   9960
         TabIndex        =   54
         Top             =   870
         Width           =   855
      End
      Begin VB.Label lblDepartment 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   345
         Left            =   8400
         TabIndex        =   52
         Top             =   2220
         Width           =   1035
      End
      Begin Threed.SSCheck chkException 
         Height          =   435
         Left            =   7620
         TabIndex        =   11
         Top             =   2190
         Visible         =   0   'False
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2730
         TabIndex        =   51
         Top             =   1800
         Width           =   435
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   525
         Left            =   6720
         TabIndex        =   30
         Top             =   8880
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditBillingSup.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAuto 
         Height          =   405
         Left            =   3870
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   840
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditBillingSup.frx":356A
         ButtonStyle     =   3
      End
      Begin VB.Label lblQueNo 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   8970
         TabIndex        =   50
         Top             =   2670
         Width           =   465
      End
      Begin Threed.SSCheck chkCommit 
         Height          =   435
         Left            =   6000
         TabIndex        =   10
         Top             =   2190
         Visible         =   0   'False
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4470
         TabIndex        =   49
         Top             =   1740
         Width           =   1455
      End
      Begin VB.Label lblSupplierNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4470
         TabIndex        =   48
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label4 
         Height          =   315
         Left            =   10950
         TabIndex        =   47
         Top             =   3660
         Width           =   585
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8280
         TabIndex        =   46
         Top             =   3660
         Width           =   1125
      End
      Begin VB.Label lblMaterialPrice 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4170
         TabIndex        =   45
         Top             =   3690
         Width           =   1695
      End
      Begin VB.Label Label2 
         Height          =   315
         Left            =   7440
         TabIndex        =   44
         Top             =   3660
         Width           =   765
      End
      Begin VB.Label Label1 
         Height          =   315
         Left            =   3900
         TabIndex        =   43
         Top             =   3660
         Width           =   765
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4710
         TabIndex        =   42
         Top             =   870
         Width           =   1155
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8400
         TabIndex        =   31
         Top             =   8880
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditBillingSup.frx":3884
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10080
         TabIndex        =   32
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
         TabIndex        =   27
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
         TabIndex        =   26
         Top             =   8880
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditBillingSup.frx":3B9E
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3480
         TabIndex        =   28
         Top             =   8880
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditBillingSup.frx":3EB8
         ButtonStyle     =   3
      End
      Begin VB.Label lblDeliveryFee 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -210
         TabIndex        =   40
         Top             =   3660
         Width           =   1695
      End
      Begin VB.Label lblTruckNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -60
         TabIndex        =   39
         Top             =   2250
         Width           =   1575
      End
      Begin VB.Label lblDeliveryNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -60
         TabIndex        =   38
         Top             =   1830
         Width           =   1575
      End
      Begin VB.Label lblDocumentNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -150
         TabIndex        =   37
         Top             =   900
         Width           =   1665
      End
      Begin VB.Label lblReceiver 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4410
         TabIndex        =   36
         Top             =   2700
         Width           =   1485
      End
      Begin VB.Label lblSender 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   30
         TabIndex        =   35
         Top             =   2670
         Width           =   1485
      End
      Begin VB.Label lblDoNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -60
         TabIndex        =   34
         Top             =   1380
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmAddEditBillingSup"
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
Private m_Weight As CWeight
Private m_Suppliers As Collection
Private m_SuppliersTrue As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public SupplierIDTrue As Long
Public AutoGenPo As Boolean

Public ID As Long
Public DocumentType As Long

Private FileName As String
Private m_SumUnit As Double
Private m_SumTotalPrice As Double
Private m_PartTxtypes As Collection
Private m_AuthenPO_Verify As Collection
Private m_AuthenPO_Approve As Collection
Private m_Cd As Collection
Private TempWeight As Collection
Private DocAdd As Long
Private CW As CWeight
Private SupCode As String
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
   
   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
      
      m_BillingDoc.BILLING_DOC_ID = ID
      m_BillingDoc.COMMIT_FLAG = ""
      If Not glbDaily.QueryBillingDoc(m_BillingDoc, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
        Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_BillingDoc.PopulateFromRS(1, m_Rs)
      
      If m_BillingDoc.AUTO_GEN_FLAG = "Y" Then
         chkGenCommitFlag.Visible = True
      End If
      
      uctlDocumentDate.ShowDate = m_BillingDoc.DOCUMENT_DATE
      txtDoNo.Text = m_BillingDoc.DO_NO
      txtTruckNo.Text = m_BillingDoc.TRUCK_NO
      txtDocumentNo.Text = m_BillingDoc.DOCUMENT_NO
      txtDeliveryFee.Text = Format(m_BillingDoc.DELIVERY_FEE, "0.00")
      txtSender.Text = m_BillingDoc.SENDER_NAME
      txtReceiver.Text = m_BillingDoc.RECEIVE_NAME
      uctlSupplierLookup.MyCombo.ListIndex = IDToListIndex(uctlSupplierLookup.MyCombo, m_BillingDoc.SUPPLIER_ID)
      uctlSupplierTrueLookup.MyCombo.ListIndex = IDToListIndex(uctlSupplierTrueLookup.MyCombo, m_BillingDoc.SUPPLIER_ID_TRUE)
      SupCode = uctlSupplierLookup.MyTextBox.Text
      cboDepartMent.ListIndex = IDToListIndex(cboDepartMent, m_BillingDoc.DEPARTMENT_ID)
'      chkCommit.Value = FlagToCheck(m_BillingDoc.COMMIT_FLAG)
      txtQueNo.Text = m_BillingDoc.QUE_NO
      txtDesc.Text = m_BillingDoc.NOTE
      uctlEntryTime.HR = HOUR(m_BillingDoc.ENTRY_DATE)
      uctlEntryTime.MI = Minute(m_BillingDoc.ENTRY_DATE)
      uctlExitTime.HR = HOUR(m_BillingDoc.EXIT_DATE)
      uctlExitTime.MI = Minute(m_BillingDoc.EXIT_DATE)
      chkException.Value = FlagToCheck(m_BillingDoc.EXCEPTION_FLAG)
      txtCredit.Text = m_BillingDoc.Credit
      txtPrNo.Text = m_BillingDoc.PR_NO
      
      uctlDueDate.ShowDate = m_BillingDoc.DUE_DATE
      txtDueAmount.Text = m_BillingDoc.DUE_AMOUNT
      
      cmdAdd.Enabled = (m_BillingDoc.OLD_COMMIT_FLAG = "N")
      cmdDelete.Enabled = (m_BillingDoc.OLD_COMMIT_FLAG = "N")
'      chkCommit.Enabled = (m_BillingDoc.OLD_COMMIT_FLAG = "N")
      cboCondition.ListIndex = IDToListIndex(cboCondition, m_BillingDoc.CONDITION)
      cboPaidType.ListIndex = IDToListIndex(cboPaidType, m_BillingDoc.PAID_TYPE)
'      txtDeliveryFee.Enabled = (m_BillingDoc.OLD_COMMIT_FLAG = "N")
      chkClose.Value = FlagToCheck(m_BillingDoc.CLOSE_FLAG)
      chkGenCommitFlag.Value = FlagToCheck(m_BillingDoc.GEN_COMMIT_FLAG)
      
      If (m_BillingDoc.TAX_FLAG = "") And (m_BillingDoc.DELIVERY_FEE > 0) Then
         chkDeliveryFee.Value = 1
      Else
         chkDeliveryFee.Value = FlagToCheck(m_BillingDoc.TAX_FLAG)
      End If
      
      
      
      
      If chkClose.Value = ssCBChecked Then
         chkClose.Enabled = False
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
               
   If ShowMode = SHOW_EDIT Then
      If Not VerifyAccessRight("LEDGER_STOCKBUY" & "_" & DocumentType & "_" & "EDIT", "���") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   End If
   
   If Not VerifyTextControl(lblDocumentNo, txtDocumentNo, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblDocumentDate, uctlDocumentDate, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblDueDate, uctlDueDate, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblSupplierNo, uctlSupplierLookup.MyCombo, True) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblDueDate, txtDueAmount, True) Then
      Exit Function
   End If
   
   
   If Not CheckUniqueNs(DO_PLAN_UNIQUE, txtDocumentNo.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("�բ�����") & " " & txtDocumentNo.Text & " " & MapText("������к�����")
      glbErrorLog.ShowUserError
      txtDocumentNo.Text = ""
      Call LoadConfigDoc(Nothing, m_Cd)
      Call cmdAuto_Click
      Exit Function
   End If
   
   If ShowMode = SHOW_EDIT Then
      Dim strTemp As String
      strTemp = getReceiptByBillingDocIDRef(ID)
      If Len(strTemp) > 0 Then
         glbErrorLog.LocalErrorMsg = MapText("�͡��� '" & txtDocumentNo.Text & "'" & vbNewLine & "  ���͡������Ѻ�Թ " & vbNewLine & " �Ţ��� '" & strTemp & "' ����" & vbNewLine & " �������ö�����")
         glbErrorLog.ShowUserError
         
         SaveData = False
         Exit Function
      End If
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   If AutoGenPo And m_BillingDoc.SupItems.Count <= 0 Then  '�óշ���Դ��Ѻ�ͧ������� PO ���� �ѹ Save ��͹���� ���� ITEM �١ �ѧ��� ����͹��� ���� ITEM �١��͹
      glbErrorLog.LocalErrorMsg = MapText("�óշ���Դ��Ѻ�ͧ������� PO �е�ͧ������¡���Ѻ�ͧ��úѹ�֡����")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   If DocumentType = 1000 Or DocumentType = 1001 Or DocumentType = 1002 Or DocumentType = 1003 Then
      If m_BillingDoc.PO_APPROVED_FLAG = "Y" Then
         glbErrorLog.LocalErrorMsg = MapText("�������ö�ѹ�֡�����ͧ�ҡ�ա��͹��ѵ�")
         glbErrorLog.ShowUserError
         Exit Function
      End If
   End If
      
   
   m_BillingDoc.AddEditMode = ShowMode
   m_BillingDoc.BILLING_DOC_ID = ID
    m_BillingDoc.DOCUMENT_DATE = uctlDocumentDate.ShowDate
   m_BillingDoc.DO_NO = txtDoNo.Text
   m_BillingDoc.TRUCK_NO = txtTruckNo.Text
   m_BillingDoc.DOCUMENT_NO = txtDocumentNo.Text
   m_BillingDoc.DELIVERY_FEE = Val(txtDeliveryFee.Text)
   m_BillingDoc.SENDER_NAME = txtSender.Text
   m_BillingDoc.RECEIVE_NAME = txtReceiver.Text
   m_BillingDoc.SUPPLIER_ID = uctlSupplierLookup.MyCombo.ItemData(Minus2Zero(uctlSupplierLookup.MyCombo.ListIndex))
   m_BillingDoc.SUPPLIER_ID_TRUE = uctlSupplierTrueLookup.MyCombo.ItemData(Minus2Zero(uctlSupplierTrueLookup.MyCombo.ListIndex))
   m_BillingDoc.DOCUMENT_TYPE = DocumentType
   m_BillingDoc.COMMIT_FLAG = Check2Flag(chkCommit.Value)
   m_BillingDoc.QUE_NO = txtQueNo.Text
   m_BillingDoc.NOTE = txtDesc.Text
   m_BillingDoc.ENTRY_DATE = uctlDocumentDate.ShowDate
   m_BillingDoc.ENTRY_DATE = DateAdd("h", uctlEntryTime.HR, m_BillingDoc.ENTRY_DATE)
   m_BillingDoc.ENTRY_DATE = DateAdd("n", uctlEntryTime.MI, m_BillingDoc.ENTRY_DATE)
   m_BillingDoc.EXIT_DATE = uctlDocumentDate.ShowDate
   m_BillingDoc.EXIT_DATE = DateAdd("h", uctlExitTime.HR, m_BillingDoc.EXIT_DATE)
   m_BillingDoc.EXIT_DATE = DateAdd("n", uctlExitTime.MI, m_BillingDoc.EXIT_DATE)
   m_BillingDoc.EXCEPTION_FLAG = Check2Flag(chkException.Value)
   m_BillingDoc.DEPARTMENT_ID = cboDepartMent.ItemData(Minus2Zero(cboDepartMent.ListIndex))
   m_BillingDoc.Credit = Val(txtCredit.Text)
   m_BillingDoc.PR_NO = txtPrNo.Text
   m_BillingDoc.DUE_AMOUNT = Val(txtDueAmount.Text)
    m_BillingDoc.CONDITION = cboCondition.ItemData(Minus2Zero(cboCondition.ListIndex))
    m_BillingDoc.PAID_TYPE = cboPaidType.ItemData(Minus2Zero(cboPaidType.ListIndex))
    m_BillingDoc.CLOSE_FLAG = Check2Flag(chkClose.Value)
   m_BillingDoc.GEN_COMMIT_FLAG = Check2Flag(chkGenCommitFlag.Value)
        
   m_BillingDoc.TOTAL_PRICE = Val(txtTotal.Text)
   
   m_BillingDoc.TOTAL_AMOUNT = Val(txtTotalAmount.Text)
   m_BillingDoc.TAX_FLAG = Check2Flag(chkDeliveryFee.Value)
      
    StrStockAmount = ""
    
       If ShowMode = SHOW_ADD Then
         If DocumentType = 1000 Or DocumentType = 1001 Or DocumentType = 1002 Or DocumentType = 1003 Then
            Set m_PartTxtypes = New Collection
            Call GetFirstLastDate(uctlDocumentDate.ShowDate, firstDate, lastDate)
               
            Set MonthlyAccums = New Collection
            Set InventoryBals1 = New Collection
            
            YYYYMM = Format(Year(DateAdd("D", -1, firstDate)), "0000") & "-" & Format(Month(DateAdd("D", -1, firstDate)), "00")
         End If
      End If
    
   For Each Sp In m_BillingDoc.SupItems
      If DocumentType = 1000 Or DocumentType = 1001 Or DocumentType = 1002 Or DocumentType = 1003 Then
         If ShowMode = SHOW_ADD Then
            Call LoadMonthlyBalancePartItem(Nothing, MonthlyAccums, YYYYMM, , Sp.PART_ITEM_ID)
            If Not MonthlyAccums Is Nothing Then
               Call glbDaily.CopyMonthlyAccumPartItem(MonthlyAccums, InventoryBals1)
            End If
            Call LoadPartTxTypeAmount(Nothing, m_PartTxtypes, firstDate, lastDate, , , , Sp.PART_ITEM_ID)
            
            Set BalanceLi = GetLotItem(InventoryBals1, Trim(str(Sp.PART_ITEM_ID)))
            Set TempLi1 = GetLotItem(m_PartTxtypes, Sp.PART_ITEM_ID & "-" & "I")
            Set TempLi2 = GetLotItem(m_PartTxtypes, Sp.PART_ITEM_ID & "-" & "E")
            
            If StrStockAmount = "" Then
                 StrStockAmount = Sp.PART_DESC & " (" & FormatNumber(BalanceLi.NEW_AMOUNT + TempLi1.TX_AMOUNT - TempLi2.TX_AMOUNT) & ") " & Sp.UNIT_NAME
              Else
                StrStockAmount = StrStockAmount & vbCrLf & Sp.PART_DESC & " (" & FormatNumber(BalanceLi.NEW_AMOUNT + TempLi1.TX_AMOUNT - TempLi2.TX_AMOUNT) & ") " & Sp.UNIT_NAME
            End If
         End If
      End If
   Next
   Set MonthlyAccums = Nothing
   
   m_BillingDoc.DUE_DATE = uctlDueDate.ShowDate
   
   If ID > 0 Then
      If DocumentType = 1000 Or DocumentType = 1001 Or DocumentType = 1002 Or DocumentType = 1003 Then
         If glbDaily.VerifyReferPo(ID, "(100,101,102,103)", glbErrorLog) Then
            SaveData = False
            Call EnableForm(Me, True)
            Exit Function
         End If
      End If
   End If
   
   Call CalculateIncludePrice
   
   Call PopulateGuiID(m_BillingDoc)
   
   Call EnableForm(Me, False)
   
   If DocumentType = 100 Or DocumentType = 101 Or DocumentType = 102 Or DocumentType = 103 Then
      If DocumentType = 100 Then   '��Ѻ����ѵ�شԺ
         Call glbDaily.SUP2InventoryDoc(m_BillingDoc, Ivd, 1)
      ElseIf DocumentType = 101 Then   '��Ѻ�����ʴ��ػ�ó�
         Call glbDaily.SUP2InventoryDoc(m_BillingDoc, Ivd, 19)
      ElseIf DocumentType = 102 Then   '㺨����͡��ʴ��ػ�ó�
         Call glbDaily.SUP2InventoryDocEx(m_BillingDoc, Ivd, 20)
      ElseIf DocumentType = 103 Then   '��Ѻ��Ңͧ������
         Call glbDaily.SUP2InventoryDoc(m_BillingDoc, Ivd, 23)
      End If
   End If
   Call glbDaily.StartTransaction
   
   If AutoGenPo Then '��͹��������Ѻ�ͧ������� PO ��͹���
      '����Դ��Ѻ�ͧ������� PO ���� Auto ���ҧ PO ������ Set Flag �����Ѻ�ͧ�ͧ���� PO ��� Auto Gen ����Ҵ���
      
      m_BillingDoc.AUTO_GEN_FLAG = "Y"
      If DocumentType = 100 Then
         m_BillingDoc.DOCUMENT_NO = GetDocumentNo(1010)
         m_BillingDoc.DOCUMENT_TYPE = 1000
      ElseIf DocumentType = 101 Then
         m_BillingDoc.DOCUMENT_NO = GetDocumentNo(1011)
         m_BillingDoc.DOCUMENT_TYPE = 1001
      ElseIf DocumentType = 102 Then
         m_BillingDoc.DOCUMENT_NO = GetDocumentNo(1012)
         m_BillingDoc.DOCUMENT_TYPE = 1002
      ElseIf DocumentType = 103 Then
         m_BillingDoc.DOCUMENT_NO = GetDocumentNo(1013)
         m_BillingDoc.DOCUMENT_TYPE = 1003
      End If
      m_BillingDoc.DUE_DATE = uctlDocumentDate.ShowDate
      
      If Not glbDaily.AddEditBillingDoc(m_BillingDoc, IsOK, False, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         SaveData = False
         Call glbDaily.RollbackTransaction
         Call EnableForm(Me, True)
         Exit Function
      End If
      
      For Each Sp In m_BillingDoc.SupItems
         Sp.PO_ID = m_BillingDoc.BILLING_DOC_ID
         Sp.PO_NO = m_BillingDoc.DOCUMENT_NO
      Next Sp
      
      m_BillingDoc.DOCUMENT_NO = txtDocumentNo.Text
      m_BillingDoc.DOCUMENT_TYPE = DocumentType
      m_BillingDoc.DUE_DATE = uctlDueDate.ShowDate
   End If
   
   If DocumentType = 100 Or DocumentType = 101 Or DocumentType = 102 Or DocumentType = 103 Then
      If Not glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         SaveData = False
         Call glbDaily.RollbackTransaction
         Call EnableForm(Me, True)
         Exit Function
      End If
      
      
      m_BillingDoc.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID
   End If
   
   If Not glbDaily.AddEditBillingDoc(m_BillingDoc, IsOK, False, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call glbDaily.RollbackTransaction
      Call EnableForm(Me, True)
      Exit Function
   End If
   If Not CW Is Nothing Then
   m_Weight.AddEditMode = ShowMode
   m_Weight.WEIGHT_ID = CW.WEIGHT_ID
   m_Weight.WEIGHT1 = CW.WEIGHT1
   m_Weight.WEIGHT2 = CW.WEIGHT2
   m_Weight.DOCUMENT_NO = txtDocumentNo.Text
   
       If Not glbDaily.AddEditWeight(m_Weight, IsOK, False, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         SaveData = False
         Call glbDaily.RollbackTransaction
         Call EnableForm(Me, True)
         Exit Function
      End If
   End If
   
   If Not IsOK Then
      Call EnableForm(Me, True)
      glbErrorLog.ShowUserError
      Call glbDaily.RollbackTransaction
      Exit Function
   End If
   
   Call glbDaily.CommitTransaction
   Call EnableForm(Me, True)
   
   Select Case DocumentType
   Case 1000
       If ShowMode = SHOW_ADD Then
      glbErrorLog.LocalErrorMsg = MapText("�ʹ�ѵ�شԺ������ʵ�͡㹢�й�� ���� ") & vbCrLf & StrStockAmount
       glbErrorLog.ShowUserError
     End If
   Case 1001
     If ShowMode = SHOW_ADD Then
      glbErrorLog.LocalErrorMsg = MapText("�ʹ��ʴ��ػ�ó������ʵ�͡㹢�й�� ���� ") & vbCrLf & StrStockAmount
       glbErrorLog.ShowUserError
     End If
   Case 1002
      If ShowMode = SHOW_ADD Then
      glbErrorLog.LocalErrorMsg = MapText("�ʹ �Ѻ��Ҩ����͡��ʴ��ػ�ó������ʵ�͡㹢�й�� ���� ") & vbCrLf & StrStockAmount
       glbErrorLog.ShowUserError
     End If
   Case 1003
        If ShowMode = SHOW_ADD Then
      glbErrorLog.LocalErrorMsg = MapText("�ʹ�ͧ�����价�����ʵ�͡㹢�й�� ���� ") & vbCrLf & StrStockAmount
       glbErrorLog.ShowUserError
     End If
   End Select

   
   
   SaveData = True
End Function

Private Sub cboCondition_Click()
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

Private Sub cboPaidType_Click()
m_HasModify = True
End Sub

Private Sub chkClose_Click(Value As Integer)
   If Not chkClose.Enabled Then
      Exit Sub
   End If
   
   frmVerifyAccRight.AccName = "LEDGER_STOCKBUY_PO-CLOSE"
   frmVerifyAccRight.AccDesc = "����ö�Դ PO ��"
   Load frmVerifyAccRight
   frmVerifyAccRight.Show 1
   
   If frmVerifyAccRight.GrantRight Then
      Unload frmVerifyAccRight
      Set frmVerifyAccRight = Nothing
   Else
      Unload frmVerifyAccRight
      Set frmVerifyAccRight = Nothing
      chkClose.Enabled = False
      chkClose.Value = ssCBUnchecked
      chkClose.Enabled = True
      Exit Sub
   End If
   m_HasModify = True
End Sub

Private Sub chkCommit_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkCommit_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub chkDeliveryFee_Click(Value As Integer)
   m_HasModify = True
  
   If chkDeliveryFee.Value = ssCBChecked Then
       If DocumentType = 1000 Or DocumentType = 100 Then
         Call GetTotalPrice
      Else
         txtDeliveryFee.Text = Format(Val(txtMaterialPrice.Text) * 0.07, "##.00")
      End If
   Else
      txtDeliveryFee.Text = ""
   End If
End Sub

Private Sub chkException_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkException_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub chkGenCommitFlag_Click(Value As Integer)
      m_HasModify = True
End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean
Dim oMenu  As cPopupMenu
Dim lMenuChosen As Long

   If Not cmdAdd.Enabled Then
      Exit Sub
   End If

   OKClick = False
   
   If Not VerifyCombo(lblSupplierNo, uctlSupplierLookup.MyCombo) Then
      Exit Sub
   End If
   
   If TabStrip1.SelectedItem.Index = 1 Then
      If Not (DocumentType = 1000 Or DocumentType = 1001 Or DocumentType = 1002 Or DocumentType = 1003) Then
         Set oMenu = New cPopupMenu
         If AutoGenPo Then
            lMenuChosen = oMenu.AddMenu(glbGuiConfigs.SupAddNoPoMenuItems)
         Else
            lMenuChosen = oMenu.AddMenu(glbGuiConfigs.SupAddMenuItems)
         End If
         Set oMenu = Nothing
         If lMenuChosen = 0 Then
            Exit Sub
         End If
      Else
         lMenuChosen = 1
      End If
      
      If lMenuChosen = 1 Then
         If DocumentType = 100 Or DocumentType = 1000 Then
            frmAddEditBLImportItem.SupplierID = uctlSupplierLookup.MyCombo.ItemData(Minus2Zero(uctlSupplierLookup.MyCombo.ListIndex))
            frmAddEditBLImportItem.DocumentType = DocumentType
            frmAddEditBLImportItem.COMMIT_FLAG = m_BillingDoc.OLD_COMMIT_FLAG
            Set frmAddEditBLImportItem.TempCollection = m_BillingDoc.SupItems
            frmAddEditBLImportItem.ParentShowMode = ShowMode
            frmAddEditBLImportItem.ShowMode = SHOW_ADD
            frmAddEditBLImportItem.HeaderText = MapText("������¡���ѵ�شԺ")
            Load frmAddEditBLImportItem
            frmAddEditBLImportItem.Show 1
            
            OKClick = frmAddEditBLImportItem.OKClick
   
            Unload frmAddEditBLImportItem
            Set frmAddEditBLImportItem = Nothing
         ElseIf (DocumentType = 101) Or DocumentType = 1001 Then
            frmAddEditBLImportItemEx.SupplierID = uctlSupplierLookup.MyCombo.ItemData(Minus2Zero(uctlSupplierLookup.MyCombo.ListIndex))
            frmAddEditBLImportItemEx.COMMIT_FLAG = m_BillingDoc.OLD_COMMIT_FLAG
            Set frmAddEditBLImportItemEx.TempCollection = m_BillingDoc.SupItems
            frmAddEditBLImportItemEx.ParentShowMode = ShowMode
            frmAddEditBLImportItemEx.ShowMode = SHOW_ADD
            frmAddEditBLImportItemEx.HeaderText = MapText("������¡����ʴ��ػ�ó�")
            Load frmAddEditBLImportItemEx
            frmAddEditBLImportItemEx.Show 1
   
            OKClick = frmAddEditBLImportItemEx.OKClick
   
            Unload frmAddEditBLImportItemEx
            Set frmAddEditBLImportItemEx = Nothing
          ElseIf DocumentType = 102 Or DocumentType = 1002 Then
            frmAddEditBLImportItemEx2.SupplierID = uctlSupplierLookup.MyCombo.ItemData(Minus2Zero(uctlSupplierLookup.MyCombo.ListIndex))
            frmAddEditBLImportItemEx2.COMMIT_FLAG = m_BillingDoc.OLD_COMMIT_FLAG
            Set frmAddEditBLImportItemEx2.TempCollection = m_BillingDoc.SupItems
            frmAddEditBLImportItemEx2.ParentShowMode = ShowMode
            frmAddEditBLImportItemEx2.ShowMode = SHOW_ADD
            frmAddEditBLImportItemEx2.HeaderText = MapText("������¡���Ѻ��Ҩ����͡��ʴ��ػ�ó�")
            Load frmAddEditBLImportItemEx2
            frmAddEditBLImportItemEx2.Show 1
   
            OKClick = frmAddEditBLImportItemEx2.OKClick
   
            Unload frmAddEditBLImportItemEx2
            Set frmAddEditBLImportItemEx2 = Nothing
         ElseIf (DocumentType = 103) Or DocumentType = 1003 Then
            frmAddEditBLImportItemEx3.SupplierID = uctlSupplierLookup.MyCombo.ItemData(Minus2Zero(uctlSupplierLookup.MyCombo.ListIndex))
            frmAddEditBLImportItemEx3.COMMIT_FLAG = m_BillingDoc.OLD_COMMIT_FLAG
            Set frmAddEditBLImportItemEx3.TempCollection = m_BillingDoc.SupItems
            frmAddEditBLImportItemEx3.ParentShowMode = ShowMode
            frmAddEditBLImportItemEx3.ShowMode = SHOW_ADD
            frmAddEditBLImportItemEx3.HeaderText = MapText("������¡�âͧ������")
            Load frmAddEditBLImportItemEx3
            frmAddEditBLImportItemEx3.Show 1
   
            OKClick = frmAddEditBLImportItemEx3.OKClick
   
            Unload frmAddEditBLImportItemEx3
            Set frmAddEditBLImportItemEx3 = Nothing
        
         End If
         If OKClick Then
            Call GetTotalPrice
            GridEX1.ItemCount = CountItem(m_BillingDoc.SupItems)
         End If
      ElseIf lMenuChosen = 4 Then
         frmAddPOSupItem.SupplierID = uctlSupplierLookup.MyCombo.ItemData(Minus2Zero(uctlSupplierLookup.MyCombo.ListIndex))
         frmAddPOSupItem.DocumentType = DocumentType
         Set frmAddPOSupItem.TempCollection = m_BillingDoc.SupItems
         frmAddPOSupItem.ShowMode = SHOW_ADD
         Select Case DocumentType
         Case 100
            frmAddPOSupItem.HeaderText = MapText("������¡�èҡ� PO �Ѻ����ѵ�شԺ")
         Case 101
            frmAddPOSupItem.HeaderText = MapText("������¡�èҡ� PO �Ѻ�����ʴ��ػ�ó�")
         Case 102
            frmAddPOSupItem.HeaderText = MapText("������¡�èҡ� PO �Ѻ��Ҩ����͡��ʴ��ػ�ó�")
         Case 103
            frmAddPOSupItem.HeaderText = MapText("������¡�èҡ� PO �Ѻ��ҷ����")
         End Select
         Load frmAddPOSupItem
         frmAddPOSupItem.Show 1
   
         OKClick = frmAddPOSupItem.OKClick
         SupplierIDTrue = frmAddPOSupItem.SupplierIDTrue
   
         Unload frmAddPOSupItem
         Set frmAddPOSupItem = Nothing
   
         If OKClick Then
            Call GetTotalPrice
            uctlSupplierTrueLookup.MyCombo.ListIndex = IDToListIndex(uctlSupplierTrueLookup.MyCombo, SupplierIDTrue)
            GridEX1.ItemCount = CountItem(m_BillingDoc.SupItems)
            GridEX1.Rebind
         End If
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

Private Sub cmdAuto_Click()
   If Trim(txtDocumentNo.Text) = "" And ShowMode = SHOW_ADD Then
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

   If DocNoType = 1000 Then
      DOC_ID = BUY_PO_RAW
   ElseIf DocNoType = 1001 Then
      DOC_ID = BUY_PO_MATERIAL
   ElseIf DocNoType = 1002 Then
      DOC_ID = BUY_PO_EXPENSE
   ElseIf DocNoType = 1003 Then
      DOC_ID = BUY_PO_GENERAL
   ElseIf DocNoType = 1010 Then
      DOC_ID = BUY_PO_RAW_AUTO
   ElseIf DocNoType = 1011 Then
      DOC_ID = BUY_PO_MATERIAL_AUTO
   ElseIf DocNoType = 1012 Then
      DOC_ID = BUY_PO_EXPENSE_AUTO
   ElseIf DocNoType = 1013 Then
      DOC_ID = BUY_PO_GENERAL_AUTO
   ElseIf DocNoType = 100 Then
      DOC_ID = BUY_RO_RAW
   ElseIf DocNoType = 101 Then
      DOC_ID = BUY_RO_MATERIAL
   ElseIf DocNoType = 102 Then
      DOC_ID = BUY_RO_EXPENSE
   ElseIf DocNoType = 103 Then
      DOC_ID = BUY_RO_GENERAL
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
               If CheckNewMounth And CheckUniqueNs(DO_PLAN_UNIQUE, GetDocumentNo & Format(1, TempStr), ID) Then  '����ѹ���� server ��ҡѺ�ѹ��� 1 ��� ��¡�� No 1 �ѧ�����
                  GetDocumentNo = GetDocumentNo & Format(1, TempStr) '������ҡ 1 ����
                  m_BillingDoc.RUNNING_NO = 1
               Else
                  GetDocumentNo = GetDocumentNo & Format(Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd, TempStr)
                 m_BillingDoc.RUNNING_NO = Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd
               End If
          Else
               GetDocumentNo = GetDocumentNo & Format(Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd, TempStr)
                m_BillingDoc.RUNNING_NO = Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd
          End If
          m_BillingDoc.CONFIG_DOC_TYPE = DOC_ID
       Else
          GetDocumentNo = ""
       End If
    End If

End Function

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
         m_BillingDoc.SupItems.Remove (ID2)
      Else
         m_BillingDoc.SupItems.Item(ID2).Flag = "D"
      End If
      
      Call GetTotalPrice
      GridEX1.ItemCount = CountItem(m_BillingDoc.SupItems)
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
Dim ID As Long
Dim OKClick As Boolean

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(2))
   OKClick = False

   If TabStrip1.SelectedItem.Index = 1 Then
      If DocumentType = 100 Or DocumentType = 1000 Then
         frmAddEditBLImportItem.ID = ID
         frmAddEditBLImportItem.SupplierID = uctlSupplierLookup.MyCombo.ItemData(Minus2Zero(uctlSupplierLookup.MyCombo.ListIndex))
         frmAddEditBLImportItem.SupplierCode = uctlSupplierLookup.MyTextBox.Text
         frmAddEditBLImportItem.DocumentType = DocumentType
         frmAddEditBLImportItem.DocumentDate = uctlDocumentDate.ShowDate
         frmAddEditBLImportItem.COMMIT_FLAG = m_BillingDoc.OLD_COMMIT_FLAG
         Set frmAddEditBLImportItem.TempCollection = m_BillingDoc.SupItems
         frmAddEditBLImportItem.HeaderText = MapText("�����¡���ѵ�شԺ")
         frmAddEditBLImportItem.ParentShowMode = ShowMode
         frmAddEditBLImportItem.ShowMode = SHOW_EDIT
         Load frmAddEditBLImportItem
         frmAddEditBLImportItem.Show 1

         OKClick = frmAddEditBLImportItem.OKClick
         If OKClick Then
            Set TempWeight = frmAddEditBLImportItem.TempWeight
            Set CW = GetObject("CWeight", TempWeight, "1")
            If Len(CW.TRUCK_ID) > 0 Then
               txtTruckNo.Text = CW.TRUCK_ID
               If Len(CW.Time1) > 0 Then
                  uctlEntryTime.HR = HOUR(CW.Time1)
                  uctlEntryTime.MI = Minute(CW.Time1)
               End If
                If Len(CW.Time2) > 0 Then
                  uctlExitTime.HR = HOUR(CW.Time2)
                  uctlExitTime.MI = Minute(CW.Time2)
               End If
            End If
         End If
         Unload frmAddEditBLImportItem
         Set frmAddEditBLImportItem = Nothing
      ElseIf (DocumentType = 101) Or DocumentType = 1001 Then
         frmAddEditBLImportItemEx.ID = ID
         frmAddEditBLImportItemEx.SupplierID = uctlSupplierLookup.MyCombo.ItemData(Minus2Zero(uctlSupplierLookup.MyCombo.ListIndex))
         frmAddEditBLImportItemEx.COMMIT_FLAG = m_BillingDoc.OLD_COMMIT_FLAG
         Set frmAddEditBLImportItemEx.TempCollection = m_BillingDoc.SupItems
         frmAddEditBLImportItemEx.HeaderText = MapText("�����¡����ʴ��ػ�ó�")
         frmAddEditBLImportItemEx.ParentShowMode = ShowMode
         frmAddEditBLImportItemEx.ShowMode = SHOW_EDIT
         Load frmAddEditBLImportItemEx
         frmAddEditBLImportItemEx.Show 1

         OKClick = frmAddEditBLImportItemEx.OKClick

         Unload frmAddEditBLImportItemEx
         Set frmAddEditBLImportItemEx = Nothing
         ElseIf DocumentType = 102 Or DocumentType = 1002 Then
         frmAddEditBLImportItemEx2.ID = ID
         frmAddEditBLImportItemEx2.SupplierID = uctlSupplierLookup.MyCombo.ItemData(Minus2Zero(uctlSupplierLookup.MyCombo.ListIndex))
         frmAddEditBLImportItemEx2.COMMIT_FLAG = m_BillingDoc.OLD_COMMIT_FLAG
         Set frmAddEditBLImportItemEx2.TempCollection = m_BillingDoc.SupItems
         frmAddEditBLImportItemEx2.HeaderText = MapText("�����¡���Ѻ��Ҩ����͡��ʴ��ػ�ó�")
         frmAddEditBLImportItemEx2.ParentShowMode = ShowMode
         frmAddEditBLImportItemEx2.ShowMode = SHOW_EDIT
         Load frmAddEditBLImportItemEx2
         frmAddEditBLImportItemEx2.Show 1

         OKClick = frmAddEditBLImportItemEx2.OKClick

         Unload frmAddEditBLImportItemEx2
         Set frmAddEditBLImportItemEx2 = Nothing
      ElseIf (DocumentType = 103) Or DocumentType = 1003 Then
         frmAddEditBLImportItemEx3.ID = ID
         frmAddEditBLImportItemEx3.SupplierID = uctlSupplierLookup.MyCombo.ItemData(Minus2Zero(uctlSupplierLookup.MyCombo.ListIndex))
         frmAddEditBLImportItemEx3.COMMIT_FLAG = m_BillingDoc.OLD_COMMIT_FLAG
         Set frmAddEditBLImportItemEx3.TempCollection = m_BillingDoc.SupItems
         frmAddEditBLImportItemEx3.HeaderText = MapText("�����¡�âͧ������")
         frmAddEditBLImportItemEx3.ParentShowMode = ShowMode
         frmAddEditBLImportItemEx3.ShowMode = SHOW_EDIT
         Load frmAddEditBLImportItemEx3
         frmAddEditBLImportItemEx3.Show 1

         OKClick = frmAddEditBLImportItemEx3.OKClick

         Unload frmAddEditBLImportItemEx3
         Set frmAddEditBLImportItemEx3 = Nothing
      
      End If

      If OKClick Then
         Call GetTotalPrice
         GridEX1.ItemCount = CountItem(m_BillingDoc.SupItems)
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
Dim II As CSupItem

   '����ͧ��� II.EXPENSE1 + II.EXPENSE2 ��������� ���Шж١��Ш�������� txtDeliveryFee ���� (�����������ͧ����) ����Ѻ��Ѻ�ѵ�شԺ
   '�����繫������ҧ��� II.EXPENSE1 + II.EXPENSE2 ���դ���� 0 ���������Ф��� txtDeliveryFee �ͧ
   For Each II In m_BillingDoc.SupItems
      If II.Flag <> "D" Then
         II.TOTAL_INCLUDE_PRICE = II.TOTAL_ACTUAL_PRICE + (MyDiff(II.TOTAL_ACTUAL_PRICE, m_SumTotalPrice) * Val(txtDeliveryFee.Text))
         II.INCLUDE_UNIT_PRICE = MyDiffEx(II.TOTAL_INCLUDE_PRICE, II.TX_AMOUNT)

         If II.Flag <> "A" Then
            II.Flag = "E"
         End If
      End If
   Next II
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
      ID = m_BillingDoc.BILLING_DOC_ID
      m_BillingDoc.QueryFlag = 1
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


Private Sub cmdOther_Click()
Dim lMenuChosen  As Long
Dim oMenu As cPopupMenu
Dim BD As CBillingDoc
Dim TempUserName As String
Dim str  As String
Dim tempKeyRight As String
Dim verifyPoFlag As Long

   
   Set oMenu = New cPopupMenu
   
   If DocumentType = 1000 Or DocumentType = 1001 Or DocumentType = 1002 Or DocumentType = 1003 Then
      lMenuChosen = oMenu.AddMenu(glbGuiConfigs.SupItemPoOtherMenuItems)
   End If
   
   Set oMenu = Nothing
   
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
'-------------------------------------------------------------------------------------------------------------------
'    Start IF  lMenuChosen ����ö͹��ѵ� PO ��
'-------------------------------------------------------------------------------------------------------------------
   If lMenuChosen = 1 Then
      If ShowMode = SHOW_ADD Then
         Exit Sub
      End If
      
      If m_BillingDoc.PO_APPROVED_FLAG = "Y" Then
         glbErrorLog.LocalErrorMsg = "PO 㺹�����Ѻ���͹��ѵ����ǡ�͹˹�ҹ�� �������ö͹��ѵԫ���� !!!!!"
         glbErrorLog.ShowUserError
         Exit Sub
      End If

      Select Case DocumentType
      Case 1000
         str = "��¡����觫����ѵ�شԺ"
         tempKeyRight = "_RAW"
      Case 1001
         str = "��¡����觫�����ʴ��ػ�ó�"
         tempKeyRight = "_MATERIAL"
      Case 1002
         str = "��¡����觫��� �Ѻ��Ҩ����͡��ʴ��ػ�ó�"
         tempKeyRight = "_EQUIPMENT"
      Case 1003
         str = "��¡����觫��ͧ͢������"
         tempKeyRight = "_GENERAL"
      End Select
      
      If Not VerifyAccessRight("LEDGER_STOCKBUY_PO-APPROVE" & tempKeyRight, "����ö͹��ѵ� PO ��" & " " & str) Then
         frmVerifyAccRight.AccName = "LEDGER_STOCKBUY_PO-APPROVE" & tempKeyRight
         frmVerifyAccRight.AccDesc = "����ö͹��ѵ� PO ��" & " " & str
         Load frmVerifyAccRight
         frmVerifyAccRight.Show 1
         
         If frmVerifyAccRight.GrantRight Then
            TempUserName = frmVerifyAccRight.UserName
            Unload frmVerifyAccRight
            Set frmVerifyAccRight = Nothing
         Else
            Unload frmVerifyAccRight
            Set frmVerifyAccRight = Nothing
            Call EnableForm(Me, True)
            Exit Sub
         End If
      Else
         TempUserName = glbUser.USER_NAME
      End If
      
     Call LoadAuthenPO_Verify(m_AuthenPO_Verify, Trim(m_BillingDoc.DOCUMENT_TYPE), m_BillingDoc.TOTAL_PRICE)
     Call LoadAuthenPO_Approve(m_AuthenPO_Approve, Trim(m_BillingDoc.DOCUMENT_TYPE), m_BillingDoc.TOTAL_PRICE)
    
      If Len(m_BillingDoc.VERIFY_BY_NAME) = 0 Then '��� PO �ѧ����Ǩ�ͺ
         If m_AuthenPO_Verify.Count > 0 Then
            glbErrorLog.LocalErrorMsg = "PO 㺹���ͧ���Ѻ��õ�Ǩ�ͺ��͹ !!!!!"
            glbErrorLog.ShowUserError
            Exit Sub
         End If
      End If
     
      verifyPoFlag = GetAuthenPO(m_AuthenPO_Approve, Trim(TempUserName), m_BillingDoc.TOTAL_PRICE)
      If verifyPoFlag = 0 Then
         glbErrorLog.LocalErrorMsg = "�س " & TempUserName & " ������Է��͹��ѵ� PO 㺹��"
         glbErrorLog.ShowUserError
         
         frmVerifyAccRight.AccName = "LEDGER_STOCKBUY_PO-APPROVE" & tempKeyRight
         frmVerifyAccRight.AccDesc = "����ö͹��ѵ� PO ��" & " " & str
         Load frmVerifyAccRight
         frmVerifyAccRight.Show 1
         
         If frmVerifyAccRight.GrantRight Then
            TempUserName = frmVerifyAccRight.UserName
            Unload frmVerifyAccRight
            Set frmVerifyAccRight = Nothing
            verifyPoFlag = GetAuthenPO(m_AuthenPO_Approve, Trim(TempUserName), m_BillingDoc.TOTAL_PRICE)
         Else
            Unload frmVerifyAccRight
            Set frmVerifyAccRight = Nothing
            Call EnableForm(Me, True)
            Exit Sub
         End If
      End If
      
      If verifyPoFlag = 2 Then
                '͹��ѵ� Po ��ҡѺ��÷��Դ PO �¡�����ҧ�ѵ���ѵ�
               '���� Close � RO ����
               Call glbDaily.StartTransaction
               Set BD = New CBillingDoc
               BD.BILLING_DOC_ID = ID
               Call BD.UpdatePoApprovedFlag(TempUserName)
               Call glbDaily.CommitTransaction
               glbErrorLog.LocalErrorMsg = "͹��ѵ������"
               glbErrorLog.ShowUserError
      ElseIf verifyPoFlag = 1 Then
               glbErrorLog.LocalErrorMsg = "PO 㺹���ѧ�������ª��ͼ��͹��ѵ� ��سҵ�駤�Ҽ�����Է���͹��ѵԡ�͹"
               glbErrorLog.ShowUserError
      End If
'-------------------------------------------------------------------------------------------------------------------
'    Start IF  lMenuChosen ����ö¡��ԡ���͹��ѵ� PO ��
'-------------------------------------------------------------------------------------------------------------------
   ElseIf lMenuChosen = 3 Then
      If ShowMode = SHOW_ADD Then
         Exit Sub
      End If
      If m_BillingDoc.PO_APPROVED_FLAG = "N" Then
            glbErrorLog.LocalErrorMsg = "PO 㺹����١¡��ԡ����͡�͹˹�ҹ������"
            glbErrorLog.ShowUserError
         Exit Sub
      End If
      
      If Not VerifyAccessRight("LEDGER_STOCKBUY_PO-APPROVE-CANCEL", "����ö¡��ԡ ���͹��ѵ� PO ��") Then
         frmVerifyAccRight.AccName = "LEDGER_STOCKBUY_PO-APPROVE-CANCEL"
         frmVerifyAccRight.AccDesc = "����ö¡��ԡ ���͹��ѵ� PO ��"
         Load frmVerifyAccRight
         frmVerifyAccRight.Show 1
         
         If frmVerifyAccRight.GrantRight Then
            TempUserName = frmVerifyAccRight.UserName
            Unload frmVerifyAccRight
            Set frmVerifyAccRight = Nothing
         Else
            Unload frmVerifyAccRight
            Set frmVerifyAccRight = Nothing
            Call EnableForm(Me, True)
            Exit Sub
         End If
      Else
         TempUserName = glbUser.USER_NAME
      End If
   If Not VerifyAccessRight("LEDGER_STOCKBUY_PO-APPROVE-CANCEL_OTHER-NAME", "����ö¡��ԡ ���͹��ѵ� PO �¼�������") Then
       If m_BillingDoc.APPROVE_NAME = TempUserName Then
         '¡��ԡ ͹��ѵ� Po
         Call glbDaily.StartTransaction
         Set BD = New CBillingDoc
         BD.BILLING_DOC_ID = ID
         Call BD.UpdatePoCancelApprovedFlag(TempUserName)
         
         Call glbDaily.CommitTransaction
         
         glbErrorLog.LocalErrorMsg = "¡��ԡ ͹��ѵ������ !!!!!!!"
         glbErrorLog.ShowUserError
       Else
         glbErrorLog.LocalErrorMsg = "�س�������ö¡��ԡ���͹��ѵ� PO 㺹�������ͧ�ҡ �س������繼��͹��ѵ� PO 㺹�� !!!!!!!"
         glbErrorLog.ShowUserError
       End If
  Else
      Call glbDaily.StartTransaction
      Set BD = New CBillingDoc
      BD.BILLING_DOC_ID = ID
      Call BD.UpdatePoCancelApprovedFlag(TempUserName)
      Call glbDaily.CommitTransaction
      glbErrorLog.LocalErrorMsg = "¡��ԡ ͹��ѵ������ !!!!!!!"
      glbErrorLog.ShowUserError
  End If
   
    
'-------------------------------------------------------------------------------------------------------------------
'    Start IF  lMenuChosen ����ö�Դ PO ��
'-------------------------------------------------------------------------------------------------------------------
    ElseIf lMenuChosen = 5 Then
      If ShowMode = SHOW_ADD Then
       Exit Sub
      End If
      
        If Not VerifyAccessRight("LEDGER_STOCKBUY_PO-CLOSE", "����ö�Դ PO ��") Then
            frmVerifyAccRight.AccName = "LEDGER_STOCKBUY_PO-CLOSE"
            frmVerifyAccRight.AccDesc = "����ö�Դ PO ��"
            Load frmVerifyAccRight
            frmVerifyAccRight.Show 1
         
         If frmVerifyAccRight.GrantRight Then
            TempUserName = frmVerifyAccRight.UserName
            Unload frmVerifyAccRight
            Set frmVerifyAccRight = Nothing
         Else
            Unload frmVerifyAccRight
            Set frmVerifyAccRight = Nothing
            Call EnableForm(Me, True)
            Exit Sub
         End If
      Else
         TempUserName = glbUser.USER_NAME
      End If
      
      Call glbDaily.StartTransaction
      Set BD = New CBillingDoc
      BD.BILLING_DOC_ID = ID
      Call BD.UpdateClosePOFlag(TempUserName)
      Call glbDaily.CommitTransaction
      
      glbErrorLog.LocalErrorMsg = "�Դ PO ����� !!!!!"
      glbErrorLog.ShowUserError
'-------------------------------------------------------------------------------------------------------------------
'    Start IF  lMenuChosen ¡��ԡ�Դ PO ��
'-------------------------------------------------------------------------------------------------------------------
   ElseIf lMenuChosen = 7 Then
      If ShowMode = SHOW_ADD Then
       Exit Sub
      End If
      If m_BillingDoc.CLOSE_FLAG = "N" Then
         glbErrorLog.LocalErrorMsg = "������ �������ö  ¡��ԡ��¡�ûԴ PO ��  ( ���ͧ�ҡ PO 㺹���ѧ���١�Դ)"
         glbErrorLog.ShowUserError
         Exit Sub
      End If
      
        If Not VerifyAccessRight("LEDGER_STOCKBUY_PO-CANCLE", "����ö¡��ԡ ��ûԴ PO ��") Then
         frmVerifyAccRight.AccName = "LEDGER_STOCKBUY_PO-CANCLE"
         frmVerifyAccRight.AccDesc = "����ö¡��ԡ ��ûԴ PO ��"
         Load frmVerifyAccRight
         frmVerifyAccRight.Show 1
         
         If frmVerifyAccRight.GrantRight Then
            TempUserName = frmVerifyAccRight.UserName
            Unload frmVerifyAccRight
            Set frmVerifyAccRight = Nothing
         Else
            Unload frmVerifyAccRight
            Set frmVerifyAccRight = Nothing
            Call EnableForm(Me, True)
            Exit Sub
         End If
      Else
         TempUserName = glbUser.USER_NAME
      End If
      
      Call glbDaily.StartTransaction
      Set BD = New CBillingDoc
      BD.BILLING_DOC_ID = ID
      Call BD.UpdateCanclePOFlag(TempUserName)
      Call glbDaily.CommitTransaction
      
      glbErrorLog.LocalErrorMsg = "¡��ԡ��ûԴ PO ����� !!!!!"
      glbErrorLog.ShowUserError

'-------------------------------------------------------------------------------------------------------------------
'    Start IF  lMenuChosen ����ö��Ǩ�ͺ PO ��
'-------------------------------------------------------------------------------------------------------------------
   ElseIf lMenuChosen = 9 Then
      If ShowMode = SHOW_ADD Then
         Exit Sub
      End If
      If m_BillingDoc.PO_APPROVED_FLAG = "Y" Then '��� PO ͹��ѵ����� ����ͧ��Ǩ�ͺ
         glbErrorLog.LocalErrorMsg = "PO ������Ѻ���͹��ѵ������������ö�ӡ�õ�Ǩ�ͺ PO �� !!!!!"
         glbErrorLog.ShowUserError
         Exit Sub
     End If
      If m_BillingDoc.CLOSE_FLAG = "Y" Then '��� PO �Դ���� ����ͧ��Ǩ�ͺ
         glbErrorLog.LocalErrorMsg = "PO �����١�Դ�����������ö¡��ԡ��÷ӡ�õ�Ǩ�ͺ PO�� !!!!!"
         glbErrorLog.ShowUserError
         Exit Sub
     End If
      If Len(m_BillingDoc.VERIFY_BY_NAME) > 0 Then '��� PO ��Ǩ�ͺ���� ����ͧ��Ǩ�ͺ
         glbErrorLog.LocalErrorMsg = "PO �����١��Ǩ�ͺ��͹˹�ҹ������ !!!!!"
         glbErrorLog.ShowUserError
         Exit Sub
      Else '��� PO�ѧ����ͧ��Ǩ�ͺ
         Call LoadAuthenPO_Verify(m_AuthenPO_Verify, Trim(m_BillingDoc.DOCUMENT_TYPE), m_BillingDoc.TOTAL_PRICE)
         
         verifyPoFlag = GetAuthenPO(m_AuthenPO_Verify, Trim(glbUser.USER_NAME), m_BillingDoc.TOTAL_PRICE)
         If verifyPoFlag = 2 Then
            Call glbDaily.StartTransaction
            Set BD = New CBillingDoc
            BD.BILLING_DOC_ID = ID
            Call BD.UpdateCanclePOVerify(glbUser.USER_NAME, 1)
            Call glbDaily.CommitTransaction
            glbErrorLog.LocalErrorMsg = "��Ǩ�ͺ PO ���º���� !!!!!"
            glbErrorLog.ShowUserError
         ElseIf m_AuthenPO_Verify.Count <= 0 Then
           glbErrorLog.LocalErrorMsg = "PO 㺹������ͧ�ռ���Ǩ�ͺ !!!!!"
           glbErrorLog.ShowUserError
         Else
           glbErrorLog.LocalErrorMsg = "�س������Է����Ǩ�ͺ PO 㺹�� !!!!!"
           glbErrorLog.ShowUserError
         End If
      End If
   '-------------------------------------------------------------------------------------------------------------------
'    Start IF  lMenuChosen ����ö¡��ԡ��Ǩ�ͺ PO ��
'-------------------------------------------------------------------------------------------------------------------
 ElseIf lMenuChosen = 11 Then
      If ShowMode = SHOW_ADD Then
         Exit Sub
      End If
      If m_BillingDoc.PO_APPROVED_FLAG = "Y" Then '��� PO ͹��ѵ����� ¡��ԡ��õ�Ǩ�ͺ�����
         glbErrorLog.LocalErrorMsg = "PO ������Ѻ���͹��ѵ������������ö¡��ԡ��õ�Ǩ�ͺ�� !!!!!"
         glbErrorLog.ShowUserError
         Exit Sub
     End If
      If m_BillingDoc.CLOSE_FLAG = "Y" Then '��� PO �Դ���� ¡��ԡ��õ�Ǩ�ͺ�����
         glbErrorLog.LocalErrorMsg = "PO �����Դ����������������ö¡��ԡ��õ�Ǩ�ͺ�� !!!!!"
         glbErrorLog.ShowUserError
         Exit Sub
     End If
      If Len(m_BillingDoc.VERIFY_BY_NAME) > 0 Then '��� PO ��Ǩ�ͺ���� ����ö���¡��ԡ�� ����ͼ��¡��ԡ��ͧ�ç�Ѻ����Ǩ�ͺ�͹�á
         If Trim(m_BillingDoc.VERIFY_BY_NAME) = Trim(glbUser.USER_NAME) Then
            Call glbDaily.StartTransaction
            Set BD = New CBillingDoc
            BD.BILLING_DOC_ID = ID
            Call BD.UpdateCanclePOVerify(glbUser.USER_NAME, 2)
            Call glbDaily.CommitTransaction
            glbErrorLog.LocalErrorMsg = "¡��ԡ��õ�Ǩ�ͺ PO ���º���� !!!!!"
            glbErrorLog.ShowUserError
         Else
           glbErrorLog.LocalErrorMsg = "�س������Է���¡��ԡ��õ�Ǩ�ͺ PO 㺹�����Ъ������ç�Ѻ����Ǩ�ͺ��͹˹�� !!!!!"
           glbErrorLog.ShowUserError
         End If
      Else '��� PO�ѧ����ͧ��Ǩ�ͺ����͡�
         glbErrorLog.LocalErrorMsg = "PO 㺹���ѧ������Ѻ��õ�Ǩ�ͺ �������ö¡��ԡ��õ�Ǩ�ͺ�� !!!!!"
         glbErrorLog.ShowUserError
         Exit Sub
      End If
   End If
End Sub
Private Sub cmdPrint_Click()
Dim Report As CReportInterface
Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
Dim ReportKey As String
Dim ReportFlag As Boolean
Dim Rc As CReportConfig
Dim iCount As Long
Dim EditMode As SHOW_MODE_TYPE
Dim ReportMode As Long
   
   ReportMode = 1
   
   If m_HasModify Or ((Not m_HasModify) And (ShowMode = SHOW_ADD)) Then
      glbErrorLog.LocalErrorMsg = "��سҷӡ�úѹ�֡������������º���¡�͹"
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
  
'   If DocumentType = 1000 Or DocumentType = 1001 Or DocumentType = 1002 Or DocumentType = 1003 Then
'      lMenuChosen = oMenu.Popup("���§ҹ���觫���", "��Ѻ���˹�ҡ�д��", "���§ҹ���觫���ŧ��д������", "��Ѻ���˹�ҡ�д��", "���§ҹ���觫���ŧ��д������ A4", "��Ѻ���˹�ҡ�д��")
'   ElseIf DocumentType = 100 Or DocumentType = 101 Or DocumentType = 102 Or DocumentType = 103 Then
'      lMenuChosen = oMenu.Popup("���§ҹ�Ѻ�ѵ�شԺ", "���§ҹ�Ѻ�ѵ�شԺ(MGP)", "���§ҹ�Ѻ�ѵ�شԺ (����ʴ��Ҥ�)", "��Ѻ���˹�ҡ�д��", "-", "���§ҹ�Ѻ�ͧ", "��Ѻ���˹�ҡ�д��")
'   End If
   Set oMenu = New cPopupMenu
   Select Case DocumentType
   Case 100
      lMenuChosen = oMenu.Popup("���§ҹ�Ѻ�ѵ�شԺ", "���§ҹ�Ѻ�ѵ�شԺ(MGP)", "���§ҹ�Ѻ�ѵ�شԺ (����ʴ��Ҥ�)", "��Ѻ���˹�ҡ�д��", "-", "���§ҹ�Ѻ�ͧ", "��Ѻ���˹�ҡ�д��", "���§ҹ�Ѻ�ѵ�شԺ(MGP-�ѾϨ�ԧ)", "���§ҹ�Ѻ�ѵ�شԺ(����ʴ��Ҥ�)(MGP)")
   Case 101
      lMenuChosen = oMenu.Popup("���§ҹ�Ѻ�����ʴ��ػ�ó�", "���§ҹ�Ѻ�����ʴ��ػ�ó�(MGP)", "���§ҹ�Ѻ�����ʴ��ػ�ó� (����ʴ��Ҥ�)", "��Ѻ���˹�ҡ�д��", "-", "���§ҹ�Ѻ�ͧ", "��Ѻ���˹�ҡ�д��")
   Case 102
      lMenuChosen = oMenu.Popup("���§ҹ�Ѻ��Ҩ����͡��ʴ��ػ�ó�", "���§ҹ�Ѻ��Ҩ����͡��ʴ��ػ�ó�(MGP)", "���§ҹ�Ѻ��Ҩ����͡��ʴ��ػ�ó� (����ʴ��Ҥ�)", "��Ѻ���˹�ҡ�д��", "-", "���§ҹ�Ѻ�ͧ", "��Ѻ���˹�ҡ�д��")
   Case 103
      lMenuChosen = oMenu.Popup("���§ҹ�Ѻ��Ңͧ������", "���§ҹ�Ѻ��Ңͧ������(MGP)", "���§ҹ�Ѻ��Ңͧ������ (����ʴ��Ҥ�)", "��Ѻ���˹�ҡ�д��", "-", "���§ҹ�Ѻ�ͧ", "��Ѻ���˹�ҡ�д��")
   Case 1000, 1001, 1002, 1003
'      lMenuChosen = oMenu.Popup("���§ҹ���觫���", "��Ѻ���˹�ҡ�д��", "���§ҹ���觫���ŧ��д������ A5 (MGP)", "���§ҹ���觫���ŧ��д������ A5", "��Ѻ���˹�ҡ�д��", "���§ҹ���觫���ŧ��д������ A4", "��Ѻ���˹�ҡ�д��", "���§ҹ��������´����Ѻ��ҵ�� PO", "��Ѻ���˹�ҡ�д��")
      lMenuChosen = oMenu.Popup("���§ҹ���觫���", "��Ѻ���˹�ҡ�д��", "���§ҹ���觫���ŧ��д������ A5 (MGP)", "���§ҹ���觫���ŧ��д������ A5", "��Ѻ���˹�ҡ�д��", "���§ҹ���觫���ŧ��д������ A4", "��Ѻ���˹�ҡ�д��", "���§ҹ��������´����Ѻ��ҵ�� PO", "��Ѻ���˹�ҡ�д��", "���§ҹ���觫���ŧ��д������ A5 (MGP-�ѾϨ�ԧ)")
      'lMenuChosen = oMenu.Popup("���§ҹ���觫���", "��Ѻ���˹�ҡ�д��", "���§ҹ���觫���ŧ��д������ A5", "��Ѻ���˹�ҡ�д��", "���§ҹ���觫���ŧ��д������ A4", "��Ѻ���˹�ҡ�д��")
   End Select
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   
   If DocumentType = 1000 Or DocumentType = 1001 Or DocumentType = 1002 Or DocumentType = 1003 Then
      If lMenuChosen = 1 Then
         ReportKey = "CReportFormPO002"
         
         Set Report = New CReportFormPO002
         ReportFlag = True
         Call Report.AddParam(1, "PREVIEW_TYPE")
      ElseIf lMenuChosen = 2 Then
         ReportKey = "CReportFormPO002"
   
         Set Rc = New CReportConfig
         Rc.REPORT_KEY = ReportKey
         Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
         Call Rc.QueryData(m_Rs, iCount)
         HeaderText = MapText("���§ҹ���觫���")
         If Not m_Rs.EOF Then
            Call Rc.PopulateFromRS(1, m_Rs)
            EditMode = SHOW_EDIT
         Else
            EditMode = SHOW_ADD
         End If
     ElseIf lMenuChosen = 3 Then 'MGP
         ReportKey = "CReportNormalPO2"
            
         Set Report = New CReportNormalPO2
         ReportFlag = True
         Call Report.AddParam(2, "PREVIEW_TYPE")
       ElseIf lMenuChosen = 4 Then
         ReportKey = "CReportNormalPO2"
            
         Set Report = New CReportNormalPO2
         ReportFlag = True
         Call Report.AddParam(1, "PREVIEW_TYPE")
      ElseIf lMenuChosen = 5 Then
          ReportKey = "CReportNormalPO2"
   
         Set Rc = New CReportConfig
         Rc.REPORT_KEY = ReportKey
         Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
         Call Rc.QueryData(m_Rs, iCount)
         HeaderText = MapText("���§ҹ���觫���")
         If Not m_Rs.EOF Then
            Call Rc.PopulateFromRS(1, m_Rs)
            EditMode = SHOW_EDIT
         Else
            EditMode = SHOW_ADD
         End If
      ElseIf lMenuChosen = 6 Then
           ReportKey = "CReportNormalPO3"
            
            Set Report = New CReportNormalPO3
            ReportFlag = True
            Call Report.AddParam(1, "PREVIEW_TYPE")
      ElseIf lMenuChosen = 7 Then
          ReportKey = "CReportNormalPO3"
   
         Set Rc = New CReportConfig
         Rc.REPORT_KEY = ReportKey
         Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
         Call Rc.QueryData(m_Rs, iCount)
         HeaderText = MapText("���§ҹ���觫���")
         If Not m_Rs.EOF Then
            Call Rc.PopulateFromRS(1, m_Rs)
            EditMode = SHOW_EDIT
         Else
            EditMode = SHOW_ADD
         End If
      ElseIf lMenuChosen = 8 Then
           ReportKey = "CReportNormalPO4"
            
            Set Report = New CReportNormalPO4
            ReportFlag = True
            Call Report.AddParam(1, "PREVIEW_TYPE")
      ElseIf lMenuChosen = 9 Then
          ReportKey = "CReportNormalPO4"
   
         Set Rc = New CReportConfig
         Rc.REPORT_KEY = ReportKey
         Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
         Call Rc.QueryData(m_Rs, iCount)
         HeaderText = MapText("���§ҹ���觫���")
         If Not m_Rs.EOF Then
            Call Rc.PopulateFromRS(1, m_Rs)
            EditMode = SHOW_EDIT
         Else
            EditMode = SHOW_ADD
         End If
       ElseIf lMenuChosen = 10 Then
         ReportKey = "CReportNormalPO2"
            
         Set Report = New CReportNormalPO2
         ReportFlag = True
         Call Report.AddParam(lMenuChosen, "PREVIEW_TYPE")
      End If
      
   ElseIf DocumentType = 100 Or DocumentType = 101 Or DocumentType = 102 Or DocumentType = 103 Then
      If CountItem(m_BillingDoc.SupItems) <> 1 And (lMenuChosen >= 0 And lMenuChosen <= 3) Then
         glbErrorLog.LocalErrorMsg = "���§ҹ�Ѻ�ͧ�е�ͧ����¡���Ѻ�������ҡѺ 1 ��¡��"
         glbErrorLog.ShowUserError
            
         Call EnableForm(Me, True)
         Exit Sub
      End If
         
         
      If lMenuChosen = 1 Then
         ReportKey = "CReportInvDoc001_1"
      
         Set Report = New CReportInvDoc001_1
         ReportFlag = True
         Call Report.AddParam(1, "PREVIEW_TYPE")
      ElseIf lMenuChosen = 2 Then
          ReportKey = "CReportInvDoc001_1"
      
          Set Report = New CReportInvDoc001_1
          ReportFlag = True
         
         Call Report.AddParam(2, "PREVIEW_TYPE")
      ElseIf lMenuChosen = 3 Then
          ReportKey = "CReportInvDoc001_1"
      
          Set Report = New CReportInvDoc001_1
          ReportFlag = True
         
         Call Report.AddParam(3, "PREVIEW_TYPE")
      ElseIf lMenuChosen = 4 Then
         ReportKey = "CReportInvDoc001_1"
      
         Set Rc = New CReportConfig
         Rc.REPORT_KEY = ReportKey
         Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
         Call Rc.QueryData(m_Rs, iCount)
         HeaderText = MapText("��Ѻ����Թ���/�ѵ�شԺ")
         If Not m_Rs.EOF Then
            Call Rc.PopulateFromRS(1, m_Rs)
            EditMode = SHOW_EDIT
         Else
            EditMode = SHOW_ADD
         End If
      ElseIf lMenuChosen = 6 Then
         Call LoadPictureFromFile(glbParameterObj.ReceiveVoucher1, Picture2)

         ReportKey = "CReportInvDoc001_2"
         Set Report = New CReportInvDoc001_2

         ReportFlag = True
         Call Report.AddParam(1, "PREVIEW_TYPE")
      ElseIf lMenuChosen = 7 Then
         ReportKey = "CReportInvDoc001_2"
         
         Set Rc = New CReportConfig
         Rc.REPORT_KEY = ReportKey
         Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
         Call Rc.QueryData(m_Rs, iCount)
         HeaderText = MapText("��Ѻ����Թ���/�ѵ�شԺ")
         If Not m_Rs.EOF Then
            Call Rc.PopulateFromRS(1, m_Rs)
            EditMode = SHOW_EDIT
         Else
            EditMode = SHOW_ADD
         End If
      ElseIf lMenuChosen = 8 Then
         ReportKey = "CReportInvDoc001_1"
      
         Set Report = New CReportInvDoc001_1
         ReportFlag = True
         Call Report.AddParam(1, "PREVIEW_TYPE")
      ElseIf lMenuChosen = 9 Then
          ReportKey = "CReportInvDoc001_1"
      
          Set Report = New CReportInvDoc001_1
          ReportFlag = True
         
         Call Report.AddParam(3, "PREVIEW_TYPE")
      End If
   End If
   
   If Not Report Is Nothing Then
   
      Call Report.AddParam(m_BillingDoc.INVENTORY_DOC_ID, "INVENTORY_DOC_ID")
       Call Report.AddParam(lMenuChosen, "REPORT_TYPE")
      Call Report.AddParam(m_BillingDoc.BILLING_DOC_ID, "BILLING_DOC_ID")
      Call Report.AddParam(ReportKey, "REPORT_KEY")
      Call Report.AddParam(Picture2.Picture, "BACK_GROUND")
   End If
   
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
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      Call LoadSupplier(uctlSupplierLookup.MyCombo, m_Suppliers)
      Set uctlSupplierLookup.MyCollection = m_Suppliers
      
      Call LoadSupplier(uctlSupplierTrueLookup.MyCombo, m_SuppliersTrue)
      Set uctlSupplierTrueLookup.MyCollection = m_SuppliersTrue
      
      
      Call LoadLayout(cboDepartMent)
      Call LoadMaster(cboCondition, , CONDITION)
      Call LoadMaster(cboPaidType, , PAID_TYPE)
      Call LoadConfigDoc(Nothing, m_Cd)
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_BillingDoc.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
         
      ElseIf ShowMode = SHOW_ADD Then
         '''Call cmdAuto_Click
'         uctlSupplierLookup.SetFocus
         uctlDocumentDate.ShowDate = Now
         uctlDueDate.ShowDate = Now
         uctlEntryTime.HR = HOUR(Now)
         uctlEntryTime.MI = Minute(Now)
         
         uctlExitTime.HR = HOUR(Now)
         uctlExitTime.MI = Minute(Now)
         m_BillingDoc.QueryFlag = 0
         Call QueryData(False)
      End If
'      Call LoadAuthenPO(m_AuthenPO_Verify, , , Trim(m_BillingDoc.DOCUMENT_NO), Trim(m_BillingDoc.DOCUMENT_TYPE))
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
  cmdOther.Top = ScaleHeight - 580
  
  
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
    cmdPrint.Left = cmdOK.Left - cmdPrint.Width - 50
    cmdOther.Left = cmdPrint.Left - cmdOther.Width - 50
End Sub
Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_BillingDoc = Nothing
   Set m_Suppliers = Nothing
   Set m_SuppliersTrue = Nothing
   Set m_Cd = Nothing
   Set m_AuthenPO_Verify = Nothing
   Set m_AuthenPO_Approve = Nothing
   Set TempWeight = Nothing
   Set m_Weight = Nothing
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
   Col.Width = 10
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 10
   Col.Caption = "Real ID"
   
   Select Case DocumentType
    Case 1000, 100
             Set Col = GridEX1.Columns.add '3
            Col.Width = 2100
            Col.Caption = MapText("�����Ţ�ѵ�شԺ")
            
            Set Col = GridEX1.Columns.add '4
            Col.Width = 4425
            Col.Caption = MapText("�ѵ�شԺ")
      Case 1001, 101
            Set Col = GridEX1.Columns.add '3
            Col.Width = 2100
            Col.Caption = MapText("�����Ţ��ʴ��ػ�ó�")
            
            Set Col = GridEX1.Columns.add '4
            Col.Width = 4425
            Col.Caption = MapText("��ʴ��ػ�ó�")
     Case 1002, 102
            Set Col = GridEX1.Columns.add '3
            Col.Width = 2100
            Col.Caption = MapText("�����Ţ�Ѻ��Ҩ����͡��ʴ��ػ�ó�")
            
            Set Col = GridEX1.Columns.add '4
            Col.Width = 4425
            Col.Caption = MapText("�Ѻ��Ҩ����͡��ʴ��ػ�ó�")
      Case 1003, 103
            Set Col = GridEX1.Columns.add '3
            Col.Width = 2100
            Col.Caption = MapText("�����Ţ�ͧ������")
            
            Set Col = GridEX1.Columns.add '4
            Col.Width = 4425
            Col.Caption = MapText("�ͧ������")
      End Select

   Set Col = GridEX1.Columns.add '5
   Col.Width = 1785
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("����ҳ")

   Set Col = GridEX1.Columns.add '6
   Col.Width = 1620
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("�Ҥ�")
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = 1620
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("�Ҥ����")
   
   Set Col = GridEX1.Columns.add '8
   Col.Width = 1980
   Col.Caption = MapText("ʶҹ���Ѵ��")
   
   Set Col = GridEX1.Columns.add '9
   Col.Width = 1980
   Col.Caption = MapText("PO NO")
End Sub

Private Sub GetTotalPrice()
Dim II As CSupItem
Dim Sum As Double
Dim Sum1 As Double
   
   Sum1 = 0
   Sum = 0
   m_SumUnit = 0
   m_SumTotalPrice = 0
   
   For Each II In m_BillingDoc.SupItems
      If II.Flag <> "D" Then
         Sum = Sum + CDbl(Format(II.TOTAL_ACTUAL_PRICE, "0.00"))
         m_SumUnit = m_SumUnit + II.TX_AMOUNT
         m_SumTotalPrice = m_SumTotalPrice + II.TOTAL_ACTUAL_PRICE
         
         Sum1 = Sum1 + II.EXPENSE1 + II.EXPENSE2
               
'''         If DocumentType = 1000 Then    '���§ҹ�Ѻ�ѵ�شԺ ����¡������          '���������¡������� �� ��ҷ����仴ٷ�� MM ���պҧ��¡���� �� ���ҡ���� 1 ��¡��
'''            'cmdAdd.Enabled = False
'''         End If
'''         If DocumentType = 100 Or DocumentType = 101 Or DocumentType = 102 Or DocumentType = 103 Then
'''            If II.PO_ID > 0 Then
'''               cmdAdd.Enabled = False
'''            End If
'''         End If
            
      End If
   Next II
   
   txtMaterialPrice.Text = Format(Sum, "0.00")
   txtTotalAmount.Text = Format(m_SumUnit, "0.00")
   If (DocumentType = 100 Or DocumentType = 1000) Then
      txtDeliveryFee.Text = Format(Sum1, "0.00")
   End If
End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame4.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblTruckNo, MapText("����¹ö"))
   Call InitNormalLabel(lblQueNo, MapText("��Ƿ��"))
   Call InitNormalLabel(lblReceiver, MapText("���������"))
   Call InitNormalLabel(lblDesc, MapText("�����˵�"))
   Call InitNormalLabel(lblDoNo, MapText("�Ţ��� PO"))
   Call InitNormalLabel(lblDeliveryNo, MapText("������� - �͡"))
   Call InitNormalLabel(Label3, MapText("-"))
   Call InitNormalLabel(lblDepartment, MapText("Ἱ�"))
   Call InitNormalLabel(lblCredit, MapText("�ôԵ"))
   Call InitNormalLabel(Label6, MapText("�ѹ"))
   
   Call InitNormalLabel(lblPrNo, MapText("�Ţ��� PR"))
   Call InitNormalLabel(lblSender, MapText("�Ţ�����觢ͧ"))
   
   Call InitNormalLabel(lblCondition, MapText("���͹���ѧ�Ѻ�Թ���"))
   Call InitNormalLabel(lblPaidType, MapText("��ê���� PO"))
   
   Call txtDueAmount.SetTextLenType(TEXT_INTEGER, glbSetting.MONEY_TYPE)
   Call InitNormalLabel(Label5, MapText("�ѹ"))
   If DocumentType = 1000 Or DocumentType = 1001 Or DocumentType = 1002 Or DocumentType = 1003 Then
      Call InitNormalLabel(lblDueDate, MapText("�ѹ����ͧ���"))
      Call InitNormalLabel(lblDocumentNo, MapText("�Ţ������觫���"))
   Else
      Call InitNormalLabel(lblDocumentNo, MapText("�Ţ������Ѻ�ͧ"))
      Call InitNormalLabel(lblDueDate, MapText("�ѹ�ú��˹�"))
      txtDueAmount.Enabled = False
   End If
   
   uctlSupplierLookup.MyTextBox.SetKeySearch ("SUPPLIER_CODE")
   
'  If (DocumentType = 101) Or (DocumentType = 102) Then
'  ' If (DocumentType = 101) Or (DocumentType = 102) Or (DocumentType = 1001) Or (DocumentType = 1002) Or (DocumentType = 1003) Then
'      Call InitNormalLabel(lblDeliveryFee, MapText("��Ť�� VAT"))
'   Else
'      Call InitNormalLabel(lblDeliveryFee, MapText("�������¨Ѵ����"))
'   End If
   Call InitNormalLabel(lblDeliveryFee, MapText("��Ť�� VAT"))
   Call InitNormalLabel(lblDocumentDate, MapText("�ѹ����͡���"))
   Call InitNormalLabel(lblMaterialPrice, MapText("�Ҥ��Թ���"))
   Call InitNormalLabel(Label1, MapText("�ҷ"))
   Call InitNormalLabel(Label2, MapText("�ҷ"))
   Call InitNormalLabel(Label4, MapText("�ҷ"))
   Call InitNormalLabel(lblTotal, MapText("�Ҥ����"))
   Call InitNormalLabel(lblVolume, MapText("����ҳ���"))
   Call InitNormalLabel(lblSupplierNo, MapText("���ʫѾ�"))
   Call InitNormalLabel(lblSupplierTrueNo, MapText("���ʫѾϨ�ԧ"))
'''   Call InitCheckBox(chkCommit, "�ӹǳ")
   Call InitCheckBox(chkException, "***")
   Call InitCheckBox(chkClose, "** �Դ PO ����㺹�� **")
   Call InitCheckBox(chkGenCommitFlag, "�Դ��͹����͡������� PO")
   Call InitCheckBox(chkDeliveryFee, "�Դ����")
   
   chkGenCommitFlag.Visible = False
   
   If DocumentType = 1000 Or DocumentType = 1001 Or DocumentType = 1002 Or DocumentType = 1003 Then
       chkClose.Visible = True
       txtDoNo.Enabled = False
   Else
        chkClose.Visible = False
   End If
   Call txtDocumentNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtDoNo.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtTruckNo.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtDeliveryFee.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtDeliveryFee.Enabled = True
   Call txtCredit.SetTextLenType(TEXT_INTEGER, glbSetting.MONEY_TYPE)
   Call txtMaterialPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtMaterialPrice.Enabled = False
   Call txtReceiver.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtTotal.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtTotal.Enabled = False
   Call txtTotalAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtTotalAmount.Enabled = False
   Call txtQueNo.SetTextLenType(TEXT_INTEGER, glbSetting.MONEY_TYPE)
   Call txtDesc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   
   Call InitCombo(cboDepartMent)
   Call InitCombo(cboCondition)
   Call InitCombo(cboPaidType)
   
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
   cmdPrint.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOther.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdPrint, MapText("�����"))
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
   Call InitMainButton(cmdAdd, MapText("���� (F7)"))
   Call InitMainButton(cmdEdit, MapText("��� (F3)"))
   Call InitMainButton(cmdDelete, MapText("ź (F6)"))
   Call InitMainButton(cmdAuto, MapText("A"))
   Call InitMainButton(cmdOther, MapText("����"))
   
   Call InitGrid1
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.NAME = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   Dim str As String
   Select Case DocumentType
    Case 1000
          str = "��¡����觫����ѵ�شԺ"
      Case 1001
            str = "��¡����觫�����ʴ��ػ�ó�"
     Case 1002
            str = "��¡����觫��� �Ѻ��Ҩ����͡��ʴ��ػ�ó�"
      Case 1003
           str = "��¡����觫��ͧ͢������"
      Case 100
          str = "��¡���Ѻ�ͧ�ѵ�شԺ"
      Case 101
            str = "��¡���Ѻ�ͧ��ʴ��ػ�ó�"
     Case 102
            str = "��¡���Ѻ�ͧ �Ѻ��Ҩ����͡��ʴ��ػ�ó�"
      Case 103
           str = "��¡���Ѻ�ͧ������"
      End Select
   TabStrip1.Tabs.add().Caption = MapText(str)
   If DocumentType = 1000 Or DocumentType = 1001 Or DocumentType = 1002 Or DocumentType = 1003 Then
      TabStrip1.Tabs.add().Caption = MapText("��������´�����")
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
   OKClick = False
   Call InitFormLayout
      
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_BillingDoc = New CBillingDoc
   Set m_Weight = New CWeight
   Set m_Suppliers = New Collection
   Set m_SuppliersTrue = New Collection
   Set m_Cd = New Collection
   Set m_AuthenPO_Verify = New Collection
   Set m_AuthenPO_Approve = New Collection
   Set TempWeight = New Collection
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
      If m_BillingDoc.SupItems Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If
      
      Dim CR As CSupItem
      If m_BillingDoc.SupItems.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_BillingDoc.SupItems, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If
      
      Values(1) = CR.SUP_ITEM_ID
      Values(2) = RealIndex
      Values(3) = CR.PART_NO
      If CR.PIG_FLAG = "Y" Then
         Values(4) = CR.ITEM_DESC
      Else
         Values(4) = CR.PART_DESC
      End If
      Values(5) = FormatNumber(CR.TX_AMOUNT)
      
'      If SupCode = "�-0012" Or SupCode = "�-1051" Then
'         Values(6) = FormatNumber(CR.ACTUAL_UNIT_PRICE, 3)
'      Else
'        Values(6) = FormatNumber(CR.ACTUAL_UNIT_PRICE)
'      End If
      Values(6) = FormatNumber(CR.ACTUAL_UNIT_PRICE)
      Values(7) = FormatNumber(CR.TOTAL_ACTUAL_PRICE)
      Values(8) = CR.LOCATION_NAME
      Values(9) = CR.PO_NO
      txtDoNo.Text = CR.PO_NO
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
  ' txtDeliveryFee.Text = FormatNumber(Val(txtMaterialPrice.Text) * 0.07)
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub TabStrip1_Click()
'  TabStrip1.Top = 4920
'   GridEX1.Top = 5400
'   GridEX1.Left = 150
   GridEX1.Visible = False

'   SSFrame4.Top = 5400
'   SSFrame4.Left = 150
   SSFrame4.Visible = False
   

   If TabStrip1.SelectedItem.Index = 1 Then
     Call EnableDisableButton(True)
      GridEX1.Visible = True
      Call InitGrid1
      
      Call GetTotalPrice
      GridEX1.ItemCount = CountItem(m_BillingDoc.SupItems)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
       SSFrame4.Visible = True
      Call EnableDisableButton(False)
       
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub txtCredit_Change()
   m_HasModify = True
   If DocumentType = 100 Or DocumentType = 101 Or DocumentType = 102 Or DocumentType = 103 Then
      uctlDueDate.ShowDate = DateAdd("D", Val(txtCredit.Text), uctlDocumentDate.ShowDate)
   End If
End Sub
Private Sub txtDeliveryFee_Change()
   m_HasModify = True
   txtTotal.Text = Format(Val(txtDeliveryFee.Text) + Val(txtMaterialPrice.Text), "0.00")
End Sub

Private Sub txtDesc_Change()
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

Private Sub txtDueAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtMaterialPrice_Change()
   m_HasModify = True
   txtTotal.Text = Format(Val(txtDeliveryFee.Text) + Val(txtMaterialPrice.Text), "0.00")
End Sub
Private Sub txtPrNo_Change()
   m_HasModify = True
End Sub

Private Sub txtQueNo_Change()
   m_HasModify = True
End Sub

Private Sub txtReceiver_Change()
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

Private Sub uctlDocumentDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlDueDate_HasChange()
   m_HasModify = True
End Sub
Private Sub uctlEntryTime_HasChange()
   m_HasModify = True
End Sub
Private Sub uctlExitTime_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlSupplierLookup_Change()
Dim ID As Long
Dim Sp As CSupplier
   
   
   ID = uctlSupplierLookup.MyCombo.ItemData(Minus2Zero(uctlSupplierLookup.MyCombo.ListIndex))
   If ID > 0 Then
      Set Sp = GetSupplier(m_Suppliers, Trim(str(ID)))
      
      txtCredit.Text = Sp.Credit
   End If
   
   m_HasModify = True
   
   
End Sub
Private Sub PopulateGuiID(BD As CBillingDoc)
Dim Di As CSupItem

   For Each Di In BD.SupItems
      If Di.Flag = "A" Then
         Di.LINK_ID = GetNextGuiID(BD)
      End If
   Next Di
End Sub

Private Function GetNextGuiID(BD As CBillingDoc) As Long
Dim Di As CSupItem
Dim MaxId As Long

   MaxId = 0
   For Each Di In BD.SupItems
      If Di.LINK_ID > MaxId Then
         MaxId = Di.LINK_ID
      End If
   Next Di
   
   GetNextGuiID = MaxId + 1
End Function
Private Sub EnableDisableButton(En As Boolean)
   If En Then
      If ShowMode = SHOW_EDIT Then
         cmdAdd.Enabled = (m_BillingDoc.OLD_COMMIT_FLAG = "N")
         cmdEdit.Enabled = True '(m_BillingDoc.COMMIT_FLAG = "N")
         cmdDelete.Enabled = (m_BillingDoc.OLD_COMMIT_FLAG = "N")
      Else
         cmdAdd.Enabled = True
         cmdEdit.Enabled = True
         cmdDelete.Enabled = True
      End If
   Else
      cmdAdd.Enabled = En
      cmdDelete.Enabled = En
      cmdEdit.Enabled = En
   End If
End Sub

Private Sub uctlSupplierTrueLookup_Change()
Dim ID As Long
Dim Sp As CSupplier
   
   ID = uctlSupplierTrueLookup.MyCombo.ItemData(Minus2Zero(uctlSupplierTrueLookup.MyCombo.ListIndex))
   If ID > 0 Then
      Set Sp = GetSupplier(m_Suppliers, Trim(str(ID)))
   End If
   
   m_HasModify = True
End Sub
