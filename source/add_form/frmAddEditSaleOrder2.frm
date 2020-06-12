VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAddEditSaleOrder2 
   ClientHeight    =   9855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditSaleOrder2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9855
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000009&
      Height          =   1275
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   1575
      TabIndex        =   49
      Top             =   480
      Visible         =   0   'False
      Width           =   1635
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   9840
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   17357
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboRateType 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   5040
         Width           =   5385
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   2595
         Left            =   1800
         TabIndex        =   50
         Top             =   6480
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   4577
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin VB.TextBox txtAgreementFinance 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   1440
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   52
            Top             =   1440
            Width           =   10095
         End
         Begin VB.TextBox txtAgreementData 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   1440
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   51
            Top             =   120
            Width           =   10095
         End
         Begin VB.Label lblAgreementFinance 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            TabIndex        =   54
            Top             =   1560
            Width           =   1155
         End
         Begin VB.Label lblAgreementData 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   1155
         End
      End
      Begin VB.ComboBox cboEnpAddress 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2250
         Width           =   9585
      End
      Begin VB.ComboBox cboCustomerAddress 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1800
         Width           =   9585
      End
      Begin VB.ComboBox cboAccount 
         Height          =   315
         Left            =   9120
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1350
         Width           =   2325
      End
      Begin prjFarmManagement.uctlTextLookup uctlCustomerLookup 
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
         Top             =   930
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   17
         Top             =   5940
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
         Width           =   2535
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   2280
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   2565
         Left            =   150
         TabIndex        =   18
         Top             =   6480
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   4524
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
         Column(1)       =   "frmAddEditSaleOrder2.frx":27A2
         Column(2)       =   "frmAddEditSaleOrder2.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditSaleOrder2.frx":290E
         FormatStyle(2)  =   "frmAddEditSaleOrder2.frx":2A6A
         FormatStyle(3)  =   "frmAddEditSaleOrder2.frx":2B1A
         FormatStyle(4)  =   "frmAddEditSaleOrder2.frx":2BCE
         FormatStyle(5)  =   "frmAddEditSaleOrder2.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditSaleOrder2.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   27
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
         TabIndex        =   9
         Top             =   2670
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTotalDiscount 
         Height          =   435
         Left            =   5910
         TabIndex        =   10
         Top             =   2670
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2408
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtNetTotal 
         Height          =   435
         Left            =   9210
         TabIndex        =   11
         Top             =   2670
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlSellByLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   14
         Top             =   4020
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtDiscount 
         Height          =   435
         Left            =   1860
         TabIndex        =   12
         Top             =   3120
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtIncludeDiscount 
         Height          =   435
         Left            =   9210
         TabIndex        =   39
         Top             =   3120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtDeposit 
         Height          =   435
         Left            =   1860
         TabIndex        =   44
         Top             =   3570
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtLeft 
         Height          =   435
         Left            =   9210
         TabIndex        =   13
         Top             =   3570
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlDeliveryCusLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   57
         Top             =   4560
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin Threed.SSCommand cmdAccessDeliveryCus 
         Height          =   405
         Left            =   7320
         TabIndex        =   60
         Top             =   4560
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditSaleOrder2.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEditCon 
         Height          =   405
         Left            =   7320
         TabIndex        =   59
         Top             =   5040
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditSaleOrder2.frx":3250
         ButtonStyle     =   3
      End
      Begin VB.Label lblRateType 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   56
         Top             =   5040
         Width           =   1635
      End
      Begin VB.Label lblDeliveryCusLookup 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   58
         Top             =   4560
         Width           =   1635
      End
      Begin Threed.SSCommand cmdCustomer 
         Height          =   405
         Left            =   7260
         TabIndex        =   5
         Top             =   1350
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditSaleOrder2.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAuto 
         Height          =   405
         Left            =   4410
         TabIndex        =   1
         Top             =   900
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditSaleOrder2.frx":3884
         ButtonStyle     =   3
      End
      Begin VB.Label lblDeposit 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         TabIndex        =   48
         Top             =   3690
         Width           =   1695
      End
      Begin VB.Label Label10 
         Height          =   315
         Left            =   3540
         TabIndex        =   47
         Top             =   3660
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label lblLeft 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7890
         TabIndex        =   46
         Top             =   3660
         Width           =   1275
      End
      Begin VB.Label Label8 
         Height          =   315
         Left            =   10860
         TabIndex        =   45
         Top             =   3630
         Width           =   585
      End
      Begin VB.Label lblDiscount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         TabIndex        =   43
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label6 
         Height          =   315
         Left            =   3540
         TabIndex        =   42
         Top             =   3210
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label lblIncludeDiscount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7890
         TabIndex        =   41
         Top             =   3210
         Width           =   1275
      End
      Begin VB.Label Label3 
         Height          =   315
         Left            =   10860
         TabIndex        =   40
         Top             =   3180
         Width           =   585
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
         TabIndex        =   38
         Top             =   4080
         Width           =   1635
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   525
         Left            =   8250
         TabIndex        =   15
         Top             =   4050
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditSaleOrder2.frx":3B9E
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   525
         Left            =   9870
         TabIndex        =   16
         Top             =   4050
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblEnpAddress 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   37
         Top             =   2340
         Width           =   1635
      End
      Begin VB.Label lblCustomerAddress 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   36
         Top             =   1890
         Width           =   1635
      End
      Begin VB.Label lblAccountNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7920
         TabIndex        =   35
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblCustomer 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   150
         TabIndex        =   34
         Top             =   1410
         Width           =   1635
      End
      Begin VB.Label Label4 
         Height          =   315
         Left            =   10860
         TabIndex        =   33
         Top             =   2730
         Width           =   585
      End
      Begin VB.Label lblNetTotal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8190
         TabIndex        =   32
         Top             =   2760
         Width           =   915
      End
      Begin VB.Label lblTotalDiscount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4080
         TabIndex        =   31
         Top             =   2760
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label2 
         Height          =   315
         Left            =   7350
         TabIndex        =   30
         Top             =   2730
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label Label1 
         Height          =   315
         Left            =   3540
         TabIndex        =   29
         Top             =   2760
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5040
         TabIndex        =   28
         Top             =   960
         Width           =   1395
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   22
         Top             =   9150
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditSaleOrder2.frx":3EB8
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   23
         Top             =   9150
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   20
         Top             =   9150
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   19
         Top             =   9150
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditSaleOrder2.frx":41D2
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   21
         Top             =   9150
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditSaleOrder2.frx":44EC
         ButtonStyle     =   3
      End
      Begin VB.Label lblTotalAmount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         TabIndex        =   26
         Top             =   2790
         Width           =   1695
      End
      Begin VB.Label lblDocumentNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   90
         TabIndex        =   25
         Top             =   960
         Width           =   1665
      End
   End
End
Attribute VB_Name = "frmAddEditSaleOrder2"
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
Private m_Customers As Collection
Private m_Employees As Collection
Private m_Resources As Collection
Private MonthlyAccums As Collection

Private m_DeliveryCus As Collection
Private m_ExWorkPricesItem As Collection
Private m_ExDeliveryCostItem As Collection
Private m_ExPromotionPartItem As Collection
Private m_ExPromotionDlcItem As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public id As Long
Public Area As Long
Public DocumentType As Long

Private FileName As String
Private m_SumUnit As Double
Private m_Cd As Collection
Private DocAdd As Long
Private CUSTOMER_ID As Long
Private DocumentDate As Date

Private CAL_RATE_DELIVERY_TYPE As Long
Private PRICE_THINK_TYPE As Long
Private ISuctlDeliveryCusLookup As Boolean
Private CAL_PRICE_PART_CENTER_FLAG As String
Private CAL_PRICE_DLC_CENTER_FLAG As String

Private NewUpdatePrice As Boolean

Private EditConditionFlag As Boolean
Public TempUserName As String

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
      m_BillingDoc.BILLING_DOC_ID = id
      If Not glbDaily.QueryBillingDoc(m_BillingDoc, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_BillingDoc.PopulateFromRS(1, m_Rs)
      
      uctlDocumentDate.ShowDate = m_BillingDoc.DOCUMENT_DATE
      txtDocumentNo.Text = m_BillingDoc.DOCUMENT_NO
      If Area = 1 Then
         uctlCustomerLookup.MyCombo.ListIndex = IDToListIndex(uctlCustomerLookup.MyCombo, m_BillingDoc.CUSTOMER_ID)
         cboAccount.ListIndex = IDToListIndex(cboAccount, m_BillingDoc.ACCOUNT_ID)
      ElseIf Area = 2 Then
         uctlCustomerLookup.MyCombo.ListIndex = IDToListIndex(uctlCustomerLookup.MyCombo, m_BillingDoc.SUPPLIER_ID)
         cboAccount.ListIndex = -1
      End If
      cboCustomerAddress.ListIndex = IDToListIndex(cboCustomerAddress, m_BillingDoc.BILLING_ADDRESS_ID)
      cboEnpAddress.ListIndex = IDToListIndex(cboEnpAddress, m_BillingDoc.ENTERPRISE_ADDRESS_ID)
      txtTotalAmount.Text = Format(m_BillingDoc.TOTAL_AMOUNT, "0.00")
      txtTotalDiscount.Text = Format(m_BillingDoc.DISCOUNT_AMOUNT, "0.00")
      uctlSellByLookup.MyCombo.ListIndex = IDToListIndex(uctlSellByLookup.MyCombo, m_BillingDoc.ACCEPT_BY)
      
      txtAgreementData.Text = m_BillingDoc.AGREEMENT_DATA
      txtAgreementFinance.Text = m_BillingDoc.AGREEMENT_FINANCE
      
      uctlDeliveryCusLookup.MyCombo.ListIndex = IDToListIndex(uctlDeliveryCusLookup.MyCombo, m_BillingDoc.DELIVERY_CUS_ITEM_ID)
      If cboRateType.ListIndex <> m_BillingDoc.PRICE_THINK_TYPE And m_BillingDoc.PRICE_THINK_TYPE > 0 Then 'ถ้า bill มีการแก้ไขเงื่อนไขการขาย ก็ให้แสดง เงื่อนไขที่ถูกแก้ไขของ bill นั้นๆ
         cboRateType.ListIndex = m_BillingDoc.PRICE_THINK_TYPE
      End If
      
      chkCommit.Value = FlagToCheck(m_BillingDoc.COMMIT_FLAG)
      cmdAdd.Enabled = (m_BillingDoc.COMMIT_FLAG = "N")
      cmdDelete.Enabled = (m_BillingDoc.COMMIT_FLAG = "N")
      chkCommit.Enabled = (m_BillingDoc.COMMIT_FLAG = "N")
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

Private Sub PopulateGuiID(BD As CBillingDoc)
Dim Di As CDoItem

   For Each Di In BD.DoItems
      If Di.Flag = "A" Then
         Di.LINK_ID = GetNextGuiID(BD)
      End If
   Next Di
End Sub

Private Function GetNextGuiID(BD As CBillingDoc) As Long
Dim Di As CDoItem
Dim MaxId As Long

   MaxId = 0
   For Each Di In BD.DoItems
      If Di.LINK_ID > MaxId Then
         MaxId = Di.LINK_ID
      End If
   Next Di
   
   GetNextGuiID = MaxId + 1
End Function

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim Ivd As CInventoryDoc
If (m_BillingDoc.SUCCESS_FLAG = "C") And (Not NewUpdatePrice) Then
      glbErrorLog.LocalErrorMsg = MapText("ใบ Sale Order เลขที่ ") & " " & m_BillingDoc.DOCUMENT_NO & " " & MapText("ออกใบส่งของเรียบร้อยแล้วไม่สามารถเปลี่ยนแปลงเอกสารได้")
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Function
End If
   If ShowMode = SHOW_EDIT Then
      If Area = 1 Then
         If Not VerifyAccessRight("LEDGER_SELL" & "_" & DocumentType & "_" & "EDIT", "แก้ไข") Then
            Call EnableForm(Me, True)
            Exit Function
         End If
      End If
      
      
      If m_BillingDoc.SaleOrders.Count > 0 And ISuctlDeliveryCusLookup Then
         glbErrorLog.LocalErrorMsg = MapText("มีการเปลี่ยนแปลงสถานที่จัดส่ง หรือ วิธีการคิดราคา กรุณาไปปรับปรุงราคาใหม่ในรายการอาหารอีกครั้ง")
         glbErrorLog.ShowUserError
         Exit Function
      End If
   End If
   If Not VerifyTextControl(lblDocumentNo, txtDocumentNo, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblDocumentDate, uctlDocumentDate, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblCustomer, uctlCustomerLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblTotalAmount, txtTotalAmount, True) Then
      Exit Function
   End If

   If Not CheckUniqueNs(DO_PLAN_UNIQUE, txtDocumentNo.Text, id) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtDocumentNo.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      txtDocumentNo.Text = ""
      DocAdd = DocAdd + 1
      Call cmdAuto_Click
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_BillingDoc.AddEditMode = ShowMode
   m_BillingDoc.BILLING_DOC_ID = id
    m_BillingDoc.DOCUMENT_DATE = uctlDocumentDate.ShowDate
   m_BillingDoc.DOCUMENT_NO = txtDocumentNo.Text
   If Area = 1 Then
      m_BillingDoc.CUSTOMER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
      m_BillingDoc.ACCOUNT_ID = cboAccount.ItemData(Minus2Zero(cboAccount.ListIndex))
      m_BillingDoc.DOCUMENT_TYPE = 19 'ใบ SO
   End If
   m_BillingDoc.BILLING_ADDRESS_ID = cboCustomerAddress.ItemData(Minus2Zero(cboCustomerAddress.ListIndex))
   m_BillingDoc.ENTERPRISE_ADDRESS_ID = cboEnpAddress.ItemData(Minus2Zero(cboEnpAddress.ListIndex))
   m_BillingDoc.EXCEPTION_FLAG = "N"
   m_BillingDoc.ACCEPT_BY = uctlSellByLookup.MyCombo.ItemData(Minus2Zero(uctlSellByLookup.MyCombo.ListIndex))
   m_BillingDoc.COMMIT_FLAG = Check2Flag(chkCommit.Value)
   m_BillingDoc.DISCOUNT_AMOUNT = Val(txtDiscount.Text)
   m_BillingDoc.DEPOSIT_AMOUNT = Val(txtDeposit.Text)
   m_BillingDoc.TOTAL_AMOUNT = Val(txtTotalAmount.Text)
   m_BillingDoc.TOTAL_PRICE = Val(txtNetTotal.Text)
   
   m_BillingDoc.AGREEMENT_DATA = txtAgreementData.Text
   m_BillingDoc.AGREEMENT_FINANCE = txtAgreementFinance.Text
   
   m_BillingDoc.DELIVERY_CUS_ITEM_ID = uctlDeliveryCusLookup.MyCombo.ItemData(Minus2Zero(uctlDeliveryCusLookup.MyCombo.ListIndex))
   m_BillingDoc.PRICE_THINK_TYPE = cboRateType.ListIndex
   
   If EditConditionFlag Then
      m_BillingDoc.USER_APPLOVE_PRICE_THINK = TempUserName
   End If


   
   Call PopulateGuiID(m_BillingDoc)
   
   Call EnableForm(Me, False)
   
   'ไม่ต้องทำการสร้าง InventoryDoc
'   Call glbDaily.DO2InventoryDoc(m_BillingDoc, Ivd, Area)
   
   If (m_BillingDoc.COMMIT_FLAG = "Y") Then
      If m_BillingDoc.OLD_COMMIT_FLAG <> "Y" Then
'         Call glbDaily.TriggerCommit(Ivd.ImportExports)
'         If Not glbDaily.VerifyStockBalance(Ivd.ImportExports, glbErrorLog) Then
'            Call EnableForm(Me, True)
'            Exit Function
'         End If
         
      End If
   End If
   
   Call glbDaily.StartTransaction

'หากมีการแก้ไขราคาหลังจากออกใบส่งของแล้วจะมีการ update Flag ตรงนี้เพื่อนำสถานะไปใช้งานต่อ
   Dim t_BillingDoc As CBillingDoc
   Set t_BillingDoc = New CBillingDoc
    If NewUpdatePrice And m_BillingDoc.SUCCESS_FLAG = "C" Then
         m_BillingDoc.EDIT_PRICE_FLAG = "Y"
         m_BillingDoc.SUCCESS_FLAG = "Y" 'บอกว่าเอกสารรออกใบส่งของ
      End If
      
   If m_BillingDoc.AddEditMode = SHOW_EDIT And Not NewUpdatePrice Then
     m_BillingDoc.SUCCESS_FLAG = "N"
   End If
   
   
   If Not glbDaily.AddEditBillingDoc(m_BillingDoc, IsOK, False, glbErrorLog) Then
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

Private Sub cboRateType_Change()
   m_HasModify = True
End Sub

Private Sub cboRateType_Click()
   m_HasModify = True
   PRICE_THINK_TYPE = cboRateType.ListIndex
End Sub

Private Sub chkCommit_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkCommit_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Public Sub RefreshGrid()
   Call GetTotalPrice

   GridEX1.ItemCount = CountItem(m_BillingDoc.DoItems)
   GridEX1.Rebind
End Sub

Private Sub cmdAccessDeliveryCus_Click()
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu
   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("ข้อมูลค่าขนส่ง", "-", "ข้อมูลส่วนลดค่าขนส่ง", "-", "ข้อมูลสถานที่จัดส่ง")
   If lMenuChosen = 0 Then
      Exit Sub
   End If

If lMenuChosen = 1 Then
     If Not VerifyAccessRight("PACKAGE-CENTER_DELIVERY-COST") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      frmExWorksPrice.Area = 2
      Load frmExWorksPrice
      frmExWorksPrice.Show 1
      Unload frmExWorksPrice
      Set frmExWorksPrice = Nothing
ElseIf lMenuChosen = 3 Then
     If Not VerifyAccessRight("PACKAGE-CENTER_PROMOTION-DELIVERY") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      frmExWorksPrice.Area = 4
      Load frmExWorksPrice
      frmExWorksPrice.Show 1
      Unload frmExWorksPrice
      Set frmExWorksPrice = Nothing
ElseIf lMenuChosen = 5 Then
      frmAddEditDeliveryCusMain.HeaderText = MapText("ข้อมูลสถานที่จัดส่ง")
      frmAddEditDeliveryCusMain.CustomerID = CUSTOMER_ID
      Load frmAddEditDeliveryCusMain
      frmAddEditDeliveryCusMain.Show 1

      OKClick = frmAddEditDeliveryCusMain.OKClick

      Unload frmAddEditDeliveryCusMain
      Set frmAddEditDeliveryCusMain = Nothing
End If

 If CUSTOMER_ID > 0 Then
   Call LoadDeliveryCus(uctlDeliveryCusLookup.MyCombo, m_DeliveryCus, CUSTOMER_ID, , , , "N") 'LOAD สถานที่จัดส่ง
   Set uctlDeliveryCusLookup.MyCollection = m_DeliveryCus
End If
Call LoadExDeliveryCusItem(Nothing, m_ExDeliveryCostItem, , 3, uctlDocumentDate.ShowDate) 'ส่ง 3 ไปดึงค่าขนส่งที่คิดให้รถรับจ้าง

End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu

   If Not cmdAdd.Enabled Then
      Exit Sub
   End If
    
   If Area = 1 Then
      If Not VerifyCombo(lblAccountNo, cboAccount) Then
         Exit Sub
      End If
      If Not VerifyDate(lblDocumentDate, uctlDocumentDate) Then
         Exit Sub
      End If
   End If
   
   If Not VerifyCombo(lblAccountNo, cboAccount) Then
      Exit Sub
   End If
    
   OKClick = False
   If TabStrip1.SelectedItem.Index = 1 Then
      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.Popup("เพิ่มรายการใหม่ (ขายอาหาร)", "-", "เพิ่มรายการใหม่ (ขายอื่นๆ)")
      If lMenuChosen = 0 Then
         Exit Sub
      End If
      Set oMenu = Nothing
      
      If lMenuChosen = 1 Then
         If PRICE_THINK_TYPE <> 1 Then
             If Not VerifyCombo(lblDeliveryCusLookup, uctlDeliveryCusLookup.MyCombo, False) Then
               Exit Sub
            End If
         End If
         
         If Not VerifyCombo(lblRateType, cboRateType, False) Then
            Exit Sub
         End If
         
         If Area = 1 Then
            frmAddEditSaleOrderItem2.AccountID = cboAccount.ItemData(cboAccount.ListIndex)
         End If
         
         Call LoadExWorksPriceItem(Nothing, m_ExWorkPricesItem, , 2, uctlDocumentDate.ShowDate, , "Y")
         Call LoadExDeliveryCusItem(Nothing, m_ExDeliveryCostItem, , 2, uctlDocumentDate.ShowDate)
         Call LoadExPromotionPartItem(Nothing, m_ExPromotionPartItem, , 2, uctlDocumentDate.ShowDate, , "Y")
         Call LoadExPromotionDlcItem(Nothing, m_ExPromotionDlcItem, , 2, uctlDocumentDate.ShowDate)
         
         frmAddEditSaleOrderItem2.DocumentType = DocumentType
         frmAddEditSaleOrderItem2.DocumentDate = uctlDocumentDate.ShowDate
         frmAddEditSaleOrderItem2.Area = Area
         frmAddEditSaleOrderItem2.COMMIT_FLAG = m_BillingDoc.COMMIT_FLAG
         Set frmAddEditSaleOrderItem2.TempCollection = m_BillingDoc.SaleOrders

          Set frmAddEditSaleOrderItem2.m_Customers = m_Customers
         Set frmAddEditSaleOrderItem2.m_ExWorkPricesItem = m_ExWorkPricesItem
         Set frmAddEditSaleOrderItem2.m_ExDeliveryCostItem = m_ExDeliveryCostItem
         Set frmAddEditSaleOrderItem2.m_ExPromotionPartItem = m_ExPromotionPartItem
         Set frmAddEditSaleOrderItem2.m_ExPromotionDlcItem = m_ExPromotionDlcItem
         Set frmAddEditSaleOrderItem2.m_DeliveryCus = m_DeliveryCus
         frmAddEditSaleOrderItem2.CUSTOMER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
         frmAddEditSaleOrderItem2.DELIVERY_CUS_ITEM_ID = uctlDeliveryCusLookup.MyCombo.ItemData(Minus2Zero(uctlDeliveryCusLookup.MyCombo.ListIndex))
         frmAddEditSaleOrderItem2.PRICE_THINK_TYPE = PRICE_THINK_TYPE ' cboRateType.ListIndex
         frmAddEditSaleOrderItem2.CAL_RATE_DELIVERY_TYPE = CAL_RATE_DELIVERY_TYPE
         frmAddEditSaleOrderItem2.CAL_PRICE_PART_CENTER_FLAG = CAL_PRICE_PART_CENTER_FLAG
         frmAddEditSaleOrderItem2.CAL_PRICE_DLC_CENTER_FLAG = CAL_PRICE_DLC_CENTER_FLAG
         'CAL_PRICE_PART_CENTER_FLAG
         frmAddEditSaleOrderItem2.TypeSale = 1 'ขายสินค้าและค่าขนส่ง
         
         frmAddEditSaleOrderItem2.ParentShowMode = ShowMode
         frmAddEditSaleOrderItem2.ShowMode = SHOW_ADD
         frmAddEditSaleOrderItem2.HeaderText = MapText("เพิ่มรายการใบ SO สินค้า")
         Load frmAddEditSaleOrderItem2
         frmAddEditSaleOrderItem2.Show 1
   
         OKClick = frmAddEditSaleOrderItem2.OKClick
   
         Unload frmAddEditSaleOrderItem2
         Set frmAddEditSaleOrderItem2 = Nothing
   
         If OKClick Then
            Call GetTotalPrice
   
            GridEX1.ItemCount = CountItem(m_BillingDoc.SaleOrders)
            GridEX1.Rebind
         End If
      ElseIf lMenuChosen = 3 Then
         If Area = 1 Then
            frmAddEditSaleOrderItem2.AccountID = cboAccount.ItemData(cboAccount.ListIndex)
         End If
         frmAddEditSaleOrderItem2.DocumentType = DocumentType
         frmAddEditSaleOrderItem2.DocumentDate = uctlDocumentDate.ShowDate
         frmAddEditSaleOrderItem2.Area = Area
         frmAddEditSaleOrderItem2.COMMIT_FLAG = m_BillingDoc.COMMIT_FLAG
         Set frmAddEditSaleOrderItem2.TempCollection = m_BillingDoc.SaleOrders
         
          Set frmAddEditSaleOrderItem2.m_Customers = m_Customers
         Set frmAddEditSaleOrderItem2.m_ExWorkPricesItem = m_ExWorkPricesItem
         Set frmAddEditSaleOrderItem2.m_ExDeliveryCostItem = m_ExDeliveryCostItem
         Set frmAddEditSaleOrderItem2.m_DeliveryCus = m_DeliveryCus
         frmAddEditSaleOrderItem2.CUSTOMER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
         frmAddEditSaleOrderItem2.DELIVERY_CUS_ITEM_ID = uctlDeliveryCusLookup.MyCombo.ItemData(Minus2Zero(uctlDeliveryCusLookup.MyCombo.ListIndex))
         frmAddEditSaleOrderItem2.PRICE_THINK_TYPE = PRICE_THINK_TYPE 'cboRateType.ListIndex
         frmAddEditSaleOrderItem2.CAL_RATE_DELIVERY_TYPE = CAL_RATE_DELIVERY_TYPE
         frmAddEditSaleOrderItem2.TypeSale = 2 'ขายอื่นๆ
         
         frmAddEditSaleOrderItem2.ParentShowMode = ShowMode
         frmAddEditSaleOrderItem2.ShowMode = SHOW_ADD
         frmAddEditSaleOrderItem2.HeaderText = MapText("เพิ่มรายการใบ SO การขายอื่นๆ")
         Load frmAddEditSaleOrderItem2
         frmAddEditSaleOrderItem2.Show 1
   
         OKClick = frmAddEditSaleOrderItem2.OKClick
   
         Unload frmAddEditSaleOrderItem2
         Set frmAddEditSaleOrderItem2 = Nothing
   
         If OKClick Then
            Call GetTotalPrice
   
            GridEX1.ItemCount = CountItem(m_BillingDoc.SaleOrders)
            GridEX1.Rebind
         End If
      Else
         frmAddEditDoItemEx.AccountID = cboAccount.ItemData(cboAccount.ListIndex)
         Set frmAddEditDoItemEx.ParentForm = Me
         frmAddEditDoItemEx.SubscriberID = -1
         frmAddEditDoItemEx.Area = Area
         frmAddEditDoItemEx.COMMIT_FLAG = m_BillingDoc.COMMIT_FLAG
         Set frmAddEditDoItemEx.TempCollection = m_BillingDoc.DoItems
         frmAddEditDoItemEx.ParentShowMode = ShowMode
         frmAddEditDoItemEx.ShowMode = SHOW_ADD
         frmAddEditDoItemEx.HeaderText = MapText("เพิ่มรายการใบ SO")
         Load frmAddEditDoItemEx
         frmAddEditDoItemEx.Show 1

         OKClick = frmAddEditDoItemEx.OKClick

         Unload frmAddEditDoItemEx
         Set frmAddEditDoItemEx = Nothing

         If OKClick Then
            Call GetTotalPrice
   
            GridEX1.ItemCount = CountItem(m_BillingDoc.DoItems)
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

Private Function GetDocumentNo(DocNoType As Long) As String
Dim No As String
Dim DOC_ID As Long
Dim Cd As CConfigDoc
Dim TempStr As String
Dim I As Long
Dim ServerDateTime As String

   DOC_ID = SELL_SO
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
'         GetDocumentNo = GetDocumentNo & Format(Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd, TempStr)
'         m_BillingDoc.RUNNING_NO = Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd
'         m_BillingDoc.CONFIG_DOC_TYPE = SELL_SO
         
         If Cd.GetFieldValue("AUTO_BEGIN_FLAG") = "Y" Then
               
               If CheckNewMounth And CheckUniqueNs(DO_PLAN_UNIQUE, GetDocumentNo & Format(1, TempStr), id) Then
                  GetDocumentNo = GetDocumentNo & Format(1, TempStr) 'เริ่มจาก 1 เสมอ
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
Private Sub cmdAuto_Click()
If Trim(txtDocumentNo.Text) = "" And ShowMode = SHOW_ADD Then
   txtDocumentNo.Text = GetDocumentNo(DocumentType)
End If
'Dim ID As Long
'Dim Cd As CConfigDoc
'Dim TempStr As String
'Dim I As Long
'Dim ServerDateTime As String
'
'   ID = SELL_SO
'   If ID > 0 Then
'      Set Cd = GetObject("CConfigDoc", m_Cd, Trim(str(ID)), False)
'      If Not (Cd Is Nothing) Then
'         txtDocumentNo.Text = Cd.GetFieldValue("PREFIX") & Cd.GetFieldValue("CODE1")
'         TempStr = ""
'         If Cd.GetFieldValue("YEAR_TYPE") = 1 Then
'            TempStr = Right(Format(Year(Now) + 543, "0000"), 2)
'         ElseIf Cd.GetFieldValue("YEAR_TYPE") = 2 Then
'            TempStr = Format(Year(Now) + 543, "0000")
'         ElseIf Cd.GetFieldValue("YEAR_TYPE") = 3 Then
'            TempStr = Right(Format(Year(Now), "0000"), 2)
'         ElseIf Cd.GetFieldValue("YEAR_TYPE") = 4 Then
'            TempStr = Format(Year(Now), "0000")
'         End If
'         txtDocumentNo.Text = txtDocumentNo.Text & TempStr & Cd.GetFieldValue("CODE2")
'         TempStr = ""
'         If Cd.GetFieldValue("MONTH_TYPE") = 1 Then
'            TempStr = Format(Month(Now), "00")
'         End If
'         txtDocumentNo.Text = txtDocumentNo.Text & TempStr & Cd.GetFieldValue("CODE3")
'         TempStr = ""
'         For I = 1 To Cd.GetFieldValue("DIGIT_AMOUNT")
'            TempStr = TempStr & "0"
'         Next I
'         txtDocumentNo.Text = txtDocumentNo.Text & Format(Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd, TempStr)
'         m_BillingDoc.RUNNING_NO = Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd
'         m_BillingDoc.CONFIG_DOC_TYPE = SELL_SO
'      Else
'         txtDocumentNo.Text = ""
'      End If
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
   
   If PRICE_THINK_TYPE = 3 Then
     If GridEX1.Value(8) = -1 Then
            glbErrorLog.LocalErrorMsg = MapText("ไม่สามารถลบค่าขนส่งได้ เนื่องจากลูกค้ารายนี้มีการคิดแบบแยกค่าขนส่ง")
            glbErrorLog.ShowUserError
            Exit Sub
      End If
   End If
   
   If TabStrip1.SelectedItem.Index = 1 Then
      If ID1 <= 0 Then
         m_BillingDoc.SaleOrders.Remove (ID2)
      Else
         m_BillingDoc.SaleOrders.Item(ID2).Flag = "D"
      End If

      Call GetTotalPrice
      GridEX1.ItemCount = CountItem(m_BillingDoc.SaleOrders)
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
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   If Area = 1 Then
      If Not VerifyCombo(lblAccountNo, cboAccount) Then
         Exit Sub
      End If
      If Not VerifyDate(lblDocumentDate, uctlDocumentDate) Then
         Exit Sub
      End If
   End If
   
   id = Val(GridEX1.Value(2))
   OKClick = False
   
   If TabStrip1.SelectedItem.Index = 1 Then
   
      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.Popup("แก้ไขรายการ (ขายอาหาร)", "-", "แก้ไขรายการ (ขายอื่นๆ)")
      If lMenuChosen = 0 Then
         Exit Sub
      End If
      Set oMenu = Nothing
   
         If Area = 1 Then
            frmAddEditSaleOrderItem2.AccountID = cboAccount.ItemData(cboAccount.ListIndex)
         End If
         

         If lMenuChosen = 1 Then
            Call LoadExWorksPriceItem(Nothing, m_ExWorkPricesItem, , 2, uctlDocumentDate.ShowDate, , "Y")
            Call LoadExDeliveryCusItem(Nothing, m_ExDeliveryCostItem, , 2, uctlDocumentDate.ShowDate)
            Call LoadExPromotionPartItem(Nothing, m_ExPromotionPartItem, , 2, uctlDocumentDate.ShowDate, , "Y")
            Call LoadExPromotionDlcItem(Nothing, m_ExPromotionDlcItem, , 2, uctlDocumentDate.ShowDate)
            
             frmAddEditSaleOrderItem2.TypeSale = 1 'แก้ไขขายสินค้าและค่าขนส่ง
             frmAddEditSaleOrderItem2.HeaderText = MapText("แก้ไขรายการใบ SO ขายอาหาร")
         ElseIf lMenuChosen = 3 Then
            frmAddEditSaleOrderItem2.TypeSale = 2 'แก้ไขขายทั่วไป
            frmAddEditSaleOrderItem2.HeaderText = MapText("แก้ไขรายการใบ SO ขายอื่นๆ")
         End If
            
         frmAddEditSaleOrderItem2.DocumentType = DocumentType
         frmAddEditSaleOrderItem2.DocumentDate = uctlDocumentDate.ShowDate
         frmAddEditSaleOrderItem2.Area = Area
         frmAddEditSaleOrderItem2.id = id
         frmAddEditSaleOrderItem2.COMMIT_FLAG = m_BillingDoc.COMMIT_FLAG
         Set frmAddEditSaleOrderItem2.TempCollection = m_BillingDoc.SaleOrders
         
         Set frmAddEditSaleOrderItem2.m_Customers = m_Customers
         Set frmAddEditSaleOrderItem2.m_ExWorkPricesItem = m_ExWorkPricesItem
         Set frmAddEditSaleOrderItem2.m_ExDeliveryCostItem = m_ExDeliveryCostItem
         Set frmAddEditSaleOrderItem2.m_ExPromotionPartItem = m_ExPromotionPartItem
         Set frmAddEditSaleOrderItem2.m_ExPromotionDlcItem = m_ExPromotionDlcItem
         Set frmAddEditSaleOrderItem2.m_DeliveryCus = m_DeliveryCus
         frmAddEditSaleOrderItem2.CUSTOMER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
         frmAddEditSaleOrderItem2.DELIVERY_CUS_ITEM_ID = uctlDeliveryCusLookup.MyCombo.ItemData(Minus2Zero(uctlDeliveryCusLookup.MyCombo.ListIndex))
         frmAddEditSaleOrderItem2.PRICE_THINK_TYPE = PRICE_THINK_TYPE 'cboRateType.ListIndex
         frmAddEditSaleOrderItem2.CAL_RATE_DELIVERY_TYPE = CAL_RATE_DELIVERY_TYPE
         frmAddEditSaleOrderItem2.CAL_PRICE_PART_CENTER_FLAG = CAL_PRICE_PART_CENTER_FLAG
         frmAddEditSaleOrderItem2.CAL_PRICE_DLC_CENTER_FLAG = CAL_PRICE_DLC_CENTER_FLAG
         frmAddEditSaleOrderItem2.ISuctlDeliveryCusLookup = ISuctlDeliveryCusLookup
         frmAddEditSaleOrderItem2.SuccessFlag = m_BillingDoc.SUCCESS_FLAG

         frmAddEditSaleOrderItem2.ParentShowMode = ShowMode
         frmAddEditSaleOrderItem2.ShowMode = SHOW_EDIT
         Load frmAddEditSaleOrderItem2
         frmAddEditSaleOrderItem2.Show 1
   
         OKClick = frmAddEditSaleOrderItem2.OKClick
         NewUpdatePrice = frmAddEditSaleOrderItem2.NewUpdatePrice
         
         ISuctlDeliveryCusLookup = frmAddEditSaleOrderItem2.ISuctlDeliveryCusLookup
   
         Unload frmAddEditSaleOrderItem2
         Set frmAddEditSaleOrderItem2 = Nothing
   
         If OKClick Then
            Call GetTotalPrice
            GridEX1.ItemCount = CountItem(m_BillingDoc.SaleOrders)
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

Private Sub cmdEditCon_Click()
''If m_BillingDoc.SaleOrders.Count > 0 Then 'ห้ามใช้ CountItem เนื่องจาก จะเช็คยอดรายการที่มาจาก database จริงๆ กันการหลอกลบรายการ แล้ว แก้ไขเงื่อนไขได้
''      glbErrorLog.LocalErrorMsg = MapText("ไม่สามารถเปลี่ยนแปลงเงื่อนไขการขายได้ เนื่องจากมีรายการใบส่งสินค้าแล้ว")
''      glbErrorLog.ShowUserError
''      Exit Sub
''   End If
   EditConditionFlag = False
   ISuctlDeliveryCusLookup = True
   frmVerifyAccRight.AccName = "CREDIT_PROMOTIONAL"
   frmVerifyAccRight.AccDesc = "สามารถเปลี่ยนแปลงเงื่อนไขส่งเสริมการขาย"
   Load frmVerifyAccRight
   frmVerifyAccRight.Show 1

   If frmVerifyAccRight.GrantRight Then
      TempUserName = frmVerifyAccRight.UserName
      Unload frmVerifyAccRight
      Set frmVerifyAccRight = Nothing
   Else
      Unload frmVerifyAccRight
      Set frmVerifyAccRight = Nothing
      Exit Sub
   End If
   EditConditionFlag = True
   cboRateType.Enabled = True
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
   
   ReportMode = 1
   
   If m_HasModify Or (m_BillingDoc.BILLING_DOC_ID <= 0) Then
      glbErrorLog.LocalErrorMsg = MapText("กรุณาทำการบันทึกข้อมูลให้เรียบร้อยก่อน")
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   ReportFlag = False

   Call LoadPictureFromFile(glbParameterObj.POPicture1, Picture2)
   
   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.AddMenu(glbGuiConfigs.SOPrintMenuItems)
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
'   If Not VerifyOnwerVersionMenu(lMenuChosen, glbParameterObj.Programowner) Then
'      Exit Sub
'   End If
   
   If lMenuChosen = 1 Then
      ReportKey = "CReportNormalPO"
      
      Set Report = New CReportNormalPO
      ReportFlag = True
   ElseIf lMenuChosen = 2 Then
      ReportKey = "CReportNormalPO"
      
      Set Rc = New CReportConfig
      Rc.REPORT_KEY = ReportKey
      Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
      Call Rc.QueryData(m_Rs, iCount)
      HeaderText = MapText("ใบ Sale Order (SO)")
      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
   ElseIf lMenuChosen = 16 Then
      ReportKey = "CReportNormalSO2"
      
      Set Report = New CReportNormalSO2
      ReportFlag = True
   ElseIf lMenuChosen = 17 Then
      ReportKey = "CReportNormalSO2"
      
      Set Rc = New CReportConfig
      Rc.REPORT_KEY = ReportKey
      Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
      Call Rc.QueryData(m_Rs, iCount)
      HeaderText = MapText("ใบขออนุมัติขายเกินวงเงิน")
      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
   ElseIf lMenuChosen = 19 Then
      ReportKey = "CReportNormalSO1"
      
      Set Report = New CReportNormalSO1
      ReportFlag = True
   ElseIf lMenuChosen = 20 Then
      ReportKey = "CReportNormalSO1"
      
      Set Rc = New CReportConfig
      Rc.REPORT_KEY = ReportKey
      Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
      Call Rc.QueryData(m_Rs, iCount)
      HeaderText = MapText("แบบฟอร์มใบ SO")
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
      Call Report.AddParam(MapText("ใบ Sale Order (SO)"), "REPORT_HEADER")
      Call Report.AddParam(Picture2.Picture, "BACK_GROUND")
      Call Report.AddParam(uctlSellByLookup.MyCombo.Text, "RECEIVE_NAME")
      Call Report.AddParam("", "ACCEPT_NAME")
   ElseIf lMenuChosen = 5 Then
      ReportKey = "CReportFormPO001"
      
      Set Rc = New CReportConfig
      Rc.REPORT_KEY = ReportKey
      Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
      Call Rc.QueryData(m_Rs, iCount)
      HeaderText = MapText("ใบ Sale Order (PO)")
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
      frmReportConfig.id = Rc.REPORT_CONFIG_ID
      frmReportConfig.ReportKey = ReportKey
      frmReportConfig.HeaderText = HeaderText
      Load frmReportConfig
      frmReportConfig.Show 1
      
      Unload frmReportConfig
      Set frmReportConfig = Nothing
   End If
   Call EnableForm(Me, True)
End Sub

Private Function VerifyOnwerVersionMenu(Menu As Long, Owner As String) As Boolean
   VerifyOnwerVersionMenu = True
   
   If (Menu <> 1) And (Menu <> 2) Then
      If Not glbGuiConfigs.VerifyGuiConfig("SELL_PO_PREFORM_PRINT", True) Then
         VerifyOnwerVersionMenu = False
         Exit Function
      End If
   End If
End Function

Private Sub cmdSave_Click()
Dim Result As Boolean
   If Not SaveData Then
      Exit Sub
   End If
   
   ShowMode = SHOW_EDIT
   id = m_BillingDoc.BILLING_DOC_ID
   m_BillingDoc.QueryFlag = 1
   QueryData (True)
   m_HasModify = False
End Sub

Private Sub Form_Activate()
Dim YYYYMM As String
Dim firstDate As Date
Dim lastDate As Date
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
'      DoEvents

      
      Call EnableForm(Me, False)
      Call LoadEnterpriseAddress(cboEnpAddress, , , True)
      Call LoadConfigDoc(Nothing, m_Cd)

      Call InitDoRateType2(cboRateType)
      
      If Area = 1 Then
         Call LoadCustomer(uctlCustomerLookup.MyCombo, m_Customers)
         Set uctlCustomerLookup.MyCollection = m_Customers
      ElseIf Area = 2 Then
         Call LoadSupplier(uctlCustomerLookup.MyCombo, m_Customers)
         Set uctlCustomerLookup.MyCollection = m_Customers
      End If
      
      Call LoadEmployee(uctlSellByLookup.MyCombo, m_Employees)
      Set uctlSellByLookup.MyCollection = m_Employees
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_BillingDoc.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         uctlDocumentDate.ShowDate = Now
         m_BillingDoc.QueryFlag = 0
         Call QueryData(False)
      End If
      DocAdd = 0
      
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
   
   Set m_BillingDoc = Nothing
   Set m_Customers = Nothing
   Set m_Employees = Nothing
   Set m_Resources = Nothing
   Set MonthlyAccums = Nothing
   Set m_Cd = Nothing
   
   Set m_DeliveryCus = Nothing
   Set m_ExWorkPricesItem = Nothing
   Set m_ExDeliveryCostItem = Nothing
   Set m_ExPromotionPartItem = Nothing
   Set m_ExPromotionDlcItem = Nothing
End Sub

Private Sub GridEX1_Click()
'  Call cmdEdit_Click
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
   Col.Width = 2235
   Col.Caption = MapText("รหัสสินค้า")

   Set Col = GridEX1.Columns.add '3
   Col.Width = 3020
   Col.Caption = MapText("รายละเอียด")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 1620
   Col.Caption = MapText("จำนวน")
   
   Set Col = GridEX1.Columns.add '5
   Col.TextAlignment = jgexAlignRight
   Col.Width = 1575
   Col.Caption = MapText("ราคารวม")
   
   Set Col = GridEX1.Columns.add '6
   Col.TextAlignment = jgexAlignRight
   Col.Width = 1300
   Col.Caption = MapText("ราคา/หน่วย")
   
   Set Col = GridEX1.Columns.add '6
   Col.Width = 1800
   Col.Caption = MapText("ผู้แก้ไขราคาขาย")
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = 0
   Col.Caption = MapText("Part_Item_id")
End Sub

Private Sub GetTotalPrice()
Dim II As CSaleOrder
Dim Sum1 As Double
Dim Sum2 As Double
Dim Sum3 As Double
Dim Sum4 As Double
Dim Sum5 As Double

   Sum1 = 0
   Sum2 = 0
   Sum3 = 0
   Sum4 = 0
   Sum5 = 0
   For Each II In m_BillingDoc.SaleOrders
      If II.Flag <> "D" Then
         Sum1 = Sum1 + II.ITEM_AMOUNT
         Sum2 = Sum2 + (II.TOTAL_PRICE + II.DISCOUNT_AMOUNT)
         Sum3 = Sum3 + II.TOTAL_WEIGHT
         Sum4 = Sum4 + II.DISCOUNT_AMOUNT
         Sum5 = Sum5 + II.DEPOSIT_AMOUNT
      End If
   Next II

   txtNetTotal.Text = Format(Sum2, "0.00")
   txtTotalDiscount.Text = Format(Sum3, "0.00")
   txtTotalAmount.Text = Format(Sum1, "0.00")
   txtDiscount.Text = Format(Sum4, "0.00")
   txtDeposit.Text = Format(Sum5, "0.00")
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame2.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblDocumentNo, MapText("เลขที่ใบ SO"))
   Call InitNormalLabel(lblAccountNo, MapText("เลขที่บัญชี"))
   If Area = 1 Then
      Call InitNormalLabel(lblCustomerAddress, MapText("ที่อยู่ลูกค้า"))
      Call InitNormalLabel(lblCustomer, MapText("รหัสลูกค้า"))
      Call InitNormalLabel(lblSellBy, MapText("พนักงานขาย"))
      Call InitNormalLabel(lblEnpAddress, MapText("ที่อยู่ออกเอกสาร"))
   ElseIf Area = 2 Then
      Call InitNormalLabel(lblCustomerAddress, MapText("ที่อยู่ซัพ ฯ"))
      Call InitNormalLabel(lblCustomer, MapText("รหัสซัพ ฯ"))
      Call InitNormalLabel(lblSellBy, MapText("ผู้รับของ"))
      Call InitNormalLabel(lblEnpAddress, MapText("ที่อยู่รับเอกสาร"))
      
      lblAccountNo.Visible = False
      cboAccount.Visible = False
      cmdAuto.Visible = False
      cmdPrint.Enabled = False
   End If
   Call InitNormalLabel(lblTotalAmount, MapText("จำนวนรวม"))
   Call InitNormalLabel(lblDocumentDate, MapText("วันที่เอกสาร"))
   Call InitNormalLabel(lblTotalDiscount, MapText("น้ำหนักรวม"))
   Call InitNormalLabel(Label1, MapText("ตัว"))
   Call InitNormalLabel(Label2, MapText("ก.ก."))
   Call InitNormalLabel(Label4, MapText("บาท"))
   Call InitNormalLabel(lblNetTotal, MapText("ราคารวม"))
   Call InitNormalLabel(lblDiscount, MapText("ส่วนลด"))
   Call InitNormalLabel(lblDeposit, MapText("มัดจำ"))
   Call InitNormalLabel(lblIncludeDiscount, MapText("รวมส่วนลด"))
   Call InitNormalLabel(lblLeft, MapText("คงค้าง"))
   Call InitCheckBox(chkCommit, "คำนวณ")
   Call InitNormalLabel(Label3, MapText("บาท"))
   Call InitNormalLabel(Label6, MapText("บาท"))
   Call InitNormalLabel(Label8, MapText("บาท"))
   Call InitNormalLabel(Label10, MapText("บาท"))
   
   Call InitNormalLabel(lblAgreementData, MapText("ฝ่ายข้อมูล"))
   Call InitNormalLabel(lblAgreementFinance, MapText("ฝ่ายสินเชื่อ"))
   
   Call InitNormalLabel(lblDeliveryCusLookup, MapText("สถานที่จัดส่ง"))
   Call InitNormalLabel(lblRateType, MapText("คิดราคาแบบ"))
   
   Call txtDocumentNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   txtDocumentNo.Enabled = False
   Call txtTotalAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtTotalAmount.Enabled = False
   Call txtTotalDiscount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtTotalDiscount.Enabled = False
   Call txtNetTotal.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtNetTotal.Enabled = False
   Call txtDeposit.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtDeposit.Enabled = False
   Call txtIncludeDiscount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtIncludeDiscount.Enabled = False
   Call txtDiscount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtDiscount.Enabled = False
   Call txtLeft.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtLeft.Enabled = False
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   GridEX1.Visible = True
   SSFrame2.Visible = False
   
   Call InitCombo(cboAccount)
   Call InitCombo(cboCustomerAddress)
   Call InitCombo(cboEnpAddress)
   
   Call InitCombo(cboRateType)
   
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
   cmdAuto.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdCustomer.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEditCon.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAccessDeliveryCus.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSave, MapText("บันทึก"))
   Call InitMainButton(cmdPrint, MapText("พิมพ์"))
   Call InitMainButton(cmdAuto, MapText("A"))
   Call InitMainButton(cmdCustomer, MapText("F"))
   Call InitMainButton(cmdEditCon, MapText("แก้ไขเงื่อนไข"))
   Call InitMainButton(cmdAccessDeliveryCus, MapText("A"))
   
   Call InitGrid1
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.NAME = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.add().Caption = MapText("รายการใบ SO")
   TabStrip1.Tabs.add().Caption = MapText("ความเห็นขออนุมัติ")
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
   Set m_Customers = New Collection
   Set m_Employees = New Collection
   Set m_Resources = New Collection
   Set MonthlyAccums = New Collection
   Set m_Cd = New Collection
   Set m_DeliveryCus = New Collection
   Set m_ExWorkPricesItem = New Collection
   Set m_ExDeliveryCostItem = New Collection
   Set m_ExPromotionPartItem = New Collection
   Set m_ExPromotionDlcItem = New Collection
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub GridEX1_DblClick()
  Call cmdEdit_Click
End Sub

'Private Sub GridEX1_DblClick()
'   Call cmdEdit_Click
'End Sub

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
      If m_BillingDoc.SaleOrders Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim CR As CSaleOrder
      If m_BillingDoc.SaleOrders.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_BillingDoc.SaleOrders, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If

      Values(1) = CR.SALE_ORDER_ID
      Values(2) = RealIndex
      If CR.PART_ITEM_ID = -1 And Mid(CR.CONFIG_CODE, 1, 1) = "Y" Then
         Values(3) = CR.FEATURE_CODE
         Values(4) = CR.FEATURE_DESC
      ElseIf CR.PART_ITEM_ID > 0 Then
         Values(3) = CR.PART_NO
         Values(4) = CR.ShowDescText
      ElseIf Mid(CR.CONFIG_CODE, 3, 1) = "Y" Then
         Values(3) = ""
         Values(4) = CR.ITEM_DESC
      End If
      Values(5) = FormatNumber(CR.ITEM_AMOUNT)
      Values(6) = FormatNumber(CR.TOTAL_PRICE)
      Values(7) = FormatNumber(CR.AVG_PRICE)
      Values(8) = CR.USER_APPLOVE_PRICE
      Values(9) = CR.PART_ITEM_ID
      
'      Values(1) = CR.SALE_ORDER_ID
'      Values(2) = RealIndex
'       Values(3) = CR.ShowDescText
'      Values(4) = FormatNumber(CR.ITEM_AMOUNT)
'      Values(5) = FormatNumber(CR.TOTAL_PRICE)
'      Values(6) = FormatNumber(CR.AVG_PRICE)
'      Values(7) = CR.USER_APPLOVE_PRICE
'      Values(8) = CR.PART_ITEM_ID
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   If m_BillingDoc.SaleOrders.Count > 0 Then
      cboRateType.Enabled = False
   End If
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub



Private Sub TabStrip1_Click()
   GridEX1.Top = TabStrip1.Top + TabStrip1.HEIGHT '5160
   GridEX1.Left = 150
   GridEX1.Visible = False
   
   SSFrame2.Top = GridEX1.Top '5160
   SSFrame2.Left = 150
   SSFrame2.Visible = False
   
   If TabStrip1.SelectedItem.Index = 1 Then
      Call InitGrid1
      GridEX1.Visible = True
      
      Call GetTotalPrice
      GridEX1.ItemCount = CountItem(m_BillingDoc.SaleOrders)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      SSFrame2.Visible = True
      
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub txtAgreementData_Change()
   m_HasModify = True
End Sub

Private Sub txtAgreementFinance_Change()
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
Private Sub txtDocumentNo_LostFocus()
   If Not CheckUniqueNs(DO_PLAN_UNIQUE, txtDocumentNo.Text, id) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtDocumentNo.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Exit Sub
   End If
End Sub
Private Sub txtIncludeDiscount_Change()
   m_HasModify = True
   Call CalculateAmount
End Sub

Private Sub txtLeft_Change()
   m_HasModify = True
End Sub

Private Sub txtNetTotal_Change()
   Call CalculateAmount
   m_HasModify = True
End Sub

Private Sub txtTotalAmount_Change()
   m_HasModify = True
End Sub

Private Sub CalculateAmount()
   txtIncludeDiscount.Text = Val(txtNetTotal.Text) - Val(txtDiscount.Text)
   txtLeft.Text = Val(txtIncludeDiscount.Text) - Val(txtDeposit.Text)
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

Private Sub txtTotalDiscount_Change()
   m_HasModify = True
   'txtNetTotal.Text = Format(Val(txtTotalAmount.Text) + Val(txtTotalDiscount.Text), "0.00")
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

Private Sub uctlApproveDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlDeliveryCusLookup_Change()
   m_HasModify = True
   
   If m_BillingDoc.DELIVERY_CUS_ITEM_ID <> uctlDeliveryCusLookup.MyCombo.ItemData(Minus2Zero(uctlDeliveryCusLookup.MyCombo.ListIndex)) Then
      ISuctlDeliveryCusLookup = True
   End If
End Sub

Private Sub uctlDocumentDate_HasChange()
   m_HasModify = True
   DocumentDate = uctlDocumentDate.ShowDate
End Sub

Private Sub uctlCustomerLookup_Change()

Dim Customer As CCustomer
Dim TempD2 As CCustomer

   CUSTOMER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
   
     
   If CUSTOMER_ID > 0 Then
    If ShowMode = SHOW_ADD Then
       Call LoadDeliveryCus(uctlDeliveryCusLookup.MyCombo, m_DeliveryCus, CUSTOMER_ID, , , , "N") 'LOAD สถานที่จัดส่ง เฉพาะที่ ที่เปิดใช้งาน
   Else
      Call LoadDeliveryCus(uctlDeliveryCusLookup.MyCombo, m_DeliveryCus, CUSTOMER_ID) 'LOAD สถานที่จัดส่ง ทุกที่
   End If
      Set uctlDeliveryCusLookup.MyCollection = m_DeliveryCus
     
     Set TempD2 = GetObject("CCustomer", m_Customers, Trim(str(CUSTOMER_ID)), False)
      If Not TempD2 Is Nothing Then
         cboRateType.ListIndex = IDToListIndex(cboRateType, TempD2.PRICE_THINK_TYPE)
         CAL_RATE_DELIVERY_TYPE = TempD2.CAL_RATE_DELIVERY_TYPE
         CAL_PRICE_PART_CENTER_FLAG = TempD2.CAL_PRICE_PART_CENTER_FLAG
         CAL_PRICE_DLC_CENTER_FLAG = TempD2.CAL_PRICE_DLC_CENTER_FLAG
      Else
        cboRateType.ListIndex = -1
      End If
      
      
      If Area = 1 Then
         Set Customer = m_Customers(Trim(str(CUSTOMER_ID)))
         Call LoadAccount(cboAccount, , CUSTOMER_ID)
         cboAccount.ListIndex = 1
   
         Call LoadCustomerAddress(cboCustomerAddress, , CUSTOMER_ID, True)
         If Customer.RESPONSE_BY > 0 Then
            uctlSellByLookup.MyCombo.ListIndex = IDToListIndex(uctlSellByLookup.MyCombo, Customer.RESPONSE_BY)
         Else
            uctlSellByLookup.MyCombo.ListIndex = -1
         End If
      ElseIf Area = 2 Then
         Call LoadAccount(cboAccount, , CUSTOMER_ID)
         cboAccount.ListIndex = -1
   
         Call LoadSupplierAddress(cboCustomerAddress, , CUSTOMER_ID, True)
      End If
   Else
      cboAccount.ListIndex = -1
      cboCustomerAddress.ListIndex = -1
   End If
   m_HasModify = True
End Sub

Private Sub uctlEstimateDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlResourceLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlSellByLookup_Change()
   m_HasModify = True
End Sub
