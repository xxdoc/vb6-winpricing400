VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddEditPackingList 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditPackingList.frx":0000
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
      TabIndex        =   43
      Top             =   0
      Visible         =   0   'False
      Width           =   1635
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   8520
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSFrame SSFrame2 
         Height          =   2835
         Left            =   150
         TabIndex        =   44
         Top             =   4920
         Visible         =   0   'False
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   5001
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin prjFarmManagement.uctlTextBox txtReferInv 
            Height          =   435
            Left            =   1830
            TabIndex        =   46
            Top             =   480
            Width           =   2925
            _ExtentX        =   2831
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtPack 
            Height          =   435
            Left            =   1830
            TabIndex        =   48
            Top             =   960
            Width           =   5205
            _ExtentX        =   2831
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtShipName 
            Height          =   435
            Left            =   1830
            TabIndex        =   50
            Top             =   1440
            Width           =   5205
            _ExtentX        =   2831
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtShippingMark 
            Height          =   435
            Left            =   1830
            TabIndex        =   52
            Top             =   1920
            Width           =   5205
            _ExtentX        =   2831
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlDate uctlShipOnDate 
            Height          =   405
            Left            =   7530
            TabIndex        =   53
            Top             =   480
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   714
         End
         Begin VB.Label lblShipOnDate 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6000
            TabIndex        =   54
            Top             =   510
            Width           =   1395
         End
         Begin VB.Label lblShippingMark 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            TabIndex        =   45
            Top             =   2040
            Width           =   1695
         End
         Begin VB.Label lblShipName 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            TabIndex        =   51
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label lblPack 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            TabIndex        =   49
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label lblReferInv 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            TabIndex        =   47
            Top             =   600
            Width           =   1695
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
         TabIndex        =   16
         Top             =   4380
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
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   2805
         Left            =   150
         TabIndex        =   17
         Top             =   4920
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
         Column(1)       =   "frmAddEditPackingList.frx":27A2
         Column(2)       =   "frmAddEditPackingList.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditPackingList.frx":290E
         FormatStyle(2)  =   "frmAddEditPackingList.frx":2A6A
         FormatStyle(3)  =   "frmAddEditPackingList.frx":2B1A
         FormatStyle(4)  =   "frmAddEditPackingList.frx":2BCE
         FormatStyle(5)  =   "frmAddEditPackingList.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditPackingList.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   26
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
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlSellByLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   13
         Top             =   3570
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
         TabIndex        =   38
         Top             =   3120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
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
         MouseIcon       =   "frmAddEditPackingList.frx":2F36
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
         MouseIcon       =   "frmAddEditPackingList.frx":3250
         ButtonStyle     =   3
      End
      Begin VB.Label lblDiscount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         TabIndex        =   42
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label6 
         Height          =   315
         Left            =   3540
         TabIndex        =   41
         Top             =   3210
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label lblIncludeDiscount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7350
         TabIndex        =   40
         Top             =   3210
         Width           =   1815
      End
      Begin VB.Label Label3 
         Height          =   315
         Left            =   10860
         TabIndex        =   39
         Top             =   3180
         Visible         =   0   'False
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
         TabIndex        =   37
         Top             =   3630
         Width           =   1635
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   525
         Left            =   8250
         TabIndex        =   14
         Top             =   3600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPackingList.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   525
         Left            =   9870
         TabIndex        =   15
         Top             =   3600
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
         TabIndex        =   36
         Top             =   2340
         Width           =   1635
      End
      Begin VB.Label lblCustomerAddress 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   35
         Top             =   1890
         Width           =   1635
      End
      Begin VB.Label lblAccountNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7920
         TabIndex        =   34
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblCustomer 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   150
         TabIndex        =   33
         Top             =   1410
         Width           =   1635
      End
      Begin VB.Label Label4 
         Height          =   315
         Left            =   10860
         TabIndex        =   32
         Top             =   2730
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label lblNetTotal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8190
         TabIndex        =   31
         Top             =   2760
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label lblTotalDiscount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4080
         TabIndex        =   30
         Top             =   2760
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label2 
         Height          =   315
         Left            =   7350
         TabIndex        =   29
         Top             =   2730
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label Label1 
         Height          =   315
         Left            =   3540
         TabIndex        =   28
         Top             =   2760
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5040
         TabIndex        =   27
         Top             =   960
         Width           =   1395
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   21
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPackingList.frx":3884
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   22
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
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPackingList.frx":3B9E
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   20
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPackingList.frx":3EB8
         ButtonStyle     =   3
      End
      Begin VB.Label lblTotalAmount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         TabIndex        =   25
         Top             =   2790
         Width           =   1695
      End
      Begin VB.Label lblDocumentNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   90
         TabIndex        =   24
         Top             =   960
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmAddEditPackingList"
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

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public id As Long
Public Area As Long
Public DocumentType As Long

Private FileName As String
Private m_SumUnit As Double

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
'      txtTotalAmount.Text = Format(m_BillingDoc.TOTAL_AMOUNT, "0.00")
      txtTotalDiscount.Text = Format(m_BillingDoc.DISCOUNT_AMOUNT, "0.00")
      uctlSellByLookup.MyCombo.ListIndex = IDToListIndex(uctlSellByLookup.MyCombo, m_BillingDoc.ACCEPT_BY)
      'uctlResourceLookup.MyCombo.ListIndex = IDToListIndex(uctlResourceLookup.MyCombo, m_BillingDoc.RESOURCE_ID)
      'uctlEstimateDate.ShowDate = m_BillingDoc.ESTIMATE_DATE
      'uctlApproveDate.ShowDate = m_BillingDoc.APPROVE_DATE
      
      txtReferInv.Text = m_BillingDoc.REFER_INV
      txtPack.Text = m_BillingDoc.PACKING_OF
      txtShipName.Text = m_BillingDoc.SHIP_NAME
      txtShippingMark.Text = m_BillingDoc.SHIPPING_MARKS
      uctlShipOnDate.ShowDate = m_BillingDoc.SHIPON_DATE
            
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
      
   If Not VerifyDate(lblShipOnDate, uctlShipOnDate, True) Then
      Exit Function
   End If
   
      
   If Not CheckUniqueNs(DO_PLAN_UNIQUE, txtDocumentNo.Text, id) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtDocumentNo.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
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
      m_BillingDoc.DOCUMENT_TYPE = DocumentType
   ElseIf Area = 2 Then
      m_BillingDoc.SUPPLIER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
      m_BillingDoc.ACCOUNT_ID = -1
      m_BillingDoc.DOCUMENT_TYPE = DocumentType
   End If
   m_BillingDoc.BILLING_ADDRESS_ID = cboCustomerAddress.ItemData(Minus2Zero(cboCustomerAddress.ListIndex))
   m_BillingDoc.ENTERPRISE_ADDRESS_ID = cboEnpAddress.ItemData(Minus2Zero(cboEnpAddress.ListIndex))
   m_BillingDoc.EXCEPTION_FLAG = "N"
   m_BillingDoc.ACCEPT_BY = uctlSellByLookup.MyCombo.ItemData(Minus2Zero(uctlSellByLookup.MyCombo.ListIndex))
   m_BillingDoc.COMMIT_FLAG = Check2Flag(chkCommit.Value)
   m_BillingDoc.DISCOUNT_AMOUNT = Val(txtDiscount.Text)
   m_BillingDoc.TOTAL_AMOUNT = Val(txtTotalAmount.Text)
   m_BillingDoc.TOTAL_PRICE = Val(txtNetTotal.Text)
   'm_BillingDoc.RESOURCE_ID = uctlResourceLookup.MyCombo.ItemData(Minus2Zero(uctlResourceLookup.MyCombo.ListIndex))
   'm_BillingDoc.ESTIMATE_DATE = uctlEstimateDate.ShowDate
   'm_BillingDoc.APPROVE_DATE = uctlApproveDate.ShowDate
   
   m_BillingDoc.REFER_INV = txtReferInv.Text
   m_BillingDoc.PACKING_OF = txtPack.Text
   m_BillingDoc.SHIP_NAME = txtShipName.Text
   m_BillingDoc.SHIPPING_MARKS = txtShippingMark.Text
   m_BillingDoc.SHIPON_DATE = uctlShipOnDate.ShowDate
   
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
'   If Not glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog) Then
'      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'      SaveData = False
'      Call glbDaily.RollbackTransaction
'      Call EnableForm(Me, True)
'      Exit Function
'   End If
   
'   m_BillingDoc.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID
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

GridEX1.ItemCount = CountItem(m_BillingDoc.Pkglsts)
   GridEX1.Rebind
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
    Set frmAddEditPackingListItem.TempCollection = m_BillingDoc.Pkglsts
      frmAddEditPackingListItem.ParentShowMode = ShowMode
      frmAddEditPackingListItem.ShowMode = SHOW_ADD
      frmAddEditPackingListItem.HeaderText = MapText("เพิ่มรายละเอียดใบบรรจุหีบห่อ")
      Load frmAddEditPackingListItem
      frmAddEditPackingListItem.Show 1

      OKClick = frmAddEditPackingListItem.OKClick

      Unload frmAddEditPackingListItem
      Set frmAddEditPackingListItem = Nothing

      If OKClick Then
           Call GetTotalPrice
         GridEX1.ItemCount = CountItem(m_BillingDoc.Pkglsts)
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

Private Sub cmdAuto_Click()
Dim No As String

   If Trim(txtDocumentNo.Text) = "" Then
      Call glbDatabaseMngr.GenerateNumber(PKGLST_NUMBER, No, glbErrorLog)
      txtDocumentNo.Text = No
   End If
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
      If ID1 <= 0 Then
         m_BillingDoc.Pkglsts.Remove (ID2)
      Else
         m_BillingDoc.Pkglsts.Item(ID2).Flag = "D"
      End If

      Call GetTotalPrice
      GridEX1.ItemCount = CountItem(m_BillingDoc.Pkglsts)
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
      Set frmAddEditPackingListItem.TempCollection = m_BillingDoc.Pkglsts
      frmAddEditPackingListItem.HeaderText = MapText("แก้ไขรายละเอียดใบบรรจุหีบห่อ")
      frmAddEditPackingListItem.ParentShowMode = ShowMode
      frmAddEditPackingListItem.id = id
      frmAddEditPackingListItem.ShowMode = SHOW_EDIT
      Load frmAddEditPackingListItem
      frmAddEditPackingListItem.Show 1

      OKClick = frmAddEditPackingListItem.OKClick

      Unload frmAddEditPackingListItem
      Set frmAddEditPackingListItem = Nothing

      If OKClick Then
         Call GetTotalPrice
         GridEX1.ItemCount = CountItem(m_BillingDoc.Pkglsts)
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
'
'Private Function GetReportClassStr(Ind As Long, ProgramOwner As String) As String
'   If (Ind = 1) Or (Ind = 2) Then
'      GetReportClassStr = "CReportNormalPO"
'   ElseIf (Ind = 4) Or (Ind = 5) Or (Ind = 6) Then
'         GetReportClassStr = ""
'   End If
'End Function
'
'Private Function GetReportClass(Ind As Long, ProgramOwner As String) As CReportInterface
'Dim Rp As CReportInterface
'
'   If Ind = 1 Then
'      Set Rp = New CReportNormalPO
'      Set GetReportClass = Rp
'   ElseIf (Ind = 4) Or (Ind = 5) Then
'         Set GetReportClass = Nothing
'   End If
'End Function

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
   lMenuChosen = oMenu.AddMenu(glbGuiConfigs.PkglistPrintMenuItems)
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   Set oMenu = Nothing
   
'   If Not VerifyOnwerVersionMenu(lMenuChosen, glbParameterObj.Programowner) Then
'      Exit Sub
'   End If
   
   If lMenuChosen = 1 Then
      ReportKey = "CReportNormalPkglistH"
      
      Set Report = New CReportNormalPkglistH
      ReportFlag = True
   ElseIf lMenuChosen = 2 Then
      ReportKey = "CReportNormalPkglistH"
      
      Set Rc = New CReportConfig
      Rc.REPORT_KEY = ReportKey
      Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
      Call Rc.QueryData(m_Rs, iCount)
      HeaderText = MapText("ใบบรรจุหีบห่อ")
      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
   ElseIf lMenuChosen = 4 Then
      ReportKey = "CReportNormalPkglist"
      
      Set Report = New CReportNormalPkglist
      ReportFlag = True
   ElseIf lMenuChosen = 5 Then
      ReportKey = "CReportNormalPkglist"
      
      Set Report = New CReportNormalPkglist
      ReportFlag = True
   ElseIf lMenuChosen = 8 Then
      ReportKey = "CReportNormalPkglist"
      ReportMode = 2
      
      Set Rc = New CReportConfig
      Rc.REPORT_KEY = ReportKey
      Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
      Call Rc.QueryData(m_Rs, iCount)
      HeaderText = MapText("ใบเสนอราคา")
      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
     ElseIf lMenuChosen = 10 Then
      ReportKey = "CReportNormalPkglistH"
      
      Set Report = New CReportNormalPkglistH
      ReportFlag = True
   ElseIf lMenuChosen = 11 Then
      ReportKey = "CReportNormalPkglistH"
      
      Set Rc = New CReportConfig
      Rc.REPORT_KEY = ReportKey
      Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
      Call Rc.QueryData(m_Rs, iCount)
      HeaderText = MapText("ใบบรรจุหีบห่อ")
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
      Call Report.AddParam(MapText("ใบบรรจุหีบห่อ"), "REPORT_HEADER")
      Call Report.AddParam(Picture2.Picture, "BACK_GROUND")
      Call Report.AddParam(uctlSellByLookup.MyCombo.Text, "RECEIVE_NAME")
      Call Report.AddParam("", "ACCEPT_NAME")
   ElseIf lMenuChosen = 5 Then
      ReportKey = "CReportNormalQua"
      
      Set Rc = New CReportConfig
      Rc.REPORT_KEY = ReportKey
      Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
      Call Rc.QueryData(m_Rs, iCount)
      HeaderText = MapText("ใบบรรจุหีบห่อ")
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
   
   VerifyOnwerVersionMenu = False
   
   If Not VerifyOnwerVersionMenu Then
      glbErrorLog.LocalErrorMsg = "โปรแกรมไม่สนับสนุนฟังก์ชันนี้ในเวอร์ชันนี้"
      glbErrorLog.ShowUserError
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
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
'      DoEvents
      
      Call EnableForm(Me, False)
      Call LoadEnterpriseAddress(cboEnpAddress, , , True)
      
      If Area = 1 Then
         Call LoadCustomer(uctlCustomerLookup.MyCombo, m_Customers)
         Set uctlCustomerLookup.MyCollection = m_Customers
      ElseIf Area = 2 Then
         Call LoadSupplier(uctlCustomerLookup.MyCombo, m_Customers)
         Set uctlCustomerLookup.MyCollection = m_Customers
      End If
      
      Call LoadEmployee(uctlSellByLookup.MyCombo, m_Employees)
      Set uctlSellByLookup.MyCollection = m_Employees
      
      'Call LoadResource(uctlResourceLookup.MyCombo, m_Resources)
      'Set uctlResourceLookup.MyCollection = m_Resources
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_BillingDoc.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         uctlDocumentDate.ShowDate = Now
         m_BillingDoc.QueryFlag = 0
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
   
   Set m_BillingDoc = Nothing
   Set m_Customers = Nothing
   Set m_Employees = Nothing
   Set m_Resources = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''''Debug.Print ColIndex & " " & NewColWidth
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
   Col.Width = 2000
   Col.Caption = MapText("หมายเลขบรรจุหีบห่อ")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 2325 + 400
   Col.Caption = MapText("รายละเอียด")
   
   Set Col = GridEX1.Columns.add '5
   Col.TextAlignment = jgexAlignRight
   Col.Width = 1575
   Col.Caption = MapText("ปริมาณ/หีบห่อ")
   
   Set Col = GridEX1.Columns.add '6
   Col.TextAlignment = jgexAlignRight
   Col.Width = 1755
   Col.Caption = MapText("จำนวนหีบห่อ")
   
   
   Set Col = GridEX1.Columns.add '7
   Col.TextAlignment = jgexAlignRight
   Col.Width = 1755
   Col.Caption = MapText("น้ำหนักสุทธิ")
   
   
   Set Col = GridEX1.Columns.add '8
   Col.TextAlignment = jgexAlignRight
   Col.Width = 1755
   Col.Caption = MapText("น้ำหนักทั้งหมด")
End Sub

Private Sub GetTotalPrice()
Dim II As CPkglst
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
   For Each II In m_BillingDoc.Pkglsts
      If II.Flag <> "D" Then
          Sum3 = Sum3 + (II.QUANTITY * II.MEASURE)
         Sum4 = Sum4 + (II.NET_WEIGHT * II.MEASURE)
         Sum5 = Sum5 + (II.GROSS_WEIGHT * II.MEASURE)
      End If
   Next II
   txtTotalAmount.Text = Sum3
   txtIncludeDiscount.Text = Format(Sum5, "0.00")
   txtDiscount.Text = Format(Sum4, "0.00")
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame2.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblDocumentNo, MapText("เลขที่เอกสาร"))
   Call InitNormalLabel(lblAccountNo, MapText("เลขที่บัญชี"))
   If Area = 1 Then
      Call InitNormalLabel(lblCustomerAddress, MapText("ที่อยู่ลูกค้า"))
      Call InitNormalLabel(lblCustomer, MapText("รหัสลูกค้า"))
      Call InitNormalLabel(lblSellBy, MapText("พนักงานขาย"))
      Call InitNormalLabel(lblEnpAddress, MapText("ที่อยู่ออกเอกสาร"))
   ElseIf Area = 2 Then
      Call InitNormalLabel(lblCustomerAddress, MapText("ที่อยู่ซัพ ฯ"))
      Call InitNormalLabel(lblCustomer, MapText("รหัสซัพ ฯ"))
      Call InitNormalLabel(lblSellBy, MapText("ผู้ทำรายการ"))
      Call InitNormalLabel(lblEnpAddress, MapText("ที่อยู่รับเอกสาร"))
      
      lblAccountNo.Visible = False
      cboAccount.Visible = False
      cmdAuto.Visible = False
      cmdCustomer.Visible = False
      cmdPrint.Enabled = False
   End If
   Call InitNormalLabel(lblTotalAmount, MapText("จำนวนรวม"))
   Call InitNormalLabel(lblDocumentDate, MapText("วันที่เอกสาร"))
   Call InitNormalLabel(lblTotalDiscount, MapText("น้ำหนักรวม"))
   Call InitNormalLabel(Label1, MapText("ตัว"))
   Call InitNormalLabel(Label2, MapText("ก.ก."))
   Call InitNormalLabel(Label4, MapText("บาท"))
   Call InitNormalLabel(lblNetTotal, MapText("ราคารวม"))
   Call InitNormalLabel(lblDiscount, MapText("น้ำหนักสุทธิรวม"))
   Call InitNormalLabel(lblIncludeDiscount, MapText("น้ำหนักทั้งหมดรวม"))
   Call InitCheckBox(chkCommit, "คำนวณ")
   Call InitNormalLabel(Label3, MapText("บาท"))
   Call InitNormalLabel(Label6, MapText("บาท"))
   
   Call InitNormalLabel(lblReferInv, MapText("REFER INV. NO."))
   Call InitNormalLabel(lblPack, MapText("PACKING OF"))
   Call InitNormalLabel(lblShipName, MapText("SHIPPED PER"))
   Call InitNormalLabel(lblShippingMark, MapText("SHIPPING MARK"))
   Call InitNormalLabel(lblShipOnDate, MapText("วันที่ขนส่ง"))
   
   Call txtReferInv.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtPack.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtShipName.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtShippingMark.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   
   Call txtDocumentNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtTotalAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtTotalAmount.Enabled = False
   Call txtTotalDiscount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtTotalDiscount.Enabled = False
   Call txtNetTotal.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtNetTotal.Enabled = False
   Call txtIncludeDiscount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtIncludeDiscount.Enabled = False
   Call txtDiscount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtDiscount.Enabled = False
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   GridEX1.Visible = True
  
   Call InitCombo(cboAccount)
   Call InitCombo(cboCustomerAddress)
   Call InitCombo(cboEnpAddress)
   
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
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSave, MapText("บันทึก"))
   Call InitMainButton(cmdPrint, MapText("พิมพ์"))
   Call InitMainButton(cmdAuto, MapText("A"))
   Call InitMainButton(cmdCustomer, MapText("F"))
   
   Call InitGrid1
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.NAME = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.add().Caption = MapText("รายการใบบรรจุหีบห่อ")
   TabStrip1.Tabs.add().Caption = MapText("รายละเอียดหลัก")
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
      If m_BillingDoc.Pkglsts Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim CR As CPkglst
      If m_BillingDoc.Pkglsts.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_BillingDoc.Pkglsts, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If

      Values(1) = CR.PKGLST_ITEM_ID
      Values(2) = RealIndex
      Values(3) = CR.PKG_NUMBER
      Values(4) = CR.DESCRIPTION
      Values(5) = CR.QUANTITY
      Values(6) = CR.MEASURE
      Values(7) = FormatNumber(CR.NET_WEIGHT)
      Values(8) = FormatNumber(CR.GROSS_WEIGHT)
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
   GridEX1.Top = 4920
   GridEX1.Left = 150
   GridEX1.Visible = False
   SSFrame2.Visible = False
   If TabStrip1.SelectedItem.Index = 1 Then
      Call InitGrid1
      GridEX1.Visible = True
      
      Call GetTotalPrice
      GridEX1.ItemCount = CountItem(m_BillingDoc.Pkglsts)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      SSFrame2.Visible = True
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub txtLeft_Change()
   m_HasModify = True
End Sub

Private Sub txtPack_Change()
   m_HasModify = True
End Sub

Private Sub txtReferInv_Change()
   m_HasModify = True
End Sub

Private Sub txtShipName_Change()
   m_HasModify = True
End Sub

Private Sub txtShippingMark_Change()
   m_HasModify = True
End Sub

Private Sub txtTotalAmount_Change()
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

Private Sub txtTotalDiscount_Change()
   m_HasModify = True
   txtNetTotal.Text = Format(Val(txtTotalAmount.Text) + Val(txtTotalDiscount.Text), "0.00")
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

Private Sub uctlDocumentDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlCustomerLookup_Change()
Dim CustomerID As Long
Dim Customer As CCustomer

   CustomerID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
   If CustomerID > 0 Then
      If Area = 1 Then
         Set Customer = m_Customers(Trim(str(CustomerID)))
         Call LoadAccount(cboAccount, , CustomerID)
         cboAccount.ListIndex = 1
   
         Call LoadCustomerAddress(cboCustomerAddress, , CustomerID, True)
         If Customer.RESPONSE_BY > 0 Then
            uctlSellByLookup.MyCombo.ListIndex = IDToListIndex(uctlSellByLookup.MyCombo, Customer.RESPONSE_BY)
         Else
            uctlSellByLookup.MyCombo.ListIndex = -1
         End If
      ElseIf Area = 2 Then
         Call LoadAccount(cboAccount, , CustomerID)
         cboAccount.ListIndex = -1
   
         Call LoadSupplierAddress(cboCustomerAddress, , CustomerID, True)
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

Private Sub uctlShipOnDate_HasChange()
   m_HasModify = True
End Sub
