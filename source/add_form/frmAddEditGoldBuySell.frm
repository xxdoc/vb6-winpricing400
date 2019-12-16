VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddEditGoldBuySell 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditGoldBuySell.frx":0000
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
      Begin VB.ComboBox cboDoType 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1500
         Width           =   2925
      End
      Begin VB.ComboBox cboEnpAddress 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2850
         Width           =   9585
      End
      Begin VB.ComboBox cboCustomerAddress 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2400
         Width           =   9585
      End
      Begin VB.ComboBox cboAccount 
         Height          =   315
         Left            =   8520
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1950
         Width           =   2925
      End
      Begin prjFarmManagement.uctlTextLookup uctlCustomerLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   4
         Top             =   1950
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlDocumentDate 
         Height          =   405
         Left            =   6570
         TabIndex        =   1
         Top             =   1080
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   9
         Top             =   4050
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
         Top             =   1050
         Width           =   2925
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
         Left            =   150
         TabIndex        =   10
         Top             =   4590
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
         Column(1)       =   "frmAddEditGoldBuySell.frx":27A2
         Column(2)       =   "frmAddEditGoldBuySell.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditGoldBuySell.frx":290E
         FormatStyle(2)  =   "frmAddEditGoldBuySell.frx":2A6A
         FormatStyle(3)  =   "frmAddEditGoldBuySell.frx":2B1A
         FormatStyle(4)  =   "frmAddEditGoldBuySell.frx":2BCE
         FormatStyle(5)  =   "frmAddEditGoldBuySell.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditGoldBuySell.frx":2D5E
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
      Begin prjFarmManagement.uctlTextLookup uctlSellByLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   8
         Top             =   3300
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin VB.Label lblDoType 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   27
         Top             =   1590
         Width           =   1575
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   525
         Left            =   9810
         TabIndex        =   26
         Top             =   3300
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   525
         Left            =   8190
         TabIndex        =   25
         Top             =   3300
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditGoldBuySell.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCheck chkCommit 
         Height          =   435
         Left            =   10500
         TabIndex        =   2
         Top             =   1050
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
         TabIndex        =   24
         Top             =   3360
         Width           =   1635
      End
      Begin VB.Label lblEnpAddress 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Top             =   2940
         Width           =   1635
      End
      Begin VB.Label lblCustomerAddress 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   2490
         Width           =   1635
      End
      Begin VB.Label lblAccountNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7320
         TabIndex        =   21
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblCustomer 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   150
         TabIndex        =   20
         Top             =   2010
         Width           =   1635
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5040
         TabIndex        =   19
         Top             =   1110
         Width           =   1395
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8505
         TabIndex        =   14
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditGoldBuySell.frx":3250
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
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditGoldBuySell.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   13
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditGoldBuySell.frx":3884
         ButtonStyle     =   3
      End
      Begin VB.Label lblDocumentNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   90
         TabIndex        =   17
         Top             =   1110
         Width           =   1665
      End
   End
End
Attribute VB_Name = "frmAddEditGoldBuySell"
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

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public id As Long

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
      uctlCustomerLookup.MyCombo.ListIndex = IDToListIndex(uctlCustomerLookup.MyCombo, m_BillingDoc.CUSTOMER_ID)
      cboAccount.ListIndex = IDToListIndex(cboAccount, m_BillingDoc.ACCOUNT_ID)
      cboCustomerAddress.ListIndex = IDToListIndex(cboCustomerAddress, m_BillingDoc.BILLING_ADDRESS_ID)
      cboEnpAddress.ListIndex = IDToListIndex(cboEnpAddress, m_BillingDoc.ENTERPRISE_ADDRESS_ID)
      uctlSellByLookup.MyCombo.ListIndex = IDToListIndex(uctlSellByLookup.MyCombo, m_BillingDoc.ACCEPT_BY)
      
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

Public Function GetExportItem(Ivd As CInventoryDoc, GuiID As Long) As CLotItem
Dim Ei As CLotItem

      For Each Ei In Ivd.ImportExports
         If Ei.LINK_ID = GuiID Then
            Set GetExportItem = Ei
            Exit Function
         End If
      Next Ei
End Function

Private Function DO2InventoryDoc(BD As CBillingDoc, Ivd As CInventoryDoc) As Boolean
Dim TempRs As ADODB.Recordset
Dim iCount As Long
Dim IsOK As Boolean
Dim Di As CDoItem
Dim Ei As CLotItem

   Set Ivd = Nothing
   Set Ivd = New CInventoryDoc

   If BD.INVENTORY_DOC_ID > 0 Then
      Set TempRs = New ADODB.Recordset
      
      Ivd.INVENTORY_DOC_ID = BD.INVENTORY_DOC_ID
      Ivd.QueryFlag = 1
      Call glbDaily.QueryInventoryDoc(Ivd, TempRs, iCount, IsOK, glbErrorLog)
      
      If TempRs.State = adStateOpen Then
         TempRs.Close
      End If
      Set TempRs = Nothing
      
      Ivd.AddEditMode = SHOW_EDIT
   Else
      Ivd.AddEditMode = SHOW_ADD
   End If
      
   Ivd.DOCUMENT_DATE = BD.DOCUMENT_DATE
   Ivd.DOCUMENT_NO = BD.DOCUMENT_NO
   Ivd.COMMIT_FLAG = BD.COMMIT_FLAG
   Ivd.DOCUMENT_TYPE = 10
   
   For Each Di In BD.DoItems
      If Di.Flag = "A" Then
         Set Ei = New CLotItem
         
         Ei.TX_TYPE = "E"
         Ei.Flag = "A"
         Ei.PART_ITEM_ID = Di.PART_ITEM_ID
         Ei.LOCATION_ID = Di.LOCATION_ID
         Ei.TX_AMOUNT = Di.ITEM_AMOUNT
         Ei.TOTAL_WEIGHT = Di.TOTAL_WEIGHT
         Ei.TOTAL_INCLUDE_PRICE = Di.TOTAL_PRICE
         Ei.LINK_ID = Di.LINK_ID
         Ei.CALCULATE_FLAG = "N"
         
         Call Ivd.ImportExports.add(Ei)
         Set Ei = Nothing
      ElseIf Di.Flag = "E" Then
         Set Ei = GetExportItem(Ivd, Di.LINK_ID)
         
         Ei.Flag = "E"
         Ei.PART_ITEM_ID = Di.PART_ITEM_ID
         Ei.LOCATION_ID = Di.LOCATION_ID
         Ei.TX_AMOUNT = Di.ITEM_AMOUNT
         Ei.TOTAL_WEIGHT = Di.TOTAL_WEIGHT
         Ei.CALCULATE_FLAG = "N"
      ElseIf Di.Flag = "D" Then
         Set Ei = GetExportItem(Ivd, Di.LINK_ID)
         Ei.Flag = "D"
      End If
   Next Di
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
   m_BillingDoc.CUSTOMER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
   m_BillingDoc.ACCOUNT_ID = cboAccount.ItemData(Minus2Zero(cboAccount.ListIndex))
   m_BillingDoc.BILLING_ADDRESS_ID = cboCustomerAddress.ItemData(Minus2Zero(cboCustomerAddress.ListIndex))
   m_BillingDoc.ENTERPRISE_ADDRESS_ID = cboEnpAddress.ItemData(Minus2Zero(cboEnpAddress.ListIndex))
   m_BillingDoc.DOCUMENT_TYPE = 1 'ใบส่งของ
   m_BillingDoc.EXCEPTION_FLAG = "N"
   m_BillingDoc.ACCEPT_BY = uctlSellByLookup.MyCombo.ItemData(Minus2Zero(uctlSellByLookup.MyCombo.ListIndex))
   m_BillingDoc.COMMIT_FLAG = Check2Flag(chkCommit.Value)
   Call PopulateGuiID(m_BillingDoc)
   
   Call EnableForm(Me, False)
   
   Call DO2InventoryDoc(m_BillingDoc, Ivd)
   
   If (m_BillingDoc.COMMIT_FLAG = "Y") Then
      If m_BillingDoc.OLD_COMMIT_FLAG <> "Y" Then
         Call glbDaily.TriggerCommit(Ivd.ImportExports)
         If Not glbDaily.VerifyStockBalance(Ivd.ImportExports, glbErrorLog) Then
            Call EnableForm(Me, True)
            Exit Function
         End If
         
      End If
   End If
   
   Call glbDaily.StartTransaction
   If Not glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call glbDaily.RollbackTransaction
      Call EnableForm(Me, True)
      Exit Function
   End If
   
   m_BillingDoc.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID
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

Private Sub cboDoType_Click()
   m_HasModify = True
End Sub

Private Sub cboDoType_KeyPress(KeyAscii As Integer)
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

Private Sub cmdAdd_Click()
Dim OKClick As Boolean

   If Not cmdAdd.Enabled Then
      Exit Sub
   End If
   
   OKClick = False
   If TabStrip1.SelectedItem.Index = 1 Then
      frmAddEditGoldDoItem1.COMMIT_FLAG = m_BillingDoc.COMMIT_FLAG
      Set frmAddEditGoldDoItem1.TempCollection = m_BillingDoc.DoItems
      frmAddEditGoldDoItem1.ParentShowMode = ShowMode
      frmAddEditGoldDoItem1.ShowMode = SHOW_ADD
      frmAddEditGoldDoItem1.HeaderText = MapText("เพิ่มรายการขาย")
      Load frmAddEditGoldDoItem1
      frmAddEditGoldDoItem1.Show 1

      OKClick = frmAddEditGoldDoItem1.OKClick

      Unload frmAddEditGoldDoItem1
      Set frmAddEditGoldDoItem1 = Nothing

      If OKClick Then
         Call GetTotalPrice

         GridEX1.ItemCount = 0 'CountItem(m_BillingDoc.DoItems)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      frmAddEditGoldDoItem2.COMMIT_FLAG = m_BillingDoc.COMMIT_FLAG
      Set frmAddEditGoldDoItem2.TempCollection = m_BillingDoc.DoItems
      frmAddEditGoldDoItem2.ParentShowMode = ShowMode
      frmAddEditGoldDoItem2.ShowMode = SHOW_ADD
      frmAddEditGoldDoItem2.HeaderText = MapText("เพิ่มรายการซื้อคืน")
      Load frmAddEditGoldDoItem2
      frmAddEditGoldDoItem2.Show 1

      OKClick = frmAddEditGoldDoItem2.OKClick

      Unload frmAddEditGoldDoItem2
      Set frmAddEditGoldDoItem2 = Nothing

      If OKClick Then
         Call GetTotalPrice

         GridEX1.ItemCount = 0
         GridEX1.Rebind
      End If
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
         m_BillingDoc.DoItems.Remove (ID2)
      Else
         m_BillingDoc.DoItems.Item(ID2).Flag = "D"
      End If

      Call GetTotalPrice
      GridEX1.ItemCount = CountItem(m_BillingDoc.DoItems)
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
      frmAddEditGoldDoItem1.id = id
      frmAddEditGoldDoItem1.COMMIT_FLAG = m_BillingDoc.COMMIT_FLAG
      Set frmAddEditGoldDoItem1.TempCollection = m_BillingDoc.DoItems
      frmAddEditGoldDoItem1.HeaderText = MapText("แก้ไขรายการขายทอง")
      frmAddEditGoldDoItem1.ParentShowMode = ShowMode
      frmAddEditGoldDoItem1.ShowMode = SHOW_EDIT
      Load frmAddEditGoldDoItem1
      frmAddEditGoldDoItem1.Show 1

      OKClick = frmAddEditGoldDoItem1.OKClick

      Unload frmAddEditGoldDoItem1
      Set frmAddEditGoldDoItem1 = Nothing

      If OKClick Then
         Call GetTotalPrice
         GridEX1.ItemCount = CountItem(m_BillingDoc.DoItems)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      frmAddEditGoldDoItem2.id = id
      frmAddEditGoldDoItem2.COMMIT_FLAG = m_BillingDoc.COMMIT_FLAG
      Set frmAddEditGoldDoItem2.TempCollection = m_BillingDoc.DoItems
      frmAddEditGoldDoItem2.HeaderText = MapText("แก้ไขรายการซื้อคืนทอง")
      frmAddEditGoldDoItem2.ParentShowMode = ShowMode
      frmAddEditGoldDoItem2.ShowMode = SHOW_EDIT
      Load frmAddEditGoldDoItem2
      frmAddEditGoldDoItem2.Show 1

      OKClick = frmAddEditGoldDoItem2.OKClick

      Unload frmAddEditGoldDoItem2
      Set frmAddEditGoldDoItem2 = Nothing

      If OKClick Then
         Call GetTotalPrice
         GridEX1.ItemCount = CountItem(m_BillingDoc.DoItems)
         GridEX1.Rebind
      End If
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
      DoEvents
      
      Call EnableForm(Me, False)
      Call LoadEnterpriseAddress(cboEnpAddress, , , True)
      
      Call LoadDoType(cboDoType)
      
      Call LoadCustomer(uctlCustomerLookup.MyCombo, m_Customers)
      Set uctlCustomerLookup.MyCollection = m_Customers
      
      Call LoadEmployee(uctlSellByLookup.MyCombo, m_Employees)
      Set uctlSellByLookup.MyCollection = m_Employees
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_BillingDoc.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
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
   Col.Width = 2010
   Col.Caption = MapText("ประเภททอง")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 3435
   Col.Caption = MapText("ชื่อทอง")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 1485
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("น้ำหนักทอง (g)")

   Set Col = GridEX1.Columns.add '6
   Col.Width = 1785
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("น้ำหนักทอง (บาท)")

   Set Col = GridEX1.Columns.add '7
   Col.Width = 1170
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ราคา/บาท")

   Set Col = GridEX1.Columns.add '8
   Col.Width = 1680
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("มูลค่าทอง (บาท)")
   
   Set Col = GridEX1.Columns.add '9
   Col.Width = 1350
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ค่าแรง (บาท)")
   
   Set Col = GridEX1.Columns.add '10
   Col.TextAlignment = jgexAlignRight
   Col.Width = 2160
   Col.Caption = MapText("รวม (บาท)")
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
   Col.Width = 2010
   Col.Caption = MapText("ประเภททอง")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 3435
   Col.Caption = MapText("ชื่อทอง")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 1485
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("น้ำหนักทอง (g)")

   Set Col = GridEX1.Columns.add '6
   Col.Width = 1785
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("น้ำหนักทอง (บาท)")

   Set Col = GridEX1.Columns.add '7
   Col.Width = 1170
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ราคา/บาท")

   Set Col = GridEX1.Columns.add '8
   Col.Width = 1680
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("มูลค่าทอง (บาท)")
   
   Set Col = GridEX1.Columns.add '9
   Col.Width = 1650
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ค่าน้ำกรด (บาท)")
   
   Set Col = GridEX1.Columns.add '10
   Col.TextAlignment = jgexAlignRight
   Col.Width = 2160
   Col.Caption = MapText("รวม (บาท)")
End Sub

Private Sub InitGrid3()
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
   Col.Width = 2010
   Col.Caption = MapText("วันที่ชำระเงิน")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 2010
   Col.Caption = MapText("ประเภทการชำระเงิน")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 1485
   Col.Caption = MapText("ทิศทาง")

   Set Col = GridEX1.Columns.add '6
   Col.Width = 1785
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("จำนวนเงิน (บาท)")

   Set Col = GridEX1.Columns.add '7
   Col.Width = 4275
   Col.Caption = MapText("หมายเหตุ")
End Sub

Private Sub InitGrid4()
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
   Col.Width = 7995
   Col.Caption = MapText("รายละเอียด")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 1545
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("จำนวนเงิน (รับ)")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 1680
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("จำนวนเงิน (จ่าย)")
End Sub

Private Sub GetTotalPrice()
Dim II As CDoItem
Dim Sum1 As Double
Dim Sum2 As Double
Dim Sum3 As Double

   Sum1 = 0
   Sum2 = 0
   Sum3 = 0
   For Each II In m_BillingDoc.DoItems
      If II.Flag <> "D" Then
         Sum1 = Sum1 + II.ITEM_AMOUNT
         Sum2 = Sum2 + II.TOTAL_PRICE
         Sum3 = Sum3 + II.TOTAL_WEIGHT
      End If
   Next II
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblDocumentNo, MapText("เลขที่บิล"))
   Call InitNormalLabel(lblAccountNo, MapText("เลขที่บัญชี"))
   Call InitNormalLabel(lblCustomerAddress, MapText("ที่อยู่ลูกค้า"))
   Call InitNormalLabel(lblDocumentDate, MapText("วันที่เอกสาร"))
   Call InitNormalLabel(lblCustomer, MapText("รหัสลูกค้า"))
   Call InitNormalLabel(lblEnpAddress, MapText("ที่อยู่ออกเอกสาร"))
   Call InitNormalLabel(lblSellBy, MapText("พนักงานขาย"))
   Call InitNormalLabel(lblDoType, MapText("ประเภทบิล"))

   Call InitCheckBox(chkCommit, "ปิดบิล")
   
   Call txtDocumentNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   Call InitCombo(cboAccount)
   Call InitCombo(cboCustomerAddress)
   Call InitCombo(cboEnpAddress)
   Call InitCombo(cboDoType)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPrint.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPrint.Enabled = False
   cmdSave.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSave, MapText("บันทึก"))
   Call InitMainButton(cmdPrint, MapText("พิมพ์"))
   
   Call InitGrid1
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.NAME = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.add().Caption = MapText("รายการขายทอง")
   TabStrip1.Tabs.add().Caption = MapText("รายการซื้อทอง/รับคืนทอง/ทองเก่า")
   TabStrip1.Tabs.add().Caption = MapText("การชำระเงิน")
   TabStrip1.Tabs.add().Caption = MapText("รายการสรุป")
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
      If m_BillingDoc.DoItems Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim CR As CDoItem
      If m_BillingDoc.DoItems.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_BillingDoc.DoItems, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If

      Values(1) = CR.DO_ITEM_ID
      Values(2) = RealIndex
      Values(3) = CR.PART_NO
      Values(4) = CR.PIG_TYPE
      Values(5) = CR.PIG_STATUS_NAME
      Values(6) = FormatNumber(CR.ITEM_AMOUNT)
      Values(7) = FormatNumber(CR.TOTAL_WEIGHT)
      Values(8) = FormatNumber(CR.TOTAL_PRICE)
      Values(9) = FormatNumber(CR.AVG_PRICE)
      Values(10) = CR.LOCATION_NAME
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
      Call InitGrid1
      
      GridEX1.ItemCount = 0 'CountItem(m_BillingDoc.DoItems)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      Call InitGrid2
      
      GridEX1.ItemCount = 0 'CountItem(m_BillingDoc.DoItems)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
      Call InitGrid3
      
      GridEX1.ItemCount = 0 'CountItem(m_BillingDoc.DoItems)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
      Call InitGrid4
      
      GridEX1.ItemCount = 0 'CountItem(m_BillingDoc.DoItems)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
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

Private Sub uctlDocumentDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlCustomerLookup_Change()
Dim CustomerID As Long

   CustomerID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
   If CustomerID > 0 Then
      Call LoadAccount(cboAccount, , CustomerID)
      cboAccount.ListIndex = 1
      
      Call LoadCustomerAddress(cboCustomerAddress, , CustomerID, True)
   Else
      cboAccount.ListIndex = -1
      cboCustomerAddress.ListIndex = -1
   End If
   m_HasModify = True
End Sub

Private Sub uctlSellByLookup_Change()
   m_HasModify = True
End Sub
