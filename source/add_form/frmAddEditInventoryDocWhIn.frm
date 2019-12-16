VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddEditInventoryDocWhIn 
   ClientHeight    =   9570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13260
   ForeColor       =   &H00000000&
   Icon            =   "frmAddEditInventoryDocWhIn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   13260
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   16320
      ScaleHeight     =   1035
      ScaleWidth      =   555
      TabIndex        =   29
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   9840
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   16215
      _ExtentX        =   28601
      _ExtentY        =   17357
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   120
         TabIndex        =   28
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
      Begin prjFarmManagement.uctlTextLookup uctlCustomerLookup 
         Height          =   435
         Left            =   6000
         TabIndex        =   3
         Top             =   1260
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlDocumentDate 
         Height          =   405
         Left            =   6000
         TabIndex        =   1
         Top             =   840
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtDoNo 
         Height          =   435
         Left            =   1560
         TabIndex        =   2
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
         Width           =   2775
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTruckNo 
         Height          =   435
         Left            =   1560
         TabIndex        =   5
         Top             =   1800
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   3375
         Left            =   120
         TabIndex        =   6
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
         Column(1)       =   "frmAddEditInventoryDocWhIn.frx":27A2
         Column(2)       =   "frmAddEditInventoryDocWhIn.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditInventoryDocWhIn.frx":290E
         FormatStyle(2)  =   "frmAddEditInventoryDocWhIn.frx":2A6A
         FormatStyle(3)  =   "frmAddEditInventoryDocWhIn.frx":2B1A
         FormatStyle(4)  =   "frmAddEditInventoryDocWhIn.frx":2BCE
         FormatStyle(5)  =   "frmAddEditInventoryDocWhIn.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditInventoryDocWhIn.frx":2D5E
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
      Begin prjFarmManagement.uctlTextBox txtDesc 
         Height          =   435
         Left            =   6000
         TabIndex        =   4
         Top             =   1710
         Width           =   5385
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin Threed.SSFrame SSFrame4 
         Height          =   3255
         Left            =   120
         TabIndex        =   23
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
            TabIndex        =   25
            Top             =   240
            Width           =   4035
         End
         Begin VB.ComboBox cboPaidType 
            Height          =   315
            Left            =   2760
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   720
            Width           =   4005
         End
         Begin VB.Label lblPaidType 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   840
            TabIndex        =   26
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label lblCondition 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   360
            TabIndex        =   27
            Top             =   240
            Width           =   2295
         End
      End
      Begin Threed.SSCommand cmdOther 
         Height          =   525
         Left            =   5100
         TabIndex        =   10
         Top             =   8880
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditInventoryDocWhIn.frx":2F36
         ButtonStyle     =   3
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2730
         TabIndex        =   22
         Top             =   1800
         Width           =   435
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   525
         Left            =   6720
         TabIndex        =   11
         Top             =   8880
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditInventoryDocWhIn.frx":3250
         ButtonStyle     =   3
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4470
         TabIndex        =   21
         Top             =   1740
         Width           =   1455
      End
      Begin VB.Label lblCustomerNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4470
         TabIndex        =   20
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4710
         TabIndex        =   19
         Top             =   870
         Width           =   1155
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8400
         TabIndex        =   12
         Top             =   8880
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditInventoryDocWhIn.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10080
         TabIndex        =   13
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   8880
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditInventoryDocWhIn.frx":3884
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3480
         TabIndex        =   9
         Top             =   8880
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditInventoryDocWhIn.frx":3B9E
         ButtonStyle     =   3
      End
      Begin VB.Label lblTruckNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   0
         TabIndex        =   17
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label lblDocumentNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -150
         TabIndex        =   16
         Top             =   900
         Width           =   1665
      End
      Begin VB.Label lblDoNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -60
         TabIndex        =   15
         Top             =   1380
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmAddEditInventoryDocWhIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_InventoryWHDoc As CInventoryWHDoc
Private m_LotItemWh As CLotItemWH
Private m_Weight As CWeight
Private m_Customers As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public AutoGenPo As Boolean

Public ID As Long
Public ID2 As Long
Public DocumentType As Long

Private FileName As String
Private m_Cd As Collection
Private TempWeight As Collection
Private CW As CWeight
Private DocAdd As Long
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long
   
   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
      
      m_InventoryWHDoc.INVENTORY_WH_DOC_ID = ID
      m_InventoryWHDoc.COMMIT_FLAG = ""
      m_InventoryWHDoc.QueryFlag = 1
      
      If Not glbDaily.QueryInventoryWhDoc(m_InventoryWHDoc, m_Rs, itemcount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
        Exit Sub
      End If
   End If
   
   If itemcount > 0 Then
     ' Call m_InventoryWHDoc.PopulateFromRS(1, m_Rs)

      uctlDocumentDate.ShowDate = m_InventoryWHDoc.DOCUMENT_DATE
      txtDoNo.Text = m_InventoryWHDoc.DO_NO
      txtTruckNo.Text = m_InventoryWHDoc.TRUCK_NO
      txtDocumentNo.Text = m_InventoryWHDoc.DOCUMENT_NO
      uctlCustomerLookup.MyCombo.ListIndex = IDToListIndex(uctlCustomerLookup.MyCombo, m_InventoryWHDoc.CUSTOMER_ID)
      txtDesc.Text = m_InventoryWHDoc.NOTE

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



Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim IvdWH As CInventoryWHDoc
Dim Sp As CSupItem
Dim LtWh As CLotItemWH
Dim StrStockAmount As String
Dim TempDocNo As String
Dim FirstDate As Date
Dim LastDate As Date
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
   If Not VerifyCombo(lblCustomerNo, uctlCustomerLookup.MyCombo, True) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblTruckNo, txtTruckNo, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_InventoryWHDoc.AddEditMode = ShowMode
   m_InventoryWHDoc.INVENTORY_WH_DOC_ID = ID
   m_InventoryWHDoc.DOCUMENT_DATE = uctlDocumentDate.ShowDate
   m_InventoryWHDoc.TRUCK_NO = txtTruckNo.Text
   m_InventoryWHDoc.DOCUMENT_NO = txtDocumentNo.Text
   m_InventoryWHDoc.CUSTOMER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
   m_InventoryWHDoc.DOCUMENT_TYPE = DocumentType
   m_InventoryWHDoc.NOTE = txtDesc.Text

   m_InventoryWHDoc.NOTE = txtDesc.Text
   
   m_InventoryWHDoc.EXCEPTION_FLAG = "N"
   m_InventoryWHDoc.SUCCESS_FLAG = "N"
   If m_InventoryWHDoc.AddEditMode = SHOW_EDIT Then
      m_InventoryWHDoc.LOAD_FLAG = "Y"
   End If
   
   Call EnableForm(Me, False)
   Call glbDaily.StartTransaction
   
   If Not glbDaily.AddEditInventoryWhDoc(m_InventoryWHDoc, IsOK, False, glbErrorLog) Then
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

Private Sub cboCon1_Click()
m_HasModify = True
End Sub
Private Sub cboCon2_Click()
m_HasModify = True
End Sub
Private Sub cboCon3_Click()
m_HasModify = True
End Sub
Private Sub cboCondition_Click()
m_HasModify = True
End Sub

Private Sub cboDepartment_Click()
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
   If Not VerifyCombo(lblCustomerNo, uctlCustomerLookup.MyCombo) Then
      Exit Sub
   End If
   
   If TabStrip1.SelectedItem.Index = 1 Then
         Set oMenu = New cPopupMenu
         If AutoGenPo Then
            lMenuChosen = oMenu.AddMenu(glbGuiConfigs.LoadGoodsAddMenuItems)
         Else
            lMenuChosen = oMenu.AddMenu(glbGuiConfigs.LoadGoodsAddMenuItems)
         End If
         Set oMenu = Nothing
         If lMenuChosen = 0 Then
            Exit Sub
         End If
   End If
If lMenuChosen = 1 Then
         ShowMode = SHOW_ADD
         Set frmAddSOItem.TempCollection = m_InventoryWHDoc.C_LotItemsWH
         frmAddSOItem.Area = 1 'lMenuChosen
         frmAddSOItem.DocumentNo = Trim(txtDoNo.Text)
         frmAddSOItem.DocumentDate = uctlDocumentDate.ShowDate
         frmAddSOItem.CustomerID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
         frmAddSOItem.TruckNo = Trim(txtTruckNo.Text)
         frmAddSOItem.ShowMode = SHOW_ADD
         frmAddSOItem.HeaderText = MapText("เพิ่มรายการใบขึ้นอาหารจากใบ SALE ORDER")
         
         Load frmAddSOItem
         frmAddSOItem.Show 1
   
         OKClick = frmAddSOItem.OKClick
   
         Unload frmAddSOItem
         Set frmAddSOItem = Nothing
   
         If OKClick Then
            ID = 1
            GridEX1.itemcount = CountItem(m_InventoryWHDoc.C_LotItemsWH)
            GridEX1.Rebind
            
            
            
'            If Not SaveData Then
'               Exit Sub
'            End If
''            ShowMode = SHOW_EDIT
'            ID = m_InventoryWHDoc.INVENTORY_WH_DOC_ID
'            m_InventoryWHDoc.QueryFlag = 1
'            QueryData (True)
'            m_HasModify = False
         End If
ElseIf lMenuChosen = 3 Then
'         frmAddSOItem.AccountID = cboAccount.ItemData(Minus2Zero(cboAccount.ListIndex))txtDoNo
         Set frmAddSOItem.TempCollection = m_InventoryWHDoc.C_LotItemsWH
         frmAddSOItem.Area = lMenuChosen
         frmAddSOItem.DocumentNo = Trim(txtDoNo.Text)
         frmAddSOItem.DocumentDate = uctlDocumentDate.ShowDate
         frmAddSOItem.CustomerID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
         frmAddSOItem.TruckNo = Trim(txtTruckNo.Text)
         frmAddSOItem.ShowMode = SHOW_ADD
         frmAddSOItem.HeaderText = MapText("เพิ่มรายการใบขึ้นอาหารจากใบ INVOICE")
         
         Load frmAddSOItem
         frmAddSOItem.Show 1
   
         OKClick = frmAddSOItem.OKClick
   
         Unload frmAddSOItem
         Set frmAddSOItem = Nothing
   
         If OKClick Then
            GridEX1.itemcount = CountItem(m_InventoryWHDoc.C_LotItemsWH)
            GridEX1.Rebind
         End If
   End If

   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdAuto_Click()
Dim No As String
   If Trim(txtDocumentNo.Text) = "" And ShowMode = SHOW_ADD Then
      Call glbDatabaseMngr.GenerateNumber(LOAD_GOODS, No, glbErrorLog)
      txtDocumentNo.Text = No
   End If
End Sub

'Private Function GetDocumentNo(DocNoType As Long) As String
'Dim No As String
'Dim DOC_ID As Long
'Dim Cd As CConfigDoc
'Dim TempStr As String
'Dim I As Long
'
'   If DocNoType = 2000 Then
'      DOC_ID = WH_LOAD_GOODS
'   End If
'
'    If DOC_ID > 0 Then
'       Set Cd = GetObject("CConfigDoc", m_Cd, Trim(str(DOC_ID)), False)
'       If Not (Cd Is Nothing) Then
'          GetDocumentNo = Cd.GetFieldValue("PREFIX") & Cd.GetFieldValue("CODE1")
'          TempStr = ""
'          If Cd.GetFieldValue("YEAR_TYPE") = 1 Then
'             TempStr = Right(Format(Year(Now) + 543, "0000"), 2)
'          ElseIf Cd.GetFieldValue("YEAR_TYPE") = 2 Then
'             TempStr = Format(Year(Now) + 543, "0000")
'          ElseIf Cd.GetFieldValue("YEAR_TYPE") = 3 Then
'             TempStr = Right(Format(Year(Now), "0000"), 2)
'          ElseIf Cd.GetFieldValue("YEAR_TYPE") = 4 Then
'             TempStr = Format(Year(Now), "0000")
'          End If
'          GetDocumentNo = GetDocumentNo & TempStr & Cd.GetFieldValue("CODE2")
'          TempStr = ""
'          If Cd.GetFieldValue("MONTH_TYPE") = 1 Then
'             TempStr = Format(Month(Now), "00")
'          End If
'          GetDocumentNo = GetDocumentNo & TempStr & Cd.GetFieldValue("CODE3")
'          TempStr = ""
'          For I = 1 To Cd.GetFieldValue("DIGIT_AMOUNT")
'             TempStr = TempStr & "0"
'          Next I
'          GetDocumentNo = GetDocumentNo & Format(Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd, TempStr)
'          m_InventoryWHDoc.RUNNING_NO = Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd
'          m_InventoryWHDoc.CONFIG_DOC_TYPE = DOC_ID
'       Else
'          GetDocumentNo = ""
'       End If
'    End If
'
'End Function

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
         m_InventoryWHDoc.C_LotItemsWH.Remove (ID2)
      Else
         m_InventoryWHDoc.C_LotItemsWH.Item(ID2).Flag = "D"
      End If

'      Call GetTotalPrice
      GridEX1.itemcount = CountItem(m_InventoryWHDoc.C_LotItemsWH)
      GridEX1.Rebind
      m_HasModify = True
   End If
End Sub
Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim itemcount As Long
Dim IsCanLock As Boolean
Dim OKClick As Boolean
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID2 = Val(GridEX1.Value(2))
   OKClick = False
   'ShowMode = SHOW_ADD
   
   If TabStrip1.SelectedItem.Index = 1 Then
      If ShowMode = SHOW_ADD Then
         frmAddEditLoadGoods.ID = ID2
         Set frmAddEditLoadGoods.TempCollection = m_InventoryWHDoc.C_LotItemsWH.Item(ID2).C_LotDoc
         frmAddEditLoadGoods.PART_ITEM_ID = GridEX1.Value(1)
         frmAddEditLoadGoods.PART_NO = GridEX1.Value(3)
         frmAddEditLoadGoods.PART_DESC = GridEX1.Value(4)
         frmAddEditLoadGoods.WEIGHT_PER_PACK = GridEX1.Value(5)
         frmAddEditLoadGoods.PACK_AMOUNT = GridEX1.Value(6)
         frmAddEditLoadGoods.HeaderText = MapText("แก้ไขข้อมูลการโหลดสินค้า")
         frmAddEditLoadGoods.ShowMode = SHOW_ADD
         Load frmAddEditLoadGoods
         frmAddEditLoadGoods.Show 1

         OKClick = frmAddEditLoadGoods.OKClick

         Unload frmAddEditLoadGoods
         Set frmAddEditLoadGoods = Nothing

      If OKClick Then
            GridEX1.itemcount = CountItem(m_InventoryWHDoc.C_LotItemsWH)
            GridEX1.Rebind
      End If
      
      Else
         frmAddEditLoadGoods.ID = ID2
         Set frmAddEditLoadGoods.TempCollection2 = m_InventoryWHDoc.C_LotItemsWH '.Item(ID).C_LotDoc
         frmAddEditLoadGoods.PART_ITEM_ID = GridEX1.Value(1)
         frmAddEditLoadGoods.PART_NO = GridEX1.Value(3)
         frmAddEditLoadGoods.PART_DESC = GridEX1.Value(4)
         frmAddEditLoadGoods.WEIGHT_PER_PACK = GridEX1.Value(10)
         frmAddEditLoadGoods.PACK_AMOUNT = GridEX1.Value(16)
         frmAddEditLoadGoods.HeaderText = MapText("แก้ไขข้อมูลการโหลดสินค้า")
         frmAddEditLoadGoods.ShowMode = SHOW_EDIT
         Load frmAddEditLoadGoods
         frmAddEditLoadGoods.Show 1

         OKClick = frmAddEditLoadGoods.OKClick

         Unload frmAddEditLoadGoods
         Set frmAddEditLoadGoods = Nothing

      If OKClick Then
            GridEX1.itemcount = CountItem(m_InventoryWHDoc.C_LotItemsWH)
            GridEX1.Rebind
      End If
      End If
   End If

   If OKClick Then
      m_HasModify = True
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
      ID = m_InventoryWHDoc.INVENTORY_WH_DOC_ID
      If Not SaveData Then
         Exit Sub
      End If
      
'      ShowMode = SHOW_EDIT
      
      QueryData (False)
      m_HasModify = False
   ElseIf lMenuChosen = 3 Then
      If Not SaveData Then
         Exit Sub
      End If
      
      OKClick = True
      Unload Me
   End If
    Call TabStrip1_Click
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
      glbErrorLog.LocalErrorMsg = "กรุณาทำการบันทึกข้อมูลให้เรียบร้อยก่อน"
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   Set oMenu = New cPopupMenu
   Select Case DocumentType
   Case 2000
      lMenuChosen = oMenu.Popup("ใบรายงานการขึ้นอาหาร", "ปรับค่าหน้ากระดาษ")
   End Select
   If lMenuChosen = 0 Then
      Exit Sub
   End If

   Call EnableForm(Me, False)

   If DocumentType = 2000 Then
      If lMenuChosen = 1 Then
         ReportKey = "CReportLD001"
         Set Report = New CReportLD001
         
'         Call LoadPictureFromFile(glbParameterObj.LoadGoodsPic, Picture1)
         Picture1.Picture = LoadPicture(glbParameterObj.LoadGoodsPic)
         Call Report.AddParam(Picture1.Picture, "BACK_GROUND")
'      Call Report.AddParam(glbParameterObj.LoadGoodsPic, "BACK_GROUND")
      
      
         ReportFlag = True
         Call Report.AddParam(1, "PREVIEW_TYPE")
      End If
   End If

   If Not Report Is Nothing Then

      Call Report.AddParam(m_InventoryWHDoc.INVENTORY_WH_DOC_ID, "INVENTORY_WH_DOC_ID")
      Call Report.AddParam(lMenuChosen, "REPORT_TYPE")
      Call Report.AddParam(ReportKey, "REPORT_KEY")
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
      Call LoadCustomer(uctlCustomerLookup.MyCombo, m_Customers)
      Set uctlCustomerLookup.MyCollection = m_Customers
      

      Call LoadMaster(cboCondition, , CONDITION)
      Call LoadMaster(cboPaidType, , PAID_TYPE)
      Call LoadConfigDoc(Nothing, m_Cd)
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_InventoryWHDoc.QueryFlag = 1 'เอาลูกหลานด้วย
         Call QueryData(True)
         Call TabStrip1_Click
         
      ElseIf ShowMode = SHOW_ADD Then
         '''Call cmdAuto_Click
'         uctlCustomerLookup.SetFocus
         uctlDocumentDate.ShowDate = Now
         m_InventoryWHDoc.QueryFlag = 0
         Call QueryData(False)
      End If
'      Call LoadAuthenPO(m_AuthenPO_Verify, , , Trim(m_InventoryWHDoc.DOCUMENT_NO), Trim(m_InventoryWHDoc.DOCUMENT_TYPE))
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
   
   Set m_InventoryWHDoc = Nothing
   Set m_Customers = Nothing
   Set m_Cd = Nothing
   Set TempWeight = Nothing
   Set m_Weight = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid1()
Dim Col As JSColumn
Dim i As Long

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.itemcount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.NAME = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 10
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 600
   Col.Caption = "ลำดับ"
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 3500
   Col.Caption = MapText("เบอร์สินค้า")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 800
   Col.Caption = MapText("ชนิดสินค้า")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 1200
   Col.Caption = MapText("Lot. No.")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 1300
   Col.Caption = MapText("วันที่ผลิต")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 1000
   Col.Caption = MapText("จำนวนแบท (B)")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 1300
   Col.Caption = MapText("บรรจุดี")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 1300
   Col.Caption = MapText("บรรจุเสีย")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 1300
   Col.Caption = MapText("ขนาด")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 1300
   Col.Caption = MapText("น้ำหนักรวม")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 1300
   Col.Caption = MapText("จำนวนเศษ (กก.)")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 1300
   Col.Caption = MapText("Bin-No.")
   
      Set Col = GridEX1.Columns.add '3
   Col.Width = 1300
   Col.Caption = MapText("เวลาเริ่มบรรจุ")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 1500
   Col.Caption = MapText("เวลาบรรจุ")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 1500
   Col.Caption = MapText("จำนวนบรรจุ (ถุง)")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 1500
   Col.Caption = MapText("พาเลทที่วาง")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2000
   Col.Caption = MapText("หมายเหตุ")

End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame4.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblDocumentNo, MapText("เลขที่ใบขึ้นอาหาร"))
   Call InitNormalLabel(lblTruckNo, MapText("ทะเบียนรถ"))
   Call InitNormalLabel(lblDesc, MapText("หมายเหตุ"))
   Call InitNormalLabel(lblDoNo, MapText("เลขที่ INV"))
   
   
   uctlCustomerLookup.MyTextBox.SetKeySearch ("CUSTOMER_CODE")
   Call InitNormalLabel(lblDocumentDate, MapText("วันที่เอกสาร"))
   Call InitNormalLabel(lblCustomerNo, MapText("รหัสลูกค้า"))
   
   Call txtDocumentNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtDoNo.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtTruckNo.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtDesc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   
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
   cmdPrint.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOther.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdPrint, MapText("พิมพ์"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdOther, MapText("อื่นๆ"))
   
   Call InitGrid1
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.NAME = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   Dim str As String
   Select Case DocumentType
    Case 2000
          str = "รายการสินค้า"
   End Select
   TabStrip1.Tabs.add().Caption = MapText(str)
'   If DocumentType = 1000 Or DocumentType = 1001 Or DocumentType = 1002 Or DocumentType = 1003 Then
'      TabStrip1.Tabs.add().Caption = MapText("รายละเอียดทั่วไป")
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
'   OKClick = False
   Call InitFormLayout
      
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_InventoryWHDoc = New CInventoryWHDoc
   Set m_Weight = New CWeight
   Set m_Customers = New Collection
   Set m_Cd = New Collection
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
Dim i As Long
Dim CountItem As Long

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Index = 1 Then
      If m_InventoryWHDoc.C_LotItemsWH Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If
      
      Dim CR As CLotItemWH
      Dim LTD As CLotDoc
      Dim PD As CPalletDoc
      If m_InventoryWHDoc.C_LotItemsWH.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_InventoryWHDoc.C_LotItemsWH, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If
      
      Values(1) = CR.PART_ITEM_ID
      Values(2) = RealIndex
      Values(3) = CR.PART_NO
      Values(4) = CR.PRODUCT_TYPE_NAME
      Values(5) = CR.LOT_NO
      Values(6) = CR.START_DATE
      Values(7) = ""
      Values(8) = CR.GOOD_AMOUNT
      Values(9) = CR.LOSE_AMOUNT
      Values(10) = CR.WEIGHT_PER_PACK
      Values(11) = CR.TX_AMOUNT
      Values(12) = CR.REST_AMOUNT
      Values(13) = CR.BIN_NAME
      Values(14) = CR.TIME_PACK_BEGIN
      Values(15) = CR.TIME_PACK_END
      Values(16) = CR.PACK_AMOUNT
      Values(17) = CR.FULL_PALLET_FROM & "-" & CR.SCRAP_PALLET
      Values(18) = CR.NOTE
       

'      If Not (CR.C_LotDoc Is Nothing) Then
'         I = 10
'            For Each LTD In CR.C_LotDoc
'              Values(7) = LTD.START_DATE
'              Values(8) = ConcateSting(Values(8), LTD.LOT_NO)
'              Values(9) = CR.LOCK_NAME
'              If Not (LTD.C_PalletDoc Is Nothing) Then
'               For Each PD In LTD.C_PalletDoc
'
'                  Values(I) = PD.CAPACITY_AMOUNT
'                  I = I + 1
'               Next PD
'               End If
'            Next LTD
'         End If
   End If
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub



Private Sub TabStrip1_Click()
   GridEX1.Visible = False
   SSFrame4.Visible = False
   If TabStrip1.SelectedItem.Index = 1 Then
     Call EnableDisableButton(True)
      GridEX1.Visible = True
      Call InitGrid1
      GridEX1.itemcount = CountItem(m_InventoryWHDoc.C_LotItemsWH)
      GridEX1.Rebind
   End If
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

Private Sub txtEntryWeight_Change()
   m_HasModify = True
End Sub

Private Sub txtExitWeight_Change()
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

Private Sub txtWeightAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtWeightNote_Change()
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

Private Sub uctlCustomerLookup_Change()
m_HasModify = True
'Dim ID As Long
'Dim Sp As CSupplier
'
'
'   ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
'   If ID > 0 Then
'      Set Sp = GetSupplier(m_Customers, Trim(str(ID)))
'   End If
'
'   m_HasModify = True

End Sub
Private Sub PopulateGuiID(Bd As CBillingDoc)
Dim Di As CSupItem

   For Each Di In Bd.SupItems
      If Di.Flag = "A" Then
         Di.LINK_ID = GetNextGuiID(Bd)
      End If
   Next Di
End Sub

Private Function GetNextGuiID(Bd As CBillingDoc) As Long
Dim Di As CSupItem
Dim MaxId As Long

   MaxId = 0
   For Each Di In Bd.SupItems
      If Di.LINK_ID > MaxId Then
         MaxId = Di.LINK_ID
      End If
   Next Di
   
   GetNextGuiID = MaxId + 1
End Function
Private Sub EnableDisableButton(En As Boolean)
   If En Then
      If ShowMode = SHOW_EDIT Then
         cmdAdd.Enabled = (m_InventoryWHDoc.OLD_COMMIT_FLAG = "N")
         cmdEdit.Enabled = True '(m_InventoryWHDoc.COMMIT_FLAG = "N")
         cmdDelete.Enabled = (m_InventoryWHDoc.OLD_COMMIT_FLAG = "N")
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
