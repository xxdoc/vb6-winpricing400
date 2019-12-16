VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmInventoryDocWH 
   BackColor       =   &H80000000&
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   11925
   Icon            =   "frmInventoryDocWH.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   11925
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlDate uctlDocumentDate 
         Height          =   405
         Left            =   6180
         TabIndex        =   1
         Top             =   750
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin VB.ComboBox cboOrderType 
         Height          =   315
         Left            =   6180
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2160
         Width           =   2625
      End
      Begin VB.ComboBox cboOrderBy 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2160
         Width           =   2985
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtDocumentNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   720
         Width           =   2985
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtCustomerCode 
         Height          =   435
         Left            =   1860
         TabIndex        =   2
         Top             =   1170
         Width           =   2985
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlToDate 
         Height          =   405
         Left            =   6180
         TabIndex        =   20
         Top             =   1200
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   4485
         Left            =   120
         TabIndex        =   23
         Top             =   3120
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   7911
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
         Column(1)       =   "frmInventoryDocWH.frx":27A2
         Column(2)       =   "frmInventoryDocWH.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmInventoryDocWH.frx":290E
         FormatStyle(2)  =   "frmInventoryDocWH.frx":2A6A
         FormatStyle(3)  =   "frmInventoryDocWH.frx":2B1A
         FormatStyle(4)  =   "frmInventoryDocWH.frx":2BCE
         FormatStyle(5)  =   "frmInventoryDocWH.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmInventoryDocWH.frx":2D5E
      End
      Begin prjFarmManagement.uctlTextBox txtTruckNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   24
         Top             =   1620
         Width           =   2985
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin VB.Label lblTruckNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   0
         TabIndex        =   25
         Top             =   1680
         Width           =   1755
      End
      Begin Threed.SSCommand cmdOther 
         Height          =   525
         Left            =   6800
         TabIndex        =   10
         Top             =   7830
         Visible         =   0   'False
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdjust 
         Height          =   525
         Left            =   5040
         TabIndex        =   22
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4920
         TabIndex        =   21
         Top             =   1230
         Width           =   1185
      End
      Begin VB.Label lblDocumentNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   19
         Top             =   780
         Width           =   1755
      End
      Begin VB.Label lblCustomerCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   18
         Top             =   1230
         Width           =   1755
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4980
         TabIndex        =   17
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4920
         TabIndex        =   16
         Top             =   780
         Width           =   1185
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   120
         TabIndex        =   15
         Top             =   2280
         Width           =   1755
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   10110
         TabIndex        =   5
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmInventoryDocWH.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   10080
         TabIndex        =   6
         Top             =   1680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   9
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmInventoryDocWH.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   7
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmInventoryDocWH.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   8
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10095
         TabIndex        =   12
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8445
         TabIndex        =   11
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmInventoryDocWH.frx":3884
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmInventoryDocWH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_InventoryWHDoc As CInventoryWHDoc
Private m_TempInventoryWHDoc As CInventoryWHDoc
Private m_Rs As ADODB.Recordset
Private m_TableName As String
Private m_IvdDocType As Long

Public OKClick As Boolean
Public DocumentType As Long
Public DocumentTypeName As String
Public ReceiptType As Long
Public Area As Long
Public DoReceiptFlag As String
Public HeaderText As String
Dim FromDate As Date
Dim ToDate As Date

Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu
Dim TempStr As String
Dim Programowner As String

   Programowner = glbParameterObj.Programowner
   
   If (DocumentType = 2000 Or DocumentType = 2001 Or DocumentType = 2004) Then
      If Not VerifyAccessRight("INVENTORY-WH_EXPORT" & "_" & DocumentType & "_" & "ADD", "เพิ่ม") Then
          Call EnableForm(Me, True)
          Exit Sub
       End If
      
      frmAddEditInventoryDocWh.DocumentType = DocumentType
      frmAddEditInventoryDocWh.HeaderText = MapText("เพิ่มข้อมูล" & HeaderText)
      frmAddEditInventoryDocWh.ShowMode = SHOW_ADD
      Load frmAddEditInventoryDocWh
      frmAddEditInventoryDocWh.Show 1
      
      OKClick = frmAddEditInventoryDocWh.OKClick
      
      Unload frmAddEditInventoryDocWh
      Set frmAddEditInventoryDocWh = Nothing
   ElseIf (DocumentType = 20 Or DocumentType = 21) Then
      If Not VerifyAccessRight("INVENTORY-WH_TRANSFER" & "_" & DocumentTypeName & "_" & "ADD", MapText("เพิ่มข้อมูล" & HeaderText)) Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub
Private Sub cmdClear_Click()
   txtDocumentNo.Text = ""
   txtCustomerCode.Text = ""
   txtTruckNo.Text = ""
   uctlDocumentDate.ShowDate = -1
   uctlToDate.ShowDate = -1
   cboOrderBy.ListIndex = -1
   cboOrderType.ListIndex = -1
End Sub

Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim id As Long
Dim m_InventoryWHDoc2 As CInventoryWHDoc
Dim m_LotItemWh As CLotItemWH
 

   If DocumentType = 2000 Or DocumentType = 2001 Or DocumentType = 2004 Then
       If Not VerifyAccessRight("INVENTORY-WH_EXPORT" & "_" & DocumentType & "_" & "DELETE", "ลบ") Then
           Call EnableForm(Me, True)
           Exit Sub
        End If
  ElseIf DocumentType = 19 Or DocumentType = 20 Then
        If Not VerifyAccessRight("INVENTORY-WH_TRANSFER" & "_" & DocumentTypeName & "_" & "DELETE", MapText("ลบข้อมูล" & HeaderText)) Then
          Call EnableForm(Me, True)
          Exit Sub
       End If
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   id = Val(GridEX1.Value(1))

   Call glbDatabaseMngr.LockTable(m_TableName, id, IsCanLock, glbErrorLog)
   If Not ConfirmDelete(GridEX1.Value(2)) Then
      Call glbDatabaseMngr.UnLockTable(m_TableName, id, IsCanLock, glbErrorLog)
      Exit Sub
   End If
   Call EnableForm(Me, False)
   
   'ดึงข้อมูลก่อนลบ
        Set m_InventoryWHDoc2 = New CInventoryWHDoc
      m_InventoryWHDoc2.INVENTORY_WH_DOC_ID = id
      m_InventoryWHDoc2.COMMIT_FLAG = ""
      m_InventoryWHDoc2.QueryFlag = 1

      If Not glbDaily.QueryInventoryWhDocForLG(m_InventoryWHDoc2, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
        Exit Sub
      End If
      '***********************
   
   If Not glbDaily.DeleteInventoryWhDoc(id, IsOK, True, glbErrorLog) Then
      m_InventoryWHDoc.INVENTORY_WH_DOC_ID = -1
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call glbDatabaseMngr.UnLockTable(m_TableName, id, IsCanLock, glbErrorLog)
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
      'Update สถานะเอกสาร ของ BillingDoc หาก ดึงใบ so แล้วก็ให้เปลี่ยนเป็น Y
   Dim BD As CBillingDoc
   For Each m_LotItemWh In m_InventoryWHDoc2.C_LotItemsWH
           Set BD = New CBillingDoc
           BD.BILLING_DOC_ID = m_LotItemWh.BILLING_DOC_ID
           Call BD.UpdateSuccessFlag("N")
   Next m_LotItemWh

   'ตรวจสอบ Stock หลังจาก update ใหม่ เพื่อเปลี่ยนสถานะ ของ Out Stock Flag
   For Each m_LotItemWh In m_InventoryWHDoc2.C_LotItemsWH
         Call LoadLotInPartIemAmount(Nothing, Nothing, , , , , m_LotItemWh.PART_ITEM_ID, 2, 1, 1, "I", m_LotItemWh.C_LotDoc, , DocumentType, m_LotItemWh.Flag)
   Next m_LotItemWh
   
   
   Call QueryData(True)
   Call glbDatabaseMngr.UnLockTable(m_TableName, id, IsCanLock, glbErrorLog)
   Call EnableForm(Me, True)
End Sub
Private Sub QueryData2(ID2 As Long, IWD As CInventoryWHDoc)
Dim IsOK As Boolean
Dim ItemCount As Long

      IWD.INVENTORY_WH_DOC_ID = ID2
      IWD.COMMIT_FLAG = ""
      IWD.QueryFlag = 1
      
      If Not glbDaily.QueryInventoryWhDocForLG(IWD, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
        Exit Sub
      End If

End Sub
Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim id As Long
Dim OKClick As Boolean
Dim TempStr As String

   Dim Programowner As String
   Programowner = glbParameterObj.Programowner
      

   Call glbDatabaseMngr.LockTable(m_TableName, id, IsCanLock, glbErrorLog)
   
   If DocumentType = 2000 Or DocumentType = 2001 Or DocumentType = 2004 Then
      If Not VerifyAccessRight("INVENTORY-WH_EXPORT" & "_" & DocumentType & "_" & "EDIT", "แก้ไข") Then
          Call EnableForm(Me, True)
          Exit Sub
       End If
       
         If Not VerifyGrid(GridEX1.Value(1)) Then
            Exit Sub
         End If
         
         id = Val(GridEX1.Value(1))
   
      frmAddEditInventoryDocWh.id = id
      frmAddEditInventoryDocWh.DocumentType = DocumentType
      frmAddEditInventoryDocWh.HeaderText = MapText("แก้ไขข้อมูล" & HeaderText)
      frmAddEditInventoryDocWh.ShowMode = SHOW_EDIT
      Load frmAddEditInventoryDocWh
      frmAddEditInventoryDocWh.Show 1
      
      OKClick = frmAddEditInventoryDocWh.OKClick
      
      Unload frmAddEditInventoryDocWh
      Set frmAddEditInventoryDocWh = Nothing
   ElseIf DocumentType = 19 Or DocumentType = 20 Then
        If Not VerifyAccessRight("INVENTORY-WH_TRANSFER" & "_" & DocumentTypeName & "_" & "EDIT", MapText("แก้ไขข้อมูล" & HeaderText)) Then
          Call EnableForm(Me, True)
          Exit Sub
       End If
       
         If Not VerifyGrid(GridEX1.Value(1)) Then
            Exit Sub
         End If
         
         id = Val(GridEX1.Value(1))
   End If
   
   If OKClick Then
      Call QueryData(True)
   End If
   Call glbDatabaseMngr.UnLockTable(m_TableName, id, IsCanLock, glbErrorLog)

End Sub

Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub

Private Sub cmdOther_Click()
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu

   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("เปิดใบรับของโดยไม่มี PO", "-", "อื่นๆ")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
   If lMenuChosen = 1 Then
      If DocumentType = 100 Or DocumentType = 101 Or DocumentType = 102 Or DocumentType = 103 Then
         If Not VerifyAccessRight("LEDGER_STOCKBUY" & "_" & DocumentType & "_" & "NO-PO", "ไม่มีPO") Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
         
         frmAddEditBillingSup.AutoGenPo = True
         frmAddEditBillingSup.DocumentType = DocumentType
         frmAddEditBillingSup.HeaderText = MapText("เพิ่มข้อมูลการนำเข้า โดยไม่มี PO")
         frmAddEditBillingSup.ShowMode = SHOW_ADD
         Load frmAddEditBillingSup
         frmAddEditBillingSup.Show 1
         
         OKClick = frmAddEditBillingSup.OKClick
         
         Unload frmAddEditBillingSup
         Set frmAddEditBillingSup = Nothing
         
         
      End If
   ElseIf lMenuChosen = 3 Then
   
   End If
      
   If OKClick Then
      Call QueryData(True)
   End If
End Sub

Private Sub cmdSearch_Click()
   Call LoadLoginTracking(Nothing, m_LoginTracking, DateAdd("M", -1, uctlDocumentDate.ShowDate), uctlToDate.ShowDate)
   Call QueryData(True)
End Sub

Private Sub Form_Activate()

   If Not m_HasActivate Then
      m_HasActivate = True
      'DocumentType = 2000 Or DocumentType = 2001 Or DocumentType = 2004
      If DocumentType = 20 Or DocumentType = 2000 Then
         DocumentTypeName = "BAG"
      ElseIf DocumentType = 21 Or DocumentType = 2001 Then
         DocumentTypeName = "BULK"
      ElseIf DocumentType = 2004 Then
         DocumentTypeName = "OTHER"
      End If
      
      Call InitBillingDocOrderBy(cboOrderBy)
      Call InitOrderType(cboOrderType)
      
      Call GetFirstLastDate(Now, FromDate, ToDate)
      uctlDocumentDate.ShowDate = FromDate
      uctlToDate.ShowDate = ToDate
      
      'Call LoadLoginTracking(Nothing, m_LoginTracking, DateAdd("M", -1, FromDate), ToDate)

      Call QueryData(True)
   End If
End Sub

Private Function GetPermissionCode() As String
Dim TempStr As String

   If Area = 1 Then
      If DocumentType = 1 Then
         TempStr = "LEDGER_DO"
      ElseIf DocumentType = 2 Then
         TempStr = "LEDGER_RC"
      ElseIf DocumentType = 3 Then
         TempStr = "LEDGER_CN"
      ElseIf DocumentType = 4 Then
         TempStr = "LEDGER_DN"
      ElseIf DocumentType = 18 Then
         TempStr = "LEDGER_RT"
      ElseIf DocumentType = 19 Then
         TempStr = "LEDGER_SO"
      End If
   End If
   
   GetPermissionCode = TempStr
End Function

Public Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      m_InventoryWHDoc.INVENTORY_WH_DOC_ID = -1
      m_InventoryWHDoc.DOCUMENT_TYPE = DocumentType
      m_InventoryWHDoc.DOCUMENT_NO = PatchWildCard(txtDocumentNo.Text)
      m_InventoryWHDoc.CUSTOMER_CODE = PatchWildCard(txtCustomerCode.Text)
      m_InventoryWHDoc.TRUCK_NO = PatchWildCard(txtTruckNo.Text)
      m_InventoryWHDoc.FROM_DATE = uctlDocumentDate.ShowDate
      m_InventoryWHDoc.TO_DATE = uctlToDate.ShowDate
      m_InventoryWHDoc.OrderBy = cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex))
      m_InventoryWHDoc.OrderType = cboOrderType.ItemData(Minus2Zero(cboOrderType.ListIndex))
      m_InventoryWHDoc.QueryFlag = 0 'ถ้าเป็น 0 ยังไม่ต้องเอาลูก หลาน แต่ถ้าเป็น 1 เอา ลูกหลานด้วย
     
      If Not glbDaily.QueryInventoryWhDoc(m_InventoryWHDoc, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
'      cmdDelete.Enabled = (m_InventoryWHDoc.COMMIT_FLAG = "N")
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   GridEX1.ItemCount = ItemCount
   GridEX1.Rebind
   
   Call EnableForm(Me, True)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.NAME
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 116 Then
      Call cmdSearch_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
      Call cmdClear_Click
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
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 123 Then
      Call AddMemoNote
      KeyCode = 0
   End If
End Sub

Private Sub InitGrid()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   
   GridEX1.FormatStyles.Clear
   Set fmsTemp = GridEX1.FormatStyles.add("N")
   fmsTemp.ForeColor = GLB_ALERT_COLOR
   
   If DocumentType = 2000 Or DocumentType = 2001 Or DocumentType = 2004 Then
      Set Col = GridEX1.Columns.add '1
      Col.Width = 0
      Col.Caption = "ID"
   
      Set Col = GridEX1.Columns.add '2
      Col.Width = 1800 '2115
      Col.Caption = MapText("เลขที่เอกสาร")
         
      Set Col = GridEX1.Columns.add '3
      Col.Width = 1500 '2055
      Col.Caption = MapText("วันที่เอกสาร")
      
      Set Col = GridEX1.Columns.add '4
      Col.Width = 2305
      Col.Caption = MapText("รหัสลูกค้า")
      
      Set Col = GridEX1.Columns.add '5
      Col.Width = 4995
      Col.Caption = MapText("ชื่อลูกค้า")
      
      Set Col = GridEX1.Columns.add '6
      Col.Width = 4995
      Col.Caption = MapText("ทะเบียนรถ")
      
      Set Col = GridEX1.Columns.add '7
      Col.Width = 2700
      Col.Caption = MapText("สถานะเอกสาร")
      
      Set Col = GridEX1.Columns.add '8
      Col.Width = 0
      Col.Caption = MapText("สถานะเอกสาร Flag")
      
      Set Col = GridEX1.Columns.add '9
      Col.Width = 2000
      Col.Caption = MapText("ใบฝากขาย")
      
      Set Col = GridEX1.Columns.add '10
      Col.Width = 1200
      Col.Caption = MapText("สร้าง")
      
      Set Col = GridEX1.Columns.add '11
      Col.Width = 1200
      Col.Caption = MapText("แก้ไข")
      
   ElseIf DocumentType = 20 Or DocumentType = 21 Then
      Set Col = GridEX1.Columns.add '1
      Col.Width = 0
      Col.Caption = "ID"
      
      Set Col = GridEX1.Columns.add '2
      Col.Width = 1800
      Col.Caption = MapText("เลขที่เอกสาร")
      
      Set Col = GridEX1.Columns.add '3
      Col.Width = 1500
      Col.Caption = MapText("วันที่เอกสาร")
      
'      Set Col = GridEX1.Columns.add '4
'      Col.Width = 1500
'      Col.Caption = MapText("วันที่เริ่มผลิต")
'
'      Set Col = GridEX1.Columns.add '5
'      Col.Width = 1500
'      Col.Caption = MapText("วันที่ผลิตเสร็จ")
'
'      Set Col = GridEX1.Columns.add '6
'      Col.Width = 1500
'      Col.Caption = MapText("จำนวนแบท")
'
'      Set Col = GridEX1.Columns.add '7
'      Col.Width = 1500
'      Col.Caption = MapText("จำนวนรับเข้า")
      
      Set Col = GridEX1.Columns.add '8
      Col.Width = 1200
      Col.Caption = MapText("สร้าง")
      
      Set Col = GridEX1.Columns.add '9
      Col.Width = 1200
      Col.Caption = MapText("แก้ไข")
      
   End If
   GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
Dim Programowner As String
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Programowner = glbParameterObj.Programowner
   Me.Caption = MapText(HeaderText)
   pnlHeader.Caption = Me.Caption
      
   Call InitGrid
   
   Call InitNormalLabel(lblDocumentDate, MapText("จากวันที่"))
   Call InitNormalLabel(lblToDate, MapText("ถึงวันที่"))
   Call InitNormalLabel(lblCustomerCode, MapText("รหัสลูกค้า"))
   Call InitNormalLabel(lblTruckNo, MapText("ทะเบียนรถ"))
   Call InitNormalLabel(lblDocumentNo, MapText("เลขที่เอกสาร"))
   Call InitNormalLabel(lblOrderBy, MapText("เรียงตาม"))
   Call InitNormalLabel(lblOrderType, MapText("เรียงจาก"))
   Call txtCustomerCode.SetKeySearch("CUSTOMER_CODE")
   
   Call InitCombo(cboOrderBy)
   Call InitCombo(cboOrderType)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdClear.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
'   cmdAdjust.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOther.Picture = LoadPicture(glbParameterObj.NormalButton1)
'   cmdDelete.Enabled = False
   
   'Call InitMainButton(cmdAdjust, MapText("ปรับยอด"))
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา (F5)"))
   Call InitMainButton(cmdClear, MapText("เคลียร์ (F4)"))
   
   Call InitMainButton(cmdOther, MapText("อื่นๆ"))
   
  pnlHeader.Caption = Me.Caption
End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   m_TableName = "USER_GROUP"
   
   Set m_InventoryWHDoc = New CInventoryWHDoc
   Set m_TempInventoryWHDoc = New CInventoryWHDoc
   Set m_Rs = New ADODB.Recordset
   
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_InventoryWHDoc = Nothing
   Set m_TempInventoryWHDoc = Nothing
   Set m_Rs = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub GridEX1_DblClick()
    Call cmdEdit_Click
End Sub

Private Sub GridEX1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim oMenu As cPopupMenu
'Dim lMenuChosen As Long
'Dim TempID1 As Long
'Dim BD As CBillingDoc
'Dim IsOK As Boolean
'Dim OKClick As Boolean
'
'   If GridEX1.ItemCount <= 0 Then
'         Exit Sub
'   End If
'
'   TempID1 = GridEX1.Value(1)
'   If Button = 2 Then
'      Set oMenu = New cPopupMenu
'     lMenuChosen = oMenu.Popup("คัดลอกข้อมูล")
'      If lMenuChosen = 0 Then
'         Exit Sub
'      End If
'      Set oMenu = Nothing
'   Else
'      Exit Sub
'   End If
'
'   Call EnableForm(Me, False)
'   If lMenuChosen = 1 Then
'      If Not (Area = 1) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If
'      Set BD = New CBillingDoc
'      BD.BILLING_DOC_ID = TempID1
'      Call glbDaily.CopyBillingDoc(BD, IsOK, True, Area, m_IvdDocType, glbErrorLog)
'      Call QueryData(True)
'      Set BD = Nothing
'   End If
'
'   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
'   RowBuffer.RowStyle = RowBuffer.Value(6)
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long
Dim fmsTemp As JSFormatStyle

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"

   If m_Rs Is Nothing Then
      Exit Sub
   End If

   If m_Rs.State <> adStateOpen Then
      Exit Sub
   End If

   If m_Rs.EOF Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If
   Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
   Call m_TempInventoryWHDoc.PopulateFromRS(1, m_Rs)
  If DocumentType = 2000 Or DocumentType = 2001 Or DocumentType = 2004 Then
      Values(1) = m_TempInventoryWHDoc.INVENTORY_WH_DOC_ID
      Values(2) = m_TempInventoryWHDoc.DOCUMENT_NO
      Values(3) = DateToStringExtEx2(m_TempInventoryWHDoc.DOCUMENT_DATE)
      Values(4) = m_TempInventoryWHDoc.CUSTOMER_CODE
      Values(5) = m_TempInventoryWHDoc.CUSTOMER_NAME
      Values(6) = m_TempInventoryWHDoc.TRUCK_NO
      Values(7) = ConvertLoadFlag(Trim(m_TempInventoryWHDoc.LOAD_FLAG))
      Values(8) = m_TempInventoryWHDoc.LOAD_FLAG
      Values(9) = m_TempInventoryWHDoc.DOCUMENT_NO_RQ
        Set Temp_LTK = GetObject("CLoginTracking", m_LoginTracking, Trim(m_TempInventoryWHDoc.CREATE_BY), False)
        If Not Temp_LTK Is Nothing Then
            Values(10) = Temp_LTK.USER_NAME
         Else
             Values(10) = ""
         End If
         
         Set Temp_LTK = GetObject("CLoginTracking", m_LoginTracking, Trim(m_TempInventoryWHDoc.MODIFY_BY), False)
        If Not Temp_LTK Is Nothing Then
            Values(11) = Temp_LTK.USER_NAME
         Else
             Values(11) = ""
         End If
   ElseIf DocumentType = 20 Or DocumentType = 21 Then
      Values(1) = m_TempInventoryWHDoc.INVENTORY_WH_DOC_ID
      Values(2) = m_TempInventoryWHDoc.DOCUMENT_NO
      Values(3) = DateToStringExtEx2(m_TempInventoryWHDoc.DOCUMENT_DATE)
'      Values(4) = m_TempInventoryWHDoc.CUSTOMER_CODE
'      Values(5) = m_TempInventoryWHDoc.CUSTOMER_NAME
'      Values(6) = m_TempInventoryWHDoc.TRUCK_NO
'      Values(7) = ConvertLoadFlag(Trim(m_TempInventoryWHDoc.LOAD_FLAG))
'      Values(8) = m_TempInventoryWHDoc.LOAD_FLAG
      
        Set Temp_LTK = GetObject("CLoginTracking", m_LoginTracking, Trim(m_TempInventoryWHDoc.CREATE_BY), False)
        If Not Temp_LTK Is Nothing Then
            Values(4) = Temp_LTK.USER_NAME
         Else
             Values(4) = ""
         End If
         
         Set Temp_LTK = GetObject("CLoginTracking", m_LoginTracking, Trim(m_TempInventoryWHDoc.MODIFY_BY), False)
        If Not Temp_LTK Is Nothing Then
            Values(5) = Temp_LTK.USER_NAME
         Else
             Values(5) = ""
         End If
   End If
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.HEIGHT = ScaleHeight
   
   pnlHeader.Width = ScaleWidth
   GridEX1.Width = ScaleWidth - 2 * GridEX1.Left
   GridEX1.HEIGHT = ScaleHeight - GridEX1.Top - 620
   cmdAdd.Top = ScaleHeight - 580
   cmdEdit.Top = ScaleHeight - 580
   cmdDelete.Top = ScaleHeight - 580
   cmdOK.Top = ScaleHeight - 580
   cmdExit.Top = ScaleHeight - 580
   cmdOther.Top = ScaleHeight - 580
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
   cmdOther.Left = cmdOK.Left - cmdOther.Width - 50
End Sub
