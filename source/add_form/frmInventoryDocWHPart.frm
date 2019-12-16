VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmInventoryDocWHPart 
   BackColor       =   &H80000000&
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   11910
   Icon            =   "frmInventoryDocWHPart.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   9255
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   13515
      _ExtentX        =   23839
      _ExtentY        =   16325
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextLookup uctlPlaceLookup 
         Height          =   375
         Left            =   1860
         TabIndex        =   22
         Top             =   2040
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   661
      End
      Begin VB.ComboBox cboPartType 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1560
         Width           =   2985
      End
      Begin VB.ComboBox cboOrderType 
         Height          =   315
         Left            =   8010
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2040
         Width           =   2955
      End
      Begin VB.ComboBox cboOrderBy 
         Height          =   315
         Left            =   8010
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1560
         Width           =   2955
      End
      Begin prjFarmManagement.uctlTextBox txtPartName 
         Height          =   435
         Left            =   8010
         TabIndex        =   1
         Top             =   1080
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   13605
         _ExtentX        =   23998
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   5055
         Left            =   180
         TabIndex        =   7
         Top             =   2640
         Width           =   13065
         _ExtentX        =   23045
         _ExtentY        =   8916
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
         Column(1)       =   "frmInventoryDocWHPart.frx":27A2
         Column(2)       =   "frmInventoryDocWHPart.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmInventoryDocWHPart.frx":290E
         FormatStyle(2)  =   "frmInventoryDocWHPart.frx":2A6A
         FormatStyle(3)  =   "frmInventoryDocWHPart.frx":2B1A
         FormatStyle(4)  =   "frmInventoryDocWHPart.frx":2BCE
         FormatStyle(5)  =   "frmInventoryDocWHPart.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmInventoryDocWHPart.frx":2D5E
      End
      Begin prjFarmManagement.uctlTextBox txtPartNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   1080
         Width           =   2985
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin VB.Label lblPlaceLookup 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   0
         TabIndex        =   21
         Top             =   2100
         Width           =   1755
      End
      Begin Threed.SSCommand cmdAdjust 
         Height          =   525
         Left            =   7920
         TabIndex        =   20
         Top             =   7830
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblPartNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   19
         Top             =   1140
         Width           =   1755
      End
      Begin VB.Label lblPartType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   18
         Top             =   1590
         Width           =   1755
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   6120
         TabIndex        =   17
         Top             =   2100
         Width           =   1755
      End
      Begin VB.Label lblPartName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   6480
         TabIndex        =   16
         Top             =   1140
         Width           =   1455
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   6150
         TabIndex        =   15
         Top             =   1590
         Width           =   1755
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   11670
         TabIndex        =   5
         Top             =   1170
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmInventoryDocWHPart.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   11670
         TabIndex        =   6
         Top             =   1740
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   10
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmInventoryDocWHPart.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   8
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmInventoryDocWHPart.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   9
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
         Left            =   11655
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
         Left            =   10005
         TabIndex        =   11
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmInventoryDocWHPart.frx":3884
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmInventoryDocWHPart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_PartItem As CLotItemWH
Private m_TempPartItem As CLotItemWH
Private m_Rs As ADODB.Recordset
Private m_TableName As String
Public DocumentType As Long
Public HeaderText As String
Public PartGroupID As Long
Public OKClick As Boolean
Public Area As Long
Public PartType As Long
Public m_Locations As Collection



Private Sub cboPartType_Click()
 If cboPartType.ListIndex = 1 Then
   uctlPlaceLookup.MyCombo.ListIndex = IDToListIndex(uctlPlaceLookup.MyCombo, 109)
 ElseIf cboPartType.ListIndex = 2 Then
  uctlPlaceLookup.MyCombo.ListIndex = IDToListIndex(uctlPlaceLookup.MyCombo, 110)
 Else
  uctlPlaceLookup.MyCombo.ListIndex = IDToListIndex(uctlPlaceLookup.MyCombo, 0)
 End If
End Sub

Private Sub cmdAdd_Click()
Dim cPopup As cPopupMenu
Dim lMenuChosen As Long
   If Not VerifyAccessRight("PRODUCT_JOB_EDIT") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   Set cPopup = New cPopupMenu
   lMenuChosen = cPopup.Popup("นำเข้าข้อมูลคงเหลือสินค้าสำเร็จรูปจากคลัง")
   Set cPopup = Nothing
   
   If lMenuChosen = 1 Then
      frmImportPacking.Area = 2
      Load frmImportPacking
      frmImportPacking.Show 1
      
      Unload frmImportPacking
      Set frmImportPacking = Nothing
   End If
End Sub

Private Sub cmdAdjust_Click()
   Load frmAdjustInventoryWH
   frmAdjustInventoryWH.Show 1
   
   OKClick = frmAdjustInventoryWH.OKClick
   
    txtPartNo.Text = frmAdjustInventoryWH.PartNo
   If frmAdjustInventoryWH.PartType = "10" Then
      cboPartType.ListIndex = 1
   ElseIf frmAdjustInventoryWH.PartType = "22" Then
      cboPartType.ListIndex = 2
   End If
   
   Unload frmAdjustInventoryWH
   Set frmAdjustInventoryWH = Nothing
   
  Call cmdSearch_Click
End Sub

Private Sub cmdClear_Click()
   txtPartName.Text = ""
   txtPartNo.Text = ""
   cboOrderBy.ListIndex = -1
   cboOrderType.ListIndex = -1
   cboPartType.ListIndex = -1
'   cboUnit.ListIndex = -1
End Sub

Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim id As Long
   If Not VerifyAccessRight("INVENTORY_PART_DELETE") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   id = GridEX1.Value(1)
   
   Call glbDatabaseMngr.LockTable(m_TableName, id, IsCanLock, glbErrorLog)
   If Not ConfirmDelete(GridEX1.Value(2)) Then
      Call glbDatabaseMngr.UnLockTable(m_TableName, id, IsCanLock, glbErrorLog)
      Exit Sub
   End If

   Call EnableForm(Me, False)
   If Not glbDaily.DeletePartItem(id, IsOK, True, glbErrorLog) Then
      m_PartItem.PART_ITEM_ID = -1
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call glbDatabaseMngr.UnLockTable(m_TableName, id, IsCanLock, glbErrorLog)
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call QueryData(True)
   
   Call glbDatabaseMngr.UnLockTable(m_TableName, id, IsCanLock, glbErrorLog)
   Call EnableForm(Me, True)
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

   id = Val(GridEX1.Value(1))
   Call glbDatabaseMngr.LockTable(m_TableName, id, IsCanLock, glbErrorLog)

   frmAddEditInventoryDocWHPart.PART_ITEM_ID = id
   frmAddEditInventoryDocWHPart.Area = Area
   frmAddEditInventoryDocWHPart.PART_NO = GridEX1.Value(2)
   frmAddEditInventoryDocWHPart.PART_DESC = GridEX1.Value(3)
   frmAddEditInventoryDocWHPart.WEIGHT_PER_PACK = GridEX1.Value(6)
   frmAddEditInventoryDocWHPart.DOCUMENT_TYPE = PartType
   frmAddEditInventoryDocWHPart.LOCATION_ID = uctlPlaceLookup.MyCombo.ItemData(Minus2Zero(uctlPlaceLookup.MyCombo.ListIndex))
   frmAddEditInventoryDocWHPart.HeaderText = MapText("รายละเอียด")
   frmAddEditInventoryDocWHPart.ShowMode = SHOW_EDIT
   Load frmAddEditInventoryDocWHPart
   frmAddEditInventoryDocWHPart.Show 1
   
   OKClick = frmAddEditInventoryDocWHPart.OKClick
   
   Unload frmAddEditInventoryDocWHPart
   Set frmAddEditInventoryDocWHPart = Nothing
               
   If OKClick Then
      Call QueryData(True)
   End If
   Call glbDatabaseMngr.UnLockTable(m_TableName, id, IsCanLock, glbErrorLog)

End Sub


Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub

Private Sub cmdSearch_Click()
   Call QueryData(True)
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Call LoadLocation(uctlPlaceLookup.MyCombo, m_Locations, 2, , , 19)
      Set uctlPlaceLookup.MyCollection = m_Locations
      
      Call InitLoadPartType2(cboPartType)
      cboPartType.ListIndex = 1
      
      Call InitGoodsOrderBy(cboOrderBy)
      Call InitOrderType(cboOrderType)
      

      
      
      
      Call QueryData(True)
   End If
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long

   If Flag Then
      Call EnableForm(Me, False)
      PartType = cboPartType.ItemData(Minus2Zero(cboPartType.ListIndex))
      Call InitGrid
      m_PartItem.PART_NO = PatchWildCard(txtPartNo.Text)
      m_PartItem.PART_DESC = PatchWildCard(txtPartName.Text)
      
      If PartType = 13 Then
         m_PartItem.DOCUMENT_TYPE_SET = "(13,16,18,20)" 'กลุ่ม bulk 13= รับเข้าปกติ ,16 =ปรับยอด ,18= Bag to Bulk,20=โอน Bulk
      Else
         m_PartItem.DOCUMENT_TYPE_SET = "(14,15,17,19)" 'กลุ่ม bulk 14= รับเข้าปกติ ,15 =ปรับยอด ,17= Bag to Bag,19= โอน Bag
      End If
      m_PartItem.LOCATION_ID = uctlPlaceLookup.MyCombo.ItemData(Minus2Zero(uctlPlaceLookup.MyCombo.ListIndex))
      m_PartItem.DOCUMENT_TYPE = -1
      m_PartItem.OrderBy = cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex))
      m_PartItem.OrderType = cboOrderType.ItemData(Minus2Zero(cboOrderType.ListIndex))
      m_PartItem.QueryFlag = 0
      m_PartItem.CANCEL_FLAG = "N"
      If Not glbDaily.QueryLotItemWhDistinctPart(m_PartItem, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   GridEX1.ItemCount = m_Rs.RecordCount 'itemcount
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
      KeyCode = 0
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
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 4000
   Col.Caption = MapText("เบอร์สินค้า")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 4000
   Col.Caption = MapText("ชื่อสินค้า")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 2520
   Col.Caption = MapText("รหัสขาย")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 3555
   Col.Caption = MapText("รายละเอียด")
   
   If PartType = 13 Then
      Set Col = GridEX1.Columns.add '6
      Col.Width = 0
      Col.Caption = MapText("น้ำหนักต่อถุง")
   Else
      Set Col = GridEX1.Columns.add '6
      Col.Width = 1500
      Col.Caption = MapText("น้ำหนักต่อถุง")
   End If
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = 0
   Col.Caption = MapText("DOCUMENT_TYPE")
      
   GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = MapText("ข้อมูลสินค้าคงเหลือ")
   pnlHeader.Caption = MapText("ข้อมูลสินค้าคงเหลือ")
   
   Call InitGrid

   Call InitNormalLabel(lblPartName, MapText("ชื่อสินค้า"))
   Call InitNormalLabel(lblPartType, MapText("ประเภทบรรจุ"))
   Call InitNormalLabel(lblPartNo, MapText("เบอร์สินค้า"))
   Call InitNormalLabel(lblPlaceLookup, MapText("สถานที่จัดเก็บ"))
   Call InitNormalLabel(lblOrderBy, MapText("เรียงตาม"))
   Call InitNormalLabel(lblOrderType, MapText("เรียงจาก"))
   
   Call txtPartName.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   Call txtPartNo.SetKeySearch("PART_NO")
   
   Call InitCombo(cboPartType)
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
   cmdAdjust.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา (F5)"))
   Call InitMainButton(cmdClear, MapText("เคลียร์ (F4)"))
   Call InitMainButton(cmdAdjust, MapText("คำนวณยอดคงเหลือ"))
   
End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   m_TableName = "USER_GROUP"

   Set m_PartItem = New CLotItemWH
   Set m_TempPartItem = New CLotItemWH
   Set m_Locations = New Collection

   Set m_Rs = New ADODB.Recordset

   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_PartItem = Nothing
   Set m_TempPartItem = Nothing
   Set m_Locations = Nothing
   Set m_Rs = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
Dim TempID1 As Long
Dim BD As CPartItem
Dim IsOK As Boolean
Dim OKClick As Boolean

   If GridEX1.ItemCount <= 0 Then
         Exit Sub
   End If
   
   TempID1 = GridEX1.Value(1)
   If Button = 2 Then
      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.Popup("คัดลอกข้อมูล")
      If lMenuChosen = 0 Then
         Exit Sub
      End If
      Set oMenu = Nothing
   Else
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   If lMenuChosen = 1 Then
      Set BD = New CPartItem
      BD.PART_ITEM_ID = TempID1
      Call glbDaily.CopyPartItem(BD, IsOK, True, -1, glbErrorLog)
      Call QueryData(True)
      Set BD = Nothing
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   RowBuffer.RowStyle = RowBuffer.Value(5)
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

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
   Call m_TempPartItem.PopulateFromRS(14, m_Rs)
   
   Values(1) = m_TempPartItem.PART_ITEM_ID
   Values(2) = m_TempPartItem.PART_NO
   Values(3) = m_TempPartItem.PART_DESC
   Values(4) = m_TempPartItem.BARCODE_NO
   Values(5) = m_TempPartItem.PART_DESC
   Values(6) = m_TempPartItem.WEIGHT_PER_PACK
   Values(7) = m_TempPartItem.DOCUMENT_TYPE
   
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
   cmdAdjust.Top = ScaleHeight - 580
   cmdOK.Top = ScaleHeight - 580
   cmdExit.Top = ScaleHeight - 580
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
   cmdAdjust.Left = cmdOK.Left - cmdAdjust.Width - 50
End Sub

