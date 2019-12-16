VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmLotItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   Icon            =   "frmLotItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   7680
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   4725
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   8334
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   7
         Top             =   0
         Width           =   7635
         _ExtentX        =   13467
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   3165
         Left            =   0
         TabIndex        =   0
         Top             =   690
         Width           =   7665
         _ExtentX        =   13520
         _ExtentY        =   5583
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
         Column(1)       =   "frmLotItem.frx":27A2
         Column(2)       =   "frmLotItem.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmLotItem.frx":290E
         FormatStyle(2)  =   "frmLotItem.frx":2A6A
         FormatStyle(3)  =   "frmLotItem.frx":2B1A
         FormatStyle(4)  =   "frmLotItem.frx":2BCE
         FormatStyle(5)  =   "frmLotItem.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmLotItem.frx":2D5E
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   3
         Top             =   3300
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmLotItem.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   1
         Top             =   3300
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmLotItem.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   2
         Top             =   3300
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   3870
         TabIndex        =   5
         Top             =   4020
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2220
         TabIndex        =   4
         Top             =   4020
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmLotItem.frx":356A
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmLotItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_LotItem As CLotItem
Private m_TempLotItem As CLotItem
Private m_Rs As ADODB.Recordset
Private m_TableName As String
Private m_Layouts As Collection
Private m_PartItems As Collection

Public OKClick As Boolean
Public PartItemID As Long
Public LocationID As Long
Public LayoutID As Long

Private Sub cmdPasswd_Click()

End Sub

Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean
      
   frmAddEditInventoryDoc1.HeaderText = MapText("เพิ่มข้อมูลการนำเข้า")
   frmAddEditInventoryDoc1.ShowMode = SHOW_ADD
   Load frmAddEditInventoryDoc1
   frmAddEditInventoryDoc1.Show 1
   
   OKClick = frmAddEditInventoryDoc1.OKClick
   
   Unload frmAddEditInventoryDoc1
   Set frmAddEditInventoryDoc1 = Nothing
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub

Private Sub cmdClear_Click()

End Sub

Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim id As Long

   If Not cmdDelete.Enabled Then
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
   If Not glbDaily.DeleteInventoryDoc(id, IsOK, True, glbErrorLog) Then
      m_LotItem.INVENTORY_DOC_ID = -1
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
   
   frmAddEditInventoryDoc1.id = id
   frmAddEditInventoryDoc1.HeaderText = MapText("แก้ไขข้อมูลการนำเข้า")
   frmAddEditInventoryDoc1.ShowMode = SHOW_EDIT
   Load frmAddEditInventoryDoc1
   frmAddEditInventoryDoc1.Show 1
   
   OKClick = frmAddEditInventoryDoc1.OKClick
   
   Unload frmAddEditInventoryDoc1
   Set frmAddEditInventoryDoc1 = Nothing
               
   If OKClick Then
      Call QueryData(True)
   End If
   Call glbDatabaseMngr.UnLockTable(m_TableName, id, IsCanLock, glbErrorLog)

End Sub

Private Sub cmdOK_Click()
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   LayoutID = Val(GridEX1.Value(1))
   
   OKClick = True
   Unload Me
End Sub

Private Sub cmdSearch_Click()
   Call QueryData(True)
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      
      Call EnableForm(Me, False)
      Call LoadPartItem(Nothing, m_PartItems)
      Call LoadLayout(Nothing, m_Layouts)
      Call EnableForm(Me, True)
      
      Call QueryData(True)
   End If
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long

   If Flag Then
      Call EnableForm(Me, False)
         
      m_LotItem.PART_ITEM_ID = PartItemID
      m_LotItem.LOCATION_ID = LocationID
      m_LotItem.COMMIT_FLAG = "Y"
      m_LotItem.TX_TYPE = "I"
      Call m_LotItem.QueryData(1, m_Rs, ItemCount)
      IsOK = True
      
      cmdDelete.Enabled = (m_LotItem.COMMIT_FLAG = "N")
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

Private Sub Form_f(KeyCode As Integer, Shift As Integer)
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
   Col.Width = 1845
   Col.Caption = MapText("วันที่เอกสาร")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 1890
   Col.Caption = MapText("เลขที่เอกสาร")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 1680
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ยอดเหลือ")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 1935
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ยอดนำเข้า")
   
   GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = MapText("ข้อมูลการนำเข้า")
   pnlHeader.Caption = MapText("ข้อมูลการนำเข้า")
   
   Call InitGrid
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
'   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
'   cmdClear.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
'   Call InitCheckBox(chkCommit, "คำนวณแล้ว")
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
'   Call InitMainButton(cmdSearch, MapText("ค้นหา (F5)"))
'   Call InitMainButton(cmdClear, MapText("เคลียร์ (F4)"))
End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   m_TableName = "USER_GROUP"
   
   Set m_LotItem = New CLotItem
   Set m_TempLotItem = New CLotItem
   Set m_Rs = New ADODB.Recordset

   Set m_Layouts = New Collection
   Set m_PartItems = New Collection

   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_Layouts = Nothing
   Set m_PartItems = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub GridEX1_DblClick()
'   Call cmdOK_Click
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
'   RowBuffer.RowStyle = RowBuffer.Value(5)
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long
Dim Pi As CPartItem
Dim Lo As CLayout

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
   Call m_TempLotItem.PopulateFromRS(1, m_Rs)
   
   Values(1) = m_TempLotItem.LAYOUT_ID
   Values(2) = DateToStringExtEx2(m_TempLotItem.DOCUMENT_DATE)
   Values(3) = m_TempLotItem.DOCUMENT_NO
   Values(4) = FormatNumber(m_TempLotItem.LEFT_AMOUNT)
   Values(5) = FormatNumber(m_TempLotItem.TX_AMOUNT)
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

