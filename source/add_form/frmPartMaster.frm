VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmPartMaster 
   BackColor       =   &H80000000&
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   13260
   Icon            =   "frmPartMaster.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   13260
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   14235
      _ExtentX        =   25109
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboUnit 
         Height          =   315
         Left            =   6450
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1530
         Width           =   2985
      End
      Begin VB.ComboBox cboPartType 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1530
         Width           =   2985
      End
      Begin VB.ComboBox cboOrderType 
         Height          =   315
         Left            =   6450
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1980
         Width           =   2955
      End
      Begin VB.ComboBox cboOrderBy 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1980
         Width           =   2985
      End
      Begin prjFarmManagement.uctlTextBox txtPartName 
         Height          =   435
         Left            =   6450
         TabIndex        =   1
         Top             =   1080
         Width           =   2985
         _extentx        =   5265
         _extenty        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   15
         Top             =   0
         Width           =   13245
         _ExtentX        =   23363
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   5055
         Left            =   180
         TabIndex        =   8
         Top             =   2640
         Width           =   11505
         _ExtentX        =   20294
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
         Column(1)       =   "frmPartMaster.frx":27A2
         Column(2)       =   "frmPartMaster.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmPartMaster.frx":290E
         FormatStyle(2)  =   "frmPartMaster.frx":2A6A
         FormatStyle(3)  =   "frmPartMaster.frx":2B1A
         FormatStyle(4)  =   "frmPartMaster.frx":2BCE
         FormatStyle(5)  =   "frmPartMaster.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmPartMaster.frx":2D5E
      End
      Begin prjFarmManagement.uctlTextBox txtPartNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   1080
         Width           =   2985
         _extentx        =   13309
         _extenty        =   767
      End
      Begin Threed.SSCheck chkCancelFlag 
         Height          =   375
         Left            =   9600
         TabIndex        =   22
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4920
         TabIndex        =   21
         Top             =   1590
         Width           =   1455
      End
      Begin VB.Label lblPartNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   20
         Top             =   1140
         Width           =   1755
      End
      Begin VB.Label lblPartType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   19
         Top             =   1590
         Width           =   1755
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4980
         TabIndex        =   18
         Top             =   2040
         Width           =   1365
      End
      Begin VB.Label lblPartName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4920
         TabIndex        =   17
         Top             =   1140
         Width           =   1455
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   16
         Top             =   2040
         Width           =   1755
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   11070
         TabIndex        =   6
         Top             =   1170
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmPartMaster.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   11070
         TabIndex        =   7
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
         TabIndex        =   11
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmPartMaster.frx":3250
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
         MouseIcon       =   "frmPartMaster.frx":356A
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
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10095
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmPartMaster.frx":3884
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmPartMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_PartMaster As CPartMaster
Private m_TempPartMaster As CPartMaster
Private m_Rs As ADODB.Recordset
Private m_TableName As String
Public OKClick As Boolean
Public DocumentType As Long
Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean
   If Not VerifyAccessRight("INVENTORY_PART-MASTER_ADD") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   frmAddEditPartMaster.PartMasterType = DocumentType
   frmAddEditPartMaster.HeaderText = MapText("เพิ่มข้อมูลหลัก")
   frmAddEditPartMaster.ShowMode = SHOW_ADD
   Load frmAddEditPartMaster
   frmAddEditPartMaster.Show 1

   OKClick = frmAddEditPartMaster.OKClick

   Unload frmAddEditPartMaster
   Set frmAddEditPartMaster = Nothing
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub

Private Sub cmdClear_Click()
   txtPartName.Text = ""
   txtPartNo.Text = ""
   cboOrderBy.ListIndex = -1
   cboOrderType.ListIndex = -1
   cboPartType.ListIndex = -1
   cboUnit.ListIndex = -1
End Sub

Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
   If Not VerifyAccessRight("INVENTORY_PART-MASTER_DELETE") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   ID = GridEX1.Value(1)
   
   Call glbDatabaseMngr.LockTable(m_TableName, ID, IsCanLock, glbErrorLog)
   If Not ConfirmDelete(GridEX1.Value(2)) Then
      Call glbDatabaseMngr.UnLockTable(m_TableName, ID, IsCanLock, glbErrorLog)
      Exit Sub
   End If

   Call EnableForm(Me, False)
   If Not glbDaily.DeletePartMaster(ID, IsOK, True, glbErrorLog) Then
      m_PartMaster.PART_MASTER_ID = -1
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call glbDatabaseMngr.UnLockTable(m_TableName, ID, IsCanLock, glbErrorLog)
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call QueryData(True)
   
   Call glbDatabaseMngr.UnLockTable(m_TableName, ID, IsCanLock, glbErrorLog)
   Call EnableForm(Me, True)
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean
   If Not VerifyAccessRight("INVENTORY_PART-MASTER_EDIT") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(1))
   Call glbDatabaseMngr.LockTable(m_TableName, ID, IsCanLock, glbErrorLog)
   
   frmAddEditPartMaster.PartMasterType = DocumentType
   frmAddEditPartMaster.ID = ID
   frmAddEditPartMaster.HeaderText = MapText("แก้ไขข้อมูล MASTER")
   frmAddEditPartMaster.ShowMode = SHOW_EDIT
   Load frmAddEditPartMaster
   frmAddEditPartMaster.Show 1

   OKClick = frmAddEditPartMaster.OKClick

   Unload frmAddEditPartMaster
   Set frmAddEditPartMaster = Nothing
               
   If OKClick Then
      Call QueryData(True)
   End If
   Call glbDatabaseMngr.UnLockTable(m_TableName, ID, IsCanLock, glbErrorLog)

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
      
'      Call LoadPartType(cboPartType, , PartGroupID)
'      Call LoadUnit(cboUnit)
      
      Call InitPartItemOrderBy(cboOrderBy)
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
      
      m_PartMaster.PART_MASTER_NO = PatchWildCard(txtPartNo.Text)
      m_PartMaster.PART_MASTER_NAME = PatchWildCard(txtPartName.Text)
      m_PartMaster.OrderBy = cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex))
      m_PartMaster.OrderType = cboOrderType.ItemData(Minus2Zero(cboOrderType.ListIndex))
      m_PartMaster.PART_MASTER_TYPE = DocumentType
      m_PartMaster.CANCEL_FLAG = Check2Flag(chkCancelFlag.Value)
      
      
      If Not glbDaily.QueryPartMaster(m_PartMaster, m_Rs, ItemCount, IsOK, glbErrorLog) Then
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
   Col.Width = 2200
   Col.Caption = MapText("หมายเลขสินค้า")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 3780
   Col.Caption = MapText("ชื่อสินค้า")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 2100
   Col.Caption = MapText("ประเภทสินค้า")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 2600
   Col.Caption = MapText("วันที่สร้าง/วันที่แก้ไข")
 
   Set Col = GridEX1.Columns.add '6
   Col.Width = 1200
   Col.Caption = MapText("สร้าง")
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = 1200
   Col.Caption = MapText("แก้ไข")
   
   GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = MapText("ข้อมูลวัตถุดิบ")
   pnlHeader.Caption = MapText("ข้อมูลวัตถุดิบ")
   
   Call InitGrid
   
   Call InitNormalLabel(lblPartName, MapText("ชื่อวัตถุดิบ"))
   Call InitNormalLabel(lblUnit, MapText("หน่วยวัด"))
   Call InitNormalLabel(lblPartType, MapText("ประเภทวัตถุดิบ"))
   Call InitNormalLabel(lblPartNo, MapText("หมายเลขวัตถุดิบ"))
   Call InitNormalLabel(lblOrderBy, MapText("เรียงตาม"))
   Call InitNormalLabel(lblOrderType, MapText("เรียงจาก"))
   
   Call txtPartName.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   
   Call InitCombo(cboUnit)
   Call InitCombo(cboPartType)
   Call InitCombo(cboOrderBy)
   Call InitCombo(cboOrderType)
   
   Call InitCheckBox(chkCancelFlag, "ยกเลิก")
   
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
   
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา (F5)"))
   Call InitMainButton(cmdClear, MapText("เคลียร์ (F4)"))

End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   m_TableName = "USER_GROUP"

   Set m_PartMaster = New CPartMaster
   Set m_TempPartMaster = New CPartMaster

   Set m_Rs = New ADODB.Recordset

   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim oMenu As cPopupMenu
'Dim lMenuChosen As Long
'Dim TempID1 As Long
'Dim BD As CPartItem
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
'      lMenuChosen = oMenu.Popup("คัดลอกข้อมูล")
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
'      Set BD = New CPartItem
'      BD.PART_ITEM_ID = TempID1
'      Call glbDaily.CopyPartItem(BD, IsOK, True, -1, glbErrorLog)
'      Call QueryData(True)
'      Set BD = Nothing
'   End If
'
'   Call EnableForm(Me, True)
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
   Call m_TempPartMaster.PopulateFromRS(1, m_Rs)
   
   Values(1) = m_TempPartMaster.PART_MASTER_ID
   Values(2) = m_TempPartMaster.PART_MASTER_NO
   Values(3) = m_TempPartMaster.PART_MASTER_NAME
   Values(4) = m_TempPartMaster.PART_MASTER_TYPE
   Values(5) = DateToStringExtEx2(m_TempPartMaster.CREATE_DATE) & "/" & DateToStringExtEx2(m_TempPartMaster.MODIFY_DATE)
     
   Set Temp_LTK = GetObject("CLoginTracking", m_LoginTracking, Trim(m_TempPartMaster.CREATE_BY), False)
   If Not Temp_LTK Is Nothing Then
      Values(6) = Temp_LTK.USER_NAME
   Else
      Values(6) = ""
   End If
   
   Set Temp_LTK = GetObject("CLoginTracking", m_LoginTracking, Trim(m_TempPartMaster.MODIFY_BY), False)
   If Not Temp_LTK Is Nothing Then
      Values(7) = Temp_LTK.USER_NAME
   Else
      Values(7) = ""
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
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
End Sub

