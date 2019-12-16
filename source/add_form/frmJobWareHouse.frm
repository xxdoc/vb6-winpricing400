VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmJobWareHouse 
   Caption         =   "Form1"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmJobWareHouse.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8445
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8520
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboOrdertype 
         Height          =   315
         Left            =   6060
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2370
         Width           =   2655
      End
      Begin VB.ComboBox cboOrderBy 
         Height          =   315
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2370
         Width           =   3100
      End
      Begin VB.ComboBox cboJobProcess 
         Height          =   315
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1920
         Width           =   3105
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   12645
         _ExtentX        =   22304
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlDate uctlJobDate 
         Height          =   405
         Left            =   6060
         TabIndex        =   1
         Top             =   1020
         Width           =   3855
         _extentx        =   6800
         _extenty        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtJobNo 
         Height          =   435
         Left            =   1650
         TabIndex        =   0
         Top             =   1020
         Width           =   2685
         _extentx        =   4736
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtBatchNo 
         Height          =   435
         Left            =   1650
         TabIndex        =   2
         Top             =   1470
         Width           =   2685
         _extentx        =   4736
         _extenty        =   767
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   4665
         Left            =   135
         TabIndex        =   10
         Top             =   3030
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   8229
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
         Column(1)       =   "frmJobWareHouse.frx":030A
         Column(2)       =   "frmJobWareHouse.frx":03D2
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmJobWareHouse.frx":0476
         FormatStyle(2)  =   "frmJobWareHouse.frx":05D2
         FormatStyle(3)  =   "frmJobWareHouse.frx":0682
         FormatStyle(4)  =   "frmJobWareHouse.frx":0736
         FormatStyle(5)  =   "frmJobWareHouse.frx":080E
         ImageCount      =   0
         PrinterProperties=   "frmJobWareHouse.frx":08C6
      End
      Begin prjFarmManagement.uctlTextBox txtPartNo 
         Height          =   435
         Left            =   6060
         TabIndex        =   5
         Top             =   1920
         Width           =   2685
         _extentx        =   4736
         _extenty        =   767
      End
      Begin Threed.SSCommand cmdOther 
         Height          =   525
         Left            =   6930
         TabIndex        =   14
         Top             =   7800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin VB.Label lblPartNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblBatchNo"
         Height          =   315
         Left            =   4740
         TabIndex        =   25
         Top             =   2010
         Width           =   1245
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "lblOrderBy"
         Height          =   315
         Left            =   390
         TabIndex        =   24
         Top             =   2430
         Width           =   1185
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "lblOrderType"
         Height          =   315
         Left            =   4890
         TabIndex        =   23
         Top             =   2460
         Width           =   1095
      End
      Begin Threed.SSCheck chkCommit 
         Height          =   435
         Left            =   6090
         TabIndex        =   3
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblJobDate 
         Alignment       =   1  'Right Justify
         Caption         =   "lblJobDate"
         Height          =   315
         Left            =   4740
         TabIndex        =   22
         Top             =   1110
         Width           =   1245
      End
      Begin VB.Label lblBatchNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblBatchNo"
         Height          =   315
         Left            =   60
         TabIndex        =   21
         Top             =   1560
         Width           =   1515
      End
      Begin VB.Label lblJobProcess 
         Alignment       =   1  'Right Justify
         Caption         =   "lblJobProcess"
         Height          =   315
         Left            =   360
         TabIndex        =   20
         Top             =   2010
         Width           =   1215
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   10140
         TabIndex        =   9
         Top             =   1650
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   10140
         TabIndex        =   8
         Top             =   1050
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmJobWareHouse.frx":0A9E
         ButtonStyle     =   3
      End
      Begin VB.Label lblJobNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblJobNo"
         Height          =   315
         Left            =   60
         TabIndex        =   19
         Top             =   1110
         Width           =   1485
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8535
         TabIndex        =   15
         Top             =   7800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10140
         TabIndex        =   16
         Top             =   7800
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
         Top             =   7800
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
         Top             =   7800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   13
         Top             =   7800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmJobWareHouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_Job As CJobWareHouse
Private m_TempJob As CJobWareHouse
Private m_Rs As ADODB.Recordset
Private m_TableName As String
Private m_HasModify As Boolean

Public OKClick As Boolean
Public JobDocType As Long
Public ProcessID As Long

Private Sub cmdOther_Click()
Dim cPopup As cPopupMenu
Dim lMenuChosen As Long
   If Not VerifyAccessRight("PRODUCT_JOB_EDIT") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   Set cPopup = New cPopupMenu
   lMenuChosen = cPopup.Popup("Import ข้อมูลจาก PLC", "-", "กระจายปริมาณวัตถุดิบ", "-", "กระจายปริมาณผลิตภัณฑ์", "-", "Import ข้อมูลจาก PACKING")
   Set cPopup = Nothing
   
   If lMenuChosen = 1 Then
      frmImportPlcItem.ProcessID = ProcessID
      frmImportPlcItem.JobDocType = JobDocType
      Load frmImportPlcItem
      frmImportPlcItem.Show 1
      
      Unload frmImportPlcItem
      Set frmImportPlcItem = Nothing
   ElseIf lMenuChosen = 3 Then
      frmAddEditQuantityExtract.ProcessID = ProcessID
      frmAddEditQuantityExtract.TX_TYPE = "E"
      frmAddEditQuantityExtract.Area = 1
      frmAddEditQuantityExtract.HeaderText = "กระจายปริมาณวัตถุดิบ"
      frmAddEditQuantityExtract.ShowMode = SHOW_ADD
      Load frmAddEditQuantityExtract
      frmAddEditQuantityExtract.Show 1
      
      Unload frmAddEditQuantityExtract
      Set frmAddEditQuantityExtract = Nothing
   ElseIf lMenuChosen = 5 Then
      frmAddEditQuantityExtract.ProcessID = ProcessID
      frmAddEditQuantityExtract.TX_TYPE = "I"
      frmAddEditQuantityExtract.Area = 2
      frmAddEditQuantityExtract.HeaderText = "กระจายปริมาณผลิตภัณฑ์"
      frmAddEditQuantityExtract.ShowMode = SHOW_ADD
      Load frmAddEditQuantityExtract
      frmAddEditQuantityExtract.Show 1
      
      Unload frmAddEditQuantityExtract
      Set frmAddEditQuantityExtract = Nothing
      ElseIf lMenuChosen = 7 Then
         Load frmImportPacking 'frmImportPacking     frmAddImportJob
         frmImportPacking.Show 1
         
         Unload frmImportPacking
         Set frmImportPacking = Nothing
   End If
End Sub

Private Sub cmdSearch_Click()
   Call QueryData(True)
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      m_Job.JOB_ID = -1
      m_Job.JOB_NO = txtJobNo.Text
      m_Job.FROM_DATE = uctlJobDate.ShowDate
      m_Job.TO_DATE = uctlJobDate.ShowDate
      m_Job.BATCH_NO = txtBatchNo.Text
      m_Job.PROCESS_ID = ProcessID 'cboJobProcess.ItemData(Minus2Zero(cboJobProcess.ListIndex))
      m_Job.OrderBy = cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex))
      m_Job.OrderType = cboOrdertype.ItemData(Minus2Zero(cboOrdertype.ListIndex))
     m_Job.COMMIT_FLAG = Check2Flag(chkCommit.Value)
     m_Job.JOB_DOC_TYPE = JobDocType
     m_Job.PART_NO = txtPartNo.Text
     
      If Not glbProductionWH.QueryJobWareHouse(m_Job, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      cmdDelete.Enabled = (m_Job.COMMIT_FLAG = "N")
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

Private Sub cboGroup_Click()
   m_HasModify = True
End Sub

Private Sub chkEnable_Click()
   m_HasModify = True
End Sub
Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean
   If Not VerifyAccessRight("PRODUCT_JOB_ADD") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   frmAddEditJobWareHouse.ProcessID = ProcessID
   frmAddEditJobWareHouse.JobDocType = JobDocType
   If JobDocType = 1 Then
      frmAddEditJobWareHouse.HeaderText = MapText("เพิ่มข้อมูลการบรรจุอาหาร")
   End If
   frmAddEditJobWareHouse.ShowMode = SHOW_ADD
   Load frmAddEditJobWareHouse
   frmAddEditJobWareHouse.Show 1
   
   OKClick = frmAddEditJobWareHouse.OKClick
   
   Unload frmAddEditJobWareHouse
   Set frmAddEditJobWareHouse = Nothing
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub


Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
   If Not VerifyAccessRight("PRODUCT_JOB_DELETE") Then
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
   If Not glbProductionWH.DeleteJob(ID, IsOK, True, glbErrorLog) Then
      m_Job.JOB_ID = -1
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
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(1))
   Call glbDatabaseMngr.LockTable(m_TableName, ID, IsCanLock, glbErrorLog)
   
   frmAddEditJobWareHouse.ProcessID = ProcessID
   frmAddEditJobWareHouse.JobDocType = JobDocType
   frmAddEditJobWareHouse.ID = ID
'   If JobDocType = 1 Then
   frmAddEditJobWareHouse.HeaderText = MapText("แก้ไขข้อมูลการบรรจุอาหาร")
'   Else
'   frmAddEditJobWareHouse.HeaderText = MapText("แก้ไขข้อมูลใบประเมินราคา")
'   End If
   
   frmAddEditJobWareHouse.ShowMode = SHOW_EDIT
   Load frmAddEditJobWareHouse
   frmAddEditJobWareHouse.Show 1
   
   OKClick = frmAddEditJobWareHouse.OKClick
   
   Unload frmAddEditJobWareHouse
   Set frmAddEditJobWareHouse = Nothing
               
   If OKClick Then
      Call QueryData(True)
   End If
   Call glbDatabaseMngr.UnLockTable(m_TableName, ID, IsCanLock, glbErrorLog)

End Sub

Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub

Private Sub Form_Activate()
 If Not m_HasActivate Then
      m_HasActivate = True
      
      Call LoadProcess(cboJobProcess, , ProcessID)
            
      Call InitJobOrderBy(cboOrderBy)
      Call InitOrderType(cboOrdertype)
      
      Call QueryData(True)
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.NAME
      glbErrorLog.ShowUserError
   ElseIf Shift = 0 And KeyCode = 116 Then
      Call cmdSearch_Click
   ElseIf Shift = 0 And KeyCode = 115 Then
      Call cmdClear_Click
   ElseIf Shift = 0 And KeyCode = 118 Then
      Call cmdAdd_Click
   ElseIf Shift = 0 And KeyCode = 117 Then
      Call cmdDelete_Click
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
   ElseIf Shift = 0 And KeyCode = 114 Then
      Call cmdEdit_Click
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
   ElseIf Shift = 0 And KeyCode = 123 Then
      Call AddMemoNote
      KeyCode = 0
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_Job = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub GridEX1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
Dim TempID1 As Long
Dim Bd As CJobWareHouse
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
      Set Bd = New CJobWareHouse
      Bd.JOB_ID = TempID1
      Call glbDaily.CopyJob(Bd, IsOK, True, JobDocType, glbErrorLog)
      Call QueryData(True)
      Set Bd = Nothing
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub InitGrid1()
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
   Col.Width = 1500
   Col.Caption = "เลขที่ใบสั่งผลิต"

   Set Col = GridEX1.Columns.add '3
   Col.Width = 4680
   Col.Caption = MapText("รายละเอียดงาน")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 2000
   Col.Caption = MapText("วันที่สั่งผลิต")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 1500
   Col.Caption = MapText("จำนวนแบต")
   
   Set Col = GridEX1.Columns.add '6
   Col.Width = 3500
   Col.Caption = MapText("โปรเซส")
 
   Set Col = GridEX1.Columns.add '7
   Col.Width = 2000
   Col.Caption = MapText("วันเริ่มผลิต")

   Set Col = GridEX1.Columns.add '8
   Col.Width = 2000
   Col.Caption = MapText("วันที่ผลิตเสร็จ")
   
   Set Col = GridEX1.Columns.add '9
   Col.Width = 2500
   Col.Caption = MapText("ผู้อนุมัติ")
   
   Set Col = GridEX1.Columns.add '10
   Col.Width = 2500
   Col.Caption = MapText("ผู้รับผิดชอบ")
      
   Set Col = GridEX1.Columns.add '11
   Col.Width = 2500
   Col.Visible = False
   Col.Caption = MapText("สถานะ")
      
   GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   If JobDocType = 1 Then
      Me.Caption = MapText("ใบสั่งผลิต")
      pnlHeader.Caption = MapText("ใบสั่งผลิต")
      
      Call InitNormalLabel(lblJobNo, MapText("เลขที่ใบสั่งผลิต"))
      Call InitNormalLabel(lblJobDate, MapText("วันที่สั่งผลิต"))
   ElseIf JobDocType = 2 Then
      Me.Caption = MapText("ใบประเมินราคา")
      pnlHeader.Caption = MapText("ใบประเมินราคา")
      
      Call InitNormalLabel(lblJobNo, MapText("เลขที่ใบประเมิน"))
      Call InitNormalLabel(lblJobDate, MapText("วันที่ประเมิน"))
   End If
   
   Call InitNormalLabel(lblBatchNo, MapText("หมายเลขแบต"))
   Call InitNormalLabel(lblJobProcess, MapText("โปรเซส"))
   Call InitNormalLabel(lblPartNo, MapText("รหัสสินค้า"))

   Call InitNormalLabel(lblOrderBy, MapText("เรียงตาม"))
   Call InitNormalLabel(lblOrderType, MapText("เรียงจาก"))
   
   Call InitCheckBox(chkCommit, "ผลิตเสร็จ")
   
   Call txtJobNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtBatchNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   Call InitCombo(cboJobProcess)
   Call InitCombo(cboOrderBy)
   Call InitCombo(cboOrdertype)
  
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdClear.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOther.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา (F5)"))
   Call InitMainButton(cmdClear, MapText("เคลียร์ (F4)"))
   Call InitMainButton(cmdOther, MapText("อื่น ๆ"))
   
   Call InitGrid1
End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   m_TableName = "USER_GROUP"
   Call InitFormLayout
      
   Set m_Rs = New ADODB.Recordset
   Set m_Job = New CJobWareHouse
   Set m_TempJob = New CJobWareHouse
   Call EnableForm(Me, True)
End Sub
Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   RowBuffer.RowStyle = RowBuffer.Value(11)
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
   Call m_TempJob.PopulateFromRS(1, m_Rs)
   
   Values(1) = m_TempJob.JOB_ID
   Values(2) = m_TempJob.JOB_NO
   Values(3) = m_TempJob.JOB_DESC
   Values(4) = DateToStringExtEx2(m_TempJob.JOB_DATE)
   Values(5) = m_TempJob.BATCH_NO
   Values(6) = m_TempJob.PROCESS_NAME
   Values(7) = DateToStringExtEx2(m_TempJob.START_DATE)
   Values(8) = DateToStringExtEx2(m_TempJob.FINISH_DATE)
   Values(9) = m_TempJob.LONG_NAMEA & " " & m_TempJob.LAST_NAMEA
   Values(10) = m_TempJob.LONG_NAMER & " " & m_TempJob.LAST_NAMER
   Values(11) = m_TempJob.COMMIT_FLAG
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub cmdClear_Click()
   txtJobNo.Text = ""
   txtBatchNo.Text = ""
   txtPartNo.Text = ""
   cboOrderBy.ListIndex = -1
   cboOrdertype.ListIndex = -1
   cboJobProcess.ListIndex = -1
   uctlJobDate.ShowDate = -1
   
   chkCommit.Value = ssCBUnchecked
End Sub

Public Sub LoadRefDoc(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CInventoryDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CInventoryDoc
Dim i As Long

   Set D = New CInventoryDoc
   Set Rs = New ADODB.Recordset
   D.COMMIT_FLAG = "Y"
   D.INVENTORY_DOC_ID = -1
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      i = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      i = i + 1
      Set TempData = New CInventoryDoc
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.DOCUMENT_NO)
         C.ItemData(i) = TempData.INVENTORY_DOC_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, str(TempData.INVENTORY_DOC_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
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
   cmdEdit.Top = cmdAdd.Top
   cmdDelete.Top = cmdAdd.Top
   cmdOK.Top = cmdAdd.Top
   cmdExit.Top = cmdAdd.Top
   cmdOther.Top = cmdAdd.Top
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
   cmdOther.Left = cmdOK.Left - cmdOther.Width - 50
End Sub

