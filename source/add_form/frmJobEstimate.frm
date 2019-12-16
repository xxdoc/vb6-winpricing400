VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmJobEstimate 
   Caption         =   "Form1"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmJobEstimate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8445
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame SSFrame1 
      Height          =   8520
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboOrdertype 
         Height          =   315
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   3240
         Width           =   3100
      End
      Begin VB.ComboBox cboOrderBy 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3240
         Width           =   3100
      End
      Begin VB.ComboBox cboJobApp 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2280
         Width           =   3100
      End
      Begin VB.ComboBox cboJobProcess 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2760
         Width           =   3100
      End
      Begin VB.ComboBox cboJobRef 
         Height          =   315
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2760
         Width           =   3100
      End
      Begin VB.ComboBox cboJobRes 
         Height          =   315
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2280
         Width           =   3100
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   12645
         _ExtentX        =   22304
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlDate uctlJobDate 
         Height          =   405
         Left            =   1320
         TabIndex        =   3
         Top             =   1320
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtJobNo 
         Height          =   435
         Left            =   1320
         TabIndex        =   1
         Top             =   840
         Width           =   3855
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtBatchNo 
         Height          =   435
         Left            =   6840
         TabIndex        =   4
         Top             =   1320
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtJobDesc 
         Height          =   435
         Left            =   6840
         TabIndex        =   2
         Top             =   840
         Width           =   3855
         _ExtentX        =   6535
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlFinishJob 
         Height          =   405
         Left            =   6840
         TabIndex        =   7
         Top             =   1800
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlDate uctlStartJob 
         Height          =   405
         Left            =   1320
         TabIndex        =   6
         Top             =   1800
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   3855
         Left            =   75
         TabIndex        =   16
         Top             =   3840
         Width           =   11745
         _ExtentX        =   20717
         _ExtentY        =   6800
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
         Column(1)       =   "frmJobEstimate.frx":030A
         Column(2)       =   "frmJobEstimate.frx":03D2
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmJobEstimate.frx":0476
         FormatStyle(2)  =   "frmJobEstimate.frx":05D2
         FormatStyle(3)  =   "frmJobEstimate.frx":0682
         FormatStyle(4)  =   "frmJobEstimate.frx":0736
         FormatStyle(5)  =   "frmJobEstimate.frx":080E
         ImageCount      =   0
         PrinterProperties=   "frmJobEstimate.frx":08C6
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "lblOrderBy"
         Height          =   315
         Left            =   120
         TabIndex        =   34
         Top             =   3360
         Width           =   1185
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "lblOrderType"
         Height          =   315
         Left            =   5640
         TabIndex        =   33
         Top             =   3360
         Width           =   1185
      End
      Begin Threed.SSCheck chkCommit 
         Height          =   435
         Left            =   8280
         TabIndex        =   5
         Top             =   1320
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblJobDate 
         Alignment       =   1  'Right Justify
         Caption         =   "lblJobDate"
         Height          =   315
         Left            =   0
         TabIndex        =   32
         Top             =   1440
         Width           =   1305
      End
      Begin VB.Label lblStartJob 
         Alignment       =   1  'Right Justify
         Caption         =   "lblStartJob"
         Height          =   315
         Left            =   0
         TabIndex        =   31
         Top             =   1920
         Width           =   1305
      End
      Begin VB.Label lblJobApp 
         Alignment       =   1  'Right Justify
         Caption         =   "lblJobApp"
         Height          =   315
         Left            =   0
         TabIndex        =   30
         Top             =   2400
         Width           =   1305
      End
      Begin VB.Label lblJobRes 
         Alignment       =   1  'Right Justify
         Caption         =   "lblJobRes"
         Height          =   315
         Left            =   4560
         TabIndex        =   29
         Top             =   2400
         Width           =   2265
      End
      Begin VB.Label lblJobRef 
         Alignment       =   1  'Right Justify
         Caption         =   "lblJobRef"
         Height          =   315
         Left            =   4440
         TabIndex        =   28
         Top             =   2880
         Width           =   2385
      End
      Begin VB.Label lblBatchNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblBatchNo"
         Height          =   315
         Left            =   5160
         TabIndex        =   27
         Top             =   1440
         Width           =   1665
      End
      Begin VB.Label lblJobDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "lblJobDesc"
         Height          =   315
         Left            =   5160
         TabIndex        =   26
         Top             =   960
         Width           =   1665
      End
      Begin VB.Label lblFinishJob 
         Alignment       =   1  'Right Justify
         Caption         =   "lblFinishJob"
         Height          =   315
         Left            =   5160
         TabIndex        =   25
         Top             =   1920
         Width           =   1665
      End
      Begin VB.Label lblJobProcess 
         Alignment       =   1  'Right Justify
         Caption         =   "lblJobProcess"
         Height          =   315
         Left            =   0
         TabIndex        =   24
         Top             =   2880
         Width           =   1305
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   10200
         TabIndex        =   15
         Top             =   3000
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   10200
         TabIndex        =   14
         Top             =   2400
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmJobEstimate.frx":0A9E
         ButtonStyle     =   3
      End
      Begin VB.Label lblJobNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblJobNo"
         Height          =   315
         Left            =   0
         TabIndex        =   23
         Top             =   960
         Width           =   1305
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8600
         TabIndex        =   20
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
         Left            =   10200
         TabIndex        =   21
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   19
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
Attribute VB_Name = "frmJobEstimate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_Job As CJob
Private m_TempJob As CJob
Private m_Rs As ADODB.Recordset
Private m_TableName As String

Public OKClick As Boolean
Private Sub cmdSearch_Click()
   Call QueryData(True)
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If Not VerifyAccessRight("PRODUCT_JOB_QUERY") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      m_Job.JOB_NO = txtJobNo.Text
      m_Job.JOB_DESC = txtJobDesc.Text
      m_Job.JOB_DATE = uctlJobDate.ShowDate
      m_Job.BATCH_NO = txtBatchNo.Text
      m_Job.APPROVED_BY = cboJobApp.ItemData(Minus2Zero(cboJobApp.ListIndex))
      m_Job.RESPONSE_BY = cboJobRes.ItemData(Minus2Zero(cboJobRes.ListIndex))
     m_Job.START_DATE = uctlStartJob.ShowDate
     m_Job.FINISH_DATE = uctlFinishJob.ShowDate
      m_Job.PROCESS_ID = cboJobProcess.ItemData(Minus2Zero(cboJobProcess.ListIndex))
     m_Job.INVENTORY_DOC_ID = cboJobRef.ItemData(Minus2Zero(cboJobRef.ListIndex))
      m_Job.OrderBy = cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex))
      m_Job.OrderType = cboOrdertype.ItemData(Minus2Zero(cboOrdertype.ListIndex))
     m_Job.COMMIT_FLAG = Check2Flag(chkCommit.Value)
      If Not glbProduction.QueryJob(m_Job, m_Rs, ItemCount, IsOK, glbErrorLog) Then
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

Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean

   If Not VerifyAccessRight("PRODUCT_ESTIMATE_ADD") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

   frmAddEditJobEstimate.HeaderText = MapText("เพิ่มงาน")
   frmAddEditJobEstimate.ShowMode = SHOW_ADD
   Load frmAddEditJobEstimate
   frmAddEditJobEstimate.Show 1
   
   OKClick = frmAddEditJobEstimate.OKClick
   
   Unload frmAddEditJobEstimate
   Set frmAddEditJobEstimate = Nothing
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub


Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long

   If Not VerifyAccessRight("PRODUCT_ESTIMATE_DELETE") Then
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
   If Not glbProduction.DeleteJob(ID, IsOK, True, glbErrorLog) Then
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

   If Not VerifyAccessRight("PRODUCT_ESTIMATE_EDIT") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(1))
   Call glbDatabaseMngr.LockTable(m_TableName, ID, IsCanLock, glbErrorLog)
   
   frmAddEditJobEstimate.ID = ID
   frmAddEditJobEstimate.HeaderText = MapText("เแก้ไขข้อมูลงาน")
   frmAddEditJobEstimate.ShowMode = SHOW_EDIT
   Load frmAddEditJobEstimate
   frmAddEditJobEstimate.Show 1
   
   OKClick = frmAddEditJobEstimate.OKClick
   
   Unload frmAddEditJobEstimate
   Set frmAddEditJobEstimate = Nothing
               
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
      
      Call LoadEmployee(cboJobApp)
      Call LoadEmployee(cboJobRes)
            Call LoadRefDoc(cboJobRef)
            Call LoadProcess(cboJobProcess)
            
      Call InitJobOrderBy(cboOrderBy)
      Call InitOrderType(cboOrdertype)
      
      Call QueryData(True)
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
   ElseIf Shift = 0 And KeyCode = 116 Then
'      Call cmdSearch_Click
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
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
   'debug.print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid1()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
      
GridEX1.FormatStyles.Clear
   Set fmsTemp = GridEX1.FormatStyles.Add("N")
   fmsTemp.ForeColor = GLB_ALERT_COLOR
      
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.Add '2
   Col.Width = 1500
   Col.Caption = "เลขที่ใบสั่งผลิต"

   Set Col = GridEX1.Columns.Add '3
   Col.Width = 4000
   Col.Caption = MapText("รายละเอียดงาน")
   Set Col = GridEX1.Columns.Add '4
   Col.Width = 2000
   Col.Caption = MapText("วันที่งาน")
   Set Col = GridEX1.Columns.Add '5
   Col.Width = 1500
   Col.Caption = MapText("หมายเลขแบท")
   
   Set Col = GridEX1.Columns.Add '6
   Col.Width = 3500
   Col.Caption = MapText("โปรเซส")
 
   Set Col = GridEX1.Columns.Add '7
   Col.Width = 2000
   Col.Caption = MapText("วันเริ่มงาน")

   Set Col = GridEX1.Columns.Add '8
   Col.Width = 2000
   Col.Caption = MapText("วันเสร็จงาน")
   Set Col = GridEX1.Columns.Add '9
   Col.Width = 2500
   Col.Caption = MapText("อนุมัติโดย")
Set Col = GridEX1.Columns.Add '10
   Col.Width = 2500
   Col.Caption = MapText("รับผิดชอบโดย")
Set Col = GridEX1.Columns.Add '11
   Col.Width = 1500
   Col.Caption = MapText("หมายเลขเอกสาร")
Set Col = GridEX1.Columns.Add '12
   Col.Width = 1500
   Col.Caption = MapText("ผลของงาน")
   
   GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = "งาน"
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblJobNo, MapText("เลขที่ใบสั่งผลิต"))
   Call InitNormalLabel(lblJobDesc, MapText("รายละเอียดงาน"))
   Call InitNormalLabel(lblJobDate, MapText("วันที่งาน"))
  Call InitNormalLabel(lblBatchNo, MapText("หมายเลขแบท"))
   Call InitNormalLabel(lblJobApp, MapText("ผู้อนุมัติ"))
   Call InitNormalLabel(lblJobRes, MapText("ผู้รับผิดชอบ"))
Call InitNormalLabel(lblStartJob, MapText("วันที่เริ่มงาน"))
   Call InitNormalLabel(lblFinishJob, MapText("วันที่เสร็จงาน"))
   Call InitNormalLabel(lblJobProcess, MapText("โปรเซส"))
Call InitNormalLabel(lblJobRef, MapText("หมายเลขเอกสารอ้างอิง"))
   Call InitNormalLabel(lblOrderBy, MapText("เรียงตาม"))
   Call InitNormalLabel(lblOrderType, MapText("เรียงจาก"))
   
   Call InitCheckBox(chkCommit, "งานเสร็จแล้ว")
   
   Call txtJobNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtJobDesc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtBatchNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   Call InitCombo(cboJobApp)
   Call InitCombo(cboJobRes)
   Call InitCombo(cboJobProcess)
   Call InitCombo(cboJobRef)
   Call InitCombo(cboOrderBy)
   Call InitCombo(cboOrdertype)
  
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdClear.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา (F5)"))
   Call InitMainButton(cmdClear, MapText("เคลียร์ (F4)"))
   
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
   Set m_Job = New CJob
   Set m_TempJob = New CJob
   Call EnableForm(Me, True)
End Sub
Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
      RowBuffer.RowStyle = RowBuffer.Value(12)
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
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
   Call m_TempJob.PopulateFromRS(m_Rs)
   
   Values(1) = m_TempJob.JOB_ID
   Values(2) = m_TempJob.JOB_NO
   Values(3) = m_TempJob.JOB_DESC
   Values(4) = DateToStringExt(m_TempJob.JOB_DATE)
   Values(5) = m_TempJob.BATCH_NO
   Values(6) = m_TempJob.PROCESS_NAME
   Values(7) = DateToStringExt(m_TempJob.START_DATE)
   Values(8) = DateToStringExt(m_TempJob.FINISH_DATE)
   Values(9) = m_TempJob.LONG_NAMEA & " " & m_TempJob.LAST_NAMEA
   Values(10) = m_TempJob.LONG_NAMER & " " & m_TempJob.LAST_NAMER
            Values(11) = m_TempJob.DOC_NO
      If m_TempJob.COMMIT_FLAG = "Y" Then
      Values(12) = "เสร็จแล้ว"
      Else
      Values(12) = "ยังไม่เสร็จ"
      End If
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub cmdClear_Click()
   txtJobNo.Text = ""
   txtJobDesc.Text = ""
   txtBatchNo.Text = ""
   cboOrderBy.ListIndex = -1
   cboOrdertype.ListIndex = -1
   cboJobApp.ListIndex = -1
   cboJobRes.ListIndex = -1
   cboJobRef.ListIndex = -1
   cboJobProcess.ListIndex = -1
   uctlJobDate.ShowDate = -1
   uctlStartJob.ShowDate = -1
   uctlFinishJob.ShowDate = -1
   chkCommit.Value = ssCBUnchecked
Call QueryData(True)
End Sub

