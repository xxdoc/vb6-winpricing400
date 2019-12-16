VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmJob 
   Caption         =   "Form1"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17160
   Icon            =   "frmJob.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   17160
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8520
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   17055
      _ExtentX        =   30083
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboLotNo 
         Height          =   315
         Left            =   6060
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1920
         Width           =   2655
      End
      Begin VB.ComboBox cboOrdertype 
         Height          =   315
         Left            =   6060
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2370
         Width           =   2655
      End
      Begin VB.ComboBox cboOrderBy 
         Height          =   315
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   5
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
         TabIndex        =   17
         Top             =   0
         Width           =   15285
         _ExtentX        =   26961
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
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtJobNo 
         Height          =   435
         Left            =   1650
         TabIndex        =   0
         Top             =   1020
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtBatchNo 
         Height          =   435
         Left            =   1650
         TabIndex        =   2
         Top             =   1470
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   767
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   4665
         Left            =   135
         TabIndex        =   9
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
         Column(1)       =   "frmJob.frx":030A
         Column(2)       =   "frmJob.frx":03D2
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmJob.frx":0476
         FormatStyle(2)  =   "frmJob.frx":05D2
         FormatStyle(3)  =   "frmJob.frx":0682
         FormatStyle(4)  =   "frmJob.frx":0736
         FormatStyle(5)  =   "frmJob.frx":080E
         ImageCount      =   0
         PrinterProperties=   "frmJob.frx":08C6
      End
      Begin prjFarmManagement.uctlTextBox txtPartNo 
         Height          =   435
         Left            =   6060
         TabIndex        =   27
         Top             =   1440
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtInputAmount 
         Height          =   435
         Left            =   10080
         TabIndex        =   30
         Top             =   1920
         Width           =   1845
         _ExtentX        =   4736
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtOutputAmount 
         Height          =   435
         Left            =   10080
         TabIndex        =   31
         Top             =   2400
         Width           =   1845
         _ExtentX        =   4736
         _ExtentY        =   767
      End
      Begin Threed.SSCommand cmdLockDate 
         Height          =   525
         Left            =   14280
         TabIndex        =   34
         Top             =   2400
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdUnLock 
         Height          =   525
         Left            =   13200
         TabIndex        =   33
         Top             =   2400
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdLock 
         Height          =   525
         Left            =   12120
         TabIndex        =   32
         Top             =   2400
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblOutputAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblOutputAmount"
         Height          =   315
         Left            =   8760
         TabIndex        =   29
         Top             =   2460
         Width           =   1245
      End
      Begin VB.Label lblInputAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblInputAmount"
         Height          =   315
         Left            =   8760
         TabIndex        =   28
         Top             =   2010
         Width           =   1245
      End
      Begin VB.Label lblLotNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblLotNo"
         Height          =   315
         Left            =   4680
         TabIndex        =   26
         Top             =   2010
         Width           =   1245
      End
      Begin Threed.SSCommand cmdOther 
         Height          =   525
         Left            =   6930
         TabIndex        =   13
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
         Caption         =   "lblPartNo"
         Height          =   315
         Left            =   4680
         TabIndex        =   24
         Top             =   1560
         Width           =   1245
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "lblOrderBy"
         Height          =   315
         Left            =   390
         TabIndex        =   23
         Top             =   2430
         Width           =   1185
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "lblOrderType"
         Height          =   315
         Left            =   4890
         TabIndex        =   22
         Top             =   2460
         Width           =   1095
      End
      Begin Threed.SSCheck chkCommit 
         Height          =   435
         Left            =   10320
         TabIndex        =   3
         Top             =   960
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
         TabIndex        =   21
         Top             =   1110
         Width           =   1245
      End
      Begin VB.Label lblBatchNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblBatchNo"
         Height          =   315
         Left            =   60
         TabIndex        =   20
         Top             =   1560
         Width           =   1515
      End
      Begin VB.Label lblJobProcess 
         Alignment       =   1  'Right Justify
         Caption         =   "lblJobProcess"
         Height          =   315
         Left            =   360
         TabIndex        =   19
         Top             =   2010
         Width           =   1215
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   12120
         TabIndex        =   8
         Top             =   1680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   12120
         TabIndex        =   7
         Top             =   1050
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmJob.frx":0A9E
         ButtonStyle     =   3
      End
      Begin VB.Label lblJobNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblJobNo"
         Height          =   315
         Left            =   60
         TabIndex        =   18
         Top             =   1110
         Width           =   1485
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8535
         TabIndex        =   14
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
         TabIndex        =   15
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
         TabIndex        =   11
         Top             =   7800
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   120
         TabIndex        =   10
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
         TabIndex        =   12
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
Attribute VB_Name = "frmJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_Job As CJob
Private m_Job2 As CJob
Private m_TempJob As CJob
Private m_Rs As ADODB.Recordset
Private m_TableName As String
Private m_HasModify As Boolean
Private m_PartItems As Collection
Public OKClick As Boolean
Public JobDocType As Long
Public ProcessID As Long
Public HeaderText As String
Public DOCUMENT_TYPE As Long
Public mainText As String

Private Sub cmdLock_Click()
   If Not VerifyAccessRight("PRODUCT_JOB_LOCK-DOC", "ล็อคเอกสาร") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   Set m_Job2 = New CJob
      m_Job2.JOB_ID = -1
      m_Job2.JOB_NO = txtJobNo.Text
      m_Job2.JOB_DATE = uctlJobDate.ShowDate
      m_Job2.PROCESS_ID = ProcessID
      m_Job2.LOCK_DOC_FLAG = "Y"
   
   If m_Job2.UpdateLockDoc() Then
      glbErrorLog.LocalErrorMsg = MapText("LOCK เอกสารวันที่ " & uctlJobDate.ShowDate & " ของโปรเซส " & m_Job2.PROCESS_NAME & " เรียบร้อยแล้ว")
      glbErrorLog.ShowUserError
      m_Job.QueryFlag = 1
     Call QueryData(True)
   End If
   Set m_Job2 = Nothing
End Sub

Private Sub cmdLockDate_Click()
   If Not VerifyAccessRight("PROGRAM_LOCK-DATE", "กำหนดวันที่เอกสาร") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   frmLockDate.HeaderText = "กำหนดวันที่เอกสาร"
   Load frmLockDate
   frmLockDate.Show 1
   
   Unload frmLockDate
   Set frmLockDate = Nothing
End Sub

Private Sub cmdOther_Click()
Dim cPopup As cPopupMenu
Dim lMenuChosen As Long
   If Not VerifyAccessRight("PRODUCT_JOB_EDIT") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   Set cPopup = New cPopupMenu
   If ProcessID = 6 Or ProcessID = 7 Then
      Exit Sub
   End If
   
   If ProcessID = 1 Or ProcessID = 2 Or ProcessID = 3 Or ProcessID = 5 Then
'      lMenuChosen = cPopup.Popup("กระจายปริมาณวัตถุดิบ", "-", "กระจายปริมาณผลิตภัณฑ์", "-", "Import ข้อมูลจาก PACKING")
     lMenuChosen = cPopup.Popup("กระจายปริมาณวัตถุดิบ", "-", "กระจายปริมาณผลิตภัณฑ์")
   ElseIf ProcessID = 4 Then
      lMenuChosen = cPopup.Popup("กระจายปริมาณวัตถุดิบ", "-", "กระจายปริมาณผลิตภัณฑ์", "-", "IMPORT ข้อมูล BULK จาก PLC")
   End If
   Set cPopup = Nothing
   
    If lMenuChosen = 1 Then
      frmAddEditQuantityExtract.ProcessID = ProcessID
      frmAddEditQuantityExtract.TX_TYPE = "E"
      frmAddEditQuantityExtract.Area = 1
      frmAddEditQuantityExtract.HeaderText = "กระจายปริมาณวัตถุดิบ"
      frmAddEditQuantityExtract.ShowMode = SHOW_ADD
      Load frmAddEditQuantityExtract
      frmAddEditQuantityExtract.Show 1
      
      Unload frmAddEditQuantityExtract
      Set frmAddEditQuantityExtract = Nothing
   ElseIf lMenuChosen = 3 Then
      frmAddEditQuantityExtract.ProcessID = ProcessID
      frmAddEditQuantityExtract.TX_TYPE = "I"
      frmAddEditQuantityExtract.Area = 2
      frmAddEditQuantityExtract.HeaderText = "กระจายปริมาณผลิตภัณฑ์"
      frmAddEditQuantityExtract.ShowMode = SHOW_ADD
      Load frmAddEditQuantityExtract
      frmAddEditQuantityExtract.Show 1
      
      Unload frmAddEditQuantityExtract
      Set frmAddEditQuantityExtract = Nothing
      ElseIf lMenuChosen = 5 Then
         If ProcessID = 2 Then
'''            frmImportPacking.Area = 1
'''            Load frmImportPacking
'''            frmImportPacking.Show 1
'''
'''            Unload frmImportPacking
'''            Set frmImportPacking = Nothing
         Else
            If DOCUMENT_TYPE > 0 Then
               If Not VerifyAccessRight("INVENTORY-WH_IMPORT" & "_" & DOCUMENT_TYPE & "_IMPORT", "อิมพอร์ต ข้อมูลจาก PLC") Then
                  Call EnableForm(Me, True)
                  Exit Sub
               End If
            End If
    
            frmImportPlcItemNew.ProcessID = ProcessID
            frmImportPlcItemNew.JobDocType = JobDocType
            Load frmImportPlcItemNew
            frmImportPlcItemNew.Show 1
            
            Unload frmImportPlcItemNew
            OKClick = frmImportPlcItemNew.OKClick
            Set frmImportPlcItemNew = Nothing
            
            If OKClick Then
            Call QueryData(True)
            End If
         End If
   End If
End Sub

Private Sub cmdSearch_Click()
m_Job.QueryFlag = 1
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
      m_Job.PROCESS_ID = ProcessID
      m_Job.OrderBy = cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex))
      m_Job.OrderType = cboOrderType.ItemData(Minus2Zero(cboOrderType.ListIndex))
     m_Job.COMMIT_FLAG = Check2Flag(chkCommit.Value)
     m_Job.JOB_DOC_TYPE = JobDocType
     m_Job.PART_NO = txtPartNo.Text
     If cboLotNo.ListIndex > -1 Then
      m_Job.LOT_ID = cboLotNo.ItemData(Minus2Zero(cboLotNo.ListIndex))
   Else
      m_Job.LOT_ID = -1
     End If
     
     
     
      If Not glbProduction.QueryJob(m_Job, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      Call CalculateTotalRatio
      
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

  If DOCUMENT_TYPE > 0 Then
      If Not VerifyAccessRight("INVENTORY-WH_IMPORT" & "_" & DOCUMENT_TYPE & "_ADD", "เพิ่ม " & mainText) Then
             Call EnableForm(Me, True)
             Exit Sub
       End If
    End If
    
   frmAddEditJob.ProcessID = ProcessID
   frmAddEditJob.DOCUMENT_TYPE = DOCUMENT_TYPE
   frmAddEditJob.JobDocType = JobDocType
  If JobDocType = 1 Then
   frmAddEditJob.HeaderText = Me.HeaderText
   Else
   frmAddEditJob.HeaderText = MapText("เพิ่มข้อมูลใบประเมินราคา")
   End If
   frmAddEditJob.mainText = mainText
   frmAddEditJob.ShowMode = SHOW_ADD
   Load frmAddEditJob
   frmAddEditJob.Show 1
   
   OKClick = frmAddEditJob.OKClick
   
   Unload frmAddEditJob
   Set frmAddEditJob = Nothing
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub


Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long

 If DOCUMENT_TYPE > 0 Then
   If Not VerifyAccessRight("INVENTORY-WH_IMPORT" & "_" & DOCUMENT_TYPE & "_DELETE", "ลบ " & mainText) Then
          Call EnableForm(Me, True)
          Exit Sub
    End If
   End If
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
   If ProcessID = 4 Then
      If Not VerifyLockInventoryDate(InternalDateToDateExGrid(GridEX1.Value(4)), InternalDateToDateExGrid(GridEX1.Value(4))) Then
         glbErrorLog.LocalErrorMsg = MapText("ไม่สามารถเปลี่ยนแปลงเอกสารตามวันที่เอกสารที่เลือกได้ กรุณาติดต่อผู้ดูแลระบบ หรือผู้มีสิทธิ์กำหนดวันที่เอกสารได้")
         glbErrorLog.ShowUserError
         Exit Sub
      End If
   End If
   
   If GridEX1.Value(22) = "Y" Then ''
      glbErrorLog.LocalErrorMsg = MapText("เอกสาร " & GridEX1.Value(2) & " ถูกล็อคไว้อยู่ไม่สามารถลบข้อมูลได้")
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   
   If Val(GridEX1.Value(23)) > 0 Then ''
      glbErrorLog.LocalErrorMsg = MapText("เอกสาร " & GridEX1.Value(2) & " เป็นเอกสารที่แยกออกมาจากเอกสารอื่นไม่สามารถลบได้")
      glbErrorLog.ShowUserError
       If Not VerifyAccessRight("PRODUCT_JOB_DELETE_SPLIT-DOC", ",ลบเอกสารที่แยกได้") Then
         Call EnableForm(Me, True)
         Exit Sub
       End If
   End If
   
   ID = GridEX1.Value(1)
   
   Call glbDatabaseMngr.LockTable(m_TableName, ID, IsCanLock, glbErrorLog)
   If Not ConfirmDelete(GridEX1.Value(2)) Then
      Call glbDatabaseMngr.UnLockTable(m_TableName, ID, IsCanLock, glbErrorLog)
      Exit Sub
   End If

   Call EnableForm(Me, False)
   If Not glbProduction.DeleteJob(ID, IsOK, True, glbErrorLog, , DOCUMENT_TYPE) Then
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

 If DOCUMENT_TYPE > 0 Then
   If Not VerifyAccessRight("INVENTORY-WH_IMPORT" & "_" & DOCUMENT_TYPE & "_EDIT", "แก้ไข " & mainText) Then
          Call EnableForm(Me, True)
          Exit Sub
    End If
End If
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(1))
   Call glbDatabaseMngr.LockTable(m_TableName, ID, IsCanLock, glbErrorLog)
   ProcessID = Val(GridEX1.Value(21)) ''
   frmAddEditJob.ProcessID = ProcessID
   frmAddEditJob.DOCUMENT_TYPE = DOCUMENT_TYPE
   frmAddEditJob.JobDocType = JobDocType
   frmAddEditJob.ID = ID
   frmAddEditJob.JobIdRef = Val(GridEX1.Value(23)) ''
   If JobDocType = 1 Then
      frmAddEditJob.HeaderText = MapText("แก้ไขข้อมูลใบสั่งผลิต")
   Else
      frmAddEditJob.HeaderText = MapText("แก้ไขข้อมูลใบประเมินราคา")
   End If
   frmAddEditJob.mainText = mainText
   
   frmAddEditJob.ShowMode = SHOW_EDIT
   Load frmAddEditJob
   frmAddEditJob.Show 1
   
   OKClick = frmAddEditJob.OKClick
   
   Unload frmAddEditJob
   Set frmAddEditJob = Nothing
               
   If OKClick Then
      Call QueryData(True)
   End If
   Call glbDatabaseMngr.UnLockTable(m_TableName, ID, IsCanLock, glbErrorLog)

End Sub

Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub

Private Sub cmdUnlock_Click()
  If Not VerifyAccessRight("PRODUCT_JOB_CANCEL-LOCK-DOC", "ปลดล็อคเอกสาร") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
    Set m_Job2 = New CJob
      m_Job2.JOB_ID = -1
      m_Job2.JOB_NO = txtJobNo.Text
      m_Job2.JOB_DATE = uctlJobDate.ShowDate
      m_Job2.PROCESS_ID = ProcessID
      m_Job2.LOCK_DOC_FLAG = "N"
   
   If m_Job2.UpdateLockDoc() Then
      glbErrorLog.LocalErrorMsg = MapText("ยกเลิกการ LOCK เอกสารวันที่ " & uctlJobDate.ShowDate & " ของโปรเซส " & m_Job2.PROCESS_NAME & " เรียบร้อยแล้ว")
      glbErrorLog.ShowUserError
      m_Job.QueryFlag = 1
      Call QueryData(True)
   End If
   Set m_Job2 = Nothing
End Sub

Private Sub Form_Activate()
 If Not m_HasActivate Then
      m_HasActivate = True
      
      Call LoadProcess(cboJobProcess, , ProcessID)
      Call InitJobOrderBy(cboOrderBy)
      Call InitOrderType(cboOrderType)
       uctlJobDate.ShowDate = Now
       
     '' Call LoadLoginTracking(Nothing, m_LoginTracking, DateAdd("M", -1, uctlJobDate.ShowDate), uctlJobDate.ShowDate)
      
      Call QueryData(True)
   End If
End Sub
Private Sub CalculateTotalRatio()
Dim D As CJobInput
Dim O As CJobOutput
Dim Sum1 As Double
Dim Sum2 As Double
Dim Sum3 As Double

   Sum1 = 0
   Sum2 = 0
   Sum3 = 0
   
 While Not m_Rs.EOF
      Call m_TempJob.PopulateFromRS(1, m_Rs)
      Sum1 = Sum1 + m_TempJob.SUM_INPUT
      Sum2 = Sum2 + m_TempJob.SUM_OUTPUT
      m_Rs.MoveNext
   Wend
   
   If m_Rs.RecordCount > 0 Then
      m_Rs.MoveFirst
   End If

   txtInputAmount.Text = FormatNumber(Sum1, 3)
   txtOutputAmount.Text = FormatNumber(Sum2, 3)
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
'   Set m_LoginTracking = Nothing
   Set m_Job = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub GridEX1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
Dim TempID1 As Long
Dim BD As CJob
Dim IsOK As Boolean
Dim OKClick As Boolean

   If GridEX1.ItemCount <= 0 Then
         Exit Sub
   End If
   If ProcessID = 2 Then
      Exit Sub
   End If
   TempID1 = GridEX1.Value(1)
   If Button = 2 Then
      Set oMenu = New cPopupMenu
      If ProcessID = 4 Then
         lMenuChosen = oMenu.Popup("คัดลอกข้อมูล", "-", "แยกข้อมูล")
      Else
         lMenuChosen = oMenu.Popup("คัดลอกข้อมูล")
      End If
      If lMenuChosen = 0 Then
         Exit Sub
      End If
      Set oMenu = Nothing
   Else
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   If lMenuChosen = 1 Then
      Set BD = New CJob
      BD.JOB_ID = TempID1
      Call glbDaily.CopyJob(BD, IsOK, True, JobDocType, glbErrorLog, ProcessID)
      Call QueryData(True)
      Set BD = Nothing
    ElseIf lMenuChosen = 3 Then 'แยกข้อมูล
      If Not VerifyAccessRight("PRODUCT_JOB_SPLIT-DOC", "แยกเอกสาร") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
    
       If GridEX1.Value(22) = "Y" Then ''
         glbErrorLog.LocalErrorMsg = MapText("เอกสาร " & GridEX1.Value(2) & " ถูกล็อคไว้อยู่ไม่สามารถแยกข้อมูลได้")
         glbErrorLog.ShowUserError
         Call EnableForm(Me, True)
         Exit Sub
      End If
    
      frmImportPlcItemNew.ProcessID = ProcessID
      frmImportPlcItemNew.JobDocType = JobDocType
      frmImportPlcItemNew.SplitFlag = True
      frmImportPlcItemNew.JobNo = GridEX1.Value(2)
      Load frmImportPlcItemNew
      frmImportPlcItemNew.Show 1

      Unload frmImportPlcItemNew
      OKClick = frmImportPlcItemNew.OKClick
      Set frmImportPlcItemNew = Nothing
      
      If OKClick Then
        Call QueryData(True)
      End If
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
   Col.Width = 2200
   Col.Caption = "เลขที่ใบสั่งผลิต"

   Set Col = GridEX1.Columns.add '3
   Col.Width = 5000
   Col.Caption = MapText("รายละเอียดงาน")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 1300
   Col.Caption = MapText("วันที่สั่งผลิต")
   
   If ProcessID = 4 Then
      Set Col = GridEX1.Columns.add '5
      Col.Width = 900
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("จากแบต")
      
      Set Col = GridEX1.Columns.add '6
      Col.Width = 900
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ถึงแบต")
      
      Set Col = GridEX1.Columns.add '7
      Col.Width = 2000
      Col.Caption = MapText("รายละเอียดแบต")
   Else
      Set Col = GridEX1.Columns.add '5
      Col.Width = 0
      Col.Caption = MapText("จากแบต")
      
      Set Col = GridEX1.Columns.add '6
      Col.Width = 0
      Col.Caption = MapText("ถึงแบต")
      
      Set Col = GridEX1.Columns.add '7
      Col.Width = 0
      Col.Caption = MapText("รายละเอียดแบต")

   End If
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = 1100
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("จำนวนแบต")
   
    If ProcessID = 4 Then
      Set Col = GridEX1.Columns.add '8
      Col.Width = 1000
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("แบตทั้งหมด")
   Else
      Set Col = GridEX1.Columns.add '8
      Col.Width = 0
      Col.Caption = MapText("แบตทั้งหมด")
   End If
   
   Set Col = GridEX1.Columns.add '9
   Col.Width = 1500
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ยอดใช้")
   
   Set Col = GridEX1.Columns.add '10
   Col.Width = 1500
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ยอดผลิต")
   
   Set Col = GridEX1.Columns.add '11
   Col.Width = 1700
   Col.Caption = MapText("โปรเซส")
 
   Set Col = GridEX1.Columns.add '12
   Col.Width = 1300
   Col.Caption = MapText("วันเริ่มผลิต")

   Set Col = GridEX1.Columns.add '13
   Col.Width = 1300
   Col.Caption = MapText("วันที่ผลิตเสร็จ")
   
   Set Col = GridEX1.Columns.add '14
   Col.Width = 2200
   Col.Caption = MapText("ผู้ตรวจสอบ")

   Set Col = GridEX1.Columns.add '15
   Col.Width = 1200
   Col.Caption = MapText("สร้าง")
   
   Set Col = GridEX1.Columns.add '16
   Col.Width = 1200
   Col.Caption = MapText("แก้ไข")
   
   Set Col = GridEX1.Columns.add '17
   Col.Width = 3000
   Col.Caption = MapText("หมายเหตุ")
   
   Set Col = GridEX1.Columns.add '17
   Col.Width = 2500
   Col.Caption = MapText("ผู้รับผิดชอบ")
      
   Set Col = GridEX1.Columns.add '18
   Col.Width = 2500
   Col.Visible = False
   Col.Caption = MapText("สถานะ")
   
   Set Col = GridEX1.Columns.add '19
   Col.Width = 0
   Col.Visible = False
   Col.Caption = MapText("ProcessID")
   
   Set Col = GridEX1.Columns.add '20
   Col.Width = 0
   Col.Visible = False
   Col.Caption = MapText("LockDocFlag")
   
   Set Col = GridEX1.Columns.add '21
   Col.Width = 0
   Col.Visible = False
   Col.Caption = MapText("JobIdRef")
   
   Set Col = GridEX1.Columns.add '23
   Col.Width = 800
   Col.Caption = MapText("ผู้อนุมัติ")
   
   Set Col = GridEX1.Columns.add '23
   Col.Width = 400
   Col.Caption = MapText("สถานะการตรวจสอบ")
      
   GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   If JobDocType = 1 Then
     If HeaderText <> "" Then
         Me.Caption = MapText(HeaderText)
         pnlHeader.Caption = Me.Caption
     Else
         Me.Caption = MapText("ใบสั่งผลิต")
         pnlHeader.Caption = MapText("ใบสั่งผลิต")
      End If
      
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
   Call InitNormalLabel(lblLotNo, MapText("LOT"))
   
   Call InitNormalLabel(lblInputAmount, MapText("ยอดใช้รวม"))
   Call InitNormalLabel(lblOutputAmount, MapText("ผลิตรวม"))

   Call InitNormalLabel(lblOrderBy, MapText("เรียงตาม"))
   Call InitNormalLabel(lblOrderType, MapText("เรียงจาก"))
   
   Call InitCheckBox(chkCommit, "ผลิตเสร็จ")
   
   Call txtJobNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtBatchNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   Call txtInputAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtOutputAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtInputAmount.Enabled = False
   txtOutputAmount.Enabled = False
   
   Call txtPartNo.SetKeySearch("PART_NO")
   
   Call InitCombo(cboJobProcess)
   Call InitCombo(cboOrderBy)
   Call InitCombo(cboOrderType)
   Call InitCombo(cboLotNo)
  
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
   cmdLock.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdUnlock.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdLockDate.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา (F5)"))
   Call InitMainButton(cmdClear, MapText("เคลียร์ (F4)"))
   Call InitMainButton(cmdOther, MapText("อื่น ๆ"))
   Call InitMainButton(cmdLock, MapText("ล็อค"))
   Call InitMainButton(cmdUnlock, MapText("ปลดล็อค"))
   Call InitMainButton(cmdLockDate, MapText("ล็อควันที่"))
  
    If ProcessID = 4 Then
      cmdLock.Visible = True
      cmdUnlock.Visible = True
      cmdLockDate.Visible = True
   Else
      cmdLock.Visible = False
      cmdUnlock.Visible = False
      cmdLockDate.Visible = False
   End If
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
   Set m_PartItems = New Collection
'   Set m_LoginTracking = New Collection
   Call EnableForm(Me, True)
End Sub



Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   RowBuffer.RowStyle = RowBuffer.Value(11)
End Sub
Private Function SplitText(str As String) As Boolean
SplitText = False
If Len(str) > 6 Then
    If Left(m_TempJob.VERIFY_NAME, 6) = "CANCEL" Then
      SplitText = True
    End If
End If
End Function
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
   If m_TempJob.LOCK_DOC_FLAG = "Y" Then
      Values(2) = m_TempJob.JOB_NO & " (" & IIf(m_TempJob.LOCK_DOC_FLAG = "Y", "L", "") & ")"
   Else
      Values(2) = m_TempJob.JOB_NO
   End If
   Values(3) = m_TempJob.JOB_DESC
   Values(4) = DateToStringExtEx2(m_TempJob.JOB_DATE)
   Values(5) = IIf(m_TempJob.FROM_BATCH_NO = -1, "", Format(m_TempJob.FROM_BATCH_NO, "000"))
   Values(6) = IIf(m_TempJob.TO_BATCH_NO = -1, "", Format(m_TempJob.TO_BATCH_NO, "000"))
   Values(7) = sortBatch(m_TempJob.BATCH_DETAIL)
   If ProcessID = 4 Then
      Values(8) = m_TempJob.TO_BATCH_NO - m_TempJob.FROM_BATCH_NO + 1
   Else
      Values(8) = m_TempJob.BATCH_NO
   End If
'   Values(8) = m_TempJob.TO_BATCH_NO - m_TempJob.FROM_BATCH_NO + 1 'm_TempJob.BATCH_NO
   Values(9) = IIf(m_TempJob.BATCH_TOTAL = -1, "", m_TempJob.BATCH_TOTAL)
   Values(10) = FormatNumber(m_TempJob.SUM_INPUT, 3)
   Values(11) = FormatNumber(m_TempJob.SUM_OUTPUT, 3)
   Values(12) = m_TempJob.PROCESS_NAME
   Values(13) = DateToStringExtEx2(m_TempJob.START_DATE)
   Values(14) = DateToStringExtEx2(m_TempJob.FINISH_DATE)
   
   If (m_TempJob.VERIFY_FLAG = "N" And SplitText(m_TempJob.VERIFY_NAME)) Or (m_TempJob.VERIFY_FLAG = "Y") Then
      Values(15) = m_TempJob.VERIFY_NAME
   Else
      Values(15) = ""
   End If
   
     Set Temp_LTK = GetObject("CLoginTracking", m_LoginTracking, Trim(m_TempJob.CREATE_BY), False)
        If Not Temp_LTK Is Nothing Then
            Values(16) = Temp_LTK.USER_NAME
         Else
             Values(16) = ""
         End If
         
         Set Temp_LTK = GetObject("CLoginTracking", m_LoginTracking, Trim(m_TempJob.MODIFY_BY), False)
        If Not Temp_LTK Is Nothing Then
            Values(17) = Temp_LTK.USER_NAME
         Else
             Values(17) = ""
         End If
   Values(18) = m_TempJob.NOTE
   Values(19) = m_TempJob.LONG_NAMER & " " & m_TempJob.LAST_NAMER
   Values(20) = m_TempJob.COMMIT_FLAG '17
   Values(21) = m_TempJob.PROCESS_ID
   Values(22) = m_TempJob.LOCK_DOC_FLAG
   Values(23) = m_TempJob.JOB_ID_REF
   Values(24) = m_TempJob.LONG_NAMEA & " " & m_TempJob.LAST_NAMEA
   Values(25) = m_TempJob.VERIFY_FLAG
'   Values(18) = m_TempJob.LONG_NAMER & " " & m_TempJob.LAST_NAMER
'   Values(19) = m_TempJob.COMMIT_FLAG '17
'   Values(20) = m_TempJob.PROCESS_ID
'   Values(21) = m_TempJob.LOCK_DOC_FLAG
'   Values(22) = m_TempJob.JOB_ID_REF
'   Values(23) = m_TempJob.LONG_NAMEA & " " & m_TempJob.LAST_NAMEA
'   Values(24) = m_TempJob.VERIFY_FLAG
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Function sortBatch(str As String) As String
   Dim strArr() As String
   Dim J As Long
   Dim Temp As Long
   Dim I As Long
      strArr = Split(str, ",")
      If UBound(strArr) > -1 Then
         For I = 0 To UBound(strArr)
           For J = I + 1 To UBound(strArr)
              If Val(strArr(J)) < Val(strArr(I)) Then
                 Temp = strArr(I)
                 strArr(I) = strArr(J)
                 strArr(J) = Temp
              End If
            Next J
         Next I
         
         For I = 0 To UBound(strArr)
           If I = 0 Then
            sortBatch = strArr(I)
           Else
             sortBatch = sortBatch & "," & strArr(I)
           End If
         Next I
      ElseIf Len(str) > 0 Then
         sortBatch = str
      End If
End Function


Private Sub cmdClear_Click()
   txtJobNo.Text = ""
   txtBatchNo.Text = ""
   txtPartNo.Text = ""
   cboOrderBy.ListIndex = -1
   cboOrderType.ListIndex = -1
   cboJobProcess.ListIndex = -1
   uctlJobDate.ShowDate = -1
   cboLotNo.ListIndex = -1
   chkCommit.Value = ssCBUnchecked
   
End Sub

Public Sub LoadRefDoc(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CInventoryDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CInventoryDoc
Dim I As Long

   Set D = New CInventoryDoc
   Set Rs = New ADODB.Recordset
   D.COMMIT_FLAG = "Y"
   D.INVENTORY_DOC_ID = -1
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CInventoryDoc
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.DOCUMENT_NO)
         C.ItemData(I) = TempData.INVENTORY_DOC_ID
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

Private Sub txtPartNo_Change()
Dim PartNo As String
    PartNo = Trim(txtPartNo.Text)
    Call LoadLotIdByPartItem(cboLotNo, Nothing, , , , , , 5, 1, 1, "I", Nothing, 1, , PartNo)
End Sub


