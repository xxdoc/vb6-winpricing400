VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditJob 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditJob.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8895
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   12720
      _ExtentX        =   22437
      _ExtentY        =   15690
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextLookup ucltApproveByLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   12
         Top             =   3180
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin VB.ComboBox cboJobProcess 
         Height          =   315
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2280
         Width           =   3100
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   120
         TabIndex        =   16
         Top             =   4900
         Width           =   11640
         _ExtentX        =   20532
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
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   2310
         Left            =   120
         TabIndex        =   17
         Top             =   5400
         Width           =   11640
         _ExtentX        =   20532
         _ExtentY        =   4075
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
         Column(1)       =   "frmAddEditJob.frx":27A2
         Column(2)       =   "frmAddEditJob.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditJob.frx":290E
         FormatStyle(2)  =   "frmAddEditJob.frx":2A6A
         FormatStyle(3)  =   "frmAddEditJob.frx":2B1A
         FormatStyle(4)  =   "frmAddEditJob.frx":2BCE
         FormatStyle(5)  =   "frmAddEditJob.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditJob.frx":2D5E
      End
      Begin prjFarmManagement.uctlTextBox txtJobNo 
         Height          =   435
         Left            =   1800
         TabIndex        =   0
         Top             =   990
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtJobDesc 
         Height          =   435
         Left            =   1800
         TabIndex        =   11
         Top             =   2730
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtBatchNo 
         Height          =   435
         Left            =   1800
         TabIndex        =   3
         Top             =   1440
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlFinishJob 
         Height          =   405
         Left            =   1800
         TabIndex        =   8
         Top             =   2310
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlDate uctlStartJob 
         Height          =   405
         Left            =   1800
         TabIndex        =   7
         Top             =   1890
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextLookup uctlResponseByLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   13
         Top             =   3600
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtInputAmount 
         Height          =   435
         Left            =   8490
         TabIndex        =   14
         Top             =   2700
         Width           =   1695
         _ExtentX        =   4683
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtOutputAmount 
         Height          =   435
         Left            =   8490
         TabIndex        =   15
         Top             =   3150
         Width           =   1695
         _ExtentX        =   4683
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlJobDate 
         Height          =   405
         Left            =   6840
         TabIndex        =   2
         Top             =   960
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtFromBatch 
         Height          =   435
         Left            =   3840
         TabIndex        =   4
         Top             =   1440
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   1720
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtToBatch 
         Height          =   435
         Left            =   5400
         TabIndex        =   5
         Top             =   1440
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   1720
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTotalBatch 
         Height          =   435
         Left            =   7320
         TabIndex        =   6
         Top             =   1440
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1720
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtNote 
         Height          =   435
         Left            =   1800
         TabIndex        =   46
         Top             =   4080
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin VB.Label lblNote 
         Alignment       =   1  'Right Justify
         Caption         =   "lblNote"
         Height          =   315
         Left            =   120
         TabIndex        =   47
         Top             =   4200
         Width           =   1605
      End
      Begin VB.Label lblTotalBatch 
         Alignment       =   1  'Right Justify
         Caption         =   "lblTotalBatch"
         Height          =   315
         Left            =   6000
         TabIndex        =   45
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblToBatch 
         Alignment       =   1  'Right Justify
         Caption         =   "lblToBatch"
         Height          =   315
         Left            =   4440
         TabIndex        =   44
         Top             =   1560
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblFromBatch 
         Alignment       =   1  'Right Justify
         Caption         =   "lblFromBatch"
         Height          =   315
         Left            =   2880
         TabIndex        =   43
         Top             =   1560
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblJobDate 
         Alignment       =   1  'Right Justify
         Caption         =   "lblJobDate"
         Height          =   315
         Left            =   5460
         TabIndex        =   42
         Top             =   1080
         Width           =   1305
      End
      Begin Threed.SSCommand cmdLock 
         Height          =   525
         Left            =   8520
         TabIndex        =   41
         Top             =   1680
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdUnlock 
         Height          =   525
         Left            =   9600
         TabIndex        =   40
         Top             =   1680
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label Label3 
         Caption         =   "lblFormulaNo"
         Height          =   315
         Left            =   10230
         TabIndex        =   39
         Top             =   2850
         Width           =   1305
      End
      Begin VB.Label Label4 
         Caption         =   "lblFormulaNo"
         Height          =   315
         Left            =   10230
         TabIndex        =   38
         Top             =   3300
         Width           =   1305
      End
      Begin VB.Label lblOutputAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblBatchNo"
         Height          =   315
         Left            =   7260
         TabIndex        =   37
         Top             =   3300
         Width           =   1125
      End
      Begin VB.Label lblInputAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblBatchNo"
         Height          =   315
         Left            =   7260
         TabIndex        =   36
         Top             =   2850
         Width           =   1125
      End
      Begin Threed.SSCommand cmdCalculate 
         Height          =   525
         Left            =   6840
         TabIndex        =   21
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJob.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAuto 
         Height          =   405
         Left            =   4470
         TabIndex        =   1
         Top             =   990
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJob.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   525
         Left            =   8490
         TabIndex        =   24
         Top             =   3630
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJob.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   525
         Left            =   10110
         TabIndex        =   25
         Top             =   3630
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblJobProcess 
         Alignment       =   1  'Right Justify
         Caption         =   "lblJobProcess"
         Height          =   315
         Left            =   5790
         TabIndex        =   35
         Top             =   2280
         Width           =   945
      End
      Begin VB.Label lblJobRes 
         Alignment       =   1  'Right Justify
         Caption         =   "lblJobRes"
         Height          =   315
         Left            =   120
         TabIndex        =   34
         Top             =   3720
         Width           =   1605
      End
      Begin VB.Label lblJobApp 
         Alignment       =   1  'Right Justify
         Caption         =   "lblJobApp"
         Height          =   315
         Left            =   420
         TabIndex        =   33
         Top             =   3300
         Width           =   1305
      End
      Begin VB.Label lblFinishJob 
         Alignment       =   1  'Right Justify
         Caption         =   "lblFinishJob"
         Height          =   315
         Left            =   60
         TabIndex        =   32
         Top             =   2430
         Width           =   1665
      End
      Begin VB.Label lblBatchNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblBatchNo"
         Height          =   315
         Left            =   150
         TabIndex        =   31
         Top             =   1590
         Width           =   1575
      End
      Begin VB.Label lblStartJob 
         Alignment       =   1  'Right Justify
         Caption         =   "lblStartJob"
         Height          =   315
         Left            =   420
         TabIndex        =   30
         Top             =   2010
         Width           =   1305
      End
      Begin Threed.SSCheck chkCommit 
         Height          =   435
         Left            =   6840
         TabIndex        =   9
         Top             =   1800
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblJobDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "lblJobDesc"
         Height          =   315
         Left            =   120
         TabIndex        =   29
         Top             =   2850
         Width           =   1605
      End
      Begin VB.Label lblJobNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblJobNo"
         Height          =   315
         Left            =   330
         TabIndex        =   28
         Top             =   1110
         Width           =   1395
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8490
         TabIndex        =   22
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJob.frx":3884
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10170
         TabIndex        =   23
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
         MouseIcon       =   "frmAddEditJob.frx":3B9E
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
         MouseIcon       =   "frmAddEditJob.frx":3EB8
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Job As CJob
Private m_Job2 As CJob
Private m_Jobs As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public JobDocType As Long
Public ProcessID As Long
Public mainText As String
Public TempCollection As Collection
Public m_CollLotExUse As Collection
Public m_CollPalletInLot As Collection
Public DOCUMENT_TYPE As Long
Public JobIdRef As Long

Private FileName As String
Private m_SumUnit As Double
Private m_Employees As Collection
Private m_FormulaID As Long
Private typeForm As Long
Private TempPDEdit As Collection
Private TempUserName As String
Private m_InventoryWhDocInput As CInventoryWHDoc

Private Sub EnableDisableButton(En As Boolean)
   If En Then
      If ShowMode = SHOW_EDIT Then
         cmdAdd.Enabled = (m_Job.COMMIT_FLAG = "N")
         cmdDelete.Enabled = (m_Job.COMMIT_FLAG = "N")
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

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
      
      m_Job.JOB_ID = ID
      m_Job.QueryFlag = 1
      If Not glbProduction.QueryJob(m_Job, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
        Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_Job.PopulateFromRS(1, m_Rs)

      m_FormulaID = m_Job.FORMULA_ID
      txtJobNo.Text = m_Job.JOB_NO
      txtJobDesc.Text = m_Job.JOB_DESC
      txtNote.Text = m_Job.NOTE
      uctlJobDate.ShowDate = m_Job.JOB_DATE
      txtBatchNo.Text = m_Job.BATCH_NO
      ucltApproveByLookup.MyCombo.ListIndex = IDToListIndex(ucltApproveByLookup.MyCombo, m_Job.APPROVED_BY)
      uctlResponseByLookup.MyCombo.ListIndex = IDToListIndex(uctlResponseByLookup.MyCombo, m_Job.RESPONSE_BY)
      uctlStartJob.ShowDate = m_Job.START_DATE
      uctlFinishJob.ShowDate = m_Job.FINISH_DATE
      cboJobProcess.ListIndex = IDToListIndex(cboJobProcess, m_Job.PROCESS_ID)
      chkCommit.Value = FlagToCheck(m_Job.COMMIT_FLAG)
      chkCommit.Enabled = (m_Job.COMMIT_FLAG = "N")
      Call EnableDisableButton(True)
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call TabStrip1_Click
   
   Call EnableForm(Me, True)
End Sub
Private Sub QueryData2(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
      m_Job.JOB_ID = ID
      m_Job.QueryFlag = 1
      If Not glbProduction.QueryJob2(m_Job, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
        Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_Job.PopulateFromRS(1, m_Rs)
      m_FormulaID = m_Job.FORMULA_ID
      txtJobNo.Text = m_Job.JOB_NO
      txtJobDesc.Text = m_Job.JOB_DESC
      txtNote.Text = m_Job.NOTE
      uctlJobDate.ShowDate = m_Job.JOB_DATE
      txtBatchNo.Text = m_Job.BATCH_NO
      ucltApproveByLookup.MyCombo.ListIndex = IDToListIndex(ucltApproveByLookup.MyCombo, m_Job.APPROVED_BY)
      uctlResponseByLookup.MyCombo.ListIndex = IDToListIndex(uctlResponseByLookup.MyCombo, m_Job.RESPONSE_BY)
      uctlStartJob.ShowDate = m_Job.START_DATE
      uctlFinishJob.ShowDate = m_Job.FINISH_DATE
      cboJobProcess.ListIndex = IDToListIndex(cboJobProcess, m_Job.PROCESS_ID)
      chkCommit.Value = FlagToCheck(m_Job.COMMIT_FLAG)
      chkCommit.Enabled = (m_Job.COMMIT_FLAG = "N")
      If ProcessID = 4 Then
         txtFromBatch.Text = m_Job.FROM_BATCH_NO
         txtToBatch.Text = m_Job.TO_BATCH_NO
         txtTotalBatch.Text = m_Job.BATCH_TOTAL
      End If
      Call EnableDisableButton(True)
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

Private Function GetJobPartItemID(Col As Collection) As CJobInput
Dim JO As CJobInput
   For Each JO In Col
      If JO.Flag <> "D" Then
         Set GetJobPartItemID = JO
      End If
   Next JO
End Function

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim Ivd As CInventoryDoc
Dim IvdWH As CInventoryWHDoc
Dim JO As CJobInput
   If Not VerifyAccessRight("PRODUCT_JOB_EDIT") Then
      Call EnableForm(Me, True)
      Exit Function
   End If
   If Not VerifyTextControl(lblJobNo, txtJobNo, False) Then
       Exit Function
   End If
   If Not VerifyTextControl(lblJobDesc, txtJobDesc, True) Then
    Exit Function
   End If
   If Not VerifyDate(lblJobDate, uctlJobDate, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblBatchNo, txtBatchNo, True) Then
    Exit Function
   End If

   If Not VerifyCombo(lblJobApp, ucltApproveByLookup.MyCombo, True) Then
      Exit Function
   End If
   If Not VerifyCombo(lblJobRes, uctlResponseByLookup.MyCombo, True) Then
      Exit Function
   End If
   
   If Not VerifyDate(lblStartJob, uctlStartJob, True) Then
     Exit Function
   End If
   
   If Not VerifyDate(lblFinishJob, uctlFinishJob, True) Then
     Exit Function
   End If
 
   If Not VerifyCombo(lblJobProcess, cboJobProcess, False) Then
      Exit Function
   End If
   
   If ProcessID = 4 Then
      If Not VerifyLockInventoryDate(uctlJobDate.ShowDate, m_Job.JOB_DATE) Then
         glbErrorLog.LocalErrorMsg = MapText("ไม่สามารถเปลี่ยนแปลงเอกสารตามวันที่เอกสารที่เลือกได้ กรุณาติดต่อผู้ดูแลระบบ หรือผู้มีสิทธิ์กำหนดวันที่เอกสารได้")
         glbErrorLog.ShowUserError
         Exit Function
      End If
   End If
      
   If Not CheckUniqueNs(JOB_NO, txtJobNo.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtJobNo.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Function
   End If
   
   If CountItem(m_Job.Outputs) <> 1 Then
      glbErrorLog.LocalErrorMsg = "ข้อมูลผลิตภัณฑ์ที่ได้จะต้องมีเพียงแค่ 1 รายการเท่านั้น"
      glbErrorLog.ShowUserError
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_Job.JOB_ID = ID
   m_Job.AddEditMode = ShowMode
   m_Job.JOB_NO = txtJobNo.Text
   m_Job.JOB_DESC = txtJobDesc.Text
   m_Job.NOTE = txtNote.Text
   m_Job.JOB_DATE = uctlJobDate.ShowDate
   m_Job.BATCH_NO = txtBatchNo.Text
   m_Job.APPROVED_BY = ucltApproveByLookup.MyCombo.ItemData(Minus2Zero(ucltApproveByLookup.MyCombo.ListIndex))
   m_Job.RESPONSE_BY = uctlResponseByLookup.MyCombo.ItemData(Minus2Zero(uctlResponseByLookup.MyCombo.ListIndex))
   m_Job.START_DATE = uctlStartJob.ShowDate
   m_Job.FINISH_DATE = uctlFinishJob.ShowDate
   m_Job.PROCESS_ID = cboJobProcess.ItemData(Minus2Zero(cboJobProcess.ListIndex))
   m_Job.COMMIT_FLAG = Check2Flag(chkCommit.Value)
   m_Job.JOB_DOC_TYPE = JobDocType
   m_Job.FORMULA_ID = m_FormulaID
   If m_Job.Outputs.Count > 0 Then
      Set JO = GetJobPartItemID(m_Job.Outputs)
      If Not (JO Is Nothing) Then
         m_Job.PART_ITEM_ID = JO.PART_ITEM_ID
         m_Job.STD_AMOUNT = JO.STD_AMOUNT
         m_Job.ACTUAL_AMOUNT = JO.TX_AMOUNT
      End If
   Else
      m_Job.PART_ITEM_ID = -1
      m_Job.STD_AMOUNT = 0
      m_Job.ACTUAL_AMOUNT = 0
   End If
   
   Call EnableForm(Me, False)
   
   Call PopulateGuiID(m_Job)
  
   If JobDocType = 1 Then
      ''COPY data to InventoryDoc
      Call glbDaily.Job2InventoryDoc(m_Job, Ivd, 1, 11)
      If (m_Job.COMMIT_FLAG = "Y") Then
         If m_Job.OLD_COMMIT_FLAG <> "Y" Then
            Call glbDaily.TriggerCommit(Ivd.ImportExports)
            If Not glbDaily.VerifyStockBalance(Ivd.ImportExports, glbErrorLog) Then
               Call EnableForm(Me, True)
               Exit Function
            End If
         End If
      End If
   End If
      
   ''insert data to InventoryDoc
   Call glbDaily.StartTransaction
   If JobDocType = 1 Then
      If Not glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         SaveData = False
         Call glbDaily.RollbackTransaction
         Call EnableForm(Me, True)
         Exit Function
      End If
      m_Job.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID
   Else
      m_Job.INVENTORY_DOC_ID = -1
   End If
    ''end insert data to InventoryDoc

   If Not glbProduction.AddEditJob(m_Job, IsOK, False, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call glbDaily.RollbackTransaction
      Call EnableForm(Me, True)
      Exit Function
   End If
   Call glbDaily.CommitTransaction
   
   If Not IsOK Then
      Call EnableForm(Me, True)
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   Call EnableForm(Me, True)
   SaveData = True
End Function
Private Function SaveData2() As Boolean
Dim IsOK As Boolean
Dim Ivd As CInventoryDoc
Dim IvdWH As CInventoryWHDoc
Dim IvdWHInput As CInventoryWHDoc
Dim JO As CJobInput
Dim J As Long
   If Not VerifyAccessRight("PRODUCT_JOB_EDIT") Then
      Call EnableForm(Me, True)
      Exit Function
   End If
   
   If Not VerifyTextControl(lblJobNo, txtJobNo, False) Then
       Exit Function
   End If
   
   If Not VerifyTextControl(lblJobDesc, txtJobDesc, True) Then
      Exit Function
   End If
   
   If Not VerifyDate(lblJobDate, uctlJobDate, False) Then
      Exit Function
   End If

   If Not VerifyTextControl(lblBatchNo, txtBatchNo, True) Then
     Exit Function
   End If

   If Not VerifyCombo(lblJobApp, ucltApproveByLookup.MyCombo, True) Then
      Exit Function
   End If
   
   If Not VerifyCombo(lblJobRes, uctlResponseByLookup.MyCombo, True) Then
      Exit Function
   End If
   
   If Not VerifyDate(lblStartJob, uctlStartJob, True) Then
     Exit Function
   End If
   
   If Not VerifyDate(lblFinishJob, uctlFinishJob, True) Then
     Exit Function
   End If
 
   If Not VerifyCombo(lblJobProcess, cboJobProcess, False) Then
      Exit Function
   End If
   
   If ProcessID = 4 Then
      If Not VerifyLockInventoryDate(uctlJobDate.ShowDate, m_Job.JOB_DATE) Then
         glbErrorLog.LocalErrorMsg = MapText("ไม่สามารถเปลี่ยนแปลงเอกสารตามวันที่เอกสารที่เลือกได้ กรุณาติดต่อผู้ดูแลระบบ หรือผู้มีสิทธิ์กำหนดวันที่เอกสารได้")
         glbErrorLog.ShowUserError
         Exit Function
      End If
   End If
      
   If Not CheckUniqueNs(JOB_NO, txtJobNo.Text, ID) Then
      Call EnableForm(Me, True)
      SaveData2 = True
      Exit Function
   End If
   
   If CountItem(m_Job.Outputs) > 1 Then
      glbErrorLog.LocalErrorMsg = "ข้อมูลผลิตภัณฑ์ที่ได้จะต้องมีเพียงแค่ 1 รายการเท่านั้น"
      glbErrorLog.ShowUserError
   End If
   
   If DOCUMENT_TYPE = 14 And m_Job.VERIFY_FLAG = "Y" Then
      glbErrorLog.LocalErrorMsg = "ข้อมูลรายการนี้มีการตรวจสอบจาก QC แล้ว หากต้องการแก้ไข ต้องยกเลิกการตรวจสอบและอนุมัติรายการนี้ใหม่อีกครั้ง"
      glbErrorLog.ShowUserError
       Exit Function
   End If
   
   
   'ตรวจสอบ Balance ระหว่าง รับเข้าจ่ายออก ของจำนวนอาหาร
      Dim SumInput As Double
      Dim SumOutput As Double
      Dim JIO As CJobInput
      Dim isWork As Boolean
   If DOCUMENT_TYPE = 14 Then 'บรรจุ bag
   '**** ปิดใช้งานไปก่อน เพราะ ตอนนี้ อาหารฝั่ง bulk ที่ออกจาก plc มีไม่พอดีกับ จำนวนที่แพ็คจริง จึง ให้รับ แบบน้ำหนักทั้งสอง ฝั่ง ไม่เท่ากัน ไปก่อน
'      isWork = True
'       For Each JIO In m_Job.Inputs
'         If JIO.PART_TYPE_ID = 21 Then
'            SumInput = SumInput + JIO.TX_AMOUNT
'         End If
'       Next JIO
'
'       For Each JIO In m_Job.Outputs
'         If JIO.PART_TYPE_ID = 10 Then
'            SumOutput = SumOutput + JIO.TX_AMOUNT
'         End If
'       Next JIO
    ElseIf DOCUMENT_TYPE = 17 And ProcessID = 6 Then  'rebag to bag
      isWork = True
       For Each JIO In m_Job.Inputs
         If JIO.PART_TYPE_ID = 10 And JIO.Flag <> "D" Then
            SumInput = SumInput + JIO.TX_AMOUNT
         End If
       Next JIO
       
       For Each JIO In m_Job.Outputs
         If JIO.PART_TYPE_ID = 10 And JIO.Flag <> "D" Then
            SumOutput = SumOutput + JIO.TX_AMOUNT
         End If
       Next JIO
   ElseIf (DOCUMENT_TYPE = 17 And ProcessID = 7) Or (DOCUMENT_TYPE = 18 And ProcessID = 7) Then  'rebag to bulk
      isWork = True
       For Each JIO In m_Job.Inputs
         If JIO.PART_TYPE_ID = 10 And JIO.Flag <> "D" Then
            SumInput = SumInput + JIO.TX_AMOUNT
         End If
       Next JIO
       
       For Each JIO In m_Job.Outputs
         If JIO.PART_TYPE_ID = 21 And JIO.Flag <> "D" Then
            SumOutput = SumOutput + JIO.TX_AMOUNT
         End If
       Next JIO
  ElseIf DOCUMENT_TYPE = 19 And ProcessID = 8 Then 'rebag to rm
      isWork = True
       For Each JIO In m_Job.Inputs
         If JIO.PART_TYPE_ID = 10 And JIO.Flag <> "D" Then
            SumInput = SumInput + JIO.TX_AMOUNT
         End If
       Next JIO
       
       For Each JIO In m_Job.Outputs
         If JIO.Flag <> "D" Then
            SumOutput = SumOutput + JIO.TX_AMOUNT
         End If
       Next JIO
   End If
   If isWork Then
       If SumInput <> SumOutput Then
              glbErrorLog.LocalErrorMsg = "จำนวนรับเข้า ไม่เท่ากับ จำนวนจ่ายออก กรุณาตรวจสอบอีกครั้ง"
              glbErrorLog.ShowUserError
              SaveData2 = False
              Exit Function
      End If
   End If
   
   If Not m_HasModify Then
      SaveData2 = True
      Exit Function
   End If
   
   m_Job.JOB_ID = ID
   m_Job.AddEditMode = ShowMode
   m_Job.JOB_NO = txtJobNo.Text
   m_Job.JOB_DESC = txtJobDesc.Text
   m_Job.NOTE = txtNote.Text
   m_Job.JOB_DATE = uctlJobDate.ShowDate
   m_Job.BATCH_NO = txtBatchNo.Text
   m_Job.APPROVED_BY = ucltApproveByLookup.MyCombo.ItemData(Minus2Zero(ucltApproveByLookup.MyCombo.ListIndex))
   m_Job.RESPONSE_BY = uctlResponseByLookup.MyCombo.ItemData(Minus2Zero(uctlResponseByLookup.MyCombo.ListIndex))
   m_Job.START_DATE = uctlStartJob.ShowDate
   m_Job.FINISH_DATE = uctlFinishJob.ShowDate
   m_Job.PROCESS_ID = cboJobProcess.ItemData(Minus2Zero(cboJobProcess.ListIndex))
   m_Job.COMMIT_FLAG = Check2Flag(chkCommit.Value)
   m_Job.JOB_DOC_TYPE = JobDocType
   m_Job.FORMULA_ID = m_FormulaID
   
   If DOCUMENT_TYPE = 14 And Not m_Job.VERIFY_FLAG = "Y" Then 'ถ้าเป็นการรับเข้า BAG จาก pack ให้เริ่มต้น เป้น N ทุกครั้ง เพื่อรอการตรวจสอบ
      m_Job.VERIFY_NAME = ""
      m_Job.VERIFY_FLAG = "N"
   End If
   
   If ProcessID = 4 Then
      If Val(txtFromBatch.Text) > Val(txtToBatch.Text) Then
         glbErrorLog.LocalErrorMsg = "ค่าจากแบท ต้องไม่มากกว่า ถึงแบท"
         glbErrorLog.ShowUserError
         SaveData2 = True
         Exit Function
      End If
      m_Job.FROM_BATCH_NO = Val(txtFromBatch.Text)
      m_Job.TO_BATCH_NO = Val(txtToBatch.Text)
      m_Job.BATCH_NO = Val(txtToBatch.Text) - Val(txtFromBatch.Text) + 1
      m_Job.BATCH_TOTAL = Val(txtTotalBatch.Text)
     For J = m_Job.FROM_BATCH_NO To m_Job.TO_BATCH_NO
         If J = m_Job.FROM_BATCH_NO Then
            m_Job.BATCH_DETAIL = J
         Else
            m_Job.BATCH_DETAIL = m_Job.BATCH_DETAIL & "," & J
         End If
      Next J
   End If
   If CountItem(m_Job.Outputs) > 0 Then
      Set JO = GetJobPartItemID(m_Job.Outputs)
      m_Job.PART_ITEM_ID = JO.PART_ITEM_ID
      m_Job.STD_AMOUNT = JO.STD_AMOUNT
      m_Job.ACTUAL_AMOUNT = JO.TX_AMOUNT
   Else
      m_Job.PART_ITEM_ID = -1
      m_Job.STD_AMOUNT = 0
      m_Job.ACTUAL_AMOUNT = 0
   End If
   
   Call EnableForm(Me, False)
   
   Call PopulateGuiID(m_Job)
  
   If JobDocType = 1 Then
      ''COPY data to InventoryDoc
      Call glbDaily.Job2InventoryDoc(m_Job, Ivd, 1, 11)
      If (m_Job.COMMIT_FLAG = "Y") Then
         If m_Job.OLD_COMMIT_FLAG <> "Y" Then
            Call glbDaily.TriggerCommit(Ivd.ImportExports)
            If Not glbDaily.VerifyStockBalance(Ivd.ImportExports, glbErrorLog) Then
               Call EnableForm(Me, True)
               Exit Function
            End If
         End If
      End If
   End If
   
   'COPY data to InventoryDocWH
   If Not m_Job.InventoryWhDoc Is Nothing Then
      If JobDocType = 1 Then
      Call glbDaily.Job2InventoryWhDoc(m_Job, IvdWH, 1, 11)
         If (m_Job.COMMIT_FLAG = "Y") Then
            If m_Job.OLD_COMMIT_FLAG <> "Y" Then
            Call glbDaily.TriggerCommit(IvdWH.C_LotItemsWH)
               If Not glbDaily.VerifyStockBalance(IvdWH.C_LotItemsWH, glbErrorLog) Then
               Call EnableForm(Me, True)
               Exit Function
               End If
            End If
         End If
      End If
   End If
   
    'COPY data to InventoryDocWHInput
   If Not m_Job.InventoryWhDocInput Is Nothing Then
      If JobDocType = 1 Then
      Call glbDaily.Job2InventoryWhDocInput(m_Job, IvdWHInput, 1, 11)
         If (m_Job.COMMIT_FLAG = "Y") Then
            If m_Job.OLD_COMMIT_FLAG <> "Y" Then
            Call glbDaily.TriggerCommit(IvdWHInput.C_LotItemsWH)
               If Not glbDaily.VerifyStockBalance(IvdWHInput.C_LotItemsWH, glbErrorLog) Then
               Call EnableForm(Me, True)
               Exit Function
               End If
            End If
         End If
      End If
   End If
      
   ''insert data to InventoryDoc
   Call glbDaily.StartTransaction
   If JobDocType = 1 Then
      If Not glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         SaveData2 = False
         Call glbDaily.RollbackTransaction
         Call EnableForm(Me, True)
         Exit Function
      End If
      m_Job.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID
   Else
      m_Job.INVENTORY_DOC_ID = -1
   End If
    ''end insert data to InventoryDoc

'     ''insert data to InventoryWhDoc
   If Not IvdWH Is Nothing Then
      If JobDocType = 1 Then
         If Not glbDaily.AddEditInventoryWhDoc(IvdWH, IsOK, False, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         SaveData2 = False
         Call glbDaily.RollbackTransaction
         Call EnableForm(Me, True)
         Exit Function
         End If
         If IvdWH.Flag = "D" Then
            m_Job.INVENTORY_WH_DOC_ID = -1
         Else
            m_Job.INVENTORY_WH_DOC_ID = IvdWH.INVENTORY_WH_DOC_ID
         End If
      Else
      m_Job.INVENTORY_WH_DOC_ID = -1
      End If
   End If
'   ''End insert data to InventoryWhDoc

'     ''insert data to InventoryWhDocInput
   If Not IvdWHInput Is Nothing Then
      If JobDocType = 1 Then
         If Not glbDaily.AddEditInventoryWhDoc(IvdWHInput, IsOK, False, glbErrorLog) Then
            glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
            SaveData2 = False
            Call glbDaily.RollbackTransaction
            Call EnableForm(Me, True)
            Exit Function
         End If
         
         If IvdWHInput.Flag = "D" Then
            m_Job.INVENTORY_WH_DOC_ID_INPUT = -1
         Else
            m_Job.INVENTORY_WH_DOC_ID_INPUT = IvdWHInput.INVENTORY_WH_DOC_ID
         End If
      Else
      m_Job.INVENTORY_WH_DOC_ID_INPUT = -1
      End If
      Call glbDaily.CommitTransaction
      Call glbDaily.StartTransaction
   End If
'   ''End insert data to InventoryWhDocInput

'แก้ไข pallet กรณีที่มีการแก้ไขเบอร์อาหารแล้ว pallet ทับกัน
   If Not TempPDEdit Is Nothing Then
      Dim LIW As CLotItemWH
      Dim PD As CPalletDoc
      Dim PrevKey As Long
     
      Set PD = New CPalletDoc
      For Each PD In TempPDEdit
          If PD.TX_TYPE = "E" Then 'หากมีข้อมูลการจ่ายออกแล้ว ก็ให้ เอา part_item_id ไป update ที่รายการจ่ายออกด้วย
            If PrevKey <> PD.LOT_ITEM_WH_ID Then
                PrevKey = PD.LOT_ITEM_WH_ID
                Set LIW = New CLotItemWH
                LIW.AddEditMode = SHOW_EDIT
                LIW.TX_TYPE = PD.TX_TYPE
                LIW.PART_ITEM_ID = PD.PART_ITEM_ID
                LIW.LOT_ITEM_WH_ID = PD.LOT_ITEM_WH_ID
                Call LIW.UpdatePartItemIdInLotItemWh
            End If
         End If
         PD.AddEditMode = SHOW_EDIT
         Call PD.AddEditPalletDocNo
      Next PD
   End If
  'end แก้ไข pallet กรณีที่มีการแก้ไขเบอร์อาหารแล้ว pallet ทับกัน
   
   If Not glbProduction.AddEditJob(m_Job, IsOK, False, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData2 = False
      Call glbDaily.RollbackTransaction
      Call EnableForm(Me, True)
      Exit Function
   End If
   Call glbDaily.CommitTransaction
   
   If Not IvdWHInput Is Nothing Then
      'ตรวจสอบ Stock หลังจาก update ใหม่ เพื่อเปลี่ยนสถานะ ของ Out Stock Flag
   Dim m_LotItemWh As CLotItemWH
   Dim LTD As CLotDoc
   Dim DocType As Long
     If DOCUMENT_TYPE = 13 Then
         DocType = 2000
     ElseIf DOCUMENT_TYPE = 14 Or DOCUMENT_TYPE = 19 Then
         DocType = 2001
     End If
      For Each m_LotItemWh In IvdWHInput.C_LotItemsWH
        For Each LTD In m_LotItemWh.C_LotDoc
            Call LoadLotInPartIemAmount(Nothing, Nothing, , , , , m_LotItemWh.PART_ITEM_ID, 2, 1, 1, "I", m_LotItemWh.C_LotDoc, , DocType, LTD.Flag)
       Next LTD
      Next m_LotItemWh
   End If
   
   
   If Not IsOK Then
      Call EnableForm(Me, True)
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   Call EnableForm(Me, True)
   SaveData2 = True
End Function
Private Sub cboJobProcess_Change()
m_HasModify = True
End Sub

Private Sub cboJobProcess_Click()
   m_HasModify = True
End Sub

Private Sub cboJobRef_Change()
m_HasModify = True
End Sub

Private Sub cboJobRef_Click()
m_HasModify = True
End Sub

Private Sub cboLotNo_Click()
m_HasModify = True
End Sub

Private Sub cboJobProcess_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
   KeyAscii = 0
End Sub

Private Sub chkCommit_Click(Value As Integer)
m_HasModify = True
End Sub

Public Sub RefreshGrid()
   GridEX1.ItemCount = CountItem(m_Job.Verifies)
   GridEX1.Rebind
End Sub

Private Sub chkCommit_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
KeyAscii = 0
End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu
Dim ID As Long
   If Not cmdAdd.Enabled Then
      Exit Sub
   End If

   OKClick = False
    If TabStrip1.SelectedItem.Index = 1 Then
      Set oMenu = New cPopupMenu
     If DOCUMENT_TYPE = 13 Then
        lMenuChosen = oMenu.Popup("เพิ่มรายการใหม่ทั่วไป", "-", "เลือกจากสูตร")
     ElseIf DOCUMENT_TYPE = 14 Then
'       lMenuChosen = oMenu.Popup("เพิ่มรายการใหม่ทั่วไป", "-", "เลือกจากสูตร", "-", "เพิ่มรายการใหม่ BULK", "-")
       lMenuChosen = oMenu.Popup("เพิ่มถุงบรรจุ", "-", "-", "-", "เพิ่มเบอร์อาหาร BULK", "-")
     ElseIf ProcessID = 6 And DOCUMENT_TYPE = 17 Then
       lMenuChosen = oMenu.Popup("เพิ่มรายการใหม่ทั่วไป", "-", "เลือกจากสูตร", "-", "เพิ่มรายการใหม่ RE-BAG -> BAG", "-")
     ElseIf DOCUMENT_TYPE = 18 Or (ProcessID = 7 And DOCUMENT_TYPE = 17) Then
       lMenuChosen = oMenu.Popup("เพิ่มรายการใหม่ทั่วไป", "-", "เลือกจากสูตร", "-", "เพิ่มรายการใหม่ RE-BAG -> BULK", "-")
       If lMenuChosen = 5 Then
         DOCUMENT_TYPE = 17 'ถ้าเป็นการป้อนวัตถุดิบที่ใช้ใน tab1 ให้เปลี่ยน type จาก 18 เป็น 17
         lMenuChosen = 7
       End If
   ElseIf ProcessID = 8 And DOCUMENT_TYPE = 19 Then
       lMenuChosen = oMenu.Popup("เพิ่มรายการใหม่ทั่วไป", "-", "เลือกจากสูตร", "-", "เพิ่มรายการใหม่ RE-BAG -> RM and Other", "-")
     Else
      lMenuChosen = oMenu.Popup("เพิ่มรายการใหม่ทั่วไป", "-", "เลือกจากสูตร")
     End If
'      lMenuChosen = oMenu.AddMenu(glbGuiConfigs.LoadAddJopProcessItems)
      If lMenuChosen = 0 Then
         Exit Sub
      End If
     
     If lMenuChosen = 1 Then
        typeForm = 1 'ให้แสดง Form ธรรมดา
        Set frmAddEditJobInput.TempCollection = m_Job.Inputs
         frmAddEditJobInput.typeForm = typeForm
         frmAddEditJobInput.ParentShowMode = ShowMode
         frmAddEditJobInput.ShowMode = SHOW_ADD
         frmAddEditJobInput.ProcessID = ProcessID
         frmAddEditJobInput.DOCUMENT_TYPE = DOCUMENT_TYPE
         Set frmAddEditJobInput.ParentForm = Me
         If DOCUMENT_TYPE = 14 Then
            frmAddEditJobInput.HeaderText = MapText("เพิ่มถุงบรรจุ")
            frmAddEditJobInput.TYPE_LIST_RM = lMenuChosen
         Else
            frmAddEditJobInput.HeaderText = MapText("เพิ่มวัตถุดิบ")
         End If
         
         Load frmAddEditJobInput
         frmAddEditJobInput.Show 1
   
         OKClick = frmAddEditJobInput.OKClick
   
         Unload frmAddEditJobInput
         Set frmAddEditJobInput = Nothing
      ElseIf lMenuChosen = 3 Then
          If ProcessID = 5 Then
            typeForm = 1 'ให้แสดง Form ธรรมดา
            GridEX1.MoveFirst
            Set frmFormulaSelect.Job = m_Job
            frmFormulaSelect.ID = ID
            frmFormulaSelect.FORMULA_ID = m_FormulaID
            frmFormulaSelect.ParentShowMode = ShowMode
            
            If Val(GridEX1.Value(1)) > 0 Then
               frmFormulaSelect.ShowMode = SHOW_EDIT
            Else
               frmFormulaSelect.ShowMode = SHOW_ADD
            End If
            frmFormulaSelect.HeaderText = MapText("เพิ่มวัตถุดิบจากสูตร")
            Load frmFormulaSelect
            frmFormulaSelect.Show 1
      
            OKClick = frmFormulaSelect.OKClick
            If OKClick Then
               m_FormulaID = frmFormulaSelect.FORMULA_ID
            End If
            
            Unload frmFormulaSelect
            Set frmFormulaSelect = Nothing
          Else
            typeForm = 2 'ให้แสดง Form แบบใหม่ ที่มี InventoryWh เข้ามาเกี่ยวข้องแล้ว
            GridEX1.MoveFirst
            Set frmFormulaSelectWh.Job = m_Job
            frmFormulaSelectWh.ID = ID
            frmFormulaSelectWh.FORMULA_ID = m_FormulaID
            frmFormulaSelectWh.ParentShowMode = ShowMode
            
            frmFormulaSelectWh.StartJob = uctlStartJob.ShowDate
            If Val(GridEX1.Value(1)) > 0 Then
               frmFormulaSelectWh.ShowMode = SHOW_EDIT
            Else
               frmFormulaSelectWh.ShowMode = SHOW_ADD
            End If
            frmFormulaSelectWh.HeaderText = MapText("เพิ่มวัตถุดิบจากสูตร")
            Load frmFormulaSelectWh
            frmFormulaSelectWh.Show 1
      
            OKClick = frmFormulaSelectWh.OKClick
            If OKClick Then
               m_FormulaID = frmFormulaSelectWh.FORMULA_ID
            End If
            
            Unload frmFormulaSelectWh
            Set frmFormulaSelectWh = Nothing
         End If
      ElseIf lMenuChosen = 5 Or lMenuChosen = 7 Then
         typeForm = 2 'ให้แสดง Form แบบใหม่ ที่มี InventoryWh เข้ามาเกี่ยวข้องแล้ว
        Set frmAddEditJobInput.TempCollection = m_Job.Inputs
        If m_Job.InventoryWhDocInput Is Nothing Then
          Set m_Job.InventoryWhDocInput = New Collection
          Set m_InventoryWhDocInput = New CInventoryWHDoc
          m_InventoryWhDocInput.AddEditMode = SHOW_ADD
         m_InventoryWhDocInput.Flag = "A"
          Call m_Job.InventoryWhDocInput.add(m_InventoryWhDocInput)
        End If
        
        If m_Job.InventoryWhDocInput.Item(1).C_LotItemsWH.Count > 0 Then
           If m_Job.InventoryWhDocInput.Item(1).C_LotItemsWH.Item(1).Flag = "D" Then
               glbErrorLog.LocalErrorMsg = "กรุณาบันทึกข้อมูลก่อน"
               glbErrorLog.ShowUserError
               Exit Sub
           Else
            glbErrorLog.LocalErrorMsg = "ข้อมูลผลิตภัณฑ์ที่ได้จะต้องมีเพียงแค่ 1 รายการเท่านั้น"
            glbErrorLog.ShowUserError
            Exit Sub
         End If
        End If
        Set frmAddEditJobInput.m_InventoryWhDocInput = m_Job.InventoryWhDocInput.Item(1)
        frmAddEditJobInput.typeForm = typeForm
         frmAddEditJobInput.ParentShowMode = ShowMode
         frmAddEditJobInput.ShowMode = SHOW_ADD
         frmAddEditJobInput.ProcessID = ProcessID
         frmAddEditJobInput.DOCUMENT_DATE = uctlJobDate.ShowDate
         Set frmAddEditJobInput.ParentForm = Me
         If DOCUMENT_TYPE = 14 Then
            frmAddEditJobInput.PartType = 21
            frmAddEditJobInput.DOCUMENT_TYPE = 18
            frmAddEditJobInput.LOCATION_ID = 110
            frmAddEditJobInput.HeaderText = MapText("เพิ่มเบอร์อาหาร BULK")
            frmAddEditJobInput.TYPE_LIST_RM = lMenuChosen
         ElseIf DOCUMENT_TYPE = 19 Then
            frmAddEditJobInput.PartType = 10
            frmAddEditJobInput.DOCUMENT_TYPE = 19
            frmAddEditJobInput.LOCATION_ID = 109
            frmAddEditJobInput.HeaderText = MapText("เพิ่มวัตถุดิบ")
         Else
            frmAddEditJobInput.PartType = 10
            frmAddEditJobInput.DOCUMENT_TYPE = DOCUMENT_TYPE
            frmAddEditJobInput.LOCATION_ID = 109
            frmAddEditJobInput.HeaderText = MapText("เพิ่มวัตถุดิบ")
         End If
         
         
         
         Load frmAddEditJobInput
         frmAddEditJobInput.Show 1
   
         OKClick = frmAddEditJobInput.OKClick
   
         Unload frmAddEditJobInput
         Set frmAddEditJobInput = Nothing
      End If
      
      If OKClick Then
         Call CalculateTotalRatio
         
         GridEX1.ItemCount = CountItem(m_Job.Inputs)
         GridEX1.Rebind
      End If
 ElseIf TabStrip1.SelectedItem.Index = 2 Then
   If Not VerifyCombo(lblJobProcess, cboJobProcess, False) Then
      Exit Sub
   End If
   
   If Not VerifyDate(lblStartJob, uctlStartJob, False) Then
      Exit Sub
   End If
   
    ProcessID = cboJobProcess.ItemData(Minus2Zero(cboJobProcess.ListIndex))
    If ProcessID = 2 Or ProcessID = 6 Then 'ถ้าเป็น Process บรรจุ BAG
      If CountItem(m_Job.Outputs) = 1 Then
         glbErrorLog.LocalErrorMsg = "ข้อมูลผลิตภัณฑ์ที่ได้จะต้องมีเพียงแค่ 1 รายการเท่านั้น"
         glbErrorLog.ShowUserError
         Exit Sub
      End If
      
      'กรณีตัดแตก
      If DOCUMENT_TYPE = 17 And ProcessID = 6 Then
         Set oMenu = New cPopupMenu
         lMenuChosen = oMenu.Popup("ถ่ายถุงทั่วไป", "-", "ตัดแตก", "-", "รับจากตัดแตก")
         If lMenuChosen = 0 Then
            Exit Sub
         End If
      ElseIf DOCUMENT_TYPE = 14 And ProcessID = 2 Then 'กรณีรับเข้าหน้าแพ็ค
         lMenuChosen = -1
      End If
      
         If lMenuChosen = 1 Or lMenuChosen = 3 Or lMenuChosen = 5 Then
             'COPY data to InventoryDocWHInput
            If Not m_Job.InventoryWhDocInput Is Nothing Then
               Call glbDaily.CopyIWDInputToIWD(m_Job, True, glbErrorLog)
            Else
               glbErrorLog.LocalErrorMsg = "กรุณาป้อนวัตุดิบที่ใช้ก่อน"
               glbErrorLog.ShowUserError
               Exit Sub
            End If
         End If
         
           If lMenuChosen = 5 Then
              If m_Job.InventoryWhDocInput.Item(1).C_LotItemsWH.Item(1).LOCATION_ID <> 78 Then
                 glbErrorLog.LocalErrorMsg = MapText("กรุณาเลือกข้อมูลจาก โกดังอาหารชั่วคราว")
                 glbErrorLog.ShowUserError
                 Exit Sub
              End If
            End If
         
              typeForm = 2
              Set frmAddEditJobOutputEx2.TempCollection = m_Job.Outputs
              Set m_Job.InventoryWhDoc = New Collection
              Set frmAddEditJobOutputEx2.TempCollection2 = m_Job.InventoryWhDoc
              Set frmAddEditJobOutputEx2.TempCollection5 = m_Job.tempIWDInput
              frmAddEditJobOutputEx2.typeInput = lMenuChosen
              frmAddEditJobOutputEx2.COMMIT_FLAG = m_Job.OLD_COMMIT_FLAG
              frmAddEditJobOutputEx2.Header = Me.HeaderText
              frmAddEditJobOutputEx2.StartJob = uctlStartJob.ShowDate
              frmAddEditJobOutputEx2.StopJob = uctlFinishJob.ShowDate
              frmAddEditJobOutputEx2.PartType = 10
               frmAddEditJobOutputEx2.ParentShowMode = ShowMode
               frmAddEditJobOutputEx2.ShowMode = SHOW_ADD
               frmAddEditJobOutputEx2.ID = 1
              frmAddEditJobOutputEx2.HeaderText = MapText("เพิ่มผลิตภัณฑ์ที่ได้")
               Load frmAddEditJobOutputEx2
               frmAddEditJobOutputEx2.Show 1
         
               OKClick = frmAddEditJobOutputEx2.OKClick
                 
               Unload frmAddEditJobOutputEx2
               Set frmAddEditJobOutputEx2 = Nothing
'               End If
'         End If
      
      

    ElseIf ProcessID = 4 Or ProcessID = 7 Then  'ถ้าเป็น Process บรรจุลง Bulk
      If CountItem(m_Job.Outputs) = 1 Then
         glbErrorLog.LocalErrorMsg = "ข้อมูลผลิตภัณฑ์ที่ได้จะต้องมีเพียงแค่ 1 รายการเท่านั้น"
         glbErrorLog.ShowUserError
         Exit Sub
      End If
      
      If ProcessID = 7 Then   'ถ้าเป็น Process บรรจุลง Bulk
       'COPY data to InventoryDocWHInput
         If Not m_Job.InventoryWhDocInput Is Nothing Then
            Call glbDaily.CopyIWDInputToIWD(m_Job, True, glbErrorLog)
          Else
            glbErrorLog.LocalErrorMsg = "กรุณาป้อนวัตุดิบที่ใช้ก่อน"
            glbErrorLog.ShowUserError
            Exit Sub
         End If
      End If

         
     typeForm = 2
     Set frmAddEditJobOutputEx4.TempCollection = m_Job.Outputs
      Set m_Job.InventoryWhDoc = New Collection
     Set frmAddEditJobOutputEx4.TempCollection2 = m_Job.InventoryWhDoc
     Set frmAddEditJobOutputEx4.TempCollection5 = m_Job.tempIWDInput
     frmAddEditJobOutputEx4.typeInput = 1
     frmAddEditJobOutputEx4.DocumentType = DOCUMENT_TYPE
     frmAddEditJobOutputEx4.COMMIT_FLAG = m_Job.OLD_COMMIT_FLAG
     frmAddEditJobOutputEx4.ID = ID
     frmAddEditJobOutputEx4.JobIdRef = JobIdRef
     frmAddEditJobOutputEx4.StartJob = uctlStartJob.ShowDate
     frmAddEditJobOutputEx4.StopJob = uctlFinishJob.ShowDate
     frmAddEditJobOutputEx4.PartType = 21
      frmAddEditJobOutputEx4.ParentShowMode = ShowMode
     frmAddEditJobOutputEx4.ShowMode = SHOW_ADD
     frmAddEditJobOutputEx4.HeaderText = MapText("แก้ไขผลผลิต")
      Load frmAddEditJobOutputEx4
      frmAddEditJobOutputEx4.Show 1

      OKClick = frmAddEditJobOutputEx4.OKClick

      Unload frmAddEditJobOutputEx4
      Set frmAddEditJobOutputEx4 = Nothing
   Else
      If ProcessID = 8 Then
       'COPY data to InventoryDocWHInput
            If Not m_Job.InventoryWhDocInput Is Nothing Then
               Call glbDaily.CopyIWDInputToIWD(m_Job, True, glbErrorLog)
             Else
               glbErrorLog.LocalErrorMsg = "กรุณาป้อนวัตุดิบที่ใช้ก่อน"
               glbErrorLog.ShowUserError
               Exit Sub
            End If
            frmAddEditJobOutputEx.typeInput = 1
            Set frmAddEditJobOutputEx.TempCollection5 = m_Job.tempIWDInput
      End If
 
      Set frmAddEditJobOutputEx.TempCollection = m_Job.Outputs
      frmAddEditJobOutputEx.ParentShowMode = ShowMode
      frmAddEditJobOutputEx.ShowMode = SHOW_ADD
      frmAddEditJobOutputEx.HeaderText = MapText("เพิ่มผลผลิต")
      Load frmAddEditJobOutputEx
      frmAddEditJobOutputEx.Show 1

      OKClick = frmAddEditJobOutputEx.OKClick

      Unload frmAddEditJobOutputEx
      Set frmAddEditJobOutputEx = Nothing
   End If
      If OKClick Then
         Call CalculateTotalRatio
         
         GridEX1.ItemCount = CountItem(m_Job.Outputs)
         GridEX1.Rebind
      End If
  
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
     Set frmAddEditJobPeople.TempCollection = m_Job.Peoples
      frmAddEditJobPeople.ParentShowMode = ShowMode
      frmAddEditJobPeople.ShowMode = SHOW_ADD
      frmAddEditJobPeople.HeaderText = MapText("เพิ่มแรงงาน")
      Load frmAddEditJobPeople
      frmAddEditJobPeople.Show 1

      OKClick = frmAddEditJobPeople.OKClick

      Unload frmAddEditJobPeople
      Set frmAddEditJobPeople = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Job.Peoples)
         GridEX1.Rebind
      End If
      
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
     Set frmAddEditJobMachineEx.TempCollection = m_Job.Machines
      frmAddEditJobMachineEx.ParentShowMode = ShowMode
      frmAddEditJobMachineEx.ShowMode = SHOW_ADD
      frmAddEditJobMachineEx.HeaderText = MapText("เพิ่มเครื่องจักรที่ใช้")
      Load frmAddEditJobMachineEx
      frmAddEditJobMachineEx.Show 1

      OKClick = frmAddEditJobMachineEx.OKClick

      Unload frmAddEditJobMachineEx
      Set frmAddEditJobMachineEx = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Job.Machines)
         GridEX1.Rebind
      End If
  ElseIf TabStrip1.SelectedItem.Index = 5 Then
     Set frmAddEditJobParameter.TempCollection = m_Job.Parameters
     If cboJobProcess.ItemData(Minus2Zero(cboJobProcess.ListIndex)) <= 0 Then
      glbErrorLog.LocalErrorMsg = "กรุณากรอก ข้อมูล โปรเซส ให้ครบถ้วน"
       glbErrorLog.ShowUserError
       Exit Sub
      End If
     frmAddEditJobParameter.Process = cboJobProcess.ItemData(Minus2Zero(cboJobProcess.ListIndex))
      frmAddEditJobParameter.ParentShowMode = ShowMode
      frmAddEditJobParameter.ShowMode = SHOW_ADD
      frmAddEditJobParameter.HeaderText = MapText("เพิ่มพารามิเตอร์ที่ใช้")
      Load frmAddEditJobParameter
      frmAddEditJobParameter.Show 1

      OKClick = frmAddEditJobParameter.OKClick

      Unload frmAddEditJobParameter
      Set frmAddEditJobParameter = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Job.Parameters)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 6 Then
     Set frmVerifyPartItemEx.TempCollection = m_Job.Verifies
     Set frmVerifyPartItemEx.TempCollection2 = m_Job.Inputs
     Set frmVerifyPartItemEx.Inputs = m_Job.Inputs
      frmVerifyPartItemEx.ParentShowMode = ShowMode
      Set frmVerifyPartItemEx.ParentForm = Me
      frmVerifyPartItemEx.ShowMode = SHOW_ADD
      frmVerifyPartItemEx.HeaderText = MapText("ตรวจสอบการใช้วัตถุดิบ")
      Load frmVerifyPartItemEx
      frmVerifyPartItemEx.Show 1

      OKClick = frmVerifyPartItemEx.OKClick

      Unload frmVerifyPartItemEx
      Set frmVerifyPartItemEx = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Job.Verifies)
         GridEX1.Rebind
      End If
   End If

   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdAdd_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
   KeyAscii = 0
End Sub

Private Sub cmdAuto_Click()
Dim No As String
   If Trim(txtJobNo.Text) = "" Then
      If JobDocType = 1 Then
         If DOCUMENT_TYPE = 17 And ProcessID = 6 Then
            Call glbDatabaseMngr.GenerateNumber(RE_BAG_JOBPLAN_NUMBER, No, glbErrorLog)
         ElseIf DOCUMENT_TYPE = 18 And ProcessID = 7 Then
            Call glbDatabaseMngr.GenerateNumber(RE_BULK_JOBPLAN_NUMBER, No, glbErrorLog)
         ElseIf DOCUMENT_TYPE = 17 And ProcessID = 7 Then ' กรณีพิเศษ
            Call glbDatabaseMngr.GenerateNumber(RE_BULK_JOBPLAN_NUMBER, No, glbErrorLog)
         ElseIf DOCUMENT_TYPE = 13 And ProcessID = 4 Then ' Create bulk
            Call glbDatabaseMngr.GenerateNumber(BULK_JOBPLAN_NUMBER, No, glbErrorLog)
         ElseIf DOCUMENT_TYPE = 19 And ProcessID = 8 Then ' Create RM
            Call glbDatabaseMngr.GenerateNumber(RE_BAG_RM_JOBPLAN_NUMBER, No, glbErrorLog)
         Else
            Call glbDatabaseMngr.GenerateNumber(JOBPLAN_NUMBER, No, glbErrorLog)
         End If
         txtJobNo.Text = No
      ElseIf JobDocType = 2 Then
         Call glbDatabaseMngr.GenerateNumber(ESTIMATE_NUMBER, No, glbErrorLog)
         txtJobNo.Text = No
      End If
   End If
End Sub

Private Sub CalculatePrice(PriceType As Long)
Dim D As CJobInput
Dim Sum As Double
Dim SumMarkup As Double
Dim FifoPrice As Double
Dim IsOK As Boolean
Dim LastPMC As Double
Dim AvgPMC As Double
Dim PL As CPartLocation
Dim TempRs As ADODB.Recordset
Dim iCount As Long

   Call EnableForm(Me, False)

   SumMarkup = 0
   Sum = 0

   For Each D In m_Job.Inputs
      If D.FROM_FORMULA > 0 Then
         Call glbProduction.GetCalculatedPrice(D.FROM_FORMULA, AvgPMC, PriceType, D.TX_AMOUNT, IsOK, glbErrorLog)
      Else
         AvgPMC = D.INCLUDE_UNIT_PRICE  'glbProduction.GetWeightedAvgPrice(D.PART_ITEM_ID, D.LOCATION_ID, D.TX_AMOUNT, PriceType)
      End If

      If D.Flag <> "D" Then
         D.AVG_PRICE = AvgPMC
         If D.Flag <> "A" Then
            D.Flag = "E"
         End If
      End If
   Next D

   For Each D In m_Job.Outputs
      If D.FROM_FORMULA > 0 Then
         Call glbProduction.GetCalculatedPrice(D.FROM_FORMULA, AvgPMC, PriceType, D.TX_AMOUNT, IsOK, glbErrorLog)
      Else
         AvgPMC = glbProduction.GetWeightedAvgPrice(D.PART_ITEM_ID, D.LOCATION_ID, D.TX_AMOUNT, PriceType)
      End If

      If D.Flag <> "D" Then
         D.AVG_PRICE = AvgPMC
         If D.Flag <> "A" Then
            D.Flag = "E"
         End If
      End If
   Next D

   Call TabStrip1_Click

   m_HasModify = True
   Call EnableForm(Me, True)
End Sub
Private Sub cmdAuto_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
  KeyAscii = 0
End Sub

Private Sub cmdCalculate_Click()
Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
Dim J As CJob

   Set oMenu = New cPopupMenu
   If DOCUMENT_TYPE = 14 Then
      lMenuChosen = oMenu.Popup("ดูข้อมูลสูตร", "-", "ปรับสูตร/ปรับปริมาณใหม่", "-", "QC ตรวจสอบให้ผ่านได้", "-", "QC ยกเลิกการตรวจสอบ")
   Else
      lMenuChosen = oMenu.Popup("ดูข้อมูลสูตร", "-", "ปรับสูตร/ปรับปริมาณใหม่")
   End If
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
   If lMenuChosen = 1 Then
      If m_FormulaID <= 0 Then
         glbErrorLog.LocalErrorMsg = "ไม่มีข้อมูลสูตรของใบสั่งผลิตนี้"
         glbErrorLog.ShowUserError
         Set oMenu = Nothing
         Exit Sub
      End If
      
      frmAddEditFormulaMain.ID = m_FormulaID
      frmAddEditFormulaMain.HeaderText = "ดูข้อมูลสูตร"
      frmAddEditFormulaMain.ShowMode = SHOW_VIEW_ONLY
      Load frmAddEditFormulaMain
      frmAddEditFormulaMain.Show 1
      
      Unload frmAddEditFormulaMain
      Set frmAddEditFormulaMain = Nothing
   ElseIf lMenuChosen = 3 Then
      If m_FormulaID <= 0 Then
         glbErrorLog.LocalErrorMsg = "ไม่มีข้อมูลสูตรของใบสั่งผลิตนี้"
         glbErrorLog.ShowUserError
         Set oMenu = Nothing
         Exit Sub
      End If
      
      If ProcessID = 5 Then
         GridEX1.MoveFirst
         Set frmFormulaSelect.Job = m_Job
         frmFormulaSelect.ID = ID
         frmFormulaSelect.FORMULA_ID = m_FormulaID
         frmFormulaSelect.ParentShowMode = ShowMode
         If Val(GridEX1.Value(1)) > 0 Then
         frmFormulaSelect.ShowMode = SHOW_EDIT
         Else
         frmFormulaSelect.ShowMode = SHOW_ADD
         End If
         frmFormulaSelect.HeaderText = MapText("เพิ่มวัตถุดิบจากสูตร")
         Load frmFormulaSelect
         frmFormulaSelect.Show 1
         
         OKClick = frmFormulaSelect.OKClick
         If OKClick Then
         m_FormulaID = frmFormulaSelect.FORMULA_ID
         End If
         
         Unload frmFormulaSelect
         Set frmFormulaSelect = Nothing
      Else
         GridEX1.MoveFirst
         Set frmFormulaSelectWh.Job = m_Job
         frmFormulaSelectWh.ID = ID
         frmFormulaSelectWh.FORMULA_ID = m_FormulaID
         frmFormulaSelectWh.ParentShowMode = ShowMode
         If Val(GridEX1.Value(1)) > 0 Then
            frmFormulaSelectWh.ShowMode = SHOW_EDIT
         Else
            frmFormulaSelectWh.ShowMode = SHOW_ADD
         End If
         frmFormulaSelectWh.HeaderText = MapText("เพิ่มวัตถุดิบจากสูตร")
         Load frmFormulaSelectWh
         frmFormulaSelectWh.Show 1
   
         OKClick = frmFormulaSelectWh.OKClick
         If OKClick Then
            m_FormulaID = frmFormulaSelectWh.FORMULA_ID
         End If
         
         Unload frmFormulaSelectWh
         Set frmFormulaSelectWh = Nothing
      End If
   ElseIf lMenuChosen = 5 Then
        
         If m_Job.VERIFY_FLAG = "Y" Then
            glbErrorLog.LocalErrorMsg = "รายการนี้ผ่านการตรวจสอบจาก " & m_Job.VERIFY_NAME & " แล้ว ไม่สามารถอนุมัติซ้ำได้"
            glbErrorLog.ShowUserError
            Exit Sub
         End If
               
         frmVerifyAccRight.AccName = "INVENTORY-WH_IMPORT" & "_" & DOCUMENT_TYPE & "_VERIFY"
         frmVerifyAccRight.AccDesc = "ตรวจสอบสินค้าก่อนขาย"
         Load frmVerifyAccRight
         frmVerifyAccRight.Show 1

         If frmVerifyAccRight.GrantRight Then
            TempUserName = frmVerifyAccRight.UserName
            Unload frmVerifyAccRight
            Set frmVerifyAccRight = Nothing
            
               Call glbDaily.StartTransaction
               Set J = New CJob
               J.JOB_ID = ID
               Call J.UpdateJobVerifyFlag(TempUserName)
               Call glbDaily.CommitTransaction
               glbErrorLog.LocalErrorMsg = "ตรวจสอบสำเร็จ"
               glbErrorLog.ShowUserError
               
               OKClick = True
              Unload Me
         Else
            Unload frmVerifyAccRight
            Set frmVerifyAccRight = Nothing
            Exit Sub
         End If
   ElseIf lMenuChosen = 7 Then
        If m_Job.VERIFY_FLAG = "N" Then
            glbErrorLog.LocalErrorMsg = "เอกสารใบนี้ยังไม่ผ่านการตรวจสอบ ไม่สามารถยกเลิกได้"
            glbErrorLog.ShowUserError
        End If
         frmVerifyAccRight.AccName = "INVENTORY-WH_IMPORT" & "_" & DOCUMENT_TYPE & "_CANCEL-VERIFY"
         frmVerifyAccRight.AccDesc = "ยกเลิกการตรวจสอบสินค้าก่อนขาย"
         Load frmVerifyAccRight
         frmVerifyAccRight.Show 1

         If frmVerifyAccRight.GrantRight Then
            TempUserName = frmVerifyAccRight.UserName
            Unload frmVerifyAccRight
            Set frmVerifyAccRight = Nothing
            
               Call glbDaily.StartTransaction
               Set J = New CJob
               J.JOB_ID = ID
               Call J.UpdateJobCancelVerifyFlag(TempUserName)
               Call glbDaily.CommitTransaction
               glbErrorLog.LocalErrorMsg = "ยกลิกสำเร็จ"
               glbErrorLog.ShowUserError
               
               OKClick = True
               Unload Me
         Else
            Unload frmVerifyAccRight
            Set frmVerifyAccRight = Nothing
            Exit Sub
         End If
   End If
   
   If Not (lMenuChosen = 5 Or lMenuChosen = 7) Then
      If OKClick Then
         Call CalculateTotalRatio
         
         m_HasModify = True
         
         GridEX1.ItemCount = CountItem(m_Job.Inputs)
         GridEX1.Rebind
      End If
   End If
      
   Set oMenu = Nothing
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
   For Each D In m_Job.Inputs
      If D.Flag <> "D" Then
         Sum1 = Sum1 + D.TX_AMOUNT
      End If
   Next D
      
   For Each D In m_Job.Outputs
      If D.Flag <> "D" Then
         Sum2 = Sum2 + D.TX_AMOUNT
      End If
   Next D
   
   txtInputAmount.Text = FormatNumber(Sum1, 3)
   txtOutputAmount.Text = FormatNumber(Sum2, 3)
End Sub

Private Sub cmdDelete_Click()
Dim ID1 As Long
Dim ID2 As Long
Dim IWH As CInventoryWHDoc
Dim IWD As CInventoryWHDoc
Dim LIW As CLotItemWH
Dim Lt As CLotDoc
Dim PD As CPalletDoc

   If Not cmdDelete.Enabled Then
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
   If Not ConfirmDelete(GridEX1.Value(3)) Then
      Exit Sub
   End If
   
   ID1 = GridEX1.Value(1)
   ID2 = GridEX1.Value(2)
   
    If TabStrip1.SelectedItem.Index = 1 Then
      If ProcessID = 2 Then
          If m_CollLotExUse.Count > 0 Then '   ถ้า lotdocid นี้ มีการเบิกจ่ายไปแล้ว จะไม่ให้สามารถแก้ไขเลขที่ lot ได้แล้ว
            glbErrorLog.LocalErrorMsg = "ข้อมูลรับเข้ารายการนี้ นี้มีการเบิกจ่ายแล้ว ไม่สามารถลบได้"
            glbErrorLog.ShowUserError
            Exit Sub
         End If
       End If
       
      If ID1 <= 0 Then
         m_Job.Inputs.Remove (ID2)
      Else
         m_Job.Inputs.Item(ID2).Flag = "D"
         
        If ProcessID = 2 Or ProcessID = 6 Or ProcessID = 7 Or ProcessID = 8 Then
        typeForm = 2
         If ID1 <= 0 Then
            m_Job.InventoryWhDocInput.Remove (ID2)
         Else
           If Not m_Job.InventoryWhDocInput Is Nothing Then
            For Each IWD In m_Job.InventoryWhDocInput
            IWD.Flag = "D"
               For Each LIW In IWD.C_LotItemsWH
                  LIW.Flag = "D"
                  For Each Lt In LIW.C_LotDoc
                     Lt.Flag = "D"
                     For Each PD In Lt.C_PalletDoc
                        PD.Flag = "D"
                     Next PD
                  Next Lt
               Next LIW
            Next IWD
            End If
         End If
      End If
      End If

      Call CalculateTotalRatio
      GridEX1.ItemCount = CountItem(m_Job.Inputs)
      GridEX1.Rebind
      m_HasModify = True
ElseIf TabStrip1.SelectedItem.Index = 2 Then
       If ProcessID = 2 Or ProcessID = 6 Then
          If m_CollLotExUse.Count > 0 Then '   ถ้า lotdocid นี้ มีการเบิกจ่ายไปแล้ว จะไม่ให้สามารถแก้ไขเลขที่ lot ได้แล้ว
            glbErrorLog.LocalErrorMsg = "ข้อมูลรับเข้ารายการนี้ นี้มีการเบิกจ่ายแล้ว ไม่สามารถลบได้"
            glbErrorLog.ShowUserError
            Exit Sub
         End If
       End If
      If ID1 <= 0 Then 'If m_Job.Outputs.Item(ID2).Flag = "I" Then
         m_Job.Outputs.Remove (ID2)
      Else
         m_Job.Outputs.Item(ID2).Flag = "D" 'm_Job.Outputs.Remove (ID2)
      End If
      
      If ProcessID = 2 Or ProcessID = 6 Then
         If ID1 <= 0 Then
            m_Job.InventoryWhDoc.Remove (ID2)
         Else
         If Not m_Job.InventoryWhDoc Is Nothing Then
            For Each IWD In m_Job.InventoryWhDoc
            IWD.Flag = "D"
               For Each LIW In IWD.C_LotItemsWH
                  LIW.Flag = "D"
                  For Each Lt In LIW.C_LotDoc
                     Lt.Flag = "D"
                     For Each PD In Lt.C_PalletDoc
                        PD.Flag = "D"
                     Next PD
                  Next Lt
               Next LIW
            Next IWD
            End If
         End If
      End If
      
      Call CalculateTotalRatio
      GridEX1.ItemCount = CountItem(m_Job.Outputs)
      GridEX1.Rebind
      m_HasModify = True
     
ElseIf TabStrip1.SelectedItem.Index = 3 Then
      If ID1 <= 0 Then
         m_Job.Peoples.Remove (ID2)
      Else
         m_Job.Peoples.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_Job.Peoples)
      GridEX1.Rebind
      m_HasModify = True
   
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
      If ID1 <= 0 Then
         m_Job.Machines.Remove (ID2)
      Else
         m_Job.Machines.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_Job.Machines)
      GridEX1.Rebind
      m_HasModify = True
   
   ElseIf TabStrip1.SelectedItem.Index = 6 Then
      If ID1 <= 0 Then
         m_Job.Verifies.Remove (ID2)
      Else
         m_Job.Verifies.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_Job.Verifies)
      GridEX1.Rebind
      m_HasModify = True
   End If
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean

   If Not cmdEdit.Enabled Then
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
   ID = Val(GridEX1.Value(2))
   OKClick = False
   
   If TabStrip1.SelectedItem.Index = 1 Then
         If Not (m_Job.InventoryWhDocInput Is Nothing) Then 'ถ้าเป็นการแก้ไข InventoryWhDocInput
            typeForm = 2
         Else
            typeForm = 1
         End If
         
         Set frmAddEditJobInput.TempCollection = m_Job.Inputs
         
          
        If Not m_Job.InventoryWhDocInput Is Nothing Then
        Set m_InventoryWhDocInput = m_Job.InventoryWhDocInput.Item(1)
         m_InventoryWhDocInput.AddEditMode = SHOW_EDIT
          m_InventoryWhDocInput.Flag = "E"
        Else
          Set m_Job.InventoryWhDocInput = New Collection
          Set m_InventoryWhDocInput = New CInventoryWHDoc
          m_InventoryWhDocInput.AddEditMode = SHOW_ADD
         m_InventoryWhDocInput.Flag = "A"
          Call m_Job.InventoryWhDocInput.add(m_InventoryWhDocInput)
        End If
         
        Set frmAddEditJobInput.m_InventoryWhDocInput = m_Job.InventoryWhDocInput.Item(1)
         frmAddEditJobInput.ID = ID
         frmAddEditJobInput.ProcessID = ProcessID
         frmAddEditJobInput.DOCUMENT_DATE = uctlJobDate.ShowDate
         
         
         If DOCUMENT_TYPE = 14 Then
            frmAddEditJobInput.DOCUMENT_TYPE = 18
            frmAddEditJobInput.TYPE_LIST_RM = 5
         Else
            frmAddEditJobInput.DOCUMENT_TYPE = DOCUMENT_TYPE
         End If
         

         
         frmAddEditJobInput.ShowMode = SHOW_EDIT
         frmAddEditJobInput.COMMIT_FLAG = m_Job.OLD_COMMIT_FLAG
         frmAddEditJobInput.typeForm = typeForm
         Set frmAddEditJobInput.ParentForm = Me
         frmAddEditJobInput.HeaderText = MapText("แก้ไขวัตถุดิบ")
         Load frmAddEditJobInput
         frmAddEditJobInput.Show 1
         
         OKClick = frmAddEditJobInput.OKClick
         
         Unload frmAddEditJobInput
         Set frmAddEditJobInput = Nothing

      If OKClick Then
         Call CalculateTotalRatio
         
         GridEX1.ItemCount = CountItem(m_Job.Inputs)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   If Not VerifyCombo(lblJobProcess, cboJobProcess, False) Then
      Exit Sub
   End If
  ProcessID = cboJobProcess.ItemData(Minus2Zero(cboJobProcess.ListIndex))
  If ProcessID = 2 Or ProcessID = 6 Then  'ถ้าเป็น Process บรรจุ
  If m_Job.InventoryWhDoc Is Nothing Then
        typeForm = 1
         Set frmAddEditJobOutputEx.TempCollection = m_Job.Outputs
         frmAddEditJobOutputEx.COMMIT_FLAG = m_Job.OLD_COMMIT_FLAG
         frmAddEditJobOutputEx.ID = ID
         frmAddEditJobOutputEx.ShowMode = SHOW_EDIT
         frmAddEditJobOutputEx.HeaderText = MapText("แก้ไขผลผลิต")
         Load frmAddEditJobOutputEx
         frmAddEditJobOutputEx.Show 1
         OKClick = frmAddEditJobOutputEx.OKClick
         m_HasModify = False
         Unload frmAddEditJobOutputEx
         Set frmAddEditJobOutputEx = Nothing
  Else
         typeForm = 2
         Set frmAddEditJobOutputEx2.TempCollection = m_Job.Outputs
         Set frmAddEditJobOutputEx2.TempCollection2 = m_Job.InventoryWhDoc
         Set frmAddEditJobOutputEx2.m_CollLotExUse = m_CollLotExUse
         Set frmAddEditJobOutputEx2.m_CollPalletInLot = m_CollPalletInLot
         
         Set frmAddEditJobOutputEx2.TempPDEdit = TempPDEdit
         frmAddEditJobOutputEx2.COMMIT_FLAG = m_Job.OLD_COMMIT_FLAG
         frmAddEditJobOutputEx2.ID = ID
         frmAddEditJobOutputEx2.DOCUMENT_TYPE = DOCUMENT_TYPE

         frmAddEditJobOutputEx2.StartJob = uctlStartJob.ShowDate
         frmAddEditJobOutputEx2.StopJob = uctlFinishJob.ShowDate
         frmAddEditJobOutputEx2.ShowMode = SHOW_EDIT
         frmAddEditJobOutputEx2.HeaderText = MapText("แก้ไขผลิตภัณฑ์ที่ได้")
         Load frmAddEditJobOutputEx2
         frmAddEditJobOutputEx2.Show 1
         
         OKClick = frmAddEditJobOutputEx2.OKClick
         
         m_HasModify = False
         Unload frmAddEditJobOutputEx2
         Set frmAddEditJobOutputEx2 = Nothing
      End If
 ElseIf ProcessID = 4 Or ProcessID = 7 Then  'ถ้าเป็น Bulk บรรจุ
  If m_Job.InventoryWhDoc Is Nothing Then
        typeForm = 1
         Set frmAddEditJobOutputEx.TempCollection = m_Job.Outputs
         frmAddEditJobOutputEx.COMMIT_FLAG = m_Job.OLD_COMMIT_FLAG
         frmAddEditJobOutputEx.ID = ID
         frmAddEditJobOutputEx.ShowMode = SHOW_EDIT
         frmAddEditJobOutputEx.HeaderText = MapText("แก้ไขผลผลิต")
         Load frmAddEditJobOutputEx
         frmAddEditJobOutputEx.Show 1
         OKClick = frmAddEditJobOutputEx.OKClick
         m_HasModify = False
         Unload frmAddEditJobOutputEx
         Set frmAddEditJobOutputEx = Nothing
  Else
         typeForm = 2
         Set frmAddEditJobOutputEx4.TempCollection = m_Job.Outputs
         Set frmAddEditJobOutputEx4.TempCollection2 = m_Job.InventoryWhDoc
         frmAddEditJobOutputEx4.DocumentType = DOCUMENT_TYPE
         frmAddEditJobOutputEx4.COMMIT_FLAG = m_Job.OLD_COMMIT_FLAG
         frmAddEditJobOutputEx4.ID = ID
         frmAddEditJobOutputEx4.JobIdRef = JobIdRef
         'JobIdRef
         frmAddEditJobOutputEx4.StartJob = uctlStartJob.ShowDate
         frmAddEditJobOutputEx4.StopJob = uctlFinishJob.ShowDate
         frmAddEditJobOutputEx4.ShowMode = SHOW_EDIT
         frmAddEditJobOutputEx4.HeaderText = MapText("แก้ไขผลผลิต")
         Load frmAddEditJobOutputEx4
         frmAddEditJobOutputEx4.Show 1
         
         OKClick = frmAddEditJobOutputEx4.OKClick
         
         m_HasModify = False
         Unload frmAddEditJobOutputEx4
         Set frmAddEditJobOutputEx4 = Nothing
      End If
  Else
      If ProcessID = 8 Then
            frmAddEditJobOutputEx.typeInput = 1
      End If
      
     Set frmAddEditJobOutputEx.TempCollection = m_Job.Outputs
     frmAddEditJobOutputEx.COMMIT_FLAG = m_Job.OLD_COMMIT_FLAG
      frmAddEditJobOutputEx.ID = ID
      frmAddEditJobOutputEx.ShowMode = SHOW_EDIT
      frmAddEditJobOutputEx.HeaderText = MapText("แก้ไขผลผลิต")
      Load frmAddEditJobOutputEx
      frmAddEditJobOutputEx.Show 1

      OKClick = frmAddEditJobOutputEx.OKClick

      Unload frmAddEditJobOutputEx
      Set frmAddEditJobOutputEx = Nothing
  End If

      If OKClick Then
         Call CalculateTotalRatio
         
         GridEX1.ItemCount = CountItem(m_Job.Outputs)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
     Set frmAddEditJobPeople.TempCollection = m_Job.Peoples
      frmAddEditJobPeople.ID = ID
      frmAddEditJobPeople.ShowMode = SHOW_EDIT
      frmAddEditJobPeople.HeaderText = MapText("แก้ไขเครื่องจักรที่ใช้")
      Load frmAddEditJobPeople
      frmAddEditJobPeople.Show 1

      OKClick = frmAddEditJobPeople.OKClick

      Unload frmAddEditJobPeople
      Set frmAddEditJobPeople = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Job.Peoples)
         GridEX1.Rebind
      End If

ElseIf TabStrip1.SelectedItem.Index = 4 Then
     Set frmAddEditJobMachineEx.TempCollection = m_Job.Machines
      frmAddEditJobMachineEx.ID = ID
      frmAddEditJobMachineEx.ShowMode = SHOW_EDIT
      frmAddEditJobMachineEx.HeaderText = MapText("แก้ไขเครื่องจักรที่ใช้")
      Load frmAddEditJobMachineEx
      frmAddEditJobMachineEx.Show 1

      OKClick = frmAddEditJobMachineEx.OKClick

      Unload frmAddEditJobMachineEx
      Set frmAddEditJobMachineEx = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Job.Machines)
         GridEX1.Rebind
      End If
ElseIf TabStrip1.SelectedItem.Index = 5 Then
     Set frmAddEditJobParameter.TempCollection = m_Job.Parameters
      frmAddEditJobParameter.ID = ID
     frmAddEditJobParameter.Process = cboJobProcess.ItemData(Minus2Zero(cboJobProcess.ListIndex))
      frmAddEditJobParameter.ShowMode = SHOW_EDIT
      frmAddEditJobParameter.HeaderText = MapText("แก้ไขพารามิเตอร์ที่ใช้")
      Load frmAddEditJobParameter
      frmAddEditJobParameter.Show 1

      OKClick = frmAddEditJobParameter.OKClick

      Unload frmAddEditJobParameter
      Set frmAddEditJobParameter = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Job.Parameters)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 6 Then
     Set frmVerifyPartItemEx.TempCollection = m_Job.Verifies
      frmVerifyPartItemEx.ID = ID
      frmVerifyPartItemEx.ShowMode = SHOW_EDIT
      Set frmVerifyPartItemEx.ParentForm = Me
      frmVerifyPartItemEx.HeaderText = MapText("ตรวจสอบการใช้วัตถุดิบ")
      Load frmVerifyPartItemEx
      frmVerifyPartItemEx.Show 1

      OKClick = frmVerifyPartItemEx.OKClick

      Unload frmVerifyPartItemEx
      Set frmVerifyPartItemEx = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Job.Verifies)
         GridEX1.Rebind
      End If
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

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
   End If
 Set m_Job2 = Nothing
End Sub

Private Sub cmdOK_Click()
Dim oMenu As cPopupMenu
Dim lMenuChosen  As Long
If DOCUMENT_TYPE > 0 Then
   If Not VerifyAccessRight("INVENTORY-WH_IMPORT" & "_" & DOCUMENT_TYPE & "_SAVE", "บันทึก " & mainText) Then
          Call EnableForm(Me, True)
          Exit Sub
    End If
   End If
    
   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("บันทึก", "-", "บันทึกและออกจากหน้าจอ")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
   If lMenuChosen = 1 Then
      If ProcessID = 2 Or ProcessID = 4 Or ProcessID = 6 Or ProcessID = 7 Or ProcessID = 8 Then
        If typeForm = 1 Then
            If Not SaveData Then
               Exit Sub
            End If
        ElseIf typeForm = 2 Then
            If Not SaveData2 Then
               Exit Sub
            End If
            ShowMode = SHOW_EDIT
            ID = m_Job.JOB_ID
        Else
             If Not SaveData2 Then
               Exit Sub
            End If
         End If
      Else
         If Not SaveData Then
            Exit Sub
         End If
      End If

   m_Job.QueryFlag = 1
   If ProcessID = 2 Or ProcessID = 4 Or ProcessID = 6 Or ProcessID = 7 Or ProcessID = 8 Then
      Call QueryData2(True)
   Else
      Call QueryData(True)
   End If
      
      m_HasModify = False 'ถ้าบันทึกธรรมดาให้เปลี่ยน m_HasModify = False ด้วย
      OKClick = True
   ElseIf lMenuChosen = 3 Then
     If ProcessID = 2 Or ProcessID = 4 Or ProcessID = 6 Or ProcessID = 7 Or ProcessID = 8 Then
        If typeForm = 1 Then
            If Not SaveData Then
               Exit Sub
            End If
        ElseIf typeForm = 2 Then
            If Not SaveData2 Then
               Exit Sub
            End If
        Else
             If Not SaveData2 Then
               Exit Sub
            End If
         End If
      Else
         If Not SaveData Then
            Exit Sub
         End If
      End If
      OKClick = True
      Unload Me
   End If
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
Dim ReportType As Long

   If m_HasModify Or ((Not m_HasModify) And (ShowMode = SHOW_ADD)) Then
      glbErrorLog.LocalErrorMsg = "กรุณาทำการบันทึกข้อมูลให้เรียบร้อยก่อน"
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   Set oMenu = New cPopupMenu
   If JobDocType = 1 Then
      lMenuChosen = oMenu.Popup("ใบสั่งผลิต", "ปรับค่าหน้ากระดาษ", "ใบสั่งผลิต+ข้อมูลมาตรฐานแบบ 1", "ใบสั่งผลิต+ข้อมูลมาตรฐานแบบ 2")
      If lMenuChosen = 0 Then
         Exit Sub
      End If
   ElseIf JobDocType = 2 Then
      lMenuChosen = oMenu.Popup("ใบประเมินราคาผลิต", "ปรับค่าหน้ากระดาษ")
      If lMenuChosen = 0 Then
         Exit Sub
      End If
   End If
   
   Call EnableForm(Me, False)
   
   If JobDocType = 1 Then
      If lMenuChosen = 1 Then
         ReportKey = "CReportJob001"
         
         Set Report = New CReportJob001
         ReportFlag = True
      ElseIf lMenuChosen = 2 Then
         ReportKey = "CReportJob001"
         
         Set Rc = New CReportConfig
         Rc.REPORT_KEY = ReportKey
         Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
         Call Rc.QueryData(m_Rs, iCount)
         HeaderText = MapText("ใบสั่งผลิต")
         If Not m_Rs.EOF Then
            Call Rc.PopulateFromRS(1, m_Rs)
            EditMode = SHOW_EDIT
         Else
            EditMode = SHOW_ADD
         End If
      ElseIf lMenuChosen = 3 Then
         ReportKey = "CReportJob002"
         Set Report = New CReportJob002
         ReportFlag = True
         ReportType = 1
      ElseIf lMenuChosen = 4 Then
         ReportKey = "CReportJob002"
         Set Report = New CReportJob002
         ReportFlag = True
         ReportType = 2
      End If
   ElseIf JobDocType = 2 Then
      If lMenuChosen = 1 Then
         ReportKey = "CReportEstimate001"
         
         Set Report = New CReportEstimate001
         ReportFlag = True
      ElseIf lMenuChosen = 2 Then
         ReportKey = "CReportEstimate001"
         
         Set Rc = New CReportConfig
         Rc.REPORT_KEY = ReportKey
         Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
         Call Rc.QueryData(m_Rs, iCount)
         HeaderText = MapText("ใบประเมินราคาผลิต")
         If Not m_Rs.EOF Then
            Call Rc.PopulateFromRS(1, m_Rs)
            EditMode = SHOW_EDIT
         Else
            EditMode = SHOW_ADD
         End If
      End If
   End If
   
   If Not Report Is Nothing Then
      Call Report.AddParam(m_Job.JOB_ID, "JOB_ID")
      Call Report.AddParam(ReportKey, "REPORT_KEY")
      Call Report.AddParam(ReportType, "REPORT_TYPE")
   End If
   
   If ReportFlag Then
      Set frmReport.ReportObject = Report
      frmReport.HeaderText = pnlHeader.Caption
      Load frmReport
      frmReport.Show 1
   
      Unload frmReport
      Set frmReport = Nothing
      Set Report = Nothing
   Else
      frmReportConfig.ReportMode = 1
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

Private Sub cmdPrint_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
   KeyAscii = 0
End Sub

Private Sub cmdSave_Click()
If DOCUMENT_TYPE > 0 Then
   If Not VerifyAccessRight("INVENTORY-WH_IMPORT" & "_" & DOCUMENT_TYPE & "_SAVE", "บันทึก " & mainText) Then
          Call EnableForm(Me, True)
          Exit Sub
   End If
End If
   
 If ProcessID = 2 Or ProcessID = 4 Or ProcessID = 6 Or ProcessID = 7 Then
        If typeForm = 1 Then
            If Not SaveData Then
               Exit Sub
            End If
        ElseIf typeForm = 2 Then
            If Not SaveData2 Then
               Exit Sub
            End If
            ShowMode = SHOW_EDIT
            ID = m_Job.JOB_ID
        Else
             If Not SaveData2 Then
               Exit Sub
            End If
         End If
      Else
         If Not SaveData Then
            Exit Sub
         End If
      End If

   m_Job.QueryFlag = 1
   If ProcessID = 2 Or ProcessID = 4 Or ProcessID = 6 Or ProcessID = 7 Then
      Call QueryData2(True)
   Else
      Call QueryData(True)
   End If
      
      m_HasModify = False 'ถ้าบันทึกธรรมดาให้เปลี่ยน m_HasModify = False ด้วย
      OKClick = True
End Sub

Private Sub cmdSave_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
   KeyAscii = 0
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
   End If
   Set m_Job2 = Nothing
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      Call LoadEmployee(ucltApproveByLookup.MyCombo, m_Employees)
      Set ucltApproveByLookup.MyCollection = m_Employees
      
      Call LoadEmployee(uctlResponseByLookup.MyCombo, m_Employees)
      Set uctlResponseByLookup.MyCollection = m_Employees
      
'      If DOCUMENT_TYPE = 13 Or DOCUMENT_TYPE = 14 Or DOCUMENT_TYPE = 17 Or DOCUMENT_TYPE = 18 Or DOCUMENT_TYPE = 19 Then
'      Call LoadEmployee(uctlEmpCheckCar.MyCombo, m_Employees, 23) '23 คือเลือกเฉพาะเจ้าหน้าที่ห้องแพ็ค
'      Set uctlEmpCheckCar.MyCollection = m_Employees
'      End If
      
      Call LoadEmployee(uctlResponseByLookup.MyCombo)
      Call LoadProcess(cboJobProcess)
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_Job.QueryFlag = 1
         If ProcessID = 2 Or ProcessID = 4 Or ProcessID = 6 Or ProcessID = 7 Or ProcessID = 8 Then
            Call QueryData2(True)
            
            Dim IWD As CInventoryWHDoc
            Dim LIW As CLotItemWH
            Dim LTD As CLotDoc
            Dim LotDocId As Long
            Dim LotId As Long

            If Not m_Job.InventoryWhDoc Is Nothing Then
               LotDocId = -1
               For Each IWD In m_Job.InventoryWhDoc
                  For Each LIW In IWD.C_LotItemsWH
                     For Each LTD In LIW.C_LotDoc
                        LotDocId = LTD.LOT_DOC_ID
                        LotId = LTD.LOT_ID
                     Next LTD
                  Next LIW
               Next IWD
               Call LoadLotRefExByLotDocId(Nothing, m_CollLotExUse, -1, -1, LotDocId)
               Call LoadPalletByLotID(Nothing, m_CollPalletInLot, LotId, "I") 'หา pallet no ที่อยู่ ใน lot นั้น ทั้งหมด
            End If
            
         Else
            Call QueryData(True)
         End If
      ElseIf ShowMode = SHOW_ADD Then
         uctlJobDate.ShowDate = Now
         uctlStartJob.ShowDate = Now
         uctlFinishJob.ShowDate = Now
         cboJobProcess.ListIndex = IDToListIndex(cboJobProcess, ProcessID)
         
        m_Job.QueryFlag = 0
         Call QueryData2(False)
      End If
      
      Call TabStrip1_Click
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

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_Job = Nothing
   Set m_Jobs = Nothing
   Set m_Employees = Nothing
   Set m_CollLotExUse = Nothing
   Set m_CollPalletInLot = Nothing
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
   
'   GridEX1.Font.Bold = False
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"
   'GridEX1.Columns.Item(1).Visible = False

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
   'GridEX1.Columns.Item(2).Visible = False
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2400
   Col.Caption = MapText("หมายเลขวัตถุดิบ")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 3500
   Col.Caption = MapText("ชื่อวัตถุดิบ")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 2000
   Col.Caption = MapText("ประเภทวัตถุดิบ")

   Set Col = GridEX1.Columns.add '6
   Col.Width = 1700
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("จำนวน")

   Set Col = GridEX1.Columns.add '7
   Col.Width = 2500
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ราคา")

   Set Col = GridEX1.Columns.add '8
   Col.Width = 2500
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ราคารวม")

   Set Col = GridEX1.Columns.add '9
   Col.Width = 2500
   Col.Caption = MapText("รหัสสถานที่จัดเก็บ")
   
   Set Col = GridEX1.Columns.add '8
   Col.Width = 2500
   Col.Caption = MapText("สถานที่จัดเก็บ")
   
   Set Col = GridEX1.Columns.add '9
   Col.Width = 2500
   Col.Visible = False
   Col.Caption = MapText("รหัสเชื่อมโยงวัตถุดิบ")
   
   Set Col = GridEX1.Columns.add '10
   Col.Width = 2500
   Col.Caption = MapText("ซีเรียล")
   
   Set Col = GridEX1.Columns.add '11
   Col.Width = 2500
   Col.Caption = MapText("หมายเลขอ้างอิง")
   
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
   
'   GridEX1.Font.Bold = False
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"
   GridEX1.Columns.Item(1).Visible = True

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
   GridEX1.Columns.Item(2).Visible = False
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2400
   Col.Caption = MapText("หมายเลขวัตถุดิบ")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 3500
   Col.Caption = MapText("ชื่อวัตถุดิบ")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 2000
   Col.Caption = MapText("ประเภทวัตถุดิบ")

   Set Col = GridEX1.Columns.add '6
   Col.Width = 1700
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("จำนวน")

   Set Col = GridEX1.Columns.add '6
   Col.Width = 2500
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ราคา")

   Set Col = GridEX1.Columns.add '6
   Col.Width = 2500
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ราคารวม")

   Set Col = GridEX1.Columns.add '7
   Col.Width = 2500
   Col.Caption = MapText("รหัสสถานที่จัดเก็บ")
   
   Set Col = GridEX1.Columns.add '8
   Col.Width = 2500
   Col.Caption = MapText("สถานที่จัดเก็บ")
   
   Set Col = GridEX1.Columns.add '9
   Col.Width = 2500
   Col.Visible = False
   Col.Caption = MapText("รหัสเชื่อมโยงวัตถุดิบ")
   
   Set Col = GridEX1.Columns.add '10
   Col.Width = 2500
   Col.Caption = MapText("ซีเรียล")
   
   Set Col = GridEX1.Columns.add '11
   Col.Width = 2500
   Col.Caption = MapText("หมายเลขอ้างอิง")
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
   
'   GridEX1.Font.Bold = False
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"
   Col.Visible = False

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
   Col.Visible = False

   Set Col = GridEX1.Columns.add '3
   Col.Width = 3500
   Col.Caption = MapText("ชื่อพนักงาน")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 2610
   Col.Caption = MapText("ตำแหน่ง")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 3000
   Col.Caption = MapText("เวลาที่ใช้")
   
   Set Col = GridEX1.Columns.add '6
   Col.Width = 2500
   Col.Caption = MapText("วันที่ทำงาน")

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
   
'   GridEX1.Font.Bold = False
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"
GridEX1.Columns.Item(1).Visible = False
   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
GridEX1.Columns.Item(2).Visible = False
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2610
   Col.Caption = MapText("หมายเลขเครื่องจักร")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 3500
   Col.Caption = MapText("ชื่อเครื่องจักร")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 3000
   Col.Caption = MapText("จำนวนเวลา(ชั่วโมง/นาที)")

 Set Col = GridEX1.Columns.add '6
   Col.Width = 2500
   Col.Caption = MapText("วันที่ใช้เครื่องจักร")
End Sub

Private Sub InitGrid5()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.NAME = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
'   GridEX1.Font.Bold = False
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"
   Col.Visible = False

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
   Col.Visible = False

   Set Col = GridEX1.Columns.add '3
   Col.Width = 3610
   Col.Caption = MapText("ค่าใช้จ่ายผลิต")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 3500 + 1500
   Col.Caption = MapText("รายละเอียด")

   Set Col = GridEX1.Columns.add '5
   Col.TextAlignment = jgexAlignRight
   Col.Width = 3000
   Col.Caption = MapText("มูลค่า")
End Sub

Private Sub InitGrid6()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

   GridEX1.FormatStyles.Clear
   Set fmsTemp = GridEX1.FormatStyles.add("N")
   fmsTemp.ForeColor = GLB_ALERT_COLOR
   fmsTemp.FontBold = True
   Set fmsTemp = GridEX1.FormatStyles.add("Y")
   fmsTemp.ForeColor = RGB(0, 255, 0)
   fmsTemp.FontBold = True
   
'   GridEX1.Font.Bold = True
   
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
   Col.Visible = False

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
   Col.Visible = False

   Set Col = GridEX1.Columns.add '3
   Col.Width = 2025
   Col.Caption = MapText("รหัสวัตถุดิบ")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 5850
   Col.Caption = MapText("ชื่อวัตถุดิบ")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 3720
   Col.Caption = MapText("หมายเหตุ")

   Set Col = GridEX1.Columns.add '6
   Col.Width = 0
   Col.Visible = False
   Col.Caption = MapText("FLAG")
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = "เพิ่มข้อมูล" & HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitGrid1
   
   If JobDocType = 1 Then
      Call InitNormalLabel(lblJobNo, MapText("เลขที่ใบสั่งผลิต"))
      If ProcessID = 2 Or ProcessID = 5 Or ProcessID = 6 Or ProcessID = 7 Or ProcessID = 8 Then
         Call InitNormalLabel(lblJobDate, MapText("วันที่เอกสาร"))
      Else
         Call InitNormalLabel(lblJobDate, MapText("วันที่สั่งผลิต"))
      End If
   Else
      Call InitNormalLabel(lblJobNo, MapText("เลขที่ใบประเมิน"))
      Call InitNormalLabel(lblJobDate, MapText("วันที่ประเมิน"))
   End If
   Call InitNormalLabel(lblJobDesc, MapText("รายละเอียด"))
   Call InitNormalLabel(lblNote, MapText("หมายเหตุ"))
   Call InitNormalLabel(lblBatchNo, MapText("จำนวนแบต"))
   Call InitNormalLabel(lblJobApp, MapText("ผู้อนุมัติ"))
   Call InitNormalLabel(lblJobRes, MapText("ผู้รับผิดชอบ"))
   Call InitNormalLabel(lblStartJob, MapText("วันที่เริ่มผลิต"))
   Call InitNormalLabel(lblFinishJob, MapText("วันที่ผลิตเสร็จ"))
   Call InitNormalLabel(lblJobProcess, MapText("โปรเซส"))
   Call InitNormalLabel(Label3, MapText("ก.ก."))
   Call InitNormalLabel(Label4, MapText("ก.ก."))
   Call InitNormalLabel(lblInputAmount, MapText("ยอดใช้รวม"))
   Call InitNormalLabel(lblOutputAmount, MapText("ผลิตรวม"))
   Call InitCheckBox(chkCommit, "งานเสร็จแล้ว")
   If ProcessID = 4 Then
      typeForm = 2 'ให้แสดง Form แบบใหม่ ที่มี InventoryWh เข้ามาเกี่ยวข้องแล้ว
      
      lblFromBatch.Visible = True
      txtFromBatch.Visible = True
      lblToBatch.Visible = True
      txtToBatch.Visible = True
      lblTotalBatch.Visible = True
      txtTotalBatch.Visible = True
      
      Call InitNormalLabel(lblFromBatch, MapText("จากแบต"))
      Call InitNormalLabel(lblToBatch, MapText("ถึงแบต"))
      Call InitNormalLabel(lblTotalBatch, MapText("แบตทั้งหมด"))
      
      Call txtFromBatch.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
      Call txtToBatch.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
      Call txtTotalBatch.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   End If
   
   Call txtJobNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtJobDesc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtNote.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtBatchNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtInputAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtInputAmount.Enabled = False
   Call txtOutputAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtOutputAmount.Enabled = False
   

   
   Call InitCombo(uctlResponseByLookup.MyCombo)
   Call InitCombo(cboJobProcess)
'   Call InitCombo(cboLotNo)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSave.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPrint.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAuto.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdCalculate.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdLock.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdUnlock.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSave, MapText("บันทึก"))
   Call InitMainButton(cmdPrint, MapText("พิมพ์"))
   Call InitMainButton(cmdAuto, MapText("A"))
   Call InitMainButton(cmdCalculate, MapText("อื่น ๆ"))
   Call InitMainButton(cmdLock, MapText("ล็อค"))
   Call InitMainButton(cmdUnlock, MapText("ปลดล็อค"))
   
   If ProcessID = 4 Then
      cmdLock.Visible = True
      cmdUnlock.Visible = True
   Else
      cmdLock.Visible = False
      cmdUnlock.Visible = False
   End If
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.NAME = GLB_FONT
   TabStrip1.Font.Size = 16
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.add().Caption = MapText("วัตถุดิบที่ใช้")
   TabStrip1.Tabs.add().Caption = MapText("ผลิตภัณฑ์ที่ได้")
   TabStrip1.Tabs.add().Caption = MapText("แรงงาน")
   TabStrip1.Tabs.add().Caption = MapText("เครื่องจักร")
   TabStrip1.Tabs.add().Caption = MapText("ค่าใช้จ่ายผลิต")
   If JobDocType = 1 Then
      TabStrip1.Tabs.add().Caption = MapText("ตรวจสอบ")
   End If
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
  ucltApproveByLookup.MyCombo.ListIndex = -1
   m_HasActivate = False
   m_HasModify = False
   
   m_FormulaID = -1
   Set m_Rs = New ADODB.Recordset
   Set m_Job = New CJob
   Set m_Jobs = New Collection
   Set m_Employees = New Collection
    Set m_CollLotExUse = New Collection
    Set m_CollPalletInLot = New Collection
   
   Set TempPDEdit = New Collection
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
   KeyAscii = 0
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   If TabStrip1.SelectedItem.Index = 6 Then
      RowBuffer.RowStyle = RowBuffer.Value(6)
   End If
End Sub
Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Index = 1 Then
     If m_Job.Inputs Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim Ci As CJobInput
      If m_Job.Inputs.Count <= 0 Then
         Exit Sub
      End If
      Set Ci = GetItem(m_Job.Inputs, RowIndex, RealIndex)
      If Ci Is Nothing Then
         Exit Sub
      End If
      
      Values(1) = Ci.JOB_INOUT_ID
      Values(2) = RealIndex
      Values(3) = Ci.PART_NO
      Values(4) = Ci.PART_DESC
      Values(5) = Ci.PART_TYPE_NAME
      Values(6) = FormatNumber(Ci.TX_AMOUNT, 3)
      Values(7) = FormatNumber(Ci.INCLUDE_UNIT_PRICE, 3)
      Values(8) = FormatNumber(Ci.TX_AMOUNT * Ci.INCLUDE_UNIT_PRICE, 3)
      Values(9) = Ci.LOCATION_NO
      Values(10) = Ci.LOCATION_NAME
      Values(11) = Ci.LINK_ID
      Values(12) = Ci.SERIAL_NUMBER
     Values(13) = Ci.INOUT_REF
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
     If m_Job.Outputs Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim CO As CJobInput
      If m_Job.Outputs.Count <= 0 Then
         Exit Sub
      End If
      Set CO = GetItem(m_Job.Outputs, RowIndex, RealIndex)
      If CO Is Nothing Then
         Exit Sub
      End If
      Values(1) = CO.JOB_INOUT_ID
      Values(2) = RealIndex
      Values(3) = CO.PART_NO
      Values(4) = CO.PART_DESC
      Values(5) = CO.PART_TYPE_NAME
      Values(6) = FormatNumber(CO.TX_AMOUNT, 3)
      Values(7) = FormatNumber(CO.INCLUDE_UNIT_PRICE, 3)
      Values(8) = FormatNumber(CO.INCLUDE_UNIT_PRICE * CO.TX_AMOUNT, 3)
      Values(9) = CO.LOCATION_NO
      Values(10) = CO.LOCATION_NAME
      Values(11) = CO.LINK_ID
      Values(12) = CO.SERIAL_NUMBER
     Values(13) = CO.INOUT_REF
    ElseIf TabStrip1.SelectedItem.Index = 3 Then
     If m_Job.Peoples Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim CR As CJobResource
      If m_Job.Peoples.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_Job.Peoples, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If
      Values(1) = CR.JOB_RESOURCE_ID
      Values(2) = RealIndex
      Values(3) = CR.LONG_NAME & "  " & CR.LAST_NAME
      Values(4) = CR.POSITION_NAME
      Values(5) = FormatNumber(CR.OCCUPY_INTERVAL)
      Values(6) = DateToStringExt(CR.OCCUPY_DATE)
      
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
      If m_Job.Machines Is Nothing Then
            Exit Sub
         End If
   
         If RowIndex <= 0 Then
            Exit Sub
         End If
   
         Dim CP As CJobResource
        Set CP = New CJobResource
         If m_Job.Machines.Count <= 0 Then
            Exit Sub
         End If
         Set CP = GetItem(m_Job.Machines, RowIndex, RealIndex)
         If CP Is Nothing Then
            Exit Sub
         End If
         Values(1) = CP.JOB_RESOURCE_ID
         Values(2) = RealIndex
         Values(3) = CP.MACHINE_NO
         Values(4) = CP.MACHINE_NAME
         Values(5) = FormatNumber(CP.OCCUPY_INTERVAL)
         Values(6) = DateToStringExt(CP.OCCUPY_DATE)
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
      If m_Job.Parameters Is Nothing Then
            Exit Sub
         End If
   
         If RowIndex <= 0 Then
            Exit Sub
         End If
   
         Dim Ca As CJobParameter
        Set Ca = New CJobParameter
         If m_Job.Parameters.Count <= 0 Then
            Exit Sub
         End If
         Set Ca = GetItem(m_Job.Parameters, RowIndex, RealIndex)
         If Ca Is Nothing Then
            Exit Sub
         End If
         Values(1) = Ca.JOB_PARAMETER_ID
         Values(2) = RealIndex
         Values(3) = Ca.PARAMETER_PROCESS_NAME
         Values(4) = Ca.JOB_PARAMETER_DESC
         Values(5) = FormatNumber(Ca.PARAM_AMOUNT)
   ElseIf TabStrip1.SelectedItem.Index = 6 Then
      If m_Job.Parameters Is Nothing Then
            Exit Sub
         End If
   
         If RowIndex <= 0 Then
            Exit Sub
         End If
   
         Dim Jv As CJobVerify
        Set Jv = New CJobVerify
         If m_Job.Verifies.Count <= 0 Then
            Exit Sub
         End If
         Set Jv = GetItem(m_Job.Verifies, RowIndex, RealIndex)
         If Jv Is Nothing Then
            Exit Sub
         End If
         Values(1) = Jv.JOB_VERIFY_ID
         Values(2) = RealIndex
         Values(3) = Jv.PART_NO
         Values(4) = Jv.PART_DESC
         Values(5) = Jv.NOTE
         Values(6) = Jv.VERIFY_FLAG
      End If
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub AllVisibleFalse()
cmdAdd.Visible = False
cmdEdit.Visible = False
cmdDelete.Visible = False
End Sub

Private Sub PopulateGuiID(BD As CJob)
Dim Di As CJobInput

   For Each Di In BD.Inputs
      If Di.Flag = "A" Then
         Di.LINK_ID = GetNextGuiID(BD)
      End If
   Next Di

   For Each Di In BD.Outputs
      If Di.Flag = "A" Then
         Di.LINK_ID = GetNextGuiID(BD)
      End If
   Next Di
End Sub

Private Function GetNextGuiID(BD As CJob) As Long
Dim Di As CJobInput
Dim MaxId As Long

   MaxId = 0
   For Each Di In BD.Inputs
      If Di.LINK_ID > MaxId Then
         MaxId = Di.LINK_ID
      End If
   Next Di
   
   For Each Di In BD.Outputs
      If Di.LINK_ID > MaxId Then
         MaxId = Di.LINK_ID
      End If
   Next Di
   
   GetNextGuiID = MaxId + 1
End Function

Private Sub Label1_Click()

End Sub

Private Sub TabStrip1_Click()

   cmdAdd.Enabled = True
   cmdEdit.Enabled = True
   cmdDelete.Enabled = True
   If TabStrip1.SelectedItem.Index = 1 Then
      Call InitGrid1
      Call CalculateTotalRatio
     GridEX1.ItemCount = CountItem(m_Job.Inputs)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
     Call InitGrid2
     GridEX1.ItemCount = CountItem(m_Job.Outputs)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
       Call InitGrid3
    GridEX1.ItemCount = CountItem(m_Job.Peoples)
     GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
     Call InitGrid4
   GridEX1.ItemCount = CountItem(m_Job.Machines)
   GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
      cmdAdd.Enabled = False
      cmdEdit.Enabled = False
      cmdDelete.Enabled = False
      Call InitGrid5
      GridEX1.ItemCount = CountItem(m_Job.Parameters)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 6 Then
      cmdEdit.Enabled = False
      Call InitGrid6
      GridEX1.ItemCount = CountItem(m_Job.Verifies)
      GridEX1.Rebind
   End If
End Sub


Private Sub TabStrip1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
   KeyAscii = 0
End Sub

Private Sub txtBatchNo_Change()
   m_HasModify = True
End Sub

Private Sub txtFromBatch_Change()
   m_HasModify = True
End Sub

Private Sub txtFromBatch_KeyPress(KeyAscii As Integer)
   KeyAscii = CheckIntAscii(KeyAscii)
End Sub

Private Sub txtJobDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtJobNo_Change()
   m_HasModify = True
End Sub

Private Sub txtNote_Change()
   m_HasModify = True
End Sub

Private Sub txtToBatch_Change()
   m_HasModify = True
End Sub

Private Sub txtToBatch_KeyPress(KeyAscii As Integer)
    KeyAscii = CheckIntAscii(KeyAscii)
End Sub

Private Sub txtTotalBatch_Change()
   m_HasModify = True
End Sub

Private Sub txtTotalBatch_KeyPress(KeyAscii As Integer)
    KeyAscii = CheckIntAscii(KeyAscii)
End Sub

Private Sub ucltApproveByLookup_Change()
'If DOCUMENT_TYPE = 14 Then
'   If Not VerifyAccessRight("PRODUCT_JOB_APPROVE", "อนุมัติการรับเข้าอาหาร") Then
'     Call EnableForm(Me, True)
'     ucltApproveByLookup.MyCombo.ListIndex = -1
'     Exit Sub
'   End If
'
'   If Not VerifyAccessRight("PRODUCT_JOB_APPROVE_DOC-BAG", "อนุมัติการตรวจสอบการรับเข้าอาหาร BAG ได้") Then
'     Call EnableForm(Me, True)
'     ucltApproveByLookup.MyCombo.ListIndex = -1
'     Exit Sub
'   End If
'ElseIf DOCUMENT_TYPE = 13 Then
'   If Not VerifyAccessRight("PRODUCT_JOB_APPROVE", "อนุมัติการรับเข้าอาหาร") Then
'     Call EnableForm(Me, True)
'     ucltApproveByLookup.MyCombo.ListIndex = -1
'     Exit Sub
'   End If
'
'   If Not VerifyAccessRight("PRODUCT_JOB_APPROVE_DOC-BULK", "อนุมัติการตรวจสอบการรับเข้าอาหาร BULK ได้") Then
'     Call EnableForm(Me, True)
'     ucltApproveByLookup.MyCombo.ListIndex = -1
'     Exit Sub
'   End If
'ElseIf DOCUMENT_TYPE = 17 Then
'   If Not VerifyAccessRight("PRODUCT_JOB_APPROVE", "อนุมัติการรับเข้าอาหาร") Then
'     Call EnableForm(Me, True)
'     ucltApproveByLookup.MyCombo.ListIndex = -1
'     Exit Sub
'   End If
'
'   If Not VerifyAccessRight("PRODUCT_JOB_APPROVE_DOC-RE-BAG-TO-BAG", "อนุมัติการตรวจสอบการรับเข้าอาหาร RE BAG TO BAG ได้") Then
'     Call EnableForm(Me, True)
'     ucltApproveByLookup.MyCombo.ListIndex = -1
'     Exit Sub
'   End If
'ElseIf DOCUMENT_TYPE = 18 Then
'   If Not VerifyAccessRight("PRODUCT_JOB_APPROVE", "อนุมัติการรับเข้าอาหาร") Then
'     Call EnableForm(Me, True)
'     ucltApproveByLookup.MyCombo.ListIndex = -1
'     Exit Sub
'   End If
'
'   If Not VerifyAccessRight("PRODUCT_JOB_APPROVE_DOC-RE-BAG-TO-BULK", "อนุมัติการตรวจสอบการรับเข้าอาหาร RE BAG TO BULK ได้") Then
'     Call EnableForm(Me, True)
'     ucltApproveByLookup.MyCombo.ListIndex = -1
'     Exit Sub
'   End If
'ElseIf DOCUMENT_TYPE = 19 Then
'   If Not VerifyAccessRight("PRODUCT_JOB_APPROVE", "อนุมัติการรับเข้าอาหาร") Then
'     Call EnableForm(Me, True)
'     ucltApproveByLookup.MyCombo.ListIndex = -1
'     Exit Sub
'   End If
'
'   If Not VerifyAccessRight("PRODUCT_JOB_APPROVE_DOC-RE-BAG-TO-RM", "อนุมัติการตรวจสอบการรับเข้าอาหาร RE BAG TO RM(OTHER) ได้") Then
'     Call EnableForm(Me, True)
'     ucltApproveByLookup.MyCombo.ListIndex = -1
'     Exit Sub
'   End If
'End If
   m_HasModify = True
End Sub

Private Sub uctlFinishJob_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlJobDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlResponseByLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlStartJob_HasChange()
   m_HasModify = True
End Sub
