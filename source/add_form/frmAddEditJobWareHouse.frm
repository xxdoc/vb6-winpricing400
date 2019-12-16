VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditJobWareHouse 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditJobWareHouse.frx":0000
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
      Height          =   8895
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   12720
      _ExtentX        =   22437
      _ExtentY        =   15690
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextLookup ucltApproveByLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   10
         Top             =   3180
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin VB.ComboBox cboJobProcess 
         Height          =   315
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1890
         Width           =   3100
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   120
         TabIndex        =   15
         Top             =   4380
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
         TabIndex        =   24
         Top             =   0
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   2790
         Left            =   120
         TabIndex        =   16
         Top             =   4920
         Width           =   11640
         _ExtentX        =   20532
         _ExtentY        =   4921
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
         Column(1)       =   "frmAddEditJobWareHouse.frx":27A2
         Column(2)       =   "frmAddEditJobWareHouse.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditJobWareHouse.frx":290E
         FormatStyle(2)  =   "frmAddEditJobWareHouse.frx":2A6A
         FormatStyle(3)  =   "frmAddEditJobWareHouse.frx":2B1A
         FormatStyle(4)  =   "frmAddEditJobWareHouse.frx":2BCE
         FormatStyle(5)  =   "frmAddEditJobWareHouse.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditJobWareHouse.frx":2D5E
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
         TabIndex        =   8
         Top             =   2730
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlJobDate 
         Height          =   405
         Left            =   6840
         TabIndex        =   2
         Top             =   990
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtBatchNo 
         Height          =   435
         Left            =   1800
         TabIndex        =   3
         Top             =   1440
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlFinishJob 
         Height          =   405
         Left            =   1800
         TabIndex        =   7
         Top             =   2310
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlDate uctlStartJob 
         Height          =   405
         Left            =   1800
         TabIndex        =   5
         Top             =   1890
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextLookup uctlResponseByLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   12
         Top             =   3600
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtInputAmount 
         Height          =   435
         Left            =   8490
         TabIndex        =   9
         Top             =   2700
         Width           =   1695
         _ExtentX        =   4683
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtOutputAmount 
         Height          =   435
         Left            =   8490
         TabIndex        =   11
         Top             =   3150
         Width           =   1695
         _ExtentX        =   4683
         _ExtentY        =   767
      End
      Begin VB.Label Label3 
         Caption         =   "lblFormulaNo"
         Height          =   315
         Left            =   10230
         TabIndex        =   37
         Top             =   2850
         Width           =   1305
      End
      Begin VB.Label Label4 
         Caption         =   "lblFormulaNo"
         Height          =   315
         Left            =   10230
         TabIndex        =   36
         Top             =   3300
         Width           =   1305
      End
      Begin VB.Label lblOutputAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblBatchNo"
         Height          =   315
         Left            =   7260
         TabIndex        =   35
         Top             =   3300
         Width           =   1125
      End
      Begin VB.Label lblInputAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblBatchNo"
         Height          =   315
         Left            =   7260
         TabIndex        =   34
         Top             =   2850
         Width           =   1125
      End
      Begin Threed.SSCommand cmdCalculate 
         Height          =   525
         Left            =   6840
         TabIndex        =   20
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJobWareHouse.frx":2F36
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
         MouseIcon       =   "frmAddEditJobWareHouse.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   525
         Left            =   8490
         TabIndex        =   13
         Top             =   3630
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJobWareHouse.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   525
         Left            =   10110
         TabIndex        =   14
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
         TabIndex        =   33
         Top             =   1950
         Width           =   945
      End
      Begin VB.Label lblJobRes 
         Alignment       =   1  'Right Justify
         Caption         =   "lblJobRes"
         Height          =   315
         Left            =   120
         TabIndex        =   32
         Top             =   3720
         Width           =   1605
      End
      Begin VB.Label lblJobApp 
         Alignment       =   1  'Right Justify
         Caption         =   "lblJobApp"
         Height          =   315
         Left            =   420
         TabIndex        =   31
         Top             =   3300
         Width           =   1305
      End
      Begin VB.Label lblFinishJob 
         Alignment       =   1  'Right Justify
         Caption         =   "lblFinishJob"
         Height          =   315
         Left            =   60
         TabIndex        =   30
         Top             =   2430
         Width           =   1665
      End
      Begin VB.Label lblBatchNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblBatchNo"
         Height          =   315
         Left            =   150
         TabIndex        =   29
         Top             =   1590
         Width           =   1575
      End
      Begin VB.Label lblStartJob 
         Alignment       =   1  'Right Justify
         Caption         =   "lblStartJob"
         Height          =   315
         Left            =   420
         TabIndex        =   28
         Top             =   2010
         Width           =   1305
      End
      Begin VB.Label lblJobDate 
         Alignment       =   1  'Right Justify
         Caption         =   "lblJobDate"
         Height          =   315
         Left            =   5460
         TabIndex        =   27
         Top             =   1110
         Width           =   1305
      End
      Begin Threed.SSCheck chkCommit 
         Height          =   435
         Left            =   6840
         TabIndex        =   4
         Top             =   1410
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
         TabIndex        =   26
         Top             =   2850
         Width           =   1605
      End
      Begin VB.Label lblJobNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblJobNo"
         Height          =   315
         Left            =   330
         TabIndex        =   25
         Top             =   1110
         Width           =   1395
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8490
         TabIndex        =   21
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJobWareHouse.frx":3884
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10170
         TabIndex        =   22
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
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJobWareHouse.frx":3B9E
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   19
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJobWareHouse.frx":3EB8
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditJobWareHouse"
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
Private m_JobWareHouse As CJobWareHouse
Private m_Jobs As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public JobDocType As Long
Public ProcessID As Long

Private FileName As String
Private m_SumUnit As Double
Private m_Employees As Collection
Private m_FormulaID As Long

Public TempCollection As Collection

Private Sub EnableDisableButton(En As Boolean)
   If En Then
      If ShowMode = SHOW_EDIT Then
         cmdAdd.Enabled = (m_JobWareHouse.COMMIT_FLAG = "N")
         cmdDelete.Enabled = (m_JobWareHouse.COMMIT_FLAG = "N")
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
      
      m_JobWareHouse.JOB_ID = ID
      m_JobWareHouse.QueryFlag = 1
      If Not glbProductionWH.QueryJob(m_JobWareHouse, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
        Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_JobWareHouse.PopulateFromRS(1, m_Rs)

      m_FormulaID = m_JobWareHouse.FORMULA_ID
      txtJobNo.Text = m_JobWareHouse.JOB_NO
      txtJobDesc.Text = m_JobWareHouse.JOB_DESC
      uctlJobDate.ShowDate = m_JobWareHouse.JOB_DATE
      txtBatchNo.Text = m_JobWareHouse.BATCH_NO
      ucltApproveByLookup.MyCombo.ListIndex = IDToListIndex(ucltApproveByLookup.MyCombo, m_JobWareHouse.APPROVED_BY)
      uctlResponseByLookup.MyCombo.ListIndex = IDToListIndex(uctlResponseByLookup.MyCombo, m_JobWareHouse.RESPONSE_BY)
      uctlStartJob.ShowDate = m_JobWareHouse.START_DATE
      uctlFinishJob.ShowDate = m_JobWareHouse.FINISH_DATE
      cboJobProcess.ListIndex = IDToListIndex(cboJobProcess, m_JobWareHouse.PROCESS_ID)
      
      chkCommit.Value = FlagToCheck(m_JobWareHouse.COMMIT_FLAG)
      chkCommit.Enabled = (m_JobWareHouse.COMMIT_FLAG = "N")
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

Private Function GetJobPartItemID(Col As Collection) As CJobInputWarehouse
Dim JO As CJobInputWarehouse
   
   For Each JO In Col
      If JO.Flag <> "D" Then
         Set GetJobPartItemID = JO
      End If
   Next JO
End Function

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim Ivd As CInventoryDoc
Dim JO As CJobInputWarehouse
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
      
   If Not CheckUniqueNs(JOB_NO, txtJobNo.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtJobNo.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Function
   End If
   
   If CountItem(m_JobWareHouse.Outputs) <> 1 Then
      glbErrorLog.LocalErrorMsg = "ข้อมูลผลิตภัณฑ์ที่ได้จะต้องมีเพียงแค่ 1 รายการเท่านั้น"
      glbErrorLog.ShowUserError
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_JobWareHouse.JOB_ID = ID
   m_JobWareHouse.AddEditMode = ShowMode
   m_JobWareHouse.JOB_NO = txtJobNo.Text
   m_JobWareHouse.JOB_DESC = txtJobDesc.Text
   m_JobWareHouse.JOB_DATE = uctlJobDate.ShowDate
'   m_JobWareHouse.BATCH_NO = txtBatchNo.Text
   m_JobWareHouse.APPROVED_BY = ucltApproveByLookup.MyCombo.ItemData(Minus2Zero(ucltApproveByLookup.MyCombo.ListIndex))
   m_JobWareHouse.RESPONSE_BY = uctlResponseByLookup.MyCombo.ItemData(Minus2Zero(uctlResponseByLookup.MyCombo.ListIndex))
   m_JobWareHouse.MIX_DATE = uctlJobDate.ShowDate
   m_JobWareHouse.START_DATE = uctlStartJob.ShowDate
   m_JobWareHouse.FINISH_DATE = uctlFinishJob.ShowDate
   m_JobWareHouse.PROCESS_ID = cboJobProcess.ItemData(Minus2Zero(cboJobProcess.ListIndex))
   m_JobWareHouse.COMMIT_FLAG = Check2Flag(chkCommit.Value)
   m_JobWareHouse.JOB_DOC_TYPE = JobDocType
   m_JobWareHouse.FORMULA_ID = m_FormulaID
   If m_JobWareHouse.Outputs.Count > 0 Then
      Set JO = GetJobPartItemID(m_JobWareHouse.Outputs)
      m_JobWareHouse.PART_ITEM_ID = JO.PART_ITEM_ID
      m_JobWareHouse.STD_AMOUNT = JO.STD_AMOUNT
      m_JobWareHouse.ACTUAL_AMOUNT = JO.TX_AMOUNT
   Else
      m_JobWareHouse.PART_ITEM_ID = -1
      m_JobWareHouse.STD_AMOUNT = 0
      m_JobWareHouse.ACTUAL_AMOUNT = 0
   End If
   Call EnableForm(Me, False)
   
   Call PopulateGuiID(m_JobWareHouse)
   
   If JobDocType = 1 Then
      Call glbDaily.Job2InventoryDoc2(m_JobWareHouse, Ivd, 1, 11)

      If (m_JobWareHouse.COMMIT_FLAG = "Y") Then
         If m_JobWareHouse.OLD_COMMIT_FLAG <> "Y" Then
            Call glbDaily.TriggerCommit(Ivd.ImportExports)
            If Not glbDaily.VerifyStockBalance(Ivd.ImportExports, glbErrorLog) Then
               Call EnableForm(Me, True)
               Exit Function
            End If
         End If
      End If
   End If
   
   Call glbDaily.StartTransaction
   If JobDocType = 1 Then
      If Not glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         SaveData = False
         Call glbDaily.RollbackTransaction
         Call EnableForm(Me, True)
         Exit Function
      End If
      m_JobWareHouse.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID
   Else
      m_JobWareHouse.INVENTORY_DOC_ID = -1
   End If

   If Not glbProductionWH.AddEditJob(m_JobWareHouse, IsOK, False, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call glbDaily.RollbackTransaction
      Call EnableForm(Me, True)
      Exit Function
   End If
   
   Call glbDaily.CommitTransaction
   
   
'''''   'เพิ่มในส่วนของบัญชี
'''''   m_Job.JOB_ID = ID
'''''   m_Job.AddEditMode = ShowMode
'''''   m_Job.JOB_NO = txtJobNo.Text
'''''   m_Job.JOB_DESC = txtJobDesc.Text
'''''   m_Job.JOB_DATE = uctlJobDate.ShowDate
''''''   m_Job.BATCH_NO = txtBatchNo.Text
'''''   m_Job.APPROVED_BY = ucltApproveByLookup.MyCombo.ItemData(Minus2Zero(ucltApproveByLookup.MyCombo.ListIndex))
'''''   m_Job.RESPONSE_BY = uctlResponseByLookup.MyCombo.ItemData(Minus2Zero(uctlResponseByLookup.MyCombo.ListIndex))
'''''   m_Job.START_DATE = uctlStartJob.ShowDate
'''''   m_Job.FINISH_DATE = uctlFinishJob.ShowDate
'''''   m_Job.PROCESS_ID = cboJobProcess.ItemData(Minus2Zero(cboJobProcess.ListIndex))
'''''   m_Job.COMMIT_FLAG = Check2Flag(chkCommit.Value)
'''''   m_Job.JOB_DOC_TYPE = JobDocType
'''''   m_Job.FORMULA_ID = m_FormulaID
'''''   If m_Job.Outputs.Count > 0 Then
'''''      Set JO = GetJobPartItemID(m_Job.Outputs)
'''''      m_Job.PART_ITEM_ID = JO.PART_ITEM_ID
'''''      m_Job.STD_AMOUNT = JO.STD_AMOUNT
'''''      m_Job.ACTUAL_AMOUNT = JO.TX_AMOUNT
'''''   Else
'''''      m_Job.PART_ITEM_ID = -1
'''''      m_Job.STD_AMOUNT = 0
'''''      m_Job.ACTUAL_AMOUNT = 0
'''''   End If
'''''
''''''   If m_Job.Inputs.Count > 0 Then
''''''      Set JO = GetJobPartItemID(m_Job.Inputs)
''''''      m_Job.PART_ITEM_ID = JO.PART_ITEM_ID
''''''      m_Job.STD_AMOUNT = JO.STD_AMOUNT
''''''      m_Job.ACTUAL_AMOUNT = JO.TX_AMOUNT
''''''   Else
''''''      m_Job.PART_ITEM_ID = -1
''''''      m_Job.STD_AMOUNT = 0
''''''      m_Job.ACTUAL_AMOUNT = 0
''''''   End If
''''''
'''''
'''''
'''''
'''''   Call EnableForm(Me, False)
'''''
'''''   Call PopulateGuiID(m_Job)
'''''
'''''   If JobDocType = 1 Then
'''''      Call glbDaily.Job2InventoryDoc(m_Job, Ivd, 1, 11)
'''''
'''''      If (m_Job.COMMIT_FLAG = "Y") Then
'''''         If m_Job.OLD_COMMIT_FLAG <> "Y" Then
'''''            Call glbDaily.TriggerCommit(Ivd.ImportExports)
'''''            If Not glbDaily.VerifyStockBalance(Ivd.ImportExports, glbErrorLog) Then
'''''               Call EnableForm(Me, True)
'''''               Exit Function
'''''            End If
'''''         End If
'''''      End If
'''''   End If
'''''
'''''   Call glbDaily.StartTransaction
'''''   If JobDocType = 1 Then
'''''      If Not glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog) Then
'''''         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'''''         SaveData = False
'''''         Call glbDaily.RollbackTransaction
'''''         Call EnableForm(Me, True)
'''''         Exit Function
'''''      End If
'''''
'''''      m_Job.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID
'''''   Else
'''''      m_Job.INVENTORY_DOC_ID = -1
'''''   End If
'''''
'''''   If Not glbProduction.AddEditJob(m_Job, IsOK, False, glbErrorLog) Then
'''''      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'''''      SaveData = False
'''''      Call glbDaily.RollbackTransaction
'''''      Call EnableForm(Me, True)
'''''      Exit Function
'''''   End If
'''''   Call glbDaily.CommitTransaction
'''''
'''''
'''''
'''''   '***************************
   
   
   
   
   
   If Not IsOK Then
      Call EnableForm(Me, True)
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   Call EnableForm(Me, True)
   SaveData = True
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

Private Sub chkCommit_Click(Value As Integer)
m_HasModify = True
End Sub

Public Sub RefreshGrid()
   GridEX1.ItemCount = CountItem(m_JobWareHouse.Verifies)
   GridEX1.Rebind
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
         Set frmAddEditJobInput.TempCollection = m_JobWareHouse.Inputs
         frmAddEditJobInput.ParentShowMode = ShowMode
         frmAddEditJobInput.ShowMode = SHOW_ADD
         frmAddEditJobInput.Area = 2 'โกดัง
         Set frmAddEditJobInput.ParentForm = Me
         frmAddEditJobInput.HeaderText = MapText("เพิ่มวัตถุดิบ")
         Load frmAddEditJobInput
         frmAddEditJobInput.Show 1
   
         OKClick = frmAddEditJobInput.OKClick
   
         Unload frmAddEditJobInput
         Set frmAddEditJobInput = Nothing
      
      If OKClick Then
         Call CalculateTotalRatio
         
         GridEX1.ItemCount = CountItem(m_JobWareHouse.Inputs)
         GridEX1.Rebind
      End If
 ElseIf TabStrip1.SelectedItem.Index = 2 Then
     Set frmAddEditJobWHOutputEx.TempCollection = m_JobWareHouse.Outputs
      'Set m_Job.Outputs = m_JobWareHouse.Outputs 'copy ไว้
      frmAddEditJobWHOutputEx.ParentShowMode = ShowMode
      frmAddEditJobWHOutputEx.ShowMode = SHOW_ADD
      frmAddEditJobWHOutputEx.HeaderText = MapText("เพิ่มผลผลิต")
      Load frmAddEditJobWHOutputEx
      frmAddEditJobWHOutputEx.Show 1

      OKClick = frmAddEditJobWHOutputEx.OKClick

      Unload frmAddEditJobWHOutputEx
      Set frmAddEditJobWHOutputEx = Nothing

      If OKClick Then
         Call CalculateTotalRatio
         
         GridEX1.ItemCount = CountItem(m_JobWareHouse.Outputs)
         GridEX1.Rebind
      End If
  
'   ElseIf TabStrip1.SelectedItem.Index = 3 Then
'     Set frmAddEditJobPeople.TempCollection = m_Job.Peoples
'      frmAddEditJobPeople.ParentShowMode = ShowMode
'      frmAddEditJobPeople.ShowMode = SHOW_ADD
'      frmAddEditJobPeople.HeaderText = MapText("เพิ่มแรงงาน")
'      Load frmAddEditJobPeople
'      frmAddEditJobPeople.Show 1
'
'      OKClick = frmAddEditJobPeople.OKClick
'
'      Unload frmAddEditJobPeople
'      Set frmAddEditJobPeople = Nothing
'
'      If OKClick Then
'         GridEX1.ItemCount = CountItem(m_Job.Peoples)
'         GridEX1.Rebind
'      End If
      
'   ElseIf TabStrip1.SelectedItem.Index = 4 Then
'     Set frmAddEditJobMachineEx.TempCollection = m_Job.Machines
'      frmAddEditJobMachineEx.ParentShowMode = ShowMode
'      frmAddEditJobMachineEx.ShowMode = SHOW_ADD
'      frmAddEditJobMachineEx.HeaderText = MapText("เพิ่มเครื่องจักรที่ใช้")
'      Load frmAddEditJobMachineEx
'      frmAddEditJobMachineEx.Show 1
'
'      OKClick = frmAddEditJobMachineEx.OKClick
'
'      Unload frmAddEditJobMachineEx
'      Set frmAddEditJobMachineEx = Nothing
'
'      If OKClick Then
'         GridEX1.ItemCount = CountItem(m_Job.Machines)
'         GridEX1.Rebind
'      End If
'  ElseIf TabStrip1.SelectedItem.Index = 5 Then
'     Set frmAddEditJobParameter.TempCollection = m_Job.Parameters
'     If cboJobProcess.ItemData(Minus2Zero(cboJobProcess.ListIndex)) <= 0 Then
'      glbErrorLog.LocalErrorMsg = "กรุณากรอก ข้อมูล โปรเซส ให้ครบถ้วน"
'       glbErrorLog.ShowUserError
'       Exit Sub
'      End If
'     frmAddEditJobParameter.Process = cboJobProcess.ItemData(Minus2Zero(cboJobProcess.ListIndex))
'      frmAddEditJobParameter.ParentShowMode = ShowMode
'      frmAddEditJobParameter.ShowMode = SHOW_ADD
'      frmAddEditJobParameter.HeaderText = MapText("เพิ่มพารามิเตอร์ที่ใช้")
'      Load frmAddEditJobParameter
'      frmAddEditJobParameter.Show 1
'
'      OKClick = frmAddEditJobParameter.OKClick
'
'      Unload frmAddEditJobParameter
'      Set frmAddEditJobParameter = Nothing
'
'      If OKClick Then
'         GridEX1.ItemCount = CountItem(m_Job.Parameters)
'         GridEX1.Rebind
'      End If
'   ElseIf TabStrip1.SelectedItem.Index = 6 Then
'     Set frmVerifyPartItemEx.TempCollection = m_Job.Verifies
'     Set frmVerifyPartItemEx.TempCollection2 = m_Job.Inputs
'     Set frmVerifyPartItemEx.Inputs = m_Job.Inputs
'      frmVerifyPartItemEx.ParentShowMode = ShowMode
'      Set frmVerifyPartItemEx.ParentForm = Me
'      frmVerifyPartItemEx.ShowMode = SHOW_ADD
'      frmVerifyPartItemEx.HeaderText = MapText("ตรวจสอบการใช้วัตถุดิบ")
'      Load frmVerifyPartItemEx
'      frmVerifyPartItemEx.Show 1
'
'      OKClick = frmVerifyPartItemEx.OKClick
'
'      Unload frmVerifyPartItemEx
'      Set frmVerifyPartItemEx = Nothing
'
'      If OKClick Then
'         GridEX1.ItemCount = CountItem(m_Job.Verifies)
'         GridEX1.Rebind
'      End If
   End If

   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdAuto_Click()
Dim No As String

   If Trim(txtJobNo.Text) = "" Then
      If JobDocType = 1 Then
         Call glbDatabaseMngr.GenerateNumber(JOBPLAN_NUMBER, No, glbErrorLog)
         txtJobNo.Text = No
      ElseIf JobDocType = 2 Then
         Call glbDatabaseMngr.GenerateNumber(ESTIMATE_NUMBER, No, glbErrorLog)
         txtJobNo.Text = No
      End If
   End If
End Sub

Private Sub CalculatePrice(PriceType As Long)
Dim D As CJobInputWarehouse
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

   For Each D In m_JobWareHouse.Inputs
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

   For Each D In m_JobWareHouse.Outputs
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

Private Sub cmdCalculate_Click()
Dim oMenu As cPopupMenu
Dim lMenuChosen As Long

   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("ดูข้อมูลสูตร", "-", "ปรับสูตร/ปรับปริมาณใหม่")
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
      
      GridEX1.MoveFirst
      
      Set frmFormulaSelect.Job = m_JobWareHouse
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
   End If
   
   If OKClick Then
      Call CalculateTotalRatio
      
      m_HasModify = True
      
      GridEX1.ItemCount = CountItem(m_JobWareHouse.Inputs)
      GridEX1.Rebind
   End If
      
   Set oMenu = Nothing
End Sub
Private Sub CalculateTotalRatio()
Dim D As CJobInputWarehouse
Dim O As CJobOutPutWarehouse
Dim Sum1 As Double
Dim Sum2 As Double
Dim Sum3 As Double

   Sum1 = 0
   Sum2 = 0
   Sum3 = 0
   For Each D In m_JobWareHouse.Inputs
      If D.Flag <> "D" Then
         Sum1 = Sum1 + D.TX_AMOUNT
      End If
   Next D
      
   For Each D In m_JobWareHouse.Outputs
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
         m_JobWareHouse.Inputs.Remove (ID2)
      Else
         m_JobWareHouse.Inputs.Item(ID2).Flag = "D"
      End If

      Call CalculateTotalRatio
      GridEX1.ItemCount = CountItem(m_JobWareHouse.Inputs)
      GridEX1.Rebind
      m_HasModify = True
ElseIf TabStrip1.SelectedItem.Index = 2 Then
      If ID1 <= 0 Then
         m_JobWareHouse.Outputs.Remove (ID2)
      Else
         m_JobWareHouse.Outputs.Item(ID2).Flag = "D"
      End If

      Call CalculateTotalRatio
      GridEX1.ItemCount = CountItem(m_JobWareHouse.Outputs)
      GridEX1.Rebind
      m_HasModify = True
     
'ElseIf TabStrip1.SelectedItem.Index = 3 Then
'      If ID1 <= 0 Then
'         m_Job.Peoples.Remove (ID2)
'      Else
'         m_Job.Peoples.Item(ID2).Flag = "D"
'      End If
'
'      GridEX1.ItemCount = CountItem(m_Job.Peoples)
'      GridEX1.Rebind
'      m_HasModify = True
'
'   ElseIf TabStrip1.SelectedItem.Index = 4 Then
'      If ID1 <= 0 Then
'         m_Job.Machines.Remove (ID2)
'      Else
'         m_Job.Machines.Item(ID2).Flag = "D"
'      End If
'
'      GridEX1.ItemCount = CountItem(m_Job.Machines)
'      GridEX1.Rebind
'      m_HasModify = True
'
'   ElseIf TabStrip1.SelectedItem.Index = 6 Then
'      If ID1 <= 0 Then
'         m_Job.Verifies.Remove (ID2)
'      Else
'         m_Job.Verifies.Item(ID2).Flag = "D"
'      End If
'
'      GridEX1.ItemCount = CountItem(m_Job.Verifies)
'      GridEX1.Rebind
'      m_HasModify = True
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
     Set frmAddEditJobInput.TempCollection = m_JobWareHouse.Inputs
      frmAddEditJobInput.ID = ID
      frmAddEditJobInput.ShowMode = SHOW_EDIT
      frmAddEditJobInput.COMMIT_FLAG = m_JobWareHouse.OLD_COMMIT_FLAG
      Set frmAddEditJobInput.ParentForm = Me
      frmAddEditJobInput.HeaderText = MapText("แก้ไขวัตถุดิบ")
      frmAddEditJobInput.Area = 2
      Load frmAddEditJobInput
      frmAddEditJobInput.Show 1

      OKClick = frmAddEditJobInput.OKClick

      Unload frmAddEditJobInput
      Set frmAddEditJobInput = Nothing

      If OKClick Then
         Call CalculateTotalRatio
         
         GridEX1.ItemCount = CountItem(m_JobWareHouse.Inputs)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
     Set frmAddEditJobWHOutputEx.TempCollection = m_JobWareHouse.Outputs
     frmAddEditJobWHOutputEx.COMMIT_FLAG = m_JobWareHouse.OLD_COMMIT_FLAG
      frmAddEditJobWHOutputEx.ID = ID
      frmAddEditJobWHOutputEx.ShowMode = SHOW_EDIT
      frmAddEditJobWHOutputEx.HeaderText = MapText("แก้ไขผลผลิต")
      Load frmAddEditJobWHOutputEx
      frmAddEditJobWHOutputEx.Show 1

      OKClick = frmAddEditJobWHOutputEx.OKClick

      Unload frmAddEditJobWHOutputEx
      Set frmAddEditJobWHOutputEx = Nothing

      If OKClick Then
         Call CalculateTotalRatio
         
         GridEX1.ItemCount = CountItem(m_JobWareHouse.Outputs)
         GridEX1.Rebind
      End If
'   ElseIf TabStrip1.SelectedItem.Index = 3 Then
'     Set frmAddEditJobPeople.TempCollection = m_Job.Peoples
'      frmAddEditJobPeople.ID = ID
'      frmAddEditJobPeople.ShowMode = SHOW_EDIT
'      frmAddEditJobPeople.HeaderText = MapText("แก้ไขเครื่องจักรที่ใช้")
'      Load frmAddEditJobPeople
'      frmAddEditJobPeople.Show 1
'
'      OKClick = frmAddEditJobPeople.OKClick
'
'      Unload frmAddEditJobPeople
'      Set frmAddEditJobPeople = Nothing
'
'      If OKClick Then
'         GridEX1.itemcount = CountItem(m_Job.Peoples)
'         GridEX1.Rebind
'      End If
'
'ElseIf TabStrip1.SelectedItem.Index = 4 Then
'     Set frmAddEditJobMachineEx.TempCollection = m_Job.Machines
'      frmAddEditJobMachineEx.ID = ID
'      frmAddEditJobMachineEx.ShowMode = SHOW_EDIT
'      frmAddEditJobMachineEx.HeaderText = MapText("แก้ไขเครื่องจักรที่ใช้")
'      Load frmAddEditJobMachineEx
'      frmAddEditJobMachineEx.Show 1
'
'      OKClick = frmAddEditJobMachineEx.OKClick
'
'      Unload frmAddEditJobMachineEx
'      Set frmAddEditJobMachineEx = Nothing
'
'      If OKClick Then
'         GridEX1.itemcount = CountItem(m_Job.Machines)
'         GridEX1.Rebind
'      End If
'ElseIf TabStrip1.SelectedItem.Index = 5 Then
'     Set frmAddEditJobParameter.TempCollection = m_Job.Parameters
'      frmAddEditJobParameter.ID = ID
'     frmAddEditJobParameter.Process = cboJobProcess.ItemData(Minus2Zero(cboJobProcess.ListIndex))
'      frmAddEditJobParameter.ShowMode = SHOW_EDIT
'      frmAddEditJobParameter.HeaderText = MapText("แก้ไขพารามิเตอร์ที่ใช้")
'      Load frmAddEditJobParameter
'      frmAddEditJobParameter.Show 1
'
'      OKClick = frmAddEditJobParameter.OKClick
'
'      Unload frmAddEditJobParameter
'      Set frmAddEditJobParameter = Nothing
'
'      If OKClick Then
'         GridEX1.itemcount = CountItem(m_Job.Parameters)
'         GridEX1.Rebind
'      End If
'   ElseIf TabStrip1.SelectedItem.Index = 6 Then
'     Set frmVerifyPartItemEx.TempCollection = m_Job.Verifies
'      frmVerifyPartItemEx.ID = ID
'      frmVerifyPartItemEx.ShowMode = SHOW_EDIT
'      Set frmVerifyPartItemEx.ParentForm = Me
'      frmVerifyPartItemEx.HeaderText = MapText("ตรวจสอบการใช้วัตถุดิบ")
'      Load frmVerifyPartItemEx
'      frmVerifyPartItemEx.Show 1
'
'      OKClick = frmVerifyPartItemEx.OKClick
'
'      Unload frmVerifyPartItemEx
'      Set frmVerifyPartItemEx = Nothing
'
'      If OKClick Then
'         GridEX1.itemcount = CountItem(m_Job.Verifies)
'         GridEX1.Rebind
'      End If
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
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
      Call Report.AddParam(m_JobWareHouse.JOB_ID, "JOB_ID")
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

Private Sub cmdSave_Click()
Dim Result As Boolean
   If Not SaveData Then
      Exit Sub
   End If
   
   ShowMode = SHOW_EDIT
   ID = m_JobWareHouse.JOB_ID
   m_JobWareHouse.QueryFlag = 1
   QueryData (True)
   m_HasModify = False
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
      
      Call LoadEmployee(uctlResponseByLookup.MyCombo)
      Call LoadProcess(cboJobProcess)
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_JobWareHouse.QueryFlag = 1
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         uctlJobDate.ShowDate = Now
         uctlStartJob.ShowDate = Now
         uctlFinishJob.ShowDate = Now
         cboJobProcess.ListIndex = IDToListIndex(cboJobProcess, ProcessID)
         
        m_JobWareHouse.QueryFlag = 0
         Call QueryData(False)
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
   
   Set m_JobWareHouse = Nothing
   Set m_Job = Nothing
   Set m_Jobs = Nothing
   Set m_Employees = Nothing
End Sub



Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   Debug.Print ColIndex & " " & NewColWidth
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
   GridEX1.Columns.Item(1).Visible = False

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
   GridEX1.Columns.Item(2).Visible = False
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2000
   Col.Caption = MapText("หมายเลขวัตถุดิบ")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 3500
   Col.Caption = MapText("ชื่อวัตถุดิบ")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 3330
   Col.Caption = MapText("ประเภทวัตถุดิบ")

   Set Col = GridEX1.Columns.add '6
   Col.Width = 2500
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
   GridEX1.Columns.Item(1).Visible = False

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
   GridEX1.Columns.Item(2).Visible = False
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2000
   Col.Caption = MapText("หมายเลขวัตถุดิบ")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 3500
   Col.Caption = MapText("ชื่อวัตถุดิบ")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 2500
   Col.Caption = MapText("ประเภทวัตถุดิบ")

   Set Col = GridEX1.Columns.add '6
   Col.Width = 2500
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
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitGrid1
   If JobDocType = 1 Then
   Call InitNormalLabel(lblJobNo, MapText("เลขที่ใบสั่งผลิต"))
   Call InitNormalLabel(lblJobDate, MapText("วันที่สั่งผลิต"))
   Else
   Call InitNormalLabel(lblJobNo, MapText("เลขที่ใบประเมิน"))
   Call InitNormalLabel(lblJobDate, MapText("วันที่ประเมิน"))
   
   End If
   Call InitNormalLabel(lblJobDesc, MapText("รายละเอียด"))
   
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
   
   Call txtJobNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtJobDesc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtBatchNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtInputAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtInputAmount.Enabled = False
   Call txtOutputAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtOutputAmount.Enabled = False
   
   Call InitCombo(uctlResponseByLookup.MyCombo)
   Call InitCombo(cboJobProcess)
   
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
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSave, MapText("บันทึก"))
   Call InitMainButton(cmdPrint, MapText("พิมพ์"))
   Call InitMainButton(cmdAuto, MapText("A"))
   Call InitMainButton(cmdCalculate, MapText("อื่น ๆ"))
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.NAME = GLB_FONT
   TabStrip1.Font.Size = 16
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.add().Caption = MapText("วัตถุดิบที่ใช้")
   TabStrip1.Tabs.add().Caption = MapText("ผลิตภัณฑ์ที่ได้")
'   TabStrip1.Tabs.add().Caption = MapText("แรงงาน")
'   TabStrip1.Tabs.add().Caption = MapText("เครื่องจักร")
'   TabStrip1.Tabs.add().Caption = MapText("ค่าใช้จ่ายผลิต")
'   If JobDocType = 1 Then
'      TabStrip1.Tabs.add().Caption = MapText("ตรวจสอบ")
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
   OKClick = False
   Call InitFormLayout
  ucltApproveByLookup.MyCombo.ListIndex = -1
   m_HasActivate = False
   m_HasModify = False
   
   m_FormulaID = -1
   Set m_Rs = New ADODB.Recordset
   Set m_Job = New CJob
   Set m_JobWareHouse = New CJobWareHouse
   Set m_Jobs = New Collection
   Set m_Employees = New Collection
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
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
     If m_JobWareHouse.Inputs Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim Ci As CJobInputWarehouse
      If m_JobWareHouse.Inputs.Count <= 0 Then
         Exit Sub
      End If
      Set Ci = GetItem(m_JobWareHouse.Inputs, RowIndex, RealIndex)
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
     If m_JobWareHouse.Outputs Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim CO As CJobInputWarehouse
      If m_JobWareHouse.Outputs.Count <= 0 Then
         Exit Sub
      End If
      Set CO = GetItem(m_JobWareHouse.Outputs, RowIndex, RealIndex)

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
'    ElseIf TabStrip1.SelectedItem.Index = 3 Then
'     If m_Job.Peoples Is Nothing Then
'         Exit Sub
'      End If
'
'      If RowIndex <= 0 Then
'         Exit Sub
'      End If
'
'      Dim CR As CJobResource
'      If m_Job.Peoples.Count <= 0 Then
'         Exit Sub
'      End If
'      Set CR = GetItem(m_Job.Peoples, RowIndex, RealIndex)
'      If CR Is Nothing Then
'         Exit Sub
'      End If
'      Values(1) = CR.JOB_RESOURCE_ID
'      Values(2) = RealIndex
'      Values(3) = CR.LONG_NAME & "  " & CR.LAST_NAME
'      Values(4) = CR.POSITION_NAME
'      Values(5) = FormatNumber(CR.OCCUPY_INTERVAL)
'      Values(6) = DateToStringExt(CR.OCCUPY_DATE)
'
'   ElseIf TabStrip1.SelectedItem.Index = 4 Then
'      If m_Job.Machines Is Nothing Then
'            Exit Sub
'         End If
'
'         If RowIndex <= 0 Then
'            Exit Sub
'         End If
'
'         Dim CP As CJobResource
'        Set CP = New CJobResource
'         If m_Job.Machines.Count <= 0 Then
'            Exit Sub
'         End If
'         Set CP = GetItem(m_Job.Machines, RowIndex, RealIndex)
'         If CP Is Nothing Then
'            Exit Sub
'         End If
'         Values(1) = CP.JOB_RESOURCE_ID
'         Values(2) = RealIndex
'         Values(3) = CP.MACHINE_NO
'         Values(4) = CP.MACHINE_NAME
'         Values(5) = FormatNumber(CP.OCCUPY_INTERVAL)
'         Values(6) = DateToStringExt(CP.OCCUPY_DATE)
'   ElseIf TabStrip1.SelectedItem.Index = 5 Then
'      If m_Job.Parameters Is Nothing Then
'            Exit Sub
'         End If
'
'         If RowIndex <= 0 Then
'            Exit Sub
'         End If
'
'         Dim Ca As CJobParameter
'        Set Ca = New CJobParameter
'         If m_Job.Parameters.Count <= 0 Then
'            Exit Sub
'         End If
'         Set Ca = GetItem(m_Job.Parameters, RowIndex, RealIndex)
'         If Ca Is Nothing Then
'            Exit Sub
'         End If
'         Values(1) = Ca.JOB_PARAMETER_ID
'         Values(2) = RealIndex
'         Values(3) = Ca.PARAMETER_PROCESS_NAME
'         Values(4) = Ca.JOB_PARAMETER_DESC
'         Values(5) = FormatNumber(Ca.PARAM_AMOUNT)
'   ElseIf TabStrip1.SelectedItem.Index = 6 Then
'      If m_Job.Parameters Is Nothing Then
'            Exit Sub
'         End If
'
'         If RowIndex <= 0 Then
'            Exit Sub
'         End If
'
'         Dim Jv As CJobVerify
'        Set Jv = New CJobVerify
'         If m_Job.Verifies.Count <= 0 Then
'            Exit Sub
'         End If
'         Set Jv = GetItem(m_Job.Verifies, RowIndex, RealIndex)
'         If Jv Is Nothing Then
'            Exit Sub
'         End If
'         Values(1) = Jv.JOB_VERIFY_ID
'         Values(2) = RealIndex
'         Values(3) = Jv.PART_NO
'         Values(4) = Jv.PART_DESC
'         Values(5) = Jv.NOTE
'         Values(6) = Jv.VERIFY_FLAG
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

Private Sub PopulateGuiID(Bd As CJobWareHouse)
Dim Di As CJobInputWarehouse

   For Each Di In Bd.Inputs
      If Di.Flag = "A" Then
         Di.LINK_ID = GetNextGuiID(Bd)
      End If
   Next Di

   For Each Di In Bd.Outputs
      If Di.Flag = "A" Then
         Di.LINK_ID = GetNextGuiID(Bd)
      End If
   Next Di
End Sub

Private Function GetNextGuiID(Bd As CJobWareHouse) As Long
Dim Di As CJobInputWarehouse
Dim MaxId As Long

   MaxId = 0
   For Each Di In Bd.Inputs
      If Di.LINK_ID > MaxId Then
         MaxId = Di.LINK_ID
      End If
   Next Di
   
   For Each Di In Bd.Outputs
      If Di.LINK_ID > MaxId Then
         MaxId = Di.LINK_ID
      End If
   Next Di
   
   GetNextGuiID = MaxId + 1
End Function

Private Sub SSCommand1_Click()

End Sub

Private Sub TabStrip1_Click()

   cmdAdd.Enabled = True
   cmdEdit.Enabled = True
   cmdDelete.Enabled = True
   If TabStrip1.SelectedItem.Index = 1 Then
      Call InitGrid1
      Call CalculateTotalRatio
      GridEX1.ItemCount = CountItem(m_JobWareHouse.Inputs)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
     Call InitGrid2
     GridEX1.ItemCount = CountItem(m_JobWareHouse.Outputs)
      GridEX1.Rebind
'   ElseIf TabStrip1.SelectedItem.Index = 3 Then
'       Call InitGrid3
'    GridEX1.ItemCount = CountItem(m_Job.Peoples)
'     GridEX1.Rebind
'   ElseIf TabStrip1.SelectedItem.Index = 4 Then
'     Call InitGrid4
'   GridEX1.ItemCount = CountItem(m_Job.Machines)
'   GridEX1.Rebind
'   ElseIf TabStrip1.SelectedItem.Index = 5 Then
'      cmdAdd.Enabled = False
'      cmdEdit.Enabled = False
'      cmdDelete.Enabled = False
'      Call InitGrid5
'      GridEX1.ItemCount = CountItem(m_Job.Parameters)
'      GridEX1.Rebind
'   ElseIf TabStrip1.SelectedItem.Index = 6 Then
'      cmdEdit.Enabled = False
'      Call InitGrid6
'      GridEX1.ItemCount = CountItem(m_Job.Verifies)
'      GridEX1.Rebind
   End If
End Sub


Private Sub txtBatchNO_Change()
   m_HasModify = True
End Sub

Private Sub txtJobDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtJobNo_Change()
   m_HasModify = True
End Sub

Private Sub ucltApproveByLookup_Change()
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
