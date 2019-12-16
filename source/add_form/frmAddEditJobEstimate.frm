VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddEditJobEstimate 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditJobEstimate.frx":0000
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
      TabIndex        =   18
      Top             =   0
      Width           =   12720
      _ExtentX        =   22437
      _ExtentY        =   15690
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboJobRes 
         Height          =   315
         Left            =   7440
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2280
         Width           =   3100
      End
      Begin VB.ComboBox cboJobRef 
         Height          =   315
         Left            =   7440
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2760
         Width           =   3100
      End
      Begin VB.ComboBox cboJobProcess 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2760
         Width           =   3100
      End
      Begin VB.ComboBox cboJobApp 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2280
         Width           =   3100
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   120
         TabIndex        =   0
         Top             =   3240
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
         TabIndex        =   19
         Top             =   0
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   3960
         Left            =   120
         TabIndex        =   12
         Top             =   3720
         Width           =   11640
         _ExtentX        =   20532
         _ExtentY        =   6985
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
         Column(1)       =   "frmAddEditJobEstimate.frx":27A2
         Column(2)       =   "frmAddEditJobEstimate.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditJobEstimate.frx":290E
         FormatStyle(2)  =   "frmAddEditJobEstimate.frx":2A6A
         FormatStyle(3)  =   "frmAddEditJobEstimate.frx":2B1A
         FormatStyle(4)  =   "frmAddEditJobEstimate.frx":2BCE
         FormatStyle(5)  =   "frmAddEditJobEstimate.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditJobEstimate.frx":2D5E
      End
      Begin prjFarmManagement.uctlTextBox txtJobNo 
         Height          =   435
         Left            =   1800
         TabIndex        =   1
         Top             =   840
         Width           =   3855
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtJobDesc 
         Height          =   435
         Left            =   7440
         TabIndex        =   2
         Top             =   840
         Width           =   3855
         _ExtentX        =   6535
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlJobDate 
         Height          =   405
         Left            =   1800
         TabIndex        =   3
         Top             =   1320
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtBatchNo 
         Height          =   435
         Left            =   7440
         TabIndex        =   4
         Top             =   1320
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlFinishJob 
         Height          =   405
         Left            =   7440
         TabIndex        =   7
         Top             =   1800
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlDate uctlStartJob 
         Height          =   405
         Left            =   1800
         TabIndex        =   6
         Top             =   1800
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin VB.Label lblJobProcess 
         Alignment       =   1  'Right Justify
         Caption         =   "lblJobProcess"
         Height          =   315
         Left            =   480
         TabIndex        =   29
         Top             =   2760
         Width           =   1305
      End
      Begin VB.Label lblJobRef 
         Alignment       =   1  'Right Justify
         Caption         =   "lblJobRef"
         Height          =   315
         Left            =   4920
         TabIndex        =   28
         Top             =   2880
         Width           =   2385
      End
      Begin VB.Label lblJobRes 
         Alignment       =   1  'Right Justify
         Caption         =   "lblJobRes"
         Height          =   315
         Left            =   5160
         TabIndex        =   27
         Top             =   2400
         Width           =   2265
      End
      Begin VB.Label lblJobApp 
         Alignment       =   1  'Right Justify
         Caption         =   "lblJobApp"
         Height          =   315
         Left            =   480
         TabIndex        =   26
         Top             =   2400
         Width           =   1305
      End
      Begin VB.Label lblFinishJob 
         Alignment       =   1  'Right Justify
         Caption         =   "lblFinishJob"
         Height          =   315
         Left            =   5760
         TabIndex        =   25
         Top             =   1920
         Width           =   1665
      End
      Begin VB.Label lblBatchNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblBatchNo"
         Height          =   315
         Left            =   5760
         TabIndex        =   24
         Top             =   1440
         Width           =   1665
      End
      Begin VB.Label lblStartJob 
         Alignment       =   1  'Right Justify
         Caption         =   "lblStartJob"
         Height          =   315
         Left            =   480
         TabIndex        =   23
         Top             =   1920
         Width           =   1305
      End
      Begin VB.Label lblJobDate 
         Alignment       =   1  'Right Justify
         Caption         =   "lblJobDate"
         Height          =   315
         Left            =   480
         TabIndex        =   22
         Top             =   1440
         Width           =   1305
      End
      Begin Threed.SSCheck chkCommit 
         Height          =   435
         Left            =   8880
         TabIndex        =   5
         Top             =   1320
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
         Left            =   5760
         TabIndex        =   21
         Top             =   960
         Width           =   1665
      End
      Begin VB.Label lblJobNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblJobNo"
         Height          =   315
         Left            =   480
         TabIndex        =   20
         Top             =   960
         Width           =   1305
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8520
         TabIndex        =   16
         Top             =   7800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJobEstimate.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10200
         TabIndex        =   17
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJobEstimate.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   15
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJobEstimate.frx":356A
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditJobEstimate"
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
Private m_Jobs As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long

Private FileName As String
Private m_SumUnit As Double
Public TempCollection As Collection
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
      Call m_Job.PopulateFromRS(m_Rs)

  txtJobNo.Text = m_Job.JOB_NO
   txtJobDesc.Text = m_Job.JOB_DESC
      uctlJobDate.ShowDate = m_Job.JOB_DATE
       txtBatchNo.Text = m_Job.BATCH_NO
       cboJobApp.ListIndex = IDToListIndex(cboJobApp, m_Job.APPROVED_BY)
       cboJobRes.ListIndex = IDToListIndex(cboJobRes, m_Job.RESPONSE_BY)
     uctlStartJob.ShowDate = m_Job.START_DATE
     uctlFinishJob.ShowDate = m_Job.FINISH_DATE
      cboJobProcess.ListIndex = IDToListIndex(cboJobProcess, m_Job.PROCESS_ID)
     cboJobRef.ListIndex = IDToListIndex(cboJobRef, m_Job.INVENTORY_DOC_ID)
   chkCommit.Value = FlagToCheck(m_Job.COMMIT_FLAG)
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
Private Function SaveData() As Boolean
Dim IsOK As Boolean

   If ShowMode = SHOW_ADD Then
      If Not VerifyAccessRight("PRODUCT_ESTIMATE_ADD") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   ElseIf ShowMode = SHOW_EDIT Then
      If Not VerifyAccessRight("PRODUCT_ESTIMATE_EDIT") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   End If
If Not VerifyTextControl(lblJobNo, txtJobNo, False) Then
      Exit Function
  End If
   If Not VerifyTextControl(lblJobDesc, txtJobDesc, False) Then
    Exit Function
   End If
   If Not VerifyDate(lblJobDate, uctlJobDate, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblBatchNo, txtBatchNo, False) Then
    Exit Function
   End If

   If Not VerifyCombo(lblJobApp, cboJobApp, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblJobRes, cboJobRes, False) Then
      Exit Function
   End If
   
If Not VerifyDate(lblStartJob, uctlStartJob, False) Then
  Exit Function
End If
If Not VerifyDate(lblFinishJob, uctlFinishJob, False) Then
  Exit Function
End If
 
   If Not VerifyCombo(lblJobProcess, cboJobProcess, False) Then
      Exit Function
   End If
   
   If Not VerifyCombo(lblJobRef, cboJobRef, False) Then
      Exit Function
   End If
   
   If Not CheckUniqueNs(JOB_NO, txtJobNo.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtJobNo.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   m_Job.JOB_ID = ID
   m_Job.AddEditMode = ShowMode
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
   m_Job.COMMIT_FLAG = Check2Flag(chkCommit.Value)
   
   Call EnableForm(Me, False)
   If Not glbProduction.AddEditJob(m_Job, IsOK, True, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call EnableForm(Me, True)
      Exit Function
   End If
   If Not IsOK Then
      Call EnableForm(Me, True)
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   Call EnableForm(Me, True)
   SaveData = True
End Function
Private Sub cboJobApp_Change()
m_HasModify = True
End Sub

Private Sub cboJobApp_Click()
m_HasModify = True
End Sub

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

Private Sub cboJobRes_Change()
m_HasModify = True
End Sub

Private Sub cboJobRes_Click()
m_HasModify = True
End Sub

Private Sub chkCommit_Click(Value As Integer)
m_HasModify = True
End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean

   If Not cmdAdd.Enabled Then
      Exit Sub
   End If

   OKClick = False
    If TabStrip1.SelectedItem.Index = 1 Then
     Set frmAddEditJobInputEstimate.TempCollection = m_Job.Inputs
      frmAddEditJobInputEstimate.ParentShowMode = ShowMode
      frmAddEditJobInputEstimate.ShowMode = SHOW_ADD
      frmAddEditJobInputEstimate.HeaderText = MapText("เพิ่มวัตถุดิบ")
      Load frmAddEditJobInputEstimate
      frmAddEditJobInputEstimate.Show 1

      OKClick = frmAddEditJobInputEstimate.OKClick

      Unload frmAddEditJobInputEstimate
      Set frmAddEditJobInputEstimate = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Job.Inputs)
         GridEX1.Rebind
      End If
 ElseIf TabStrip1.SelectedItem.Index = 2 Then
     Set frmAddEditJobOutputEstimate.TempCollection = m_Job.Outputs
      frmAddEditJobOutputEstimate.ParentShowMode = ShowMode
      frmAddEditJobOutputEstimate.ShowMode = SHOW_ADD
      frmAddEditJobOutputEstimate.HeaderText = MapText("เพิ่มผลผลิต")
      Load frmAddEditJobOutputEstimate
      frmAddEditJobOutputEstimate.Show 1

      OKClick = frmAddEditJobOutputEstimate.OKClick

      Unload frmAddEditJobOutputEstimate
      Set frmAddEditJobOutputEstimate = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Job.Outputs)
         GridEX1.Rebind
      End If
  
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
     Set frmAddEditJobPeopleEstimate.TempCollection = m_Job.Peoples
      frmAddEditJobPeopleEstimate.ParentShowMode = ShowMode
      frmAddEditJobPeopleEstimate.ShowMode = SHOW_ADD
      frmAddEditJobPeopleEstimate.HeaderText = MapText("เพิ่มแรงงาน")
      Load frmAddEditJobPeopleEstimate
      frmAddEditJobPeopleEstimate.Show 1

      OKClick = frmAddEditJobPeopleEstimate.OKClick

      Unload frmAddEditJobPeopleEstimate
      Set frmAddEditJobPeopleEstimate = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Job.Peoples)
         GridEX1.Rebind
      End If
      
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
     Set frmAddEditJobMachineEstimate.TempCollection = m_Job.Machines
      frmAddEditJobMachineEstimate.ParentShowMode = ShowMode
      frmAddEditJobMachineEstimate.ShowMode = SHOW_ADD
      frmAddEditJobMachineEstimate.HeaderText = MapText("เพิ่มเครื่องจักรที่ใช้")
      Load frmAddEditJobMachineEstimate
      frmAddEditJobMachineEstimate.Show 1

      OKClick = frmAddEditJobMachineEstimate.OKClick

      Unload frmAddEditJobMachineEstimate
      Set frmAddEditJobMachineEstimate = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Job.Machines)
         GridEX1.Rebind
      End If
       
   End If

   If OKClick Then
      m_HasModify = True
   End If
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
         m_Job.Inputs.Remove (ID2)
      Else
         m_Job.Inputs.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_Job.Inputs)
      GridEX1.Rebind
      m_HasModify = True
ElseIf TabStrip1.SelectedItem.Index = 2 Then
      If ID1 <= 0 Then
         m_Job.Outputs.Remove (ID2)
      Else
         m_Job.Outputs.Item(ID2).Flag = "D"
      End If

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
   
   
   End If
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
   ID = Val(GridEX1.Value(2))
   OKClick = False
         If TabStrip1.SelectedItem.Index = 1 Then
     Set frmAddEditJobInputEstimate.TempCollection = m_Job.Inputs
      frmAddEditJobInputEstimate.ID = ID
      frmAddEditJobInputEstimate.ShowMode = SHOW_EDIT
      frmAddEditJobInputEstimate.HeaderText = MapText("แก้ไขวัตถุดิบ")
      Load frmAddEditJobInputEstimate
      frmAddEditJobInputEstimate.Show 1

      OKClick = frmAddEditJobInputEstimate.OKClick

      Unload frmAddEditJobInputEstimate
      Set frmAddEditJobInputEstimate = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Job.Inputs)
         GridEX1.Rebind
      End If
ElseIf TabStrip1.SelectedItem.Index = 2 Then
     Set frmAddEditJobOutputEstimate.TempCollection = m_Job.Outputs
      frmAddEditJobOutputEstimate.ID = ID
      frmAddEditJobOutputEstimate.ShowMode = SHOW_EDIT
      frmAddEditJobOutputEstimate.HeaderText = MapText("แก้ไขผลผลิตที่ใช้")
      Load frmAddEditJobOutputEstimate
      frmAddEditJobOutputEstimate.Show 1

      OKClick = frmAddEditJobOutputEstimate.OKClick

      Unload frmAddEditJobOutputEstimate
      Set frmAddEditJobOutputEstimate = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Job.Outputs)
         GridEX1.Rebind
      End If
ElseIf TabStrip1.SelectedItem.Index = 3 Then
     Set frmAddEditJobPeopleEstimate.TempCollection = m_Job.Peoples
      frmAddEditJobPeopleEstimate.ID = ID
      frmAddEditJobPeopleEstimate.ShowMode = SHOW_EDIT
      frmAddEditJobPeopleEstimate.HeaderText = MapText("แก้ไขเครื่องจักรที่ใช้")
      Load frmAddEditJobPeopleEstimate
      frmAddEditJobPeopleEstimate.Show 1

      OKClick = frmAddEditJobPeopleEstimate.OKClick

      Unload frmAddEditJobPeopleEstimate
      Set frmAddEditJobPeopleEstimate = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Job.Peoples)
         GridEX1.Rebind
      End If

ElseIf TabStrip1.SelectedItem.Index = 4 Then
     Set frmAddEditJobMachineEstimate.TempCollection = m_Job.Machines
      frmAddEditJobMachineEstimate.ID = ID
      frmAddEditJobMachineEstimate.ShowMode = SHOW_EDIT
      frmAddEditJobMachineEstimate.HeaderText = MapText("แก้ไขเครื่องจักรที่ใช้")
      Load frmAddEditJobMachineEstimate
      frmAddEditJobMachineEstimate.Show 1

      OKClick = frmAddEditJobMachineEstimate.OKClick

      Unload frmAddEditJobMachineEstimate
      Set frmAddEditJobMachineEstimate = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Job.Machines)
         GridEX1.Rebind
      End If
   
   
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

Private Sub cmdPictureAdd_Click()
On Error Resume Next
Dim strDescription As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Picture Files (*.jpg, *.gif)|*.jpg;*.gif"
   dlgAdd.DialogTitle = "Select Picture to Add to Database"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
   m_HasModify = True
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      Call LoadEmployee(cboJobApp)
      Call LoadEmployee(cboJobRes)
            Call LoadRefDoc(cboJobRef)
            Call LoadProcess(cboJobProcess)
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_Job.QueryFlag = 1
         Call QueryData(True)
         
      ElseIf ShowMode = SHOW_ADD Then
        m_Job.QueryFlag = 0
         Call QueryData(False)
      End If
      Call TabStrip1_Click
      Call EnableForm(Me, True)
      m_HasModify = False
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
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
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_Job = Nothing
   Set m_Jobs = Nothing
End Sub



Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''''''debug.print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid1()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.Name = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"
GridEX1.Columns.Item(1).Visible = False
   Set Col = GridEX1.Columns.Add '2
   Col.Width = 0
   Col.Caption = "Real ID"
   GridEX1.Columns.Item(2).Visible = False
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 2000
   Col.Caption = MapText("หมายเลขวัตถุดิบ")

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 3500
   Col.Caption = MapText("ชื่อวัตถุดิบ")

   Set Col = GridEX1.Columns.Add '5
   Col.Width = 2500
   Col.Caption = MapText("ประเภทวัตถุดิบ")

   Set Col = GridEX1.Columns.Add '6
   Col.Width = 2500
   Col.Caption = MapText("จำนวน")

   Set Col = GridEX1.Columns.Add '7
   Col.Width = 2500
   Col.Caption = MapText("รหัสสถานที่เบิก")
   
   Set Col = GridEX1.Columns.Add '8
   Col.Width = 2500
   Col.Caption = MapText("สถานที่เบิก")
   
   Set Col = GridEX1.Columns.Add '9
   Col.Width = 2500
   Col.Caption = MapText("รหัสเชื่อมโยงวัตถุดิบ")
   
   Set Col = GridEX1.Columns.Add '10
   Col.Width = 2500
   Col.Caption = MapText("รหัสสินค้าขาย")
   
   Set Col = GridEX1.Columns.Add '11
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
   GridEX1.ColumnHeaderFont.Name = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"
  GridEX1.Columns.Item(1).Visible = False
   Set Col = GridEX1.Columns.Add '2
   Col.Width = 0
   Col.Caption = "Real ID"
  GridEX1.Columns.Item(2).Visible = False

   Set Col = GridEX1.Columns.Add '3
   Col.Width = 2000
   Col.Caption = MapText("หมายเลขผลิตภัณฑ์")

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 3500
   Col.Caption = MapText("ชื่อผลิตภัณฑ์")

   Set Col = GridEX1.Columns.Add '5
   Col.Width = 2500
   Col.Caption = MapText("ประเภทผลิตภัณฑ์")

   Set Col = GridEX1.Columns.Add '6
   Col.Width = 2500
   Col.Caption = MapText("จำนวน")

   Set Col = GridEX1.Columns.Add '7
   Col.Width = 2500
   Col.Caption = MapText("รหัสสถานที่จัดเก็บ")
   
   Set Col = GridEX1.Columns.Add '8
   Col.Width = 2500
   Col.Caption = MapText("สถานที่จัดเก็บ")
   
   Set Col = GridEX1.Columns.Add '9
   Col.Width = 2500
   Col.Caption = MapText("รหัสเชื่อมโยงผลิตภัณฑ์")
   
   Set Col = GridEX1.Columns.Add '10
   Col.Width = 2500
   Col.Caption = MapText("รหัสสินค้าขาย")
   
   Set Col = GridEX1.Columns.Add '11
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
   GridEX1.ColumnHeaderFont.Name = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"
GridEX1.Columns.Item(1).Visible = False
   Set Col = GridEX1.Columns.Add '2
   Col.Width = 0
   Col.Caption = "Real ID"
GridEX1.Columns.Item(2).Visible = False
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 3500
   Col.Caption = MapText("ชื่อพนักงาน")

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 2610
   Col.Caption = MapText("ตำแหน่ง")

 Set Col = GridEX1.Columns.Add '5
   Col.Width = 3000
   Col.Caption = MapText("จำนวนเวลา(ชั่วโมง/นาที)")
 Set Col = GridEX1.Columns.Add '6
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
   GridEX1.ColumnHeaderFont.Name = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"
GridEX1.Columns.Item(1).Visible = False
   Set Col = GridEX1.Columns.Add '2
   Col.Width = 0
   Col.Caption = "Real ID"
GridEX1.Columns.Item(2).Visible = False
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 2610
   Col.Caption = MapText("หมายเลขเครื่องจักร")

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 3500
   Col.Caption = MapText("ชื่อเครื่องจักร")

   Set Col = GridEX1.Columns.Add '5
   Col.Width = 3000
   Col.Caption = MapText("จำนวนเวลา(ชั่วโมง/นาที)")

 Set Col = GridEX1.Columns.Add '6
   Col.Width = 2500
   Col.Caption = MapText("วันที่ใช้เครื่องจักร")
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
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
   
   Call InitCheckBox(chkCommit, "งานเสร็จแล้ว")
   
   Call txtJobNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtJobDesc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtBatchNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   Call InitCombo(cboJobApp)
   Call InitCombo(cboJobRes)
   Call InitCombo(cboJobProcess)
   Call InitCombo(cboJobRef)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   TabStrip1.Font.Bold = True
   TabStrip1.Font.Name = GLB_FONT
   TabStrip1.Font.Size = 16
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.Add().Caption = MapText("           วัตถุดิบที่ใช้          ")
   TabStrip1.Tabs.Add().Caption = MapText("        ผลิตภัณฑ์ที่ได้          ")
   TabStrip1.Tabs.Add().Caption = MapText("          แรงงานที่ใช้         ")
   TabStrip1.Tabs.Add().Caption = MapText("        เครื่องจักรที่ใช้         ")
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
  cboJobApp.ListIndex = -1
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_Job = New CJob
   Set m_Jobs = New Collection
   Set TempCollection = New Collection
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Index = 1 Then
     If m_Job.Inputs Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim CI As CJobInput
      If m_Job.Inputs.Count <= 0 Then
         Exit Sub
      End If
      Set CI = GetItem(m_Job.Inputs, RowIndex, RealIndex)
      If CI Is Nothing Then
         Exit Sub
      End If
      Values(1) = CI.JOB_INOUT_ID
      Values(2) = RealIndex
      Values(3) = CI.PART_NO
      Values(4) = CI.PART_DESC
      Values(5) = CI.PART_TYPE_NAME
      Values(6) = CI.TX_AMOUNT
      Values(7) = CI.LOCATION_NO
      Values(8) = CI.LOCATION_NAME
      Values(9) = CI.LINK_ID
      Values(10) = CI.SERIAL_NUMBER
     Values(11) = CI.INOUT_REF
      
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
      Values(6) = CO.TX_AMOUNT
      Values(7) = CO.LOCATION_NO
      Values(8) = CO.LOCATION_NAME
      Values(9) = CO.LINK_ID
      Values(10) = CO.SERIAL_NUMBER
     Values(11) = CO.INOUT_REF
      
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
      Values(5) = CR.EMP_ID_HOUR & " ชั่วโมง " & CR.EMP_ID_HOURN & " นาที "
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
      Values(5) = CP.MACHINE_ID_HOUR & " ชั่วโมง " & CP.MACHINE_ID_HOURN & " นาที "
      Values(6) = DateToStringExt(CP.OCCUPY_DATE)
     
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



Private Sub TabStrip1_Click()

   If TabStrip1.SelectedItem.Index = 1 Then
   Call InitGrid1
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
   
   End If
End Sub


Private Sub txtBatchNo_Change()
m_HasModify = True
End Sub

Private Sub txtJobDesc_Change()
m_HasModify = True
End Sub

Private Sub txtJobNo_Change()
m_HasModify = True
End Sub

Private Sub uctlFinishJob_HasChange()
m_HasModify = True
End Sub

Private Sub uctlJobDate_HasChange()
m_HasModify = True
End Sub

Private Sub uctlStartJob_HasChange()
m_HasModify = True
End Sub
