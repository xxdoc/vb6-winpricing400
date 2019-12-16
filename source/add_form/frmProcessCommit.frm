VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProcssCommit 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Icon            =   "frmProcessCommit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   10410
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   4005
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   7064
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlDate uctlFromDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   0
         Top             =   1020
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   1860
         TabIndex        =   4
         Top             =   2370
         Width           =   8235
         _ExtentX        =   14526
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   10
         Top             =   0
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtPercent 
         Height          =   465
         Left            =   1860
         TabIndex        =   5
         Top             =   2700
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlDate uctlToDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   1
         Top             =   1470
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin Threed.SSCheck chkNewStatus 
         Height          =   465
         Left            =   4230
         TabIndex        =   3
         Top             =   1890
         Visible         =   0   'False
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   820
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin Threed.SSCheck chkOldStatus 
         Height          =   465
         Left            =   1890
         TabIndex        =   2
         Top             =   1890
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   820
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   15
         Top             =   1530
         Width           =   1575
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   1860
         TabIndex        =   6
         Top             =   3180
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmProcessCommit.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3600
         TabIndex        =   14
         Top             =   2820
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   13
         Top             =   2430
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   12
         Top             =   2850
         Width           =   1575
      End
      Begin VB.Label lblFileName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   11
         Top             =   1080
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   8535
         TabIndex        =   8
         Top             =   3180
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   6885
         TabIndex        =   7
         Top             =   3180
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmProcessCommit.frx":2ABC
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmProcssCommit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Employee As CEmployee

Public id As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String

Private m_TempSearchs1 As Collection
Private m_MovementItemSearchs1 As Collection
Private m_MovementItemSearchs2 As Collection
Private m_MovementItemSearchs3 As Collection

Private m_ProductStatuss As Collection

Public DocumentCategory As Long
Public DocumentType As Long
Public Area As Long

Private Sub cmdPasswd_Click()

End Sub


Private Sub cboPartType_Click()
   m_HasModify = True
End Sub

Private Sub cboPosition_Click()
   m_HasModify = True
End Sub

Private Sub chkPigFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cboCommitType_Click()
   m_HasModify = True
End Sub

Private Sub chkBalanceFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
'   Call LoadAccessRight(Nothing, glbAccessRight, glbUser.GROUP_ID)
   OKClick = True
   Unload Me
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
   End If
   
   If ItemCount > 0 Then
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_Employee.EMP_ID = id
   m_Employee.AddEditMode = ShowMode
   m_Employee.PASS_STATUS = "Y"
   
   m_Employee.EmpName.AddEditMode = ShowMode
   m_Employee.EName.AddEditMode = ShowMode
      
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditEmployee(m_Employee, IsOK, True, glbErrorLog) Then
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

Private Sub cmdStart_Click()
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim Ivd As CBillingDoc
Dim IsOK As Boolean
Dim iCount As Long
Dim RecordCount As Long
Dim Percent As Double
Dim I As Long
Dim HasBegin As Boolean
Dim Result As Boolean

'   If Not VerifyDate(lblFileName, uctlFromDate, False) Then
'      Exit Sub
'   End If

   HasBegin = False
   prgProgress.MAX = 100
   prgProgress.MIN = 0
   
   Set Ivd = New CBillingDoc
   Ivd.BILLING_DOC_ID = -1
   Ivd.DOCUMENT_TYPE = DocumentType
   Ivd.COMMIT_FLAG = Check2Flag(chkOldStatus.Value)
   Ivd.FROM_DATE = uctlFromDate.ShowDate
   Ivd.TO_DATE = uctlToDate.ShowDate
   Call glbDaily.QueryBillingDoc(Ivd, m_Rs, iCount, IsOK, glbErrorLog)
   
   I = 0
   While Not m_Rs.EOF
      I = I + 1
      prgProgress.Value = MyDiff(I, iCount) * 100
      txtPercent.Text = prgProgress.Value
      
      Call Ivd.PopulateFromRS(1, m_Rs)
      
      Set BD = New CBillingDoc
      BD.BILLING_DOC_ID = Ivd.BILLING_DOC_ID
      Call glbDaily.CopyBillingDoc(BD, IsOK, True, Area, 10, glbErrorLog)
      Set BD = Nothing
      
      m_Rs.MoveNext
   Wend
    prgProgress.Value = 100
    
   Set Ivd = Nothing
   Set BD = Nothing
   Exit Sub
   
ErrorHandler:
   If HasBegin Then
      Call glbDaily.RollbackTransaction
   End If
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadProductStatus(Nothing, m_ProductStatuss)
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         id = 0
      End If
      
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
'      Call cmdAdd_Click
      KeyCode = 0
'   ElseIf Shift = 0 And KeyCode = 117 Then
'      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
'      Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 123 Then
      Call AddMemoNote
      KeyCode = 0
   End If

End Sub

Private Sub ResetStatus()
   prgProgress.MAX = 100
   prgProgress.MIN = 0
   prgProgress.Value = 0
   txtPercent.Text = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = MapText("คัดลอกข้อมูลแบบกำหนดเอง")
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitNormalLabel(lblFileName, "จากวันที่")
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(Label1, "%")
   Call InitNormalLabel(lblToDate, "ถึงวันที่")
   
   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPercent.Enabled = False
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdStart.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitCheckBox(chkOldStatus, "คำนวณแล้ว")
   Call InitCheckBox(chkNewStatus, "คัดลอกแล้วเปลี่ยนสถานะเป็นคำนวณ")
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdStart, MapText("เริ่ม"))
   
   Call ResetStatus
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   Call EnableForm(Me, False)
   m_HasActivate = False
   
   Set m_Employee = New CEmployee
   Set m_Rs = New ADODB.Recordset
   
   Set m_MovementItemSearchs1 = New Collection
   Set m_MovementItemSearchs2 = New Collection
   Set m_MovementItemSearchs3 = New Collection
   Set m_ProductStatuss = New Collection
   Set m_TempSearchs1 = New Collection
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub txtLastName_Change()
   m_HasModify = True
End Sub

Private Sub txtCode_Change()
   m_HasModify = True
End Sub

Private Sub txtName_Change()
   m_HasModify = True
End Sub

Private Sub txtMaxUser_Change()
   m_HasModify = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_TempSearchs1 = Nothing
   Set m_MovementItemSearchs1 = Nothing
   Set m_MovementItemSearchs2 = Nothing
   Set m_MovementItemSearchs3 = Nothing
   Set m_ProductStatuss = Nothing
End Sub

Private Sub uctlToDate_HasChange()
   m_HasModify = True
End Sub
