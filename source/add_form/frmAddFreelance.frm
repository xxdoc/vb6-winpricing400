VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddFreelance 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15195
   Icon            =   "frmAddFreelance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   15195
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3765
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   6641
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextBox txtName 
         Height          =   435
         Left            =   1860
         TabIndex        =   1
         Top             =   1440
         Width           =   4485
         _extentx        =   13309
         _extenty        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   2520
         TabIndex        =   7
         Top             =   -120
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtCode 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   990
         Width           =   2955
         _extentx        =   5212
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtLastName 
         Height          =   435
         Left            =   1860
         TabIndex        =   2
         Top             =   1890
         Width           =   4485
         _extentx        =   13309
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlFreelanceLookup 
         Height          =   435
         Left            =   1920
         TabIndex        =   12
         Top             =   2400
         Width           =   5385
         _extentx        =   9499
         _extenty        =   767
      End
      Begin Threed.SSCheck chkEmpResignFlag 
         Height          =   345
         Left            =   6480
         TabIndex        =   11
         Top             =   1920
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   609
         _Version        =   131073
         Caption         =   "SSCheck1"
         TripleState     =   -1  'True
      End
      Begin VB.Label lblLastName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   10
         Top             =   1950
         Width           =   1575
      End
      Begin Threed.SSCheck chkPigFlag 
         Height          =   345
         Left            =   6420
         TabIndex        =   5
         Top             =   1020
         Visible         =   0   'False
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   609
         _Version        =   131073
         Caption         =   "SSCheck1"
         TripleState     =   -1  'True
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   9
         Top             =   1500
         Width           =   1575
      End
      Begin VB.Label lblCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   8
         Top             =   1080
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5235
         TabIndex        =   4
         Top             =   2910
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3585
         TabIndex        =   3
         Top             =   2910
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddFreelance.frx":27A2
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddFreelance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Freelance As CFreelance
Public ParentForm As Form
Public TempCollection As Collection
Public TempFreelance As Collection

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String

Private Sub cboPartType_Click()
   m_HasModify = True
End Sub

Private Sub chkEmpResignFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkPigFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If

   OKClick = True
   Unload Me
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)

      m_Freelance.FREELANCE_ID = ID
      m_Freelance.QueryFlag = 1
      If Not glbDaily.QueryFreelance(m_Freelance, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If

   If ItemCount > 0 Then
      Call m_Freelance.PopulateFromRS(1, m_Rs)

      txtCode.Text = m_Freelance.FREELANCE_CODE
      txtName.Text = m_Freelance.FREELANCE_NAME
      txtLastName.Text = m_Freelance.FREELANCE_LASTNAME
      chkEmpResignFlag = FlagToCheck(m_Freelance.FREELANCE_RESIGN_FLAG)
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
   
'   If Not VerifyTextControl(lblLastName, txtLastName, False) Then
'      Exit Function
'   End If

   
'   If Not m_HasModify Then
'      SaveData = True
'      Exit Function
'   End If
   
   m_Freelance.FREELANCE_ID = ID
   m_Freelance.AddEditMode = ShowMode
   m_Freelance.FREELANCE_RESIGN_FLAG = Check2Flag(chkEmpResignFlag.Value)
   m_Freelance.FREELANCE_CODE = txtCode.Text
   m_Freelance.FREELANCE_NAME = txtName.Text
   m_Freelance.FREELANCE_LASTNAME = txtLastName.Text
'   uctlFreelanceLookup.MyCombo.ListIndex = IDToListIndex(uctlFreelanceLookup.MyCombo, m_Customer.FREELANCE_ID)
      
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditFreelance(m_Freelance, IsOK, True, glbErrorLog) Then
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

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadFreelance(uctlFreelanceLookup.MyCombo, TempFreelance)
      Set uctlFreelanceLookup.MyCollection = TempFreelance
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
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
   ElseIf Shift = 0 And KeyCode = 117 Then
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

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblCode, MapText("รหัสพนักงาน"))
   Call InitNormalLabel(lblName, MapText("ชื่อ"))
   Call InitNormalLabel(lblLastName, MapText("นามสกุล"))
   
   Call InitCheckBox(chkPigFlag, "PIG FLAG")
   Call InitCheckBox(chkEmpResignFlag, "ลาออก")
   
   Call txtName.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   Call txtCode.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtLastName.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
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
   
   Set m_Freelance = New CFreelance
   Set m_Rs = New ADODB.Recordset
   Set TempFreelance = New Collection
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set TempFreelance = Nothing
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

Private Sub uctlFreelanceLookup_Change()
   m_HasModify = True
End Sub
