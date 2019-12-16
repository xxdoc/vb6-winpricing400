VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditSupplierAccount 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9435
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditSupplierAccount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   9435
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   3645
      Left            =   0
      TabIndex        =   8
      Top             =   600
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   6429
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextBox txtSupAccNo 
         Height          =   435
         Left            =   1800
         TabIndex        =   0
         Top             =   240
         Width           =   6465
         _ExtentX        =   6112
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtSupAccName 
         Height          =   435
         Left            =   1800
         TabIndex        =   1
         Top             =   720
         Width           =   6465
         _ExtentX        =   11404
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtSupAccBank 
         Height          =   435
         Left            =   1800
         TabIndex        =   2
         Top             =   1200
         Width           =   6465
         _ExtentX        =   11404
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtSupAccBranch 
         Height          =   435
         Left            =   1800
         TabIndex        =   3
         Top             =   1680
         Width           =   6465
         _ExtentX        =   11404
         _ExtentY        =   767
      End
      Begin VB.Label lblSupAccNo 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label lblSupAccBranch 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1680
         Width           =   1485
      End
      Begin Threed.SSCheck chkUseTransFlag 
         Height          =   255
         Left            =   1800
         TabIndex        =   4
         Top             =   2280
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   450
         _Version        =   131073
         Caption         =   "sscUseTransFlag"
      End
      Begin VB.Label lblSupAccName 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   10
         Top             =   780
         Width           =   1485
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3120
         TabIndex        =   5
         Top             =   2880
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditSupplierAccount.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4800
         TabIndex        =   6
         Top             =   2880
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblSupAccBank 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmAddEditSupplierAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Header As String
Public ShowMode As SHOW_MODE_TYPE
Public ParentShowMode As SHOW_MODE_TYPE
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public TempCollection As Collection
Public COMMIT_FLAG As String

Private Sub chkUseTransFlag_Click(Value As Integer)
m_HasModify = True
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If

   OKClick = False
   Unload Me
End Sub

Private Sub InitFormLayout()
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)

   Me.KeyPreview = True
   pnlHeader.Caption = HeaderText
   Me.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_HEAD_COLOR
   SSFrame1.BackColor = GLB_FORM_COLOR

   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText

   Call InitNormalLabel(lblSupAccNo, MapText("เลขที่บัญชี"))
   Call InitNormalLabel(lblSupAccName, MapText("ชื่อบัญชี"))
   Call InitNormalLabel(lblSupAccBank, MapText("ชื่อธนาคาร"))
   Call InitNormalLabel(lblSupAccBranch, MapText("สาขา"))
   Call InitCheckBox(chkUseTransFlag, "แสดงในค่าขนส่ง")
   'chkFlag
   Call txtSupAccNo.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtSupAccName.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtSupAccBank.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtSupAccBranch.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)

   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)

      If ShowMode = SHOW_EDIT Then
         Dim EnpAcc As CSupplierAccount

         Set EnpAcc = TempCollection.Item(ID)

         txtSupAccNo.Text = EnpAcc.SUPPLIER_ACCOUNT_NO
         txtSupAccName.Text = EnpAcc.SUPPLIER_ACCOUNT_NAME
         txtSupAccBank.Text = EnpAcc.SUPPLIER_ACCOUNT_BANK
         txtSupAccBranch.Text = EnpAcc.SUPPLIER_ACCOUNT_BRANCH
         chkUseTransFlag.Value = FlagToCheck(EnpAcc.USE_TRANSPORT_FLAG)

      End If
   End If

   Call EnableForm(Me, True)
End Sub


Private Sub cmdOK_Click()
   If Not cmdOK.Enabled Then
      Exit Sub
   End If

   If Not SaveData Then
      Exit Sub
   End If

   OKClick = True
   Unload Me
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long

   If Not VerifyTextControl(lblSupAccNo, txtSupAccNo, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblSupAccName, txtSupAccName, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblSupAccBank, txtSupAccBank, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblSupAccBranch, txtSupAccBranch, False) Then
      Exit Function
   End If

   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
'   Dim D As CSupplierAccount 'ตรวจสอบใน Collection ก่อนการบันทึกด้วย
'   Set D = GetObject("CSupplierAccount", TempCollection, Trim(txtSupAccNo.Text), False)
'   If Not D Is Nothing Then
'       glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtSupAccNo.Text & " " & MapText("อยู่ในระบบแล้ว")
'      glbErrorLog.ShowUserError
'      Call EnableForm(Me, True)
'      txtSupAccNo.SetFocus
'      Exit Function
'   End If
'
'   If Not CheckUniqueNs(SUPPLIER_ACCOUNT_NO_UNIQUE, txtSupAccNo.Text, ID) Then
'      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtSupAccNo.Text & " " & MapText("อยู่ในระบบแล้ว")
'      glbErrorLog.ShowUserError
'      Call EnableForm(Me, True)
'      txtSupAccNo.SetFocus
'      Exit Function
'   End If

   Dim EnpAcc As CSupplierAccount
   If ShowMode = SHOW_ADD Then
      Set EnpAcc = New CSupplierAccount
      EnpAcc.Flag = "A"
      Call TempCollection.add(EnpAcc, Trim(txtSupAccNo.Text))
   Else
      Set EnpAcc = TempCollection.Item(ID)
      If EnpAcc.Flag <> "A" Then
         EnpAcc.Flag = "E"
      End If
   End If
   
   EnpAcc.SUPPLIER_ACCOUNT_NO = txtSupAccNo.Text
   EnpAcc.SUPPLIER_ACCOUNT_NAME = txtSupAccName.Text
   EnpAcc.SUPPLIER_ACCOUNT_BANK = txtSupAccBank.Text
   EnpAcc.SUPPLIER_ACCOUNT_BRANCH = txtSupAccBranch.Text
   EnpAcc.USE_TRANSPORT_FLAG = Check2Flag(chkUseTransFlag.Value)

   Set EnpAcc = Nothing
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents

      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
         Call QueryData(True)
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

Private Sub Form_Load()

   OKClick = False
   Call InitFormLayout

   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
End Sub


Private Sub txtSupAccBank_Change()
m_HasModify = True
End Sub

Private Sub txtSupAccBranch_Change()
m_HasModify = True
End Sub

Private Sub txtSupAccName_Change()
m_HasModify = True
End Sub

Private Sub txtSupAccNo_Change()
m_HasModify = True
End Sub
