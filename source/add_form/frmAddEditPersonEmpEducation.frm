VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditPersonEmpEducation 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5760
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditPersonEmpEducation.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   5760
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   4125
      Left            =   0
      TabIndex        =   9
      Top             =   600
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   7276
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboQualification 
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   9
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1680
         Width           =   3855
      End
      Begin prjFarmManagement.uctlTextBox txtInstitute 
         Height          =   435
         Left            =   1560
         TabIndex        =   0
         Top             =   270
         Width           =   3855
         _ExtentX        =   9763
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtScore 
         Height          =   435
         Left            =   1560
         TabIndex        =   5
         Top             =   2640
         Width           =   3855
         _ExtentX        =   14790
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlInDate 
         Height          =   405
         Left            =   1560
         TabIndex        =   1
         Top             =   720
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlDate uctlOutDate 
         Height          =   405
         Left            =   1560
         TabIndex        =   2
         Top             =   1200
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtMaster 
         Height          =   435
         Left            =   1560
         TabIndex        =   4
         Top             =   2160
         Width           =   3855
         _ExtentX        =   14790
         _ExtentY        =   767
      End
      Begin VB.Label lblMaster 
         Alignment       =   1  'Right Justify
         Caption         =   "lblMaster"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   2280
         Width           =   1365
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   960
         TabIndex        =   6
         Top             =   3360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPersonEmpEducation.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   3240
         TabIndex        =   7
         Top             =   3360
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblScore 
         Alignment       =   1  'Right Justify
         Caption         =   "lblScore"
         Height          =   375
         Left            =   0
         TabIndex        =   14
         Top             =   2640
         Width           =   1485
      End
      Begin VB.Label lblQualification 
         Alignment       =   1  'Right Justify
         Caption         =   "lblQualification"
         Height          =   465
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   1365
      End
      Begin VB.Label lblInstitute 
         Alignment       =   1  'Right Justify
         Caption         =   "lblInstitute"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label lblInDate 
         Alignment       =   1  'Right Justify
         Caption         =   "lblInDate"
         Height          =   375
         Left            =   0
         TabIndex        =   11
         Top             =   720
         Width           =   1485
      End
      Begin VB.Label lblOutDate 
         Alignment       =   1  'Right Justify
         Caption         =   "lblOutDate"
         Height          =   375
         Left            =   0
         TabIndex        =   10
         Top             =   1230
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmAddEditPersonEmpEducation"
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
Private Sub cboQualification_Change()
m_HasModify = True
End Sub
Private Sub cboQualification_Click()
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
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
         
   Call InitNormalLabel(lblInstitute, MapText("สถาบัน"))
   Call InitNormalLabel(lblInDate, MapText("จากวันที่"))
   Call InitNormalLabel(lblOutDate, MapText("ถึงวันที่"))
   Call InitNormalLabel(lblQualification, MapText("วุฒิ"))
   Call InitNormalLabel(lblScore, MapText("เกรดเฉลี่ย"))
   Call InitNormalLabel(lblMaster, MapText("วิชาหลัก"))
   
   Call InitCombo(cboQualification)

   Call txtInstitute.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtMaster.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtScore.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   
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
                 Dim D As CEmpEducation
         Set D = TempCollection.Item(ID)
         
         txtInstitute.Text = D.INSTITUTE
         cboQualification.ListIndex = IDToListIndex(cboQualification, D.QUALIFICATION_ID)
         uctlInDate.ShowDate = D.FROM_DATE
         uctlOutDate.ShowDate = D.TO_DATE
         txtMaster.Text = D.MASTER
         txtScore.Text = D.SCORE

      End If
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long
   If Not VerifyTextControl(lblInstitute, txtInstitute) Then
      Exit Function
   End If
   If Not VerifyDate(lblInDate, uctlInDate, True) Then
      Exit Function
   End If
   If Not VerifyDate(lblOutDate, uctlOutDate, True) Then
      Exit Function
   End If
   If Not VerifyCombo(lblQualification, cboQualification) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblScore, txtScore) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim D As CEmpEducation
   If ShowMode = SHOW_ADD Then
      Set D = New CEmpEducation
         Else
      Set D = TempCollection.Item(ID)
   End If
   
   D.INSTITUTE = txtInstitute.Text
   D.QUALIFICATION_NAME = cboQualification.Text
   D.QUALIFICATION_ID = cboQualification.ItemData(Minus2Zero(cboQualification.ListIndex))
   D.FROM_DATE = uctlInDate.ShowDate
   D.TO_DATE = uctlOutDate.ShowDate
   D.MASTER = txtMaster.Text
   D.SCORE = Val(txtScore.Text)
   

   If ShowMode = SHOW_ADD Then
      D.Flag = "A"
      Call TempCollection.add(D)
Else
      If D.Flag <> "A" Then
      D.Flag = "E"
      End If
         End If
   
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
            Call LoadQualificationType(cboQualification)
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
         Call QueryData(False)
      End If
      
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

Private Sub SSCommand2_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
End Sub



Private Sub txtInstitute_Change()
  m_HasModify = True
End Sub

Private Sub txtMaster_Change()
  m_HasModify = True
End Sub

Private Sub txtScore_Change()
  m_HasModify = True
End Sub

Private Sub uctlInDate_HasChange()
  m_HasModify = True
End Sub

Private Sub uctlOutDate_HasChange()
  m_HasModify = True
End Sub
