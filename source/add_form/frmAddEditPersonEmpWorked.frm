VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditPersonEmpWorked 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditPersonEmpWorked.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   8910
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   4005
      Left            =   0
      TabIndex        =   8
      Top             =   600
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   7064
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboResign 
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
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2160
         Width           =   2685
      End
      Begin prjFarmManagement.uctlTextBox txtWorkPlace 
         Height          =   435
         Left            =   1800
         TabIndex        =   0
         Top             =   270
         Width           =   6735
         _ExtentX        =   9763
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlInDate 
         Height          =   405
         Left            =   1800
         TabIndex        =   1
         Top             =   720
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlDate uctlOutDate 
         Height          =   405
         Left            =   1800
         TabIndex        =   2
         Top             =   1200
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtPosition 
         Height          =   435
         Left            =   1800
         TabIndex        =   3
         Top             =   1680
         Width           =   6735
         _ExtentX        =   9763
         _ExtentY        =   767
      End
      Begin VB.Label lblResign 
         Alignment       =   1  'Right Justify
         Caption         =   "lblResign"
         Height          =   465
         Left            =   240
         TabIndex        =   13
         Top             =   2160
         Width           =   1485
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2280
         TabIndex        =   5
         Top             =   3000
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPersonEmpWorked.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4920
         TabIndex        =   6
         Top             =   3000
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblPosition 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPosition"
         Height          =   465
         Left            =   240
         TabIndex        =   12
         Top             =   1680
         Width           =   1485
      End
      Begin VB.Label lblWorkPlace 
         Alignment       =   1  'Right Justify
         Caption         =   "lblWorkPlace"
         Height          =   375
         Left            =   0
         TabIndex        =   11
         Top             =   300
         Width           =   1725
      End
      Begin VB.Label lblInDate 
         Alignment       =   1  'Right Justify
         Caption         =   "lblInDate"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   780
         Width           =   1485
      End
      Begin VB.Label lblOutDate 
         Alignment       =   1  'Right Justify
         Caption         =   "lblOutDate"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1230
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmAddEditPersonEmpWorked"
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

Private Sub cboResign_Change()
m_HasModify = True
End Sub

Private Sub cboResign_Click()
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
      
   Call InitNormalLabel(lblWorkPlace, MapText("สถานที่ทำงาน"))
   Call InitNormalLabel(lblInDate, MapText("จากวันที่"))
   Call InitNormalLabel(lblOutDate, MapText("ถึงวันที่"))
   Call InitNormalLabel(lblPosition, MapText("ตำแหน่ง"))
   Call InitNormalLabel(lblResign, MapText("สาเหตุที่ออก"))
   
   Call txtWorkPlace.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtPosition.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitCombo(cboResign)
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
                 Dim D As CEmpWorked
         Set D = TempCollection.Item(ID)
         
         txtWorkPlace.Text = D.WORK_PLACE
        
         cboResign.ListIndex = IDToListIndex(cboResign, D.RESIGN_REASON)
         uctlInDate.ShowDate = D.FROM_DATE
         uctlOutDate.ShowDate = D.TO_DATE
         txtPosition.Text = D.EMP_POSITION

        
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

   If Not VerifyTextControl(lblWorkPlace, txtWorkPlace, False) Then
      Exit Function
   End If
      'If Not VerifyCombo(lblPosition, cboPosition, False) Then
    '  Exit Function
   'End If

   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
      
      Dim D As CEmpWorked
   If ShowMode = SHOW_ADD Then
      Set D = New CEmpWorked
   Else
      Set D = TempCollection.Item(ID)
   End If
   
   D.WORK_PLACE = txtWorkPlace.Text
D.RESIGN_REASON_NAME = cboResign.Text
   D.RESIGN_REASON = cboResign.ItemData(Minus2Zero(cboResign.ListIndex))
   D.FROM_DATE = uctlInDate.ShowDate
   D.TO_DATE = uctlOutDate.ShowDate
   D.EMP_POSITION = txtPosition.Text
   
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
            Call LoadResignReason(cboResign)
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



Private Sub txtPosition_Change()
m_HasModify = True
End Sub

Private Sub txtWorkPlace_Change()
m_HasModify = True
End Sub

Private Sub uctlInDate_HasChange()
  m_HasModify = True
End Sub

Private Sub uctlOutDate_HasChange()
  m_HasModify = True
End Sub
