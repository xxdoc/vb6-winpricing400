VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditMoneyPerson 
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
   Icon            =   "frmAddEditMoneyPerson.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   9435
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   4605
      Left            =   0
      TabIndex        =   10
      Top             =   600
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   8123
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextBox txtAmount 
         Height          =   435
         Left            =   1785
         TabIndex        =   6
         Top             =   2060
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtLendCode 
         Height          =   435
         Left            =   1800
         TabIndex        =   0
         Top             =   360
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtLender 
         Height          =   435
         Left            =   1800
         TabIndex        =   3
         Top             =   1200
         Width           =   5115
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtDesc 
         Height          =   435
         Left            =   1800
         TabIndex        =   5
         Top             =   1640
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlLendDate 
         Height          =   405
         Left            =   1800
         TabIndex        =   2
         Top             =   800
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin Threed.SSCommand cmdAuto 
         Height          =   405
         Left            =   3840
         TabIndex        =   1
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditMoneyPerson.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdLayout 
         Height          =   405
         Left            =   6960
         TabIndex        =   4
         Top             =   1200
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditMoneyPerson.frx":0BE4
         ButtonStyle     =   3
      End
      Begin VB.Label lblUnit 
         Caption         =   "lblUnit"
         Height          =   375
         Left            =   3840
         TabIndex        =   16
         Top             =   2130
         Width           =   1005
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3090
         TabIndex        =   7
         Top             =   2760
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditMoneyPerson.frx":0EFE
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4740
         TabIndex        =   8
         Top             =   2760
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblAmount"
         Height          =   375
         Left            =   195
         TabIndex        =   15
         Top             =   2160
         Width           =   1485
      End
      Begin VB.Label lblLendDate 
         Alignment       =   1  'Right Justify
         Caption         =   "lblLendDate"
         Height          =   375
         Left            =   195
         TabIndex        =   14
         Top             =   840
         Width           =   1485
      End
      Begin VB.Label lblLender 
         Alignment       =   1  'Right Justify
         Caption         =   "lblLender"
         Height          =   375
         Left            =   195
         TabIndex        =   13
         Top             =   1260
         Width           =   1485
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "lblDesc"
         Height          =   375
         Left            =   195
         TabIndex        =   12
         Top             =   1710
         Width           =   1485
      End
      Begin VB.Label lblLendCode 
         Alignment       =   1  'Right Justify
         Caption         =   "lblLendCode"
         Height          =   375
         Left            =   180
         TabIndex        =   11
         Top             =   420
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmAddEditMoneyPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_EmpReceivable As CEmpReceivable
Private m_EmpReceivables As Collection
Private m_Employee As CEmployee
Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Private FileName As String
Private m_SumUnit As Double
Public Header As String
Public TempCollection As Collection
Public TempCollection2 As Collection

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

   Call InitNormalLabel(lblLendCode, MapText("หมายเลขยืม"))
   Call InitNormalLabel(lblLendDate, MapText("วันที่ยืม"))
   Call InitNormalLabel(lblLender, MapText("ผู้ยืม"))
   Call InitNormalLabel(lblDesc, MapText("รายละเอียด"))
   Call InitNormalLabel(lblAmount, MapText("จำนวน"))
   Call InitNormalLabel(lblUnit, MapText("บาท"))
   
   Call txtLendCode.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtLender.SetTextLenType(TEXT_STRING, glbSetting.NAME_TYPE)
   
   Call txtDesc.SetTextLenType(TEXT_STRING, glbSetting.ADDRESS_TYPE)
   Call txtAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   
txtLender.Enabled = False
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdLayout.Picture = LoadPicture(glbParameterObj.NormalButton1)
cmdAuto.Picture = LoadPicture(glbParameterObj.NormalButton1)
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdLayout, MapText("..."))
   Call InitMainButton(cmdAuto, MapText(" อัตโนมัติ "))
End Sub
Private Sub cmdAuto_Click()
Dim No As String

   Call glbDatabaseMngr.GenerateNumber(BORROW_NUMBER, No, glbErrorLog)
   txtLendCode.Text = No
End Sub
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long
IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
      m_EmpReceivable.EMP_RECEIVABLE_ID = ID
      m_EmpReceivable.QueryFlag = 1
      If Not glbDaily.QueryEmpReceivable(m_EmpReceivable, m_Rs, itemcount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
        Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If itemcount > 0 Then
      Call m_EmpReceivable.PopulateFromRS(m_Rs)
   txtLendCode.Text = m_EmpReceivable.BORROW_NO
    uctlLendDate.ShowDate = m_EmpReceivable.BORROW_DATE
    txtLender.Text = m_EmpReceivable.LONG_NAME & " " & m_EmpReceivable.LAST_NAME
       txtDesc.Text = m_EmpReceivable.BORROW_DESC
    txtAmount.Text = FormatNumber(m_EmpReceivable.BORROW_AMOUNT)
    cmdLayout.Tag = m_EmpReceivable.EMP_ID
      End If
   Call EnableForm(Me, True)
End Sub

Private Sub cmdLayout_Click()
Dim OKClick As Boolean
Dim LayoutID As Long
Dim TempID As Long
Dim TempStr As String

   Load frmDataPersonMoney
   frmDataPersonMoney.Show 1
   If frmDataPersonMoney.OKClick Then
      TempID = frmDataPersonMoney.PersonID
      TempStr = frmDataPersonMoney.PersonName
   Else
      TempID = Val(cmdLayout.Tag)
      TempStr = txtLender.Text
   End If
   
   Unload frmDataPersonMoney
   Set frmDataPersonMoney = Nothing
   
   cmdLayout.Tag = TempID
   txtLender.Text = TempStr
   m_HasModify = True
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
Dim itemcount As Long
If Not VerifyTextControl(lblLender, txtLender, False) Then
     Exit Function
  End If
   If Not VerifyTextControl(lblLendCode, txtLendCode, False) Then
     Exit Function
  End If
  If Not VerifyTextControl(lblDesc, txtDesc, False) Then
      Exit Function
   End If
  If Not VerifyTextControl(lblAmount, txtAmount, False) Then
      Exit Function
   End If
If Not VerifyDate(lblLendDate, uctlLendDate, False) Then
      Exit Function
   End If
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_EmpReceivable.EMP_RECEIVABLE_ID = ID
   m_EmpReceivable.BORROW_NO = txtLendCode.Text
      m_EmpReceivable.BORROW_DATE = uctlLendDate.ShowDate
      m_EmpReceivable.EMP_ID = cmdLayout.Tag
       m_EmpReceivable.EMP_NAME = txtLender.Text
      m_EmpReceivable.BORROW_DESC = txtDesc.Text
      m_EmpReceivable.BORROW_AMOUNT = Val(txtAmount.Text)
      m_EmpReceivable.CLOSED_FLAG = "N"
      
   Call EnableForm(Me, False)
   m_EmpReceivable.AddEditMode = ShowMode
   If Not glbDaily.AddEditEmpReceivable(m_EmpReceivable, IsOK, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call EnableForm(Me, True)
      Exit Function
   End If
   
        m_Employee.EMP_ID = m_EmpReceivable.EMP_ID
                 If Not glbDaily.QueryEmployeeMoney(m_Employee, m_Rs, itemcount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Function
      End If
      Call m_Employee.PopulateFromRSMoney(1, m_Rs)
      m_Employee.EMP_ID = cmdLayout.Tag
m_Employee.TOTBORROW = m_Employee.TOTBORROW + Val(txtAmount.Text)
    If Not glbDaily.AddEditEmployeeMoney(m_Employee, IsOK, True, glbErrorLog) Then
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
   Set m_EmpReceivable = New CEmpReceivable
   Set m_EmpReceivables = New Collection
Set m_Employee = New CEmployee
End Sub


Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_EmpReceivables = Nothing
   Set m_EmpReceivable = Nothing
   Set m_Employee = Nothing
End Sub

Private Sub txtAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtKeyName_Change()
   m_HasModify = True
End Sub

Private Sub txtThaiMsg_Change()
   m_HasModify = True
End Sub

Private Sub txtAmphur_Change()
   m_HasModify = True
End Sub

Private Sub txtDistrict_Change()
   m_HasModify = True
End Sub

Private Sub txtFax_Change()
   m_HasModify = True
End Sub

Private Sub txtHomeNo_Change()
   m_HasModify = True
End Sub

Private Sub txtQuantity_Change()
   m_HasModify = True
End Sub

Private Sub txtPhone_Change()
   m_HasModify = True
End Sub

Private Sub txtProvince_Change()
   m_HasModify = True
End Sub

Private Sub txtRoad_Change()
   m_HasModify = True
End Sub

Private Sub txtSoi_Change()
   m_HasModify = True
End Sub

Private Sub txtPrice_Change()
   m_HasModify = True
End Sub

Private Sub txtZipcode_Change()
   m_HasModify = True
End Sub

Private Sub uctlLayoutLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlToLocationLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlLocationLookup_Change()
   m_HasModify = True
End Sub
Private Sub txtLendCode_Change()
   m_HasModify = True
End Sub

Private Sub txtLender_Change()
   m_HasModify = True
End Sub
Private Sub uctlLendDate_HasChange()
   m_HasModify = True
End Sub
