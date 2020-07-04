VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditDeliveryCus 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Icon            =   "frmAddEditDeliveryCus.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   10410
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3165
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   5583
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextBox txtName 
         Height          =   435
         Left            =   1860
         TabIndex        =   1
         Top             =   1440
         Width           =   7845
         _extentx        =   13309
         _extenty        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   5
         Top             =   0
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
         Width           =   2355
         _extentx        =   2672
         _extenty        =   767
      End
      Begin Threed.SSCheck chkHide 
         Height          =   435
         Left            =   1860
         TabIndex        =   8
         Top             =   1920
         Visible         =   0   'False
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "chkHide"
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   7
         Top             =   1500
         Width           =   1575
      End
      Begin VB.Label lblCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   6
         Top             =   1080
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5280
         TabIndex        =   3
         Top             =   2400
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3600
         TabIndex        =   2
         Top             =   2400
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditDeliveryCus.frx":27A2
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditDeliveryCus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Header As String
Public ShowMode As SHOW_MODE_TYPE
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_m_ExDeliveryCostItem As Collection

Public HeaderText As String
Public id As Long
Public ID2 As Long
Public OKClick As Boolean
Public ParentForm As Form
Public TempCollection As Collection
Public TempFreelance As Collection
Public CustomerCode As String

Private Sub cmdCancel_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub InitFormLayout()
   Me.KeyPreview = True
   pnlHeader.Caption = HeaderText
   Me.BackColor = GLB_FORM_COLOR
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   Call InitNormalLabel(lblCode, MapText("รหัสสถานที่"))
   Call InitNormalLabel(lblName, MapText("ชื่อสถานที่จัดส่ง"))
'   Call InitCheckBox(chkHide, "ยกเลิก")
   
   Call txtCode.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtName.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)

   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim DC As CDeliveryCus

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
      
         Set DC = TempCollection.Item(id)
         
         txtCode.Text = DC.DELIVERY_CUS_ITEM_CODE
         txtName.Text = DC.DELIVERY_CUS_ITEM_NAME
'         chkHide.Value = FlagToCheck(DC.HIDE_FLAG)
      Else
        txtCode.Text = CustomerCode & "-"
      End If
   End If
   
   Call EnableForm(Me, True)
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long

   If Not VerifyTextControl(lblCode, txtCode, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblName, txtName, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim D As CDeliveryCus 'ตรวจสอบใน Collection ก่อนการบันทึกด้วย
   Set D = GetObject("CDeliveryCus", TempCollection, Trim(txtCode.Text), False)
   If Not D Is Nothing Then
       glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtName.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      txtName.SetFocus
      Exit Function
   End If
   
   If Not CheckUniqueNs(DELIVERY_CUS_UNIQUE, txtCode.Text, ID2) Then
      glbErrorLog.LocalErrorMsg = MapText("มี  " & lblCode.Caption & " ") & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      txtName.SetFocus
      Exit Function
   End If
   
   Dim DC As CDeliveryCus
   If ShowMode = SHOW_ADD Then
      Set DC = New CDeliveryCus
      DC.Flag = "A"
      DC.DELIVERY_CUS_ITEM_CODE = txtCode.Text
      DC.DELIVERY_CUS_ITEM_NAME = txtName.Text
'      DC.HIDE_FLAG = Check2Flag(chkHide.Value)
      Call TempCollection.add(DC, Trim(txtCode.Text))
   Else
      Set DC = TempCollection.Item(id)
      If DC.Flag <> "A" Then
         DC.Flag = "E"
         DC.DELIVERY_CUS_ITEM_CODE = txtCode.Text
         DC.DELIVERY_CUS_ITEM_NAME = txtName.Text
'         DC.HIDE_FLAG = Check2Flag(chkHide.Value)
      End If
   End If
   SaveData = True
End Function

'Private Sub chkHide_Click(Value As Integer)
'   m_HasModify = True
'End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
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

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
'      Call LoadExDeliveryCusItem(Nothing, m_ExDeliveryCostItem, , 2)
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         id = 0
         Call QueryData(True)
      End If
      
      m_HasModify = False
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.NAME
      glbErrorLog.ShowUserError
   ElseIf Shift = 1 And KeyCode = 112 Then
      If glbUser.EXCEPTION_FLAG = "Y" Then
         glbUser.EXCEPTION_FLAG = "N"
      Else
         glbUser.EXCEPTION_FLAG = "Y"
      End If
   ElseIf Shift = 0 And KeyCode = 116 Then
'      Call cmdSearch_Click
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
   ElseIf Shift = 0 And KeyCode = 118 Then
'      Call cmdAdd_Click
'   ElseIf Shift = 0 And KeyCode = 117 Then
'      Call cmdDelete_Click
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
   ElseIf Shift = 0 And KeyCode = 114 Then
'      Call cmdEdit_Click
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
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

Private Sub txtCode_Change()
   m_HasModify = True
End Sub

Private Sub txtName_Change()
   m_HasModify = True
End Sub
