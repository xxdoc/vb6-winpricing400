VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditEvaluatePay 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Icon            =   "frmAddEditEvaluatePay.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   10410
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3765
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   6641
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextBox txtEvaluateDesc 
         Height          =   435
         Left            =   1920
         TabIndex        =   2
         Top             =   1800
         Width           =   8205
         _ExtentX        =   14473
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   7
         Top             =   0
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlDate uctlEvaluateDate 
         Height          =   405
         Left            =   1920
         TabIndex        =   0
         Top             =   840
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextLookup uctlSupplierLookup 
         Height          =   435
         Left            =   1920
         TabIndex        =   1
         Top             =   1320
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtEvaluateAmount 
         Height          =   435
         Left            =   1920
         TabIndex        =   3
         Top             =   2280
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   767
      End
      Begin VB.Label lblEvaluateAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   11
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label lblSupplierNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   360
         TabIndex        =   10
         Top             =   1380
         Width           =   1455
      End
      Begin VB.Label lblEvaluateDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   9
         Top             =   990
         Width           =   1575
      End
      Begin VB.Label lblEvaluateDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   8
         Top             =   1860
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5235
         TabIndex        =   5
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
         TabIndex        =   4
         Top             =   2910
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditEvaluatePay.frx":27A2
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditEvaluatePay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_EvaluatePay As CEvaluatePay

Private m_Suppliers As Collection

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
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
Dim itemcount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      Call m_EvaluatePay.SetFieldValue("EVALUATE_PAY_ID", ID)
      m_EvaluatePay.QueryFlag = 1
      If Not glbDaily.QueryEvaluatePay(m_EvaluatePay, m_Rs, itemcount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If itemcount > 0 Then
      Call m_EvaluatePay.PopulateFromRS(1, m_Rs)
      
      UctlEvaluateDate.ShowDate = m_EvaluatePay.GetFieldValue("EVALUATE_DATE")
      uctlSupplierLookup.MyCombo.ListIndex = IDToListIndex(uctlSupplierLookup.MyCombo, m_EvaluatePay.GetFieldValue("SUPPLIER_ID"))
      txtEvaluateDesc.Text = m_EvaluatePay.GetFieldValue("EVALUATE_PAY_DESC")
      txtEvaluateAmount.Text = m_EvaluatePay.GetFieldValue("EVALUATE_AMOUNT")
      
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
   
   If Not VerifyDate(lblEvaluateDate, UctlEvaluateDate, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblEvaluateAmount, txtEvaluateAmount, True) Then
      Exit Function
   End If
   
'   If Not CheckUniqueNs(EMPCODE_UNIQUE, txtName.Text, ID) Then
'      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtName.Text & " " & MapText("อยู่ในระบบแล้ว")
'      glbErrorLog.ShowUserError
'      Exit Function
'   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Call m_EvaluatePay.SetFieldValue("EVALUATE_PAY_ID", ID)
   m_EvaluatePay.ShowMode = ShowMode
   Call m_EvaluatePay.SetFieldValue("EVALUATE_DATE", UctlEvaluateDate.ShowDate)
   Call m_EvaluatePay.SetFieldValue("SUPPLIER_ID", uctlSupplierLookup.MyCombo.ItemData(Minus2Zero(uctlSupplierLookup.MyCombo.ListIndex)))
   Call m_EvaluatePay.SetFieldValue("EVALUATE_AMOUNT", Val(txtEvaluateAmount.Text))
   Call m_EvaluatePay.SetFieldValue("EVALUATE_PAY_DESC", txtEvaluateDesc.Text)
      
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditEvaluatePay(m_EvaluatePay, IsOK, True, glbErrorLog) Then
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
      
      Call LoadSupplier(uctlSupplierLookup.MyCombo, m_Suppliers)
      Set uctlSupplierLookup.MyCollection = m_Suppliers
      
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

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblSupplierNo, MapText("ซัพพลายเออร์"))
   Call InitNormalLabel(lblEvaluateDate, MapText("วันที่"))
   Call InitNormalLabel(lblEvaluateAmount, MapText("จำนวน"))
   Call InitNormalLabel(lblEvaluateDesc, MapText("รายละเอียด"))
   
   
   Call txtEvaluateAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
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
   
   Set m_EvaluatePay = New CEvaluatePay
   Set m_Rs = New ADODB.Recordset
   Set m_Suppliers = New Collection
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set m_Suppliers = Nothing
End Sub

Private Sub txtEvaluateAmount_Change()
m_HasModify = True
End Sub

Private Sub txtEvaluateDesc_Change()
m_HasModify = True
End Sub

Private Sub uctlEvaluateDate_HasChange()
m_HasModify = True
End Sub

Private Sub uctlSupplierLookup_Change()
m_HasModify = True
End Sub
