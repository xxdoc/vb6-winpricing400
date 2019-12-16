VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditExpenseDetail 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11520
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditExpenseDetail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   11520
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   3135
      Left            =   0
      TabIndex        =   9
      Top             =   600
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   5530
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextLookup uctlExpenseTypeLookup 
         Height          =   405
         Left            =   1860
         TabIndex        =   0
         Top             =   300
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtDesc 
         Height          =   495
         Left            =   1860
         TabIndex        =   1
         Top             =   780
         Width           =   9375
         _ExtentX        =   3413
         _ExtentY        =   873
      End
      Begin prjFarmManagement.uctlTextBox txtAmount 
         Height          =   495
         Left            =   1860
         TabIndex        =   2
         Top             =   1440
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
      End
      Begin prjFarmManagement.uctlTextBox txtAvg 
         Height          =   495
         Left            =   5460
         TabIndex        =   3
         Top             =   1440
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
      End
      Begin prjFarmManagement.uctlTextBox txtPrice 
         Height          =   495
         Left            =   9180
         TabIndex        =   4
         Top             =   1440
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   3480
         TabIndex        =   5
         Top             =   2190
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditExpenseDetail.frx":08CA
         ButtonStyle     =   3
      End
      Begin VB.Label lblPrice 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPrice"
         Height          =   375
         Left            =   7440
         TabIndex        =   14
         Top             =   1500
         Width           =   1665
      End
      Begin VB.Label lblAvg 
         Alignment       =   1  'Right Justify
         Caption         =   "lblAvg"
         Height          =   375
         Left            =   3720
         TabIndex        =   13
         Top             =   1500
         Width           =   1665
      End
      Begin VB.Label lblAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblAmount"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   1665
      End
      Begin VB.Label lblExpenseTypeLookup 
         Alignment       =   1  'Right Justify
         Caption         =   "lblExpenseTypeLookup"
         Height          =   315
         Left            =   30
         TabIndex        =   11
         Top             =   330
         Width           =   1725
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "lblDesc"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1665
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   5130
         TabIndex        =   6
         Top             =   2190
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditExpenseDetail.frx":0BE4
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   6780
         TabIndex        =   7
         Top             =   2190
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditExpenseDetail"
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

Private m_ProcessParams As Collection
Public ParentForm As Form

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
      
   Call InitNormalLabel(lblExpenseTypeLookup, MapText("ต้นทุนผลิต"))
   Call InitNormalLabel(lblDesc, MapText("รายละเอียด"))
   Call InitNormalLabel(lblAmount, MapText("จำนวน"))
   Call InitNormalLabel(lblAvg, MapText("ราคาเฉลี่ย"))
   Call InitNormalLabel(lblPrice, MapText("มูลค่า"))
   
   Call txtAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtAvg.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtPrice.Enabled = False
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdNext, MapText("ถัดไป"))
End Sub
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
         Dim MA As CExpenseDetail
         Set MA = TempCollection.Item(ID)
         
      If ShowMode = SHOW_EDIT Then
         uctlExpenseTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlExpenseTypeLookup.MyCombo, MA.GetFieldValue("EXPENSE_DETAIL_TYPE"))
         txtDesc.Text = MA.GetFieldValue("EXPENSE_DETAIL_DESC")
         txtAmount.Text = MA.GetFieldValue("EXPENSE_DETAIL_AMOUNT")
         txtAvg.Text = MA.GetFieldValue("EXPENSE_DETAIL_AVG")
         txtPrice.Text = MA.GetFieldValue("EXPENSE_DETAIL_PRICE")
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
         
   If Not VerifyCombo(lblExpenseTypeLookup, uctlExpenseTypeLookup.MyCombo, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim MA As CExpenseDetail
   If ShowMode = SHOW_ADD Then
      Set MA = New CExpenseDetail
   Else
      Set MA = TempCollection.Item(ID)
   End If
   
   If ShowMode = SHOW_ADD Then
      MA.Flag = "A"
      Call TempCollection.add(MA)
   Else
      If MA.Flag <> "A" Then
         MA.Flag = "E"
      End If
   End If
   
   Call MA.SetFieldValue("EXPENSE_DETAIL_TYPE", uctlExpenseTypeLookup.MyCombo.ItemData(Minus2Zero(uctlExpenseTypeLookup.MyCombo.ListIndex)))
   Call MA.SetFieldValue("EXPENSE_DETAIL_DESC", txtDesc.Text)
   Call MA.SetFieldValue("EXPENSE_DETAIL_AMOUNT", Val(txtAmount.Text))
   Call MA.SetFieldValue("EXPENSE_DETAIL_AVG", Val(txtAvg.Text))
   Call MA.SetFieldValue("EXPENSE_DETAIL_PRICE", Val(txtPrice.Text))
   
   Call MA.SetFieldValue("PARAMETER_PROCESS_NAME", uctlExpenseTypeLookup.MyCombo.Text)
   
   SaveData = True
End Function
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadParameterProcess(uctlExpenseTypeLookup.MyCombo, m_ProcessParams)
      Set uctlExpenseTypeLookup.MyCollection = m_ProcessParams
            
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
   
   Set m_ProcessParams = New Collection
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_ProcessParams = Nothing
End Sub

Private Sub txtAmount_Change()
   txtPrice.Text = Val(txtAmount.Text) * Val(txtAvg.Text)
   m_HasModify = True
End Sub

Private Sub txtAvg_Change()
   txtPrice.Text = Val(txtAmount.Text) * Val(txtAvg.Text)
   m_HasModify = True
End Sub

Private Sub txtDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtPrice_Change()
   m_HasModify = True
End Sub

Private Sub uctlExpenseTypeLookup_Change()
   m_HasModify = True
End Sub
Private Sub cmdNext_Click()
Dim NewID As Long

   If Not SaveData Then
      Exit Sub
   End If
   
   If ShowMode = SHOW_EDIT Then
      NewID = GetNextID(ID, TempCollection)
      If ID = NewID Then
         glbErrorLog.LocalErrorMsg = "ถึงเรคคอร์ดสุดท้ายแล้ว"
         glbErrorLog.ShowUserError
         
         Call ParentForm.RefreshGrid
         uctlExpenseTypeLookup.SetFocus
         Exit Sub
      End If
      
      ID = NewID
      Call QueryData(True)
   ElseIf ShowMode = SHOW_ADD Then
      uctlExpenseTypeLookup.MyCombo.ListIndex = -1
      txtDesc.Text = ""
      txtAmount.Text = ""
      txtAvg.Text = ""
      txtPrice.Text = ""
   End If
   
   uctlExpenseTypeLookup.SetFocus
   Call ParentForm.RefreshGrid
End Sub

