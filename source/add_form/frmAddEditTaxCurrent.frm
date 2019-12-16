VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditTaxCurrent 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   Icon            =   "frmAddEditTaxCurrent.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5430
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3255
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   5741
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txttax3 
         Height          =   435
         Left            =   2280
         TabIndex        =   2
         Top             =   1800
         Width           =   1755
         _ExtentX        =   6800
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTax1 
         Height          =   435
         Left            =   2280
         TabIndex        =   0
         Top             =   840
         Width           =   1755
         _ExtentX        =   6800
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txttax2 
         Height          =   435
         Left            =   2280
         TabIndex        =   1
         Top             =   1320
         Width           =   1755
         _ExtentX        =   6800
         _ExtentY        =   767
      End
      Begin VB.Label lblCurrent2 
         Alignment       =   1  'Right Justify
         Caption         =   "lblCurrent2"
         Height          =   315
         Left            =   360
         TabIndex        =   9
         Top             =   1440
         Width           =   1875
      End
      Begin VB.Label lblCurrent3 
         Alignment       =   1  'Right Justify
         Caption         =   "lblCurrent3"
         Height          =   315
         Left            =   600
         TabIndex        =   8
         Top             =   1920
         Width           =   1635
      End
      Begin VB.Label lblCurrent1 
         Alignment       =   1  'Right Justify
         Caption         =   "lblCurrent1"
         Height          =   315
         Left            =   360
         TabIndex        =   7
         Top             =   960
         Width           =   1875
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   2880
         TabIndex        =   4
         Top             =   2400
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   1200
         TabIndex        =   3
         Top             =   2400
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditTaxCurrent.frx":27A2
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditTaxCurrent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean


Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Public tax1 As Double
Public tax2 As Double
Public tax3 As Double


Private Sub cmdOK_Click()
 Call QueryData(True)
   
   OKClick = True
   Unload Me
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean

   If Not VerifyTextControl(lblCurrent1, txtTax1, True) Then
      Exit Sub
   End If
If Not VerifyTextControl(lblCurrent2, txttax2, True) Then
      Exit Sub
   End If
If Not VerifyTextControl(lblCurrent2, txttax2, True) Then
      Exit Sub
   End If
   Call EnableForm(Me, False)
   tax1 = Val(txtTax1.Text)
   tax2 = Val(txttax2.Text)
   tax3 = Val(txttax3.Text)

   Call EnableForm(Me, True)
   
   End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
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
   
   Call InitNormalLabel(lblCurrent1, MapText("TAX1"))
   Call InitNormalLabel(lblCurrent2, MapText("TAX2"))
   Call InitNormalLabel(lblCurrent3, MapText("TAX3"))
   
   Call txtTax1.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txttax2.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txttax3.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
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
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

