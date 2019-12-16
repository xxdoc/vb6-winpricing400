VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditSliptSalarySub 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   Icon            =   "frmAddEditSliptSalarySub.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   6150
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3765
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   6641
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboType 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1200
         Width           =   2955
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   6195
         _ExtentX        =   10927
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtAmount 
         Height          =   435
         Left            =   1860
         TabIndex        =   1
         Top             =   1680
         Width           =   2925
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin VB.Label lblAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblAmount"
         Height          =   315
         Left            =   210
         TabIndex        =   8
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label lblType 
         Alignment       =   1  'Right Justify
         Caption         =   "lblType"
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblBath 
         Alignment       =   1  'Right Justify
         Caption         =   "lblBath"
         Height          =   315
         Left            =   4800
         TabIndex        =   6
         Top             =   1800
         Width           =   495
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   3480
         TabIndex        =   3
         Top             =   2520
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   1200
         TabIndex        =   2
         Top             =   2520
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditSliptSalarySub.frx":27A2
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditSliptSalarySub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Public id As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public ParentShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Public TempCollection As Collection

Private Sub cboType_Change()
   m_HasModify = True
End Sub
Private Sub cboType_Click()
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
      
      If ShowMode = SHOW_EDIT Then
           Dim SB As CSliptSub
         Set SB = TempCollection.Item(id)
         
         txtAmount.Text = SB.MONTHLY_AMOUNT
         cboType.ListIndex = IDToListIndex(cboType, SB.MONTHLY_SUB)
      End If
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
   
   If Not VerifyTextControl(lblAmount, txtAmount, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblType, cboType, False) Then
      Exit Function
   End If
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
        
         Dim SB As CSliptSub
   If ShowMode = SHOW_ADD Then
      Set SB = New CSliptSub
         Else
      Set SB = TempCollection.Item(id)
   End If
   
   SB.MONTHLY_AMOUNT = txtAmount.Text
   SB.MONTHLY_NAME = cboType.Text
   SB.MONTHLY_SUB = cboType.ItemData(Minus2Zero(cboType.ListIndex))
   If ShowMode = SHOW_ADD Then
      SB.Flag = "A"
      Call TempCollection.add(SB)
      Else
      If SB.Flag <> "A" Then
      SB.Flag = "E"
      End If
         End If
   
   SaveData = True

End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadSliptSub(cboType)
      
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
   
   Call InitNormalLabel(lblType, MapText("ประเภทเงินหัก"))
   Call InitNormalLabel(lblAmount, MapText("จำนวนเงิน"))
   Call InitNormalLabel(lblBath, MapText("บาท"))
   
   Call txtAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.CODE_TYPE)
  
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitCombo(cboType)
   
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
   
   Set m_Rs = New ADODB.Recordset
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''''''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub txtAmount_Change()
   m_HasModify = True
End Sub

