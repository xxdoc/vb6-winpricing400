VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditStepTier 
   ClientHeight    =   2895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9885
   Icon            =   "frmAddEditStepTier.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   9885
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   2925
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   5159
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel pnlFooter 
         Height          =   675
         Left            =   30
         TabIndex        =   6
         Top             =   2220
         Width           =   9870
         _ExtentX        =   17410
         _ExtentY        =   1191
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin Threed.SSCommand cmdExit 
            Cancel          =   -1  'True
            Height          =   525
            Left            =   4965
            TabIndex        =   14
            Top             =   90
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   926
            _Version        =   131073
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdOK 
            Height          =   525
            Left            =   3315
            TabIndex        =   13
            Top             =   90
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   926
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmAddEditStepTier.frx":27A2
            ButtonStyle     =   3
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   675
         Left            =   30
         TabIndex        =   5
         Top             =   30
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   1191
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtRate 
         Height          =   435
         Left            =   1710
         TabIndex        =   3
         Top             =   1500
         Width           =   1695
         _ExtentX        =   6165
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtFrom 
         Height          =   435
         Left            =   1710
         TabIndex        =   0
         Top             =   1050
         Width           =   1695
         _ExtentX        =   6165
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtWidth 
         Height          =   435
         Left            =   7620
         TabIndex        =   2
         Top             =   1020
         Width           =   1695
         _ExtentX        =   6165
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTo 
         Height          =   435
         Left            =   4590
         TabIndex        =   1
         Top             =   1020
         Width           =   1695
         _ExtentX        =   6165
         _ExtentY        =   767
      End
      Begin VB.Label lblTo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3480
         TabIndex        =   12
         Top             =   1140
         Width           =   1005
      End
      Begin VB.Label lblWidth 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6330
         TabIndex        =   11
         Top             =   1140
         Width           =   1185
      End
      Begin VB.Label lblFrom 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   150
         TabIndex        =   10
         Top             =   1170
         Width           =   1485
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   0
         TabIndex        =   8
         Top             =   480
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label lblRate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   150
         TabIndex        =   7
         Top             =   1620
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmAddEditStepTier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Features As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public id As Long
Public TempCollection As Collection
Public SocCode As String
Public FROM_QUANTITY As Double
Public RATE_TYPE As Long
Public SocPartType As Long

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim D As CStpTierVol

   If Flag Then
      Call EnableForm(Me, False)
      
      Set D = TempCollection.Item(id)
      
      txtFrom.Text = D.FROM_QUANTITY
      txtTo.Text = D.TO_QUANTITY
      txtWidth.Text = D.TO_QUANTITY - D.FROM_QUANTITY
      txtRate.Text = D.RATE_AMOUNT
      
      Call EnableForm(Me, True)
   End If
   
   If ItemCount > 0 Then
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cboFeatureLevel_Click()
   m_HasModify = True
End Sub

Private Sub cboGroup_Click()
   m_HasModify = True
End Sub

Private Sub chkEnable_Click()
   m_HasModify = True
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim Sf As CStpTierVol

   If Not VerifyTextControl(lblWidth, txtWidth, True) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblRate, txtRate, True) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   If ShowMode = SHOW_ADD Then
      Set Sf = New CStpTierVol

      Sf.Flag = "A"
      Call TempCollection.add(Sf)
   Else
      Set Sf = TempCollection(id)
      Sf.Flag = "E"
   End If
   
   Sf.FROM_QUANTITY = Val(txtFrom.Text)
   Sf.TO_QUANTITY = Val(txtTo.Text)
   Sf.RATE_AMOUNT = Val(txtRate.Text)
   Sf.Width = Val(txtWidth.Text)
   
   Call EnableForm(Me, True)
   SaveData = True
End Function

Private Sub cboRateType_Click()
   m_HasModify = True
End Sub

Private Sub cboUnit_Click()
   m_HasModify = True
End Sub

Private Sub chkStartEndFlag_Click()
   m_HasModify = True
End Sub

Private Sub cmdOK_Click()

   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Sub Form_Activate()
Dim Sp As CSystemParam

   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         Call QueryData(True)
      Else
         txtFrom.Text = FROM_QUANTITY
         txtTo.Text = FROM_QUANTITY
         txtWidth.Text = 0
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

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_Features = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub InitFormLayout()
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   Me.KeyPreview = True
   
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   pnlFooter.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)

   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
    If SocPartType = 1 Then
   Call InitNormalLabel(lblRate, MapText("อัตราค่าบริการ"))
   Else
   Call InitNormalLabel(lblRate, MapText("อัตราค่าสินค้า"))
   End If
   Call InitNormalLabel(lblFrom, MapText("จาก"))
   Call InitNormalLabel(lblTo, MapText("ถึง"))
   Call InitNormalLabel(lblWidth, MapText("ความกว้าง"))
   
   Call txtFrom.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtTo.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtWidth.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtRate.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   
   txtFrom.Enabled = False
   txtTo.Enabled = False
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   OKClick = False
   Call InitFormLayout
      
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_Features = New Collection
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub


Private Sub txtFeatureCode_Change()
   m_HasModify = True
End Sub

Private Sub txtPassword_Change()
   m_HasModify = True
End Sub

Private Sub txtAC_Change()
   m_HasModify = True
End Sub

Private Sub txtFix_Change()
   m_HasModify = True
End Sub

Private Sub txtPercent_Change()
   m_HasModify = True
End Sub

Private Sub txtQuoata_Change()
   m_HasModify = True
End Sub

Private Sub txtOC_Change()
   m_HasModify = True
End Sub

Private Sub txtRate_Change()
   m_HasModify = True
End Sub

Private Sub txtRC_Change()
   m_HasModify = True
End Sub

Private Sub txtRoundingFactor_Change()
   m_HasModify = True
End Sub

Private Sub txtSocCode_Change()
   m_HasModify = True
End Sub

Private Sub uctlEffectiveDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlEffectiveTIme_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlExpireDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlTime1_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlTime2_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlExpireTime_HasChange()
   m_HasModify = True
End Sub

Private Sub txtWidth_Change()
   txtTo.Text = Val(txtFrom.Text) + Val(txtWidth.Text)
   m_HasModify = True
End Sub
