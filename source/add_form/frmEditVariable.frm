VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmEditVariable 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   Icon            =   "frmEditVariable.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   6840
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   2955
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   5212
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextBox txtLotNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   1020
         Width           =   3285
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   5
         Top             =   0
         Width           =   6825
         _ExtentX        =   12039
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtItemAmount 
         Height          =   435
         Left            =   1860
         TabIndex        =   1
         Top             =   1470
         Width           =   1575
         _ExtentX        =   13361
         _ExtentY        =   767
      End
      Begin VB.Label lblLotNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   7
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblItemAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   6
         Top             =   1560
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   3450
         TabIndex        =   3
         Top             =   2130
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   1800
         TabIndex        =   2
         Top             =   2130
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmEditVariable.frx":27A2
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmEditVariable"
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
Public HeaderText As String
Public TempCollection As Collection

Private Sub cmdPasswd_Click()

End Sub


Private Sub cboUserGroup_Click()
   m_HasModify = True
End Sub

Private Sub chkEnable_Click(Value As Integer)
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
Dim Fv As CFormulaVariable

   If Flag Then
      Set Fv = TempCollection.Item(id)
      
      txtLotNo.Text = VariableToText(Fv.VARIABLE_ID)
      txtItemAmount.Text = Fv.VARIABLE_VALUE
   End If
   
   If ItemCount > 0 Then

   End If
   
'   If Not IsOK Then
'      glbErrorLog.ShowUserError
'      Call EnableForm(Me, True)
'      Exit Sub
'   End If
   
   Call EnableForm(Me, True)
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
      
   If Not VerifyTextControl(lblItemAmount, txtItemAmount, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
      
   Dim Fv As CFormulaVariable
   If ShowMode = SHOW_ADD Then
      Set Fv = New CFormulaVariable

      Fv.Flag = "A"
      Call TempCollection.add(Fv)
   Else
      Set Fv = TempCollection.Item(id)
      If Fv.Flag <> "A" Then
         Fv.Flag = "E"
      End If
   End If
   
   Fv.AddEditMode = ShowMode
   Fv.VARIABLE_VALUE = Val(txtItemAmount.Text)
   
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
   
   Call InitNormalLabel(lblLotNo, MapText("�����"))
   Call InitNormalLabel(lblItemAmount, MapText("�ӹǹ"))
   
   Call txtLotNo.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   txtLotNo.Enabled = False
   Call txtItemAmount.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
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
   
'   Set FV = New CFV
   Set m_Rs = New ADODB.Recordset
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub txtLeftAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtItemAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtLotNo_Change()
   m_HasModify = True
End Sub

Private Sub txtMaxUser_Change()
   m_HasModify = True
End Sub
