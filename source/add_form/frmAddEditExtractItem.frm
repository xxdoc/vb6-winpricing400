VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditExtractItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditExtractItem.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2625
      Left            =   0
      TabIndex        =   7
      Top             =   600
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   4630
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextBox txtTotalAmount 
         Height          =   435
         Left            =   1800
         TabIndex        =   2
         Top             =   1140
         Width           =   1485
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPartDesc 
         Height          =   435
         Left            =   1800
         TabIndex        =   1
         Top             =   690
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPartNo 
         Height          =   435
         Left            =   1800
         TabIndex        =   0
         Top             =   240
         Width           =   2355
         _ExtentX        =   10398
         _ExtentY        =   767
      End
      Begin VB.Label lblPartNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblType"
         Height          =   315
         Left            =   270
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   1545
         TabIndex        =   3
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditExtractItem.frx":08CA
         ButtonStyle     =   3
      End
      Begin VB.Label lblTotalAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblAmount"
         Height          =   375
         Left            =   480
         TabIndex        =   9
         Top             =   1200
         Width           =   1245
      End
      Begin VB.Label lblPartDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "lblType"
         Height          =   315
         Left            =   270
         TabIndex        =   8
         Top             =   690
         Width           =   1455
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3195
         TabIndex        =   4
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditExtractItem.frx":0BE4
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4845
         TabIndex        =   5
         Top             =   1800
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditExtractItem"
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
Private m_Input_combo As Collection
Private m_Input1_combo As Collection
Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public TempCollection As Collection
Public COMMIT_FLAG As String
Public ExtractItem As CExtractItem
Public ParentForm As Form

Private m_PartTypes As Collection
Private m_PartItems As Collection
Private m_Locations As Collection
Private m_Formulas As Collection

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
      
   Call InitNormalLabel(lblPartNo, MapText("รหัสวัตถุดิบ"))
   Call InitNormalLabel(lblPartDesc, MapText("ชื่อวัตถุดิบ"))
   Call InitNormalLabel(lblTotalAmount, MapText("จำนวนรวม"))
   
   Call txtPartNo.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtPartNo.Enabled = False
   Call txtPartDesc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtPartDesc.Enabled = False
   Call txtTotalAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdNext, MapText("ถัดไป"))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Dim Ma As CExtractItem
         Set Ma = TempCollection.Item(ID)
         
         txtPartNo.Text = Ma.PART_NO
         txtPartDesc.Text = Ma.PART_DESC
         txtTotalAmount.Text = Ma.TOTAL_AMT
      End If
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Function GetNextID(OldID As Long) As Long
Dim Ei As CExtractItem
Dim TempIndex As Long
Dim J As Long

'   TempIndex = 0
'   For Each Ei In TempCollection
'      TempIndex = TempIndex + 1
'      If ID = OldID Then
'         j = TempIndex
'         Exit For
'      End If
'   Next Ei
   
   If OldID >= TempCollection.Count Then
      J = TempCollection.Count
      glbErrorLog.LocalErrorMsg = "ถึงเรคคอร์ดสุดท้ายแล้ว"
      glbErrorLog.ShowUserError
   Else
      J = OldID + 1
   End If
   
   GetNextID = J
End Function

Private Sub cmdNext_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   ID = GetNextID(ID)
   Call ParentForm.ShowGrid
   
   Call QueryData(True)
   
   txtTotalAmount.SetFocus
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

   If Not VerifyTextControl(lblTotalAmount, txtTotalAmount, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim Ma As CExtractItem
   If ShowMode = SHOW_ADD Then
      Set Ma = New CExtractItem
   Else
      Set Ma = TempCollection.Item(ID)
   End If
   
   Ma.TOTAL_AMT = Val(txtTotalAmount.Text)
   
   If ShowMode = SHOW_ADD Then
      Ma.Flag = "A"
      Call TempCollection.add(Ma)
   Else
      If Ma.Flag <> "A" Then
         Ma.Flag = "E"
      End If
   End If
   
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
   Set m_Input_combo = New Collection
   Set m_Input1_combo = New Collection
   Set m_Rs = New ADODB.Recordset
   
   Set m_PartTypes = New Collection
   Set m_PartItems = New Collection
   Set m_Locations = New Collection
   Set m_Formulas = New Collection
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_PartTypes = Nothing
   Set m_PartItems = Nothing
   Set m_Locations = Nothing
   Set m_Formulas = Nothing
End Sub

Private Sub txtAvgPrice_Change()
   m_HasModify = True
End Sub

Private Sub txtGroupNo_Change()
   m_HasModify = True
End Sub

Private Sub txtStdAmount_Change()
   m_HasModify = True
End Sub

Private Sub uctlFromFormulaLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlMixDate_HasChange()
   m_HasModify = True
End Sub

Private Sub txtAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtLink_Change()
   m_HasModify = True
End Sub

Private Sub txtRef_Change()
   m_HasModify = True
End Sub

Private Sub txtSerialNo_Change()
   m_HasModify = True
End Sub

Private Sub uctlPlaceLookup_Change()
   m_HasModify = True
End Sub

Private Sub txtPartDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtTotalAmount_Change()
   m_HasModify = True
End Sub
