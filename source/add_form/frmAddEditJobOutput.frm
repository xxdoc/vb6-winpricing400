VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditJobOutput 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5040
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
   Icon            =   "frmAddEditJobOutput.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   4485
      Left            =   0
      TabIndex        =   10
      Top             =   600
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   7911
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboType 
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
         TabIndex        =   0
         Top             =   240
         Width           =   2565
      End
      Begin prjFarmManagement.uctlTextLookup uctlProductLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   1
         Top             =   720
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtAmount 
         Height          =   435
         Left            =   1800
         TabIndex        =   2
         Top             =   1200
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPlaceLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   3
         Top             =   1680
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtRef 
         Height          =   435
         Left            =   1800
         TabIndex        =   6
         Top             =   3120
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtSerialNo 
         Height          =   435
         Left            =   1800
         TabIndex        =   5
         Top             =   2640
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtLink 
         Height          =   435
         Left            =   1800
         TabIndex        =   4
         Top             =   2160
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin VB.Label lblLink 
         Alignment       =   1  'Right Justify
         Caption         =   "lblLink"
         Height          =   375
         Left            =   480
         TabIndex        =   17
         Top             =   2280
         Width           =   1245
      End
      Begin VB.Label lblSerialNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblSerialNo"
         Height          =   375
         Left            =   0
         TabIndex        =   16
         Top             =   2760
         Width           =   1725
      End
      Begin VB.Label lblRef 
         Alignment       =   1  'Right Justify
         Caption         =   "lblRef"
         Height          =   375
         Left            =   0
         TabIndex        =   15
         Top             =   3240
         Width           =   1725
      End
      Begin VB.Label lblPlace 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPlace"
         Height          =   315
         Left            =   240
         TabIndex        =   14
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label lblAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblAmount"
         Height          =   375
         Left            =   480
         TabIndex        =   13
         Top             =   1320
         Width           =   1245
      End
      Begin VB.Label lblType 
         Alignment       =   1  'Right Justify
         Caption         =   "lblType"
         Height          =   315
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblProduct 
         Alignment       =   1  'Right Justify
         Caption         =   "lblProduct"
         Height          =   315
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   1455
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   1920
         TabIndex        =   7
         Top             =   3720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJobOutput.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4560
         TabIndex        =   8
         Top             =   3720
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditJobOutput"
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
Public tempCollection As Collection

Private Sub cboType_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
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
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
      
   Call InitNormalLabel(lblType, MapText("ประเภท"))
   Call InitNormalLabel(lblProduct, MapText("ผลผลิต"))
   Call InitNormalLabel(lblAmount, MapText("จำนวน"))
   Call InitNormalLabel(lblPlace, MapText("สถานที่เก็บ"))
   Call InitNormalLabel(lblLink, MapText("รหัสเชื่อมโยง"))
   Call InitNormalLabel(lblSerialNo, MapText("รหัสสินค้าขาย"))
   Call InitNormalLabel(lblRef, MapText("หมายเลขอ้างอิง"))
   
   Call txtAmount.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtLink.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtSerialNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtRef.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   Call InitCombo(cboType)
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
           Dim Ma As CJobInput
         Set Ma = tempCollection.Item(ID)
        cboType.ListIndex = IDToListIndex(cboType, Ma.PART_TYPE_ID)
        uctlProductLookup.MyCombo.ListIndex = IDToListIndex(uctlProductLookup.MyCombo, Ma.PART_ITEM_ID)
        txtAmount.Text = Ma.TX_AMOUNT
        uctlPlaceLookup.MyCombo.ListIndex = IDToListIndex(uctlPlaceLookup.MyCombo, Ma.LOCATION_ID)
        txtLink.Text = Ma.LINK_ID
        txtSerialNo.Text = Ma.SERIAL_NUMBER
        txtRef.Text = Ma.INOUT_REF
        
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

   
   If Not VerifyCombo(lblType, cboType, False) Then
      Exit Function
   End If

   If Not VerifyCombo(lblProduct, uctlProductLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblAmount, txtAmount, False) Then
      Exit Function
   End If
  If Not VerifyCombo(lblPlace, uctlPlaceLookup.MyCombo, False) Then
      Exit Function
   End If
   
   
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   Dim Ma As CJobInput
   If ShowMode = SHOW_ADD Then
      Set Ma = New CJobInput
   Else
      Set Ma = tempCollection.Item(ID)
   End If
   Ma.PART_ITEM_ID = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
   Ma.PART_DESC = uctlProductLookup.MyCombo.Text
   Ma.PART_NO = uctlProductLookup.MyTextBox.Text
   Ma.PART_TYPE_ID = cboType.ItemData(Minus2Zero(cboType.ListIndex))
   Ma.PART_TYPE_NAME = cboType.Text
   Ma.TX_AMOUNT = txtAmount.Text
   Ma.LOCATION_ID = uctlPlaceLookup.MyCombo.ItemData(Minus2Zero(uctlPlaceLookup.MyCombo.ListIndex))
   Ma.LOCATION_NO = uctlPlaceLookup.MyTextBox.Text
   Ma.LOCATION_NAME = uctlPlaceLookup.MyCombo.Text
   Ma.LINK_ID = txtLink.Text
   Ma.SERIAL_NUMBER = txtSerialNo.Text
   Ma.INOUT_REF = txtRef.Text
   Ma.TX_TYPE = "I"
   If ShowMode = SHOW_ADD Then
      Ma.Flag = "A"
      Call tempCollection.add(Ma)
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
      Call LoadPartType(cboType)
      Call LoadPartItem(uctlProductLookup.MyCombo, m_Input_combo)
      Set uctlProductLookup.MyCollection = m_Input_combo
      Call LoadLocation(uctlPlaceLookup.MyCombo, m_Input1_combo, 2)
      Set uctlPlaceLookup.MyCollection = m_Input1_combo
    
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
End Sub
Private Sub cboType_Change()
m_HasModify = True
End Sub

Private Sub cboType_Click()
Dim ID As Long
   ID = cboType.ItemData(Minus2Zero(cboType.ListIndex))
   If ID <> 0 Then
   Call LoadPartItem(uctlProductLookup.MyCombo, m_Input_combo, ID)
   End If
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

Private Sub uctlProductLookup_Change()
   m_HasModify = True
End Sub
