VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAdjustInventoryWH 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7995
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAdjustInventoryWH.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   7995
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2835
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   5001
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextLookup uctlProductLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   1
         Top             =   750
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPartTypeLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   0
         Top             =   300
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin Threed.SSCommand cmdStartAdjust 
         Height          =   525
         Left            =   5040
         TabIndex        =   8
         Top             =   1320
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAdjustInventoryWH.frx":08CA
         ButtonStyle     =   3
      End
      Begin VB.Label lblPartTypeLookup 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPartTypeLookup"
         Height          =   315
         Left            =   270
         TabIndex        =   7
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label lblProductLookup 
         Alignment       =   1  'Right Justify
         Caption         =   "lblProductLookup"
         Height          =   315
         Left            =   270
         TabIndex        =   6
         Top             =   780
         Width           =   1455
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2280
         TabIndex        =   2
         Top             =   2040
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAdjustInventoryWH.frx":0BE4
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4080
         TabIndex        =   3
         Top             =   2040
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAdjustInventoryWH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Header As String
Public ShowMode As SHOW_MODE_TYPE
Private m_PartTypes As Collection
Private m_PartItems As Collection
Private m_HasModify As Boolean

Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public PartNo As String
Public PartType As String
Public DocumentType As Long

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
      
   Call InitNormalLabel(lblPartTypeLookup, MapText("ประเภทสินค้า"))
   Call InitNormalLabel(lblProductLookup, MapText("ชื่อสินค้า"))
   
   cmdStartAdjust.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdStartAdjust, MapText("เริ่มคำนวณยอด"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
End Sub

Private Sub cmdOK_Click()
   PartType = uctlPartTypeLookup.MyTextBox.Text
   PartNo = uctlProductLookup.MyTextBox.Text
   OKClick = True
   Unload Me
End Sub
Private Sub cmdStartAdjust_Click()
   If Not VerifyCombo(lblPartTypeLookup, uctlPartTypeLookup.MyCombo, False) Then
      Exit Sub
   End If
   
   If CalAdjustByPartItem(uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex)), 2, 1, 1, "I", True) Then
      glbErrorLog.LocalErrorMsg = "คำนวณยอดเสร็จสิ้น"
      glbErrorLog.ShowUserError
   Else
      glbErrorLog.LocalErrorMsg = "คำนวณยอดไม่สำเร็จ"
      glbErrorLog.ShowUserError
   End If
End Sub

Private Sub Form_Activate()
   Call LoadPartType(uctlPartTypeLookup.MyCombo, m_PartTypes)
   Set uctlPartTypeLookup.MyCollection = m_PartTypes
         
   If DocumentType = 14 Then
         uctlPartTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPartTypeLookup.MyCombo, 10)
         uctlProductLookup.MyTextBox.Text = PartNo
   ElseIf DocumentType = 13 Then
         uctlPartTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPartTypeLookup.MyCombo, 21)
         uctlProductLookup.MyTextBox.Text = PartNo
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

   Set m_PartTypes = New Collection
   Set m_PartItems = New Collection
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_PartTypes = Nothing
   Set m_PartItems = Nothing
End Sub

Private Sub uctlPartTypeLookup_Change()
Dim PartTypeID As Long
Dim Pt As CPartType
m_HasModify = True

   PartTypeID = uctlPartTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPartTypeLookup.MyCombo.ListIndex))
   
   If PartTypeID > 0 Then
      Set Pt = GetPartType(m_PartTypes, Trim(str(PartTypeID)))
      Call LoadPartItem(uctlProductLookup.MyCombo, m_PartItems, PartTypeID, "N")
      Set uctlProductLookup.MyCollection = m_PartItems
   End If
End Sub

Private Sub uctlProductLookup_Change()
   m_HasModify = True
End Sub
