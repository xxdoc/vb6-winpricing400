VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditPlanningItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3840
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
   Icon            =   "frmAddEditPlanningItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
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
      Height          =   6225
      Left            =   0
      TabIndex        =   7
      Top             =   600
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   10980
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextLookup uctlPartLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   0
         Top             =   720
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPlanAmount 
         Height          =   435
         Left            =   1800
         TabIndex        =   1
         Top             =   1170
         Width           =   1485
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtNote 
         Height          =   435
         Left            =   1800
         TabIndex        =   2
         Top             =   1680
         Width           =   5325
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   1440
         TabIndex        =   3
         Top             =   2460
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblNote 
         Alignment       =   1  'Right Justify
         Caption         =   "lblNote"
         Height          =   345
         Left            =   30
         TabIndex        =   10
         Top             =   1710
         Width           =   1695
      End
      Begin VB.Label lblPlanAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPlanAmount"
         Height          =   375
         Left            =   480
         TabIndex        =   9
         Top             =   1230
         Width           =   1245
      End
      Begin VB.Label lblPart 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPart"
         Height          =   315
         Left            =   270
         TabIndex        =   8
         Top             =   720
         Width           =   1455
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3090
         TabIndex        =   4
         Top             =   2460
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPlanningItem.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4740
         TabIndex        =   5
         Top             =   2460
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditPlanningItem"
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

Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public TempCollection As Collection

Public ParentForm As Form
Public ParentTag As String

Private m_PartItems As Collection
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
      
   Call InitNormalLabel(lblPart, MapText("สินค้า/วัตถุดิบ"))
   Call InitNormalLabel(lblPlanAmount, MapText("ยอดประมาณ"))
   Call InitNormalLabel(lblNote, MapText("หมายเหตุ"))
   
   Call txtPlanAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   
   Call uctlPartLookup.MyTextBox.SetKeySearch("PART_NO")
   
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
Dim Pni As CPlanningItem

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Set Pni = TempCollection.Item(ID)
        uctlPartLookup.MyCombo.ListIndex = IDToListIndex(uctlPartLookup.MyCombo, Pni.PART_ITEM_ID)
        txtPlanAmount.Text = Pni.PLAN_AMOUNT
        txtNote.Text = Pni.NOTE
      End If
   End If
   
   Call EnableForm(Me, True)
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

         Call ParentForm.RefreshGrid(ParentTag)
         Exit Sub
      End If
      ID = NewID
      Call QueryData(True)
   ElseIf ShowMode = SHOW_ADD Then
      uctlPartLookup.MyCombo.ListIndex = -1
      txtPlanAmount.Text = ""
      txtNote.Text = ""
   End If
   
   Call ParentForm.RefreshGrid(ParentTag)
   
   Call uctlPartLookup.SetFocus
   m_HasModify = False
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
   
   If Not VerifyCombo(lblPart, uctlPartLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblPlanAmount, txtPlanAmount, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim Pni As CPlanningItem
   If ShowMode = SHOW_ADD Then
      Set Pni = New CPlanningItem
   Else
      Set Pni = TempCollection.Item(ID)
   End If
   
   Pni.PART_ITEM_ID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
   Pni.PART_DESC = uctlPartLookup.MyCombo.Text
   Pni.PART_NO = uctlPartLookup.MyTextBox.Text
   Pni.PLAN_AMOUNT = txtPlanAmount.Text
   Pni.NOTE = txtNote.Text
   If ShowMode = SHOW_ADD Then
      Pni.Flag = "A"
      Call TempCollection.add(Pni)
   Else
      If Pni.Flag <> "A" Then
         Pni.Flag = "E"
      End If
   End If
   
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadPartItem(uctlPartLookup.MyCombo, m_PartItems, , "N", , , "N")
      Set uctlPartLookup.MyCollection = m_PartItems
      
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
   
   Set m_PartItems = New Collection
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_PartItems = Nothing
End Sub

Private Sub txtNote_Change()
   m_HasModify = True
End Sub

Private Sub txtPlanAmount_Change()
   m_HasModify = True
End Sub

Private Sub uctlPartLookup_Change()
   m_HasModify = True
End Sub
