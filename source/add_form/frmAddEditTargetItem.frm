VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditTargetItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15330
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditTargetItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   15330
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   19035
      _ExtentX        =   33576
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   6225
      Left            =   0
      TabIndex        =   17
      Top             =   600
      Width           =   19065
      _ExtentX        =   33629
      _ExtentY        =   10980
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextLookup uctlSaleLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   0
         Top             =   720
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTargetPrice1 
         Height          =   435
         Left            =   1800
         TabIndex        =   1
         Top             =   1650
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTargetPrice2 
         Height          =   435
         Left            =   2880
         TabIndex        =   2
         Top             =   1650
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTargetPrice3 
         Height          =   435
         Left            =   3960
         TabIndex        =   3
         Top             =   1650
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTargetPrice4 
         Height          =   435
         Left            =   5040
         TabIndex        =   4
         Top             =   1650
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTargetPrice5 
         Height          =   435
         Left            =   6120
         TabIndex        =   5
         Top             =   1650
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTargetPrice6 
         Height          =   435
         Left            =   7200
         TabIndex        =   6
         Top             =   1650
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTargetPrice7 
         Height          =   435
         Left            =   8280
         TabIndex        =   7
         Top             =   1650
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTargetPrice8 
         Height          =   435
         Left            =   9360
         TabIndex        =   8
         Top             =   1650
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTargetPrice9 
         Height          =   435
         Left            =   10440
         TabIndex        =   9
         Top             =   1650
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTargetPrice10 
         Height          =   435
         Left            =   11520
         TabIndex        =   10
         Top             =   1650
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTargetPrice11 
         Height          =   435
         Left            =   12600
         TabIndex        =   11
         Top             =   1650
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTargetPrice12 
         Height          =   435
         Left            =   13680
         TabIndex        =   12
         Top             =   1650
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   767
      End
      Begin VB.Label lblMonth12 
         Alignment       =   1  'Right Justify
         Caption         =   "lblMonth1"
         Height          =   375
         Left            =   13920
         TabIndex        =   31
         Top             =   1200
         Width           =   765
      End
      Begin VB.Label lblMonth11 
         Alignment       =   1  'Right Justify
         Caption         =   "lblMonth1"
         Height          =   375
         Left            =   12840
         TabIndex        =   30
         Top             =   1200
         Width           =   765
      End
      Begin VB.Label lblMonth10 
         Alignment       =   1  'Right Justify
         Caption         =   "lblMonth1"
         Height          =   375
         Left            =   11760
         TabIndex        =   29
         Top             =   1200
         Width           =   765
      End
      Begin VB.Label lblMonth9 
         Alignment       =   1  'Right Justify
         Caption         =   "lblMonth1"
         Height          =   375
         Left            =   10680
         TabIndex        =   28
         Top             =   1200
         Width           =   765
      End
      Begin VB.Label lblMonth8 
         Alignment       =   1  'Right Justify
         Caption         =   "lblMonth1"
         Height          =   375
         Left            =   9600
         TabIndex        =   27
         Top             =   1200
         Width           =   765
      End
      Begin VB.Label lblMonth7 
         Alignment       =   1  'Right Justify
         Caption         =   "lblMonth1"
         Height          =   375
         Left            =   8520
         TabIndex        =   26
         Top             =   1200
         Width           =   765
      End
      Begin VB.Label lblMonth6 
         Alignment       =   1  'Right Justify
         Caption         =   "lblMonth1"
         Height          =   375
         Left            =   7440
         TabIndex        =   25
         Top             =   1200
         Width           =   765
      End
      Begin VB.Label lblMonth5 
         Alignment       =   1  'Right Justify
         Caption         =   "lblMonth1"
         Height          =   375
         Left            =   6360
         TabIndex        =   24
         Top             =   1200
         Width           =   765
      End
      Begin VB.Label lblMonth4 
         Alignment       =   1  'Right Justify
         Caption         =   "lblMonth1"
         Height          =   375
         Left            =   5280
         TabIndex        =   23
         Top             =   1200
         Width           =   765
      End
      Begin VB.Label lblMonth3 
         Alignment       =   1  'Right Justify
         Caption         =   "lblMonth1"
         Height          =   375
         Left            =   4200
         TabIndex        =   22
         Top             =   1200
         Width           =   765
      End
      Begin VB.Label lblMonth2 
         Alignment       =   1  'Right Justify
         Caption         =   "lblMonth1"
         Height          =   375
         Left            =   3120
         TabIndex        =   21
         Top             =   1200
         Width           =   765
      End
      Begin VB.Label lblMonth1 
         Alignment       =   1  'Right Justify
         Caption         =   "lblMonth1"
         Height          =   375
         Left            =   2040
         TabIndex        =   20
         Top             =   1200
         Width           =   765
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   5400
         TabIndex        =   13
         Top             =   2940
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblTarget 
         Alignment       =   1  'Right Justify
         Caption         =   "lblTarget"
         Height          =   375
         Left            =   480
         TabIndex        =   19
         Top             =   1800
         Width           =   1245
      End
      Begin VB.Label lblSale 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPart"
         Height          =   315
         Left            =   270
         TabIndex        =   18
         Top             =   720
         Width           =   1455
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   7050
         TabIndex        =   14
         Top             =   2940
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditTargetItem.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   8700
         TabIndex        =   15
         Top             =   2940
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditTargetItem"
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

Private collEmp As Collection

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
      
   Call InitNormalLabel(lblSale, MapText("พนักงานขาย"))
   Call InitNormalLabel(lblMonth1, MapText("1"))
   Call InitNormalLabel(lblMonth2, MapText("2"))
   Call InitNormalLabel(lblMonth3, MapText("3"))
   Call InitNormalLabel(lblMonth4, MapText("4"))
   Call InitNormalLabel(lblMonth5, MapText("5"))
   Call InitNormalLabel(lblMonth6, MapText("6"))
   Call InitNormalLabel(lblMonth7, MapText("7"))
   Call InitNormalLabel(lblMonth8, MapText("8"))
   Call InitNormalLabel(lblMonth9, MapText("9"))
   Call InitNormalLabel(lblMonth10, MapText("10"))
   Call InitNormalLabel(lblMonth11, MapText("11"))
   Call InitNormalLabel(lblMonth12, MapText("12"))
   
   Call InitNormalLabel(lblTarget, MapText("เป้าการขาย(ยอดตัน)"))
   
   Call uctlSaleLookup.MyTextBox.SetKeySearch("EMP_CODE")
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdNext, MapText("ถัดไป"))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long
Dim Tgdt As CTargetDetail

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Set Tgdt = TempCollection.Item(ID)
        uctlSaleLookup.MyCombo.ListIndex = IDToListIndex(uctlSaleLookup.MyCombo, Tgdt.EMP_ID)
        
        txtTargetPrice1.Text = Tgdt.TARGET_PRICE1
        txtTargetPrice2.Text = Tgdt.TARGET_PRICE2
        txtTargetPrice3.Text = Tgdt.TARGET_PRICE3
        txtTargetPrice4.Text = Tgdt.TARGET_PRICE4
        txtTargetPrice5.Text = Tgdt.TARGET_PRICE5
        txtTargetPrice6.Text = Tgdt.TARGET_PRICE6
        txtTargetPrice7.Text = Tgdt.TARGET_PRICE7
        txtTargetPrice8.Text = Tgdt.TARGET_PRICE8
        txtTargetPrice9.Text = Tgdt.TARGET_PRICE9
        txtTargetPrice10.Text = Tgdt.TARGET_PRICE10
        txtTargetPrice11.Text = Tgdt.TARGET_PRICE11
        txtTargetPrice12.Text = Tgdt.TARGET_PRICE12
        
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
      uctlSaleLookup.MyCombo.ListIndex = -1
      txtTargetPrice1.Text = ""
      txtTargetPrice2.Text = ""
      txtTargetPrice3.Text = ""
      txtTargetPrice4.Text = ""
      txtTargetPrice5.Text = ""
      txtTargetPrice6.Text = ""
      txtTargetPrice7.Text = ""
      txtTargetPrice8.Text = ""
      txtTargetPrice9.Text = ""
      txtTargetPrice10.Text = ""
      txtTargetPrice11.Text = ""
      txtTargetPrice12.Text = ""
      
   End If
   
   Call ParentForm.RefreshGrid(ParentTag)
   
   Call uctlSaleLookup.SetFocus
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
   
   If Not VerifyCombo(lblSale, uctlSaleLookup.MyCombo, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim Tgdt As CTargetDetail
   If ShowMode = SHOW_ADD Then
      Set Tgdt = New CTargetDetail
   Else
      Set Tgdt = TempCollection.Item(ID)
   End If
   
   Tgdt.EMP_ID = uctlSaleLookup.MyCombo.ItemData(Minus2Zero(uctlSaleLookup.MyCombo.ListIndex))
   Tgdt.EMP_NAME = uctlSaleLookup.MyCombo.Text
   Tgdt.EMP_CODE = uctlSaleLookup.MyTextBox.Text
   
   Tgdt.TARGET_PRICE1 = Val(txtTargetPrice1.Text)
   Tgdt.TARGET_PRICE2 = Val(txtTargetPrice2.Text)
   Tgdt.TARGET_PRICE3 = Val(txtTargetPrice3.Text)
   Tgdt.TARGET_PRICE4 = Val(txtTargetPrice4.Text)
   Tgdt.TARGET_PRICE5 = Val(txtTargetPrice5.Text)
   Tgdt.TARGET_PRICE6 = Val(txtTargetPrice6.Text)
   Tgdt.TARGET_PRICE7 = Val(txtTargetPrice7.Text)
   Tgdt.TARGET_PRICE8 = Val(txtTargetPrice8.Text)
   Tgdt.TARGET_PRICE9 = Val(txtTargetPrice9.Text)
   Tgdt.TARGET_PRICE10 = Val(txtTargetPrice10.Text)
   Tgdt.TARGET_PRICE11 = Val(txtTargetPrice11.Text)
   Tgdt.TARGET_PRICE12 = Val(txtTargetPrice12.Text)
   
      
   If ShowMode = SHOW_ADD Then
      Tgdt.Flag = "A"
      Call TempCollection.add(Tgdt)
   Else
      If Tgdt.Flag <> "A" Then
         Tgdt.Flag = "E"
      End If
   End If
   
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadEmployee(uctlSaleLookup.MyCombo, collEmp)
      Set uctlSaleLookup.MyCollection = collEmp
      
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
   
   Set collEmp = New Collection
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set collEmp = Nothing
End Sub


Private Sub uctlSaleLookup_Change()
   m_HasModify = True
End Sub

Private Sub txtTargetPrice1_Change()
   m_HasModify = True
End Sub
Private Sub txtTargetPrice2_Change()
   m_HasModify = True
End Sub
Private Sub txtTargetPrice3_Change()
   m_HasModify = True
End Sub
Private Sub txtTargetPrice4_Change()
   m_HasModify = True
End Sub
Private Sub txtTargetPrice5_Change()
   m_HasModify = True
End Sub
Private Sub txtTargetPrice6_Change()
   m_HasModify = True
End Sub
Private Sub txtTargetPrice7_Change()
   m_HasModify = True
End Sub
Private Sub txtTargetPrice8_Change()
   m_HasModify = True
End Sub
Private Sub txtTargetPrice9_Change()
   m_HasModify = True
End Sub
Private Sub txtTargetPrice10_Change()
   m_HasModify = True
End Sub
Private Sub txtTargetPrice11_Change()
   m_HasModify = True
End Sub
Private Sub txtTargetPrice12_Change()
   m_HasModify = True
End Sub
