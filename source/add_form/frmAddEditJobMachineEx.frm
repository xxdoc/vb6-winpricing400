VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditJobMachineEx 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3300
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
   Icon            =   "frmAddEditJobMachineEx.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2715
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   4789
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextBox uctlTextBox1 
         Height          =   465
         Left            =   1800
         TabIndex        =   2
         Top             =   1200
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlDate uctlUseDate 
         Height          =   405
         Left            =   1800
         TabIndex        =   1
         Top             =   780
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextLookup uctlPeopleLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   0
         Top             =   330
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin VB.Label lblHour 
         Caption         =   "lblHour"
         Height          =   375
         Left            =   3180
         TabIndex        =   10
         Top             =   1260
         Width           =   1005
      End
      Begin VB.Label lblPeople 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPeople"
         Height          =   315
         Left            =   270
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2385
         TabIndex        =   3
         Top             =   1890
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJobMachineEx.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4005
         TabIndex        =   4
         Top             =   1890
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblUseTime 
         Alignment       =   1  'Right Justify
         Caption         =   "lblUseTime"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1230
         Width           =   1485
      End
      Begin VB.Label lblUseDate 
         Alignment       =   1  'Right Justify
         Caption         =   "lblUseDate"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   780
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmAddEditJobMachineEx"
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
Private m_People_combo As Collection
Private m_Machines As Collection

Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public TempCollection As Collection

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
      
   Call InitNormalLabel(lblPeople, MapText("เครื่องจักร"))
   Call InitNormalLabel(lblUseDate, MapText("วันที่ใช้"))
   Call InitNormalLabel(lblUseTime, MapText("จำนวน"))
   Call InitNormalLabel(lblHour, MapText("(ชั่วโมง)"))
   
   Call uctlTextBox1.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
           Dim Ma As CJobResource
         Set Ma = TempCollection.Item(ID)
         
         uctlPeopleLookup.MyCombo.ListIndex = IDToListIndex(uctlPeopleLookup.MyCombo, Ma.MACHINE_ID)
         uctlUseDate.ShowDate = Ma.OCCUPY_DATE
         uctlTextBox1.Text = Ma.OCCUPY_INTERVAL
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

   If Not VerifyCombo(lblPeople, uctlPeopleLookup.MyCombo, True) Then
      Exit Function
   End If
   If Not VerifyDate(lblUseDate, uctlUseDate, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim Ma As CJobResource
   If ShowMode = SHOW_ADD Then
      Set Ma = New CJobResource
   Else
      Set Ma = TempCollection.Item(ID)
   End If
   
   Ma.POSITION_ID = -1
   Ma.POSITION_NAME = ""
   Ma.EMP_ID = -1
   Ma.MACHINE_ID = uctlPeopleLookup.MyCombo.ItemData(Minus2Zero(uctlPeopleLookup.MyCombo.ListIndex))
   Ma.MACHINE_NO = uctlPeopleLookup.MyTextBox.Text
   Ma.MACHINE_NAME = uctlPeopleLookup.MyCombo.Text
   Ma.OCCUPY_DATE = uctlUseDate.ShowDate
   Ma.OCCUPY_INTERVAL = Val(uctlTextBox1.Text)
   
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
      
      Call LoadMachine(uctlPeopleLookup.MyCombo, m_Machines)
      Set uctlPeopleLookup.MyCollection = m_Machines
    
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
   Set m_People_combo = New Collection
   Set m_Rs = New ADODB.Recordset
   Set m_Machines = New Collection
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_Machines = Nothing
End Sub

Private Sub uctlPeopleLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlTime_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlTextBox1_Change()
   m_HasModify = True
End Sub

Private Sub uctlUseDate_HasChange()
   m_HasModify = True
End Sub

