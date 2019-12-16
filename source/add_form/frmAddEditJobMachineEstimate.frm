VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditJobMachineEstimate 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3750
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
   Icon            =   "frmAddEditJobMachineEstimate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
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
      Height          =   3165
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   5583
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTime uctlTime 
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   1440
         Width           =   1095
         _ExtentX        =   4683
         _ExtentY        =   661
      End
      Begin prjFarmManagement.uctlDate uctlUseDate 
         Height          =   405
         Left            =   1800
         TabIndex        =   1
         Top             =   840
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextLookup uctlMachineLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   0
         Top             =   240
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin VB.Label lblHour 
         Alignment       =   1  'Right Justify
         Caption         =   "lblHour"
         Height          =   375
         Left            =   3000
         TabIndex        =   10
         Top             =   1440
         Width           =   1245
      End
      Begin VB.Label lblMachine 
         Alignment       =   1  'Right Justify
         Caption         =   "lblMachine"
         Height          =   315
         Left            =   360
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   1920
         TabIndex        =   3
         Top             =   2280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJobMachineEstimate.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4560
         TabIndex        =   4
         Top             =   2280
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
         Top             =   1440
         Width           =   1485
      End
      Begin VB.Label lblUseDate 
         Alignment       =   1  'Right Justify
         Caption         =   "lblUseDate"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmAddEditJobMachineEstimate"
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
Private m_Machine_combo As Collection
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
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
      
   Call InitNormalLabel(lblMachine, MapText("เครื่องจักร"))
   Call InitNormalLabel(lblUseDate, MapText("วันที่ใช้"))
   Call InitNormalLabel(lblUseTime, MapText("เวลาที่ใช้"))
   Call InitNormalLabel(lblHour, MapText("(ชั่วโมง/นาที)"))
   
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
           Dim MA As CJobResource
         Set MA = TempCollection.Item(ID)
         
         uctlMachineLookup.MyCombo.ListIndex = IDToListIndex(uctlMachineLookup.MyCombo, MA.MACHINE_ID)
         uctlUseDate.ShowDate = MA.OCCUPY_DATE
          uctlTime.HR = MA.MACHINE_ID_HOUR
        uctlTime.MI = MA.MACHINE_ID_HOURN
      
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

   If Not VerifyCombo(lblMachine, uctlMachineLookup.MyCombo, False) Then
      Exit Function
   End If
      If Not VerifyTime(lblUseTime, uctlTime, False) Then
      Exit Function
   End If
If Not VerifyDate(lblUseDate, uctlUseDate, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
      Dim MA As CJobResource
   If ShowMode = SHOW_ADD Then
      Set MA = New CJobResource
   Else
      Set MA = TempCollection.Item(ID)
   End If
   MA.MACHINE_ID = uctlMachineLookup.MyCombo.ItemData(Minus2Zero(uctlMachineLookup.MyCombo.ListIndex))
   MA.MACHINE_NO = uctlMachineLookup.MyTextBox.Text
   MA.MACHINE_NAME = uctlMachineLookup.MyCombo.Text
   MA.OCCUPY_DATE = uctlUseDate.ShowDate
   MA.MACHINE_ID_HOUR = uctlTime.HR
   MA.MACHINE_ID_HOURN = uctlTime.MI
   If ShowMode = SHOW_ADD Then
      MA.Flag = "A"
      Call TempCollection.Add(MA)
      Else
      If MA.Flag <> "A" Then
      MA.Flag = "E"
      End If
          End If
   
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadMachine(uctlMachineLookup.MyCombo, m_Machine_combo)
      Set uctlMachineLookup.MyCollection = m_Machine_combo
      
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
   End If
End Sub

Private Sub Form_Load()

   OKClick = False
   Call InitFormLayout
   
   m_HasActivate = False
   m_HasModify = False
   Set m_Machine_combo = New Collection
   Set m_Rs = New ADODB.Recordset
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
End Sub

Private Sub uctlMachineLookup_Change()
m_HasModify = True
End Sub

Private Sub uctlTime_HasChange()
m_HasModify = True
End Sub

Private Sub uctlUseDate_HasChange()
m_HasModify = True
End Sub
