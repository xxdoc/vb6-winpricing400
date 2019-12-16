VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditPackingListItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditPackingListItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   6105
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   4005
      Left            =   0
      TabIndex        =   9
      Top             =   600
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   7064
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextBox txtPkgNo 
         Height          =   435
         Left            =   1800
         TabIndex        =   0
         Top             =   270
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtMeasure 
         Height          =   435
         Left            =   1800
         TabIndex        =   3
         Top             =   1605
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtDesc 
         Height          =   435
         Left            =   1800
         TabIndex        =   1
         Top             =   720
         Width           =   3855
         _ExtentX        =   14790
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtQuantity 
         Height          =   435
         Left            =   1800
         TabIndex        =   2
         Top             =   1155
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtNetWeight 
         Height          =   435
         Left            =   1800
         TabIndex        =   4
         Top             =   2040
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtGrossWeight 
         Height          =   435
         Left            =   1800
         TabIndex        =   5
         Top             =   2475
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   767
      End
      Begin Threed.SSCheck setflag 
         Height          =   435
         Left            =   3240
         TabIndex        =   16
         Top             =   1200
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblNetWeight 
         Alignment       =   1  'Right Justify
         Caption         =   "lblNetWeight"
         Height          =   465
         Left            =   240
         TabIndex        =   15
         Top             =   2040
         Width           =   1485
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   1440
         TabIndex        =   6
         Top             =   3150
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPackingListItem.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   3060
         TabIndex        =   7
         Top             =   3150
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblMeasure 
         Alignment       =   1  'Right Justify
         Caption         =   "lblMeasure"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   1680
         Width           =   1485
      End
      Begin VB.Label lblGrossWeight 
         Alignment       =   1  'Right Justify
         Caption         =   "lblGrossWeight"
         Height          =   465
         Left            =   240
         TabIndex        =   13
         Top             =   2520
         Width           =   1485
      End
      Begin VB.Label lblPkgNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPkgNo"
         Height          =   375
         Left            =   0
         TabIndex        =   12
         Top             =   300
         Width           =   1725
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "lblDesc"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   780
         Width           =   1485
      End
      Begin VB.Label lblQuantity 
         Alignment       =   1  'Right Justify
         Caption         =   "lblQuantity"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1230
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmAddEditPackingListItem"
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
      
   Call InitNormalLabel(lblPkgNo, MapText("หมายเลขหีบห่อ"))
   Call InitNormalLabel(lblDesc, MapText("รายละเอียด"))
   Call InitNormalLabel(lblQuantity, MapText("ปริมาณ/หีบห่อ"))
   Call InitNormalLabel(lblMeasure, MapText("จำนวนหีบห่อ"))
   Call InitNormalLabel(lblNetWeight, MapText("น้ำหนักสุทธิ"))
   Call InitNormalLabel(lblGrossWeight, MapText("น้ำหนักรวม"))
   
   Call txtPkgNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtDesc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   
   Call txtQuantity.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.CODE_TYPE)
   Call txtMeasure.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.CODE_TYPE)
   
   Call txtNetWeight.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.CODE_TYPE)
   Call txtGrossWeight.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.CODE_TYPE)
   
      Call InitCheckBox(setflag, "SET UNIT")
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
           Dim CustProof As CPkglst
         Set CustProof = TempCollection.Item(ID)
         setflag.Value = FlagToCheck(CustProof.SET_FLAG)
         txtPkgNo.Text = CustProof.PKG_NUMBER
         txtDesc.Text = CustProof.DESCRIPTION
         txtQuantity.Text = CustProof.QUANTITY
         txtMeasure.Text = CustProof.MEASURE
         txtNetWeight.Text = CustProof.NET_WEIGHT
         txtGrossWeight.Text = CustProof.GROSS_WEIGHT
         
         
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

   If Not VerifyTextControl(lblPkgNo, txtPkgNo, False) Then
      Exit Function
   End If

   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
      Dim CustProof As CPkglst
   If ShowMode = SHOW_ADD Then
      Set CustProof = New CPkglst
    Else
      Set CustProof = TempCollection.Item(ID)
   End If
   
   CustProof.PKG_NUMBER = txtPkgNo.Text
   CustProof.DESCRIPTION = txtDesc.Text
   CustProof.QUANTITY = Val(txtQuantity.Text)
   CustProof.MEASURE = Val(txtMeasure.Text)
   CustProof.NET_WEIGHT = Val(txtNetWeight.Text)
   CustProof.GROSS_WEIGHT = Val(txtGrossWeight.Text)
    CustProof.SET_FLAG = Check2Flag(setflag.Value)
   If ShowMode = SHOW_ADD Then
      CustProof.Flag = "A"
      Call TempCollection.add(CustProof)
      Else
      If CustProof.Flag <> "A" Then
      CustProof.Flag = "E"
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
   Set m_Rs = New ADODB.Recordset
End Sub


Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
End Sub


Private Sub setflag_Click(Value As Integer)
m_HasModify = True
End Sub

Private Sub txtDesc_Change()
m_HasModify = True
End Sub

Private Sub txtGrossWeight_Change()
m_HasModify = True
End Sub

Private Sub txtMeasure_Change()
m_HasModify = True
End Sub

Private Sub txtNetWeight_Change()
m_HasModify = True
End Sub

Private Sub txtPkgNo_Change()
m_HasModify = True
End Sub

Private Sub txtQuantity_Change()
m_HasModify = True
End Sub
