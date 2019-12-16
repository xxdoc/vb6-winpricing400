VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditGlDetail 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9300
   Icon            =   "frmAddEditGlDetail.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   9300
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3705
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   6535
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   9
         Top             =   0
         Width           =   9285
         _ExtentX        =   16378
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextLookup uctlGl 
         Height          =   465
         Left            =   1920
         TabIndex        =   0
         Top             =   840
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtGlDesc 
         Height          =   435
         Left            =   1920
         TabIndex        =   1
         Top             =   1320
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtGlAmount 
         Height          =   435
         Left            =   1920
         TabIndex        =   4
         Top             =   2280
         Width           =   1515
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin Threed.SSOption RadCredit 
         Height          =   375
         Left            =   3600
         TabIndex        =   3
         Top             =   1800
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSOption4"
      End
      Begin Threed.SSOption radDebit 
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   1800
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSOption4"
         Value           =   -1
      End
      Begin VB.Label lblGlAmount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   2340
         Width           =   1485
      End
      Begin VB.Label lblGlDesc 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   1380
         Width           =   1485
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   2205
         TabIndex        =   5
         Top             =   2820
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditGlDetail.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label lblGl 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   10
         Top             =   900
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5505
         TabIndex        =   7
         Top             =   2820
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3855
         TabIndex        =   6
         Top             =   2820
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditGlDetail.frx":2ABC
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditGlDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_GLDetail As CGLDetail

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Public ParentForm As Object
Public TempCollection As Collection
Public ListType As Long

Private m_Gl As Collection
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
         
         'Call ParentForm.RefreshGridAccountList()
         Exit Sub
      End If
      
      ID = NewID
      Call QueryData(True)
   ElseIf ShowMode = SHOW_ADD Then
      uctlGl.MyCombo.ListIndex = -1
      uctlGl.MyTextBox.Text = ""
      txtGlDesc.Text = ""
      txtGlAmount.Text = ""
      uctlGl.MyTextBox.SetFocus
   End If
   
   Call ParentForm.RefreshGridAccountList(ListType)

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

   If Flag Then
      Call EnableForm(Me, False)
      
      Dim Acl As CGLDetail
      Set Acl = TempCollection.Item(ID)
      
      uctlGl.MyCombo.ListIndex = IDToListIndex(uctlGl.MyCombo, Acl.GetFieldValue("GL_ID"))
      txtGlDesc.Text = Acl.GetFieldValue("GL_DESC")
      If Acl.GetFieldValue("GL_TYPE") = 1 Then  'Dr
         radDebit.Value = True
      ElseIf Acl.GetFieldValue("GL_TYPE") = 2 Then 'Cr
         RadCredit.Value = True
      End If
      txtGlAmount.Text = Val(Acl.GetFieldValue("GL_AMOUNT"))
      
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim PaymentType As Long
Dim GL As CGLDetail
   
  If Not VerifyCombo(lblGl, uctlGl.MyCombo, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim Acl As CGLDetail
   
   If ShowMode = SHOW_ADD Then
      Set Acl = New CGLDetail

      Acl.Flag = "A"

      Call TempCollection.add(Acl)
   Else
      Set Acl = TempCollection.Item(ID)
      If Acl.Flag <> "A" Then
         Acl.Flag = "E"
      End If
   End If

   Call Acl.SetFieldValue("Gl_ID", uctlGl.MyCombo.ItemData(Minus2Zero(uctlGl.MyCombo.ListIndex)))
   Call Acl.SetFieldValue("Gl_NO", uctlGl.MyTextBox.Text)
   Call Acl.SetFieldValue("Gl_NAME", uctlGl.MyCombo.Text)
   
   Call Acl.SetFieldValue("GL_DESC", txtGlDesc.Text)
   
   Dim tempSum As Double
   If Len(txtGlAmount.Text) = 0 And RadCredit.Value Then
      For Each GL In TempCollection
         If GL.GetFieldValue("GL_TYPE") = 1 Then
               tempSum = tempSum + GL.GetFieldValue("GL_AMOUNT")
         End If
      Next GL
      Call Acl.SetFieldValue("GL_AMOUNT", Val(tempSum))
   Else
      Call Acl.SetFieldValue("GL_AMOUNT", Val(txtGlAmount.Text))
   End If
   
   If radDebit.Value Then
      Call Acl.SetFieldValue("GL_TYPE", 1)
   Else
      Call Acl.SetFieldValue("GL_TYPE", 2)
   End If
   
   Set Acl = Nothing

   Call EnableForm(Me, True)
   SaveData = True
End Function
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadMaster(uctlGl.MyCombo, m_Gl, ACCOUNT_LIST)
      Set uctlGl.MyCollection = m_Gl
      
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
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
   
  
   Call InitNormalLabel(lblGl, MapText("เลขที่บัญชี"))
   Call InitNormalLabel(lblGlDesc, MapText("รายละเอียด"))
   Call InitNormalLabel(lblGlAmount, MapText("จำนวนเงิน"))
  
      
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitOptionEx(radDebit, "Dr.")
   Call InitOptionEx(RadCredit, "Cr.")
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdNext, MapText("ถัดไป"))
   
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
   
   Set m_GLDetail = New CGLDetail
   Set m_Rs = New ADODB.Recordset
   
   Set m_Gl = New Collection
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set m_GLDetail = Nothing
   
   Set m_Gl = Nothing

End Sub

Private Sub RadCredit_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub RadCredit_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub radDebit_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub radDebit_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub txtGlAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtGlDesc_Change()
   m_HasModify = True
End Sub

Private Sub uctlGl_Change()
   m_HasModify = True
End Sub
