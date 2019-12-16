VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditCustomerMKTFol 
   Caption         =   "Form1"
   ClientHeight    =   3345
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9330
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   9330
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin prjFarmManagement.uctlDate uctlDate 
      Height          =   405
      Left            =   1800
      TabIndex        =   7
      Top             =   840
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   714
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2835
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   5001
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextBox txtDesc 
         Height          =   465
         Left            =   1800
         TabIndex        =   2
         Top             =   780
         Width           =   6405
         _ExtentX        =   7488
         _ExtentY        =   820
      End
      Begin Threed.SSCheck chkStatusCustomer 
         Height          =   375
         Left            =   1800
         TabIndex        =   10
         Top             =   1440
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4740
         TabIndex        =   6
         Top             =   1980
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3090
         TabIndex        =   5
         Top             =   1980
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCustomerMKTFol.frx":0000
         ButtonStyle     =   3
      End
      Begin VB.Label lblDate 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   4
         Top             =   360
         Width           =   1485
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   3
         Top             =   840
         Width           =   1485
      End
   End
   Begin Threed.SSCheck chkBangkok 
      Height          =   405
      Left            =   1800
      TabIndex        =   9
      Top             =   1920
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   714
      _Version        =   131073
      Caption         =   "SSCheck1"
   End
   Begin VB.Label lblUseDate 
      Alignment       =   1  'Right Justify
      Caption         =   "lblUseDate"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1485
   End
End
Attribute VB_Name = "frmAddEditCustomerMKTFol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_CustomerMKTFol As CMKTFol

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Public ParentForm As Object
Public TempCollection As Collection
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   
   Call InitNormalLabel(lblDate, MapText("วันที่"))
   Call InitNormalLabel(lblDesc, MapText("รายละเอียด"))
    Call InitCheckBox(chkStatusCustomer, "สถานะประวัติของลูกค้า")
      
      
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
  
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
  
   
End Sub

Private Sub chkStatusCustomer_Change()
 m_HasModify = True
End Sub

Private Sub chkStatusCustomer_Click(Value As Integer)
m_HasModify = True
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub
Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
'   Call LoadAccessRight(Nothing, glbAccessRight, glbUser.GROUP_ID)
   OKClick = True
   Unload Me

End Sub

Private Sub Form_Activate()
If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
'      If AccountListType = 1 Then
'         Call LoadPartGroup(uctlItemLookup.MyCombo, m_Item)
'         Set uctlItemLookup.MyCollection = m_Item
'      ElseIf AccountListType = 2 Then
'         Call LoadFeatureType(uctlItemLookup.MyCombo, m_Item)
'         Set uctlItemLookup.MyCollection = m_Item
'      ElseIf AccountListType = 3 Then
'         Call LoadMaster(uctlItemLookup.MyCombo, m_Item, BANK_ACCOUNT)
'         Set uctlItemLookup.MyCollection = m_Item
'      End If
      
'      Call LoadMaster(uctlDebit.MyCombo, m_Debit, ACCOUNT_LIST)
'      Set uctlDebit.MyCollection = m_Debit
'
'      Call LoadMaster(uctlCredit.MyCombo, m_Credit, ACCOUNT_LIST)
'      Set uctlCredit.MyCollection = m_Credit
      
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
      End If
      
      m_HasModify = False
   End If
End Sub

Private Sub Form_Load()
   Call EnableForm(Me, False)
   m_HasActivate = False
   
   Set m_CustomerMKTFol = New CMKTFol
   Set m_Rs = New ADODB.Recordset
   
'   Set m_Item = New Collection
'   Set m_Debit = New Collection
'   Set m_Credit = New Collection
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set m_CustomerMKTFol = Nothing
   
'   Set m_Item = Nothing
'   Set m_Debit = Nothing
'   Set m_Credit = Nothing

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
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long


   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
      
      Dim MKTFol As CMKTFol
         Set MKTFol = TempCollection.Item(ID)

      
       uctlDate.ShowDate = MKTFol.FOL_DATE
        txtDesc.Text = MKTFol.FOL_NOTE
        chkStatusCustomer.Value = FlagToCheck(MKTFol.CANCEL_FLAG)
      End If
   End If
   
   Call EnableForm(Me, True)
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean

'   If Not CheckUniqueNs(USERNAME_UNIQUE, txtCustomerAccountListNo.Text, ID) Then
'      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCustomerAccountListNo.Text & " " & MapText("อยู่ในระบบแล้ว")
'      glbErrorLog.ShowUserError
'      Exit Function
'   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim MKTFol As CMKTFol
   
   If ShowMode = SHOW_ADD Then
      Set MKTFol = New CMKTFol

     MKTFol.Flag = "A"

      Call TempCollection.add(MKTFol)
   Else
      Set MKTFol = TempCollection.Item(ID)
      If MKTFol.Flag <> "A" Then
         MKTFol.Flag = "E"
      End If
   End If

  MKTFol.FOL_DATE = uctlDate.ShowDate
  MKTFol.FOL_NOTE = txtDesc.Text
  MKTFol.CANCEL_FLAG = Check2Flag(chkStatusCustomer.Value)

   Set MKTFol = Nothing

   Call EnableForm(Me, True)
   SaveData = True
End Function
Private Sub txtDesc_Change()
 m_HasModify = True
End Sub

Private Sub uctlDate_Change()
 m_HasModify = True
End Sub

Private Sub uctlDate_HasChange()
m_HasModify = True
End Sub
