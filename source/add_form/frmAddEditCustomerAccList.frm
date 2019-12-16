VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditCustomerAccList 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9300
   Icon            =   "frmAddEditCustomerAccList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   9300
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3345
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   5900
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   4
         Top             =   0
         Width           =   9285
         _ExtentX        =   16378
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextLookup uctlItemLookup 
         Height          =   465
         Left            =   1920
         TabIndex        =   8
         Top             =   960
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextLookup uctlDebit 
         Height          =   465
         Left            =   1920
         TabIndex        =   9
         Top             =   1440
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextLookup uctlCredit 
         Height          =   465
         Left            =   1920
         TabIndex        =   10
         Top             =   1920
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   820
      End
      Begin VB.Label lblCredit 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   7
         Top             =   1950
         Width           =   1575
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   2205
         TabIndex        =   0
         Top             =   2580
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCustomerAccList.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label lblItemLookUp 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   270
         TabIndex        =   6
         Top             =   1080
         Width           =   1485
      End
      Begin VB.Label lblDebit 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   5
         Top             =   1500
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5505
         TabIndex        =   2
         Top             =   2580
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3855
         TabIndex        =   1
         Top             =   2580
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCustomerAccList.frx":2ABC
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditCustomerAccList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_CustomerAccountList As CCustomerAccountList

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Public ParentForm As Object
Public TempCollection As Collection

Public AccountListType As Long

Private m_Item As Collection
Private m_Debit As Collection
Private m_Credit As Collection
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
         
         Call ParentForm.RefreshGridAccountList(AccountListType)
         Exit Sub
      End If
      
      ID = NewID
      Call QueryData(True)
   ElseIf ShowMode = SHOW_ADD Then
      uctlItemLookup.MyCombo.ListIndex = -1
      uctlItemLookup.MyTextBox.Text = ""
      uctlDebit.MyCombo.ListIndex = -1
      uctlDebit.MyTextBox.Text = ""
      uctlCredit.MyCombo.ListIndex = -1
      uctlCredit.MyTextBox.Text = ""

   End If
   
   Call ParentForm.RefreshGridAccountList(AccountListType)
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
      
      Dim Acl As CCustomerAccountList
      Set Acl = TempCollection.Item(ID)
      
      If AccountListType = 1 Then
         uctlItemLookup.MyCombo.ListIndex = IDToListIndex(uctlItemLookup.MyCombo, Acl.GetFieldValue("PART_GROUP_ID"))
      ElseIf AccountListType = 2 Then
         uctlItemLookup.MyCombo.ListIndex = IDToListIndex(uctlItemLookup.MyCombo, Acl.GetFieldValue("FEATURE_TYPE"))
      ElseIf AccountListType = 3 Then
         uctlItemLookup.MyCombo.ListIndex = IDToListIndex(uctlItemLookup.MyCombo, Acl.GetFieldValue("BANK_ACCOUNT_ID"))
      End If
      
      uctlDebit.MyCombo.ListIndex = IDToListIndex(uctlDebit.MyCombo, Acl.GetFieldValue("DEBIT_ID"))
      uctlCredit.MyCombo.ListIndex = IDToListIndex(uctlCredit.MyCombo, Acl.GetFieldValue("CREDIT_ID"))
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim PaymentType As Long
   
'   If Not CheckUniqueNs(USERNAME_UNIQUE, txtCustomerAccountListNo.Text, ID) Then
'      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCustomerAccountListNo.Text & " " & MapText("อยู่ในระบบแล้ว")
'      glbErrorLog.ShowUserError
'      Exit Function
'   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim Acl As CCustomerAccountList
   
   If ShowMode = SHOW_ADD Then
      Set Acl = New CCustomerAccountList

      Acl.Flag = "A"

      Call TempCollection.add(Acl)
   Else
      Set Acl = TempCollection.Item(ID)
      If Acl.Flag <> "A" Then
         Acl.Flag = "E"
      End If
   End If

   Call Acl.SetFieldValue("ACCOUNT_LIST_TYPE", AccountListType)
   Call Acl.SetFieldValue("DEBIT_ID", uctlDebit.MyCombo.ItemData(Minus2Zero(uctlDebit.MyCombo.ListIndex)))
   
   Call Acl.SetFieldValue("DEBIT_NO", uctlDebit.MyTextBox.Text)
   Call Acl.SetFieldValue("DEBIT_NAME", uctlDebit.MyCombo.Text)
   
   Call Acl.SetFieldValue("CREDIT_ID", uctlCredit.MyCombo.ItemData(Minus2Zero(uctlCredit.MyCombo.ListIndex)))
   
   Call Acl.SetFieldValue("CREDIT_NO", uctlCredit.MyTextBox.Text)
   Call Acl.SetFieldValue("CREDIT_NAME", uctlCredit.MyCombo.Text)
   
   If AccountListType = 1 Then
      Call Acl.SetFieldValue("PART_GROUP_ID", uctlItemLookup.MyCombo.ItemData(Minus2Zero(uctlItemLookup.MyCombo.ListIndex)))
      Call Acl.SetFieldValue("PART_GROUP_NO", uctlItemLookup.MyTextBox.Text)
      Call Acl.SetFieldValue("PART_GROUP_NAME", uctlItemLookup.MyCombo.Text)
   ElseIf AccountListType = 2 Then
      Call Acl.SetFieldValue("FEATURE_TYPE", uctlItemLookup.MyCombo.ItemData(Minus2Zero(uctlItemLookup.MyCombo.ListIndex)))
      Call Acl.SetFieldValue("FEATURE_TYPE_NO", uctlItemLookup.MyTextBox.Text)
      Call Acl.SetFieldValue("FEATURE_TYPE_NAME", uctlItemLookup.MyCombo.Text)
   ElseIf AccountListType = 3 Then
      Call Acl.SetFieldValue("BANK_ACCOUNT_ID", uctlItemLookup.MyCombo.ItemData(Minus2Zero(uctlItemLookup.MyCombo.ListIndex)))
      Call Acl.SetFieldValue("BANK_ACCOUNT_NO", uctlItemLookup.MyTextBox.Text)
      Call Acl.SetFieldValue("BANK_ACCOUNT_NAME", uctlItemLookup.MyCombo.Text)
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
      
      If AccountListType = 1 Then
         Call LoadPartGroup(uctlItemLookup.MyCombo, m_Item)
         Set uctlItemLookup.MyCollection = m_Item
      ElseIf AccountListType = 2 Then
         Call LoadFeatureType(uctlItemLookup.MyCombo, m_Item)
         Set uctlItemLookup.MyCollection = m_Item
      ElseIf AccountListType = 3 Then
         Call LoadMaster(uctlItemLookup.MyCombo, m_Item, BANK_ACCOUNT)
         Set uctlItemLookup.MyCollection = m_Item
      End If
      
      Call LoadMaster(uctlDebit.MyCombo, m_Debit, ACCOUNT_LIST)
      Set uctlDebit.MyCollection = m_Debit
      
      Call LoadMaster(uctlCredit.MyCombo, m_Credit, ACCOUNT_LIST)
      Set uctlCredit.MyCollection = m_Credit
      
      
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

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   If AccountListType = 1 Then
      Call InitNormalLabel(lblItemLookUp, MapText("กลุ่มสินค้า"))
   ElseIf AccountListType = 2 Then
      Call InitNormalLabel(lblItemLookUp, MapText("ประเภทบริการ"))
   ElseIf AccountListType = 3 Then
      Call InitNormalLabel(lblItemLookUp, MapText("เลขที่บัญชี"))
   End If
   
   Call InitNormalLabel(lblDebit, MapText("DEBIT"))
   Call InitNormalLabel(lblCredit, MapText("CREDIT"))
      
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
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
   
   Set m_CustomerAccountList = New CCustomerAccountList
   Set m_Rs = New ADODB.Recordset
   
   Set m_Item = New Collection
   Set m_Debit = New Collection
   Set m_Credit = New Collection
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set m_CustomerAccountList = Nothing
   
   Set m_Item = Nothing
   Set m_Debit = Nothing
   Set m_Credit = Nothing

End Sub

Private Sub uctlCredit_Change()
   m_HasModify = True
End Sub

Private Sub uctlDebit_Change()
   m_HasModify = True
End Sub

Private Sub uctlItemLookup_Change()
   m_HasModify = True
End Sub
