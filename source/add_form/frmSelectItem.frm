VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelectItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Icon            =   "frmSelectItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   10410
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8445
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   14896
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextLookup uctlItem 
         Height          =   495
         Left            =   1860
         TabIndex        =   0
         Top             =   1320
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   873
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   5655
         Left            =   1860
         TabIndex        =   2
         Top             =   1920
         Width           =   8145
         _ExtentX        =   14367
         _ExtentY        =   9975
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "ImageList1"
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSelectItem.frx":27A2
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   7800
         TabIndex        =   1
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmSelectItem.frx":307C
         ButtonStyle     =   3
      End
      Begin Threed.SSCheck chkSelectAll 
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   7680
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSCSelectAll"
      End
      Begin VB.Label lblItem 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   9
         Top             =   1980
         Width           =   1575
      End
      Begin VB.Label lblItemName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   8400
         TabIndex        =   5
         Top             =   7800
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   6720
         TabIndex        =   4
         Top             =   7800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmSelectItem.frx":3396
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmSelectItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Employee As CEmployee

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String

Public itemType As Long
Public TempCollection As Collection
Public TempCollection2 As Collection

Private Sub chkSelectAll_Click(value As Integer)
Dim N As Node
For Each N In TreeView1.Nodes
  If value = 1 Then
      N.Checked = True
  Else
      N.Checked = False
  End If
Next N
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub
Private Sub PopulateItem(Col As Collection)
Dim N As Node
Dim D1 As CCustomerGrade
Dim TempD1 As CCustomerGrade
Dim D2 As CCustomer
Dim TempD2 As CCustomer
   If itemType = 1 Then
      For Each D1 In Col
          Set TempD1 = GetObject("CCustomerGrade", TempCollection, Trim(str(D1.CSTGRADE_ID)), False)
         If TempD1 Is Nothing Then
            Call TempCollection.add(D1, Trim(str(D1.CSTGRADE_ID)))
         End If
      Next D1
      For Each N In TreeView1.Nodes
        If N.Checked Then
            Set D1 = GetObject("CCustomerGrade", TempCollection, Trim(str(N.Tag)), False)
            If Not D1 Is Nothing Then
               D1.SELECT_FLAG = "Y"
            End If
         Else
            Set D1 = GetObject("CCustomerGrade", TempCollection, Trim(str(N.Tag)), False)
            If Not D1 Is Nothing Then
               D1.SELECT_FLAG = "N"
            End If
         End If
         Set D1 = Nothing
      Next N
      
      Set Col = Nothing
      Set Col = New Collection
       For Each D1 In TempCollection 'copy กลับ
            Call Col.add(D1, Trim(str(D1.CSTGRADE_ID)))
      Next D1
      
  ElseIf itemType = 2 Then
    For Each D2 In Col
          Set TempD2 = GetObject("CCustomer", TempCollection, Trim(str(D2.CUSTOMER_ID)), False)
         If TempD2 Is Nothing Then
            Call TempCollection.add(D2, Trim(str(D2.CUSTOMER_ID)))
         End If
      Next D2
      For Each N In TreeView1.Nodes
        If N.Checked Then
            Set D2 = GetObject("CCustomer", TempCollection, Trim(str(N.Tag)), False)
            If Not D2 Is Nothing Then
               D2.SELECT_FLAG = "Y"
            End If
         Else
            Set D2 = GetObject("CCustomer", TempCollection, Trim(str(N.Tag)), False)
            If Not D2 Is Nothing Then
               D2.SELECT_FLAG = "N"
            End If
         End If
         Set D2 = Nothing
      Next N
      Set Col = Nothing
      Set Col = New Collection
       For Each D2 In TempCollection 'copy กลับ
          Call Col.add(D2, Trim(str(D2.CUSTOMER_ID)))
      Next D2
  End If
End Sub
Private Function SaveData() As Boolean
  Call PopulateItem(TempCollection2)
   SaveData = True
End Function

Private Sub cmdSearch_Click()
   Dim ID As Long
   Dim TempD1 As CCustomerGrade
   Dim TempD2 As CCustomer
   Dim TempColl As Collection
   If itemType = 1 Then
      ID = uctlItem.MyCombo.ItemData(Minus2Zero(uctlItem.MyCombo.ListIndex))
      If ID > 0 Then
         Set TempColl = New Collection
           Set TempD1 = GetObject("CCustomerGrade", TempCollection2, Trim(str(ID)), False)
            If Not TempD1 Is Nothing Then
               Call TempColl.add(TempD1, Trim(str(ID)))
            End If
      Else
            Set TempColl = New Collection
            Set TempD1 = Nothing
             For Each TempD1 In TempCollection2
               Call TempColl.add(TempD1, Trim(str(TempD1.CSTGRADE_ID)))
             Next TempD1
      End If
      If Not TempColl Is Nothing Then
         Call LoadItemView(TempColl)
      End If
      Set TempD1 = Nothing
   ElseIf itemType = 2 Then
      ID = uctlItem.MyCombo.ItemData(Minus2Zero(uctlItem.MyCombo.ListIndex))
      If ID > 0 Then
         Set TempColl = New Collection
           Set TempD2 = GetObject("CCustomer", TempCollection2, Trim(str(ID)), False)
            If Not TempD2 Is Nothing Then
               Call TempColl.add(TempD2, Trim(str(ID)))
            End If
      ElseIf ID = 0 And uctlItem.MyTextBox.Text = "" Then
            Set TempColl = New Collection
            Set TempD2 = Nothing
             For Each TempD2 In TempCollection2
               Call TempColl.add(TempD2, Trim(str(TempD2.CUSTOMER_ID)))
             Next TempD2
      End If
      If Not TempColl Is Nothing Then
         Call LoadItemView(TempColl)
      End If
      Set TempD2 = Nothing
   End If
End Sub

Private Sub Form_Activate()
Dim DG As CCustomerGrade
Dim DC As CCustomer
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      
      If itemType = 1 Then
       If TempCollection.Count = 0 Then
         Call LoadCustomerGrade(uctlItem.MyCombo, TempCollection2)
         Set uctlItem.MyCollection = TempCollection2
       Else
       Call LoadCustomerGrade(uctlItem.MyCombo)
      For Each DG In TempCollection
         Call TempCollection2.add(DG, Trim(str(DG.CSTGRADE_ID)))
      Next DG
      Set uctlItem.MyCollection = TempCollection2
      End If
      ElseIf itemType = 2 Then
        If TempCollection.Count = 0 Then
         Call LoadCustomer(uctlItem.MyCombo, TempCollection2)
         Set uctlItem.MyCollection = TempCollection2
       Else
        Call LoadCustomer(uctlItem.MyCombo)
         For Each DC In TempCollection
            Call TempCollection2.add(DC, Trim(str(DC.CUSTOMER_ID)))
         Next DC
         Set uctlItem.MyCollection = TempCollection2
      End If
      End If
      
      Call LoadItemView(TempCollection2)
      
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
'   ElseIf Shift = 0 And KeyCode = 117 Then
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

Private Sub LoadItemView(Col As Collection)
Dim C1 As CCustomerGrade
Dim C2 As CCustomer
Dim N As Node
Dim Np As Node

Call TreeView1.Nodes.Clear
If itemType = 1 Then
   For Each C1 In Col
         Set N = TreeView1.Nodes.add(, tvwFirst, Trim(str(C1.CSTGRADE_ID) & "-C"), C1.CSTGRADE_NAME & " (" & C1.CSTGRADE_NO & ")", 1, 1)
         N.Tag = C1.CSTGRADE_ID
         N.Checked = False
         N.Expanded = False
         
         If C1.SELECT_FLAG = "Y" Then
            N.Checked = True
         End If
   Next C1
 ElseIf itemType = 2 Then
   For Each C2 In Col
      Set N = TreeView1.Nodes.add(, tvwFirst, Trim(str(C2.CUSTOMER_ID) & "-C"), C2.CUSTOMER_NAME & " (" & C2.CUSTOMER_CODE & ")", 1, 1)
      N.Tag = C2.CUSTOMER_ID
      N.Checked = False
      N.Expanded = False
      
      If C2.SELECT_FLAG = "Y" Then
         N.Checked = True
      End If
      
   Next C2
 End If
End Sub
Private Sub EditItemView(Key As String)
Dim C1 As CCustomerGrade
Dim C2 As CCustomer
Dim N As Node
Dim Np As Node
Dim I As Long
Dim N2 As TreeView

For Each N In TreeView1.Nodes
  If N.Key = Trim(Key & "-C") Then ' Trim(str(C1.CSTGRADE_ID) & "-C") Then
  N.Selected = True
  Else
  N.Selected = False
  End If
Next N

End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   If itemType = 1 Then
      pnlHeader.Caption = "เลือกระดับลูกค้า"
      Call InitNormalLabel(lblItemName, "ระดับลูกค้า")
      Call InitNormalLabel(lblItem, "รายการระดับลูกค้า")
   ElseIf itemType = 2 Then
      pnlHeader.Caption = "เลือกชื่อลูกค้า"
      Call InitNormalLabel(lblItemName, "ชื่อลูกค้า")
      Call InitNormalLabel(lblItem, "รายการชื่อลูกค้า")
   End If
  Call InitCheckBox(chkSelectAll, "เลือกทั้งหมด")
  
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdSearch, MapText("ค้นหา"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
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
   Set TempCollection2 = New Collection
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set TempCollection2 = Nothing
End Sub



Private Sub TreeView1_Click()
 If Not SaveData Then
      Exit Sub
   End If
End Sub

Private Sub uctlItem_Change()
'Dim ID As Long
'Dim TempD1 As CCustomerGrade
'Dim TempD2 As CCustomer
'Dim TempColl As Collection
'If itemType = 1 Then
'   ID = uctlItem.MyCombo.ItemData(Minus2Zero(uctlItem.MyCombo.ListIndex))
'   If ID > 0 Then
'      Set TempColl = New Collection
'        Set TempD1 = GetObject("CCustomerGrade", TempCollection2, Trim(str(ID)), False)
'         If Not TempD1 Is Nothing Then
'            Call TempColl.add(TempD1, Trim(str(ID)))
'         End If
'   Else
'         Set TempColl = New Collection
'         Set TempD1 = Nothing
'          For Each TempD1 In TempCollection2
'            Call TempColl.add(TempD1, Trim(str(TempD1.CSTGRADE_ID)))
'          Next TempD1
'   End If
'   Call LoadItemView(TempColl)
'   Set TempD1 = Nothing
'ElseIf itemType = 2 Then
'   ID = uctlItem.MyCombo.ItemData(Minus2Zero(uctlItem.MyCombo.ListIndex))
'   If ID > 0 Then
'      Set TempColl = New Collection
'        Set TempD2 = GetObject("CCustomer", TempCollection2, Trim(str(ID)), False)
'         If Not TempD2 Is Nothing Then
'            Call TempColl.add(TempD2, Trim(str(ID)))
'         End If
'   Else
'         Set TempColl = New Collection
'         Set TempD2 = Nothing
'          For Each TempD2 In TempCollection2
'            Call TempColl.add(TempD2, Trim(str(TempD2.CUSTOMER_ID)))
'          Next TempD2
'   End If
'   Call LoadItemView(TempColl)
'   Set TempD2 = Nothing
'End If
End Sub
