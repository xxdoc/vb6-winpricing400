VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAddEditPrtItemSet 
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8490
   Icon            =   "frmAddEditPrtItemSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   8490
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   5700
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   10054
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   240
         Top             =   690
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
               Picture         =   "frmAddEditPrtItemSet.frx":27A2
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   2955
         Left            =   180
         TabIndex        =   2
         Top             =   1890
         Width           =   8145
         _ExtentX        =   14367
         _ExtentY        =   5212
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "ImageList1"
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtCode 
         Height          =   435
         Left            =   2430
         TabIndex        =   0
         Top             =   840
         Width           =   1845
         _ExtentX        =   4683
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtName 
         Height          =   435
         Left            =   2430
         TabIndex        =   1
         Top             =   1290
         Width           =   5745
         _ExtentX        =   4683
         _ExtentY        =   767
      End
      Begin VB.Label lblCode 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   330
         TabIndex        =   8
         Top             =   870
         Width           =   1965
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Height          =   465
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   2055
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2625
         TabIndex        =   3
         Top             =   5010
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPrtItemSet.frx":307C
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4275
         TabIndex        =   4
         Top             =   5010
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditPrtItemSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Private m_MasterRef As CMasterRef
Private m_PrtItemSet As CPrtItemSet

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long

Public m_PartItemGroupType As Collection

Private Sub LoadTreeView(Col As Collection)
Dim C As CPartItem
Dim N As Node
Dim Np As Node
Dim Key1 As String
Dim Key2 As String
Dim p As CPrtItemSet
Dim I As Long
      
      I = 0
      If I <= 0 Then
         Set N = TreeView1.Nodes.add(, tvwFirst, "ROOT", "รายการ", 1, 1)
         N.Tag = "R-99999"
         N.Expanded = True
      End If
      
      I = 1
      
      For Each C In Col
         If Key1 <> Trim(Str(C.PART_GROUP_ID)) & "-G" Then
            Set N = TreeView1.Nodes.add("ROOT", tvwChild, Trim(Str(C.PART_GROUP_ID)) & "-G", C.PART_GROUP_NAME & " (" & C.PART_GROUP_NO & ")", 1, 1)
            N.Tag = Trim(Str(C.PART_GROUP_ID))
            N.Expanded = True
            
            Key1 = Trim(Str(C.PART_GROUP_ID)) & "-G"
         End If
         
         If Key2 <> Trim(Str(C.PART_TYPE)) & "-T" Then
            Set N = TreeView1.Nodes.add(Key1, tvwChild, Trim(Str(C.PART_TYPE)) & "-T", C.PART_TYPE_NAME & " (" & C.PART_TYPE_NO & ")", 1, 1)
            N.Tag = Trim(Str(C.PART_TYPE))
            N.Expanded = True
            
            Key2 = Trim(Str(C.PART_TYPE)) & "-T"
         End If
         
         Set N = TreeView1.Nodes.add(Key2, tvwChild, "I-" & Trim(Str(C.PART_ITEM_ID)), C.PART_DESC & " (" & C.PART_NO & ")", 1, 1)
         N.Tag = Trim(Str(C.PART_ITEM_ID))
         For Each p In m_MasterRef.PrtItemSets
            If (p.GetFieldValue("PART_ITEM_ID") = C.PART_ITEM_ID) And (p.Flag <> "D") Then
               N.Checked = True
               Exit For
            End If
         Next p
         
         N.Expanded = False
      Next C
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long

   IsOK = True
   
   If Flag Then
      Call EnableForm(Me, False)
            
      m_MasterRef.KEY_ID = ID
      m_MasterRef.QueryFlag = 1
      If Not glbMaster.QueryMasterRef(m_MasterRef, m_Rs, itemcount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If itemcount > 0 Then
      Call m_MasterRef.PopulateFromRS(1, m_Rs)
      txtCode.Text = m_MasterRef.KEY_CODE
      txtName.Text = m_MasterRef.KEY_NAME
   End If
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
      
   Call LoadPartItemGroupType(Nothing, m_PartItemGroupType)
   Call LoadTreeView(m_PartItemGroupType)
   Call EnableForm(Me, True)
End Sub
Private Sub PopulateItem(Col As Collection)
Dim N As Node
Dim C As CPrtItemSet

   For Each C In Col
      C.Flag = "D"
   Next C
   
   For Each N In TreeView1.Nodes
      If N.Checked Then
         Set C = New CPrtItemSet
         C.Flag = "A"
         If Left(N.Key, 1) = "I" Then
            Call C.SetFieldValue("PART_ITEM_ID", N.Tag)
            Call Col.add(C)
            Set C = Nothing
         End If
      End If
   Next N
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean

      
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Call EnableForm(Me, False)
      
   Call PopulateItem(m_MasterRef.PrtItemSets)
   
   m_MasterRef.AddEditMode = ShowMode
   m_MasterRef.MASTER_AREA = PRTITEM_SET
   m_MasterRef.KEY_NAME = txtName.Text
   m_MasterRef.KEY_CODE = txtCode.Text
   
   If Not glbMaster.AddEditMasterRef(m_MasterRef, IsOK, True, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call EnableForm(Me, True)
      Exit Function
   End If
   
   If Not IsOK Then
      Call EnableForm(Me, True)
      glbErrorLog.ShowUserError
      Call glbDaily.RollbackTransaction
      Exit Function
   End If
   
   Call EnableForm(Me, True)
   SaveData = True
End Function


Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_PrtItemSet.QueryFlag = 1
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         m_PrtItemSet.QueryFlag = 0
         Call QueryData(False)
      End If
      
      Call EnableForm(Me, True)
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
Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_PrtItemSet = Nothing
   
   Set m_PartItemGroupType = Nothing
   Set m_MasterRef = Nothing
End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitNormalLabel(lblCode, MapText("รหัสเซตข้อมูล"))
   Call InitNormalLabel(lblName, MapText("เซตข้อมูล"))
   
   Call InitTreeView
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
End Sub
Private Sub InitTreeView()
   TreeView1.Font.Name = GLB_FONT
   TreeView1.Font.Size = 14
End Sub
Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub
Private Sub Form_Load()
   OKClick = False
   Call InitFormLayout
      
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_PrtItemSet = New CPrtItemSet
   
   Set m_PartItemGroupType = New Collection
   Set m_MasterRef = New CMasterRef
End Sub

Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
   m_HasModify = True
   Call UpdateChild(Node.Child, Node.Checked)
End Sub

Private Sub txtCode_Change()
   m_HasModify = True
End Sub

Private Sub txtName_Change()
   m_HasModify = True
End Sub
Private Sub UpdateChild(ByVal Node As MSComctlLib.Node, Flag As Boolean)
Dim N As Node

   If Node Is Nothing Then
      Exit Sub
   End If
   
   Node.Checked = Flag
   Set N = Node
   While Not (N Is Nothing)
      N.Checked = Flag
      Call UpdateChild(N.Child, Flag)
      Set N = N.Next
   Wend
End Sub

