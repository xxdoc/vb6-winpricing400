VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAddEditParameter 
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8490
   Icon            =   "frmAddEditParameter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
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
               Picture         =   "frmAddEditParameter.frx":27A2
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   2835
         Left            =   180
         TabIndex        =   2
         Top             =   2010
         Width           =   8145
         _ExtentX        =   14367
         _ExtentY        =   5001
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "ImageList1"
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin prjFarmManagement.uctlTextBox txtDocumentNo 
         Height          =   435
         Left            =   2280
         TabIndex        =   0
         Top             =   930
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtMotherNo 
         Height          =   435
         Left            =   2280
         TabIndex        =   1
         Top             =   1380
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
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
         MouseIcon       =   "frmAddEditParameter.frx":307C
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
      Begin VB.Label lblMotherNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   150
         TabIndex        =   7
         Top             =   1470
         Width           =   2025
      End
      Begin VB.Label lblDocumentNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   90
         TabIndex        =   6
         Top             =   990
         Width           =   2115
      End
   End
End
Attribute VB_Name = "frmAddEditParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Process As CProcess
'Private m_Houses As Collection
'Private m_Employees As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long

Private FileName As String
'Private m_SumUnit As Double
'Private m_OldPartItemID As Long
Private m_Locations As Collection

Private Sub LoadTreeView(Col As Collection)
Dim C As CParameterItem
Dim N As Node
Dim Np As Node

      For Each C In Col
         Set N = TreeView1.Nodes.add(, tvwFirst, Trim(str(C.HGI_PARAMETER_ID)) & "-X", C.PARAMETER_PROCESS_NAME & " (" & C.PARAMETER_PROCESS_NO & ")", 1, 1)
         N.Tag = C.PARAMETER_ID
         N.Checked = (C.SELECT_FLAG = "Y")

         N.Expanded = False
      Next C
      
      Dim check As Long
      Call LoadParameterProcess(Nothing, m_Locations)
        Dim AA As CParameterProcess
        For Each AA In m_Locations
       For Each C In Col
       If AA.PARAMETER_PROCESS_ID = C.PARAMETER_ID Then
       check = -1
       Exit For
       Else
       check = AA.PARAMETER_PROCESS_ID
       End If
       Next C

       If check > -1 Then
         Set N = TreeView1.Nodes.add(, tvwFirst, Trim(str(AA.PARAMETER_PROCESS_ID)) & "-X", AA.PARAMETER_PROCESS_NAME & " (" & AA.PARAMETER_PROCESS_NO & ")", 1, 1)
         N.Tag = check
         N.Checked = False

         N.Expanded = False

       End If
       Next AA

      
End Sub

Private Sub LoadLocationTreeView(Col As Collection)
Dim C As CParameterProcess
Dim N As Node
Dim Np As Node
 
       
      For Each C In Col
         Set N = TreeView1.Nodes.add(, tvwFirst, Trim(str(C.PARAMETER_PROCESS_ID)) & "-X", C.PARAMETER_PROCESS_NAME & " (" & C.PARAMETER_PROCESS_NO & ")", 1, 1)
         N.Tag = C.PARAMETER_PROCESS_ID
         N.Checked = False
         
         N.Expanded = False
      Next C
      
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
      m_Process.PROCESS_ID = ID
      If Not glbMaster.QueryProcess(m_Process, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_Process.PopulateFromRS(1, m_Rs)
      txtDocumentNo.Text = m_Process.PROCESS_NO
      txtMotherNo.Text = m_Process.PROCESS_NAME
      
      Dim II As CParameterItem
      If m_Process.ParameterItems.Count > 0 Then
         Call LoadTreeView(m_Process.ParameterItems)
    
      End If
   Else
     ' Call LoadTreeView(m_Process.ParameterItems)
   End If
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub


Private Sub PopulateItem(Col As Collection)
Dim N As Node
Dim C As CParameterItem

   For Each C In Col
      C.Flag = "D"
   Next C
   
   For Each N In TreeView1.Nodes
      Set C = New CParameterItem
      C.Flag = "A"
      C.PARAMETER_ID = Val(N.Key)
      If N.Checked Then
         C.SELECT_FLAG = "Y"
      Else
         C.SELECT_FLAG = "N"
      End If
      Call Col.add(C)
      Set C = Nothing
   Next N
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim Pi As CPartItem

   If Not VerifyTextControl(lblDocumentNo, txtDocumentNo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblMotherNo, txtMotherNo, False) Then
      Exit Function
   End If
   
'   If Not CheckUniqueNs(IMPORT_UNIQUE, txtDocumentNo.Text, ID) Then
'      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtDocumentNo.Text & " " & MapText("อยู่ในระบบแล้ว")
'      glbErrorLog.ShowUserError
'      Exit Function
'   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_Process.AddEditMode = ShowMode
   m_Process.PROCESS_ID = ID
   m_Process.PROCESS_NO = txtDocumentNo.Text
   m_Process.PROCESS_NAME = txtMotherNo.Text
   Call PopulateItem(m_Process.ParameterItems)
   
   Call EnableForm(Me, False)
   If Not glbMaster.AddEditProcess(m_Process, IsOK, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call EnableForm(Me, True)
      Exit Function
   End If
   If Not IsOK Then
      Call EnableForm(Me, True)
      glbErrorLog.ShowUserError
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
         m_Process.QueryFlag = 1
'         Call LoadParameterProcess(Nothing, m_Locations)
'         Call LoadLocationTreeView(m_Locations)

         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         m_Process.QueryFlag = 0
         Call QueryData(False)
         Call LoadParameterProcess(Nothing, m_Locations)
         Call LoadLocationTreeView(m_Locations)
      End If
      
      Call EnableForm(Me, True)
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

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_Process = Nothing
'   Set m_Houses = Nothing
'   Set m_Employees = Nothing
   Set m_Locations = Nothing
End Sub



Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblMotherNo, MapText("ชื่อโปรเซส"))
   Call InitNormalLabel(lblDocumentNo, MapText("หมายเลขโปรเซส"))
   
   Call txtDocumentNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtMotherNo.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitTreeView
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
End Sub

Private Sub InitTreeView()
   TreeView1.Font.NAME = GLB_FONT
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
   Set m_Process = New CProcess
'   Set m_Houses = New Collection
'   Set m_Employees = New Collection
   Set m_Locations = New Collection
End Sub


Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
   m_HasModify = True
End Sub

Private Sub txtDocumentNo_Change()
   m_HasModify = True
End Sub


Private Sub txtMotherNo_Change()
   m_HasModify = True
End Sub

