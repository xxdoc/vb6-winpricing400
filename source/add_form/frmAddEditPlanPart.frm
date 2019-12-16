VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditPlanPart 
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7695
   Icon            =   "frmAddEditPlanPart.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   7695
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   4380
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   7726
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextLookup uctlPartItemLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   1
         Top             =   1410
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlPlanDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   0
         Top             =   990
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   -240
         TabIndex        =   10
         Top             =   0
         Width           =   8205
         _ExtentX        =   14473
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtPlanIn 
         Height          =   435
         Left            =   1860
         TabIndex        =   2
         Top             =   1860
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPlanOut 
         Height          =   435
         Left            =   1860
         TabIndex        =   3
         Top             =   2310
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtNote 
         Height          =   435
         Left            =   1860
         TabIndex        =   4
         Top             =   2760
         Width           =   5445
         _ExtentX        =   9604
         _ExtentY        =   767
      End
      Begin Threed.SSCheck chkCancelFlag 
         Height          =   375
         Left            =   6000
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   1680
         TabIndex        =   5
         Top             =   3630
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPlanPart.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label lblNote 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   0
         TabIndex        =   14
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label lblPartItem 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   13
         Top             =   1470
         Width           =   1605
      End
      Begin VB.Label lblPlanOut 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         TabIndex        =   12
         Top             =   2430
         Width           =   1695
      End
      Begin VB.Label lblPlanDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   11
         Top             =   1020
         Width           =   1575
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3315
         TabIndex        =   6
         Top             =   3630
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPlanPart.frx":2ABC
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4965
         TabIndex        =   7
         Top             =   3630
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblPlanIn 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         TabIndex        =   9
         Top             =   1980
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmAddEditPlanPart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_PlanPart As CPlanPart
Private m_PartItems As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long

Public Area As Long
Public m_IndexCollections As Collection
Public CurrentIndex As Long
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
      m_PlanPart.PLAN_PART_ID = ID
      
      If Not glbDaily.QueryPlanPart(m_PlanPart, m_Rs, itemcount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If itemcount > 0 Then
      Call m_PlanPart.PopulateFromRS(1, m_Rs)

      uctlPlanDate.ShowDate = m_PlanPart.PLAN_DATE
      uctlPartItemLookup.MyCombo.ListIndex = IDToListIndex(uctlPartItemLookup.MyCombo, m_PlanPart.PART_ITEM_ID)
      txtPlanIn.Text = m_PlanPart.PLAN_IN
      txtPlanOut.Text = m_PlanPart.PLAN_OUT
      txtNote.Text = m_PlanPart.NOTE
      chkCancelFlag.Value = FlagToCheck(m_PlanPart.CANCEL_FLAG)
   End If
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
   
   If Not VerifyDate(lblPlanDate, uctlPlanDate, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblPartItem, uctlPartItemLookup.MyCombo, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_PlanPart.AddEditMode = ShowMode
   m_PlanPart.PLAN_PART_ID = ID
   m_PlanPart.PLAN_DATE = uctlPlanDate.ShowDate
   m_PlanPart.PLAN_AREA = Area
   m_PlanPart.PART_ITEM_ID = uctlPartItemLookup.MyCombo.ItemData(Minus2Zero(uctlPartItemLookup.MyCombo.ListIndex))
   
   m_PlanPart.PLAN_IN = Val(txtPlanIn.Text)
   m_PlanPart.PLAN_OUT = Val(txtPlanOut.Text)
   m_PlanPart.NOTE = txtNote.Text
   m_PlanPart.CANCEL_FLAG = Check2Flag(chkCancelFlag.Value)
   
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditPlanPart(m_PlanPart, IsOK, True, glbErrorLog) Then
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

Private Sub chkCancelFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cmdNext_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   
   If ShowMode = SHOW_ADD Then
      Call ClearData
      uctlPlanDate.SetFocus
   Else
      If (CurrentIndex < m_IndexCollections.Count) Then
         CurrentIndex = CurrentIndex + 1
         ID = m_IndexCollections(Trim(Str(CurrentIndex)))
         Set m_PlanPart = New CPlanPart
         Call QueryData(True)
       End If
   End If
End Sub

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
      Call LoadPartItem(uctlPartItemLookup.MyCombo, m_PartItems, , , , , "N")
      Set uctlPartItemLookup.MyCollection = m_PartItems
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_PlanPart.QueryFlag = 1
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         m_PlanPart.QueryFlag = 0
         uctlPlanDate.ShowDate = Now
         Call QueryData(False)
      End If
      
      Call EnableForm(Me, True)
      m_HasModify = False
   End If
End Sub
Private Sub ClearData()
   uctlPartItemLookup.MyCombo.ListIndex = -1
   txtPlanIn.Text = ""
   txtPlanOut.Text = ""
   txtNote.Text = ""
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
   
   Set m_PlanPart = Nothing
   Set m_PartItems = Nothing
End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText

   Call InitNormalLabel(lblPlanDate, MapText("วันที่"))
   If Area = 3 Then
      Call InitNormalLabel(lblPartItem, MapText("ผลิตภัณฑ์"))
   Else
      Call InitNormalLabel(lblPartItem, MapText("วัตถุดิบ"))
   End If
   Call InitNormalLabel(lblPlanIn, MapText("ยอดรับเข้า"))
   Call InitNormalLabel(lblPlanOut, MapText("ยอดเบิกใช้"))
   Call InitNormalLabel(lblNote, MapText("หมายเหตุ"))
   
   If Area = 1 Then
      txtPlanIn.Enabled = False
   ElseIf Area = 2 Or Area = 3 Then
      txtPlanOut.Enabled = False
   End If
   
   Call InitCheckBox(chkCancelFlag, "ยกเลิก")
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdNext, MapText("ถัดไป"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("บันทึก ออก"))
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
   Set m_PlanPart = New CPlanPart
   Set m_PartItems = New Collection
End Sub

Private Sub txtNote_Change()
   m_HasModify = True
End Sub

Private Sub txtPlanIn_Change()
   m_HasModify = True
End Sub

Private Sub txtPlanOut_Change()
   m_HasModify = True
End Sub

Private Sub uctlPartItemLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlPlanDate_HasChange()
   m_HasModify = True
End Sub

