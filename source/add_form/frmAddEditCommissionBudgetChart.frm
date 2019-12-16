VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditCommissionBudgetChart 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9885
   Icon            =   "frmAddEditCommissionBudgetChart.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   9885
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   4245
      Left            =   0
      TabIndex        =   3
      Top             =   540
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   7488
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboParent 
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   720
         Width           =   5295
      End
      Begin prjFarmManagement.uctlTextLookup uctlSaleLookup 
         Height          =   465
         Left            =   2520
         TabIndex        =   7
         Top             =   1200
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   820
      End
      Begin VB.Label lblSale 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   480
         TabIndex        =   6
         Top             =   1320
         Width           =   1935
      End
      Begin Threed.SSCommand cmdCancel 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5040
         TabIndex        =   2
         Top             =   2280
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3360
         TabIndex        =   1
         Top             =   2280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCommissionBudgetChart.frx":08CA
         ButtonStyle     =   3
      End
      Begin VB.Label lblParent 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   780
         TabIndex        =   4
         Top             =   810
         Width           =   1605
      End
   End
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
End
Attribute VB_Name = "frmAddEditCommissionBudgetChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MODULE_NAME = "frmAddEditCommissionBudgetChart"

Private HasActivate As Boolean
Private m_HasModify As Boolean
Public HeaderText As String
Public OKClick As Boolean
Public ID As Long
Public FK_ID As Long
Public ShowMode As SHOW_MODE_TYPE
Private m_Rs As ADODB.Recordset
Public ParentID As Long

Private m_CommissionBudgetChart As CCommissionBgChart
Private m_CollBudgetCharts As Collection

Private SaleColl As Collection
Private Sub cboParent_Click()
   m_HasModify = True
End Sub

Private Sub chkNotSumComFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cmdCancel_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   OKClick = False
   Unload Me
End Sub
Private Sub cmdOK_Click()
On Error GoTo ErrorHandler
Dim IsOK As Boolean
   
   glbErrorLog.ModuleName = MODULE_NAME
   glbErrorLog.RoutineName = "Form_Activate"
   
   If Not VerifyCombo(lblParent, cboParent, True) Then
      Exit Sub
   End If
   
   If Not VerifyCombo(lblSale, uctlSaleLookup.MyCombo, False) Then
      Exit Sub
   End If
   
   If Not m_HasModify Then
      Unload Me
      Exit Sub
   End If
   
   m_CommissionBudgetChart.AddEditMode = ShowMode
   m_CommissionBudgetChart.MASTER_VALID_ID = FK_ID
   If cboParent.ListIndex >= 0 Then
      m_CommissionBudgetChart.PARENT_ID = cboParent.ItemData(cboParent.ListIndex)
   Else
      m_CommissionBudgetChart.PARENT_ID = 0
   End If

   m_CommissionBudgetChart.EMP_ID = uctlSaleLookup.MyCombo.ItemData(Minus2Zero(uctlSaleLookup.MyCombo.ListIndex))
     
   Call EnableForm(Me, False)
   
   If Not glbDaily.AddEditCommissionBudgetChart(m_CommissionBudgetChart, IsOK, True, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call EnableForm(Me, True)
      Exit Sub
   End If
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   Call EnableForm(Me, True)
   
   OKClick = True
   Unload Me
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Activate()
On Error GoTo ErrorHandler
Dim itemcount As Long
Dim IsOK As Boolean

   glbErrorLog.ModuleName = MODULE_NAME
   
   glbErrorLog.RoutineName = "Form_Load"

   If Not HasActivate Then
      HasActivate = True
      Me.Refresh

      Call LoadCommissionBudgetChart(cboParent, m_CollBudgetCharts, FK_ID)
      cboParent.ListIndex = IDToListIndex(cboParent, ParentID)
      
      Call LoadEmployee(uctlSaleLookup.MyCombo, SaleColl)
      Set uctlSaleLookup.MyCollection = SaleColl
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         Call EnableForm(Me, False)
         m_CommissionBudgetChart.COMMISSION_BUDGET_CHART_ID = ID
         m_CommissionBudgetChart.MASTER_VALID_ID = FK_ID
         If Not glbDaily.QueryCommissionBudgetChart(m_CommissionBudgetChart, m_Rs, itemcount, IsOK, glbErrorLog) Then
            glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
            Call EnableForm(Me, True)
            Exit Sub
         End If
         
         If m_Rs.RecordCount > 0 Then
            Call m_CommissionBudgetChart.PopulateFromRS(1, m_Rs)
            If m_CommissionBudgetChart.PARENT_ID > 0 Then
                  cboParent.ListIndex = IDToListIndex(cboParent, m_CommissionBudgetChart.PARENT_ID)
            End If
            
            uctlSaleLookup.MyCombo.ListIndex = IDToListIndex(uctlSaleLookup.MyCombo, m_CommissionBudgetChart.EMP_ID)
         End If
         Call EnableForm(Me, True)
         m_HasModify = False
      End If
   End If
   
Call EnableForm(Me, True)
Exit Sub
   
ErrorHandler:
   Call EnableForm(Me, True)
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      MsgBox Me.NAME
   ElseIf Shift = 1 And KeyCode = 112 Then
      If glbUser.EXCEPTION_FLAG = "Y" Then
         glbUser.EXCEPTION_FLAG = "N"
      Else
         glbUser.EXCEPTION_FLAG = "Y"
      End If
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
   End If
End Sub
Private Sub Form_Load()
   Set m_Rs = New ADODB.Recordset
   
   Set m_CommissionBudgetChart = New CCommissionBgChart
   Set SaleColl = New Collection
   Set m_CollBudgetCharts = New Collection
   
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
   
   Call InitNormalLabel(lblParent, MapText("ภายใต้"))
   Call InitNormalLabel(lblSale, MapText("พนักงานขาย"))
   
   Call InitCombo(cboParent)
   
   cmdCancel.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdCancel, MapText("ยกเลิก (ESC)"))
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set m_CommissionBudgetChart = Nothing
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set SaleColl = Nothing
   Set m_CollBudgetCharts = Nothing
End Sub

Private Sub uctlSaleLookup_Change()
   m_HasModify = True
End Sub
