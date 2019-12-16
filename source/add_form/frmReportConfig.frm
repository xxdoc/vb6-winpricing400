VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmReportConfig 
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8490
   Icon            =   "frmReportConfig.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   8490
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   6060
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   10689
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboPaperSize 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1080
         Width           =   2355
      End
      Begin VB.ComboBox cboOrientation 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   4200
         Width           =   2415
      End
      Begin VB.ComboBox cboFontName 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3780
         Width           =   3525
      End
      Begin prjFarmManagement.uctlTextBox txtPaperWidth 
         Height          =   435
         Left            =   1860
         TabIndex        =   2
         Top             =   1530
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox uctlTextBox1 
         Height          =   435
         Left            =   6000
         TabIndex        =   1
         Top             =   1080
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPaperHeight 
         Height          =   435
         Left            =   6000
         TabIndex        =   3
         Top             =   1530
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtMarginBottom 
         Height          =   435
         Left            =   6000
         TabIndex        =   5
         Top             =   1980
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtMarginHeader 
         Height          =   435
         Left            =   1860
         TabIndex        =   8
         Top             =   2880
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtMarginFooter 
         Height          =   435
         Left            =   6000
         TabIndex        =   9
         Top             =   2880
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtMarginTop 
         Height          =   435
         Left            =   1860
         TabIndex        =   4
         Top             =   1980
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtMarginLeft 
         Height          =   435
         Left            =   1860
         TabIndex        =   6
         Top             =   2430
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox uctlTextBox8 
         Height          =   435
         Left            =   1860
         TabIndex        =   10
         Top             =   3360
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtMarginRight 
         Height          =   435
         Left            =   6000
         TabIndex        =   7
         Top             =   2430
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox uctlTextBox10 
         Height          =   435
         Left            =   6000
         TabIndex        =   11
         Top             =   3330
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtHeadOffset 
         Height          =   435
         Left            =   1860
         TabIndex        =   14
         Top             =   4620
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtDummyOffset 
         Height          =   435
         Left            =   6000
         TabIndex        =   15
         Top             =   4620
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtFontSize 
         Height          =   435
         Left            =   6000
         TabIndex        =   48
         Top             =   3780
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin VB.Label lblHeadOffset 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   47
         Top             =   4710
         Width           =   1575
      End
      Begin VB.Label lblDummyOffset 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4320
         TabIndex        =   46
         Top             =   4710
         Width           =   1575
      End
      Begin VB.Label Label3 
         Height          =   315
         Left            =   3480
         TabIndex        =   45
         Top             =   4650
         Width           =   525
      End
      Begin VB.Label Label1 
         Height          =   315
         Left            =   7620
         TabIndex        =   44
         Top             =   4680
         Width           =   525
      End
      Begin VB.Label lblCm10 
         Height          =   315
         Left            =   7620
         TabIndex        =   43
         Top             =   3390
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label lblCm9 
         Height          =   315
         Left            =   7620
         TabIndex        =   42
         Top             =   2970
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label lblCm8 
         Height          =   315
         Left            =   7620
         TabIndex        =   41
         Top             =   2520
         Width           =   435
      End
      Begin VB.Label lblCm7 
         Height          =   315
         Left            =   7620
         TabIndex        =   40
         Top             =   2100
         Width           =   435
      End
      Begin VB.Label lblCm6 
         Height          =   315
         Left            =   7620
         TabIndex        =   39
         Top             =   1590
         Width           =   435
      End
      Begin VB.Label lblCm5 
         Height          =   315
         Left            =   3480
         TabIndex        =   38
         Top             =   3360
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label lblCm4 
         Height          =   315
         Left            =   3480
         TabIndex        =   37
         Top             =   2940
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label lblCm3 
         Height          =   315
         Left            =   3480
         TabIndex        =   36
         Top             =   2490
         Width           =   435
      End
      Begin VB.Label lblCm2 
         Height          =   315
         Left            =   3480
         TabIndex        =   35
         Top             =   2070
         Width           =   435
      End
      Begin VB.Label lblCm1 
         Height          =   315
         Left            =   3480
         TabIndex        =   34
         Top             =   1560
         Width           =   435
      End
      Begin VB.Label lblOrientation 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   33
         Top             =   4290
         Width           =   1575
      End
      Begin VB.Label lblFontName 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   32
         Top             =   3870
         Width           =   1575
      End
      Begin VB.Label lblMarginRight 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4230
         TabIndex        =   31
         Top             =   2490
         Width           =   1665
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4320
         TabIndex        =   30
         Top             =   3420
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblMarginLeft 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   90
         TabIndex        =   29
         Top             =   2490
         Width           =   1665
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   28
         Top             =   3420
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblMarginFooter 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4230
         TabIndex        =   27
         Top             =   2940
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label lblMarginTop 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   26
         Top             =   2070
         Width           =   1575
      End
      Begin VB.Label lblMarginBottom 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4230
         TabIndex        =   25
         Top             =   2040
         Width           =   1665
      End
      Begin VB.Label lblMarginHeader 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   24
         Top             =   2970
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4230
         TabIndex        =   23
         Top             =   1140
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label lblPaperHeight 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4320
         TabIndex        =   22
         Top             =   1620
         Width           =   1575
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2625
         TabIndex        =   16
         Top             =   5280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmReportConfig.frx":27A2
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4275
         TabIndex        =   18
         Top             =   5280
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblPaperWidth 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   20
         Top             =   1620
         Width           =   1575
      End
      Begin VB.Label lblPaperSize 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   90
         TabIndex        =   19
         Top             =   1110
         Width           =   1665
      End
   End
End
Attribute VB_Name = "frmReportConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_ReportConfig As CReportConfig
Private m_Houses As Collection
Private m_Employees As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public id As Long
Public ReportKey As String
Public ReportMode As Long

Private FileName As String
Private m_SumUnit As Double
Private m_OldPartItemID As Long
Private m_PigStatus As Collection

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
      m_ReportConfig.REPORT_CONFIG_ID = id
      If Not glbDaily.QueryReportConfig(m_ReportConfig, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_ReportConfig.PopulateFromRS(1, m_Rs)
      txtPaperWidth.Text = m_ReportConfig.PAPER_WIDTH
      txtPaperHeight.Text = m_ReportConfig.PAPER_HEIGHT
      txtMarginTop.Text = m_ReportConfig.MARGIN_TOP
      txtMarginBottom.Text = m_ReportConfig.MARGIN_BOTTOM
      txtMarginLeft.Text = m_ReportConfig.MARGIN_LEFT
      txtMarginRight.Text = m_ReportConfig.MARGIN_RIGHT
      txtMarginFooter.Text = m_ReportConfig.MARGIN_FOOTER
      txtMarginHeader.Text = m_ReportConfig.MARGIN_HEADER
      cboPaperSize.ListIndex = IDToListIndex(cboPaperSize, m_ReportConfig.PAPER_SIZE)
      cboOrientation.ListIndex = IDToListIndex(cboOrientation, m_ReportConfig.ORIENTATION)
      cboFontName.ListIndex = IDToListIndex(cboFontName, m_ReportConfig.FONT_NAME)
      txtHeadOffset.Text = m_ReportConfig.HEAD_OFFSET
      txtFontSize.Text = m_ReportConfig.FONT_SIZE
      txtDummyOffset.Text = m_ReportConfig.DUMMY_OFFSET
   Else
   End If
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cboGroup_Click()
   m_HasModify = True
End Sub

Private Sub chkEnable_Click()
   m_HasModify = True
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim Pi As CPartItem
   
   If Not VerifyCombo(lblPaperSize, cboPaperSize, (ReportMode <> 1)) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblPaperWidth, txtPaperWidth, True) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblPaperHeight, txtPaperHeight, True) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblMarginTop, txtMarginTop, (ReportMode <> 1)) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblMarginBottom, txtMarginBottom, (ReportMode <> 1)) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblMarginLeft, txtMarginLeft, (ReportMode <> 1)) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblMarginRight, txtMarginRight, (ReportMode <> 1)) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblMarginHeader, txtMarginHeader, True) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblMarginFooter, txtMarginFooter, True) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_ReportConfig.AddEditMode = ShowMode
   m_ReportConfig.REPORT_KEY = ReportKey
   m_ReportConfig.REPORT_CONFIG_ID = id
   m_ReportConfig.PAPER_WIDTH = Val(txtPaperWidth.Text)
   m_ReportConfig.PAPER_HEIGHT = Val(txtPaperHeight.Text)
   m_ReportConfig.MARGIN_TOP = Val(txtMarginTop.Text)
   m_ReportConfig.MARGIN_BOTTOM = Val(txtMarginBottom.Text)
   m_ReportConfig.MARGIN_LEFT = Val(txtMarginLeft.Text)
   m_ReportConfig.MARGIN_RIGHT = Val(txtMarginRight.Text)
   m_ReportConfig.MARGIN_FOOTER = Val(txtMarginBottom.Text)
   m_ReportConfig.MARGIN_HEADER = Val(txtMarginTop.Text)
   m_ReportConfig.PAPER_SIZE = cboPaperSize.ItemData(Minus2Zero(cboPaperSize.ListIndex))
   m_ReportConfig.ORIENTATION = cboOrientation.ItemData(Minus2Zero(cboOrientation.ListIndex))
   m_ReportConfig.FONT_NAME = cboFontName.ItemData(Minus2Zero(cboFontName.ListIndex))
   m_ReportConfig.HEAD_OFFSET = Val(txtHeadOffset.Text)
   m_ReportConfig.DUMMY_OFFSET = Val(txtDummyOffset.Text)
   m_ReportConfig.FONT_SIZE = Val(txtFontSize.Text)
   m_ReportConfig.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
   
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditReportConfig(m_ReportConfig, IsOK, True, glbErrorLog) Then
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

Private Sub cboBusinessGroup_Click()
   m_HasModify = True
End Sub

Private Sub cboBusinessType_Click()
   m_HasModify = True
End Sub

Private Sub cboEnterpriseType_Click()
   m_HasModify = True
End Sub
Private Sub chkCommit_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkExtraFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cboOrientation_Click()
   m_HasModify = True
End Sub

Private Sub cboPaperSize_Click()
   m_HasModify = True
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
      Call InitOrientation(cboOrientation)
      Call InitPaperSize(cboPaperSize)
      Call InitFontName(cboFontName)
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_ReportConfig.QueryFlag = 1
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         m_ReportConfig.QueryFlag = 0
         Call QueryData(False)
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
   
   Set m_ReportConfig = Nothing
   Set m_Houses = Nothing
   Set m_Employees = Nothing
   Set m_PigStatus = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
      
   Call InitNormalLabel(lblPaperSize, MapText("��Ҵ��д��"))
   Call InitNormalLabel(lblPaperWidth, MapText("�������ҧ"))
   Call InitNormalLabel(lblPaperHeight, MapText("�����٧"))
   Call InitNormalLabel(lblMarginTop, MapText("���˹�Һ�"))
   Call InitNormalLabel(lblMarginBottom, MapText("���˹����ҧ"))
   Call InitNormalLabel(lblMarginLeft, MapText("���˹�ҫ���"))
   Call InitNormalLabel(lblMarginRight, MapText("���˹�Ң��"))
   Call InitNormalLabel(lblFontName, MapText("���Ϳ͹��"))
   Call InitNormalLabel(lblOrientation, MapText("��èѴ���§˹��"))
   Call InitNormalLabel(lblHeadOffset, MapText("��Ѻ��"))
   Call InitNormalLabel(lblDummyOffset, MapText("��Ѻ����"))
   
   Call InitNormalLabel(lblCm1, MapText("�.�."))
   Call InitNormalLabel(lblCm2, MapText("�.�."))
   Call InitNormalLabel(lblCm3, MapText("�.�."))
   Call InitNormalLabel(lblCm4, MapText("�.�."))
   Call InitNormalLabel(lblCm5, MapText("�.�."))
   Call InitNormalLabel(lblCm6, MapText("�.�."))
   Call InitNormalLabel(lblCm7, MapText("�.�."))
   Call InitNormalLabel(lblCm8, MapText("�.�."))
   Call InitNormalLabel(lblCm9, MapText("�.�."))
   Call InitNormalLabel(lblCm10, MapText("�.�."))
   Call InitNormalLabel(Label1, MapText("TWIP"))
   Call InitNormalLabel(Label3, MapText("TWIP"))
   
   Call txtPaperWidth.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   txtPaperWidth.Enabled = False
   Call txtPaperHeight.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   txtPaperHeight.Enabled = False
   Call txtMarginTop.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   txtMarginTop.Enabled = (ReportMode = 1)
   Call txtMarginBottom.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   txtMarginBottom.Enabled = (ReportMode = 1)
   Call txtMarginLeft.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   txtMarginLeft.Enabled = (ReportMode = 1)
   Call txtMarginRight.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   txtMarginRight.Enabled = (ReportMode = 1)
   Call txtMarginHeader.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   txtMarginHeader.Enabled = (ReportMode = 1)
   Call txtMarginFooter.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   txtMarginFooter.Enabled = (ReportMode = 1)
   Call txtHeadOffset.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   txtHeadOffset.Enabled = (ReportMode <> 1)
   Call txtDummyOffset.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   txtDummyOffset.Enabled = (ReportMode <> 1)
   Call txtFontSize.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   
   Call InitCombo(cboPaperSize)
   cboPaperSize.Enabled = (ReportMode = 1)
   Call InitCombo(cboFontName)
   cboFontName.Enabled = (ReportMode = 1)
   Call InitCombo(cboOrientation)
   cboOrientation.Enabled = (ReportMode = 1)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
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
   
   If ReportMode <= 0 Then
      ReportMode = 1
   End If
   Call InitFormLayout
      
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_ReportConfig = New CReportConfig
   Set m_Houses = New Collection
   Set m_Employees = New Collection
   Set m_PigStatus = New Collection
End Sub

Private Sub txtDoNo_Change()
   m_HasModify = True
End Sub

Private Sub txtParentNo_Change()
   m_HasModify = True
End Sub

Private Sub txtSellBy_Change()
   m_HasModify = True
End Sub

Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
   m_HasModify = True
End Sub

Private Sub txtPaperSize_Change()
   m_HasModify = True
End Sub

Private Sub txtPassword_Change()
   m_HasModify = True
End Sub

Private Sub txtSender_Change()
   m_HasModify = True
End Sub

Private Sub txtTotal_Change()
   m_HasModify = True
End Sub

Private Sub txtDummyOffset_Change()
   m_HasModify = True
End Sub

Private Sub txtFontSize_Change()
   m_HasModify = True
End Sub

Private Sub txtHeadOffset_Change()
   m_HasModify = True
End Sub

Private Sub txtMarginBottom_Change()
   m_HasModify = True
End Sub

Private Sub txtMarginFooter_Change()
   m_HasModify = True
End Sub

Private Sub txtMarginHeader_Change()
   m_HasModify = True
End Sub

Private Sub txtMarginLeft_Change()
   m_HasModify = True
End Sub

Private Sub txtMarginRight_Change()
   m_HasModify = True
End Sub

Private Sub txtMarginTop_Change()
   m_HasModify = True
End Sub

Private Sub txtPaperHeight_Change()
   m_HasModify = True
End Sub

Private Sub txtPaperWidth_Change()
   m_HasModify = True
End Sub

Private Sub uctlSetupDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlTextBox1_Change()
   m_HasModify = True
End Sub

Private Sub uctlDeliveryLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlDocumentDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlHouseLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlResponseByLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox10_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox8_Change()
   m_HasModify = True
End Sub
