VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditCostItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditCostItem.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   5715
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   10081
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextBox txtRef 
         Height          =   435
         Left            =   1800
         TabIndex        =   1
         Top             =   750
         Width           =   5805
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtSerialNo 
         Height          =   435
         Left            =   1800
         TabIndex        =   0
         Top             =   300
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtCostAmount 
         Height          =   435
         Index           =   0
         Left            =   1800
         TabIndex        =   2
         Top             =   900
         Visible         =   0   'False
         Width           =   1485
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin VB.Label lblBaht 
         Caption         =   "บาท"
         Height          =   375
         Index           =   0
         Left            =   3330
         TabIndex        =   10
         Top             =   810
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label lblCostAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblAmount"
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   9
         Top             =   960
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label lblSerialNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblSerialNo"
         Height          =   375
         Left            =   60
         TabIndex        =   8
         Top             =   330
         Width           =   1665
      End
      Begin VB.Label lblRef 
         Alignment       =   1  'Right Justify
         Caption         =   "lblRef"
         Height          =   345
         Left            =   30
         TabIndex        =   7
         Top             =   780
         Width           =   1695
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2370
         TabIndex        =   3
         Top             =   4920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCostItem.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4020
         TabIndex        =   4
         Top             =   4920
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditCostItem"
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
Private m_Input_combo As Collection
Private m_Input1_combo As Collection
Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public TempCollection As Collection
Public COMMIT_FLAG As String

Private m_ProcessParams As Collection
Private m_PartItems As Collection
Private m_Locations As Collection
Private m_Formulas As Collection

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
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
      
   Call InitNormalLabel(lblSerialNo, MapText("รหัสผลิตภัณฑ์"))
   Call InitNormalLabel(lblRef, MapText("ชื่อผลิตภัณฑ์"))
   
   Call txtSerialNo.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtSerialNo.Enabled = False
   Call txtRef.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtRef.Enabled = False
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
End Sub

Private Function GetCostItem(TempCol As Collection, Ind As Long) As CCostItem
Dim Ci As CCostItem

   For Each Ci In TempCol
      If Ci.PARAM_PROCESS_ID = Ind Then
         Set GetCostItem = Ci
         Exit Function
      End If
   Next Ci
   
   Set GetCostItem = Nothing
End Function

Private Sub PopulateGui(Ma As CCostPrdItem)
Dim PP As CParameterProcess
Dim I As Long
Dim TempCi As CCostItem

   I = 0
   For Each PP In m_ProcessParams
      I = I + 1
      Set TempCi = GetCostItem(Ma.CostItems, PP.PARAMETER_PROCESS_ID)
      If TempCi Is Nothing Then
         txtCostAmount(I).Text = 0
      Else
         txtCostAmount(I).Text = TempCi.ITEM_COST
      End If
   Next PP
End Sub

Private Sub PopulateCostItem(Ma As CCostPrdItem)
Dim PP As CParameterProcess
Dim I As Long
Dim TempCi As CCostItem
Dim Tempsum As Double

   I = 0
   Tempsum = 0
   For Each PP In m_ProcessParams
      I = I + 1
      Set TempCi = GetCostItem(Ma.CostItems, PP.PARAMETER_PROCESS_ID)
      If TempCi Is Nothing Then
         Set TempCi = New CCostItem
         TempCi.Flag = "A"
         TempCi.PARAM_PROCESS_ID = PP.PARAMETER_PROCESS_ID
         TempCi.ITEM_COST = Val(txtCostAmount(I).Text)
         Call Ma.CostItems.add(TempCi)
         Set TempCi = Nothing
      Else
        TempCi.ITEM_COST = Val(txtCostAmount(I).Text)
        If TempCi.Flag <> "A" Then
            TempCi.Flag = "E"
        End If
      End If
      Tempsum = Tempsum + Val(txtCostAmount(I).Text)
   Next PP
   Ma.EXPENSE_AMOUNT = Tempsum
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Dim Ma As CCostPrdItem
         Set Ma = TempCollection.Item(ID)
         
         txtSerialNo.Text = Ma.PART_NO
         txtRef.Text = Ma.PART_DESC
         
         Call PopulateGui(Ma)
      End If
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub LoadControl(Ind As Long, TempCol As Collection)
Dim Mr As CParameterProcess
Dim I As Long
      
   I = 0
   For Each Mr In m_ProcessParams
      I = I + 1
      
      Load lblBaht(I)
      Call InitNormalLabel(lblBaht(I), "บาท")
      Load lblCostAmount(I)
      Call InitNormalLabel(lblCostAmount(I), Mr.PARAMETER_PROCESS_NAME)
      lblCostAmount(I).Visible = True
      lblCostAmount(I).Top = txtCostAmount(0).Top + txtCostAmount(0).HEIGHT * I
      lblBaht(I).Visible = True
      lblBaht(I).Top = txtCostAmount(0).Top + txtCostAmount(0).HEIGHT * I
      
      Load txtCostAmount(I)
      Call txtCostAmount(I).SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
      txtCostAmount(I).Visible = True
      txtCostAmount(I).Top = txtCostAmount(0).Top + txtCostAmount(0).HEIGHT * I
   Next Mr
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
         
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim Ma As CCostPrdItem
   If ShowMode = SHOW_ADD Then
      Set Ma = New CJobInput
   Else
      Set Ma = TempCollection.Item(ID)
   End If
   
   
   If ShowMode = SHOW_ADD Then
      Ma.Flag = "A"
      Call TempCollection.add(Ma)
   Else
      If Ma.Flag <> "A" Then
         Ma.Flag = "E"
      End If
   End If
   
   Call PopulateCostItem(Ma)
   
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadParameterProcess(Nothing, m_ProcessParams)
      Call LoadControl(1, m_ProcessParams)
                                                       
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

Private Sub Form_Load()

   OKClick = False
   Call InitFormLayout
   
   m_HasActivate = False
   m_HasModify = False
   Set m_Input_combo = New Collection
   Set m_Input1_combo = New Collection
   Set m_Rs = New ADODB.Recordset
   
   Set m_ProcessParams = New Collection
   Set m_PartItems = New Collection
   Set m_Locations = New Collection
   Set m_Formulas = New Collection
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_ProcessParams = Nothing
   Set m_PartItems = Nothing
   Set m_Locations = Nothing
   Set m_Formulas = Nothing
End Sub

Private Sub txtAvgPrice_Change()
   m_HasModify = True
End Sub

Private Sub txtGroupNo_Change()
   m_HasModify = True
End Sub

Private Sub txtStdAmount_Change()
   m_HasModify = True
End Sub

Private Sub uctlFromFormulaLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlMixDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlPartTypeLookup_Change()

End Sub

Private Sub txtAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtLink_Change()
   m_HasModify = True
End Sub

Private Sub txtCostAmount_Change(Index As Integer)
   m_HasModify = True
End Sub

Private Sub txtRef_Change()
   m_HasModify = True
End Sub

Private Sub txtSerialNo_Change()
   m_HasModify = True
End Sub

Private Sub uctlPlaceLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlProductLookup_Change()

End Sub

Private Sub uctlTextBox1_Change()

End Sub
