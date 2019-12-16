VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmFormulaSelect 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3810
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
   Icon            =   "frmFormulaSelect.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   3225
      Left            =   0
      TabIndex        =   7
      Top             =   600
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   5689
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextBox txtAmount 
         Height          =   435
         Left            =   1800
         TabIndex        =   3
         Top             =   1710
         Width           =   1485
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctFormulaLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   1
         Top             =   780
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlFormulaTypeLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   0
         Top             =   330
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtFormulaWeight 
         Height          =   435
         Left            =   1800
         TabIndex        =   2
         Top             =   1230
         Width           =   1485
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin VB.Label lblFormulaWeight 
         Alignment       =   1  'Right Justify
         Caption         =   "lblAmount"
         Height          =   375
         Left            =   150
         TabIndex        =   11
         Top             =   1260
         Width           =   1575
      End
      Begin VB.Label lblFormulaType 
         Alignment       =   1  'Right Justify
         Caption         =   "lblType"
         Height          =   315
         Left            =   270
         TabIndex        =   10
         Top             =   330
         Width           =   1455
      End
      Begin VB.Label lblAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblAmount"
         Height          =   375
         Left            =   480
         TabIndex        =   9
         Top             =   1740
         Width           =   1245
      End
      Begin VB.Label lblType 
         Alignment       =   1  'Right Justify
         Caption         =   "lblType"
         Height          =   315
         Left            =   270
         TabIndex        =   8
         Top             =   780
         Width           =   1455
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2370
         TabIndex        =   4
         Top             =   2400
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmFormulaSelect.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4020
         TabIndex        =   5
         Top             =   2400
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmFormulaSelect"
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
Public Job As CJob

Private m_FormulaTypes As Collection
Private m_Formulas As Collection
Private m_PartItems As Collection
Private m_Locations As Collection
Public FORMULA_ID As Long

 
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
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
      
   Call InitNormalLabel(lblType, MapText("สูตรการผลิต"))
   Call InitNormalLabel(lblAmount, MapText("จำนวน"))
   Call InitNormalLabel(lblFormulaType, MapText("ประเภทสูตร"))
   Call InitNormalLabel(lblFormulaWeight, MapText("น.น. ตามสูตร"))
   
   Call txtAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtFormulaWeight.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtFormulaWeight.Enabled = False
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim JO As CJobInput
   
   If FORMULA_ID <= 0 Then
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   
   If CountItem(Job.Outputs) Then
      
      Call LoadFormula(uctFormulaLookup.MyCombo, m_Formulas)
      Set uctFormulaLookup.MyCollection = m_Formulas
      
      Dim F As CFormula
      Set F = m_Formulas(Trim(Str(FORMULA_ID)))
      
      uctlFormulaTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlFormulaTypeLookup.MyCombo, F.FORMULA_TYPE)
      uctFormulaLookup.MyCombo.ListIndex = IDToListIndex(uctFormulaLookup.MyCombo, F.FORMULA_ID)
      For Each JO In Job.Outputs
         txtAmount.Text = JO.TX_AMOUNT
       Next JO
   End If
   
   
   Call EnableForm(Me, True)
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
Dim F As CFormula
   
   If Not VerifyCombo(lblType, uctFormulaLookup.MyCombo, False) Then
      Exit Function
   End If

   If Not VerifyTextControl(lblAmount, txtAmount, False) Then
      Exit Function
   End If
         
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Call ClearDataBefore
   
   Call CreateJobInput(Job, uctFormulaLookup.MyCombo.ItemData(Minus2Zero(uctFormulaLookup.MyCombo.ListIndex)), Val(txtAmount.Text))
   FORMULA_ID = uctFormulaLookup.MyCombo.ItemData(Minus2Zero(uctFormulaLookup.MyCombo.ListIndex))
   
   SaveData = True
End Function

Public Sub CreateJobInput(Job As CJob, FormulaID As Long, ItemAmount As Double)
Dim I As CJobInput
Dim O As CJobInput
Dim F As CFormula
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim IsOK As Boolean
Dim Fi As CFormulaItem

   Set TempRs = New ADODB.Recordset
   
   Set F = New CFormula
   
   F.FORMULA_ID = FormulaID
   F.QueryFlag = 1
   Call glbProduction.QueryFormula(F, TempRs, iCount, IsOK, glbErrorLog)
   If Not TempRs.EOF Then
      Call F.PopulateFromRS(1, TempRs)
   End If
   
   For Each Fi In F.Inputs
      Set I = New CJobInput
      I.TX_TYPE = "E"
      I.Flag = "A"
      I.PART_ITEM_ID = Fi.PART_ITEM_ID
      I.PART_NO = Fi.PART_NO
      I.PART_DESC = Fi.PART_ITEM_NAME
      I.LOCATION_ID = Fi.LOCATION_ID
      I.LOCATION_NAME = Fi.LOCATION_ID
      I.LOCATION_NO = Fi.LOCATION_NO
      I.LOCATION_NAME = Fi.LOCATION_NAME
      I.PART_TYPE_ID = Fi.PART_TYPE_ID
      I.PART_TYPE_NAME = Fi.PART_TYPE_NAME
      I.AVG_PRICE = 0
      I.FROM_FORMULA = Fi.FROM_FORMULA
      I.TX_AMOUNT = (Fi.ITEM_PERCENT / 100) * ItemAmount
      I.GROUP_NO = Fi.GROUP_NO
      I.MIX_DATE = Now
      I.MIX_DATE = DateAdd("h", 0, I.MIX_DATE)
      I.MIX_DATE = DateAdd("n", 0, I.MIX_DATE)
      I.STD_AMOUNT = I.TX_AMOUNT
      
      Call Job.Inputs.add(I)
      Set I = Nothing
   Next Fi
   
   Set O = New CJobInput
   O.Flag = "A"
   O.TX_TYPE = "I"
   O.PART_ITEM_ID = F.PART_ITEM_ID
   O.PART_NO = F.PART_NO
   O.PART_DESC = F.PART_ITEM_NAME
   O.LOCATION_ID = F.LOCATION_ID
   O.LOCATION_NAME = F.LOCATION_NAME
   O.LOCATION_NO = F.LOCATION_NO
   O.PART_TYPE_ID = F.PART_TYPE_ID
   O.PART_TYPE_NAME = F.PART_TYPE_NAME
   O.TX_AMOUNT = ItemAmount
   O.FROM_FORMULA = F.FORMULA_ID
   O.STD_AMOUNT = O.TX_AMOUNT
   Call Job.Outputs.add(O)
   Set O = Nothing
   
   Set F = Nothing
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadFormulaType(uctlFormulaTypeLookup.MyCombo, m_FormulaTypes)
      Set uctlFormulaTypeLookup.MyCollection = m_FormulaTypes
    
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

Private Sub Form_Load()

   OKClick = False
   Call InitFormLayout
   
   m_HasActivate = False
   m_HasModify = False
   Set m_Input_combo = New Collection
   Set m_Input1_combo = New Collection
   Set m_Rs = New ADODB.Recordset
   
   Set m_Formulas = New Collection
   Set m_PartItems = New Collection
   Set m_Locations = New Collection
   Set m_Formulas = New Collection
   Set m_FormulaTypes = New Collection
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_Formulas = Nothing
   Set m_PartItems = Nothing
   Set m_Locations = Nothing
   Set m_Formulas = Nothing
   Set m_FormulaTypes = Nothing
End Sub

Private Sub txtAvgPrice_Change()
   m_HasModify = True
End Sub

Private Sub uctlFromFormulaLookup_Change()
   m_HasModify = True
End Sub

Private Sub txtAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtLink_Change()
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

Private Sub uctlTextBox1_Change()

End Sub

Private Sub uctlTextLookup1_Change()
   m_HasModify = True
End Sub

Private Sub uctFormulaLookup_Change()
Dim Fm As CFormula
Dim TempID As Long

   If uctFormulaLookup.MyCollection.Count <= 0 Then
      Exit Sub
   End If
   
   TempID = uctFormulaLookup.MyCombo.ItemData(Minus2Zero(uctFormulaLookup.MyCombo.ListIndex))
   If TempID > 0 Then
      Set Fm = m_Formulas(Trim(Str(TempID)))
      txtFormulaWeight.Text = FormatNumber(Fm.SUM_REAL_AMOUNT, 3)
   Else
      txtFormulaWeight.Text = ""
   End If
End Sub

Private Sub uctlFormulaTypeLookup_Change()
Dim FormulaTypeID As Long

   FormulaTypeID = uctlFormulaTypeLookup.MyCombo.ItemData(Minus2Zero(uctlFormulaTypeLookup.MyCombo.ListIndex))
   If FormulaTypeID > 0 Then
      Call LoadFormula(uctFormulaLookup.MyCombo, m_Formulas, FormulaTypeID)
      Set uctFormulaLookup.MyCollection = m_Formulas
   End If
   
   m_HasModify = True
End Sub
Private Sub ClearDataBefore()
Dim I As Long
Dim Ji As CJobInput
Dim JO As CJobInput

   If CountItem(Job.Inputs) > 0 Or CountItem(Job.Outputs) > 0 Then
         For Each Ji In Job.Inputs
            I = I + 1
            If Ji.Flag = "A" Then
               Job.Inputs.Remove (I)
               I = I - 1
            Else
               Ji.Flag = "D"
            End If
         Next Ji
         
         I = 0
         For Each JO In Job.Outputs
            I = I + 1
            If JO.Flag = "A" Then
               Job.Outputs.Remove (I)
               I = I - 1
            Else
               JO.Flag = "D"
            End If
         Next JO
   End If

End Sub
