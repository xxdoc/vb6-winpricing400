VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditFormulaInput 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4965
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
   Icon            =   "frmAddEditFormulaInput.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   4395
      Left            =   0
      TabIndex        =   10
      Top             =   600
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   7752
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextLookup uctlProductLookup 
         Height          =   435
         Left            =   1830
         TabIndex        =   1
         Top             =   660
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtAmount 
         Height          =   435
         Left            =   1830
         TabIndex        =   3
         Top             =   1560
         Width           =   1515
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlTypeLookup 
         Height          =   435
         Left            =   1830
         TabIndex        =   0
         Top             =   210
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlFromFormulaLookup 
         Height          =   435
         Left            =   1830
         TabIndex        =   5
         Top             =   2460
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlLocationLookup 
         Height          =   435
         Left            =   1830
         TabIndex        =   4
         Top             =   2010
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtGroupNo 
         Height          =   435
         Left            =   1830
         TabIndex        =   6
         Top             =   2910
         Width           =   1515
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtRealAmount 
         Height          =   435
         Left            =   1830
         TabIndex        =   2
         Top             =   1110
         Width           =   1515
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin VB.Label lblRealAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblAmount"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   1140
         Width           =   1635
      End
      Begin VB.Label Label2 
         Caption         =   "lblAmount"
         Height          =   375
         Left            =   3420
         TabIndex        =   18
         Top             =   1140
         Width           =   1635
      End
      Begin VB.Label lblGroupNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblAmount"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   2940
         Width           =   1635
      End
      Begin VB.Label lblFromFormula 
         Alignment       =   1  'Right Justify
         Caption         =   "lblProduct"
         Height          =   315
         Left            =   300
         TabIndex        =   16
         Top             =   2430
         Width           =   1455
      End
      Begin VB.Label lblLocation 
         Alignment       =   1  'Right Justify
         Caption         =   "lblType"
         Height          =   315
         Left            =   300
         TabIndex        =   15
         Top             =   2010
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "lblAmount"
         Height          =   375
         Left            =   3420
         TabIndex        =   14
         Top             =   1590
         Width           =   1635
      End
      Begin VB.Label lblAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblAmount"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1590
         Width           =   1635
      End
      Begin VB.Label lblType 
         Alignment       =   1  'Right Justify
         Caption         =   "lblType"
         Height          =   315
         Left            =   300
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblProduct 
         Alignment       =   1  'Right Justify
         Caption         =   "lblProduct"
         Height          =   315
         Left            =   300
         TabIndex        =   11
         Top             =   630
         Width           =   1455
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2370
         TabIndex        =   7
         Top             =   3570
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditFormulaInput.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4020
         TabIndex        =   8
         Top             =   3570
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditFormulaInput"
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

Private m_PartTypes As Collection
Private m_PartItems As Collection
Private m_Formulas As Collection
Private m_Locations As Collection
Private m_FormulaType As Collection

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
      
   Call InitNormalLabel(lblType, MapText("ประเภท"))
   Call InitNormalLabel(lblProduct, MapText("วัตถุดิบ"))
   Call InitNormalLabel(lblAmount, MapText("จำนวนเปอร์เซ็นต์"))
   Call InitNormalLabel(Label1, MapText("%"))
   Call InitNormalLabel(lblLocation, MapText("สถานที่จัดเก็บ"))
   Call InitNormalLabel(lblFromFormula, MapText("จากสูตร"))
   Call InitNormalLabel(lblGroupNo, MapText("กลุ่ม"))
   Call InitNormalLabel(lblRealAmount, MapText("ปริมาณ"))
   Call InitNormalLabel(Label2, MapText("หน่วย"))
   
   Call txtAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtAmount.Enabled = False
   Call txtRealAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtGroupNo.SetTextLenType(TEXT_INTEGER, glbSetting.MONEY_TYPE)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)

   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
        Dim Ma As CFormulaItem
         Set Ma = TempCollection.Item(ID)
         
        uctlTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlTypeLookup.MyCombo, Ma.PART_TYPE_ID)
        uctlProductLookup.MyCombo.ListIndex = IDToListIndex(uctlProductLookup.MyCombo, Ma.PART_ITEM_ID)
        uctlLocationLookup.MyCombo.ListIndex = IDToListIndex(uctlLocationLookup.MyCombo, Ma.LOCATION_ID)
        uctlFromFormulaLookup.MyCombo.ListIndex = IDToListIndex(uctlFromFormulaLookup.MyCombo, Ma.FROM_FORMULA)
        txtAmount.Text = Ma.ITEM_PERCENT
        txtGroupNo.Text = Ma.GROUP_NO
        txtRealAmount.Text = Ma.REAL_AMOUNT
        
        End If
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

   
   If Not VerifyCombo(lblType, uctlTypeLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblProduct, uctlProductLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblRealAmount, txtRealAmount, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblGroupNo, txtGroupNo, True) Then
      Exit Function
   End If
   If Not VerifyCombo(lblLocation, uctlLocationLookup.MyCombo, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
      Dim Ma As CFormulaItem
   If ShowMode = SHOW_ADD Then
      Set Ma = New CFormulaItem
   Else
      Set Ma = TempCollection.Item(ID)
   End If
      
   Ma.PART_ITEM_ID = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
   Ma.PART_ITEM_NAME = uctlProductLookup.MyCombo.Text
   Ma.PART_NO = uctlProductLookup.MyTextBox.Text
   Ma.PART_TYPE_ID = uctlTypeLookup.MyCombo.ItemData(Minus2Zero(uctlTypeLookup.MyCombo.ListIndex))
   Ma.PART_TYPE_NAME = uctlTypeLookup.MyCombo.Text
   Ma.ITEM_PERCENT = Val(txtAmount.Text)
    Ma.LOCATION_ID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
    Ma.FROM_FORMULA = uctlFromFormulaLookup.MyCombo.ItemData(Minus2Zero(uctlFromFormulaLookup.MyCombo.ListIndex))
    Ma.GROUP_NO = Val(txtGroupNo.Text)
    Ma.REAL_AMOUNT = Val(txtRealAmount.Text)
    
   If ShowMode = SHOW_ADD Then
      Ma.Flag = "A"
         Call TempCollection.add(Ma)
      Else
      If Ma.Flag <> "A" Then
         Ma.Flag = "E"
      End If
   End If
   
   Call ReArrangeRatio(TempCollection)
   
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadPartType(uctlTypeLookup.MyCombo, m_PartTypes)
      Set uctlTypeLookup.MyCollection = m_PartTypes
            
      Call LoadLocation(uctlLocationLookup.MyCombo, m_Locations, 2)
      Set uctlLocationLookup.MyCollection = m_Locations
            
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

   Set m_PartTypes = New Collection
   Set m_PartItems = New Collection
   Set m_Formulas = New Collection
   Set m_Locations = New Collection
   Set m_FormulaType = New Collection
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If

   Set m_Rs = Nothing
   Set m_PartTypes = Nothing
   Set m_PartItems = Nothing
   Set m_Formulas = Nothing
   Set m_Locations = Nothing
   Set m_FormulaType = Nothing
End Sub

Private Sub cboType_Change()
m_HasModify = True
End Sub

Private Sub txtAmount_Change()
   m_HasModify = True
End Sub

Private Sub uctlFormulaTypeLookup_Change()
   m_HasModify = True
End Sub

Private Sub txtGroupNo_Change()
   m_HasModify = True
End Sub

Private Sub txtRealAmount_Change()
   m_HasModify = True
End Sub

Private Sub uctlFromFormulaLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlLocationLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlProductLookup_Change()
Dim ID As Long

   ID = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
   If ID <> 0 Then
      Call LoadFormula(uctlFromFormulaLookup.MyCombo, m_Formulas, , ID)
      Set uctlFromFormulaLookup.MyCollection = m_Formulas
   End If
   m_HasModify = True
End Sub

Private Sub uctlTextBox1_Change()
   m_HasModify = True
End Sub

Private Sub uctlTypeLookup_Change()
Dim PartTypeID As Long
Dim Pt As CPartType

   PartTypeID = uctlTypeLookup.MyCombo.ItemData(Minus2Zero(uctlTypeLookup.MyCombo.ListIndex))
   
   If PartTypeID > 0 Then
      Set Pt = GetPartType(m_PartTypes, Trim(Str(PartTypeID)))
      Call LoadPartItem(uctlProductLookup.MyCombo, m_PartItems, PartTypeID, "N")
      Set uctlProductLookup.MyCollection = m_PartItems
   
      Call LoadLocation(uctlLocationLookup.MyCombo, m_Locations, 2, , , Pt.PART_GROUP_ID)
      Set uctlLocationLookup.MyCollection = m_Locations
   End If
   
   m_HasModify = True
End Sub
