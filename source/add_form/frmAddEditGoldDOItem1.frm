VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditGoldDoItem1 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9435
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditGoldDOItem1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   9435
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   4935
      Left            =   0
      TabIndex        =   15
      Top             =   600
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   8705
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextBox txtQuantity 
         Height          =   435
         Left            =   1785
         TabIndex        =   3
         Top             =   1620
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPartLookup 
         Height          =   435
         Left            =   1785
         TabIndex        =   2
         Top             =   1200
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlToLocationLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   0
         Top             =   300
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPigTypeLookup 
         Height          =   435
         Left            =   1785
         TabIndex        =   1
         Top             =   750
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPrice 
         Height          =   435
         Left            =   1770
         TabIndex        =   7
         Top             =   2520
         Width           =   2025
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtWeight1 
         Height          =   435
         Left            =   1770
         TabIndex        =   5
         Top             =   2070
         Width           =   2025
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtWeight2 
         Height          =   435
         Left            =   6000
         TabIndex        =   6
         Top             =   2070
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtWage 
         Height          =   435
         Left            =   6000
         TabIndex        =   9
         Top             =   2520
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTotal 
         Height          =   435
         Left            =   1770
         TabIndex        =   10
         Top             =   2970
         Width           =   2025
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtActualWeight 
         Height          =   435
         Left            =   1770
         TabIndex        =   11
         Top             =   3420
         Width           =   2025
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin Threed.SSCheck chkGoldFlag 
         Height          =   435
         Left            =   6000
         TabIndex        =   4
         Top             =   1620
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin Threed.SSCommand cmdPriceSelect 
         Height          =   405
         Left            =   3810
         TabIndex        =   8
         Top             =   2550
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditGoldDOItem1.frx":08CA
         ButtonStyle     =   3
      End
      Begin VB.Label lblActualWeight 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   450
         TabIndex        =   28
         Top             =   3480
         Width           =   1185
      End
      Begin VB.Label Label2 
         Height          =   375
         Left            =   3870
         TabIndex        =   27
         Top             =   3450
         Width           =   1245
      End
      Begin VB.Label Label3 
         Height          =   375
         Left            =   3870
         TabIndex        =   26
         Top             =   3000
         Width           =   1245
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   450
         TabIndex        =   25
         Top             =   3030
         Width           =   1185
      End
      Begin VB.Label lblWage 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4440
         TabIndex        =   24
         Top             =   2580
         Width           =   1515
      End
      Begin VB.Label lblWeight2 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4410
         TabIndex        =   23
         Top             =   2130
         Width           =   1545
      End
      Begin VB.Label lblWeight1 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   180
         TabIndex        =   22
         Top             =   2130
         Width           =   1485
      End
      Begin VB.Label lblPrice 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   450
         TabIndex        =   21
         Top             =   2580
         Width           =   1185
      End
      Begin VB.Label lblPigType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   180
         TabIndex        =   20
         Top             =   780
         Width           =   1485
      End
      Begin VB.Label Label1 
         Height          =   345
         Left            =   7200
         TabIndex        =   19
         Top             =   1680
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblToLocation 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   18
         Top             =   360
         Width           =   1485
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3090
         TabIndex        =   12
         Top             =   4140
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditGoldDOItem1.frx":0BE4
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4740
         TabIndex        =   13
         Top             =   4140
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblPart 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   17
         Top             =   1260
         Width           =   1485
      End
      Begin VB.Label lblQuantity 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   16
         Top             =   1680
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmAddEditGoldDoItem1"
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

Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public TempCollection As Collection
Public TempCollection2 As Collection
Public COMMIT_FLAG As String

Private m_PartTypes As Collection
Private m_Parts As Collection
Private m_Locations As Collection
Private m_Houses As Collection
Private m_Pigs As Collection
Private m_PigStatuss As Collection

Private Sub cboTextType_Click()
   m_HasModify = True
End Sub

Private Sub cboCountry_Click()
   m_HasModify = True
End Sub

Private Sub chkBangkok_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkGoldFlag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

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
   
   Call InitNormalLabel(lblPart, MapText("สินค้า"))
   Call InitNormalLabel(lblQuantity, MapText("ปริมาณ"))
   Call InitNormalLabel(lblToLocation, MapText("สถานที่จัดเก็บ"))
   Call InitNormalLabel(Label1, MapText("ก.ก."))
   Call InitNormalLabel(lblPigType, MapText("ประเภทสินค้า"))
   Call InitNormalLabel(lblPrice, MapText("ราคา/บาท"))
   Call InitNormalLabel(lblWeight1, MapText("น้ำหนักบาท"))
   Call InitNormalLabel(lblWeight2, MapText("น้ำหนักกรัม"))
   Call InitNormalLabel(lblWage, MapText("ค่าแรง/หน่วย"))
   Call InitNormalLabel(lblTotal, MapText("ราคารวม"))
   Call InitNormalLabel(Label3, MapText("บาท"))
   Call InitNormalLabel(Label2, MapText("กรัม"))
   Call InitNormalLabel(lblActualWeight, MapText("น้ำหนักจริง"))
   
   Call InitCheckBox(chkGoldFlag, "เปลี่ยนเป็นเงิน")
   
   Call txtQuantity.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   Call txtWeight1.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtWeight2.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtWage.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   Call txtTotal.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   Call txtActualWeight.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPriceSelect.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdPriceSelect, MapText("..."))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Dim Di As CDoItem
         
         Set Di = TempCollection.Item(ID)
         
         uctlPigTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPigTypeLookup.MyCombo, Di.PART_TYPE)
         uctlPartLookup.MyCombo.ListIndex = IDToListIndex(uctlPartLookup.MyCombo, Di.PART_ITEM_ID)
         txtQuantity.Text = Di.ITEM_AMOUNT
         uctlToLocationLookup.MyCombo.ListIndex = IDToListIndex(uctlToLocationLookup.MyCombo, Di.LOCATION_ID)
         txtPrice.Text = Di.TOTAL_PRICE
         
         cmdOK.Enabled = (COMMIT_FLAG <> "Y")
      End If
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cmdOK_Click()
   If Not cmdOK.Enabled Then
      Exit Sub
   End If
   
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long

   If Not VerifyCombo(lblToLocation, uctlToLocationLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblPart, uctlPartLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblQuantity, txtQuantity, False) Then
      Exit Function
   End If
'   If Not VerifyTextControl(lblWeight, txtWeight, False) Then
'      Exit Function
'   End If
   If Not VerifyTextControl(lblPrice, txtPrice, False) Then
      Exit Function
   End If
      
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim Di As CDoItem
   If ShowMode = SHOW_ADD Then
      Set Di = New CDoItem
      
      Di.Flag = "A"
      Call TempCollection.add(Di)
   Else
      Set Di = TempCollection.Item(ID)
      If Di.Flag <> "A" Then
         Di.Flag = "E"
      End If
   End If

   Di.PART_ITEM_ID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
   Di.PART_NO = uctlPartLookup.MyTextBox.Text
   Di.ITEM_AMOUNT = txtQuantity.Text
   Di.LOCATION_ID = uctlToLocationLookup.MyCombo.ItemData(Minus2Zero(uctlToLocationLookup.MyCombo.ListIndex))
   Di.LOCATION_NAME = uctlToLocationLookup.MyCombo.Text
   Di.PART_TYPE_NAME = uctlPigTypeLookup.MyCombo.Text
   Di.PART_TYPE = uctlPigTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPigTypeLookup.MyCombo.ListIndex))
   Di.TOTAL_PRICE = Val(txtPrice.Text)
   Di.AVG_WEIGHT = 0
      
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadPartType(uctlPigTypeLookup.MyCombo, m_PartTypes)
      Set uctlPigTypeLookup.MyCollection = m_PartTypes
      
      Call LoadLocation(uctlToLocationLookup.MyCombo, m_Houses, 2)
      Set uctlToLocationLookup.MyCollection = m_Houses

      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
         Call QueryData(True)
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
   Set m_Rs = New ADODB.Recordset
   Set m_PartTypes = New Collection
   Set m_Parts = New Collection
   Set m_Locations = New Collection
   Set m_Houses = New Collection
   Set m_Pigs = New Collection
   Set m_PigStatuss = New Collection
   Set m_PartTypes = New Collection
End Sub

Private Sub SSCommand2_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_PartTypes = Nothing
   Set m_Parts = Nothing
   Set m_Locations = Nothing
   Set m_Houses = Nothing
   Set m_Pigs = Nothing
   Set m_PigStatuss = Nothing
   Set m_PartTypes = Nothing
End Sub

Private Sub txtDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtKeyName_Change()
   m_HasModify = True
End Sub

Private Sub txtThaiMsg_Change()
   m_HasModify = True
End Sub

Private Sub txtAmphur_Change()
   m_HasModify = True
End Sub

Private Sub txtDistrict_Change()
   m_HasModify = True
End Sub

Private Sub txtFax_Change()
   m_HasModify = True
End Sub

Private Sub txtHomeNo_Change()
   m_HasModify = True
End Sub

Private Sub txtAvgPrice_Change()
   m_HasModify = True
End Sub

Private Sub txtPhone_Change()
   m_HasModify = True
End Sub

Private Sub txtProvince_Change()
   m_HasModify = True
End Sub

Private Sub txtRoad_Change()
   m_HasModify = True
End Sub

Private Sub txtSoi_Change()
   m_HasModify = True
End Sub

Private Sub txtActualWeight_Change()
   m_HasModify = True
End Sub

Private Sub txtPrice_Change()
   m_HasModify = True
End Sub

Private Sub txtZipcode_Change()
   m_HasModify = True
End Sub

Private Sub uctlPigStatusLookup_Change()
   m_HasModify = True
End Sub

Private Sub txtTotal_Change()
   m_HasModify = True
End Sub

Private Sub txtWage_Change()
   m_HasModify = True
End Sub

Private Sub txtWeight1_Change()
   m_HasModify = True
End Sub

Private Sub txtWeight2_Change()
   m_HasModify = True
End Sub

Private Sub uctlPigTypeLookup_Change()
Dim PartTypeID As Long

   m_HasModify = True

   PartTypeID = uctlPigTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPigTypeLookup.MyCombo.ListIndex))
   If PartTypeID > 0 Then
      Call LoadPartItem(uctlPartLookup.MyCombo, m_Pigs, PartTypeID)
      Set uctlPartLookup.MyCollection = m_Pigs
   End If
End Sub

Private Sub uctlToLocationLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlLocationLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlPartLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlPartTypeLookup_Change()
Dim PartTypeID As Long
   
   Call LoadPartItem(uctlPartLookup.MyCombo, m_Parts, PartTypeID, "")
   Set uctlPartLookup.MyCollection = m_Parts
   
   m_HasModify = True
End Sub

Private Sub uctlPigWeekLookup_Change()
   m_HasModify = True
End Sub
