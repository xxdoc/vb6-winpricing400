VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditExPromotionPartItem 
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditExPromotionPartItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   4665
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   8229
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextLookup uctlPartLookup 
         Height          =   405
         Left            =   2520
         TabIndex        =   2
         Top             =   2640
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtSocCode 
         Height          =   435
         Left            =   2520
         TabIndex        =   0
         Top             =   1140
         Width           =   4005
         _ExtentX        =   11615
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlFooter 
         Height          =   705
         Left            =   0
         TabIndex        =   8
         Top             =   3960
         Width           =   12240
         _ExtentX        =   21590
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin Threed.SSCommand cmdPrev 
            Height          =   525
            Left            =   2520
            TabIndex        =   22
            Top             =   90
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   926
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmAddEditExPromotionPartItem.frx":27A2
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdNext 
            Height          =   525
            Left            =   4200
            TabIndex        =   19
            Top             =   90
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   926
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmAddEditExPromotionPartItem.frx":2ABC
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdOK 
            Height          =   525
            Left            =   8490
            TabIndex        =   4
            Top             =   90
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   926
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmAddEditExPromotionPartItem.frx":2DD6
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdExit 
            Cancel          =   -1  'True
            Height          =   525
            Left            =   10140
            TabIndex        =   5
            Top             =   90
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   926
            _Version        =   131073
            ButtonStyle     =   3
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   10
         TabIndex        =   7
         Top             =   0
         Width           =   11835
         _ExtentX        =   20876
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtDiscountRate 
         Height          =   435
         Left            =   2520
         TabIndex        =   3
         Top             =   3120
         Width           =   1695
         _ExtentX        =   6165
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPartTypeLookup 
         Height          =   405
         Left            =   2520
         TabIndex        =   1
         Top             =   2160
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextLookup uctlCustomerLookup 
         Height          =   405
         Left            =   2520
         TabIndex        =   20
         Top             =   1680
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   714
      End
      Begin Threed.SSCheck chkEditPrice 
         Height          =   345
         Left            =   7920
         TabIndex        =   23
         Top             =   3240
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   609
         _Version        =   131073
         Caption         =   "chkEditPrice"
         TripleState     =   -1  'True
      End
      Begin VB.Label lblPartLookup 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   600
         TabIndex        =   21
         Top             =   2640
         Width           =   1755
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6240
         TabIndex        =   18
         Top             =   3060
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.Label Label1 
         Height          =   315
         Left            =   10110
         TabIndex        =   17
         Top             =   3030
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label lblCustomerLookup 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   330
         TabIndex        =   16
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label lblDiscountRate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   720
         TabIndex        =   15
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Label lblBath5 
         Height          =   315
         Left            =   4350
         TabIndex        =   14
         Top             =   3930
         Width           =   705
      End
      Begin VB.Label lblBath4 
         Height          =   315
         Left            =   10050
         TabIndex        =   13
         Top             =   3900
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   6630
         TabIndex        =   12
         Top             =   300
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblBath1 
         Height          =   315
         Left            =   4320
         TabIndex        =   11
         Top             =   3120
         Width           =   1065
      End
      Begin VB.Label lblPartTypeLookup 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   600
         TabIndex        =   10
         Top             =   2160
         Width           =   1755
      End
      Begin VB.Label lblSocCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   510
         TabIndex        =   9
         Top             =   1230
         Width           =   1845
      End
   End
End
Attribute VB_Name = "frmAddEditExPromotionPartItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
'Private m_Features As Collection
'Private m_FeatureTypes As Collection
Private m_TempCol As Collection
Private EX_WORKS_PRICE_ITEM_ID As Long
Private EX_WORKS_PRICE_ID As Long
Private RATE_TYPE As Long
Private RATE_AMOUNT As Double
Private PART_ITEM_ID As Long
Public SocPartType As Long
Public SocCode As String
Public m_FeatureTypes As Collection
Public m_Customers As Collection
Private m_TempFeatures As Collection
Public PartType As Long
Public ProductType As Long
Public ParentForm As Form

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public id As Long
Public SocID As Long
Public TempCollection As Collection
Public m_ExPromotionPartItem As Collection
Public ID_MUM As Long
Private CurrentKey As String

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim D As CExPromotionPartItem

   If Flag Then
      Call EnableForm(Me, False)

      Set D = TempCollection.Item(id)
      uctlCustomerLookup.MyCombo.ListIndex = IDToListIndex(uctlCustomerLookup.MyCombo, D.CUSTOMER_ID)
      
      uctlPartTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPartTypeLookup.MyCombo, D.PART_TYPE)
      uctlPartLookup.MyCombo.ListIndex = IDToListIndex(uctlPartLookup.MyCombo, D.PART_ITEM_ID)
      txtDiscountRate.Text = Val(D.DISCOUNT_AMOUNT)
      chkEditPrice.Value = FlagToCheck(D.LAST_EDIT_FLAG)
      
      CurrentKey = Trim(str(D.CUSTOMER_ID)) & "-" & Trim(str(D.PART_ITEM_ID))

      Call EnableForm(Me, True)
   End If

   Call EnableForm(Me, True)
End Sub

Private Sub chkEnable_Click()
   m_HasModify = True
End Sub
'
Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim EPPI As CExPromotionPartItem
Dim tempEPPI As CExPromotionPartItem

   If Not VerifyCombo(lblCustomerLookup, uctlCustomerLookup.MyCombo, False) Then
      Exit Function
   End If
   
   If Not VerifyCombo(lblPartTypeLookup, uctlPartTypeLookup.MyCombo, False) Then
      Exit Function
   End If
   
  If Not VerifyCombo(lblPartLookup, uctlPartLookup.MyCombo, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblDiscountRate, txtDiscountRate, False) Then
      Exit Function
   End If

   If Val(txtDiscountRate.Text) <= 0 Then
      glbErrorLog.LocalErrorMsg = "ราคา " & lblDiscountRate.Caption & " ต้องมีค่ามากกว่า 0"
      glbErrorLog.ShowUserError
      SaveData = True
      Exit Function
   End If
   
    If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If

   Dim TempEWP As CExWorksPrice
   Dim TempEWP2 As CExWorksPrice
   Dim PartID As Long
   Dim CusID As Long
   Dim Key As String
   PartID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
   CusID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
   Key = Trim(str(CusID)) & "-" & Trim(str(PartID))
   Set TempEWP = GetObject("CExWorksPrice", m_ExPromotionPartItem, Key, False)
    If Not TempEWP Is Nothing Then
       'If (TempEWP.EX_PROMOTION_PART_ITEM_ID <> ID_MUM And TempEWP.EX_PROMOTION_PART_ITEM_ID > 0) Or ShowMode = SHOW_ADD Then 'ถ้าไม่เป็นการแก้ไขตัวเอง หรือเป็นการเพิ่มใหม่
        If (TempEWP.EX_PROMOTION_PART_ITEM_ID <> ID_MUM) Or CurrentKey <> Key Then
          glbErrorLog.LocalErrorMsg = "มีข้อมูลของลูกค้า " & uctlCustomerLookup.MyCombo.Text & " และเบอร์สินค้า " & uctlPartLookup.MyCombo.Text & " ในเอกสารชุดนี้แล้ว"
          glbErrorLog.ShowUserError
          Exit Function
       End If
    ElseIf ShowMode = SHOW_ADD Then
       Set TempEWP2 = New CExWorksPrice
       TempEWP2.Flag = "A"
       Call m_ExPromotionPartItem.add(TempEWP2, Trim(str(CusID)) & "-" & Trim(str(PartID)))
       Set TempEWP2 = Nothing
    End If
   
   If ShowMode = SHOW_ADD Then
      Set EPPI = New CExPromotionPartItem
      EPPI.CUSTOMER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
      EPPI.PART_TYPE = uctlPartTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPartTypeLookup.MyCombo.ListIndex))
      EPPI.PART_ITEM_ID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
      EPPI.PART_NO = uctlPartLookup.MyTextBox.Text
      EPPI.PART_DESC = uctlPartLookup.MyCombo.Text
      EPPI.DISCOUNT_AMOUNT = Val(txtDiscountRate.Text)
      EPPI.LAST_EDIT_FLAG = "Y" 'ถ้าเป็นการเพิ่มใหม่บังคับให้ Flag แก้ไขราคาเปิดใช้อัตโนมัติ
      EPPI.DECLARE_NEW_FLAG = "Y"
      EPPI.RATE_TYPE = 1
      EPPI.Flag = "A"
      Call TempCollection.add(EPPI)
   Else
      Set tempEPPI = TempCollection(id)
      If Check2Flag(chkEditPrice.Value) = "Y" Then 'ต้องให้กดติ๊กเลือก แก้ไขข้อมูลก่อน
         tempEPPI.CUSTOMER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
         tempEPPI.PART_TYPE = uctlPartTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPartTypeLookup.MyCombo.ListIndex))
         tempEPPI.PART_ITEM_ID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
         tempEPPI.PART_NO = uctlPartLookup.MyTextBox.Text
         tempEPPI.PART_DESC = uctlPartLookup.MyCombo.Text
         tempEPPI.DISCOUNT_AMOUNT = Val(txtDiscountRate.Text)
         
         tempEPPI.VERIFY_FLAG = "N"
         tempEPPI.VERIFY_NAME = ""
         tempEPPI.APPROVED_FLAG = "N"
         tempEPPI.APPROVED_NAME = ""
      
         tempEPPI.RATE_TYPE = 1
         tempEPPI.LAST_EDIT_FLAG = Check2Flag(chkEditPrice.Value)
       If tempEPPI.Flag <> "A" Then
         tempEPPI.Flag = "E"
       End If
      End If
   End If
   Call EnableForm(Me, True)
   SaveData = True
End Function

Private Sub chkEditPrice_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cmdNext_Click()
Dim D As CExPromotionPartItem
Dim Cm As CCustomer
Dim Pt As CPartItem
 If Not SaveData Then
      Exit Sub
   End If
If ShowMode = SHOW_EDIT Then
   id = GetNextID(id, TempCollection)
   Set D = TempCollection(id)
   uctlCustomerLookup.MyCombo.ListIndex = IDToListIndex(uctlCustomerLookup.MyCombo, D.CUSTOMER_ID)
   uctlPartTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPartTypeLookup.MyCombo, D.PART_TYPE)
   uctlPartLookup.MyCombo.ListIndex = IDToListIndex(uctlPartLookup.MyCombo, D.PART_ITEM_ID)
   txtDiscountRate.Text = Val(D.DISCOUNT_AMOUNT)
   
   ID_MUM = D.EX_PROMOTION_PART_ITEM_ID
   CurrentKey = Trim(str(D.CUSTOMER_ID)) & "-" & Trim(str(D.PART_ITEM_ID))
   chkEditPrice.Value = FlagToCheck(D.LAST_EDIT_FLAG)
Else
  id = GetNextID(id, uctlPartLookup.MyCollection)
  Set Cm = uctlCustomerLookup.MyCollection(id)
  uctlCustomerLookup.MyCombo.ListIndex = IDToListIndex(uctlCustomerLookup.MyCombo, Cm.CUSTOMER_ID)
  Set Pt = uctlPartLookup.MyCollection(id)
  uctlPartTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPartTypeLookup.MyCombo, Pt.PART_TYPE)
  uctlPartLookup.MyCombo.ListIndex = IDToListIndex(uctlPartLookup.MyCombo, Pt.PART_ITEM_ID)
   txtDiscountRate.Text = ""
   txtDiscountRate.SetFocus
   Call ParentForm.ShowGridItem
End If
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If

   OKClick = True
   Unload Me
End Sub

Private Sub cmdPrev_Click()
Dim D As CExPromotionPartItem
Dim Pt As CPartItem
 If Not SaveData Then
      Exit Sub
   End If
If ShowMode = SHOW_EDIT Then
   id = GetPrevID(id, TempCollection)
   Set D = TempCollection(id)
   uctlCustomerLookup.MyCombo.ListIndex = IDToListIndex(uctlCustomerLookup.MyCombo, D.CUSTOMER_ID)
   uctlPartTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPartTypeLookup.MyCombo, D.PART_TYPE)
   uctlPartLookup.MyCombo.ListIndex = IDToListIndex(uctlPartLookup.MyCombo, D.PART_ITEM_ID)
   txtDiscountRate.Text = Val(D.DISCOUNT_AMOUNT)
   
   ID_MUM = D.EX_PROMOTION_PART_ITEM_ID
   CurrentKey = Trim(str(D.CUSTOMER_ID)) & "-" & Trim(str(D.PART_ITEM_ID))
   chkEditPrice.Value = FlagToCheck(D.LAST_EDIT_FLAG)
Else
  id = GetPrevID(id, uctlPartLookup.MyCollection)
  Set Pt = uctlPartLookup.MyCollection(id)
  uctlCustomerLookup.MyCombo.ListIndex = IDToListIndex(uctlPartTypeLookup.MyCombo, D.CUSTOMER_ID)
  uctlPartTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPartTypeLookup.MyCombo, D.PART_TYPE)
  uctlPartLookup.MyCombo.ListIndex = IDToListIndex(uctlPartLookup.MyCombo, Pt.PART_ITEM_ID)
   txtDiscountRate.Text = ""
   txtDiscountRate.SetFocus
   Call ParentForm.ShowGridItem
End If
End Sub

'
Private Sub Form_Activate()
Dim Sp As CSystemParam
Dim FeatureTypeID As Long

   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadCustomer(uctlCustomerLookup.MyCombo, m_Customers)
      Set uctlCustomerLookup.MyCollection = m_Customers

      Call LoadPartType(uctlPartTypeLookup.MyCombo, m_FeatureTypes)
      Set uctlPartTypeLookup.MyCollection = m_FeatureTypes

      txtSocCode.Text = SocCode
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         Call QueryData(True)
      Else
         uctlPartTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPartTypeLookup.MyCombo, PartType)
'         cboRateType.ListIndex = IDToListIndex(cboRateType, RATE_FLAT)
      End If

      m_HasModify = False
   End If
End Sub
'
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.NAME
      glbErrorLog.ShowUserError
   ElseIf Shift = 1 And KeyCode = 112 Then
      If glbUser.EXCEPTION_FLAG = "Y" Then
         glbUser.EXCEPTION_FLAG = "N"
      Else
         glbUser.EXCEPTION_FLAG = "Y"
      End If
   ElseIf Shift = 0 And KeyCode = 116 Then
'      Call cmdSearch_Click
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
   ElseIf Shift = 0 And KeyCode = 118 Then
'      Call cmdAdd_Click
'   ElseIf Shift = 0 And KeyCode = 117 Then
'      Call cmdDelete_Click
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
   ElseIf Shift = 0 And KeyCode = 114 Then
'      Call cmdEdit_Click
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
   ElseIf Shift = 0 And KeyCode = 123 Then
      Call AddMemoNote
      KeyCode = 0
   End If
End Sub
'
Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_TempCol = Nothing
   Set m_FeatureTypes = Nothing
   Set m_TempFeatures = Nothing
   Set m_Customers = Nothing
End Sub

Private Sub InitFormLayout()
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   Me.KeyPreview = True

   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   pnlFooter.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)

   Call InitHeaderFooter(pnlHeader, pnlFooter)

   Call txtSocCode.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call InitNormalLabel(lblSocCode, MapText("แพคเกจ"))
   txtSocCode.Enabled = False

   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPrev.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)

   Call InitNormalLabel(lblCustomerLookup, MapText("ชื่อลูกค้า"))
   Call InitNormalLabel(lblPartTypeLookup, MapText("ประเภทสินค้า"))
   Call InitNormalLabel(lblPartLookup, MapText("รหัสสินค้า"))
   Call txtDiscountRate.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
     If ProductType = 1 Then
         Call InitNormalLabel(lblDiscountRate, MapText("ส่วนลด/ถุง"))
     ElseIf ProductType = 2 Then
         Call InitNormalLabel(lblDiscountRate, MapText("ส่วนลด/กก."))
     Else
         Call InitNormalLabel(lblDiscountRate, MapText("ส่วนลด/หน่วย"))
     End If
   Call InitNormalLabel(lblBath1, MapText("บาท"))
   
   chkEditPrice.Visible = False
   If ShowMode = SHOW_EDIT Then
      Call InitCheckBox(chkEditPrice, "ปรับปรุงข้อมูล")
      chkEditPrice.Visible = True
   End If
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdPrev, MapText("ก่อนหน้า"))
   Call InitMainButton(cmdNext, MapText("ถัดไป"))
End Sub
'
Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If

   OKClick = False
   Unload Me
End Sub
'
Private Sub Form_Load()
   OKClick = False
   Call InitFormLayout

   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_TempCol = New Collection
   Set m_FeatureTypes = New Collection
   Set m_TempFeatures = New Collection
   Set m_Customers = New Collection
End Sub

Private Sub txtPackageRate_Change()
   m_HasModify = True
End Sub

Private Sub txtDiscountRate_Change()
   m_HasModify = True
End Sub

Private Sub uctlCustomerLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlPartLookup_Change()
   m_HasModify = True
End Sub
'
Private Sub uctlPartTypeLookup_Change()
Dim PartTypeID As Long

   m_HasModify = True

'   If SocPartType = 3 Then
      PartTypeID = uctlPartTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPartTypeLookup.MyCombo.ListIndex))
      If PartTypeID > 0 Then
         Call LoadPartItem(uctlPartLookup.MyCombo, m_TempFeatures, PartTypeID)
         Set uctlPartLookup.MyCollection = m_TempFeatures
      End If
'   End If
End Sub
Private Function GetPrevID(OldID As Long, TempColl As Collection) As Long
Dim TempIndex As Long
Dim J As Long
   
   If OldID <= 1 Then
      J = 1
      glbErrorLog.LocalErrorMsg = "ถึงเรคคอร์ดแรกแล้ว"
      glbErrorLog.ShowUserError
   Else
      J = OldID - 1
   End If
   GetPrevID = J
End Function
Private Function GetNextID(OldID As Long, TempColl As Collection) As Long
Dim TempIndex As Long
Dim J As Long
   
   If OldID >= TempColl.Count Then
      J = TempColl.Count
      glbErrorLog.LocalErrorMsg = "ถึงเรคคอร์ดสุดท้ายแล้ว"
      glbErrorLog.ShowUserError
   Else
      J = OldID + 1
   End If
   
   GetNextID = J
End Function
