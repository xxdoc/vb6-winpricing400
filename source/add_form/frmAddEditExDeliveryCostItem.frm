VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditExDeliveryCostItem 
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditExDeliveryCostItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   4545
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   8017
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboUnit2 
         Height          =   315
         Left            =   9120
         TabIndex        =   26
         Text            =   "Combo1"
         Top             =   3000
         Width           =   1455
      End
      Begin VB.ComboBox cboUnit 
         Height          =   315
         Left            =   9120
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   2520
         Width           =   1455
      End
      Begin prjFarmManagement.uctlTextLookup uctlDeliveryCusLookup 
         Height          =   405
         Left            =   2520
         TabIndex        =   2
         Top             =   2040
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtPackageCode 
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
         TabIndex        =   12
         Top             =   3840
         Width           =   12240
         _ExtentX        =   21590
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin Threed.SSCommand cmdPrev 
            Height          =   525
            Left            =   2520
            TabIndex        =   31
            Top             =   90
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   926
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmAddEditExDeliveryCostItem.frx":27A2
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdNext 
            Height          =   525
            Left            =   4200
            TabIndex        =   7
            Top             =   90
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   926
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmAddEditExDeliveryCostItem.frx":2ABC
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdOK 
            Height          =   525
            Left            =   8490
            TabIndex        =   8
            Top             =   90
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   926
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmAddEditExDeliveryCostItem.frx":2DD6
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdExit 
            Cancel          =   -1  'True
            Height          =   525
            Left            =   10140
            TabIndex        =   9
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
         TabIndex        =   11
         Top             =   0
         Width           =   11835
         _ExtentX        =   20876
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtRateDelivery 
         Height          =   435
         Left            =   2520
         TabIndex        =   3
         Top             =   2520
         Width           =   1695
         _ExtentX        =   6165
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlCustomerLookup 
         Height          =   405
         Left            =   2520
         TabIndex        =   1
         Top             =   1590
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtRateCustomer 
         Height          =   435
         Left            =   2520
         TabIndex        =   4
         Top             =   3000
         Width           =   1695
         _ExtentX        =   6165
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtWeightPerPack 
         Height          =   435
         Left            =   6480
         TabIndex        =   5
         Top             =   2520
         Width           =   1455
         _ExtentX        =   6165
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtWeightPerPackCus 
         Height          =   435
         Left            =   6480
         TabIndex        =   27
         Top             =   3000
         Width           =   1455
         _ExtentX        =   6165
         _ExtentY        =   767
      End
      Begin Threed.SSCheck chkDeclareNew 
         Height          =   345
         Left            =   9120
         TabIndex        =   32
         Top             =   2040
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   609
         _Version        =   131073
         Caption         =   "chkDeclareNew"
         TripleState     =   -1  'True
      End
      Begin Threed.SSCommand cmdDeliveryCusData 
         Height          =   405
         Left            =   8040
         TabIndex        =   30
         Top             =   2040
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditExDeliveryCostItem.frx":30F0
         ButtonStyle     =   3
      End
      Begin VB.Label lblUnit2 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   8040
         TabIndex        =   29
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label lblWeightPerPackCus 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   5040
         TabIndex        =   28
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label lblWeightPerPack 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   5040
         TabIndex        =   25
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label4 
         Height          =   315
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   1065
      End
      Begin VB.Label lblRateCustomer 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   0
         TabIndex        =   23
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Label lblBath2 
         Height          =   315
         Left            =   4350
         TabIndex        =   22
         Top             =   3000
         Width           =   945
      End
      Begin VB.Label Label1 
         Height          =   315
         Left            =   10110
         TabIndex        =   21
         Top             =   3030
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label lblCustomerLookup 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   330
         TabIndex        =   20
         Top             =   1590
         Width           =   2055
      End
      Begin VB.Label lblRateDelivery 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   0
         TabIndex        =   19
         Top             =   2520
         Width           =   2415
      End
      Begin VB.Label lblBath5 
         Height          =   315
         Left            =   4350
         TabIndex        =   18
         Top             =   3930
         Width           =   705
      End
      Begin VB.Label lblBath4 
         Height          =   315
         Left            =   10050
         TabIndex        =   17
         Top             =   3900
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   8040
         TabIndex        =   16
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label lblBath1 
         Height          =   315
         Left            =   4350
         TabIndex        =   15
         Top             =   2520
         Width           =   585
      End
      Begin VB.Label lblDeliveryCusLookup 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         TabIndex        =   14
         Top             =   2040
         Width           =   2355
      End
      Begin VB.Label lblPackageCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   510
         TabIndex        =   13
         Top             =   1140
         Width           =   1845
      End
   End
End
Attribute VB_Name = "frmAddEditExDeliveryCostItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_TempCol As Collection
Private EX_DELIVERY_COST_ITEM_ID As Long
Private EX_WORKS_PRICE_ID As Long
Private RATE_TYPE As Long
Private RATE_TYPE_CUS As Long
Private RATE_DELIVERY As Double
Private RATE_CUSTOMER As Double
Public PackageCode As String
Private CUSTOMER_ID As Long
Public m_Customers As Collection
Private m_DeliveryCus As Collection
Public UnitType As Long
Public UnitTypeCus As Long
Public ParentForm As Form


Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public id As Long
Public SocID As Long
Public TempCollection As Collection
Public m_ExDeliveryCostItem As Collection
Public ID_MUM As Long
Private CurrentKey As String

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim D As CExDeliveryCostItem

   If Flag Then
      Call EnableForm(Me, False)

      Set D = TempCollection.Item(id)
      uctlCustomerLookup.MyCombo.ListIndex = IDToListIndex(uctlCustomerLookup.MyCombo, D.CUSTOMER_ID)
      uctlDeliveryCusLookup.MyCombo.ListIndex = IDToListIndex(uctlDeliveryCusLookup.MyCombo, D.DELIVERY_CUS_ITEM_ID)
     txtRateDelivery.Text = Val(D.RATE_DELIVERY)
     txtRateCustomer.Text = Val(D.RATE_CUSTOMER)
     txtWeightPerPack.Text = Val(D.WEIGHT_PER_PACK)
     txtWeightPerPackCus.Text = Val(D.WEIGHT_PER_PACK_CUS)
     cboUnit.ListIndex = IDToListIndex(cboUnit, D.RATE_TYPE)
     cboUnit2.ListIndex = IDToListIndex(cboUnit2, D.RATE_TYPE_CUS)
     chkDeclareNew.Value = FlagToCheck(D.DECLARE_NEW_FLAG)
     CurrentKey = Trim(str(D.CUSTOMER_ID)) & "-" & Trim(str(D.DELIVERY_CUS_ITEM_ID)) & "-" & Trim(str(D.RATE_TYPE)) & "-" & Trim(str(D.RATE_TYPE_CUS))

      Call EnableForm(Me, True)
   End If

   Call EnableForm(Me, True)
End Sub
'
Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim EDC As CExDeliveryCostItem
Dim TempEDC  As CExDeliveryCostItem

   If Not VerifyCombo(lblCustomerLookup, uctlCustomerLookup.MyCombo, False) Then
      Exit Function
   End If
   
   If Not VerifyCombo(lblDeliveryCusLookup, uctlDeliveryCusLookup.MyCombo, False) Then
      Exit Function
   End If
   
   If Not VerifyCombo(lblUnit, cboUnit, False) Then
      Exit Function
   End If
   
   If Not VerifyCombo(lblUnit2, cboUnit2, False) Then
      Exit Function
   End If
   
  If Not VerifyTextControl(lblRateDelivery, txtRateDelivery, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblRateCustomer, txtRateCustomer, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblWeightPerPack, txtWeightPerPack, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblWeightPerPackCus, txtWeightPerPackCus, False) Then
      Exit Function
   End If
   
   If Val(txtRateDelivery.Text) <= 0 Then
      glbErrorLog.LocalErrorMsg = "ราคา " & lblRateDelivery.Caption & " ต้องมีค่ามากกว่า 0"
      glbErrorLog.ShowUserError
      SaveData = True
      Exit Function
   End If
   
   If Val(txtRateCustomer.Text) <= 0 Then
      glbErrorLog.LocalErrorMsg = "ราคา " & lblRateCustomer.Caption & " ต้องมีค่ามากกว่า 0"
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
   Dim CusID As Long
   Dim DelCusItemId As Long
   Dim RateType As Long
   Dim RateTypeCus As Long
   Dim Key As String
   
     CusID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
     DelCusItemId = uctlDeliveryCusLookup.MyCombo.ItemData(Minus2Zero(uctlDeliveryCusLookup.MyCombo.ListIndex))
     RateType = cboUnit.ItemData(Minus2Zero(cboUnit.ListIndex))
     RateTypeCus = cboUnit2.ItemData(Minus2Zero(cboUnit2.ListIndex))
     Key = Trim(str(CusID)) & "-" & Trim(str(DelCusItemId)) & "-" & Trim(str(RateType)) & "-" & Trim(str(RateTypeCus))
     Set TempEWP = GetObject("CExWorksPrice", m_ExDeliveryCostItem, Key, False)
      If Not TempEWP Is Nothing Then
         If (TempEWP.EX_DELIVERY_COST_ITEM_ID <> ID_MUM) Or CurrentKey <> Key Then
            glbErrorLog.LocalErrorMsg = "มีข้อมูลของสถานที่จัดส่ง " & uctlDeliveryCusLookup.MyCombo.Text & " ในเอกสารชุดนี้แล้ว"
            glbErrorLog.ShowUserError
            Exit Function
         End If
      Else
         Set TempEWP2 = New CExWorksPrice
         TempEWP2.Flag = "A"
         Call m_ExDeliveryCostItem.add(TempEWP2, Key)
         Set TempEWP2 = Nothing
      End If


   If ShowMode = SHOW_ADD Then
      Set EDC = New CExDeliveryCostItem
      EDC.CUSTOMER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
      EDC.CUSTOMER_CODE = uctlCustomerLookup.MyTextBox.Text
      EDC.CUSTOMER_NAME = uctlCustomerLookup.MyCombo.Text
      EDC.DELIVERY_CUS_ITEM_ID = uctlDeliveryCusLookup.MyCombo.ItemData(Minus2Zero(uctlDeliveryCusLookup.MyCombo.ListIndex))
      EDC.DELIVERY_CUS_ITEM_CODE = uctlDeliveryCusLookup.MyTextBox.Text
      EDC.DELIVERY_CUS_ITEM_NAME = uctlDeliveryCusLookup.MyCombo.Text
      EDC.RATE_DELIVERY = Val(txtRateDelivery.Text)
      EDC.RATE_CUSTOMER = Val(txtRateCustomer.Text)
      EDC.WEIGHT_PER_PACK = Val(txtWeightPerPack.Text)
      EDC.WEIGHT_PER_PACK_CUS = Val(txtWeightPerPackCus.Text)
      EDC.RATE_TYPE = cboUnit.ItemData(Minus2Zero(cboUnit.ListIndex))
      EDC.RATE_TYPE_CUS = cboUnit2.ItemData(Minus2Zero(cboUnit2.ListIndex))
      EDC.LAST_EDIT_FLAG = "Y" 'ถ้าเป็นการเพิ่มใหม่บังคับให้ Flag แก้ไขราคาเปิดใช้อัตโนมัติ
      EDC.DECLARE_NEW_FLAG = "Y"
      EDC.Flag = "A"
      Call TempCollection.add(EDC)
   Else
      Set EDC = TempCollection(id)
      If Check2Flag(chkDeclareNew.Value) = "Y" Then 'เข้าแก้ไขได้ต่อเมื่อ ยังไม่เคยประกาศราคามาก่อนเท่านั้น
         EDC.CUSTOMER_CODE = uctlCustomerLookup.MyTextBox.Text
         EDC.CUSTOMER_NAME = uctlCustomerLookup.MyCombo.Text
         EDC.DELIVERY_CUS_ITEM_ID = uctlDeliveryCusLookup.MyCombo.ItemData(Minus2Zero(uctlDeliveryCusLookup.MyCombo.ListIndex))
         EDC.DELIVERY_CUS_ITEM_CODE = uctlDeliveryCusLookup.MyTextBox.Text
         EDC.DELIVERY_CUS_ITEM_NAME = uctlDeliveryCusLookup.MyCombo.Text
         EDC.RATE_DELIVERY = Val(txtRateDelivery.Text)
         EDC.RATE_CUSTOMER = Val(txtRateCustomer.Text)
         EDC.WEIGHT_PER_PACK = Val(txtWeightPerPack.Text)
         EDC.WEIGHT_PER_PACK_CUS = Val(txtWeightPerPackCus.Text)
         EDC.RATE_TYPE = cboUnit.ItemData(Minus2Zero(cboUnit.ListIndex))
         EDC.RATE_TYPE_CUS = cboUnit2.ItemData(Minus2Zero(cboUnit2.ListIndex))
         
         EDC.VERIFY_FLAG = "N"
         EDC.VERIFY_NAME = ""
         EDC.APPROVED_FLAG = "N"
         EDC.APPROVED_NAME = ""
         EDC.LAST_EDIT_FLAG = "Y"
      Else
         EDC.LAST_EDIT_FLAG = "N"
      End If
      EDC.DECLARE_NEW_FLAG = Check2Flag(chkDeclareNew.Value)
      If EDC.Flag <> "A" Then
         EDC.Flag = "E"
      End If
   End If

   Call EnableForm(Me, True)
   SaveData = True
End Function

Private Sub cboUnit_Change()
   m_HasModify = True
End Sub

Private Sub cboUnit_Click()
   m_HasModify = True
   If cboUnit.ListIndex = 1 Then
      txtWeightPerPack.Enabled = True
   ElseIf cboUnit.ListIndex = 2 Then
      txtWeightPerPack.Enabled = False
      txtWeightPerPack.Text = "1"
   ElseIf cboUnit.ListIndex = 3 Then
      txtWeightPerPack.Enabled = False
      txtWeightPerPack.Text = "999"
   End If
End Sub

Private Sub cboUnit_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cboUnit2_Change()
   m_HasModify = True
End Sub

Private Sub cboUnit2_Click()
   m_HasModify = True
  If cboUnit2.ListIndex = 1 Then
      txtWeightPerPackCus.Enabled = True
   ElseIf cboUnit2.ListIndex = 2 Then
      txtWeightPerPackCus.Enabled = False
      txtWeightPerPackCus.Text = "1"
   ElseIf cboUnit2.ListIndex = 3 Then
      txtWeightPerPackCus.Enabled = False
      txtWeightPerPackCus.Text = "999"
   End If
End Sub

Private Sub chkEditPrice_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkDeclareNew_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cmdDeliveryCusData_Click()
      frmAddEditDeliveryCusMain.HeaderText = MapText("ข้อมูลสถานที่จัดส่ง")
      frmAddEditDeliveryCusMain.CustomerID = CUSTOMER_ID
      Load frmAddEditDeliveryCusMain
      frmAddEditDeliveryCusMain.Show 1

      OKClick = frmAddEditDeliveryCusMain.OKClick

      Unload frmAddEditDeliveryCusMain
      Set frmAddEditDeliveryCusMain = Nothing

End Sub

Private Sub cmdNext_Click()
Dim D As CExDeliveryCostItem
Dim DC As CDeliveryCus
 If Not SaveData Then
      Exit Sub
   End If
If ShowMode = SHOW_EDIT Then
   id = GetNextID(id, TempCollection)
   Set D = TempCollection(id)
      uctlCustomerLookup.MyCombo.ListIndex = IDToListIndex(uctlCustomerLookup.MyCombo, D.CUSTOMER_ID)
      uctlDeliveryCusLookup.MyCombo.ListIndex = IDToListIndex(uctlDeliveryCusLookup.MyCombo, D.DELIVERY_CUS_ITEM_ID)
     txtRateDelivery.Text = Val(D.RATE_DELIVERY)
     txtRateCustomer.Text = Val(D.RATE_CUSTOMER)
     txtWeightPerPack.Text = Val(D.WEIGHT_PER_PACK)
     txtWeightPerPackCus.Text = Val(D.WEIGHT_PER_PACK_CUS)
     cboUnit.ListIndex = IDToListIndex(cboUnit, D.RATE_TYPE)
     cboUnit2.ListIndex = IDToListIndex(cboUnit2, D.RATE_TYPE_CUS)
     
     ID_MUM = D.EX_DELIVERY_COST_ITEM_ID
     CurrentKey = Trim(str(D.CUSTOMER_ID)) & "-" & Trim(str(D.DELIVERY_CUS_ITEM_ID)) & "-" & Trim(str(D.RATE_TYPE)) & "-" & Trim(str(D.RATE_TYPE_CUS))
     chkDeclareNew.Value = FlagToCheck(D.DECLARE_NEW_FLAG)
Else
  id = GetNextID(id, uctlDeliveryCusLookup.MyCollection)
  Set DC = uctlDeliveryCusLookup.MyCollection(id)
  uctlDeliveryCusLookup.MyCombo.ListIndex = IDToListIndex(uctlDeliveryCusLookup.MyCombo, DC.DELIVERY_CUS_ITEM_ID)
   txtRateDelivery.Text = ""
   txtRateCustomer.Text = ""
   txtWeightPerPack.Text = ""
   txtWeightPerPackCus.Text = ""
   cboUnit.ListIndex = -1
   cboUnit2.ListIndex = -1
   txtRateDelivery.SetFocus
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
Dim D As CExDeliveryCostItem
Dim DC As CDeliveryCus
 If Not SaveData Then
      Exit Sub
   End If
If ShowMode = SHOW_EDIT Then
   id = GetPrevID(id, TempCollection)
   Set D = TempCollection(id)
      uctlCustomerLookup.MyCombo.ListIndex = IDToListIndex(uctlCustomerLookup.MyCombo, D.CUSTOMER_ID)
      uctlDeliveryCusLookup.MyCombo.ListIndex = IDToListIndex(uctlDeliveryCusLookup.MyCombo, D.DELIVERY_CUS_ITEM_ID)
     txtRateDelivery.Text = Val(D.RATE_DELIVERY)
     txtRateCustomer.Text = Val(D.RATE_CUSTOMER)
     txtWeightPerPack.Text = Val(D.WEIGHT_PER_PACK)
     txtWeightPerPackCus.Text = Val(D.WEIGHT_PER_PACK_CUS)
     cboUnit.ListIndex = IDToListIndex(cboUnit, D.RATE_TYPE)
     cboUnit2.ListIndex = IDToListIndex(cboUnit2, D.RATE_TYPE_CUS)
     
     ID_MUM = D.EX_DELIVERY_COST_ITEM_ID
     CurrentKey = Trim(str(D.CUSTOMER_ID)) & "-" & Trim(str(D.DELIVERY_CUS_ITEM_ID)) & "-" & Trim(str(D.RATE_TYPE)) & "-" & Trim(str(D.RATE_TYPE_CUS))
     chkDeclareNew.Value = FlagToCheck(D.DECLARE_NEW_FLAG)
Else
  id = GetPrevID(id, uctlDeliveryCusLookup.MyCollection)
  Set DC = uctlDeliveryCusLookup.MyCollection(id)
  uctlDeliveryCusLookup.MyCombo.ListIndex = IDToListIndex(uctlDeliveryCusLookup.MyCombo, DC.DELIVERY_CUS_ITEM_ID)
   txtRateDelivery.Text = ""
   txtRateCustomer.Text = ""
   txtWeightPerPack.Text = ""
   txtWeightPerPackCus.Text = ""
   cboUnit.ListIndex = -1
   cboUnit2.ListIndex = -1
   txtRateDelivery.SetFocus
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
      
      Call InitDeliveryType(cboUnit)
      cboUnit.ListIndex = UnitType
      
      Call InitDeliveryType(cboUnit2)
      cboUnit2.ListIndex = UnitTypeCus
  

      txtPackageCode.Text = PackageCode
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         Call QueryData(True)
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
   Set m_Customers = Nothing
   Set m_DeliveryCus = Nothing
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

   Call txtPackageCode.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call InitNormalLabel(lblPackageCode, MapText("แพคเกจ"))
   txtPackageCode.Enabled = False

   Call InitNormalLabel(lblCustomerLookup, MapText("ลูกค้า"))
   Call InitNormalLabel(lblDeliveryCusLookup, MapText("สถานที่จัดส่ง"))
   Call txtRateDelivery.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtRateCustomer.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtWeightPerPack.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtWeightPerPackCus.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtWeightPerPack.Enabled = True
   txtWeightPerPack.Text = ""
   
     If UnitType = 1 Then
         Call InitNormalLabel(lblRateDelivery, MapText("จ่ายให้รถรับจ้าง/ถุง"))
         Call InitNormalLabel(lblRateCustomer, MapText("ค่าขนส่งคิดลูกค้า/ถุง"))
     ElseIf UnitType = 2 Then
         Call InitNormalLabel(lblRateDelivery, MapText("จ่ายให้รถรับจ้าง/กก."))
         Call InitNormalLabel(lblRateCustomer, MapText("ค่าขนส่งคิดลูกค้า/กก."))
         
         txtWeightPerPack.Enabled = False
         txtWeightPerPack.Text = "1"
         
   ElseIf UnitType = 3 Then
         Call InitNormalLabel(lblRateDelivery, MapText("จ่ายให้รถรับจ้าง/เที่ยว"))
         Call InitNormalLabel(lblRateCustomer, MapText("ค่าขนส่งคิดลูกค้า/เที่ยว"))
         
         txtWeightPerPack.Enabled = False
         txtWeightPerPack.Text = "30"
     Else
         Call InitNormalLabel(lblRateDelivery, MapText("จ่ายให้รถรับจ้าง/หน่วย"))
         Call InitNormalLabel(lblRateCustomer, MapText("ค่าขนส่งคิดลูกค้า/หน่วย"))
     End If
     
   txtWeightPerPackCus.Enabled = True
   txtWeightPerPackCus.Text = ""
   
   If UnitTypeCus = 1 Then
         Call InitNormalLabel(lblRateDelivery, MapText("จ่ายให้รถรับจ้าง/ถุง"))
         Call InitNormalLabel(lblRateCustomer, MapText("ค่าขนส่งคิดลูกค้า/ถุง"))
     ElseIf UnitTypeCus = 2 Then
         Call InitNormalLabel(lblRateDelivery, MapText("จ่ายให้รถรับจ้าง/กก."))
         Call InitNormalLabel(lblRateCustomer, MapText("ค่าขนส่งคิดลูกค้า/กก."))
         
         txtWeightPerPackCus.Enabled = False
         txtWeightPerPackCus.Text = "1"
         
   ElseIf UnitTypeCus = 3 Then
         Call InitNormalLabel(lblRateDelivery, MapText("จ่ายให้รถรับจ้าง/เที่ยว"))
         Call InitNormalLabel(lblRateCustomer, MapText("ค่าขนส่งคิดลูกค้า/เที่ยว"))
         
         txtWeightPerPackCus.Enabled = False
         txtWeightPerPackCus.Text = "30"
     Else
         Call InitNormalLabel(lblRateDelivery, MapText("จ่ายให้รถรับจ้าง/หน่วย"))
         Call InitNormalLabel(lblRateCustomer, MapText("ค่าขนส่งคิดลูกค้า/หน่วย"))
     End If

   Call InitNormalLabel(lblBath1, MapText("บาท"))
   Call InitNormalLabel(lblBath2, MapText("บาท"))
   Call InitNormalLabel(lblWeightPerPack, MapText("น้ำหนัก(กก.)"))
   Call InitNormalLabel(lblWeightPerPackCus, MapText("น้ำหนัก(กก.)"))
   Call InitNormalLabel(lblUnit, MapText("หน่วย"))
   Call InitNormalLabel(lblUnit2, MapText("หน่วย"))
   
   chkDeclareNew.Visible = False
   If ShowMode = SHOW_EDIT Then
      Call InitCheckBox(chkDeclareNew, "ประกาศราคาใหม่")
      chkDeclareNew.Visible = True
   End If
   
   Call InitCombo(cboUnit)
   Call InitCombo(cboUnit2)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPrev.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDeliveryCusData.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdPrev, MapText("ก่อนหน้า"))
   Call InitMainButton(cmdNext, MapText("ถัดไป"))
   Call InitMainButton(cmdDeliveryCusData, MapText("เพิ่ม"))

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
   Set m_Customers = New Collection
   Set m_DeliveryCus = New Collection
End Sub

Private Sub txtPackageRate_Change()
   m_HasModify = True
End Sub

Private Sub uctlPartLookup_Change()
   m_HasModify = True
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

Private Sub txtRateCustomer_Change()
   m_HasModify = True
End Sub

Private Sub txtRateDelivery_Change()
   m_HasModify = True
End Sub

Private Sub txtWeightPerPack_Change()
   m_HasModify = True
End Sub

Private Sub txtWeightPerPackCus_Change()
   m_HasModify = True
End Sub

Private Sub uctlCustomerLookup_Change()

   m_HasModify = True
   
   CUSTOMER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
   If CUSTOMER_ID > 0 Then
      Call LoadDeliveryCus(uctlDeliveryCusLookup.MyCombo, m_DeliveryCus, CUSTOMER_ID)
      Set uctlDeliveryCusLookup.MyCollection = m_DeliveryCus
   End If

End Sub

Private Sub uctlDeliveryCusLookup_Change()
   m_HasModify = True
End Sub
