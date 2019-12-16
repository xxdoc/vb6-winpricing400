VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditTransportItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditTransportItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   8040
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   4875
      Left            =   0
      TabIndex        =   12
      Top             =   600
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   8599
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSFrame SSFrame2 
         Height          =   975
         Left            =   4080
         TabIndex        =   29
         Top             =   2400
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   1720
         _Version        =   131073
         Begin Threed.SSOption ssoCalCustomer 
            Height          =   450
            Left            =   120
            TabIndex        =   8
            Top             =   480
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   794
            _Version        =   131073
            Caption         =   "ssoCalCustomer"
         End
         Begin Threed.SSOption ssoCalDriver 
            Height          =   450
            Left            =   120
            TabIndex        =   7
            Top             =   120
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   794
            _Version        =   131073
            Caption         =   "ssoCalDriver"
         End
      End
      Begin prjFarmManagement.uctlTextLookup uctlBillType 
         Height          =   495
         Left            =   1800
         TabIndex        =   0
         Top             =   360
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   873
      End
      Begin prjFarmManagement.uctlTextBox txtWeightPerUnit 
         Height          =   435
         Left            =   1800
         TabIndex        =   1
         Top             =   840
         Width           =   1695
         _ExtentX        =   2514
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtRatePerUnit 
         Height          =   435
         Left            =   1800
         TabIndex        =   3
         Top             =   1800
         Width           =   1695
         _ExtentX        =   2514
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPriceTotal 
         Height          =   435
         Left            =   1800
         TabIndex        =   4
         Top             =   2280
         Width           =   1695
         _ExtentX        =   2514
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtNote 
         Height          =   435
         Left            =   1800
         TabIndex        =   5
         Top             =   3480
         Width           =   5775
         _ExtentX        =   2514
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPackAmount 
         Height          =   435
         Left            =   1800
         TabIndex        =   2
         Top             =   1320
         Width           =   1695
         _ExtentX        =   2514
         _ExtentY        =   767
      End
      Begin Threed.SSCheck sscCalDirec 
         Height          =   375
         Left            =   4080
         TabIndex        =   31
         Top             =   2040
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "sscCalDirec"
      End
      Begin VB.Label Label1 
         Height          =   855
         Left            =   3600
         TabIndex        =   30
         Top             =   840
         Width           =   4125
      End
      Begin Threed.SSCheck sscCalPriceInProduct 
         Height          =   375
         Left            =   4080
         TabIndex        =   6
         Top             =   1680
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "sscCalPriceInProduct"
      End
      Begin VB.Label lblPackAmount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   1320
         Width           =   1485
      End
      Begin VB.Label lblBillType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   1485
      End
      Begin VB.Label lblNoteTransport 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   3480
         Width           =   1485
      End
      Begin VB.Label Label10 
         Height          =   345
         Left            =   8880
         TabIndex        =   25
         Top             =   1320
         Width           =   405
      End
      Begin VB.Label Label9 
         Height          =   345
         Left            =   8880
         TabIndex        =   24
         Top             =   840
         Width           =   405
      End
      Begin VB.Label lblPriceTotal 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   2280
         Width           =   1485
      End
      Begin VB.Label Label8 
         Height          =   345
         Left            =   8865
         TabIndex        =   22
         Top             =   360
         Width           =   405
      End
      Begin VB.Label lblRatePerUnit 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   1860
         Width           =   1485
      End
      Begin VB.Label lblWeightPerUnit 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   900
         Width           =   1485
      End
      Begin VB.Label lblStdTrfCharge 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   5190
         TabIndex        =   19
         Top             =   7230
         Width           =   1785
      End
      Begin VB.Label Label6 
         Height          =   345
         Left            =   8865
         TabIndex        =   18
         Top             =   7200
         Width           =   405
      End
      Begin VB.Label lblExcludeDiscount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   90
         TabIndex        =   17
         Top             =   5070
         Width           =   1575
      End
      Begin VB.Label Label7 
         Height          =   345
         Left            =   8880
         TabIndex        =   16
         Top             =   5010
         Width           =   495
      End
      Begin VB.Label Label4 
         Height          =   345
         Left            =   8850
         TabIndex        =   15
         Top             =   4620
         Width           =   495
      End
      Begin VB.Label lblTotalPrice 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   5370
         TabIndex        =   14
         Top             =   5070
         Width           =   1305
      End
      Begin VB.Label Label2 
         Height          =   345
         Left            =   8865
         TabIndex        =   13
         Top             =   4110
         Width           =   435
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2400
         TabIndex        =   9
         Top             =   4080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4080
         TabIndex        =   10
         Top             =   4080
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditTransportItem"
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
Public id As Long
Public OKClick As Boolean
Public TempCollection As Collection
Public TempCollection2 As Collection
Public m_ExDeliveryCostItem As Collection
Public m_DeliveryCus As New Collection
Public CAL_RATE_DELIVERY_TYPE As Long
Public COMMIT_FLAG As String
Public Area As Long
Private m_BillType As Collection
Public DocumentDate As Date
Public StatusFlag As Long
Private RatePerUnit As Double
Public CUSTOMER_ID As Long
Public DeliveryCusId As Long
Public FixRateDetail As String
Dim SearchCusDly As CDeliveryCus
Dim TempD As CExWorksPrice

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
   SSFrame2.BackColor = GLB_FORM_COLOR
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   
   Call txtWeightPerUnit.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtPackAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtRatePerUnit.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtPriceTotal.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)

      
   Call InitNormalLabel(lblBillType, MapText("รายการ"))
   Call InitNormalLabel(lblWeightPerUnit, MapText("น้ำหนัก/หน่วย"))
   Call InitNormalLabel(lblPackAmount, MapText("จำนวนหน่วย"))
   Call InitNormalLabel(lblRatePerUnit, MapText("ราคา/หน่วย"))
   Call InitNormalLabel(lblPriceTotal, MapText("ราคารวม"))
   Call InitNormalLabel(lblNoteTransport, MapText("หมายเหตุ"))
   Call InitNormalLabel(Label1, MapText("*** น้ำหนัก/หน่วย เป็น 1 กรณีที่เป็น BULK" & vbNewLine & "*** น้ำหนัก/หน่วย เป็น 999 กรณี เหมาเที่ยว"))
   Label1.ForeColor = vbRed
   
   Call InitOptionEx(ssoCalDriver, MapText("คิดจากรถขนส่ง"))
   Call InitOptionEx(ssoCalCustomer, MapText("คิดจากลูกค้า"))
   Call InitCheckBox(sscCalPriceInProduct, MapText("คิดค่าขนส่งในสินค้า"))
   Call InitCheckBox(sscCalDirec, MapText("คิดราคาตามอัตราจริง"))
  
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)

   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim iCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
            Dim BTI As CBillTransportItem
            Set BTI = TempCollection.Item(id)
            If BTI.Flag = "N" Then
               uctlBillType.MyTextBox.Text = BTI.BILL_TYPE_CODE
            Else
               uctlBillType.MyCombo.ListIndex = IDToListIndex(uctlBillType.MyCombo, BTI.BILL_TYPE_ID)
            End If
            
            txtPriceTotal.Text = BTI.TOTAL_PRICE
            txtWeightPerUnit.Text = BTI.WEIGHT_PER_UNIT
            txtPackAmount.Text = BTI.PACK_AMOUNT
            RatePerUnit = BTI.RATE_PER_UNIT
            txtRatePerUnit.Text = BTI.RATE_PER_UNIT
            
            ssoCalDriver.Value = FlagToCheck(BTI.CAL_DRIVER)
            ssoCalCustomer.Value = FlagToCheck(BTI.CAL_CUSTOMER)
            sscCalPriceInProduct.Value = FlagToCheck(BTI.CAL_PRICE_IN_PRODUCT)
            sscCalDirec.Value = FlagToCheck(BTI.CAL_DIRECT)
            txtNote.Text = BTI.NOTE
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

   If Not VerifyCombo(lblBillType, uctlBillType.MyCombo, True) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   

      Dim BTI As CBillTransportItem
      If ShowMode = SHOW_ADD Then
         Set BTI = New CBillTransportItem
      Else
          Set BTI = TempCollection.Item(id)
      End If

      BTI.BILL_TYPE_ID = uctlBillType.MyCombo.ItemData(Minus2Zero(uctlBillType.MyCombo.ListIndex))
      BTI.BILL_TYPE_CODE = uctlBillType.MyTextBox.Text
      BTI.BILL_TYPE_NAME = uctlBillType.MyCombo.Text
      BTI.WEIGHT_PER_UNIT = Val(txtWeightPerUnit.Text)
      BTI.PACK_AMOUNT = Val(txtPackAmount.Text)
      BTI.RATE_PER_UNIT = Val(txtRatePerUnit.Text)
      BTI.CAL_DRIVER = Option2Flag(ssoCalDriver.Value)
      BTI.CAL_CUSTOMER = Option2Flag(ssoCalCustomer.Value)
      BTI.CAL_PRICE_IN_PRODUCT = Check2Flag(sscCalPriceInProduct.Value)
      BTI.CAL_DIRECT = Check2Flag(sscCalDirec.Value)
      BTI.TOTAL_PRICE = Val(txtPriceTotal.Text)
      BTI.NOTE = txtNote.Text

   If ShowMode = SHOW_ADD Then
      BTI.Flag = "A"
      Call TempCollection.add(BTI)
   Else
         If BTI.Flag <> "A" And BTI.Flag <> "N" Then
            BTI.Flag = "E"
         End If
   End If
   
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents

      Call LoadMaster(uctlBillType.MyCombo, m_BillType, TRANSPORT_DETAIL)
     Set uctlBillType.MyCollection = m_BillType
           
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
         
         Dim BTI As CBillTransportItem
        Set BTI = TempCollection.Item(id)
         If BTI.Flag = "N" Then
            m_HasModify = True
         Else
            m_HasModify = False
         End If
         
      
         If uctlBillType.MyCombo.ListIndex = 1 Then
            If CAL_RATE_DELIVERY_TYPE = 2 Then 'ถ้าเป็น เหมาเที่ยว ให้บังคับ คิดที่ 30 โล
                txtWeightPerUnit.Text = "999"
            End If
            Set SearchCusDly = GetObject("CDeliveryCus", m_DeliveryCus, Trim(str(DeliveryCusId)), False)
            If Not SearchCusDly Is Nothing Then
               Set TempD = GetObject("CExWorksPrice", m_ExDeliveryCostItem, Trim(str(SearchCusDly.DELIVERY_CUS_ITEM_ID)) & "-" & Trim(txtWeightPerUnit.Text), False)   'ค้นหาราคาค่าขนส่ง  ที่ 1 กิโลเลย
               If Not TempD Is Nothing Then
                 If CAL_RATE_DELIVERY_TYPE = 2 Then 'ถ้าเป็น เหมาเที่ยว
                     txtRatePerUnit.Text = TempD.RATE_DELIVERY
                     txtPriceTotal.Text = TempD.RATE_DELIVERY
                     txtNote.Text = "เหมาเที่ยว"
                     
                     If txtWeightPerUnit.Text = "999" Then
                        txtWeightPerUnit.Text = "1"
                     End If
                     
'                     txtWeightPerUnit.Enabled = False
'                     txtPackAmount.Enabled = False
'                     txtRatePerUnit.Enabled = False
'                     txtPriceTotal.Enabled = False
                  Else
                     txtRatePerUnit.Text = TempD.RATE_DELIVERY
                  End If
               Else
                     glbErrorLog.LocalErrorMsg = "ไม่มีข้อมูลราคาค่าขนส่งคิดให้รถรับจ้าง กรุณาระบุสถานที่จัดส่ง หรือ เงื่อนไขการคิดค่าขนส่งรถรับจ้าง ใหม่อีกครั้ง"
                    glbErrorLog.ShowUserError
      
                  txtRatePerUnit.Text = ""
               End If
            End If
         ElseIf uctlBillType.MyCombo.ListIndex = 2 Then
'            If CAL_RATE_DELIVERY_TYPE = 2 Then 'ถ้าเป็น เหมาเที่ยว ให้บังคับ คิดที่ 30 โล
'                txtWeightPerUnit.Text = "999"
                txtRatePerUnit.Text = txtRatePerUnit.Text
'            End If
             
         End If
      ElseIf ShowMode = SHOW_ADD Then
         id = 0
         m_HasModify = False
      End If
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
   
   Set m_Rs = New ADODB.Recordset
   Set m_BillType = New Collection

End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   
   Set m_Rs = Nothing
   Set m_BillType = Nothing

End Sub

Private Sub txtDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtKeyName_Change()
   m_HasModify = True
End Sub



Private Sub txtRateDriverTransport_Change()
   m_HasModify = True
End Sub

Private Sub SSCheck1_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub SSCheck2_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub sscCalCustomer_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub sscCalDriver_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub sscCalDirec_Click(Value As Integer)
   m_HasModify = True
   txtPriceTotal.Text = Val(txtRatePerUnit.Text) * Val(txtPackAmount.Text)
End Sub

Private Sub sscCalPriceInProduct_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub ssoCalCustomer_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub ssoCalDriver_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub txtNote_Change()
   m_HasModify = True
End Sub

Private Sub txtPackAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtPackAmount_KeyPress(KeyAscii As Integer)
 KeyAscii = CheckIntAscii(KeyAscii)
End Sub

Private Sub txtPriceTotal_Change()
   m_HasModify = True
End Sub

Private Sub txtPriceTotal_KeyPress(KeyAscii As Integer)
 KeyAscii = CheckIntAscii(KeyAscii)
End Sub

Private Sub txtRatePerUnit_Change()
   m_HasModify = True
   If uctlBillType.MyCombo.ListIndex = 1 Then
    If CAL_RATE_DELIVERY_TYPE = 2 Then
      txtPriceTotal.Text = txtRatePerUnit.Text
    Else
      If Val(txtRatePerUnit.Text) * Val(txtPackAmount.Text) = 0 Then
      ElseIf Val(txtWeightPerUnit.Text) > 1 And Val(txtWeightPerUnit.Text) < 30 Then
            txtPriceTotal.Text = ((Val(txtRatePerUnit.Text) * Val(txtWeightPerUnit.Text)) / 30) * Val(txtPackAmount.Text)
      Else
             txtPriceTotal.Text = Val(txtRatePerUnit.Text) * Val(txtPackAmount.Text)
      End If
   End If
   Else
      If Val(txtRatePerUnit.Text) * Val(txtPackAmount.Text) = 0 Then
      ElseIf Val(txtWeightPerUnit.Text) > 1 And Val(txtWeightPerUnit.Text) < 30 Then
            txtPriceTotal.Text = ((Val(txtRatePerUnit.Text) * Val(txtWeightPerUnit.Text)) / 30) * Val(txtPackAmount.Text)
      Else
             txtPriceTotal.Text = Val(txtRatePerUnit.Text) * Val(txtPackAmount.Text)
      End If
   End If
End Sub

Private Sub txtRatePerUnit_KeyPress(KeyAscii As Integer)
 KeyAscii = CheckIntAscii(KeyAscii)
End Sub

Private Sub txtWeightPerUnit_Change()
   m_HasModify = True
End Sub


Private Sub txtWeightPerUnit_KeyPress(KeyAscii As Integer)
   KeyAscii = CheckIntAscii(KeyAscii)
End Sub

Private Sub uctlBillType_Change()
   m_HasModify = True
   
    If uctlBillType.MyCombo.ListIndex = 5 Or uctlBillType.MyCombo.ListIndex = 9 Then  'รายได้อื่นๆ,รายได้จากการเรียกคืนขนส่ง
      ssoCalDriver.Caption = "คิดจากรถขนส่ง"
       ssoCalCustomer.Caption = "คิดจากลูกค้า"
      SSFrame2.Visible = True
   ElseIf uctlBillType.MyCombo.ListIndex = 6 Or uctlBillType.MyCombo.ListIndex = 10 Then   'รายจ่ายอื่นๆ,รายจ่ายเพิ่มเติมขนส่ง
      ssoCalDriver.Caption = "คิดให้รถขนส่ง"
      ssoCalCustomer.Caption = "คิดให้ลูกค้า"
      SSFrame2.Visible = True
   Else
      SSFrame2.Visible = False
   End If
End Sub


