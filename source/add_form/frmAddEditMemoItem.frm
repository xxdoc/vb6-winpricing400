VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditMemoItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditMemoItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   6765
      Left            =   0
      TabIndex        =   19
      Top             =   600
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   11933
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextBox txtNo 
         Height          =   435
         Left            =   1800
         TabIndex        =   0
         Top             =   270
         Width           =   3375
         _ExtentX        =   4471
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlInDate 
         Height          =   435
         Left            =   6720
         TabIndex        =   1
         Top             =   240
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTicketType 
         Height          =   435
         Left            =   1800
         TabIndex        =   2
         Top             =   720
         Width           =   3375
         _ExtentX        =   4471
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtInvoiceNo 
         Height          =   435
         Left            =   6720
         TabIndex        =   3
         Top             =   720
         Width           =   3855
         _ExtentX        =   14790
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtCurrencyOther 
         Height          =   435
         Left            =   1800
         TabIndex        =   4
         Top             =   1200
         Width           =   2415
         _ExtentX        =   14790
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtRatio 
         Height          =   435
         Left            =   5640
         TabIndex        =   5
         Top             =   1200
         Width           =   2295
         _ExtentX        =   14790
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtCostBaht 
         Height          =   435
         Left            =   9480
         TabIndex        =   6
         Top             =   1200
         Width           =   2295
         _ExtentX        =   14790
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTaxPercent 
         Height          =   435
         Left            =   1800
         TabIndex        =   7
         Top             =   1680
         Width           =   2415
         _ExtentX        =   14790
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTaxIn 
         Height          =   435
         Left            =   5640
         TabIndex        =   8
         Top             =   1680
         Width           =   2295
         _ExtentX        =   14790
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtVat 
         Height          =   435
         Left            =   9480
         TabIndex        =   9
         Top             =   1680
         Width           =   2295
         _ExtentX        =   14790
         _ExtentY        =   767
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   2895
         Left            =   840
         TabIndex        =   30
         Top             =   3000
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   5106
         _Version        =   131073
         Enabled         =   0   'False
         PictureBackgroundStyle=   1
         Begin prjFarmManagement.uctlTextBox txtAmount 
            Height          =   435
            Left            =   2400
            TabIndex        =   14
            Top             =   1800
            Width           =   2535
            _ExtentX        =   2355
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtDesc 
            Height          =   435
            Left            =   2400
            TabIndex        =   11
            Top             =   360
            Width           =   3855
            _ExtentX        =   14790
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlDate uctlDate 
            Height          =   435
            Left            =   2400
            TabIndex        =   12
            Top             =   840
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtDol 
            Height          =   435
            Left            =   2400
            TabIndex        =   13
            Top             =   1320
            Width           =   2535
            _ExtentX        =   2355
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtTax 
            Height          =   435
            Left            =   2400
            TabIndex        =   15
            Top             =   2280
            Width           =   2535
            _ExtentX        =   2355
            _ExtentY        =   767
         End
         Begin VB.Label lblTax 
            Alignment       =   1  'Right Justify
            Caption         =   "lblTax"
            Height          =   375
            Left            =   840
            TabIndex        =   35
            Top             =   2280
            Width           =   1485
         End
         Begin VB.Label lblDol 
            Alignment       =   1  'Right Justify
            Caption         =   "lblDol"
            Height          =   375
            Left            =   360
            TabIndex        =   34
            Top             =   1320
            Width           =   1965
         End
         Begin VB.Label lblAmount 
            Alignment       =   1  'Right Justify
            Caption         =   "lblAmount"
            Height          =   375
            Left            =   840
            TabIndex        =   33
            Top             =   1800
            Width           =   1485
         End
         Begin VB.Label lblDesc 
            Alignment       =   1  'Right Justify
            Caption         =   "lblDesc"
            Height          =   375
            Left            =   840
            TabIndex        =   32
            Top             =   480
            Width           =   1485
         End
         Begin VB.Label lblDate 
            Alignment       =   1  'Right Justify
            Caption         =   "lblDate"
            Height          =   375
            Left            =   840
            TabIndex        =   31
            Top             =   870
            Width           =   1485
         End
      End
      Begin Threed.SSCheck chkPaid 
         Height          =   495
         Left            =   840
         TabIndex        =   10
         Top             =   2400
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   131073
         Caption         =   "chkPaid"
      End
      Begin VB.Label lblVat 
         Alignment       =   1  'Right Justify
         Caption         =   "lblVat"
         Height          =   375
         Left            =   8040
         TabIndex        =   29
         Top             =   1740
         Width           =   1365
      End
      Begin VB.Label lblTaxIn 
         Alignment       =   1  'Right Justify
         Caption         =   "lblTaxIn"
         Height          =   375
         Left            =   4200
         TabIndex        =   28
         Top             =   1740
         Width           =   1365
      End
      Begin VB.Label lblTaxPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "lblTaxPercent"
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   1740
         Width           =   1485
      End
      Begin VB.Label lblCostBaht 
         Alignment       =   1  'Right Justify
         Caption         =   "lblCostBaht"
         Height          =   375
         Left            =   8040
         TabIndex        =   26
         Top             =   1260
         Width           =   1365
      End
      Begin VB.Label lblRatio 
         Alignment       =   1  'Right Justify
         Caption         =   "lblRatio"
         Height          =   375
         Left            =   4200
         TabIndex        =   25
         Top             =   1260
         Width           =   1365
      End
      Begin VB.Label lblCurrencyOther 
         Alignment       =   1  'Right Justify
         Caption         =   "lblCurrencyOther"
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   1260
         Width           =   1485
      End
      Begin VB.Label lblInvoiceNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblInvoiceNo"
         Height          =   375
         Left            =   5160
         TabIndex        =   23
         Top             =   780
         Width           =   1485
      End
      Begin VB.Label lblTicketType 
         Alignment       =   1  'Right Justify
         Caption         =   "lblTicketType"
         Height          =   375
         Left            =   0
         TabIndex        =   22
         Top             =   750
         Width           =   1725
      End
      Begin VB.Label lblInDate 
         Alignment       =   1  'Right Justify
         Caption         =   "lblInDate"
         Height          =   375
         Left            =   5160
         TabIndex        =   21
         Top             =   360
         Width           =   1485
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   4320
         TabIndex        =   16
         Top             =   6000
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditMemoItem.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   6000
         TabIndex        =   17
         Top             =   6000
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblNo"
         Height          =   375
         Left            =   0
         TabIndex        =   20
         Top             =   300
         Width           =   1725
      End
   End
End
Attribute VB_Name = "frmAddEditMemoItem"
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
Private DOLLAR As Double
Private TAX As Double
Private COEF As Double
Private Date1 As Date
Private UNIT As Long


Private Sub chkPaid_Click(Value As Integer)
   m_HasModify = True
   If chkPaid.Value = ssCBChecked Then
   SSFrame2.Enabled = True
   Else
   SSFrame2.Enabled = False
   txtDesc.Text = ""
   uctlDate.ShowDate = -1
   txtDol.Text = ""
   txtAmount.Text = ""
   txtTax.Text = ""
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
   SSFrame2.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   Me.KeyPreview = True
   pnlHeader.Caption = HeaderText
   Me.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_HEAD_COLOR
   SSFrame1.BackColor = GLB_FORM_COLOR
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   
   Call InitCheckBox(chkPaid, "จ่ายเงินแล้ว")
   Call InitNormalLabel(lblNo, MapText("หมายเลข MEMO"))
   Call InitNormalLabel(lblInDate, MapText("วันที่เข้า"))
   Call InitNormalLabel(lblTicketType, MapText("ลักษณะตั๋ว"))
   Call InitNormalLabel(lblInvoiceNo, MapText("หมายเลข INV"))
   Call InitNormalLabel(lblCurrencyOther, MapText("เงิน ต.ป.ท."))
   Call InitNormalLabel(lblRatio, MapText("อัตรา"))
   Call InitNormalLabel(lblCostBaht, MapText("ราคาบาท"))
   Call InitNormalLabel(lblTaxPercent, MapText("อัตราภาษี"))
   Call InitNormalLabel(lblTaxIn, MapText("ภาษีนำเข้า"))
   Call InitNormalLabel(lblVat, MapText("ภาษีมูลค่าเพิ่ม"))
   
   Call InitNormalLabel(lblDesc, MapText("รายละเอียด"))
   Call InitNormalLabel(lblDate, MapText("วันที่จ่าย"))
   Call InitNormalLabel(lblAmount, MapText("เงินบาท"))
   Call InitNormalLabel(lblDol, MapText("อัตราแลกเปลี่ยน"))
   Call InitNormalLabel(lblTax, MapText("ดอกเบี้ย"))
   
   Call txtNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtTicketType.SetTextLenType(TEXT_STRING, glbSetting.NAME_TYPE)
   Call txtInvoiceNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtCurrencyOther.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtRatio.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   Call txtCostBaht.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   Call txtTaxPercent.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   Call txtTaxIn.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   Call txtVat.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   
   Call txtDesc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
    Call txtAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
    Call txtDol.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
    Call txtTax.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   
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
           Dim CustProof As CMemoBank
         Set CustProof = TempCollection.Item(ID)
         
         txtNo.Text = CustProof.MEMO_BANK_NO
         uctlInDate.ShowDate = CustProof.EXCHANGE_DATE
         txtTicketType.Text = CustProof.TICKET_TYPE
         txtInvoiceNo.Text = CustProof.INVOICE_NO
         txtCurrencyOther.Text = CustProof.CURRENCY_OTHER
         txtRatio.Text = Val(CustProof.RATIO)
         txtCostBaht.Text = Val(CustProof.COST_BAHT)
         txtTaxPercent.Text = Val(CustProof.TAX_PERCENT)
         txtTaxIn.Text = Val(CustProof.TAX_IN)
          txtVat.Text = Val(CustProof.VAT)
          
          chkPaid.Value = FlagToCheck(CustProof.PAID_FLAG)
          
         txtDesc.Text = CustProof.DESCRIPTION
         txtAmount.Text = CustProof.AMOUNT_THAI
         uctlDate.ShowDate = CustProof.MEMO_BANK_DATE
         txtTax.Text = CustProof.TAX
         txtDol.Text = CustProof.AMOUNT_OTHER
         
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

   If Not VerifyTextControl(lblNo, txtNo, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblInDate, uctlInDate, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblRatio, txtRatio, True) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblCostBaht, txtCostBaht, True) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblTaxPercent, txtTaxPercent, True) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblTaxIn, txtTaxIn, True) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblVat, txtVat, True) Then
      Exit Function
   End If
   If chkPaid.Value = ssCBChecked Then
   If Not VerifyDate(lblDate, uctlDate, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblAmount, txtAmount, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblDol, txtDol, True) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblTax, txtTax, True) Then
      Exit Function
   End If
   End If
 
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
      Dim CustProof As CMemoBank
   If ShowMode = SHOW_ADD Then
      Set CustProof = New CMemoBank
    Else
      Set CustProof = TempCollection.Item(ID)
   End If
   
   CustProof.MEMO_BANK_NO = txtNo.Text
   CustProof.EXCHANGE_DATE = uctlInDate.ShowDate
   CustProof.TICKET_TYPE = txtTicketType.Text
   CustProof.INVOICE_NO = txtInvoiceNo.Text
   CustProof.CURRENCY_OTHER = txtCurrencyOther.Text
   CustProof.RATIO = Val(txtRatio.Text)
   CustProof.COST_BAHT = Val(txtCostBaht.Text)
   CustProof.TAX_PERCENT = Val(txtTaxPercent.Text)
   CustProof.TAX_IN = Val(txtTaxIn.Text)
   CustProof.VAT = Val(txtVat.Text)
   
   CustProof.PAID_FLAG = Check2Flag(chkPaid.Value)
   
   CustProof.DESCRIPTION = txtDesc.Text
   CustProof.AMOUNT_THAI = Val(txtAmount.Text)
   CustProof.AMOUNT_OTHER = Val(txtDol.Text)
   CustProof.MEMO_BANK_DATE = uctlDate.ShowDate
   CustProof.COEFFICIENT = COEF
   CustProof.TAX = Val(txtTax.Text)
'   CustProof.EXCHANGE_DATE = Date1
   CustProof.UNIT = UNIT
   
   If ShowMode = SHOW_ADD Then
      CustProof.Flag = "A"
      Call TempCollection.add(CustProof)
      Else
      If CustProof.Flag <> "A" Then
      CustProof.Flag = "E"
      End If
         End If
   
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
         uctlInDate.ShowDate = Now
         Call chkPaid_Click(2)
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
   Set m_Rs = New ADODB.Recordset
End Sub


Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
End Sub

'Private Sub cmddollar_Click()
'Dim OKClick As Boolean
'Dim LayoutID As Long
'Dim TempID As Double
'Dim TempID1 As Long
'Dim TempStr As Double
'
'frmAddEditDollar.HeaderText = "เปลี่ยนแปลงเงินตรา"
'   Load frmAddEditDollar
'   frmAddEditDollar.Show 1
'   If frmAddEditDollar.OKClick Then
'      TempID = frmAddEditDollar.Dollar1
'      TempStr = frmAddEditDollar.dollar2
'      TempID1 = frmAddEditDollar.dollarID
'       COEF = frmAddEditDollar.COEF
'       Date1 = frmAddEditDollar.Date1
'       UNIT = frmAddEditDollar.COUNTRY_CURRENCY1
'   End If
'
'   Unload frmAddEditDollar
'   Set frmAddEditDollar = Nothing
'
'   DOLLAR = TempStr
'   txtDol.Text = DOLLAR
'   txtAmount.Text = TempID
'
' m_HasModify = True
'   End Sub
'
'
'Private Sub cmdTAXCURRENT_Click()
'Dim OKClick As Boolean
'Dim tax1 As Double
'Dim tax2 As Double
'Dim tax3 As Double
'
'    frmAddEditTaxCurrent.HeaderText = "แปลงเงินปัจจุบันเป็นเงินบวกภาษี"
'   Load frmAddEditTaxCurrent
'   frmAddEditTaxCurrent.Show 1
'   If frmAddEditTaxCurrent.OKClick Then
'      tax1 = frmAddEditTaxCurrent.tax1
'      tax2 = frmAddEditTaxCurrent.tax2
'      tax3 = frmAddEditTaxCurrent.tax3
'     TAX = tax1 + tax2 + tax3
'     txtTax.Text = Val(txtAmount.Text) + (Val(txtAmount.Text) * TAX / 100)
'   End If
'
'   Unload frmAddEditTaxCurrent
'   Set frmAddEditTaxCurrent = Nothing
'
'   m_HasModify = True
'   End Sub

Private Sub txtAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtCostBaht_Change()
   m_HasModify = True
End Sub

Private Sub txtCurrencyOther_Change()
   m_HasModify = True
End Sub

Private Sub txtDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtDol_Change()
   m_HasModify = True
End Sub

Private Sub txtInvoiceNo_Change()
   m_HasModify = True
End Sub

Private Sub txtNo_Change()
   m_HasModify = True
End Sub

Private Sub txtRatio_Change()
   m_HasModify = True
End Sub

Private Sub txtTax_Change()
   m_HasModify = True
End Sub

Private Sub txtTaxIn_Change()
   m_HasModify = True
End Sub

Private Sub txtTaxPercent_Change()
   m_HasModify = True
End Sub

Private Sub txtTicketType_Change()
   m_HasModify = True
End Sub

Private Sub txtVat_Change()
   m_HasModify = True
End Sub

Private Sub uctlDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlInDate_HasChange()
   m_HasModify = True
End Sub
