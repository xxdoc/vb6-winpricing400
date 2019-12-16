VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditTransport 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12090
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditTransport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10875
   ScaleWidth      =   12090
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   12105
      _ExtentX        =   21352
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   10275
      Left            =   0
      TabIndex        =   16
      Top             =   600
      Width           =   12105
      _ExtentX        =   21352
      _ExtentY        =   18124
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlDate uctlShippingDate 
         Height          =   495
         Left            =   1800
         TabIndex        =   4
         Top             =   1800
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   873
      End
      Begin prjFarmManagement.uctlTextLookup uctlSupplierTransport 
         Height          =   435
         Left            =   1800
         TabIndex        =   1
         Top             =   840
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtSupplierName 
         Height          =   435
         Left            =   3360
         TabIndex        =   2
         Top             =   1320
         Width           =   3075
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtDocumentNo 
         Height          =   435
         Left            =   1800
         TabIndex        =   0
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtNoteTransport 
         Height          =   435
         Left            =   1800
         TabIndex        =   8
         Top             =   3720
         Width           =   6975
         _ExtentX        =   2514
         _ExtentY        =   767
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   4815
         Left            =   240
         TabIndex        =   9
         Top             =   4440
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   8493
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         TabKeyBehavior  =   1
         MethodHoldFields=   -1  'True
         AllowColumnDrag =   0   'False
         AllowEdit       =   0   'False
         BorderStyle     =   3
         GroupByBoxVisible=   0   'False
         DataMode        =   99
         HeaderFontBold  =   -1  'True
         HeaderFontWeight=   700
         FontSize        =   9.75
         BackColorBkg    =   16777215
         ColumnHeaderHeight=   480
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmAddEditTransport.frx":08CA
         Column(2)       =   "frmAddEditTransport.frx":0992
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditTransport.frx":0A36
         FormatStyle(2)  =   "frmAddEditTransport.frx":0B92
         FormatStyle(3)  =   "frmAddEditTransport.frx":0C42
         FormatStyle(4)  =   "frmAddEditTransport.frx":0CF6
         FormatStyle(5)  =   "frmAddEditTransport.frx":0DCE
         ImageCount      =   0
         PrinterProperties=   "frmAddEditTransport.frx":0E86
      End
      Begin prjFarmManagement.uctlTextBox txtAround 
         Height          =   435
         Left            =   7800
         TabIndex        =   3
         Top             =   1320
         Width           =   915
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtShipping 
         Height          =   435
         Left            =   1800
         TabIndex        =   6
         Top             =   2760
         Width           =   6975
         _ExtentX        =   2514
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtFixDetail 
         Height          =   435
         Left            =   1800
         TabIndex        =   7
         Top             =   3240
         Width           =   6975
         _ExtentX        =   2514
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtSupplierCode 
         Height          =   435
         Left            =   1800
         TabIndex        =   35
         Top             =   1320
         Width           =   1515
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlDeliveryCusLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   36
         Top             =   2280
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   767
      End
      Begin Threed.SSFrame SSFrame3 
         Height          =   1455
         Left            =   8880
         TabIndex        =   38
         Top             =   2760
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   2566
         _Version        =   131073
         Caption         =   "SSFrame3"
         Begin Threed.SSOption ssoVolume 
            Height          =   375
            Left            =   240
            TabIndex        =   40
            Top             =   360
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   131073
            Caption         =   "ssoVolume"
         End
         Begin Threed.SSOption ssoRound 
            Height          =   375
            Left            =   240
            TabIndex        =   39
            Top             =   840
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   131073
            Caption         =   "ssoRound"
         End
      End
      Begin Threed.SSCommand cmdAccessDeliveryCus 
         Height          =   405
         Left            =   8760
         TabIndex        =   41
         Top             =   2280
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditTransport.frx":105E
         ButtonStyle     =   3
      End
      Begin VB.Label lblDeliveryCusLookup 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   37
         Top             =   2280
         Width           =   1485
      End
      Begin Threed.SSCheck sscNotCalFlag 
         Height          =   375
         Left            =   7560
         TabIndex        =   5
         Top             =   1800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "sscNotCalFlag"
      End
      Begin Threed.SSCommand cmdAddData 
         Height          =   405
         Left            =   8760
         TabIndex        =   34
         Top             =   840
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditTransport.frx":1378
         ButtonStyle     =   3
      End
      Begin VB.Label lblFixDetail 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   3240
         Width           =   1485
      End
      Begin VB.Label lblShippingDate 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   1800
         Width           =   1485
      End
      Begin VB.Label lblShipping 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   2760
         Width           =   1485
      End
      Begin VB.Label lblAround 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   6480
         TabIndex        =   30
         Top             =   1320
         Width           =   1245
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3600
         TabIndex        =   12
         Top             =   9480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1920
         TabIndex        =   11
         Top             =   9480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   240
         TabIndex        =   10
         Top             =   9480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin VB.Label lblNoteTransport 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   3720
         Width           =   1485
      End
      Begin VB.Label lblDocumentNo 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   180
         TabIndex        =   28
         Top             =   420
         Width           =   1485
      End
      Begin VB.Label lblSupplierTransport 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   840
         Width           =   1485
      End
      Begin VB.Label lblSupplierCode 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   1320
         Width           =   1485
      End
      Begin VB.Label lblStdTrfCharge 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   5190
         TabIndex        =   25
         Top             =   7230
         Width           =   1785
      End
      Begin VB.Label Label6 
         Height          =   345
         Left            =   8865
         TabIndex        =   24
         Top             =   7200
         Width           =   405
      End
      Begin VB.Label lblExcludeDiscount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   90
         TabIndex        =   23
         Top             =   5070
         Width           =   1575
      End
      Begin VB.Label lblPackAmount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   3240
         TabIndex        =   22
         Top             =   3660
         Width           =   1275
      End
      Begin VB.Label Label7 
         Height          =   345
         Left            =   8880
         TabIndex        =   21
         Top             =   5010
         Width           =   495
      End
      Begin VB.Label lblUnit 
         Height          =   345
         Left            =   3300
         TabIndex        =   20
         Top             =   4110
         Width           =   1215
      End
      Begin VB.Label Label4 
         Height          =   345
         Left            =   8850
         TabIndex        =   19
         Top             =   4620
         Width           =   495
      End
      Begin VB.Label lblTotalPrice 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   5370
         TabIndex        =   18
         Top             =   5070
         Width           =   1305
      End
      Begin VB.Label Label2 
         Height          =   345
         Left            =   8865
         TabIndex        =   17
         Top             =   4110
         Width           =   435
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8640
         TabIndex        =   13
         Top             =   9480
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
         Left            =   10320
         TabIndex        =   14
         Top             =   9480
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditTransport"
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
Private m_SuppliersTransport As Collection
Private m_BillTransport As CBillTransport
Public CAL_RATE_DELIVERY_TYPE As Long
Public DocumentDate As Date
Public BillingdocID As Long
Public DocumentNo As String
Public TruckNo As String
Public CUSTOMER_ID As Long
Public m_ExDeliveryCostItem As Collection
Public m_DeliveryCus As Collection
Public DELIVERY_CUS_ITEM_ID As Long
Private DeliveryCusId As Long
Private TempD As CExWorksPrice

Private Sub cmdAccessDeliveryCus_Click()
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu
   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("ข้อมูลค่าขนส่ง", "-", "ข้อมูลสถานที่จัดส่ง")
   If lMenuChosen = 0 Then
      Exit Sub
   End If

If lMenuChosen = 1 Then
     If Not VerifyAccessRight("PACKAGE-CENTER_DELIVERY-COST") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      frmExWorksPrice.Area = 2
      Load frmExWorksPrice
      frmExWorksPrice.Show 1
      Unload frmExWorksPrice
      Set frmExWorksPrice = Nothing
ElseIf lMenuChosen = 3 Then
      frmAddEditDeliveryCusMain.HeaderText = MapText("ข้อมูลสถานที่จัดส่ง")
      frmAddEditDeliveryCusMain.CustomerID = CUSTOMER_ID
      Load frmAddEditDeliveryCusMain
      frmAddEditDeliveryCusMain.Show 1

      OKClick = frmAddEditDeliveryCusMain.OKClick

      Unload frmAddEditDeliveryCusMain
      Set frmAddEditDeliveryCusMain = Nothing
End If

 If CUSTOMER_ID > 0 Then
   Call LoadDeliveryCus(uctlDeliveryCusLookup.MyCombo, m_DeliveryCus, CUSTOMER_ID) 'LOAD สถานที่จัดส่ง
   Set uctlDeliveryCusLookup.MyCollection = m_DeliveryCus
End If
Call LoadExDeliveryCusItem(Nothing, m_ExDeliveryCostItem, , 3, DocumentDate) 'ส่ง 3 ไปดึงค่าขนส่งที่คิดให้รถรับจ้าง
      
End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean

   If Not VerifyCombo(lblDeliveryCusLookup, uctlDeliveryCusLookup.MyCombo, False) Then
      Exit Sub
   End If
   
      Set frmAddEditTransportItem.TempCollection = m_BillTransport.C_BillTransportItem
      frmAddEditTransportItem.ShowMode = SHOW_ADD
      frmAddEditTransportItem.HeaderText = MapText("เพิ่มรายการค่าขนส่ง")
      frmAddEditTransportItem.CUSTOMER_ID = CUSTOMER_ID
      frmAddEditTransportItem.DeliveryCusId = DeliveryCusId
      frmAddEditTransportItem.CAL_RATE_DELIVERY_TYPE = CAL_RATE_DELIVERY_TYPE
      Set frmAddEditTransportItem.m_ExDeliveryCostItem = m_ExDeliveryCostItem
      Set frmAddEditTransportItem.m_DeliveryCus = m_DeliveryCus
      Load frmAddEditTransportItem
      frmAddEditTransportItem.Show 1
      
      OKClick = frmAddEditTransportItem.OKClick
      
      Unload frmAddEditTransportItem
      Set frmAddEditTransportItem = Nothing
   
      If OKClick Then
         m_HasModify = True
         GridEX1.ItemCount = CountItem(m_BillTransport.C_BillTransportItem)
         GridEX1.Rebind
      End If
End Sub

Private Sub cmdCustomer_Click()

End Sub

Private Sub cmdAddData_Click()
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu
Dim OKClick As Boolean
Dim TempCol As Collection
Dim Cs As CCustomer
   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("เพิ่มข้อมูลซัพฯ")
   If lMenuChosen = 0 Then
      Exit Sub
   End If

If lMenuChosen = 1 Then
   If Not VerifyAccessRight("MAIN_SUPPLIER") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If

      Load frmSupplier
      frmSupplier.Show 1
      
      Unload frmSupplier
      Set frmSupplier = Nothing

      Call LoadSupplierTransport(uctlSupplierTransport.MyCombo, m_SuppliersTransport)
      Set uctlSupplierTransport.MyCollection = m_SuppliersTransport
      
   End If
End Sub

Private Sub cmdDelete_Click()
Dim ID1 As Long
Dim ID2 As Long

   If Not cmdDelete.Enabled Then
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
   If Not ConfirmDelete(GridEX1.Value(3)) Then
      Exit Sub
   End If
   
   ID1 = GridEX1.Value(1)
   ID2 = GridEX1.Value(2)
   
   
    If ID1 <= 0 Then
      Call m_BillTransport.C_BillTransportItem.Remove(ID2)
   Else
      m_BillTransport.C_BillTransportItem.Item(ID2).Flag = "D"
   End If

   GridEX1.ItemCount = CountItem(m_BillTransport.C_BillTransportItem)
   GridEX1.Rebind
   m_HasModify = True
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean

   If Not VerifyCombo(lblDeliveryCusLookup, uctlDeliveryCusLookup.MyCombo, False) Then
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(2))
   frmAddEditTransportItem.ID = ID
   frmAddEditTransportItem.ShowMode = SHOW_EDIT
   Set frmAddEditTransportItem.TempCollection = m_BillTransport.C_BillTransportItem
   Set frmAddEditTransportItem.m_ExDeliveryCostItem = m_ExDeliveryCostItem
   Set frmAddEditTransportItem.m_DeliveryCus = m_DeliveryCus
   frmAddEditTransportItem.HeaderText = HeaderText
   frmAddEditTransportItem.CUSTOMER_ID = CUSTOMER_ID
   frmAddEditTransportItem.CAL_RATE_DELIVERY_TYPE = SetValue
   frmAddEditTransportItem.DeliveryCusId = DeliveryCusId
   Load frmAddEditTransportItem
   frmAddEditTransportItem.Show 1
   
   OKClick = frmAddEditTransportItem.OKClick
   
   Unload frmAddEditTransportItem
   Set frmAddEditTransportItem = Nothing
               
   If OKClick Then
       m_HasModify = True
      GridEX1.ItemCount = CountItem(m_BillTransport.C_BillTransportItem)
      GridEX1.Rebind
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
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   
   Call txtAround.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtShipping.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)

   Call InitGrid1
   
   Call InitNormalLabel(lblDocumentNo, MapText("เลขที่เอกสาร"))
   Call InitNormalLabel(lblSupplierTransport, MapText("จ่ายให้กับ"))
   Call InitNormalLabel(lblSupplierCode, MapText("ชื่อซัพฯ"))
   Call InitNormalLabel(lblAround, MapText("จำนวนเที่ยว"))
   Call InitNormalLabel(lblShippingDate, MapText("วันที่ส่งของ"))
   Call InitNormalLabel(lblShipping, MapText("สถานที่จัดส่ง"))
   Call InitNormalLabel(lblFixDetail, MapText("Rate อ้างอิง"))
   Call InitNormalLabel(lblNoteTransport, MapText("หมายเหตุ"))
   Call InitCheckBox(sscNotCalFlag, MapText("ไม่คิดภาษี"))
   Call InitNormalLabel(lblDeliveryCusLookup, MapText("สถานที่จัดส่ง"))
   
   Call InitNormalFrame(SSFrame3, "เงื่อนไขคิดค่าขนส่งรถรับจ้าง")
   Call InitOptionEx(ssoVolume, "คิดตามปริมาณ")
   Call InitOptionEx(ssoRound, "คิดตามเที่ยว")
   
   ssoVolume.Value = True
   
   'sscNotCalFlag
   txtDocumentNo.Enabled = False
'   txtSupplierName.Enabled = False
   uctlShippingDate.ShowDate = DocumentDate
   txtAround.Text = "1"
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdExit, MapText("ออก (ESC)"))
   
   Call InitMainButton(cmdAddData, MapText("A"))
   Call InitMainButton(cmdAccessDeliveryCus, MapText("A"))
   
      
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAddData.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAccessDeliveryCus.Picture = LoadPicture(glbParameterObj.NormalButton1)
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim iCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
    txtDocumentNo.Text = DocumentNo
    uctlSupplierTransport.MyTextBox.Text = TruckNo
      If ShowMode = SHOW_EDIT Then
         Set m_BillTransport = TempCollection2.Item(1)
         m_BillTransport.AddEditMode = SHOW_EDIT
         m_BillTransport.Flag = "E"
         uctlSupplierTransport.MyCombo.ListIndex = IDToListIndex(uctlSupplierTransport.MyCombo, m_BillTransport.SUPPLIER_TRANSPORT_ID)
         txtAround.Text = m_BillTransport.AROUND
         txtShipping.Text = m_BillTransport.SHIPPING
         uctlShippingDate.ShowDate = m_BillTransport.SHIPPING_DATE
         sscNotCalFlag.Value = FlagToCheck(m_BillTransport.NOT_CAL_VAT)
         txtFixDetail.Text = m_BillTransport.FIX_RATE_DETAIL
         txtNoteTransport.Text = m_BillTransport.NOTE
         Call GetValue(m_BillTransport.CAL_RATE_DELIVERY_TYPE)
         uctlDeliveryCusLookup.MyCombo.ListIndex = IDToListIndex(uctlDeliveryCusLookup.MyCombo, m_BillTransport.EX_DELIVERY_COST_ITEM_ID)
      ElseIf ShowMode = SHOW_ADD Then
      If CountItem(TempCollection2) = 0 Then
         Set m_BillTransport = New CBillTransport
         m_BillTransport.AddEditMode = SHOW_ADD
         m_BillTransport.Flag = "A"
         m_BillTransport.BILLING_DOC_ID = BillingdocID
         
         Dim Di As CDoItem
         Dim BTI As CBillTransportItem
         Dim TempBTI As CBillTransportItem
         Set TempCollection2 = New Collection
         
         For Each Di In TempCollection
           If Di.PART_ITEM_ID > 0 Then
               Set TempBTI = GetObject("CBillTransportItem", m_BillTransport.C_BillTransportItem, Trim(str(Di.WEIGHT_PER_PACK)) & "-" & Trim("01"), False)
               If Not (TempBTI Is Nothing) Then
                  TempBTI.PACK_AMOUNT = TempBTI.PACK_AMOUNT + Di.PACK_AMOUNT
               Else
                  Set BTI = New CBillTransportItem
                  BTI.PACK_AMOUNT = Di.PACK_AMOUNT
                  BTI.WEIGHT_PER_UNIT = Di.WEIGHT_PER_PACK
                  BTI.BILL_TRANSPORT_ITEM_ID = -1 'ค่าขนส่ง
                  BTI.BILL_TYPE_NAME = "ค่าขนส่งรถรับจ้าง"
                  BTI.BILL_TYPE_CODE = "01" 'FIX CODE ค่าขนส่งรถรับจ้าง
                  BTI.Flag = "N"
                  Call m_BillTransport.C_BillTransportItem.add(BTI, Trim(str(Di.WEIGHT_PER_PACK)) & "-" & Trim(BTI.BILL_TYPE_CODE))
               End If
           End If
         Next Di
         
         Dim TempCollection3 As Collection
         Dim TempDI As CDoItem
         Set TempCollection3 = New Collection
         For Each Di In TempCollection
             If Di.PART_ITEM_ID = -1 Then
                Set TempDI = GetObject("CDoItem", TempCollection3, Trim(str(Di.WEIGHT_PER_PACK)), False)
                If (TempDI Is Nothing) Then
                  Call TempCollection3.add(Di, Trim(str(Di.WEIGHT_PER_PACK)))
                End If
             End If
         Next Di
         
         For Each Di In TempCollection
           If Di.PART_ITEM_ID > 0 Then
               Set TempBTI = GetObject("CBillTransportItem", m_BillTransport.C_BillTransportItem, Trim(str(Di.WEIGHT_PER_PACK)) & "-" & Trim("02"), False)
               If Not (TempBTI Is Nothing) Then
                  TempBTI.PACK_AMOUNT = TempBTI.PACK_AMOUNT + Di.PACK_AMOUNT
               Else
                  Set BTI = New CBillTransportItem
                  BTI.PACK_AMOUNT = Di.PACK_AMOUNT
                  BTI.WEIGHT_PER_UNIT = Di.WEIGHT_PER_PACK
                  
                  Set TempDI = GetObject("CDoItem", TempCollection3, Trim(str(Di.WEIGHT_PER_PACK)), False)
                   If Not (TempDI Is Nothing) Then
                     BTI.RATE_PER_UNIT = TempDI.PRICE_PER_PACK
                   End If
                  
                  BTI.BILL_TRANSPORT_ITEM_ID = -1 'ค่าขนส่ง
                  BTI.BILL_TYPE_NAME = "ค่าขนส่งคิดลูกค้า"
                  BTI.BILL_TYPE_CODE = "02" 'FIX CODE ค่าขนส่งคิดลูกค้า
                  BTI.Flag = "N"
                  Call m_BillTransport.C_BillTransportItem.add(BTI, Trim(str(Di.WEIGHT_PER_PACK)) & "-" & Trim(BTI.BILL_TYPE_CODE))
               End If
           End If
         Next Di
         Else
            Set m_BillTransport = TempCollection2.Item(1)
            uctlSupplierTransport.MyCombo.ListIndex = IDToListIndex(uctlSupplierTransport.MyCombo, m_BillTransport.SUPPLIER_TRANSPORT_ID)
            txtAround.Text = m_BillTransport.AROUND
            txtShipping.Text = m_BillTransport.SHIPPING
            txtFixDetail.Text = m_BillTransport.FIX_RATE_DETAIL
            txtNoteTransport.Text = m_BillTransport.NOTE
            sscNotCalFlag.Value = FlagToCheck(m_BillTransport.NOT_CAL_VAT)
            Call GetValue(m_BillTransport.CAL_RATE_DELIVERY_TYPE)
           uctlDeliveryCusLookup.MyCombo.ListIndex = IDToListIndex(uctlDeliveryCusLookup.MyCombo, m_BillTransport.EX_DELIVERY_COST_ITEM_ID)
         End If
      End If
      
       
      
      GridEX1.ItemCount = CountItem(m_BillTransport.C_BillTransportItem)
      GridEX1.Rebind
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

   If Not VerifyTextControl(lblAround, txtAround, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblShipping, txtShipping, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblShippingDate, uctlShippingDate, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblSupplierTransport, uctlSupplierTransport.MyCombo, False) Then
      Exit Function
   End If
   
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_BillTransport.SUPPLIER_TRANSPORT_ID = uctlSupplierTransport.MyCombo.ItemData(Minus2Zero(uctlSupplierTransport.MyCombo.ListIndex))
   m_BillTransport.AROUND = Val(txtAround.Text)
   m_BillTransport.SHIPPING = txtShipping.Text
   m_BillTransport.SHIPPING_DATE = uctlShippingDate.ShowDate
   m_BillTransport.NOT_CAL_VAT = Check2Flag(sscNotCalFlag.Value)
   m_BillTransport.FIX_RATE_DETAIL = txtFixDetail.Text
   m_BillTransport.NOTE = txtNoteTransport.Text
   m_BillTransport.CAL_RATE_DELIVERY_TYPE = SetValue
   m_BillTransport.EX_DELIVERY_COST_ITEM_ID = uctlDeliveryCusLookup.MyCombo.ItemData(Minus2Zero(uctlDeliveryCusLookup.MyCombo.ListIndex))

   If Not glbDaily.AddEditBillTransport(m_BillTransport, IsOK, True, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call EnableForm(Me, True)
      Exit Function
   End If

'update ค่าขนส่งใน Doitem ด้วย
   Dim BillingDoc As CBillingDoc
   Dim DoItem As CDoItem
   Dim TempBTI As CBillTransportItem
   Dim TempBTI2 As CBillTransportItem
    Dim TempDI As CDoItem
    For Each DoItem In TempCollection
    If DoItem.PART_ITEM_ID = -1 Then
         For Each TempBTI In m_BillTransport.C_BillTransportItem
            If TempBTI.BILL_TYPE_CODE = "01" Then  'ถ้าเป็น ค่าขนส่งของรถรับจ้าง ให้รวมเอาทุกน้ำหนักเข้าด้วยกัน
                If DoItem.WEIGHT_PER_PACK = TempBTI.WEIGHT_PER_UNIT Then
                  DoItem.TRANSFER_WAGE = TempBTI.TOTAL_PRICE
                End If
            End If
            
         Next TempBTI
         Call DoItem.UpdateTransfer_Wage
'         Exit For

     End If
   Next DoItem
   
   Set BillingDoc = New CBillingDoc
   BillingDoc.BILLING_DOC_ID = BillingdocID
   Call BillingDoc.UpdateModify 'update คนที่แก้ไขเอกสารด้วย
         
   SaveData = True
End Function

Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
    
      Call LoadSupplierTransport(uctlSupplierTransport.MyCombo, m_SuppliersTransport)
      Set uctlSupplierTransport.MyCollection = m_SuppliersTransport
      
     If CUSTOMER_ID > 0 Then
         Call LoadDeliveryCus(uctlDeliveryCusLookup.MyCombo, m_DeliveryCus, CUSTOMER_ID) 'LOAD สถานที่จัดส่ง
         Set uctlDeliveryCusLookup.MyCollection = m_DeliveryCus
         
         uctlDeliveryCusLookup.MyCombo.ListIndex = IDToListIndex(uctlDeliveryCusLookup.MyCombo, DELIVERY_CUS_ITEM_ID)

      End If

        Call LoadExDeliveryCusItem(Nothing, m_ExDeliveryCostItem, , 3, DocumentDate) 'ส่ง 3 ไปดึงค่าขนส่งที่คิดให้รถรับจ้าง

       Call QueryData(True)
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
   
   Set m_Rs = New ADODB.Recordset
   Set m_SuppliersTransport = New Collection
   Set m_DeliveryCus = New Collection
   Set m_ExDeliveryCostItem = New Collection
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   
   Set m_Rs = Nothing
   Set m_SuppliersTransport = Nothing
   Set m_DeliveryCus = Nothing
   Set m_ExDeliveryCostItem = Nothing
End Sub

Private Sub txtDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtKeyName_Change()
   m_HasModify = True
End Sub




Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long
Dim I As Long

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"

      If RowIndex <= 0 Then
         Exit Sub
      End If
      
      Dim BTI As CBillTransportItem
      If m_BillTransport.C_BillTransportItem.Count <= 0 Then
         Exit Sub
      End If
      Set BTI = GetItem(m_BillTransport.C_BillTransportItem, RowIndex, RealIndex)
      I = 0
      
      I = I + 1
      Values(I) = BTI.BILL_TRANSPORT_ITEM_ID
      I = I + 1
      Values(I) = RealIndex
      I = I + 1
      Values(I) = BTI.BILL_TYPE_CODE
      I = I + 1
      Values(I) = BTI.BILL_TYPE_NAME
      I = I + 1
      Values(I) = BTI.WEIGHT_PER_UNIT
      I = I + 1
      Values(I) = FormatNumber(BTI.PACK_AMOUNT)
      I = I + 1
      Values(I) = BTI.RATE_PER_UNIT
      I = I + 1
      Values(I) = FormatNumber(BTI.TOTAL_PRICE)
      I = I + 1
      Values(I) = BTI.NOTE
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)

End Sub

Private Sub txtRateDriverTransport_Change()
   m_HasModify = True
End Sub

Private Sub txtRateDriverTransport_KeyPress(KeyAscii As Integer)
' If KeyAscii = 13 Then
'   Dim Di As CDoItem
'   Dim SumDI As Double
'   For Each Di In TempCollection
'     If Di.PART_ITEM_ID > 0 Then
'         SumDI = SumDI + (((Val(txtRateDriverTransport.Text) / 30) * Di.WEIGHT_PER_PACK) * Di.PACK_AMOUNT)
'     End If
'   Next Di
'   txtTransferWage.Text = SumDI
' End If
'
End Sub
Private Sub InitGrid1()
Dim Col As JSColumn
Dim I As Long

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.NAME = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 10
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 600
   Col.Caption = "ลำดับ"
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 600
   Col.TextAlignment = jgexAlignLeft
   Col.Caption = MapText("รหัส")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2400
   Col.TextAlignment = jgexAlignLeft
   Col.Caption = MapText("รายการ")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 1300
   Col.Caption = MapText("น้ำหนัก/ถุง")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 1300
   Col.Caption = MapText("จำนวน")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 1500
   Col.Caption = MapText("ราคา/หน่วย")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 1600
   Col.Caption = MapText("ราคา")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 2500
   Col.Caption = MapText("หมายเหตุ")

End Sub



Private Sub sscNotCalFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub ssoRound_Click(Value As Integer)
    m_HasModify = True
End Sub

Private Sub ssoVolume_Click(Value As Integer)
    m_HasModify = True
End Sub

Private Sub txtAround_Change()
    m_HasModify = True
End Sub

Private Sub txtAround_KeyPress(KeyAscii As Integer)
 KeyAscii = CheckIntAscii(KeyAscii)
End Sub

Private Sub txtFixDetail_Change()
   m_HasModify = True
End Sub

Private Sub txtNoteTransport_Change()
    m_HasModify = True
End Sub

Private Sub txtShipping_Change()
    m_HasModify = True
End Sub

Private Sub uctlDeliveryCusLookup_Change()
Dim TempStr As String
   DeliveryCusId = uctlDeliveryCusLookup.MyCombo.ItemData(Minus2Zero(uctlDeliveryCusLookup.MyCombo.ListIndex))
   txtShipping.Text = uctlDeliveryCusLookup.MyCombo.Text
   
   Set TempD = GetObject("CExWorksPrice", m_ExDeliveryCostItem, Trim(str(DeliveryCusId)) & "-" & Trim("30"), False)   'ค้นหาราคาค่าขนส่ง  ที่  กิโลเลย
   If TempD Is Nothing Then
      Set TempD = GetObject("CExWorksPrice", m_ExDeliveryCostItem, Trim(str(DeliveryCusId)) & "-" & Trim("1"), False)
      If TempD Is Nothing Then
          Set TempD = GetObject("CExWorksPrice", m_ExDeliveryCostItem, Trim(str(DeliveryCusId)) & "-" & Trim("999"), False)
         If TempD Is Nothing Then
              Set TempD = New CExWorksPrice
         End If
      End If
   End If
   txtFixDetail.Text = TempD.EX_WORKS_PRICE_CODE
   
   If Not TempD Is Nothing Then
     txtFixDetail.Text = TempD.EX_WORKS_PRICE_CODE
   Else
      Set TempD = GetObject("CExWorksPrice", m_ExDeliveryCostItem, Trim(str(DeliveryCusId)) & "-" & Trim("1"), False)
       If Not TempD Is Nothing Then
           txtFixDetail.Text = TempD.EX_WORKS_PRICE_CODE
      Else
         Set TempD = GetObject("CExWorksPrice", m_ExDeliveryCostItem, Trim(str(DeliveryCusId)) & "-" & Trim("999"), False)
          If Not TempD Is Nothing Then
           txtFixDetail.Text = TempD.EX_WORKS_PRICE_CODE
         End If
      End If
   End If
   
   m_HasModify = True
End Sub

Private Sub uctlShippingDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlSupplierTransport_Change()
Dim SupTransID As Long
Dim St As CSupplierTranSport
SupTransID = uctlSupplierTransport.MyCombo.ItemData(Minus2Zero(uctlSupplierTransport.MyCombo.ListIndex))
   If SupTransID > 0 Then
      Set St = GetSupplierTrans(m_SuppliersTransport, Trim(str(SupTransID)))
      txtSupplierCode.Text = St.SUPPLIER_CODE
      txtSupplierName.Text = St.SUPPLIER_NAME
   End If
   m_HasModify = True
End Sub
Private Function SetValue() As Long
   If ssoVolume.Value Then
      SetValue = 1
   ElseIf ssoRound.Value Then
      SetValue = 2
   Else
      SetValue = 1
   End If
End Function
Public Sub GetValue(ID As Long)
   If ID = 1 Then
      ssoVolume.Value = True
   ElseIf ID = 2 Then
      ssoRound.Value = True
   Else
      ssoVolume.Value = True
      ssoRound.Value = False
   End If
End Sub


