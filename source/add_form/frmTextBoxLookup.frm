VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmTextBoxLookup 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   Icon            =   "frmTextBoxLookup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   7320
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   6855
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   12091
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   15
         TabIndex        =   3
         Top             =   0
         Width           =   7350
         _ExtentX        =   12965
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   4605
         Left            =   60
         TabIndex        =   1
         Top             =   2190
         Width           =   7185
         _ExtentX        =   12674
         _ExtentY        =   8123
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
         HeaderFontName  =   "AngsanaUPC"
         HeaderFontBold  =   -1  'True
         HeaderFontSize  =   14.25
         HeaderFontWeight=   700
         FontSize        =   9.75
         BackColorBkg    =   16777215
         ColumnHeaderHeight=   480
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmTextBoxLookup.frx":27A2
         Column(2)       =   "frmTextBoxLookup.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmTextBoxLookup.frx":290E
         FormatStyle(2)  =   "frmTextBoxLookup.frx":2A6A
         FormatStyle(3)  =   "frmTextBoxLookup.frx":2B1A
         FormatStyle(4)  =   "frmTextBoxLookup.frx":2BCE
         FormatStyle(5)  =   "frmTextBoxLookup.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmTextBoxLookup.frx":2D5E
      End
      Begin prjFarmManagement.uctlTextBox txtSearchText 
         Height          =   435
         Left            =   1680
         TabIndex        =   0
         Top             =   840
         Width           =   2265
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtSearchName 
         Height          =   435
         Left            =   1650
         TabIndex        =   5
         Top             =   1320
         Width           =   5025
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin VB.Label lblSearchName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   0
         TabIndex        =   6
         Top             =   1380
         Width           =   1635
      End
      Begin VB.Label lblSearchText 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   4
         Top             =   900
         Width           =   1635
      End
   End
End
Attribute VB_Name = "frmTextBoxLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_Rs As ADODB.Recordset
Public KEYWORD As String
Public KeySearch As String
Public KeyId As Long
Public SearchType As Long

Public OKClick As Boolean
Public HeaderText As String
Private m_Customer As CCustomer
Private m_Employee As CEmployee
Private m_PartItem As CPartItem
Private m_PartMasterItem As CPartMaster
Private m_Supplier As CSupplier
Private m_FreeLance As CFreelance
Private m_ExWorksPrice As CExWorksPrice
Private m_TruckNo As CSupplierTranSport
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      
      Call RunQuery
   End If
End Sub
Private Sub QueryCustomer()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long
      
   Call EnableForm(Me, False)
         
   Dim m_Customer As CCustomer
   Set m_Customer = New CCustomer
   
   m_Customer.CUSTOMER_ID = -1
   m_Customer.CUSTOMER_CODE = PatchWildCard(txtSearchText.Text)
   m_Customer.CUSTOMER_NAME = PatchWildCard(txtSearchName.Text)
   
   m_Customer.OrderType = 1
   Call m_Customer.QueryData5(m_Rs, ItemCount)
   
   Call InitGrid
   
   GridEX1.ItemCount = ItemCount
   GridEX1.Rebind
   
   Call EnableForm(Me, True)
End Sub
Private Sub QueryEmployee()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long
      
   Call EnableForm(Me, False)
         
   Dim m_Employee As CEmployee
   Set m_Employee = New CEmployee
   
  m_Employee.EMP_ID = -1
  m_Employee.EMP_CODE = PatchWildCard(txtSearchText.Text)
  m_Employee.EMP_NAME = PatchWildCard(txtSearchName.Text)
  m_Employee.EMP_RESIGN_FLAG = "N"
   m_Employee.OrderType = 1
   Call m_Employee.QueryData5(m_Rs, ItemCount)
   
   Call InitGrid
   
   GridEX1.ItemCount = ItemCount
   GridEX1.Rebind
   
   Call EnableForm(Me, True)
End Sub
Private Sub QueryFreeLance()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long
      
   Call EnableForm(Me, False)
         
   Dim m_FreeLance As CFreelance
   Set m_FreeLance = New CFreelance
   
   m_FreeLance.FREELANCE_ID = -1
   m_FreeLance.FREELANCE_CODE = PatchWildCard(txtSearchText.Text)
   m_FreeLance.FREELANCE_NAME = PatchWildCard(txtSearchName.Text)
   m_FreeLance.FREELANCE_RESIGN_FLAG = "N"
   m_FreeLance.OrderType = 1
   Call m_FreeLance.QueryData(m_Rs, ItemCount)
   
   Call InitGrid
   
   GridEX1.ItemCount = ItemCount
   GridEX1.Rebind
   
   Call EnableForm(Me, True)
End Sub

Private Sub QueryPartItem()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long
      
   Call EnableForm(Me, False)
         
   Dim m_PartItem As CPartItem
   Set m_PartItem = New CPartItem
   
   m_PartItem.PART_ITEM_ID = -1
   m_PartItem.PART_NO = PatchWildCard(txtSearchText.Text)
   m_PartItem.PART_DESC = PatchWildCard(txtSearchName.Text)
   m_PartItem.OrderType = 1
   m_PartItem.CANCEL_FLAG = "N"
   Call m_PartItem.QueryData2(101, m_Rs, ItemCount)
   
   Call InitGrid
   
   GridEX1.ItemCount = ItemCount
   GridEX1.Rebind
   
   Call EnableForm(Me, True)
End Sub
Private Sub QueryPartMasterItem()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long
      
   Call EnableForm(Me, False)
         
   Dim m_PartMasterItem As CPartMaster
   Set m_PartMasterItem = New CPartMaster
   
   m_PartMasterItem.PART_MASTER_ID = -1
   m_PartMasterItem.PART_MASTER_NO = PatchWildCard(txtSearchText.Text)
   m_PartMasterItem.PART_MASTER_NAME = PatchWildCard(txtSearchName.Text)
   m_PartMasterItem.OrderType = 1
   m_PartMasterItem.CANCEL_FLAG = "N"
   Call m_PartMasterItem.QueryData(1, m_Rs, ItemCount)
   
   Call InitGrid
   
   GridEX1.ItemCount = ItemCount
   GridEX1.Rebind
   
   Call EnableForm(Me, True)
End Sub
Private Sub QuerySupplier(Optional SupplierType As Long = -1)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long
      
   Call EnableForm(Me, False)
         
   Dim m_Supplier As CSupplier
   Set m_Supplier = New CSupplier
   
   m_Supplier.SUPPLIER_ID = -1
   m_Supplier.SUPPLIER_CODE = PatchWildCard(txtSearchText.Text)
   m_Supplier.SUPPLIER_NAME = PatchWildCard(txtSearchName.Text)
   
   m_Supplier.SUPPLIER_TYPE = SupplierType
   m_Supplier.OrderType = 1
   Call m_Supplier.QueryData2(m_Rs, ItemCount)
   
   Call InitGrid
   
   GridEX1.ItemCount = ItemCount
   GridEX1.Rebind
   
   Call EnableForm(Me, True)
End Sub

Private Sub QueryExWorksPrice()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long
      
   Call EnableForm(Me, False)
         
   Dim m_ExWorksPrice As CExWorksPrice
   Set m_ExWorksPrice = New CExWorksPrice
   
   m_ExWorksPrice.EX_WORKS_PRICE_ID = -1
   m_ExWorksPrice.EX_WORKS_PRICE_CODE = PatchWildCard(txtSearchText.Text)
   m_ExWorksPrice.EX_WORKS_PRICE_DESC = PatchWildCard(txtSearchName.Text)
   
   m_ExWorksPrice.OrderType = 1
   Call m_ExWorksPrice.QueryData(1, m_Rs, ItemCount)
   
   Call InitGrid
   
   GridEX1.ItemCount = ItemCount
   GridEX1.Rebind
   
   Call EnableForm(Me, True)
End Sub
Private Sub QueryTruckNo()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long
      
   Call EnableForm(Me, False)
         
   Dim m_TruckNo As CSupplierTranSport
   Set m_TruckNo = New CSupplierTranSport
   
   m_TruckNo.SUPPLIER_TRANSPORT_ID = -1
   m_TruckNo.SUPPLIER_TRANSPORT_CODE = PatchWildCard(txtSearchText.Text)
   m_TruckNo.SUPPLIER_TRANSPORT_DETAIL = PatchWildCard(txtSearchName.Text)
   
   m_TruckNo.OrderType = 1
   Call m_TruckNo.QueryData(1, m_Rs, ItemCount)
   
   Call InitGrid
   
   GridEX1.ItemCount = ItemCount
   GridEX1.Rebind
   
   Call EnableForm(Me, True)
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.NAME
      glbErrorLog.ShowUserError
      KeyCode = 0
   End If
End Sub
Private Sub InitGrid()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   
   GridEX1.FormatStyles.Clear
   Set fmsTemp = GridEX1.FormatStyles.add("N")
   fmsTemp.ForeColor = GLB_ALERT_COLOR
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 2700
   Col.Caption = MapText("รหัส")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = ScaleWidth - 2300
   Col.Caption = MapText("รายละเอียด")
   
   GridEX1.ItemCount = 0
End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   Me.Caption = HeaderText
   
   Call InitGrid
   
   Call InitNormalLabel(lblSearchText, MapText("รหัส"))
   Call InitNormalLabel(lblSearchName, MapText("รายละเอียด"))
      
   txtSearchText.Text = KEYWORD
   txtSearchText.Enabled = False
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
End Sub
Private Sub Form_Load()
   m_HasActivate = False
   
   Set m_Rs = New ADODB.Recordset
   
   Set m_Customer = New CCustomer
   Set m_Employee = New CEmployee
   Set m_PartItem = New CPartItem
   Set m_PartMasterItem = New CPartMaster
   Set m_Supplier = New CSupplier
   Set m_FreeLance = New CFreelance
   Set m_ExWorksPrice = New CExWorksPrice
   Set m_TruckNo = New CSupplierTranSport
   
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set m_Customer = Nothing
   Set m_Employee = Nothing
   Set m_PartItem = Nothing
   Set m_PartMasterItem = Nothing
   Set m_Supplier = Nothing
   Set m_FreeLance = Nothing
   Set m_ExWorksPrice = Nothing
   Set m_TruckNo = Nothing
   
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
End Sub

Private Sub GridEX1_DblClick()
   Call ReturnKeyWord
End Sub
Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long
   
   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"

   If m_Rs Is Nothing Then
      Exit Sub
   End If

   If m_Rs.State <> adStateOpen Then
      Exit Sub
   End If

   If m_Rs.EOF Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If
   
   Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
   
   If KeySearch = "CUSTOMER_CODE" Then
      Call m_Customer.PopulateFromRS(3, m_Rs)
      Values(1) = m_Customer.CUSTOMER_ID
      Values(2) = m_Customer.CUSTOMER_CODE
      Values(3) = m_Customer.CUSTOMER_NAME
   ElseIf KeySearch = "EMP_CODE" Then
      Call m_Employee.PopulateFromRS(2, m_Rs)
      Values(1) = m_Employee.EMP_ID
      Values(2) = m_Employee.EMP_CODE
       Values(3) = m_Employee.EMP_NAME & "  " & m_Employee.LAST_NAME
   ElseIf KeySearch = "PART_NO" Then
      Call m_PartItem.PopulateFromRS2(101, m_Rs)
      Values(1) = m_PartItem.PART_ITEM_ID
      Values(2) = m_PartItem.PART_NO
       Values(3) = m_PartItem.PART_DESC
   ElseIf KeySearch = "PART_MASTER_NO" Then
      Call m_PartMasterItem.PopulateFromRS(1, m_Rs)
      Values(1) = m_PartMasterItem.PART_MASTER_ID
      Values(2) = m_PartMasterItem.PART_MASTER_NO
       Values(3) = m_PartMasterItem.PART_MASTER_NAME
   ElseIf KeySearch = "SUPPLIER_CODE" Then
      Call m_Supplier.PopulateFromRS(3, m_Rs)
      Values(1) = m_Supplier.SUPPLIER_ID
      Values(2) = m_Supplier.SUPPLIER_CODE
      Values(3) = m_Supplier.SUPPLIER_NAME
   ElseIf KeySearch = "FREELANCE_CODE" Then
      Call m_FreeLance.PopulateFromRS(1, m_Rs)
      Values(1) = m_FreeLance.FREELANCE_ID
      Values(2) = m_FreeLance.FREELANCE_CODE
      Values(3) = m_FreeLance.FREELANCE_NAME
   ElseIf KeySearch = "WORKS_PRICE_CODE" Then
      Call m_ExWorksPrice.PopulateFromRS(1, m_Rs)
      Values(1) = m_ExWorksPrice.EX_WORKS_PRICE_ID
      Values(2) = m_ExWorksPrice.EX_WORKS_PRICE_CODE
      Values(3) = m_ExWorksPrice.EX_WORKS_PRICE_DESC
   ElseIf KeySearch = "TRUCK_NO" Then
      Call m_TruckNo.PopulateFromRS(1, m_Rs)
      Values(1) = m_TruckNo.SUPPLIER_TRANSPORT_ID
      Values(2) = m_TruckNo.SUPPLIER_TRANSPORT_CODE
      Values(3) = m_TruckNo.SUPPLIER_TRANSPORT_DETAIL
   ElseIf KeySearch = "SUPPLIER_CODE_TRANSPORT" Then
      Call m_Supplier.PopulateFromRS(3, m_Rs)
      Values(1) = m_Supplier.SUPPLIER_ID
      Values(2) = m_Supplier.SUPPLIER_CODE
      Values(3) = m_Supplier.SUPPLIER_NAME
   End If
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub GridEX1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = DUMMY_KEY Then
      KeyCode = 0
      Unload Me
   ElseIf KeyCode = 13 Or KeyCode = 32 Then
      Call ReturnKeyWord
   End If
End Sub
Private Sub ReturnKeyWord()
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
'   If KeySearch = 1 Then
'      KeyId = GridEX1.Value(1)
'
'         frmAddEditPartItemPicture.ID = 1
''         Set frmAddEditPartItemPicture.ParentForm = Me
'         Set frmAddEditPartItemPicture.TempCollection = m_PartItem.Pictures
'         frmAddEditPartItemPicture.ShowMode = SHOW_EDIT
'         frmAddEditPartItemPicture.PictureType = HEAD_PART
'         frmAddEditPartItemPicture.HeaderText = MapText("แก้ไข ") & PictureTypeToText(HEAD_PART)
'         Load frmAddEditPartItemPicture
'         frmAddEditPartItemPicture.Show 1
'
'      OKClick = frmAddEditPartItemPicture.OKClick
'
'      Unload frmAddEditPartItemPicture
'      Set frmAddEditPartItemPicture = Nothing
'   Else
      KEYWORD = GridEX1.Value(2)
      Unload Me
'   End If
End Sub
Private Sub RunQuery()
   If KeySearch = "CUSTOMER_CODE" Then
      Call QueryCustomer
   ElseIf KeySearch = "EMP_CODE" Then
      Call QueryEmployee
   ElseIf KeySearch = "PART_NO" Then
      Call QueryPartItem
   ElseIf KeySearch = "PART_MASTER_NO" Then
      Call QueryPartMasterItem
   ElseIf KeySearch = "SUPPLIER_CODE" Then
      Call QuerySupplier
   ElseIf KeySearch = "FREELANCE_CODE" Then
      Call QueryFreeLance
   ElseIf KeySearch = "WORKS_PRICE_CODE" Then
      Call QueryExWorksPrice
   ElseIf KeySearch = "TRUCK_NO" Then
      Call QueryTruckNo
   ElseIf KeySearch = "SUPPLIER_CODE_TRANSPORT" Then
      Call QuerySupplier(19)
   End If
End Sub

Private Sub txtSearchName_Change()
   Call RunQuery
End Sub
