VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmInventoryWH 
   BackColor       =   &H80000000&
   ClientHeight    =   9510
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   20100
   Icon            =   "frmInventoryWH.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9510
   ScaleWidth      =   20100
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   9495
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   20115
      _ExtentX        =   35481
      _ExtentY        =   16748
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextLookup uctlPlaceLookup 
         Height          =   375
         Left            =   6450
         TabIndex        =   37
         Top             =   2520
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   661
      End
      Begin VB.ComboBox cboBalanceType 
         Height          =   315
         Left            =   6420
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   3000
         Width           =   2985
      End
      Begin VB.ComboBox cboPartType2 
         Height          =   315
         Left            =   6450
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   2040
         Width           =   2985
      End
      Begin prjFarmManagement.uctlDate uctlDateStock 
         Height          =   375
         Left            =   6450
         TabIndex        =   18
         Top             =   1080
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
      End
      Begin VB.ComboBox cboPartType 
         Height          =   315
         Left            =   6450
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1560
         Width           =   2985
      End
      Begin VB.ComboBox cboOrderType 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2520
         Width           =   2985
      End
      Begin VB.ComboBox cboOrderBy 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2040
         Width           =   2985
      End
      Begin prjFarmManagement.uctlTextBox txtPartName 
         Height          =   435
         Left            =   1860
         TabIndex        =   1
         Top             =   1560
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   19485
         _ExtentX        =   34369
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   4455
         Left            =   120
         TabIndex        =   7
         Top             =   3840
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   7858
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
         Column(1)       =   "frmInventoryWH.frx":27A2
         Column(2)       =   "frmInventoryWH.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmInventoryWH.frx":290E
         FormatStyle(2)  =   "frmInventoryWH.frx":2A6A
         FormatStyle(3)  =   "frmInventoryWH.frx":2B1A
         FormatStyle(4)  =   "frmInventoryWH.frx":2BCE
         FormatStyle(5)  =   "frmInventoryWH.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmInventoryWH.frx":2D5E
      End
      Begin prjFarmManagement.uctlTextBox txtPartNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   1080
         Width           =   2985
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtBalanceAmount 
         Height          =   435
         Left            =   11760
         TabIndex        =   31
         Top             =   1080
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtImportAmount 
         Height          =   435
         Left            =   11760
         TabIndex        =   32
         Top             =   1560
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtAdjustAmount 
         Height          =   435
         Left            =   11760
         TabIndex        =   33
         Top             =   2040
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtExportAmount 
         Height          =   435
         Left            =   15720
         TabIndex        =   34
         Top             =   1080
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTotalAmount 
         Height          =   435
         Left            =   15720
         TabIndex        =   35
         Top             =   1560
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTonAmount 
         Height          =   435
         Left            =   15720
         TabIndex        =   36
         Top             =   2040
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   767
      End
      Begin VB.Label lblPlaceLookup 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   4560
         TabIndex        =   38
         Top             =   2520
         Width           =   1755
      End
      Begin VB.Label lblTonAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblTonAmount"
         Height          =   435
         Left            =   13800
         TabIndex        =   30
         Top             =   2040
         Width           =   1755
      End
      Begin VB.Label lblTotalAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblTotalAmount"
         Height          =   435
         Left            =   13800
         TabIndex        =   29
         Top             =   1560
         Width           =   1755
      End
      Begin VB.Label lblExportAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblExportAmount"
         Height          =   435
         Left            =   13800
         TabIndex        =   28
         Top             =   1080
         Width           =   1755
      End
      Begin VB.Label lblAdjustAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblAdjustAmount"
         Height          =   435
         Left            =   9840
         TabIndex        =   27
         Top             =   2040
         Width           =   1755
      End
      Begin VB.Label lblImportAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblImportAmount"
         Height          =   435
         Left            =   9840
         TabIndex        =   26
         Top             =   1560
         Width           =   1755
      End
      Begin VB.Label lblBalanceAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblBalanceAmount"
         Height          =   435
         Left            =   9840
         TabIndex        =   25
         Top             =   1080
         Width           =   1755
      End
      Begin Threed.SSCommand cmdAdjust 
         Height          =   525
         Left            =   4560
         TabIndex        =   24
         Top             =   8430
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblBalanceType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   4560
         TabIndex        =   23
         Top             =   3000
         Width           =   1755
      End
      Begin VB.Label lblPartType2 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   4560
         TabIndex        =   21
         Top             =   2040
         Width           =   1755
      End
      Begin VB.Label lblDateStock 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   4560
         TabIndex        =   19
         Top             =   1080
         Width           =   1755
      End
      Begin VB.Label lblPartNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   17
         Top             =   1080
         Width           =   1755
      End
      Begin VB.Label lblPartType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   4560
         TabIndex        =   16
         Top             =   1560
         Width           =   1755
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   15
         Top             =   2520
         Width           =   1755
      End
      Begin VB.Label lblPartName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   360
         TabIndex        =   14
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   13
         Top             =   2040
         Width           =   1755
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   17910
         TabIndex        =   5
         Top             =   1050
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmInventoryWH.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   17910
         TabIndex        =   6
         Top             =   1620
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   525
         Left            =   6720
         TabIndex        =   8
         Top             =   8430
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmInventoryWH.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10095
         TabIndex        =   10
         Top             =   8430
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8445
         TabIndex        =   9
         Top             =   8430
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmInventoryWH.frx":356A
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmInventoryWH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_InventoryWHDoc As CInventoryWHDoc
Private m_TempInventoryWHDoc As CInventoryWHDoc
Private m_ColJobInOutWH As Collection
Private m_CollLotItemWh As Collection
Private m_CollLotItemWhImport As Collection
Private m_CollLotItemWhExport As Collection
Private m_CollLotItemWhImportBal As Collection
Private m_CollLotItemWhExportBal As Collection
Private m_Locations As Collection
Private m_PartTypes As Collection

Public m_LotItemWh As CLotItemWH

Private m_Rs As ADODB.Recordset
Private m_PartItem As CPartItem
Private m_PartItem2 As CPartItem
Private m_TempPartItem As CPartItem
Private mCollPartItem As Collection
Private m_TableName As String
Private DOCUMENT_TYPE As Long
Public JobDocType As Long
Public PartGroupID As Long
Public OKClick As Boolean
Public HeaderText As String

Private m_PartItems As Collection
Private Total(100) As Double
Private Sub cboPartType_Click()
   DOCUMENT_TYPE = cboPartType.ItemData(Minus2Zero(cboPartType.ListIndex))
   If DOCUMENT_TYPE = 13 Then
      cboPartType2.ListIndex = IDToListIndex(cboPartType2, 21)
   ElseIf DOCUMENT_TYPE = 14 Then
      cboPartType2.ListIndex = IDToListIndex(cboPartType2, 10)
   Else
      cboPartType2.ListIndex = IDToListIndex(cboPartType2, 0)
   End If
   
   If cboPartType.ListIndex = 1 Then
     uctlPlaceLookup.MyCombo.ListIndex = IDToListIndex(uctlPlaceLookup.MyCombo, 109)
   ElseIf cboPartType.ListIndex = 2 Then
    uctlPlaceLookup.MyCombo.ListIndex = IDToListIndex(uctlPlaceLookup.MyCombo, 110)
   Else
    uctlPlaceLookup.MyCombo.ListIndex = IDToListIndex(uctlPlaceLookup.MyCombo, 0)
   End If
End Sub



Private Sub cmdAdjust_Click()
   Load frmAdjustInventoryWH
   frmAdjustInventoryWH.Show 1
   
   OKClick = frmAdjustInventoryWH.OKClick
   txtPartNo.Text = frmAdjustInventoryWH.PartNo
   If frmAdjustInventoryWH.PartType = "10" Then
      cboPartType.ListIndex = 1
   ElseIf frmAdjustInventoryWH.PartType = "22" Then
      cboPartType.ListIndex = 2
   End If
   
   Unload frmAdjustInventoryWH
   Set frmAdjustInventoryWH = Nothing
   
   Call cmdSearch_Click
End Sub

Private Sub cmdClear_Click()
   txtPartName.Text = ""
   txtPartNo.Text = ""
   cboOrderBy.ListIndex = -1
   cboOrderType.ListIndex = -1
   cboPartType.ListIndex = -1
   cboPartType2.ListIndex = -1
   uctlDateStock.ShowDate = -1
End Sub
Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub

Private Sub cmdPrint_Click()
Dim Report As CReportInterface
Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
Dim ReportKey As String
Dim ReportFlag As Boolean
Dim Rc As CReportConfig
Dim iCount As Long
Dim EditMode As SHOW_MODE_TYPE
Dim ReportMode As Long

   ReportMode = 1
   
   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("ใบรายงานสินค้าคงเหลือ", "ปรับค่าหน้ากระดาษ")
   If lMenuChosen = 0 Then
      Exit Sub
   End If

   Call EnableForm(Me, False)

   If lMenuChosen = 1 Then
      ReportKey = "CReportInventoryWh"
      Set Report = New CReportInventoryWh
'      Report.TempColl (m_CollLotItemWh)
      ReportFlag = True
      Call Report.AddParam(1, "PREVIEW_TYPE")
   End If

   If Not Report Is Nothing Then
      Call Report.AddParam(DOCUMENT_TYPE, "DOCUMENT_TYPE")
      Call Report.AddParam(m_CollLotItemWh, "LOT_ITEM_WH")
      Call Report.AddParam(uctlDateStock.ShowDate, "FROM_DATE")
      Call Report.AddParam(lMenuChosen, "REPORT_TYPE")
      Call Report.AddParam(ReportKey, "REPORT_KEY")
   End If

   If ReportFlag Then
    frmReport.ClassName = ReportKey
      Set frmReport.ReportObject = Report

      frmReport.HeaderText = pnlHeader.Caption
      Load frmReport
      frmReport.Show 1

      Unload frmReport
      Set frmReport = Nothing
      Set Report = Nothing

   Else
   
   ReportKey = "CReportInventoryWh"
   ReportMode = 1
   
   Set Rc = New CReportConfig
   Rc.REPORT_KEY = ReportKey
   Call Rc.QueryData(m_Rs, iCount)
   
   If Not m_Rs.EOF Then
      Call Rc.PopulateFromRS(1, m_Rs)
      frmReportConfig.ShowMode = SHOW_EDIT
      frmReportConfig.ID = Rc.REPORT_CONFIG_ID
   Else
      frmReportConfig.ShowMode = SHOW_ADD
   End If
   
   frmReportConfig.ReportMode = ReportMode
   frmReportConfig.ReportKey = ReportKey
   frmReportConfig.HeaderText = HeaderText
   Load frmReportConfig
   frmReportConfig.Show 1
   
   Unload frmReportConfig
   Set frmReportConfig = Nothing
   
   Set Rc = Nothing
   End If

   Call EnableForm(Me, True)
End Sub

Private Sub cmdSearch_Click()
   cmdSearch.Enabled = False
   cmdClear.Enabled = False
   Call QueryData(True)
   cmdSearch.Enabled = True
   cmdClear.Enabled = True
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Call LoadLocation(uctlPlaceLookup.MyCombo, m_Locations, 2, , , 19)
      Set uctlPlaceLookup.MyCollection = m_Locations
      
      Call InitLoadPartType2(cboPartType)
      Call InitLoadBalanceType(cboBalanceType)
      cboPartType.ListIndex = 1

      Call InitGoodsOrderBy2(cboOrderBy)
      Call InitOrderType(cboOrderType)
      
      Call LoadPartType(cboPartType2)
   

      
''     Call QueryData(True)
   End If
End Sub
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long
Dim LWH As CLotItemWH
Dim TempLotItemWh As CLotItemWH
Dim I As Long
Dim Data1 As Double
Dim Data2 As Double
Dim Data3 As Double
Dim Data4 As Double
Dim Data5 As Double
Dim Key As String
Dim PartTypeID As Long
Dim balanceType As Long

   If Not VerifyCombo(lblPartType, cboPartType) Then
      Exit Sub
   End If

   If Flag Then
      Call EnableForm(Me, False)
      Set m_CollLotItemWh = Nothing
      Set m_CollLotItemWh = New Collection
      Set m_LotItemWh = Nothing
      Set m_LotItemWh = New CLotItemWH
      m_LotItemWh.FROM_DATE = -1
      m_LotItemWh.TO_DATE = uctlDateStock.ShowDate
      m_LotItemWh.PART_NO = PatchWildCard(txtPartNo.Text)
      m_LotItemWh.PART_DESC = PatchWildCard(txtPartName.Text)
      m_LotItemWh.LOCATION_ID = uctlPlaceLookup.MyCombo.ItemData(Minus2Zero(uctlPlaceLookup.MyCombo.ListIndex))
      m_LotItemWh.OrderBy = cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex))
      m_LotItemWh.OrderType = cboOrderType.ItemData(Minus2Zero(cboOrderType.ListIndex))
      DOCUMENT_TYPE = cboPartType.ItemData(Minus2Zero(cboPartType.ListIndex))
      balanceType = cboBalanceType.ItemData(Minus2Zero(cboBalanceType.ListIndex))
      
      m_LotItemWh.BALANCE_FLAG = "N"
      m_LotItemWh.CANCEL_FLAG = "N"
      m_LotItemWh.TX_TYPE = "I"
      m_LotItemWh.VERIFY_FLAG = "Y" 'จะดึงเฉพาะตัวที่ผ่านการตรวจสอบแล้วเท่านั้น
      If DOCUMENT_TYPE = 13 Then 'bulk
         Call InitGridBulk
         m_LotItemWh.DOCUMENT_TYPE_SET = "(13,16,18,21)"
      ElseIf DOCUMENT_TYPE = 14 Then 'bag
         Call InitGridBag
         m_LotItemWh.DOCUMENT_TYPE_SET = "(14,15,17,20)"
      End If

      If glbDaily.QueryLotItemWh(m_LotItemWh, m_Rs, ItemCount, IsOK, glbErrorLog) Then
           PartTypeID = cboPartType2.ItemData(Minus2Zero(cboPartType2.ListIndex))

            Call LoadPartStockFromLotItemWh(Nothing, m_CollLotItemWhImportBal, -1, uctlDateStock.ShowDate - 1, PartTypeID, , "I", 1, 3, DOCUMENT_TYPE, 2, m_LotItemWh.LOCATION_ID)  'ยอดรับเข้ายกมา'
            Call LoadPartStockFromLotItemWh(Nothing, m_CollLotItemWhExportBal, -1, uctlDateStock.ShowDate - 1, PartTypeID, , "E", 1, 6, DOCUMENT_TYPE, , m_LotItemWh.LOCATION_ID) 'ยอดจ่ายออกยกมา'
            'เอา m_CollLotItemWhImportBal-m_CollLotItemWhExportBal ก็จะได้ยอดยกมา
            
            Call LoadPartStockFromLotItemWh(Nothing, m_CollLotItemWhImport, uctlDateStock.ShowDate, uctlDateStock.ShowDate, PartTypeID, , "I", 1, 3, DOCUMENT_TYPE, 2, m_LotItemWh.LOCATION_ID) 'ยอดรับเข้าวันนี้
            Call LoadPartStockFromLotItemWh(Nothing, m_CollLotItemWhExport, uctlDateStock.ShowDate, uctlDateStock.ShowDate, PartTypeID, , "E", 1, 6, DOCUMENT_TYPE, , m_LotItemWh.LOCATION_ID) 'ยอดจ่ายออกวันนี้
      
      I = 0
         Set m_LotItemWh = Nothing
         While Not m_Rs.EOF
            I = I + 1
             Set m_LotItemWh = New CLotItemWH
            Call m_LotItemWh.PopulateFromRS(2, m_Rs)
            
            Set TempLotItemWh = New CLotItemWH
            TempLotItemWh.AddEditMode = SHOW_VIEW
            TempLotItemWh.PART_ITEM_ID = m_LotItemWh.PART_ITEM_ID
            TempLotItemWh.PART_NO = m_LotItemWh.PART_NO
            TempLotItemWh.BARCODE_NO = m_LotItemWh.BARCODE_NO
            TempLotItemWh.PART_DESC = m_LotItemWh.PART_DESC
            TempLotItemWh.DOCUMENT_TYPE = m_LotItemWh.DOCUMENT_TYPE
            TempLotItemWh.LOCATION_ID = m_LotItemWh.LOCATION_ID
           
            Set LWH = Nothing
            If DOCUMENT_TYPE = 14 Then 'Bag
               Key = Trim(str(m_LotItemWh.PART_ITEM_ID) & "-" & str(m_LotItemWh.WEIGHT_PER_PACK) & "-" & str(m_LotItemWh.LOT_DOC_ID) & "-" & str(m_LotItemWh.LOT_ID) & "-" & str(m_LotItemWh.BIN_NO) & "-" & str(m_LotItemWh.LOCK_NO) & "-" & "I" & "-" & str(m_LotItemWh.LOCATION_ID))
'                Key = Trim(str(m_LotItemWh.PART_ITEM_ID) & "-" & str(m_LotItemWh.WEIGHT_PER_PACK) & "-" & str(m_LotItemWh.LOT_DOC_ID) & "-" & str(m_LotItemWh.LOT_ID) & "-" & str(m_LotItemWh.BIN_NO) & "-" & str(m_LotItemWh.LOCK_NO) & "-" & "I")
            ElseIf DOCUMENT_TYPE = 13 Then 'Bulk
              Key = Trim(str(m_LotItemWh.PART_ITEM_ID) & "-" & str(m_LotItemWh.LOT_ID) & "-" & "I" & "-" & str(m_LotItemWh.LOT_DOC_ID) & "-" & str(m_LotItemWh.LOCATION_ID))
'              Key = Trim(str(m_LotItemWh.PART_ITEM_ID) & "-" & str(m_LotItemWh.LOT_ID) & "-" & "I" & "-" & str(m_LotItemWh.LOT_DOC_ID))
            End If
            
            Set LWH = GetObject("CLotItemWH", m_CollLotItemWhImportBal, Key, False)
            If Not LWH Is Nothing Then
               Data1 = LWH.CAPACITY_AMOUNT  'รับเข้ายกมา
'               TempLotItemWh.LOCATION_NAME_IN_BAL = LWH.LOCATION_NAME
            Else
               Data1 = 0
            End If
            
            Set LWH = Nothing
            If DOCUMENT_TYPE = 14 Then 'Bag
                 Key = Trim(str(m_LotItemWh.PART_ITEM_ID) & "-" & str(m_LotItemWh.WEIGHT_PER_PACK) & "-" & str(m_LotItemWh.LOT_DOC_ID) & "-" & str(m_LotItemWh.LOT_ID) & "-" & str(m_LotItemWh.BIN_NO) & "-" & str(m_LotItemWh.LOCK_NO) & "-" & "E" & "-" & str(m_LotItemWh.LOCATION_ID))
'                 Key = Trim(str(m_LotItemWh.PART_ITEM_ID) & "-" & str(m_LotItemWh.WEIGHT_PER_PACK) & "-" & str(m_LotItemWh.LOT_DOC_ID) & "-" & str(m_LotItemWh.LOT_ID) & "-" & str(m_LotItemWh.BIN_NO) & "-" & str(m_LotItemWh.LOCK_NO) & "-" & "E")
            ElseIf DOCUMENT_TYPE = 13 Then  'Bulk
                  Key = Trim(str(m_LotItemWh.PART_ITEM_ID) & "-" & str(m_LotItemWh.LOT_ID) & "-" & "E" & "-" & str(m_LotItemWh.LOT_DOC_ID) & "-" & str(m_LotItemWh.LOCATION_ID))
'                  Key = Trim(str(m_LotItemWh.PART_ITEM_ID) & "-" & str(m_LotItemWh.LOT_ID) & "-" & "E" & "-" & str(m_LotItemWh.LOT_DOC_ID))
            End If
            
            Set LWH = GetObject("CLotItemWH", m_CollLotItemWhExportBal, Key, False)
            If Not LWH Is Nothing Then
               Data2 = LWH.CAPACITY_AMOUNT 'จ่ายออกยกมา
'               TempLotItemWh.LOCATION_NAME_OUT_BAL = LWH.LOCATION_NAME
            Else
               Data2 = 0
            End If
            TempLotItemWh.BALANCE_AMOUNT = Data1 - Data2 'Abs(Data1 - Data2)
            
            Set LWH = Nothing
            If DOCUMENT_TYPE = 14 Then  'Bag
                 Key = Trim(str(m_LotItemWh.PART_ITEM_ID) & "-" & str(m_LotItemWh.WEIGHT_PER_PACK) & "-" & str(m_LotItemWh.LOT_DOC_ID) & "-" & str(m_LotItemWh.LOT_ID) & "-" & str(m_LotItemWh.BIN_NO) & "-" & str(m_LotItemWh.LOCK_NO) & "-" & "I" & "-" & str(m_LotItemWh.LOCATION_ID))
                '  Key = Trim(str(m_LotItemWh.PART_ITEM_ID) & "-" & str(m_LotItemWh.WEIGHT_PER_PACK) & "-" & str(m_LotItemWh.LOT_DOC_ID) & "-" & str(m_LotItemWh.LOT_ID) & "-" & str(m_LotItemWh.BIN_NO) & "-" & str(m_LotItemWh.LOCK_NO) & "-" & "I")
            ElseIf DOCUMENT_TYPE = 13 Then 'Bulk
               Key = Trim(str(m_LotItemWh.PART_ITEM_ID) & "-" & str(m_LotItemWh.LOT_ID) & "-" & "I" & "-" & str(m_LotItemWh.LOT_DOC_ID) & "-" & str(m_LotItemWh.LOCATION_ID))
'               Key = Trim(str(m_LotItemWh.PART_ITEM_ID) & "-" & str(m_LotItemWh.LOT_ID) & "-" & "I" & "-" & str(m_LotItemWh.LOT_DOC_ID))
            End If
           
            Set LWH = GetObject("CLotItemWH", m_CollLotItemWhImport, Key, False)
            If Not LWH Is Nothing Then
               Data3 = LWH.CAPACITY_AMOUNT
'               TempLotItemWh.LOCATION_NAME_IN = LWH.LOCATION_NAME
            Else
               Data3 = 0
            End If
            TempLotItemWh.IMPORT_AMOUNT = Data3 'รับเข้า
            
            Set LWH = Nothing
            If DOCUMENT_TYPE = 14 Then 'Bag
                 Key = Trim(str(m_LotItemWh.PART_ITEM_ID) & "-" & str(m_LotItemWh.WEIGHT_PER_PACK) & "-" & str(m_LotItemWh.LOT_DOC_ID) & "-" & str(m_LotItemWh.LOT_ID) & "-" & str(m_LotItemWh.BIN_NO) & "-" & str(m_LotItemWh.LOCK_NO) & "-" & "E" & "-" & str(m_LotItemWh.LOCATION_ID))
'                 Key = Trim(str(m_LotItemWh.PART_ITEM_ID) & "-" & str(m_LotItemWh.WEIGHT_PER_PACK) & "-" & str(m_LotItemWh.LOT_DOC_ID) & "-" & str(m_LotItemWh.LOT_ID) & "-" & str(m_LotItemWh.BIN_NO) & "-" & str(m_LotItemWh.LOCK_NO) & "-" & "E")
            ElseIf DOCUMENT_TYPE = 13 Then 'Bulk
               Key = Trim(str(m_LotItemWh.PART_ITEM_ID) & "-" & str(m_LotItemWh.LOT_ID) & "-" & "E" & "-" & str(m_LotItemWh.LOT_DOC_ID) & "-" & str(m_LotItemWh.LOCATION_ID))
'               Key = Trim(str(m_LotItemWh.PART_ITEM_ID) & "-" & str(m_LotItemWh.LOT_ID) & "-" & "E" & "-" & str(m_LotItemWh.LOT_DOC_ID))
            End If
            
            Set LWH = GetObject("CLotItemWH", m_CollLotItemWhExport, Key, False)
            If Not LWH Is Nothing Then
               Data4 = LWH.CAPACITY_AMOUNT
'               TempLotItemWh.LOCATION_NAME_OUT = LWH.LOCATION_NAME
            Else
               Data4 = 0
            End If
            TempLotItemWh.EXPORT_AMOUNT = Data4 'จ่ายออก
            
            Data5 = ((Data1 - Data2) + Data3) - Data4
            TempLotItemWh.ACTUAL_AMOUNT = Data5 'คงเหลือ
            TempLotItemWh.TOTAL_WEIGHT = Data5 * m_LotItemWh.WEIGHT_PER_PACK 'ยอดตัน
            TempLotItemWh.BILL_DESC = m_LotItemWh.BILL_DESC  'ประเภท
            If TempLotItemWh.DOCUMENT_TYPE = 15 Or TempLotItemWh.DOCUMENT_TYPE = 16 Then 'ปรับยอด
               TempLotItemWh.BL_START_DATE = m_LotItemWh.BL_START_DATE      'วันที่ผลิต
            Else 'ปกติ
               TempLotItemWh.START_DATE = m_LotItemWh.START_DATE     'วันที่ผลิต
            End If
'            TempLotItemWh.START_DATE = m_LotItemWh.START_DATE     'วันที่ผลิต
            TempLotItemWh.PACK_DATE = m_LotItemWh.PACK_DATE  'วันที่บรรจุ
            TempLotItemWh.TIME_PACK_BEGIN = m_LotItemWh.TIME_PACK_BEGIN 'เวลาเริ่มบรรจุ
            TempLotItemWh.TIME_PACK_END = m_LotItemWh.TIME_PACK_END 'เวลาบรรจุเสร็จ
            TempLotItemWh.LOT_NO = m_LotItemWh.LOT_NO   'Lot
            TempLotItemWh.BIN_NAME = m_LotItemWh.BIN_NAME   'ถังบรรจุ
            TempLotItemWh.LOCK_NAME = m_LotItemWh.LOCK_NAME    'ล๊อค
            TempLotItemWh.LOCATION_NAME = m_LotItemWh.LOCATION_NAME 'สถานที่จัดเก็บ
            TempLotItemWh.NOTE = m_LotItemWh.NOTE  'หมายเหตุ
            TempLotItemWh.WEIGHT_PER_PACK = m_LotItemWh.WEIGHT_PER_PACK
            
            If balanceType = 0 Then
               Call m_CollLotItemWh.add(TempLotItemWh)
            ElseIf balanceType = 1 Then
               If TempLotItemWh.BALANCE_AMOUNT > 0 Then
                  Call m_CollLotItemWh.add(TempLotItemWh)
               End If
            ElseIf balanceType = 2 Then
               If TempLotItemWh.IMPORT_AMOUNT > 0 Then
                  Call m_CollLotItemWh.add(TempLotItemWh)
               End If
            ElseIf balanceType = 3 Then
               If TempLotItemWh.EXPORT_AMOUNT > 0 Then
                  Call m_CollLotItemWh.add(TempLotItemWh)
               End If
            ElseIf balanceType = 4 Then
               If TempLotItemWh.ACTUAL_AMOUNT > 0 Then
                  Call m_CollLotItemWh.add(TempLotItemWh)
               End If
            ElseIf balanceType = 5 Then
                If TempLotItemWh.ACTUAL_AMOUNT <= 0 Then
                  Call m_CollLotItemWh.add(TempLotItemWh)
               End If
            ElseIf balanceType = 6 Then
                If TempLotItemWh.BALANCE_AMOUNT > 0 Or TempLotItemWh.IMPORT_AMOUNT > 0 Or TempLotItemWh.EXPORT_AMOUNT > 0 Or TempLotItemWh.ACTUAL_AMOUNT > 0 Then
                  Call m_CollLotItemWh.add(TempLotItemWh)
               End If
            End If

            Set TempLotItemWh = Nothing
            m_Rs.MoveNext
         Wend
   End If
   End If

   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   GridEX1.ItemCount = CountItem(m_CollLotItemWh)
   GridEX1.Rebind
   
   For I = 0 To UBound(Total)
      Total(I) = 0
   Next I
   
    For Each LWH In m_CollLotItemWh
      Total(1) = Total(1) + LWH.BALANCE_AMOUNT
      If LWH.DOCUMENT_TYPE = 15 Or LWH.DOCUMENT_TYPE = 16 Then       '  ปรับยอด Bag Bulk
         Total(3) = Total(3) + LWH.IMPORT_AMOUNT
      ElseIf DOCUMENT_TYPE = 13 Or DOCUMENT_TYPE = 14 Then
         Total(2) = Total(2) + LWH.IMPORT_AMOUNT
      End If
      Total(4) = Total(4) + LWH.EXPORT_AMOUNT
      Total(5) = Total(5) + LWH.ACTUAL_AMOUNT
     If DOCUMENT_TYPE = 14 Then
         Total(6) = Total(6) + MyDiffEx(LWH.ACTUAL_AMOUNT * LWH.WEIGHT_PER_PACK, 1000)
      ElseIf DOCUMENT_TYPE = 13 Then
          Total(6) = Total(6) + MyDiffEx(LWH.ACTUAL_AMOUNT, 1000)
      End If
      Next LWH
   
   txtBalanceAmount.Text = FormatNumber(Total(1))
   txtImportAmount.Text = FormatNumber(Total(2))
   txtAdjustAmount.Text = FormatNumber(Total(3))
   txtExportAmount.Text = FormatNumber(Total(4))
   txtTotalAmount.Text = FormatNumber(Total(5))
   txtTonAmount.Text = FormatNumber(Total(6))
   
   Call EnableForm(Me, True)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.NAME
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 116 Then
      Call cmdSearch_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
      Call cmdClear_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 118 Then
'     Call cmdAdd_Click
'      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
'      Call cmdDelete_Click
'      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
'      Call cmdEdit_Click
'      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
      Call cmdPrint_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 123 Then
      Call AddMemoNote
      KeyCode = 0
   End If
End Sub

Private Sub InitGridBag()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   
   GridEX1.FormatStyles.Clear
   Set fmsTemp = GridEX1.FormatStyles.add("N")
   fmsTemp.ForeColor = GLB_ALERT_COLOR
   
   Set Col = GridEX1.Columns.add '0
   Col.Width = 0
   Col.Caption = "ID"
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 600
   Col.Caption = "ลำดับ"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 2900
   Col.Caption = MapText("เบอร์สินค้า")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 1300
   Col.Caption = MapText("รหัสขาย")
      
   Set Col = GridEX1.Columns.add '4
   Col.Width = 2700
   Col.Caption = MapText("ชนิดสินค้า")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 1000
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ยกมา")
   
   Set Col = GridEX1.Columns.add '6
   Col.Width = 1000
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("รับเข้า")
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = 10
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ปรับยอด")
   
   Set Col = GridEX1.Columns.add '8
   Col.Width = 1000
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("จ่ายออก")
   
   Set Col = GridEX1.Columns.add '9
   Col.Width = 1000
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("คงเหลือ")
   
   Set Col = GridEX1.Columns.add '10
   Col.Width = 1000
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ยอดตัน")
   
   Set Col = GridEX1.Columns.add '11
   Col.Width = 2200
   Col.Caption = MapText("ประเภท")
   
   Set Col = GridEX1.Columns.add '12
   Col.Width = 1500
   Col.Caption = MapText("วันที่ผลิต")
   
   Set Col = GridEX1.Columns.add '12
   Col.Width = 2000
   Col.Caption = MapText("วันที่บรรจุ")
   
   Set Col = GridEX1.Columns.add '13
   Col.Width = 1900
   Col.Caption = MapText("LOT")
   
   Set Col = GridEX1.Columns.add '14
   Col.Width = 1000
   Col.Caption = MapText("ถังบรรจุ")
   
   Set Col = GridEX1.Columns.add '15
   Col.Width = 1000
   Col.Caption = MapText("ล๊อค")
   
   Set Col = GridEX1.Columns.add '16
   Col.Width = 1700
   Col.Caption = MapText("ที่จัดเก็บ")
   
   Set Col = GridEX1.Columns.add '18
   Col.Width = 3500
   Col.Caption = MapText("หมายเหตุ")
   
   GridEX1.ItemCount = 0
End Sub
Private Sub InitGridBulk()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   
   GridEX1.FormatStyles.Clear
   Set fmsTemp = GridEX1.FormatStyles.add("N")
   fmsTemp.ForeColor = GLB_ALERT_COLOR
   
   Set Col = GridEX1.Columns.add '0
   Col.Width = 0
   Col.Caption = "ID"
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 600
   Col.Caption = "ลำดับ"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 2900
   Col.Caption = MapText("เบอร์สินค้า")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 1300
   Col.Caption = MapText("รหัสขาย")
      
   Set Col = GridEX1.Columns.add '4
   Col.Width = 2700
   Col.Caption = MapText("ชนิดสินค้า")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 1500
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ยกมา")
   
   Set Col = GridEX1.Columns.add '6
   Col.Width = 1500
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("รับเข้า")
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = 10
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ปรับยอด")
   
   Set Col = GridEX1.Columns.add '8
   Col.Width = 1500
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("จ่ายออก")
   
   Set Col = GridEX1.Columns.add '9
   Col.Width = 1500
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("คงเหลือ")
   
   Set Col = GridEX1.Columns.add '10
   Col.Width = 1300
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ยอดตัน")
   
   Set Col = GridEX1.Columns.add '11
   Col.Width = 2200
   Col.Caption = MapText("ประเภท")
   
   Set Col = GridEX1.Columns.add '12
   Col.Width = 1500
   Col.Caption = MapText("วันที่ผลิต")
   
   Set Col = GridEX1.Columns.add '12
   Col.Width = 2000
   Col.Caption = MapText("วันที่บรรจุ")
   
   Set Col = GridEX1.Columns.add '13
   Col.Width = 1900
   Col.Caption = MapText("LOT")
   
   Set Col = GridEX1.Columns.add '14
   Col.Width = 1000
   Col.Caption = MapText("ถังบรรจุ")
   
   Set Col = GridEX1.Columns.add '15
   Col.Width = 1700
   Col.Caption = MapText("ที่จัดเก็บ")
   
   Set Col = GridEX1.Columns.add '16
   Col.Width = 3500
   Col.Caption = MapText("หมายเหตุ")
   
   GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = MapText("ข้อมูลสินค้าคงเหลือรายวัน")
   pnlHeader.Caption = MapText("ข้อมูลสินค้าคงเหลือรายวัน")
   
   Call InitGridBag
   
   Call InitNormalLabel(lblPartName, MapText("ชื่อสินค้า"))
   Call InitNormalLabel(lblDateStock, MapText("วันที่สต๊อก"))
   Call InitNormalLabel(lblPartType, MapText("ประเภทบรรจุ"))
   Call InitNormalLabel(lblPartType2, MapText("ชนิดสินค้า"))
   Call InitNormalLabel(lblPartNo, MapText("เบอร์สินค้า"))
   Call InitNormalLabel(lblOrderBy, MapText("เรียงตาม"))
   Call InitNormalLabel(lblOrderType, MapText("เรียงจาก"))
   Call InitNormalLabel(lblBalanceType, MapText("ประเภทคงเหลือ"))
   
   Call InitNormalLabel(lblPlaceLookup, MapText("สถานที่จัดเก็บ"))
   Call InitNormalLabel(lblBalanceAmount, MapText("รวมยอดยกมา"))
   Call InitNormalLabel(lblImportAmount, MapText("รวมยอดรับเข้า"))
   Call InitNormalLabel(lblAdjustAmount, MapText("รวมปรับยอด"))
   Call InitNormalLabel(lblExportAmount, MapText("รวมยอดจ่ายออก"))
   Call InitNormalLabel(lblTotalAmount, MapText("รวมยอดคงเหลือ"))
   Call InitNormalLabel(lblTonAmount, MapText("รวมยอดตัน"))

   Call txtPartName.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   Call txtBalanceAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.NAME_LEN)
   Call txtImportAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.NAME_LEN)
   Call txtAdjustAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.NAME_LEN)
   Call txtExportAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.NAME_LEN)
   Call txtTotalAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.NAME_LEN)
   Call txtTonAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.NAME_LEN)
  
   Call txtPartNo.SetKeySearch("PART_NO")

   Call InitCombo(cboPartType)
   Call InitCombo(cboPartType2)
   Call InitCombo(cboOrderBy)
   Call InitCombo(cboOrderType)
   Call InitCombo(cboBalanceType)
   
   
   uctlDateStock.ShowDate = Now
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdClear.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPrint.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdjust.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdPrint, MapText("พิมพ์"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา (F5)"))
   Call InitMainButton(cmdClear, MapText("เคลียร์ (F4)"))
   Call InitMainButton(cmdAdjust, MapText("คำนวณยอดคงเหลือ"))

End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   m_TableName = "USER_GROUP"

   Set m_InventoryWHDoc = New CInventoryWHDoc
   Set m_TempInventoryWHDoc = New CInventoryWHDoc
   Set m_CollLotItemWh = New Collection
   Set mCollPartItem = New Collection
   Set m_CollLotItemWhImport = New Collection
   Set m_CollLotItemWhExport = New Collection
   Set m_CollLotItemWhImportBal = New Collection
   Set m_CollLotItemWhExportBal = New Collection
   Set m_Locations = New Collection
   Set m_PartTypes = New Collection
   Set m_PartItem = New CPartItem
   Set m_PartItem2 = New CPartItem
   Set m_TempPartItem = New CPartItem
   Set m_LotItemWh = New CLotItemWH
   
   Set m_PartItems = New Collection

   Set m_Rs = New ADODB.Recordset

   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Terminate()
   Set m_CollLotItemWh = Nothing
   Set mCollPartItem = Nothing
   Set m_PartItem = Nothing
   Set m_PartItem2 = Nothing
   Set m_TempPartItem = Nothing
   Set m_CollLotItemWhImport = Nothing
   Set m_CollLotItemWhExport = Nothing
   Set m_CollLotItemWhImportBal = Nothing
   Set m_CollLotItemWhExportBal = Nothing
   Set m_Locations = Nothing
   Set m_PartTypes = Nothing
   Set m_LotItemWh = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub GridEX1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim oMenu As cPopupMenu
'Dim lMenuChosen As Long
'Dim TempID1 As Long
'Dim BD As CPartItem
'Dim IsOK As Boolean
'Dim OKClick As Boolean
'
'   If GridEX1.ItemCount <= 0 Then
'         Exit Sub
'   End If
'
'   TempID1 = GridEX1.Value(1)
'   If Button = 2 Then
'      Set oMenu = New cPopupMenu
'      lMenuChosen = oMenu.Popup("คัดลอกข้อมูล")
'      If lMenuChosen = 0 Then
'         Exit Sub
'      End If
'      Set oMenu = Nothing
'   Else
'      Exit Sub
'   End If
'
'   Call EnableForm(Me, False)
'   If lMenuChosen = 1 Then
'      Set BD = New CPartItem
'      BD.PART_ITEM_ID = TempID1
'      Call glbDaily.CopyPartItem(BD, IsOK, True, -1, glbErrorLog)
'      Call QueryData(True)
'      Set BD = Nothing
'   End If
'
'   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   RowBuffer.RowStyle = RowBuffer.Value(5)
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long
Dim LWH As CLotItemWH
   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"

   If RowIndex <= 0 Then
      Exit Sub
   End If
     RealIndex = RowIndex
   
      Set LWH = GetItem(m_CollLotItemWh, RowIndex, RealIndex)
      If LWH Is Nothing Then
         Exit Sub
      End If

      Values(1) = LWH.PART_ITEM_ID
      Values(2) = RealIndex
      Values(3) = LWH.PART_NO
      Values(4) = LWH.BARCODE_NO
      Values(5) = LWH.PART_DESC
      If DOCUMENT_TYPE = 13 Then
         Values(6) = FormatNumberToNull(LWH.BALANCE_AMOUNT, -1)
      ElseIf DOCUMENT_TYPE = 14 Then
         Values(6) = FormatNumberToNull(LWH.BALANCE_AMOUNT, 0)
      End If

      If LWH.DOCUMENT_TYPE = 15 Then      '  ปรับยอด Bag
         Values(7) = "" 'รับเข้า
         Values(8) = FormatNumberToNull(LWH.IMPORT_AMOUNT, 0)  ')'ปรับยอด
         Values(9) = FormatNumberToNull(LWH.EXPORT_AMOUNT, -1) 'จ่ายออก
     ElseIf LWH.DOCUMENT_TYPE = 16 Then       '  ปรับยอด Bulk
         Values(7) = "" 'รับเข้า
         Values(8) = FormatNumberToNull(LWH.IMPORT_AMOUNT, -1)  ')'ปรับยอด
         Values(9) = FormatNumberToNull(LWH.EXPORT_AMOUNT, -1) 'จ่ายออก
      ElseIf DOCUMENT_TYPE = 13 Then
         Values(7) = FormatNumberToNull(LWH.IMPORT_AMOUNT, -1)  'รับเข้า
         Values(8) = "" 'ปรับยอด
         Values(9) = FormatNumberToNull(LWH.EXPORT_AMOUNT, -1) 'จ่ายออก
    ElseIf DOCUMENT_TYPE = 14 Then
         Values(7) = FormatNumberToNull(LWH.IMPORT_AMOUNT, 0)  'รับเข้า
         Values(8) = "" 'ปรับยอด
         Values(9) = FormatNumberToNull(LWH.EXPORT_AMOUNT, 0) 'จ่ายออก
      End If
      
'      Values(9) = FormatNumberToNull(LWH.EXPORT_AMOUNT, 0) 'จ่ายออก
     If DOCUMENT_TYPE = 14 Then
         Values(10) = FormatNumberToNull(LWH.ACTUAL_AMOUNT, 0)  '"" 'คงเหลือ
         Values(11) = FormatNumberToNull(MyDiffEx(LWH.ACTUAL_AMOUNT * LWH.WEIGHT_PER_PACK, 1000), 2) '"" 'ยอดตัน
      ElseIf DOCUMENT_TYPE = 13 Then
          Values(10) = FormatNumberToNull(LWH.ACTUAL_AMOUNT, -1)  '"" 'คงเหลือ
          Values(11) = FormatNumberToNull(MyDiffEx(LWH.ACTUAL_AMOUNT, 1000), 3) '"" 'ยอดตัน
      End If
      
      Values(12) = LWH.BILL_DESC  'ประเภท
      
      If LWH.DOCUMENT_TYPE = 15 Or LWH.DOCUMENT_TYPE = 16 Then     '  ปรับยอด Bag,Bulk
        Values(13) = DateToStringExtEx2(LWH.BL_START_DATE)    'วันที่ผลิต
      Else
         Values(13) = DateToStringExtEx2(LWH.START_DATE)   'วันที่ผลิต
      End If
      Values(14) = DateToStringExtEx2(LWH.PACK_DATE) & " " & Format(LWH.TIME_PACK_BEGIN, "HH:mm")  'วันที่บรรจุ
      
      Values(15) = LWH.LOT_NO   'Lot
      Values(16) = LWH.BIN_NAME   'ถังบรรจุ
      If DOCUMENT_TYPE = 13 Then
         Values(17) = LWH.LOCATION_NAME     'ที่จัดเก็บ
         Values(18) = LWH.NOTE 'หมายเหตุ
      Else
         Values(17) = LWH.LOCK_NAME   'ล๊อค
         Values(18) = LWH.LOCATION_NAME     'ที่จัดเก็บ
         Values(19) = LWH.NOTE 'หมายเหตุ
      End If
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.HEIGHT = ScaleHeight
   pnlHeader.Width = ScaleWidth
   GridEX1.Width = ScaleWidth - 2 * GridEX1.Left
   GridEX1.HEIGHT = ScaleHeight - GridEX1.Top - 620
   cmdAdjust.Top = ScaleHeight - 580
   cmdPrint.Top = ScaleHeight - 580
   cmdOK.Top = ScaleHeight - 580
   cmdExit.Top = ScaleHeight - 580
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
   cmdPrint.Left = cmdOK.Left - cmdExit.Width - 50
   cmdAdjust.Left = cmdPrint.Left - cmdAdjust.Width - 50
End Sub


