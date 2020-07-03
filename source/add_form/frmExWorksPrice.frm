VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmExWorksPrice 
   BackColor       =   &H80000000&
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   11910
   Icon            =   "frmExWorksPrice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboOrderType 
         Height          =   315
         Left            =   6450
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1620
         Width           =   2985
      End
      Begin VB.ComboBox cboOrderBy 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1620
         Width           =   2985
      End
      Begin prjFarmManagement.uctlTextBox txtDesc 
         Height          =   435
         Left            =   6450
         TabIndex        =   1
         Top             =   1080
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   5175
         Left            =   180
         TabIndex        =   7
         Top             =   2520
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   9128
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
         Column(1)       =   "frmExWorksPrice.frx":27A2
         Column(2)       =   "frmExWorksPrice.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmExWorksPrice.frx":290E
         FormatStyle(2)  =   "frmExWorksPrice.frx":2A6A
         FormatStyle(3)  =   "frmExWorksPrice.frx":2B1A
         FormatStyle(4)  =   "frmExWorksPrice.frx":2BCE
         FormatStyle(5)  =   "frmExWorksPrice.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmExWorksPrice.frx":2D5E
      End
      Begin prjFarmManagement.uctlTextBox txtPackageCode 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   1080
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   767
      End
      Begin Threed.SSCommand cmdImport 
         Height          =   525
         Left            =   6720
         TabIndex        =   19
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmExWorksPrice.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCheck chkMarket 
         Height          =   375
         Left            =   1860
         TabIndex        =   2
         Top             =   2040
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblPackageCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   18
         Top             =   1140
         Width           =   1755
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4980
         TabIndex        =   17
         Top             =   1680
         Width           =   1365
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4080
         TabIndex        =   16
         Top             =   1140
         Width           =   2295
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   15
         Top             =   1680
         Width           =   1755
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   10110
         TabIndex        =   5
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmExWorksPrice.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   10110
         TabIndex        =   6
         Top             =   1680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   10
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmExWorksPrice.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   8
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmExWorksPrice.frx":3884
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   9
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10095
         TabIndex        =   12
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8445
         TabIndex        =   11
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmExWorksPrice.frx":3B9E
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmExWorksPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_ExWorksPrice As CExWorksPrice
Private m_TempExWorksPrice  As CExWorksPrice
Private m_Rs As ADODB.Recordset
Private m_Rs2 As ADODB.Recordset
Private m_TableName As String
Dim ItemCount As Long
Public Area As Long

Public OKClick As Boolean
Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean
If Area = 1 Then
   If Not VerifyAccessRight("PACKAGE-CENTER_EX-WORKS-PRICE_ADD", "เพิ่มข้อมูลราคาสินค้า") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   frmAddEditExWorksPrice.HeaderText = MapText("เพิ่มข้อมูลราคาสินค้า")
ElseIf Area = 2 Then
   If Not VerifyAccessRight("PACKAGE-CENTER_DELIVERY-COST_ADD", "เพิ่มข้อมูลราคาค่าขนส่ง") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   frmAddEditExWorksPrice.HeaderText = MapText("เพิ่มข้อมูลราคาค่าขนส่ง")
ElseIf Area = 3 Then
   If Not VerifyAccessRight("PACKAGE-CENTER_PROMOTION-PART_ADD", "เพิ่มข้อมูลราคาส่วนลดค่าสินค้า(หน้าบิล)") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   frmAddEditExWorksPrice.HeaderText = MapText("เพิ่มข้อมูลราคาส่วนลดค่าสินค้า(หน้าบิล)")
ElseIf Area = 4 Then
   If Not VerifyAccessRight("PACKAGE-CENTER_PROMOTION-DELIVERY_ADD", "เพิ่มข้อมูลราคาส่วนลดค่าขนส่ง") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   frmAddEditExWorksPrice.HeaderText = MapText("เพิ่มข้อมูลราคาส่วนลดค่าขนส่ง")
ElseIf Area = 5 Then
   If Not VerifyAccessRight("PACKAGE-CENTER_PROMOTION-PART-EXTRA_ADD", "เพิ่มข้อมูลราคาส่วนลดพิเศษค่าสินค้า(หลังบิล)") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   frmAddEditExWorksPrice.HeaderText = MapText("เพิ่มข้อมูลราคาส่วนลดพิเศษค่าสินค้า(หลังบิล)")
End If
  frmAddEditExWorksPrice.Area = Area
   frmAddEditExWorksPrice.ShowMode = SHOW_ADD
   Load frmAddEditExWorksPrice
   frmAddEditExWorksPrice.Show 1
   
   OKClick = frmAddEditExWorksPrice.OKClick
   
   Unload frmAddEditExWorksPrice
   Set frmAddEditExWorksPrice = Nothing
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub

Private Sub cmdClear_Click()
   txtDesc.Text = ""
   txtPackageCode.Text = ""
   cboOrderBy.ListIndex = -1
   cboOrderType.ListIndex = -1
End Sub

Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim id As Long
  If Area = 1 Then
      If Not VerifyAccessRight("PACKAGE-CENTER_EX-WORKS-PRICE_DELETE", "ลบข้อมูลราคาสินค้า") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf Area = 2 Then
      If Not VerifyAccessRight("PACKAGE-CENTER_DELIVERY-COST_DELETE", "ลบข้อมูลราคาค่าขนส่ง") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf Area = 3 Then
      If Not VerifyAccessRight("PACKAGE-CENTER_PROMOTION-PART_DELETE", "ลบข้อมูลราคาส่วนลดค่าสินค้า(หน้าบิล)") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf Area = 4 Then
      If Not VerifyAccessRight("PACKAGE-CENTER_PROMOTION-DELIVERY_DELETE", "ลบข้อมูลโปรโมชั่นราคาค่าขนส่ง") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf Area = 5 Then
      If Not VerifyAccessRight("PACKAGE-CENTER_PROMOTION-PART-EXTRA_DELETE", "ลบข้อมูลราคาส่วนลดพิเศษค่าสินค้า(หลังบิล)") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   id = GridEX1.Value(1)
   
   Call glbDatabaseMngr.LockTable(m_TableName, id, IsCanLock, glbErrorLog)
   If Not ConfirmDelete(GridEX1.Value(2)) Then
      Call glbDatabaseMngr.UnLockTable(m_TableName, id, IsCanLock, glbErrorLog)
      Exit Sub
   End If

   Call EnableForm(Me, False)
   If Not glbDaily.DeleteExWorksPrice(id, IsOK, True, glbErrorLog) Then
      m_ExWorksPrice.EX_WORKS_PRICE_ID = -1
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call glbDatabaseMngr.UnLockTable(m_TableName, id, IsCanLock, glbErrorLog)
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call QueryData(True)
   
   Call glbDatabaseMngr.UnLockTable(m_TableName, id, IsCanLock, glbErrorLog)
   Call EnableForm(Me, True)
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim id As Long
Dim OKClick As Boolean
   If Area = 1 Then
   If Not VerifyAccessRight("PACKAGE-CENTER_EX-WORKS-PRICE_EDIT", "แก้ไขข้อมูลราคาสินค้า") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   id = Val(GridEX1.Value(1))
   Call glbDatabaseMngr.LockTable(m_TableName, id, IsCanLock, glbErrorLog)
   
   frmAddEditExWorksPrice.id = id
   frmAddEditExWorksPrice.HeaderText = MapText("แก้ไขข้อมูลราคาสินค้า")
ElseIf Area = 2 Then
      If Not VerifyAccessRight("PACKAGE-CENTER_DELIVERY-COST_EDIT", "แก้ไขข้อมูลราคาค่าขนส่ง") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      If Not VerifyGrid(GridEX1.Value(1)) Then
         Exit Sub
      End If
      
      id = Val(GridEX1.Value(1))
      Call glbDatabaseMngr.LockTable(m_TableName, id, IsCanLock, glbErrorLog)
      
      frmAddEditExWorksPrice.id = id
      frmAddEditExWorksPrice.HeaderText = MapText("แก้ไขข้อมูลราคาค่าขนส่ง")
   ElseIf Area = 3 Then
      If Not VerifyAccessRight("PACKAGE-CENTER_PROMOTION-PART_EDIT", "แก้ไขข้อมูลราคาส่วนลดค่าสินค้า(หน้าบิล)") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      If Not VerifyGrid(GridEX1.Value(1)) Then
         Exit Sub
      End If
      
      id = Val(GridEX1.Value(1))
      Call glbDatabaseMngr.LockTable(m_TableName, id, IsCanLock, glbErrorLog)
      
      frmAddEditExWorksPrice.id = id
      frmAddEditExWorksPrice.HeaderText = MapText("แก้ไขข้อมูลราคาส่วนลดค่าสินค้า(หน้าบิล)")
   ElseIf Area = 4 Then
      If Not VerifyAccessRight("PACKAGE-CENTER_PROMOTION-DELIVERY_EDIT", "แก้ไขข้อมูลส่วนลดราคาค่าขนส่ง") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      If Not VerifyGrid(GridEX1.Value(1)) Then
         Exit Sub
      End If
      
      id = Val(GridEX1.Value(1))
      Call glbDatabaseMngr.LockTable(m_TableName, id, IsCanLock, glbErrorLog)
      
      frmAddEditExWorksPrice.id = id
      frmAddEditExWorksPrice.HeaderText = MapText("แก้ไขข้อมูลส่วนลดราคาค่าขนส่ง")
   ElseIf Area = 5 Then
      If Not VerifyAccessRight("PACKAGE-CENTER_PROMOTION-PART-EXTRA_EDIT", "แก้ไขข้อมูลราคาส่วนลดพิเศษค่าสินค้า(หลังบิล)") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      If Not VerifyGrid(GridEX1.Value(1)) Then
         Exit Sub
      End If
      
      id = Val(GridEX1.Value(1))
      Call glbDatabaseMngr.LockTable(m_TableName, id, IsCanLock, glbErrorLog)
      
      frmAddEditExWorksPrice.id = id
      frmAddEditExWorksPrice.HeaderText = MapText("แก้ไขข้อมูลราคาส่วนลดพิเศษค่าสินค้า(หลังบิล)")
   End If '
   frmAddEditExWorksPrice.Area = Area
   frmAddEditExWorksPrice.ShowMode = SHOW_EDIT
   Load frmAddEditExWorksPrice
   frmAddEditExWorksPrice.Show 1
   
   OKClick = frmAddEditExWorksPrice.OKClick
   
   Unload frmAddEditExWorksPrice
   Set frmAddEditExWorksPrice = Nothing
               
   If OKClick Then
      Call QueryData(True)
   End If
   Call glbDatabaseMngr.UnLockTable(m_TableName, id, IsCanLock, glbErrorLog)

End Sub

Private Sub cmdImport_Click()
Dim cPopup As cPopupMenu
Dim lMenuChosen As Long
 If Area = 1 Then
  If Not VerifyAccessRight("PACKAGE-CENTER_EX-WORKS-PRICE_IMPORT", "นำเข้าข้อมูลราคาสินค้า") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
  
 ElseIf Area = 2 Then
    If Not VerifyAccessRight("PACKAGE-CENTER_DELIVERY-COST_IMPORT", "นำเข้าข้อมูลราคาค่าขนส่ง") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

 ElseIf Area = 3 Then
    If Not VerifyAccessRight("PACKAGE-CENTER_PROMOTION-PART_IMPORT", "นำเข้าข้อมูลโปรโมชั่นราคาสินค้า") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

 ElseIf Area = 4 Then
    If Not VerifyAccessRight("PACKAGE-CENTER_PROMOTION-DELIVERY_IMPORT", "นำเข้าข้อมูลโปรโมชั่นราคาค่าขนส่ง") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

 End If
 
    frmImportWorkPrice.Area = Area
   Load frmImportWorkPrice
   frmImportWorkPrice.Show 1

   Unload frmImportWorkPrice
   Set frmImportWorkPrice = Nothing

   Call QueryData(True)
End Sub

Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub

Private Sub cmdSearch_Click()
   Call QueryData(True)
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      
      Call InitPackageOrderBy(cboOrderBy)
      Call InitOrderType(cboOrderType)
      
      Call QueryData(True)
   End If
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean

Dim Temp As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      m_ExWorksPrice.EX_WORKS_PRICE_CODE = PatchWildCard(txtPackageCode.Text)
      m_ExWorksPrice.EX_WORKS_PRICE_DESC = PatchWildCard(txtDesc.Text)
      m_ExWorksPrice.EX_WORKS_PRICE_LEVEL = Check2Flag(chkMarket.Value)
      m_ExWorksPrice.OrderBy = cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex))
      m_ExWorksPrice.OrderType = cboOrderType.ItemData(Minus2Zero(cboOrderType.ListIndex))
      
         m_ExWorksPrice.EX_WORKS_PRICE_TYPE = Area
         If Not glbDaily.QueryExWorksPrice(m_ExWorksPrice, m_Rs, ItemCount, IsOK, glbErrorLog) Then
            glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
            Call EnableForm(Me, True)
            Exit Sub
         End If
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If

   GridEX1.ItemCount = ItemCount
   GridEX1.Rebind

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
      Call cmdAdd_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
      Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 123 Then
      Call AddMemoNote
      KeyCode = 0
   End If
End Sub

Private Sub InitGrid1()
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
   Col.Width = 1500
   Col.Caption = MapText("แพคเกจ")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 6500
   Col.Caption = MapText("รายละเอียดแพคเกจ")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 1700
   Col.Caption = MapText("วันที่ประกาศ")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 1700
   Col.Caption = MapText("วันที่มีผล")
   
   Set Col = GridEX1.Columns.add '6
   Col.Width = 1700
   Col.Caption = MapText("วันที่สิ้นสุด")
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = 1000
   Col.Caption = MapText("สร้าง")
   
   Set Col = GridEX1.Columns.add '8
   Col.Width = 1000
   Col.Caption = MapText("แก้ไข")

   GridEX1.ItemCount = 0
End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   If Area = 1 Then
      Me.Caption = MapText("ข้อมูลแพคเกจราคาค่าสินค้าหน้าโรง")
      pnlHeader.Caption = MapText("ข้อมูลแพคเกจราคาค่าสินค้าหน้าโรง")
   ElseIf Area = 2 Then
      Me.Caption = MapText("ข้อมูลแพคเกจราคาค่าขนส่ง")
      pnlHeader.Caption = MapText("ข้อมูลแพคเกจราคาค่าขนส่ง")
   ElseIf Area = 3 Then
      Me.Caption = MapText("ข้อมูลแพคเกจส่วนลดราคาสินค้า")
      pnlHeader.Caption = MapText("ข้อมูลแพคเกจส่วนลดราคาสินค้า")
   ElseIf Area = 4 Then
      Me.Caption = MapText("ข้อมูลแพคเกจส่วนลดราคาค่าขนส่ง")
      pnlHeader.Caption = MapText("ข้อมูลแพคเกจส่วนลดราคาค่าขนส่ง")
   ElseIf Area = 5 Then
      Me.Caption = MapText("ข้อมูลแพคเกจส่วนลดพิเศษราคาสินค้า")
      pnlHeader.Caption = MapText("ข้อมูลแพคเกจส่วนลดพิเศษราคาสินค้า")
   End If
   Call InitGrid1
   Call InitNormalLabel(lblPackageCode, MapText("แพคเกจ"))
   Call InitNormalLabel(lblDesc, MapText("รายละเอียดแพคเกจ"))
   Call InitNormalLabel(lblOrderBy, MapText("เรียงตาม"))
   Call InitNormalLabel(lblOrderType, MapText("เรียงจาก"))
   
   Call txtDesc.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   
   Call InitCheckBox(chkMarket, "เปิดใช้งาน")
   chkMarket.Value = FlagToCheck("Y")
   
   Call InitCombo(cboOrderBy)
   Call InitCombo(cboOrderType)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdImport.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdClear.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdImport, MapText("IMPORT"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา (F5)"))
   Call InitMainButton(cmdClear, MapText("เคลียร์ (F4)"))
   
   cmdImport.Visible = False
   
End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   m_TableName = "EX_WORKS_PRICE"
   
   Set m_ExWorksPrice = New CExWorksPrice
   Set m_TempExWorksPrice = New CExWorksPrice
   Set m_Rs = New ADODB.Recordset
   Set m_Rs2 = New ADODB.Recordset

   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
Dim TempID1 As Long
Dim TempID2 As Long
Dim Ms As CExWorksPrice
Dim IsOK As Boolean
Dim OKClick As Boolean

   If GridEX1.ItemCount <= 0 Then
         Exit Sub
   End If
   
   TempID1 = GridEX1.Value(1)
   
   If Button = 2 Then
      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.Popup("COPY")
      If lMenuChosen = 0 Then
         Exit Sub
      End If
      Set oMenu = Nothing
   Else
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   If lMenuChosen = 1 Then
   
   If Area = 1 Then
      If Not VerifyAccessRight("PACKAGE-CENTER_EX-WORKS-PRICE_ADD", "เพิ่มข้อมูลราคาสินค้า") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      frmAddEditExWorksPrice.HeaderText = MapText("เพิ่มข้อมูลราคาสินค้า")
   ElseIf Area = 2 Then
      If Not VerifyAccessRight("PACKAGE-CENTER_DELIVERY-COST_ADD", "เพิ่มข้อมูลราคาค่าขนส่ง") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      frmAddEditExWorksPrice.HeaderText = MapText("เพิ่มข้อมูลราคาค่าขนส่ง")
   ElseIf Area = 3 Then
      If Not VerifyAccessRight("PACKAGE-CENTER_PROMOTION-PART_ADD", "เพิ่มข้อมูลส่วนลดราคาสินค้า") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      frmAddEditExWorksPrice.HeaderText = MapText("เพิ่มข้อมูลราคาโปรโมชั่นสินค้า")
   ElseIf Area = 4 Then
      If Not VerifyAccessRight("PACKAGE-CENTER_PROMOTION-DELIVERY_ADD", "เพิ่มข้อมูลราคาส่วนลดค่าขนส่ง") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      frmAddEditExWorksPrice.HeaderText = MapText("เพิ่มข้อมูลราคาโปรโมชั่นค่าขนส่ง")
   ElseIf Area = 5 Then
      If Not VerifyAccessRight("PACKAGE-CENTER_PROMOTION-PART-EXTRA_ADD", "เพิ่มข้อมูลส่วนลดพิเศษราคาสินค้า") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      frmAddEditExWorksPrice.HeaderText = MapText("เพิ่มข้อมูลราคาโปรโมชั่นสินค้า")
   End If
   
   
   
      Set Ms = New CExWorksPrice
      Ms.EX_WORKS_PRICE_ID = TempID1
      Call glbDaily.CopyExWorksPrice(Ms, IsOK, True, glbErrorLog)
      Call QueryData(True)
      Set Ms = Nothing
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
'   RowBuffer.RowStyle = RowBuffer.Value(5)
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
   Call m_TempExWorksPrice.PopulateFromRS(1, m_Rs)
   Values(1) = m_TempExWorksPrice.EX_WORKS_PRICE_ID
   Values(2) = m_TempExWorksPrice.EX_WORKS_PRICE_CODE
   Values(3) = m_TempExWorksPrice.EX_WORKS_PRICE_DESC
   Values(4) = DateToStringExtEx2(m_TempExWorksPrice.EX_WORKS_PRICE_DATE)
   Values(5) = DateToStringExtEx2(m_TempExWorksPrice.FROM_ACTIVE_DATE)
   Values(6) = DateToStringExtEx2(m_TempExWorksPrice.TO_VALID_DATE)
   
   Set Temp_LTK = GetObject("CLoginTracking", m_LoginTracking, Trim(m_TempExWorksPrice.CREATE_BY), False)
   If Not Temp_LTK Is Nothing Then
      Values(7) = Temp_LTK.USER_NAME
   Else
      Values(7) = ""
   End If
   
   Set Temp_LTK = GetObject("CLoginTracking", m_LoginTracking, Trim(m_TempExWorksPrice.MODIFY_BY), False)
   If Not Temp_LTK Is Nothing Then
      Values(8) = Temp_LTK.USER_NAME
   Else
      Values(8) = ""
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
   cmdAdd.Top = ScaleHeight - 580
   cmdEdit.Top = ScaleHeight - 580
   cmdDelete.Top = ScaleHeight - 580
   
   cmdImport.Top = ScaleHeight - 580
   cmdOK.Top = ScaleHeight - 580
   cmdExit.Top = ScaleHeight - 580
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
   cmdImport.Left = cmdOK.Left - cmdImport.Width - 50
End Sub
