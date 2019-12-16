VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditInventoryDocWHPart 
   ClientHeight    =   9795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16515
   Icon            =   "frmAddEditInventoryDocWHPart.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9795
   ScaleWidth      =   16515
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   9780
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   16515
      _ExtentX        =   29131
      _ExtentY        =   17251
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel pnlHeader 
         Height          =   585
         Left            =   10
         TabIndex        =   4
         Top             =   0
         Width           =   16485
         _ExtentX        =   29078
         _ExtentY        =   1032
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtPartNo 
         Height          =   435
         Left            =   2040
         TabIndex        =   0
         Top             =   720
         Width           =   3885
         _ExtentX        =   9604
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtDesc 
         Height          =   435
         Left            =   2040
         TabIndex        =   7
         Top             =   1200
         Width           =   3885
         _ExtentX        =   9604
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtWeightPerPack 
         Height          =   435
         Left            =   7080
         TabIndex        =   9
         Top             =   720
         Width           =   765
         _ExtentX        =   9604
         _ExtentY        =   767
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   6495
         Left            =   240
         TabIndex        =   11
         Top             =   2325
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   11456
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         TabKeyBehavior  =   1
         MethodHoldFields=   -1  'True
         AllowColumnDrag =   0   'False
         AllowEdit       =   0   'False
         BorderStyle     =   3
         GroupByBoxVisible=   0   'False
         ItemCount       =   0
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
         Column(1)       =   "frmAddEditInventoryDocWHPart.frx":27A2
         Column(2)       =   "frmAddEditInventoryDocWHPart.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditInventoryDocWHPart.frx":290E
         FormatStyle(2)  =   "frmAddEditInventoryDocWHPart.frx":2A6A
         FormatStyle(3)  =   "frmAddEditInventoryDocWHPart.frx":2B1A
         FormatStyle(4)  =   "frmAddEditInventoryDocWHPart.frx":2BCE
         FormatStyle(5)  =   "frmAddEditInventoryDocWHPart.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditInventoryDocWHPart.frx":2D5E
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   240
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1800
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   979
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "AngsanaUPC"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin GridEX20.GridEX GridEX2 
         Height          =   6510
         Left            =   11880
         TabIndex        =   13
         Top             =   2325
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   11483
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
         Column(1)       =   "frmAddEditInventoryDocWHPart.frx":2F36
         Column(2)       =   "frmAddEditInventoryDocWHPart.frx":2FFE
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditInventoryDocWHPart.frx":30A2
         FormatStyle(2)  =   "frmAddEditInventoryDocWHPart.frx":31FE
         FormatStyle(3)  =   "frmAddEditInventoryDocWHPart.frx":32AE
         FormatStyle(4)  =   "frmAddEditInventoryDocWHPart.frx":3362
         FormatStyle(5)  =   "frmAddEditInventoryDocWHPart.frx":343A
         ImageCount      =   0
         PrinterProperties=   "frmAddEditInventoryDocWHPart.frx":34F2
      End
      Begin MSComctlLib.TabStrip TabStrip2 
         Height          =   555
         Left            =   11880
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1800
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   979
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "AngsanaUPC"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin prjFarmManagement.uctlTextBox txtTotalAmount 
         Height          =   435
         Left            =   10560
         TabIndex        =   15
         Top             =   720
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTotalPallet 
         Height          =   435
         Left            =   10560
         TabIndex        =   18
         Top             =   1200
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   767
      End
      Begin VB.Label Label4 
         Caption         =   "Label3"
         Height          =   315
         Left            =   12600
         TabIndex        =   20
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblTotalPallet 
         Alignment       =   1  'Right Justify
         Caption         =   "lblTotalPallet"
         Height          =   315
         Left            =   9000
         TabIndex        =   19
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblTotalAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblWeightPerPack"
         Height          =   315
         Left            =   9000
         TabIndex        =   17
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Label3"
         Height          =   315
         Left            =   12600
         TabIndex        =   16
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   315
         Left            =   7920
         TabIndex        =   10
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblWeightPerPack 
         Alignment       =   1  'Right Justify
         Caption         =   "lblWeightPerPack"
         Height          =   315
         Left            =   6120
         TabIndex        =   8
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblPartDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPartDesc"
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label lblPartNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPartNo"
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   1695
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   12960
         TabIndex        =   1
         Top             =   9000
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditInventoryDocWHPart.frx":36CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   14640
         TabIndex        =   2
         Top             =   9000
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditInventoryDocWHPart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_LotItemWh As CLotItemWH
Private m_TempLotItemWh As CLotItemWH
Private m_InventoryWHDoc As CInventoryWHDoc
Private TempCollection As Collection

Private m_CollPallet As Collection
Private m_Lot As cLot
Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public id As Long
Public PART_ITEM_ID As Long
Public PART_NO As String
Public PART_DESC As String
Public WEIGHT_PER_PACK As Double
Public LOT_ITEM_WH_ID As Long
Public PART_TYPE As Long
Public DOCUMENT_TYPE As Long
Public LOCATION_ID As Long
Public Area As Long
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim m_PD As CPalletDoc
Dim m_TempPD As CPalletDoc
Dim I As Long
Dim TotalPart As Double

   IsOK = True
   If Flag Then
   Call EnableForm(Me, False)
   
   m_LotItemWh.PART_ITEM_ID = PART_ITEM_ID
   m_LotItemWh.LOCATION_ID = LOCATION_ID
   m_LotItemWh.OrderBy = 1
   m_LotItemWh.OrderType = 1
   m_LotItemWh.QueryFlag = 1
   m_LotItemWh.TX_TYPE = "I"
   m_LotItemWh.DOCUMENT_TYPE = DOCUMENT_TYPE

   If Not glbDaily.QueryLotItemWhPart(m_LotItemWh, m_Rs, ItemCount, IsOK, glbErrorLog, TotalPart) Then
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

   If DOCUMENT_TYPE = 13 Then
      txtTotalAmount.Text = FormatNumber(TotalPart, 3)
'      txtTotalAmount.Text = FormatNumber(GetTotalAmount2(m_LotItemWh.C_LotDoc, DOCUMENT_TYPE, m_LotItemWh.PART_ITEM_ID), 3)
   Else
'      txtTotalAmount.Text = FormatNumber(GetTotalAmount2(m_LotItemWh.C_LotDoc, DOCUMENT_TYPE, m_LotItemWh.PART_ITEM_ID), 0)
      txtTotalAmount.Text = FormatNumber(TotalPart, 0)
   End If

   GridEX1.ItemCount = CountItem(m_LotItemWh.C_LotDoc)
   GridEX1.Rebind
   Call EnableForm(Me, True)
End Sub
Private Function SaveData() As Boolean
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   SaveData = True
End Function
Private Function SaveData2() As Boolean
Dim IsOK As Boolean
Dim PD As CPalletDoc

   If Not m_HasModify Then
      SaveData2 = True
      Exit Function
   End If
   
   Call EnableForm(Me, False)
   Call glbDaily.StartTransaction
   
   For Each PD In TempCollection
      PD.AddEditData
   Next PD
'   If Not glbDaily.AddEditInventoryWhDoc(m_InventoryWHDoc, IsOK, False, glbErrorLog) Then
'      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'      SaveData = False
'      Call glbDaily.RollbackTransaction
'      Call EnableForm(Me, True)
'      Exit Function
'   End If
   
'   If Not IsOK Then
'      Call EnableForm(Me, True)
'      glbErrorLog.ShowUserError
'      Call glbDaily.RollbackTransaction
'      Exit Function
'   End If
   
   Call glbDaily.CommitTransaction
   Call EnableForm(Me, True)
   SaveData2 = True
End Function



Private Sub cmdOK_Click()
Dim oMenu As cPopupMenu
Dim lMenuChosen  As Long
'If Area = 2 Then
'   Set oMenu = New cPopupMenu
'   lMenuChosen = oMenu.Popup("บันทึก", "-", "บันทึกและออกจากหน้าจอ")
'   If lMenuChosen = 0 Then
'      Exit Sub
'   End If
'
'   If lMenuChosen = 1 Then
'      If Not SaveData2 Then
'         Exit Sub
'      End If
''      ShowMode = SHOW_EDIT
'
'      Call QueryData(True)
'      m_HasModify = False
'   ElseIf lMenuChosen = 3 Then
'      If Not SaveData2 Then
'         Exit Sub
'      End If
'
'      OKClick = True
'      Unload Me
'   End If
'Else
   If Not SaveData Then
      Exit Sub
   End If

   OKClick = True
   Unload Me
'End If
   
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      txtPartNo.Text = PART_NO
      txtDesc.Text = PART_DESC
      txtWeightPerPack.Text = WEIGHT_PER_PACK

      Call EnableForm(Me, False)
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         Call QueryData(True)
      End If
      
      Call EnableForm(Me, True)
      m_HasModify = False
   End If
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
   Col.Width = 1700
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("Lot ผลิต")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 1000
   Col.Caption = MapText("ถัง")
   
    If DOCUMENT_TYPE = 14 Or DOCUMENT_TYPE = 15 Then
      Set Col = GridEX1.Columns.add '5
      Col.Width = 1000
      Col.Caption = MapText("ล๊อค")
      
      Set Col = GridEX1.Columns.add '6
      Col.Width = 0
      Col.Caption = MapText("LOT_DOC_ID")
      
      Set Col = GridEX1.Columns.add '7
      Col.Width = 1400
      Col.Caption = MapText("วันที่ผลิต")
      
      Set Col = GridEX1.Columns.add '8
      Col.Width = 2200
      Col.Caption = MapText("วันที่/เวลา แพ็ค")
   ElseIf DOCUMENT_TYPE = 13 Or DOCUMENT_TYPE = 16 Then
      Set Col = GridEX1.Columns.add '5
      Col.Width = 0
      Col.Caption = MapText("ล๊อค")
      
      Set Col = GridEX1.Columns.add '6
      Col.Width = 0
      Col.Caption = MapText("LOT_DOC_ID")
      
      Set Col = GridEX1.Columns.add '7
      Col.Width = 1400
      Col.Caption = MapText("วันที่ผลิต")
      
      Set Col = GridEX1.Columns.add '8
      Col.Width = 2200
      Col.Caption = MapText("วันที่/เวลา แพ็ค")
   End If
   
   Set Col = GridEX1.Columns.add '9
   Col.Width = 1900
   Col.Caption = MapText("เลขที่เอกสาร")
   
   Set Col = GridEX1.Columns.add '10
   Col.Width = 1400
   Col.Caption = MapText("วันที่เอกสาร")
   
   Set Col = GridEX1.Columns.add '11
   Col.Width = 0
   Col.Caption = MapText("หัวจ่าย")
   
   Set Col = GridEX1.Columns.add '12
   Col.Width = 0
   Col.Caption = MapText("LOT_ITEM_WH_ID")
   
   Set Col = GridEX1.Columns.add '13
   Col.Width = 0
   Col.Caption = MapText("DOCUMENT_TYPE")
   
End Sub
Private Sub InitGrid2()
Dim Col As JSColumn
Dim I As Long

   GridEX2.Columns.Clear
   GridEX2.BackColor = GLB_GRID_COLOR
   GridEX2.ItemCount = 0
   GridEX2.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX2.ColumnHeaderFont.Bold = True
   GridEX2.ColumnHeaderFont.NAME = GLB_FONT
   GridEX2.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX2.Columns.add '1
   Col.Width = 10
   Col.Caption = "ID"

   Set Col = GridEX2.Columns.add '2
   Col.Width = 0
   Col.Caption = "RealIndex"
   
   If DOCUMENT_TYPE = 14 Or DOCUMENT_TYPE = 15 Then
      Set Col = GridEX2.Columns.add '3
      Col.Width = 1200
      Col.TextAlignment = jgexAlignCenter
      Col.Caption = MapText("ชื่อพาเลท")
      
      Set Col = GridEX2.Columns.add '4
      Col.Width = 1000
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("จำนวนถุง")
      
      Set Col = GridEX2.Columns.add '4
      Col.Width = 2000
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("จำนวนคงเหลือ(ถุง)")
   ElseIf DOCUMENT_TYPE = 13 Or DOCUMENT_TYPE = 16 Then
      Set Col = GridEX2.Columns.add '3
      Col.Width = 0
      Col.TextAlignment = jgexAlignCenter
      Col.Caption = MapText("อาหาร BULK")
      
      Set Col = GridEX2.Columns.add '4
      Col.Width = 2000
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("จำนวนทั้งหมด(กก.)")
      
      Set Col = GridEX2.Columns.add '4
      Col.Width = 2000
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("จำนวนคงเหลือ(กก.)")
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
   ElseIf Shift = 0 And KeyCode = 114 Then
'      Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 118 Then
'      Call SendData
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

Private Sub Form_Resize()
   TabStrip1.Width = GridEX1.Width
   TabStrip2.Width = GridEX2.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_LotItemWh = Nothing
   Set m_TempLotItemWh = Nothing
End Sub
Private Sub InitFormLayout()
Dim I As Long
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblPartNo, MapText("รหัสสินค้า"))
   Call InitNormalLabel(lblPartDesc, MapText("รายละเอียด"))
   Call InitNormalLabel(lblWeightPerPack, MapText("ถุงละ"))
   Call InitNormalLabel(Label3, MapText("กก."))
   Call InitNormalLabel(lblTotalAmount, MapText("คงเหลือทั้งหมด"))
'   Call InitNormalLabel(lblTotalBalance, MapText("คงเหลือยกมา"))
   Call InitNormalLabel(lblTotalPallet, MapText("คงเหลือทั้งล็อต"))
   Call InitNormalLabel(Label1, MapText("ถุง"))
   Call InitNormalLabel(Label4, MapText("ถุง"))
'   Call InitNormalLabel(Label5, MapText("ถุง"))
   
   txtTotalAmount.Enabled = False
   txtTotalPallet.Enabled = False
  
'   txtPartNo.Enabled = False
'   txtDesc.Enabled = False
'   txtWeightPerPack.Enabled = False
   
   Call InitGrid1
   TabStrip1.Font.Bold = True
   TabStrip1.Font.NAME = GLB_FONT
   TabStrip1.Font.Size = 16
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.add().Caption = MapText("รายการล๊อตการผลิต")
   
   Call InitGrid2
   TabStrip2.Font.Bold = True
   TabStrip2.Font.NAME = GLB_FONT
   TabStrip2.Font.Size = 16
   TabStrip2.Tabs.Clear
   TabStrip2.Tabs.add().Caption = MapText("รายการพาเลทที่วาง")

   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19

   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
  Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
'   If Area = 2 Then
'      cmdBalance.Visible = True
'      cmdBalance.Picture = LoadPicture(glbParameterObj.NormalButton1)
'      Call InitMainButton(cmdBalance, MapText("ปรับยอดคงเหลือ"))
'   End If
End Sub
Private Sub InitFormLayout2()
Dim I As Long
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblPartNo, MapText("รหัสสินค้า"))
   Call InitNormalLabel(lblPartDesc, MapText("รายละเอียด"))
   lblWeightPerPack.Visible = False
   txtWeightPerPack.Visible = False
   Label3.Visible = False

   Call InitNormalLabel(lblTotalAmount, MapText("คงเหลือทั้งหมด"))
   Call InitNormalLabel(lblTotalPallet, MapText("คงเหลือทั้งล็อต"))
   Call InitNormalLabel(Label1, MapText("กก."))
   Call InitNormalLabel(Label4, MapText("กก."))
   
   txtTotalAmount.Enabled = False
   txtTotalPallet.Enabled = False
  
'   txtPartNo.Enabled = False
'   txtDesc.Enabled = False
'   txtWeightPerPack.Enabled = False
   
   Call InitGrid1
   TabStrip1.Font.Bold = True
   TabStrip1.Font.NAME = GLB_FONT
   TabStrip1.Font.Size = 16
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.add().Caption = MapText("รายการล๊อตการผลิต")
   
   Call InitGrid2
   TabStrip2.Font.Bold = True
   TabStrip2.Font.NAME = GLB_FONT
   TabStrip2.Font.Size = 16
   TabStrip2.Tabs.Clear
   TabStrip2.Tabs.add().Caption = MapText("จำนวนอาหาร BULK")

   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19

   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
  Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
'   If Area = 2 Then
'      cmdBalance.Visible = True
'      cmdBalance.Picture = LoadPicture(glbParameterObj.NormalButton1)
'      Call InitMainButton(cmdBalance, MapText("ปรับยอดคงเหลือ"))
'   End If
End Sub
Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   OKClick = False
   If DOCUMENT_TYPE = 13 Or DOCUMENT_TYPE = 16 Then 'BULK
      Call InitFormLayout2
   ElseIf DOCUMENT_TYPE = 14 Or DOCUMENT_TYPE = 15 Then  'BAG
      Call InitFormLayout
   End If
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_LotItemWh = New CLotItemWH
   Set m_TempLotItemWh = New CLotItemWH
End Sub

Private Sub txtNote_Change()
   m_HasModify = True
End Sub
Private Sub txtPlanOut_Change()
   m_HasModify = True
End Sub
Private Sub uctlPlanDate_HasChange()
   m_HasModify = True
End Sub

Private Sub GridEX1_Click()
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   Call EnableForm(Me, False)
   Call TabStrip2_Click
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_DblClick()
'   Call cmdBalance_Click
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long
Dim LTD As CLotDoc

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"

     If m_LotItemWh Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If
      
   If CountItem(m_LotItemWh.C_LotDoc) <= 0 Then
      Exit Sub
   End If
   Set LTD = GetItem(m_LotItemWh.C_LotDoc, RowIndex, RealIndex)
   If LTD Is Nothing Then
      Exit Sub
   End If
   
      Values(1) = LTD.LOT_DOC_ID
      Values(2) = RealIndex
      Values(3) = LTD.LOT_NO
      Values(4) = LTD.BIN_NAME
      Values(5) = LTD.LOCK_NAME
      Values(6) = LTD.LOT_ID
      If LTD.DOCUMENT_TYPE = 15 Or LTD.DOCUMENT_TYPE = 16 Then
         Values(7) = DateToStringExtEx2(LTD.BL_START_DATE)
      Else
         Values(7) = DateToStringExtEx2(LTD.START_DATE)
      End If
      
      Values(8) = DateToStringExtEx2(LTD.PACK_DATE) & " " & Format(LTD.TIME_PACK_BEGIN, "HH:mm")
      Values(9) = LTD.DOCUMENT_NO
      Values(10) = DateToStringExtEx2(LTD.DOCUMENT_DATE)
      Values(11) = LTD.HEAD_PACK_NO
      Values(12) = LTD.LOT_ITEM_WH_ID
      Values(13) = LTD.DOCUMENT_TYPE

   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'On Error GoTo ErrorHandler
'Dim RealIndex As Long

End Sub

Private Sub GridEX2_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long
Dim I As Long

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip2.SelectedItem.Index = 1 Then
      If RowIndex <= 0 Then
         Exit Sub
      End If
      
      Dim PD As CPalletDoc
      Set PD = GetItem(m_LotItemWh.C_LotDoc.Item(id).C_PalletDoc, RowIndex, RealIndex)
      If PD Is Nothing Then
         Exit Sub
      End If
      
      Values(1) = PD.PALLET_DOC_ID
      Values(2) = RealIndex
      Values(3) = PD.PALLET_DOC_NO
      If DOCUMENT_TYPE = 13 Then
         Values(4) = FormatNumber(PD.CAPACITY_AMOUNT, 3)
      Else
         Values(4) = FormatNumber(PD.CAPACITY_AMOUNT, 0)
      End If
      
      
      Dim PD2 As CPalletDoc
      Set PD2 = GetObject("CPalletDoc", m_CollPallet, Trim(PD.PALLET_DOC_NO & "-" & PD.HEAD_PACK_NO), False)
      If Not (PD2 Is Nothing) Then
         If DOCUMENT_TYPE = 13 Then
            Values(5) = FormatNumber(PD2.PALLET_CAP_LAST, 3)
         Else
            Values(5) = FormatNumber(PD2.PALLET_CAP_LAST, 0)
         End If
      Else
          Values(5) = "0"
      End If
     
   End If
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub TabStrip2_Click()
Dim LotId As Long
Dim LotDocId As Long
Dim HeadPackNo  As Long
Dim LotItemWhId As Long
Dim DocumentType As Long
'Dim ID As Long
   If TabStrip2.SelectedItem.Index = 1 Then
      Call InitGrid2
      LotDocId = Val(GridEX1.Value(1))
      LotId = Val(GridEX1.Value(6))
      HeadPackNo = Val(GridEX1.Value(11))
      LotItemWhId = Val(GridEX1.Value(12))
      DocumentType = Val(GridEX1.Value(13))
      id = Val(GridEX1.Value(2))
      Set m_CollPallet = New Collection
      Call LoadPalletDocAmount(Nothing, m_CollPallet, LotId, 2, , 2, "I", , , LotDocId, HeadPackNo, LotItemWhId, DocumentType, PART_ITEM_ID, LotDocId, LOCATION_ID)
      If DOCUMENT_TYPE = 13 Then
         txtTotalPallet.Text = FormatNumber(GetTotalAmountPallet(m_CollPallet), 3)
      Else
         txtTotalPallet.Text = FormatNumber(GetTotalAmountPallet(m_CollPallet), 0)
      End If
      If CountItem(m_LotItemWh.C_LotDoc) > 0 Then
            GridEX2.Visible = True
            GridEX2.ItemCount = CountItem(m_LotItemWh.C_LotDoc.Item(id).C_PalletDoc)
            GridEX2.Rebind
            Set TempCollection = m_LotItemWh.C_LotDoc.Item(id).C_PalletDoc
         End If
     End If
End Sub
