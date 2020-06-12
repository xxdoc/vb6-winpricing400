VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditPartItem 
   BackColor       =   &H80000000&
   ClientHeight    =   9420
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   11835
   Icon            =   "frmAddEditPartItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9420
   ScaleWidth      =   11835
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   9495
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   16748
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboProductType 
         Height          =   315
         Left            =   9690
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   3840
         Width           =   1875
      End
      Begin prjFarmManagement.uctlTextBox txtNumberPLCID 
         Height          =   435
         Left            =   7800
         TabIndex        =   32
         Top             =   2880
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   767
      End
      Begin VB.ComboBox cboParcelType 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2880
         Width           =   2955
      End
      Begin VB.ComboBox cboUnit 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2400
         Width           =   2955
      End
      Begin VB.ComboBox cboPartType 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1920
         Width           =   4485
      End
      Begin prjFarmManagement.uctlTextBox txtName 
         Height          =   435
         Left            =   1860
         TabIndex        =   2
         Top             =   1470
         Width           =   4485
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtPartNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   1020
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   767
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   3315
         Left            =   150
         TabIndex        =   17
         Top             =   5160
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   5847
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
         Column(1)       =   "frmAddEditPartItem.frx":27A2
         Column(2)       =   "frmAddEditPartItem.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditPartItem.frx":290E
         FormatStyle(2)  =   "frmAddEditPartItem.frx":2A6A
         FormatStyle(3)  =   "frmAddEditPartItem.frx":2B1A
         FormatStyle(4)  =   "frmAddEditPartItem.frx":2BCE
         FormatStyle(5)  =   "frmAddEditPartItem.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditPartItem.frx":2D5E
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   18
         Top             =   4620
         Width           =   11595
         _ExtentX        =   20452
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
      Begin prjFarmManagement.uctlTextBox txtBarcode 
         Height          =   435
         Left            =   7800
         TabIndex        =   1
         Top             =   1020
         Width           =   2175
         _ExtentX        =   5212
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtBillDesc 
         Height          =   435
         Left            =   7800
         TabIndex        =   3
         Top             =   1470
         Width           =   3945
         _ExtentX        =   5212
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtWeightPerPack 
         Height          =   435
         Left            =   7800
         TabIndex        =   5
         Top             =   1920
         Width           =   1455
         _ExtentX        =   5212
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtMin 
         Height          =   435
         Left            =   7800
         TabIndex        =   26
         Top             =   2400
         Width           =   1455
         _ExtentX        =   5212
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtMax 
         Height          =   435
         Left            =   10320
         TabIndex        =   28
         Top             =   1920
         Width           =   1455
         _ExtentX        =   5212
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlLocationLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   33
         Top             =   3360
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtNumberLabID 
         Height          =   435
         Left            =   9720
         TabIndex        =   34
         Top             =   3360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPartMaster 
         Height          =   435
         Left            =   1860
         TabIndex        =   38
         Top             =   3840
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin VB.Label lblProductType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   7440
         TabIndex        =   40
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label lblPartMaster 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPartMaster"
         Height          =   375
         Left            =   240
         TabIndex        =   37
         Top             =   3960
         Width           =   1485
      End
      Begin VB.Label lblNumberLabID 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   255
         Left            =   7320
         TabIndex        =   36
         Top             =   3480
         Width           =   2295
      End
      Begin VB.Label lblLocation 
         Alignment       =   1  'Right Justify
         Caption         =   "lblLocation"
         Height          =   375
         Left            =   240
         TabIndex        =   35
         Top             =   3420
         Width           =   1485
      End
      Begin VB.Label lblNumberPLCID 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   255
         Left            =   5040
         TabIndex        =   31
         Top             =   3000
         Width           =   2655
      End
      Begin Threed.SSCheck chkCancelFlag 
         Height          =   375
         Left            =   9960
         TabIndex        =   30
         Top             =   2880
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblMax 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   345
         Left            =   9480
         TabIndex        =   29
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label lblMin 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   345
         Left            =   6960
         TabIndex        =   27
         Top             =   2490
         Width           =   735
      End
      Begin VB.Label lblParcelType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   25
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label lblWeightPerPack 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   345
         Left            =   6390
         TabIndex        =   24
         Top             =   2010
         Width           =   1335
      End
      Begin VB.Label lblBillDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   345
         Left            =   6390
         TabIndex        =   23
         Top             =   1560
         Width           =   1335
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   22
         Top             =   8550
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPartItem.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   21
         Top             =   8550
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPartItem.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   20
         Top             =   8550
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblBarcode 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   345
         Left            =   6390
         TabIndex        =   19
         Top             =   1110
         Width           =   1335
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   16
         Top             =   2490
         Width           =   1575
      End
      Begin Threed.SSCheck chkPigFlag 
         Height          =   345
         Left            =   4920
         TabIndex        =   9
         Top             =   2370
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   609
         _Version        =   131073
         Caption         =   "SSCheck1"
         TripleState     =   -1  'True
      End
      Begin VB.Label lblPartType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   15
         Top             =   1980
         Width           =   1575
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   14
         Top             =   1530
         Width           =   1575
      End
      Begin VB.Label lblPartNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   13
         Top             =   1110
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10155
         TabIndex        =   8
         Top             =   8550
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8505
         TabIndex        =   7
         Top             =   8550
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPartItem.frx":356A
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditPartItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_PartItem As CPartItem
Private m_Sp As CSystemParam

Public id As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Public PartGroupID As Long

Private m_Locations As Collection
Private m_CusType As Collection
Private m_PartMaster As Collection
Private Sub cboParcelType_Click()
   m_HasModify = True
End Sub

Private Sub cboPartType_Click()
   m_HasModify = True
End Sub

Private Sub cboProductType_Change()
      m_HasModify = True
End Sub

Private Sub cboProductType_Click()
      m_HasModify = True
End Sub

Private Sub cboUnit_Click()
   m_HasModify = True
End Sub
Private Sub chkPigFlag_Click(Value As Integer)
   m_HasModify = True
End Sub
Private Sub chkCancelFlag_Click(Value As Integer)
' If Not VerifyAccessRight("INVENTORY_PART_SELL-CANCEL-FLAG", "สามารถตั้งค่าการยกเลิกได้") Then
'         Call EnableForm(Me, True)
'         Exit Sub
' End If
   m_HasModify = True
End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean

   If Not cmdAdd.Enabled Then
      Exit Sub
   End If
   
   OKClick = False
   If TabStrip1.SelectedItem.Index = 3 Then
      Set frmAddEditPartItemSpec.TempCollection = m_PartItem.HumidRates
      frmAddEditPartItemSpec.ShowMode = SHOW_ADD
      frmAddEditPartItemSpec.HeaderText = MapText("เพิ่มเกณฑ์ความชื้น")
      Load frmAddEditPartItemSpec
      frmAddEditPartItemSpec.Show 1

      OKClick = frmAddEditPartItemSpec.OKClick

      Unload frmAddEditPartItemSpec
      Set frmAddEditPartItemSpec = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_PartItem.HumidRates)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
         Set frmAddEditPartItemPicture.ParentForm = Me
         Set frmAddEditPartItemPicture.TempCollection = m_PartItem.Pictures
         frmAddEditPartItemPicture.ShowMode = SHOW_ADD
         frmAddEditPartItemPicture.PictureType = HEAD_PART
         frmAddEditPartItemPicture.HeaderText = MapText("เพิ่ม ") & PictureTypeToText(HEAD_PART)
         Load frmAddEditPartItemPicture
         frmAddEditPartItemPicture.Show 1
   
         OKClick = frmAddEditPartItemPicture.OKClick
   
         Unload frmAddEditPartItemPicture
         Set frmAddEditPartItemPicture = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_PartItem.Pictures)
         GridEX1.Rebind
      End If
   End If
   
   If OKClick Then
      m_HasModify = True
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
   
   ID2 = GridEX1.Value(2)
   ID1 = GridEX1.Value(1)
   
   If TabStrip1.SelectedItem.Index = 3 Then
      If ID1 <= 0 Then
         m_PartItem.HumidRates.Remove (ID2)
      Else
         m_PartItem.HumidRates.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_PartItem.HumidRates)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
      If ID1 <= 0 Then
         m_PartItem.Pictures.Remove (ID2)
      Else
         m_PartItem.Pictures.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_PartItem.Pictures)
      GridEX1.Rebind
      m_HasModify = True
   End If
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim id As Long
Dim OKClick As Boolean
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   id = Val(GridEX1.Value(2))
   OKClick = False
   
   If TabStrip1.SelectedItem.Index = 3 Then
      frmAddEditPartItemSpec.id = id
      Set frmAddEditPartItemSpec.TempCollection = m_PartItem.HumidRates
      frmAddEditPartItemSpec.HeaderText = MapText("แก้ไขเกณฑ์ความชื้น")
      frmAddEditPartItemSpec.ShowMode = SHOW_EDIT
      Load frmAddEditPartItemSpec
      frmAddEditPartItemSpec.Show 1

      OKClick = frmAddEditPartItemSpec.OKClick

      Unload frmAddEditPartItemSpec
      Set frmAddEditPartItemSpec = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_PartItem.HumidRates)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
        frmAddEditPartItemPicture.id = id
         Set frmAddEditPartItemPicture.ParentForm = Me
         Set frmAddEditPartItemPicture.TempCollection = m_PartItem.Pictures
         frmAddEditPartItemPicture.ShowMode = SHOW_EDIT
         frmAddEditPartItemPicture.PictureType = HEAD_PART
         frmAddEditPartItemPicture.HeaderText = MapText("แก้ไข ") & PictureTypeToText(HEAD_PART)
         Load frmAddEditPartItemPicture
         frmAddEditPartItemPicture.Show 1

      OKClick = frmAddEditPartItemPicture.OKClick

      Unload frmAddEditPartItemPicture
      Set frmAddEditPartItemPicture = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_PartItem.Pictures)
         GridEX1.Rebind
      End If
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdOK_Click()
Dim oMenu As cPopupMenu
Dim lMenuChosen  As Long

   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("บันทึก", "-", "บันทึกและออกจากหน้าจอ")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
   If lMenuChosen = 1 Then
      If Not SaveData Then
         Exit Sub
      End If
      
      ShowMode = SHOW_EDIT
      id = m_PartItem.PART_ITEM_ID
      m_PartItem.QueryFlag = 1
      QueryData (True)
      m_HasModify = False
   ElseIf lMenuChosen = 3 Then
      If Not SaveData Then
         Exit Sub
      End If
      
      OKClick = True
      Unload Me
   End If

'   If Not SaveData Then
'      Exit Sub
'   End If
'
''   Call LoadAccessRight(Nothing, glbAccessRight, glbUser.GROUP_ID)
'   OKClick = True
'   Unload Me
End Sub

Private Sub InitGrid1()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.NAME = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   Set Col = GridEX1.Columns.add '3
   Col.Width = 3465
   Col.Caption = MapText("สถานที่จัดเก็บ")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 2745
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("จำนวนคงคลัง")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 2790
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ราคาเฉลี่ย")

   Set Col = GridEX1.Columns.add '6
   Col.Width = 2565
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ราคาหลังสุด")
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = 0
   Col.Visible = False
   Col.Caption = MapText("LocationID")
End Sub

Private Sub InitGrid2()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.NAME = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   Set Col = GridEX1.Columns.add '3
   Col.Width = 6030
   Col.Caption = MapText("ชื่อซัพพลายเออร์")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 2745
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ยอดนำเข้ารวม")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 2790
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("มูลค่านำเข้ารวม")
End Sub

Private Sub InitGrid3()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.NAME = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   Set Col = GridEX1.Columns.add '3
   Col.Width = 3270
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("จาก % ความชื้น")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 4545
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ถึง % ความชื้น")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 3375
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("หัก/ตัน")
End Sub
Private Sub InitGrid4()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.NAME = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 3270
   Col.TextAlignment = jgexAlignLeft
   Col.Caption = MapText("ประเภทที่จัดเก็บ")

   Set Col = GridEX1.Columns.add '3
   Col.Width = 9270
   Col.TextAlignment = jgexAlignLeft
   Col.Caption = MapText("ที่จัดเก็บรูป")

End Sub
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      m_PartItem.PART_ITEM_ID = id
      m_PartItem.QueryFlag = 1
      If Not glbDaily.QueryPartItem(m_PartItem, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_PartItem.PopulateFromRS(1, m_Rs)
      
      txtName.Text = m_PartItem.PART_DESC
      txtPartNo.Text = m_PartItem.PART_NO
      cboPartType.ListIndex = IDToListIndex(cboPartType, m_PartItem.PART_TYPE)
      cboUnit.ListIndex = IDToListIndex(cboUnit, m_PartItem.UNIT_COUNT)
      chkPigFlag.Value = FlagToCheck(m_PartItem.PIG_FLAG)
      chkCancelFlag.Value = FlagToCheck(m_PartItem.CANCEL_FLAG)
      txtBarcode.Text = m_PartItem.BARCODE_NO
      txtBillDesc.Text = m_PartItem.BILL_DESC
      txtWeightPerPack.Text = m_PartItem.WEIGHT_PER_PACK
      cboParcelType.ListIndex = IDToListIndex(cboParcelType, m_PartItem.PARCEL_TYPE)
      txtMin.Text = m_PartItem.MINIMUM_ALLOW
      txtMax.Text = m_PartItem.MAXIMUM_ALLOW
      txtNumberPLCID.Text = m_PartItem.NUMBER_PLC_ID
      uctlLocationLookup.MyCombo.ListIndex = IDToListIndex(uctlLocationLookup.MyCombo, m_PartItem.DEFAULT_LOCATION)
      uctlPartMaster.MyCombo.ListIndex = IDToListIndex(uctlPartMaster.MyCombo, m_PartItem.PART_MASTER_ID)
      cboProductType.ListIndex = IDToListIndex(cboProductType, m_PartItem.PRODUCT_TYPE_ID)
      txtNumberLabID.Text = m_PartItem.NUMBER_LAB_ID
      TabStrip1_Click
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Resize()
   SSFrame1.Width = ScaleWidth
   SSFrame1.HEIGHT = ScaleHeight
   
   pnlHeader.Width = ScaleWidth
   GridEX1.Width = ScaleWidth - 2 * GridEX1.Left
   GridEX1.HEIGHT = ScaleHeight - GridEX1.Top - 620
   
   TabStrip1.Width = GridEX1.Width
   
   cmdAdd.Top = ScaleHeight - 580
   cmdEdit.Top = cmdAdd.Top
   cmdDelete.Top = cmdAdd.Top
   cmdOK.Top = cmdAdd.Top
   cmdExit.Top = cmdAdd.Top
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_Locations = Nothing
   Set m_CusType = Nothing
   Set m_PartMaster = Nothing
End Sub

Private Sub GridEX1_DblClick()
Dim LocationID As Long

   If TabStrip1.SelectedItem.Index = 1 Then
      If Not VerifyGrid(GridEX1.Value(1)) Then
         Exit Sub
      End If

      LocationID = Val(GridEX1.Value(7))
      
      frmLotItem.PartItemID = m_PartItem.PART_ITEM_ID
      frmLotItem.LocationID = LocationID
      Load frmLotItem
      frmLotItem.Show 1
      
      Unload frmLotItem
      Set frmLotItem = Nothing
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
      Call cmdEdit_Click
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
      Call cmdEdit_Click
   End If
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Index = 1 Then
      If m_PartItem.PartLocations Is Nothing Then
         Exit Sub
      End If
   
      If RowIndex <= 0 Then
         Exit Sub
      End If
   
      Dim CR As CPartLocation
      If m_PartItem.PartLocations.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_PartItem.PartLocations, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If
   
      Values(1) = CR.PART_LOCATION_ID
      Values(2) = RealIndex
      Values(3) = CR.LOCATION_NAME
      Values(4) = FormatNumber(CR.CURRENT_AMOUNT)
      Values(5) = FormatNumber(CR.AVG_PRICE)
      Values(6) = FormatNumber(CR.LAST_PRICE)
      Values(7) = FormatNumber(CR.LOCATION_ID)
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      If m_PartItem.Suppliers Is Nothing Then
         Exit Sub
      End If
   
      If RowIndex <= 0 Then
         Exit Sub
      End If
   
      Dim Sp As CSupplier
      If m_PartItem.Suppliers.Count <= 0 Then
         Exit Sub
      End If
      Set Sp = GetItem(m_PartItem.Suppliers, RowIndex, RealIndex)
      If Sp Is Nothing Then
         Exit Sub
      End If
   
      Values(1) = Sp.SUPPLIER_ID
      Values(2) = RealIndex
      Values(3) = Sp.SUPPLIER_NAME
      Values(4) = FormatNumber(Sp.TX_AMOUNT)
      Values(5) = FormatNumber(Sp.TOTAL_INCLUDE_PRICE)
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
      If m_PartItem.HumidRates Is Nothing Then
         Exit Sub
      End If
   
      If RowIndex <= 0 Then
         Exit Sub
      End If
   
      Dim Ps As CPartItemSpec
      If m_PartItem.HumidRates.Count <= 0 Then
         Exit Sub
      End If
      Set Ps = GetItem(m_PartItem.HumidRates, RowIndex, RealIndex)
      If Ps Is Nothing Then
         Exit Sub
      End If
   
      Values(1) = Ps.PARTITEM_SPEC_ID
      Values(2) = RealIndex
      Values(3) = FormatNumber(Ps.FROM_RATE)
      Values(4) = FormatNumber(Ps.TO_RATE)
      Values(5) = FormatNumber(Ps.HUMIDITY_WEIGHT)
    ElseIf TabStrip1.SelectedItem.Index = 4 Then
      If m_PartItem.Pictures Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim Cs As CPartItemPicture
      If m_PartItem.Pictures.Count <= 0 Then
         Exit Sub
      End If
      Set Cs = GetItem(m_PartItem.Pictures, RowIndex, RealIndex)
      If Cs Is Nothing Then
         Exit Sub
      End If

      Values(1) = Cs.GetFieldValue("PART_ITEM_PICTURE_ID")
      Values(2) = RealIndex
      Values(3) = PictureTypeToText(Cs.GetFieldValue("PART_ITEM_PICTURE_TYPE"))
      Values(4) = Cs.GetFieldValue("PART_ITEM_PICTURE_PATH")
   End If
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
   If ShowMode = SHOW_EDIT Then
      If Not VerifyAccessRight("INVENTORY_PART_EDIT") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   End If
   If Not VerifyTextControl(lblPartNo, txtPartNo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblName, txtName, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblPartType, cboPartType, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblUnit, cboUnit, False) Then
      Exit Function
   End If
   If PartGroupID = 17 Or PartGroupID = 19 Then
      If Not VerifyCombo(lblParcelType, cboParcelType, False) Then
         Exit Function
      End If
      
      If Not VerifyCombo(lblPartMaster, uctlPartMaster.MyCombo, False) Then
         Exit Function
      End If
      
      If Not VerifyCombo(lblProductType, cboProductType, False) Then
         Exit Function
      End If
      
      'uctlPartMaster
      
   End If
   
     If PartGroupID = 19 Then
            If Not Val(txtWeightPerPack.Text) > 0 Then
               glbErrorLog.LocalErrorMsg = MapText("กรุณาระบุน้ำหนัก/ถุง ด้วย")
               glbErrorLog.ShowUserError
               Exit Function
            End If
      End If
   
   If chkCancelFlag.Value = 0 Then
      If PartGroupID = 17 Then
            If Not InStr(1, txtPartNo.Text, "-BK") > 0 Then
               glbErrorLog.LocalErrorMsg = MapText("การตั้งรหัสวัตถุดิบที่ใช้ผลิตอาหารต้องมีคำว่า -BK ต่อท้ายชื่อเสมอ")
               glbErrorLog.ShowUserError
               Exit Function
            End If
      End If
   End If
   
   If Not CheckUniqueNs(PARTNO_UNIQUE, txtPartNo.Text, id) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtPartNo.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_PartItem.PART_ITEM_ID = id
   m_PartItem.AddEditMode = ShowMode
   m_PartItem.PIG_FLAG = Check2Flag(chkPigFlag.Value)
   m_PartItem.CANCEL_FLAG = Check2Flag(chkCancelFlag.Value)
   m_PartItem.PART_NO = txtPartNo.Text
   m_PartItem.PART_DESC = txtName.Text
   m_PartItem.PART_TYPE = cboPartType.ItemData(Minus2Zero(cboPartType.ListIndex))
   m_PartItem.BARCODE_NO = txtBarcode.Text
   m_PartItem.BILL_DESC = txtBillDesc.Text
   m_PartItem.WEIGHT_PER_PACK = Val(txtWeightPerPack.Text)
   m_PartItem.MINIMUM_ALLOW = Val(txtMin.Text)
   m_PartItem.MAXIMUM_ALLOW = Val(txtMax.Text)
   m_PartItem.NUMBER_PLC_ID = txtNumberPLCID.Text
   m_PartItem.UNIT_COUNT = cboUnit.ItemData(Minus2Zero(cboUnit.ListIndex))
   m_PartItem.PARCEL_TYPE = cboParcelType.ItemData(Minus2Zero(cboParcelType.ListIndex))
   m_PartItem.DEFAULT_LOCATION = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
   m_PartItem.PART_MASTER_ID = uctlPartMaster.MyCombo.ItemData(Minus2Zero(uctlPartMaster.MyCombo.ListIndex))
   m_PartItem.PRODUCT_TYPE_ID = cboProductType.ItemData(Minus2Zero(cboProductType.ListIndex))
   m_PartItem.NUMBER_LAB_ID = txtNumberLabID.Text
   
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditPartItem(m_PartItem, IsOK, True, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call EnableForm(Me, True)
      Exit Function
   End If
   If Not IsOK Then
      Call EnableForm(Me, True)
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   Call EnableForm(Me, True)
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadUnit(cboUnit)
      Call InitParcelTypeEx(cboParcelType)
      
      Call LoadLocation(uctlLocationLookup.MyCombo, m_Locations, 2)
      Set uctlLocationLookup.MyCollection = m_Locations
     
      Call LoadPartMaster(uctlPartMaster.MyCombo, m_PartMaster)
     Set uctlPartMaster.MyCollection = m_PartMaster
     
     Call LoadMaster(cboProductType, , PRODUCT_TYPE)
      
      If ShowMode = SHOW_EDIT Then
         Call LoadPartType(cboPartType)
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         Call LoadPartType(cboPartType, , PartGroupID)
         id = 0
      End If
      
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

Private Sub InitFormLayout()
   Set m_Sp = GetSystemParam(glbSystemParams, "PROGRAM_OWNER")
   
   Call InitGrid1
   
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblName, MapText("ชื่อวัตถุดิบ"))
   Call InitNormalLabel(lblPartNo, MapText("รหัสวัตถุดิบ"))
   Call InitNormalLabel(lblPartType, MapText("ประเภทวัตถุดิบ"))
   
   Call InitNormalLabel(lblBarcode, MapText("รหัสขาย"))
   Call InitNormalLabel(lblBillDesc, MapText("ชื่อขาย"))
   Call InitNormalLabel(lblWeightPerPack, MapText("น้ำหนัก/ถุง"))
   Call InitNormalLabel(lblNumberPLCID, MapText("หมายเลขวัตถุดิบ PLC"))
   Call InitNormalLabel(lblNumberLabID, MapText("หมายเลขวัตถุดิบ LAB"))
   Call InitNormalLabel(lblProductType, MapText("รูปแบบวัตถุดิบ"))
   
   Call InitNormalLabel(lblUnit, MapText("หน่วยวัด"))
   Call InitNormalLabel(lblParcelType, MapText("ประเภทบรรจุ"))
   Call InitNormalLabel(lblLocation, MapText("คลังหลัก PLC"))
'   Call InitNormalLabel(lblCustomerType, MapText("ประเภทลูกค้า"))
   Call InitNormalLabel(lblPartMaster, MapText("ชื่อหลัก"))
  
   
   Call InitNormalLabel(lblMin, MapText("MIN"))
   Call InitNormalLabel(lblMax, MapText("MAX"))
   
   Call InitCheckBox(chkPigFlag, "รับเข้าจ่ายออก")
   Call InitCheckBox(chkCancelFlag, "ยกเลิก")
   
   Call txtName.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtPartNo.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtBarcode.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtBillDesc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtWeightPerPack.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   
   Call txtMin.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtMax.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtNumberPLCID.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitCombo(cboPartType)
   Call InitCombo(cboUnit)
   Call InitCombo(cboParcelType)
   Call InitCombo(cboProductType)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.NAME = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.add().Caption = MapText("สถานที่จัดเก็บ")
   TabStrip1.Tabs.add().Caption = MapText("ซัพพลายเออร์")
   TabStrip1.Tabs.add().Caption = MapText("เกณฑ์ความชื้น")
   TabStrip1.Tabs.add().Caption = MapText("ภาพประกอบ")
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   Set m_PartItem = New CPartItem
   Set m_Rs = New ADODB.Recordset

   Call EnableForm(Me, False)
   m_HasActivate = False
      
   Set m_Locations = New Collection
   Set m_CusType = New Collection
   Set m_PartMaster = New Collection
      
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub TabStrip1_Click()
   cmdAdd.Enabled = False
   cmdEdit.Enabled = False
   cmdDelete.Enabled = False
   
   If TabStrip1.SelectedItem.Index = 1 Then
      Call InitGrid1
      
      GridEX1.ItemCount = CountItem(m_PartItem.PartLocations)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      Call InitGrid2
      
      GridEX1.ItemCount = CountItem(m_PartItem.Suppliers)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
      Call InitGrid3
      
      GridEX1.ItemCount = CountItem(m_PartItem.HumidRates)
      GridEX1.Rebind
      
      cmdAdd.Enabled = True
      cmdEdit.Enabled = True
      cmdDelete.Enabled = True
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
      Call InitGrid4
      
      GridEX1.ItemCount = CountItem(m_PartItem.Pictures)
      GridEX1.Rebind
      
      cmdAdd.Enabled = True
'      cmdEdit.Enabled = True
      cmdDelete.Enabled = True
   
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub txtBarcode_Change()
   m_HasModify = True
End Sub

Private Sub txtBillDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtNumberLabID_Change()
   m_HasModify = True
End Sub

Private Sub txtNumberPLCID_Change()
   m_HasModify = True
End Sub
Private Sub txtMax_Change()
   m_HasModify = True
End Sub

Private Sub txtMin_Change()
   m_HasModify = True
End Sub

Private Sub txtPartNo_Change()
   m_HasModify = True
End Sub

Private Sub txtName_Change()
   m_HasModify = True
End Sub
Private Sub txtWeightPerPack_Change()
   m_HasModify = True
End Sub

Private Sub uctlCustomerType_Change()
m_HasModify = True
End Sub

Private Sub uctlLocationLookup_Change()
   m_HasModify = True
End Sub
Public Sub RefreshGrid()
   GridEX1.ItemCount = CountItem(m_PartItem.Pictures)
   GridEX1.Rebind
End Sub

Private Sub uctlPartMaster_Change()
   m_HasModify = True
End Sub
