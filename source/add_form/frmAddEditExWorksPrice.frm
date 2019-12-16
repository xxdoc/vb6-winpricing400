VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditExWorksPrice 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15480
   Icon            =   "frmAddEditExWorksPrice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9450
   ScaleWidth      =   15480
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   9495
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   15525
      _ExtentX        =   27384
      _ExtentY        =   16748
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlDate uctlFromActiveDate 
         Height          =   495
         Left            =   1860
         TabIndex        =   4
         Top             =   2520
         Width           =   4335
         _extentx        =   7646
         _extenty        =   873
      End
      Begin prjFarmManagement.uctlTextBox txtDesc 
         Height          =   435
         Left            =   1860
         TabIndex        =   2
         Top             =   1560
         Width           =   4485
         _extentx        =   13309
         _extenty        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   15495
         _ExtentX        =   27331
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtPackageNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   1080
         Width           =   2955
         _extentx        =   5212
         _extenty        =   767
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   4455
         Left            =   120
         TabIndex        =   7
         Top             =   4080
         Width           =   15195
         _ExtentX        =   26802
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
         Column(1)       =   "frmAddEditExWorksPrice.frx":27A2
         Column(2)       =   "frmAddEditExWorksPrice.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditExWorksPrice.frx":290E
         FormatStyle(2)  =   "frmAddEditExWorksPrice.frx":2A6A
         FormatStyle(3)  =   "frmAddEditExWorksPrice.frx":2B1A
         FormatStyle(4)  =   "frmAddEditExWorksPrice.frx":2BCE
         FormatStyle(5)  =   "frmAddEditExWorksPrice.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditExWorksPrice.frx":2D5E
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   16
         Top             =   3540
         Width           =   15195
         _ExtentX        =   26802
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
      Begin prjFarmManagement.uctlDate uctlToValidDate 
         Height          =   495
         Left            =   1860
         TabIndex        =   5
         Top             =   3000
         Width           =   4335
         _extentx        =   7646
         _extenty        =   873
      End
      Begin prjFarmManagement.uctlDate uctlDocumentDate 
         Height          =   495
         Left            =   1860
         TabIndex        =   3
         Top             =   2040
         Width           =   4335
         _extentx        =   7646
         _extenty        =   873
      End
      Begin prjFarmManagement.uctlTextBox txtPartNo 
         Height          =   435
         Left            =   9120
         TabIndex        =   13
         Top             =   1020
         Width           =   4485
         _extentx        =   13309
         _extenty        =   767
      End
      Begin VB.Label lblPartNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   7320
         TabIndex        =   24
         Top             =   1080
         Width           =   1695
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   13800
         TabIndex        =   15
         Top             =   1600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   13800
         TabIndex        =   14
         Top             =   1020
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditExWorksPrice.frx":2F36
         ButtonStyle     =   3
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   23
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label lblToValidDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   0
         TabIndex        =   22
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label lblFromActiveDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   120
         TabIndex        =   21
         Top             =   2640
         Width           =   1695
      End
      Begin Threed.SSCommand cmdAuto 
         Height          =   405
         Left            =   4920
         TabIndex        =   1
         Top             =   1080
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditExWorksPrice.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   10
         Top             =   8670
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditExWorksPrice.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   8
         Top             =   8670
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditExWorksPrice.frx":3884
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   9
         Top             =   8670
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCheck chkMarket 
         Height          =   345
         Left            =   5760
         TabIndex        =   6
         Top             =   1080
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   609
         _Version        =   131073
         Caption         =   "SSCheck1"
         TripleState     =   -1  'True
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   120
         TabIndex        =   20
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lblPackageNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   60
         TabIndex        =   19
         Top             =   1170
         Width           =   1725
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   13635
         TabIndex        =   12
         Top             =   8670
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   11985
         TabIndex        =   11
         Top             =   8670
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditExWorksPrice.frx":3B9E
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditExWorksPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_ExWorksPrice As CExWorksPrice
Private m_Sp As CSystemParam
Private m_ExWorkPricesItem As Collection
Private m_ExDeliveryCostItem As Collection
Private m_ExPromotionPartItem As Collection
Private m_ExPromotionDlcItem As Collection
Public Area As Long

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String

Private Sub cboPartType_Click()
   m_HasModify = True
End Sub

Private Sub cboUnit_Click()
   m_HasModify = True
End Sub

Private Sub chkMarket_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu

Set oMenu = New cPopupMenu
If Area = 1 Then
lMenuChosen = oMenu.Popup("สินค้า BAG", "-", "สินค้า BULK")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   OKClick = False
   If TabStrip1.SelectedItem.Index = 1 Then
    If lMenuChosen = 1 Then
      frmAddEditExWorksPriceItem.HeaderText = MapText("เพิ่มสินค้า BAG")
      frmAddEditExWorksPriceItem.PartType = 10
      frmAddEditExWorksPriceItem.ProductType = 1
    ElseIf lMenuChosen = 3 Then
      frmAddEditExWorksPriceItem.HeaderText = MapText("เพิ่มสินค้า BULK")
      frmAddEditExWorksPriceItem.PartType = 21
      frmAddEditExWorksPriceItem.ProductType = 2
    End If
      frmAddEditExWorksPriceItem.SocPartType = 3
      Set frmAddEditExWorksPriceItem.ParentForm = Me
      Set frmAddEditExWorksPriceItem.TempCollection = m_ExWorksPrice.ExWorksPriceItem
      Set frmAddEditExWorksPriceItem.m_ExWorkPricesItem = m_ExWorkPricesItem
      frmAddEditExWorksPriceItem.SocCode = txtPackageNo.Text
      frmAddEditExWorksPriceItem.ShowMode = SHOW_ADD
      Load frmAddEditExWorksPriceItem
      frmAddEditExWorksPriceItem.Show 1
   
      OKClick = frmAddEditExWorksPriceItem.OKClick
   
      Unload frmAddEditExWorksPriceItem
      Set frmAddEditExWorksPriceItem = Nothing
   
      If OKClick Then
         GridEX1.ItemCount = CountItem(m_ExWorksPrice.ExWorksPriceItem)
         GridEX1.Rebind
      End If
   End If
ElseIf Area = 2 Then
lMenuChosen = oMenu.Popup("ค่าขนส่งสินค้า BAG", "-", "ค่าขนส่งสินค้า BULK", "-", "ค่าขนส่งสินค้า เหมาเที่ยว")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   OKClick = False
   If TabStrip1.SelectedItem.Index = 1 Then
    If lMenuChosen = 1 Then
      frmAddEditExDeliveryCostItem.HeaderText = MapText("เพิ่มค่าขนส่งสินค้า BAG")
      frmAddEditExDeliveryCostItem.UnitType = 1
      frmAddEditExDeliveryCostItem.UnitTypeCus = 1
    ElseIf lMenuChosen = 3 Then
      frmAddEditExDeliveryCostItem.HeaderText = MapText("เพิ่มค่าขนส่งสินค้า BULK")
      frmAddEditExDeliveryCostItem.UnitType = 2
      frmAddEditExDeliveryCostItem.UnitTypeCus = 2
   ElseIf lMenuChosen = 5 Then
      frmAddEditExDeliveryCostItem.HeaderText = MapText("เพิ่มค่าขนส่งสินค้า เหมาเที่ยว")
      frmAddEditExDeliveryCostItem.UnitType = 3
      frmAddEditExDeliveryCostItem.UnitTypeCus = 3
    End If
      Set frmAddEditExDeliveryCostItem.ParentForm = Me
      Set frmAddEditExDeliveryCostItem.TempCollection = m_ExWorksPrice.ExDeliveryCost
      Set frmAddEditExDeliveryCostItem.m_ExDeliveryCostItem = m_ExDeliveryCostItem
      frmAddEditExDeliveryCostItem.PackageCode = txtPackageNo.Text
      frmAddEditExDeliveryCostItem.ShowMode = SHOW_ADD
      Load frmAddEditExDeliveryCostItem
      frmAddEditExDeliveryCostItem.Show 1
   
      OKClick = frmAddEditExDeliveryCostItem.OKClick
   
      Unload frmAddEditExDeliveryCostItem
      Set frmAddEditExDeliveryCostItem = Nothing
   
      If OKClick Then
         GridEX1.ItemCount = CountItem(m_ExWorksPrice.ExDeliveryCost)
         GridEX1.Rebind
      End If
   End If
ElseIf Area = 3 Then
  lMenuChosen = oMenu.Popup("ราคาโปรโมชั่นสินค้า BAG", "-", "ราคาโปรโมชั่นสินค้า BULK")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   OKClick = False
   If TabStrip1.SelectedItem.Index = 1 Then
    If lMenuChosen = 1 Then
      frmAddEditExPromotionPartItem.HeaderText = MapText("ราคาโปรโมชั่นสินค้า BAG")
      frmAddEditExPromotionPartItem.PartType = 10
      frmAddEditExPromotionPartItem.ProductType = 1
    ElseIf lMenuChosen = 3 Then
      frmAddEditExPromotionPartItem.HeaderText = MapText("ราคาโปรโมชั่นสินค้า BULK")
      frmAddEditExPromotionPartItem.PartType = 21
      frmAddEditExPromotionPartItem.ProductType = 2
    End If
      Set frmAddEditExPromotionPartItem.ParentForm = Me
      Set frmAddEditExPromotionPartItem.TempCollection = m_ExWorksPrice.ExPromotionPart
      Set frmAddEditExPromotionPartItem.m_ExPromotionPartItem = m_ExPromotionPartItem
      frmAddEditExPromotionPartItem.SocCode = txtPackageNo.Text
      frmAddEditExPromotionPartItem.ShowMode = SHOW_ADD
      Load frmAddEditExPromotionPartItem
      frmAddEditExPromotionPartItem.Show 1
   
      OKClick = frmAddEditExPromotionPartItem.OKClick
   
      Unload frmAddEditExPromotionPartItem
      Set frmAddEditExPromotionPartItem = Nothing
   
      If OKClick Then
         GridEX1.ItemCount = CountItem(m_ExWorksPrice.ExPromotionPart)
         GridEX1.Rebind
      End If
   End If
ElseIf Area = 4 Then
  lMenuChosen = oMenu.Popup("โปรโมชั่นขนส่งสินค้า BAG", "-", "โปรโมชั่นขนส่งสินค้า BULK")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   OKClick = False
   If TabStrip1.SelectedItem.Index = 1 Then
    If lMenuChosen = 1 Then
      frmAddEditExPromotionDlcItem.HeaderText = MapText("เพิ่มโปรโมชั่นขนส่งสินค้า BAG")
      frmAddEditExPromotionDlcItem.UnitType = 1
      frmAddEditExPromotionDlcItem.UnitTypeCus = 1
    ElseIf lMenuChosen = 3 Then
      frmAddEditExPromotionDlcItem.HeaderText = MapText("เพิ่มโปรโมชั่นขนส่งสินค้า BULK")
      frmAddEditExPromotionDlcItem.UnitType = 2
      frmAddEditExPromotionDlcItem.UnitTypeCus = 2
    End If
      Set frmAddEditExPromotionDlcItem.ParentForm = Me
      Set frmAddEditExPromotionDlcItem.TempCollection = m_ExWorksPrice.ExPromotionDlc
      Set frmAddEditExPromotionDlcItem.m_ExPromotionDlcItem = m_ExPromotionDlcItem
      frmAddEditExPromotionDlcItem.PackageCode = txtPackageNo.Text
      frmAddEditExPromotionDlcItem.ShowMode = SHOW_ADD
      Load frmAddEditExPromotionDlcItem
      frmAddEditExPromotionDlcItem.Show 1

      OKClick = frmAddEditExPromotionDlcItem.OKClick

      Unload frmAddEditExPromotionDlcItem
      Set frmAddEditExPromotionDlcItem = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_ExWorksPrice.ExPromotionDlc)
         GridEX1.Rebind
      End If
   End If
End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdAuto_Click()
Dim No As String
   If Trim(txtPackageNo.Text) = "" Then
         Call glbDatabaseMngr.GenerateNumber(EX_WORKS_PRICE, No, glbErrorLog)
         If Area = 1 Then
           No = "P" & No
         ElseIf Area = 2 Then
           No = "D" & No
         ElseIf Area = 3 Then
           No = "PP" & No
         ElseIf Area = 4 Then
           No = "PD" & No
         End If
         txtPackageNo.Text = No
   End If
End Sub
Private Sub cmdAuto_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
  KeyAscii = 0
End Sub

Private Sub cmdClear_Click()
   txtPartNo.Text = ""
End Sub

Private Sub cmdDelete_Click()
Dim ID1 As Long
Dim ID2 As Long

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
   If Not ConfirmDelete(GridEX1.Value(3)) Then
      Exit Sub
   End If
   
   ID2 = GridEX1.Value(2)
   ID1 = GridEX1.Value(1)
   
   If Area = 1 Then
      If TabStrip1.SelectedItem.Index = 1 Then
         If ID1 <= 0 Then
            m_ExWorksPrice.ExWorksPriceItem.Remove (ID2)
         Else
            m_ExWorksPrice.ExWorksPriceItem.Item(ID2).Flag = "D"
         End If

      GridEX1.ItemCount = CountItem(m_ExWorksPrice.ExWorksPriceItem)
      GridEX1.Rebind
      m_HasModify = True
      End If
      
   ElseIf Area = 2 Then
      If TabStrip1.SelectedItem.Index = 1 Then
         If ID1 <= 0 Then
            m_ExWorksPrice.ExDeliveryCost.Remove (ID2)
         Else
            m_ExWorksPrice.ExDeliveryCost.Item(ID2).Flag = "D"
         End If

      GridEX1.ItemCount = CountItem(m_ExWorksPrice.ExDeliveryCost)
      GridEX1.Rebind
      m_HasModify = True
      End If
   
   ElseIf Area = 3 Then
      If TabStrip1.SelectedItem.Index = 1 Then
         If ID1 <= 0 Then
            m_ExWorksPrice.ExPromotionPart.Remove (ID2)
         Else
            m_ExWorksPrice.ExPromotionPart.Item(ID2).Flag = "D"
         End If

      GridEX1.ItemCount = CountItem(m_ExWorksPrice.ExPromotionPart)
      GridEX1.Rebind
      m_HasModify = True
      End If
   End If

End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim ID2 As Long
Dim OKClick As Boolean
Dim lMenuChosen As Long

Dim RateType As Long
Dim RateType_Cus As Long

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(2))
   ID2 = Val(GridEX1.Value(1))
   lMenuChosen = Val(GridEX1.Value(6))
   OKClick = False
   
   If TabStrip1.SelectedItem.Index = 1 Then
     If Area = 1 Then
      Set frmAddEditExWorksPriceItem.ParentForm = Me
      frmAddEditExWorksPriceItem.SocPartType = 3
      frmAddEditExWorksPriceItem.ID = ID
      frmAddEditExWorksPriceItem.SocCode = txtPackageNo.Text
      Set frmAddEditExWorksPriceItem.TempCollection = m_ExWorksPrice.ExWorksPriceItem
      Set frmAddEditExWorksPriceItem.m_ExWorkPricesItem = m_ExWorkPricesItem
      frmAddEditExWorksPriceItem.ID_MUM = ID2
      frmAddEditExWorksPriceItem.HeaderText = MapText("แก้ไขสินค้า/บริการ")
      If lMenuChosen = 3 Then
      frmAddEditExWorksPriceItem.HeaderText = MapText("แก้ไขสินค้า/วัตถุดิบ")
      End If
      frmAddEditExWorksPriceItem.ShowMode = SHOW_EDIT
      Load frmAddEditExWorksPriceItem
      frmAddEditExWorksPriceItem.Show 1

      OKClick = frmAddEditExWorksPriceItem.OKClick

      Unload frmAddEditExWorksPriceItem
      Set frmAddEditExWorksPriceItem = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_ExWorksPrice.ExWorksPriceItem)
         GridEX1.Rebind
      End If
      ElseIf Area = 2 Then

         RateType = Val(GridEX1.Value(14))
         RateType_Cus = Val(GridEX1.Value(15))
         Set frmAddEditExDeliveryCostItem.ParentForm = Me
         frmAddEditExDeliveryCostItem.ID = ID
         frmAddEditExDeliveryCostItem.PackageCode = txtPackageNo.Text
         If RateType = 1 Then
            frmAddEditExDeliveryCostItem.HeaderText = MapText("แก้ไขค่าขนส่งสินค้า BAG")
            frmAddEditExDeliveryCostItem.UnitType = 1
            frmAddEditExDeliveryCostItem.UnitTypeCus = 1
          ElseIf RateType = 2 Then
            frmAddEditExDeliveryCostItem.HeaderText = MapText("แก้ไขค่าขนส่งสินค้า BULK")
            frmAddEditExDeliveryCostItem.UnitType = 2
            frmAddEditExDeliveryCostItem.UnitTypeCus = 2
         ElseIf RateType = 3 Then
            frmAddEditExDeliveryCostItem.HeaderText = MapText("แก้ไขค่าขนส่งสินค้า เหมาเที่ยว")
            frmAddEditExDeliveryCostItem.UnitType = 3
            frmAddEditExDeliveryCostItem.UnitTypeCus = 3
          End If
    
         Set frmAddEditExDeliveryCostItem.TempCollection = m_ExWorksPrice.ExDeliveryCost
         Set frmAddEditExDeliveryCostItem.m_ExDeliveryCostItem = m_ExDeliveryCostItem
         frmAddEditExDeliveryCostItem.ID_MUM = ID2
         frmAddEditExDeliveryCostItem.ShowMode = SHOW_EDIT
         Load frmAddEditExDeliveryCostItem
         frmAddEditExDeliveryCostItem.Show 1
   
         OKClick = frmAddEditExDeliveryCostItem.OKClick
   
         Unload frmAddEditExDeliveryCostItem
         Set frmAddEditExDeliveryCostItem = Nothing
   
         If OKClick Then
            GridEX1.ItemCount = CountItem(m_ExWorksPrice.ExDeliveryCost)
            GridEX1.Rebind
         End If
      ElseIf Area = 3 Then
         RateType = Val(GridEX1.Value(9))
'         RateType_Cus = Val(GridEX1.Value(15))
         Set frmAddEditExPromotionPartItem.ParentForm = Me
         frmAddEditExPromotionPartItem.ID = ID
         frmAddEditExPromotionPartItem.SocCode = txtPackageNo.Text
         If RateType = 1 Then
            frmAddEditExPromotionPartItem.HeaderText = MapText("แก้ไขราคาโปรโมชั่นสินค้า BAG")
            frmAddEditExWorksPriceItem.PartType = 10
            frmAddEditExWorksPriceItem.ProductType = 1
          ElseIf RateType = 2 Then
            frmAddEditExPromotionPartItem.HeaderText = MapText("แก้ไขราคาโปรโมชั่นสินค้า BULK")
            frmAddEditExWorksPriceItem.PartType = 21
            frmAddEditExWorksPriceItem.ProductType = 2
          End If
    
         Set frmAddEditExPromotionPartItem.TempCollection = m_ExWorksPrice.ExPromotionPart
         Set frmAddEditExPromotionPartItem.m_ExPromotionPartItem = m_ExPromotionPartItem
         frmAddEditExPromotionPartItem.ID_MUM = ID2
         frmAddEditExPromotionPartItem.ShowMode = SHOW_EDIT
         Load frmAddEditExPromotionPartItem
         frmAddEditExPromotionPartItem.Show 1
   
         OKClick = frmAddEditExPromotionPartItem.OKClick
   
         Unload frmAddEditExPromotionPartItem
         Set frmAddEditExPromotionPartItem = Nothing
   
         If OKClick Then
            GridEX1.ItemCount = CountItem(m_ExWorksPrice.ExPromotionPart)
            GridEX1.Rebind
         End If
   ElseIf Area = 4 Then
         RateType_Cus = Val(GridEX1.Value(10))
         Set frmAddEditExPromotionDlcItem.ParentForm = Me
         frmAddEditExPromotionDlcItem.ID = ID
         frmAddEditExPromotionDlcItem.PackageCode = txtPackageNo.Text
         If RateType = 1 Then
            frmAddEditExPromotionDlcItem.HeaderText = MapText("แก้ไขโปรโมชั่นค่าขนส่งสินค้า BAG")
            frmAddEditExPromotionDlcItem.UnitType = 1
            frmAddEditExPromotionDlcItem.UnitTypeCus = 1
          ElseIf RateType = 2 Then
            frmAddEditExPromotionDlcItem.HeaderText = MapText("แก้ไขโปรโมชั่นค่าขนส่งสินค้า BULK")
            frmAddEditExPromotionDlcItem.UnitType = 2
            frmAddEditExPromotionDlcItem.UnitTypeCus = 2
          End If
    
         Set frmAddEditExPromotionDlcItem.TempCollection = m_ExWorksPrice.ExPromotionDlc
         Set frmAddEditExPromotionDlcItem.m_ExPromotionDlcItem = m_ExPromotionDlcItem
         frmAddEditExPromotionDlcItem.ID_MUM = ID2
         frmAddEditExPromotionDlcItem.ShowMode = SHOW_EDIT
         Load frmAddEditExPromotionDlcItem
         frmAddEditExPromotionDlcItem.Show 1
   
         OKClick = frmAddEditExPromotionDlcItem.OKClick
   
         Unload frmAddEditExPromotionDlcItem
         Set frmAddEditExPromotionDlcItem = Nothing
   
         If OKClick Then
            GridEX1.ItemCount = CountItem(m_ExWorksPrice.ExPromotionDlc)
            GridEX1.Rebind
         End If
      End If
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
'   Call LoadAccessRight(Nothing, glbAccessRight, glbUser.GROUP_ID)
   OKClick = True
   Unload Me
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

If Area = 1 Then
   Set Col = GridEX1.Columns.add '3
   Col.Width = 4000
   Col.Caption = MapText("รหัสสินค้า/บริการ")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 3500
   Col.Caption = MapText("ชื่อสินค้า/บริการ")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 1800
   Col.Caption = MapText("ราคา/ถุง")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.add '6
   Col.Width = 1500
   Col.Caption = MapText("ประเภทสินค้า")
ElseIf Area = 2 Then
   Set Col = GridEX1.Columns.add '1
   Col.Width = 1200
   Col.Caption = MapText("รหัสลูกค้า")
   
   Set Col = GridEX1.Columns.add '2
   Col.Width = 3500
   Col.Caption = MapText("ชื่อลูกค้า")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 1300
   Col.Caption = MapText("รหัสสถานที")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 3500
   Col.Caption = MapText("สถานที่จัดส่ง")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 1600
   Col.Caption = MapText("ค่าขนส่ง/หน่วย")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.add '6
   Col.Width = 800
   Col.Caption = MapText("หน่วย")
   Col.TextAlignment = jgexAlignCenter
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = 1300
   Col.Caption = MapText("น้ำหนัก(กก.)")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.add '8
   Col.Width = 2000
   Col.Caption = MapText("ประเภทรถขนส่ง")
   
   Set Col = GridEX1.Columns.add '9
   Col.Width = 1600
   Col.Caption = MapText("คิดลูกค้า/หน่วย")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.add '10
   Col.Width = 800
   Col.Caption = MapText("หน่วย")
   Col.TextAlignment = jgexAlignCenter
   
   Set Col = GridEX1.Columns.add '11
   Col.Width = 1300
   Col.Caption = MapText("น้ำหนัก(กก.)")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.add '12
   Col.Width = 0
   Col.Caption = MapText("rate_type")
   
   Set Col = GridEX1.Columns.add '13
   Col.Width = 0
   Col.Caption = MapText("rate_type_cus")
ElseIf Area = 3 Then
   Set Col = GridEX1.Columns.add '1
   Col.Width = 1200
   Col.Caption = MapText("รหัสลูกค้า")
   
   Set Col = GridEX1.Columns.add '2
   Col.Width = 3500
   Col.Caption = MapText("ชื่อลูกค้า")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2500
   Col.Caption = MapText("รหัสสินค้า")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 3500
   Col.Caption = MapText("ชื่อสินค้า")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 1600
   Col.Caption = MapText("ส่วนลด/หน่วย")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.add '6
   Col.Width = 800
   Col.Caption = MapText("หน่วย")
   Col.TextAlignment = jgexAlignCenter
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = 0
   Col.Caption = MapText("rate_type_cus")
ElseIf Area = 4 Then
   Set Col = GridEX1.Columns.add '1
   Col.Width = 1200
   Col.Caption = MapText("รหัสลูกค้า")
   
   Set Col = GridEX1.Columns.add '2
   Col.Width = 3500
   Col.Caption = MapText("ชื่อลูกค้า")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 1300
   Col.Caption = MapText("รหัสสถานที")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 3500
   Col.Caption = MapText("สถานที่จัดส่ง")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 1600
   Col.Caption = MapText("ส่วนลด/หน่วย")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.add '6
   Col.Width = 800
   Col.Caption = MapText("หน่วย")
   Col.TextAlignment = jgexAlignCenter
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = 1300
   Col.Caption = MapText("น้ำหนัก(กก.)")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.add '8
   Col.Width = 0
   Col.Caption = MapText("rate_type")
   
   Set Col = GridEX1.Columns.add '9
   Col.Width = 0
   Col.Caption = MapText("rate_type_cus")
End If
End Sub


Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      m_ExWorksPrice.EX_WORKS_PRICE_ID = ID
      m_ExWorksPrice.QueryFlag = 1
      If Area = 1 Or Area = 3 Then
         m_ExWorksPrice.PART_NO_SEARCH = txtPartNo.Text
      ElseIf Area = 2 Or Area = 4 Then
         m_ExWorksPrice.CUSTOMER_CODE_SEARCH = txtPartNo.Text
      End If
      If Not glbDaily.QueryExWorksPrice(m_ExWorksPrice, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_ExWorksPrice.PopulateFromRS(1, m_Rs)

      txtPackageNo.Text = m_ExWorksPrice.EX_WORKS_PRICE_CODE
      txtDesc.Text = m_ExWorksPrice.EX_WORKS_PRICE_DESC
      chkMarket.Value = FlagToCheck(m_ExWorksPrice.EX_WORKS_PRICE_LEVEL)
      uctlDocumentDate.ShowDate = m_ExWorksPrice.EX_WORKS_PRICE_DATE
      uctlFromActiveDate.ShowDate = m_ExWorksPrice.FROM_ACTIVE_DATE
      uctlToValidDate.ShowDate = m_ExWorksPrice.TO_VALID_DATE
      
      TabStrip1_Click
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cmdSearch_Click()
     Call QueryData(True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set m_ExWorkPricesItem = Nothing
  Set m_ExDeliveryCostItem = Nothing
  Set m_ExPromotionPartItem = Nothing
  Set m_ExPromotionDlcItem = Nothing
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Index = 1 Then
   If Area = 1 Then
      If m_ExWorksPrice.ExWorksPriceItem Is Nothing Then
         Exit Sub
      End If
   
      If RowIndex <= 0 Then
         Exit Sub
      End If
   
      Dim EWPI As CExWorksPriceItem
      If m_ExWorksPrice.ExWorksPriceItem.Count <= 0 Then
         Exit Sub
      End If
      Set EWPI = GetItem(m_ExWorksPrice.ExWorksPriceItem, RowIndex, RealIndex)
      If EWPI Is Nothing Then
         Exit Sub
      End If
   
      Values(1) = EWPI.EX_WORKS_PRICE_ITEM_ID
      Values(2) = RealIndex
      Values(3) = EWPI.PART_NO
      Values(4) = EWPI.PART_DESC
      Values(5) = FormatNumber(EWPI.PACKAGE_RATE)
      If EWPI.PART_TYPE = 10 Then
         Values(6) = "BAG"
      ElseIf EWPI.PART_TYPE = 21 Then
         Values(6) = "BULK"
      Else
         Values(6) = ""
      End If
   ElseIf Area = 2 Then
      If m_ExWorksPrice.ExDeliveryCost Is Nothing Then
         Exit Sub
      End If
   
      If RowIndex <= 0 Then
         Exit Sub
      End If
   
      Dim EDCI As CExDeliveryCostItem
      If m_ExWorksPrice.ExDeliveryCost.Count <= 0 Then
         Exit Sub
      End If
      Set EDCI = GetItem(m_ExWorksPrice.ExDeliveryCost, RowIndex, RealIndex)
      If EDCI Is Nothing Then
         Exit Sub
      End If
   
      Values(1) = EDCI.EX_DELIVERY_COST_ITEM_ID
      Values(2) = RealIndex
      Values(3) = EDCI.CUSTOMER_CODE
      Values(4) = EDCI.CUSTOMER_NAME
      Values(5) = EDCI.DELIVERY_CUS_ITEM_CODE
      Values(6) = EDCI.DELIVERY_CUS_ITEM_NAME
      
      Values(7) = FormatNumber(EDCI.RATE_DELIVERY, 3)
      Values(8) = DeliveryUnit(EDCI.RATE_TYPE)
      Values(9) = FormatNumber(EDCI.WEIGHT_PER_PACK, 0)
      Values(10) = DeliveryType(EDCI.RATE_TYPE)
      
      Values(11) = FormatNumber(EDCI.RATE_CUSTOMER, 3)
      Values(12) = DeliveryUnit(EDCI.RATE_TYPE_CUS)
      Values(13) = FormatNumber(EDCI.WEIGHT_PER_PACK_CUS, 0)
      
      
      Values(14) = EDCI.RATE_TYPE
      Values(15) = EDCI.RATE_TYPE_CUS
   ElseIf Area = 3 Then
      If m_ExWorksPrice.ExPromotionPart Is Nothing Then
         Exit Sub
      End If
   
      If RowIndex <= 0 Then
         Exit Sub
      End If
   
      Dim EPPI    As CExPromotionPartItem
      If m_ExWorksPrice.ExPromotionPart.Count <= 0 Then
         Exit Sub
      End If
      Set EPPI = GetItem(m_ExWorksPrice.ExPromotionPart, RowIndex, RealIndex)
      If EPPI Is Nothing Then
         Exit Sub
      End If
   
      Values(1) = EPPI.EX_PROMOTION_PART_ITEM_ID
      Values(2) = RealIndex
      Values(3) = EPPI.CUSTOMER_CODE
      Values(4) = EPPI.CUSTOMER_NAME
      Values(5) = EPPI.PART_NO
      Values(6) = EPPI.PART_DESC
      Values(7) = FormatNumber(EPPI.DISCOUNT_AMOUNT)
      If EPPI.PART_TYPE = 10 Then
         Values(8) = "BAG"
      ElseIf EPPI.PART_TYPE = 21 Then
         Values(8) = "BULK"
      Else
         Values(8) = ""
      End If
      Values(9) = EPPI.RATE_TYPE
   ElseIf Area = 4 Then
      If m_ExWorksPrice.ExPromotionDlc Is Nothing Then
         Exit Sub
      End If
   
      If RowIndex <= 0 Then
         Exit Sub
      End If
   
      Dim EPDI As CExPromotionDlcItem
      If m_ExWorksPrice.ExPromotionDlc.Count <= 0 Then
         Exit Sub
      End If
      Set EPDI = GetItem(m_ExWorksPrice.ExPromotionDlc, RowIndex, RealIndex)
      If EPDI Is Nothing Then
         Exit Sub
      End If
   
      Values(1) = EPDI.EX_PROMOTION_DLC_ITEM_ID
      Values(2) = RealIndex
      Values(3) = EPDI.CUSTOMER_CODE
      Values(4) = EPDI.CUSTOMER_NAME
      Values(5) = EPDI.DELIVERY_CUS_ITEM_CODE
      Values(6) = EPDI.DELIVERY_CUS_ITEM_NAME
      
      Values(7) = FormatNumber(EPDI.DISCOUNT_AMOUNT, 3)
      Values(8) = DeliveryUnit(EPDI.RATE_TYPE_CUS)
      Values(9) = FormatNumber(EPDI.WEIGHT_PER_PACK_CUS, 0)
      Values(10) = EPDI.RATE_TYPE_CUS
   End If
   End If
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean

   If Not VerifyTextControl(lblPackageNo, txtPackageNo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblDesc, txtDesc, False) Then
      Exit Function
   End If
   
   If uctlFromActiveDate.ShowDate > uctlToValidDate.ShowDate Then
       glbErrorLog.LocalErrorMsg = MapText(lblFromActiveDate.Caption) & " ต้องไม่น้อยกว่า " & MapText(lblToValidDate.Caption)
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      cmdExit.Enabled = True
      cmdOK.Enabled = True
      Exit Function
   End If
   
   If Not CheckUniqueNs(WORKS_PRICE_ACTIVE_DATE_UNIQUE, Trim(DateToStringInt(uctlFromActiveDate.ShowDate)), ID, , 4, , Area) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล ช่วงวันทีมีผล ") & " " & uctlFromActiveDate.ShowDate & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      cmdExit.Enabled = True
      cmdOK.Enabled = True
      Exit Function
   End If
   
   If Not CheckUniqueNs(WORKS_PRICE_ACTIVE_DATE_UNIQUE, Trim(DateToStringInt(uctlToValidDate.ShowDate)), ID, , 4, , Area) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล ช่วงวันทีสิ้นสุด ") & " " & uctlToValidDate.ShowDate & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      cmdExit.Enabled = True
      cmdOK.Enabled = True
      Exit Function
   End If
   
'   If Not CheckUniqueNs(SOCNO_UNIQUE, txtDesc.Text, ID) Then
'      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtDesc.Text & " " & MapText("อยู่ในระบบแล้ว")
'      glbErrorLog.ShowUserError
'      Exit Function
'   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_ExWorksPrice.EX_WORKS_PRICE_ID = ID
   m_ExWorksPrice.AddEditMode = ShowMode
   m_ExWorksPrice.EX_WORKS_PRICE_LEVEL = Check2Flag(chkMarket.Value)
   m_ExWorksPrice.EX_WORKS_PRICE_CODE = txtPackageNo.Text
   m_ExWorksPrice.EX_WORKS_PRICE_DESC = txtDesc.Text
   m_ExWorksPrice.EX_WORKS_PRICE_STATUS = 0
   m_ExWorksPrice.EX_WORKS_PRICE_DATE = uctlDocumentDate.ShowDate
   m_ExWorksPrice.FROM_ACTIVE_DATE = uctlFromActiveDate.ShowDate
   m_ExWorksPrice.TO_VALID_DATE = uctlToValidDate.ShowDate
   If Area = 1 Then
      m_ExWorksPrice.EX_WORKS_PRICE_TYPE = 1 'ค่าสินค้า
   ElseIf Area = 2 Then
      m_ExWorksPrice.EX_WORKS_PRICE_TYPE = 2 'ค่าขนส่ง
   ElseIf Area = 3 Then
      m_ExWorksPrice.EX_WORKS_PRICE_TYPE = 3 'โปรโมชั่น สินค้า
   ElseIf Area = 4 Then
      m_ExWorksPrice.EX_WORKS_PRICE_TYPE = 4 'โปรโมชั่น ขนส่ง
   End If
   
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditExWorksPrice(m_ExWorksPrice, IsOK, True, glbErrorLog) Then
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
      
      uctlDocumentDate.ShowDate = Now
      uctlFromActiveDate.ShowDate = Now
      uctlToValidDate.ShowDate = Now
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
      End If
      
      If Area = 1 Then
         Call LoadExWorksPriceItem(Nothing, m_ExWorkPricesItem, m_ExWorksPrice.EX_WORKS_PRICE_ID, 2, -1, -1)
      ElseIf Area = 2 Then
         Call LoadExDeliveryCusItem(Nothing, m_ExDeliveryCostItem, m_ExWorksPrice.EX_WORKS_PRICE_ID, 4, -1, -1)
      ElseIf Area = 3 Then
         Call LoadExPromotionPartItem(Nothing, m_ExPromotionPartItem, m_ExWorksPrice.EX_WORKS_PRICE_ID, 2, -1, -1)
      ElseIf Area = 4 Then
         Call LoadExPromotionDlcItem(Nothing, m_ExPromotionDlcItem, m_ExWorksPrice.EX_WORKS_PRICE_ID, 3, -1, -1)
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

Private Sub InitFormLayout()
   Set m_Sp = GetSystemParam(glbSystemParams, "PROGRAM_OWNER")
   
   Call InitGrid1
   
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblPackageNo, MapText("แพคเกจ"))
   Call InitNormalLabel(lblDesc, MapText("ข้อมูลแพคเกจ"))
   Call InitNormalLabel(lblDocumentDate, MapText("วันที่ประกาศ"))
   Call InitNormalLabel(lblFromActiveDate, MapText("วันที่มีผล"))
   Call InitNormalLabel(lblToValidDate, MapText("วันที่สิ้นสุด"))
   If Area = 1 Or Area = 3 Then
      Call InitNormalLabel(lblPartNo, MapText("รหัสสินค้า"))
   ElseIf Area = 2 Or Area = 4 Then
      Call InitNormalLabel(lblPartNo, MapText("รหัสลูกค้า"))
   End If
   
   'lblPartNo
   
   Call InitCheckBox(chkMarket, "เปิดใช้งาน")
   If ShowMode = SHOW_ADD Then
      chkMarket.Value = ssCBChecked
   End If
   
   Call txtDesc.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   Call txtPackageNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtPartNo.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   
   If Area = 1 Or Area = 3 Then
      Call txtPartNo.SetKeySearch("PART_NO")
   ElseIf Area = 2 Or Area = 4 Then
      Call txtPartNo.SetKeySearch("CUSTOMER_CODE")
   End If
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdAuto.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdClear.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdAuto, MapText("A"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา (F5)"))
   Call InitMainButton(cmdClear, MapText("เคลียร์ (F4)"))
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.NAME = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.add().Caption = MapText("สินค้า/บริการ")
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   Set m_ExWorksPrice = New CExWorksPrice
   Set m_ExWorkPricesItem = New Collection
   Set m_ExDeliveryCostItem = New Collection
   Set m_ExPromotionPartItem = New Collection
  Set m_ExPromotionDlcItem = New Collection
  
   Set m_Rs = New ADODB.Recordset

   Call EnableForm(Me, False)
   m_HasActivate = False
      
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub txtLastName_Change()
   m_HasModify = True
End Sub

Private Sub TabStrip1_Click()
   If TabStrip1.SelectedItem.Index = 1 Then
      Call InitGrid1
      If Area = 1 Then
         GridEX1.ItemCount = CountItem(m_ExWorksPrice.ExWorksPriceItem)
         GridEX1.Rebind
      ElseIf Area = 2 Then
         GridEX1.ItemCount = CountItem(m_ExWorksPrice.ExDeliveryCost)
         GridEX1.Rebind
      ElseIf Area = 3 Then
         GridEX1.ItemCount = CountItem(m_ExWorksPrice.ExPromotionPart)
         GridEX1.Rebind
      ElseIf Area = 4 Then
         GridEX1.ItemCount = CountItem(m_ExWorksPrice.ExPromotionDlc)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub txtPackageNo_Change()
   m_HasModify = True
End Sub

Private Sub txtDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtMaxUser_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox1_Change()
   m_HasModify = True
End Sub

Private Sub txtUnitWeight_Change()
   m_HasModify = True
End Sub

Private Sub uctlDocumentDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlFromActiveDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlToValidDate_HasChange()
   m_HasModify = True
End Sub
Public Sub ShowGridItem()
   If TabStrip1.SelectedItem.Index = 1 Then
     If Area = 1 Then
      GridEX1.ItemCount = CountItem(m_ExWorksPrice.ExWorksPriceItem)
      GridEX1.Rebind
   ElseIf Area = 2 Then
      GridEX1.ItemCount = CountItem(m_ExWorksPrice.ExDeliveryCost)
      GridEX1.Rebind
   End If
   End If
   m_HasModify = True
End Sub

