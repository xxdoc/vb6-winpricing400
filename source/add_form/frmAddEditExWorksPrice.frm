VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditExWorksPrice 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18945
   Icon            =   "frmAddEditExWorksPrice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9450
   ScaleWidth      =   18945
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   9495
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   19005
      _ExtentX        =   33523
      _ExtentY        =   16748
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboDeclareCount 
         Height          =   315
         Left            =   17160
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   3240
         Width           =   1065
      End
      Begin VB.TextBox txtNote 
         Height          =   1935
         Left            =   7440
         MultiLine       =   -1  'True
         TabIndex        =   28
         Top             =   1560
         Width           =   9495
      End
      Begin prjFarmManagement.uctlDate uctlFromActiveDate 
         Height          =   495
         Left            =   1860
         TabIndex        =   3
         Top             =   2520
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   873
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   10
         TabIndex        =   17
         Top             =   0
         Width           =   18975
         _ExtentX        =   33470
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
         _ExtentX        =   5212
         _ExtentY        =   767
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   4455
         Left            =   120
         TabIndex        =   6
         Top             =   4080
         Width           =   18795
         _ExtentX        =   33152
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
         TabIndex        =   15
         Top             =   3540
         Width           =   18765
         _ExtentX        =   33099
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
         TabIndex        =   4
         Top             =   3000
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   873
      End
      Begin prjFarmManagement.uctlDate uctlDocumentDate 
         Height          =   495
         Left            =   1860
         TabIndex        =   2
         Top             =   2040
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   873
      End
      Begin prjFarmManagement.uctlTextBox txtPartNo 
         Height          =   435
         Left            =   12480
         TabIndex        =   12
         Top             =   1020
         Width           =   4485
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtDesc 
         Height          =   435
         Left            =   1860
         TabIndex        =   30
         Top             =   1560
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   767
      End
      Begin VB.Label lblDeclareCount 
         Caption         =   "Label1"
         Height          =   435
         Left            =   17160
         TabIndex        =   32
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   60
         TabIndex        =   29
         Top             =   1680
         Width           =   1725
      End
      Begin Threed.SSCommand cmdVerify 
         Height          =   525
         Left            =   8400
         TabIndex        =   27
         Top             =   8670
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditExWorksPrice.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdApproved 
         Height          =   525
         Left            =   10080
         TabIndex        =   26
         Top             =   8670
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditExWorksPrice.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   525
         Left            =   13920
         TabIndex        =   25
         Top             =   8670
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditExWorksPrice.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdApprove 
         Height          =   525
         Left            =   12240
         TabIndex        =   24
         Top             =   8670
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditExWorksPrice.frx":3884
         ButtonStyle     =   3
      End
      Begin VB.Label lblPartNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   10680
         TabIndex        =   23
         Top             =   1080
         Width           =   1695
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   17160
         TabIndex        =   14
         Top             =   1680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   17160
         TabIndex        =   13
         Top             =   1020
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditExWorksPrice.frx":3B9E
         ButtonStyle     =   3
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   22
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label lblToValidDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   0
         TabIndex        =   21
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label lblFromActiveDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   120
         TabIndex        =   20
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
         MouseIcon       =   "frmAddEditExWorksPrice.frx":3EB8
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   9
         Top             =   8670
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditExWorksPrice.frx":41D2
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   7
         Top             =   8670
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditExWorksPrice.frx":44EC
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   8
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
         TabIndex        =   5
         Top             =   1080
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   609
         _Version        =   131073
         Caption         =   "chkMarket"
         TripleState     =   -1  'True
      End
      Begin VB.Label lblNote 
         Caption         =   "Label1"
         Height          =   435
         Left            =   7440
         TabIndex        =   19
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblPackageNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   60
         TabIndex        =   18
         Top             =   1170
         Width           =   1725
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   17235
         TabIndex        =   11
         Top             =   8670
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   15600
         TabIndex        =   10
         Top             =   8670
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditExWorksPrice.frx":4806
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
Public canShowGP As Boolean

Public id As Long
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
      frmAddEditExWorksPriceItem.canShowGP = canShowGP
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

Private Sub cmdApprove_Click()
Dim oMenu As cPopupMenu
Dim TempUserName As String
Dim lMenuChosen As Long
Dim TempStr As String
''If Area = 2 Or Area = 4 Then
''   Exit Sub
''End If

   If m_HasModify Or ((Not m_HasModify) And (ShowMode = SHOW_ADD)) Then
      glbErrorLog.LocalErrorMsg = "กรุณาทำการบันทึกข้อมูลให้เรียบร้อยก่อน"
      glbErrorLog.ShowUserError
      Exit Sub
   End If

   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("ตรวจสอบ ราคา", "-", "อนุมัติ ราคา ")
   If lMenuChosen = 0 Then
      Exit Sub
   End If

'ตรวจสอบสิทธิ์
If lMenuChosen = 1 Then 'ตรวจสอบสิทธิ์
   If Area = 1 Then
      frmVerifyAccRight.AccName = "PACKAGE-CENTER_EX-WORKS-PRICE_VERIFY"
      frmVerifyAccRight.AccDesc = "สามารถตรวจสอบราคาประกาศสินค้าหน้าโรงงานได้"
   ElseIf Area = 2 Then
      frmVerifyAccRight.AccName = "PACKAGE-CENTER_DELIVERY-COST_VERIFY"
      frmVerifyAccRight.AccDesc = "สามารถตรวจสอบราคาค่าขนส่งได้"
   ElseIf Area = 3 Then
      frmVerifyAccRight.AccName = "PACKAGE-CENTER_PROMOTION-PART_VERIFY"
      frmVerifyAccRight.AccDesc = "สามารถตรวจสอบราคาประกาศสินค้าโปรโมชั่นได้"
   ElseIf Area = 4 Then
      frmVerifyAccRight.AccName = "PACKAGE-CENTER_PROMOTION-DELIVERY_VERIFY"
      frmVerifyAccRight.AccDesc = "สามารถตรวจสอบราคาโปรโมชั่นค่าขนส่งได้"
   End If
ElseIf lMenuChosen = 3 Then 'อนุมัติ
   If Area = 1 Then
      frmVerifyAccRight.AccName = "PACKAGE-CENTER_EX-WORKS-PRICE_APPROVE"
      frmVerifyAccRight.AccDesc = "สามารถอนุมัติราคาประกาศสินค้าหน้าโรงงานได้"
   ElseIf Area = 2 Then
      frmVerifyAccRight.AccName = "PACKAGE-CENTER_DELIVERY-COST_APPROVE"
      frmVerifyAccRight.AccDesc = "สามารถอนุมัติราคาค่าขนส่งได้"
   ElseIf Area = 3 Then
      frmVerifyAccRight.AccName = "PACKAGE-CENTER_PROMOTION-PART_APPROVE"
      frmVerifyAccRight.AccDesc = "สามารถอนุมัติราคาประกาศสินค้าโปรโมชั่นได้"
   ElseIf Area = 4 Then
      frmVerifyAccRight.AccName = "PACKAGE-CENTER_PROMOTION-DELIVERY_APPROVE"
      frmVerifyAccRight.AccDesc = "สามารถอนุมัติราคาโปรโมชั่นค่าขนส่งได้"
   End If
End If
Load frmVerifyAccRight
frmVerifyAccRight.Show 1

   If frmVerifyAccRight.GrantRight Then
      TempUserName = frmVerifyAccRight.UserName
      Unload frmVerifyAccRight
      Set frmVerifyAccRight = Nothing
      
      m_ExWorksPrice.EX_WORKS_PRICE_ID = id
      
      Dim tempEWPI As CExWorksPriceItem
      Dim tempEDCI As CExDeliveryCostItem
      Dim tempEPPI As CExPromotionPartItem
      Dim TempEPDI As CExPromotionDlcItem
      Dim CountUpdate As Long
      Dim SumCU As Long
      Dim CountEdit As Long
      Dim DeclareCount As Long
      
      DeclareCount = LoadDeclareCount(m_ExWorksPrice.EX_WORKS_PRICE_ID, Area)
         If Area = 1 Then
            For Each tempEWPI In m_ExWorksPrice.ExWorksPriceItem
               If tempEWPI.LAST_EDIT_FLAG = "Y" Then
                     If lMenuChosen = 1 Then
                        tempEWPI.VERIFY_FLAG = "Y"
                        tempEWPI.VERIFY_NAME = TempUserName
                        TempStr = "ตรวจสอบเอกสาร"
                     ElseIf lMenuChosen = 3 Then
                        tempEWPI.APPROVED_FLAG = "Y"
                        tempEWPI.APPROVED_NAME = TempUserName
                        tempEWPI.LAST_EDIT_FLAG = "N" 'เมื่อสั่งอนุมัติแล้ว ให้เปลี่ยนสถานะการ update เอกสาร
                        tempEWPI.DECLARE_COUNT = DeclareCount + 1
                        TempStr = "อนุมัติเอกสาร"
                     End If
                  CountUpdate = 0
                  Call tempEWPI.UpdateApprovedFlag(lMenuChosen, CountUpdate)
                 SumCU = SumCU + CountUpdate
                 CountEdit = CountEdit + 1
               End If
            Next tempEWPI
         ElseIf Area = 2 Then
            For Each tempEDCI In m_ExWorksPrice.ExDeliveryCost
               If tempEDCI.LAST_EDIT_FLAG = "Y" Then
                     If lMenuChosen = 1 Then
                        tempEDCI.VERIFY_FLAG = "Y"
                        tempEDCI.VERIFY_NAME = TempUserName
                        TempStr = "ตรวจสอบเอกสาร"
                     ElseIf lMenuChosen = 3 Then
                        tempEDCI.APPROVED_FLAG = "Y"
                        tempEDCI.APPROVED_NAME = TempUserName
                        tempEDCI.LAST_EDIT_FLAG = "N" 'เมื่อสั่งอนุมัติแล้ว ให้เปลี่ยนสถานะการ update เอกสาร
                        tempEDCI.DECLARE_COUNT = DeclareCount + 1
                        TempStr = "อนุมัติเอกสาร"
                     End If
                  CountUpdate = 0
                  Call tempEDCI.UpdateApprovedFlag(lMenuChosen, CountUpdate)
                 SumCU = SumCU + CountUpdate
                 CountEdit = CountEdit + 1
               End If
            Next tempEDCI
         ElseIf Area = 3 Then
            For Each tempEPPI In m_ExWorksPrice.ExPromotionPart
               If tempEPPI.LAST_EDIT_FLAG = "Y" Then
                    If lMenuChosen = 1 Then
                        tempEPPI.VERIFY_FLAG = "Y"
                        tempEPPI.VERIFY_NAME = TempUserName
                        TempStr = "ตรวจสอบเอกสาร"
                     ElseIf lMenuChosen = 3 Then
                        tempEPPI.APPROVED_FLAG = "Y"
                        tempEPPI.APPROVED_NAME = TempUserName
                        tempEPPI.LAST_EDIT_FLAG = "N" 'เมื่อสั่งอนุมัติแล้ว ให้เปลี่ยนสถานะการ update เอกสาร
                        tempEPPI.DECLARE_COUNT = DeclareCount + 1
                        TempStr = "อนุมัติเอกสาร"
                     End If
                  CountUpdate = 0
                  Call tempEPPI.UpdateApprovedFlag(lMenuChosen, CountUpdate)
                 SumCU = SumCU + CountUpdate
                 CountEdit = CountEdit + 1
               End If
            Next tempEPPI
         ElseIf Area = 4 Then
            For Each TempEPDI In m_ExWorksPrice.ExPromotionDlc
               If TempEPDI.LAST_EDIT_FLAG = "Y" Then
                    If lMenuChosen = 1 Then
                        TempEPDI.VERIFY_FLAG = "Y"
                        TempEPDI.VERIFY_NAME = TempUserName
                        TempStr = "ตรวจสอบเอกสาร"
                     ElseIf lMenuChosen = 3 Then
                        TempEPDI.APPROVED_FLAG = "Y"
                        TempEPDI.APPROVED_NAME = TempUserName
                        TempEPDI.LAST_EDIT_FLAG = "N" 'เมื่อสั่งอนุมัติแล้ว ให้เปลี่ยนสถานะการ update เอกสาร
                        TempEPDI.DECLARE_COUNT = DeclareCount + 1
                        TempStr = "อนุมัติเอกสาร"
                     End If
                  CountUpdate = 0
                  Call TempEPDI.UpdateApprovedFlag(lMenuChosen, CountUpdate)
                 SumCU = SumCU + CountUpdate
                 CountEdit = CountEdit + 1
               End If
            Next TempEPDI
          End If
      glbErrorLog.LocalErrorMsg = TempStr & "สำเร็จ " & SumCU & " รายการ จากทั้งหมด " & CountEdit & " รายการ"
      glbErrorLog.ShowUserError
   Else
      Unload frmVerifyAccRight
      Set frmVerifyAccRight = Nothing
      Exit Sub
   End If
   Call LoadDeclareCountList(cboDeclareCount, , id, Area)
   Call cmdSearch_Click
End Sub

Private Sub cmdApproved_Click()
   If m_HasModify Or ((Not m_HasModify) And (ShowMode = SHOW_ADD)) Then
      glbErrorLog.LocalErrorMsg = "กรุณาทำการบันทึกข้อมูลให้เรียบร้อยก่อน"
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   Call CONDITION(3)
   Call QueryData(True)
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

   ID2 = GridEX1.Value(2)
   ID1 = GridEX1.Value(1)
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
   If Not ConfirmDelete(GridEX1.Value(3)) Then
      Exit Sub
   End If
   

   
   If Area = 1 Then
         If TabStrip1.SelectedItem.Index = 1 Then
            If m_ExWorksPrice.ExWorksPriceItem.Item(ID2).APPROVED_FLAG = "Y" Then
               glbErrorLog.LocalErrorMsg = MapText("ข้อมูล") & " " & GridEX1.Value(3) & " " & MapText("ได้รับการอนุมัติแล้วไม่สามารถลบได้")
               glbErrorLog.ShowUserError
               Exit Sub
            End If
   
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
   ElseIf Area = 4 Then
      If TabStrip1.SelectedItem.Index = 1 Then
         If ID1 <= 0 Then
            m_ExWorksPrice.ExPromotionDlc.Remove (ID2)
         Else
            m_ExWorksPrice.ExPromotionDlc.Item(ID2).Flag = "D"
         End If

      GridEX1.ItemCount = CountItem(m_ExWorksPrice.ExPromotionDlc)
      GridEX1.Rebind
      m_HasModify = True
      End If
   End If

End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim id As Long
Dim ID2 As Long
Dim OKClick As Boolean
Dim lMenuChosen As Long

Dim RateType As Long
Dim RateType_Cus As Long

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   id = Val(GridEX1.Value(2))
   ID2 = Val(GridEX1.Value(1))
   
'   If m_ExWorksPrice.ExWorksPriceItem.Item(ID).APPROVED_FLAG = "Y" Then
'      glbErrorLog.LocalErrorMsg = MapText("ข้อมูล") & " " & GridEX1.Value(3) & " " & MapText("ได้รับการอนุมัติแล้วไม่สามารถแก้ไขได้")
'      glbErrorLog.ShowUserError
'      Exit Sub
'   End If
   
   lMenuChosen = Val(GridEX1.Value(6))
   OKClick = False
   
   If TabStrip1.SelectedItem.Index = 1 Then
     If Area = 1 Then
      Set frmAddEditExWorksPriceItem.ParentForm = Me
      frmAddEditExWorksPriceItem.SocPartType = 3
      frmAddEditExWorksPriceItem.id = id
      frmAddEditExWorksPriceItem.SocCode = txtPackageNo.Text
      Set frmAddEditExWorksPriceItem.TempCollection = m_ExWorksPrice.ExWorksPriceItem
      Set frmAddEditExWorksPriceItem.m_ExWorkPricesItem = m_ExWorkPricesItem
      frmAddEditExWorksPriceItem.ID_MUM = ID2
      frmAddEditExWorksPriceItem.HeaderText = MapText("แก้ไขสินค้า/บริการ")
      If lMenuChosen = 3 Then
      frmAddEditExWorksPriceItem.HeaderText = MapText("แก้ไขสินค้า/วัตถุดิบ")
      End If
      frmAddEditExWorksPriceItem.canShowGP = canShowGP
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
         frmAddEditExDeliveryCostItem.id = id
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
         frmAddEditExPromotionPartItem.id = id
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
         frmAddEditExPromotionDlcItem.id = id
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
      id = m_ExWorksPrice.EX_WORKS_PRICE_ID
      m_ExWorksPrice.QueryFlag = 1
      QueryData (True)
      m_HasModify = False
   ElseIf lMenuChosen = 3 Then
      If Not SaveData Then
         Exit Sub
      End If
      
      OKClick = True
      Unload Me
   End If
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
   Col.Width = 1000
   Col.Caption = MapText("ราคา/ถุง")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.add '6
   Col.Width = 900
   Col.Caption = MapText("ประเภท")
   Col.TextAlignment = jgexAlignCenter
   
   
   If canShowGP Then
      Set Col = GridEX1.Columns.add '7
      Col.Width = 800
      Col.Caption = MapText("% GP")
     Col.TextAlignment = jgexAlignRight
   End If
   
   Set Col = GridEX1.Columns.add '8
   Col.Width = 1200
   Col.Caption = MapText("ผู้สร้าง")
   Col.TextAlignment = jgexAlignCenter
   
   Set Col = GridEX1.Columns.add '9
   Col.Width = 1200
   Col.Caption = MapText("ผู้แก้ไข")
   Col.TextAlignment = jgexAlignCenter
   
   Set Col = GridEX1.Columns.add '10
   Col.Width = 1200
   Col.Caption = MapText("ผู้ตรวจสอบ")
   Col.TextAlignment = jgexAlignCenter
   
   Set Col = GridEX1.Columns.add '11
   Col.Width = 1200
   Col.Caption = MapText("ผู้อนุมัติ")
   Col.TextAlignment = jgexAlignCenter
   
   Set Col = GridEX1.Columns.add '12
   Col.Width = 1200
   Col.Caption = MapText("แก้ไขราคา")
   Col.TextAlignment = jgexAlignCenter
   
   Set Col = GridEX1.Columns.add '13
   Col.Width = 1200
   Col.Caption = MapText("ประกาศใหม่")
   Col.TextAlignment = jgexAlignCenter
   
   Set Col = GridEX1.Columns.add '14
   Col.Width = 1000
   Col.Caption = MapText("ครั้งที่")
   Col.TextAlignment = jgexAlignCenter
ElseIf Area = 2 Then
   Set Col = GridEX1.Columns.add '1
   Col.Width = 1000
   Col.Caption = MapText("รหัสลูกค้า")
   
   Set Col = GridEX1.Columns.add '2
   Col.Width = 2700
   Col.Caption = MapText("ชื่อลูกค้า")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 1200
   Col.Caption = MapText("รหัสสถานที่")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 2800
   Col.Caption = MapText("สถานที่จัดส่ง")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 1100
   Col.Caption = MapText("ค่าขนส่ง/")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.add '6
   Col.Width = 600
   Col.Caption = MapText("หน่วย")
   Col.TextAlignment = jgexAlignCenter
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = 850
   Col.Caption = MapText("นน.(กก.)")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.add '8
   Col.Width = 1000
   Col.Caption = MapText("ประเภท")
   Col.TextAlignment = jgexAlignCenter
   
   Set Col = GridEX1.Columns.add '9
   Col.Width = 1100
   Col.Caption = MapText("คิดลูกค้า/")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.add '10
   Col.Width = 600
   Col.Caption = MapText("หน่วย")
   Col.TextAlignment = jgexAlignCenter
   
   Set Col = GridEX1.Columns.add '11
   Col.Width = 850
   Col.Caption = MapText("นน.(กก.)")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.add '12
   Col.Width = 0
   Col.Caption = MapText("rate_type")
   
   Set Col = GridEX1.Columns.add '13
   Col.Width = 0
   Col.Caption = MapText("rate_type_cus")
   
   Set Col = GridEX1.Columns.add '8
   Col.Width = 1200
   Col.Caption = MapText("ผู้สร้าง")
   Col.TextAlignment = jgexAlignCenter
   
   Set Col = GridEX1.Columns.add '9
   Col.Width = 1200
   Col.Caption = MapText("ผู้แก้ไข")
   Col.TextAlignment = jgexAlignCenter
   
   Set Col = GridEX1.Columns.add '14
   Col.Width = 1200
   Col.Caption = MapText("ผู้ตรวจสอบ")
   Col.TextAlignment = jgexAlignCenter
   
   Set Col = GridEX1.Columns.add '15
   Col.Width = 1200
   Col.Caption = MapText("ผู้อนุมัติ")
   Col.TextAlignment = jgexAlignCenter
   
   Set Col = GridEX1.Columns.add '16
   Col.Width = 1200
   Col.Caption = MapText("แก้ไขราคา")
   Col.TextAlignment = jgexAlignCenter
   
   Set Col = GridEX1.Columns.add '17
   Col.Width = 1200
   Col.Caption = MapText("ประกาศใหม่")
   Col.TextAlignment = jgexAlignCenter
   
   Set Col = GridEX1.Columns.add '18
   Col.Width = 1000
   Col.Caption = MapText("ครั้งที่")
   Col.TextAlignment = jgexAlignCenter
ElseIf Area = 3 Then
   Set Col = GridEX1.Columns.add '1
   Col.Width = 1000
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
   Col.Width = 900
   Col.Caption = MapText("ส่วนลด/")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.add '6
   Col.Width = 600
   Col.Caption = MapText("หน่วย")
   Col.TextAlignment = jgexAlignCenter
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = 0
   Col.Caption = MapText("rate_type_cus")
   
   Set Col = GridEX1.Columns.add '8
   Col.Width = 1200
   Col.Caption = MapText("ผู้สร้าง")
   Col.TextAlignment = jgexAlignCenter
   
   Set Col = GridEX1.Columns.add '9
   Col.Width = 1200
   Col.Caption = MapText("ผู้แก้ไข")
   Col.TextAlignment = jgexAlignCenter
   
   Set Col = GridEX1.Columns.add '8
   Col.Width = 1200
   Col.Caption = MapText("ผู้ตรวจสอบ")
   Col.TextAlignment = jgexAlignCenter
   
   Set Col = GridEX1.Columns.add '9
   Col.Width = 1200
   Col.Caption = MapText("ผู้อนุมัติ")
   Col.TextAlignment = jgexAlignCenter
   
   Set Col = GridEX1.Columns.add '10
   Col.Width = 1200
   Col.Caption = MapText("แก้ไขราคา")
   Col.TextAlignment = jgexAlignCenter
   
   Set Col = GridEX1.Columns.add '11
   Col.Width = 1200
   Col.Caption = MapText("ประกาศใหม่")
   Col.TextAlignment = jgexAlignCenter
   
   Set Col = GridEX1.Columns.add '12
   Col.Width = 1000
   Col.Caption = MapText("ครั้งที่")
   Col.TextAlignment = jgexAlignCenter
ElseIf Area = 4 Then
   Set Col = GridEX1.Columns.add '3
   Col.Width = 1200
   Col.Caption = MapText("รหัสลูกค้า")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 3500
   Col.Caption = MapText("ชื่อลูกค้า")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 1300
   Col.Caption = MapText("รหัสสถานที่")

   Set Col = GridEX1.Columns.add '6
   Col.Width = 3500
   Col.Caption = MapText("สถานที่จัดส่ง")
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = 900
   Col.Caption = MapText("ส่วนลด/")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.add '8
   Col.Width = 600
   Col.Caption = MapText("หน่วย")
   Col.TextAlignment = jgexAlignCenter
   
   Set Col = GridEX1.Columns.add '9
   Col.Width = 1300
   Col.Caption = MapText("น้ำหนัก(กก.)")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.add '10
   Col.Width = 0
   Col.Caption = MapText("rate_type")
   
   Set Col = GridEX1.Columns.add '8
   Col.Width = 1200
   Col.Caption = MapText("ผู้สร้าง")
   Col.TextAlignment = jgexAlignCenter
   
   Set Col = GridEX1.Columns.add '9
   Col.Width = 1200
   Col.Caption = MapText("ผู้แก้ไข")
   Col.TextAlignment = jgexAlignCenter
   
    Set Col = GridEX1.Columns.add '11
   Col.Width = 1200
   Col.Caption = MapText("ผู้ตรวจสอบ")
   Col.TextAlignment = jgexAlignCenter
   
   Set Col = GridEX1.Columns.add '12
   Col.Width = 1200
   Col.Caption = MapText("ผู้อนุมัติ")
   Col.TextAlignment = jgexAlignCenter
   
   Set Col = GridEX1.Columns.add '13
   Col.Width = 1200
   Col.Caption = MapText("แก้ไขราคา")
   Col.TextAlignment = jgexAlignCenter
   
   Set Col = GridEX1.Columns.add '14
   Col.Width = 1200
   Col.Caption = MapText("ประกาศใหม่")
   Col.TextAlignment = jgexAlignCenter
   
   Set Col = GridEX1.Columns.add '12
   Col.Width = 1000
   Col.Caption = MapText("ครั้งที่")
   Col.TextAlignment = jgexAlignCenter
End If
End Sub


Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      m_ExWorksPrice.EX_WORKS_PRICE_ID = id
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
      txtNote.Text = m_ExWorksPrice.EX_WORKS_PRICE_NOTE
      chkMarket.Value = FlagToCheck(m_ExWorksPrice.EX_WORKS_PRICE_LEVEL)
      uctlDocumentDate.ShowDate = m_ExWorksPrice.EX_WORKS_PRICE_DATE
      uctlFromActiveDate.ShowDate = m_ExWorksPrice.FROM_ACTIVE_DATE
      uctlToValidDate.ShowDate = m_ExWorksPrice.TO_VALID_DATE
      
      If Area = 1 Then
         Call LoadExWorksPriceItem(Nothing, m_ExWorkPricesItem, m_ExWorksPrice.EX_WORKS_PRICE_ID, 2, -1, -1, "")
      ElseIf Area = 2 Then
         Call LoadExDeliveryCusItem(Nothing, m_ExDeliveryCostItem, m_ExWorksPrice.EX_WORKS_PRICE_ID, 4, -1, -1, "")
      ElseIf Area = 3 Then
         Call LoadExPromotionPartItem(Nothing, m_ExPromotionPartItem, m_ExWorksPrice.EX_WORKS_PRICE_ID, 2, -1, -1, "")
      ElseIf Area = 4 Then
         Call LoadExPromotionDlcItem(Nothing, m_ExPromotionDlcItem, m_ExWorksPrice.EX_WORKS_PRICE_ID, 3, -1, -1, "")
      End If
      
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
   If m_HasModify Or ((Not m_HasModify) And (ShowMode = SHOW_ADD)) Then
      glbErrorLog.LocalErrorMsg = "กรุณาทำการบันทึกข้อมูลให้เรียบร้อยก่อน"
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
     Call CONDITION(-1)
     Call QueryData(True)
End Sub

Private Sub cmdVerify_Click()
   If m_HasModify Or ((Not m_HasModify) And (ShowMode = SHOW_ADD)) Then
      glbErrorLog.LocalErrorMsg = "กรุณาทำการบันทึกข้อมูลให้เรียบร้อยก่อน"
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   Call CONDITION(1)
   Call QueryData(True)
End Sub
Function CONDITION(Con As Long)
m_ExWorksPrice.APPROVED_FLAG = ""
m_ExWorksPrice.VERIFY_FLAG = ""
m_ExWorksPrice.LAST_EDIT_FLAG = ""
m_ExWorksPrice.DECLARE_NEW_FLAG = ""
m_ExWorksPrice.DECLARE_COUNT = Val(cboDeclareCount.Text)

If Con = 1 Then 'รายการรอการตรวจสอบ
   If Area = 1 Then
       m_ExWorksPrice.APPROVED_FLAG = ""
       m_ExWorksPrice.VERIFY_FLAG = "N"
       m_ExWorksPrice.DECLARE_NEW_FLAG = ""
   ElseIf Area = 2 Then
       m_ExWorksPrice.APPROVED_FLAG = ""
       m_ExWorksPrice.VERIFY_FLAG = "N"
       m_ExWorksPrice.DECLARE_NEW_FLAG = ""
    ElseIf Area = 3 Then
       m_ExWorksPrice.APPROVED_FLAG = ""
       m_ExWorksPrice.VERIFY_FLAG = "N"
       m_ExWorksPrice.LAST_EDIT_FLAG = "Y"
       m_ExWorksPrice.DECLARE_NEW_FLAG = ""
   ElseIf Area = 4 Then
       m_ExWorksPrice.APPROVED_FLAG = ""
       m_ExWorksPrice.VERIFY_FLAG = "N"
       m_ExWorksPrice.LAST_EDIT_FLAG = "Y"
       m_ExWorksPrice.DECLARE_NEW_FLAG = ""
    End If
ElseIf Con = 3 Then 'รายการรอการอนุมัติ
   If Area = 1 Then
      m_ExWorksPrice.VERIFY_FLAG = "Y"
      m_ExWorksPrice.APPROVED_FLAG = "N"
      m_ExWorksPrice.DECLARE_NEW_FLAG = "Y"
   ElseIf Area = 2 Then
      m_ExWorksPrice.VERIFY_FLAG = "Y"
      m_ExWorksPrice.APPROVED_FLAG = "N"
      m_ExWorksPrice.DECLARE_NEW_FLAG = "Y"
   ElseIf Area = 3 Then
      m_ExWorksPrice.APPROVED_FLAG = "N"
      m_ExWorksPrice.VERIFY_FLAG = "Y"
      m_ExWorksPrice.LAST_EDIT_FLAG = "Y"
      m_ExWorksPrice.DECLARE_NEW_FLAG = "Y"
   ElseIf Area = 4 Then
      m_ExWorksPrice.APPROVED_FLAG = "N"
      m_ExWorksPrice.VERIFY_FLAG = "Y"
      m_ExWorksPrice.LAST_EDIT_FLAG = "Y"
      m_ExWorksPrice.DECLARE_NEW_FLAG = "Y"
   End If
ElseIf Con = 5 Then 'รายการประกาศใหม่ที่อนุมัติแล้ว
     If Area = 1 Then
         m_ExWorksPrice.VERIFY_FLAG = "Y" 'ต้องผ่านการตรวจสอบมาก่อน
         m_ExWorksPrice.APPROVED_FLAG = "Y" 'ต้องเคยอนุมัติมาก่อน
         m_ExWorksPrice.DECLARE_NEW_FLAG = "Y"  'ต้องเป็นเอกสารที่ประกาศใหม่เท่านั้น
      ElseIf Area = 2 Then
         m_ExWorksPrice.VERIFY_FLAG = "Y" 'ต้องผ่านการตรวจสอบมาก่อน
         m_ExWorksPrice.APPROVED_FLAG = "Y" 'ต้องเคยอนุมัติมาก่อน
         m_ExWorksPrice.DECLARE_NEW_FLAG = "Y"  'ต้องเป็นเอกสารที่ประกาศใหม่เท่านั้น
      ElseIf Area = 3 Then
         m_ExWorksPrice.VERIFY_FLAG = "Y" 'ต้องผ่านการตรวจสอบมาก่อน
         m_ExWorksPrice.APPROVED_FLAG = "Y" 'ต้องเคยอนุมัติมาก่อน
         m_ExWorksPrice.DECLARE_NEW_FLAG = "Y"  'ต้องเป็นเอกสารที่ประกาศใหม่เท่านั้น
      ElseIf Area = 4 Then
         m_ExWorksPrice.VERIFY_FLAG = "Y" 'ต้องผ่านการตรวจสอบมาก่อน
         m_ExWorksPrice.APPROVED_FLAG = "Y" 'ต้องเคยอนุมัติมาก่อน
         m_ExWorksPrice.DECLARE_NEW_FLAG = "Y"  'ต้องเป็นเอกสารที่ประกาศใหม่เท่านั้น
      End If
ElseIf Con = 7 Then 'รายการประกาศทั้งหมดที่อนุมัติแล้ว
   If Area = 1 Then
      m_ExWorksPrice.VERIFY_FLAG = "Y" 'ต้องผ่านการตรวจสอบมาก่อน
      m_ExWorksPrice.APPROVED_FLAG = "Y" 'ต้องเคยอนุมัติมาก่อน
   ElseIf Area = 2 Then
      m_ExWorksPrice.VERIFY_FLAG = "Y" 'ต้องผ่านการตรวจสอบมาก่อน
      m_ExWorksPrice.APPROVED_FLAG = "Y" 'ต้องเคยอนุมัติมาก่อน
   ElseIf Area = 3 Then
      m_ExWorksPrice.VERIFY_FLAG = "Y" 'ต้องผ่านการตรวจสอบมาก่อน
      m_ExWorksPrice.APPROVED_FLAG = "Y" 'ต้องเคยอนุมัติมาก่อน
   ElseIf Area = 4 Then
      m_ExWorksPrice.VERIFY_FLAG = "Y" 'ต้องผ่านการตรวจสอบมาก่อน
      m_ExWorksPrice.APPROVED_FLAG = "Y" 'ต้องเคยอนุมัติมาก่อน
   End If
ElseIf Con = 11 Then 'ใบประกาศราคาอาหารสัตว์ในเครือมิตรภาพ
   If Area = 1 Then
      m_ExWorksPrice.VERIFY_FLAG = "Y" 'ต้องผ่านการตรวจสอบมาก่อน
      m_ExWorksPrice.APPROVED_FLAG = "Y" 'ต้องเคยอนุมัติมาก่อน
      m_ExWorksPrice.DECLARE_NEW_FLAG = "Y"  'ต้องเป็นเอกสารที่ประกาศใหม่เท่านั้น
   End If
End If
End Function
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
Dim I As Long
I = 0
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
         Values(6) = ConvertPerPack(EWPI.PART_TYPE)
         I = 7
         If canShowGP Then
            Values(I) = FormatNumber(EWPI.GP_VALUE)
            I = I + 1
          End If
          
         Set Temp_LTK = GetObject("CLoginTracking", m_LoginTracking, Trim(EWPI.CREATE_BY), False)
         If Not Temp_LTK Is Nothing Then
            Values(I) = Temp_LTK.USER_NAME
         Else
            Values(I) = ""
         End If
         
         I = I + 1
         Set Temp_LTK = GetObject("CLoginTracking", m_LoginTracking, Trim(EWPI.MODIFY_BY), False)
         If Not Temp_LTK Is Nothing Then
            Values(I) = Temp_LTK.USER_NAME
         Else
            Values(I) = ""
         End If
         
         I = I + 1
         Values(I) = EWPI.VERIFY_NAME
         I = I + 1
         Values(I) = EWPI.APPROVED_NAME
         I = I + 1
         Values(I) = EWPI.LAST_EDIT_FLAG
         I = I + 1
         Values(I) = EWPI.DECLARE_NEW_FLAG
         I = I + 1
         Values(I) = IIf(EWPI.DECLARE_COUNT <= 0, "", FormatNumber(EWPI.DECLARE_COUNT, 0))
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
         
         Set Temp_LTK = GetObject("CLoginTracking", m_LoginTracking, Trim(EDCI.CREATE_BY), False)
         If Not Temp_LTK Is Nothing Then
            Values(16) = Temp_LTK.USER_NAME
         Else
            Values(16) = ""
         End If
         
         Set Temp_LTK = GetObject("CLoginTracking", m_LoginTracking, Trim(EDCI.MODIFY_BY), False)
         If Not Temp_LTK Is Nothing Then
            Values(17) = Temp_LTK.USER_NAME
         Else
            Values(17) = ""
         End If
         
         Values(18) = EDCI.VERIFY_NAME
         Values(19) = EDCI.APPROVED_NAME
         Values(20) = EDCI.LAST_EDIT_FLAG
         Values(21) = EDCI.DECLARE_NEW_FLAG
         Values(22) = IIf(EDCI.DECLARE_COUNT <= 0, "", FormatNumber(EDCI.DECLARE_COUNT, 0))
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
         Values(8) = ConvertPerPack(EPPI.PART_TYPE)
         Values(9) = EPPI.RATE_TYPE
         
         Set Temp_LTK = GetObject("CLoginTracking", m_LoginTracking, Trim(EPPI.CREATE_BY), False)
         If Not Temp_LTK Is Nothing Then
            Values(10) = Temp_LTK.USER_NAME
         Else
            Values(10) = ""
         End If
         
         Set Temp_LTK = GetObject("CLoginTracking", m_LoginTracking, Trim(EPPI.MODIFY_BY), False)
         If Not Temp_LTK Is Nothing Then
            Values(11) = Temp_LTK.USER_NAME
         Else
            Values(11) = ""
         End If
         
         Values(12) = EPPI.VERIFY_NAME
         Values(13) = EPPI.APPROVED_NAME
         Values(14) = EPPI.LAST_EDIT_FLAG
         Values(15) = EPPI.DECLARE_NEW_FLAG
         Values(16) = IIf(EPPI.DECLARE_COUNT <= 0, "", FormatNumber(EPPI.DECLARE_COUNT, 0))
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
         
         Set Temp_LTK = GetObject("CLoginTracking", m_LoginTracking, Trim(EPDI.CREATE_BY), False)
         If Not Temp_LTK Is Nothing Then
            Values(11) = Temp_LTK.USER_NAME
         Else
            Values(11) = ""
         End If
         
         Set Temp_LTK = GetObject("CLoginTracking", m_LoginTracking, Trim(EPDI.MODIFY_BY), False)
         If Not Temp_LTK Is Nothing Then
            Values(12) = Temp_LTK.USER_NAME
         Else
            Values(12) = ""
         End If
         Values(13) = EPDI.VERIFY_NAME
         Values(14) = EPDI.APPROVED_NAME
         Values(15) = EPDI.LAST_EDIT_FLAG
         Values(16) = EPDI.DECLARE_NEW_FLAG
         Values(17) = IIf(EPDI.DECLARE_COUNT <= 0, "", FormatNumber(EPDI.DECLARE_COUNT, 0))
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
   
   If Not CheckUniqueNs(WORKS_PRICE_ACTIVE_DATE_UNIQUE, Trim(DateToStringInt(uctlFromActiveDate.ShowDate)), id, , 4, , Area) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล ช่วงวันทีมีผล ") & " " & uctlFromActiveDate.ShowDate & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      cmdExit.Enabled = True
      cmdOK.Enabled = True
      Exit Function
   End If
   
   If Not CheckUniqueNs(WORKS_PRICE_ACTIVE_DATE_UNIQUE, Trim(DateToStringInt(uctlToValidDate.ShowDate)), id, , 4, , Area) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล ช่วงวันทีสิ้นสุด ") & " " & uctlToValidDate.ShowDate & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      cmdExit.Enabled = True
      cmdOK.Enabled = True
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_ExWorksPrice.EX_WORKS_PRICE_ID = id
   m_ExWorksPrice.AddEditMode = ShowMode
   m_ExWorksPrice.EX_WORKS_PRICE_LEVEL = Check2Flag(chkMarket.Value)
   m_ExWorksPrice.EX_WORKS_PRICE_CODE = txtPackageNo.Text
   m_ExWorksPrice.EX_WORKS_PRICE_DESC = txtDesc.Text
   m_ExWorksPrice.EX_WORKS_PRICE_NOTE = txtNote.Text
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
      
     If VerifyAccessRight("PACKAGE-CENTER_EX-WORKS-PRICE_SHOW-GP", "แสดงเปอร์เซ็นต์ GP", 2) Then
         canShowGP = True
     End If
   
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
         Call LoadDeclareCountList(cboDeclareCount, , id, Area)
      ElseIf ShowMode = SHOW_ADD Then
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
   Call InitNormalLabel(lblNote, MapText("หมายเหตุ"))
   Call InitNormalLabel(lblDocumentDate, MapText("วันที่ประกาศ"))
   Call InitNormalLabel(lblFromActiveDate, MapText("วันที่มีผล"))
   Call InitNormalLabel(lblToValidDate, MapText("วันที่สิ้นสุด"))
   If Area = 1 Or Area = 3 Then
      Call InitNormalLabel(lblPartNo, MapText("รหัสสินค้า"))
   ElseIf Area = 2 Or Area = 4 Then
      Call InitNormalLabel(lblPartNo, MapText("รหัสลูกค้า"))
   End If
   Call InitNormalLabel(lblDeclareCount, MapText("ประกาศครั้งที่"))
   Call InitCombo(cboDeclareCount)
   
   'lblDeclareCount
   
   Call InitCheckBox(chkMarket, "แสดง")
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
   cmdApprove.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPrint.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdClear.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdVerify.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdApproved.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdAuto, MapText("A"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdApprove, MapText("อื่นๆ"))
   Call InitMainButton(cmdPrint, MapText("พิมพ์"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา (F5)"))
   Call InitMainButton(cmdClear, MapText("เคลียร์ (F4)"))
   Call InitMainButton(cmdVerify, MapText("รอตรวจสอบ"))
   Call InitMainButton(cmdApproved, MapText("รออนุมัติ"))
   
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

''If Area = 2 Or Area = 4 Then
''   Exit Sub
''End If

   If m_HasModify Or ((Not m_HasModify) And (ShowMode = SHOW_ADD)) Then
      glbErrorLog.LocalErrorMsg = "กรุณาทำการบันทึกข้อมูลให้เรียบร้อยก่อน"
      glbErrorLog.ShowUserError
      Exit Sub
   End If

   ReportMode = 1
   
   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("รายการรอการตรวจสอบ", "-", "รายการรอการอนุมัติ", "-", "รายการประกาศใหม่ที่อนุมัติแล้ว", "-", "รายการประกาศทั้งหมดที่อนุมัติแล้ว", "-", "ปรับค่าหน้ากระดาษ", "-", "ใบประกาศราคาอาหารสัตว์ในเครือมิตรภาพ", "-", "ปรับค่าหน้ากระดาษ")
   If lMenuChosen = 0 Then
      Exit Sub
   End If

   Call EnableForm(Me, False)
   Call CONDITION(lMenuChosen)
    If lMenuChosen = 1 Or lMenuChosen = 3 Or lMenuChosen = 5 Or lMenuChosen = 7 Then
         Call QueryData(True)
         ReportKey = "CReportExWPAppr"
         Set Report = New CReportExWPAppr
         ReportFlag = True
         Call Report.AddParam(1, "PREVIEW_TYPE")
   ElseIf lMenuChosen = 11 Then
         Call QueryData(True)
         ReportKey = "CReportExWPDeclare"
         Set Report = New CReportExWPDeclare
         ReportFlag = True
         Call Report.AddParam(1, "PREVIEW_TYPE")
         Call Report.AddParam(Val(cboDeclareCount.Text), "DECLARE_COUNT")
   End If
      If Not Report Is Nothing Then
         Call Report.AddParam(lMenuChosen, "DOCUMENT_TYPE")
         If Area = 1 Then
            Call Report.AddParam(m_ExWorksPrice.ExWorksPriceItem, "EX_WORK_PRICE_APPROVED")
         ElseIf Area = 2 Then
            Call Report.AddParam(m_ExWorksPrice.ExDeliveryCost, "EX_WORK_PRICE_APPROVED")
         ElseIf Area = 3 Then
            Call Report.AddParam(m_ExWorksPrice.ExPromotionPart, "EX_WORK_PRICE_APPROVED")
         ElseIf Area = 4 Then
            Call Report.AddParam(m_ExWorksPrice.ExPromotionDlc, "EX_WORK_PRICE_APPROVED")
         End If
         Call Report.AddParam(m_ExWorksPrice.EX_WORKS_PRICE_CODE, "DOCUMENT_NO")
         Call Report.AddParam(m_ExWorksPrice.EX_WORKS_PRICE_DESC, "DOCUMENT_DESC")
         Call Report.AddParam(m_ExWorksPrice.EX_WORKS_PRICE_NOTE, "DOCUMENT_NOTE")
         Call Report.AddParam(m_ExWorksPrice.EX_WORKS_PRICE_DATE, "DOCUMENT_DATE")
         Call Report.AddParam(m_ExWorksPrice.FROM_ACTIVE_DATE, "FROM_DATE")
         Call Report.AddParam(m_ExWorksPrice.TO_VALID_DATE, "TO_DATE")
         Call Report.AddParam(canShowGP, "CAN_SHOW_GP")
         Call Report.AddParam(Area, "AREA")
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
   If lMenuChosen = 9 Then
      ReportKey = "CReportExWPAppr"
      ReportMode = 1
   ElseIf lMenuChosen = 13 Then
      ReportKey = "CReportExWPDeclare"
      ReportMode = 1
   End If
      
      Set Rc = New CReportConfig
      Rc.REPORT_KEY = ReportKey
      Call Rc.QueryData(m_Rs, iCount)
   
      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         frmReportConfig.ShowMode = SHOW_EDIT
         frmReportConfig.id = Rc.REPORT_CONFIG_ID
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

Private Sub SSCommand2_Click()

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

Private Sub txtNote_Change()
   m_HasModify = True
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

