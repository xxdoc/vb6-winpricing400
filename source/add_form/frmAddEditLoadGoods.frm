VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditLoadGoods 
   ClientHeight    =   10365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13980
   Icon            =   "frmAddEditLoadGoods.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   10365
   ScaleWidth      =   13980
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   10380
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   13995
      _ExtentX        =   24686
      _ExtentY        =   18309
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboLotNo 
         Height          =   315
         Left            =   2040
         TabIndex        =   30
         Top             =   2160
         Width           =   4605
      End
      Begin VB.ComboBox cboBinNo 
         Height          =   315
         ItemData        =   "frmAddEditLoadGoods.frx":27A2
         Left            =   1560
         List            =   "frmAddEditLoadGoods.frx":27A4
         TabIndex        =   22
         Top             =   3240
         Visible         =   0   'False
         Width           =   3885
      End
      Begin VB.ComboBox cboLockNo 
         Height          =   315
         Left            =   1560
         TabIndex        =   19
         Top             =   3720
         Visible         =   0   'False
         Width           =   3885
      End
      Begin VB.ComboBox cboPallet 
         Height          =   315
         Left            =   10680
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   2640
         Width           =   975
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   585
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   13965
         _ExtentX        =   24633
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
      Begin prjFarmManagement.uctlTextBox txtBag 
         Height          =   435
         Left            =   11760
         TabIndex        =   8
         Top             =   2640
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtDesc 
         Height          =   435
         Left            =   2040
         TabIndex        =   11
         Top             =   1200
         Width           =   3885
         _ExtentX        =   9604
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtWeightPerPack 
         Height          =   435
         Left            =   7560
         TabIndex        =   14
         Top             =   720
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPackAmount 
         Height          =   435
         Left            =   11280
         TabIndex        =   15
         Top             =   720
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   767
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   4335
         Left            =   240
         TabIndex        =   20
         Top             =   4485
         Width           =   9795
         _ExtentX        =   17277
         _ExtentY        =   7646
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
         Column(1)       =   "frmAddEditLoadGoods.frx":27A6
         Column(2)       =   "frmAddEditLoadGoods.frx":286E
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditLoadGoods.frx":2912
         FormatStyle(2)  =   "frmAddEditLoadGoods.frx":2A6E
         FormatStyle(3)  =   "frmAddEditLoadGoods.frx":2B1E
         FormatStyle(4)  =   "frmAddEditLoadGoods.frx":2BD2
         FormatStyle(5)  =   "frmAddEditLoadGoods.frx":2CAA
         ImageCount      =   0
         PrinterProperties=   "frmAddEditLoadGoods.frx":2D62
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   240
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   3960
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
         Height          =   4335
         Left            =   10320
         TabIndex        =   24
         Top             =   4485
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   7646
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
         Column(1)       =   "frmAddEditLoadGoods.frx":2F3A
         Column(2)       =   "frmAddEditLoadGoods.frx":3002
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditLoadGoods.frx":30A6
         FormatStyle(2)  =   "frmAddEditLoadGoods.frx":3202
         FormatStyle(3)  =   "frmAddEditLoadGoods.frx":32B2
         FormatStyle(4)  =   "frmAddEditLoadGoods.frx":3366
         FormatStyle(5)  =   "frmAddEditLoadGoods.frx":343E
         ImageCount      =   0
         PrinterProperties=   "frmAddEditLoadGoods.frx":34F6
      End
      Begin MSComctlLib.TabStrip TabStrip2 
         Height          =   555
         Left            =   10320
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   3960
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
      Begin prjFarmManagement.uctlTextBox txtLotNo 
         Height          =   435
         Left            =   10680
         TabIndex        =   28
         Top             =   2160
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtLoadAmount 
         Height          =   435
         Left            =   11280
         TabIndex        =   37
         Top             =   1200
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTxAmount 
         Height          =   435
         Left            =   7560
         TabIndex        =   39
         Top             =   1200
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTotalPallet 
         Height          =   435
         Left            =   10680
         TabIndex        =   42
         Top             =   3120
         Width           =   1245
         _ExtentX        =   3889
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtBrokenAmount 
         Height          =   435
         Left            =   7560
         TabIndex        =   44
         Top             =   1680
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   767
      End
      Begin Threed.SSCommand cmdEdit2 
         Height          =   525
         Left            =   11520
         TabIndex        =   49
         Top             =   9000
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditLoadGoods.frx":36CE
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdjust 
         Height          =   525
         Left            =   2040
         TabIndex        =   48
         Top             =   2880
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdUse 
         Height          =   405
         Left            =   10680
         TabIndex        =   47
         Top             =   1680
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditLoadGoods.frx":39E8
         ButtonStyle     =   3
      End
      Begin VB.Label lblBrokenAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblBrokenAmount"
         Height          =   315
         Left            =   6120
         TabIndex        =   46
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   315
         Left            =   9000
         TabIndex        =   45
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lblTotalPallet 
         Alignment       =   1  'Right Justify
         Caption         =   "lblTotalPallet"
         Height          =   315
         Left            =   8880
         TabIndex        =   43
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label lblTxAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblTxAmount"
         Height          =   315
         Left            =   5880
         TabIndex        =   41
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Label3"
         Height          =   315
         Left            =   9000
         TabIndex        =   40
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   315
         Left            =   12720
         TabIndex        =   38
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblLoadAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblLoadAmount"
         Height          =   315
         Left            =   9840
         TabIndex        =   36
         Top             =   1320
         Width           =   1335
      End
      Begin Threed.SSCommand cmdAdd3 
         Height          =   405
         Left            =   12120
         TabIndex        =   35
         Top             =   3120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditLoadGoods.frx":3D02
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   240
         TabIndex        =   34
         Top             =   9000
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditLoadGoods.frx":401C
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1920
         TabIndex        =   33
         Top             =   9000
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditLoadGoods.frx":4336
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3600
         TabIndex        =   32
         Top             =   9000
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditLoadGoods.frx":4650
         ButtonStyle     =   3
      End
      Begin VB.Label lblLotNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblLotNo"
         Height          =   315
         Left            =   1200
         TabIndex        =   31
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label lblLotNo2 
         Alignment       =   1  'Right Justify
         Caption         =   "lblLotNo2"
         Height          =   315
         Left            =   8880
         TabIndex        =   29
         Top             =   2280
         Width           =   1695
      End
      Begin Threed.SSCommand cmdDelete2 
         Height          =   525
         Left            =   12720
         TabIndex        =   27
         Top             =   9000
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditLoadGoods.frx":496A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd2 
         Height          =   525
         Left            =   10320
         TabIndex        =   26
         Top             =   9000
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditLoadGoods.frx":4C84
         ButtonStyle     =   3
      End
      Begin VB.Label lblBinNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblBinNo"
         Height          =   315
         Left            =   -360
         TabIndex        =   23
         Top             =   3240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblLock 
         Alignment       =   1  'Right Justify
         Caption         =   "lblLock"
         Height          =   315
         Left            =   480
         TabIndex        =   18
         Top             =   3720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   315
         Left            =   12720
         TabIndex        =   17
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   315
         Left            =   9000
         TabIndex        =   16
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblAmount"
         Height          =   315
         Left            =   9840
         TabIndex        =   13
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblWeightPerPack 
         Alignment       =   1  'Right Justify
         Caption         =   "lblWeightPerPack"
         Height          =   315
         Left            =   5880
         TabIndex        =   12
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblPartDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPartDesc"
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label lblPartNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPartNo"
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lblBags 
         Caption         =   "lblBag"
         Height          =   315
         Left            =   13320
         TabIndex        =   7
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label lblPalletNames 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPalletName"
         Height          =   315
         Left            =   8880
         TabIndex        =   6
         Top             =   2760
         Width           =   1695
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   10320
         TabIndex        =   1
         Top             =   9720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditLoadGoods.frx":4F9E
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   12240
         TabIndex        =   2
         Top             =   9720
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditLoadGoods"
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
Private m_CollLotItemWh As Collection
Private m_TempPallets As Collection
Private m_InventoryWHDoc As CInventoryWHDoc
Private m_Lot As cLot

Public TempCollection As Collection
Public TempCollection2 As Collection
Public TempLotItemsWH As Collection
Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public ID2 As Long
Public ID_LOT As Long
Public PART_ITEM_ID As Long
Public PART_NO As String
Public PART_DESC As String
Public PART_TYPE As String
Public WEIGHT_PER_PACK As Double
Public PACK_AMOUNT As Long
Public BARCODE_NO As String
Public LOT_ITEM_WH_ID As Long
Public LOCATION_ID As Long
Public AutoSave As Boolean

Public Area As Long
Public m_IndexCollections As Collection
Public CurrentIndex As Long
Public LotId As Long
Public LotDocId As Long
Public LotDocIdRef As Long
Public HeadPackNo As Long
Public LotItemWhId As Long
Private m_CollPallet As Collection
Private m_CollPallet2 As Collection
Private m_Pallet As CPalletDoc
Private palletIdNotIn As String
Private FlagEvens As String
Public DOCUMENT_TYPE As Long
Private DOCUMENT_TYPE_INPUT As Long
Public DOCUMENT_DATE As Date
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim m_PD As CPalletDoc
Dim m_TempPD As CPalletDoc
Dim I As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
      Call Clear
            
      txtPartNo.Text = m_LotItemWh.PART_NO
      txtDesc.Text = m_LotItemWh.PART_DESC
      txtWeightPerPack.Text = m_LotItemWh.WEIGHT_PER_PACK
      txtTxAmount.Text = m_LotItemWh.TX_AMOUNT
      txtPackAmount.Text = m_LotItemWh.PACK_AMOUNT
      cboLockNo.ListIndex = -1
      cboBinNo.ListIndex = -1
   End If
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub
Private Sub PopulateDestColl2()
Dim Ri As CPalletDoc
Dim D As CPalletDoc
Dim TempPD As CPalletDoc

   If m_CollPallet Is Nothing Then
      Exit Sub
   End If
  
   Set m_CollPallet2 = Nothing
   Set m_CollPallet2 = New Collection
   For Each Ri In m_CollPallet
      
   Set TempPD = GetObject("CPalletDoc", TempCollection.Item(ID_LOT).C_PalletDoc, Trim(Ri.PALLET_DOC_NO & "-" & str(Ri.LOT_DOC_ID)), False)    'lotid
     If Not (TempPD Is Nothing) Then
         If Not (TempPD Is Nothing) Then
            Ri.CAPACITY_AMOUNT = Ri.CAPACITY_AMOUNT - TempPD.CAPACITY_AMOUNT
            Ri.PALLET_CAP_LAST = Ri.CAPACITY_AMOUNT
         End If
      End If
      If Ri.Flag <> "D" And Ri.CAPACITY_AMOUNT > 0 Then
         Set D = New CPalletDoc
         Call D.CopyObject(1, Ri)
            Call m_CollPallet2.add(D, Trim(D.PALLET_DOC_NO & "-" & str(D.LOT_DOC_ID)))  'lotid
         Set D = Nothing
      End If
   Next Ri
End Sub
Private Sub PopulateDestColl()
Dim Ri As CPalletDoc
Dim TempPD As CPalletDoc

   If m_CollPallet Is Nothing Then
      Exit Sub
   End If
  
   Set m_CollPallet2 = Nothing
   Set m_CollPallet2 = New Collection
   Dim TempCapLast As Long
   For Each Ri In m_CollPallet
      Set TempPD = GetObject("CPalletDoc", TempCollection.Item(ID_LOT).C_PalletDoc, Trim(Ri.PALLET_DOC_NO & "-" & str(Ri.LOT_ID)), False)    'lotid
      If Not (TempPD Is Nothing) Then
             If TempPD.Flag = "A" Then
               Ri.TEMP_PALLET_CAP_LAST = Val(Ri.PALLET_CAP_LAST) - Val(TempPD.CAPACITY_AMOUNT)
            ElseIf TempPD.Flag = "E" Then
               Ri.TEMP_PALLET_CAP_LAST = Val(Ri.PALLET_CAP_LAST) - Val(TempPD.TEMP_PALLET_CAP_LAST)
             ElseIf TempPD.Flag = "I" Then
               Ri.TEMP_PALLET_CAP_LAST = Val(Ri.PALLET_CAP_LAST)
            ElseIf TempPD.Flag = "D" Then
               Ri.TEMP_PALLET_CAP_LAST = Val(Ri.PALLET_CAP_LAST) + Val(TempPD.PALLET_CAP_LAST)
             End If
      Else
         Ri.TEMP_PALLET_CAP_LAST = Val(Ri.PALLET_CAP_LAST)
      End If
'      If Ri.Flag <> "D" And Ri.CAPACITY_AMOUNT > 0 Then
      If Ri.CAPACITY_AMOUNT > 0 Then
            Call m_CollPallet2.add(Ri, Trim(Ri.PALLET_DOC_NO & "-" & str(Ri.LOT_ID) & "-" & str(Ri.HEAD_PACK_NO)))    'LOT_DOC_ID
      End If
   Next Ri
End Sub

Private Function SaveData() As Boolean
Dim AMOUNT As Double
Dim LTD As CLotDoc
Dim I As Long
If DOCUMENT_TYPE = 2001 Then
   AMOUNT = Val(txtTxAmount.Text)
Else
   AMOUNT = Val(txtPackAmount.Text)
End If
''''
''''If Val(txtLoadAmount.Text) = 0 Then
''''   glbErrorLog.LocalErrorMsg = MapText("จำนวนที่โหลดต้องไม่มีค่าเป็น 0")
''''   glbErrorLog.ShowUserError
''''   Exit Function
''''End If
I = 0

'Check Lot Empty
For Each LTD In TempCollection
I = I + 1
  If LTD.Flag <> "D" Then
     If Not LTD.C_PalletDoc Is Nothing Then
        If CountItem(LTD.C_PalletDoc) = 0 Then
            If MsgBox("ยังไม่มียอดจำนวนพาเลทจาก LOT : " & LTD.LOT_NO & " ลำดับที่ " & I & vbNewLine & " หากคุณต้องการเพิ่มพาเลทของ LOT นี้ให้กดปุ่ม Yes" & vbNewLine & "หากคุณไม่ต้องการเพิ่มพาเลทของ LOT นี้ให้กดปุ่ม No", vbYesNo, "แจ้งเตือน") = vbNo Then
               LTD.Flag = "D" 'ลบข้อมูล Lot โดย อัตโนมัติ
            Else
               SaveData = False
               Exit Function
            End If
       End If
     End If
  End If
Next LTD


'      If Val(txtLoadAmount.Text) <> AMOUNT Then
''         glbErrorLog.LocalErrorMsg = MapText("ขณะนี้ยอดที่เบิกอาหารไม่เท่ากับยอดที่ต้องการเดิม")
''         glbErrorLog.ShowUserError
'         Exit Function
'      End If
'      If Val(txtLoadAmount.Text) <> AMOUNT Then
'         If MsgBox("ขณะนี้ยอดที่เบิกอาหารไม่เท่ากับยอดที่ต้องการเดิม" & vbNewLine & "ต้องการยอดใหม่กด ปุ่ม Yes" & vbNewLine & "ต้องการยอดเก่ากดปุ่ม No", vbYesNo, "แจ้งเตือน") = vbYes Then
'            If m_LotItemWh.Flag <> "A" Then
'                  m_LotItemWh.Flag = "E"
'            End If
'            If DOCUMENT_TYPE = 2001 Then
'               m_LotItemWh.TX_AMOUNT = Val(txtLoadAmount.Text)
'            Else
'               m_LotItemWh.PACK_AMOUNT = Val(txtLoadAmount.Text)
'               m_LotItemWh.TX_AMOUNT = m_LotItemWh.PACK_AMOUNT * m_LotItemWh.WEIGHT_PER_PACK
'            End If
'         End If
'      End If
      
      
      
          If m_LotItemWh.Flag <> "A" Then
            m_LotItemWh.Flag = "E"
         End If
         If DOCUMENT_TYPE = 2001 Then
            m_LotItemWh.LOAD_TRUE = Val(txtLoadAmount.Text)
         Else
            m_LotItemWh.LOAD_TRUE = Val(txtLoadAmount.Text) * m_LotItemWh.WEIGHT_PER_PACK
         End If
         
         If m_LotItemWh.TX_AMOUNT <> m_LotItemWh.LOAD_TRUE Then
            m_LotItemWh.LOAD_AMOUNT_FLAG = "N"
         ElseIf m_LotItemWh.TX_AMOUNT = m_LotItemWh.LOAD_TRUE And m_LotItemWh.LOAD_TRUE > 0 Then
            m_LotItemWh.LOAD_AMOUNT_FLAG = "Y"
         ElseIf m_LotItemWh.LOAD_TRUE = 0 Then
            m_LotItemWh.LOAD_AMOUNT_FLAG = "N"
         Else
            m_LotItemWh.LOAD_AMOUNT_FLAG = "N"
         End If
         
         m_LotItemWh.LOCK_NO = GridEX1.Value(15)
         m_LotItemWh.PRODUCT_TYPE_ID = GridEX1.Value(13)
         m_LotItemWh.BIN_NO = GridEX1.Value(14)
   
   SaveData = True
End Function
Private Function SaveData2() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long
Dim I As Long
Dim m_Lot As cLot
Dim LIW As CLotItemWH
Dim LTD As CLotDoc
Dim PD As CPalletDoc
Dim TempLTD As CLotDoc
Dim TempPD As CPalletDoc
Dim m_CollPD As CPalletDoc
Dim SumGoodAmount As Long
Dim FIRST As Boolean
Dim FIRST2 As Boolean
Dim Key As String

   If Not VerifyCombo(lblPalletNames, cboPallet, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData2 = True
      Exit Function
   End If
   
   If CountItem(TempCollection) > 0 Then
      Set LTD = TempCollection.Item(ID_LOT)
   Else
      Set LTD = New CLotDoc
   End If
         If txtBag.Text <> "" Then
            Set PD = New CPalletDoc
                  Key = Trim(cboPallet.Text & "-" & str(LotId))
                  Set TempPD = GetObject("CPalletDoc", LTD.C_PalletDoc, Key, False) 'lotid
                  If Not (TempPD Is Nothing) Then
                     'copy ลงใน collection เดิม
                     TempPD.CAPACITY_AMOUNT = TempPD.CAPACITY_AMOUNT + Val(txtBag.Text)
                     TempPD.TEMP_PALLET_CAP_LAST = TempPD.TEMP_PALLET_CAP_LAST + Val(txtBag.Text)
                     TempPD.AddEditMode = SHOW_EDIT
                     If TempPD.Flag <> "A" Then
                        TempPD.Flag = "E"
                     End If
                     LTD.LOT_AMOUNT = LTD.LOT_AMOUNT + TempPD.CAPACITY_AMOUNT   'บวก สินค้าที่เบิกไปแต่ละ Pallet
                  Else
                     PD.PALLET_DOC_NO = cboPallet.Text
                     PD.LOT_DOC_ID = LTD.LOT_DOC_ID
                     PD.Flag = "A"
                     PD.AddEditMode = SHOW_ADD
                     PD.CAPACITY_AMOUNT = CDbl(txtBag.Text)
                     PD.TEMP_PALLET_CAP_LAST = PD.TEMP_PALLET_CAP_LAST + Val(txtBag.Text)
                     PD.TX_TYPE = "E" 'จ่ายออก
                     Call LTD.C_PalletDoc.add(PD, Trim(PD.PALLET_DOC_NO & "-" & str(LotId)))
                     LTD.LOT_AMOUNT = LTD.LOT_AMOUNT + PD.CAPACITY_AMOUNT  'บวก สินค้าที่เบิกไปแต่ละ Pallet
                     Set PD = Nothing
                  End If
          End If
   Set LTD = Nothing
   SaveData2 = True
End Function
Private Sub chkCancelFlag_Click(Value As Integer)
   m_HasModify = True
End Sub
Private Function Clear()
   cboPallet.Clear
   txtBag.Text = ""
End Function
Private Sub cboPallets_Change(Index As Integer)
   m_HasModify = True
End Sub
Private Sub LoadNewCbo(C As ComboBox, Cl As Collection)
Dim TempData As CLotDoc
Dim I As Long
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If
 For Each TempData In Cl
   If Not (C Is Nothing) Then
         I = I + 1
         C.AddItem (TempData.LOT_NO & "-" & Format(TempData.TIME_PACK_BEGIN, "HH:mm") & " " & TempData.BIN_NAME & " " & TempData.LOCK_NAME)
         C.ItemData(I) = TempData.LOT_ID & TempData.LOT_DOC_ID
      End If
   Next TempData
End Sub
Private Sub cboLotNo_Click()
Dim LTD As CLotDoc
Dim TempLTD As CLotDoc
Dim Key As String
Dim Key2 As String

Call EnableForm(Me, False)

   Key = Trim(str(cboLotNo.ItemData(Minus2Zero(cboLotNo.ListIndex))))
   Set LTD = GetObject("CLotDoc", m_CollLotItemWh, Key, False)
   If Not (LTD Is Nothing) Then
      Key2 = Trim(str(LTD.LOT_ID) & "-" & str(LTD.LOT_DOC_ID))
     Set TempLTD = GetObject("CLotDoc", TempCollection, Key2, False)
     If TempLTD Is Nothing Then 'ถ้ายังไม่มี lot
         LTD.AddEditMode = SHOW_ADD
         LTD.Flag = "A"
         LTD.LOT_DOC_ID = 0 'เป็นการเพิ่มใหม่จะไม่เอา id lot เดิมมาใช้เด็ดขาด
         Call TempCollection.add(LTD, Key2)
    End If
   Call m_CollLotItemWh.Remove(Key)
   Call LoadNewCbo(cboLotNo, m_CollLotItemWh)

   Call TabStrip1_Click
   End If
   
   Call EnableForm(Me, True)
End Sub
Private Sub cboPallet_Change()
   m_HasModify = True
End Sub
Private Sub cboPallet_Click()
   m_HasModify = True
   Call getPalletFromLot
End Sub
Private Sub getPalletFromLot()
   Dim I As Long
   Dim ID3 As String
   Dim LTD As CLotDoc
   Dim PD As CPalletDoc
   Dim TempPD As CPalletDoc
   m_HasModify = True

   If cboPallet.ListIndex > -1 Then
      ID3 = cboPallet.Text
         Set m_Pallet = GetObject("CPalletDoc", m_CollPallet2, Trim(ID3 & "-" & str(LotId) & "-" & str(HeadPackNo)), False)  'LotDocId
      If Not m_Pallet Is Nothing Then
          If DOCUMENT_TYPE = 2001 Then
             txtBag.Text = m_Pallet.TEMP_PALLET_CAP_LAST 'Format(m_Pallet.TEMP_PALLET_CAP_LAST, "#.000") 'm_Pallet.TEMP_PALLET_CAP_LAST
          Else
            txtBag.Text = m_Pallet.TEMP_PALLET_CAP_LAST
         End If
      Else
         txtBag.Text = ""
      End If
   End If
End Sub

Private Sub AddToList()
   ShowMode = SHOW_ADD
   If Not SaveData2 Then
     Exit Sub
   End If
   Call TabStrip2_Click
End Sub

Private Sub cmdAdd_Click()
   If FlagEvens = "D" Then
         glbErrorLog.LocalErrorMsg = MapText("โปรแกรมจะปิดหน้าต่างนี้ เพื่อรีเซ็ต ข้อมูลเดิม ก่อนเพิ่มข้อมูลล๊อตใหม่")
         glbErrorLog.ShowUserError
         AutoSave = True
         Call cmdOK_Click
         Exit Sub
   End If
   ShowMode = SHOW_ADD
   If DOCUMENT_TYPE_INPUT = 14 Then 'เป็นการ Load BAG
      Call LoadLotByPartItem(cboLotNo, m_CollLotItemWh, , -1, DOCUMENT_DATE, , m_LotItemWh.PART_ITEM_ID, 2, 1, 1, "I", TempCollection, , DOCUMENT_TYPE_INPUT, , 109)
   ElseIf DOCUMENT_TYPE_INPUT = 13 Then 'เป็นการ Load BULK
      Call LoadLotByPartItem(cboLotNo, m_CollLotItemWh, , -1, DOCUMENT_DATE, , m_LotItemWh.PART_ITEM_ID, 2, 1, 1, "I", TempCollection, , DOCUMENT_TYPE_INPUT, , 110)
  End If
   cboLotNo.Enabled = True
   Call TabStrip1_Click
End Sub

Private Sub cmdAdd3_Click()
Dim Value As Double
Dim CheckVerify As Boolean

Call EnableForm(Me, False)
     If cboPallet.ListCount <= 1 Then
         glbErrorLog.LocalErrorMsg = MapText("ไม่มียอดนี้อยู่ในสต๊อกจริงแล้ว หากต้องการยอดนี้ให้ติดต่อผู้ดูแลระบบ")
         glbErrorLog.ShowUserError
         CheckVerify = True
'         Exit Sub
     End If
   If cboPallet.ListIndex > -1 Then
       Set m_Pallet = GetObject("CPalletDoc", m_CollPallet2, Trim(cboPallet.Text & "-" & str(LotId) & "-" & str(HeadPackNo)), False) 'lotid
       If Not m_Pallet Is Nothing Then
          If Val(txtBag.Text) > Val(m_Pallet.PALLET_CAP_LAST) Then
             txtBag.Text = m_Pallet.PALLET_CAP_LAST
             MsgBox "จำนวนที่ป้อนมากกว่าจำนวนที่มีอยู่จริง"
              Call EnableForm(Me, True)
             Exit Sub
         Else
            m_Pallet.TEMP_PALLET_CAP_LAST = m_Pallet.TEMP_PALLET_CAP_LAST - Val(txtBag.Text)
            m_Pallet.PALLET_CAP_LAST = m_Pallet.TEMP_PALLET_CAP_LAST
            
          End If
          Call SetTotal
       End If
   Else
      CheckVerify = True
    End If
    
    If CheckVerify Then
      If DOCUMENT_TYPE = 2000 Or DOCUMENT_TYPE = 2001 Then
         If Not VerifyAccessRight("INVENTORY-WH_EXPORT" & "_" & DOCUMENT_TYPE & "_" & "ADD" & "_ADD-OVER", "เพิ่มข้อมูลสินค้าเกินสต๊อกได้") Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
      End If
    End If
    
   If ChekPackAmount() Then
      Call AddToList
   End If
   
   If ID_LOT > 0 Then
      Call LoadPalletDocAmount2(cboPallet, m_CollPallet2, LotId, 2, , 2, "I", , , TempCollection.Item(ID_LOT).C_PalletDoc)
   End If
   Call EnableForm(Me, True)
End Sub
Function ChekPackAmount() As Boolean
Dim TempData As Double
Dim TempData2 As Double
Dim NewValue As Double
If FlagEvens <> "D" Then
   TempData2 = GetTotalAmount(TempCollection) + Val(txtBag.Text)
Else
   TempData2 = GetTotalAmount(TempCollection)
End If
TempData = TempData2 - Val(txtLoadAmount.Text)

 If DOCUMENT_TYPE = 2001 Then
   NewValue = Val(txtTxAmount.Text)
Else
   NewValue = Val(txtPackAmount.Text)
End If

  If Val(TempData2) > Val(NewValue) Then
      glbErrorLog.LocalErrorMsg = MapText("ขณะนี้คุณได้เบิกเกินยอดที่ต้องการแล้ว")
      glbErrorLog.ShowUserError
      ChekPackAmount = False
      Exit Function
  End If

   txtBrokenAmount.Text = Val(NewValue) - Val(TempData2)
   txtLoadAmount.Text = Trim(str(TempData2))
   ChekPackAmount = True

End Function

Private Sub cmdAdjust_Click()
   Load frmAdjustInventoryWH
   frmAdjustInventoryWH.DocumentType = DOCUMENT_TYPE_INPUT
   frmAdjustInventoryWH.PartNo = txtPartNo.Text
   frmAdjustInventoryWH.Show 1
   
   OKClick = frmAdjustInventoryWH.OKClick
   
   Unload frmAdjustInventoryWH
   Set frmAdjustInventoryWH = Nothing
   
   Call cmdAdd_Click
End Sub

Private Sub cmdDelete_Click()
Dim ID1 As Long
Dim PD As CPalletDoc
Dim LTD As CLotDoc
FlagEvens = "D"
   If Not cmdDelete.Enabled Then
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   cboLotNo.Enabled = False
   ID1 = GridEX1.Value(2)
   If TabStrip1.SelectedItem.Index = 1 Then
     Set LTD = TempCollection.Item(ID1)
         LTD.Flag = "D"
      For Each PD In LTD.C_PalletDoc
         PD.Flag = "D"
      Next PD
      
      Call TabStrip1_Click
       Call ChekPackAmount
      Call TabStrip2_Click
      
      cboPallet.Clear
      txtBag.Text = ""
      m_HasModify = True
   End If
End Sub

Private Sub cmdDelete2_Click()
Dim ID1 As Long
Dim ID2 As Long
Dim TempPD As CPalletDoc
Dim LTD As CLotDoc
Dim PD As CPalletDoc
Dim NewValue As Double
Dim TempData As Double
Dim TempData2 As Double
   If Not cmdDelete2.Enabled Then
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX2.Value(1)) Then
      Exit Sub
   End If
   
   ID1 = GridEX1.Value(2)
   ID2 = GridEX2.Value(2)
   Set PD = TempCollection.Item(ID_LOT).C_PalletDoc.Item(ID2)
'''   If PD.BALANCE_FLAG_F_I = "N" Then
'''      MsgBox "ไม่สามารถลบการเบิกจาก พาเลท " & PD.PALLET_DOC_NO & " นี้ได้เนื่องจากถูกปรับยอดไปแล้ว", vbOKOnly, "แจ้งเตือน"
'''      Exit Sub
'''   End If
   If TabStrip2.SelectedItem.Index = 1 Then
      If PD.PALLET_DOC_ID > 0 Then
         PD.Flag = "D"
      Else
        TempCollection.Item(ID_LOT).C_PalletDoc.Remove (ID2)
      End If
      Call TabStrip2_Click
      Call cmdAdd2_Click
      
TempData2 = GetTotalAmount(TempCollection)
txtLoadAmount.Text = TempData2
If DOCUMENT_TYPE = 2001 Then
   NewValue = CDbl(txtTxAmount.Text)
Else
   NewValue = Val(txtPackAmount.Text)
End If
txtBrokenAmount.Text = Val(NewValue) - Val(TempData2)
m_HasModify = True
   End If
End Sub

Private Sub cmdEdit_Click()
ShowMode = SHOW_EDIT
cboLotNo.Enabled = False
cboBinNo.Enabled = False
cboLockNo.Enabled = False
End Sub

Private Sub cmdAdd2_Click()
  If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   Call LoadPallet
'   cboPallet.SetFocus
End Sub

Private Sub cmdEdit2_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean
Dim LTD As CLotDoc

   If Not VerifyGrid(GridEX2.Value(1)) Then
      Exit Sub
   End If
   
   FlagEvens = "D"
   ID = Val(GridEX2.Value(2))
   Set LTD = TempCollection.Item(ID_LOT)

   frmAddEditLocation.ID = ID
   If DOCUMENT_TYPE_INPUT = 14 Then
      frmAddEditLocation.HeaderText = MapText("แก้ไขจำนวนถุง")
   ElseIf DOCUMENT_TYPE_INPUT = 13 Then
      frmAddEditLocation.HeaderText = MapText("แก้ไขน้ำหนัก")
   End If
   frmAddEditLocation.DocumentTypeInput = DOCUMENT_TYPE_INPUT
   frmAddEditLocation.Area = 3
   frmAddEditLocation.FlagNotEditOver = True
   Set frmAddEditLocation.TempCollection = LTD.C_PalletDoc
   frmAddEditLocation.ShowMode = SHOW_EDIT
   Load frmAddEditLocation
   frmAddEditLocation.Show 1

   OKClick = frmAddEditLocation.OKClick

   Unload frmAddEditLocation
   Set frmAddEditLocation = Nothing
   
   Call checkReturnValue(LTD.C_PalletDoc, ID)

   If OKClick Then
      If ChekPackAmount() Then
         Call TabStrip2_Click
      End If
   End If

End Sub
Private Sub checkReturnValue(Cl As Collection, ID As Long)
  Dim t_PD1 As CPalletDoc
  Dim t_PD2 As CPalletDoc
     Set t_PD1 = Cl.Item(ID)
     If Not t_PD1 Is Nothing Then
          Set t_PD2 = GetObject("CPalletDoc", m_CollPallet2, Trim(t_PD1.PALLET_DOC_NO & "-" & str(LotId) & "-" & str(HeadPackNo)), False) 'lotid
           If Not t_PD2 Is Nothing Then
               t_PD2.TEMP_PALLET_CAP_LAST = t_PD2.TEMP_PALLET_CAP_LAST + (t_PD1.TEMP_PALLET_CAP_LAST - t_PD1.CAPACITY_AMOUNT)
               t_PD2.PALLET_CAP_LAST = t_PD2.TEMP_PALLET_CAP_LAST
           End If
    End If
    Call SetTotal
    Call LoadPalletDocAmount2(cboPallet, m_CollPallet2, LotId, 2, , 2, "I", , , TempCollection.Item(ID_LOT).C_PalletDoc)
End Sub
Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If

   OKClick = True
   Unload Me
End Sub

Private Sub cmdUse_Click()
   txtBag.Text = txtBrokenAmount.Text
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      
      If DOCUMENT_TYPE = 2000 Then 'ถ้าเป็น Bag
         DOCUMENT_TYPE_INPUT = 14
      ElseIf DOCUMENT_TYPE = 2001 Then 'ถ้าเป็น Bulk
         DOCUMENT_TYPE_INPUT = 13
      End If
      
      If Not TempLotItemsWH Is Nothing Then
         Set TempCollection = TempLotItemsWH.Item(ID).C_LotDoc
         Set m_LotItemWh = TempLotItemsWH.Item(ID)
      Else
         Set TempCollection = New Collection
         Set m_LotItemWh = New CLotItemWH
      End If

      Call QueryData(True)
      Call TabStrip1_Click
      Call ChekPackAmount
      Call TabStrip2_Click
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
   
   If DOCUMENT_TYPE = 2001 Then
      Set Col = GridEX1.Columns.add '5
      Col.Width = 0
      Col.Caption = MapText("ล๊อค")
   Else
      Set Col = GridEX1.Columns.add '5
      Col.Width = 1000
      Col.Caption = MapText("ล๊อค")
   End If
   
   Set Col = GridEX1.Columns.add '6
   Col.Width = 0
   Col.Caption = MapText("LOT_DOC_ID")
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = 1400
   Col.Caption = MapText("วันที่ผลิต")
   
   If DOCUMENT_TYPE = 2001 Then
      Set Col = GridEX1.Columns.add '8
      Col.Width = 0
      Col.Caption = MapText("วันที่/เวลา แพ็ค")
   Else
      Set Col = GridEX1.Columns.add '8
      Col.Width = 2000
      Col.Caption = MapText("วันที่/เวลา แพ็ค")
   End If
   
   Set Col = GridEX1.Columns.add '9
   Col.Width = 0
   Col.Caption = MapText("LOT_DOC_ID_REF")
   
   Set Col = GridEX1.Columns.add '10
   Col.Width = 0
   Col.Caption = MapText("HEAD_PACK_NO")
   
   Set Col = GridEX1.Columns.add '11
   Col.Width = 0
   Col.Caption = MapText("LOT_ITEM_WH_ID")
   
   Set Col = GridEX1.Columns.add '12
   Col.Width = 0
   Col.Caption = MapText("LOT_ID")
   
   Set Col = GridEX1.Columns.add '13
   Col.Width = 0
   Col.Caption = MapText("PRODUCT_TYPE_ID")
   
   Set Col = GridEX1.Columns.add '14
   Col.Width = 0
   Col.Caption = MapText("BIN_NO")
   
   Set Col = GridEX1.Columns.add '15
   Col.Width = 0
   Col.Caption = MapText("LOCK_NO")
   
   Set Col = GridEX1.Columns.add '16
   Col.Width = 2500
   Col.Caption = MapText("สถานที่จัดเก็บ")

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
   
   If DOCUMENT_TYPE = 2001 Then
      Set Col = GridEX2.Columns.add '3
      Col.Width = 0
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("")
      
      Set Col = GridEX2.Columns.add '4
      Col.Width = 2300
      Col.Caption = MapText("จำนวน")
   Else
      Set Col = GridEX2.Columns.add '3
      Col.Width = 1500
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ชื่อพาเลท")
      
      Set Col = GridEX2.Columns.add '4
      Col.Width = 1500
      Col.Caption = MapText("จำนวนถุง")
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
      Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 118 Then
      Call LoadPallet
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
      Call cmdDelete_Click
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
  Call cmdOK_Click
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_LotItemWh = Nothing
   Set m_CollPallet = Nothing
   Set m_CollPallet2 = Nothing
   Set m_CollLotItemWh = Nothing
   Set m_TempPallets = Nothing
End Sub
Private Sub InitFormLayout()
Dim I As Long
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   cmdAdjust.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblPartNo, MapText("รหัสสินค้า"))
   Call InitNormalLabel(lblPartDesc, MapText("รายละเอียด"))
   Call InitNormalLabel(lblWeightPerPack, MapText("นน./ถุง"))
   Call InitNormalLabel(lblAmount, MapText("จำนวนที่เบิก"))
   Call InitNormalLabel(lblTxAmount, MapText("น้ำหนักรวม"))
   Call InitNormalLabel(lblLoadAmount, MapText("จำนวนที่โหลด"))
   Call InitNormalLabel(lblBrokenAmount, MapText("จำนวนที่ขาด"))
   Call InitNormalLabel(Label2, MapText("ถุง"))
   Call InitNormalLabel(Label3, MapText("กก."))
   Call InitNormalLabel(Label1, MapText("กก."))
   Call InitNormalLabel(Label4, MapText("ถุง"))
   Call InitNormalLabel(Label5, MapText("ถุง"))
   Call InitNormalLabel(lblLotNo, MapText("ล๊อต"))
   Call InitNormalLabel(lblLotNo2, MapText("ล๊อต"))
   Call InitNormalLabel(lblBinNo, MapText("ถัง"))
   Call InitNormalLabel(lblLock, MapText("ล๊อค"))
   Call InitNormalLabel(lblTotalPallet, MapText("คงเหลือทั้งล๊อต"))
   Call InitMainButton(cmdAdjust, MapText("คำนวณยอดคงเหลือ"))
   
   Call InitCombo(cboLotNo)
   Call InitCombo(cboBinNo)
   Call InitCombo(cboLockNo)
   cboLotNo.Enabled = False
   cboBinNo.Enabled = False
   cboLockNo.Enabled = False

   Call InitCombo(cboPallet)
   If DOCUMENT_TYPE = 2001 Then
      Call InitNormalLabel(lblPalletNames, MapText("จำนวนที่เหลือ "))
   Else
      Call InitNormalLabel(lblPalletNames, MapText("พาเลทที่ "))
   End If
   Call txtBag.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   Call txtTotalPallet.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   Call txtBrokenAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   txtTotalPallet.Enabled = False
   Call InitNormalLabel(lblBags, MapText("ถุง"))
  
   txtPartNo.Enabled = False
   txtDesc.Enabled = False
   txtLotNo.Enabled = False
   txtWeightPerPack.Enabled = False
   txtTxAmount.Enabled = False
   txtPackAmount.Enabled = False
   txtLoadAmount.Enabled = False
   txtBrokenAmount.Enabled = False
   
   If DOCUMENT_TYPE = 2001 Then
      lblWeightPerPack.Visible = False
      txtWeightPerPack.Visible = False
      Label3.Visible = False
      lblAmount.Visible = False
      txtPackAmount.Visible = False
      Label4.Visible = False
      Call InitNormalLabel(Label2, MapText("กก."))
      Call InitNormalLabel(lblBags, MapText("กก."))
      Call InitNormalLabel(Label5, MapText("กก."))
   End If
   
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
   If DOCUMENT_TYPE = 2001 Then
      TabStrip2.Tabs.add().Caption = MapText("จำนวนอาหารที่เบิก")
   Else
      TabStrip2.Tabs.add().Caption = MapText("รายการพาเลทที่วาง")
   End If

   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19

   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   cmdAdd2.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit2.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete2.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   cmdAdd3.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   cmdUse.Picture = LoadPicture(glbParameterObj.NormalButton1)


   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdAdd2, MapText("เพิ่ม"))
   Call InitMainButton(cmdUse, MapText("ใช้ค่า"))
   Call InitMainButton(cmdAdd3, MapText("เพิ่ม"))
   Call InitMainButton(cmdEdit2, MapText("แก้ไข"))
   Call InitMainButton(cmdDelete2, MapText("ลบ"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("บันทึก ออก"))
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   Call cmdOK_Click
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   OKClick = False
   Call InitFormLayout
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_LotItemWh = New CLotItemWH
   Set m_CollPallet = New Collection
   Set m_CollPallet2 = New Collection
   Set m_CollLotItemWh = New Collection
   Set m_TempPallets = New Collection
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
   Call TabStrip2_Click
End Sub
Private Sub LoadPallet()
Dim LTD As CLotDoc
   If Not CountItem(TempCollection) > 0 Then
      Exit Sub
   End If

   ID_LOT = GridEX1.Value(2)
   HeadPackNo = GridEX1.Value(10)
   LotItemWhId = GridEX1.Value(11)
   LotId = GridEX1.Value(12)
   
   Set LTD = GetItem(TempCollection, ID_LOT, 0)
   If Not (LTD Is Nothing) Then
      If LTD.Flag = "A" Then
         LotDocId = GridEX1.Value(6)
      Else
         LotDocId = GridEX1.Value(9) 'LOT_DOC_ID_REF
      End If
   End If
   LotDocIdRef = GridEX1.Value(9) 'LOT_DOC_ID_REF
   
   Call Clear
   Call LoadPalletDocAmount(Nothing, m_CollPallet, LotId, 2, , 2, "I", , , LotDocId, HeadPackNo, LotItemWhId, DOCUMENT_TYPE_INPUT, PART_ITEM_ID, LotDocIdRef, LOCATION_ID) 'load ข้อมูลคงเหลือทั้งหมดของ lot นี้
   Call PopulateDestColl 'เอาค่าจาก Collection มาวางไว้ที่ m_CollPallet2
   Call LoadPalletDocAmount2(cboPallet, m_CollPallet2, LotId, 2, , 2, "I", , , TempCollection.Item(ID_LOT).C_PalletDoc) 'เอาค่าจาก m_CollPallet2 มาแสดงที่ cboPallet
   
   Call SetTotal
   If TempCollection.Count > 0 Then
      Call TabStrip2_Click
   End If
End Sub
Private Sub SetTotal()
   txtTotalPallet.Text = GetTotalAmountPallet(m_CollPallet2)
End Sub
Private Sub GridEX1_DblClick()
   Call EnableForm(Me, False)
   Call TabStrip2_Click
   Call cmdAdd2_Click
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 And KeyCode = 13 Then
      Call LoadPallet
      KeyCode = 0
End If
End Sub

Private Sub GridEX1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
Dim TempID1 As Long
Dim BD As CBillingDoc
Dim IsOK As Boolean
Dim OKClick As Boolean

   If GridEX1.ItemCount <= 0 Then
         Exit Sub
   End If
   
   TempID1 = GridEX1.Value(1)
   If Button = 2 Then
      Set oMenu = New cPopupMenu
     lMenuChosen = oMenu.Popup("ตรวจสอบความผิดพลาด")
      If lMenuChosen = 0 Then
         Exit Sub
      End If
      Set oMenu = Nothing
   Else
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
      If lMenuChosen = 1 Then
         Dim LTD As CLotDoc
         If Not CountItem(TempCollection) > 0 Then
            Exit Sub
         End If
      
         ID_LOT = GridEX1.Value(2)
         HeadPackNo = GridEX1.Value(10)
         LotItemWhId = GridEX1.Value(11)
         LotId = GridEX1.Value(12)

         
         Set LTD = GetItem(TempCollection, ID_LOT, 0)
         If Not (LTD Is Nothing) Then
            If LTD.Flag = "A" Then
               LotDocId = GridEX1.Value(6)
            Else
               LotDocId = GridEX1.Value(9) 'LOT_DOC_ID_REF
            End If
         End If
         LotDocIdRef = GridEX1.Value(9) 'LOT_DOC_ID_REF
      
         frmShowEvents.LotId = ID_LOT
         frmShowEvents.HeadPackNo = HeadPackNo
         frmShowEvents.LotItemWhId = LotItemWhId
         frmShowEvents.LotId = LotId
         frmShowEvents.LotDocId = LotDocId
         frmShowEvents.LotDocIdRef = LotDocIdRef
         frmShowEvents.DOCUMENT_TYPE_INPUT = DOCUMENT_TYPE_INPUT
         frmShowEvents.PART_ITEM_ID = PART_ITEM_ID
         frmShowEvents.LOCATION_ID = LOCATION_ID
'         frmShowEvents.KeyType = 3
         frmShowEvents.PART_NO = txtPartNo.Text
         
         frmShowEvents.HeadertText = MapText("ข้อมูลในการตรวจสอบ")
         Load frmShowEvents
         frmShowEvents.Show 1
         
         Unload frmShowEvents
         Set frmShowEvents = Nothing
   
   End If
   Call EnableForm(Me, True)

End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long
Dim I As Long

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Index = 1 Then
      If RowIndex <= 0 Then
         Exit Sub
      End If
      
      Dim LTD As CLotDoc
      If TempCollection.Count <= 0 Then
         Exit Sub
      End If
      Set LTD = GetItem(TempCollection, RowIndex, RealIndex)
      Values(1) = LTD.LOT_DOC_ID
      Values(2) = RealIndex
      Values(3) = LTD.LOT_NO
      Values(4) = LTD.BIN_NAME
      Values(5) = LTD.LOCK_NAME
      Values(6) = LTD.LOT_DOC_ID
      If LTD.DOCUMENT_TYPE = 15 Or LTD.DOCUMENT_TYPE = 16 Then
         Values(7) = DateToStringExtEx2(LTD.BL_START_DATE)
      Else
         Values(7) = DateToStringExtEx2(LTD.START_DATE)
      End If
      Values(8) = DateToStringExtEx2(LTD.PACK_DATE) & " " & Format(LTD.TIME_PACK_BEGIN, "HH:mm")
      Values(9) = LTD.LOT_DOC_ID_REF
      Values(10) = LTD.HEAD_PACK_NO
      Values(11) = LTD.LOT_ITEM_WH_ID
      Values(12) = LTD.LOT_ID
      Values(13) = LTD.PRODUCT_TYPE_ID
      Values(14) = LTD.BIN_NO
      Values(15) = LTD.LOCK_NO
      Values(16) = LTD.LOCATION_NAME
   End If
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub SSCommand3_Click()

End Sub

Private Sub GridEX2_DblClick()
Call cmdEdit2_Click
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
      
      Dim LTD As CLotDoc
      Dim PD As CPalletDoc
      Set LTD = TempCollection.Item(ID_LOT)
      If CountItem(LTD.C_PalletDoc) <= 0 Then
         Exit Sub
      End If
      Set PD = GetItem(LTD.C_PalletDoc, RowIndex, RealIndex)
      If PD Is Nothing Then
         Exit Sub
      End If
      If PD.Flag = "D" Then
         Exit Sub
      End If
      Values(1) = PD.PALLET_DOC_ID
      Values(2) = RealIndex
      Values(3) = PD.PALLET_DOC_NO
      Values(4) = PD.CAPACITY_AMOUNT
   End If
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub TabStrip1_Click()
   GridEX1.Visible = False
   If TabStrip1.SelectedItem.Index = 1 Then
      GridEX1.Visible = True
      Call InitGrid1
      GridEX1.ItemCount = CountItem(TempCollection)
      GridEX1.Rebind
   End If
End Sub
Private Sub cboPallet_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
     txtBag.SetFocus
Else
   KeyAscii = 0
End If
End Sub

Private Sub TabStrip2_Click()
Dim LTD As CLotDoc
   If TabStrip2.SelectedItem.Index = 1 Then
      Call InitGrid2
      ID_LOT = GridEX1.Value(2)
      txtLotNo.Text = GridEX1.Value(3)
      If TempCollection.Count > 0 Then
         If ID_LOT > 0 Then
            GridEX2.Visible = True
            GridEX2.ItemCount = CountItem(TempCollection.Item(ID_LOT).C_PalletDoc)
            GridEX2.Rebind
         End If
      End If
   End If
End Sub

Private Sub txtBag_Change()
   m_HasModify = True
End Sub

Private Sub txtBag_KeyPress(KeyAscii As Integer)
m_HasModify = True
 KeyAscii = CheckIntAscii(KeyAscii)
  If KeyAscii = 13 Then
     Call cmdAdd3_Click
   End If
End Sub

