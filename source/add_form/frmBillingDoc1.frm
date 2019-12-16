VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmBillingDoc1 
   BackColor       =   &H80000000&
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   11925
   Icon            =   "frmBillingDoc1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   11925
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboOtherFilter 
         Height          =   315
         Left            =   6180
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   1680
         Width           =   5505
      End
      Begin prjFarmManagement.uctlDate uctlDocumentDate 
         Height          =   405
         Left            =   6180
         TabIndex        =   1
         Top             =   750
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin VB.ComboBox cboOrderType 
         Height          =   315
         Left            =   6180
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2550
         Width           =   2625
      End
      Begin VB.ComboBox cboOrderBy 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2550
         Width           =   2985
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   18
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtDocumentNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   720
         Width           =   2985
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtCustomerCode 
         Height          =   435
         Left            =   1860
         TabIndex        =   2
         Top             =   1170
         Width           =   2985
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtAccountNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   3
         Top             =   1650
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPartNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   5
         Top             =   2100
         Width           =   2985
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtFeatureCode 
         Height          =   435
         Left            =   6180
         TabIndex        =   6
         Top             =   2100
         Width           =   2625
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlToDate 
         Height          =   405
         Left            =   6180
         TabIndex        =   27
         Top             =   1200
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   4485
         Left            =   120
         TabIndex        =   31
         Top             =   3120
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   7911
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
         Column(1)       =   "frmBillingDoc1.frx":27A2
         Column(2)       =   "frmBillingDoc1.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmBillingDoc1.frx":290E
         FormatStyle(2)  =   "frmBillingDoc1.frx":2A6A
         FormatStyle(3)  =   "frmBillingDoc1.frx":2B1A
         FormatStyle(4)  =   "frmBillingDoc1.frx":2BCE
         FormatStyle(5)  =   "frmBillingDoc1.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmBillingDoc1.frx":2D5E
      End
      Begin VB.Label lblOtherFilter 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4920
         TabIndex        =   33
         Top             =   1740
         Width           =   1185
      End
      Begin Threed.SSCommand cmdOther 
         Height          =   525
         Left            =   6800
         TabIndex        =   14
         Top             =   7830
         Visible         =   0   'False
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdjust 
         Height          =   525
         Left            =   5040
         TabIndex        =   30
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCheck chkPayFlag 
         Height          =   435
         Left            =   8910
         TabIndex        =   29
         Top             =   2490
         Visible         =   0   'False
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4920
         TabIndex        =   28
         Top             =   1230
         Width           =   1185
      End
      Begin VB.Label lblFeatureCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   4920
         TabIndex        =   26
         Top             =   2190
         Width           =   1155
      End
      Begin VB.Label lblPartNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   25
         Top             =   2160
         Width           =   1755
      End
      Begin Threed.SSCheck chkCommit 
         Height          =   435
         Left            =   8910
         TabIndex        =   4
         Top             =   2040
         Visible         =   0   'False
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblAccountNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   600
         TabIndex        =   24
         Top             =   1710
         Width           =   1155
      End
      Begin VB.Label lblDocumentNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   23
         Top             =   780
         Width           =   1755
      End
      Begin VB.Label lblCustomerCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   22
         Top             =   1230
         Width           =   1755
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4980
         TabIndex        =   21
         Top             =   2610
         Width           =   1095
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4920
         TabIndex        =   20
         Top             =   780
         Width           =   1185
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   19
         Top             =   2610
         Width           =   1755
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   10110
         TabIndex        =   9
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBillingDoc1.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   10080
         TabIndex        =   10
         Top             =   2160
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   13
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBillingDoc1.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   11
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBillingDoc1.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   12
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
         TabIndex        =   16
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
         TabIndex        =   15
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBillingDoc1.frx":3884
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmBillingDoc1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_BillingDoc As CBillingDoc
Private m_TempBillingDoc As CBillingDoc
Private m_Rs As ADODB.Recordset
Private m_TableName As String
Private m_IvdDocType As Long

Public OKClick As Boolean
Public DocumentType As Long
Public ReceiptType As Long
Public Area As Long
Public DoReceiptFlag As String
Private TempUserName As String
Private m_BillTrans As Collection

Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu
Dim TempStr As String
Dim Programowner As String

   Programowner = glbParameterObj.Programowner
        
   If Area = 1 Then
      TempStr = "(ขาย)"
   ElseIf Area = 2 Then
      TempStr = "(ซื้อ)"
   End If
   
   If DoReceiptFlag = "Y" Then
      Set oMenu = New cPopupMenu
'      lMenuChosen = oMenu.Popup("ขายเชื่อ", "-", "ขายสด", "-", "ทดสอบ")
      lMenuChosen = oMenu.Popup("ขายเชื่อ", "-", "ขายสด")
      If lMenuChosen = 0 Then
         Exit Sub
      End If
      
      If lMenuChosen = 1 Then
         DocumentType = 1
      ElseIf lMenuChosen = 3 Then
         DocumentType = 2
         ReceiptType = 1
      ElseIf lMenuChosen = 5 Then
         DocumentType = 1
      End If
   End If
   
   If Area = 1 Then
      If Not VerifyAccessRight("LEDGER_SELL" & "_" & DocumentType & "_" & "ADD", "เพิ่ม") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf DocumentType = 1000 Or DocumentType = 1001 Or DocumentType = 1002 Or DocumentType = 1003 Or DocumentType = 100 Or DocumentType = 101 Or DocumentType = 102 Or DocumentType = 103 Then
      If Not VerifyAccessRight("LEDGER_STOCKBUY" & "_" & DocumentType & "_" & "ADD", "เพิ่ม") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf Area = 2 Then
      If Not VerifyAccessRight("LEDGER_BUY" & "_" & DocumentType & "_" & "ADD", "เพิ่ม") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If (DocumentType = 1) Or (DocumentType = 7) Then
       If lMenuChosen = 5 Then
         frmAddEditDO2.Area = Area
         frmAddEditDO2.DocumentType = DocumentType
         frmAddEditDO2.HeaderText = MapText("เพิ่มข้อมูลใบส่งสินค้า " & TempStr)
         frmAddEditDO2.ShowMode = SHOW_ADD
         Load frmAddEditDO2
         frmAddEditDO2.Show 1
         
         OKClick = frmAddEditDO2.OKClick
         
         Unload frmAddEditDO2
         Set frmAddEditDO2 = Nothing
    Else
       
         frmAddEditDO2.Area = Area
         frmAddEditDO2.DocumentType = DocumentType
         frmAddEditDO2.HeaderText = MapText("เพิ่มข้อมูลใบส่งสินค้า " & TempStr)
         frmAddEditDO2.ShowMode = SHOW_ADD
         Load frmAddEditDO2
         frmAddEditDO2.Show 1
         
         OKClick = frmAddEditDO2.OKClick
         
         Unload frmAddEditDO2
         Set frmAddEditDO2 = Nothing
         
''      frmAddEditDO.Area = Area
''      frmAddEditDO.DocumentType = DocumentType
''      frmAddEditDO.HeaderText = MapText("เพิ่มข้อมูลใบส่งสินค้า " & TempStr)
''      frmAddEditDO.ShowMode = SHOW_ADD
''      Load frmAddEditDO
''      frmAddEditDO.Show 1
''
''      OKClick = frmAddEditDO.OKClick
''
''      Unload frmAddEditDO
''      Set frmAddEditDO = Nothing
      End If
   ElseIf (DocumentType = 2) Or (DocumentType = 8) Then
      'Set oMenu = New cPopupMenu
      lMenuChosen = ReceiptType 'oMenu.AddMenu(glbGuiConfigs.ReceiptAddMenuItems)
   
'      If lMenuChosen = 0 Then
'         Exit Sub
'      End If
      
      If lMenuChosen = 3 Then
         frmAddEditReceipt.Area = Area
         frmAddEditReceipt.DocumentType = DocumentType
         frmAddEditReceipt.ReceiptType = lMenuChosen
         frmAddEditReceipt.HeaderText = MapText("เพิ่มข้อมูลใบเสร็จ " & TempStr)
         frmAddEditReceipt.ShowMode = SHOW_ADD
         Load frmAddEditReceipt
         frmAddEditReceipt.Show 1
         
         OKClick = frmAddEditReceipt.OKClick
         
         Unload frmAddEditReceipt
         Set frmAddEditReceipt = Nothing
      ElseIf lMenuChosen = 1 Then
         frmAddEditDO.Area = Area
         frmAddEditDO.DocumentType = DocumentType
         frmAddEditDO.ReceiptType = lMenuChosen
         frmAddEditDO.HeaderText = MapText("เพิ่มข้อมูลใบเสร็จ " & TempStr)
         frmAddEditDO.ShowMode = SHOW_ADD
         Load frmAddEditDO
         frmAddEditDO.Show 1
         
         OKClick = frmAddEditDO.OKClick
         
         Unload frmAddEditDO
         Set frmAddEditDO = Nothing
      End If
   ElseIf (DocumentType = 3) Or (DocumentType = 9) Then
      frmAddEditDebitCreditNote.Area = Area
      frmAddEditDebitCreditNote.DebitCreditType = 2
      frmAddEditDebitCreditNote.HeaderText = MapText("เพิ่มข้อมูลใบลดหนี้ " & TempStr)
      frmAddEditDebitCreditNote.ShowMode = SHOW_ADD
      Load frmAddEditDebitCreditNote
      frmAddEditDebitCreditNote.Show 1
      
      OKClick = frmAddEditDebitCreditNote.OKClick
      
      Unload frmAddEditDebitCreditNote
      Set frmAddEditDebitCreditNote = Nothing
   ElseIf (DocumentType = 4) Or (DocumentType = 10) Then
      frmAddEditDebitCreditNote.Area = Area
      frmAddEditDebitCreditNote.DebitCreditType = 1
      frmAddEditDebitCreditNote.HeaderText = MapText("เพิ่มข้อมูลใบเพิ่มหนี้ " & TempStr)
      frmAddEditDebitCreditNote.ShowMode = SHOW_ADD
      Load frmAddEditDebitCreditNote
      frmAddEditDebitCreditNote.Show 1
      
      OKClick = frmAddEditDebitCreditNote.OKClick
      
      Unload frmAddEditDebitCreditNote
      Set frmAddEditDebitCreditNote = Nothing
   ElseIf (DocumentType = 5) Or (DocumentType = 11) Then
      frmAddEditInvoice.DocumentType = DocumentType
      frmAddEditInvoice.Area = Area
      frmAddEditInvoice.HeaderText = MapText("เพิ่มข้อมูลใบกำกับภาษี " & TempStr)
      frmAddEditInvoice.ShowMode = SHOW_ADD
      Load frmAddEditInvoice
      frmAddEditInvoice.Show 1
      
      OKClick = frmAddEditInvoice.OKClick
      
      Unload frmAddEditInvoice
      Set frmAddEditInvoice = Nothing
   ElseIf DocumentType = 6 Then
      frmAddEditBillSummary.Area = Area
      frmAddEditBillSummary.HeaderText = MapText("เพิ่มข้อมูลใบสรุปวางบิล " & TempStr)
      frmAddEditBillSummary.ShowMode = SHOW_ADD
      Load frmAddEditBillSummary
      frmAddEditBillSummary.Show 1

      OKClick = frmAddEditBillSummary.OKClick

      Unload frmAddEditBillSummary
      Set frmAddEditBillSummary = Nothing
   ElseIf DocumentType = 12 Then
      frmAddEditPO.Area = Area
      frmAddEditPO.DocumentType = DocumentType
      frmAddEditPO.HeaderText = MapText("เพิ่มข้อมูลใบรับงาน/สั่งงาน " & TempStr)
      frmAddEditPO.ShowMode = SHOW_ADD
      Load frmAddEditPO
      frmAddEditPO.Show 1

      OKClick = frmAddEditPO.OKClick

      Unload frmAddEditPO
      Set frmAddEditPO = Nothing
   ElseIf (DocumentType = 14) Or (DocumentType = 13) Then
      frmAddEditQuoatation.Area = Area
      frmAddEditQuoatation.DocumentType = DocumentType
      frmAddEditQuoatation.HeaderText = MapText("เพิ่มข้อมูลใบเสนอราคา" & TempStr)
      frmAddEditQuoatation.ShowMode = SHOW_ADD
      Load frmAddEditQuoatation
      frmAddEditQuoatation.Show 1

      OKClick = frmAddEditQuoatation.OKClick

      Unload frmAddEditQuoatation
      Set frmAddEditQuoatation = Nothing
   
   ElseIf (DocumentType = 15) Or (DocumentType = 17) Then
      frmAddEditPackingList.Area = Area
      frmAddEditPackingList.DocumentType = DocumentType
      frmAddEditPackingList.HeaderText = MapText("เพิ่มข้อมูลใบบรรจุหีบห่อ" & TempStr)
      frmAddEditPackingList.ShowMode = SHOW_ADD
      Load frmAddEditPackingList
      frmAddEditPackingList.Show 1

      OKClick = frmAddEditPackingList.OKClick

      Unload frmAddEditPackingList
      Set frmAddEditPackingList = Nothing
   
   ElseIf (DocumentType = 16) Then
      frmAddEditMemo.Area = Area
      frmAddEditMemo.DocumentType = DocumentType
      frmAddEditMemo.HeaderText = MapText("เพิ่มข้อมูลใบ MEMO ธนาคาร" & TempStr)
      frmAddEditMemo.ShowMode = SHOW_ADD
      Load frmAddEditMemo
      frmAddEditMemo.Show 1

      OKClick = frmAddEditMemo.OKClick

      Unload frmAddEditMemo
      Set frmAddEditMemo = Nothing
   ElseIf (DocumentType = 18) Then
      frmAddEditReturn.Area = Area
      frmAddEditReturn.DocumentType = DocumentType
      frmAddEditReturn.HeaderText = MapText("เพิ่มข้อมูลใบรับคืนสินค้า " & TempStr)
      
      frmAddEditReturn.ShowMode = SHOW_ADD
      Load frmAddEditReturn
      frmAddEditReturn.Show 1
      
      OKClick = frmAddEditReturn.OKClick
      
      Unload frmAddEditReturn
      Set frmAddEditReturn = Nothing
   ElseIf DocumentType = 19 Then
      Set oMenu = New cPopupMenu
'      lMenuChosen = oMenu.Popup("Sale order", "-", "ทดสอบ")
      lMenuChosen = oMenu.Popup("Sale order")
      If lMenuChosen = 0 Then
         Exit Sub
      End If
      
      If lMenuChosen = 1 Then
'''         frmAddEditSaleOrder.Area = Area
'''         frmAddEditSaleOrder.DocumentType = DocumentType
'''         frmAddEditSaleOrder.HeaderText = MapText("เพิ่มข้อมูลใบ Sale Order " & TempStr)
'''         frmAddEditSaleOrder.ShowMode = SHOW_ADD
'''         Load frmAddEditSaleOrder
'''         frmAddEditSaleOrder.Show 1
'''
'''         OKClick = frmAddEditSaleOrder.OKClick
'''
'''         Unload frmAddEditSaleOrder
'''         Set frmAddEditSaleOrder = Nothing

         frmAddEditSaleOrder2.Area = Area
         frmAddEditSaleOrder2.DocumentType = DocumentType
         frmAddEditSaleOrder2.HeaderText = MapText("เพิ่มข้อมูลใบ Sale Order " & TempStr)
         frmAddEditSaleOrder2.ShowMode = SHOW_ADD
         Load frmAddEditSaleOrder2
         frmAddEditSaleOrder2.Show 1
   
         OKClick = frmAddEditSaleOrder2.OKClick
   
         Unload frmAddEditSaleOrder2
         Set frmAddEditSaleOrder2 = Nothing
      ElseIf lMenuChosen = 3 Then
         frmAddEditSaleOrder2.Area = Area
         frmAddEditSaleOrder2.DocumentType = DocumentType
         frmAddEditSaleOrder2.HeaderText = MapText("เพิ่มข้อมูลใบ Sale Order " & TempStr)
         frmAddEditSaleOrder2.ShowMode = SHOW_ADD
         Load frmAddEditSaleOrder2
         frmAddEditSaleOrder2.Show 1
   
         OKClick = frmAddEditSaleOrder2.OKClick
   
         Unload frmAddEditSaleOrder2
         Set frmAddEditSaleOrder2 = Nothing
      End If
      
'      frmAddEditSaleOrder.Area = Area
'      frmAddEditSaleOrder.DocumentType = DocumentType
'      frmAddEditSaleOrder.HeaderText = MapText("เพิ่มข้อมูลใบ Sale Order " & TempStr)
'      frmAddEditSaleOrder.ShowMode = SHOW_ADD
'      Load frmAddEditSaleOrder
'      frmAddEditSaleOrder.Show 1
'
'      OKClick = frmAddEditSaleOrder.OKClick
'
'      Unload frmAddEditSaleOrder
'      Set frmAddEditSaleOrder = Nothing
      
   ElseIf DocumentType = 100 Or DocumentType = 101 Or DocumentType = 102 Or DocumentType = 103 Then
      frmAddEditBillingSup.DocumentType = DocumentType
      Select Case DocumentType
      Case 100
           frmAddEditBillingSup.HeaderText = MapText("เพิ่มข้อมูลการรับเข้าวัตถุดิบ")
           
      Case 101
            frmAddEditBillingSup.HeaderText = MapText("เพิ่มข้อมูลการรับเข้าวัสดุอุปกรณ์")
     Case 102
            frmAddEditBillingSup.HeaderText = MapText("เพิ่มข้อมูลการรับเข้าจ่ายออกวัสดุอุปกรณ์")
      Case 103
           frmAddEditBillingSup.HeaderText = MapText("เพิ่มข้อมูลการรับเข้าทั่วไป")
      End Select
      frmAddEditBillingSup.ShowMode = SHOW_ADD
      Load frmAddEditBillingSup
      frmAddEditBillingSup.Show 1
      
      OKClick = frmAddEditBillingSup.OKClick
      
      Unload frmAddEditBillingSup
      Set frmAddEditBillingSup = Nothing
   ElseIf DocumentType = 110 Then
      frmAddEditBillingSupRe.DocumentType = DocumentType
      frmAddEditBillingSupRe.HeaderText = MapText("เพิ่มข้อมูลรับคืน (ซื้อ)")
      frmAddEditBillingSupRe.ShowMode = SHOW_ADD
      Load frmAddEditBillingSupRe
      frmAddEditBillingSupRe.Show 1
      
      OKClick = frmAddEditBillingSupRe.OKClick
      
      Unload frmAddEditBillingSupRe
      Set frmAddEditBillingSupRe = Nothing
   ElseIf DocumentType = 111 Then
      frmAddEditBillingSupRe.DocumentType = DocumentType
      frmAddEditBillingSupRe.HeaderText = MapText("เพิ่มข้อมูลรับคืน (ซื้อ)")
      frmAddEditBillingSupRe.ShowMode = SHOW_ADD
      Load frmAddEditBillingSupRe
      frmAddEditBillingSupRe.Show 1
      
      OKClick = frmAddEditBillingSupRe.OKClick
      
      Unload frmAddEditBillingSupRe
      Set frmAddEditBillingSupRe = Nothing
   ElseIf DocumentType = 1000 Or DocumentType = 1001 Or DocumentType = 1002 Or DocumentType = 1003 Then
      frmAddEditBillingSup.DocumentType = DocumentType
      Select Case DocumentType
      Case 1000
           frmAddEditBillingSup.HeaderText = MapText("เพิ่มข้อมูลใบ PO สั่งซื้อวัตถุดิบ")
      Case 1001
            frmAddEditBillingSup.HeaderText = MapText("เพิ่มข้อมูลใบ PO สั่งซื้อวัสดุอุปกรณ์")
     Case 1002
            frmAddEditBillingSup.HeaderText = MapText("เพิ่มข้อมูลใบ PO สั่งซื้อ รับเข้าจ่ายออกวัสดุอุปกรณ์")
      Case 1003
           frmAddEditBillingSup.HeaderText = MapText("เพิ่มข้อมูลใบ PO สั่งซื้อทั่วไป")
      End Select
     
      frmAddEditBillingSup.ShowMode = SHOW_ADD
      Load frmAddEditBillingSup
      frmAddEditBillingSup.Show 1
      
      OKClick = frmAddEditBillingSup.OKClick
      
      Unload frmAddEditBillingSup
      Set frmAddEditBillingSup = Nothing
   End If
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub

'Private Sub cmdAdjust_Click()
'Dim IsOK As Boolean
'Dim itemcount As Long
'Dim IsCanLock As Boolean
'Dim ID As Long
'Dim PaymentID As Long
'
'   If Not cmdAdjust.Enabled Then
'      Exit Sub
'   End If
'
'   If Not VerifyGrid(GridEX1.Value(1)) Then
'      Exit Sub
'   End If
'   ID = GridEX1.Value(1)
'
'   Call EnableForm(Me, False)
'
'   Dim Bd As CBillingDoc
'   Dim X As Double
'
'   Set Bd = New CBillingDoc
'   Bd.BILLING_DOC_ID = ID
'   Call Bd.UpdatePaidAmount
'   Call Bd.UpdateCnDnAmount
'   Set Bd = Nothing
'
'   Call EnableForm(Me, True)
'
'End Sub

Private Sub cmdClear_Click()
   txtDocumentNo.Text = ""
   txtCustomerCode.Text = ""
   txtAccountNo.Text = ""
   txtFeatureCode.Text = ""
   txtPartNo.Text = ""
   
   uctlDocumentDate.ShowDate = -1
   uctlToDate.ShowDate = -1
   
   cboOtherFilter.ListIndex = -1
   
   cboOrderBy.ListIndex = -1
   cboOrderType.ListIndex = -1
End Sub

Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim PaymentID As Long
   If Area = 1 Then
      If Not VerifyAccessRight("LEDGER_SELL" & "_" & DocumentType & "_" & "DELETE", "ลบ") Then
         frmVerifyAccRight.AccName = "LEDGER_SELL" & "_" & DocumentType & "_" & "DELETE"
         frmVerifyAccRight.AccDesc = "ลบ"
         Load frmVerifyAccRight
         frmVerifyAccRight.Show 1
         
         If frmVerifyAccRight.GrantRight Then
            Unload frmVerifyAccRight
            Set frmVerifyAccRight = Nothing
         Else
            Unload frmVerifyAccRight
            Set frmVerifyAccRight = Nothing
            Call EnableForm(Me, True)
            Exit Sub
         End If
      End If
   ElseIf Area = 2 Then
      If DocumentType = 1000 Or DocumentType = 1001 Or DocumentType = 1002 Or DocumentType = 1003 Or DocumentType = 100 Or DocumentType = 101 Or DocumentType = 102 Or DocumentType = 103 Then
         If Not VerifyAccessRight("LEDGER_STOCKBUY" & "_" & DocumentType & "_" & "DELETE", "ลบ") Then
            frmVerifyAccRight.AccName = "LEDGER_STOCKBUY" & "_" & DocumentType & "_" & "DELETE"
            frmVerifyAccRight.AccDesc = "ลบ"
            Load frmVerifyAccRight
            frmVerifyAccRight.Show 1

            If frmVerifyAccRight.GrantRight Then
               Unload frmVerifyAccRight
               Set frmVerifyAccRight = Nothing
            Else
               Unload frmVerifyAccRight
               Set frmVerifyAccRight = Nothing
               Call EnableForm(Me, True)
               Exit Sub
            End If
         End If
      Else
         If Not VerifyAccessRight("LEDGER_BUY" & "_" & DocumentType & "_" & "DELETE", "ลบ") Then
            frmVerifyAccRight.AccName = "LEDGER_BUY" & "_" & DocumentType & "_" & "DELETE"
            frmVerifyAccRight.AccDesc = "ลบ"
            Load frmVerifyAccRight
            frmVerifyAccRight.Show 1
            
            If frmVerifyAccRight.GrantRight Then
               Unload frmVerifyAccRight
               Set frmVerifyAccRight = Nothing
            Else
               Unload frmVerifyAccRight
               Set frmVerifyAccRight = Nothing
               Call EnableForm(Me, True)
               Exit Sub
            End If
         End If
      End If
   End If
   
   If Not cmdDelete.Enabled Then
      Exit Sub
   End If
        
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   ID = GridEX1.Value(1)
   PaymentID = GridEX1.Value(8)
   
   If m_TempBillingDoc.PO_APPROVED_FLAG = "Y" Then
     MsgBox "ไม่สามารถลบได้เนื่องจากมีการอนุมัติแล้ว"
     Exit Sub
   End If
   
    'ตรวจสอบการแก้ไขเอกสารที่ออกบิลแล้ว
   If m_TempBillingDoc.POST_FLAG = "Y" Then
         glbErrorLog.LocalErrorMsg = MapText("เอกสารใบได้ทำการออกบิลขายเรียบร้อยแล้ว หากต้องการเปลี่ยนแปลงเอกสารต้องให้ผู้ควบคุม อนุมัติก่อน")
         glbErrorLog.ShowUserError
      
         frmVerifyAccRight.AccName = "LEDGER_SELL_" & DocumentType & "_" & "MANAGE"
         frmVerifyAccRight.AccDesc = "สามารถอนุมัติการเปลี่ยนแปลงเอกสารได้"
         Load frmVerifyAccRight
         frmVerifyAccRight.Show 1
         If frmVerifyAccRight.GrantRight Then
            TempUserName = frmVerifyAccRight.UserName
            Unload frmVerifyAccRight
            Set frmVerifyAccRight = Nothing
            m_BillingDoc.APPROVE_MANAGE_NAME = TempUserName
         Else
            Unload frmVerifyAccRight
            Set frmVerifyAccRight = Nothing
            Exit Sub
         End If
   End If
   
   
   Call glbDatabaseMngr.LockTable(m_TableName, ID, IsCanLock, glbErrorLog)
   If Not ConfirmDelete(GridEX1.Value(2)) Then
      Call glbDatabaseMngr.UnLockTable(m_TableName, ID, IsCanLock, glbErrorLog)
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   
      'ดึงข้อมูลก่อนลบ
      Dim m_BillingDoc2 As CBillingDoc
      Set m_BillingDoc2 = New CBillingDoc
      m_BillingDoc2.BILLING_DOC_ID = ID
      m_BillingDoc2.QueryFlag = 1

       If Not glbDaily.QueryBillingDoc(m_BillingDoc2, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   
      '***********************
      
   If Not glbDaily.DeleteBillingDoc(ID, IsOK, True, glbErrorLog, PaymentID) Then
      m_BillingDoc.BILLING_DOC_ID = -1
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call glbDatabaseMngr.UnLockTable(m_TableName, ID, IsCanLock, glbErrorLog)
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   If ItemCount > 0 And IsOK Then
      Call m_BillingDoc2.PopulateFromRS(1, m_Rs)
      Call EditStatusFlagInInventoryWHDoc(m_BillingDoc2, "D")
      Call EditStatusFlagInBillingDoc(m_BillingDoc2) 'ไป uddate ในใบ SO ด้วย
   End If

   
   Call QueryData(True)
   
   Call glbDatabaseMngr.UnLockTable(m_TableName, ID, IsCanLock, glbErrorLog)
   Call EnableForm(Me, True)
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean
Dim TempStr As String

   Dim Programowner As String
   Programowner = glbParameterObj.Programowner
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(1))
   Call glbDatabaseMngr.LockTable(m_TableName, ID, IsCanLock, glbErrorLog)
   
   If Area = 1 Then
      TempStr = "(ขาย)"
   ElseIf Area = 2 Then
      TempStr = "(ซื้อ)"
   End If
   
   If DoReceiptFlag = "Y" Then
      DocumentType = Val(GridEX1.Value(9))
      ReceiptType = Val(GridEX1.Value(7))
   End If
   
   If (DocumentType = 1) Or (DocumentType = 7) Then
       frmAddEditDO2.Area = Area
      frmAddEditDO2.DocumentType = DocumentType
      frmAddEditDO2.ID = Val(GridEX1.Value(1)) 'ID
      frmAddEditDO2.HeaderText = MapText("แก้ไขข้อมูลใบส่งสินค้า ")
      frmAddEditDO2.ShowMode = SHOW_EDIT
      Load frmAddEditDO2
      frmAddEditDO2.Show 1
      
      OKClick = frmAddEditDO2.OKClick
      
      Unload frmAddEditDO2
      Set frmAddEditDO2 = Nothing
      
'      frmAddEditDO.Area = Area
'      frmAddEditDO.DocumentType = DocumentType
'      frmAddEditDO.id = id
'      frmAddEditDO.HeaderText = MapText("แก้ไขข้อมูลใบส่งสินค้า " & TempStr)
'      frmAddEditDO.ShowMode = SHOW_EDIT
'      Load frmAddEditDO
'      frmAddEditDO.Show 1
'
'      OKClick = frmAddEditDO.OKClick
'
'      Unload frmAddEditDO
'      Set frmAddEditDO = Nothing
   ElseIf (DocumentType = 2) Or (DocumentType = 8) Then
      If Val(GridEX1.Value(7)) = 1 Then
         frmAddEditDO.Area = Area
         frmAddEditDO.DocumentType = DocumentType
         frmAddEditDO.ReceiptType = Val(GridEX1.Value(7))
         frmAddEditDO.ID = ID
         frmAddEditDO.HeaderText = MapText("แก้ไขข้อมูลใบเสร็จ " & TempStr)
         frmAddEditDO.ShowMode = SHOW_EDIT
         Load frmAddEditDO
         frmAddEditDO.Show 1
         
         OKClick = frmAddEditDO.OKClick
         
         Unload frmAddEditDO
         Set frmAddEditDO = Nothing
      ElseIf Val(GridEX1.Value(7)) = 3 Then
         frmAddEditReceipt.Area = Area
         frmAddEditReceipt.DocumentType = DocumentType
         frmAddEditReceipt.ReceiptType = Val(GridEX1.Value(7))
         frmAddEditReceipt.ID = ID
         frmAddEditReceipt.HeaderText = MapText("แก้ไขข้อมูลใบเสร็จ " & TempStr)
         frmAddEditReceipt.ShowMode = SHOW_EDIT
         Load frmAddEditReceipt
         frmAddEditReceipt.Show 1
         
         OKClick = frmAddEditReceipt.OKClick
         
         Unload frmAddEditReceipt
         Set frmAddEditReceipt = Nothing
      End If
   ElseIf (DocumentType = 3) Or (DocumentType = 9) Then
      frmAddEditDebitCreditNote.Area = Area
      frmAddEditDebitCreditNote.DebitCreditType = 2
      frmAddEditDebitCreditNote.ID = ID
      frmAddEditDebitCreditNote.HeaderText = MapText("แก้ไขข้อมูลใบลดหนี้ " & TempStr)
      frmAddEditDebitCreditNote.ShowMode = SHOW_EDIT
      Load frmAddEditDebitCreditNote
      frmAddEditDebitCreditNote.Show 1
      
      OKClick = frmAddEditDebitCreditNote.OKClick
      
      Unload frmAddEditDebitCreditNote
      Set frmAddEditDebitCreditNote = Nothing
   ElseIf (DocumentType = 4) Or (DocumentType = 10) Then
      frmAddEditDebitCreditNote.Area = Area
      frmAddEditDebitCreditNote.DebitCreditType = 1
      frmAddEditDebitCreditNote.ID = ID
      frmAddEditDebitCreditNote.HeaderText = MapText("แก้ไขข้อมูลใบเพิ่มหนี้ " & TempStr)
      frmAddEditDebitCreditNote.ShowMode = SHOW_EDIT
      Load frmAddEditDebitCreditNote
      frmAddEditDebitCreditNote.Show 1
      
      OKClick = frmAddEditDebitCreditNote.OKClick
      
      Unload frmAddEditDebitCreditNote
      Set frmAddEditDebitCreditNote = Nothing
   ElseIf (DocumentType = 5) Or (DocumentType = 11) Then
      frmAddEditInvoice.DocumentType = DocumentType
      frmAddEditInvoice.Area = Area
      frmAddEditInvoice.ID = ID
      frmAddEditInvoice.HeaderText = MapText("แก้ไขข้อมูลใบกำกับภาษี " & TempStr)
      frmAddEditInvoice.ShowMode = SHOW_EDIT
      Load frmAddEditInvoice
      frmAddEditInvoice.Show 1
      
      OKClick = frmAddEditInvoice.OKClick
      
      Unload frmAddEditInvoice
      Set frmAddEditInvoice = Nothing
   ElseIf DocumentType = 6 Then
      frmAddEditBillSummary.Area = Area
      frmAddEditBillSummary.ID = ID
      frmAddEditBillSummary.HeaderText = MapText("แก้ไขข้อมูลใบสรุปวางบิล " & TempStr)
      frmAddEditBillSummary.ShowMode = SHOW_EDIT
      Load frmAddEditBillSummary
      frmAddEditBillSummary.Show 1

      OKClick = frmAddEditBillSummary.OKClick

      Unload frmAddEditBillSummary
      Set frmAddEditBillSummary = Nothing
   ElseIf DocumentType = 12 Then
      frmAddEditPO.DocumentType = DocumentType
      frmAddEditPO.Area = Area
      frmAddEditPO.ID = ID
      frmAddEditPO.HeaderText = MapText("แก้ไขข้อมูลใบรับงาน/สั่งงาน " & TempStr)
      frmAddEditPO.ShowMode = SHOW_EDIT
      Load frmAddEditPO
      frmAddEditPO.Show 1

      OKClick = frmAddEditPO.OKClick

      Unload frmAddEditPO
      Set frmAddEditPO = Nothing
   ElseIf (DocumentType = 13) Or (DocumentType = 14) Then
      frmAddEditQuoatation.DocumentType = DocumentType
      frmAddEditQuoatation.Area = Area
      frmAddEditQuoatation.ID = ID
      frmAddEditQuoatation.HeaderText = MapText("แก้ไขข้อมูลใบเสนอราคา " & TempStr)
      frmAddEditQuoatation.ShowMode = SHOW_EDIT
      Load frmAddEditQuoatation
      frmAddEditQuoatation.Show 1

      OKClick = frmAddEditQuoatation.OKClick

      Unload frmAddEditQuoatation
      Set frmAddEditQuoatation = Nothing
     ElseIf (DocumentType = 15) Or (DocumentType = 17) Then
      frmAddEditPackingList.DocumentType = DocumentType
      frmAddEditPackingList.Area = Area
      frmAddEditPackingList.ID = ID
      frmAddEditPackingList.HeaderText = MapText("แก้ไขข้อมูลใบบรรจุหีบห่อ " & TempStr)
      frmAddEditPackingList.ShowMode = SHOW_EDIT
      Load frmAddEditPackingList
      frmAddEditPackingList.Show 1

      OKClick = frmAddEditPackingList.OKClick

      Unload frmAddEditPackingList
      Set frmAddEditPackingList = Nothing
  
  ElseIf (DocumentType = 16) Then
      frmAddEditMemo.DocumentType = DocumentType
      frmAddEditMemo.Area = Area
      frmAddEditMemo.ID = ID
      frmAddEditMemo.HeaderText = MapText("แก้ไขข้อมูลใบ MEMO ธนาคาร " & TempStr)
      frmAddEditMemo.ShowMode = SHOW_EDIT
      Load frmAddEditMemo
      frmAddEditMemo.Show 1

      OKClick = frmAddEditMemo.OKClick

      Unload frmAddEditMemo
      Set frmAddEditMemo = Nothing
   ElseIf (DocumentType = 18) Then
      frmAddEditReturn.DocumentType = DocumentType
      frmAddEditReturn.Area = Area
      frmAddEditReturn.ID = ID
      frmAddEditReturn.HeaderText = MapText("แก้ไขข้อมูลใบรับคืนสินค้า " & TempStr)
      frmAddEditReturn.ShowMode = SHOW_EDIT
      Load frmAddEditReturn
      frmAddEditReturn.Show 1

      OKClick = frmAddEditReturn.OKClick

      Unload frmAddEditReturn
      Set frmAddEditReturn = Nothing
   ElseIf DocumentType = 19 Then

'   Dim TempStr As String
'     If Area = 1 Then
'      TempStr = "(ขาย)"
'   ElseIf Area = 2 Then
'      TempStr = "(ซื้อ)"
'   End If
   
       frmAddEditSaleOrder2.Area = Area
      frmAddEditSaleOrder2.DocumentType = DocumentType
      frmAddEditSaleOrder2.ID = Val(GridEX1.Value(1)) 'ID
       frmAddEditSaleOrder2.HeaderText = MapText("แก้ไขข้อมูลใบ SO " & TempStr)
      frmAddEditSaleOrder2.ShowMode = SHOW_EDIT
      Load frmAddEditSaleOrder2
      frmAddEditSaleOrder2.Show 1
      
      OKClick = frmAddEditSaleOrder2.OKClick
      
      Unload frmAddEditSaleOrder2
      Set frmAddEditSaleOrder2 = Nothing
      
'      If OKClick Then
'         Call QueryData(True)
'      End If
      
'''''      frmAddEditSaleOrder.DocumentType = DocumentType
'''''      frmAddEditSaleOrder.Area = Area
'''''      frmAddEditSaleOrder.id = id
'''''      frmAddEditSaleOrder.HeaderText = MapText("แก้ไขข้อมูลใบ SO " & TempStr)
'''''      frmAddEditSaleOrder.ShowMode = SHOW_EDIT
'''''      Load frmAddEditSaleOrder
'''''      frmAddEditSaleOrder.Show 1
'''''
'''''      OKClick = frmAddEditSaleOrder.OKClick
'''''
'''''      Unload frmAddEditSaleOrder
'''''      Set frmAddEditSaleOrder = Nothing
   ElseIf DocumentType = 100 Or DocumentType = 101 Or DocumentType = 102 Or DocumentType = 103 Then
      frmAddEditBillingSup.DocumentType = DocumentType
      frmAddEditBillingSup.ID = ID
       Select Case DocumentType
      Case 100
      frmAddEditBillingSup.HeaderText = MapText("แก้ไขข้อมูลการรับเข้าวัตุดิบ")
      Case 101
      frmAddEditBillingSup.HeaderText = MapText("แก้ไขข้อมูลการรับเข้าวัสดุอุปกรณ์")
      Case 102
      frmAddEditBillingSup.HeaderText = MapText("แก้ไขข้อมูลการรับเข้าจ่ายออกวัสดุอุปกรณ์")
      Case 103
      frmAddEditBillingSup.HeaderText = MapText("แก้ไขข้อมูลการรับเข้าของใช้ทั่วไป")
      End Select
      frmAddEditBillingSup.ShowMode = SHOW_EDIT
      Load frmAddEditBillingSup
      frmAddEditBillingSup.Show 1
      
      OKClick = frmAddEditBillingSup.OKClick
      
      Unload frmAddEditBillingSup
      Set frmAddEditBillingSup = Nothing
   ElseIf DocumentType = 110 Then
      frmAddEditBillingSupRe.DocumentType = DocumentType
      frmAddEditBillingSupRe.ID = ID
      frmAddEditBillingSupRe.HeaderText = MapText("แก้ไขข้อมูลการรับเข้า (ซื้อ)")
      frmAddEditBillingSupRe.ShowMode = SHOW_EDIT
      Load frmAddEditBillingSupRe
      frmAddEditBillingSupRe.Show 1
      
      OKClick = frmAddEditBillingSupRe.OKClick
      
      Unload frmAddEditBillingSupRe
      Set frmAddEditBillingSupRe = Nothing
   ElseIf DocumentType = 1000 Or DocumentType = 1001 Or DocumentType = 1002 Or DocumentType = 1003 Then
      frmAddEditBillingSup.DocumentType = DocumentType
      frmAddEditBillingSup.ID = ID
      Select Case DocumentType
      Case 1000
      frmAddEditBillingSup.HeaderText = MapText("แก้ไขข้อมูล PO สั่งซื้อวัตุดิบ")
      Case 1001
      frmAddEditBillingSup.HeaderText = MapText("แก้ไขข้อมูล PO สั่งซื้อวัสดุอุปกรณ์")
      Case 1002
      frmAddEditBillingSup.HeaderText = MapText("แก้ไขข้อมูล PO สั่งซื้อ รับเข้าจ่ายออกวัสดุอุปกรณ์")
      Case 1003
      frmAddEditBillingSup.HeaderText = MapText("แก้ไขข้อมูล PO สั่งซื้อของใช้ทั่วไป")
      End Select
      
      frmAddEditBillingSup.ShowMode = SHOW_EDIT
      Load frmAddEditBillingSup
      frmAddEditBillingSup.Show 1
      
      OKClick = frmAddEditBillingSup.OKClick
      
      Unload frmAddEditBillingSup
      Set frmAddEditBillingSup = Nothing
   End If
   
   If OKClick Then
      Call QueryData(True)
   End If
   Call glbDatabaseMngr.UnLockTable(m_TableName, ID, IsCanLock, glbErrorLog)

End Sub

Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub

Private Sub cmdOther_Click()
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu

   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("เปิดใบรับของโดยไม่มี PO", "-", "อื่นๆ")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
   If lMenuChosen = 1 Then
      If DocumentType = 100 Or DocumentType = 101 Or DocumentType = 102 Or DocumentType = 103 Then
         If Not VerifyAccessRight("LEDGER_STOCKBUY" & "_" & DocumentType & "_" & "NO-PO", "ไม่มีPO") Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
         
         frmAddEditBillingSup.AutoGenPo = True
         frmAddEditBillingSup.DocumentType = DocumentType
         frmAddEditBillingSup.HeaderText = MapText("เพิ่มข้อมูลการนำเข้า โดยไม่มี PO")
         frmAddEditBillingSup.ShowMode = SHOW_ADD
         Load frmAddEditBillingSup
         frmAddEditBillingSup.Show 1
         
         OKClick = frmAddEditBillingSup.OKClick
         
         Unload frmAddEditBillingSup
         Set frmAddEditBillingSup = Nothing
         
         
      End If
   ElseIf lMenuChosen = 3 Then
   
   End If
      
   If OKClick Then
      Call QueryData(True)
   End If
End Sub

Private Sub cmdSearch_Click()
   Call QueryData(True)
End Sub

Private Sub Form_Activate()
Dim FromDate As Date
Dim ToDate As Date
   If Not m_HasActivate Then
      m_HasActivate = True
      
      Call InitBillingDocOtherFilterOrderBy(cboOtherFilter)
      
      Call InitBillingDocOrderBy(cboOrderBy)
      Call InitOrderType(cboOrderType)
      
      
      Call GetFirstLastDate(Now, FromDate, ToDate)
      uctlDocumentDate.ShowDate = FromDate
      uctlToDate.ShowDate = ToDate
      
      Call QueryData(True)
   End If
End Sub

Private Function GetPermissionCode() As String
Dim TempStr As String

   If Area = 1 Then
      If DocumentType = 1 Then
         TempStr = "LEDGER_DO"
      ElseIf DocumentType = 2 Then
         TempStr = "LEDGER_RC"
      ElseIf DocumentType = 3 Then
         TempStr = "LEDGER_CN"
      ElseIf DocumentType = 4 Then
         TempStr = "LEDGER_DN"
      ElseIf DocumentType = 18 Then
         TempStr = "LEDGER_RT"
      ElseIf DocumentType = 19 Then
         TempStr = "LEDGER_SO"
      End If
   End If
   
'   If ActionCode = 1 Then 'add
'      TempStr = TempStr & "_ADD"
'   ElseIf ActionCode = 2 Then 'edit
'      TempStr = TempStr & "_EDIT"
'   ElseIf ActionCode = 3 Then 'delete
'      TempStr = TempStr & "_DELETE"
'   ElseIf ActionCode = 4 Then 'query
'      TempStr = TempStr & "_QUERY"
'   Else 'print
'      TempStr = TempStr & "_PRINT"
'   End If
   
   GetPermissionCode = TempStr
End Function

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      Call LoadBillTransportByTypePrice(Nothing, m_BillTrans, uctlDocumentDate.ShowDate, uctlToDate.ShowDate, 3, 1) 'ค่าขน KEY BY BILL
      
      m_BillingDoc.BILLING_DOC_ID = -1
      m_BillingDoc.DOCUMENT_NO = PatchWildCard(txtDocumentNo.Text)
      If Area = 2 Then
         m_BillingDoc.SUPPLIER_CODE = PatchWildCard(txtCustomerCode.Text)
         m_BillingDoc.PART_NO_SUPITEM_SEARCH = txtPartNo.Text
      Else
         m_BillingDoc.CUSTOMER_CODE = PatchWildCard(txtCustomerCode.Text)
         m_BillingDoc.PART_NO = txtPartNo.Text
      End If
      m_BillingDoc.ACCOUNT_NO = PatchWildCard(txtAccountNo.Text)
      m_BillingDoc.FROM_DATE = uctlDocumentDate.ShowDate
      m_BillingDoc.TO_DATE = uctlToDate.ShowDate
      m_BillingDoc.DOCUMENT_TYPE = DocumentType
      m_BillingDoc.RECEIPT_TYPE = ReceiptType
      m_BillingDoc.COMMIT_FLAG = Check2Flag(chkCommit.Value)
      
      m_BillingDoc.FEATURE_CODE = txtFeatureCode.Text
      m_BillingDoc.OrderBy = cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex))
      If m_BillingDoc.OrderBy <= 0 Then
         'm_BillingDoc.OrderBy = 1
      End If
      m_BillingDoc.OrderType = cboOrderType.ItemData(Minus2Zero(cboOrderType.ListIndex))
      m_BillingDoc.DO_RECEIPT_FLAG = DoReceiptFlag
      m_BillingDoc.PAY_FLAG = Check2Flag(chkPayFlag.Value)
      If m_BillingDoc.PAY_FLAG = "N" Then
         m_BillingDoc.PAY_FLAG = ""
      End If
      
      m_BillingDoc.PO_APPROVED_FLAG = ""
      m_BillingDoc.AUTO_GEN_FLAG = ""
      If (cboOtherFilter.ItemData(Minus2Zero(cboOtherFilter.ListIndex))) = 1 Then
         m_BillingDoc.PO_APPROVED_FLAG = "N"
      ElseIf (cboOtherFilter.ItemData(Minus2Zero(cboOtherFilter.ListIndex))) = 2 Then
         m_BillingDoc.PO_APPROVED_FLAG = "Y"
      ElseIf (cboOtherFilter.ItemData(Minus2Zero(cboOtherFilter.ListIndex))) = 3 Then
         m_BillingDoc.AUTO_GEN_FLAG = "Y"
         m_BillingDoc.PO_APPROVED_FLAG = "N"
      ElseIf (cboOtherFilter.ItemData(Minus2Zero(cboOtherFilter.ListIndex))) = 4 Then
         m_BillingDoc.AUTO_GEN_FLAG = "Y"
         m_BillingDoc.PO_APPROVED_FLAG = "Y"
      End If
      
      If DoReceiptFlag = "Y" Then
         m_BillingDoc.DOCUMENT_TYPE = -1
         m_BillingDoc.RECEIPT_TYPE = -1
      End If
      If Not glbDaily.QueryBillingDoc(m_BillingDoc, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      cmdDelete.Enabled = (m_BillingDoc.COMMIT_FLAG = "N")
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
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 123 Then
      Call AddMemoNote
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
   Col.Width = 2000 '2115
   Col.Caption = MapText("เลขที่เอกสาร")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 1500 '2055
   Col.Caption = MapText("วันที่เอกสาร")

   If Area = 1 Then
      Set Col = GridEX1.Columns.add '4
      Col.Width = 1600
      Col.Caption = MapText("รหัสลูกค้า")
      
      Set Col = GridEX1.Columns.add '5
      Col.Width = 4995
      Col.Caption = MapText("ชื่อลูกค้า")
   ElseIf Area = 2 Then
      Set Col = GridEX1.Columns.add '4
      Col.Width = 1700 '2305
      Col.Caption = MapText("รหัสซัพพลายเออร์")
      
      Set Col = GridEX1.Columns.add '5
      Col.Width = 4995
      Col.Caption = MapText("ชื่อซัพพลายเออร์")
   End If
   Set Col = GridEX1.Columns.add '6
   Col.Width = 0
   Col.Visible = False
   Col.Caption = MapText("COMMIT FLAG")
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = 0
   Col.Visible = False
   Col.Caption = MapText("RECEIPT_TYPE")
   
   Set Col = GridEX1.Columns.add '8
   Col.Width = 0
   Col.Visible = False
   Col.Caption = MapText("PAYMENT_ID")
   
   Set Col = GridEX1.Columns.add '9
   Col.Width = 0
   Col.Visible = False
   Col.Caption = MapText("DOCUMENT_TYPE")
   
   Set Col = GridEX1.Columns.add '10
   Col.Width = 1000
   'Col.Visible = False
   Col.Caption = MapText("INV")
   
   Set Col = GridEX1.Columns.add '9
   Col.Width = 0
   Col.Visible = False
   Col.Caption = MapText("AUTO GEN")
      
   Set Col = GridEX1.Columns.add '2
   Col.Width = 1000
   Col.Caption = MapText("สร้าง")
   
   Set Col = GridEX1.Columns.add '2
   Col.Width = 1000
   Col.Caption = MapText("แก้ไข")

If DocumentType = 2 Or DocumentType = 3 Or DocumentType = 4 Or DocumentType = 18 Or DocumentType = 19 Or DocumentType = 100 Or DocumentType = 101 Or DocumentType = 102 Or DocumentType = 103 Then
   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = MapText("อนุมัติ")
   
   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = MapText("อนุมัติแก้ไขบิล")
Else
   Set Col = GridEX1.Columns.add '2
   Col.Width = 1000
   Col.Caption = MapText("อนุมัติ")
   
   Set Col = GridEX1.Columns.add '2
   Col.Width = 1000
   Col.Caption = MapText("อนุมัติแก้ไขบิล")
End If

   If DocumentType = 1000 Or DocumentType = 1001 Or DocumentType = 1002 Or DocumentType = 1003 Then
      Set Col = GridEX1.Columns.add '2
      Col.Width = 1100
      Col.Caption = MapText("สถานะ PO")
   
      Set Col = GridEX1.Columns.add '2
      Col.Width = 1500
      Col.Caption = MapText("ผู้ตรวจสอบ PO")
   Else
      Set Col = GridEX1.Columns.add '2
      Col.Width = 0
      Col.Caption = MapText("สถานะ PO")
   
      Set Col = GridEX1.Columns.add '2
      Col.Width = 0
      Col.Caption = MapText("ผู้ตรวจสอบ PO")
   End If
      
   Set Col = GridEX1.Columns.add '2
   Col.Width = 1700
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ยอดเงิน")
   
   If DocumentType = 1 Then
      Set Col = GridEX1.Columns.add '2
      Col.Width = 1700
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("สถานะเครดิต")
   Else
      Set Col = GridEX1.Columns.add '2
      Col.Width = 0
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("สถานะเครดิต")
   End If
   
   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("INVENTORY_WH_DOC")
   
   If DocumentType = 1 Then
      Set Col = GridEX1.Columns.add '2
      Col.Width = 1700
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("เลขที่ใบขึ้นอาหาร")
'
      Set Col = GridEX1.Columns.add '2
      Col.Width = 2200
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("สถานะเอกสาร")
      
      Set Col = GridEX1.Columns.add '2
      Col.Width = 1000
      Col.TextAlignment = jgexAlignLeft
      Col.Caption = MapText("คิดแบบ")
   ElseIf DocumentType = 19 Then
      Set Col = GridEX1.Columns.add '2
      Col.Width = 0
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("เลขที่ใบขึ้นอาหาร")
'
      Set Col = GridEX1.Columns.add '2
      Col.Width = 2200
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("สถานะเอกสาร")
      
      Set Col = GridEX1.Columns.add '2
      Col.Width = 0
      Col.Caption = MapText("ผู้แก้ไขเงื่อนไขการคิดราคา")
   Else
      Set Col = GridEX1.Columns.add '2
      Col.Width = 0
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("เลขที่ใบขึ้นอาหาร")
'
      Set Col = GridEX1.Columns.add '2
      Col.Width = 0
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("สถานะเอกสาร")
   End If
   
   GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
Dim Programowner As String
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Programowner = glbParameterObj.Programowner
      
   If DocumentType = 1 Then
      Me.Caption = MapText("ใบส่งสินค้า (ขาย)")
   ElseIf DocumentType = 2 Then
      Me.Caption = MapText("ใบเสร็จรับเงิน (ขาย)")
   ElseIf DocumentType = 3 Then
      Me.Caption = MapText("ใบลดหนี้ (ขาย)")
   ElseIf DocumentType = 4 Then
      Me.Caption = MapText("ใบเพิ่มหนี้ (ขาย)")
   ElseIf DocumentType = 5 Then
      Me.Caption = MapText("ใบกำกับภาษี (ขาย)")
   ElseIf DocumentType = 6 Then
      Me.Caption = MapText("ใบสรุปวางบิล (ขาย)")
   ElseIf DocumentType = 7 Then
      Me.Caption = MapText("ใบส่งสินค้า (ซื้อ)")
   ElseIf DocumentType = 8 Then
      Me.Caption = MapText("ใบเสร็จรับเงิน (ซื้อ)")
   ElseIf DocumentType = 9 Then
      Me.Caption = MapText("ใบลดหนี้ (ซื้อ)")
   ElseIf DocumentType = 10 Then
      Me.Caption = MapText("ใบเพิ่มหนี้ (ซื้อ)")
   ElseIf DocumentType = 11 Then
      Me.Caption = MapText("ใบกำกับภาษี (ซื้อ)")
   ElseIf DocumentType = 12 Then
      Me.Caption = MapText("ใบรับงาน/สั่งงาน (PO ขาย)")
   ElseIf DocumentType = 13 Then
      Me.Caption = MapText("ใบเสนอราคา (ซื้อ)")
   ElseIf DocumentType = 14 Then
      Me.Caption = MapText("ใบเสนอราคา (ขาย)")
   ElseIf DocumentType = 15 Then
      Me.Caption = MapText("ใบบรรจุหีบห่อ(ซื้อ)")
   ElseIf DocumentType = 16 Then
      Me.Caption = MapText("ใบ MEMO ธนาคาร")
   ElseIf DocumentType = 17 Then
      Me.Caption = MapText("ใบบรรจุหีบห่อ(ขาย)")
   ElseIf DocumentType = 18 Then
      Me.Caption = MapText("ใบรับคืนสินค้า (ขาย)")
   ElseIf DocumentType = 19 Then
      Me.Caption = MapText("ใบ Sale Order")
   ElseIf DocumentType = 100 Then
      Me.Caption = MapText("ใบรับเข้าวัตถุดิบ")
   ElseIf DocumentType = 101 Then
      Me.Caption = MapText("ใบรับเข้าวัสดุอุปกรณ์")
   ElseIf DocumentType = 102 Then
      Me.Caption = MapText("ใบรับเข้าจ่ายออกวัสดุอุปกรณ์")
   ElseIf DocumentType = 103 Then
      Me.Caption = MapText("ใบรับเข้าทั่วไป")
   ElseIf DocumentType = 110 Then
      Me.Caption = MapText("ใบรับคืนสินค้า (ซื้อ)")
    ElseIf DocumentType = 111 Then
      Me.Caption = MapText("ใบค่าใช้จ่ายโรงงาน (เงินสดย่อย)")
   ElseIf DocumentType = 1000 Then
      Me.Caption = MapText("PO สั่งซื้อวัตถุดิบ")
   ElseIf DocumentType = 1001 Then
      Me.Caption = MapText("PO สั่งซื้อวัสดุอุปกรณ์")
   ElseIf DocumentType = 1002 Then
      Me.Caption = MapText("PO สั่งซื้อ รับเข้าจ่ายออกวัสดุอุปกรณ์")
   ElseIf DocumentType = 1003 Then
      Me.Caption = MapText("PO สั่งซื้อทั่วไป")
   
   End If
   
   Call InitGrid
   
   Call InitNormalLabel(lblDocumentDate, MapText("จากวันที่"))
   Call InitNormalLabel(lblToDate, MapText("ถึงวันที่"))
   If Area = 1 Then
      Call InitNormalLabel(lblCustomerCode, MapText("รหัสลูกค้า"))
   Else
      Call InitNormalLabel(lblCustomerCode, MapText("รหัสซัพพลายเออร์"))
   End If
   Call InitNormalLabel(lblDocumentNo, MapText("เลขที่เอกสาร"))
   Call InitNormalLabel(lblAccountNo, MapText("เลขที่บัญชี"))
   Call InitNormalLabel(lblPartNo, MapText("รหัสสินค้า"))
   Call InitNormalLabel(lblFeatureCode, MapText("รหัสบริการ"))
   Call InitNormalLabel(lblOrderBy, MapText("เรียงตาม"))
   Call InitNormalLabel(lblOrderType, MapText("เรียงจาก"))
   
   Call InitNormalLabel(lblOtherFilter, MapText("ค้นหาอื่นๆ"))
   
   Call InitCheckBox(chkCommit, "คำนวณ")
   Call InitCheckBox(chkPayFlag, "สรุปจ่าย")
   
   
   If Area = 1 Then
      Call txtCustomerCode.SetKeySearch("CUSTOMER_CODE")
   Else
      Call txtCustomerCode.SetKeySearch("SUPPLIER_CODE")
   End If
   Call txtPartNo.SetKeySearch("PART_NO")
   
   
   Call InitCombo(cboOtherFilter)
   Call InitCombo(cboOrderBy)
   Call InitCombo(cboOrderType)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdClear.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
'   cmdAdjust.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOther.Picture = LoadPicture(glbParameterObj.NormalButton1)
'   cmdDelete.Enabled = False
   
   'Call InitMainButton(cmdAdjust, MapText("ปรับยอด"))
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา (F5)"))
   Call InitMainButton(cmdClear, MapText("เคลียร์ (F4)"))
   
   Call InitMainButton(cmdOther, MapText("อื่นๆ"))

   If DocumentType = 100 Or DocumentType = 101 Or DocumentType = 102 Or DocumentType = 103 Then
      cmdOther.Visible = True
   End If
   
   lblOtherFilter.Visible = False
   cboOtherFilter.Visible = False
   If DocumentType = 100 Or DocumentType = 101 Or DocumentType = 102 Or DocumentType = 103 Or DocumentType = 1000 Or DocumentType = 1001 Or DocumentType = 1002 Or DocumentType = 1003 Then
      lblOtherFilter.Visible = True
      cboOtherFilter.Visible = True
   End If
   
  pnlHeader.Caption = Me.Caption
End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   m_TableName = "USER_GROUP"
   
   Set m_BillingDoc = New CBillingDoc
   Set m_TempBillingDoc = New CBillingDoc
   Set m_Rs = New ADODB.Recordset
   Set m_BillTrans = New Collection

   If DocumentType = 1 Then
      m_IvdDocType = 10
   ElseIf DocumentType = 2 Then
      m_IvdDocType = 21
   End If
   
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_BillTrans = Nothing
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
Dim BD As CBillingDoc
Dim IsOK As Boolean
Dim OKClick As Boolean

   If GridEX1.ItemCount <= 0 Then
         Exit Sub
   End If
   
   TempID1 = GridEX1.Value(1)
   If Button = 2 Then
      Set oMenu = New cPopupMenu
'     lMenuChosen = oMenu.Popup("คัดลอกข้อมูล", "-", "ทดสอบ")
     lMenuChosen = oMenu.Popup("คัดลอกข้อมูล")
      If lMenuChosen = 0 Then
         Exit Sub
      End If
      Set oMenu = Nothing
   Else
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   If lMenuChosen = 1 Then
      If Not (Area = 1) Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      Set BD = New CBillingDoc
      BD.BILLING_DOC_ID = TempID1
      Call glbDaily.CopyBillingDoc(BD, IsOK, True, Area, m_IvdDocType, glbErrorLog)
      Call QueryData(True)
      Set BD = Nothing
   ElseIf lMenuChosen = 3 Then
   If DocumentType = 19 Then
   Dim TempStr As String
     If Area = 1 Then
      TempStr = "(ขาย)"
   ElseIf Area = 2 Then
      TempStr = "(ซื้อ)"
   End If
   
       frmAddEditSaleOrder2.Area = Area
      frmAddEditSaleOrder2.DocumentType = DocumentType
      frmAddEditSaleOrder2.ID = Val(GridEX1.Value(1)) 'ID
       frmAddEditSaleOrder2.HeaderText = MapText("แก้ไขข้อมูลใบ SO " & TempStr)
      frmAddEditSaleOrder2.ShowMode = SHOW_EDIT
      Load frmAddEditSaleOrder2
      frmAddEditSaleOrder2.Show 1
      
      OKClick = frmAddEditSaleOrder2.OKClick
      
      Unload frmAddEditSaleOrder2
      Set frmAddEditSaleOrder2 = Nothing
      
      If OKClick Then
         Call QueryData(True)
      End If
   ElseIf DocumentType = 1 Then
         frmAddEditDO2.Area = Area
      frmAddEditDO2.DocumentType = DocumentType
      frmAddEditDO2.ID = Val(GridEX1.Value(1)) 'ID
      frmAddEditDO2.HeaderText = MapText("แก้ไขข้อมูลใบส่งสินค้า ")
      frmAddEditDO2.ShowMode = SHOW_EDIT
      Load frmAddEditDO2
      frmAddEditDO2.Show 1
      
      OKClick = frmAddEditDO2.OKClick
      
      Unload frmAddEditDO2
      Set frmAddEditDO2 = Nothing
   End If
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   RowBuffer.RowStyle = RowBuffer.Value(6)
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long
Dim fmsTemp As JSFormatStyle

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
   Call m_TempBillingDoc.PopulateFromRS(1, m_Rs)

   Values(1) = m_TempBillingDoc.BILLING_DOC_ID
   Values(2) = m_TempBillingDoc.DOCUMENT_NO
   Values(3) = DateToStringExtEx2(m_TempBillingDoc.DOCUMENT_DATE)
   If Area = 1 Then
      Values(4) = m_TempBillingDoc.CUSTOMER_CODE
      Values(5) = m_TempBillingDoc.CUSTOMER_NAME
   ElseIf Area = 2 Then
      Values(4) = m_TempBillingDoc.SUPPLIER_CODE
      Values(5) = m_TempBillingDoc.SUPPLIER_NAME
   End If
   Values(6) = m_TempBillingDoc.COMMIT_FLAG
   Values(7) = m_TempBillingDoc.RECEIPT_TYPE
   Values(8) = m_TempBillingDoc.PAYMENT_ID
   Values(9) = m_TempBillingDoc.DOCUMENT_TYPE
   Values(10) = m_TempBillingDoc.INVENTORY_DOC_ID
   Values(11) = m_TempBillingDoc.AUTO_GEN_FLAG
   Values(12) = m_TempBillingDoc.CREATE_NAME
   Values(13) = m_TempBillingDoc.MODIFY_NAME
   Values(14) = m_TempBillingDoc.APPROVE_NAME
   Values(15) = m_TempBillingDoc.APPROVE_MANAGE_NAME
   Values(16) = IIf(m_TempBillingDoc.CLOSE_FLAG = "Y", "ปิดแล้ว", "")
   Values(17) = m_TempBillingDoc.VERIFY_BY_NAME
   Values(18) = FormatNumber(m_TempBillingDoc.TOTAL_PRICE)
   If m_TempBillingDoc.CASH_FLAG = "Y" Then
      Values(19) = "ชำระเงินสด"
   Else
      Values(19) = "เงินเชื่อ"
   End If
   Values(20) = m_TempBillingDoc.INVENTORY_WH_DOC_ID
   Values(21) = m_TempBillingDoc.REF
   Values(22) = ConvertSuccessFlagBillingSo(m_TempBillingDoc.SUCCESS_FLAG)
   If DocumentType = 19 Then
      Values(23) = m_TempBillingDoc.USER_APPLOVE_PRICE_THINK
   ElseIf DocumentType = 1 Then
      Dim Cd As CBillTransport
        Set Cd = GetObject("CBillTransport", m_BillTrans, Trim(str(m_TempBillingDoc.BILLING_DOC_ID)), False)
       If (Cd Is Nothing) And (m_TempBillingDoc.PRICE_THINK_TYPE = 2 Or m_TempBillingDoc.PRICE_THINK_TYPE = 3) Then
         Values(23) = "* " & ConvertThinkRate(m_TempBillingDoc.PRICE_THINK_TYPE)
      Else
          Values(23) = ConvertThinkRate(m_TempBillingDoc.PRICE_THINK_TYPE)
       End If
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
   cmdOK.Top = ScaleHeight - 580
   cmdExit.Top = ScaleHeight - 580
   cmdOther.Top = ScaleHeight - 580
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
   cmdOther.Left = cmdOK.Left - cmdOther.Width - 50
End Sub

