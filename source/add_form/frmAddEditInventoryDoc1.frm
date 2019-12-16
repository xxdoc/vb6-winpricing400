VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddEditInventoryDoc1 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditInventoryDoc1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8520
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000009&
         Height          =   1275
         Left            =   0
         ScaleHeight     =   1215
         ScaleWidth      =   1575
         TabIndex        =   47
         Top             =   0
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.ComboBox cboDepartMent 
         Height          =   315
         Left            =   9510
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2160
         Width           =   1905
      End
      Begin prjFarmManagement.uctlTime uctlEntryTime 
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   1740
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
      End
      Begin prjFarmManagement.uctlTextLookup uctlSupplierLookup 
         Height          =   435
         Left            =   6000
         TabIndex        =   4
         Top             =   1260
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlDocumentDate 
         Height          =   405
         Left            =   6000
         TabIndex        =   2
         Top             =   840
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   19
         Top             =   4170
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
      Begin prjFarmManagement.uctlTextBox txtDoNo 
         Height          =   435
         Left            =   1560
         TabIndex        =   3
         Top             =   1290
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtDocumentNo 
         Height          =   435
         Left            =   1560
         TabIndex        =   0
         Top             =   840
         Width           =   2295
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTruckNo 
         Height          =   435
         Left            =   1560
         TabIndex        =   9
         Top             =   2160
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   3015
         Left            =   150
         TabIndex        =   20
         Top             =   4710
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   5318
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
         Column(1)       =   "frmAddEditInventoryDoc1.frx":27A2
         Column(2)       =   "frmAddEditInventoryDoc1.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditInventoryDoc1.frx":290E
         FormatStyle(2)  =   "frmAddEditInventoryDoc1.frx":2A6A
         FormatStyle(3)  =   "frmAddEditInventoryDoc1.frx":2B1A
         FormatStyle(4)  =   "frmAddEditInventoryDoc1.frx":2BCE
         FormatStyle(5)  =   "frmAddEditInventoryDoc1.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditInventoryDoc1.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   35
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtSender 
         Height          =   435
         Left            =   1560
         TabIndex        =   13
         Top             =   2610
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtReceiver 
         Height          =   435
         Left            =   6000
         TabIndex        =   14
         Top             =   2640
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtDeliveryFee 
         Height          =   435
         Left            =   1560
         TabIndex        =   16
         Top             =   3540
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtMaterialPrice 
         Height          =   435
         Left            =   6000
         TabIndex        =   17
         Top             =   3570
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTotal 
         Height          =   435
         Left            =   9510
         TabIndex        =   18
         Top             =   3540
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtQueNo 
         Height          =   435
         Left            =   9510
         TabIndex        =   15
         Top             =   2610
         Width           =   1365
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtDesc 
         Height          =   435
         Left            =   6000
         TabIndex        =   8
         Top             =   1710
         Width           =   5385
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTime uctlExitTime 
         Height          =   375
         Left            =   3270
         TabIndex        =   7
         Top             =   1770
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
      End
      Begin prjFarmManagement.uctlTextBox txtCredit 
         Height          =   435
         Left            =   10830
         TabIndex        =   48
         Top             =   840
         Width           =   525
         _ExtentX        =   2408
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPrNo 
         Height          =   435
         Left            =   1560
         TabIndex        =   51
         Top             =   3080
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin VB.Label lblPrNo 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   0
         TabIndex        =   52
         Top             =   3180
         Width           =   1485
      End
      Begin VB.Label Label6 
         Height          =   315
         Left            =   11400
         TabIndex        =   50
         Top             =   870
         Width           =   405
      End
      Begin VB.Label lblCredit 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   9960
         TabIndex        =   49
         Top             =   870
         Width           =   855
      End
      Begin Threed.SSCommand cmdSupplierSearch 
         Height          =   405
         Left            =   11370
         TabIndex        =   5
         Top             =   1260
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditInventoryDoc1.frx":2F36
         ButtonStyle     =   3
      End
      Begin VB.Label lblDepartment 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   345
         Left            =   8400
         TabIndex        =   46
         Top             =   2220
         Width           =   1035
      End
      Begin Threed.SSCheck chkException 
         Height          =   435
         Left            =   7620
         TabIndex        =   11
         Top             =   2190
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2730
         TabIndex        =   45
         Top             =   1800
         Width           =   435
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   525
         Left            =   6840
         TabIndex        =   24
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditInventoryDoc1.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAuto 
         Height          =   405
         Left            =   3870
         TabIndex        =   1
         Top             =   840
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditInventoryDoc1.frx":356A
         ButtonStyle     =   3
      End
      Begin VB.Label lblQueNo 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   8970
         TabIndex        =   44
         Top             =   2670
         Width           =   465
      End
      Begin Threed.SSCheck chkCommit 
         Height          =   435
         Left            =   6000
         TabIndex        =   10
         Top             =   2190
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4470
         TabIndex        =   43
         Top             =   1740
         Width           =   1455
      End
      Begin VB.Label lblSupplierNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4470
         TabIndex        =   42
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label4 
         Height          =   315
         Left            =   10950
         TabIndex        =   41
         Top             =   3660
         Width           =   585
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8280
         TabIndex        =   40
         Top             =   3660
         Width           =   1125
      End
      Begin VB.Label lblMaterialPrice 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4170
         TabIndex        =   39
         Top             =   3690
         Width           =   1695
      End
      Begin VB.Label Label2 
         Height          =   315
         Left            =   7440
         TabIndex        =   38
         Top             =   3660
         Width           =   765
      End
      Begin VB.Label Label1 
         Height          =   315
         Left            =   3900
         TabIndex        =   37
         Top             =   3660
         Width           =   765
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4710
         TabIndex        =   36
         Top             =   870
         Width           =   1155
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   25
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditInventoryDoc1.frx":3884
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   26
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   22
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   21
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditInventoryDoc1.frx":3B9E
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   23
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditInventoryDoc1.frx":3EB8
         ButtonStyle     =   3
      End
      Begin VB.Label lblDeliveryFee 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -210
         TabIndex        =   34
         Top             =   3660
         Width           =   1695
      End
      Begin VB.Label lblTruckNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -60
         TabIndex        =   33
         Top             =   2250
         Width           =   1575
      End
      Begin VB.Label lblDeliveryNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -60
         TabIndex        =   32
         Top             =   1830
         Width           =   1575
      End
      Begin VB.Label lblDocumentNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -150
         TabIndex        =   31
         Top             =   900
         Width           =   1665
      End
      Begin VB.Label lblReceiver 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4410
         TabIndex        =   30
         Top             =   2700
         Width           =   1485
      End
      Begin VB.Label lblSender 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   30
         TabIndex        =   29
         Top             =   2670
         Width           =   1485
      End
      Begin VB.Label lblDoNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -60
         TabIndex        =   28
         Top             =   1380
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmAddEditInventoryDoc1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_InventoryDoc As CInventoryDoc
Private m_Suppliers As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean

Public id As Long
Public DocumentType As Long

Private FileName As String
Private m_SumUnit As Double
Private m_SumTotalPrice As Double

Private Sub CreateImportExportItems()
Dim Ti As CTransferItem
Dim Ei As CLotItem
Dim II As CLotItem

   Set m_InventoryDoc.ImportExports = Nothing
   Set m_InventoryDoc.ImportExports = New Collection

   For Each Ti In m_InventoryDoc.TransferItems
      Set Ei = Ti.ExportItem
      Set II = Ti.ImportItem

      Ei.Flag = Ti.Flag
      II.Flag = Ti.Flag

      Call m_InventoryDoc.ImportExports.add(II)
      Call m_InventoryDoc.ImportExports.add(Ei)
   Next Ti
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
      m_InventoryDoc.INVENTORY_DOC_ID = id
      m_InventoryDoc.COMMIT_FLAG = ""
      If Not glbDaily.QueryInventoryDoc(m_InventoryDoc, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_InventoryDoc.PopulateFromRS(1, m_Rs)
      
      uctlDocumentDate.ShowDate = m_InventoryDoc.DOCUMENT_DATE
      txtDoNo.Text = m_InventoryDoc.DO_NO
      txtTruckNo.Text = m_InventoryDoc.TRUCK_NO
      txtDocumentNo.Text = m_InventoryDoc.DOCUMENT_NO
      txtDeliveryFee.Text = Format(m_InventoryDoc.DELIVERY_FEE, "0.00")
      txtSender.Text = m_InventoryDoc.SENDER_NAME
      txtReceiver.Text = m_InventoryDoc.RECEIVE_NAME
      uctlSupplierLookup.MyCombo.ListIndex = IDToListIndex(uctlSupplierLookup.MyCombo, m_InventoryDoc.SUPPLIER_ID)
      cboDepartMent.ListIndex = IDToListIndex(cboDepartMent, m_InventoryDoc.DEPARTMENT_ID)
      chkCommit.Value = FlagToCheck(m_InventoryDoc.COMMIT_FLAG)
      txtQueNo.Text = m_InventoryDoc.QUE_NO
      txtDesc.Text = m_InventoryDoc.DOCUMENT_DESC
      uctlEntryTime.HR = HOUR(m_InventoryDoc.ENTRY_DATE)
      uctlEntryTime.MI = Minute(m_InventoryDoc.ENTRY_DATE)
      uctlExitTime.HR = HOUR(m_InventoryDoc.EXIT_DATE)
      uctlExitTime.MI = Minute(m_InventoryDoc.EXIT_DATE)
      chkException.Value = FlagToCheck(m_InventoryDoc.EXCEPTION_FLAG)
      txtCredit.Text = m_InventoryDoc.Credit
      txtPrNo.Text = m_InventoryDoc.PR_NO
      
      cmdAdd.Enabled = (m_InventoryDoc.OLD_COMMIT_FLAG = "N")
      cmdDelete.Enabled = (m_InventoryDoc.OLD_COMMIT_FLAG = "N")
      chkCommit.Enabled = (m_InventoryDoc.OLD_COMMIT_FLAG = "N")
'      txtDeliveryFee.Enabled = (m_InventoryDoc.OLD_COMMIT_FLAG = "N")

      If DocumentType = 20 Then
         Call glbDaily.CreateTransferItems(m_InventoryDoc)
      End If
   End If
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call TabStrip1_Click
   
   Call EnableForm(Me, True)
End Sub

Private Sub cboGroup_Click()
   m_HasModify = True
End Sub

Private Sub chkEnable_Click()
   m_HasModify = True
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
   If ShowMode = SHOW_EDIT Then
      If Not VerifyAccessRight("INVENTORY_IMPORT_EDIT") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   End If
   If Not VerifyTextControl(lblDocumentNo, txtDocumentNo, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblDocumentDate, uctlDocumentDate, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblSupplierNo, uctlSupplierLookup.MyCombo, True) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblDeliveryFee, txtDeliveryFee, True) Then
      Exit Function
   End If
   
   If Not CheckUniqueNs(IMPORT_UNIQUE, txtDocumentNo.Text, id) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtDocumentNo.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_InventoryDoc.AddEditMode = ShowMode
   m_InventoryDoc.INVENTORY_DOC_ID = id
    m_InventoryDoc.DOCUMENT_DATE = uctlDocumentDate.ShowDate
   m_InventoryDoc.DO_NO = txtDoNo.Text
   m_InventoryDoc.TRUCK_NO = txtTruckNo.Text
   m_InventoryDoc.DOCUMENT_NO = txtDocumentNo.Text
   m_InventoryDoc.DELIVERY_FEE = Val(txtDeliveryFee.Text)
   m_InventoryDoc.SENDER_NAME = txtSender.Text
   m_InventoryDoc.RECEIVE_NAME = txtReceiver.Text
   m_InventoryDoc.SUPPLIER_ID = uctlSupplierLookup.MyCombo.ItemData(Minus2Zero(uctlSupplierLookup.MyCombo.ListIndex))
   m_InventoryDoc.DELIVERY_ID = -1
   m_InventoryDoc.DOCUMENT_TYPE = DocumentType
   m_InventoryDoc.COMMIT_FLAG = Check2Flag(chkCommit.Value)
   m_InventoryDoc.QUE_NO = txtQueNo.Text
   m_InventoryDoc.DOCUMENT_DESC = txtDesc.Text
   m_InventoryDoc.ENTRY_DATE = uctlDocumentDate.ShowDate
   m_InventoryDoc.ENTRY_DATE = DateAdd("h", uctlEntryTime.HR, m_InventoryDoc.ENTRY_DATE)
   m_InventoryDoc.ENTRY_DATE = DateAdd("n", uctlEntryTime.MI, m_InventoryDoc.ENTRY_DATE)
   m_InventoryDoc.EXIT_DATE = uctlDocumentDate.ShowDate
   m_InventoryDoc.EXIT_DATE = DateAdd("h", uctlExitTime.HR, m_InventoryDoc.EXIT_DATE)
   m_InventoryDoc.EXIT_DATE = DateAdd("n", uctlExitTime.MI, m_InventoryDoc.EXIT_DATE)
   m_InventoryDoc.EXCEPTION_FLAG = Check2Flag(chkException.Value)
   m_InventoryDoc.DEPARTMENT_ID = cboDepartMent.ItemData(Minus2Zero(cboDepartMent.ListIndex))
   m_InventoryDoc.Credit = Val(txtCredit.Text)
   m_InventoryDoc.PR_NO = txtPrNo.Text
   
   If DocumentType = 20 Then
      Call CreateImportExportItems
   End If
   
   If m_InventoryDoc.COMMIT_FLAG = "Y" Then
      If m_InventoryDoc.OLD_COMMIT_FLAG <> "Y" Then
         Call glbDaily.TriggerCommit(m_InventoryDoc.ImportExports)
      End If
   End If
   Call CalculateIncludePrice
   
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditInventoryDoc(m_InventoryDoc, IsOK, True, glbErrorLog) Then
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

Private Sub cboBusinessGroup_Click()
   m_HasModify = True
End Sub

Private Sub cboBusinessType_Click()
   m_HasModify = True
End Sub

Private Sub cboEnterpriseType_Click()
   m_HasModify = True
End Sub

Private Sub cboDepartment_Click()
   m_HasModify = True
End Sub

Private Sub cboDepartMent_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub chkCommit_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkCommit_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub chkException_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkException_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean

   If Not cmdAdd.Enabled Then
      Exit Sub
   End If
   
   OKClick = False
   If TabStrip1.SelectedItem.Index = 1 Then
      If DocumentType = 1 Then
         frmAddEditImportItem.SupplierID = uctlSupplierLookup.MyCombo.ItemData(Minus2Zero(uctlSupplierLookup.MyCombo.ListIndex))
         frmAddEditImportItem.COMMIT_FLAG = m_InventoryDoc.OLD_COMMIT_FLAG
         Set frmAddEditImportItem.TempCollection = m_InventoryDoc.ImportExports
         frmAddEditImportItem.ParentShowMode = ShowMode
         frmAddEditImportItem.ShowMode = SHOW_ADD
         frmAddEditImportItem.HeaderText = MapText("เพิ่มรายการนำเข้า")
         Load frmAddEditImportItem
         frmAddEditImportItem.Show 1
   
         OKClick = frmAddEditImportItem.OKClick
   
         Unload frmAddEditImportItem
         Set frmAddEditImportItem = Nothing
      ElseIf (DocumentType = 19) Then
         frmAddEditImportItemEx.SupplierID = uctlSupplierLookup.MyCombo.ItemData(Minus2Zero(uctlSupplierLookup.MyCombo.ListIndex))
         frmAddEditImportItemEx.COMMIT_FLAG = m_InventoryDoc.OLD_COMMIT_FLAG
         Set frmAddEditImportItemEx.TempCollection = m_InventoryDoc.ImportExports
         frmAddEditImportItemEx.ParentShowMode = ShowMode
         frmAddEditImportItemEx.ShowMode = SHOW_ADD
         frmAddEditImportItemEx.HeaderText = MapText("เพิ่มรายการนำเข้า")
         Load frmAddEditImportItemEx
         frmAddEditImportItemEx.Show 1
   
         OKClick = frmAddEditImportItemEx.OKClick
   
         Unload frmAddEditImportItemEx
         Set frmAddEditImportItemEx = Nothing
      ElseIf (DocumentType = 23) Then
         frmAddEditImportItemEx3.SupplierID = uctlSupplierLookup.MyCombo.ItemData(Minus2Zero(uctlSupplierLookup.MyCombo.ListIndex))
         frmAddEditImportItemEx3.COMMIT_FLAG = m_InventoryDoc.OLD_COMMIT_FLAG
         Set frmAddEditImportItemEx3.TempCollection = m_InventoryDoc.ImportExports
         frmAddEditImportItemEx3.ParentShowMode = ShowMode
         frmAddEditImportItemEx3.ShowMode = SHOW_ADD
         frmAddEditImportItemEx3.HeaderText = MapText("เพิ่มรายการนำเข้า")
         Load frmAddEditImportItemEx3
         frmAddEditImportItemEx3.Show 1
   
         OKClick = frmAddEditImportItemEx3.OKClick
   
         Unload frmAddEditImportItemEx3
         Set frmAddEditImportItemEx3 = Nothing
      ElseIf DocumentType = 20 Then
         frmAddEditImportItemEx2.SupplierID = uctlSupplierLookup.MyCombo.ItemData(Minus2Zero(uctlSupplierLookup.MyCombo.ListIndex))
         frmAddEditImportItemEx2.COMMIT_FLAG = m_InventoryDoc.OLD_COMMIT_FLAG
         Set frmAddEditImportItemEx2.TempCollection = m_InventoryDoc.TransferItems
         frmAddEditImportItemEx2.ParentShowMode = ShowMode
         frmAddEditImportItemEx2.ShowMode = SHOW_ADD
         frmAddEditImportItemEx2.HeaderText = MapText("เพิ่มรายการนำเข้า")
         Load frmAddEditImportItemEx2
         frmAddEditImportItemEx2.Show 1
   
         OKClick = frmAddEditImportItemEx2.OKClick
   
         Unload frmAddEditImportItemEx2
         Set frmAddEditImportItemEx2 = Nothing
      End If
      If OKClick Then
         Call GetTotalPrice
         
         If DocumentType <> 20 Then
            GridEX1.ItemCount = CountItem(m_InventoryDoc.ImportExports)
         Else
            GridEX1.ItemCount = CountItem(m_InventoryDoc.TransferItems)
         End If
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdAuto_Click()
Dim No As String

   If Trim(txtDocumentNo.Text) = "" Then
      Call glbDatabaseMngr.GenerateNumber(IMPORT_NUMBER, No, glbErrorLog)
      txtDocumentNo.Text = No
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
   
   If TabStrip1.SelectedItem.Index = 1 Then
      If DocumentType <> 20 Then
         If ID1 <= 0 Then
            m_InventoryDoc.ImportExports.Remove (ID2)
         Else
            m_InventoryDoc.ImportExports.Item(ID2).Flag = "D"
         End If
      Else
         If ID1 <= 0 Then
            m_InventoryDoc.TransferItems.Remove (ID2)
         Else
            m_InventoryDoc.TransferItems.Item(ID2).Flag = "D"
         End If
      End If
      
      Call GetTotalPrice
      If DocumentType <> 20 Then
         GridEX1.ItemCount = CountItem(m_InventoryDoc.ImportExports)
      Else
         GridEX1.ItemCount = CountItem(m_InventoryDoc.TransferItems)
      End If
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
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
   
   If TabStrip1.SelectedItem.Index = 1 Then
      If DocumentType = 1 Then
         frmAddEditImportItem.id = id
         frmAddEditImportItem.SupplierID = uctlSupplierLookup.MyCombo.ItemData(Minus2Zero(uctlSupplierLookup.MyCombo.ListIndex))
         frmAddEditImportItem.COMMIT_FLAG = m_InventoryDoc.OLD_COMMIT_FLAG
         Set frmAddEditImportItem.TempCollection = m_InventoryDoc.ImportExports
         frmAddEditImportItem.HeaderText = MapText("แก้ไขรายการนำเข้า")
         frmAddEditImportItem.ParentShowMode = ShowMode
         frmAddEditImportItem.ShowMode = SHOW_EDIT
         Load frmAddEditImportItem
         frmAddEditImportItem.Show 1
         
         OKClick = frmAddEditImportItem.OKClick
   
         Unload frmAddEditImportItem
         Set frmAddEditImportItem = Nothing
      ElseIf (DocumentType = 19) Then
         frmAddEditImportItemEx.id = id
         frmAddEditImportItemEx.SupplierID = uctlSupplierLookup.MyCombo.ItemData(Minus2Zero(uctlSupplierLookup.MyCombo.ListIndex))
         frmAddEditImportItemEx.COMMIT_FLAG = m_InventoryDoc.OLD_COMMIT_FLAG
         Set frmAddEditImportItemEx.TempCollection = m_InventoryDoc.ImportExports
         frmAddEditImportItemEx.HeaderText = MapText("แก้ไขรายการนำเข้า")
         frmAddEditImportItemEx.ParentShowMode = ShowMode
         frmAddEditImportItemEx.ShowMode = SHOW_EDIT
         Load frmAddEditImportItemEx
         frmAddEditImportItemEx.Show 1
         
         OKClick = frmAddEditImportItemEx.OKClick
   
         Unload frmAddEditImportItemEx
         Set frmAddEditImportItemEx = Nothing
      ElseIf (DocumentType = 23) Then
         frmAddEditImportItemEx3.id = id
         frmAddEditImportItemEx3.SupplierID = uctlSupplierLookup.MyCombo.ItemData(Minus2Zero(uctlSupplierLookup.MyCombo.ListIndex))
         frmAddEditImportItemEx3.COMMIT_FLAG = m_InventoryDoc.OLD_COMMIT_FLAG
         Set frmAddEditImportItemEx3.TempCollection = m_InventoryDoc.ImportExports
         frmAddEditImportItemEx3.HeaderText = MapText("แก้ไขรายการนำเข้า")
         frmAddEditImportItemEx3.ParentShowMode = ShowMode
         frmAddEditImportItemEx3.ShowMode = SHOW_EDIT
         Load frmAddEditImportItemEx3
         frmAddEditImportItemEx3.Show 1
         
         OKClick = frmAddEditImportItemEx3.OKClick
   
         Unload frmAddEditImportItemEx3
         Set frmAddEditImportItemEx3 = Nothing
      ElseIf DocumentType = 20 Then
         frmAddEditImportItemEx2.id = id
         frmAddEditImportItemEx2.SupplierID = uctlSupplierLookup.MyCombo.ItemData(Minus2Zero(uctlSupplierLookup.MyCombo.ListIndex))
         frmAddEditImportItemEx2.COMMIT_FLAG = m_InventoryDoc.OLD_COMMIT_FLAG
         Set frmAddEditImportItemEx2.TempCollection = m_InventoryDoc.TransferItems
         frmAddEditImportItemEx2.HeaderText = MapText("แก้ไขรายการนำเข้า")
         frmAddEditImportItemEx2.ParentShowMode = ShowMode
         frmAddEditImportItemEx2.ShowMode = SHOW_EDIT
         Load frmAddEditImportItemEx2
         frmAddEditImportItemEx2.Show 1
         
         OKClick = frmAddEditImportItemEx2.OKClick
   
         Unload frmAddEditImportItemEx2
         Set frmAddEditImportItemEx2 = Nothing
      End If

      If OKClick Then
         Call GetTotalPrice
         If DocumentType <> 20 Then
            GridEX1.ItemCount = CountItem(m_InventoryDoc.ImportExports)
         Else
            GridEX1.ItemCount = CountItem(m_InventoryDoc.TransferItems)
         End If
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub CalculateIncludePrice()
Dim II As CLotItem
   
   'ไม่ต้องเอา II.EXPENSE1 + II.EXPENSE2 มารวมด้วย เพราะจะถูกกระจายไปไว้ที่ txtDeliveryFee แล้ว (ยูสเซอร์ไม่ต้องคีย์) สำหรับใบรับวัตถุดิบ
   'แต่ถ้าเป็นซื้ออย่างอื่น II.EXPENSE1 + II.EXPENSE2 จะมีค่าเป็น 0 แต่ยูสเซอร์จะคีย์ txtDeliveryFee เอง
   For Each II In m_InventoryDoc.ImportExports
      If II.Flag <> "D" Then
         II.TOTAL_INCLUDE_PRICE = II.TOTAL_ACTUAL_PRICE + (MyDiff(II.TOTAL_ACTUAL_PRICE, m_SumTotalPrice) * Val(txtDeliveryFee.Text))
         II.INCLUDE_UNIT_PRICE = MyDiffEx(II.TOTAL_INCLUDE_PRICE, II.TX_AMOUNT)
         
         If II.Flag <> "A" Then
            II.Flag = "E"
         End If
      End If
   Next II
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
      id = m_InventoryDoc.INVENTORY_DOC_ID
      m_InventoryDoc.QueryFlag = 1
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

Private Sub cmdPictureAdd_Click()
On Error Resume Next
Dim strDescription As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Picture Files (*.jpg, *.gif)|*.jpg;*.gif"
   dlgAdd.DialogTitle = "Select Picture to Add to Database"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
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
Dim ClassName As String

   If m_HasModify Or ((Not m_HasModify) And (ShowMode = SHOW_ADD)) Then
      glbErrorLog.LocalErrorMsg = "กรุณาทำการบันทึกข้อมูลให้เรียบร้อยก่อน"
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("ใบรายงานรับวัตถุดิบ", "ปรับค่าหน้ากระดาษ", "-", "ใบรายงานรับของ", "ปรับค่าหน้ากระดาษ")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   
   If lMenuChosen = 1 Then
      If CountItem(m_InventoryDoc.ImportExports) <> 1 Then
         glbErrorLog.LocalErrorMsg = "ใบรายงานรับของจะต้องมีรายการรับเข้าได้เท่ากับ 1 รายการ"
         glbErrorLog.ShowUserError
         
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      ReportKey = "CReportInvDoc001_1"
      ClassName = "CReportInvDoc001_1"
      Set Report = New CReportInvDoc001_1
      ReportFlag = True
   ElseIf lMenuChosen = 2 Then
      ReportKey = "CReportInvDoc001_1"
      ClassName = "CReportInvDoc001_1"
      Set Rc = New CReportConfig
      Rc.REPORT_KEY = ReportKey
      Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
      Call Rc.QueryData(m_Rs, iCount)
      HeaderText = MapText("ใบรายงานรับวัตถุดิบ")
      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
   ElseIf lMenuChosen = 4 Then
      Call LoadPictureFromFile(glbParameterObj.ReceiveVoucher1, Picture2)

      ReportKey = "CReportInvDoc001_2"
      ClassName = "CReportInvDoc001_2"
      Set Report = New CReportInvDoc001_2
      ReportFlag = True
   ElseIf lMenuChosen = 5 Then
      ReportKey = "CReportInvDoc001_2"
      ClassName = "CReportInvDoc001_2"
      Set Rc = New CReportConfig
      Rc.REPORT_KEY = ReportKey
      Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
      Call Rc.QueryData(m_Rs, iCount)
      HeaderText = MapText("ใบรับเข้าสินค้า/วัตถุดิบ")
      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
   End If
   
   If Not Report Is Nothing Then
      Call Report.AddParam(m_InventoryDoc.INVENTORY_DOC_ID, "INVENTORY_DOC_ID")
      Call Report.AddParam(ReportKey, "REPORT_KEY")
      Call Report.AddParam(Picture2.Picture, "BACK_GROUND")
   End If
   
   If ReportFlag Then
      Set frmReport.ReportObject = Report
      frmReport.HeaderText = pnlHeader.Caption
      frmReport.ClassName = ClassName
      Load frmReport
      frmReport.Show 1
   
      Unload frmReport
      Set frmReport = Nothing
      Set Report = Nothing
   Else
      frmReportConfig.ReportMode = 1
      frmReportConfig.ShowMode = EditMode
      frmReportConfig.id = Rc.REPORT_CONFIG_ID
      frmReportConfig.ReportKey = ReportKey
      frmReportConfig.HeaderText = HeaderText
      Load frmReportConfig
      frmReportConfig.Show 1
      
      Unload frmReportConfig
      Set frmReportConfig = Nothing
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cmdSupplierSearch_Click()
Dim OKClick As Boolean
Dim TempCol As Collection
Dim Cs As CSupplier

   Set TempCol = New Collection
   
   Set frmQuerySupplier.TempCollection = TempCol
   frmQuerySupplier.ShowMode = SHOW_ADD
   Load frmQuerySupplier
   frmQuerySupplier.Show 1
   
   OKClick = frmQuerySupplier.OKClick
   
   Unload frmQuerySupplier
   Set frmQuerySupplier = Nothing
   
   If OKClick Then
      Set Cs = TempCol(1)
      uctlSupplierLookup.MyCombo.ListIndex = IDToListIndex(uctlSupplierLookup.MyCombo, Cs.SUPPLIER_ID)
      m_HasModify = True
   End If
   
   Set TempCol = Nothing
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      Call LoadSupplier(uctlSupplierLookup.MyCombo, m_Suppliers)
      Set uctlSupplierLookup.MyCollection = m_Suppliers
      
      Call LoadLayout(cboDepartMent)
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_InventoryDoc.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         uctlDocumentDate.ShowDate = Now
         uctlEntryTime.HR = HOUR(Now)
         uctlEntryTime.MI = Minute(Now)
         
         uctlExitTime.HR = HOUR(Now)
         uctlExitTime.MI = Minute(Now)
         
         m_InventoryDoc.QueryFlag = 0
         Call QueryData(False)
      End If
      
      Call EnableForm(Me, True)
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

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_InventoryDoc = Nothing
   Set m_Suppliers = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
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
   Col.Width = 2100
   Col.Caption = MapText("หมายเลขวัตถุดิบ")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 4425
   Col.Caption = MapText("วัตถุดิบ")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 1785
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ปริมาณ")

   Set Col = GridEX1.Columns.add '6
   Col.Width = 1620
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ราคา")
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = 1620
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ราคารวม")
   
   Set Col = GridEX1.Columns.add '8
   Col.Width = 1980
   Col.Caption = MapText("สถานที่จัดเก็บ")
End Sub

Private Sub GetTotalPrice()
Dim II As CLotItem
Dim Tr As CTransferItem
Dim Sum As Double
Dim Sum1 As Double

   Sum1 = 0
   Sum = 0
   m_SumUnit = 0
   m_SumTotalPrice = 0
   
   If DocumentType <> 20 Then
      For Each II In m_InventoryDoc.ImportExports
         If II.Flag <> "D" Then
            Sum = Sum + CDbl(Format(II.TOTAL_ACTUAL_PRICE, "0.00"))
            m_SumUnit = m_SumUnit + II.TX_AMOUNT
            m_SumTotalPrice = m_SumTotalPrice + II.TOTAL_ACTUAL_PRICE
            
            Sum1 = Sum1 + II.EXPENSE1 + II.EXPENSE2
         End If
      Next II
   Else
      For Each Tr In m_InventoryDoc.TransferItems
         If Tr.Flag <> "D" Then
            Sum = Sum + CDbl(Format(Tr.ImportItem.TOTAL_ACTUAL_PRICE, "0.00"))
            m_SumUnit = m_SumUnit + Tr.ImportItem.TX_AMOUNT
            m_SumTotalPrice = m_SumTotalPrice + Tr.ImportItem.TOTAL_ACTUAL_PRICE
            
            Sum1 = Sum1 + Tr.ImportItem.EXPENSE1 + Tr.ImportItem.EXPENSE2
         End If
      Next Tr
   End If
   
   txtMaterialPrice.Text = Format(Sum, "0.00")
   If (DocumentType = 1) Then
      txtDeliveryFee.Text = Format(Sum1, "0.00")
   End If
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblTruckNo, MapText("ทะเบียนรถ"))
   Call InitNormalLabel(lblQueNo, MapText("คิวที่"))
   Call InitNormalLabel(lblDocumentNo, MapText("เลขที่บิลรับของ"))
   Call InitNormalLabel(lblReceiver, MapText("กรรมกรสาย"))
   Call InitNormalLabel(lblDesc, MapText("หมายเหตุ"))
   Call InitNormalLabel(lblDoNo, MapText("เลขที่ PO"))
   Call InitNormalLabel(lblDeliveryNo, MapText("เวลาเข้า - ออก"))
   Call InitNormalLabel(Label3, MapText("-"))
   Call InitNormalLabel(lblDepartment, MapText("แผนก"))
   Call InitNormalLabel(lblCredit, MapText("เครดิต"))
   Call InitNormalLabel(Label6, MapText("วัน"))
   
   Call InitNormalLabel(lblPrNo, MapText("เลขที่ PR"))
   Call InitNormalLabel(lblSender, MapText("เลขที่ใบส่งของ"))
   If (DocumentType = 19) Or (DocumentType = 20) Then
      Call InitNormalLabel(lblDeliveryFee, MapText("มูลค่า VAT"))
   Else
      Call InitNormalLabel(lblDeliveryFee, MapText("ค่าใช้จ่ายจัดซื้อ"))
   End If
   Call InitNormalLabel(lblDocumentDate, MapText("วันที่เอกสาร"))
   Call InitNormalLabel(lblMaterialPrice, MapText("ราคาวัตถุดิบ"))
   Call InitNormalLabel(Label1, MapText("บาท"))
   Call InitNormalLabel(Label2, MapText("บาท"))
   Call InitNormalLabel(Label4, MapText("บาท"))
   Call InitNormalLabel(lblTotal, MapText("ราคารวม"))
   Call InitNormalLabel(lblSupplierNo, MapText("รหัสซัพ ฯ"))
   Call InitCheckBox(chkCommit, "คำนวณ")
   Call InitCheckBox(chkException, "***")
   
   Call txtDocumentNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtDoNo.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtTruckNo.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtDeliveryFee.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtCredit.SetTextLenType(TEXT_INTEGER, glbSetting.MONEY_TYPE)
   If (DocumentType = 19) Or (DocumentType = 20) Then
      txtDeliveryFee.Enabled = True
   Else
      txtDeliveryFee.Enabled = False
   End If
   Call txtMaterialPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtMaterialPrice.Enabled = False
   Call txtReceiver.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtTotal.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtTotal.Enabled = False
   Call txtQueNo.SetTextLenType(TEXT_INTEGER, glbSetting.MONEY_TYPE)
   Call txtDesc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   
   Call InitCombo(cboDepartMent)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAuto.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPrint.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSupplierSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdPrint, MapText("พิมพ์"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdAuto, MapText("A"))
   Call InitMainButton(cmdSupplierSearch, MapText("..."))
   
   Call InitGrid1
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.NAME = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.add().Caption = MapText("รายการรับวัตถุดิบ")
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
   Call InitFormLayout
      
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_InventoryDoc = New CInventoryDoc
   Set m_Suppliers = New Collection
End Sub


Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   If TabStrip1.SelectedItem.Index = 5 Then
      RowBuffer.RowStyle = RowBuffer.Value(7)
   End If
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Index = 1 Then
      If m_InventoryDoc.ImportExports Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      If DocumentType <> 20 Then
         Dim CR As CLotItem
         If m_InventoryDoc.ImportExports.Count <= 0 Then
            Exit Sub
         End If
         Set CR = GetItem(m_InventoryDoc.ImportExports, RowIndex, RealIndex)
         If CR Is Nothing Then
            Exit Sub
         End If
   
         Values(1) = CR.LOT_ITEM_ID
         Values(2) = RealIndex
         Values(3) = CR.PART_NO
         If CR.PIG_FLAG = "Y" Then
            Values(4) = CR.ITEM_DESC
         Else
            Values(4) = CR.PART_DESC
         End If
         Values(5) = FormatNumber(CR.TX_AMOUNT)
         Values(6) = FormatNumber(CR.ACTUAL_UNIT_PRICE, 4)
         Values(7) = FormatNumber(CR.TOTAL_ACTUAL_PRICE)
         Values(8) = CR.LOCATION_NAME
      Else
         Dim Tr As CTransferItem
         If m_InventoryDoc.TransferItems.Count <= 0 Then
            Exit Sub
         End If
         Set Tr = GetItem(m_InventoryDoc.TransferItems, RowIndex, RealIndex)
         If Tr Is Nothing Then
            Exit Sub
         End If
   
         Values(1) = Tr.ImportItem.LOT_ITEM_ID
         Values(2) = RealIndex
         Values(3) = Tr.ImportItem.PART_NO
         If Tr.ExportItem.PIG_FLAG = "Y" Then
            Values(4) = Tr.ExportItem.ITEM_DESC
         Else
            Values(4) = Tr.ExportItem.PART_DESC
         End If
         Values(5) = FormatNumber(Tr.ImportItem.TX_AMOUNT)
         Values(6) = FormatNumber(Tr.ImportItem.ACTUAL_UNIT_PRICE)
         Values(7) = FormatNumber(Tr.ImportItem.TOTAL_ACTUAL_PRICE)
         Values(8) = Tr.ImportItem.LOCATION_NAME
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
      
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub TabStrip1_Click()
   If TabStrip1.SelectedItem.Index = 1 Then
      Call InitGrid1
      
      Call GetTotalPrice
      If DocumentType <> 20 Then
         GridEX1.ItemCount = CountItem(m_InventoryDoc.ImportExports)
         GridEX1.Rebind
      Else
         GridEX1.ItemCount = CountItem(m_InventoryDoc.TransferItems)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub txtCredit_Change()
   m_HasModify = True
End Sub
Private Sub txtDeliveryFee_Change()
   m_HasModify = True
   txtTotal.Text = Format(Val(txtDeliveryFee.Text) + Val(txtMaterialPrice.Text), "0.00")
End Sub

Private Sub txtDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtDoNo_Change()
   m_HasModify = True
End Sub

Private Sub txtDeliveryNo_Change()
   m_HasModify = True
End Sub

Private Sub txtSellBy_Change()
   m_HasModify = True
End Sub

Private Sub txtDocumentNo_Change()
   m_HasModify = True
End Sub

Private Sub txtPassword_Change()
   m_HasModify = True
End Sub

Private Sub txtMaterialPrice_Change()
   m_HasModify = True
   txtTotal.Text = Format(Val(txtDeliveryFee.Text) + Val(txtMaterialPrice.Text), "0.00")
End Sub
Private Sub txtPrNo_Change()
   m_HasModify = True
End Sub

Private Sub txtQueNo_Change()
   m_HasModify = True
End Sub

Private Sub txtReceiver_Change()
   m_HasModify = True
End Sub

Private Sub txtSender_Change()
   m_HasModify = True
End Sub

Private Sub txtTotal_Change()
   m_HasModify = True
End Sub

Private Sub txtTruckNo_Change()
   m_HasModify = True
End Sub

Private Sub uctlSetupDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlTextBox1_Change()
   m_HasModify = True
End Sub

Private Sub uctlDeliveryLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlDocumentDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlEntryTime_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlExitTime_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlSupplierLookup_Change()
   m_HasModify = True
End Sub
