VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAddEditInventoryDoc2 
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmaddeditinventorydoc2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9315
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   9360
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   16510
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextLookup uctlEmployeeLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   3
         Top             =   1530
         Width           =   5385
         _extentx        =   9499
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlDate uctlDocumentDate 
         Height          =   405
         Left            =   6300
         TabIndex        =   2
         Top             =   1110
         Width           =   3855
         _extentx        =   6800
         _extenty        =   714
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   8
         Top             =   3900
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
      Begin prjFarmManagement.uctlTextBox txtDocumentNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   1080
         Width           =   2385
         _extentx        =   5001
         _extenty        =   767
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   4125
         Left            =   150
         TabIndex        =   9
         Top             =   4440
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   7276
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
         Column(1)       =   "frmaddeditinventorydoc2.frx":27A2
         Column(2)       =   "frmaddeditinventorydoc2.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmaddeditinventorydoc2.frx":290E
         FormatStyle(2)  =   "frmaddeditinventorydoc2.frx":2A6A
         FormatStyle(3)  =   "frmaddeditinventorydoc2.frx":2B1A
         FormatStyle(4)  =   "frmaddeditinventorydoc2.frx":2BCE
         FormatStyle(5)  =   "frmaddeditinventorydoc2.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmaddeditinventorydoc2.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtDeliveryFee 
         Height          =   435
         Left            =   1860
         TabIndex        =   5
         Top             =   2460
         Width           =   1485
         _extentx        =   2619
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtNote 
         Height          =   435
         Left            =   1860
         TabIndex        =   7
         Top             =   2910
         Width           =   8325
         _extentx        =   5001
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlReceivedByLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   25
         Top             =   2000
         Width           =   5385
         _extentx        =   9499
         _extenty        =   767
      End
      Begin VB.Label lblReceivedBy 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   360
         TabIndex        =   26
         Top             =   2000
         Width           =   1455
      End
      Begin VB.Label lblNote 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   390
         TabIndex        =   24
         Top             =   2970
         Width           =   1365
      End
      Begin Threed.SSCommand cmdAuto 
         Height          =   405
         Left            =   4260
         TabIndex        =   1
         Top             =   1080
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmaddeditinventorydoc2.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   525
         Left            =   6840
         TabIndex        =   13
         Top             =   8670
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmaddeditinventorydoc2.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCheck chkSaleFlag 
         Height          =   435
         Left            =   7320
         TabIndex        =   4
         Top             =   2490
         Visible         =   0   'False
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin Threed.SSCheck chkCommit 
         Height          =   435
         Left            =   7320
         TabIndex        =   6
         Top             =   1530
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblEmployeeNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   330
         TabIndex        =   23
         Top             =   1590
         Width           =   1455
      End
      Begin VB.Label Label4 
         Height          =   315
         Left            =   11250
         TabIndex        =   22
         Top             =   3420
         Width           =   585
      End
      Begin VB.Label Label1 
         Height          =   315
         Left            =   3450
         TabIndex        =   21
         Top             =   2070
         Width           =   765
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5010
         TabIndex        =   20
         Top             =   1140
         Width           =   1155
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   14
         Top             =   8670
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmaddeditinventorydoc2.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   15
         Top             =   8670
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   11
         Top             =   8670
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   10
         Top             =   8670
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmaddeditinventorydoc2.frx":3884
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   12
         Top             =   8670
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmaddeditinventorydoc2.frx":3B9E
         ButtonStyle     =   3
      End
      Begin VB.Label lblDeliveryFee 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         TabIndex        =   18
         Top             =   2580
         Width           =   1695
      End
      Begin VB.Label lblDocumentNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   90
         TabIndex        =   17
         Top             =   1140
         Width           =   1665
      End
   End
End
Attribute VB_Name = "frmAddEditInventoryDoc2"
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
Private m_Employees As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public id As Long

Private FileName As String
Private m_SumUnit As Double

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
      txtDocumentNo.Text = m_InventoryDoc.DOCUMENT_NO
      txtDeliveryFee.Text = Format(m_InventoryDoc.DELIVERY_FEE, "0.00")
      uctlEmployeeLookup.MyCombo.ListIndex = IDToListIndex(uctlEmployeeLookup.MyCombo, m_InventoryDoc.EMP_ID)
      chkCommit.Value = FlagToCheck(m_InventoryDoc.COMMIT_FLAG)
      chkSaleFlag.Value = FlagToCheck(m_InventoryDoc.SALE_FLAG)
      txtNote.Text = m_InventoryDoc.DOCUMENT_DESC
      
      cmdAdd.Enabled = (m_InventoryDoc.COMMIT_FLAG = "N")
      cmdDelete.Enabled = (m_InventoryDoc.COMMIT_FLAG = "N")
      chkCommit.Enabled = (m_InventoryDoc.COMMIT_FLAG = "N")
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
      If Not VerifyAccessRight("INVENTORY_EXPORT_EDIT") Then
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
   If Not VerifyCombo(lblEmployeeNo, uctlEmployeeLookup.MyCombo, True) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblDeliveryFee, txtDeliveryFee, True) Then
      Exit Function
   End If
   
   If Not CheckUniqueNs(EXPORT_UNIQUE, txtDocumentNo.Text, id) Then
      glbErrorLog.LocalErrorMsg = MapText("�բ�����") & " " & txtDocumentNo.Text & " " & MapText("������к�����")
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
   m_InventoryDoc.DOCUMENT_NO = txtDocumentNo.Text
   m_InventoryDoc.DELIVERY_FEE = 0
   m_InventoryDoc.EMP_ID = uctlEmployeeLookup.MyCombo.ItemData(Minus2Zero(uctlEmployeeLookup.MyCombo.ListIndex))
   m_InventoryDoc.DOCUMENT_TYPE = 2
   m_InventoryDoc.COMMIT_FLAG = Check2Flag(chkCommit.Value)
   m_InventoryDoc.SALE_FLAG = Check2Flag(chkSaleFlag.Value)
   m_InventoryDoc.EXCEPTION_FLAG = "Y"
   m_InventoryDoc.DOCUMENT_DESC = txtNote.Text
   
   Call EnableForm(Me, False)
   
   If (m_InventoryDoc.COMMIT_FLAG = "Y") Then
      If m_InventoryDoc.OLD_COMMIT_FLAG <> "Y" Then
         Call glbDaily.TriggerCommit(m_InventoryDoc.ImportExports)
         If Not glbDaily.VerifyStockBalance(m_InventoryDoc.ImportExports, glbErrorLog) Then
            Call EnableForm(Me, True)
            Exit Function
         End If
      End If
   End If
   
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

Private Sub cboReason_Click()
   m_HasModify = True
End Sub

Private Sub chkCommit_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkCommit_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub chkSaleFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub GetTotalPrice()
Dim D As CLotItem
Dim Sum As Double

   Sum = 0
   For Each D In m_InventoryDoc.ImportExports
      Sum = Sum + D.INCLUDE_UNIT_PRICE * D.TX_AMOUNT
   Next D
   txtDeliveryFee.Text = FormatNumber(Sum)
End Sub

Private Sub chkSaleFlag_KeyPress(KeyAscii As Integer)
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
      frmAddEditExportItem.COMMIT_FLAG = m_InventoryDoc.COMMIT_FLAG
      Set frmAddEditExportItem.TempCollection = m_InventoryDoc.ImportExports
      frmAddEditExportItem.ParentShowMode = ShowMode
      frmAddEditExportItem.ShowMode = SHOW_ADD
      frmAddEditExportItem.HeaderText = MapText("������¡���ԡ����")
      Load frmAddEditExportItem
      frmAddEditExportItem.Show 1

      OKClick = frmAddEditExportItem.OKClick

      Unload frmAddEditExportItem
      Set frmAddEditExportItem = Nothing

      If OKClick Then
         Call GetTotalPrice
         
         GridEX1.ItemCount = CountItem(m_InventoryDoc.ImportExports)
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

Private Sub cmdAuto_Click()
Dim No As String

   If Trim(txtDocumentNo.Text) = "" Then
      Call glbDatabaseMngr.GenerateNumber(EXPORT_NUMBER, No, glbErrorLog)
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
      If ID1 <= 0 Then
         m_InventoryDoc.ImportExports.Remove (ID2)
      Else
         m_InventoryDoc.ImportExports.Item(ID2).Flag = "D"
      End If

      Call GetTotalPrice
      GridEX1.ItemCount = CountItem(m_InventoryDoc.ImportExports)
      GridEX1.Rebind
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
      frmAddEditExportItem.id = id
      frmAddEditExportItem.COMMIT_FLAG = m_InventoryDoc.COMMIT_FLAG
      Set frmAddEditExportItem.TempCollection = m_InventoryDoc.ImportExports
      frmAddEditExportItem.HeaderText = MapText("�����¡���ԡ����")
      frmAddEditExportItem.ParentShowMode = ShowMode
      frmAddEditExportItem.ShowMode = SHOW_EDIT
      Load frmAddEditExportItem
      frmAddEditExportItem.Show 1

      OKClick = frmAddEditExportItem.OKClick

      Unload frmAddEditExportItem
      Set frmAddEditExportItem = Nothing

      If OKClick Then
         Call GetTotalPrice
         GridEX1.ItemCount = CountItem(m_InventoryDoc.ImportExports)
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

Private Sub cmdOK_Click()
Dim oMenu As cPopupMenu
Dim lMenuChosen  As Long

   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("�ѹ�֡", "-", "�ѹ�֡����͡�ҡ˹�Ҩ�")
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
      glbErrorLog.LocalErrorMsg = "��سҷӡ�úѹ�֡������������º���¡�͹"
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("��ԡ�Թ���/�ѵ�شԺ", "��Ѻ���˹�ҡ�д��")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   
   If lMenuChosen = 1 Then
      ReportKey = "CReportInvDoc002"
      ClassName = "CReportInvDoc002"
      Set Report = New CReportInvDoc002
      ReportFlag = True
   ElseIf lMenuChosen = 2 Then
      ReportKey = "CReportInvDoc002"
      Set Rc = New CReportConfig
      Rc.REPORT_KEY = ReportKey
      Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
      Call Rc.QueryData(m_Rs, iCount)
      HeaderText = MapText("��ԡ�Թ���/�ѵ�شԺ")
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

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      
      Call LoadEmployee(uctlEmployeeLookup.MyCombo, m_Employees)
      Set uctlEmployeeLookup.MyCollection = m_Employees
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_InventoryDoc.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         uctlDocumentDate.ShowDate = Now
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
   Set m_Employees = Nothing
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
   Col.Width = 1710
   Col.Caption = MapText("�����ѵ�شԺ")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 4335
   Col.Caption = MapText("�ѵ�شԺ")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 3360
   Col.Caption = MapText("��������´")

   Set Col = GridEX1.Columns.add '6
   Col.Width = 1785
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("����ҳ")

   Set Col = GridEX1.Columns.add '7
   Col.Width = 1620
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("�Ҥ�")
   
   Set Col = GridEX1.Columns.add '8
   Col.Width = 1620
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("�Ҥ����")
   
   Set Col = GridEX1.Columns.add '9
   Col.Width = 1440
   Col.Caption = MapText("ʶҹ���Ѵ��")
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblDocumentNo, MapText("�����Ţ��ԡ"))
   Call InitNormalLabel(lblDeliveryFee, MapText("��Ť�����"))
   Call InitNormalLabel(lblDocumentDate, MapText("�ѹ����͡���"))
   Call InitNormalLabel(Label1, MapText("�ҷ"))
   Call InitNormalLabel(Label4, MapText("�ҷ"))
   Call InitNormalLabel(lblEmployeeNo, MapText("����Ѻ�Դ�ͺ"))
   Call InitNormalLabel(lblNote, MapText("�����˵�"))
   
   Call txtDocumentNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtDeliveryFee.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtDeliveryFee.Enabled = False
   Call txtNote.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPrint.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAuto.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
   Call InitMainButton(cmdAdd, MapText("���� (F7)"))
   Call InitMainButton(cmdEdit, MapText("��� (F3)"))
   Call InitMainButton(cmdDelete, MapText("ź (F6)"))
   Call InitMainButton(cmdPrint, MapText("�����"))
   Call InitMainButton(cmdAuto, MapText("A"))

   Call InitCheckBox(chkCommit, MapText("�ӹǳ"))
   Call InitCheckBox(chkSaleFlag, MapText("�ԡ���"))
   
   Call InitGrid1
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.NAME = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.add().Caption = MapText("��¡���ԡ�ѵ�شԺ")
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
   Set m_Employees = New Collection
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

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
      Values(5) = CR.ITEM_DESC_NAME & " " & CR.ITEM_DESC
      Values(6) = FormatNumber(CR.TX_AMOUNT)
      Values(7) = FormatNumber(CR.INCLUDE_UNIT_PRICE)
      Values(8) = FormatNumber(CR.TX_AMOUNT * CR.INCLUDE_UNIT_PRICE)
      Values(9) = CR.LOCATION_NAME
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

Private Sub SSCommand1_Click()

End Sub

Private Sub TabStrip1_Click()
   If TabStrip1.SelectedItem.Index = 1 Then
      Call InitGrid1
      
      Call GetTotalPrice
      GridEX1.ItemCount = CountItem(m_InventoryDoc.ImportExports)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
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

Private Sub txtNote_Change()
   m_HasModify = True
End Sub

Private Sub uctlDocumentDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlEmployeeLookup_Change()
   m_HasModify = True
End Sub
Private Sub GridEX1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
Dim TempID1 As Long
Dim BD As CLotItem
Dim IsOK As Boolean
Dim OKClick As Boolean
Dim TempBd As CLotItem
   
   If GridEX1.ItemCount <= 0 Then
         Exit Sub
   End If
   
   TempID1 = Val(GridEX1.Value(2))
   If Button = 2 Then
      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.Popup("�Ѵ�͡������")
      If lMenuChosen = 0 Then
         Exit Sub
      End If
      Set oMenu = Nothing
   Else
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   If lMenuChosen = 1 Then
      Set BD = m_InventoryDoc.ImportExports(TempID1)
      Set TempBd = New CLotItem
      TempBd.Flag = "A"
      TempBd.PART_TYPE = BD.PART_TYPE
      TempBd.PART_ITEM_ID = BD.PART_ITEM_ID
      TempBd.LOCATION_ID = BD.LOCATION_ID
      TempBd.TX_AMOUNT = BD.TX_AMOUNT
      TempBd.INCLUDE_UNIT_PRICE = BD.INCLUDE_UNIT_PRICE
      TempBd.ACTUAL_UNIT_PRICE = BD.ACTUAL_UNIT_PRICE
      TempBd.PART_TYPE_NAME = BD.PART_TYPE_NAME
      TempBd.LOCATION_NAME = BD.LOCATION_NAME
      TempBd.PART_NO = BD.PART_NO
      TempBd.PART_DESC = BD.PART_DESC
      TempBd.TO_DEPARTMENT = BD.TO_DEPARTMENT
      TempBd.ITEM_DESC = BD.ITEM_DESC
      TempBd.CALCULATE_FLAG = BD.CALCULATE_FLAG
      TempBd.TX_TYPE = BD.TX_TYPE
      TempBd.EXPENSE_TYPE = BD.EXPENSE_TYPE
      TempBd.PIG_FLAG = BD.PIG_FLAG
      
      TempBd.ITEM_DESC_ID = BD.ITEM_DESC_ID
      TempBd.ITEM_DESC_NAME = BD.ITEM_DESC_NAME
      
      Call m_InventoryDoc.ImportExports.add(TempBd)
      
      Set BD = Nothing
      Set TempBd = Nothing
      
      m_HasModify = True
      Call TabStrip1_Click
   End If
   Call EnableForm(Me, True)
End Sub
