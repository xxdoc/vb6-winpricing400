VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddEditPigDoc3 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditPigDoc3.frx":0000
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
      TabIndex        =   11
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextLookup uctlEmployeeLookup 
         Height          =   435
         Left            =   2280
         TabIndex        =   3
         Top             =   1950
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlDocumentDate 
         Height          =   405
         Left            =   2280
         TabIndex        =   2
         Top             =   1530
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   4
         Top             =   2760
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
         Left            =   2280
         TabIndex        =   0
         Top             =   1080
         Width           =   2655
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
         Height          =   4425
         Left            =   150
         TabIndex        =   5
         Top             =   3300
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   7805
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
         Column(1)       =   "frmAddEditPigDoc3.frx":27A2
         Column(2)       =   "frmAddEditPigDoc3.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditPigDoc3.frx":290E
         FormatStyle(2)  =   "frmAddEditPigDoc3.frx":2A6A
         FormatStyle(3)  =   "frmAddEditPigDoc3.frx":2B1A
         FormatStyle(4)  =   "frmAddEditPigDoc3.frx":2BCE
         FormatStyle(5)  =   "frmAddEditPigDoc3.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditPigDoc3.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin Threed.SSCheck chkCommit 
         Height          =   435
         Left            =   4980
         TabIndex        =   1
         Top             =   1080
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblEmployeeNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   750
         TabIndex        =   16
         Top             =   2010
         Width           =   1455
      End
      Begin VB.Label Label4 
         Height          =   315
         Left            =   11250
         TabIndex        =   15
         Top             =   3420
         Width           =   585
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   990
         TabIndex        =   14
         Top             =   1560
         Width           =   1155
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   9
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPigDoc3.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   10
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
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPigDoc3.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   8
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPigDoc3.frx":356A
         ButtonStyle     =   3
      End
      Begin VB.Label lblDocumentNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   1140
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmAddEditPigDoc3"
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
Public ID As Long

Private FileName As String
Private m_SumUnit As Double

Private Sub SplitImprtExport()
Dim O As Object

   For Each O In m_InventoryDoc.ImportExports
      If O.TX_TYPE = "I" Then
         Call m_InventoryDoc.ImportItems.Add(O)
      ElseIf O.TX_TYPE = "E" Then
         Call m_InventoryDoc.ExportItems.Add(O)
      End If
   Next O
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
      m_InventoryDoc.INVENTORY_DOC_ID = ID
      m_InventoryDoc.COMMIT_FLAG = ""
      m_InventoryDoc.QueryFlag = 1
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
      uctlEmployeeLookup.MyCombo.ListIndex = IDToListIndex(uctlEmployeeLookup.MyCombo, m_InventoryDoc.EMP_ID)
      chkCommit.Value = FlagToCheck(m_InventoryDoc.COMMIT_FLAG)
      
      cmdAdd.Enabled = (m_InventoryDoc.COMMIT_FLAG = "N")
      cmdDelete.Enabled = (m_InventoryDoc.COMMIT_FLAG = "N")
      chkCommit.Enabled = (m_InventoryDoc.COMMIT_FLAG = "N")
      
      Call SplitImprtExport
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

Private Function GetNextID() As Long
Dim Ei As CExportItem
Dim Ii As CLotItem
Dim Max As Long

   Max = 0
   For Each Ei In m_InventoryDoc.ExportItems
      If Ei.TRANSACTION_SEQ > Max Then
         Max = Ei.TRANSACTION_SEQ
      End If
   Next Ei
   
   For Each Ii In m_InventoryDoc.ImportItems
      If Ii.TRANSACTION_SEQ > Max Then
         Max = Ii.TRANSACTION_SEQ
      End If
   Next Ii
   
   GetNextID = Max + 1
End Function

Private Function SaveData() As Boolean
Dim IsOK As Boolean

   If ShowMode = SHOW_ADD Then
'      If Not VerifyAccessRight("DAILY_CUSTOMER_ADD") Then
'         Call EnableForm(Me, True)
'         Exit Function
'      End If
   ElseIf ShowMode = SHOW_EDIT Then
'      If Not VerifyAccessRight("DAILY_CUSTOMER_EDIT") Then
'         Call EnableForm(Me, True)
'         Exit Function
'      End If
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
   
   If Not CheckUniqueNs(EXPORT_UNIQUE, txtDocumentNo.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("�բ�����") & " " & txtDocumentNo.Text & " " & MapText("������к�����")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_InventoryDoc.AddEditMode = ShowMode
   m_InventoryDoc.INVENTORY_DOC_ID = ID
    m_InventoryDoc.DOCUMENT_DATE = uctlDocumentDate.ShowDate
   m_InventoryDoc.DOCUMENT_NO = txtDocumentNo.Text
   m_InventoryDoc.EMP_ID = uctlEmployeeLookup.MyCombo.ItemData(Minus2Zero(uctlEmployeeLookup.MyCombo.ListIndex))
   m_InventoryDoc.DOCUMENT_TYPE = 9
   m_InventoryDoc.COMMIT_FLAG = Check2Flag(chkCommit.Value)
   
   Call EnableForm(Me, False)
   
   Set m_InventoryDoc.ImportExports = Nothing
   Set m_InventoryDoc.ImportExports = New Collection
   
   Call glbDaily.MergeImportExport(m_InventoryDoc.ImportItems, m_InventoryDoc.ExportItems, m_InventoryDoc.ImportExports)
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

Private Sub CreateImportExportItems()
Dim Ti As CTransferItem
Dim Ei As CExportItem
Dim Ii As CLotItem

   Set m_InventoryDoc.ImportExports = Nothing
   Set m_InventoryDoc.ImportExports = New Collection
   
   For Each Ti In m_InventoryDoc.TransferItems
      Set Ei = Ti.ExportItem
      Set Ii = Ti.ImportItem
      
      Ei.Flag = Ti.Flag
      Ii.Flag = Ti.Flag
      
      Call m_InventoryDoc.ImportExports.Add(Ei)
      Call m_InventoryDoc.ImportExports.Add(Ii)
   Next Ti
End Sub

Private Sub CreateTransferItems()
Dim Ti As CTransferItem
Dim O As Object
Dim Ei As CExportItem
Dim Ii As CLotItem
Dim i As Long
Dim j As Long
Dim Count1 As Long
Dim Count2 As Long
Dim ImportCount As Long
Dim ExportCount As Long

   Count1 = m_InventoryDoc.ImportExports.Count \ 2
   Count2 = m_InventoryDoc.ImportExports.Count
   For i = 1 To Count1
      ImportCount = 0
      ExportCount = 0
      j = 1
      While j <= Count2
         Set O = m_InventoryDoc.ImportExports(j)
         If (O.TX_TYPE = "I") Then
            ImportCount = ImportCount + 1
            If ImportCount = i Then
               Set Ii = O
            End If
         ElseIf (O.TX_TYPE = "E") Then
            ExportCount = ExportCount + 1
            If ExportCount = i Then
               Set Ei = O
            End If
         End If
         j = j + 1
      Wend
         
      Set Ti = New CTransferItem
      Set Ti.ImportItem = Ii
      Set Ti.ExportItem = Ei
      Ti.Flag = Ii.Flag
      Call m_InventoryDoc.TransferItems.Add(Ti)
      Set Ti = Nothing
   Next i
End Sub

Private Sub chkCommit_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean

   If Not cmdAdd.Enabled Then
      Exit Sub
   End If
   
   OKClick = False
   If TabStrip1.SelectedItem.Index = 1 Then
      frmAddEditPigAdjustItem1.COMMIT_FLAG = m_InventoryDoc.COMMIT_FLAG
      frmAddEditPigAdjustItem1.TxSeq = GetNextID
      Set frmAddEditPigAdjustItem1.TempCollection = m_InventoryDoc.ImportItems
      frmAddEditPigAdjustItem1.ParentShowMode = ShowMode
      frmAddEditPigAdjustItem1.ShowMode = SHOW_ADD
      frmAddEditPigAdjustItem1.HeaderText = MapText("������¡�û�Ѻ�ʹ (����)")
      Load frmAddEditPigAdjustItem1
      frmAddEditPigAdjustItem1.Show 1

      OKClick = frmAddEditPigAdjustItem1.OKClick

      Unload frmAddEditPigAdjustItem1
      Set frmAddEditPigAdjustItem1 = Nothing

      If OKClick Then
         Call GetTotalPrice
         
         GridEX1.ItemCount = CountItem(m_InventoryDoc.ImportItems)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      frmAddEditPigAdjustItem2.COMMIT_FLAG = m_InventoryDoc.COMMIT_FLAG
      frmAddEditPigAdjustItem2.TxSeq = GetNextID
      Set frmAddEditPigAdjustItem2.TempCollection = m_InventoryDoc.ExportItems
      frmAddEditPigAdjustItem2.ParentShowMode = ShowMode
      frmAddEditPigAdjustItem2.ShowMode = SHOW_ADD
      frmAddEditPigAdjustItem2.HeaderText = MapText("������¡�û�Ѻ�ʹ (Ŵ)")
      Load frmAddEditPigAdjustItem2
      frmAddEditPigAdjustItem2.Show 1

      OKClick = frmAddEditPigAdjustItem2.OKClick

      Unload frmAddEditPigAdjustItem2
      Set frmAddEditPigAdjustItem2 = Nothing

      If OKClick Then
         Call GetTotalPrice
         
         GridEX1.ItemCount = CountItem(m_InventoryDoc.ExportItems)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
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
   
   If TabStrip1.SelectedItem.Index = 1 Then
      If ID1 <= 0 Then
         m_InventoryDoc.ImportItems.Remove (ID2)
      Else
         m_InventoryDoc.ImportItems.Item(ID2).Flag = "D"
      End If

      Call GetTotalPrice
      GridEX1.ItemCount = CountItem(m_InventoryDoc.ImportItems)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      If ID1 <= 0 Then
         m_InventoryDoc.ExportItems.Remove (ID2)
      Else
         m_InventoryDoc.ExportItems.Item(ID2).Flag = "D"
      End If

      Call GetTotalPrice
      GridEX1.ItemCount = CountItem(m_InventoryDoc.ExportItems)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean

'   If Not VerifyAccessRight("GROUP_QUERY_RIGHT") Then
'      Exit Sub
'   End If
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(2))
   OKClick = False
   
   If TabStrip1.SelectedItem.Index = 1 Then
      frmAddEditPigAdjustItem1.ID = ID
      frmAddEditPigAdjustItem1.COMMIT_FLAG = m_InventoryDoc.COMMIT_FLAG
      Set frmAddEditPigAdjustItem1.TempCollection = m_InventoryDoc.ImportItems
      frmAddEditPigAdjustItem1.HeaderText = MapText("�����¡�û�Ѻ�ʹ (����)")
      frmAddEditPigAdjustItem1.ParentShowMode = ShowMode
      frmAddEditPigAdjustItem1.ShowMode = SHOW_EDIT
      Load frmAddEditPigAdjustItem1
      frmAddEditPigAdjustItem1.Show 1

      OKClick = frmAddEditPigAdjustItem1.OKClick

      Unload frmAddEditPigAdjustItem1
      Set frmAddEditPigAdjustItem1 = Nothing

      If OKClick Then
         Call GetTotalPrice
         GridEX1.ItemCount = CountItem(m_InventoryDoc.ImportItems)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      frmAddEditPigAdjustItem2.ID = ID
      frmAddEditPigAdjustItem2.COMMIT_FLAG = m_InventoryDoc.COMMIT_FLAG
      Set frmAddEditPigAdjustItem2.TempCollection = m_InventoryDoc.ExportItems
      frmAddEditPigAdjustItem2.HeaderText = MapText("�����¡�û�Ѻ�ʹ (Ŵ)")
      frmAddEditPigAdjustItem2.ParentShowMode = ShowMode
      frmAddEditPigAdjustItem2.ShowMode = SHOW_EDIT
      Load frmAddEditPigAdjustItem2
      frmAddEditPigAdjustItem2.Show 1

      OKClick = frmAddEditPigAdjustItem2.OKClick

      Unload frmAddEditPigAdjustItem2
      Set frmAddEditPigAdjustItem2 = Nothing

      If OKClick Then
         Call GetTotalPrice
         GridEX1.ItemCount = CountItem(m_InventoryDoc.ExportItems)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
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
         m_InventoryDoc.QueryFlag = 0
         Call QueryData(False)
      End If
      
      Call EnableForm(Me, True)
      m_HasModify = False
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
   ElseIf Shift = 0 And KeyCode = 116 Then
'      Call cmdSearch_Click
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
   ElseIf Shift = 0 And KeyCode = 118 Then
      Call cmdAdd_Click
   ElseIf Shift = 0 And KeyCode = 117 Then
      Call cmdDelete_Click
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
   ElseIf Shift = 0 And KeyCode = 114 Then
      Call cmdEdit_Click
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
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
   Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid1()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.Name = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.Add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   Set Col = GridEX1.Columns.Add '3
   Col.Width = 2100
   Col.Caption = MapText("�ѻ�����Դ")

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 4425 + 3240
   Col.Caption = MapText("��������´�ء�")

   Set Col = GridEX1.Columns.Add '5
   Col.Width = 1785
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("�ӹǹ")

   Set Col = GridEX1.Columns.Add '8
   Col.Width = 1980
   Col.Caption = MapText("�ç���͹")
End Sub

Private Sub GetTotalPrice()
Dim Ii As CTransferItem
Dim Sum As Double

   Sum = 0
   For Each Ii In m_InventoryDoc.TransferItems
      If Ii.Flag <> "D" Then
         Sum = Sum + CDbl(Format(Ii.ExportItem.EXPORT_AVG_PRICE, "0.00")) * CDbl(Format(Ii.ExportItem.EXPORT_AMOUNT, "0.00"))
      End If
   Next Ii
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblDocumentNo, MapText("�����Ţ㺻�Ѻ�ʹ"))
   Call InitNormalLabel(lblDocumentDate, MapText("�ѹ����͡���"))
   Call InitNormalLabel(Label4, MapText("�ҷ"))
   Call InitNormalLabel(lblEmployeeNo, MapText("����Ѻ�Դ�ͺ"))
   
   Call txtDocumentNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitCheckBox(chkCommit, MapText("�ӹǳ"))
   
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
   Call InitMainButton(cmdAdd, MapText("���� (F7)"))
   Call InitMainButton(cmdEdit, MapText("��� (F3)"))
   Call InitMainButton(cmdDelete, MapText("ź (F6)"))
   
   Call InitGrid1
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.Name = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.Add().Caption = MapText("��Ѻ�ʹ�ء� (����)")
   TabStrip1.Tabs.Add().Caption = MapText("��Ѻ�ʹ�ء� (Ŵ)")
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

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Index = 1 Then
      If m_InventoryDoc.ImportItems Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim CR As CLotItem
      If m_InventoryDoc.ImportItems.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_InventoryDoc.ImportItems, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If

      Values(1) = CR.LOT_ITEM_ID
      Values(2) = RealIndex
      Values(3) = CR.PART_NO
      Values(4) = CR.PART_DESC
      Values(5) = FormatNumber(CR.TX_AMOUNT)
      Values(6) = CR.LOCATION_NAME
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      If m_InventoryDoc.ExportItems Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim Ei As CExportItem
      If m_InventoryDoc.ExportItems.Count <= 0 Then
         Exit Sub
      End If
      Set Ei = GetItem(m_InventoryDoc.ExportItems, RowIndex, RealIndex)
      If Ei Is Nothing Then
         Exit Sub
      End If

      Values(1) = Ei.EXPORT_ITEM_ID
      Values(2) = RealIndex
      Values(3) = Ei.PART_NO
      Values(4) = Ei.PART_DESC
      Values(5) = FormatNumber(Ei.EXPORT_AMOUNT)
      Values(6) = Ei.LOCATION_NAME
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
      GridEX1.ItemCount = CountItem(m_InventoryDoc.ImportItems)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      Call InitGrid1
      
      Call GetTotalPrice
      GridEX1.ItemCount = CountItem(m_InventoryDoc.ExportItems)
      GridEX1.Rebind
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

Private Sub uctlDocumentDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlEmployeeLookup_Change()
   m_HasModify = True
End Sub
