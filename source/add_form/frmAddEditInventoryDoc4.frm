VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAddEditInventoryDoc4 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditInventoryDoc4.frx":0000
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
      TabIndex        =   13
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboReason 
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   9
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2400
         Width           =   4005
      End
      Begin prjFarmManagement.uctlTextLookup uctlEmployeeLookup 
         Height          =   435
         Left            =   2280
         TabIndex        =   2
         Top             =   1950
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlDocumentDate 
         Height          =   405
         Left            =   2280
         TabIndex        =   1
         Top             =   1530
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   6
         Top             =   3060
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
         Height          =   4125
         Left            =   150
         TabIndex        =   7
         Top             =   3600
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
         Column(1)       =   "frmAddEditInventoryDoc4.frx":27A2
         Column(2)       =   "frmAddEditInventoryDoc4.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditInventoryDoc4.frx":290E
         FormatStyle(2)  =   "frmAddEditInventoryDoc4.frx":2A6A
         FormatStyle(3)  =   "frmAddEditInventoryDoc4.frx":2B1A
         FormatStyle(4)  =   "frmAddEditInventoryDoc4.frx":2BCE
         FormatStyle(5)  =   "frmAddEditInventoryDoc4.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditInventoryDoc4.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtTotalAmount 
         Height          =   435
         Left            =   9030
         TabIndex        =   5
         Top             =   2400
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   767
      End
      Begin VB.Label lblTotalAmount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6870
         TabIndex        =   21
         Top             =   2460
         Width           =   2055
      End
      Begin Threed.SSCheck chkAdjustFlag 
         Height          =   435
         Left            =   7740
         TabIndex        =   20
         Top             =   0
         Visible         =   0   'False
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblReason 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   180
         TabIndex        =   19
         Top             =   2460
         Width           =   1965
      End
      Begin Threed.SSCheck chkCommit 
         Height          =   435
         Left            =   7710
         TabIndex        =   3
         Top             =   1950
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
         TabIndex        =   18
         Top             =   2010
         Width           =   1455
      End
      Begin VB.Label Label4 
         Height          =   315
         Left            =   11250
         TabIndex        =   17
         Top             =   3420
         Width           =   585
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   990
         TabIndex        =   16
         Top             =   1560
         Width           =   1155
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   11
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditInventoryDoc4.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   12
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
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditInventoryDoc4.frx":3250
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
         MouseIcon       =   "frmAddEditInventoryDoc4.frx":356A
         ButtonStyle     =   3
      End
      Begin VB.Label lblDocumentNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   1140
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmAddEditInventoryDoc4"
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
Public DocumentType As Long

Private FileName As String
Private m_SumUnit As Double

Private Sub SplitImprtExport()
Dim O As Object

   Set m_InventoryDoc.ImportItems = Nothing
   Set m_InventoryDoc.ImportItems = New Collection
   
   Set m_InventoryDoc.ExportItems = Nothing
   Set m_InventoryDoc.ExportItems = New Collection
   
   For Each O In m_InventoryDoc.ImportExports
      If O.TX_TYPE = "I" Then
         Call m_InventoryDoc.ImportItems.add(O)
      ElseIf O.TX_TYPE = "E" Then
         Call m_InventoryDoc.ExportItems.add(O)
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
      cboReason.ListIndex = IDToListIndex(cboReason, m_InventoryDoc.REASON_ID)
      
      chkCommit.Value = FlagToCheck(m_InventoryDoc.COMMIT_FLAG)
      chkAdjustFlag.Value = FlagToCheck(m_InventoryDoc.ADJUST_FLAG)
      
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

'Private Function GetNextID() As Long
'Dim Ei As CExportItem
'Dim Ii As CLotItem
'Dim Max As Long
'
'   Max = 0
'   For Each Ei In m_InventoryDoc.ExportItems
'      If Ei.TRANSACTION_SEQ > Max Then
'         Max = Ei.TRANSACTION_SEQ
'      End If
'   Next Ei
'
'   For Each Ii In m_InventoryDoc.ImportItems
'      If Ii.TRANSACTION_SEQ > Max Then
'         Max = Ii.TRANSACTION_SEQ
'      End If
'   Next Ii
'
'   GetNextID = Max + 1
'End Function

Private Function SaveData() As Boolean
Dim IsOK As Boolean
   If ShowMode = SHOW_EDIT Then
      If Not VerifyAccessRight("INVENTORY_ADJUST_EDIT") Then
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
   m_InventoryDoc.DELIVERY_FEE = 0
   m_InventoryDoc.EMP_ID = uctlEmployeeLookup.MyCombo.ItemData(Minus2Zero(uctlEmployeeLookup.MyCombo.ListIndex))
   m_InventoryDoc.DOCUMENT_TYPE = DocumentType
   m_InventoryDoc.COMMIT_FLAG = Check2Flag(chkCommit.Value)
   m_InventoryDoc.REASON_ID = cboReason.ItemData(Minus2Zero(cboReason.ListIndex))
   m_InventoryDoc.EXCEPTION_FLAG = "Y"
   m_InventoryDoc.ADJUST_FLAG = "Y"
   
   Call EnableForm(Me, False)
   
'   Set m_InventoryDoc.ImportExports = Nothing
'   Set m_InventoryDoc.ImportExports = New Collection
   
'   Call glbDaily.MergeImportExport(m_InventoryDoc.ImportItems, m_InventoryDoc.ExportItems, m_InventoryDoc.ImportExports)
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
'Dim Ti As CTransferItem
'Dim Ei As CExportItem
'Dim Ii As CLotItem
'
'   Set m_InventoryDoc.ImportExports = Nothing
'   Set m_InventoryDoc.ImportExports = New Collection
'
'   For Each Ti In m_InventoryDoc.TransferItems
'      Set Ei = Ti.ExportItem
'      Set Ii = Ti.ImportItem
'
'      Ei.Flag = Ti.Flag
'      Ii.Flag = Ti.Flag
'
'      Call m_InventoryDoc.ImportExports.Add(Ei)
'      Call m_InventoryDoc.ImportExports.Add(Ii)
'   Next Ti
End Sub

Private Sub CreateTransferItems()
'Dim Ti As CTransferItem
'Dim O As Object
'Dim Ei As CExportItem
'Dim Ii As CLotItem
'Dim i As Long
'Dim j As Long
'Dim Count1 As Long
'Dim Count2 As Long
'Dim ImportCount As Long
'Dim ExportCount As Long
'
'   Count1 = m_InventoryDoc.ImportExports.Count \ 2
'   Count2 = m_InventoryDoc.ImportExports.Count
'   For i = 1 To Count1
'      ImportCount = 0
'      ExportCount = 0
'      j = 1
'      While j <= Count2
'         Set O = m_InventoryDoc.ImportExports(j)
'         If (O.TX_TYPE = "I") Then
'            ImportCount = ImportCount + 1
'            If ImportCount = i Then
'               Set Ii = O
'            End If
'         ElseIf (O.TX_TYPE = "E") Then
'            ExportCount = ExportCount + 1
'            If ExportCount = i Then
'               Set Ei = O
'            End If
'         End If
'         j = j + 1
'      Wend
'
'      Set Ti = New CTransferItem
'      Set Ti.ImportItem = Ii
'      Set Ti.ExportItem = Ei
'      Ti.Flag = Ii.Flag
'      Call m_InventoryDoc.TransferItems.Add(Ti)
'      Set Ti = Nothing
'   Next i
End Sub

Private Sub cboReason_Click()
   m_HasModify = True
End Sub

Private Sub chkAdjustFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkCommit_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Function GetNextID() As Long
Dim Ei As CLotItem
Dim II As CLotItem
Dim MAX As Long

   MAX = 0
   For Each Ei In m_InventoryDoc.ExportItems
      If Ei.TRANSACTION_SEQ > MAX Then
         MAX = Ei.TRANSACTION_SEQ
      End If
   Next Ei
   
   For Each II In m_InventoryDoc.ImportItems
      If II.TRANSACTION_SEQ > MAX Then
         MAX = II.TRANSACTION_SEQ
      End If
   Next II
   
   GetNextID = MAX + 1
End Function

Public Sub ShowGridItem()
   If TabStrip1.SelectedItem.Index = 1 Then
      GridEX1.ItemCount = CountItem(m_InventoryDoc.ImportExports)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      GridEX1.ItemCount = CountItem(m_InventoryDoc.ExportItems)
      GridEX1.Rebind
   End If
   m_HasModify = True
End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean

   If Not cmdAdd.Enabled Then
      Exit Sub
   End If

   OKClick = False
   
   If TabStrip1.SelectedItem.Index = 1 Then
      frmAddEditAdjustItemEx2.DocumentType = DocumentType
      Set frmAddEditAdjustItemEx2.ParentForm = Me
      frmAddEditAdjustItemEx2.COMMIT_FLAG = m_InventoryDoc.COMMIT_FLAG
      frmAddEditAdjustItemEx2.TxSeq = GetNextID
      Set frmAddEditAdjustItemEx2.TempCollection = m_InventoryDoc.ImportExports
      frmAddEditAdjustItemEx2.ParentShowMode = ShowMode
      frmAddEditAdjustItemEx2.ShowMode = SHOW_ADD
      frmAddEditAdjustItemEx2.HeaderText = MapText("������¡�û�Ѻ�ʹ")
      Load frmAddEditAdjustItemEx2
      frmAddEditAdjustItemEx2.Show 1

      OKClick = frmAddEditAdjustItemEx2.OKClick

      Unload frmAddEditAdjustItemEx2
      Set frmAddEditAdjustItemEx2 = Nothing

      If OKClick Then
         Call GetTotalPrice

         GridEX1.ItemCount = CountItem(m_InventoryDoc.ImportExports)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      frmAddEditAdjustItem2.COMMIT_FLAG = m_InventoryDoc.COMMIT_FLAG
      frmAddEditAdjustItem2.TxSeq = GetNextID
      Set frmAddEditAdjustItem2.TempCollection = m_InventoryDoc.ExportItems
      frmAddEditAdjustItem2.ParentShowMode = ShowMode
      frmAddEditAdjustItem2.ShowMode = SHOW_ADD
      frmAddEditAdjustItem2.HeaderText = MapText("������¡�û�Ѻ�ʹ (Ŵ)")
      Load frmAddEditAdjustItem2
      frmAddEditAdjustItem2.Show 1

      OKClick = frmAddEditAdjustItem2.OKClick

      Unload frmAddEditAdjustItem2
      Set frmAddEditAdjustItem2 = Nothing

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
         m_InventoryDoc.ImportExports.Remove (ID2)
      Else
         m_InventoryDoc.ImportExports.Item(ID2).Flag = "D"
      End If

      Call GetTotalPrice
      GridEX1.ItemCount = CountItem(m_InventoryDoc.ImportExports)
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
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(2))
   OKClick = False
   
   If TabStrip1.SelectedItem.Index = 1 Then
      frmAddEditAdjustItemEx2.DocumentType = DocumentType
      frmAddEditAdjustItemEx2.ID = ID
      Set frmAddEditAdjustItemEx2.ParentForm = Me
      frmAddEditAdjustItemEx2.COMMIT_FLAG = m_InventoryDoc.COMMIT_FLAG
      Set frmAddEditAdjustItemEx2.TempCollection = m_InventoryDoc.ImportExports
      frmAddEditAdjustItemEx2.HeaderText = MapText("�����¡�û�Ѻ�ʹ")
      frmAddEditAdjustItemEx2.ParentShowMode = ShowMode
      frmAddEditAdjustItemEx2.ShowMode = SHOW_EDIT
      Load frmAddEditAdjustItemEx2
      frmAddEditAdjustItemEx2.Show 1

      OKClick = frmAddEditAdjustItemEx2.OKClick

      Unload frmAddEditAdjustItemEx2
      Set frmAddEditAdjustItemEx2 = Nothing

      If OKClick Then
         Call GetTotalPrice
         GridEX1.ItemCount = CountItem(m_InventoryDoc.ImportExports)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      frmAddEditAdjustItem2.ID = ID
      frmAddEditAdjustItem2.COMMIT_FLAG = m_InventoryDoc.COMMIT_FLAG
      Set frmAddEditAdjustItem2.TempCollection = m_InventoryDoc.ExportItems
      frmAddEditAdjustItem2.HeaderText = MapText("�����¡�û�Ѻ�ʹ (Ŵ)")
      frmAddEditAdjustItem2.ParentShowMode = ShowMode
      frmAddEditAdjustItem2.ShowMode = SHOW_EDIT
      Load frmAddEditAdjustItem2
      frmAddEditAdjustItem2.Show 1

      OKClick = frmAddEditAdjustItem2.OKClick

      Unload frmAddEditAdjustItem2
      Set frmAddEditAdjustItem2 = Nothing

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
      ID = m_InventoryDoc.INVENTORY_DOC_ID
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

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      Call LoadReason(cboReason, , 2)
      
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
   Col.Width = 2100
   Col.Caption = MapText("�����Ţ�ѵ�شԺ")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 4425
   Col.Caption = MapText("�ѵ�شԺ")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 1785
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("����ҳ")

   Set Col = GridEX1.Columns.add '6
   Col.Width = 1620
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("�Ҥ�")
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = 1620
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("�Ҥ����")
   
   Set Col = GridEX1.Columns.add '8
   Col.Width = 1980
   Col.Caption = MapText("ʶҹ���Ѵ��")
End Sub

Private Sub GetTotalPrice()
Dim II As CLotItem
Dim Sum As Double

   Sum = 0
   For Each II In m_InventoryDoc.ImportExports
      If II.Flag <> "D" Then
         Sum = Sum + II.NEED_TOTAL_AMOUNT
      End If
   Next II
   
   txtTotalAmount.Text = FormatNumber(Sum, 2)
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblDocumentNo, MapText("�����Ţ㺻�Ѻ�ʹ"))
   Call InitNormalLabel(lblDocumentDate, MapText("�ѹ����͡���"))
   Call InitNormalLabel(Label4, MapText("�ҷ"))
   Call InitNormalLabel(lblEmployeeNo, MapText("����Ѻ�Դ�ͺ"))
   Call InitNormalLabel(lblReason, MapText("���˵ء�û�Ѻ�ʹ"))
   Call InitNormalLabel(lblTotalAmount, MapText("�ӹǹ���"))
   
   Call txtDocumentNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtTotalAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtTotalAmount.Enabled = False
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
   Call InitMainButton(cmdAdd, MapText("���� (F7)"))
   Call InitMainButton(cmdEdit, MapText("��� (F3)"))
   Call InitMainButton(cmdDelete, MapText("ź (F6)"))
   
   Call InitCheckBox(chkCommit, MapText("�ӹǳ"))
   Call InitCheckBox(chkAdjustFlag, MapText("��Ѻ�ʹ�ҡ��Ǩ�Ѻ��ԧ"))
   
   Call InitCombo(cboReason)
   
   Call InitGrid1
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.NAME = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.add().Caption = MapText("��¡�û�Ѻ�ʹ")
'   TabStrip1.Tabs.add().Caption = MapText("��Ѻ�ʹ�ѵ�شԺ (Ŵ)")
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
      Values(4) = CR.PART_DESC
      Values(5) = FormatNumber(CR.NEED_TOTAL_AMOUNT)
      Values(6) = FormatNumber(CR.NEED_AVG_PRICE)
      Values(7) = FormatNumber(CR.NEED_TOTAL_PRICE)
      Values(8) = CR.LOCATION_NAME
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      If m_InventoryDoc.ExportItems Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim Ei As CLotItem
      If m_InventoryDoc.ExportItems.Count <= 0 Then
         Exit Sub
      End If
      Set Ei = GetItem(m_InventoryDoc.ExportItems, RowIndex, RealIndex)
      If Ei Is Nothing Then
         Exit Sub
      End If

      Values(1) = Ei.LOT_ITEM_ID
      Values(2) = RealIndex
      Values(3) = Ei.PART_NO
      Values(4) = Ei.PART_DESC
      Values(5) = FormatNumber(Ei.TX_AMOUNT)
      Values(6) = FormatNumber(MyDiffEx(Ei.TOTAL_INCLUDE_PRICE, Ei.TX_AMOUNT))
      Values(7) = FormatNumber(Ei.TOTAL_INCLUDE_PRICE)
      Values(8) = Ei.LOCATION_NAME
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
      GridEX1.ItemCount = CountItem(m_InventoryDoc.ImportExports)
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

Private Sub txtTotalAmount_Change()
   m_HasModify = True
End Sub

Private Sub uctlDocumentDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlEmployeeLookup_Change()
   m_HasModify = True
End Sub
