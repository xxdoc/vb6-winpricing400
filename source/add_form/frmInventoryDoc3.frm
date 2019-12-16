VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmInventoryDoc3 
   BackColor       =   &H80000000&
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   11910
   Icon            =   "frmInventoryDoc3.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlDate uctlDocumentDate 
         Height          =   405
         Left            =   6180
         TabIndex        =   1
         Top             =   1110
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin VB.ComboBox cboOrderType 
         Height          =   315
         Left            =   6180
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1980
         Width           =   2595
      End
      Begin VB.ComboBox cboOrderBy 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1980
         Width           =   2985
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   15
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   5055
         Left            =   180
         TabIndex        =   8
         Top             =   2640
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   8916
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
         Column(1)       =   "frmInventoryDoc3.frx":27A2
         Column(2)       =   "frmInventoryDoc3.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmInventoryDoc3.frx":290E
         FormatStyle(2)  =   "frmInventoryDoc3.frx":2A6A
         FormatStyle(3)  =   "frmInventoryDoc3.frx":2B1A
         FormatStyle(4)  =   "frmInventoryDoc3.frx":2BCE
         FormatStyle(5)  =   "frmInventoryDoc3.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmInventoryDoc3.frx":2D5E
      End
      Begin prjFarmManagement.uctlTextBox txtDocumentNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   1080
         Width           =   2985
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPartNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   2
         Top             =   1530
         Width           =   2985
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin Threed.SSCheck chkCommit 
         Height          =   435
         Left            =   6180
         TabIndex        =   3
         Top             =   1530
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblDocumentNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   20
         Top             =   1140
         Width           =   1755
      End
      Begin VB.Label lblPartNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   19
         Top             =   1590
         Width           =   1755
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4980
         TabIndex        =   18
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4920
         TabIndex        =   17
         Top             =   1140
         Width           =   1185
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   16
         Top             =   2040
         Width           =   1755
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   10110
         TabIndex        =   6
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmInventoryDoc3.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   10110
         TabIndex        =   7
         Top             =   1650
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   11
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmInventoryDoc3.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   9
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmInventoryDoc3.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   10
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmInventoryDoc3.frx":3884
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmInventoryDoc3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_InventoryDoc As CInventoryDoc
Private m_TempInventoryDoc As CInventoryDoc
Private m_Rs As ADODB.Recordset
Private m_TableName As String
Public DocumentType As Long
Public DocPartType As Long
Public OKClick As Boolean
Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean
   If Not VerifyAccessRight("INVENTORY_TRANSFER_ADD") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   frmAddEditInventoryDoc3.DocumentType = DocumentType
  If DocPartType = 1 Then
     frmAddEditInventoryDoc3.HeaderText = MapText("���������š���͹�Թ��������ҧ��ѧ")
     frmAddEditInventoryDoc3.DocPartType = DocPartType
  Else
     frmAddEditInventoryDoc3.HeaderText = MapText("���������š���͹�ѵ�شԺ")
  End If
   
   frmAddEditInventoryDoc3.ShowMode = SHOW_ADD
   Load frmAddEditInventoryDoc3
   frmAddEditInventoryDoc3.Show 1
   
   OKClick = frmAddEditInventoryDoc3.OKClick
   
   Unload frmAddEditInventoryDoc3
   Set frmAddEditInventoryDoc3 = Nothing
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub

Private Sub cmdClear_Click()
   txtDocumentNo.Text = ""
   txtPartNo.Text = ""
   uctlDocumentDate.ShowDate = -1
   cboOrderBy.ListIndex = -1
   cboOrderType.ListIndex = -1
End Sub

Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim id As Long
   If Not VerifyAccessRight("INVENTORY_TRANSFER_DELETE") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   If Not cmdDelete.Enabled Then
      Exit Sub
   End If

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   id = GridEX1.Value(1)
   
   Call glbDatabaseMngr.LockTable(m_TableName, id, IsCanLock, glbErrorLog)
   If Not ConfirmDelete(GridEX1.Value(2)) Then
      Call glbDatabaseMngr.UnLockTable(m_TableName, id, IsCanLock, glbErrorLog)
      Exit Sub
   End If

   Call EnableForm(Me, False)
   If Not glbDaily.DeleteInventoryDoc(id, IsOK, True, glbErrorLog) Then
      m_InventoryDoc.INVENTORY_DOC_ID = -1
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call glbDatabaseMngr.UnLockTable(m_TableName, id, IsCanLock, glbErrorLog)
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call QueryData(True)
   
   Call glbDatabaseMngr.UnLockTable(m_TableName, id, IsCanLock, glbErrorLog)
   Call EnableForm(Me, True)
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

   id = Val(GridEX1.Value(1))
   Call glbDatabaseMngr.LockTable(m_TableName, id, IsCanLock, glbErrorLog)
   
   frmAddEditInventoryDoc3.DocumentType = DocumentType
   frmAddEditInventoryDoc3.id = id
     If DocPartType = 1 Then
     frmAddEditInventoryDoc3.HeaderText = MapText("��䢢����š���͹�Թ��������ҧ��ѧ")
     frmAddEditInventoryDoc3.DocPartType = DocPartType
  Else
     frmAddEditInventoryDoc3.HeaderText = MapText("��䢢����š���͹�ѵ�شԺ")
  End If

   frmAddEditInventoryDoc3.ShowMode = SHOW_EDIT
   Load frmAddEditInventoryDoc3
   frmAddEditInventoryDoc3.Show 1
   
   OKClick = frmAddEditInventoryDoc3.OKClick
   
   Unload frmAddEditInventoryDoc3
   Set frmAddEditInventoryDoc3 = Nothing
               
   If OKClick Then
      Call QueryData(True)
   End If
   Call glbDatabaseMngr.UnLockTable(m_TableName, id, IsCanLock, glbErrorLog)

End Sub

Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub

Private Sub cmdSearch_Click()
   Call QueryData(True)
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      
      Call InitInventoryDoc3OrderBy(cboOrderBy)
      Call InitOrderType(cboOrderType)
      Call QueryData(True)
   End If
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      m_InventoryDoc.DOCUMENT_NO = PatchWildCard(txtDocumentNo.Text)
      m_InventoryDoc.PART_NO = PatchWildCard(txtPartNo.Text)
      m_InventoryDoc.DOCUMENT_DATE = uctlDocumentDate.ShowDate
      m_InventoryDoc.DOCUMENT_TYPE = DocumentType '��͹
'      If DocPartType = 1 Then
'         m_InventoryDoc.INVENTORY_WH_DOC_ID = -2 'CHECK NOT NULL
'       Else
'         m_InventoryDoc.INVENTORY_WH_DOC_ID = -3 'CHECK  NULL
'      End If
      m_InventoryDoc.OrderBy = cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex))
      m_InventoryDoc.OrderType = cboOrderType.ItemData(Minus2Zero(cboOrderType.ListIndex))
      m_InventoryDoc.COMMIT_FLAG = Check2Flag(chkCommit.Value)
      If Not glbDaily.QueryInventoryDoc(m_InventoryDoc, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      cmdDelete.Enabled = (m_InventoryDoc.COMMIT_FLAG = "N")
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
Private Sub QueryData2(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      m_InventoryDoc.DOCUMENT_NO = PatchWildCard(txtDocumentNo.Text)
      m_InventoryDoc.PART_NO = PatchWildCard(txtPartNo.Text)
      m_InventoryDoc.DOCUMENT_DATE = uctlDocumentDate.ShowDate
      m_InventoryDoc.DOCUMENT_TYPE = DocumentType '��͹
      m_InventoryDoc.OrderBy = cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex))
      m_InventoryDoc.OrderType = cboOrderType.ItemData(Minus2Zero(cboOrderType.ListIndex))
      m_InventoryDoc.COMMIT_FLAG = Check2Flag(chkCommit.Value)
      If Not glbDaily.QueryInventoryDoc(m_InventoryDoc, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      cmdDelete.Enabled = (m_InventoryDoc.COMMIT_FLAG = "N")
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
      KeyCode = 0
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
   Col.Width = 2115
   Col.Caption = MapText("�����Ţ��͹")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2055
   Col.Caption = MapText("�ѹ����͡���")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 7305
   Col.Caption = MapText("����Ѻ�Դ�ͺ")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 0
   Col.Visible = False
   Col.Caption = "Commit flag"
   
   GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = MapText("�͹�����ѵ�شԺ")
   pnlHeader.Caption = MapText("�͹�����ѵ�شԺ")
   
   Call InitGrid
   
   Call InitNormalLabel(lblDocumentDate, MapText("�ѹ����͡���"))
   Call InitNormalLabel(lblPartNo, MapText("�����Ţ�ѵ�شԺ"))
   Call InitNormalLabel(lblDocumentNo, MapText("�Ţ�����͹"))
   Call InitNormalLabel(lblOrderBy, MapText("���§���"))
   Call InitNormalLabel(lblOrderType, MapText("���§�ҡ"))
   
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
   
   Call InitCheckBox(chkCommit, "�ӹǳ����")
   
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
   Call InitMainButton(cmdAdd, MapText("���� (F7)"))
   Call InitMainButton(cmdEdit, MapText("��� (F3)"))
   Call InitMainButton(cmdDelete, MapText("ź (F6)"))
   Call InitMainButton(cmdSearch, MapText("���� (F5)"))
   Call InitMainButton(cmdClear, MapText("������ (F4)"))
End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   m_TableName = "USER_GROUP"
   
   Set m_InventoryDoc = New CInventoryDoc
   Set m_TempInventoryDoc = New CInventoryDoc
   Set m_Rs = New ADODB.Recordset

   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   RowBuffer.RowStyle = RowBuffer.Value(5)
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

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
   Call m_TempInventoryDoc.PopulateFromRS(1, m_Rs)
   
   Values(1) = m_TempInventoryDoc.INVENTORY_DOC_ID
   Values(2) = m_TempInventoryDoc.DOCUMENT_NO
   Values(3) = DateToStringExtEx2(m_TempInventoryDoc.DOCUMENT_DATE)
   Values(4) = m_TempInventoryDoc.RESPONSE_NAME & " " & m_TempInventoryDoc.RESPONSE_LNAME
   Values(5) = m_TempInventoryDoc.COMMIT_FLAG
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
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
End Sub
