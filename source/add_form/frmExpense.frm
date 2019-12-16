VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmExpense 
   Caption         =   "Form1"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmExpense.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8445
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8520
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboOrdertype 
         Height          =   315
         Left            =   6060
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1530
         Width           =   2655
      End
      Begin VB.ComboBox cboOrderBy 
         Height          =   315
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1530
         Width           =   3100
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   12645
         _ExtentX        =   22304
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlDate uctlExpenseDate 
         Height          =   405
         Left            =   6060
         TabIndex        =   1
         Top             =   1080
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtExpenseNo 
         Height          =   435
         Left            =   1650
         TabIndex        =   0
         Top             =   1080
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   767
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   5205
         Left            =   135
         TabIndex        =   6
         Top             =   2490
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   9181
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
         Column(1)       =   "frmExpense.frx":030A
         Column(2)       =   "frmExpense.frx":03D2
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmExpense.frx":0476
         FormatStyle(2)  =   "frmExpense.frx":05D2
         FormatStyle(3)  =   "frmExpense.frx":0682
         FormatStyle(4)  =   "frmExpense.frx":0736
         FormatStyle(5)  =   "frmExpense.frx":080E
         ImageCount      =   0
         PrinterProperties=   "frmExpense.frx":08C6
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "lblOrderBy"
         Height          =   315
         Left            =   390
         TabIndex        =   17
         Top             =   1590
         Width           =   1185
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "lblOrderType"
         Height          =   315
         Left            =   4890
         TabIndex        =   16
         Top             =   1620
         Width           =   1095
      End
      Begin VB.Label lblExpenseDate 
         Alignment       =   1  'Right Justify
         Caption         =   "lblJobDate"
         Height          =   315
         Left            =   4740
         TabIndex        =   15
         Top             =   1170
         Width           =   1245
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   10140
         TabIndex        =   5
         Top             =   1650
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   10140
         TabIndex        =   4
         Top             =   1050
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmExpense.frx":0A9E
         ButtonStyle     =   3
      End
      Begin VB.Label lblExpenseNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblExpenseNo"
         Height          =   315
         Left            =   60
         TabIndex        =   14
         Top             =   1170
         Width           =   1485
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8535
         TabIndex        =   10
         Top             =   7800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10140
         TabIndex        =   11
         Top             =   7800
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   8
         Top             =   7800
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   7
         Top             =   7800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   9
         Top             =   7800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmExpense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_Expense As CExpense
Private m_TempExpense As CExpense
Private m_Rs As ADODB.Recordset
Private m_TableName As String
Private m_HasModify As Boolean

Public OKClick As Boolean

Private Sub cmdSearch_Click()
   Call QueryData(True)
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long

   If Flag Then
      Call EnableForm(Me, False)

      Call m_Expense.SetFieldValue("EXPENSE_ID", -1)
      Call m_Expense.SetFieldValue("EXPENSE_NO", txtExpenseNo.Text)
      Call m_Expense.SetFieldValue("FROM_DATE", uctlExpenseDate.ShowDate)
      Call m_Expense.SetFieldValue("TO_DATE", uctlExpenseDate.ShowDate)
      Call m_Expense.SetFieldValue("ORDER_BY", cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex)))
      Call m_Expense.SetFieldValue("ORDER_TYPE", cboOrderType.ItemData(Minus2Zero(cboOrderType.ListIndex)))
      
      If Not glbDaily.QueryExpense(m_Expense, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
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
Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean

   frmAddEditExpense.HeaderText = MapText("เพิ่มข้อมูลค่าใช้จ่าย")
   frmAddEditExpense.ShowMode = SHOW_ADD
   Load frmAddEditExpense
   frmAddEditExpense.Show 1

   OKClick = frmAddEditExpense.OKClick

   Unload frmAddEditExpense
   Set frmAddEditExpense = Nothing
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub
Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim id As Long

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
   If Not glbDaily.DeleteExpense(id, IsOK, True, glbErrorLog) Then
      Call m_Expense.SetFieldValue("EXPENSE_ID", -1)
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
   
   frmAddEditExpense.id = id
   frmAddEditExpense.HeaderText = MapText("แก้ไขข้อมูลค่าใช้จ่าย")
   frmAddEditExpense.ShowMode = SHOW_EDIT
   Load frmAddEditExpense
   frmAddEditExpense.Show 1

   OKClick = frmAddEditExpense.OKClick

   Unload frmAddEditExpense
   Set frmAddEditExpense = Nothing
               
   If OKClick Then
      Call QueryData(True)
   End If
   Call glbDatabaseMngr.UnLockTable(m_TableName, id, IsCanLock, glbErrorLog)

End Sub
Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub
Private Sub Form_Activate()
 If Not m_HasActivate Then
      m_HasActivate = True
      
      Call InitExpenseOrderBy(cboOrderBy)
      Call InitOrderType(cboOrderType)
      
      Call QueryData(True)
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.NAME
      glbErrorLog.ShowUserError
   ElseIf Shift = 0 And KeyCode = 116 Then
      Call cmdSearch_Click
   ElseIf Shift = 0 And KeyCode = 115 Then
      Call cmdClear_Click
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
   
   Set m_Expense = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
End Sub
Private Sub InitGrid1()
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
   Col.Width = 2805
   Col.Caption = "เลขที่เอกสาร"

   Set Col = GridEX1.Columns.add '3
   Col.Width = 3360
   Col.Caption = MapText("วันที่เอกสาร")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 5000
   Col.Caption = MapText("รายละเอียด")
   
   GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = MapText("รายการค่าใช้จ่าย")
   pnlHeader.Caption = MapText("รายการค่าใช้จ่าย")
   
   Call InitNormalLabel(lblExpenseNo, MapText("เลขที่เอกสาร"))
   Call InitNormalLabel(lblExpenseDate, MapText("วันที่เอกสาร"))

   Call InitNormalLabel(lblOrderBy, MapText("เรียงตาม"))
   Call InitNormalLabel(lblOrderType, MapText("เรียงจาก"))
   
   Call txtExpenseNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   Call InitCombo(cboOrderBy)
   Call InitCombo(cboOrderType)
  
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdClear.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา (F5)"))
   Call InitMainButton(cmdClear, MapText("เคลียร์ (F4)"))
   
   Call InitGrid1
End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   m_TableName = "EXPENSE"
   Call InitFormLayout
      
   Set m_Rs = New ADODB.Recordset
   Set m_Expense = New CExpense
   Set m_TempExpense = New CExpense
   Call EnableForm(Me, True)
End Sub
Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
'   RowBuffer.RowStyle = RowBuffer.Value(11)
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
   Call m_TempExpense.PopulateFromRS(1, m_Rs)
   
   Values(1) = m_TempExpense.GetFieldValue("EXPENSE_ID")
   Values(2) = m_TempExpense.GetFieldValue("EXPENSE_NO")
   Values(3) = DateToStringExtEx2(m_TempExpense.GetFieldValue("EXPENSE_DATE"))
   Values(4) = m_TempExpense.GetFieldValue("EXPENSE_DESC")
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub cmdClear_Click()
   txtExpenseNo.Text = ""
   cboOrderBy.ListIndex = -1
   cboOrderType.ListIndex = -1
   uctlExpenseDate.ShowDate = -1
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
