VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditExpense 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditExpense.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8895
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   12720
      _ExtentX        =   22437
      _ExtentY        =   15690
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   120
         TabIndex        =   2
         Top             =   2010
         Width           =   11640
         _ExtentX        =   20532
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
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   5190
         Left            =   120
         TabIndex        =   3
         Top             =   2550
         Width           =   11640
         _ExtentX        =   20532
         _ExtentY        =   9155
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
         Column(1)       =   "frmAddEditExpense.frx":27A2
         Column(2)       =   "frmAddEditExpense.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditExpense.frx":290E
         FormatStyle(2)  =   "frmAddEditExpense.frx":2A6A
         FormatStyle(3)  =   "frmAddEditExpense.frx":2B1A
         FormatStyle(4)  =   "frmAddEditExpense.frx":2BCE
         FormatStyle(5)  =   "frmAddEditExpense.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditExpense.frx":2D5E
      End
      Begin prjFarmManagement.uctlTextBox txtExpenseNo 
         Height          =   435
         Left            =   1800
         TabIndex        =   0
         Top             =   990
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlExpenseDate 
         Height          =   405
         Left            =   7500
         TabIndex        =   1
         Top             =   990
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtExpenseDesc 
         Height          =   435
         Left            =   1800
         TabIndex        =   14
         Top             =   1440
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   767
      End
      Begin VB.Label lblExpenseDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "lblExpenseDesc"
         Height          =   315
         Left            =   210
         TabIndex        =   15
         Top             =   1560
         Width           =   1395
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   525
         Left            =   6840
         TabIndex        =   7
         Top             =   7830
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditExpense.frx":2F36
         ButtonStyle     =   3
      End
      Begin VB.Label lblExpenseDate 
         Alignment       =   1  'Right Justify
         Caption         =   "lblExpenseDate"
         Height          =   315
         Left            =   6120
         TabIndex        =   13
         Top             =   1050
         Width           =   1305
      End
      Begin VB.Label lblExpenseNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblExpenseNo"
         Height          =   315
         Left            =   210
         TabIndex        =   12
         Top             =   1110
         Width           =   1395
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8490
         TabIndex        =   8
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditExpense.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10170
         TabIndex        =   9
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1740
         TabIndex        =   5
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   120
         TabIndex        =   4
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditExpense.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3390
         TabIndex        =   6
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditExpense.frx":3884
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditExpense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Expense As CExpense

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public id As Long
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
      
      Call m_Expense.SetFieldValue("EXPENSE_ID", id)
      m_Expense.QueryFlag = 1
      If Not glbDaily.QueryExpense(m_Expense, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
        Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_Expense.PopulateFromRS(1, m_Rs)

      txtExpenseNo.Text = m_Expense.GetFieldValue("EXPENSE_NO")
      txtExpenseDesc.Text = m_Expense.GetFieldValue("EXPENSE_DESC")
      uctlExpenseDate.ShowDate = m_Expense.GetFieldValue("EXPENSE_DATE")
      
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call TabStrip1_Click
   
   Call EnableForm(Me, True)
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim Ivd As CInventoryDoc
   
   If Not VerifyTextControl(lblExpenseNo, txtExpenseNo, False) Then
       Exit Function
   End If
   If Not VerifyDate(lblExpenseDate, uctlExpenseDate, False) Then
      Exit Function
   End If
       
'   If Not CheckUniqueNs(JOB_NO, txtJobNo.Text, ID) Then
'      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtJobNo.Text & " " & MapText("อยู่ในระบบแล้ว")
'      glbErrorLog.ShowUserError
'      Exit Function
'   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Call m_Expense.SetFieldValue("EXPENSE_ID", id)
   m_Expense.ShowMode = ShowMode
   Call m_Expense.SetFieldValue("EXPENSE_NO", txtExpenseNo.Text)
   Call m_Expense.SetFieldValue("EXPENSE_DATE", uctlExpenseDate.ShowDate)
   Call m_Expense.SetFieldValue("EXPENSE_DESC", txtExpenseDesc.Text)
   
   Call EnableForm(Me, False)
      
   If Not glbDaily.AddEditExpense(m_Expense, IsOK, True, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call glbDaily.RollbackTransaction
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
Private Sub cmdAdd_Click()
Dim iCount As Long

   Call EnableForm(Me, False)
     
   If TabStrip1.SelectedItem.Index = 1 Then
      Set frmAddEditExpenseDetail.TempCollection = m_Expense.ExpenseDetail
      Set frmAddEditExpenseDetail.ParentForm = Me
      frmAddEditExpenseDetail.ShowMode = SHOW_ADD
      frmAddEditExpenseDetail.HeaderText = MapText("เพิ่มต้นทุนผลิต")
      Load frmAddEditExpenseDetail
      frmAddEditExpenseDetail.Show 1

      OKClick = frmAddEditExpenseDetail.OKClick

      Unload frmAddEditExpenseDetail
      Set frmAddEditExpenseDetail = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Expense.ExpenseDetail)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   
   End If
   Call EnableForm(Me, True)
   
   m_HasModify = True
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
         m_Expense.ExpenseDetail.Remove (ID2)
      Else
         m_Expense.ExpenseDetail.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_Expense.ExpenseDetail)
      GridEX1.Rebind
      m_HasModify = True
    ElseIf TabStrip1.SelectedItem.Index = 2 Then
   End If
End Sub
Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim id As Long
Dim OKClick As Boolean

   If Not cmdEdit.Enabled Then
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
   id = Val(GridEX1.Value(2))
   OKClick = False
   
   If TabStrip1.SelectedItem.Index = 1 Then
      Set frmAddEditExpenseDetail.TempCollection = m_Expense.ExpenseDetail
      Set frmAddEditExpenseDetail.ParentForm = Me
      frmAddEditExpenseDetail.id = id
      frmAddEditExpenseDetail.ShowMode = SHOW_EDIT
      frmAddEditExpenseDetail.HeaderText = MapText("แก้ไขต้นทุนผลิต")
      Load frmAddEditExpenseDetail
      frmAddEditExpenseDetail.Show 1

      OKClick = frmAddEditExpenseDetail.OKClick

      Unload frmAddEditExpenseDetail
      Set frmAddEditExpenseDetail = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Expense.ExpenseDetail)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
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

Private Sub cmdSave_Click()
Dim Result As Boolean
   If Not SaveData Then
      Exit Sub
   End If
   
   ShowMode = SHOW_EDIT
   id = m_Expense.GetFieldValue("EXPENSE_ID")
   m_Expense.QueryFlag = 1
   QueryData (True)
   m_HasModify = False
End Sub
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      
'      Call LoadPartItem(Nothing, m_PartItems)
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_Expense.QueryFlag = 1
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         uctlExpenseDate.ShowDate = Now
         
        m_Expense.QueryFlag = 0
         Call QueryData(False)
      End If
      
      Call TabStrip1_Click
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
   
   Set m_Expense = Nothing
   
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
   
'   GridEX1.Font.Bold = False
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"
   GridEX1.Columns.Item(1).Visible = False

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
   GridEX1.Columns.Item(2).Visible = False
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2000
   Col.Caption = MapText("ประเภทผลิต")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 5000
   Col.Caption = MapText("รายละเอียด")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 1500
   Col.Caption = MapText("จำนวน")

   Set Col = GridEX1.Columns.add '6
   Col.Width = 1500
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ราคาเฉลี่ย")
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = 1500
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("มูลค่า")
End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitGrid1
   Call InitNormalLabel(lblExpenseNo, MapText("เลขที่เอกสาร"))
   Call InitNormalLabel(lblExpenseDate, MapText("วันที่เอกสาร"))
   Call InitNormalLabel(lblExpenseDesc, MapText("รายละเอียด"))
   
   Call txtExpenseNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSave.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSave, MapText("บันทึก"))
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.NAME = GLB_FONT
   TabStrip1.Font.Size = 16
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.add().Caption = MapText("ค่าใช้จ่าย")
   
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
   Set m_Expense = New CExpense
End Sub
Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub
Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   If TabStrip1.SelectedItem.Index = 6 Then
      RowBuffer.RowStyle = RowBuffer.Value(6)
   End If
End Sub
Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Index = 1 Then
     If m_Expense.ExpenseDetail Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim Ce As CExpenseDetail
      If m_Expense.ExpenseDetail.Count <= 0 Then
         Exit Sub
      End If
      Set Ce = GetItem(m_Expense.ExpenseDetail, RowIndex, RealIndex)
      If Ce Is Nothing Then
         Exit Sub
      End If

      Values(1) = Ce.GetFieldValue("EXPENSE_DETAIL_ID")
      Values(2) = RealIndex
      Values(3) = Ce.GetFieldValue("PARAMETER_PROCESS_NAME")
      Values(4) = Ce.GetFieldValue("EXPENSE_DETAIL_DESC")
      Values(5) = FormatNumber(Ce.GetFieldValue("EXPENSE_DETAIL_AMOUNT"))
      Values(6) = FormatNumber(Ce.GetFieldValue("EXPENSE_DETAIL_AVG"))
      Values(7) = FormatNumber(Ce.GetFieldValue("EXPENSE_DETAIL_PRICE"))
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   End If
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub TabStrip1_Click()

   If TabStrip1.SelectedItem.Index = 1 Then
      cmdAdd.Enabled = True
      cmdEdit.Enabled = True
      cmdDelete.Enabled = True
      
      Call InitGrid1
     GridEX1.ItemCount = CountItem(m_Expense.ExpenseDetail)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      
   End If
End Sub

Private Sub txtExpense_Change()
   m_HasModify = True
End Sub

Private Sub txtExpenseDesc_Change()
   m_HasModify = True
End Sub

Private Sub uctlExpenseDate_HasChange()
   m_HasModify = True
End Sub
Public Sub RefreshGrid()
   GridEX1.ItemCount = CountItem(m_Expense.ExpenseDetail)
   GridEX1.Rebind
End Sub

