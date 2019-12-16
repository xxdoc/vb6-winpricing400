VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmMoneyPerson 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmMoneyPerson.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
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
      Begin VB.ComboBox cboOrderType 
         Height          =   315
         Left            =   6120
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2400
         Width           =   2955
      End
      Begin VB.ComboBox cboOrderBy 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2400
         Width           =   2955
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
      Begin GridEX20.GridEX GridEX1 
         Height          =   4695
         Left            =   180
         TabIndex        =   8
         Top             =   3000
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   8281
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
         Column(1)       =   "frmMoneyPerson.frx":27A2
         Column(2)       =   "frmMoneyPerson.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmMoneyPerson.frx":290E
         FormatStyle(2)  =   "frmMoneyPerson.frx":2A6A
         FormatStyle(3)  =   "frmMoneyPerson.frx":2B1A
         FormatStyle(4)  =   "frmMoneyPerson.frx":2BCE
         FormatStyle(5)  =   "frmMoneyPerson.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmMoneyPerson.frx":2D5E
      End
      Begin prjFarmManagement.uctlTextBox txtLendCode 
         Height          =   435
         Left            =   1680
         TabIndex        =   0
         Top             =   960
         Width           =   2985
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlFromDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   1440
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin prjFarmManagement.uctlDate uctlToDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   1920
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin prjFarmManagement.uctlTextBox txtLender 
         Height          =   435
         Left            =   6120
         TabIndex        =   1
         Top             =   960
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   767
      End
      Begin VB.Label lblLender 
         Alignment       =   1  'Right Justify
         Caption         =   "lblLender"
         Height          =   315
         Left            =   4800
         TabIndex        =   21
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Caption         =   "lblToDate"
         Height          =   315
         Left            =   0
         TabIndex        =   20
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Caption         =   "lblFromDate"
         Height          =   315
         Left            =   0
         TabIndex        =   19
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label lblLendCode 
         Alignment       =   1  'Right Justify
         Caption         =   "lblLendCode"
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   255
         Left            =   4800
         TabIndex        =   17
         Top             =   2520
         Width           =   1245
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   2520
         Width           =   1455
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   10110
         TabIndex        =   6
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmMoneyPerson.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   10110
         TabIndex        =   7
         Top             =   1440
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
         MouseIcon       =   "frmMoneyPerson.frx":3250
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
         MouseIcon       =   "frmMoneyPerson.frx":356A
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
         MouseIcon       =   "frmMoneyPerson.frx":3884
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmMoneyPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_EmpReceivable As CEmpReceivable
Private m_TempEmpReceivable As CEmpReceivable
Private m_Employee As CEmployee
Private m_Rs As ADODB.Recordset
Private m_TableName As String

Public OKClick As Boolean
Private Sub cmdAdd_Click()
Dim itemcount As Long
Dim OKClick As Boolean

   frmAddEditMoneyPerson.HeaderText = MapText("เพิ่มข้อมูลการยืม")
   frmAddEditMoneyPerson.ShowMode = SHOW_ADD
   Load frmAddEditMoneyPerson
   frmAddEditMoneyPerson.Show 1
   
   OKClick = frmAddEditMoneyPerson.OKClick
   
   Unload frmAddEditMoneyPerson
   Set frmAddEditMoneyPerson = Nothing
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub

Private Sub cmdClear_Click()
   txtLendCode.Text = ""
   txtLender.Text = ""
   cboOrderBy.ListIndex = -1
   cboOrdertype.ListIndex = -1
   
Call QueryData(True)
End Sub

Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim itemcount As Long
Dim IsCanLock As Boolean
Dim ID As Long

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   ID = GridEX1.Value(1)
   
   Call glbDatabaseMngr.LockTable(m_TableName, ID, IsCanLock, glbErrorLog)
   If Not ConfirmDelete(GridEX1.Value(2)) Then
      Call glbDatabaseMngr.UnLockTable(m_TableName, ID, IsCanLock, glbErrorLog)
      Exit Sub
   End If

   Call EnableForm(Me, False)
m_EmpReceivable.EMP_RECEIVABLE_ID = ID
        If Not glbDaily.QueryEmpReceivable(m_EmpReceivable, m_Rs, itemcount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
If itemcount <> 0 Then
Call m_EmpReceivable.PopulateFromRS(m_Rs)
m_Employee.EMP_ID = m_EmpReceivable.EMP_ID
End If

   Call glbDaily.StartTransaction
   If Not glbDaily.DeleteEmpReceivable(ID, IsOK, glbErrorLog) Then
      m_EmpReceivable.EMP_ID = -1
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call glbDatabaseMngr.UnLockTable(m_TableName, ID, IsCanLock, glbErrorLog)
      Call EnableForm(Me, True)
      Call glbDaily.RollbackTransaction
      Exit Sub
   End If
   
   m_Employee.EMP_ID = m_EmpReceivable.EMP_ID
   If Not glbDaily.QueryEmployee(m_Employee, m_Rs, itemcount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
        Call EnableForm(Me, True)
         Exit Sub
      End If
      Call m_Employee.PopulateFromRSMoney(1, m_Rs)
      m_Employee.TOTBORROW = m_Employee.TOTBORROW - m_EmpReceivable.BORROW_AMOUNT
 If Not glbDaily.AddEditEmployeeMoney(m_Employee, IsOK, False, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call glbDatabaseMngr.UnLockTable(m_TableName, ID, IsCanLock, glbErrorLog)
      Call glbDaily.RollbackTransaction
      Call EnableForm(Me, True)
      Exit Sub
   End If
Call glbDaily.CommitTransaction

   Call QueryData(True)
   
   Call glbDatabaseMngr.UnLockTable(m_TableName, ID, IsCanLock, glbErrorLog)
   Call EnableForm(Me, True)
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim itemcount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean

   If Not cmdEdit.Enabled Then
      Exit Sub
   End If
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(1))
   Call glbDatabaseMngr.LockTable(m_TableName, ID, IsCanLock, glbErrorLog)
   
   frmAddEditMoneyPerson.ID = ID
   frmAddEditMoneyPerson.HeaderText = MapText("แก้ไขข้อมูลการยืม")
   frmAddEditMoneyPerson.ShowMode = SHOW_EDIT
   Load frmAddEditMoneyPerson
   frmAddEditMoneyPerson.Show 1
   
   OKClick = frmAddEditMoneyPerson.OKClick
   
   Unload frmAddEditMoneyPerson
   Set frmAddEditMoneyPerson = Nothing
               
   If OKClick Then
      Call QueryData(True)
   End If
   Call glbDatabaseMngr.UnLockTable(m_TableName, ID, IsCanLock, glbErrorLog)

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
      
      Call InitEmpReceivableOrderBy(cboOrderBy)
      Call InitOrderType(cboOrdertype)
      
      Call QueryData(True)
   End If
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long
Dim Temp As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      m_EmpReceivable.EMP_ID = -1
      m_EmpReceivable.EMP_RECEIVABLE_ID = -1
      m_EmpReceivable.BORROW_NO = txtLendCode.Text
      m_EmpReceivable.FROM_DATE = uctlFromDate.ShowDate
      m_EmpReceivable.TO_DATE = uctlToDate.ShowDate
      m_EmpReceivable.EMP_NAME = txtLender.Text
      m_EmpReceivable.OrderBy = cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex))
      m_EmpReceivable.OrderType = cboOrdertype.ItemData(Minus2Zero(cboOrdertype.ListIndex))
     If Not glbDaily.QueryEmpReceivable(m_EmpReceivable, m_Rs, itemcount, IsOK, glbErrorLog) Then
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
   
   GridEX1.itemcount = itemcount
   GridEX1.Rebind
   
   Call EnableForm(Me, True)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
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
Dim fmsTemp2 As JSFormatStyle

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.FormatStyles.Clear
   Set fmsTemp = GridEX1.FormatStyles.add("N")
   fmsTemp.ForeColor = GLB_ALERT_COLOR

   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"
GridEX1.Columns.Item(1).Visible = False
   Set Col = GridEX1.Columns.add '2
   Col.Width = 2000
   Col.Caption = MapText("หมายเลขใบยืม")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2000
   Col.Caption = MapText("วันที่ยืม")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 2500
   Col.Caption = MapText("ผู้ยืม")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 2500
   Col.Caption = MapText("รายละเอียด")
Set Col = GridEX1.Columns.add '6
   Col.Width = 2480
   Col.Caption = MapText("จำนวน")
   Col.HeaderAlignment = jgexAlignRight
   Col.TextAlignment = jgexAlignRight
   GridEX1.itemcount = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = MapText("ข้อมูลเงินยืมพนักงาน")
   pnlHeader.Caption = MapText("ข้อมูลเงินยืมพนักงาน")
   
   Call InitGrid
   
   Call InitNormalLabel(lblLendCode, MapText("หมายเลขใบยืม"))
   Call InitNormalLabel(lblFromDate, MapText("จากวันที่ยืม"))
   Call InitNormalLabel(lblToDate, MapText("ถึงวันที่ยืม"))
   Call InitNormalLabel(lblLender, MapText("ผู้ยืม"))
      Call InitNormalLabel(lblOrderBy, MapText("เรียงตาม"))
   Call InitNormalLabel(lblOrderType, MapText("เรียงจาก"))
   
   Call txtLendCode.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   Call txtLender.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
      
      Call InitCombo(cboOrderBy)
   Call InitCombo(cboOrdertype)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdClear.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา (F5)"))
   Call InitMainButton(cmdClear, MapText("เคลียร์ (F4)"))
 cmdEdit.Enabled = False
End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   m_TableName = "USER_GROUP"
   
   Set m_EmpReceivable = New CEmpReceivable
   Set m_TempEmpReceivable = New CEmpReceivable
   Set m_Employee = New CEmployee
   Set m_Rs = New ADODB.Recordset

   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''''debug.print ColIndex & " " & NewColWidth
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   RowBuffer.RowStyle = RowBuffer.Value(6)
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
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
            Call m_TempEmpReceivable.PopulateFromRS(m_Rs)

          Values(1) = m_TempEmpReceivable.EMP_RECEIVABLE_ID
   Values(2) = m_TempEmpReceivable.BORROW_NO
   Values(3) = DateToStringExt(m_TempEmpReceivable.BORROW_DATE)
   Values(4) = m_TempEmpReceivable.LONG_NAME & " " & m_TempEmpReceivable.LAST_NAME
   Values(5) = m_TempEmpReceivable.BORROW_DESC
   Values(6) = FormatNumber(m_TempEmpReceivable.BORROW_AMOUNT)


            
            
            Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

