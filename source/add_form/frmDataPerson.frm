VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmDataPerson 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmDataPerson.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboStatus 
      Height          =   315
      Left            =   6600
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2160
      Width           =   2955
   End
   Begin VB.ComboBox cboSex 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2160
      Width           =   2955
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboPosition 
         Height          =   315
         Left            =   6600
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1080
         Width           =   2955
      End
      Begin VB.ComboBox cboOrderType 
         Height          =   315
         Left            =   6600
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2640
         Width           =   2955
      End
      Begin VB.ComboBox cboOrderBy 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2640
         Width           =   2955
      End
      Begin prjFarmManagement.uctlTextBox txtEmpName 
         Height          =   435
         Left            =   1560
         TabIndex        =   5
         Top             =   1530
         Width           =   2985
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   4215
         Left            =   180
         TabIndex        =   0
         Top             =   3480
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   7435
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
         Column(1)       =   "frmDataPerson.frx":27A2
         Column(2)       =   "frmDataPerson.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmDataPerson.frx":290E
         FormatStyle(2)  =   "frmDataPerson.frx":2A6A
         FormatStyle(3)  =   "frmDataPerson.frx":2B1A
         FormatStyle(4)  =   "frmDataPerson.frx":2BCE
         FormatStyle(5)  =   "frmDataPerson.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmDataPerson.frx":2D5E
      End
      Begin prjFarmManagement.uctlTextBox txtEmpCode 
         Height          =   435
         Left            =   1560
         TabIndex        =   3
         Top             =   1080
         Width           =   2985
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtLastName 
         Height          =   435
         Left            =   6600
         TabIndex        =   6
         Top             =   1560
         Width           =   2955
         _ExtentX        =   5265
         _ExtentY        =   767
      End
      Begin VB.Label lblStatus 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4680
         TabIndex        =   25
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label lblSex 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   0
         TabIndex        =   24
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label lblLastName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   5040
         TabIndex        =   23
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label lblEmpCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   22
         Top             =   1140
         Width           =   1455
      End
      Begin VB.Label lblPosition 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   5040
         TabIndex        =   21
         Top             =   1140
         Width           =   1455
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   5160
         TabIndex        =   20
         Top             =   2760
         Width           =   1365
      End
      Begin VB.Label lblEmpName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   19
         Top             =   1590
         Width           =   1455
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   18
         Top             =   2760
         Width           =   1455
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   10110
         TabIndex        =   11
         Top             =   1170
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmDataPerson.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   10110
         TabIndex        =   12
         Top             =   1740
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   15
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmDataPerson.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   13
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmDataPerson.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   14
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
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmDataPerson.frx":3884
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmDataPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_Employee As CEmployee
Private m_TempEmployee As CEmployee
Private m_Rs As ADODB.Recordset
Private m_TableName As String

Public OKClick As Boolean
Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean

   frmAddEditPerson.HeaderText = MapText("เพิ่มข้อมูลพนักงาน")
   frmAddEditPerson.ShowMode = SHOW_ADD
   Load frmAddEditPerson
   frmAddEditPerson.Show 1
   
   OKClick = frmAddEditPerson.OKClick
   
   Unload frmAddEditPerson
   Set frmAddEditPerson = Nothing
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub

Private Sub cmdClear_Click()
   txtEmpName.Text = ""
   txtEmpCode.Text = ""
   txtLastName.Text = ""
   cboOrderBy.ListIndex = -1
   cboOrderType.ListIndex = -1
   cboPosition.ListIndex = -1
   cboSex.ListIndex = -1
   cboStatus.ListIndex = -1
Call QueryData(True)
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
   If Not glbDaily.DeleteEmployee(id, IsOK, True, glbErrorLog) Then
      m_Employee.EMP_ID = -1
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
   
   frmAddEditPerson.id = id
   frmAddEditPerson.HeaderText = MapText("แก้ไขข้อมูลพนักงาน")
   frmAddEditPerson.ShowMode = SHOW_EDIT
   Load frmAddEditPerson
   frmAddEditPerson.Show 1
   
   OKClick = frmAddEditPerson.OKClick
   
   Unload frmAddEditPerson
   Set frmAddEditPerson = Nothing
               
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
      
      Call LoadPosition(cboPosition)
      Call LoadSex(cboSex)
            Call LoadWorkStatus(cboStatus)
      Call InitEmployeeOrderBy(cboOrderBy)
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
      
      m_Employee.EMP_CODE = txtEmpCode.Text
      m_Employee.NAME = txtEmpName.Text
      m_Employee.LASTNAME = txtLastName.Text
      m_Employee.OrderBy = cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex))
      m_Employee.OrderType = cboOrderType.ItemData(Minus2Zero(cboOrderType.ListIndex))
      m_Employee.CURRENT_POSITION = cboPosition.ItemData(Minus2Zero(cboPosition.ListIndex))
m_Employee.SEX_ID = cboSex.ItemData(Minus2Zero(cboSex.ListIndex))
m_Employee.WORK_STATUS_ID = cboStatus.ItemData(Minus2Zero(cboStatus.ListIndex))
      If Not glbDaily.QueryEmployee(m_Employee, m_Rs, ItemCount, IsOK, glbErrorLog) Then
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
'Dim fmsTemp2 As JSFormatStyle

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   
   GridEX1.FormatStyles.Clear
   Set fmsTemp = GridEX1.FormatStyles.add("N")
   'Set fmsTemp2 = fmsTemp
   fmsTemp.ForeColor = GLB_ALERT_COLOR
   'fmsTemp.ForeColor = BLACK
   
   'Set fmsTemp3 = GridEX1.FormatStyles.Add("Y")

   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"
   GridEX1.Columns.Item(1).Visible = False
   Set Col = GridEX1.Columns.add '2
   Col.Width = 1620
   Col.Caption = MapText("รหัสพนักงาน")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 3780
   Col.Caption = MapText("ชื่อ")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 3555
   Col.Caption = MapText("นามสกุล")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 2520
   Col.Caption = MapText("ตำแหน่ง")
Set Col = GridEX1.Columns.add '6
   Col.Width = 2520
   Col.Caption = MapText("วันเกิด")
   
   GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = MapText("ข้อมูลพนักงาน")
   pnlHeader.Caption = MapText("ข้อมูลพนักงาน")
   
   Call InitGrid
   
   Call InitNormalLabel(lblEmpName, MapText("ชื่อ"))
   Call InitNormalLabel(lblLastName, MapText("นามสกุล"))
   Call InitNormalLabel(lblPosition, MapText("ตำแหน่ง"))
   Call InitNormalLabel(lblEmpCode, MapText("รหัสพนักงาน"))
   Call InitNormalLabel(lblSex, MapText("เพศ"))
   Call InitNormalLabel(lblStatus, MapText("สถานะการทำงาน"))
   Call InitNormalLabel(lblOrderBy, MapText("เรียงตาม"))
   Call InitNormalLabel(lblOrderType, MapText("เรียงจาก"))
   
   Call txtEmpName.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   
   Call InitCombo(cboPosition)
   Call InitCombo(cboOrderBy)
   Call InitCombo(cboOrderType)
    Call InitCombo(cboSex)
   Call InitCombo(cboStatus)
   
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
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา (F5)"))
   Call InitMainButton(cmdClear, MapText("เคลียร์ (F4)"))
End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   m_TableName = "USER_GROUP"
   
   Set m_Employee = New CEmployee
   Set m_TempEmployee = New CEmployee
   Set m_Rs = New ADODB.Recordset

   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''''''Debug.Print ColIndex & " " & NewColWidth
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
   Call m_TempEmployee.PopulateFromRS(1, m_Rs)
   
   Values(1) = m_TempEmployee.EMP_ID
   Values(2) = m_TempEmployee.EMP_CODE
   Values(3) = m_TempEmployee.NAME
   Values(4) = m_TempEmployee.LASTNAME
   Values(5) = m_TempEmployee.POSITION_NAME
   Values(6) = DateToStringExt(m_TempEmployee.BIRTH_DATE)
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

