VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmMemoNote 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmMemoNote.frx":0000
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
      TabIndex        =   19
      Top             =   0
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox uctlCreateBy 
         Height          =   315
         Left            =   6120
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   2280
         Width           =   2595
      End
      Begin VB.ComboBox uctlCreateTo 
         Height          =   315
         Left            =   6120
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   2760
         Width           =   2595
      End
      Begin VB.ComboBox cboMemoStatus 
         Height          =   315
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2760
         Width           =   2595
      End
      Begin VB.ComboBox cboMemoType 
         Height          =   315
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2280
         Width           =   2595
      End
      Begin VB.ComboBox cboOrderType 
         Height          =   315
         Left            =   6150
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   3300
         Width           =   1995
      End
      Begin VB.ComboBox cboOrderBy 
         Height          =   315
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   3300
         Width           =   2595
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   20
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   3735
         Left            =   180
         TabIndex        =   13
         Top             =   3960
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   6588
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
         Column(1)       =   "frmMemoNote.frx":27A2
         Column(2)       =   "frmMemoNote.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmMemoNote.frx":290E
         FormatStyle(2)  =   "frmMemoNote.frx":2A6A
         FormatStyle(3)  =   "frmMemoNote.frx":2B1A
         FormatStyle(4)  =   "frmMemoNote.frx":2BCE
         FormatStyle(5)  =   "frmMemoNote.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmMemoNote.frx":2D5E
      End
      Begin prjFarmManagement.uctlDate uctlFromDateCreate 
         Height          =   405
         Left            =   1560
         TabIndex        =   0
         Top             =   840
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlDate uctlToDateCreate 
         Height          =   405
         Left            =   5520
         TabIndex        =   1
         Top             =   840
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlDate uctlFromDateFinish 
         Height          =   405
         Left            =   1560
         TabIndex        =   2
         Top             =   1320
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlDate uctlToDateFinish 
         Height          =   405
         Left            =   5520
         TabIndex        =   3
         Top             =   1320
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlDate uctlFromDateFinishReal 
         Height          =   405
         Left            =   1560
         TabIndex        =   4
         Top             =   1800
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlDate uctlToDateFinishReal 
         Height          =   405
         Left            =   5520
         TabIndex        =   5
         Top             =   1800
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin Threed.SSCheck chkWarn 
         Height          =   435
         Left            =   10200
         TabIndex        =   10
         Top             =   3240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblCreateTo 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4560
         TabIndex        =   29
         Top             =   2790
         Width           =   1485
      End
      Begin VB.Label lblCreateBy 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4560
         TabIndex        =   28
         Top             =   2310
         Width           =   1485
      End
      Begin VB.Label lblMemoStatus 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   360
         TabIndex        =   27
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label lblFinishReal 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   0
         TabIndex        =   26
         Top             =   1980
         Width           =   1455
      End
      Begin VB.Label lblFinishDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   0
         TabIndex        =   25
         Top             =   1500
         Width           =   1455
      End
      Begin VB.Label lblCreateDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   24
         Top             =   1020
         Width           =   1455
      End
      Begin VB.Label lblMemoType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   360
         TabIndex        =   23
         Top             =   2340
         Width           =   1095
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4680
         TabIndex        =   22
         Top             =   3360
         Width           =   1365
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   21
         Top             =   3360
         Width           =   1455
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   10110
         TabIndex        =   11
         Top             =   930
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmMemoNote.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   10110
         TabIndex        =   12
         Top             =   1500
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   16
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmMemoNote.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   14
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmMemoNote.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   15
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
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmMemoNote.frx":3884
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmMemoNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_MemoNote As CMemoNote
Private m_TempMemoNote As CMemoNote
Private m_Rs As ADODB.Recordset
Private m_TableName As String

Public OKClick As Boolean
Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean

   frmAddEditMemoNote.HeaderText = MapText("เพิ่ม MEMO")
   frmAddEditMemoNote.ShowMode = SHOW_ADD
   Load frmAddEditMemoNote
   frmAddEditMemoNote.Show 1
   
   OKClick = frmAddEditMemoNote.OKClick
   
   Unload frmAddEditMemoNote
   Set frmAddEditMemoNote = Nothing
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub

Private Sub cmdClear_Click()
   uctlFromDateCreate.ShowDate = -1
   uctlToDateCreate.ShowDate = -1
   uctlFromDateFinish.ShowDate = -1
   uctlToDateFinish.ShowDate = -1
   uctlFromDateFinishReal.ShowDate = -1
   uctlToDateFinishReal.ShowDate = -1
   cboMemoType.ListIndex = -1
   cboMemoStatus.ListIndex = -1
   uctlCreateBy.ListIndex = -1
   uctlCreateTo.ListIndex = -1
   cboOrderBy.ListIndex = -1
   cboOrderType.ListIndex = -1
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
   
   If Not ConfirmDelete(GridEX1.Value(2)) Then
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   Call m_MemoNote.SetFieldValue("MEMO_NOTE_ID", id)
   If Not glbDaily.DeleteMemoNote(id, IsOK, True, glbErrorLog) Then
      Call m_MemoNote.SetFieldValue("EMP_ID", -1)
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call QueryData(True)
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
   
   frmAddEditMemoNote.id = id
   frmAddEditMemoNote.HeaderText = MapText("แก้ไข MEMO")
   frmAddEditMemoNote.ShowMode = SHOW_EDIT
   Load frmAddEditMemoNote
   frmAddEditMemoNote.Show 1
   
   OKClick = frmAddEditMemoNote.OKClick
   
   Unload frmAddEditMemoNote
   Set frmAddEditMemoNote = Nothing
               
   If OKClick Then
      Call QueryData(True)
   End If

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
      
      Call LoadMaster(cboMemoType)
      Call LoadMaster(cboMemoStatus)
      
      Call LoadUserAccount(uctlCreateBy)
      
      Call LoadUserAccount(uctlCreateTo)

      uctlCreateTo.ListIndex = IDToListIndex(uctlCreateTo, Trim(str(glbUser.REAL_USER_ID)))
      
      Call InitMemoNoteOrderBy(cboOrderBy)
      Call InitOrderType(cboOrderType)
      
      chkWarn.Value = ssCBChecked
      
      Call QueryData(True)
   End If
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      Call m_MemoNote.SetFieldValue("MEMO_NOTE_ID", -1)
      Call m_MemoNote.SetFieldValue("FROM_DATE_CREATE", uctlFromDateCreate.ShowDate)
      Call m_MemoNote.SetFieldValue("TO_DATE_CREATE", uctlToDateCreate.ShowDate)
      Call m_MemoNote.SetFieldValue("FROM_DATE_FINISH", uctlFromDateFinish.ShowDate)
      Call m_MemoNote.SetFieldValue("TO_DATE_FINISH", uctlToDateFinish.ShowDate)
      Call m_MemoNote.SetFieldValue("FROM_DATE_FINISH_REAL", uctlFromDateFinishReal.ShowDate)
      Call m_MemoNote.SetFieldValue("TO_DATE_FINISH_REAL", uctlToDateFinishReal.ShowDate)
      
      Call m_MemoNote.SetFieldValue("MEMO_NOTE_CREATE_BY", uctlCreateBy.ItemData(Minus2Zero(uctlCreateBy.ListIndex)))
      Call m_MemoNote.SetFieldValue("MEMO_NOTE_CREATE_TO", uctlCreateTo.ItemData(Minus2Zero(uctlCreateTo.ListIndex)))
      
      Call m_MemoNote.SetFieldValue("MEMO_NOTE_TYPE", cboMemoType.ItemData(Minus2Zero(cboMemoType.ListIndex)))
      Call m_MemoNote.SetFieldValue("MEMO_NOTE_STATUS", cboMemoStatus.ItemData(Minus2Zero(cboMemoStatus.ListIndex)))
      
      Call m_MemoNote.SetFieldValue("ORDER_BY", cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex)))
      Call m_MemoNote.SetFieldValue("ORDER_TYPE", cboOrderType.ItemData(Minus2Zero(cboOrderType.ListIndex)))
      
      Call m_MemoNote.SetFieldValue("MEMO_NOTE_WARN", Check2Flag(chkWarn.Value))
      
      If Not glbDaily.QueryMemoNote(m_MemoNote, m_Rs, ItemCount, IsOK, glbErrorLog) Then
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
   Col.Width = 1500
   Col.Caption = MapText("วันที่สร้าง")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 1500
   Col.Caption = MapText("วันกำหนดเสร็จ")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 1500
   Col.Caption = MapText("วันที่เสร็จ")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 1000
   Col.Caption = MapText("ประเภท")
   
   Set Col = GridEX1.Columns.add '6
   Col.Width = 1000
   Col.Caption = MapText("สถานะ")
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = 3000
   Col.Caption = MapText("หัวข้อ")
   
   Set Col = GridEX1.Columns.add '8
   Col.Width = 3000
   Col.Caption = MapText("สร้างโดย")
   
   Set Col = GridEX1.Columns.add '9
   Col.Width = 3000
   Col.Caption = MapText("มอบหมายให้")
   
   Set Col = GridEX1.Columns.add '10
   Col.Width = 0
   Col.Caption = MapText("เตือน")
   
   GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = MapText("ข้อมูล MEMO")
   pnlHeader.Caption = MapText("ข้อมูล MEMO")
   
   Call InitGrid
   
   Call InitNormalLabel(lblCreateDate, MapText("วันที่สร้าง"))
   Call InitNormalLabel(lblFinishDate, MapText("วันกำหนดเสร็จ"))
   Call InitNormalLabel(lblFinishReal, MapText("วันที่เสร็จ"))
   Call InitNormalLabel(lblMemoType, MapText("ประเภท"))
   Call InitNormalLabel(lblMemoStatus, MapText("สถานะ"))
   Call InitNormalLabel(lblCreateBy, MapText("สร้างโดย"))
   Call InitNormalLabel(lblCreateTo, MapText("มอบหมายให้"))
   Call InitNormalLabel(lblOrderBy, MapText("เรียงตาม"))
   Call InitNormalLabel(lblOrderType, MapText("เรียงจาก"))
   
   Call InitCombo(cboMemoType)
   Call InitCombo(cboMemoStatus)
   Call InitCombo(uctlCreateBy)
   Call InitCombo(uctlCreateTo)
   Call InitCombo(cboOrderBy)
   Call InitCombo(cboOrderType)
   
   Call InitCheckBox(chkWarn, "เตือน")
   
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
   
'   If Not VerifyAccessRight("MEMO_NOTE_QUERY_ALL") Then
'      uctlCreateTo.Enabled = False
'   End If
   
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
   
   Set m_MemoNote = New CMemoNote
   Set m_TempMemoNote = New CMemoNote
   Set m_Rs = New ADODB.Recordset
   
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      Call m_Rs.Close
   End If
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   RowBuffer.RowStyle = RowBuffer.Value(10)
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
   Call m_TempMemoNote.PopulateFromRS(1, m_Rs)
   Values(1) = m_TempMemoNote.GetFieldValue("MEMO_NOTE_ID")
   Values(2) = DateToStringExtEx2(m_TempMemoNote.GetFieldValue("MEMO_NOTE_DATE_CREATE"))
   Values(3) = DateToStringExtEx2(m_TempMemoNote.GetFieldValue("MEMO_NOTE_DATE_FINISH"))
   Values(4) = DateToStringExtEx2(m_TempMemoNote.GetFieldValue("MEMO_NOTE_DATE_FINISH_REAL"))
   Values(5) = m_TempMemoNote.GetFieldValue("MEMO_NOTE_TYPE_NAME")
   Values(6) = m_TempMemoNote.GetFieldValue("MEMO_NOTE_STATUS_NAME")
   Values(7) = m_TempMemoNote.GetFieldValue("MEMO_NOTE_SUBJECT")
   Values(8) = m_TempMemoNote.GetFieldValue("MEMO_NOTE_CREATE_BY_NAME")
   Values(9) = m_TempMemoNote.GetFieldValue("MEMO_NOTE_CREATE_TO_NAME")
   Values(10) = m_TempMemoNote.GetFieldValue("MEMO_NOTE_WARN")
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
