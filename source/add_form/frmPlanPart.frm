VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmPlanPart 
   BackColor       =   &H80000000&
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   11910
   Icon            =   "frmPlanPart.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlDate uctlFromPlanDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   0
         Top             =   840
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin VB.ComboBox cboOrderType 
         Height          =   315
         Left            =   6120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1890
         Width           =   2595
      End
      Begin VB.ComboBox cboOrderBy 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1890
         Width           =   2985
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   12
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   5175
         Left            =   180
         TabIndex        =   5
         Top             =   2520
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   9128
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
         Column(1)       =   "frmPlanPart.frx":27A2
         Column(2)       =   "frmPlanPart.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmPlanPart.frx":290E
         FormatStyle(2)  =   "frmPlanPart.frx":2A6A
         FormatStyle(3)  =   "frmPlanPart.frx":2B1A
         FormatStyle(4)  =   "frmPlanPart.frx":2BCE
         FormatStyle(5)  =   "frmPlanPart.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmPlanPart.frx":2D5E
      End
      Begin prjFarmManagement.uctlDate uctlToPlanDate 
         Height          =   405
         Left            =   6120
         TabIndex        =   16
         Top             =   840
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtFromPartNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   17
         Top             =   1320
         Width           =   2985
         _ExtentX        =   4630
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtToPartNo 
         Height          =   435
         Left            =   6120
         TabIndex        =   19
         Top             =   1320
         Width           =   2985
         _ExtentX        =   4630
         _ExtentY        =   767
      End
      Begin VB.Label lblToPartNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   5160
         TabIndex        =   20
         Top             =   1380
         Width           =   885
      End
      Begin VB.Label lblFromPartNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   300
         TabIndex        =   18
         Top             =   1380
         Width           =   1485
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4980
         TabIndex        =   15
         Top             =   1950
         Width           =   1095
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   300
         TabIndex        =   14
         Top             =   870
         Width           =   1485
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   13
         Top             =   1950
         Width           =   1755
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   10230
         TabIndex        =   3
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmPlanPart.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   10230
         TabIndex        =   4
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
         TabIndex        =   8
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmPlanPart.frx":3250
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
         MouseIcon       =   "frmPlanPart.frx":356A
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
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10095
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmPlanPart.frx":3884
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmPlanPart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_PlanPart As CPlanPart
Private m_TempPlanPart As CPlanPart
Private m_Rs As ADODB.Recordset
Private m_TableName As String

Public Area As Long
Public HeaderText As String
Public OKClick As Boolean

Private m_IndexCollections As Collection
Private m_IndexCollectionsSearch As Collection
Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean
   
   frmAddEditPlanPart.HeaderText = MapText("เพิ่ม") & " " & HeaderText
   frmAddEditPlanPart.Area = Area
   frmAddEditPlanPart.ShowMode = SHOW_ADD
   Load frmAddEditPlanPart
   frmAddEditPlanPart.Show 1
   
   OKClick = frmAddEditPlanPart.OKClick
   
   Unload frmAddEditPlanPart
   Set frmAddEditPlanPart = Nothing
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub

Private Sub cmdClear_Click()
   uctlFromPlanDate.ShowDate = -1
   uctlToPlanDate.ShowDate = -1
   txtFromPartNo.Text = ""
   txtToPartNo.Text = ""
   cboOrderBy.ListIndex = -1
   cboOrderType.ListIndex = -1
End Sub
Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim id As Long

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
   If Not glbDaily.DeletePlanPart(id, IsOK, True, glbErrorLog) Then
      m_PlanPart.PLAN_PART_ID = -1
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
   
   frmAddEditPlanPart.id = id
   frmAddEditPlanPart.HeaderText = MapText("แก้ไข") & " " & HeaderText
   Set frmAddEditPlanPart.m_IndexCollections = m_IndexCollections
   frmAddEditPlanPart.CurrentIndex = m_IndexCollectionsSearch(Trim(str(id)))
   frmAddEditPlanPart.Area = Area
   frmAddEditPlanPart.ShowMode = SHOW_EDIT
   Load frmAddEditPlanPart
   frmAddEditPlanPart.Show 1
   
   OKClick = frmAddEditPlanPart.OKClick
   
   Unload frmAddEditPlanPart
   Set frmAddEditPlanPart = Nothing
               
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
      
      Call InitPlanPartOrderBy(cboOrderBy)
      Call InitOrderType(cboOrderType)
      
      uctlFromPlanDate.ShowDate = Now
      uctlToPlanDate.ShowDate = Now
      
      Call QueryData(True)
   End If
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      m_PlanPart.FROM_DATE = uctlFromPlanDate.ShowDate
      m_PlanPart.TO_DATE = uctlToPlanDate.ShowDate
      m_PlanPart.PLAN_AREA = Area
      m_PlanPart.FROM_PART_NO = txtFromPartNo.Text
      m_PlanPart.TO_PART_NO = txtToPartNo.Text
      
      m_PlanPart.OrderBy = cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex))
      m_PlanPart.OrderType = cboOrderType.ItemData(Minus2Zero(cboOrderType.ListIndex))
      
      If Not glbDaily.QueryPlanPart(m_PlanPart, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If (ItemCount <> m_IndexCollectionsSearch.Count) Then
      Set m_IndexCollections = New Collection
      Set m_IndexCollectionsSearch = New Collection
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
   Col.Width = 1500
   Col.Caption = MapText("วันที่")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 1000
   Col.Caption = MapText("รหัส")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 2500
   If Area = 3 Then
      Col.Caption = MapText("ผลิตภัณฑ์")
   Else
      Col.Caption = MapText("วัตถุดิบ")
   End If
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 1000
   Col.Caption = MapText("รับเข้า")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.add '6
   Col.Width = 1000
   Col.Caption = MapText("เบิกใช้")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = 5250
   Col.Caption = "หมายเหตุ"
   
   Set Col = GridEX1.Columns.add '8
   Col.Width = 1000
   Col.Caption = "ยกเลิก"
   
   Set Col = GridEX1.Columns.add '9
   Col.Width = 2265
   Col.Caption = MapText("วันที่ปรับปรุงล่าสุด")
   
   GridEX1.ItemCount = 0
End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitGrid
   
   Call InitNormalLabel(lblDocumentDate, MapText("วันที่"))
   Call InitNormalLabel(lblFromPartNo, MapText("จากรหัสวัตถดิบ"))
   Call InitNormalLabel(lblToPartNo, MapText("ถึง"))
   Call InitNormalLabel(lblOrderBy, MapText("เรียงตาม"))
   Call InitNormalLabel(lblOrderType, MapText("เรียงจาก"))
   
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
   m_TableName = "PLAN_PART"
   
   Set m_PlanPart = New CPlanPart
   Set m_TempPlanPart = New CPlanPart
   Set m_Rs = New ADODB.Recordset
   
   Set m_IndexCollections = New Collection
   Set m_IndexCollectionsSearch = New Collection
   
   Call InitFormLayout
   Call EnableForm(Me, True)
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

Private Sub Form_Unload(Cancel As Integer)
   Set m_IndexCollections = Nothing
   Set m_IndexCollectionsSearch = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   RowBuffer.RowStyle = RowBuffer.Value(8)
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
   Call m_TempPlanPart.PopulateFromRS(1, m_Rs)

   Values(1) = m_TempPlanPart.PLAN_PART_ID
   Values(2) = DateToStringExtEx2(m_TempPlanPart.PLAN_DATE)
   Values(3) = m_TempPlanPart.PART_NO
   Values(4) = m_TempPlanPart.PART_DESC
   Values(5) = m_TempPlanPart.PLAN_IN
   Values(6) = m_TempPlanPart.PLAN_OUT
   Values(7) = m_TempPlanPart.NOTE
   Values(8) = m_TempPlanPart.CANCEL_FLAG
   Values(9) = DateToStringExtEx3(m_TempPlanPart.MODIFY_DATE)
   
   Call AddDataCollection(m_TempPlanPart.PLAN_PART_ID)
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub AddDataCollection(id As Long)
On Error Resume Next
   Dim Temp As Long
   Temp = m_IndexCollectionsSearch(Trim(str(id)))
   If (Temp <= 0) Then
      Call m_IndexCollectionsSearch.add(m_IndexCollectionsSearch.Count + 1, Trim(str(id)))
      Call m_IndexCollections.add(id, Trim(str(m_IndexCollectionsSearch.Count)))
   End If
End Sub
