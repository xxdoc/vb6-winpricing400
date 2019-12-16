VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmInventoryWhAct 
   Caption         =   "Form1"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11820
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   11820
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8520
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboOrdertype 
         Height          =   315
         Left            =   6660
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1560
         Width           =   2985
      End
      Begin VB.ComboBox cboOrderBy 
         Height          =   315
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1560
         Width           =   3225
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   12645
         _ExtentX        =   22304
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlDate uctlInventoryActDate 
         Height          =   405
         Left            =   1650
         TabIndex        =   0
         Top             =   1020
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   4755
         Left            =   120
         TabIndex        =   15
         Top             =   2520
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   8387
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
         Column(1)       =   "frmInventoryWhAct.frx":0000
         Column(2)       =   "frmInventoryWhAct.frx":00C8
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmInventoryWhAct.frx":016C
         FormatStyle(2)  =   "frmInventoryWhAct.frx":02C8
         FormatStyle(3)  =   "frmInventoryWhAct.frx":0378
         FormatStyle(4)  =   "frmInventoryWhAct.frx":042C
         FormatStyle(5)  =   "frmInventoryWhAct.frx":0504
         ImageCount      =   0
         PrinterProperties=   "frmInventoryWhAct.frx":05BC
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "lblOrderBy"
         Height          =   315
         Left            =   360
         TabIndex        =   14
         Top             =   1680
         Width           =   1185
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "lblOrderType"
         Height          =   315
         Left            =   5130
         TabIndex        =   13
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label lblInventoryDate 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPlanDate"
         Height          =   315
         Left            =   300
         TabIndex        =   12
         Top             =   1110
         Width           =   1245
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   9900
         TabIndex        =   4
         Top             =   1410
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   9900
         TabIndex        =   3
         Top             =   810
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8535
         TabIndex        =   8
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
         TabIndex        =   9
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   7
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
Attribute VB_Name = "frmInventoryWhAct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_InventoryWhAct As CInventoryWhAct
Private m_TempInventoryWhAct As CInventoryWhAct
Private m_Rs As ADODB.Recordset

Private m_HasModify As Boolean

Public OKClick As Boolean
Public InventoryWhActArea As Long
Public HeaderText   As String
Public ProcessID As Long

Private Sub cmdSearch_Click()
   Call QueryData(True)
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      m_InventoryWhAct.INVENTORY_WH_ACT_ID = -1
      m_InventoryWhAct.FROM_DATE = uctlInventoryActDate.ShowDate
      m_InventoryWhAct.TO_DATE = uctlInventoryActDate.ShowDate
      
      m_InventoryWhAct.OrderBy = cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex))
      m_InventoryWhAct.OrderType = cboOrderType.ItemData(Minus2Zero(cboOrderType.ListIndex))
     m_InventoryWhAct.INVENTORY_WH_ACT_AREA = InventoryWhActArea
    
    If Not glbInventoryWhAct.QueryInventoryWhAct(m_InventoryWhAct, m_Rs, ItemCount, IsOK, glbErrorLog) Then
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

   If Not cmdAdd.Enabled Then
      Exit Sub
   End If

   If Not VerifyAccessRight("INVENTORY-WH_ACTUAL_" & InventoryWhActArea2Text2(InventoryWhActArea) & "_ADD", "เพิ่ม" & InventoryWhActArea2Text(InventoryWhActArea)) Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   frmAddEditInventoryWhAct.InventoryWhActArea = InventoryWhActArea
   frmAddEditInventoryWhAct.ShowMode = SHOW_ADD
   frmAddEditInventoryWhAct.HeaderText = HeaderText
   Load frmAddEditInventoryWhAct
   frmAddEditInventoryWhAct.Show 1

   OKClick = frmAddEditInventoryWhAct.OKClick

   Unload frmAddEditInventoryWhAct
   Set frmAddEditInventoryWhAct = Nothing

   If OKClick Then
      Call QueryData(True)
   End If
End Sub
Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim id As Long
   If Not VerifyAccessRight("INVENTORY-WH_ACTUAL_" & InventoryWhActArea2Text2(InventoryWhActArea) & "_DELETE", "ลบ" & InventoryWhActArea2Text(InventoryWhActArea)) Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   id = GridEX1.Value(1)
   
   If Not ConfirmDelete(GridEX1.Value(2)) Then
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   If Not glbInventoryAct.DeleteInventoryAct(id, IsOK, True, glbErrorLog) Then
      m_InventoryWhAct.INVENTORY_WH_ACT_AREA = -1
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
    If Not VerifyAccessRight("INVENTORY-WH_ACTUAL_" & InventoryWhActArea2Text2(InventoryWhActArea) & "_EDIT", "แก้ไข" & InventoryWhActArea2Text(InventoryWhActArea)) Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   id = Val(GridEX1.Value(1))
   frmAddEditInventoryWhAct.InventoryWhActArea = InventoryWhActArea
   frmAddEditInventoryWhAct.id = id
   frmAddEditInventoryWhAct.ShowMode = SHOW_EDIT
   frmAddEditInventoryWhAct.HeaderText = HeaderText
   Load frmAddEditInventoryWhAct
   frmAddEditInventoryWhAct.Show 1
   
   OKClick = frmAddEditInventoryWhAct.OKClick
   
   Unload frmAddEditInventoryWhAct
   Set frmAddEditInventoryWhAct = Nothing
               
   If OKClick Then
      Call QueryData(True)
   End If
End Sub
Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub

Private Sub Form_Activate()
 If Not m_HasActivate Then
      m_HasActivate = True
      
      Call InitInvActOrderBy(cboOrderBy)
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
   
   Set m_InventoryWhAct = Nothing
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
   
   If InventoryWhActArea = 1 Or InventoryWhActArea = 2 Or InventoryWhActArea = 3 Then
      Set Col = GridEX1.Columns.add '2
      Col.Width = 1500
      Col.Caption = "วันที่ประมาณ"
   End If
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 10000
   Col.Caption = MapText("รายละเอียด")
   
   GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
      
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblInventoryDate, MapText("วันที่นับสต๊อก"))
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
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdClear.Picture = LoadPicture(glbParameterObj.NormalButton1)
'   cmdImport.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา (F5)"))
   Call InitMainButton(cmdClear, MapText("เคลียร์ (F4)"))
'   Call InitMainButton(cmdImport, MapText("IMPORT"))
   
  If InventoryWhActArea = 1 Or InventoryWhActArea = 2 Then
      cmdAdd.Enabled = False
  ElseIf InventoryWhActArea = 3 Then
      cmdAdd.Enabled = True
  End If
  
   Call InitGrid1
End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   Call InitFormLayout
      
   Set m_Rs = New ADODB.Recordset
   Set m_InventoryWhAct = New CInventoryWhAct
   Set m_TempInventoryWhAct = New CInventoryWhAct

   Call EnableForm(Me, True)
End Sub
Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   'RowBuffer.RowStyle = RowBuffer.Value(3)
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long
Dim I As Long

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
   Call m_TempInventoryWhAct.PopulateFromRS(1, m_Rs)
   I = 1
   Values(I) = m_TempInventoryWhAct.INVENTORY_WH_ACT_ID
   If InventoryWhActArea = 1 Or InventoryWhActArea = 2 Or InventoryWhActArea = 3 Then
      I = I + 1
      Values(I) = DateToStringExtEx2(m_TempInventoryWhAct.INVENTORY_WH_ACT_DATE)
   End If
   I = I + 1
   Values(I) = m_TempInventoryWhAct.INVENTORY_WH_ACT_DESC
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub cmdClear_Click()
   cboOrderBy.ListIndex = -1
   cboOrderType.ListIndex = -1
   uctlInventoryActDate.ShowDate = -1
End Sub

Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.HEIGHT = ScaleHeight
   
   pnlHeader.Width = ScaleWidth
   GridEX1.Width = ScaleWidth - 2 * GridEX1.Left
   GridEX1.HEIGHT = ScaleHeight - GridEX1.Top - 620
   cmdAdd.Top = ScaleHeight - 580
   cmdEdit.Top = cmdAdd.Top
   cmdDelete.Top = cmdAdd.Top
   cmdOK.Top = cmdAdd.Top
   cmdExit.Top = cmdAdd.Top
'   cmdImport.Top = cmdAdd.Top
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
'   cmdImport.Left = cmdOK.Left - cmdImport.Width - 50
End Sub
