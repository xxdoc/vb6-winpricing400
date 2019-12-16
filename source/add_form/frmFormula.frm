VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmFormula 
   Caption         =   "Form1"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12810
   Icon            =   "frmFormula.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   12810
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8520
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboOrdertype 
         Height          =   315
         Left            =   6120
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1830
         Width           =   3100
      End
      Begin VB.ComboBox cboOrderBy 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1830
         Width           =   3135
      End
      Begin VB.ComboBox cboFormulaType 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1380
         Width           =   3135
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   12795
         _ExtentX        =   22569
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlDate uctlFormulaDate 
         Height          =   405
         Left            =   6120
         TabIndex        =   4
         Top             =   1380
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtFormulaNo 
         Height          =   435
         Left            =   1320
         TabIndex        =   1
         Top             =   930
         Width           =   3135
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtFormulaDesc 
         Height          =   435
         Left            =   6120
         TabIndex        =   2
         Top             =   930
         Width           =   3855
         _ExtentX        =   6535
         _ExtentY        =   767
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   5265
         Left            =   135
         TabIndex        =   9
         Top             =   2430
         Width           =   12555
         _ExtentX        =   22146
         _ExtentY        =   9287
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
         Column(1)       =   "frmFormula.frx":030A
         Column(2)       =   "frmFormula.frx":03D2
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmFormula.frx":0476
         FormatStyle(2)  =   "frmFormula.frx":05D2
         FormatStyle(3)  =   "frmFormula.frx":0682
         FormatStyle(4)  =   "frmFormula.frx":0736
         FormatStyle(5)  =   "frmFormula.frx":080E
         ImageCount      =   0
         PrinterProperties=   "frmFormula.frx":08C6
      End
      Begin Threed.SSCheck chkCancelFlag 
         Height          =   375
         Left            =   9360
         TabIndex        =   23
         Top             =   1800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin Threed.SSCommand cmdOther 
         Height          =   525
         Left            =   7800
         TabIndex        =   22
         Top             =   7800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "lblOrderBy"
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   1950
         Width           =   1125
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "lblOrderType"
         Height          =   315
         Left            =   4890
         TabIndex        =   20
         Top             =   1920
         Width           =   1185
      End
      Begin VB.Label lblFormulaDate 
         Alignment       =   1  'Right Justify
         Caption         =   "lblFormulaDate"
         Height          =   315
         Left            =   4500
         TabIndex        =   19
         Top             =   1500
         Width           =   1575
      End
      Begin VB.Label lblFormulaType 
         Alignment       =   1  'Right Justify
         Caption         =   "lblFormulaType"
         Height          =   315
         Left            =   150
         TabIndex        =   18
         Top             =   1500
         Width           =   1125
      End
      Begin VB.Label lblFormulaDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "lblFormulaDesc"
         Height          =   315
         Left            =   4500
         TabIndex        =   17
         Top             =   1050
         Width           =   1575
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   11160
         TabIndex        =   8
         Top             =   1560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   11160
         TabIndex        =   7
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmFormula.frx":0A9E
         ButtonStyle     =   3
      End
      Begin VB.Label lblFormulaNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblFormulaNo"
         Height          =   315
         Left            =   30
         TabIndex        =   16
         Top             =   1050
         Width           =   1245
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   9465
         TabIndex        =   13
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
         Left            =   11070
         TabIndex        =   14
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   12
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
Attribute VB_Name = "frmFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_Formula As CFormula
Private m_TempFormula As CFormula
Private m_Rs As ADODB.Recordset
Private m_TableName As String

Public OKClick As Boolean

Private Sub cmdOther_Click()
   Load frmImportPlcFormula
   frmImportPlcFormula.Show 1
   
   Unload frmImportPlcFormula
   Set frmImportPlcFormula = Nothing
End Sub

Private Sub cmdSearch_Click()
   Call QueryData(True)
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long

   If Flag Then
      Call EnableForm(Me, False)
            
      m_Formula.FORMULA_NO = txtFormulaNo.Text
      m_Formula.PART_NO = txtFormulaDesc.Text
      m_Formula.FORMULA_DATE = uctlFormulaDate.ShowDate
      m_Formula.FORMULA_TYPE = cboFormulaType.ItemData(Minus2Zero(cboFormulaType.ListIndex))
      m_Formula.CANCEL_FLAG = Check2Flag(chkCancelFlag.Value)
       m_Formula.OrderBy = cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex))
      m_Formula.OrderType = cboOrdertype.ItemData(Minus2Zero(cboOrdertype.ListIndex))
      If Not glbProduction.QueryFormula(m_Formula, m_Rs, ItemCount, IsOK, glbErrorLog) Then
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
   If Not VerifyAccessRight("PRODUCT_FORMULA_ADD") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   frmAddEditFormulaMain.HeaderText = MapText("เพิ่มข้อมูลสูตรการผลิต")
   frmAddEditFormulaMain.ShowMode = SHOW_ADD
   Load frmAddEditFormulaMain
   frmAddEditFormulaMain.Show 1
   
   OKClick = frmAddEditFormulaMain.OKClick
   
   Unload frmAddEditFormulaMain
   Set frmAddEditFormulaMain = Nothing
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub


Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim id As Long
   If Not VerifyAccessRight("PRODUCT_FORMULA_DELETE") Then
      Call EnableForm(Me, True)
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
   If Not glbProduction.DeleteFormula(id, IsOK, True, glbErrorLog) Then
      m_Formula.FORMULA_ID = -1
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
   
   frmAddEditFormulaMain.id = id
   frmAddEditFormulaMain.HeaderText = MapText("แก้ไขข้อมูลสูตรการผลิต")
   frmAddEditFormulaMain.ShowMode = SHOW_EDIT
   Load frmAddEditFormulaMain
   frmAddEditFormulaMain.Show 1
   
   OKClick = frmAddEditFormulaMain.OKClick
   
   Unload frmAddEditFormulaMain
   Set frmAddEditFormulaMain = Nothing
               
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
      
      Call LoadFormulaType(cboFormulaType)
            
      Call InitFormulaOrderBy(cboOrderBy)
      Call InitOrderType(cboOrdertype)
      
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
   
   Set m_Formula = Nothing
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
   Col.Width = 2085
   Col.Caption = "รหัสสูตร"

   Set Col = GridEX1.Columns.add '3
   Col.Width = 4320
   Col.Caption = MapText("รายละเอียดสูตร")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 2475
   Col.Caption = MapText("วันที่สร้างสูตร")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 2685
   Col.Caption = MapText("ประเภทสูตร")
   
   Set Col = GridEX1.Columns.add '6
   Col.Width = 3500
   Col.Caption = MapText("ผู้สร้างสูตร")
 
   Set Col = GridEX1.Columns.add '7
   Col.Width = 3500
   Col.Caption = MapText("ผลิตภัณฑ์")
   
   GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = MapText("สูตรการผลิต")
   pnlHeader.Caption = MapText("สูตรการผลิต")
   
   Call InitNormalLabel(lblFormulaNo, MapText("รหัสสูตร"))
   Call InitNormalLabel(lblFormulaDesc, MapText("รหัสสินค้า"))
   Call InitNormalLabel(lblFormulaDate, MapText("วันที่สูตร"))
   Call InitNormalLabel(lblFormulaType, MapText("ประเภทสูตร"))
   Call InitNormalLabel(lblOrderBy, MapText("เรียงตาม"))
   Call InitNormalLabel(lblOrderType, MapText("เรียงจาก"))
   
   Call txtFormulaNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtFormulaDesc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   
   Call InitCombo(cboFormulaType)
   Call InitCombo(cboOrderBy)
   Call InitCombo(cboOrdertype)
   
   Call InitCheckBox(chkCancelFlag, "ยกเลิก")
   
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
   
   cmdOther.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา (F5)"))
   Call InitMainButton(cmdClear, MapText("เคลียร์ (F4)"))
   
   Call InitMainButton(cmdOther, MapText("IMPORT"))
   
   Call InitGrid1
End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   m_TableName = "USER_GROUP"
   Call InitFormLayout
      
   Set m_Rs = New ADODB.Recordset
   Set m_Formula = New CFormula
   Set m_TempFormula = New CFormula
   Call EnableForm(Me, True)
End Sub
Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
Dim TempID1 As Long
Dim BD As CFormula
Dim IsOK As Boolean
Dim OKClick As Boolean

   If GridEX1.ItemCount <= 0 Then
         Exit Sub
   End If
   
   TempID1 = GridEX1.Value(1)
   If Button = 2 Then
      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.Popup("คัดลอกข้อมูล")
      If lMenuChosen = 0 Then
         Exit Sub
      End If
      Set oMenu = Nothing
   Else
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   If lMenuChosen = 1 Then
         If Not VerifyAccessRight("PRODUCT_FORMULA_ADD") Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
      Set BD = New CFormula
      BD.FORMULA_ID = TempID1
      Call glbProduction.CopyFormula(BD, IsOK, True, 1, glbErrorLog)
      Call QueryData(True)
      Set BD = Nothing
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
      RowBuffer.RowStyle = RowBuffer.Value(7)
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
   Call m_TempFormula.PopulateFromRS(1, m_Rs)
   
   Values(1) = m_TempFormula.FORMULA_ID
   Values(2) = m_TempFormula.FORMULA_NO
   Values(3) = m_TempFormula.FORMULA_DESC
   Values(4) = DateToStringExtEx2(m_TempFormula.FORMULA_DATE)
   Values(5) = m_TempFormula.FORMULA_TYPE_NAME
   Values(6) = m_TempFormula.LONG_NAME & " " & m_TempFormula.LAST_NAME
   Values(7) = m_TempFormula.PART_ITEM_NAME
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub cmdClear_Click()
   txtFormulaNo.Text = ""
   txtFormulaDesc.Text = ""
   cboOrderBy.ListIndex = -1
   cboOrdertype.ListIndex = -1
   uctlFormulaDate.ShowDate = -1
   cboFormulaType.ListIndex = -1
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
   cmdOther.Top = ScaleHeight - 580
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
   cmdOther.Left = cmdOK.Left - cmdOther.Width - 50
End Sub
