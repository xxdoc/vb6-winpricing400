VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmBillingDocPayment 
   BackColor       =   &H80000000&
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   11925
   Icon            =   "frmBillingDocPayment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   11925
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlDate uctlDocumentDate 
         Height          =   405
         Left            =   6180
         TabIndex        =   1
         Top             =   750
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin VB.ComboBox cboOrderType 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2070
         Width           =   2985
      End
      Begin VB.ComboBox cboOrderBy 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1680
         Width           =   2985
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   13
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtDocumentNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   720
         Width           =   2985
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtDocumentNo_JV 
         Height          =   435
         Left            =   1860
         TabIndex        =   2
         Top             =   1170
         Width           =   2985
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlToDate 
         Height          =   405
         Left            =   6180
         TabIndex        =   19
         Top             =   1200
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   4965
         Left            =   120
         TabIndex        =   21
         Top             =   2640
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   8758
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
         Column(1)       =   "frmBillingDocPayment.frx":27A2
         Column(2)       =   "frmBillingDocPayment.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmBillingDocPayment.frx":290E
         FormatStyle(2)  =   "frmBillingDocPayment.frx":2A6A
         FormatStyle(3)  =   "frmBillingDocPayment.frx":2B1A
         FormatStyle(4)  =   "frmBillingDocPayment.frx":2BCE
         FormatStyle(5)  =   "frmBillingDocPayment.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmBillingDocPayment.frx":2D5E
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4920
         TabIndex        =   20
         Top             =   1230
         Width           =   1185
      End
      Begin VB.Label lblDocumentNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   18
         Top             =   780
         Width           =   1755
      End
      Begin VB.Label lblDocumentNo_JV 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   17
         Top             =   1230
         Width           =   1755
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   16
         Top             =   2130
         Width           =   1755
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4920
         TabIndex        =   15
         Top             =   780
         Width           =   1185
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   14
         Top             =   1680
         Width           =   1755
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   10110
         TabIndex        =   5
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBillingDocPayment.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   10110
         TabIndex        =   6
         Top             =   1680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   9
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBillingDocPayment.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   7
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBillingDocPayment.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   8
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBillingDocPayment.frx":3884
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmBillingDocPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_BillingPayment As CBillingPayment
Private m_TempBillingPayment As CBillingPayment
Private m_Rs As ADODB.Recordset
Private m_TableName As String
Public OKClick As Boolean
Public DocumentType As Long

Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu
Dim TempStr As String
Dim Programowner As String

Programowner = glbParameterObj.Programowner
        
If DocumentType = 111 Or DocumentType = 112 Then
      If Not VerifyAccessRight("LEDGER_BUY" & "_" & DocumentType & "_" & "ADD", "เพิ่ม") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
  If DocumentType = 111 Then
      frmAddEditBillingPayment.DocumentType = DocumentType
      frmAddEditBillingPayment.HeaderText = MapText("เพิ่มข้อมูล PV และ JV")
      frmAddEditBillingPayment.ShowMode = SHOW_ADD
      Load frmAddEditBillingPayment
      frmAddEditBillingPayment.Show 1
      
      OKClick = frmAddEditBillingPayment.OKClick
      
      Unload frmAddEditBillingPayment
      Set frmAddEditBillingPayment = Nothing
   End If
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub

Private Sub cmdAdjust_Click()

End Sub

Private Sub cmdClear_Click()
   txtDocumentNo.Text = ""
   txtDocumentNo_JV.Text = ""
   
   uctlDocumentDate.ShowDate = -1
   uctlToDate.ShowDate = -1
      
   cboOrderBy.ListIndex = -1
   cboOrderType.ListIndex = -1
End Sub

Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim id As Long
Dim ID2 As Long
Dim PaymentID As Long
Dim str As String
   If DocumentType = 111 Or DocumentType = 112 Then
      If Not VerifyAccessRight("LEDGER_BUY" & "_" & DocumentType & "_" & "DELETE", "ลบ") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If

   If Not cmdDelete.Enabled Then
      Exit Sub
   End If

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   id = GridEX1.Value(1)
   ID2 = GridEX1.Value(4)
   
   If Len(GridEX1.Value(5)) > 0 Then
        str = GridEX1.Value(2) & " และ " & GridEX1.Value(5)
   Else
      str = GridEX1.Value(2)
   End If
 

   Call glbDatabaseMngr.LockTable(m_TableName, id, IsCanLock, glbErrorLog)
   If Not ConfirmDelete(str) Then
      Call glbDatabaseMngr.UnLockTable(m_TableName, id, IsCanLock, glbErrorLog)
      Exit Sub
   End If

   Call EnableForm(Me, False)

   If Not glbDaily.DeleteBillingPayment(id, ID2, IsOK, True, glbErrorLog) Then
      m_BillingPayment.BILLING_PAYMENT_ID = -1
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
Dim ID2 As Long
Dim OKClick As Boolean
Dim TempStr As String

   Dim Programowner As String
   Programowner = glbParameterObj.Programowner
   
   If DocumentType = 111 Or DocumentType = 112 Then
      If Not VerifyAccessRight("LEDGER_BUY" & "_" & DocumentType & "_" & "EDIT", "แก้ไข") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   id = Val(GridEX1.Value(1))
   ID2 = Val(GridEX1.Value(4))
   Call glbDatabaseMngr.LockTable(m_TableName, id, IsCanLock, glbErrorLog)
    
   If DocumentType = 111 Or DocumentType = 112 Then
      frmAddEditBillingPayment.DocumentType = DocumentType
      frmAddEditBillingPayment.id = id
      frmAddEditBillingPayment.ID2 = ID2
      frmAddEditBillingPayment.HeaderText = MapText("แก้ไขข้อมูล PV และ JV")
      frmAddEditBillingPayment.ShowMode = SHOW_EDIT
      Load frmAddEditBillingPayment
      frmAddEditBillingPayment.Show 1
      
      OKClick = frmAddEditBillingPayment.OKClick
      
      Unload frmAddEditBillingPayment
      Set frmAddEditBillingPayment = Nothing
   End If
   
   If OKClick Then
      Call QueryData(True)
   End If
   Call glbDatabaseMngr.UnLockTable(m_TableName, id, IsCanLock, glbErrorLog)

End Sub

Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub

Private Sub cmdOther_Click()

End Sub

Private Sub cmdSearch_Click()
   Call QueryData(True)
End Sub

Private Sub Form_Activate()
Dim FromDate As Date
Dim ToDate As Date
   If Not m_HasActivate Then
      m_HasActivate = True
      
      Call InitBillingPaymentOrderBy(cboOrderBy)
      Call InitOrderType(cboOrderType)
      
      Call GetFirstLastDate(Now, FromDate, ToDate)
      uctlDocumentDate.ShowDate = FromDate
      uctlToDate.ShowDate = ToDate
      
      Call QueryData(True)
   End If
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      m_BillingPayment.BILLING_PAYMENT_ID = -1
      m_BillingPayment.DOCUMENT_NO = PatchWildCard(txtDocumentNo.Text)
      m_BillingPayment.DOCUMENT_NO_JV = PatchWildCard(txtDocumentNo_JV.Text)
      m_BillingPayment.FROM_DATE = uctlDocumentDate.ShowDate
      m_BillingPayment.TO_DATE = uctlToDate.ShowDate
      m_BillingPayment.DOCUMENT_TYPE = DocumentType
      m_BillingPayment.OrderBy = cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex))
      m_BillingPayment.OrderType = cboOrderType.ItemData(Minus2Zero(cboOrderType.ListIndex))
     
      If Not glbDaily.QueryBillingPayment(m_BillingPayment, m_Rs, ItemCount, IsOK, glbErrorLog) Then
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
   Col.Width = 2000
   Col.Caption = MapText("เลขที่ใบ PV")
   
   Set Col = GridEX1.Columns.add '2
   Col.Width = 2000
   Col.Caption = MapText("วันที่ใบ PV")
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID2"
   
   Set Col = GridEX1.Columns.add '2
   Col.Width = 2000
   Col.Caption = MapText("เลขที่ใบ JV")
   
   Set Col = GridEX1.Columns.add '2
   Col.Width = 2000
   Col.Caption = MapText("วันที่ใบ JV")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 1500
   Col.Caption = MapText("เพื่อจ่ายให้")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 1500
   Col.Caption = MapText("เป็นการชำระค่า")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 1500
   Col.Caption = MapText("กำหนดจ่าย")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 1500
   Col.Caption = MapText("จำนวนเงิน")
 
   
   GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
Dim Programowner As String
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Programowner = glbParameterObj.Programowner
      
  If DocumentType = 111 Then
      Me.Caption = MapText("ใบค่าใช้จ่ายโรงงาน (เงินสดย่อย)")
   End If
   
   Call InitGrid
   
   Call InitNormalLabel(lblDocumentDate, MapText("จากวันที่"))
   Call InitNormalLabel(lblToDate, MapText("ถึงวันที่"))
 
   Call InitNormalLabel(lblDocumentNo, MapText("เลขที่เอกสาร PV"))
   Call InitNormalLabel(lblDocumentNo_JV, MapText("เลขที่เอกสาร JV"))
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
   
  pnlHeader.Caption = Me.Caption
End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   m_TableName = "USER_GROUP"
   
   Set m_BillingPayment = New CBillingPayment
   Set m_TempBillingPayment = New CBillingPayment
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

'Private Sub GridEX1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim oMenu As cPopupMenu
'Dim lMenuChosen As Long
'Dim TempID1 As Long
'Dim Bd As CBillingDoc
'Dim IsOK As Boolean
'Dim OKClick As Boolean
'
'   If GridEX1.ItemCount <= 0 Then
'         Exit Sub
'   End If
'
'   TempID1 = GridEX1.Value(1)
'   If Button = 2 Then
'      Set oMenu = New cPopupMenu
'     lMenuChosen = oMenu.Popup("คัดลอกข้อมูล")
'      If lMenuChosen = 0 Then
'         Exit Sub
'      End If
'      Set oMenu = Nothing
'   Else
'      Exit Sub
'   End If
'
'   Call EnableForm(Me, False)
'   If lMenuChosen = 1 Then
'      If Not (Area = 1) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If
'      Set Bd = New CBillingDoc
'      Bd.BILLING_DOC_ID = TempID1
'      Call glbDaily.CopyBillingDoc(Bd, IsOK, True, Area, m_IvdDocType, glbErrorLog)
'      Call QueryData(True)
'      Set Bd = Nothing
'   End If
'
'   Call EnableForm(Me, True)
'End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   RowBuffer.RowStyle = RowBuffer.Value(6)
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long
Dim fmsTemp As JSFormatStyle

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
   Call m_TempBillingPayment.PopulateFromRS(1, m_Rs)

   Values(1) = m_TempBillingPayment.BILLING_PAYMENT_ID
   Values(2) = m_TempBillingPayment.DOCUMENT_NO
   Values(3) = DateToStringExtEx2(m_TempBillingPayment.DOCUMENT_DATE)
   Values(4) = m_TempBillingPayment.BILLING_PAYMENT_ID_REF
   Values(5) = m_TempBillingPayment.DOCUMENT_NO_JV
   Values(6) = DateToStringExtEx2(m_TempBillingPayment.DOCUMENT_DATE_JV)
   Values(7) = m_TempBillingPayment.PAYMENT_TO
   Values(8) = m_TempBillingPayment.PAYMENT_COST
   Values(9) = m_TempBillingPayment.PAYMENT_DUE
   Values(10) = m_TempBillingPayment.PAYMENT_AMOUNT
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

