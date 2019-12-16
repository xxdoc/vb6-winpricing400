VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAlertBox 
   BackColor       =   &H80000000&
   ClientHeight    =   8670
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   12060
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   12060
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   11130
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   19632
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboAlertBoxType 
         Height          =   315
         Left            =   2460
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   990
         Width           =   4335
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   12150
         _ExtentX        =   21431
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   5925
         Left            =   240
         TabIndex        =   12
         Top             =   1680
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   10451
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
         Column(1)       =   "frmAlertBox.frx":0000
         Column(2)       =   "frmAlertBox.frx":00C8
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAlertBox.frx":016C
         FormatStyle(2)  =   "frmAlertBox.frx":02C8
         FormatStyle(3)  =   "frmAlertBox.frx":0378
         FormatStyle(4)  =   "frmAlertBox.frx":042C
         FormatStyle(5)  =   "frmAlertBox.frx":0504
         ImageCount      =   0
         PrinterProperties=   "frmAlertBox.frx":05BC
      End
      Begin Threed.SSCheck chkCancelFlag 
         Height          =   435
         Left            =   6960
         TabIndex        =   11
         Top             =   960
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblAlertBoxType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   10
         Top             =   1050
         Width           =   2175
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   8190
         TabIndex        =   1
         Top             =   930
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   9840
         TabIndex        =   2
         Top             =   930
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3660
         TabIndex        =   5
         Top             =   7950
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   390
         TabIndex        =   3
         Top             =   7950
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   2010
         TabIndex        =   4
         Top             =   7950
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   9975
         TabIndex        =   7
         Top             =   7950
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8325
         TabIndex        =   6
         Top             =   7950
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAlertBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_AlertBox As CAlertBox
Private m_TempAlertBox As CAlertBox
Private m_Rs As ADODB.Recordset
Public OKClick As Boolean
Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean
   
   frmAddEditAlertBox.HeaderText = MapText("เพิ่มข้อมูลการแจ้งเตือน")
   frmAddEditAlertBox.ShowMode = SHOW_ADD
   Load frmAddEditAlertBox
   frmAddEditAlertBox.Show 1
   
   OKClick = frmAddEditAlertBox.OKClick
   
   Unload frmAddEditAlertBox
   Set frmAddEditAlertBox = Nothing
      
   If OKClick Then
      Call QueryData(True)
   End If
End Sub

Private Sub cmdClear_Click()
   cboAlertBoxType.ListIndex = -1
   chkCancelFlag.Value = ssCBUnchecked
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
   If Not glbDaily.DeleteAlertBox(id, IsOK, True, glbErrorLog) Then
      m_AlertBox.ALERT_BOX_ID = -1
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
               
   frmAddEditAlertBox.id = id
   frmAddEditAlertBox.HeaderText = MapText("แก้ไขข้อมูลการแจ้งเตือน")
   frmAddEditAlertBox.ShowMode = SHOW_EDIT
   Load frmAddEditAlertBox
   frmAddEditAlertBox.Show 1
   
   OKClick = frmAddEditAlertBox.OKClick
   
   Unload frmAddEditAlertBox
   Set frmAddEditAlertBox = Nothing
   
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
      
      Call InitAlertBoxType(cboAlertBoxType)
      
      Call QueryData(True)
   End If
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long
   
   If Flag Then
      Call EnableForm(Me, False)
      m_AlertBox.ALERT_BOX_TYPE = cboAlertBoxType.ItemData(Minus2Zero(cboAlertBoxType.ListIndex))
      m_AlertBox.ALERT_CANCEL_FLAG = Check2Flag(chkCancelFlag.Value)
      
      If Not glbDaily.QueryAlertBox(m_AlertBox, m_Rs, ItemCount, IsOK, glbErrorLog) Then
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
   Col.Width = 2000
   Col.Caption = MapText("จากวันที่")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2000
   Col.Caption = MapText("ถึงวันที่")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 4000
   Col.Caption = MapText("ประเภท")
  
   Set Col = GridEX1.Columns.add '6
   Col.Width = 1000
   Col.Caption = MapText("ถึงทุกคน")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 10000
   Col.Caption = "รายละเอียด"
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = 0
   Col.Caption = MapText("ยกเลิก")
   
   GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = MapText("ข้อมูลการแจ้งเตือน")
   
   pnlHeader.Caption = MapText("ข้อมูลการแจ้งเตือน")
   
   Call InitGrid
   Call InitNormalLabel(lblAlertBoxType, MapText("ประเภทการแจ้งเตือน"))
   
   Call InitCheckBox(chkCancelFlag, "ยกเลิก")
   
   Call InitCombo(cboAlertBoxType)
   
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
       
    Call InitGrid
End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   
   Set m_AlertBox = New CAlertBox
   Set m_TempAlertBox = New CAlertBox
   Set m_Rs = New ADODB.Recordset
   
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

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
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
   Call m_TempAlertBox.PopulateFromRS(1, m_Rs)
   
   Values(1) = m_TempAlertBox.ALERT_BOX_ID
   Values(2) = DateToStringExtEx2(m_TempAlertBox.ALERT_BOX_FROM)
   Values(3) = DateToStringExtEx2(m_TempAlertBox.ALERT_BOX_TO)
   Values(4) = AlertBoxType2Text(m_TempAlertBox.ALERT_BOX_TYPE)
   Values(5) = m_TempAlertBox.ALERT_ALL_FLAG
   Values(6) = m_TempAlertBox.ALERT_BOX_DESC
   Values(7) = m_TempAlertBox.ALERT_CANCEL_FLAG
   
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
