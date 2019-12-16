VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddEditAlertBox 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13650
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   10650
   ScaleWidth      =   13650
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   11130
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   19632
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboAlertBoxType 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1950
         Width           =   4335
      End
      Begin prjFarmManagement.uctlDate uctlAlertBoxTo 
         Height          =   405
         Left            =   7800
         TabIndex        =   1
         Top             =   1050
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   6
         Top             =   2640
         Width           =   13395
         _ExtentX        =   23627
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
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   15360
         _ExtentX        =   27093
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtAlertBoxDesc 
         Height          =   435
         Left            =   2160
         TabIndex        =   2
         Top             =   1500
         Width           =   11355
         _ExtentX        =   3254
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlAlertBoxFrom 
         Height          =   405
         Left            =   2160
         TabIndex        =   0
         Top             =   1080
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   6615
         Left            =   120
         TabIndex        =   18
         Top             =   3120
         Width           =   13425
         _ExtentX        =   23680
         _ExtentY        =   11668
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         TabKeyBehavior  =   1
         MultiSelect     =   -1  'True
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
         Column(1)       =   "frmAddEditAlertBox.frx":0000
         Column(2)       =   "frmAddEditAlertBox.frx":00C8
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditAlertBox.frx":016C
         FormatStyle(2)  =   "frmAddEditAlertBox.frx":02C8
         FormatStyle(3)  =   "frmAddEditAlertBox.frx":0378
         FormatStyle(4)  =   "frmAddEditAlertBox.frx":042C
         FormatStyle(5)  =   "frmAddEditAlertBox.frx":0504
         ImageCount      =   0
         PrinterProperties=   "frmAddEditAlertBox.frx":05BC
      End
      Begin Threed.SSCheck chkAllFlag 
         Height          =   435
         Left            =   7920
         TabIndex        =   4
         Top             =   1920
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
         Left            =   600
         TabIndex        =   17
         Top             =   2010
         Width           =   1500
      End
      Begin Threed.SSCheck chkCancelFlag 
         Height          =   435
         Left            =   10680
         TabIndex        =   5
         Top             =   1920
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblAlertBoxFrom 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   600
         TabIndex        =   16
         Top             =   1170
         Width           =   1500
      End
      Begin VB.Label lblAlertBoxTo 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   6240
         TabIndex        =   15
         Top             =   1140
         Width           =   1485
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   525
         Left            =   8610
         TabIndex        =   9
         Top             =   9990
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   10275
         TabIndex        =   10
         Top             =   9990
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
         Left            =   11925
         TabIndex        =   11
         Top             =   9990
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
         Top             =   9990
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   1800
         TabIndex        =   8
         Top             =   9990
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin VB.Label lblAlertBoxDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   315
         Left            =   600
         TabIndex        =   13
         Top             =   1650
         Width           =   1500
      End
   End
End
Attribute VB_Name = "frmAddEditAlertBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_AlertBox As CAlertBox

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public id As Long
Private Sub cboAlertBoxType_Click()
   m_HasModify = True
End Sub

Private Sub chkAllFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkCancelFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      
      Call InitAlertBoxType(cboAlertBoxType)
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_AlertBox.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         uctlAlertBoxFrom.ShowDate = Now
         uctlAlertBoxTo.ShowDate = Now
         m_AlertBox.QueryFlag = 0
         'Call QueryData(False)
      End If

      Call EnableForm(Me, True)
      m_HasModify = False
   End If
End Sub

Private Sub TabStrip1_Click()
   If TabStrip1.SelectedItem.Index = 1 Then
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub


Private Sub Form_Load()
   OKClick = False
   Call InitFormLayout
      
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_AlertBox = New CAlertBox
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
      
      
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblAlertBoxFrom, MapText("จากวันที่"))
   Call InitNormalLabel(lblAlertBoxTo, MapText("ถึงวันที่"))
   
   Call InitNormalLabel(lblAlertBoxDesc, MapText("รายละเอียด"))
   Call InitNormalLabel(lblAlertBoxType, MapText("ประเภท"))
   
   Call txtAlertBoxDesc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   
   Call InitCheckBox(chkCancelFlag, "ยกเลิก")
   Call InitCheckBox(chkAllFlag, "ส่งถึงทุกคน")
   
   Call InitCombo(cboAlertBoxType)
   
   
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSave.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSave, MapText("บันทึก (F10)"))
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.NAME = GLB_FONT
   TabStrip1.Font.Size = 16
   
   Call InitGrid1
   
   TabStrip1.Tabs.Clear
   
   TabStrip1.Tabs.add().Caption = MapText("รายชื่อผู้ที่ต้องการแจ้งเตือน")
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
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   Set Col = GridEX1.Columns.add '3
   Col.Width = 4000
   Col.Caption = MapText("รหัสผู้ใช้งาน")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 1000
   Col.Caption = MapText("อ่านแล้ว")
End Sub
Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
      m_AlertBox.ALERT_BOX_ID = id
      If Not glbDaily.QueryAlertBox(m_AlertBox, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_AlertBox.PopulateFromRS(1, m_Rs)
      
      uctlAlertBoxFrom.ShowDate = m_AlertBox.ALERT_BOX_FROM
      uctlAlertBoxTo.ShowDate = m_AlertBox.ALERT_BOX_TO
      txtAlertBoxDesc.Text = m_AlertBox.ALERT_BOX_DESC
      cboAlertBoxType.ListIndex = IDToListIndex(cboAlertBoxType, m_AlertBox.ALERT_BOX_TYPE)
      chkAllFlag.Value = FlagToCheck(m_AlertBox.ALERT_ALL_FLAG)
      chkCancelFlag.Value = FlagToCheck(m_AlertBox.ALERT_CANCEL_FLAG)
      
      GridEX1.ItemCount = CountItem(m_AlertBox.CollAlertDetail)
      GridEX1.Rebind
   End If
      
      
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub txtAlertBoxDesc_Change()
   m_HasModify = True
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim ItemCount As Long

   If Not VerifyDate(lblAlertBoxFrom, uctlAlertBoxFrom, False) Then
      Exit Function
   End If
   
   If Not VerifyDate(lblAlertBoxFrom, uctlAlertBoxFrom, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblAlertBoxDesc, txtAlertBoxDesc, False) Then
      Exit Function
   End If
   
   If Not VerifyCombo(lblAlertBoxType, cboAlertBoxType, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_AlertBox.AddEditMode = ShowMode
   m_AlertBox.ALERT_BOX_FROM = uctlAlertBoxFrom.ShowDate
   m_AlertBox.ALERT_BOX_TO = uctlAlertBoxTo.ShowDate
   m_AlertBox.ALERT_BOX_DESC = txtAlertBoxDesc.Text
   
   m_AlertBox.ALERT_BOX_TYPE = cboAlertBoxType.ItemData(Minus2Zero(cboAlertBoxType.ListIndex))
   m_AlertBox.ALERT_ALL_FLAG = Check2Flag(chkAllFlag.Value)
   m_AlertBox.ALERT_CANCEL_FLAG = Check2Flag(chkCancelFlag.Value)
      
   Call EnableForm(Me, False)
   
   If Not glbDaily.AddEditAlertBox(m_AlertBox, IsOK, True, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
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
Dim OKClick As Boolean
   
   OKClick = False
   If TabStrip1.SelectedItem.Index = 1 Then
      Set frmAddAlertDetail.TempCollection = m_AlertBox.CollAlertDetail
      frmAddAlertDetail.ShowMode = SHOW_ADD
      frmAddAlertDetail.HeaderText = MapText("เลือกบัญชีผู้ใช้")
      Load frmAddAlertDetail
      frmAddAlertDetail.Show 1
      
      OKClick = frmAddAlertDetail.OKClick
      
      Unload frmAddAlertDetail
      Set frmAddAlertDetail = Nothing
      
      If OKClick Then
         GridEX1.ItemCount = CountItem(m_AlertBox.CollAlertDetail)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdDelete_Click()
Dim ID1 As Long
Dim ID2 As Long

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
         m_AlertBox.CollAlertDetail.Remove (ID2)
      Else
         m_AlertBox.CollAlertDetail.Item(ID2).Flag = "D"
      End If
      
      GridEX1.ItemCount = CountItem(m_AlertBox.CollAlertDetail)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
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
   If Not SaveData Then
      Exit Sub
   End If
   
   ShowMode = SHOW_EDIT
   id = m_AlertBox.ALERT_BOX_ID
   m_AlertBox.QueryFlag = 1
   QueryData (True)
   m_HasModify = False
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
      'Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
      KeyCode = 0
   End If
End Sub
Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
End Sub
Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Index = 1 Then
      If m_AlertBox.CollAlertDetail Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If
'
      Dim CR As CAlertDetail
      If m_AlertBox.CollAlertDetail.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_AlertBox.CollAlertDetail, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If
      
      Values(1) = CR.ALERT_DETAIL_ID
      Values(2) = RealIndex
      Values(3) = CR.USER_NAME
      Values(4) = CR.READ_FLAG
      
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
      
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub GridEX1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
Dim TempID1 As Long
Dim TempID2 As Long
Dim IsOK As Boolean
Dim OKClick As Boolean
Dim ItemCount As Long
Dim Status As String
   
   If GridEX1.ItemCount <= 0 Then
      Exit Sub
   End If
   
   TempID1 = GridEX1.Value(1)
   TempID2 = GridEX1.Value(2)
   
   If Button = 2 Then
      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.Popup("อ่านแล้ว", "-", "ยังไม่อ่าน")
      If lMenuChosen = 0 Then
         Exit Sub
      End If
      
      If lMenuChosen = 0 Then
         Exit Sub
      End If
      Set oMenu = Nothing
   Else
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   
    Dim Ald As CAlertDetail
    Set Ald = m_AlertBox.CollAlertDetail(TempID2)
    If Ald.Flag <> "A" Then
      Ald.Flag = "E"
   End If
   
   If lMenuChosen = 1 Then
      Ald.READ_FLAG = "Y"
   ElseIf lMenuChosen = 3 Then
      Ald.READ_FLAG = "N"
   End If
      
   GridEX1.ItemCount = CountItem(m_AlertBox.CollAlertDetail)
   GridEX1.Rebind
  
   Call EnableForm(Me, True)
   m_HasModify = True
   
End Sub

Private Sub uctlAlertBoxFrom_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlAlertBoxTo_HasChange()
   m_HasModify = True
End Sub
