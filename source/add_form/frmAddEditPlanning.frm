VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditPlanning 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14865
   Icon            =   "frmAddEditPlanning.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   14865
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8895
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   14880
      _ExtentX        =   26247
      _ExtentY        =   15690
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   120
         TabIndex        =   5
         Top             =   2220
         Width           =   14640
         _ExtentX        =   25823
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
         Left            =   10
         TabIndex        =   13
         Top             =   0
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   4935
         Left            =   120
         TabIndex        =   6
         Top             =   2760
         Width           =   14640
         _ExtentX        =   25823
         _ExtentY        =   8705
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
         Column(1)       =   "frmAddEditPlanning.frx":27A2
         Column(2)       =   "frmAddEditPlanning.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditPlanning.frx":290E
         FormatStyle(2)  =   "frmAddEditPlanning.frx":2A6A
         FormatStyle(3)  =   "frmAddEditPlanning.frx":2B1A
         FormatStyle(4)  =   "frmAddEditPlanning.frx":2BCE
         FormatStyle(5)  =   "frmAddEditPlanning.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditPlanning.frx":2D5E
      End
      Begin prjFarmManagement.uctlTextBox txtPlanningDesc 
         Height          =   435
         Left            =   1800
         TabIndex        =   3
         Top             =   1770
         Width           =   6705
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlPlanningDate 
         Height          =   405
         Left            =   1800
         TabIndex        =   0
         Top             =   870
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlDate uctlPlanningTo 
         Height          =   405
         Left            =   7680
         TabIndex        =   2
         Top             =   1290
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlDate uctlPlanningFrom 
         Height          =   405
         Left            =   1800
         TabIndex        =   1
         Top             =   1290
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtNetTotal 
         Height          =   435
         Left            =   7680
         TabIndex        =   18
         Top             =   870
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
      End
      Begin VB.Label lblNetTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPlanningTo"
         Height          =   315
         Left            =   5880
         TabIndex        =   19
         Top             =   960
         Width           =   1665
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   525
         Left            =   9840
         TabIndex        =   4
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblPlanningTo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPlanningTo"
         Height          =   315
         Left            =   5940
         TabIndex        =   17
         Top             =   1410
         Width           =   1665
      End
      Begin VB.Label lblPlanningFrom 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPlanningFrom"
         Height          =   315
         Left            =   60
         TabIndex        =   16
         Top             =   1410
         Width           =   1665
      End
      Begin VB.Label lblPlanningDate 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPlanningDate"
         Height          =   315
         Left            =   420
         TabIndex        =   15
         Top             =   990
         Width           =   1305
      End
      Begin VB.Label lblPlanningDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPlanningDesc"
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   1890
         Width           =   1605
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   11520
         TabIndex        =   10
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPlanning.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   13200
         TabIndex        =   11
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
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
         MouseIcon       =   "frmAddEditPlanning.frx":3250
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
         MouseIcon       =   "frmAddEditPlanning.frx":356A
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditPlanning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Planning As CPlanning

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public id As Long
Public PlanningArea As Long

Private m_ReportControls As Collection
Private m_Texts As Collection
Private m_Dates As Collection
Private m_Labels As Collection
Private m_Combos As Collection
Private m_TextLookups As Collection
Private m_Checks As Collection
Private m_CyclePerMonth As Long
Private m_PartGroups As Collection
Private Mr As CMasterRef
Private m_FromDate As Date
Private m_ToDate As Date
Private m_ToRcp As Date
Private m_PrintDate As Date
Private TempKey  As String

Public TempCollection As Collection
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
      
      m_Planning.PLANNING_ID = id
      m_Planning.QueryFlag = 1
      If Not glbPlanning.QueryPlanning(m_Planning, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
        Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_Planning.PopulateFromRS(1, m_Rs)

      txtPlanningDesc.Text = m_Planning.PLANNING_DESC
      uctlPlanningDate.ShowDate = m_Planning.PLANNING_DATE
      uctlPlanningFrom.ShowDate = m_Planning.PLANNING_FROM
      uctlPlanningTo.ShowDate = m_Planning.PLANNING_TO
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
Dim JO As CPlanningItem
Dim planVersion As Long
   
   If ShowMode = SHOW_EDIT Then
      If Not VerifyAccessRight("PLANNING_" & PlanningArea & "_EDIT", "แก้ไข" & PlanningArea2Text(PlanningArea)) Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   End If
   If PlanningArea = 1 Or PlanningArea = 3 Then
      If Not VerifyDate(lblPlanningDate, uctlPlanningDate, False) Then
         Exit Function
      End If
      
'      If Not CheckUniqueNs(PLANNING_UNIQUE, Trim(DateToStringInt(uctlPlanningDate.ShowDate)), ID, Trim(str(PlanningArea))) Then
'         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & uctlPlanningDate.ShowDate & " " & MapText("อยู่ในระบบแล้ว")
'         glbErrorLog.ShowUserError
'         Exit Function
'      End If
   End If
   If PlanningArea = 2 Or PlanningArea = 4 Then
      If Not VerifyDate(lblPlanningFrom, uctlPlanningFrom, False) Then
         Exit Function
      End If
      
      If Not VerifyDate(lblPlanningTo, uctlPlanningTo, False) Then
         Exit Function
      End If
      
'      If Not CheckUniqueNs(PLANNING_UNIQUE, Trim(DateToStringInt(uctlPlanningFrom.ShowDate)), ID, Trim(str(PlanningArea))) Then
'         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & uctlPlanningFrom.ShowDate & " " & MapText("อยู่ในระบบแล้ว")
'         glbErrorLog.ShowUserError
'         Exit Function
'      End If
   
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_Planning.PLANNING_ID = id
   m_Planning.AddEditMode = ShowMode
   If PlanningArea = 1 Or PlanningArea = 3 Then
     planVersion = LoadPlanVersion(uctlPlanningDate.ShowDate, uctlPlanningDate.ShowDate, Trim(str(PlanningArea)))
      uctlPlanningFrom.ShowDate = uctlPlanningDate.ShowDate
      uctlPlanningTo.ShowDate = uctlPlanningDate.ShowDate
   Else
      planVersion = LoadPlanVersion(uctlPlanningFrom.ShowDate, uctlPlanningTo.ShowDate, Trim(str(PlanningArea)))
      uctlPlanningDate.ShowDate = uctlPlanningFrom.ShowDate
   End If
   m_Planning.PLANNING_DATE = uctlPlanningDate.ShowDate
   m_Planning.PLANNING_FROM = uctlPlanningFrom.ShowDate
   m_Planning.PLANNING_TO = uctlPlanningTo.ShowDate
   
   m_Planning.PLANNING_DESC = txtPlanningDesc.Text
   m_Planning.PLANNING_AREA = PlanningArea
   m_Planning.PLAN_VERSION = planVersion + 1
   Call EnableForm(Me, False)
   
   
   Call glbDaily.StartTransaction
   
   If Not glbPlanning.AddEditPlanning(m_Planning, IsOK, False, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call glbDaily.RollbackTransaction
      Call EnableForm(Me, True)
      Exit Function
   End If
   Call glbDaily.CommitTransaction
   
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
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu
Dim id As Long
   If Not cmdAdd.Enabled Then
      Exit Sub
   End If

   OKClick = False
   Set frmAddEditPlanningItem.TempCollection = GetCollection(TabStrip1.SelectedItem.Tag)
   frmAddEditPlanningItem.ParentShowMode = ShowMode
   frmAddEditPlanningItem.ShowMode = SHOW_ADD
   Set frmAddEditPlanningItem.ParentForm = Me
   frmAddEditPlanningItem.ParentTag = TabStrip1.SelectedItem.Tag
   frmAddEditPlanningItem.HeaderText = MapText("เพิ่มรายการ")
   Load frmAddEditPlanningItem
   frmAddEditPlanningItem.Show 1
   
   OKClick = frmAddEditPlanningItem.OKClick

   Unload frmAddEditPlanningItem
   Set frmAddEditPlanningItem = Nothing
      
   If OKClick Then
      GridEX1.ItemCount = CountItem(GetCollection(TabStrip1.SelectedItem.Tag))
      GridEX1.Rebind
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
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
   
    If ID1 <= 0 Then
      GetCollection(TabStrip1.SelectedItem.Tag).Remove (ID2)
   Else
      GetCollection(TabStrip1.SelectedItem.Tag).Item(ID2).Flag = "D"
   End If

   GridEX1.ItemCount = CountItem(GetCollection(TabStrip1.SelectedItem.Tag))
   GridEX1.Rebind
   m_HasModify = True
   
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
   
    Set frmAddEditPlanningItem.TempCollection = GetCollection(TabStrip1.SelectedItem.Tag)
    frmAddEditPlanningItem.id = id
    frmAddEditPlanningItem.ShowMode = SHOW_EDIT
    Set frmAddEditPlanningItem.ParentForm = Me
    frmAddEditPlanningItem.ParentTag = TabStrip1.SelectedItem.Tag
    frmAddEditPlanningItem.HeaderText = MapText("แก้ไขรายการ")
    Load frmAddEditPlanningItem
    frmAddEditPlanningItem.Show 1

   OKClick = frmAddEditPlanningItem.OKClick

   Unload frmAddEditPlanningItem
   Set frmAddEditPlanningItem = Nothing

   If OKClick Then
      GridEX1.ItemCount = CountItem(GetCollection(TabStrip1.SelectedItem.Tag))
      GridEX1.Rebind
   End If
      
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdOK_Click()
Dim oMenu As cPopupMenu
Dim lMenuChosen  As Long
   If PlanningArea = 1 Or PlanningArea = 2 Or PlanningArea = 4 Then
      OKClick = True
      Unload Me
      Exit Sub
   End If
   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("บันทึก", "-", "บันทึกและออกจากหน้าจอ")
   If lMenuChosen = 0 Then
      Exit Sub
   End If

   If lMenuChosen = 1 Then
      If Not SaveData Then
         Exit Sub
      End If

      ShowMode = SHOW_EDIT
      id = m_Planning.PLANNING_ID
      m_Planning.QueryFlag = 1
      QueryData (True)
      m_HasModify = False
   ElseIf lMenuChosen = 3 Then
      If Not SaveData Then
         Exit Sub
      End If

      OKClick = True
      Unload Me
   End If


End Sub

Private Sub cmdPrint_Click()
Dim Report As CReportInterface
Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
Dim ReportKey As String
Dim ReportFlag As Boolean
Dim Rc As CReportConfig
Dim iCount As Long
Dim EditMode As SHOW_MODE_TYPE
Dim ReportMode As Long
Dim ClassName As String

   If m_HasModify Or ((Not m_HasModify) And (ShowMode = SHOW_ADD)) Then
      glbErrorLog.LocalErrorMsg = "กรุณาทำการบันทึกข้อมูลให้เรียบร้อยก่อน"
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   Set oMenu = New cPopupMenu
   If PlanningArea = 1 Then
         lMenuChosen = oMenu.Popup("ประมาณการการใช้วัตถุดิบรายวัน", "ปรับค่าหน้ากระดาษ")
   ElseIf PlanningArea = 2 Then
         lMenuChosen = oMenu.Popup("ประมาณการการใช้วัตถุดิบรายสัปดาห์", "ปรับค่าหน้ากระดาษ")
   ElseIf PlanningArea = 3 Then
         lMenuChosen = oMenu.Popup("ประมาณการรับเข้าวัตถุดิบรายวันจากซัพพลายเออร์", "ปรับค่าหน้ากระดาษ")
   ElseIf PlanningArea = 4 Then
         lMenuChosen = oMenu.Popup("ประมาณการการใช้วัตถุดิบรายเดือน", "ปรับค่าหน้ากระดาษ")
   End If
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   
   If lMenuChosen = 1 Then
      ReportKey = "CReportPlanning002"
      ClassName = "CReportPlanning002"
      Set Report = New CReportPlanning002
      ReportFlag = True
   ElseIf lMenuChosen = 2 Then
      ReportKey = "CReportPlanning002"
      ClassName = "CReportPlanning002"
      Set Rc = New CReportConfig
      Rc.REPORT_KEY = ReportKey
      Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
      Call Rc.QueryData(m_Rs, iCount)
      HeaderText = MapText("ใบเบิกสินค้า/วัตถุดิบ")
      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
   End If
   
   If Not Report Is Nothing Then
      Call Report.AddParam(m_Planning.PLANNING_ID, "PLANNING_ID")
      Call Report.AddParam(ReportKey, "REPORT_KEY")
   End If
   
   If ReportFlag Then
      Set frmReport.ReportObject = Report
      frmReport.HeaderText = pnlHeader.Caption
      frmReport.ClassName = ClassName
      Load frmReport
      frmReport.Show 1
   
      Unload frmReport
      Set frmReport = Nothing
      Set Report = Nothing
   Else
      frmReportConfig.ReportMode = 1
      frmReportConfig.ShowMode = EditMode
      frmReportConfig.id = Rc.REPORT_CONFIG_ID
      frmReportConfig.ReportKey = ReportKey
      frmReportConfig.HeaderText = HeaderText
      Load frmReportConfig
      frmReportConfig.Show 1
      
      Unload frmReportConfig
      Set frmReportConfig = Nothing
   End If
   
   Call EnableForm(Me, True)
End Sub
Private Sub FillReportInput(R As CReportInterface)
Dim C As CReportControl

'   Call R.AddParam(Picture1.Picture, "PICTURE")
   For Each C In m_ReportControls
      If (C.ControlType = "C") Then
         If C.Param1 <> "" Then
            Call R.AddParam(m_Combos(C.ControlIndex).Text, C.Param1)
         End If
         
         If C.Param2 <> "" Then
            Call R.AddParam(m_Combos(C.ControlIndex).ItemData(Minus2Zero(m_Combos(C.ControlIndex).ListIndex)), C.Param2)
         End If
      End If
   
      If (C.ControlType = "T") Then
         If C.Param1 <> "" Then
            Call R.AddParam(m_Texts(C.ControlIndex).Text, C.Param1)
         End If
         
         If C.Param2 <> "" Then
            Call R.AddParam(m_Texts(C.ControlIndex).Text, C.Param2)
         End If
      End If
   
      If (C.ControlType = "D") Then
         If C.Param1 <> "" Then
            Call R.AddParam(m_Dates(C.ControlIndex).ShowDate, C.Param1)
         End If
         
         If C.Param2 <> "" Then
            If m_Dates(C.ControlIndex).ShowDate <= 0 Then
               If C.Param2 = "TO_DOC_DATE" Then
                  m_Dates(C.ControlIndex).ShowDate = -1
               ElseIf C.Param2 = "FROM_DOC_DATE" Then
                  m_Dates(C.ControlIndex).ShowDate = -2
               ElseIf C.Param2 = "TO_PAY_DATE" Then
                  m_Dates(C.ControlIndex).ShowDate = -1
               ElseIf C.Param2 = "PRINT_DATE" Then
                  m_Dates(C.ControlIndex).ShowDate = -1
               End If
            End If
            If C.Param2 = "FROM_DOC_DATE" Or C.Param2 = "FROM_DATE" Then
               m_FromDate = m_Dates(C.ControlIndex).ShowDate
            ElseIf C.Param2 = "TO_DOC_DATE" Or C.Param2 = "TO_DATE" Then
               m_ToDate = m_Dates(C.ControlIndex).ShowDate
            ElseIf C.Param2 = "TO_PAY_DATE" Then
               m_ToRcp = m_Dates(C.ControlIndex).ShowDate
            ElseIf C.Param2 = "PRINT_DATE" Then
               m_PrintDate = m_Dates(C.ControlIndex).ShowDate
            End If
            Call R.AddParam(m_Dates(C.ControlIndex).ShowDate, C.Param2)
         End If
      End If
   
        If (C.ControlType = "CH") Then
         If C.Param1 <> "" Then
            Call R.AddParam(m_Checks(C.ControlIndex).Value, C.Param1)
         End If
         
         If C.Param2 <> "" Then
            Call R.AddParam(m_Checks(C.ControlIndex).Value, C.Param2)
         End If
      End If
    
   Next C
End Sub
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_Planning.QueryFlag = 1
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         m_Planning.QueryFlag = 0
         Call QueryData(False)
      End If
      
      'Call TabStrip1_Click
      Call EnableForm(Me, True)
      m_HasModify = False
      
      If PlanningArea = 1 Or PlanningArea = 3 Then
         uctlPlanningFrom.Enable = False
         uctlPlanningFrom.TabStop = False
         uctlPlanningTo.Enable = False
         uctlPlanningTo.TabStop = False
         uctlPlanningDate.SetFocus
      Else
         uctlPlanningDate.Enable = False
         uctlPlanningDate.TabStop = False
         uctlPlanningFrom.SetFocus
      End If
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

Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.HEIGHT = ScaleHeight
   pnlHeader.Width = ScaleWidth
   GridEX1.Width = ScaleWidth - 2 * GridEX1.Left
   GridEX1.HEIGHT = ScaleHeight - GridEX1.Top - 620
   TabStrip1.Width = GridEX1.Width
   cmdAdd.Top = ScaleHeight - 580
   cmdEdit.Top = ScaleHeight - 580
   cmdDelete.Top = ScaleHeight - 580
   cmdOK.Top = ScaleHeight - 580
   cmdExit.Top = ScaleHeight - 580
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
   cmdPrint.Top = ScaleHeight - 580
   cmdPrint.Left = cmdOK.Left - cmdPrint.Width - 50
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_Planning = Nothing
End Sub
Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   '''Debug.Print ColIndex & " " & NewColWidth
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
   Col.Caption = MapText("รหัส")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 5000
   Col.Caption = MapText("ชื่อ")

   Set Col = GridEX1.Columns.add '6
   Col.Width = 2000
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("จำนวน")

   Set Col = GridEX1.Columns.add '7
   Col.Width = 2500
   Col.Caption = MapText("หน่วย")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = 4000
   Col.Caption = MapText("หมายเหตุ")
End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitGrid1
   
   Call InitNormalLabel(lblPlanningDate, MapText("วันที่ประมาณ"))
   Call InitNormalLabel(lblPlanningFrom, MapText("จากวันที่ประมาณ"))
   Call InitNormalLabel(lblPlanningTo, MapText("ถึงวันที่ประมาณ"))
   
   Call InitNormalLabel(lblPlanningDesc, MapText("รายละเอียด"))
   Call InitNormalLabel(lblNetTotal, MapText("ยอดรวม"))
   
   txtNetTotal.Enabled = False
   Call txtNetTotal.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPrint.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdPrint, MapText("พิมพ์"))
   
   If PlanningArea = 1 Or PlanningArea = 2 Or PlanningArea = 4 Then
      cmdAdd.Enabled = False
      cmdEdit.Enabled = False
      cmdDelete.Enabled = False
  End If
  
   TabStrip1.Font.Bold = True
   TabStrip1.Font.NAME = GLB_FONT
   TabStrip1.Font.Size = 16
   
   Dim T As Object
   TabStrip1.Tabs.Clear
   
   If PlanningArea = 1 Or PlanningArea = 2 Or PlanningArea = 4 Then
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText("วัตถดุดิบที่ใช้ผลิต")
      T.Tag = "PART_USE"
      
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText("สินค้าที่ผลิตได้")
      T.Tag = "PRODUCT_GET"
   Else
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText("วัตถุดิบที่ใช้ผลิตรับจากซัพพลายเออร์")
      T.Tag = "PART_SUP"
   End If
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
   Set m_Planning = New CPlanning
   
   
   Set m_ReportControls = New Collection
   Set m_Texts = New Collection
   Set m_Dates = New Collection
   Set m_Labels = New Collection
   Set m_Combos = New Collection
   Set m_TextLookups = New Collection
   Set m_Checks = New Collection
   Set m_PartGroups = New Collection
End Sub
Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub
Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long
   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"

   If GetCollection(TabStrip1.SelectedItem.Tag) Is Nothing Then
       Exit Sub
    End If

   If RowIndex <= 0 Then
      Exit Sub
   End If

   Dim Pni As CPlanningItem
   If GetCollection(TabStrip1.SelectedItem.Tag).Count <= 0 Then
      Exit Sub
   End If
   Set Pni = GetItem(GetCollection(TabStrip1.SelectedItem.Tag), RowIndex, RealIndex)
   If Pni Is Nothing Then
      Exit Sub
   End If
   
   Values(1) = Pni.PLANNING_ITEM_ID
   Values(2) = RealIndex
   Values(3) = Pni.PART_NO
   Values(4) = Pni.PART_DESC
   Values(5) = FormatNumber(Pni.PLAN_AMOUNT, 2)
   Values(6) = Pni.UNIT_NAME
   Values(7) = Pni.NOTE
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub TabStrip1_Click()
   Call InitGrid1
   GridEX1.ItemCount = CountItem(GetCollection(TabStrip1.SelectedItem.Tag))
   GridEX1.Rebind
   
   Call GetTotalAmount
End Sub
Private Sub txtPlanningDesc_Change()
   m_HasModify = True
End Sub
Private Sub uctlPlanningDate_HasChange()
   m_HasModify = True
End Sub
Private Sub uctlPlanningFrom_HasChange()
   m_HasModify = True
End Sub
Private Sub uctlPlanningTo_HasChange()
   m_HasModify = True
End Sub
Public Sub RefreshGrid(Optional Tag As String = "")
   If Len(Tag) > 0 Then
      GridEX1.ItemCount = CountItem(GetCollection(Tag))
      GridEX1.Rebind
      m_HasModify = True
   Else
      GridEX1.ItemCount = CountItem(GetCollection(TabStrip1.SelectedItem.Tag))
      GridEX1.Rebind
   End If
End Sub
Private Function GetCollection(Tag As String) As Collection
   If Tag = "PART_USE" Then
      Set GetCollection = m_Planning.CollPartUse
   ElseIf Tag = "PRODUCT_GET" Then
      Set GetCollection = m_Planning.CollProductGet
   ElseIf Tag = "PART_SUP" Then
      Set GetCollection = m_Planning.CollPartSup
   End If
End Function
Private Sub GetTotalAmount()
Dim II As CPlanningItem
Dim SumAmount As Double
   
   SumAmount = 0

   For Each II In GetCollection(TabStrip1.SelectedItem.Tag)
      If II.Flag <> "D" Then
         SumAmount = SumAmount + II.PLAN_AMOUNT
      End If
   Next II

   txtNetTotal.Text = FormatNumber(SumAmount, "0.00")
End Sub
