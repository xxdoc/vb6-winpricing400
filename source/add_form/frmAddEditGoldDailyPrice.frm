VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditGoldDailyPrice 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmAddEditGoldDailyPrice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboUnit 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2340
         Width           =   2955
      End
      Begin VB.ComboBox cboPartType 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1920
         Width           =   2955
      End
      Begin prjFarmManagement.uctlTextBox txtName 
         Height          =   435
         Left            =   1860
         TabIndex        =   1
         Top             =   1470
         Width           =   4485
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   9
         Top             =   0
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtPartNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   1020
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtUnitWeight 
         Height          =   435
         Left            =   7800
         TabIndex        =   2
         Top             =   1470
         Width           =   1305
         _ExtentX        =   5212
         _ExtentY        =   767
      End
      Begin VB.Label lblKg 
         Caption         =   "Label1"
         Height          =   345
         Left            =   9180
         TabIndex        =   15
         Top             =   1560
         Width           =   1365
      End
      Begin VB.Label lblUnitWeight 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   6360
         TabIndex        =   14
         Top             =   1560
         Width           =   1365
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   13
         Top             =   2400
         Width           =   1575
      End
      Begin Threed.SSCheck chkPigFlag 
         Height          =   345
         Left            =   6420
         TabIndex        =   7
         Top             =   1050
         Visible         =   0   'False
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   609
         _Version        =   131073
         Caption         =   "SSCheck1"
         TripleState     =   -1  'True
      End
      Begin VB.Label lblPartType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   12
         Top             =   1980
         Width           =   1575
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   11
         Top             =   1530
         Width           =   1575
      End
      Begin VB.Label lblPartNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   10
         Top             =   1110
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5985
         TabIndex        =   6
         Top             =   3270
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   4335
         TabIndex        =   5
         Top             =   3270
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditGoldDailyPrice.frx":27A2
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditGoldDailyPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_PartItem As CPartItem

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String

Private Sub cmdPasswd_Click()

End Sub

Private Sub cboPartType_Click()
   m_HasModify = True
End Sub

Private Sub cboUnit_Click()
   m_HasModify = True
End Sub

Private Sub chkPigFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
'   Call LoadAccessRight(Nothing, glbAccessRight, glbUser.GROUP_ID)
   OKClick = True
   Unload Me
End Sub

Private Sub InitGrid1()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.Name = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.Add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   Set Col = GridEX1.Columns.Add '3
   Col.Width = 3465
   Col.Caption = MapText("สถานที่จัดเก็บ")

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 2745
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("จำนวนคงคลัง")

   Set Col = GridEX1.Columns.Add '5
   Col.Width = 2790
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ราคาเฉลี่ย")

   Set Col = GridEX1.Columns.Add '6
   Col.Width = 2565
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ราคาหลังสุด")
   
   Set Col = GridEX1.Columns.Add '7
   Col.Width = 0
   Col.Visible = False
   Col.Caption = MapText("LocationID")
End Sub

Private Sub InitGrid2()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.Name = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.Add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   Set Col = GridEX1.Columns.Add '3
   Col.Width = 6030
   Col.Caption = MapText("ชื่อซัพพลายเออร์")

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 2745
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ยอดนำเข้ารวม")

   Set Col = GridEX1.Columns.Add '5
   Col.Width = 2790
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("มูลค่านำเข้ารวม")
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      m_PartItem.PART_ITEM_ID = ID
      m_PartItem.QueryFlag = 1
      If Not glbDaily.QueryPartItem(m_PartItem, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_PartItem.PopulateFromRS(1, m_Rs)
      
      txtName.Text = m_PartItem.PART_DESC
      txtPartNo.Text = m_PartItem.PART_NO
      cboPartType.ListIndex = IDToListIndex(cboPartType, m_PartItem.PART_TYPE)
      cboUnit.ListIndex = IDToListIndex(cboUnit, m_PartItem.UNIT_COUNT)
      chkPigFlag.Value = FlagToCheck(m_PartItem.PIG_FLAG)
      txtUnitWeight.Text = m_PartItem.UNIT_WEIGHT
      
      TabStrip1_Click
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_DblClick()
Dim LocationID As Long

   If TabStrip1.SelectedItem.Index = 1 Then
      If Not VerifyGrid(GridEX1.Value(1)) Then
         Exit Sub
      End If

      LocationID = Val(GridEX1.Value(7))
      
      frmLotItem.PartItemID = m_PartItem.PART_ITEM_ID
      frmLotItem.LocationID = LocationID
      Load frmLotItem
      frmLotItem.Show 1
      
      Unload frmLotItem
      Set frmLotItem = Nothing
   End If
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Index = 1 Then
      If m_PartItem.PartLocations Is Nothing Then
         Exit Sub
      End If
   
      If RowIndex <= 0 Then
         Exit Sub
      End If
   
      Dim CR As CPartLocation
      If m_PartItem.PartLocations.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_PartItem.PartLocations, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If
   
      Values(1) = CR.PART_LOCATION_ID
      Values(2) = RealIndex
      Values(3) = CR.LOCATION_NAME
      Values(4) = FormatNumber(CR.CURRENT_AMOUNT)
      Values(5) = FormatNumber(CR.AVG_PRICE)
      Values(6) = FormatNumber(CR.LAST_PRICE)
      Values(7) = FormatNumber(CR.LOCATION_ID)
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      If m_PartItem.Suppliers Is Nothing Then
         Exit Sub
      End If
   
      If RowIndex <= 0 Then
         Exit Sub
      End If
   
      Dim Sp As CSupplier
      If m_PartItem.Suppliers.Count <= 0 Then
         Exit Sub
      End If
      Set Sp = GetItem(m_PartItem.Suppliers, RowIndex, RealIndex)
      If Sp Is Nothing Then
         Exit Sub
      End If
   
      Values(1) = Sp.SUPPLIER_ID
      Values(2) = RealIndex
      Values(3) = Sp.SUPPLIER_NAME
      Values(4) = FormatNumber(Sp.TX_AMOUNT)
      Values(5) = FormatNumber(Sp.TOTAL_INCLUDE_PRICE)
   End If
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean

   If ShowMode = SHOW_ADD Then
      If Not VerifyAccessRight("INVENTORY_PART_ADD") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   ElseIf ShowMode = SHOW_EDIT Then
      If Not VerifyAccessRight("INVENTORY_PART_EDIT") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   End If
   
   If Not VerifyTextControl(lblPartNo, txtPartNo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblName, txtName, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblUnitWeight, txtUnitWeight, True) Then
      Exit Function
   End If
   If Not VerifyCombo(lblPartType, cboPartType, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblUnit, cboUnit, False) Then
      Exit Function
   End If
   
   If Not CheckUniqueNs(PARTNO_UNIQUE, txtName.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtName.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_PartItem.PART_ITEM_ID = ID
   m_PartItem.AddEditMode = ShowMode
   m_PartItem.PIG_FLAG = Check2Flag(chkPigFlag.Value)
   m_PartItem.PART_NO = txtPartNo.Text
   m_PartItem.PART_DESC = txtName.Text
   m_PartItem.PART_TYPE = cboPartType.ItemData(Minus2Zero(cboPartType.ListIndex))
   m_PartItem.UNIT_COUNT = cboUnit.ItemData(Minus2Zero(cboUnit.ListIndex))
   m_PartItem.UNIT_WEIGHT = Val(txtUnitWeight.Text)
   
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditPartItem(m_PartItem, IsOK, True, glbErrorLog) Then
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

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadPartType(cboPartType)
      Call LoadUnit(cboUnit)
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
      End If
      
      m_HasModify = False
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 116 Then
'      Call cmdSearch_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 118 Then
'      Call cmdAdd_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
'      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
'      Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
      KeyCode = 0
   End If
End Sub

Private Sub InitFormLayout()
   Call InitGrid1
   
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblName, MapText("ชื่อวัตถุดิบ"))
   Call InitNormalLabel(lblPartNo, MapText("รหัสวัตถุดิบ"))
   Call InitNormalLabel(lblPartType, MapText("ประเภทวัตถุดิบ"))
   Call InitNormalLabel(lblUnit, MapText("หน่วยวัด"))
   Call InitNormalLabel(lblUnitWeight, MapText("ความหนาแน่น"))
   Call InitNormalLabel(lblKg, MapText(""))
   
   Call InitCheckBox(chkPigFlag, "PIG FLAG")
   
   Call txtName.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   Call txtPartNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtUnitWeight.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitCombo(cboPartType)
   Call InitCombo(cboUnit)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.Name = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.Add().Caption = MapText("สถานที่จัดเก็บ")
   TabStrip1.Tabs.Add().Caption = MapText("ซัพพลายเออร์")
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   Set m_PartItem = New CPartItem
   Set m_Rs = New ADODB.Recordset

   Call EnableForm(Me, False)
   m_HasActivate = False
      
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub txtLastName_Change()
   m_HasModify = True
End Sub

Private Sub TabStrip1_Click()
   If TabStrip1.SelectedItem.Index = 1 Then
      Call InitGrid1
      
      GridEX1.ItemCount = CountItem(m_PartItem.PartLocations)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      Call InitGrid2
      
      GridEX1.ItemCount = CountItem(m_PartItem.Suppliers)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub txtPartNo_Change()
   m_HasModify = True
End Sub

Private Sub txtName_Change()
   m_HasModify = True
End Sub

Private Sub txtMaxUser_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox1_Change()
   m_HasModify = True
End Sub

Private Sub txtUnitWeight_Change()
   m_HasModify = True
End Sub
