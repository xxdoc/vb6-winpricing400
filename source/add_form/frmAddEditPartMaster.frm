VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditPartMaster 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8970
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditPartMaster.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   8970
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   6285
      Left            =   0
      TabIndex        =   9
      Top             =   600
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   11086
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextBox txtPartMasterName 
         Height          =   435
         Left            =   1785
         TabIndex        =   1
         Top             =   780
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPartMasterNo 
         Height          =   435
         Left            =   1800
         TabIndex        =   0
         Top             =   300
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   767
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   2805
         Left            =   120
         TabIndex        =   5
         Top             =   2340
         Width           =   8715
         _ExtentX        =   15372
         _ExtentY        =   4948
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
         HeaderFontBold  =   -1  'True
         HeaderFontWeight=   700
         FontSize        =   9.75
         BackColorBkg    =   16777215
         ColumnHeaderHeight=   480
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmAddEditPartMaster.frx":08CA
         Column(2)       =   "frmAddEditPartMaster.frx":0992
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditPartMaster.frx":0A36
         FormatStyle(2)  =   "frmAddEditPartMaster.frx":0B92
         FormatStyle(3)  =   "frmAddEditPartMaster.frx":0C42
         FormatStyle(4)  =   "frmAddEditPartMaster.frx":0CF6
         FormatStyle(5)  =   "frmAddEditPartMaster.frx":0DCE
         ImageCount      =   0
         PrinterProperties=   "frmAddEditPartMaster.frx":0E86
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   8715
         _ExtentX        =   15372
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
      Begin prjFarmManagement.uctlTextLookup uctlAnimalType 
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   1320
         Width           =   4455
         _ExtentX        =   9975
         _ExtentY        =   661
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3390
         TabIndex        =   15
         Top             =   5520
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPartMaster.frx":105E
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   120
         TabIndex        =   14
         Top             =   5520
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPartMaster.frx":1378
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1740
         TabIndex        =   13
         Top             =   5520
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblAnimalType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   1485
      End
      Begin Threed.SSCheck chkCancelFlag 
         Height          =   375
         Left            =   5760
         TabIndex        =   3
         Top             =   300
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "chkCancelFlag"
      End
      Begin VB.Label lblPartMasterNo 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   11
         Top             =   360
         Width           =   1485
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   5520
         TabIndex        =   6
         Top             =   5520
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPartMaster.frx":1692
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   7200
         TabIndex        =   7
         Top             =   5520
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblPartMasterName 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   10
         Top             =   840
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmAddEditPartMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Header As String
Public ShowMode As SHOW_MODE_TYPE
Public ParentShowMode As SHOW_MODE_TYPE
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_PartMaster As CPartMaster

Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public TempCollection As Collection
Public COMMIT_FLAG As String
Public PartItemID As Long
Public PartType As Long
Public PartMasterType As Long
Private m_AnimalType As Collection
Private m_PartGroupMenus As Collection

'Private m_PartTypes As Collection
'Private m_Parts As Collection

Private Sub cboTextType_Click()
   m_HasModify = True
End Sub

Private Sub cboCountry_Click()
   m_HasModify = True
End Sub

Private Sub chkBangkok_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkCancelFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu


   If Not cmdAdd.Enabled Then
      Exit Sub
   End If
   
   OKClick = False
   If TabStrip1.SelectedItem.Index = 1 Then
      If Not VerifyAccessRight("INVENTORY_PART") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.AddMenu(m_PartGroupMenus)
      If lMenuChosen = 0 Then
         Exit Sub
      End If

      frmPartItem.PartGroupID = lMenuChosen
      Load frmPartItem
      frmPartItem.Show 1
      
      Unload frmPartItem
      Set frmPartItem = Nothing
      
      Call QueryData(True)
   ElseIf TabStrip1.SelectedItem.Index = 2 Then

      Set frmAddEditCusGroup.TempCollection = m_PartMaster.CusGroups 'SetKeyTempCollection(m_PartMaster.CusGroups, "CPartItem") '
      Set frmAddEditCusGroup.ParentForm = Me
      frmAddEditCusGroup.ParentTag = TabStrip1.SelectedItem.Tag
      frmAddEditCusGroup.ShowMode = SHOW_ADD
      frmAddEditCusGroup.HeaderText = MapText("เพิ่มรายการกลุ่มลูกค้า")
      Load frmAddEditCusGroup
      frmAddEditCusGroup.Show 1

      OKClick = frmAddEditCusGroup.OKClick

      Unload frmAddEditCusGroup
      Set frmAddEditCusGroup = Nothing
      
      If OKClick Then
'         Call GetTotalPrice

         GridEX1.ItemCount = CountItem(m_PartMaster.CusGroups)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub
Public Function SetKeyTempCollection(Cl As Collection, Cs As String) As Collection
 Dim T As Object
 Dim tColl As Collection
If Cs = "CPartItem" Then
  Set T = New CPartItem
ElseIf Cs = "" Then

End If

Set tColl = New Collection
For Each T In Cl 'Copy
   Call tColl.add(T)
Next T

For Each T In tColl 'Remove
   Call Cl.Remove(1)
Next T

For Each T In tColl 'Add
  If Cs = "CPartItem" Then
     Call Cl.add(T, str(T.PART_CUS_GROUPS_ID))
  End If
Next T

Set T = Nothing

Set SetKeyTempCollection = Cl
End Function
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
   
   If TabStrip1.SelectedItem.Index = 1 Then

   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      If ID1 <= 0 Then
         m_PartMaster.CusGroups.Remove (ID2)
      Else
         m_PartMaster.CusGroups.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_PartMaster.CusGroups)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(2))
   OKClick = False

   If TabStrip1.SelectedItem.Index = 1 Then
      If Not VerifyAccessRight("INVENTORY_PART") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
'      Set oMenu = New cPopupMenu
'      lMenuChosen = oMenu.AddMenu(m_PartGroupMenus)
'      If lMenuChosen = 0 Then
'         Exit Sub
'      End If

'      frmPartItem.PartGroupID = lMenuChosen
      frmPartItem.txtPartNo.Text = Trim(GridEX1.Value(3))
      Load frmPartItem
      frmPartItem.Show 1
      
      Unload frmPartItem
      Set frmPartItem = Nothing
      
      Call QueryData(True)
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      frmAddEditCusGroup.ID = ID
'      frmAddEditCusGroup.COMMIT_FLAG = m_Payment.COMMIT_FLAG
      Set frmAddEditCusGroup.ParentForm = Me
      Set frmAddEditCusGroup.TempCollection = m_PartMaster.CusGroups
      frmAddEditCusGroup.ParentTag = TabStrip1.SelectedItem.Tag
      frmAddEditCusGroup.HeaderText = MapText("แก้ไขรายการนำฝากธนาคาร")
      frmAddEditCusGroup.ParentShowMode = ShowMode
      frmAddEditCusGroup.ShowMode = SHOW_EDIT
      Load frmAddEditCusGroup
      frmAddEditCusGroup.Show 1

      OKClick = frmAddEditCusGroup.OKClick

      Unload frmAddEditCusGroup
      Set frmAddEditCusGroup = Nothing
      
      If OKClick Then
'         Call GetTotalPrice
         GridEX1.ItemCount = CountItem(m_PartMaster.CusGroups)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub InitFormLayout()
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)

   Me.KeyPreview = True
   pnlHeader.Caption = HeaderText
   Me.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_HEAD_COLOR
   SSFrame1.BackColor = GLB_FORM_COLOR
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText

   Call InitNormalLabel(lblPartMasterNo, MapText("รหัสสินค้า"))
   Call InitNormalLabel(lblPartMasterName, MapText("ชื่อสินค้า"))
   Call InitNormalLabel(lblAnimalType, MapText("ชนิดสัตว์"))
   Call InitCheckBox(chkCancelFlag, "ยกเลิก")

   Call txtPartMasterNo.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
    Call txtPartMasterName.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)

   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.NAME = GLB_FONT
   TabStrip1.Font.Size = 16
   
   Dim T As Object
   TabStrip1.Tabs.Clear
   
   Set T = TabStrip1.Tabs.add()
   T.Caption = MapText("เบอร์อาหาร")
   T.Tag = "PART_NO"
   
   Set T = TabStrip1.Tabs.add()
   T.Caption = MapText("กลุ่มลูกค้า")
   T.Tag = "CUS_GROUP"
   
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
      m_PartMaster.PART_MASTER_ID = ID
      m_PartMaster.QueryFlag = 1
      If Not glbDaily.QueryPartMaster(m_PartMaster, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_PartMaster.PopulateFromRS(1, m_Rs)
      
      txtPartMasterNo.Text = m_PartMaster.PART_MASTER_NO
      txtPartMasterName.Text = m_PartMaster.PART_MASTER_NAME
      chkCancelFlag.Value = FlagToCheck(m_PartMaster.CANCEL_FLAG)
      uctlAnimalType.MyCombo.ListIndex = IDToListIndex(uctlAnimalType.MyCombo, m_PartMaster.ANIMAL_TYPE)
   End If
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call TabStrip1_Click
   Call EnableForm(Me, True)
End Sub


Private Sub cmdOK_Click()
   If Not cmdOK.Enabled Then
      Exit Sub
   End If
   
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long

   If Not VerifyTextControl(lblPartMasterNo, txtPartMasterNo, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblPartMasterName, txtPartMasterName, False) Then
      Exit Function
   End If
   
    If Not CheckUniqueNs(PART_MASTER_NO_UNIQUE, txtPartMasterNo.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtPartMasterNo.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If

   
   m_PartMaster.PART_MASTER_ID = ID
   m_PartMaster.AddEditMode = ShowMode
   m_PartMaster.PART_MASTER_NO = txtPartMasterNo.Text
   m_PartMaster.PART_MASTER_NAME = txtPartMasterName.Text
   m_PartMaster.CANCEL_FLAG = Check2Flag(chkCancelFlag.Value)
   m_PartMaster.PART_MASTER_TYPE = PartMasterType
   m_PartMaster.ANIMAL_TYPE = uctlAnimalType.MyCombo.ItemData(Minus2Zero(uctlAnimalType.MyCombo.ListIndex))
   
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditPartMaster(m_PartMaster, IsOK, True, glbErrorLog) Then
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
   
   
'   Dim PartMaster As CPartMaster
'   If ShowMode = SHOW_ADD Then
'      Set PartMaster = New CPartMaster
'      PartMaster.Flag = "A"
'      Call TempCollection.add(PartMaster)
'   Else
'      Set PartMaster = TempCollection.Item(ID)
'      If PartMaster.Flag <> "A" Then
'         PartMaster.Flag = "E"
'      End If
'   End If
         

'
'   Set PartMaster = Nothing
'   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
     Call LoadMaster(uctlAnimalType.MyCombo, m_AnimalType, ANIMAL_TYPE)
     Set uctlAnimalType.MyCollection = m_AnimalType
     
     Call GeneratePartGroupMenu(m_PartGroupMenus)
      
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
   ElseIf Shift = 0 And KeyCode = 123 Then
      Call AddMemoNote
      KeyCode = 0
   End If
End Sub

Private Sub Form_Load()

   OKClick = False
   Call InitFormLayout
   
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_PartMaster = New CPartMaster
   Set m_AnimalType = New Collection
   Set m_PartGroupMenus = New Collection
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_PartMaster = Nothing
   Set m_AnimalType = Nothing
   Set m_PartGroupMenus = Nothing
End Sub

Private Sub GridEX1_DblClick()
  Call cmdEdit_Click
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Index = 1 Then
      If m_PartMaster.PartItems Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim Pi As CPartItem
      If m_PartMaster.PartItems.Count <= 0 Then
         Exit Sub
      End If
      Set Pi = GetItem(m_PartMaster.PartItems, RowIndex, RealIndex)
      If Pi Is Nothing Then
         Exit Sub
      End If

      Values(1) = Pi.PART_ITEM_ID
      Values(2) = RealIndex
      Values(3) = Pi.PART_NO
      Values(4) = Pi.PART_DESC
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      If m_PartMaster.CusGroups Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim CR As CCusGroup
      If m_PartMaster.CusGroups.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_PartMaster.CusGroups, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If

      Values(1) = CR.PART_CUS_GROUPS_ID
      Values(2) = RealIndex
      Values(3) = CR.CUS_GROUPS_NO
      Values(4) = CR.CUS_GROUPS_NAME
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
      
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub TabStrip1_Click()
   If TabStrip1.SelectedItem.Index = 1 Then
      Call InitGrid1
      
      GridEX1.ItemCount = CountItem(m_PartMaster.PartItems)
      GridEX1.Rebind
      GridEX1.Visible = True
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      Call InitGrid2
      
      GridEX1.ItemCount = CountItem(m_PartMaster.CusGroups)
      GridEX1.Rebind
      GridEX1.Visible = True
      
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
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
   Col.Width = 3500
   Col.Caption = MapText("รหัสสินค้า")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 5000
   Col.Caption = MapText("ชื่อสินค้า")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 0
   Col.Caption = MapText("PART_ITEM_ID")

End Sub
Private Sub InitGrid2()
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

   Set Col = GridEX1.Columns.add '4
   Col.Width = 2415
   Col.Caption = MapText("รหัสประเภทลูกค้า")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 4000
   Col.Caption = MapText("ชื่อประเภทลูกค้า")

End Sub
Private Sub txtPartMasterName_Change()
   m_HasModify = True
End Sub

Private Sub txtPartMasterNo_Change()
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
   If Tag = "PART_NO" Then
      Set GetCollection = m_PartMaster.PartItems
   ElseIf Tag = "CUS_GROUP" Then
      Set GetCollection = m_PartMaster.CusGroups
   End If
End Function

Private Sub uctlAnimalType_Change()
   m_HasModify = True
End Sub
