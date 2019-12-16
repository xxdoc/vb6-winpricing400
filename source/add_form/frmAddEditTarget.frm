VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditTarget 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11865
   Icon            =   "frmAddEditTarget.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11865
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8895
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   12720
      _ExtentX        =   22437
      _ExtentY        =   15690
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   120
         TabIndex        =   2
         Top             =   2220
         Width           =   11640
         _ExtentX        =   20532
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
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   4935
         Left            =   120
         TabIndex        =   3
         Top             =   2760
         Width           =   11640
         _ExtentX        =   20532
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
         Column(1)       =   "frmAddEditTarget.frx":27A2
         Column(2)       =   "frmAddEditTarget.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditTarget.frx":290E
         FormatStyle(2)  =   "frmAddEditTarget.frx":2A6A
         FormatStyle(3)  =   "frmAddEditTarget.frx":2B1A
         FormatStyle(4)  =   "frmAddEditTarget.frx":2BCE
         FormatStyle(5)  =   "frmAddEditTarget.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditTarget.frx":2D5E
      End
      Begin prjFarmManagement.uctlTextBox txtTargetDesc 
         Height          =   435
         Left            =   1800
         TabIndex        =   1
         Top             =   1410
         Width           =   9705
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtYearNo 
         Height          =   435
         Left            =   1800
         TabIndex        =   0
         Top             =   990
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
      End
      Begin VB.Label lblYearNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblYearNo"
         Height          =   315
         Left            =   0
         TabIndex        =   12
         Top             =   1080
         Width           =   1665
      End
      Begin VB.Label lblTargetDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "lblTargetDesc"
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   1530
         Width           =   1605
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8490
         TabIndex        =   7
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditTarget.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10170
         TabIndex        =   8
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
         TabIndex        =   5
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
         TabIndex        =   4
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditTarget.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   6
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditTarget.frx":356A
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditTarget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public id As Long

Private m_Target As CTarget

Public TempCollection As Collection
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
      
      m_Target.TARGET_ID = id
      m_Target.QueryFlag = 1
      If Not glbTargetCom.QueryTarget(m_Target, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
        Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_Target.PopulateFromRS(1, m_Rs)

      txtYearNo.Text = m_Target.YEAR_NO + 543
      txtTargetDesc.Text = m_Target.TARGET_DESC
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
      
   If ShowMode = SHOW_EDIT Then
      If Not VerifyAccessRight("COMMISSION_TARGET_EDIT") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   End If
   
   If Not VerifyTextControl(lblYearNo, txtYearNo, False) Then
      Exit Function
   End If
      
   If Not CheckUniqueNs(TARGET_UNIQUE, txtYearNo.Text, id) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtYearNo.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_Target.TARGET_ID = id
   m_Target.AddEditMode = ShowMode
   
   m_Target.YEAR_NO = Val(txtYearNo.Text) - 543
   m_Target.TARGET_DESC = txtTargetDesc.Text
   
   Call EnableForm(Me, False)
   
   
   Call glbDaily.StartTransaction
   
   If Not glbTargetCom.AddEditTarget(m_Target, IsOK, False, glbErrorLog) Then
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
   Set frmAddEditTargetItem.TempCollection = GetCollection(TabStrip1.SelectedItem.Tag)
   frmAddEditTargetItem.ParentShowMode = ShowMode
   frmAddEditTargetItem.ShowMode = SHOW_ADD
   Set frmAddEditTargetItem.ParentForm = Me
   frmAddEditTargetItem.ParentTag = TabStrip1.SelectedItem.Tag
   frmAddEditTargetItem.HeaderText = MapText("เพิ่มรายการ")
   Load frmAddEditTargetItem
   frmAddEditTargetItem.Show 1
   
   OKClick = frmAddEditTargetItem.OKClick

   Unload frmAddEditTargetItem
   Set frmAddEditTargetItem = Nothing
      
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
   
    Set frmAddEditTargetItem.TempCollection = GetCollection(TabStrip1.SelectedItem.Tag)
    frmAddEditTargetItem.id = id
    frmAddEditTargetItem.ShowMode = SHOW_EDIT
    Set frmAddEditTargetItem.ParentForm = Me
    frmAddEditTargetItem.ParentTag = TabStrip1.SelectedItem.Tag
    frmAddEditTargetItem.HeaderText = MapText("แก้ไขรายการ")
    Load frmAddEditTargetItem
    frmAddEditTargetItem.Show 1

   OKClick = frmAddEditTargetItem.OKClick

   Unload frmAddEditTargetItem
   Set frmAddEditTargetItem = Nothing

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
      id = m_Target.TARGET_ID
      m_Target.QueryFlag = 1
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
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_Target.QueryFlag = 1
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         m_Target.QueryFlag = 0
         Call QueryData(False)
      End If
      
      Call EnableForm(Me, True)
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_Target = Nothing
End Sub
Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   '''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid1()
Dim Col As JSColumn
Dim I As Long
   
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
   Col.Width = 1500
   Col.Caption = MapText("รหัส SALE")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 3000
   Col.Caption = MapText("ชื่อ SALE")
   
   For I = 1 To 12
      Set Col = GridEX1.Columns.add '6
      Col.Width = 1500
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("เป้าเดือน " & I)
   Next I
End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitGrid1
   
   Call InitNormalLabel(lblYearNo, MapText("ปี"))
   Call InitNormalLabel(lblTargetDesc, MapText("รายละเอียด"))
   
   Call txtYearNo.SetTextLenType(TEXT_INTEGER, glbSetting.AMOUNT_LEN)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.NAME = GLB_FONT
   TabStrip1.Font.Size = 16
   
   Dim T As Object
   TabStrip1.Tabs.Clear
   
   Set T = TabStrip1.Tabs.add()
   T.Caption = MapText("ประมาณการรายเดือน")
   T.Tag = "TARGET_DETAIL"

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
   Set m_Target = New CTarget
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

   Dim Tgdt As CTargetDetail
   If GetCollection(TabStrip1.SelectedItem.Tag).Count <= 0 Then
      Exit Sub
   End If
   Set Tgdt = GetItem(GetCollection(TabStrip1.SelectedItem.Tag), RowIndex, RealIndex)
   If Tgdt Is Nothing Then
      Exit Sub
   End If
   
   Values(1) = Tgdt.TARGET_DETAIL_ID
   Values(2) = RealIndex
   Values(3) = Tgdt.EMP_CODE
   Values(4) = Tgdt.EMP_NAME
   
   Values(5) = Tgdt.TARGET_PRICE1
   Values(6) = Tgdt.TARGET_PRICE2
   Values(7) = Tgdt.TARGET_PRICE3
   Values(8) = Tgdt.TARGET_PRICE4
   Values(9) = Tgdt.TARGET_PRICE5
   Values(10) = Tgdt.TARGET_PRICE6
   Values(11) = Tgdt.TARGET_PRICE7
   Values(12) = Tgdt.TARGET_PRICE8
   Values(13) = Tgdt.TARGET_PRICE9
   Values(14) = Tgdt.TARGET_PRICE10
   Values(15) = Tgdt.TARGET_PRICE11
   Values(16) = Tgdt.TARGET_PRICE12
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub TabStrip1_Click()
   Call InitGrid1
   GridEX1.ItemCount = CountItem(GetCollection(TabStrip1.SelectedItem.Tag))
   GridEX1.Rebind
End Sub
Private Sub txtTargetDesc_Change()
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
   If Tag = "TARGET_DETAIL" Then
      Set GetCollection = m_Target.CollTargerDetail
   End If
End Function

Private Sub txtYearNo_Change()
   m_HasModify = True
End Sub
