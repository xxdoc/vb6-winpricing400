VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmQueryPartItem 
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10005
   Icon            =   "frmQueryPartItem.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   10005
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   6285
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   11086
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboPartType 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1950
         Width           =   3585
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   2865
         Left            =   90
         TabIndex        =   12
         Top             =   2610
         Width           =   9825
         _ExtentX        =   17330
         _ExtentY        =   5054
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin GridEX20.GridEX GridEX1 
            Height          =   2835
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   9795
            _ExtentX        =   17277
            _ExtentY        =   5001
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            MethodHoldFields=   -1  'True
            AllowColumnDrag =   0   'False
            AllowEdit       =   0   'False
            GroupByBoxVisible=   0   'False
            DataMode        =   99
            HeaderFontName  =   "AngsanaUPC"
            FontSize        =   12
            ColumnHeaderHeight=   480
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            ColumnsCount    =   2
            Column(1)       =   "frmQueryPartItem.frx":27A2
            Column(2)       =   "frmQueryPartItem.frx":286A
            FormatStylesCount=   5
            FormatStyle(1)  =   "frmQueryPartItem.frx":290E
            FormatStyle(2)  =   "frmQueryPartItem.frx":2A6A
            FormatStyle(3)  =   "frmQueryPartItem.frx":2B1A
            FormatStyle(4)  =   "frmQueryPartItem.frx":2BCE
            FormatStyle(5)  =   "frmQueryPartItem.frx":2CA6
            ImageCount      =   0
            PrinterProperties=   "frmQueryPartItem.frx":2D5E
         End
      End
      Begin prjFarmManagement.uctlTextBox txtSocCode 
         Height          =   435
         Left            =   2400
         TabIndex        =   0
         Top             =   1050
         Width           =   2505
         _ExtentX        =   11615
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlFooter 
         Height          =   705
         Left            =   30
         TabIndex        =   9
         Top             =   5550
         Width           =   9990
         _ExtentX        =   17621
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin Threed.SSCommand cmdOK 
            Height          =   525
            Left            =   3375
            TabIndex        =   5
            Top             =   90
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   926
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmQueryPartItem.frx":2F36
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdExit 
            Cancel          =   -1  'True
            Height          =   525
            Left            =   5025
            TabIndex        =   6
            Top             =   90
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   926
            _Version        =   131073
            ButtonStyle     =   3
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   8
         Top             =   30
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtMinimum 
         Height          =   435
         Left            =   2400
         TabIndex        =   1
         Top             =   1500
         Width           =   4905
         _ExtentX        =   6165
         _ExtentY        =   767
      End
      Begin VB.Label lblPartType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   180
         TabIndex        =   13
         Top             =   2010
         Width           =   2145
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   8130
         TabIndex        =   3
         Top             =   990
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmQueryPartItem.frx":3250
         ButtonStyle     =   3
      End
      Begin VB.Label lblMinimum 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   1620
         Width           =   2175
      End
      Begin VB.Label lblSocCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   90
         TabIndex        =   10
         Top             =   1140
         Width           =   2235
      End
   End
End
Attribute VB_Name = "frmQueryPartItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_PartItem As CPartItem

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public id As Long
Public TempCollection As Collection

Private Sub InitGrid1()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.NAME = GLB_FONT
   GridEX1.ColumnHeaderFont.Size = 16
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = MapText("ID1")

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = MapText("ID2")

   Set Col = GridEX1.Columns.add '3
   Col.Width = 2055
   Col.Caption = MapText("รหัสสินค้า/วัตถุดิบ")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 7410
   Col.Caption = MapText("ชื่อสินค้า/วัตถุดิบ")
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim D As CSocFeature
Dim Cm As CPartItem

   If Flag Then
      Call EnableForm(Me, False)
      
      Set Cm = New CPartItem
      Cm.PART_ITEM_ID = -1
      Cm.PART_NO = PatchWildCard(txtSocCode.Text)
      Cm.PART_DESC = PatchWildCard(txtMinimum.Text)
      Cm.PART_TYPE = cboPartType.ItemData(Minus2Zero(cboPartType.ListIndex))
      Cm.OrderBy = 3
      Cm.OrderType = 1
      If Not glbDaily.QueryPartItem(Cm, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Set Cm = Nothing
         Exit Sub
      End If
      
      Set Cm = Nothing
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Call EnableForm(Me, True)
   End If
   
   If ItemCount > 0 Then
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cboFeatureLevel_Click()
   m_HasModify = True
End Sub

Private Sub cboGroup_Click()
   m_HasModify = True
End Sub

Private Sub chkEnable_Click()
   m_HasModify = True
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim Sf As CSocFeature
Dim TempID As Long
Dim ItemCount As Long
Dim C As CPartItem

'   If Not m_HasModify Then
'      SaveData = True
'      Exit Function
'   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Function
   End If

   TempID = Val(GridEX1.Value(1))
   
   If ShowMode = SHOW_ADD Then
      Set C = New CPartItem
      C.PART_ITEM_ID = TempID
      Call glbDaily.QueryPartItem(C, m_Rs, ItemCount, IsOK, glbErrorLog)
      If Not m_Rs.EOF Then
         Call C.PopulateFromRS(1, m_Rs)
         Call TempCollection.add(C)
      End If
      Set C = Nothing
   End If
      
   Call EnableForm(Me, True)
   SaveData = True
End Function

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Sub cmdSearch_Click()
   Call QueryData(True)
End Sub

Private Sub Form_Activate()
Dim Sp As CSystemParam
Dim FeatureTypeID As Long

   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadPartType(cboPartType)
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         Call QueryData(True)
      Else
         Call QueryData(False)
      End If
      
      m_HasModify = False
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.NAME
      glbErrorLog.ShowUserError
   ElseIf Shift = 1 And KeyCode = 112 Then
      If glbUser.EXCEPTION_FLAG = "Y" Then
         glbUser.EXCEPTION_FLAG = "N"
      Else
         glbUser.EXCEPTION_FLAG = "Y"
      End If
   ElseIf Shift = 0 And KeyCode = 116 Then
      Call cmdSearch_Click
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
   ElseIf Shift = 0 And KeyCode = 118 Then
'      Call cmdAdd_Click
'   ElseIf Shift = 0 And KeyCode = 117 Then
'      Call cmdDelete_Click
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
   ElseIf Shift = 0 And KeyCode = 114 Then
'      Call cmdEdit_Click
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
   Set m_PartItem = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub InitFormLayout()
   HeaderText = MapText("ค้นหาข้อมูลสินค้า/วัตถุดิบ")
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   Me.KeyPreview = True
   
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   pnlFooter.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame2.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Call InitHeaderFooter(pnlHeader, pnlFooter)
      
   Call txtSocCode.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call InitNormalLabel(lblPartType, MapText("ประเภทสินค้า/วัตถุดิบ"))
      
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา (F5)"))
   
   Call txtSocCode.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtMinimum.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   
   Call InitNormalLabel(lblSocCode, MapText("รหัสสินค้า/วัตถุดิบ"))
   Call InitNormalLabel(lblMinimum, MapText("ชื่อสินค้า/วัตถุดิบ"))
   Call InitNormalLabel(lblPartType, MapText("ประเภทสินค้า/วัตถุดิบ"))
   
   Call InitCombo(cboPartType)
   Call InitGrid1
End Sub

Private Sub cmdExit_Click()
'   If Not ConfirmExit(m_HasModify) Then
'      Exit Sub
'   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   OKClick = False
   Call InitFormLayout
      
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_PartItem = New CPartItem
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub


Private Sub txtFeatureCode_Change()
   m_HasModify = True
End Sub

Private Sub txtPassword_Change()
   m_HasModify = True
End Sub

Private Sub Label2_Click()

End Sub

Private Sub GridEX1_DblClick()
   Call cmdOK_Click
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
   Call m_PartItem.PopulateFromRS(1, m_Rs)
   
   Values(1) = m_PartItem.PART_ITEM_ID
   Values(2) = RealIndex
   Values(3) = m_PartItem.PART_NO
   Values(4) = m_PartItem.PART_DESC
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub txtAC_Change()
   m_HasModify = True
End Sub

Private Sub txtMinimum_Change()
   m_HasModify = True
End Sub

Private Sub txtOC_Change()
   m_HasModify = True
End Sub

Private Sub txtRate_Change()
   m_HasModify = True
End Sub

Private Sub txtRC_Change()
   m_HasModify = True
End Sub

Private Sub txtRoundingFactor_Change()
   m_HasModify = True
End Sub

Private Sub txtSocCode_Change()
   m_HasModify = True
End Sub

Private Sub uctlExpireDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlTime1_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlTime2_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlFeatureLookup_Change()
   m_HasModify = True
End Sub
