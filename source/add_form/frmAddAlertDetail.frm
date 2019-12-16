VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAddAlertDetail 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddAlertDetail.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8520
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboUserGroup 
         Height          =   315
         Left            =   1770
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   960
         Width           =   2955
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
         TabIndex        =   9
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         Enabled         =   0   'False
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   6075
         Left            =   120
         TabIndex        =   2
         Top             =   1650
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   10716
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
         Column(1)       =   "frmAddAlertDetail.frx":27A2
         Column(2)       =   "frmAddAlertDetail.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddAlertDetail.frx":290E
         FormatStyle(2)  =   "frmAddAlertDetail.frx":2A6A
         FormatStyle(3)  =   "frmAddAlertDetail.frx":2B1A
         FormatStyle(4)  =   "frmAddAlertDetail.frx":2BCE
         FormatStyle(5)  =   "frmAddAlertDetail.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddAlertDetail.frx":2D5E
      End
      Begin GridEX20.GridEX GridEX2 
         Height          =   6075
         Left            =   6360
         TabIndex        =   5
         Top             =   1650
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   10716
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
         Column(1)       =   "frmAddAlertDetail.frx":2F36
         Column(2)       =   "frmAddAlertDetail.frx":2FFE
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddAlertDetail.frx":30A2
         FormatStyle(2)  =   "frmAddAlertDetail.frx":31FE
         FormatStyle(3)  =   "frmAddAlertDetail.frx":32AE
         FormatStyle(4)  =   "frmAddAlertDetail.frx":3362
         FormatStyle(5)  =   "frmAddAlertDetail.frx":343A
         ImageCount      =   0
         PrinterProperties=   "frmAddAlertDetail.frx":34F2
      End
      Begin VB.Label lblUserGroup 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   1575
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   9240
         TabIndex        =   1
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSelectAll 
         Height          =   525
         Left            =   5680
         TabIndex        =   4
         Top             =   5280
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSelect 
         Height          =   525
         Left            =   5680
         TabIndex        =   3
         Top             =   4530
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   7560
         TabIndex        =   0
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   4320
         TabIndex        =   6
         Top             =   7860
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
         Left            =   5970
         TabIndex        =   7
         Top             =   7860
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddAlertDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Private m_UserAccount As CUserAccount

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public TempCollection As Collection

Private m_TempCol1 As Collection
Private m_TempCol2 As Collection
Private Sub PopulateDestColl()
Dim Ri As CAlertDetail
Dim D As CUserAccount
   
   For Each Ri In TempCollection
      Set D = New CUserAccount

      If Ri.Flag <> "D" Then
         D.USER_ID = Ri.USER_ID
         D.USER_NAME = Ri.USER_NAME
         Call m_TempCol2.add(D)
      End If

      Set D = Nothing
   Next Ri
End Sub

Private Function IsIn(TempCol As Collection, TempID As Long) As Boolean
Dim D As CUserAccount
Dim Found As Boolean

   Found = False
   For Each D In TempCol
      If D.USER_ID = TempID Then
         Found = True
         Exit For
      End If
   Next D

   IsIn = Found
End Function

Private Sub GenerateSourceItem(Rs As ADODB.Recordset, TempCol As Collection)
Dim Bd As CUserAccount
Dim RefCount As Long
Dim Found As Boolean

   Set m_TempCol1 = Nothing
   Set m_TempCol1 = New Collection
   While Not Rs.EOF
      Set Bd = New CUserAccount
      Call Bd.PopulateFromRS(1, Rs)
      
      If Bd.USER_ID > 0 Then
         If IsIn(m_TempCol2, Bd.USER_ID) Then
            Found = True
         Else
            Found = False
         End If
   
         If Not Found Then
            Call TempCol.add(Bd)
         End If
      End If

      Set Bd = Nothing
      Rs.MoveNext
   Wend
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)

      m_UserAccount.USER_ID = -1
      m_UserAccount.GROUP_ID = cboUserGroup.ItemData(Minus2Zero(cboUserGroup.ListIndex))
      
      If Not glbAdmin.QueryUserAccount(m_UserAccount, m_Rs, itemcount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If

   If itemcount > 0 Then
      Call GenerateSourceItem(m_Rs, m_TempCol1)
      GridEX1.itemcount = m_TempCol1.Count
      GridEX1.Rebind
   Else
      GridEX1.itemcount = 0
      GridEX1.Rebind
   End If

   GridEX2.itemcount = m_TempCol2.Count
   GridEX2.Rebind

   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If

   Call EnableForm(Me, True)
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean

   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If

   Call PopulateTempColl

   Call EnableForm(Me, True)
   SaveData = True
End Function

Private Sub cmdClear_Click()
   cboUserGroup.ListIndex = -1
End Sub

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

Public Sub CopyItem(TempCol1 As Collection, TempCol2 As Collection, ID As Long)
Dim L As CUserAccount

   If ID > 0 Then
      TempCol1(ID).Flag = "A"
      Call TempCol2.add(TempCol1(ID))
      TempCol1.Remove (ID)
   End If
End Sub

Public Sub CopyAllItem(TempCol1 As Collection, TempCol2 As Collection)
Dim j As Long

   For j = 1 To TempCol1.Count
      TempCol1(j).Flag = "A"
      Call TempCol2.add(TempCol1(j))
   Next j
   Set TempCol1 = Nothing
   Set TempCol1 = New Collection
End Sub

Private Sub cmdSelect_Click()
Dim TempID As Long
Dim check As CUserAccount
Dim ID As Long
Dim I As Long
Dim row As Long
   If m_TempCol1.Count <= 0 Then
      Exit Sub
   End If

   m_HasModify = True
   'Id = GridEX1.Value(1)
   For row = 1 To GridEX1.RowCount
      I = 0
      If GridEX1.RowSelected(row) = True Then
         ID = GridEX1.GetRowData(row).Value(1)
         For Each check In m_TempCol1
           I = I + 1
           If check.USER_ID = ID Then
              TempID = I
           End If
         Next check
   
          Call CopyItem(m_TempCol1, m_TempCol2, TempID)
       End If

   Next row

   GridEX1.itemcount = m_TempCol1.Count
   GridEX1.Rebind

   GridEX2.itemcount = m_TempCol2.Count
   GridEX2.Rebind

End Sub

Private Sub cmdSelectAll_Click()
   m_HasModify = True
   Call CopyAllItem(m_TempCol1, m_TempCol2)

   GridEX1.itemcount = m_TempCol1.Count
   GridEX1.Rebind

   GridEX2.itemcount = m_TempCol2.Count
   GridEX2.Rebind
End Sub

Public Sub PopulateTempColl()
Dim D As CUserAccount
Dim Ri As CAlertDetail

   For Each D In m_TempCol2
      Set Ri = New CAlertDetail

      If D.Flag = "A" Then
         Ri.Flag = "A"
         Ri.USER_ID = D.USER_ID
         Ri.USER_NAME = D.USER_NAME
         
         Ri.READ_FLAG = "N"
         
         Call TempCollection.add(Ri)
      End If

      Set Ri = Nothing
   Next D
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadUserGroup(cboUserGroup)
      
      Call EnableForm(Me, False)
      Call PopulateDestColl
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_UserAccount.QueryFlag = 1
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         m_UserAccount.QueryFlag = 0
         Call QueryData(True)
      End If

      Call EnableForm(Me, True)
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

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_UserAccount = Nothing
   Set m_TempCol1 = Nothing
   Set m_TempCol2 = Nothing
   
End Sub
Private Sub InitGrid1()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.itemcount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.Name = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation

   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
   
   '==
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2500
   Col.Caption = MapText("บัญชีผู้ใช้งาน")
End Sub


Private Sub InitGrid2()
Dim Col As JSColumn

   GridEX2.Columns.Clear
   GridEX2.BackColor = GLB_GRID_COLOR
   GridEX2.itemcount = 0
   GridEX2.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX2.ColumnHeaderFont.Bold = True
   GridEX2.ColumnHeaderFont.Name = GLB_FONT
   GridEX2.TabKeyBehavior = jgexControlNavigation

   Set Col = GridEX2.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX2.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   '==
   Set Col = GridEX2.Columns.add '3
   Col.Width = 2500
   Col.Caption = MapText("บัญชีผู้ใช้งาน")
End Sub


Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)

   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText

   Call InitNormalLabel(lblUserGroup, MapText("กลุ่มผู้ใช้งาน"))
   
   Call InitCombo(cboUserGroup)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)

   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19

   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSelect.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSelectAll.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdClear.Picture = LoadPicture(glbParameterObj.NormalButton1)
  
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา (F5)"))
   Call InitMainButton(cmdClear, MapText("เคลียร์ (F4)"))
   Call InitMainButton(cmdSelect, MapText("->"))
   Call InitMainButton(cmdSelectAll, MapText(">>"))
   
   Call InitGrid1
   Call InitGrid2
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
   
   Set m_UserAccount = New CUserAccount
   Set m_TempCol1 = New Collection
   Set m_TempCol2 = New Collection
End Sub
Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"


   If m_TempCol1 Is Nothing Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If

   Dim CR As CUserAccount
   If m_TempCol1.Count <= 0 Then
      Exit Sub
   End If
   Set CR = GetItem(m_TempCol1, RowIndex, RealIndex)
   If CR Is Nothing Then
      Exit Sub
   End If

      Values(1) = CR.USER_ID
      Values(2) = RealIndex
      Values(3) = CR.USER_NAME
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub GridEX2_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"


   If m_TempCol2 Is Nothing Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If

   Dim CR As CUserAccount
   If m_TempCol2.Count <= 0 Then
      Exit Sub
   End If
   Set CR = GetItem(m_TempCol2, RowIndex, RealIndex)
   If CR Is Nothing Then
      Exit Sub
   End If

      Values(1) = CR.USER_ID
      Values(2) = RealIndex
      Values(3) = CR.USER_NAME
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

