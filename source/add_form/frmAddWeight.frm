VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddWeight 
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15465
   Icon            =   "frmAddWeight.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   15465
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8520
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   15495
      _ExtentX        =   27331
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextBox txtSupplierCode 
         Height          =   495
         Left            =   1680
         TabIndex        =   14
         Top             =   1440
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   873
      End
      Begin prjFarmManagement.uctlDate uctlWeightDate 
         Height          =   405
         Left            =   1680
         TabIndex        =   11
         Top             =   960
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   120
         Top             =   720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   4695
         Left            =   120
         TabIndex        =   1
         Top             =   2520
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   8281
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
         Column(1)       =   "frmAddWeight.frx":27A2
         Column(2)       =   "frmAddWeight.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddWeight.frx":290E
         FormatStyle(2)  =   "frmAddWeight.frx":2A6A
         FormatStyle(3)  =   "frmAddWeight.frx":2B1A
         FormatStyle(4)  =   "frmAddWeight.frx":2BCE
         FormatStyle(5)  =   "frmAddWeight.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddWeight.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   15525
         _ExtentX        =   27384
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtTruckNo 
         Height          =   495
         Left            =   1680
         TabIndex        =   15
         Top             =   1920
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   873
      End
      Begin Threed.SSCommand cmdEditTempWeight 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   8280
         TabIndex        =   19
         Top             =   7440
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSetTempWeight 
         Height          =   525
         Left            =   6600
         TabIndex        =   18
         Top             =   7440
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSetEmpty 
         Height          =   525
         Left            =   4920
         TabIndex        =   17
         Top             =   7440
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   5760
         TabIndex        =   16
         Top             =   1560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddWeight.frx":2F36
         ButtonStyle     =   3
      End
      Begin VB.Label lblTruckNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   360
         TabIndex        =   13
         Top             =   2040
         Width           =   1155
      End
      Begin VB.Label lblSupplierCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   360
         TabIndex        =   12
         Top             =   1560
         Width           =   1155
      End
      Begin Threed.SSCommand cmdSelectAll 
         Height          =   525
         Left            =   2640
         TabIndex        =   3
         Top             =   4440
         Visible         =   0   'False
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddWeight.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSelect 
         Height          =   525
         Left            =   1680
         TabIndex        =   2
         Top             =   4440
         Visible         =   0   'False
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddWeight.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   5760
         TabIndex        =   0
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddWeight.frx":3884
         ButtonStyle     =   3
      End
      Begin VB.Label lblWeightDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   360
         TabIndex        =   10
         Top             =   960
         Width           =   1155
      End
      Begin VB.Label Label4 
         Height          =   315
         Left            =   10920
         TabIndex        =   9
         Top             =   3420
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   570
         TabIndex        =   8
         Top             =   1320
         Visible         =   0   'False
         Width           =   1155
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3240
         TabIndex        =   4
         Top             =   7440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddWeight.frx":3B9E
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   525
         Left            =   9960
         TabIndex        =   5
         Top             =   7440
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddWeight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Rs2 As ADODB.Recordset
Private m_Weight As CWeight

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public CancelWeigth As Boolean
Public TempWeigth As Boolean
Public EditTempWeigth As Boolean
Public id As Long
Public TempCollection As Collection
Public SumLotAmount As Double

Private FileName As String
Private m_SumUnit As Double
Private m_TempCol1 As Collection
Public m_TempCol2 As Collection

Public SupplierID As Long
Public SupplierCode As String
Public DocumentType As Long
Public DocumentDate As Date
Public TruckNo As String
Private Sub PopulateDestColl()
Dim Ri As CSupItem
Dim D As CSupItem

   If TempCollection Is Nothing Then
      Exit Sub
   End If
   
   For Each Ri In TempCollection
      Set D = New CSupItem
      
      If Ri.Flag <> "D" Then
         Call D.CopyObject(1, Ri)
         Call m_TempCol2.add(D)
      End If
      
      Set D = Nothing
   Next Ri
End Sub

Private Function IsIn(TempCol As Collection, TempID As Long) As Boolean
Dim D As CSupItem
Dim Found As Boolean
   
   Found = False
   For Each D In TempCol
      If D.PO_ID = TempID Then
         Found = True
      End If
   Next D
   
   IsIn = Found
End Function

Private Sub GenerateSourceItem(Rs As ADODB.Recordset, Rs2 As ADODB.Recordset, TempCol As Collection)
Dim CW As CWeight

   Set m_TempCol1 = Nothing
   Set m_TempCol1 = New Collection
   While Not Rs2.EOF
      Set CW = New CWeight
      Call CW.PopulateFromRS(2, Rs2)
      Call TempCol.add(CW)
      Set CW = Nothing
      Rs2.MoveNext
   Wend
   
   While Not Rs.EOF
      Set CW = New CWeight
      Call CW.PopulateFromRS(1, Rs)
      Call TempCol.add(CW)
      Set CW = Nothing
      Rs.MoveNext
   Wend
   
   
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Not VerifyDate(lblWeightDate, uctlWeightDate, False) Then
      Exit Sub
   End If
   
   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
      m_Weight.Date1 = uctlWeightDate.ShowDate
      m_Weight.CUST_ID = Trim(txtSupplierCode.Text) 'SupplierCode
      m_Weight.TRUCK_ID = Trim(txtTruckNo.Text) 'TruckNo
   
      Call glbDaily.QueryLegacyWeight(m_Weight, m_Rs, m_Rs2, ItemCount, IsOK, glbErrorLog)
      
   End If

   If ItemCount > 0 Then
      Call GenerateSourceItem(m_Rs, m_Rs2, m_TempCol1)
      GridEX1.ItemCount = m_TempCol1.Count
      GridEX1.Rebind
   Else
      GridEX1.ItemCount = 0
      GridEX1.Rebind
   End If

   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cboGroup_Click()
   m_HasModify = True
End Sub

Private Sub chkEnable_Click()
   m_HasModify = True
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
      
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Call PopulateTempColl(SumLotAmount)

   Call EnableForm(Me, True)
   SaveData = True
End Function

Private Sub cboBusinessGroup_Click()
   m_HasModify = True
End Sub

Private Sub cboBusinessType_Click()
   m_HasModify = True
End Sub

Private Sub cboEnterpriseType_Click()
   m_HasModify = True
End Sub

Private Sub chkCommit_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkCommit_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub chkSaleFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkSaleFlag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub


Private Sub cmdClear_Click()
'   uctlWeightDate.ShowDate = -1
   txtSupplierCode.Text = ""
   txtTruckNo.Text = ""
End Sub

Private Sub cmdEditTempWeight_Click()
   EditTempWeigth = True
End Sub

Private Sub cmdOK_Click()
Dim TempID As Long
m_HasModify = True
TempID = GridEX1.row
Call CopyItem(m_TempCol1, m_TempCol2, TempID)

GridEX1.ItemCount = m_TempCol1.Count
GridEX1.Rebind

   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Sub cmdPictureAdd_Click()
On Error Resume Next
Dim strDescription As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Picture Files (*.jpg, *.gif)|*.jpg;*.gif"
   dlgAdd.DialogTitle = "Select Picture to Add to Database"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
   m_HasModify = True
End Sub

Private Sub cmdSearch_Click()
   Call QueryData(True)
End Sub

Public Sub CopyItem(TempCol1 As Collection, TempCol2 As Collection, id As Long)
Dim L As CWeight

   If id > 0 Then
      Set L = TempCol1(id)
      
      L.Flag = "A"
      Call TempCol2.add(L)
      TempCol1.Remove (id)
   End If
End Sub
Public Sub CopyAllItem(TempCol1 As Collection, TempCol2 As Collection)
Dim J As Long

   For J = 1 To TempCol1.Count
      TempCol1(J).Flag = "A"
      TempCol1(J).IncludeFlag = True
      Call TempCol2.add(TempCol1(J))
   Next J
   Set TempCol1 = Nothing
   Set TempCol1 = New Collection
End Sub
Private Sub cmdSelect_Click()
Dim TempID As Long
m_HasModify = True
TempID = GridEX1.row
Call CopyItem(m_TempCol1, m_TempCol2, TempID)

GridEX1.ItemCount = m_TempCol1.Count
GridEX1.Rebind
Call cmdOK_Click
End Sub

Private Sub cmdSelectAll_Click()
   m_HasModify = True
   Call CopyAllItem(m_TempCol1, m_TempCol2)

   GridEX1.ItemCount = m_TempCol1.Count
   GridEX1.Rebind

'   GridEX2.ItemCount = m_TempCol2.Count
'   GridEX2.Rebind
End Sub

Public Sub PopulateTempColl(Tempsum As Double)
Dim D As CWeight
Dim Ri As CWeight
Dim Sum As Double
   Sum = 0
   For Each D In m_TempCol2
      Set Ri = New CWeight '
            Ri.WEIGHT_ID = D.WEIGHT_ID
            Ri.WEIGHT1 = D.WEIGHT1
            Ri.WEIGHT2 = D.WEIGHT2
            Ri.NetWeight = D.NetWeight
            Ri.TRUCK_ID = D.TRUCK_ID
            Ri.Date1 = D.Date1
            Ri.Date2 = D.Date2
            Ri.DateShow1 = D.DateShow1
            Ri.DateShow2 = D.DateShow2
            Ri.Time1 = D.Time1
            Ri.Time2 = D.Time2
            Ri.REMARK = D.REMARK
            Ri.Flag = "A"
            Call TempCollection.add(Ri, "1")
      Set Ri = Nothing
   Next D
   
   Tempsum = Sum
End Sub

Private Sub cmdSetEmpty_Click()
   CancelWeigth = True
   Unload Me
End Sub

Private Sub cmdSetTempWeight_Click()
   TempWeigth = True
   Unload Me
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      ' Call PopulateDestColl
      
        Set glbDatabaseMngr2 = New clsDatabaseMngr
       If Not glbDatabaseMngr2.ConnectLegacyDatabase(glbParameterObj.DBFileACCESSS, glbParameterObj.UserNameAccess, glbParameterObj.PasswordAccess, glbErrorLog) Then
       End If
      uctlWeightDate.ShowDate = DocumentDate
      txtSupplierCode.Text = SupplierCode
      txtTruckNo.Text = TruckNo
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_Weight.QueryFlag = 1
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         m_Weight.QueryFlag = 0
         Call QueryData(True)
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

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   If m_Rs2.State = adStateOpen Then
      m_Rs2.Close
   End If
   Set m_Rs2 = Nothing
   
   Set m_Weight = Nothing
   Set m_TempCol1 = Nothing
   Set m_TempCol2 = Nothing
   Set glbDatabaseMngr2 = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub GridEX2_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
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
   Col.Width = 1600
   Col.Caption = "เลขที่ใบชั่ง"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 1200
   Col.Caption = "รหัสซัพฯ"

   Set Col = GridEX1.Columns.add '3
   Col.Width = 1500
   Col.Caption = MapText("วันที่เอกสาร")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 1200
   Col.Caption = MapText("น้ำหนักเข้า")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 1200
   Col.Caption = MapText("น้ำหนักออก")
   
   Set Col = GridEX1.Columns.add '6
   Col.Width = 1200
   Col.Caption = MapText("น้ำหนักสุทธิ")
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = 2300
   Col.Caption = MapText("เวลาเข้า")
   
   Set Col = GridEX1.Columns.add '8
   Col.Width = 2300
   Col.Caption = MapText("เวลาออก")
   
   Set Col = GridEX1.Columns.add '9
   Col.Width = 2500
   Col.Caption = MapText("ทะเบียนรถ")
   
   Set Col = GridEX1.Columns.add '10
   Col.Width = 3500
   Col.Caption = MapText("หมายเหตุ")
End Sub

Private Sub GetTotalPrice()
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSetEmpty.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSetTempWeight.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEditTempWeight.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdClear.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSelect.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSelectAll.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call txtSupplierCode.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtTruckNo.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdSetEmpty, MapText("ยกเลิกน้ำหนัก"))
   Call InitMainButton(cmdSetTempWeight, MapText("น้ำหนักชั่วคราว"))
   Call InitMainButton(cmdEditTempWeight, MapText("แก้ไขน้ำหนัก"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา"))
   Call InitMainButton(cmdClear, MapText("เคลียร์ (F4)"))
   Call InitMainButton(cmdSelect, MapText(">"))
   Call InitMainButton(cmdSelectAll, MapText(">>"))
   
   Call InitNormalLabel(lblWeightDate, MapText("วันที่ชั่ง"))
   Call InitNormalLabel(lblSupplierCode, MapText("รหัสซัพฯ"))
   Call InitNormalLabel(lblTruckNo, MapText("ทะเบียนรถ"))
   
   Call InitGrid1
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
   Set m_Rs2 = New ADODB.Recordset
   Set m_Weight = New CWeight
   Set m_TempCol1 = New Collection
   Set m_TempCol2 = New Collection
   Set TempCollection = New Collection
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub GridEX1_DblClick()
 If GridEX1.RowCount > 0 Then
   Call cmdOK_Click
 End If
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"


   If m_TempCol1 Is Nothing Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If

   Dim CW As CWeight
   If m_TempCol1.Count <= 0 Then
      Exit Sub
   End If
   Set CW = GetItem(m_TempCol1, RowIndex, RealIndex)
   If CW Is Nothing Then
      Exit Sub
   End If

   Values(1) = CW.WEIGHT_ID
   Values(2) = RealIndex
   Values(3) = CW.CUST_ID
   Values(4) = CW.DateShow1
   Values(5) = CW.WEIGHT1
   Values(6) = CW.WEIGHT2
   Values(7) = CW.NetWeight
   Values(8) = CW.DateShow1 & " " & CW.Time1
   Values(9) = IIf(CW.DateShow2 = "", "", CW.DateShow2 & " " & CW.Time2)
   Values(10) = CW.TRUCK_ID
   Values(11) = CW.REMARK
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

'Private Sub GridEX2_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
'On Error GoTo ErrorHandler
'Dim RealIndex As Long
'
'   glbErrorLog.ModuleName = Me.NAME
'   glbErrorLog.RoutineName = "UnboundReadData"
'
'   If m_TempCol2 Is Nothing Then
'      Exit Sub
'   End If
'
'   If RowIndex <= 0 Then
'      Exit Sub
'   End If
'
'   Dim CR As CSupItem
'   If m_TempCol2.Count <= 0 Then
'      Exit Sub
'   End If
'   Set CR = GetItem(m_TempCol2, RowIndex, RealIndex)
'   If CR Is Nothing Then
'      Exit Sub
'   End If
'
'   Values(1) = CR.SUP_ITEM_ID
'   Values(2) = RealIndex
'   If CR.PIG_FLAG = "Y" Then
'      Values(3) = CR.ITEM_DESC
'   Else
'      Values(3) = CR.PART_DESC
'   End If
'   Values(4) = FormatNumber(CR.TOTAL_INCLUDE_PRICE)
'   Exit Sub
'
'ErrorHandler:
'   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
'   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'End Sub

Private Sub txtDoNo_Change()
   m_HasModify = True
End Sub

Private Sub txtDeliveryNo_Change()
   m_HasModify = True
End Sub

Private Sub txtSellBy_Change()
   m_HasModify = True
End Sub

Private Sub txtDocumentNo_Change()
   m_HasModify = True
End Sub

Private Sub txtPassword_Change()
   m_HasModify = True
End Sub

Private Sub txtSender_Change()
   m_HasModify = True
End Sub

Private Sub txtTotal_Change()
   m_HasModify = True
End Sub

Private Sub txtTruckNo_Change()
   m_HasModify = True
End Sub

Private Sub uctlSetupDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlTextBox1_Change()
   m_HasModify = True
End Sub

Private Sub uctlDeliveryLookup_Change()
   m_HasModify = True
End Sub

Private Sub SSCommand2_Click()

End Sub

Private Sub uctlDocumentDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlEmployeeLookup_Change()
   m_HasModify = True
End Sub
