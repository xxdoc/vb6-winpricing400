VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAddPOSupItem 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddPOSupItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
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
      TabIndex        =   9
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   6855
         Left            =   150
         TabIndex        =   3
         Top             =   870
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   12091
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
         Column(1)       =   "frmAddPOSupItem.frx":27A2
         Column(2)       =   "frmAddPOSupItem.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddPOSupItem.frx":290E
         FormatStyle(2)  =   "frmAddPOSupItem.frx":2A6A
         FormatStyle(3)  =   "frmAddPOSupItem.frx":2B1A
         FormatStyle(4)  =   "frmAddPOSupItem.frx":2BCE
         FormatStyle(5)  =   "frmAddPOSupItem.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddPOSupItem.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX2 
         Height          =   6855
         Left            =   6540
         TabIndex        =   6
         Top             =   870
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   12091
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
         Column(1)       =   "frmAddPOSupItem.frx":2F36
         Column(2)       =   "frmAddPOSupItem.frx":2FFE
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddPOSupItem.frx":30A2
         FormatStyle(2)  =   "frmAddPOSupItem.frx":31FE
         FormatStyle(3)  =   "frmAddPOSupItem.frx":32AE
         FormatStyle(4)  =   "frmAddPOSupItem.frx":3362
         FormatStyle(5)  =   "frmAddPOSupItem.frx":343A
         ImageCount      =   0
         PrinterProperties=   "frmAddPOSupItem.frx":34F2
      End
      Begin prjFarmManagement.uctlDate uctlFromDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   0
         Top             =   930
         Visible         =   0   'False
         Width           =   3855
         _extentx        =   6800
         _extenty        =   714
      End
      Begin prjFarmManagement.uctlDate uctlDocumentDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   1
         Top             =   1350
         Visible         =   0   'False
         Width           =   3855
         _extentx        =   6800
         _extenty        =   714
      End
      Begin Threed.SSCommand cmdSelectAll 
         Height          =   525
         Left            =   5648
         TabIndex        =   5
         Top             =   5040
         Visible         =   0   'False
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddPOSupItem.frx":36CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSelect 
         Height          =   525
         Left            =   5648
         TabIndex        =   4
         Top             =   4470
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddPOSupItem.frx":39E4
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   10140
         TabIndex        =   2
         Top             =   930
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddPOSupItem.frx":3CFE
         ButtonStyle     =   3
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   570
         TabIndex        =   13
         Top             =   900
         Width           =   1155
      End
      Begin VB.Label Label4 
         Height          =   315
         Left            =   11250
         TabIndex        =   12
         Top             =   3420
         Width           =   585
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   570
         TabIndex        =   11
         Top             =   1320
         Width           =   1155
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   4320
         TabIndex        =   7
         Top             =   7860
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddPOSupItem.frx":4018
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5970
         TabIndex        =   8
         Top             =   7860
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddPOSupItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_PO As CBillingDoc

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public TempCollection As Collection
Public SumLotAmount As Double

Private FileName As String
Private m_SumUnit As Double
Private m_TempCol1 As Collection
Private m_TempCol2 As Collection

Public SupplierID As Long
Public SupplierIDTrue As Long
Public DocumentType As Long
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

Private Sub GenerateSourceItem(Rs As ADODB.Recordset, TempCol As Collection)
Dim BD As CBillingDoc

   Set m_TempCol1 = Nothing
   Set m_TempCol1 = New Collection
   While Not Rs.EOF
      Set BD = New CBillingDoc
      Call BD.PopulateFromRS(1, Rs)
      
      If Not IsIn(m_TempCol2, BD.BILLING_DOC_ID) Then
         Call TempCol.add(BD)
      End If

      Set BD = Nothing
      Rs.MoveNext
   Wend
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)

      m_PO.BILLING_DOC_ID = -1
      m_PO.FROM_DATE = -1
      m_PO.TO_DATE = -1
      m_PO.SUPPLIER_ID = SupplierID
      m_PO.CLOSE_FLAG = "N"
      m_PO.PO_APPROVED_FLAG = "Y"
      m_PO.CheckTotalAmount = True
      If DocumentType = 100 Then
         m_PO.DOCUMENT_TYPE = 1000
      ElseIf DocumentType = 101 Then
         m_PO.DOCUMENT_TYPE = 1001
      ElseIf DocumentType = 102 Then
         m_PO.DOCUMENT_TYPE = 1002
      ElseIf DocumentType = 103 Then
         m_PO.DOCUMENT_TYPE = 1003
      End If
      Call glbDaily.QueryBillingDoc(m_PO, m_Rs, ItemCount, IsOK, glbErrorLog)
   End If

   If ItemCount > 0 Then
      Call GenerateSourceItem(m_Rs, m_TempCol1)
      GridEX1.ItemCount = m_TempCol1.Count
      GridEX1.Rebind
   Else
      GridEX1.ItemCount = 0
      GridEX1.Rebind
   End If
   
   If Not (TempCollection Is Nothing) Then
      GridEX2.ItemCount = TempCollection.Count
      GridEX2.Rebind
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


Private Sub cmdOK_Click()
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

Public Sub CopyItem(TempCol1 As Collection, TempCol2 As Collection, ID As Long)
Dim L As CBillingDoc
Dim OKClick As Boolean
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim Poi As CSupItem
Dim Di As CSupItem
Dim IsOK As Boolean
Dim TempColl As Collection
Dim GetUpdate  As CSupItem
   Set TempColl = New Collection
   Set TempRs = New ADODB.Recordset
   
   If ID > 0 Then
      Set L = TempCol1(ID)
      L.QueryFlag = 1
      
      SupplierIDTrue = L.SUPPLIER_ID_TRUE
      Call glbDaily.QueryBillingDoc(L, TempRs, iCount, IsOK, glbErrorLog)
      
      Call GetUpDateFromPo(L.BILLING_DOC_ID, TempColl)
      
      For Each Poi In L.SupItems
         Set Di = New CSupItem
         Call Di.CopyObject(1, Poi)
         Di.Flag = "A"
         Di.PO_ID = Poi.DO_ID
         Di.PO_NO = Poi.DOCUMENT_NO
         
         Set GetUpdate = GetObject("CSupItem", TempColl, Poi.PART_ITEM_ID, True)
         
         Di.TX_AMOUNT = Di.TX_AMOUNT - GetUpdate.TX_AMOUNT
         Di.WEIGHT_AMOUNT = Di.TX_AMOUNT
         Di.TOTAL_WEIGHT = Di.TX_AMOUNT
         Di.TOTAL_ACTUAL_PRICE = Di.TX_AMOUNT * Di.ACTUAL_UNIT_PRICE
         Di.TOTAL_INCLUDE_PRICE = Di.TX_AMOUNT * Di.ACTUAL_UNIT_PRICE
         
         Call TempCol2.add(Di)
         Set Di = Nothing
      Next Poi
      Call TempCol1.Remove(ID)
      
   End If
   
   If TempRs.State = adStateOpen Then
      Call TempRs.Close
   End If
   Set TempRs = Nothing
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
   If GridEX2.ItemCount <= 0 Then
      
      TempID = GridEX1.row
      Call CopyItem(m_TempCol1, m_TempCol2, TempID)
      
      GridEX1.ItemCount = m_TempCol1.Count
      GridEX1.Rebind
      
      GridEX2.ItemCount = m_TempCol2.Count
      GridEX2.Rebind
      
      Call cmdOK_Click
   End If
End Sub

Private Sub cmdSelectAll_Click()
   m_HasModify = True
   Call CopyAllItem(m_TempCol1, m_TempCol2)

   GridEX1.ItemCount = m_TempCol1.Count
   GridEX1.Rebind

   GridEX2.ItemCount = m_TempCol2.Count
   GridEX2.Rebind
End Sub

Public Sub PopulateTempColl(Tempsum As Double)
Dim D As CSupItem
Dim Ri As CSupItem
Dim Sum As Double

   Sum = 0
   For Each D In m_TempCol2
      Set Ri = New CSupItem

      If (D.Flag = "A") Then
         Call Ri.CopyObject(1, D)
         If Val(Ri.TX_AMOUNT) > 0 Then
            Ri.PO_ID = D.PO_ID
            Ri.PO_NO = D.PO_NO
            Ri.Flag = "A"
            Call TempCollection.add(Ri)
         End If
      End If

      If D.Flag <> "D" Then
'         Sum = Sum + D.USED_AMOUNT
      End If
      
      Set Ri = Nothing
   Next D
   
   Tempsum = Sum
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      Call PopulateDestColl
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_PO.QueryFlag = 1
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         m_PO.QueryFlag = 0
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
   
   Set m_PO = Nothing
   Set m_TempCol1 = Nothing
   Set m_TempCol2 = Nothing
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
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   Set Col = GridEX1.Columns.add '3
   Col.Width = 1575
   Col.Caption = MapText("�ѹ����͡���")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 1740 + 1410
   Col.Caption = MapText("�����Ţ�͡���")

'   Set Col = GridEX1.Columns.Add '5
'   Col.Width = 1410
'   Col.TextAlignment = jgexAlignRight
'   Col.Caption = MapText("�ӹǹ��������")
End Sub


Private Sub InitGrid2()
Dim Col As JSColumn

   GridEX2.Columns.Clear
   GridEX2.BackColor = GLB_GRID_COLOR
   GridEX2.ItemCount = 0
   GridEX2.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX2.ColumnHeaderFont.Bold = True
   GridEX2.ColumnHeaderFont.NAME = GLB_FONT
   GridEX2.TabKeyBehavior = jgexControlNavigation

   Set Col = GridEX2.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX2.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   Set Col = GridEX2.Columns.add '3
   Col.Width = 3015
   Col.Caption = MapText("��������´")

   Set Col = GridEX2.Columns.add '4
   Col.Width = 1740
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("�Ҥ����")
End Sub

Private Sub GetTotalPrice()
'Dim II As CExportItem
'Dim Sum As Double
'
'   Sum = 0
'   For Each II In m_PO.ImportExports
'      If II.Flag <> "D" Then
'         Sum = Sum + CDbl(Format(II.EXPORT_AVG_PRICE, "0.00")) * CDbl(Format(II.EXPORT_AMOUNT, "0.00"))
'      End If
'   Next II
''
''   txtDeliveryFee.Text = Format(Sum, "0.00")
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblDocumentDate, MapText("�֧�ѹ���"))
   Call InitNormalLabel(lblFromDate, MapText("�ҡ�ѹ���"))
   Call InitNormalLabel(Label4, MapText("�ҷ"))
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSelect.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSelectAll.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
   Call InitMainButton(cmdSearch, MapText("����"))
   Call InitMainButton(cmdSelect, MapText(">"))
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
   Set m_PO = New CBillingDoc
   Set m_TempCol1 = New Collection
   Set m_TempCol2 = New Collection
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

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

   Dim CR As CBillingDoc
   If m_TempCol1.Count <= 0 Then
      Exit Sub
   End If
   Set CR = GetItem(m_TempCol1, RowIndex, RealIndex)
   If CR Is Nothing Then
      Exit Sub
   End If

   Values(1) = CR.BILLING_DOC_ID
   Values(2) = RealIndex
   Values(3) = DateToStringExtEx2(CR.DOCUMENT_DATE)
   Values(4) = CR.DOCUMENT_NO
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub GridEX2_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"

   If m_TempCol2 Is Nothing Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If

   Dim CR As CSupItem
   If m_TempCol2.Count <= 0 Then
      Exit Sub
   End If
   Set CR = GetItem(m_TempCol2, RowIndex, RealIndex)
   If CR Is Nothing Then
      Exit Sub
   End If

   Values(1) = CR.SUP_ITEM_ID
   Values(2) = RealIndex
   If CR.PIG_FLAG = "Y" Then
      Values(3) = CR.ITEM_DESC
   Else
      Values(3) = CR.PART_DESC
   End If
   Values(4) = FormatNumber(CR.TOTAL_INCLUDE_PRICE)
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

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
Public Sub GetUpDateFromPo(ID As Long, TmpCol As Collection)
Dim Doc As CSupItem
Dim m_Rs As ADODB.Recordset
Dim ItemCount As Long
Set Doc = New CSupItem
Set m_Rs = New ADODB.Recordset
Set TmpCol = New Collection

   Doc.PO_ID = ID
   Call Doc.QueryData(104, m_Rs, ItemCount)
   
   While Not m_Rs.EOF
      Set Doc = New CSupItem
      Call Doc.PopulateFromRS(104, m_Rs)
      Call TmpCol.add(Doc, Trim(str(Doc.PART_ITEM_ID)))
      m_Rs.MoveNext
   Wend
   
   Set m_Rs = Nothing
   Set Doc = Nothing
End Sub

