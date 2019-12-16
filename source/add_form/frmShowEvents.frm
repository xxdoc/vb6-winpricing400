VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmShowEvents 
   BackColor       =   &H80000000&
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   16200
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   16200
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   16155
      _ExtentX        =   28496
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboTxType 
         Height          =   315
         Left            =   11760
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   840
         Width           =   1335
      End
      Begin prjFarmManagement.uctlTextBox txtName 
         Height          =   375
         Left            =   1920
         TabIndex        =   4
         Top             =   795
         Width           =   7935
         _extentx        =   2143
         _extenty        =   661
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   10
         TabIndex        =   2
         Top             =   0
         Width           =   16125
         _ExtentX        =   28443
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   6255
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   15705
         _ExtentX        =   27702
         _ExtentY        =   11033
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
         Column(1)       =   "frmShowEvents.frx":0000
         Column(2)       =   "frmShowEvents.frx":00C8
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmShowEvents.frx":016C
         FormatStyle(2)  =   "frmShowEvents.frx":02C8
         FormatStyle(3)  =   "frmShowEvents.frx":0378
         FormatStyle(4)  =   "frmShowEvents.frx":042C
         FormatStyle(5)  =   "frmShowEvents.frx":0504
         ImageCount      =   0
         PrinterProperties=   "frmShowEvents.frx":05BC
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   13440
         TabIndex        =   8
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin VB.Label lblTxType 
         Alignment       =   1  'Right Justify
         Caption         =   "lblName"
         Height          =   435
         Left            =   9960
         TabIndex        =   7
         Top             =   840
         Width           =   1665
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Caption         =   "lblName"
         Height          =   435
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   1665
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   7320
         TabIndex        =   0
         Top             =   7800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmShowEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Public TempCollection As Collection
Private m_Rs As ADODB.Recordset
Private m_TableName As String
Public HeadertText As String
Public OKClick As Boolean
Public id As Long
Private m_HasModify As Boolean

Public ID_LOT As Long
Public HeadPackNo As Long
Public LotItemWhId As Long
Public LotId As Long
Public LotDocId As Long
Public LotDocIdRef As Long
Public DOCUMENT_TYPE_INPUT As Long
Public PART_ITEM_ID As Long
Public PART_NO As String
Public LOCATION_ID As Long
Public KeyType As Long
Private m_Coll As Collection

Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub

Private Sub QueryData(Flag As Boolean)
  txtName.Text = PART_NO
   If cboTxType.ListIndex = 1 Then
     Call LoadPalletDoc(Nothing, m_Coll, LotId, 8, , 5, "I", , , , PART_ITEM_ID, DOCUMENT_TYPE_INPUT, , LOCATION_ID)
  ElseIf cboTxType.ListIndex = 2 Then
     Call LoadPalletDoc(Nothing, m_Coll, LotId, 8, , 5, "E", , , , PART_ITEM_ID, , LotDocIdRef, LOCATION_ID)
  End If
   GridEX1.ItemCount = CountItem(m_Coll)
   GridEX1.Rebind
   
   Call EnableForm(Me, True)
End Sub

Private Sub cmdSearch_Click()
  Call QueryData(True)
End Sub

Private Sub Form_Activate()
Dim ItemCount As Long
   If Not m_HasActivate Then
      Me.Refresh
      DoEvents
      
     Call QueryData(True)
      m_HasActivate = True
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
   Col.Width = 0
   Col.Caption = "RealIndex"
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2000
   Col.Caption = "เอกสารอ้างอิง"
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 1500
   Col.Caption = MapText("PALLET_NO")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 2000
   Col.Caption = MapText("LOT_NO")
   
   Set Col = GridEX1.Columns.add '6
   Col.Width = 2000
   Col.Caption = "HEAD_PACK_NO"
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = 2000
   Col.Caption = "PART_ITEM_ID"
   
   Set Col = GridEX1.Columns.add '8
   Col.Width = 1000
   Col.Caption = "TX_TYPE"
   
   Set Col = GridEX1.Columns.add '9
   Col.Width = 2000
   Col.Caption = "BIN_NO"
   
   Set Col = GridEX1.Columns.add '10
   Col.Width = 2000
   Col.Caption = "จำนวนบรรจุ"
   
   
   GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = MapText(HeadertText)
   pnlHeader.Caption = MapText(HeadertText)
   
   Call InitGrid
   
   Call InitCombo(cboTxType)
   Call InitTxType(cboTxType)
   
   Call InitNormalLabel(lblName, MapText("รายละเอียด"))
   Call InitNormalLabel(lblTxType, MapText("ประเภท"))
   
   Call txtName.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdSearch, MapText("ค้นหา (F5)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
End Sub



Private Sub Form_Load()
   Set m_Rs = New ADODB.Recordset
   Set m_Coll = New Collection
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set m_Coll = Nothing
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
End Sub


Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"

      If m_Coll Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If
      
      Dim CR As CPalletDoc
      If m_Coll.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_Coll, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If
      
      Values(1) = CR.PALLET_DOC_ID
      Values(2) = RealIndex
      Values(3) = CR.DOCUMENT_NO
      Values(4) = CR.PALLET_DOC_NO
      Values(5) = CR.LOT_NO
      Values(6) = CR.HEAD_PACK_NO
      Values(7) = CR.PART_ITEM_ID
      Values(8) = CR.TX_TYPE
      Values(9) = CR.BIN_NO
      Values(10) = CR.CAPACITY_AMOUNT
      Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)

End Sub
Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.HEIGHT = ScaleHeight
   
   pnlHeader.Width = ScaleWidth
   GridEX1.Width = ScaleWidth - 2 * GridEX1.Left
   GridEX1.HEIGHT = ScaleHeight - GridEX1.Top - 620
   cmdOK.Top = ScaleHeight - 580
   cmdOK.Left = cmdOK.Width - 50
End Sub

