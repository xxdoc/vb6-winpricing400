VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmLocation 
   BackColor       =   &H80000000&
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   8850
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   8850
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextBox txtTotalPallet 
         Height          =   375
         Left            =   1920
         TabIndex        =   8
         Top             =   800
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   6255
         Left            =   150
         TabIndex        =   7
         Top             =   1320
         Width           =   8505
         _ExtentX        =   15002
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
         Column(1)       =   "frmLocation.frx":0000
         Column(2)       =   "frmLocation.frx":00C8
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmLocation.frx":016C
         FormatStyle(2)  =   "frmLocation.frx":02C8
         FormatStyle(3)  =   "frmLocation.frx":0378
         FormatStyle(4)  =   "frmLocation.frx":042C
         FormatStyle(5)  =   "frmLocation.frx":0504
         ImageCount      =   0
         PrinterProperties=   "frmLocation.frx":05BC
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   3240
         TabIndex        =   10
         Top             =   840
         Width           =   465
      End
      Begin VB.Label lblPalletAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPalletAmount"
         Height          =   435
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1665
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   2
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   0
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   1
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   7080
         TabIndex        =   4
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   5400
         TabIndex        =   3
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_PalletDoc As CPalletDoc
Private m_TempPalletDoc As CPalletDoc
Public TempCollection As Collection
Public TempLotItemWh As CLotItemWH
Public m_CollLotExUse As Collection
Public m_CollPalletInLot As Collection
Private m_Rs As ADODB.Recordset
Private m_TableName As String
Public LocationGroup As String
Public LotNo As String
Public HeadertText As String
Public OKClick As Boolean
Public id As Long
Private m_HasModify As Boolean
Public BALANCE_FLAG As String
Public DOCUMENT_TYPE As Long
Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean

   If BALANCE_FLAG = "Y" Then
      MsgBox "เอกสารนี้ได้ถูกปรับยอดไปแล้ว ดูได้อย่างเดียวไม่สามารถเพิ่ม แก้ไข หรือลบ ได้อีกแล้ว", vbOKOnly, "แจ้งเตือน"
      Exit Sub
   End If
   
   frmAddEditLocation.HeaderText = MapText("เพิ่มข้อมูลพาเลท")
   Set frmAddEditLocation.TempCollection = TempCollection
   Set frmAddEditLocation.m_CollPalletInLot = m_CollPalletInLot
   Set frmAddEditLocation.TempLotItemWh = TempLotItemWh
    frmAddEditLocation.DocumentTypeInput = DOCUMENT_TYPE
   frmAddEditLocation.ShowMode = SHOW_ADD
   Load frmAddEditLocation
   frmAddEditLocation.Show 1

   OKClick = frmAddEditLocation.OKClick

   Unload frmAddEditLocation
   Set frmAddEditLocation = Nothing

   If OKClick Then
      Call QueryData(True)
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
   
'   If GridEX1.Value(5) = "N" Then
'      MsgBox "เอกสารนี้ได้ถูกปรับยอดไปแล้ว ดูได้อย่างเดียวไม่สามารถเพิ่ม แก้ไข หรือลบ ได้อีกแล้ว", vbOKOnly, "แจ้งเตือน"
'      Exit Sub
'   End If
   If BALANCE_FLAG = "Y" Then
      MsgBox "เอกสารนี้ได้ถูกปรับยอดไปแล้ว ดูได้อย่างเดียวไม่สามารถเพิ่ม แก้ไข หรือลบ ได้อีกแล้ว", vbOKOnly, "แจ้งเตือน"
      Exit Sub
   End If
   
    Dim t_LTD As CLotDoc
   Dim PalletNo As String
   PalletNo = GridEX1.Value(3)
   Set t_LTD = GetObject("CLotDoc", m_CollLotExUse, Trim(LotNo) & "-" & Trim(PalletNo), False)
   If Not (t_LTD Is Nothing) Then
      MsgBox "พาเลท " & PalletNo & " ของล๊อต : " & LotNo & " ได้ถูกตัดจ่ายไปแล้วจำนวน " & t_LTD.CAPACITY_AMOUNT & " ถุง " & vbNewLine & "ไม่สามารถลบได้ในขณะนี้ กรุณาไปแก้ไขข้อมูลการตัดจ่ายก่อน", vbOKOnly, "แจ้งเตือน"
      Exit Sub
   End If
   
   If Not ConfirmDelete(GridEX1.Value(3)) Then
      Exit Sub
   End If

   ID1 = GridEX1.Value(1)
   ID2 = GridEX1.Value(2)

   If ID1 <= 0 Then
      TempCollection.Remove (ID2)
   Else
      TempCollection.Item(ID2).Flag = "D"
   End If

   Call QueryData(True)
   m_HasModify = True
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim id As Long
Dim OKClick As Boolean

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
   If BALANCE_FLAG = "Y" Then
      MsgBox "เอกสารนี้ได้ถูกปรับยอดไปแล้ว ดูได้อย่างเดียวไม่สามารถเพิ่ม แก้ไข หรือลบ ได้อีกแล้ว", vbOKOnly, "แจ้งเตือน"
      Exit Sub
   End If
   
   Dim t_LTD As CLotDoc
   Dim PalletNo As String
   PalletNo = GridEX1.Value(3)
   Set t_LTD = GetObject("CLotDoc", m_CollLotExUse, Trim(LotNo) & "-" & Trim(PalletNo), False)
   If Not (t_LTD Is Nothing) Then
      MsgBox "พาเลท " & PalletNo & " ของล๊อต : " & LotNo & " ได้ถูกตัดจ่ายไปแล้วจำนวน " & t_LTD.CAPACITY_AMOUNT & " ถุง " & vbNewLine & "ไม่สามารถแก้ไขได้ในขณะนี้ กรุณาไปแก้ไขข้อมูลการตัดจ่ายก่อน", vbOKOnly, "แจ้งเตือน"
      Exit Sub
   End If

   id = Val(GridEX1.Value(2))
'   Call glbDatabaseMngr.LockTable(m_TableName, ID, IsCanLock, glbErrorLog)

   frmAddEditLocation.id = id
   frmAddEditLocation.HeaderText = MapText("แก้ไขพาเลท")
   Set frmAddEditLocation.TempCollection = TempCollection
   Set frmAddEditLocation.m_CollPalletInLot = m_CollPalletInLot
   Set frmAddEditLocation.TempLotItemWh = TempLotItemWh
   frmAddEditLocation.DocumentTypeInput = DOCUMENT_TYPE
   frmAddEditLocation.ShowMode = SHOW_EDIT
   Load frmAddEditLocation
   frmAddEditLocation.Show 1

   OKClick = frmAddEditLocation.OKClick

   Unload frmAddEditLocation
   Set frmAddEditLocation = Nothing

   If OKClick Then
      Call QueryData(True)
   End If
'   Call glbDatabaseMngr.UnLockTable(m_TableName, ID, IsCanLock, glbErrorLog)
End Sub

Private Sub cmdOK_Click()
'  ''Debug.Print TempCollection.Count
   OKClick = True
   Unload Me
End Sub

Private Sub QueryData(Flag As Boolean)

   txtTotalPallet.Text = FormatNumber(GetTotalAmountPallet(TempCollection), 0)
   GridEX1.ItemCount = CountItem(TempCollection)
   GridEX1.Rebind
   
   Call EnableForm(Me, True)
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
   Col.Width = 1500
   Col.Caption = MapText("เลขพาเลท")
      
   Set Col = GridEX1.Columns.add '4
   Col.Width = 2000
   Col.Caption = "จำนวนรับเข้า"
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 10
   Col.Caption = MapText("สถานะการปรับยอด")
   
   GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = MapText("ข้อมูลสถานที่จัดเก็บ")
   pnlHeader.Caption = MapText("ข้อมูลสถานที่จัดเก็บ")
   
   Call InitGrid
   
   
   Call InitNormalLabel(lblPalletAmount, MapText("จำนวนทั้งหมด"))
   Call InitNormalLabel(Label1, MapText("ถุง"))
   
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
End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   Set m_PalletDoc = New CPalletDoc
   Set m_TempPalletDoc = New CPalletDoc
   Set m_Rs = New ADODB.Recordset
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_PalletDoc = Nothing
   Set m_TempPalletDoc = Nothing
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub GridEX1_DblClick()

'If GridEX1.row = GridEX1.RowCount Then

   Call cmdEdit_Click
'End If
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"

      If TempCollection Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If
      
      Dim CR As CPalletDoc
      If TempCollection.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(TempCollection, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If
      
      Values(1) = CR.PALLET_DOC_ID
      Values(2) = RealIndex
      Values(3) = CR.PALLET_DOC_NO
      Values(4) = CR.CAPACITY_AMOUNT
      Values(5) = CR.BALANCE_FLAG
      Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'On Error GoTo ErrorHandler
'Dim RealIndex As Long
'
'   glbErrorLog.ModuleName = Me.NAME
'   glbErrorLog.RoutineName = "UnboundReadData"
'
'   If RowIndex <= 0 Then
'      Exit Sub
'   End If
'    Dim Cpd As CPalletDoc
'      If TempCollection.Count <= 0 Then
'         Exit Sub
'      End If
'    Set Cpd = GetItem(TempCollection, RowIndex, RealIndex)
'      If Cpd Is Nothing Then
'         Exit Sub
'      End If
'
'   'Set m_TempPalletDoc = TempCollection.Item(RowIndex)
'
'   Values(1) = Cpd.PALLET_DOC_ID
'   Values(2) = RowIndex
'   Values(3) = Cpd.PALLET_DOC_NO
'   Values(4) = Cpd.CAPACITY_AMOUNT
'
'   Exit Sub
'
'ErrorHandler:
'   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
'   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.HEIGHT = ScaleHeight
   
   pnlHeader.Width = ScaleWidth
   GridEX1.Width = ScaleWidth - 2 * GridEX1.Left
   GridEX1.HEIGHT = ScaleHeight - GridEX1.Top - 620
   cmdAdd.Top = ScaleHeight - 580
   cmdEdit.Top = ScaleHeight - 580
   cmdDelete.Top = ScaleHeight - 580
   cmdOK.Top = ScaleHeight - 580
   cmdExit.Top = ScaleHeight - 580
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
End Sub
Function GetTotalAmountPallet(Cl As Collection) As Double
Dim PD As CPalletDoc
Dim SumAmount As Double
   SumAmount = 0
   For Each PD In Cl
      If PD.Flag <> "D" Then
         SumAmount = SumAmount + PD.CAPACITY_AMOUNT
      End If
   Next PD
   GetTotalAmountPallet = SumAmount
End Function
