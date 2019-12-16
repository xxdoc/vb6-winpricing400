VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportInventoryAct 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13560
   Icon            =   "frmImportInventoryAct.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   13560
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   7845
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   13785
      _ExtentX        =   24315
      _ExtentY        =   13838
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextBox txtFileName 
         Height          =   435
         Left            =   1860
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2310
         Width           =   10515
         _ExtentX        =   18547
         _ExtentY        =   767
      End
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   1860
         TabIndex        =   5
         Top             =   2760
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   13755
         _ExtentX        =   24262
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtPercent 
         Height          =   465
         Left            =   1860
         TabIndex        =   6
         Top             =   3090
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   9780
         Top             =   750
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin prjFarmManagement.uctlDate uctlInventoryActDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   0
         Top             =   1005
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtFileName2 
         Height          =   435
         Left            =   1860
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1800
         Visible         =   0   'False
         Width           =   10515
         _ExtentX        =   18547
         _ExtentY        =   767
      End
      Begin VB.Label lblFileName2 
         Alignment       =   1  'Right Justify
         Caption         =   "lblFileName2"
         Height          =   435
         Left            =   210
         TabIndex        =   18
         Top             =   1830
         Visible         =   0   'False
         Width           =   1575
      End
      Begin Threed.SSCommand cmdFileName2 
         Height          =   405
         Left            =   12480
         TabIndex        =   17
         Top             =   1815
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportInventoryAct.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label lblNote 
         Caption         =   "Label1"
         Height          =   2955
         Left            =   480
         TabIndex        =   15
         Top             =   4320
         Width           =   12585
      End
      Begin VB.Label lblInventoryActDate 
         Alignment       =   1  'Right Justify
         Caption         =   "lblInventoryActDate"
         Height          =   315
         Left            =   480
         TabIndex        =   14
         Top             =   1095
         Width           =   1305
      End
      Begin Threed.SSCommand cmdFileName 
         Height          =   405
         Left            =   12480
         TabIndex        =   1
         Top             =   2310
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportInventoryAct.frx":2ABC
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   1890
         TabIndex        =   2
         Top             =   3630
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportInventoryAct.frx":2DD6
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3600
         TabIndex        =   12
         Top             =   3210
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   11
         Top             =   2820
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   10
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label lblFileName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   9
         Top             =   2340
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10935
         TabIndex        =   4
         Top             =   3630
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   9285
         TabIndex        =   3
         Top             =   3630
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportInventoryAct.frx":30F0
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmImportInventoryAct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Public InventoryActArea As Long

Private m_ExcelApp As Object
Private m_ExcelSheet As Object

Private PartUctlColls As Collection
Private PartColls As Collection
Private PartLabColls  As Collection
Private PartLabUpdateColls  As Collection

Private m_PartItems As Collection
Private Sub cmdFileName_Click()
 On Error Resume Next
Dim strDescription As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Access Files (*.XLS)|*..xls;*.XLS;"
   dlgAdd.DialogTitle = "Select access file to import"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
   txtFileName.Text = dlgAdd.FileName
   m_HasModify = True
End Sub

Private Sub cmdFileName2_Click()
On Error Resume Next
Dim strDescription As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Access Files (*.XLS)|*..xls;*.XLS;"
   dlgAdd.DialogTitle = "Select access file to import"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
   txtFileName2.Text = dlgAdd.FileName
   m_HasModify = True
End Sub

Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub
Private Sub cmdStart_Click()
Dim TempID As Long
   If Not VerifyTextControl(lblFileName, txtFileName) Then
      Exit Sub
   End If
   
   If InventoryActArea = 2 Then
      If Not VerifyTextControl(lblFileName2, txtFileName2) Then
         Exit Sub
      End If
   End If
   
   If InventoryActArea = 1 Or InventoryActArea = 2 Or InventoryActArea = 3 Then
      If Not VerifyDate(lblInventoryActDate, uctlInventoryActDate, False) Then
         Exit Sub
      End If
      
      If Not CheckUniqueNs(INVENTORY_ACT_UNIQUE, Trim(DateToStringInt(uctlInventoryActDate.ShowDate)), ID, Trim(str(InventoryActArea))) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & uctlInventoryActDate.ShowDate & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
         Exit Sub
      End If
   End If
   Call EnableForm(Me, False)
   Call LoadPartItem(Nothing, PartColls, , , , 2)
   Call LoadPartItem(Nothing, PartLabColls, , , , 4)
   If InventoryActArea = 2 Then
      Call ImportInventoryActDrug
   Else
      Call ImportInventoryAct
   End If
   Call EnableForm(Me, True)
End Sub

Private Sub ImportInventoryAct()
On Error GoTo ErrorHandler
Dim MaxRow As Long
Dim MaxCol As Long
Dim ID As Long
Dim i As Long
Dim j As Long
Dim row As Long
Dim Col As Long
Dim ErrorCount As Long
Dim SuccessCount As Long
Dim ProgressCount As Long
Dim HasBegin As Boolean
Dim Ia As CInventoryAct
Dim Iai As CInventoryActItem
Dim IsOK As Boolean
Dim SearchItemNo As CPartItem
Dim StatusInsert As Boolean
Dim CodeForInsert As String
Dim SearchItemNo2 As CInventoryActItem

   m_ExcelApp.Workbooks.Close
   m_ExcelApp.Workbooks.Open (txtFileName.Text)
   
   StatusInsert = True
   HasBegin = False

   ID = 1
   
   Set m_ExcelSheet = m_ExcelApp.Sheets(ID)
   
   MaxRow = m_ExcelSheet.UsedRange.Rows.Count
   MaxCol = m_ExcelSheet.UsedRange.Columns.Count
   
   Call EnableForm(Me, False)
   cmdStart.Enabled = False
   cmdExit.Enabled = False
   cmdOK.Enabled = False
       
   ProgressCount = 0
   ErrorCount = 0
   SuccessCount = 0
   
   prgProgress.MIN = 1
   prgProgress.MAX = (MaxRow * 3) + 1
   
   Set Ia = New CInventoryAct
   Ia.AddEditMode = SHOW_ADD
   If InventoryActArea = 1 Or InventoryActArea = 2 Or InventoryActArea = 3 Then
      Ia.INVENTORY_ACT_DATE = uctlInventoryActDate.ShowDate
   Else
      Ia.INVENTORY_ACT_DATE = Now
   End If
   Ia.INVENTORY_ACT_AREA = InventoryActArea
   Ia.INVENTORY_ACT_DESC = "IMPORTED " & DateToStringExtEx3(Now)
   
   For row = 2 To MaxRow - 1
      DoEvents
      Me.Refresh
      For j = row + 1 To MaxRow 'ตรวจสอบบรรทัดที่เบอร์ซ้ำกัน
         If Trim(m_ExcelSheet.Cells(row, 1).Value) = Trim(m_ExcelSheet.Cells(j, 1).Value) Then
              MsgBox "เอกสาร : <" & txtFileName2.Text & ">  เบอร์วัตถุดิบ : " & m_ExcelSheet.Cells(row, 1).Value & " บรรทัดที่ " & row & " ซ้ำกับบรรทัดที่ " & j & " กรุณาแก้ไขไม่ให้ซ้ำกัน"
               Call EnableForm(Me, True)
               cmdExit.Enabled = True
               cmdOK.Enabled = True
               glbDatabaseMngr.DBConnection.RollbackTrans
               Exit Sub
         End If
      Next j
   Next row
   'รอบแรก วัตถุดิบ หน่วยตัน ต้องคูณ 1000 แต่ถ้าเป็น ยา ไม่ต้องคูณ
   prgProgress.MAX = 100
   For row = 2 To MaxRow
      DoEvents
      Me.Refresh
      
      If Val(m_ExcelSheet.Cells(row, 1).Value) > 0 Then
         Set Iai = New CInventoryActItem
         Iai.Flag = "A"
         
         If InventoryActArea = 2 Then 'ถ้าเป็น ยา ไม่ต้องคูณ 1000 ให้เป็น kg เลย
            Iai.INVENTORY_ACT_AMOUNT = Val(m_ExcelSheet.Cells(row, 3).Value)
         Else
            Iai.INVENTORY_ACT_AMOUNT = Val(m_ExcelSheet.Cells(row, 3).Value) * 1000
         End If
         If Not SearchLabCode(SearchItemNo, Trim(m_ExcelSheet.Cells(row, 1).Value), Trim(m_ExcelSheet.Cells(row, 2).Value)) Then
             StatusInsert = False
             CodeForInsert = CodeForInsert & Trim(m_ExcelSheet.Cells(row, 1).Value) & " : " & Trim(m_ExcelSheet.Cells(row, 2).Value) & ","
         End If
         
         Iai.PART_ITEM_ID = SearchItemNo.PART_ITEM_ID
         If InventoryActArea = 1 Then
            Call Ia.CollRawMaterials.add(Iai)
         ElseIf InventoryActArea = 2 Then
            Call Ia.CollPhamacyRoom.add(Iai)
         ElseIf InventoryActArea = 3 Then
            Call Ia.CollSilo.add(Iai)
         End If
         Set Iai = Nothing
      End If
      
      ProgressCount = ProgressCount + 1
      'prgProgress.Value = ProgressCount
      prgProgress.Value = MyDiff(ProgressCount, MaxRow) * 100
      txtPercent.Text = prgProgress.Value
   Next row
   
   glbDatabaseMngr.DBConnection.BeginTrans
   HasBegin = True
   
   If StatusInsert = True Then
      Call glbInventoryAct.AddEditInventoryAct(Ia, IsOK, False, glbErrorLog)
      Call EnableForm(Me, True)
      glbDatabaseMngr.DBConnection.CommitTrans
      HasBegin = False
   Else
      lblNote.Caption = CodeForInsert
      Call EnableForm(Me, True)
      cmdExit.Enabled = True
      cmdOK.Enabled = True
      glbDatabaseMngr.DBConnection.RollbackTrans
      Exit Sub
   End If
   
'   For Each SearchItemNo In PartLabUpdateColls
'      Call SearchItemNo.UpdateLabPartNo
'   Next SearchItemNo
   
   Set Ia = Nothing
   prgProgress.Value = prgProgress.MAX
   

   
   Set m_ExcelSheet = Nothing
   
   'cmdStart.Enabled = True
   cmdExit.Enabled = True
   cmdOK.Enabled = True
   m_ExcelApp.Workbooks.Close
   
   Exit Sub
   
ErrorHandler:
   If HasBegin Then
      glbDatabaseMngr.DBConnection.RollbackTrans
   End If
   m_ExcelApp.Workbooks.Close
   Call EnableForm(Me, True)
   cmdStart.Enabled = True
   cmdExit.Enabled = True
   cmdOK.Enabled = True
   
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
End Sub
Private Sub ImportInventoryActDrug()
On Error GoTo ErrorHandler
Dim MaxRow As Long
Dim MaxCol As Long
Dim ID As Long
Dim i As Long
Dim j As Long
Dim row As Long
Dim Col As Long
Dim ErrorCount As Long
Dim SuccessCount As Long
Dim ProgressCount As Long
Dim HasBegin As Boolean
Dim Ia As CInventoryAct
Dim Iai As CInventoryActItem
Dim IsOK As Boolean
Dim SearchItemNo As CPartItem
Dim StatusInsert As Boolean
Dim CodeForInsert As String
Dim SearchItemNo2 As CInventoryActItem
Dim CollPrue As Collection
Dim CollOC As Collection
   Set CollPrue = New Collection
   Set CollOC = New Collection
   StatusInsert = True
   HasBegin = False
   Call EnableForm(Me, False)
   cmdStart.Enabled = False
   cmdExit.Enabled = False
   cmdOK.Enabled = False
       
   ProgressCount = 0
   ErrorCount = 0
   SuccessCount = 0

   
   Set Ia = New CInventoryAct
   Ia.AddEditMode = SHOW_ADD
   If InventoryActArea = 2 Then
      Ia.INVENTORY_ACT_DATE = uctlInventoryActDate.ShowDate
   Else
      Ia.INVENTORY_ACT_DATE = Now
   End If
   Ia.INVENTORY_ACT_AREA = InventoryActArea
   Ia.INVENTORY_ACT_DESC = "IMPORTED " & DateToStringExtEx3(Now)
   
   'เอกสารที่ 1 เพรียว
   m_ExcelApp.Workbooks.Close
   m_ExcelApp.Workbooks.Open (txtFileName2.Text)
   ID = 1
   Set m_ExcelSheet = m_ExcelApp.Sheets(ID)
   MaxRow = m_ExcelSheet.UsedRange.Rows.Count
   MaxCol = m_ExcelSheet.UsedRange.Columns.Count
'   prgProgress.MIN = 1
''   prgProgress.MAX = (MaxRow * 3) + 1
'   prgProgress.MAX = 100
   
   For row = 5 To MaxRow
      DoEvents
      Me.Refresh
      
      If Val(m_ExcelSheet.Cells(row, 1).Value) > 0 Then
         Set Iai = New CInventoryActItem
         Iai.Flag = "A"
         Iai.INVENTORY_ACT_AMOUNT = Val(m_ExcelSheet.Cells(row, "M").Value)
         If Not SearchLabCode(SearchItemNo, Trim(m_ExcelSheet.Cells(row, 1).Value), Trim(m_ExcelSheet.Cells(row, 3).Value)) Then
             'StatusInsert = False
             CodeForInsert = CodeForInsert & Trim(m_ExcelSheet.Cells(row, 1).Value) & " : " & Trim(m_ExcelSheet.Cells(row, 3).Value) & ","
             SearchItemNo.PART_ITEM_ID = -1
         End If
         If SearchItemNo.PART_ITEM_ID <> -1 Then
            Iai.PART_ITEM_ID = SearchItemNo.PART_ITEM_ID
            Set SearchItemNo2 = GetObject("CInventoryActItem", CollPrue, Trim(str(Iai.PART_ITEM_ID)), False)
            If Not SearchItemNo2 Is Nothing Then 'ถ้าเจอ
               SearchItemNo2.INVENTORY_ACT_AMOUNT = SearchItemNo2.INVENTORY_ACT_AMOUNT + Iai.INVENTORY_ACT_AMOUNT
            Else
               Call CollPrue.add(Iai, Trim(str(Iai.PART_ITEM_ID)))
            End If
         End If
         Set Iai = Nothing
      End If
      
'      ProgressCount = ProgressCount + row
'      prgProgress.Value = ProgressCount
'      txtPercent.Text = prgProgress.Value
   Next row
   
   'เอกสารที่ 2 เศษเพรียว
   m_ExcelApp.Workbooks.Close
   m_ExcelApp.Workbooks.Open (txtFileName.Text)
   ID = 1
   Set m_ExcelSheet = m_ExcelApp.Sheets(ID)
   MaxRow = m_ExcelSheet.UsedRange.Rows.Count
   MaxCol = m_ExcelSheet.UsedRange.Columns.Count
   prgProgress.MAX = 100
   For row = 5 To MaxRow
      DoEvents
      Me.Refresh
      
      If Val(m_ExcelSheet.Cells(row, 1).Value) > 0 Then
         Set Iai = New CInventoryActItem
         Iai.Flag = "A"
         
        
         Iai.INVENTORY_ACT_AMOUNT = Val(m_ExcelSheet.Cells(row, "EA").Value)
         If Not SearchLabCode(SearchItemNo, Trim(m_ExcelSheet.Cells(row, 1).Value), Trim(m_ExcelSheet.Cells(row, 2).Value)) Then
'             StatusInsert = False
             CodeForInsert = CodeForInsert & Trim(m_ExcelSheet.Cells(row, 1).Value) & " : " & Trim(m_ExcelSheet.Cells(row, 2).Value) & ","
             SearchItemNo.PART_ITEM_ID = -1
         End If
         If SearchItemNo.PART_ITEM_ID <> -1 Then
         Iai.PART_ITEM_ID = SearchItemNo.PART_ITEM_ID
         Set SearchItemNo2 = GetObject("CInventoryActItem", CollPrue, Trim(str(SearchItemNo.PART_ITEM_ID)), False)
            If Not SearchItemNo2 Is Nothing Then 'ถ้าเจอ
            SearchItemNo2.INVENTORY_ACT_AMOUNT = SearchItemNo2.INVENTORY_ACT_AMOUNT + Iai.INVENTORY_ACT_AMOUNT
            Else
            Call CollPrue.add(Iai, Trim(str(Iai.PART_ITEM_ID)))
            End If
         End If
      Set Iai = Nothing
      End If
      
'      ProgressCount = ProgressCount + 1
'      prgProgress.Value = MyDiff(ProgressCount, MaxRow) * 100
'      txtPercent.Text = prgProgress.Value
   Next row
   
   For Each Iai In CollPrue
      Call Ia.CollPhamacyRoom.add(Iai)
   Next Iai
   
   glbDatabaseMngr.DBConnection.BeginTrans
   HasBegin = True
   
   If StatusInsert = True Then
      Call glbInventoryAct.AddEditInventoryAct(Ia, IsOK, False, glbErrorLog)
      Call EnableForm(Me, True)
      glbDatabaseMngr.DBConnection.CommitTrans
      HasBegin = False
   Else
      lblNote.Caption = CodeForInsert
      Call EnableForm(Me, True)
      cmdExit.Enabled = True
      cmdOK.Enabled = True
      glbDatabaseMngr.DBConnection.RollbackTrans
      Exit Sub
   End If
   
   Set Ia = Nothing
   prgProgress.Value = prgProgress.MAX
   Set m_ExcelSheet = Nothing
   cmdExit.Enabled = True
   cmdOK.Enabled = True
   m_ExcelApp.Workbooks.Close
   Exit Sub
   
ErrorHandler:
   If HasBegin Then
      glbDatabaseMngr.DBConnection.RollbackTrans
   End If
   
   m_ExcelApp.Workbooks.Close
   Call EnableForm(Me, True)
   cmdStart.Enabled = True
   cmdExit.Enabled = True
   cmdOK.Enabled = True
   
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
End Sub
Private Sub Form_Activate()
Dim FromDate As Date
Dim ToDate As Date

   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      If InventoryActArea = 1 Or InventoryActArea = 2 Or InventoryActArea = 3 Then
         uctlInventoryActDate.SetFocus
         uctlInventoryActDate.ShowDate = Now
      Else
         uctlInventoryActDate.Enable = False
         uctlInventoryActDate.TabStop = False
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
'   ElseIf Shift = 0 And KeyCode = 117 Then
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

Private Sub ResetStatus()
   prgProgress.MAX = 100
   prgProgress.MIN = 0
   prgProgress.Value = 0
   txtPercent.Text = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = "อิมพอร์ต" & HeaderText
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitNormalLabel(lblInventoryActDate, MapText("วันที่นับสต๊อก"))
   Dim str As String

   If InventoryActArea = 2 Then
      lblFileName2.Visible = True
      txtFileName2.Visible = True
      cmdFileName2.Visible = True
      Call InitNormalLabel(lblFileName2, "ไฟล์ยาเพรียว")
      Call InitNormalLabel(lblFileName, "ไฟล์เศษยาเพรียว")
      str = "- ไฟล์ยาเพรียว : เริ่ม Import ที่ Row ที่ 5 ,Column A ของ Sheet1 เท่านั้น โดย" & vbCrLf & "Col A = รหัสวัตถุดิบ,Col C =รายละเอียดวัตถุดิบ,Col M =น้ำหนักคงเหลือ (กก.)" & vbCrLf & " ***โดยห้ามเพิ่มหรือลบ Column และให้ Sheet ที่ต้องการต้องเป็น Sheet แรกเท่านั้น***"
      str = str & vbCrLf & "- ไฟล์เศษยาเพรียว : เริ่ม Import ที่ Row ที่ 5 ,Column A ของ Sheet1 เท่านั้น โดย" & vbCrLf & "Col A = รหัสวัตถุดิบ,Col EA =น้ำหนักคงเหลือ/นับจริง (กก.)" & vbCrLf & " ***โดยห้ามเพิ่มหรือลบ Column และให้ Sheet ที่ต้องการต้องเป็น Sheet แรกเท่านั้น***"
   Else
      Call InitNormalLabel(lblFileName, "ชื่อไฟล์")
       str = "- เริ่ม Import ที่ Row ที่ 2 ,Column A ของ Sheet1 เท่านั้น โดย" & vbCrLf & "Col A = รหัสวัตถุดิบ,Col B =รายละเอียดวัตถุดิบ,Col C =จำนวนเป็นตัน/กิโลกรัม" & vbCrLf & " ***โดยห้ามให้เบอร์วัตถุดิบซ้ำกัน***"
   End If
   
   Call InitNormalLabel(lblNote, str)
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(Label1, "%")

   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPercent.Enabled = False
   Call txtFileName.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtFileName.Enabled = False
   Call txtFileName2.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtFileName2.Enabled = False
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdStart.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdFileName.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdFileName2.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdStart, MapText("เริ่ม"))
   Call InitMainButton(cmdFileName, MapText("..."))
   Call InitMainButton(cmdFileName2, MapText("..."))
   
   Call ResetStatus
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   Call EnableForm(Me, False)
   m_HasActivate = False
   
   Set PartUctlColls = New Collection
   Set PartColls = New Collection
   Set PartLabColls = New Collection
   Set PartLabUpdateColls = New Collection
   
   Set m_ExcelApp = CreateObject("Excel.application")
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set PartColls = Nothing
   Set PartLabColls = Nothing
   Set PartLabUpdateColls = Nothing
   Set PartUctlColls = Nothing
End Sub
Private Function SearchLabCode(SearchItemNo As CPartItem, PartNo As String, PartName As String) As Boolean
   SearchLabCode = True
   Set SearchItemNo = GetObject("CPartItem", PartColls, PartNo, False)
   If SearchItemNo Is Nothing Then
      Set SearchItemNo = GetObject("CPartItem", PartLabColls, PartNo, False)
      If SearchItemNo Is Nothing Then
         Set SearchItemNo = GetObject("CPartItem", PartLabUpdateColls, Trim(PartNo), False)
         If SearchItemNo Is Nothing Then
            'LoadForm
            Set SearchItemNo = New CPartItem
            Set frmMapPlcProductItem.PartItem = SearchItemNo
'            Set frmMapPlcProductItem.mPartItemColl = PartUctlColls
            If Trim(PartNo) = Trim(PartName) Then
               MsgBox MapText("MAP ข้อมูล รหัสสินค้า/วัตถุดิบ " & PartNo)
            Else
               MsgBox MapText("MAP ข้อมูล รหัสสินค้า/วัตถุดิบ " & PartNo & "-" & PartName & "กรุณาติดต่อบัญชีให้เพิ่มรหัสข้อมูลเข้าระบบ")
            End If
'            frmMapPlcProductItem.ShowMode = SHOW_ADD
'            Load frmMapPlcProductItem
'            frmMapPlcProductItem.Show 1

'            OKClick = frmMapPlcProductItem.OKClick
'
'            Unload frmMapPlcProductItem
'            Set frmMapPlcProductItem = Nothing

            'AddDataTo PartPlcUpdateColls
            If Len(Trim(SearchItemNo.PART_NO)) <= 0 Then
               glbErrorLog.LocalErrorMsg = "ไม่พบรหัสที่อ้างอิง วัตถุดิบ สำหรับ " & PartNo & "-" & PartName
               glbErrorLog.ShowUserError

               SearchLabCode = False
               Exit Function
            End If
            SearchItemNo.NUMBER_LAB_ID = Trim(PartNo)
'            Call PartLabUpdateColls.add(SearchItemNo, Trim(PartNo))
         End If
      End If
   End If
End Function

