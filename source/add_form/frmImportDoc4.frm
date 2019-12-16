VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportDoc4 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Icon            =   "frmImportDoc4.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   10410
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3765
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   6641
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextBox txtFileName 
         Height          =   435
         Left            =   1860
         TabIndex        =   11
         Top             =   1470
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   767
      End
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   1860
         TabIndex        =   0
         Top             =   1920
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   6
         Top             =   0
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtPercent 
         Height          =   465
         Left            =   1860
         TabIndex        =   1
         Top             =   2250
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
      Begin Threed.SSCheck SSCheck1 
         Height          =   255
         Left            =   1860
         TabIndex        =   13
         Top             =   960
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   450
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin Threed.SSCommand cmdFileName 
         Height          =   405
         Left            =   8670
         TabIndex        =   12
         Top             =   1470
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportDoc4.frx":27A2
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   1890
         TabIndex        =   2
         Top             =   2910
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportDoc4.frx":2ABC
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3600
         TabIndex        =   10
         Top             =   2370
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   9
         Top             =   1980
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   8
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label lblMasterName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   7
         Top             =   1500
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   8535
         TabIndex        =   4
         Top             =   2910
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   6885
         TabIndex        =   3
         Top             =   2910
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportDoc4.frx":2DD6
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmImportDoc4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_PartItem As CPartItem
Private m_PartItemSelect As CPartItemSelect

Public id As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String

Private m_PartGroups As Collection
Private m_FeatureTypeDateLocations As Collection
Private m_ExcelApp As Object
Private m_ExcelSheet As Object
Private isSave As Boolean
Private PartColls As Collection
Private m_PartItems As Collection

Private Sub cboPartType_Click()
   m_HasModify = True
End Sub

Private Sub cboPosition_Click()
   m_HasModify = True
End Sub

Private Sub chkPigFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkBalanceFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

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

Private Sub cmdOK_Click()
   If isSave Then
      glbDatabaseMngr.DBConnection.CommitTrans
      OKClick = True
   End If
   Unload Me
End Sub

Private Sub QueryData(Flag As Boolean)
End Sub

Private Sub cmdStart_Click()
Dim TempID As Long

   If Not VerifyTextControl(lblMasterName, txtFileName) Then
      Exit Sub
   End If
   
      Call EnableForm(Me, False)
      
      'สินค้าบริการ
      Call LoadPartItem(Nothing, PartColls, , , , 2)
      Call ImportPartItemSelect
      
      m_ExcelApp.Workbooks.Close
      Call EnableForm(Me, True)

End Sub

Private Sub ImportPartItemSelect()
On Error GoTo ErrorHandler
Dim MaxRow As Long
Dim MaxCol As Long
Dim id As Long
Dim I As Long
Dim J As Long
Dim row As Long
Dim Col As Long
Dim ErrorCount As Long
Dim SuccessCount As Long
Dim ProgressCount As Long
Dim HasBegin As Boolean
Dim PIS As CPartItemSelect
Dim PISI As CPartItemSelect
Dim IsOK As Boolean
Dim SearchItemNo As CPartItem
Dim StatusInsert As Boolean
Dim CodeForInsert As String
Dim Beginrow As Long

   m_ExcelApp.Workbooks.Close
   m_ExcelApp.Workbooks.Open (txtFileName.Text)
   
   StatusInsert = True
   HasBegin = False

   id = 1
   Beginrow = 2
   
   Set m_ExcelSheet = m_ExcelApp.Sheets(id)
   
   MaxRow = m_ExcelSheet.UsedRange.Rows.Count
   MaxCol = m_ExcelSheet.UsedRange.Columns.Count
   
   Call EnableForm(Me, False)
   cmdStart.Enabled = False
   cmdExit.Enabled = False
   cmdOK.Enabled = False
       
   ProgressCount = 0
   ErrorCount = 0
   SuccessCount = 0
   
   prgProgress.MIN = 0
   prgProgress.MAX = (MaxRow * 3) + 1

   
   For row = Beginrow To MaxRow - 1
      DoEvents
      Me.Refresh
      If Len(m_ExcelSheet.Cells(row, 1).Value) = 0 Then
         MsgBox "เอกสาร : <" & txtFileName.Text & ">  บรรทัดที่ " & row & " ต้องไม่เป็นช่องว่างและต้องมีค่ามากกว่า 0 กรุณาตรวจสอบ"
         Call EnableForm(Me, True)
         cmdExit.Enabled = True
         cmdOK.Enabled = True
         Exit Sub
      End If
      If SSCheck1.Value = ssCBChecked Then
      For J = row + 1 To MaxRow 'ตรวจสอบบรรทัดที่เบอร์ซ้ำกัน
         If Trim(m_ExcelSheet.Cells(row, 1).Value) = Trim(m_ExcelSheet.Cells(J, 1).Value) And Len(m_ExcelSheet.Cells(row, 1).Value) > 0 Then
              MsgBox "เอกสาร : <" & txtFileName.Text & ">  เบอร์วัตถุดิบ : " & m_ExcelSheet.Cells(row, 1).Value & " บรรทัดที่ " & row & " ซ้ำกับบรรทัดที่ " & J & " กรุณาแก้ไขไม่ให้ซ้ำกัน"
               Call EnableForm(Me, True)
               cmdExit.Enabled = True
               cmdOK.Enabled = True
               Exit Sub
         End If
      Next J
      End If
   Next row


   Set PIS = New CPartItemSelect
   PIS.AddEditMode = SHOW_ADD
   prgProgress.MAX = 100
   For row = Beginrow To MaxRow
      DoEvents
      Me.Refresh

      If Len(m_ExcelSheet.Cells(row, 1).Value) > 0 Then
         Set PISI = New CPartItemSelect
         PISI.AddEditMode = SHOW_ADD
         PISI.Flag = "A"
         
         PISI.PART_ITEM_SELECT_NO = "001"
       
       Set SearchItemNo = GetObject("CPartItem", PartColls, Trim(m_ExcelSheet.Cells(row, 1).Value), False)
         If Not SearchItemNo Is Nothing Then
            PISI.PART_ITEM_ID = SearchItemNo.PART_ITEM_ID
         End If
         
        Call PIS.Part_Sel_Coll.add(PISI)
         Set PISI = Nothing
      End If
      
      ProgressCount = ProgressCount + 1
      prgProgress.Value = MyDiff(ProgressCount, MaxRow) * 100
      txtPercent.Text = prgProgress.Value
   Next row
   
   glbDatabaseMngr.DBConnection.BeginTrans
   HasBegin = True
   
   If StatusInsert = True Then
       Call PIS.DeleteData
       
      If Not glbDaily.AddEditPartItemSelect(PIS, IsOK, False, glbErrorLog) Then
          Call EnableForm(Me, True)
          HasBegin = False
       End If
       If Not IsOK Then
          Call EnableForm(Me, True)
          glbErrorLog.ShowUserError
      Else
         isSave = True
       End If
   Else
      Call EnableForm(Me, True)
      cmdExit.Enabled = True
      cmdOK.Enabled = True
      glbDatabaseMngr.DBConnection.RollbackTrans
      Exit Sub
   End If
   
   Set PIS = Nothing
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
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         id = 0
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
   pnlHeader.Caption = "อิมพอร์ตข้อมูล"
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitNormalLabel(lblMasterName, "ชื่อไฟล์")
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(Label1, "%")
   
   Call InitCheckBox(SSCheck1, "ตรวจสอบรหัสสินค้าซ้ำก่อน Import")
   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPercent.Enabled = False
   Call txtFileName.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtFileName.Enabled = False
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdStart.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdFileName.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdStart, MapText("เริ่ม"))
   Call InitMainButton(cmdFileName, MapText("..."))
   
   Call ResetStatus
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   If isSave Then
      glbDatabaseMngr.DBConnection.RollbackTrans
   End If
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   Call EnableForm(Me, False)
   m_HasActivate = False
   Set m_Rs = New ADODB.Recordset
   Set m_PartItems = New Collection
   Set PartColls = New Collection
   
   Set m_ExcelApp = CreateObject("Excel.application")
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub txtLastName_Change()
   m_HasModify = True
End Sub

Private Sub txtCode_Change()
   m_HasModify = True
End Sub

Private Sub txtName_Change()
   m_HasModify = True
End Sub

Private Sub txtMaxUser_Change()
   m_HasModify = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Not OKClick Then
      Call cmdExit_Click
   End If
   Set PartColls = Nothing
   Set m_PartItems = Nothing
End Sub

