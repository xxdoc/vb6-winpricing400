VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportPostItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Icon            =   "frmImportPostItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   10410
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3405
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   6006
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboExportType 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   840
         Width           =   3135
      End
      Begin prjFarmManagement.uctlTextBox txtFileName 
         Height          =   435
         Left            =   1860
         TabIndex        =   11
         Top             =   1350
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   767
      End
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   1860
         TabIndex        =   0
         Top             =   1800
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
         Top             =   2130
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
      Begin VB.Label lblExportType 
         Alignment       =   1  'Right Justify
         Caption         =   "lblOrderBy"
         Height          =   315
         Left            =   600
         TabIndex        =   14
         Top             =   900
         Width           =   1125
      End
      Begin Threed.SSCommand cmdFileName 
         Height          =   405
         Left            =   8670
         TabIndex        =   12
         Top             =   1350
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportPostItem.frx":27A2
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   1890
         TabIndex        =   2
         Top             =   2670
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportPostItem.frx":2ABC
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3600
         TabIndex        =   10
         Top             =   2250
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   9
         Top             =   1860
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   8
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label lblMasterName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   7
         Top             =   1380
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   8535
         TabIndex        =   4
         Top             =   2670
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
         Top             =   2670
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportPostItem.frx":2DD6
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmImportPostItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Public id As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String

Private CountBill As Long
Private CountDown As Double

Private Cd As CCashDoc

Private DocumentType As Long

Private Sub cmdFileName_Click()
On Error Resume Next
Dim strDescription As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Text Files (*.TXT)|*..txt;*.TXT;"
   dlgAdd.DialogTitle = "Select access file to import"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
   txtFileName.Text = dlgAdd.FileName
   m_HasModify = True
End Sub
Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub
Private Sub cmdStart_Click()
Dim TempID As Long

   If Not VerifyTextControl(lblMasterName, txtFileName) Then
      Exit Sub
   End If
   
   If Not VerifyCombo(lblExportType, cboExportType, False) Then
      Exit Sub
   End If
         
   Call EnableForm(Me, False)
   
   TempID = cboExportType.ItemData(Minus2Zero(cboExportType.ListIndex))
   
   If TempID = 1 Then
      Call ImportSupItem
   End If
   
   Call EnableForm(Me, True)
   
End Sub

Private Sub ImportSupItem()
On Error GoTo ErrorHandler
Dim HasBegin As Boolean
Dim FileName As String
Dim F As Long
Dim TempStr As String
Dim LineCount As Long
Dim SuccessCount As Long
Dim OutRance As Long
Dim ErrorCount As Long
Dim Sum As Long
Dim I As Long

Dim FromDate As Date
Dim ToDate As Date
Dim FileName1 As String
   
   Call glbDatabaseMngr.DBConnection.BeginTrans
   
   CountBill = 0
   
   FileName1 = Dir(txtFileName.Text)
   
   FromDate = DateSerial(Mid(FileName1, 10, 4), Mid(FileName1, 14, 2), Mid(FileName1, 16, 2))
   ToDate = DateSerial(Mid(FileName1, 19, 4), Mid(FileName1, 23, 2), Mid(FileName1, 25, 2))
   'Call LoadCashDocPostDistinctDocumentNo(Nothing, c_DocumentNos, FromDate, ToDate)
   
   prgProgress.MAX = 100
   prgProgress.MIN = 0
   FileName = txtFileName.Text
   F = FreeFile()
   Close #F
   Open FileName For Input As #F
   Sum = 0
   While Not EOF(F)
      Line Input #F, TempStr
      Sum = Sum + 1
   Wend
   
   HasBegin = True
      
   CountDown = Sum
   
   SuccessCount = 0
   ErrorCount = 0
   F = FreeFile()
   Close #F
   Open FileName For Input As #F
   While Not EOF(F)
      I = I + 1
      Line Input #F, TempStr
      prgProgress.Value = MyDiff(I, Sum) * 100
      txtPercent.Text = prgProgress.Value
      LineCount = LineCount + 1
      Me.Refresh
      DoEvents
      
      CountDown = CountDown - 1
      
      If CountDown = 19 Then
         'Debug.Print
      End If
      
      If ProcessLine(TempStr) Then
         SuccessCount = SuccessCount + 1
      Else
         ErrorCount = ErrorCount + 1
      End If
   Wend
   Close #F
   prgProgress.Value = 100
   txtPercent.Text = 100
   
   glbErrorLog.LocalErrorMsg = "อิมพอร์ต สำเสร็จจำนวน " & SuccessCount & "และ ล้มเหลวจำนวน " & ErrorCount & " ข้อมูล"
   glbErrorLog.ShowUserError
   
   If ConfirmSave Then
      glbDatabaseMngr.DBConnection.CommitTrans
   Else
      glbDatabaseMngr.DBConnection.RollbackTrans
   End If
   Exit Sub
   
ErrorHandler:
   If HasBegin Then
      glbDatabaseMngr.DBConnection.RollbackTrans
   End If
   
   Call EnableForm(Me, True)
   
   cmdStart.Enabled = True
   cmdExit.Enabled = True
   cmdOK.Enabled = True
   
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
End Sub
Private Function ProcessLine(LineStr As String) As Boolean
On Error GoTo ErrorHandler
'Dim TimeStamp As Date
'Dim TimeStampStr As String
Dim TempAsc As Long
Dim OldTempAsc As Long
'Dim FirstDate As Date
'Dim LastDate As Date
Dim I As Long
Dim IsOK As Boolean
Dim Si As CCashDocPost
'Dim Key1 As String
'Dim Key2 As String
'Dim Key3 As String
'Dim Key4 As String
Dim BD As CBillingDoc

Dim m_Rs  As ADODB.Recordset
Dim ItemCount  As Long
   
   If Left(LineStr, 2) = "CD" Then
      If CountBill > 0 Then
         Call CashDocPost2BillingDoc(Cd, BD, 15000)        ' ใบสร้างจากใบเช็ครอจ่าย
         
         Call glbDaily.AddEditCashDoc(Cd, IsOK, False, glbErrorLog)
         
         Set Cd = New CCashDoc
         
      End If
      CountBill = 1
      
      TempAsc = 3
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      
      
      Set m_Rs = New ADODB.Recordset
      
      Cd.QueryFlag = 1
      Call Cd.SetFieldValue("DOCUMENT_NO", Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1))
      
      If Cd.GetFieldValue("DOCUMENT_NO") = "15062550" Then
         'Debug.Print
      End If
      If Not glbDaily.QueryCashDocEx(Cd, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Function
      End If
      
      If ItemCount > 0 Then
         Cd.ShowMode = SHOW_EDIT
         Call Cd.PopulateFromRS(1, m_Rs)
      Else
         Cd.ShowMode = SHOW_ADD
         
         Call Cd.SetFieldValue("DOCUMENT_NO", Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1))
         OldTempAsc = TempAsc
         
         Call Cd.SetFieldValue("DOCUMENT_DATE", InternalDateToDate(StingToVariable(TempAsc, OldTempAsc, LineStr)))
         Call Cd.SetFieldValue("BANK_ACCOUNT", StingToVariable(TempAsc, OldTempAsc, LineStr))
         Call Cd.SetFieldValue("DOCUMENT_TYPE", StingToVariable(TempAsc, OldTempAsc, LineStr))
         
         Call Cd.SetFieldValue("ENTERPRISE_ID", StingToVariable(TempAsc, OldTempAsc, LineStr))
         Call Cd.SetFieldValue("EMP_ID", StingToVariable(TempAsc, OldTempAsc, LineStr))
         Call Cd.SetFieldValue("CUSTOMER_ID", StingToVariable(TempAsc, OldTempAsc, LineStr))
         Call Cd.SetFieldValue("SUPPLIER_ID", StingToVariable(TempAsc, OldTempAsc, LineStr))
         
      End If

   End If

   Dim TempCheque As String
   
   If Left(LineStr, 2) = "CP" Then
      Dim CP As CCashDocPost
      Dim TempCp As CCashDocPost
            
      TempAsc = 3
      OldTempAsc = TempAsc
      
      TempCheque = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      For Each TempCp In Cd.PostItems
         '''Debug.Print (TempCp.GetFieldValue("CHEQUE_NO"))
         If TempCheque = TempCp.GetFieldValue("CHEQUE_NO") Then
            ProcessLine = True
            Exit Function
         End If
      Next
         
      Set CP = New CCashDocPost
      CP.Flag = "A"
      
      Call CP.SetFieldValue("CHEQUE_ID", StingToVariable(TempAsc, OldTempAsc, LineStr))
      Call CP.SetFieldValue("BANK_BRANCH", StingToVariable(TempAsc, OldTempAsc, LineStr))
      Call CP.SetFieldValue("BANK_ID", StingToVariable(TempAsc, OldTempAsc, LineStr))
      Call CP.SetFieldValue("CHEQUE_AMOUNT", StingToVariable(TempAsc, OldTempAsc, LineStr))
      Call CP.SetFieldValue("POST_TYPE", StingToVariable(TempAsc, OldTempAsc, LineStr))
      Call CP.SetFieldValue("BILLING_DOC_ID", StingToVariable(TempAsc, OldTempAsc, LineStr))
      Call CP.SetFieldValue("WH_AMOUNT", StingToVariable(TempAsc, OldTempAsc, LineStr))
      Call CP.SetFieldValue("INTERREST_AMOUNT", StingToVariable(TempAsc, OldTempAsc, LineStr))
      Call CP.SetFieldValue("BILLING_DOC_NO", StingToVariable(TempAsc, OldTempAsc, LineStr))
      Call CP.SetFieldValue("CHEQUE_SUPPLIER_ID", StingToVariable(TempAsc, OldTempAsc, LineStr))
      
      If CP.GetFieldValue("BILLING_DOC_NO") = "5005249" Then
         'Debug.Print
      End If
      
      Call Cd.PostItems.add(CP)
      
   End If
   
   If CountDown = 0 Then
      Call CashDocPost2BillingDoc(Cd, BD, 15000)        ' ใบสร้างจากใบเช็ครอจ่าย
      Call glbDaily.AddEditCashDoc(Cd, IsOK, False, glbErrorLog)
   End If
   
   ProcessLine = True

   Exit Function
ErrorHandler:
   ProcessLine = False
End Function

Private Sub Form_Activate()
Dim FromDate As Date
Dim ToDate As Date

   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call InitExportPostType(cboExportType)
      
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
   Call InitNormalLabel(lblExportType, "ประเภท")
   Call InitNormalLabel(Label1, "%")

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
   
   Call InitCombo(cboExportType)
   
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
   
   Set m_Rs = New ADODB.Recordset
   
   Set Cd = New CCashDoc
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   
   Set Cd = Nothing
   
End Sub
Private Function StingToVariable(TempAsc As Long, OldTempAsc As Long, LineStr As String) As Variant
   TempAsc = InStr(TempAsc + 1, LineStr, ";")
   StingToVariable = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)
   OldTempAsc = TempAsc
End Function
Public Function CashDocPost2BillingDoc(Cd As CCashDoc, BD As CBillingDoc, IvdDocType As Long) As Boolean
Dim IsOK As Boolean
Dim CP As CCashDocPost
   
   For Each CP In Cd.PostItems
      If CP.Post2BD.Count > 0 Then
         If CP.Flag = "" Then
            CP.Flag = "E"
         End If
         Set BD = CP.Post2BD(1)
         BD.Flag = "E"
      Else
         If CP.Flag = "" Then
            CP.Flag = "E"
         End If
         Set BD = New CBillingDoc
         BD.Flag = "A"
         Call CP.Post2BD.add(BD)
      End If
      
      BD.DOCUMENT_NO = CP.GetFieldValue("BILLING_DOC_NO")         'หมายเลข PV NO
      BD.DOCUMENT_DATE = Cd.GetFieldValue("DOCUMENT_DATE")
      BD.SUPPLIER_ID = CP.GetFieldValue("CHEQUE_SUPPLIER_ID")
      BD.PAID_AMOUNT = Val(CP.GetFieldValue("CHEQUE_AMOUNT")) + Val(CP.GetFieldValue("WH_AMOUNT")) - Val(CP.GetFieldValue("INTERREST_AMOUNT"))
      BD.DOCUMENT_TYPE = IvdDocType
      BD.COMMIT_FLAG = "N"
      BD.EXCEPTION_FLAG = "N"
   Next CP
End Function


