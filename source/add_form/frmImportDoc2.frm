VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportDoc2 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Icon            =   "frmImportDoc2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
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
      Begin VB.ComboBox cboImportType 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1020
         Width           =   3105
      End
      Begin prjFarmManagement.uctlTextBox txtFileName 
         Height          =   435
         Left            =   1860
         TabIndex        =   12
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
      Begin Threed.SSCommand SSCommand1 
         Height          =   525
         Left            =   6270
         TabIndex        =   15
         Top             =   780
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportDoc2.frx":27A2
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdFileName 
         Height          =   405
         Left            =   8670
         TabIndex        =   13
         Top             =   1470
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportDoc2.frx":2ABC
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
         MouseIcon       =   "frmImportDoc2.frx":2DD6
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3600
         TabIndex        =   11
         Top             =   2370
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   10
         Top             =   1980
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   9
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label lblMasterName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   8
         Top             =   1500
         Width           =   1575
      End
      Begin VB.Label lblFileName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   7
         Top             =   1080
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
         MouseIcon       =   "frmImportDoc2.frx":30F0
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmImportDoc2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Employee As CEmployee

Public id As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String

Private m_PartGroups As Collection
Private m_FeatureTypeDateLocations As Collection
Private m_ExcelApp As Object
Private m_ExcelSheet As Object
Private m_FeatureType As Collection
Private m_BankAccounts As Collection
Private m_Customers As Collection
Private m_Accounts As Collection
Private m_AccountCodes As Collection
Private m_SheetBalances As Collection

Private m_Suppliers As Collection

Private Sub cmdPasswd_Click()

End Sub

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
'   If Not SaveData Then
'      Exit Sub
'   End If
   
'   Call LoadAccessRight(Nothing, glbAccessRight, glbUser.GROUP_ID)
   OKClick = True
   Unload Me
End Sub

Private Sub QueryData(Flag As Boolean)
End Sub

Private Sub cmdStart_Click()
Dim TempID As Long

   If Not VerifyCombo(lblFileName, cboImportType) Then
      Exit Sub
   End If
   If Not VerifyTextControl(lblMasterName, txtFileName) Then
      Exit Sub
   End If
   
   TempID = cboImportType.ItemData(Minus2Zero(cboImportType.ListIndex))
   If TempID <= 0 Then
      Exit Sub
   End If
   
   If TempID = 1 Then
      Call EnableForm(Me, False)
      m_ExcelApp.Workbooks.Close
      m_ExcelApp.Workbooks.Open (txtFileName.Text)
      
      'สินค้าบริการ
      Call ImportData(1)
      
      m_ExcelApp.Workbooks.Close
      Call EnableForm(Me, True)
   ElseIf TempID = 2 Then
      Call EnableForm(Me, False)
      m_ExcelApp.Workbooks.Close
      m_ExcelApp.Workbooks.Open (txtFileName.Text)
      
      Call ImportData(2)
      
      m_ExcelApp.Workbooks.Close
      Call EnableForm(Me, True)
   ElseIf TempID = 3 Then
      Call EnableForm(Me, False)
      m_ExcelApp.Workbooks.Close
      m_ExcelApp.Workbooks.Open (txtFileName.Text)
      
      Call ImportData(3)

      
      m_ExcelApp.Workbooks.Close
      Call EnableForm(Me, True)
   End If
End Sub

Private Sub ImportStock()
On Error GoTo ErrorHandler
Dim MaxRow As Long
Dim MaxCol As Long
Dim id As Long
Dim FieldNames() As String
Dim FieldTypes() As String
Dim I As Long
Dim TabField As String
Dim StateMent As String
Dim NewValue As String
Dim row As Long
Dim Col As Long
Dim ErrorCount As Long
Dim SuccessCount As Long
Dim ProgressCount As Long
Dim ErrorFlag As Boolean
Dim ServerDtm As String
Dim HasBegin As Boolean
Dim Ivd As CInventoryDoc
Dim II As CLotItem
Dim IsOK As Boolean
Dim Lc As CLocation
Dim Pi As CPartItem
   
   
   HasBegin = False

   id = 1
   
   Set m_ExcelSheet = m_ExcelApp.Sheets(id)
   
   MaxRow = m_ExcelSheet.UsedRange.Rows.Count
   MaxCol = m_ExcelSheet.UsedRange.Columns.Count

   ReDim FieldNames(MaxCol)
   ReDim FieldTypes(MaxCol)
   
   Call EnableForm(Me, False)
   cmdStart.Enabled = False
   cmdExit.Enabled = False
   cmdOK.Enabled = False
       
   ProgressCount = 0
   ErrorCount = 0
   SuccessCount = 0
   
   prgProgress.MIN = 1
   prgProgress.MAX = (MaxRow) + 1
   
   glbDatabaseMngr.DBConnection.BeginTrans
   HasBegin = True
   
   Set Ivd = New CInventoryDoc
   Ivd.AddEditMode = SHOW_ADD
   Ivd.DOCUMENT_NO = "สต็อคยกมา"
   Ivd.DOCUMENT_DATE = InternalDateToDateEx2(m_ExcelApp.Sheets(id).NAME)
   Ivd.COMMIT_FLAG = "N"
   Ivd.DOCUMENT_TYPE = 1
   Ivd.EXCEPTION_FLAG = "N"
   
   For row = 3 To MaxRow
      DoEvents
      Me.Refresh
      
      Set II = New CLotItem
      II.Flag = "A"
      II.TOTAL_INCLUDE_PRICE = Val(m_ExcelSheet.Cells(row, 6).Value)
      II.TOTAL_ACTUAL_PRICE = II.TOTAL_INCLUDE_PRICE
      II.CALCULATE_FLAG = "Y"
      II.TX_AMOUNT = Val(m_ExcelSheet.Cells(row, 4).Value)
      II.INCLUDE_UNIT_PRICE = MyDiff(II.TOTAL_INCLUDE_PRICE, II.TX_AMOUNT)
      II.ACTUAL_UNIT_PRICE = II.INCLUDE_UNIT_PRICE
      II.CALCULATE_TYPE = 3
      II.TOTAL_WEIGHT = II.TX_AMOUNT
      Set Lc = GetLocation(m_FeatureType, Trim(m_ExcelSheet.Cells(row, 1).Value))
      II.LOCATION_ID = Lc.LOCATION_ID
      Set Pi = GetPartItem(m_FeatureType, Trim("S-" & m_ExcelSheet.Cells(row, 2).Value))
      II.PART_ITEM_ID = Pi.PART_ITEM_ID
      II.TX_TYPE = "I"
      If (II.PART_ITEM_ID > 0) And (II.LOCATION_ID > 0) Then
         Call Ivd.ImportExports.add(II)
      Else
         glbErrorLog.LocalErrorMsg = "ไม่พบข้อมูลวัตถุดิบ '" & Trim(m_ExcelSheet.Cells(row, 3).Value) & "' รหัส '" & Trim(m_ExcelSheet.Cells(row, 2).Value) & "' ในฐานข้อมูล"
         glbErrorLog.ShowUserError
      End If
      Set II = Nothing
      
      ProgressCount = ProgressCount + 1
      prgProgress.Value = ProgressCount
   Next row
   Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog)
   
   Set Ivd = Nothing
   prgProgress.Value = prgProgress.MAX
   
   Call EnableForm(Me, True)
   glbDatabaseMngr.DBConnection.CommitTrans
   HasBegin = False
   
   Set m_ExcelSheet = Nothing
   
   cmdStart.Enabled = True
   cmdExit.Enabled = True
   cmdOK.Enabled = True
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

Private Sub ImportData(Ind As Long)
On Error GoTo ErrorHandler
Dim MaxRow As Long
Dim MaxCol As Long
Dim id As Long
Dim I As Long
Dim TabField As String
Dim StateMent As String
Dim NewValue As String
Dim row As Long
Dim Col As Long
Dim ProgressCount As Long
Dim ErrorFlag As Boolean
Dim Cm As CCustomer
Dim Pg As CPartGroup
Dim Cai As CCustomerAccountList
Dim TmpCai As CCustomerAccountList
Dim TempCol As Collection
Dim Key1 As String
Dim Key2 As String
Dim FoundFlag As Boolean
Dim DrID As Long
Dim CrID As Long
Dim Acc1 As String
Dim Acc2 As String
Dim HasBegin As Boolean
Dim Cols As Collection
Dim Obj As Object
Dim TempCode As String
Dim TempID As Long
Dim TempName As String
Dim TempPg As CPartGroup
Dim TempFt As CFeatureType
Dim TempBa As CMasterRef

   HasBegin = False

   id = 1
   
   Set TempCol = New Collection
   Set m_ExcelSheet = m_ExcelApp.Sheets(Ind)
   
   MaxRow = m_ExcelSheet.UsedRange.Rows.Count
   MaxCol = m_ExcelSheet.UsedRange.Columns.Count

   Call EnableForm(Me, False)
   cmdStart.Enabled = False
   cmdExit.Enabled = False
   cmdOK.Enabled = False
       
   ProgressCount = 0
'   ErrorCount = 0
'   SuccessCount = 0
   
   prgProgress.MIN = 1
   prgProgress.MAX = m_Customers.Count + 1
         
   For row = 2 To MaxRow
      DoEvents
      Me.Refresh
      
      Set Cai = New CCustomerAccountList
      Call Cai.SetFieldValue("CUSTOMER_CODE", m_ExcelSheet.Cells(row, 1).Value)
      Call Cai.SetFieldValue("TEMP_CODE", m_ExcelSheet.Cells(row, 2).Value)
      Call Cai.SetFieldValue("DEBIT_NO", m_ExcelSheet.Cells(row, 3).Value)
      Call Cai.SetFieldValue("CREDIT_NO", m_ExcelSheet.Cells(row, 4).Value)
      
      If (Trim(m_ExcelSheet.Cells(row, 1).Value) <> "") Or _
         (Trim(m_ExcelSheet.Cells(row, 2).Value) <> "") Or _
         (Trim(m_ExcelSheet.Cells(row, 3).Value) <> "") Or _
         (Trim(m_ExcelSheet.Cells(row, 4).Value) <> "") Then
         Call TempCol.add(Cai, Cai.GetFieldValue("CUSTOMER_CODE") & "-" & Cai.GetFieldValue("TEMP_CODE"))
      End If
'''Debug.Print Cai.GetFieldValue("CUSTOMER_CODE") & "-" & Cai.GetFieldValue("TEMP_CODE")
      Set Cai = Nothing
   Next row
   
   glbDatabaseMngr.DBConnection.BeginTrans
   HasBegin = True
   
   If Ind = 1 Then
      Set Cols = m_PartGroups
      Set TempPg = New CPartGroup
      TempPg.PART_GROUP_ID = -1
      TempPg.PART_GROUP_NO = "*"
      Call Cols.add(TempPg)
      Set TempPg = Nothing
   ElseIf Ind = 2 Then
      Set TempFt = New CFeatureType
      TempFt.FEATURE_TYPE_ID = -1
      TempFt.FEATURE_TYPE_NO = "*"
      Set Cols = m_FeatureType
      Call Cols.add(TempFt)
      Set TempFt = Nothing
   Else '3
      Set TempBa = New CMasterRef
      TempBa.KEY_ID = -1
      TempBa.KEY_CODE = "*"
      Set Cols = m_BankAccounts
      Call Cols.add(TempBa)
      Set TempBa = Nothing
   End If

   For Each Cm In m_Customers
      For Each Obj In Cols
         If Ind = 1 Then
            TempCode = Obj.PART_GROUP_NO
            TempID = Obj.PART_GROUP_ID
            TempName = "PART_GROUP_ID"
         ElseIf Ind = 2 Then
            TempCode = Obj.FEATURE_TYPE_NO
            TempID = Obj.FEATURE_TYPE_ID
            TempName = "FEATURE_TYPE"
         Else '3
            TempCode = Obj.KEY_CODE
            TempID = Obj.KEY_ID
            TempName = "BANK_ACCOUNT_ID"
         End If
'If Cm.CUSTOMER_CODE = "อ-0082" Then
''Debug.Print
'End If
         Key1 = Cm.CUSTOMER_CODE & "-" & TempCode
         Key2 = "*" & "-" & TempCode
         Set TmpCai = GetCustomerAccountList(TempCol, Key1)
         
         If TmpCai Is Nothing Then
            'หาไม่พบ
            Set TmpCai = GetCustomerAccountList(TempCol, Key2)
            If TmpCai Is Nothing Then
               FoundFlag = False
            Else
               FoundFlag = True
            End If
         Else
            'หาพบ
            FoundFlag = True
         End If
         
         If FoundFlag Then
            Acc1 = TmpCai.GetFieldValue("DEBIT_NO")
            Acc2 = TmpCai.GetFieldValue("CREDIT_NO")
            
            Set TmpCai = Nothing
            
            Set TmpCai = New CCustomerAccountList
            
            TmpCai.ShowMode = SHOW_ADD
            Call TmpCai.SetFieldValue("CUSTOMER_ID", Cm.CUSTOMER_ID)
            Call TmpCai.SetFieldValue(TempName, TempID)
            Call TmpCai.SetFieldValue("ACCOUNT_LIST_TYPE", Ind)
            DrID = GetAccountNoID(Acc1)
            CrID = GetAccountNoID(Acc2)
            Call TmpCai.SetFieldValue("DEBIT_ID", DrID)
            Call TmpCai.SetFieldValue("CREDIT_ID", CrID)
            Call TmpCai.AddEditData
            
            Set TmpCai = Nothing
         End If
      Next Obj
      
      ProgressCount = ProgressCount + 1
      prgProgress.Value = ProgressCount
   Next Cm
      
   prgProgress.Value = prgProgress.MAX
   
   Call EnableForm(Me, True)
   glbDatabaseMngr.DBConnection.CommitTrans
   HasBegin = False
   
   Set m_ExcelSheet = Nothing
   Set TempCol = Nothing
   
   cmdStart.Enabled = True
   cmdExit.Enabled = True
   cmdOK.Enabled = True
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

Private Function GetAccountNoID(AccCode As String) As Long
Dim Mr As CMasterRef
Dim TempRs As ADODB.Recordset
Dim iCount As Long

   If Trim(AccCode) = "" Then
      GetAccountNoID = -1
      Exit Function
   End If
   
   Set TempRs = New ADODB.Recordset
   
   Set Mr = New CMasterRef
   Mr.KEY_ID = -1
   Mr.KEY_CODE = AccCode
   Mr.MASTER_AREA = ACCOUNT_LIST
   Call Mr.QueryData(TempRs, iCount)
   If TempRs.EOF Then
      Mr.AddEditMode = SHOW_ADD
      Mr.SUM_FLAG = "N"
      Mr.MASTER_FLAG = "N"
      Call Mr.AddEditData
   Else
      Call Mr.PopulateFromRS(1, TempRs)
   End If
   
   GetAccountNoID = Mr.KEY_ID
   
   If Not TempRs.EOF Then
      Call TempRs.Close
   End If
   Set TempRs = Nothing
   
   Set Mr = Nothing
End Function

Private Function GetVal(row As Long, Col As Long) As Double
On Error Resume Next

   GetVal = m_ExcelSheet.Cells(row, Col).Value
End Function

Private Sub Form_Activate()
Dim FromDate As Date
Dim ToDate As Date

   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call InitImportType2(cboImportType)
      Call LoadCustomer(Nothing, m_Customers)
      Call LoadPartGroup(Nothing, m_PartGroups)
      Call LoadFeatureType(Nothing, m_FeatureType)
      Call LoadMaster(Nothing, m_BankAccounts, BANK_ACCOUNT)
      
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
   
   Call InitNormalLabel(lblFileName, "ประเภท")
   Call InitNormalLabel(lblMasterName, "ชื่อไฟล์")
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(Label1, "%")

'   Call InitCheckBox(chkBalanceFlag, "ลบยอดยกมา")
'   chkBalanceFlag.Value = ssCBUnchecked

   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPercent.Enabled = False
   Call txtFileName.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtFileName.Enabled = False
   
   Call InitCombo(cboImportType)
   
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
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   Call EnableForm(Me, False)
   m_HasActivate = False
   
   Set m_Employee = New CEmployee
   Set m_Rs = New ADODB.Recordset
   Set m_PartGroups = New Collection
   Set m_FeatureTypeDateLocations = New Collection
   Set m_FeatureType = New Collection
   Set m_BankAccounts = New Collection
   Set m_Customers = New Collection
   Set m_Accounts = New Collection
   Set m_AccountCodes = New Collection
   Set m_SheetBalances = New Collection
   Set m_Suppliers = New Collection
   
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
   Set m_PartGroups = Nothing
   Set m_FeatureTypeDateLocations = Nothing
   Set m_FeatureType = Nothing
   Set m_BankAccounts = Nothing
   Set m_Customers = Nothing
   Set m_Accounts = Nothing
   Set m_SheetBalances = Nothing
   Set m_AccountCodes = Nothing
   Set m_Suppliers = Nothing
End Sub

