VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportDoc 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Icon            =   "frmImportDoc.frx":0000
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
         MouseIcon       =   "frmImportDoc.frx":27A2
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
         MouseIcon       =   "frmImportDoc.frx":2ABC
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
         MouseIcon       =   "frmImportDoc.frx":2DD6
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
         MouseIcon       =   "frmImportDoc.frx":30F0
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmImportDoc"
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

Private m_Balances As Collection
Private m_PartItemsDateLocations As Collection
Private m_ExcelApp As Object
Private m_ExcelSheet As Object
Private m_Locations As Collection
Private m_PartItems As Collection
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
      
      Call ImportStock
      
      m_ExcelApp.Workbooks.Close
      Call EnableForm(Me, True)
   ElseIf TempID = 2 Then
      Call EnableForm(Me, False)
      m_ExcelApp.Workbooks.Close
      m_ExcelApp.Workbooks.Open (txtFileName.Text)
      
      Call ImportARBalance
      
      m_ExcelApp.Workbooks.Close
      Call EnableForm(Me, True)
   ElseIf TempID = 3 Then
      Call EnableForm(Me, False)
      m_ExcelApp.Workbooks.Close
      m_ExcelApp.Workbooks.Open (txtFileName.Text)
      
      Call ImportCustomer
      
      m_ExcelApp.Workbooks.Close
      Call EnableForm(Me, True)
   ElseIf TempID = 4 Then
      Call EnableForm(Me, False)
      m_ExcelApp.Workbooks.Close
      m_ExcelApp.Workbooks.Open (txtFileName.Text)
      
      Call ImportAdjust
      
      m_ExcelApp.Workbooks.Close
      Call EnableForm(Me, True)
   ElseIf TempID = 5 Then
      Call EnableForm(Me, False)
      m_ExcelApp.Workbooks.Close
      m_ExcelApp.Workbooks.Open (txtFileName.Text)
      
      Call ImportChangeCustomer
      
      m_ExcelApp.Workbooks.Close
      Call EnableForm(Me, True)
   ElseIf TempID = 6 Then
      Call EnableForm(Me, False)
      m_ExcelApp.Workbooks.Close
      m_ExcelApp.Workbooks.Open (txtFileName.Text)
      
      Call AdjustStock
      
      m_ExcelApp.Workbooks.Close
      Call EnableForm(Me, True)
   ElseIf TempID = 7 Then
      Call EnableForm(Me, False)
      m_ExcelApp.Workbooks.Close
      m_ExcelApp.Workbooks.Open (txtFileName.Text)
      
      Call ImportSpBalance
      
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
   Ivd.DOCUMENT_NO = "ʵ�ͤ¡��"
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
      Set Lc = GetLocation(m_Locations, Trim(m_ExcelSheet.Cells(row, 1).Value))
      II.LOCATION_ID = Lc.LOCATION_ID
      Set Pi = GetPartItem(m_PartItems, Trim("S-" & m_ExcelSheet.Cells(row, 2).Value))
      II.PART_ITEM_ID = Pi.PART_ITEM_ID
      II.TX_TYPE = "I"
      If (II.PART_ITEM_ID > 0) And (II.LOCATION_ID > 0) Then
         Call Ivd.ImportExports.add(II)
      Else
         glbErrorLog.LocalErrorMsg = "��辺�������ѵ�شԺ '" & Trim(m_ExcelSheet.Cells(row, 3).Value) & "' ���� '" & Trim(m_ExcelSheet.Cells(row, 2).Value) & "' 㹰ҹ������"
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
Private Sub AdjustStock()
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
   Ivd.DOCUMENT_NO = "ʵ�ͤ¡��"
   Ivd.DOCUMENT_DATE = InternalDateToDateEx2(m_ExcelApp.Sheets(id).NAME)
   Ivd.COMMIT_FLAG = "N"
   Ivd.DOCUMENT_TYPE = 4
   Ivd.ADJUST_FLAG = "Y"
   Ivd.EXCEPTION_FLAG = "N"
      
   For row = 3 To MaxRow
      DoEvents
      Me.Refresh
      
      Set II = New CLotItem
      II.Flag = "A"
      II.NEED_TOTAL_PRICE = Val(m_ExcelSheet.Cells(row, 6).Value)
      II.CALCULATE_FLAG = "Y"
      II.NEED_TOTAL_AMOUNT = Val(m_ExcelSheet.Cells(row, 4).Value)
      II.NEED_AVG_PRICE = MyDiff(II.NEED_TOTAL_PRICE, II.NEED_TOTAL_AMOUNT)
      II.CALCULATE_TYPE = 3
      II.TOTAL_WEIGHT = II.NEED_TOTAL_AMOUNT
      Set Lc = GetLocation(m_Locations, Trim(m_ExcelSheet.Cells(row, 1).Value))
      II.LOCATION_ID = Lc.LOCATION_ID
      Set Pi = GetPartItem(m_PartItems, Trim("S-" & m_ExcelSheet.Cells(row, 2).Value))
      II.PART_ITEM_ID = Pi.PART_ITEM_ID
      II.TX_TYPE = "I"
      If (II.PART_ITEM_ID > 0) And (II.LOCATION_ID > 0) Then
         Call Ivd.ImportExports.add(II)
      Else
         glbErrorLog.LocalErrorMsg = "��辺�������ѵ�شԺ '" & Trim(m_ExcelSheet.Cells(row, 3).Value) & "' ���� '" & Trim(m_ExcelSheet.Cells(row, 2).Value) & "' 㹰ҹ������"
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

Private Sub LoadBalances()
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
Dim II As CLotItem

   HasBegin = False

   id = 2
   
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
   
'   glbDatabaseMngr.DBConnection.BeginTrans
   HasBegin = True
      
   For row = 2 To MaxRow
      DoEvents
      
      Set II = New CLotItem
      
      II.TOTAL_INCLUDE_PRICE = Val(m_ExcelSheet.Cells(row, 5).Value)
      II.TX_AMOUNT = Val(m_ExcelSheet.Cells(row, 4).Value)
      II.LOCATION_NO = Trim(m_ExcelSheet.Cells(row, 1).Value)
      II.PART_NO = Trim(m_ExcelSheet.Cells(row, 2).Value)

      Call m_SheetBalances.add(II, II.LOCATION_NO & "-" & II.PART_NO)
      
      Set II = Nothing
      
      ProgressCount = ProgressCount + 1
      prgProgress.Value = ProgressCount
   Next row
   
   prgProgress.Value = prgProgress.MAX
   
   Call EnableForm(Me, True)
'   glbDatabaseMngr.DBConnection.CommitTrans
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

Private Function GetVal(row As Long, Col As Long) As Double
On Error Resume Next

   GetVal = m_ExcelSheet.Cells(row, Col).Value
End Function

Private Sub ImportAdjust()
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
Dim TxAmount As Double
Dim TotalPrice As Double
Dim NeedAmount As Double
Dim NeedPrice As Double
Dim NeedAvgPrice As Double
Dim Ba As CBalanceAccum

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
   
'   Set Ivd = New CInventoryDoc
'   Ivd.AddEditMode = SHOW_ADD
'   Ivd.DOCUMENT_NO = "ADJ-SEUB"
'   Ivd.DOCUMENT_DATE = InternalDateToDateEx2(m_ExcelApp.Sheets(ID).Name)
'   Ivd.COMMIT_FLAG = "N"
'   Ivd.DOCUMENT_TYPE = 4
'   Ivd.EXCEPTION_FLAG = "N"
'   Ivd.ADJUST_FLAG = "Y"
   
   For row = 2 To MaxRow
      DoEvents
      TxAmount = m_ExcelSheet.Cells(row, 8).Value
      TotalPrice = Val(m_ExcelSheet.Cells(row, 9).Value)
      NeedAmount = Val(m_ExcelSheet.Cells(row, 6).Value)
      NeedPrice = Val(m_ExcelSheet.Cells(row, 7).Value)
      NeedAvgPrice = GetVal(row, 10)
      
      Set Lc = GetLocation(m_Locations, Trim(m_ExcelSheet.Cells(row, 1).Value))
      Set Pi = GetPartItem(m_PartItems, Trim(m_ExcelSheet.Cells(row, 2).Value))
      
      Set Ba = New CBalanceAccum
      Ba.AddEditMode = SHOW_ADD
      Ba.DOCUMENT_DATE = InternalDateToDateEx2(m_ExcelApp.Sheets(id).NAME)
      Ba.IMPORT_AMOUNT = TxAmount
      Ba.EXPORT_AMOUNT = 0
      Ba.BALANCE_AMOUNT = TxAmount
      Ba.TOTAL_INCLUDE_PRICE = TotalPrice
      Ba.AVG_PRICE = NeedAvgPrice
      Ba.PART_ITEM_ID = Pi.PART_ITEM_ID
      Ba.LOCATION_ID = Lc.LOCATION_ID
      If (Ba.PART_ITEM_ID > 0) And (Ba.LOCATION_ID > 0) Then
         Call Ba.AddEditData
      Else
         glbErrorLog.LocalErrorMsg = "��辺�������ѵ�شԺ '" & Trim(m_ExcelSheet.Cells(row, 2).Value) & "' ��ѧ '" & Trim(m_ExcelSheet.Cells(row, 1).Value) & "' 㹰ҹ������"
         glbErrorLog.ShowUserError
      End If
      Set Ba = Nothing
      
'''Debug.Print NeedAvgPrice
'      Set II = New CLotItem
'      II.Flag = "A"
'
'      II.TX_AMOUNT = Abs(TxAmount)
'      If TxAmount > 0 Then
'         II.TX_TYPE = "I"
'         II.TOTAL_INCLUDE_PRICE = TotalPrice
'      Else
'         II.TX_TYPE = "E"
'         II.TOTAL_INCLUDE_PRICE = -1 * TotalPrice
'      End If
'      II.ACTUAL_UNIT_PRICE = II.INCLUDE_UNIT_PRICE
'      II.TOTAL_ACTUAL_PRICE = II.TOTAL_INCLUDE_PRICE
'      II.CALCULATE_FLAG = "Y"
'      II.INCLUDE_UNIT_PRICE = Minus2Zero(MyDiff(II.TOTAL_INCLUDE_PRICE, II.TX_AMOUNT))
'      II.ACTUAL_UNIT_PRICE = II.INCLUDE_UNIT_PRICE
'      II.CALCULATE_TYPE = 0
'      II.TOTAL_WEIGHT = II.TX_AMOUNT
'      II.NEED_TOTAL_AMOUNT = NeedAmount
'      II.NEED_TOTAL_PRICE = NeedPrice
'      II.NEED_AVG_PRICE = NeedAvgPrice
'      Set Lc = GetLocation(m_Locations, Trim(m_ExcelSheet.Cells(Row, 1).Value))
'      II.LOCATION_ID = Lc.LOCATION_ID
'      Set Pi = GetPartItem(m_PartItems, Trim(m_ExcelSheet.Cells(Row, 2).Value))
'      II.PART_ITEM_ID = Pi.PART_ITEM_ID
'      If (II.PART_ITEM_ID > 0) And (II.LOCATION_ID > 0) Then
'         Call Ivd.ImportExports.add(II)
'      Else
'         glbErrorLog.LocalErrorMsg = "��辺�������ѵ�شԺ '" & Trim(m_ExcelSheet.Cells(Row, 2).Value) & "' ��ѧ '" & Trim(m_ExcelSheet.Cells(Row, 1).Value) & "' 㹰ҹ������"
'         glbErrorLog.ShowUserError
'      End If
'
'      Set II = Nothing
      
      ProgressCount = ProgressCount + 1
      prgProgress.Value = ProgressCount
   Next row
'   Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog)
   
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

Private Sub UpdateBalance()
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
Dim TxAmount As Double
Dim TotalPrice As Double

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
   
   HasBegin = True
      
   For row = 2 To MaxRow
      DoEvents
      Set II = GetLotItem(m_SheetBalances, Trim(m_ExcelSheet.Cells(row, 1).Value) & "-" & Trim(m_ExcelSheet.Cells(row, 2).Value))
      m_ExcelSheet.Cells(row, 8).Value = II.TX_AMOUNT
      m_ExcelSheet.Cells(row, 9).Value = II.TOTAL_INCLUDE_PRICE
      
      ProgressCount = ProgressCount + 1
      prgProgress.Value = ProgressCount
   Next row
   
   Set Ivd = Nothing
   prgProgress.Value = prgProgress.MAX
   
   Call EnableForm(Me, True)
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

Private Function GetCustomerFromExpCode(Col As Collection, Key As String) As CCustomer
Dim Cm As CCustomer

   For Each Cm In Col
      If Cm.EXP_CODE = Key Then
         Set GetCustomerFromExpCode = Cm
         Exit Function
      End If
   Next Cm
   
   Set GetCustomerFromExpCode = Nothing
End Function

Private Function GetAccountFromCustID(Col As Collection, id As Long) As CAccount
Dim Cm As CAccount

   For Each Cm In Col
      If Cm.CUSTOMER_ID = id Then
         Set GetAccountFromCustID = Cm
         Exit Function
      End If
   Next Cm
   
   Set GetAccountFromCustID = Nothing
End Function

Private Sub ImportCustomer()
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
Dim Cm As CCustomer
Dim CstName As CCustomerName
Dim NAME As CName
Dim Acc As CAccount
Dim IsOK As Boolean
Dim tempCm As CCustomer
Dim SB As CSubscriber
Dim Agr As CAgreement

   HasBegin = False

   id = 2
   
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
   
   For row = 2 To MaxRow
      DoEvents

      Set tempCm = Nothing
       Set tempCm = GetCustomerFromExpCode(m_Customers, Trim(m_ExcelSheet.Cells(row, 1).Value))
       
      If tempCm Is Nothing Then
         Set Cm = New CCustomer
         
         Cm.AddEditMode = SHOW_ADD
         Cm.CUSTOMER_CODE = Trim(m_ExcelSheet.Cells(row, 1).Value)
         Cm.EXP_CODE = Cm.CUSTOMER_CODE
         Cm.Credit = Val(Trim(m_ExcelSheet.Cells(row, 3).Value))
         Cm.CUSTOMER_NAME = Trim(m_ExcelSheet.Cells(row, 2).Value)
         Cm.CUSTOMER_GRADE = -1
         Cm.CUSTOMER_TYPE = -1
         
         If Cm.CstNames.Count <= 0 Then
            Set CstName = New CCustomerName
            CstName.Flag = "A"
            
            Set NAME = CstName.NAME
            NAME.LONG_NAME = Cm.CUSTOMER_NAME
            NAME.SHORT_NAME = ""
            NAME.Flag = "A"
            
            Call Cm.CstNames.add(CstName)
         End If
         
         If Cm.CstAccounts.Count <= 0 Then
            Set Acc = New CAccount
            Acc.ACCOUNT_NO = Cm.CUSTOMER_CODE
            Acc.MASTER_FLAG = "Y"
            Acc.ENABLE_FLAG = "Y"
            Acc.Flag = "A"
                        
            Set SB = New CSubscriber
            SB.Flag = "A"
            SB.DUMMY_FLAG = "Y"
            SB.SUBSCRIBER_NO = Cm.CUSTOMER_CODE
            SB.SUBSCRIBER_STATUS = "Y"
            Call Acc.ActSubs.add(SB)
            Set SB = Nothing
            
            Set Agr = New CAgreement
            Agr.Flag = "A"
            Agr.EXCLUDE_FLAG = "N"
            Agr.SOC_ID = -1
            Call Acc.ActAgrmnts.add(Agr)
            Set Agr = Nothing
            
            Call Cm.CstAccounts.add(Acc)
            Set Acc = Nothing
         End If
         
         Call glbDaily.AddEditCustomer(Cm, IsOK, False, glbErrorLog)
         
         ProgressCount = ProgressCount + 1
         prgProgress.Value = ProgressCount
         
         Set Cm = Nothing
      End If
   Next row
   
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

Private Sub ImportChangeCustomer()
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
Dim Cm1 As CCustomer
Dim Cm2 As CCustomer
Dim BD As CBillingDoc
Dim Ca1 As CAccount
Dim Ca2 As CAccount

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
   
   For row = 2 To MaxRow
      DoEvents

      For Each Ca1 In m_AccountCodes
         If Ca1.CUSTOMER_CODE = Trim(m_ExcelSheet.Cells(row, 1).Value) Then
            Exit For
         End If
      Next Ca1
      
      For Each Ca2 In m_AccountCodes
         If Ca2.CUSTOMER_CODE = Trim(m_ExcelSheet.Cells(row, 2).Value) Then
            Exit For
         End If
      Next Ca2
                     
      If (Not Ca1 Is Nothing) And (Not Ca2 Is Nothing) Then
         Set BD = New CBillingDoc
         BD.AddEditMode = SHOW_EDIT
         BD.ACCOUNT_ID = Ca1.ACCOUNT_ID
         BD.NEW_ACCOUNT_ID = Ca2.ACCOUNT_ID
         Call BD.UpdateNewAccount
         Set BD = Nothing
      End If
   Next row
   
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

Private Sub ImportARBalance()
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
Dim IsOK As Boolean
Dim BD As CBillingDoc
Dim Cm As CCustomer
Dim Ac As CAccount
Dim Di As CDoItem
Dim Ft As CFeature

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
      
   For row = 2 To MaxRow
      DoEvents
      
      Set BD = New CBillingDoc
      Set Di = New CDoItem
      
      BD.AddEditMode = SHOW_ADD
      BD.DOCUMENT_NO = m_ExcelSheet.Cells(row, 3).Value
      BD.DOCUMENT_DATE = m_ExcelSheet.Cells(row, 4).Value
      BD.DUE_DATE = m_ExcelSheet.Cells(row, 5).Value
      BD.COMMIT_FLAG = "N"
      BD.EXCEPTION_FLAG = "N"
      BD.DOCUMENT_TYPE = 1
      
      Set Cm = GetCustomerFromExpCode(m_Customers, Trim(m_ExcelSheet.Cells(row, 2).Value))
'      Set Cm = GetCustomerFromExpCode(m_Customers, "XXX")
      If Not (Cm Is Nothing) Then
         BD.CUSTOMER_ID = Cm.CUSTOMER_ID
         Set Ac = GetAccountFromCustID(m_Accounts, Cm.CUSTOMER_ID)
         BD.ACCOUNT_ID = Ac.ACCOUNT_ID
      
         Di.PART_ITEM_ID = -1
         Di.FEATURE_ID = -1
         Di.LOCATION_ID = -1
         Di.ITEM_AMOUNT = 0
         Di.TOTAL_PRICE = Val(m_ExcelSheet.Cells(row, 6).Value)
         Di.CONFIG_CODE = "NNY"
         Di.ITEM_DESC = "��¡���ʹ¡��"
         Di.Flag = "A"
         Di.MANUAL_FLAG = "N"
         Di.DISPLAY_ID = 1
         Di.ITEM_AMOUNT = 1
         Di.AVG_PRICE = Di.TOTAL_PRICE
         
         Call BD.DoItems.add(Di)
         
         Call glbDaily.AddEditBillingDoc(BD, IsOK, False, glbErrorLog)
      Else
''Debug.Print Trim(m_ExcelSheet.Cells(row, 2).Value)
      End If
      
      Set Di = Nothing
      Set BD = Nothing
      
      ProgressCount = ProgressCount + 1
      prgProgress.Value = ProgressCount
   Next row
   
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
Private Sub ImportSpBalance()
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
Dim IsOK As Boolean
Dim BD As CBillingDoc
Dim Cm As CSupplier
Dim Di As CSupItem
Dim TempSetDate As String
Dim DD As String
Dim MM As String
Dim YYYY As String
Dim Lc As Byte
Dim TempLc As Byte
Dim TempCode As String
Dim TempDate As Date

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
   
   For row = 1 To MaxRow
      DoEvents
      
      Set BD = New CBillingDoc
      Set Di = New CSupItem
      
      BD.AddEditMode = SHOW_ADD
      BD.DOCUMENT_NO = m_ExcelSheet.Cells(row, 4).Value
      
      BD.DOCUMENT_DATE = DateSerial(Mid(m_ExcelSheet.Cells(row, 5).Value, 1, 4), Mid(m_ExcelSheet.Cells(row, 5).Value, 8, 2), Mid(m_ExcelSheet.Cells(row, 5).Value, 13, 2))
      
      TempDate = DateSerial(Mid(m_ExcelSheet.Cells(row, 3).Value, 1, 4), Mid(m_ExcelSheet.Cells(row, 3).Value, 8, 2), Mid(m_ExcelSheet.Cells(row, 3).Value, 13, 2))
      BD.DUE_DATE = TempDate
      
      BD.COMMIT_FLAG = "N"
      BD.EXCEPTION_FLAG = "N"
      BD.DOCUMENT_TYPE = 100
      
      Set Cm = GetSupplier(m_Suppliers, Trim(m_ExcelSheet.Cells(row, 2).Value))
      If Not (Cm Is Nothing) Then
         BD.SUPPLIER_ID = Cm.SUPPLIER_ID
         If BD.SUPPLIER_ID <= 0 And Not (TempCode = Trim(m_ExcelSheet.Cells(row, 2).Value)) Then
            TempCode = Trim(m_ExcelSheet.Cells(row, 2).Value)
            glbErrorLog.LocalErrorMsg = "��������� �Ѿ��������� " & Trim(m_ExcelSheet.Cells(row, 2).Value)
            Call glbErrorLog.ShowErrorLog(LOG_TO_FILE)
         End If
         
         
         Di.PART_ITEM_ID = -1
         'Di.FEATURE_ID = -1
         Di.LOCATION_ID = -1
         Di.TOTAL_INCLUDE_PRICE = Val(m_ExcelSheet.Cells(row, 6).Value)
         Di.ITEM_DESC = "��¡���ʹ¡��"
         Di.Flag = "A"
         Di.TX_AMOUNT = 1
         Di.TOTAL_ACTUAL_PRICE = Di.TOTAL_INCLUDE_PRICE
         Di.INCLUDE_UNIT_PRICE = Di.TOTAL_INCLUDE_PRICE
         Di.ACTUAL_UNIT_PRICE = Di.TOTAL_INCLUDE_PRICE
         Di.CALCULATE_FLAG = "N"
         
         Call BD.SupItems.add(Di)
         
         Call glbDaily.AddEditBillingDoc(BD, IsOK, False, glbErrorLog)
      Else
         glbErrorLog.LocalErrorMsg = "��������� �Ѿ��������� " & Trim(m_ExcelSheet.Cells(row, 2).Value)
         Call glbErrorLog.ShowErrorLog(LOG_TO_FILE)
      End If
      
      Set Di = Nothing
      Set BD = Nothing
      
      ProgressCount = ProgressCount + 1
      prgProgress.Value = ProgressCount
   Next row
   
   prgProgress.Value = prgProgress.MAX
   
   Call EnableForm(Me, True)
   
   If ConfirmSave Then
      glbDatabaseMngr.DBConnection.CommitTrans
   Else
      glbDatabaseMngr.DBConnection.RollbackTrans
   End If
   
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

Private Sub Form_Activate()
Dim FromDate As Date
Dim ToDate As Date

   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call InitImportType(cboImportType)
      Call LoadLocationCodeKey(Nothing, m_Locations)
      Call LoadPartItemCodeKey(Nothing, m_PartItems, , "")
      Call LoadCustomer(Nothing, m_Customers)
      Call LoadSupplier(Nothing, m_Suppliers, 2)
      Call LoadAccount(Nothing, m_Accounts)
      Call LoadAccountByCustCode(Nothing, m_AccountCodes)
      
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
   pnlHeader.Caption = "������쵢�����"
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitNormalLabel(lblFileName, "������")
   Call InitNormalLabel(lblMasterName, "�������")
   Call InitNormalLabel(lblProgress, "�����׺˹��")
   Call InitNormalLabel(lblPercent, "����ૹ��")
   Call InitNormalLabel(Label1, "%")

'   Call InitCheckBox(chkBalanceFlag, "ź�ʹ¡��")
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
   
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
   Call InitMainButton(cmdStart, MapText("�����"))
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
   Set m_Balances = New Collection
   Set m_PartItemsDateLocations = New Collection
   Set m_Locations = New Collection
   Set m_PartItems = New Collection
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
   Set m_Balances = Nothing
   Set m_PartItemsDateLocations = Nothing
   Set m_Locations = Nothing
   Set m_PartItems = Nothing
   Set m_Customers = Nothing
   Set m_Accounts = Nothing
   Set m_SheetBalances = Nothing
   Set m_AccountCodes = Nothing
   Set m_Suppliers = Nothing
End Sub

Private Sub SSCommand1_Click()
Dim Li As CLotItem

   Call EnableForm(Me, False)
   m_ExcelApp.Workbooks.Close
   m_ExcelApp.Workbooks.Open (txtFileName.Text)
   
   Call LoadBalances
   Call UpdateBalance
   
   m_ExcelApp.Workbooks.Close
   Call EnableForm(Me, True)
End Sub
