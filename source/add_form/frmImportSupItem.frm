VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmImportSupItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Icon            =   "frmImportSupItem.frx":0000
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
         MouseIcon       =   "frmImportSupItem.frx":27A2
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
         MouseIcon       =   "frmImportSupItem.frx":2ABC
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
         MouseIcon       =   "frmImportSupItem.frx":2DD6
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmImportSupItem"
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

Private c_DocumentNos As Collection

Private CountBill As Long
Private CountDown As Double

Private Bl As CBillingDoc
Private Ivd As CInventoryDoc

Private EmpColls As Collection
Private SupColls As Collection
Private PartColls As Collection
Private LocationColls As Collection
Private CnDnRtColls As Collection

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
      Call ImportSupplier
   ElseIf TempID = 2 Then
      Call ImportPartItem
   ElseIf TempID = 3 Then
      Call ImportPo
   ElseIf TempID = 4 Then
      Call ImportSupItem
   ElseIf TempID = 5 Then
      Call ImportCnDnRt
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
   
   'FromDate = DateSerial(Mid(FileName1, 10, 4), Mid(FileName1, 14, 2), Mid(FileName1, 16, 2))
   ToDate = DateSerial(Mid(FileName1, 19, 4), Mid(FileName1, 23, 2), Mid(FileName1, 25, 2))
   Call LoadBillingDocDistinctDocumentNo(Nothing, c_DocumentNos, FromDate, ToDate)
   
   Call LoadSupplier(Nothing, SupColls, 2)
   Call LoadPartItem(Nothing, PartColls, , , , 2)
   
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
Dim EmpCode As String

Dim ChkUnigueBillingDoc As CBillingDoc

Dim TimeStamp As Date
Dim TimeStampStr As String
Dim TempAsc As Long
Dim OldTempAsc As Long
Dim firstDate As Date
Dim lastDate As Date
Dim I As Long
Dim ItemCount As Long
Dim IsOK As Boolean
Dim Si As CSupItem
Dim Key1 As String
Dim Key2 As String
Dim Key3 As String
Dim Key4 As String

   If Left(LineStr, 2) = "BD" Then
      If CountBill > 0 Then
         If DocumentType = 100 Then   'ใบรับเข้าวัตถุดิบ
            Call glbDaily.SUP2InventoryDoc(Bl, Ivd, 1)
         ElseIf DocumentType = 101 Then   'ใบรับเข้าวัสดุอุปกรณ์
            Call glbDaily.SUP2InventoryDoc(Bl, Ivd, 19)
         ElseIf DocumentType = 102 Then   'ใบรับเข้าจ่ายออกวัสดุอุปกรณ์
            Call glbDaily.SUP2InventoryDocEx(Bl, Ivd, 20)
         ElseIf DocumentType = 103 Then   'ใบรับเข้าทั่วไป
            Call glbDaily.SUP2InventoryDoc(Bl, Ivd, 23)
         End If
         
         If Ivd.INVENTORY_DOC_ID = 109780 Then
            'Debug.Print
         End If
         
         Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog)
         
         Bl.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID
         
         
         Call glbDaily.AddEditBillingDoc(Bl, IsOK, False, glbErrorLog)
         
         Set Bl = New CBillingDoc
         Set Ivd = New CInventoryDoc
      
      End If
      CountBill = 1
      
      TempAsc = 3
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      
      If Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1) = "01-0263-09" Then
         'Debug.Print
      End If
      If Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1) = "01-0264-09" Then
         'Debug.Print
      End If
      
      Set ChkUnigueBillingDoc = GetObject("CBillingDoc", c_DocumentNos, Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1), False)
      If ChkUnigueBillingDoc Is Nothing Then
         Bl.AddEditMode = SHOW_ADD
      Else
         Bl.AddEditMode = SHOW_EDIT
         
         Set Si = New CSupItem
         Si.DO_ID = ChkUnigueBillingDoc.BILLING_DOC_ID
         Si.INVENTORY_DOC_ID = ChkUnigueBillingDoc.INVENTORY_DOC_ID
         Bl.BILLING_DOC_ID = ChkUnigueBillingDoc.BILLING_DOC_ID
         Bl.INVENTORY_DOC_ID = ChkUnigueBillingDoc.INVENTORY_DOC_ID
         
         Call Si.DeleteFromBillInv
         Set Si = Nothing
         
         Bl.EditonlyFromChild = True
      End If
      Bl.DOCUMENT_NO = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)
      OldTempAsc = TempAsc
      
      Bl.DOCUMENT_DATE = InternalDateToDate(StingToVariable(TempAsc, OldTempAsc, LineStr))
      Bl.DOCUMENT_TYPE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      DocumentType = Bl.DOCUMENT_TYPE
      Bl.DUE_DATE = InternalDateToDate(StingToVariable(TempAsc, OldTempAsc, LineStr))
      Bl.NOTE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Bl.BILLING_ADDRESS_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.BILLING_ADDRESS_ID = -1
      Bl.ENTERPRISE_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.ENTERPRISE_ADDRESS_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.TOTAL_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.VAT_PERCENT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Bl.VAT_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.WH_PERCENT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.WH_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.DISCOUNT_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.PAYMENT_TERM = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
'      Dim Emp As CEmployee
'      TempAsc = InStr(TempAsc + 1, LineStr, ";")
'      Set Emp = GetEmployee(EmpColls, Trim(Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)))
'      If Emp.EMP_CODE <> Trim(Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)) Then
'         Call MsgBox("ยังไม่มีรหัสพนักงาน " & Trim(Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)) & " นี้ในระบบ กรุณาใส่เพิ่มเติ่มในภายหลัง ", vbOKOnly, PROJECT_NAME)
'      End If
'      Bl.ACCEPT_BY = Emp.EMP_ID                                                     '13
'      OldTempAsc = TempAsc
      
      Bl.ACCEPT_BY = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
'      TempAsc = InStr(TempAsc + 1, LineStr, ";")
'      Set Emp = GetEmployee(EmpColls, Trim(Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)))
'      Bl.RECEIVE_BY = Emp.EMP_ID                                                     '14
'      OldTempAsc = TempAsc
      Bl.RECEIVE_BY = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Bl.EXCEPTION_FLAG = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.PAYEE_NAME = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.COMMIT_FLAG = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Dim Sp As CSupplier
      Bl.SUPPLIER_CODE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Set Sp = GetSupplier(SupColls, Trim(Bl.SUPPLIER_CODE))
      If Sp.SUPPLIER_CODE <> Bl.SUPPLIER_CODE Then
         glbErrorLog.SystemErrorMsg = "ยังไม่มีรหัสซัพพลายเออร์ " & Trim(Bl.SUPPLIER_CODE) & " นี้ในระบบ กรุณาใส่เพิ่มเติ่มในภายหลัง "
         glbErrorLog.ShowErrorLog (LOG_TO_FILE)
      End If
      Bl.SUPPLIER_ID = Sp.SUPPLIER_ID

      'Bl.SUPPLIER_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.RECEIPT_TYPE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.ACCOUNT_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.DEPOSIT_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.TOTAL_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Bl.APPROVE_DATE = InternalDateToDate(StingToVariable(TempAsc, OldTempAsc, LineStr))
      Bl.ESTIMATE_DATE = InternalDateToDate(StingToVariable(TempAsc, OldTempAsc, LineStr))
      
      Bl.RESOURCE_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.PAYMENT_TYPE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.BANK_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.BBRANCH_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.CHECK_NO = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Bl.CHECK_DATE = InternalDateToDate(StingToVariable(TempAsc, OldTempAsc, LineStr))
      Bl.VADILITY = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.DELIVERY = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.PAYMENT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.PAYMENT_DESC = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Bl.SHIPMENT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.REFER_INV = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.PACKING_OF = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.SHIPON_DATE = InternalDateToDate(StingToVariable(TempAsc, OldTempAsc, LineStr))
      Bl.REF = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Bl.Credit = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.AREA_CODE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.SHIP_NAME = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.SHIPPING_MARKS = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.CD_PERCENT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Bl.CD_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.PACKAGE_TYPE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.TEMP_DO_NO = StingToVariable(TempAsc, OldTempAsc, LineStr)
      'Bl.PAID_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)                               ' Paid Amount ไม่ต้องเนื่องจาก Key มาจากโรงงานเดี่ยวมันจะทับ
      Call StingToVariable(TempAsc, OldTempAsc, LineStr)
      'Bl.DEBIT_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)                             ' Debit Amount ไม่ต้องเนื่องจาก Key มาจากโรงงานเดี่ยวมันจะทับ
      Call StingToVariable(TempAsc, OldTempAsc, LineStr)
      'Bl.CREDIT_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)                             ' Credit Amount ไม่ต้องเนื่องจาก Key มาจากโรงงานเดี่ยวมันจะทับ
      Call StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.DO_TOTAL_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.REVENUE_TOTAL_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.PAYMENT_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.BANK_BRANCH_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Bl.BANK_NOTE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.TOTAL_RCP = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.RUNNING_NO = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.AGREEMENT_DATA = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.AGREEMENT_FINANCE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Bl.OLD_CREDIT_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.ENTRY_DATE = InternalDateToDate(StingToVariable(TempAsc, OldTempAsc, LineStr))
      Bl.EXIT_DATE = InternalDateToDate(StingToVariable(TempAsc, OldTempAsc, LineStr))
      Bl.DO_NO = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.TRUCK_NO = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Bl.DELIVERY_FEE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.SENDER_NAME = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.RECEIVE_NAME = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.DEPARTMENT_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.DEPARTMENT_ID = -1
      Bl.QUE_NO = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.PR_NO = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      
         
      'FK = Bl.BILLING_DOC_ID
      
   End If
   
   
   If Left(LineStr, 2) = "SI" Then
      Dim Ti As CSupItem
      Set Ti = New CSupItem
      Ti.Flag = "A"
      
      'TI.DO_ID = FK
      
      TempAsc = 3
      OldTempAsc = TempAsc

      Ti.DO_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Dim Pi As CPartItem
      Ti.PART_NO = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Set Pi = GetPartItem(PartColls, Trim(Ti.PART_NO))
      If Ti.PART_NO <> Pi.PART_NO Then
         glbErrorLog.SystemErrorMsg = "ยังไม่มีรหัสสินค้า/วัตถุดิบ " & Trim(Ti.PART_NO) & " นี้ในระบบ กรุณาใส่เพิ่มเติ่มในภายหลัง "
         glbErrorLog.ShowErrorLog (LOG_TO_FILE)
      End If
      Ti.PART_ITEM_ID = Pi.PART_ITEM_ID
      
      'Ti.PART_ITEM_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Dim Lc As CLocation
      Ti.LOCATION_NO = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Set Lc = GetLocation(LocationColls, Trim(Ti.LOCATION_NO))
      If Lc.LOCATION_NO <> Ti.LOCATION_NO Then
         glbErrorLog.SystemErrorMsg = "ยังไม่มีรหัสคลัง " & Trim(Ti.LOCATION_NO) & " นี้ในระบบ กรุณาใส่เพิ่มเติ่มในภายหลัง "
         glbErrorLog.ShowErrorLog (LOG_TO_FILE)
      End If
      Ti.LOCATION_ID = Lc.LOCATION_ID
      
      'Ti.LOCATION_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.ACTUAL_UNIT_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.INCLUDE_UNIT_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.PREVIOUS_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Ti.PREVIOUS_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.NEW_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.TX_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.NEW_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.TRANSACTION_SEQ = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Ti.GUI_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.CALCULATE_FLAG = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.TOTAL_ACTUAL_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.TOTAL_INCLUDE_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.TX_TYPE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Ti.LEFT_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.LAYOUT_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.LINK_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.PACKAGING_AMT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.ENTRY_WEIGHT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Ti.EXIT_WEIGHT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.WEIGHT_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.PACKAGE_WEIGHT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.OTHER_WEIGHT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.PERCENT_HUMID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Ti.HUMID_WEIGHT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.PACKAGING_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.SUPPLIER_WEIGHT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.CALCULATE_TYPE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.PACKAGE_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Ti.ACTUAL_PKG_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.PUREXP_ID1 = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.PUREXP_ID2 = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.EXPENSE1 = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.EXPENSE2 = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Ti.TOTAL_WEIGHT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.RAW_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.DISCOUNT_AMT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.RAW_TOT_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.TO_DEPARTMENT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.TO_DEPARTMENT = -1
      
      Ti.ITEM_DESC = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.EXTRA_NAME = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.SALE_TOT_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.CALCULATE_WEIGHT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.RAW_COST = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Ti.EXPENSE_COST = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.TOTAL_NEW_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.BAG_RETURN = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.ACTUAL_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.ACTUAL_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Ti.CURRENT_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.CURRENT_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.NEED_TOTAL_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.NEED_TOTAL_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.NEED_AVG_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Ti.MANUAL_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.EXPENSE_TYPE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.AUTO_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.ITEM_DESC_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Ti.PO_NO = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Set ChkUnigueBillingDoc = GetObject("CBillingDoc", c_DocumentNos, Trim(Ti.PO_NO), False)
      If ChkUnigueBillingDoc Is Nothing Then
         Ti.PO_ID = -1
      Else
         Ti.PO_ID = ChkUnigueBillingDoc.BILLING_DOC_ID
      End If
      
      Call Bl.SupItems.add(Ti)
      
   End If
   
   If CountDown = 0 Then
      If DocumentType = 100 Then   'ใบรับเข้าวัตถุดิบ
         Call glbDaily.SUP2InventoryDoc(Bl, Ivd, 1)
      ElseIf DocumentType = 101 Then   'ใบรับเข้าวัสดุอุปกรณ์
         Call glbDaily.SUP2InventoryDoc(Bl, Ivd, 19)
      ElseIf DocumentType = 102 Then   'ใบรับเข้าจ่ายออกวัสดุอุปกรณ์
         Call glbDaily.SUP2InventoryDocEx(Bl, Ivd, 20)
      ElseIf DocumentType = 103 Then   'ใบรับเข้าทั่วไป
         Call glbDaily.SUP2InventoryDoc(Bl, Ivd, 23)
      End If
      
      Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog)
      
      Bl.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID
      
      
      Call glbDaily.AddEditBillingDoc(Bl, IsOK, False, glbErrorLog)
      
      Set Bl = New CBillingDoc
      Set Ivd = New CInventoryDoc
      
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
      
      Call InitExportType(cboExportType)
      
      'Call LoadEmployeeCode(Nothing, EmpColls)
      Call LoadSupplier(Nothing, SupColls, 2)
      
      Call LoadPartItem(Nothing, PartColls, , , , 2)
      Call LoadLocation(Nothing, LocationColls, , , , , 2)
      
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
   Set c_DocumentNos = New Collection
   
   Set Bl = New CBillingDoc
   Set Ivd = New CInventoryDoc
   
   Set EmpColls = New Collection
   Set SupColls = New Collection
   Set PartColls = New Collection
   Set LocationColls = New Collection
   Set CnDnRtColls = New Collection
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set c_DocumentNos = Nothing
   
   Set Bl = Nothing
   Set Ivd = Nothing
   
   Set SupColls = Nothing
   Set EmpColls = Nothing
   Set PartColls = Nothing
   Set LocationColls = Nothing
   Set CnDnRtColls = Nothing
   
End Sub
Private Function StingToVariable(TempAsc As Long, OldTempAsc As Long, LineStr As String) As Variant
   TempAsc = InStr(TempAsc + 1, LineStr, ";")
   StingToVariable = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)
   OldTempAsc = TempAsc
End Function
Private Sub ImportSupplier()
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
   'Call LoadBillingDocDistinctDocumentNo(Nothing, c_DocumentNos, FromDate, ToDate)
   
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
      
      If ProcessLine1(TempStr) Then
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
Private Function ProcessLine1(LineStr As String) As Boolean
On Error GoTo ErrorHandler
Dim EmpCode As String

Dim ChkUnigueSupplier As CSupplier
Dim Sp As CSupplier

Dim TimeStamp As Date
Dim TimeStampStr As String
Dim TempAsc As Long
Dim OldTempAsc As Long
Dim firstDate As Date
Dim lastDate As Date
Dim I As Long
Dim ItemCount As Long
Dim IsOK As Boolean
Dim Key1 As String
Dim Key2 As String
Dim Key3 As String
Dim Key4 As String

   If Left(LineStr, 2) = "SP" Then
               
      Set Sp = New CSupplier
      
      TempAsc = 3
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Set ChkUnigueSupplier = GetObject("CSupplier", SupColls, Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1), False)
      If ChkUnigueSupplier Is Nothing Then
         Sp.AddEditMode = SHOW_ADD
      Else
         Sp.AddEditMode = SHOW_EDIT
         Sp.SUPPLIER_ID = ChkUnigueSupplier.SUPPLIER_ID
         Sp.QueryFlag = 1
         If Not glbDaily.QuerySupplier(Sp, m_Rs, ItemCount, IsOK, glbErrorLog) Then
            glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
            Call EnableForm(Me, True)
            Exit Function
         End If
      End If
      
      Sp.SUPPLIER_CODE = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)
      OldTempAsc = TempAsc
      
      Sp.SUPPLIER_GRADE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Sp.Credit = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Sp.TAX_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Sp.SUPPLIER_TYPE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Sp.EMAIL = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Sp.WEBSITE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Sp.BIRTH_DATE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Sp.PASSWORD1 = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Sp.SUPPLIER_STATUS = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Sp.BUSINESS_DESC = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Dim CstName As CSupplierName
      If Sp.AddEditMode = SHOW_ADD Then
         Set CstName = New CSupplierName
         CstName.Flag = "A"
         Call Sp.CstNames.add(CstName)
      Else
         Set CstName = Sp.CstNames.Item(1)
         CstName.Flag = "E"
      End If
      
      Dim NAME As cName
      If Sp.AddEditMode = SHOW_ADD Then
         Set NAME = CstName.NAME
         NAME.LONG_NAME = StingToVariable(TempAsc, OldTempAsc, LineStr)
         NAME.SHORT_NAME = Sp.SUPPLIER_CODE
         NAME.Flag = "A"
      Else
         Set NAME = CstName.NAME
         NAME.LONG_NAME = StingToVariable(TempAsc, OldTempAsc, LineStr)
         NAME.SHORT_NAME = Sp.SUPPLIER_CODE
         NAME.Flag = "E"
      End If
      
      Call glbDaily.AddEditSupplier(Sp, IsOK, False, glbErrorLog)
   End If
   
   ProcessLine1 = True
   
   Exit Function
ErrorHandler:
   ProcessLine1 = False
End Function
Private Sub ImportPartItem()
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
   'Call LoadBillingDocDistinctDocumentNo(Nothing, c_DocumentNos, FromDate, ToDate)
   
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
      
      If ProcessLine2(TempStr) Then
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
Private Function ProcessLine2(LineStr As String) As Boolean
On Error GoTo ErrorHandler
Dim EmpCode As String

Dim ChkUniguePartItem As CPartItem
Dim Sp As CPartItem

Dim TimeStamp As Date
Dim TimeStampStr As String
Dim TempAsc As Long
Dim OldTempAsc As Long
Dim firstDate As Date
Dim lastDate As Date
Dim I As Long
Dim ItemCount As Long
Dim IsOK As Boolean
Dim Key1 As String
Dim Key2 As String
Dim Key3 As String
Dim Key4 As String

   If Left(LineStr, 2) = "PI" Then
               
      Set Sp = New CPartItem
      
      TempAsc = 3
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Set ChkUniguePartItem = GetObject("CPartItem", PartColls, Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1), False)
      If ChkUniguePartItem Is Nothing Then
         Sp.AddEditMode = SHOW_ADD
      Else
         Sp.AddEditMode = SHOW_EDIT
         Sp.PART_ITEM_ID = ChkUniguePartItem.PART_ITEM_ID
'         Sp.QueryFlag = 1
'         If Not glbDaily.QueryPartItem(Sp, m_Rs, ItemCount, IsOK, glbErrorLog) Then
'            glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'            Call EnableForm(Me, True)
'            Exit Function
'         End If
      End If
      
      Sp.PART_NO = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)
      OldTempAsc = TempAsc
      
      Sp.UNIT_COUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Sp.MINIMUM_ALLOW = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Sp.MAXIMUM_ALLOW = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Sp.PART_TYPE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Sp.PIG_FLAG = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Sp.PART_DESC = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Sp.UNIT_WEIGHT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Sp.BARCODE_NO = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Sp.BILL_DESC = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Sp.WEIGHT_PER_PACK = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Sp.PARCEL_TYPE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Sp.CANCEL_FLAG = "N"
      
      Call glbDaily.AddEditPartItem(Sp, IsOK, False, glbErrorLog)
   End If
   
   ProcessLine2 = True
   
   Exit Function
ErrorHandler:
   ProcessLine2 = False
End Function
Private Sub ImportCnDnRt()
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
   
   'FromDate = DateSerial(Mid(FileName1, 10, 4), Mid(FileName1, 14, 2), Mid(FileName1, 16, 2))
   'ToDate = DateSerial(Mid(FileName1, 19, 4), Mid(FileName1, 23, 2), Mid(FileName1, 25, 2))
   Call LoadBillingDocDistinctDocumentNo(Nothing, c_DocumentNos, FromDate, ToDate)
   
   Call LoadSupplier(Nothing, SupColls, 2)
   Call LoadPartItem(Nothing, PartColls, , , , 2)
   Call LoadMaster(Nothing, CnDnRtColls, DRCR_REASON, , , 2)
   
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
      
      If ProcessLine3(TempStr) Then
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
Private Function ProcessLine3(LineStr As String) As Boolean
On Error GoTo ErrorHandler
Dim EmpCode As String

Dim ChkUnigueBillingDoc As CBillingDoc

Dim TimeStamp As Date
Dim TimeStampStr As String
Dim TempAsc As Long
Dim OldTempAsc As Long
Dim firstDate As Date
Dim lastDate As Date
Dim I As Long
Dim ItemCount As Long
Dim IsOK As Boolean
Dim Si As CReceiptItem
Dim Key1 As String
Dim Key2 As String
Dim Key3 As String
Dim Key4 As String

   If Left(LineStr, 2) = "BD" Then
      If CountBill > 0 Then
         If DocumentType = 110 Then   'ใบรับคืน
            Call glbDaily.RT2InventoryDoc(Bl, Ivd, 2, 110)
         End If
         
         Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog)
         
         Bl.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID
         
         
         Call glbDaily.AddEditBillingDoc(Bl, IsOK, False, glbErrorLog)
         
         Set Bl = New CBillingDoc
         Set Ivd = New CInventoryDoc
      
      End If
      CountBill = 1
      
      TempAsc = 3
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Set ChkUnigueBillingDoc = GetObject("CBillingDoc", c_DocumentNos, Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1), False)
      If ChkUnigueBillingDoc Is Nothing Then
         Bl.AddEditMode = SHOW_ADD
      Else
         Bl.AddEditMode = SHOW_EDIT
         
         Set Si = New CReceiptItem
         Si.BILLING_DOC_ID = ChkUnigueBillingDoc.BILLING_DOC_ID
         Si.INVENTORY_DOC_ID = ChkUnigueBillingDoc.INVENTORY_DOC_ID
         Bl.BILLING_DOC_ID = ChkUnigueBillingDoc.BILLING_DOC_ID
         Bl.INVENTORY_DOC_ID = ChkUnigueBillingDoc.INVENTORY_DOC_ID
         
         Call Si.DeleteFromBillInv
         Set Si = Nothing
      End If
      Bl.DOCUMENT_NO = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)
      OldTempAsc = TempAsc
      
      Bl.DOCUMENT_DATE = InternalDateToDate(StingToVariable(TempAsc, OldTempAsc, LineStr))
      Bl.DOCUMENT_TYPE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      DocumentType = Bl.DOCUMENT_TYPE
      Bl.DUE_DATE = InternalDateToDate(StingToVariable(TempAsc, OldTempAsc, LineStr))
      Bl.NOTE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Bl.BILLING_ADDRESS_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.BILLING_ADDRESS_ID = -1
      Bl.ENTERPRISE_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.ENTERPRISE_ADDRESS_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.TOTAL_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.VAT_PERCENT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Bl.VAT_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.WH_PERCENT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.WH_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.DISCOUNT_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.PAYMENT_TERM = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
'      Dim Emp As CEmployee
'      TempAsc = InStr(TempAsc + 1, LineStr, ";")
'      Set Emp = GetEmployee(EmpColls, Trim(Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)))
'      If Emp.EMP_CODE <> Trim(Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)) Then
'         Call MsgBox("ยังไม่มีรหัสพนักงาน " & Trim(Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)) & " นี้ในระบบ กรุณาใส่เพิ่มเติ่มในภายหลัง ", vbOKOnly, PROJECT_NAME)
'      End If
'      Bl.ACCEPT_BY = Emp.EMP_ID                                                     '13
'      OldTempAsc = TempAsc
      
      Bl.ACCEPT_BY = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
'      TempAsc = InStr(TempAsc + 1, LineStr, ";")
'      Set Emp = GetEmployee(EmpColls, Trim(Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)))
'      Bl.RECEIVE_BY = Emp.EMP_ID                                                     '14
'      OldTempAsc = TempAsc
      Bl.RECEIVE_BY = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Bl.EXCEPTION_FLAG = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.PAYEE_NAME = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.COMMIT_FLAG = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Dim Sp As CSupplier
      Bl.SUPPLIER_CODE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Set Sp = GetSupplier(SupColls, Trim(Bl.SUPPLIER_CODE))
      If Sp.SUPPLIER_CODE <> Bl.SUPPLIER_CODE Then
         glbErrorLog.SystemErrorMsg = "ยังไม่มีรหัสซัพพลายเออร์ " & Trim(Bl.SUPPLIER_CODE) & " นี้ในระบบ กรุณาใส่เพิ่มเติ่มในภายหลัง "
         glbErrorLog.ShowErrorLog (LOG_TO_FILE)
      End If
      Bl.SUPPLIER_ID = Sp.SUPPLIER_ID

      'Bl.SUPPLIER_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.RECEIPT_TYPE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.ACCOUNT_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.DEPOSIT_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.TOTAL_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Bl.APPROVE_DATE = InternalDateToDate(StingToVariable(TempAsc, OldTempAsc, LineStr))
      Bl.ESTIMATE_DATE = InternalDateToDate(StingToVariable(TempAsc, OldTempAsc, LineStr))
      
      Bl.RESOURCE_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.PAYMENT_TYPE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.BANK_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.BBRANCH_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.CHECK_NO = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Bl.CHECK_DATE = InternalDateToDate(StingToVariable(TempAsc, OldTempAsc, LineStr))
      Bl.VADILITY = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.DELIVERY = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.PAYMENT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.PAYMENT_DESC = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Bl.SHIPMENT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.REFER_INV = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.PACKING_OF = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.SHIPON_DATE = InternalDateToDate(StingToVariable(TempAsc, OldTempAsc, LineStr))
      Bl.REF = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Bl.Credit = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.AREA_CODE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.SHIP_NAME = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.SHIPPING_MARKS = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.CD_PERCENT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Bl.CD_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.PACKAGE_TYPE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.TEMP_DO_NO = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.PAID_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.DEBIT_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Bl.CREDIT_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.DO_TOTAL_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.REVENUE_TOTAL_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.PAYMENT_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.BANK_BRANCH_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Bl.BANK_NOTE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.TOTAL_RCP = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.RUNNING_NO = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.AGREEMENT_DATA = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.AGREEMENT_FINANCE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Bl.OLD_CREDIT_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.ENTRY_DATE = InternalDateToDate(StingToVariable(TempAsc, OldTempAsc, LineStr))
      Bl.EXIT_DATE = InternalDateToDate(StingToVariable(TempAsc, OldTempAsc, LineStr))
      
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      
         
      'FK = Bl.BILLING_DOC_ID
      
   End If
   
   
   If Left(LineStr, 3) = "RCP" Then
      Dim Ti As CReceiptItem
      Set Ti = New CReceiptItem
      Ti.Flag = "A"
      
      'TI.DO_ID = FK
      
      TempAsc = 4
      OldTempAsc = TempAsc

      Dim Mr As CMasterRef
      Ti.DRCR_REASON_NO = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Set Mr = GetMasterRef(CnDnRtColls, Trim(Ti.DRCR_REASON_NO))
      If Ti.DRCR_REASON_NO <> Mr.KEY_CODE Then
         glbErrorLog.SystemErrorMsg = "ยังไม่มี สาเหตุ " & Trim(Ti.DRCR_REASON_NO) & " นี้ในระบบ กรุณาใส่เพิ่มเติ่มในภายหลัง "
         glbErrorLog.ShowErrorLog (LOG_TO_FILE)
      End If
      Ti.DRCR_REASON = Mr.KEY_ID
      
      Dim Lc As CLocation
      Ti.LOCATION_NO = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Set Lc = GetLocation(LocationColls, Trim(Ti.LOCATION_NO))
      If Lc.LOCATION_NO <> Ti.LOCATION_NO Then
         glbErrorLog.SystemErrorMsg = "ยังไม่มีรหัสคลัง " & Trim(Ti.LOCATION_NO) & " นี้ในระบบ กรุณาใส่เพิ่มเติ่มในภายหลัง "
         glbErrorLog.ShowErrorLog (LOG_TO_FILE)
      End If
      Ti.LOCATION_ID = Lc.LOCATION_ID
      
      Dim Pi As CPartItem
      Ti.PART_NO = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Set Pi = GetPartItem(PartColls, Trim(Ti.PART_NO))
      If Ti.PART_NO <> Pi.PART_NO Then
         glbErrorLog.SystemErrorMsg = "ยังไม่มีรหัสสินค้า/วัตถุดิบ " & Trim(Ti.PART_NO) & " นี้ในระบบ กรุณาใส่เพิ่มเติ่มในภายหลัง "
         glbErrorLog.ShowErrorLog (LOG_TO_FILE)
      End If
      Ti.PART_ITEM_ID = Pi.PART_ITEM_ID
      
      Ti.RECEIPT_ITEM_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.DISCOUNT_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.DEPOSIT_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.PAID_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.DEBIT_CREDIT_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Ti.CASH_DISCOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.LINK_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.RETURN_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.RETURN_TOTAL_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Ti.RETURN_AVG_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.AVG_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.AVG_WEIGHT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.RETURN_DISCOUNT_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.CONFIG_CODE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Ti.DISPLAY_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.COUNTRY_CURRENCY1 = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.COUNTRY_CURRENCY2 = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.WEIGHT_PER_PACK = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.PACK_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Ti.PRICE_PER_PACK = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.DISCOUNT_PER_PACK = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.MANUAL_FLAG = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.MANUAL_CODE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.MANUAL_NAME = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Ti.RATE_TYPE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.TRANSFER_WAGE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.STD_TRANSFER_CHARGE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.ITEM_DESC = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Ti.BBRANCH_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
           
      
      Ti.BILL_NO = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Set ChkUnigueBillingDoc = GetObject("CBillingDoc", c_DocumentNos, Ti.BILL_NO, True)
      If Ti.BILL_NO <> ChkUnigueBillingDoc.DOCUMENT_NO Then
         glbErrorLog.SystemErrorMsg = "ยังไม่มีรายการอ้างอิง " & Trim(Ti.BILL_NO) & " นี้ในระบบ กรุณาใส่เพิ่มเติ่มในภายหลัง "
         glbErrorLog.ShowErrorLog (LOG_TO_FILE)
      End If
      Ti.DO_ID = ChkUnigueBillingDoc.BILLING_DOC_ID
      
      
      Call Bl.ReceiptItems.add(Ti)
      
   End If
   
   If CountDown = 0 Then
      If DocumentType = 110 Then   'ใบรับคืน
         Call glbDaily.RT2InventoryDoc(Bl, Ivd, 2, 110)
      End If
      
      Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog)
      
      Bl.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID

      Call glbDaily.AddEditBillingDoc(Bl, IsOK, False, glbErrorLog)
      
      Set Bl = New CBillingDoc
      Set Ivd = New CInventoryDoc
      
   End If
   
   ProcessLine3 = True
   
   Exit Function
ErrorHandler:
   ProcessLine3 = False
End Function
Private Sub ImportPo()
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
   Call LoadBillingDocDistinctDocumentNo(Nothing, c_DocumentNos, FromDate, ToDate)
   
   Call LoadSupplier(Nothing, SupColls, 2)
   Call LoadPartItem(Nothing, PartColls, , , , 2)
   
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
      
      If ProcessLine4(TempStr) Then
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
Private Function ProcessLine4(LineStr As String) As Boolean
On Error GoTo ErrorHandler
Dim EmpCode As String

Dim ChkUnigueBillingDoc As CBillingDoc

Dim TimeStamp As Date
Dim TimeStampStr As String
Dim TempAsc As Long
Dim OldTempAsc As Long
Dim firstDate As Date
Dim lastDate As Date
Dim I As Long
Dim ItemCount As Long
Dim IsOK As Boolean
Dim Si As CSupItem
Dim Key1 As String
Dim Key2 As String
Dim Key3 As String
Dim Key4 As String

   If Left(LineStr, 2) = "BD" Then
      If CountBill > 0 Then
'         If DocumentType = 100 Then   'ใบรับเข้าวัตถุดิบ
'            Call glbDaily.SUP2InventoryDoc(Bl, Ivd, 1)
'         ElseIf DocumentType = 101 Then   'ใบรับเข้าวัสดุอุปกรณ์
'            Call glbDaily.SUP2InventoryDoc(Bl, Ivd, 19)
'         ElseIf DocumentType = 102 Then   'ใบรับเข้าจ่ายออกวัสดุอุปกรณ์
'            Call glbDaily.SUP2InventoryDocEx(Bl, Ivd, 20)
'         ElseIf DocumentType = 103 Then   'ใบรับเข้าทั่วไป
'            Call glbDaily.SUP2InventoryDoc(Bl, Ivd, 23)
'         End If
'
'         Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog)
'
'         Bl.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID
         
         
         Call glbDaily.AddEditBillingDoc(Bl, IsOK, False, glbErrorLog)
         
         Set Bl = New CBillingDoc
 '        Set Ivd = New CInventoryDoc
      
      End If
      CountBill = 1
      
      TempAsc = 3
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Set ChkUnigueBillingDoc = GetObject("CBillingDoc", c_DocumentNos, Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1), False)
      If ChkUnigueBillingDoc Is Nothing Then
         Bl.AddEditMode = SHOW_ADD
      Else
         Bl.AddEditMode = SHOW_EDIT
         
         Set Si = New CSupItem
         Si.DO_ID = ChkUnigueBillingDoc.BILLING_DOC_ID
'         Si.INVENTORY_DOC_ID = ChkUnigueBillingDoc.INVENTORY_DOC_ID
         Bl.BILLING_DOC_ID = ChkUnigueBillingDoc.BILLING_DOC_ID
'         Bl.INVENTORY_DOC_ID = ChkUnigueBillingDoc.INVENTORY_DOC_ID
         
         Call Si.DeleteFromBillInv
         Set Si = Nothing
      End If
      Bl.DOCUMENT_NO = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)
      OldTempAsc = TempAsc
      
      Bl.DOCUMENT_DATE = InternalDateToDate(StingToVariable(TempAsc, OldTempAsc, LineStr))
      Bl.DOCUMENT_TYPE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      DocumentType = Bl.DOCUMENT_TYPE
      Bl.DUE_DATE = InternalDateToDate(StingToVariable(TempAsc, OldTempAsc, LineStr))
      Bl.NOTE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Bl.BILLING_ADDRESS_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.BILLING_ADDRESS_ID = -1
      Bl.ENTERPRISE_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.ENTERPRISE_ADDRESS_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.TOTAL_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.VAT_PERCENT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Bl.VAT_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.WH_PERCENT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.WH_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.DISCOUNT_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.PAYMENT_TERM = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
'      Dim Emp As CEmployee
'      TempAsc = InStr(TempAsc + 1, LineStr, ";")
'      Set Emp = GetEmployee(EmpColls, Trim(Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)))
'      If Emp.EMP_CODE <> Trim(Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)) Then
'         Call MsgBox("ยังไม่มีรหัสพนักงาน " & Trim(Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)) & " นี้ในระบบ กรุณาใส่เพิ่มเติ่มในภายหลัง ", vbOKOnly, PROJECT_NAME)
'      End If
'      Bl.ACCEPT_BY = Emp.EMP_ID                                                     '13
'      OldTempAsc = TempAsc
      
      Bl.ACCEPT_BY = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
'      TempAsc = InStr(TempAsc + 1, LineStr, ";")
'      Set Emp = GetEmployee(EmpColls, Trim(Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)))
'      Bl.RECEIVE_BY = Emp.EMP_ID                                                     '14
'      OldTempAsc = TempAsc
      Bl.RECEIVE_BY = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Bl.EXCEPTION_FLAG = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.PAYEE_NAME = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.COMMIT_FLAG = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Dim Sp As CSupplier
      Bl.SUPPLIER_CODE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Set Sp = GetSupplier(SupColls, Trim(Bl.SUPPLIER_CODE))
      If Sp.SUPPLIER_CODE <> Bl.SUPPLIER_CODE Then
         glbErrorLog.SystemErrorMsg = "ยังไม่มีรหัสซัพพลายเออร์ " & Trim(Bl.SUPPLIER_CODE) & " นี้ในระบบ กรุณาใส่เพิ่มเติ่มในภายหลัง "
         glbErrorLog.ShowErrorLog (LOG_TO_FILE)
      End If
      Bl.SUPPLIER_ID = Sp.SUPPLIER_ID

      'Bl.SUPPLIER_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.RECEIPT_TYPE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.ACCOUNT_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.DEPOSIT_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.TOTAL_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Bl.APPROVE_DATE = InternalDateToDate(StingToVariable(TempAsc, OldTempAsc, LineStr))
      Bl.ESTIMATE_DATE = InternalDateToDate(StingToVariable(TempAsc, OldTempAsc, LineStr))
      
      Bl.RESOURCE_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.PAYMENT_TYPE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.BANK_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.BBRANCH_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.CHECK_NO = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Bl.CHECK_DATE = InternalDateToDate(StingToVariable(TempAsc, OldTempAsc, LineStr))
      Bl.VADILITY = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.DELIVERY = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.PAYMENT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.PAYMENT_DESC = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Bl.SHIPMENT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.REFER_INV = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.PACKING_OF = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.SHIPON_DATE = InternalDateToDate(StingToVariable(TempAsc, OldTempAsc, LineStr))
      Bl.REF = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Bl.Credit = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.AREA_CODE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.SHIP_NAME = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.SHIPPING_MARKS = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.CD_PERCENT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Bl.CD_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.PACKAGE_TYPE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.TEMP_DO_NO = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.PAID_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.DEBIT_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Bl.CREDIT_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.DO_TOTAL_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.REVENUE_TOTAL_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.PAYMENT_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.BANK_BRANCH_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Bl.BANK_NOTE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.TOTAL_RCP = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.RUNNING_NO = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.AGREEMENT_DATA = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.AGREEMENT_FINANCE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Bl.OLD_CREDIT_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.ENTRY_DATE = InternalDateToDate(StingToVariable(TempAsc, OldTempAsc, LineStr))
      Bl.EXIT_DATE = InternalDateToDate(StingToVariable(TempAsc, OldTempAsc, LineStr))
      Bl.DO_NO = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.TRUCK_NO = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Bl.DELIVERY_FEE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.SENDER_NAME = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.RECEIVE_NAME = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.DEPARTMENT_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.QUE_NO = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Bl.PR_NO = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      
         
      'FK = Bl.BILLING_DOC_ID
      
   End If
   
   
   If Left(LineStr, 2) = "SI" Then
      Dim Ti As CSupItem
      Set Ti = New CSupItem
      Ti.Flag = "A"
      
      'TI.DO_ID = FK
      
      TempAsc = 3
      OldTempAsc = TempAsc

      Ti.DO_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Dim Pi As CPartItem
      Ti.PART_NO = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Set Pi = GetPartItem(PartColls, Trim(Ti.PART_NO))
      If Ti.PART_NO <> Pi.PART_NO Then
         glbErrorLog.SystemErrorMsg = "ยังไม่มีรหัสสินค้า/วัตถุดิบ " & Trim(Ti.PART_NO) & " นี้ในระบบ กรุณาใส่เพิ่มเติ่มในภายหลัง "
         glbErrorLog.ShowErrorLog (LOG_TO_FILE)
      End If
      Ti.PART_ITEM_ID = Pi.PART_ITEM_ID
      
      'Ti.PART_ITEM_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Dim Lc As CLocation
      Ti.LOCATION_NO = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Set Lc = GetLocation(LocationColls, Trim(Ti.LOCATION_NO))
      If Lc.LOCATION_NO <> Ti.LOCATION_NO Then
         glbErrorLog.SystemErrorMsg = "ยังไม่มีรหัสคลัง " & Trim(Ti.LOCATION_NO) & " นี้ในระบบ กรุณาใส่เพิ่มเติ่มในภายหลัง "
         glbErrorLog.ShowErrorLog (LOG_TO_FILE)
      End If
      Ti.LOCATION_ID = Lc.LOCATION_ID
      
      'Ti.LOCATION_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.ACTUAL_UNIT_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.INCLUDE_UNIT_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.PREVIOUS_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Ti.PREVIOUS_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.NEW_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.TX_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.NEW_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.TRANSACTION_SEQ = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Ti.GUI_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.CALCULATE_FLAG = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.TOTAL_ACTUAL_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.TOTAL_INCLUDE_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.TX_TYPE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Ti.LEFT_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.LAYOUT_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.LINK_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.PACKAGING_AMT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.ENTRY_WEIGHT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Ti.EXIT_WEIGHT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.WEIGHT_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.PACKAGE_WEIGHT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.OTHER_WEIGHT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.PERCENT_HUMID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Ti.HUMID_WEIGHT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.PACKAGING_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.SUPPLIER_WEIGHT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.CALCULATE_TYPE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.PACKAGE_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Ti.ACTUAL_PKG_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.PUREXP_ID1 = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.PUREXP_ID2 = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.EXPENSE1 = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.EXPENSE2 = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Ti.TOTAL_WEIGHT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.RAW_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.DISCOUNT_AMT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.RAW_TOT_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.TO_DEPARTMENT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Ti.ITEM_DESC = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.EXTRA_NAME = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.SALE_TOT_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.CALCULATE_WEIGHT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.RAW_COST = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Ti.EXPENSE_COST = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.TOTAL_NEW_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.BAG_RETURN = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.ACTUAL_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.ACTUAL_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Ti.CURRENT_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.CURRENT_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.NEED_TOTAL_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.NEED_TOTAL_AMOUNT = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.NEED_AVG_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
      Ti.MANUAL_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.EXPENSE_TYPE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.AUTO_PRICE = StingToVariable(TempAsc, OldTempAsc, LineStr)
      Ti.ITEM_DESC_ID = StingToVariable(TempAsc, OldTempAsc, LineStr)
      
              
      Call Bl.SupItems.add(Ti)
      
   End If
   
   If CountDown = 0 Then
'      If DocumentType = 100 Then   'ใบรับเข้าวัตถุดิบ
'         Call glbDaily.SUP2InventoryDoc(Bl, Ivd, 1)
'      ElseIf DocumentType = 101 Then   'ใบรับเข้าวัสดุอุปกรณ์
'         Call glbDaily.SUP2InventoryDoc(Bl, Ivd, 19)
'      ElseIf DocumentType = 102 Then   'ใบรับเข้าจ่ายออกวัสดุอุปกรณ์
'         Call glbDaily.SUP2InventoryDocEx(Bl, Ivd, 20)
'      ElseIf DocumentType = 103 Then   'ใบรับเข้าทั่วไป
'         Call glbDaily.SUP2InventoryDoc(Bl, Ivd, 23)
'      End If
'
'      Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog)
'
'      Bl.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID

      
      Call glbDaily.AddEditBillingDoc(Bl, IsOK, False, glbErrorLog)
      
      Set Bl = New CBillingDoc
'      Set Ivd = New CInventoryDoc
      
   End If
   
   ProcessLine4 = True
   
   Exit Function
ErrorHandler:
   ProcessLine4 = False
End Function

