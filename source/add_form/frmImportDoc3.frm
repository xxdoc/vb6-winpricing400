VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmImportDoc3 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Icon            =   "frmImportDoc3.frx":0000
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
         TabIndex        =   13
         Top             =   1020
         Width           =   3105
      End
      Begin prjFarmManagement.uctlTextBox txtFileName 
         Height          =   435
         Left            =   1860
         TabIndex        =   12
         Top             =   1470
         Width           =   7305
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
         MouseIcon       =   "frmImportDoc3.frx":27A2
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
         MouseIcon       =   "frmImportDoc3.frx":2ABC
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmImportDoc3"
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
Private m_InventoryDocsDateLocations As Collection
Private m_ExcelApp As Object
Private m_ExcelSheet As Object

Private m_InventoryDocs As Collection
Private m_InventoryDocExts As Collection

Private m_BillingDocs As Collection
Private m_PartItems As Collection
Private m_Customers As Collection
Private m_CustomerExts As Collection
Private m_Suppliers As Collection
Private m_PartItemExts As Collection
Private m_SupplierExts As Collection
Private m_SheetBalances As Collection
Private m_Formulas As Collection
Private m_Jobs As Collection
Private m_Cheques As Collection
Private m_CashDocs As Collection
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
   
   TempID = cboImportType.ItemData(Minus2Zero(cboImportType.ListIndex))
   If TempID <= 0 Then
      Exit Sub
   End If
   
   If TempID = 1 Then
      Call PatchPartItem
   ElseIf TempID = 2 Then
      Call PatchSupplier
   ElseIf TempID = 4 Then
      Call PatchInventoryDoc
   ElseIf TempID = 3 Then
      Call PatchCustomer
   ElseIf TempID = 5 Then
      Call PatchCheque
   ElseIf TempID = 6 Then
      Call PatchJob
   ElseIf TempID = 7 Then
      Call PatchFormula
   ElseIf TempID = 8 Then
      Call PatchCashDoc
   ElseIf TempID = 9 Then
      Call PatchBillingDoc
   End If
End Sub

Private Function IsExist(Key As String, Col As Collection) As Boolean
Dim Obj As Object

   Set Obj = GetInventoryDoc(Col, Key)
   
   IsExist = Not (Obj Is Nothing)
End Function

Private Sub PatchInventoryDoc()
Dim IvdE As CInventoryDocExt
Dim SpE As CSupplierExt
Dim Sp As CSupplier
Dim PiE As CPartItemExt
Dim Pi As CPartItem
Dim LID As CLotItemDup
Dim TempRs1 As ADODB.Recordset
Dim TempRs2 As ADODB.Recordset
Dim iCount As Long
Dim ItemCount As Long
Dim I As Long
Dim Percent As Double
Dim OldID As Long

   Set TempRs1 = New ADODB.Recordset
   Set TempRs2 = New ADODB.Recordset
   Set IvdE = New CInventoryDocExt
   
   txtFileName.Text = "INVENTORY_DOC"
   Call IvdE.QueryData(TempRs1, iCount)
   
   glbDaily.StartTransaction
   I = 0
   prgProgress.Value = 0
   prgProgress.MAX = 100
   
   While Not TempRs1.EOF
      I = I + 1
      Percent = MyDiffEx(I, iCount) * 100
      
      Call IvdE.PopulateFromRS(1, TempRs1)
      If Not IsExist(IvdE.DOCUMENT_NO, m_InventoryDocs) Then
         OldID = IvdE.INVENTORY_DOC_ID
         txtFileName.Text = "INVENTORY_DOC - " & IvdE.DOCUMENT_NO
         
         If IvdE.SUPPLIER_ID > 0 Then
            Set SpE = GetSupplierExt(m_SupplierExts, Trim(str(IvdE.SUPPLIER_ID)))
            Set Sp = GetSupplier(m_Suppliers, SpE.SUPPLIER_CODE)
            IvdE.SUPPLIER_ID = Sp.SUPPLIER_ID
         Else
            IvdE.SUPPLIER_ID = -1
         End If
                  
         IvdE.AddEditMode = SHOW_ADD
         Call IvdE.AddEditData
         
         Set LID = New CLotItemDup
         LID.INVENTORY_DOC_ID = OldID
         Call LID.QueryData(1, TempRs2, ItemCount)
         While Not TempRs2.EOF
            Call LID.PopulateFromRS(1, TempRs2)
            
            Set PiE = GetPartItemExt(m_PartItemExts, Trim(str(LID.PART_ITEM_ID)))
            Set Pi = GetPartItem(m_PartItems, PiE.PART_NO)
            
            LID.AddEditMode = SHOW_ADD
            LID.PART_ITEM_ID = Pi.PART_ITEM_ID
            LID.INVENTORY_DOC_ID = IvdE.INVENTORY_DOC_ID
            Call LID.AddEditData
            
            TempRs2.MoveNext
         Wend
         Set LID = Nothing
      End If
      
      TempRs1.MoveNext
      
      txtPercent.Text = FormatNumber(Percent)
      prgProgress.Value = Percent
      Me.Refresh
   Wend
   
   prgProgress.Value = prgProgress.MAX
   
   glbDaily.CommitTransaction
   
   Set IvdE = Nothing
   If TempRs1.State = adStateOpen Then
      TempRs1.Close
   End If
   Set TempRs1 = Nothing
   
   If TempRs2.State = adStateOpen Then
      TempRs2.Close
   End If
   Set TempRs2 = Nothing
End Sub

Private Sub PatchFormula()
Dim IvdE As CFormulaExt
Dim SpE As CSupplierExt
Dim Sp As CSupplier
Dim PiE As CPartItemExt
Dim Pi As CPartItem
Dim LID As CFormulaItemDup
Dim TempRs1 As ADODB.Recordset
Dim TempRs2 As ADODB.Recordset
Dim iCount As Long
Dim ItemCount As Long
Dim I As Long
Dim Percent As Double
Dim OldID As Long

   Set TempRs1 = New ADODB.Recordset
   Set TempRs2 = New ADODB.Recordset
   Set IvdE = New CFormulaExt
   
   txtFileName.Text = "FORMULA"
   Call IvdE.QueryData(1, TempRs1, iCount)
   
   glbDaily.StartTransaction
   I = 0
   prgProgress.Value = 0
   prgProgress.MAX = 100
   
   While Not TempRs1.EOF
      I = I + 1
      Percent = MyDiffEx(I, iCount) * 100
      
      Call IvdE.PopulateFromRS(1, TempRs1)
      If Not IsExist(IvdE.FORMULA_NO, m_Formulas) Then
         OldID = IvdE.FORMULA_ID
         txtFileName.Text = "FORMULA - " & IvdE.FORMULA_NO
         
         Set PiE = GetPartItemExt(m_PartItemExts, Trim(str(IvdE.PART_ITEM_ID)))
         Set Pi = GetPartItem(m_PartItems, PiE.PART_NO)
         
         IvdE.PART_ITEM_ID = Pi.PART_ITEM_ID
         IvdE.AddEditMode = SHOW_ADD
         Call IvdE.AddEditData
         
         Set LID = New CFormulaItemDup
         LID.FORMULA_ID = OldID
         Call LID.QueryData(1, TempRs2, ItemCount)
         While Not TempRs2.EOF
            Call LID.PopulateFromRS(1, TempRs2)
            
            Set PiE = GetPartItemExt(m_PartItemExts, Trim(str(LID.PART_ITEM_ID)))
            Set Pi = GetPartItem(m_PartItems, PiE.PART_NO)
            
            LID.AddEditMode = SHOW_ADD
            LID.PART_ITEM_ID = Pi.PART_ITEM_ID
            LID.FORMULA_ID = IvdE.FORMULA_ID
            Call LID.AddEditData
            
            TempRs2.MoveNext
         Wend
         Set LID = Nothing
      Else

      End If
      
      TempRs1.MoveNext
      
      txtPercent.Text = FormatNumber(Percent)
      prgProgress.Value = Percent
      Me.Refresh
   Wend
   
   prgProgress.Value = prgProgress.MAX
   
   glbDaily.CommitTransaction
   
   Set IvdE = Nothing
   If TempRs1.State = adStateOpen Then
      TempRs1.Close
   End If
   Set TempRs1 = Nothing
   
   If TempRs2.State = adStateOpen Then
      TempRs2.Close
   End If
   Set TempRs2 = Nothing
End Sub
Private Sub PatchCashDoc()
Dim IvdE As CCashDocExt
Dim SpE As CSupplierExt
Dim Sp As CSupplier
Dim PiE As CPartItemExt
Dim Pi As CPartItem
Dim LID As CFormulaItemDup
Dim TempRs1 As ADODB.Recordset
Dim TempRs2 As ADODB.Recordset
Dim iCount As Long
Dim ItemCount As Long
Dim I As Long
Dim Percent As Double
Dim OldID As Long

   Set TempRs1 = New ADODB.Recordset
   Set TempRs2 = New ADODB.Recordset
   Set IvdE = New CCashDocExt
   
   txtFileName.Text = "CASH_DOC"
   Call IvdE.QueryData(1, TempRs1, iCount)
   
   glbDaily.StartTransaction
   I = 0
   prgProgress.Value = 0
   prgProgress.MAX = 100
   
   While Not TempRs1.EOF
      I = I + 1
      Percent = MyDiffEx(I, iCount) * 100
      
      Call IvdE.PopulateFromRS(1, TempRs1)
      If Not IsExist(IvdE.GetFieldValue("DOCUMENT_NO"), m_CashDocs) Then
         OldID = IvdE.GetFieldValue("CASH_DOC_ID")
         txtFileName.Text = "CASH_DOC - " & IvdE.GetFieldValue("DOCUMENT_NO")
   
         Call IvdE.SetFieldValue("CUSTOMER_ID", -1)
         IvdE.ShowMode = SHOW_ADD
         Call IvdE.AddEditData
         
'         Set LID = New CFormulaItemDup
'         LID.FORMULA_ID = OldID
'         Call LID.QueryData(1, TempRs2, ItemCount)
'         While Not TempRs2.EOF
'            Call LID.PopulateFromRS(1, TempRs2)
'
'            Set PiE = GetPartItemExt(m_PartItemExts, Trim(Str(LID.PART_ITEM_ID)))
'            Set Pi = GetPartItem(m_PartItems, PiE.PART_NO)
'
'            LID.AddEditMode = SHOW_ADD
'            LID.PART_ITEM_ID = Pi.PART_ITEM_ID
'            LID.FORMULA_ID = IvdE.FORMULA_ID
'            Call LID.AddEditData
'
'            TempRs2.MoveNext
'         Wend
'         Set LID = Nothing
      Else

      End If
      
      TempRs1.MoveNext
      
      txtPercent.Text = FormatNumber(Percent)
      prgProgress.Value = Percent
      Me.Refresh
   Wend
   
   prgProgress.Value = prgProgress.MAX
   
   glbDaily.CommitTransaction
   
   Set IvdE = Nothing
   If TempRs1.State = adStateOpen Then
      TempRs1.Close
   End If
   Set TempRs1 = Nothing
   
   If TempRs2.State = adStateOpen Then
      TempRs2.Close
   End If
   Set TempRs2 = Nothing
End Sub

Private Sub PatchJob()
Dim Ivd As CInventoryDoc
Dim IvdE As CJobExt
Dim SpE As CSupplierExt
Dim Sp As CSupplier
Dim PiE As CPartItemExt
Dim Pi As CPartItem
Dim LID As CJobInOutDup
Dim JpD As CJobParameterDup
Dim TempRs1 As ADODB.Recordset
Dim TempRs2 As ADODB.Recordset
Dim iCount As Long
Dim ItemCount As Long
Dim I As Long
Dim Percent As Double
Dim OldID As Long

   Set TempRs1 = New ADODB.Recordset
   Set TempRs2 = New ADODB.Recordset
   Set IvdE = New CJobExt
   
   txtFileName.Text = "JOB"
   Call IvdE.QueryData(1, TempRs1, iCount)
   
   glbDaily.StartTransaction
   I = 0
   prgProgress.Value = 0
   prgProgress.MAX = 100
   
   While Not TempRs1.EOF
      I = I + 1
      Percent = MyDiffEx(I, iCount) * 100
      
      Call IvdE.PopulateFromRS(1, TempRs1)
      If Not IsExist(IvdE.JOB_NO, m_Jobs) Then
         OldID = IvdE.JOB_ID
         txtFileName.Text = "JOB - " & IvdE.JOB_NO
         
         Set PiE = GetPartItemExt(m_PartItemExts, Trim(str(IvdE.PART_ITEM_ID)))
         Set Pi = GetPartItem(m_PartItems, PiE.PART_NO)
         Set Ivd = GetInventoryDoc(m_InventoryDocs, IvdE.JOB_NO)
         
         IvdE.PART_ITEM_ID = Pi.PART_ITEM_ID
         IvdE.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID
         IvdE.FORMULA_ID = -1
         IvdE.AddEditMode = SHOW_ADD
         Call IvdE.AddEditData
         
         '=====
         Set LID = New CJobInOutDup
         LID.JOB_ID = OldID
         Call LID.QueryData(1, TempRs2, ItemCount)
         While Not TempRs2.EOF
            Call LID.PopulateFromRS(1, TempRs2)
            
            Set PiE = GetPartItemExt(m_PartItemExts, Trim(str(LID.PART_ITEM_ID)))
            Set Pi = GetPartItem(m_PartItems, PiE.PART_NO)
            
            LID.AddEditMode = SHOW_ADD
            LID.PART_ITEM_ID = Pi.PART_ITEM_ID
            LID.JOB_ID = IvdE.JOB_ID
            Call LID.AddEditData
            
            TempRs2.MoveNext
         Wend
         Set LID = Nothing
      
         '=====
         Set JpD = New CJobParameterDup
         JpD.JOB_ID = OldID
         Call JpD.QueryData(1, TempRs2, ItemCount)
         While Not TempRs2.EOF
            Call JpD.PopulateFromRS(1, TempRs2)
            
            JpD.AddEditMode = SHOW_ADD
            JpD.JOB_ID = IvdE.JOB_ID
            Call JpD.AddEditData
            
            TempRs2.MoveNext
         Wend
         Set JpD = Nothing
      Else
'Debug.Print
      End If
      
      TempRs1.MoveNext
      
      txtPercent.Text = FormatNumber(Percent)
      prgProgress.Value = Percent
      Me.Refresh
   Wend
   
   prgProgress.Value = prgProgress.MAX
   
   glbDaily.CommitTransaction
   
   Set IvdE = Nothing
   If TempRs1.State = adStateOpen Then
      TempRs1.Close
   End If
   Set TempRs1 = Nothing
   
   If TempRs2.State = adStateOpen Then
      TempRs2.Close
   End If
   Set TempRs2 = Nothing
End Sub

Private Sub PatchPartItem()
Dim PiE As CPartItemExt
Dim Pi As CPartItem
Dim TempRs1 As ADODB.Recordset
Dim TempRs2 As ADODB.Recordset
Dim iCount As Long
Dim ItemCount As Long
Dim I As Long
Dim Percent As Double
Dim OldID As Long

   Set TempRs1 = New ADODB.Recordset
   Set TempRs2 = New ADODB.Recordset
   Set PiE = New CPartItemExt
   
   txtFileName.Text = "PART_ITEM"
   Call PiE.QueryData(1, TempRs1, iCount)
   
   glbDaily.StartTransaction
   I = 0
   prgProgress.Value = 0
   prgProgress.MAX = 100
   
   While Not TempRs1.EOF
      I = I + 1
      Percent = MyDiffEx(I, iCount) * 100
      
      Call PiE.PopulateFromRS(1, TempRs1)
      If Not IsExist(PiE.PART_NO, m_PartItems) Then
         OldID = PiE.PART_ITEM_ID
         txtFileName.Text = "PART_ITEM  -  " & PiE.PART_NO
                 
         PiE.OLD_PART_ID = OldID
         PiE.AddEditMode = SHOW_ADD
         PiE.PIG_FLAG = "N"
         Call PiE.AddEditData
      Else
'Debug.Print
      End If
      
      TempRs1.MoveNext
      
      txtPercent.Text = FormatNumber(Percent)
      prgProgress.Value = Percent
      Me.Refresh
   Wend
   
   prgProgress.Value = prgProgress.MAX
   
   glbDaily.CommitTransaction
   
   Set PiE = Nothing
   If TempRs1.State = adStateOpen Then
      TempRs1.Close
   End If
   Set TempRs1 = Nothing
   
   If TempRs2.State = adStateOpen Then
      TempRs2.Close
   End If
   Set TempRs2 = Nothing
End Sub

Private Sub PatchCheque()
Dim PiE As CChequeExt
Dim Pi As CCheque
Dim TempRs1 As ADODB.Recordset
Dim TempRs2 As ADODB.Recordset
Dim iCount As Long
Dim ItemCount As Long
Dim I As Long
Dim Percent As Double
Dim OldID As Long

   Set TempRs1 = New ADODB.Recordset
   Set TempRs2 = New ADODB.Recordset
   Set PiE = New CChequeExt
   
   txtFileName.Text = "CCHEQUE"
   Call PiE.QueryData(1, TempRs1, iCount)
   
   glbDaily.StartTransaction
   I = 0
   prgProgress.Value = 0
   prgProgress.MAX = 100
   
   While Not TempRs1.EOF
      I = I + 1
      Percent = MyDiffEx(I, iCount) * 100
      
      Call PiE.PopulateFromRS(1, TempRs1)
      If Not IsExist(PiE.GetFieldValue("CHEQUE_NO"), m_Cheques) Then
         OldID = PiE.GetFieldValue("CHEQUE_ID")
         txtFileName.Text = "CHEQUE  -  " & PiE.GetFieldValue("CHEQUE_NO")
                 
         PiE.ShowMode = SHOW_ADD
         Call PiE.AddEditData
      Else
''Debug.Print
      End If
      
      TempRs1.MoveNext
      
      txtPercent.Text = FormatNumber(Percent)
      prgProgress.Value = Percent
      Me.Refresh
   Wend
   
   prgProgress.Value = prgProgress.MAX
   
   glbDaily.CommitTransaction
   
   Set PiE = Nothing
   If TempRs1.State = adStateOpen Then
      TempRs1.Close
   End If
   Set TempRs1 = Nothing
   
   If TempRs2.State = adStateOpen Then
      TempRs2.Close
   End If
   Set TempRs2 = Nothing
End Sub

Private Sub PatchSupplier()
Dim PiE As CSupplierExt
Dim Pi As CSupplier
Dim TempRs1 As ADODB.Recordset
Dim TempRs2 As ADODB.Recordset
Dim iCount As Long
Dim ItemCount As Long
Dim I As Long
Dim Percent As Double
Dim OldID As Long

   Set TempRs1 = New ADODB.Recordset
   Set TempRs2 = New ADODB.Recordset
   Set PiE = New CSupplierExt
   
   txtFileName.Text = "SUPPLIER"
   PiE.SUPPLIER_ID = -1
   Call PiE.QueryData(TempRs1, iCount)
   
   glbDaily.StartTransaction
   I = 0
   prgProgress.Value = 0
   prgProgress.MAX = 100
   
   While Not TempRs1.EOF
      I = I + 1
      Percent = MyDiffEx(I, iCount) * 100
      
      Call PiE.PopulateFromRS(1, TempRs1)
      If Not IsExist(PiE.SUPPLIER_CODE, m_Suppliers) Then
         txtFileName.Text = "SUPPLIER  -  " & PiE.SUPPLIER_CODE
                 
         PiE.AddEditMode = SHOW_ADD
         Call PiE.AddEditData
      Else
'Debug.Print
      End If
      
      TempRs1.MoveNext
      
      txtPercent.Text = FormatNumber(Percent)
      prgProgress.Value = Percent
      Me.Refresh
   Wend
   
   prgProgress.Value = prgProgress.MAX
   
   glbDaily.CommitTransaction
   
   Set PiE = Nothing
   If TempRs1.State = adStateOpen Then
      TempRs1.Close
   End If
   Set TempRs1 = Nothing
   
   If TempRs2.State = adStateOpen Then
      TempRs2.Close
   End If
   Set TempRs2 = Nothing
End Sub

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
      
      Call InitPatchTable(cboImportType)
      Call LoadInventoryDoc(Nothing, m_InventoryDocs, InternalDateToDate("2007-05-01 00:00:00"), InternalDateToDate("2007-06-30 23:59:59"))
      Call LoadInventoryDocExt(Nothing, m_InventoryDocExts, InternalDateToDate("2007-05-01 00:00:00"), InternalDateToDate("2007-06-30 23:59:59"))
      
      Call LoadBillingDoc(Nothing, m_BillingDocs)
      
      Call LoadFormula(Nothing, m_Formulas, , , InternalDateToDate("2007-05-01 00:00:00"), -1, 2)
      Call LoadJob(Nothing, m_Jobs, InternalDateToDate("2007-05-01 00:00:00"), -1, 2)
'      Call LoadCashDoc(Nothing, m_CashDocs, -1, -1, 2)
      Call LoadCheque(Nothing, m_Cheques)
      Call LoadSupplier(Nothing, m_Suppliers, 2)
      Call LoadSupplierExt(Nothing, m_SupplierExts, 1)
      Call LoadPartItem(Nothing, m_PartItems, , , , 2)
      Call LoadPartItemExt(Nothing, m_PartItemExts, , , , 1)
      
      Call LoadCustomer(Nothing, m_Customers, 2)
      Call LoadCustomerExt(Nothing, m_CustomerExts)
      
      
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
   pnlHeader.Caption = "แพตข้อมูล"
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitNormalLabel(lblFileName, "ตาราง")
   Call InitNormalLabel(lblMasterName, "รายละเอียด")
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
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdStart, MapText("เริ่ม"))
   
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
   Set m_InventoryDocsDateLocations = New Collection
   Set m_InventoryDocs = New Collection
   Set m_InventoryDocExts = New Collection
   
   Set m_PartItems = New Collection
   Set m_Customers = New Collection
   Set m_CustomerExts = New Collection
   Set m_Suppliers = New Collection
   Set m_PartItemExts = New Collection
   Set m_SheetBalances = New Collection
   Set m_Suppliers = New Collection
   Set m_SupplierExts = New Collection
   Set m_Formulas = New Collection
   Set m_Jobs = New Collection
   Set m_Cheques = New Collection
   Set m_CashDocs = New Collection
   Set m_BillingDocs = New Collection
   
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
   Set m_InventoryDocsDateLocations = Nothing
   Set m_InventoryDocs = Nothing
   Set m_InventoryDocExts = Nothing
   Set m_PartItems = Nothing
   Set m_Customers = Nothing
   Set m_CustomerExts = Nothing
   Set m_Suppliers = Nothing
   Set m_SheetBalances = Nothing
   Set m_PartItemExts = Nothing
   Set m_Suppliers = Nothing
   Set m_SupplierExts = Nothing
   Set m_Formulas = Nothing
   Set m_Jobs = Nothing
   Set m_Cheques = Nothing
   Set m_CashDocs = Nothing
   Set m_BillingDocs = Nothing
End Sub
Private Sub PatchCustomer()
Dim PiE As CCustomerExt
Dim Pi As CCustomer
Dim TempRs1 As ADODB.Recordset
Dim TempRs2 As ADODB.Recordset
Dim iCount As Long
Dim ItemCount As Long
Dim I As Long
Dim Percent As Double
Dim OldID As Long

   Set TempRs1 = New ADODB.Recordset
   Set TempRs2 = New ADODB.Recordset
   Set PiE = New CCustomerExt
   
   txtFileName.Text = "CUSTOMER"
   PiE.CUSTOMER_ID = -1
   Call PiE.QueryData(TempRs1, iCount)
   
   glbDaily.StartTransaction
   I = 0
   prgProgress.Value = 0
   prgProgress.MAX = 100
   
   While Not TempRs1.EOF
      I = I + 1
      Percent = MyDiffEx(I, iCount) * 100
      
      Call PiE.PopulateFromRS(1, TempRs1)
      If Not IsExist(PiE.CUSTOMER_CODE, m_Customers) Then
         txtFileName.Text = "CUSTOMER  -  " & PiE.CUSTOMER_CODE
         
         PiE.AddEditMode = SHOW_ADD
         Call PiE.AddEditData
      Else
'Debug.Print
      End If
      
      TempRs1.MoveNext
      
      txtPercent.Text = FormatNumber(Percent)
      prgProgress.Value = Percent
      Me.Refresh
   Wend
   
   prgProgress.Value = prgProgress.MAX
   
   glbDaily.CommitTransaction
   
   Set PiE = Nothing
   If TempRs1.State = adStateOpen Then
      TempRs1.Close
   End If
   Set TempRs1 = Nothing
   
   If TempRs2.State = adStateOpen Then
      TempRs2.Close
   End If
   Set TempRs2 = Nothing
End Sub
Private Sub PatchBillingDoc()
Dim TempRs1 As ADODB.Recordset
Dim TempRs2 As ADODB.Recordset
Dim iCount As Long
Dim ItemCount As Long
Dim I As Long
Dim Percent As Double
Dim OldID As Long

Dim BD As CBillingDoc
Dim BdE As CBillingDocDup

Dim SpE As CSupplierExt
Dim Sp As CSupplier

Dim IvdE As CInventoryDocExt
Dim Ivd As CInventoryDoc

Dim PiE As CPartItemExt
Dim Pi As CPartItem

'---------------------------------------------------------------------------------------------------------- Table ลูก
Dim Bdd As CBillingDiscountDup
Dim BuH As CBulkHoleDup
Dim Sup As CSupItemDup
Dim Rec As CReceiptItemDup
Dim ct As CCashTranDup
Dim Doc As CDoItemDup
'---------------------------------------------------------------------------------------------------------- Table ลูก
   
   Set TempRs1 = New ADODB.Recordset
   Set TempRs2 = New ADODB.Recordset
   Set BdE = New CBillingDocDup
   
   txtFileName.Text = "BILLING_DOC"
   BdE.FROM_DATE = DateSerial(2007, 5, 1)
   BdE.TO_DATE = DateSerial(2007, 6, 30)
   Call BdE.QueryData(1, TempRs1, iCount)
   
   glbDaily.StartTransaction
   I = 0
   prgProgress.Value = 0
   prgProgress.MAX = 100
   
   While Not TempRs1.EOF
      I = I + 1
      Percent = MyDiffEx(I, iCount) * 100
      
      Call BdE.PopulateFromRS(1, TempRs1)
      If Not IsExist(BdE.DOCUMENT_NO, m_BillingDocs) Then
         OldID = BdE.BILLING_DOC_ID
         txtFileName.Text = "BILLING_DOC - " & BdE.BILLING_DOC_ID
         
         If BdE.SUPPLIER_ID > 0 Then
            Set SpE = GetSupplierExt(m_SupplierExts, Trim(str(BdE.SUPPLIER_ID)))
            Set Sp = GetSupplier(m_Suppliers, SpE.SUPPLIER_CODE)
            BdE.SUPPLIER_ID = Sp.SUPPLIER_ID
         Else
            BdE.SUPPLIER_ID = -1
         End If
         
         If BdE.INVENTORY_DOC_ID > 0 Then
            Set IvdE = GetInventoryDoc(m_InventoryDocExts, Trim(str(BdE.INVENTORY_DOC_ID)))
            Set Ivd = GetInventoryDoc(m_InventoryDocs, IvdE.DOCUMENT_NO)
            BdE.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID
         Else
            BdE.INVENTORY_DOC_ID = -1
         End If
         
         BdE.AddEditMode = SHOW_ADD
         Call BdE.AddEditData
         
         
         '---------------------------------------------------------------------------- CBillingDiscountExt        1
         Set Bdd = New CBillingDiscountDup
         Bdd.BILLING_DOC_ID = OldID
         Call Bdd.QueryData(1, TempRs2, ItemCount)
         While Not TempRs2.EOF
            Call Bdd.PopulateFromRS(1, TempRs2)
            
            Set PiE = GetPartItemExt(m_PartItemExts, Trim(str(Bdd.PART_ITEM_ID)))
            Set Pi = GetPartItem(m_PartItems, PiE.PART_NO)

            Bdd.AddEditMode = SHOW_ADD
            Bdd.PART_ITEM_ID = Pi.PART_ITEM_ID
            Bdd.BILLING_DOC_ID = BdE.BILLING_DOC_ID
            Call Bdd.AddEditData

            TempRs2.MoveNext
         Wend
         Set Bdd = Nothing
         '---------------------------------------------------------------------------- CBillingDiscountExt     1
         
         '---------------------------------------------------------------------------- CBulkHoleExt                        2
         Set BuH = New CBulkHoleDup
         BuH.BILLING_DOC_ID = OldID
         Call BuH.QueryData(1, TempRs2, ItemCount)
         While Not TempRs2.EOF
            Call BuH.PopulateFromRS(1, TempRs2)
            
            Set PiE = GetPartItemExt(m_PartItemExts, Trim(str(BuH.PART_ITEM_ID)))
            Set Pi = GetPartItem(m_PartItems, PiE.PART_NO)
            
            BuH.AddEditMode = SHOW_ADD
            BuH.PART_ITEM_ID = Pi.PART_ITEM_ID
            BuH.BILLING_DOC_ID = BdE.BILLING_DOC_ID
            Call BuH.AddEditData
            
            TempRs2.MoveNext
         Wend
         Set BuH = Nothing
         '---------------------------------------------------------------------------- CBulkHoleExt                        2
         
         '---------------------------------------------------------------------------- CSupItem                         3
         Set Sup = New CSupItemDup
         Sup.DO_ID = OldID
         Call Sup.QueryData(1, TempRs2, ItemCount)
         While Not TempRs2.EOF
            Call Sup.PopulateFromRS(1, TempRs2)
            
            Set PiE = GetPartItemExt(m_PartItemExts, Trim(str(Sup.PART_ITEM_ID)))
            Set Pi = GetPartItem(m_PartItems, PiE.PART_NO)

            Sup.AddEditMode = SHOW_ADD
            Sup.PART_ITEM_ID = Pi.PART_ITEM_ID
            Sup.DO_ID = BdE.BILLING_DOC_ID
            Call Sup.AddEditData

            TempRs2.MoveNext
         Wend
         Set Sup = Nothing
         '---------------------------------------------------------------------------- CSupItem                         3
         
         '---------------------------------------------------------------------------- CDoItem                          4
         Set Doc = New CDoItemDup
         Doc.DO_ID = OldID
         Call Doc.QueryData(1, TempRs2, ItemCount)
         While Not TempRs2.EOF
            Call Doc.PopulateFromRS(1, TempRs2)
            
            Set PiE = GetPartItemExt(m_PartItemExts, Trim(str(Doc.PART_ITEM_ID)))
            Set Pi = GetPartItem(m_PartItems, PiE.PART_NO)

            Doc.AddEditMode = SHOW_ADD
            Doc.PART_ITEM_ID = Pi.PART_ITEM_ID
            Doc.DO_ID = BdE.BILLING_DOC_ID
            Call Doc.AddEditData
            
            TempRs2.MoveNext
         Wend
         Set Doc = Nothing
         '---------------------------------------------------------------------------- CDoItem                          4
         
         
         '---------------------------------------------------------------------------- CCashTran                          5
         Set ct = New CCashTranDup
         Call ct.SetFieldValue("BILLING_DOC_iD", OldID)
         Call ct.QueryData(1, TempRs2, ItemCount)
         While Not TempRs2.EOF
            Call ct.PopulateFromRS(1, TempRs2)
            
            ct.ShowMode = SHOW_ADD
            Call ct.SetFieldValue("BILLING_DOC_ID", BdE.BILLING_DOC_ID)
            Call ct.SetFieldValue("SUPPLIER_ID", BdE.SUPPLIER_ID)
            
            '---------------------------------------------------------------> หาเช็ค
            
            
            Call ct.AddEditData
            
            TempRs2.MoveNext
         Wend
         Set ct = Nothing
         '---------------------------------------------------------------------------- CCashTran                          5
         
         
         '---------------------------------------------------------------------------- CReceiptItem                          6
         Set Rec = New CReceiptItemDup
         Rec.BILLING_DOC_ID = OldID
         Call Rec.QueryData(1, TempRs2, ItemCount)
         While Not TempRs2.EOF
            Call Rec.PopulateFromRS(1, TempRs2)
            
            Set PiE = GetPartItemExt(m_PartItemExts, Trim(str(Rec.PART_ITEM_ID)))
            Set Pi = GetPartItem(m_PartItems, PiE.PART_NO)
            
            Set BD = GetInventoryDoc(m_BillingDocs, Rec.DOCUMENT_NO)
            
            Rec.AddEditMode = SHOW_ADD
            Rec.PART_ITEM_ID = Pi.PART_ITEM_ID
            Rec.BILLING_DOC_ID = BdE.BILLING_DOC_ID
            Rec.DO_ID = BD.BILLING_DOC_ID
            Call Rec.AddEditData
            
            TempRs2.MoveNext
         Wend
         Set Rec = Nothing
         '---------------------------------------------------------------------------- CReceiptItem                          6
         
      End If
      
      TempRs1.MoveNext
      
      txtPercent.Text = FormatNumber(Percent)
      prgProgress.Value = Percent
      Me.Refresh
   Wend
   
   prgProgress.Value = prgProgress.MAX
   
   glbDaily.CommitTransaction
   
   Set IvdE = Nothing
   If TempRs1.State = adStateOpen Then
      TempRs1.Close
   End If
   Set TempRs1 = Nothing
   
   If TempRs2.State = adStateOpen Then
      TempRs2.Close
   End If
   Set TempRs2 = Nothing
End Sub
