VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportPacking 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13785
   Icon            =   "frmImportPacking.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   13785
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   6525
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   13785
      _ExtentX        =   24315
      _ExtentY        =   11509
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlDate uctlFromDate 
         Height          =   405
         Left            =   2520
         TabIndex        =   17
         Top             =   3360
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtFileName 
         Height          =   435
         Left            =   2520
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   960
         Width           =   9795
         _ExtentX        =   17277
         _ExtentY        =   767
      End
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   2520
         TabIndex        =   4
         Top             =   2280
         Width           =   10425
         _ExtentX        =   18389
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   13755
         _ExtentX        =   24262
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtPercent 
         Height          =   465
         Left            =   2520
         TabIndex        =   5
         Top             =   2760
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
      Begin prjFarmManagement.uctlTextBox txtFileName2 
         Height          =   435
         Left            =   2520
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1560
         Width           =   9795
         _ExtentX        =   17277
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlToDate 
         Height          =   405
         Left            =   2520
         TabIndex        =   19
         Top             =   3960
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   714
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   840
         TabIndex        =   18
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   840
         TabIndex        =   20
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label lblFileName2 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   2295
      End
      Begin Threed.SSCommand cmdFileName2 
         Height          =   405
         Left            =   12480
         TabIndex        =   15
         Top             =   1560
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportPacking.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label lblNote 
         Height          =   795
         Left            =   480
         TabIndex        =   13
         Top             =   5520
         Width           =   12585
      End
      Begin Threed.SSCommand cmdFileName 
         Height          =   405
         Left            =   12480
         TabIndex        =   0
         Top             =   960
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportPacking.frx":2ABC
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   2640
         TabIndex        =   1
         Top             =   4800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportPacking.frx":2DD6
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   4320
         TabIndex        =   11
         Top             =   2880
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   120
         TabIndex        =   10
         Top             =   2280
         Width           =   2295
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   840
         TabIndex        =   9
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label lblFileName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   2295
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   11160
         TabIndex        =   3
         Top             =   4800
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   9360
         TabIndex        =   2
         Top             =   4800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportPacking.frx":30F0
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmImportPacking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Public id As Long
Public Area As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Public InventoryActArea As Long

Private m_ExcelApp As Object
Private m_ExcelSheet As Object
Private PartProduct As Collection
Private PartProductStock As Collection
Private PartProductNoData As Collection
Private m_JobInOut As Collection
Private m_Lot As Collection
Private m_Bin As Collection
Private m_Lock As Collection
Private m_PartItems As Collection
Private JobDocType As Long
Private m_Job As CJob
Private TempJob As CJob
Private TempJobIn As CJobInput
Public Count2 As Long
Public ChkRound As Boolean
Private Sub cmdFileName_Click()
 On Error Resume Next
 Dim strDescription As String
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
Call EnableForm(Me, False)
    If Not VerifyTextControl(lblFileName, txtFileName) Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   If Not VerifyTextControl(lblFileName2, txtFileName2) Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   Call LoadJobInputByDate(Nothing, m_JobInOut, DateSerial(2017, 1, 1), DateSerial(2017, 1, 31), 2)

   Call LoadLotFromLot(Nothing, m_Lot, , , , , , 1, , 2)
   Call LoadLocation(Nothing, m_Bin, 2, , , , 2, "BIN")
   Call LoadLocation(Nothing, m_Lock, 2, , , , 2, "LOCK")
   Call LoadPartItem(Nothing, m_PartItems, , "", , 2)
   If Area = 1 Then
     Call ImportPacking
   ElseIf Area = 2 Then
     Call ImportInventory
   End If
   Call EnableForm(Me, True)
End Sub
Private Sub ImportPacking()
Dim MaxRow As Long
Dim MaxCol As Long
Dim id As Long
Dim ID_Count As Long
Dim I As Long
Dim row As Long
Dim Col As Long
Dim ErrorCount As Long
Dim SuccessCount As Long
Dim ProgressCount As Long
Dim HasBegin As Boolean
Dim IsOK As Boolean
Dim SearchItemNo As CPartItem
Dim SearchItemNo2 As CJobInput
Dim SheetName As String
Dim MaxSheet As Long
Dim cData As CPartItem
Dim cDataStock As CJobInput
Dim strDate As String
Dim Key As String
Dim Ma As CJobInput
Dim TX_AMOUNT As Double
Dim TmpDate As Date
Dim DDMMYYYY As String

   HasBegin = False
   JobDocType = 1
   
   If Not ChkRound Then
      m_ExcelApp.Workbooks.Close
      m_ExcelApp.Workbooks.Open (txtFileName.Text)
      id = 1
      Set m_ExcelSheet = m_ExcelApp.Sheets(id)
      MaxRow = m_ExcelSheet.UsedRange.Rows.Count
      MaxCol = m_ExcelSheet.UsedRange.Columns.Count
      For row = 2 To MaxRow
         DoEvents
         Me.Refresh
          Set cData = New CPartItem
           lblNote.Caption = "จากเอกสาร : " & txtFileName.Text & " บรรทัดที่ : " & row
          
          cData.PART_NO = Trim(m_ExcelSheet.Cells(row, 2).Value)
          cData.PART_NO_PRODUCT = Trim(m_ExcelSheet.Cells(row, 3).Value)
          cData.PART_TYPE_BAG = Val(m_ExcelSheet.Cells(row, 4).Value)
          Set SearchItemNo = GetPartItem(m_PartItems, cData.PART_NO)
          cData.PART_ITEM_ID = SearchItemNo.PART_ITEM_ID
         
         Set SearchItemNo = GetObject("CPartItem", PartProduct, cData.PART_NO_PRODUCT & "-" & cData.PART_TYPE_BAG, False)
         If SearchItemNo Is Nothing Then
            Call PartProduct.add(cData, cData.PART_NO_PRODUCT & "-" & cData.PART_TYPE_BAG)
         End If
   
      Next row
      Set m_ExcelSheet = Nothing
      m_ExcelApp.Workbooks.Close
      ChkRound = True
   End If
   
   
   'File ที่ 2
   m_ExcelApp.Workbooks.Open (txtFileName2.Text)
   MaxSheet = m_ExcelApp.Sheets.Count
   Call glbDaily.StartTransaction
   TmpDate = uctlFromDate.ShowDate
   DDMMYYYY = Format(Day(TmpDate), "00") & "-" & Format(Month(TmpDate), "00") & "-" & Format(Year(TmpDate) + 543, "0000")
   For id = 1 To MaxSheet
      Set m_ExcelSheet = m_ExcelApp.Sheets(id)
      MaxRow = m_ExcelSheet.UsedRange.Rows.Count
      MaxCol = m_ExcelSheet.UsedRange.Columns.Count
      SheetName = m_ExcelApp.Sheets(id).NAME
      strDate = SplitStringToDate2(Trim(SheetName))
      If strDate >= uctlFromDate.ShowDate Then
        For row = 5 To 300 'MaxRow
            DoEvents
            Me.Refresh
             lblNote.Caption = "จากเอกสาร : " & txtFileName2.Text & " วันที่ : " & strDate & "  บรรทัดที่ : " & row
            If Val(m_ExcelSheet.Cells(row, 1).Value) > 0 Then 'ตรวจสอบว่า column เป็นข้อมูลที่ต้องการหรือไม่
                If (Not SplitStringToDate(Trim(m_ExcelSheet.Cells(row, 1).Value)) = SheetName) And (Val(m_ExcelSheet.Cells(row, 5).Value) > 0) And (Trim(m_ExcelSheet.Cells(row, 2).Value) <> "รวม") Then ' ตรวจสอบว่า ข้อมูลไม่เป็นวันที่ใช่หรือไม่ และ Column ที่ 5 มีค่ามากกว่า 0 หรือไม่
                     If Not (PartProductStock Is Nothing) Then
                     ' set ค่าเพื่อ insert แม่
                     Set m_Job = New CJob
                      m_Job.AddEditMode = SHOW_ADD
                      m_Job.FORMULA_ID = -1
                      m_Job.JOB_NO = getJobplan()
                      m_Job.JOB_DESC = Trim(m_ExcelSheet.Cells(row, 2).Value) & "(" & Trim(m_ExcelSheet.Cells(row, 11).Value) & ")"
                      m_Job.JOB_DATE = CDate(strDate)
                      m_Job.BATCH_NO = ""
                      m_Job.START_DATE = CDate(m_ExcelSheet.Cells(row, 10).Value)
                      m_Job.FINISH_DATE = CDate(m_ExcelSheet.Cells(row, 10).Value)
                      m_Job.COMMIT_FLAG = "N"
                      m_Job.PROCESS_ID = 2
                      m_Job.JOB_DOC_TYPE = JobDocType
                      
                      'เก็บข้อมูลจาก Stock โกดังเข้า collection PartProductStock
                    Set cDataStock = New CJobInput
                     cDataStock.MIX_DATE = strDate
                     cDataStock.PART_NO = Trim(m_ExcelSheet.Cells(row, 2).Value)
                     cDataStock.WEIGHT_PER_PACK = splitStr(Trim(m_ExcelSheet.Cells(row, 8).Formula))
                     cDataStock.PACK_AMOUNT = Val(m_ExcelSheet.Cells(row, 5).Value)
                       
                       'ดึง PART_ITEM_ID จาก PartProduct มาเก็บที่ cDataStock
                       Key = cDataStock.PART_NO & "-" & cDataStock.WEIGHT_PER_PACK
                       Set SearchItemNo = GetObject("CPartItem", PartProduct, Key, False)
                       If Not SearchItemNo Is Nothing Then
                         cDataStock.PART_ITEM_ID = SearchItemNo.PART_ITEM_ID ' ดึง PART_ITEM_ID จาก PartProduct มาเก็บที่ PartProductStock
                       Else
                          cDataStock.PART_ITEM_ID = -1
                          Key = cDataStock.PART_NO & "-" & cDataStock.WEIGHT_PER_PACK
                          Set SearchItemNo2 = GetObject("CPartItem", PartProductNoData, Key, False)
                           If SearchItemNo2 Is Nothing Then
                              Call PartProductNoData.add(cDataStock, Key)
                              Count2 = Count2 + 1
                              ''Debug.Print Count2 & ".  " & cDataStock.PART_NO & "," & cDataStock.WEIGHT_PER_PACK
                           End If
                       End If
                       Set SearchItemNo2 = Nothing
   
                       'ค้นหาเบอร์อาหารที่ข้อมูลในแต่ละวันซ้ำกัน และให้บวกจำนวน PACK_AMOUNT ด้วย
                        Key = cDataStock.MIX_DATE & "-" & cDataStock.PART_NO & "-" & cDataStock.WEIGHT_PER_PACK
                        Set SearchItemNo2 = GetObject("CPartItem", PartProductStock, Key, False)
                         If SearchItemNo2 Is Nothing Then
                            Call PartProductStock.add(cDataStock, Key)
                         Else
                              SearchItemNo2.PACK_AMOUNT = SearchItemNo2.PACK_AMOUNT + Val(m_ExcelSheet.Cells(row, 5).Value)
                         End If
                      
                      TX_AMOUNT = 0
                       'Input อาหารสำเร็จรูปที่ได้
                        Set Ma = Nothing
                        Set Ma = New CJobInput
                        Ma.PART_ITEM_ID = cDataStock.PART_ITEM_ID
                        Ma.PART_DESC = Trim(m_ExcelSheet.Cells(row, 3).Value)
                        Ma.PART_NO = Trim(m_ExcelSheet.Cells(row, 2).Value)
                        Ma.PART_TYPE_ID = 10
                        Ma.PART_TYPE_NAME = "สินค้าสำเร็จรูป"
                        Ma.LOCATION_ID = 109
                        Ma.LOCATION_NO = ".GO"
                        Ma.LOCATION_NAME = ".โกดังอาหาร"
                        Ma.SERIAL_NUMBER = "LOT" & Trim(m_ExcelSheet.Cells(row, 11).Value)
                        Ma.INOUT_REF = "BIN" & Trim(m_ExcelSheet.Cells(row, 11).Value)
                        Ma.TX_TYPE = "I"
                        Ma.WEIGHT_PER_PACK = splitStr(Trim(m_ExcelSheet.Cells(row, 8).Formula))
                        Ma.PACK_AMOUNT = Val(m_ExcelSheet.Cells(row, 5).Value)
                        Ma.TX_AMOUNT = Ma.WEIGHT_PER_PACK * Ma.PACK_AMOUNT
                        TX_AMOUNT = Ma.TX_AMOUNT
                        Ma.STD_AMOUNT = Ma.TX_AMOUNT
                        Ma.Flag = "A"
                        Call m_Job.Outputs.add(Ma)
                        'สิ้นสุดการ Input อาหารสำเร็จรูปที่ได้
                      
                       'Input ส่วนผสมที่ใช้
                       Set TempJob = GetObject("Cjob", m_JobInOut, cDataStock.PART_ITEM_ID)
                       If Not TempJob Is Nothing Then
                       For Each TempJobIn In TempJob.Inputs
                         Set Ma = New CJobInput
                        Ma.PART_ITEM_ID = TempJobIn.PART_ITEM_ID
                        Ma.PART_TYPE_ID = TempJobIn.PART_TYPE_ID
                        Ma.PART_TYPE_NAME = TempJobIn.PART_TYPE_NAME
                        If Ma.PART_TYPE_ID = 26 Or Ma.PART_TYPE_ID = 29 Or Ma.PART_TYPE_ID = 30 Or Ma.PART_TYPE_ID = 31 Or Ma.PART_TYPE_ID = 47 Or Ma.PART_TYPE_ID = 48 Then
                          Ma.TX_AMOUNT = (TX_AMOUNT * 2) / 100
                          Ma.PART_TYPE_ID = 22
                          Ma.LOCATION_ID = 117
                          Ma.LOCATION_NO = ".PACK"
                       Else
                          Ma.TX_AMOUNT = (TX_AMOUNT * 95) / 100
                          Ma.PART_TYPE_ID = 22
                          Ma.LOCATION_ID = 110
                          Ma.LOCATION_NO = ".BK"
                        End If
                        Ma.TX_TYPE = "E" 'TempJobIn.TX_TYPE
                        Ma.Flag = "A"
                        Call m_Job.Inputs.add(Ma)
                       Next TempJobIn
                       End If
                       If cDataStock.PART_ITEM_ID > 0 Then
                          Call SaveData
                       End If
                           Set cDataStock = Nothing
                         Set m_Job = Nothing
                     End If
                End If
           End If
         Next row
      ElseIf strDate < uctlFromDate.ShowDate Then
      Else
         id = id - 1
         ID_Count = ID_Count + 1
         If ID_Count > 31 Then
            lblNote.Caption = "เอกสารไม่ตรงกับวันที่ที่เลือก !!!"
            Exit For ' แสดงว่า file นี้ กับวันที่ที่เลือก ไม่สัมพันธ์กัน
         End If
       End If
       TmpDate = TmpDate - 1
       DDMMYYYY = Format(Day(TmpDate), "00") & "-" & Format(Month(TmpDate), "00") & "-" & Format(Year(TmpDate) + 543, "0000")
   Next id
   Call glbDaily.CommitTransaction
   Set m_ExcelSheet = Nothing
   m_ExcelApp.Workbooks.Close
   m_ExcelApp.Quit
End Sub
Private Sub ImportInventory()
Dim MaxRow As Long
Dim MaxCol As Long
Dim id As Long
Dim ID_Count As Long
Dim I As Long
Dim row As Long
Dim Col As Long
Dim ErrorCount As Long
Dim SuccessCount As Long
Dim ProgressCount As Long
Dim HasBegin As Boolean
Dim IsOK As Boolean
Dim SearchItemNo As CPartItem
Dim SLotItemWH As CLotItemWH
Dim SLocation As CLocation
Dim SLotDoc As CLotDoc
Dim SearchItemNo2 As CJobInput
Dim SheetName As String
Dim MaxSheet As Long
Dim cData As CPartItem
Dim cDataStock As CJobInput
Dim strDate As String
Dim Key As String
Dim Ma As CJobInput
Dim TX_AMOUNT As Double
Dim TmpDate As Date
Dim DDMMYYYY As String
Dim m_InventoryWh  As CInventoryWHDoc

   HasBegin = False
   JobDocType = 1
   
   If Not ChkRound Then
      m_ExcelApp.Workbooks.Close
      m_ExcelApp.Workbooks.Open (txtFileName.Text)
      id = 1
      Set m_ExcelSheet = m_ExcelApp.Sheets(id)
      MaxRow = m_ExcelSheet.UsedRange.Rows.Count
      MaxCol = m_ExcelSheet.UsedRange.Columns.Count
      For row = 2 To MaxRow
         DoEvents
         Me.Refresh
          Set cData = New CPartItem
           lblNote.Caption = "จากเอกสาร : " & txtFileName.Text & " บรรทัดที่ : " & row
          
          cData.PART_NO = Trim(m_ExcelSheet.Cells(row, 2).Value)
          cData.PART_NO_PRODUCT = Trim(m_ExcelSheet.Cells(row, 3).Value)
          cData.PART_TYPE_BAG = Val(m_ExcelSheet.Cells(row, 4).Value)
          Set SearchItemNo = GetPartItem(m_PartItems, cData.PART_NO) 'หา PART_ITEM_ID
          cData.PART_ITEM_ID = SearchItemNo.PART_ITEM_ID
         
         Set SearchItemNo = GetObject("CPartItem", PartProduct, cData.PART_NO_PRODUCT & "-" & cData.PART_TYPE_BAG, False)
         If SearchItemNo Is Nothing Then
            Call PartProduct.add(cData, cData.PART_NO_PRODUCT & "-" & cData.PART_TYPE_BAG)
         End If
   
      Next row
      Set m_ExcelSheet = Nothing
      m_ExcelApp.Workbooks.Close
      ChkRound = True
   End If
   

   
 
   'File ที่ 2
   Dim IWH As CInventoryWHDoc
   Dim LIW As CLotItemWH
   Dim LTD As CLotDoc
   Dim PD As CPalletDoc
   Set IWH = New CInventoryWHDoc
   m_ExcelApp.Workbooks.Open (txtFileName2.Text)
   MaxSheet = m_ExcelApp.Sheets.Count
   Call glbDaily.StartTransaction
   TmpDate = uctlToDate.ShowDate
   DDMMYYYY = Format(Day(TmpDate), "00") & "-" & Format(Month(TmpDate), "00") & "-" & Format(Year(TmpDate) + 543, "0000")
   For id = MaxSheet To 1 Step -1
      Set m_ExcelSheet = m_ExcelApp.Sheets(id)
      MaxRow = m_ExcelSheet.UsedRange.Rows.Count
      MaxCol = m_ExcelSheet.UsedRange.Columns.Count
      SheetName = m_ExcelApp.Sheets(id).NAME
      strDate = SplitStringToDate2(Trim(SheetName))
      If strDate >= uctlFromDate.ShowDate Then
        For row = 5 To 300 'MaxRow
            DoEvents
            Me.Refresh
            If row = 138 Then
               ''Debug.Print row
            End If
             lblNote.Caption = "จากเอกสาร : " & txtFileName2.Text & " วันที่ : " & strDate & "  บรรทัดที่ : " & row
            If Not Val(m_ExcelSheet.Cells(row, 1).Value) > 0 Then 'ตรวจสอบว่า column เป็นข้อมูลที่ต้องการหรือไม่
               GoTo Continue 'ถ้าไม่ใช่ ให้ไปแถวต่อไป
            End If
            If Not (Not SplitStringToDate(Trim(m_ExcelSheet.Cells(row, 1).Value)) = SheetName) And (Val(m_ExcelSheet.Cells(row, 7).Value) > 0) And (Trim(m_ExcelSheet.Cells(row, 2).Value) <> "รวม") Then ' ตรวจสอบว่า ข้อมูลไม่เป็นวันที่ใช่หรือไม่ และ Column ที่ 5 มีค่ามากกว่า 0 หรือไม่
               GoTo Continue 'ถ้าไม่ใช่ ให้ไปแถวต่อไป
            End If
            If (PartProductStock Is Nothing) Then
               GoTo Continue 'ถ้าไม่มี ให้ไปแถวต่อไป
            End If
 'เก็บข้อมูลเข้า LotItemWH
           Dim tmpPART_ITEM_ID As Long
           Key = Trim(m_ExcelSheet.Cells(row, 2).Value) & "-" & splitStr(Trim(m_ExcelSheet.Cells(row, 8).Formula))
           Set SearchItemNo = GetObject("CPartItem", PartProduct, Key, False)
           If Not SearchItemNo Is Nothing Then
               tmpPART_ITEM_ID = SearchItemNo.PART_ITEM_ID
           Else
           ''Debug.Print row & "," & m_ExcelSheet.Cells(row, 2).Value
               tmpPART_ITEM_ID = -1
               GoTo Continue 'ถ้าไม่มี ให้ไปแถวต่อไป
           End If 'If Not SearchItemNo Is Nothing Then
           
              'load ข้อมูลสินค้าทั้งหมด ที่มีทั้ง lot ทั้ง pallet
            Dim ItemCount As Long
            Dim c_PartItem As Collection
            Dim m_PartItem As CLotItemWH
            Set c_PartItem = New Collection
            Set m_PartItem = New CLotItemWH
         
            m_PartItem.PART_ITEM_ID = tmpPART_ITEM_ID
            m_PartItem.QueryFlag = 1
            If Not glbDaily.QueryLotItemWhPart2(m_PartItem, c_PartItem, ItemCount, IsOK, glbErrorLog) Then
               GoTo Continue 'ถ้าไม่มี ให้ไปแถวต่อไป
            End If
'            'เก็บข้อมูลเข้า LotDoc
            Dim TempKey As String
            Dim ID2 As Long
'            Set LTD = New CLotDoc
            'Key = ConvertString2Lot(Trim(m_ExcelSheet.Cells(row, 11).Value), m_ExcelSheet.Cells(row, 10).Value)
            Set m_PartItem = Nothing
            ID2 = 0
            For Each m_PartItem In c_PartItem
               TempKey = Mid(m_PartItem.LOT_NO, 9, 3) 'เอา 3 ตัวหลังออกมาจาก LG171201 222
               Key = "BIN" & Trim(m_ExcelSheet.Cells(row, 12).Value)
               If Not (Format(m_ExcelSheet.Cells(row, 11).Value, "000") = TempKey And Key = m_PartItem.BIN_NAME) Then 'เช็ค lot และ bin จาก excel ว่าตรงกับ ใน collection หรือไม่
                  GoTo Continue 'ถ้าไม่มี ให้ไปแถวต่อไป
               End If
                     
               ID2 = ID2 + 1
               'ถ้าทุกอย่างถูกต้องแล้วให้เริ่ม บรรจุข้อมูล
               'เก็บข้อมูลเข้า InventoryWHDoc
               Dim No As String
               
               Call glbDatabaseMngr.GenerateNumber(AVG_GOODS, No, glbErrorLog)
               IWH.DOCUMENT_NO = No
               IWH.DOCUMENT_DATE = strDate
               IWH.DOCUMENT_TYPE = 2000 'ปรับยอด
               IWH.ENTRY_DATE = -1
               IWH.EXIT_DATE = -1
               IWH.YYYYMM = SplitStringToDate3(Trim(SheetName))
               IWH.CUSTOMER_ID = -1
              
              'เก็บข้อมูลเข้า LotItemWH
               Set LIW = New CLotItemWH
               Key = "BIN" & Trim(m_ExcelSheet.Cells(row, 12).Value)
               Set SLocation = GetObject("CLotItemWH", m_Bin, Key, False)
               If Not SLocation Is Nothing Then
                  LIW.BIN_NO = SLocation.KEY_ID
               Else
                  LIW.BIN_NO = -1
               End If 'If Not SLocation Is Nothing Then
               Set SLocation = Nothing
               Key = ConvertString2Lock(Trim(m_ExcelSheet.Cells(row, 13).Value))
               Set SLocation = GetObject("CLotItemWH", m_Lock, Key, False)
               If Not SLocation Is Nothing Then
                  LIW.LOCK_NO = SLocation.KEY_ID
               Else
                  LIW.LOCK_NO = -1
               End If 'If Not SLocation Is Nothing Then
               Set SLocation = Nothing
               LIW.PART_ITEM_ID = tmpPART_ITEM_ID
               LIW.LOCATION_ID = 109
               LIW.CALCULATE_FLAG = "N"
               LIW.TX_TYPE = "E"
               LIW.WEIGHT_PER_PACK = splitStr(Trim(m_ExcelSheet.Cells(row, 8).Formula))
               LIW.BILLING_DOC_ID = -1
               
               
               'เก็บข้อมูลเข้า LotDoc
               Set LTD = New CLotDoc
               LTD.LOT_DOC_ID_REF = m_PartItem.LOT_DOC_ID
               LTD.BIN_NO = LIW.BIN_NO
               
               'เก็บข้อมูลเข้า PalletDoc
               Dim c_LTD As Collection
               Dim c_PD As Collection
               Dim s_PD As CPalletDoc
               Dim palletAmount As Long 'คงเหลือในระบบ
               Dim tAmount As Long 'คงเหลือใน excel
               Dim C As Long
               
               Set c_LTD = c_PartItem.Item(1).C_LotDoc
               If Not c_LTD Is Nothing Then
               
               
               tAmount = CInt(m_ExcelSheet.Cells(row, 7).Value)
               For Each LTD In c_LTD
                 C = C + 1
                 If m_PartItem.LOT_NO = LTD.LOT_NO Then
                    Set c_PD = LTD.C_PalletDoc
                    palletAmount = GetTotalAmountPallet(c_PD)
                     For Each s_PD In c_PD
                        Set PD = New CPalletDoc
                         If (palletAmount = tAmount) Then
                            'Debug.Print
                         End If
                        If (palletAmount - tAmount) - s_PD.CAPACITY_AMOUNT > 0 Then
                           PD.CAPACITY_AMOUNT = s_PD.CAPACITY_AMOUNT
                           PD.TX_TYPE = "E"
                           palletAmount = palletAmount - s_PD.CAPACITY_AMOUNT
   '                        PD.LOT_DOC_ID
                           Call LTD.C_PalletDoc.add(PD)
                        ElseIf (s_PD.CAPACITY_AMOUNT - (palletAmount - tAmount) > 0) And (palletAmount - tAmount) <> 0 Then
                           PD.CAPACITY_AMOUNT = palletAmount - tAmount
                           PD.TX_TYPE = "E"
                           Call LTD.C_PalletDoc.add(PD)
                           Exit For
                        End If
                        Set PD = Nothing
                     Next s_PD
                     Call LIW.C_LotDoc.add(LTD)
                  End If
               Next LTD
                  Call IWH.C_LotItemsWH.add(LIW)
               End If
            Next m_PartItem

Continue:
          ID2 = 0
         Next row
      ElseIf strDate < uctlFromDate.ShowDate Then
      Else
         id = id - 1
         ID_Count = ID_Count + 1
         If ID_Count > 31 Then
            lblNote.Caption = "เอกสารไม่ตรงกับวันที่ที่เลือก !!!"
            Exit For ' แสดงว่า file นี้ กับวันที่ที่เลือก ไม่สัมพันธ์กัน
         End If
       End If
       TmpDate = TmpDate + 1
       DDMMYYYY = Format(Day(TmpDate), "00") & "-" & Format(Month(TmpDate), "00") & "-" & Format(Year(TmpDate) + 543, "0000")
   Next id
   Call glbDaily.CommitTransaction
   Set m_ExcelSheet = Nothing
   m_ExcelApp.Workbooks.Close
   m_ExcelApp.Quit
End Sub
Function ConvertString2Lock(str As String) As String
   Dim TempStr() As String
   If str = "" Then
      ConvertString2Lock = ""
      Exit Function
   End If
   TempStr() = Split(str, "-")
   If UBound(TempStr()) = 1 Then
      ConvertString2Lock = "L-" & Trim(TempStr(0) & Format(TempStr(1), "00"))
   Else
      ConvertString2Lock = ""
   End If
   
End Function
Function ConvertString2Lot(LotStr As String, DateStr As String) As String
    ConvertString2Lot = "LG" & Right(Format(Year(DateStr), "0000"), 2) & Format(CDate(DateStr), "mm") & Format(CDate(DateStr), "dd") & Format(CInt(LotStr), "00") '"LG" & Format(Year(DateStr), "00") & Format(Month(DateStr), "00") & Format(Day(DateStr), "00") & Format(CInt(LotStr), "00")
End Function
Private Sub PopulateGuiID(BD As CJob)
Dim Di As CJobInput

   For Each Di In BD.Inputs
      If Di.Flag = "A" Then
         Di.LINK_ID = GetNextGuiID(BD)
      End If
   Next Di

   For Each Di In BD.Outputs
      If Di.Flag = "A" Then
         Di.LINK_ID = GetNextGuiID(BD)
      End If
   Next Di
End Sub

Private Function GetNextGuiID(BD As CJob) As Long
Dim Di As CJobInput
Dim MaxId As Long

   MaxId = 0
   For Each Di In BD.Inputs
      If Di.LINK_ID > MaxId Then
         MaxId = Di.LINK_ID
      End If
   Next Di
   
   For Each Di In BD.Outputs
      If Di.LINK_ID > MaxId Then
         MaxId = Di.LINK_ID
      End If
   Next Di
   
   GetNextGuiID = MaxId + 1
End Function
Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim Ivd As CInventoryDoc
Dim JO As CJobInput
   Call EnableForm(Me, False)
   Call PopulateGuiID(m_Job)
   If JobDocType = 1 Then
      Call glbDaily.Job2InventoryDoc(m_Job, Ivd, 1, 11)

      If (m_Job.COMMIT_FLAG = "Y") Then
         If m_Job.OLD_COMMIT_FLAG <> "Y" Then
            Call glbDaily.TriggerCommit(Ivd.ImportExports)
            If Not glbDaily.VerifyStockBalance(Ivd.ImportExports, glbErrorLog) Then
               Call EnableForm(Me, True)
               Exit Function
            End If
         End If
      End If
   End If

'   Call glbDaily.StartTransaction
   If JobDocType = 1 Then
      If Not glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         SaveData = False
         Call glbDaily.RollbackTransaction
         Call EnableForm(Me, True)
         Exit Function
      End If

      m_Job.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID
   Else
      m_Job.INVENTORY_DOC_ID = -1
   End If
   If Not glbProduction.AddEditJob(m_Job, IsOK, False, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call glbDaily.RollbackTransaction
      Call EnableForm(Me, True)
      Exit Function
   End If
'   Call glbDaily.CommitTransaction

   If Not IsOK Then
      Call EnableForm(Me, True)
      glbErrorLog.ShowUserError
      Exit Function
   End If

   Call EnableForm(Me, True)
   SaveData = True
End Function
Private Function getJobplan() As String
Dim No As String
      If JobDocType = 1 Then
         Call glbDatabaseMngr.GenerateNumber(JOBPLAN_AUTO_NUMBER, No, glbErrorLog)
         getJobplan = No
      ElseIf JobDocType = 2 Then
         Call glbDatabaseMngr.GenerateNumber(ESTIMATE_NUMBER, No, glbErrorLog)
         getJobplan = No
      End If
End Function
Private Function splitStr(str As String) As String
Dim data() As String
   If str = "" Then
      splitStr = -1
      Exit Function
   End If
   data = Split(str, "/")
   data = Split(data(0), "*")
   If Len(data(1)) > 0 Then
      splitStr = data(1)
   Else
      splitStr = "-1"
   End If
End Function
Private Sub Form_Activate()
Dim FromDate As Date
Dim ToDate As Date

  
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      Count2 = 0
      ChkRound = False
'      Call LoadJobInputByDate(Nothing, m_JobInOut, DateSerial(2017, 1, 1), DateSerial(2017, 1, 31), 2)
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
   
'   Call InitNormalLabel(lblJobDate, MapText("วันที่สั่งผลิต"))
   
  ' Call InitNormalLabel(lblNote, "- เริ่ม Import ที่ Row ที่ 2 ,Column A ของ Sheet1 เท่านั้น โดย" & vbCrLf & "Col A = รหัสวัตถุดิบ,Col B =รายละเอียดวัตถุดิบ,Col C =จำนวนเป็นตัน/กิโลกรัม" & vbCrLf & " ***โดยห้ามให้เบอร์วัตถุดิบซ้ำกัน***")
   
   Call InitNormalLabel(lblFileName, "ชื่อไฟล์เบอร์อาหาร")
   Call InitNormalLabel(lblFileName2, "ชื่อไฟล์สต๊อกอาหาร")
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(Label1, "%")
   Call InitNormalLabel(lblFromDate, MapText("จากวันที่"))
   Call InitNormalLabel(lblToDate, MapText("ถึงวันที่"))

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

   Set m_PartItems = New Collection
   Set PartProduct = New Collection
   Set PartProductStock = New Collection
   Set PartProductNoData = New Collection
   Set m_JobInOut = New Collection
   Set m_Lot = New Collection
   Set m_Bin = New Collection
   Set m_Lock = New Collection
   
   Set m_ExcelApp = CreateObject("Excel.application")
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set PartProduct = Nothing
   Set PartProductStock = Nothing
   Set PartProductNoData = Nothing
   Set m_PartItems = Nothing
   Set m_JobInOut = Nothing
   Set m_Lot = Nothing
   Set m_Bin = Nothing
   Set m_Lock = Nothing
End Sub


Private Sub lblFileName_Click()
   ChkRound = False
End Sub
