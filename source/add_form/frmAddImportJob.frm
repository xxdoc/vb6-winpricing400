VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddImportJob 
   ClientHeight    =   2715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13755
   Icon            =   "frmAddImportJob.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   13755
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog dlgAdd 
      Left            =   15000
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13560
      _ExtentX        =   23918
      _ExtentY        =   4471
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextBox txtFileName 
         Height          =   435
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   360
         Width           =   10515
         _ExtentX        =   18547
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtFileName2 
         Height          =   435
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   840
         Width           =   10515
         _ExtentX        =   18547
         _ExtentY        =   767
      End
      Begin VB.Label lblNote 
         Caption         =   "Label1"
         Height          =   855
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   11175
      End
      Begin Threed.SSCommand cmdFileName 
         Height          =   405
         Left            =   10740
         TabIndex        =   5
         Top             =   360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddImportJob.frx":27A2
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdFileName2 
         Height          =   405
         Left            =   10740
         TabIndex        =   4
         Top             =   825
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddImportJob.frx":2ABC
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   11520
         TabIndex        =   1
         Top             =   720
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddImportJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Job As CJob
Private m_Jobs As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public JobDocType As Long
Public ProcessID As Long

Private FileName As String
Private m_SumUnit As Double
Private m_Employees As Collection
Private m_FormulaID As Long

Public TempCollection As Collection
Private PartProduct As Collection
Private PartProductStock As Collection
Private m_PartItems As Collection
Private m_ExcelApp As Object
Private m_ExcelSheet As Object

Private Sub EnableDisableButton(En As Boolean)
'   If En Then
'      If ShowMode = SHOW_EDIT Then
'         cmdAdd.Enabled = (m_Job.COMMIT_FLAG = "N")
'         cmdDelete.Enabled = (m_Job.COMMIT_FLAG = "N")
'      Else
'         cmdAdd.Enabled = True
'         cmdEdit.Enabled = True
'         cmdDelete.Enabled = True
'      End If
'   Else
'      cmdAdd.Enabled = En
'      cmdDelete.Enabled = En
'      cmdEdit.Enabled = En
'   End If
End Sub

Private Sub cboGroup_Click()
   m_HasModify = True
End Sub

Private Sub chkEnable_Click()
   m_HasModify = True
End Sub

Private Function GetJobPartItemID(Col As Collection) As CJobInput
Dim JO As CJobInput

   For Each JO In Col
      If JO.Flag <> "D" Then
         Set GetJobPartItemID = JO
      End If
   Next JO
End Function


Private Function SaveData2() As Boolean
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

   Call glbDaily.StartTransaction
   If JobDocType = 1 Then
      If Not glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         SaveData2 = False
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
      SaveData2 = False
      Call glbDaily.RollbackTransaction
      Call EnableForm(Me, True)
      Exit Function
   End If
   Call glbDaily.CommitTransaction

   If Not IsOK Then
      Call EnableForm(Me, True)
      glbErrorLog.ShowUserError
      Exit Function
   End If

   Call EnableForm(Me, True)
   SaveData2 = True
End Function
Private Sub cboJobProcess_Change()
m_HasModify = True
End Sub

Private Sub cboJobProcess_Click()
   m_HasModify = True
End Sub

Private Sub cboJobRef_Change()
m_HasModify = True
End Sub

Private Sub cboJobRef_Click()
m_HasModify = True
End Sub

Private Sub chkCommit_Click(Value As Integer)
m_HasModify = True
End Sub
Private Function getJobplan() As String
Dim No As String
      If JobDocType = 1 Then
         Call glbDatabaseMngr.GenerateNumber(JOBPLAN_NUMBER, No, glbErrorLog)
         getJobplan = No
      ElseIf JobDocType = 2 Then
         Call glbDatabaseMngr.GenerateNumber(ESTIMATE_NUMBER, No, glbErrorLog)
         getJobplan = No
      End If
End Function


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
   dlgAdd.Filter = "Access Files (*.XLS)|*..xls;*.XLS;"
   dlgAdd.DialogTitle = "Select access file to import"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If

   txtFileName2.Text = dlgAdd.FileName
   m_HasModify = True
End Sub
Private Sub cmdStart_Click()
 Call LoadPartItem(Nothing, m_PartItems, , "", , 2)
 Call ImportPacking
End Sub
Private Sub ImportPacking()
Dim MaxRow As Long
Dim MaxCol As Long
Dim ID As Long
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
Dim strDate As Date
Dim Key As String
Dim Ma As CJobInput

   HasBegin = False

   m_ExcelApp.Workbooks.Close
   m_ExcelApp.Workbooks.Open (txtFileName.Text)

   JobDocType = 1
   ID = 1
   Set m_ExcelSheet = m_ExcelApp.Sheets(ID)
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


   'File ที่ 2
   m_ExcelApp.Workbooks.Open (txtFileName2.Text)
   MaxSheet = m_ExcelApp.Sheets.Count

   For ID = 1 To MaxSheet
      Set m_ExcelSheet = m_ExcelApp.Sheets(ID)
      MaxRow = m_ExcelSheet.UsedRange.Rows.Count
      MaxCol = m_ExcelSheet.UsedRange.Columns.Count
      SheetName = m_ExcelApp.Sheets(ID).NAME
      strDate = SplitStringToDate(Trim(SheetName))

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

                     Key = cDataStock.PART_NO & "-" & cDataStock.WEIGHT_PER_PACK
                     Set SearchItemNo = GetObject("CPartItem", PartProduct, Key, False)
                     If Not SearchItemNo Is Nothing Then
                       cDataStock.PART_ITEM_ID = SearchItemNo.PART_ITEM_ID ' ดึง PART_ITEM_ID จาก PartProduct มาเก็บที่ PartProductStock
                     Else
                        cDataStock.PART_ITEM_ID = -1
                        Debug.Print cDataStock.PART_NO
                     End If


                      Key = cDataStock.MIX_DATE & "-" & cDataStock.PART_NO & "-" & cDataStock.WEIGHT_PER_PACK
                      Set SearchItemNo2 = GetObject("CPartItem", PartProductStock, Key, False)
                       If SearchItemNo2 Is Nothing Then
                          Call PartProductStock.add(cDataStock, Key)
                       Else
                            SearchItemNo2.WEIGHT_PER_PACK = SearchItemNo2.WEIGHT_PER_PACK + Val(m_ExcelSheet.Cells(row, 5).Value)
                       End If


                     'Input อาหารสำเร็จรูปที่ได้
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
                      Ma.STD_AMOUNT = Ma.TX_AMOUNT
                      Ma.Flag = "A"
                      
                      Call m_Job.Outputs.add(Ma)
                      'สิ้นสุดการ Input อาหารสำเร็จรูปที่ได้

                     If cDataStock.PART_ITEM_ID > 0 Then
                        Call SaveData2
                     End If
                         Set cDataStock = Nothing
                       Set m_Job = Nothing
                   End If
              End If
         End If
       Next row
   Next ID
   Set m_ExcelSheet = Nothing
End Sub
Private Function InputJobInput()
   Dim Ma As CJobInput
   If ShowMode = SHOW_ADD Then
      Set Ma = New CJobInput
   Else
      Set Ma = TempCollection.Item(ID)
   End If

'   Ma.PART_ITEM_ID = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
'   Ma.PART_DESC = uctlProductLookup.MyCombo.Text
'   Ma.PART_NO = uctlProductLookup.MyTextBox.Text
'   Ma.PART_TYPE_ID = uctlPartTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPartTypeLookup.MyCombo.ListIndex))
'   Ma.PART_TYPE_NAME = uctlPartTypeLookup.MyCombo.Text
'   Ma.TX_AMOUNT = txtAmount.Text
'   Ma.LOCATION_ID = uctlPlaceLookup.MyCombo.ItemData(Minus2Zero(uctlPlaceLookup.MyCombo.ListIndex))
'   Ma.LOCATION_NO = uctlPlaceLookup.MyTextBox.Text
'   Ma.LOCATION_NAME = uctlPlaceLookup.MyCombo.Text
'   Ma.SERIAL_NUMBER = txtSerialNo.Text
'   Ma.INOUT_REF = txtRef.Text
'   Ma.TX_TYPE = "I"
'   Ma.STD_AMOUNT = Val(txtStdAmount.Text)
'   Ma.WEIGHT_PER_PACK = Val(txtWeightPerPack.Text)
'   Ma.PACK_AMOUNT = Val(txtPackAmount.Text)

   If ShowMode = SHOW_ADD Then
      Ma.Flag = "A"
      Call TempCollection.add(Ma)
   Else
      If Ma.Flag <> "A" Then
         Ma.Flag = "E"
      End If
   End If
End Function
Private Function splitStr(str As String) As String
Dim Data() As String
   Data = Split(str, "/")
   Data = Split(Data(0), "*")
   If Len(Data(1)) > 0 Then
      splitStr = Data(1)
   Else
      splitStr = "-1"
   End If
End Function
Private Sub Form_Activate()
'   If Not m_HasActivate Then
'      m_HasActivate = True
'      Me.Refresh
'      DoEvents
'      m_HasModify = False
'   End If
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
'      Call LoadEmployee(ucltApproveByLookup.MyCombo, m_Employees)
'      Set ucltApproveByLookup.MyCollection = m_Employees
'
'      Call LoadEmployee(uctlResponseByLookup.MyCombo, m_Employees)
'      Set uctlResponseByLookup.MyCollection = m_Employees
      
'      Call LoadEmployee(uctlResponseByLookup.MyCombo)
'      Call LoadProcess(cboJobProcess)
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_Job.QueryFlag = 1
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
'         uctlJobDate.ShowDate = Now
'         uctlStartJob.ShowDate = Now
'         uctlFinishJob.ShowDate = Now
'         cboJobProcess.ListIndex = IDToListIndex(cboJobProcess, ProcessID)
         
        m_Job.QueryFlag = 0
         Call QueryData(False)
      End If
      
     ' Call TabStrip1_Click
      Call EnableForm(Me, True)
      m_HasModify = False
   End If
End Sub
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
      
      m_Job.JOB_ID = ID
      m_Job.QueryFlag = 1
      If Not glbProduction.QueryJob(m_Job, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
        Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_Job.PopulateFromRS(1, m_Rs)

      m_FormulaID = m_Job.FORMULA_ID
      Call EnableDisableButton(True)
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
'   Call TabStrip1_Click
   
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing

   Set m_Job = Nothing
   Set m_Jobs = Nothing
   Set m_Employees = Nothing

   Set PartProduct = Nothing
   Set PartProductStock = Nothing
   Set m_PartItems = Nothing
   Set TempCollection = Nothing
End Sub





Private Sub InitFormLayout()
''   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
''   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
''
''   Me.Caption = HeaderText
''   pnlHeader.Caption = HeaderText

'   Call InitNormalLabel(lblJobDesc, MapText("รายละเอียด"))
'
'   Call InitNormalLabel(lblBatchNo, MapText("จำนวนแบต"))
'   Call InitNormalLabel(lblJobApp, MapText("ผู้อนุมัติ"))
'   Call InitNormalLabel(lblJobRes, MapText("ผู้รับผิดชอบ"))
'   Call InitNormalLabel(lblStartJob, MapText("วันที่เริ่มผลิต"))
'   Call InitNormalLabel(lblFinishJob, MapText("วันที่ผลิตเสร็จ"))
'   Call InitNormalLabel(lblJobProcess, MapText("โปรเซส"))
'   Call InitNormalLabel(Label3, MapText("ก.ก."))
'   Call InitNormalLabel(Label4, MapText("ก.ก."))
'   Call InitNormalLabel(lblInputAmount, MapText("ยอดใช้รวม"))
'   Call InitNormalLabel(lblOutputAmount, MapText("ผลิตรวม"))
'
'   Call InitCheckBox(chkCommit, "งานเสร็จแล้ว")
'
'   Call txtJobNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
'   Call txtJobDesc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
'   Call txtBatchNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
'   Call txtInputAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
'   txtInputAmount.Enabled = False
'   Call txtOutputAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
'   txtOutputAmount.Enabled = False
'
'   Call InitCombo(uctlResponseByLookup.MyCombo)
'   Call InitCombo(cboJobProcess)
'
'   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
'
'   pnlHeader.Font.NAME = GLB_FONT
'   pnlHeader.Font.Bold = True
'   pnlHeader.Font.Size = 19
'
'   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
'   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
'   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
'   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
'   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
'   cmdSave.Picture = LoadPicture(glbParameterObj.NormalButton1)
'   cmdPrint.Picture = LoadPicture(glbParameterObj.NormalButton1)
'   cmdAuto.Picture = LoadPicture(glbParameterObj.NormalButton1)
'   cmdCalculate.Picture = LoadPicture(glbParameterObj.NormalButton1)
'
'   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
'   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
'   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
'   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
'   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
'   Call InitMainButton(cmdSave, MapText("บันทึก"))
'   Call InitMainButton(cmdPrint, MapText("พิมพ์"))
'   Call InitMainButton(cmdAuto, MapText("A"))
'   Call InitMainButton(cmdCalculate, MapText("อื่น ๆ"))
'
'   TabStrip1.Font.Bold = True
'   TabStrip1.Font.NAME = GLB_FONT
'   TabStrip1.Font.Size = 16
'   TabStrip1.Tabs.Clear
'   TabStrip1.Tabs.add().Caption = MapText("วัตถุดิบที่ใช้")
'   TabStrip1.Tabs.add().Caption = MapText("ผลิตภัณฑ์ที่ได้")
'   TabStrip1.Tabs.add().Caption = MapText("แรงงาน")
'   TabStrip1.Tabs.add().Caption = MapText("เครื่องจักร")
'   TabStrip1.Tabs.add().Caption = MapText("ค่าใช้จ่ายผลิต")
'   If JobDocType = 1 Then
'      TabStrip1.Tabs.add().Caption = MapText("ตรวจสอบ")
'   End If
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If

   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   OKClick = False
   Call InitFormLayout
  'ucltApproveByLookup.MyCombo.ListIndex = -1
   m_HasActivate = False
   m_HasModify = False

   m_FormulaID = -1
   Set m_Rs = New ADODB.Recordset
   Set m_Job = New CJob
   Set m_Jobs = New Collection
   Set m_Employees = New Collection

   Set PartProduct = New Collection
   Set PartProductStock = New Collection
   Set m_PartItems = New Collection
   Set TempCollection = New Collection

   Set m_ExcelApp = CreateObject("Excel.application")
End Sub

Private Sub PopulateGuiID(Bd As CJob)
Dim Di As CJobInput

   For Each Di In Bd.Inputs
      If Di.Flag = "A" Then
         Di.LINK_ID = GetNextGuiID(Bd)
      End If
   Next Di

   For Each Di In Bd.Outputs
      If Di.Flag = "A" Then
         Di.LINK_ID = GetNextGuiID(Bd)
      End If
   Next Di
End Sub
Private Function GetNextGuiID(Bd As CJob) As Long
Dim Di As CJobInput
Dim MaxId As Long

   MaxId = 0
   For Each Di In Bd.Inputs
      If Di.LINK_ID > MaxId Then
         MaxId = Di.LINK_ID
      End If
   Next Di
   
   For Each Di In Bd.Outputs
      If Di.LINK_ID > MaxId Then
         MaxId = Di.LINK_ID
      End If
   Next Di
   
   GetNextGuiID = MaxId + 1
End Function
Private Sub uctlStartJob_HasChange()
   m_HasModify = True
End Sub


