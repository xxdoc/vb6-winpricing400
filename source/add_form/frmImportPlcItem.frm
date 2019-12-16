VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmImportPlcItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Icon            =   "frmImportPlcItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   10410
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3525
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   6218
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextBox txtFileName 
         Height          =   435
         Left            =   1860
         TabIndex        =   11
         Top             =   1350
         Width           =   7875
         _ExtentX        =   11986
         _ExtentY        =   767
      End
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   1860
         TabIndex        =   0
         Top             =   1800
         Width           =   7785
         _ExtentX        =   13732
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
      Begin Threed.SSCommand cmdFileName 
         Height          =   405
         Left            =   9750
         TabIndex        =   12
         Top             =   1350
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportPlcItem.frx":27A2
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
         MouseIcon       =   "frmImportPlcItem.frx":2ABC
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
      Begin VB.Label lblFileName 
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
         MouseIcon       =   "frmImportPlcItem.frx":2DD6
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmImportPlcItem"
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

Private PartUctlColls As Collection
Private PartColls As Collection
Private PartPlcColls As Collection
Private PartPlcUpdateColls As Collection

Private LocationColls As Collection
Private LocationUpdateColls As Collection

Private JobNoColls As Collection

Private m_JobCollection As Collection

Public ProcessID As Long
Public JobDocType As Long
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

   If Not VerifyTextControl(lblFileName, txtFileName) Then
      Exit Sub
   End If
         
   Call EnableForm(Me, False)
   
   Call ImportPlcProductionItem
   
   Call EnableForm(Me, True)
   
End Sub

Private Sub ImportPlcProductionItem()
On Error GoTo ErrorHandler
Dim HasBegin As Boolean
Dim FileName As String
Dim F As Long
Dim TempStr As String
Dim SuccessCount As Long
Dim ErrorCount As Long
Dim Sum As Long
Dim I As Long

   Call LoadPartItem(Nothing, PartUctlColls, , , , 1)
   
   Call LoadPartItem(Nothing, PartColls, , , , 2)
   Call LoadPartItem(Nothing, PartPlcColls, , , , 3)
   
   Call LoadLocation(Nothing, LocationColls, 2)
   
   Call LoadDistinctJobNo(Nothing, JobNoColls)
   
   
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
   
      
   SuccessCount = 0
   ErrorCount = 0
   F = FreeFile()
   Close #F
   Open FileName For Input As #F
   While Not EOF(F)
      I = I + 1
      Line Input #F, TempStr
      prgProgress.Value = Round(MyDiff(I, Sum) * 90, 2)
      txtPercent.Text = prgProgress.Value
      
      Me.Refresh
      DoEvents
      
      If ProcessLine(TempStr) Then
         SuccessCount = SuccessCount + 1
      Else
         ErrorCount = ErrorCount + 1
      End If
   Wend
   Close #F
   
   Dim TempPi As CPartItem
   Dim TempLc As CLocation
      
      
   If (ErrorCount > 0) Then
      glbErrorLog.LocalErrorMsg = "อิมพอร์ต สำเสร็จจำนวน " & SuccessCount & "และ ล้มเหลวจำนวน " & ErrorCount & " ข้อมูล ระบบจะบันทึกการ MAP เท่านั้น"
      glbErrorLog.ShowUserError
      
      For Each TempPi In PartPlcUpdateColls
         Call TempPi.UpdatePlcPartNo
      Next TempPi
      For Each TempLc In LocationUpdateColls
         Set TempPi = New CPartItem
         TempPi.PART_ITEM_ID = TempLc.KEY_ID
         TempPi.DEFAULT_LOCATION = TempLc.LOCATION_ID
         TempPi.UpdatePlcPartLocation
      Next TempLc
      
      Exit Sub
   End If
   
   Call glbDatabaseMngr.DBConnection.BeginTrans
   
   HasBegin = True
   
   Dim TempJob As CJob
   Dim Ivd As CInventoryDoc
   Dim IsOK As Boolean
   
   For Each TempJob In m_JobCollection
      Call PopulateGuiID(TempJob)
      
      Call glbDaily.Job2InventoryDoc(TempJob, Ivd, 1, 11)
         
      If Not glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog) Then
         ErrorCount = ErrorCount + 1
         glbErrorLog.LocalErrorMsg = " บันทึกเข้า INVENTORY ERROR"
         glbErrorLog.ShowUserError
      End If
      
      TempJob.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID
      
      If Not glbProduction.AddEditJob(TempJob, IsOK, False, glbErrorLog) Then
         ErrorCount = ErrorCount + 1
         glbErrorLog.LocalErrorMsg = " บันทึกเข้า JOB ERROR"
         glbErrorLog.ShowUserError
      End If
   Next TempJob
   
   prgProgress.Value = 95
   txtPercent.Text = 95
   Me.Refresh
   DoEvents
      
   For Each TempPi In PartPlcUpdateColls
      Call TempPi.UpdatePlcPartNo
   Next TempPi
   For Each TempLc In LocationUpdateColls
      Set TempPi = New CPartItem
      TempPi.PART_ITEM_ID = TempLc.KEY_ID
      TempPi.DEFAULT_LOCATION = TempLc.LOCATION_ID
      TempPi.UpdatePlcPartLocation
   Next TempLc
   
   prgProgress.Value = 100
   txtPercent.Text = 100
   Me.Refresh
   glbErrorLog.LocalErrorMsg = "อิมพอร์ต สำเสร็จจำนวน " & SuccessCount & "และ ล้มเหลวจำนวน " & ErrorCount & " ข้อมูล"
   glbErrorLog.ShowUserError
   
   If (ErrorCount > 0) Then
      glbDatabaseMngr.DBConnection.RollbackTrans
   Else
      If ConfirmSave Then
         glbDatabaseMngr.DBConnection.CommitTrans
      Else
         glbDatabaseMngr.DBConnection.RollbackTrans
      End If
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
   
   glbErrorLog.LocalErrorMsg = "Error จากการบันทึกเข้า DATABASE " & Err.DESCRIPTION
   glbErrorLog.ShowUserError
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
Private Function ProcessLine(LineStr As String) As Boolean
On Error GoTo ErrorHandler

Dim TempAsc As Long
Dim OldTempAsc As Long

Dim SearchJobNo As CJob
Dim MainJob As CJob

Dim SearchProductNo As CPartItem
Dim SearchLocation As CLocation

Dim SearchItemNo As CPartItem

Dim PlanCode As String
Dim ProductionDate As String
Dim ProductionNumber As String
Dim BatchNumber As String
Dim FormulaCode As String
Dim FormulaName As String
Dim FormulaDate As String
Dim BatchStartDate As String
Dim BatchEndDate As String
Dim DestinationBin As String
Dim ProductionWeight As Double
Dim TotalBatch As Double
Dim TargetDryMix  As Double
Dim TargetWetMix  As Double
Dim TargetAfterWetMix  As Double
Dim ActualDryMix  As Double
Dim ActualWetMix  As Double
Dim ActualAfterWetMix  As Double
Dim RuningIngredient  As Double
Dim IngredientCode As String
Dim IngredientName  As String
Dim IngredientType As String
Dim BinCode As String
Dim IngredientTargetWeight As String
Dim IngredientActualWeight As String
Dim IngredientDeviationWeight As String

Dim TempDate As String

   OldTempAsc = 1
   PlanCode = StingToVariable2(10, OldTempAsc, LineStr) 'รหัสโรงงาน
   ProductionDate = StingToVariable2(10, OldTempAsc, LineStr) 'วันที่ผลิต
   ProductionNumber = StingToVariable2(10, OldTempAsc, LineStr) 'หมายเลขการผลิต
   
   BatchNumber = StingToVariable2(5, OldTempAsc, LineStr) 'เลขที่ชุดที่ผลิต
   FormulaCode = StingToVariable2(20, OldTempAsc, LineStr)  'รหัสสูตร --> เราใช้เป็นรหัสผลิตภัณฑ์เลย
   
' DateSerial(Right(ProductionDate, 4), Mid(ProductionDate, 4, 2), Left(ProductionDate, 2))
   
   FormulaName = StingToVariable2(50, OldTempAsc, LineStr)  'ชื่อสูตร
   
   FormulaDate = StingToVariable2(10, OldTempAsc, LineStr)  'วันที่สูตร
   BatchStartDate = StingToVariable2(20, OldTempAsc, LineStr) 'วันเวลาเริ่มผลิต
   BatchEndDate = StingToVariable2(20, OldTempAsc, LineStr) 'วันเวลาผลิตเสร็จ
      
   DestinationBin = StingToVariable2(10, OldTempAsc, LineStr) 'รหัสถังปลายทาง
   
   ProductionWeight = StingToVariable2(16, OldTempAsc, LineStr) 'น้ำหนักที่ชั่งจริงรวมทั้งชุด
   TotalBatch = StingToVariable2(5, OldTempAsc, LineStr) 'Total Batch
   TargetDryMix = StingToVariable2(11, OldTempAsc, LineStr) 'Target Dry Mix
   TargetWetMix = StingToVariable2(11, OldTempAsc, LineStr)   'Target Wet Mix
   TargetAfterWetMix = StingToVariable2(11, OldTempAsc, LineStr) 'Target After Wet Mix
   ActualDryMix = StingToVariable2(11, OldTempAsc, LineStr) 'Actual Dry Mix
   ActualWetMix = StingToVariable2(11, OldTempAsc, LineStr) 'Actual Wet Mix
   ActualAfterWetMix = StingToVariable2(11, OldTempAsc, LineStr) 'Actual After Wet Mix
   RuningIngredient = StingToVariable2(2, OldTempAsc, LineStr) 'ลำดับของวัตถุดิบในสูตร
   
   IngredientCode = StingToVariable2(20, OldTempAsc, LineStr)  'รหัสวัตถุดิบ
   IngredientName = StingToVariable2(50, OldTempAsc, LineStr) 'ชื่อวัตถุดิบ
   IngredientType = StingToVariable2(10, OldTempAsc, LineStr)  'ชนิดวัตถุดิบ
   BinCode = StingToVariable2(10, OldTempAsc, LineStr)  'รหัสถังที่ชั่งจริง
   
   IngredientTargetWeight = StingToVariable2(16, OldTempAsc, LineStr)  'นน ที่ต้องการชั่ง
   IngredientActualWeight = StingToVariable2(16, OldTempAsc, LineStr)  'นน ที่ชั่งได้จริง
   IngredientDeviationWeight = StingToVariable2(16, OldTempAsc, LineStr)  'นน Diff

   
   Set SearchJobNo = GetObject("CJob", JobNoColls, Trim(ProductionNumber), False)
   If Not SearchJobNo Is Nothing Then 'แสดงว่ามีแล้ว ถ้าจะอัพเดดให้ลบของเดิมเองก่อน จะดีกว่า
      ProcessLine = False
      glbErrorLog.LocalErrorMsg = " มีข้อมูล JOB " & ProductionNumber & " แล้ว"
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   Set MainJob = GetObject("CJob", m_JobCollection, Trim(ProductionNumber), False)
   If MainJob Is Nothing Then 'ถ้าไม่มีก็ Set New พร้อมทั้งตั้งค่าของ Job ก่อน ส่วนถ้ามี Job แล้วให้สร้าง JobInOut อย่างเดียว
      Set MainJob = New CJob
      MainJob.JOB_ID = -1
      MainJob.AddEditMode = SHOW_ADD
      MainJob.JOB_NO = ProductionNumber
      MainJob.JOB_DESC = "PLC " & FormulaCode & "-" & FormulaName & "-" & FormulaDate
      MainJob.JOB_DATE = DateSerial(Right(ProductionDate, 4), Mid(ProductionDate, 4, 2), Left(ProductionDate, 2))
      MainJob.BATCH_NO = Val(TotalBatch)
      MainJob.START_DATE = DateSerial(Mid(BatchStartDate, 7, 4), Mid(BatchStartDate, 4, 2), Mid(BatchStartDate, 1, 2))
      MainJob.START_DATE = DateAdd("h", Val(Mid(BatchStartDate, 12, 2)), MainJob.START_DATE)
      MainJob.START_DATE = DateAdd("n", Val(Mid(BatchStartDate, 15, 2)), MainJob.START_DATE)
      MainJob.START_DATE = DateAdd("s", Val(Mid(BatchStartDate, 18, 2)), MainJob.START_DATE)
      
      MainJob.FINISH_DATE = DateSerial(Mid(BatchEndDate, 7, 4), Mid(BatchEndDate, 4, 2), Mid(BatchEndDate, 1, 2))
      MainJob.FINISH_DATE = DateAdd("h", Val(Mid(BatchEndDate, 12, 2)), MainJob.START_DATE)
      MainJob.FINISH_DATE = DateAdd("n", Val(Mid(BatchEndDate, 15, 2)), MainJob.START_DATE)
      MainJob.FINISH_DATE = DateAdd("s", Val(Mid(BatchEndDate, 18, 2)), MainJob.START_DATE)
      
      MainJob.PROCESS_ID = ProcessID
      MainJob.COMMIT_FLAG = "N"
      MainJob.JOB_DOC_TYPE = JobDocType
      MainJob.FORMULA_ID = -1
         
'      If MainJob.PART_ITEM_ID = 6555 Then
'         'Debug.Print
'      End If
      ' Search หา จาก FormulaCode ไปยัง PartColls ถ้ายังไม่เจอให้ ไปหาที่ PartPlcColls และถ้ายังไม่เจออีกให้ขึ้น Form มาให้ใส่ แล้ว Save เข้า PartPlcColls และ UpdatePartColls
'      If FormulaCode = "3300" Then
'         'Debug.Print
'      End If
   
      Set SearchProductNo = GetObject("CPartItem", PartColls, Trim(FormulaCode), False)
      If SearchProductNo Is Nothing Then
         Set SearchProductNo = GetObject("CPartItem", PartPlcColls, Trim(FormulaCode), False)
         If SearchProductNo Is Nothing Then
            Set SearchProductNo = GetObject("CPartItem", PartPlcUpdateColls, Trim(FormulaCode), False)
            If SearchProductNo Is Nothing Then
               'LoadForm
               Set SearchProductNo = New CPartItem
               Set frmMapPlcProductItem.PartItem = SearchProductNo
               Set frmMapPlcProductItem.mPartItemColl = PartUctlColls
               frmMapPlcProductItem.HeaderText = MapText("MAP ข้อมูล รหัสผลิตภัณฑ์ " & FormulaCode & "-" & FormulaName)
               frmMapPlcProductItem.ShowMode = SHOW_ADD
               Load frmMapPlcProductItem
               frmMapPlcProductItem.Show 1
               
               OKClick = frmMapPlcProductItem.OKClick
               
               Unload frmMapPlcProductItem
               Set frmMapPlcProductItem = Nothing
   
               'AddDataTo PartPlcUpdateColls
               If Len(Trim(SearchProductNo.PART_NO)) <= 0 Then
                  glbErrorLog.LocalErrorMsg = "ไม่พบรหัสที่อ้างอิง ผลิตภัณฑ์ สำหรับ " & FormulaCode & "-" & FormulaName
                  glbErrorLog.ShowUserError
                  
                  ProcessLine = False
                  Exit Function
               End If
               SearchProductNo.NUMBER_PLC_ID = Trim(FormulaCode)
               Call PartPlcUpdateColls.add(SearchProductNo, Trim(FormulaCode))
            End If
         End If
      End If
      
      'เช็คต่อว่ามี Default Location หรือยัง
      If SearchProductNo.DEFAULT_LOCATION <= 0 Then
         Set SearchLocation = GetObject("CLocation", LocationUpdateColls, Trim(SearchProductNo.PART_NO), False)
         If SearchLocation Is Nothing Then
            'LoadForm
            Set SearchLocation = New CLocation
            Set frmMapPlcProductLocation.Location = SearchLocation
            Set frmMapPlcProductLocation.mLocationColl = LocationColls
            frmMapPlcProductLocation.HeaderText = MapText("MAP ข้อมูล สถานที่จัดเก็บ " & FormulaCode & "-" & FormulaName)
            frmMapPlcProductLocation.ShowMode = SHOW_ADD
            Load frmMapPlcProductLocation
            frmMapPlcProductLocation.Show 1
            
            OKClick = frmMapPlcProductLocation.OKClick
            
            Unload frmMapPlcProductLocation
            Set frmMapPlcProductLocation = Nothing

            'AddDataTo PartPlcUpdateColls
            If Len(Trim(SearchLocation.LOCATION_NO)) <= 0 Then
               glbErrorLog.LocalErrorMsg = "ไม่พบรหัสที่อ้างอิง สถานที่จัดเก็บ สำหรับ " & FormulaCode & "-" & FormulaName
               glbErrorLog.ShowUserError
               
               ProcessLine = False
               Exit Function
            End If
            
            SearchLocation.KEY_ID = SearchProductNo.PART_ITEM_ID
            Call LocationUpdateColls.add(SearchLocation, Trim(SearchProductNo.PART_NO))
         End If
         SearchProductNo.DEFAULT_LOCATION = SearchLocation.LOCATION_ID
      End If
        
      MainJob.PART_ITEM_ID = SearchProductNo.PART_ITEM_ID
      MainJob.STD_AMOUNT = 0          'เดี่ยวรอคำนวณใหม่จาก Input
      MainJob.ACTUAL_AMOUNT = 0 'เดี่ยวรอคำนวณใหม่จาก Input
      
      
      'สำหรับ JobOutPut Collection
      Dim Ma As CJobInput
      Set Ma = New CJobInput
   
      Ma.Flag = "A"
      Ma.PART_ITEM_ID = SearchProductNo.PART_ITEM_ID
      Ma.TX_AMOUNT = 0 'เดี่ยวรอคำนวณใหม่จาก Input
      Ma.LOCATION_ID = SearchProductNo.DEFAULT_LOCATION
      Ma.SERIAL_NUMBER = ""
      Ma.INOUT_REF = ""
      Ma.STD_AMOUNT = 0 'เดี่ยวรอคำนวณใหม่จาก Input
      Ma.WEIGHT_PER_PACK = 0
      Ma.PACK_AMOUNT = 0
      Ma.TX_TYPE = "I"
      Call MainJob.Outputs.add(Ma, Trim(str(SearchProductNo.PART_ITEM_ID)))
      
      Call m_JobCollection.add(MainJob, Trim(ProductionNumber))
      
      Set Ma = Nothing
   End If
   
'   If MainJob.PART_ITEM_ID = 6555 Then
'      'Debug.Print
'   End If
   
   ' Input
   Set SearchItemNo = GetObject("CPartItem", PartColls, Trim(IngredientCode), False)
   If SearchItemNo Is Nothing Then
      Set SearchItemNo = GetObject("CPartItem", PartPlcColls, Trim(IngredientCode), False)
      If SearchItemNo Is Nothing Then
         Set SearchItemNo = GetObject("CPartItem", PartPlcUpdateColls, Trim(IngredientCode), False)
         If SearchItemNo Is Nothing Then
            'LoadForm
            Set SearchItemNo = New CPartItem
            Set frmMapPlcProductItem.PartItem = SearchItemNo
            Set frmMapPlcProductItem.mPartItemColl = PartUctlColls
            frmMapPlcProductItem.HeaderText = MapText("MAP ข้อมูล รหัสวัตถุดิบ " & IngredientCode & "-" & IngredientName)
            frmMapPlcProductItem.ShowMode = SHOW_ADD
            Load frmMapPlcProductItem
            frmMapPlcProductItem.Show 1
               
            OKClick = frmMapPlcProductItem.OKClick
               
            Unload frmMapPlcProductItem
            Set frmMapPlcProductItem = Nothing
   
            'AddDataTo PartPlcUpdateColls
            If Len(Trim(SearchItemNo.PART_NO)) <= 0 Then
               glbErrorLog.LocalErrorMsg = "ไม่พบรหัสที่อ้างอิง วัตถุดิบ สำหรับ " & IngredientCode & "-" & IngredientName
               glbErrorLog.ShowUserError
                  
               ProcessLine = False
               Exit Function
            End If
            SearchItemNo.NUMBER_PLC_ID = Trim(IngredientCode)
            Call PartPlcUpdateColls.add(SearchItemNo, Trim(IngredientCode))
         End If
      End If
   End If
   'เช็คต่อว่ามี Default Location หรือยัง
   If SearchItemNo.DEFAULT_LOCATION <= 0 Then
      Set SearchLocation = GetObject("CLocation", LocationUpdateColls, Trim(SearchItemNo.PART_NO), False)
      If SearchLocation Is Nothing Then
         'LoadForm
         Set SearchLocation = New CLocation
         Set frmMapPlcProductLocation.Location = SearchLocation
         Set frmMapPlcProductLocation.mLocationColl = LocationColls
         frmMapPlcProductLocation.HeaderText = MapText("MAP ข้อมูล สถานที่จัดเก็บ " & IngredientCode & "-" & IngredientName)
         frmMapPlcProductLocation.ShowMode = SHOW_ADD
         Load frmMapPlcProductLocation
         frmMapPlcProductLocation.Show 1
         
         OKClick = frmMapPlcProductLocation.OKClick
            
         Unload frmMapPlcProductLocation
         Set frmMapPlcProductLocation = Nothing

         'AddDataTo PartPlcUpdateColls
         If Len(Trim(SearchLocation.LOCATION_NO)) <= 0 Then
            glbErrorLog.LocalErrorMsg = "ไม่พบรหัสที่อ้างอิง สถานที่จัดเก็บ สำหรับ " & IngredientCode & "-" & IngredientName
            glbErrorLog.ShowUserError
            
            ProcessLine = False
            Exit Function
         End If
         
         SearchLocation.KEY_ID = SearchItemNo.PART_ITEM_ID
         Call LocationUpdateColls.add(SearchLocation, Trim(SearchItemNo.PART_NO))
      End If
      SearchItemNo.DEFAULT_LOCATION = SearchLocation.LOCATION_ID
   End If
      
   'สำหรับ JobInPut Collection
   Dim MI As CJobInput
   Set MI = GetObject("CJobInput", MainJob.Inputs, Trim(str(SearchItemNo.PART_ITEM_ID)), False)
   If MI Is Nothing Then
      Set MI = New CJobInput
      
      MI.Flag = "A"
      MI.PART_ITEM_ID = SearchItemNo.PART_ITEM_ID
      MI.TX_AMOUNT = Val(IngredientActualWeight)
      MI.LOCATION_ID = SearchItemNo.DEFAULT_LOCATION
      MI.FROM_FORMULA = -1
      MI.TX_TYPE = "E"
      MI.AVG_PRICE = 0
      MI.GROUP_NO = 0
      MI.MIX_DATE = MainJob.START_DATE
      MI.STD_AMOUNT = Val(IngredientTargetWeight)
      MI.PARAM_ID = -1
      
      ' Add Data To Collection
      Call MainJob.Inputs.add(MI, Trim(str(SearchItemNo.PART_ITEM_ID)))
   Else
      MI.TX_AMOUNT = MI.TX_AMOUNT + Val(IngredientActualWeight)
      MI.STD_AMOUNT = MI.STD_AMOUNT + Val(IngredientTargetWeight)
   End If
   
   MainJob.STD_AMOUNT = MainJob.STD_AMOUNT + Val(IngredientTargetWeight)
   MainJob.ACTUAL_AMOUNT = MainJob.ACTUAL_AMOUNT + Val(IngredientActualWeight)
   MainJob.BATCH_NO = Val(BatchNumber)
   
   Set Ma = GetObject("CJobInput", MainJob.Outputs, Trim(str(MainJob.PART_ITEM_ID)), False)
   If Not Ma Is Nothing Then
      Ma.TX_AMOUNT = MainJob.ACTUAL_AMOUNT
      Ma.STD_AMOUNT = MainJob.STD_AMOUNT
   End If
   
   ProcessLine = True
   
   Exit Function
ErrorHandler:
   ProcessLine = False
   glbErrorLog.LocalErrorMsg = "Runtime error. At ProductionNumber = " & ProductionNumber & " BatchNo = " & BatchNumber
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
End Function

Private Sub Form_Activate()
Dim FromDate As Date
Dim ToDate As Date

   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      
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
   
   Call InitNormalLabel(lblFileName, "ชื่อไฟล์")
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
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
   
   Call ResetStatus
End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub
Private Sub Form_Load()
   Call EnableForm(Me, False)
   m_HasActivate = False
   
   Set m_Rs = New ADODB.Recordset
   
   Set PartUctlColls = New Collection
   Set PartColls = New Collection
   Set PartPlcColls = New Collection
   Set PartPlcUpdateColls = New Collection
   
   Set LocationColls = New Collection
   Set LocationUpdateColls = New Collection
   
   Set JobNoColls = New Collection
   
   
   Set m_JobCollection = New Collection
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set PartUctlColls = Nothing
   Set PartColls = Nothing
   Set PartPlcColls = Nothing
   Set PartPlcUpdateColls = Nothing
   
   Set LocationColls = Nothing
   Set LocationUpdateColls = Nothing
   Set JobNoColls = Nothing
   
   Set m_JobCollection = Nothing
End Sub
Private Function StingToVariable(TempAsc As Long, OldTempAsc As Long, LineStr As String) As Variant
   TempAsc = InStr(TempAsc + 1, LineStr, ";")
   StingToVariable = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)
   OldTempAsc = TempAsc
End Function
Private Function StingToVariable2(TempAsc As Long, OldTempAsc As Long, LineStr As String) As Variant
   While (Asc(Mid(LineStr, OldTempAsc, 1)) = 32) '32 = ช่องว่าง
      OldTempAsc = OldTempAsc + 1
   Wend
   StingToVariable2 = Trim(Mid(LineStr, OldTempAsc, TempAsc))
   OldTempAsc = OldTempAsc + TempAsc
End Function

