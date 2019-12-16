VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmImportPlcFormula 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Icon            =   "frmImportPlcFormula.frx":0000
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
      TabIndex        =   6
      Top             =   0
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   6218
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboFormulaType 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   840
         Width           =   2625
      End
      Begin prjFarmManagement.uctlTextBox txtFileName 
         Height          =   435
         Left            =   1860
         TabIndex        =   12
         Top             =   1350
         Width           =   7875
         _ExtentX        =   11986
         _ExtentY        =   767
      End
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   1860
         TabIndex        =   1
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
         TabIndex        =   7
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
         TabIndex        =   2
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
      Begin VB.Label lblFormulaType 
         Alignment       =   1  'Right Justify
         Caption         =   "lblFormulaType"
         Height          =   315
         Left            =   600
         TabIndex        =   14
         Top             =   930
         Width           =   1185
      End
      Begin Threed.SSCommand cmdFileName 
         Height          =   405
         Left            =   9750
         TabIndex        =   13
         Top             =   1350
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportPlcFormula.frx":27A2
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   1890
         TabIndex        =   3
         Top             =   2670
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportPlcFormula.frx":2ABC
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3600
         TabIndex        =   11
         Top             =   2250
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   10
         Top             =   1860
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   9
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label lblFileName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   8
         Top             =   1380
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   8535
         TabIndex        =   5
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
         TabIndex        =   4
         Top             =   2670
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportPlcFormula.frx":2DD6
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmImportPlcFormula"
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

Private FormulaNoColls As Collection

Private m_FormulaCollection As Collection

Private SearchFormulaNo As CFormula
Private MainFormula As CFormula

Private SearchProductNo As CPartItem
Private SearchLocation As CLocation

Private SearchItemNo As CPartItem

Private FormulaCode As String
Private FormulaName As String
Private FormulaDate As String
Private IngredientCode As String
Private IngredientAmount As Double

Private FormulaType As Long
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
   
   If Not VerifyCombo(lblFormulaType, cboFormulaType, False) Then
      Exit Sub
   End If
   
   If Not VerifyTextControl(lblFileName, txtFileName) Then
      Exit Sub
   End If
         
   Call EnableForm(Me, False)
   
   Call ImportPlcFormulaItem
   
   Call EnableForm(Me, True)
   
End Sub

Private Sub ImportPlcFormulaItem()
On Error GoTo ErrorHandler
Dim HasBegin As Boolean
Dim FileName As String
Dim F As Long
Dim TempStr As String
Dim SuccessCount As Long
Dim ErrorCount As Long
Dim Sum As Long
Dim I As Long
Dim LineNo As Long

   Call LoadPartItem(Nothing, PartUctlColls, , , , 1)
   
   Call LoadPartItem(Nothing, PartColls, , , , 2)
   Call LoadPartItem(Nothing, PartPlcColls, , , , 3)
   
   Call LoadLocation(Nothing, LocationColls, 2)
   
   Call LoadDistinctFormulaNo(Nothing, FormulaNoColls)
   
   FormulaType = cboFormulaType.ItemData(Minus2Zero(cboFormulaType.ListIndex))
   
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
   
   LineNo = 0
   SuccessCount = 0
   ErrorCount = 0
   F = FreeFile()
   Close #F
   Open FileName For Input As #F
   While Not EOF(F)
      I = I + 1
      Line Input #F, TempStr
      prgProgress.Value = MyDiff(I, Sum) * 90
      txtPercent.Text = prgProgress.Value
      LineNo = LineNo + 1
      Me.Refresh
      DoEvents
      
      If ProcessLine(TempStr, LineNo) Then
         SuccessCount = SuccessCount + 1
      Else
         ErrorCount = ErrorCount + 1
      End If
   Wend
   Close #F
   
   
   Call glbDatabaseMngr.DBConnection.BeginTrans
   
   HasBegin = True
   
   Dim TempFormula As CFormula
   Dim IsOK As Boolean
   
   For Each TempFormula In m_FormulaCollection
      If Not glbProduction.AddEditFormula(TempFormula, IsOK, False, glbErrorLog) Then
         ErrorCount = ErrorCount + 1
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      End If
   Next TempFormula
   
   prgProgress.Value = 95
   txtPercent.Text = 95
   Me.Refresh
   DoEvents
      
   Dim TempPi As CPartItem
   Dim TempLc As CLocation
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
   
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
End Sub
Private Function ProcessLine(LineStr As String, LineNo As Long) As Boolean
On Error GoTo ErrorHandler
   
   If Trim(LineStr) = "|" Or Trim(LineStr) = "||" Then
      LineNo = 0
      ProcessLine = True
      Exit Function
   End If
   
   If LineNo = 1 Then
      FormulaCode = Trim(LineStr)
   ElseIf LineNo = 2 Then
      FormulaName = Trim(LineStr)
   ElseIf LineNo = 3 Then
      FormulaDate = Trim(LineStr)
      FormulaCode = FormulaCode & "-" & Trim(LineStr)
      
      Set SearchFormulaNo = GetObject("CFormula", FormulaNoColls, Trim(FormulaCode), False)
      If Not SearchFormulaNo Is Nothing Then 'แสดงว่ามีแล้ว ถ้าจะอัพเดดให้ลบของเดิมเองก่อน จะดีกว่า แต่อณุญาติให้ SAVE ได้
         Set MainFormula = Nothing
         ProcessLine = True
         Exit Function
      End If
      
      Set MainFormula = GetObject("CFormula", m_FormulaCollection, Trim(FormulaCode), False)
      If MainFormula Is Nothing Then 'ถ้าไม่มีก็ Set New พร้อมทั้งตั้งค่าของ Formula ก่อน ส่วนถ้ามี Formula แล้วให้สร้าง FormulaInOut อย่างเดียว
         Set MainFormula = New CFormula
         
         MainFormula.FORMULA_ID = -1
         MainFormula.AddEditMode = SHOW_ADD
         MainFormula.FORMULA_DATE = DateSerial(Right(FormulaDate, 2), Mid(FormulaDate, 4, 2), Left(FormulaDate, 2))
         MainFormula.FORMULA_NO = FormulaCode
         MainFormula.FORMULA_DESC = "PLC " & FormulaCode & "-" & FormulaName
         MainFormula.FORMULA_TYPE = FormulaType
         
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
                     glbErrorLog.ShowErrorLog (LOG_TO_FILE)
                     
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
                  glbErrorLog.ShowErrorLog (LOG_TO_FILE)
                  
                  ProcessLine = False
                  Exit Function
               End If
               
               SearchLocation.KEY_ID = SearchProductNo.PART_ITEM_ID
               Call LocationUpdateColls.add(SearchLocation, Trim(SearchProductNo.PART_NO))
            End If
            SearchProductNo.DEFAULT_LOCATION = SearchLocation.LOCATION_ID
         End If
           
         MainFormula.PART_ITEM_ID = SearchProductNo.PART_ITEM_ID
         MainFormula.LOCATION_ID = SearchProductNo.DEFAULT_LOCATION
         
         Call m_FormulaCollection.add(MainFormula, Trim(FormulaCode))
      
      End If
   
            
   ElseIf (LineNo >= 4) And (LineNo <= 6) Then
      ProcessLine = True
      Exit Function
   ElseIf ((LineNo - 6) Mod 3) = 1 Then
      IngredientCode = Trim(LineStr)
   ElseIf ((LineNo - 6) Mod 3) = 2 Then
      ProcessLine = True
      Exit Function
   ElseIf ((LineNo - 6) Mod 3) = 0 Then
      If MainFormula Is Nothing Then
         ProcessLine = True
         Exit Function
      End If
      
      IngredientAmount = Val(Trim(LineStr))
            
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
               frmMapPlcProductItem.HeaderText = MapText("MAP ข้อมูล รหัสวัตถุดิบ " & IngredientCode)
               frmMapPlcProductItem.ShowMode = SHOW_ADD
               Load frmMapPlcProductItem
               frmMapPlcProductItem.Show 1
                  
               OKClick = frmMapPlcProductItem.OKClick
                  
               Unload frmMapPlcProductItem
               Set frmMapPlcProductItem = Nothing
      
               'AddDataTo PartPlcUpdateColls
               If Len(Trim(SearchItemNo.PART_NO)) <= 0 Then
                  glbErrorLog.LocalErrorMsg = "ไม่พบรหัสที่อ้างอิง วัตถุดิบ สำหรับ " & IngredientCode
                  glbErrorLog.ShowErrorLog (LOG_TO_FILE)
                     
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
            frmMapPlcProductLocation.HeaderText = MapText("MAP ข้อมูล สถานที่จัดเก็บ " & IngredientCode)
            frmMapPlcProductLocation.ShowMode = SHOW_ADD
            Load frmMapPlcProductLocation
            frmMapPlcProductLocation.Show 1
            
            OKClick = frmMapPlcProductLocation.OKClick
               
            Unload frmMapPlcProductLocation
            Set frmMapPlcProductLocation = Nothing
   
            'AddDataTo PartPlcUpdateColls
            If Len(Trim(SearchLocation.LOCATION_NO)) <= 0 Then
               glbErrorLog.LocalErrorMsg = "ไม่พบรหัสที่อ้างอิง สถานที่จัดเก็บ สำหรับ " & IngredientCode
               glbErrorLog.ShowErrorLog (LOG_TO_FILE)
               
               ProcessLine = False
               Exit Function
            End If
            
            SearchLocation.KEY_ID = SearchItemNo.PART_ITEM_ID
            Call LocationUpdateColls.add(SearchLocation, Trim(SearchItemNo.PART_NO))
         End If
         SearchItemNo.DEFAULT_LOCATION = SearchLocation.LOCATION_ID
      End If
         
      'สำหรับ FormulaInPut Collection
      Dim MI As CFormulaItem
      Set MI = GetObject("CFormulaInput", MainFormula.Inputs, Trim(Str(SearchItemNo.PART_ITEM_ID)), False)
      If MI Is Nothing Then
         Set MI = New CFormulaItem
         
         MI.Flag = "A"
         MI.PART_ITEM_ID = SearchItemNo.PART_ITEM_ID
         MI.REAL_AMOUNT = IngredientAmount
         MI.LOCATION_ID = SearchItemNo.DEFAULT_LOCATION
         
         ' Add Data To Collection
         Call MainFormula.Inputs.add(MI, Trim(Str(SearchItemNo.PART_ITEM_ID)))
      End If
      
      Call ReArrangeRatio(MainFormula.Inputs)
      
      MainFormula.RMC = MainFormula.RMC + IngredientAmount
      
   End If
   
   ProcessLine = True
   
   Exit Function
ErrorHandler:
   ProcessLine = False
End Function
Public Sub ReArrangeRatio(Col As Collection)
Dim Fi As CFormulaItem
Dim Sum As Double

   Sum = 0
   For Each Fi In Col
      If Fi.Flag <> "D" Then
         Sum = Sum + Fi.REAL_AMOUNT
      End If
   Next Fi

   For Each Fi In Col
      If Fi.Flag <> "D" Then
         Fi.ITEM_PERCENT = MyDiffEx(Fi.REAL_AMOUNT, Sum) * 100

         If Fi.Flag <> "A" Then
            Fi.Flag = "E"
         End If
      End If
   Next Fi
End Sub
Private Sub Form_Activate()
Dim FromDate As Date
Dim ToDate As Date

   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadFormulaType(cboFormulaType)
      
      m_HasModify = False
   End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
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
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitNormalLabel(lblFileName, "ชื่อไฟล์")
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(Label1, "%")
   Call InitNormalLabel(lblFormulaType, MapText("ประเภทสูตร"))
   
   Call InitCombo(cboFormulaType)
   
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
   
   Set FormulaNoColls = New Collection
   
   
   Set m_FormulaCollection = New Collection
   
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
   Set FormulaNoColls = Nothing
   
   Set m_FormulaCollection = Nothing
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

