VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmImportPlcItemNew 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   20355
   Icon            =   "frmImportPlcItemNew.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   20355
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   7125
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   20385
      _ExtentX        =   35957
      _ExtentY        =   12568
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlDate uctlDateSel 
         Height          =   495
         Left            =   7200
         TabIndex        =   15
         Top             =   870
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   873
      End
      Begin prjFarmManagement.uctlTextBox txtFileName 
         Height          =   435
         Left            =   1440
         TabIndex        =   11
         Top             =   870
         Width           =   5535
         _ExtentX        =   14843
         _ExtentY        =   767
      End
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   2340
         TabIndex        =   0
         Top             =   5520
         Width           =   7785
         _ExtentX        =   13732
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   20355
         _ExtentX        =   35904
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtPercent 
         Height          =   465
         Left            =   2340
         TabIndex        =   1
         Top             =   5850
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   820
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   13080
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin GridEX20.GridEX GridEX2 
         Height          =   3735
         Left            =   5160
         TabIndex        =   13
         Top             =   1560
         Width           =   15015
         _ExtentX        =   26485
         _ExtentY        =   6588
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         TabKeyBehavior  =   1
         MethodHoldFields=   -1  'True
         AllowColumnDrag =   0   'False
         AllowEdit       =   0   'False
         BorderStyle     =   3
         GroupByBoxVisible=   0   'False
         DataMode        =   99
         HeaderFontName  =   "AngsanaUPC"
         HeaderFontBold  =   -1  'True
         HeaderFontSize  =   14.25
         HeaderFontWeight=   700
         FontSize        =   9.75
         BackColorBkg    =   16777215
         ColumnHeaderHeight=   480
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmImportPlcItemNew.frx":27A2
         Column(2)       =   "frmImportPlcItemNew.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmImportPlcItemNew.frx":290E
         FormatStyle(2)  =   "frmImportPlcItemNew.frx":2A6A
         FormatStyle(3)  =   "frmImportPlcItemNew.frx":2B1A
         FormatStyle(4)  =   "frmImportPlcItemNew.frx":2BCE
         FormatStyle(5)  =   "frmImportPlcItemNew.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmImportPlcItemNew.frx":2D5E
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   3735
         Left            =   120
         TabIndex        =   17
         Top             =   1560
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   6588
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         TabKeyBehavior  =   1
         MethodHoldFields=   -1  'True
         AllowColumnDrag =   0   'False
         AllowEdit       =   0   'False
         BorderStyle     =   3
         GroupByBoxVisible=   0   'False
         DataMode        =   99
         HeaderFontName  =   "AngsanaUPC"
         HeaderFontBold  =   -1  'True
         HeaderFontSize  =   14.25
         HeaderFontWeight=   700
         FontSize        =   9.75
         BackColorBkg    =   16777215
         ColumnHeaderHeight=   480
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmImportPlcItemNew.frx":2F36
         Column(2)       =   "frmImportPlcItemNew.frx":2FFE
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmImportPlcItemNew.frx":30A2
         FormatStyle(2)  =   "frmImportPlcItemNew.frx":31FE
         FormatStyle(3)  =   "frmImportPlcItemNew.frx":32AE
         FormatStyle(4)  =   "frmImportPlcItemNew.frx":3362
         FormatStyle(5)  =   "frmImportPlcItemNew.frx":343A
         ImageCount      =   0
         PrinterProperties=   "frmImportPlcItemNew.frx":34F2
      End
      Begin Threed.SSCommand cmdSelect 
         Height          =   405
         Left            =   4560
         TabIndex        =   19
         Top             =   3000
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportPlcItemNew.frx":36CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSelectAll 
         Height          =   405
         Left            =   4560
         TabIndex        =   18
         Top             =   3600
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportPlcItemNew.frx":39E4
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   12000
         TabIndex        =   16
         Top             =   5400
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportPlcItemNew.frx":3CFE
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdRunAuto 
         Height          =   405
         Left            =   11400
         TabIndex        =   14
         Top             =   870
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportPlcItemNew.frx":4018
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdFileName 
         Height          =   405
         Left            =   14880
         TabIndex        =   12
         Top             =   840
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportPlcItemNew.frx":4332
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   2340
         TabIndex        =   2
         Top             =   6360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportPlcItemNew.frx":464C
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   4080
         TabIndex        =   10
         Top             =   5970
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   690
         TabIndex        =   9
         Top             =   5580
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   690
         TabIndex        =   8
         Top             =   6000
         Width           =   1575
      End
      Begin VB.Label lblFileName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   330
         TabIndex        =   7
         Top             =   900
         Width           =   975
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   13680
         TabIndex        =   4
         Top             =   6360
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   12000
         TabIndex        =   3
         Top             =   6360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportPlcItemNew.frx":4966
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmImportPlcItemNew"
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
Private JobNoColls2 As Collection
Private JobNoColls3 As Collection
Private m_JobCollection As Collection
Private m_CollLotItemWh As Collection
Public TempCollection3 As Collection
Public m_CollBin As Collection
Public m_CollList1 As Collection
Public m_CollList2  As Collection
Public LotColls As Collection
Public ListPartName As Collection

Private PartItemID As Long

Public ProcessID As Long
Public JobDocType As Long
Public StartJob As Date
Public StopJob As Date
Public SplitFlag As Boolean
Public JobNo As String
Private Lt As cLot
Dim ItemCount As Long
Dim strDate As String
Dim isRunFirst As Boolean
Dim TempBatchNumber As Long

Dim SearchJobNo As CJob
Dim SearchJobNo2 As CJob
Dim strDateSel As String

Private Type BROWSEINFO
  hOwner As Long
  pidlRoot As Long
  pszDisplayName As String
  lpszTitle As String
  ulFlags As Long
  lpfn As Long
  lParam As Long
  iImage As Long
End Type
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Const BIF_RETURNONLYFSDIRS = &H1

Private Sub cboLotNo_Click()
   m_HasModify = True
End Sub

Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
'Dim TempBatch As CBacthing

   If Not cmdDelete.Enabled Then
      Exit Sub
   End If
   
   If SplitFlag Then 'GridEX2.Value(13) <> "Sp" And SplitFlag
      Exit Sub
   End If

   If Not VerifyGrid(GridEX2.Value(1)) Then
      Exit Sub
   End If
   ID = GridEX2.Value(2)
   
   Call EnableForm(Me, False)
   
   If ID <= 0 Then
      m_CollList2.Remove (ID)
   Else
         m_CollList2.Item(ID).Flag = "D"
   End If
   
   GridEX2.ItemCount = CountItem(m_CollList2)
   GridEX2.Rebind
   m_HasModify = True
   Call EnableForm(Me, True)
End Sub

Private Sub cmdFileName_Click()
On Error Resume Next
'Dim strDescription As String
   m_HasModify = True
   
''   'edit the filter to support more image types
'   dlgAdd.Filter = "Text Files (*.TXT)|*..txt;*.TXT;"
'   dlgAdd.Filter = ""
'   dlgAdd.fileName = "t"
'   dlgAdd.DialogTitle = "Select access file to import"
'   dlgAdd.ShowOpen
'   If dlgAdd.fileName = "" Then
'      Exit Sub
'   End If
''
    Dim FileName As String
    FileName = BrowseFolder("Select a folder")
    If FileName <> "" Then
         txtFileName.Text = FileName
         glbParameterObj.PartImportPLC = FileName
    End If

  Call checkData(FileName)
   m_HasModify = True
End Sub
Private Function checkData(FileName As String)
   If Not ListFolder(FileName) Then
      Exit Function
   End If
   
   ItemCount = 0
   Call SetNothing
   Call SetNew
   
   Call LoadDistinctJobNo(Nothing, JobNoColls, 4)
   Call loadFileToList
   
   If ItemCount = 0 Then
      glbErrorLog.LocalErrorMsg = "ไม่มีรายการคงเหลือให้ต้อง อิมพอร์ท ของ วันที่ " & uctlDateSel.ShowDate & " "
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   GridEX1.ItemCount = CountItem(m_CollList1)
   GridEX1.Rebind
End Function
Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If

   SaveData = False
End Function

Private Sub cmdRunAuto_Click()
   Dim FileName As String
   GridEX1.ItemCount = 0
   GridEX1.Rebind
   GridEX2.ItemCount = 0
   GridEX2.Rebind
   strDateSel = Format(Year(uctlDateSel.ShowDate), "0000") & Format(Month(uctlDateSel.ShowDate), "00") & Format(Day(uctlDateSel.ShowDate), "00")
    txtFileName.Text = glbParameterObj.PartImportPLC & "\" & strDateSel
   FileName = glbParameterObj.PartImportPLC
   Call checkData(txtFileName.Text)
   m_HasModify = True
End Sub



Private Sub cmdSelect_Click()
Dim TempID As Long
   m_HasModify = True
   
   TempID = GridEX1.row
 
   Call CopyItem(m_CollList1, m_CollList2, TempID)

   GridEX1.ItemCount = CountItem(m_CollList1)
   GridEX1.Rebind
   
   GridEX2.ItemCount = CountItem(m_CollList2)
   GridEX2.Rebind
End Sub

Private Sub cmdSelectAll_Click()
   m_HasModify = True
   Call CopyAllItem(m_CollList1, m_CollList2)

   GridEX1.ItemCount = CountItem(m_CollList1) 'm_CollList1.Count
   GridEX1.Rebind

   GridEX2.ItemCount = CountItem(m_CollList2) 'm_CollList2.Count
   GridEX2.Rebind
End Sub
Public Sub CopyItem(TempCol1 As Collection, TempCol2 As Collection, ID As Long)
Dim L As CUserAccount
Dim strDate As Date
   If ID > 0 Then
      strDate = DateSerial(Val(Mid(TempCol1(ID).BatchStartDate, 7, 4)), Val(Mid(TempCol1(ID).BatchStartDate, 4, 2)), Val(Mid(TempCol1(ID).BatchStartDate, 1, 2)))
      If Not VerifyLockInventoryDate(strDate, strDate) Then
         glbErrorLog.LocalErrorMsg = MapText("ไม่สามารถเปลี่ยนแปลงเอกสารตามวันที่เอกสารที่เลือกได้ กรุณาติดต่อผู้ดูแลระบบ หรือผู้มีสิทธิ์กำหนดวันที่เอกสารได้")
         glbErrorLog.ShowUserError
         Exit Sub
      End If
       
      Call TempCol2.add(TempCol1(ID), Trim(TempCol1(ID).TempProductionNumber))
      TempCol1.Remove (ID)
   End If
End Sub
Public Sub CopyAllItem(TempCol1 As Collection, TempCol2 As Collection)
Dim J As Long
Dim strDate As Date
Dim C As Long
Dim tempBatch As CBacthing


For Each tempBatch In TempCol1
'   strDate = DateToStringIntLow(DateSerial(Mid(tempBatch.BatchStartDate, 7, 4), Mid(tempBatch.BatchStartDate, 4, 2), Mid(tempBatch.BatchStartDate, 1, 2)))
   strDate = DateSerial(Val(Mid(tempBatch.BatchStartDate, 7, 4)), Val(Mid(tempBatch.BatchStartDate, 4, 2)), Val(Mid(tempBatch.BatchStartDate, 1, 2)))
     If VerifyLockInventoryDate(strDate, strDate) Then
          Call TempCol2.add(tempBatch, Trim(tempBatch.TempProductionNumber))
      End If
Next

   
'   For J = 1 To C 'TempCol1.Count
'      strDate = InternalDateToDateExGrid(DateSerial(Right(TempCol1(J).ProductionDate, 4), Mid(TempCol1(J).ProductionDate, 4, 2), Left(TempCol1(J).ProductionDate, 2)))
'      If Not VerifyLockInventoryDate(strDate, strDate) Then
'          Call TempCol2.add(TempCol1(J), Trim(TempCol1(J).TempProductionNumber))
''          TempCol1.Remove (J)
'      End If
'   Next J
   Set TempCol1 = Nothing
   Set TempCol1 = New Collection
End Sub
Private Sub cmdStart_Click()
Dim TempID As Long
Dim tempBatch As CBacthing
Dim ID As Long
Dim SearchBacthing As CBacthing

   If Not VerifyTextControl(lblFileName, txtFileName) Then
      Exit Sub
   End If
   
   If m_CollList2 Is Nothing Or m_CollList2.Count = 0 Then
      Exit Sub
   End If
   
   If m_CollList2.Count > 0 Then
   ID = 0
       For Each tempBatch In m_CollList2
         ID = ID + 1
         If tempBatch.SKIP_PART_ITEM_NO = False Then  'ถ้าไม่ใช่รำล้างไลน์ หรือ ข้าวโพดล้างไลน์'If TempBatch.FormulaCode <> "10541" And TempBatch.FormulaCode <> "10101" Then  'ถ้าไม่ใช่รำล้างไลน์ หรือ ข้าวโพดล้างไลน์
            If Len(tempBatch.LotNo) = 0 And tempBatch.Flag <> "D" Then
               glbErrorLog.LocalErrorMsg = "การผลิตอาหารเบอร์ " & tempBatch.FormulaCode & " ยังไม่มีเลข Lot การผลิต กรุณาป้อนเลข Lot การผลิต"
               glbErrorLog.ShowUserError
               Call ShowLot(ID)
               Exit Sub
            End If
         End If
       Next tempBatch
   End If
   Call EnableForm(Me, False)
   
   Call LoadPartItem(Nothing, PartUctlColls, , , , 1)
   Call LoadPartItem(Nothing, PartColls, , , , 2)
   Call LoadPartItem(Nothing, PartPlcColls, , , , 3)
   Call LoadLocation(Nothing, LocationColls, 2)

Dim IsOK As Boolean
   If SplitFlag Then
     For Each SearchBacthing In m_CollList2
      If SearchBacthing.SplitFlag = "Sp" And SearchBacthing.Flag <> "D" Then
         Call ImportPlcProductionItem
      ElseIf SearchBacthing.SplitFlag = "Sp" And SearchBacthing.Flag = "D" Then
         Call cmdOK_Click
      ElseIf SearchBacthing.SplitFlag = "" Then
          If Not glbProduction.DeleteJobSplit(-1, IsOK, True, glbErrorLog, SearchBacthing.TempProductionNumber) Then
            glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
            Exit Sub
         End If
         Call ImportPlcProductionItem2
         If m_CollList2.Count = 1 Then
            Call cmdOK_Click
         End If
      End If
     Next SearchBacthing
   Else
      Call ImportPlcProductionItem
   End If
          
   Call EnableForm(Me, True)
   OKClick = True
End Sub
Function Clear()
   txtFileName.Text = ""
   prgProgress.Value = 0
   txtPercent.Text = ""
End Function
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
Dim ExitDo As Boolean
Dim tempBatch As CBacthing
Dim TempDate As Date

Dim TempJob As CJob
Dim Ivd As CInventoryDoc
Dim IvdWH As CInventoryWHDoc
Dim IsOK As Boolean
Dim File As File
Dim refSum As Long

Dim TempPi As CPartItem
Dim TempLc As CLocation

   HasBegin = True
   Call glbDatabaseMngr.DBConnection.BeginTrans
   
   SuccessCount = 0
   ErrorCount = 0
   
   Sum = 0
   For Each File In ListPartName
    FileName = File
   F = FreeFile()
   Close #F
   Open FileName For Input As #F
      While Not EOF(F)
         Line Input #F, TempStr
         Sum = Sum + 1
      Wend
   Next File
   
   prgProgress.MAX = 100
   prgProgress.MIN = 0
   prgProgress.Value = 0
   txtPercent.Text = ""
   I = 0
   Dim strFileName As String
   For Each File In ListPartName
    I = I + 1
    
   Set Ivd = New CInventoryDoc
   Set IvdWH = New CInventoryWHDoc
   Set m_JobCollection = New Collection
   Set JobNoColls2 = New Collection
   
   strFileName = "BK" & Mid(File.NAME, 1, 9)
   Call LoadJobByJobNo(Nothing, m_JobCollection, , , 2, Trim(strFileName))
   
     If m_CollList2.Count > 0 Then
    strDate = "('"
      For Each tempBatch In m_CollList2
         If strDate <> "('" Then
           If DateToStringIntLow(Trim(DateSerial(Mid(TempStr, 7, 4), Mid(TempStr, 4, 2), Mid(TempStr, 1, 2)))) <> DateToStringIntLow(Trim(DateSerial(Mid(tempBatch.BatchStartDate, 7, 4), Mid(tempBatch.BatchStartDate, 4, 2), Mid(tempBatch.BatchStartDate, 1, 2)))) Then
               strDate = strDate & "','" & DateToStringIntLow(Trim(DateSerial(Mid(tempBatch.BatchStartDate, 7, 4), Mid(tempBatch.BatchStartDate, 4, 2), Mid(tempBatch.BatchStartDate, 1, 2))))
              TempStr = Trim(tempBatch.BatchStartDate)
            End If
         Else
            strDate = strDate & DateToStringIntLow(Trim(DateSerial(Mid(tempBatch.BatchStartDate, 7, 4), Mid(tempBatch.BatchStartDate, 4, 2), Mid(tempBatch.BatchStartDate, 1, 2))))
            TempStr = Trim(tempBatch.BatchStartDate)
         End If
      Next tempBatch
      strDate = strDate & "')"
   Else
     strDate = ""
   End If

   If strDate = "" Then
      Call LoadLotFromLot(Nothing, LotColls, , , , 2, , 4)
   Else
      Call LoadLotFromLot(Nothing, LotColls, , , , 2, , 4, , , strDate)
   End If
    
   FileName = File
   F = FreeFile()
   Close #F
   Open FileName For Input As #F
   While (Not EOF(F)) And (ExitDo = False)
      I = I + 1
      Line Input #F, TempStr
      prgProgress.Value = ROUND(MyDiff(I, Sum) * 90, 2)
      txtPercent.Text = prgProgress.Value

      Me.Refresh
      DoEvents
     
      If ProcessLine(TempStr, refSum) Then
         SuccessCount = SuccessCount + 1 + refSum
      Else
         If MsgBox("ต้องการจะทำการ อิมพอร์ทข้อมูลชุดนี้ต่อหรือไม่", vbOKCancel, "แจ้งเตือน") = vbCancel Then
             ExitDo = True
             glbDatabaseMngr.DBConnection.RollbackTrans
         End If
         ErrorCount = ErrorCount + 1
      End If
   Wend
   Close #F
      
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
   
   HasBegin = True
   
      For Each TempJob In m_JobCollection
         Call PopulateGuiID(TempJob)
         Call glbDaily.Job2InventoryDoc(TempJob, Ivd, 1, 11)
         If Not glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog) Then
            ErrorCount = ErrorCount + 1
            glbErrorLog.LocalErrorMsg = " บันทึกเข้า INVENTORY ERROR"
            glbErrorLog.ShowUserError
            
            If MsgBox("ต้องการจะทำการ อิมพอร์ทข้อมูลชุดนี้ต่อหรือไม่", vbOKCancel, "แจ้งเตือน") = vbCancel Then
                ExitDo = True
                glbDatabaseMngr.DBConnection.RollbackTrans
            End If
         End If
         TempJob.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID
         
         'เข้าคลัง
         Call glbDaily.Job2InventoryWhDoc(TempJob, IvdWH, 1, 11, 1)
         If Not glbDaily.AddEditInventoryWhDoc(IvdWH, IsOK, False, glbErrorLog) Then
            ErrorCount = ErrorCount + 1
            glbErrorLog.LocalErrorMsg = " บันทึกเข้า INVENTORY WH ERROR"
            glbErrorLog.ShowUserError
            
            If MsgBox("ต้องการจะทำการ อิมพอร์ทข้อมูลชุดนี้ต่อหรือไม่", vbOKCancel, "แจ้งเตือน") = vbCancel Then
                ExitDo = True
                glbDatabaseMngr.DBConnection.RollbackTrans
            End If
         End If
         
         TempJob.INVENTORY_WH_DOC_ID = IvdWH.INVENTORY_WH_DOC_ID
         If Not glbProduction.AddEditJob(TempJob, IsOK, False, glbErrorLog) Then
            ErrorCount = ErrorCount + 1
            glbErrorLog.LocalErrorMsg = " บันทึกเข้า JOB ERROR"
            glbErrorLog.ShowUserError
            If MsgBox("ต้องการจะทำการ อิมพอร์ทข้อมูลชุดนี้ต่อหรือไม่", vbOKCancel, "แจ้งเตือน") = vbCancel Then
                ExitDo = True
                glbDatabaseMngr.DBConnection.RollbackTrans
            End If
         End If
      Next TempJob
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
      Me.Refresh
   Next File
   
   prgProgress.Value = 100
   txtPercent.Text = 100
   
   glbErrorLog.LocalErrorMsg = "อิมพอร์ต สำเร็จจำนวน " & SuccessCount & "และ ล้มเหลวจำนวน " & ErrorCount & " ข้อมูล"
   glbErrorLog.ShowUserError
   
   If (ErrorCount > 0) Then
        glbDatabaseMngr.DBConnection.RollbackTrans
     Else
        If ConfirmSave Then
           glbDatabaseMngr.DBConnection.CommitTrans
        
           glbErrorLog.LocalErrorMsg = "บันทึกข้อมูลเสร็จเรียบร้อยแล้ว"
           glbErrorLog.ShowUserError
           m_HasModify = False
                  
           Call cmdOK_Click
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
Private Sub ImportPlcProductionItem2()
On Error GoTo ErrorHandler
Dim HasBegin As Boolean
Dim FileName As String
Dim F As Long
Dim TempStr As String
Dim SuccessCount As Long
Dim ErrorCount As Long
Dim Sum As Long
Dim I As Long
Dim ExitDo As Boolean
Dim tempBatch As CBacthing
Dim TempDate As Date

Dim TempJob As CJob
Dim Ivd As CInventoryDoc
Dim IvdWH As CInventoryWHDoc
Dim IsOK As Boolean
Dim File As File
Dim refSum As Long

Dim TempPi As CPartItem
Dim TempLc As CLocation
Dim Find As Boolean

'   Call LoadPartItem(Nothing, PartUctlColls, , , , 1)
'   Call LoadPartItem(Nothing, PartColls, , , , 2)
'   Call LoadPartItem(Nothing, PartPlcColls, , , , 3)
'   Call LoadLocation(Nothing, LocationColls, 2)
   
   HasBegin = True
   Call glbDatabaseMngr.DBConnection.BeginTrans
   
   SuccessCount = 0
   ErrorCount = 0
   
   Sum = 0
   For Each File In ListPartName
    FileName = File
   F = FreeFile()
   Close #F
   Open FileName For Input As #F
      While Not EOF(F)
         Line Input #F, TempStr
         Sum = Sum + 1
      Wend
   Next File
   
   prgProgress.MAX = 100
   prgProgress.MIN = 0
   prgProgress.Value = 0
   txtPercent.Text = ""
   I = 0
   Dim strFileName As String
   For Each File In ListPartName
    I = I + 1
    
   Set Ivd = New CInventoryDoc
   Set m_JobCollection = New Collection
   Set JobNoColls2 = New Collection
   
   strFileName = "BK" & Mid(File.NAME, 1, 9)
   
   Find = False
   If m_CollList2.Count > 0 Then
    For Each tempBatch In m_CollList2
         'หาตัวที่ไม่ใช่ File ทีต้องการ ก็ให้ออกไป
         If Mid(tempBatch.TempProductionNumber, 1, 11) = strFileName Then
            Find = True
            Exit For
         End If
      Next tempBatch
   End If
   
   If Find Then
   
   Call LoadJobByJobNo2(Nothing, m_JobCollection, , , 2, Trim(strFileName))
   
     If m_CollList2.Count > 0 Then
    strDate = "('"
      For Each tempBatch In m_CollList2
         If strDate <> "('" Then
           If DateToStringIntLow(Trim(DateSerial(Mid(TempStr, 7, 4), Mid(TempStr, 4, 2), Mid(TempStr, 1, 2)))) <> DateToStringIntLow(Trim(DateSerial(Mid(tempBatch.BatchStartDate, 7, 4), Mid(tempBatch.BatchStartDate, 4, 2), Mid(tempBatch.BatchStartDate, 1, 2)))) Then
               strDate = strDate & "','" & DateToStringIntLow(Trim(DateSerial(Mid(tempBatch.BatchStartDate, 7, 4), Mid(tempBatch.BatchStartDate, 4, 2), Mid(tempBatch.BatchStartDate, 1, 2))))
              TempStr = Trim(tempBatch.BatchStartDate)
            End If
         Else
            strDate = strDate & DateToStringIntLow(Trim(DateSerial(Mid(tempBatch.BatchStartDate, 7, 4), Mid(tempBatch.BatchStartDate, 4, 2), Mid(tempBatch.BatchStartDate, 1, 2))))
            TempStr = Trim(tempBatch.BatchStartDate)
         End If
      Next tempBatch
      strDate = strDate & "')"
   Else
     strDate = ""
   End If

   If strDate = "" Then
      Call LoadLotFromLot(Nothing, LotColls, , , , 2, , 4)
   Else
      Call LoadLotFromLot(Nothing, LotColls, , , , 2, , 4, , , strDate)
   End If
    
   FileName = File
   F = FreeFile()
   Close #F
   Open FileName For Input As #F
   While (Not EOF(F)) And (ExitDo = False)
      I = I + 1
      Line Input #F, TempStr
      prgProgress.Value = ROUND(MyDiff(I, Sum) * 90, 2)
      txtPercent.Text = prgProgress.Value

      Me.Refresh
      DoEvents
     
      If ProcessLine2(TempStr, refSum, SuccessCount) Then
         SuccessCount = SuccessCount + 1 + refSum
      Else
         If MsgBox("ต้องการจะทำการ อิมพอร์ทข้อมูลชุดนี้ต่อหรือไม่", vbOKCancel, "แจ้งเตือน") = vbCancel Then
             ExitDo = True
             glbDatabaseMngr.DBConnection.RollbackTrans
         End If
         ErrorCount = ErrorCount + 1
      End If
      
   Wend
   Close #F
      
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
      
   HasBegin = True
   
      For Each TempJob In m_JobCollection
         Call PopulateGuiID(TempJob)
         Call glbDaily.Job2InventoryDoc(TempJob, Ivd, 1, 11)
         If Not glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog) Then
            ErrorCount = ErrorCount + 1
            glbErrorLog.LocalErrorMsg = " บันทึกเข้า INVENTORY ERROR"
            glbErrorLog.ShowUserError
            
            If MsgBox("ต้องการจะทำการ อิมพอร์ทข้อมูลชุดนี้ต่อหรือไม่", vbOKCancel, "แจ้งเตือน") = vbCancel Then
                ExitDo = True
                glbDatabaseMngr.DBConnection.RollbackTrans
            End If
         End If
         TempJob.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID

         If Not glbProduction.AddEditJob(TempJob, IsOK, False, glbErrorLog) Then
            ErrorCount = ErrorCount + 1
            glbErrorLog.LocalErrorMsg = " บันทึกเข้า JOB ERROR"
            glbErrorLog.ShowUserError
            If MsgBox("ต้องการจะทำการ อิมพอร์ทข้อมูลชุดนี้ต่อหรือไม่", vbOKCancel, "แจ้งเตือน") = vbCancel Then
                ExitDo = True
                glbDatabaseMngr.DBConnection.RollbackTrans
            End If
         End If
      Next TempJob

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
     
     End If
      Me.Refresh
   Next File
   
   prgProgress.Value = 100
   txtPercent.Text = 100
   
   glbErrorLog.LocalErrorMsg = "อิมพอร์ต สำเร็จจำนวน " & SuccessCount & "และ ล้มเหลวจำนวน " & ErrorCount & " ข้อมูล"
   glbErrorLog.ShowUserError
   
   If (ErrorCount > 0) Then
        glbDatabaseMngr.DBConnection.RollbackTrans
     Else
        If ConfirmSave Then
           glbDatabaseMngr.DBConnection.CommitTrans
        
           glbErrorLog.LocalErrorMsg = "บันทึกข้อมูลเสร็จเรียบร้อยแล้ว"
           glbErrorLog.ShowUserError
           m_HasModify = False
           'Call cmdOK_Click
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
Private Sub loadFileToList()
On Error GoTo ErrorHandler
Dim HasBegin As Boolean
Dim FileName As String
Dim F As Long
Dim LineStr As String
Dim SuccessCount As Long
Dim ErrorCount As Long
Dim Sum As Long
Dim I As Long
Dim J As Long
Dim ExitDo As Boolean
Dim Cb As CBacthing
Dim SearchCB As CBacthing
Dim OldTempAsc As Long
Dim Key1 As String
Dim Key2 As String
Dim strLot() As String

Dim ProductionNumberNew As String
Dim ProductionNumberNewTemp As String

Dim TempBatchSys As Collection
Set TempBatchSys = New Collection
Dim TempCBSys As CBacthing
Dim strArr() As String

 Dim step As Long
 Dim TempBatchStart As String
Dim File As File
    For Each File In ListPartName
   '****************************************หา FromBatch ToBatch
   FileName = File 'txtFileName.Text
   I = 0
   step = 0
   Key1 = ""
   Key2 = ""
   F = FreeFile()
   ExitDo = False
   Close #F
   Open FileName For Input As #F
   Sum = 0
   While Not EOF(F)
      Line Input #F, LineStr
      Sum = Sum + 1
   Wend
     
   SuccessCount = 0
   ErrorCount = 0
   F = FreeFile()
   Close #F
   Open FileName For Input As #F
   While (Not EOF(F)) And (ExitDo = False)
      I = I + 1
      Line Input #F, LineStr
      Me.Refresh
      DoEvents
       Set Cb = New CBacthing
        '000000000104/02/201800030341  006  M-914600527         เบอร์914rework         10/04/201704/02/2018 06:23:00 04/02/2018 06:53:00 B301      000000002055.925000200000001.0000000210.0000000000.10000000001.0000000471.3000000471.2001NCP914              NCP914                                            H1        H1        000000000090.000000000000090.000000000000000.000XXX                 XXX                 MIXER     05/02/2018 09:08:34
       OldTempAsc = 1
       Cb.PlanCode = StingToVariable2(10, OldTempAsc, LineStr) 'รหัสโรงงาน  =0000000001
       Cb.ProductionDate = StingToVariable2(10, OldTempAsc, LineStr) 'วันที่ผลิต = 04/02/2018
       Cb.ProductionNumber = StingToVariable2(10, OldTempAsc, LineStr) 'หมายเลขการผลิต=00030341
       Cb.BatchNumber = StingToVariable2(5, OldTempAsc, LineStr) 'เลขที่ชุดที่ผลิต= 006
       
       ProductionNumberNew = "BK-" & Cb.ProductionNumber 'รหัสนี้ต้องไปหาทุกรหัสที่ซ้ำกัน ที่เป็น -2,-3.. ในระบบได้
      
       Cb.FormulaCode = StingToVariable2(20, OldTempAsc, LineStr)  'รหัสสูตร --> เราใช้เป็นรหัสผลิตภัณฑ์เลย=M-914600527
       Cb.FormulaName = StingToVariable2(50, OldTempAsc, LineStr)  'ชื่อสูตร=เบอร์914rework
       Cb.FormulaDate = StingToVariable2(10, OldTempAsc, LineStr)  'วันที่สูตร=10/04/2017
       Cb.BatchStartDate = StingToVariable2(20, OldTempAsc, LineStr) 'วันเวลาเริ่มผลิต=04/02/2018 06:23:00
       Cb.BatchEndDate = StingToVariable2(20, OldTempAsc, LineStr) 'วันเวลาผลิตเสร็จ=04/02/2018 06:53:00
       Cb.DestinationBin = StingToVariable2(10, OldTempAsc, LineStr) 'รหัสถังปลายทาง=B301
       Cb.ProductionWeight = StingToVariable2(16, OldTempAsc, LineStr) 'น้ำหนักที่ชั่งจริงรวมทั้งชุด=000000002055.925
       Cb.TotalBatch = StingToVariable2(5, OldTempAsc, LineStr) 'Total Batch=00020
       
       If SplitFlag Then  'กรณีที่เป็นการแยก Job งาน
            If JobNo = ProductionNumberNew Then
                  
                     'กระจาย Batch
                     If Key1 <> Trim(ProductionNumberNew) Then 'จะต้องเข้าทำเบอร์ละครั้งเท่านั้น
                       step = 0 'ตั้งค่า step ของ รหัสผลิตเบอร์ใหม่
                        Call LoadDistinctJobNoToLot(Nothing, JobNoColls2, 4, ProductionNumberNew)
                        For Each SearchJobNo In JobNoColls2
                           If Len(SearchJobNo.BATCH_DETAIL) > 0 Then
                           SearchJobNo.BATCH_DETAIL = ""
                            strArr = Split(SearchJobNo.BATCH_DETAIL, ",")
                              If UBound(strArr) > -1 Then
                                 For J = 0 To UBound(strArr)
                                    Set TempCBSys = New CBacthing
                                    TempCBSys.ProductionNumber = ProductionNumberNew 'ให้เป็นชื่อเดิม ที่ไม่มีขีดต่อท้าย
                                    TempCBSys.BatchNumber = strArr(J)
                                    Call TempBatchSys.add(TempCBSys, Trim(TempCBSys.ProductionNumber) & "-" & Trim(TempCBSys.BatchNumber))
                                    Set TempCBSys = Nothing
                                 Next J
                              ElseIf Len(SearchJobNo.BATCH_DETAIL) > 0 Then
                                 Set TempCBSys = New CBacthing
                                 TempCBSys.ProductionNumber = ProductionNumberNew 'ให้เป็นชื่อเดิม ที่ไม่มีขีดต่อท้าย
                                 TempCBSys.BatchNumber = SearchJobNo.BATCH_DETAIL
                                 Call TempBatchSys.add(TempCBSys, Trim(TempCBSys.ProductionNumber) & "-" & Trim(TempCBSys.BatchNumber))
                                 Set TempCBSys = Nothing
                              End If
                           Else
                              For J = SearchJobNo.FROM_BATCH_NO To SearchJobNo.TO_BATCH_NO
                                 Set TempCBSys = New CBacthing
                                 TempCBSys.ProductionNumber = ProductionNumberNew 'ให้เป็นชื่อเดิม ที่ไม่มีขีดต่อท้าย
                                 TempCBSys.BatchNumber = J
                                 Call TempBatchSys.add(TempCBSys, Trim(TempCBSys.ProductionNumber) & "-" & Trim(TempCBSys.BatchNumber))
                                 Set TempCBSys = Nothing
                              Next J
                           End If
                        Next SearchJobNo
                      End If
                      
                      'หาแบตเริ่มต้น สิ้นสุดของแต่ละProductionNumber และ BatchNumber
                     If Key2 <> Trim(ProductionNumberNew) & "-" & Val(Cb.BatchNumber) Then 'จะต้องเข้าทำเบอร์และแบตละครั้งเท่านั้น
                        Key2 = Trim(ProductionNumberNew) & "-" & Trim(Val(Cb.BatchNumber))
                        Set SearchCB = GetObject("CBacthing", TempBatchSys, Key2, False)
                        If SearchCB Is Nothing Then
                           If step = 0 Then
                              Cb.FromBatch = Cb.BatchNumber
                              Cb.ToBatch = Cb.BatchNumber
                              Cb.TotalBatch = Cb.TotalBatch
                              Cb.BatchDetail = "" & Val(Cb.BatchNumber)
                              Cb.tempBatchDetail = Cb.BatchDetail
                              
                              Set SearchJobNo = GetObject("CJob", JobNoColls2, Trim(ProductionNumberNew), False) 'หา Lot และ ถัง หากเคยมีข้อมูลก่อนหน้าแล้ว โดยไม่ต้อง key ใหม่อีก
                              If Not SearchJobNo Is Nothing Then
                                 Cb.LotId = SearchJobNo.LOT_ID
                                 Cb.LotNo = SearchJobNo.LOT_NO
                                 Cb.BIN_NO = SearchJobNo.BIN_NO
                                 Cb.BIN_NAME = SearchJobNo.BIN_NAME
                              End If
                                 
                               ItemCount = ItemCount + 1
                              Cb.ProductionId = ItemCount
                              Key1 = Trim(ProductionNumberNew)
                              Cb.TempProductionNumber = ProductionNumberNew
                              
                              
                              If Cb.FormulaCode = "10541" Or Cb.FormulaCode = "10101" Then    'ถ้าเป็นรำล้างไลน์ หรือ ข้าวโพดล้างไลน์
                                 Cb.SKIP_PART_ITEM_NO = True
                              End If
                              Cb.JOB_ID = SearchJobNo.JOB_ID
                              Call m_CollList1.add(Cb, Key1)
                              I = 0
                              step = Val(Cb.BatchNumber)
               
                           ElseIf Val(Cb.BatchNumber) - step = 1 Then
                             Set SearchCB = GetObject("CBacthing", m_CollList1, Trim(ProductionNumberNew), False)
                              If Not SearchCB Is Nothing Then
                                 SearchCB.ToBatch = Cb.BatchNumber
                                 SearchCB.TotalBatch = Cb.TotalBatch
                                 If Len(SearchCB.BatchDetail) > 0 Then
                                    SearchCB.BatchDetail = SearchCB.BatchDetail & "," & Val(Cb.BatchNumber)
                                    SearchCB.tempBatchDetail = SearchCB.BatchDetail
                                 Else
                                    SearchCB.BatchDetail = "" & Val(Cb.BatchNumber)
                                    SearchCB.tempBatchDetail = SearchCB.BatchDetail
                                 End If
                              End If
                              step = Val(Cb.BatchNumber)
                           ElseIf Val(Cb.BatchNumber) - step > 1 Then 'ถ้าช่องห่างมากกว่า 1 ก็ให้ตั้งช่วงใหม่และตั้งชื่อ ProductionNumber ใหม่
               
                                 Set SearchCB = GetObject("CBacthing", m_CollList1, Trim(ProductionNumberNew), False)
                                 If Not SearchCB Is Nothing Then
                                    SearchCB.ToBatch = Cb.BatchNumber
                                    SearchCB.TotalBatch = Cb.TotalBatch
                                    If Len(SearchCB.BatchDetail) > 0 Then
                                    SearchCB.BatchDetail = SearchCB.BatchDetail & "," & Val(Cb.BatchNumber)
                                    SearchCB.tempBatchDetail = SearchCB.BatchDetail
                                  Else
                                    SearchCB.BatchDetail = "" & Val(Cb.BatchNumber)
                                    SearchCB.tempBatchDetail = SearchCB.BatchDetail
                                  End If
                                 End If
                                 step = Val(Cb.BatchNumber)
                           End If
                        End If
                     End If
                  
                     If Key1 = "" And I = 1 Then
                        Key1 = Trim(ProductionNumberNew) 'Trim(CB.ProductionNumber)
                        I = 0
                      ElseIf Key1 <> Trim(ProductionNumberNew) Then
                        Key1 = Trim(ProductionNumberNew) 'Trim(CB.ProductionNumber)
                        I = 0
                     End If
                  Set Cb = Nothing
                  
            Else
               ExitDo = True
            End If
       
       Else '******************************************************************ห้ามยุ่งข้างล่างนี้
       
      'กระจาย Batch
      
      If Key1 <> Trim(ProductionNumberNew) Then 'จะต้องเข้าทำเบอร์ละครั้งเท่านั้น
        step = 0 'ตั้งค่า step ของ รหัสผลิตเบอร์ใหม่
'         ProductionNumberNew = "BK-" & CB.ProductionNumber 'รหัสนี้ต้องไปหาทุกรหัสที่ซ้ำกัน ที่เป็น -2,-3.. ในระบบได้
         Call LoadDistinctJobNoToLot(Nothing, JobNoColls2, 4, ProductionNumberNew)
         For Each SearchJobNo In JobNoColls2
            If Len(SearchJobNo.BATCH_DETAIL) > 0 Then
             strArr = Split(SearchJobNo.BATCH_DETAIL, ",")
               If UBound(strArr) > -1 Then
                  For J = 0 To UBound(strArr)
                     Set TempCBSys = New CBacthing
                     TempCBSys.ProductionNumber = ProductionNumberNew 'ให้เป็นชื่อเดิม ที่ไม่มีขีดต่อท้าย
                     TempCBSys.BatchNumber = strArr(J)
                     Call TempBatchSys.add(TempCBSys, Trim(TempCBSys.ProductionNumber) & "-" & Trim(TempCBSys.BatchNumber))
                     Set TempCBSys = Nothing
                  Next J
               ElseIf Len(SearchJobNo.BATCH_DETAIL) > 0 Then
                  Set TempCBSys = New CBacthing
                  TempCBSys.ProductionNumber = ProductionNumberNew 'ให้เป็นชื่อเดิม ที่ไม่มีขีดต่อท้าย
                  TempCBSys.BatchNumber = SearchJobNo.BATCH_DETAIL
                  Call TempBatchSys.add(TempCBSys, Trim(TempCBSys.ProductionNumber) & "-" & Trim(TempCBSys.BatchNumber))
                  Set TempCBSys = Nothing
               End If
            Else
               For J = SearchJobNo.FROM_BATCH_NO To SearchJobNo.TO_BATCH_NO
                  Set TempCBSys = New CBacthing
                  TempCBSys.ProductionNumber = ProductionNumberNew 'ให้เป็นชื่อเดิม ที่ไม่มีขีดต่อท้าย
                  TempCBSys.BatchNumber = J
                  Call TempBatchSys.add(TempCBSys, Trim(TempCBSys.ProductionNumber) & "-" & Trim(TempCBSys.BatchNumber))
                  Set TempCBSys = Nothing
               Next J
            End If
         Next SearchJobNo
       End If
       
       'หาแบตเริ่มต้น สิ้นสุดของแต่ละProductionNumber และ BatchNumber
      If Key2 <> Trim(ProductionNumberNew) & "-" & Val(Cb.BatchNumber) Then 'จะต้องเข้าทำเบอร์และแบตละครั้งเท่านั้น
         Key2 = Trim(ProductionNumberNew) & "-" & Trim(Val(Cb.BatchNumber))
         Set SearchCB = GetObject("CBacthing", TempBatchSys, Key2, False)

         If SearchCB Is Nothing Then
            If step = 0 Then
               Cb.FromBatch = Cb.BatchNumber
               Cb.ToBatch = Cb.BatchNumber
               Cb.TotalBatch = Cb.TotalBatch
               Cb.BatchDetail = "" & Val(Cb.BatchNumber)
               
               Set SearchJobNo = GetObject("CJob", JobNoColls2, Trim(ProductionNumberNew), False) 'หา Lot และ ถัง หากเคยมีข้อมูลก่อนหน้าแล้ว โดยไม่ต้อง key ใหม่อีก
               If Not SearchJobNo Is Nothing Then
                  Cb.LotId = SearchJobNo.LOT_ID
                  Cb.LotNo = SearchJobNo.LOT_NO
                  Cb.BIN_NO = SearchJobNo.BIN_NO
                  Cb.BIN_NAME = SearchJobNo.BIN_NAME
               End If
                  
                ItemCount = ItemCount + 1
               Cb.ProductionId = ItemCount
               Key1 = Trim(ProductionNumberNew)
               Cb.TempProductionNumber = ProductionNumberNew
               
               
               If Cb.FormulaCode = "10541" Or Cb.FormulaCode = "10101" Then    'ถ้าเป็นรำล้างไลน์ หรือ ข้าวโพดล้างไลน์
                  Cb.SKIP_PART_ITEM_NO = True
               End If

               Call m_CollList1.add(Cb, Key1)
               I = 0
               step = Val(Cb.BatchNumber)

            ElseIf Val(Cb.BatchNumber) - step = 1 Then
              Set SearchCB = GetObject("CBacthing", m_CollList1, Trim(ProductionNumberNew), False)
               If Not SearchCB Is Nothing Then
                  SearchCB.ToBatch = Cb.BatchNumber
                  SearchCB.TotalBatch = Cb.TotalBatch
                  If Len(SearchCB.BatchDetail) > 0 Then
                     SearchCB.BatchDetail = SearchCB.BatchDetail & "," & Val(Cb.BatchNumber)
                  Else
                     SearchCB.BatchDetail = "" & Val(Cb.BatchNumber)
                  End If
               End If
               step = Val(Cb.BatchNumber)
            ElseIf Val(Cb.BatchNumber) - step > 1 Then 'ถ้าช่องห่างมากกว่า 1 ก็ให้ตั้งช่วงใหม่และตั้งชื่อ ProductionNumber ใหม่

                  Set SearchCB = GetObject("CBacthing", m_CollList1, Trim(ProductionNumberNew), False)
                  If Not SearchCB Is Nothing Then
                     SearchCB.ToBatch = Cb.BatchNumber
                     SearchCB.TotalBatch = Cb.TotalBatch
                     If Len(SearchCB.BatchDetail) > 0 Then
                     SearchCB.BatchDetail = SearchCB.BatchDetail & "," & Val(Cb.BatchNumber)
                   Else
                     SearchCB.BatchDetail = "" & Val(Cb.BatchNumber)
                   End If
                  End If
                  step = Val(Cb.BatchNumber)
            End If
         End If
      End If
   
      If Key1 = "" And I = 1 Then
         Key1 = Trim(ProductionNumberNew) 'Trim(CB.ProductionNumber)
         I = 0
       ElseIf Key1 <> Trim(ProductionNumberNew) Then
         Key1 = Trim(ProductionNumberNew) 'Trim(CB.ProductionNumber)
         I = 0
      End If
    End If 'end if splitFlag
   Set Cb = Nothing
   Wend
   Close #F
   '****************************************จบการหา FromBatch ToBatch
   Next File
   
   Exit Sub
   
ErrorHandler:
End Sub
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
Private Function ProcessLine(LineStr As String, ByRef refSum As Long) As Boolean
On Error GoTo ErrorHandler

Dim TempAsc As Long
Dim OldTempAsc As Long

'Dim SearchJobNo As CJob
Dim MainJob As CJob
Dim SearchLotNo As cLot

Dim SearchProductNo As CPartItem
Dim SearchLocation As CLocation
Dim SearchBacthing As CBacthing

Dim SearchItemNo As CPartItem
Dim SearchBinNo As CLocation

Dim IWD As CInventoryWHDoc
Dim LWH As CLotItemWH
Dim LTD As CLotDoc
Dim PD As CPalletDoc

Dim PlanCode As String
Dim ProductionDate As String
Dim ProductionNumber As String
Dim ProductionNumberNew As String
Dim ProductionNumberNewTemp As String
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

Dim Ma As CJobInput
Dim MI As CJobInput
Dim strArr() As String
Dim I As Long
Dim J As Long
Dim FindResult As Boolean
Dim ExitDo As Boolean
Dim TempDate As Date


'000000000104/02/201800030341  006  M-914600527         เบอร์914rework         10/04/201704/02/2018 06:23:00 04/02/2018 06:53:00 B301      000000002055.925000200000001.0000000210.0000000000.10000000001.0000000471.3000000471.2001NCP914              NCP914                                            H1        H1        000000000090.000000000000090.000000000000000.000XXX                 XXX                 MIXER     05/02/2018 09:08:34
   OldTempAsc = 1
   PlanCode = StingToVariable2(10, OldTempAsc, LineStr) 'รหัสโรงงาน  =0000000001
   ProductionDate = StingToVariable2(10, OldTempAsc, LineStr) 'วันที่ผลิต = 04/02/2018
   ProductionNumber = StingToVariable2(10, OldTempAsc, LineStr) 'หมายเลขการผลิต=00030341
   
   BatchNumber = StingToVariable2(5, OldTempAsc, LineStr) 'เลขที่ชุดที่ผลิต= 006
   FormulaCode = StingToVariable2(20, OldTempAsc, LineStr)  'รหัสสูตร --> เราใช้เป็นรหัสผลิตภัณฑ์เลย=M-914600527
   
   
   FormulaName = StingToVariable2(50, OldTempAsc, LineStr)  'ชื่อสูตร=เบอร์914rework
   
   FormulaDate = StingToVariable2(10, OldTempAsc, LineStr)  'วันที่สูตร=10/04/2017
   BatchStartDate = StingToVariable2(20, OldTempAsc, LineStr) 'วันเวลาเริ่มผลิต=04/02/2018 06:23:00
   BatchEndDate = StingToVariable2(20, OldTempAsc, LineStr) 'วันเวลาผลิตเสร็จ=04/02/2018 06:53:00
      
   DestinationBin = StingToVariable2(10, OldTempAsc, LineStr) 'รหัสถังปลายทาง=B301
  
   ProductionWeight = StingToVariable2(16, OldTempAsc, LineStr) 'น้ำหนักที่ชั่งจริงรวมทั้งชุด=000000002055.925
   TotalBatch = StingToVariable2(5, OldTempAsc, LineStr) 'Total Batch=00020
   TargetDryMix = StingToVariable2(11, OldTempAsc, LineStr) 'Target Dry Mix=0000001.00
   TargetWetMix = StingToVariable2(11, OldTempAsc, LineStr)   'Target Wet Mix=00000210.00
   TargetAfterWetMix = StingToVariable2(11, OldTempAsc, LineStr) 'Target After Wet Mix=00000000.10
   ActualDryMix = StingToVariable2(11, OldTempAsc, LineStr) 'Actual Dry Mix=00000001.00
   ActualWetMix = StingToVariable2(11, OldTempAsc, LineStr) 'Actual Wet Mix=00000471.30
   ActualAfterWetMix = StingToVariable2(11, OldTempAsc, LineStr) 'Actual After Wet Mix=00000471.20
   RuningIngredient = StingToVariable2(2, OldTempAsc, LineStr) 'ลำดับของวัตถุดิบในสูตร=01
   
   IngredientCode = StingToVariable2(20, OldTempAsc, LineStr)  'รหัสวัตถุดิบ=NCP914
   IngredientName = StingToVariable2(50, OldTempAsc, LineStr) 'ชื่อวัตถุดิบ=NCP914
   IngredientType = StingToVariable2(10, OldTempAsc, LineStr)  'ชนิดวัตถุดิบ=H1
   BinCode = StingToVariable2(10, OldTempAsc, LineStr)  'รหัสถังที่ชั่งจริง=H1
   IngredientTargetWeight = StingToVariable2(16, OldTempAsc, LineStr)  'นน ที่ต้องการชั่ง=000000000090.000
   IngredientActualWeight = StingToVariable2(16, OldTempAsc, LineStr)  'นน ที่ชั่งได้จริง=000000000090.000
   IngredientDeviationWeight = StingToVariable2(16, OldTempAsc, LineStr)  'นน Diff=000000000000.000
     
   ProductionNumberNew = "BK-" & ProductionNumber

   I = 1
   
   'ตรวจหาว่า File นี้เคยมีในระบบหรือไม่
   If SplitFlag Then 'ถ้าเป็นการ แยก job ให้ แก้ไข job ทีแยกออกด้วยการเพิ่ม "Sp" ต่อท้าย ProductionNumberNew ด้วย
      ProductionNumberNew = ProductionNumberNew & "-Sp"
   End If
   
   refSum = 0
   Set SearchBacthing = GetObject("CBacthing", m_CollList2, Trim(ProductionNumberNew), False)
   If Not SearchBacthing Is Nothing Then
        If SearchBacthing.Flag = "D" Then
          ProcessLine = True
          refSum = -1
          Exit Function
        End If
        
        'ถ้าแบตใหม่ไม่อยู่ในช่วงที่ต้องการก็ให้ออกจากบรรทัดนั้นไป

         strArr = Split(SearchBacthing.BatchDetail, ",")
         If UBound(strArr) > -1 Then
            For I = 0 To UBound(strArr)
                If Val(BatchNumber) = strArr(I) Then
                   FindResult = True
                   Exit For
                End If
            Next I
            If Not FindResult Then
                ProcessLine = True
                refSum = -1
                Exit Function
            End If
      Else
         ProcessLine = True
         refSum = -1
         Exit Function
        End If
        '*********************
   Else 'หากไม่มีก็ให้ออกไปเหมือนกัน
      ProcessLine = True
      refSum = -1
      Exit Function
   End If
      
      'ตรวจหาวันที่ ที่แก้ไข
      Set SearchBacthing = GetObject("CBacthing", m_CollList2, Trim(ProductionNumberNew), False)
      If Not SearchBacthing Is Nothing Then
         BatchStartDate = SearchBacthing.BatchStartDate
      End If
   
   ExitDo = False
   ProductionNumberNewTemp = ProductionNumberNew
   Set MainJob = GetObject("CJob", m_JobCollection, Trim(ProductionNumberNew), False)
   If Not MainJob Is Nothing Then
      If MainJob.LOCK_DOC_FLAG = "Y" Then
         ProductionNumberNewTemp = ProductionNumberNew
         Set SearchJobNo = GetObject("CJob", JobNoColls, Trim(ProductionNumberNewTemp), False) 'หาชื่อใหม่ที่ไม่ซ้ำกันจากในระบบ
         J = 1
            While (Not SearchJobNo Is Nothing And Not ExitDo)
               If SearchJobNo.LOCK_DOC_FLAG = "Y" Then
                  J = J + 1
                  ProductionNumberNewTemp = ProductionNumberNew & "-" & J
                  Set SearchJobNo = GetObject("CJob", JobNoColls, Trim(ProductionNumberNewTemp), False)
               Else
                  ExitDo = True
               End If
            Wend
      End If
   End If
   
'   JobNoColls
   Set MainJob = GetObject("CJob", m_JobCollection, Trim(ProductionNumberNewTemp), False)
   If MainJob Is Nothing Then 'ถ้าไม่มีก็ Set New พร้อมทั้งตั้งค่าของ Job ก่อน ส่วนถ้ามี Job แล้วให้สร้าง JobInOut อย่างเดียว
      Set MainJob = New CJob
      Set MainJob.InventoryWhDoc = New Collection
      MainJob.JOB_ID = -1
      MainJob.AddEditMode = SHOW_ADD
      MainJob.JOB_NO = ProductionNumberNewTemp
      MainJob.JOB_DESC = "PLC " & FormulaCode & "-" & FormulaName & "-" & FormulaDate
      MainJob.JOB_DATE = DateSerial(Mid(BatchStartDate, 7, 4), Mid(BatchStartDate, 4, 2), Mid(BatchStartDate, 1, 2))
      
      Dim SearchCB As CBacthing
       Set SearchCB = GetObject("CBacthing", m_CollList2, Trim(ProductionNumberNew), False)
       If Not SearchCB Is Nothing Then
         MainJob.BATCH_NO = SearchCB.BatchNumber
         MainJob.FROM_BATCH_NO = SearchCB.FromBatch
         MainJob.TO_BATCH_NO = SearchCB.ToBatch
         MainJob.BATCH_TOTAL = SearchCB.TotalBatch
         MainJob.BATCH_DETAIL = SearchCB.BatchDetail
         SearchCB.SKIP = True
         MainJob.JOB_ID_REF = SearchCB.JOB_ID
      Else
         MainJob.BATCH_NO = Val(MainJob.TO_BATCH_NO) - Val(MainJob.FROM_BATCH_NO) + 1 'Val(BatchNumber) 'Val(TotalBatch)
         MainJob.FROM_BATCH_NO = Val(BatchNumber)
         MainJob.TO_BATCH_NO = Val(BatchNumber)
         MainJob.BATCH_TOTAL = Val(TotalBatch)
         MainJob.BATCH_DETAIL = "" & Val(BatchNumber)
      End If
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
         

      ' Search หา จาก FormulaCode ไปยัง PartColls ถ้ายังไม่เจอให้ ไปหาที่ PartPlcColls และถ้ายังไม่เจออีกให้ขึ้น Form มาให้ใส่ แล้ว Save เข้า PartPlcColls และ UpdatePartColls
      strArr = Split(FormulaCode, "-BK")
      If UBound(strArr) > -1 Then
         FormulaCode = Trim(strArr(0)) & "-BK"
      End If

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
'      Dim Ma As CJobInput
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
  
      Set Ma = Nothing


      If Not SplitFlag Then
         'เข้าคลังอาหาร INVENTORY_WH_DOC
         If Not SearchProductNo Is Nothing Then
         Call LoadLocation(Nothing, m_CollBin, 2, , , , 2, "BIN")
         
         Set IWD = New CInventoryWHDoc
         IWD.DOCUMENT_NO = ProductionNumberNewTemp 'ProductionNumberNew
         IWD.DOCUMENT_DATE = DateSerial(Mid(BatchStartDate, 7, 4), Mid(BatchStartDate, 4, 2), Mid(BatchStartDate, 1, 2)) 'DateSerial(Right(ProductionDate, 4), Mid(ProductionDate, 4, 2), Left(ProductionDate, 2))
         IWD.DOCUMENT_DESC = "PLC " & FormulaCode & "-" & FormulaName & "-" & FormulaDate
         IWD.YYYYMM = Format(DateSerial(Mid(BatchStartDate, 7, 4), Mid(BatchStartDate, 4, 2), Mid(BatchStartDate, 1, 2)), "YYYYMM") 'Format(DateSerial(Right(ProductionDate, 4), Mid(ProductionDate, 4, 2), Left(ProductionDate, 2)), "YYYYMM")
         IWD.PROCESS_ID = 4
         IWD.DOCUMENT_TYPE = JobDocType
         IWD.BATCH_NO = Val(MainJob.TO_BATCH_NO) - Val(MainJob.FROM_BATCH_NO) + 1 'Val(BatchNumber)
         
         
         IWD.START_DATE = DateSerial(Mid(BatchStartDate, 7, 4), Mid(BatchStartDate, 4, 2), Mid(BatchStartDate, 1, 2))
         IWD.START_DATE = DateAdd("h", Val(Mid(BatchStartDate, 12, 2)), MainJob.START_DATE)
         IWD.START_DATE = DateAdd("n", Val(Mid(BatchStartDate, 15, 2)), MainJob.START_DATE)
         IWD.START_DATE = DateAdd("s", Val(Mid(BatchStartDate, 18, 2)), MainJob.START_DATE)
         
         IWD.FINISH_DATE = DateSerial(Mid(BatchEndDate, 7, 4), Mid(BatchEndDate, 4, 2), Mid(BatchEndDate, 1, 2))
         IWD.FINISH_DATE = DateAdd("h", Val(Mid(BatchEndDate, 12, 2)), MainJob.START_DATE)
         IWD.FINISH_DATE = DateAdd("n", Val(Mid(BatchEndDate, 15, 2)), MainJob.START_DATE)
         IWD.FINISH_DATE = DateAdd("s", Val(Mid(BatchEndDate, 18, 2)), MainJob.START_DATE)
         IWD.Flag = "A"
         
               Set LTD = New CLotDoc
               Set PD = New CPalletDoc
               PD.Flag = "A"
               PD.PALLET_DOC_NO = "1000"
               PD.CAPACITY_AMOUNT = 0
               PD.TX_TYPE = "I"
               PD.AddEditMode = SHOW_ADD
               Call LTD.C_PalletDoc.add(PD, Trim(str(SearchProductNo.PART_ITEM_ID)))
               Set PD = Nothing
   
         TempDate = DateSerial(Mid(BatchStartDate, 7, 4), Mid(BatchStartDate, 4, 2), Mid(BatchStartDate, 1, 2))
         Set SearchBacthing = GetObject("CBacthing", m_CollList2, Trim(ProductionNumberNew), False)
         If Not SearchBacthing Is Nothing Then
            If SearchBacthing.SKIP_PART_ITEM_NO = False Then   'ถ้าไม่ใช่รำล้างไลน์ หรือ ข้าวโพดล้างไลน์'If SearchBacthing.LotNo <> "" And SearchBacthing.FormulaCode <> "10541" And SearchBacthing.FormulaCode <> "10101" Then   'ถ้าไม่ใช่รำล้างไลน์ หรือ ข้าวโพดล้างไลน์
                Call LoadLotFromLot(Nothing, LotColls, , , , 2, , 4, , , strDate)
                LTD.Flag = "A"
                If Not SearchBacthing.LotId > 0 Then
                  LTD.LOT_ID = CreateLotAuto(TempDate, Val(SearchBacthing.LotNo), LotColls)
                Else
                  LTD.LOT_ID = SearchBacthing.LotId
                End If
                LTD.BIN_NO = SearchBacthing.BIN_NO
            End If
         End If
         
         LTD.AddEditMode = SHOW_ADD
         Set LWH = New CLotItemWH
         LWH.PART_ITEM_ID = SearchProductNo.PART_ITEM_ID
         LWH.PRODUCT_TYPE_ID = 222
         LWH.BIN_NO = LTD.BIN_NO ' SearchBinNo.LOCATION_ID
         LWH.TX_AMOUNT = MainJob.ACTUAL_AMOUNT
         LWH.GOOD_AMOUNT = MainJob.ACTUAL_AMOUNT
         LWH.PACK_DATE = DateSerial(Mid(BatchStartDate, 7, 4), Mid(BatchStartDate, 4, 2), Mid(BatchStartDate, 1, 2)) 'StingToVariable2(10, 1, BatchStartDate) '04/02/2018 06:23:00
         LWH.TIME_PACK_BEGIN = StingToVariable2(5, 11, BatchStartDate) '04/02/2018 06:23:00
         LWH.TIME_PACK_END = StingToVariable2(5, 11, BatchEndDate) '04/02/2018 06:23:00
         LWH.CALCULATE_FLAG = "N"
         LWH.LOCATION_ID = SearchProductNo.DEFAULT_LOCATION
         LWH.TX_TYPE = "I" 'รับเข้า
         LWH.AddEditMode = SHOW_ADD
         LWH.Flag = "A"
         Call LWH.C_LotDoc.add(LTD)
         Call IWD.C_LotItemsWH.add(LWH, Trim(str(SearchProductNo.PART_ITEM_ID)))
         Call MainJob.InventoryWhDoc.add(IWD)
         Set LTD = Nothing
         Set LWH = Nothing
      End If
   End If 'If Not splitFlag Then
    Call m_JobCollection.add(MainJob, Trim(ProductionNumberNewTemp))
Else 'If MainJob Is Nothing Then *******************************************
     If MainJob.AddEditMode <> SHOW_ADD Then
       MainJob.AddEditMode = SHOW_EDIT
     End If
     
      Set SearchBacthing = GetObject("CBacthing", m_CollList2, Trim(ProductionNumberNew), False)
      If Not SearchBacthing Is Nothing Then
          If SearchBacthing.FromBatch < MainJob.FROM_BATCH_NO Then
            MainJob.FROM_BATCH_NO = SearchBacthing.FromBatch
          End If
          If SearchBacthing.ToBatch > MainJob.TO_BATCH_NO Then
            MainJob.TO_BATCH_NO = SearchBacthing.ToBatch
          End If
          
          If Len(SearchBacthing.BatchDetail) > 0 And Not SearchBacthing.SKIP Then
             MainJob.BATCH_DETAIL = MainJob.BATCH_DETAIL & "," & SearchBacthing.BatchDetail
            SearchBacthing.SKIP = True
          End If
      End If

  If Not SplitFlag Then
      ''''      'เข้าคลังอาหาร INVENTORY_WH_DOC
      If Not MainJob.InventoryWhDoc Is Nothing Then
          Set IWD = MainJob.InventoryWhDoc.Item(1)
          If Not IWD Is Nothing Then
            If IWD.Flag <> "A" Then
              IWD.Flag = "E"
            End If
          End If
      End If
      End If ' end If MainJob Is Nothing Then
 End If
   'end
   
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
'   Dim Mi As CJobInput
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
      If MI.Flag <> "A" Then
         MI.Flag = "E"
      End If
      MI.TX_AMOUNT = MI.TX_AMOUNT + Val(IngredientActualWeight)
      MI.STD_AMOUNT = MI.STD_AMOUNT + Val(IngredientTargetWeight)
   End If
   
   MainJob.STD_AMOUNT = MainJob.STD_AMOUNT + Val(IngredientTargetWeight)
   MainJob.ACTUAL_AMOUNT = MainJob.ACTUAL_AMOUNT + Val(IngredientActualWeight)
   MainJob.BATCH_NO = Val(MainJob.TO_BATCH_NO) - Val(MainJob.FROM_BATCH_NO) + 1 'Val(BatchNumber)
   
   Set Ma = GetObject("CJobInput", MainJob.Outputs, Trim(str(MainJob.PART_ITEM_ID)), False)
   If Not Ma Is Nothing Then
       If Ma.Flag <> "A" Then
         Ma.Flag = "E"
      End If
      Ma.TX_AMOUNT = MainJob.ACTUAL_AMOUNT
      Ma.STD_AMOUNT = MainJob.STD_AMOUNT
    Else
     
   End If
   
    If Not SplitFlag Then
   Set LWH = GetObject("CInventoryWhDoc", MainJob.InventoryWhDoc.Item(1).C_LotItemsWH, Trim(str(MainJob.PART_ITEM_ID)), False)
   If Not LWH Is Nothing Then
      If LWH.Flag <> "A" Then
         LWH.Flag = "E"
       End If
      LWH.TX_AMOUNT = MainJob.ACTUAL_AMOUNT
      LWH.GOOD_AMOUNT = MainJob.ACTUAL_AMOUNT
       
       If LWH.C_LotDoc.Count > 0 Then
         Set PD = GetObject("CInventoryWhDoc", LWH.C_LotDoc.Item(1).C_PalletDoc, Trim(str(MainJob.PART_ITEM_ID)), False)
         If Not PD Is Nothing Then
           If PD.Flag <> "A" Then
              PD.Flag = "E"
           End If
           PD.CAPACITY_AMOUNT = MainJob.ACTUAL_AMOUNT
         End If
       End If
   End If
   Set IWD = Nothing
   End If
   


   ProcessLine = True
   
   Exit Function
ErrorHandler:
   ProcessLine = False
   glbErrorLog.LocalErrorMsg = "Runtime error. At ProductionNumber = " & ProductionNumberNew & " BatchNo = " & BatchNumber
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
End Function
Private Function ProcessLine2(LineStr As String, ByRef refSum As Long, ByRef SuccessCount As Long) As Boolean
On Error GoTo ErrorHandler

Dim TempAsc As Long
Dim OldTempAsc As Long

'Dim SearchJobNo As CJob
Dim MainJob As CJob
Dim SearchLotNo As cLot

Dim SearchProductNo As CPartItem
Dim SearchLocation As CLocation
Dim SearchBacthing As CBacthing

Dim SearchItemNo As CPartItem
Dim SearchBinNo As CLocation

Dim IWD As CInventoryWHDoc
Dim LWH As CLotItemWH
Dim LTD As CLotDoc
Dim PD As CPalletDoc

Dim PlanCode As String
Dim ProductionDate As String
Dim ProductionNumber As String
Dim ProductionNumberNew As String
Dim ProductionNumberNewTemp As String
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

Dim Ma As CJobInput
Dim MI As CJobInput
Dim strArr() As String
Dim I As Long
Dim J As Long
Dim FindResult As Boolean
Dim ExitDo As Boolean
Dim TempDate As Date


'000000000104/02/201800030341  006  M-914600527         เบอร์914rework         10/04/201704/02/2018 06:23:00 04/02/2018 06:53:00 B301      000000002055.925000200000001.0000000210.0000000000.10000000001.0000000471.3000000471.2001NCP914              NCP914                                            H1        H1        000000000090.000000000000090.000000000000000.000XXX                 XXX                 MIXER     05/02/2018 09:08:34
   OldTempAsc = 1
   PlanCode = StingToVariable2(10, OldTempAsc, LineStr) 'รหัสโรงงาน  =0000000001
   ProductionDate = StingToVariable2(10, OldTempAsc, LineStr) 'วันที่ผลิต = 04/02/2018
   ProductionNumber = StingToVariable2(10, OldTempAsc, LineStr) 'หมายเลขการผลิต=00030341
   
   BatchNumber = StingToVariable2(5, OldTempAsc, LineStr) 'เลขที่ชุดที่ผลิต= 006
   FormulaCode = StingToVariable2(20, OldTempAsc, LineStr)  'รหัสสูตร --> เราใช้เป็นรหัสผลิตภัณฑ์เลย=M-914600527
   
   
   FormulaName = StingToVariable2(50, OldTempAsc, LineStr)  'ชื่อสูตร=เบอร์914rework
   
   FormulaDate = StingToVariable2(10, OldTempAsc, LineStr)  'วันที่สูตร=10/04/2017
   BatchStartDate = StingToVariable2(20, OldTempAsc, LineStr) 'วันเวลาเริ่มผลิต=04/02/2018 06:23:00
   BatchEndDate = StingToVariable2(20, OldTempAsc, LineStr) 'วันเวลาผลิตเสร็จ=04/02/2018 06:53:00
      
   DestinationBin = StingToVariable2(10, OldTempAsc, LineStr) 'รหัสถังปลายทาง=B301
  
   ProductionWeight = StingToVariable2(16, OldTempAsc, LineStr) 'น้ำหนักที่ชั่งจริงรวมทั้งชุด=000000002055.925
   TotalBatch = StingToVariable2(5, OldTempAsc, LineStr) 'Total Batch=00020
   TargetDryMix = StingToVariable2(11, OldTempAsc, LineStr) 'Target Dry Mix=0000001.00
   TargetWetMix = StingToVariable2(11, OldTempAsc, LineStr)   'Target Wet Mix=00000210.00
   TargetAfterWetMix = StingToVariable2(11, OldTempAsc, LineStr) 'Target After Wet Mix=00000000.10
   ActualDryMix = StingToVariable2(11, OldTempAsc, LineStr) 'Actual Dry Mix=00000001.00
   ActualWetMix = StingToVariable2(11, OldTempAsc, LineStr) 'Actual Wet Mix=00000471.30
   ActualAfterWetMix = StingToVariable2(11, OldTempAsc, LineStr) 'Actual After Wet Mix=00000471.20
   RuningIngredient = StingToVariable2(2, OldTempAsc, LineStr) 'ลำดับของวัตถุดิบในสูตร=01
   
   IngredientCode = StingToVariable2(20, OldTempAsc, LineStr)  'รหัสวัตถุดิบ=NCP914
   IngredientName = StingToVariable2(50, OldTempAsc, LineStr) 'ชื่อวัตถุดิบ=NCP914
   IngredientType = StingToVariable2(10, OldTempAsc, LineStr)  'ชนิดวัตถุดิบ=H1
   BinCode = StingToVariable2(10, OldTempAsc, LineStr)  'รหัสถังที่ชั่งจริง=H1
   IngredientTargetWeight = StingToVariable2(16, OldTempAsc, LineStr)  'นน ที่ต้องการชั่ง=000000000090.000
   IngredientActualWeight = StingToVariable2(16, OldTempAsc, LineStr)  'นน ที่ชั่งได้จริง=000000000090.000
   IngredientDeviationWeight = StingToVariable2(16, OldTempAsc, LineStr)  'นน Diff=000000000000.000
     
   ProductionNumberNew = "BK-" & ProductionNumber

   I = 1
   
   'ตรวจหาว่า File นี้เคยมีในระบบหรือไม่
   refSum = 0
   Set SearchBacthing = GetObject("CBacthing", m_CollList2, Trim(ProductionNumberNew), False)
   If Not SearchBacthing Is Nothing Then
        If SearchBacthing.Flag = "D" Then
          ProcessLine2 = True
          refSum = -1
          Exit Function
        End If
        
        'ถ้าแบตใหม่ไม่อยู่ในช่วงที่ต้องการก็ให้ออกจากบรรทัดนั้นไป

         strArr = Split(SearchBacthing.BatchDetail, ",")
         If UBound(strArr) > -1 Then
            For I = 0 To UBound(strArr)
                If Val(BatchNumber) = strArr(I) Then
                   FindResult = True
                   Exit For
                End If
            Next I
            If Not FindResult Then
                ProcessLine2 = True
                refSum = -1
                Exit Function
            End If
      Else
         ProcessLine2 = True
         refSum = -1
         Exit Function
        End If
        '*********************
   Else 'หากไม่มีก็ให้ออกไปเหมือนกัน
      ProcessLine2 = True
      refSum = -1
      Exit Function
   End If
      
      'ตรวจหาวันที่ ที่แก้ไข
      Set SearchBacthing = GetObject("CBacthing", m_CollList2, Trim(ProductionNumberNew), False)
      If Not SearchBacthing Is Nothing Then
         BatchStartDate = SearchBacthing.BatchStartDate
      End If
   
   ExitDo = False
   ProductionNumberNewTemp = ProductionNumberNew
   Set MainJob = GetObject("CJob", m_JobCollection, Trim(ProductionNumberNewTemp), False)
   If Not MainJob Is Nothing And SuccessCount = 0 Then
         For Each MI In MainJob.Inputs
            MI.Flag = "D"
         Next MI
         For Each Ma In MainJob.Outputs
            Ma.Flag = "D"
         Next Ma
         Set MainJob.InventoryWhDoc = Nothing
   End If
   
   If Not MainJob Is Nothing And SuccessCount = 0 Then  'ถ้าไม่มีก็ Set New พร้อมทั้งตั้งค่าของ Job ก่อน ส่วนถ้ามี Job แล้วให้สร้าง JobInOut อย่างเดียว
      MainJob.AddEditMode = SHOW_EDIT
      MainJob.JOB_NO = ProductionNumberNewTemp
      MainJob.JOB_DESC = "PLC " & FormulaCode & "-" & FormulaName & "-" & FormulaDate
      MainJob.JOB_DATE = DateSerial(Mid(BatchStartDate, 7, 4), Mid(BatchStartDate, 4, 2), Mid(BatchStartDate, 1, 2))
      
      Dim SearchCB As CBacthing
       Set SearchCB = GetObject("CBacthing", m_CollList2, Trim(ProductionNumberNew), False)
       If Not SearchCB Is Nothing Then
         MainJob.BATCH_NO = SearchCB.BatchNumber
         MainJob.FROM_BATCH_NO = SearchCB.FromBatch
         MainJob.TO_BATCH_NO = SearchCB.ToBatch
         MainJob.BATCH_TOTAL = SearchCB.TotalBatch
         MainJob.BATCH_DETAIL = SearchCB.BatchDetail
         SearchCB.SKIP = True
      Else
         MainJob.BATCH_NO = Val(MainJob.TO_BATCH_NO) - Val(MainJob.FROM_BATCH_NO) + 1 'Val(BatchNumber) 'Val(TotalBatch)
         MainJob.FROM_BATCH_NO = Val(BatchNumber)
         MainJob.TO_BATCH_NO = Val(BatchNumber)
         MainJob.BATCH_TOTAL = Val(TotalBatch)
         MainJob.BATCH_DETAIL = "" & Val(BatchNumber)
      End If
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
         

      ' Search หา จาก FormulaCode ไปยัง PartColls ถ้ายังไม่เจอให้ ไปหาที่ PartPlcColls และถ้ายังไม่เจออีกให้ขึ้น Form มาให้ใส่ แล้ว Save เข้า PartPlcColls และ UpdatePartColls
      
      strArr = Split(FormulaCode, "-BK")
      If UBound(strArr) > -1 Then
         FormulaCode = strArr(0) & "-BK"
      End If

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
                  
                  ProcessLine2 = False
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
               
               ProcessLine2 = False
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
'      Dim Ma As CJobInput
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
  
      Set Ma = Nothing
Else 'If MainJob Is Nothing Then *******************************************
     If MainJob.AddEditMode <> SHOW_ADD Then
       MainJob.AddEditMode = SHOW_EDIT
     End If
     
      Set SearchBacthing = GetObject("CBacthing", m_CollList2, Trim(ProductionNumberNew), False)
      If Not SearchBacthing Is Nothing Then
          If SearchBacthing.FromBatch < MainJob.FROM_BATCH_NO Then
            MainJob.FROM_BATCH_NO = SearchBacthing.FromBatch
          End If
          If SearchBacthing.ToBatch > MainJob.TO_BATCH_NO Then
            MainJob.TO_BATCH_NO = SearchBacthing.ToBatch
          End If
          
          If Len(SearchBacthing.BatchDetail) > 0 And Not SearchBacthing.SKIP Then
             MainJob.BATCH_DETAIL = MainJob.BATCH_DETAIL & "," & SearchBacthing.BatchDetail
            SearchBacthing.SKIP = True
          End If
      End If
End If ' end If MainJob Is Nothing Then
   'end
   
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
                  
               ProcessLine2 = False
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
            
            ProcessLine2 = False
            Exit Function
         End If
         
         SearchLocation.KEY_ID = SearchItemNo.PART_ITEM_ID
         Call LocationUpdateColls.add(SearchLocation, Trim(SearchItemNo.PART_NO))
      End If
      SearchItemNo.DEFAULT_LOCATION = SearchLocation.LOCATION_ID
   End If
      
   'สำหรับ JobInPut Collection
'   Dim Mi As CJobInput
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
      If MI.Flag <> "A" Then
         MI.Flag = "E"
      End If
      MI.TX_AMOUNT = MI.TX_AMOUNT + Val(IngredientActualWeight)
      MI.STD_AMOUNT = MI.STD_AMOUNT + Val(IngredientTargetWeight)
   End If
   
   MainJob.STD_AMOUNT = MainJob.STD_AMOUNT + Val(IngredientTargetWeight)
   MainJob.ACTUAL_AMOUNT = MainJob.ACTUAL_AMOUNT + Val(IngredientActualWeight)
   MainJob.BATCH_NO = Val(MainJob.TO_BATCH_NO) - Val(MainJob.FROM_BATCH_NO) + 1 'Val(BatchNumber)
   
   Set Ma = GetObject("CJobInput", MainJob.Outputs, Trim(str(MainJob.PART_ITEM_ID)), False)
   If Not Ma Is Nothing Then
       If Ma.Flag <> "A" Then
         Ma.Flag = "E"
      End If
      Ma.TX_AMOUNT = MainJob.ACTUAL_AMOUNT
      Ma.STD_AMOUNT = MainJob.STD_AMOUNT
    Else
     
   End If
   
'   Set LWH = GetObject("CInventoryWhDoc", MainJob.InventoryWhDoc.Item(1).C_LotItemsWH, Trim(str(MainJob.PART_ITEM_ID)), False)
'   If Not LWH Is Nothing Then
'      If LWH.Flag <> "A" Then
'         LWH.Flag = "E"
'       End If
'      LWH.TX_AMOUNT = MainJob.ACTUAL_AMOUNT
'      LWH.GOOD_AMOUNT = MainJob.ACTUAL_AMOUNT
'
'       If LWH.C_LotDoc.Count > 0 Then
'         Set PD = GetObject("CInventoryWhDoc", LWH.C_LotDoc.Item(1).C_PalletDoc, Trim(str(MainJob.PART_ITEM_ID)), False)
'         If Not PD Is Nothing Then
'           If PD.Flag <> "A" Then
'              PD.Flag = "E"
'           End If
'           PD.CAPACITY_AMOUNT = MainJob.ACTUAL_AMOUNT
'         End If
'       End If
'   End If
'   Set IWD = Nothing
   
   


   ProcessLine2 = True
   
   Exit Function
ErrorHandler:
   ProcessLine2 = False
   glbErrorLog.LocalErrorMsg = "Runtime error. At ProductionNumber = " & ProductionNumberNew & " BatchNo = " & BatchNumber
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
End Function

Private Function ListFolder(sFolderPath As String) As Boolean
    Dim FS As New FileSystemObject
    Dim FSfolder As Folder
    Dim File As File
    Dim I As Integer
    ListFolder = True
    If Dir(sFolderPath, vbDirectory) = vbNullString Then
      glbErrorLog.LocalErrorMsg = "ไม่มีข้อมูลของวันที่ " & uctlDateSel.ShowDate
      glbErrorLog.ShowUserError
      ListFolder = False
      Exit Function
    End If
    Set ListPartName = New Collection
    Set FSfolder = FS.GetFolder(sFolderPath)
    For Each File In FSfolder.Files
        DoEvents
        If Mid(File.NAME, 1, 1) <> "-" Then
            glbErrorLog.LocalErrorMsg = "ชื่อไฟล์ " & File.NAME & "ไม่มีเครื่องหมาย '-' ตำแหน่งหน้าสุด ซึ่งไม่ถูกต้องตามความต้องการของระบบโปรดตรวจสอบชื่อให้เป็นรูปแบบดังตัวอย่าง '-000000000.txt' หรือ '-000000000'"
            glbErrorLog.ShowUserError
            ListFolder = False
            Exit Function
        End If
       Call ListPartName.add(File)
    Next File
    Set FSfolder = Nothing
End Function





Private Sub Form_Activate()
Dim FromDate As Date
Dim ToDate As Date

   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      uctlDateSel.ShowDate = Now
      Call InitGrid
      Call InitGrid2
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
   pnlHeader.Caption = "อิมพอร์ตข้อมูลผลผลิตจาก PLC"
   
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
   cmdRunAuto.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSelect.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSelectAll.Picture = LoadPicture(glbParameterObj.NormalButton1)

   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdStart, MapText("เริ่ม"))
   Call InitMainButton(cmdFileName, MapText("..."))
   Call InitMainButton(cmdRunAuto, MapText("ตรวจสอบ"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSelect, MapText(">"))
   Call InitMainButton(cmdSelectAll, MapText(">>"))
   
   Call ResetStatus
End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Call Clear
   Unload Me
End Sub
Private Sub Form_Load()
   Call EnableForm(Me, False)
   m_HasActivate = False
    Call SetNew
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Function SetNew()
   Set m_Rs = New ADODB.Recordset
   
   Set PartUctlColls = New Collection
   Set PartColls = New Collection
   Set PartPlcColls = New Collection
   Set PartPlcUpdateColls = New Collection
   
   Set LocationColls = New Collection
   Set LocationUpdateColls = New Collection
   Set JobNoColls = New Collection
   Set JobNoColls2 = New Collection
   Set JobNoColls3 = New Collection
   Set m_JobCollection = New Collection
   Set m_CollLotItemWh = New Collection
   Set TempCollection3 = New Collection
   Set m_CollBin = New Collection
   Set m_CollList1 = New Collection
   Set m_CollList2 = New Collection
   Set LotColls = New Collection
End Function
Private Sub Form_Unload(Cancel As Integer)
 Call SetNothing
End Sub
Private Function SetNothing()
   Set PartUctlColls = Nothing
   Set PartColls = Nothing
   Set PartPlcColls = Nothing
   Set PartPlcUpdateColls = Nothing
   
   Set LocationColls = Nothing
   Set LocationUpdateColls = Nothing
   Set JobNoColls = Nothing
   Set JobNoColls2 = Nothing
   Set JobNoColls3 = Nothing
   
   Set m_CollLotItemWh = Nothing
   Set TempCollection3 = Nothing
   Set m_CollBin = Nothing
   
   Set m_JobCollection = Nothing
   
   Set m_CollList1 = Nothing
   Set m_CollList2 = Nothing
   
   Set LotColls = Nothing
End Function
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

Private Sub GridEX1_DblClick()
   Call cmdSelect_Click
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"


   If m_CollList1 Is Nothing Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If

   Dim Cb As CBacthing
   If m_CollList1.Count <= 0 Then
      Exit Sub
   End If
   Set Cb = GetItem(m_CollList1, RowIndex, RealIndex)
   If Cb Is Nothing Then
      Exit Sub
   End If
   
   Values(1) = Cb.ProductionId
   Values(2) = RealIndex
   Values(3) = Cb.ProductionNumber
   Values(4) = Cb.FormulaName

   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)

End Sub

Private Sub GridEX2_Change()
   m_HasModify = True
End Sub

Private Sub GridEX2_DblClick()
  If Not Val(GridEX2.Value(1)) > 0 Then
      Exit Sub
  End If
  Call ShowLot(Val(GridEX2.Value(2)))
End Sub
Private Function ShowLot(ID As Long)
   frmAddEditLotNo.ID = ID
   frmAddEditLotNo.HeaderText = MapText("แก้ไขข้อมูล LOT การผลิต")
   frmAddEditLotNo.ShowMode = SHOW_EDIT
   frmAddEditLotNo.SplitFlag = SplitFlag
   Set frmAddEditLotNo.ParentForm = Me
   Set frmAddEditLotNo.TempCollection = m_CollList2
   Load frmAddEditLotNo
   frmAddEditLotNo.Show 1
   
   OKClick = frmAddEditLotNo.OKClick
   
   Unload frmAddEditLotNo
   Set frmAddEditLotNo = Nothing
               
   If OKClick Then
      GridEX2.ItemCount = CountItem(m_CollList2) 'itemcount
      GridEX2.Rebind
   End If
End Function
'Private Sub txtLotNoNew_Change()
'   m_HasModify = True
'End Sub

'Private Sub uctlStartDate_HasChange()
'   Call LoadLotFromLot(cboLotNo, Nothing, , uctlStartDate.ShowDate, uctlStartDate.ShowDate, , , 1, , 2)
'End Sub

Private Sub GridEX2_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"


   If m_CollList2 Is Nothing Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If

   Dim Cb As CBacthing
   If m_CollList2.Count <= 0 Then
      Exit Sub
   End If
   Set Cb = GetItem(m_CollList2, RowIndex, RealIndex)
   If Cb Is Nothing Then
      Exit Sub
   End If
   
   Values(1) = Cb.ProductionId
   Values(2) = RealIndex
   Values(3) = Cb.BatchStartDate
   Values(4) = Cb.ProductionNumber
   Values(5) = Cb.FormulaCode
   Values(6) = Cb.FormulaName
   Values(7) = Cb.LotNo
   Values(8) = Cb.BIN_NAME
   Values(9) = Format(Cb.FromBatch, "000")
   Values(10) = Format(Cb.ToBatch, "000")
   Values(11) = Cb.BatchDetail
   Values(12) = Cb.TotalBatch
    Values(13) = Cb.SplitFlag
   
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub InitGrid()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   
   GridEX1.FormatStyles.Clear
   Set fmsTemp = GridEX1.FormatStyles.add("N")
   fmsTemp.ForeColor = GLB_ALERT_COLOR
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 500
   Col.TextAlignment = jgexAlignCenter
   Col.Caption = "ลำดับ"
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 1200
   Col.Caption = MapText("เลขการผลิต")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 2500
   Col.Caption = "ชื่อสูตร"
   
   GridEX1.ItemCount = 0
End Sub
Private Sub InitGrid2()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

   GridEX2.Columns.Clear
   GridEX2.BackColor = GLB_GRID_COLOR
   GridEX2.BackColorHeader = GLB_GRIDHD_COLOR
   
   GridEX2.FormatStyles.Clear
   Set fmsTemp = GridEX2.FormatStyles.add("N")
   fmsTemp.ForeColor = GLB_ALERT_COLOR
   
   Set Col = GridEX2.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"
   
   Set Col = GridEX2.Columns.add '1
   Col.Width = 500
   Col.TextAlignment = jgexAlignCenter
   Col.Caption = "ลำดับ"

   Set Col = GridEX2.Columns.add '2
   Col.Width = 2300
   Col.Caption = MapText("วันที่ผลิต")
      
   Set Col = GridEX2.Columns.add '3
   Col.Width = 1200
   Col.Caption = MapText("เลขการผลิต")

   Set Col = GridEX2.Columns.add '4
   Col.Width = 2500
   Col.Caption = MapText("รหัสสูตร")
   
   Set Col = GridEX2.Columns.add '5
   Col.Width = 2500
   Col.Caption = "ชื่อสูตร"
   
   Set Col = GridEX2.Columns.add '7
   Col.Width = 1700
   Col.Caption = "Lot No"
   
   Set Col = GridEX2.Columns.add '6
   Col.Width = 1000
   Col.Caption = "ถัง"
   
   Set Col = GridEX2.Columns.add '7
   Col.Width = 700
   Col.TextAlignment = jgexAlignRight
   Col.Caption = "จากแบต"
   
   Set Col = GridEX2.Columns.add '7
   Col.Width = 700
   Col.TextAlignment = jgexAlignRight
   Col.Caption = "ถึงแบต"
   
   Set Col = GridEX2.Columns.add '7
   Col.Width = 1100
   Col.TextAlignment = jgexAlignLeft
   Col.Caption = "รายละเอียดแบต"
   
   Set Col = GridEX2.Columns.add '7
   Col.Width = 800
   Col.TextAlignment = jgexAlignRight
   Col.Caption = "จำนวนแบต"
   
   Set Col = GridEX2.Columns.add '6
   Col.Width = 0
   Col.Caption = "SPLITFLAG"
   
   GridEX2.ItemCount = 0
End Sub
Public Sub RefreshGrid()
   GridEX2.ItemCount = CountItem(m_CollList2)
   GridEX2.Rebind
End Sub

Private Sub SSCommand1_Click()

End Sub

Private Sub txtFileName_Change()
   m_HasModify = True
End Sub
Public Function BrowseFolder(szDialogTitle As String) As String
  Dim X As Long
  Dim Bi As BROWSEINFO
  Dim dwIList As Long
  Dim szPath As String, wPos As Integer
  
    With Bi
        .hOwner = 0
        .lpszTitle = szDialogTitle
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    
    dwIList = SHBrowseForFolder(Bi)
    szPath = Space$(512)
    X = SHGetPathFromIDList(ByVal dwIList, ByVal szPath)
    
    If X Then
        wPos = InStr(szPath, Chr(0))
        BrowseFolder = Left$(szPath, wPos - 1)
    Else
        BrowseFolder = ""
    End If
End Function
