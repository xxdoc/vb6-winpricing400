VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmImportPartItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Enabled         =   0   'False
   Icon            =   "frmImportPartItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   10410
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8085
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   14261
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboLocation 
         Height          =   315
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2520
         Width           =   2955
      End
      Begin VB.ComboBox cboUnit 
         Height          =   315
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1560
         Width           =   2955
      End
      Begin VB.ComboBox cboParcelType 
         Height          =   315
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2040
         Width           =   2955
      End
      Begin VB.ComboBox cboPartType 
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
         Top             =   3630
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   767
      End
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   1860
         TabIndex        =   0
         Top             =   4080
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
         Top             =   4410
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
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   20
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label lblParcelType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   19
         Top             =   2190
         Width           =   1575
      End
      Begin VB.Label lblLocation 
         Alignment       =   1  'Right Justify
         Caption         =   "lblLocation"
         Height          =   375
         Left            =   270
         TabIndex        =   18
         Top             =   2610
         Width           =   1485
      End
      Begin Threed.SSCommand cmdFileName 
         Height          =   405
         Left            =   8670
         TabIndex        =   13
         Top             =   3630
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportPartItem.frx":27A2
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   1890
         TabIndex        =   2
         Top             =   5070
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportPartItem.frx":2ABC
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3600
         TabIndex        =   11
         Top             =   4530
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   10
         Top             =   4140
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   9
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label lblMasterName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   8
         Top             =   3660
         Width           =   1575
      End
      Begin VB.Label lblPartType 
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
         Top             =   5070
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
         Top             =   5070
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportPartItem.frx":2DD6
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmImportPartItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String

Public PartGroupID As Long
Private m_PartItem As CPartItem
Private m_PartItems As Collection
Private m_PartTypes As Collection

Private m_ExcelApp As Object
Private m_ExcelSheet As Object





Private Sub cboPartType_Click()
Dim PartTypeID As Long
Dim Pt As CPartType

   PartTypeID = cboPartType.ItemData(Minus2Zero(cboPartType.ListIndex))
   
   If PartTypeID > 0 Then
      Set Pt = GetPartType(m_PartTypes, Trim(str(PartTypeID)))
      Call LoadPartItem(Nothing, m_PartItems, PartTypeID, "N", , 2)
'      Set uctlProductLookup.MyCollection = m_PartItems
   
'      Call LoadLocation(uctlPlaceLookup.MyCombo, m_Locations, 2, , , Pt.PART_GROUP_ID)
'      Set uctlPlaceLookup.MyCollection = m_Locations
   End If
   
   m_HasModify = True
End Sub

Private Sub cmdFileName_Click()
On Error Resume Next
Dim strDescription As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Access Files (*.XLS)|*..xls;*.XLS;"
   dlgAdd.DialogTitle = "Select Excel file to import"
   dlgAdd.ShowOpen
   If dlgAdd.fileName = "" Then
      Exit Sub
   End If
    
   txtFileName.Text = dlgAdd.fileName
   m_HasModify = True
End Sub

Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub
Private Sub cmdStart_Click()
Dim TempID As Long
   
   If Not VerifyCombo(lblPartType, cboPartType) Then
      Exit Sub
   End If
   If Not VerifyTextControl(lblMasterName, txtFileName) Then
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   m_ExcelApp.Workbooks.Close
   m_ExcelApp.Workbooks.Open (txtFileName.Text)
   
   Call ImportStock
   
   m_ExcelApp.Workbooks.Close
   Call EnableForm(Me, True)
End Sub

Private Function GetVal(row As Long, Col As Long) As Double
On Error Resume Next

   GetVal = m_ExcelSheet.Cells(row, Col).Value
End Function

Private Sub ImportStock()
On Error GoTo ErrorHandler
Dim MaxRow As Long
Dim MaxCol As Long
Dim ID As Long
Dim FieldNames() As String
Dim FieldTypes() As String
Dim i As Long
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
Dim TempPi As CPartItem

Dim IsOK As Boolean

   HasBegin = False
   
   Set m_ExcelSheet = m_ExcelApp.Sheets(1)
      
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
   For Each TempPi In m_PartItems 'ยกเลิก อาหารทุกเบอร์ก่อน
      TempPi.CANCEL_FLAG = "Y"
      TempPi.UpdateCancelFlag
   Next TempPi
   
   For row = 11 To MaxRow
      DoEvents
      If Len(Trim(m_ExcelSheet.Cells(row, 5).Value)) > 0 Then

      Set TempPi = GetObject("CPartItem", m_PartItems, Trim(m_ExcelSheet.Cells(row, 9).Value))
       Set m_PartItem = New CPartItem
            If Not TempPi Is Nothing And TempPi.DuptCheck = False Then  'ถ้ามีแล้วให้ update
              TempPi.DuptCheck = True
               m_PartItem.AddEditMode = SHOW_EDIT
               m_PartItem.PART_ITEM_ID = TempPi.PART_ITEM_ID
            Else
               m_PartItem.AddEditMode = SHOW_ADD
            End If

            m_PartItem.WEIGHT_PER_PACK = Val(m_ExcelSheet.Cells(row, 4).Value)
            m_PartItem.PART_NO = Trim(m_ExcelSheet.Cells(row, 5).Value)
            m_PartItem.PART_DESC = Trim(m_ExcelSheet.Cells(row, 6).Value)
            m_PartItem.BARCODE_NO = Trim(m_ExcelSheet.Cells(row, 7).Value)
            m_PartItem.BILL_DESC = Trim(m_ExcelSheet.Cells(row, 8).Value)
            m_PartItem.PART_TYPE = cboPartType.ItemData(Minus2Zero(cboPartType.ListIndex))

            m_PartItem.PIG_FLAG = "N"
            m_PartItem.UNIT_COUNT = cboUnit.ItemData(Minus2Zero(cboUnit.ListIndex))
            m_PartItem.PARCEL_TYPE = cboParcelType.ItemData(Minus2Zero(cboParcelType.ListIndex))
            m_PartItem.DEFAULT_LOCATION = cboLocation.ItemData(Minus2Zero(cboLocation.ListIndex))
            m_PartItem.CANCEL_FLAG = "N" 'เปิดใช้งานอาหารเบอร์นี้
            If Trim(m_ExcelSheet.Cells(row, 14).Value) = 1 Then
               m_PartItem.ANIMAL_TYPE = 235
            ElseIf Trim(m_ExcelSheet.Cells(row, 14).Value) = 2 Then
               m_PartItem.ANIMAL_TYPE = 236
            ElseIf Trim(m_ExcelSheet.Cells(row, 14).Value) = 3 Then
               m_PartItem.ANIMAL_TYPE = 237
            End If

            Call glbDaily.AddEditPartItem(m_PartItem, IsOK, False, glbErrorLog)

            ProgressCount = ProgressCount + 1
            prgProgress.Value = ProgressCount

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

Private Sub Form_Activate()
Dim FromDate As Date
Dim ToDate As Date

   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadPartType(cboPartType, , PartGroupID)
      
      Call LoadUnit(cboUnit)
      Call InitParcelTypeEx(cboParcelType)
      
      Call LoadLocation(cboLocation, Nothing, 2)
      
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
   pnlHeader.Caption = "อิมพอร์ตข้อมูลสินค้าวัตถุดิบ"
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitNormalLabel(lblPartType, "ประเภทสินค้า")
   Call InitNormalLabel(lblMasterName, "ชื่อไฟล์")
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(lblUnit, MapText("หน่วยวัด"))
   Call InitNormalLabel(lblParcelType, MapText("ประเภทบรรจุ"))
   Call InitNormalLabel(lblLocation, MapText("คลังหลัก PLC"))
   
   Call InitNormalLabel(Label1, "%")


   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPercent.Enabled = False
   Call txtFileName.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtFileName.Enabled = False
   
   Call InitCombo(cboPartType)
   Call InitCombo(cboLocation)
   Call InitCombo(cboUnit)
   Call InitCombo(cboParcelType)
   
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
   
   Set m_PartItem = New CPartItem
   Set m_PartItems = New Collection
   Set m_PartTypes = New Collection
   
   Set m_ExcelApp = CreateObject("Excel.application")
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_PartItem = Nothing
   Set m_PartItems = Nothing
   Set m_PartTypes = Nothing
End Sub


