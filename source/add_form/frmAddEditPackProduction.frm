VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditPackProduction 
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15705
   Icon            =   "frmAddEditPackProduction.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   15705
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   7575
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   15720
      _ExtentX        =   27728
      _ExtentY        =   13361
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   120
         TabIndex        =   2
         Top             =   1920
         Width           =   11640
         _ExtentX        =   20532
         _ExtentY        =   979
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "AngsanaUPC"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   15735
         _ExtentX        =   27755
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   4230
         Left            =   120
         TabIndex        =   3
         Top             =   2470
         Width           =   15240
         _ExtentX        =   26882
         _ExtentY        =   7461
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
         Column(1)       =   "frmAddEditPackProduction.frx":27A2
         Column(2)       =   "frmAddEditPackProduction.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditPackProduction.frx":290E
         FormatStyle(2)  =   "frmAddEditPackProduction.frx":2A6A
         FormatStyle(3)  =   "frmAddEditPackProduction.frx":2B1A
         FormatStyle(4)  =   "frmAddEditPackProduction.frx":2BCE
         FormatStyle(5)  =   "frmAddEditPackProduction.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditPackProduction.frx":2D5E
      End
      Begin prjFarmManagement.uctlTextBox txtDesc 
         Height          =   435
         Left            =   1560
         TabIndex        =   0
         Top             =   1320
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlPackDate 
         Height          =   405
         Left            =   7800
         TabIndex        =   1
         Top             =   840
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtPackNo 
         Height          =   435
         Left            =   1560
         TabIndex        =   16
         Top             =   840
         Width           =   2895
         _ExtentX        =   18018
         _ExtentY        =   767
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   525
         Left            =   10440
         TabIndex        =   17
         Top             =   6840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPackProduction.frx":2F36
         ButtonStyle     =   3
      End
      Begin VB.Label lblPackNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPackNo"
         Height          =   315
         Left            =   0
         TabIndex        =   15
         Top             =   960
         Width           =   1395
      End
      Begin Threed.SSCommand cmdAuto 
         Height          =   405
         Left            =   4560
         TabIndex        =   14
         Top             =   840
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPackProduction.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   525
         Left            =   5400
         TabIndex        =   7
         Top             =   6960
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPackProduction.frx":356A
         ButtonStyle     =   3
      End
      Begin VB.Label lblPackDate 
         Alignment       =   1  'Right Justify
         Caption         =   "lblJobDate"
         Height          =   315
         Left            =   6360
         TabIndex        =   13
         Top             =   840
         Width           =   1305
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "lblDesc"
         Height          =   315
         Left            =   0
         TabIndex        =   12
         Top             =   1440
         Width           =   1395
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   12120
         TabIndex        =   8
         Top             =   6840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPackProduction.frx":3884
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   13800
         TabIndex        =   9
         Top             =   6840
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1800
         TabIndex        =   5
         Top             =   6840
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   120
         TabIndex        =   4
         Top             =   6840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPackProduction.frx":3B9E
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3480
         TabIndex        =   6
         Top             =   6840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPackProduction.frx":3EB8
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditPackProduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_PackProduction As CPackProduction
Private m_PackProductions As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public id As Long
Public JobDocType As Long
Public ProcessID As Long

Private FileName As String
Private m_SumUnit As Double
Private m_ExtractItems As Collection
Private m_PartItems As Collection
Private m_ProductPartUseds As Collection
Private m_ProductPartUsed As Collection

Public TempCollection As Collection

Private Sub EnableDisableButton(En As Boolean)
   If En Then
      If ShowMode = SHOW_EDIT Then
'         cmdAdd.Enabled = (m_PackProduction.COMMIT_FLAG = "N")
'         cmdDelete.Enabled = (m_PackProduction.COMMIT_FLAG = "N")
      Else
         cmdAdd.Enabled = True
         cmdEdit.Enabled = True
         cmdDelete.Enabled = True
      End If
   Else
      cmdAdd.Enabled = En
      cmdDelete.Enabled = En
      cmdEdit.Enabled = En
   End If
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
      
      m_PackProduction.PACK_PRODUCTION_ID = id
      m_PackProduction.PACK_PRODUCTION_DATE = -1
      m_PackProduction.QueryFlag = 1
      If Not glbProduction.QueryPackProduction(m_PackProduction, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
        Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_PackProduction.PopulateFromRS(1, m_Rs)
      txtPackNo.Text = m_PackProduction.PACK_PRODUCTION_NO
      txtDesc.Text = m_PackProduction.PACK_PRODUCTION_DESC
      uctlPackDate.ShowDate = m_PackProduction.PACK_PRODUCTION_DATE
      cmdAdd.Enabled = (m_PackProduction.PackItems.Count <= 0)
      
      Call EnableDisableButton(True)
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call TabStrip1_Click
   
   Call EnableForm(Me, True)
End Sub

Private Sub cboGroup_Click()
   m_HasModify = True
End Sub

Private Sub chkEnable_Click()
   m_HasModify = True
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim Ivd As CInventoryDoc
   If ShowMode = SHOW_EDIT Then
      If Not VerifyAccessRight("PRODUCT_PACK_EDIT") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   End If
   If Not VerifyTextControl(lblPackNo, txtPackNo, False) Then
       Exit Function
   End If

   
   If Not VerifyDate(lblPackDate, uctlPackDate, False) Then
      Exit Function
   End If
      
       
   If Not CheckUniqueNs(PACK_PRODUCTION_UNIQUE, Trim(DateToStringInt(uctlPackDate.ShowDate)), id) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูลวันที่ ") & " " & Trim(DateToStringInt(uctlPackDate.ShowDate)) & " " & MapText(" อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Exit Function
   End If

'      If Not CheckUniqueNs(PLANNING_UNIQUE, Trim(DateToStringInt(uctlPlanningFrom.ShowDate)), ID, Trim(str(PlanningArea))) Then
'         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & uctlPlanningFrom.ShowDate & " " & MapText("อยู่ในระบบแล้ว")
'         glbErrorLog.ShowUserError
'         Exit Function
'      End If

   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_PackProduction.PACK_PRODUCTION_ID = id
   m_PackProduction.AddEditMode = ShowMode
   m_PackProduction.PACK_PRODUCTION_NO = txtPackNo.Text
   m_PackProduction.PACK_PRODUCTION_DESC = txtDesc.Text
   m_PackProduction.PACK_PRODUCTION_DATE = uctlPackDate.ShowDate
   m_PackProduction.PACK_PRODUCTION_AREA = 1
'   m_PackProduction.JOB_FROM_DATE = uctlPackDate.ShowDate
'   m_PackProduction.JOB_TO_DATE = uctlPackDate.ShowDate
   
   Call EnableForm(Me, False)
      
   If Not glbProduction.AddEditPackProduction(m_PackProduction, IsOK, True, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call glbDaily.RollbackTransaction
      Call EnableForm(Me, True)
      Exit Function
   End If
   
   If Not IsOK Then
      Call EnableForm(Me, True)
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   Call EnableForm(Me, True)
   SaveData = True
End Function
Private Sub chkCommit_Click(Value As Integer)
m_HasModify = True
End Sub
Private Sub cmdAuto_Click()
Dim No As String

   If Trim(txtPackNo.Text) = "" Then
      Call glbDatabaseMngr.GenerateNumber(PACK_PRODUCTION, No, glbErrorLog)
      txtPackNo.Text = No
   End If
End Sub

Private Sub CalculateTotalRatio()
   
End Sub

Private Sub DeleteAllItem()
Dim Ei As CCostPrdItem

   For Each Ei In m_PackProduction.PackItems
      Ei.Flag = "D"
   Next Ei
End Sub

Private Sub cmdAdd_Click()
Dim Ji As CJobInput
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim Ei As CCostPrdItem
Dim Pi As CPartItem
Dim TempEi As CCostPrdItem

'   If Not VerifyDate(lblStartJob, uctlStartJob, False) Then
'      Exit Sub
'   End If
'   If Not VerifyDate(lblFinishJob, uctlFinishJob, False) Then
'      Exit Sub
'   End If
   
   Call EnableForm(Me, False)
   
   Set TempRs = New ADODB.Recordset
   
   If TabStrip1.SelectedItem.Index = 1 Then
     Set frmAddEditPackProductItem.TempCollection = m_PackProduction.PackItems
      frmAddEditPackProductItem.ShowMode = SHOW_ADD
      frmAddEditPackProductItem.HeaderText = MapText("เพิ่มข้อมูลใบสั่งแพ็คอาหาร")
      frmAddEditPackProductItem.PartType = 10
'      frmAddEditCostExpense.FromDate = uctlStartJob.ShowDate
'      frmAddEditCostExpense.ToDate = uctlFinishJob.ShowDate
      Load frmAddEditPackProductItem
      frmAddEditPackProductItem.Show 1

      OKClick = frmAddEditPackProductItem.OKClick

      Unload frmAddEditPackProductItem
      Set frmAddEditPackProductItem = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_PackProduction.PackItems)
         GridEX1.Rebind
      End If
   End If
   Call EnableForm(Me, True)
   
   m_HasModify = True
End Sub

Private Sub cmdDelete_Click()
Dim ID1 As Long
Dim ID2 As Long

   If Not cmdDelete.Enabled Then
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
   If Not ConfirmDelete(GridEX1.Value(3)) Then
      Exit Sub
   End If
   
   ID2 = GridEX1.Value(2)
   ID1 = GridEX1.Value(1)
   
    If TabStrip1.SelectedItem.Index = 1 Then
      If ID1 <= 0 Then
         m_PackProduction.PackItems.Remove (ID2)
      Else
         m_PackProduction.PackItems.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_PackProduction.PackItems)
      GridEX1.Rebind
      m_HasModify = True
   End If
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim id As Long
Dim OKClick As Boolean

   If Not cmdEdit.Enabled Then
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
   id = Val(GridEX1.Value(2))
   OKClick = False
   
   If TabStrip1.SelectedItem.Index = 1 Then
     Set frmAddEditPackProductItem.TempCollection = m_PackProduction.PackItems
      frmAddEditPackProductItem.id = id
      frmAddEditPackProductItem.PartType = 10
      frmAddEditPackProductItem.ShowMode = SHOW_EDIT
      frmAddEditPackProductItem.HeaderText = MapText("แก้ไขต้นทุนผลิต")
      Load frmAddEditPackProductItem
      frmAddEditPackProductItem.Show 1

      OKClick = frmAddEditPackProductItem.OKClick

      Unload frmAddEditPackProductItem
      Set frmAddEditPackProductItem = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_PackProduction.PackItems)
         GridEX1.Rebind
      End If
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdOK_Click()
Dim oMenu As cPopupMenu
Dim lMenuChosen  As Long

   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("บันทึก", "-", "บันทึกและออกจากหน้าจอ")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
    
   If lMenuChosen = 1 Then
      If Not SaveData Then
         Exit Sub
      End If
      
      ShowMode = SHOW_EDIT
      m_PackProduction.QueryFlag = 1
      QueryData (True)
      m_HasModify = False
   ElseIf lMenuChosen = 3 Then
      If Not SaveData Then
         Exit Sub
      End If
      
      OKClick = True
      Unload Me
   End If
End Sub

Private Sub cmdPrint_Click()
Dim Report As CReportInterface
Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
Dim ReportKey As String
Dim ReportFlag As Boolean
Dim Rc As CReportConfig
Dim iCount As Long
Dim EditMode As SHOW_MODE_TYPE
Dim ReportMode As Long

   ReportMode = 1

   If m_HasModify Or ((Not m_HasModify) And (ShowMode = SHOW_ADD)) Then
      glbErrorLog.LocalErrorMsg = "กรุณาทำการบันทึกข้อมูลให้เรียบร้อยก่อน"
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("ใบรายงานใบสั่งแพ็คอาหาร", "ปรับค่าหน้ากระดาษ")
   If lMenuChosen = 0 Then
      Exit Sub
   End If

   Call EnableForm(Me, False)

   If lMenuChosen = 1 Then
      ReportKey = "CReportPackProduction"
      Set Report = New CReportPackProduction
      ReportFlag = True
      Call Report.AddParam(1, "PREVIEW_TYPE")
   End If

   If Not Report Is Nothing Then

      Call Report.AddParam(m_PackProduction, "m_PACK_PRODUCTION")
      Call Report.AddParam(uctlPackDate.ShowDate, "DOCUMENT_DATE")
      Call Report.AddParam(ReportKey, "REPORT_KEY")
   End If

   If ReportFlag Then
    frmReport.ClassName = ReportKey
      Set frmReport.ReportObject = Report

      frmReport.HeaderText = pnlHeader.Caption
      Load frmReport
      frmReport.Show 1

      Unload frmReport
      Set frmReport = Nothing
      Set Report = Nothing
   End If

   Call EnableForm(Me, True)
End Sub

Private Sub cmdSave_Click()
Dim Result As Boolean
   If Not SaveData Then
      Exit Sub
   End If
   
   ShowMode = SHOW_EDIT
   id = m_PackProduction.PACK_PRODUCTION_ID
   m_PackProduction.QueryFlag = 1
   QueryData (True)
   m_HasModify = False
End Sub

Private Function GetTotalRelateAmount(Ce As CCostExpense, Optional NetRaw As Double) As Double
'Dim Cpi As CCostPrdItem
'Dim Tempsum As Double
'Dim Ji As CJobInput
'
'   Tempsum = 0
'   NetRaw = 0
'   For Each Cpi In m_PackProduction.CostItems
'      If Cpi.Flag <> "D" Then
'
'         If (Ce.PACKAGE_TYPE = PARCEL_BAG) And (Ce.RATIO_TYPE = RATIO_QUANTITY) And (Cpi.PARCEL_TYPE = PARCEL_BAG) Then
'            Tempsum = Tempsum + Cpi.COST_AMOUNT
'         ElseIf (Ce.PACKAGE_TYPE = PARCEL_BULK) And (Ce.RATIO_TYPE = RATIO_QUANTITY) And (Cpi.PARCEL_TYPE = PARCEL_BULK) Then
'            Tempsum = Tempsum + Cpi.COST_AMOUNT
'         ElseIf (Ce.PACKAGE_TYPE = PARCEL_ALL) And (Ce.RATIO_TYPE = RATIO_QUANTITY) Then
'            Tempsum = Tempsum + Cpi.COST_AMOUNT
'
'         ElseIf (Ce.PACKAGE_TYPE = PARCEL_BAG) And (Ce.RATIO_TYPE = RATIO_COST) And (Cpi.PARCEL_TYPE = PARCEL_BAG) Then
'            Tempsum = Tempsum + Cpi.RAW_AMOUNT
'         ElseIf (Ce.PACKAGE_TYPE = PARCEL_BULK) And (Ce.RATIO_TYPE = RATIO_COST) And (Cpi.PARCEL_TYPE = PARCEL_BULK) Then
'            Tempsum = Tempsum + Cpi.RAW_AMOUNT
'         ElseIf (Ce.PACKAGE_TYPE = PARCEL_ALL) And (Ce.RATIO_TYPE = RATIO_COST) Then
'            Tempsum = Tempsum + Cpi.RAW_AMOUNT
'
'         ElseIf (Ce.PACKAGE_TYPE = PARCEL_BAG) And (Ce.RATIO_TYPE = RATIO_RAW) And (Cpi.PARCEL_TYPE = PARCEL_BAG) Then
'            Set Ji = GetJobInOut(m_ProductPartUseds, Cpi.PART_ITEM_ID & "-E")
'            Tempsum = Tempsum + Ji.TX_AMOUNT
'         ElseIf (Ce.PACKAGE_TYPE = PARCEL_BULK) And (Ce.RATIO_TYPE = RATIO_RAW) And (Cpi.PARCEL_TYPE = PARCEL_BULK) Then
'            Set Ji = GetJobInOut(m_ProductPartUseds, Cpi.PART_ITEM_ID & "-E")
'            Tempsum = Tempsum + Ji.TX_AMOUNT
'         ElseIf (Ce.PACKAGE_TYPE = PARCEL_ALL) And (Ce.RATIO_TYPE = RATIO_RAW) Then
'            Set Ji = GetJobInOut(m_ProductPartUseds, Cpi.PART_ITEM_ID & "-E")
'            Tempsum = Tempsum + Ji.TX_AMOUNT
'
'         ElseIf (Ce.PACKAGE_TYPE = PARCEL_BAG) And (Ce.RATIO_TYPE = RATIO_VARY) And (Cpi.PARCEL_TYPE = PARCEL_BAG) Then
'            Tempsum = Tempsum + Cpi.COST_AMOUNT
'         ElseIf (Ce.PACKAGE_TYPE = PARCEL_BULK) And (Ce.RATIO_TYPE = RATIO_VARY) And (Cpi.PARCEL_TYPE = PARCEL_BULK) Then
'            Tempsum = Tempsum + Cpi.COST_AMOUNT
'         ElseIf (Ce.PACKAGE_TYPE = PARCEL_ALL) And (Ce.RATIO_TYPE = RATIO_VARY) Then
'            Tempsum = Tempsum + Cpi.COST_AMOUNT
'         ElseIf (Ce.PACKAGE_TYPE = PARCEL_BAG) And (Ce.RATIO_TYPE = RATIO_PERCENT) And (Cpi.PARCEL_TYPE = PARCEL_BAG) Then
'            Tempsum = Tempsum + Cpi.COST_AMOUNT
'            NetRaw = NetRaw + Cpi.RAW_AMOUNT
'         ElseIf (Ce.PACKAGE_TYPE = PARCEL_BULK) And (Ce.RATIO_TYPE = RATIO_PERCENT) And (Cpi.PARCEL_TYPE = PARCEL_BULK) Then
'            Tempsum = Tempsum + Cpi.COST_AMOUNT
'            NetRaw = NetRaw + Cpi.RAW_AMOUNT
'         ElseIf (Ce.PACKAGE_TYPE = PARCEL_ALL) And (Ce.RATIO_TYPE = RATIO_PERCENT) Then
'            Tempsum = Tempsum + Cpi.COST_AMOUNT
'            NetRaw = NetRaw + Cpi.RAW_AMOUNT
'         End If
'      End If
'   Next Cpi
'
'   GetTotalRelateAmount = Tempsum
End Function

Private Sub ShareExpenseAmount(Ce As CCostExpense, TotalAmount As Double, Optional NetRaw As Double)
'Dim Cpi As CCostPrdItem
'Dim ExpenseAmount As Double
'Dim FoundFlag As Boolean
'Dim Ci As CCostItem
'Dim Ji As CJobInput
'Dim ExpenseFoundFlag As Boolean
'Dim Cir As CCostItemRaw
'Dim ItemRawFlag As Boolean
'Dim Tempsum As Double
'
'   For Each Cpi In m_PackProduction.PackItems
'      If Cpi.Flag <> "D" Then
'         FoundFlag = False
'         ItemRawFlag = False
'         If (Ce.PACKAGE_TYPE = PARCEL_BAG) And (Ce.RATIO_TYPE = RATIO_QUANTITY) And (Cpi.PARCEL_TYPE = PARCEL_BAG) Then
'            ExpenseAmount = MyDiffEx(Cpi.COST_AMOUNT, TotalAmount) * Ce.EXPENSE_AMOUNT
'            FoundFlag = True
'            '''Debug.Print (Cpi.PART_NO)
'
'            'TempSum = TempSum + ExpenseAmount
'         ElseIf (Ce.PACKAGE_TYPE = PARCEL_BULK) And (Ce.RATIO_TYPE = RATIO_QUANTITY) And (Cpi.PARCEL_TYPE = PARCEL_BULK) Then
'            ExpenseAmount = MyDiffEx(Cpi.COST_AMOUNT, TotalAmount) * Ce.EXPENSE_AMOUNT
'            FoundFlag = True
'         ElseIf (Ce.PACKAGE_TYPE = PARCEL_ALL) And (Ce.RATIO_TYPE = RATIO_QUANTITY) Then
'            ExpenseAmount = MyDiffEx(Cpi.COST_AMOUNT, TotalAmount) * Ce.EXPENSE_AMOUNT
'            FoundFlag = True
'
'         ElseIf (Ce.PACKAGE_TYPE = PARCEL_BAG) And (Ce.RATIO_TYPE = RATIO_VARY) And (Cpi.PARCEL_TYPE = PARCEL_BAG) Then
'            ExpenseAmount = Cpi.COST_AMOUNT * Ce.EXPENSE_AMOUNT / 1000
'            FoundFlag = True
'         ElseIf (Ce.PACKAGE_TYPE = PARCEL_BULK) And (Ce.RATIO_TYPE = RATIO_VARY) And (Cpi.PARCEL_TYPE = PARCEL_BULK) Then
'            ExpenseAmount = Cpi.COST_AMOUNT * Ce.EXPENSE_AMOUNT / 1000
'            FoundFlag = True
'         ElseIf (Ce.PACKAGE_TYPE = PARCEL_ALL) And (Ce.RATIO_TYPE = RATIO_VARY) Then
'            ExpenseAmount = Cpi.COST_AMOUNT * Ce.EXPENSE_AMOUNT / 1000
'            FoundFlag = True
'
'         ElseIf (Ce.PACKAGE_TYPE = PARCEL_BAG) And (Ce.RATIO_TYPE = RATIO_COST) And (Cpi.PARCEL_TYPE = PARCEL_BAG) Then
'            ExpenseAmount = MyDiffEx(Cpi.RAW_AMOUNT, TotalAmount) * Ce.EXPENSE_AMOUNT
'            FoundFlag = True
'         ElseIf (Ce.PACKAGE_TYPE = PARCEL_BULK) And (Ce.RATIO_TYPE = RATIO_COST) And (Cpi.PARCEL_TYPE = PARCEL_BULK) Then
'            ExpenseAmount = MyDiffEx(Cpi.RAW_AMOUNT, TotalAmount) * Ce.EXPENSE_AMOUNT
'            FoundFlag = True
'         ElseIf (Ce.PACKAGE_TYPE = PARCEL_ALL) And (Ce.RATIO_TYPE = RATIO_COST) Then
'            ExpenseAmount = MyDiffEx(Cpi.RAW_AMOUNT, TotalAmount) * Ce.EXPENSE_AMOUNT
'            FoundFlag = True
'
'         ElseIf (Ce.PACKAGE_TYPE = PARCEL_BAG) And (Ce.RATIO_TYPE = RATIO_RAW) And (Cpi.PARCEL_TYPE = PARCEL_BAG) Then
'            Set Ji = GetJobInOut(m_ProductPartUseds, Cpi.PART_ITEM_ID & "-E")
'            ExpenseAmount = MyDiffEx(Ji.TX_AMOUNT, TotalAmount) * Ce.EXPENSE_AMOUNT
'            FoundFlag = True
'            ItemRawFlag = True
'         ElseIf (Ce.PACKAGE_TYPE = PARCEL_BULK) And (Ce.RATIO_TYPE = RATIO_RAW) And (Cpi.PARCEL_TYPE = PARCEL_BULK) Then
'            Set Ji = GetJobInOut(m_ProductPartUseds, Cpi.PART_ITEM_ID & "-E")
'            'ExpenseAmount = MyDiffEx(Ji.TX_AMOUNT, TotalAmount) * Ce.EXPENSE_AMOUNT
'            ExpenseAmount = MyDiffEx(Ji.TX_AMOUNT, 1000) * Ce.EXPENSE_AMOUNT
'            FoundFlag = True
'            ItemRawFlag = True
'
'         ElseIf (Ce.PACKAGE_TYPE = PARCEL_ALL) And (Ce.RATIO_TYPE = RATIO_RAW) Then
'            Set Ji = GetJobInOut(m_ProductPartUseds, Cpi.PART_ITEM_ID & "-E")
'            ExpenseAmount = MyDiffEx(Ji.TX_AMOUNT, TotalAmount) * Ce.EXPENSE_AMOUNT
'            FoundFlag = True
'            ItemRawFlag = True
'         ElseIf (Ce.PACKAGE_TYPE = PARCEL_BAG) And (Ce.RATIO_TYPE = RATIO_PERCENT) And (Cpi.PARCEL_TYPE = PARCEL_BAG) Then
'            ExpenseAmount = MyDiffEx(Cpi.RAW_AMOUNT, NetRaw) * NetRaw * Ce.EXPENSE_AMOUNT / 100
'            FoundFlag = True
'         ElseIf (Ce.PACKAGE_TYPE = PARCEL_BULK) And (Ce.RATIO_TYPE = RATIO_PERCENT) And (Cpi.PARCEL_TYPE = PARCEL_BULK) Then
'            ExpenseAmount = MyDiffEx(Cpi.RAW_AMOUNT, NetRaw) * NetRaw * Ce.EXPENSE_AMOUNT / 100
'            FoundFlag = True
'         ElseIf (Ce.PACKAGE_TYPE = PARCEL_ALL) And (Ce.RATIO_TYPE = RATIO_PERCENT) Then
'            ExpenseAmount = MyDiffEx(Cpi.RAW_AMOUNT, NetRaw) * NetRaw * Ce.EXPENSE_AMOUNT / 100
'            'ExpenseAmount = MyDiffEx(Cpi.COST_AMOUNT, TotalAmount) * NetRaw * Ce.EXPENSE_AMOUNT / 100
'            FoundFlag = True
'         End If
'
'         If FoundFlag Then
'            ExpenseFoundFlag = False
'            For Each Ci In Cpi.CostItems
'               If Ci.Flag <> "D" Then
'                  If Ci.PARAM_PROCESS_ID = Ce.EXPENSE_TYPE Then
'                     ExpenseFoundFlag = True
'                     Exit For
'                  End If
'               End If
'            Next Ci
'
'
'            If ExpenseFoundFlag Then
'               If Ci.Flag <> "A" Then
'                  Ci.Flag = "E"
'               End If
'               Ci.PARAM_PROCESS_ID = Ce.EXPENSE_TYPE
'               Ci.ITEM_COST = ExpenseAmount
'               'Tempsum = Tempsum + Ci.ITEM_COST
'               If ItemRawFlag Then
'                  Ci.ITEM_COST = GetExpenese(Ce, Cpi.PART_ITEM_ID, Ci)
'               End If
'            Else
'               Set Ci = New CCostItem
'               Ci.Flag = "A"
'               Ci.PARAM_PROCESS_ID = Ce.EXPENSE_TYPE
'               Ci.ITEM_COST = ExpenseAmount
'               Call Cpi.CostItems.add(Ci)
'
'               If ItemRawFlag Then
'                  Ci.ITEM_COST = GetExpenese(Ce, Cpi.PART_ITEM_ID, Ci)
'               End If
'               Set Ci = Nothing
'            End If
'         End If
'      End If
'   Next Cpi
End Sub

Private Sub GenerateCostPrdItem()
'Dim Ce As CCostExpense
'Dim TotalAmount As Double
'Dim PartItemSet As String
'Dim NetRaw As Double
'   For Each Ce In m_PackProduction.PackItems
'      If Ce.Flag <> "D" Then
'         If Ce.RATIO_TYPE = RATIO_RAW Then
'            Set m_ProductPartUseds = Nothing
'            Set m_ProductPartUseds = New Collection
'
'            PartItemSet = GeneratePartItemSet(Ce.CostRaws)
'            'อาหารแต่ละเบอร์มีจำนวน มูลค่าเท่าใด โดยที่ใช้วัตถุดิบตามที่อยู่ในไซโล
'            Call LoadJobProductRMAmount(Nothing, m_ProductPartUseds, uctlStartJob.ShowDate, uctlFinishJob.ShowDate, , , "E", PartItemSet)
'
'            'ยอดใช้วัตถุดิบตามอาหาร
'            Call LoadProductPartUsed(Nothing, m_ProductPartUsed, uctlStartJob.ShowDate, uctlFinishJob.ShowDate, , , , "E", PartItemSet)
'         End If
'
'         'ดูว่ามีจำนวน อาหารที่ผลิตออกมา ตรงตามเงื่อนไขใน Ce เป็นจำนวนเท่าใด
'         TotalAmount = GetTotalRelateAmount(Ce, NetRaw)
'
'         'คำนวณหาสัดส่วนว่าจะกระจายจำนวนเงินในแต่ละเบอร์เป็นจำนวนเงินเท่าใด
'         Call ShareExpenseAmount(Ce, TotalAmount, NetRaw)
'
'      End If
'   Next Ce
'
'   Call CalculateSumExpense
End Sub
Private Sub CalculateSumExpense()
'Dim Cpi As CCostPrdItem
'Dim Ci As CCostItem
'Dim SumExpense As Double
'Dim Tempsum As Double
'
'   For Each Cpi In m_PackProduction.CostItems
'      If Cpi.Flag <> "D" Then
'         SumExpense = 0
'         For Each Ci In Cpi.CostItems
'            If Ci.Flag <> "D" Then
'               SumExpense = SumExpense + Ci.ITEM_COST
'            End If
'         Next Ci
'
'
'         Cpi.EXPENSE_AMOUNT = SumExpense
'         'Tempsum = Tempsum + SumExpense
'         If Cpi.Flag <> "A" Then
'            Cpi.Flag = "E"
'         End If
'      End If
'   Next Cpi
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      
'      Call LoadPartItem(Nothing, m_PartItems)
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_PackProduction.QueryFlag = 1
         txtPackNo.Enabled = False
         cmdAuto.Enabled = False
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         uctlPackDate.ShowDate = Now
         
        m_PackProduction.QueryFlag = 0
         Call QueryData(False)
      End If
      
      Call TabStrip1_Click
      Call EnableForm(Me, True)
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
      Call cmdAdd_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
      Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 123 Then
      Call AddMemoNote
      KeyCode = 0
   End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.HEIGHT = ScaleHeight
   pnlHeader.Width = ScaleWidth
   GridEX1.Width = ScaleWidth - 2 * GridEX1.Left
   GridEX1.HEIGHT = ScaleHeight - GridEX1.Top - 620
   TabStrip1.Width = GridEX1.Width
   cmdAdd.Top = ScaleHeight - 580
   cmdEdit.Top = ScaleHeight - 580
   cmdDelete.Top = ScaleHeight - 580
   cmdPrint.Top = ScaleHeight - 580
   cmdOK.Top = ScaleHeight - 580
   cmdExit.Top = ScaleHeight - 580
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
   cmdPrint.Left = cmdExit.Left - cmdOK.Width - cmdOK.Width - 50
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_PackProduction = Nothing
   Set m_PackProductions = Nothing
   Set m_ExtractItems = Nothing
   Set m_PartItems = Nothing
   Set m_ProductPartUseds = Nothing
   Set m_ProductPartUsed = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid1()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.NAME = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
'   GridEX1.Font.Bold = False
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"
   GridEX1.Columns.Item(1).Visible = False

   Set Col = GridEX1.Columns.add '2
   Col.Width = 1000
   Col.Caption = "ลำดับ"
'   GridEX1.Columns.Item(2).Visible = False
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 3570
   Col.Caption = MapText("เบอร์อาหาร")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 1200
   Col.Caption = MapText("จำนวน")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 2760
   Col.Caption = MapText("ชนิดถุง")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 1500
   Col.Caption = MapText("ขนาด (กก.)")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 1500
   Col.Caption = MapText("จำนวนถุง")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 3000
   Col.Caption = MapText("ป้ายบ่งชี้สีเหลือง")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 3000
   Col.Caption = MapText("ป้ายบ่งชี้สีเขียว")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 1500
   Col.Caption = MapText("ด้ายที่เย็บ")

   Set Col = GridEX1.Columns.add '6
   Col.Width = 4000
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("หมายเหตุ")
End Sub


Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitGrid1
   Call InitNormalLabel(lblPackNo, MapText("เลขที่เอกสาร"))
   Call InitNormalLabel(lblDesc, MapText("รายละเอียด"))
   Call InitNormalLabel(lblPackDate, MapText("วันที่เอกสาร"))
   
   Call txtPackNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtDesc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPrint.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSave.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAuto.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSave, MapText("บันทึก"))
   Call InitMainButton(cmdAuto, MapText("A"))
   Call InitMainButton(cmdPrint, MapText("พิมพ์"))
   'cmdAuto
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.NAME = GLB_FONT
   TabStrip1.Font.Size = 16
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.add().Caption = MapText("รายการอาหาร")
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

   m_HasActivate = False
   m_HasModify = False
   
   Set m_Rs = New ADODB.Recordset
   Set m_PackProduction = New CPackProduction
   Set m_PackProductions = New Collection
   Set m_ExtractItems = New Collection
   Set m_PartItems = New Collection
   Set m_ProductPartUseds = New Collection
   Set m_ProductPartUsed = New Collection
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   If TabStrip1.SelectedItem.Index = 6 Then
      RowBuffer.RowStyle = RowBuffer.Value(6)
   End If
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Index = 1 Then
     If m_PackProduction.PackItems Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim PPI As CPackProductionItem
      If m_PackProduction.PackItems.Count <= 0 Then
         Exit Sub
      End If
      Set PPI = GetItem(m_PackProduction.PackItems, RowIndex, RealIndex)
      If PPI Is Nothing Then
         Exit Sub
      End If

      Values(1) = PPI.PACK_PRODUCTION_ITEM_ID
      Values(2) = RealIndex
      Values(3) = PPI.PART_ITEM_ID
      Values(4) = PPI.TX_AMOUNT
      Values(5) = PPI.PART_DESC
      If PPI.WEIGHT_PER_PACK = 1 Then
         Values(6) = "30"
      ElseIf PPI.WEIGHT_PER_PACK = 2 Then
         Values(6) = "50"
      Else
         Values(6) = ""
      End If
      Values(7) = PPI.PACK_AMOUNT
      Values(8) = PPI.PALLET_LABEL_YELLOW
      Values(9) = PPI.PALLET_LABEL_GREEN
      If PPI.SEWING_THREAD = 1 Then
         Values(10) = "ขาว"
      ElseIf PPI.SEWING_THREAD = 2 Then
         Values(10) = "ขาว-แดง"
      Else
         Values(10) = ""
      End If
      Values(11) = PPI.NOTE
   End If
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub AllVisibleFalse()
cmdAdd.Visible = False
cmdEdit.Visible = False
cmdDelete.Visible = False
End Sub

Private Sub TabStrip1_Click()

   If TabStrip1.SelectedItem.Index = 1 Then
      cmdAdd.Enabled = True
      cmdEdit.Enabled = True
      cmdDelete.Enabled = True
      
      Call InitGrid1
     GridEX1.ItemCount = CountItem(m_PackProduction.PackItems)
      GridEX1.Rebind
   End If
End Sub

Private Sub txtBatchNo_Change()
   m_HasModify = True
End Sub

Private Sub txtJobDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtPackNo_Change()
   m_HasModify = True
End Sub

Private Sub ucltApproveByLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlFinishJob_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlPackDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlResponseByLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlStartJob_HasChange()
   m_HasModify = True
End Sub
Private Function GetExpenese(Ce As CCostExpense, TempID As Long, Ci As CCostItem) As Double
Dim CR As CCostRaw
Dim Cir As CCostItemRaw
Dim Ji As CJobInput
Dim Expense As Double
   Expense = 0
   
   For Each Cir In Ci.CostItemRaws
      Cir.Flag = "D"
   Next
   
   For Each CR In Ce.CostRaws
      Set Ji = GetObject("CJobInput", m_ProductPartUsed, CR.GetFieldValue("PART_ITEM_ID") & "-" & TempID)
      Expense = Expense + (Ji.TX_AMOUNT * Ce.EXPENSE_AMOUNT / 1000)
      'Set Cir = GetObject("CCostItemRaw", Ci.CostItemRaws, Trim(Ci.COST_ITEM_ID & "-" & CR.GetFieldValue("PART_ITEM_ID")))
'      If Cir.Flag = "I" Then
'         Cir.Flag = "E"
'      Else
      Set Cir = New CCostItemRaw
      Cir.Flag = "A"
      Call Ci.CostItemRaws.add(Cir, Trim(str(CR.GetFieldValue("PART_ITEM_ID"))))
'      End If
      Call Cir.SetFieldValue("PART_ITEM_ID", CR.GetFieldValue("PART_ITEM_ID"))
      Call Cir.SetFieldValue("ITEM_COST", Ji.TX_AMOUNT * Ce.EXPENSE_AMOUNT / 1000)
      Call Cir.SetFieldValue("ITEM_AMOUNT", Ji.TX_AMOUNT)
   Next CR
   GetExpenese = Expense
End Function
