VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditCostProduction 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditCostProduction.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8895
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   12720
      _ExtentX        =   22437
      _ExtentY        =   15690
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   120
         TabIndex        =   7
         Top             =   2970
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
         TabIndex        =   16
         Top             =   0
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   4230
         Left            =   120
         TabIndex        =   8
         Top             =   3510
         Width           =   11640
         _ExtentX        =   20532
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
         Column(1)       =   "frmAddEditCostProduction.frx":27A2
         Column(2)       =   "frmAddEditCostProduction.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditCostProduction.frx":290E
         FormatStyle(2)  =   "frmAddEditCostProduction.frx":2A6A
         FormatStyle(3)  =   "frmAddEditCostProduction.frx":2B1A
         FormatStyle(4)  =   "frmAddEditCostProduction.frx":2BCE
         FormatStyle(5)  =   "frmAddEditCostProduction.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditCostProduction.frx":2D5E
      End
      Begin prjFarmManagement.uctlTextBox txtJobNo 
         Height          =   435
         Left            =   1800
         TabIndex        =   0
         Top             =   990
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlJobDate 
         Height          =   405
         Left            =   7500
         TabIndex        =   1
         Top             =   990
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlDate uctlFinishJob 
         Height          =   405
         Left            =   7500
         TabIndex        =   3
         Top             =   1440
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlDate uctlStartJob 
         Height          =   405
         Left            =   1800
         TabIndex        =   2
         Top             =   1440
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   1800
         TabIndex        =   4
         Top             =   1920
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin prjFarmManagement.uctlTextBox txtProgress 
         Height          =   435
         Left            =   10170
         TabIndex        =   5
         Top             =   1860
         Width           =   1185
         _ExtentX        =   4683
         _ExtentY        =   767
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   1800
         TabIndex        =   6
         Top             =   2250
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCostProduction.frx":2F36
         ButtonStyle     =   3
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "lblJobNo"
         Height          =   315
         Left            =   330
         TabIndex        =   21
         Top             =   1950
         Width           =   1395
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   525
         Left            =   6840
         TabIndex        =   12
         Top             =   7830
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCostProduction.frx":3250
         ButtonStyle     =   3
      End
      Begin VB.Label lblFinishJob 
         Alignment       =   1  'Right Justify
         Caption         =   "lblFinishJob"
         Height          =   315
         Left            =   5760
         TabIndex        =   20
         Top             =   1560
         Width           =   1665
      End
      Begin VB.Label lblStartJob 
         Alignment       =   1  'Right Justify
         Caption         =   "lblStartJob"
         Height          =   315
         Left            =   420
         TabIndex        =   19
         Top             =   1560
         Width           =   1305
      End
      Begin VB.Label lblJobDate 
         Alignment       =   1  'Right Justify
         Caption         =   "lblJobDate"
         Height          =   315
         Left            =   6120
         TabIndex        =   18
         Top             =   1050
         Width           =   1305
      End
      Begin VB.Label lblJobNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblJobNo"
         Height          =   315
         Left            =   330
         TabIndex        =   17
         Top             =   1110
         Width           =   1395
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8490
         TabIndex        =   13
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCostProduction.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10170
         TabIndex        =   14
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1740
         TabIndex        =   10
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   120
         TabIndex        =   9
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCostProduction.frx":3884
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3390
         TabIndex        =   11
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCostProduction.frx":3B9E
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditCostProduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_CostProduction As CCostProduction
Private m_CostProductions As Collection

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
'         cmdAdd.Enabled = (m_CostProduction.COMMIT_FLAG = "N")
'         cmdDelete.Enabled = (m_CostProduction.COMMIT_FLAG = "N")
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
      
      m_CostProduction.COST_PRODUCTION_ID = id
      m_CostProduction.QueryFlag = 1
      If Not glbProduction.QueryCostProduction(m_CostProduction, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
        Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_CostProduction.PopulateFromRS(1, m_Rs)

      txtJobNo.Text = m_CostProduction.DOCUMENT_NO
      uctlJobDate.ShowDate = m_CostProduction.DOCUMENT_DATE
      uctlStartJob.ShowDate = m_CostProduction.JOB_FROM_DATE
      uctlFinishJob.ShowDate = m_CostProduction.JOB_TO_DATE
      cmdAdd.Enabled = (m_CostProduction.CostItems.Count <= 0)
      
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
      If Not VerifyAccessRight("PRODUCT_ESTIMATE_EDIT") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   End If
   If Not VerifyTextControl(lblJobNo, txtJobNo, False) Then
       Exit Function
   End If
   If Not VerifyDate(lblJobDate, uctlJobDate, False) Then
      Exit Function
   End If
      
   If Not VerifyDate(lblStartJob, uctlStartJob, True) Then
     Exit Function
   End If
   If Not VerifyDate(lblFinishJob, uctlFinishJob, True) Then
     Exit Function
   End If
       
'   If Not CheckUniqueNs(JOB_NO, txtJobNo.Text, ID) Then
'      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtJobNo.Text & " " & MapText("อยู่ในระบบแล้ว")
'      glbErrorLog.ShowUserError
'      Exit Function
'   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_CostProduction.COST_PRODUCTION_ID = id
   m_CostProduction.AddEditMode = ShowMode
   m_CostProduction.DOCUMENT_NO = txtJobNo.Text
   m_CostProduction.DOCUMENT_DATE = uctlJobDate.ShowDate
   m_CostProduction.JOB_FROM_DATE = uctlStartJob.ShowDate
   m_CostProduction.JOB_TO_DATE = uctlFinishJob.ShowDate
   
   Call EnableForm(Me, False)
      
   If Not glbProduction.AddEditCostProduction(m_CostProduction, IsOK, True, glbErrorLog) Then
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

Public Sub RefreshGrid()

End Sub

Private Sub cmdAuto_Click()
Dim No As String

   If Trim(txtJobNo.Text) = "" Then
      If JobDocType = 1 Then
         Call glbDatabaseMngr.GenerateNumber(JOBPLAN_NUMBER, No, glbErrorLog)
         txtJobNo.Text = No
      ElseIf JobDocType = 2 Then
         Call glbDatabaseMngr.GenerateNumber(ESTIMATE_NUMBER, No, glbErrorLog)
         txtJobNo.Text = No
      End If
   End If
End Sub

Private Sub CalculateTotalRatio()
   
End Sub

Private Sub DeleteAllItem()
Dim Ei As CCostPrdItem

   For Each Ei In m_CostProduction.CostItems
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

   If Not VerifyDate(lblStartJob, uctlStartJob, False) Then
      Exit Sub
   End If
   If Not VerifyDate(lblFinishJob, uctlFinishJob, False) Then
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   
   Set TempRs = New ADODB.Recordset
   
   If TabStrip1.SelectedItem.Index = 1 Then
     Set frmAddEditCostExpense.TempCollection = m_CostProduction.ExpenseItem
      frmAddEditCostExpense.ShowMode = SHOW_ADD
      frmAddEditCostExpense.HeaderText = MapText("เพิ่มต้นทุนผลิต")
      frmAddEditCostExpense.FromDate = uctlStartJob.ShowDate
      frmAddEditCostExpense.ToDate = uctlFinishJob.ShowDate
      Load frmAddEditCostExpense
      frmAddEditCostExpense.Show 1

      OKClick = frmAddEditCostExpense.OKClick

      Unload frmAddEditCostExpense
      Set frmAddEditCostExpense = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_CostProduction.ExpenseItem)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      Set Ji = New CJobInput
      Ji.JOB_INOUT_ID = -1
      Ji.TX_TYPE = "I"
      'Ji.JOB_PART_ITEM_ID = 4126                             'DeleteIt
      Ji.FROM_DATE = uctlStartJob.ShowDate
      Ji.TO_DATE = uctlFinishJob.ShowDate
      Ji.JOB_DOC_TYPE = 1
      Call Ji.QueryData(2, TempRs, iCount)
         
'      Set m_CostProduction.CostItems = Nothing
'      Set m_CostProduction.CostItems = New Collection
      Call DeleteAllItem
      While Not TempRs.EOF
         'ต้องเอา PARCEL_TYPE มาด้วย
         Call Ji.PopulateFromRS(2, TempRs)
   '      Set Pi = GetPartItem(m_PartItems, Ji.PART_ITEM_ID)
         
         Set Ei = New CCostPrdItem
         Ei.Flag = "A"
         Ei.COST_AMOUNT = Ji.TX_AMOUNT
         Ei.RAW_AMOUNT = Ji.RAW_COST
         Ei.PART_ITEM_ID = Ji.PART_ITEM_ID
         Ei.PART_NO = Ji.PART_NO
         Ei.PART_DESC = Ji.PART_DESC
         Ei.PARCEL_TYPE = Ji.PARCEL_TYPE
         Call m_CostProduction.CostItems.add(Ei, Trim(str(Ji.PART_ITEM_ID)))
         
   '      Set TempEi = New CCostPrdItem
   '      TempEi.COST_AMOUNT = Ei.TOTAL_AMT
   '      TempEi.PART_ITEM_ID = Ei.PART_ITEM_ID
   '      Call m_ExtractItems.add(TempEi, Trim(Str(Pi.PART_ITEM_ID)))
   '      Set TempEi = Nothing
         
         Set Ei = Nothing
         
         TempRs.MoveNext
      Wend
      
      Set Ji = Nothing
      If TempRs.State = adStateOpen Then
         TempRs.Close
      End If
      Set TempRs = Nothing
      
      GridEX1.ItemCount = CountItem(m_CostProduction.CostItems)
      GridEX1.Rebind
   '   Call CalculateTotalAmount
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
         m_CostProduction.ExpenseItem.Remove (ID2)
      Else
         m_CostProduction.ExpenseItem.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_CostProduction.ExpenseItem)
      GridEX1.Rebind
      m_HasModify = True
    ElseIf TabStrip1.SelectedItem.Index = 2 Then
      If ID1 <= 0 Then
         m_CostProduction.CostItems.Remove (ID2)
      Else
         m_CostProduction.CostItems.Item(ID2).Flag = "D"
      End If

      Call CalculateTotalRatio
      GridEX1.ItemCount = CountItem(m_CostProduction.CostItems)
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
     Set frmAddEditCostExpense.TempCollection = m_CostProduction.ExpenseItem
      frmAddEditCostExpense.id = id
      frmAddEditCostExpense.ShowMode = SHOW_EDIT
      frmAddEditCostExpense.HeaderText = MapText("แก้ไขต้นทุนผลิต")
      Load frmAddEditCostExpense
      frmAddEditCostExpense.Show 1

      OKClick = frmAddEditCostExpense.OKClick

      Unload frmAddEditCostExpense
      Set frmAddEditCostExpense = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_CostProduction.ExpenseItem)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
     Set frmAddEditCostItem.TempCollection = m_CostProduction.CostItems
      frmAddEditCostItem.id = id
      frmAddEditCostItem.ShowMode = SHOW_EDIT
      frmAddEditCostItem.HeaderText = MapText("แก้ไขปริมาณต้นทุน")
      Load frmAddEditCostItem
      frmAddEditCostItem.Show 1

      OKClick = frmAddEditCostItem.OKClick

      Unload frmAddEditCostItem
      Set frmAddEditCostItem = Nothing

      If OKClick Then
         Call CalculateTotalRatio
         
         GridEX1.ItemCount = CountItem(m_CostProduction.CostItems)
         GridEX1.Rebind
      End If
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Sub cmdSave_Click()
Dim Result As Boolean
   If Not SaveData Then
      Exit Sub
   End If
   
   ShowMode = SHOW_EDIT
   id = m_CostProduction.COST_PRODUCTION_ID
   m_CostProduction.QueryFlag = 1
   QueryData (True)
   m_HasModify = False
End Sub

Private Function GetTotalRelateAmount(Ce As CCostExpense, Optional NetRaw As Double) As Double
Dim Cpi As CCostPrdItem
Dim Tempsum As Double
Dim Ji As CJobInput

   Tempsum = 0
   NetRaw = 0
   For Each Cpi In m_CostProduction.CostItems
      If Cpi.Flag <> "D" Then
      
         If (Ce.PACKAGE_TYPE = PARCEL_BAG) And (Ce.RATIO_TYPE = RATIO_QUANTITY) And (Cpi.PARCEL_TYPE = PARCEL_BAG) Then
            Tempsum = Tempsum + Cpi.COST_AMOUNT
         ElseIf (Ce.PACKAGE_TYPE = PARCEL_BULK) And (Ce.RATIO_TYPE = RATIO_QUANTITY) And (Cpi.PARCEL_TYPE = PARCEL_BULK) Then
            Tempsum = Tempsum + Cpi.COST_AMOUNT
         ElseIf (Ce.PACKAGE_TYPE = PARCEL_ALL) And (Ce.RATIO_TYPE = RATIO_QUANTITY) Then
            Tempsum = Tempsum + Cpi.COST_AMOUNT
         
         ElseIf (Ce.PACKAGE_TYPE = PARCEL_BAG) And (Ce.RATIO_TYPE = RATIO_COST) And (Cpi.PARCEL_TYPE = PARCEL_BAG) Then
            Tempsum = Tempsum + Cpi.RAW_AMOUNT
         ElseIf (Ce.PACKAGE_TYPE = PARCEL_BULK) And (Ce.RATIO_TYPE = RATIO_COST) And (Cpi.PARCEL_TYPE = PARCEL_BULK) Then
            Tempsum = Tempsum + Cpi.RAW_AMOUNT
         ElseIf (Ce.PACKAGE_TYPE = PARCEL_ALL) And (Ce.RATIO_TYPE = RATIO_COST) Then
            Tempsum = Tempsum + Cpi.RAW_AMOUNT
         
         ElseIf (Ce.PACKAGE_TYPE = PARCEL_BAG) And (Ce.RATIO_TYPE = RATIO_RAW) And (Cpi.PARCEL_TYPE = PARCEL_BAG) Then
            Set Ji = GetJobInOut(m_ProductPartUseds, Cpi.PART_ITEM_ID & "-E")
            Tempsum = Tempsum + Ji.TX_AMOUNT
         ElseIf (Ce.PACKAGE_TYPE = PARCEL_BULK) And (Ce.RATIO_TYPE = RATIO_RAW) And (Cpi.PARCEL_TYPE = PARCEL_BULK) Then
            Set Ji = GetJobInOut(m_ProductPartUseds, Cpi.PART_ITEM_ID & "-E")
            Tempsum = Tempsum + Ji.TX_AMOUNT
         ElseIf (Ce.PACKAGE_TYPE = PARCEL_ALL) And (Ce.RATIO_TYPE = RATIO_RAW) Then
            Set Ji = GetJobInOut(m_ProductPartUseds, Cpi.PART_ITEM_ID & "-E")
            Tempsum = Tempsum + Ji.TX_AMOUNT
         
         ElseIf (Ce.PACKAGE_TYPE = PARCEL_BAG) And (Ce.RATIO_TYPE = RATIO_VARY) And (Cpi.PARCEL_TYPE = PARCEL_BAG) Then
            Tempsum = Tempsum + Cpi.COST_AMOUNT
         ElseIf (Ce.PACKAGE_TYPE = PARCEL_BULK) And (Ce.RATIO_TYPE = RATIO_VARY) And (Cpi.PARCEL_TYPE = PARCEL_BULK) Then
            Tempsum = Tempsum + Cpi.COST_AMOUNT
         ElseIf (Ce.PACKAGE_TYPE = PARCEL_ALL) And (Ce.RATIO_TYPE = RATIO_VARY) Then
            Tempsum = Tempsum + Cpi.COST_AMOUNT
         ElseIf (Ce.PACKAGE_TYPE = PARCEL_BAG) And (Ce.RATIO_TYPE = RATIO_PERCENT) And (Cpi.PARCEL_TYPE = PARCEL_BAG) Then
            Tempsum = Tempsum + Cpi.COST_AMOUNT
            NetRaw = NetRaw + Cpi.RAW_AMOUNT
         ElseIf (Ce.PACKAGE_TYPE = PARCEL_BULK) And (Ce.RATIO_TYPE = RATIO_PERCENT) And (Cpi.PARCEL_TYPE = PARCEL_BULK) Then
            Tempsum = Tempsum + Cpi.COST_AMOUNT
            NetRaw = NetRaw + Cpi.RAW_AMOUNT
         ElseIf (Ce.PACKAGE_TYPE = PARCEL_ALL) And (Ce.RATIO_TYPE = RATIO_PERCENT) Then
            Tempsum = Tempsum + Cpi.COST_AMOUNT
            NetRaw = NetRaw + Cpi.RAW_AMOUNT
         End If
      End If
   Next Cpi
   
   GetTotalRelateAmount = Tempsum
End Function

Private Sub ShareExpenseAmount(Ce As CCostExpense, TotalAmount As Double, Optional NetRaw As Double)
Dim Cpi As CCostPrdItem
Dim ExpenseAmount As Double
Dim FoundFlag As Boolean
Dim Ci As CCostItem
Dim Ji As CJobInput
Dim ExpenseFoundFlag As Boolean
Dim Cir As CCostItemRaw
Dim ItemRawFlag As Boolean
Dim Tempsum As Double
   
   For Each Cpi In m_CostProduction.CostItems
      If Cpi.Flag <> "D" Then
         FoundFlag = False
         ItemRawFlag = False
         If (Ce.PACKAGE_TYPE = PARCEL_BAG) And (Ce.RATIO_TYPE = RATIO_QUANTITY) And (Cpi.PARCEL_TYPE = PARCEL_BAG) Then
            ExpenseAmount = MyDiffEx(Cpi.COST_AMOUNT, TotalAmount) * Ce.EXPENSE_AMOUNT
            FoundFlag = True
            '''Debug.Print (Cpi.PART_NO)
            
            'TempSum = TempSum + ExpenseAmount
         ElseIf (Ce.PACKAGE_TYPE = PARCEL_BULK) And (Ce.RATIO_TYPE = RATIO_QUANTITY) And (Cpi.PARCEL_TYPE = PARCEL_BULK) Then
            ExpenseAmount = MyDiffEx(Cpi.COST_AMOUNT, TotalAmount) * Ce.EXPENSE_AMOUNT
            FoundFlag = True
         ElseIf (Ce.PACKAGE_TYPE = PARCEL_ALL) And (Ce.RATIO_TYPE = RATIO_QUANTITY) Then
            ExpenseAmount = MyDiffEx(Cpi.COST_AMOUNT, TotalAmount) * Ce.EXPENSE_AMOUNT
            FoundFlag = True
            
         ElseIf (Ce.PACKAGE_TYPE = PARCEL_BAG) And (Ce.RATIO_TYPE = RATIO_VARY) And (Cpi.PARCEL_TYPE = PARCEL_BAG) Then
            ExpenseAmount = Cpi.COST_AMOUNT * Ce.EXPENSE_AMOUNT / 1000
            FoundFlag = True
         ElseIf (Ce.PACKAGE_TYPE = PARCEL_BULK) And (Ce.RATIO_TYPE = RATIO_VARY) And (Cpi.PARCEL_TYPE = PARCEL_BULK) Then
            ExpenseAmount = Cpi.COST_AMOUNT * Ce.EXPENSE_AMOUNT / 1000
            FoundFlag = True
         ElseIf (Ce.PACKAGE_TYPE = PARCEL_ALL) And (Ce.RATIO_TYPE = RATIO_VARY) Then
            ExpenseAmount = Cpi.COST_AMOUNT * Ce.EXPENSE_AMOUNT / 1000
            FoundFlag = True
            
         ElseIf (Ce.PACKAGE_TYPE = PARCEL_BAG) And (Ce.RATIO_TYPE = RATIO_COST) And (Cpi.PARCEL_TYPE = PARCEL_BAG) Then
            ExpenseAmount = MyDiffEx(Cpi.RAW_AMOUNT, TotalAmount) * Ce.EXPENSE_AMOUNT
            FoundFlag = True
         ElseIf (Ce.PACKAGE_TYPE = PARCEL_BULK) And (Ce.RATIO_TYPE = RATIO_COST) And (Cpi.PARCEL_TYPE = PARCEL_BULK) Then
            ExpenseAmount = MyDiffEx(Cpi.RAW_AMOUNT, TotalAmount) * Ce.EXPENSE_AMOUNT
            FoundFlag = True
         ElseIf (Ce.PACKAGE_TYPE = PARCEL_ALL) And (Ce.RATIO_TYPE = RATIO_COST) Then
            ExpenseAmount = MyDiffEx(Cpi.RAW_AMOUNT, TotalAmount) * Ce.EXPENSE_AMOUNT
            FoundFlag = True
            
         ElseIf (Ce.PACKAGE_TYPE = PARCEL_BAG) And (Ce.RATIO_TYPE = RATIO_RAW) And (Cpi.PARCEL_TYPE = PARCEL_BAG) Then
            Set Ji = GetJobInOut(m_ProductPartUseds, Cpi.PART_ITEM_ID & "-E")
            ExpenseAmount = MyDiffEx(Ji.TX_AMOUNT, TotalAmount) * Ce.EXPENSE_AMOUNT
            FoundFlag = True
            ItemRawFlag = True
         ElseIf (Ce.PACKAGE_TYPE = PARCEL_BULK) And (Ce.RATIO_TYPE = RATIO_RAW) And (Cpi.PARCEL_TYPE = PARCEL_BULK) Then
            Set Ji = GetJobInOut(m_ProductPartUseds, Cpi.PART_ITEM_ID & "-E")
            'ExpenseAmount = MyDiffEx(Ji.TX_AMOUNT, TotalAmount) * Ce.EXPENSE_AMOUNT
            ExpenseAmount = MyDiffEx(Ji.TX_AMOUNT, 1000) * Ce.EXPENSE_AMOUNT
            FoundFlag = True
            ItemRawFlag = True
            
         ElseIf (Ce.PACKAGE_TYPE = PARCEL_ALL) And (Ce.RATIO_TYPE = RATIO_RAW) Then
            Set Ji = GetJobInOut(m_ProductPartUseds, Cpi.PART_ITEM_ID & "-E")
            ExpenseAmount = MyDiffEx(Ji.TX_AMOUNT, TotalAmount) * Ce.EXPENSE_AMOUNT
            FoundFlag = True
            ItemRawFlag = True
         ElseIf (Ce.PACKAGE_TYPE = PARCEL_BAG) And (Ce.RATIO_TYPE = RATIO_PERCENT) And (Cpi.PARCEL_TYPE = PARCEL_BAG) Then
            ExpenseAmount = MyDiffEx(Cpi.RAW_AMOUNT, NetRaw) * NetRaw * Ce.EXPENSE_AMOUNT / 100
            FoundFlag = True
         ElseIf (Ce.PACKAGE_TYPE = PARCEL_BULK) And (Ce.RATIO_TYPE = RATIO_PERCENT) And (Cpi.PARCEL_TYPE = PARCEL_BULK) Then
            ExpenseAmount = MyDiffEx(Cpi.RAW_AMOUNT, NetRaw) * NetRaw * Ce.EXPENSE_AMOUNT / 100
            FoundFlag = True
         ElseIf (Ce.PACKAGE_TYPE = PARCEL_ALL) And (Ce.RATIO_TYPE = RATIO_PERCENT) Then
            ExpenseAmount = MyDiffEx(Cpi.RAW_AMOUNT, NetRaw) * NetRaw * Ce.EXPENSE_AMOUNT / 100
            'ExpenseAmount = MyDiffEx(Cpi.COST_AMOUNT, TotalAmount) * NetRaw * Ce.EXPENSE_AMOUNT / 100
            FoundFlag = True
         End If
         
         If FoundFlag Then
            ExpenseFoundFlag = False
            For Each Ci In Cpi.CostItems
               If Ci.Flag <> "D" Then
                  If Ci.PARAM_PROCESS_ID = Ce.EXPENSE_TYPE Then
                     ExpenseFoundFlag = True
                     Exit For
                  End If
               End If
            Next Ci
            
            
            If ExpenseFoundFlag Then
               If Ci.Flag <> "A" Then
                  Ci.Flag = "E"
               End If
               Ci.PARAM_PROCESS_ID = Ce.EXPENSE_TYPE
               Ci.ITEM_COST = ExpenseAmount
               'Tempsum = Tempsum + Ci.ITEM_COST
               If ItemRawFlag Then
                  Ci.ITEM_COST = GetExpenese(Ce, Cpi.PART_ITEM_ID, Ci)
               End If
            Else
               Set Ci = New CCostItem
               Ci.Flag = "A"
               Ci.PARAM_PROCESS_ID = Ce.EXPENSE_TYPE
               Ci.ITEM_COST = ExpenseAmount
               Call Cpi.CostItems.add(Ci)
               
               If ItemRawFlag Then
                  Ci.ITEM_COST = GetExpenese(Ce, Cpi.PART_ITEM_ID, Ci)
               End If
               Set Ci = Nothing
            End If
         End If
      End If
   Next Cpi
End Sub

Private Sub GenerateCostPrdItem()
Dim Ce As CCostExpense
Dim TotalAmount As Double
Dim PartItemSet As String
Dim NetRaw As Double
   For Each Ce In m_CostProduction.ExpenseItem
      If Ce.Flag <> "D" Then
         If Ce.RATIO_TYPE = RATIO_RAW Then
            Set m_ProductPartUseds = Nothing
            Set m_ProductPartUseds = New Collection
            
            PartItemSet = GeneratePartItemSet(Ce.CostRaws)
            'อาหารแต่ละเบอร์มีจำนวน มูลค่าเท่าใด โดยที่ใช้วัตถุดิบตามที่อยู่ในไซโล
            Call LoadJobProductRMAmount(Nothing, m_ProductPartUseds, uctlStartJob.ShowDate, uctlFinishJob.ShowDate, , , "E", PartItemSet)
            
            'ยอดใช้วัตถุดิบตามอาหาร
            Call LoadProductPartUsed(Nothing, m_ProductPartUsed, uctlStartJob.ShowDate, uctlFinishJob.ShowDate, , , , "E", PartItemSet)
         End If
         
         'ดูว่ามีจำนวน อาหารที่ผลิตออกมา ตรงตามเงื่อนไขใน Ce เป็นจำนวนเท่าใด
         TotalAmount = GetTotalRelateAmount(Ce, NetRaw)
         
         'คำนวณหาสัดส่วนว่าจะกระจายจำนวนเงินในแต่ละเบอร์เป็นจำนวนเงินเท่าใด
         Call ShareExpenseAmount(Ce, TotalAmount, NetRaw)
         
      End If
   Next Ce
   
   Call CalculateSumExpense
End Sub
Private Sub CalculateSumExpense()
Dim Cpi As CCostPrdItem
Dim Ci As CCostItem
Dim SumExpense As Double
Dim Tempsum As Double

   For Each Cpi In m_CostProduction.CostItems
      If Cpi.Flag <> "D" Then
         SumExpense = 0
         For Each Ci In Cpi.CostItems
            If Ci.Flag <> "D" Then
               SumExpense = SumExpense + Ci.ITEM_COST
            End If
         Next Ci
         
         
         Cpi.EXPENSE_AMOUNT = SumExpense
         'Tempsum = Tempsum + SumExpense
         If Cpi.Flag <> "A" Then
            Cpi.Flag = "E"
         End If
      End If
   Next Cpi
End Sub
Private Sub cmdStart_Click()
Dim Jb As CJob
Dim TempRs As ADODB.Recordset
Dim TempJb As CJob
Dim IsOK As Boolean
Dim iCount As Long
Dim Ji As CJobInput
Dim Ei1 As CExtractItem
Dim Ei2 As CExtractItem
Dim Ratio As Double
Dim Ivd As CInventoryDoc
Dim RCount As Long
Dim I As Long
Dim Percent As Double
Dim TempCol As Collection
Dim Tempsum As Double

   If CountItem(m_CostProduction.CostItems) <= 0 Then
      glbErrorLog.LocalErrorMsg = "กรุณาทำการเพิ่มปริมาณผลิตภัณฑ์ก่อน (กดปุ่มเพิ่ม)"
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   If CountItem(m_CostProduction.ExpenseItem) <= 0 Then
      glbErrorLog.LocalErrorMsg = "กรุณาทำการเพิ่มต้นทุนผลิตก่อน (กดปุ่มเพิ่ม)"
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   If m_HasModify Or (m_CostProduction.COST_PRODUCTION_ID <= 0) Then
      glbErrorLog.LocalErrorMsg = MapText("กรุณาทำการบันทึกข้อมูลให้เรียบร้อยก่อน")
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   Set TempRs = New ADODB.Recordset
   
   Set Jb = New CJob
   Jb.JOB_ID = -1
   'Jb.PART_ITEM_ID = 4126
   Jb.FROM_DATE = uctlStartJob.ShowDate
   Jb.TO_DATE = uctlFinishJob.ShowDate
   Call Jb.QueryData(1, TempRs, RCount)
   
   Call glbDaily.StartTransaction
   I = 0
   prgProgress.MAX = 100
   prgProgress.MIN = 0
   
   Call m_CostProduction.DeleteDataCstItemRaw
   
   Call GenerateCostPrdItem
   
   Call Jb.DeleteDataParameter
   
   While Not TempRs.EOF
      I = I + 1
      Percent = MyDiffEx(I, RCount) * 100
      prgProgress.Value = Percent
      prgProgress.Refresh
      txtProgress.Text = FormatNumber(Percent)
      DoEvents

      Call Jb.PopulateFromRS(1, TempRs)

      Set TempJb = New CJob
      TempJb.JOB_ID = Jb.JOB_ID
      TempJb.QueryFlag = 1
      Call glbProduction.QueryJob(TempJb, m_Rs, iCount, IsOK, glbErrorLog)
      If Not m_Rs.EOF Then
          Call TempJb.PopulateFromRS(1, m_Rs)
      End If
      
      Tempsum = Tempsum + GenerateJobParameter(TempJb)
      
'      Call glbDaily.Job2InventoryDoc(TempJb, Ivd, 1)
'      Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog)
'      TempJb.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID
      TempJb.AddEditMode = SHOW_EDIT
      Call glbProduction.AddEditJob(TempJb, IsOK, False, glbErrorLog)

      Set TempJb = Nothing
      TempRs.MoveNext
   Wend
   
   Call glbDaily.CommitTransaction
   
   Set Jb = Nothing
   If TempRs.State = adStateOpen Then
      Call TempRs.Close
   End If
   Set TempRs = Nothing
   
   Call TabStrip1_Click
   m_HasModify = True
   Call EnableForm(Me, True)
End Sub

Public Sub FindJobActualAmount(Jb As CJob, Amt As Double)
Dim Ji As CJobInput
Dim TmpAmt As Double

   TmpAmt = 0
   For Each Ji In Jb.Outputs
      If (Ji.Flag <> "D") And (Ji.PART_ITEM_ID = Jb.PART_ITEM_ID) Then
         TmpAmt = Ji.TX_AMOUNT
         Exit For
      End If
   Next Ji
   
   Amt = TmpAmt
   Jb.ACTUAL_AMOUNT = Amt
End Sub

Private Function GenerateJobParameter(Jb As CJob) As Double
Dim Cpi As CCostPrdItem
Dim Ci As CCostItem
Dim Ratio As Double
Dim Jp As CJobParameter
Dim Amt As Double
Dim Tempsum As Double
   For Each Cpi In m_CostProduction.CostItems
      If Cpi.PART_ITEM_ID = Jb.PART_ITEM_ID Then
         Exit For
      End If
   Next Cpi
   
'   If Jb.PART_ITEM_ID = 4126 Then
'      ''Debug.Print (Jb.JOB_ID)
'   End If
   
   If Not (Cpi Is Nothing) Then
      Call FindJobActualAmount(Jb, Amt)
      'jb.actual_amount ถูก modify ค่าแล้ว
      Ratio = Jb.ACTUAL_AMOUNT / Cpi.COST_AMOUNT
   Else
      Ratio = 0
   End If
   
   For Each Jp In Jb.Parameters
      Jp.Flag = "D"
   Next Jp

'   Set Jb.Parameters = Nothing
'   Set Jb.Parameters = New Collection
   
   Tempsum = 0
   
   If Not (Cpi Is Nothing) Then
      For Each Ci In Cpi.CostItems
         Set Jp = New CJobParameter
         Jp.Flag = "A"
         Jp.PARAMETER_PROCESS_ID = Ci.PARAM_PROCESS_ID
         Jp.PARAM_AMOUNT = Ci.ITEM_COST * Ratio
'         If Jb.PART_ITEM_ID = 4126 Then
'            Tempsum = Tempsum + Jp.PARAM_AMOUNT
'         End If
         
         Call Jb.Parameters.add(Jp)
         Set Jp = Nothing
      Next Ci
   End If
   
   GenerateJobParameter = Tempsum
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      
'      Call LoadPartItem(Nothing, m_PartItems)
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_CostProduction.QueryFlag = 1
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         uctlJobDate.ShowDate = Now
         
        m_CostProduction.QueryFlag = 0
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

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_CostProduction = Nothing
   Set m_CostProductions = Nothing
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
   Col.Width = 0
   Col.Caption = "Real ID"
   GridEX1.Columns.Item(2).Visible = False
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 3570
   Col.Caption = MapText("ต้นทุนผลิต")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 3045
   Col.Caption = MapText("ปันให้กับ")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 2760
   Col.Caption = MapText("อัตราส่วนตาม")

   Set Col = GridEX1.Columns.add '6
   Col.Width = 2220
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("มูลค่า")
End Sub

Private Sub InitGrid2()
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
   Col.Width = 0
   Col.Caption = "Real ID"
   GridEX1.Columns.Item(2).Visible = False
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2625
   Col.Caption = MapText("รหัสผลิตภัณฑ์")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 3480
   Col.Caption = MapText("ชื่อผลิตภัณฑ์")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 1905
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ยอดผลิต")

   Set Col = GridEX1.Columns.add '6
   Col.Width = 1950
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ค่าใช้จ่ายรวม")

   Set Col = GridEX1.Columns.add '7
   Col.Width = 1380
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ต้นทุน RM")
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitGrid1
   Call InitNormalLabel(lblJobNo, MapText("เลขที่เอกสาร"))
   Call InitNormalLabel(lblJobDate, MapText("วันที่เอกสาร"))
   Call InitNormalLabel(lblPercent, MapText("ความคืบหน้า"))
   
   Call InitNormalLabel(lblStartJob, MapText("จากวันที่"))
   Call InitNormalLabel(lblFinishJob, MapText("ถึงวันที่"))
      
   Call txtJobNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtProgress.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtProgress.Enabled = False
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSave.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdStart.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSave, MapText("บันทึก"))
   Call InitMainButton(cmdStart, MapText("เริ่ม"))
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.NAME = GLB_FONT
   TabStrip1.Font.Size = 16
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.add().Caption = MapText("ต้นทุนผลิต")
   TabStrip1.Tabs.add().Caption = MapText("ผลิตภัณฑ์")
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
   Set m_CostProduction = New CCostProduction
   Set m_CostProductions = New Collection
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
     If m_CostProduction.ExpenseItem Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim Ce As CCostExpense
      If m_CostProduction.ExpenseItem.Count <= 0 Then
         Exit Sub
      End If
      Set Ce = GetItem(m_CostProduction.ExpenseItem, RowIndex, RealIndex)
      If Ce Is Nothing Then
         Exit Sub
      End If

      Values(1) = Ce.COST_EXPENSE_ID
      Values(2) = RealIndex
      Values(3) = Ce.PARAMETER_PROCESS_NAME
      Values(4) = ParcelTypeToText(Ce.PACKAGE_TYPE)
      Values(5) = RatioTypeToText(Ce.RATIO_TYPE)
      Values(6) = FormatNumber(Ce.EXPENSE_AMOUNT, 2)
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
     If m_CostProduction.CostItems Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim Ci As CCostPrdItem
      If m_CostProduction.CostItems.Count <= 0 Then
         Exit Sub
      End If
      Set Ci = GetItem(m_CostProduction.CostItems, RowIndex, RealIndex)
      If Ci Is Nothing Then
         Exit Sub
      End If

      Values(1) = Ci.COSTPRD_ITEM_ID
      Values(2) = RealIndex
      Values(3) = Ci.PART_NO
      Values(4) = Ci.PART_DESC
      Values(5) = FormatNumber(Ci.COST_AMOUNT, 3)
      Values(6) = FormatNumber(Ci.EXPENSE_AMOUNT, 2)
      Values(7) = FormatNumber(Ci.RAW_AMOUNT, 2)
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
     GridEX1.ItemCount = CountItem(m_CostProduction.ExpenseItem)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      cmdEdit.Enabled = True
      
      Call InitGrid2
      Call CalculateTotalRatio
     GridEX1.ItemCount = CountItem(m_CostProduction.CostItems)
      GridEX1.Rebind
   End If
End Sub

Private Sub txtBatchNo_Change()
   m_HasModify = True
End Sub

Private Sub txtJobDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtJobNo_Change()
   m_HasModify = True
End Sub

Private Sub ucltApproveByLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlFinishJob_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlJobDate_HasChange()
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
