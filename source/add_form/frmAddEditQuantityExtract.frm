VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditQuantityExtract 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   Icon            =   "frmAddEditQuantityExtract.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   9390
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   7335
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   12938
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboProcess2 
         Height          =   315
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2580
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.ComboBox cboProcess1 
         Height          =   315
         Left            =   2070
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2580
         Width           =   2325
      End
      Begin prjFarmManagement.uctlTextBox txtTotal 
         Height          =   465
         Left            =   2070
         TabIndex        =   4
         Top             =   2100
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlDate uctlFromDate 
         Height          =   405
         Left            =   2070
         TabIndex        =   0
         Top             =   900
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   15
         Top             =   0
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   2895
         Left            =   150
         TabIndex        =   8
         Top             =   3660
         Width           =   9105
         _ExtentX        =   16060
         _ExtentY        =   5106
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
         Column(1)       =   "frmAddEditQuantityExtract.frx":27A2
         Column(2)       =   "frmAddEditQuantityExtract.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditQuantityExtract.frx":290E
         FormatStyle(2)  =   "frmAddEditQuantityExtract.frx":2A6A
         FormatStyle(3)  =   "frmAddEditQuantityExtract.frx":2B1A
         FormatStyle(4)  =   "frmAddEditQuantityExtract.frx":2BCE
         FormatStyle(5)  =   "frmAddEditQuantityExtract.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditQuantityExtract.frx":2D5E
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   7
         Top             =   3120
         Width           =   9105
         _ExtentX        =   16060
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
      Begin prjFarmManagement.uctlDate uctlToDate 
         Height          =   405
         Left            =   2070
         TabIndex        =   1
         Top             =   1350
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   2070
         TabIndex        =   3
         Top             =   1770
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblProcess2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4440
         TabIndex        =   20
         Top             =   2670
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblProcess1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   870
         TabIndex        =   19
         Top             =   2670
         Width           =   1095
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   390
         TabIndex        =   18
         Top             =   2250
         Width           =   1575
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   7560
         TabIndex        =   2
         Top             =   900
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditQuantityExtract.frx":2F36
         ButtonStyle     =   3
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   390
         TabIndex        =   17
         Top             =   1440
         Width           =   1575
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   10
         Top             =   6660
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   9
         Top             =   6660
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditQuantityExtract.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   11
         Top             =   6660
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditQuantityExtract.frx":356A
         ButtonStyle     =   3
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   390
         TabIndex        =   16
         Top             =   990
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   7665
         TabIndex        =   13
         Top             =   6660
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   6015
         TabIndex        =   12
         Top             =   6660
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditQuantityExtract.frx":3884
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditQuantityExtract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_QuantityExtract As CQuantityExtract
Private m_Sp As CSystemParam
Private m_PartItems As Collection
Private m_ExtractItems As Collection

Public id As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Public TX_TYPE As String
Public Area As Long
Public ProcessID As Long

Private Sub cmdPasswd_Click()

End Sub

Private Sub cboPartType_Click()
   m_HasModify = True
End Sub

Private Sub cboUnit_Click()
   m_HasModify = True
End Sub

Private Sub chkPigFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cboProcess1_Click()
   m_HasModify = True
End Sub

Private Sub cboProcess2_Click()
   m_HasModify = True
End Sub

Private Sub cmdAdd_Click()
Dim Ji As CJobInput
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim Ei As CExtractItem
Dim Pi As CPartItem
Dim TempEi As CExtractItem
Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
Dim Qt As CQuantityExtract

   If Not VerifyDate(lblFromDate, uctlFromDate, False) Then
      Exit Sub
   End If
   If Not VerifyDate(lblToDate, uctlToDate, False) Then
      Exit Sub
   End If
   
   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("รายการใหม่", "-", "รายการล่าสุด")
   Set oMenu = Nothing
   If lMenuChosen <= 0 Then
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   
   Set TempRs = New ADODB.Recordset
   If lMenuChosen = 1 Then
      Set Ji = New CJobInput
      Ji.JOB_INOUT_ID = -1
      Ji.TX_TYPE = TX_TYPE
      Ji.FROM_DATE = uctlFromDate.ShowDate
      Ji.TO_DATE = uctlToDate.ShowDate
      Ji.JOB_DOC_TYPE = 1
      Ji.ProcessSet = GetProcessSet
      Call Ji.QueryData(2, TempRs, iCount)
      
      Set m_ExtractItems = Nothing
      Set m_ExtractItems = New Collection
      
      Set m_QuantityExtract.ExtractItems = Nothing
      Set m_QuantityExtract.ExtractItems = New Collection
      While Not TempRs.EOF
         Call Ji.PopulateFromRS(2, TempRs)
   '      Set Pi = GetPartItem(m_PartItems, Ji.PART_ITEM_ID)
         
         Set Ei = New CExtractItem
         Ei.Flag = "A"
         Ei.TOTAL_AMT = Ji.TX_AMOUNT
         Ei.STD_AMOUNT = Ji.TX_AMOUNT
         Ei.PART_ITEM_ID = Ji.PART_ITEM_ID
         Ei.PART_NO = Ji.PART_NO
         Ei.PART_DESC = Ji.PART_DESC
         Call m_QuantityExtract.ExtractItems.add(Ei, Trim(str(Ji.PART_ITEM_ID)))

         Set Ei = Nothing
         
         TempRs.MoveNext
      Wend
      
      Set Ji = Nothing
      
      GridEX1.ItemCount = CountItem(m_QuantityExtract.ExtractItems)
      GridEX1.Rebind
      Call CalculateTotalAmount
      ShowMode = SHOW_ADD
   ElseIf lMenuChosen = 3 Then
      Set Qt = New CQuantityExtract
      Qt.QUANTITY_EXTRACT_ID = -1
      Qt.Area = Area
      Qt.PROCESS_TYPE = ProcessID
      Call Qt.QueryData(2, TempRs, iCount)
      If Not TempRs.EOF Then
         Call Qt.PopulateFromRS(2, TempRs)
         id = Qt.QUANTITY_EXTRACT_ID
      Else
         id = -1
      End If
      If id > 0 Then
         Call QueryData(True)
         ShowMode = SHOW_EDIT
      End If
      Set Qt = Nothing
   End If
      
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
      
   Call EnableForm(Me, True)
   
   m_HasModify = True
End Sub

Private Sub CalculateTotalAmount()
Dim Ei As CExtractItem
Dim Sum As Double
   
   Sum = 0
   For Each Ei In m_QuantityExtract.ExtractItems
      If Ei.Flag <> "D" Then
         Sum = Sum + Ei.TOTAL_AMT
      End If
   Next Ei
   
   txtTotal.Text = FormatNumber(Sum, 3)
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
      Set frmAddEditExtractItem.ParentForm = Me
     Set frmAddEditExtractItem.TempCollection = m_QuantityExtract.ExtractItems
      frmAddEditExtractItem.id = id
      frmAddEditExtractItem.ShowMode = SHOW_EDIT
      frmAddEditExtractItem.HeaderText = MapText("แก้ไขจำนวนวัตถุดิบ")
      Load frmAddEditExtractItem
      frmAddEditExtractItem.Show 1

      OKClick = frmAddEditExtractItem.OKClick

      Unload frmAddEditExtractItem
      Set frmAddEditExtractItem = Nothing

      If OKClick Then
         Call ShowGrid
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   ElseIf TabStrip1.SelectedItem.Index = 6 Then
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Public Sub ShowGrid()
   GridEX1.ItemCount = CountItem(m_QuantityExtract.ExtractItems)
   GridEX1.Rebind
   
   Call CalculateTotalAmount
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
'   Call LoadAccessRight(Nothing, glbAccessRight, glbUser.GROUP_ID)
   OKClick = True
   Unload Me
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
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   Set Col = GridEX1.Columns.add '3
   Col.Width = 2855
   Col.Caption = MapText("รหัสวัตถุดิบ")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 4000
   Col.Caption = MapText("ชื่อวัตถุดิบ")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 1890
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("จำนวน")
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
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   Set Col = GridEX1.Columns.add '3
   Col.Width = 6030
   Col.Caption = MapText("ชื่อซัพพลายเออร์")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 2745
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ยอดนำเข้ารวม")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 2790
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("มูลค่านำเข้ารวม")
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      m_QuantityExtract.QUANTITY_EXTRACT_ID = id
      m_QuantityExtract.QueryFlag = 1
      If Not glbDaily.QueryQuantityExtract(m_QuantityExtract, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_QuantityExtract.PopulateFromRS(1, m_Rs)
      
      cboProcess1.ListIndex = IDToListIndex(cboProcess1, ProcessID)
      uctlFromDate.ShowDate = m_QuantityExtract.FROM_JOB_DATE
      uctlToDate.ShowDate = m_QuantityExtract.TO_JOB_DATE

      TabStrip1_Click
      Call CalculateTotalAmount
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cmdStart_Click()
Dim Jb As CJob
Dim TempRs As ADODB.Recordset
Dim TempJb As CJob
Dim TempJb2 As CJob
Dim IsOK As Boolean
Dim iCount As Long
Dim Ji As CJobInput
Dim LWH As CLotItemWH
Dim LTD As CLotDoc
Dim PD As CPalletDoc
Dim Ei1 As CExtractItem
Dim Ei2 As CExtractItem
Dim Ratio As Double
Dim Ivd As CInventoryDoc
Dim IvdWH As CInventoryWHDoc
Dim RCount As Long
Dim I As Long
Dim Percent As Double
Dim TempCol As Collection
Dim TempCol2 As Collection
Dim TempCol3 As Collection
Dim Temp_LWH As CLotItemWH
Dim TempOutput As CJobOutput
Dim m_JobCollection As Collection
Dim strJobNo As String


   If CountItem(m_QuantityExtract.ExtractItems) <= 0 Then
      glbErrorLog.LocalErrorMsg = "กรุณาทำการเพิ่มปริมาณวัตถุดิดิบที่ใช้ก่อน (กดปุ่มเพิ่ม)"
      glbErrorLog.ShowUserError
      Exit Sub
   End If
      
   Call EnableForm(Me, False)
   Set TempRs = New ADODB.Recordset
   
   Set Jb = New CJob
   Jb.JOB_ID = -1
   Jb.FROM_DATE = uctlFromDate.ShowDate
   Jb.TO_DATE = uctlToDate.ShowDate
   Jb.ProcessSet = GetProcessSet
   Call Jb.QueryData(1, TempRs, RCount)
   
   Call glbDaily.StartTransaction
   I = 0
   prgProgress.MAX = 100
   prgProgress.MIN = 0
   While Not TempRs.EOF
      I = I + 1
      Percent = MyDiffEx(I, RCount) * 100
      prgProgress.Value = Percent
      prgProgress.Refresh
      
      Call Jb.PopulateFromRS(1, TempRs)
        
      Set TempJb = New CJob
      TempJb.JOB_ID = Jb.JOB_ID
      TempJb.QueryFlag = 1
      Call glbProduction.QueryJob2(TempJb, m_Rs, iCount, IsOK, glbErrorLog)
      If Not m_Rs.EOF Then
          Call TempJb.PopulateFromRS(1, m_Rs)
      End If
        
      TempJb.AddEditMode = SHOW_EDIT
      If TX_TYPE = "E" Then
         Set TempCol = TempJb.Inputs
      ElseIf TX_TYPE = "I" Then
         Set TempCol = TempJb.Outputs
         If ProcessID = 4 Then 'กรณี การบรรจุ BAG รอคิดก่อน
            Set TempCol2 = TempJb.InventoryWhDoc.Item(1).C_LotItemsWH 'new
         End If
      End If
      
      For Each Ji In TempCol
'          Set Ei1 = m_ExtractItems(Trim(Str(Ji.PART_ITEM_ID))) 'มาตรฐาน
          Set Ei2 = m_QuantityExtract.ExtractItems(Trim(str(Ji.PART_ITEM_ID))) 'อัตรส่วนจริง
          'Ei2.STD_AMOUNT เป็นค่ามาตรฐาน
          Ratio = MyDiffEx(Ei2.TOTAL_AMT, Ei2.STD_AMOUNT)
          
          Ji.Flag = "E"
          'Ji.STD_AMOUNT = Ji.TX_AMOUNT 'สลับมาเก็บไว้เป็นค่ามาตรฐาน       จิวลบออก วันที่ วันจันทร์ ที่ 05 เดือน 08  ปี 2556 เพราะว่าพอใช้ PLC แล้วยอดมันไปทับค่ามาตรฐาน
          Ji.TX_AMOUNT = Ratio * Ji.TX_AMOUNT
         
         Set m_JobCollection = New Collection
         Set TempCol3 = New Collection
         
         If Val(Jb.JOB_ID_REF) > 0 Then 'ถ้าเป็นการกระจาย  JOB SPLIT
              strJobNo = Left(Jb.JOB_NO, 11)
              Call LoadJobByJobNo(Nothing, m_JobCollection, , , 1, Trim(strJobNo)) 'ดึง jobno ทุก job ที่มีเลข job เดียวกัน
            For Each TempJb2 In m_JobCollection
               If TempJb2.JOB_ID = Jb.JOB_ID_REF Then
                 Set LWH = New CLotItemWH
                 LWH.TX_AMOUNT = TempJb2.Outputs.Item(1).TX_AMOUNT
                 If TempJb2.JOB_ID > 0 Then
                    Call TempCol3.add(LWH, str(TempJb2.JOB_ID))
                 End If
                    
                  Set TempCol2 = TempJb2.InventoryWhDoc.Item(1).C_LotItemsWH 'new
                 Set Temp_LWH = GetObject("CLotItemWH", TempCol3, str(Jb.JOB_ID_REF), False) 'ทำรองรับกรณีที่เป็นการแยก Job เพราะเมื่อกระจาย ทุก job
                 If (Not Temp_LWH Is Nothing) Then
                       For Each LWH In TempCol2 'new
                       LWH.Flag = "E"
                       Temp_LWH.TX_AMOUNT = Ji.TX_AMOUNT + Temp_LWH.TX_AMOUNT
                       LWH.TX_AMOUNT = Temp_LWH.TX_AMOUNT
                       LWH.GOOD_AMOUNT = Temp_LWH.TX_AMOUNT
                          For Each LTD In LWH.C_LotDoc  'new
                             For Each PD In LTD.C_PalletDoc   'new
                                PD.Flag = "E"
                                PD.CAPACITY_AMOUNT = Temp_LWH.TX_AMOUNT
                             Next PD
                          Next LTD
                       Next LWH
                    End If
               
                     Call glbDaily.Job2InventoryWhDoc(TempJb2, IvdWH, 1, 11, , 2)
                     Call glbDaily.AddEditInventoryWhDoc(IvdWH, IsOK, False, glbErrorLog)
'                     TempJb.INVENTORY_WH_DOC_ID = IvdWH.INVENTORY_WH_DOC_ID

                     Set TempCol2 = Nothing 'ไม่ให้เข้าทำเมื่อออกไป
                  End If
            Next TempJb2
        Else 'ถ้าเป็นการกระจาย Job แม่
         Call LoadJobByJobNo2(Nothing, m_JobCollection, , , 1, Trim(Jb.JOB_NO)) 'ดึง jobno ทุก job ที่มีเลข job เดียวกัน
         For Each TempJb2 In m_JobCollection
               If TempJb2.JOB_ID_REF = Jb.JOB_ID Then
                  Set LWH = New CLotItemWH
                  LWH.TX_AMOUNT = TempJb2.Outputs.Item(1).TX_AMOUNT
                     If TempJb2.JOB_ID_REF > 0 Then
                     Call TempCol3.add(LWH, str(TempJb2.JOB_ID_REF))
                     End If
               End If
         Next
      End If
         
      If Not TempCol2 Is Nothing Then
            Set Temp_LWH = GetObject("CLotItemWH", TempCol3, str(Jb.JOB_ID), False) 'ทำรองรับกรณีที่เป็นการแยก Job เพราะเมื่อกระจาย ทุก job
            If (Not Temp_LWH Is Nothing) Then
                  For Each LWH In TempCol2 'new
                  LWH.Flag = "E"
                  Temp_LWH.TX_AMOUNT = Ji.TX_AMOUNT + Temp_LWH.TX_AMOUNT
                  LWH.TX_AMOUNT = Temp_LWH.TX_AMOUNT
                  LWH.GOOD_AMOUNT = Temp_LWH.TX_AMOUNT
                     For Each LTD In LWH.C_LotDoc  'new
                        For Each PD In LTD.C_PalletDoc   'new
                           PD.Flag = "E"
                           PD.CAPACITY_AMOUNT = Temp_LWH.TX_AMOUNT
                        Next PD
                     Next LTD
                  Next LWH
               Else
                  For Each LWH In TempCol2 'new
                  LWH.Flag = "E"
                  LWH.TX_AMOUNT = Ji.TX_AMOUNT
                  LWH.GOOD_AMOUNT = Ji.TX_AMOUNT
                     For Each LTD In LWH.C_LotDoc  'new
                        For Each PD In LTD.C_PalletDoc   'new
                           PD.Flag = "E"
                           PD.CAPACITY_AMOUNT = Ji.TX_AMOUNT
                        Next PD
                     Next LTD
                  Next LWH
               End If
            End If
      Next Ji
      
      Call glbDaily.Job2InventoryDoc(TempJb, Ivd, 1, 11)
      Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog)
      TempJb.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID
      
      If ProcessID = 4 And Not TempCol2 Is Nothing Then    'กรณี การบรรจุ BAG รอคิดก่อน
         Call glbDaily.Job2InventoryWhDoc(TempJb, IvdWH, 1, 11, , 2)
         Call glbDaily.AddEditInventoryWhDoc(IvdWH, IsOK, False, glbErrorLog)
         TempJb.INVENTORY_WH_DOC_ID = IvdWH.INVENTORY_WH_DOC_ID
      End If
      
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
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_PartItems = Nothing
   Set m_ExtractItems = Nothing
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Index = 1 Then
      If m_QuantityExtract.ExtractItems Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim CR As CExtractItem
      If m_QuantityExtract.ExtractItems.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_QuantityExtract.ExtractItems, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If

      Values(1) = CR.PART_ITEM_ID
      Values(2) = RealIndex
      Values(3) = CR.PART_NO
      Values(4) = CR.PART_DESC
      Values(5) = FormatNumber(CR.TOTAL_AMT, 3)
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   End If

   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
   
   If Not VerifyDate(lblFromDate, uctlFromDate, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblToDate, uctlToDate, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_QuantityExtract.QUANTITY_EXTRACT_ID = id
   m_QuantityExtract.AddEditMode = ShowMode
   m_QuantityExtract.FROM_JOB_DATE = uctlFromDate.ShowDate
   m_QuantityExtract.TO_JOB_DATE = uctlToDate.ShowDate
   m_QuantityExtract.Area = Area
   m_QuantityExtract.PROCESS_TYPE = ProcessID
   
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditQuantityExtract(m_QuantityExtract, IsOK, True, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
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

Private Function GetProcessSet() As String
Dim Process1 As Long
Dim Process2 As Long
Dim TempArr(1 To 10) As Long
Dim I As Long
Dim TempSet As String
Dim MAX As Long
Dim Count As Long

   I = 0
   Process1 = cboProcess1.ItemData(Minus2Zero(cboProcess1.ListIndex))
   If Process1 > 0 Then
      I = I + 1
      TempArr(I) = Process1
   End If
   
   Process2 = cboProcess2.ItemData(Minus2Zero(cboProcess2.ListIndex))
   If Process2 > 0 Then
      I = I + 1
      TempArr(I) = Process2
   End If
   
   TempSet = ""
   Count = 0
   MAX = I
   For I = 1 To MAX
      TempSet = TempSet & TempArr(I)
      Count = Count + 1
      
      If I < MAX Then
         TempSet = TempSet & ", "
      Else
         TempSet = TempSet & ""
      End If
   Next I
   
   If Count > 0 Then
      TempSet = "(" & TempSet & ")"
   Else
      TempSet = ""
   End If
   GetProcessSet = TempSet
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
'      Call LoadPartItem(Nothing, m_PartItems)
      Call LoadProcess(cboProcess1, , ProcessID)
      Call LoadProcess(cboProcess2)
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         uctlFromDate.ShowDate = Now
         uctlToDate.ShowDate = Now
         id = 0
         cboProcess1.ListIndex = IDToListIndex(cboProcess1, ProcessID)
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
   ElseIf Shift = 0 And KeyCode = 117 Then
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

Private Sub InitFormLayout()
   Call InitGrid1
   
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblFromDate, MapText("จากวันที่"))
   Call InitNormalLabel(lblToDate, MapText("ถึงวันที่"))
   Call InitNormalLabel(lblTotal, MapText("ยอดรวม"))
   Call InitNormalLabel(lblProcess1, MapText("โปรเซส 1"))
   Call InitNormalLabel(lblProcess2, MapText("โปรเซส 2"))
   
   Call txtTotal.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtTotal.Enabled = False
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdStart.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdStart, MapText("เริ่ม"))
   cmdDelete.Enabled = False
   
   Call InitCombo(cboProcess1)
   Call InitCombo(cboProcess2)
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.NAME = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.add().Caption = MapText("วัตถุดิบ")
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   Set m_QuantityExtract = New CQuantityExtract
   Set m_Rs = New ADODB.Recordset
   Set m_PartItems = New Collection
   Set m_ExtractItems = New Collection
   
   Call EnableForm(Me, False)
   m_HasActivate = False
      
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

Private Sub TabStrip1_Click()
   If TabStrip1.SelectedItem.Index = 1 Then
      Call InitGrid1
      
      GridEX1.ItemCount = CountItem(m_QuantityExtract.ExtractItems)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub txtBarcode_Change()
   m_HasModify = True
End Sub

Private Sub txtPartNo_Change()
   m_HasModify = True
End Sub

Private Sub txtName_Change()
   m_HasModify = True
End Sub

Private Sub txtMaxUser_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox1_Change()
   m_HasModify = True
End Sub

Private Sub txtUnitWeight_Change()
   m_HasModify = True
End Sub

Private Sub uctlFromDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlToDate_HasChange()
   m_HasModify = True
End Sub
